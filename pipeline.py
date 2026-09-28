# -*- coding: utf-8 -*-
"""
Pipeline de génération automatique de bilan orthodontique.
Fonctions locales : traitement photos, rendu STL, figure Steiner, assemblage du bilan Word/PDF.
Aucune donnée ne quitte l'ordinateur.
"""
import os, io, json, shutil, zipfile, subprocess, sys
from PIL import Image, ImageOps

# Dossier de données (partagé avec l'app) + modèle de reconnaissance auto-appris
PDATA = os.environ.get("BILANODF_DATA") or os.path.expanduser("~/BilanODF_Data")
MODEL_PATH = os.path.join(PDATA, "photo_model.npz")

# ----------------------------------------------------------------------------
# Noms de fichiers attendus dans <patient>/01_photos_brutes et 04_stl_bruts
PHOTO_SLOTS = [
    "exo_face_repos", "exo_face_sourire", "exo_profil",
    "endo_occlusion_frontale", "endo_laterale_droite", "endo_laterale_gauche",
    "endo_occlusal_maxillaire", "endo_occlusal_mandibulaire",
]
RADIO_SLOTS = ["panoramique", "teleradiographie_profil"]
STL_SLOTS = ["UpperJawScan", "LowerJawScan"]

# ----------------------------------------------------------------------------
# 1) TRAITEMENT DES PHOTOS  (recadrage aux proportions du modèle)
#    aspect = largeur/hauteur cible ; cx,cy centre ; zoom = fraction de hauteur utilisée
_CROP = {
    # slot: (aspect, cx, cy, zoom, transform)   transform: None|'mirror'|'flip'|'lat'
    "exo_face_repos":            (697/768, 0.50, 0.47, 0.84, None),
    "exo_face_sourire":          (697/768, 0.50, 0.47, 0.84, None),
    "exo_profil":                (758/768, 0.51, 0.47, 0.86, None),
    "endo_occlusion_frontale":   (819/599, 0.50, 0.52, 0.62, None),
    "endo_laterale_droite":      (860/599, 0.50, 0.52, 0.82, "lat"),
    "endo_laterale_gauche":      (860/599, 0.50, 0.52, 0.82, "lat"),
    "endo_occlusal_maxillaire":  (861/645, 0.50, 0.45, 0.74, "mirror"), # incisives en HAUT
    "endo_occlusal_mandibulaire":(861/645, 0.50, 0.50, 0.74, "flip"),   # incisives en BAS
}

def _crop_aspect_box(W, H, aspect, cx, cy, zoom):
    """Boîte auto centrée -> (x, y, cw, ch) en pixels."""
    ch = int(H * zoom); cw = int(ch * aspect)
    if cw > W:
        cw = W; ch = int(cw / aspect)
    x = int(W * cx - cw / 2); y = int(H * cy - ch / 2)
    x = max(0, min(x, W - cw)); y = max(0, min(y, H - ch))
    return x, y, cw, ch

def _crop_aspect(im, aspect, cx, cy, zoom):
    x, y, cw, ch = _crop_aspect_box(im.width, im.height, aspect, cx, cy, zoom)
    return im.crop((x, y, x + cw, y + ch))

def _family(slot):
    if slot.startswith("exo"): return "exo"
    if "occlusal" in slot: return "occlusal"
    return "endo"

def roi_box(im, family):
    """Détecte la ZONE UTILE et renvoie (x, y, w, h) en pixels, ou None si peu fiable.
    - exo  : tête/visage = ce qui se détache du fond clair uniforme (coins).
    - endo/occlusal : contenu buccal = dents (clair/peu saturé) + muqueuse (rose/rouge),
      en ignorant les bords sombres (lèvres/écarteurs) et la peau."""
    try:
        import numpy as np
        small = im.convert("RGB").resize((200, max(1, int(200 * im.height / im.width))))
        a = np.asarray(small).astype(float) / 255.0
        H0, W0 = a.shape[:2]
        R, G, B = a[..., 0], a[..., 1], a[..., 2]
        mx = a.max(-1); mn = a.min(-1); sat = (mx - mn) / np.clip(mx, 1e-6, 1)
        val = a.mean(-1)
        if family == "exo":
            # couleur de fond = médiane des 4 coins ; premier plan = ce qui s'en écarte
            k = max(6, W0 // 12)
            corners = np.concatenate([a[:k, :k].reshape(-1, 3), a[:k, -k:].reshape(-1, 3),
                                      a[-k:, :k].reshape(-1, 3), a[-k:, -k:].reshape(-1, 3)])
            bg = np.median(corners, 0)
            dist = np.sqrt(((a - bg) ** 2).sum(-1))
            mask = dist > 0.18
        else:
            teeth = (mx > 0.55) & (sat < 0.22)
            muc = (R > 0.4) & ((R - G) > 0.10) & ((R - B) > 0.10) & (sat > 0.22)
            dark = val < 0.16
            mask = (teeth | muc) & (~dark)
        ys, xs = np.where(mask)
        if len(xs) < 0.02 * H0 * W0:      # trop peu -> pas fiable
            return None
        x0, x1 = np.percentile(xs, [2, 98]); y0, y1 = np.percentile(ys, [2, 98])
        bw = (x1 - x0) / W0; bh = (y1 - y0) / H0
        if bw < 0.08 or bh < 0.08 or (bw > 0.97 and bh > 0.97):
            return None                    # dégénéré ou couvre tout -> repli
        sx = im.width / W0; sy = im.height / H0
        return (x0 * sx, y0 * sy, (x1 - x0) * sx, (y1 - y0) * sy)
    except Exception:
        return None

CROP_CALIB_PATH = os.path.join(PDATA, "crop_calib.json")
# marges par défaut (resserrées) : fraction ajoutée autour de la zone détectée
DEFAULT_MARGIN = {"exo": 0.10, "endo": 0.05, "occlusal": 0.03}

def load_crop_calib():
    try:
        if os.path.exists(CROP_CALIB_PATH):
            return json.load(open(CROP_CALIB_PATH))
    except Exception:
        pass
    return {}

def axis_angle(im, family):
    """Angle (degrés, [-90,90]) de l'axe principal de la bande dentaire par rapport à
    l'horizontale, ou None si peu fiable. 0 = plan d'occlusion horizontal.
    Intra-oral uniquement."""
    if family == "exo":
        return None
    try:
        import numpy as np
        small = im.convert("RGB").resize((200, max(1, int(200 * im.height / im.width))))
        a = np.asarray(small).astype(float) / 255.0
        mx = a.max(-1); mn = a.min(-1); sat = (mx - mn) / np.clip(mx, 1e-6, 1)
        teeth = (mx > 0.55) & (sat < 0.22)
        ys, xs = np.where(teeth)
        if len(xs) < 200:
            return None
        x = xs - xs.mean(); y = ys - ys.mean()
        cxx = float((x * x).mean()); cyy = float((y * y).mean()); cxy = float((x * y).mean())
        if abs(cxx - cyy) < 1e-9 and abs(cxy) < 1e-9:
            return None
        theta = 0.5 * np.degrees(np.arctan2(2 * cxy, cxx - cyy))   # y vers le bas
        if theta > 90: theta -= 180
        if theta < -90: theta += 180
        return float(theta)
    except Exception:
        return None

def estimate_occlusal_tilt(im, family):
    """Compat : renvoie une estimation signée de l'inclinaison (None si douteuse)."""
    ang = axis_angle(im, family)
    if ang is None or abs(ang) > 30:
        return None
    return -ang

def best_straighten(base_im, family):
    """Angle de redressement (degrés) à appliquer pour rendre le plan d'occlusion
    horizontal. DÉTERMINISTE : teste les deux sens et garde celui qui aplatit le plus
    la bande dentaire -> plus d'ambiguïté de sens. 0 si inutile ou peu fiable."""
    if family == "exo":
        return 0.0
    a0 = axis_angle(base_im, family)
    if a0 is None or abs(a0) < 0.6 or abs(a0) > 30:
        return 0.0
    m = min(18.0, abs(a0))
    best, best_dev = 0.0, abs(a0)
    for cand in (m, -m):
        try:
            dev = axis_angle(base_im.rotate(cand, expand=True), family)
        except Exception:
            dev = None
        if dev is not None and abs(dev) < best_dev - 0.3:
            best_dev, best = abs(dev), cand
    return round(best, 2)

def _box_for_roi(W, H, aspect, roi, margin):
    """Plus petite boîte de ratio `aspect` contenant la ROI + marge, centrée sur elle,
    puis recadrée dans l'image. Renvoie (x, y, cw, ch) en pixels."""
    rx, ry, rw, rh = roi
    cx = rx + rw / 2.0; cy = ry + rh / 2.0
    need_w = rw * (1 + 2 * margin); need_h = rh * (1 + 2 * margin)
    cw = max(need_w, need_h * aspect); ch = cw / aspect
    cw = min(cw, W); ch = cw / aspect
    if ch > H: ch = H; cw = ch * aspect
    x = cx - cw / 2.0; y = cy - ch / 2.0
    x = max(0, min(x, W - cw)); y = max(0, min(y, H - ch))
    return int(round(x)), int(round(y)), int(round(cw)), int(round(ch))

def auto_rotation(base, slot):
    """Rotation automatique par défaut : latérales en portrait -> 90° pour rendre
    le plan d'occlusion horizontal. 0 sinon."""
    p = os.path.join(base, "01_photos_brutes", slot + ".jpg")
    if not os.path.exists(p):
        return 0
    if _CROP.get(slot, (0,0,0,0,None))[4] == "lat":
        im = ImageOps.exif_transpose(Image.open(p))
        return 90 if im.height > im.width else 0
    return 0

def default_crop(base, slot):
    """Cadrage automatique par défaut exprimé dans le repère de l'image ORIENTEE
    (après rotation + miroir + retournement). Fractions [0..1]. C'est la source de
    vérité partagée avec l'éditeur de l'app : ce que l'appli affiche = ce que le
    pipeline produit."""
    cfg = _CROP.get(slot)
    if not cfg:
        return None
    aspect, cx, cy, zoom, tf = cfg
    p = os.path.join(base, "01_photos_brutes", slot + ".jpg")
    if not os.path.exists(p):
        return {"rot": 0, "fine": 0, "mirror": False, "flip": False,
                "x": 0.0, "y": 0.0, "w": 1.0, "h": 1.0, "aspect": aspect}
    im0 = ImageOps.exif_transpose(Image.open(p))
    rot = 90 if (tf == "lat" and im0.height > im0.width) else 0
    fam = _family(slot)
    calib = load_crop_calib().get(fam, {})
    # redressement automatique DÉSACTIVÉ (à la demande) : aucune rotation fine imposée.
    # La molette de redressement reste disponible manuellement dans l'éditeur.
    fine = 0.0
    total = rot % 360
    im = im0.rotate(total, expand=True, resample=Image.BICUBIC) if total else im0
    W, H = im.size
    # 2) cadrage sur la zone détectée, marge calibrée (resserrée par défaut)
    margin = float(calib.get("margin", DEFAULT_MARGIN.get(fam, 0.08)))
    roi = roi_box(im, fam)
    if roi is not None:
        x, y, cw, ch = _box_for_roi(W, H, aspect, roi, margin)
    else:
        x, y, cw, ch = _crop_aspect_box(W, H, aspect, cx, cy, zoom)
    mirror = (tf == "mirror"); flip = (tf == "flip")
    # la boîte est calculée AVANT miroir/flip ; on la transpose dans le repère orienté
    if mirror: x = W - (x + cw)
    if flip:   y = H - (y + ch)
    return {"rot": rot, "fine": round(fine, 2), "mirror": mirror, "flip": flip,
            "x": round(x / W, 5), "y": round(y / H, 5),
            "w": round(cw / W, 5), "h": round(ch / H, 5), "aspect": aspect}

def train_crop_calib(patients_root=None, model_path=None):
    """Apprend TA façon de cadrer en comparant, pour chaque photo, la version BRUTE
    (01_photos_brutes) et la version RECADRÉE (02_photos_traitees) : la fraction que la
    zone utile occupe dans la photo recadrée donne ta marge préférée par type de vue.
    Aucun recadrage manuel nécessaire. Pondéré par récence. Enregistre crop_calib.json."""
    import numpy as np
    patients_root = patients_root or os.path.join(PDATA, "patients")
    model_path = model_path or CROP_CALIB_PATH
    fams = ("exo", "endo", "occlusal")
    margins = {f: [] for f in fams}   # (mtime, marge implicite)
    n_ex = 0
    if os.path.isdir(patients_root):
        for slug in os.listdir(patients_root):
            recs = os.path.join(patients_root, slug, "records")
            if not os.path.isdir(recs): continue
            for rid in os.listdir(recs):
                out = os.path.join(recs, rid, "02_photos_traitees")
                if not os.path.isdir(out): continue
                for fn in os.listdir(out):
                    if not fn.lower().endswith(".jpg"): continue
                    slot = os.path.splitext(fn)[0]
                    if slot not in _CROP: continue
                    fam = _family(slot)
                    p = os.path.join(out, fn)
                    try:
                        im = ImageOps.exif_transpose(Image.open(p))
                        W, H = im.size
                        roi = roi_box(im, fam)
                        if roi is None: continue
                        _rx, _ry, rw, rh = roi
                        maxf = max(rw / max(W, 1), rh / max(H, 1))   # fraction occupée (dim. la plus serrée)
                        if maxf <= 0.05: continue
                        imp = (1.0 / maxf - 1.0) / 2.0               # marge implicite
                        if 0.0 <= imp <= 0.6:
                            margins[fam].append((os.path.getmtime(p), imp)); n_ex += 1
                    except Exception:
                        pass
    calib = {}
    RECENT = 30   # on privilégie tes photos recadrées les plus récentes
    for f in fams:
        ms = [v for _t, v in sorted(margins[f], key=lambda z: z[0], reverse=True)[:RECENT]]
        if ms:
            calib[f] = {"margin": round(float(np.median(ms)), 3), "n_margin": len(margins[f])}
    try:
        json.dump(calib, open(model_path, "w"), ensure_ascii=False, indent=1)
    except Exception as e:
        return {"ok": False, "msg": str(e)}
    return {"ok": True, "n": n_ex, "calib": calib}

def process_photos(base, rotations=None, crops=None):
    """rotations: {slot: degrés} rotation simple (legacy, sinon auto).
       crops: {slot: {rot,mirror,flip,x,y,w,h}} cadrage manuel complet, prioritaire.
       Sans crop pour un slot -> comportement automatique historique identique."""
    rotations = rotations or {}; crops = crops or {}
    src = os.path.join(base, "01_photos_brutes")
    out = os.path.join(base, "02_photos_traitees"); os.makedirs(out, exist_ok=True)
    done = []
    for slot, (aspect, cx, cy, zoom, tf) in _CROP.items():
        p = os.path.join(src, slot + ".jpg")
        if not os.path.exists(p):
            continue
        im = ImageOps.exif_transpose(Image.open(p)).convert("RGB")
        c = crops.get(slot)
        if c is None and slot not in rotations:
            # pas de réglage manuel -> cadrage AUTO détecté (identique à ce que montre l'éditeur)
            c = default_crop(base, slot)
        if c:
            # --- cadrage manuel complet (orientation puis rognage, repère orienté) ---
            rot = int(c.get("rot", 0)) % 360
            try: fine = max(-45.0, min(45.0, float(c.get("fine", 0))))
            except Exception: fine = 0.0
            ang = rot + fine
            if ang:
                im = im.rotate(ang, expand=True, resample=Image.BICUBIC)
            if c.get("mirror"): im = ImageOps.mirror(im)
            if c.get("flip"):   im = ImageOps.flip(im)
            W, H = im.size
            x = int(round(float(c.get("x", 0)) * W)); y = int(round(float(c.get("y", 0)) * H))
            cw = int(round(float(c.get("w", 1)) * W)); ch = int(round(float(c.get("h", 1)) * H))
            cw = max(8, min(cw, W)); ch = max(8, min(ch, H))
            x = max(0, min(x, W - cw)); y = max(0, min(y, H - ch))
            im = im.crop((x, y, x + cw, y + ch))
        else:
            # --- cadrage automatique (comportement historique) ---
            rot = rotations.get(slot)
            if rot is None:
                rot = 90 if (tf == "lat" and im.height > im.width) else 0
            rot = int(rot) % 360
            if rot:
                im = im.rotate(rot, expand=True)     # plan d'occlusion horizontal, etc.
            im = _crop_aspect(im, aspect, cx, cy, zoom)
            if tf == "mirror":
                im = ImageOps.mirror(im)             # occlusal maxillaire : dé-miroir
            elif tf == "flip":
                im = ImageOps.flip(im)               # occlusal mandibulaire : incisives en bas
        im.thumbnail((1000, 1400), Image.LANCZOS)
        im.save(os.path.join(out, slot + ".jpg"), quality=92)
        done.append(slot)
    return done

# ----------------------------------------------------------------------------
# 2) RENDU DES MODELES 3D (STL)  — gris sur blanc, style MeshLab clair
def render_stl(base):
    import numpy as np, trimesh, pyvista as pv
    pv.OFF_SCREEN = True
    UP = os.path.join(base, "04_stl_bruts", "UpperJawScan.stl")
    LO = os.path.join(base, "04_stl_bruts", "LowerJawScan.stl")
    OUT = os.path.join(base, "05_stl_rendus"); os.makedirs(OUT, exist_ok=True)
    has_u = os.path.exists(UP); has_l = os.path.exists(LO)
    if not (has_u or has_l): return
    mu = trimesh.load(UP, force="mesh") if has_u else None
    ml = trimesh.load(LO, force="mesh") if has_l else None
    if (mu is not None) and (len(getattr(mu, "vertices", [])) == 0): mu = None
    if (ml is not None) and (len(getattr(ml, "vertices", [])) == 0): ml = None
    if (mu is None) and (ml is None): return
    allv = np.vstack([np.asarray(m.vertices) for m in (mu, ml) if m is not None]); c = allv.mean(0)
    X = allv - c; w, Vec = np.linalg.eigh(X.T @ X / len(X))
    vert = Vec[:, 0]; ap = Vec[:, 1]
    def proj(v, a): return (v - c) @ a
    zc = proj(allv, ap); xc = proj(allv, Vec[:, 2]); zr = zc.max() - zc.min()
    def cf(m, xx):
        if m.sum() < 50: return 0
        return np.mean(np.abs(xx[m]) < 0.12 * (xx.max() - xx.min()))
    if cf(zc < zc.min() + 0.15 * zr, xc) > cf(zc > zc.min() + 0.85 * zr, xc): ap = -ap
    if (mu is not None) and (ml is not None):
        if proj(np.asarray(mu.vertices), vert).mean() < proj(np.asarray(ml.vertices), vert).mean(): vert = -vert
    lr = np.cross(vert, ap); lr /= np.linalg.norm(lr); R = np.vstack([lr, vert, ap])
    def canon(m):
        return pv.PolyData((np.asarray(m.vertices) - c) @ R.T,
                           np.hstack([np.full((len(m.faces), 1), 3), m.faces]).astype(np.int64).ravel())
    U0 = canon(mu) if mu is not None else None; L0 = canon(ml) if ml is not None else None
    MAT = "#dcdcdc"; SS = 1
    # Rendu en gris uni (coloration dents/gencive abandonnée à la demande).
    U, L = U0, L0; COLORED = False
    def shot(meshes, pos, up, fn, zoom=1.35):
        meshes = [m for m in meshes if m is not None]
        if not meshes: return
        p = pv.Plotter(off_screen=True, window_size=[1100*SS, 850*SS])
        p.set_background([0.33, 0.49, 0.78], top=[0.88, 0.92, 0.98])   # dégradé bleu type MeshLab
        for mm in meshes:
            if COLORED and ("rgb" in mm.point_data):
                p.add_mesh(mm, scalars="rgb", rgb=True, smooth_shading=True, specular=0.14,
                           specular_power=18, ambient=0.34, diffuse=0.82)
            else:
                p.add_mesh(mm, color=MAT, smooth_shading=True, specular=0.18,
                           specular_power=15, ambient=0.35, diffuse=0.75)
        p.enable_lightkit()
        try: p.enable_anti_aliasing('ssaa')
        except Exception: pass
        import numpy as _np
        b = p.bounds
        ctr = ((b[0] + b[1]) / 2.0, (b[2] + b[3]) / 2.0, (b[4] + b[5]) / 2.0)
        dv = _np.array(pos, dtype=float); nn = _np.linalg.norm(dv); dv = (dv / nn) if nn else dv
        diag = ((b[1] - b[0]) ** 2 + (b[3] - b[2]) ** 2 + (b[5] - b[4]) ** 2) ** 0.5 or 100.0
        p.camera.focal_point = ctr
        p.camera.position = (ctr[0] + dv[0] * diag, ctr[1] + dv[1] * diag, ctr[2] + dv[2] * diag)
        p.camera.up = up
        p.reset_camera(); p.camera.zoom(zoom)
        im = Image.fromarray(p.screenshot(return_img=True)); p.close()
        im.thumbnail((1100, 1100), Image.LANCZOS)
        im.save(os.path.join(OUT, fn))
    D = 240
    shot([U], (0, -D, 0), (0, 0,  1), "upper_occlusal.png", 1.35)   # maxillaire : incisives en HAUT
    shot([L], (0, -D, 0), (0, 0, -1), "lower_occlusal.png", 1.35)   # mandibulaire : incisives en BAS
    shot([U, L], (0, 0,  D), (0, 1, 0), "front_occ.png", 1.35)
    # occlusion latérale : on retire la moitié opposée pour dégager la classe d'Angle
    xr = np.vstack([m.points for m in (U, L) if m is not None])[:, 0]; thr = -0.10 * (xr.max() - xr.min())
    def half(mesh, keep_pos):
        return mesh.clip(normal=(1, 0, 0), origin=(thr if keep_pos else -thr, 0, 0), invert=not keep_pos)
    shot([half(U, True) if U is not None else None, half(L, True) if L is not None else None],  ( D, 0, 0), (0, 1, 0), "right_occ.png", 1.5)
    shot([half(U, False) if U is not None else None, half(L, False) if L is not None else None], (-D, 0, 0), (0, 1, 0), "left_occ.png", 1.5)
    # courbe de Spee : arcade inférieure seule, moitié proche, vue latérale
    if L is not None:
        shot([half(L, True)],  ( D, D*0.12, D*0.08), (0, 1, 0), "spee_lower.png", 1.4)
        shot([half(L, False)], (-D, D*0.12, D*0.08), (0, 1, 0), "spee_left.png", 1.4)
    # Ré-applique les retournements manuels (marqueurs .flip_<vue>) après chaque rendu
    try:
        for _f in os.listdir(OUT):
            if _f.startswith(".flip_"):
                _p = os.path.join(OUT, _f[6:] + ".png")
                if os.path.exists(_p):
                    ImageOps.flip(Image.open(_p)).save(_p)
    except Exception:
        pass

# ----------------------------------------------------------------------------
# 3) FIGURE STEINER  (recrée le "Graphique" WebCeph en PNG branded)
STEINER_NORMS = [
    # (nom, moyenne, ecart_type)  -- ordre d'affichage
    ("SNA (°)", 80.22, 2.8), ("SNB (°)", 76.18, 2.5), ("ANB (°)", 4.08, 1.6),
    ("U1-NA (°)", 22, 5.0), ("U1-NA (mm)", 4, 3.0),
    ("L1-NB (°)", 25, 5.0), ("L1-NB (mm)", 4, 2.0),
    ("Wits (mm)", -0.33, 2.7), ("Pog-NB (mm)", 2.25, 1.3),
    ("Plan mandibulaire (Go-Gn/SN) (°)", 32, 4.0), ("Plan occlusal/SN (°)", 14, 4.0),
    ("Angle interincisif (°)", 128, 5.3),
]

def build_steiner_figure(base):
    """Tableau propre : Mesure | Moyenne | Écart-type | Résultat | Interprétation.
    Sans sévérité ni graphique polygonal."""
    import matplotlib
    matplotlib.use("Agg"); import matplotlib.pyplot as plt
    data = json.load(open(os.path.join(base, "06_webceph", "steiner.json")))["mesures"]
    lab = {"SNA (°)":"SNA (°)","SNB (°)":"SNB (°)","ANB (°)":"ANB (°)","U1-NA (°)":"I-NA (°)","U1-NA (mm)":"I-NA (mm)",
           "L1-NB (°)":"i-NB (°)","L1-NB (mm)":"i-NB (mm)","Wits (mm)":"Wits (mm)","Pog-NB (mm)":"Pog-NB (mm)",
           "Plan mandibulaire (Go-Gn/SN) (°)":"Go-Gn / SN (°)","Plan occlusal/SN (°)":"Plan occlusal / SN (°)",
           "Angle interincisif (°)":"Angle inter-incisif (°)"}
    rows = list(data); N = len(rows)
    BLUE="#2F5CA8"; GREEN="#1E7D32"; RED="#C0392B"
    fig = plt.figure(figsize=(8.6, 0.42 * (N + 2)), dpi=150)
    ax = fig.add_axes([0, 0, 1, 1]); ax.axis("off"); ax.set_xlim(0, 1); ax.set_ylim(0, 1)
    top = 0.93; bot = 0.04; ys = [top - (top - bot) * (i + 1) / (N + 1) for i in range(N)]
    COLS = [(0.01, "Mesure", "left"), (0.44, "Moyenne", "center"),
            (0.57, "Écart-type", "center"), (0.70, "Résultat", "center"), (0.80, "Interprétation", "left")]
    ax.text(0.5, 0.975, "Analyse de Steiner", ha="center", va="center", fontsize=15, weight="bold", color=BLUE)
    hy = top - (top - bot) * 0.5 / (N + 1)
    for x, t, ha in COLS: ax.text(x, hy, t, ha=ha, va="center", fontsize=8.4, weight="bold", color="#222")
    ax.plot([0.005, 0.995], [hy - 0.018, hy - 0.018], color="#888", lw=0.8)
    for m, y in zip(rows, ys):
        mean, sd, res = m["moyenne"], m["ecart_type"], m["resultat"]
        out = bool(m.get("severite"))
        ax.text(0.01, y, lab.get(m["nom"], m["nom"]), ha="left", va="center", fontsize=8, color=BLUE)
        ax.text(0.44, y, f"{mean:g}", ha="center", va="center", fontsize=8, color="#333")
        ax.text(0.57, y, f"{sd:g}", ha="center", va="center", fontsize=8, color="#333")
        ax.text(0.70, y, f"{res:g}", ha="center", va="center", fontsize=8, color=(RED if out else GREEN), weight="bold")
        ax.text(0.80, y, m.get("signification") or "Dans la norme", ha="left", va="center",
                fontsize=7.6, color=(RED if out else GREEN))
    fig.savefig(os.path.join(base, "06_webceph", "steiner_graphique.png"), dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)

# ----------------------------------------------------------------------------
# 4) ASSEMBLAGE DU BILAN  (édition en place du modèle .docx)
# rId -> (fichier source relatif au dossier patient, dimensions cible)
_MEDIA = {
    "0afc091204a7df8f2ff99ddd4784b6ec5ab547e1.jpg": ("02_photos_traitees/exo_face_repos.jpg", (697,768)),
    "951cd8e02991963d36b9edd37db9fad200e94d7e.jpg": ("02_photos_traitees/exo_face_sourire.jpg", (697,768)),
    "aeeb8832c0f1d6bee03810da70bef8423ed097f6.jpg": ("02_photos_traitees/exo_profil.jpg", (758,768)),
    "c8580be4cbc5c738824f2ff1efc2b30b243f7d9b.jpg": ("02_photos_traitees/endo_occlusal_maxillaire.jpg", (861,645)),
    "59b0becb1a9147213871efb27f0d62a1bcfafa6e.jpg": ("02_photos_traitees/endo_occlusal_mandibulaire.jpg", (861,645)),
    "3f04a16d4efe799a1de0e9df5b0eab12cb69888e.jpg": ("02_photos_traitees/endo_laterale_droite.jpg", (860,599)),
    "0c3a43f5a8f493700f9edc6c5686d239e4c93bdc.jpg": ("02_photos_traitees/endo_laterale_gauche.jpg", (860,599)),
    "c2a48fecf384fec186f75982bc17a12924250f5b.jpg": ("02_photos_traitees/endo_occlusion_frontale.jpg", (819,599)),
    "5ac3c59ca092a8ac4ac9ce2614f5667de997ab61.png": ("05_stl_rendus/upper_occlusal.png", (1100,1047)),
    "e5bb2fba0f127090e814382a36e543cf9b0e0c55.png": ("05_stl_rendus/lower_occlusal.png", (1100,1003)),
    "e25d01e8b430f5fc609fabf3695feb64699c767e.png": ("05_stl_rendus/right_occ.png", (1100,683)),
    "fed5c52fe9e4586bed8ce1c8fffea7d859460347.png": ("05_stl_rendus/left_occ.png", (1100,603)),
    "9a92d772e65948c2fafd2d95b6f4c8abcbe0b8dd.png": ("05_stl_rendus/front_occ.png", (1100,614)),
    "1bcea7ef245074d7b6e8ef85cf56e355888b4483.png": ("05_stl_rendus/spee_lower.png", (1100,578)),
    "cf7f64d963f1544894f85c3ba30c736d9c012501.jpg": ("03_radios/panoramique.jpg", (2551,1470)),
    "456fdc28f5e27ebe5849e88df6d8508f4be06286.jpg": ("03_radios/teleradiographie_profil.jpg", (2926,2174)),
    "6332fdb5101edc7f7ec9a61fcedd51012700e3bc.png": ("06_webceph/steiner_graphique.png", (1327,1140)),
}
# Table Steiner (T9) : libellé de ligne -> nom dans steiner.json
_T9_MAP = {
    "SNA (°)":"SNA (°)", "SNB (°)":"SNB (°)", "ANB (°)":"ANB (°)", "AoBo (mm)":"Wits (mm)",
    "I-NA (mm)":"U1-NA (mm)", "I-NA (°)":"U1-NA (°)", "i-NB  (mm)":"L1-NB (mm)", "i-NB  (°)":"L1-NB (°)",
    "Pog-NB (mm)":"Pog-NB (mm)", "i-I  (°)":"Angle interincisif (°)", "O-SN (°)":"Plan occlusal/SN (°)",
    "GoGn-SN (°)":"Plan mandibulaire (Go-Gn/SN) (°)",
}
NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

def _fit(src, dims, is_png):
    W, H = dims; im = Image.open(src).convert("RGB")
    im2 = im.copy(); im2.thumbnail((W, H), Image.LANCZOS)
    canvas = Image.new("RGB", (W, H), "white")
    canvas.paste(im2, ((W-im2.width)//2, (H-im2.height)//2))
    buf = io.BytesIO()
    canvas.save(buf, "PNG") if is_png else canvas.save(buf, "JPEG", quality=92)
    return buf.getvalue()

def build_bilan(base, template_docx, info):
    """info: dict avec nom, dob_court ('02/2016'), age ('10 ans'), sexe ('féminin'),
       synthese: {classe_sagittale, divergence, incisives}"""
    from docx import Document
    sjson = os.path.join(base, "06_webceph", "steiner.json")
    steiner = json.load(open(sjson))["mesures"] if os.path.exists(sjson) else []
    by_nom = {m["nom"]: m for m in steiner}
    import tempfile
    tmpdir = tempfile.mkdtemp(prefix="bilanodf_")   # dossier LOCAL (hors iCloud)
    out_name = "Bilan_%s.docx" % info["nom"].replace(" ", "_")
    out_docx = os.path.join(base, out_name)
    out_tmp = os.path.join(tmpdir, out_name)         # docx final construit en local
    tmp = os.path.join(tmpdir, "_work.docx")
    # On OUVRE le modèle directement (lecture seule OK) et on enregistre vers un
    # fichier NEUF inscriptible — évite de recopier des permissions read-only.
    d = Document(template_docx)
    def tlist(el): return list(el.iter(NS + 't'))
    # titre
    firstp = d.element.body.find(NS + 'p')
    tl = tlist(firstp)
    if tl: tl[0].text = "%s, %s (%s)" % (info["nom"], info["dob_court"], info["age"])
    # identité (T7 = tables[6])
    t7 = d.tables[6]
    tlist(t7.rows[0].cells[0]._tc)[0].text = info["nom"]
    tlist(t7.rows[0].cells[1]._tc)[0].text = " Sexe %s — %s" % (info["sexe"], info["age"])
    # valeurs Steiner (T9 = tables[8])
    t9 = d.tables[8]
    for r in t9.rows:
        lbl = r.cells[0].text.strip()
        if lbl in _T9_MAP:                       # ligne connue -> valeur saisie ou "-" (pas de fuite du modèle)
            ts = tlist(r.cells[1]._tc)
            if ts:
                nom = _T9_MAP[lbl]
                ts[0].text = ("%g" % by_nom[nom]["resultat"]) if nom in by_nom else "-"
                for e in ts[1:]: e.text = ""
    # synthèse diagnostique — grille complète T15 = tables[14]
    t15 = d.tables[14]
    syn = dict(info.get("synthese", {}))
    syn.setdefault("sq_ap", syn.get("classe_sagittale", ""))
    syn.setdefault("sq_ve", syn.get("divergence", ""))
    syn.setdefault("al_ap", syn.get("incisives", ""))
    CELLS = {"sq_ap": (1, 1), "al_ap": (1, 2), "oc_ap": (1, 3), "cu_ap": (1, 4),
             "sq_tr": (2, 1), "al_tr": (2, 2), "oc_tr": (2, 3), "cu_tr": (2, 4),
             "sq_ve": (3, 1), "al_ve": (3, 2), "oc_ve": (3, 3), "cu_ve": (3, 4)}
    def set_cell(tc, text):
        ts = tlist(tc)
        if not ts: return
        ts[0].text = " / ".join(str(text).split("\n"))
        for e in ts[1:]: e.text = ""
    for k, (rr, cc) in CELLS.items():
        set_cell(t15.rows[rr].cells[cc]._tc, syn.get(k, ""))   # vide si non fourni
    # motif de consultation (T14 = tables[13])
    t14cell = d.tables[13].rows[0].cells[0]._tc
    for t in tlist(t14cell):
        if (t.text or "").strip() == "fin": t.text = ""
    motif = info.get("motif", "")
    if motif:
        for t in tlist(t14cell):
            if "Motif de consultation" in (t.text or ""):
                t.text = "Motif de consultation : " + motif; break
    # --- Plan de traitement : tableau d'analyse d'espace (T09 = tables[9]) ---
    espace = info.get("espace", {}) or {}
    T09_ROWS = {"encombrement": 1, "repositionnement_incisive": 2, "courbe_spee": 3, "_ddm": 4,
                "repositionnement_6": 5, "espace_derive_mesiale": 6, "_total": 9,
                "extraction_stripping": 10, "individualisation": 11, "_net": 12}
    try:
        t_esp = d.tables[9]
        for k, ri in T09_ROWS.items():
            dd = espace.get(k)
            if not isinstance(dd, dict): continue
            p = float(dd.get("p", 0) or 0); m = float(dd.get("m", 0) or 0)
            t_esp.rows[ri].cells[1].text = ("%g" % p) if abs(p) > 1e-9 else ""   # colonne +
            t_esp.rows[ri].cells[2].text = ("%g" % m) if abs(m) > 1e-9 else ""   # colonne −
    except Exception:
        pass
    # --- Stratégies et moyens : Objectifs / Moyens (T16 = tables[16]) ---
    plan = info.get("plan", {}) or {}
    obj = (plan.get("objectifs") or "").strip(); moy = (plan.get("moyens") or "").strip()
    if obj or moy:
        try:
            t_str = d.tables[16]
            cell = None
            for rr in t_str.rows:
                for cc in rr.cells:
                    if "Objectifs" in cc.text:
                        cell = cc; break
                if cell is not None: break
            if cell is not None:
                cell.text = ""                                   # repart d'un paragraphe vide
                cell.paragraphs[0].add_run("Objectifs :").bold = True
                for line in (obj.split("\n") if obj else [""]):
                    cell.add_paragraph(line)
                pm = cell.add_paragraph(); pm.add_run("Moyens :").bold = True
                for line in (moy.split("\n") if moy else [""]):
                    cell.add_paragraph(line)
        except Exception:
            pass
    d.save(tmp)
    # remplacement des images
    newbytes = {}
    for fn, (rel, dims) in _MEDIA.items():
        src = os.path.join(base, rel)
        if os.path.exists(src):
            newbytes["word/media/" + fn] = _fit(src, dims, fn.endswith(".png"))
    zin = zipfile.ZipFile(tmp); zout = zipfile.ZipFile(out_tmp, "w", zipfile.ZIP_DEFLATED)
    for it in zin.infolist():
        data = zin.read(it.filename)
        if it.filename in newbytes: data = newbytes[it.filename]
        zout.writestr(it, data)
    zin.close(); zout.close()
    # PDF via LibreOffice (converti DANS le dossier temporaire local)
    soffice = shutil.which("libreoffice") or shutil.which("soffice") or \
              "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if os.path.exists(soffice) or shutil.which(os.path.basename(soffice)):
        try:
            subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", tmpdir, out_tmp],
                           timeout=240, check=False,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception:
            pass
    pdf_tmp = out_tmp[:-5] + ".pdf"
    # on retire d'éventuels anciens livrables (parfois en lecture seule) avant de recopier
    for pth in (out_docx, out_docx[:-5] + ".pdf"):
        try:
            if os.path.exists(pth):
                os.chmod(pth, 0o644); os.remove(pth)
        except Exception:
            pass
    shutil.copy(out_tmp, out_docx)
    try: os.chmod(out_docx, 0o644)
    except Exception: pass
    pdf_final = None
    if os.path.exists(pdf_tmp):
        pdf_final = out_docx[:-5] + ".pdf"; shutil.copy(pdf_tmp, pdf_final)
        try: os.chmod(pdf_final, 0o644)
        except Exception: pass
    # aperçu des pages (PNG) pour affichage direct dans l'app
    apercu = os.path.join(base, "apercu")
    if os.path.isdir(apercu):
        for f in os.listdir(apercu):
            try: os.remove(os.path.join(apercu, f))
            except Exception: pass
    if pdf_final:
        try: render_pdf_preview(pdf_final, apercu)
        except Exception: pass
    shutil.rmtree(tmpdir, ignore_errors=True)
    return out_docx, pdf_final

def render_pdf_preview(pdf_path, out_dir, dpi=110):
    """Rend chaque page du PDF en PNG (pour l'aperçu dans l'app). Nécessite PyMuPDF (optionnel)."""
    try:
        import fitz  # PyMuPDF
    except Exception:
        return []
    os.makedirs(out_dir, exist_ok=True)
    doc = fitz.open(pdf_path); out = []
    for i, page in enumerate(doc):
        pix = page.get_pixmap(dpi=dpi)
        fn = os.path.join(out_dir, "page_%02d.png" % (i + 1)); pix.save(fn); out.append(fn)
    doc.close(); return out

def make_steiner_json(base, resultats, significations=None, severites=None):
    """resultats: {nom: valeur}. Construit 06_webceph/steiner.json avec les normes standard."""
    significations = significations or {}
    severites = severites or {}
    mes = []
    for nom, mean, sd in STEINER_NORMS:
        if nom not in resultats or resultats[nom] in (None, ""):
            continue
        mes.append({"nom": nom, "moyenne": mean, "ecart_type": sd,
                    "resultat": float(resultats[nom]),
                    "severite": severites.get(nom, ""),
                    "signification": significations.get(nom, "")})
    os.makedirs(os.path.join(base, "06_webceph"), exist_ok=True)
    json.dump({"mesures": mes}, open(os.path.join(base, "06_webceph", "steiner.json"), "w"),
              ensure_ascii=False, indent=1)

# ----------------------------------------------------------------------------
# 5) CLASSIFICATION LOCALE DES FICHIERS (import intelligent)
#    Renvoie un jeton "photo:<slot>" | "radio:<slot>" | "stl:<slot>" | "".
#    Type (photo/radio/stl) fiable ; la VUE photo est une meilleure estimation
#    (à confirmer par l'utilisateur dans l'écran de vérification).
# ============================================================================
#  Reconnaissance des vues photo — modèle AUTO-APPRIS sur les photos déjà rangées
#  (aucune donnée à copier : le nom de fichier 01_photos_brutes/<vue>.jpg = étiquette)
# ============================================================================
def extract_features(path):
    """Vecteur de caractéristiques d'une photo (couleurs, muqueuse/dents/peau,
    symétries signées, et grille spatiale 4x4) — robuste pour distinguer les vues."""
    import numpy as np
    im = ImageOps.exif_transpose(Image.open(path)).convert("RGB")
    W, H = im.size; aspect = W / max(H, 1)
    a = np.asarray(im.resize((96, 96))).astype(float) / 255.0
    R, G, B = a[..., 0], a[..., 1], a[..., 2]
    mx = a.max(-1); mn = a.min(-1); sat = (mx - mn) / np.clip(mx, 1e-6, 1)
    gray = a.mean(-1); h, w = a.shape[:2]
    colored = float((sat > 0.15).mean())
    skin = float(((R > 0.4) & (R >= G) & (G >= B) & ((R - B) > 0.05) & (sat < 0.5)).mean())
    muc = float(((R > 0.4) & ((R - G) > 0.12) & ((R - B) > 0.12) & (sat > 0.25)).mean())
    teeth = (mx > 0.6) & (sat < 0.18); teethf = float(teeth.mean())
    asym = float(np.abs(gray - gray[:, ::-1]).mean())
    vasym = float(np.abs(gray - gray[::-1, :]).mean())
    cy0, cy1, cx0, cx1 = int(h * .3), int(h * .7), int(w * .3), int(w * .7)
    cmuc = float((((R > 0.4) & ((R - G) > 0.12) & (sat > 0.25)))[cy0:cy1, cx0:cx1].mean())
    cteeth = float(teeth[cy0:cy1, cx0:cx1].mean())
    bm = np.ones((h, w), bool); bm[cy0:cy1, cx0:cx1] = False; bteeth = float(teeth[bm].mean())
    lr_bright = float(gray[:, :w // 2].mean() - gray[:, w // 2:].mean())   # gauche - droite (signé)
    lr_teeth = float(teeth[:, :w // 2].mean() - teeth[:, w // 2:].mean())
    scal = [aspect, colored, skin, muc, teethf, asym, vasym, cmuc, cteeth, bteeth, lr_bright, lr_teeth]
    grid = []
    gs = 4
    for gy in range(gs):
        for gx in range(gs):
            ys = slice(gy * h // gs, (gy + 1) * h // gs); xs = slice(gx * w // gs, (gx + 1) * w // gs)
            grid += [float(R[ys, xs].mean()), float(G[ys, xs].mean()),
                     float(B[ys, xs].mean()), float(sat[ys, xs].mean())]
    return np.array(scal + grid, float)

def _loo_accuracy(Xs, y, k=5):
    import numpy as np
    from collections import defaultdict
    n = len(y); correct = 0
    for i in range(n):
        d = np.sqrt(((Xs - Xs[i]) ** 2).sum(1)); d[i] = 1e18
        idx = np.argsort(d)[:k]; sc = defaultdict(float)
        for j in idx: sc[y[j]] += 1.0 / (d[j] + 1e-6)
        if max(sc.items(), key=lambda kv: kv[1])[0] == y[i]: correct += 1
    return round(correct / max(n, 1), 3)

def train_photo_model(patients_root=None, model_path=None):
    """Apprend un modèle de reconnaissance des vues à partir des photos DÉJÀ rangées
    dans 01_photos_brutes de tous les patients. Retourne un compte-rendu."""
    import numpy as np
    from collections import Counter
    patients_root = patients_root or os.path.join(PDATA, "patients")
    model_path = model_path or MODEL_PATH
    X, y = [], []
    if os.path.isdir(patients_root):
        for slug in os.listdir(patients_root):
            recs = os.path.join(patients_root, slug, "records")
            if not os.path.isdir(recs): continue
            for rid in os.listdir(recs):
                raw = os.path.join(recs, rid, "01_photos_brutes")
                if not os.path.isdir(raw): continue
                for fn in os.listdir(raw):
                    if not fn.lower().endswith(".jpg"): continue
                    slot = os.path.splitext(fn)[0]
                    try: X.append(extract_features(os.path.join(raw, fn))); y.append("photo:" + slot)
                    except Exception: pass
    if len(X) < 8:
        return {"ok": False, "n": len(X), "msg": "Pas assez de photos déjà rangées (min. 8) pour calibrer."}
    X = np.array(X); y = np.array(y)
    mean = X.mean(0); std = X.std(0); std[std < 1e-6] = 1.0
    Xs = (X - mean) / std
    acc = _loo_accuracy(Xs, y, k=min(5, max(1, len(y) // 4)))
    np.savez(model_path, X=Xs.astype("float32"), y=y, mean=mean, std=std)
    global _MODEL, _MODEL_MTIME; _MODEL = None; _MODEL_MTIME = None   # recharge au prochain usage
    return {"ok": True, "n": int(len(y)), "acc": acc,
            "counts": dict(Counter([c.split(":", 1)[1] for c in y.tolist()]))}

_MODEL = None; _MODEL_MTIME = None
def _load_model(model_path=None):
    global _MODEL, _MODEL_MTIME
    import numpy as np
    model_path = model_path or MODEL_PATH
    if not os.path.exists(model_path): return None
    mt = os.path.getmtime(model_path)
    if _MODEL is None or _MODEL_MTIME != mt:
        d = np.load(model_path, allow_pickle=True)
        _MODEL = {"X": d["X"].astype(float), "y": d["y"], "mean": d["mean"], "std": d["std"]}
        _MODEL_MTIME = mt
    return _MODEL

def _knn_predict(feat, model, k=5):
    import numpy as np
    from collections import defaultdict
    fs = (feat - model["mean"]) / model["std"]
    d = np.sqrt(((model["X"] - fs) ** 2).sum(1))
    idx = np.argsort(d)[:min(k, len(d))]; sc = defaultdict(float)
    for j in idx: sc[model["y"][j]] += 1.0 / (d[j] + 1e-6)
    return max(sc.items(), key=lambda kv: kv[1])[0]

def guess_category(path, filename=""):
    name = (filename or os.path.basename(path)).lower()
    if name.endswith((".stl", ".ply", ".obj")):
        if any(w in name for w in ["lower", "mand", "inf", "bas", "low"]):
            return "stl:LowerJawScan"
        return "stl:UpperJawScan"
    # modèle auto-appris prioritaire (si l'utilisateur a calibré)
    try:
        model = _load_model()
        if model is not None:
            feat = extract_features(path)
            if feat[1] < 0.04:      # fraction de couleur quasi nulle -> radiographie
                return "radio:panoramique" if feat[0] > 1.5 else "radio:teleradiographie_profil"
            return _knn_predict(feat, model)
    except Exception:
        pass
    return _guess_heuristic(path, filename)

def _guess_heuristic(path, filename=""):
    name = (filename or os.path.basename(path)).lower()
    if name.endswith((".stl", ".ply", ".obj")):
        if any(w in name for w in ["lower", "mand", "inf", "bas", "low"]):
            return "stl:LowerJawScan"
        return "stl:UpperJawScan"
    try:
        import numpy as np
        im = ImageOps.exif_transpose(Image.open(path)).convert("RGB")
        W, H = im.size; aspect = W / max(H, 1)
        thumb = im.resize((160, max(1, int(160 * H / W)))) if W >= H else im.resize((max(1, int(160 * W / H)), 160))
        a = np.asarray(thumb).astype(float) / 255.0
        R, G, B = a[..., 0], a[..., 1], a[..., 2]
        mx = a.max(-1); mn = a.min(-1); sat = (mx - mn) / np.clip(mx, 1e-6, 1)
        h, w = a.shape[:2]
        colored = float((sat > 0.15).mean())        # fraction de pixels réellement colorés
        # --- radiographie : image quasi sans couleur (rayons X en niveaux de gris) ---
        if colored < 0.04:
            return "radio:panoramique" if aspect > 1.5 else "radio:teleradiographie_profil"
        skin = float(((R > 0.4) & (R >= G) & (G >= B) & ((R - B) > 0.05) & (sat < 0.5)).mean())
        muc = float(((R > 0.4) & ((R - G) > 0.12) & ((R - B) > 0.12) & (sat > 0.25)).mean())  # muqueuse rose/rouge
        teeth = (mx > 0.6) & (sat < 0.18)
        g = a.mean(-1); asym = float(np.abs(g - g[:, ::-1]).mean())
        cy0, cy1, cx0, cx1 = int(h * .3), int(h * .7), int(w * .3), int(w * .7)
        cmuc = float((((R > 0.4) & ((R - G) > 0.12) & (sat > 0.25))[cy0:cy1, cx0:cx1]).mean())  # centre muqueux
        cteeth = float(teeth[cy0:cy1, cx0:cx1].mean())
        bm = np.ones((h, w), bool); bm[cy0:cy1, cx0:cx1] = False
        bteeth = float(teeth[bm].mean())
        # --- extra-oral : très peu de pixels colorés (fond clair + peau désaturée) ---
        if colored < 0.30:
            if skin < 0.12:                          # profil : peau frontale réduite, forte asymétrie
                return "photo:exo_profil"
            return "photo:exo_face_repos"            # repos/sourire indiscernables -> défaut repos (à confirmer)
        # --- intra-oral ---
        if cmuc > 0.75:                              # vue occlusale : centre = palais/langue, dents en couronne
            return "photo:endo_occlusal_maxillaire" if bteeth > 0.10 else "photo:endo_occlusal_mandibulaire"
        if muc < 0.20:                               # peu de gencive exposée -> occlusion frontale
            return "photo:endo_occlusion_frontale"
        return "photo:endo_laterale_droite" if cteeth >= 0.12 else "photo:endo_laterale_gauche"
    except Exception:
        return ""

# ----------------------------------------------------------------------------
# Entrée ligne de commande : permet d'exécuter le rendu 3D dans un SOUS-PROCESSUS
# isolé (un crash OpenGL/VTK ne tue alors pas l'application).
if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == "render_stl":
        render_stl(sys.argv[2])
