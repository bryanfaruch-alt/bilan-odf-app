# -*- coding: utf-8 -*-
"""
Importation massive de dossiers patients (clé USB / disque) vers Bilan ODF.

Ce module ne fait QUE l'analyse (scan) d'une arborescence : il ne copie rien et
n'écrit rien dans les données de l'app. Il renvoie un « plan » par patient que
l'app affiche dans un récapitulatif, puis exécute (copie disque à disque) après
validation. Aucune donnée ne quitte la machine.

Structure attendue (tolérante aux variantes vues en vrai) :
  DOSSIER PATIENT « NOM Prénom [date] [✅] »
    ├── bilan NOM Prénom….docx              (identité + analyses)
    ├── MODELE/  (ou sous-dossier 1288…/)    STL + PNG de rendu (occlusion, empreinte, spee)
    ├── PHOTO/                               soit PHOTO 1, PHOTO 2… (= temps), soit photos en vrac
    ├── RADIO/                               radios (noms DICOM), ou RX/TRP en vrac à la racine
Cas « Word seul » : uniquement un .docx -> on extraira les images du Word (autre étape).
"""
import os
import re
import datetime
import unicodedata

PHOTO_EXTS = (".jpg", ".jpeg", ".png", ".heic", ".heif", ".tif", ".tiff", ".bmp", ".webp")
STL_EXTS = (".stl", ".ply", ".obj")

# Rendu de modèle : mots-clés du nom de fichier -> emplacement de rendu dans l'app.
RENDER_MATCHERS = [
    ("upper_occlusal", ["empreinte maxil", "empreinte max", "emp maxil", "emp max", "maxilaire", "maxillaire", "arcade sup", "upperjaw", "sup00", "max00"]),
    ("lower_occlusal", ["empreinte mand", "empreinte mendib", "emp mand", "emp mendib", "emp bas", "mandibulaire", "mendibulaire", "arcade inf", "lowerjaw"]),
    ("front_occ", ["occlusion anter", "occlusion ant", "occ anter", "occ ant", "anterieur", "anterieure", "frontal"]),
    ("right_occ", ["occlusion droite", "occ droite", "occ droit", "droite", "droit"]),
    ("left_occ", ["occlusion gauche", "occ gauche", "gauche"]),
    ("spee_lower", ["courbe de spee", "courbe spee", "spee"]),
]

RADIO_HINTS = ["pano", "trp", "teleradio", "tele radio", "teleradiographie", "telerad",
               "cephalo", "cephalom", "ceph", "rx pano", "rx tele", "rx ", "orthopan", "opt "]


def _norm(s):
    s = unicodedata.normalize("NFKD", s or "").encode("ascii", "ignore").decode().lower()
    return re.sub(r"\s+", " ", s).strip()


def _visible(names):
    return [e for e in names if not e.startswith(".") and not e.startswith("~$")]


def _ls(dirpath):
    try:
        return sorted(_visible(os.listdir(dirpath)))
    except Exception:
        return []


def _images_in(dirpath):
    return [os.path.join(dirpath, e) for e in _ls(dirpath)
            if e.lower().endswith(PHOTO_EXTS) and os.path.isfile(os.path.join(dirpath, e))]


def _find_subdir(folder, *names):
    want = {_norm(n) for n in names}
    for e in _ls(folder):
        p = os.path.join(folder, e)
        if os.path.isdir(p) and _norm(e) in want:
            return p
    return None


def render_key_for(filename):
    n = _norm(filename)
    for key, kws in RENDER_MATCHERS:
        if any(k in n for k in kws):
            return key
    return None


def looks_like_radio(filename):
    n = _norm(filename)
    if re.match(r"^\d+\.\d+\.\d", n):          # nom DICOM exporté : 1.76.380.18.18…
        return True
    return any(h in n for h in RADIO_HINTS)


def parse_identity_from_name(name):
    """« LAMRINI Mohamed Amine 22.01.25 ✅ » -> nom_famille/prenom/nom complet."""
    s = re.sub(r"\.docx?$", "", name, flags=re.I)
    s = re.sub(r"[✅✔☑️⭐\U0001F7E2]", " ", s)   # ✅ ✔ ✲ étoiles…
    s = re.sub(r"^\s*bilan\s+", "", s, flags=re.I)
    s = re.sub(r"\b\d{1,2}[./]\d{1,2}[./]\d{2,4}\b", " ", s)            # date 22.01.25
    s = re.sub(r"\b\d{1,2}\s*ans(\s+\d+)?\b", " ", s, flags=re.I)       # « 11 ans 26 »
    s = re.sub(r"[+_].*$", " ", s)                                      # « … + transferts »
    s = re.sub(r"\b\d+\b", " ", s)                                      # nombres résiduels
    s = re.sub(r"\s+", " ", s).strip(" -–—")
    nom_toks, pre_toks = [], []
    for t in s.split(" "):
        if not t:
            continue
        is_caps = t == t.upper() and re.search(r"[A-Z]", unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode())
        if is_caps and not pre_toks:
            nom_toks.append(t)
        else:
            pre_toks.append(t)
    nom = " ".join(nom_toks)
    prenom = " ".join(pre_toks)
    full = (nom + " " + prenom).strip()
    return {"nom_famille": nom, "prenom": prenom, "nom": full or s}


def _resolve_render_keys(renders):
    """Attribue les emplacements de rendu ; 2e courbe de spee -> spee_left ;
    doublons et non reconnus -> clé None (capture « extra »)."""
    spees = sorted([fp for rk, fp in renders if rk == "spee_lower"], key=lambda p: _norm(os.path.basename(p)))
    spee_map = {}
    if len(spees) >= 1: spee_map[spees[0]] = "spee_lower"
    if len(spees) >= 2: spee_map[spees[1]] = "spee_left"
    used, out = set(), []
    for rk, fp in renders:
        if rk == "spee_lower":
            rk = spee_map.get(fp)
        if rk and rk in used:
            rk = None
        if rk:
            used.add(rk)
        out.append((rk, fp))
    return out


def _photo_timepoints(photo_dir):
    """Renvoie la liste ordonnée des temps : [{label, dir, photos[]}]."""
    subdirs = [e for e in _ls(photo_dir) if os.path.isdir(os.path.join(photo_dir, e))]
    numbered = []
    for e in subdirs:
        m = re.search(r"(\d+)", e)
        if m and (_norm(e).startswith("photo") or _norm(e).startswith("reev") or _norm(e).startswith("temps")):
            numbered.append((int(m.group(1)), e))
    tps = []
    if numbered:
        numbered.sort()
        for i, (_num, e) in enumerate(numbered):
            imgs = _images_in(os.path.join(photo_dir, e))
            tps.append({"label": "Bilan initial" if i == 0 else "Réévaluation %d" % i,
                        "dir": os.path.join(photo_dir, e), "photos": imgs})
        return tps
    # pas de sous-dossiers numérotés : soit des images en vrac, soit des sous-dossiers quelconques
    direct = _images_in(photo_dir)
    if direct:
        return [{"label": "Bilan initial", "dir": photo_dir, "photos": direct}]
    img_subdirs = [(e, _images_in(os.path.join(photo_dir, e))) for e in subdirs]
    img_subdirs = [(e, im) for e, im in img_subdirs if im]
    for i, (e, im) in enumerate(sorted(img_subdirs)):
        tps.append({"label": "Bilan initial" if i == 0 else "Réévaluation %d" % i,
                    "dir": os.path.join(photo_dir, e), "photos": im})
    return tps


def _files_under(folder):
    out = []
    for dp, _d, fs in os.walk(folder):
        for f in _visible(fs):
            out.append(os.path.join(dp, f))
    return out


def _photo_num(name):
    """Numéro d'un dossier/temps : 'PHOTO 3' -> 3 ; sinon None."""
    m = re.search(r"(\d+)", name)
    return int(m.group(1)) if m else None


def _mtime_date(path):
    try:
        return datetime.date.fromtimestamp(os.path.getmtime(path))
    except Exception:
        return None

_YYMMDD = re.compile(r"(2[0-9])(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])")

def _exif_date(path):
    """Date de PRISE DE VUE depuis l'EXIF (fiable, survit à la copie sur la clé,
    contrairement au mtime qui vaut souvent la date de copie)."""
    try:
        from PIL import Image
        ex = Image.open(path)._getexif() or {}
        v = ex.get(36867) or ex.get(36868) or ex.get(306)   # DateTimeOriginal / Digitized / DateTime
        if v and len(str(v)) >= 10:
            v = str(v)
            return datetime.date(int(v[0:4]), int(v[5:7]), int(v[8:10]))
    except Exception:
        pass
    return None

def _photo_date(path):
    """EXIF d'abord (date réelle), puis mtime en dernier recours."""
    return _exif_date(path) or _mtime_date(path)

def _model_date(path):
    """Date d'un modèle : dossier de scan daté (…_AAMMJJhhmmss) dans le chemin, sinon mtime."""
    for yy, mm, dd in reversed(_YYMMDD.findall(path)):
        try:
            return datetime.date(2000 + int(yy), int(mm), int(dd))
        except Exception:
            continue
    return _mtime_date(path)

def _dicom_date(path):
    """Date d'examen depuis un nom DICOM (…AAMMJJhhmmss…). On prend la DERNIÈRE date
    plausible : le préfixe DICOM contient un identifiant machine fixe (…1210429141329…)
    qui ressemble à une date ; la vraie date d'examen vient après."""
    ms = _YYMMDD.findall(os.path.basename(path))
    for yy, mm, dd in reversed(ms):
        try:
            return datetime.date(2000 + int(yy), int(mm), int(dd))
        except Exception:
            continue
    return None

def _cluster_by_date(items, gap_days=25):
    """items = [(path, date|None)] -> [(date_repr|None, [paths])] triés du + ancien au + récent.
    Une nouvelle séance démarre quand l'écart avec la photo précédente dépasse gap_days.
    Les photos d'appareils différents prises le même jour restent donc dans la même séance."""
    dated = sorted([(p, d) for p, d in items if d], key=lambda x: x[1])
    undated = [p for p, d in items if not d]
    clusters = []
    cur, start, last = [], None, None
    for p, d in dated:
        if not cur:
            cur, start, last = [p], d, d
        elif (d - last).days <= gap_days:
            cur.append(p); last = d
        else:
            clusters.append((start, cur)); cur, start, last = [p], d, d
    if cur:
        clusters.append((start, cur))
    if undated:
        if clusters:
            clusters[0] = (clusters[0][0], clusters[0][1] + undated)
        else:
            clusters = [(None, undated)]
    return clusters


def scan_patient(folder):
    """Détection PAR CONTENU + regroupement des temps PAR DATE (et non par dossier) :
    - MODÈLE = dossier contenant un .stl/.ply/.obj/.3ox (nom quelconque : 'MODELE', '1288…', 'MOULAGE'…) ; ses images = rendus.
    - RADIO = nom DICOM (1.76.380…) / 'PANO/TRP/téléradio/cephalo…' / dossier 'radio…'.
    - PHOTOS cliniques = le reste ; on les REGROUPE PAR DATE (mtime) : les dossiers 'PHOTO 1/2'
      séparent souvent les APPAREILS (Panasonic/iPhone) d'une même séance, pas les temps.
      La séance la plus ancienne = Bilan initial, les suivantes = Réévaluation 1, 2…
    Radios et modèles sont rattachés à la séance la plus proche par leur date."""
    warnings = []
    entries = _ls(folder)
    docx = [e for e in entries if e.lower().endswith(".docx")]
    docx.sort(key=lambda e: (0 if _norm(e).startswith("bilan") else 1, len(e)))
    word_path = os.path.join(folder, docx[0]) if docx else None

    stl, renders, radios, clinical = [], [], [], []
    subdirs = [e for e in entries if os.path.isdir(os.path.join(folder, e))]

    for e in subdirs:
        sp = os.path.join(folder, e); n = _norm(e)
        files = _files_under(sp)
        stls = [f for f in files if f.lower().endswith(STL_EXTS)]
        scan_folder = bool(re.match(r"^\d{6,}_\d", e)) or any(
            re.match(r"^\d{6,}_\d", os.path.basename(os.path.dirname(f))) for f in files)
        has_scan = bool(stls) or scan_folder or any(f.lower().endswith((".3ox", ".ox", ".dcm")) for f in files)
        imgs = [f for f in files if f.lower().endswith(PHOTO_EXTS)]
        if has_scan or "modele" in n or "modèle" in n or "moulage" in n or "empreinte" in n:   # -> MODÈLE
            stl += stls
            renders += [(render_key_for(os.path.basename(f)), f) for f in imgs]
            continue
        dicom_ratio = (sum(1 for f in imgs if looks_like_radio(os.path.basename(f))) / len(imgs)) if imgs else 0
        if n.startswith("radio") or n.startswith("rx") or (imgs and dicom_ratio >= 0.5):   # -> RADIO
            radios += imgs
            continue
        # tout autre dossier porteur d'images (PHOTO, PHOTO 1/2, 'X photos', vrac…) -> photos cliniques
        clinical += [f for f in imgs if not looks_like_radio(os.path.basename(f))]

    # fichiers en vrac à la racine
    for e in entries:
        p = os.path.join(folder, e)
        if not os.path.isfile(p):
            continue
        low = e.lower()
        if low.endswith(STL_EXTS):
            stl.append(p)
        elif low.endswith(PHOTO_EXTS):
            (radios if looks_like_radio(e) else clinical).append(p)

    renders = _resolve_render_keys(renders)

    # --- temps d'intervention PAR DATE (EXIF de prise de vue des photos cliniques) ---
    photo_clusters = _cluster_by_date([(p, _photo_date(p)) for p in clinical])
    timepoints = []
    for i, (d, paths) in enumerate(photo_clusters):
        timepoints.append({"label": "Bilan initial" if i == 0 else "Réévaluation %d" % i,
                           "dir": folder, "photos": paths, "date": (d.isoformat() if d else None),
                           "radios": [], "stl": [], "renders": []})
    # pas de photo mais radios/modèles -> créer les temps depuis leurs dates
    if not timepoints:
        seed = [(p, _dicom_date(p) or _mtime_date(p)) for p in radios] or [(p, _model_date(p)) for p in stl]
        for i, (d, _paths) in enumerate(_cluster_by_date(seed) if seed else []):
            timepoints.append({"label": "Bilan initial" if i == 0 else "Réévaluation %d" % i,
                               "dir": folder, "photos": [], "date": (d.isoformat() if d else None),
                               "radios": [], "stl": [], "renders": []})
    if not timepoints:
        timepoints = [{"label": "Bilan initial", "dir": folder, "photos": [],
                       "date": None, "radios": [], "stl": [], "renders": []}]

    # --- rattacher radios / stl / rendus à la séance la plus proche par date ---
    def _nearest(d):
        if d is None:
            return 0
        best, bestdiff = 0, None
        for i, tp in enumerate(timepoints):
            td = tp.get("date")
            try:
                diff = abs((datetime.date.fromisoformat(td) - d).days) if td else 10 ** 6
            except Exception:
                diff = 10 ** 6
            if bestdiff is None or diff < bestdiff:
                best, bestdiff = i, diff
        return best
    for p in radios:
        timepoints[_nearest(_dicom_date(p) or _mtime_date(p))]["radios"].append(p)
    for p in stl:
        timepoints[_nearest(_model_date(p))]["stl"].append(p)
    for rk, p in renders:
        timepoints[_nearest(_model_date(p))]["renders"].append((rk, p))

    # ne garder que les temps réellement porteurs de contenu (sinon 1 temps vide autorisé)
    tp_nonempty = [t for t in timepoints if t["photos"] or t["radios"] or t["stl"] or t["renders"]]
    if tp_nonempty:
        timepoints = tp_nonempty
        for i, t in enumerate(timepoints):
            t["label"] = "Bilan initial" if i == 0 else "Réévaluation %d" % i

    has_files = bool(any(t["photos"] for t in timepoints) or radios or stl or renders)
    mode = "files" if has_files else ("word_only" if word_path else "empty")
    if mode == "empty":
        warnings.append("Aucun fichier exploitable (ni photo, ni radio, ni modèle, ni Word).")
    elif mode == "files" and not any(t["photos"] for t in timepoints):
        warnings.append("Aucune photo trouvée (radios/modèles seulement).")

    return {"folder": folder, "ident": parse_identity_from_name(os.path.basename(folder)),
            "word_path": word_path, "mode": mode, "timepoints": timepoints,
            "radios": radios, "stl": stl, "renders": renders, "warnings": warnings}


_NON_PATIENT_DIRS = {"modele", "modeles", "photo", "photos", "radio", "radios", "rx"}


def looks_like_patient_folder(folder):
    """Un dossier patient contient un bilan Word OU un sous-dossier MODELE/PHOTO.
    (Un dossier RADIO / MODELE / PHOTO seul n'est PAS un patient.)"""
    entries = _ls(folder)
    if any(e.lower().endswith(".docx") for e in entries):
        return True
    if _find_subdir(folder, "modele", "modèle", "modeles", "modèles", "photo", "photos"):
        return True
    return False


def _word_only_plan(path, name=None):
    return {"folder": path, "ident": parse_identity_from_name(name or os.path.basename(path)),
            "word_path": path, "mode": "word_only", "timepoints": [], "radios": [],
            "stl": [], "renders": [], "warnings": []}


def scan_source(root):
    """Analyse une racine et renvoie un plan PAR PATIENT.

    Robuste au niveau pointé : que l'on désigne le dossier parent, un niveau de
    regroupement au-dessus, ou directement UN dossier patient, on obtient toujours
    un plan par patient — jamais un « patient » MODELE / PHOTO / RADIO.
    """
    subdirs = [os.path.join(root, e) for e in _ls(root) if os.path.isdir(os.path.join(root, e))]
    child_patients = [d for d in subdirs
                      if _norm(os.path.basename(d)) not in _NON_PATIENT_DIRS and looks_like_patient_folder(d)]
    has_own_clinical = _find_subdir(root, "modele", "modèle", "modeles", "modèles", "photo", "photos") is not None
    # racine = UN seul patient uniquement si elle a sa propre structure clinique
    # (MODELE/PHOTO) ET aucun sous-dossier patient. Un simple .docx à la racine ne
    # doit pas faire prendre tout le dossier pour un seul patient.
    if has_own_clinical and not child_patients:
        return [scan_patient(root)]
    if not child_patients and not has_own_clinical:
        entries = _ls(root)
        docx_at_root = [e for e in entries if e.lower().endswith(".docx")]
        has_subdir = any(os.path.isdir(os.path.join(root, e)) for e in entries)
        if docx_at_root and not has_subdir:
            # UN seul .docx -> dossier d'un patient « Word seul » (peut aussi contenir
            # des photos/radios en vrac -> scan_patient les récupère).
            if len(docx_at_root) == 1:
                return [scan_patient(root)]
            # PLUSIEURS .docx en vrac dans un même dossier -> un patient « Word seul »
            # PAR fichier (cas d'un dossier « Bilans » regroupant tous les bilans).
            return [_word_only_plan(os.path.join(root, e), e) for e in docx_at_root]
    return _scan_dir(root, 0)


def _scan_dir(root, depth):
    plans = []
    for e in _ls(root):
        p = os.path.join(root, e)
        if os.path.isdir(p):
            if _norm(e) in _NON_PATIENT_DIRS:        # jamais un patient
                continue
            if looks_like_patient_folder(p):
                plans.append(scan_patient(p))
            elif depth < 2:                          # niveau de regroupement -> on descend
                plans += _scan_dir(p, depth + 1)
        elif e.lower().endswith(".docx"):            # bilan Word seul posé à ce niveau
            plans.append(_word_only_plan(p, e))
    return plans


# ----------------------------------------------------------------------------
#  MODE « WORD SEUL » : extraction des images intégrées au .docx
# ----------------------------------------------------------------------------
# Gabarit constaté (3 bilans réels) : 10 images de contenu, ordre fixe =
#   1 latérale droite · 2 latérale gauche · 3 occlusion frontale ·
#   4 occlusal maxillaire · 5 occlusal mandibulaire ·
#   6 profil · 7 visage sourire · 8 visage repos · 9 téléradio · 10 panoramique.
# Les petits logos du modèle (emf + images < 8 Ko) sont écartés.
WORD_VIEW_ORDER = ["endo_laterale_droite", "endo_laterale_gauche", "endo_occlusion_frontale",
                   "endo_occlusal_maxillaire", "endo_occlusal_mandibulaire",
                   "exo_profil", "exo_face_sourire", "exo_face_repos"]
WORD_LOGO_MIN_SIZE = 8000   # octets : en dessous = logo/décoration du gabarit


def _media_entries(docx_path):
    import zipfile
    out = []
    with zipfile.ZipFile(docx_path) as z:
        for n in z.namelist():
            if n.startswith("word/media/") and not n.lower().endswith(".emf"):
                m = re.search(r"(\d+)", os.path.basename(n))
                out.append((int(m.group(1)) if m else 0, n, z.getinfo(n).file_size))
    out.sort()
    return out


def count_word_media(docx_path):
    """Nombre d'images de contenu (hors logos) — pour le récapitulatif, sans décoder."""
    try:
        return sum(1 for _i, _n, sz in _media_entries(docx_path) if sz >= WORD_LOGO_MIN_SIZE)
    except Exception:
        return 0


# Ordre des images dans les deux gabarits de bilan Word constatés :
#  - gabarit COMPLET (>=16 images) : 3 photos de visage, 5 intra-orales, 6 rendus, 2 radios.
#  - gabarit COURT « Word seul » (10 images) : 5 intra-orales, 3 visages, 2 radios.
WORD_ORDER_16 = ["exo_face_sourire", "exo_profil", "exo_face_repos", "endo_occlusion_frontale",
                 "endo_laterale_droite", "endo_laterale_gauche",
                 "endo_occlusal_maxillaire", "endo_occlusal_mandibulaire"]
WORD_ORDER_10 = ["endo_laterale_droite", "endo_laterale_gauche", "endo_occlusion_frontale",
                 "endo_occlusal_maxillaire", "endo_occlusal_mandibulaire",
                 "exo_profil", "exo_face_sourire", "exo_face_repos"]


def word_photo_plan(paths):
    """Mappe positionnellement les images d'un bilan Word (ordonnées) vers les 8 vues.
    Renvoie (photos{vue: chemin}, radios[2 chemins], autres[chemins de rendus]) ou None
    si le nombre d'images ne correspond à aucun gabarit connu (-> classer autrement)."""
    n = len(paths)
    if n >= 16:
        order = WORD_ORDER_16
    elif n == 10:
        order = WORD_ORDER_10
    else:
        return None
    photos = {order[i]: paths[i] for i in range(min(8, len(order)))}
    radios = paths[-2:] if n >= 10 else []
    autres = paths[8:-2] if n >= 16 else []
    return photos, radios, autres


def extract_word_media(docx_path, out_dir):
    """Extrait les images de contenu dans out_dir (ordre du document). Renvoie [chemins]."""
    import zipfile
    os.makedirs(out_dir, exist_ok=True)
    paths = []
    with zipfile.ZipFile(docx_path) as z:
        for _idx, name, sz in _media_entries(docx_path):
            if sz < WORD_LOGO_MIN_SIZE:
                continue
            ext = os.path.splitext(name)[1].lower() or ".png"
            dst = os.path.join(out_dir, "w%02d%s" % (len(paths), ext))
            with open(dst, "wb") as fh:
                fh.write(z.read(name))
            paths.append(dst)
    return paths


def plan_summary(plan):
    """Petit résumé chiffré pour le récapitulatif."""
    n_photos = sum(len(t["photos"]) for t in plan["timepoints"])
    n_renders = sum(1 for rk, _ in plan["renders"] if rk)
    word_imgs = count_word_media(plan["word_path"]) if plan["mode"] == "word_only" and plan["word_path"] else 0
    return {"temps": len(plan["timepoints"]), "photos": n_photos, "radios": len(plan["radios"]),
            "stl": len(plan["stl"]), "renders": n_renders, "renders_extra": len(plan["renders"]) - n_renders,
            "word": bool(plan["word_path"]), "mode": plan["mode"], "word_imgs": word_imgs}
