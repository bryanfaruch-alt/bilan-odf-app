# -*- coding: utf-8 -*-
"""
Bilan ODF — application locale de gestion des bilans orthodontiques.
Login + bibliothèque + génération + suivi dans le temps (réévaluations). 100% local.
Lancer :  python3 app.py   puis http://127.0.0.1:5000
"""
import os, json, re, datetime, secrets, unicodedata, traceback, shutil, subprocess, sys
from flask import (Flask, request, redirect, url_for, session, send_file,
                   render_template_string, flash, abort, jsonify)
from werkzeug.security import generate_password_hash, check_password_hash
import pipeline as P
import word_import as WI
import backup as BK
import mass_import as MI

# Répertoire des ressources : dossier du script, ou dossier temporaire du bundle
# PyInstaller (_MEIPASS) quand l'app est « gelée » en exécutable autonome.
if getattr(sys, "frozen", False):
    HERE = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(sys.executable)))
else:
    HERE = os.path.dirname(os.path.abspath(__file__))
APP_VERSION = "1.2"          # version de l'app (source unique : lue aussi par run_native pour la MAJ auto)
DATA = os.environ.get("BILANODF_DATA") or os.path.expanduser("~/BilanODF_Data")
PATIENTS = os.path.join(DATA, "patients")
CONFIG = os.path.join(DATA, "config.json")
BACKUP_DIR = os.environ.get("BILANODF_BACKUPS") or os.path.expanduser("~/BilanODF_Sauvegardes")
TEMPLATE_DOCX = os.path.join(HERE, "modele", "Bilan_MODELE.docx")
os.makedirs(PATIENTS, exist_ok=True)

def load_config():
    if os.path.exists(CONFIG): return json.load(open(CONFIG))
    return {"secret": secrets.token_hex(16), "users": {}}
def save_config(c): json.dump(c, open(CONFIG, "w"), ensure_ascii=False, indent=1)
cfg = load_config(); save_config(cfg)
def get_settings(): return cfg.get("settings") or {}
def save_settings(s): cfg["settings"] = s; save_config(cfg)

app = Flask(__name__)
app.secret_key = cfg["secret"]

@app.errorhandler(Exception)
def _show_error(e):
    """Affiche l'erreur réelle (appli locale) au lieu d'un 500 générique, pour faciliter le diagnostic."""
    from werkzeug.exceptions import HTTPException
    if isinstance(e, HTTPException): return e
    tb = traceback.format_exc()
    return ("<div style='font:14px -apple-system,sans-serif;padding:24px;max-width:900px;margin:auto'>"
            "<h2 style='color:#C0392B'>Une erreur est survenue</h2>"
            "<p>Envoie cette capture à ton assistant pour correction rapide.</p>"
            "<pre style='white-space:pre-wrap;background:#f6f8fb;border:1px solid #dde;border-radius:8px;padding:12px;font-size:12px'>%s</pre>"
            "<p><a href='/'>← Retour à la bibliothèque</a></p></div>") % tb, 500

def slugify(s):
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_") or "patient"
def logged(): return session.get("user")

# ---- Steiner : (nom interne, libellé court, moyenne, écart-type) ----
STEINER_FIELDS = [
    ("SNA (°)", "SNA", 80.22, 2.8), ("SNB (°)", "SNB", 76.18, 2.5), ("ANB (°)", "ANB", 4.08, 1.6),
    ("U1-NA (°)", "U1-NA (deg)", 22, 5.0), ("U1-NA (mm)", "U1-NA (mm)", 4, 3.0),
    ("L1-NB (°)", "L1-NB (deg)", 25, 5.0), ("L1-NB (mm)", "L1-NB (mm)", 4, 2.0),
    ("Wits (mm)", "Wits", -0.33, 2.7), ("Pog-NB (mm)", "Pog-NB", 2.25, 1.3),
    ("Plan mandibulaire (Go-Gn/SN) (°)", "Plan mandibulaire", 32, 4.0),
    ("Plan occlusal/SN (°)", "Plan occlusal", 14, 4.0),
    ("Angle interincisif (°)", "Angle inter-incisif", 128, 5.3),
]
PHOTO_FIELDS = [
    ("exo_face_repos", "Visage repos"), ("exo_face_sourire", "Visage sourire"), ("exo_profil", "Profil"),
    ("endo_occlusion_frontale", "Occlusion frontale"), ("endo_laterale_droite", "Latérale droite"),
    ("endo_laterale_gauche", "Latérale gauche"), ("endo_occlusal_maxillaire", "Occlusal maxillaire"),
    ("endo_occlusal_mandibulaire", "Occlusal mandibulaire"),
]
RENDER_FIELDS = [("upper_occlusal", "Occlusal sup."), ("lower_occlusal", "Occlusal inf."),
                 ("front_occ", "Face"), ("right_occ", "Droite"), ("left_occ", "Gauche"),
                 ("spee_lower", "Spee droite"), ("spee_left", "Spee gauche")]

def steiner_signif(nom, v, mean, sd):
    z = (v - mean) / sd if sd else 0
    sev = "**" if abs(z) > 2 else ("*" if abs(z) > 1 else "")
    hi = v > mean
    M = {"SNA (°)": ("Maxillaire antéposé", "Maxillaire rétroposé"),
         "SNB (°)": ("Mandibule antéposée (prognathe)", "Mandibule rétroposée"),
         "ANB (°)": ("Tendance Classe II squelettique", "Tendance Classe III squelettique"),
         "Wits (mm)": ("Tendance Classe II", "Tendance Classe III"),
         "Plan mandibulaire (Go-Gn/SN) (°)": ("Hyperdivergent", "Hypodivergent"),
         "Plan occlusal/SN (°)": ("Plan occlusal augmenté", "Plan occlusal diminué"),
         "U1-NA (°)": ("Incisive sup. vestibuloversée", "Incisive sup. rétroversée"),
         "U1-NA (mm)": ("Incisive sup. antéposée", "Incisive sup. rétroposée"),
         "L1-NB (°)": ("Incisive inf. vestibuloversée", "Incisive inf. rétroversée"),
         "L1-NB (mm)": ("Incisive inf. antéposée", "Incisive inf. rétroposée"),
         "Angle interincisif (°)": ("Angle inter-incisif fermé", "Angle inter-incisif ouvert"),
         "Pog-NB (mm)": ("Menton proéminent", "Menton effacé")}
    if not sev: return "Dans la norme", ""
    pair = M.get(nom)
    return (pair[0] if hi else pair[1]) if pair else "Hors norme", sev

COURRIER_TPL = """<!doctype html><html lang=fr><head><meta charset=utf-8>
<meta name=viewport content="width=device-width,initial-scale=1"><title>Courrier &mdash; __NOM__</title>
<style>
*{box-sizing:border-box}
body{margin:0;background:#5c6b7a;font-family:Georgia,'Times New Roman',serif;color:#1a2430}
#bar{position:sticky;top:0;z-index:5;background:#2F5CA8;color:#fff;display:flex;gap:10px;align-items:center;padding:10px 16px;font-family:-apple-system,sans-serif}
#bar a,#bar button{background:rgba(255,255,255,.15);color:#fff;border:1px solid rgba(255,255,255,.3);border-radius:8px;padding:7px 14px;font-size:13px;cursor:pointer;text-decoration:none}
#bar button.p{background:#fff;color:#2F5CA8;border-color:#fff;font-weight:600}
#bar .sp{flex:1}
#bar .hint{font-size:12px;opacity:.9}
.sheet{background:#fff;width:210mm;min-height:297mm;margin:22px auto;padding:26mm 24mm;box-shadow:0 6px 30px rgba(0,0,0,.3);line-height:1.6;font-size:15px}
.sheet:focus{outline:none}
.top{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:26px;gap:20px}
.exp{font-size:13px;color:#33475b;line-height:1.5}
.exp b{font-size:15px;color:#1a2430}
.date{font-size:14px;text-align:right;white-space:nowrap}
.sheet p{margin:0 0 12px}
.sheet h2{font-size:15px;margin:0 0 4px}
[contenteditable]{outline:none}
[contenteditable]:hover{background:#fbfcfe}
@media print{
  body{background:#fff}
  #bar{display:none}
  .sheet{box-shadow:none;margin:0;width:auto;min-height:auto;padding:18mm 20mm}
  @page{size:A4;margin:0}
}
</style></head><body>
<div id=bar>
  <a href="__BACK__">&larr; Retour</a>
  <button class=p onclick="window.print()">&#128424; Imprimer / PDF</button>
  <span class=sp></span>
  <span class=hint>Clique dans le texte pour l&rsquo;ajuster avant d&rsquo;imprimer</span>
</div>
<div class=sheet contenteditable=true spellcheck=false>
  <div class=top>
    <div class=exp contenteditable=true><b>Dr [Votre nom]</b><br>Orthodontie / Orthop&eacute;die dento-faciale<br>[Adresse du cabinet]<br>[T&eacute;l&eacute;phone &middot; e-mail]</div>
    <div class=date contenteditable=true>Le __DATE__</div>
  </div>
  __LETTER__
</div>
</body></html>"""


def _diag_summary(r):
    syn = r.get("synthese", {}) or {}
    parts = []
    for cl, cc in DIAG_COLS:
        vals = []
        for _rl, rc in DIAG_ROWS:
            v = (syn.get(cc + "_" + rc) or "").strip()
            if v: vals.append(v)
        if vals:
            parts.append("%s : %s" % (cl.lower(), ", ".join(vals)))
    return " ; ".join(parts)


def _angle_summary(r):
    ang = r.get("angle", {}) or {}
    seg = []
    md, mg = ang.get("molaire_d"), ang.get("molaire_g")
    cd, cg = ang.get("canine_d"), ang.get("canine_g")
    if md or mg: seg.append("rapports molaires %s à droite et %s à gauche" % (md or "non précisé", mg or "non précisé"))
    if cd or cg: seg.append("rapports canins %s à droite et %s à gauche" % (cd or "non précisé", cg or "non précisé"))
    return ", ".join(seg)


def _steiner_line(r):
    st = r.get("steiner", {}) or {}
    lbl = [("SNA (°)", "SNA"), ("SNB (°)", "SNB"), ("ANB (°)", "ANB"),
           ("Plan mandibulaire (Go-Gn/SN) (°)", "Go-Gn/SN"), ("Angle interincisif (°)", "angle inter-incisif")]
    bits = []
    for k, sh in lbl:
        v = st.get(k)
        if v not in (None, ""):
            try: bits.append("%s = %g°" % (sh, float(v)))
            except Exception: pass
    return ", ".join(bits)


@app.route("/record/<slug>/<rid>/courrier")
def courrier(slug, rid):
    if not logged(): return redirect(url_for("login"))
    import datetime as _dt
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r: abort(404)
    nom = pt.get("nom", slug); sexe = (pt.get("sexe") or "").lower()
    qual = "votre patiente" if "f" in sexe and "précisé" not in sexe else ("votre patient" if "m" in sexe and "féminin" not in sexe else "votre patient(e)")
    age = pt.get("age", "")
    age_bit = (", âgé%s de %s" % ("e" if "f" in sexe and "précisé" not in sexe else "", age)) if age else ""
    motif = (r.get("motif") or "").strip()
    diag = _diag_summary(r); angle = _angle_summary(r); stein = _steiner_line(r)
    P = []
    P.append("<p>Cher Confrère, Chère Consœur,</p>")
    P.append("<p>Je vous remercie de m'avoir adressé %s <b>%s</b>%s, et vous fais part des éléments de mon examen orthodontique.</p>" % (qual, nom, age_bit))
    P.append("<p><b>Motif de consultation :</b> %s</p>" % (motif if motif else "…"))
    if diag: P.append("<p><b>Synthèse diagnostique :</b> %s.</p>" % diag)
    else: P.append("<p><b>Synthèse diagnostique :</b> …</p>")
    if angle: P.append("<p><b>Rapports occlusaux :</b> %s.</p>" % angle)
    if stein: P.append("<p><b>Analyse céphalométrique (Steiner) :</b> %s.</p>" % stein)
    P.append("<p><b>Plan de traitement envisagé :</b> [à compléter].</p>")
    P.append("<p>Je reste à votre disposition pour tout renseignement complémentaire et vous prie d'agréer, Cher Confrère, Chère Consœur, l'expression de mes salutations confraternelles.</p>")
    P.append("<p style=\"text-align:right;margin-top:34px\">Dr [Votre nom]</p>")
    letter = "".join(P)
    mois = ["janvier","février","mars","avril","mai","juin","juillet","août","septembre","octobre","novembre","décembre"]
    d = _dt.date.today(); datestr = "%d %s %d" % (d.day, mois[d.month-1], d.year)
    html = (COURRIER_TPL.replace("__NOM__", nom).replace("__DATE__", datestr)
            .replace("__LETTER__", letter).replace("__BACK__", url_for("record", slug=slug, rid=rid)))
    return html


def steiner_table_html(r):
    """Tableau Steiner : Mesure | Moyenne | Écart-type | Résultat | Interprétation (sans sévérité)."""
    st = r.get("steiner", {}) or {}
    rows = ""
    for nom, short, mean, sd in STEINER_FIELDS:
        v = st.get(nom)
        if v in (None, ""):
            res, interp, col = "–", "", "#8a93a3"
        else:
            try:
                fv = float(v); txt, sev = steiner_signif(nom, fv, mean, sd)
                res, interp = ("%g" % fv), txt; col = "#C0392B" if sev else "#1E7D32"
            except Exception:
                res, interp, col = str(v), "", "#333"
        rows += ('<tr><td>%s</td><td style="text-align:center">%g</td><td style="text-align:center">%g</td>'
                 '<td style="text-align:center;font-weight:600;color:%s">%s</td>'
                 '<td style="color:%s">%s</td></tr>') % (short, mean, sd, col, res, col, interp)
    return ('<table class=st style="width:100%%"><tr><th style="text-align:left">Mesure</th><th>Moyenne</th>'
            '<th>Écart-type</th><th>Résultat</th><th style="text-align:left">Interprétation</th></tr>%s</table>') % rows

# ---- data helpers ----
def pdir(slug): return os.path.join(PATIENTS, slug)
def rdir(slug, rid): return os.path.join(PATIENTS, slug, "records", rid)
def load_patient(slug):
    p = os.path.join(pdir(slug), "patient.json")
    return json.load(open(p)) if os.path.exists(p) else None
def save_patient(slug, d): json.dump(d, open(os.path.join(pdir(slug), "patient.json"), "w"), ensure_ascii=False, indent=1)
def load_rec(slug, rid):
    p = os.path.join(rdir(slug, rid), "meta.json")
    return json.load(open(p)) if os.path.exists(p) else None
def save_rec(slug, rid, d): json.dump(d, open(os.path.join(rdir(slug, rid), "meta.json"), "w"), ensure_ascii=False, indent=1)

def list_patients():
    out = []
    if not os.path.isdir(PATIENTS): return out
    for slug in os.listdir(PATIENTS):
        pt = load_patient(slug)
        if pt: out.append(pt)
    out.sort(key=lambda m: m.get("date_creation", ""), reverse=True)
    return out

def age_from_dob(dob):
    try:
        dt = datetime.date.fromisoformat(dob); t = datetime.date.today()
        y = t.year - dt.year - ((t.month, t.day) < (dt.month, dt.day))
        return "%02d/%d" % (dt.month, dt.year), "%d ans" % y
    except Exception:
        return "—", "—"

# ---- génération d'un enregistrement (bilan d'un timepoint) ----
def run_generation(base, info, steiner_res, rotations, crops=None, skip_stl=False):
    log, err = [], None
    docx = pdf = None
    try:
        if any(os.path.exists(os.path.join(base, "01_photos_brutes", k + ".jpg")) for k, _ in PHOTO_FIELDS):
            P.process_photos(base, rotations, crops); log.append("photos")
        up = os.path.join(base, "04_stl_bruts", "UpperJawScan.stl")
        lo = os.path.join(base, "04_stl_bruts", "LowerJawScan.stl")
        _rd = os.path.join(base, "05_stl_rendus")
        renders_ok = os.path.exists(os.path.join(_rd, "upper_occlusal.png")) or os.path.exists(os.path.join(_rd, "lower_occlusal.png"))
        # Empreintes ajoutées / remplacées APRÈS le dernier rendu -> forcer un nouveau rendu
        _stl_newer = False
        try:
            _rt = max([os.path.getmtime(os.path.join(_rd, f)) for f in os.listdir(_rd)] or [0]) if os.path.isdir(_rd) else 0
            _st = max([os.path.getmtime(p) for p in (up, lo) if os.path.exists(p)] or [0])
            _stl_newer = _st > _rt + 1
        except Exception:
            pass
        if (os.path.exists(up) or os.path.exists(lo)) and not (skip_stl and renders_ok and not _stl_newer):
            # rendu 3D dans un SOUS-PROCESSUS isolé : un crash OpenGL/VTK ne tue pas l'app.
            # Réessais : le contexte OpenGL off-screen échoue parfois de façon transitoire.
            if getattr(sys, "frozen", False):
                cmd = [sys.executable, "render_stl", base]
            else:
                cmd = [sys.executable, os.path.join(HERE, "pipeline.py"), "render_stl", base]
            _want = "upper_occlusal.png" if os.path.exists(up) else "lower_occlusal.png"
            _ok = False
            for _attempt in range(3):
                try:
                    res = subprocess.run(cmd, timeout=420, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
                    if res.returncode == 0 and os.path.exists(os.path.join(_rd, _want)):
                        _ok = True; break
                except Exception:
                    pass
                import time as _time; _time.sleep(1.5)
            log.append("rendus 3D" if _ok else "rendus 3D (échec)")
        if steiner_res:
            sig, sev = {}, {}
            for nom, v in steiner_res.items():
                mean, sd = next((m, s) for (n, _l, m, s) in STEINER_FIELDS if n == nom)
                t, x = steiner_signif(nom, v, mean, sd); sig[nom] = t; sev[nom] = x
            P.make_steiner_json(base, steiner_res, sig, sev); P.build_steiner_figure(base); log.append("Steiner")
        docx, pdf = P.build_bilan(base, TEMPLATE_DOCX, info)
    except Exception as e:
        err = str(e); traceback.print_exc()
    return docx, pdf, log, err

def build_info(pt, rec):
    return {"nom": pt["nom"], "dob_court": pt["dob_court"], "age": pt["age"], "sexe": pt["sexe"],
            "synthese": rec.get("synthese", {}), "motif": rec.get("motif", ""),
            "plan": rec.get("plan", {}), "espace": espace_effective(rec.get("plan", {}), rec.get("steiner", {}))}

def parse_steiner_form(form):
    res = {}
    for nom, short, mean, sd in STEINER_FIELDS:
        v = form.get("st_" + nom, "").replace(",", ".").strip()
        if v:
            try: res[nom] = float(v)
            except ValueError: pass
    return res

# --- Synthèse diagnostique : grille (colonnes × sens) ---
DIAG_COLS = [("Squelettique", "sq"), ("Alvéolaire", "al"), ("Occlusal", "oc"), ("Cutané", "cu")]
DIAG_ROWS = [("Sagittal", "ap"), ("Transversal", "tr"), ("Vertical", "ve")]
DIAG_OPTIONS = {
    "sq_ap": ["Classe I squelettique", "Classe II squelettique", "Classe III squelettique", "Classe III (tendance)"],
    "sq_tr": ["Symétrique", "Endognathie maxillaire", "Asymétrie mandibulaire"],
    "sq_ve": ["Normodivergent", "Hyperdivergent", "Hypodivergent"],
    "al_ap": ["Incisives normo-inclinées", "Proalvéolie supérieure", "Rétroalvéolie supérieure", "Proalvéolie inférieure", "Rétroalvéolie inférieure"],
    "al_tr": ["Normo-alvéolie", "Endo-alvéolie", "Exo-alvéolie"],
    "al_ve": ["Normo-alvéolie", "Supra-alvéolie", "Infra-alvéolie"],
    "oc_ap": ["Classe I molaire et canine", "Classe II molaire", "Classe III molaire", "Surplomb normal", "Surplomb augmenté", "Surplomb inversé"],
    "oc_tr": ["Milieux inter-incisifs coïncidents", "Déviation du milieu supérieur", "Déviation du milieu inférieur", "Articulé inversé"],
    "oc_ve": ["Recouvrement normal", "Supraclusion", "Infraclusion", "Béance antérieure"],
    "cu_ap": ["Profil équilibré", "Profil convexe", "Profil concave"],
    "cu_tr": ["Symétrie faciale", "Asymétrie faciale"],
    "cu_ve": ["Proportions verticales équilibrées", "Étage inférieur augmenté", "Étage inférieur diminué"],
}

# --- Examen clinique du patient ---
EXAM_ENV = [
    ("langue_volume", "Langue — volume", ["augmenté", "normal", "diminué"]),
    ("langue_posture", "Posture de la langue", ["basse", "haute", "interposée", "normale"]),
    ("amygdales", "Amygdales", ["normales", "hypertrophiées", "absentes"]),
    ("tonicite", "Tonicité musculaire", ["normale", "hypotonique", "hypertonique"]),
    ("deglutition", "Déglutition", ["mature (typique)", "atypique primaire", "atypique avec interposition linguale"]),
    ("ventilation", "Ventilation", ["nasale", "buccale", "mixte"]),
    ("hygiene", "Hygiène bucco-dentaire", ["bonne", "moyenne", "insuffisante"]),
    ("gencive", "Gencive", ["saine", "inflammatoire", "récession(s)"]),
    ("freins", "Freins", ["normaux", "frein lingual court", "frein labial hypertrophié"]),
    ("habitudes", "Habitudes", ["aucune", "succion du pouce", "tétine", "onychophagie", "succion labiale"]),
    ("bruxisme", "Bruxisme", ["absent", "présent (diurne)", "présent (nocturne)"]),
    ("ronflement", "Ronflement / SAOS", ["absent", "ronflement", "suspicion de SAOS"]),
]
EXAM_FACE = [
    ("symetrie", "Symétrie", ["symétrique", "asymétrie droite", "asymétrie gauche"]),
    ("typologie", "Typologie faciale", ["mésofaciale", "dolichofaciale", "brachyfaciale"]),
    ("profil", "Profil", ["droit", "convexe", "concave"]),
    ("etage_inf", "Étage sous-nasal / inférieur", ["équilibré", "augmenté", "diminué"]),
    ("transversal", "Proportions transversales", ["équilibrées", "étroit", "large"]),
    ("sillon_lm", "Sillon labio-mentonnier", ["normal", "marqué", "effacé"]),
    ("cervico_ment", "Distance cervico-mentonnière", ["normale", "courte", "longue"]),
    ("angle_nl", "Angle naso-labial", ["normal", "ouvert", "fermé"]),
    ("contact_bl", "Contact bilabial", ["présent", "absent (inocclusion labiale)"]),
]
# --- Suivi de traitement (par rendez-vous / réévaluation) ---
# (clé, libellé, indice de saisie, multiligne)
SUIVI_FIELDS = [
    ("appareil", "Type d'appareil", "ex. Multibague MBT .022 · Damon · Invisalign · Quad Helix · FR…", False),
    ("arc_haut", "Arc supérieur", "ex. NiTi .014 · Acier .019×.025", False),
    ("arc_bas", "Arc inférieur", "ex. NiTi .016 · TMA .017×.025", False),
    ("attaches", "Chaînettes / ligatures / ressorts", "ex. chaînette 13→23 · ressort ouvert 15-16", False),
    ("elastiques", "Élastiques", "ex. Classe II 3/16 6oz D+G la nuit", False),
    ("actes", "Gestes réalisés ce rendez-vous", "collage, changement d'arc, activation…", True),
    ("commentaire", "Commentaire / prochaine étape", "", True),
]

# --- Journal de séances (une entrée par rendez-vous) : menus pré-remplis + texte libre ---
SG_ARC = [".012 NiTi", ".014 NiTi", ".016 NiTi", ".016×.022 NiTi", ".017×.025 NiTi",
          ".018 Acier", ".016×.022 Acier", ".019×.025 Acier", ".017×.025 TMA", ".019×.025 TMA", "—"]
SG_APP = ["Multibague MBT .022", "Multibague Roth .018", "Damon", "Aligneurs (Invisalign)",
          "Aligneurs (Spark)", "Quad Helix", "Disjoncteur (Hyrax)", "Plaque amovible",
          "Arc de traction extra-oral", "Contention collée", "Gouttière de contention"]
SG_ATT = ["Chaînette élastomérique", "Ressort ouvert (NiTi)", "Ressort fermé (NiTi)",
          "Ligatures métalliques", "Ligature en 8", "Lace-back", "Butée"]
SG_ELA = ["Aucun", "Classe II 3/16 6oz", "Classe III 3/16 6oz", "Boîte antérieure",
          "Triangle", "Cross-elastic", "Vertical antérieur"]
SG_PHASE = ["Nivellement / alignement", "Fermeture d'espaces", "Correction de classe",
            "Coordination des arcs", "Finitions", "Contention"]
SG_ACTE = ["Collage complet", "Changement d'arc", "Activation", "Repositionnement de bracket",
           "Stripping", "Pose de mini-vis", "Dépose de l'appareil", "Empreinte / scan", "Contrôle simple"]
SG_HYG = ["Bonne", "Moyenne", "Insuffisante"]
SG_NEXT = ["4 semaines", "5 semaines", "6 semaines", "8 semaines", "3 mois", "6 mois"]

# (clé, libellé, options|None, multiligne, afficher dans « état actuel »)
JOURNAL_FIELDS = [
    ("phase", "Phase de traitement", SG_PHASE, False, True),
    ("appareil", "Type d'appareil", SG_APP, False, True),
    ("arc_haut", "Arc supérieur", SG_ARC, False, True),
    ("arc_bas", "Arc inférieur", SG_ARC, False, True),
    ("attaches", "Chaînettes / ressorts / ligatures", SG_ATT, False, True),
    ("elastiques", "Élastiques", SG_ELA, False, True),
    ("actes", "Gestes réalisés ce RDV", SG_ACTE, True, False),
    ("hygiene", "Hygiène / observance", SG_HYG, False, False),
    ("prochain", "Prochain RDV (délai)", SG_NEXT, False, False),
    ("commentaire", "Commentaire", None, True, False),
]
# champs mis en avant dans l'encart « état actuel »
CURRENT_FIELDS = [("appareil", "Appareil"), ("arc_haut", "Arc supérieur"), ("arc_bas", "Arc inférieur"),
                  ("elastiques", "Élastiques"), ("attaches", "Attaches"), ("phase", "Phase")]

def parse_journal(form, prefix="jv_"):
    """Garde toute valeur non vide telle quelle (y compris les choix explicites
    « Aucun » / « — », qui sont des informations utiles : pas d'élastique, arc passif…)."""
    e = {}
    for k, _l, _o, _ml, _c in JOURNAL_FIELDS:
        v = form.get(prefix + k, "").strip()
        if v:
            e[k] = v
    return e

def journal_current_state(pt):
    """Remonte les séances (de la plus récente à la plus ancienne) et prend, pour
    chaque champ, la 1re valeur renseignée : donne l'état ACTUEL du traitement."""
    entries = sorted(pt.get("journal", []) or [], key=lambda x: x.get("date", ""), reverse=True)
    state = {}
    for e in entries:
        for k, _l, _o, _ml, cur in JOURNAL_FIELDS:
            if cur and k not in state and e.get(k):
                state[k] = {"val": e[k], "date": e.get("date", "")}
    return state, entries

ANGLE_FIELDS = [
    ("molaire_d", "Molaire droite", ["Classe I", "Classe II div.1", "Classe II div.2", "Classe III"]),
    ("molaire_g", "Molaire gauche", ["Classe I", "Classe II div.1", "Classe II div.2", "Classe III"]),
    ("canine_d", "Canine droite", ["Classe I", "Classe II", "Classe III"]),
    ("canine_g", "Canine gauche", ["Classe I", "Classe II", "Classe III"]),
]

# ======================================================================
#  PLAN DE TRAITEMENT — analyse d'espace de Steiner + chevrons + stratégies
# ======================================================================
# Lignes du tableau d'analyse d'espace. inp=True -> saisie (mm signés :
# + = espace gagné, − = espace perdu). inp=False -> ligne calculée.
ESPACE_ROWS = [
    ("encombrement", "Encombrement", True),
    ("repositionnement_incisive", "Repositionnement incisive", True),
    ("courbe_spee", "Courbe de Spee", True),
    ("_ddm", "DDM", False),
    ("repositionnement_6", "Repositionnement 6", True),
    ("espace_derive_mesiale", "Espace de dérive mésiale", True),
    ("_total", "TOTAL", False),
    ("extraction_stripping", "Extraction / stripping", True),
    ("individualisation", "Individualisation", True),
    ("_net", "NET", False),
]
# 4 valeurs par chevron : ANB (haut-ext), I-NA mm (bras haut), I-NB mm (bras bas), Pog-NB (bas-ext)
CHEVRON_POS = ["anb", "ina", "inb", "pog"]
CHEVRON_POS_META = [("anb", "ANB"), ("ina", "I-NA mm"), ("inb", "I-NB mm"), ("pog", "Pog-NB")]
_CHV_STYLE = {"anb": "top:1%;left:0", "ina": "top:23%;left:36%", "inb": "top:57%;left:36%", "pog": "top:80%;left:0"}
# enchaînement des chevrons du plan de traitement (avec les positions visibles par chevron)
CHEVRON_FLOW = [
    ("probleme", "Problème", ("anb", "ina", "inb", "pog")),
    ("anb_solution", "ANB solution", ("anb", "ina", "inb")),        # pas de Pog
    ("pog_solution", "Pog-NB solution", ("ina", "inb", "pog")),     # pas d'ANB
    ("solution", "Solution", ("anb", "ina", "inb", "pog")),
    ("individualisation", "Individualisation", ("anb", "ina", "inb", "pog")),
]
_CHV_SVG = ('<svg viewBox="0 0 120 120" style="position:absolute;inset:0;width:100%;height:100%;pointer-events:none">'
            '<polyline points="26,14 98,60 26,106" fill="none" stroke="#2F5CA8" stroke-width="4" '
            'stroke-linecap="round" stroke-linejoin="round"/></svg>')

def compromis_targets(anb):
    """Positions incisives cibles EN MM — compromis acceptable de Steiner (FMA 28°, Wits −2 mm)."""
    if anb in (None, ""): return None
    try: a = float(anb)
    except Exception: return None
    return {"ul": round(6 - a, 2), "ll": round(3.5 + 0.25 * a, 2)}

def repos_incisive_suggest(steiner):
    """Espace (mm) lié au repositionnement des incisives inférieures, déduit du Steiner :
    ~2 mm d'espace par mm d'écart de i-NB (mm) mesuré vs la cible du compromis (3,5 + 0,25·ANB ;
    4 mm si ANB inconnu). + = incisives à avancer (gain) ; − = à reculer (perte)."""
    inb = steiner.get("L1-NB (mm)")
    if inb in (None, ""): return None
    try:
        anb = steiner.get("ANB (°)")
        target = (3.5 + 0.25 * float(anb)) if anb not in (None, "") else 4.0
        return round(2.0 * (target - float(inb)), 1)
    except Exception: return None

ESPACE_INPUTS = [k for k, _l, inp in ESPACE_ROWS if inp]
ESPACE_DDM = ["encombrement", "repositionnement_incisive", "courbe_spee"]
ESPACE_TOTAL = ESPACE_DDM + ["repositionnement_6", "espace_derive_mesiale"]
ESPACE_NET = ESPACE_TOTAL + ["extraction_stripping", "individualisation"]

def esp_cell(d):
    """Normalise une case du tableau d'espace. Accepte le nouveau format {p,m}
    ET l'ancien (nombre signé : + = colonne +, − = colonne −)."""
    if isinstance(d, dict):
        def f(x):
            try: return float(x or 0)
            except Exception: return 0.0
        return {"p": f(d.get("p")), "m": f(d.get("m"))}
    try:
        v = float(d); return {"p": v if v > 0 else 0.0, "m": -v if v < 0 else 0.0}
    except Exception:
        return {"p": 0.0, "m": 0.0}

def espace_effective(plan, steiner):
    """Colonnes + / − par ligne (repositionnement incisif déduit du Steiner si vide) + sommes DDM/TOTAL/NET."""
    src = (plan or {}).get("espace", {})
    if not isinstance(src, dict): src = {}
    e = {k: esp_cell(src.get(k)) for k in ESPACE_INPUTS}
    ri = e["repositionnement_incisive"]
    if not (ri["p"] or ri["m"]):
        s = repos_incisive_suggest(steiner or {})
        if s is not None:
            if s >= 0: ri["p"] = round(s, 1)
            else: ri["m"] = round(-s, 1)
    def grp(keys):
        return {"p": round(sum(e[k]["p"] for k in keys), 1), "m": round(sum(e[k]["m"] for k in keys), 1)}
    e["_ddm"], e["_total"], e["_net"] = grp(ESPACE_DDM), grp(ESPACE_TOTAL), grp(ESPACE_NET)
    return e

def parse_plan(form):
    plan = {"espace": {}, "chevrons": {}, "objectifs": "", "moyens": ""}
    for k in ESPACE_INPUTS:
        d = {}
        for side in ("p", "m"):
            v = form.get("esp_%s_%s" % (k, side), "").replace(",", ".").strip()
            if v:
                try: d[side] = round(float(v), 2)
                except Exception: pass
        if d: plan["espace"][k] = d
    for pk, _pl, pos in CHEVRON_FLOW:   # tous saisis à la main ; seules les positions visibles
        d = {}
        for ck in pos:
            v = form.get("chv_%s_%s" % (pk, ck), "").strip()
            if v: d[ck] = v
        if d: plan["chevrons"][pk] = d
    plan["objectifs"] = form.get("pt_objectifs", "").strip()
    plan["moyens"] = form.get("pt_moyens", "").strip()
    return plan

def _fmt(v):
    if v in (None, ""): return ""
    try:
        f = float(v); return ("%g" % f)
    except Exception: return str(v)

def _num(x):
    """3.25 -> '3,25' ; 4.0 -> '4'."""
    return ("%g" % x).replace(".", ",")

def compromis_chevron_strip(anb):
    """Planche de référence (valeurs EN MM) : une bande de chevrons pour ANB de −1 à 8,
    chacun portant U1-NA (mm) sur le bras haut et i-NB (mm) sur le bras bas.
    La cellule de l'ANB du patient est surlignée."""
    try: a = int(round(float(anb))) if anb not in (None, "") else None
    except Exception: a = None
    # chevron décalé à droite : la colonne de gauche reste libre pour les deux valeurs (lisibles)
    svg = ('<svg viewBox="0 0 92 150" style="position:absolute;inset:0;width:100%;height:100%;pointer-events:none">'
           '<polyline points="48,34 82,75 48,116" fill="none" stroke="#2F5CA8" stroke-width="3" '
           'stroke-linecap="round" stroke-linejoin="round"/></svg>')
    cells = ""
    for A in range(-1, 9):
        bg = "background:#dbe9ff;" if (a is not None and A == a) else "background:#fff;"
        cells += ('<div style="%sflex:0 0 92px;border:1px solid #9fb2d4;border-radius:3px;position:relative;height:150px">'
                  '<div style="position:absolute;top:5px;left:8px;font-weight:700;font-size:13px">%d&deg;</div>'
                  '%s'
                  '<div style="position:absolute;top:52px;left:8px;font-size:12px;font-weight:600">%s</div>'
                  '<div style="position:absolute;top:104px;left:8px;font-size:12px;font-weight:600">%s</div>'
                  '</div>') % (bg, A, svg, _num(6 - A), _num(3.5 + 0.25 * A))
    return ('<h3 style="margin:16px 0 6px;color:var(--acc)">Valeurs de référence — compromis de Steiner (mm)</h3>'
            '<div class=muted style="font-size:12px;margin:0 0 6px">FMA = 28°, Wits = −2 mm — position cible des incisives '
            'selon l\'ANB, en mm. Bras haut : U1-NA mm ; bras bas : i-NB mm. La colonne surlignée = ANB du patient.</div>'
            '<div style="display:flex;gap:5px;overflow-x:auto;padding-bottom:6px">%s</div>') % cells

def _chevron_box(inner, style):
    return ('<div style="position:absolute;%s;width:54px;text-align:center">%s</div>') % (style, inner)

def _chevron_frame(title, boxes, accent=False):
    hd = 'color:var(--acc)' if accent else ''
    return ('<div style="text-align:center"><div class=muted style="margin-bottom:4px;font-weight:600;%s">%s</div>'
            '<div style="position:relative;width:168px;height:132px;margin:0 auto">%s%s</div></div>') % (hd, title, _CHV_SVG, boxes)

def _chevron_display(title, idpfx, data):
    """Chevron en lecture seule (valeurs calculées), avec ids pour mise à jour live JS."""
    boxes = ""
    for ck, lab in CHEVRON_POS_META:
        v = data.get(ck) if data else None
        txt = _fmt(v) if v not in (None, "") else "–"
        inner = ('<div id="%s_%s" style="border:1px solid #b9c6de;border-radius:5px;padding:2px;background:#f3f8ff;font-size:13px">%s</div>'
                 '<div class=muted style="font-size:9px">%s</div>') % (idpfx, ck, txt, lab)
        boxes += _chevron_box(inner, _CHV_STYLE[ck])
    return _chevron_frame(title, boxes, accent=True)

def _chevron_widget(pk, pl, data, positions=None):
    """Chevron éditable (rempli à la main). positions = sous-ensemble de clés à afficher."""
    if not isinstance(data, dict): data = {}
    if positions is None: positions = CHEVRON_POS
    boxes = ""
    for ck, lab in CHEVRON_POS_META:
        if ck not in positions: continue
        inner = ('<input name="chv_%s_%s" value="%s" autocomplete=off placeholder="mm" '
                 'style="width:54px;text-align:center;font-size:13px;border:1px solid #ccd;border-radius:5px;padding:2px">'
                 '<div class=muted style="font-size:9px">%s</div>') % (pk, ck, str(data.get(ck, "") or "").replace('"', "&quot;"), lab)
        boxes += _chevron_box(inner, _CHV_STYLE[ck])
    return _chevron_frame(pl, boxes)

def plan_card(r):
    """Carte 'Plan de traitement' : tableau d'analyse d'espace (totaux auto + repositionnement
    incisif déduit du Steiner), chevrons éditables, objectifs/moyens."""
    plan = r.get("plan") or {}
    if not isinstance(plan, dict): plan = {}
    eff = espace_effective(plan, r.get("steiner", {}) or {})   # {k:{p,m}} avec repos déduit + sommes
    src = plan.get("espace", {})
    if not isinstance(src, dict): src = {}
    # --- tableau d'analyse d'espace : 2 colonnes + / − ---
    INP = 'inputmode=decimal autocomplete=off oninput="recalcPlan()" style="width:62px;text-align:center;font-size:13px" placeholder="mm"'
    rows = ""
    for k, lbl, inp in ESPACE_ROWS:
        if inp:
            d = esp_cell(src.get(k))
            pv, mv = (_fmt(d["p"]) if d["p"] else ""), (_fmt(d["m"]) if d["m"] else "")
            note = ""
            if k == "repositionnement_incisive" and not (d["p"] or d["m"]):
                dd = eff.get(k, {})
                pv, mv = (_fmt(dd.get("p")) if dd.get("p") else ""), (_fmt(dd.get("m")) if dd.get("m") else "")
                note = '<div class=muted style="font-size:11px">déduit du Steiner (modifiable)</div>'
            rows += ('<tr><td style="padding:3px 8px">%s%s</td>'
                     '<td style="padding:3px 6px"><input id="esp_p_%s" name="esp_%s_p" value="%s" %s></td>'
                     '<td style="padding:3px 6px"><input id="esp_m_%s" name="esp_%s_m" value="%s" %s></td></tr>'
                     ) % (lbl, note, k, k, pv, INP, k, k, mv, INP)
        else:
            base = k[1:]
            strong = "font-weight:700" if k in ("_ddm", "_total", "_net") else ""
            rows += ('<tr style="background:#f3f6fb;%s"><td style="padding:4px 8px">%s</td>'
                     '<td style="text-align:center"><span id="esp_p_%s">–</span></td>'
                     '<td style="text-align:center"><span id="esp_m_%s">–</span></td></tr>') % (strong, lbl, base, base)
    table = ('<h3 style="margin:4px 0 6px;color:var(--acc)">Analyse d\'espace (Steiner)</h3>'
             '<p class=muted style="font-size:12px;margin:0 0 8px">Reporte chaque ligne dans la colonne <b>+</b> (espace gagné) '
             'ou <b>−</b> (déficit), en mm. DDM, TOTAL et NET se somment automatiquement par colonne.</p>'
             '<table class=st style="max-width:420px"><tr><th style="text-align:left">Mesure</th>'
             '<th style="width:70px">+</th><th style="width:70px">−</th></tr>%s</table>') % rows
    # --- chevrons : flux Problème → (ANB sol / Pog-NB sol) → Solution → Individualisation ---
    chv = plan.get("chevrons", {})
    if not isinstance(chv, dict): chv = {}
    anb = (r.get("steiner", {}) or {}).get("ANB (°)")
    wd = {pk: (pl, pos) for pk, pl, pos in CHEVRON_FLOW}
    def w(pk): return _chevron_widget(pk, wd[pk][0], chv.get(pk, {}), wd[pk][1])
    arrow = '<div style="font-size:26px;color:#9fb0c8;align-self:center">&#8594;</div>'
    flow = ('<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:12px">'
            + w("probleme") + arrow
            + '<div style="display:flex;flex-direction:column;gap:10px">'
            + w("anb_solution") + w("pog_solution") + '</div>'
            + arrow + w("solution") + arrow + w("individualisation")
            + '</div>')
    chev = ('<h3 style="margin:16px 0 6px;color:var(--acc)">Chevrons — plan de traitement</h3>'
            '<p class=muted style="font-size:12px;margin:0 0 8px">Chaque chevron porte ANB, I-NA (mm), I-NB (mm) et Pog-NB — '
            'tout se saisit à la main. Le Problème se sépare en une hypothèse ANB et une hypothèse Pog-NB, '
            'qui fusionnent en Solution, puis Individualisation. La planche ci-dessous rappelle le compromis de Steiner selon l\'ANB.</p>'
            + flow + compromis_chevron_strip(anb))
    # --- stratégies / moyens ---
    obj = str(plan.get("objectifs", "") or "").replace("<", "&lt;")
    moy = str(plan.get("moyens", "") or "").replace("<", "&lt;")
    strat = ('<h3 style="margin:16px 0 6px;color:var(--acc)">Stratégies et moyens thérapeutiques</h3>'
             '<label>Objectifs</label><textarea name=pt_objectifs rows=3 style="width:100%%">%s</textarea>'
             '<label style="margin-top:8px;display:block">Moyens</label>'
             '<textarea name=pt_moyens rows=6 style="width:100%%" '
             'placeholder="1. …\n2. …\n3. …">%s</textarea>') % (obj, moy)
    # --- JS recalcul live (sommes par colonne + / −) ---
    js = ("<script>function _gp(id){var e=document.getElementById(id);if(!e)return 0;"
          "var v=parseFloat((e.value||'').replace(',','.'));return isNaN(v)?0:v;}"
          "var _DDM=['encombrement','repositionnement_incisive','courbe_spee'];"
          "var _TOT=_DDM.concat(['repositionnement_6','espace_derive_mesiale']);"
          "var _NET=_TOT.concat(['extraction_stripping','individualisation']);"
          "function _sum(keys,side){var s=0;keys.forEach(function(k){s+=_gp('esp_'+side+'_'+k);});return Math.round(s*10)/10;}"
          "function _st(id,val){var e=document.getElementById(id);if(e)e.textContent=val;}"
          "function recalcPlan(){[['ddm',_DDM],['total',_TOT],['net',_NET]].forEach(function(g){"
          "_st('esp_p_'+g[0],_sum(g[1],'p'));_st('esp_m_'+g[0],_sum(g[1],'m'));});}"
          "window.addEventListener('DOMContentLoaded',recalcPlan);</script>")
    return '<div class=card><h2>Plan de traitement</h2>%s%s%s%s</div>' % (table, chev, strat, js)

def field_input(name, value, options, ph=""):
    lid = "dl_" + name
    opts = "".join('<option value="%s">' % o for o in options)
    v = (value or "").replace('"', "&quot;")
    return ('<input name="%s" value="%s" list="%s" placeholder="%s" style="width:100%%;font-size:13px" autocomplete=off>'
            '<datalist id="%s">%s</datalist>') % (name, v, lid, ph, lid, opts)

def parse_exam(form):
    ex = {}
    for k, _l, _o in EXAM_ENV + EXAM_FACE:
        v = form.get("ex_" + k, "").strip()
        if v: ex[k] = v
    return ex

def parse_angle(form):
    a = {}
    for k, _l, _o in ANGLE_FIELDS:
        v = form.get("ang_" + k, "").strip()
        if v: a[k] = v
    return a

def parse_suivi(form):
    sv = {}
    for k, _l, _ph, _ml in SUIVI_FIELDS:
        v = form.get("sv_" + k, "").strip()
        if v: sv[k] = v
    return sv

def parse_synthese(form):
    syn = {}
    for _, rk in DIAG_ROWS:
        for _, ck in DIAG_COLS:
            k = ck + "_" + rk; v = form.get("d_" + k, "").strip()
            if v: syn[k] = v
    return syn

def auto_synthese(res):
    """Pré-remplissage squelettique d'après le Steiner (modifiable ensuite)."""
    s = {}
    anb = res.get("ANB (°)")
    if anb is not None:
        s["sq_ap"] = ("Classe I squelettique" if 1.5 <= anb <= 4.5
                      else ("Classe II squelettique (tendance)" if anb > 4.5 else "Classe III squelettique (tendance)"))
    pm = res.get("Plan mandibulaire (Go-Gn/SN) (°)")
    if pm is not None:
        s["sq_ve"] = ("Normodivergent" if 28 <= pm <= 36 else ("Hyperdivergent" if pm > 36 else "Hypodivergent"))
    return s

def apply_word_to_record(r, res):
    """Reporte les valeurs lues d'un bilan Word dans un enregistrement existant.
    - Steiner / analyse d'espace / chevrons : les valeurs lues COMPLÈTENT (mise à jour par clé) ;
    - synthèse : ne remplit que les cases encore vides ;
    - motif / objectifs / moyens : ne remplit que si le champ est vide."""
    r.setdefault("steiner", {}); r.setdefault("synthese", {})
    r["steiner"].update(res.get("steiner", {}) or {})
    syn = dict(res.get("synthese", {}) or {})
    for k, v in auto_synthese(r["steiner"]).items():
        syn.setdefault(k, v)
    for k, v in syn.items():
        if not r["synthese"].get(k):
            r["synthese"][k] = v
    plan = r.get("plan") or {"espace": {}, "chevrons": {}, "objectifs": "", "moyens": ""}
    plan.setdefault("espace", {}); plan.setdefault("chevrons", {})
    for k, v in (res.get("espace", {}) or {}).items():
        plan["espace"][k] = v
    for k, v in (res.get("chevrons", {}) or {}).items():
        plan["chevrons"][k] = v
    if not plan.get("objectifs"): plan["objectifs"] = res.get("objectifs", "")
    if not plan.get("moyens"): plan["moyens"] = res.get("moyens", "")
    r["plan"] = plan
    if not r.get("motif"): r["motif"] = res.get("motif", "")
    return r

def apercu_html(slug, rid, pt, r, base):
    """Rendu HTML du bilan directement dans l'app (aucune dépendance : ni LibreOffice ni PyMuPDF)."""
    def ex(p): return os.path.exists(os.path.join(base, p))
    def u(p): return url_for("rfile", slug=slug, rid=rid, path=p)
    PAGE = 'style="background:#fff;border:1px solid #e3e8ef;border-radius:8px;padding:20px;margin-bottom:16px;box-shadow:0 1px 4px rgba(0,0,0,.06)"'
    IMG = 'style="width:100%;border-radius:4px;border:1px solid #eef"'
    h = []
    # Page photos
    h.append(f'<div {PAGE}>')
    h.append(f'<div style="text-align:center;font:600 16px Times,serif;margin-bottom:12px">{pt["nom"]}, {pt["dob_court"]} ({pt["age"]})</div>')
    def line(paths, dirpref, ext):
        cells = [p for p in paths if ex(f"{dirpref}/{p}.{ext}")]
        if not cells: return ""
        return (f'<div style="display:grid;grid-template-columns:repeat({len(cells)},1fr);gap:8px;margin-bottom:8px">' +
                "".join(f'<img src="{u(f"{dirpref}/{p}.{ext}")}" {IMG}>' for p in cells) + '</div>')
    # ligne exobuccales / ligne occlusales / ligne occlusion ant.+D+G
    h.append(line(["exo_face_repos", "exo_face_sourire", "exo_profil"], "02_photos_traitees", "jpg"))
    h.append(line(["endo_occlusal_maxillaire", "endo_occlusal_mandibulaire"], "02_photos_traitees", "jpg"))
    h.append(line(["endo_occlusion_frontale", "endo_laterale_droite", "endo_laterale_gauche"], "02_photos_traitees", "jpg"))
    h.append('</div>')
    # Page rendus 3D — même configuration
    if any(ex(f"05_stl_rendus/{k}.png") for k, _ in RENDER_FIELDS):
        h.append(f'<div {PAGE}>')
        h.append(line(["upper_occlusal", "lower_occlusal"], "05_stl_rendus", "png"))
        h.append(line(["front_occ", "right_occ", "left_occ"], "05_stl_rendus", "png"))
        h.append(line(["spee_lower", "spee_left"], "05_stl_rendus", "png"))
        h.append('</div>')
    # Page radios
    rad = [k for k in ["panoramique", "teleradiographie_profil"] if ex(f"03_radios/{k}.jpg")]
    if rad:
        h.append(f'<div {PAGE}>' + "".join(f'<img src="{u(f"03_radios/{k}.jpg")}" style="width:100%;border-radius:4px;margin-bottom:10px">' for k in rad) + '</div>')
    # Page identité + tableau Steiner + grille diagnostique
    h.append(f'<div {PAGE}>')
    h.append(f'<div style="font:600 18px sans-serif;color:var(--acc)">{pt["nom"]}</div><div class=muted>Sexe {pt["sexe"]} — {pt["age"]}</div>')
    if r.get("motif"): h.append(f'<p style="margin:8px 0"><b>Motif de consultation :</b> {r["motif"]}</p>')
    # Examen clinique (uniquement les champs renseignés)
    exd = r.get("exam", {}); angd = r.get("angle", {})
    def exam_table(title, fields, data):
        rows = "".join(f'<tr><td class=muted style="width:48%">{lbl}</td><td>{data.get(k)}</td></tr>' for k, lbl, opts in fields if data.get(k))
        return (f'<h4 style="margin:12px 0 4px;color:var(--acc)">{title}</h4><table class=st style="max-width:520px">{rows}</table>') if rows else ""
    h.append(exam_table("Examen — Environnement", EXAM_ENV, exd))
    h.append(exam_table("Examen — Face", EXAM_FACE, exd))
    h.append(exam_table("Classe d'Angle dentaire", ANGLE_FIELDS, angd))
    if r.get("steiner"):
        h.append(f'<h4 style="margin:12px 0 4px;color:var(--acc)">Analyse de Steiner</h4>{steiner_table_html(r)}')
    syn = r.get("synthese", {})
    g = '<h4 style="margin:14px 0 4px;color:var(--acc)">Synthèse diagnostique</h4><table class=st><tr><td></td>' + "".join(f'<th>{cn}</th>' for cn, _ in DIAG_COLS) + '</tr>'
    for rn, rk in DIAG_ROWS:
        g += f'<tr><td class=muted>{rn}</td>' + "".join(f'<td>{(syn.get(ck + "_" + rk, "") or "")}</td>' for cn, ck in DIAG_COLS) + '</tr>'
    h.append(g + '</table></div>')
    # Page plan de traitement
    plan = r.get("plan") or {}
    if not isinstance(plan, dict): plan = {}
    esp = espace_effective(plan, r.get("steiner", {}))
    if any((v.get("p") or v.get("m")) for v in esp.values()) or plan.get("chevrons") or plan.get("objectifs") or plan.get("moyens"):
        h.append(f'<div {PAGE}><h4 style="margin:0 0 8px;color:var(--acc)">Plan de traitement</h4>')
        er = ""
        for k, lbl, inp in ESPACE_ROWS:
            d = esp.get(k) or {}
            p, m = d.get("p", 0), d.get("m", 0)
            if not (p or m): continue
            strong = ' style="font-weight:700"' if k in ("_ddm", "_total", "_net") else ""
            er += (f'<tr{strong}><td class=muted style="width:60%">{lbl}</td>'
                   f'<td style="text-align:center">{_fmt(p) if p else ""}</td>'
                   f'<td style="text-align:center">{_fmt(m) if m else ""}</td></tr>')
        if er: h.append('<h5 style="margin:8px 0 2px">Analyse d\'espace (Steiner)</h5>'
                        f'<table class=st style="max-width:360px"><tr><th style="text-align:left">Mesure</th><th>+</th><th>−</th></tr>{er}</table>')
        chv = plan.get("chevrons", {})
        if not isinstance(chv, dict): chv = {}
        cc = ""
        for pk, pl, pos in CHEVRON_FLOW:
            d = chv.get(pk, {})
            if not isinstance(d, dict): d = {}
            vals = " · ".join("%s %s" % (lab, _fmt(d.get(ck))) for ck, lab in CHEVRON_POS_META if ck in pos and d.get(ck) not in (None, ""))
            if vals: cc += f'<tr><td class=muted>{pl}</td><td>{vals}</td></tr>'
        if cc: h.append(f'<h5 style="margin:10px 0 2px">Chevrons — plan de traitement</h5><table class=st style="max-width:460px">{cc}</table>')
        if plan.get("objectifs"): h.append(f'<h5 style="margin:10px 0 2px">Objectifs</h5><div>{plan["objectifs"]}</div>')
        if plan.get("moyens"): h.append(f'<h5 style="margin:10px 0 2px">Moyens</h5><div style="white-space:pre-line">{plan["moyens"]}</div>')
        h.append('</div>')
    return "".join(h)

def diag_inputs(syn):
    h = '<table style="width:100%;border-collapse:collapse"><tr><td></td>'
    h += "".join('<th style="font-size:12px;color:var(--mut);padding:4px;text-align:center">%s</th>' % cn for cn, _ in DIAG_COLS) + "</tr>"
    for rn, rk in DIAG_ROWS:
        h += '<tr><td style="font-size:12px;color:var(--mut);padding:4px;white-space:nowrap">%s</td>' % rn
        for cn, ck in DIAG_COLS:
            k = ck + "_" + rk
            h += '<td style="padding:3px">%s</td>' % field_input("d_" + k, syn.get(k, ""), DIAG_OPTIONS.get(k, []))
        h += "</tr>"
    return h + "</table>"

def save_uploads(base, files):
    for key, _ in PHOTO_FIELDS:
        f = files.get("photo_" + key)
        if f and f.filename: os.makedirs(os.path.join(base, "01_photos_brutes"), exist_ok=True); f.save(os.path.join(base, "01_photos_brutes", key + ".jpg"))
    for key in P.RADIO_SLOTS:
        f = files.get("radio_" + key)
        if f and f.filename: os.makedirs(os.path.join(base, "03_radios"), exist_ok=True); f.save(os.path.join(base, "03_radios", key + ".jpg"))
    for key in P.STL_SLOTS:
        f = files.get("stl_" + key)
        if f and f.filename: os.makedirs(os.path.join(base, "04_stl_bruts"), exist_ok=True); f.save(os.path.join(base, "04_stl_bruts", key + ".stl"))

# ======================================================================
LOGO_URI = "data:image/png;base64,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"
BASE = """<!doctype html><html lang=fr><head><meta charset=utf-8>
<meta name=viewport content="width=device-width,initial-scale=1"><title>Bilan ODF</title><script>(function(){try{var t=localStorage.getItem('bilan-theme')||'light';document.documentElement.setAttribute('data-theme',t);}catch(e){}})();</script><style>
:root{--bg:#f4f6f9;--card:#ffffff;--card2:#f8fafc;--ink:#0c1526;--ink2:#3a475c;--mut:#6b7891;--acc:#2f6bff;--acc2:#1c4fd6;--accw:#eaf1ff;--teal:#3cc0c8;
--line:#e7ebf1;--line2:#eef1f6;--ok:#169d5b;--bad:#d9483f;
--sb:#ffffff;--sb1:#ffffff;--sb2:#ffffff;--sb-ink:#3a475c;--sb-mut:#6b7891;
--rad:18px;--rad-sm:12px;--rad-xs:9px;
--sh:0 1px 2px rgba(16,24,40,.04),0 8px 24px rgba(16,24,40,.06);
--sh-sm:0 1px 2px rgba(16,24,40,.05)}
html[data-theme="dark"]{--bg:#0a0d13;--card:#12161f;--card2:#0e131b;--ink:#eef2f8;--ink2:#c2cbdb;--mut:#8a96ab;--acc:#4d8bff;--acc2:#8fb6ff;--accw:rgba(77,139,255,.14);--teal:#3cc0c8;
--line:#1d2531;--line2:#161d29;--ok:#4ade80;--bad:#ff8078;
--sb:#12161f;--sb1:#12161f;--sb2:#12161f;--sb-ink:#c2cbdb;--sb-mut:#8a96ab;
--sh:0 1px 2px rgba(0,0,0,.4),0 12px 32px rgba(0,0,0,.35);
--sh-sm:0 1px 2px rgba(0,0,0,.4)}
html{background:var(--bg)}
*{box-sizing:border-box}
body{margin:0;font:15px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;background:var(--bg);color:var(--ink);-webkit-font-smoothing:antialiased}
a{color:var(--acc);text-decoration:none}a:hover{text-decoration:none}
/* ===== SHELL ===== */
.shell{display:flex;min-height:100vh}
.sidebar{width:256px;flex:0 0 256px;min-height:100vh;background:var(--sb);border-right:1px solid var(--line);color:var(--sb-ink);display:flex;flex-direction:column;padding:20px 14px;position:sticky;top:0;height:100vh}
.brand{display:flex;align-items:center;gap:11px;padding:6px 8px 18px}
.brand img{width:40px;height:40px;border-radius:11px;box-shadow:0 4px 12px rgba(0,0,0,.28)}
.brand .bt{font-weight:750;font-size:16px;color:var(--ink);letter-spacing:-.2px}
.brand .bs{font-size:11px;color:var(--sb-mut);margin-top:1px}
.brand .bv{display:inline-block;margin-top:4px;font-size:10px;font-weight:700;letter-spacing:.3px;
  color:var(--acc);background:rgba(47,92,168,.12);border:1px solid rgba(47,92,168,.30);
  border-radius:999px;padding:1px 7px}
.snav{display:flex;flex-direction:column;gap:3px;margin-top:6px}
.snav .lbl{font-size:10.5px;text-transform:uppercase;letter-spacing:1.2px;color:var(--sb-mut);margin:14px 12px 6px;font-weight:700}
.snav a{display:flex;align-items:center;gap:12px;padding:10px 12px;border-radius:11px;color:var(--sb-ink);font-weight:600;font-size:14px;transition:.15s}
.snav a svg{width:20px;height:20px;flex:0 0 20px;opacity:.85}
.snav a:hover{background:var(--card2);color:var(--ink)}
.snav a.active{background:var(--accw);color:var(--acc2);box-shadow:none}
.snav a.active svg{opacity:1}
.sb-foot{margin-top:auto;padding:12px 8px 4px;border-top:1px solid var(--line)}
.sb-foot .u{display:flex;align-items:center;gap:10px}
.sb-foot .av{width:34px;height:34px;border-radius:50%;background:linear-gradient(135deg,var(--acc),var(--teal));color:#fff;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:13px}
.sb-foot .un{font-size:13px;font-weight:700;color:var(--ink)}
.sb-foot .ur{font-size:11px;color:var(--sb-mut)}
.local-badge{display:inline-flex;align-items:center;gap:6px;margin-top:10px;font-size:11px;color:#8fe3c4;background:rgba(23,166,115,.14);padding:5px 9px;border-radius:8px;font-weight:600}
.main{flex:1;min-width:0;display:flex;flex-direction:column}
.topbar{height:64px;display:flex;align-items:center;gap:12px;padding:0 30px;background:var(--card);border-bottom:1px solid var(--line);position:sticky;top:0;z-index:40}
.topbar .search{flex:1;max-width:440px;position:relative;margin:0}
.topbar .search input{width:100%;height:40px;border:1px solid var(--line);border-radius:11px;padding:0 14px 0 40px;font-size:14px;background:var(--card);outline:none;transition:.15s}
.topbar .search input:focus{border-color:var(--acc);box-shadow:0 0 0 3px rgba(30,138,208,.12)}
.topbar .search svg{position:absolute;left:12px;top:11px;width:18px;height:18px;color:var(--mut)}
.topbar .spacer{flex:1}
.content{padding:26px 30px 60px;max-width:1240px;width:100%;margin:0 auto}
.content.wide{max-width:none;margin:0;padding:14px 12px 48px}
/* ===== TYPO ===== */
h1{font-size:25px;margin:0 0 4px;font-weight:800;letter-spacing:-.4px}
h2{font-size:16px;color:var(--ink);font-weight:800;margin:0 0 14px;letter-spacing:-.2px}
h3{font-size:14px;margin:0 0 10px}
h5{font-size:13px;margin:10px 0 4px}
.phead{margin-bottom:20px}
.phead .sub{color:var(--mut);font-size:13.5px;margin-top:2px}
/* ===== CARDS ===== */
.card{background:var(--card);border:1px solid var(--line);border-radius:var(--rad);padding:22px 24px;margin-bottom:20px;box-shadow:var(--sh-sm)}
.cardphotos{padding:12px 12px 10px}
/* ===== BUTTONS ===== */
.btn{display:inline-flex;align-items:center;gap:8px;background:var(--acc);color:#fff;padding:0 16px;height:40px;border-radius:11px;border:0;font-size:14px;font-weight:600;cursor:pointer;box-shadow:0 4px 12px rgba(47,107,255,.28);transition:.15s;white-space:nowrap}
.btn:hover{filter:brightness(1.05);text-decoration:none}
.btn.sec{background:var(--card);color:var(--ink);border:1px solid var(--line);box-shadow:none}
.btn.sec:hover{border-color:#cdd9e6;box-shadow:var(--sh-sm);filter:none}
.btn.sm{height:34px;padding:0 12px;font-size:13px;border-radius:9px;box-shadow:none}
/* ===== FORMS ===== */
input,select,textarea{width:100%;padding:0 13px;height:42px;border:1px solid var(--line);border-radius:10px;font-size:14px;background:var(--card);color:var(--ink);outline:none;transition:.15s;font-family:inherit}
textarea{height:auto;padding:11px 13px;min-height:84px;resize:vertical}
input:focus,select:focus,textarea:focus{border-color:var(--acc);box-shadow:0 0 0 3px rgba(30,138,208,.12)}
label{font-size:12.5px;color:var(--ink);font-weight:700;display:block;margin:0 0 6px}
/* ===== GRIDS ===== */
.grid{display:grid;gap:14px}.g2{grid-template-columns:1fr 1fr}.g3{grid-template-columns:1fr 1fr 1fr}.g4{grid-template-columns:repeat(4,1fr)}.grid3{grid-template-columns:1fr 1fr 1fr;gap:14px}
/* ===== PATIENT GRID ===== */
.plist{display:grid;grid-template-columns:repeat(auto-fill,minmax(228px,1fr));gap:18px}
.pcard{border:1px solid var(--line);border-radius:var(--rad);overflow:visible;background:var(--card);transition:.18s;display:block;position:relative;box-shadow:var(--sh-sm)}
.pcard:hover{box-shadow:var(--sh);transform:translateY(-3px);border-color:#d9e3ee}
.pcard:focus-within{z-index:60}
.pcard:has(details.menu[open]){z-index:60}
.pcard .thumb{position:relative;overflow:hidden;height:176px;background-color:var(--card2);display:flex;align-items:center;justify-content:center;color:#9fb2c6;font-size:38px;border-radius:var(--rad) var(--rad) 0 0}
.pcard .thumb::before{content:"";position:absolute;inset:0;background:var(--u) center/cover no-repeat;filter:blur(20px) brightness(.97);transform:scale(1.18)}
.pcard .thumb::after{content:"";position:absolute;inset:0;background:var(--u) center/contain no-repeat}
.pcard .body{padding:13px 15px 15px}.pcard .nm{font-weight:700;font-size:15px;letter-spacing:-.2px}.pcard .mt{font-size:12.5px;color:var(--mut)}
.plist.icone{grid-template-columns:repeat(auto-fill,minmax(172px,1fr))}
.pcard.nopic .tic{font-size:34px;text-align:center;margin:6px 0 2px;color:#8fb0d4}
.pcard.nopic .body{padding:14px 15px 8px;text-align:center}
.plist.liste{display:block}
.lrow{position:relative;display:flex;align-items:center;gap:14px;padding:12px 16px;border:1px solid var(--line);border-radius:12px;background:var(--card);margin-bottom:8px;box-shadow:var(--sh-sm);transition:.15s}
.lrow:hover{border-color:#d9e3ee;box-shadow:var(--sh)}
.lrow .lnm{font-weight:700;font-size:15px;color:inherit;text-decoration:none;flex:1;min-width:140px}
.lrow .lage{color:var(--mut);font-size:13px;min-width:70px}
.lrow .lstat{margin-left:auto;margin-right:34px}
.lrow:has(details.menu[open]){z-index:60}
/* ===== MENUS ===== */
.menu{position:relative}
.menu>summary{list-style:none;cursor:pointer}
.menu>summary::-webkit-details-marker{display:none}
.menu>summary::marker{content:""}
.dots{width:30px;height:30px;border-radius:9px;background:rgba(255,255,255,.9);color:#3a4a5e;display:flex;align-items:center;justify-content:center;font-size:20px;font-weight:900;box-shadow:0 2px 6px rgba(0,0,0,.14)}
.dots:hover{background:#fff;color:#24405f}
.menu-pop{position:absolute;z-index:30;top:calc(100% + 4px);right:0;background:var(--card);border:1px solid var(--line);border-radius:12px;box-shadow:0 12px 34px rgba(16,40,70,.20);padding:6px;min-width:190px}
.menu-pop .mi{display:flex;align-items:center;gap:9px;padding:8px 11px;border-radius:8px;color:var(--ink);font-size:13.5px;white-space:nowrap;cursor:pointer}
.menu-pop .mi:hover{background:var(--card2)}
.menu-pop .sep{height:1px;background:var(--line2);margin:4px 2px}
.mi{display:flex;align-items:center;gap:9px;padding:9px 12px;color:var(--ink);font-size:14px}
.mi:hover{background:var(--card2)}
/* ===== TABLES ===== */
.mtab{border-collapse:collapse;width:100%;font-size:13.5px}
.mtab th,.mtab td{text-align:left;padding:11px 12px;border-bottom:1px solid var(--line2);vertical-align:top}
.mtab th{color:var(--mut);font-weight:700;font-size:11.5px;text-transform:uppercase;letter-spacing:.05em;border-bottom:1px solid var(--line)}
.mtab tbody tr:hover{background:var(--card2)}
table.st{width:100%;border-collapse:collapse}table.st td{padding:6px 8px;border-bottom:1px solid var(--line2);font-size:13px}
/* ===== MISC ===== */
.flash{background:#fff6d8;border:1px solid #f0d97a;padding:11px 15px;border-radius:11px;margin-bottom:16px}
.err{background:#fdecea;border-color:#f5b3ab}
.tag{display:inline-block;background:var(--accw);color:var(--acc2);border-radius:20px;padding:3px 11px;font-size:12px;font-weight:600;margin:2px 3px}
.muted{color:var(--mut);font-size:13px}.center{text-align:center}
.thumbs{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}.thumbs img{width:100%;border-radius:10px;border:1px solid var(--line)}
.ph{position:relative}.ph img{width:100%;display:block;border-radius:10px;border:1px solid var(--line)}
.ph .ctrl{position:absolute;top:5px;right:5px;display:flex;gap:3px}
.ph .ctrl a{background:rgba(255,255,255,.92);border:1px solid var(--line);border-radius:7px;padding:1px 6px;font-size:13px}
.timeline{display:flex;flex-direction:column;gap:12px}
.rec{border:1px solid var(--line);border-radius:12px;padding:14px 16px;display:flex;align-items:center;gap:14px;background:var(--card)}
.rec .dot{width:10px;height:10px;border-radius:50%;background:var(--acc);flex:none}
.rec .rt{font-weight:700}.rec .rm{font-size:12px;color:var(--mut)}
.reclist{display:flex;flex-direction:column}
.recrow{display:flex;align-items:center;gap:13px;padding:12px 4px;border-bottom:1px solid var(--line2)}
.recrow:last-child{border-bottom:none}
.recdot{width:9px;height:9px;border-radius:50%;background:var(--acc);flex:none}
.recname{font-weight:700;font-size:14.5px;line-height:1.25}
.recsub{font-size:12.5px;color:var(--mut);margin-top:1px}
.recrow .dots{width:28px;height:28px;font-size:17px;box-shadow:none;background:var(--card2);border-radius:8px}
.recrow .dots:hover{background:#e2e9f5}
.spin{display:inline-block;width:15px;height:15px;border:3px solid #cdd;border-top-color:var(--acc);border-radius:50%;animation:s 1s linear infinite;vertical-align:-2px}@keyframes s{to{transform:rotate(360deg)}}
.sub{border:1px solid var(--line);border-radius:12px;padding:16px;margin-bottom:14px}.sub h3{margin:0 0 10px;font-size:14px}
/* ===== LOGIN ===== */
.loginwrap{min-height:100vh;display:flex;align-items:center;justify-content:center;background:linear-gradient(135deg,#0d2135 0%,#123a5c 55%,#186a92 100%);padding:20px}
.loginwrap .card{width:400px;max-width:100%;border:none;border-radius:22px;box-shadow:0 30px 80px rgba(0,0,0,.35);padding:38px 36px;margin:0}
.loginwrap .card h1{text-align:center}.loginwrap .card h2{text-align:center;color:var(--mut);font-weight:600}
.loginwrap .btn{width:100%;justify-content:center;height:46px;font-size:15px;margin-top:6px}
@media(max-width:820px){.sidebar{display:none}.g4{grid-template-columns:1fr 1fr}}
</style></head><body>
{% if session.get('user') %}
{% set ep=request.endpoint %}
<div class="shell">
  <aside class="sidebar">
    <div class="brand"><img src="{{logo}}" alt=""><div><div class="bt">Bilan ODF</div><div class="bs">Gestion locale</div><div class="bv">v{{version}}</div></div></div>
    <nav class="snav">
      <div class="lbl">Patient</div>
      <a href="{{url_for('dashboard')}}" class="{{'active' if ep in ('dashboard','suivi','patient','record','evolution','patient_edit') else ''}}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="3" width="7" height="7" rx="1.5"/><rect x="3" y="14" width="7" height="7" rx="1.5"/><rect x="14" y="14" width="7" height="7" rx="1.5"/></svg> Bibliothèque</a>
      <a href="{{url_for('nouveau')}}" class="{{'active' if ep in ('nouveau','creer') else ''}}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg> Nouveau patient</a>
      <a href="{{url_for('import_massif')}}" class="{{'active' if ep and ('massif' in ep or ep=='word_import') else ''}}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/></svg> Importer patient / bilan (Word)</a>
      <a href="{{url_for('appareils_home')}}" class="{{'active' if ep and ep.startswith('appareil') else ''}}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 8v2a8 8 0 0 0 16 0V8"/><path d="M4 10h16"/><path d="M8 10v2M12 10v3M16 10v2"/></svg> Appareils</a>
      <div class="lbl">Applications tierces</div>
      <a href="#" onclick="try{window.webkit.messageHandlers.bilan.postMessage('open_xero')}catch(e){}; return false;"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="9"/><path d="M12 3v18M4.5 8h15M4.5 16h15"/></svg> XERO</a>
      <a href="#" onclick="try{window.webkit.messageHandlers.bilan.postMessage('open_webceph')}catch(e){}; return false;"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 20V6a2 2 0 0 1 2-2h9l5 5v11a2 2 0 0 1-2 2z"/><circle cx="11" cy="12" r="3"/></svg> WebCeph</a>
      <a href="{{url_for('radios')}}" class="{{'active' if ep=='radios' else ''}}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="16" rx="2"/><line x1="8" y1="4" x2="8" y2="20"/><line x1="16" y1="4" x2="16" y2="20"/></svg> Radios &amp; tracés</a>
      <a href="{{url_for('steiner_queue')}}" class="{{'active' if ep=='steiner_queue' else ''}}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 3v18h18"/><path d="M18 8l-5 5-3-3-4 4"/></svg> Analyses Steiner</a>
      <a href="#" onclick="try{window.webkit.messageHandlers.bilan.postMessage('open_doctolib')}catch(e){}; return false;"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg> Doctolib</a>
      <div class="lbl">Réglages</div>
      <a href="{{url_for('reglages')}}" class="{{'active' if ep=='reglages' else ''}}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg> Réglages</a>
      <a href="{{url_for('logout')}}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg> Déconnexion</a>
    </nav>
    <div class="sb-foot">
      <div class="u"><div class="av">{{ session['user'][:2]|upper }}</div><div><div class="un">{{ session['user'] }}</div><div class="ur">Praticien</div></div></div>
      <div class="local-badge">🔒 Sauvegarde locale</div>
    </div>
  </aside>
  <div class="main">
    <header class="topbar">
      <form class="search" method="get" action="{{url_for('dashboard')}}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16.5" y2="16.5"/></svg><input name="q" placeholder="Rechercher un patient…" value="{{ request.args.get('q','') }}"></form>
      <div class="spacer"></div>
      <button type="button" class="btn sec" onclick="(function(){var h=document.documentElement,n=h.getAttribute('data-theme')==='dark'?'light':'dark';h.setAttribute('data-theme',n);try{localStorage.setItem('bilan-theme',n);}catch(e){}})()" title="Clair / sombre" style="width:40px;height:40px;padding:0;justify-content:center;margin-right:2px"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" style="width:18px;height:18px"><path d="M21 12.8A8.5 8.5 0 1 1 11.2 3a6.5 6.5 0 0 0 9.8 9.8Z"/></svg></button>
      <a class="btn" href="{{url_for('nouveau')}}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" style="width:17px;height:17px"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg> Nouveau patient</a>
    </header>
    <div class="content{{' wide' if wide else ''}}">{% with m=get_flashed_messages() %}{% if m %}<div class=flash>{{m|join(' · ')}}</div>{% endif %}{% endwith %}{{body|safe}}</div>
  </div>
</div>
{% else %}
<div class="loginwrap"><div style="width:400px;max-width:100%">{% with m=get_flashed_messages() %}{% if m %}<div class=flash>{{m|join(' · ')}}</div>{% endif %}{% endwith %}{{body|safe}}</div></div>
{% endif %}
<script>function openPlat(which,el,ev){if(ev){ev.preventDefault();ev.stopPropagation();}var nom=(el&&el.getAttribute('data-nom'))||'';try{window.webkit.messageHandlers.bilan.postMessage('open_'+which+':'+nom);}catch(e){}return false;}
// Ferme tout menu déroulant (badge statut, ⋮) quand on clique ailleurs, sans rien changer.
document.addEventListener('click',function(e){document.querySelectorAll('details.menu[open]').forEach(function(d){if(!d.contains(e.target))d.removeAttribute('open');});},true);</script>
<div id=majModal style="display:none;position:fixed;inset:0;background:rgba(8,12,20,.55);z-index:99999;align-items:center;justify-content:center">
  <div style="background:var(--card,#fff);color:var(--ink,#1f2a37);max-width:450px;width:92%;border-radius:16px;padding:24px;box-shadow:0 24px 70px rgba(0,0,0,.4);border:1px solid var(--line,#e5e7eb)">
    <div style="display:flex;align-items:center;gap:11px;margin-bottom:10px">
      <span style="width:38px;height:38px;border-radius:11px;background:linear-gradient(135deg,#16324f,#2f5ca8);color:#fff;display:inline-flex;align-items:center;justify-content:center;font-size:20px">&#8635;</span>
      <div style="font-size:18px;font-weight:800">Mise &agrave; jour disponible</div>
    </div>
    <div id=majTxt style="font-size:14px;color:var(--mut,#6b7280);line-height:1.55;margin-bottom:16px;white-space:pre-line"></div>
    <div id=majMsg style="font-size:13px;color:var(--acc,#2f5ca8);margin-bottom:12px;display:none"></div>
    <div style="display:flex;gap:10px;justify-content:flex-end">
      <button type=button class="btn sec" onclick="majLater()">Plus tard</button>
      <button type=button class="btn" onclick="majApply(this)">Installer et relancer</button>
    </div>
  </div>
</div>
<script>
(function(){var tries=0;function show(d){var t=document.getElementById('majTxt');
    var s='Une nouvelle version (v'+d.version+') de Bilan ODF est disponible';
    if(d.current)s+=' (vous avez la v'+d.current+')';s+='.';
    if(d.notes)s+='\n\n'+d.notes;
    s+='\n\nInstaller maintenant et relancer l\'application ?';
    t.textContent=s;document.getElementById('majModal').style.display='flex';}
  function chk(){fetch('/maj/state').then(function(r){return r.json();}).then(function(d){
    if(d&&d.version){show(d);} else if(tries++<8){setTimeout(chk,3000);}
  }).catch(function(){if(tries++<8)setTimeout(chk,3000);});}
  try{chk();}catch(e){}})();
window.majLater=function(){document.getElementById('majModal').style.display='none';};
window.majApply=function(btn){btn.disabled=true;var m=document.getElementById('majMsg');m.style.display='block';m.textContent='Téléchargement et installation...';
  fetch('/maj/apply',{method:'POST'}).then(function(r){return r.text();}).then(function(t){
    if(t==='ok'){m.textContent='Installation terminée. Redémarrage en cours, la page se recharge dans ~10 s...';setTimeout(function(){location.reload();},10000);}
    else{m.textContent='Échec : '+t;btn.disabled=false;}
  }).catch(function(){m.textContent='Redémarrage en cours... rechargement dans ~10 s';setTimeout(function(){location.reload();},10000);});};
</script>
</body></html>"""
def page(body, wide=False): return render_template_string(BASE, body=body, wide=wide, logo=LOGO_URI, version=APP_VERSION)


@app.route("/radios")
def radios():
    if not logged(): return redirect(url_for("login"))
    import glob as _glob, json as _json
    inbox = os.path.join(DATA, "_radios_inbox")
    files = []
    if os.path.isdir(inbox):
        files = [os.path.basename(f) for f in _glob.glob(os.path.join(inbox, "*"))
                 if os.path.isfile(f) and f.lower().endswith((".jpg", ".jpeg", ".png", ".pdf", ".webp"))]
        files.sort(key=lambda f: os.path.getmtime(os.path.join(inbox, f)), reverse=True)
    recs = {}; names = {}
    for slug in sorted(os.listdir(PATIENTS)):
        pdir = os.path.join(PATIENTS, slug); rd = os.path.join(pdir, "records")
        if not os.path.isdir(rd): continue
        try: nom = _json.load(open(os.path.join(pdir, "patient.json"), encoding="utf-8")).get("nom", slug)
        except Exception: nom = slug
        lst = []
        for rid in sorted(os.listdir(rd)):
            try: lbl = _json.load(open(os.path.join(rd, rid, "meta.json"), encoding="utf-8")).get("label", rid)
            except Exception: lbl = rid
            lst.append({"rid": rid, "label": lbl})
        if lst: recs[slug] = lst; names[slug] = nom
    popts = "".join('<option value="%s">%s</option>' % (sl, names[sl]) for sl in sorted(names, key=lambda x: names[x].lower()))
    fopts = "".join('<option value="%s">%s</option>' % (f, f) for f in files)
    finfo = ("%d radio(s) telecharged(s) en attente." % len(files)) if files else "Aucune radio telechargee pour l'instant. Ouvre un site, connecte-toi, telecharge une radio : elle apparaitra ici."
    import json as _j
    _HEAD = '<div class=phead><h1>Radios &amp; tracés à classer</h1><div class=sub>Récupérer une radio (XERO) ou un tracé (WebCeph) et le placer sur un bilan</div></div>'
    _STYLE = '<style>.rcard{background:var(--card);border:1px solid var(--line);border-radius:var(--rad);padding:22px 24px;margin-bottom:20px;box-shadow:var(--sh-sm)}.rhint{margin-top:16px;padding:12px 14px;background:var(--card2);border:1px solid var(--line);border-radius:12px;font-size:13px;color:var(--mut)}.inbgal{display:flex;gap:12px;flex-wrap:wrap;margin:6px 0 2px}.inbthumb{position:relative;cursor:pointer;border:2px solid var(--line);border-radius:12px;overflow:hidden;width:154px;background:var(--card2);transition:.15s}.inbthumb:hover{border-color:#cdd9e6}.inbthumb img{display:block;width:154px;height:112px;object-fit:cover}.inbthumb .cap{font-size:11px;color:var(--mut);padding:5px 8px;text-align:center;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.inbthumb input{position:absolute;opacity:0;pointer-events:none}.inbthumb:has(input:checked){border-color:var(--acc);box-shadow:0 0 0 3px rgba(30,138,208,.16)}.inbthumb:has(input:checked)::after{content:"";position:absolute;top:7px;right:7px;width:20px;height:20px;border-radius:50%;background:var(--acc);border:2px solid #fff;box-shadow:0 1px 3px rgba(0,0,0,.3)}.rempty{padding:26px;text-align:center;color:var(--mut);background:var(--card2);border:1px dashed var(--line);border-radius:12px}.curradio{display:inline-block;margin:10px 14px 0 0;vertical-align:top;text-align:center}.curradio img{max-height:150px;max-width:230px;border:1px solid var(--line);border-radius:10px;display:block}</style>'
    _C1 = '<div class=rcard><div style="display:flex;gap:16px;align-items:center;flex-wrap:wrap"><div style="flex:0 0 auto;width:50px;height:50px;border-radius:14px;background:linear-gradient(135deg,var(--acc),var(--teal));display:flex;align-items:center;justify-content:center"><svg width=26 height=26 viewBox="0 0 24 24" fill=none stroke=#fff stroke-width=2><rect x=3 y=4 width=18 height=16 rx=2></rect><line x1=8 y1=4 x2=8 y2=20></line><line x1=16 y1=4 x2=16 y2=20></line></svg></div><div style="flex:1;min-width:220px"><div style="font-weight:800;font-size:15px;color:var(--ink)">XERO &mdash; CHU Nice</div><div class=muted>Le site s\'ouvre dans une fenetre integree a l\'app et garde ta session. Tes identifiants ne sont jamais enregistres.</div></div><a class=btn href="#" onclick="try{window.webkit.messageHandlers.bilan.postMessage(`open_xero`)}catch(e){}; return false;">Ouvrir XERO</a></div><div class=rhint><b style="color:var(--ink)">Comment faire</b> &mdash; connecte-toi, affiche la radio du patient, puis clique le bouton bleu &laquo; Recuperer cette radio &raquo; en bas de la fenetre. Elle apparaitra ci-dessous.</div></div>'
    _C2A = '<div class=rcard><div style="display:flex;justify-content:space-between;align-items:center;gap:12px;margin-bottom:8px"><h2 style="margin:0">Radios recuperees</h2>'
    _C2B = '</div><p class=muted style="margin:0 0 4px">Choisis la radio a placer, puis le patient et le temps.</p><form method=post action="/radios/attach" enctype=multipart/form-data>'
    _C2C = '<div class="grid g2" style="margin-top:16px"><div><label>Patient</label><select id=selp name=slug onchange=fillRecs()>'
    _C2D = '</select></div><div><label>Reevaluation</label><select id=selr name=rid onchange=loadCur()></select></div><div><label>Type de radio</label><select name=type><option value=panoramique>Panoramique</option><option value=teleradiographie_profil>Teleradiographie de profil</option><option value=trace_cephalo>Tracé céphalométrique</option></select></div><div><label>... ou importer un fichier</label><input type=file name=file accept="image/*"></div></div><div style="margin-top:16px"><button class=btn type=submit>Attacher la radio</button></div></form><div id=curradios style="margin-top:16px"></div></div>'
    _EMPTY = "<div class=rempty>Aucune radio recuperee pour l'instant.<br>Ouvre XERO, affiche une radio et clique &laquo; Recuperer cette radio &raquo;.</div>"
    _JSA = '<script>var RECS='
    _JSB = ';\nfunction fillRecs(){var s=document.getElementById(\'selp\').value;var rs=document.getElementById(\'selr\');rs.innerHTML=\'\';(RECS[s]||[]).forEach(function(r){var o=document.createElement(\'option\');o.value=r.rid;o.textContent=r.label;rs.appendChild(o);});loadCur();}\nfunction loadCur(){var s=document.getElementById(\'selp\').value,r=document.getElementById(\'selr\').value;var box=document.getElementById(\'curradios\');if(!s||!r){box.innerHTML=\'\';return;}\nfetch(\'/radios/record_radios?slug=\'+encodeURIComponent(s)+\'&rid=\'+encodeURIComponent(r)).then(function(x){return x.json();}).then(function(d){var LBL={panoramique:\'Panoramique\',teleradiographie_profil:\'Teleradiographie de profil\'};var h=\'\';Object.keys(d).forEach(function(k){h+=\'<div class=curradio><div class=muted style="font-size:12px;margin-bottom:4px">\'+LBL[k]+\'</div><img src="\'+d[k]+\'"><form method=post action="/radios/delete" style="margin:6px 0 0"><input type=hidden name=slug value="\'+s+\'"><input type=hidden name=rid value="\'+r+\'"><input type=hidden name=type value="\'+k+\'"><button class="btn sec" type=submit style="font-size:12px;padding:4px 10px">Supprimer</button></form></div>\';});box.innerHTML=h?(\'<div style="font-weight:700;font-size:13px;margin:6px 0 2px;color:var(--ink)">Deja sur ce bilan</div>\'+h):\'\';});}\nwindow.addEventListener(\'load\',function(){fillRecs();});</script>'
    _vider = ('<form method=post action="/radios/clear_inbox" style="margin:0"><button class="btn sec" type=submit>Vider la file</button></form>') if files else ''
    _thumbs = ''
    for _f in files:
        _thumbs += '<label class=inbthumb><input type=radio name=inbox value="' + _f + '"><img src="/radios/inbox_file/' + _f + '"><div class=cap>' + _f + '</div></label>'
    _inbgal = ('<div class=inbgal>' + _thumbs + '</div>') if files else _EMPTY
    body = _HEAD + _STYLE + _C1 + _C2A + _vider + _C2B + _inbgal + _C2C + popts + _C2D + _JSA + _j.dumps(recs) + _JSB
    return page(body)


RADIO_CAP_TOKEN = secrets.token_urlsafe(16)


@app.route("/radios/inbox_file/<path:name>")
def radios_inbox_file(name):
    if not logged(): abort(403)
    fp = os.path.join(DATA, "_radios_inbox", os.path.basename(name))
    if not os.path.exists(fp): abort(404)
    return send_file(fp)


@app.route("/radios/clear_inbox", methods=["POST"])
def radios_clear_inbox():
    if not logged(): return redirect(url_for("login"))
    import glob as _g
    inbox = os.path.join(DATA, "_radios_inbox"); n = 0
    if os.path.isdir(inbox):
        for f in _g.glob(os.path.join(inbox, "*")):
            try: os.remove(f); n += 1
            except Exception: pass
    flash("%d radio(s) en attente supprimee(s)." % n)
    return redirect(url_for("radios"))


@app.route("/radios/record_radios")
def radios_record_radios():
    if not logged(): return jsonify({})
    slug = request.args.get("slug", ""); rid = request.args.get("rid", "")
    try: base = rdir(slug, rid)
    except Exception: return jsonify({})
    out = {}
    for t in ("panoramique", "teleradiographie_profil"):
        fp = os.path.join(base, "03_radios", t + ".jpg")
        if os.path.exists(fp):
            try: cb = int(os.path.getmtime(fp))
            except Exception: cb = 0
            out[t] = url_for("rfile", slug=slug, rid=rid, path="03_radios/%s.jpg" % t) + ("?t=%d" % cb)
    return jsonify(out)


@app.route("/radios/delete", methods=["POST"])
def radios_delete():
    if not logged(): return redirect(url_for("login"))
    slug = request.form.get("slug", ""); rid = request.form.get("rid", ""); typ = request.form.get("type", "")
    if typ not in ("panoramique", "teleradiographie_profil", "trace_cephalo"):
        flash("Type de radio invalide."); return redirect(url_for("radios"))
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r:
        flash("Patient ou reevaluation introuvable."); return redirect(url_for("radios"))
    fp = os.path.join(rdir(slug, rid), "03_radios", typ + ".jpg")
    if os.path.exists(fp):
        try: os.remove(fp)
        except Exception: pass
        try: _regenerate(slug, rid, pt, r, skip_stl=True)
        except Exception: pass
        flash("Radio supprimee.")
    else:
        flash("Aucune radio de ce type sur cette reevaluation.")
    return redirect(url_for("radios"))


@app.route("/radios/capture", methods=["POST"])
def radios_capture():
    # recoit l'image affichee dans la fenetre XERO (captureee par le bouton injecte) -> inbox
    if request.form.get("t", "") != RADIO_CAP_TOKEN: return ("", 403)
    d = request.form.get("img", "")
    if not d.startswith("data:"): return ("", 400)
    import base64, time
    try: raw = base64.b64decode(d.split(",", 1)[1])
    except Exception: return ("", 400)
    inbox = os.path.join(DATA, "_radios_inbox"); os.makedirs(inbox, exist_ok=True)
    open(os.path.join(inbox, "radio_%d.jpg" % int(time.time() * 1000)), "wb").write(raw)
    return ("ok", 200)


@app.route("/radios/attach", methods=["POST"])
def radios_attach():
    if not logged(): return redirect(url_for("login"))
    from PIL import Image as _I, ImageOps as _IO
    slug = request.form.get("slug", ""); rid = request.form.get("rid", ""); typ = request.form.get("type", "")
    if typ not in ("panoramique", "teleradiographie_profil", "trace_cephalo"):
        flash("Type de radio invalide."); return redirect(url_for("radios"))
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r:
        flash("Patient ou reevaluation introuvable."); return redirect(url_for("radios"))
    base = rdir(slug, rid)
    if typ == "trace_cephalo":
        import time as _t
        os.makedirs(os.path.join(base, "06_webceph"), exist_ok=True)
        dest = os.path.join(base, "06_webceph", "trace_%d.jpg" % int(_t.time() * 1000))
    else:
        os.makedirs(os.path.join(base, "03_radios"), exist_ok=True)
        dest = os.path.join(base, "03_radios", typ + ".jpg")
    saved = False; used_inbox = None
    up = request.files.get("file")
    if up and up.filename:
        try:
            _IO.exif_transpose(_I.open(up.stream)).convert("RGB").save(dest, quality=92); saved = True
        except Exception as e:
            flash("Image illisible : %s" % e)
    else:
        fn = os.path.basename(request.form.get("inbox", ""))
        srcp = os.path.join(DATA, "_radios_inbox", fn)
        if fn and os.path.exists(srcp):
            try:
                _IO.exif_transpose(_I.open(srcp)).convert("RGB").save(dest, quality=92); saved = True; used_inbox = srcp
            except Exception as e:
                flash("Image illisible : %s (les PDF ne sont pas encore geres, exporte en image)" % e)
        else:
            flash("Choisis une radio telechargee ou importe un fichier.")
    if saved:
        try: _regenerate(slug, rid, pt, r, skip_stl=True)
        except Exception: pass
        if used_inbox:
            try: os.remove(used_inbox)
            except Exception: pass
        lbl = {"panoramique": "Panoramique", "teleradiographie_profil": "Teleradiographie de profil", "trace_cephalo": "Trace cephalometrique"}[typ]
        _w = "Trace" if typ == "trace_cephalo" else "Radio"
        flash("%s (%s) attache a %s." % (_w, lbl, (pt.get("nom") or slug)))
    return redirect(url_for("radios"))


def form_fields(rec=None, with_identity=True):
    """champs upload + steiner (partagés nouveau/réévaluation/édition)."""
    steiner = (rec or {}).get("steiner", {})
    ident = ""
    if with_identity:
        ident = """<div class=card><h2>Identité</h2><div class="grid g3">
        <div><label>Nom</label><input name=nom required></div>
        <div><label>Prénom</label><input name=prenom></div>
        <div><label>Date de naissance</label><input name=dob type=date></div>
        <div><label>Sexe</label><select name=sexe><option>non précisé</option><option>féminin</option><option>masculin</option></select></div></div></div>"""
    photos = "".join("""<div><label>%s</label><input type=file name="photo_%s" accept="image/*"></div>""" % (l, k) for k, l in PHOTO_FIELDS)
    stein = "".join("""<div><label>%s <span class=muted>(%g±%g)</span></label>
        <input name="st_%s" inputmode="decimal" value="%s" placeholder="résultat"></div>"""
        % (short, mean, sd, nom, (("%g" % steiner[nom]) if nom in steiner else "")) for nom, short, mean, sd in STEINER_FIELDS)
    return """%s
    <div class=card><h2>Photos cliniques</h2><div class="grid g4">%s</div>
      <p class=muted style=margin-top:8px>Les latérales prises en portrait sont pivotées automatiquement ; tu pourras aussi les faire pivoter à la main sur la fiche.</p></div>
    <div class=card><h2>Radiographies</h2><div class="grid g2">
      <div><label>Panoramique</label><input type=file name=radio_panoramique accept="image/*"></div>
      <div><label>Téléradiographie de profil</label><input type=file name=radio_teleradiographie_profil accept="image/*"></div></div></div>
    <div class=card><h2>Modèles 3D (STL)</h2><div class="grid g2">
      <div><label>Arcade supérieure (UpperJawScan.stl)</label><input type=file name=stl_UpperJawScan accept=".stl"></div>
      <div><label>Arcade inférieure (LowerJawScan.stl)</label><input type=file name=stl_LowerJawScan accept=".stl"></div></div></div>
    <div class=card><h2>Analyse de Steiner <span class=muted>— valeurs WebCeph</span></h2><div class="grid g3">%s</div></div>""" % (ident, photos, stein)

# ======================================================================
#  AUTH
# ======================================================================
@app.route("/setup", methods=["GET", "POST"])
def setup():
    if cfg["users"]: return redirect(url_for("login"))
    if request.method == "POST":
        u = request.form.get("user", "").strip(); p = request.form.get("pwd", "")
        if u and len(p) >= 4:
            cfg["users"][u] = generate_password_hash(p); save_config(cfg); session["user"] = u
            return redirect(url_for("dashboard"))
        flash("Identifiant requis, mot de passe ≥ 4 caractères.")
    return page("""<div class=card style="max-width:420px;margin:8vh auto"><h1>Bienvenue</h1>
    <h2>Créez le premier compte de cette machine</h2><form method=post class=grid>
    <div><label>Identifiant</label><input name=user autofocus></div>
    <div><label>Mot de passe</label><input name=pwd type=password></div>
    <button class=btn>Créer le compte</button></form></div>""")

def _app_creds_get():
    import subprocess, json as _j
    try:
        r = subprocess.run(["security", "find-generic-password", "-s", "BilanODF-App", "-a", "app", "-w"],
                           capture_output=True, text=True, timeout=6)
        if r.returncode != 0: return {}
        raw = (r.stdout or "").strip()
        return _j.loads(raw) if raw else {}
    except Exception:
        return {}

def _app_creds_set(u, p):
    import subprocess, json as _j
    try:
        subprocess.run(["security", "add-generic-password", "-U", "-s", "BilanODF-App", "-a", "app", "-w", _j.dumps({"id": u, "pw": p})],
                       capture_output=True, text=True, timeout=6)
    except Exception:
        pass

def _app_creds_clear():
    import subprocess
    try:
        subprocess.run(["security", "delete-generic-password", "-s", "BilanODF-App", "-a", "app"],
                       capture_output=True, text=True, timeout=6)
    except Exception:
        pass

@app.route("/login", methods=["GET", "POST"])
def login():
    if not cfg["users"]: return redirect(url_for("setup"))
    if request.method == "POST":
        u = request.form.get("user", "").strip(); p = request.form.get("pwd", "")
        if cfg["users"].get(u) and check_password_hash(cfg["users"][u], p):
            if request.form.get("remember"): _app_creds_set(u, p)
            else: _app_creds_clear()
            session["user"] = u; return redirect(url_for("dashboard"))
        flash("Identifiant ou mot de passe incorrect.")
    _sc = _app_creds_get()
    _su = (_sc.get("id", "") or "").replace('"', "&quot;")
    _sp = (_sc.get("pw", "") or "").replace('"', "&quot;")
    _ck = "checked" if (_su or _sp) else ""
    return page("""<div class=card><div style="text-align:center;margin-bottom:8px"><img src="%s" style="width:72px;height:72px;border-radius:19px;box-shadow:0 10px 26px rgba(18,58,92,.28)"></div>
    <h1>Bilan ODF</h1><h2>Gestion locale des bilans orthodontiques</h2>
    <form method=post class=grid><div><label>Identifiant</label><input name=user autofocus value="%s"></div>
    <div><label>Mot de passe</label><input name=pwd type=password value="%s"></div>
    <label style="display:flex;align-items:center;gap:8px;font-size:13px;color:var(--mut);font-weight:500;margin:2px 0 2px"><input type=checkbox name=remember %s style="width:auto;margin:0"> Se souvenir de mes identifiants sur cet ordinateur</label>
    <button class=btn>Se connecter</button></form>
    <div style="margin-top:16px;text-align:center;font-size:11.5px;color:#0f8a5c;background:rgba(23,166,115,.1);padding:9px;border-radius:9px;font-weight:600">&#128274; 100&nbsp;%% local — aucune donnée ne quitte cet ordinateur</div></div>""" % (LOGO_URI, _su, _sp, _ck))

@app.route("/logout")
def logout(): session.clear(); return redirect(url_for("login"))

# ======================================================================
#  STATUT DE TRAITEMENT + TABLEAU DE BORD
# ======================================================================
# (clé, libellé, couleur, emoji) — badge « médaille » (double cercle) + emoji au centre
STATUTS = [("bilan", "Bilan", "#6366f1", "\U0001F50E"),
           ("en_cours", "En cours", "#16a34a", "\U0001F62C"),
           ("surveillance", "Surveillance", "#d97706", "\U0001F440"),
           ("termine", "Terminé", "#64748b", "\U0001F389"),
           ("perdu", "Perdu de vue", "#b91c1c", "\U0001F441")]
STATUT_MAP = {k: (lbl, color, emoji) for k, lbl, color, emoji in STATUTS}

def statut_of(pt):
    s = pt.get("statut")
    return s if s in STATUT_MAP else "en_cours"

def _medal(color, size, emoji=None):
    """Badge « médaille » : double cercle coloré (cercle plein + anneau blanc),
    sans ruban ni icône. `emoji` est ignoré (conservé pour compat des appels)."""
    c = size / 2.0; outer = size * 0.42; inner = size * 0.26
    sw = max(1.6, size * 0.09)
    _sz = int(round(size))
    return ('<span style="display:inline-block;flex:none;vertical-align:middle;width:%dpx;height:%dpx;'
            'border-radius:50%%;background:radial-gradient(circle,%s 51%%,#fff 51%% 71%%,%s 71%%)"></span>'
            ) % (_sz, _sz, color, color)

def statut_badge(pt, small=False):
    s = statut_of(pt); lbl, color, emoji = STATUT_MAP[s]
    sz = 20 if small else 27
    fs = "12.5px" if small else "15px"
    return ('<span style="display:inline-flex;align-items:center;gap:8px;white-space:nowrap">'
            '%s<span style="font-weight:800;color:%s;font-size:%s">%s</span></span>'
            ) % (_medal(color, sz), color, fs, lbl)

def status_menu(slug, pt):
    """Badge cliquable -> petit menu pour changer le statut directement."""
    cur = statut_of(pt)
    items = ""
    for k, lbl, color, emoji in STATUTS:
        chk = ' <span style="margin-left:auto;color:%s">&#10003;</span>' % color if k == cur else ''
        items += ('<a class=mi href="%s" style="min-width:170px">%s <span style="color:%s;font-weight:700">%s</span>%s</a>'
                  ) % (url_for("set_statut", slug=slug, key=k), _medal(color, 20, emoji), color, lbl, chk)
    return ('<details class=menu><summary style="cursor:pointer" title="Changer le statut">%s</summary>'
            '<div class="menu-pop" style="left:0;right:auto">%s</div></details>'
            ) % (statut_badge(pt, small=True), items)

def actions_menu(slug, pt=None):
    staff = ""
    if pt is not None:
        _on = bool(pt.get("a_staffer"))
        _lbl = "\u2605 Retirer du staff" if _on else "\u2606 Marquer \u00e0 staffer"
        _col = "#d97706" if _on else "#24405f"
        staff = ('<a class=mi href="%s" style="color:%s">%s</a><div class=sep></div>'
                 % (url_for("toggle_staffer", slug=slug), _col, _lbl))
    return ('<details class="menu" style="position:absolute;top:8px;right:8px;z-index:25"><summary class=dots>&#8942;</summary>'
            '<div class="menu-pop">'
            '%s'
            '<a class=mi href="%s">&#9998; Modifier la fiche</a>'
            '<div class=sep></div>'
            '<a class=mi href="%s" style="color:#C0392B">&#128465; Supprimer</a>'
            '</div></details>') % (staff, url_for("patient_edit", slug=slug), url_for("patient_delete", slug=slug))

# --- Boutons plateformes (Doctolib / WebCeph / XERO) : ouvrent la plateforme et cherchent le patient ---
# which, libellé, couleur pastille, lettre (repli si pas de logo)
PLAT = [
    ("doctolib", "Doctolib", "#0596DE", "D"),
    ("webceph",  "WebCeph",  "#5A45C8", "W"),
    ("xero",     "XERO",     "#0E7C7B", "X"),
]
# Rempli avec les vrais logos (data URI) quand fournis ; sinon pastille couleur.
PLAT_LOGOS = {
    'doctolib': 'data:image/png;base64,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',
    'webceph': 'data:image/png;base64,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',
    'xero': 'data:image/png;base64,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',
}

def _attr_esc(s):
    return (str(s or "").replace("&", "&amp;").replace('"', "&quot;")
            .replace("<", "&lt;").replace(">", "&gt;"))

def plat_btns(nom, size=26, center=False):
    """Rangée de 3 boutons plateformes pour un patient (nom prérempli à la recherche)."""
    nomq = _attr_esc(nom)
    just = "center" if center else "flex-start"
    out = ['<div class="platrow" onclick="event.stopPropagation()" '
           'style="display:inline-flex;gap:6px;align-items:center;justify-content:%s">' % just]
    for which, title, color, letter in PLAT:
        logo = PLAT_LOGOS.get(which)
        if logo:
            inner = ('<img src="%s" alt="%s" style="width:82%%;height:82%%;object-fit:contain">' % (logo, title))
            bg = "#fff"
        else:
            inner = ('<span style="color:#fff;font-weight:800;font-size:%dpx;'
                     'font-family:system-ui,-apple-system,sans-serif;line-height:1">%s</span>'
                     % (max(11, int(size * 0.5)), letter))
            bg = color
        out.append(
            '<button type="button" class="platbtn" title="Ouvrir %s (recherche : %s)" data-nom="%s" '
            'onclick="return openPlat(&#39;%s&#39;,this,event)" '
            'style="width:%dpx;height:%dpx;border-radius:8px;border:1px solid rgba(0,0,0,.10);'
            'background:%s;display:inline-flex;align-items:center;justify-content:center;cursor:pointer;'
            'padding:0;box-shadow:0 1px 3px rgba(0,0,0,.14);transition:.12s" '
            'onmouseover="this.style.transform=&#39;translateY(-1px)&#39;" '
            'onmouseout="this.style.transform=&#39;&#39;">%s</button>'
            % (title, nomq, nomq, which, size, size, bg, inner))
    out.append('</div>')
    return "".join(out)

def first_record_date(slug, pt):
    ds = []
    for rid in pt.get("records", []):
        r = load_rec(slug, rid)
        if r and r.get("date"):
            ds.append((r["date"] or "")[:10])
    return min(ds) if ds else None

def treatment_duration_txt(slug, pt):
    d = first_record_date(slug, pt)
    if not d:
        return ""
    try:
        b = datetime.date.fromisoformat(d); t = datetime.date.today()
        m = max(0, (t.year - b.year) * 12 + (t.month - b.month) - (1 if t.day < b.day else 0))
        if m < 1: return "depuis moins d'un mois"
        if m < 12: return "%d mois de traitement" % m
        a, mo = divmod(m, 12)
        return ("%d an%s" % (a, "s" if a > 1 else "")) + ((" %d mois" % mo) if mo else "") + " de traitement"
    except Exception:
        return ""

# ======================================================================
#  TRI DE LA BIBLIOTHÈQUE
# ======================================================================
SORTS = [("ouvert", "Derniers ouverts"), ("recent", "Ajout récent"), ("nom", "Nom (A → Z)"),
         ("age", "Âge (jeune → âgé)"), ("debut", "Début de traitement (récent)"),
         ("statut", "Statut du traitement")]
SORT_KEYS = {k for k, _ in SORTS}
_STATUT_ORDER = {k: i for i, (k, _l, _c, _e) in enumerate(STATUTS)}

def _age_num(m):
    """Âge numérique du patient (depuis 'age' type '14 ans', repli sur la date de naissance)."""
    try:
        return int(re.match(r"\s*(\d+)", m.get("age", "") or "").group(1))
    except Exception:
        pass
    try:
        d = datetime.date.fromisoformat((m.get("dob", "") or "")[:10]); t = datetime.date.today()
        return t.year - d.year - ((t.month, t.day) < (d.month, d.day))
    except Exception:
        return None

def _touch_opened(slug):
    try:
        st = get_settings(); o = dict(st.get("opened", {}))
        o[slug] = datetime.datetime.now().isoformat(timespec="seconds")
        st["opened"] = o; save_settings(st)
    except Exception:
        pass

def sort_patients(pts, tri):
    nom = lambda m: (m.get("nom", "") or "").lower()
    if tri == "ouvert":
        om = get_settings().get("opened", {})
        return sorted(pts, key=lambda m: (om.get(m["slug"]) or m.get("date_creation", "")), reverse=True)
    if tri == "nom":
        return sorted(pts, key=nom)
    if tri == "age":
        return sorted(pts, key=lambda m: (_age_num(m) is None, _age_num(m) or 0, nom(m)))
    if tri == "debut":   # début de traitement le plus récent d'abord ; sans date -> à la fin
        return sorted(pts, key=lambda m: (first_record_date(m["slug"], m) or ""), reverse=True)
    if tri == "statut":  # ordre du parcours : Bilan -> En cours -> Surveillance -> Terminé
        return sorted(pts, key=lambda m: (_STATUT_ORDER.get(statut_of(m), 99), nom(m)))
    return sorted(pts, key=lambda m: m.get("date_creation", ""), reverse=True)  # récent (défaut)

# ======================================================================
#  DASHBOARD
# ======================================================================
@app.route("/")
def dashboard():
    if not logged(): return redirect(url_for("login"))
    auto_backup_if_due()
    if "crop_calib" not in get_settings():   # 1re fois : apprend sur tes photos déjà recadrées
        auto_crop_calib()
    q = request.args.get("q", "").strip().lower()
    fstat = request.args.get("statut", "").strip()
    fstaff = request.args.get("staffer", "").strip()
    fpres = request.args.get("presenter", "").strip()
    _sett = get_settings()
    tri = request.args.get("tri")
    if tri is None: tri = _sett.get("lib_tri", "ouvert")
    tri = (tri or "").strip()
    if tri not in SORT_KEYS: tri = "ouvert"
    vue = request.args.get("vue")
    if vue is None: vue = _sett.get("lib_vue", "photo")
    vue = (vue or "").strip()
    if vue not in ("photo", "icone", "liste"): vue = "photo"
    if _sett.get("lib_tri") != tri or _sett.get("lib_vue") != vue:
        _sett["lib_tri"] = tri; _sett["lib_vue"] = vue; save_settings(_sett)
    allpts = list_patients()
    counts = {k: 0 for k, _l, _c, _e in STATUTS}
    for m in allpts:
        counts[statut_of(m)] = counts.get(statut_of(m), 0) + 1
    nstaff = sum(1 for m in allpts if m.get("a_staffer"))
    npres = sum(1 for m in allpts if m.get("a_presenter"))
    pts = allpts
    if q: pts = [m for m in pts if q in m.get("nom", "").lower()]
    if fstat in STATUT_MAP: pts = [m for m in pts if statut_of(m) == fstat]
    if fstaff: pts = [m for m in pts if m.get("a_staffer")]
    if fpres: pts = [m for m in pts if m.get("a_presenter")]
    pts = sort_patients(pts, tri)
    # Les "perdus de vue" descendent tout en bas (tri stable), sauf si on filtre justement sur ce statut.
    if fstat != "perdu":
        pts = sorted(pts, key=lambda m: 1 if statut_of(m) == "perdu" else 0)
    def _latest_photo_iso(slug, recs):
        best = ""
        for rid in recs:
            base = rdir(slug, rid); has = False
            for sub in ("02_photos_traitees", "01_photos_brutes"):
                d = os.path.join(base, sub)
                try:
                    if os.path.isdir(d) and any(f.lower().endswith(".jpg") for f in os.listdir(d)):
                        has = True; break
                except Exception: pass
            if has:
                dt = ((load_rec(slug, rid) or {}).get("date", "") or "")[:10]
                if dt and dt > best: best = dt
        return best
    def _older_than_months(iso, months=3):
        try: d = datetime.date.fromisoformat((iso or "")[:10])
        except Exception: return False
        import calendar
        mm = d.month - 1 + months; y = d.year + mm // 12; mo = mm % 12 + 1
        day = min(d.day, calendar.monthrange(y, mo)[1])
        return datetime.date(y, mo, day) < datetime.date.today()
    cards = []
    for m in pts:
        slug = m["slug"]; recs = m.get("records", [])
        thumb = ""
        # photo de FACE du BILAN INITIAL (1er enregistrement) en priorité : repos -> sourire -> profil.
        # Replis : autre photo du bilan initial, puis mêmes essais sur les enregistrements suivants.
        face = ["exo_face_repos", "exo_face_sourire", "exo_profil"]
        order = recs[:1] + recs[1:]            # initial d'abord, puis les suivants
        for pref in (face, [k for k, _ in PHOTO_FIELDS]):   # d'abord les vues de face, puis n'importe quelle photo
            for rid in order:
                for key in pref:
                    rel = "02_photos_traitees/%s.jpg" % key
                    if os.path.exists(os.path.join(rdir(slug, rid), rel)):
                        thumb = "data-u=\"%s\"" % url_for("rthumb", slug=slug, rid=rid, path=rel)
                        break
                if thumb: break
            if thumb: break
        _su = url_for("suivi", slug=slug); _nom = m.get("nom", ""); _age = m.get("age", "") or ""
        _plat = plat_btns(_nom)
        _staff = bool(m.get("a_staffer"))
        _present = bool(m.get("a_presenter"))
        _stale = _older_than_months(_latest_photo_iso(slug, recs), 3)
        # Étoile "à staffer" CLIQUABLE (bascule le statut) — pleine et dorée si à staffer, sinon contour gris.
        def _starbtn(size=19):
            return ('<a href="%s" onclick="event.stopPropagation()" title="%s" '
                    'style="text-decoration:none;font-size:%dpx;line-height:1;cursor:pointer;color:%s">%s</a>'
                    ) % (url_for("toggle_staffer", slug=slug),
                         "Retirer du staff" if _staff else "Marquer à staffer",
                         size, "#f59e0b" if _staff else "#c3ccd8",
                         "★" if _staff else "☆")
        # "à présenter" CLIQUABLE — écran de présentation (indigo si actif, gris sinon).
        def _presbtn(size=19):
            col = "#6366f1" if _present else "#c3ccd8"
            return ('<a href="%s" onclick="event.stopPropagation()" title="%s" '
                    'style="display:inline-flex;text-decoration:none;line-height:1;cursor:pointer;color:%s">'
                    '<svg viewBox="0 0 24 24" width="%d" height="%d" fill="none" stroke="currentColor" '
                    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
                    '<path d="M3 4h18"/><path d="M4 4v10a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1V4"/>'
                    '<path d="M12 15v4"/><path d="M9 21l3-2 3 2"/><path d="M8 11l2.5-2.5 2 2L16 7"/></svg></a>'
                    ) % (url_for("toggle_presenter", slug=slug),
                         "Retirer « à présenter »" if _present else "Marquer « à présenter »",
                         col, size, size)
        # Appareil photo si les photos datent de plus de 3 mois (auto, non cliquable).
        def _cam(inline=False):
            if not _stale: return ""
            svg = ('<svg viewBox="0 0 24 24" width="%d" height="%d" fill="none" stroke="currentColor" '
                   'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
                   '<path d="M23 19a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h4l2-3h6l2 3h4a2 2 0 0 1 2 2z"/>'
                   '<circle cx="12" cy="13" r="3.5"/></svg>')
            if inline:
                return ('<span title="Photos de plus de 3 mois — à refaire" style="display:inline-flex;'
                        'align-items:center;color:#e8590c">%s</span>') % (svg % (18, 18))
            return ('<div title="Photos de plus de 3 mois — à refaire" style="position:absolute;top:8px;left:8px;'
                    'z-index:22;display:inline-flex;align-items:center;gap:4px;background:#e8590c;color:#fff;'
                    'padding:4px 9px;border-radius:999px;box-shadow:0 2px 6px rgba(0,0,0,.25);font-size:11px;'
                    'font-weight:800">%s+3 mois</div>') % (svg % (14, 14))
        _flags = ('<span style="display:inline-flex;gap:9px;align-items:center">%s%s%s</span>'
                  % (_starbtn(), _presbtn(), _cam(inline=True)))
        if vue == "liste":
            cards.append(('<div class=lrow>%s<span style="display:inline-flex;gap:7px;align-items:center;margin-right:2px">%s%s%s</span>'
                          '<a href="%s" class=lnm>%s</a>'
                          '<span class=lage>%s</span>%s<span class=lstat>%s</span></div>')
                         % (actions_menu(slug, m), _starbtn(16), _presbtn(16), _cam(inline=True),
                            _su, _nom, _age, plat_btns(_nom, size=22), status_menu(slug, m)))
        elif vue == "icone":
            cards.append(('<div class="pcard nopic">%s'
                          '<a href="%s" style="display:block;color:inherit;text-decoration:none">'
                          '<div class=body><div class=tic>🦷</div>'
                          '<div class=nm>%s</div><div class=mt style="margin-top:2px">%s</div></div></a>'
                          '<div style="padding:0 15px 14px;text-align:center">%s'
                          '<div style="margin-top:8px;display:flex;justify-content:center">%s</div>'
                          '<div style="margin-top:8px;display:flex;justify-content:center">%s</div></div></div>')
                         % (actions_menu(slug, m), _su, _nom, _age, status_menu(slug, m), _flags, _plat))
        else:
            cards.append("""<div class=pcard>%s%s
        <a href="%s" style="display:block;color:inherit;text-decoration:none"><div class=thumb %s>%s</div></a>
        <div class=body>
          <a href="%s" class=nm style="display:block;color:inherit;text-decoration:none">%s</a>
          <div class=mt style="margin-top:2px">%s</div>
          <div style="margin-top:10px;display:flex;align-items:center;justify-content:space-between;gap:8px">%s%s</div>
          <div style="margin-top:8px">%s</div>
        </div></div>""" % (
            actions_menu(slug, m), _cam(inline=False), _su, thumb, "" if thumb else "🦷",
            _su, _nom, _age, status_menu(slug, m), _flags, _plat))
    if not cards:
        cards = ["<div class='card center muted'>Aucun patient%s. <a href='%s'>Créer le premier</a>.</div>"
                 % ((" pour ce filtre" if (q or fstat or fstaff or fpres) else ""), url_for("nouveau"))]
    # barre de filtres par statut
    def fchip(key, lbl, n, active, color="#2F5CA8"):
        base = ("display:inline-flex;align-items:center;gap:7px;padding:7px 15px;border-radius:999px;"
                "text-decoration:none;font-size:13.5px;margin:0 8px 9px 0;font-weight:700;border:1.5px solid ")
        if active:
            stl = base + "%s;background:%s;color:#fff;box-shadow:0 2px 6px rgba(0,0,0,.15)" % (color, color)
            dot = '<span style="width:8px;height:8px;border-radius:50%%;background:rgba(255,255,255,.9)"></span>'
        else:
            stl = base + "%s;background:#fff;color:#24405f" % (color + "55")
            dot = '<span style="width:8px;height:8px;border-radius:50%%;background:%s"></span>' % color
        args = {}
        if key: args["statut"] = key
        if tri != "ouvert": args["tri"] = tri
        if vue != "photo": args["vue"] = vue
        if fstaff: args["staffer"] = "1"
        if fpres: args["presenter"] = "1"
        href = url_for("dashboard", **args)
        cnt = '<span style="opacity:.75;font-weight:600">%d</span>' % n
        return '<a href="%s" style="%s">%s%s %s</a>' % (href, stl, (dot if key else ""), lbl, cnt)
    total = len(allpts)
    chips = fchip("", "Tous", total, not fstat)
    for k, lbl, _c, _e in STATUTS:
        chips += fchip(k, lbl, counts.get(k, 0), fstat == k, _c)
    _scb = ("display:inline-flex;align-items:center;gap:7px;padding:7px 15px;border-radius:999px;"
            "text-decoration:none;font-size:13.5px;margin:0 8px 9px 14px;font-weight:700;border:1.5px solid ")
    # --- Filtre "\u00c0 staffer" ---
    _stargs = {}
    if fstat in STATUT_MAP: _stargs["statut"] = fstat
    if tri != "ouvert": _stargs["tri"] = tri
    if vue != "photo": _stargs["vue"] = vue
    if fpres: _stargs["presenter"] = "1"
    if not fstaff: _stargs["staffer"] = "1"
    if fstaff:
        _scs = _scb + "#d97706;background:#d97706;color:#fff;box-shadow:0 2px 6px rgba(0,0,0,.15)"
    else:
        _scs = _scb + "#d9770655;background:#fff;color:#8a5a00"
    chips += '<a href="%s" style="%s">\u2605 \u00c0 staffer <span style="opacity:.75;font-weight:600">%d</span></a>' % (url_for("dashboard", **_stargs), _scs, nstaff)
    # --- Filtre "\u00c0 pr\u00e9senter" ---
    _prargs = {}
    if fstat in STATUT_MAP: _prargs["statut"] = fstat
    if tri != "ouvert": _prargs["tri"] = tri
    if vue != "photo": _prargs["vue"] = vue
    if fstaff: _prargs["staffer"] = "1"
    if not fpres: _prargs["presenter"] = "1"
    _prcb = ("display:inline-flex;align-items:center;gap:7px;padding:7px 15px;border-radius:999px;"
             "text-decoration:none;font-size:13.5px;margin:0 8px 9px 0;font-weight:700;border:1.5px solid ")
    if fpres:
        _prs = _prcb + "#6366f1;background:#6366f1;color:#fff;box-shadow:0 2px 6px rgba(0,0,0,.15)"
    else:
        _prs = _prcb + "#6366f155;background:#fff;color:#4338ca"
    _prico = ('<svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" '
              'stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="vertical-align:-2px">'
              '<path d="M3 4h18"/><path d="M4 4v10a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1V4"/>'
              '<path d="M12 15v4"/><path d="M9 21l3-2 3 2"/><path d="M8 11l2.5-2.5 2 2L16 7"/></svg>')
    chips += '<a href="%s" style="%s">%s \u00c0 pr\u00e9senter <span style="opacity:.75;font-weight:600">%d</span></a>' % (url_for("dashboard", **_prargs), _prs, _prico, npres)
    triopts = "".join('<option value="%s"%s>%s</option>' % (k, " selected" if k == tri else "", lbl)
                      for k, lbl in SORTS)
    trisel = ('<label style="display:inline-flex;align-items:center;gap:8px;font-size:13.5px;color:var(--mut)">'
              'Trier&nbsp;: <select name=tri onchange="this.form.submit()" '
              'style="width:auto;padding:7px 10px;border-radius:9px">%s</select></label>') % triopts
    _vopts = "".join('<option value="%s"%s>%s</option>' % (k, " selected" if k == vue else "", lbl)
                     for k, lbl in (("photo", "Icônes (photo)"), ("icone", "Icônes (sans photo)"), ("liste", "Liste")))
    vuesel = ('<label style="display:inline-flex;align-items:center;gap:8px;font-size:13.5px;color:var(--mut)">'
              'Affichage&nbsp;: <select name=vue onchange="this.form.submit()" '
              'style="width:auto;padding:7px 10px;border-radius:9px">%s</select></label>') % _vopts
    _scroll_js = ("<script>(function(){var K='libScroll';"
      "try{if(sessionStorage.getItem('libBack')){var y=parseInt(sessionStorage.getItem(K)||'0',10)||0;"
      "window.scrollTo(0,y);sessionStorage.removeItem('libBack');}}catch(e){}"
      "document.addEventListener('click',function(ev){var a=ev.target.closest&&ev.target.closest('a[href]');"
      "if(a&&a.getAttribute('href')&&a.getAttribute('href').indexOf('/patient/')===0){"
      "try{sessionStorage.setItem(K,String(window.scrollY));sessionStorage.setItem('libBack','1');}catch(e){}}},true);"
      "})();</script>"
      "<script>(function(){try{var io=new IntersectionObserver(function(es){es.forEach(function(e){"
      "if(e.isIntersecting){var t=e.target,u=t.getAttribute('data-u');if(u){t.style.setProperty('--u',\"url('\"+u+\"')\");}io.unobserve(t);}});},{rootMargin:'500px'});"
      "document.querySelectorAll('.thumb[data-u]').forEach(function(t){io.observe(t);});"
      "}catch(e){document.querySelectorAll('.thumb[data-u]').forEach(function(t){t.style.setProperty('--u',\"url('\"+t.getAttribute('data-u')+\"')\");});}})();</script>")
    return page("""<div class=phead><h1>Bibliothèque</h1><div class=sub>%d patient%s</div></div>
    <form method=get style="display:flex;gap:12px;margin-bottom:14px;flex-wrap:wrap;align-items:center">
    <input type=hidden name=q value="%s"><input type=hidden name=statut value="%s">
    <input type=hidden name=staffer value="%s"><input type=hidden name=presenter value="%s">%s%s</form>
    <div style="margin-bottom:18px">%s</div>
    <div class="plist %s">%s</div>%s""" % (total, "s" if total > 1 else "",
        q, (fstat if fstat in STATUT_MAP else ""),
        ("1" if fstaff else ""), ("1" if fpres else ""),
        trisel, vuesel, chips, vue, "".join(cards), _scroll_js))

# ======================================================================
#  NOUVEAU PATIENT (+ bilan initial)
# ======================================================================
@app.route("/nouveau")
def nouveau():
    if not logged(): return redirect(url_for("login"))
    return page("""<h1>Nouveau patient</h1>
    <div class=flash style="background:var(--accw);border-color:var(--line)">💡 <b>Import intelligent</b> : renseigne juste l'identité ci-dessous, clique « <b>Créer puis importer en vrac</b> », et dépose tous les fichiers d'un coup — l'app devine photo / radio / STL et la vue de chaque photo (tu confirmes en un écran).</div>
    <form method=post action="%s" enctype=multipart/form-data>
    %s<div style="position:sticky;bottom:0;z-index:20;display:flex;gap:10px;flex-wrap:wrap;align-items:center;background:var(--card);padding:12px;margin-top:16px;border:1px solid var(--line);border-radius:12px;box-shadow:0 -6px 20px rgba(0,0,0,.10)">
    <button type=submit formaction="%s" class=btn onclick="this.innerHTML='<span class=spin></span> Création…'">&#9889; Créer puis importer en vrac (IA)</button>
    <button type=submit class="btn sec" onclick="this.innerHTML='<span class=spin></span> Génération…'">Créer avec les champs ci-dessus</button></div></form>""" % (
        url_for("creer"), form_fields(), url_for("creer_import")))

@app.route("/from_doctolib")
def from_doctolib():
    if not logged(): return redirect(url_for("login"))
    import json as _j
    fp = os.path.join(DATA, "_doctolib_inbox.json")
    if not os.path.exists(fp):
        flash("Aucune donnée Doctolib en attente."); return redirect(url_for("dashboard"))
    try: o = _j.load(open(fp, encoding="utf-8"))
    except Exception:
        flash("Lecture Doctolib impossible."); return redirect(url_for("dashboard"))
    try: os.remove(fp)
    except Exception: pass
    nom_raw = (o.get("nom", "") or "").strip(); prenom = (o.get("prenom", "") or "").strip()
    dob = (o.get("dob", "") or "").strip()
    if not nom_raw:
        flash("Identité Doctolib illisible — fiche non créée."); return redirect(url_for("dashboard"))
    nom = (nom_raw.upper() + " " + prenom).strip()
    slug = slugify(nom) + "_" + datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    dob_court, age = age_from_dob(dob) if dob else ("—", "—")
    pt = {"slug": slug, "nom": nom, "dob": dob, "dob_court": dob_court, "age": age,
          "sexe": (o.get("sexe") or "non précisé"), "auteur": session.get("user", ""),
          "tel": (o.get("tel", "") or ""), "email": (o.get("email", "") or ""), "source": "doctolib",
          "date_creation": datetime.datetime.now().isoformat(timespec="seconds"), "statut": "bilan", "records": []}
    os.makedirs(pdir(slug), exist_ok=True); save_patient(slug, pt)
    rid = create_empty_record(slug, "Bilan initial")
    flash("Fiche créée depuis Doctolib : %s." % nom)
    return redirect(url_for("record", slug=slug, rid=rid))


@app.route("/creer", methods=["POST"])
def creer():
    if not logged(): return redirect(url_for("login"))
    nom = (request.form.get("nom", "").strip().upper() + " " + request.form.get("prenom", "").strip()).strip()
    slug = slugify(nom) + "_" + datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    dob = request.form.get("dob", ""); dob_court, age = age_from_dob(dob) if dob else ("—", "—")
    pt = {"slug": slug, "nom": nom, "dob": dob, "dob_court": dob_court, "age": age,
          "sexe": request.form.get("sexe", "non précisé"), "auteur": session["user"],
          "date_creation": datetime.datetime.now().isoformat(timespec="seconds"), "statut": "bilan", "records": []}
    os.makedirs(pdir(slug), exist_ok=True); save_patient(slug, pt)
    rid = _new_record(slug, "Bilan initial", request)
    return redirect(url_for("record", slug=slug, rid=rid))

def create_empty_record(slug, label):
    """Crée un enregistrement vide (dossiers + meta), sans fichiers ni génération."""
    pt = load_patient(slug)
    rid = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    base = rdir(slug, rid); _k = 0
    while os.path.exists(base):        # évite les collisions quand on crée plusieurs temps dans la même seconde
        _k += 1; rid = datetime.datetime.now().strftime("%Y%m%d%H%M%S") + "_%d" % _k; base = rdir(slug, rid)
    for sub in ["01_photos_brutes", "02_photos_traitees", "03_radios", "04_stl_bruts", "05_stl_rendus", "06_webceph"]:
        os.makedirs(os.path.join(base, sub), exist_ok=True)
    rec = {"rid": rid, "label": label, "date": datetime.datetime.now().isoformat(timespec="seconds"),
           "steiner": {}, "synthese": {}, "exam": {}, "angle": {}, "motif": "", "rotations": {}, "crops": {}, "status": "empty"}
    save_rec(slug, rid, rec)
    pt["records"] = pt.get("records", []) + [rid]; save_patient(slug, pt)
    return rid

@app.route("/creer_import", methods=["POST"])
def creer_import():
    if not logged(): return redirect(url_for("login"))
    nom = (request.form.get("nom", "").strip().upper() + " " + request.form.get("prenom", "").strip()).strip()
    slug = slugify(nom) + "_" + datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    dob = request.form.get("dob", ""); dob_court, age = age_from_dob(dob) if dob else ("—", "—")
    pt = {"slug": slug, "nom": nom, "dob": dob, "dob_court": dob_court, "age": age,
          "sexe": request.form.get("sexe", "non précisé"), "auteur": session["user"],
          "date_creation": datetime.datetime.now().isoformat(timespec="seconds"), "statut": "bilan", "records": []}
    os.makedirs(pdir(slug), exist_ok=True); save_patient(slug, pt)
    rid = create_empty_record(slug, "Bilan initial")
    return redirect(url_for("import_upload", slug=slug, rid=rid))

@app.route("/patient/<slug>/reeval_import", methods=["POST"])
def reeval_import(slug):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug)
    if not pt: abort(404)
    nre = sum(1 for rid in pt.get("records", []) if str((load_rec(slug, rid) or {}).get("label", "")).lower().startswith("réév"))
    rid = create_empty_record(slug, "Réévaluation %d" % (nre + 1))
    return redirect(url_for("import_upload", slug=slug, rid=rid))

# ---- Import intelligent : dépôt en vrac -> classification -> vérification -> rangement ----
IMPORT_TARGETS = ([("photo:%s" % k, "Photo · %s" % l) for k, l in PHOTO_FIELDS]
                  + [("radio:panoramique", "Radio · Panoramique"), ("radio:teleradiographie_profil", "Radio · Téléradiographie de profil")]
                  + [("stl:UpperJawScan", "Modèle STL · Arcade supérieure"), ("stl:LowerJawScan", "Modèle STL · Arcade inférieure")])

def _import_options(sel):
    opts = '<option value="">— Ignorer ce fichier —</option>'
    return opts + "".join('<option value="%s" %s>%s</option>' % (v, "selected" if v == sel else "", l) for v, l in IMPORT_TARGETS)

@app.route("/import/<slug>/<rid>", methods=["GET", "POST"])
def import_upload(slug, rid):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r: abort(404)
    base = rdir(slug, rid); stage = os.path.join(base, "_import_stage")
    if request.method == "POST":
        if os.path.isdir(stage): shutil.rmtree(stage, ignore_errors=True)
        os.makedirs(stage, exist_ok=True)
        manifest = []
        files = request.files.getlist("files")
        for i, f in enumerate(files):
            if not f or not f.filename: continue
            ext = os.path.splitext(f.filename)[1].lower() or ".dat"
            sn = "%03d%s" % (i, ext)
            f.save(os.path.join(stage, sn))
            if ext == ".docx":
                guess = "word:bilan"      # bilan Word -> lecture des valeurs
            else:
                guess = P.guess_category(os.path.join(stage, sn), f.filename)
            manifest.append({"file": sn, "orig": f.filename, "guess": guess})
        json.dump(manifest, open(os.path.join(stage, "manifest.json"), "w"), ensure_ascii=False)
        return redirect(url_for("import_review", slug=slug, rid=rid))
    # GET : formulaire de dépôt
    body = """<p class=muted><a href="%s">← Fiche patient</a></p>
    <h1>Import intelligent — %s</h1><h2>%s</h2>
    <div class=card><p>Dépose <b>tous les fichiers d'un coup</b> (photos, radios, modèles STL, <b>et même un bilan Word .docx</b>).
    L'app devine le type et la vue de chacun, lit les valeurs du Word, puis te montre un écran de vérification.</p>
    <p class=muted>Tu peux importer <b>en plusieurs fois</b> (fichiers dans des dossiers différents) : valide un lot, puis reviens
    en ajouter d'autres — les fichiers déjà rangés sont conservés.</p>
    <form method=post enctype=multipart/form-data onsubmit="this.querySelector('button').innerHTML='<span class=spin></span> Analyse…'">
      <input type=file name=files multiple style="padding:20px;border:2px dashed #b9c6de;background:#f7f9fd">
      <p class=muted style="margin:10px 0 0">Photos (JPG/PNG), radios, modèles <b>STL</b> et bilan <b>Word (.docx)</b> — tu peux tout sélectionner d'un coup.</p>
      <div style="margin-top:14px"><button class=btn>Analyser les fichiers</button></div>
    </form></div>""" % (url_for("patient", slug=slug), pt["nom"], r.get("label", ""))
    return page(body)

@app.route("/import/<slug>/<rid>/review")
def import_review(slug, rid):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r: abort(404)
    stage = os.path.join(rdir(slug, rid), "_import_stage")
    mpath = os.path.join(stage, "manifest.json")
    if not os.path.exists(mpath):
        return redirect(url_for("import_upload", slug=slug, rid=rid))
    manifest = json.load(open(mpath))
    import base64, io as _io
    from PIL import Image as _Image, ImageOps as _IO
    def thumb(sn):
        p = os.path.join(stage, sn)
        if sn.lower().endswith((".stl", ".ply", ".obj")):
            return ""
        try:
            im = _IO.exif_transpose(_Image.open(p)).convert("RGB"); im.thumbnail((190, 190))
            buf = _io.BytesIO(); im.save(buf, "JPEG", quality=80)
            return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()
        except Exception:
            return ""
    LOW_CONF = {"photo:exo_face_repos", "photo:exo_face_sourire", "photo:endo_laterale_droite", "photo:endo_laterale_gauche", ""}
    cards = ""
    for e in manifest:
        g = e.get("guess", "")
        if e["file"].lower().endswith(".docx") or g.startswith("word:"):
            # carte spéciale « bilan Word » : lecture des valeurs (pas de vue à choisir)
            vis = ('<div style="height:150px;display:flex;flex-direction:column;align-items:center;justify-content:center;'
                   'background:#eef7ee;border-radius:8px;color:#2e7d32;font-weight:700">&#128196;<div style="font-size:12px;margin-top:6px">Bilan Word</div></div>')
            opts = ('<option value="word:bilan" selected>Bilan Word — lire les valeurs</option>'
                    '<option value="">— Ignorer ce fichier —</option>')
            cards += ("""<div class="impcard" style="border:1px solid #cfe6cf;border-radius:10px;padding:10px">%s
            <div class=muted style="font-size:11px;margin:6px 0;word-break:break-all">%s</div>
            <select name="s_%s" class=impsel onchange="dupCheck()" style="font-size:13px">%s</select>
            <button type=button onclick="ignoreCard(this)" style="margin-top:6px;width:100%%;background:#fbecec;color:#C0392B;border:1px solid #f0c9c4;border-radius:7px;padding:5px;cursor:pointer;font-size:12px">&#128465; Retirer</button>
            </div>""") % (vis, e["orig"], e["file"], opts)
            continue
        t = thumb(e["file"])
        vis = ('<img src="%s" style="width:100%%;height:150px;object-fit:contain;background:#fafbfd;border-radius:8px;border:1px solid #e3e8ef">' % t
               if t else '<div style="height:150px;display:flex;align-items:center;justify-content:center;background:var(--card2);border-radius:8px;color:var(--acc2);font-weight:700">&#128190; STL / modèle 3D</div>')
        badge = ('<span style="background:#fff3d6;border:1px solid #e6c65b;color:#8a5a00;border-radius:20px;padding:1px 8px;font-size:11px;margin-left:6px">à vérifier</span>'
                 if g in LOW_CONF else '')
        cards += ("""<div class="impcard" style="border:1px solid var(--line);border-radius:10px;padding:10px">%s
        <div class=muted style="font-size:11px;margin:6px 0;word-break:break-all">%s%s</div>
        <select name="s_%s" class=impsel onchange="dupCheck()" style="font-size:13px">%s</select>
        <button type=button onclick="ignoreCard(this)" style="margin-top:6px;width:100%%;background:#fbecec;color:#C0392B;border:1px solid #f0c9c4;border-radius:7px;padding:5px;cursor:pointer;font-size:12px">&#128465; Retirer / doublon</button>
        </div>""") % (vis, e["orig"], badge, e["file"], _import_options(g))
    js = ("<script>"
          "function ignoreCard(b){var c=b.closest('.impcard');var s=c.querySelector('.impsel');s.value='';"
          "c.style.opacity=(c.style.opacity=='0.45'?'1':'0.45');dupCheck();}"
          "function dupCheck(){var sels=[].slice.call(document.querySelectorAll('.impsel'));var seen={};"
          "sels.forEach(function(s){s.style.borderColor='';});var dup=false;"
          "sels.forEach(function(s){if(!s.value)return;if(seen[s.value]){dup=true;s.style.borderColor='#C0392B';seen[s.value].style.borderColor='#C0392B';}else{seen[s.value]=s;}});"
          "var w=document.getElementById('dupwarn');if(w)w.style.display=dup?'block':'none';}"
          "window.addEventListener('DOMContentLoaded',dupCheck);</script>")
    body = """<p class=muted><a href="%s">← Recommencer le dépôt</a></p>
    <h1>Vérification du classement</h1><h2>%s · %s</h2>
    <div class=flash style="background:var(--accw);border-color:var(--line)">Le type (photo/radio/STL) est fiable ; la <b>vue</b> est une estimation.
    Les cases <b>« à vérifier »</b> sont les plus ambiguës (repos/sourire, droite/gauche). Si tu as <b>2 photos du même élément</b>,
    clique <b>« Retirer »</b> sur celle à écarter (ou mets-la sur « Ignorer »).</div>
    <div id=dupwarn class="flash err" style="display:none">&#9888; Deux fichiers pointent vers le même emplacement (surlignés en rouge) — un seul sera gardé. Retire le doublon.</div>
    <form method=post action="%s">
      <div class=plist style="grid-template-columns:repeat(auto-fill,minmax(200px,1fr))">%s</div>
      <div style="margin-top:16px;display:flex;gap:10px;flex-wrap:wrap">
        <button class=btn name=then value=done onclick="this.innerHTML='<span class=spin></span> Rangement…'">Valider et ranger</button>
        <button class="btn sec" name=then value=more onclick="this.innerHTML='<span class=spin></span> Rangement…'">Valider et importer d'autres fichiers</button>
        <a class="btn sec" href="%s">Annuler</a></div>
    </form>%s""" % (url_for("import_upload", slug=slug, rid=rid), pt["nom"], r.get("label", ""),
                    url_for("import_commit", slug=slug, rid=rid), cards, url_for("record", slug=slug, rid=rid), js)
    return page(body)

@app.route("/import/<slug>/<rid>/commit", methods=["POST"])
def import_commit(slug, rid):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r: abort(404)
    base = rdir(slug, rid); stage = os.path.join(base, "_import_stage")
    mpath = os.path.join(stage, "manifest.json")
    if not os.path.exists(mpath): return redirect(url_for("record", slug=slug, rid=rid))
    manifest = json.load(open(mpath))
    from PIL import Image as _Image, ImageOps as _IO
    r.setdefault("crops", {}); r.setdefault("rotations", {})
    placed = 0; word_read = 0
    for e in manifest:
        token = request.form.get("s_" + e["file"], "").strip()
        if not token or ":" not in token: continue
        cat, key = token.split(":", 1)
        src = os.path.join(stage, e["file"])
        try:
            if cat == "photo":
                _IO.exif_transpose(_Image.open(src)).convert("RGB").save(os.path.join(base, "01_photos_brutes", key + ".jpg"), quality=92)
                r["crops"].pop(key, None); r["rotations"].pop(key, None)   # nouveau cliché -> cadrage auto
                placed += 1
            elif cat == "radio":
                _IO.exif_transpose(_Image.open(src)).convert("RGB").save(os.path.join(base, "03_radios", key + ".jpg"), quality=92)
                placed += 1
            elif cat == "stl":
                shutil.copy(src, os.path.join(base, "04_stl_bruts", key + ".stl"))
                placed += 1
            elif cat == "word":
                res = WI.parse_word_bilan(src)
                apply_word_to_record(r, res)
                word_read += 1
        except Exception:
            pass
    shutil.rmtree(stage, ignore_errors=True)
    if r.get("status") == "empty": r["status"] = "processing"
    save_rec(slug, rid, r)
    _regenerate(slug, rid, pt, r)
    wtxt = (" + valeurs d'un bilan Word lues" if word_read else "")
    if request.form.get("then") == "more":
        flash("%d fichier(s) rangé(s)%s. Dépose le lot suivant — les fichiers déjà rangés sont conservés." % (placed, wtxt))
        return redirect(url_for("import_upload", slug=slug, rid=rid))
    flash("%d fichier(s) rangé(s)%s et bilan régénéré." % (placed, wtxt))
    return redirect(url_for("record", slug=slug, rid=rid))

def _new_record(slug, label, req):
    pt = load_patient(slug)
    rid = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    base = rdir(slug, rid)
    for sub in ["01_photos_brutes", "02_photos_traitees", "03_radios", "04_stl_bruts", "05_stl_rendus", "06_webceph"]:
        os.makedirs(os.path.join(base, sub), exist_ok=True)
    save_uploads(base, req.files)
    steiner = parse_steiner_form(req.form)
    synth = parse_synthese(req.form)
    for k, v in auto_synthese(steiner).items():   # pré-remplissage squelettique si non saisi
        synth.setdefault(k, v)
    rec = {"rid": rid, "label": label, "date": datetime.datetime.now().isoformat(timespec="seconds"),
           "steiner": steiner, "synthese": synth, "exam": parse_exam(req.form), "angle": parse_angle(req.form),
           "motif": req.form.get("motif", "").strip(), "rotations": {}, "status": "processing"}
    save_rec(slug, rid, rec)
    docx, pdf, log, err = run_generation(base, build_info(pt, rec), steiner, rec["rotations"])
    rec.update({"docx": os.path.basename(docx) if docx else None, "pdf": os.path.basename(pdf) if pdf else None,
                "log": log, "status": "error" if err else "done", "error": err})
    save_rec(slug, rid, rec)
    pt["records"] = pt.get("records", []) + [rid]; save_patient(slug, pt)
    return rid

@app.route("/patient/<slug>/reeval")
def reeval(slug):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug)
    if not pt: abort(404)
    return page("""<h1>Réévaluation — %s</h1><h2>Nouvel enregistrement daté (photos, radios, STL, Steiner)</h2>
    <div class=flash style="background:var(--accw);border-color:var(--line)">💡 Tu peux aussi « <b>Ajouter puis importer en vrac</b> » : dépose tous les fichiers d'un coup, l'app les trie.</div>
    <form method=post action="%s" enctype=multipart/form-data>
    %s<div style="position:sticky;bottom:0;z-index:20;display:flex;gap:10px;flex-wrap:wrap;align-items:center;background:var(--card);padding:12px;margin-top:16px;border:1px solid var(--line);border-radius:12px;box-shadow:0 -6px 20px rgba(0,0,0,.10)">
    <button type=submit formaction="%s" class=btn onclick="this.innerHTML='<span class=spin></span> Création…'">&#9889; Ajouter puis importer en vrac (IA)</button>
    <button type=submit class="btn sec" onclick="this.innerHTML='<span class=spin></span> Génération…'">Ajouter avec les champs ci-dessus</button></div></form>
    <p><a href="%s">← Retour à la fiche</a></p>""" % (
        pt["nom"], url_for("reeval_post", slug=slug), form_fields(with_identity=False),
        url_for("reeval_import", slug=slug), url_for("patient", slug=slug)))

@app.route("/patient/<slug>/reeval", methods=["POST"])
def reeval_post(slug):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug)
    if not pt: abort(404)
    nre = sum(1 for rid in pt.get("records", []) if str((load_rec(slug, rid) or {}).get("label", "")).lower().startswith("réév"))
    rid = _new_record(slug, "Réévaluation %d" % (nre + 1), request)
    return redirect(url_for("record", slug=slug, rid=rid))

# ======================================================================
#  FICHE PATIENT (timeline des enregistrements)
# ======================================================================
def patient_tabs(slug, active):
    tabs = [("suivi", "Suivi", url_for("suivi", slug=slug)),
            ("bilans", "Bilans / R\xe9\xe9valuations", url_for("patient", slug=slug)),
            ("evolution", "\xc9volution", url_for("evolution", slug=slug)),
            ("staff", "Staff", url_for("staff_page", slug=slug)),
            ("documents", "Documents", url_for("documents_page", slug=slug)),
            ("fiche", "Fiche patient", url_for("patient_edit", slug=slug))]
    out = '<div style="display:flex;gap:2px;border-bottom:2px solid var(--line);margin:14px 0 18px;flex-wrap:wrap">'
    for key, lbl, href in tabs:
        on = (key == active)
        stl = ("border-bottom:3px solid var(--acc);color:var(--acc);font-weight:700"
               if on else "color:var(--mut);border-bottom:3px solid transparent")
        out += '<a href="%s" style="padding:9px 18px;text-decoration:none;margin-bottom:-2px;%s">%s</a>' % (href, stl, lbl)
    return out + '</div>'

def phead(slug, pt, active):
    """En-tête patient commun aux onglets : identité + statut + durée + barre d'onglets."""
    dur = treatment_duration_txt(slug, pt)
    dur_html = (' &middot; <span class=muted>%s</span>' % dur) if dur else ""
    _ct = []
    if pt.get("tel"): _ct.append('&#128222; ' + str(pt.get("tel")))
    if pt.get("email"): _ct.append('&#9993; ' + str(pt.get("email")))
    contact_html = (' &middot; ' + ' &middot; '.join(_ct)) if _ct else ""
    return ('<div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px">'
            '<div><div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap"><h1 style="margin:0">%s</h1>%s</div>'
            '<div class=muted style="margin-top:2px">%s &middot; %s &middot; n\xe9(e) %s%s%s</div></div>'
            '<div style="display:flex;flex-direction:column;align-items:flex-end;gap:8px">'
            '%s<a class="btn sec" href="%s">← Biblioth\xe8que</a></div></div>%s'
            ) % (pt.get("nom", ""), statut_badge(pt), pt.get("age", ""), pt.get("sexe", ""),
                 pt.get("dob_court", ""), dur_html, contact_html,
                 plat_btns(pt.get("nom", ""), size=30), url_for("dashboard"), patient_tabs(slug, active))

@app.route("/patient/<slug>")
def patient(slug):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug)
    if not pt: abort(404)
    # enregistrements triés par date (récent en haut) — permet d'intercaler par la date
    recs_sorted = sorted([(rid, load_rec(slug, rid)) for rid in pt.get("records", [])],
                         key=lambda x: (x[1] or {}).get("date", ""), reverse=True)
    rows = ""
    for rid, r in recs_sorted:
        if not r: continue
        st = {"processing": "<span class=spin></span> ", "error": "&#9888; "}.get(r.get("status"), "")
        dcur = (r.get("date", "") or "")[:10]
        lbl = (r.get("label", "") or "").replace('"', "&quot;")
        datetxt = _date_fr(r.get("date", "")) or "date non renseignée"
        age = age_at(pt.get("dob", ""), r.get("date", ""))
        sub = datetxt + (" · %s" % age if age else "")
        rows += ("""<div class=recrow><span class=recdot></span>
        <div style="flex:1;min-width:0">
          <div class=recname>%s%s</div><div class=recsub>%s</div></div>
        <a class="btn sec sm" href="%s">Ouvrir</a>
        <details class=menu><summary class=dots title="Modifier / supprimer">&#8942;</summary>
          <div class=menu-pop style="min-width:230px;padding:11px">
            <form method=post action="%s">
              <label style="font-size:12px;color:var(--mut)">Date</label>
              <input type=date name=date value="%s" style="width:100%%;margin:3px 0 9px">
              <label style="font-size:12px;color:var(--mut)">Libellé</label>
              <input name=label value="%s" style="width:100%%;margin:3px 0 11px">
              <button class="btn sm" style="width:100%%">Enregistrer</button></form>
            <div class=sep></div>
            <a class=mi href="%s" style="color:#C0392B">&#128465; Supprimer (corbeille)</a>
          </div></details></div>""") % (
            st, r.get("label", ""), sub, url_for("record", slug=slug, rid=rid),
            url_for("record_setmeta", slug=slug, rid=rid), dcur, lbl,
            url_for("record_delete", slug=slug, rid=rid))
    return page("""%s
    <div class=card><div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;gap:10px;flex-wrap:wrap">
    <b>Bilans &amp; réévaluations</b>
    <div style="display:flex;gap:8px;flex-wrap:wrap">
      <a class="btn sec" href="%s" title="Génère un Word complet + un dossier brut (Photos/Radios/STL) dans Téléchargements">&#128194; Transfert de dossier</a>
      <a class=btn href="%s">+ Nouvelle réévaluation</a></div></div>
    <div class=reclist>%s</div></div>""" % (
        phead(slug, pt, "bilans"), url_for("transfert_dossier", slug=slug), url_for("reeval", slug=slug),
        rows or "<div class=recsub style='padding:12px 4px'>Aucun enregistrement.</div>"))

# ======================================================================
#  TRANSFERT DE DOSSIER : Word complet + dossier brut (zip) -> Téléchargements
# ======================================================================
def _safe_name(s):
    s = unicodedata.normalize("NFKD", (s or "")).encode("ascii", "ignore").decode()
    s = re.sub(r"[^A-Za-z0-9 _-]", "", s).strip().replace(" ", "_")
    return s or "patient"

def _docx_text(path):
    try:
        from docx import Document as _D
        d = _D(path)
        parts = [p.text for p in d.paragraphs if (p.text or "").strip()]
        return "\n".join(parts)
    except Exception:
        return ""

def _ordered_records(slug, pt):
    """Bilan initial d'abord, puis par date croissante."""
    recs = [(rid, load_rec(slug, rid)) for rid in pt.get("records", [])]
    recs = [(rid, r) for rid, r in recs if r]
    recs.sort(key=lambda x: (x[1].get("date", "") or "", x[0]))
    return recs

def _rec_label(idx, r):
    lbl = (r.get("label", "") or "").strip()
    if lbl: return lbl
    return "Bilan initial" if idx == 0 else "Réévaluation %d" % idx

@app.route("/patient/<slug>/transfert")
def transfert_dossier(slug):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug)
    if not pt: abort(404)
    import zipfile, tempfile
    import bilan_pdf_rl
    recs = _ordered_records(slug, pt)
    safe = _safe_name(pt.get("nom", "patient"))
    dl = os.path.expanduser("~/Downloads")
    if not os.path.isdir(dl): dl = os.path.expanduser("~/Desktop")
    if not os.path.isdir(dl): dl = os.path.expanduser("~")
    logo = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "chu_nice.png")
    prac = (get_settings().get("praticien") or "").strip() or "Dr Bryan Faruch"
    # --- PDF premium (rendu ReportLab, 100% local) : genere en temporaire, il ira DANS le zip ---
    _tmp = tempfile.mkdtemp()
    pdf_path = os.path.join(_tmp, "Dossier ODF - %s.pdf" % safe)
    try:
        bilan_pdf_rl.build_pdf(slug, pdf_path, practitioner=prac, logo=logo if os.path.exists(logo) else None)
    except Exception as e:
        shutil.rmtree(_tmp, ignore_errors=True)
        flash("Erreur lors de la generation du PDF : %s" % e)
        return redirect(url_for("patient", slug=slug))
    # --- dossier complet zippe (le PDF + Photos / Radios / STL par temps) ---
    zip_path = os.path.join(dl, "Dossier ODF - %s.zip" % safe)
    _n = 1
    while os.path.exists(zip_path):
        zip_path = os.path.join(dl, "Dossier ODF - %s (%d).zip" % (safe, _n)); _n += 1
    root = "Dossier ODF - %s" % safe
    RAW = [("Photos", "02_photos_traitees"), ("Photos_brutes", "01_photos_brutes"),
           ("Radios", "03_radios"), ("Modeles_STL", "04_stl_bruts"), ("Rendus_3D", "05_stl_rendus"),
           ("Traces_WebCeph", "06_webceph")]
    try:
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
            if os.path.exists(pdf_path):
                z.write(pdf_path, arcname="%s/%s" % (root, os.path.basename(pdf_path)))
            for idx, (rid, r) in enumerate(recs):
                base = rdir(slug, rid)
                folder = "%02d - %s" % (idx + 1, _safe_name(_rec_label(idx, r)))
                for label, sub in RAW:
                    sd = os.path.join(base, sub)
                    if not os.path.isdir(sd): continue
                    for f in sorted(os.listdir(sd)):
                        if f.startswith("."): continue
                        fp = os.path.join(sd, f)
                        if os.path.isfile(fp):
                            z.write(fp, arcname="%s/%s/%s/%s" % (root, folder, label, f))
    except Exception as e:
        shutil.rmtree(_tmp, ignore_errors=True)
        flash("Erreur lors de la creation du dossier : %s" % e)
        return redirect(url_for("patient", slug=slug))
    shutil.rmtree(_tmp, ignore_errors=True)
    flash("Dossier enregistre dans Telechargements : %s" % os.path.basename(zip_path))
    return redirect(url_for("patient", slug=slug))

# ======================================================================
#  DOCUMENTS DU PATIENT (fichiers libres : PDF, Word, radios, photos...)
# ======================================================================
def _docs_dir(slug):
    d = os.path.join(pdir(slug), "documents"); os.makedirs(d, exist_ok=True); return d

def _docs_load(slug):
    try: return json.load(open(os.path.join(_docs_dir(slug), "_docs.json"), encoding="utf-8"))
    except Exception: return []

def _docs_save(slug, items):
    json.dump(items, open(os.path.join(_docs_dir(slug), "_docs.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)

def _human_size(n):
    n = float(n)
    for u in ("o", "Ko", "Mo", "Go"):
        if n < 1024 or u == "Go":
            return ("%d %s" % (int(n), u)) if u == "o" else ("%.1f %s" % (n, u))
        n /= 1024.0

def _doc_badge(ext):
    ext = (ext or "").lower()
    if ext == "pdf": return ("#C0392B", "PDF")
    if ext in ("doc", "docx"): return ("#2B579A", "DOC")
    if ext in ("xls", "xlsx", "csv"): return ("#1E7D45", "XLS")
    if ext in ("jpg", "jpeg", "png", "gif", "webp", "heic", "tif", "tiff", "bmp"): return ("#8E44AD", "IMG")
    if ext in ("ppt", "pptx"): return ("#D24726", "PPT")
    if ext in ("zip", "rar", "7z"): return ("#7F8C8D", "ZIP")
    if ext in ("stl", "ply", "obj"): return ("#0E7C7B", "3D")
    return ("#5B6B7F", (ext[:3].upper() if ext else "FIC"))

def _doc_previewable(ext):
    return (ext or "").lower() in ("pdf", "jpg", "jpeg", "png", "gif", "webp", "bmp", "tif", "tiff")

@app.route("/patient/<slug>/documents")
def documents_page(slug):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug)
    if not pt: abort(404)
    items = sorted(_docs_load(slug), key=lambda x: x.get("date", ""), reverse=True)
    rows = ""
    for it in items:
        col, tag = _doc_badge(it.get("ext"))
        did = it.get("id"); nom = (it.get("name") or "document").replace("<", "&lt;")
        nom_val = nom.replace('"', "&quot;")
        dt = _date_fr(it.get("date", "")) or (it.get("date", "") or "")[:10]
        sz = _human_size(it.get("size", 0) or 0)
        ext = (it.get("ext") or "").lower()
        is_img = 1 if ext in ("jpg", "jpeg", "png", "gif", "webp", "bmp", "tif", "tiff") else 0
        open_url = url_for("documents_file", slug=slug, did=did)
        dl_url = open_url + "?dl=1"
        del_url = url_for("documents_delete", slug=slug, did=did)
        ren_url = url_for("documents_rename", slug=slug, did=did)
        voir = ('<button type=button class="btn sec sm" onclick="docView(\'%s\',%d)">&#128065; Voir</button> '
                % (open_url, is_img)) if _doc_previewable(ext) else ""
        menu = ('<details class=menu><summary class=dots title="Renommer / supprimer">&#8942;</summary>'
                '<div class=menu-pop style="min-width:250px;padding:11px;right:0;left:auto">'
                '<form method=post action="%s"><label style="font-size:12px;color:var(--mut)">Nom du document</label>'
                '<input name=name value="%s" style="width:100%%;margin:3px 0 9px"><button class="btn sm" style="width:100%%">Renommer</button></form>'
                '<div class=sep></div>'
                '<form method=post action="%s" onsubmit="return confirm(\'Supprimer ce document ?\')">'
                '<button type=submit class=mi style="color:#C0392B;width:100%%;text-align:left;background:none;border:none;cursor:pointer;padding:8px 6px">&#128465; Supprimer</button></form>'
                '</div></details>') % (ren_url, nom_val, del_url)
        rows += ('<div class=drow style="display:flex;align-items:center;gap:14px;padding:11px 4px;border-bottom:1px solid var(--line)">'
                 '<span style="flex:0 0 auto;width:44px;height:44px;border-radius:10px;background:%s;color:#fff;'
                 'display:flex;align-items:center;justify-content:center;font-weight:800;font-size:12px;letter-spacing:.5px">%s</span>'
                 '<div style="flex:1;min-width:0"><div style="font-weight:700;word-break:break-word">%s</div>'
                 '<div class=muted style="font-size:12.5px">%s &middot; %s</div></div>'
                 '%s<a class="btn sec sm" href="%s">&#8681; Telecharger</a> %s'
                 '</div>') % (col, tag, nom, dt, sz, voir, dl_url, menu)
    if not items:
        rows = '<div class=muted style="padding:22px 4px;text-align:center">Aucun document. Ajoute un fichier ci-dessus (PDF, Word, radio, photo...).</div>'
    # visionneuse en superposition (aperçu sans téléchargement)
    rows += ('<div id=docov style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.82);z-index:9999;'
             'align-items:center;justify-content:center" onclick="if(event.target===this)docClose()">'
             '<div style="position:relative;width:92%;height:92%;background:#111;border-radius:10px;overflow:hidden;box-shadow:0 10px 40px rgba(0,0,0,.5)">'
             '<button type=button onclick="docClose()" style="position:absolute;top:8px;right:10px;z-index:2;width:34px;height:34px;'
             'border-radius:50%;border:none;background:rgba(255,255,255,.92);cursor:pointer;font-size:16px;font-weight:700">&times;</button>'
             '<div id=docbody style="width:100%;height:100%"></div></div></div>'
             '<script>function docView(u,img){var o=document.getElementById("docov"),b=document.getElementById("docbody");'
             'b.innerHTML=img?(\'<img src="\'+u+\'" style="width:100%;height:100%;object-fit:contain;background:#111">\')'
             ':(\'<iframe src="\'+u+\'" style="width:100%;height:100%;border:0;background:#fff"></iframe>\');'
             'o.style.display="flex";}'
             'function docClose(){var o=document.getElementById("docov");o.style.display="none";'
             'document.getElementById("docbody").innerHTML="";}'
             'document.addEventListener("keydown",function(e){if(e.key==="Escape")docClose();});</script>')
    _tpl = ('<div class=card><h2 style="margin-top:0">Ajouter des documents</h2>'
            '<p class=muted style="margin-top:0">Consentement, courrier d\'adressage, analyse medicale, radio, photo... tout type de fichier.</p>'
            '<form method=post action="%s" enctype=multipart/form-data style="display:flex;gap:10px;flex-wrap:wrap;align-items:center">'
            '<input type=file name=files multiple required style="flex:1;min-width:240px">'
            '<button class=btn type=submit onclick="this.innerHTML=\'<span class=spin></span> Ajout...\'">Ajouter</button>'
            '</form></div>'
            '<div class=card><h2 style="margin-top:0">Documents (%d)</h2>%s</div>'
            ) % (url_for("documents_upload", slug=slug), len(items), rows)
    body = phead(slug, pt, "documents") + _tpl
    return page(body)

@app.route("/patient/<slug>/documents/upload", methods=["POST"])
def documents_upload(slug):
    if not logged(): return redirect(url_for("login"))
    if not load_patient(slug): abort(404)
    files = request.files.getlist("files")
    items = _docs_load(slug); n = 0
    for f in files:
        if not f or not f.filename: continue
        name = os.path.basename(f.filename.replace("\\", "/"))
        ext = os.path.splitext(name)[1].lower().lstrip(".")
        did = "%d_%s" % (int(datetime.datetime.now().timestamp() * 1000), os.urandom(3).hex())
        stored = did + (("." + ext) if ext else "")
        dest = os.path.join(_docs_dir(slug), stored)
        try:
            f.save(dest)
            items.append({"id": did, "stored": stored, "name": name, "ext": ext,
                          "date": datetime.datetime.now().isoformat(timespec="seconds"),
                          "size": os.path.getsize(dest)})
            n += 1
        except Exception:
            pass
    _docs_save(slug, items)
    flash(("%d document(s) ajoute(s)." % n) if n else "Aucun fichier importe.")
    return redirect(url_for("documents_page", slug=slug))

@app.route("/patient/<slug>/documents/file/<did>")
def documents_file(slug, did):
    if not logged(): return redirect(url_for("login"))
    it = next((x for x in _docs_load(slug) if x.get("id") == did), None)
    if not it: abort(404)
    fp = os.path.join(_docs_dir(slug), it.get("stored", ""))
    if not os.path.exists(fp): abort(404)
    return send_file(fp, as_attachment=bool(request.args.get("dl")),
                     download_name=it.get("name") or "document")

@app.route("/patient/<slug>/documents/delete/<did>", methods=["POST"])
def documents_delete(slug, did):
    if not logged(): return redirect(url_for("login"))
    items = _docs_load(slug)
    it = next((x for x in items if x.get("id") == did), None)
    if it:
        src = os.path.join(_docs_dir(slug), it.get("stored", ""))
        cb = os.path.join(_docs_dir(slug), "_corbeille"); os.makedirs(cb, exist_ok=True)
        try:
            if os.path.exists(src): shutil.move(src, os.path.join(cb, it.get("stored")))
        except Exception: pass
        _docs_save(slug, [x for x in items if x.get("id") != did])
        flash("Document supprime (corbeille).")
    return redirect(url_for("documents_page", slug=slug))

@app.route("/patient/<slug>/documents/rename/<did>", methods=["POST"])
def documents_rename(slug, did):
    if not logged(): return redirect(url_for("login"))
    items = _docs_load(slug)
    it = next((x for x in items if x.get("id") == did), None)
    if it:
        newname = os.path.basename((request.form.get("name", "") or "").strip().replace("\\", "/"))
        if newname:
            base, ne = os.path.splitext(newname)
            if not ne and it.get("ext"):
                newname = newname + "." + it["ext"]      # garde l'extension d'origine
            elif ne:
                it["ext"] = ne.lower().lstrip(".")
            it["name"] = newname
            _docs_save(slug, items)
            flash("Document renomme.")
    return redirect(url_for("documents_page", slug=slug))

STAFF_QA_TPL = r'''<form class=eslide data-autosave-url="__SURL__" style="display:flex;flex-direction:column;height:88vh;width:100%;max-width:1400px;gap:1.1vh;padding:0 2vw">
<div style="flex:1;min-height:0;display:flex;flex-direction:column">
<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:.5vh">
<span style="font-size:clamp(13px,1.9vh,18px);color:#9fb4d6;text-transform:uppercase;letter-spacing:.6px">Questions</span>
<span class=autosave-status style="font-size:13px;color:#7f8ea8"></span></div>
<textarea name=questions placeholder="Une question par ligne..." style="flex:1;width:100%;font-size:clamp(16px,2.3vh,24px);padding:14px 16px;border:none;border-radius:12px;background:#f5f8fc;color:#12233a;resize:none;line-height:1.5">__Q__</textarea>
</div>
<div style="flex:1;min-height:0;display:flex;flex-direction:column">
<span style="font-size:clamp(13px,1.9vh,18px);color:#57d98a;text-transform:uppercase;letter-spacing:.6px;margin-bottom:.5vh">R&eacute;ponses / d&eacute;cisions</span>
<textarea name=answers placeholder="Reponses decidees au staff..." style="flex:1;width:100%;font-size:clamp(16px,2.3vh,24px);padding:14px 16px;border:none;border-radius:12px;background:#eef7ef;color:#12233a;resize:none;line-height:1.5">__A__</textarea>
</div>
<div style="display:flex;justify-content:flex-end;align-items:center;gap:14px;padding:.2vh 0">
<span class=exp-status style="font-size:14px;color:#8ea3c4"></span>
<button type=button onclick="exportDecision(this)" data-exp-url="__EXPURL__" style="background:linear-gradient(135deg,#2f7ff0,#2f6ad0);color:#fff;border:none;border-radius:11px;padding:11px 20px;font-size:15px;font-weight:600;cursor:pointer;box-shadow:0 4px 14px rgba(47,127,240,.35)">&#8595;&nbsp; Exporter la d&eacute;cision dans le suivi</button>
</div>
</form>'''

STAFF_PLAN_TPL = r'''<form class=eslide data-autosave-url="__PURL__" style="display:flex;flex-direction:column;height:88vh;width:100%;max-width:1400px;gap:1.4vh;padding:0 2vw">
<div style="flex:1;min-height:0;display:flex;flex-direction:column">
<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:.6vh">
<span style="font-size:clamp(13px,1.9vh,18px);color:#9fb4d6;text-transform:uppercase;letter-spacing:.6px">Objectifs</span>
<span class=autosave-status style="font-size:13px;color:#7f8ea8"></span></div>
<textarea name=objectifs placeholder="Objectifs du traitement..." style="flex:1;width:100%;font-size:clamp(16px,2.3vh,24px);padding:14px 16px;border:none;border-radius:12px;background:#f5f8fc;color:#12233a;resize:none;line-height:1.5">__OBJ__</textarea>
</div>
<div style="flex:1.5;min-height:0;display:flex;flex-direction:column">
<span style="font-size:clamp(13px,1.9vh,18px);color:#7db8ff;text-transform:uppercase;letter-spacing:.6px;margin-bottom:.6vh">Moyens th&eacute;rapeutiques</span>
<textarea name=moyens placeholder="Moyens therapeutiques..." style="flex:1;width:100%;font-size:clamp(16px,2.3vh,24px);padding:14px 16px;border:none;border-radius:12px;background:#eef3fb;color:#12233a;resize:none;line-height:1.5">__MOY__</textarea>
</div></form>'''

# ======================================================================
#  RUBRIQUE STAFF (presentation prete a montrer : temps + evolution + questions)
# ======================================================================
def _hesc(s):
    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def _staff_list(pt):
    return sorted(pt.get("staffs") or [], key=lambda x: (x.get("date", "") or ""), reverse=True)

def _sl(title, inner, tp=""):
    t = ('<div class=stitle>%s</div>' % title) if title else ""
    tpa = (' data-tp="%s"' % tp.replace('"', "&quot;")) if tp else ""
    return '<section class=pslide%s>%s<div class=pbody>%s</div></section>' % (tpa, t, inner)

def _recs_chrono(slug, pt):
    rs = [(rid, load_rec(slug, rid)) for rid in pt.get("records", [])]
    rs = [(rid, r) for rid, r in rs if r]
    rs.sort(key=lambda x: (x[1].get("date", "") or "", x[0]))
    return rs

def _rec_labels(slug, pt):
    out = {}
    for idx, (rid, r) in enumerate(_recs_chrono(slug, pt)):
        out[rid] = r.get("label", "") or ("Bilan initial" if idx == 0 else "Reevaluation %d" % idx)
    return out

def _rec_options_html(slug, pt, checked=None, name="records"):
    checked = set(checked or [])
    out = ""
    for idx, (rid, r) in enumerate(_recs_chrono(slug, pt)):
        lbl = r.get("label", "") or ("Bilan initial" if idx == 0 else "Reevaluation %d" % idx)
        dd = _date_fr(r.get("date", "")) or (r.get("date", "") or "")[:10]
        ck = " checked" if rid in checked else ""
        out += ('<label style="display:flex;align-items:center;gap:9px;padding:8px 4px;border-bottom:1px solid var(--line);cursor:pointer">'
                '<input type=checkbox name=%s value="%s"%s style="width:auto">'
                '<span><b>%s</b> <span class=muted>&middot; %s</span></span></label>') % (name, rid, ck, _hesc(lbl), dd)
    return out or '<div class=muted>Aucun temps enregistre.</div>'

EVO_VIEWS = [
    ("exo_face_repos", "Visage au repos", "photo"),
    ("exo_face_sourire", "Sourire", "photo"),
    ("exo_profil", "Profil", "photo"),
    ("endo_occlusion_frontale", "Occlusion frontale", "photo"),
    ("endo_laterale_droite", "Laterale droite", "photo"),
    ("endo_laterale_gauche", "Laterale gauche", "photo"),
    ("endo_occlusal_maxillaire", "Occlusal maxillaire", "photo"),
    ("endo_occlusal_mandibulaire", "Occlusal mandibulaire", "photo"),
    ("panoramique", "Panoramique", "radio"),
    ("teleradiographie_profil", "Teleradiographie de profil", "radio"),
]

def _view_options_html(checked=None):
    checked = set(checked or [])
    out = ""
    for key, lbl, kind in EVO_VIEWS:
        ck = " checked" if key in checked else ""
        out += ('<label style="display:inline-flex;align-items:center;gap:7px;padding:6px 10px;margin:0 6px 6px 0;'
                'border:1px solid var(--line);border-radius:999px;cursor:pointer">'
                '<input type=checkbox name=evo_views value="%s"%s style="width:auto">%s</label>') % (key, ck, _hesc(lbl))
    return out

def _view_src(slug, rid, key, kind):
    base = rdir(slug, rid)
    if kind == "radio":
        rel = "03_radios/%s.jpg" % key
        return url_for("rfile", slug=slug, rid=rid, path=rel) if os.path.exists(os.path.join(base, rel)) else None
    for rel in ("02_photos_traitees/%s.jpg" % key, "01_photos_brutes/%s.jpg" % key):
        if os.path.exists(os.path.join(base, rel)):
            return url_for("rfile", slug=slug, rid=rid, path=rel)
    return None

def _slides_for_record(slug, rid, pt, r, section_label=None, is_initial=False):
    """Diapos d'un temps : separateur + photos + radios + traces + 3D + Steiner + diagnostic + plan."""
    base = rdir(slug, rid)
    PLBL = dict(PHOTO_FIELDS)
    def psrc(rel): return url_for("rfile", slug=slug, rid=rid, path=rel)
    def photo_src(key):
        for rel in ("02_photos_traitees/%s.jpg" % key, "01_photos_brutes/%s.jpg" % key):
            if os.path.exists(os.path.join(base, rel)): return psrc(rel)
        return None
    label = section_label or (r.get("label", "") or "Bilan")
    date = _date_fr(r.get("date", "")) or (r.get("date", "") or "")[:10]
    age = age_at(pt.get("dob", ""), r.get("date", ""))
    tp = "%s%s" % (label, (" · " + date) if date else "")
    S = []
    dmeta = " &middot; ".join([x for x in (date, age) if x])
    div = ('<div class=divider><div class=dkick>Temps clinique</div>'
           '<div class=dlabel>%s</div><div class=dmeta>%s</div></div>') % (_hesc(label), dmeta)
    S.append(_sl("", div, tp))
    def photo_slide(title, keys):
        cells = []
        for key in keys:
            src = photo_src(key)
            if src:
                cells.append('<figure class=pfig><img loading=lazy src="%s" onclick="zoom(this)">'
                             '<figcaption>%s</figcaption></figure>' % (src, PLBL.get(key, key)))
        if cells:
            S.append(_sl(title, '<div class="pgrid n%d">%s</div>' % (len(cells), "".join(cells)), tp))
    photo_slide("Photographies exobuccales", ["exo_face_repos", "exo_face_sourire", "exo_profil"])
    photo_slide("Vues endobuccales", ["endo_laterale_droite", "endo_occlusion_frontale", "endo_laterale_gauche"])
    photo_slide("Vues occlusales", ["endo_occlusal_maxillaire", "endo_occlusal_mandibulaire"])
    rad_cells = []
    for k, lbl in (("panoramique", "Panoramique"), ("teleradiographie_profil", "T&eacute;l&eacute;radiographie de profil")):
        rel = "03_radios/%s.jpg" % k
        if os.path.exists(os.path.join(base, rel)):
            rad_cells.append('<figure class=pfig><img loading=lazy src="%s" onclick="zoom(this)">'
                             '<figcaption>%s</figcaption></figure>' % (psrc(rel), lbl))
    if rad_cells:
        S.append(_sl("Radiographies", '<div class="pgrid n%d">%s</div>' % (len(rad_cells), "".join(rad_cells)), tp))
    wc = os.path.join(base, "06_webceph"); tr_cells = []
    if os.path.isdir(wc):
        for tf in sorted(os.listdir(wc)):
            if tf.lower().endswith((".jpg", ".jpeg", ".png")):
                tr_cells.append('<figure class=pfig><img loading=lazy src="%s" onclick="zoom(this)"></figure>'
                                % psrc("06_webceph/%s" % tf))
    if tr_cells:
        S.append(_sl("Trac&eacute; c&eacute;phalom&eacute;trique",
                     '<div class="pgrid n%d">%s</div>' % (min(len(tr_cells), 2), "".join(tr_cells[:2])), tp))
    rend = []
    for k, lbl in RENDER_FIELDS:
        rel = "05_stl_rendus/%s.png" % k
        if os.path.exists(os.path.join(base, rel)):
            rend.append((psrc(rel), lbl))
    if rend:
        ncol = 3 if len(rend) >= 3 else len(rend)
        nrows = (len(rend) + ncol - 1) // ncol
        maxh = max(20, int((80 - 5 * nrows) / nrows))
        figs = "".join('<figure class=pfig style="height:auto;justify-content:flex-start">'
                       '<img loading=lazy src="%s" onclick="zoom(this)" style="max-height:%dvh">'
                       '<figcaption>%s</figcaption></figure>' % (src, maxh, lbl) for src, lbl in rend)
        grid = ('<div style="display:grid;grid-template-columns:repeat(%d,1fr);gap:8px 16px;'
                'width:100%%;max-height:84vh;align-content:center;justify-items:center;margin-top:2vh">%s</div>') % (ncol, figs)
        S.append(_sl("Mod&egrave;les 3D", grid, tp))
    st = r.get("steiner") or {}
    srows = []
    for nom, short, mean, sd in STEINER_FIELDS:
        v = st.get(nom)
        if v in (None, ""):
            continue
        try: v = float(v)
        except Exception: continue
        z = (v - mean) / sd if sd else 0
        cls = "ok" if abs(z) <= 1 else ("warn" if abs(z) <= 2 else "bad")
        sig, sev = steiner_signif(nom, v, mean, sd)
        srows.append('<div class="srow %s"><span class=sname>%s</span><span class=sval>%g</span>'
                     '<span class=snorm>norme %g</span><span class=ssig>%s</span></div>'
                     % (cls, short, v, mean, sig if sev else "Dans la norme"))
    if srows:
        S.append(_sl("Analyse c&eacute;phalom&eacute;trique de Steiner", '<div class=steiner>%s</div>' % "".join(srows), tp))
    syn = r.get("synthese", {}) or {}
    dcols = []
    for cl, cc in DIAG_COLS:
        items = []
        for rl, rc in DIAG_ROWS:
            val = (syn.get(cc + "_" + rc) or "").strip()
            if val: items.append('<div class=drow><span class=drl>%s</span><span class=dv>%s</span></div>' % (rl, val))
        if items: dcols.append('<div class=dcol><h3>%s</h3>%s</div>' % (cl, "".join(items)))
    if dcols:
        S.append(_sl("Synth&egrave;se diagnostique", '<div class=diag>%s</div>' % "".join(dcols), tp))
    # Plan de traitement (objectifs / moyens) : editable + synchro dossier.
    # Affiche si bilan initial OU si deja rempli.
    plan = r.get("plan") or {}
    obj = str(plan.get("objectifs", "") or "")
    moy = str(plan.get("moyens", "") or "")
    if is_initial or obj.strip() or moy.strip():
        purl = url_for("plan_obj_save", slug=slug, rid=rid)
        ptpl = (STAFF_PLAN_TPL.replace("__OBJ__", _hesc(obj))
                .replace("__MOY__", _hesc(moy)).replace("__PURL__", purl))
        S.append(_sl("Plan de traitement", ptpl, tp))
    return S

def _evolution_slides(slug, pt, evo_records, evo_views):
    order = [rid for rid, _r in _recs_chrono(slug, pt) if rid in set(evo_records or [])]
    if len(order) < 2 or not evo_views:
        return []
    labels = _rec_labels(slug, pt)
    dates = {rid: (_date_fr((load_rec(slug, rid) or {}).get("date", "")) or "") for rid in order}
    S = []
    for key, vlabel, kind in EVO_VIEWS:
        if key not in evo_views:
            continue
        cells = []
        for rid in order:
            src = _view_src(slug, rid, key, kind)
            if not src:
                continue
            cap = _hesc(labels.get(rid, "")) + (("<br><span style='opacity:.7'>%s</span>" % dates.get(rid, "")) if dates.get(rid) else "")
            cells.append('<figure class=pfig><img loading=lazy src="%s" onclick="zoom(this)">'
                         '<figcaption>%s</figcaption></figure>' % (src, cap))
        if len(cells) >= 2:
            S.append(_sl("Évolution &mdash; %s" % _hesc(vlabel),
                         '<div class="pgrid n%d">%s</div>' % (min(len(cells), 6), "".join(cells)), "Évolution"))
    return S

def _staff_form(slug, pt, action, date, records, questions, evo_records, evo_views, submit):
    return ('<form method=post action="%s">'
            '<label>Date</label><input type=date name=date value="%s" style="max-width:220px">'
            '<label style="margin-top:10px;display:block">Temps a presenter <span class=muted>(dans l\'ordre chronologique)</span></label>'
            '<div style="border:1px solid var(--line);border-radius:10px;padding:2px 10px;margin-top:4px;max-height:200px;overflow:auto">%s</div>'
            '<label style="margin-top:12px;display:block">Questions a poser <span class=muted>(une par ligne)</span></label>'
            '<textarea name=questions rows=3 style="width:100%%" placeholder="Ex : Faut-il extraire 14/24 ?">%s</textarea>'
            '<details style="margin-top:12px;border:1px solid var(--line);border-radius:10px;padding:10px 12px">'
            '<summary style="cursor:pointer;font-weight:700;color:var(--acc)">&#43; Ajouter une evolution (comparatif avant/apres) &mdash; optionnel</summary>'
            '<div style="margin-top:10px">'
            '<label style="font-size:12px;color:var(--mut)">Temps a comparer <span class=muted>(coche 2 temps ou plus)</span></label>'
            '<div style="border:1px solid var(--line);border-radius:10px;padding:2px 10px;margin:4px 0 10px;max-height:170px;overflow:auto">%s</div>'
            '<label style="font-size:12px;color:var(--mut)">Vues a comparer</label>'
            '<div style="margin-top:6px">%s</div>'
            '<div class=muted style="font-size:12px;margin-top:6px">Une diapo d\'evolution par vue s\'ajoutera a la fin. Laisse vide pour ne rien ajouter.</div>'
            '</div></details>'
            '<div style="margin-top:12px"><button class=btn onclick="this.innerHTML=\'<span class=spin></span> ...\'">%s</button></div>'
            '</form>') % (action, date, _rec_options_html(slug, pt, records), _hesc(questions),
                         _rec_options_html(slug, pt, evo_records, name="evo_records"), _view_options_html(evo_views), submit)

@app.route("/patient/<slug>/staff")
def staff_page(slug):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug)
    if not pt: abort(404)
    entries = _staff_list(pt)
    cards = ""
    for e in entries:
        sid = e.get("id"); d = _date_fr(e.get("date", "")) or (e.get("date", "") or "")[:10] or "date ?"
        ntemps = len([x for x in (e.get("records") or []) if x in pt.get("records", [])])
        nq = len([l for l in (e.get("questions", "") or "").splitlines() if l.strip()])
        nevo = len(e.get("evo_records") or []) if (e.get("evo_records") and e.get("evo_views")) else 0
        show = url_for("staff_show", slug=slug, sid=sid)
        efrm = _staff_form(slug, pt, url_for("staff_edit", slug=slug, sid=sid), (e.get("date", "") or "")[:10],
                           e.get("records"), e.get("questions", ""), e.get("evo_records"), e.get("evo_views"), "Enregistrer")
        edit = ('<details class=menu><summary class="btn sec sm" style="list-style:none">&#9998; Modifier / supprimer</summary>'
                '<div class="menu-pop" style="right:0;left:auto;min-width:380px;max-width:94vw;padding:12px">%s'
                '<div class=sep></div>'
                '<form method=post action="%s" onsubmit="return confirm(\'Supprimer ce staff ?\')">'
                '<button type=submit class=mi style="color:#C0392B;width:100%%;text-align:left;background:none;border:none;cursor:pointer;padding:8px 6px">&#128465; Supprimer</button></form>'
                '</div></details>') % (efrm, url_for("staff_delete", slug=slug, sid=sid))
        evo_txt = (" &middot; evolution (%d temps)" % nevo) if nevo else ""
        cards += ('<div class=card><div style="display:flex;justify-content:space-between;align-items:center;gap:12px">'
                  '<a href="%s" style="text-decoration:none;color:inherit;display:flex;align-items:center;gap:12px">'
                  '<span style="width:46px;height:46px;border-radius:12px;background:linear-gradient(135deg,#16324f,#2f5ca8);color:#fff;'
                  'display:flex;align-items:center;justify-content:center;font-size:20px">&#9654;</span>'
                  '<span><span style="font-weight:800;font-size:16px">Staff du %s</span><br>'
                  '<span class=muted style="font-size:13px">%d temps &middot; %d question(s)%s &mdash; cliquer pour ouvrir</span></span></a>'
                  '%s</div></div>') % (show, d, ntemps, nq, evo_txt, edit)
    if not entries:
        cards = '<div class=card center muted>Aucun staff. Cree un staff : choisis les temps a montrer et tes questions.</div>'
    today = datetime.date.today().isoformat()
    add = ('<details class=card style="margin-bottom:16px">'
           '<summary style="cursor:pointer;list-style:none;display:flex;align-items:center;gap:11px;font-weight:800;font-size:16px">'
           '<span style="width:34px;height:34px;border-radius:9px;background:linear-gradient(135deg,#16324f,#2f5ca8);color:#fff;'
           'display:inline-flex;align-items:center;justify-content:center;font-size:22px;line-height:1">+</span> Nouveau staff</summary>'
           '<div style="margin-top:14px">%s</div></details>') % _staff_form(
               slug, pt, url_for("staff_add", slug=slug), today, None, "", None, None, "Creer le staff")
    hdr = '<h2 style="margin:0 0 6px">Staffs enregistr\xe9s</h2>'
    return page(phead(slug, pt, "staff") + add + hdr + cards)

@app.route("/patient/<slug>/staff/add", methods=["POST"])
def staff_add(slug):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug)
    if not pt: abort(404)
    e = {"id": "%d_%s" % (int(datetime.datetime.now().timestamp() * 1000), os.urandom(3).hex()),
         "date": (request.form.get("date", "") or datetime.date.today().isoformat())[:10],
         "records": request.form.getlist("records"),
         "questions": (request.form.get("questions", "") or "").strip(),
         "evo_records": request.form.getlist("evo_records"),
         "evo_views": request.form.getlist("evo_views")}
    if not e["records"] and not e["questions"]:
        flash("Choisis au moins un temps ou ecris des questions.")
        return redirect(url_for("staff_page", slug=slug))
    pt.setdefault("staffs", []).append(e); save_patient(slug, pt)
    flash("Staff cree.")
    return redirect(url_for("staff_page", slug=slug))

@app.route("/patient/<slug>/staff/edit/<sid>", methods=["POST"])
def staff_edit(slug, sid):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug)
    if not pt: abort(404)
    for e in pt.get("staffs", []) or []:
        if e.get("id") == sid:
            e["date"] = (request.form.get("date", "") or e.get("date", ""))[:10]
            e["records"] = request.form.getlist("records")
            e["questions"] = (request.form.get("questions", "") or "").strip()
            e["evo_records"] = request.form.getlist("evo_records")
            e["evo_views"] = request.form.getlist("evo_views")
            save_patient(slug, pt); flash("Staff modifie."); break
    return redirect(url_for("staff_page", slug=slug))

@app.route("/patient/<slug>/staff/delete/<sid>", methods=["POST"])
def staff_delete(slug, sid):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug)
    if not pt: abort(404)
    pt["staffs"] = [e for e in (pt.get("staffs", []) or []) if e.get("id") != sid]
    save_patient(slug, pt); flash("Staff supprime.")
    return redirect(url_for("staff_page", slug=slug))

@app.route("/patient/<slug>/staff/<sid>/show")
def staff_show(slug, sid):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug)
    if not pt: abort(404)
    e = next((x for x in (pt.get("staffs") or []) if x.get("id") == sid), None)
    if not e: abort(404)
    S = []
    dd = _date_fr(e.get("date", "")) or (e.get("date", "") or "")[:10]
    age = pt.get("age", "")
    cover = ('<div class=cover><div class=ckick>Staff</div>'
             '<div class=cnom>%s</div><div class=clabel>Staff du %s</div>'
             '<div class=cmeta>%s%s</div></div>') % (
        _hesc(pt.get("nom", "")), _hesc(dd),
        _hesc(pt.get("dob_court", "")), (" &middot; " + _hesc(age)) if age else "")
    S.append(_sl("", cover))
    chosen = set(rid for rid in (e.get("records") or []) if rid in pt.get("records", []))
    labels = _rec_labels(slug, pt)
    chrono = _recs_chrono(slug, pt)
    initial_rid = chrono[0][0] if chrono else None
    for rid, _r in chrono:
        if rid in chosen:
            r = load_rec(slug, rid)
            if r: S += _slides_for_record(slug, rid, pt, r, section_label=labels.get(rid),
                                          is_initial=(rid == initial_rid))
    S += _evolution_slides(slug, pt, e.get("evo_records"), e.get("evo_views"))
    qtext = _hesc(e.get("questions", "") or "")
    atext = _hesc(e.get("answers", "") or "")
    save_url = url_for("staff_save", slug=slug, sid=sid)
    exp_url = url_for("staff_export_suivi", slug=slug, sid=sid)
    qa = ((STAFF_QA_TPL).replace("__Q__", qtext).replace("__A__", atext)
          .replace("__SURL__", save_url).replace("__EXPURL__", exp_url))
    S.append(_sl("Questions &amp; r\xe9ponses", qa))
    html = (PRESENT_TPL.replace("__NOM__", pt.get("nom", ""))
            .replace("__BACK__", url_for("staff_page", slug=slug))
            .replace("__SLIDES__", "".join(S)))
    return html

@app.route("/patient/<slug>/staff/<sid>/save", methods=["POST"])
def staff_save(slug, sid):
    if not logged(): return ("", 403)
    pt = load_patient(slug)
    if not pt: return ("", 404)
    for e in pt.get("staffs", []) or []:
        if e.get("id") == sid:
            if "questions" in request.form: e["questions"] = request.form.get("questions", "") or ""
            if "answers" in request.form: e["answers"] = request.form.get("answers", "") or ""
            save_patient(slug, pt)
            return ("ok", 200)
    return ("", 404)

@app.route("/record/<slug>/<rid>/planobj/save", methods=["POST"])
def plan_obj_save(slug, rid):
    """Enregistre uniquement objectifs/moyens du plan (depuis la diapo staff),
    en preservant les autres champs du plan (espace, chevrons). Synchro deux sens
    avec la rubrique bilan initial (meme champ r['plan'])."""
    if not logged(): return ("", 403)
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r: return ("", 404)
    plan = r.get("plan") or {"espace": {}, "chevrons": {}, "objectifs": "", "moyens": ""}
    if "objectifs" in request.form: plan["objectifs"] = request.form.get("objectifs", "") or ""
    if "moyens" in request.form: plan["moyens"] = request.form.get("moyens", "") or ""
    r["plan"] = plan
    save_rec(slug, rid, r)
    return ("ok", 200)

@app.route("/patient/<slug>/staff/<sid>/export_suivi", methods=["POST"])
def staff_export_suivi(slug, sid):
    """Exporte la decision (reponses du staff) dans le suivi du patient, date du jour."""
    if not logged(): return ("", 403)
    pt = load_patient(slug)
    if not pt: return ("", 404)
    e = next((x for x in (pt.get("staffs") or []) if x.get("id") == sid), None)
    if not e: return ("", 404)
    txt = (request.form.get("answers", "") or e.get("answers", "") or "").strip()
    if not txt: return ("empty", 200)
    dd = _date_fr(e.get("date", "")) or (e.get("date", "") or "")[:10]
    body = ("D\xe9cision de staff" + ((" du %s" % dd) if dd else "") + "\n" + txt)
    nt = {"id": datetime.datetime.now().strftime("%Y%m%d%H%M%S") + secrets.token_hex(2),
          "date": datetime.date.today().isoformat(), "text": body}
    pt.setdefault("notes", []).append(nt)
    save_patient(slug, pt)
    return ("ok", 200)

# ======================================================================
#  MISE A JOUR AUTOMATIQUE (verifiee en tache de fond, appliquee dans l'app)
#  - version.json + fichiers de code lus sur un depot GitHub public
#  - ne touche JAMAIS aux donnees patients (~/BilanODF_Data)
# ======================================================================
GH_OWNER_UPD = "bryanfaruch-alt"
GH_REPO_UPD = "bilan-odf-app"
GH_BRANCH_UPD = "main"
UPDATE_INFO = None          # {"version":.., "notes":.., "man":..} si une MAJ plus recente existe

def _upd_raw_base():
    return "https://raw.githubusercontent.com/%s/%s/%s/" % (GH_OWNER_UPD, GH_REPO_UPD, GH_BRANCH_UPD)

def _upd_ssl():
    import ssl
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        try:
            return ssl.create_default_context()
        except Exception:
            return None

def _upd_vt(s):
    try:
        return tuple(int(x) for x in str(s).strip().split("."))
    except Exception:
        return (0,)

def _upd_get(url, timeout):
    import urllib.request, time as _t
    u = url + ("&" if "?" in url else "?") + "_=" + str(int(_t.time()))
    req = urllib.request.Request(u, headers={"Cache-Control": "no-cache", "User-Agent": "BilanODF"})
    with urllib.request.urlopen(req, timeout=timeout, context=_upd_ssl()) as r:
        return r.read()

def _upd_check_bg():
    """Verifie en tache de fond si une version plus recente existe."""
    global UPDATE_INFO
    if getattr(sys, "frozen", False):
        return
    try:
        import json
        man = json.loads(_upd_get(_upd_raw_base() + "version.json", 5).decode("utf-8"))
        remote = str(man.get("version", "")).strip()
        if remote and _upd_vt(remote) > _upd_vt(APP_VERSION):
            UPDATE_INFO = {"version": remote, "notes": str(man.get("notes", "") or ""), "man": man}
    except Exception:
        pass

def _upd_apply(man):
    """Telecharge tout dans un dossier temporaire PUIS remplace en place (sauvegarde .bak_upd)."""
    import tempfile, shutil
    base = man.get("base") or _upd_raw_base()
    files = man.get("files") or []
    if not files:
        raise RuntimeError("manifest sans fichiers")
    tmp = tempfile.mkdtemp(prefix="bilanodf_upd_")
    try:
        got = []
        for rel in files:
            data = _upd_get(base + rel, 60)
            if not data:
                raise RuntimeError("fichier vide: %s" % rel)
            p = os.path.join(tmp, rel)
            os.makedirs(os.path.dirname(p) or ".", exist_ok=True)
            with open(p, "wb") as f:
                f.write(data)
            got.append(rel)
        for rel in got:
            dst = os.path.join(HERE, rel)
            os.makedirs(os.path.dirname(dst) or ".", exist_ok=True)
            try:
                if os.path.exists(dst):
                    shutil.copy2(dst, dst + ".bak_upd")
            except Exception:
                pass
            shutil.move(os.path.join(tmp, rel), dst)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

@app.route("/maj/state")
def maj_state():
    if not logged():
        return jsonify({})
    if not UPDATE_INFO:
        return jsonify({})
    return jsonify({"version": UPDATE_INFO["version"], "notes": UPDATE_INFO["notes"], "current": APP_VERSION})

@app.route("/maj/apply", methods=["POST"])
def maj_apply():
    if not logged():
        return ("", 403)
    if not UPDATE_INFO:
        return ("aucune mise a jour", 200)
    try:
        _upd_apply(UPDATE_INFO["man"])
    except Exception as e:
        return ("erreur: %s" % e, 200)

    def _relaunch():
        import time as _t
        _t.sleep(0.7)
        try:
            os.execv(sys.executable, [sys.executable] + sys.argv)
        except Exception:
            os._exit(0)
    import threading as _th
    _th.Thread(target=_relaunch, daemon=True).start()
    return ("ok", 200)

# Verifie la disponibilite d'une MAJ au demarrage (thread, non bloquant)
try:
    import threading as _upd_th
    _upd_th.Thread(target=_upd_check_bg, daemon=True).start()
except Exception:
    pass

# ======================================================================
#  IMPORTATION MASSIVE (dossiers d'une clé USB / disque, en local)
# ======================================================================
MASS_STAGE = os.path.join(DATA, "_mass_import.json")

def _name_key(nom):
    return re.sub(r"[^a-z0-9]", "", unicodedata.normalize("NFKD", (nom or "")).encode("ascii", "ignore").decode().lower())

def _existing_name_keys():
    return {_name_key(m.get("nom", "")) for m in list_patients()}

def _browse_roots():
    roots = []
    for v in sorted(_ls_dir("/Volumes")):
        p = os.path.join("/Volumes", v)
        if os.path.isdir(p): roots.append(p)
    home = os.path.expanduser("~")
    for sub in ("", "Desktop", "Documents", "Downloads"):
        p = os.path.join(home, sub) if sub else home
        if os.path.isdir(p): roots.append(p)
    return roots

def _ls_dir(path):
    try: return [e for e in sorted(os.listdir(path)) if not e.startswith(".")]
    except Exception: return []

@app.route("/import_massif")
def import_massif():
    if not logged(): return redirect(url_for("login"))
    path = request.args.get("path", "").strip()
    # écran d'accueil : choix d'un emplacement
    if not path:
        items = ""
        for r in _browse_roots():
            items += ('<a class=mi href="%s" style="border:1px solid var(--line);border-radius:9px;margin-bottom:8px">'
                      '&#128193; <b>%s</b></a>') % (url_for("import_massif", path=r), r)
        return page("""<p class=muted><a href="%s">← Bibliothèque</a></p>
        <h1>Importation massive</h1>
        <div class=flash style="background:var(--accw);border-color:var(--line)">Choisis le <b>dossier qui contient tes dossiers patients</b>
        (ta clé USB, un dossier « BILAN PERSO »…). Tout se passe <b>en local</b> : rien n'est envoyé sur internet.</div>
        <div class=card><h2 style="margin-top:0">Emplacements</h2>%s
        <form method=get style="margin-top:10px;display:flex;gap:8px"><input name=path placeholder="…ou colle un chemin de dossier ici">
        <button class="btn sec">Ouvrir</button></form></div>""" % (url_for("dashboard"), items))
    if not os.path.isdir(path):
        flash("Dossier introuvable : %s" % path); return redirect(url_for("import_massif"))
    # navigateur : sous-dossiers + bouton « analyser »
    subdirs = [e for e in _ls_dir(path) if os.path.isdir(os.path.join(path, e))]
    parent = os.path.dirname(path.rstrip("/")) or "/"
    listing = ('<a class=mi href="%s" style="border:1px solid var(--line);border-radius:9px;margin-bottom:6px">&#8617; Dossier parent</a>'
               % url_for("import_massif", path=parent))
    for e in subdirs[:400]:
        listing += ('<a class=mi href="%s" style="border:1px solid var(--line);border-radius:9px;margin-bottom:6px">&#128193; %s</a>'
                    ) % (url_for("import_massif", path=os.path.join(path, e)), e)
    return page("""<p class=muted><a href="%s">← Emplacements</a></p>
    <h1>Importation massive</h1>
    <div class=card><div class=muted style="word-break:break-all;margin-bottom:6px">%s</div>
    <div style="margin:10px 0"><form method=post action="%s" style="display:inline">
      <input type=hidden name=path value="%s">
      <button class=btn onclick="this.innerHTML='<span class=spin></span> Analyse…'">&#128269; Analyser ce dossier (%d sous-dossiers)</button>
    </form></div>
    <h2>Entrer dans un sous-dossier</h2>%s</div>""" % (
        url_for("import_massif"), path, url_for("import_massif_scan"), path.replace('"', "&quot;"),
        len(subdirs), listing))

@app.route("/import_massif/scan", methods=["POST"])
def import_massif_scan():
    if not logged(): return redirect(url_for("login"))
    path = request.form.get("path", "").strip()
    if not os.path.isdir(path):
        flash("Dossier introuvable."); return redirect(url_for("import_massif"))
    plans = MI.scan_source(path)
    existing = _existing_name_keys()
    seen = set()
    data = {"source": path, "patients": []}
    for pl in plans:
        s = MI.plan_summary(pl)
        nom = pl["ident"]["nom"]
        # identité fiable = celle du bilan Word quand il existe (évite les doublons de nom)
        if pl.get("word_path") and os.path.exists(pl["word_path"]):
            try:
                pn = (WI.parse_word_bilan(pl["word_path"]).get("patient") or {}).get("nom")
                if pn: nom = pn
            except Exception:
                pass
        key = _name_key(nom)
        dup_in_batch = key in seen
        seen.add(key)
        data["patients"].append({
            "folder": pl["folder"], "nom": nom, "nom_famille": pl["ident"]["nom_famille"],
            "prenom": pl["ident"]["prenom"], "mode": pl["mode"], "summary": s,
            "warnings": pl["warnings"], "exists": (key in existing) or dup_in_batch})
    json.dump(data, open(MASS_STAGE, "w"), ensure_ascii=False)
    return redirect(url_for("import_massif_recap"))

@app.route("/import_massif/recap")
def import_massif_recap():
    if not logged(): return redirect(url_for("login"))
    if not os.path.exists(MASS_STAGE):
        return redirect(url_for("import_massif"))
    data = json.load(open(MASS_STAGE))
    pts = data["patients"]
    n_new = sum(1 for p in pts if not p["exists"] and p["mode"] != "empty")
    rows = ""
    for i, p in enumerate(pts):
        s = p["summary"]
        if p["mode"] == "word_only":
            content = "Word seul · %d image(s) à extraire" % s.get("word_imgs", 0)
        elif p["mode"] == "empty":
            content = "<span style='color:#C0392B'>vide</span>"
        else:
            bits = []
            if s["temps"]: bits.append("%d temps · %d photos" % (s["temps"], s["photos"]))
            if s["radios"]: bits.append("%d radio(s)" % s["radios"])
            if s["stl"]: bits.append("%d STL" % s["stl"])
            if s["renders"]: bits.append("%d rendu(s)" % s["renders"])
            content = " · ".join(bits) or "—"
        warn = (" <span style='color:#b45309'>⚠ %s</span>" % "; ".join(p["warnings"])) if p["warnings"] else ""
        if p["exists"]:
            state = "<span style='color:#64748b'>déjà présent — ignoré</span>"; chk = ""
        elif p["mode"] == "empty":
            state = "à ignorer"; chk = ""
        else:
            state = "<b style='color:#16a34a'>nouveau</b>"
            chk = "<input type=checkbox name=sel value='%d' checked>" % i
        rows += ("<tr><td style='text-align:center'>%s</td><td><b>%s</b></td><td>%s</td>"
                 "<td>%s%s</td><td>%s</td></tr>") % (
            chk, p["nom"], {"files": "Fichiers", "word_only": "Word seul", "empty": "—"}.get(p["mode"], p["mode"]),
            content, warn, state)
    return page("""<p class=muted><a href="%s">← Changer de dossier</a></p>
    <h1>Récapitulatif de l'import</h1>
    <div class=flash style="background:rgba(34,197,94,.12);border-color:rgba(34,197,94,.35)">Source : <b>%s</b><br>
    <b>%d patient(s)</b> analysés · <b>%d nouveau(x)</b> à créer. Les patients déjà présents sont ignorés (aucun écrasement).
    Décoche ceux que tu ne veux pas importer, puis valide.</div>
    <form method=post action="%s">
    <div class=card style="overflow:auto"><table class=mtab>
    <thead><tr><th></th><th>Patient</th><th>Type</th><th>Contenu détecté</th><th>État</th></tr></thead>
    <tbody>%s</tbody></table></div>
    <div style="margin-top:14px;display:flex;gap:10px"><button class=btn onclick="this.innerHTML='<span class=spin></span> Import en cours… (ne ferme pas)'">Importer les patients cochés</button>
    <a class="btn sec" href="%s">Annuler</a></div></form>""" % (
        url_for("import_massif"), data["source"].replace('<', '&lt;'), len(pts), n_new,
        url_for("import_massif_commit"), rows, url_for("dashboard")))

def _mass_process_record(base, r):
    """Traitement léger d'un enregistrement importé : figure Steiner uniquement.
    (Les photos du Word sont placées TELLES QUELLES — pas de recadrage carré — donc on ne
    lance pas process_photos ; génération Word/PDF et rendu 3D évités car trop lourds en masse.)"""
    try:
        st = r.get("steiner", {})
        if st:
            sig, sev = {}, {}
            for nom, v in st.items():
                mean, sd = next(((m, s) for (n, _l, m, s) in STEINER_FIELDS if n == nom), (0, 1))
                t, x = steiner_signif(nom, v, mean, sd); sig[nom] = t; sev[nom] = x
            P.make_steiner_json(base, st, sig, sev); P.build_steiner_figure(base)
    except Exception:
        pass

def _save_photo(src, base, view):
    from PIL import Image as _I, ImageOps as _IO
    try:
        _IO.exif_transpose(_I.open(src)).convert("RGB").save(
            os.path.join(base, "01_photos_brutes", view + ".jpg"), quality=92)
        return True
    except Exception:
        return False

def _radio_key(src):
    """panoramique (large) vs téléradiographie de profil, d'après le nom puis le ratio."""
    n = MI._norm(os.path.basename(src))
    if "pano" in n or "opt" in n: return "panoramique"
    if "trp" in n or "tele" in n or "teleradio" in n or "profil" in n or "ceph" in n: return "teleradiographie_profil"
    try:
        from PIL import Image as _I
        w, h = _I.open(src).size
        return "panoramique" if w / max(1, h) > 1.5 else "teleradiographie_profil"
    except Exception:
        return "teleradiographie_profil"

def _save_radio(src, base):
    from PIL import Image as _I, ImageOps as _IO
    try:
        _IO.exif_transpose(_I.open(src)).convert("RGB").save(
            os.path.join(base, "03_radios", _radio_key(src) + ".jpg"), quality=92)
        return True
    except Exception:
        return False

def _fast_open(path, box=1400):
    """Ouvre une image en décodant à résolution réduite (JPEG DCT) — beaucoup plus rapide
    sur les gros clichés de 5–6 Mo."""
    from PIL import Image as _I, ImageOps as _IO, ImageFile as _IF
    _IF.LOAD_TRUNCATED_IMAGES = True   # certaines photos de la clé sont légèrement tronquées
    im = _I.open(path)
    try: im.draft("RGB", (box, box))
    except Exception: pass
    return _IO.exif_transpose(im).convert("RGB")

def _img_feats(path):
    """Caractéristiques robustes d'une image (pour trier visages / intra-oral / radio / rendu)."""
    import numpy as _np
    im = _fast_open(path, 240); w, h = im.size
    a = _np.asarray(im.resize((100, 100))).astype(float) / 255.0
    R, G, B = a[..., 0], a[..., 1], a[..., 2]
    mx = a.max(-1); mn = a.min(-1); sat = (mx - mn) / _np.clip(mx, 1e-6, 1)
    muc = float(((R > 0.35) & ((R - G) > 0.10) & ((R - B) > 0.10) & (sat > 0.20)).mean())  # muqueuse rose/rouge
    grayfrac = float((sat > 0.15).mean())
    cor = _np.concatenate([a[:12, :12].reshape(-1, 3), a[:12, -12:].reshape(-1, 3),
                           a[-12:, :12].reshape(-1, 3), a[-12:, -12:].reshape(-1, 3)])
    purple = bool((cor[:, 2].mean() - cor[:, 0].mean()) > 0.03 and cor.std(0).mean() < 0.18)
    g = a.mean(-1); asym = float(_np.abs(g - g[:, ::-1]).mean())
    teeth = float(((mx > 0.6) & (sat < 0.18)).mean())
    return {"aspect": w / max(1, h), "gray": grayfrac < 0.06, "muc": muc,
            "purple": purple, "asym": asym, "teeth": teeth}

def _clin_feats(path):
    """Descripteurs pour le classement automatique des vues (visage/intra/occlusal/latérale)."""
    import numpy as _np
    im = _fast_open(path, 200); w, h = im.size
    a = _np.asarray(im.resize((120, 120))).astype(float) / 255.0
    R, G, B = a[..., 0], a[..., 1], a[..., 2]
    mx = a.max(-1); mn = a.min(-1); sat = (mx - mn) / _np.clip(mx, 1e-6, 1)
    muc = ((R > 0.35) & ((R - G) > 0.12) & ((R - B) > 0.10) & (sat > 0.22))
    # peau : tons chauds clairs, PAS la muqueuse rouge saturée (permet de repérer les visages
    # quelle que soit l'orientation — les intra-orales n'ont quasi pas de peau)
    skin = ((R > 0.35) & (R >= G) & (G >= B) & ((R - B) > 0.03) & ((R - B) < 0.42) & (sat < 0.40) & ~muc)
    teeth = ((mx > 0.55) & (sat < 0.20)); g = a.mean(-1); c = slice(40, 80)
    return {"p": path, "aspect": w / max(1, h), "grayfrac": float((sat > 0.15).mean()),
            "muc": float(muc.mean()), "skin": float(skin.mean()), "asym": float(_np.abs(g - g[:, ::-1]).mean()),
            "cmuc": float(muc[c, c].mean()), "csat": float(sat[c, c].mean()), "cR": float(R[c, c].mean()),
            "lft": float(teeth[:, :60].mean()), "rgt": float(teeth[:, 60:].mean()),
            "midteeth": float(teeth[45:80, 30:90].mean())}

def _view_feat(path):
    """Vecteur RÉGIONAL (grille 3x3 sur muqueuse/dents/peau/luminance/saturation) : capture la
    DISPOSITION spatiale (palais au centre = occlusal, dents d'un côté = latérale, peau étalée =
    visage…). Sert au classement APPRIS à partir des fiches déjà classées par l'utilisateur."""
    import numpy as _np
    im = _fast_open(path, 120)
    a = _np.asarray(im.resize((96, 96))).astype(float) / 255.0
    R, G, B = a[..., 0], a[..., 1], a[..., 2]
    mx = a.max(-1); mn = a.min(-1); sat = (mx - mn) / _np.clip(mx, 1e-6, 1)
    muc = ((R > 0.35) & ((R - G) > 0.12) & ((R - B) > 0.10) & (sat > 0.22)).astype(float)
    teeth = ((mx > 0.55) & (sat < 0.20)).astype(float)
    skin = ((R > 0.30) & (R >= G) & (G >= B) & ((R - B) > 0.02) & ((R - B) < 0.45) & (sat < 0.42) & (muc < 0.5)).astype(float)
    gray = a.mean(-1)
    vec = [im.size[0] / max(1, im.size[1]), float(_np.abs(gray - gray[:, ::-1]).mean())]
    for M in (muc, teeth, skin, gray, sat):
        for i in range(3):
            for j in range(3):
                vec.append(float(M[i * 32:(i + 1) * 32, j * 32:(j + 1) * 32].mean()))
    return _np.array(vec, dtype=float)

_VIEW_CEN = None   # (centroids[8], mu, sd) appris ; False si pas assez d'exemples
def _get_view_centroids():
    """Construit (une fois) les centroïdes de chaque vue à partir des photos DÉJÀ classées
    par l'utilisateur (01_photos_brutes/<vue>.jpg de toutes ses fiches). Renvoie None si
    l'utilisateur n'a pas encore assez d'exemples (-> repli sur les règles)."""
    global _VIEW_CEN
    if _VIEW_CEN is not None:
        return _VIEW_CEN or None
    import numpy as _np
    views = [k for k, _ in PHOTO_FIELDS]; vidx = {v: i for i, v in enumerate(views)}
    feats = [[] for _ in range(8)]
    try:
        for slug in os.listdir(PATIENTS):
            rd = os.path.join(PATIENTS, slug, "records")
            if not os.path.isdir(rd): continue
            for rid in os.listdir(rd):
                braw = os.path.join(rd, rid, "01_photos_brutes")
                if not os.path.isdir(braw): continue
                # N'apprendre QUE sur les fiches validees par l'utilisateur
                # (sinon le systeme apprend ses propres erreurs de classement auto).
                try:
                    if not json.load(open(os.path.join(rd, rid, "meta.json"), encoding="utf-8")).get("photos_verified"):
                        continue
                except Exception:
                    continue
                for v in views:
                    p = os.path.join(braw, v + ".jpg")
                    if os.path.exists(p):
                        try: feats[vidx[v]].append(_view_feat(p))
                        except Exception: pass
    except Exception:
        _VIEW_CEN = False; return None
    if any(len(feats[c]) < 4 for c in range(8)):   # pas assez d'exemples sur au moins une vue
        _VIEW_CEN = False; return None
    allf = _np.array([f for lst in feats for f in lst])
    mu = allf.mean(0); sd = allf.std(0) + 1e-6
    cent = _np.array([((_np.array(feats[c]) - mu) / sd).mean(0) for c in range(8)])
    _VIEW_CEN = (cent, mu, sd)
    return _VIEW_CEN

def _classify_learned(paths, cen):
    """Classe via les centroïdes appris + affectation 1-à-1, puis réordonne les visages par règle.
    Renvoie (photos{vue:chemin}, radios[], renders[]) ou None si non applicable."""
    import numpy as _np
    cent, mu, sd = cen
    views = [k for k, _ in PHOTO_FIELDS]; vidx = {v: i for i, v in enumerate(views)}
    RE, SO, PR = vidx["exo_face_repos"], vidx["exo_face_sourire"], vidx["exo_profil"]
    aux = {p: _clin_feats(p) for p in paths}
    radios = [p for p in paths if aux[p]["muc"] < 0.05 and aux[p]["grayfrac"] < 0.12]
    clin = [p for p in paths if p not in radios]
    if not clin: return None
    vecs = {p: (_view_feat(p) - mu) / sd for p in clin}
    D = {(p, c): float(_np.linalg.norm(vecs[p] - cent[c])) for p in clin for c in range(8)}
    # Barriere anti-confusion visage/intra-oral, robuste a la couleur de peau :
    # une photo SANS muqueuse (muc faible) est un visage -> seulement les 3 cases visage ;
    # une photo AVEC muqueuse est intra-orale -> seulement les 5 cases intra-orales.
    FACE_SLOTS = {RE, SO, PR}; INTRA_SLOTS = set(range(8)) - FACE_SLOTS
    def _allowed(p):
        return FACE_SLOTS if aux[p]["muc"] < 0.25 else INTRA_SLOTS
    def _best(p):
        al = [c for c in _allowed(p)]
        return min(D[(p, c)] for c in al) if al else 1e9
    face_cands = [p for p in clin if aux[p]["muc"] < 0.25]
    intra_cands = [p for p in clin if aux[p]["muc"] >= 0.25]
    used = set(); pred = {}
    # INTRA-ORAL : affectation apprise 1-a-1 (les distinctions fines profitent des exemples).
    for p in sorted(intra_cands, key=lambda p: min(D[(p, c)] for c in INTRA_SLOTS)):
        free = [c for c in INTRA_SLOTS if c not in used]
        if not free: break
        c = min(free, key=lambda c: D[(p, c)]); pred[p] = c; used.add(c)
    # VISAGES : regle fiable et independante de la peau. profil = le plus asymetrique ;
    # repos/sourire = les 2 photos les plus FRONTALES (asym faible), jamais un 2e profil.
    if face_cands:
        pf = max(face_cands, key=lambda p: aux[p]["asym"] + abs(aux[p]["lft"] - aux[p]["rgt"]))
        pred[pf] = PR
        frontals = sorted([p for p in face_cands if p != pf], key=lambda p: aux[p]["asym"])[:2]
        fr = sorted(frontals, key=lambda p: aux[p]["midteeth"], reverse=True)
        if len(fr) >= 1: pred[fr[0]] = SO
        if len(fr) >= 2: pred[fr[1]] = RE
    photos = {views[c]: p for p, c in pred.items()}
    radlist = sorted(radios, key=lambda p: aux[p]["aspect"])[:2]
    return photos, radlist, []

def _classify_photo_set(paths):
    """Classe un JEU PROPRE de photos de bilan vers les 8 vues + radios.
    D'abord par APPRENTISSAGE (centroïdes des fiches déjà classées par l'utilisateur) ; à défaut,
    par règles relatives. Renvoie (photos{vue:chemin}, radios[chemins], autres[rendus])."""
    cen = _get_view_centroids()
    if cen:
        try:
            res = _classify_learned(paths, cen)
            if res is not None: return res
        except Exception:
            pass
    F = [_clin_feats(p) for p in paths]
    # RADIOS = quasi pas de muqueuse et peu coloré (niveaux de gris).
    radios = [x for x in F if x["muc"] < 0.05 and x["grayfrac"] < 0.12]
    nonrad = [x for x in F if x not in radios]
    # VISAGES = pas de muqueuse au CENTRE (les occlusales montrent le palais/gencive au centre)
    # ET présence de peau. INDÉPENDANT de l'orientation (portrait ou paysage) : c'est ce qui
    # faisait échouer le classement quand toutes les photos étaient en paysage.
    # VISAGE = aucune muqueuse au centre (les occlusales montrent le palais/gencive au centre).
    # On NE filtre PAS sur la peau (ça rejetait à tort des visages peu lumineux / peau foncée).
    # VISAGE = photo quasiment sans muqueuse (muc faible). Separateur robuste a la couleur de
    # peau : l'ancien filtre exigeait cmuc<0.15, ce qui rejetait les visages dont les levres
    # (au centre) sont vues comme muqueuse -> patients a peau foncee mal classes.
    face_cand = [x for x in nonrad if x["muc"] < 0.25]
    intra = [x for x in nonrad if x["muc"] >= 0.25]
    faces = sorted(face_cand, key=lambda x: x["muc"])[:3]   # les 3 avec le moins de muqueuse
    renders = []
    photos = {}
    rs = sorted(radios, key=lambda x: x["aspect"])
    if rs: photos["__radio1"] = rs[0]["p"]
    if len(rs) >= 2: photos["__radio2"] = rs[-1]["p"]
    if face_cand:
        # profil = le plus asymetrique ; repos/sourire = les 2 photos les plus FRONTALES
        # (asym la plus faible) -> jamais un 2e profil dans une case de face.
        pf = max(face_cand, key=lambda x: x["asym"] + abs(x["lft"] - x["rgt"])); photos["exo_profil"] = pf["p"]
        frontals = sorted([x for x in face_cand if x is not pf], key=lambda x: x["asym"])[:2]
        fr = sorted(frontals, key=lambda x: x["midteeth"], reverse=True)
        if len(fr) >= 1: photos["exo_face_sourire"] = fr[0]["p"]        # sourire = plus de dents
        if len(fr) >= 2: photos["exo_face_repos"] = fr[1]["p"]
    occ = sorted(intra, key=lambda x: x["cmuc"], reverse=True)[:2]
    rest3 = [x for x in intra if x not in occ]
    if len(occ) >= 2:
        o = sorted(occ, key=lambda x: x["cR"] - 0.6 * x["csat"], reverse=True)  # occmax = palais (plus clair)
        photos["endo_occlusal_maxillaire"] = o[0]["p"]; photos["endo_occlusal_mandibulaire"] = o[1]["p"]
    elif occ:
        photos["endo_occlusal_maxillaire"] = occ[0]["p"]
    if rest3:
        fr = min(rest3, key=lambda x: x["asym"]); photos["endo_occlusion_frontale"] = fr["p"]
        for x in [y for y in rest3 if y is not fr]:
            k = "endo_laterale_droite" if x["lft"] > x["rgt"] else "endo_laterale_gauche"
            if k in photos: k = "endo_laterale_gauche" if k == "endo_laterale_droite" else "endo_laterale_droite"
            photos[k] = x["p"]
    rad_paths = [photos.pop(k) for k in ("__radio1", "__radio2") if k in photos]
    return photos, rad_paths, [x["p"] for x in renders]

def _to_gallery(base, paths):
    """Copie toutes les photos dans 08_galerie (rien n'est perdu ; base de l'assignation manuelle).
    Réduites à ~1600 px pour rester rapides et légères (largement suffisant pour le bilan)."""
    gal = os.path.join(base, "08_galerie"); os.makedirs(gal, exist_ok=True)
    existing = len([f for f in os.listdir(gal) if f.lower().endswith((".jpg", ".jpeg", ".png"))])
    for i, src in enumerate(paths):
        try:
            im = _fast_open(src, 1600); im.thumbnail((1600, 1600))
            im.save(os.path.join(gal, "g%03d.jpg" % (existing + i)), quality=85)
        except Exception: pass

def _auto_faces(clinical_feats):
    """Attribue les 3 vues de visage (fiable) : profil = plus asymétrique ; sourire = plus de dents."""
    res = {}
    faces = [(p, f) for p, f in clinical_feats if f["aspect"] < 1.2]
    if not faces: return res
    faces.sort(key=lambda pf: pf[1]["asym"], reverse=True)
    res["exo_profil"] = faces[0][0]
    rest = sorted(faces[1:], key=lambda pf: pf[1]["teeth"], reverse=True)
    if len(rest) >= 1: res["exo_face_sourire"] = rest[0][0]
    if len(rest) >= 2: res["exo_face_repos"] = rest[1][0]
    return res

def _place_as_is(src, base, view, box=1600):
    """Place une photo déjà cadrée (issue du Word) TELLE QUELLE dans la vue, sans recadrage :
    on écrit à la fois la brute et la version 'traitée' (affichée) pour éviter le cadre carré."""
    from PIL import Image as _I
    try:
        im = _fast_open(src, box); im.thumbnail((box, box))
        im.save(os.path.join(base, "01_photos_brutes", view + ".jpg"), quality=92)
        im.save(os.path.join(base, "02_photos_traitees", view + ".jpg"), quality=92)
        return True
    except Exception:
        return False

def _render_slot(base, slot, c):
    """Rend UNE seule photo (orientation + recadrage) depuis la brute vers 02_photos_traitees,
    SANS régénérer le Word ni les rendus 3D. Rapide -> la rotation/recadrage est instantané."""
    from PIL import Image as _I, ImageOps as _IO, ImageFile as _IF
    _IF.LOAD_TRUNCATED_IMAGES = True
    src = os.path.join(base, "01_photos_brutes", slot + ".jpg")
    if not os.path.exists(src): return False
    try:
        im = _IO.exif_transpose(_I.open(src)).convert("RGB")
        ang = int(c.get("rot", 0) or 0) + float(c.get("fine", 0) or 0)
        if ang: im = im.rotate(ang, expand=True, resample=_I.BICUBIC)
        if c.get("mirror"): im = _IO.mirror(im)
        if c.get("flip"): im = _IO.flip(im)
        W, H = im.size
        x, y, w, h = (float(c.get("x", 0)), float(c.get("y", 0)),
                      float(c.get("w", 1)), float(c.get("h", 1)))
        if (x, y, w, h) != (0.0, 0.0, 1.0, 1.0):
            im = im.crop((int(x * W), int(y * H), int((x + w) * W), int((y + h) * H)))
        im.thumbnail((1600, 1600), _I.LANCZOS)
        os.makedirs(os.path.join(base, "02_photos_traitees"), exist_ok=True)
        im.save(os.path.join(base, "02_photos_traitees", slot + ".jpg"), quality=92)
        return True
    except Exception:
        return False

def _place_word_photos(base, docx_path):
    """Depuis un bilan Word : classe automatiquement les 8 vues (classement relatif, fiable),
    les place TELLES QUELLES (pas de recadrage carré), place les radios, et met toutes les
    photos cliniques dans la galerie pour corriger un éventuel échange en un glisser."""
    from PIL import Image as _I
    import tempfile
    counts = {"photos": 0, "radios": 0, "captures": 0}
    tmp = tempfile.mkdtemp()
    try:
        try:
            imgs = MI.extract_word_media(docx_path, tmp)
        except Exception:
            return False, counts
        if not imgs:
            return False, counts
        photos, radios, renders = _classify_photo_set(imgs)
        for view, src in photos.items():
            if _place_as_is(src, base, view): counts["photos"] += 1
        for src in radios:
            if _save_radio(src, base): counts["radios"] += 1
        if renders:
            cap = os.path.join(base, "07_modele_captures"); os.makedirs(cap, exist_ok=True)
            for i, src in enumerate(renders):
                try:
                    _I.open(src).convert("RGB").save(os.path.join(cap, "capture_%02d.jpg" % i), quality=90)
                    counts["captures"] += 1
                except Exception: pass
        # galerie = toutes les photos cliniques (visages + intra-orales) pour corriger au besoin
        clin = [photos[v] for v in photos]
        _to_gallery(base, clin)
        return True, counts
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

def _place_folder_photos(base, photo_paths):
    """Photos BRUTES d'un dossier (réévaluation, ou bilan sans Word) : classe automatiquement
    les vues (mêmes règles que le Word), place TELLES QUELLES (pas de recadrage carré), et met
    toutes les photos dans la galerie pour corriger au besoin."""
    if not photo_paths:
        return 0
    n = 0
    try:
        photos, radios, _renders = _classify_photo_set(photo_paths)
        for view, src in photos.items():
            if _place_as_is(src, base, view): n += 1
        for src in radios:
            _save_radio(src, base)
    except Exception:
        pass
    _to_gallery(base, photo_paths)
    return n

def _place_raw_photos(base, photo_paths):
    """Dossier brut : on met TOUTES les photos dans la galerie (assignation manuelle par
    glisser-déposer) et on ne place automatiquement que les radios (niveaux de gris, fiable).
    Aucune photo clinique n'est pré-placée -> plus jamais de photo dans la mauvaise vue."""
    if not photo_paths:
        return 0
    clinical = []
    for src in photo_paths:
        try:
            if _img_feats(src)["gray"]:
                _save_radio(src, base)               # radio glissée dans le dossier photo
            else:
                clinical.append(src)
        except Exception:
            clinical.append(src)
    _to_gallery(base, clinical)
    return 0

def _capture_date(paths):
    """Date d'un temps = date de fichier médiane (rapide, pas de décodage d'image)."""
    dates = []
    for p in paths[:60]:
        try: dates.append(datetime.date.fromtimestamp(os.path.getmtime(p)).isoformat())
        except Exception: pass
    if not dates: return None
    return sorted(dates)[len(dates) // 2]

@app.route("/import_massif/commit", methods=["POST"])
def import_massif_commit():
    if not logged(): return redirect(url_for("login"))
    if not os.path.exists(MASS_STAGE): return redirect(url_for("import_massif"))
    data = json.load(open(MASS_STAGE))
    sel = set(request.form.getlist("sel"))
    existing = _existing_name_keys()
    created = 0; skipped = 0; warns = []
    for i, p in enumerate(data["patients"]):
        if str(i) not in sel: continue
        if p["mode"] == "empty" or _name_key(p["nom"]) in existing:
            skipped += 1; continue
        folder = p["folder"]
        plan = MI.scan_patient(folder) if os.path.isdir(folder) else MI._word_only_plan(folder)
        try:
            n = _do_import_patient(p, plan); created += n
            existing.add(_name_key(p["nom"]))
        except Exception as e:
            warns.append("%s : %s" % (p["nom"], e))
    try: os.remove(MASS_STAGE)
    except Exception: pass
    msg = "Import terminé : %d patient(s) créé(s), %d ignoré(s)." % (created, skipped)
    if warns: msg += " Soucis : " + " | ".join(warns[:5])
    flash(msg)
    return redirect(url_for("dashboard"))

def _do_import_patient(meta, plan):
    """Crée le patient + ses enregistrements et copie/classe les fichiers. Renvoie 1 si créé."""
    from PIL import Image as _I, ImageOps as _IO
    nom = meta["nom"]
    # identité depuis le Word si disponible
    dob = ""; sexe = "non précisé"; age = "—"; word_res = None
    if plan.get("word_path") and os.path.exists(plan["word_path"]):
        try:
            word_res = WI.parse_word_bilan(plan["word_path"])
            pat = word_res.get("patient", {})
            if pat.get("nom"): nom = pat["nom"]
            dob = pat.get("dob", "") or ""
            if pat.get("sexe"): sexe = "masculin" if str(pat["sexe"]).lower().startswith(("h", "m")) else "féminin"
        except Exception:
            word_res = None
    dob_court, age = age_from_dob(dob) if dob else ("—", "—")
    slug = slugify(nom) + "_" + datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")[:-3]
    reeval = plan["mode"] == "files" and len(plan["timepoints"]) > 1
    pt = {"slug": slug, "nom": nom, "dob": dob, "dob_court": dob_court, "age": age, "sexe": sexe,
          "auteur": session.get("user", ""), "date_creation": datetime.datetime.now().isoformat(timespec="seconds"),
          "statut": "en_cours" if reeval else "bilan", "records": []}
    os.makedirs(pdir(slug), exist_ok=True); save_patient(slug, pt)

    def apply_word(r):
        if word_res:
            try: apply_word_to_record(r, word_res)
            except Exception: pass

    word_ok = bool(plan.get("word_path") and os.path.exists(plan["word_path"]))

    if plan["mode"] == "word_only":
        rid = create_empty_record(slug, "Bilan initial")
        base = rdir(slug, rid); r = load_rec(slug, rid)
        _place_word_photos(base, plan["word_path"])
        apply_word(r)
        r["status"] = "done"; r["import_note"] = "Extrait du bilan Word"
        save_rec(slug, rid, r)
        _mass_process_record(base, r)
        return 1

    # --- mode fichiers : un enregistrement par temps (regroupé par DATE) ---
    timepoints = plan["timepoints"] or [{"label": "Bilan initial", "dir": plan["folder"], "photos": []}]
    for ti, tp in enumerate(timepoints):
        rid = create_empty_record(slug, tp["label"]); base = rdir(slug, rid); r = load_rec(slug, rid)
        d = tp.get("date") or _capture_date(tp.get("photos") or [])
        if d:
            r["date"] = d[:10] + "T09:00:00"
        # PHOTOS — on privilégie TOUJOURS les photos BRUTES du dossier (pleines, non rognées)
        # et on les classe automatiquement. Le Word (photos déjà rognées) ne sert que de repli
        # quand le dossier ne contient aucune photo (cas « Word seul »).
        if tp.get("photos"):
            _place_folder_photos(base, tp["photos"])   # brutes -> classement auto + galerie
        elif ti == 0 and word_ok:
            _place_word_photos(base, plan["word_path"])
        # RADIOS / STL / RENDUS de CETTE séance (rattachés par date au scan)
        for src in tp.get("radios", []):
            _save_radio(src, base)
        for src in tp.get("stl", []):
            low = os.path.basename(src).lower()
            if "bite" in low: continue
            dest = "LowerJawScan.stl" if any(w in low for w in ("lower", "mand", "inf", "bas")) else "UpperJawScan.stl"
            try: shutil.copy(src, os.path.join(base, "04_stl_bruts", dest))
            except Exception: pass
        for rk, src in tp.get("renders", []):
            if not rk: continue
            try: _I.open(src).convert("RGB").save(os.path.join(base, "05_stl_rendus", rk + ".png"))
            except Exception: pass
        if ti == 0:
            apply_word(r)
        r["status"] = "done"
        save_rec(slug, rid, r)
        _mass_process_record(base, r)
    return 1

_MOIS_FR = ["", "janvier", "février", "mars", "avril", "mai", "juin",
            "juillet", "août", "septembre", "octobre", "novembre", "décembre"]
def _date_fr(date_iso):
    """'2026-03-12...' -> '12 mars 2026' (chaîne vide si non parsable)."""
    try:
        d = datetime.date.fromisoformat((date_iso or "")[:10])
        return "%d %s %d" % (d.day, _MOIS_FR[d.month], d.year)
    except Exception:
        return ""

def age_at(dob, date_iso):
    try:
        d = datetime.date.fromisoformat((date_iso or "")[:10]); b = datetime.date.fromisoformat(dob)
        return "%d ans" % (d.year - b.year - ((d.month, d.day) < (b.month, b.day)))
    except Exception:
        return ""

# ---- édition de la fiche patient ----
@app.route("/patient/<slug>/edit", methods=["GET", "POST"])
def patient_edit(slug):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug)
    if not pt: abort(404)
    if request.method == "POST":
        nom = (request.form.get("nom", "").strip().upper() + " " + request.form.get("prenom", "").strip()).strip()
        dob = request.form.get("dob", "").strip()
        dc, age = age_from_dob(dob) if dob else (pt.get("dob_court", "—"), pt.get("age", "—"))
        stt = request.form.get("statut", "")
        pt.update({"nom": nom or pt.get("nom", ""), "dob": dob, "dob_court": dc, "age": age,
                   "sexe": request.form.get("sexe", pt.get("sexe", "non précisé")),
                   "statut": stt if stt in STATUT_MAP else statut_of(pt)})
        save_patient(slug, pt); flash("Fiche patient mise à jour.")
        return redirect(url_for("patient_edit", slug=slug))
    parts = (pt.get("nom", "") or "").split(" ", 1)
    nom0 = parts[0] if parts else ""; prenom0 = parts[1] if len(parts) > 1 else ""
    sx = pt.get("sexe", "non précisé")
    opts = "".join('<option %s>%s</option>' % ("selected" if s == sx else "", s) for s in ("non précisé", "féminin", "masculin"))
    cur = statut_of(pt)
    stopts = "".join('<option value="%s" %s>%s</option>' % (k, "selected" if k == cur else "", lbl)
                     for k, lbl, _c, _e in STATUTS)
    return page("""%s
    <form method=post class=card style="max-width:560px"><h2 style="margin:0 0 12px">Informations du patient</h2>
      <div class="grid g2"><div><label>Nom</label><input name=nom value="%s"></div>
      <div><label>Prénom</label><input name=prenom value="%s"></div></div>
      <div class="grid g2" style="margin-top:12px"><div><label>Date de naissance</label><input type=date name=dob value="%s"></div>
      <div><label>Sexe</label><select name=sexe>%s</select></div></div>
      <div class="grid g2" style="margin-top:12px"><div><label>Statut du traitement</label><select name=statut>%s</select></div>
      <div><label>&nbsp;</label><div class=muted style="font-size:13px;padding-top:8px">%s</div></div></div>
      <div style="margin-top:16px"><button class=btn>Enregistrer</button></div></form>
    <div class=card style="max-width:560px"><h3 style="margin:0 0 6px;color:#C0392B">Zone sensible</h3>
      <p class=muted style="margin:0 0 10px">Supprimer ce patient le déplace dans une corbeille (récupérable), il n'est pas effacé définitivement.</p>
      <a class="btn sec" href="%s" style="color:#C0392B">&#128465; Supprimer ce patient</a></div>""" % (
        phead(slug, pt, "fiche"), nom0.replace('"', "&quot;"), prenom0.replace('"', "&quot;"),
        pt.get("dob", ""), opts, stopts, (treatment_duration_txt(slug, pt) or "—"),
        url_for("patient_delete", slug=slug)))

# ---- changement rapide de statut (depuis la bibliothèque ou l'en-tête) ----
@app.route("/patient/<slug>/statut/<key>")
def set_statut(slug, key):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug)
    if not pt: abort(404)
    if key in STATUT_MAP:
        pt["statut"] = key; save_patient(slug, pt)
    return redirect(request.referrer or url_for("dashboard"))

@app.route("/patient/<slug>/staffer")
def toggle_staffer(slug):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug)
    if not pt: abort(404)
    pt["a_staffer"] = not bool(pt.get("a_staffer")); save_patient(slug, pt)
    return redirect(request.referrer or url_for("dashboard"))


@app.route("/patient/<slug>/presenter")
def toggle_presenter(slug):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug)
    if not pt: abort(404)
    pt["a_presenter"] = not bool(pt.get("a_presenter")); save_patient(slug, pt)
    return redirect(request.referrer or url_for("dashboard"))


# ---- suppression d'un patient (corbeille récupérable) ----
@app.route("/patient/<slug>/delete", methods=["GET", "POST"])
def patient_delete(slug):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug)
    if not pt: abort(404)
    if request.method == "POST":
        trash = os.path.join(DATA, "_corbeille_patients"); os.makedirs(trash, exist_ok=True)
        src = pdir(slug)
        if os.path.isdir(src):
            dest = os.path.join(trash, slug)
            if os.path.exists(dest): dest += "_" + secrets.token_hex(3)
            try: shutil.move(src, dest)
            except Exception: pass
        flash("Patient déplacé dans la corbeille.")
        return redirect(url_for("dashboard"))
    nrec = len([1 for rid in pt.get("records", []) if load_rec(slug, rid)])
    return page("""<h1>Supprimer un patient</h1>
    <div class=card><p>Confirmer la suppression de <b>%s</b> (%d enregistrement(s)) ?</p>
    <p class=muted>Le dossier complet sera déplacé dans une corbeille (dans %s), pas effacé définitivement — récupérable manuellement si besoin.</p>
    <form method=post style="display:inline"><button class=btn style="background:#C0392B">Oui, mettre à la corbeille</button></form>
    <a class="btn sec" href="%s">Annuler</a></div>""" % (
        (pt.get("nom", "") or "").replace("<", "&lt;"), nrec,
        os.path.join("~", "BilanODF_Data", "_corbeille_patients"), url_for("dashboard")))

# ======================================================================
#  SUIVI DE TRAITEMENT — journal de séances daté + récap des bilans, état actuel
# ======================================================================
# --- état actuel DÉDUIT des commentaires (Max = maxillaire, Mand = mandibulaire) ---
# appareils : (motif, libellé, arch-based ?) ; arch-based=True => arcs pertinents
_APP_PATTERNS = [
    (r"quad[\s\-]?helix", "Quad Helix", False),
    (r"disjoncteur|hyrax|disjonction", "Disjoncteur", False),
    (r"\bplaque\b", "Plaque amovible", False),
    (r"aligneur|invisalign|spark|goutti[e\xe8]re", "Aligneurs", False),
    (r"frankel|\bFR\b", "FR", False),
    (r"mainteneur", "Mainteneur d'espace", False),
    (r"multi[\s\-]?attache", "Multi-attache", True),
    (r"multi[\s\-]?bague", "Multibague", True),
    (r"\bMBT\b", "Multibague MBT", True),
    (r"\bdamon\b", "Damon", True),
]
ETAT_ORDER = [("appareil", "Appareil"), ("arc_haut", "Max"), ("arc_bas", "Mand"), ("elastiques", "\xc9lastiques")]

def _clean_val(v):
    return (v or "").strip().strip(",|").rstrip(" .\xb7-").strip()

def parse_note_state(text):
    """Comprend un commentaire libre segment par segment (séparés par , | retour ligne) :
    Max/Mand -> arcs ; TIM/Classe/élastique -> élastiques (+ qualifiers D/G, 24h…) ;
    mots-clés -> appareil (avec info arch-based)."""
    out = {}
    elas_open = False
    for raw in re.split(r"[\n,|]+", text or ""):
        seg = raw.strip()
        if not seg:
            continue
        m = re.match(r"(?:arc\s+)?max(?:illaire)?\s*[:=\-]?\s*(.+)$", seg, re.I)
        if m:
            v = _arc_without_accessories(_clean_val(m.group(1)))
            if v: out["arc_haut"] = v
            elas_open = False; continue
        m = re.match(r"(?:arc\s+)?mand(?:ibulaire)?\s*[:=\-]?\s*(.+)$", seg, re.I)
        if m:
            v = _arc_without_accessories(_clean_val(m.group(1)))
            if v: out["arc_bas"] = v
            elas_open = False; continue
        me = re.search(r"((?:[e\xe9]lastiques?|tim\b|classe\s+i{1,3})\b.*)$", seg, re.I)
        if me:
            if re.search(r"arr[e\xea]t|\bstop\b|\bfin\b|plus\s+d|\bsans\b|d[e\xe9]pose|retrait|enlev|\bplus\b", seg, re.I):
                out["_elas_stop"] = True; elas_open = False; continue   # arrêt des élastiques -> on vide
            v = re.sub(r"^(?:mise en place\s+)?[e\xe9]lastiques?\s*[:=\-]?\s*", "", me.group(1), flags=re.I)
            out["elastiques"] = _clean_val(v); elas_open = True; continue
        if elas_open and re.match(r"^(?:[dg](?:\s*/\s*[dg])?|\d+\s*h|nuit|jour(?:s)?|permanent|\d+\s*oz|24\s*/\s*24)$", seg, re.I):
            out["elastiques"] = (out.get("elastiques", "") + " " + seg).strip(); continue
        matched = False
        for pat, lbl, arch in _APP_PATTERNS:
            if re.search(pat, seg, re.I):
                out["appareil"] = lbl; out["_arch"] = arch; matched = True; break
        elas_open = False
    return out

def _ordered_notes(pt):
    """Commentaires du plus récent au plus ancien (date, puis ordre d'ajout)."""
    idx = list(enumerate(pt.get("notes", []) or []))
    idx.sort(key=lambda t: ((t[1].get("date", "") or ""), t[0]), reverse=True)
    return [n for _i, n in idx]

def deduce_etat(pt):
    """Pour chaque champ, valeur du commentaire le plus récent qui le mentionne.
    Appareil : explicite si mentionné ; sinon « Multi-attache » dès qu'un arc est présent."""
    state = {}
    elas_done = False
    for nt in _ordered_notes(pt):
        ps = parse_note_state(nt.get("text", ""))
        for k, _lbl in ETAT_ORDER:
            if k == "elastiques":
                # le commentaire le plus récent qui parle des élastiques décide (valeur OU arrêt)
                if not elas_done and (ps.get("elastiques") or ps.get("_elas_stop")):
                    elas_done = True
                    if ps.get("elastiques") and not ps.get("_elas_stop"):
                        state["elastiques"] = {"val": ps["elastiques"], "date": nt.get("date", "")}
                continue
            if k not in state and ps.get(k):
                state[k] = {"val": ps[k], "date": nt.get("date", "")}
        if "appareil" in ps and "_arch" not in state:
            state["_arch"] = ps.get("_arch", True)
    if "appareil" not in state and ("arc_haut" in state or "arc_bas" in state):
        d = (state.get("arc_haut") or state.get("arc_bas"))["date"]
        state["appareil"] = {"val": "Multi-attache", "date": d}; state["_arch"] = True
    return state

# --- Accessoires : repérés dans les commentaires, avec leur but ---
_ACC_REMOVAL = re.compile(r"d[e\xe9]pos|retrait|retir|enlev|enl[e\xe8]v|arr[e\xea]t|\bfin\b|termin", re.I)
ACCESSORIES = [
    ("cale_retro", r"cale\s*r[e\xe9]tro(?:[\s\-]?incisi\w*)?|butoir\s+incisi\w*", "Cale r\xe9tro-incisive", "d\xe9socclusion (d\xe9gager l'occlusion, lever la supraclusion)"),
    ("ressort", r"ressort", "Ressort", None),
    ("chainette", r"cha[i\xee]nette", "Cha\xeenette \xe9lastom\xe9rique", "fermeture d'espace / traction"),
    ("laceback", r"lace[\s\-]?back", "Lace-back", "recul et contr\xf4le des canines (ancrage)"),
    ("ligature", r"ligature", "Ligature", "maintien du fil"),
    ("butee", r"but[e\xe9]e|lip[\s\-]?bumper|\bbumper\b", "But\xe9e / lip bumper", "recul des molaires, gain de place (ancrage)"),
    ("minivis", r"mini[\s\-]?vis|micro[\s\-]?vis|\bTAD\b|mini[\s\-]?implant", "Mini-vis (ancrage)", "ancrage squelettique"),
    ("crochet", r"crochet", "Crochet", "point d'ancrage (\xe9lastiques / traction)"),
    ("disjonction", r"disjonct|v[e\xe9]rin", "V\xe9rin de disjonction", "expansion maxillaire"),
    ("stop", r"\bstops?\b", "Stop", "emp\xeache le glissement du fil"),
]

def _arc_without_accessories(val):
    """Retire d'une valeur d'arc les mentions d'accessoires (ex. « Acier 19.25 + pose
    cale rétro » -> « Acier 19.25 ») pour ne garder que le fil."""
    if not val:
        return val
    parts = re.split(r"\s*\+\s*", val)
    kept = []
    for p in parts:
        ps = p.strip()
        if not ps:
            continue
        if any(re.search(pat, ps, re.I) for _k, pat, _l, _d in ACCESSORIES):
            continue                      # segment d'accessoire -> retiré de l'arc
        kept.append(ps)
    out = " + ".join(kept)
    return _clean_val(out)

def parse_note_accessories(text):
    """Repère les accessoires d'un commentaire, avec but déduit (ou précisé après « : »)
    et détection d'un éventuel retrait."""
    t = text or ""
    segs = re.split(r"[\n,|;]+|\.\s+", t)   # coupe aussi en fin de phrase (« . »)
    out = {}
    for key, pat, label, dpurp in ACCESSORIES:
        seg = m = None
        for s in segs:
            ss = s.strip()
            mm = re.search(pat, ss, re.I)
            if mm:
                seg, m = ss, mm; break
        if seg is None:
            continue
        removed = bool(_ACC_REMOVAL.search(seg))
        tail = seg[m.start():]                # de l'accessoire jusqu'à la fin (évite le « Max: »)
        purpose = dpurp
        mc = re.search(r":\s*(.+)$", tail)   # une précision « : … » qui suit l'accessoire
        if mc and mc.group(1).strip():
            purpose = mc.group(1).strip()
        elif key == "ressort":
            if re.search(r"ferm", seg, re.I): purpose = "fermeture d'espace"
            elif re.search(r"ouvert", seg, re.I): purpose = "ouverture d'espace"
            else: purpose = "ressort (pr\xe9ciser ouvert / ferm\xe9)"
        detail = _clean_val(tail.split(":", 1)[0])
        out[key] = {"label": label, "purpose": purpose or "", "detail": detail, "removed": removed}
    return out

def deduce_accessories(pt):
    """Accessoires actuellement en place : le dernier commentaire qui mentionne un
    accessoire décide (posé -> affiché ; retiré -> non affiché)."""
    acc = {}; seen = set()
    for nt in _ordered_notes(pt):
        for key, info in parse_note_accessories(nt.get("text", "")).items():
            if key in seen:
                continue
            seen.add(key)
            if not info["removed"]:
                acc[key] = dict(info, date=nt.get("date", ""))
    return acc

def etat_arch_based(state):
    if "_arch" in state:
        return bool(state["_arch"])
    return ("arc_haut" in state or "arc_bas" in state)

def etat_line(pt):
    st = deduce_etat(pt)
    arch = etat_arch_based(st)
    keys = ["appareil"] + (["arc_haut", "arc_bas"] if arch else []) + ["elastiques"]
    labels = dict(ETAT_ORDER)
    return " &nbsp;\xb7&nbsp; ".join("%s : <b>%s</b>" % (labels[k], str(st[k]["val"]).replace("<", "&lt;"))
                                     for k in keys if k in st)

@app.route("/patient/<slug>/suivi")
def suivi(slug):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug)
    if not pt: abort(404)
    _touch_opened(slug)
    notes = _ordered_notes(pt)
    st = deduce_etat(pt)
    arch = etat_arch_based(st)
    today = datetime.date.today().isoformat()
    def fdate(d):
        try: return datetime.date.fromisoformat((d or "")[:10]).strftime("%d/%m/%Y")
        except Exception: return d or ""
    def esc(x): return str(x or "").replace("<", "&lt;")
    def tile(lbl, k, wide=False):
        v = st[k]["val"] if k in st else "—"
        return ('<div style="background:var(--card2);border:1px solid var(--line);border-radius:12px;padding:12px 16px;%s">'
                '<div class=muted style="font-size:12px;text-transform:uppercase;letter-spacing:.03em">%s</div>'
                '<div style="font-size:19px;font-weight:700;color:var(--ink);margin-top:2px">%s</div></div>'
                ) % ("min-width:220px" if wide else "flex:1;min-width:150px", lbl, esc(v))
    arch_tiles = (tile("Max (maxillaire)", "arc_haut") + tile("Mand (mandibulaire)", "arc_bas")) if arch else ""
    # accessoires en place (avec but)
    acc = deduce_accessories(pt)
    acc_block = ""
    if acc:
        items = ""
        for _k, info in acc.items():
            det = info.get("detail", "")
            extra = ('<div class=muted style="font-size:11px;margin-top:2px">%s</div>' % esc(det)) \
                if det and det.lower() not in info["label"].lower() else ""
            items += ('<div style="background:var(--card2);border:1px solid var(--line);border-radius:10px;padding:8px 12px;min-width:190px">'
                      '<div style="font-weight:700;color:var(--ink)">%s</div>'
                      '<div class=muted style="font-size:12px">%s</div>%s</div>'
                      ) % (esc(info["label"]), esc(info["purpose"]), extra)
        acc_block = ('<div style="margin-top:14px"><div class=muted style="font-size:12px;text-transform:uppercase;'
                     'letter-spacing:.03em;margin-bottom:6px">Accessoires</div>'
                     '<div style="display:flex;gap:10px;flex-wrap:wrap">%s</div></div>') % items
    etat_card = ('<div class=card style="background:var(--card2);border-color:var(--line)">'
                 '<h2 style="margin:0 0 12px">\xc9tat actuel du traitement</h2>'
                 '<div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:10px">%s</div>'
                 '<div style="display:flex;gap:12px;flex-wrap:wrap">%s%s</div>'
                 '%s'
                 '<p class=muted style="font-size:12px;margin:10px 0 0">D\xe9duit de tes commentaires (le plus r\xe9cent fait foi). '
                 'Accessoires : \xe9cris-les dans le commentaire (ex. <i>ressort ouvert : place pour la 12</i>, <i>cale r\xe9tro</i>) ; '
                 'ils restent affich\xe9s jusqu\'\xe0 ce que tu notes leur retrait (ex. <i>d\xe9pose ressort</i>).</p></div>'
                 ) % (tile("Appareil", "appareil", wide=True), arch_tiles, tile("\xc9lastiques", "elastiques"), acc_block)

    rows = ""
    for nt in list(reversed(notes)):          # plus ancien en haut, plus récent en bas (près de la saisie)
        ps = parse_note_state(nt.get("text", ""))
        eid = nt.get("id", "")
        edit = ('<details><summary style="cursor:pointer;color:var(--acc);font-size:12px">modifier</summary>'
                '<form method=post action="%s" style="margin-top:6px">'
                '<input type=date name=n_date value="%s" style="width:auto;margin-bottom:6px">'
                '<textarea name=n_text rows=2 style="width:100%%">%s</textarea>'
                '<div style="margin-top:6px"><button class="btn sec sm">Enregistrer</button></div></form>'
                '<form method=post action="%s" onsubmit="return confirm(\'Supprimer ce commentaire ?\')">'
                '<button class="btn sec sm" style="color:#C0392B">Supprimer</button></form></details>'
                ) % (url_for("note_edit", slug=slug, eid=eid), (nt.get("date", "") or today)[:10],
                     esc(nt.get("text", "")), url_for("note_delete", slug=slug, eid=eid))
        rows += ('<tr>'
                 '<td style="white-space:nowrap;font-weight:600;vertical-align:top">%s</td>'
                 '<td style="vertical-align:top"><div style="white-space:pre-line">%s</div></td>'
                 '<td style="vertical-align:top">%s</td>'
                 '<td style="vertical-align:top">%s</td>'
                 '<td style="vertical-align:top">%s</td>'
                 '<td style="vertical-align:top;text-align:right;white-space:nowrap">%s</td></tr>'
                 ) % (fdate(nt.get("date", "")), esc(nt.get("text", "")),
                      esc(ps.get("arc_haut", "")), esc(ps.get("arc_bas", "")), esc(ps.get("elastiques", "")), edit)
    if not rows:
        rows = '<tr><td colspan=6 class=muted>Aucun commentaire pour l\'instant.</td></tr>'
    table = ('<table class=st style="width:100%%;border-collapse:collapse">'
             '<tr><th style="text-align:left">Date</th><th style="text-align:left">Commentaire</th>'
             '<th style="text-align:left">Arc Max.</th><th style="text-align:left">Arc Mand.</th>'
             '<th style="text-align:left">\xc9lastique</th><th style="text-align:right"></th></tr>%s</table>') % rows

    add = ('<form method=post action="%s" style="margin-top:14px;border-top:1px solid var(--line);padding-top:12px">'
           '<div style="display:flex;gap:12px;align-items:flex-start;flex-wrap:wrap">'
           '<div><label>Date du jour</label><input type=date name=n_date value="%s" style="width:auto"></div>'
           '<div style="flex:1;min-width:280px"><label>Nouveau commentaire</label>'
           '<textarea name=n_text rows=3 placeholder="Max: 17.25 NiTi, Mand: 17.25 NiTi, TIM II D/G 24h  — ou  Quad helix / arrêt élastiques" style="width:100%%"></textarea></div>'
           '</div><div style="margin-top:10px"><button class=btn onclick="this.innerHTML=\'<span class=spin></span> Ajout…\'">Ajouter le commentaire</button></div></form>'
           ) % (url_for("note_add", slug=slug), today)
    suivi_card = '<div class=card><h2 style="margin:0 0 10px">Suivi</h2>%s%s</div>' % (table, add)

    body = '%s%s%s' % (phead(slug, pt, "suivi"), etat_card, suivi_card)
    return page(body)

@app.route("/patient/<slug>/suivi/note/add", methods=["POST"])
def note_add(slug):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug)
    if not pt: abort(404)
    txt = request.form.get("n_text", "").strip()
    if txt:
        nt = {"id": datetime.datetime.now().strftime("%Y%m%d%H%M%S") + secrets.token_hex(2),
              "date": (request.form.get("n_date", "") or datetime.date.today().isoformat())[:10], "text": txt}
        pt.setdefault("notes", []).append(nt); save_patient(slug, pt); flash("Commentaire ajout\xe9.")
    else:
        flash("Commentaire vide.")
    return redirect(url_for("suivi", slug=slug))

@app.route("/patient/<slug>/suivi/note/<eid>/edit", methods=["POST"])
def note_edit(slug, eid):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug)
    if not pt: abort(404)
    for nt in pt.get("notes", []) or []:
        if nt.get("id") == eid:
            nt["text"] = request.form.get("n_text", "").strip()
            nt["date"] = (request.form.get("n_date", "") or nt.get("date", ""))[:10]
            break
    save_patient(slug, pt); flash("Commentaire mis \xe0 jour.")
    return redirect(url_for("suivi", slug=slug))

@app.route("/patient/<slug>/suivi/note/<eid>/delete", methods=["POST"])
def note_delete(slug, eid):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug)
    if not pt: abort(404)
    pt["notes"] = [x for x in (pt.get("notes", []) or []) if x.get("id") != eid]
    save_patient(slug, pt); flash("Commentaire supprim\xe9.")
    return redirect(url_for("suivi", slug=slug))

# ---- gestion d'un enregistrement : date/libellé + suppression (corbeille) ----
@app.route("/record/<slug>/<rid>/setmeta", methods=["POST"])
def record_setmeta(slug, rid):
    if not logged(): return redirect(url_for("login"))
    r = load_rec(slug, rid)
    if not r: abort(404)
    d = request.form.get("date", "").strip()
    lbl = request.form.get("label", "").strip()
    if d:
        old = r.get("date", "")
        r["date"] = (d + old[10:]) if (len(old) > 10 and old[10] == "T") else (d + "T09:00:00")
    if lbl: r["label"] = lbl
    save_rec(slug, rid, r); flash("Enregistrement mis à jour.")
    return redirect(url_for("patient", slug=slug))

@app.route("/record/<slug>/<rid>/delete", methods=["GET", "POST"])
def record_delete(slug, rid):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r: abort(404)
    if request.method == "POST":
        base = rdir(slug, rid)
        trash = os.path.join(pdir(slug), "_corbeille"); os.makedirs(trash, exist_ok=True)
        if os.path.isdir(base):
            dest = os.path.join(trash, rid)
            if os.path.exists(dest): dest += "_" + secrets.token_hex(3)
            try: shutil.move(base, dest)
            except Exception: pass
        pt["records"] = [x for x in pt.get("records", []) if x != rid]
        save_patient(slug, pt); flash("Enregistrement déplacé dans la corbeille.")
        return redirect(url_for("patient", slug=slug))
    return page("""<h1>Supprimer un enregistrement</h1>
    <div class=card><p>Confirmer la suppression de <b>%s</b> (%s) ?</p>
    <p class=muted>Il sera déplacé dans une corbeille (récupérable dans le dossier du patient), pas effacé définitivement.</p>
    <form method=post style="display:inline"><button class=btn style="background:#C0392B">Oui, mettre à la corbeille</button></form>
    <a class="btn sec" href="%s">Annuler</a></div>""" % (
        r.get("label", ""), r.get("date", "")[:10], url_for("patient", slug=slug)))

# ======================================================================
#  RÉGLAGES
# ======================================================================
DOC_KC_SERVICE = "BilanODF-Doctolib"
DOC_KC_ACCOUNT = "doctolib"

def doctolib_creds_get():
    import subprocess, json as _j
    try:
        r = subprocess.run(["security", "find-generic-password", "-s", DOC_KC_SERVICE, "-a", DOC_KC_ACCOUNT, "-w"],
                           capture_output=True, text=True, timeout=6)
        if r.returncode != 0: return {}
        raw = (r.stdout or "").strip()
        if not raw: return {}
        try: return _j.loads(raw)
        except Exception: return {}
    except Exception:
        return {}

def doctolib_creds_set(idv, pw):
    import subprocess, json as _j
    blob = _j.dumps({"id": idv, "pw": pw})
    try:
        subprocess.run(["security", "add-generic-password", "-U", "-s", DOC_KC_SERVICE, "-a", DOC_KC_ACCOUNT, "-w", blob],
                       capture_output=True, text=True, timeout=6)
        return True
    except Exception:
        return False

def _doctolib_kc_clear():
    import subprocess
    try:
        subprocess.run(["security", "delete-generic-password", "-s", DOC_KC_SERVICE, "-a", DOC_KC_ACCOUNT],
                       capture_output=True, text=True, timeout=6)
    except Exception:
        pass

@app.route("/doctolib/creds", methods=["POST"])
def doctolib_creds():
    if not logged(): return redirect(url_for("login"))
    idv = request.form.get("doc_id", "").strip()
    pw = request.form.get("doc_pw", "")
    if not pw:
        old = doctolib_creds_get(); pw = old.get("pw", "")
    if idv or pw:
        doctolib_creds_set(idv, pw); flash("Identifiants Doctolib enregistres (Trousseau macOS).")
    else:
        flash("Renseigne l'identifiant et le mot de passe.")
    return redirect(url_for("reglages"))

@app.route("/doctolib/creds/clear")
def doctolib_creds_clear():
    if not logged(): return redirect(url_for("login"))
    _doctolib_kc_clear(); flash("Identifiants Doctolib supprimes.")
    return redirect(url_for("reglages"))


XERO_KC_SERVICE = "BilanODF-XERO"
XERO_KC_ACCOUNT = "xero"

def xero_creds_get():
    import subprocess, json as _j
    try:
        r = subprocess.run(["security", "find-generic-password", "-s", XERO_KC_SERVICE, "-a", XERO_KC_ACCOUNT, "-w"],
                           capture_output=True, text=True, timeout=6)
        if r.returncode != 0: return {}
        raw = (r.stdout or "").strip()
        if not raw: return {}
        try: return _j.loads(raw)
        except Exception: return {}
    except Exception:
        return {}

def xero_creds_set(idv, pw):
    import subprocess, json as _j
    blob = _j.dumps({"id": idv, "pw": pw})
    try:
        subprocess.run(["security", "add-generic-password", "-U", "-s", XERO_KC_SERVICE, "-a", XERO_KC_ACCOUNT, "-w", blob],
                       capture_output=True, text=True, timeout=6)
        return True
    except Exception:
        return False

def _xero_kc_clear():
    import subprocess
    try:
        subprocess.run(["security", "delete-generic-password", "-s", XERO_KC_SERVICE, "-a", XERO_KC_ACCOUNT],
                       capture_output=True, text=True, timeout=6)
    except Exception:
        pass

@app.route("/xero/creds", methods=["POST"])
def xero_creds():
    if not logged(): return redirect(url_for("login"))
    idv = request.form.get("xero_id", "").strip()
    pw = request.form.get("xero_pw", "")
    if not pw:
        old = xero_creds_get(); pw = old.get("pw", "")
    if idv or pw:
        xero_creds_set(idv, pw); flash("Identifiants XERO enregistres (Trousseau macOS).")
    else:
        flash("Renseigne l'identifiant et le mot de passe.")
    return redirect(url_for("reglages"))

@app.route("/xero/creds/clear")
def xero_creds_clear():
    if not logged(): return redirect(url_for("login"))
    _xero_kc_clear(); flash("Identifiants XERO supprimes.")
    return redirect(url_for("reglages"))


WEBCEPH_KC_SERVICE = "BilanODF-WebCeph"
WEBCEPH_KC_ACCOUNT = "webceph"

def webceph_creds_get():
    import subprocess, json as _j
    try:
        r = subprocess.run(["security", "find-generic-password", "-s", WEBCEPH_KC_SERVICE, "-a", WEBCEPH_KC_ACCOUNT, "-w"],
                           capture_output=True, text=True, timeout=6)
        if r.returncode != 0: return {}
        raw = (r.stdout or "").strip()
        if not raw: return {}
        try: return _j.loads(raw)
        except Exception: return {}
    except Exception:
        return {}

def webceph_creds_set(idv, pw):
    import subprocess, json as _j
    blob = _j.dumps({"id": idv, "pw": pw})
    try:
        subprocess.run(["security", "add-generic-password", "-U", "-s", WEBCEPH_KC_SERVICE, "-a", WEBCEPH_KC_ACCOUNT, "-w", blob],
                       capture_output=True, text=True, timeout=6)
        return True
    except Exception:
        return False

def _webceph_kc_clear():
    import subprocess
    try:
        subprocess.run(["security", "delete-generic-password", "-s", WEBCEPH_KC_SERVICE, "-a", WEBCEPH_KC_ACCOUNT],
                       capture_output=True, text=True, timeout=6)
    except Exception:
        pass

@app.route("/webceph/creds", methods=["POST"])
def webceph_creds():
    if not logged(): return redirect(url_for("login"))
    idv = request.form.get("wc_id", "").strip()
    pw = request.form.get("wc_pw", "")
    if not pw:
        old = webceph_creds_get(); pw = old.get("pw", "")
    if idv or pw:
        webceph_creds_set(idv, pw); flash("Identifiants WebCeph enregistres (Trousseau macOS).")
    else:
        flash("Renseigne l'identifiant et le mot de passe.")
    return redirect(url_for("reglages"))

@app.route("/webceph/creds/clear")
def webceph_creds_clear():
    if not logged(): return redirect(url_for("login"))
    _webceph_kc_clear(); flash("Identifiants WebCeph supprimes.")
    return redirect(url_for("reglages"))


@app.route("/reglages", methods=["GET", "POST"])
def reglages():
    if not logged(): return redirect(url_for("login"))
    st = get_settings()
    if request.method == "POST":
        st["praticien"] = request.form.get("praticien", "").strip()
        st["evo_blocks"] = {k: (k in request.form.getlist("evo_blocks")) for k, _ in EVO_BLOCK_DEFS}
        st["evo_photos"] = request.form.getlist("evo_photos")
        st["auto_backup"] = bool(request.form.get("auto_backup"))
        save_settings(st); flash("Réglages enregistrés.")
        return redirect(url_for("reglages"))
    blocks = st.get("evo_blocks") or {}
    photos_on = st.get("evo_photos")
    if photos_on is None: photos_on = [k for k, _ in EVO_PHOTOS]
    CHK = "display:inline-flex;align-items:center;gap:7px;white-space:nowrap;border:1px solid var(--line);border-radius:8px;padding:6px 11px;margin:0 8px 8px 0;background:var(--card2);font-size:13px;cursor:pointer"
    def chk(name, val, lbl, on):
        return ('<label style="%s"><input type=checkbox name=%s value="%s" %s style="width:auto;margin:0"> %s</label>'
                % (CHK, name, val, "checked" if on else "", lbl))
    bx = "".join(chk("evo_blocks", k, lbl, blocks.get(k, True)) for k, lbl in EVO_BLOCK_DEFS)
    px = "".join(chk("evo_photos", k, lbl, k in photos_on) for k, lbl in EVO_PHOTOS)
    # --- sauvegarde & sécurité ---
    auto_html = chk("auto_backup", "1", "Sauvegarde automatique quotidienne (à l'ouverture de l'app)", bool(st.get("auto_backup")))
    def _hsize(n):
        for u in ("o", "Ko", "Mo", "Go"):
            if n < 1024 or u == "Go": return "%.0f %s" % (n, u)
            n /= 1024.0
    rows = ""
    for name, size, mt in BK.list_backups(BACKUP_DIR):
        enc = BK.is_encrypted(os.path.join(BACKUP_DIR, name))
        when = datetime.datetime.fromtimestamp(mt).strftime("%d/%m/%Y %H:%M")
        lock = ' 🔒' if enc else ''
        rows += ('<form method=post action="%s" style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;'
                 'border-top:1px solid #eef;padding:8px 0;margin:0">'
                 '<input type=hidden name=name value="%s">'
                 '<div style="flex:1;min-width:220px"><b>%s</b>%s<div class=muted style="font-size:12px">%s · %s</div></div>'
                 '%s'
                 '<button class="btn sec sm" onclick="return confirm(\'Restaurer cette sauvegarde ? Les données actuelles seront mises de côté.\')">Restaurer</button>'
                 '</form>') % (
                    url_for("backup_restore"), name.replace('"', "&quot;"), name, lock, when, _hsize(size),
                    ('<input type=password name=password placeholder="mot de passe" style="width:auto;max-width:170px">' if enc else ''))
    if not rows:
        rows = '<p class=muted>Aucune sauvegarde pour le moment.</p>'
    aes_note = "" if BK.HAVE_AES else '<p class=muted style="color:#8a5a00">⚠ Chiffrement indisponible sur cette installation (module « pyzipper » manquant) — les sauvegardes seront non chiffrées.</p>'
    backup_card = ("""<div class=card><h2>Sauvegarde &amp; sécurité</h2>
      <p class=muted style="margin:0 0 6px">Les sauvegardes sont enregistrées dans <b>~/BilanODF_Sauvegardes</b>. Copie-les régulièrement
      sur un disque externe ou un espace sécurisé. Pour chiffrer <b>le dossier de données vivant</b>, active <b>FileVault</b>
      dans Réglages système macOS (chiffrement intégral du disque, recommandé pour des données de santé).</p>
      %s
      <form method=post action="%s" style="display:flex;gap:10px;align-items:end;flex-wrap:wrap;margin:10px 0">
        <div><label>Mot de passe (optionnel, chiffre la sauvegarde en AES-256)</label>
        <input type=password name=password placeholder="laisser vide = non chiffrée" style="width:auto;min-width:260px"></div>
        <button class=btn onclick="this.innerHTML='<span class=spin></span> Sauvegarde…'">Créer une sauvegarde maintenant</button>
      </form>
      <h3 style="margin:14px 0 4px;color:var(--acc);font-size:14px">Sauvegardes disponibles</h3>
      %s
    </div>""") % (aes_note, url_for("backup_create"), rows)
    # --- reconnaissance photo (IA locale auto-apprise) ---
    pm = st.get("photo_model") or {}
    model_exists = os.path.exists(os.path.join(DATA, "photo_model.npz"))
    if pm:
        det = "".join("%s : %d · " % (dict(PHOTO_FIELDS).get(k, k), v) for k, v in sorted(pm.get("counts", {}).items()))
        status = ('<div class=flash style="background:rgba(34,197,94,.12);border-color:rgba(34,197,94,.35)">Modèle actif — calibré le <b>%s</b> sur '
                  '<b>%d photos</b>, fiabilité estimée <b>%d%%</b> (validation croisée).<div class=muted style="font-size:12px;margin-top:4px">%s</div></div>'
                  ) % (pm.get("when", "?"), pm.get("n", 0), round(pm.get("acc", 0) * 100), det.rstrip(" ·"))
    elif model_exists:
        status = '<div class=flash style="background:rgba(34,197,94,.12);border-color:rgba(34,197,94,.35)">Un modèle est actif.</div>'
    else:
        status = '<p class=muted>Aucune calibration pour l\'instant — l\'app utilise des règles génériques (parfois imprécises sur les vues endobuccales).</p>'
    ia_card = ("""<div class=card><h2>Reconnaissance des photos (IA locale)</h2>
      <p class=muted style="margin:0 0 6px">Améliore le classement automatique de l'import « tout auto » en apprenant sur <b>tes propres photos déjà rangées</b>
      (aucune donnée ne quitte ton Mac, rien à copier). Relance-la de temps en temps, surtout après avoir ajouté ou corrigé des photos.</p>
      %s
      <form method=post action="%s"><button class=btn onclick="this.innerHTML='<span class=spin></span> Calibration…'">Recalibrer la reconnaissance maintenant</button></form>
    </div>""") % (status, url_for("recalibrate"))
    # --- calibrage du recadrage (marges + redressement) ---
    cc = st.get("crop_calib") or {}
    FAML = {"exo": "Exobuccales", "endo": "Endobuccales", "occlusal": "Occlusales"}
    if cc and cc.get("calib"):
        det = ""
        for f, e in cc.get("calib", {}).items():
            bits = []
            if "margin" in e: bits.append("marge %d%%" % round(e["margin"] * 100))
            if "rot_gain" in e: bits.append("redressement appris")
            det += "%s : %s · " % (FAML.get(f, f), ", ".join(bits) if bits else "—")
        cstatus = ('<div class=flash style="background:rgba(34,197,94,.12);border-color:rgba(34,197,94,.35)">Calibré le <b>%s</b> sur <b>%d</b> cadrage(s) ajusté(s).'
                   '<div class=muted style="font-size:12px;margin-top:4px">%s</div></div>') % (cc.get("when", "?"), cc.get("n", 0), det.rstrip(" ·"))
    else:
        cstatus = '<p class=muted>Pas encore calibré : l\'app utilise des marges resserrées par défaut et un redressement automatique du plan d\'occlusion.</p>'
    crop_card = ("""<div class=card><h2>Recadrage (calibrage)</h2>
      <p class=muted style="margin:0 0 6px">Le recadrage auto détecte la zone utile et la resserre (aucune rotation automatique — le redressement se fait
      seulement à la main, avec la molette de l'éditeur). Il <b>apprend ta marge préférée</b> en comparant tes photos
      <b>brutes</b> et <b>recadrées</b> déjà présentes — rien à recadrer à nouveau. Le bouton ne sert qu'à forcer un recalcul.</p>
      %s
      <form method=post action="%s"><button class="btn sec" onclick="this.innerHTML='<span class=spin></span> Calibrage…'">Réapprendre depuis mes photos recadrées</button></form>
    </div>""") % (cstatus, url_for("recalibrate_crop"))
    _dc = doctolib_creds_get()
    _dc_saved = bool(_dc.get("id") or _dc.get("pw"))
    _dc_status = ('<div class=flash style="background:rgba(34,197,94,.12);border-color:rgba(34,197,94,.35)">Identifiants enregistrés (Trousseau macOS). Ils sont pré-remplis automatiquement sur la page de connexion Doctolib.</div>' if _dc_saved else '<p class=muted>Aucun identifiant enregistré pour l\'instant.</p>')
    _dc_del = (('<a class="btn sec" href="%s" onclick="return confirm(\'Supprimer les identifiants Doctolib enregistrés ?\')">Supprimer</a>' % url_for("doctolib_creds_clear")) if _dc_saved else "")
    doctolib_card = ("""<div class=card><h2>Connexion Doctolib (pré-remplissage)</h2>
      <p class=muted style="margin:0 0 6px">Saisis tes identifiants Doctolib <b>pro</b>. Ils sont stockés <b>uniquement sur ce Mac, dans le Trousseau (Keychain)</b> — jamais en clair, jamais envoyés ailleurs — et servent à pré-remplir la page de connexion Doctolib.</p>
      %s
      <form method=post action="%s" class="grid g2" style="max-width:620px;align-items:end">
        <div><label>Identifiant Doctolib (e-mail)</label><input name=doc_id value="%s" autocomplete=off></div>
        <div><label>Mot de passe Doctolib</label><input type=password name=doc_pw placeholder="%s" autocomplete=new-password></div>
        <div style="grid-column:1/-1;display:flex;gap:10px"><button class=btn>Enregistrer</button>%s</div>
      </form>
    </div>""") % (_dc_status, url_for("doctolib_creds"), (_dc.get("id", "") or "").replace('"', "&quot;"),
                  ("•••••• (déjà enregistré, laisser vide pour ne pas changer)" if _dc_saved else "mot de passe"), _dc_del)
    _xc = xero_creds_get()
    _xc_saved = bool(_xc.get("id") or _xc.get("pw"))
    _xc_status = ('<div class=flash style="background:rgba(34,197,94,.12);border-color:rgba(34,197,94,.35)">Identifiants enregistrés (Trousseau macOS). Ils sont pré-remplis automatiquement sur la page de connexion XERO.</div>' if _xc_saved else '<p class=muted>Aucun identifiant enregistré pour l\'instant.</p>')
    _xc_del = (('<a class="btn sec" href="%s" onclick="return confirm(\'Supprimer les identifiants XERO enregistrés ?\')">Supprimer</a>' % url_for("xero_creds_clear")) if _xc_saved else "")
    xero_card = ("""<div class=card><h2>Connexion XERO (pré-remplissage)</h2>
      <p class=muted style="margin:0 0 6px">Saisis tes identifiants XERO (CHU Nice). Ils sont stockés <b>uniquement sur ce Mac, dans le Trousseau (Keychain)</b> — jamais en clair, jamais envoyés ailleurs — et servent à pré-remplir la page de connexion XERO.</p>
      %s
      <form method=post action="%s" class="grid g2" style="max-width:620px;align-items:end">
        <div><label>Identifiant XERO</label><input name=xero_id value="%s" autocomplete=off></div>
        <div><label>Mot de passe XERO</label><input type=password name=xero_pw placeholder="%s" autocomplete=new-password></div>
        <div style="grid-column:1/-1;display:flex;gap:10px"><button class=btn>Enregistrer</button>%s</div>
      </form>
    </div>""") % (_xc_status, url_for("xero_creds"), (_xc.get("id", "") or "").replace('"', "&quot;"),
                  ("•••••• (déjà enregistré, laisser vide pour ne pas changer)" if _xc_saved else "identifiant / mot de passe"), _xc_del)
    _wc = webceph_creds_get()
    _wc_saved = bool(_wc.get("id") or _wc.get("pw"))
    _wc_status = ('<div class=flash style="background:rgba(34,197,94,.12);border-color:rgba(34,197,94,.35)">Identifiants enregistrés (Trousseau macOS). Ils sont pré-remplis automatiquement sur la page de connexion WebCeph.</div>' if _wc_saved else '<p class=muted>Aucun identifiant enregistré pour l\'instant.</p>')
    _wc_del = (('<a class="btn sec" href="%s" onclick="return confirm(\'Supprimer les identifiants WebCeph enregistrés ?\')">Supprimer</a>' % url_for("webceph_creds_clear")) if _wc_saved else "")
    webceph_card = ("""<div class=card><h2>Connexion WebCeph (pré-remplissage)</h2>
      <p class=muted style="margin:0 0 6px">Saisis tes identifiants WebCeph. Ils sont stockés <b>uniquement sur ce Mac, dans le Trousseau (Keychain)</b> — jamais en clair, jamais envoyés ailleurs — et servent à pré-remplir la page de connexion WebCeph.</p>
      %s
      <form method=post action="%s" class="grid g2" style="max-width:620px;align-items:end">
        <div><label>Identifiant WebCeph (e-mail)</label><input name=wc_id value="%s" autocomplete=off></div>
        <div><label>Mot de passe WebCeph</label><input type=password name=wc_pw placeholder="%s" autocomplete=new-password></div>
        <div style="grid-column:1/-1;display:flex;gap:10px"><button class=btn>Enregistrer</button>%s</div>
      </form>
    </div>""") % (_wc_status, url_for("webceph_creds"), (_wc.get("id", "") or "").replace('"', "&quot;"),
                  ("•••••• (déjà enregistré, laisser vide pour ne pas changer)" if _wc_saved else "identifiant / mot de passe"), _wc_del)
    body = """<h1>⚙ Réglages</h1>
    <div class=card><h2>Compte</h2>
      <p>Connecté en tant que <b>%s</b>.</p>
      <form method=post action="%s" class="grid g3" style="max-width:620px;align-items:end">
        <div><label>Mot de passe actuel</label><input type=password name=current></div>
        <div><label>Nouveau mot de passe</label><input type=password name=new></div>
        <div><label>Confirmer</label><input type=password name=confirm></div>
        <div style="grid-column:1/-1"><button class="btn sec">Changer le mot de passe</button></div>
      </form>
    </div>
    <form method=post>
      <div class=card><h2>Bilans</h2>
        <label>Nom du praticien / cabinet (affiché sur les bilans)</label>
        <input name=praticien value="%s" style="max-width:520px">
      </div>
      <div class=card><h2>Affichage de la vue Évolution</h2>
        <p class=muted style="margin:0 0 8px">Choisis les blocs à afficher :</p>
        <div style="display:flex;flex-wrap:wrap">%s</div>
        <p class=muted style="margin:12px 0 8px">Et les vues photo à comparer :</p>
        <div style="display:flex;flex-wrap:wrap">%s</div>
      </div>
      <div class=card><h2>Sauvegarde automatique</h2>
        <div style="display:flex;flex-wrap:wrap">%s</div>
        <p class=muted style="margin:6px 0 0;font-size:12px">Une sauvegarde (non chiffrée) est alors créée une fois par jour à l'ouverture de l'app. Les 15 plus récentes sont conservées.</p>
      </div>
      <button class=btn>Enregistrer les réglages</button>
    </form>
    %s
    %s
    %s
    %s
    %s
    %s""" % (session["user"], url_for("reglages_password"),
                  (st.get("praticien", "") or "").replace('"', "&quot;"), bx, px, auto_html, ia_card, crop_card, backup_card, doctolib_card, xero_card, webceph_card)
    return page(body)

@app.route("/reglages/password", methods=["POST"])
def reglages_password():
    if not logged(): return redirect(url_for("login"))
    u = session["user"]
    cur = request.form.get("current", ""); new = request.form.get("new", ""); conf = request.form.get("confirm", "")
    if not check_password_hash(cfg["users"].get(u, ""), cur):
        flash("Mot de passe actuel incorrect.")
    elif len(new) < 4 or new != conf:
        flash("Nouveau mot de passe trop court ou non confirmé (min. 4 caractères).")
    else:
        cfg["users"][u] = generate_password_hash(new); save_config(cfg); flash("Mot de passe modifié.")
    return redirect(url_for("reglages"))

# ---- Sauvegarde & sécurité ----
def _do_backup(password=None):
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    return BK.create_backup(DATA, BACKUP_DIR, stamp, password=password or None, keep=15)

def auto_backup_if_due():
    """Sauvegarde automatique quotidienne (silencieuse) si activée dans les réglages."""
    st = get_settings()
    if not st.get("auto_backup"):
        return
    today = datetime.date.today().isoformat()
    if st.get("last_auto_backup") == today:
        return
    try:
        _do_backup(password=None)
        st["last_auto_backup"] = today; save_settings(st)
    except Exception:
        pass

@app.route("/backup/create", methods=["POST"])
def backup_create():
    if not logged(): return redirect(url_for("login"))
    pw = request.form.get("password", "").strip()
    if pw and not BK.HAVE_AES:
        flash("Chiffrement indisponible (module manquant) — sauvegarde créée sans mot de passe.")
        pw = ""
    try:
        dest = _do_backup(password=pw or None)
        flash("Sauvegarde créée : %s%s" % (os.path.basename(dest), " (chiffrée)" if pw else ""))
    except Exception as e:
        flash("Échec de la sauvegarde : %s" % e)
    return redirect(url_for("reglages"))

@app.route("/backup/restore", methods=["POST"])
def backup_restore():
    if not logged(): return redirect(url_for("login"))
    name = os.path.basename(request.form.get("name", ""))
    pw = request.form.get("password", "").strip()
    path = os.path.join(BACKUP_DIR, name)
    if not name or not os.path.exists(path):
        flash("Sauvegarde introuvable."); return redirect(url_for("reglages"))
    ok, msg = BK.restore_backup(path, DATA, password=pw or None)
    flash(msg + (" Les données précédentes ont été mises de côté (…_avant_restauration)." if ok else ""))
    return redirect(url_for("reglages"))

# ---- Reconnaissance des photos : calibration sur les photos déjà rangées ----
@app.route("/recalibrate", methods=["POST"])
def recalibrate():
    if not logged(): return redirect(url_for("login"))
    try:
        rep = P.train_photo_model(PATIENTS, os.path.join(DATA, "photo_model.npz"))
    except Exception as e:
        flash("Échec de la calibration : %s" % e); return redirect(url_for("reglages"))
    if rep.get("ok"):
        st = get_settings()
        st["photo_model"] = {"n": rep["n"], "acc": rep["acc"], "counts": rep.get("counts", {}),
                             "when": datetime.datetime.now().strftime("%d/%m/%Y %H:%M")}
        save_settings(st)
        flash("Reconnaissance calibrée sur %d photos déjà rangées — fiabilité estimée %d%% (validation croisée)."
              % (rep["n"], round(rep["acc"] * 100)))
    else:
        flash(rep.get("msg", "Pas assez de photos déjà rangées pour calibrer."))
    return redirect(url_for("reglages"))

def auto_crop_calib():
    """Recalibre le recadrage automatiquement (silencieux) après chaque ajustement
    manuel, en se basant sur tes cadrages les plus récents."""
    try:
        rep = P.train_crop_calib(PATIENTS, os.path.join(DATA, "crop_calib.json"))
        if rep.get("ok"):
            st = get_settings()
            st["crop_calib"] = {"n": rep["n"], "calib": rep.get("calib", {}),
                                "when": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"), "auto": True}
            save_settings(st)
    except Exception:
        pass

@app.route("/recalibrate_crop", methods=["POST"])
def recalibrate_crop():
    if not logged(): return redirect(url_for("login"))
    try:
        rep = P.train_crop_calib(PATIENTS, os.path.join(DATA, "crop_calib.json"))
    except Exception as e:
        flash("Échec du calibrage du recadrage : %s" % e); return redirect(url_for("reglages"))
    if rep.get("ok"):
        st = get_settings()
        st["crop_calib"] = {"n": rep["n"], "calib": rep.get("calib", {}),
                            "when": datetime.datetime.now().strftime("%d/%m/%Y %H:%M")}
        save_settings(st)
        if rep["n"]:
            flash("Recadrage calibré sur %d cadrage(s) que tu as ajusté(s) — marges et redressement affinés." % rep["n"])
        else:
            flash("Aucun cadrage ajusté trouvé pour l'instant : ajuste quelques photos à la main (cadre + molette de redressement), puis relance le calibrage.")
    else:
        flash("Échec : %s" % rep.get("msg", "inconnu"))
    return redirect(url_for("reglages"))

# ordre d'affichage des photos et rendus dans la vue évolution
EVO_PHOTOS = [("exo_face_repos", "Visage repos"), ("exo_face_sourire", "Visage sourire"), ("exo_profil", "Profil"),
              ("endo_occlusal_maxillaire", "Occlusal maxillaire"), ("endo_occlusal_mandibulaire", "Occlusal mandibulaire"),
              ("endo_occlusion_frontale", "Occlusion frontale"), ("endo_laterale_droite", "Latérale droite"),
              ("endo_laterale_gauche", "Latérale gauche")]
EVO_BLOCK_DEFS = [("photos", "Photos"), ("radios", "Radiographies"), ("stl", "Modèles 3D"), ("steiner", "Évolution du Steiner")]

@app.route("/patient/<slug>/evolution")
def evolution(slug):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug)
    if not pt: abort(404)
    all_recs = [(rid, load_rec(slug, rid)) for rid in pt.get("records", [])]
    all_recs = [(rid, r) for rid, r in all_recs if r]
    def _evo_order(item):
        # Ordre logique du traitement : Bilan initial d'abord (à gauche), puis Réévaluation 1, 2…
        # (on n'utilise PAS la date de fichier, peu fiable, mais l'étape de traitement.)
        rid, r = item; lbl = (r.get("label", "") or "").lower()
        if lbl.startswith("bilan"):
            return (0, 0, r.get("date", ""))
        m = re.search(r"(\d+)", lbl)
        return (1, int(m.group(1)) if m else 99, r.get("date", ""))
    all_recs.sort(key=_evo_order)
    if not all_recs:
        return page(phead(slug, pt, "evolution") + "<p class=muted>Aucun enregistrement pour l'instant.</p>")
    # sélection des temps affichés (par défaut tous)
    sel = request.args.getlist("cols")
    recs = [(rid, r) for rid, r in all_recs if rid in sel] if sel else all_recs
    if not recs: recs = all_recs
    n = len(recs); LABW = 130
    tindex = {rid: i + 1 for i, (rid, _r) in enumerate(all_recs)}   # T1, T2… (ordre chrono, tous temps)
    def uf(rid, path): return url_for("rfile", slug=slug, rid=rid, path=path)
    def has(rid, path): return os.path.exists(os.path.join(rdir(slug, rid), path))
    # colonnes proportionnelles : la 1re (libellés) fixe, les temps se partagent 100% -> photos au max
    colgroup = '<colgroup><col style="width:%dpx">%s</colgroup>' % (LABW, "<col>" * n)
    tstyle = "border-collapse:collapse;table-layout:fixed;width:100%"
    IMGST = "width:100%;height:auto;display:block;border-radius:6px;border:1px solid #e3e8ef"
    # sélecteur de temps (chips) — width:auto sur la case pour éviter le débordement
    chk = ""
    for rid, r in all_recs:
        on = "checked" if (rid in sel or not sel) else ""
        chk += ('<label style="display:inline-flex;align-items:center;gap:7px;white-space:nowrap;border:1px solid var(--line);'
                'border-radius:8px;padding:6px 11px;margin:0 8px 8px 0;background:var(--card2);font-size:13px;cursor:pointer">'
                '<input type=checkbox name=cols value="%s" %s onchange="this.form.submit()" style="width:auto;margin:0">'
                '<span><b>T%d</b> · %s <span class=muted>%s</span></span></label>'
                ) % (rid, on, tindex[rid], r.get("label", ""), r.get("date", "")[:10])
    selector = ('<form method=get class=card style="padding:12px 14px;margin-bottom:12px">'
                '<div style="font-weight:700;margin-bottom:8px">Temps affichés</div>'
                '<div style="display:flex;flex-wrap:wrap">%s</div>'
                '<noscript><button class="btn sec sm">Appliquer</button></noscript></form>') % chk
    # entête colonnes (temps)
    th = '<th style="text-align:left;position:sticky;left:0;background:var(--card);z-index:2"></th>'
    for rid, r in recs:
        th += ('<th style="padding:8px;vertical-align:bottom">'
               '<div style="font-weight:700;color:var(--acc)">T%d · %s</div>'
               '<div class=muted style="font-size:12px">%s%s</div></th>') % (
            tindex[rid], r.get("label", ""), (r.get("date", "")[:10]),
            (" · " + age_at(pt.get("dob", ""), r.get("date", ""))) if age_at(pt.get("dob", ""), r.get("date", "")) else "")
    grp = [0]
    def img_row(label, relf):
        grp[0] += 1; gid = "g%d" % grp[0]; cells = ""
        for rid, r in recs:
            if has(rid, relf):
                src = uf(rid, relf)
                cap = ("%s — T%d · %s %s" % (label, tindex[rid], r.get("label", ""), r.get("date", "")[:10])).replace('"', "&quot;")
                inner = ('<a href="%s" class=evlink data-grp="%s" data-cap="%s" onclick="return evOpen(this)" style="cursor:zoom-in">'
                         '<img src="%s" style="%s"></a>') % (src, gid, cap, src, IMGST)
            else:
                inner = '<span class=muted>—</span>'
            cells += '<td style="padding:5px;text-align:center;vertical-align:top">%s</td>' % inner
        return ('<tr><td style="padding:5px;font-weight:600;position:sticky;left:0;background:var(--card);border-right:1px solid var(--line)">%s</td>%s</tr>'
                % (label, cells))
    def section(title):
        return '<tr><td colspan="%d" style="background:var(--card2);color:var(--acc2);font-weight:700;padding:6px 8px;position:sticky;left:0">%s</td></tr>' % (n + 1, title)
    warn = ("<div class='flash' style='background:#fff6e5;border:1px solid #f0d59a;color:#8a5a00'>Il n'y a qu'un seul enregistrement pour ce patient. "
            "Ajoute une <b>réévaluation</b> (fiche patient → « + Nouvelle réévaluation ») pour comparer les temps côte à côte.</div>") if n < 2 else ""
    body = [phead(slug, pt, "evolution")
            + "<p class=muted>Chaque ligne = une vue ; chaque colonne = un temps. Décoche des temps pour comparer seulement ceux qui t'intéressent — les photos s'agrandissent automatiquement.</p>"
            + "%s%s<div style='overflow-x:auto'><table style='%s'>%s"
            % (selector, warn, tstyle, colgroup)]
    body.append("<tr>" + th + "</tr>")
    # réglages d'affichage (quels blocs / quelles vues photo)
    st = get_settings(); blocks = st.get("evo_blocks") or {}
    photos_on = st.get("evo_photos")
    if photos_on is None: photos_on = [k for k, _ in EVO_PHOTOS]
    def on(k): return blocks.get(k, True)
    # PHOTOS
    if on("photos"):
        body.append(section("Photos"))
        for key, lbl in EVO_PHOTOS:
            if key in photos_on:
                body.append(img_row(lbl, "02_photos_traitees/%s.jpg" % key))
    # RADIOS
    if on("radios"):
        body.append(section("Radiographies"))
        for key, lbl in [("panoramique", "Panoramique"), ("teleradiographie_profil", "Téléradiographie de profil")]:
            body.append(img_row(lbl, "03_radios/%s.jpg" % key))
    # MODELES STL
    if on("stl"):
        body.append(section("Modèles 3D"))
        for key, lbl in RENDER_FIELDS:
            body.append(img_row(lbl, "05_stl_rendus/%s.png" % key))
    # STEINER
    if on("steiner"):
        body.append(section("Évolution du Steiner"))
        hdr = '<tr><td style="padding:5px;font-weight:600;position:sticky;left:0;background:var(--card)">Mesure (norme)</td>'
        for rid, r in recs: hdr += '<td style="text-align:center;font-weight:600;padding:5px">T%d</td>' % tindex[rid]
        body.append(hdr + "</tr>")
        for nom, short, mean, sd in STEINER_FIELDS:
            line = '<tr><td style="padding:5px;position:sticky;left:0;background:var(--card)">%s <span class=muted style="font-size:11px">(%g±%g)</span></td>' % (short, mean, sd)
            first = None                                   # 1re valeur mesurée (référence pour Δ)
            for i, (rid, r) in enumerate(recs):
                v = (r.get("steiner") or {}).get(nom)
                if v in (None, ""):
                    line += '<td style="text-align:center;color:#aab">—</td>'; continue
                try:
                    fv = float(v); _t, sev = steiner_signif(nom, fv, mean, sd)
                    col = "#C0392B" if sev else "#1E7D32"
                    if first is None:
                        first = fv; delta = ""
                    else:
                        d = round(fv - first, 1)
                        delta = ('<div style="font-size:10px;color:#8a93a3">%s%g</div>' % ("+" if d > 0 else "", d)) if d else '<div style="font-size:10px;color:#8a93a3">=</div>'
                    line += '<td style="text-align:center;font-weight:600;color:%s">%g%s</td>' % (col, fv, delta)
                except Exception:
                    line += '<td style="text-align:center">%s</td>' % v
            body.append(line + "</tr>")
    body.append("</table></div>")
    # --- visionneuse plein écran : clic sur une photo -> grand format, flèches = défiler dans le temps sur CE type de vue ---
    NAVB = ("position:fixed;top:50%;transform:translateY(-50%);width:56px;height:76px;border:none;border-radius:12px;"
            "background:rgba(255,255,255,.14);color:#fff;font-size:34px;cursor:pointer;z-index:10000")
    lightbox = """
<div id=evlb style="display:none;position:fixed;inset:0;background:rgba(8,12,20,.94);z-index:9999;flex-direction:column;align-items:center;justify-content:center">
  <div id=evcap style="color:#fff;font:600 15px -apple-system,sans-serif;margin-bottom:10px;text-align:center;padding:0 60px"></div>
  <img id=evbig style="max-width:92vw;max-height:78vh;border-radius:10px;box-shadow:0 12px 50px rgba(0,0,0,.7);background:#fff">
  <div id=evpos style="color:#c7d2e0;margin-top:12px;font-size:13px"></div>
  <button onclick="evNav(-1)" title="Precedent" style="__NAVB__;left:2vw">&#8249;</button>
  <button onclick="evNav(1)" title="Suivant" style="__NAVB__;right:2vw">&#8250;</button>
  <button onclick="evClose()" title="Fermer" style="position:fixed;top:16px;right:20px;width:44px;height:44px;border:none;border-radius:10px;background:rgba(255,255,255,.14);color:#fff;font-size:22px;cursor:pointer;z-index:10000">&#10005;</button>
</div>
<script>
var EVG={}, evCur=null, evIdx=0;
(function(){document.querySelectorAll('.evlink').forEach(function(a){var g=a.dataset.grp;(EVG[g]=EVG[g]||[]).push(a);});})();
function evOpen(a){evCur=a.dataset.grp;evIdx=EVG[evCur].indexOf(a);evShow();document.getElementById('evlb').style.display='flex';return false;}
function evShow(){var arr=EVG[evCur],a=arr[evIdx];document.getElementById('evbig').src=a.getAttribute('href');
  document.getElementById('evcap').textContent=a.dataset.cap;
  document.getElementById('evpos').textContent=(evIdx+1)+' / '+arr.length+'   \\u2014   \\u2190 \\u2192 : defiler dans le temps sur cette vue';}
function evNav(d){var arr=EVG[evCur];if(!arr||arr.length<2){return;}evIdx=(evIdx+d+arr.length)%arr.length;evShow();}
function evClose(){document.getElementById('evlb').style.display='none';}
document.addEventListener('keydown',function(e){var lb=document.getElementById('evlb');if(!lb||lb.style.display==='none')return;
  if(e.key==='ArrowLeft'){evNav(-1);}else if(e.key==='ArrowRight'){evNav(1);}else if(e.key==='Escape'){evClose();}});
document.getElementById('evlb').addEventListener('click',function(e){if(e.target.id==='evlb')evClose();});
</script>"""
    body.append(lightbox.replace("__NAVB__", NAVB))
    return page("".join(body), wide=True)

# ======================================================================
#  FICHE ENREGISTREMENT (bilan d'un timepoint) — vue + édition
# ======================================================================
@app.route("/record/<slug>/<rid>")
def record(slug, rid):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r: abort(404)
    base = rdir(slug, rid)
    def u(p): return url_for("rfile", slug=slug, rid=rid, path=p)
    def ex_(p): return os.path.exists(os.path.join(base, p))
    banner = ""
    if r.get("status") == "error":
        banner = "<div class='flash err'>&#9888; Erreur : %s</div>" % r.get("error", "")
    # barre d'actions
    dl = '<a class=btn href="%s">&#128196; Voir le bilan</a> ' % url_for("apercu", slug=slug, rid=rid)
    if r.get("docx"): dl += '<a class="btn sec" href="%s">&#8681; Word</a> ' % u(r["docx"])
    if r.get("pdf"): dl += '<a class="btn sec" href="%s">&#8681; PDF</a> ' % u(r["pdf"])
    dl += '<a class="btn" href="%s" style="background:var(--ink);border-color:var(--ink);color:var(--card)">&#128421; Pr&eacute;senter</a> ' % url_for("presentation", slug=slug, rid=rid)
    dl += '<a class="btn sec" href="%s">&#128229; Import intelligent</a> <a class="btn sec" href="%s">&#128196; Valeurs depuis un Word</a> <a class="btn sec" href="%s">Regenerer</a> <a class="btn sec" href="%s">Ouvrir le dossier</a> <a class="btn sec" href="%s">&#128200; Évolution</a>' % (
        url_for("import_upload", slug=slug, rid=rid), url_for("record_word_values", slug=slug, rid=rid), url_for("regen", slug=slug, rid=rid), url_for("openfolder", slug=slug, rid=rid), url_for("evolution", slug=slug))
    # 1) PHOTOS — emplacements de vues (déposer/assigner) + galerie + visionneuse
    PLBL = dict(PHOTO_FIELDS)
    VIEW_LINES = [("exo_face_repos", "exo_face_sourire", "exo_profil"),
                  ("endo_occlusal_maxillaire", "endo_occlusal_mandibulaire"),
                  ("endo_occlusion_frontale", "endo_laterale_droite", "endo_laterale_gauche")]
    def _slot_src(key):
        for rel in ("02_photos_traitees/%s.jpg" % key, "01_photos_brutes/%s.jpg" % key):
            if ex_(rel):
                try: cb = int(os.path.getmtime(os.path.join(base, rel)))
                except Exception: cb = 0
                return "%s?t=%d" % (u(rel), cb)
        return None
    def _slot_thumb(key):
        for rel in ("02_photos_traitees/%s.jpg" % key, "01_photos_brutes/%s.jpg" % key):
            if ex_(rel):
                try: cb = int(os.path.getmtime(os.path.join(base, rel)))
                except Exception: cb = 0
                return "%s?s=700&t=%d" % (url_for("rthumb", slug=slug, rid=rid, path=rel), cb)
        return None
    def slot_cell(key):
        img = _slot_src(key); disp = _slot_thumb(key) or img; lbl = PLBL.get(key, key)
        crop_url = url_for("crop_editor", slug=slug, rid=rid, slot=key)
        if img:
            del_url = url_for("clear_view", slug=slug, rid=rid, view=key)
            inner = ('<img src="%s" class=slotimg loading=lazy decoding=async onclick="lbShow(\'%s\',null)">'
                     '<a class=slotcrop href="%s" title="Recadrer / orienter">&#9986;</a>'
                     '<form method=post action="%s" style="margin:0">'
                     '<button type=submit class=slotdel title="Retirer cette photo de cette vue">&times;</button></form>'
                     '<div class=slotorient style="position:absolute;left:6px;bottom:6px;display:flex;gap:4px">'
                     '<form method=post action="%s" style="margin:0"><button type=submit title="Pivoter a gauche" style="width:26px;height:26px;padding:0;border:none;border-radius:6px;background:rgba(0,0,0,.55);color:#fff;cursor:pointer;font-size:14px">&#8634;</button></form>'
                     '<form method=post action="%s" style="margin:0"><button type=submit title="Pivoter a droite" style="width:26px;height:26px;padding:0;border:none;border-radius:6px;background:rgba(0,0,0,.55);color:#fff;cursor:pointer;font-size:14px">&#8635;</button></form>'
                     '<form method=post action="%s" style="margin:0"><button type=submit title="Miroir horizontal" style="width:26px;height:26px;padding:0;border:none;border-radius:6px;background:rgba(0,0,0,.55);color:#fff;cursor:pointer;font-size:14px">&#8646;</button></form>'
                     '<form method=post action="%s" style="margin:0"><button type=submit title="Retourner vertical" style="width:26px;height:26px;padding:0;border:none;border-radius:6px;background:rgba(0,0,0,.55);color:#fff;cursor:pointer;font-size:14px">&#8693;</button></form>'
                     '</div>'
                     ) % (disp, img, crop_url, del_url,
                          url_for("orient_slot", slug=slug, rid=rid, slot=key, op="rotl"),
                          url_for("orient_slot", slug=slug, rid=rid, slot=key, op="rotr"),
                          url_for("orient_slot", slug=slug, rid=rid, slot=key, op="mirror"),
                          url_for("orient_slot", slug=slug, rid=rid, slot=key, op="flip"))
        else:
            inner = '<div class=slotempty>Glisser une photo ici</div>'
        return ('<div class=slot data-view="%s" ondragover="event.preventDefault();this.classList.add(\'over\')" '
                'ondragleave="this.classList.remove(\'over\')" ondrop="slotDrop(event,this)">'
                '<div class=slotlbl>%s</div>%s</div>') % (key, lbl, inner)
    def line(keys):
        return '<div class=slotline style="grid-template-columns:repeat(%d,1fr)">%s</div>' % (
            len(keys), "".join(slot_cell(k) for k in keys))
    slots_html = "".join(line(ks) for ks in VIEW_LINES)
    # galerie de toutes les photos importées
    gal_dir = os.path.join(base, "08_galerie"); gthumbs = ""
    if os.path.isdir(gal_dir):
        for f in sorted(os.listdir(gal_dir)):
            if f.lower().endswith((".jpg", ".jpeg", ".png")):
                gsrc = u("08_galerie/%s" % f)
                gth = url_for("rthumb", slug=slug, rid=rid, path="08_galerie/%s" % f)
                gthumbs += ('<img class=galimg loading=lazy decoding=async draggable=true data-file="%s" src="%s" '
                            'ondragstart="event.dataTransfer.setData(\'f\',\'%s\')" '
                            'onclick="lbShow(\'%s\',\'%s\')">') % (f, gth, f, gsrc, f)
    gallery_block = ""
    if gthumbs:
        gallery_block = ('<div style="margin-top:14px"><div style="font-weight:700;margin-bottom:6px">Galerie — toutes les photos '
                         '<span class=muted style="font-weight:400">(glisse une photo vers une vue ci-dessus, ou clique pour l\'agrandir et l\'assigner)</span></div>'
                         '<div class=galgrid>%s</div></div>') % gthumbs
    # boutons d'assignation de la visionneuse (une vue par bouton)
    lb_buttons = "".join('<button type=button class="btn sec sm" onclick="lbAssign(\'%s\')">%s</button>' % (k, l)
                         for k, l in PHOTO_FIELDS)
    assign_action = url_for("assign_photo", slug=slug, rid=rid)
    photos_css = """<style>
    .slotline{display:grid;gap:6px;margin-bottom:6px}
    .slot{border:1px solid var(--line);border-radius:9px;padding:4px;position:relative;min-height:60px;background:#fff}
    .slot.over{border-color:var(--acc);background:#eef4ff}
    .slotlbl{font-size:11.5px;color:var(--mut);text-align:center;margin-bottom:3px}
    .slotimg{width:100%;border-radius:7px;cursor:zoom-in;display:block}
    .slotcrop{position:absolute;top:6px;right:6px;background:rgba(255,255,255,.9);border:1px solid var(--line);border-radius:6px;padding:0 6px;font-size:13px;text-decoration:none}
    .slotdel{position:absolute;top:6px;left:6px;width:24px;height:24px;line-height:1;background:rgba(255,255,255,.92);border:1px solid #f3c0ba;border-radius:6px;color:#C0392B;font-size:17px;font-weight:800;cursor:pointer;padding:0;display:flex;align-items:center;justify-content:center}
    .slotdel:hover{background:#C0392B;color:#fff;border-color:#C0392B}
    .slotempty{display:flex;align-items:center;justify-content:center;height:90px;border:2px dashed #c7d2e5;border-radius:8px;color:#93a3bd;font-size:12.5px;text-align:center}
    .galgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(110px,1fr));gap:8px}
    .galimg{width:100%;height:90px;object-fit:cover;border-radius:8px;border:1px solid var(--line);cursor:grab;background:#fafbfd}
    .galimg:active{cursor:grabbing}
    #lbov{display:none;position:fixed;inset:0;background:rgba(15,25,45,.85);z-index:1000;align-items:center;justify-content:center;flex-direction:column;padding:20px}
    #lbov img{max-width:88vw;max-height:70vh;border-radius:10px;box-shadow:0 10px 40px rgba(0,0,0,.5)}
    #lbbtns{display:flex;flex-wrap:wrap;gap:8px;justify-content:center;margin-top:14px;max-width:900px}
    #lbclose{position:absolute;top:16px;right:22px;color:#fff;font-size:30px;cursor:pointer;text-decoration:none}
    </style>"""
    photos_js = """<script>
    var _lbFile=null;
    function lbShow(src,file){_lbFile=file;document.getElementById('lbimg').src=src;
      document.getElementById('lbassign').style.display=file?'block':'none';
      document.getElementById('lbov').style.display='flex';}
    function lbClose(){document.getElementById('lbov').style.display='none';_lbFile=null;}
    function lbAssign(view){if(_lbFile)doAssign(view,_lbFile);}
    function slotDrop(e,el){e.preventDefault();el.classList.remove('over');
      var f=e.dataTransfer.getData('f');if(f)doAssign(el.dataset.view,f);}
    function doAssign(view,file){var fm=document.getElementById('afForm');
      fm.view.value=view;fm.file.value=file;fm.submit();}
    </script>"""
    photos_html = (photos_css + photos_js
                   + '<form id=afForm method=post action="%s"><input type=hidden name=view><input type=hidden name=file></form>' % assign_action
                   + slots_html + gallery_block
                   + '<div id=lbov onclick="if(event.target.id==\'lbov\')lbClose()"><a id=lbclose onclick="lbClose()">&times;</a>'
                     '<img id=lbimg src=""><div id=lbassign><div style="color:#cfe0ff;text-align:center;margin-top:12px;font-size:13px">Assigner cette photo à :</div>'
                     '<div id=lbbtns>' + lb_buttons + '</div></div></div>')
    def img_line(cells):
        cells = [c for c in cells if c]
        if not cells: return ""
        return '<div style="display:grid;grid-template-columns:repeat(%d,1fr);gap:10px;margin-bottom:12px">%s</div>' % (len(cells), "".join(cells))
    # 2) MODELES 3D — même configuration (occlusales / occlusion ant.+D+G / courbe de Spee)
    RLBL = dict(RENDER_FIELDS)
    def stl_cell(key):
        rel = "05_stl_rendus/%s.png" % key
        if not ex_(rel): return ""
        return ('<div><img src="%s" onclick="lbShow(this.src,\'\')" style="width:100%%;border-radius:8px;border:1px solid var(--line);cursor:zoom-in;display:block">'
                '<div style="display:flex;align-items:center;justify-content:center;gap:7px;margin-top:3px">'
                '<span class=muted>%s</span>'
                '<form method=post action="%s" style="margin:0"><button class="btn sec" style="padding:1px 7px;font-size:12px;line-height:1.4" title="Retourner haut/bas">\u21c5</button></form>'
                '</div></div>') % (u(rel), RLBL.get(key, key), url_for('stl_flip', slug=slug, rid=rid, name=key))
    stl_html = (img_line([stl_cell(k) for k in ("upper_occlusal", "lower_occlusal")])
                + img_line([stl_cell(k) for k in ("front_occ", "right_occ", "left_occ")])
                + img_line([stl_cell(k) for k in ("spee_lower", "spee_left")]))
    has_up = ex_("04_stl_bruts/UpperJawScan.stl"); has_lo = ex_("04_stl_bruts/LowerJawScan.stl")
    has_stl = has_up or has_lo
    n_rend = sum(1 for _k, _l in RENDER_FIELDS if ex_("05_stl_rendus/%s.png" % _k))
    n_exp = 7 if (has_up and has_lo) else (2 if has_stl else 0)
    viewer_btn = ('<a class="btn sec sm" href="%s">&#128260; Voir en 3D (rotation libre)</a>' % url_for("viewer3d", slug=slug, rid=rid)) if has_stl else ""
    if has_stl:
        _glbl = "Générer les aperçus" if n_rend == 0 else "Régénérer les aperçus"
        gen_btn = ('<a class="btn sec sm" href="%s" onclick="this.innerHTML=\'<span class=spin></span> Rendu 3D…\'">%s</a>'
                   % (url_for("render_stl_now", slug=slug, rid=rid), _glbl))
    else:
        gen_btn = ""
    stl_ctrl = (viewer_btn + (" " + gen_btn if gen_btn else "")) or "&nbsp;"
    if not has_stl:
        stl_body = "<span class=muted>Pas de modèle 3D importé.</span>"
    elif n_rend == 0:
        stl_body = ('<div class=flash style="background:var(--accw);border-color:var(--line);margin:0">'
                    '<b>Modèle 3D importé</b>, mais les aperçus 2D ne sont pas encore générés. '
                    'Clique sur <b>Générer les aperçus</b> (quelques secondes) — ou ouvre <b>Voir en 3D</b> pour la rotation libre.</div>')
    else:
        _warn = ('<div class=flash style="background:#fff7e6;border-color:#f0d9a8;margin:0 0 10px">'
                 'Aperçus incomplets (%d/%d) — clique sur <b>Régénérer les aperçus</b>.</div>' % (n_rend, n_exp)) if n_rend < n_exp else ""
        stl_body = _warn + stl_html
    # 2b) CAPTURES DES MODELES (extraites d'un bilan Word, quand il n'y a pas de STL)
    cap_dir = os.path.join(base, "07_modele_captures")
    cap_cells = []
    if os.path.isdir(cap_dir):
        for f in sorted(os.listdir(cap_dir)):
            if f.lower().endswith((".jpg", ".jpeg", ".png")):
                rel = "07_modele_captures/%s" % f
                cap_cells.append('<div><a href="%s" target=_blank><img src="%s" style="width:100%%;border-radius:8px;border:1px solid var(--line)"></a></div>' % (u(rel), u(rel)))
    cap_html = (('<div class=card><h2>Captures des modèles</h2>'
                 '<p class=muted style="margin:0 0 10px">Captures d\'écran des modèles 3D extraites du bilan Word (aucun fichier STL disponible).</p>'
                 '<div class="grid g3">%s</div></div>') % "".join(cap_cells)) if cap_cells else ""
    # galerie « toutes les photos de la séance » (import massif : rien n'est perdu)
    gal_dir = os.path.join(base, "08_galerie"); gal_cells = []
    if os.path.isdir(gal_dir):
        for f in sorted(os.listdir(gal_dir)):
            if f.lower().endswith((".jpg", ".jpeg", ".png")):
                rel = "08_galerie/%s" % f
                gal_cells.append('<div><a href="%s" target=_blank><img src="%s" style="width:100%%;border-radius:8px;border:1px solid var(--line)"></a></div>' % (u(rel), u(rel)))
    gal_html = (('<div class=card><details><summary style="cursor:pointer;font-weight:600;font-size:16px;color:var(--acc)">Toutes les photos de la séance (%d) — non classées</summary>'
                 '<p class=muted style="margin:8px 0 10px">Photos importées non rangées automatiquement. Tu peux les consulter ici ; le classement par vue reste modifiable.</p>'
                 '<div class="grid g4">%s</div></details></div>') % (len(gal_cells), "".join(gal_cells))) if gal_cells else ""
    # 3) RADIOS
    rad = ""
    for key, lbl in [("panoramique", "Panoramique"), ("teleradiographie_profil", "Teleradiographie de profil")]:
        rel = "03_radios/%s.jpg" % key
        if ex_(rel): rad += ('<div style="position:relative"><img src="%s" onclick="lbShow(this.src,\'\')" style="width:100%%;border-radius:8px;border:1px solid #e3e8ef;cursor:zoom-in;display:block">' '<form method=post action="%s" style="margin:0;position:absolute;top:8px;right:8px"><button type=submit title="Supprimer cette radio" style="width:30px;height:30px;border:none;border-radius:8px;background:rgba(200,40,40,.9);color:#fff;font-size:17px;line-height:1;cursor:pointer">&times;</button></form>' '<div class=muted style=text-align:center>%s</div>' '<form method=post action="%s" style="margin:3px 0 0;text-align:center"><button class="btn sec" type=submit style="font-size:11.5px;padding:3px 9px" title="Reclasser cette radio">&#8646; Reclasser en %s</button></form>' '</div>') % (u(rel), url_for("record_radio_delete", slug=slug, rid=rid, key=key), lbl, url_for("record_radio_retype", slug=slug, rid=rid, key=key), ("Téléradiographie de profil" if key == "panoramique" else "Panoramique"))
    # 4) ANALYSE RADIO (Steiner)
    steiner_graph = ('<div style="margin-bottom:14px">%s</div>' % steiner_table_html(r)) if r.get("steiner") else ""
    steiner_inputs = "".join('<div><label>%s</label><input name="st_%s" inputmode=decimal value="%s" autocomplete=off></div>'
        % (short, nom, (("%g" % r["steiner"][nom]) if nom in r.get("steiner", {}) else "")) for nom, short, mean, sd in STEINER_FIELDS)
    # 5) SYNTHESE
    syn = r.get("synthese", {}); ex = r.get("exam", {}); ang = r.get("angle", {})
    # 6) EXAMEN + MDC
    def exsec(title, fields, data):
        inner = "".join('<div><label>%s</label>%s</div>' % (lbl, field_input("ex_" + k, data.get(k, ""), opts)) for k, lbl, opts in fields)
        return '<h3 style="margin:14px 0 6px;color:var(--acc)">%s</h3><div class="grid g3">%s</div>' % (title, inner)
    angle_html = ('<h3 style="margin:14px 0 6px;color:var(--acc)">Classe d\'Angle dentaire</h3><div class="grid g4">'
                  + "".join('<div><label>%s</label>%s</div>' % (lbl, field_input("ang_" + k, ang.get(k, ""), opts)) for k, lbl, opts in ANGLE_FIELDS) + '</div>')
    motif_val = (r.get("motif", "") or "").replace('"', "&quot;")
    # 7) MODIF fichiers
    photo_re = "".join('<div><label>%s</label><input type=file name="photo_%s" accept="image/*"></div>' % (l, k) for k, l in PHOTO_FIELDS)
    # 8) lien vers le SUIVI de traitement (état actuel + commentaires, gérés à part)
    suivi_card = ('<div class=card><h2 style="margin:0 0 6px">Suivi de traitement</h2>'
                  '<p class=muted style="margin:0 0 10px">L\'état actuel (appareil, arcs en place…) et les commentaires de séances '
                  'se gèrent dans la section Suivi du patient.</p>'
                  '<a class="btn sec" href="%s">🦷 Ouvrir le suivi de traitement</a> '
                  '<a class="btn sec" href="%s">✉️ Courrier au correspondant</a></div>') % (url_for("suivi", slug=slug), url_for("courrier", slug=slug, rid=rid))

    form = ('<form method=post action="%s" enctype=multipart/form-data onsubmit="this.querySelector(\'button[type=submit]\').innerHTML=\'<span class=spin></span> Enregistrement...\'">'
      '<div class=card id=steiner><h2 style="margin:0 0 6px">Analyse radiographique (Steiner)</h2>%s'
        '<details style="margin-top:10px"><summary style="cursor:pointer;font-weight:600;color:var(--acc)">Modifier les valeurs de la céphalométrie</summary>'
        '<div class="grid g3" style="margin-top:10px">%s</div></details></div>'
      '<div class=card><h2>Synthese diagnostique</h2>%s</div>'
      '%s'
      '%s'
      '<div class=card><details><summary style="cursor:pointer;font-weight:600;font-size:16px;color:var(--acc)">Examen clinique &amp; motif de consultation</summary>'
        '<div style="margin-top:12px"><label>Motif de consultation</label><input name=motif value="%s" placeholder="Motif..." style="width:100%%">'
        '%s%s%s</div></details></div>'
      '<div class=card><details><summary style="cursor:pointer;font-weight:600;font-size:16px;color:var(--acc)">Modification des elements (photos / radios / STL)</summary>'
        '<p class=muted style=margin:8px_0>Depose un fichier seulement pour le remplacer, puis Enregistrer.</p>'
        '<div class="grid g4">%s</div><div class="grid g2" style="margin-top:10px">'
          '<div><label>Panoramique</label><input type=file name=radio_panoramique accept="image/*"></div>'
          '<div><label>Teleradiographie</label><input type=file name=radio_teleradiographie_profil accept="image/*"></div>'
          '<div><label>STL superieur</label><input type=file name=stl_UpperJawScan accept=".stl"></div>'
          '<div><label>STL inferieur</label><input type=file name=stl_LowerJawScan accept=".stl"></div></div></details></div>'
      '<button type=submit class=btn>Enregistrer &amp; regenerer</button></form>') % (
        url_for("edit_record", slug=slug, rid=rid), steiner_graph, steiner_inputs, diag_inputs(syn),
        plan_card(r), suivi_card,
        motif_val, exsec("Environnement", EXAM_ENV, ex), exsec("Face", EXAM_FACE, ex), angle_html, photo_re)

    _wc_dir = os.path.join(base, "06_webceph"); _tcells = []
    if os.path.isdir(_wc_dir):
        for _tf in sorted(os.listdir(_wc_dir)):
            if _tf.lower().endswith((".jpg", ".jpeg", ".png")):
                _tsrc = u("06_webceph/%s" % _tf)
                _tcells.append(('<div style="position:relative"><img src="%s" onclick="lbShow(this.src,\'\')" style="width:100%%;border-radius:8px;border:1px solid #e3e8ef;cursor:zoom-in;display:block">'
                    '<form method=post action="%s" style="margin:0;position:absolute;top:8px;right:8px"><button type=submit title="Supprimer ce trace" style="width:30px;height:30px;border:none;border-radius:8px;background:rgba(200,40,40,.9);color:#fff;font-size:17px;line-height:1;cursor:pointer">&times;</button></form></div>')
                    % (_tsrc, url_for("record_trace_delete", slug=slug, rid=rid, name=_tf)))
    trace_html = "".join(_tcells) if _tcells else "<span class=muted>-</span>"
    body = ("%s<p class=muted><a href=\"%s\">&#8592; %s</a></p>"
      "<h1>%s <span class=tag>%s</span></h1><h2>%s &middot; %s &middot; %s</h2>"
      "<div class=card>%s</div>"
      "<div class=\"card cardphotos\"><h2>Photos</h2>%s</div>"
      "<div class=card><div style=\"display:flex;justify-content:space-between;align-items:center;margin-bottom:14px\"><h2 style=margin:0>Modeles 3D</h2>%s</div>%s</div>"
      "%s%s"
      "<div class=card><h2>Radiographies</h2><div class=\"grid g2\">%s</div></div>"
      "<div class=card id=trace><h2>Tracé céphalométrique</h2><div class=\"grid g2\">%s</div></div>"
      "<div class=card id=odonto><h2>Odontogramme</h2>%s</div>"
      "%s") % (
        banner, url_for("patient", slug=slug), pt["nom"], pt["nom"], r.get("label", ""),
        pt["age"], pt["sexe"], r.get("date", "")[:16].replace("T", " "),
        dl, photos_html or "<span class=muted>-</span>",
        stl_ctrl, stl_body,
        cap_html, gal_html,
        rad or "<span class=muted>-</span>", trace_html, odonto_card(slug, rid, r), form)
    return page(body, wide=True)

@app.route("/record/<slug>/<rid>/trace_delete/<name>", methods=["POST"])
def record_trace_delete(slug, rid, name):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r: abort(404)
    fp = os.path.join(rdir(slug, rid), "06_webceph", os.path.basename(name))
    if os.path.exists(fp):
        try: os.remove(fp); flash("Trace supprime.")
        except Exception: pass
    return redirect(url_for("record", slug=slug, rid=rid) + "#trace")


STEINER_WEBCEPH_MAP = {
    "SNA": "SNA (°)", "SNB": "SNB (°)", "ANB": "ANB (°)",
    "U1 to NA(deg)": "U1-NA (°)", "U1 to NA(mm)": "U1-NA (mm)",
    "L1 to NB(deg)": "L1-NB (°)", "L1 to NB(mm)": "L1-NB (mm)",
    "Wits appraisal": "Wits (mm)", "Pog to NB": "Pog-NB (mm)",
    "Mandibular plane angle(Go-Gn to SN)": "Plan mandibulaire (Go-Gn/SN) (°)",
    "Occlusal plane to SN angle": "Plan occlusal/SN (°)",
    "Interincisal angle": "Angle interincisif (°)",
}

def _steiner_queue_dir():
    return os.path.join(DATA, "_steiner_inbox")


def _steiner_queue_list():
    import glob as _g, json as _j
    qd = _steiner_queue_dir(); out = []
    if os.path.isdir(qd):
        for fp in sorted(_g.glob(os.path.join(qd, "*.json")), key=lambda f: os.path.getmtime(f), reverse=True):
            try:
                d = _j.load(open(fp, encoding="utf-8"))
                out.append({"file": os.path.basename(fp), "patient": (d.get("patient") or "").strip(), "vals": d.get("vals", {}) or {}})
            except Exception:
                pass
    return out


@app.route("/steiner")
def steiner_queue():
    if not logged(): return redirect(url_for("login"))
    import json as _j
    items = _steiner_queue_list()
    recs = {}; names = {}
    for slug in sorted(os.listdir(PATIENTS)):
        pdir = os.path.join(PATIENTS, slug); rd = os.path.join(pdir, "records")
        if not os.path.isdir(rd): continue
        try: nom = _j.load(open(os.path.join(pdir, "patient.json"), encoding="utf-8")).get("nom", slug)
        except Exception: nom = slug
        lst = []
        for rid in sorted(os.listdir(rd)):
            try: lbl = _j.load(open(os.path.join(rd, rid, "meta.json"), encoding="utf-8")).get("label", rid)
            except Exception: lbl = rid
            lst.append({"rid": rid, "label": lbl})
        if lst: recs[slug] = lst; names[slug] = nom
    popts = "".join('<option value="%s">%s</option>' % (sl, names[sl]) for sl in sorted(names, key=lambda x: names[x].lower()))
    _STYLE = '<style>.rcard{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:18px 20px;margin-bottom:16px;box-shadow:var(--sh-sm)}.rcard table{border-collapse:collapse}.rempty{padding:26px;text-align:center;color:var(--mut);background:var(--card2);border:1px dashed var(--line);border-radius:12px}.rcard label{font-size:12px;color:var(--mut)}.rcard select{padding:8px 10px;border:1px solid var(--line);border-radius:9px;min-width:190px;background:var(--card)}</style>'
    _HEAD = '<div class=phead><h1>Analyses Steiner à classer</h1><div class=sub>Chaque capture WebCeph attend ici avec le nom du patient — choisis la fiche où l\'importer.</div></div>'
    cards = []
    for it in items:
        rows = ""
        for lab, nom in STEINER_WEBCEPH_MAP.items():
            if lab in it["vals"]:
                v = it["vals"][lab]; vs = ("%g" % v) if isinstance(v, (int, float)) else str(v)
                rows += '<tr><td style="padding:1px 14px 1px 0;color:var(--mut)">%s</td><td style="padding:1px 0;font-weight:600">%s</td></tr>' % (nom, vs)
        nval = len([1 for lab in it["vals"] if lab in STEINER_WEBCEPH_MAP])
        pt_name = (it["patient"] or "Patient inconnu")
        card = ('<div class=rcard>'
                '<div style="font-weight:800;font-size:16px;color:var(--ink);margin-bottom:2px">%s <span class=muted style="font-weight:600;font-size:13px">· %d valeurs</span></div>'
                '<table style="margin:8px 0 12px;font-size:13px"><tbody>%s</tbody></table>'
                '<form method=post action="%s" style="display:flex;gap:12px;flex-wrap:wrap;align-items:flex-end">'
                '<input type=hidden name=file value="%s">'
                '<div><label>Patient</label><br><select name=slug onchange="fillQ(this)"><option value="">—</option>%s</select></div>'
                '<div><label>Réévaluation</label><br><select name=rid></select></div>'
                '<button class=btn type=submit>Importer dans cette fiche</button>'
                '<button class="btn sec" type=submit formaction="%s">Supprimer</button>'
                '</form></div>') % (pt_name, nval, rows, url_for("steiner_queue_import"), it["file"], popts, url_for("steiner_queue_delete"))
        cards.append(card)
    body_cards = "".join(cards) if cards else '<div class=rempty>Aucune analyse en attente.<br>Ouvre WebCeph, affiche l\'analyse Steiner (Analyse + Graphique) et clique « Importer valeurs Steiner ».</div>'
    _JS = ('<script>var RECS=' + _j.dumps(recs) + ';\n'
           'function fillQ(sel){var form=sel.form;var rid=form.querySelector("select[name=rid]");rid.innerHTML="";(RECS[sel.value]||[]).forEach(function(r){var o=document.createElement("option");o.value=r.rid;o.textContent=r.label;rid.appendChild(o);});}\n'
           '</script>')
    return page(_HEAD + _STYLE + body_cards + _JS)


@app.route("/steiner/import", methods=["POST"])
def steiner_queue_import():
    if not logged(): return redirect(url_for("login"))
    import json as _j
    fn = os.path.basename(request.form.get("file", ""))
    slug = request.form.get("slug", ""); rid = request.form.get("rid", "")
    fp = os.path.join(_steiner_queue_dir(), fn)
    if not fn or not os.path.exists(fp):
        flash("Capture introuvable (déjà classée ?)."); return redirect(url_for("steiner_queue"))
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r:
        flash("Choisis un patient et une réévaluation."); return redirect(url_for("steiner_queue"))
    try:
        d = _j.load(open(fp, encoding="utf-8")); vals = d.get("vals", {}) or {}
    except Exception:
        flash("Capture illisible."); return redirect(url_for("steiner_queue"))
    st = dict(r.get("steiner", {}) or {}); n = 0
    for lab, v in vals.items():
        nom = STEINER_WEBCEPH_MAP.get(lab)
        if nom and isinstance(v, (int, float)): st[nom] = v; n += 1
    if n:
        r["steiner"] = st; save_rec(slug, rid, r)
        try: _regenerate(slug, rid, pt, r, skip_stl=True)
        except Exception: pass
        try: os.remove(fp)
        except Exception: pass
        flash("%d valeurs Steiner importées dans %s." % (n, pt.get("nom") or slug))
    else:
        flash("Aucune valeur reconnue.")
    return redirect(url_for("steiner_queue"))


@app.route("/steiner/delete", methods=["POST"])
def steiner_queue_delete():
    if not logged(): return redirect(url_for("login"))
    fn = os.path.basename(request.form.get("file", ""))
    fp = os.path.join(_steiner_queue_dir(), fn)
    if fn and os.path.exists(fp):
        try: os.remove(fp); flash("Capture supprimée.")
        except Exception: pass
    return redirect(url_for("steiner_queue"))


@app.route("/record/<slug>/<rid>/steiner_from_webceph")
def steiner_from_webceph(slug, rid):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r: abort(404)
    fp = os.path.join(DATA, "_steiner_inbox.json")
    if not os.path.exists(fp):
        flash("Aucune valeur WebCeph en attente. Sur WebCeph (onglet Analyse + Graphique), clique \u00ab Importer valeurs Steiner \u00bb d'abord.")
        return redirect(url_for("record", slug=slug, rid=rid))
    try:
        payload = json.load(open(fp, encoding="utf-8"))
        vals = payload.get("vals", {}) if isinstance(payload, dict) else {}
    except Exception:
        flash("Import WebCeph illisible."); return redirect(url_for("record", slug=slug, rid=rid))
    st = dict(r.get("steiner", {}) or {}); n = 0; done = []
    for lab, v in vals.items():
        nom = STEINER_WEBCEPH_MAP.get(lab)
        if nom and isinstance(v, (int, float)):
            st[nom] = v; n += 1; done.append(nom)
    if n:
        r["steiner"] = st; save_rec(slug, rid, r)
        try: _regenerate(slug, rid, pt, r, skip_stl=True)
        except Exception: pass
        flash("%d valeurs Steiner import\u00e9es depuis WebCeph." % n)
    else:
        flash("Aucune valeur reconnue dans l'import WebCeph.")
    return redirect(url_for("record", slug=slug, rid=rid) + "#steiner")


@app.route("/record/<slug>/<rid>/radio_delete/<key>", methods=["POST"])
def record_radio_delete(slug, rid, key):
    if not logged(): return redirect(url_for("login"))
    if key not in ("panoramique", "teleradiographie_profil"): abort(404)
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r: abort(404)
    fp = os.path.join(rdir(slug, rid), "03_radios", key + ".jpg")
    if os.path.exists(fp):
        try: os.remove(fp)
        except Exception: pass
        try: _regenerate(slug, rid, pt, r, skip_stl=True)
        except Exception: pass
        flash("Radio supprimee.")
    return redirect(url_for("record", slug=slug, rid=rid) + "#radios")


@app.route("/record/<slug>/<rid>/render_stl")
def render_stl_now(slug, rid):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r: abort(404)
    base = rdir(slug, rid)
    up = os.path.join(base, "04_stl_bruts", "UpperJawScan.stl")
    lo = os.path.join(base, "04_stl_bruts", "LowerJawScan.stl")
    if not (os.path.exists(up) or os.path.exists(lo)):
        flash("Aucun modèle STL sur cette réévaluation."); return redirect(url_for("record", slug=slug, rid=rid))
    try:
        if getattr(sys, "frozen", False):
            cmd = [sys.executable, "render_stl", base]
        else:
            cmd = [sys.executable, os.path.join(HERE, "pipeline.py"), "render_stl", base]
        _rd = os.path.join(base, "05_stl_rendus")
        _want = "upper_occlusal.png" if os.path.exists(up) else "lower_occlusal.png"
        res = None
        for _a in range(3):
            res = subprocess.run(cmd, timeout=180, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
            if res.returncode == 0 and os.path.exists(os.path.join(_rd, _want)): break
            import time as _t; _t.sleep(1.5)
        if res is not None and res.returncode == 0 and os.path.exists(os.path.join(_rd, _want)):
            flash("Aperçus 3D générés.")
        else:
            flash("Échec du rendu 3D : %s" % ((res.stderr or b"").decode("utf-8", "ignore")[-300:] if res else "erreur"))
    except Exception as e:
        flash("Échec du rendu 3D : %s" % e)
    return redirect(url_for("record", slug=slug, rid=rid))


@app.route("/record/<slug>/<rid>/stl_flip/<name>", methods=["POST"])
def stl_flip(slug, rid, name):
    if not logged(): return redirect(url_for("login"))
    if name not in {k for k, _ in RENDER_FIELDS}: abort(404)
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r: abort(404)
    OUT = os.path.join(rdir(slug, rid), "05_stl_rendus")
    png = os.path.join(OUT, name + ".png")
    marker = os.path.join(OUT, ".flip_" + name)
    try:
        from PIL import Image as _I, ImageOps as _IO
        os.makedirs(OUT, exist_ok=True)
        if os.path.exists(png):
            _IO.flip(_I.open(png)).save(png)
        if os.path.exists(marker):
            os.remove(marker)
        else:
            open(marker, "w").close()
    except Exception:
        pass
    return redirect(request.referrer or url_for("record", slug=slug, rid=rid))


@app.route("/record/<slug>/<rid>/radio_retype/<key>", methods=["POST"])
def record_radio_retype(slug, rid, key):
    if not logged(): return redirect(url_for("login"))
    if key not in ("panoramique", "teleradiographie_profil"): abort(404)
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r: abort(404)
    rd = os.path.join(rdir(slug, rid), "03_radios")
    other = "teleradiographie_profil" if key == "panoramique" else "panoramique"
    src = os.path.join(rd, key + ".jpg"); dst = os.path.join(rd, other + ".jpg")
    if not os.path.exists(src):
        flash("Radio introuvable."); return redirect(url_for("record", slug=slug, rid=rid))
    try:
        if os.path.exists(dst):
            tmp = os.path.join(rd, "_swap_tmp.jpg")
            os.replace(src, tmp); os.replace(dst, src); os.replace(tmp, dst)
            flash("Les deux radios ont ete interverties.")
        else:
            os.replace(src, dst)
            flash("Radio reclassee en %s." % ("Teleradiographie de profil" if other == "teleradiographie_profil" else "Panoramique"))
        try: _regenerate(slug, rid, pt, r, skip_stl=True)
        except Exception: pass
    except Exception as e:
        flash("Echec du changement de type : %s" % e)
    return redirect(url_for("record", slug=slug, rid=rid) + "#radios")


@app.route("/record/<slug>/<rid>/apercu")
def apercu(slug, rid):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r: abort(404)
    base = rdir(slug, rid)
    body = ('<p class=muted><a href="%s">&#8592; Retour a la fiche</a></p><h1>Apercu du bilan</h1>%s'
            % (url_for("record", slug=slug, rid=rid), apercu_html(slug, rid, pt, r, base)))
    return page(body)

@app.route("/record/<slug>/<rid>/edit", methods=["POST"])
def edit_record(slug, rid):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r: abort(404)
    base = rdir(slug, rid)
    save_uploads(base, request.files)
    r["steiner"] = parse_steiner_form(request.form)
    r["synthese"] = parse_synthese(request.form)
    r["exam"] = parse_exam(request.form)
    r["angle"] = parse_angle(request.form)
    r["motif"] = request.form.get("motif", "").strip()
    r["plan"] = parse_plan(request.form)
    # Un STL ajouté rend la VISIONNEUSE 3D (three.js) disponible immédiatement — pas besoin
    # du rendu VTK (lourd, 420 s, absent du .app) qui faisait « planter » l'enregistrement.
    # On régénère le Word/photos sans forcer le rendu 3D ; les captures existantes sont conservées.
    save_rec(slug, rid, r)
    _regenerate(slug, rid, pt, r, skip_stl=True)
    flash("Enregistré.")
    return redirect(url_for("record", slug=slug, rid=rid))

@app.route("/record/<slug>/<rid>/regen")
def regen(slug, rid):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r: abort(404)
    _regenerate(slug, rid, pt, r); flash("Bilan régénéré.")
    return redirect(url_for("record", slug=slug, rid=rid))

@app.route("/record/<slug>/<rid>/rotate/<slot>")
def rotate(slug, rid, slot):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r: abort(404)
    base = rdir(slug, rid)
    # rotation rapide de 90° : on tourne UNIQUEMENT cette photo (pas de re-génération Word/3D)
    c = r.setdefault("crops", {}).get(slot) or {"rot": 0, "fine": 0, "mirror": False,
                                                "flip": False, "x": 0, "y": 0, "w": 1, "h": 1}
    c["rot"] = (int(c.get("rot", 0) or 0) + 90) % 360
    r["crops"][slot] = c
    _render_slot(base, slot, c)
    save_rec(slug, rid, r)
    return redirect(url_for("record", slug=slug, rid=rid) + "#photos")

def _regenerate(slug, rid, pt, r, skip_stl=False):
    base = rdir(slug, rid)
    docx, pdf, log, err = run_generation(base, build_info(pt, r), r.get("steiner", {}),
                                         r.get("rotations", {}), r.get("crops", {}), skip_stl)
    r.update({"docx": os.path.basename(docx) if docx else None, "pdf": os.path.basename(pdf) if pdf else None,
              "log": log, "status": "error" if err else "done", "error": err})
    save_rec(slug, rid, r)

# ---- visualiseur 3D interactif ----
VIEWER_HTML = """<!doctype html><html lang=fr><head><meta charset=utf-8>
<meta name=viewport content="width=device-width,initial-scale=1"><title>Modèle 3D</title>
<style>html,body{margin:0;height:100%;overflow:hidden;font:14px -apple-system,sans-serif}
#c{position:fixed;inset:0;background:linear-gradient(#e0eaf7,#4f76b8)}
#bar{position:fixed;top:10px;left:10px;z-index:10;display:flex;gap:8px;align-items:center}
.b{background:rgba(255,255,255,.92);border:1px solid #cdd;border-radius:8px;padding:6px 12px;cursor:pointer;font-size:13px;color:#2F5CA8;text-decoration:none}
.b:hover{background:#fff}#hint{position:fixed;bottom:12px;left:0;right:0;text-align:center;color:#fff;opacity:.85;z-index:10;text-shadow:0 1px 2px rgba(0,0,0,.3)}
#err{position:fixed;inset:0;display:none;align-items:center;justify-content:center;color:#fff;text-align:center;padding:30px;z-index:20}
#measbtn.on{background:#ff3b30;color:#fff;border-color:#ff3b30}
#anglebtn.on{background:#1e8ad0;color:#fff;border-color:#1e8ad0}
#meas{position:fixed;top:52px;left:10px;z-index:10;background:rgba(255,255,255,.96);border:1px solid #cdd;border-radius:10px;padding:9px 12px;min-width:150px;display:none;box-shadow:0 4px 16px rgba(0,0,0,.15)}
#meas h4{margin:0 0 6px;font-size:12px;color:#2F5CA8}
.mrow{display:flex;justify-content:space-between;gap:14px;font-size:13px;padding:2px 0;color:#2a3b52}
#anabtn.on{background:#2F5CA8;color:#fff;border-color:#2F5CA8}
#ana{position:fixed;top:52px;right:10px;z-index:10;background:rgba(255,255,255,.97);border:1px solid #cdd;border-radius:10px;padding:10px 12px;width:252px;max-height:82vh;overflow:auto;display:none;box-shadow:0 4px 16px rgba(0,0,0,.15)}
#ana h4{margin:0;font-size:13px;color:#2F5CA8}
#anatabs .bs{padding:4px 8px;font-size:12px;background:rgba(255,255,255,.92);border:1px solid #cdd;border-radius:7px;cursor:pointer;color:#2F5CA8}
#anatabs .bs.on{background:#2F5CA8;color:#fff;border-color:#2F5CA8}
.anag{font-size:11px;color:#8794a8;text-transform:uppercase;letter-spacing:.03em;margin:8px 0 3px}
.arow{display:flex;justify-content:space-between;gap:10px;font-size:13px;padding:4px 7px;border-radius:6px;cursor:pointer;border:1px solid transparent}
.arow:hover{background:#f2f6fc}.arow.armed{background:#eef4ff;border-color:#2F5CA8}.arow b{color:#16283d}
.ares{margin-top:10px;padding:9px 10px;background:#f4f8ff;border:1px solid #d7e3f5;border-radius:8px;font-size:12.5px;line-height:1.55;color:#24405f}
.nrm{color:#8794a8;font-weight:400}
</style></head><body>
<div id=c></div>
<div id=bar><a class=b href="__BACK__">← Retour</a>
<button class=b onclick="setView('both')">Haut+Bas</button>
<button class=b onclick="setView('upper')">Maxillaire</button>
<button class=b onclick="setView('lower')">Mandibule</button>
<button class=b onclick="resetCam()">Recentrer</button><button class=b id=measbtn onclick="toggleMeasure()">📏 Mesurer</button><button class=b id=anglebtn onclick="toggleAngle()">📐 Angle</button><button class=b onclick="undoLast()">↶ Annuler</button><button class=b onclick="clearMeasure()">Effacer</button><button class=b id=anabtn onclick="toggleAna()">🧮 Analyses</button><button class=b id=savebtn onclick="saveMeasures()">💾 Enregistrer</button></div>
<div id=meas><h4>Mesures (mm)</h4><div id=measlist></div></div>
<div id=ana><div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px"><h4>Analyses cliniques</h4><a onclick="clearAna()" style="font-size:11px;color:#c0392b;cursor:pointer;text-decoration:none">réinit.</a></div><div id=anatabs style="display:flex;gap:5px;flex-wrap:wrap;margin-bottom:6px"><button class=bs data-a=bolton_ant onclick="selAna('bolton_ant')">Bolton ant.</button><button class=bs data-a=bolton_tot onclick="selAna('bolton_tot')">Bolton total</button><button class=bs data-a=pont onclick="selAna('pont')">Pont</button><button class=bs data-a=ddm onclick="selAna('ddm')">Espace (DDM)</button></div><div id=anabody><div class=nrm style="font-size:12px">Choisis une analyse ci-dessus, puis mesure chaque dent (clic 2 points sur les faces mésiale et distale).</div></div></div>
<div id=hint>Fais glisser pour tourner • molette pour zoomer • clic droit pour déplacer</div>
<div id=err><div><b>Modèle 3D non disponible</b><br>Le modèle 3D n'a pas pu être chargé.</div></div>
<script type="importmap">{"imports":{"three":"/lib/three.module.js","three/addons/":"/lib/addons/"}}</script>
<script type="module">
import * as THREE from 'three';
import { OrbitControls } from 'three/addons/controls/OrbitControls.js';
import { STLLoader } from 'three/addons/loaders/STLLoader.js';
const UP="__UPURL__", LO="__LOURL__";
const el=document.getElementById('c');
const scene=new THREE.Scene();
const camera=new THREE.PerspectiveCamera(35, window.innerWidth/window.innerHeight, 1, 5000);
const renderer=new THREE.WebGLRenderer({antialias:true, alpha:true});
renderer.setPixelRatio(window.devicePixelRatio); renderer.setSize(window.innerWidth, window.innerHeight);
el.appendChild(renderer.domElement);
scene.add(new THREE.HemisphereLight(0xffffff, 0x888899, 1.05));
const d1=new THREE.DirectionalLight(0xffffff,0.55); d1.position.set(1,1,1); scene.add(d1);
const d2=new THREE.DirectionalLight(0xffffff,0.35); d2.position.set(-1,0.5,-0.5); scene.add(d2);
const mat=new THREE.MeshStandardMaterial({color:0xdcdcdc, roughness:0.55, metalness:0.0, side:THREE.DoubleSide});
const controls=new OrbitControls(camera, renderer.domElement);
controls.enableDamping=true; controls.rotateSpeed=0.9;
const group=new THREE.Group(); scene.add(group);
const meshes={};
const loader=new STLLoader();
let loaded=0, toLoad=0, center=new THREE.Vector3(), radius=100;
function fit(){
  const box=new THREE.Box3().setFromObject(group);
  if(box.isEmpty()) return;
  box.getCenter(center); const size=box.getSize(new THREE.Vector3()); radius=Math.max(size.x,size.y,size.z);
  resetCam();
}
function resetCam(){
  camera.position.set(center.x, center.y+radius*0.15, center.z+radius*2.4);
  controls.target.copy(center); controls.update();
}
function add(url,name){ if(!url) return; toLoad++;
  loader.load(url, g=>{ const m=new THREE.Mesh(g,mat); group.add(m); meshes[name]=m; loaded++; if(loaded===toLoad) fit(); },
   undefined, e=>{ document.getElementById('err').style.display='flex'; });
}
window.setView=function(v){ if(meshes.upper) meshes.upper.visible=(v!=='lower'); if(meshes.lower) meshes.lower.visible=(v!=='upper'); };
window.resetCam=resetCam;
add(UP,'upper'); add(LO,'lower');
addEventListener('resize', ()=>{ camera.aspect=window.innerWidth/window.innerHeight; camera.updateProjectionMatrix(); renderer.setSize(window.innerWidth,window.innerHeight); });
// ---- MESURE & ANALYSES 3D (mm ; les STL sont en mm) ----
const measGroup=new THREE.Group(); scene.add(measGroup);
const raycaster=new THREE.Raycaster(), mouse=new THREE.Vector2();
let downXY=null, mode='off', drag=null;
let freeData=[], _fp=[], _ap=[], chainPts=[], chainTmp=[], chainSum=0;
const units=[];
let anaKey=null; const anaSaved={}; const SAVEURL="__SAVEURL__";
window.__anaOpen=false; window.__armedRow=null; window.__chainMode=false;
function mkMarker(p){ const s=new THREE.Mesh(new THREE.SphereGeometry(Math.max(radius*0.009,0.2),16,16), new THREE.MeshBasicMaterial({color:0xff3b30,depthTest:false})); s.renderOrder=999; s.position.copy(p); measGroup.add(s); return s; }
function mkLine(a,b,c){ const l=new THREE.Line(new THREE.BufferGeometry().setFromPoints([a,b]), new THREE.LineBasicMaterial({color:c,depthTest:false})); l.renderOrder=998; measGroup.add(l); return l; }
function rm(o){ measGroup.remove(o); if(o.geometry)o.geometry.dispose(); if(o.material){ if(o.material.map)o.material.map.dispose(); o.material.dispose(); } }
function _V(a){ return new THREE.Vector3(a[0],a[1],a[2]); }
function makeUnit(kind, markers, meta){ meta=meta||{};
  const u={kind:kind, markers:markers, lines:[], pts:[], slot:meta.slot||null, analysis:meta.analysis||null, val:0, data:null};
  if(kind==='dist'||kind==='anapair') u.lines.push(mkLine(markers[0].position,markers[1].position,0xff3b30));
  else if(kind==='ang'){ u.lines.push(mkLine(markers[1].position,markers[0].position,0x1e8ad0)); u.lines.push(mkLine(markers[1].position,markers[2].position,0x1e8ad0)); }
  else { for(let i=1;i<markers.length;i++) u.lines.push(mkLine(markers[i-1].position,markers[i].position,0xff9500)); }
  markers.forEach((m,i)=>{ m.userData={unit:u,i:i}; });
  if(kind==='dist'||kind==='ang'){ u.data={type:(kind==='ang'?'ang':'dist'),val:0,pts:u.pts,unit:u}; freeData.push(u.data); }
  units.push(u); recompute(u); return u; }
function removeUnit(u){ u.markers.forEach(rm); u.lines.forEach(rm); }
function recompute(u){ u.pts.length=0; u.markers.forEach(m=>u.pts.push([m.position.x,m.position.y,m.position.z]));
  const Pp=u.markers.map(m=>m.position);
  if(u.kind==='dist'||u.kind==='anapair'){ u.lines[0].geometry.setFromPoints([Pp[0],Pp[1]]); u.val=Pp[0].distanceTo(Pp[1]); }
  else if(u.kind==='ang'){ u.lines[0].geometry.setFromPoints([Pp[1],Pp[0]]); u.lines[1].geometry.setFromPoints([Pp[1],Pp[2]]); u.val=Pp[0].clone().sub(Pp[1]).angleTo(Pp[2].clone().sub(Pp[1]))*180/Math.PI; }
  else { let sum=0; for(let i=1;i<Pp.length;i++){ u.lines[i-1].geometry.setFromPoints([Pp[i-1],Pp[i]]); sum+=Pp[i-1].distanceTo(Pp[i]); } u.val=sum; }
  if(u.data){ u.data.val=u.val; renderList(); }
  if(u.analysis){ if(!anaSaved[u.analysis]) anaSaved[u.analysis]={vals:{},pts:{}}; anaSaved[u.analysis].vals[u.slot]=u.val; anaSaved[u.analysis].pts[u.slot]=u.pts.map(a=>a.slice()); if(u.analysis===anaKey) renderAna(); } }
function surfacePick(e){ const r=renderer.domElement.getBoundingClientRect(); mouse.x=((e.clientX-r.left)/r.width)*2-1; mouse.y=-((e.clientY-r.top)/r.height)*2+1; raycaster.setFromCamera(mouse,camera); const hits=raycaster.intersectObjects(group.children.filter(m=>m.visible),true); return hits.length?hits[0].point:null; }
function grabHandle(e){ const r=renderer.domElement.getBoundingClientRect(); let best=null,bd=15;
  measGroup.children.forEach(o=>{ if(!(o.userData&&o.userData.unit)) return; const v=o.position.clone().project(camera); const sx=r.left+(v.x*0.5+0.5)*r.width, sy=r.top+(-v.y*0.5+0.5)*r.height; const d=Math.hypot(e.clientX-sx,e.clientY-sy); if(d<bd){ bd=d; best=o; } }); return best; }
renderer.domElement.addEventListener('pointerdown',e=>{ downXY=[e.clientX,e.clientY]; const g=grabHandle(e); if(g){ drag=g.userData; controls.enabled=false; } });
renderer.domElement.addEventListener('pointermove',e=>{ if(!drag) return; const hit=surfacePick(e); if(hit){ drag.unit.markers[drag.i].position.copy(hit); recompute(drag.unit); } });
renderer.domElement.addEventListener('pointerup',e=>{
  if(drag){ drag=null; controls.enabled=true; autosave(); return; }
  const d=downXY; downXY=null; if(!d||e.button!==0) return;
  if(Math.hypot(e.clientX-d[0],e.clientY-d[1])>5) return;
  const active=(mode!=='off')||(window.__anaOpen&&window.__armedRow); if(!active) return;
  const p=surfacePick(e); if(!p) return;
  if(window.__anaOpen&&window.__armedRow){ if(window.__chainMode) chainAdd(p); else anaPair(p); }
  else if(mode==='angle') freeAngle(p); else if(mode==='dist') freeDist(p); });
function freeDist(p){ _fp.push(mkMarker(p)); if(_fp.length===2){ makeUnit('dist',_fp.slice(),{}); _fp=[]; renderList(); autosave(); } }
function freeAngle(p){ _fp.push(mkMarker(p)); if(_fp.length===3){ makeUnit('ang',_fp.slice(),{}); _fp=[]; renderList(); autosave(); } }
function renderList(){ const el=document.getElementById('measlist'); if(!el) return; let h='',tot=0;
  freeData.forEach((m,i)=>{ if(m.type==='ang'){ h+='<div class=mrow><span>#'+(i+1)+'</span><b>'+m.val.toFixed(1)+'°</b></div>'; } else { tot+=m.val; h+='<div class=mrow><span>#'+(i+1)+'</span><b>'+m.val.toFixed(1)+' mm</b></div>'; } });
  if(freeData.filter(m=>m.type!=='ang').length>1) h+='<div class=mrow style="border-top:1px solid #e2e8f0;margin-top:4px;padding-top:4px"><span>Somme</span><b>'+tot.toFixed(1)+' mm</b></div>';
  if(!freeData.length) h='<div style="color:#8794a8;font-size:12px">Clique 2 points (ou 3 pour un angle). Glisse un point pour lajuster.</div>';
  el.innerHTML=h; }
window.undoLast=function(){ if(_fp.length){ _fp.forEach(rm); _fp=[]; renderList(); return; } const last=freeData.pop(); if(last&&last.unit){ removeUnit(last.unit); const ix=units.indexOf(last.unit); if(ix>=0) units.splice(ix,1); } renderList(); autosave(); };
window.clearMeasure=function(){ units.slice().forEach(removeUnit); units.length=0; while(measGroup.children.length){ rm(measGroup.children[0]); } freeData=[]; _fp=[]; _ap=[]; chainPts=[]; chainTmp=[]; chainSum=0; for(const k in anaSaved) delete anaSaved[k]; renderList(); if(anaKey) renderAna(); autosave(); };
function _syncBtns(){ const mb=document.getElementById('measbtn'), ab=document.getElementById('anglebtn'); if(mb) mb.classList.toggle('on',mode==='dist'); if(ab) ab.classList.toggle('on',mode==='angle'); document.getElementById('meas').style.display=(mode!=='off')?'block':'none'; document.getElementById('hint').textContent=mode==='dist'?'Mesure : clique 2 points (mm). Glisse un point pour lajuster.':(mode==='angle'?'Angle : clique 3 points (le 2e = sommet).':'Fais glisser pour tourner • molette pour zoomer • clic droit pour déplacer'); }
window.toggleMeasure=function(){ mode=(mode==='dist')?'off':'dist'; _fp.forEach(rm); _fp=[]; _syncBtns(); };
window.toggleAngle=function(){ mode=(mode==='angle')?'off':'angle'; _fp.forEach(rm); _fp=[]; _syncBtns(); };
function _sumv(v,keys){ let s=0; for(let i=0;i<keys.length;i++){ if(v[keys[i]]==null) return null; s+=v[keys[i]]; } return s; }
function boltonDef(name,maxT,mandT,norm){ const rows=[]; maxT.forEach(t=>rows.push({k:'x'+t,g:'Maxillaire (mésio-distal)',t:t})); mandT.forEach(t=>rows.push({k:'n'+t,g:'Mandibule (mésio-distal)',t:t}));
  return {name:name,rows:rows,compute:function(v){ const sM=_sumv(v,maxT.map(t=>'x'+t)), sm=_sumv(v,mandT.map(t=>'n'+t)); if(sM==null||sm==null) return null;
    const ratio=sm/sM*100, f=norm/100; let r='Σ max = <b>'+sM.toFixed(1)+'</b> · Σ mand = <b>'+sm.toFixed(1)+'</b> mm<br>Rapport : <b>'+ratio.toFixed(1)+' %</b> <span class=nrm>(norme '+norm.toFixed(1).replace('.',',')+')</span><br>';
    if(ratio>norm) r+='Excès mandibulaire ≈ <b>'+(sm-sM*f).toFixed(1)+' mm</b>'; else r+='Excès maxillaire ≈ <b>'+(sM-sm/f).toFixed(1)+' mm</b>'; return r; }}; }
const ANA={
  bolton_ant: boltonDef('Bolton antérieur',['13','12','11','21','22','23'],['43','42','41','31','32','33'],77.2),
  bolton_tot: boltonDef('Bolton total',['16','15','14','13','12','11','21','22','23','24','25','26'],['46','45','44','43','42','41','31','32','33','34','35','36'],91.3),
  pont:{name:'Pont (transversal)',rows:[
    {k:'pmand',g:'Inter-molaire mandibulaire',t:'Cuspide à cuspide (46↔36)'},
    {k:'pmax',g:'Inter-molaire maxillaire',t:'Sillon à sillon (16↔26)'}],
    compute:function(v){ if(v.pmand==null&&v.pmax==null) return null; let r='';
      if(v.pmand!=null) r+='Mandibulaire (cuspide) = <b>'+v.pmand.toFixed(1)+' mm</b><br>';
      if(v.pmax!=null) r+='Maxillaire (sillon) = <b>'+v.pmax.toFixed(1)+' mm</b><br>';
      if(v.pmand!=null&&v.pmax!=null){ const dd=v.pmax-v.pmand; r+='Écart (maxi - mand) = <b>'+dd.toFixed(1)+' mm</b> — '+(Math.abs(dd)<1?'<b>coordonné</b>':(dd<0?('<b style="color:#c0392b">endognathie maxillaire ('+(-dd).toFixed(1)+' mm)</b>'):('<b style="color:#1e8ad0">excès transversal maxillaire ('+dd.toFixed(1)+' mm)</b>'))); }
      return r; }},
  ddm:{name:'Espace (DDM)',rows:[
    {k:'e1',g:'Dents (canine → canine)',t:'Dent 1'},{k:'e2',g:'Dents (canine → canine)',t:'Dent 2'},{k:'e3',g:'Dents (canine → canine)',t:'Dent 3'},
    {k:'e4',g:'Dents (canine → canine)',t:'Dent 4'},{k:'e5',g:'Dents (canine → canine)',t:'Dent 5'},{k:'e6',g:'Dents (canine → canine)',t:'Dent 6'},
    {k:'arch',g:'Arcade disponible',t:'Périmètre (chaîne)',chain:true}],
    compute:function(v){ const req=_sumv(v,['e1','e2','e3','e4','e5','e6']); let out='';
      out+= (req!=null)?('Requis (Σ 6 dents) = <b>'+req.toFixed(1)+' mm</b><br>'):'<span class=nrm>Mesure les 6 dents…</span><br>';
      if(v.arch!=null) out+='Disponible (arcade) = <b>'+v.arch.toFixed(1)+' mm</b><br>';
      if(req!=null&&v.arch!=null){ const ddm=v.arch-req; out+='DDM = <b>'+ddm.toFixed(1)+' mm</b> — '+(ddm<-0.3?('<b style="color:#c0392b">encombrement '+(-ddm).toFixed(1)+' mm</b>'):(ddm>0.3?('<b style="color:#1e8ad0">espace '+ddm.toFixed(1)+' mm</b>'):'<b>équilibré</b>')); }
      return out; }}
};
function gv(slot){ const c=anaSaved[anaKey]; return (c&&c.vals[slot]!=null)?c.vals[slot]:null; }
function vmap(){ const c=anaSaved[anaKey]; return c?c.vals:{}; }
function anaPair(p){ _ap.push(mkMarker(p)); if(_ap.length===2){ const slot=window.__armedRow; makeUnit('anapair',_ap.slice(),{slot:slot,analysis:anaKey}); _ap=[]; advanceArm(); renderAna(); autosave(); } }
function chainAdd(p){ const mk=mkMarker(p); if(chainPts.length){ chainTmp.push(mkLine(chainPts[chainPts.length-1].position,mk.position,0xff9500)); } chainPts.push(mk); chainSum=0; for(let i=1;i<chainPts.length;i++) chainSum+=chainPts[i-1].position.distanceTo(chainPts[i].position); renderAna(); }
window.chainDone=function(){ chainTmp.forEach(rm); chainTmp=[]; if(chainPts.length>=2&&window.__armedRow){ makeUnit('anachain',chainPts.slice(),{slot:window.__armedRow,analysis:anaKey}); } chainPts=[]; chainSum=0; window.__chainMode=false; advanceArm(); renderAna(); autosave(); };
function _syncChain(){ const ar=(anaKey&&window.__armedRow)?ANA[anaKey].rows.find(r=>r.k===window.__armedRow):null; if(ar&&ar.chain){ window.__chainMode=true; chainTmp.forEach(rm); chainTmp=[]; chainPts=[]; chainSum=0; } else window.__chainMode=false; }
function advanceArm(){ const rows=ANA[anaKey].rows, idx=rows.findIndex(r=>r.k===window.__armedRow); let nxt=null; for(let i=idx+1;i<rows.length;i++){ if(gv(rows[i].k)==null){ nxt=rows[i].k; break; } } if(!nxt) for(let j=0;j<rows.length;j++){ if(gv(rows[j].k)==null){ nxt=rows[j].k; break; } } window.__armedRow=nxt; _syncChain(); }
function renderAna(){ const host=document.getElementById('anabody'); if(!host) return;
  if(!anaKey){ host.innerHTML='<div class=nrm style="font-size:12px">Choisis une analyse, puis mesure chaque dent (clic 2 points). Pour DDM, mesure l arcade en chaîne de points.</div>'; return; }
  const def=ANA[anaKey]; let h='', lastg=null;
  def.rows.forEach(function(r){ if(r.g!==lastg){ h+='<div class=anag>'+r.g+'</div>'; lastg=r.g; } const val=gv(r.k), armed=(window.__armedRow===r.k); const disp=val!=null?val.toFixed(1)+' mm':(armed?(r.chain?'clique le long de l arcade…':'clique 2 points…'):'—'); h+='<div class="arow'+(armed?' armed':'')+'" onclick="armRow(&quot;'+r.k+'&quot;)"><span>'+r.t+'</span><b>'+disp+'</b></div>'; });
  const ar=(window.__armedRow)?def.rows.find(r=>r.k===window.__armedRow):null;
  if(ar&&ar.chain&&chainPts.length>0){ h+='<button class=bs style="width:100%;margin-top:6px" onclick="chainDone()">Terminer la chaîne ('+chainSum.toFixed(1)+' mm)</button>'; }
  const res=def.compute(vmap()); h+='<div class=ares>'+(res||'<span class=nrm>Complète les mesures.</span>')+'</div>'; host.innerHTML=h; }
window.armRow=function(k){ window.__armedRow=k; _syncChain(); renderAna(); };
window.selAna=function(k){ anaKey=k; if(!anaSaved[k]) anaSaved[k]={vals:{},pts:{}}; window.__armedRow=ANA[k].rows[0].k; _syncChain(); document.querySelectorAll('#anatabs .bs').forEach(function(b){ b.classList.toggle('on', b.dataset.a===k); }); renderAna(); };
window.clearAna=function(){ if(anaKey){ units.slice().filter(u=>u.analysis===anaKey).forEach(u=>{ removeUnit(u); const ix=units.indexOf(u); if(ix>=0) units.splice(ix,1); }); anaSaved[anaKey]={vals:{},pts:{}}; window.__armedRow=ANA[anaKey].rows[0].k; _syncChain(); } renderAna(); autosave(); };
window.toggleAna=function(){ window.__anaOpen=!window.__anaOpen; document.getElementById('ana').style.display=window.__anaOpen?'block':'none'; document.getElementById('anabtn').classList.toggle('on', window.__anaOpen); document.getElementById('hint').textContent=window.__anaOpen?'Analyses : choisis une analyse, puis mesure chaque dent (clic 2 points). Glisse un point pour lajuster.':'Fais glisser pour tourner • molette pour zoomer • clic droit pour déplacer'; };
function serialize(){ return { free: freeData.map(m=>({type:m.type,val:m.val,pts:m.pts})), ana: Object.keys(anaSaved).map(k=>({analysis:k, vals:anaSaved[k].vals, pts:anaSaved[k].pts})) }; }
function autosave(){ try{ fetch(SAVEURL,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(serialize())}); }catch(e){} }
window.saveMeasures=function(){ const b=document.getElementById('savebtn'); try{ fetch(SAVEURL,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(serialize())}).then(function(){ if(b){ b.textContent='✓ Enregistré'; setTimeout(function(){ b.textContent='💾 Enregistrer'; },1600);} }); }catch(e){ if(b) b.textContent='Échec'; } };
function restore(d){ if(!d) return;
  (d.free||[]).forEach(function(m){ const mks=(m.pts||[]).map(a=>mkMarker(_V(a))); if(m.type==='dist'&&mks.length>=2) makeUnit('dist',mks,{}); else if(m.type==='ang'&&mks.length>=3) makeUnit('ang',mks,{}); });
  (d.ana||[]).forEach(function(a){ if(!anaSaved[a.analysis]) anaSaved[a.analysis]={vals:{},pts:{}}; const rows=(ANA[a.analysis]||{}).rows||[]; for(const slot in (a.pts||{})){ const mks=(a.pts[slot]||[]).map(x=>mkMarker(_V(x))); const rd=rows.find(r=>r.k===slot); if(rd&&rd.chain){ if(mks.length>=2) makeUnit('anachain',mks,{slot:slot,analysis:a.analysis}); } else if(mks.length>=2) makeUnit('anapair',mks,{slot:slot,analysis:a.analysis}); } });
  renderList(); if(anaKey) renderAna(); }
function loadMeasures(){ try{ fetch(SAVEURL).then(r=>r.json()).then(restore).catch(function(){}); }catch(e){} }
loadMeasures();
(function loop(){ requestAnimationFrame(loop); controls.update(); renderer.render(scene,camera); })();
setTimeout(()=>{ if(loaded===0) document.getElementById('err').style.display='flex'; }, 12000);
</script></body></html>"""

@app.route("/lib/<path:p>")
def lib_file(p):
    root = os.path.join(HERE, "vendor", "three")
    fp = os.path.normpath(os.path.join(root, p))
    if (not fp.startswith(root)) or (not os.path.exists(fp)): abort(404)
    return send_file(fp)


ODONTO_STATES = [("", "Saine"), ("carie", "Carie"), ("soin", "Soin / couronne"),
                 ("absente", "Absente"), ("agenesie", "Agénésie"), ("incluse", "Incluse"),
                 ("extraire", "À extraire"), ("eruption", "En éruption")]
ODONTO_UP = [18,17,16,15,14,13,12,11,21,22,23,24,25,26,27,28]
ODONTO_LO = [48,47,46,45,44,43,42,41,31,32,33,34,35,36,37,38]


def odonto_card(slug, rid, r):
    od = r.get("odonto", {}) or {}
    def cell(n):
        s = od.get(str(n), "")
        return '<div class="tooth s-%s" data-n="%d" onclick="odClick(this)"><b>%d</b></div>' % (s or "ok", n, n)
    def row(nums):
        return ('<div class=odrow><div class=odq>%s</div><div class=odmid></div><div class=odq>%s</div></div>'
                % ("".join(cell(n) for n in nums[:8]), "".join(cell(n) for n in nums[8:])))
    tools = "".join('<button type=button class="odb s-%s%s" data-k="%s" onclick="odBrush(this)">%s</button>'
                    % (k or "ok", " on" if k == "carie" else "", k, lbl) for k, lbl in ODONTO_STATES)
    import json as _j
    style = ('<style>'
             '.odwrap{--od-ok:#ffffff;--od-carie:#e05252;--od-soin:#3b82f6;--od-absente:#9aa5b1;'
             '--od-agenesie:#a855f7;--od-incluse:#f59e0b;--od-extraire:#b91c1c;--od-eruption:#22c55e}'
             '.odtools{display:flex;flex-wrap:wrap;gap:6px;margin:2px 0 14px}'
             '.odb{display:flex;align-items:center;gap:7px;font-size:12.5px;padding:5px 10px;border:1px solid var(--line);border-radius:8px;background:var(--card);cursor:pointer;color:var(--ink)}'
             '.odb::before{content:"";width:13px;height:13px;border-radius:4px;border:1px solid rgba(0,0,0,.15)}'
             '.odb.on{border-color:var(--acc);box-shadow:0 0 0 2px rgba(30,138,208,.18)}'
             '.odb.s-ok::before{background:#fff}.odb.s-carie::before{background:var(--od-carie)}.odb.s-soin::before{background:var(--od-soin)}'
             '.odb.s-absente::before{background:var(--od-absente)}.odb.s-agenesie::before{background:var(--od-agenesie)}'
             '.odb.s-incluse::before{background:var(--od-incluse)}.odb.s-extraire::before{background:var(--od-extraire)}.odb.s-eruption::before{background:var(--od-eruption)}'
             '.odrow{display:flex;align-items:stretch;gap:3px;justify-content:center;margin:5px 0}'
             '.odq{display:flex;gap:3px}.odmid{width:2px;background:var(--line);margin:0 6px}'
             '.tooth{width:38px;height:46px;border:1px solid #cfd8e0;border-radius:7px;display:flex;align-items:center;justify-content:center;'
             'cursor:pointer;background:#fff;color:#334;font-size:12px;user-select:none;transition:.1s;position:relative}'
             '.tooth:hover{transform:translateY(-2px);box-shadow:0 3px 8px rgba(0,0,0,.12)}'
             '.tooth b{font-weight:600;pointer-events:none}'
             '.tooth.s-carie{background:var(--od-carie);color:#fff;border-color:var(--od-carie)}'
             '.tooth.s-soin{background:var(--od-soin);color:#fff;border-color:var(--od-soin)}'
             '.tooth.s-absente{background:var(--od-absente);color:#fff;border-color:var(--od-absente);opacity:.6}'
             '.tooth.s-agenesie{background:var(--od-agenesie);color:#fff;border-color:var(--od-agenesie)}'
             '.tooth.s-incluse{background:var(--od-incluse);color:#fff;border-color:var(--od-incluse)}'
             '.tooth.s-extraire{background:var(--od-extraire);color:#fff;border-color:var(--od-extraire)}'
             '.tooth.s-extraire::after{content:"✕";position:absolute;font-size:20px;color:rgba(255,255,255,.85)}'
             '.tooth.s-eruption{background:var(--od-eruption);color:#fff;border-color:var(--od-eruption)}'
             '.odsave{font-size:12px;color:var(--mut);margin-left:8px}'
             '</style>')
    js = ('<script>'
          '(function(){var ODATA=%s;var BRUSH="carie";var SAVEURL=%s;var LBL=%s;'
          'window.odBrush=function(b){BRUSH=b.dataset.k;document.querySelectorAll(".odb").forEach(function(x){x.classList.remove("on");});b.classList.add("on");};'
          'window.odClick=function(el){var n=el.dataset.n;if(BRUSH===""){delete ODATA[n];}else{ODATA[n]=BRUSH;}'
          'el.className="tooth s-"+(ODATA[n]||"ok");el.title=LBL[ODATA[n]||""]||"Saine";odSave();};'
          'function odSave(){var s=document.getElementById("odsave");if(s)s.textContent="enregistrement…";'
          'try{fetch(SAVEURL,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(ODATA)}).then(function(){if(s)s.textContent="✓ enregistré";setTimeout(function(){if(s)s.textContent="";},1200);});}catch(e){}}'
          '})();</script>') % (_j.dumps(od), _j.dumps(url_for("record_odonto", slug=slug, rid=rid)),
                               _j.dumps({k: l for k, l in ODONTO_STATES if k}))
    hint = '<p class=muted style="font-size:12.5px;margin:0 0 6px">Choisis un état ci-dessus, puis clique les dents concernées. « Saine » efface. <span id=odsave class=odsave></span></p>'
    return ('<div class=odwrap>' + style + hint + '<div class=odtools>' + tools + '</div>'
            + row(ODONTO_UP) + row(ODONTO_LO) + js + '</div>')


@app.route("/record/<slug>/<rid>/odonto", methods=["POST"])
def record_odonto(slug, rid):
    if not logged(): abort(403)
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r: abort(404)
    try:
        data = request.get_json(force=True, silent=True)
        if not isinstance(data, dict): data = {}
        clean = {}
        valid = {k for k, _l in ODONTO_STATES if k}
        for k, v in data.items():
            if str(k).isdigit() and v in valid: clean[str(k)] = v
        r["odonto"] = clean; save_rec(slug, rid, r)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "err": str(e)}), 500


@app.route("/record/<slug>/<rid>/measures", methods=["GET", "POST"])
def record_measures(slug, rid):
    if not logged(): abort(403)
    base = rdir(slug, rid); fp = os.path.join(base, "measures3d.json")
    if request.method == "POST":
        try:
            data = request.get_json(force=True, silent=True) or {}
            json.dump(data, open(fp, "w"), ensure_ascii=False)
            return jsonify({"ok": True})
        except Exception as e:
            return jsonify({"ok": False, "err": str(e)}), 500
    if os.path.exists(fp):
        try: return jsonify(json.load(open(fp)))
        except Exception: return jsonify(None)
    return jsonify(None)


@app.route("/record/<slug>/<rid>/viewer3d")
def viewer3d(slug, rid):
    if not logged(): return redirect(url_for("login"))
    base = rdir(slug, rid)
    up = "04_stl_bruts/UpperJawScan.stl"; lo = "04_stl_bruts/LowerJawScan.stl"
    upurl = url_for("rfile", slug=slug, rid=rid, path=up) if os.path.exists(os.path.join(base, up)) else ""
    lourl = url_for("rfile", slug=slug, rid=rid, path=lo) if os.path.exists(os.path.join(base, lo)) else ""
    html = (VIEWER_HTML.replace("__UPURL__", upurl).replace("__LOURL__", lourl)
            .replace("__BACK__", url_for("record", slug=slug, rid=rid)).replace("__SAVEURL__", url_for("record_measures", slug=slug, rid=rid)))
    return html

# ---- éditeur de cadrage photo (WYSIWYG : l'image orientée est rendue côté serveur) ----
@app.route("/record/<slug>/<rid>/oriented/<slot>")
def oriented(slug, rid, slot):
    """Renvoie la photo d'origine avec l'orientation demandée (rotation/miroir/retournement),
    exactement comme le pipeline l'appliquera. Sert de fond WYSIWYG à l'éditeur."""
    if not logged(): return redirect(url_for("login"))
    from PIL import Image, ImageOps
    import io as _io
    base = rdir(slug, rid)
    p = os.path.join(base, "01_photos_brutes", slot + ".jpg")
    if not os.path.exists(p): abort(404)
    im = ImageOps.exif_transpose(Image.open(p)).convert("RGB")
    try: rot = int(request.args.get("rot", 0)) % 360
    except Exception: rot = 0
    try: fine = max(-45.0, min(45.0, float(request.args.get("fine", 0))))
    except Exception: fine = 0.0
    ang = rot + fine
    if ang: im = im.rotate(ang, expand=True, resample=Image.BICUBIC)
    if request.args.get("mirror") == "1": im = ImageOps.mirror(im)
    if request.args.get("flip") == "1": im = ImageOps.flip(im)
    im.thumbnail((1200, 1200), Image.LANCZOS)
    buf = _io.BytesIO(); im.save(buf, "JPEG", quality=88); buf.seek(0)
    return send_file(buf, mimetype="image/jpeg")

CROP_HTML = """<!doctype html><html lang=fr><head><meta charset=utf-8>
<meta name=viewport content="width=device-width,initial-scale=1"><title>Recadrage — __LABEL__</title>
<style>
:root{--acc:#2F5CA8}
*{box-sizing:border-box}
html,body{margin:0;height:100%;font:14px -apple-system,BlinkMacSystemFont,sans-serif;background:linear-gradient(#eef3fb,#cdd9ee);overflow:hidden}
#bar{position:fixed;top:0;left:0;right:0;height:54px;z-index:10;display:flex;gap:6px;align-items:center;padding:0 10px;background:rgba(255,255,255,.94);border-bottom:1px solid #d7deea;flex-wrap:wrap}
#bar .lbl{font-weight:700;color:var(--acc);margin-right:6px}
.b{background:#fff;border:1px solid #c7d0df;border-radius:8px;padding:7px 11px;cursor:pointer;font-size:13px;color:#24405f}
.b:hover{background:#f2f6fc}
.b.pri{background:var(--acc);border-color:var(--acc);color:#fff;font-weight:600}
.b.pri:hover{background:#264c8a}
.b.on{background:var(--acc);color:#fff;border-color:var(--acc)}
#stage{position:fixed;top:54px;left:0;right:0;bottom:40px;display:flex;align-items:center;justify-content:center;padding:14px}
#wrap{position:relative;display:inline-block;line-height:0;box-shadow:0 6px 24px rgba(0,0,0,.25)}
#img{max-width:100%;max-height:100%;display:block;user-select:none;-webkit-user-drag:none}
#crop{position:absolute;border:2px solid #fff;box-shadow:0 0 0 9999px rgba(0,0,0,.55);cursor:move;touch-action:none}
.h{position:absolute;width:16px;height:16px;background:#fff;border:2px solid var(--acc);border-radius:50%;touch-action:none}
.h.nw{left:-9px;top:-9px;cursor:nwse-resize}.h.ne{right:-9px;top:-9px;cursor:nesw-resize}
.h.sw{left:-9px;bottom:-9px;cursor:nesw-resize}.h.se{right:-9px;bottom:-9px;cursor:nwse-resize}
#hint{position:fixed;bottom:10px;left:0;right:0;text-align:center;color:#3a4a63}
.grid3{position:absolute;inset:0;pointer-events:none}
.grid3:before,.grid3:after{content:"";position:absolute;background:rgba(255,255,255,.35)}
.grid3:before{left:33.33%;right:33.33%;top:0;bottom:0;border-left:1px solid rgba(255,255,255,.35);border-right:1px solid rgba(255,255,255,.35)}
.grid3:after{top:33.33%;bottom:33.33%;left:0;right:0;border-top:1px solid rgba(255,255,255,.35);border-bottom:1px solid rgba(255,255,255,.35)}
.fgrp{display:flex;align-items:center;gap:7px;padding:0 8px;margin:0 2px;border-left:1px solid #e2e7f0;border-right:1px solid #e2e7f0}
.fgrp .cap{font-size:12px;color:#5a6b86;font-weight:600}
.knob{position:relative;width:44px;height:44px;border-radius:50%;background:radial-gradient(circle at 50% 36%,#ffffff,#e6edf7);border:1px solid #b6c4dd;box-shadow:inset 0 1px 2px rgba(255,255,255,.9),0 1px 3px rgba(0,0,0,.18);cursor:grab;touch-action:none;flex:none}
.knob:active{cursor:grabbing}
.knob .ind{position:absolute;left:50%;top:4px;width:3px;height:15px;background:var(--acc);border-radius:2px;transform-origin:50% 18px;margin-left:-1.5px}
.knob .dot{position:absolute;inset:0;margin:auto;width:6px;height:6px;background:#93a7c8;border-radius:50%}
.fineval{font-variant-numeric:tabular-nums;min-width:50px;text-align:center;font-weight:700;color:#24405f;font-size:13px}
.bmini{background:#fff;border:1px solid #c7d0df;border-radius:7px;width:26px;height:28px;cursor:pointer;font-size:16px;line-height:1;color:#24405f}
.bmini:hover{background:#f2f6fc}
</style></head><body>
<div id=bar>
  <span class=lbl>__LABEL__</span>
  <button class=b onclick="rotate(-90)" title="Rotation gauche">&#8634; Gauche</button>
  <button class=b onclick="rotate(90)" title="Rotation droite">&#8635; Droite</button>
  <span class=fgrp>
    <span class=cap>Redresser</span>
    <div class=knob id=knob title="Tourne la molette pour incliner &middot; molette souris pour ajuster &middot; double-clic = 0"><div class=ind id=knobind></div><div class=dot></div></div>
    <button class=bmini onclick="nudge(-0.5)" title="&minus;0,5&deg;">&minus;</button>
    <span class=fineval id=fineval>0,0&deg;</span>
    <button class=bmini onclick="nudge(0.5)" title="+0,5&deg;">+</button>
  </span>
  <button class=b id=bmir onclick="tog('mirror')" title="Miroir horizontal">&#8646; Miroir</button>
  <button class=b id=bflip onclick="tog('flip')" title="Retourner vertical">&#8645; Retourner</button>
  <button class=b onclick="resetAuto()" title="Revenir au cadrage automatique">Cadrage auto</button>
  <span style=flex:1></span>
  <a class=b href="__BACK__">Annuler</a>
  <button class="b pri" onclick="save()">Enregistrer</button>
</div>
<div id=stage><div id=wrap><img id=img draggable=false>
  <div id=crop><div class=grid3></div>
    <div class="h nw" data-c=nw></div><div class="h ne" data-c=ne></div>
    <div class="h sw" data-c=sw></div><div class="h se" data-c=se></div>
  </div>
</div></div>
<div id=hint>Glisse le cadre pour le déplacer &middot; tire un coin ou utilise la <b>molette</b> pour zoomer &middot; les proportions du bilan sont conservées</div>
<div id=prevwrap style="position:fixed;bottom:12px;right:14px;background:rgba(255,255,255,.96);border:1px solid #cdd6e6;border-radius:10px;padding:8px;box-shadow:0 4px 16px rgba(0,0,0,.18);z-index:20">
  <div style="font-size:11px;color:#5a6b86;font-weight:600;margin-bottom:5px;text-align:center">Aperçu du résultat</div>
  <canvas id=prev width=170 height=170 style="display:block;border-radius:6px;background:#f2f4f8"></canvas>
</div>
<script>
const IMGBASE="__IMGBASE__", SAVEURL="__SAVEURL__", BACK="__BACK__";
const ASPECT=__ASPECT__;
const INIT=__INIT__, DEF=__DEF__;
let st={rot:INIT.rot|0, fine:+(INIT.fine||0), mirror:!!INIT.mirror, flip:!!INIT.flip};
let box={x:0,y:0,w:0,h:0};          // px affichés, repère image
let dw=1,dh=1;                       // dims affichées de l'image
let pendingFrac=null;                // fractions à appliquer au prochain chargement
let reloadTimer=null;
const img=document.getElementById('img'), crop=document.getElementById('crop');
const wrap=document.getElementById('wrap');
function url(){return IMGBASE+"?rot="+((st.rot%360+360)%360)+"&fine="+st.fine.toFixed(1)+"&mirror="+(st.mirror?1:0)+"&flip="+(st.flip?1:0)+"&_="+Date.now();}
function loadImg(frac){pendingFrac=frac; img.src=url();}
// --- rotation fine (redressement) ---
function curFrac(){return {x:box.x/dw, y:box.y/dh, w:box.w/dw, h:box.h/dh};}
function fmtFine(v){return (v>0?'+':(v<0?'−':''))+Math.abs(v).toFixed(1).replace('.',',')+'°';}
function renderFine(){var el=document.getElementById('fineval');if(el)el.textContent=fmtFine(st.fine);
  var k=document.getElementById('knobind');if(k)k.style.transform='rotate('+st.fine+'deg)';}
function scheduleReload(){if(reloadTimer)clearTimeout(reloadTimer);reloadTimer=setTimeout(function(){loadImg(curFrac());},90);}
function setFine(v){v=Math.max(-45,Math.min(45,v));st.fine=Math.round(v*10)/10;renderFine();scheduleReload();}
function nudge(d){setFine(st.fine+d);}
img.onload=function(){
  const r=img.getBoundingClientRect(); dw=r.width; dh=r.height;
  wrap.style.width=dw+"px"; wrap.style.height=dh+"px";
  var cv=document.getElementById('prev'); if(cv){ cv.height=Math.round(cv.width/ASPECT); }
  let f=pendingFrac; pendingFrac=null;
  if(f){ box={x:f.x*dw, y:f.y*dh, w:f.w*dw, h:f.h*dh}; }
  else { seedCenter(); }
  clampBox(); draw();
};
wrap.addEventListener('wheel',function(e){ e.preventDefault(); zoomBox(e.deltaY<0?0.94:1.06); },{passive:false});
document.getElementById('bmir').classList.toggle('on', st.mirror);
document.getElementById('bflip').classList.toggle('on', st.flip);
function seedCenter(){ // plus grand cadre au ratio ASPECT, centré, 92%
  let w=dw*0.92, h=w/ASPECT; if(h>dh*0.92){ h=dh*0.92; w=h*ASPECT; }
  box={x:(dw-w)/2, y:(dh-h)/2, w:w, h:h};
}
function clampBox(){
  box.w=Math.max(24,Math.min(box.w,dw)); box.h=box.w/ASPECT;
  if(box.h>dh){ box.h=dh; box.w=box.h*ASPECT; }
  box.x=Math.max(0,Math.min(box.x,dw-box.w));
  box.y=Math.max(0,Math.min(box.y,dh-box.h));
}
function draw(){ crop.style.left=box.x+"px"; crop.style.top=box.y+"px"; crop.style.width=box.w+"px"; crop.style.height=box.h+"px"; drawPrev(); }
function drawPrev(){
  var cv=document.getElementById('prev'); if(!cv||!img.naturalWidth||!dw) return;
  var ctx=cv.getContext('2d'); ctx.clearRect(0,0,cv.width,cv.height);
  var sx=box.x/dw*img.naturalWidth, sy=box.y/dh*img.naturalHeight,
      sw=box.w/dw*img.naturalWidth, sh=box.h/dh*img.naturalHeight;
  try{ ctx.drawImage(img, sx,sy,sw,sh, 0,0,cv.width,cv.height); }catch(e){}
}
function zoomBox(factor){
  var cx=box.x+box.w/2, cy=box.y+box.h/2;
  box.w=Math.max(24,Math.min(box.w*factor,dw)); box.h=box.w/ASPECT;
  if(box.h>dh){ box.h=dh; box.w=box.h*ASPECT; }
  box.x=cx-box.w/2; box.y=cy-box.h/2; clampBox(); draw();
}
function rotate(d){
  // conserve la région visée : bascule la boîte dans le nouveau repère (dims inversées)
  st.rot=((st.rot+d)%360+360)%360;
  loadImg(null); // re-seed centré (dims inversées) — l'utilisateur ajuste ensuite
}
function tog(k){
  st[k]=!st[k]; document.getElementById(k=='mirror'?'bmir':'bflip').classList.toggle('on', st[k]);
  // miroir/retournement : mêmes dims -> on reflète la boîte pour garder la même zone
  let f={x:box.x/dw, y:box.y/dh, w:box.w/dw, h:box.h/dh};
  if(k=='mirror') f.x=1-f.x-f.w; else f.y=1-f.y-f.h;
  loadImg(f);
}
function resetAuto(){
  st={rot:DEF.rot|0, fine:+(DEF.fine||0), mirror:!!DEF.mirror, flip:!!DEF.flip};
  document.getElementById('bmir').classList.toggle('on', st.mirror);
  document.getElementById('bflip').classList.toggle('on', st.flip);
  renderFine();
  loadImg({x:DEF.x,y:DEF.y,w:DEF.w,h:DEF.h});
}
// --- déplacement / redimensionnement ---
let drag=null;
function pt(e){const t=e.touches?e.touches[0]:e; return {x:t.clientX,y:t.clientY};}
crop.addEventListener('pointerdown',e=>{
  if(e.target.classList.contains('h')) return; // géré par les poignées
  const p=pt(e); drag={mode:'move',sx:p.x,sy:p.y,bx:box.x,by:box.y}; crop.setPointerCapture(e.pointerId); e.preventDefault();
});
document.querySelectorAll('.h').forEach(h=>{
  h.addEventListener('pointerdown',e=>{
    const p=pt(e); drag={mode:'resize',corner:h.dataset.c}; h.setPointerCapture(e.pointerId); e.stopPropagation(); e.preventDefault();
  });
});
window.addEventListener('pointermove',e=>{
  if(!drag) return; const p=pt(e);
  if(drag.mode=='move'){ box.x=drag.bx+(p.x-drag.sx); box.y=drag.by+(p.y-drag.sy); box.x=Math.max(0,Math.min(box.x,dw-box.w)); box.y=Math.max(0,Math.min(box.y,dh-box.h)); draw(); return; }
  // resize : coin opposé = ancre fixe
  const r=wrap.getBoundingClientRect(); let mx=p.x-r.left, my=p.y-r.top;
  mx=Math.max(0,Math.min(mx,dw)); my=Math.max(0,Math.min(my,dh));
  const c=drag.corner; let ax,ay,maxW,w;
  if(c=='se'){ ax=box.x; ay=box.y; maxW=Math.min(dw-ax,(dh-ay)*ASPECT); w=Math.min(Math.max(mx-ax,24),maxW); box.x=ax; box.y=ay; }
  else if(c=='nw'){ ax=box.x+box.w; ay=box.y+box.h; maxW=Math.min(ax,ay*ASPECT); w=Math.min(Math.max(ax-mx,24),maxW); box.x=ax-w; box.y=ay-w/ASPECT; }
  else if(c=='ne'){ ax=box.x; ay=box.y+box.h; maxW=Math.min(dw-ax,ay*ASPECT); w=Math.min(Math.max(mx-ax,24),maxW); box.x=ax; box.y=ay-w/ASPECT; }
  else { ax=box.x+box.w; ay=box.y; maxW=Math.min(ax,(dh-ay)*ASPECT); w=Math.min(Math.max(ax-mx,24),maxW); box.x=ax-w; box.y=ay; }
  box.w=w; box.h=w/ASPECT; draw();
});
window.addEventListener('pointerup',()=>{drag=null;});
// --- molette de redressement (rotation fine) ---
const knob=document.getElementById('knob');
let kdrag=null;
function knobAngle(e){const r=knob.getBoundingClientRect();const cx=r.left+r.width/2,cy=r.top+r.height/2;const p=pt(e);return Math.atan2(p.y-cy,p.x-cx)*180/Math.PI;}
knob.addEventListener('pointerdown',e=>{kdrag={a0:knobAngle(e),f0:st.fine};knob.setPointerCapture(e.pointerId);e.preventDefault();e.stopPropagation();});
knob.addEventListener('pointermove',e=>{if(!kdrag)return;let d=knobAngle(e)-kdrag.a0;while(d>180)d-=360;while(d<-180)d+=360;setFine(kdrag.f0+d);e.preventDefault();});
knob.addEventListener('pointerup',()=>{kdrag=null;});
knob.addEventListener('dblclick',e=>{e.preventDefault();setFine(0);});
knob.addEventListener('wheel',e=>{e.preventDefault();setFine(st.fine+(e.deltaY<0?0.2:-0.2));},{passive:false});
function save(){
  const f={rot:((st.rot%360+360)%360), fine:st.fine, mirror:st.mirror, flip:st.flip,
           x:box.x/dw, y:box.y/dh, w:box.w/dw, h:box.h/dh};
  const btn=document.querySelector('.b.pri'); btn.textContent="Enregistrement…"; btn.disabled=true;
  fetch(SAVEURL,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(f)})
    .then(r=>r.json()).then(d=>{ window.location=d.redirect||BACK; })
    .catch(()=>{ btn.textContent="Erreur — réessayer"; btn.disabled=false; });
}
// premier chargement : orientation + boîte initiales
renderFine();
loadImg({x:INIT.x,y:INIT.y,w:INIT.w,h:INIT.h});
</script></body></html>"""

@app.route("/record/<slug>/<rid>/crop/<slot>", methods=["GET"])
def crop_editor(slug, rid, slot):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r: abort(404)
    base = rdir(slug, rid)
    if not os.path.exists(os.path.join(base, "01_photos_brutes", slot + ".jpg")): abort(404)
    df = P.default_crop(base, slot) or {"rot":0,"fine":0,"mirror":False,"flip":False,"x":0,"y":0,"w":1,"h":1,"aspect":1}
    df.setdefault("fine", 0)
    init = r.get("crops", {}).get(slot) or df
    aspect = P._CROP.get(slot, (1,))[0]
    label = dict(PHOTO_FIELDS).get(slot, slot)
    html = (CROP_HTML
        .replace("__LABEL__", label)
        .replace("__IMGBASE__", url_for("oriented", slug=slug, rid=rid, slot=slot))
        .replace("__SAVEURL__", url_for("crop_save", slug=slug, rid=rid, slot=slot))
        .replace("__BACK__", url_for("record", slug=slug, rid=rid))
        .replace("__ASPECT__", "%.6f" % aspect)
        .replace("__INIT__", json.dumps({k: init.get(k, df.get(k)) for k in ("rot","fine","mirror","flip","x","y","w","h")}))
        .replace("__DEF__", json.dumps({k: df.get(k) for k in ("rot","fine","mirror","flip","x","y","w","h")})))
    return html

@app.route("/record/<slug>/<rid>/crop/<slot>", methods=["POST"])
def crop_save(slug, rid, slot):
    if not logged(): return jsonify(ok=False), 403
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r: return jsonify(ok=False), 404
    d = request.get_json(force=True, silent=True) or {}
    def cl(v):
        try: return max(0.0, min(1.0, float(v)))
        except Exception: return 0.0
    def fcl(v):
        try: return max(-45.0, min(45.0, float(v)))
        except Exception: return 0.0
    c = {"rot": int(d.get("rot", 0)) % 360, "fine": round(fcl(d.get("fine", 0)), 2),
         "mirror": bool(d.get("mirror")), "flip": bool(d.get("flip")),
         "x": cl(d.get("x", 0)), "y": cl(d.get("y", 0)),
         "w": max(0.02, cl(d.get("w", 1))), "h": max(0.02, cl(d.get("h", 1))),
         "_ts": datetime.datetime.now().isoformat(timespec="seconds")}   # horodatage -> apprentissage récent
    r.setdefault("crops", {})[slot] = c
    _render_slot(rdir(slug, rid), slot, c)         # rend UNIQUEMENT cette photo -> instantané
    save_rec(slug, rid, r)
    return jsonify(ok=True, redirect=url_for("record", slug=slug, rid=rid))

def _rotate_file_inplace(path, op):
    try:
        from PIL import Image as _I, ImageOps as _IO
        im = _IO.exif_transpose(_I.open(path)).convert("RGB")
        if op == "rotl": im = im.rotate(-90, expand=True, resample=_I.BICUBIC)
        elif op == "rotr": im = im.rotate(90, expand=True, resample=_I.BICUBIC)
        elif op == "mirror": im = _IO.mirror(im)
        elif op == "flip": im = _IO.flip(im)
        else: return False
        im.save(path, "JPEG", quality=92); return True
    except Exception:
        return False


@app.route("/record/<slug>/<rid>/orient/<slot>/<op>", methods=["POST"])
def orient_slot(slug, rid, slot, op):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r: abort(404)
    if slot not in {k for k, _ in PHOTO_FIELDS}: abort(404)
    base = rdir(slug, rid)
    _raw = os.path.join(base, "01_photos_brutes", slot + ".jpg")
    _proc = os.path.join(base, "02_photos_traitees", slot + ".jpg")
    if not os.path.exists(_raw):
        if os.path.exists(_proc) and _rotate_file_inplace(_proc, op):
            r["photos_verified"] = True; save_rec(slug, rid, r)
        else:
            flash("Photo introuvable.")
        return redirect(url_for("record", slug=slug, rid=rid) + "#photos")
    c = {"rot": 0, "fine": 0, "mirror": False, "flip": False, "x": 0, "y": 0, "w": 1, "h": 1}
    c.update(r.get("crops", {}).get(slot) or {})
    c.setdefault("fine", 0)
    if op == "rotl": c["rot"] = (int(c.get("rot", 0)) - 90) % 360
    elif op == "rotr": c["rot"] = (int(c.get("rot", 0)) + 90) % 360
    elif op == "mirror": c["mirror"] = not bool(c.get("mirror"))
    elif op == "flip": c["flip"] = not bool(c.get("flip"))
    else: return redirect(url_for("record", slug=slug, rid=rid) + "#photos")
    c["_ts"] = datetime.datetime.now().isoformat(timespec="seconds")
    r.setdefault("crops", {})[slot] = c
    r["photos_verified"] = True
    _render_slot(base, slot, c)
    save_rec(slug, rid, r)
    return redirect(url_for("record", slug=slug, rid=rid) + "#photos")


# ---- corriger la vue attribuée à une photo (erreur d'attribution à l'import) ----
@app.route("/record/<slug>/<rid>/reassign_photo/<slot>", methods=["POST"])
def reassign_photo(slug, rid, slot):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r: abort(404)
    to = request.form.get("to", "").strip()
    valid = {k for k, _ in PHOTO_FIELDS}
    if to not in valid or slot not in valid or to == slot:
        return redirect(url_for("record", slug=slug, rid=rid))
    raw = os.path.join(rdir(slug, rid), "01_photos_brutes")
    src = os.path.join(raw, slot + ".jpg")
    if not os.path.exists(src):
        flash("Photo introuvable."); return redirect(url_for("record", slug=slug, rid=rid))
    dst = os.path.join(raw, to + ".jpg")
    r.setdefault("crops", {}); r.setdefault("rotations", {})
    tmp = os.path.join(raw, "_swap_tmp.jpg")
    try:
        if os.path.exists(dst):
            # la vue cible est déjà occupée -> ÉCHANGE des deux photos (et de leur cadrage)
            os.replace(src, tmp); os.replace(dst, src); os.replace(tmp, dst)
            cs, ct = r["crops"].get(slot), r["crops"].get(to)
            r["crops"].pop(slot, None); r["crops"].pop(to, None)
            if ct is not None: r["crops"][slot] = ct
            if cs is not None: r["crops"][to] = cs
            rs, rt = r["rotations"].get(slot), r["rotations"].get(to)
            r["rotations"].pop(slot, None); r["rotations"].pop(to, None)
            if rt is not None: r["rotations"][slot] = rt
            if rs is not None: r["rotations"][to] = rs
            msg = "Vues échangées : « %s » ↔ « %s »." % (dict(PHOTO_FIELDS).get(slot), dict(PHOTO_FIELDS).get(to))
        else:
            os.replace(src, dst)
            if slot in r["crops"]: r["crops"][to] = r["crops"].pop(slot)
            if slot in r["rotations"]: r["rotations"][to] = r["rotations"].pop(slot)
            msg = "Photo ré-attribuée à « %s »." % dict(PHOTO_FIELDS).get(to)
    except Exception as e:
        flash("Impossible de ré-attribuer : %s" % e); return redirect(url_for("record", slug=slug, rid=rid))
    finally:
        if os.path.exists(tmp):
            try: os.remove(tmp)
            except Exception: pass
    r["photos_verified"] = True
    save_rec(slug, rid, r)
    _regenerate(slug, rid, pt, r, skip_stl=True)
    flash(msg)
    return redirect(url_for("record", slug=slug, rid=rid))

@app.route("/record/<slug>/<rid>/clearview/<view>", methods=["POST"])
def clear_view(slug, rid, view):
    """Retire la photo placée dans une vue (bilan/réévaluation). Non destructif :
    la photo reste dans la galerie (08_galerie) et peut être réassignée."""
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r: abort(404)
    if view not in {k for k, _ in PHOTO_FIELDS}:
        flash("Vue inconnue."); return redirect(url_for("record", slug=slug, rid=rid))
    for rel in ("01_photos_brutes/%s.jpg" % view, "02_photos_traitees/%s.jpg" % view):
        fp = os.path.join(rdir(slug, rid), rel)
        if os.path.exists(fp):
            try: os.remove(fp)
            except Exception: pass
    r.setdefault("crops", {}).pop(view, None)
    r.setdefault("rotations", {}).pop(view, None)
    r["photos_verified"] = True
    save_rec(slug, rid, r)
    flash("Photo retirée de « %s » (elle reste dans la galerie ci-dessous)." % dict(PHOTO_FIELDS).get(view, view))
    return redirect(url_for("record", slug=slug, rid=rid) + "#photos")

@app.route("/record/<slug>/<rid>/assign", methods=["POST"])
def assign_photo(slug, rid):
    """Assigne une photo de la galerie (08_galerie) à une vue : copie -> 01_photos_brutes/<vue>.jpg,
    réinitialise le cadrage de cette vue et régénère (recadrage). Glisser-déposer ou visionneuse."""
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r: abort(404)
    view = request.form.get("view", "").strip()
    fname = os.path.basename(request.form.get("file", "").strip())
    valid = {k for k, _ in PHOTO_FIELDS}
    src = os.path.join(rdir(slug, rid), "08_galerie", fname)
    if view not in valid or not fname or not os.path.exists(src):
        flash("Assignation impossible (photo ou vue invalide)."); return redirect(url_for("record", slug=slug, rid=rid))
    if not _place_as_is(src, rdir(slug, rid), view):
        flash("Échec de l'assignation."); return redirect(url_for("record", slug=slug, rid=rid))
    r.setdefault("crops", {}); r.setdefault("rotations", {})
    r["crops"].pop(view, None); r["rotations"].pop(view, None)
    r["photos_verified"] = True
    save_rec(slug, rid, r)
    flash("Photo assignée à « %s » (recadrage possible avec l'outil ✂)." % dict(PHOTO_FIELDS).get(view, view))
    return redirect(url_for("record", slug=slug, rid=rid) + "#photos")

# ======================================================================
#  IMPORT D'UN BILAN WORD EXISTANT (.docx) -> lecture et report dans un dossier
# ======================================================================
WORD_STAGE = os.path.join(DATA, "_word_stage.docx")

def _word_summary_html(res):
    p = res.get("patient", {})
    def esc(s): return str(s or "").replace("<", "&lt;")
    ident = "%s &middot; %s &middot; %s" % (esc(p.get("nom", "—")),
             {"masculin": "Homme", "féminin": "Femme"}.get(p.get("sexe"), "—"),
             esc(p.get("age") or p.get("dob") or "—"))
    def chip(label, val):
        return '<div style="display:flex;justify-content:space-between;gap:12px;padding:2px 0"><span class=muted>%s</span><b>%s</b></div>' % (label, esc(val))
    stein = "".join(chip(dict((n, s) for n, s, _m, _sd in STEINER_FIELDS).get(k, k), ("%g" % v))
                     for k, v in res.get("steiner", {}).items())
    esp = ""
    for k, cell in res.get("espace", {}).items():
        lbl = dict((kk, ll) for kk, ll, _ in ESPACE_ROWS).get(k, k)
        v = ("+%g" % cell["p"]) if cell.get("p") else ("−%g" % cell["m"] if cell.get("m") else "")
        esp += chip(lbl, v)
    syn = ""
    RN = dict((rk, rn) for rn, rk in DIAG_ROWS); CN = dict((ck, cn) for cn, ck in DIAG_COLS)
    for key, txt in res.get("synthese", {}).items():
        ck, rk = key.split("_", 1)
        syn += chip("%s — %s" % (CN.get(ck, ck), RN.get(rk, rk)), txt[:70])
    chv = ""
    CHVL = dict((k, l) for k, l, _pos in CHEVRON_FLOW)
    for k, d in res.get("chevrons", {}).items():
        vals = " · ".join("%s %s" % (dict(CHEVRON_POS_META).get(pp, pp), d[pp]) for pp in CHEVRON_POS if pp in d)
        chv += chip(CHVL.get(k, k), vals)
    def box(title, inner, empty="—"):
        return ('<div class=card style="flex:1;min-width:240px"><h3 style="margin:0 0 6px;color:var(--acc);font-size:14px">%s</h3>%s</div>'
                % (title, inner or ('<span class=muted>%s</span>' % empty)))
    txtblocks = ""
    for t, v in (("Motif", res.get("motif")), ("Objectifs", res.get("objectifs")), ("Moyens", res.get("moyens"))):
        if v: txtblocks += '<div style="margin-top:6px"><b>%s :</b><div style="white-space:pre-line">%s</div></div>' % (t, esc(v))
    return ('<div style="display:flex;gap:14px;flex-wrap:wrap">'
            + box("Identité", '<b>%s</b>' % ident)
            + box("Analyse de Steiner", stein, "aucune valeur lue")
            + box("Analyse d'espace", esp, "aucune valeur lue")
            + box("Synthèse diagnostique", syn, "aucune case lue")
            + box("Chevrons", chv, "aucun chevron lu")
            + '</div>' + (('<div class=card>%s</div>' % txtblocks) if txtblocks else ''))

@app.route("/word_import", methods=["GET", "POST"])
def word_import():
    if not logged(): return redirect(url_for("login"))
    if request.method == "POST":
        f = request.files.get("docx")
        if not f or not f.filename.lower().endswith(".docx"):
            flash("Choisis un fichier Word (.docx).")
            return redirect(url_for("word_import"))
        os.makedirs(DATA, exist_ok=True)
        f.save(WORD_STAGE)
        try:
            res = WI.parse_word_bilan(WORD_STAGE)
        except Exception as e:
            return page('<div class="flash err">Lecture impossible : %s</div><p><a href="%s">← Retour</a></p>'
                        % (str(e).replace("<", "&lt;"), url_for("word_import")))
        body = ('<p class=muted><a href="%s">← Bibliothèque</a></p><h1>Import d\'un bilan Word</h1>'
                '<div class=flash style="background:var(--accw);border-color:var(--line)">Voici ce que j\'ai lu dans <b>%s</b>. '
                'Vérifie, puis crée le dossier — tout restera modifiable ensuite (et tu pourras ajouter les photos/radios/STL via l\'import intelligent).</div>'
                '%s'
                '<form method=post action="%s" style="margin-top:14px;display:flex;gap:10px;flex-wrap:wrap">'
                '<button class=btn onclick="this.innerHTML=\'<span class=spin></span> Création…\'">Créer le dossier patient</button>'
                '<a class="btn sec" href="%s">Annuler</a></form>') % (
                    url_for("dashboard"), (f.filename or "").replace("<", "&lt;"),
                    _word_summary_html(res), url_for("word_import_commit"), url_for("dashboard"))
        return page(body)
    # GET : formulaire de dépôt
    body = ('<p class=muted><a href="%s">← Bibliothèque</a></p><h1>Importer un bilan Word</h1>'
            '<div class=card><p>Dépose un bilan existant au format <b>Word (.docx)</b>. L\'app le lit et reporte '
            'automatiquement l\'identité, l\'analyse de Steiner, l\'analyse d\'espace, la synthèse diagnostique, '
            'le motif et les objectifs/moyens dans un nouveau dossier patient.</p>'
            '<form method=post enctype=multipart/form-data onsubmit="this.querySelector(\'button\').innerHTML=\'<span class=spin></span> Lecture…\'">'
            '<input type=file name=docx accept=".docx" style="padding:18px;border:2px dashed #b9c6de;background:#f7f9fd">'
            '<div style="margin-top:14px"><button class=btn>Lire le bilan</button></div></form></div>') % url_for("dashboard")
    return page(body)

@app.route("/word_import/commit", methods=["POST"])
def word_import_commit():
    if not logged(): return redirect(url_for("login"))
    if not os.path.exists(WORD_STAGE):
        flash("Aucun bilan en attente — recommence l'import."); return redirect(url_for("word_import"))
    res = WI.parse_word_bilan(WORD_STAGE)
    p = res.get("patient", {})
    nom = (p.get("nom") or "Patient importé").strip()
    slug = slugify(nom) + "_" + datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    dob = p.get("dob", "")
    if dob:
        dob_court, age = age_from_dob(dob)
    else:
        dob_court, age = "—", (p.get("age") or "—")
    pt = {"slug": slug, "nom": nom, "dob": dob, "dob_court": dob_court, "age": age,
          "sexe": p.get("sexe", "non précisé"), "auteur": session["user"],
          "date_creation": datetime.datetime.now().isoformat(timespec="seconds"), "statut": "bilan", "records": []}
    os.makedirs(pdir(slug), exist_ok=True); save_patient(slug, pt)
    rid = create_empty_record(slug, "Bilan initial")
    r = load_rec(slug, rid)
    apply_word_to_record(r, res)
    save_rec(slug, rid, r)
    try:
        _regenerate(slug, rid, pt, r, skip_stl=True)
    except Exception:
        pass
    try: os.remove(WORD_STAGE)
    except Exception: pass
    flash("Bilan Word importé — vérifie/complète, puis ajoute les photos, radios et STL via « Import intelligent ».")
    return redirect(url_for("record", slug=slug, rid=rid))

# ---- import des valeurs d'un bilan Word DANS un enregistrement existant ----
@app.route("/record/<slug>/<rid>/word_values", methods=["GET", "POST"])
def record_word_values(slug, rid):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r: abort(404)
    if request.method == "POST":
        f = request.files.get("docx")
        if not f or not f.filename.lower().endswith(".docx"):
            flash("Choisis un fichier Word (.docx)."); return redirect(url_for("record_word_values", slug=slug, rid=rid))
        tmp = os.path.join(rdir(slug, rid), "_word_values.docx")
        f.save(tmp)
        try:
            res = WI.parse_word_bilan(tmp)
            apply_word_to_record(r, res)
            if r.get("status") == "empty": r["status"] = "processing"
            save_rec(slug, rid, r)
            try: _regenerate(slug, rid, pt, r, skip_stl=True)
            except Exception: pass
            flash("Valeurs du bilan Word ajoutées à cet enregistrement.")
        except Exception as e:
            flash("Lecture impossible : %s" % str(e))
        try: os.remove(tmp)
        except Exception: pass
        return redirect(url_for("record", slug=slug, rid=rid))
    body = ('<p class=muted><a href="%s">← Retour au dossier</a></p>'
            '<h1>Ajouter les valeurs d\'un bilan Word</h1><h2>%s · %s</h2>'
            '<div class=card><p>Dépose un bilan <b>Word (.docx)</b> : ses valeurs (Steiner, analyse d\'espace, '
            'synthèse, chevrons, motif, objectifs/moyens) seront <b>ajoutées à cet enregistrement</b>. '
            'Les champs déjà remplis ne sont pas écrasés ; les valeurs Steiner/espace/chevrons sont mises à jour.</p>'
            '<form method=post enctype=multipart/form-data onsubmit="this.querySelector(\'button\').innerHTML=\'<span class=spin></span> Lecture…\'">'
            '<input type=file name=docx accept=".docx" style="padding:18px;border:2px dashed #b9c6de;background:#f7f9fd">'
            '<div style="margin-top:14px"><button class=btn>Lire et ajouter</button>'
            ' <a class="btn sec" href="%s">Annuler</a></div></form></div>') % (
            url_for("record", slug=slug, rid=rid), pt["nom"], r.get("label", ""), url_for("record", slug=slug, rid=rid))
    return page(body)

# ---- fichiers ----
def _gen_thumb(full, tp, size=340, q=82):
    try:
        if (not os.path.exists(tp)) or (os.path.getmtime(tp) < os.path.getmtime(full)):
            from PIL import Image as _I, ImageOps as _IO
            im = _IO.exif_transpose(_I.open(full)).convert("RGB")
            im.thumbnail((size, size), _I.LANCZOS)
            d = os.path.dirname(tp)
            if d: os.makedirs(d, exist_ok=True)
            im.save(tp, "JPEG", quality=q)
        return True
    except Exception:
        return False


def _prewarm_thumbs():
    import time as _t
    _t.sleep(4.0)
    try:
        if not os.path.isdir(PATIENTS): return
        for slug in os.listdir(PATIENTS):
            rdr = os.path.join(PATIENTS, slug, "records")
            if not os.path.isdir(rdr): continue
            for rid in os.listdir(rdr):
                base = os.path.join(rdr, rid); tdir = os.path.join(base, "_thumbs")
                for key in ("exo_face_repos", "exo_face_sourire", "exo_profil"):
                    rel = "02_photos_traitees/%s.jpg" % key
                    full = os.path.join(base, rel)
                    if os.path.exists(full):
                        tp = os.path.join(tdir, rel.replace("/", "_") + ".t.jpg")
                        _gen_thumb(full, tp); _t.sleep(0.03)
                        break
    except Exception:
        pass


@app.route("/f/<slug>/<rid>/<path:path>")
def rfile(slug, rid, path):
    if not logged(): return redirect(url_for("login"))
    full = os.path.realpath(os.path.join(rdir(slug, rid), path))
    if not full.startswith(os.path.realpath(PATIENTS)) or not os.path.exists(full): abort(404)
    resp = send_file(full)
    resp.headers["Cache-Control"] = "private, no-cache"
    return resp


@app.route("/thumb/<slug>/<rid>/<path:path>")
def rthumb(slug, rid, path):
    if not logged(): return redirect(url_for("login"))
    base = rdir(slug, rid)
    full = os.path.realpath(os.path.join(base, path))
    if not full.startswith(os.path.realpath(PATIENTS)) or not os.path.exists(full): abort(404)
    tdir = os.path.join(base, "_thumbs")
    try: os.makedirs(tdir, exist_ok=True)
    except Exception: pass
    try: sz = int(request.args.get("s", "340"))
    except Exception: sz = 340
    if sz not in (340, 700): sz = 340
    tp = os.path.join(tdir, path.replace("/", "_") + (".t.jpg" if sz == 340 else ".m.jpg"))
    try:
        if (not os.path.exists(tp)) or (os.path.getmtime(tp) < os.path.getmtime(full)):
            from PIL import Image as _I, ImageOps as _IO
            im = _IO.exif_transpose(_I.open(full)).convert("RGB")
            im.thumbnail((sz, sz), _I.LANCZOS)
            im.save(tp, "JPEG", quality=(82 if sz == 340 else 88))
        resp = send_file(tp)
    except Exception:
        resp = send_file(full)
    resp.headers["Cache-Control"] = "private, no-cache"
    return resp

PRESENT_TPL = """<!doctype html><html lang=fr><head><meta charset=utf-8>
<meta name=viewport content="width=device-width,initial-scale=1"><title>Presentation &mdash; __NOM__</title>
<style>
*{box-sizing:border-box}
html,body{margin:0;height:100%;overflow:hidden;color:#eef2f8;
  font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;-webkit-font-smoothing:antialiased}
body{background:radial-gradient(1400px 900px at 50% -12%,#17223a 0%,#0d1220 46%,#080b12 100%)}
#progress{position:fixed;top:0;left:0;height:3px;width:0;z-index:60;
  background:linear-gradient(90deg,#2f7ff0,#57d98a);box-shadow:0 0 12px rgba(47,127,240,.5);transition:width .28s ease}
#topbar{position:fixed;top:0;left:0;right:0;height:52px;display:flex;align-items:center;gap:14px;
  padding:0 18px;z-index:40;background:linear-gradient(#0b0e14ee,#0b0e1400);opacity:.22;transition:opacity .2s}
#topbar:hover{opacity:1}
#topbar .nm{font-weight:600;font-size:15px;letter-spacing:.2px}
#topbar .sp{flex:1}
#topbar a,#topbar button{background:rgba(255,255,255,.10);color:#eef2f8;border:1px solid rgba(255,255,255,.18);
  border-radius:8px;padding:7px 13px;font-size:13px;cursor:pointer;text-decoration:none}
#topbar button.p{background:#2f7ff0;border-color:#2f7ff0;font-weight:600}
.pslide{position:fixed;inset:0;display:none;flex-direction:column;align-items:center;justify-content:center;
  padding:64px 8px 30px;text-align:center}
.pslide.on{display:flex;animation:fadein .22s ease}
@keyframes fadein{from{opacity:0;transform:scale(.994)}to{opacity:1;transform:none}}
.stitle{position:absolute;top:20px;left:50%;transform:translateX(-50%);max-width:66%;
  font-size:clamp(12px,1.75vh,17px);font-weight:700;color:#cfe0ff;letter-spacing:2.4px;text-transform:uppercase;
  z-index:45;pointer-events:none;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;
  padding-bottom:9px}
.stitle::after{content:"";position:absolute;left:50%;bottom:0;transform:translateX(-50%);
  width:46px;height:2px;border-radius:2px;background:linear-gradient(90deg,#2f7ff0,#57d98a)}
.pbody{flex:1;width:100%;display:flex;align-items:center;justify-content:center;min-height:0}
.pgrid{display:flex;gap:12px;align-items:center;justify-content:center;width:100%;height:100%}
.pfig{display:flex;flex:1 1 0;flex-direction:column;align-items:center;min-width:0;height:100%;justify-content:center}
.pfig img{max-height:88vh;max-width:100%;object-fit:contain;border-radius:10px;
  box-shadow:0 10px 40px rgba(0,0,0,.55),0 0 0 1px rgba(255,255,255,.05);cursor:zoom-in;background:#000}
.pgrid.n1 .pfig img{max-height:94vh}
.pfig figcaption{margin-top:10px;font-size:clamp(12px,1.55vh,16px);color:#c2d2ea;font-weight:600;letter-spacing:.4px}
/* Separateur de temps clinique */
.divider{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:16px;text-align:center}
.divider .dkick{font-size:clamp(12px,1.7vh,16px);font-weight:700;letter-spacing:4px;text-transform:uppercase;color:#7fa8e6}
.divider .dlabel{font-size:clamp(38px,8vh,86px);font-weight:800;letter-spacing:.5px;line-height:1.02;
  background:linear-gradient(120deg,#eaf1ff,#9fc0ff);-webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent}
.divider .dmeta{font-size:clamp(15px,2.4vh,25px);color:#9fb4d6}
.divider::before{content:"";width:70px;height:3px;border-radius:3px;background:linear-gradient(90deg,#2f7ff0,#57d98a);margin-bottom:6px}
/* Cover */
.cover{display:flex;flex-direction:column;align-items:center;gap:16px}
.cover .ckick{display:inline-block;font-size:clamp(12px,1.7vh,16px);font-weight:700;letter-spacing:4px;text-transform:uppercase;
  color:#bcd2ff;padding:7px 18px;border:1px solid rgba(120,160,255,.35);border-radius:999px;background:rgba(47,127,240,.10)}
.cover .cnom{font-size:clamp(34px,7.4vh,78px);font-weight:800;letter-spacing:.5px;
  background:linear-gradient(120deg,#ffffff,#bcd6ff);-webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent}
.cover .clabel{font-size:clamp(20px,3.4vh,34px);color:#4f9bff;font-weight:700}
.cover .cmeta{font-size:clamp(15px,2.4vh,24px);color:#9fb4d6}
.cover .cmotif{margin-top:20px;max-width:60ch;font-size:clamp(16px,2.6vh,26px);color:#dbe4f2;line-height:1.5;font-style:italic}
.cover::after{content:"";width:120px;height:3px;border-radius:3px;background:linear-gradient(90deg,#2f7ff0,#57d98a);margin-top:10px;opacity:.8}
/* Steiner */
.steiner{display:grid;grid-template-columns:1fr 1fr;gap:10px 46px;width:100%;max-width:1500px}
.srow{display:grid;grid-template-columns:1.4fr auto auto 1.6fr;align-items:center;gap:14px;
  padding:10px 16px;border-radius:10px;background:rgba(255,255,255,.045);font-size:clamp(14px,2.1vh,22px)}
.srow .sname{text-align:left;color:#cfdaec;font-weight:600}
.srow .sval{font-weight:700;font-variant-numeric:tabular-nums;font-size:1.12em}
.srow .snorm{color:#7f8ea8;font-size:.82em}
.srow .ssig{text-align:right;color:#aebbd2;font-size:.92em}
.srow.ok .dot,.srow .dot{width:12px;height:12px;border-radius:50%;display:inline-block}
.srow.ok .sval{color:#57d98a}.srow.warn .sval{color:#f2c14e}.srow.bad .sval{color:#f4746b}
.srow.warn{background:rgba(242,193,78,.09)}.srow.bad{background:rgba(244,116,107,.11)}
/* Diagnostic */
.diag{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:20px;width:100%;max-width:1500px}
.dcol{background:rgba(255,255,255,.045);border:1px solid rgba(255,255,255,.08);border-radius:14px;padding:18px 18px 8px}
.dcol h3{margin:0 0 12px;font-size:clamp(15px,2.2vh,22px);color:#4f9bff;text-transform:uppercase;letter-spacing:.4px}
.drow{display:flex;flex-direction:column;gap:2px;text-align:left;margin-bottom:12px}
.drow .drl{font-size:clamp(11px,1.5vh,14px);color:#8394b0;text-transform:uppercase;letter-spacing:.3px}
.drow .dv{font-size:clamp(15px,2.2vh,22px);color:#eef2f8}
/* nav */
.nav{position:fixed;top:0;bottom:0;width:9%;z-index:20;border:none;background:transparent;color:#fff;
  font-size:44px;cursor:pointer;opacity:0;transition:opacity .2s}
.nav:hover{opacity:.5;background:linear-gradient(90deg,#0b0e1466,#0b0e1400)}
.nav.r{right:0;background:linear-gradient(270deg,#0b0e1466,#0b0e1400)}
.nav.l{left:0}
#tpband{position:fixed;bottom:15px;left:20px;color:#8ea3c4;font-size:13px;z-index:30;letter-spacing:.4px;
  max-width:40vw;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;opacity:.8}
#count{position:fixed;bottom:16px;left:0;right:0;text-align:center;color:#6f809c;font-size:14px;z-index:30;
  font-variant-numeric:tabular-nums;letter-spacing:1px}
#hint{position:fixed;bottom:16px;right:20px;color:#5a6a86;font-size:12px;z-index:30}
#pzoom{position:fixed;inset:0;background:#000;display:none;align-items:center;justify-content:center;z-index:100;cursor:zoom-out}
#pzoom img{max-width:100%;max-height:100%;object-fit:contain}
</style></head><body>
<div id=progress></div>
<div id=topbar>
  <span class=nm>__NOM__</span>
  <span class=sp></span>
  <button class=p onclick="togFS()">&#9974; Plein &eacute;cran</button>
  <a href="__BACK__">&times; Quitter</a>
</div>
__SLIDES__
<button class="nav l" onclick="prev()">&#8249;</button>
<button class="nav r" onclick="next()">&#8250;</button>
<div id=tpband></div>
<div id=count></div>
<div id=hint>&#8592; &#8594; naviguer &middot; F plein &eacute;cran &middot; Echap quitter</div>
<div id=pzoom onclick="unzoom()"><img alt=""></div>
<script>
var BACK="__BACK__";
var slides=Array.prototype.slice.call(document.querySelectorAll('.pslide'));
var N=slides.length,idx=0;
function show(i){idx=(i+N)%N;for(var k=0;k<N;k++)slides[k].classList.toggle('on',k===idx);
  document.getElementById('count').textContent=(idx+1)+' / '+N;
  document.getElementById('progress').style.width=(N>1?(idx/(N-1)*100):100)+'%';
  var tp=slides[idx].getAttribute('data-tp')||'';document.getElementById('tpband').textContent=tp;}
function next(){show(idx+1);}function prev(){show(idx-1);}
function togFS(){if(document.fullscreenElement){document.exitFullscreen();}
  else if(document.documentElement.requestFullscreen){document.documentElement.requestFullscreen();}}
function zoom(img){var o=document.getElementById('pzoom');o.querySelector('img').src=img.src;o.style.display='flex';}
function unzoom(){document.getElementById('pzoom').style.display='none';}
document.addEventListener('keydown',function(e){
  var _ae=document.activeElement;
  if(_ae&&(_ae.tagName==='TEXTAREA'||_ae.tagName==='INPUT')){if(e.key==='Escape')_ae.blur();return;}
  if(document.getElementById('pzoom').style.display==='flex'){if(e.key==='Escape')unzoom();return;}
  if(e.key==='ArrowRight'||e.key===' '||e.key==='PageDown'){e.preventDefault();next();}
  else if(e.key==='ArrowLeft'||e.key==='PageUp'){e.preventDefault();prev();}
  else if(e.key==='Escape'){if(document.fullscreenElement){document.exitFullscreen();}else{location.href=BACK;}}
  else if(e.key==='f'||e.key==='F'){togFS();}});
/* --- Enregistrement automatique generique (diapos editables : Q/R, plan) --- */
(function(){
  var timers=new WeakMap();
  function body(form){var d=[];form.querySelectorAll('textarea[name]').forEach(function(t){d.push(encodeURIComponent(t.name)+'='+encodeURIComponent(t.value));});return d.join('&');}
  function save(form,beacon){var url=form.getAttribute('data-autosave-url');if(!url)return;var b=body(form);var st=form.querySelector('.autosave-status');
    if(beacon&&navigator.sendBeacon){try{navigator.sendBeacon(url,new Blob([b],{type:'application/x-www-form-urlencoded'}));}catch(e){}return;}
    if(st)st.textContent='Enregistrement...';fetch(url,{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:b}).then(function(r){if(st)st.textContent=r.ok?'Enregistré':'Erreur';}).catch(function(){if(st)st.textContent='Erreur';});}
  document.addEventListener('input',function(e){var f=e.target.closest&&e.target.closest('form[data-autosave-url]');if(!f)return;var st=f.querySelector('.autosave-status');if(st)st.textContent='…';if(timers.get(f))clearTimeout(timers.get(f));timers.set(f,setTimeout(function(){save(f,false);},700));});
  function saveAll(beacon){document.querySelectorAll('form[data-autosave-url]').forEach(function(f){save(f,beacon);});}
  window.addEventListener('beforeunload',function(){saveAll(true);});
  document.addEventListener('visibilitychange',function(){if(document.visibilityState==='hidden')saveAll(true);});
})();
/* --- Export de la decision du staff vers le suivi --- */
window.exportDecision=function(btn){
  var form=btn.closest('form');var ans=form.querySelector('textarea[name=answers]');
  var url=btn.getAttribute('data-exp-url');var msg=form.querySelector('.exp-status');
  var txt=ans?ans.value.trim():'';
  if(!txt){if(msg){msg.style.color='#f4746b';msg.textContent='Aucune décision à exporter.';}return;}
  if(msg){msg.style.color='#8ea3c4';msg.textContent='Export...';}btn.disabled=true;
  fetch(url,{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:'answers='+encodeURIComponent(txt)})
    .then(function(r){return r.text();}).then(function(t){
      if(msg){if(t==='ok'){msg.style.color='#57d98a';msg.textContent='Décision ajoutée au suivi ✓';}
        else if(t==='empty'){msg.style.color='#f4746b';msg.textContent='Aucune décision à exporter.';}
        else{msg.style.color='#f4746b';msg.textContent='Erreur.';}}
      setTimeout(function(){btn.disabled=false;},1600);})
    .catch(function(){if(msg){msg.style.color='#f4746b';msg.textContent='Erreur réseau.';}btn.disabled=false;});};
show(0);
</script>
</body></html>"""


@app.route("/record/<slug>/<rid>/presentation")
def presentation(slug, rid):
    if not logged(): return redirect(url_for("login"))
    pt = load_patient(slug); r = load_rec(slug, rid)
    if not pt or not r: abort(404)
    base = rdir(slug, rid)
    PLBL = dict(PHOTO_FIELDS)
    def psrc(rel):
        return url_for("rfile", slug=slug, rid=rid, path=rel)
    def photo_src(key):
        for rel in ("02_photos_traitees/%s.jpg" % key, "01_photos_brutes/%s.jpg" % key):
            if os.path.exists(os.path.join(base, rel)): return psrc(rel)
        return None
    S = []
    def sl(title, inner):
        t = ('<div class=stitle>%s</div>' % title) if title else ""
        return '<section class=pslide>%s<div class=pbody>%s</div></section>' % (t, inner)
    # 1) Cover
    label = r.get("label", "") or "Bilan"
    date = (r.get("date", "") or "")[:10]
    motif = (r.get("motif") or "").strip()
    meta = "%s &middot; %s ans" % (pt.get("dob_court", ""), pt.get("age", ""))
    if date: meta += " &middot; " + date
    cover = ('<div class=cover><div class=cnom>%s</div><div class=clabel>%s</div>'
             '<div class=cmeta>%s</div>%s</div>') % (
        pt.get("nom", ""), label, meta, ('<div class=cmotif>&laquo;&nbsp;%s&nbsp;&raquo;</div>' % motif) if motif else "")
    S.append(sl("", cover))
    # 2) Photos
    def photo_slide(title, keys):
        cells = []
        for key in keys:
            src = photo_src(key)
            if src:
                cells.append('<figure class=pfig><img loading=lazy src="%s" onclick="zoom(this)">'
                             '<figcaption>%s</figcaption></figure>' % (src, PLBL.get(key, key)))
        if not cells: return
        S.append(sl(title, '<div class="pgrid n%d">%s</div>' % (len(cells), "".join(cells))))
    photo_slide("Photographies exobuccales", ["exo_face_repos", "exo_face_sourire", "exo_profil"])
    photo_slide("Vues endobuccales", ["endo_laterale_droite", "endo_occlusion_frontale", "endo_laterale_gauche"])
    photo_slide("Vues occlusales", ["endo_occlusal_maxillaire", "endo_occlusal_mandibulaire"])
    # 3) Radios
    rad_cells = []
    for k, lbl in (("panoramique", "Panoramique"), ("teleradiographie_profil", "T&eacute;l&eacute;radiographie de profil")):
        rel = "03_radios/%s.jpg" % k
        if os.path.exists(os.path.join(base, rel)):
            rad_cells.append('<figure class=pfig><img loading=lazy src="%s" onclick="zoom(this)">'
                             '<figcaption>%s</figcaption></figure>' % (psrc(rel), lbl))
    if rad_cells:
        S.append(sl("Radiographies", '<div class="pgrid n%d">%s</div>' % (len(rad_cells), "".join(rad_cells))))
    # 4) Traces cephalo
    wc = os.path.join(base, "06_webceph"); tr_cells = []
    if os.path.isdir(wc):
        for tf in sorted(os.listdir(wc)):
            if tf.lower().endswith((".jpg", ".jpeg", ".png")):
                tr_cells.append('<figure class=pfig><img loading=lazy src="%s" onclick="zoom(this)"></figure>'
                                % psrc("06_webceph/%s" % tf))
    if tr_cells:
        S.append(sl("Trac&eacute; c&eacute;phalom&eacute;trique", '<div class="pgrid n%d">%s</div>' % (min(len(tr_cells), 2), "".join(tr_cells[:2]))))
    # 5) Steiner
    st = r.get("steiner") or {}
    srows = []
    for nom, short, mean, sd in STEINER_FIELDS:
        if nom in st and st[nom] not in (None, ""):
            try: v = float(st[nom])
            except Exception: continue
            z = (v - mean) / sd if sd else 0
            cls = "ok" if abs(z) <= 1 else ("warn" if abs(z) <= 2 else "bad")
            sig, sev = steiner_signif(nom, v, mean, sd)
            srows.append('<div class="srow %s"><span class=sname>%s</span><span class=sval>%g</span>'
                         '<span class=snorm>norme %g</span><span class=ssig>%s</span></div>'
                         % (cls, short, v, mean, sig if sev else "Dans la norme"))
    if srows:
        S.append(sl("Analyse c&eacute;phalom&eacute;trique de Steiner", '<div class=steiner>%s</div>' % "".join(srows)))
    # 6) Diagnostic
    syn = r.get("synthese", {}) or {}
    dcols = []
    for cl, cc in DIAG_COLS:
        items = []
        for rl, rc in DIAG_ROWS:
            val = (syn.get(cc + "_" + rc) or "").strip()
            if val: items.append('<div class=drow><span class=drl>%s</span><span class=dv>%s</span></div>' % (rl, val))
        if items: dcols.append('<div class=dcol><h3>%s</h3>%s</div>' % (cl, "".join(items)))
    if dcols:
        S.append(sl("Synth&egrave;se diagnostique", '<div class=diag>%s</div>' % "".join(dcols)))
    slides_html = "".join(S)
    html = (PRESENT_TPL.replace("__NOM__", pt.get("nom", ""))
            .replace("__BACK__", url_for("record", slug=slug, rid=rid))
            .replace("__SLIDES__", slides_html))
    return html


@app.route("/open/<slug>/<rid>")
def openfolder(slug, rid):
    if not logged(): return redirect(url_for("login"))
    import subprocess, sys
    folder = rdir(slug, rid)
    try:
        if sys.platform == "darwin": subprocess.Popen(["open", folder])
        elif sys.platform.startswith("win"): os.startfile(folder)
        else: subprocess.Popen(["xdg-open", folder])
    except Exception: pass
    return redirect(url_for("record", slug=slug, rid=rid))

import appareils_ext
try:
    appareils_ext.register(app, {"page": page, "DATA": DATA, "logged": logged})
except Exception as _e:
    print("appareils register failed:", _e)

if __name__ == "__main__":
    import webbrowser, threading
    print("Bibliothèque locale :", DATA)
    try: threading.Timer(1.2, lambda: webbrowser.open("http://127.0.0.1:5000")).start()
    except Exception: pass
    app.run(host="127.0.0.1", port=5000, debug=False)
