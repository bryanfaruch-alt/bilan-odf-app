# -*- coding: utf-8 -*-
"""Dossier orthodontique en PDF — 100% pur Python (ReportLab), aucune dépendance système.
Photos = EXACTEMENT celles du dossier (02_photos_traitees). Importe app pour les données."""
import os
import datetime as _dt
import app as A
from PIL import Image as PILImage
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.colors import HexColor, white, Color
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
                                Image as RLImage, PageBreak, KeepTogether, Flowable, HRFlowable)
from reportlab.pdfgen import canvas

NAVY = HexColor(0x16324F)
NAVY2 = HexColor(0x27548C)
ACCENT = HexColor(0x2F5CA8)
GREY = HexColor(0x6B7280)
LINE = HexColor(0xD7E2F0)
ALT = HexColor(0xF6F9FD)
G_BG, O_BG, R_BG = HexColor(0xE6F4E7), HexColor(0xFCECD9), HexColor(0xF9DEDC)
G_TX, O_TX, R_TX = HexColor(0x1E7D32), HexColor(0xB05A00), HexColor(0xB02828)

PAGE_W, PAGE_H = A4
LM = RM = 1.2 * cm
TM = 1.5 * cm
BM = 1.3 * cm
AVAIL = PAGE_W - LM - RM

_st_name = ParagraphStyle("nm", fontName="Helvetica-Bold", fontSize=24, textColor=ACCENT,
                          alignment=TA_CENTER, spaceBefore=10, spaceAfter=4, leading=28)
_st_info = ParagraphStyle("info", fontName="Helvetica", fontSize=12, textColor=GREY,
                          alignment=TA_CENTER, spaceAfter=2, leading=16)
_st_prac = ParagraphStyle("prac", fontName="Helvetica-Bold", fontSize=12.5, textColor=NAVY,
                          alignment=TA_CENTER, leading=15)
_st_date = ParagraphStyle("dt", fontName="Helvetica", fontSize=10, textColor=GREY,
                          alignment=TA_CENTER, leading=13)
_st_sub = ParagraphStyle("sub", fontName="Helvetica-Bold", fontSize=8.5, textColor=ACCENT,
                         spaceBefore=8, spaceAfter=1, leading=11)
_st_cap = ParagraphStyle("cap", fontName="Helvetica-Oblique", fontSize=8, textColor=GREY,
                         alignment=TA_CENTER, leading=10)
_st_body = ParagraphStyle("body", fontName="Helvetica", fontSize=10, textColor=HexColor(0x1F2A37),
                          leading=13.5, spaceAfter=2)
_st_lbl = ParagraphStyle("lbl", fontName="Helvetica-Bold", fontSize=10, textColor=HexColor(0x1F2A37),
                         spaceBefore=3, spaceAfter=1, leading=13)
_st_crd = ParagraphStyle("crd", fontName="Helvetica-Bold", fontSize=9.5, textColor=ACCENT, leading=13)


def _esc(s):
    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _ml(s):
    return _esc(s).replace("\n", "<br/>")


def _pth(base, slot):
    for rel in ("02_photos_traitees/%s.jpg" % slot, "01_photos_brutes/%s.jpg" % slot):
        p = os.path.join(base, rel)
        if os.path.exists(p):
            return p
    return None


def _img(path, width, maxh=250):
    try:
        iw, ih = PILImage.open(path).size
    except Exception:
        return None
    h = width * ih / float(iw)
    w = width
    if h > maxh:
        h = maxh; w = h * iw / float(ih)
    return RLImage(path, width=w, height=h)


class Band(Flowable):
    def __init__(self, text, sub="", height=30, big=False):
        Flowable.__init__(self)
        self.text = text; self.sub = sub; self.height = height; self.big = big
        self.spaceBefore = 12; self.spaceAfter = 7

    def wrap(self, aw, ah):
        self.width = aw
        return (aw, self.height)

    def draw(self):
        c = self.canv
        c.saveState()
        c.setFillColor(NAVY)
        c.roundRect(0, 0, self.width, self.height, 4, stroke=0, fill=1)
        c.setFillColor(white)
        if self.big:
            c.setFont("Helvetica-Bold", 25)
            c.drawCentredString(self.width / 2.0, self.height / 2.0 + 4, self.text)
            if self.sub:
                c.setFont("Helvetica", 10)
                c.setFillColor(Color(1, 1, 1, 0.85))
                c.drawCentredString(self.width / 2.0, self.height / 2.0 - 14, self.sub)
        else:
            c.setFont("Helvetica-Bold", 13)
            c.drawString(12, self.height / 2.0 - 4, self.text)
            if self.sub:
                tw = c.stringWidth(self.text, "Helvetica-Bold", 13)
                c.setFont("Helvetica", 10); c.setFillColor(Color(1, 1, 1, 0.82))
                c.drawString(12 + tw + 8, self.height / 2.0 - 4, self.sub)
        c.restoreState()


def _sub(text):
    return [Paragraph(text.upper(), _st_sub),
            HRFlowable(width="100%", thickness=0.5, color=LINE, spaceBefore=1, spaceAfter=3)]


def _grid(items, ncols):
    items = [(p, c) for p, c in items if p and os.path.exists(p)]
    if not items:
        return []
    cellw = AVAIL / ncols
    imgw = cellw - 8
    out = []
    for i in range(0, len(items), ncols):
        chunk = items[i:i + ncols]
        imgrow, caprow = [], []
        for p, cap in chunk:
            im = _img(p, imgw)
            imgrow.append(im if im else "")
            caprow.append(Paragraph(_esc(cap), _st_cap))
        while len(imgrow) < ncols:
            imgrow.append(""); caprow.append("")
        t = Table([imgrow, caprow], colWidths=[cellw] * ncols)
        t.setStyle(TableStyle([('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                               ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                               ('TOPPADDING', (0, 0), (-1, 0), 2),
                               ('BOTTOMPADDING', (0, 0), (-1, 0), 1),
                               ('TOPPADDING', (0, 1), (-1, 1), 0),
                               ('BOTTOMPADDING', (0, 1), (-1, 1), 8)]))
        out.append(KeepTogether(t))
    return out


def _steiner(r):
    stv = r.get("steiner") or {}
    data = [["Mesure", "Valeur", "Norme", "Écart"]]
    styles = [('BACKGROUND', (0, 0), (-1, 0), ACCENT), ('TEXTCOLOR', (0, 0), (-1, 0), white),
              ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'), ('FONTSIZE', (0, 0), (-1, -1), 9),
              ('ALIGN', (1, 0), (-1, -1), 'CENTER'), ('ALIGN', (0, 0), (0, -1), 'LEFT'),
              ('LINEBELOW', (0, 0), (-1, -1), 0.3, LINE),
              ('TOPPADDING', (0, 0), (-1, -1), 3.5), ('BOTTOMPADDING', (0, 0), (-1, -1), 3.5),
              ('LEFTPADDING', (0, 0), (-1, -1), 6), ('RIGHTPADDING', (0, 0), (-1, -1), 6)]
    ri = 1
    for label, nom, mean, sd in A.STEINER_FIELDS:
        v = stv.get(label)            # cle de stockage = libelle complet, ex. "SNA (°)"
        if v in (None, ""):
            v = stv.get(nom)          # repli sur le nom court par securite
        if v in (None, ""):
            continue
        try:
            fv = float(v)
        except Exception:
            continue
        dev = fv - mean; ad = abs(dev)
        if ad <= sd: bg, tx, tag = G_BG, G_TX, "normal"
        elif ad <= 2 * sd: bg, tx, tag = O_BG, O_TX, "limite"
        else: bg, tx, tag = R_BG, R_TX, "hors norme"
        data.append([label, "%g" % fv, "%g ± %g" % (mean, sd), "%+.1f · %s" % (dev, tag)])
        styles += [('BACKGROUND', (1, ri), (1, ri), bg), ('TEXTCOLOR', (1, ri), (1, ri), tx),
                   ('FONTNAME', (1, ri), (1, ri), 'Helvetica-Bold'), ('TEXTCOLOR', (3, ri), (3, ri), tx)]
        ri += 1
    if ri == 1:
        return []
    t = Table(data, colWidths=[AVAIL * 0.40, AVAIL * 0.18, AVAIL * 0.22, AVAIL * 0.20])
    t.setStyle(TableStyle(styles))
    return _sub("Analyse de Steiner") + [t]


def _motif(r):
    m = (r.get("motif") or "").strip()
    return (_sub("Motif de consultation") + [Paragraph(_ml(m), _st_body)]) if m else []


def _plan(r):
    plan = r.get("plan") or {}
    if not isinstance(plan, dict):
        plan = {}
    obj = (plan.get("objectifs") or "").strip()
    moy = (plan.get("moyens") or "").strip()
    if not obj and not moy:
        return []
    out = _sub("Plan de traitement")
    if obj:
        out.append(Paragraph("Objectifs", _st_lbl))
        for l in obj.splitlines():
            if l.strip():
                out.append(Paragraph("• " + _esc(l), _st_body))
    if moy:
        out.append(Paragraph("Moyens thérapeutiques", _st_lbl))
        for l in moy.splitlines():
            if l.strip():
                out.append(Paragraph("• " + _esc(l), _st_body))
    return out


def _crs(pt):
    notes = sorted((pt.get("notes") or []), key=lambda n: (n.get("date", "") or ""))
    notes = [n for n in notes if (n.get("text") or "").strip()]
    if not notes:
        return []
    data = []
    for n in notes:
        d = A._date_fr(n.get("date", "")) or (n.get("date", "") or "")[:10] or "—"
        data.append([Paragraph(_esc(d), _st_crd), Paragraph(_ml((n.get("text") or "").strip()), _st_body)])
    t = Table(data, colWidths=[AVAIL * 0.22, AVAIL * 0.78])
    t.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'),
                           ('LINEBELOW', (0, 0), (-1, -1), 0.3, HexColor(0xEEF2F7)),
                           ('TOPPADDING', (0, 0), (-1, -1), 4), ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                           ('LEFTPADDING', (0, 0), (0, -1), 0)]))
    return [Band("Comptes rendus")] + [t]


def _temp(slug, idx, rid, r, pt):
    base = A.rdir(slug, rid)
    title = "%s — %s" % (A._rec_label(idx, r), A._date_fr(r.get("date", "")) or "date non renseignée")
    age = A.age_at(pt.get("dob", ""), r.get("date", ""))
    fl = [Band(title, sub=("· %s" % age) if age else "")]
    fl += _motif(r)
    exo = [(_pth(base, k), c) for k, c in [("exo_face_repos", "Visage au repos"),
                                           ("exo_face_sourire", "Sourire"), ("exo_profil", "Profil")]]
    endo = [(_pth(base, k), c) for k, c in [("endo_laterale_droite", "Latérale droite"),
                                            ("endo_occlusion_frontale", "Occlusion frontale"),
                                            ("endo_laterale_gauche", "Latérale gauche")]]
    occl = [(_pth(base, k), c) for k, c in [("endo_occlusal_maxillaire", "Occlusal maxillaire"),
                                            ("endo_occlusal_mandibulaire", "Occlusal mandibulaire")]]
    if any(p for p, _ in exo):
        fl += _sub("Photographies exobuccales") + _grid(exo, 3)
    if any(p for p, _ in endo):
        fl += _sub("Photographies endobuccales") + _grid(endo, 3)
    if any(p for p, _ in occl):
        fl += _sub("Vues occlusales") + _grid(occl, 2)
    rad = [(os.path.join(base, "03_radios", "panoramique.jpg"), "Panoramique"),
           (os.path.join(base, "03_radios", "teleradiographie_profil.jpg"), "Téléradiographie de profil")]
    rad = [(p, c) for p, c in rad if os.path.exists(p)]
    if rad:
        fl += _sub("Radiographies") + _grid(rad, len(rad))
    rend = [(os.path.join(base, "05_stl_rendus", k + ".png"), c) for k, c in A.RENDER_FIELDS]
    rend = [(p, c) for p, c in rend if os.path.exists(p)]
    if rend:
        fl += _sub("Modèles 3D") + _grid(rend, 3)
    wc = os.path.join(base, "06_webceph"); wp = []
    if os.path.isdir(wc):
        for f in sorted(os.listdir(wc)):
            if f.lower().endswith(".jpg"):
                wp.append((os.path.join(wc, f), "Tracé céphalométrique"))
    if wp:
        fl += _sub("Tracés céphalométriques (WebCeph)") + _grid(wp, 2)
    fl += _steiner(r)
    fl += _plan(r)
    return fl


def _make_numbered(patient_name):
    class NumberedCanvas(canvas.Canvas):
        def __init__(self, *a, **k):
            canvas.Canvas.__init__(self, *a, **k)
            self._saved = []

        def showPage(self):
            self._saved.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            total = len(self._saved)
            for st in self._saved:
                self.__dict__.update(st)
                self._deco(total)
                canvas.Canvas.showPage(self)
            canvas.Canvas.save(self)

        def _deco(self, total):
            pg = self._pageNumber
            if pg <= 1:
                return
            self.saveState()
            self.setFont("Helvetica", 7); self.setFillColor(GREY)
            self.drawRightString(PAGE_W - RM, PAGE_H - 26, "CHU de Nice — Orthopédie dento-faciale")
            self.drawString(LM, 18, patient_name)
            self.drawCentredString(PAGE_W / 2.0, 18, "Confidentiel — usage médical")
            self.drawRightString(PAGE_W - RM, 18, "Page %d / %d" % (pg, total))
            self.restoreState()
    return NumberedCanvas


def build_pdf(slug, out_path, practitioner="Dr Bryan Faruch", logo=None):
    pt = A.load_patient(slug)
    if not pt:
        raise ValueError("patient introuvable: %s" % slug)
    recs = A._ordered_records(slug, pt)
    story = []
    # --- couverture ---
    if logo and os.path.exists(logo):
        im = _img(logo, 130, maxh=90)
        if im:
            im.hAlign = 'RIGHT'; story.append(im)
    story.append(Spacer(1, 70))
    story.append(Band("DOSSIER ORTHODONTIQUE", sub="CHU DE NICE", height=64, big=True))
    story.append(Spacer(1, 26))
    story.append(Paragraph(_esc(pt.get("nom", "")), _st_name))
    story.append(Paragraph("Né(e) %s&nbsp;&nbsp;·&nbsp;&nbsp;%s&nbsp;&nbsp;·&nbsp;&nbsp;%s"
                           % (_esc(pt.get("dob_court", "—")), _esc(pt.get("sexe", "—")),
                              _esc(pt.get("age", "—"))), _st_info))
    story.append(Spacer(1, 150))
    story.append(HRFlowable(width="34%", thickness=0.8, color=ACCENT, hAlign='CENTER', spaceAfter=8))
    story.append(Paragraph(_esc(practitioner), _st_prac))
    story.append(Paragraph("Édité le %s" % _dt.date.today().strftime("%d/%m/%Y"), _st_date))
    story.append(PageBreak())
    # --- temps ---
    for idx, (rid, r) in enumerate(recs):
        story += _temp(slug, idx, rid, r, pt)
    story += _crs(pt)

    doc = SimpleDocTemplate(out_path, pagesize=A4, leftMargin=LM, rightMargin=RM,
                            topMargin=TM, bottomMargin=BM, title="Dossier orthodontique — %s" % pt.get("nom", ""))
    doc.build(story, canvasmaker=_make_numbered(pt.get("nom", "")))
    return out_path
