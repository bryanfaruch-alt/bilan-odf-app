# -*- coding: utf-8 -*-
"""Bibliotheque d'appareils orthodontiques a montrer au patient.
Module autonome branche sur l'app Flask existante via register(app, ctx).
Donnees dans DATA/_appareils.json ; photos dans DATA/_appareils_img/<id>.jpg.
Illustration SVG par defaut, remplacable par une photo."""
import os, json, re, datetime, shutil
from flask import request, session, redirect, url_for, abort, send_file

APPAREILS_CATS = [
    ("fixe", "Appareils fixes"),
    ("amovible", "Appareils amovibles"),
    ("aligneurs", "Aligneurs"),
    ("contention", "Contention"),
    ("extra", "Extra-oral"),
]

SEED = [
    ('bagues_metal', 'fixe', 'bracket', 'Bagues métalliques',
     ''),
    ('quad_helix', 'fixe', 'quad', 'Quad hélix',
     ''),
    ('disjoncteur', 'fixe', 'expander', 'Disjoncteur (expansion du palais)',
     ''),
    ('plaque_amovible', 'amovible', 'plate', 'Plaque amovible',
     ''),
    ('activateur', 'amovible', 'activator', 'Appareil fonctionnel (activateur)',
     ''),
    ('aligneurs', 'aligneurs', 'aligner', 'Gouttières transparentes (aligneurs)',
     ''),
    ('fil_contention', 'contention', 'wire', 'Fil de contention collé',
     ''),
    ('gouttiere_contention', 'contention', 'tray', 'Gouttière de contention',
     ''),
    ('masque_delaire', 'extra', 'facemask', 'Masque facial (Delaire)',
     ''),
    ('fronde', 'extra', 'chincup', 'Fronde mentonnière',
     ''),
    ('feo', 'extra', 'headgear', 'Force extra-orale (headgear)',
     ''),
    ('pendulum', 'fixe', 'quad', 'Pendulum',
     ''),
    ('arc_lingual', 'fixe', 'quad', 'Arc lingual',
     ''),
    ('arc_transpalatin', 'fixe', 'quad', 'Arc transpalatin',
     ''),
    ('nance', 'fixe', 'quad', 'Gosh Nance',
     ''),
    ('plaque_verin_schwarz', 'amovible', 'plate', 'Plaque à vérin (ressort de Schwarz)',
     ''),
    ('forsus', 'fixe', 'bracket', 'Forsus',
     ''),
    ('bouton_metal', 'fixe', 'bracket', 'Bouton métallique collé',
     ''),
    ('lip_bumper', 'fixe', 'quad', 'Lip bumper',
     ''),
    ('herbst', 'fixe', 'bracket', 'Bielle de Herbst',
     ''),
    ('carriere_motion', 'fixe', 'bracket', 'Carrière Motion',
     ''),
    ('elastiques_classe_2_3', 'fixe', 'bracket', 'Élastiques de classe II / III',
     ''),
]

# --- Illustrations SVG (stroke = currentColor -> suivent le theme) ---
def _svg(inner):
    return ('<svg viewBox="0 0 120 96" fill="none" stroke="currentColor" stroke-width="3" '
            'stroke-linecap="round" stroke-linejoin="round">%s</svg>') % inner

_TEETH = ('<path d="M30 30h16v20a8 8 0 0 1-16 0Z" fill="currentColor" opacity=".07"/>'
          '<path d="M52 30h16v20a8 8 0 0 1-16 0Z" fill="currentColor" opacity=".07"/>'
          '<path d="M74 30h16v20a8 8 0 0 1-16 0Z" fill="currentColor" opacity=".07"/>')
_PROFILE = ('<path d="M74 20c-15 0-25 12-25 27 0 6 3 9 3 15s-6 6-6 13"/>')

SVGS = {
    "bracket": _svg('<line x1="16" y1="42" x2="104" y2="42"/>' + _TEETH +
        '<rect x="34" y="37" width="8" height="8" rx="1.5" fill="currentColor" opacity=".35"/>'
        '<rect x="56" y="37" width="8" height="8" rx="1.5" fill="currentColor" opacity=".35"/>'
        '<rect x="78" y="37" width="8" height="8" rx="1.5" fill="currentColor" opacity=".35"/>'),
    "bracketc": _svg('<line x1="16" y1="42" x2="104" y2="42" opacity=".5"/>' + _TEETH +
        '<rect x="34" y="37" width="8" height="8" rx="1.5" opacity=".5"/>'
        '<rect x="56" y="37" width="8" height="8" rx="1.5" opacity=".5"/>'
        '<rect x="78" y="37" width="8" height="8" rx="1.5" opacity=".5"/>'),
    "lingual": _svg(_TEETH +
        '<path d="M30 58c14 8 46 8 60 0" stroke-dasharray="4 4"/>'
        '<rect x="46" y="52" width="7" height="7" rx="1.5" opacity=".35" fill="currentColor"/>'
        '<rect x="67" y="52" width="7" height="7" rx="1.5" opacity=".35" fill="currentColor"/>'),
    "quad": _svg('<path d="M30 28h60v6a30 30 0 0 1-60 0Z" fill="currentColor" opacity=".07"/>'
        '<path d="M44 40c-7 0-7 9 0 9M76 40c7 0 7 9 0 9M50 52c-7 0-7 9 0 9M70 52c7 0 7 9 0 9"/>'
        '<path d="M44 44h32M50 56h20"/>'),
    "expander": _svg('<path d="M28 28h64v8a26 26 0 0 1-26 26h-12a26 26 0 0 1-26-26Z" fill="currentColor" opacity=".07"/>'
        '<rect x="51" y="38" width="18" height="15" rx="2"/><line x1="60" y1="38" x2="60" y2="53"/>'
        '<line x1="51" y1="45" x2="69" y2="45"/>'),
    "spacer": _svg('<rect x="32" y="32" width="20" height="28" rx="3" fill="currentColor" opacity=".07"/>'
        '<path d="M52 40c16 0 16 12 0 12"/><rect x="72" y="34" width="16" height="24" rx="3"/>'),
    "plate": _svg('<path d="M26 34h68v6a26 26 0 0 1-26 26h-16a26 26 0 0 1-26-26Z" fill="currentColor" opacity=".1"/>'
        '<path d="M30 34c-5-7-12-4-11 3M90 34c5-7 12-4 11 3"/><line x1="60" y1="40" x2="60" y2="60"/>'),
    "activator": _svg('<path d="M28 34c0 20 14 28 32 28s32-8 32-28"/>'
        '<path d="M28 62c0-20 14 28 32-28s32 8 32 28"/>'),
    "aligner": _svg('<path d="M24 26c0 34 14 48 36 48s36-14 36-48"/>'
        '<path d="M34 30c0 28 10 40 26 40s26-12 26-40" opacity=".45"/>'),
    "wire": _svg('<path d="M30 30h60v6a30 30 0 0 1-60 0Z" fill="currentColor" opacity=".08"/>'
        '<path d="M34 50c12 7 40 7 52 0"/>'),
    "tray": _svg('<path d="M26 28c0 33 14 47 34 47s34-14 34-47"/>'
        '<path d="M26 28h68" opacity=".45"/>'),
    "facemask": _svg(_PROFILE + '<line x1="38" y1="14" x2="38" y2="82"/>'
        '<path d="M38 24h14"/><path d="M38 74h17"/><line x1="38" y1="49" x2="52" y2="49" opacity=".5"/>'),
    "chincup": _svg(_PROFILE + '<path d="M50 72a13 9 0 0 0 22 0"/><path d="M72 68l16-30"/>'
        '<path d="M50 40l14-14" opacity=".5"/>'),
    "headgear": _svg(_PROFILE + '<line x1="52" y1="49" x2="98" y2="43"/>'
        '<path d="M98 43c7 6 7 20 0 26"/>'),
}
def svg_for(key):
    return SVGS.get(key, SVGS["bracket"])

def _slug(s):
    s = re.sub(r"[^a-z0-9]+", "_", (s or "").lower()).strip("_")
    return s or ("app_" + datetime.datetime.now().strftime("%H%M%S"))


def register(app, ctx):
    page = ctx["page"]; DATA = ctx["DATA"]; logged = ctx["logged"]
    FILE = os.path.join(DATA, "_appareils.json")
    IMGDIR = os.path.join(DATA, "_appareils_img")

    def load_items():
        if os.path.exists(FILE):
            try: return json.load(open(FILE, encoding="utf-8")).get("items", [])
            except Exception: pass
        # 1re utilisation : amorcer depuis les valeurs par defaut fournies (photos incluses)
        _defdir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "appareils_defaults")
        _defjson = os.path.join(_defdir, "_appareils.json")
        if os.path.exists(_defjson):
            try:
                _items = (json.load(open(_defjson, encoding="utf-8")) or {}).get("items", [])
                _dimg = os.path.join(_defdir, "_appareils_img")
                if os.path.isdir(_dimg):
                    os.makedirs(IMGDIR, exist_ok=True)
                    for _f in os.listdir(_dimg):
                        try: shutil.copy2(os.path.join(_dimg, _f), os.path.join(IMGDIR, _f))
                        except Exception: pass
                save_items(_items)
                return _items
            except Exception: pass
        items = [{"id": i, "cat": c, "svg": s, "nom": n, "desc": d} for (i, c, s, n, d) in SEED]
        save_items(items)
        return items

    def save_items(items):
        os.makedirs(DATA, exist_ok=True)
        tmp = FILE + ".tmp"
        json.dump({"items": items}, open(tmp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        os.replace(tmp, FILE)

    def get_item(items, iid):
        return next((x for x in items if x.get("id") == iid), None)

    def has_photo(iid):
        return os.path.exists(os.path.join(IMGDIR, iid + ".jpg"))

    def esc(x): return str(x or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    def escat(x): return esc(x).replace('"', "&quot;")

    def visual(it, big=False):
        iid = it.get("id", "")
        if has_photo(iid):
            u = url_for("appareil_img", iid=iid) + ("?t=%d" % int(os.path.getmtime(os.path.join(IMGDIR, iid + ".jpg"))))
            return '<img src="%s" class="apimg%s" alt="">' % (u, " big" if big else "")
        return '<div class="apsvg%s">%s</div>' % (" big" if big else "", svg_for(it.get("svg")))

    CSS = """<style>
    .apintro{color:var(--mut);font-size:14px;margin:2px 0 20px;max-width:70ch}
    .apcat{font-size:12px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:var(--acc2);margin:26px 0 12px}
    .apgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:18px}
    .apcard{background:var(--card);border:1px solid var(--line);border-radius:16px;overflow:hidden;box-shadow:var(--sh-sm);transition:.16s;display:flex;flex-direction:column}
    .apcard:hover{box-shadow:var(--sh);transform:translateY(-3px);border-color:var(--acc)}
    .apthumb{aspect-ratio:4/3;background:var(--card2);display:flex;align-items:center;justify-content:center;cursor:pointer;overflow:hidden}
    .apimg{width:100%;height:100%;object-fit:cover}
    .apsvg{width:58%;color:var(--mut)}
    .apbody{padding:13px 15px 14px}
    .apnom{font-weight:650;font-size:15px;letter-spacing:-.01em}
    .apdesc{font-size:12.5px;color:var(--mut);margin-top:4px;line-height:1.45;display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden}
    .apfoot{margin-top:10px;display:flex;gap:8px}
    .aplink{font-size:12.5px;font-weight:600;color:var(--acc2);cursor:pointer}
    .aplink.mut{color:var(--mut)}
    </style>"""

    @app.route("/appareils")
    def appareils_home():
        if not logged(): return redirect(url_for("login"))
        items = load_items()
        body = [CSS,
            '<div style="display:flex;align-items:flex-end;justify-content:space-between;gap:14px;flex-wrap:wrap">',
            '<div><h1>Appareils</h1><p class=apintro>Une bibliothèque visuelle des appareils, à montrer au patient pour lui expliquer ce qu’on va lui poser. Clique une carte pour l’afficher en grand.</p></div>',
            '<a class=btn href="%s">+ Ajouter un appareil</a>' % url_for("appareil_edit", iid="new"),
            '</div>',
            '<input id=apsearch type=search placeholder="Rechercher un appareil\u2026" oninput="apFilter()" autocomplete=off style="width:100%;max-width:460px;padding:10px 14px;border:1px solid var(--line);border-radius:10px;margin:2px 0 20px;font-size:14px;background:var(--card);color:var(--ink)">']
        by = {}
        for it in items: by.setdefault(it.get("cat", "fixe"), []).append(it)
        for cat, lbl in APPAREILS_CATS:
            lst = by.get(cat) or []
            if not lst: continue
            body.append('<div class=apsec><div class=apcat>%s</div><div class=apgrid>' % esc(lbl))
            for it in lst:
                iid = it.get("id", "")
                show = url_for("appareil_show", iid=iid); edit = url_for("appareil_edit", iid=iid)
                crl = ('<a class="aplink mut" href="%s">\u2702 Recadrer</a>' % url_for("appareil_crop", iid=iid)) if has_photo(iid) else ''
                body.append(
                    '<div class=apcard data-nom="%s">'
                    '<div class=apthumb onclick="location.href=\'%s\'">%s</div>'
                    '<div class=apbody><div class=apnom>%s</div>'
                    '<div class=apfoot><a class=aplink href="%s">\U0001f5a5 Montrer</a>'
                    '<a class="aplink mut" href="%s">Modifier</a>%s</div></div></div>'
                    % (escat((it.get("nom") or "").lower()), show, visual(it), esc(it.get("nom")), show, edit, crl))
            body.append('</div></div>')
        body.append('<script>function apFilter(){var q=(document.getElementById("apsearch").value||"").toLowerCase().trim();document.querySelectorAll(".apsec").forEach(function(sec){var any=false;sec.querySelectorAll(".apcard").forEach(function(c){var n=(c.getAttribute("data-nom")||"");var sh=!q||n.indexOf(q)>=0;c.style.display=sh?"":"none";if(sh)any=true;});sec.style.display=any?"":"none";});}</script>')
        return page("".join(body), wide=True)

    @app.route("/appareils/img/<iid>")
    def appareil_img(iid):
        if not logged(): return redirect(url_for("login"))
        p = os.path.join(IMGDIR, iid + ".jpg")
        if not os.path.exists(p): abort(404)
        resp = send_file(p); resp.headers["Cache-Control"] = "private, no-cache"; return resp

    @app.route("/appareils/<iid>/show")
    def appareil_show(iid):
        if not logged(): return redirect(url_for("login"))
        items = load_items(); it = get_item(items, iid)
        if not it: abort(404)
        vis = visual(it, big=True)
        html = SHOW_TPL.replace("__NOM__", escat(it.get("nom"))) \
            .replace("__VIS__", vis).replace("__DESC__", esc(it.get("desc")).replace("\n", "<br>")) \
            .replace("__BACK__", url_for("appareils_home"))
        return html

    @app.route("/appareils/<iid>/edit", methods=["GET", "POST"])
    def appareil_edit(iid):
        if not logged(): return redirect(url_for("login"))
        items = load_items()
        is_new = (iid == "new")
        it = None if is_new else get_item(items, iid)
        if not is_new and not it: abort(404)
        if request.method == "POST":
            nom = (request.form.get("nom", "") or "").strip()
            cat = request.form.get("cat", "fixe")
            desc = (request.form.get("desc", "") or "").strip()
            if cat not in dict(APPAREILS_CATS): cat = "fixe"
            if not nom:
                return redirect(url_for("appareil_edit", iid=iid))
            if is_new:
                nid = _slug(nom)
                base = nid; k = 2
                while get_item(items, nid): nid = "%s_%d" % (base, k); k += 1
                it = {"id": nid, "cat": cat, "svg": "bracket", "nom": nom, "desc": desc}
                items.append(it); iid = nid
            else:
                it["nom"] = nom; it["cat"] = cat; it["desc"] = desc
            save_items(items)
            f = request.files.get("photo")
            purl = (request.form.get("photo_url", "") or "").strip()
            src = None
            if f and f.filename:
                src = f.stream
            elif purl:
                try:
                    import urllib.request as _ur, io as _io, ssl as _ssl
                    try:
                        import certifi as _cf; _ctx = _ssl.create_default_context(cafile=_cf.where())
                    except Exception:
                        _ctx = _ssl._create_unverified_context()
                    _rq = _ur.Request(purl, headers={"User-Agent": "Mozilla/5.0 BilanODF"})
                    _raw = _ur.urlopen(_rq, timeout=20, context=_ctx).read(30 * 1024 * 1024)
                    src = _io.BytesIO(_raw)
                except Exception:
                    src = None
            if src is not None:
                try:
                    from PIL import Image as _I, ImageOps as _IO
                    os.makedirs(IMGDIR, exist_ok=True)
                    im = _IO.exif_transpose(_I.open(src)).convert("RGB")
                    _o = im.copy(); _o.thumbnail((2200, 2200), _I.LANCZOS)
                    _o.save(os.path.join(IMGDIR, iid + "_orig.jpg"), "JPEG", quality=90)
                    _d = im.copy(); _d.thumbnail((1400, 1400), _I.LANCZOS)
                    _d.save(os.path.join(IMGDIR, iid + ".jpg"), "JPEG", quality=88)
                    it["crop"] = None; save_items(items)
                except Exception:
                    pass
            return redirect(url_for("appareils_home"))
        cur_cat = (it or {}).get("cat", "fixe")
        opts = "".join('<option value="%s"%s>%s</option>'
                       % (c, " selected" if c == cur_cat else "", esc(l)) for c, l in APPAREILS_CATS)
        title = "Nouvel appareil" if is_new else "Modifier l’appareil"
        photo_block = ""
        if not is_new and has_photo(iid):
            photo_block = ('<div style="margin:6px 0 10px"><img src="%s" style="max-height:160px;border-radius:12px;border:1px solid var(--line)"></div>'
                           '<div style="margin:2px 0 10px"><a class="btn sec" href="%s">\u2702 Recadrer la photo</a></div>'
                           % (url_for("appareil_img", iid=iid) + "?t=%d" % int(os.path.getmtime(os.path.join(IMGDIR, iid + ".jpg"))), url_for("appareil_crop", iid=iid)))
        del_block = ""
        if not is_new:
            del_block = ('<form method=post action="%s" style="margin-top:20px" onsubmit="return confirm(\'Supprimer cet appareil ?\')">'
                         '<button class="btn sec" type=submit style="color:var(--bad);border-color:var(--bad)">Supprimer</button></form>'
                         % url_for("appareil_delete", iid=iid))
        body = ('<p class=muted><a href="%s">← Appareils</a></p><h1>%s</h1>'
                '<form method=post enctype=multipart/form-data class=card style="max-width:620px">'
                '<label>Nom de l’appareil</label><input name=nom value="%s" required>'
                '<div style="margin-top:12px"><label>Catégorie</label><select name=cat>%s</select></div>'
                                '<div style="margin-top:12px"><label>Photo (remplace l’illustration)</label>%s'
                '<input type=file name=photo accept="image/*">'
                '<div style="margin-top:8px"><input name=photo_url placeholder="\u2026 ou coller l\u2019URL d\u2019une image (libre de droit)" style="width:100%%;padding:8px 10px;border:1px solid var(--line);border-radius:8px"></div></div>'
                '<div style="margin-top:16px"><button class=btn type=submit>Enregistrer</button> '
                '<a class="btn sec" href="%s">Annuler</a></div></form>%s') % (
                    url_for("appareils_home"), esc(title),
                    escat((it or {}).get("nom", "")), opts,
                    photo_block, url_for("appareils_home"), del_block)
        return page(body)

    def _ap_src_path(iid):
        o = os.path.join(IMGDIR, iid + "_orig.jpg")
        p = os.path.join(IMGDIR, iid + ".jpg")
        return o if os.path.exists(o) else (p if os.path.exists(p) else None)

    def _ap_oriented(iid, rot, fine, mirror, flip, maxpx):
        from PIL import Image as _I, ImageOps as _IO
        sp = _ap_src_path(iid)
        if not sp: return None
        im = _IO.exif_transpose(_I.open(sp)).convert("RGB")
        ang = (int(rot) % 360) + float(fine)
        if ang: im = im.rotate(ang, expand=True, resample=_I.BICUBIC)
        if mirror: im = _IO.mirror(im)
        if flip: im = _IO.flip(im)
        if maxpx: im.thumbnail((maxpx, maxpx), _I.LANCZOS)
        return im

    @app.route("/appareils/<iid>/oriented")
    def appareil_oriented(iid):
        if not logged(): return redirect(url_for("login"))
        import io as _io
        try: rot = int(request.args.get("rot", 0)) % 360
        except Exception: rot = 0
        try: fine = max(-45.0, min(45.0, float(request.args.get("fine", 0))))
        except Exception: fine = 0.0
        im = _ap_oriented(iid, rot, fine, request.args.get("mirror") == "1", request.args.get("flip") == "1", 1200)
        if im is None: abort(404)
        buf = _io.BytesIO(); im.save(buf, "JPEG", quality=88); buf.seek(0)
        return send_file(buf, mimetype="image/jpeg")

    @app.route("/appareils/<iid>/crop", methods=["GET"])
    def appareil_crop(iid):
        if not logged(): return redirect(url_for("login"))
        items = load_items(); it = get_item(items, iid)
        if not it or not _ap_src_path(iid): abort(404)
        import app as _app
        try:
            _im0 = _ap_oriented(iid, 0, 0, False, False, None)
            _W, _H = _im0.size; _ar = 4.0 / 3.0
            if _W / float(_H) > _ar:
                _h = 1.0; _w = (_H * _ar) / float(_W)
            else:
                _w = 1.0; _h = (_W / _ar) / float(_H)
            df = {"rot": 0, "fine": 0, "mirror": False, "flip": False,
                  "x": (1.0 - _w) / 2.0, "y": (1.0 - _h) / 2.0, "w": _w, "h": _h}
        except Exception:
            df = {"rot": 0, "fine": 0, "mirror": False, "flip": False, "x": 0.0, "y": 0.0, "w": 1.0, "h": 1.0}
        init = it.get("crop") or df
        html = (_app.CROP_HTML
                .replace("__LABEL__", esc(it.get("nom", "")))
                .replace("__IMGBASE__", url_for("appareil_oriented", iid=iid))
                .replace("__SAVEURL__", url_for("appareil_crop_save", iid=iid))
                .replace("__BACK__", url_for("appareils_home"))
                .replace("__ASPECT__", "1.333333")
                .replace("__INIT__", json.dumps({k: init.get(k, df[k]) for k in ("rot", "fine", "mirror", "flip", "x", "y", "w", "h")}))
                .replace("__DEF__", json.dumps(df)))
        return html

    @app.route("/appareils/<iid>/crop", methods=["POST"])
    def appareil_crop_save(iid):
        from flask import jsonify
        if not logged(): return jsonify(ok=False), 403
        items = load_items(); it = get_item(items, iid)
        if not it: return jsonify(ok=False), 404
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
             "w": max(0.02, cl(d.get("w", 1))), "h": max(0.02, cl(d.get("h", 1)))}
        from PIL import Image as _I
        im = _ap_oriented(iid, c["rot"], c["fine"], c["mirror"], c["flip"], None)
        if im is None: return jsonify(ok=False), 404
        W, H = im.size
        x0 = int(c["x"] * W); y0 = int(c["y"] * H)
        x1 = int((c["x"] + c["w"]) * W); y1 = int((c["y"] + c["h"]) * H)
        x0 = max(0, min(W - 1, x0)); y0 = max(0, min(H - 1, y0))
        x1 = max(x0 + 1, min(W, x1)); y1 = max(y0 + 1, min(H, y1))
        im = im.crop((x0, y0, x1, y1))
        im.thumbnail((1400, 1400), _I.LANCZOS)
        os.makedirs(IMGDIR, exist_ok=True)
        im.save(os.path.join(IMGDIR, iid + ".jpg"), "JPEG", quality=90)
        it["crop"] = c; save_items(items)
        return jsonify(ok=True, redirect=url_for("appareils_home"))

    @app.route("/appareils/<iid>/delete", methods=["POST"])
    def appareil_delete(iid):
        if not logged(): return redirect(url_for("login"))
        items = load_items()
        items = [x for x in items if x.get("id") != iid]
        save_items(items)
        try: os.remove(os.path.join(IMGDIR, iid + ".jpg"))
        except Exception: pass
        return redirect(url_for("appareils_home"))


SHOW_TPL = """<!doctype html><html lang=fr><head><meta charset=utf-8>
<meta name=viewport content="width=device-width,initial-scale=1"><title>__NOM__</title>
<script>(function(){try{var t=localStorage.getItem('bilan-theme')||'light';document.documentElement.setAttribute('data-theme',t);}catch(e){}})();</script>
<style>
:root{--bg:#f4f6f9;--card:#fff;--ink:#0c1526;--mut:#6b7891;--line:#e7ebf1;--acc:#2f6bff}
html[data-theme="dark"]{--bg:#0a0d13;--card:#12161f;--ink:#eef2f8;--mut:#8a96ab;--line:#1d2531;--acc:#4d8bff}
*{box-sizing:border-box}
html,body{margin:0;height:100%;background:var(--bg);color:var(--ink);
  font-family:-apple-system,"SF Pro Text","Helvetica Neue",system-ui,sans-serif;letter-spacing:-.01em}
.bar{position:fixed;top:0;left:0;right:0;height:58px;display:flex;align-items:center;justify-content:space-between;padding:0 22px;z-index:5}
.bar a{display:inline-flex;align-items:center;gap:8px;color:var(--mut);text-decoration:none;font-size:14px;font-weight:600;
  border:1px solid var(--line);border-radius:10px;padding:8px 14px;background:var(--card)}
.wrap{min-height:100%;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:74px 40px 48px;text-align:center}
.vis{max-width:min(880px,80vw)}
.apimg.big{max-height:64vh;border-radius:18px;box-shadow:0 18px 60px rgba(0,0,0,.28);display:block}
.apsvg.big{color:var(--mut);width:340px;max-width:70vw}
.apsvg.big svg{width:100%;height:auto}
h1{font-size:clamp(26px,4.4vh,44px);font-weight:700;letter-spacing:-.02em;margin:30px 0 14px}
.desc{font-size:clamp(16px,2.5vh,24px);color:var(--mut);line-height:1.55;max-width:60ch}
</style></head><body>
<div class=bar><a href="__BACK__">← Retour</a><span></span></div>
<div class=wrap><div class=vis>__VIS__</div><h1>__NOM__</h1></div>
</body></html>"""
