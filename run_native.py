# -*- coding: utf-8 -*-
"""
Lance Bilan ODF dans une VRAIE fenêtre native macOS, sans Terminal et sans
dépendre d'un navigateur type Chrome.

Stratégie :
  1) Fenêtre WebKit native via pyobjc (WKWebView = le moteur de Safari, fourni
     par macOS). Indépendant de tout navigateur -> base idéale pour une future .app.
  2) Repli sur pywebview seulement s'il est complet (souvent cassé sur Python 3.14).
  3) Dernier recours : navigateur par défaut.
Fermer la fenêtre quitte l'application.
"""
import os, sys, time, threading

# --- Mode « rendu STL » (sous-processus isolé, y compris en exécutable gelé) ---
# app.py relance l'exécutable avec  render_stl <dossier>  : on dispatche AVANT
# de démarrer le serveur web pour que le rendu VTK tourne dans un process séparé.
if len(sys.argv) >= 3 and sys.argv[1] == "render_stl":
    import pipeline as _P
    _P.render_stl(sys.argv[2])
    sys.exit(0)

from app import app as flask_app

import socket

def _free_port():
    """Choisit un port local libre en évitant 5000 (occupé par le « Récepteur AirPlay »
    de macOS, qui renvoyait des erreurs 403 et empêchait l'app de démarrer)."""
    for p in (5137, 5173, 5321, 5480, 8137, 8765, 8099):
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            s.bind(("127.0.0.1", p)); s.close(); return p
        except OSError:
            try: s.close()
            except Exception: pass
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.bind(("127.0.0.1", 0)); p = s.getsockname()[1]; s.close(); return p

PORT = _free_port()
URL = "http://127.0.0.1:%d" % PORT

def serve():
    flask_app.run(host="127.0.0.1", port=PORT, threaded=True, use_reloader=False)

def run_wkwebview():
    """Fenêtre native via WebKit du système (pyobjc). Lève ImportError si pyobjc absent."""
    from AppKit import (NSApplication, NSWindow, NSObject, NSBackingStoreBuffered,
                        NSApplicationActivationPolicyRegular, NSOpenPanel,
                        NSRunningApplication, NSApplicationActivateIgnoringOtherApps,
                        NSFloatingWindowLevel, NSNormalWindowLevel, NSView)
    from WebKit import WKWebView, WKWebViewConfiguration
    from Foundation import NSURL, NSURLRequest, NSMakeRect
    import objc

    STYLE = (1 << 0) | (1 << 1) | (1 << 2) | (1 << 3)   # Titled|Closable|Miniaturizable|Resizable

    class AppDelegate(NSObject):
        def applicationShouldTerminateAfterLastWindowClosed_(self, sender):
            return True   # fermer la fenêtre quitte l'app

    INBOX = os.path.expanduser("~/BilanODF_Data/_radios_inbox")
    _browser_windows = []
    _docstate = {}
    class _MsgHandler(NSObject):
        def userContentController_didReceiveScriptMessage_(self, ucc, message):
            try: d = str(message.body())
            except Exception: return
            # Ouverture d'une plateforme AVEC recherche d'un patient : "open_<which>:<nom>"
            for _pfx in ("open_doctolib:", "open_xero:", "open_webceph:"):
                if d.startswith(_pfx):
                    _which = _pfx[5:-1]           # doctolib / xero / webceph
                    _nom = d[len(_pfx):].strip()
                    if _nom:
                        try:
                            import subprocess as _sp
                            _sp.run(["pbcopy"], input=_fmt_name(_which, _nom).encode("utf-8"))   # prêt pour Cmd+V (format plateforme)
                        except Exception: pass
                        _docstate["search_" + _which] = _nom
                    try:
                        _t = _docstate.get("toggle")
                        if _t: _t(True, _which)
                    except Exception: pass
                    return
            if d in ("open_doctolib", "close_doctolib", "open_xero", "close_xero", "open_webceph", "close_webceph"):
                try:
                    _t = _docstate.get("toggle")
                    if _t: _t(d.startswith("open_"), "xero" if "xero" in d else ("webceph" if "webceph" in d else "doctolib"))
                except Exception: pass
                return
            if d.startswith("doctopat:"):
                try:
                    _dd = os.path.dirname(INBOX)
                    os.makedirs(_dd, exist_ok=True)
                    open(os.path.join(_dd, "_doctolib_inbox.json"), "w", encoding="utf-8").write(d[9:])
                    _t = _docstate.get("toggle")
                    if _t: _t(False, _docstate.get("shown") or "doctolib")
                    _mw = _docstate.get("mainweb"); _bu = _docstate.get("baseurl")
                    if _mw is not None and _bu:
                        from Foundation import NSURL as _NU, NSURLRequest as _NR
                        _mw.loadRequest_(_NR.requestWithURL_(_NU.URLWithString_(_bu.rstrip("/") + "/from_doctolib")))
                except Exception as _e: print("doctopat", _e)
                return
            if d.startswith("steiner:"):
                try:
                    import time as _tst
                    _qd = os.path.join(os.path.dirname(INBOX), "_steiner_inbox")
                    os.makedirs(_qd, exist_ok=True)
                    open(os.path.join(_qd, "st_%d.json" % int(_tst.time() * 1000)), "w", encoding="utf-8").write(d[8:])
                except Exception as _e: print("steiner", _e)
                return
            if d.startswith("snap"):
                try: _do_snap(message, d)
                except Exception as _e: print("snap", _e)
                return
            if not d.startswith("data:"): return
            import base64 as _b64, time as _t
            try: raw = _b64.b64decode(d.split(",", 1)[1])
            except Exception: return
            try: os.makedirs(INBOX, exist_ok=True)
            except Exception: pass
            try: open(os.path.join(INBOX, "radio_%d.jpg" % int(_t.time() * 1000)), "wb").write(raw)
            except Exception: pass
    _msgh = _MsgHandler.alloc().init(); _browser_windows.append(_msgh)
    def _do_snap(message, d):
        rect = None
        if d.startswith("snap:"):
            try: rect = [float(x) for x in d[5:].split(",")]
            except Exception: rect = None
        try: wv = message.webView()
        except Exception: wv = None
        if wv is None: return
        from WebKit import WKSnapshotConfiguration
        cfg = WKSnapshotConfiguration.alloc().init()
        if rect and len(rect) == 4:
            try: cfg.setRect_(NSMakeRect(rect[0], rect[1], rect[2], rect[3]))
            except Exception: pass
        def _cb(img, err):
            try:
                if img is None: return
                from AppKit import NSBitmapImageRep
                rep2 = NSBitmapImageRep.imageRepWithData_(img.TIFFRepresentation())
                data = rep2.representationUsingType_properties_(3, {})
                try: os.makedirs(INBOX, exist_ok=True)
                except Exception: pass
                import time as _t2
                data.writeToFile_atomically_(os.path.join(INBOX, "webceph_%d.jpg" % int(_t2.time() * 1000)), True)
            except Exception as _e:
                print("snap save", _e)
        wv.takeSnapshotWithConfiguration_completionHandler_(cfg, _cb)
    try:
        import app as _appmod
        _CAPTOK = getattr(_appmod, "RADIO_CAP_TOKEN", "")
    except Exception:
        _CAPTOK = ""
    _CAP_JS = (r'''(function(){
if(window.__bilanCapInit)return;window.__bilanCapInit=1;
var LBL="Recuperer une image (clique dessus)";var picking=false;
function send(m){try{window.webkit.messageHandlers.bilan.postMessage(m);}catch(e){}}
function done(b){if(b){b.textContent="Recuperee ! (Radios & traces)";setTimeout(function(){b.textContent=LBL;},2400);}}
function cap(el,b){var tag=el.tagName.toLowerCase();
if(tag==="svg"){try{var xml=new XMLSerializer().serializeToString(el);var W=el.clientWidth||600,H=el.clientHeight||600;var im=new Image();im.onload=function(){var cv=document.createElement("canvas");cv.width=W;cv.height=H;var cx=cv.getContext("2d");cx.fillStyle="#fff";cx.fillRect(0,0,W,H);cx.drawImage(im,0,0,W,H);send(cv.toDataURL("image/jpeg",0.92));done(b);};im.onerror=function(){if(b)b.textContent="SVG non capturable";};im.src="data:image/svg+xml;base64,"+btoa(unescape(encodeURIComponent(xml)));return;}catch(e){if(b)b.textContent="SVG non capturable";return;}}
try{var cv=document.createElement("canvas");cv.width=el.naturalWidth||el.width||el.clientWidth;cv.height=el.naturalHeight||el.height||el.clientHeight;cv.getContext("2d").drawImage(el,0,0);send(cv.toDataURL("image/jpeg",0.92));done(b);return;}catch(e){}
if(tag==="img"&&el.src){fetch(el.src).then(function(r){return r.blob();}).then(function(bl){var fr=new FileReader();fr.onload=function(){send(fr.result);done(b);};fr.readAsDataURL(bl);}).catch(function(){if(b)b.textContent="Image protegee";});return;}
if(b)b.textContent="Impossible a capturer";}
function pickAt(x,y){var st=document.elementsFromPoint(x,y)||[];for(var i=0;i<st.length;i++){var t=st[i].tagName?st[i].tagName.toLowerCase():"";if(t==="img"||t==="canvas")return st[i];}for(var j=0;j<st.length;j++){var e=st[j];if(e.tagName&&e.tagName.toLowerCase()==="svg")return e;if(e.closest){var sv=e.closest("svg");if(sv)return sv;}}return null;}
function onDown(ev){if(!picking)return;var el=pickAt(ev.clientX,ev.clientY);if(el){ev.preventDefault();ev.stopPropagation();picking=false;document.body.style.cursor="";cap(el,document.getElementById("__bilanCapBtn"));}}
function mk(){if(document.getElementById("__bilanCapBtn"))return;var b=document.createElement("button");b.id="__bilanCapBtn";b.textContent=LBL;b.style.cssText="position:fixed;z-index:2147483647;right:16px;bottom:16px;padding:11px 15px;background:#2563eb;color:#fff;border:none;border-radius:9px;font-size:14px;font-family:sans-serif;cursor:pointer;box-shadow:0 3px 12px rgba(0,0,0,.35)";b.onclick=function(ev){ev.stopPropagation();if(picking){picking=false;document.body.style.cursor="";b.textContent=LBL;return;}picking=true;document.body.style.cursor="crosshair";b.textContent="Clique l'image a recuperer (Echap = annuler)";};document.body.appendChild(b);}
document.addEventListener("pointerdown",onDown,true);
document.addEventListener("keydown",function(e){if(e.key==="Escape"&&picking){picking=false;document.body.style.cursor="";var b=document.getElementById("__bilanCapBtn");if(b)b.textContent=LBL;}});
mk();setInterval(mk,1500);
})();''').replace("__TOK__", _CAPTOK).replace("__BASE__", URL)
    _CAP_JS_SNAP = (r'''(function(){
if(window.__bilanSnapInit)return;window.__bilanSnapInit=1;
var LBL="Capturer une zone (radio/trace)";var active=false,sx=0,sy=0,box=null;
function bt(){return document.getElementById("__bilanSnapBtn");}
function reset(){active=false;if(box){box.remove();box=null;}document.body.style.cursor="";var b=bt();if(b)b.textContent=LBL;}
function onDown(e){if(!active)return;if(e.target&&e.target.id==="__bilanSnapBtn")return;e.preventDefault();e.stopPropagation();sx=e.clientX;sy=e.clientY;box=document.createElement("div");box.style.cssText="position:fixed;border:2px dashed #2563eb;background:rgba(37,99,235,.12);z-index:2147483646;pointer-events:none;left:"+sx+"px;top:"+sy+"px;width:0;height:0";document.body.appendChild(box);}
function onMove(e){if(!active||!box)return;var x=Math.min(e.clientX,sx),y=Math.min(e.clientY,sy),w=Math.abs(e.clientX-sx),h=Math.abs(e.clientY-sy);box.style.left=x+"px";box.style.top=y+"px";box.style.width=w+"px";box.style.height=h+"px";}
function onUp(e){if(!active||!box)return;e.preventDefault();e.stopPropagation();var x=Math.min(e.clientX,sx),y=Math.min(e.clientY,sy),w=Math.abs(e.clientX-sx),h=Math.abs(e.clientY-sy);box.remove();box=null;active=false;document.body.style.cursor="";var b=bt();if(w>15&&h>15){try{window.webkit.messageHandlers.bilan.postMessage("snap:"+Math.round(x)+","+Math.round(y)+","+Math.round(w)+","+Math.round(h));}catch(err){}if(b)b.textContent="Zone capturee ! (Radios & traces)";setTimeout(function(){if(b)b.textContent=LBL;},2200);}else{if(b)b.textContent=LBL;}}
function mk(){if(document.getElementById("__bilanSnapBtn"))return;var b=document.createElement("button");b.id="__bilanSnapBtn";b.textContent=LBL;b.style.cssText="position:fixed;z-index:2147483647;right:16px;bottom:16px;padding:11px 15px;background:#2563eb;color:#fff;border:none;border-radius:9px;font-size:14px;font-family:sans-serif;cursor:pointer;box-shadow:0 3px 12px rgba(0,0,0,.35)";b.onclick=function(ev){ev.stopPropagation();active=true;document.body.style.cursor="crosshair";b.textContent="Trace un rectangle sur la zone (Echap=annuler)";};document.body.appendChild(b);}
document.addEventListener("pointerdown",onDown,true);
document.addEventListener("pointermove",onMove,true);
document.addEventListener("pointerup",onUp,true);
document.addEventListener("keydown",function(e){if(e.key==="Escape"&&active)reset();});
mk();setInterval(mk,1500);
})();''')
    _RET_TMPL = ('(function(){if(window.__bRet_WHICH)return;window.__bRet_WHICH=1;'
                 'function mk(){if(document.getElementById("__bRetBtn_WHICH"))return;'
                 'var b=document.createElement("button");b.id="__bRetBtn_WHICH";b.textContent="← Retour a Bilan ODF";'
                 'b.style.cssText="position:fixed;z-index:2147483647;left:16px;bottom:16px;padding:11px 16px;background:#16283d;color:#fff;border:none;border-radius:10px;font-size:14px;font-family:-apple-system,sans-serif;font-weight:600;cursor:pointer;box-shadow:0 4px 14px rgba(0,0,0,.4)";'
                 'b.onclick=function(){try{window.webkit.messageHandlers.bilan.postMessage("close_WHICH");}catch(e){}};'
                 'document.body.appendChild(b);}'
                 'mk();setInterval(mk,1500);})();')
    _FILL_TMPL = ('(function(){if(window.__bFill_WHICH)return;'
                  'var ID=__ID__,PW=__PW__;if(!(ID||PW))return;'
                  'function sv(el,v){try{var d=Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype,"value").set;d.call(el,v);}catch(e){el.value=v;}'
                  'el.dispatchEvent(new Event("input",{bubbles:true}));el.dispatchEvent(new Event("change",{bubbles:true}));}'
                  'function fill(){window.__ft_WHICH=(window.__ft_WHICH||0)+1;'
                  'var u=document.querySelector("input[type=email],input[name=email],input[name=username],input[name=login],input[name*=user],input[autocomplete=username],input#username");'
                  'if(!u){ if(window.__ft_WHICH>40){window.__bFill_WHICH=1;if(window.__fiv_WHICH)clearInterval(window.__fiv_WHICH);} return; }'
                  'var p=document.querySelector("input[autocomplete=current-password],input[type=password]");'
                  'if(ID&&!u.value){sv(u,ID);}'
                  'if(PW&&p&&!p.value){sv(p,PW);}'
                  'if((PW&&p&&p.value)||(!PW&&ID&&u&&u.value)||window.__ft_WHICH>50){window.__bFill_WHICH=1;if(window.__fiv_WHICH)clearInterval(window.__fiv_WHICH);}}'
                  'window.__fiv_WHICH=setInterval(fill,800);fill();})();')
    def _ext_creds(which):
        try:
            import app as _am
            c = (_am.doctolib_creds_get() if which == "doctolib" else (_am.webceph_creds_get() if which == "webceph" else _am.xero_creds_get())) or {}
            return (c.get("id", "") or "", c.get("pw", "") or "")
        except Exception:
            return ("", "")
    def _ext_js(which):
        import json as _j
        i, p = _ext_creds(which)
        js = _RET_TMPL.replace("WHICH", which)
        if i or p:
            js += _FILL_TMPL.replace("WHICH", which).replace("__ID__", _j.dumps(i)).replace("__PW__", _j.dumps(p))
        return js

    def _name_parts(nom):
        """(nom_de_famille[list], prenom[list]) — le nom = les jetons MAJUSCULES de tete."""
        toks = (nom or "").strip().split()
        surn = []; i = 0
        while i < len(toks) and toks[i] == toks[i].upper() and any(c.isalpha() for c in toks[i]):
            surn.append(toks[i]); i += 1
        if not surn and toks:            # repli : 1er jeton = nom
            surn = [toks[0]]; i = 1
        return surn, toks[i:]

    def _fmt_name(which, nom):
        """Texte a taper dans la recherche selon la plateforme.
        XERO : 'NOM, Prenom' (avec virgule) ; WebCeph : 'NOM Prenom' (sans virgule) ; Doctolib : nom complet."""
        surn, first = _name_parts(nom)
        nom = (nom or "").strip()
        if which == "xero":
            if surn and first:
                return " ".join(surn) + ", " + " ".join(first)
            return nom
        if which == "webceph":
            if surn and first:
                return " ".join(surn) + " " + " ".join(first)
            return nom
        return nom

    def _search_js(which, nom):
        """Ouvre la recherche de la plateforme et cherche le patient :
        - WebCeph : va d'abord sur 'Liste des Patients' (meme si une fiche est ouverte), tape, clique Rechercher.
        - Doctolib : vide le champ (efface un patient deja cherche) puis tape.
        - XERO : tape 'NOM, Prenom'.
        - Si un SEUL resultat -> ouvre sa fiche automatiquement. Le nom reste dans le presse-papiers (Cmd+V)."""
        import json as _j
        surn_list, _first = _name_parts(nom)
        typed = _fmt_name(which, nom)
        disp = (nom or "").strip()
        surn = " ".join(surn_list)
        clear_first = "1" if which in ("doctolib", "xero") else "0"
        tmpl = r'''(function(){
var WHICH=__WHICH__, TYPED=__TYPED__, DISP=__DISP__, SURN=__SURN__, CLEARFIRST=(__CLEAR__===1);
function toast(t){try{var o=document.getElementById("__bilanSrchTst");if(o)o.remove();
var e=document.createElement("div");e.id="__bilanSrchTst";e.textContent=t;
e.style.cssText="position:fixed;left:50%;bottom:22px;transform:translateX(-50%);z-index:2147483647;background:#0b3550;color:#fff;padding:10px 16px;border-radius:10px;font:14px/1.3 -apple-system,sans-serif;box-shadow:0 6px 22px rgba(0,0,0,.35);max-width:80%";
document.body.appendChild(e);setTimeout(function(){e.style.transition="opacity .5s";e.style.opacity="0";setTimeout(function(){e.remove();},600);},4200);}catch(x){}}
function vis(el){return el&&el.offsetParent!==null;}
function clickText(txts){var els=document.querySelectorAll('a,button,[role=tab],[role=button],li,span,div,input[type=button],input[type=submit],input[type=reset]');for(var i=0;i<els.length;i++){var el=els[i];if(!vis(el))continue;var t=((el.tagName==="INPUT")?(el.value||el.getAttribute("value")||""):(el.textContent||"")).replace(/\s+/g,' ').trim().toLowerCase();if(!t||t.length>34)continue;for(var j=0;j<txts.length;j++){if(t===txts[j]||t.indexOf(txts[j])===0){try{el.click();return true;}catch(e){}}}}return false;}
function labelEl(txts){var els=document.querySelectorAll('label,td,th,span,div,b,strong,p');for(var i=0;i<els.length;i++){var el=els[i];if(!vis(el))continue;var t=(el.textContent||"").replace(/\s+/g,' ').replace(/[:*]/g,'').trim().toLowerCase();if(!t||t.length>22)continue;for(var j=0;j<txts.length;j++){if(t===txts[j])return el;}}return null;}
function inputRightOf(lab){if(!lab)return null;var lr=lab.getBoundingClientRect();var best=null,bd=1e9;var ins=document.querySelectorAll('input');for(var i=0;i<ins.length;i++){var el=ins[i];if(!vis(el))continue;var ty=(el.getAttribute("type")||"text").toLowerCase();if(ty!=="text"&&ty!=="search")continue;var r=el.getBoundingClientRect();if(r.width<40)continue;var dyc=Math.abs((r.top+r.height/2)-(lr.top+lr.height/2));if(dyc>18)continue;if(r.left<lr.left-2)continue;var dx=r.left-lr.right;if(dx<-40)continue;var d=Math.abs(dx)+dyc*3;if(d<bd){bd=d;best=el;}}return best;}
function box(){
if(WHICH==="xero"){var xe=inputRightOf(labelEl(["nom du patient","nom patient","nom"]));if(vis(xe))return xe;}
var s=['input[placeholder*="nom du patient" i]','input[aria-label*="nom du patient" i]','input[placeholder*="echerch" i]','input[placeholder*="atient" i]','input[type=search]','input[aria-label*="echerch" i]','input[aria-label*="earch" i]','[role=searchbox]','input[name*="search" i]','input[name*="recherch" i]','input[name*="query" i]'];
if(WHICH!=="webceph"&&WHICH!=="xero"){s.push('input[placeholder*="earch" i]');s.push('input[type=text]');}
for(var i=0;i<s.length;i++){var el=document.querySelector(s[i]);if(vis(el))return el;}return null;}
function nativeSet(el,v){try{el.focus();var pr=Object.getPrototypeOf(el);var de=Object.getOwnPropertyDescriptor(pr,"value");if(de&&de.set)de.set.call(el,v);else el.value=v;["input","change","keyup"].forEach(function(t){el.dispatchEvent(new Event(t,{bubbles:true}));});}catch(e){}}
function enter(el){try{["keydown","keypress","keyup"].forEach(function(t){el.dispatchEvent(new KeyboardEvent(t,{bubbles:true,key:"Enter",code:"Enter",keyCode:13,which:13}));});var f=el.form||el.closest("form");if(f){try{f.requestSubmit?f.requestSubmit():f.submit();}catch(e){}}}catch(e){}}
function esc(s){return s.replace(/[.*+?^${}()|[\]\\]/g,"\\$&");}
var RE=SURN?new RegExp("(^|\\s)"+esc(SURN),"i"):null;
function clickable(el){var n=el;for(var d=0;d<6&&n;d++){var tg=(n.tagName||"").toLowerCase();if(tg==="a"||tg==="tr"||tg==="li"||tg==="button")return n;try{if(n.getAttribute&&(n.getAttribute("role")==="button"||n.hasAttribute("onclick")))return n;}catch(e){}try{if(getComputedStyle(n).cursor==="pointer")return n;}catch(e){}n=n.parentElement;}return el;}
function rows(){if(!RE)return [];var els=document.querySelectorAll('a,tr,li,td,div,span,button'),map={};for(var i=0;i<els.length;i++){var el=els[i];if(!vis(el))continue;var t=(el.textContent||"").replace(/\s+/g," ").trim();if(!t||t.length>48)continue;if(!RE.test(t))continue;var r=el.getBoundingClientRect();if(r.width<20||r.height<8)continue;var key=Math.round(r.top/8);var cur=map[key];if(!cur||t.length<(cur.textContent||"").replace(/\s+/g," ").trim().length)map[key]=el;}var out=[];for(var k in map)out.push(map[k]);return out;}
function openIfSingle(){try{var rs=rows();if(rs.length===1){clickable(rs[0]).click();toast("Fiche ouverte : « "+DISP+" »");return true;}}catch(e){}return false;}
function after(el){enter(el);if(WHICH==="webceph"||WHICH==="xero"){setTimeout(function(){clickText(["rechercher"]);},250);}toast("🔎 Recherche de « "+DISP+" »");setTimeout(openIfSingle,1700);}
function typeIn(el){if(CLEARFIRST){nativeSet(el,"");setTimeout(function(){nativeSet(el,TYPED);after(el);},320);}else{nativeSet(el,TYPED);after(el);}}
// XERO / Doctolib : cliquer "Effacer" d'abord (au cas ou un patient etait deja selectionne), puis reprendre la barre.
function preclear(cb){if(WHICH==="xero"||WHICH==="doctolib"){clickText(["effacer","effacer la recherche","vider la recherche","vider","reinitialiser","réinitialiser","clear"]);setTimeout(cb,400);}else{cb();}}
var navied=false,n=0,iv=setInterval(function(){n++;
if(WHICH==="webceph"&&!navied){navied=true;clickText(["liste des patients","liste des patient"]);}
var el=box();
if(el){clearInterval(iv);preclear(function(){var e2=box()||el;typeIn(e2);});return;}
if(n>60){clearInterval(iv);toast("Nom copié : « "+DISP+" » — collez-le (Cmd+V) dans la recherche.");}
},400);
})();'''
        return (tmpl.replace("__WHICH__", _j.dumps(which)).replace("__TYPED__", _j.dumps(typed))
                .replace("__DISP__", _j.dumps(disp)).replace("__SURN__", _j.dumps(surn))
                .replace("__CLEAR__", clear_first))

    _DOCTO_JS = (r'''(function(){
if(window.__bilanDocInit)return;window.__bilanDocInit=1;
function toISO(d){var m=d.match(/(\d{2})\/(\d{2})\/(\d{4})/);return m?m[3]+"-"+m[2]+"-"+m[1]:"";}
function scrape(){
var t=document.body.innerText||"";
var lines=t.split("\n").map(function(s){return s.replace(/\s+/g," ").trim();}).filter(Boolean);
var o={nom:"",prenom:"",dob:"",sexe:"",tel:"",email:""};
for(var i=0;i<lines.length;i++){var mb=lines[i].match(/^([HFhf])\s*[,\.]\s*(\d{2}\/\d{2}\/\d{4})/);if(mb){o.dob=toISO(mb[2]);o.sexe=(mb[1].toUpperCase()==="F")?"féminin":"masculin";if(i>=1)o.prenom=lines[i-1];if(i>=2)o.nom=lines[i-2];break;}}
if(o.nom&&/^(m\.|mme|mlle|monsieur|madame|mademoiselle|h|f)$/i.test(o.nom))o.nom="";
if(o.prenom&&/^(m\.|mme|mlle|monsieur|madame|mademoiselle)$/i.test(o.prenom))o.prenom="";
if(!o.dob){var md=t.match(/(\d{2}\/\d{2}\/\d{4})/);if(md)o.dob=toISO(md[1]);}
if(!o.sexe){var lt=t.toLowerCase();if(/née\s+le/.test(lt)||/\bmadame\b|\bmme\b/.test(lt))o.sexe="féminin";else if(/né\s+le/.test(lt)||/\bmonsieur\b/.test(lt))o.sexe="masculin";}
if(!o.nom){var h=document.querySelector("h1");if(h){var hn=(h.innerText||"").replace(/\s+/g," ").trim().split(" ").filter(function(w){return w&&!/^(m\.|mr|mme|mlle|dr|pr|monsieur|madame|mademoiselle)$/i.test(w);});if(hn.length>=2){o.nom=hn[hn.length-1];if(!o.prenom)o.prenom=hn.slice(0,hn.length-1).join(" ");}else if(hn.length===1){o.nom=hn[0];}}}
var mem=t.match(/e[\s\-]?mail\s*:?\s*([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})/i);if(mem)o.email=mem[1];
var at=document.querySelector('a[href^="tel:"]');if(at){o.tel=(at.getAttribute("href")||"").replace(/^tel:/i,"").replace(/[^\d+]/g,"");}
if(!o.tel){var mp=t.match(/(?:\+33|0)[\s.\-]?[1-9](?:[\s.\-]?\d{2}){4}/);if(mp)o.tel=mp[0].replace(/[\s.\-]/g,"");}
return o;}
function mk(){if(document.getElementById("__bilanDocBtn"))return;var b=document.createElement("button");b.id="__bilanDocBtn";b.textContent="➕ Créer une fiche Bilan ODF";b.style.cssText="position:fixed;z-index:2147483647;right:16px;bottom:16px;padding:11px 15px;background:#0b6bcb;color:#fff;border:none;border-radius:9px;font-size:14px;font-family:sans-serif;cursor:pointer;box-shadow:0 3px 12px rgba(0,0,0,.35)";b.onclick=function(ev){ev.stopPropagation();var o=scrape();if(!o.nom||!o.dob){b.textContent="Identité illisible — ouvre la fiche patient";setTimeout(function(){b.textContent="➕ Créer une fiche Bilan ODF";},2800);return;}try{window.webkit.messageHandlers.bilan.postMessage("doctopat:"+JSON.stringify(o));}catch(e){}b.textContent="Fiche créée ✓";};document.body.appendChild(b);}
mk();setInterval(mk,1500);
})();''')

    _STEINER_JS = (r'''(function(){
if(window.__bilanStInit)return;window.__bilanStInit=1;
var LABELS=["SNA","SNB","ANB","U1 to NA(deg)","U1 to NA(mm)","L1 to NB(deg)","L1 to NB(mm)","Wits appraisal","Pog to NB","Mandibular plane angle(Go-Gn to SN)","Occlusal plane to SN angle","Interincisal angle"];
function isNumLine(s){if(s==="")return false;for(var i=0;i<s.length;i++){var c=s.charAt(i);if(!((c>="0"&&c<="9")||c==="."||c===","||c==="-"||c==="+"))return false;}var v=parseFloat(s.split(",").join("."));return !isNaN(v);}
function scrape(){var lines=(document.body.innerText||"").split(String.fromCharCode(10));var norm=[];for(var i=0;i<lines.length;i++){norm.push(lines[i].trim().split(String.fromCharCode(8722)).join("-"));}var o={};LABELS.forEach(function(lab){for(var i=0;i<norm.length;i++){if(norm[i]===lab){var got=[];for(var j=i+1;j<norm.length&&got.length<3;j++){if(isNumLine(norm[j])){got.push(parseFloat(norm[j].split(",").join(".")));}else if(got.length>0){break;}}if(got.length>=3){o[lab]=got[2];}break;}}});return o;}
function patName(){var lines=(document.body.innerText||"").split(String.fromCharCode(10));for(var i=0;i<lines.length;i++){var ln=lines[i].trim();var op=ln.indexOf("(");var cp=ln.indexOf(")");if(op>0&&cp>op){var code=ln.substring(op+1,cp);if(code.length>=5){var ok=true;for(var k=0;k<code.length;k++){var ch=code.charAt(k);if(!((ch>="A"&&ch<="Z")||(ch>="0"&&ch<="9"))){ok=false;break;}}if(ok){return ln.substring(0,op).trim();}}}}return "";}
function mk(){if(document.getElementById("__bilanStBtn"))return;var b=document.createElement("button");b.id="__bilanStBtn";b.textContent="Importer valeurs Steiner";b.style.cssText="position:fixed;z-index:2147483647;right:16px;bottom:62px;padding:11px 15px;background:#16a34a;color:#fff;border:none;border-radius:9px;font-size:14px;font-family:sans-serif;cursor:pointer;box-shadow:0 3px 12px rgba(0,0,0,.35)";b.onclick=function(ev){ev.stopPropagation();var oo=scrape();var k=Object.keys(oo).length;try{window.webkit.messageHandlers.bilan.postMessage("steiner:"+JSON.stringify({patient:patName(),vals:oo,raw:(document.body.innerText||"").substring(0,1500)}));}catch(e){}b.textContent=k+" valeurs -> Analyses Steiner a classer";setTimeout(function(){b.textContent="Importer valeurs Steiner";},3000);};document.body.appendChild(b);}
mk();setInterval(mk,1500);
})();''')

    class _DL(NSObject):
        # destination des telechargements (radios) -> boite de reception
        def download_decideDestinationUsingResponse_suggestedFilename_completionHandler_(self, dl, resp, name, ch):
            import time as _t
            try: os.makedirs(INBOX, exist_ok=True)
            except Exception: pass
            fn = name or ("radio_%d.jpg" % int(_t.time()))
            dest = os.path.join(INBOX, fn)
            b_, e_ = os.path.splitext(dest); i = 1
            while os.path.exists(dest):
                dest = "%s_%d%s" % (b_, i, e_); i += 1
            ch(NSURL.fileURLWithPath_(dest))
        def downloadDidFinish_(self, dl): pass
        def download_didFailWithError_resumeData_(self, dl, err, rd): pass

    class _Nav(NSObject):
        # si la reponse ne peut pas s'afficher (fichier a telecharger) -> download
        def webView_decidePolicyForNavigationResponse_decisionHandler_(self, web, resp, dh):
            try: can = bool(resp.canShowMIMEType())
            except Exception: can = True
            dh(1 if can else 2)   # 1=Allow, 2=Download
        def webView_navigationResponse_didBecomeDownload_(self, web, resp, download):
            d = _DL.alloc().init(); _browser_windows.append(d); download.setDelegate_(d)
        def webView_navigationAction_didBecomeDownload_(self, web, action, download):
            d = _DL.alloc().init(); _browser_windows.append(d); download.setDelegate_(d)
        def webView_didFinishNavigation_(self, web, nav):
            try: u = str(web.URL().absoluteString())
            except Exception: u = ""
            if (not u.startswith("http")) or ("127.0.0.1" in u):
                return
            _ul = u.lower()
            _which = None
            if "doctolib" in _ul: _js = _DOCTO_JS + _ext_js("doctolib"); _which = "doctolib"
            elif "xero" in _ul: _js = _CAP_JS + _ext_js("xero"); _which = "xero"
            elif "webceph" in _ul: _js = _CAP_JS_SNAP + _STEINER_JS + _ext_js("webceph"); _which = "webceph"
            else: _js = _CAP_JS
            # recherche du patient demandee depuis la bibliotheque / fiche : une seule fois
            if _which:
                _nom = _docstate.get("search_" + _which)
                if _nom:
                    _js += _search_js(_which, _nom)
                    _docstate["search_" + _which] = None
            try: web.evaluateJavaScript_completionHandler_(_js, None)
            except Exception: pass

    class UIDelegate(NSObject):
        def webView_runJavaScriptAlertPanelWithMessage_initiatedByFrame_completionHandler_(self, web, message, frame, completionHandler):
            try:
                from AppKit import NSAlert
                a = NSAlert.alloc().init(); a.setMessageText_("Bilan ODF"); a.setInformativeText_(str(message)); a.addButtonWithTitle_("OK"); a.runModal()
            except Exception: pass
            completionHandler()
        def webView_runJavaScriptConfirmPanelWithMessage_initiatedByFrame_completionHandler_(self, web, message, frame, completionHandler):
            ok = False
            try:
                from AppKit import NSAlert
                a = NSAlert.alloc().init(); a.setMessageText_("Bilan ODF"); a.setInformativeText_(str(message)); a.addButtonWithTitle_("OK"); a.addButtonWithTitle_("Annuler")
                ok = (a.runModal() == 1000)
            except Exception: ok = False
            completionHandler(ok)
        def webView_runJavaScriptTextInputPanelWithPrompt_defaultText_initiatedByFrame_completionHandler_(self, web, prompt, defaultText, frame, completionHandler):
            val = None
            try:
                from AppKit import NSAlert, NSTextField
                a = NSAlert.alloc().init(); a.setMessageText_("Bilan ODF"); a.setInformativeText_(str(prompt)); a.addButtonWithTitle_("OK"); a.addButtonWithTitle_("Annuler")
                tf = NSTextField.alloc().initWithFrame_(NSMakeRect(0, 0, 240, 24)); tf.setStringValue_(defaultText or "")
                a.setAccessoryView_(tf)
                val = tf.stringValue() if a.runModal() == 1000 else None
            except Exception: val = None
            completionHandler(val)
        # INDISPENSABLE : sans ce gestionnaire, WKWebView n'ouvre PAS le selecteur de fichiers.
        def webView_runOpenPanelWithParameters_initiatedByFrame_completionHandler_(
                self, web, params, frame, completionHandler):
            panel = NSOpenPanel.openPanel()
            # Autorise le choix d'un DOSSIER quand la page le demande (input webkitdirectory),
            # sinon reste sur le choix de fichiers. Le sélecteur natif (powerbox) peut atteindre
            # le Bureau/Téléchargements même si l'app n'a pas l'accès disque : c'est la sélection
            # de l'utilisateur qui accorde l'accès.
            try: _allow_dirs = bool(params.allowsDirectories())
            except Exception: _allow_dirs = False
            panel.setCanChooseFiles_(not _allow_dirs)
            panel.setCanChooseDirectories_(_allow_dirs)
            panel.setResolvesAliases_(True)
            try:
                panel.setAllowsMultipleSelection_(bool(params.allowsMultipleSelection()))
            except Exception:
                panel.setAllowsMultipleSelection_(True)
            if panel.runModal() == 1:          # NSModalResponseOK
                completionHandler(panel.URLs())
            else:
                completionHandler(None)
        # ouvre les liens target=_blank / window.open dans une fenetre navigateur INTEGREE
        def webView_createWebViewWithConfiguration_forNavigationAction_windowFeatures_(
                self, web, config, navaction, features):
            try:
                try: config.userContentController().addScriptMessageHandler_name_(_msgh, "bilan")
                except Exception: pass
                r2 = NSMakeRect(0, 0, 1200, 820)
                newweb = WKWebView.alloc().initWithFrame_configuration_(r2, config)
                newweb.setUIDelegate_(self)
                nav = _Nav.alloc().init()
                newweb.setNavigationDelegate_(nav)
                win2 = NSWindow.alloc().initWithContentRect_styleMask_backing_defer_(r2, STYLE, NSBackingStoreBuffered, False)
                win2.setTitle_("Navigateur - Radios (CHU Nice)")
                win2.setMinSize_((700, 500))
                win2.setContentView_(newweb)
                win2.center(); win2.makeKeyAndOrderFront_(None)
                _browser_windows.append((win2, newweb, nav))
                return newweb
            except Exception as _e:
                print("createWebView err", _e); return None

    app = NSApplication.sharedApplication()
    app.setActivationPolicy_(NSApplicationActivationPolicyRegular)
    delegate = AppDelegate.alloc().init()
    app.setDelegate_(delegate)

    # --- Menu principal (INDISPENSABLE pour que Cmd+C / Cmd+V / Cmd+X fonctionnent
    #     dans les champs de la WebView : sans menu Edition, WKWebView ne recoit pas
    #     ces raccourcis, donc IMPOSSIBLE de coller du texte quelque part dans l'app). ---
    try:
        from AppKit import NSMenu, NSMenuItem, NSEventModifierFlagShift, NSEventModifierFlagCommand
        _mainmenu = NSMenu.alloc().init()
        # Menu application (Quitter)
        _appItem = NSMenuItem.alloc().init(); _mainmenu.addItem_(_appItem)
        _appMenu = NSMenu.alloc().init()
        _appMenu.addItemWithTitle_action_keyEquivalent_("Masquer Bilan ODF", "hide:", "h")
        _appMenu.addItem_(NSMenuItem.separatorItem())
        _appMenu.addItemWithTitle_action_keyEquivalent_("Quitter Bilan ODF", "terminate:", "q")
        _appItem.setSubmenu_(_appMenu)
        # Menu Edition (Annuler / Retablir / Couper / Copier / Coller / Tout selectionner)
        _edItem = NSMenuItem.alloc().init(); _mainmenu.addItem_(_edItem)
        _edMenu = NSMenu.alloc().initWithTitle_("Édition")
        _edMenu.addItemWithTitle_action_keyEquivalent_("Annuler", "undo:", "z")
        _rit = _edMenu.addItemWithTitle_action_keyEquivalent_("Rétablir", "redo:", "z")
        try: _rit.setKeyEquivalentModifierMask_(NSEventModifierFlagShift | NSEventModifierFlagCommand)
        except Exception: pass
        _edMenu.addItem_(NSMenuItem.separatorItem())
        _edMenu.addItemWithTitle_action_keyEquivalent_("Couper", "cut:", "x")
        _edMenu.addItemWithTitle_action_keyEquivalent_("Copier", "copy:", "c")
        _edMenu.addItemWithTitle_action_keyEquivalent_("Coller", "paste:", "v")
        _edMenu.addItemWithTitle_action_keyEquivalent_("Tout sélectionner", "selectAll:", "a")
        _edItem.setSubmenu_(_edMenu)
        app.setMainMenu_(_mainmenu)
    except Exception as _e:
        print("menu principal indisponible (%s)" % _e)

    rect = NSMakeRect(0, 0, 1260, 860)
    win = NSWindow.alloc().initWithContentRect_styleMask_backing_defer_(rect, STYLE, NSBackingStoreBuffered, False)
    win.setTitle_("Bilan ODF")
    win.setMinSize_((900, 600))
    conf = WKWebViewConfiguration.alloc().init()
    try: conf.userContentController().addScriptMessageHandler_name_(_msgh, "bilan")
    except Exception: pass
    web = WKWebView.alloc().initWithFrame_configuration_(rect, conf)
    uidelegate = UIDelegate.alloc().init()
    web.setUIDelegate_(uidelegate)
    _mainnav = _Nav.alloc().init(); web.setNavigationDelegate_(_mainnav); _browser_windows.append(_mainnav)          # active le sélecteur de fichiers natif
    container = NSView.alloc().initWithFrame_(rect)
    web.setFrame_(container.bounds()); web.setAutoresizingMask_(18)
    container.addSubview_(web)
    win.setContentView_(container)
    def _toggle_ext(show, which):
        try:
            _urls = {"doctolib": "https://pro.doctolib.fr/", "xero": "https://xero.chu-nice.fr", "webceph": "https://webceph.com/"}
            _titles = {"doctolib": "Doctolib", "xero": "XERO", "webceph": "WebCeph"}
            if show:
                for _k in ("doctolib", "xero", "webceph"):
                    if _k != which:
                        _w = _docstate.get(_k)
                        if _w is not None:
                            try: _w.removeFromSuperview()
                            except Exception: pass
                dw = _docstate.get(which)
                _fresh = dw is None
                if dw is None:
                    dconf = WKWebViewConfiguration.alloc().init()
                    try: dconf.userContentController().addScriptMessageHandler_name_(_msgh, "bilan")
                    except Exception: pass
                    dw = WKWebView.alloc().initWithFrame_configuration_(container.bounds(), dconf)
                    dw.setNavigationDelegate_(_mainnav)
                    dw.setUIDelegate_(uidelegate)
                    dw.setAutoresizingMask_(18)
                    dw.loadRequest_(NSURLRequest.requestWithURL_(NSURL.URLWithString_(_urls[which])))
                    _docstate[which] = dw
                dw.setFrame_(container.bounds())
                container.addSubview_(dw)
                win.setTitle_(_titles[which])
                _docstate["shown"] = which
                # Fenetre deja chargee (pas de nouvelle navigation) : on lance la recherche tout de suite.
                if not _fresh:
                    _nom = _docstate.get("search_" + which)
                    if _nom:
                        try: dw.evaluateJavaScript_completionHandler_(_search_js(which, _nom), None)
                        except Exception: pass
                        _docstate["search_" + which] = None
            else:
                dw = _docstate.get(which)
                if dw is not None: dw.removeFromSuperview()
                win.setTitle_("Bilan ODF")
                _docstate["shown"] = None
        except Exception as _e:
            print("toggle ext err", _e)
    _docstate["toggle"] = _toggle_ext
    web.loadRequest_(NSURLRequest.requestWithURL_(NSURL.URLWithString_(URL)))
    _docstate["mainweb"] = web; _docstate["baseurl"] = URL
    win.center()
    # Recadre sur l'ecran principal visible (evite l'ouverture hors champ si un
    # ancien 2e ecran a laisse une position memorisee negative -> icone 'sans effet').
    try:
        from AppKit import NSScreen as _NSScreen
        _vf = _NSScreen.mainScreen().visibleFrame()
        _cw = min(1260.0, _vf.size.width - 40.0)
        _ch = min(900.0, _vf.size.height - 40.0)
        _nx = _vf.origin.x + (_vf.size.width - _cw) / 2.0
        _ny = _vf.origin.y + (_vf.size.height - _ch) / 2.0
        win.setFrame_display_(NSMakeRect(_nx, _ny, _cw, _ch), True)
    except Exception:
        pass
    # Passage au premier plan robuste (sinon la fenetre s'ouvre DERRIERE le Terminal)
    win.setLevel_(NSFloatingWindowLevel)
    win.makeKeyAndOrderFront_(None)
    app.activateIgnoringOtherApps_(True)
    try:
        NSRunningApplication.currentApplication().activateWithOptions_(NSApplicationActivateIgnoringOtherApps)
    except Exception:
        pass
    win.setLevel_(NSNormalWindowLevel)
    # garder des références vivantes (sinon les delegates sont libérés -> plantage)
    global _refs; _refs = (app, delegate, uidelegate, win, web)
    app.run()

def pywebview_ok():
    try:
        import webview
        jsdir = os.path.join(os.path.dirname(webview.__file__), "js")
        return os.path.isdir(jsdir) and any(f.endswith(".js") for f in os.listdir(jsdir))
    except Exception:
        return False

def _wait_ready(timeout=30):
    """Attend que le serveur Flask accepte les connexions AVANT d'afficher la fenêtre.
    Évite la fenêtre blanche quand la machine est chargée (import massif en cours) et
    que le serveur met plus d'une seconde à démarrer : la WebView ne recharge pas seule."""
    end = time.time() + timeout
    while time.time() < end:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            s.settimeout(0.4); s.connect(("127.0.0.1", PORT)); s.close(); return True
        except Exception:
            try: s.close()
            except Exception: pass
            time.sleep(0.25)
    return False

# ======================================================================
#  MISE A JOUR AUTOMATIQUE (depuis un depot GitHub public)
#  - au demarrage, lit version.json ; si plus recent -> dialogue natif,
#    telecharge les fichiers de code, remplace en place, relance.
#  - ne touche JAMAIS aux donnees patients (~/BilanODF_Data).
#  - silencieux si hors-ligne. Desactive si GH_OWNER vide ou app "gelee".
# ======================================================================
try:
    from app import APP_VERSION          # source unique de la version (definie dans app.py)
except Exception:
    APP_VERSION = "1.0"
GH_OWNER = "bryanfaruch-alt"            # compte GitHub (vide => MAJ desactivee)
GH_REPO = "bilan-odf-app"
GH_BRANCH = "main"

def _vt(s):
    try:
        return tuple(int(x) for x in str(s).strip().split("."))
    except Exception:
        return (0,)

def _app_dir():
    return os.path.dirname(os.path.abspath(__file__))

def _raw_base():
    return "https://raw.githubusercontent.com/%s/%s/%s/" % (GH_OWNER, GH_REPO, GH_BRANCH)

def _ssl_ctx():
    """Contexte SSL avec les certificats CA (certifi) — le Python framework macOS
    n'a pas de CA systeme, sinon 'CERTIFICATE_VERIFY_FAILED'."""
    import ssl
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        try:
            return ssl.create_default_context()
        except Exception:
            return None

def _check_update():
    """Retourne (version_distante, manifest) si une MAJ plus recente existe, sinon None."""
    if not GH_OWNER or getattr(sys, "frozen", False):
        return None
    import json, urllib.request
    url = _raw_base() + "version.json?_=" + str(int(time.time()))
    try:
        req = urllib.request.Request(url, headers={"Cache-Control": "no-cache", "User-Agent": "BilanODF"})
        with urllib.request.urlopen(req, timeout=4, context=_ssl_ctx()) as r:
            man = json.loads(r.read().decode("utf-8"))
    except Exception:
        return None
    remote = str(man.get("version", "")).strip()
    if remote and _vt(remote) > _vt(APP_VERSION):
        return remote, man
    return None

def _download_all(man, dest):
    import urllib.request
    base = man.get("base") or _raw_base()
    files = man.get("files") or []
    if not files:
        raise RuntimeError("manifest sans fichiers")
    got = []
    for rel in files:
        u = base + rel + "?_=" + str(int(time.time()))
        req = urllib.request.Request(u, headers={"Cache-Control": "no-cache", "User-Agent": "BilanODF"})
        with urllib.request.urlopen(req, timeout=40, context=_ssl_ctx()) as r:
            data = r.read()
        if not data:
            raise RuntimeError("fichier vide: %s" % rel)
        p = os.path.join(dest, rel)
        os.makedirs(os.path.dirname(p) or ".", exist_ok=True)
        with open(p, "wb") as f:
            f.write(data)
        got.append(rel)
    return got

def _apply_update(man):
    """Telecharge tout dans un dossier temporaire PUIS remplace en place (avec sauvegarde .bak_upd)."""
    import tempfile, shutil
    d = _app_dir()
    tmp = tempfile.mkdtemp(prefix="bilanodf_upd_")
    try:
        files = _download_all(man, tmp)          # tout ou rien : si un fichier echoue, on garde l'ancien
        for rel in files:
            src = os.path.join(tmp, rel)
            dst = os.path.join(d, rel)
            os.makedirs(os.path.dirname(dst) or ".", exist_ok=True)
            try:
                if os.path.exists(dst):
                    shutil.copy2(dst, dst + ".bak_upd")
            except Exception:
                pass
            shutil.move(src, dst)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

def _update_dialog(remote, notes):
    """Fenetre native. True si l'utilisateur veut installer maintenant."""
    try:
        from AppKit import (NSApplication, NSAlert, NSAlertFirstButtonReturn,
                            NSApplicationActivationPolicyRegular, NSRunningApplication,
                            NSApplicationActivateIgnoringOtherApps)
        app = NSApplication.sharedApplication()
        try:
            app.setActivationPolicy_(NSApplicationActivationPolicyRegular)
            NSRunningApplication.currentApplication().activateWithOptions_(NSApplicationActivateIgnoringOtherApps)
        except Exception:
            pass
        a = NSAlert.alloc().init()
        a.setMessageText_("Mise à jour disponible")
        info = "Une nouvelle version de Bilan ODF est disponible (v%s → v%s)." % (APP_VERSION, remote)
        if notes:
            info += "\n\n%s" % notes
        info += "\n\nInstaller maintenant et relancer l'application ?"
        a.setInformativeText_(info)
        a.addButtonWithTitle_("Installer et relancer")
        a.addButtonWithTitle_("Plus tard")
        return a.runModal() == NSAlertFirstButtonReturn
    except Exception:
        return False

def _info_dialog(msg):
    try:
        from AppKit import NSAlert
        a = NSAlert.alloc().init()
        a.setMessageText_("Bilan ODF")
        a.setInformativeText_(str(msg))
        a.addButtonWithTitle_("OK")
        a.runModal()
    except Exception:
        pass

def maybe_update():
    """Verifie et, si l'utilisateur accepte, installe puis relance. Silencieux sinon."""
    try:
        res = _check_update()
    except Exception:
        res = None
    if not res:
        return
    remote, man = res
    if not _update_dialog(remote, str(man.get("notes", "") or "")):
        return
    try:
        _apply_update(man)
    except Exception as e:
        _info_dialog("La mise à jour a échoué (%s).\nL'application démarre avec la version actuelle." % e)
        return
    # relance le processus avec le nouveau code
    try:
        os.execv(sys.executable, [sys.executable] + sys.argv)
    except Exception as e:
        _info_dialog("Mise à jour installée. Merci de relancer l'application.\n(%s)" % e)
        os._exit(0)

def main():
    # La mise a jour est geree DANS l'app (pop-up HTML fiable) : voir app.py (/maj/state, /maj/apply).
    threading.Thread(target=serve, daemon=True).start()
    _wait_ready()          # attend le serveur (au lieu d'un simple sleep) -> plus de fenêtre blanche
    try:
        run_wkwebview(); return
    except Exception as e:
        print("Fenêtre WebKit native indisponible (%s)." % e)
    if pywebview_ok():
        try:
            import webview
            webview.create_window("Bilan ODF", URL, width=1260, height=860, min_size=(900, 600))
            webview.start(); return
        except Exception as e:
            print("pywebview indisponible (%s)." % e)
    import webbrowser
    print("Ouverture dans le navigateur par défaut.")
    webbrowser.open(URL)
    try:
        while True: time.sleep(3600)
    except KeyboardInterrupt:
        pass

if __name__ == "__main__":
    main()
