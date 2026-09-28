# -*- coding: utf-8 -*-
"""
Sauvegarde et restauration des données de Bilan ODF.

- Crée une archive .zip horodatée de tout le dossier de données (patients +
  configuration), dans un dossier de sauvegardes séparé.
- Chiffrement AES-256 optionnel protégé par mot de passe (via pyzipper) : idéal
  pour une copie qui sort de la machine (clé USB, cloud). Sans mot de passe, zip
  standard.
- Restauration : remet une sauvegarde en place, après avoir mis les données
  actuelles de côté (sécurité, jamais d'effacement définitif).
- Purge : ne conserve que les N sauvegardes les plus récentes.

Le chiffrement « au repos » du dossier de données vivant relève de macOS
(FileVault, recommandé) ; ici on garantit surtout des SAUVEGARDES chiffrées.
"""
import os
import shutil
import zipfile

try:
    import pyzipper  # zip AES-256
    HAVE_AES = True
except Exception:
    HAVE_AES = False


def _iter_files(root):
    for dirpath, _dirs, files in os.walk(root):
        # on ignore le dossier des sauvegardes lui-même et les corbeilles volumineuses ? -> on garde tout sauf le stage
        for f in files:
            full = os.path.join(dirpath, f)
            if os.path.sep + "_import_stage" + os.path.sep in full:
                continue
            yield full


def create_backup(data_dir, backup_dir, stamp, password=None, keep=15):
    """Crée une sauvegarde datée. `stamp` = 'AAAAMMJJ_HHMMSS' (fourni par l'appelant).
    Retourne le chemin du fichier créé."""
    os.makedirs(backup_dir, exist_ok=True)
    name = "BilanODF_sauvegarde_%s.zip" % stamp
    dest = os.path.join(backup_dir, name)
    k = 2
    while os.path.exists(dest):   # évite d'écraser une sauvegarde de la même seconde
        dest = os.path.join(backup_dir, "BilanODF_sauvegarde_%s_%d.zip" % (stamp, k)); k += 1
    files = list(_iter_files(data_dir))
    if password and HAVE_AES:
        with pyzipper.AESZipFile(dest, "w", compression=pyzipper.ZIP_DEFLATED,
                                 encryption=pyzipper.WZ_AES) as z:
            z.setpassword(password.encode("utf-8"))
            for full in files:
                z.write(full, os.path.relpath(full, data_dir))
    else:
        with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as z:
            for full in files:
                z.write(full, os.path.relpath(full, data_dir))
    _prune(backup_dir, keep)
    return dest


def _prune(backup_dir, keep):
    try:
        zips = sorted([f for f in os.listdir(backup_dir) if f.startswith("BilanODF_sauvegarde_") and f.endswith(".zip")])
        for old in zips[:-keep] if keep and len(zips) > keep else []:
            try: os.remove(os.path.join(backup_dir, old))
            except Exception: pass
    except Exception:
        pass


def list_backups(backup_dir):
    """Retourne [(nom, taille_octets, mtime)] du plus récent au plus ancien."""
    out = []
    if not os.path.isdir(backup_dir):
        return out
    for f in os.listdir(backup_dir):
        if f.startswith("BilanODF_sauvegarde_") and f.endswith(".zip"):
            p = os.path.join(backup_dir, f)
            try:
                out.append((f, os.path.getsize(p), os.path.getmtime(p)))
            except Exception:
                pass
    out.sort(key=lambda x: x[2], reverse=True)
    return out


def is_encrypted(path):
    """Détecte si un zip est chiffré (AES ou classique)."""
    try:
        with zipfile.ZipFile(path) as z:
            for info in z.infolist():
                if info.flag_bits & 0x1:
                    return True
        return False
    except Exception:
        # AESZipFile peut lever avec zipfile standard : on suppose chiffré si pyzipper le lit
        return True


def restore_backup(path, data_dir, password=None):
    """Restaure une sauvegarde : extrait dans un dossier temporaire, met les
    données actuelles de côté (…_avant_restauration_STAMP) puis remet en place.
    Retourne (ok, message)."""
    if not os.path.exists(path):
        return False, "Sauvegarde introuvable."
    tmp = data_dir.rstrip("/\\") + "_restore_tmp"
    if os.path.isdir(tmp):
        shutil.rmtree(tmp, ignore_errors=True)
    os.makedirs(tmp, exist_ok=True)
    try:
        opened = False
        if password and HAVE_AES:
            try:
                with pyzipper.AESZipFile(path) as z:
                    z.setpassword(password.encode("utf-8"))
                    z.extractall(tmp)
                    opened = True
            except Exception:
                opened = False
        if not opened:
            with zipfile.ZipFile(path) as z:
                z.extractall(tmp, pwd=password.encode("utf-8") if password else None)
    except RuntimeError:
        shutil.rmtree(tmp, ignore_errors=True)
        return False, "Mot de passe incorrect ou sauvegarde chiffrée."
    except Exception as e:
        shutil.rmtree(tmp, ignore_errors=True)
        return False, "Lecture impossible : %s" % e
    # bascule : patients/ et config.json
    try:
        for name in ("patients", "config.json"):
            live = os.path.join(data_dir, name)
            newp = os.path.join(tmp, name)
            if os.path.exists(newp):
                if os.path.exists(live):
                    safety = live + "_avant_restauration"
                    if os.path.exists(safety):
                        shutil.rmtree(safety, ignore_errors=True) if os.path.isdir(safety) else os.remove(safety)
                    shutil.move(live, safety)
                shutil.move(newp, live)
        return True, "Restauration effectuée."
    except Exception as e:
        return False, "Restauration partielle : %s" % e
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
