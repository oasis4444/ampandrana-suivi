"""Import du classeur Excel « Ampandrana_Suivi_des_commandes » dans la base.

Lit les feuilles Commandes, Paiements (si présente), Equipe, Listes et RENDEZ- VOUS.
Sans feuille Paiements (ancienne version), chaque « Avance payée » devient un paiement
avec le « Mode de paiement » de la commande, ou « À préciser ».
"""
import re
from datetime import datetime

import openpyxl

A_PRECISER = "À préciser"


def _d(v):
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, str):
        m = re.match(r"^\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*$", v)
        if m:
            return f"{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"
    return None


def _s(v):
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    if isinstance(v, int):
        return "0" + str(v) if 300000000 <= v < 400000000 else str(v)   # téléphone sans le 0
    if isinstance(v, datetime):
        return v.strftime("%d/%m/%Y")
    v = str(v).strip()
    return v or None


def _n(v):
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.replace(" ", "").replace(" ", "").replace(",", "."))
        except ValueError:
            return None
    return None


def _header_map(ws, row):
    out = {}
    for c in ws[row]:
        if c.value:
            out[str(c.value).strip().lower()] = c.column
    return out


def _col(hm, *starts, default=None):
    for key, col in hm.items():
        for s in starts:
            if key.startswith(s.lower()):
                return col
    return default


def import_workbook(db, path, replace=False):
    wb = openpyxl.load_workbook(path, data_only=True)
    if "Commandes" not in wb.sheetnames:
        raise ValueError("feuille « Commandes » introuvable")
    if replace:
        for t in ("paiements", "commandes", "rdv", "equipe"):
            db.execute(f"DELETE FROM {t}")

    # --- Equipe
    n_eq = 0
    if "Equipe" in wb.sheetnames:
        ws = wb["Equipe"]
        for r in range(4, ws.max_row + 1):
            nom = _s(ws.cell(r, 1).value)
            if nom and not db.execute("SELECT 1 FROM equipe WHERE nom=?", (nom,)).fetchone():
                db.execute("INSERT INTO equipe(nom, role, tel) VALUES (?,?,?)", (nom, _s(ws.cell(r, 2).value), _s(ws.cell(r, 3).value)))
                n_eq += 1

    # --- Listes (ajoute les valeurs manquantes)
    if "Listes" in wb.sheetnames:
        ws = wb["Listes"]
        cats = {1: "statut_commande", 2: "statut_livraison", 3: "mode_paiement"}
        for col, cat in cats.items():
            for r in range(2, 8):
                v = _s(ws.cell(r, col).value)
                if v and v != A_PRECISER:
                    n = db.execute("SELECT COALESCE(MAX(ordre),0)+1 FROM listes WHERE categorie=?", (cat,)).fetchone()[0]
                    db.execute("INSERT OR IGNORE INTO listes(categorie, valeur, ordre) VALUES (?,?,?)", (cat, v, n))

    # --- Commandes (en-têtes ligne 4, données à partir de la ligne 5 ; N° = ligne - 4)
    ws = wb["Commandes"]
    hm = _header_map(ws, 4)
    C = {
        "date": _col(hm, "date commande", default=2), "client": _col(hm, "nom du client", default=3),
        "tel": _col(hm, "téléphone", default=4), "situation": _col(hm, "situation", default=5),
        "produit": _col(hm, "produit", default=6), "qte": _col(hm, "quantité", default=7),
        "pu": _col(hm, "prix unitaire", default=8), "total": _col(hm, "montant total", default=9),
        "paye": _col(hm, "avance payée", default=10), "statut": _col(hm, "statut commande", default=12),
        "livraison": _col(hm, "statut livraison", default=14), "date_livraison": _col(hm, "date livraison", default=15),
        "recu_par": _col(hm, "reçu par", default=16), "responsable": _col(hm, "responsable", default=17),
        "equipe_msg": _col(hm, "équipe qui a répondu", default=18), "mode": _col(hm, "mode de paiement", default=19),
        "remarques": _col(hm, "remarques", default=20), "canal": _col(hm, "statu en ligne", "statut en ligne", "en ligne", default=22),
    }
    n_cmd = 0
    legacy = []
    for r in range(5, ws.max_row + 1):
        g = lambda k: ws.cell(r, C[k]).value if C[k] else None  # noqa: E731
        if not any(g(k) not in (None, "") for k in ("date", "client", "produit")):
            continue
        cid = r - 4
        qte, pu = _n(g("qte")), _n(g("pu"))
        if pu is None and _n(g("total")):          # montant saisi sans prix unitaire
            pu, qte = _n(g("total")), 1
        db.execute("""INSERT OR REPLACE INTO commandes(id, date, client, tel, situation, produit, qte, pu, statut, livraison,
                      date_livraison, recu_par, responsable, equipe_msg, canal, remarques)
                      VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                   (cid, _d(g("date")), _s(g("client")) or "(sans nom)", _s(g("tel")), _s(g("situation")), _s(g("produit")),
                    qte, pu, _s(g("statut")), _s(g("livraison")), _d(g("date_livraison")), _s(g("recu_par")),
                    _s(g("responsable")), _s(g("equipe_msg")), _s(g("canal")), _s(g("remarques"))))
        n_cmd += 1
        if _n(g("paye")):
            legacy.append((cid, _d(g("date")), _n(g("paye")), _s(g("mode")) or A_PRECISER, _s(g("remarques"))))

    # --- Paiements
    n_pay, total = 0, 0.0
    if "Paiements" in wb.sheetnames:
        ws = wb["Paiements"]
        for r in range(5, ws.max_row + 1):
            montant = _n(ws.cell(r, 4).value)
            if not montant:
                continue
            m = re.search(r"(\d+)", str(ws.cell(r, 2).value or ""))
            if not m:
                continue
            db.execute("""INSERT INTO paiements(commande_id, date, montant, mode, reference, encaisse_par, remarque, importe)
                          VALUES (?,?,?,?,?,?,?,1)""",
                       (int(m.group(1)), _d(ws.cell(r, 1).value) or "2000-01-01", montant, _s(ws.cell(r, 5).value) or A_PRECISER,
                        _s(ws.cell(r, 6).value), _s(ws.cell(r, 7).value),
                        (_s(ws.cell(r, 9).value) or "").replace("Nafindra avy amin'ny Commandes (daty = daty commande)", "").lstrip(" —") or None))
            n_pay += 1
            total += montant
    else:
        for cid, d, montant, mode, rem in legacy:
            db.execute("""INSERT INTO paiements(commande_id, date, montant, mode, remarque, importe) VALUES (?,?,?,?,?,1)""",
                       (cid, d or "2000-01-01", montant, mode, rem))
            n_pay += 1
            total += montant

    # --- Rendez-vous
    n_rdv = 0
    sheet = next((s for s in wb.sheetnames if s.upper().replace(" ", "").startswith("RENDEZ")), None)
    if sheet:
        ws = wb[sheet]
        for r in range(2, ws.max_row + 1):
            projet = _s(ws.cell(r, 1).value)
            if not projet:
                continue
            raw_date = ws.cell(r, 5).value
            d = _d(raw_date)
            heure = _s(ws.cell(r, 6).value)
            if not d and raw_date:
                heure = " ".join(x for x in (_s(raw_date), heure) if x)
            notes = " / ".join(x for x in (_s(ws.cell(r, c).value) for c in range(9, 12)) if x) or None
            db.execute("""INSERT INTO rdv(projet, tel, lieu, frais, date, heure, responsable, statut, notes)
                          VALUES (?,?,?,?,?,?,?,?,?)""",
                       (projet, _s(ws.cell(r, 2).value), _s(ws.cell(r, 3).value), _s(ws.cell(r, 4).value), d, heure,
                        _s(ws.cell(r, 7).value), _s(ws.cell(r, 8).value), notes))
            n_rdv += 1
    return {"commandes": n_cmd, "paiements": n_pay, "montant": total, "equipe": n_eq, "rdv": n_rdv}
