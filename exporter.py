"""Export de toute la base vers un classeur Excel (Commandes, Paiements, Mr, Rendez-vous, Equipe)."""
import io
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

NAVY = "1F4E78"
AR = '#,##0" Ar";-#,##0" Ar";"-"'


def _date(s):
    try:
        return datetime.strptime(s, "%Y-%m-%d") if s else None
    except ValueError:
        return s


def _sheet(wb, title, headers, rows, money=(), dates=(), widths=None):
    ws = wb.create_sheet(title)
    ws.append(headers)
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in rows:
        ws.append(list(r))
    for col in range(1, len(headers) + 1):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = (widths or {}).get(col, 16)
        for cell in ws[letter][1:]:
            cell.font = Font(name="Arial", size=10)
            if col in money:
                cell.number_format = AR
            if col in dates:
                cell.number_format = "dd/mm/yyyy"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    return ws


def build_workbook(db, cmds, modes):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    _sheet(wb, "Commandes",
           ["N° commande", "Date", "Client", "Téléphone", "Situation", "Produit / Service", "Quantité", "Prix unitaire (Ar)",
            "Montant (Ar)", "Payé (Ar)", "Reste (Ar)", "Statut", "Statut paiement", "Livraison", "Date livraison",
            "Reçu par", "Responsable", "Équipe message", "En ligne / local", "Remarques"],
           [[c["code"], _date(c["date"]), c["client"], c["tel"], c["situation"], c["produit"], c["qte"], c["pu"],
             c["total"] or None, c["paid"] or None, c["reste"] or None, c["statut"], c["sp"], c["livraison"],
             _date(c["date_livraison"]), c["recu_par"], c["responsable"], c["equipe_msg"], c["canal"], c["remarques"]] for c in cmds],
           money=(8, 9, 10, 11), dates=(2, 15), widths={3: 30, 6: 45, 20: 40})
    pays = db.execute("""SELECT p.*, c.client, u.nom AS saisi_nom, ua.nom AS annule_nom FROM paiements p
                         LEFT JOIN commandes c ON c.id=p.commande_id LEFT JOIN users u ON u.id=p.saisi_par
                         LEFT JOIN users ua ON ua.id=p.annule_par ORDER BY p.date, p.id""").fetchall()
    _sheet(wb, "Paiements",
           ["Date paiement", "N° commande", "Client", "Montant (Ar)", "Mode", "Référence", "Nandray ny vola (encaissé par)",
            "Nampiditra (saisi par)", "Saisi le", "Annulé", "Annulé par", "Motif", "Remarque"],
           [[_date(p["date"]), f"CMD-{p['commande_id']:03d}", p["client"], p["montant"], p["mode"], p["reference"],
             p["encaisse_par"], p["saisi_nom"] or ("Import Excel" if p["importe"] else ""), p["saisi_le"],
             "Oui" if p["annule"] else "", p["annule_nom"], p["motif_annulation"], p["remarque"]] for p in pays],
           money=(4,), dates=(1,), widths={3: 30, 13: 40})
    # Mr : par mois
    months = {}
    for c in cmds:
        if not c["date"]:
            continue
        m = months.setdefault(c["date"][:7], [0, 0, 0, 0])
        m[0] += 1
        if c["statut"] != "Annulé":
            m[1] += 1 if c["total"] > 0 else 0
            m[2] += c["total"]
            m[3] += c["paid"]
    enc = {}
    for p in pays:
        if p["annule"]:
            continue
        k = p["mode"] if p["mode"] in modes else "À préciser"
        enc.setdefault(p["date"][:7], {}).setdefault(k, 0)
        enc[p["date"][:7]][k] += p["montant"]
    cols = modes + ["À préciser"]
    keys = sorted(set(months) | set(enc))
    _sheet(wb, "Mr (par mois)",
           ["Mois", "Nombre de messages", "Nombre de projets", "Montant des projets (Ar)", "Avances reçues (Ar)",
            "Reste à payer (Ar)"] + [f"Encaissé {m} (Ar)" for m in cols],
           [[k, *(months.get(k, [0, 0, 0, 0])[:4]), months.get(k, [0, 0, 0, 0])[2] - months.get(k, [0, 0, 0, 0])[3]] +
            [enc.get(k, {}).get(m, 0) for m in cols] for k in keys],
           money=tuple(range(4, 7 + len(cols))))
    rdv = db.execute("SELECT * FROM rdv ORDER BY date").fetchall()
    _sheet(wb, "Rendez-vous", ["Projet", "Téléphone", "Lieu", "Frais", "Date", "Heure", "Responsable", "Statut", "Notes"],
           [[r["projet"], r["tel"], r["lieu"], r["frais"], _date(r["date"]), r["heure"], r["responsable"], r["statut"], r["notes"]]
            for r in rdv], dates=(5,), widths={9: 40})
    eq = db.execute("SELECT * FROM equipe ORDER BY id").fetchall()
    _sheet(wb, "Equipe", ["Nom", "Rôle", "Téléphone", "Actif"], [[e["nom"], e["role"], e["tel"], "Oui" if e["actif"] else "Non"] for e in eq])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
