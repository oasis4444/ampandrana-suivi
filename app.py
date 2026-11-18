"""Ampandrana Suivi — commandes, paiements (Espèces / MVola / Orange Money…), Mr, tableau de bord.

Lancer en local :   flask --app app run --debug
PythonAnywhere :    voir README.md
"""
import csv

import click
import io
import os
import secrets
import sqlite3
from datetime import date, datetime, timedelta
from functools import wraps
from zoneinfo import ZoneInfo

from flask import (Flask, abort, flash, g, redirect, render_template, request,
                   send_file, session, url_for)
from werkzeug.security import check_password_hash, generate_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TZ = ZoneInfo(os.environ.get("APP_TZ", "Indian/Antananarivo"))

MOIS = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août",
        "septembre", "octobre", "novembre", "décembre"]
JOURS = ["lun.", "mar.", "mer.", "jeu.", "ven.", "sam.", "dim."]
A_PRECISER = "À préciser"
MOBILE = ("MVola", "Orange Money", "Airtel Money")
DEFAULT_LISTES = {
    "statut_commande": ["OK", "En attente", "En cours", "Annulé"],
    "statut_livraison": ["Non livré", "En préparation", "En cours de livraison", "Livré"],
    "mode_paiement": ["Espèces", "MVola", "Orange Money", "Airtel Money", "Virement"],
    "canal": ["En ligne", "Local", "Hybride"],
}
LISTE_LABELS = {
    "statut_commande": "Statut commande",
    "statut_livraison": "Statut livraison",
    "mode_paiement": "Mode de paiement",
    "canal": "En ligne / local",
}
MODE_COLORS = {"Espèces": "#2e7d4f", "MVola": "#c9a200", "Orange Money": "#e56b00",
               "Airtel Money": "#c62828", "Virement": "#5b6b8c", A_PRECISER: "#8a94a0"}

app = Flask(__name__, root_path=BASE_DIR, instance_path=os.path.join(BASE_DIR, "instance"))
os.makedirs(app.instance_path, exist_ok=True)
app.config["DATABASE"] = os.environ.get("DATABASE", os.path.join(app.instance_path, "ampandrana.db"))
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024


def _secret_key():
    if os.environ.get("SECRET_KEY"):
        return os.environ["SECRET_KEY"]
    path = os.path.join(app.instance_path, "secret_key")
    if not os.path.exists(path):
        with open(path, "w") as f:
            f.write(secrets.token_hex(32))
    with open(path) as f:
        return f.read().strip()


app.secret_key = _secret_key()


# ---------------------------------------------------------------- database
def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(app.config["DATABASE"])
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db


@app.teardown_appcontext
def close_db(_exc):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db():
    db = get_db()
    with open(os.path.join(BASE_DIR, "schema.sql"), encoding="utf-8") as f:
        db.executescript(f.read())
    if not db.execute("SELECT 1 FROM listes LIMIT 1").fetchone():
        for cat, vals in DEFAULT_LISTES.items():
            for i, v in enumerate(vals):
                db.execute("INSERT INTO listes(categorie, valeur, ordre) VALUES (?,?,?)", (cat, v, i))
    db.commit()


@app.cli.command("init-db")
def init_db_command():
    """Crée les tables."""
    init_db()
    print("Base initialisée :", app.config["DATABASE"])


@app.cli.command("import-excel")
@click.argument("path")
def import_excel_command(path):
    """flask --app app import-excel fichier.xlsx  (remplace les données existantes)"""
    init_db()
    from ampandrana_importer import import_workbook
    res = import_workbook(get_db(), path, replace=True)
    get_db().commit()
    print(res)


_db_ready = False


@app.before_request
def ensure_db():
    global _db_ready
    if not _db_ready:
        init_db()
        _db_ready = True


# ---------------------------------------------------------------- helpers
def now_iso():
    return datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")


def today():
    return datetime.now(TZ).date()


def iso(d):
    return d.strftime("%Y-%m-%d")


def parse_date(s):
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except (TypeError, ValueError):
        return None


def to_num(s):
    if s is None:
        return None
    s = str(s).replace(" ", "").replace(" ", "").replace(",", ".").strip()
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def period_range(periode, du=None, au=None):
    t = today()
    if periode == "today":
        return iso(t), iso(t)
    if periode == "month":
        first = t.replace(day=1)
        nxt = (first + timedelta(days=32)).replace(day=1)
        return iso(first), iso(nxt - timedelta(days=1))
    if periode == "prevmonth":
        first_this = t.replace(day=1)
        last_prev = first_this - timedelta(days=1)
        return iso(last_prev.replace(day=1)), iso(last_prev)
    if periode == "custom":
        return (du or "0000-01-01"), (au or "9999-12-31")
    return "0000-01-01", "9999-12-31"


def period_label(periode, a, b):
    if periode == "all" or (a == "0000-01-01" and b == "9999-12-31"):
        return "Toutes les dates"
    if a == b:
        return fmt_date(a)
    return f"du {fmt_date(a) if a != '0000-01-01' else '…'} au {fmt_date(b) if b != '9999-12-31' else '…'}"


def liste(cat):
    return [r["valeur"] for r in get_db().execute(
        "SELECT valeur FROM listes WHERE categorie=? ORDER BY ordre, valeur", (cat,))]


def modes():
    return liste("mode_paiement")


def equipe_noms(include=None):
    noms = [r["nom"] for r in get_db().execute("SELECT nom FROM equipe WHERE actif=1 ORDER BY id")]
    if include and include not in noms:
        noms.append(include)
    return noms


def log(action, objet, details=""):
    get_db().execute("INSERT INTO journal(quand, user_id, action, objet, details) VALUES (?,?,?,?,?)",
                     (now_iso(), g.user["id"] if g.get("user") else None, action, objet, details))


def cmd_code(n):
    return f"CMD-{int(n):03d}"


def cmd_calc(c):
    """Montant, payé, reste et statut de paiement d'une commande (règles de l'Excel)."""
    pu = c["pu"]
    total = (c["qte"] or 1) * pu if pu else 0
    paid = c["paye"] or 0
    annule = c["statut"] == "Annulé"
    reste = 0 if (not total or annule) else total - paid
    if total:
        sp = "Annulé" if annule else ("Payé" if paid >= total else ("Avance versée" if paid > 0 else "Non payé"))
    else:
        sp = "Avance versée" if paid > 0 else ""
    return {"total": total, "paid": paid, "reste": reste, "sp": sp}


def commandes_with_paid(where="", params=()):
    sql = f"""SELECT c.*, COALESCE((SELECT SUM(p.montant) FROM paiements p
                                     WHERE p.commande_id=c.id AND p.annule=0),0) AS paye
              FROM commandes c {where}"""
    rows = []
    for r in get_db().execute(sql, params):
        d = dict(r)
        d.update(cmd_calc(r))
        d["code"] = cmd_code(r["id"])
        rows.append(d)
    return rows


def pay_issues(p, cmd_exists=True):
    out = []
    if not p["mode"] or p["mode"] == A_PRECISER:
        out.append("mode à préciser")
    if not p["encaisse_par"]:
        out.append("encaissé par ?")
    if p["mode"] in MOBILE and not p["reference"]:
        out.append("référence manquante")
    if not cmd_exists:
        out.append("commande introuvable")
    return out


# ---------------------------------------------------------------- template filters
def fmt_ar(v):
    v = round(v or 0)
    return f"{v:,.0f}".replace(",", " ") + " Ar"


def fmt_num(v):
    if not v:
        return "–"
    return f"{round(v):,.0f}".replace(",", " ")


def fmt_date(s):
    if not s or len(str(s)) < 10:
        return s or ""
    s = str(s)
    return f"{s[8:10]}/{s[5:7]}/{s[0:4]}"


def fmt_dt(s):
    if not s:
        return ""
    return fmt_date(s[:10]) + " " + s[11:16]


app.jinja_env.filters.update(ar=fmt_ar, num=fmt_num, fd=fmt_date, fdt=fmt_dt, code=cmd_code)


@app.context_processor
def inject():
    return {"csrf_token": csrf_token, "MODE_COLORS": MODE_COLORS, "A_PRECISER": A_PRECISER,
            "MOBILE": MOBILE, "today_iso": iso(today())}


# ---------------------------------------------------------------- auth & csrf
def csrf_token():
    if "_csrf" not in session:
        session["_csrf"] = secrets.token_hex(16)
    return session["_csrf"]


@app.before_request
def load_user_and_check_csrf():
    g.user = None
    uid = session.get("uid")
    if uid:
        g.user = get_db().execute("SELECT * FROM users WHERE id=? AND actif=1", (uid,)).fetchone()
    if request.method == "POST":
        if request.form.get("_csrf") != session.get("_csrf"):
            abort(400, "Formulaire expiré : rechargez la page et recommencez.")


def login_required(view):
    @wraps(view)
    def wrapped(*a, **kw):
        if not get_db().execute("SELECT 1 FROM users LIMIT 1").fetchone():
            return redirect(url_for("setup"))
        if g.user is None:
            return redirect(url_for("login", next=request.path))
        return view(*a, **kw)
    return wrapped


def admin_required(view):
    @wraps(view)
    @login_required
    def wrapped(*a, **kw):
        if g.user["role"] != "admin":
            abort(403)
        return view(*a, **kw)
    return wrapped


@app.route("/setup", methods=["GET", "POST"])
def setup():
    db = get_db()
    if db.execute("SELECT 1 FROM users LIMIT 1").fetchone():
        return redirect(url_for("login"))
    if request.method == "POST":
        u, nom, pw = (request.form.get(k, "").strip() for k in ("username", "nom", "password"))
        if not u or not nom or len(pw) < 6:
            flash("Remplissez tous les champs (mot de passe : 6 caractères minimum).", "bad")
        else:
            db.execute("INSERT INTO users(username, nom, password_hash, role, created_at) VALUES (?,?,?,?,?)",
                       (u.lower(), nom, generate_password_hash(pw), "admin", now_iso()))
            db.commit()
            flash("Compte administrateur créé. Connectez-vous.", "good")
            return redirect(url_for("login"))
    return render_template("setup.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        u = request.form.get("username", "").strip().lower()
        row = get_db().execute("SELECT * FROM users WHERE username=? AND actif=1", (u,)).fetchone()
        if row and check_password_hash(row["password_hash"], request.form.get("password", "")):
            session.clear()
            session["uid"] = row["id"]
            session.permanent = True
            nxt = request.args.get("next") or url_for("dashboard")
            return redirect(nxt if nxt.startswith("/") else url_for("dashboard"))
        flash("Identifiant ou mot de passe incorrect.", "bad")
    return render_template("login.html")


@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/mon-compte", methods=["GET", "POST"])
@login_required
def mon_compte():
    if request.method == "POST":
        old, new = request.form.get("old", ""), request.form.get("new", "")
        if not check_password_hash(g.user["password_hash"], old):
            flash("Ancien mot de passe incorrect.", "bad")
        elif len(new) < 6:
            flash("Le nouveau mot de passe doit avoir 6 caractères minimum.", "bad")
        else:
            get_db().execute("UPDATE users SET password_hash=? WHERE id=?", (generate_password_hash(new), g.user["id"]))
            get_db().commit()
            flash("Mot de passe modifié.", "good")
    return render_template("mon_compte.html")


# ---------------------------------------------------------------- dashboard
@app.route("/")
@login_required
def index():
    return redirect(url_for("dashboard"))


@app.route("/tableau-de-bord")
@login_required
def dashboard():
    db = get_db()
    periode = request.args.get("periode", "month")
    a, b = period_range(periode, request.args.get("du"), request.args.get("au"))
    mlist = modes()
    pays = db.execute("""SELECT p.*, u.nom AS saisi_nom FROM paiements p LEFT JOIN users u ON u.id=p.saisi_par
                         WHERE p.annule=0 AND p.date BETWEEN ? AND ?""", (a, b)).fetchall()

    def mkey(m):
        return m if m in mlist else A_PRECISER
    by_mode = {m: {"s": 0, "n": 0} for m in mlist + [A_PRECISER]}
    total = 0
    members, saisie = {}, {}
    for p in pays:
        m = mkey(p["mode"])
        by_mode[m]["s"] += p["montant"]
        by_mode[m]["n"] += 1
        total += p["montant"]
        k = p["encaisse_par"] or ""
        members.setdefault(k, {mm: 0 for mm in mlist + [A_PRECISER, "total"]})
        members[k][m] += p["montant"]
        members[k]["total"] += p["montant"]
        sk = p["saisi_nom"] or ""
        s = saisie.setdefault(sk, {"n": 0, "s": 0, "last": ""})
        s["n"] += 1
        s["s"] += p["montant"]
        s["last"] = max(s["last"], p["saisi_le"] or "")
    order = equipe_noms()
    member_rows = [(k, members[k]) for k in order if k in members]
    member_rows += [(k, v) for k, v in members.items() if k not in order and k]
    if "" in members:
        member_rows.append(("", members[""]))
    mobile = sum(by_mode[m]["s"] for m in MOBILE if m in by_mode)

    # alerts (all time)
    allp = db.execute("""SELECT p.*, (SELECT 1 FROM commandes c WHERE c.id=p.commande_id) AS ok
                         FROM paiements p WHERE annule=0""").fetchall()
    alerts = {
        "prec": sum(1 for p in allp if p["mode"] not in mlist),
        "enc": sum(1 for p in allp if not p["encaisse_par"]),
        "ref": sum(1 for p in allp if p["mode"] in MOBILE and not p["reference"]),
        "lost": sum(1 for p in allp if not p["ok"]),
    }
    cmds = commandes_with_paid()
    stats = {"n": len(cmds), "montant": sum(c["total"] for c in cmds if c["statut"] != "Annulé"),
             "encaisse": sum(p["montant"] for p in allp), "reste": sum(c["reste"] for c in cmds),
             "par_statut": {s: sum(1 for c in cmds if c["statut"] == s) for s in liste("statut_commande")}}
    chart = monthly_by_mode(6)
    livraisons = livraisons_a_surveiller(cmds)
    return render_template("dashboard.html", livraisons=livraisons, periode=periode, du=request.args.get("du", ""),
                           au=request.args.get("au", ""), label=period_label(periode, a, b),
                           by_mode=by_mode, total=total, npay=len(pays), mobile=mobile, mlist=mlist,
                           member_rows=member_rows, saisie=sorted(saisie.items(), key=lambda x: -x[1]["s"]),
                           alerts=alerts, stats=stats, chart=chart)


def livraisons_a_surveiller(cmds, jours=3):
    """Projets non livrés dont la date de livraison est passée (en retard) ou arrive dans `jours` jours."""
    t = today()
    retard, bientot = [], []
    for c in cmds:
        d = parse_date(c.get("date_livraison"))
        if not d or c.get("livraison") == "Livré" or c.get("statut") == "Annulé":
            continue
        c = dict(c, jours=(t - d).days)
        if d < t:
            retard.append(c)
        elif d <= t + timedelta(days=jours):
            bientot.append(c)
    retard.sort(key=lambda c: c["date_livraison"])
    bientot.sort(key=lambda c: c["date_livraison"])
    return {"retard": retard, "bientot": bientot, "jours": jours}


def monthly_by_mode(n):
    t = today().replace(day=1)
    months = []
    for i in range(n - 1, -1, -1):
        y, m = t.year, t.month - i
        while m <= 0:
            m += 12
            y -= 1
        months.append(f"{y}-{m:02d}")
    mlist = modes()
    data = {mm: [0] * n for mm in mlist + [A_PRECISER]}
    for r in get_db().execute("""SELECT substr(date,1,7) AS ym, mode, SUM(montant) AS s FROM paiements
                                  WHERE annule=0 AND substr(date,1,7) >= ? GROUP BY ym, mode""", (months[0],)):
        if r["ym"] in months:
            k = r["mode"] if r["mode"] in mlist else A_PRECISER
            data[k][months.index(r["ym"])] += r["s"]
    labels = [MOIS[int(m[5:]) - 1][:4] + ". " + m[:4] for m in months]
    return {"labels": labels, "series": [{"label": k, "data": v, "color": MODE_COLORS.get(k, "#5b6b8c")}
                                         for k, v in data.items() if any(v)]}


# ---------------------------------------------------------------- Mr (résumé jour / mois)
@app.route("/mr")
@login_required
def mr():
    db = get_db()
    mlist = modes()
    cmds = commandes_with_paid("WHERE c.date IS NOT NULL")
    pays = db.execute("SELECT date, mode, montant FROM paiements WHERE annule=0").fetchall()
    dates = [c["date"] for c in cmds] + [p["date"] for p in pays]
    if not dates:
        return render_template("mr.html", days=[], months=[], totals=None, mois_sel="", mois_opts=[], mlist=mlist, chart=None)
    start = parse_date(min(dates))
    end = max(today(), parse_date(max(dates)))

    def blank():
        return {"messages": 0, "projets": 0, "montant": 0, "avances": 0, "reste": 0,
                "enc": {m: 0 for m in mlist + [A_PRECISER]}, "enc_total": 0}
    days, months = {}, {}
    d = start
    while d <= end:
        days[iso(d)] = blank()
        d += timedelta(days=1)
    for c in cmds:
        for key, bucket in ((c["date"], days), (c["date"][:7], months)):
            row = bucket.setdefault(key, blank())
            row["messages"] += 1
            if c["statut"] != "Annulé":
                if c["total"] > 0:
                    row["projets"] += 1
                row["montant"] += c["total"]
                row["avances"] += c["paid"]
    for p in pays:
        k = p["mode"] if p["mode"] in mlist else A_PRECISER
        for key, bucket in ((p["date"], days), (p["date"][:7], months)):
            row = bucket.setdefault(key, blank())
            row["enc"][k] += p["montant"]
            row["enc_total"] += p["montant"]
    for bucket in (days, months):
        for row in bucket.values():
            row["reste"] = row["montant"] - row["avances"]

    mois_opts = sorted(months.keys(), reverse=True)
    mois_sel = request.args.get("mois", "")
    day_rows = [(k, v) for k, v in sorted(days.items()) if not mois_sel or k.startswith(mois_sel)]
    totals = blank()
    for _, v in day_rows:
        for f in ("messages", "projets", "montant", "avances", "reste", "enc_total"):
            totals[f] += v[f]
        for m in v["enc"]:
            totals[m] = totals.get(m, 0) + v["enc"][m]
    month_rows = sorted(months.items())
    mtot = blank()
    for _, v in month_rows:
        for f in ("messages", "projets", "montant", "avances", "reste", "enc_total"):
            mtot[f] += v[f]
        for m in v["enc"]:
            mtot["enc"][m] += v["enc"][m]
    labels = [f"{fmt_date(k)[:5]}" for k, _ in day_rows]
    chart = {"labels": labels,
             "messages": [v["messages"] for _, v in day_rows], "projets": [v["projets"] for _, v in day_rows],
             "montant": [v["montant"] for _, v in day_rows], "avances": [v["avances"] for _, v in day_rows],
             "reste": [v["reste"] for _, v in day_rows]}
    enc_cols = [m for m in mlist + [A_PRECISER] if mtot["enc"][m]] or ["Espèces", "MVola", "Orange Money"]
    return render_template("mr.html", days=day_rows, months=month_rows, totals=totals, mtot=mtot,
                           mois_sel=mois_sel, mois_opts=mois_opts, mlist=mlist, enc_cols=enc_cols, chart=chart,
                           jour=lambda s: JOURS[parse_date(s).weekday()], mois_nom=lambda ym: f"{MOIS[int(ym[5:]) - 1]} {ym[:4]}")


# ---------------------------------------------------------------- commandes
@app.route("/commandes")
@login_required
def commandes():
    q = request.args.get("q", "").strip().lower()
    st = request.args.get("statut", "")
    sp = request.args.get("paiement", "")
    prix = request.args.get("prix") == "1"
    retard = request.args.get("retard") == "1"
    late_ids = {c["id"] for c in livraisons_a_surveiller(commandes_with_paid())["retard"]} if retard else set()
    page = max(1, int(request.args.get("page", 1) or 1))
    rows = commandes_with_paid("ORDER BY c.id DESC")
    out = []
    for c in rows:
        if st == "-" and c["statut"]:
            continue
        if st and st != "-" and c["statut"] != st:
            continue
        if sp == "-" and c["total"]:
            continue
        if sp and sp != "-" and c["sp"] != sp:
            continue
        if prix and not c["total"]:
            continue
        if retard and c["id"] not in late_ids:
            continue
        if q:
            hay = " ".join(str(c.get(k) or "") for k in ("code", "client", "tel", "produit", "situation", "remarques", "responsable")).lower()
            if q not in hay:
                continue
        out.append(c)
    sums = {"total": sum(c["total"] for c in out), "paid": sum(c["paid"] for c in out), "reste": sum(c["reste"] for c in out)}
    per = 50
    pages = max(1, (len(out) + per - 1) // per)
    page = min(page, pages)
    return render_template("commandes.html", rows=out[(page - 1) * per: page * per], count=len(out), sums=sums,
                           page=page, pages=pages, statuts=liste("statut_commande"), args=request.args)


def cmd_from_form(f):
    return {
        "date": f.get("date") or None, "client": f.get("client", "").strip(), "tel": f.get("tel", "").strip() or None,
        "situation": f.get("situation", "").strip() or None, "produit": f.get("produit", "").strip() or None,
        "qte": to_num(f.get("qte")), "pu": to_num(f.get("pu")), "statut": f.get("statut") or None,
        "livraison": f.get("livraison") or None, "date_livraison": f.get("date_livraison") or None,
        "recu_par": f.get("recu_par") or None, "responsable": f.get("responsable") or None,
        "equipe_msg": f.get("equipe_msg") or None, "canal": f.get("canal", "").strip() or None,
        "remarques": f.get("remarques", "").strip() or None,
    }


@app.route("/commandes/nouvelle", methods=["GET", "POST"])
@login_required
def commande_nouvelle():
    db = get_db()
    if request.method == "POST":
        data = cmd_from_form(request.form)
        if not data["client"]:
            flash("Le nom du client est obligatoire.", "bad")
        else:
            new_id = (db.execute("SELECT MAX(id) FROM commandes").fetchone()[0] or 0) + 1
            data.update(id=new_id, created_by=g.user["id"], created_at=now_iso())
            db.execute(f"INSERT INTO commandes({','.join(data)}) VALUES ({','.join('?' * len(data))})", list(data.values()))
            log("création", cmd_code(new_id), data["client"])
            db.commit()
            flash(f"{cmd_code(new_id)} créée.", "good")
            return redirect(url_for("commande", cid=new_id))
    c = dict.fromkeys(cmd_from_form({}), None)
    c.update(cmd_from_form(request.form) if request.method == "POST" else
             {"date": iso(today()), "statut": "En attente", "livraison": "Non livré"})
    return render_template("commande.html", c=c, new=True, **form_lists(c))


def form_lists(c):
    return {"statuts": liste("statut_commande"), "livraisons": liste("statut_livraison"), "canaux": liste("canal"),
            "team_recu": equipe_noms(c.get("recu_par")), "team_resp": equipe_noms(c.get("responsable")),
            "team_msg": equipe_noms(c.get("equipe_msg")), "mlist": modes(), "team": equipe_noms()}


@app.route("/commandes/<int:cid>", methods=["GET", "POST"])
@login_required
def commande(cid):
    db = get_db()
    rows = commandes_with_paid("WHERE c.id=?", (cid,))
    if not rows:
        abort(404)
    c = rows[0]
    if request.method == "POST":
        data = cmd_from_form(request.form)
        if not data["client"]:
            flash("Le nom du client est obligatoire.", "bad")
        else:
            changes = [k for k in data if (c.get(k) or None) != data[k]]
            data.update(updated_by=g.user["id"], updated_at=now_iso())
            db.execute(f"UPDATE commandes SET {','.join(k + '=?' for k in data)} WHERE id=?", list(data.values()) + [cid])
            log("modification", cmd_code(cid), ", ".join(changes))
            db.commit()
            flash("Commande enregistrée.", "good")
            return redirect(url_for("commande", cid=cid))
    pays = db.execute("""SELECT p.*, u.nom AS saisi_nom FROM paiements p LEFT JOIN users u ON u.id=p.saisi_par
                         WHERE commande_id=? ORDER BY date, id""", (cid,)).fetchall()
    who = {r["id"]: r["nom"] for r in db.execute("SELECT id, nom FROM users")}
    return render_template("commande.html", c=c, new=False, pays=pays, who=who, **form_lists(c))


# ---------------------------------------------------------------- paiements
def validate_payment(f, mlist, new=True):
    p = {"date": f.get("date") or "", "montant": to_num(f.get("montant")), "mode": f.get("mode", ""),
         "reference": f.get("reference", "").strip() or None, "encaisse_par": f.get("encaisse_par") or None,
         "remarque": f.get("remarque", "").strip() or None}
    if not parse_date(p["date"]):
        return p, "Indiquez la date du paiement."
    if new and not (p["montant"] and p["montant"] > 0):
        return p, "Indiquez un montant supérieur à 0."
    if p["mode"] not in mlist and not (not new and p["mode"] == A_PRECISER):
        return p, "Choisissez le mode de paiement."
    if not p["encaisse_par"]:
        return p, "Indiquez qui a reçu l'argent (Nandray ny vola)."
    if p["mode"] in MOBILE and not p["reference"]:
        return p, f"Indiquez la référence de la transaction {p['mode']}."
    return p, None


@app.route("/commandes/<int:cid>/paiement", methods=["POST"])
@login_required
def paiement_ajout(cid):
    db = get_db()
    if not db.execute("SELECT 1 FROM commandes WHERE id=?", (cid,)).fetchone():
        abort(404)
    p, err = validate_payment(request.form, modes())
    if err:
        flash(err, "bad")
        return redirect(url_for("commande", cid=cid) + "#paiement")
    db.execute("""INSERT INTO paiements(commande_id, date, montant, mode, reference, encaisse_par, remarque, saisi_par, saisi_le)
                  VALUES (?,?,?,?,?,?,?,?,?)""",
               (cid, p["date"], p["montant"], p["mode"], p["reference"], p["encaisse_par"], p["remarque"], g.user["id"], now_iso()))
    log("paiement", cmd_code(cid), f"{fmt_ar(p['montant'])} {p['mode']} encaissé par {p['encaisse_par']}")
    db.commit()
    flash(f"Paiement de {fmt_ar(p['montant'])} enregistré au nom de {g.user['nom']}.", "good")
    return redirect(request.form.get("back") or url_for("commande", cid=cid))


@app.route("/paiements")
@login_required
def paiements():
    db = get_db()
    a = request.args
    periode = a.get("periode", "all")
    d1, d2 = period_range(periode, a.get("du"), a.get("au"))
    mlist = modes()
    rows = db.execute("""SELECT p.*, c.client, (c.id IS NOT NULL) AS cmd_ok, u.nom AS saisi_nom, ua.nom AS annule_nom
                         FROM paiements p LEFT JOIN commandes c ON c.id=p.commande_id
                         LEFT JOIN users u ON u.id=p.saisi_par LEFT JOIN users ua ON ua.id=p.annule_par
                         WHERE p.date BETWEEN ? AND ? ORDER BY p.date DESC, p.id DESC""", (d1, d2)).fetchall()
    out = []
    q = a.get("q", "").strip().lower()
    for p in rows:
        if p["annule"] and a.get("annules") != "1":
            continue
        if a.get("mode"):
            if a["mode"] == A_PRECISER and p["mode"] in mlist:
                continue
            if a["mode"] != A_PRECISER and p["mode"] != a["mode"]:
                continue
        if a.get("enc") == "-" and p["encaisse_par"]:
            continue
        if a.get("enc") and a["enc"] != "-" and p["encaisse_par"] != a["enc"]:
            continue
        if a.get("saisi") == "-" and p["saisi_par"]:
            continue
        if a.get("saisi") and a["saisi"] != "-" and str(p["saisi_par"]) != a["saisi"]:
            continue
        issues = pay_issues(p, p["cmd_ok"])
        if a.get("verifier") == "1" and not issues:
            continue
        if q and q not in " ".join(str(x or "") for x in (cmd_code(p["commande_id"]), p["client"], p["reference"], p["remarque"], p["encaisse_par"])).lower():
            continue
        out.append((p, issues))
    valid = [p for p, _ in out if not p["annule"]]
    by_mode = {}
    for p in valid:
        k = p["mode"] if p["mode"] in mlist else A_PRECISER
        by_mode[k] = by_mode.get(k, 0) + p["montant"]
    users = db.execute("SELECT id, nom FROM users ORDER BY nom").fetchall()
    if a.get("format") == "csv":
        return csv_response("paiements", [["Date", "N° commande", "Client", "Montant (Ar)", "Mode", "Référence", "Encaissé par",
                                           "Saisi par", "Saisi le", "Annulé", "Motif", "Remarque"]] +
                            [[fmt_date(p["date"]), cmd_code(p["commande_id"]), p["client"], p["montant"], p["mode"], p["reference"],
                              p["encaisse_par"], p["saisi_nom"] or "Import Excel", p["saisi_le"], "Oui" if p["annule"] else "",
                              p["motif_annulation"], p["remarque"]] for p, _ in out])
    return render_template("paiements.html", rows=out, total=sum(p["montant"] for p in valid), by_mode=by_mode,
                           mlist=mlist + [A_PRECISER], team=equipe_noms(), users=users, args=a, periode=periode,
                           label=period_label(periode, d1, d2))


@app.route("/paiements/nouveau", methods=["GET", "POST"])
@login_required
def paiement_nouveau():
    db = get_db()
    cmds = db.execute("SELECT id, client FROM commandes ORDER BY id DESC").fetchall()
    if request.method == "POST":
        raw = request.form.get("commande", "")
        import re
        m = re.search(r"(\d+)", raw)
        cid = int(m.group(1)) if m else None
        if not cid or not db.execute("SELECT 1 FROM commandes WHERE id=?", (cid,)).fetchone():
            flash("Choisissez une commande existante (ex. CMD-012).", "bad")
        else:
            return paiement_ajout(cid)
    return render_template("paiement_nouveau.html", cmds=cmds, mlist=modes(), team=equipe_noms(), f=request.form)


@app.route("/paiements/<int:pid>", methods=["GET", "POST"])
@login_required
def paiement(pid):
    db = get_db()
    p = db.execute("""SELECT p.*, c.client, u.nom AS saisi_nom, um.nom AS modifie_nom, ua.nom AS annule_nom
                      FROM paiements p LEFT JOIN commandes c ON c.id=p.commande_id LEFT JOIN users u ON u.id=p.saisi_par
                      LEFT JOIN users um ON um.id=p.modifie_par LEFT JOIN users ua ON ua.id=p.annule_par
                      WHERE p.id=?""", (pid,)).fetchone()
    if not p:
        abort(404)
    if request.method == "POST":
        if p["annule"]:
            abort(400)
        data, err = validate_payment(request.form, modes(), new=False)
        if err:
            flash(err, "bad")
        else:
            db.execute("""UPDATE paiements SET date=?, mode=?, reference=?, encaisse_par=?, remarque=?, modifie_par=?, modifie_le=?
                          WHERE id=?""", (data["date"], data["mode"], data["reference"], data["encaisse_par"], data["remarque"],
                                          g.user["id"], now_iso(), pid))
            log("modification paiement", f"P{pid} {cmd_code(p['commande_id'])}",
                f"mode {p['mode']}→{data['mode']}, encaissé par {p['encaisse_par']}→{data['encaisse_par']}")
            db.commit()
            flash("Paiement mis à jour.", "good")
            return redirect(url_for("paiement", pid=pid))
    return render_template("paiement.html", p=p, mlist=modes(), team=equipe_noms(p["encaisse_par"]))


@app.route("/paiements/<int:pid>/annuler", methods=["POST"])
@login_required
def paiement_annuler(pid):
    db = get_db()
    p = db.execute("SELECT * FROM paiements WHERE id=?", (pid,)).fetchone()
    if not p or p["annule"]:
        abort(404)
    motif = request.form.get("motif", "").strip()
    if not motif:
        flash("Le motif de l'annulation est obligatoire.", "bad")
        return redirect(url_for("paiement", pid=pid))
    db.execute("UPDATE paiements SET annule=1, annule_par=?, annule_le=?, motif_annulation=? WHERE id=?",
               (g.user["id"], now_iso(), motif, pid))
    log("annulation paiement", f"P{pid} {cmd_code(p['commande_id'])}", f"{fmt_ar(p['montant'])} — {motif}")
    db.commit()
    flash("Paiement annulé. Il reste visible dans l'historique.", "good")
    return redirect(url_for("paiement", pid=pid))


# ---------------------------------------------------------------- rendez-vous
@app.route("/rendez-vous")
@login_required
def rdv_list():
    rows = get_db().execute("SELECT * FROM rdv ORDER BY COALESCE(date,'') DESC, id DESC").fetchall()
    return render_template("rdv_list.html", rows=rows)


@app.route("/rendez-vous/nouveau", methods=["GET", "POST"])
@app.route("/rendez-vous/<int:rid>", methods=["GET", "POST"])
@login_required
def rdv_edit(rid=None):
    db = get_db()
    r = db.execute("SELECT * FROM rdv WHERE id=?", (rid,)).fetchone() if rid else None
    if rid and not r:
        abort(404)
    if request.method == "POST":
        if request.form.get("supprimer") and rid:
            db.execute("DELETE FROM rdv WHERE id=?", (rid,))
            log("suppression", f"RDV {rid}", r["projet"])
            db.commit()
            flash("Rendez-vous supprimé.", "good")
            return redirect(url_for("rdv_list"))
        f = request.form
        data = {k: (f.get(k, "").strip() or None) for k in ("projet", "tel", "lieu", "frais", "date", "heure", "responsable", "statut", "notes")}
        if not data["projet"]:
            flash("Indiquez le projet.", "bad")
        else:
            data.update(updated_by=g.user["id"], updated_at=now_iso())
            if rid:
                db.execute(f"UPDATE rdv SET {','.join(k + '=?' for k in data)} WHERE id=?", list(data.values()) + [rid])
            else:
                db.execute(f"INSERT INTO rdv({','.join(data)}) VALUES ({','.join('?' * len(data))})", list(data.values()))
            db.commit()
            flash("Rendez-vous enregistré.", "good")
            return redirect(url_for("rdv_list"))
    return render_template("rdv_edit.html", r=r or {}, team=equipe_noms(r["responsable"] if r else None))


# ---------------------------------------------------------------- équipe, listes, utilisateurs
@app.route("/equipe", methods=["GET", "POST"])
@login_required
def equipe():
    db = get_db()
    if request.method == "POST":
        nom = request.form.get("nom", "").strip()
        if not nom:
            flash("Indiquez un nom.", "bad")
        elif db.execute("SELECT 1 FROM equipe WHERE lower(nom)=lower(?)", (nom,)).fetchone():
            flash("Ce nom existe déjà.", "bad")
        else:
            db.execute("INSERT INTO equipe(nom, role, tel) VALUES (?,?,?)",
                       (nom, request.form.get("role", "").strip() or None, request.form.get("tel", "").strip() or None))
            log("ajout membre", nom)
            db.commit()
            flash(f"{nom} ajouté(e).", "good")
        return redirect(url_for("equipe"))
    rows = db.execute("""SELECT e.*, COALESCE((SELECT SUM(montant) FROM paiements WHERE annule=0 AND encaisse_par=e.nom),0) AS enc,
                                (SELECT COUNT(*) FROM commandes WHERE responsable=e.nom) AS ncmd
                         FROM equipe e ORDER BY e.actif DESC, e.id""").fetchall()
    return render_template("equipe.html", rows=rows)


@app.route("/equipe/<int:eid>/actif", methods=["POST"])
@login_required
def equipe_toggle(eid):
    db = get_db()
    db.execute("UPDATE equipe SET actif = 1 - actif WHERE id=?", (eid,))
    db.commit()
    return redirect(url_for("equipe"))


@app.route("/listes", methods=["GET", "POST"])
@login_required
def listes():
    db = get_db()
    if request.method == "POST":
        if g.user["role"] != "admin":
            abort(403)
        cat, val = request.form.get("categorie"), request.form.get("valeur", "").strip()
        if cat in LISTE_LABELS and val:
            n = db.execute("SELECT COALESCE(MAX(ordre),0)+1 FROM listes WHERE categorie=?", (cat,)).fetchone()[0]
            try:
                db.execute("INSERT INTO listes(categorie, valeur, ordre) VALUES (?,?,?)", (cat, val, n))
                db.commit()
                flash(f"« {val} » ajouté.", "good")
            except sqlite3.IntegrityError:
                flash("Cette valeur existe déjà.", "bad")
        return redirect(url_for("listes"))
    data = {cat: db.execute("SELECT * FROM listes WHERE categorie=? ORDER BY ordre, valeur", (cat,)).fetchall() for cat in LISTE_LABELS}
    usage = {
        "statut_commande": dict(db.execute("SELECT statut, COUNT(*) FROM commandes GROUP BY statut").fetchall()),
        "statut_livraison": dict(db.execute("SELECT livraison, COUNT(*) FROM commandes GROUP BY livraison").fetchall()),
        "mode_paiement": dict(db.execute("SELECT mode, COUNT(*) FROM paiements WHERE annule=0 GROUP BY mode").fetchall()),
        "canal": dict(db.execute("SELECT canal, COUNT(*) FROM commandes GROUP BY canal").fetchall()),
    }
    return render_template("listes.html", data=data, labels=LISTE_LABELS, usage=usage)


@app.route("/listes/<int:lid>/supprimer", methods=["POST"])
@admin_required
def liste_supprimer(lid):
    db = get_db()
    db.execute("DELETE FROM listes WHERE id=?", (lid,))
    db.commit()
    flash("Valeur retirée de la liste (les données existantes ne changent pas).", "good")
    return redirect(url_for("listes"))


@app.route("/listes/<int:lid>/monter", methods=["POST"])
@admin_required
def liste_monter(lid):
    db = get_db()
    r = db.execute("SELECT * FROM listes WHERE id=?", (lid,)).fetchone()
    if r:
        rows = db.execute("SELECT id FROM listes WHERE categorie=? ORDER BY ordre, valeur", (r["categorie"],)).fetchall()
        ids = [x["id"] for x in rows]
        i = ids.index(lid)
        if i > 0:
            ids[i - 1], ids[i] = ids[i], ids[i - 1]
            for n, x in enumerate(ids):
                db.execute("UPDATE listes SET ordre=? WHERE id=?", (n, x))
            db.commit()
    return redirect(url_for("listes"))


@app.route("/utilisateurs", methods=["GET", "POST"])
@admin_required
def utilisateurs():
    db = get_db()
    if request.method == "POST":
        u, nom, pw, role = (request.form.get(k, "").strip() for k in ("username", "nom", "password", "role"))
        if not u or not nom or len(pw) < 6:
            flash("Identifiant, nom et mot de passe (6 caractères minimum) sont obligatoires.", "bad")
        elif db.execute("SELECT 1 FROM users WHERE username=?", (u.lower(),)).fetchone():
            flash("Cet identifiant existe déjà.", "bad")
        else:
            db.execute("INSERT INTO users(username, nom, password_hash, role, created_at) VALUES (?,?,?,?,?)",
                       (u.lower(), nom, generate_password_hash(pw), "admin" if role == "admin" else "membre", now_iso()))
            log("création utilisateur", u.lower(), nom)
            db.commit()
            flash(f"Compte « {u.lower()} » créé pour {nom}.", "good")
        return redirect(url_for("utilisateurs"))
    rows = db.execute("""SELECT u.*, (SELECT COUNT(*) FROM paiements WHERE saisi_par=u.id) AS npay,
                                COALESCE((SELECT SUM(montant) FROM paiements WHERE saisi_par=u.id AND annule=0),0) AS spay
                         FROM users u ORDER BY u.actif DESC, u.nom""").fetchall()
    return render_template("utilisateurs.html", rows=rows)


@app.route("/utilisateurs/<int:uid>", methods=["POST"])
@admin_required
def utilisateur_action(uid):
    db = get_db()
    action = request.form.get("action")
    if uid == g.user["id"] and action in ("desactiver", "role"):
        flash("Vous ne pouvez pas modifier votre propre accès.", "bad")
    elif action == "desactiver":
        db.execute("UPDATE users SET actif = 1 - actif WHERE id=?", (uid,))
    elif action == "role":
        db.execute("UPDATE users SET role = CASE role WHEN 'admin' THEN 'membre' ELSE 'admin' END WHERE id=?", (uid,))
    elif action == "mdp":
        pw = request.form.get("password", "")
        if len(pw) < 6:
            flash("Mot de passe : 6 caractères minimum.", "bad")
        else:
            db.execute("UPDATE users SET password_hash=? WHERE id=?", (generate_password_hash(pw), uid))
            flash("Mot de passe réinitialisé.", "good")
    db.commit()
    return redirect(url_for("utilisateurs"))


@app.route("/journal")
@admin_required
def journal():
    rows = get_db().execute("""SELECT j.*, u.nom FROM journal j LEFT JOIN users u ON u.id=j.user_id
                               ORDER BY j.id DESC LIMIT 500""").fetchall()
    return render_template("journal.html", rows=rows)


# ---------------------------------------------------------------- import / export
@app.route("/import", methods=["GET", "POST"])
@admin_required
def import_page():
    db = get_db()
    counts = {t: db.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0] for t in ("commandes", "paiements", "rdv", "equipe")}
    if request.method == "POST":
        f = request.files.get("fichier")
        if not f or not f.filename.lower().endswith(".xlsx"):
            flash("Choisissez un fichier Excel (.xlsx).", "bad")
            return redirect(url_for("import_page"))
        replace = request.form.get("remplacer") == "1"
        if counts["commandes"] and not replace:
            flash("La base contient déjà des commandes : cochez « Remplacer » pour réimporter.", "bad")
            return redirect(url_for("import_page"))
        from ampandrana_importer import import_workbook
        path = os.path.join(app.instance_path, "import.xlsx")
        f.save(path)
        try:
            res = import_workbook(db, path, replace=replace)
            log("import Excel", f.filename, str(res))
            db.commit()
            flash(f"Import terminé : {res['commandes']} commandes, {res['paiements']} paiements "
                  f"({fmt_ar(res['montant'])}), {res['equipe']} membres, {res['rdv']} rendez-vous.", "good")
        except Exception as e:  # noqa: BLE001
            db.rollback()
            flash(f"Import impossible : {e}", "bad")
        finally:
            os.remove(path)
        return redirect(url_for("import_page"))
    return render_template("import.html", counts=counts)


def csv_response(name, rows):
    buf = io.StringIO()
    buf.write("﻿")
    csv.writer(buf, delimiter=";").writerows(rows)
    data = io.BytesIO(buf.getvalue().encode("utf-8"))
    return send_file(data, mimetype="text/csv", as_attachment=True, download_name=f"{name}_{iso(today())}.csv")


@app.route("/export.xlsx")
@login_required
def export_xlsx():
    from ampandrana_exporter import build_workbook
    data = build_workbook(get_db(), commandes_with_paid("ORDER BY c.id"), modes())
    return send_file(data, as_attachment=True, download_name=f"Ampandrana_export_{iso(today())}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.errorhandler(403)
def forbidden(_e):
    return render_template("message.html", titre="Accès refusé", texte="Cette page est réservée aux administrateurs."), 403


@app.errorhandler(404)
def notfound(_e):
    return render_template("message.html", titre="Introuvable", texte="Cette page ou cet élément n'existe pas."), 404


if __name__ == "__main__":
    app.run(debug=True)
