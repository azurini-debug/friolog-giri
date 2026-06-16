import os
import uuid
import secrets
import sqlite3
import csv
import xml.etree.ElementTree as ET
from datetime import datetime, date
from functools import wraps
from flask import (
    Flask, render_template, request, redirect, url_for,
    flash, session, send_file, jsonify, abort, g
)
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
import io

# ---------------------------------------------------------------------------
# App setup
# ---------------------------------------------------------------------------
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", secrets.token_hex(32))

ADMIN_PASSWORD_HASH = os.environ.get(
    "ADMIN_PASSWORD_HASH",
    generate_password_hash("changeme")   # Cambiare in produzione via env var
)

ALLOWED_EXTENSIONS = {"xls"}
DATABASE = os.path.join(app.instance_path, "friolog.db")
os.makedirs(app.instance_path, exist_ok=True)


# ---------------------------------------------------------------------------
# Database helpers
# ---------------------------------------------------------------------------
def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DATABASE, detect_types=sqlite3.PARSE_DECLTYPES)
        g.db.row_factory = sqlite3.Row
    return g.db


@app.teardown_appcontext
def close_db(exc=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db():
    db = get_db()
    db.executescript("""
        CREATE TABLE IF NOT EXISTS giri (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            numero_giro TEXT    NOT NULL,
            data_giro   DATE    NOT NULL,
            token       TEXT    NOT NULL UNIQUE,
            autista     TEXT,
            telefono    TEXT,
            completato  INTEGER NOT NULL DEFAULT 0,
            created_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS consegne (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            giro_id       INTEGER NOT NULL REFERENCES giri(id) ON DELETE CASCADE,
            seq_originale INTEGER NOT NULL,
            seq_autista   INTEGER,
            ragione_soc   TEXT    NOT NULL,
            indirizzo     TEXT,
            citta         TEXT,
            provincia     TEXT,
            codice_cliente TEXT   NOT NULL,
            telefono_cl   TEXT,
            note          TEXT
        );
    """)
    db.commit()


with app.app_context():
    init_db()


# ---------------------------------------------------------------------------
# Auth helper
# ---------------------------------------------------------------------------
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("admin_logged_in"):
            return redirect(url_for("admin_login"))
        return f(*args, **kwargs)
    return decorated


# ---------------------------------------------------------------------------
# XLS parser (SpreadsheetML format generato da AS400)
# ---------------------------------------------------------------------------
def parse_xls_as400(filepath):
    """
    Colonne attese (0-based):
      0  seq_originale
      1  ragione_soc
      2  indirizzo
      3  citta
      4  provincia
      5  n_colli (non usato nell'output)
      6  codice_cliente
      7  telefono_cl
      8  numero_giro
      9  note
    """
    with open(filepath, encoding="iso-8859-15", errors="replace") as f:
        content = f.read()

    # Rimuovi namespace per semplificare il parsing
    content = content.replace(' xmlns="urn:schemas-microsoft-com:office:spreadsheet"', "")
    for prefix in ("ss:", "x:", "o:", "html:"):
        content = content.replace(prefix, "")

    root = ET.fromstring(content.encode("utf-8"))
    worksheet = root.find(".//Worksheet")
    if worksheet is None:
        raise ValueError("Nessun foglio trovato nel file XLS")
    table = worksheet.find(".//Table")
    if table is None:
        raise ValueError("Nessuna tabella trovata nel foglio")

    rows = []
    for row_el in table.findall("Row"):
        cells = []
        for cell in row_el.findall("Cell"):
            data = cell.find("Data")
            cells.append(data.text.strip() if data is not None and data.text else "")
        if len(cells) >= 9 and cells[6]:   # codice_cliente obbligatorio
            rows.append({
                "seq_originale":  int(cells[0]) if cells[0].isdigit() else 0,
                "ragione_soc":    cells[1],
                "indirizzo":      cells[2],
                "citta":          cells[3],
                "provincia":      cells[4],
                "codice_cliente": cells[6],
                "telefono_cl":    cells[7],
                "numero_giro":    cells[8],
                "note":           cells[9] if len(cells) > 9 else "",
            })
    return rows


# ---------------------------------------------------------------------------
# Admin: login / logout
# ---------------------------------------------------------------------------
@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        pwd = request.form.get("password", "")
        if check_password_hash(ADMIN_PASSWORD_HASH, pwd):
            session["admin_logged_in"] = True
            return redirect(url_for("admin_dashboard"))
        flash("Password errata.", "danger")
    return render_template("login.html")


@app.route("/admin/logout")
def admin_logout():
    session.clear()
    return redirect(url_for("admin_login"))


# ---------------------------------------------------------------------------
# Admin: dashboard
# ---------------------------------------------------------------------------
@app.route("/admin")
@login_required
def admin_dashboard():
    db = get_db()
    giri = db.execute(
        """SELECT g.*, COUNT(c.id) as n_clienti
           FROM giri g
           LEFT JOIN consegne c ON c.giro_id = g.id
           WHERE g.data_giro = ?
           GROUP BY g.id
           ORDER BY g.numero_giro""",
        (date.today().isoformat(),)
    ).fetchall()

    # Genera link wa.me per ogni giro non completato
    giri_con_link = []
    for g_row in giri:
        giro = dict(g_row)
        link_giro = url_for("giro_autista", token=giro["token"], _external=True)
        tel = (giro.get("telefono") or "").strip().replace(" ", "").replace("-", "")
        if tel and not tel.startswith("+"):
            tel = "+39" + tel.lstrip("0")
        testo = (
            f"Ciao! Ecco il link per ordinare le consegne di oggi "
            f"(Giro {giro['numero_giro']}): {link_giro}"
        )
        import urllib.parse
        wa_link = f"https://wa.me/{tel}?text={urllib.parse.quote(testo)}" if tel else None
        giro["link_giro"] = link_giro
        giro["wa_link"] = wa_link
        giri_con_link.append(giro)

    return render_template("dashboard.html", giri=giri_con_link, oggi=date.today())


# ---------------------------------------------------------------------------
# Admin: caricamento file XLS
# ---------------------------------------------------------------------------
@app.route("/admin/upload", methods=["GET", "POST"])
@login_required
def admin_upload():
    if request.method == "POST":
        files = request.files.getlist("files")
        if not files or all(f.filename == "" for f in files):
            flash("Nessun file selezionato.", "warning")
            return redirect(request.url)

        db = get_db()
        oggi = date.today().isoformat()
        nuovi = 0
        errori = []

        for f in files:
            fname = secure_filename(f.filename)
            if not fname.lower().endswith(".xls"):
                errori.append(f"{fname}: formato non supportato (solo .xls)")
                continue

            tmp_path = os.path.join("/tmp", fname)
            f.save(tmp_path)

            try:
                righe = parse_xls_as400(tmp_path)
            except Exception as e:
                errori.append(f"{fname}: errore parsing — {e}")
                os.remove(tmp_path)
                continue
            finally:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)

            if not righe:
                errori.append(f"{fname}: nessuna riga valida trovata")
                continue

            numero_giro = righe[0]["numero_giro"]

            # Elimina giro esistente per oggi (ricaricamento)
            existing = db.execute(
                "SELECT id FROM giri WHERE numero_giro=? AND data_giro=?",
                (numero_giro, oggi)
            ).fetchone()
            if existing:
                db.execute("DELETE FROM giri WHERE id=?", (existing["id"],))

            token = secrets.token_urlsafe(16)
            cur = db.execute(
                "INSERT INTO giri (numero_giro, data_giro, token) VALUES (?,?,?)",
                (numero_giro, oggi, token)
            )
            giro_id = cur.lastrowid

            for r in righe:
                db.execute(
                    """INSERT INTO consegne
                       (giro_id, seq_originale, ragione_soc, indirizzo, citta,
                        provincia, codice_cliente, telefono_cl, note)
                       VALUES (?,?,?,?,?,?,?,?,?)""",
                    (giro_id, r["seq_originale"], r["ragione_soc"],
                     r["indirizzo"], r["citta"], r["provincia"],
                     r["codice_cliente"], r["telefono_cl"], r["note"])
                )
            nuovi += 1

        db.commit()

        if nuovi:
            flash(f"{nuovi} giro/i caricato/i con successo.", "success")
        for e in errori:
            flash(e, "danger")

        return redirect(url_for("admin_dashboard"))

    return render_template("upload.html")


# ---------------------------------------------------------------------------
# Admin: imposta autista e telefono per un giro
# ---------------------------------------------------------------------------
@app.route("/admin/giro/<int:giro_id>/autista", methods=["POST"])
@login_required
def set_autista(giro_id):
    autista = request.form.get("autista", "").strip()
    telefono = request.form.get("telefono", "").strip()
    db = get_db()
    db.execute(
        "UPDATE giri SET autista=?, telefono=? WHERE id=?",
        (autista, telefono, giro_id)
    )
    db.commit()
    return redirect(url_for("admin_dashboard"))


# ---------------------------------------------------------------------------
# Admin: download CSV output per AS400
# ---------------------------------------------------------------------------
@app.route("/admin/export")
@login_required
def admin_export():
    db = get_db()
    oggi = date.today().isoformat()
    giri = db.execute(
        "SELECT id, numero_giro FROM giri WHERE data_giro=? AND completato=1",
        (oggi,)
    ).fetchall()

    output = io.StringIO()
    writer = csv.writer(output, delimiter=";")
    writer.writerow(["GIRO", "SEQ_AUTISTA", "CODICE_CLIENTE", "RAGIONE_SOCIALE"])

    for giro in giri:
        consegne = db.execute(
            """SELECT seq_autista, codice_cliente, ragione_soc
               FROM consegne
               WHERE giro_id=? AND seq_autista IS NOT NULL
               ORDER BY seq_autista""",
            (giro["id"],)
        ).fetchall()
        for c in consegne:
            writer.writerow([
                giro["numero_giro"],
                c["seq_autista"],
                c["codice_cliente"],
                c["ragione_soc"]
            ])

    output.seek(0)
    filename = f"sequenze_{oggi}.csv"
    return send_file(
        io.BytesIO(output.getvalue().encode("utf-8-sig")),
        mimetype="text/csv",
        as_attachment=True,
        download_name=filename
    )


# ---------------------------------------------------------------------------
# Admin: export anche giri incompleti (con sequenza originale)
# ---------------------------------------------------------------------------
@app.route("/admin/export_all")
@login_required
def admin_export_all():
    db = get_db()
    oggi = date.today().isoformat()
    giri = db.execute(
        "SELECT id, numero_giro FROM giri WHERE data_giro=?", (oggi,)
    ).fetchall()

    output = io.StringIO()
    writer = csv.writer(output, delimiter=";")
    writer.writerow(["GIRO", "SEQ", "CODICE_CLIENTE", "RAGIONE_SOCIALE", "CONFERMATO"])

    for giro in giri:
        consegne = db.execute(
            """SELECT seq_autista, seq_originale, codice_cliente, ragione_soc
               FROM consegne WHERE giro_id=? ORDER BY COALESCE(seq_autista, seq_originale)""",
            (giro["id"],)
        ).fetchall()
        for c in consegne:
            seq = c["seq_autista"] if c["seq_autista"] is not None else c["seq_originale"]
            confermato = "SI" if c["seq_autista"] is not None else "NO"
            writer.writerow([giro["numero_giro"], seq, c["codice_cliente"],
                             c["ragione_soc"], confermato])

    output.seek(0)
    filename = f"sequenze_completo_{oggi}.csv"
    return send_file(
        io.BytesIO(output.getvalue().encode("utf-8-sig")),
        mimetype="text/csv",
        as_attachment=True,
        download_name=filename
    )


# ---------------------------------------------------------------------------
# Pagina autista (pubblica, protetta da token)
# ---------------------------------------------------------------------------
@app.route("/giro/<token>")
def giro_autista(token):
    db = get_db()
    giro = db.execute(
        "SELECT * FROM giri WHERE token=?", (token,)
    ).fetchone()
    if giro is None:
        abort(404)

    consegne = db.execute(
        """SELECT * FROM consegne WHERE giro_id=? ORDER BY seq_originale""",
        (giro["id"],)
    ).fetchall()

    return render_template(
        "giro.html",
        giro=dict(giro),
        consegne=[dict(c) for c in consegne]
    )


# ---------------------------------------------------------------------------
# API: salva sequenza dell'autista
# ---------------------------------------------------------------------------
@app.route("/api/giro/<token>/salva", methods=["POST"])
def salva_sequenza(token):
    db = get_db()
    giro = db.execute(
        "SELECT * FROM giri WHERE token=?", (token,)
    ).fetchone()
    if giro is None:
        return jsonify({"ok": False, "msg": "Giro non trovato"}), 404
    if giro["completato"]:
        return jsonify({"ok": False, "msg": "Giro già confermato"}), 400

    data = request.get_json(force=True)
    ordine = data.get("ordine", [])  # lista di codice_cliente in ordine

    if not ordine:
        return jsonify({"ok": False, "msg": "Ordine vuoto"}), 400

    # Verifica che tutti i clienti del giro siano presenti
    consegne = db.execute(
        "SELECT codice_cliente FROM consegne WHERE giro_id=?", (giro["id"],)
    ).fetchall()
    codici_attesi = {c["codice_cliente"] for c in consegne}
    codici_ricevuti = set(ordine)
    if codici_attesi != codici_ricevuti:
        return jsonify({"ok": False, "msg": "Lista clienti non corrisponde"}), 400

    for seq, codice in enumerate(ordine, start=1):
        db.execute(
            """UPDATE consegne SET seq_autista=?
               WHERE giro_id=? AND codice_cliente=?""",
            (seq, giro["id"], codice)
        )
    db.execute("UPDATE giri SET completato=1 WHERE id=?", (giro["id"],))
    db.commit()

    return jsonify({"ok": True, "msg": "Sequenza salvata con successo!"})


# ---------------------------------------------------------------------------
# Health check
# ---------------------------------------------------------------------------
@app.route("/health")
def health():
    return "ok", 200


if __name__ == "__main__":
    app.run(debug=False, host="0.0.0.0", port=5000)
