import os
import uuid
import secrets
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
import urllib.parse
import io
import psycopg2
import psycopg2.extras
import psycopg2.pool

# ---------------------------------------------------------------------------
# App setup
# ---------------------------------------------------------------------------
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", secrets.token_hex(32))

ADMIN_PASSWORD_HASH = os.environ.get(
    "ADMIN_PASSWORD_HASH",
    generate_password_hash("changeme")
)

DATABASE_URL = os.environ.get("DATABASE_URL", "")

# ---------------------------------------------------------------------------
# Connection pool (invece di aprire/chiudere una connessione per richiesta)
# minconn=1, maxconn=3: sufficiente per 1 worker gunicorn su istanza 512MB,
# evita di esaurire memoria/connessioni come con connessioni "usa e getta".
# connect_timeout evita che richieste restino appese se Supabase è lento.
# ---------------------------------------------------------------------------
db_pool = psycopg2.pool.ThreadedConnectionPool(
    minconn=1,
    maxconn=3,
    dsn=DATABASE_URL,
    sslmode="require",
    connect_timeout=10
)


def get_db():
    if "db" not in g:
        g.db = db_pool.getconn()
    return g.db


@app.teardown_appcontext
def close_db(exc=None):
    db = g.pop("db", None)
    if db is not None:
        if exc:
            db.rollback()
        db_pool.putconn(db)


def query(sql, params=(), fetchone=False, fetchall=False, commit=False):
    db = get_db()
    with db.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
        cur.execute(sql, params)
        result = None
        if fetchone:
            result = cur.fetchone()
        elif fetchall:
            result = cur.fetchall()
        if commit:
            db.commit()
        return result


def execute(sql, params=(), commit=False):
    db = get_db()
    with db.cursor() as cur:
        cur.execute(sql, params)
        if commit:
            db.commit()


def execute_returning(sql, params=()):
    db = get_db()
    with db.cursor() as cur:
        cur.execute(sql, params)
        result = cur.fetchone()
        db.commit()
        return result


def init_db():
    conn = db_pool.getconn()
    try:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS giri (
                    id          SERIAL PRIMARY KEY,
                    numero_giro TEXT    NOT NULL,
                    data_giro   DATE    NOT NULL,
                    token       TEXT    NOT NULL UNIQUE,
                    autista     TEXT,
                    telefono    TEXT,
                    completato  INTEGER NOT NULL DEFAULT 0,
                    created_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );

                CREATE TABLE IF NOT EXISTS consegne (
                    id             SERIAL PRIMARY KEY,
                    giro_id        INTEGER NOT NULL REFERENCES giri(id) ON DELETE CASCADE,
                    seq_originale  INTEGER NOT NULL,
                    seq_autista    INTEGER,
                    ragione_soc    TEXT    NOT NULL,
                    indirizzo      TEXT,
                    citta          TEXT,
                    provincia      TEXT,
                    codice_cliente TEXT    NOT NULL,
                    telefono_cl    TEXT,
                    note           TEXT
                );
            """)
            conn.commit()
    finally:
        db_pool.putconn(conn)


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
# XLS parser (SpreadsheetML formato AS400)
# ---------------------------------------------------------------------------
def parse_xls_as400(filepath):
    with open(filepath, encoding="iso-8859-15", errors="replace") as f:
        content = f.read()

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
        if len(cells) >= 9 and cells[6]:
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
    giri = query(
        """SELECT g.*, COUNT(c.id) as n_clienti
           FROM giri g
           LEFT JOIN consegne c ON c.giro_id = g.id
           WHERE g.data_giro = %s
           GROUP BY g.id
           ORDER BY g.numero_giro""",
        (date.today(),),
        fetchall=True
    )

    giri_con_link = []
    for g_row in (giri or []):
        giro = dict(g_row)
        link_giro = url_for("giro_autista", token=giro["token"], _external=True)
        tel = (giro.get("telefono") or "").strip().replace(" ", "").replace("-", "")
        if tel and not tel.startswith("+"):
            tel = "+39" + tel.lstrip("0")
        testo = (
            f"Ciao! Ecco il link per ordinare le consegne di oggi "
            f"(Giro {giro['numero_giro']}): {link_giro}"
        )
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

        oggi = date.today()
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

            existing = query(
                "SELECT id FROM giri WHERE numero_giro=%s AND data_giro=%s",
                (numero_giro, oggi),
                fetchone=True
            )
            if existing:
                execute("DELETE FROM giri WHERE id=%s", (existing["id"],), commit=True)

            token = secrets.token_urlsafe(16)
            row = execute_returning(
                "INSERT INTO giri (numero_giro, data_giro, token) VALUES (%s,%s,%s) RETURNING id",
                (numero_giro, oggi, token)
            )
            giro_id = row[0]

            db = get_db()
            with db.cursor() as cur:
                for r in righe:
                    cur.execute(
                        """INSERT INTO consegne
                           (giro_id, seq_originale, ragione_soc, indirizzo, citta,
                            provincia, codice_cliente, telefono_cl, note)
                           VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
                        (giro_id, r["seq_originale"], r["ragione_soc"],
                         r["indirizzo"], r["citta"], r["provincia"],
                         r["codice_cliente"], r["telefono_cl"], r["note"])
                    )
                db.commit()
            nuovi += 1

        if nuovi:
            flash(f"{nuovi} giro/i caricato/i con successo.", "success")
        for e in errori:
            flash(e, "danger")

        return redirect(url_for("admin_dashboard"))

    return render_template("upload.html")


# ---------------------------------------------------------------------------
# Admin: imposta autista e telefono
# ---------------------------------------------------------------------------
@app.route("/admin/giro/<int:giro_id>/autista", methods=["POST"])
@login_required
def set_autista(giro_id):
    autista = request.form.get("autista", "").strip()
    telefono = request.form.get("telefono", "").strip()
    execute(
        "UPDATE giri SET autista=%s, telefono=%s WHERE id=%s",
        (autista, telefono, giro_id),
        commit=True
    )
    return redirect(url_for("admin_dashboard"))


# ---------------------------------------------------------------------------
# Admin: export CSV singolo giro
# ---------------------------------------------------------------------------
@app.route("/admin/giro/<int:giro_id>/export")
@login_required
def export_giro(giro_id):
    giro = query("SELECT * FROM giri WHERE id=%s", (giro_id,), fetchone=True)
    if giro is None:
        abort(404)

    consegne = query(
        """SELECT COALESCE(seq_autista, seq_originale) as seq,
                  codice_cliente, ragione_soc
           FROM consegne WHERE giro_id=%s
           ORDER BY COALESCE(seq_autista, seq_originale)""",
        (giro_id,),
        fetchall=True
    )

    output = io.StringIO()
    writer = csv.writer(output, delimiter=";")
    writer.writerow(["GIRO", "SEQ_AUTISTA", "CODICE_CLIENTE", "RAGIONE_SOCIALE"])
    for c in (consegne or []):
        writer.writerow([giro["numero_giro"], c["seq"], c["codice_cliente"], c["ragione_soc"]])

    output.seek(0)
    oggi_fname = date.today().strftime("%Y%m%d")
    filename = f"sequenza_giro{giro['numero_giro']}_{oggi_fname}.csv"
    return send_file(
        io.BytesIO(output.getvalue().encode("utf-8-sig")),
        mimetype="text/csv",
        as_attachment=True,
        download_name=filename
    )


# ---------------------------------------------------------------------------
# Admin: export CSV tutti i giri confermati
# ---------------------------------------------------------------------------
@app.route("/admin/export")
@login_required
def admin_export():
    oggi = date.today()
    oggi_fname = oggi.strftime("%Y%m%d")
    giri = query(
        "SELECT id, numero_giro FROM giri WHERE data_giro=%s AND completato=1 ORDER BY numero_giro",
        (oggi,),
        fetchall=True
    )

    output = io.StringIO()
    writer = csv.writer(output, delimiter=";")
    writer.writerow(["GIRO", "SEQ_AUTISTA", "CODICE_CLIENTE", "RAGIONE_SOCIALE"])

    numeri = []
    for giro in (giri or []):
        numeri.append(giro["numero_giro"])
        consegne = query(
            """SELECT seq_autista, codice_cliente, ragione_soc
               FROM consegne WHERE giro_id=%s AND seq_autista IS NOT NULL
               ORDER BY seq_autista""",
            (giro["id"],),
            fetchall=True
        )
        for c in (consegne or []):
            writer.writerow([giro["numero_giro"], c["seq_autista"],
                             c["codice_cliente"], c["ragione_soc"]])

    output.seek(0)
    giri_str = "-".join(numeri) if numeri else "nessuno"
    filename = f"sequenze_giri{giri_str}_{oggi_fname}.csv"
    return send_file(
        io.BytesIO(output.getvalue().encode("utf-8-sig")),
        mimetype="text/csv",
        as_attachment=True,
        download_name=filename
    )


# ---------------------------------------------------------------------------
# Admin: export tutti i giri (anche incompleti)
# ---------------------------------------------------------------------------
@app.route("/admin/export_all")
@login_required
def admin_export_all():
    oggi = date.today()
    oggi_fname = oggi.strftime("%Y%m%d")
    giri = query(
        "SELECT id, numero_giro FROM giri WHERE data_giro=%s ORDER BY numero_giro",
        (oggi,),
        fetchall=True
    )

    output = io.StringIO()
    writer = csv.writer(output, delimiter=";")
    writer.writerow(["GIRO", "SEQ", "CODICE_CLIENTE", "RAGIONE_SOCIALE", "CONFERMATO"])

    numeri = []
    for giro in (giri or []):
        numeri.append(giro["numero_giro"])
        consegne = query(
            """SELECT COALESCE(seq_autista, seq_originale) as seq,
                      codice_cliente, ragione_soc,
                      CASE WHEN seq_autista IS NOT NULL THEN 'SI' ELSE 'NO' END as confermato
               FROM consegne WHERE giro_id=%s
               ORDER BY COALESCE(seq_autista, seq_originale)""",
            (giro["id"],),
            fetchall=True
        )
        for c in (consegne or []):
            writer.writerow([giro["numero_giro"], c["seq"],
                             c["codice_cliente"], c["ragione_soc"], c["confermato"]])

    output.seek(0)
    giri_str = "-".join(numeri) if numeri else "tutti"
    filename = f"sequenze_completo_giri{giri_str}_{oggi_fname}.csv"
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
    giro = query("SELECT * FROM giri WHERE token=%s", (token,), fetchone=True)
    if giro is None:
        abort(404)

    consegne = query(
        "SELECT * FROM consegne WHERE giro_id=%s ORDER BY seq_originale",
        (giro["id"],),
        fetchall=True
    )

    return render_template(
        "giro.html",
        giro=dict(giro),
        consegne=[dict(c) for c in consegne]
    )


# ---------------------------------------------------------------------------
# API: salva sequenza autista
# ---------------------------------------------------------------------------
@app.route("/api/giro/<token>/salva", methods=["POST"])
def salva_sequenza(token):
    giro = query("SELECT * FROM giri WHERE token=%s", (token,), fetchone=True)
    if giro is None:
        return jsonify({"ok": False, "msg": "Giro non trovato"}), 404
    if giro["completato"]:
        return jsonify({"ok": False, "msg": "Giro già confermato"}), 400

    data = request.get_json(force=True)
    ordine = data.get("ordine", [])

    if not ordine:
        return jsonify({"ok": False, "msg": "Ordine vuoto"}), 400

    consegne = query(
        "SELECT codice_cliente FROM consegne WHERE giro_id=%s", (giro["id"],), fetchall=True
    )
    codici_attesi = {c["codice_cliente"] for c in consegne}
    if codici_attesi != set(ordine):
        return jsonify({"ok": False, "msg": "Lista clienti non corrisponde"}), 400

    db = get_db()
    with db.cursor() as cur:
        for seq, codice in enumerate(ordine, start=1):
            cur.execute(
                "UPDATE consegne SET seq_autista=%s WHERE giro_id=%s AND codice_cliente=%s",
                (seq, giro["id"], codice)
            )
        cur.execute("UPDATE giri SET completato=1 WHERE id=%s", (giro["id"],))
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
