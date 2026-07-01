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

# Validità a tempo del link dell'autista: scade dopo N ore dalla generazione
# (oltre a scadere alla conferma). Cambia qui il numero per regolarla.
ORE_VALIDITA_GIRO = 12

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

                CREATE TABLE IF NOT EXISTS articoli (
                    id              SERIAL PRIMARY KEY,
                    codice          TEXT NOT NULL,
                    descrizione     TEXT,
                    data_scadenza   TEXT,
                    lotto           TEXT,
                    data_produzione TEXT,
                    fornitore       TEXT
                );
                CREATE INDEX IF NOT EXISTS idx_articoli_scad  ON articoli (data_scadenza);
                CREATE INDEX IF NOT EXISTS idx_articoli_lotto ON articoli (lotto);

                CREATE TABLE IF NOT EXISTS articoli_import (
                    id          SERIAL PRIMARY KEY,
                    filename    TEXT,
                    righe       INTEGER,
                    n_articoli  INTEGER,
                    n_terne     INTEGER,
                    created_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );

                CREATE TABLE IF NOT EXISTS riconoscimenti (
                    id                SERIAL PRIMARY KEY,
                    letto_lotto       TEXT,
                    letto_scadenza    TEXT,
                    letto_fornitore   TEXT,
                    letto_descrizione TEXT,
                    codice_proposto   TEXT,
                    confidenza        INTEGER,
                    n_foto            INTEGER,
                    esito             TEXT,
                    codice_corretto   TEXT,
                    created_at        TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );

                ALTER TABLE riconoscimenti ADD COLUMN IF NOT EXISTS esito           TEXT;
                ALTER TABLE riconoscimenti ADD COLUMN IF NOT EXISTS codice_corretto TEXT;
                ALTER TABLE riconoscimenti ADD COLUMN IF NOT EXISTS letto_ean       TEXT;

                CREATE TABLE IF NOT EXISTS correzioni (
                    id         SERIAL PRIMARY KEY,
                    codice     TEXT NOT NULL,
                    descr_norm TEXT,
                    ean        TEXT,
                    conferme   INTEGER NOT NULL DEFAULT 1,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );

                CREATE TABLE IF NOT EXISTS scanner_token (
                    id         SERIAL PRIMARY KEY,
                    token      TEXT NOT NULL UNIQUE,
                    nome       TEXT,
                    telefono   TEXT,
                    scadenza   TIMESTAMP NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
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
            f"(Giro {giro['numero_giro']}), valido 12 ore: {link_giro}"
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
# Admin: elimina giro
# ---------------------------------------------------------------------------
@app.route("/admin/giro/<int:giro_id>/elimina", methods=["POST"])
@login_required
def elimina_giro(giro_id):
    giro = query("SELECT numero_giro FROM giri WHERE id=%s", (giro_id,), fetchone=True)
    if giro is None:
        flash("Giro non trovato.", "danger")
        return redirect(url_for("admin_dashboard"))

    # Le consegne collegate vengono rimosse automaticamente (ON DELETE CASCADE)
    execute("DELETE FROM giri WHERE id=%s", (giro_id,), commit=True)
    flash(f"Giro {giro['numero_giro']} eliminato.", "success")
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
    giro = query(
        f"SELECT *, (created_at < NOW() - INTERVAL '{ORE_VALIDITA_GIRO} hours') AS scaduto "
        "FROM giri WHERE token=%s",
        (token,), fetchone=True
    )
    if giro is None:
        abort(404)

    if giro["completato"]:
        # Link monouso: una volta confermato il giro, il link non è più operativo.
        return render_template("giro.html", giro=dict(giro), consegne=[])

    if giro["scaduto"]:
        # Scadenza di sicurezza: il link non è più valido dopo 12 ore.
        return render_template("giro_scaduto.html"), 410

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
    giro = query(
        f"SELECT *, (created_at < NOW() - INTERVAL '{ORE_VALIDITA_GIRO} hours') AS scaduto "
        "FROM giri WHERE token=%s",
        (token,), fetchone=True
    )
    if giro is None:
        return jsonify({"ok": False, "msg": "Giro non trovato"}), 404
    if giro["completato"]:
        return jsonify({"ok": False, "msg": "Giro già confermato"}), 400
    if giro["scaduto"]:
        return jsonify({"ok": False, "msg": f"Link scaduto (oltre {ORE_VALIDITA_GIRO} ore). Chiedi un nuovo link all'ufficio."}), 410

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
        # Check-and-set atomico: solo il PRIMO invio "spegne" il link.
        # Se un secondo invio (doppio tap, refresh, corsa) arriva dopo,
        # completato è già 1 e rowcount vale 0 -> viene rifiutato.
        cur.execute(
            "UPDATE giri SET completato=1 "
            f"WHERE id=%s AND completato=0 AND created_at >= NOW() - INTERVAL '{ORE_VALIDITA_GIRO} hours'",
            (giro["id"],)
        )
        if cur.rowcount == 0:
            db.rollback()
            return jsonify({"ok": False, "msg": "Giro già confermato o link scaduto"}), 400
        for seq, codice in enumerate(ordine, start=1):
            cur.execute(
                "UPDATE consegne SET seq_autista=%s WHERE giro_id=%s AND codice_cliente=%s",
                (seq, giro["id"], codice)
            )
        db.commit()

    return jsonify({"ok": True, "msg": "Sequenza salvata con successo!"})


# ---------------------------------------------------------------------------
# Cartoni: parser del CSV movimenti (export gestionale)
# ---------------------------------------------------------------------------
def _clean_cell(x):
    """Rimuove l'involucro stile Excel  ="..."  e gli spazi."""
    if x is None:
        return ""
    x = x.strip()
    if x.startswith('="') and x.endswith('"'):
        x = x[2:-1]
    elif x.startswith('=') and len(x) > 1:
        x = x[1:]
    return x.strip().strip('"')


def _norm_data(x):
    """Tiene solo GG/MM/AAAA, scarta l'orario '00:00:00'."""
    return (_clean_cell(x) or "").split(" ")[0]


# Intestazioni accettate per ogni campo (confronto sul NOME colonna, non sulla
# posizione: il CSV può avere quante colonne vuoi, in qualsiasi ordine).
COLONNE_CARTONI = {
    "codice":          ["codice articolo", "cod articolo", "cod art", "codice art", "codice"],
    "descrizione":     ["descrizione articolo", "descr articolo", "descrizione art", "descrizione"],
    "data_scadenza":   ["data scadenza", "scadenza", "data scad"],
    "lotto":           ["lotto", "n lotto", "numero lotto", "nr lotto"],
    "data_produzione": ["data produzione", "data prod", "produzione"],
    "fornitore":       ["fornitore del pallet", "fornitore pallet"],
}

# Campi che DEVONO esserci, con il nome "ufficiale" da mostrare in caso di errore
CARTONI_OBBLIGATORIE = {
    "codice":        "Codice Articolo",
    "descrizione":   "Descrizione Articolo",
    "data_scadenza": "Data Scadenza",
    "lotto":         "Lotto",
}


def parse_movimenti_csv(filepath):
    import csv as _csv
    with open(filepath, encoding="latin1", errors="replace", newline="") as fh:
        reader = _csv.reader(fh, delimiter=";")
        try:
            header = next(reader)
        except StopIteration:
            raise ValueError("Il file è vuoto.")

        norm_header = [h.strip().lower() for h in header]
        idx = {}
        for chiave, alias in COLONNE_CARTONI.items():
            for a in alias:
                if a in norm_header:
                    idx[chiave] = norm_header.index(a)
                    break

        mancanti = [CARTONI_OBBLIGATORIE[c] for c in CARTONI_OBBLIGATORIE if c not in idx]
        if mancanti:
            raise ValueError(
                "Colonne obbligatorie non trovate: " + ", ".join(mancanti)
                + ". Controlla le intestazioni del CSV."
            )

        def get(row, k):
            i = idx.get(k)
            return row[i] if (i is not None and i < len(row)) else ""

        righe = []
        for row in reader:
            if not row:
                continue
            codice = _clean_cell(get(row, "codice"))
            if not codice:
                continue
            righe.append({
                "codice":          codice,
                "descrizione":     _clean_cell(get(row, "descrizione")),
                "data_scadenza":   _norm_data(get(row, "data_scadenza")),
                "lotto":           _clean_cell(get(row, "lotto")),
                "data_produzione": _norm_data(get(row, "data_produzione")),
                "fornitore":       _clean_cell(get(row, "fornitore")),
            })
        return righe


# ---------------------------------------------------------------------------
# Cartoni: sezione dashboard + caricamento CSV movimenti
# ---------------------------------------------------------------------------
@app.route("/admin/cartoni")
@login_required
def admin_cartoni():
    n_art = query("SELECT COUNT(*) AS n FROM articoli", fetchone=True)
    ultimo = query("SELECT * FROM articoli_import ORDER BY id DESC LIMIT 1", fetchone=True)
    tok_rows = query(
        "SELECT * FROM scanner_token WHERE scadenza > NOW() ORDER BY id DESC",
        fetchall=True,
    ) or []
    links = []
    for t in tok_rows:
        t = dict(t)
        url = url_for("scanner_operatore", token=t["token"], _external=True)
        tel = (t.get("telefono") or "").strip().replace(" ", "").replace("-", "")
        if tel and not tel.startswith("+"):
            tel = "+39" + tel.lstrip("0")
        testo = f"Ciao! Ecco il link per identificare i cartoni (valido 24h): {url}"
        t["url"] = url
        t["wa_link"] = (f"https://wa.me/{tel}?text={urllib.parse.quote(testo)}" if tel else None)
        links.append(t)
    return render_template(
        "cartoni.html",
        n_terne=(n_art["n"] if n_art else 0),
        ultimo=ultimo,
        links=links,
    )


@app.route("/admin/cartoni/upload-csv", methods=["POST"])
@login_required
def admin_cartoni_upload_csv():
    f = request.files.get("csv")
    if not f or f.filename == "":
        flash("Nessun file selezionato.", "warning")
        return redirect(url_for("admin_cartoni"))

    fname = secure_filename(f.filename)
    if not fname.lower().endswith(".csv"):
        flash("Formato non supportato: serve un file .csv", "danger")
        return redirect(url_for("admin_cartoni"))

    tmp_path = os.path.join("/tmp", fname)
    f.save(tmp_path)
    try:
        righe = parse_movimenti_csv(tmp_path)
    except Exception as e:
        flash(f"Errore nella lettura del CSV: {e}", "danger")
        return redirect(url_for("admin_cartoni"))
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

    if not righe:
        flash("Nessuna riga valida trovata nel CSV.", "warning")
        return redirect(url_for("admin_cartoni"))

    # Dedup sulle terne (codice, lotto, scadenza): più movimenti dello stesso
    # articolo/lotto non servono al riconoscimento. Conservo i fornitori
    # incontrati per quella terna (separati da ' | ') perché lo stesso articolo
    # può arrivare da fornitori/pallet diversi.
    agg = {}
    codici = set()
    for r in righe:
        codici.add(r["codice"])
        k = (r["codice"], r["lotto"], r["data_scadenza"])
        if k not in agg:
            agg[k] = {
                "codice": r["codice"], "descrizione": r["descrizione"],
                "data_scadenza": r["data_scadenza"], "lotto": r["lotto"],
                "data_produzione": r["data_produzione"], "fornitori": set(),
            }
        if r["fornitore"]:
            agg[k]["fornitori"].add(r["fornitore"])
        if r["data_produzione"] and not agg[k]["data_produzione"]:
            agg[k]["data_produzione"] = r["data_produzione"]

    terne = list(agg.values())

    db = get_db()
    with db.cursor() as cur:
        cur.execute("TRUNCATE articoli")
        psycopg2.extras.execute_values(
            cur,
            """INSERT INTO articoli
               (codice, descrizione, data_scadenza, lotto, data_produzione, fornitore)
               VALUES %s""",
            [(t["codice"], t["descrizione"], t["data_scadenza"], t["lotto"],
              t["data_produzione"], " | ".join(sorted(t["fornitori"])))
             for t in terne],
            page_size=1000,
        )
        cur.execute(
            """INSERT INTO articoli_import (filename, righe, n_articoli, n_terne)
               VALUES (%s,%s,%s,%s)""",
            (fname, len(righe), len(codici), len(terne)),
        )
        db.commit()

    flash(
        f"CSV caricato: {len(righe)} righe lette, {len(codici)} articoli, "
        f"{len(terne)} combinazioni lotto/scadenza in archivio.",
        "success",
    )
    return redirect(url_for("admin_cartoni"))


# ---------------------------------------------------------------------------
# Cartoni: lettura etichetta via API Claude (visione) + motore di match
# ---------------------------------------------------------------------------
MODEL_VISION = "claude-sonnet-4-6"

PROMPT_ETICHETTA = (
    "Leggi l'etichetta di questo cartone di prodotto alimentare (spesso surgelato) "
    "e restituisci SOLO un oggetto JSON valido, senza testo prima o dopo e senza backtick.\n"
    'Campi: {"lotto":"","scadenza":"","data_produzione":"","descrizione":"","fornitore":"","ean":"","incertezze":""}\n'
    "Istruzioni:\n"
    "- lotto: il codice di lotto stampato (anche scritto a mano), esattamente com'e'. \"\" se assente.\n"
    "- scadenza: data 'da consumarsi entro / TMC / best before', formato GG/MM/AAAA; se e' indicata solo mese/anno usa MM/AAAA. \"\" se assente.\n"
    "- data_produzione: data di produzione o congelamento, GG/MM/AAAA. \"\" se assente.\n"
    "- descrizione: nome/descrizione del prodotto come scritto sull'etichetta.\n"
    "- fornitore: marca, produttore o importatore se visibile (nome azienda). \"\" se assente.\n"
    "- ean: codice a barre numerico EAN/GTIN se leggibile. \"\" se assente.\n"
    "- incertezze: breve nota su caratteri ambigui (es. 'lotto: 2a cifra 0 o O'). \"\" se nessuna.\n"
    "Non inventare: se non leggi un campo lascia \"\"."
)

# Pesi del punteggio (regolabili: la chiave e' scadenza+descrizione, il lotto conferma)
W_LOTTO_EXACT = 45
W_LOTTO_DATE  = 42   # il lotto a gestionale coincide con la data (ri-lottatura)
W_LOTTO_FUZZY = 24
W_SCAD        = 40
W_SCAD_MY     = 28   # scadenza letta solo mese/anno
W_DESCR       = 12
W_DESCR_MAX   = 40
W_FORN        = 16   # conferma: bonus se coincide, MAI penalita' se diverso
SOGLIA_AFFIDABILE = 85


def _norm_lot(s):
    import re
    return re.sub(r'[^A-Z0-9]', '', (s or '').upper())


def _lev(a, b):
    a, b = a.upper(), b.upper()
    if a == b:
        return 0
    la, lb = len(a), len(b)
    if abs(la - lb) > 2:
        return 9
    prev = list(range(lb + 1))
    for i in range(1, la + 1):
        cur = [i] + [0] * lb
        for j in range(1, lb + 1):
            cur[j] = min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] != b[j - 1]))
        prev = cur
    return prev[lb]


def _date_as_lot_variants(d):
    """Da 'GG/MM/AAAA' genera ['GGMMAA','GGMMAAAA'] (per la ri-lottatura con data)."""
    import re
    m = re.match(r'^(\d{2})/(\d{2})/(\d{4})$', (d or '').strip())
    if not m:
        return []
    gg, mm, aaaa = m.groups()
    return [gg + mm + aaaa[2:], gg + mm + aaaa]


def _tokens(s):
    import re
    return [t for t in re.split(r'[^A-Za-z0-9]+', (s or '').upper()) if len(t) >= 3]


def _score_candidati(righe, letto, limite=3):
    import re
    lotto_letto = _norm_lot(letto.get("lotto", ""))
    scad = (letto.get("scadenza", "") or "").strip()
    scad_full = bool(re.match(r'^\d{2}/\d{2}/\d{4}$', scad))
    m_my = re.match(r'^(\d{1,2})/(\d{4})$', scad)
    scad_my = (m_my.group(1).zfill(2) + "/" + m_my.group(2)) if m_my else ""

    date_lots = set(_date_as_lot_variants(scad)) | set(_date_as_lot_variants(letto.get("data_produzione", "")))
    date_lots_n = {_norm_lot(x) for x in date_lots if x}

    toks = _tokens(letto.get("descrizione", ""))
    forn_toks = [t for t in _tokens(letto.get("fornitore", "")) if len(t) >= 4]

    best = {}
    for r in righe:
        sc = 0
        reasons = []
        rl = _norm_lot(r.get("lotto", ""))
        rdescr = (r.get("descrizione", "") or "").upper()
        rscad = (r.get("data_scadenza", "") or "")
        rforn = (r.get("fornitore", "") or "").upper()

        if lotto_letto and rl:
            if lotto_letto == rl:
                sc += W_LOTTO_EXACT; reasons.append("lotto")
            elif len(rl) >= 4 and (_lev(lotto_letto, rl) <= 1 or lotto_letto in rl or rl in lotto_letto):
                sc += W_LOTTO_FUZZY; reasons.append("lotto~")
        if rl and rl in date_lots_n and "lotto" not in reasons:
            sc += W_LOTTO_DATE; reasons.append("lotto=data")

        if scad_full and rscad == scad:
            sc += W_SCAD; reasons.append("scadenza")
        elif scad_my and rscad.endswith(scad_my):
            sc += W_SCAD_MY; reasons.append("scad mese/anno")

        if toks:
            hit = [t for t in toks if t[:5] in rdescr]
            if hit:
                sc += min(W_DESCR_MAX, W_DESCR * len(hit)); reasons.append("descrizione")

        if forn_toks and rforn and any(t in rforn for t in forn_toks):
            sc += W_FORN; reasons.append("fornitore")

        if sc <= 0:
            continue
        cod = r.get("codice", "")
        if cod not in best or sc > best[cod]["score"]:
            best[cod] = {
                "codice": cod,
                "descrizione": r.get("descrizione", ""),
                "scadenza": rscad,
                "lotto": r.get("lotto", ""),
                "fornitore": r.get("fornitore", ""),
                "score": sc,
                "motivi": reasons,
            }

    ranked = sorted(best.values(), key=lambda d: -d["score"])[:limite]
    if not ranked:
        return [], 0
    top = ranked[0]["score"]
    second = ranked[1]["score"] if len(ranked) > 1 else 0
    gap = top - second
    uni = 15 if gap >= 20 else (5 if gap >= 10 else -5)
    confidenza = max(15, min(98, top + uni))
    return ranked, confidenza


def cerca_candidati(letto, limite=3):
    righe = query(
        "SELECT codice, descrizione, data_scadenza, lotto, data_produzione, fornitore FROM articoli",
        fetchall=True,
    ) or []
    righe = [dict(r) for r in righe]
    return _score_candidati(righe, letto, limite)


def leggi_etichetta(image_bytes, media_type):
    import base64, json
    from anthropic import Anthropic
    client = Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))
    b64 = base64.b64encode(image_bytes).decode("ascii")
    msg = client.messages.create(
        model=MODEL_VISION,
        max_tokens=600,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": b64}},
                {"type": "text", "text": PROMPT_ETICHETTA},
            ],
        }],
    )
    txt = "".join(b.text for b in msg.content if getattr(b, "type", "") == "text").strip()
    if "```" in txt:
        parts = txt.split("```")
        txt = parts[1] if len(parts) > 1 else txt
        if txt.lstrip().lower().startswith("json"):
            txt = txt.lstrip()[4:]
    i, j = txt.find("{"), txt.rfind("}")
    if i != -1 and j != -1:
        txt = txt[i:j + 1]
    try:
        data = json.loads(txt)
    except Exception:
        data = {}
    campi = ["lotto", "scadenza", "data_produzione", "descrizione", "fornitore", "ean", "incertezze"]
    return {k: (str(data.get(k, "")).strip() if data.get(k) is not None else "") for k in campi}


def _merge_letture(letture):
    """Unisce le letture di piu' foto dello STESSO cartone nel miglior set di campi."""
    out = {k: "" for k in ["lotto", "scadenza", "data_produzione", "descrizione", "fornitore", "ean", "incertezze"]}
    descr_parts, incert_parts = [], []
    import re
    for L in letture:
        for k in ["lotto", "data_produzione", "fornitore", "ean"]:
            if not out[k] and L.get(k):
                out[k] = L[k]
        # scadenza: preferisci la data completa a mese/anno
        s = (L.get("scadenza") or "").strip()
        if s:
            if not out["scadenza"]:
                out["scadenza"] = s
            elif re.match(r'^\d{2}/\d{2}/\d{4}$', s) and not re.match(r'^\d{2}/\d{2}/\d{4}$', out["scadenza"]):
                out["scadenza"] = s
        if L.get("descrizione") and L["descrizione"] not in descr_parts:
            descr_parts.append(L["descrizione"])
        if L.get("incertezze"):
            incert_parts.append(L["incertezze"])
    out["descrizione"] = " | ".join(descr_parts)
    out["incertezze"] = " | ".join(incert_parts)
    return out


# ---------------------------------------------------------------------------
# Cartoni: memoria delle correzioni
# Impara dai feedback dell'operatore ("questo prodotto = questo codice") e,
# quando ripassa un cartone simile, richiama la risposta confermata.
# ---------------------------------------------------------------------------
def _solo_cifre(s):
    import re
    return re.sub(r'\D', '', s or '')


def _descr_tokens_set(s):
    # token significativi (>=3 char), come insieme ordinato
    return sorted(set(_tokens(s)))


def _overlap(a, b):
    if not a or not b:
        return 0.0
    return len(a & b) / min(len(a), len(b))


def _scegli_correzione(read_tokens, read_ean, rows):
    """Sceglie la correzione memorizzata che meglio corrisponde alla lettura."""
    best, best_key = None, (0.0, 0)
    for r in rows:
        rt = set((r.get("descr_norm") or "").split())
        r_ean = r.get("ean") or ""
        if read_ean and r_ean and read_ean == r_ean:
            score = 1.0
        else:
            if len(read_tokens & rt) < 2:      # servono almeno 2 parole in comune
                continue
            score = _overlap(read_tokens, rt)
            if score < 0.6:                    # e una sovrapposizione netta
                continue
        key = (score, int(r.get("conferme") or 1))
        if key > best_key:
            best, best_key = r, key
    return best


def registra_correzione(codice, descrizione, ean=""):
    """Salva/rinforza una correzione confermata dall'operatore."""
    codice = (codice or "").strip()
    if not codice:
        return
    toks = _descr_tokens_set(descrizione)
    descr_norm = " ".join(toks)
    ean_n = _solo_cifre(ean)
    rows = query("SELECT * FROM correzioni WHERE codice=%s", (codice,), fetchall=True) or []
    best = None
    for row in rows:
        if ean_n and (row.get("ean") or "") == ean_n:
            best = row; break
        if _overlap(set(toks), set((row.get("descr_norm") or "").split())) >= 0.6:
            best = row; break
    if best:
        union = " ".join(sorted(set(toks) | set((best.get("descr_norm") or "").split())))
        execute(
            "UPDATE correzioni SET descr_norm=%s, ean=COALESCE(NULLIF(%s,''), ean), "
            "conferme=conferme+1, updated_at=NOW() WHERE id=%s",
            (union, ean_n, best["id"]), commit=True,
        )
    else:
        execute(
            "INSERT INTO correzioni (codice, descr_norm, ean) VALUES (%s,%s,%s)",
            (codice, descr_norm, ean_n), commit=True,
        )


def applica_memoria(merged, candidati, confidenza):
    """Se un cartone simile è già stato confermato, porta quel codice in cima."""
    toks = set(_descr_tokens_set(merged.get("descrizione", "")))
    ean_n = _solo_cifre(merged.get("ean", ""))
    if not toks and not ean_n:
        return candidati, confidenza
    rows = query("SELECT codice, descr_norm, ean, conferme FROM correzioni", fetchall=True) or []
    scelta = _scegli_correzione(toks, ean_n, [dict(r) for r in rows])
    if not scelta:
        return candidati, confidenza

    cod = scelta["codice"]
    conf_n = int(scelta.get("conferme") or 1)
    motivo = f"già confermato ({conf_n})" if conf_n > 1 else "già confermato"

    trovato = next((c for c in candidati if c["codice"] == cod), None)
    if trovato:
        if motivo not in trovato["motivi"]:
            trovato["motivi"] = [motivo] + trovato["motivi"]
        candidati = [trovato] + [c for c in candidati if c is not trovato]
    else:
        art = query(
            "SELECT codice, descrizione, data_scadenza, lotto, fornitore "
            "FROM articoli WHERE codice=%s LIMIT 1",
            (cod,), fetchone=True,
        )
        nuovo = {
            "codice": cod,
            "descrizione": (art["descrizione"] if art else "(da correzione confermata)"),
            "scadenza": (art["data_scadenza"] if art else ""),
            "lotto": (art["lotto"] if art else ""),
            "fornitore": (art["fornitore"] if art else ""),
            "score": 999,
            "motivi": [motivo],
        }
        candidati = [nuovo] + candidati

    candidati = candidati[:3]
    confidenza = max(confidenza, min(98, 88 + conf_n * 2))
    return candidati, confidenza


# ---------------------------------------------------------------------------
# Cartoni: motore di identificazione (condiviso) + link operatore a scadenza
# ---------------------------------------------------------------------------
def _esegui_identificazione(files):
    files = [f for f in files if f and f.filename][:4]
    if not files:
        return {"ok": False, "msg": "Nessuna foto ricevuta."}, 400

    letture, errori = [], []
    for f in files:
        mt = f.mimetype or ""
        if mt not in ("image/jpeg", "image/png", "image/gif", "image/webp"):
            name = (f.filename or "").lower()
            mt = ("image/png" if name.endswith(".png")
                  else "image/webp" if name.endswith(".webp")
                  else "image/gif" if name.endswith(".gif")
                  else "image/jpeg")
        try:
            letture.append(leggi_etichetta(f.read(), mt))
        except Exception as e:
            errori.append(str(e))

    if not letture:
        return {"ok": False, "msg": "Lettura non riuscita. " + " ".join(errori)}, 502

    merged = _merge_letture(letture)
    candidati, confidenza = cerca_candidati(merged)
    candidati, confidenza = applica_memoria(merged, candidati, confidenza)
    proposto = candidati[0]["codice"] if candidati else ""
    ric_id = None
    try:
        row = execute_returning(
            """INSERT INTO riconoscimenti
               (letto_lotto, letto_scadenza, letto_fornitore, letto_descrizione,
                letto_ean, codice_proposto, confidenza, n_foto)
               VALUES (%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id""",
            (merged.get("lotto"), merged.get("scadenza"), merged.get("fornitore"),
             merged.get("descrizione"), merged.get("ean"), proposto,
             int(confidenza), len(letture)),
        )
        ric_id = row["id"] if row else None
    except Exception:
        pass

    return {
        "ok": True, "id": ric_id, "letto": merged, "candidati": candidati,
        "confidenza": confidenza, "soglia": SOGLIA_AFFIDABILE, "n_foto": len(letture),
    }, 200


def _scanner_token_valido(token):
    return query(
        "SELECT * FROM scanner_token WHERE token=%s AND scadenza > NOW()",
        (token,), fetchone=True,
    )


# --- Operatore: scanner via link a scadenza, senza login ---
@app.route("/scan/<token>")
def scanner_operatore(token):
    if _scanner_token_valido(token) is None:
        return render_template("scan_scaduto.html"), 410
    return render_template("scan.html", token=token)


@app.route("/api/scan/<token>/identifica", methods=["POST"])
def scanner_identifica(token):
    if _scanner_token_valido(token) is None:
        return jsonify({"ok": False, "msg": "Link scaduto o non valido."}), 410
    if not os.environ.get("ANTHROPIC_API_KEY"):
        return jsonify({"ok": False, "msg": "Chiave API non configurata sul server."}), 500
    payload, status = _esegui_identificazione(request.files.getlist("foto"))
    return jsonify(payload), status


@app.route("/api/scan/<token>/feedback", methods=["POST"])
def scanner_feedback(token):
    if _scanner_token_valido(token) is None:
        return jsonify({"ok": False, "msg": "Link scaduto o non valido."}), 410
    data = request.get_json(force=True) or {}
    ric_id = data.get("id")
    esito = (data.get("esito") or "").strip()          # 'corretto' | 'alternativa' | 'nessuno'
    codice_corretto = (data.get("codice_corretto") or "").strip() or None
    if not ric_id or esito not in ("corretto", "alternativa", "nessuno"):
        return jsonify({"ok": False, "msg": "Feedback non valido."}), 400
    try:
        execute(
            "UPDATE riconoscimenti SET esito=%s, codice_corretto=%s WHERE id=%s",
            (esito, codice_corretto, ric_id),
            commit=True,
        )
    except Exception:
        return jsonify({"ok": False, "msg": "Errore nel salvataggio."}), 500

    # Se conosciamo il codice giusto, impariamolo (memoria delle correzioni)
    if codice_corretto:
        try:
            ric = query(
                "SELECT letto_descrizione, letto_ean FROM riconoscimenti WHERE id=%s",
                (ric_id,), fetchone=True,
            )
            if ric:
                registra_correzione(codice_corretto, ric.get("letto_descrizione", ""),
                                    ric.get("letto_ean", ""))
        except Exception:
            pass

    return jsonify({"ok": True})


# --- Ufficio: genera un link operatore valido 24 ore ---
@app.route("/admin/cartoni/genera-link", methods=["POST"])
@login_required
def admin_cartoni_genera_link():
    nome = request.form.get("nome", "").strip()
    telefono = request.form.get("telefono", "").strip()
    token = secrets.token_urlsafe(16)
    execute(
        """INSERT INTO scanner_token (token, nome, telefono, scadenza)
           VALUES (%s, %s, %s, NOW() + INTERVAL '24 hours')""",
        (token, nome, telefono),
        commit=True,
    )
    flash("Link operatore generato: valido 24 ore.", "success")
    return redirect(url_for("admin_cartoni"))


# ---------------------------------------------------------------------------
# Health check
# ---------------------------------------------------------------------------
@app.route("/health")
def health():
    return "ok", 200


if __name__ == "__main__":
    app.run(debug=False, host="0.0.0.0", port=5000)
