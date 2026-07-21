"""
Wily Facturador - App Web
Factura C para cuenta personal (Monotributista)
"""

from flask import Flask, render_template, request, redirect, url_for, session, flash
from functools import wraps
import sqlite3
import json
import os
import csv
import io
from datetime import datetime

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "wily-secret-2024")

DB_PATH = os.environ.get("DB_PATH", "facturador.db")

USUARIOS = {
    "facturador": {
        "password": os.environ.get("PASS_FACTURADOR", "wily123"),
        "nombre": "Facturador",
    },
    "admin": {
        "password": os.environ.get("PASS_ADMIN", "admin2024"),
        "nombre": "Admin",
    },
}

PUNTO_VENTA = int(os.environ.get("PUNTO_VENTA", "1"))
TIPO_COMPROBANTE = 11  # Factura C

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with get_db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS comprobantes (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                numero       INTEGER,
                punto_venta  INTEGER,
                fecha        TEXT,
                destinatario TEXT,
                doc_tipo     INTEGER,
                doc_nro      TEXT,
                monto        REAL,
                cae          TEXT,
                cae_vto      TEXT,
                estado       TEXT,
                error        TEXT,
                usuario      TEXT,
                created_at   TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.commit()

init_db()

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "usuario" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated

@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        usuario = request.form.get("usuario", "").lower()
        password = request.form.get("password", "")
        if usuario in USUARIOS and USUARIOS[usuario]["password"] == password:
            session["usuario"] = usuario
            session["nombre"] = USUARIOS[usuario]["nombre"]
            return redirect(url_for("dashboard"))
        error = "Usuario o contrasena incorrectos"
    return render_template("login.html", error=error)

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/")
@login_required
def dashboard():
    with get_db() as conn:
        rows = conn.execute(
            "SELECT * FROM comprobantes ORDER BY created_at DESC LIMIT 100"
        ).fetchall()
        total_mes = conn.execute("""
            SELECT COALESCE(SUM(monto),0) FROM comprobantes
            WHERE estado='APROBADO'
            AND strftime('%Y-%m', fecha) = strftime('%Y-%m', 'now')
        """).fetchone()[0]
        cant_mes = conn.execute("""
            SELECT COUNT(*) FROM comprobantes
            WHERE estado='APROBADO'
            AND strftime('%Y-%m', fecha) = strftime('%Y-%m', 'now')
        """).fetchone()[0]

    return render_template("dashboard.html",
        comprobantes=rows,
        total_mes=total_mes,
        cant_mes=cant_mes,
        usuario=session["usuario"],
    )

@app.route("/emitir", methods=["GET", "POST"])
@login_required
def emitir():
    if request.method == "POST":
        data = {
            "nombre":   request.form.get("nombre"),
            "doc_tipo": int(request.form.get("doc_tipo", 99)),
            "doc_nro":  request.form.get("doc_nro", ""),
            "monto":    float(request.form.get("monto", 0)),
        }
        resultado = emitir_factura_arca(data)
        guardar_comprobante(resultado)

        if resultado["estado"] == "APROBADO":
            flash("Factura C #{} emitida. CAE: {}".format(
                resultado["numero"], resultado["cae"]), "success")
        else:
            flash("Error: {}".format(resultado["error"]), "error")

        return redirect(url_for("dashboard"))

    return render_template("emitir.html", usuario=session["usuario"])

@app.route("/masivo", methods=["GET", "POST"])
@login_required
def masivo():
    resultados = []
    if request.method == "POST":
        archivo = request.files.get("archivo")
        if not archivo:
            flash("Selecciona un archivo Excel o CSV", "error")
            return redirect(url_for("masivo"))

        nombre_archivo = archivo.filename.lower()
        filas = []

        if nombre_archivo.endswith(".xlsx") or nombre_archivo.endswith(".xls"):
            try:
                import openpyxl
                wb = openpyxl.load_workbook(archivo)
                ws = wb.active

                header_row = None
                headers = []
                for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
                    valores = [str(v).strip().upper() if v else "" for v in row]
                    if any(v in ("RAZONSOCIAL", "NOMBRE", "TIPODOCUMENTO", "DOC_TIPO") for v in valores):
                        header_row = i
                        headers = [str(v).strip().lower() if v else "" for v in row]
                        break

                if not header_row:
                    flash("No se encontraron los encabezados en el Excel.", "error")
                    return redirect(url_for("masivo"))

                for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
                    if not any(row):
                        continue
                    fila = {headers[i]: (str(v).strip() if v is not None else "") for i, v in enumerate(row)}
                    filas.append(fila)
            except Exception as e:
                flash("Error al leer el Excel: {}".format(str(e)), "error")
                return redirect(url_for("masivo"))
        else:
            content = archivo.read().decode("utf-8")
            reader = csv.DictReader(io.StringIO(content))
            filas = [{k.lower().strip(): v for k, v in row.items()} for row in reader]

        for fila in filas:
            nombre = (fila.get("razonsocial") or fila.get("nombre") or "").strip()
            monto = float(fila.get("monto") or 0)

            tipo_doc_raw = (fila.get("tipodocumento") or fila.get("doc_tipo") or "CF").strip().upper()
            if tipo_doc_raw in ("CUIT", "80"):
                doc_tipo = 80
            elif tipo_doc_raw in ("DNI", "96"):
                doc_tipo = 96
            else:
                doc_tipo = 99

            doc_nro = (fila.get("nrodocumento") or fila.get("doc_nro") or "").strip()
            doc_nro = doc_nro.replace("-", "").replace(".", "")

            data = {
                "nombre":   nombre,
                "doc_tipo": doc_tipo,
                "doc_nro":  doc_nro,
                "monto":    monto,
            }
            resultado = emitir_factura_arca(data)
            guardar_comprobante(resultado)
            resultados.append(resultado)

    return render_template("masivo.html",
        usuario=session["usuario"],
        resultados=resultados,
    )

def emitir_factura_arca(data):
    import random, string

    MODO_DEMO = not os.environ.get("CUIT")
    hoy = datetime.now().strftime("%Y-%m-%d")

    if MODO_DEMO:
        numero = random.randint(1000, 9999)
        cae = "".join(random.choices(string.digits, k=14))
        return {
            "numero":        numero,
            "punto_venta":   PUNTO_VENTA,
            "fecha":         hoy,
            "destinatario":  data["nombre"],
            "doc_tipo":      data["doc_tipo"],
            "doc_nro":       data["doc_nro"],
            "monto":         data["monto"],
            "cae":           cae,
            "cae_vto":       "20251231",
            "estado":        "APROBADO",
            "error":         None,
            "usuario":       session["usuario"],
        }

    try:
        from zeep import Client
        import tempfile
        import xml.etree.ElementTree as ET
        from datetime import timedelta

        CUIT       = os.environ["CUIT"]
        CERT_PATH  = os.environ["CERT_PATH"]
        KEY_PATH   = os.environ["KEY_PATH"]
        PRODUCCION = os.environ.get("PRODUCCION", "false").lower() == "true"
        CONCEPTO   = int(os.environ.get("CONCEPTO", "2"))

        WSAA_URL = (
            "https://wsaa.afip.gov.ar/ws/services/LoginCms?wsdl"
            if PRODUCCION else
            "https://wsaahomo.afip.gov.ar/ws/services/LoginCms?wsdl"
        )
        WSFE_URL = (
            "https://servicios1.afip.gov.ar/wsfev1/service.asmx?WSDL"
            if PRODUCCION else
            "https://wswhomo.afip.gov.ar/wsfev1/service.asmx?WSDL"
        )

        cache_file = "/tmp/ticket_{}.json".format(CUIT)
        ticket = None
        if os.path.exists(cache_file):
            cached = json.loads(open(cache_file).read())
            if datetime.now() < datetime.fromisoformat(cached["expiry"]) - timedelta(minutes=10):
                ticket = cached

        if not ticket:
            from datetime import timezone
            now = datetime.now(timezone.utc).astimezone(timezone(timedelta(hours=-3)))
            exp = now + timedelta(hours=12)
            xml_tra = """<?xml version="1.0" encoding="UTF-8"?>
<loginTicketRequest version="1.0">
  <header>
    <uniqueId>{}</uniqueId>
    <generationTime>{}</generationTime>
    <expirationTime>{}</expirationTime>
  </header>
  <service>wsfe</service>
</loginTicketRequest>""".format(
                int(now.timestamp()),
                now.strftime('%Y-%m-%dT%H:%M:%S'),
                exp.strftime('%Y-%m-%dT%H:%M:%S')
            ).encode()

            from cryptography.hazmat.primitives.serialization.pkcs7 import PKCS7SignatureBuilder
            from cryptography.hazmat.primitives.serialization import Encoding
            from cryptography.hazmat.primitives import hashes
            from cryptography.hazmat.backends import default_backend
            from cryptography.hazmat.primitives.serialization import load_pem_private_key
            from cryptography.x509 import load_pem_x509_certificate
            import base64

            with open(CERT_PATH, "rb") as fc:
                cert = load_pem_x509_certificate(fc.read(), default_backend())
            with open(KEY_PATH, "rb") as fk:
                key = load_pem_private_key(fk.read(), password=None, backend=default_backend())

            cms_der = (
                PKCS7SignatureBuilder()
                .set_data(xml_tra)
                .add_signer(cert, key, hashes.SHA256())
                .sign(Encoding.DER, options=[])
            )
            cms_b64 = base64.b64encode(cms_der).decode()

            wsaa = Client(WSAA_URL)
            resp_xml = wsaa.service.loginCms(in0=cms_b64)
            root = ET.fromstring(resp_xml)

            ticket = {
                "token":  root.find(".//token").text,
                "sign":   root.find(".//sign").text,
                "expiry": exp.isoformat(),
            }
            open(cache_file, "w").write(json.dumps(ticket))

        auth = {"Token": ticket["token"], "Sign": ticket["sign"], "Cuit": int(CUIT)}

        wsfe = Client(WSFE_URL)
        ultimo = wsfe.service.FECompUltimoAutorizado(
            Auth=auth, PtoVta=PUNTO_VENTA, CbteTipo=TIPO_COMPROBANTE
        ).CbteNro
        numero = ultimo + 1

        hoy_arca = datetime.now().strftime("%Y%m%d")
        monto = round(data["monto"], 2)

        detalle = {
            "Concepto":     CONCEPTO,
            "DocTipo":      data["doc_tipo"],
            "DocNro":       int(data["doc_nro"]) if data["doc_nro"] else 0,
            "CbteDesde":    numero,
            "CbteHasta":    numero,
            "CbteFch":      hoy_arca,
            "ImpTotal":     monto,
            "ImpTotConc":   0,
            "ImpNeto":      monto,
            "ImpOpEx":      0,
            "ImpIVA":       0,
            "ImpTrib":      0,
            "FchServDesde": hoy_arca if CONCEPTO in (2, 3) else None,
            "FchServHasta": hoy_arca if CONCEPTO in (2, 3) else None,
            "FchVtoPago":   hoy_arca if CONCEPTO in (2, 3) else None,
            "MonId":        "PES",
            "MonCotiz":     1,
        }

        resp = wsfe.service.FECAESolicitar(Auth=auth, FeCAEReq={
            "FeCabReq": {"CantReg": 1, "PtoVta": PUNTO_VENTA, "CbteTipo": TIPO_COMPROBANTE},
            "FeDetReq": {"FECAEDetRequest": [detalle]},
        })

        det = resp.FeDetResp.FECAEDetResponse[0]
        if det.Resultado == "A":
            return {
                "numero": numero, "punto_venta": PUNTO_VENTA,
                "fecha": hoy, "destinatario": data["nombre"],
                "doc_tipo": data["doc_tipo"], "doc_nro": data["doc_nro"],
                "monto": monto, "cae": det.CAE, "cae_vto": det.CAEFchVto,
                "estado": "APROBADO", "error": None,
                "usuario": session["usuario"],
            }
        else:
            errores = []
            if det.Observaciones:
                for o in det.Observaciones.Obs:
                    errores.append("[{}] {}".format(o.Code, o.Msg))
            raise Exception("; ".join(errores))

    except Exception as e:
        return {
            "numero": None, "punto_venta": PUNTO_VENTA,
            "fecha": hoy, "destinatario": data["nombre"],
            "doc_tipo": data.get("doc_tipo"), "doc_nro": data.get("doc_nro"),
            "monto": data["monto"], "cae": None, "cae_vto": None,
            "estado": "RECHAZADO", "error": str(e),
            "usuario": session["usuario"],
        }

def guardar_comprobante(r):
    with get_db() as conn:
        conn.execute("""
            INSERT INTO comprobantes
            (numero, punto_venta, fecha, destinatario, doc_tipo, doc_nro,
             monto, cae, cae_vto, estado, error, usuario)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            r.get("numero"), r.get("punto_venta"), r.get("fecha"),
            r.get("destinatario"), r.get("doc_tipo"), r.get("doc_nro"),
            r.get("monto"), r.get("cae"), r.get("cae_vto"),
            r.get("estado"), r.get("error"), r.get("usuario"),
        ))
        conn.commit()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
