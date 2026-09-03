from __future__ import annotations

import argparse
import base64
import calendar
import hashlib
import json
import re
import os
import threading
import time
import shutil
import webbrowser
import urllib.parse
import urllib.request
from datetime import datetime, timedelta, timezone
from http.server import BaseHTTPRequestHandler, HTTPServer
from email.message import EmailMessage
from pathlib import Path
from secrets import token_urlsafe

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
SQL_FILE = BASE_DIR / "Aventis_1.sql"
GMAIL_CREDENTIALS_FILE = BASE_DIR / "credenciales.json"
GMAIL_TOKEN_FILE = BASE_DIR / "token.json"
STATE_FILE = BASE_DIR / "aventis_state.json"
OUTPUT_DIR = BASE_DIR / "salidas"

SERVER = "150.1.1.152"
DATABASE = "SIR"
USERNAME = "ConsultaBD"
PASSWORD = "5D$bc#kM&5W2T8J40?s%"

SHEET_IMPORT = "Cont PT IMP"
SHEET_EXPORT = "Cont PT EXP"

GMAIL_RECIPIENTS = [
    "gerencia.ver@abcsc.mx",
    "sgonzalez@abcsc.mx",
    "jperez@abcsc.mx",
    "myanez@abcsc.mx",
    "ccarbajal@abcsc.mx",
    "ssalguero@abcsc.mx",
    "imedrano@abcsc.mx",
    "apalacios@abcsc.mx",
]
GMAIL_FROM = "reportes.bi@abcsc.mx"
GMAIL_SCOPE = "https://www.googleapis.com/auth/gmail.send"

SEND_DAY = 1
SEND_HOUR = 0
SEND_MINUTE = 0
DAEMON_SLEEP_SECONDS = 900

EXPORT_COLUMNS = [
    "Pedimento Original A1",
    "Tipo Operación Desc",
    "Clave Pedimento",
    "Contenedores",
    "QTY Contenedor",
    "Peso Bruto",
    "Clave de País Origen/Destino",
    "Mercancía",
    "Fecha de Pago funcion",
    "Total de Bultos",
]

MONTH_NAMES = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}

SQL_START_DATE = "2026-05-01"
SQL_END_DATE = "2026-05-31"


def leer_query(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError:
        return path.read_text(encoding="latin-1")


def ajustar_rango_fechas(query: str, fecha_ini: str, fecha_fin: str) -> str:
    query = re.sub(
        r"(TRY_CONVERT\(DATE,\s*\[Fecha de Pago funcion\],\s*103\)\s*>=\s*)'[^']+'",
        rf"\1'{fecha_ini}'",
        query,
    )
    query = re.sub(
        r"(TRY_CONVERT\(DATE,\s*\[Fecha de Pago funcion\],\s*103\)\s*<=\s*)'[^']+'",
        rf"\1'{fecha_fin}'",
        query,
    )
    return query


def periodo_anterior(ref: datetime) -> tuple[int, int]:
    if ref.month == 1:
        return ref.year - 1, 12
    return ref.year, ref.month - 1


def primer_dia_mes(year: int, month: int) -> datetime:
    return datetime(year, month, 1)


def ultimo_dia_mes(year: int, month: int) -> datetime:
    ultimo = calendar.monthrange(year, month)[1]
    return datetime(year, month, ultimo)


def formatear_periodo(year: int, month: int) -> str:
    return f"{MONTH_NAMES[month].upper()} {year}"


def periodo_key(year: int, month: int) -> str:
    return f"{year:04d}-{month:02d}"


def period_due_datetime(year: int, month: int) -> datetime:
    if month == 12:
        due_year = year + 1
        due_month = 1
    else:
        due_year = year
        due_month = month + 1
    return datetime(due_year, due_month, SEND_DAY, SEND_HOUR, SEND_MINUTE)


def cargar_estado() -> dict:
    if not STATE_FILE.exists():
        return {"reports": {}}
    try:
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {"reports": {}}


def guardar_estado(state: dict) -> None:
    STATE_FILE.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def marcar_estado(state: dict, year: int, month: int, **updates) -> None:
    reports = state.setdefault("reports", {})
    key = periodo_key(year, month)
    current = reports.get(key, {})
    current.update(updates)
    reports[key] = current


def estado_enviado(state: dict, year: int, month: int) -> bool:
    return state.get("reports", {}).get(periodo_key(year, month), {}).get("status") == "sent"


def confirmar_envio_anticipado(periodo: str) -> bool:
    while True:
        respuesta = input(
            f"Aun no toca enviar el reporte de {periodo}.\n"
            "¿Desea enviarlo de forma anticipada? (S/N): "
        ).strip().upper()

        if respuesta in {"S", "SI"}:
            return True
        if respuesta in {"N", "NO"}:
            return False

        print("Respuesta invalida. Escriba S, SI, N o NO.")


def cargar_configuracion_gmail() -> dict:
    if not GMAIL_CREDENTIALS_FILE.exists():
        raise FileNotFoundError(
            f"No existe {GMAIL_CREDENTIALS_FILE.name}. "
            "Ese archivo debe contener el cliente OAuth de Gmail."
        )

    data = json.loads(GMAIL_CREDENTIALS_FILE.read_text(encoding="utf-8"))
    cfg = data.get("installed") or data.get("web")
    if not cfg:
        raise ValueError(
            f"{GMAIL_CREDENTIALS_FILE.name} no tiene bloque 'installed' ni 'web'."
        )
    return cfg


def cargar_token() -> dict | None:
    if not GMAIL_TOKEN_FILE.exists():
        return None
    try:
        return json.loads(GMAIL_TOKEN_FILE.read_text(encoding="utf-8"))
    except Exception:
        return None


def guardar_token(token_data: dict) -> None:
    GMAIL_TOKEN_FILE.write_text(
        json.dumps(token_data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def _token_expira_en(token_data: dict | None) -> bool:
    if not token_data or not token_data.get("expiry"):
        return True
    try:
        expiry = datetime.fromisoformat(token_data["expiry"])
    except Exception:
        return True
    return expiry <= datetime.now(timezone.utc)


def _codigo_pkce() -> tuple[str, str]:
    verifier = token_urlsafe(64)
    challenge = base64.urlsafe_b64encode(
        hashlib.sha256(verifier.encode("ascii")).digest()
    ).decode("ascii").rstrip("=")
    return verifier, challenge


def _solicitud_form(url: str, datos: dict) -> dict:
    body = urllib.parse.urlencode(datos).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=body,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _solicitud_json(url: str, headers: dict | None = None) -> dict:
    req = urllib.request.Request(url, headers=headers or {})
    with urllib.request.urlopen(req, timeout=60) as resp:
        return json.loads(resp.read().decode("utf-8"))


def abrir_navegador_autorizacion(url: str) -> None:
    candidatos = []
    if shutil.which("microsoft-edge"):
        candidatos.append("microsoft-edge")
    if shutil.which("edge"):
        candidatos.append("edge")
    if shutil.which("firefox"):
        candidatos.append("firefox")

    for nombre in candidatos:
        try:
            navegador = webbrowser.get(nombre)
            if navegador.open(url, new=1, autoraise=True):
                return
        except webbrowser.Error:
            continue

    webbrowser.open(url, new=1, autoraise=True)


def obtener_access_token() -> dict:
    cfg = cargar_configuracion_gmail()
    token_data = cargar_token()

    if token_data and not _token_expira_en(token_data):
        return token_data

    if token_data and token_data.get("refresh_token"):
        refreshed = _solicitud_form(
            cfg["token_uri"],
            {
                "client_id": cfg["client_id"],
                "client_secret": cfg.get("client_secret", ""),
                "refresh_token": token_data["refresh_token"],
                "grant_type": "refresh_token",
            },
        )
        token_data["access_token"] = refreshed["access_token"]
        expires_in = int(refreshed.get("expires_in", 3600))
        token_data["expiry"] = (
            datetime.now(timezone.utc) + timedelta(seconds=expires_in)
        ).isoformat()
        guardar_token(token_data)
        return token_data

    verifier, challenge = _codigo_pkce()
    state = token_urlsafe(24)
    resultado = {}

    class CallbackHandler(BaseHTTPRequestHandler):
        def do_GET(self):
            parsed = urllib.parse.urlparse(self.path)
            params = urllib.parse.parse_qs(parsed.query)
            resultado["code"] = params.get("code", [None])[0]
            resultado["state"] = params.get("state", [None])[0]
            resultado["error"] = params.get("error", [None])[0]
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(
                b"<html><body><h3>Autorizacion completada.</h3>"
                b"<p>Puedes cerrar esta ventana y regresar al script.</p></body></html>"
            )

        def log_message(self, format, *args):
            return

    servidor = HTTPServer(("localhost", 0), CallbackHandler)
    redirect_uri = f"http://localhost:{servidor.server_address[1]}/"

    auth_params = {
        "client_id": cfg["client_id"],
        "redirect_uri": redirect_uri,
        "response_type": "code",
        "scope": GMAIL_SCOPE,
        "access_type": "offline",
        "prompt": "consent",
        "state": state,
        "code_challenge": challenge,
        "code_challenge_method": "S256",
    }
    auth_url = f"{cfg['auth_uri']}?{urllib.parse.urlencode(auth_params)}"

    thread = threading.Thread(target=servidor.serve_forever, daemon=True)
    thread.start()

    print("Abriendo navegador para autorizar Gmail...")
    print(auth_url)
    abrir_navegador_autorizacion(auth_url)

    timeout = time.time() + 300
    while time.time() < timeout and "code" not in resultado and "error" not in resultado:
        time.sleep(0.5)

    servidor.shutdown()
    servidor.server_close()

    if not resultado.get("code") and not resultado.get("error"):
        print(
            "\nNo se recibió el callback local de OAuth.\n"
            "Copia y pega aquí el codigo que aparece en la URL de redireccion, "
            "o pega la URL completa que te dio Google."
        )
        entrada = input("Codigo o URL: ").strip()
        if "code=" in entrada:
            parsed = urllib.parse.urlparse(entrada)
            params = urllib.parse.parse_qs(parsed.query)
            resultado["code"] = params.get("code", [None])[0]
            resultado["state"] = params.get("state", [None])[0]
        else:
            resultado["code"] = entrada
            resultado["state"] = state

    if resultado.get("error"):
        raise RuntimeError(f"OAuth cancelado o fallido: {resultado['error']}")
    if not resultado.get("code"):
        raise RuntimeError(
            "No se recibió el código OAuth. Abre la URL impresa, autoriza y reintenta."
        )
    if resultado.get("state") != state:
        raise RuntimeError("La respuesta OAuth no coincide con el estado esperado.")

    token_data = _solicitud_form(
        cfg["token_uri"],
        {
            "code": resultado["code"],
            "client_id": cfg["client_id"],
            "client_secret": cfg.get("client_secret", ""),
            "redirect_uri": redirect_uri,
            "grant_type": "authorization_code",
            "code_verifier": verifier,
        },
    )

    expires_in = int(token_data.get("expires_in", 3600))
    token_data["expiry"] = (
        datetime.now(timezone.utc) + timedelta(seconds=expires_in)
    ).isoformat()
    guardar_token(token_data)
    return token_data


def obtener_datos(fecha_ini: str, fecha_fin: str) -> pd.DataFrame:
    import pyodbc

    conn_str = (
        "DRIVER={ODBC Driver 18 for SQL Server};"
        f"SERVER={SERVER};DATABASE={DATABASE};"
        f"UID={USERNAME};PWD={PASSWORD};"
        "TrustServerCertificate=yes;"
    )

    query = ajustar_rango_fechas(leer_query(SQL_FILE), fecha_ini, fecha_fin)

    with pyodbc.connect(conn_str, timeout=30) as conn:
        df = pd.read_sql_query(query, conn)

    df.columns = df.columns.str.strip()
    return df


def filtrar_tipo_operacion(df: pd.DataFrame, texto: str) -> pd.DataFrame:
    columna = "Tipo Operación Desc"
    if columna not in df.columns:
        raise KeyError(f"No existe la columna requerida: {columna}")

    serie = df[columna].fillna("").astype(str).str.strip().str.lower()
    return df[serie.str.contains(texto.lower(), na=False)].copy()


def descomponer_total_bultos(df: pd.DataFrame) -> pd.DataFrame:
    faltantes = [col for col in EXPORT_COLUMNS if col not in df.columns]
    if faltantes:
        raise KeyError(f"Faltan columnas requeridas para exportar: {faltantes}")

    salida = df.copy()
    salida["Descripcion de la mercancía"] = salida["Mercancía"]

    fecha = pd.to_datetime(salida["Fecha de Pago funcion"], dayfirst=True, errors="coerce")
    salida["Mes"] = fecha.dt.month.map(MONTH_NAMES)
    tipo_contenedor = []
    pallet = []
    bultos = []

    for _, fila in salida.iterrows():
        valor = str(fila["Total de Bultos"]).strip()
        valor_up = valor.upper()
        tipo = ""
        pal = ""
        bult = ""

        if valor and valor.lower() != "nan":
            if "BULT" in valor_up and "CONTENEDOR" not in valor_up:
                bult = valor
            else:
                match = re.match(
                    r"^(?P<count>\d+)\s+CONTENEDORES?\s+(?P<tipo>[^()]+?)(?:\s*\((?P<pallet>[^()]+)\))?$",
                    valor,
                    flags=re.IGNORECASE,
                )
                if match:
                    tipo = match.group("tipo").strip()
                    pal = (match.group("pallet") or "").strip()
                    try:
                        qty_texto = int(match.group("count").strip())
                        qty_df = pd.to_numeric(fila["QTY Contenedor"], errors="coerce")
                        if pd.notna(qty_df) and int(qty_df) != qty_texto:
                            pass
                    except Exception:
                        pass

        tipo_contenedor.append(tipo)
        pallet.append(pal)
        bultos.append(bult)

    salida["Tipo Contenedor"] = tipo_contenedor
    salida["# Pallet"] = pallet
    salida["# Bultos"] = bultos

    return salida


def preparar_hoja_general(df: pd.DataFrame, sheet_all_name: str) -> pd.DataFrame:
    salida = descomponer_total_bultos(df)
    salida = salida.rename(
        columns={
            "Fecha de Pago funcion": "Pedimento Fecha Pago",
            "Entrada de pago": "Custom Clearance Time (Days)",
            "Mercancía": "Descripcion de la mercancía",
        }
    )

    columnas = [
        "Sucursal",
        "Tipo Sucursal",
        "Referencia",
        "Cliente",
        "Patente",
        "Pedimento Original A1",
        "Pedimento R1",
        "Tipo Operación Desc",
        "Clave Pedimento",
        "Tipo de Cambio de Pedimento",
        "Aduana Despacho",
        "Contenedores",
        "QTY Contenedor",
        "Peso Bruto",
        "Tipo Contenedor",
        "# Pallet",
        "# Bultos",
        "Valor Comercial MXP",
        "Valor Aduana",
        "Seguros",
        "Fletes",
        "Embalajes",
        "Otros",
        "TASA IGI/IGE",
        "IGI/IGIE",
        "Recargos",
        "DTA",
        "IVA",
        "Prevalidación",
        "IVA Prevalidación",
        "Impuestos",
        "PECE",
        "Facturas",
        "RazonSocial de Proveedores",
        "Clave Incoterm",
        "Clave de País Origen/Destino",
        "Clave de País Vendedor/Comprador",
        "Moneda",
        "Bls MASTER",
        "Bls HOUSE",
        "Fracciones",
        "Descripcion de la mercancía",
        "PT",
        "Tempreratura",
        "UNIDAD",
        "PLACAS",
        "Fecha Entrada/Presentación",
        "Pedimento Fecha Pago",
        "Custom Clearance Time (Days)",
        "MOTIVO DE RETRASO COMPLETO",
        "Cuenta de Gastos",
        "Permisos",
    ]
    return salida[columnas]


def preparar_hoja_contable(df: pd.DataFrame) -> pd.DataFrame:
    salida = descomponer_total_bultos(df)
    columnas = [
        "Pedimento Original A1",
        "Tipo Operación Desc",
        "Clave Pedimento",
        "Contenedores",
        "QTY Contenedor",
        "Peso Bruto",
        "Clave de País Origen/Destino",
        "Descripcion de la mercancía",
        "Fecha de Pago funcion",
        "Mes",
        "Tipo Contenedor",
        "# Pallet",
        "# Bultos",
    ]

    return salida[columnas]


def guardar_excel(df: pd.DataFrame, output_path: Path, sheet_all_name: str) -> Path:
    df_general = preparar_hoja_general(df, sheet_all_name)
    df_import = preparar_hoja_contable(filtrar_tipo_operacion(df, "Importación"))
    df_export = preparar_hoja_contable(filtrar_tipo_operacion(df, "Exportación"))

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_general.to_excel(writer, sheet_name=sheet_all_name[:31], index=False)
        df_import.to_excel(writer, sheet_name=SHEET_IMPORT, index=False)
        df_export.to_excel(writer, sheet_name=SHEET_EXPORT, index=False)

    return output_path


def enviar_correo_gmail(archivo: Path, subject: str, body: str) -> dict:
    token_data = obtener_access_token()

    mensaje = EmailMessage()
    mensaje["From"] = GMAIL_FROM
    mensaje["To"] = ", ".join(GMAIL_RECIPIENTS)
    mensaje["Subject"] = subject
    mensaje.set_content(body)

    contenido = archivo.read_bytes()
    mensaje.add_attachment(
        contenido,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=archivo.name,
    )

    raw = base64.urlsafe_b64encode(mensaje.as_bytes()).decode("ascii").rstrip("=")
    payload = json.dumps({"raw": raw}).encode("utf-8")
    req = urllib.request.Request(
        "https://gmail.googleapis.com/gmail/v1/users/me/messages/send",
        data=payload,
        headers={
            "Authorization": f"Bearer {token_data['access_token']}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        respuesta = json.loads(resp.read().decode("utf-8"))

    if not respuesta.get("id"):
        raise RuntimeError(f"Gmail no regreso id de mensaje. Respuesta: {respuesta}")

    return respuesta


def procesar_periodo(
    year: int,
    month: int,
    state: dict,
    solicitud: bool = False,
    fecha_fin_override: datetime | None = None,
) -> Path:
    fecha_ini = primer_dia_mes(year, month).strftime("%Y-%m-%d")
    fecha_fin_dt = fecha_fin_override or ultimo_dia_mes(year, month)
    fecha_fin = fecha_fin_dt.strftime("%Y-%m-%d")
    periodo = formatear_periodo(year, month)
    if solicitud:
        output_path = OUTPUT_DIR / f"Aventis_{year}_{month:02d}_solicitud_{fecha_fin}.xlsx"
    else:
        output_path = OUTPUT_DIR / f"Aventis_{year}_{month:02d}.xlsx"
    sheet_all_name = f"Lay out {periodo}"

    df = obtener_datos(fecha_ini, fecha_fin)
    guardar_excel(df, output_path, sheet_all_name)

    subject = f"Aventis - {periodo}"
    if solicitud:
        subject = f"{subject} - Solicitud extraordinaria al {fecha_fin}"

    body = (
        f"Adjunto el archivo de Aventis correspondiente a {periodo}.\n\n"
        f"Periodo analizado: {fecha_ini} a {fecha_fin}\n"
        f"Archivo generado: {output_path.name}\n"
    )

    respuesta_gmail = enviar_correo_gmail(output_path, subject, body)
    print(
        "Correo aceptado por Gmail "
        f"(message_id: {respuesta_gmail['id']}) para: {', '.join(GMAIL_RECIPIENTS)}"
    )
    marcar_estado(
        state,
        year,
        month,
        status="request_sent" if solicitud else "sent",
        sent_at=datetime.now().isoformat(timespec="seconds"),
        gmail_message_id=respuesta_gmail["id"],
        solicitud_extraordinaria=solicitud,
        subject=subject,
        file=str(output_path),
        sheet_all=sheet_all_name,
    )
    guardar_estado(state)
    return output_path


def ejecutar_envio_pendiente(
    now: datetime | None = None,
    force: bool = False,
    solicitud: bool = False,
) -> None:
    now = now or datetime.now()
    if solicitud:
        target_year, target_month = now.year, now.month
        fecha_fin_override = now
    else:
        target_year, target_month = periodo_anterior(now)
        fecha_fin_override = None

    due_dt = period_due_datetime(target_year, target_month)
    state = cargar_estado()

    if now < due_dt and not solicitud:
        periodo = formatear_periodo(target_year, target_month)
        if not confirmar_envio_anticipado(periodo):
            print(f"Aun no toca enviar el reporte de {periodo}.")
            return

    reenvio_permitido = force or solicitud

    if estado_enviado(state, target_year, target_month) and not reenvio_permitido:
        print(
            f"Ya fue enviado el reporte de {formatear_periodo(target_year, target_month)}. "
            "No se envio otro correo. Usa --solicitud para envio extraordinario "
            "o --force para reenviarlo."
        )
        return

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    try:
        if solicitud:
            print(
                "Envio extraordinario solicitado para "
                f"{formatear_periodo(target_year, target_month)} "
                f"del {primer_dia_mes(target_year, target_month).strftime('%Y-%m-%d')} "
                f"al {fecha_fin_override.strftime('%Y-%m-%d')}."
            )
        archivo = procesar_periodo(
            target_year,
            target_month,
            state,
            solicitud=solicitud,
            fecha_fin_override=fecha_fin_override,
        )
        print(f"Archivo generado y enviado: {archivo}")
    except Exception as exc:
        marcar_estado(
            state,
            target_year,
            target_month,
            status="pending",
            last_attempt=now.isoformat(timespec="seconds"),
            last_error=str(exc),
        )
        guardar_estado(state)
        raise


def main() -> None:
    parser = argparse.ArgumentParser(description="Genera y envía el reporte Aventis por Gmail.")
    parser.add_argument(
        "--watch",
        action="store_true",
        help="Mantiene el proceso activo y reintenta cada cierto tiempo.",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Reenvia el periodo aunque aventis_state.json ya lo marque como enviado.",
    )
    parser.add_argument(
        "--solicitud",
        action="store_true",
        help="Envia el reporte por solicitud extraordinaria aunque ya figure como enviado.",
    )
    args = parser.parse_args()

    try:
        if args.watch:
            while True:
                try:
                    ejecutar_envio_pendiente(force=args.force, solicitud=args.solicitud)
                except Exception as exc:
                    print(f"Fallo el envio: {exc}")
                time.sleep(DAEMON_SLEEP_SECONDS)
        else:
            ejecutar_envio_pendiente(force=args.force, solicitud=args.solicitud)
    except ModuleNotFoundError as exc:
        if exc.name == "pyodbc":
            raise SystemExit(
                "Falta pyodbc en este Python. Ejecuta el script con el entorno "
                "que tenga instalada esa dependencia."
            ) from exc
        raise
    except Exception as exc:
        raise SystemExit(f"No se pudo completar el proceso: {exc}") from exc


if __name__ == "__main__":
    main()
