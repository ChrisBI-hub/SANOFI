"""
╔══════════════════════════════════════════════════════════════════════════════╗
║          MERGE & DEDUPLICACIÓN — TRACK AND TRACE ORGANON                   ║
╠══════════════════════════════════════════════════════════════════════════════╣
║  Descripción:                                                               ║
║    Consolida múltiples archivos bimestrales de Track and Trace en un        ║
║    único Excel limpio, eliminando operaciones duplicadas entre archivos.    ║
║                                                                             ║
║  Hojas procesadas:                                                          ║
║    - "Track and Trace"    -> operaciones productivas                        ║
║    - "T&T No Productivo"  -> operaciones no productivas                     ║
║                                                                             ║
║  Lógica de lectura:                                                         ║
║    - La fila 0 de cada hoja contiene instrucciones (se descarta)            ║
║    - La fila 1 contiene los encabezados reales                              ║
║    - Los datos van de la fila 2 en adelante                                 ║
║    - Las fórmulas se leen con su valor calculado (data_only)                ║
║    - Las filas completamente vacías se descartan automáticamente            ║
║                                                                             ║
║  Lógica de deduplicación (en orden de prioridad):                          ║
║    1. Columna CONCAT  (ID único por operación/producto)                     ║
║    2. Columna ID                                                            ║
║    3. PEDIMENTO COMPLETO + REFERENCE / TRAFICO                              ║
║    4. MAWB + DOC. DE TRANSPORTE + ETA  (T&T)                               ║
║       MAWB + HAWB / BL + ETA           (No Productivo)                     ║
║                                                                             ║
║  En caso de duplicado se conserva la fila del archivo más reciente.        ║
╚══════════════════════════════════════════════════════════════════════════════╝

REQUISITOS:
    pip install pandas openpyxl

USO:
    1. Pon este script en la misma carpeta que los archivos .xlsx, o
       edita FILES_TO_PROCESS con las rutas que necesites.
    2. Ejecuta:  python merge_track_and_trace.py
    3. Se genera un Excel:  Track_and_Trace_CONSOLIDADO_<fecha>.xlsx
"""

import sys
import logging
import warnings
from datetime import datetime
from pathlib import Path

import pandas as pd

warnings.filterwarnings("ignore")

# ══════════════════════════════════════════════════════════════════════════════
#  CONFIGURACION  <- edita esta sección
# ══════════════════════════════════════════════════════════════════════════════

# Carpeta donde están los .xlsx.
# None = misma carpeta donde está este script.
FOLDER = None

# Lista explícita de archivos (rutas absolutas o relativas).
# Si se deja vacía, el script toma TODOS los .xlsx de FOLDER automáticamente.
# El ORDEN importa: los archivos al final "ganan" en caso de duplicado.
FILES_TO_PROCESS = [
    # Ejemplos:
    # "Track_and_Trace_México_v2.xlsx",
    # "Track_and_Trace_México_2026_enero.xlsx",
    # "Track_and_Trace_México_2026_abril.xlsx",
]

# Hojas a procesar. Si una hoja no existe en algún archivo, se omite sin error.
TARGET_SHEETS = ["Track and Trace", "T&T No Productivo"]

# Índice (0-based) de la fila de encabezados dentro de la hoja.
# Fila 0 = instrucciones, Fila 1 = encabezados reales -> header=1
HEADER_ROW = 1

# Nombre del archivo de salida
OUTPUT_FILENAME = "Track_and_Trace_CONSOLIDADO_{}.xlsx".format(
    datetime.now().strftime("%Y%m%d_%H%M%S")
)

# Mínimo de columnas con datos para considerar una fila como real
MIN_COLS_WITH_DATA = 2

# ══════════════════════════════════════════════════════════════════════════════
#  CLAVES DE DEDUPLICACIÓN por hoja (en orden de preferencia)
# ══════════════════════════════════════════════════════════════════════════════
DEDUP_KEYS = {
    "Track and Trace": [
        ["CONCAT"],
        ["ID"],
        ["PEDIMENTO COMPLETO", "REFERENCE / TRAFICO"],
        ["MAWB", "DOC. DE TRANSPORTE", "ETA"],
    ],
    "T&T No Productivo": [
        ["CONCAT"],
        ["ID"],
        ["PEDIMENTO COMPLETO", "REFERENCE / TRAFICO"],
        ["MAWB", "HAWB / BL", "ETA"],
    ],
}

# ══════════════════════════════════════════════════════════════════════════════
#  LOGGING
# ══════════════════════════════════════════════════════════════════════════════
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════════════════
#  FUNCIONES
# ══════════════════════════════════════════════════════════════════════════════

def resolve_files(folder, explicit):
    """Devuelve la lista ordenada de archivos a procesar."""
    if explicit:
        paths = [Path(p) for p in explicit]
        missing = [p for p in paths if not p.exists()]
        if missing:
            raise FileNotFoundError("Archivos no encontrados: {}".format(missing))
        return paths

    base = Path(folder) if folder else Path(__file__).parent
    paths = sorted(base.glob("*.xlsx"), key=lambda p: p.stat().st_mtime)
    paths = [p for p in paths if "CONSOLIDADO" not in p.name]
    if not paths:
        raise FileNotFoundError("No se encontraron .xlsx en '{}'".format(base))
    return paths


def clean_col(name):
    """Normaliza un nombre de columna."""
    if name is None:
        return "__SIN_NOMBRE__"
    return " ".join(str(name).split())


def read_sheet(filepath, sheet_name):
    """Lee una hoja con pandas y devuelve un DataFrame limpio o None."""
    try:
        df = pd.read_excel(
            filepath,
            sheet_name=sheet_name,
            header=HEADER_ROW,
            engine="openpyxl",
        )
    except Exception as e:
        msg = str(e)
        if "not found" in msg.lower() or sheet_name in msg:
            log.warning("  ⚠  Hoja '%s' no encontrada en '%s' -> se omite",
                        sheet_name, filepath.name)
        else:
            log.error("  ✗  Error leyendo '%s'/'%s': %s", filepath.name, sheet_name, e)
        return None

    if df.empty:
        log.warning("  ⚠  Hoja '%s' en '%s' esta vacia -> se omite",
                    sheet_name, filepath.name)
        return None

    # Limpiar nombres de columnas
    df.columns = [clean_col(c) for c in df.columns]

    # Quitar columnas sin nombre (suelen ser auxiliares o de formato)
    df = df.loc[:, ~df.columns.str.fullmatch(r"__SIN_NOMBRE__(\.\d+)?")]

    # Quitar filas con menos de MIN_COLS_WITH_DATA valores reales
    df = df.dropna(thresh=MIN_COLS_WITH_DATA).reset_index(drop=True)

    if df.empty:
        log.warning("  ⚠  Hoja '%s' en '%s' quedo sin filas validas -> se omite",
                    sheet_name, filepath.name)
        return None

    # Agregar columnas de trazabilidad
    df["_ARCHIVO_ORIGEN"] = filepath.name
    df["_HOJA_ORIGEN"]    = sheet_name

    log.info("     ✓  %d filas  |  %d columnas", len(df), df.shape[1] - 2)
    return df


def find_dedup_key(df, sheet_name):
    """Retorna la primera combinación de claves disponible en el df."""
    available = set(df.columns)
    for key_combo in DEDUP_KEYS.get(sheet_name, []):
        norm = [clean_col(k) for k in key_combo]
        if all(k in available for k in norm):
            if df[norm].notna().any().any():
                return norm
    return None


def make_composite_key(df, key_cols):
    """Crea una Serie con la clave compuesta como string."""
    return (
        df[key_cols]
        .fillna("")
        .astype(str)
        .apply(lambda col: col.str.strip().str.upper())
        .apply(lambda row: "||".join(row.tolist()), axis=1)
    )


def deduplicate(df, key_cols):
    """Elimina duplicados conservando la última ocurrencia (archivo más reciente).
    
    Retorna (df_limpio, df_reporte_de_duplicados).
    """
    composite = make_composite_key(df, key_cols)

    # Filas sin clave -> no participan en dedup, pasan directo
    empty_mask = composite.str.fullmatch(r"(\|\|)*")
    df_keyed = df[~empty_mask].copy()
    df_nokey = df[empty_mask].copy()

    df_keyed["__CK__"] = composite[~empty_mask].values

    # Reporte: filas que aparecen más de una vez
    cnt = df_keyed["__CK__"].value_counts()
    dup_keys = cnt[cnt > 1].index
    dup_report = df_keyed[df_keyed["__CK__"].isin(dup_keys)].drop(columns="__CK__").copy()

    # Deduplicar: keep='last' -> gana el último archivo (el más reciente)
    df_keyed = df_keyed.drop_duplicates(subset="__CK__", keep="last")
    df_keyed = df_keyed.drop(columns="__CK__")

    result = pd.concat([df_keyed, df_nokey], ignore_index=True)

    removed = len(df) - len(result)
    if removed:
        log.info("     🔁  Duplicados eliminados: %d  (de %d -> %d filas)",
                 removed, len(df), len(result))
    return result, dup_report


def reorder_cols(df):
    """Mueve las columnas de metadatos al final."""
    meta = ["_ARCHIVO_ORIGEN", "_HOJA_ORIGEN"]
    other = [c for c in df.columns if c not in meta]
    meta_present = [m for m in meta if m in df.columns]
    return df[other + meta_present]


# ══════════════════════════════════════════════════════════════════════════════
#  FLUJO PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════

def main():
    log.info("══════════════════════════════════════════")
    log.info("  MERGE & DEDUPLICACION — TRACK AND TRACE")
    log.info("══════════════════════════════════════════")

    # 1. Resolver archivos
    try:
        files = resolve_files(FOLDER, FILES_TO_PROCESS)
    except FileNotFoundError as e:
        log.error("%s", e)
        sys.exit(1)

    log.info("Archivos a procesar (%d), orden de prioridad (ultimo = gana):", len(files))
    for f in files:
        log.info("  • %s", f.name)

    # 2. Procesar cada hoja
    results = {}

    for sheet_name in TARGET_SHEETS:
        log.info("")
        log.info("━━━ Hoja: '%s' ━━━", sheet_name)
        frames = []

        for filepath in files:
            log.info("  📂  %s", filepath.name)
            df = read_sheet(filepath, sheet_name)
            if df is not None:
                frames.append(df)

        if not frames:
            log.warning("  ⚠  Sin datos para '%s'", sheet_name)
            continue

        # Concatenar todos los archivos para esta hoja
        combined = pd.concat(frames, ignore_index=True, sort=False)
        log.info("")
        log.info("  📊  Total filas antes de deduplicar: %d", len(combined))

        # Deduplicar
        key_cols = find_dedup_key(combined, sheet_name)
        dup_report = pd.DataFrame()

        if key_cols:
            log.info("  🔑  Clave de deduplicación elegida: %s", key_cols)
            combined, dup_report = deduplicate(combined, key_cols)
        else:
            log.warning(
                "  ⚠  Ninguna columna clave disponible para '%s'. "
                "Se conservan todas las filas.", sheet_name
            )

        combined = reorder_cols(combined)
        log.info("  ✅  Filas finales: %d  |  Columnas: %d",
                 len(combined), combined.shape[1] - 2)

        results[sheet_name] = {
            "data":       combined,
            "key_cols":   key_cols,
            "dup_report": dup_report,
        }

    if not results:
        log.error("No se genero ningun resultado. Revisa los archivos de entrada.")
        sys.exit(1)

    # 3. Guardar Excel de salida
    base = Path(FOLDER) if FOLDER else Path(__file__).parent
    output_path = base / OUTPUT_FILENAME

    log.info("")
    log.info("━━━ Guardando Excel ━━━")
    log.info("  📄  %s", output_path)

    with pd.ExcelWriter(output_path, engine="openpyxl",
                        datetime_format="YYYY-MM-DD") as writer:

        # Hoja RESUMEN
        summary = []
        for sn, info in results.items():
            summary.append({
                "Hoja":                         sn,
                "Filas consolidadas":           len(info["data"]),
                "Columnas (sin metadatos)":     info["data"].shape[1] - 2,
                "Clave de deduplicación":       " + ".join(info["key_cols"])
                                                if info["key_cols"] else "N/A",
                "Filas duplicadas encontradas": len(info["dup_report"]),
            })
        pd.DataFrame(summary).to_excel(writer, sheet_name="RESUMEN", index=False)

        # Hojas de datos
        for sn, info in results.items():
            info["data"].to_excel(writer, sheet_name=sn[:31], index=False)

            if not info["dup_report"].empty:
                dup_sheet = ("DUPS " + sn)[:31]
                info["dup_report"].to_excel(writer, sheet_name=dup_sheet, index=False)
                log.info("  ℹ  Duplicados detallados en hoja: '%s'", dup_sheet)

    # 4. Resumen en consola
    log.info("")
    log.info("══════════════════════════════════════════")
    log.info("  ✅  COMPLETADO")
    log.info("  📁  Archivo: %s", output_path.name)
    for sn, info in results.items():
        log.info("      %-30s -> %d filas", sn, len(info["data"]))
    log.info("══════════════════════════════════════════")


if __name__ == "__main__":
    main()
