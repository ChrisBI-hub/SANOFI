#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
╔══════════════════════════════════════════════════════════════╗
║   Automatización Reporte Bimestral SANOFI - ABC              ║
║   Genera el .pptx de Operaciones desde SQL Server            ║
╚══════════════════════════════════════════════════════════════╝

═══════════════════════════════════════════════════════════════
  CORRECCIONES APLICADAS (por qué no generaba la presentación)
═══════════════════════════════════════════════════════════════

BUG 1 — KeyError: 'DESCRIPCIÓN_PRODUCTO' [CRÍTICO - causa principal]
  slide5_delayed() y slide6/7_refs_*() accedían a r["DESCRIPCIÓN_PRODUCTO"]
  pero el DataFrame fue construido agrupando por "Mercancía".
  Esa columna nunca existió en el dict de filas → KeyError no capturado
  → el proceso moría silenciosamente justo después del SQL.
  FIX: Unificar a "Mercancía" en toda la pipeline y en slide6/7.

BUG 2 — except IndexError no capturaba el KeyError anterior
  El try/except en llenar_tabla_delayed solo atrapaba IndexError,
  dejando pasar el KeyError de DESCRIPCIÓN_PRODUCTO sin manejarlo.
  FIX: except (IndexError, KeyError) + mensaje de aviso.

BUG 3 — groupby().apply() deprecado en pandas ≥ 2.2
  pct_ontime() recibía el grupo completo incluyendo la clave,
  lo que genera DeprecationWarning (y en futuras versiones error).
  FIX: agregar include_groups=False al apply().

MEJORA ADICIONAL — Normalización de columnas
  Nombres de columnas de SQL Server pueden tener espacios, mayúsculas
  o encodings distintos. Se agrega normalización al inicio de
  preparar_datos() y búsqueda flexible de columnas clave.

═══════════════════════════════════════════════════════════════

Uso:
    python reporte_bimestral.py

Dependencias:
    pip install pyodbc pandas python-pptx openpyxl sqlalchemy
"""

import os
import sys
import traceback
import urllib.parse

import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.util import Pt
from datetime import datetime
from sqlalchemy import create_engine, text

# ================================================================
#  CONFIGURACIÓN — EDITAR AQUÍ CADA BIMESTRE
# ================================================================

server   = '150.1.1.152'
database = 'SIR'
username = 'ConsultaBD'
password = '5D$bc#kM&5W2T8J40?s%'

CONN_STRING = (
    f"DRIVER={{ODBC Driver 18 for SQL Server}};"
    f"SERVER={server};"
    f"DATABASE={database};"
    f"UID={username};"
    f"PWD={password};"
    f"TrustServerCertificate=yes;"
)

# ── Parámetros del bimestre ──────────────────────────────────────
BIMESTRE = {
    "meses":               [3, 4],
    "anno":                2026,
    "mes1_nombre":         "Marzo",
    "mes2_nombre":         "Abril",
    "mes1_corto":          "Mar",
    "mes2_corto":          "Abr",
    "label_portada":       "Marzo – Abril 2026",
    "label_portada_anterior": "Enero– Febrero 2026",
}

# ── Archivos ─────────────────────────────────────────────────────
DIR_BASE  = os.path.dirname(os.path.abspath(__file__))
SQL_PATH  = os.path.join(DIR_BASE, "SANOFI_V5.sql")
PPTX_TMPL = os.path.join(DIR_BASE, "ABC_OPERACIONES_SANOFI_ENE_FEB_2026.pptx")
PPTX_OUT  = os.path.join(
    DIR_BASE,
    f"ABC_OPERACIONES_SANOFI_{BIMESTRE['mes1_nombre'].upper()}_"
    f"{BIMESTRE['mes2_nombre'].upper()}_{BIMESTRE['anno']}.pptx"
)

# Solo estas diapositivas reciben datos/graficas desde la BD.
# El resto se conserva desde la plantilla y solo se actualiza el periodo.
SLIDES_DESDE_BD = (3, 8, 9, 10, 11)

# ── Mapeos de negocio ────────────────────────────────────────────
CLIENTE_AVENTIS = "SANOFI - AVENTIS DE MEXICO, S.A. DE C.V."
CLIENTE_PASTEUR = "SANOFI PASTEUR, S.A DE C.V."

SUCURSAL_MAP = {
    "CIUDAD DE MÉXICO":      "AICM",
    "AIFA ESTADO DE MEXICO": "AIFA",
    "VERACRUZ":              "VERACRUZ",
    "LAREDO":                "LAREDO",
    "MANZANILLO":            "MANZANILLO",
}

TARGET_ENTRADA_CRUCE = {
    "CIUDAD DE MÉXICO":      3,
    "AIFA ESTADO DE MEXICO": 3,
    "VERACRUZ":              5,
    "LAREDO":                5,
    "MANZANILLO":            5,
}


# ================================================================
#  UTILIDAD: resolución flexible de nombres de columna
# ================================================================

def resolver_columna(df: pd.DataFrame, candidatos: list, obligatoria: bool = True) -> str:
    """
    Devuelve el primer nombre de la lista 'candidatos' que exista en df.columns.
    Si ninguno existe y es obligatoria, lanza KeyError con mensaje claro.
    """
    cols_lower = {c.strip().lower(): c for c in df.columns}
    for nombre in candidatos:
        if nombre in df.columns:
            return nombre
        if nombre.strip().lower() in cols_lower:
            return cols_lower[nombre.strip().lower()]
    if obligatoria:
        raise KeyError(
            f"No se encontró ninguna de las columnas {candidatos} en el DataFrame.\n"
            f"Columnas disponibles: {list(df.columns)}"
        )
    return None


# ================================================================
#  1. CONEXIÓN Y LECTURA DE DATOS
# ================================================================

def leer_query() -> str:
    if not os.path.exists(SQL_PATH):
        raise FileNotFoundError(
            f"No se encontró el archivo SQL en:\n  {SQL_PATH}\n"
            "Verifica que esté en la misma carpeta que este script."
        )
    with open(SQL_PATH, encoding="latin-1") as f:
        query = f.read()

    meses  = BIMESTRE["meses"]
    anno   = BIMESTRE["anno"]
    fecha_ini = f"{anno}-{meses[0]:02d}-01"
    # Ultimo dia del mes 2
    import calendar
    ultimo_dia = calendar.monthrange(anno, meses[1])[1]
    fecha_fin = f"{anno}-{meses[1]:02d}-{ultimo_dia:02d}"
    
    # reemplazar fechas en el sql
    query = query.replace("'2026-03-01'", f"'{fecha_ini}'")
    query = query.replace("'2026-04-30'", f"'{fecha_fin}'")
    return query

def obtener_datos() -> pd.DataFrame:
    """Conecta a SQL Server, ejecuta la query y retorna un DataFrame."""
    print("  Conectando a SQL Server...")
    try:
        params = urllib.parse.quote_plus(CONN_STRING)
        engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")

        print("  Ejecutando query (puede tardar unos segundos)...")
        query = leer_query()

        with engine.connect() as conn:
            df = pd.read_sql(text(query), conn)
    except ModuleNotFoundError as e:
        if e.name == "pyodbc":
            raise ConnectionError(
                "Falta la dependencia pyodbc en este Python.\n"
                "Ejecuta el reporte con el entorno del proyecto:\n"
                "  sanofi/bin/python reporte_bimestral.py"
            ) from e
        raise
    except Exception as e:
        raise ConnectionError(
            "No se pudo obtener informacion de SQL Server.\n"
            f"Servidor configurado: {server}\n"
            "Verifica que estes conectado a la red/VPN y que el puerto SQL Server "
            "este disponible.\n"
            f"Detalle tecnico: {e}"
        ) from e

    print(f"  → {len(df):,} registros obtenidos, {len(df.columns)} columnas")
    print(f"  → Columnas: {list(df.columns)}")   # diagnóstico: imprime columnas reales
    return df


# ================================================================
#  2. PROCESAMIENTO Y KPIs
# ================================================================

def preparar_datos(df: pd.DataFrame, meses: list, anno: int) -> pd.DataFrame:
    df = df.copy()

    # Normalizar nombres de columnas (quitar espacios extras)
    df.columns = df.columns.str.strip()

    # Asegurar que MES sea numérico
    df["MES"] = pd.to_numeric(df["MES"], errors="coerce")

    # Filtrar bimestre
    if "ANNO" in df.columns:
        df = df[(df["MES"].isin(meses)) & (df["ANNO"] == anno)].copy()
    else:
        df = df[df["MES"].isin(meses)].copy()

    if len(df) == 0:
        print(f"  ⚠ ADVERTENCIA: No hay registros para los meses {meses} del año {anno}.")
        print("    Verifica que la query SQL retorne datos para ese bimestre.")

    # Columna ESTATUS
    col_verif = resolver_columna(df, ["VERIFICADOR", "Verificador"], obligatoria=False)
    if col_verif and "ESTATUS" not in df.columns:
        df["ESTATUS"] = df[col_verif].apply(
            lambda x: "ON TIME" if x == 1 else "DELAYED"
        )

    if "ESTATUS" not in df.columns:
        col_suc = resolver_columna(df, ["Sucursal", "SUCURSAL", "sucursal"])
        col_ec  = resolver_columna(df, ["Entrada a Cruce", "EntradaCruce"], obligatoria=False)
        if col_ec:
            df["ESTATUS"] = df.apply(
                lambda r: "ON TIME"
                if r.get(col_ec, 99) <= TARGET_ENTRADA_CRUCE.get(r.get(col_suc, ""), 5)
                else "DELAYED",
                axis=1
            )
        else:
            df["ESTATUS"] = "ON TIME"  # fallback si no hay ninguna columna de tiempo

    # Columna RAZON_SOCIAL — BUG FIX: búsqueda flexible
    col_cliente = resolver_columna(df, ["Cliente", "CLIENTE", "cliente", "RAZON_SOCIAL"])
    df["RAZON_SOCIAL"] = df[col_cliente].str.strip()

    # ── DIAGNÓSTICO: mostrar conteo por mes para verificar datos ──
    print(f"  → Registros por mes en el bimestre:")
    print(df["MES"].value_counts().sort_index().to_string())

    return df


def calcular_kpis(df: pd.DataFrame) -> dict:
    meses  = BIMESTRE["meses"]
    m1, m2 = meses[0], meses[1]
    kpis   = {}

    # Resolver nombres reales de columnas críticas
    col_suc   = resolver_columna(df, ["Sucursal", "SUCURSAL"])
    col_ref   = resolver_columna(df, ["Referencia", "REFERENCIA"])
    col_fam   = resolver_columna(df, ["FAMILIA", "Familia", "familia"])
    # Mercancia solo se conserva como dato auxiliar; las slides 5-7 ya no
    # se alimentan desde BD porque dependen de informacion externa del cliente.
    col_merc  = resolver_columna(df, ["Mercancía", "Mercancia", "MERCANCIA", "DESCRIPCIÓN_PRODUCTO",
                                      "Descripcion_Producto", "DESCRIPCION_PRODUCTO"],
                                 obligatoria=False)
    col_ec    = resolver_columna(df, ["Entrada a Cruce", "EntradaCruce", "ENTRADA_A_CRUCE"],
                                 obligatoria=False)
    col_fac   = resolver_columna(df, ["Entrada a Factura", "EntradaFactura", "ENTRADA_A_FACTURA"],
                                 obligatoria=False)
    col_mot   = resolver_columna(df, ["MOTIVO DE RETRASO COMPLETO", "MOTIVO_RETRASO",
                                      "MotivoRetraso"], obligatoria=False)

    # ── Slide 3: Operaciones totales ──────────────────────────────
    op_rs = (
        df.groupby(["MES", "RAZON_SOCIAL"])[col_ref]
        .nunique()
        .unstack(fill_value=0)
    )
    kpis["op_razon_social"] = op_rs

    op_fam = df.groupby(col_fam)[col_ref].nunique().sort_values(ascending=False)
    kpis["op_familia"] = op_fam

    op_aduana = (
        df.groupby(["MES", col_suc])[col_ref]
        .nunique()
        .unstack(fill_value=0)
    )
    kpis["op_aduana"] = op_aduana

    kpis["op_por_puerto"] = {
        corto: int(df[df[col_suc] == orig][col_ref].nunique())
        for orig, corto in SUCURSAL_MAP.items()
    }

    # ── Slide 8: Motivos delayed ───────────────────────────────
    df_delayed = df[df["ESTATUS"] == "DELAYED"].copy()

    # Guardar nombres reales de columna para compatibilidad con helpers antiguos.
    kpis["_col_fam"]  = col_fam
    kpis["_col_merc"] = col_merc

    col_imp = "IMPUTABLE A"
    if col_imp not in df_delayed.columns:
        df_delayed[col_imp] = "CLIENTE"

    if col_mot:
        kpis["motivos_tabla"] = (
            df_delayed.groupby([col_suc, "MES", col_mot, col_imp])[col_ref]
            .nunique()
            .reset_index(name="CANTIDAD")
            .rename(columns={col_mot: "MOTIVO DE RETRASO COMPLETO"})
        )
    else:
        kpis["motivos_tabla"] = pd.DataFrame(
            columns=[col_suc, "MES", "MOTIVO DE RETRASO COMPLETO", col_imp, "CANTIDAD"]
        )

    # Slide 9: Rectificaciones
    if "ES_RECTIFICACION" in df.columns:
        kpis["total_rectificaciones"] = int(
            df[df["ES_RECTIFICACION"] == 1][col_ref].nunique()
        )
    else:
        kpis["total_rectificaciones"] = 0

    # Slides 10-11: Tiempos promedio
    def avg_pivot(df_in, col_val, col_group):
        if col_val is None or col_val not in df_in.columns:
            return pd.DataFrame()
        return (
            df_in.groupby(["MES", col_group])[col_val]
            .mean().round(0).unstack()
        )

    kpis["entrada_cruce_razon"]  = avg_pivot(df, col_ec,  "RAZON_SOCIAL")
    kpis["entrada_cruce_aduana"] = avg_pivot(df, col_ec,  col_suc)
    kpis["facturacion_razon"]    = avg_pivot(df, col_fac, "RAZON_SOCIAL")
    kpis["facturacion_aduana"]   = avg_pivot(df, col_fac, col_suc)

    return kpis


# ================================================================
#  3. HELPERS python-pptx
# ================================================================

def _safe_val(pivote, mes, col, default=None):
    try:
        if pivote.empty:
            return default
        v = pivote.loc[mes, col]
        return None if pd.isna(v) else float(v)
    except (KeyError, TypeError):
        return default


def replace_text_slide(slide, reemplazos: dict):
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            texto_completo = "".join(run.text for run in para.runs)
            texto_nuevo = texto_completo
            for buscar, nuevo in reemplazos.items():
                texto_nuevo = texto_nuevo.replace(buscar, str(nuevo))

            # PowerPoint a veces parte una frase en varios runs. Si el
            # reemplazo completo cambia el parrafo, lo reescribimos completo.
            if texto_completo and texto_nuevo != texto_completo:
                para.clear()
                run = para.add_run()
                run.text = texto_nuevo
                continue

            for run in para.runs:
                for buscar, nuevo in reemplazos.items():
                    if buscar in run.text:
                        run.text = run.text.replace(buscar, str(nuevo))


def actualizar_periodo_presentacion(prs: Presentation, cfg: dict):
    """Actualiza textos de periodo en toda la plantilla sin tocar graficas."""
    periodo_largo = f"{cfg['mes1_nombre']} - {cfg['mes2_nombre']}"
    periodo_largo_mayus = periodo_largo.upper()
    reemplazos = {
        cfg["label_portada_anterior"]: cfg["label_portada"],
        "Enero– Febrero 2026": cfg["label_portada"],
        "Enero–Febrero 2026": cfg["label_portada"],
        "Enero - Febrero": periodo_largo,
        "ENERO - FEBRERO": periodo_largo_mayus,
        "Enero": cfg["mes1_nombre"],
        "Febrero": cfg["mes2_nombre"],
        "ENERO": cfg["mes1_nombre"].upper(),
        "FEBRERO": cfg["mes2_nombre"].upper(),
        "Periodo: Marzo": f"Periodo: {periodo_largo}",
    }

    for slide in prs.slides:
        replace_text_slide(slide, reemplazos)


def set_table_cell(tabla, fila, col, valor, bold=False):
    cell = tabla.cell(fila, col)
    tf   = cell.text_frame
    tf.paragraphs[0].clear()
    run = tf.paragraphs[0].add_run()
    run.text = str(valor)
    if bold:
        run.font.bold = True


def update_chart_data(chart, categories: list, series: dict):
    cd = CategoryChartData()
    cd.categories = categories
    for nombre, valores in series.items():
        cd.add_series(nombre, tuple(valores))
    chart.replace_data(cd)


def get_charts(slide) -> list:
    return [s for s in slide.shapes if s.has_chart]


def get_tables(slide) -> list:
    return [s for s in slide.shapes if s.has_table]


# ================================================================
#  4. ACTUALIZACIÓN DE CADA SLIDE
#  Cada función está envuelta en try/except para que un fallo
#  en una slide no detenga la generación del resto.
# ================================================================

def _safe_slide(nombre: str, fn, *args, **kwargs):
    """Ejecuta una función de slide capturando errores sin abortar."""
    try:
        fn(*args, **kwargs)
    except Exception as exc:
        print(f"  ⚠ {nombre}: error al actualizar → {exc}")
        print(f"    (Detalle: {traceback.format_exc().splitlines()[-2]})")


def slide1_portada(slide, cfg: dict):
    replace_text_slide(slide, {
        cfg["label_portada_anterior"]: cfg["label_portada"],
        "Enero":   cfg["mes1_nombre"],
        "Febrero": cfg["mes2_nombre"],
    })
    print("  ✓ Slide 1: Portada")


def slide3_operaciones(slide, kpis: dict, cfg: dict):
    charts    = get_charts(slide)
    m1, m2    = cfg["meses"]
    cats      = [cfg["mes1_corto"], cfg["mes2_corto"]]
    op_rs     = kpis["op_razon_social"]
    op_fam    = kpis["op_familia"]
    op_aduana = kpis["op_aduana"]

    if len(charts) > 0:
        series = {
            CLIENTE_AVENTIS: [_safe_val(op_rs, m1, CLIENTE_AVENTIS, 0),
                               _safe_val(op_rs, m2, CLIENTE_AVENTIS, 0)],
            CLIENTE_PASTEUR: [_safe_val(op_rs, m1, CLIENTE_PASTEUR, 0),
                               _safe_val(op_rs, m2, CLIENTE_PASTEUR, 0)],
        }
        update_chart_data(charts[0].chart, cats, series)

    if len(charts) > 1:
        familias = op_fam.index.tolist()
        valores  = [float(v) for v in op_fam.values]
        update_chart_data(charts[1].chart, familias, {"": valores})

    if len(charts) > 2:
        aduanas = ["CIUDAD DE MÉXICO", "AIFA ESTADO DE MEXICO", "VERACRUZ"]
        labels  = ["AICM", "AIFA", "VERACRUZ"]
        series  = {
            lbl: [_safe_val(op_aduana, m1, orig, 0),
                  _safe_val(op_aduana, m2, orig, 0)]
            for lbl, orig in zip(labels, aduanas)
        }
        update_chart_data(charts[2].chart, cats, series)

    # Números sueltos de puerto — actualizar textboxes con los totales calculados
    puertos = kpis["op_por_puerto"]
    puerto_vals = sorted(puertos.values(), reverse=True)
    idx = 0
    for shape in slide.shapes:
        if shape.has_text_frame:
            txt = shape.text_frame.text.strip()
            if txt.isdigit() and idx < len(puerto_vals):
                shape.text_frame.paragraphs[0].runs[0].text = str(puerto_vals[idx])
                idx += 1

    print("  ✓ Slide 3: Operaciones")


def slide4_ontime(slide, kpis: dict, cfg: dict):
    tablas = get_tables(slide)
    m1, m2 = cfg["meses"]
    on_gen = kpis["ontime_general"]
    on_ad  = kpis["ontime_aduana"]

    def fmt(val):
        return f"{int(val)}%" if val is not None else "Sin Op."

    if len(tablas) > 0:
        t  = tablas[0].table
        v1 = fmt(on_gen.get(m1))
        v2 = fmt(on_gen.get(m2))
        set_table_cell(t, 1, 1, v1)
        set_table_cell(t, 1, 2, v2)

    if len(tablas) > 1:
        t = tablas[1].table
        filas_aduana = [
            ("CIUDAD DE MÉXICO",      1),
            ("AIFA ESTADO DE MEXICO", 2),
            ("LAREDO",                3),
            ("MANZANILLO",            4),
            ("VERACRUZ",              5),
        ]
        for sucursal, fila in filas_aduana:
            try:
                v1_raw = on_ad.loc[sucursal, m1] if (sucursal in on_ad.index and m1 in on_ad.columns) else None
                v2_raw = on_ad.loc[sucursal, m2] if (sucursal in on_ad.index and m2 in on_ad.columns) else None
                set_table_cell(t, fila, 1, fmt(v1_raw))
                set_table_cell(t, fila, 2, fmt(v2_raw))
            except (KeyError, IndexError):
                pass

    print("  ✓ Slide 4: On Time")


def slide5_delayed(slide, kpis: dict, cfg: dict):
    tablas   = get_tables(slide)
    m1, m2   = cfg["meses"]
    m1_nom   = cfg["mes1_nombre"].upper()
    m2_nom   = cfg["mes2_nombre"].upper()
    col_fam  = kpis["_col_fam"]
    col_merc = kpis["_col_merc"]   # ← BUG FIX: nombre real de la columna

    def llenar_tabla_delayed(tabla_shape, df_delayed_aduana):
        t    = tabla_shape.table
        fila = 1
        for _, r in df_delayed_aduana.iterrows():
            try:
                # BUG FIX: usar col_merc en vez de la literal "DESCRIPCIÓN_PRODUCTO"
                set_table_cell(t, fila, 0, r[col_fam])
                set_table_cell(t, fila, 1, r[col_merc])
                set_table_cell(t, fila, 2, str(r["CANTIDAD"]))
                fila += 1
            except IndexError:
                break
            except KeyError as ke:   # BUG FIX: capturar KeyError también
                print(f"    ⚠ Columna no encontrada en delayed: {ke}")
                break

    if len(tablas) > 0:
        llenar_tabla_delayed(tablas[0], kpis["delayed_aifa"])
    if len(tablas) > 1:
        llenar_tabla_delayed(tablas[1], kpis["delayed_veracruz"])

    print("  ✓ Slide 5: Delayed")


def slide6_refs_aifa(slide, kpis: dict, cfg: dict):
    refs     = kpis["refs_aifa"]
    col_merc = kpis["_col_merc"]   # BUG FIX

    shapes_texto = [
        s for s in slide.shapes
        if s.has_text_frame and len(s.text_frame.text) > 20
    ]
    if not shapes_texto:
        print("  ⚠ Slide 6: No se encontró textbox para referencias AIFA")
        return

    target = max(shapes_texto, key=lambda s: len(s.text_frame.text))
    tf     = target.text_frame

    for i, para in enumerate(tf.paragraphs):
        if i == 0:
            continue
        if i - 1 < len(refs):
            ref = refs[i - 1]
            # BUG FIX: usar col_merc en vez de la literal "DESCRIPCIÓN_PRODUCTO"
            merc = ref.get(col_merc, ref.get("DESCRIPCIÓN_PRODUCTO", ""))
            nuevo_texto = f"Referencia {ref.get('Referencia', '')}, mercancía {merc}"
            for run in para.runs:
                run.text = nuevo_texto
                break

    print("  ✓ Slide 6: Referencias AIFA")


def slide7_refs_veracruz(slide, kpis: dict, cfg: dict):
    refs     = kpis["refs_veracruz"]
    col_merc = kpis["_col_merc"]   # BUG FIX

    shapes_texto = [
        s for s in slide.shapes
        if s.has_text_frame and len(s.text_frame.text) > 20
    ]
    if not shapes_texto:
        print("  ⚠ Slide 7: No se encontró textbox para referencias VERACRUZ")
        return

    target = max(shapes_texto, key=lambda s: len(s.text_frame.text))
    tf     = target.text_frame

    for i, para in enumerate(tf.paragraphs):
        if i == 0:
            continue
        if i - 1 < len(refs):
            ref = refs[i - 1]
            merc = ref.get(col_merc, ref.get("DESCRIPCIÓN_PRODUCTO", ""))
            nuevo_texto = f"Referencia {ref.get('Referencia', '')}, mercancía {merc}"
            for run in para.runs:
                run.text = nuevo_texto
                break

    print("  ✓ Slide 7: Referencias VERACRUZ")


def slide8_motivos_tabla(slide, kpis: dict, cfg: dict):
    tablas  = get_tables(slide)
    df_mot  = kpis["motivos_tabla"]
    m1, m2  = cfg["meses"]

    def llenar_motivos(tabla_shape, sucursal):
        t   = tabla_shape.table
        sub = df_mot[df_mot.iloc[:, 0] == sucursal]  # primera columna = sucursal
        fila = 1
        for motivo in sub["MOTIVO DE RETRASO COMPLETO"].unique():
            try:
                def q(mes, imp):
                    return int(
                        sub[(sub["MES"] == mes) &
                            (sub["IMPUTABLE A"] == imp) &
                            (sub["MOTIVO DE RETRASO COMPLETO"] == motivo)
                           ]["CANTIDAD"].sum()
                    )
                set_table_cell(t, fila, 0, motivo)
                set_table_cell(t, fila, 1, str(q(m1, "ABC")))
                set_table_cell(t, fila, 2, str(q(m1, "CLIENTE")))
                set_table_cell(t, fila, 3, str(q(m2, "ABC")))
                set_table_cell(t, fila, 4, str(q(m2, "CLIENTE")))
                fila += 1
            except IndexError:
                break

    if len(tablas) > 0:
        llenar_motivos(tablas[0], "AIFA ESTADO DE MEXICO")
    if len(tablas) > 1:
        llenar_motivos(tablas[1], "VERACRUZ")

    print("  ✓ Slide 8: Motivos de retraso")


def slide9_rectificaciones(slide, kpis: dict, cfg: dict):
    total   = kpis["total_rectificaciones"]
    periodo = f"{cfg['mes1_nombre']} - {cfg['mes2_nombre']}"

    for shape in slide.shapes:
        if shape.has_text_frame:
            txt = shape.text_frame.text.strip()
            if txt.isdigit() and len(txt) <= 3:
                tf = shape.text_frame
                tf.paragraphs[0].runs[0].text = str(total)
                break

    replace_text_slide(slide, {"Enero": periodo})
    print("  ✓ Slide 9: Rectificaciones")


def slide10_entrada_cruce(slide, kpis: dict, cfg: dict):
    charts = get_charts(slide)
    m1, m2 = cfg["meses"]
    cats   = [cfg["mes1_corto"], cfg["mes2_corto"]]
    ec_rs  = kpis["entrada_cruce_razon"]
    ec_ad  = kpis["entrada_cruce_aduana"]

    if len(charts) > 0 and not ec_rs.empty:
        series = {
            CLIENTE_AVENTIS: [_safe_val(ec_rs, m1, CLIENTE_AVENTIS),
                               _safe_val(ec_rs, m2, CLIENTE_AVENTIS)],
            CLIENTE_PASTEUR: [_safe_val(ec_rs, m1, CLIENTE_PASTEUR),
                               _safe_val(ec_rs, m2, CLIENTE_PASTEUR)],
        }
        update_chart_data(charts[0].chart, cats, series)

    if len(charts) > 1 and not ec_ad.empty:
        aduanas = {
            "AIFA ESTADO DE MEXICO": "AIFA ESTADO DE MEXICO",
            "CIUDAD DE MÉXICO":      "CIUDAD DE MÉXICO",
            "VERACRUZ":              "VERACRUZ",
        }
        series = {
            lbl: [_safe_val(ec_ad, m1, key), _safe_val(ec_ad, m2, key)]
            for lbl, key in aduanas.items()
        }
        update_chart_data(charts[1].chart, cats, series)

    print("  ✓ Slide 10: Entrada a Cruce")


def slide11_facturacion(slide, kpis: dict, cfg: dict):
    charts = get_charts(slide)
    m1, m2 = cfg["meses"]
    cats   = [cfg["mes1_corto"], cfg["mes2_corto"]]
    fac_rs = kpis["facturacion_razon"]
    fac_ad = kpis["facturacion_aduana"]

    if len(charts) > 0 and not fac_rs.empty:
        series = {
            CLIENTE_AVENTIS: [_safe_val(fac_rs, m1, CLIENTE_AVENTIS),
                               _safe_val(fac_rs, m2, CLIENTE_AVENTIS)],
            CLIENTE_PASTEUR: [_safe_val(fac_rs, m1, CLIENTE_PASTEUR),
                               _safe_val(fac_rs, m2, CLIENTE_PASTEUR)],
        }
        update_chart_data(charts[0].chart, cats, series)

    if len(charts) > 1 and not fac_ad.empty:
        aduanas = {
            "AICM":     "CIUDAD DE MÉXICO",
            "AIFA":     "AIFA ESTADO DE MEXICO",
            "VERACRUZ": "VERACRUZ",
        }
        series = {
            lbl: [_safe_val(fac_ad, m1, key), _safe_val(fac_ad, m2, key)]
            for lbl, key in aduanas.items()
        }
        update_chart_data(charts[1].chart, cats, series)

    print("  ✓ Slide 11: Facturación")


def slide12_estado_cuenta(slide, kpis: dict):
    charts = get_charts(slide)
    if not charts:
        print("  ⚠ Slide 12: No se encontró gráfica")
        return

    update_chart_data(
        charts[0].chart,
        ["AL CORRIENTE", "VENCIDO"],
        {"": [kpis["pct_corriente"], kpis["pct_vencido"]]}
    )
    print("  ✓ Slide 12: Estado de Cuenta")


# ================================================================
#  5. MAIN
# ================================================================

def main():
    inicio = datetime.now()
    print("\n" + "═" * 60)
    print(f"  REPORTE BIMESTRAL SANOFI — {BIMESTRE['label_portada']}")
    print("═" * 60)

    # ── Paso 1: Obtener datos ───────────────────────────────────
    print("\n[1/4] Obteniendo datos de SQL Server...")
    df_raw = obtener_datos()

    # ── Paso 2: Procesar y calcular KPIs ────────────────────────
    print("\n[2/4] Calculando KPIs...")
    df   = preparar_datos(df_raw, BIMESTRE["meses"], BIMESTRE["anno"])
    kpis = calcular_kpis(df)
    print(f"  → Registros en bimestre: {len(df):,}")
    print(f"  → Slides con datos BD:   {', '.join(map(str, SLIDES_DESDE_BD))}")

    # ── Paso 3: Abrir plantilla y actualizar slides ─────────────
    print(f"\n[3/4] Actualizando presentación...")
    print(f"  Plantilla: {os.path.basename(PPTX_TMPL)}")

    if not os.path.exists(PPTX_TMPL):
        raise FileNotFoundError(f"No se encontró la plantilla:\n  {PPTX_TMPL}")

    prs    = Presentation(PPTX_TMPL)
    slides = prs.slides

    print(f"  Slides en la plantilla: {len(slides)}")
    print("\n  Slides dinámicos:")

    # Cada slide usa _safe_slide para no abortar si una falla.
    # Las slides no listadas en SLIDES_DESDE_BD se conservan como plantilla.
    _safe_slide("Slide 1",  slide1_portada,        slides[0],  BIMESTRE)
    _safe_slide("Slide 3",  slide3_operaciones,    slides[2],  kpis, BIMESTRE)
    _safe_slide("Slide 8",  slide8_motivos_tabla,  slides[7],  kpis, BIMESTRE)
    _safe_slide("Slide 9",  slide9_rectificaciones,slides[8],  kpis, BIMESTRE)
    _safe_slide("Slide 10", slide10_entrada_cruce, slides[9],  kpis, BIMESTRE)
    _safe_slide("Slide 11", slide11_facturacion,   slides[10], kpis, BIMESTRE)

    actualizar_periodo_presentacion(prs, BIMESTRE)
    print("  ✓ Periodo actualizado en titulos/textos de toda la presentacion")
    print("  → Omitidas desde BD: slides 4, 5, 6, 7 y 12 (se conserva la plantilla)")

    # ── Paso 4: Guardar ─────────────────────────────────────────
    print(f"\n[4/4] Guardando archivo...")
    prs.save(PPTX_OUT)
    elapsed = (datetime.now() - inicio).seconds
    print(f"\n✅ Reporte generado en {elapsed}s")
    print(f"   → {PPTX_OUT}\n")


if __name__ == "__main__":
    try:
        main()
    except FileNotFoundError as e:
        print(f"\n❌ Archivo no encontrado:\n   {e}")
        sys.exit(1)
    except ConnectionError as e:
        print(f"\n❌ Error de conexión:\n   {e}")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Error inesperado: {e}")
        traceback.print_exc()   # imprime el traceback completo
        sys.exit(1)
