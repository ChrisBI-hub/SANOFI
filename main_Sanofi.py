#!/usr/bin/env python3
"""
main.py — Orquestador del Reporte Mensual SANOFI
==================================================

Flujo completo:
  1. organon_2_1.py       → REPORTE_ADUANAL_SANOFI_{año}_{mes:02d}.xlsx
  2. oñate_extraccion.py  → REPORTE_MAESTRO_{año}_{mes:02d}.xlsx  (refs MNS)
  3. laredo.py            → PDFs en Descargas_LT/               (refs LT)
     └── analisis.py      → ANALISIS_LT_{timestamp}.xlsx
  4. union_2_1.py         → Reporte_SANOFI.xlsx
  5. revision_profunda.py → Reporte_SANOFI_{mes:02d}.xlsx
  6a. [Revisado]   limpieza.py → separador_pestañas.py → correo.py
  6b. [No revisado]                                  correo_revision.py

Uso:
  python main.py              # Pide interactivamente los años y meses a procesar
  python main.py --desde 4    # Inicia cada periodo desde el paso 4 (útil si algo falló)

Selección de periodos:
  Al arrancar, el script pide:
    1. Año(s) — ej. "2026" o "2023,2024" o "2023-2025"
    2. Mes(es) por NÚMERO — ej. "3" o "1,2,5" o "1-6", o "todos" para los 12
  Si se elige más de un periodo, genera una sola extracción SQL general
  por rango calendario, sin segmentar por mes.
"""

import os
import re
import sys
import importlib.util
import calendar
import argparse
from datetime import date, timedelta

# ===========================================================================
# CONFIGURACIÓN CENTRAL
# ===========================================================================

PATH_BASE = "/home/christian/Documentos/SANOFI"

MESES_ES = {
    1: "Enero",    2: "Febrero",   3: "Marzo",    4: "Abril",
    5: "Mayo",     6: "Junio",     7: "Julio",    8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}

# ===========================================================================
# UTILIDADES
# ===========================================================================

def titulo(texto, nivel=1):
    if nivel == 1:
        print(f"\n{'═' * 62}")
        print(f"  {texto}")
        print(f"{'═' * 62}")
    else:
        print(f"\n  ── {texto} ──")


def detectar_periodo():
    """Retorna (mes, año) del mes inmediatamente anterior al actual."""
    hoy = date.today()
    primer_dia = hoy.replace(day=1)
    mes_anterior = primer_dia - timedelta(days=1)
    return mes_anterior.month, mes_anterior.year


def parse_años(raw: str) -> list[int]:
    """
    Parsea años desde texto. Soporta:
      "2026"          → [2026]
      "2023,2024"     → [2023, 2024]
      "2023-2025"     → [2023, 2024, 2025]
      "2023,2025-2026" → [2023, 2025, 2026]
    """
    años = set()
    for parte in raw.split(","):
        parte = parte.strip()
        if not parte:
            continue
        if "-" in parte:
            a, b = parte.split("-")
            a, b = int(a.strip()), int(b.strip())
            años.update(range(min(a, b), max(a, b) + 1))
        else:
            años.add(int(parte))
    if not años:
        raise ValueError("No se especificó ningún año.")
    return sorted(años)


def parse_meses(raw: str) -> list[int]:
    """
    Parsea meses por NÚMERO desde texto. Soporta:
      "todos" / "all" / "*"  → los 12 meses
      "3"                    → [3]
      "1,2,5"                → [1, 2, 5]
      "1-6"                  → [1, 2, 3, 4, 5, 6]
    """
    raw = raw.strip().lower()
    if raw in ("todos", "all", "*"):
        return list(range(1, 13))

    meses = set()
    for parte in raw.split(","):
        parte = parte.strip()
        if not parte:
            continue
        if "-" in parte:
            a, b = parte.split("-")
            a, b = int(a.strip()), int(b.strip())
            meses.update(range(min(a, b), max(a, b) + 1))
        else:
            meses.add(int(parte))

    if not meses:
        raise ValueError("No se especificó ningún mes.")
    if not all(1 <= m <= 12 for m in meses):
        raise ValueError("Los meses deben estar entre 1 y 12.")
    return sorted(meses)


def pedir_periodos():
    """
    Pide interactivamente los años y meses a procesar y retorna la lista
    completa de periodos (mes, año) a correr, en orden cronológico.
    Enter vacío usa el mes anterior / año actual como valor por defecto.
    """
    mes_auto, año_auto = detectar_periodo()

    titulo("Selección de periodo(s) a procesar", nivel=2)

    while True:
        raw_años = input(
            f"\n  Año(s) a procesar (ej. 2026 | 2023,2024 | 2023-2025) "
            f"[Enter = {año_auto}]: "
        ).strip()
        if not raw_años:
            años = [año_auto]
            break
        try:
            años = parse_años(raw_años)
            break
        except ValueError as e:
            print(f"  ⚠️  Entrada inválida ({e}). Intenta de nuevo.")

    while True:
        raw_meses = input(
            f"  Mes(es) por NÚMERO (ej. 3 | 1,2,5 | 1-6 | todos) "
            f"[Enter = {mes_auto}]: "
        ).strip()
        if not raw_meses:
            meses = [mes_auto]
            break
        try:
            meses = parse_meses(raw_meses)
            break
        except ValueError as e:
            print(f"  ⚠️  Entrada inválida ({e}). Intenta de nuevo.")

    periodos = [(mes, año) for año in años for mes in meses]

    print(f"\n  📋 Se procesarán {len(periodos)} periodo(s):")
    for mes, año in periodos:
        print(f"     - {MESES_ES[mes]} {año}")

    if not confirmar("¿Continuar con estos periodos?"):
        print("\n  🔁 Volviendo a pedir los periodos...")
        return pedir_periodos()

    return periodos


def preguntar_modo_envio():
    """
    Pregunta UNA sola vez cómo manejar el paso 6 (envío) para todos los
    periodos del lote, para no tener que confirmar manualmente en cada uno
    cuando se procesan muchos meses seguidos (ej. extracción histórica).
    """
    titulo("Manejo del envío (Paso 6) para todos los periodos", nivel=2)
    print("  1) Preguntar en cada periodo si ya fue revisado (flujo normal)")
    print("  2) Enviar automáticamente como REVISADO en todos (limpieza + envío al cliente)")
    print("  3) Enviar automáticamente PARA REVISIÓN en todos (correo a Claudia)")
    print("  4) No enviar ningún correo — solo generar los reportes hasta el paso 5")

    while True:
        opcion = input("\n  Elige una opción [1-4]: ").strip()
        if opcion in ("1", "2", "3", "4"):
            return opcion
        print("  ⚠️  Opción inválida.")


def calcular_rango_portal(mes, año):
    """
    Calcula FECHA_INI y FECHA_FIN para el portal Owcia (oñate).
    El portal muestra operaciones con ~2 meses de desfase respecto al mes de factura.
    Rango: primer día de M-2  →  último día de M-1
    """
    primer_dia_mes = date(año, mes, 1)

    # Último día del mes anterior (M-1)
    fin_m1 = primer_dia_mes - timedelta(days=1)
    # Primer día del mes M-1
    ini_m1 = fin_m1.replace(day=1)

    # Último día de M-2
    fin_m2 = ini_m1 - timedelta(days=1)
    # Primer día de M-2
    ini_m2 = fin_m2.replace(day=1)

    return ini_m2.strftime("%Y-%m-%d"), fin_m1.strftime("%Y-%m-%d")


def cargar_modulo(alias, nombre_archivo):
    """
    Importa un módulo desde su ruta sin ejecutar el bloque __main__.
    Devuelve el objeto de módulo para que se puedan parchear sus variables.
    """
    ruta = os.path.join(PATH_BASE, nombre_archivo)
    spec = importlib.util.spec_from_file_location(alias, ruta)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def exec_con_vars(nombre_archivo, substituciones: dict):
    """
    Ejecuta un script que corre lógica a nivel de módulo (sin funciones),
    inyectando las variables dinámicas ANTES de que se evalúen las f-strings.

    Se usa exclusivamente para organon_2_1.py, donde las queries SQL
    son f-strings calculadas en la raíz del módulo.
    """
    ruta = os.path.join(PATH_BASE, nombre_archivo)
    with open(ruta, "r", encoding="utf-8") as f:
        source = f.read()

    for patron, reemplazo in substituciones.items():
        source = re.sub(patron, reemplazo, source)

    # Asegurar que los archivos generados caigan en PATH_BASE
    os.chdir(PATH_BASE)
    ns = {"__name__": "__main__", "__file__": ruta}
    exec(compile(source, ruta, "exec"), ns)  # noqa: S102


def etiqueta_lote(periodos):
    """Construye una etiqueta estable para archivos generados por rango."""
    ordenados = sorted(periodos, key=lambda p: (p[1], p[0]))
    mes_ini, año_ini = ordenados[0]
    mes_fin, año_fin = ordenados[-1]
    return f"{año_ini}_{mes_ini:02d}_{año_fin}_{mes_fin:02d}"


def rango_lote(periodos):
    """Devuelve fecha inicial y final calendario para un lote de meses."""
    ordenados = sorted(periodos, key=lambda p: (p[1], p[0]))
    mes_ini, año_ini = ordenados[0]
    mes_fin, año_fin = ordenados[-1]
    ultimo_dia_fin = calendar.monthrange(año_fin, mes_fin)[1]
    return (
        date(año_ini, mes_ini, 1).strftime("%Y-%m-%d"),
        date(año_fin, mes_fin, ultimo_dia_fin).strftime("%Y-%m-%d"),
    )


def confirmar(pregunta):
    """Retorna True si el usuario responde sí."""
    r = input(f"\n  {pregunta} [s/N]: ").strip().lower()
    return r in ("s", "si", "sí", "y", "yes")

# ===========================================================================
# PASOS DEL PIPELINE
# ===========================================================================

def paso_1_extraccion_base(mes, año):
    titulo(f"PASO 1 — Extracción Base SQL  |  {MESES_ES[mes]} {año}")
    # organon_2_1.py ejecuta queries SQL a nivel de módulo (no hay función),
    # por lo que se inyectan los valores antes de ejecutar el archivo.
    exec_con_vars(
        "organon_2_1.py",
        {
            r"MES_VALIDACION\s*=\s*\d+":  f"MES_VALIDACION = {mes}",
            r"ANIO_VALIDACION\s*=\s*\d+": f"ANIO_VALIDACION = {año}",
        },
    )
    print(f"\n  ✅ Archivo generado: REPORTE_ADUANAL_SANOFI_{año}_{mes:02d}.xlsx")


def paso_1_extraccion_base_general(periodos):
    fecha_ini, fecha_fin = rango_lote(periodos)
    etiqueta = etiqueta_lote(periodos)

    titulo(f"PASO 1 — Extracción Base SQL GENERAL  |  {fecha_ini} → {fecha_fin}")
    exec_con_vars(
        "organon_2_1.py",
        {
            r"FECHA_INICIO_VALIDACION\s*=\s*None": f'FECHA_INICIO_VALIDACION = "{fecha_ini}"',
            r"FECHA_FIN_VALIDACION\s*=\s*None": f'FECHA_FIN_VALIDACION = "{fecha_fin}"',
            r"ETIQUETA_PERIODO\s*=\s*None": f'ETIQUETA_PERIODO = "{etiqueta}"',
        },
    )
    print(f"\n  ✅ Archivo general generado: REPORTE_ADUANAL_SANOFI_{etiqueta}.xlsx")


def paso_2_extraccion_manzanillo(mes, año):
    titulo("PASO 2 — Corresponsalías Manzanillo  |  oñate_extraccion.py")

    fecha_ini, fecha_fin = calcular_rango_portal(mes, año)
    print(f"  📅 Rango portal: {fecha_ini}  →  {fecha_fin}")

    mod = cargar_modulo("onate_extraccion", "oñate_extraccion.py")

    # Parchear variables de configuración ANTES de instanciar la clase
    mod.MES_VALIDACION    = mes
    mod.ANIO_VALIDACION   = año
    mod.FECHA_INI         = fecha_ini
    mod.FECHA_FIN         = fecha_fin
    mod.ARCHIVO_EXCEL_FINAL = os.path.join(PATH_BASE, f"REPORTE_MAESTRO_{año}_{mes:02d}.xlsx")

    resultado = mod.SuperScraperCorresponsalias(headless=False).ejecutar()
    if resultado is False:
        print("  ℹ️  El pipeline continuará sin archivo de corresponsalías Manzanillo.")
    elif resultado is None:
        print("  ⚠️  No se pudo consultar Manzanillo por SQL. El pipeline seguirá, pero revisa la conectividad si esperabas datos.")


def paso_3_laredo_y_analisis(mes, año):
    titulo("PASO 3 — Laredo: descarga PDFs + análisis")

    # 3a — Descarga de PDFs vía Selenium
    titulo("PASO 3a — Descarga PDFs  |  laredo.py", nivel=2)
    laredo = cargar_modulo("laredo", "laredo.py")
    laredo.MES_VALIDACION  = mes
    laredo.ANIO_VALIDACION = año
    laredo.LaredoExtractor().ejecutar()

    # 3b — Análisis de los PDFs descargados
    titulo("PASO 3b — Análisis PDFs  |  analisis.py", nivel=2)
    analisis = cargar_modulo("analisis", "analisis.py")
    analisis.AnalizadorLaredo().analizar_archivos()


def paso_4_union(mes, año):
    titulo("PASO 4 — Unificación de Reportes  |  union_2_1.py")

    mod = cargar_modulo("union_2_1", "union_2_1.py")

    # Rutas dinámicas que en el script original están hardcodeadas
    mod.ARCHIVO_ADUANAL = os.path.join(PATH_BASE, f"REPORTE_ADUANAL_SANOFI_{año}_{mes:02d}.xlsx")
    mod.ARCHIVO_MAESTRO = os.path.join(PATH_BASE, f"REPORTE_MAESTRO_{año}_{mes:02d}.xlsx")
    mod.ARCHIVO_SALIDA  = os.path.join(PATH_BASE, "Reporte_SANOFI.xlsx")

    mod.unificar_reportes()
    print(f"\n  ✅ Reporte unificado: Reporte_SANOFI.xlsx")


def paso_5_revision_profunda(mes):
    titulo("PASO 5 — Revisión Profunda  |  revision_profunda.py")

    mod = cargar_modulo("revision_profunda", "revision_profunda.py")

    mod.MES_ANALIZADO  = f"{mes:02d}"
    mod.ARCHIVO_ENTRADA = os.path.join(PATH_BASE, "Reporte_SANOFI.xlsx")
    mod.ARCHIVO_SALIDA  = os.path.join(PATH_BASE, f"Reporte_SANOFI_{mes:02d}.xlsx")
    mod.PATH_DESCARGAS  = os.path.join(PATH_BASE, "Descargas_ZIP")

    mod.ejecutar_revision()
    print(f"\n  ✅ Reporte revisado: Reporte_SANOFI_{mes:02d}.xlsx")


def paso_6a_envio_revision(mes, año):
    """Envía el reporte a Claudia para revisión interna."""
    titulo("PASO 6b — Envío para revisión  |  correo_revision.py")

    mod = cargar_modulo("correo_revision", "correo_revision.py")

    mod.MES_NOMBRE   = MESES_ES[mes]
    mod.ANIO         = str(año)
    mod.RUTA_ARCHIVO = os.path.join(PATH_BASE, f"Reporte_SANOFI_{mes:02d}.xlsx")

    mod.enviar_correo_pro()
    print("\n  📬 Correo de revisión enviado a Claudia.")
    print("  ℹ️  Cuando confirme, ejecuta nuevamente desde el paso 6.")


def paso_6b_envio_final(mes, año):
    """Limpia descargas, separa por pestañas y envía al cliente."""
    titulo("PASO 6a — Limpieza + Separación + Envío Final")

    archivo_final = os.path.join(PATH_BASE, f"Reporte_SANOFI_{mes:02d}.xlsx")

    # Limpieza de carpetas de descargas temporales
    titulo("Limpiando carpetas temporales  |  limpieza.py", nivel=2)
    limpieza = cargar_modulo("limpieza", "limpieza.py")
    limpieza.limpiar_carpeta_descargas(os.path.join(PATH_BASE, "Descargas_ZIP"))
    limpieza.limpiar_carpeta_descargas(os.path.join(PATH_BASE, "Descargas_LT"))

    # Separación por pestañas (referencia por hoja)
    titulo("Separando por pestañas  |  separador_pestañas.py", nivel=2)
    separador = cargar_modulo("separador_pestañas", "separador_pestañas.py")
    separador.separar_por_referencias(archivo_final)

    # Envío al cliente
    titulo("Enviando correo al cliente  |  correo.py", nivel=2)
    correo = cargar_modulo("correo", "correo.py")
    correo.ASUNTO        = f"Reporte Aduanal SANOFI - {MESES_ES[mes]} {año}"
    correo.RUTA_ARCHIVO  = archivo_final
    correo.enviar_correo_pro()

    print(f"\n  ✅ Reporte de {MESES_ES[mes]} {año} enviado al cliente.")

# ===========================================================================
# MAIN
# ===========================================================================

def parse_args():
    parser = argparse.ArgumentParser(
        description="Orquestador del Reporte Mensual SANOFI",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument(
        "--desde",
        type=int,
        choices=range(1, 7),
        default=1,
        metavar="PASO",
        help=(
            "Inicia el pipeline desde este paso (1-6).\n"
            "  1 = Extracción Base SQL\n"
            "  2 = Corresponsalías Manzanillo\n"
            "  3 = Laredo (descarga + análisis)\n"
            "  4 = Unificación\n"
            "  5 = Revisión Profunda\n"
            "  6 = Envío (revisión o final)"
        ),
    )
    return parser.parse_args()


def procesar_periodo(mes, año, paso_inicio, modo_envio):
    """Corre el pipeline completo (o desde paso_inicio) para un único periodo."""
    if paso_inicio > 1:
        print(f"\n  ⏭️  Iniciando desde el Paso {paso_inicio}.")

    if paso_inicio <= 1:
        paso_1_extraccion_base(mes, año)

    if paso_inicio <= 2:
        paso_2_extraccion_manzanillo(mes, año)

    if paso_inicio <= 3:
        paso_3_laredo_y_analisis(mes, año)

    if paso_inicio <= 4:
        paso_4_union(mes, año)

    if paso_inicio <= 5:
        paso_5_revision_profunda(mes)

    # ── Punto de decisión: ¿Ya revisó Claudia? ──────────────────────────────
    print("\n" + "─" * 62)
    print(f"  📋 Reporte listo: Reporte_SANOFI_{mes:02d}.xlsx")

    if modo_envio == "1":
        if confirmar("¿Ya fue revisado y aprobado por Claudia?"):
            paso_6b_envio_final(mes, año)
        else:
            paso_6a_envio_revision(mes, año)
    elif modo_envio == "2":
        paso_6b_envio_final(mes, año)
    elif modo_envio == "3":
        paso_6a_envio_revision(mes, año)
    elif modo_envio == "4":
        print("  ℹ️  Modo sin envío: reporte generado, no se envía ningún correo.")


def main():
    args = parse_args()
    paso_inicio = args.desde

    titulo("REPORTE MENSUAL SANOFI — AUTOMATIZACIÓN")

    periodos = pedir_periodos()

    os.chdir(PATH_BASE)

    if len(periodos) > 1:
        if paso_inicio != 1:
            print("\n  ⚠️  El modo general por varios meses solo inicia desde el paso 1.")
            print("  ℹ️  Ignorando --desde para generar una extracción consolidada.")

        paso_1_extraccion_base_general(periodos)
        etiqueta = etiqueta_lote(periodos)
        titulo("RESUMEN DEL LOTE")
        print(f"\n  ✅ Extracción general completada: REPORTE_ADUANAL_SANOFI_{etiqueta}.xlsx")
        print("  ℹ️  No se generaron reportes mensuales segmentados.")
        titulo("✅  PROCESO GENERAL COMPLETADO")
        return

    modo_envio = preguntar_modo_envio() if paso_inicio <= 6 else "4"

    exitosos = []
    fallidos = []

    for i, (mes, año) in enumerate(periodos, start=1):
        titulo(f"PERIODO {i}/{len(periodos)} — {MESES_ES[mes]} {año}")
        try:
            procesar_periodo(mes, año, paso_inicio, modo_envio)
            exitosos.append((mes, año))
            titulo(f"✅  PERIODO COMPLETADO — {MESES_ES[mes]} {año}")
        except Exception as e:
            fallidos.append((mes, año, str(e)))
            print(f"\n  ❌ Error procesando {MESES_ES[mes]} {año}: {e}")
            print(f"  ⏭️  Se continúa con el siguiente periodo del lote.")

    # ── Resumen final del lote ───────────────────────────────────────────────
    titulo("RESUMEN DEL LOTE")
    print(f"\n  ✅ Completados ({len(exitosos)}):")
    for mes, año in exitosos:
        print(f"     - {MESES_ES[mes]} {año}")

    if fallidos:
        print(f"\n  ❌ Con errores ({len(fallidos)}):")
        for mes, año, err in fallidos:
            print(f"     - {MESES_ES[mes]} {año}  →  {err}")
        print(
            f"\n  💡 Para reintentar un periodo fallido, corre main.py de nuevo "
            f"indicando solo ese mes/año (con --desde N si ya generó archivos parciales)."
        )

    titulo("✅  PROCESO DE LOTE COMPLETADO")


if __name__ == "__main__":
    main()
