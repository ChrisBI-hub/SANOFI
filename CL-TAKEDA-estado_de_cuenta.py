import pyodbc
import pandas as pd
from datetime import datetime
import os
import openpyxl
from openpyxl.styles import numbers
from openpyxl.styles import PatternFill
from openpyxl.styles import Font, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage
# -*- coding: utf-8 -*- 
from sys import platform
import locale
from openpyxl.styles import Border, Side

try:
    locale.setlocale(locale.LC_TIME, 'es_MX.UTF-8')
except:
    locale.setlocale(locale.LC_TIME, 'Spanish_Mexico.1252')
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = (os.path.dirname(os.path.dirname(BASE_DIR)))  # <- dos niveles arriba
carpeta_destino = os.path.join(PARENT_DIR, "estados_de_cuenta")

fecha = datetime.today()
mes = fecha.strftime('%m')
anio = fecha.strftime('%Y')
anioANTERIOR = str(int(anio) - 1)
# Configuración de la base de datos
SERVER = "150.1.1.152"
DATABASE = "SIR"
USERNAME = "ConsultaBD"
PASSWORD = "5D$bc#kM&5W2T8J40?s%"

# --- Conexión a SQL Server ---
conn = conn = pyodbc.connect(
        f"DRIVER={{ODBC Driver 18 for SQL Server}};"
        f"SERVER={SERVER};"
        f"DATABASE={DATABASE};"
        f"UID={USERNAME};"
        f"PWD={PASSWORD};"
        f"TrustServerCertificate=yes;"
    )
query = F"""
/*  TAKEDA ESTADOS DE CUENTA Geraldine */
/*  Fecha de elabporación: 26/08/2025  */
/*  Elaborado por: Miguel Angel Yañez Jacinto  */
/*  Version 1.0  */

SELECT 
CuentaDeGasto AS "FACTURA",
EstatusCGA AS "ESTATUS CGA",
Fecha AS "FECHA",
Referencia AS "REFERENCIA",
Pedimento AS "PEDIMENTO",
DescripcionMercancia AS "MERCANCIA",

/* Funcion para calsificar la mercancia por GMS y LOC */

/* Clasificación LOC : PRODUCTO TERMINADO */
CASE
    when Operacion LIKE '%EXPORTACION%' then 'GMS'

        WHEN DescripcionMercancia LIKE '% material de empaque%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%Alogliptina benzoato (syr-322)%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%B8610-014%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%B8610-044%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%B9110-021%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%B8510-028%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%Byk347611%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%B7499-107%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%B7605-024 ( byk11684 ) mc-4-sa%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%B7805-005 ( mc-4-sa)%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%B7705-012%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%Pantoprazol Hemi Magnesico%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%Pantoprazol Sodico Sesquihidratado%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%Pioglitazona Clorhidrato%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%Tak-491 ( Azilsartan Medoxomil )%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%Tak-491 Stress%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%Tak-390 ( Dexlansoprazol )%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%Policresuleno 50%%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%Policresuleno%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%Alogliptina%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%Metformina%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%Alogliptina%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%Pioglitazona%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%Cumarina%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%Azilsartan%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%Medoxomilo%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%Pantozol%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%Dexlansoprazol%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%Bromuro de Pinaverio%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%Esencia de Naranja%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%Pancreatina%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%tapas e insertos%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%papel riopan%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%ALOGLIPTINA METFORMINA%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%ALOGLIPTINA METFORMINA 12.5 /500%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%DEXILANT LR CAP 60 MG W2 MS%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%ALIGLIPTINA BENZOATE%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%ALEVIAN DUO%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%TECTA%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%POLIETILENGLICOL 8000 (POLYGLYKOL 8000 PF)%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%DICLUCONATO DE CLORHEXIDINA AL 20%%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%PAPEL IMPRESO%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%ITOPRIDE HYDROCHLORIDE%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%DEXILANT LR CAP 60 MG%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%DEXILANT LR CAP 30 MG%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%DEXILANT CAP 30 MG W/2 MS%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%DEXILANT CAP 30 MG%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%DEXILANT XR CAP 60 MG C/14%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%DEXILANT 60 MG W/14%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%DEXILANT CAP 60 MG W/2 MS%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%FAKTU RL OINT 0.5/0.1G W10 G AP MS%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%FAKTU OINTMENT TUBE W20 G & APLIC%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%TECTA TAB 40 MG BOX W/2 TAB MS%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%DEXILANT LR CAP 30MG W2 MS%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%ALOGLIPTINA, AZILSARTAN%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%AZILSARTAN MEDOXOMILO%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%ALGOLIPTINA/METFORMINA%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%FAKTU OINTMENT%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%ALOGLIPTINA7PIOGLITAZONA%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%ALEVIAN DUO CAP 100/300 MG BX W64%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%ALEVIAN DUO CAP 100/300MG BX%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%RIOPAN GEL BOTTLE 1/8 G W/250 ML%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%FAKTU OINTMENT 0.5/0.1G W10G MS%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%TECTA TABLET 40 MG W2 MS%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%TECTA TABLET 40 MG C14%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%DEXILANT CAP 30 MG W/14%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%MUESTRAS DE ALEVIAN DUO 100MG/300MG%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%AZILSARTAN%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%ITROPRIDE CLORHIDRATE%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%ALEVIAN DUO CAP 100/300MG W64%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%RIOPAN GEL FLASK 8/1 G W/250 ML%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%INCRESINA%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%ALOGLIPTINA%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%RIOPAN GEL FLASK 8/1 G W/250 ML%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%RIOPAN GEL FLASK 8/1 G W/250ML%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%ALEVIAN DUO CAP 100/300MG W64%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%ALEVIAN DUO CAO 100MG SLE W48%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%DEXILANT LR CAP 30 MG W2 MS%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%DEXILANT LR CAP 30 MG W2 MS%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%ALEVIAN DUO CAP 100MG SLE W48%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%RIOPAN GEL 8/1 BOX C20 SCH C10 ML%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%RIOPAN GEL 8/1G BOX C20 SCH C10 ML%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%PAPEL IMPRESO PRESENTADO EN BOBINAS (ROLLOS) RIOPAN%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%RIOPAN GEL80/10 MG W20 SACH%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%RIOPAN GEL 8/1 G BOX C20 SCH C10 ML%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%AZILSARTAN%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%RIOPAN GEL FCO 80/10MG 250 ML%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%POLICRESULENO 36%%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%ALOGLIPTINA METFORMINA%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%POLICRESULENO 50%%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%BROMURO DE PINAVERIO%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%ALEVIAN DUO%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%ALEVIAN DUO%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%POLICRESULEN 50%%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%ITOPRIDE HYDROCHLORIDE%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%PAPEL CON REFUERZO DE ALUMINIO Y RECUBIERTO DE PLASTICO, IMPRESO%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%MANUFACTURAS DE PLASTICO (INSERTOS)  (SPRITZEINSATZ E E-24 PE-LD WHITE-COLOURED)%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%MANUFACTURAS DE PLASTICO (INSERTOS)  (SPRITZEINSATZ E E-24 PE-LD WHITE-COLOURED)%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%HIERRO POLIMALTOSADO%' THEN 'GMS'
        WHEN DescripcionMercancia LIKE '%POLICRESULENO CONCENTRATE 36%%' THEN 'GMS'
       -- Sin Clasificación que deberían tener categoría
WHEN DescripcionMercancia LIKE '%ALBOTHYL%' THEN 'GMS'
WHEN DescripcionMercancia LIKE '%POLICRESULEN%' THEN 'GMS'
WHEN DescripcionMercancia LIKE '%DEXILANT%' THEN 'GMS'
WHEN DescripcionMercancia LIKE '%ZURCAL%' THEN 'GMS'
WHEN DescripcionMercancia LIKE '%PANTECTA%' THEN 'GMS'
WHEN DescripcionMercancia LIKE '%PANTECTAP%' THEN 'GMS'
WHEN DescripcionMercancia LIKE '%FAKTU%' THEN 'GMS'
    WHEN DescripcionMercancia LIKE '%RIOPAN GEL 8/1G CJA C20 SOB C10 ML' then 'GMS'
    WHEN DescripcionMercancia LIKE '%RIOPAN GEL 80/10MG C/1 SOB MM' then 'GMS'
    WHEN DescripcionMercancia LIKE '%RIOPAN GEL 80/10MG C/1 SOB MM' then 'GMS'
    WHEN DescripcionMercancia LIKE '%RECURSO TECNICO' then 'GMS'
    WHEN DescripcionMercancia LIKE '%VIALOK ' then 'GMS'
    WHEN DescripcionMercancia LIKE '%HIERRO POLIMALTISADO' then 'GMS'
    WHEN DescripcionMercancia LIKE '%PANTOXOL' then 'GMS'
    WHEN DescripcionMercancia LIKE '%RECURSO TECNICO' then 'GMS'
    WHEN DescripcionMercancia LIKE '%RIOPAN GEL 80/10MG FCO 250ML' then 'GMS'
    WHEN DescripcionMercancia LIKE '%PANTOPRAZOL' then 'GMS'
    WHEN DescripcionMercancia LIKE '%ESTANDAR REFERENCIA F II-IX-X PUR' then 'GMS'
    WHEN DescripcionMercancia LIKE '%CLORHIDRATO DE CINCOCAINA (DIBUCAINE CHLORHYDRATE)' then 'GMS'
    WHEN DescripcionMercancia LIKE '%RECURSO TECNICO' then 'GMS'
    WHEN DescripcionMercancia LIKE '%RECURSO TECNICO' then 'GMS'
    WHEN DescripcionMercancia LIKE '%DULBECCOS' then 'GMS'
    WHEN DescripcionMercancia LIKE '%ALUNBRING  BRIGATINIB 90 MG' then 'GMS'
    WHEN DescripcionMercancia LIKE '%ESTANDAR DE REFERENCIA' then 'GMS'
    WHEN DescripcionMercancia LIKE '%DILUENT SOLVENT ' then 'GMS'
    WHEN DescripcionMercancia LIKE '%RIOPAN GEL 8010 MG' then 'GMS'

        ---loc
        WHEN DescripcionMercancia LIKE '%Ninlaro%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%Advate%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%Adynovate%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%ADCETRIS%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%Cuvitru%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%Elaprase%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%Entyvio%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%Entyvio SC%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%Feiba%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%Firazyr%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%Flowease%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%Hemofil%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%Human Albumin%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%Hyqvia%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%Immunine%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%Intuniu%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%Kiovig%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%MEPACT%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%Qdenga%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%Recombinate%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%Replagal%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%Rixubis%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%Takhzyro%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%Vpriv%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%Vyvanse%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%PRODUCTO%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%STANDARES%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%BOC-L-LYS(BOC)-L%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%BOC-L-LYS(BOC)-D%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%BOC-D-LYS(BOC)-D-ANFETAMINA%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%BOC-D-LYS(BOC)-L-ANFETAMINA%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%LYS (LYS) DEX (ISOMERO 1 )%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%LYS (LYS) DEX (ISOMERO 2 )%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%FVIII DEFICIENT PLASMA (IMMUNODEPLETED)%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%SUBSTRATE FXa-1+aNAPAP%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%VEDOLIZUMAB WORKING REFERENCE STANDARD (WRS)%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%PRECICONTROL CLINCHEM MULTI 1%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%REFERENCE STANDARD RHUPH20%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%TESTKONTROLLE PH20-ASSAY%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%ACIDO TRIFLUORACETICO%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%HIALURONATO DE SODIO%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%PLACEBO BLEND%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%FRUQUINTINIB CAPSULAS 1 y 5 mg%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%HMPL-013-SM1 /   HMPL-013-SM2   / DHQ%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%ESTANDAR DE REFERENCIA / CMB%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%SYRINGE FILTER, HYDROPHILIC POLYETHERSULFONE (PTFE)%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%SYRINGE FILTER, HYDROPHILIC POLYETHERSULFONE (PES)%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%ACTIN FSL%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%SUBSTRATE PK1%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%ALBUMINA HUMANA 20%%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%FEIBA REACTIVO ANALITICO PREPARADO%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%CÉLULAS VERO%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%POLICRESULENO 50%%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%DEXLANSOPRAZOL (TAK-390) ESTÁNDAR DE REFERENCIA%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%PLACEBO PANTOPRAZOL SÓDICO 20/40 MG RECUBIERTO%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%PANTOPRAZOL HEMI MAGNESICO%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%PLACEBO PANTOPRAZOL NUCLEO 20/40 mg%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%B8401-026%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%B8510-028%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%B8610-014%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%B8810-044%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%B9110-021%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%ALBUMINA HUMANA 25%%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%ALKALINE PHOSPHATASE CONJUGATED ANTI-DEN-1%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%MONOCLONAL ANTIBODY%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%ALKALINE PHOSPHATASE CONJUGATED ANTI-DEN-2%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%MONOCLONAL ANTIBODY%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%ALKALINE PHOSPHATASE CONJUGATED ANTIDEN-%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%3 MONOCLONAL ANTIBODY%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%ALKALINE PHOSPHATASE CONJUGATED ANTIDEN-%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%4 MONOCLONAL ANTIBODY%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%AVICEL RC-591 NF%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%DULBECCO´S MODIFIED EAGLE MEDIUM%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%TAKEDA DENGUE TETRAVALENT VACCINE (TDV) DILUENT%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%TAKEDA TDV QC REFERENCE MATERIAL%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%DAGLA MUESTRA%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%DAGLA PLACEBO%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%FERRANINA FOL MUESTRA%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%FERRANINA FOL PLACEBO%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%FX (FACTOR X) BUFFER%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%PREKALLIKREIN-REAGENT%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%CONTROL PREPARATION FEIBA PUR%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%REFERENCE STANDARD PKA%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%TEST BUFFER PLASMIN%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%CITRATE SODIUM CHLORIDE BUFFER%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%Dexoamphetamine Mesylate Reference Standard%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%Lisdexamfetamine Dimesylate Reference Standard%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%KIT PARA LA TOMA DE MUESTRA EN PAPEL FILTRO%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%HI BUMIN 250%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%TAKHZYRO (LANADELUMAB) (MEDICAMENTO DE USO HUMANO)%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%ENTYVIO%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%KIT PARA LA TOMA DE MUESTRA EN PAPEL FILTRO%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%ALUNBRIG%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%HI BUMIN 250%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%TAKHZYRO (LANADELUMAB) (MEDICAMENTO DE USO HUMANO)%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%ADCETRIS%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%ADVATE OCTOG ALFA 500UI INJECTABLE%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%TAKHZYRO (LANADELUMAB)%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%HI BUMIN 250%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%DEXLANSOPRAZOL CAPSULAS DE 60  MG (MEDICAMENTO DE USO HUMANO)%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%CITRATE SODIUM CHLORIDE BUFFER%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%ADCETRIS%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%NINLARO%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%PREKALLIKREIN-REAGENT%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%ADVATE, ADYNOVATE%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%ADVATE, ADYNOVATE, KIOVIG%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%ADYNOVATE,1000IU,SWFI 5ML BJII%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%ADYNOVATE,500IU,SWFI 5ML BJII%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%ESTANDARES%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%ESTANDARES DE REFERENCIA%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%ESTANDARES FRUQUINIB%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%ESTANDARES MAT%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%REPLAGAL (AGALSIDASA ALFA) 3.5ML (1MG/ML) (MEDICAMENTO DE USO HUMANO)%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%SUBSTRATE PK-1%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%ADCETRIS (BRENTUXIMAB VEDOTIN) CAJA CON UN FRASCO AMPULA CON 50 MG (MEDICAMENTO DE USO HUMANO)%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%ELAPRASE%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%HYQVIA DEMO KIT%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%FEIBA REAGENT%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%VPRIV%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%REPLAGAL AGALSIDASA ALFA%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%ESTANDARES (4)%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%HI-BUMIN%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%HI-BUMIN%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%IMMNUNINE Y FEIBA%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%IMMNUNINE Y FEIBA%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%HI BUMIN%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%HI BUMIN%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%NINLARO4 MG%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%HI-BUMIN%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%AGUA%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%HI-BUMIN SOLUCION DE ALBUMINA HUMANA FRASCO CON 50 ML (MEDICAMENTO DE USO HUMANO)%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%HI-BUMIN 50 ML (MEDICAMENTO DE USO HUMANO)%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%IMMUNINE CAJA CON UN FRASCO AMPULA CON 600UI (MEDICAMENTO DE USO HUMANO)%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%B7499-107  M-CRESOL-6-SULFONIC ACID ESTANDAR DE REFERENCIA%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%KIT PARA LA TOMA DE MUESTRA EN PAPEL FILTRO%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%DEXLANSOPRAZOL  (ESTANDAR DE REFERENCIA)%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%CITRATE SODIUM CHLORIDE BUFFER  50ML%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%CONTROL PREPARATION ALBUMIN HUMAN LYO%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%SUBSTRATE PK-1%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%PREKALLIKREIN REAGENT%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%TEST BUFFER PLASMIN (AUPL)%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%HYQVIA INMUNOGLOBULINA HUMANA NORMAL 10 G/ 100 ML (MEDICAMENTO DE USO HUMANO)%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%HYQVIA INMUNOGLOBULINA HUMANA NORMAL 10 G/100 ML%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%MEPACT (MIFAMURTIDA) CAJA CON UN FRASCO CON 4 MG (MEDICAMENTO DE USO HUMANO)%' THEN 'LOC'
        WHEN DescripcionMercancia LIKE '%Alunbrig%' THEN 'LOC'
WHEN DescripcionMercancia LIKE '%13MM VIALOK%' THEN 'LOC'
WHEN DescripcionMercancia LIKE '%VIAL ACCESS DEVICE%' THEN 'LOC'
WHEN DescripcionMercancia LIKE '%ALBUMIN HUMAN%' THEN 'LOC'
WHEN DescripcionMercancia LIKE '%ALBUM HUMAN 25%' THEN 'LOC'
WHEN DescripcionMercancia LIKE '%HUMAN ALBUMIN%' THEN 'LOC'
WHEN DescripcionMercancia LIKE '%LIVTENCITY%' THEN 'LOC'
WHEN DescripcionMercancia LIKE '%MARIBAVIR%' THEN 'LOC'
WHEN DescripcionMercancia LIKE '%EXKRUTHERA%' THEN 'LOC'
WHEN DescripcionMercancia LIKE '%FRUQUINTINIB%' THEN 'LOC'
WHEN DescripcionMercancia LIKE '%FRUQUINTINIB REFERENCE STANDARD%' THEN 'LOC'
WHEN DescripcionMercancia LIKE '%CONTROL PREPARATION PLASMIN%' THEN 'LOC'
WHEN DescripcionMercancia LIKE '%ESTANDAR DE REFERENCIA F II-IX-X%' THEN 'LOC'
WHEN DescripcionMercancia LIKE '%ENVASES DE PLASTICO%' THEN 'LOC'
WHEN DescripcionMercancia LIKE '%TUBO DE PROCTOACID%' THEN 'LOC'
        ELSE 'Sin Ubicación'
END AS CLASIFICACIÓN,


Pedido AS "PEDIDO",
sObservacionesCGA AS "OBSERVACIONES",
Operacion AS "OPERACION",
Cliente AS "CLIENTE",
PagosHechosME AS "PAGOS TERCEROS",
HonorariosME AS "HONORARIOS ABC",
ComplementariosME AS "COMPLEMENTARIOS",
IVA AS "IVA",
AnticiposME AS "ANTICIPOS",
TotalME AS "TOTAL FACTURA",

LiquidacionME AS "LIQUIDACION / ANTICIPO",

/* Funcion para sacar el saldo acumulado */

 SaldoME AS "SALDO FACTURA",
 
/* Funcion para sacar los dias vencidos */

  CASE 
        WHEN DATEDIFF(DAY, GETDATE(), DATEADD(DAY, DiasCred, TRY_CAST(Fecha AS DATE))) > 0 THEN 0
        WHEN DATEDIFF(DAY, GETDATE(), DATEADD(DAY, DiasCred, TRY_CAST(Fecha AS DATE))) < 0 
            THEN -DATEDIFF(DAY, GETDATE(), DATEADD(DAY, DiasCred, TRY_CAST(Fecha AS DATE)))
        ELSE ''
    END AS "DIAS VENCIDOS",

/* Funcion para sacar el estatus */

CASE 
        WHEN DATEDIFF(DAY, GETDATE(), DATEADD(DAY, DiasCred, TRY_CAST(Fecha AS DATE))) <= 0 THEN 'VENCIDO'
        WHEN DATEDIFF(DAY, GETDATE(), DATEADD(DAY, DiasCred, TRY_CAST(Fecha AS DATE))) > 0 THEN 'CORRIENTE'
        ELSE 'N/A'
    END AS [ESTATUS],

--	'' AS "SALDO ACUMULADO",

/* Función para sacar la fecha limite de pago */

  CONVERT(VARCHAR(10), DATEADD(DAY, DiasCred, TRY_CAST(Fecha AS DATE)), 103) AS "FECHA LIMITE DE PAGO",


UUID AS "UUID",
DiasCred as "Días Credito"


/* Conexión a la base de datos */


FROM [SIR].[Admin].[SIR_VT_CCEstadoDeCuenta]
    WHERE Cliente = 'TAKEDA MEXICO S.A. DE C.V.'
--and Referencia = '25-001514'
--AND Sucursal IN ('VERACRUZ','AIFA ESTADO DE MEXICO')
AND EstatusCGA = 'FACTURADA'
  AND (
        (EstatusCGA = 'PROFORMA' AND LEFT(Fecha,4) in ('{anio}','{anioANTERIOR}'))
        OR
        (EstatusCGA <> 'PROFORMA')
      )
--AND MONTH(CuentaG_FechaFactura) IN (1,2,3,4,5,6,7,8,9,10,11)
--AND YEAR(CuentaG_FechaFactura)=2025
	 ORDER BY  left(Fecha,10) ASC;

        """
# Lista de columnas que deseas conservar

janssen = pd.read_sql(query, conn)
janssen['FECHA'] = pd.to_datetime(janssen['FECHA'], errors='coerce')
CLIENTE = janssen['CLIENTE'].dropna().unique()[0] if not janssen.empty else 'N/A'

janssen['FECHA LIMITE DE PAGO'] = janssen['FECHA'] + pd.to_timedelta(
    janssen['Días Credito'].fillna(0), unit='D'
)
ruta_salida = os.path.join(carpeta_destino, f"ESTADO DE CUENTA {CLIENTE}.xlsx")

# --- Convertir FECHA a datetime ---
#janssen['FECHA'] = pd.to_datetime(janssen['FECHA CGA'], format='%Y-%m-%d', errors='coerce')
#janssen['Días Credito'] = pd.to_numeric(janssen['Días Credito'], errors='coerce')
# --- Calcular PAGO PARCIAL ---
#janssen['PAGO PARCIAL'] = janssen['PAGO PARCIA'].fillna(0)

# --- Calcular SALDO FACTURA ---
#janssen['SALDO FACTURA'] = janssen['TOTAL FACTURA'] - janssen['PAGO PARCIAL']

# --- Calcular SALDO ACUMULADO incluyendo ANTICIPO ---
#janssen['SALDO ACUMULADO'] = janssen['SALDO FACTURA'] + janssen['ANTICIPO'].fillna(0)

# --- Calcular FECHA LIMITE DE PAGO ---
janssen['FECHA LIMITE DE PAGO'] = janssen['FECHA'] + pd.to_timedelta(janssen['Días Credito'].fillna(0), unit='D')

# --- Calcular DÍAS VENCIDOS ---
hoy = pd.to_datetime(datetime.now().date())
janssen['Días VENCIDOS'] = (hoy - janssen['FECHA LIMITE DE PAGO']).dt.days

# --- Calcular ESTATUS ---
def calcular_estatus(vencidos):
    if pd.isna(vencidos):
        return "SIN FECHA"
    elif vencidos >= 0:
        return "VENCIDO"
    else:
        return "CORRIENTE"

janssen['ESTATUS'] = janssen['Días VENCIDOS'].apply(calcular_estatus)

# --- Ordenar por FECHA ---
janssen = janssen.sort_values(by='FECHA')

# --- Formatear fechas para reportes (solo para mostrar, no afecta cálculos) ---
janssen['FECHA'] = janssen['FECHA'].dt.strftime('%d/%m/%Y')
janssen['FECHA LIMITE DE PAGO'] = janssen['FECHA LIMITE DE PAGO'].dt.strftime('%d/%m/%Y')
fecha_actual = datetime.now().strftime("%d/%m/%Y")
fecha_actual = datetime.now().strftime("%A, %d de %B de %Y")

# --- Seleccionar columnas deseadas ---
columnas_deseadas = [
    'FACTURA', 'ESTATUS CGA', 'FECHA', 'REFERENCIA','PEDIMENTO', 'MERCANCIA', 'PEDIDO',
    'OBSERVACIONES', 'OPERACION', 'CLIENTE', 'PAGOS TERCEROS', 'HONORARIOS ABC',
    'COMPLEMENTARIOS', 'IVA', 'LIQUIDACION / ANTICIPO', 'TOTAL FACTURA', 'PAGO PARCIAL', 'SALDO FACTURA',
    'Días VENCIDOS', 'ESTATUS', 'SALDO ACUMULADO', 'FECHA LIMITE DE PAGO', 'UUID'
]
# Limpia nombres de columnas
janssen.columns = janssen.columns.str.strip()


with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
    hoja_creada = False

    for clasificacion in ["LOC", "GMS", "Sin Ubicación"]:
        if "CLASIFICACIÓN" not in janssen.columns:
            break

        df_clas = janssen[janssen["CLASIFICACIÓN"] == clasificacion]

        if df_clas.empty:
            continue

        # 👉 SOLO COLUMNAS QUE EXISTEN
        columnas_presentes = [c for c in columnas_deseadas if c in df_clas.columns]
        df_clas = df_clas[columnas_presentes]

        df_clas.to_excel(
            writer,
            sheet_name=clasificacion[:31],
            index=False,
            startcol=1,
            startrow=13
        )
        hoja_creada = True

    # 👉 GARANTÍA DE HOJA
    if not hoja_creada:
        pd.DataFrame(
            {"MENSAJE": ["SIN INFORMACIÓN"]}
        ).to_excel(writer, sheet_name="MENSAJE", index=False)

libro = openpyxl.load_workbook(ruta_salida)
for hoja in ["LOC", "GMS", "Sin Ubicación"]:
    Facturas = libro[hoja]

    # --------------------------------------------------
    # FORMATO GENERAL
    # --------------------------------------------------

    for row in range(3, 8):
        for col in ['E', 'F']:
            cell = Facturas[f'{col}{row}']
            cell.font = Font(size=10, bold=True, color='000000')

    Facturas.merge_cells(start_row=3, start_column=6, end_row=3, end_column=8)
    Facturas.merge_cells(start_row=5, start_column=6, end_row=5, end_column=8)
	# Crear un borde blanco (líneas blancas)
    borde_blanco = Border(
        left=Side(border_style="thin", color="FFFFFF"),
        right=Side(border_style="thin", color="FFFFFF"),
        top=Side(border_style="thin", color="FFFFFF"),
        bottom=Side(border_style="thin", color="FFFFFF")
    )

    black_border = Border(
        left=Side(border_style="medium", color="000000"),
        right=Side(border_style="medium", color="000000"),
        top=Side(border_style="medium", color="000000"),
        bottom=Side(border_style="medium", color="000000")
    )

    Facturas['W14'] = 'saldo acumulado'
    # --------------------------------------------------
    # ACUMULATIVO
    # --------------------------------------------------
	
    fila_inicio = 15
    fila_fin = Facturas.max_row

    acumulado = 0
    for fila in range(fila_inicio, fila_fin + 1):
        valor = Facturas[f'R{fila}'].value
        if valor is not None and isinstance(valor, (int, float)):
            acumulado += valor
            Facturas[f'W{fila}'] = acumulado
            #Facturas[f'W{fila}'].number_format = u'[$$-409]#,##0.00;[Rojo]\-[$$-409]#,##0.00'
        else:
            Facturas[f'W{fila}'] = None


    # --------------------------------------------------
    # BORDES FILAS 1–13 COLS B–X
    # --------------------------------------------------

    for row in range(1, 14):
        for col in range(1, 25):
            cell = Facturas[f"{get_column_letter(col)}{row}"]
            cell.border = borde_blanco

    bold_font = Font(bold=True)
    center_alignment = Alignment(horizontal="center", vertical="center")
    fill_yellow = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")

    CLIENTE = janssen['CLIENTE'].dropna().unique()[0] if not janssen.empty else 'N/A'
    credito = janssen['Días Credito'].dropna().unique()[0] if not janssen.empty else 0
    Facturas['F2'] = 'ESTADO DE CUENTA'
    Facturas['F3'] = fecha_actual
    Facturas['F4'] = int(3942)
    Facturas['F5'] = CLIENTE
    Facturas['F6'] = int(credito)

    Facturas['E3'] = 'FECHA DE CORTE'
    Facturas['E4'] = 'PATENTE'
    Facturas['E5'] = 'TITULAR'
    Facturas['E6'] = 'Dias de credito'
    Facturas['D8'] = 'Resumen'


    for cell in ['F3','F5','G3','H3','G5','H5']:
        c = Facturas[cell]
        c.alignment = center_alignment
        c.fill = fill_yellow
        c.border = black_border

    for cell in ['F2','F4','F6']:
        c = Facturas[cell]
        c.border = black_border

    # --------------------------------------------------
    # FILA 14 FORMATO
    # --------------------------------------------------

    Facturas.row_dimensions[14].height = 40  # Puedes cambiar 40 por la altura que desees

    fila_a_colorear = 14
    color_fila = PatternFill(start_color='aebdff', end_color='aebdff', fill_type='solid')
    Facturas.row_dimensions[14].height = 40

    for col in range(2, 26):
        celda = f"{get_column_letter(col)}{fila_a_colorear}"
        Facturas[celda].fill = color_fila
        Facturas[celda].font = Font(size=10, bold=True, color='000000')
        Facturas[celda].alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')

    # --------------------------------------------------
    # MINITABLA
    # --------------------------------------------------

    Facturas.merge_cells(start_row=9, start_column=4, end_row=9, end_column=5)
    Facturas.merge_cells(start_row=10, start_column=4, end_row=10, end_column=5)
    Facturas.merge_cells(start_row=11, start_column=4, end_row=11, end_column=5)
    Facturas.merge_cells(start_row=12, start_column=4, end_row=12, end_column=5)

    Facturas['D9'] = 'CONCEPTO'
    Facturas['F9'] = 'IMPORTE'
    Facturas['G9'] = '%'
    Facturas['D10'] = 'TOTAL CARTERA'
    Facturas['D11'] = 'TOTAL VENCIDO'
    Facturas['D12'] = 'TOTAL CORRIENTE'

    header_fill = PatternFill(start_color='000000', end_color='000000', fill_type='solid')
    header_font = Font(size=10, bold=True, color='FFFFFF')

    for letra in ['D9','F9','G9']:
        Facturas[letra].fill = header_fill
        Facturas[letra].font = header_font

    body_font = Font(size=10, bold=True, color='000000')

    for row in range(10, 13):
        for col in range(4, 8):
            cell = Facturas[f"{get_column_letter(col)}{row}"]
            cell.font = body_font

    for row in range(9, 13):
        for col in range(4, 8):
            cell = Facturas[f"{get_column_letter(col)}{row}"]
            cell.border = black_border
            cell.alignment = Alignment(horizontal='center', vertical='center')

    for row in range(3, 7):
        cell = Facturas[f"F{row}"]
        cell.border = black_border
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # --------------------------------------------------
    # COLUMNAS
    # --------------------------------------------------

    widths = {
        'A':2,'B':12,'C':12,'D':12,'E':12,'F':18,'G':29,'H':14,'I':66,'J':12,'K':12,
        'L':12,'M':12,'N':12,'O':12,'P':12,'Q':12,'R':12,'S':12,'T':12,'U':12,'V':20,'W':12,'X':12
    }
    for col, w in widths.items():
        Facturas.column_dimensions[col].width = w

    Facturas['G10'].number_format = '0%'
    Facturas['G11'].number_format = '0%'
    Facturas['G12'].number_format = '0%'

    # --------------------------------------------------
    # LOGO
    # --------------------------------------------------

    Facturas['K4'] = 'ABSOLUTE BROKERAGE CUSTOMS, S.C.'
    Facturas['K4'].font = Font(size=20, bold=True, color='000000')

    
    imagen_path = os.path.join(BASE_DIR, "ABC.png")
    img_pil = PILImage.open(imagen_path)
    img_pil = img_pil.resize((132, 151))
    img_pil.save("imagen_redimensionada.png")

    img_excel = XLImage("imagen_redimensionada.png")
    Facturas.add_image(img_excel, "I2")

    # --------------------------------------------------
    # CALCULAR TOTALES
    # --------------------------------------------------

for hoja in ["LOC", "GMS","Sin Ubicación"]:
    Facturas = libro[hoja]

    # FILTRAR DATOS PARA ESA HOJA
    df_filtrado = janssen[
        (janssen['ESTATUS CGA'] == 'FACTURADA') &
        (janssen['CLASIFICACIÓN'] == hoja)
    ].copy()

    df_filtrado['ESTATUS_NORM'] = df_filtrado['ESTATUS'].astype(str).str.strip().str.upper()

    df_filtrado['SALDO FACTURA'] = pd.to_numeric(
        df_filtrado['SALDO FACTURA'].astype(str).str.replace(r'[\$\s,]', '', regex=True),
        errors='coerce'
    ).fillna(0)

    suma_saldo = df_filtrado['SALDO FACTURA'].sum()
    suma_vencido = df_filtrado.loc[df_filtrado['ESTATUS_NORM'] == 'VENCIDO', 'SALDO FACTURA'].sum()
    suma_corriente = df_filtrado.loc[df_filtrado['ESTATUS_NORM'] == 'CORRIENTE', 'SALDO FACTURA'].sum()

    # -----------------------------
    # FORMULAS EXCEL PARA RESUMEN
    # -----------------------------
    formulas = {
        "F10": "=SUM(R:R)",
        "F11": '=SUMIF(T:T,"VENCIDO",R:R)',
        "F12": '=SUMIF(T:T,"CORRIENTE",R:R)'
    }

    for celda, formula in formulas.items():
        Facturas[celda] = formula

    # -----------------------------
    # PORCENTAJES
    # -----------------------------
    porcentaje_vencido = round((suma_vencido / suma_saldo * 100)) if suma_saldo else 0
    porcentaje_corriente = round((suma_corriente / suma_saldo * 100)) if suma_saldo else 0

    Facturas['G10'] = "100%"
    Facturas['G11'] = f"{porcentaje_vencido}%"
    Facturas['G12'] = f"{porcentaje_corriente}%"


    # --------------------------------------------------
    # COLORES POR ESTATUS
    # --------------------------------------------------

    color_rojo = PatternFill(start_color='ff0000', end_color='ff0000', fill_type='solid')
    color_amarillo = PatternFill(start_color='fff300', end_color='fff300', fill_type='solid')

    for fila in range(9, Facturas.max_row + 1):
        exep = Facturas[f'T{fila}'].value
        tiempo = Facturas[f'S{fila}'].value

        if exep == 'VENCIDO':
            Facturas[f'T{fila}'].fill = color_rojo
            Facturas[f'R{fila}'].font = Font(size=10, bold=True)
            Facturas['I11'].fill = color_rojo
            Facturas['J11'] = "FACTURAS VENCIDAS"
        if exep == 'CORRIENTE' and tiempo and tiempo > -7 and tiempo <= 0:
            Facturas[f'T{fila}'].fill = color_amarillo
            Facturas[f'S{fila}'] = 0
            Facturas['J12'] = "FACTURAS A VENCER"
            Facturas['I12'].fill = color_amarillo
        if exep == 'CORRIENTE' and tiempo and tiempo <= -7:
            Facturas[f'S{fila}'] = 0
        
        # Aplicar formato contabilidad SIN símbolo de moneda
        account_fmt = '#,##0.00;[Red]-#,##0.00'  # sin símbolo de pesos, negativos en rojo

        # Encabezado de la columna acumulativo
        # Aplicar formato contabilidad (moneda sin símbolo) a las columnas W,L,M,N,O,P,Q,R,S
        cols = ['W','L','M','N','O','P','Q','R']

        for col in cols:
            for row in range(fila_inicio, Facturas.max_row + 1):
                cell = Facturas[f"{col}{row}"]
                if cell.value is not None:
                    cell.number_format = account_fmt

        # Aplicar formato contabilidad a las celdas F10:F12 (totales) sin símbolo
        for row in range(10, 13):  # 10,11,12
            Facturas[f'F{row}'].number_format = account_fmt
    #color_naraja = PatternFill(start_color='ff9900', end_color='ff9900', fill_type='solid')
    #fila_inicio = 15
    #for fila in range(fila_inicio, Facturas.max_row + 1):
    #    anticipo = Facturas[f'Q{fila}'].value
    #    if anticipo is not None and anticipo > 0:
    #        Facturas[f'Q{fila}'].fill = color_naraja
    #        Facturas['J13'] = "PAGO PARCIAL"
    #        Facturas['I13'].fill = color_naraja

# ================
# GUARDAR ARCHIVO
# ================
libro.save(ruta_salida)

print(f"Archivo 'ESTADO DE CUENTA {CLIENTE}.xlsx' creado exitosamente.")

to = ['judith.limas@takeda.com', 'mario-angel.mora@takeda.com (LOC)', 'erika.aguirre@takeda.com', 'ivan.garcia@takeda.com (LOC)']
cc = ['ymontoya@abcsc.mx','gerenteadmin@abcsc.mx','contraloria@abcsc.mx','sgonzalez@abcsc.mx','administracion@abcsc.mx','cxc@abcsc.mx']
#to=['sgonzalez@abcsc.mx']
#cc = ['ccarbajal@abcsc.mx']

CLIENTE = janssen['CLIENTE'].dropna().unique()[0].upper() if not janssen.empty else 'N/A'
fecha_actual = datetime.now().strftime("%d de %B de %Y").encode('latin1').decode('utf-8').upper()

asunto = f"“ESTADO DE CUENTA AGENCIA ADUANAL ABC- {CLIENTE} al {fecha_actual}”".upper()

cuerpo = f"""
<span style="font-size:14px;color:black;">
Hola buenos dias,
<br><br>
un gusto saludarles 
<br><br>
comparto el estado de cuenta actualizado al {fecha_actual}
<br><br>
cualquier aclaración o duda, favor de contactar al departamento de Cuentas por cobrar:
<br><br>
Coordinador de Cuentas por cobrar:
<br><br>
Geraldine Aguilar Sánchez
<br>
Correo: Administracion@abcsc.mx
<BR>
Telefono: 56-1007-8251
<br><br>
Vieyra Calderon Fernando Daniel
<br>
Analista de Cuentas por cobrar:
<br><br>
Correo: cxc@abcsc.mx
<BR>
Telefono: 56-1007-8251
<br><br>
Saludos
<br><br>
Atte. 
<br><br>
Departamento de BI
<br><br>
No responda a este correo es sólo de ENVÍO
</span><br><br>
"""
firma_html = """
<br><br>
<span style="font-size:11px;color:gray;">
este correo es generado automáticamente por el sistema.
</span>
<br><br>
<br><br>
<img src="cid:firma_bi" width="589" height="138"><br><br>
<i>"The simplicity in total quality"</i><br>
<i>"La simpleza de la Calidad Total"</i><br><br>

<span style="color:green;">
🌱 Por favor solo imprime este correo si es realmente necesario.
Tu ayuda es muy importante para salvar este planeta.<br>
Please only print this mail if it is really necessary.
Your help is very important to save this planet.
</span>

<br><br>

<span style="font-size:11px;color:gray;">
Aviso de Privacidad: En cumplimiento con lo previsto en la Ley Federal de Protección 
de Datos Personales en Posesión de los Particulares (LFPDPPP) y su Reglamento, 
pueden acceder al portal 
<a href="https://www.abcsc.mx">www.abcsc.mx</a>
para consultar nuestro aviso de privacidad.

</span>
"""

import base64
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders
import mimetypes

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
FIRMAIMAGE_PATH = "firma_bi.jpg"

SCOPES = ['https://www.googleapis.com/auth/gmail.send']
CREDENTIALS_FILE = 'credenciales.json'
TOKEN_FILE = 'token.json'


def autenticar():

    creds = None

    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:

        flow = InstalledAppFlow.from_client_secrets_file(
            CREDENTIALS_FILE, SCOPES
        )

        creds = flow.run_local_server(port=0)

        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())

    return creds


def enviar_correo(to, cc, asunto, cuerpo, archivo):

    creds = autenticar()
    service = build('gmail', 'v1', credentials=creds)

    mensaje = MIMEMultipart()

    mensaje['To'] = ', '.join(to)
    mensaje['Cc'] = ', '.join(cc)
    mensaje['Subject'] = asunto

    mensaje.attach(MIMEText(cuerpo + firma_html, 'html'))

    # ADJUNTAR IMAGEN FIRMA
    with open(FIRMAIMAGE_PATH, "rb") as img:

        imagen = MIMEImage(img.read())

        imagen.add_header('Content-ID', '<firma_bi>')
        imagen.add_header('Content-Disposition', 'inline', filename="firma_bi.png")

        mensaje.attach(imagen)

    # ADJUNTAR ARCHIVO

    ctype, encoding = mimetypes.guess_type(archivo)

    if ctype is None:
        ctype = 'application/octet-stream'

    maintype, subtype = ctype.split('/', 1)

    with open(archivo, 'rb') as f:

        parte = MIMEBase(maintype, subtype)
        parte.set_payload(f.read())

    encoders.encode_base64(parte)

    parte.add_header(
        'Content-Disposition',
        'attachment',
        filename=os.path.basename(archivo)
    )

    mensaje.attach(parte)

    raw = base64.urlsafe_b64encode(mensaje.as_bytes()).decode()

    message = {'raw': raw}

    service.users().messages().send(
        userId='me',
        body=message
    ).execute()

#enviar_correo(
#    to=to,
#    cc=cc,
#    asunto=asunto,
#    cuerpo=cuerpo,
#    archivo=ruta_salida
#)



print(f"Archivo 'ESTADO DE CUENTA {CLIENTE}.xlsx' creado exitosamente.")
