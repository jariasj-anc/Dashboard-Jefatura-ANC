import io
import os
import re
import glob
import unicodedata
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

try:
    import pyodbc
except ImportError:
    pyodbc = None

# ==============================================================================
# 1. CONFIGURACIÓN GENERAL Y PARÁMETROS PARA PRODUCCIÓN (UTI)
# ==============================================================================
MODO_DATOS = "EXCEL"  # Opciones: "EXCEL" (piloto) | "SYBASE" (tiempo real ODBC)

# Fechas de corte institucionales (auditoría estática para paridad exacta):
FECHA_CORTE_CARGA = pd.Timestamp("2026-11-13")
FECHA_CORTE_PRESC = pd.Timestamp("2026-11-14")

# --- OPCIÓN DINÁMICA EN TIEMPO REAL (PARA DESPLIEGUE UTI) ---
# FECHA_CORTE_CARGA = pd.Timestamp.today().normalize()
# FECHA_CORTE_PRESC = pd.Timestamp.today().normalize()

st.set_page_config(
    page_title="ANC - Análisis Estadístico Oficial (UIA)",
    page_icon="⚖️",
    layout="wide"
)

# ==============================================================================
# 2. ESTILOS CSS CORPORATIVOS
# ==============================================================================
st.markdown("""
<style>
div[data-testid="stHorizontalBlock"] div.stButton > button {
    height: 115px;
    min-height: 115px;
    border-radius: 10px;
    border: 1px solid #1f2937;
    background-color: #111827;
    color: #f8fafc;
    text-align: left !important;
    padding: 0.85rem 1.1rem !important;
    display: flex !important;
    flex-direction: column !important;
    justify-content: center !important;
    align-items: flex-start !important;
    white-space: pre-line;
    line-height: 1.25;
    box-shadow: 0 2px 6px rgba(0, 0, 0, 0.25);
    transition: all 0.2s ease;
}
div[data-testid="stHorizontalBlock"] div.stButton > button > div {
    width: 100%;
    text-align: left !important;
    display: flex;
    flex-direction: column;
    align-items: flex-start;
}
div[data-testid="stHorizontalBlock"] div.stButton > button:hover {
    border-color: #ef4444;
    background-color: #1e293b;
    transform: translateY(-2px);
}
div[data-testid="stHorizontalBlock"] div.stButton > button p {
    font-size: 0.78rem;
    font-weight: 700;
    color: #94a3b8;
    letter-spacing: 0.03em;
    margin: 0;
    text-align: left !important;
    width: 100%;
}
div[data-testid="stHorizontalBlock"] div.stButton > button p strong {
    font-size: 1.85rem;
    font-weight: 800;
    color: #ffffff;
    display: block;
    margin: 3px 0 5px 0;
    line-height: 1.1;
}
div[data-testid="stHorizontalBlock"] div.stButton > button p em {
    font-style: normal;
    display: inline-block;
    background-color: rgba(34, 197, 94, 0.16);
    color: #4ade80;
    font-size: 0.78rem;
    font-weight: 600;
    padding: 2px 9px;
    border-radius: 999px;
    letter-spacing: 0.01em;
}

/* Estilos para los 6 botones semáforo de Prescripción */
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(6)) div.stButton > button {
    height: 122px;
    min-height: 122px;
    padding: 12px 12px !important;
    border: 1px solid #334155;
    margin-bottom: 10px;
}
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(6)) div.stButton > button p {
    font-size: 13px;
    font-weight: 700;
    color: #f8fafc;
    letter-spacing: 0.01em;
}
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(6)) div.stButton > button p strong {
    font-size: 1.45rem;
    margin: 4px 0 4px 0;
}
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(6)) div.stButton > button p em {
    background-color: transparent !important;
    padding: 0 !important;
    border-radius: 0 !important;
    font-size: 11.5px;
    font-weight: 600;
    line-height: 1.3;
}
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(6)) > div:nth-child(1) div.stButton > button {
    background: rgba(239, 68, 68, 0.14);
    border-left: 5px solid #ef4444;
}
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(6)) > div:nth-child(1) div.stButton > button p em { color: #ef4444; }

div[data-testid="stHorizontalBlock"]:has(> div:nth-child(6)) > div:nth-child(2) div.stButton > button {
    background: rgba(250, 204, 21, 0.14);
    border-left: 5px solid #facc15;
}
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(6)) > div:nth-child(2) div.stButton > button p em { color: #facc15; }

div[data-testid="stHorizontalBlock"]:has(> div:nth-child(6)) > div:nth-child(3) div.stButton > button {
    background: rgba(101, 163, 13, 0.14);
    border-left: 5px solid #65a30d;
}
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(6)) > div:nth-child(3) div.stButton > button p em { color: #84cc16; }

div[data-testid="stHorizontalBlock"]:has(> div:nth-child(6)) > div:nth-child(4) div.stButton > button {
    background: rgba(37, 99, 235, 0.14);
    border-left: 5px solid #2563eb;
}
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(6)) > div:nth-child(4) div.stButton > button p em { color: #60a5fa; }

div[data-testid="stHorizontalBlock"]:has(> div:nth-child(6)) > div:nth-child(5) div.stButton > button {
    background: rgba(6, 182, 212, 0.14);
    border-left: 5px solid #06b6d4;
}
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(6)) > div:nth-child(5) div.stButton > button p em { color: #22d3ee; }

div[data-testid="stHorizontalBlock"]:has(> div:nth-child(6)) > div:nth-child(6) div.stButton > button {
    background: rgba(139, 92, 246, 0.14);
    border-left: 5px solid #8b5cf6;
}
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(6)) > div:nth-child(6) div.stButton > button p em { color: #a78bfa; }

div[data-testid="stMetric"] {
    background-color: #111827;
    border: 1px solid #1f2937;
    padding: 0.85rem 1.1rem;
    border-radius: 10px;
    height: 115px;
    min-height: 115px;
    display: flex;
    flex-direction: column;
    justify-content: center;
}
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 3. CONSTANTES Y REGLAS DE NEGOCIO INSTITUCIONALES (UIA)
# ==============================================================================
MESES_ABR = {
    1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Ago", 9: "Set", 10: "Oct", 11: "Nov", 12: "Dic"
}
MESES_NOM = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

COLUMNAS_EXACTAS_DETALLE = [
    "N°", "Distrito Actual", "Número", "Año", "Tipo", "Distrito Judicial",
    "Encargado Actual", "Fecha creación. Exp.", "Estado Exp.", "Nota de Inventario",
    "Personal Asignado Actual", "Tipo de Atención", "Despacho o Unidad",
    "Etapa", "Subetapa", "Fecha Envio/Recep. Despacho", "Magistrado"
]

# Orden procesal estricto aplicado en todas las secciones y gráficos
ORDEN_ETAPAS = [
    "Calificación",
    "Instrucción",
    "Sanción",
    "Apelación",
    "Otros (Abstenciones, incidentes, etc.)"
]

ETAPAS_CORTAS = {
    "Calificación": "Calificación",
    "Instrucción": "Instrucción",
    "Sanción": "Sanción",
    "Apelación": "Apelación",
    "Otros (Abstenciones, incidentes, etc.)": "Otros"
}

ORDEN_RANGOS_PRESC = [
    "Prescritos (<= 0 días)",
    "Crítico: De 1 a 15 días",
    "Alto: De 16 a 30 días",
    "Moderado: De 31 a 60 días",
    "Preventivo: De 61 a 90 días",
    "Monitoreo: De 91 a 240 días"
]

CONFIG_COLORES_RANGOS = [
    ("PRESC_RANGO_0", "Prescritos (<= 0 días)", "Prescritos", "<= 0 días", "#ef4444", "🔴"),
    ("PRESC_RANGO_1", "Crítico: De 1 a 15 días", "Crítico", "De 1 a 15 días", "#facc15", "🟡"),
    ("PRESC_RANGO_2", "Alto: De 16 a 30 días", "Alto", "De 16 a 30 días", "#65a30d", "🟢"),
    ("PRESC_RANGO_3", "Moderado: De 31 a 60 días", "Moderado", "De 31 a 60 días", "#2563eb", "🔵"),
    ("PRESC_RANGO_4", "Preventivo: De 61 a 90 días", "Preventivo", "De 61 a 90 días", "#06b6d4", "🔵"),
    ("PRESC_RANGO_5", "Monitoreo: De 91 a 240 días", "Monitoreo", "De 91 a 240 días", "#8b5cf6", "🟣"),
]

# ==============================================================================
# 4. MOTOR AUTOMÁTICO DE NORMALIZACIÓN Y CÁLCULOS
# ==============================================================================
def norm_txt(val):
    if pd.isna(val):
        return ""
    s = str(val).strip().upper()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    if "SEDE CENTRAL" in s:
        return "A.N.C. SEDE CENTRAL"
    return s

def limpiar_columnas(df):
    df.columns = [str(c).strip() for c in df.columns]
    return df.loc[:, ~df.columns.duplicated()].copy()

def encontrar_columna(df, candidatos_norm):
    mapa = {norm_txt(c): c for c in df.columns}
    for cand in candidatos_norm:
        cand_n = norm_txt(cand)
        if cand_n in mapa:
            return mapa[cand_n]
    for cand in candidatos_norm:
        cand_n = norm_txt(cand)
        for k_norm, col_real in mapa.items():
            if cand_n in k_norm:
                return col_real
    return None

def categorizar_etapa_auto(etapa_val):
    e = norm_txt(etapa_val)
    if e in ("", "NAN", "NONE", "NULL"):
        return "Otros (Abstenciones, incidentes, etc.)"
    if "CALIFICAC" in e:
        return "Calificación"
    if "INSTRUCC" in e or "INVESTIGAC" in e:
        return "Instrucción"
    if "SANCION" in e:
        return "Sanción"
    if "APELAC" in e:
        return "Apelación"
    return "Otros (Abstenciones, incidentes, etc.)"

def subcat_calificacion(sub_val):
    s = norm_txt(sub_val)
    if s == "CALIFICACION":
        return "   • Calificación propiamente dicha"
    if s == "APELACION":
        return "   • Apelación (en Calificación)"
    return "   • Pend. Informe / Por Elevar / En Recepción / Otros"

def clasificar_presc_rango(dias):
    if pd.isna(dias):
        return None
    if dias <= 0:
        return "Prescritos (<= 0 días)"
    if dias <= 15:
        return "Crítico: De 1 a 15 días"
    if dias <= 30:
        return "Alto: De 16 a 30 días"
    if dias <= 60:
        return "Moderado: De 31 a 60 días"
    if dias <= 90:
        return "Preventivo: De 61 a 90 días"
    if dias <= 240:
        return "Monitoreo: De 91 a 240 días"
    return None

def formatear_detalle_exp_presc(row):
    raw_exp = str(row.get("N° EXPEDIENTE", "")).strip()
    base_exp = re.split(r'\s+de\s+Origen |\s+DE\s+ORIGEN ', raw_exp)[0].strip()

    m = re.match(r'^\s*(\d+)\s*-\s*(\d{4})\s*/\s*([^/]+)\s*/\s*(.+?)\s*$', base_exp)
    if m:
        num_e, anio_e, tipo_e, dist_j = m.groups()
        return f"{num_e}-{anio_e} | {tipo_e.strip().title()} | {dist_j.strip().title()}"

    num_c = str(row.get("NUMERO", row.get("Número", ""))).strip()
    anio_c = str(row.get("AÑO", row.get("Año", ""))).strip()
    tipo_c = str(row.get("TIPO", row.get("Tipo", ""))).strip()
    dist_c = str(row.get("DISTRITO JUDICIAL", row.get("Distrito Judicial", ""))).strip()
    if num_c and anio_c and num_c not in ("-", "None", "nan"):
        partes = [f"{num_c}-{anio_c}"]
        if tipo_c and tipo_c not in ("-", "None", "nan"):
            partes.append(tipo_c.title())
        if dist_c and dist_c not in ("-", "None", "nan"):
            partes.append(dist_c.title())
        return " | ".join(partes)

    return base_exp if base_exp else "-"

def map_unidad_produccion(row):
    s = norm_txt(row.get("ENCARGADO", ""))
    if row.get("Es_Sede_Central", False):
        if "JEFATURA" in s:
            return "JEFATURA ANC"
        m = re.search(r'-(UCIP|UPAD|USAP|OCPAD|UPSI|CII|OCCPC) ', s)
        return m.group(1) if m else (s if s else "NO ESPECIFICADO")
    else:
        if "JEFATURA" in s:
            return "Jefatura ODANC"
        if any(k in s for k in ("CALIFICACION", "UDCIP", "UCIP", "CALIF")):
            return "Calificación e Inv. Preliminar (UDCIP/UCIP)"
        if any(k in s for k in ("PROCEDIMIENTO", "UDPAD", "UPAD", "OCPAD", "PAD", "DISCIPLINARIO", "INSTRUCCION")):
            return "Procedimiento Adm. Disciplinario (UDPAD/UPAD)"
        if any(k in s for k in ("SANCION", "APELACION", "UDSAP", "USAP")):
            return "Sanción y Apelación (UDSAP/USAP)"
        if any(k in s for k in ("PREVENCION", "SUPERVISION", "INSPECCION", "UPSI", "UDPSI")):
            return "Prevención, Supervisión e Inspección (UDPSI/UPSI)"
        if "INTERVENCION INMEDIATA" in s or "-CII" in s:
            return "Comisión de Intervención Inmediata (CII)"
        return "Otros / Especializados"

def map_unidad_todos(enc):
    s = norm_txt(enc)
    if "JEFATURA" in s:
        return "Jefatura (ANC / ODANC)"
    if any(k in s for k in ("CALIFICACION", "UDCIP", "UCIP", "CALIF")):
        return "Calificación e Inv. Preliminar (UCIP / UDCIP)"
    if any(k in s for k in ("PROCEDIMIENTO", "UDPAD", "UPAD", "OCPAD", "PAD", "DISCIPLINARIO", "INSTRUCCION")):
        return "Procedimiento Adm. Disciplinario (UPAD / OCPAD / UDPAD)"
    if any(k in s for k in ("SANCION", "APELACION", "UDSAP", "USAP")):
        return "Sanción y Apelación (USAP / UDSAP)"
    if any(k in s for k in ("PREVENCION", "SUPERVISION", "INSPECCION", "UPSI", "UDPSI")):
        return "Prevención, Supervisión e Inspección (UPSI / UDPSI)"
    if "OCCPC" in s or "CONTROL PREVENTIVO" in s:
        return "Control Preventivo y Concurrente (OCCPC)"
    if "INTERVENCION INMEDIATA" in s or "-CII" in s:
        return "Comisión de Intervención Inmediata (CII)"
    return "Otros / Especializados"

def map_magistrado_produccion(row):
    mag = str(row.get("MAGISTRADO", "")).strip().upper()
    if mag in ("", "NAN", "NONE", "NULL"):
        mag = "NO ESPECIFICADO"
    enc = norm_txt(row.get("ENCARGADO", ""))
    if row.get("Es_Sede_Central", False) and ("JEFATURA" in enc or "PALACIOS BRAN" in mag):
        return "JEFATURA SUPREMA"
    return mag

def estandarizar_tablas(df_c_raw, df_p_raw, df_pr_raw):
    # A. CARGA ACTUAL
    df_c_orig = limpiar_columnas(df_c_raw.copy())
    mask_no_inv_global = df_c_orig.apply(
        lambda fila: fila.astype(str).str.upper().str.contains("SIN INVENTARIAR", na=False).any(),
        axis=1
    )

    cols_base_c = [
        "Distrito Actual", "Número", "Año", "Tipo", "Distrito Judicial",
        "Encargado Actual", "Fecha creación. Exp.", "Estado Exp.", "Es Carga",
        "Nota de Inventario", "Personal Asignado Actual", "Tipo de Atención",
        "Despacho o Unidad", "Etapa", "Subetapa", "Fecha Envio/Recep. Despacho",
        "Magistrado"
    ]
    cols_c_ok = [c for c in cols_base_c if c in df_c_orig.columns]
    df_c = df_c_orig[cols_c_ok].copy()

    for col_num in ["Número", "Año"]:
        if col_num in df_c.columns:
            df_c[col_num] = (
                df_c[col_num]
                .fillna("-")
                .astype(str)
                .str.replace(r"\.0$", "", regex=True)
                .str.strip()
            )

    df_c["Es_No_Inventariado"] = mask_no_inv_global
    df_c["Distrito_Norm"] = df_c["Distrito Actual"].apply(norm_txt)
    df_c["Es_Sede_Central"] = df_c["Distrito_Norm"] == "A.N.C. SEDE CENTRAL"
    df_c["Estado_Norm"] = df_c["Estado Exp."].apply(norm_txt)

    col_etapa_c = encontrar_columna(df_c, ["Etapa"]) or "Etapa"
    df_c["Etapa_Cat"] = df_c[col_etapa_c].apply(categorizar_etapa_auto)
    df_c["Subetapa_Raw"] = df_c["Subetapa"].fillna("SIN SUBETAPA").astype(str).str.strip()
    df_c["Subetapa_Cat"] = df_c["Subetapa"].apply(subcat_calificacion)
    df_c["Despacho_Norm"] = df_c["Despacho o Unidad"].apply(norm_txt)
    df_c["Anio_Num"] = pd.to_numeric(df_c["Año"], errors="coerce").fillna(0).astype(int)

    df_c["Fecha_Creacion_Dt"] = pd.to_datetime(df_c["Fecha creación. Exp."], errors="coerce", dayfirst=True).dt.normalize()
    df_c["Antiguedad_Dias"] = (FECHA_CORTE_CARGA - df_c["Fecha_Creacion_Dt"]).dt.days
    df_c["Antiguedad_Anios"] = (df_c["Antiguedad_Dias"] / 365.25).round(2)

    for col_f in ["Fecha creación. Exp.", "Fecha Envio/Recep. Despacho"]:
        if col_f in df_c.columns:
            f_dt = pd.to_datetime(df_c[col_f], errors="coerce", dayfirst=True)
            df_c[col_f] = f_dt.dt.strftime("%d/%m/%Y").fillna("-")

    for col_txt in COLUMNAS_EXACTAS_DETALLE:
        if col_txt in df_c.columns and col_txt != "N°":
            df_c[col_txt] = (
                df_c[col_txt]
                .replace({None: "-", "None": "-", "nan": "-", "NaN": "-", "NAT": "-", "NaT": "-", "": "-"})
                .fillna("-")
                .astype(str)
                .str.strip()
            )
            df_c[col_txt] = df_c[col_txt].replace({"": "-", "None": "-", "nan": "-"})

    # B. PRESCRIPCIONES (<= 240 días)
    df_p = limpiar_columnas(df_p_raw.copy())
    cols_manuales_obsoletas = [
        c for c in df_p.columns
        if any(k in norm_txt(c) for k in ("DIAS", "RANGO", "CATEGORIA", "SEMAFORO"))
        and "FECHA" not in norm_txt(c)
    ]
    if cols_manuales_obsoletas:
        df_p = df_p.drop(columns=cols_manuales_obsoletas)

    col_exp_p = encontrar_columna(df_p, ["N° EXPEDIENTE", "EXPEDIENTE"])
    if col_exp_p:
        df_p = df_p[
            df_p[col_exp_p].notna() &
            (~df_p[col_exp_p].astype(str).str.strip().str.upper().isin(["", "NONE", "NAN"]))
        ].copy()

    col_dist_p = encontrar_columna(df_p, ["DISTRITO ACTUAL"]) or "DISTRITO ACTUAL"
    df_p["Distrito_Norm"] = df_p[col_dist_p].apply(norm_txt)
    df_p["Es_Sede_Central"] = df_p["Distrito_Norm"] == "A.N.C. SEDE CENTRAL"

    col_etapa_p = encontrar_columna(df_p, ["ETAPA"])
    df_p["Etapa_Presc_Cat"] = df_p[col_etapa_p].apply(categorizar_etapa_auto) if col_etapa_p else "Otros (Abstenciones, incidentes, etc.)"

    col_f_presc = encontrar_columna(df_p, ["FECHA DE PRESCRIPCION", "FECHA PRESCRIPCION"]) or "FECHA DE PRESCRIPCION"
    df_p["Fecha_Presc_Dt"] = pd.to_datetime(
        df_p[col_f_presc].astype(str).str.strip(),
        dayfirst=True,
        errors="coerce"
    ).dt.normalize()

    col_ap = encontrar_columna(df_p, ["ATENCION PREFERENCIAL", "PREFERENCIAL"])
    df_p["Atencion_Pref_Norm"] = df_p[col_ap].apply(norm_txt) if col_ap else "SIN ATENCION PREFERENCIAL"
    df_p["Es_Atencion_Pref"] = ~df_p["Atencion_Pref_Norm"].isin(["", "SIN ATENCION PREFERENCIAL", "NO"])

    col_mag_p = encontrar_columna(df_p, ["MAGISTRADO"])
    df_p["MAGISTRADO"] = df_p[col_mag_p].fillna("-") if col_mag_p else "-"

    df_p["Dias_Calc"] = (df_p["Fecha_Presc_Dt"] - FECHA_CORTE_PRESC).dt.days
    df_p["Rango_Presc"] = df_p["Dias_Calc"].apply(clasificar_presc_rango)
    df_p = df_p[df_p["Rango_Presc"].notna()].copy()

    # C. PRODUCCIÓN MENSUAL
    df_pr = limpiar_columnas(df_pr_raw.copy())
    cols_pr_ok = [
        c for c in [
            "Distrito Actual", "EXPEDIENTE", "ENCARGADO", "MAGISTRADO",
            "INFORME", "RESOLUCIÓN", "FECHA_DESCARGA",
            "DETALLE DEL TIPO DE PRONUNCIAMIENTO", "TIPO DE AUTO"
        ] if c in df_pr.columns
    ]
    df_pr = df_pr[cols_pr_ok].copy()
    df_pr["Distrito_Norm"] = df_pr["Distrito Actual"].apply(norm_txt)
    df_pr["Es_Sede_Central"] = df_pr["Distrito_Norm"] == "A.N.C. SEDE CENTRAL"
    df_pr["Fecha_Desc_Dt"] = pd.to_datetime(df_pr["FECHA_DESCARGA"], errors="coerce", dayfirst=True)
    df_pr["Es_Informe"] = (
        df_pr["INFORME"].notna() & (df_pr["INFORME"].astype(str).str.strip() != "")
        if "INFORME" in df_pr.columns else False
    )

    if "ENCARGADO" in df_pr.columns:
        df_pr["Unidad_Prod"] = df_pr.apply(map_unidad_produccion, axis=1)
        df_pr["Unidad_Prod_Todos"] = df_pr["ENCARGADO"].apply(map_unidad_todos)
    else:
        df_pr["Unidad_Prod"] = "NO ESPECIFICADO"
        df_pr["Unidad_Prod_Todos"] = "NO ESPECIFICADO"

    if "MAGISTRADO" in df_pr.columns:
        df_pr["Magistrado_Prod"] = df_pr.apply(map_magistrado_produccion, axis=1)
    else:
        df_pr["Magistrado_Prod"] = "NO ESPECIFICADO"

    return df_c, df_p, df_pr

# ==============================================================================
# 5. CONECTORES DE DATOS (EXCEL / SYBASE ODBC EN TIEMPO REAL)
# ==============================================================================
@st.cache_data(show_spinner="Procesando base de datos institucional...")
def cargar_desde_excel(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    hojas = {norm_txt(h): h for h in xls.sheet_names}
    h_carga = next((v for k, v in hojas.items() if "CARGA" in k), xls.sheet_names[0])
    h_prod = next((v for k, v in hojas.items() if "PRODUCCION" in k), xls.sheet_names[1])
    h_presc = next((v for k, v in hojas.items() if "PRESCRIP" in k), xls.sheet_names[2])
    return estandarizar_tablas(
        pd.read_excel(xls, sheet_name=h_carga),
        pd.read_excel(xls, sheet_name=h_presc),
        pd.read_excel(xls, sheet_name=h_prod)
    )

@st.cache_data(ttl=300, show_spinner="Consultando base de datos Sybase en tiempo real...")
def cargar_desde_sybase():
    if pyodbc is None:
        st.error("La librería pyodbc no está instalada. Ejecuta: pip install pyodbc")
        st.stop()
    conn_str = (
        f"DRIVER={{Adaptive Server Enterprise}};"
        f"SERVER={os.getenv('SYBASE_HOST', 'IP_SERVIDOR_SYBASE')};"
        f"PORT={os.getenv('SYBASE_PORT', '5000')};"
        f"DATABASE={os.getenv('SYBASE_DB', 'BD_ANC')};"
        f"UID={os.getenv('SYBASE_USER', 'USUARIO_SOLO_LECTURA')};"
        f"PWD={os.getenv('SYBASE_PASS', 'CLAVE_SEGURA')};"
    )
    with pyodbc.connect(conn_str) as conn:
        df_c_sql = pd.read_sql("SELECT * FROM vw_anc_carga_actual", conn)
        df_p_sql = pd.read_sql("SELECT * FROM vw_anc_prescripciones", conn)
        df_pr_sql = pd.read_sql("SELECT * FROM vw_anc_produccion", conn)
    return estandarizar_tablas(df_c_sql, df_p_sql, df_pr_sql)

if MODO_DATOS == "EXCEL":
    archivos_excel = [f for f in glob.glob("*.xlsx") if not os.path.basename(f).startswith("~$")]
    if not archivos_excel:
        st.error("No se encontró ningún archivo Excel (.xlsx) en el directorio.")
        st.stop()
    with open(archivos_excel[0], "rb") as f:
        df_carga, df_presc, df_prod = cargar_desde_excel(f.read())
else:
    df_carga, df_presc, df_prod = cargar_desde_sybase()

# ==============================================================================
# 6. UTILIDADES DE EXPORTACIÓN Y RENDERIZADO DE TABLAS / TARJETAS
# ==============================================================================
def limpiar_latin1(txt):
    s = (
        str(txt)
        .replace("•", "-")
        .replace("🔴", "").replace("🟡", "").replace("🟢", "")
        .replace("🔵", "").replace("🟠", "").replace("🟣", "")
    )
    return s.encode("latin-1", "replace").decode("latin-1").strip()

def exportar_detalle_excel(df_exp: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_exp.to_excel(writer, index=False, sheet_name="Detalle_Expedientes")
        ws = writer.sheets["Detalle_Expedientes"]
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col) + 3
            ws.column_dimensions[col[0].column_letter].width = min(max_len, 45)
    return buffer.getvalue()

def exportar_detalle_pdf(df_exp: pd.DataFrame, titulo: str = "Detalle de Expedientes") -> bytes:
    from fpdf import FPDF
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=10)
    pdf.add_page()

    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 8, limpiar_latin1(titulo), ln=True, align="L")
    pdf.ln(2)

    cols = df_exp.columns.tolist()
    n_cols = min(len(cols), 12)
    cols_show = cols[:n_cols]
    ancho_total = 277.0
    w_prim = 10.0
    w_resto = (ancho_total - w_prim) / max(1, (n_cols - 1))
    anchos = [w_prim] + [w_resto] * (n_cols - 1)

    pdf.set_font("Helvetica", "B", 6.5)
    pdf.set_fill_color(21, 94, 138)
    pdf.set_text_color(255, 255, 255)
    for i, col_name in enumerate(cols_show):
        max_h = max(3, int(anchos[i] / 1.8))
        pdf.cell(anchos[i], 7, limpiar_latin1(str(col_name)[:max_h]), border=1, align="C", fill=True)
    pdf.ln()

    pdf.set_font("Helvetica", "", 6)
    pdf.set_text_color(0, 0, 0)
    for _, row in df_exp.iterrows():
        for i, col_name in enumerate(cols_show):
            val = str(row[col_name]) if pd.notna(row[col_name]) else "-"
            max_chars = max(4, int(anchos[i] / 1.75))
            align = "C" if i == 0 or any(k in str(col_name).upper() for k in ("FECHA", "DIAS", "DÍAS")) else "L"
            pdf.cell(anchos[i], 6, limpiar_latin1(val[:max_chars]), border=1, align=align)
        pdf.ln()

    return bytes(pdf.output())

def mostrar_tabla_centrada(df_in):
    if df_in.empty:
        st.info("Sin registros disponibles para mostrar.")
        return

    cols = df_in.columns.tolist()
    cols_num = df_in.select_dtypes(include=["number"]).columns.tolist()

    ths = "".join([
        f"<th style='padding:8px 10px; border:1px solid #334155; text-align:{'left' if (i == 0 and c != 'N°') else 'center'}; color:#f8fafc; white-space:nowrap;'>{c}</th>"
        for i, c in enumerate(cols)
    ])

    filas_html = ""
    for idx, row in df_in.iterrows():
        texto_fila = " ".join([str(row[c]) for c in cols[:2]])
        es_total = any(palabra in texto_fila.upper() for palabra in ["TOTAL", "SUBTOTAL"])
        bg_fila = "#1e293b" if es_total else ("#0f172a" if idx % 2 == 1 else "#111827")
        borde_fila = "1px solid #334155" if es_total else "1px solid #1f2937"
        peso = "700" if es_total else "400"

        celdas = ""
        for i, c in enumerate(cols):
            val = row[c]
            if c in cols_num and pd.notna(val):
                txt_val = f"{val:,.2f}%" if "%" in str(c) else f"{int(val):,}"
            else:
                txt_val = str(val) if pd.notna(val) else ""

            es_texto_izq = (i == 0 and c != "N°") or (i == 1 and cols[0] == "N°")
            alineacion = "left" if es_texto_izq else "center; white-space:nowrap"
            if es_texto_izq and txt_val.startswith("   •"):
                estilo_extra = "padding-left:22px; color:#cbd5e1;"
            elif es_total and i == len(cols) - 2:
                estilo_extra = "color:#fbbf24;"
            elif es_total:
                estilo_extra = "color:#f8fafc;"
            else:
                estilo_extra = "color:#e2e8f0;"

            celdas += f"<td style='padding:7px 10px; border:{borde_fila}; text-align:{alineacion}; font-weight:{peso}; {estilo_extra}'>{txt_val}</td>"

        filas_html += f"<tr style='background:{bg_fila};'>{celdas}</tr>"

    st.markdown(f"""
    <div style="overflow-x:auto; margin-bottom:16px; border-radius:8px; border:1px solid #334155;">
    <table style="width:100%; border-collapse:collapse; font-size:13px; background:#111827;">
    <thead><tr style="background:#1e293b; border-bottom:2px solid #475569;">{ths}</tr></thead>
    <tbody>{filas_html}</tbody>
    </table>
    </div>
    """, unsafe_allow_html=True)

def abreviar_unidad_card(raw_unidad, es_sede_central):
    u = norm_txt(raw_unidad)
    if u in ("", "-", "NONE", "NAN"):
        return "Sin unidad"
    if "JEFATURA" in u:
        return "Jefatura ANC" if es_sede_central else "Jefatura"
    if any(k in u for k in ("SANCION", "APELACION", "UDSAP", "USAP")):
        return "USAP" if es_sede_central else "UDSAP"
    if any(k in u for k in ("CALIFICACION", "UDCIP", "UCIP", "CALIF")):
        return "UCIP" if es_sede_central else "UDCIP"
    if any(k in u for k in ("PROCEDIMIENTO", "UDPAD", "UPAD", "OCPAD", "PAD", "DISCIPLINARIO")):
        return "UPAD" if es_sede_central else "UDPAD"
    if any(k in u for k in ("PREVENCION", "SUPERVISION", "INSPECCION", "UDPSI", "UPSI")):
        return "UPSI" if es_sede_central else "UDPSI"
    if "OCCPC" in u or "CONTROL PREVENTIVO" in u:
        return "OCCPC"
    if "INTERVENCION INMEDIATA" in u or "CII" in u:
        return "CII"
    return str(raw_unidad).strip().title()

def render_tarjeta_mas_antiguo(df_origen, es_nac, titulo_card="EXP. MÁS ANTIGUO"):
    df_fechas = df_origen.dropna(subset=["Fecha_Creacion_Dt"]).sort_values("Fecha_Creacion_Dt", ascending=True)
    if df_fechas.empty:
        mes_anio, unidad_txt, unidad_full, etapa_txt, estado_txt, sede_txt = "Sin datos", "-", "-", "-", "-", ""
        color_est, bg_est = "#94a3b8", "rgba(148, 163, 184, 0.15)"
    else:
        exp = df_fechas.iloc[0]
        dt_f = exp["Fecha_Creacion_Dt"]
        mes_anio = f"{MESES_ABR[dt_f.month]} {dt_f.year}"
        raw_unidad = str(exp.get("Despacho o Unidad", "-")).strip()
        unidad_full = raw_unidad.title() if raw_unidad not in ("", "-", "None", "nan") else "Sin unidad"
        unidad_txt = abreviar_unidad_card(raw_unidad, exp["Es_Sede_Central"])
        etapa_txt = exp["Etapa_Cat"]
        estado_txt = exp["Estado_Grupo"]
        if exp["Es_Sede_Central"]:
            sede_txt = "ANC Sede Central"
        else:
            dist_raw = str(exp.get("Distrito Actual", "")).strip()
            sede_txt = re.sub(r'^\s*ODANC\s*[-:]?\s*', '', dist_raw, flags=re.IGNORECASE).strip().title()
        color_est = "#38bdf8" if estado_txt == "Trámite" else "#fbbf24"
        bg_est = "rgba(56, 189, 248, 0.15)" if estado_txt == "Trámite" else "rgba(251, 191, 36, 0.15)"

    cabecera_sede = (
        f'<span title="{sede_txt}" style="background: rgba(245, 158, 11, 0.18); color: #fde68a; font-weight: 700; font-size: 0.70rem; padding: 1px 6px; border-radius: 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 55%;">📍 {sede_txt}</span>'
        if es_nac and sede_txt else ""
    )

    st.markdown(
        f'<div style="background: linear-gradient(135deg, #111827 0%, #1e1b18 100%); border: 1px solid #334155; border-left: 4px solid #f59e0b; border-radius: 10px; padding: 0.65rem 0.9rem; height: 115px; min-height: 115px; display: flex; flex-direction: column; justify-content: center; box-shadow: 0 2px 6px rgba(0,0,0,0.25);">'
        f'<div style="display: flex; align-items: center; justify-content: space-between; gap: 6px;">'
        f'<span style="font-size: 0.70rem; font-weight: 700; color: #f59e0b; letter-spacing: 0.03em; text-transform: uppercase; flex-shrink: 0;">⏳ {titulo_card}</span>'
        f'{cabecera_sede}'
        f'</div>'
        f'<div style="font-size: 1.55rem; font-weight: 800; color: #ffffff; line-height: 1.15; margin: 3px 0 4px 0;">{mes_anio}</div>'
        f'<div style="display: flex; flex-wrap: nowrap; gap: 5px; align-items: center; overflow: hidden;">'
        f'<span title="{unidad_full}" style="background: rgba(148, 163, 184, 0.16); color: #f8fafc; font-size: 10px; font-weight: 700; padding: 2px 6px; border-radius: 4px; white-space: nowrap; flex-shrink: 0;">🏛️ {unidad_txt}</span>'
        f'<span style="background: rgba(255,255,255,0.08); color: #e2e8f0; font-size: 10px; font-weight: 600; padding: 2px 6px; border-radius: 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">{etapa_txt}</span>'
        f'<span style="background: {bg_est}; color: {color_est}; font-size: 10px; font-weight: 700; padding: 2px 6px; border-radius: 4px; flex-shrink: 0;">{estado_txt}</span>'
        f'</div></div>',
        unsafe_allow_html=True
    )

# ==============================================================================
# 7. FILTRO LATERAL POR SEDE Y CABECERA INSTITUCIONAL
# ==============================================================================
distritos_unicos = sorted([d for d in df_carga["Distrito_Norm"].unique() if d not in ("", "A.N.C. SEDE CENTRAL")])
opciones_filtro = ["TODOS", "A.N.C. SEDE CENTRAL"] + distritos_unicos

with st.sidebar:
    st.subheader("🎯 Filtro por Sede / ODANC")
    dist_sel = st.selectbox("Visualizar:", opciones_filtro, index=1)

es_nacional = (dist_sel == "TODOS")
nombre_ambito = "TODOS (ANC SEDE CENTRAL + 34 ODANCs)" if es_nacional else dist_sel
sub_titulo_sec = "ANC vs. ODANC" if es_nacional else dist_sel

st.markdown(f"""
<div style="background: linear-gradient(90deg, #7f1d1d 0%, #991b1b 60%, #1e293b 100%); padding: 18px 24px; border-radius: 10px; border-left: 6px solid #f59e0b; margin-bottom: 18px;">
<h2 style="color: white; margin: 0; font-size: 22px;">⚖️ Autoridad Nacional de Control del Poder Judicial</h2>
<p style="color: #e2e8f0; margin: 4px 0 0 0; font-size: 13px;">ANÁLISIS ESTADÍSTICO DE CARGA, PRESCRIPCIONES Y PRODUCCIÓN: <b>{nombre_ambito}</b> (Corte Carga: {FECHA_CORTE_CARGA.strftime('%d/%m/%Y')} | Corte Presc.: {FECHA_CORTE_PRESC.strftime('%d/%m/%Y')}) — UIA</p>
</div>
""", unsafe_allow_html=True)

def aplicar_filtro_sede(df):
    if es_nacional:
        return df
    if dist_sel == "A.N.C. SEDE CENTRAL":
        return df[df["Es_Sede_Central"]]
    return df[df["Distrito_Norm"] == dist_sel]

df_c_f = aplicar_filtro_sede(df_carga)
df_p_f = aplicar_filtro_sede(df_presc)
df_pr_f = aplicar_filtro_sede(df_prod)

# Preparar base unificada de la sede (Activos + Sin Inventariar) con las mismas columnas normalizadas
df_c_sede = df_c_f.copy()
df_c_sede["Anio_Grupo"] = np.where(df_c_sede["Anio_Num"] >= 2022, df_c_sede["Anio_Num"].astype(str), "2021 y ant.")
df_c_sede["Despacho_Clean"] = df_c_sede["Despacho o Unidad"].fillna("-").astype(str).str.strip()
df_c_sede["Subetapa_Clean"] = df_c_sede["Subetapa_Raw"].fillna("-").astype(str).str.strip()

df_no_inv = df_c_sede[df_c_sede["Es_No_Inventariado"]].copy()
df_tram_sede = df_c_sede[(~df_c_sede["Es_No_Inventariado"]) & (df_c_sede["Estado_Norm"] == "TRAMITE")].copy()
df_imp_sede = df_c_sede[(~df_c_sede["Es_No_Inventariado"]) & (df_c_sede["Estado_Norm"].str.contains("RESUELTO|IMPUGNACION", na=False))].copy()

df_act_sede = pd.concat([df_tram_sede, df_imp_sede], ignore_index=True)
df_act_sede["Estado_Grupo"] = np.where(df_act_sede["Estado_Norm"] == "TRAMITE", "Trámite", "Res. Impug.")

df_p_act = df_p_f.copy()
df_p_vigente = df_p_act[df_p_act["Dias_Calc"] >= 1].copy()
df_p_atencion_pref = df_p_vigente[df_p_vigente["Es_Atencion_Pref"]].copy()

# ==============================================================================
# 8. CONSTRUCTORES DE CUADROS ESTADÍSTICOS (ORDEN ESTRICTO DE ETAPAS)
# ==============================================================================
def construir_df_etapas_anio(df_in):
    if df_in.empty:
        return pd.DataFrame(), []
    df_v = df_in.copy()
    df_v["Anio_Col"] = np.where(df_v["Anio_Num"] > 0, df_v["Anio_Num"].astype(str), "S/A")
    anios_disp = sorted([a for a in df_v["Anio_Col"].unique() if a != "S/A"])
    if "S/A" in df_v["Anio_Col"].unique():
        anios_disp.append("S/A")

    tot_gen = len(df_v)
    ord_sub = [
        "   • Calificación propiamente dicha",
        "   • Apelación (en Calificación)",
        "   • Pend. Informe / Por Elevar / En Recepción / Otros"
    ]

    def crear_fila(num_str, etiqueta, sub_df, es_sub=False, es_tot=False):
        fila = {"N°": num_str, "Etapa / Sub-etapa Procesal": etiqueta, "_es_sub": es_sub, "_es_total": es_tot}
        for a in anios_disp:
            sub_a = sub_df[sub_df["Anio_Col"] == a]
            fila[f"{a} - Trámite"] = int((sub_a["Estado_Grupo"] == "Trámite").sum())
            fila[f"{a} - Res. Impug."] = int((sub_a["Estado_Grupo"] == "Res. Impug.").sum())
        fila["Total - Trámite"] = int((sub_df["Estado_Grupo"] == "Trámite").sum())
        fila["Total - Res. Impug."] = int((sub_df["Estado_Grupo"] == "Res. Impug.").sum())
        fila["Total"] = len(sub_df)
        fila["%"] = round(len(sub_df) / tot_gen * 100, 1) if tot_gen > 0 else 0.0
        return fila

    filas = []
    for idx_e, et_nom in enumerate(ORDEN_ETAPAS, start=1):
        sub_et = df_v[df_v["Etapa_Cat"] == et_nom]
        filas.append(crear_fila(str(idx_e), et_nom, sub_et))
        if et_nom == "Calificación":
            for sub_nom in ord_sub:
                filas.append(crear_fila("", sub_nom, sub_et[sub_et["Subetapa_Cat"] == sub_nom], es_sub=True))

    filas.append(crear_fila("", "TOTAL POR ETAPAS", df_v, es_tot=True))
    return pd.DataFrame(filas), anios_disp

def construir_df_mes_anio_estado(df_in, nombre_total="TOTAL CALIFICACIÓN"):
    if df_in.empty:
        return pd.DataFrame(), []
    df_v = df_in.copy()
    df_v["Anio_Col"] = np.where(df_v["Anio_Num"] > 0, df_v["Anio_Num"].astype(str), "S/A")
    anios_disp = sorted([a for a in df_v["Anio_Col"].unique() if a != "S/A"])
    if "S/A" in df_v["Anio_Col"].unique():
        anios_disp.append("S/A")

    df_v["Mes_Num"] = df_v["Fecha_Creacion_Dt"].dt.month.fillna(0).astype(int)
    tot_gen = len(df_v)

    meses_presentes = [m for m in range(1, 13) if (df_v["Mes_Num"] == m).any()]
    if (df_v["Mes_Num"] == 0).any():
        meses_presentes.append(0)

    filas = []
    for idx_m, m_num in enumerate(meses_presentes, start=1):
        sub_m = df_v[df_v["Mes_Num"] == m_num]
        etiq_mes = MESES_NOM[m_num] if m_num in MESES_NOM else "Sin fecha"
        fila = {"N°": str(idx_m), "Mes de Creación": etiq_mes, "_es_total": False}
        for a in anios_disp:
            sub_ma = sub_m[sub_m["Anio_Col"] == a]
            fila[f"{a} - Trámite"] = int((sub_ma["Estado_Grupo"] == "Trámite").sum())
            fila[f"{a} - Res. Impug."] = int((sub_ma["Estado_Grupo"] == "Res. Impug.").sum())
        fila["Total - Trámite"] = int((sub_m["Estado_Grupo"] == "Trámite").sum())
        fila["Total - Res. Impug."] = int((sub_m["Estado_Grupo"] == "Res. Impug.").sum())
        fila["Total"] = len(sub_m)
        fila["%"] = round(len(sub_m) / tot_gen * 100, 1) if tot_gen > 0 else 0.0
        filas.append(fila)

    fila_tot = {"N°": "", "Mes de Creación": nombre_total, "_es_total": True}
    for a in anios_disp:
        sub_a = df_v[df_v["Anio_Col"] == a]
        fila_tot[f"{a} - Trámite"] = int((sub_a["Estado_Grupo"] == "Trámite").sum())
        fila_tot[f"{a} - Res. Impug."] = int((sub_a["Estado_Grupo"] == "Res. Impug.").sum())
    fila_tot["Total - Trámite"] = int((df_v["Estado_Grupo"] == "Trámite").sum())
    fila_tot["Total - Res. Impug."] = int((df_v["Estado_Grupo"] == "Res. Impug.").sum())
    fila_tot["Total"] = tot_gen
    fila_tot["%"] = 100.0 if tot_gen > 0 else 0.0
    filas.append(fila_tot)

    return pd.DataFrame(filas), anios_disp

def construir_df_dimension_anio(df_in, col_dim_src, label_dim, es_magistrado=False):
    if df_in.empty:
        return pd.DataFrame(), []
    df_v = df_in.copy()
    val_nulo = "<SIN MAGISTRADO>" if es_magistrado else "(No reconocida)"

    if col_dim_src in df_v.columns:
        df_v["Dim_Exacta"] = (
            df_v[col_dim_src].fillna(val_nulo).astype(str).str.strip()
            .replace({"": val_nulo, "-": val_nulo, "NONE": val_nulo, "NAN": val_nulo, "None": val_nulo, "nan": val_nulo})
        )
    else:
        df_v["Dim_Exacta"] = val_nulo

    if es_magistrado:
        mask_jef_sup = df_v["Es_Sede_Central"] & df_v["Dim_Exacta"].str.upper().str.contains("PALACIOS BRAN ROBERTO ALEJANDRO", na=False)
        df_v.loc[mask_jef_sup, "Dim_Exacta"] = "JEFATURA SUPREMA"

    df_v["Anio_Col"] = np.where(df_v["Anio_Num"] > 0, df_v["Anio_Num"].astype(str), "S/A")
    anios_disp = sorted([a for a in df_v["Anio_Col"].unique() if a != "S/A"])
    if "S/A" in df_v["Anio_Col"].unique():
        anios_disp.append("S/A")

    tot_gen = len(df_v)
    dims_ord = df_v.groupby("Dim_Exacta").size().sort_values(ascending=False).index.tolist()

    filas = []
    for idx_d, d_nom in enumerate(dims_ord, start=1):
        sub_d = df_v[df_v["Dim_Exacta"] == d_nom]
        fila_d = {"N°": str(idx_d), label_dim: d_nom, "_es_total": False}
        for a in anios_disp:
            sub_da = sub_d[sub_d["Anio_Col"] == a]
            fila_d[f"{a} - Trámite"] = int((sub_da["Estado_Grupo"] == "Trámite").sum())
            fila_d[f"{a} - Res. Impug."] = int((sub_da["Estado_Grupo"] == "Res. Impug.").sum())
        fila_d["Total - Trámite"] = int((sub_d["Estado_Grupo"] == "Trámite").sum())
        fila_d["Total - Res. Impug."] = int((sub_d["Estado_Grupo"] == "Res. Impug.").sum())
        fila_d["Total"] = len(sub_d)
        fila_d["%"] = round(len(sub_d) / tot_gen * 100, 1) if tot_gen > 0 else 0.0
        filas.append(fila_d)

    fila_tot = {"N°": "", label_dim: "Total carga procedimental", "_es_total": True}
    for a in anios_disp:
        sub_a = df_v[df_v["Anio_Col"] == a]
        fila_tot[f"{a} - Trámite"] = int((sub_a["Estado_Grupo"] == "Trámite").sum())
        fila_tot[f"{a} - Res. Impug."] = int((sub_a["Estado_Grupo"] == "Res. Impug.").sum())
    fila_tot["Total - Trámite"] = int((df_v["Estado_Grupo"] == "Trámite").sum())
    fila_tot["Total - Res. Impug."] = int((df_v["Estado_Grupo"] == "Res. Impug.").sum())
    fila_tot["Total"] = tot_gen
    fila_tot["%"] = 100.0 if tot_gen > 0 else 0.0
    filas.append(fila_tot)

    return pd.DataFrame(filas), anios_disp

def generar_cuadro_produccion_mes_actual(df_mes_act, col_grupo, label_col, nombre_total):
    if df_mes_act.empty:
        return pd.DataFrame()
    g_tot = df_mes_act.groupby(col_grupo).size()
    orden_grupos = g_tot.sort_values(ascending=False).index.tolist()

    res = pd.DataFrame(index=orden_grupos)
    res["Resoluciones"] = df_mes_act[~df_mes_act["Es_Informe"]].groupby(col_grupo).size().reindex(orden_grupos, fill_value=0)
    res["Informes"] = df_mes_act[df_mes_act["Es_Informe"]].groupby(col_grupo).size().reindex(orden_grupos, fill_value=0)
    res["Total"] = g_tot.reindex(orden_grupos, fill_value=0)

    tot_gen = int(res["Total"].sum())
    res["% Part."] = (res["Total"] / tot_gen * 100).round(2) if tot_gen > 0 else 0.0
    res = res.reset_index().rename(columns={"index": label_col})

    fila_tot = {
        label_col: nombre_total,
        "Resoluciones": int(res["Resoluciones"].sum()),
        "Informes": int(res["Informes"].sum()),
        "Total": tot_gen,
        "% Part.": 100.0 if tot_gen > 0 else 0.0
    }
    return pd.concat([res, pd.DataFrame([fila_tot])], ignore_index=True)

# Cuadros de Prescripción
etiquetas_rangos_color = [f"{emoji} {clave_rg}" for _, clave_rg, _, _, _, emoji in CONFIG_COLORES_RANGOS]
if es_nacional:
    p_anc_s = df_p_act[df_p_act["Es_Sede_Central"]].groupby("Rango_Presc").size().reindex(ORDEN_RANGOS_PRESC, fill_value=0)
    p_oda_s = df_p_act[~df_p_act["Es_Sede_Central"]].groupby("Rango_Presc").size().reindex(ORDEN_RANGOS_PRESC, fill_value=0)
    t_2_1_presc = pd.DataFrame({
        "Rango de Riesgo de Expedientes Prescritos y por Prescribir": etiquetas_rangos_color,
        "ANC Sede Central": p_anc_s.values,
        "% en ANC": (p_anc_s.values / p_anc_s.sum() * 100).round(2) if p_anc_s.sum() else 0,
        "ODANCs (34 Dist.)": p_oda_s.values,
        "% en ODANC": (p_oda_s.values / p_oda_s.sum() * 100).round(2) if p_oda_s.sum() else 0,
        "Total Nacional": p_anc_s.values + p_oda_s.values
    })
    tot_p_nac = t_2_1_presc["Total Nacional"].sum()
    t_2_1_presc["% Nac."] = (t_2_1_presc["Total Nacional"] / tot_p_nac * 100).round(2) if tot_p_nac else 0
    fila_tot_p = pd.DataFrame([{
        "Rango de Riesgo de Expedientes Prescritos y por Prescribir": "TOTAL PRESCRIPCIONES",
        "ANC Sede Central": int(p_anc_s.sum()), "% en ANC": 100.0 if p_anc_s.sum() else 0.0,
        "ODANCs (34 Dist.)": int(p_oda_s.sum()), "% en ODANC": 100.0 if p_oda_s.sum() else 0.0,
        "Total Nacional": int(tot_p_nac), "% Nac.": 100.0 if tot_p_nac else 0.0
    }])
    t_2_1_presc = pd.concat([t_2_1_presc, fila_tot_p], ignore_index=True)
else:
    p_sede_s = df_p_act.groupby("Rango_Presc").size().reindex(ORDEN_RANGOS_PRESC, fill_value=0)
    tot_p_s = p_sede_s.sum()
    t_2_1_presc = pd.DataFrame({
        "Rango de Riesgo de Expedientes Prescritos y por Prescribir": etiquetas_rangos_color,
        "Expedientes": p_sede_s.values,
        "% Part.": (p_sede_s.values / tot_p_s * 100).round(2) if tot_p_s else 0
    })
    fila_tot_p = pd.DataFrame([{
        "Rango de Riesgo de Expedientes Prescritos y por Prescribir": f"TOTAL PRESCRIPCIONES {dist_sel}",
        "Expedientes": int(tot_p_s), "% Part.": 100.0 if tot_p_s else 0.0
    }])
    t_2_1_presc = pd.concat([t_2_1_presc, fila_tot_p], ignore_index=True)

ct_et_rg = pd.crosstab(df_p_act["Etapa_Presc_Cat"], df_p_act["Rango_Presc"]).reindex(index=ORDEN_ETAPAS, columns=ORDEN_RANGOS_PRESC, fill_value=0)
t_2_2_et = pd.DataFrame(index=ORDEN_ETAPAS)
t_2_2_et["🔴 Prescritos (<=0d)"] = ct_et_rg["Prescritos (<= 0 días)"]
t_2_2_et["🟡 Crítico (1-15d)"] = ct_et_rg["Crítico: De 1 a 15 días"]
t_2_2_et["🟢 Alto (16-30d)"] = ct_et_rg["Alto: De 16 a 30 días"]
t_2_2_et["🔵 Moderado (31-60d)"] = ct_et_rg["Moderado: De 31 a 60 días"]
t_2_2_et["🔵 Preventivo (61-90d)"] = ct_et_rg["Preventivo: De 61 a 90 días"]
t_2_2_et["🟣 Monitoreo (91-240d)"] = ct_et_rg["Monitoreo: De 91 a 240 días"]
t_2_2_et["Total"] = t_2_2_et.sum(axis=1)
tot_et_p_gen = int(t_2_2_et["Total"].sum())
t_2_2_et["% Part."] = (t_2_2_et["Total"] / tot_et_p_gen * 100).round(2) if tot_et_p_gen > 0 else 0.0
t_2_2_et = t_2_2_et.reset_index().rename(columns={"index": "Etapa Procesal"})
fila_tot_22et = {
    "Etapa Procesal": "TOTAL POR ETAPA PROCESAL",
    **{c: int(t_2_2_et[c].sum()) for c in t_2_2_et.columns if c not in ("Etapa Procesal", "% Part.")},
    "% Part.": 100.0 if tot_et_p_gen > 0 else 0.0
}
t_2_2_et = pd.concat([t_2_2_et, pd.DataFrame([fila_tot_22et])], ignore_index=True)

# Cuadros de Producción
fechas_pr = df_pr_f["Fecha_Desc_Dt"].dropna()
mes_act_lbl = "Mes Actual"
t_3_1, t_3_2, t_3_3, t_3_4 = pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
per_act, per_ant = None, None
df_pr_mes_act = df_pr_f.copy()

if not fechas_pr.empty:
    per_act = fechas_pr.max().to_period("M")
    per_ant = per_act - 1
    mes_act_lbl = f"{MESES_NOM[per_act.month]} {per_act.year}"
    df_pr_f = df_pr_f.copy()
    df_pr_f["Periodo_M"] = df_pr_f["Fecha_Desc_Dt"].dt.to_period("M")
    df_pr_mes_act = df_pr_f[df_pr_f["Periodo_M"] == per_act].copy()

    filas_pr = []
    for per_obj, etiq in [(per_ant, f"Mes anterior ({MESES_NOM[per_ant.month]} {per_ant.year})"), (per_act, f"Mes actual ({mes_act_lbl})")]:
        sub_p = df_pr_f[df_pr_f["Periodo_M"] == per_obj]
        inf_p = int(sub_p["Es_Informe"].sum())
        res_p = len(sub_p) - inf_p
        if es_nacional:
            filas_pr.append({
                "Periodo Estadístico": etiq,
                "ANC Sede Central": int(sub_p["Es_Sede_Central"].sum()),
                "ODANCs (34 Dist.)": int((~sub_p["Es_Sede_Central"]).sum()),
                "Resoluciones": res_p, "Informes": inf_p, "Total Nacional": len(sub_p)
            })
        else:
            filas_pr.append({"Periodo Estadístico": etiq, "Resoluciones": res_p, "Informes": inf_p, "Total": len(sub_p)})

    t_3_1 = pd.DataFrame(filas_pr)
    col_tot_31 = "Total Nacional" if es_nacional else "Total"
    tot_3_1 = t_3_1[col_tot_31].sum()
    t_3_1["% Acum."] = (t_3_1[col_tot_31] / tot_3_1 * 100).round(2) if tot_3_1 else 0.0
    fila_tot_31 = {"Periodo Estadístico": "TOTAL PRODUCCIÓN ACUMULADA", **{c: int(t_3_1[c].sum()) for c in t_3_1.columns if c not in ("Periodo Estadístico", "% Acum.")}, "% Acum.": 100.0 if tot_3_1 else 0.0}
    t_3_1 = pd.concat([t_3_1, pd.DataFrame([fila_tot_31])], ignore_index=True)

if "DETALLE DEL TIPO DE PRONUNCIAMIENTO" in df_pr_mes_act.columns:
    df_pr_mes_act["Pron_Limpio"] = df_pr_mes_act["DETALLE DEL TIPO DE PRONUNCIAMIENTO"].fillna("NO ESPECIFICADO").astype(str).str.strip()
    if es_nacional:
        t_3_2_full = pd.DataFrame({
            "ANC Sede Central": df_pr_mes_act[df_pr_mes_act["Es_Sede_Central"]].groupby("Pron_Limpio").size(),
            "ODANCs (34 Dist.)": df_pr_mes_act[~df_pr_mes_act["Es_Sede_Central"]].groupby("Pron_Limpio").size()
        }).fillna(0).astype(int)
        t_3_2_full["Total Nacional"] = t_3_2_full["ANC Sede Central"] + t_3_2_full["ODANCs (34 Dist.)"]
        t_3_2_full = t_3_2_full.sort_values("Total Nacional", ascending=False).reset_index().rename(columns={"Pron_Limpio": "Tipo de Pronunciamiento / Sanción"})
        tot_32_nac = t_3_2_full["Total Nacional"].sum()

        if len(t_3_2_full) > 15:
            resto = t_3_2_full.iloc[15:]
            fila_otros = pd.DataFrame([{
                "Tipo de Pronunciamiento / Sanción": f"Otros pronunciamientos ({len(resto)} tipos)",
                "ANC Sede Central": int(resto["ANC Sede Central"].sum()),
                "ODANCs (34 Dist.)": int(resto["ODANCs (34 Dist.)"].sum()),
                "Total Nacional": int(resto["Total Nacional"].sum())
            }])
            t_3_2 = pd.concat([t_3_2_full.head(15), fila_otros], ignore_index=True)
        else:
            t_3_2 = t_3_2_full.copy()

        t_3_2["% Nac."] = (t_3_2["Total Nacional"] / tot_32_nac * 100).round(2) if tot_32_nac else 0.0
        fila_tot_32 = pd.DataFrame([{
            "Tipo de Pronunciamiento / Sanción": f"TOTAL {mes_act_lbl.upper()}",
            "ANC Sede Central": int(t_3_2["ANC Sede Central"].sum()),
            "ODANCs (34 Dist.)": int(t_3_2["ODANCs (34 Dist.)"].sum()),
            "Total Nacional": int(tot_32_nac), "% Nac.": 100.0 if tot_32_nac else 0.0
        }])
        t_3_2 = pd.concat([t_3_2, fila_tot_32], ignore_index=True)
    else:
        t_3_2_full = df_pr_mes_act["Pron_Limpio"].value_counts().reset_index()
        t_3_2_full.columns = ["Tipo de Pronunciamiento / Sanción", "Total"]
        tot_32_s = t_3_2_full["Total"].sum()
        if len(t_3_2_full) > 15:
            resto = t_3_2_full.iloc[15:]
            t_3_2 = pd.concat([t_3_2_full.head(15), pd.DataFrame([{"Tipo de Pronunciamiento / Sanción": f"Otros pronunciamientos ({len(resto)} tipos)", "Total": int(resto["Total"].sum())}])], ignore_index=True)
        else:
            t_3_2 = t_3_2_full.copy()
        t_3_2["% Part."] = (t_3_2["Total"] / tot_32_s * 100).round(2) if tot_32_s else 0.0
        t_3_2 = pd.concat([t_3_2, pd.DataFrame([{"Tipo de Pronunciamiento / Sanción": f"TOTAL {mes_act_lbl.upper()} ({dist_sel})", "Total": int(tot_32_s), "% Part.": 100.0 if tot_32_s else 0.0}])], ignore_index=True)

t_3_3 = generar_cuadro_produccion_mes_actual(df_pr_mes_act, "Unidad_Prod_Todos" if es_nacional else "Unidad_Prod", "Unidad Orgánica / Funcional", f"TOTAL {mes_act_lbl.upper()} ({'TODOS' if es_nacional else dist_sel})")
t_3_4 = generar_cuadro_produccion_mes_actual(df_pr_mes_act, "Magistrado_Prod", "Magistrado", f"TOTAL {mes_act_lbl.upper()} ({'TODOS' if es_nacional else dist_sel})")

# ==============================================================================
# 9. CONTROL DE NAVEGACIÓN Y PANTALLAS DE DETALLE POR KPI
# ==============================================================================
if "kpi_activo" not in st.session_state:
    st.session_state["kpi_activo"] = None

def alternar_kpi(nombre_kpi: str):
    st.session_state["kpi_activo"] = None if st.session_state["kpi_activo"] == nombre_kpi else nombre_kpi

def mostrar_pantalla_detalle(df_base: pd.DataFrame, titulo: str, clave: str):
    df_det = df_base.copy()
    col_t, col_close = st.columns([5, 1])
    with col_t:
        st.subheader(f"📋 {titulo} — {len(df_det):,} expedientes")
    with col_close:
        if st.button("⬅ Volver al Dashboard", key=f"volver_{clave}", use_container_width=True):
            st.session_state["kpi_activo"] = None
            st.rerun()

    f1, f2, f3 = st.columns([1, 1, 2])
    with f1:
        if "Etapa" in df_det.columns:
            sel_e = st.selectbox("Filtrar por Etapa:", ["Todas"] + sorted([x for x in df_det["Etapa"].unique() if x != "-"]), key=f"f_etapa_{clave}")
            if sel_e != "Todas":
                df_det = df_det[df_det["Etapa"] == sel_e]
    with f2:
        if "Año" in df_det.columns:
            sel_a = st.selectbox("Filtrar por Año:", ["Todos"] + sorted([x for x in df_det["Año"].unique() if x != "-"], reverse=True), key=f"f_anio_{clave}")
            if sel_a != "Todos":
                df_det = df_det[df_det["Año"] == sel_a]
    with f3:
        busq = st.text_input("Buscar por Número, Magistrado o Personal Asignado:", key=f"f_busq_{clave}")
        if busq:
            mask = df_det.astype(str).apply(lambda x: x.str.contains(busq, case=False, na=False, regex=False)).any(axis=1)
            df_det = df_det[mask]

    df_det = df_det.reset_index(drop=True)
    if "N°" in df_det.columns:
        df_det = df_det.drop(columns=["N°"])
    df_det.insert(0, "N°", range(1, len(df_det) + 1))

    cols_finales = [c for c in COLUMNAS_EXACTAS_DETALLE if c in df_det.columns]
    df_mostrar = df_det[cols_finales]

    st.caption(f"Mostrando **{len(df_mostrar):,}** expedientes encontrados.")
    st.dataframe(df_mostrar, use_container_width=True, hide_index=True, height=min(480, max(75, (len(df_mostrar) + 1) * 35 + 3)))

    b1, b2, _ = st.columns([1, 1, 3])
    nombre_arch = titulo.replace(" ", "_").replace("-", "").replace(".", "")
    with b1:
        st.download_button("📥 Exportar en Excel (.xlsx)", exportar_detalle_excel(df_mostrar), f"{nombre_arch}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"xlsx_{clave}", use_container_width=True)
    with b2:
        st.download_button("📄 Exportar en PDF", exportar_detalle_pdf(df_mostrar, titulo=titulo), f"{nombre_arch}.pdf", "application/pdf", key=f"pdf_{clave}", use_container_width=True)

def mostrar_pantalla_detalle_presc(df_base: pd.DataFrame, titulo: str, clave: str):
    df_det = df_base.copy().sort_values("Dias_Calc", ascending=True)
    col_t, col_close = st.columns([5, 1])
    with col_t:
        st.subheader(f"⏳ {titulo} — {len(df_det):,} expedientes")
    with col_close:
        if st.button("⬅ Volver al Dashboard", key=f"volver_{clave}", use_container_width=True):
            st.session_state["kpi_activo"] = None
            st.rerun()

    f1, f2, f3 = st.columns([1, 1.3, 1.7])
    with f1:
        etapas_disp_p = [e for e in ORDEN_ETAPAS if e in df_det["Etapa_Presc_Cat"].unique()]
        sel_et = st.selectbox("Filtrar por Etapa Procesal:", ["Todas"] + etapas_disp_p, key=f"f_etapa_p_{clave}")
        if sel_et != "Todas":
            df_det = df_det[df_det["Etapa_Presc_Cat"] == sel_et]
    with f2:
        opts_m = ["Todos"] + sorted([str(x) for x in df_det["MAGISTRADO"].dropna().unique() if str(x).strip() not in ("", "-", "None", "nan")])
        sel_m = st.selectbox("Filtrar por Magistrado:", opts_m, key=f"f_mag_p_{clave}")
        if sel_m != "Todos":
            df_det = df_det[df_det["MAGISTRADO"].astype(str) == sel_m]
    with f3:
        busq = st.text_input("Buscar por N° Expediente, Distrito o cualquier campo:", key=f"f_busq_p_{clave}")
        if busq:
            mask = df_det.astype(str).apply(lambda x: x.str.contains(busq, case=False, na=False, regex=False)).any(axis=1)
            df_det = df_det[mask]

    cols_internas = {"Distrito_Norm", "Es_Sede_Central", "Etapa_Presc_Cat", "Fecha_Presc_Dt", "Dias_Calc", "Rango_Presc", "Es_Atencion_Pref", "Atencion_Pref_Norm"}
    cols_orig_p = [c for c in df_det.columns if c not in cols_internas and c != "N°"]
    df_mostrar = df_det[cols_orig_p].copy()

    col_f_p = encontrar_columna(df_mostrar, ["FECHA DE PRESCRIPCION", "FECHA PRESCRIPCION"])
    if col_f_p:
        f_p_dt = pd.to_datetime(df_mostrar[col_f_p], errors="coerce", dayfirst=True)
        df_mostrar[col_f_p] = f_p_dt.dt.strftime("%d/%m/%Y").fillna(df_mostrar[col_f_p].astype(str))

    df_mostrar["DÍAS P/ PRESCRIBIR"] = df_det["Dias_Calc"].astype(int)
    mapa_iconos_rango = {clave_rg: f"{icono_rg} {clave_rg}".strip() for _, clave_rg, icono_rg, _, _, _ in CONFIG_COLORES_RANGOS}
    df_mostrar["RANGO DE RIESGO"] = df_det["Rango_Presc"].map(lambda x: mapa_iconos_rango.get(x, str(x)))

    df_mostrar = df_mostrar.reset_index(drop=True)
    df_mostrar.insert(0, "N°", range(1, len(df_mostrar) + 1))

    st.caption(f"Mostrando **{len(df_mostrar):,}** expedientes encontrados (ordenados del más próximo a prescribir).")
    st.dataframe(df_mostrar, use_container_width=True, hide_index=True, height=min(480, max(75, (len(df_mostrar) + 1) * 35 + 3)))

    b1, b2, _ = st.columns([1, 1, 3])
    nombre_arch = titulo.replace(" ", "_").replace("-", "").replace(".", "").replace("(", "").replace(")", "").replace("<=", "menor_igual_").replace(":", "")
    with b1:
        st.download_button("📥 Exportar en Excel (.xlsx)", exportar_detalle_excel(df_mostrar), f"{nombre_arch}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"xlsx_p_{clave}", use_container_width=True)
    with b2:
        st.download_button("📄 Exportar en PDF", exportar_detalle_pdf(df_mostrar, titulo=titulo), f"{nombre_arch}.pdf", "application/pdf", key=f"pdf_p_{clave}", use_container_width=True)

def mostrar_pantalla_detalle_prod(df_base: pd.DataFrame, titulo: str, clave: str):
    df_det = df_base.copy()
    if "Fecha_Desc_Dt" in df_det.columns:
        df_det = df_det.sort_values("Fecha_Desc_Dt", ascending=False)

    col_t, col_close = st.columns([5, 1])
    with col_t:
        st.subheader(f"⚖️ {titulo} — {len(df_det):,} registros")
    with col_close:
        if st.button("⬅ Volver al Dashboard", key=f"volver_{clave}", use_container_width=True):
            st.session_state["kpi_activo"] = None
            st.rerun()

    if "Fecha_Desc_Dt" in df_det.columns:
        df_det["_anio_desc"] = df_det["Fecha_Desc_Dt"].apply(lambda d: str(int(d.year)) if pd.notna(d) else "Sin año")
        df_det["_mes_num"] = df_det["Fecha_Desc_Dt"].apply(lambda d: int(d.month) if pd.notna(d) else 99)
        df_det["_mes_lbl"] = df_det["Fecha_Desc_Dt"].apply(lambda d: MESES_NOM[d.month] if pd.notna(d) else "Sin fecha")
    else:
        df_det["_anio_desc"], df_det["_mes_num"], df_det["_mes_lbl"] = "Sin año", 99, "Sin fecha"

    df_det["_tipo_doc"] = np.where(df_det["Es_Informe"], "Informe", "Resolución")

    f1, f2, f3, f4, f5 = st.columns([0.9, 1.1, 1.4, 1.1, 1.5])
    with f1:
        anios_ord = sorted([a for a in df_det["_anio_desc"].unique() if a != "Sin año"], reverse=True)
        if "Sin año" in df_det["_anio_desc"].unique():
            anios_ord.append("Sin año")
        opts_anio = ["Todos"] + anios_ord
        idx_anio_def = opts_anio.index("2026") if "2026" in opts_anio else (1 if len(opts_anio) > 1 else 0)
        sel_anio = st.selectbox("Año (Descarga):", opts_anio, index=idx_anio_def, key=f"f_anio_pr_{clave}")
        if sel_anio != "Todos":
            df_det = df_det[df_det["_anio_desc"] == sel_anio]
    with f2:
        meses_ord = df_det[["_mes_num", "_mes_lbl"]].drop_duplicates().sort_values("_mes_num")["_mes_lbl"].tolist()
        opts_mes = ["Todos"] + meses_ord
        mes_actual_def = MESES_NOM[int(df_det["Fecha_Desc_Dt"].dropna().max().month)] if "Fecha_Desc_Dt" in df_det.columns and df_det["Fecha_Desc_Dt"].notna().any() else None
        sel_mes = st.selectbox("Mes de Resol.:", opts_mes, index=opts_mes.index(mes_actual_def) if mes_actual_def in opts_mes else 0, key=f"f_mes_pr_{clave}")
        if sel_mes != "Todos":
            df_det = df_det[df_det["_mes_lbl"] == sel_mes]
    with f3:
        col_mag_pr = "MAGISTRADO" if "MAGISTRADO" in df_det.columns else "Magistrado_Prod"
        opts_mag = ["Todos"] + sorted([str(x).strip() for x in df_det[col_mag_pr].dropna().unique() if str(x).strip() not in ("", "-", "None", "nan", "NO ESPECIFICADO")])
        sel_mag = st.selectbox("Magistrado:", opts_mag, key=f"f_mag_pr_{clave}")
        if sel_mag != "Todos":
            df_det = df_det[df_det[col_mag_pr].astype(str).str.strip() == sel_mag]
    with f4:
        sel_tipo = st.selectbox("Tipo (Inf. / Res.):", ["Todos", "Resolución", "Informe"], key=f"f_tipo_pr_{clave}")
        if sel_tipo != "Todos":
            df_det = df_det[df_det["_tipo_doc"] == sel_tipo]
    with f5:
        busq = st.text_input("Buscar Expediente / Pron.:", key=f"f_busq_pr_{clave}")
        if busq:
            mask = df_det.astype(str).apply(lambda x: x.str.contains(busq, case=False, na=False, regex=False)).any(axis=1)
            df_det = df_det[mask]

    cols_internas_pr = {"Distrito_Norm", "Es_Sede_Central", "Fecha_Desc_Dt", "Es_Informe", "Unidad_Prod", "Unidad_Prod_Todos", "Magistrado_Prod", "Periodo_M", "Pron_Limpio", "_anio_desc", "_mes_num", "_mes_lbl", "_tipo_doc"}
    cols_orig_pr = [c for c in df_det.columns if c not in cols_internas_pr and c != "N°"]
    df_mostrar = df_det[cols_orig_pr].copy()
    if "FECHA_DESCARGA" in df_mostrar.columns:
        df_mostrar["FECHA_DESCARGA"] = pd.to_datetime(df_mostrar["FECHA_DESCARGA"], errors="coerce", dayfirst=True).dt.strftime("%d/%m/%Y").fillna("")

    for col in df_mostrar.columns:
        df_mostrar[col] = df_mostrar[col].fillna("").astype(str).replace({"None": "", "none": "", "NONE": "", "nan": "", "NaN": "", "NAN": "", "NaT": "", "NAT": "", "<NA>": ""}).str.strip()

    df_mostrar.insert(0, "Tipo Doc.", df_det["_tipo_doc"].values)
    df_mostrar = df_mostrar.reset_index(drop=True)
    df_mostrar.insert(0, "N°", range(1, len(df_mostrar) + 1))

    st.caption(f"Mostrando **{len(df_mostrar):,}** registros de producción encontrados.")
    st.dataframe(df_mostrar, use_container_width=True, hide_index=True, height=min(480, max(75, (len(df_mostrar) + 1) * 35 + 3)))

    b1, b2, _ = st.columns([1, 1, 3])
    nombre_arch = titulo.replace(" ", "_").replace("-", "").replace(".", "")
    with b1:
        st.download_button("📥 Exportar en Excel (.xlsx)", exportar_detalle_excel(df_mostrar), f"{nombre_arch}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"xlsx_pr_{clave}", use_container_width=True)
    with b2:
        st.download_button("📄 Exportar en PDF", exportar_detalle_pdf(df_mostrar, titulo=titulo), f"{nombre_arch}.pdf", "application/pdf", key=f"pdf_pr_{clave}", use_container_width=True)

def render_kpis_carga(df_vista, df_no_inv_vista):
    df_tram_v = df_vista[df_vista["Estado_Grupo"] == "Trámite"]
    df_imp_v = df_vista[df_vista["Estado_Grupo"] == "Res. Impug."]
    df_jef_v = df_vista[df_vista["Despacho_Norm"].str.contains("JEFATURA", na=False)]

    c_anc_tot = int(df_vista["Es_Sede_Central"].sum())
    c_oda_tot = int((~df_vista["Es_Sede_Central"]).sum())

    t1, t2, t3, t4 = st.columns(4)
    with t1:
        sub_tot = f"↑ ANC: {c_anc_tot:,} | ODANC: {c_oda_tot:,}" if es_nacional else f"↑ 100.0% de {dist_sel}"
        st.button(f"CARGA PROCESAL TOTAL\n**{len(df_vista):,} exp.**\n*{sub_tot}*", key="kpi_btn_total", on_click=alternar_kpi, args=("TOTAL",), use_container_width=True)
    with t2:
        pct_tr = (len(df_tram_v) / len(df_vista) * 100) if len(df_vista) > 0 else 0.0
        st.button(f"EN TRÁMITE\n**{len(df_tram_v):,} exp.**\n*↑ {pct_tr:.1f}% de la carga*", key="kpi_btn_tramite", on_click=alternar_kpi, args=("TRAMITE",), use_container_width=True)
    with t3:
        pct_im = (len(df_imp_v) / len(df_vista) * 100) if len(df_vista) > 0 else 0.0
        st.button(f"RESUELTO - IMPUG. TRÁNSITO\n**{len(df_imp_v):,} exp.**\n*↑ {pct_im:.1f}% de la carga*", key="kpi_btn_resuelto", on_click=alternar_kpi, args=("RESUELTO",), use_container_width=True)
    with t4:
        j_anc = int(df_jef_v["Es_Sede_Central"].sum())
        j_oda = int((~df_jef_v["Es_Sede_Central"]).sum())
        sub_jef = f"↑ ANC: {j_anc:,} | ODANC: {j_oda:,}" if es_nacional else f"↑ {len(df_jef_v)/max(1,len(df_vista))*100:.1f}% de la carga filtrada"
        st.button(f"EN JEFATURA\n**{len(df_jef_v):,} exp.**\n*{sub_jef}*", key="kpi_btn_jef_total", on_click=alternar_kpi, args=("JEF_TOTAL",), use_container_width=True)

    ni_anc = int(df_no_inv_vista["Es_Sede_Central"].sum())
    ni_oda = int((~df_no_inv_vista["Es_Sede_Central"]).sum())
    sub_ni = f"↑ ANC: {ni_anc:,} | ODANC: {ni_oda:,}" if es_nacional else f"↑ Sede: {dist_sel}"

    conteo_etapas = df_vista["Etapa_Cat"].value_counts()
    et_dom_nom = conteo_etapas.index[0] if not conteo_etapas.empty else "Sin datos"
    et_dom_cnt = int(conteo_etapas.iloc[0]) if not conteo_etapas.empty else 0
    et_dom_pct = (et_dom_cnt / max(1, len(df_vista)) * 100)

    df_ant_val = df_vista[df_vista["Antiguedad_Dias"].notna() & (df_vista["Antiguedad_Dias"] >= 0)]
    rezago_cnt = int((df_ant_val["Antiguedad_Dias"] > 365).sum())
    rezago_pct = (rezago_cnt / max(1, len(df_ant_val)) * 100)

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("ETAPA CON MAYOR CARGA", et_dom_nom, f"{et_dom_cnt:,} exp. ({et_dom_pct:.1f}% de la carga)")
    k2.metric("ANTIGÜEDAD > 1 AÑO (REZAGO)", f"{rezago_cnt:,} exp.", f"{rezago_pct:.1f}% supera los 365 días")
    with k3:
        st.button(f"SIN INVENTARIAR\n**{len(df_no_inv_vista):,} exp.**\n*{sub_ni}*", key="kpi_btn_no_inv", on_click=alternar_kpi, args=("NO_INVENTARIADO",), use_container_width=True)
    with k4:
        render_tarjeta_mas_antiguo(df_vista, es_nacional, "EXP. MÁS ANTIGUO")
    st.markdown("<div style='margin-bottom: 10px;'></div>", unsafe_allow_html=True)

def render_tabla_multinivel_anio(df_tab, anios_list, col_nombre, col_num_orden=None, max_height_css=""):
    if df_tab.empty:
        st.info("Sin registros disponibles para desagregar.")
        return

    th_anios = "".join([f"<th colspan='2' style='padding:6px; border:1px solid #334155; text-align:center; color:#f8fafc;'>{a}</th>" for a in anios_list])
    th_sub = "".join([
        "<th style='padding:5px; border:1px solid #334155; text-align:center; font-size:11.5px; color:#38bdf8; white-space:nowrap;'>Trámite</th>"
        "<th style='padding:5px; border:1px solid #334155; text-align:center; font-size:11.5px; color:#f59e0b; white-space:nowrap;'>Res. Impug.</th>"
        for _ in anios_list
    ]) + (
        "<th style='padding:5px; border:1px solid #334155; text-align:center; font-size:11.5px; color:#38bdf8; white-space:nowrap;'>Trámite</th>"
        "<th style='padding:5px; border:1px solid #334155; text-align:center; font-size:11.5px; color:#f59e0b; white-space:nowrap;'>Res. Impug.</th>"
    )

    sticky_head = "style='position:sticky; top:0; z-index:2;'" if max_height_css else ""
    cols_excluir = {col_nombre, "_es_sub", "_es_total", "%"}
    if col_num_orden:
        cols_excluir.add(col_num_orden)
    cols_val = [c for c in df_tab.columns if c not in cols_excluir]

    html = f"""
    <div style="overflow-x:auto; {max_height_css} margin-bottom:16px; border-radius:8px; border:1px solid #334155;">
    <table style="width:100%; border-collapse:collapse; font-size:12.5px; color:#f8fafc; background:#111827; text-align:center;">
    <thead {sticky_head}>
    <tr style="background:#1e293b; border-bottom:1px solid #334155;">
    {f"<th rowspan='2' style='padding:8px 6px; border:1px solid #334155; vertical-align:middle; text-align:center; width:38px; background:#1e293b;'>N.º</th>" if col_num_orden else ""}
    <th rowspan="2" style="padding:8px 10px; border:1px solid #334155; vertical-align:middle; text-align:left; background:#1e293b;">{col_nombre}</th>
    {th_anios}
    <th colspan="2" style="padding:6px; border:1px solid #334155; text-align:center; color:#f8fafc; background:#0f172a; white-space:nowrap;">Total por estado</th>
    <th rowspan="2" style="padding:8px 10px; border:1px solid #334155; vertical-align:middle; text-align:center; background:#0f172a; white-space:nowrap;">Total</th>
    <th rowspan="2" style="padding:8px 10px; border:1px solid #334155; vertical-align:middle; text-align:center; background:#0f172a; white-space:nowrap; min-width:62px;">%</th>
    </tr>
    <tr style="background:#1e293b; border-bottom:2px solid #475569;">{th_sub}</tr>
    </thead>
    <tbody>
    """

    for idx, r in df_tab.iterrows():
        es_s = bool(r.get("_es_sub", False))
        es_t = bool(r.get("_es_total", False))
        if es_t:
            bg, fw, txt_color = "#1e293b", "700", "#fbbf24"
        elif es_s:
            bg, fw, txt_color = "#0f172a", "400", "#cbd5e1"
        else:
            bg = "#111827" if idx % 2 == 0 else "#0f172a"
            fw, txt_color = "600", "#f8fafc"

        sticky_tot = "position:sticky; bottom:0; z-index:2;" if (es_t and max_height_css) else ""
        celdas_num = ""
        for c in cols_val:
            v = int(r[c])
            v_str = "–" if v == 0 else f"{v:,}"
            c_estilo = "color:#fbbf24;" if (es_t and c == "Total") else (f"color:{txt_color};" if v > 0 else "color:#64748b;")
            fw_c = "700" if (es_t or (not es_s and c == "Total")) else "400"
            celdas_num += f"<td style='padding:6px 8px; border:1px solid #1f2937; text-align:center; white-space:nowrap; font-weight:{fw_c}; {c_estilo}'>{v_str}</td>"

        pct_str = f"{r['%']:.1f}%" if not es_t else "100%"
        fw_pct = "700" if (es_t or not es_s) else "400"
        pad_nom = "padding-left:24px;" if es_s else "padding-left:10px;"
        td_num = f"<td style='padding:6px; border:1px solid #1f2937; text-align:center; color:#94a3b8; font-weight:600; white-space:nowrap;'>{r[col_num_orden]}</td>" if col_num_orden else ""

        html += (
            f"<tr style='background:{bg}; {sticky_tot}'>{td_num}"
            f"<td style='padding:6px 10px; {pad_nom} border:1px solid #1f2937; text-align:left; font-weight:{fw}; color:{txt_color};'>{r[col_nombre]}</td>"
            f"{celdas_num}<td style='padding:6px 10px; border:1px solid #1f2937; text-align:center; white-space:nowrap; font-weight:{fw_pct}; color:{txt_color};'>{pct_str}</td></tr>"
        )
    html += "</tbody></table></div>"
    st.markdown(html, unsafe_allow_html=True)

# ==============================================================================
# 10. RENDERIZADO PRINCIPAL (DETALLE O 3 PESTAÑAS EJECUTIVAS)
# ==============================================================================
kpi_sel = st.session_state["kpi_activo"]

if kpi_sel is not None:
    df_act_det = st.session_state.get("df_act_filtrado_actual", df_act_sede)
    df_no_inv_det = st.session_state.get("df_no_inv_filtrado_actual", df_no_inv)
    if kpi_sel in ("TOTAL", "TRAMITE", "RESUELTO", "NO_INVENTARIADO", "JEF_TOTAL"):
        mapa_kpis_carga = {
            "TOTAL": (f"Detalle de Carga Procesal Filtrada — {nombre_ambito}", df_act_det),
            "TRAMITE": (f"Detalle de Expedientes En Trámite — {nombre_ambito}", df_act_det[df_act_det["Estado_Grupo"] == "Trámite"]),
            "RESUELTO": (f"Detalle de Resuelto - Impug. Tránsito — {nombre_ambito}", df_act_det[df_act_det["Estado_Grupo"] == "Res. Impug."]),
            "NO_INVENTARIADO": (f"Detalle de Expedientes Sin Inventariar — {nombre_ambito}", df_no_inv_det),
            "JEF_TOTAL": (f"Detalle de Expedientes en Jefatura — {nombre_ambito}", df_act_det[df_act_det["Despacho_Norm"].str.contains("JEFATURA", na=False)]),
        }
        tit_kpi, df_kpi = mapa_kpis_carga[kpi_sel]
        mostrar_pantalla_detalle(df_kpi, tit_kpi, kpi_sel)
    elif kpi_sel == "PROD_ACUMULADA":
        mostrar_pantalla_detalle_prod(df_pr_f, f"Detalle de Producción Acumulada — {nombre_ambito}", kpi_sel)
    elif kpi_sel == "PRESC_TOTAL_VIG":
        mostrar_pantalla_detalle_presc(df_p_vigente, f"Detalle de Riesgo de Prescripción (1-240d) – {nombre_ambito}", kpi_sel)
    elif kpi_sel == "PRESC_ATENCION_PREF":
        mostrar_pantalla_detalle_presc(df_p_atencion_pref, f"Detalle de Riesgo de Prescripción con Atención Preferencial — {nombre_ambito}", kpi_sel)
    elif kpi_sel.startswith("PRESC_RANGO_"):
        mapa_kpis_presc = {
            cod_k: (f"Detalle — Rango de Riesgo: {clave_rg} — {nombre_ambito}", df_p_act[df_p_act["Rango_Presc"] == clave_rg])
            for cod_k, clave_rg, _, _, _, _ in CONFIG_COLORES_RANGOS
        }
        tit_p, df_p_sub = mapa_kpis_presc[kpi_sel]
        mostrar_pantalla_detalle_presc(df_p_sub, tit_p, kpi_sel)
else:
    tab_carga, tab_presc, tab_prod = st.tabs([
        "📊 1. Carga Procesal Integral (Etapas, Sub-etapas y Jefatura)",
        "⏳ 2. Riesgo de Prescripción",
        "⚖️ 3. Producción Mensual y Pronunciamientos"
    ])

    # --------------------------------------------------------------------------
    # PESTAÑA 1: CARGA PROCESAL INTEGRAL (CUADRO + GRÁFICO ESPECIALIZADO)
    # --------------------------------------------------------------------------
    with tab_carga:
        # Base combinada para que los selectboxes incluyan valores tanto de Inventariados como de Sin Inventariar
        df_base_opciones = pd.concat([df_act_sede, df_no_inv], ignore_index=True)

        # Barra de Filtros Cruzados Interactivos
        with st.expander("🎛️ Filtros Dinámicos de Carga (Etapa, Sub-etapa, Unidad / Jefatura y Año de Creación)", expanded=True):
            fc1, fc2, fc3, fc4 = st.columns([1.1, 1.3, 1.5, 1.1])
            with fc1:
                etapas_disp_f = ["Todas"] + [e for e in ORDEN_ETAPAS if e in df_base_opciones["Etapa_Cat"].unique()]
                f_etapa = st.selectbox("Etapa Procesal:", etapas_disp_f, key="filt_c_etapa")
            with fc2:
                df_tmp_sub = df_base_opciones if f_etapa == "Todas" else df_base_opciones[df_base_opciones["Etapa_Cat"] == f_etapa]
                subs_disp_f = ["Todas"] + sorted([s for s in df_tmp_sub["Subetapa_Clean"].unique() if s not in ("", "-", "SIN SUBETAPA")])
                f_subetapa = st.selectbox("Sub-etapa:", subs_disp_f, key="filt_c_subetapa")
            with fc3:
                unid_disp_f = ["Todas (Inc. Jefatura)"] + sorted([u for u in df_base_opciones["Despacho_Clean"].unique() if u not in ("", "-")])
                f_unidad = st.selectbox("Unidad Orgánica / Jefatura:", unid_disp_f, key="filt_c_unidad")
            with fc4:
                anios_orden_f = ["Todos", "2026", "2025", "2024", "2023", "2022", "2021 y ant."]
                anios_presentes_f = ["Todos"] + [a for a in anios_orden_f[1:] if a in df_base_opciones["Anio_Grupo"].unique()]
                f_anio = st.selectbox("Año de Creación:", anios_presentes_f, key="filt_c_anio")

        # Función única de filtrado aplicada simultáneamente a Carga Activa y Sin Inventariar
        def aplicar_filtros_dinamicos(df_target):
            res = df_target.copy()
            if f_etapa != "Todas":
                res = res[res["Etapa_Cat"] == f_etapa]
            if f_subetapa != "Todas":
                res = res[res["Subetapa_Clean"] == f_subetapa]
            if f_unidad != "Todas (Inc. Jefatura)":
                res = res[res["Despacho_Clean"] == f_unidad]
            if f_anio != "Todos":
                res = res[res["Anio_Grupo"] == f_anio]
            return res

        df_act = aplicar_filtros_dinamicos(df_act_sede)
        df_no_inv_filt = aplicar_filtros_dinamicos(df_no_inv)

        st.session_state["df_act_filtrado_actual"] = df_act
        st.session_state["df_no_inv_filtrado_actual"] = df_no_inv_filt
        df_cal = df_act[df_act["Etapa_Cat"] == "Calificación"].copy()

        # KPIs Ejecutivos Superiores (100% dinámicos ante cualquier filtro, incluido Sin Inventariar)
        render_kpis_carga(df_act, df_no_inv_filt)
        st.markdown("---")

        # ======================================================================
        # BLOQUE 1.1: CARGA POR AÑO DE CREACIÓN + GRÁFICO COMBINADO PARETO (80/20)
        # ======================================================================
        st.subheader(f"1.1. Carga Total por Año de Creación y Estado ({sub_titulo_sec})")
        orden_anios_cuadro = ["2021 y ant.", "2022", "2023", "2024", "2025", "2026"]
        tot_carga_cuadro = max(1, len(df_act))

        if es_nacional:
            filas_1_1 = []
            for a_lbl in orden_anios_cuadro:
                sub_a = df_act[df_act["Anio_Grupo"] == a_lbl]
                sub_anc = sub_a[sub_a["Es_Sede_Central"]]
                sub_oda = sub_a[~sub_a["Es_Sede_Central"]]
                tot_a = len(sub_a)
                filas_1_1.append({
                    "Año de Creación": a_lbl,
                    "ANC - En Trámite": int((sub_anc["Estado_Grupo"] == "Trámite").sum()),
                    "ANC - Res. Impug.": int((sub_anc["Estado_Grupo"] == "Res. Impug.").sum()),
                    "Total ANC": len(sub_anc),
                    "ODANC - En Trámite": int((sub_oda["Estado_Grupo"] == "Trámite").sum()),
                    "ODANC - Res. Impug.": int((sub_oda["Estado_Grupo"] == "Res. Impug.").sum()),
                    "Total ODANC": len(sub_oda),
                    "Total Nacional": tot_a,
                    "% Part.": round(tot_a / tot_carga_cuadro * 100, 2) if len(df_act) > 0 else 0.0
                })

            t_1_1_anio = pd.DataFrame(filas_1_1).sort_values(
                by=["Total Nacional", "Año de Creación"], ascending=[False, False]
            ).reset_index(drop=True)
            t_1_1_anio.insert(0, "N°", [str(i) for i in range(1, len(t_1_1_anio) + 1)])

            fila_tot_11 = {
                "N°": "",
                "Año de Creación": "TOTAL CARGA PROCESAL",
                "ANC - En Trámite": int(t_1_1_anio["ANC - En Trámite"].sum()),
                "ANC - Res. Impug.": int(t_1_1_anio["ANC - Res. Impug."].sum()),
                "Total ANC": int(t_1_1_anio["Total ANC"].sum()),
                "ODANC - En Trámite": int(t_1_1_anio["ODANC - En Trámite"].sum()),
                "ODANC - Res. Impug.": int(t_1_1_anio["ODANC - Res. Impug."].sum()),
                "Total ODANC": int(t_1_1_anio["Total ODANC"].sum()),
                "Total Nacional": len(df_act),
                "% Part.": 100.0 if len(df_act) > 0 else 0.0
            }
            t_1_1_anio = pd.concat([t_1_1_anio, pd.DataFrame([fila_tot_11])], ignore_index=True)
        else:
            filas_1_1 = []
            for a_lbl in orden_anios_cuadro:
                sub_a = df_act[df_act["Anio_Grupo"] == a_lbl]
                tr_a = int((sub_a["Estado_Grupo"] == "Trámite").sum())
                im_a = int((sub_a["Estado_Grupo"] == "Res. Impug.").sum())
                tot_a = len(sub_a)
                filas_1_1.append({
                    "Año de Creación": a_lbl,
                    "En Trámite": tr_a,
                    "Resuelto - Impug. Tránsito": im_a,
                    "Total Expedientes": tot_a,
                    "% Part.": round(tot_a / tot_carga_cuadro * 100, 2) if len(df_act) > 0 else 0.0
                })

            t_1_1_anio = pd.DataFrame(filas_1_1).sort_values(
                by=["Total Expedientes", "Año de Creación"], ascending=[False, False]
            ).reset_index(drop=True)
            t_1_1_anio.insert(0, "N°", [str(i) for i in range(1, len(t_1_1_anio) + 1)])

            fila_tot_11 = {
                "N°": "",
                "Año de Creación": f"TOTAL CARGA {dist_sel}",
                "En Trámite": int(t_1_1_anio["En Trámite"].sum()),
                "Resuelto - Impug. Tránsito": int(t_1_1_anio["Resuelto - Impug. Tránsito"].sum()),
                "Total Expedientes": len(df_act),
                "% Part.": 100.0 if len(df_act) > 0 else 0.0
            }
            t_1_1_anio = pd.concat([t_1_1_anio, pd.DataFrame([fila_tot_11])], ignore_index=True)

        mostrar_tabla_centrada(t_1_1_anio)

        # Gráfico 1.1: Pareto Ejecutivo (Barras de Volumen + Curva de Concentración Acumulada %)
        df_par_anio = []
        for a_lbl in orden_anios_cuadro:
            sub_a = df_act[df_act["Anio_Grupo"] == a_lbl]
            df_par_anio.append({
                "Anio_Etq": f"Año {a_lbl}" if a_lbl != "2021 y ant." else "2021 y ant.",
                "Anio_Raw": a_lbl,
                "Tramite": int((sub_a["Estado_Grupo"] == "Trámite").sum()),
                "Res_Impug": int((sub_a["Estado_Grupo"] == "Res. Impug.").sum()),
                "Total": len(sub_a)
            })
        df_par_a = pd.DataFrame(df_par_anio).sort_values(by=["Total", "Anio_Raw"], ascending=[False, False]).reset_index(drop=True)
        tot_par_sum = max(1, int(df_par_a["Total"].sum()))
        df_par_a["Pct"] = (df_par_a["Total"] / tot_par_sum * 100).round(2)
        df_par_a["Acum_Pct"] = (df_par_a["Total"].cumsum() / tot_par_sum * 100).round(1)
        max_par_y = max(1, int(df_par_a["Total"].max()))

        custom_par_a = np.stack([
            df_par_a["Tramite"], df_par_a["Res_Impug"], df_par_a["Pct"], df_par_a["Acum_Pct"]
        ], axis=-1)

        fig_1_1 = go.Figure()
        fig_1_1.add_trace(go.Bar(
            x=df_par_a["Anio_Etq"],
            y=df_par_a["Total"],
            name="Total Expedientes",
            marker=dict(color="#1d4ed8", line=dict(color="#60a5fa", width=1.2)),
            text=[f"<b>{v:,}</b> ({p:.1f}%)" for v, p in zip(df_par_a["Total"], df_par_a["Pct"])],
            textposition="outside",
            textfont=dict(color="#f8fafc", size=11.5),
            customdata=custom_par_a,
            hovertemplate=(
                "<b>📅 Periodo:</b> %{x}<br>"
                "<b>📊 Total de Expedientes:</b> %{y:,} exp. (%{customdata[2]:.2f}% del total)<br>"
                "──────────────────────────<br>"
                "<b>🔵 En Trámite:</b> %{customdata[0]:,} exp.<br>"
                "<b>🟠 Resuelto - Impug. Tránsito:</b> %{customdata[1]:,} exp.<br>"
                "<b>📈 % Acumulado Pareto:</b> %{customdata[3]:.1f}%"
                "<extra></extra>"
            )
        ))
        fig_1_1.add_trace(go.Scatter(
            x=df_par_a["Anio_Etq"],
            y=df_par_a["Acum_Pct"],
            name="% Acumulado (Pareto)",
            yaxis="y2",
            mode="lines+markers+text",
            text=[f"{p:.1f}%" for p in df_par_a["Acum_Pct"]],
            textposition="top center",
            textfont=dict(color="#fbbf24", size=10.5),
            line=dict(color="#f59e0b", width=3),
            marker=dict(size=8, color="#f59e0b", line=dict(width=1.5, color="#111827")),
            hovertemplate="<b>📅 Periodo:</b> %{x}<br><b>📈 Concentración Acumulada:</b> %{y:.1f}%<extra></extra>"
        ))
        fig_1_1.add_hline(
            y=80, yref="y2", line_dash="dash", line_color="#ef4444", line_width=1.5,
            annotation_text="Umbral 80% Pareto", annotation_position="bottom right",
            annotation_font=dict(color="#fca5a5", size=11)
        )
        fig_1_1.update_layout(
            title=dict(text="Gráfico 1.1 — Análisis Pareto (80/20): Concentración de Carga por Año de Creación", font=dict(size=13.5, color="#f8fafc"), x=0.01, y=0.95),
            bargap=0.38,
            template="plotly_dark",
            paper_bgcolor="#111827",
            plot_bgcolor="#0f172a",
            xaxis=dict(type="category", categoryorder="array", categoryarray=df_par_a["Anio_Etq"].tolist(), title="", tickfont=dict(size=12, color="#f8fafc")),
            yaxis=dict(title="Total de Expedientes", range=[0, max_par_y * 1.24], gridcolor="rgba(148, 163, 184, 0.14)"),
            yaxis2=dict(title="% Acumulado", overlaying="y", side="right", range=[0, 112], showgrid=False, tickfont=dict(color="#fbbf24")),
            legend=dict(orientation="h", y=-0.18, x=0.5, xanchor="center", font=dict(size=11.5, color="#f8fafc")),
            height=360,
            margin=dict(l=15, r=25, t=50, b=60),
            hoverlabel=dict(bgcolor="#0f172a", bordercolor="#38bdf8", font_size=12, font_color="#f8fafc")
        )
        st.plotly_chart(fig_1_1, use_container_width=True)
        st.markdown("---")

        # ======================================================================
        # BLOQUE 1.2: CARGA POR ETAPA PROCESAL + ANILLO (DONUT) Y MAPA DE CALOR
        # ======================================================================
        t_1_2_etapas, anios_12e = construir_df_etapas_anio(df_act)
        st.subheader(f"1.2. Carga Procesal por Etapa y Sub-etapas de Calificación ({sub_titulo_sec})")
        st.caption("Orden procesal: Calificación (con desglose de sub-etapas) → Instrucción → Sanción → Apelación → Otros.")
        render_tabla_multinivel_anio(t_1_2_etapas, anios_12e, "Etapa / Sub-etapa Procesal", col_num_orden="N°")

        df_cal_otros = df_cal[df_cal["Subetapa_Cat"] == "   • Pend. Informe / Por Elevar / En Recepción / Otros"]
        if es_nacional:
            det_anc = df_cal_otros[df_cal_otros["Es_Sede_Central"]]["Subetapa_Raw"].value_counts()
            det_oda = df_cal_otros[~df_cal_otros["Es_Sede_Central"]]["Subetapa_Raw"].value_counts()
            if not det_anc.empty:
                st.caption(f"*Detalle Otros en Calificación Sede Central ({det_anc.sum()}): " + ", ".join([f"{k.title()} ({v})" for k, v in det_anc.items()]))
            if not det_oda.empty:
                st.caption(f"*Detalle Otros en Calificación ODANC ({det_oda.sum()}): " + ", ".join([f"{k.title()} ({v})" for k, v in det_oda.items()]))
        else:
            det_s = df_cal_otros["Subetapa_Raw"].value_counts()
            if not det_s.empty:
                st.caption(f"*Detalle Otros en Calificación {dist_sel} ({det_s.sum()}): " + ", ".join([f"{k.title()} ({v})" for k, v in det_s.items()]))

        g1_c, g2_c = st.columns([1.0, 1.25])

        # Gráfico 1.2.A: Gráfico de Anillo (Donut Chart) con etiquetas corregidas
        with g1_c:
            etapas_ord_cortas = [ETAPAS_CORTAS[e] for e in ORDEN_ETAPAS]
            df_donut = []
            for et_full in ORDEN_ETAPAS:
                sub_e = df_act[df_act["Etapa_Cat"] == et_full]
                df_donut.append({
                    "Etapa_Full": et_full,
                    "Etapa_Corta": ETAPAS_CORTAS[et_full],
                    "Tramite": int((sub_e["Estado_Grupo"] == "Trámite").sum()),
                    "Res_Impug": int((sub_e["Estado_Grupo"] == "Res. Impug.").sum()),
                    "Total": len(sub_e)
                })
            df_d = pd.DataFrame(df_donut)
            tot_d_sum = int(df_d["Total"].sum())
            df_d["Pct"] = (df_d["Total"] / max(1, tot_d_sum) * 100).round(1)

            colores_donut = ["#0284c7", "#6366f1", "#f59e0b", "#ef4444", "#10b981"]
            custom_d = np.stack([df_d["Etapa_Full"], df_d["Tramite"], df_d["Res_Impug"], df_d["Pct"]], axis=-1)

            fig_donut = go.Figure(data=[go.Pie(
                labels=df_d["Etapa_Corta"],
                values=df_d["Total"],
                hole=0.58,
                sort=False,
                direction="clockwise",
                marker=dict(colors=colores_donut, line=dict(color="#111827", width=2)),
                textinfo="percent",
                textposition="inside",
                insidetextorientation="horizontal",
                textfont=dict(size=12, color="#ffffff"),
                customdata=custom_d,
                hovertemplate=(
                    "<b>📂 Etapa:</b> %{customdata[0][0]}<br>"
                    "<b>📊 Total de Expedientes:</b> %{value:,} exp. (%{customdata[0][3]:.1f}% del total)<br>"
                    "──────────────────────────<br>"
                    "<b>🔵 En Trámite:</b> %{customdata[0][1]:,} exp.<br>"
                    "<b>🟠 Resuelto - Impug. Tránsito:</b> %{customdata[0][2]:,} exp."
                    "<extra></extra>"
                )
            )])
            fig_donut.add_annotation(
                text=f"<span style='font-size:11px; color:#94a3b8;'>TOTAL CARGA</span><br><b style='font-size:20px; color:#f8fafc;'>{tot_d_sum:,}</b><br><span style='font-size:11px; color:#38bdf8;'>expedientes</span>",
                x=0.5, y=0.5, font=dict(size=13), showarrow=False
            )
            fig_donut.update_layout(
                title=dict(text="Gráfico 1.2.A — Composición Porcentual por Etapa Procesal", font=dict(size=13.5, color="#f8fafc"), x=0.02, y=0.95),
                template="plotly_dark",
                paper_bgcolor="#111827",
                legend=dict(orientation="h", y=-0.12, x=0.5, xanchor="center", font=dict(size=11.5, color="#f8fafc")),
                height=390,
                margin=dict(l=15, r=15, t=50, b=55),
                hoverlabel=dict(bgcolor="#0f172a", bordercolor="#38bdf8", font_size=12, font_color="#f8fafc")
            )
            st.plotly_chart(fig_donut, use_container_width=True)

        # Gráfico 1.2.B: Matriz de Antigüedad por Etapa (Heatmap)
        # Gráfico 1.2.B: Matriz Semáforo de Criticidad Jurídica por Etapa y Antigüedad
        # Gráfico 1.2.B: Matriz de Criticidad Jurídica con Gradiente Moderno y Barra de Medición
        with g2_c:
            df_c_valid = df_act[df_act["Antiguedad_Dias"].notna() & (df_act["Antiguedad_Dias"] >= 0)].copy()
            if not df_c_valid.empty:
                df_c_valid["Etapa_Corta"] = df_c_valid["Etapa_Cat"].map(ETAPAS_CORTAS)
                etapas_ord_cortas = [ETAPAS_CORTAS[e] for e in ORDEN_ETAPAS]
                tramos_cols = ["0 a 2 años", "> 2 a 4 años", "> 4 años"]

                df_c_valid["Tramo_Antig"] = np.select(
                    [df_c_valid["Antiguedad_Dias"] <= 730, df_c_valid["Antiguedad_Dias"] <= 1461],
                    ["0 a 2 años", "> 2 a 4 años"],
                    default="> 4 años"
                )

                ct_hm = (
                    pd.crosstab(df_c_valid["Etapa_Corta"], df_c_valid["Tramo_Antig"])
                    .reindex(index=etapas_ord_cortas, columns=tramos_cols, fill_value=0)
                )
                tot_hm_gen = max(1, int(ct_hm.values.sum()))

                # Máximos por columna para dar gradiente interno dentro de cada nivel jurídico
                max_col_tramo = [max(1, int(ct_hm[col].max())) for col in tramos_cols]

                # Escala continua por tramos (0.0=Sin carga | 0.25-0.48=Turquesa | 0.52-0.73=Ámbar | 0.77-1.0=Carmesí/Coral)
                z_moderno = np.zeros((len(etapas_ord_cortas), len(tramos_cols)), dtype=float)
                texto_celdas = []
                custom_hm = np.empty((len(etapas_ord_cortas), len(tramos_cols), 4), dtype=object)

                for i_et, et_c in enumerate(etapas_ord_cortas):
                    fila_txt = []
                    sub_et_df = df_c_valid[df_c_valid["Etapa_Corta"] == et_c]
                    for j_tr, tr_c in enumerate(tramos_cols):
                        val_celda = int(ct_hm.loc[et_c, tr_c])
                        pct_celda = round(val_celda / tot_hm_gen * 100, 1)
                        ratio_col = val_celda / max_col_tramo[j_tr]

                        if val_celda == 0:
                            z_moderno[i_et, j_tr] = 0.05
                        elif j_tr == 0:
                            # 0 a 2 años: rango [0.27, 0.48] -> Azul Cian a Turquesa Vibrante
                            z_moderno[i_et, j_tr] = 0.27 + (0.21 * ratio_col)
                        elif j_tr == 1:
                            # > 2 a 4 años: rango [0.53, 0.73] -> Ámbar a Naranja Intenso
                            z_moderno[i_et, j_tr] = 0.53 + (0.20 * ratio_col)
                        else:
                            # > 4 años: rango [0.78, 1.00] -> Carmesí a Rojo Coral de Impacto
                            z_moderno[i_et, j_tr] = 0.78 + (0.22 * ratio_col)

                        fila_txt.append(f"<b>{val_celda:,}</b><br><span style='font-size:11px;'>({pct_celda:.1f}%)</span>")

                        sub_celda = sub_et_df[sub_et_df["Tramo_Antig"] == tr_c]
                        tr_cnt = int((sub_celda["Estado_Grupo"] == "Trámite").sum())
                        im_cnt = int((sub_celda["Estado_Grupo"] == "Res. Impug.").sum())
                        custom_hm[i_et, j_tr] = [val_celda, pct_celda, tr_cnt, im_cnt]
                    texto_celdas.append(fila_txt)

                # Paleta moderna con transiciones suaves y barra de medición lateral
                escala_moderna = [
                    [0.00, "#0f172a"], [0.22, "#131c31"],   # Sin carga
                    [0.25, "#075985"], [0.48, "#06b6d4"],   # Plazo Regular (Azul -> Turquesa)
                    [0.51, "#92400e"], [0.73, "#f97316"],   # Prioritario (Ámbar -> Naranja)
                    [0.76, "#881337"], [1.00, "#f43f5e"]    # Atención Inmediata (Carmesí -> Coral Neón)
                ]

                fig_hm_ant = go.Figure(data=go.Heatmap(
                    z=z_moderno,
                    zmin=0.0,
                    zmax=1.0,
                    x=tramos_cols,
                    y=etapas_ord_cortas,
                    text=texto_celdas,
                    texttemplate="%{text}",
                    textfont=dict(size=12.5, color="#ffffff"),
                    colorscale=escala_moderna,
                    xgap=3,
                    ygap=3,
                    showscale=True,
                    colorbar=dict(
                        title=dict(text="Prioridad", font=dict(size=11.5, color="#cbd5e1")),
                        tickmode="array",
                        tickvals=[0.10, 0.38, 0.63, 0.89],
                        ticktext=["0 Exp.", "Regular", "Alerta", "Crítico"],
                        tickfont=dict(size=10.5, color="#e2e8f0"),
                        thickness=14,
                        len=0.92
                    ),
                    customdata=custom_hm,
                    hovertemplate=(
                        "<b>📂 Etapa:</b> %{y} | <b>⏳ Antigüedad:</b> %{x}<br>"
                        "<b>📊 Total de Expedientes:</b> %{customdata[0]:,} exp. (%{customdata[1]:.1f}% del total)<br>"
                        "──────────────────────────<br>"
                        "<b>🔵 En Trámite:</b> %{customdata[2]:,} exp.<br>"
                        "<b>🟠 Resuelto - Impug. Tránsito:</b> %{customdata[3]:,} exp."
                        "<extra></extra>"
                    )
                ))
                fig_hm_ant.update_layout(
                    title=dict(text="Gráfico 1.2.B — Matriz de Antigüedad y Prioridad Jurídica por Etapa", font=dict(size=13.5, color="#f8fafc"), x=0.02, y=0.95),
                    template="plotly_dark",
                    paper_bgcolor="#111827",
                    plot_bgcolor="#111827",
                    xaxis=dict(title="Rango de Antigüedad", side="bottom", tickfont=dict(size=12, color="#f8fafc")),
                    yaxis=dict(categoryorder="array", categoryarray=list(reversed(etapas_ord_cortas)), title="", tickfont=dict(size=12, color="#f8fafc")),
                    height=355,
                    margin=dict(l=15, r=15, t=48, b=45),
                    hoverlabel=dict(bgcolor="#0f172a", bordercolor="#f43f5e", font_size=12, font_color="#f8fafc")
                )
                st.plotly_chart(fig_hm_ant, use_container_width=True)

                # Nota Ejecutiva Jurídica debajo del gráfico con totales y % consolidados
                tot_0_2 = int(ct_hm["0 a 2 años"].sum())
                tot_2_4 = int(ct_hm["> 2 a 4 años"].sum())
                tot_4_m = int(ct_hm["> 4 años"].sum())

                pct_0_2 = tot_0_2 / tot_hm_gen * 100
                pct_2_4 = tot_2_4 / tot_hm_gen * 100
                pct_4_m = tot_4_m / tot_hm_gen * 100

                etapa_critica_4m = ct_hm["> 4 años"].idxmax() if tot_4_m > 0 else "Ninguna"
                val_critica_4m = int(ct_hm["> 4 años"].max()) if tot_4_m > 0 else 0

                st.markdown(
                    f"""
                    <div style="background-color:#0f172a; border-left:4px solid #f43f5e; padding:10px 14px; border-radius:6px; font-size:12.5px; color:#e2e8f0; margin-top:-8px;">
                        <b>⚖️ Lectura Jurídica de Priorización (sobre {tot_hm_gen:,} exp.):</b><br>
                        • <span style="color:#22d3ee;"><b>Plazo Ordinario (0 a 2 años):</b></span> <b>{tot_0_2:,} exp. ({pct_0_2:.1f}%)</b> — Tramitación regular.<br>
                        • <span style="color:#fb923c;"><b>Seguimiento Prioritario (&gt; 2 a 4 años):</b></span> <b>{tot_2_4:,} exp. ({pct_2_4:.1f}%)</b> — Requiere control de impulsos procesales.<br>
                        • <span style="color:#fb7185;"><b>Atención Inmediata (&gt; 4 años):</b></span> <b>{tot_4_m:,} exp. ({pct_4_m:.1f}%)</b> — Máxima prioridad por riesgo de prescripción/caducidad (mayor foco en <b>{etapa_critica_4m}: {val_critica_4m:,} exp.</b>).
                    </div>
                    """,
                    unsafe_allow_html=True
                )
        st.markdown("---")

        # ======================================================================
        # BLOQUE 1.3: CALIFICACIÓN POR MES + GRÁFICO DE TENDENCIA EN LÍNEAS/ÁREA
        # ======================================================================
        t_1_3_cal_mes, anios_13c = construir_df_mes_anio_estado(df_cal, "TOTAL CALIFICACIÓN")
        st.subheader(f"1.3. Carga en Etapa de Calificación por Mes de Creación, Año y Estado ({sub_titulo_sec})")
        render_tabla_multinivel_anio(t_1_3_cal_mes, anios_13c, "Mes de Creación", col_num_orden="N°")

        if not df_cal.empty:
            df_cal_g = df_cal.copy()
            df_cal_g["Mes_Num"] = df_cal_g["Fecha_Creacion_Dt"].dt.month.fillna(0).astype(int)
            meses_activos_cal = [m for m in range(1, 13) if (df_cal_g["Mes_Num"] == m).any()]
            if meses_activos_cal:
                df_cal_m = (
                    df_cal_g[df_cal_g["Mes_Num"].isin(meses_activos_cal)]
                    .groupby(["Mes_Num", "Estado_Grupo"]).size().unstack(fill_value=0)
                    .reindex(meses_activos_cal, fill_value=0).reset_index()
                )
                for col_st in ["Trámite", "Res. Impug."]:
                    if col_st not in df_cal_m.columns:
                        df_cal_m[col_st] = 0
                df_cal_m["Mes_Nom"] = df_cal_m["Mes_Num"].map(MESES_ABR)
                df_cal_m["Total_Mes"] = df_cal_m["Trámite"] + df_cal_m["Res. Impug."]
                df_cal_m["Pct_Mes"] = (df_cal_m["Total_Mes"] / max(1, int(df_cal_m["Total_Mes"].sum())) * 100).round(1)
                max_cal_m = max(1, int(df_cal_m[["Trámite", "Res. Impug."]].max().max()))

                fig_1_3 = go.Figure()
                fig_1_3.add_trace(go.Scatter(
                    x=df_cal_m["Mes_Nom"], y=df_cal_m["Trámite"],
                    name="En Trámite", mode="lines+markers+text",
                    fill="tozeroy", fillcolor="rgba(16, 185, 129, 0.16)",
                    line=dict(color="#10b981", width=3),
                    marker=dict(size=8, color="#10b981"),
                    text=[f"{v:,}" for v in df_cal_m["Trámite"]],
                    textposition="top center", textfont=dict(color="#34d399", size=11),
                    customdata=np.stack([df_cal_m["Total_Mes"], df_cal_m["Pct_Mes"]], axis=-1),
                    hovertemplate="<b>🗓️ Mes:</b> %{x}<br><b>🟢 En Trámite:</b> %{y:,} exp.<br><b>📊 Total de Expedientes:</b> %{customdata[0]:,} exp. (%{customdata[1]:.1f}%)<extra></extra>"
                ))
                fig_1_3.add_trace(go.Scatter(
                    x=df_cal_m["Mes_Nom"], y=df_cal_m["Res. Impug."],
                    name="Resuelto - Impug. Tránsito", mode="lines+markers+text",
                    fill="tozeroy", fillcolor="rgba(244, 63, 94, 0.14)",
                    line=dict(color="#f43f5e", width=3, dash="dot"),
                    marker=dict(size=8, color="#f43f5e"),
                    text=[f"{v:,}" for v in df_cal_m["Res. Impug."]],
                    textposition="bottom center", textfont=dict(color="#fda4af", size=11),
                    customdata=np.stack([df_cal_m["Total_Mes"], df_cal_m["Pct_Mes"]], axis=-1),
                    hovertemplate="<b>🗓️ Mes:</b> %{x}<br><b>🔴 Res. Impug. Tránsito:</b> %{y:,} exp.<br><b>📊 Total de Expedientes:</b> %{customdata[0]:,} exp. (%{customdata[1]:.1f}%)<extra></extra>"
                ))
                fig_1_3.update_layout(
                    title=dict(text="Gráfico 1.3 — Tendencia Mensual en Etapa de Calificación (En Trámite vs. Res. Impug.)", font=dict(size=13.5, color="#f8fafc"), x=0.01, y=0.95),
                    template="plotly_dark", paper_bgcolor="#111827", plot_bgcolor="#0f172a",
                    xaxis=dict(type="category", title="Mes de Creación", tickfont=dict(size=11.5, color="#f8fafc")),
                    yaxis=dict(title="Expedientes", range=[0, max_cal_m * 1.24], gridcolor="rgba(148, 163, 184, 0.14)"),
                    legend=dict(orientation="h", y=-0.22, x=0.5, xanchor="center", font=dict(size=11.5, color="#f8fafc")),
                    height=330, margin=dict(l=15, r=15, t=45, b=65),
                    hoverlabel=dict(bgcolor="#0f172a", bordercolor="#10b981", font_size=12, font_color="#f8fafc")
                )
                st.plotly_chart(fig_1_3, use_container_width=True)

        # ======================================================================
        # BLOQUE 1.4: UNIDAD ORGÁNICA Y MAGISTRADO (BARRAS APILADAS CON NUEVA PALETA)
        # ======================================================================
        t_1_2_unidad, anios_12u = construir_df_dimension_anio(df_act, "Despacho o Unidad", "Unidad Orgánica", es_magistrado=False)
        t_1_2_magistrado, anios_12m = construir_df_dimension_anio(df_act, "Magistrado", "Magistrado de Control", es_magistrado=True)

        with st.expander("🔍 Ver Desglose Operativo por Unidad Orgánica y Magistrado de Control (Cuadros + Gráficos)", expanded=False):
            st.markdown(f"**1.4.A. Carga Procesal por Unidad Orgánica ({sub_titulo_sec})**")
            render_tabla_multinivel_anio(t_1_2_unidad, anios_12u, "Unidad Orgánica", col_num_orden="N°")

            # Gráfico 1.4.A: Barras Apiladas Horizontales (Paleta Esmeralda / Coral)
            df_u_graf = t_1_2_unidad[~t_1_2_unidad["_es_total"]].head(10).iloc[::-1]
            if not df_u_graf.empty:
                max_u = max(1, int(df_u_graf["Total"].max()))
                fig_1_4_u = go.Figure()
                fig_1_4_u.add_trace(go.Bar(
                    y=df_u_graf["Unidad Orgánica"], x=df_u_graf["Total - Trámite"], name="En Trámite", orientation="h",
                    marker=dict(color="#0d9488", line=dict(color="#2dd4bf", width=1)),
                    text=[f"<b>{v:,}</b>" if v >= max_u * 0.07 else "" for v in df_u_graf["Total - Trámite"]],
                    textposition="inside", insidetextanchor="middle", textfont=dict(color="#ffffff", size=11),
                    customdata=np.stack([df_u_graf["Total"], df_u_graf["%"]], axis=-1),
                    hovertemplate="<b>🏛️ Unidad:</b> %{y}<br><b>🟢 En Trámite:</b> %{x:,} exp.<br><b>📊 Total de Expedientes:</b> %{customdata[0]:,} exp. (%{customdata[1]:.1f}%)<extra></extra>"
                ))
                fig_1_4_u.add_trace(go.Bar(
                    y=df_u_graf["Unidad Orgánica"], x=df_u_graf["Total - Res. Impug."], name="Res. Impug. Tránsito", orientation="h",
                    marker=dict(color="#e11d48", line=dict(color="#fb7185", width=1)),
                    text=[f"<b>{v:,}</b>" if v >= max_u * 0.07 else "" for v in df_u_graf["Total - Res. Impug."]],
                    textposition="inside", insidetextanchor="middle", textfont=dict(color="#ffffff", size=11),
                    customdata=np.stack([df_u_graf["Total"], df_u_graf["%"]], axis=-1),
                    hovertemplate="<b>🏛️ Unidad:</b> %{y}<br><b>🔴 Res. Impug. Tránsito:</b> %{x:,} exp.<br><b>📊 Total de Expedientes:</b> %{customdata[0]:,} exp. (%{customdata[1]:.1f}%)<extra></extra>"
                ))
                anot_u = [
                    dict(
                        x=int(r["Total"]) + (max_u * 0.02), y=r["Unidad Orgánica"],
                        text=f"<b>{int(r['Total']):,}</b> ({r['%']:.1f}%)", xanchor="left", yanchor="middle",
                        showarrow=False, font=dict(size=11, color="#f8fafc")
                    ) for _, r in df_u_graf.iterrows()
                ]
                fig_1_4_u.update_layout(
                    title=dict(text="Gráfico 1.4.A — Top Unidades Orgánicas por Carga Procesal", font=dict(size=13, color="#f8fafc"), x=0.01),
                    barmode="stack", template="plotly_dark", paper_bgcolor="#111827", plot_bgcolor="#0f172a",
                    xaxis=dict(title="Expedientes", range=[0, max_u * 1.25], gridcolor="rgba(148, 163, 184, 0.14)"),
                    legend=dict(orientation="h", y=-0.22, x=0.5, xanchor="center"),
                    annotations=anot_u, height=max(260, len(df_u_graf) * 38 + 90), margin=dict(l=15, r=25, t=40, b=60),
                    hoverlabel=dict(bgcolor="#0f172a", bordercolor="#2dd4bf", font_size=12, font_color="#f8fafc")
                )
                st.plotly_chart(fig_1_4_u, use_container_width=True)

            st.markdown("---")
            st.markdown(f"**1.4.B. Carga Procesal por Magistrado de Control ({sub_titulo_sec})**")
            render_tabla_multinivel_anio(t_1_2_magistrado, anios_12m, "Magistrado de Control", col_num_orden="N°", max_height_css="max-height:480px; overflow-y:auto;")

            # Gráfico 1.4.B: Barras Apiladas Horizontales (Paleta Índigo / Ámbar)
            df_m_graf = t_1_2_magistrado[~t_1_2_magistrado["_es_total"]].head(10).iloc[::-1]
            if not df_m_graf.empty:
                max_m = max(1, int(df_m_graf["Total"].max()))
                fig_1_4_m = go.Figure()
                fig_1_4_m.add_trace(go.Bar(
                    y=df_m_graf["Magistrado de Control"], x=df_m_graf["Total - Trámite"], name="En Trámite", orientation="h",
                    marker=dict(color="#4f46e5", line=dict(color="#818cf8", width=1)),
                    text=[f"<b>{v:,}</b>" if v >= max_m * 0.07 else "" for v in df_m_graf["Total - Trámite"]],
                    textposition="inside", insidetextanchor="middle", textfont=dict(color="#ffffff", size=11),
                    customdata=np.stack([df_m_graf["Total"], df_m_graf["%"]], axis=-1),
                    hovertemplate="<b>⚖️ Magistrado:</b> %{y}<br><b>🟣 En Trámite:</b> %{x:,} exp.<br><b>📊 Total de Expedientes:</b> %{customdata[0]:,} exp. (%{customdata[1]:.1f}%)<extra></extra>"
                ))
                fig_1_4_m.add_trace(go.Bar(
                    y=df_m_graf["Magistrado de Control"], x=df_m_graf["Total - Res. Impug."], name="Res. Impug. Tránsito", orientation="h",
                    marker=dict(color="#d97706", line=dict(color="#fbbf24", width=1)),
                    text=[f"<b>{v:,}</b>" if v >= max_m * 0.07 else "" for v in df_m_graf["Total - Res. Impug."]],
                    textposition="inside", insidetextanchor="middle", textfont=dict(color="#ffffff", size=11),
                    customdata=np.stack([df_m_graf["Total"], df_m_graf["%"]], axis=-1),
                    hovertemplate="<b>⚖️ Magistrado:</b> %{y}<br><b>🟠 Res. Impug. Tránsito:</b> %{x:,} exp.<br><b>📊 Total de Expedientes:</b> %{customdata[0]:,} exp. (%{customdata[1]:.1f}%)<extra></extra>"
                ))
                anot_m = [
                    dict(
                        x=int(r["Total"]) + (max_m * 0.02), y=r["Magistrado de Control"],
                        text=f"<b>{int(r['Total']):,}</b> ({r['%']:.1f}%)", xanchor="left", yanchor="middle",
                        showarrow=False, font=dict(size=11, color="#f8fafc")
                    ) for _, r in df_m_graf.iterrows()
                ]
                fig_1_4_m.update_layout(
                    title=dict(text="Gráfico 1.4.B — Top 10 Magistrados de Control por Carga Procesal", font=dict(size=13, color="#f8fafc"), x=0.01),
                    barmode="stack", template="plotly_dark", paper_bgcolor="#111827", plot_bgcolor="#0f172a",
                    xaxis=dict(title="Expedientes", range=[0, max_m * 1.25], gridcolor="rgba(148, 163, 184, 0.14)"),
                    legend=dict(orientation="h", y=-0.22, x=0.5, xanchor="center"),
                    annotations=anot_m, height=max(260, len(df_m_graf) * 38 + 90), margin=dict(l=15, r=25, t=40, b=60),
                    hoverlabel=dict(bgcolor="#0f172a", bordercolor="#818cf8", font_size=12, font_color="#f8fafc")
                )
                st.plotly_chart(fig_1_4_m, use_container_width=True)
    # --------------------------------------------------------------------------
    # PESTAÑA 2: RIESGO DE PRESCRIPCIÓN (+ MAPA DE CALOR Y PARETO)
    # --------------------------------------------------------------------------
    with tab_presc:
        tot_p_sem = len(df_p_act)
        tot_p_vig = len(df_p_vigente)
        p_anc_vig = int(df_p_vigente["Es_Sede_Central"].sum())
        p_oda_vig = int((~df_p_vigente["Es_Sede_Central"]).sum())

        tot_ap = len(df_p_atencion_pref)
        cnt_m60 = int((df_p_atencion_pref["Atencion_Pref_Norm"].str.contains("60", na=False)).sum())
        cnt_m75 = int((df_p_atencion_pref["Atencion_Pref_Norm"].str.contains("75", na=False)).sum())
        sub_ap = f"↑ +60 años: {cnt_m60:,} | +75 años: {cnt_m75:,}" if (cnt_m60 > 0 or cnt_m75 > 0) else f"↑ {(tot_ap / max(1, tot_p_vig) * 100):.1f}% del riesgo vigente"

        etapa_top_p = df_p_vigente["Etapa_Presc_Cat"].value_counts() if not df_p_vigente.empty else df_p_act["Etapa_Presc_Cat"].value_counts()
        et_top_nom = etapa_top_p.index[0] if not etapa_top_p.empty else "Sin datos"
        et_top_cnt = int(etapa_top_p.iloc[0]) if not etapa_top_p.empty else 0
        base_et_pct = max(1, tot_p_vig if not df_p_vigente.empty else tot_p_sem)

        pk1, pk2, pk3, pk4 = st.columns(4)
        with pk1:
            sub_vig = f"ANC: {p_anc_vig:,} | ODANC: {p_oda_vig:,}" if es_nacional else f"Sede: {dist_sel}"
            if tot_p_vig > 0:
                st.button(f"TOTAL RIESGO PRESC. (1-240d)\n**{tot_p_vig:,} exp.**\n*{sub_vig}*", key="btn_presc_total_vig", on_click=alternar_kpi, args=("PRESC_TOTAL_VIG",), use_container_width=True)
            else:
                st.metric("TOTAL RIESGO PRESC. (1-240d)", "0 exp.", "Sin expedientes en riesgo", delta_color="off")
        with pk2:
            if tot_ap > 0:
                st.button(f"ATENCIÓN PREFERENCIAL\n**{tot_ap:,} exp.**\n*{sub_ap}*", key="btn_presc_atencion_pref", on_click=alternar_kpi, args=("PRESC_ATENCION_PREF",), use_container_width=True)
            else:
                st.metric("ATENCIÓN PREFERENCIAL", "0 exp.", "Sin atención preferencial", delta_color="off")
        pk3.metric("ETAPA CON MAYOR RIESGO", f"{et_top_nom}", f"{et_top_cnt:,} exp. ({et_top_cnt/base_et_pct*100:.1f}%)")
        if es_nacional:
            top1_dist = df_p_vigente[~df_p_vigente["Es_Sede_Central"]]["Distrito_Norm"].value_counts()
            pk4.metric("ODANC CON MAYOR RIESGO", f"{top1_dist.index[0].title() if not top1_dist.empty else '-'}", f"{int(top1_dist.iloc[0]) if not top1_dist.empty else 0:,} exp. en riesgo (1-240d)")
        else:
            df_p_ord = df_p_vigente.sort_values("Dias_Calc", ascending=True)
            if not df_p_ord.empty:
                fila_min = df_p_ord.iloc[0]
                dias_min = int(fila_min["Dias_Calc"])
                pk4.metric("EXP. MÁS PRÓXIMO A VENCER", f"Vence ({dias_min} {'día' if dias_min == 1 else 'días'})", formatear_detalle_exp_presc(fila_min))
            else:
                pk4.metric("EXP. MÁS PRÓXIMO A VENCER", "Sin exp. por vencer", "-", delta_color="off")

        st.markdown("<div style='margin-bottom: 12px;'></div>", unsafe_allow_html=True)

        conteos_rangos = df_p_act.groupby("Rango_Presc").size().reindex(ORDEN_RANGOS_PRESC, fill_value=0)
        sem_cols = st.columns(6)
        for col_s, (cod_k, clave_rg, titulo_rg, rango_txt, _, emoji) in zip(sem_cols, CONFIG_COLORES_RANGOS):
            val_s = int(conteos_rangos.loc[clave_rg])
            pct_s = (val_s / tot_p_sem * 100) if tot_p_sem > 0 else 0.0
            with col_s:
                st.button(f"{emoji} {titulo_rg}\n**{val_s:,} exp.**\n*{rango_txt} ({pct_s:.2f}%)*", key=f"btn_{cod_k}", on_click=alternar_kpi, args=(cod_k,), use_container_width=True)

        if not df_p_act.empty:
            gp1, gp2 = st.columns(2)
            with gp1:
                hm_data = pd.crosstab(df_p_act["Etapa_Presc_Cat"], df_p_act["Rango_Presc"]).reindex(index=ORDEN_ETAPAS, columns=ORDEN_RANGOS_PRESC, fill_value=0)
                hm_data.index = [ETAPAS_CORTAS[e] for e in ORDEN_ETAPAS]
                hm_data.columns = ["Presc (<=0d)", "Crítico (1-15d)", "Alto (16-30d)", "Moderado (31-60d)", "Prev. (61-90d)", "Monit. (91-240d)"]
                fig_hm = px.imshow(
                    hm_data, text_auto=True, aspect="auto", color_continuous_scale="YlOrRd",
                    title="Mapa de Calor: Riesgo por Etapa Procesal y Plazo",
                    labels=dict(x="Rango de Riesgo", y="", color="Exp.")
                )
                fig_hm.update_layout(template="plotly_dark", paper_bgcolor="#111827", height=360, margin=dict(l=15, r=15, t=45, b=20))
                st.plotly_chart(fig_hm, use_container_width=True)

            with gp2:
                pareto_s = df_p_act["Distrito_Norm"].value_counts().head(12) if es_nacional else df_p_act["Etapa_Presc_Cat"].map(ETAPAS_CORTAS).value_counts()
                if not es_nacional:
                    pareto_s = pareto_s[pareto_s > 0].sort_values(ascending=False)
                titulo_par = "Concentración de Riesgo (Pareto 80/20): Top 12 Sedes" if es_nacional else f"Concentración de Riesgo por Etapa — {dist_sel}"
                df_par = pd.DataFrame({"Categoria": [x.title() for x in pareto_s.index], "Expedientes": pareto_s.values})
                df_par["Acum_Pct"] = (df_par["Expedientes"].cumsum() / max(1, len(df_p_act)) * 100).round(1)

                fig_par = go.Figure()
                fig_par.add_trace(go.Bar(
                    x=df_par["Categoria"], y=df_par["Expedientes"], name="Expedientes",
                    marker_color="#3b82f6", text=df_par["Expedientes"], textposition="outside",
                    textfont=dict(color="#f8fafc", size=11)
                ))
                fig_par.add_trace(go.Scatter(x=df_par["Categoria"], y=df_par["Acum_Pct"], name="% Acumulado", yaxis="y2", mode="lines+markers", line=dict(color="#f59e0b", width=3)))
                fig_par.add_hline(y=80, yref="y2", line_dash="dash", line_color="#ef4444", annotation_text="80% del Riesgo", annotation_position="top left")
                fig_par.update_layout(
                    title=titulo_par, template="plotly_dark", paper_bgcolor="#111827", plot_bgcolor="#0f172a",
                    yaxis=dict(title="Expedientes", range=[0, max(1, df_par["Expedientes"].max()) * 1.18]),
                    yaxis2=dict(title="% Acumulado", overlaying="y", side="right", range=[0, 108]),
                    legend=dict(orientation="h", y=1.12), height=360, margin=dict(l=15, r=15, t=55, b=35)
                )
                st.plotly_chart(fig_par, use_container_width=True)

        st.subheader(f"2.1. Rango de Riesgo de Expedientes Prescritos y por Prescribir ({sub_titulo_sec})")
        mostrar_tabla_centrada(t_2_1_presc)

        st.markdown("---")
        st.subheader(f"2.2. Riesgo de Prescripción según Etapa Procesal ({sub_titulo_sec})")
        mostrar_tabla_centrada(t_2_2_et)

    # --------------------------------------------------------------------------
    # PESTAÑA 3: PRODUCCIÓN MENSUAL Y PRONUNCIAMIENTOS (+ EVOLUCIÓN Y RANKING)
    # --------------------------------------------------------------------------
    with tab_prod:
        inf_mes_pr = int(df_pr_mes_act["Es_Informe"].sum())
        res_mes_pr = len(df_pr_mes_act) - inf_mes_pr
        pr_anc_mes = int(df_pr_mes_act["Es_Sede_Central"].sum())
        pr_oda_mes = int((~df_pr_mes_act["Es_Sede_Central"]).sum())

        prk1, prk2, prk3, prk4 = st.columns(4)
        prk1.metric(f"PRODUCCIÓN {mes_act_lbl.upper()}", f"{len(df_pr_mes_act):,} exp.", f"ANC: {pr_anc_mes:,} | ODANC: {pr_oda_mes:,}" if es_nacional else f"Sede: {dist_sel}")
        prk2.metric(f"RESOLUCIONES ({mes_act_lbl.upper()})", f"{res_mes_pr:,} res.", f"{res_mes_pr/max(1,len(df_pr_mes_act))*100:.1f}% del mes actual")
        prk3.metric(f"INFORMES ({mes_act_lbl.upper()})", f"{inf_mes_pr:,} inf.", f"{inf_mes_pr/max(1,len(df_pr_mes_act))*100:.1f}% del mes actual")
        with prk4:
            sub_pr_acum = f"↑ Mes ant.: {int((df_pr_f['Periodo_M'] == per_ant).sum()):,} | Mes act.: {len(df_pr_mes_act):,}" if (not fechas_pr.empty and per_act is not None and per_ant is not None) else "↑ Ver detalle acumulado"
            st.button(f"PRODUCCIÓN ACUMULADA\n**{len(df_pr_f):,} exp.**\n*{sub_pr_acum}*", key="kpi_btn_prod_acumulada", on_click=alternar_kpi, args=("PROD_ACUMULADA",), use_container_width=True)

        st.markdown("<div style='margin-bottom: 10px;'></div>", unsafe_allow_html=True)

        df_pr_valid = df_pr_f.dropna(subset=["Fecha_Desc_Dt"]).copy()
        if not df_pr_valid.empty:
            gpr1, gpr2 = st.columns(2)
            with gpr1:
                df_pr_valid["Periodo_Str"] = df_pr_valid["Fecha_Desc_Dt"].dt.to_period("M").astype(str)
                serie_m = df_pr_valid.groupby("Periodo_Str").size().sort_index()
                media_pr = float(serie_m.mean())

                fig_evol = go.Figure()
                fig_evol.add_trace(go.Scatter(
                    x=serie_m.index, y=serie_m.values, mode="lines+markers+text",
                    text=[f"{v:,}" for v in serie_m.values], textposition="top center",
                    textfont=dict(color="#f8fafc", size=11),
                    name="Pronunciamientos", line=dict(color="#38bdf8", width=3), marker=dict(size=9, color="#38bdf8")
                ))
                fig_evol.add_hline(y=media_pr, line_dash="dash", line_color="#22c55e", annotation_text=f"Promedio Mensual ({media_pr:,.0f})", annotation_position="top left")
                fig_evol.update_layout(
                    title="Evolución Mensual de Producción vs. Promedio Histórico",
                    template="plotly_dark", paper_bgcolor="#111827", plot_bgcolor="#0f172a",
                    xaxis_title="Mes", yaxis_title="Pronunciamientos Emitidos",
                    yaxis=dict(range=[0, max(1, serie_m.max()) * 1.18]),
                    height=360, margin=dict(l=15, r=15, t=45, b=25)
                )
                st.plotly_chart(fig_evol, use_container_width=True)

            with gpr2:
                prod_mag = df_pr_mes_act["Magistrado_Prod"].value_counts().head(10).sort_values(ascending=True)
                med_mag = float(df_pr_mes_act["Magistrado_Prod"].value_counts().median()) if not df_pr_mes_act.empty else 0.0
                fig_mag = px.bar(
                    x=prod_mag.values, y=prod_mag.index, orientation="h", text=prod_mag.values,
                    title=f"Top Magistrados por Producción ({mes_act_lbl})",
                    labels={"x": "Pronunciamientos", "y": ""}
                )
                fig_mag.update_traces(marker_color="#6366f1", textposition="outside", textfont=dict(color="#f8fafc", size=11))
                fig_mag.add_vline(x=med_mag, line_dash="dash", line_color="#f59e0b", annotation_text=f"Mediana ({med_mag:.0f})", annotation_position="top right")
                fig_mag.update_layout(
                    template="plotly_dark", paper_bgcolor="#111827", plot_bgcolor="#0f172a",
                    xaxis=dict(range=[0, max(1, prod_mag.max()) * 1.18]),
                    height=360, margin=dict(l=15, r=25, t=45, b=25)
                )
                st.plotly_chart(fig_mag, use_container_width=True)

        st.subheader(f"3.1. Comparativo Mensual de Producción: Mes Anterior vs. Mes Actual ({sub_titulo_sec})")
        if not t_3_1.empty:
            mostrar_tabla_centrada(t_3_1)

        st.markdown("---")
        st.subheader(f"3.2. Producción por Tipo de Pronunciamiento y/o Sanción — {mes_act_lbl} ({sub_titulo_sec})")
        if not t_3_2.empty:
            mostrar_tabla_centrada(t_3_2)

        st.markdown("---")
        c_p3, c_p4 = st.columns(2)
        with c_p3:
            st.subheader(f"3.3. Producción por Unidad — {mes_act_lbl}")
            if not t_3_3.empty:
                mostrar_tabla_centrada(t_3_3)
        with c_p4:
            st.subheader(f"3.4. Producción por Magistrado — {mes_act_lbl}")
            if not t_3_4.empty:
                mostrar_tabla_centrada(t_3_4)

    # ==========================================================================
    # 11. EXPORTACIÓN DE INFORME EJECUTIVO GLOBAL (EXCEL Y PDF EN BARRA LATERAL)
    # ==========================================================================
    buffer_anexo = io.BytesIO()
    with pd.ExcelWriter(buffer_anexo, engine="openpyxl") as writer:
        t_1_1_anio.to_excel(writer, sheet_name="1_0_Carga_Por_Anio", index=False)
        t_1_2_etapas.to_excel(writer, sheet_name="1_1_Carga_Etapas", index=False)
        t_1_3_cal_mes.to_excel(writer, sheet_name="1_2_Calif_Mes_Anio", index=False)
        t_1_2_unidad.to_excel(writer, sheet_name="1_3_Carga_Unidad", index=False)
        t_1_2_magistrado.to_excel(writer, sheet_name="1_4_Carga_Magistrado", index=False)
        t_2_1_presc.to_excel(writer, sheet_name="2_1_Riesgo_Presc", index=False)
        t_2_2_et.to_excel(writer, sheet_name="2_2_Presc_Etapas", index=False)
        if not t_3_1.empty:
            t_3_1.to_excel(writer, sheet_name="3_1_Produccion_Mensual", index=False)
        if not t_3_2.empty:
            t_3_2.to_excel(writer, sheet_name="3_2_Prod_Pronunciamiento", index=False)
        if not t_3_3.empty:
            t_3_3.to_excel(writer, sheet_name="3_3_Prod_Unidad", index=False)
        if not t_3_4.empty:
            t_3_4.to_excel(writer, sheet_name="3_4_Prod_Magistrado", index=False)

    nombre_file_limpio = "TODOS" if es_nacional else dist_sel.replace(" ", "_").replace(".", "")

    def construir_pdf_anexo():
        from fpdf import FPDF
        pdf = FPDF(orientation="L", unit="mm", format="A4")
        pdf.set_auto_page_break(auto=True, margin=12)
        pdf.add_page()

        pdf.set_fill_color(127, 29, 29)
        pdf.rect(10, 10, 277, 18, "F")
        pdf.set_fill_color(245, 158, 11)
        pdf.rect(10, 10, 3, 18, "F")

        pdf.set_xy(16, 12)
        pdf.set_font("Helvetica", "B", 13)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(0, 6, limpiar_latin1("AUTORIDAD NACIONAL DE CONTROL DEL PODER JUDICIAL"), ln=True)

        pdf.set_x(16)
        pdf.set_font("Helvetica", "", 9)
        pdf.set_text_color(226, 232, 240)
        pdf.cell(0, 5, limpiar_latin1(f"ANÁLISIS ESTADÍSTICO DE CARGA, PRESCRIPCIONES Y PRODUCCIÓN: {nombre_ambito} | Corte Carga: {FECHA_CORTE_CARGA.strftime('%d/%m/%Y')} | Presc.: {FECHA_CORTE_PRESC.strftime('%d/%m/%Y')} (UIA)"), ln=True)
        pdf.ln(6)

        pdf.set_font("Helvetica", "B", 8.5)
        pdf.set_text_color(30, 41, 59)
        resumen_kpis = (
            f"CARGA PROCESAL: {len(df_act):,} exp.   |   "
            f"EN JEFATURA: {len(df_act[df_act['Despacho_Norm'].str.contains('JEFATURA', na=False)]):,} exp.   |   "
            f"RIESGO PRESC. (<=240d): {len(df_p_act):,} exp.   |   "
            f"PROD. {mes_act_lbl.upper()}: {len(df_pr_mes_act):,} exp.   |   "
            f"SIN INVENTARIAR: {len(df_no_inv):,} exp."
        )
        pdf.set_fill_color(241, 245, 249)
        pdf.cell(277, 8, limpiar_latin1(resumen_kpis), border=1, ln=True, align="C", fill=True)
        pdf.ln(4)

        def dibujar_tabla_pdf(titulo, df_tab, max_filas=45):
            if df_tab is None or df_tab.empty:
                return
            df_show = pd.concat([df_tab.head(max_filas - 1), df_tab.tail(1)], ignore_index=True) if len(df_tab) > max_filas else df_tab.copy()
            cols = [c for c in df_show.columns if not str(c).startswith("_")]
            n_cols = len(cols)
            w_primera = 95.0 if n_cols > 3 else 160.0
            w_resto = (277.0 - w_primera) / max(1, (n_cols - 1))
            anchos = [w_primera] + [w_resto] * (n_cols - 1)

            pdf.set_font("Helvetica", "B", 9.5)
            pdf.set_text_color(127, 29, 29)
            pdf.cell(0, 6, limpiar_latin1(titulo), ln=True)

            pdf.set_font("Helvetica", "B", 7.5)
            pdf.set_fill_color(30, 41, 59)
            pdf.set_text_color(255, 255, 255)
            for i, c in enumerate(cols):
                etiq_col = "Rango de Riesgo Presc. / Por Presc." if (len(str(c)) > 36 and i == 0) else str(c)
                pdf.cell(anchos[i], 6.5, limpiar_latin1(etiq_col[:38]), border=1, align="L" if i == 0 else "C", fill=True)
            pdf.ln()

            for idx, row in df_show.iterrows():
                prim = str(row[cols[0]])
                es_tot = any(p in prim.upper() for p in ["TOTAL", "SUBTOTAL"])
                pdf.set_font("Helvetica", "B" if es_tot else "", 7.5)
                pdf.set_fill_color(226, 232, 240) if es_tot else pdf.set_fill_color(255 if idx % 2 == 0 else 248, 255 if idx % 2 == 0 else 250, 255 if idx % 2 == 0 else 252)
                pdf.set_text_color(15, 23, 42) if es_tot else pdf.set_text_color(30, 41, 59)

                for i, c in enumerate(cols):
                    val = row[c]
                    txt = (f"{val:,.2f}" if "%" in str(c) else f"{int(val):,}") if isinstance(val, (int, float, np.integer, np.floating)) and pd.notna(val) else (str(val) if pd.notna(val) else "")
                    pdf.cell(anchos[i], 5.8, limpiar_latin1(txt[:45]), border=1, align="L" if i == 0 else "C", fill=True)
                pdf.ln()
            pdf.ln(3)

        dibujar_tabla_pdf(f"1. Carga Total por Año de Creación y Estado ({sub_titulo_sec})", t_1_1_anio)
        dibujar_tabla_pdf(f"1.1. Carga Procesal por Etapa y Sub-etapas de Calificación ({sub_titulo_sec})", t_1_2_etapas)
        dibujar_tabla_pdf(f"1.2. Carga en Calificación por Mes de Creación, Año y Estado ({sub_titulo_sec})", t_1_3_cal_mes)
        dibujar_tabla_pdf(f"2.1. Rango de Riesgo de Expedientes Prescritos y por Prescribir ({sub_titulo_sec})", t_2_1_presc)
        dibujar_tabla_pdf(f"2.2. Riesgo de Prescripción por Etapa Procesal ({sub_titulo_sec})", t_2_2_et)
        if not t_3_1.empty:
            dibujar_tabla_pdf(f"3.1. Comparativo Mensual de Producción ({sub_titulo_sec})", t_3_1)
        if not t_3_2.empty:
            dibujar_tabla_pdf(f"3.2. Producción por Tipo de Pronunciamiento / Sanción - {mes_act_lbl}", t_3_2)
        if not t_3_3.empty:
            dibujar_tabla_pdf(f"3.3. Producción por Unidad - {mes_act_lbl}", t_3_3, max_filas=60)
        if not t_3_4.empty:
            dibujar_tabla_pdf(f"3.4. Producción por Magistrado - {mes_act_lbl}", t_3_4, max_filas=60)

        return bytes(pdf.output())

    with st.sidebar:
        st.subheader("📥 Exportar Informe Oficial")
        st.download_button("📊 Descargar Informe en Excel", buffer_anexo.getvalue(), f"ANALISIS_ESTADISTICO_{nombre_file_limpio}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        st.download_button("📄 Descargar Reporte Ejecutivo (PDF)", construir_pdf_anexo(), f"REPORTE_ESTADISTICO_{nombre_file_limpio}.pdf", "application/pdf", use_container_width=True)