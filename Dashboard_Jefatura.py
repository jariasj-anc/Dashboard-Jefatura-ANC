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

# Fechas de corte institucionales (modo Excel activo):
FECHA_CORTE_CARGA = pd.Timestamp("2026-11-13")
FECHA_CORTE_PRESC = pd.Timestamp("2026-11-14")

# --- OPCIÓN DINÁMICA EN TIEMPO REAL (COMENTADA HASTA DESPLIEGUE UTI) ---
# def obtener_fechas_corte():
#     hoy = pd.Timestamp.today().normalize()
#     return hoy, hoy
# FECHA_CORTE_CARGA, FECHA_CORTE_PRESC = obtener_fechas_corte()

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

def obtener_nombre_unidad_por_sigla(enc, es_sc=False):
    s = norm_txt(enc)
    if "JEFATURA" in s:
        return "JEFATURA-ANC" if es_sc else "JEFATURA DE ODANC"
    if "-CII" in s or "INTERVENCION INMEDIATA" in s:
        return "COMISIÓN DE INTERVENCIÓN INMEDIATA"
    if "-OCCPC" in s or "CONTROL PREVENTIVO" in s:
        return "OFICINA CENTRAL DE CONTROL PREVENTIVO Y CONCURRENTE"
    if "-OCPAD" in s:
        return "OFICINA CENTRAL DE PROCEDIMIENTO ADMINISTRATIVO DISCIPLINARIO"
    # Unidades Descentralizadas (ODANC)
    if "-UDCIP" in s:
        return "UNIDAD DESCENTRALIZADA DE CALIFICACIÓN E INVESTIGACIÓN PRELIMINAR"
    if "-UDPSI" in s:
        return "UNIDAD DESCENTRALIZADA DE PREVENCIÓN, SUPERVISIÓN E INSPECCIÓN"
    if "-UDPAD" in s:
        return "UNIDAD DESCENTRALIZADA DE PROCEDIMIENTO ADMINISTRATIVO DISCIPLINARIO"
    if "-UDSAP" in s:
        return "UNIDAD DESCENTRALIZADA DE SANCIÓN Y APELACIÓN"
    # Unidades Centrales (ANC Sede Central)
    if "-UCIP" in s or "CALIFICACION" in s:
        return "UNIDAD DE CALIFICACIÓN E INVESTIGACIÓN PRELIMINAR" if es_sc else "UNIDAD DESCENTRALIZADA DE CALIFICACIÓN E INVESTIGACIÓN PRELIMINAR"
    if "-UPSI" in s or "PREVENCION" in s or "SUPERVISION" in s:
        return "UNIDAD DE PREVENCIÓN, SUPERVISIÓN E INSPECCIÓN" if es_sc else "UNIDAD DESCENTRALIZADA DE PREVENCIÓN, SUPERVISIÓN E INSPECCIÓN"
    if "-UPAD" in s or "PROCEDIMIENTO" in s:
        return "UNIDAD DE PROCEDIMIENTO ADMINISTRATIVO DISCIPLINARIO" if es_sc else "UNIDAD DESCENTRALIZADA DE PROCEDIMIENTO ADMINISTRATIVO DISCIPLINARIO"
    if "-USAP" in s or "SANCION" in s or "APELACION" in s:
        return "UNIDAD DE SANCIÓN Y APELACIÓN" if es_sc else "UNIDAD DESCENTRALIZADA DE SANCIÓN Y APELACIÓN"
    return "OTROS / ESPECIALIZADOS"

def map_unidad_produccion(row):
    return obtener_nombre_unidad_por_sigla(row.get("ENCARGADO", ""), bool(row.get("Es_Sede_Central", False)))

def map_unidad_todos(enc):
    s = norm_txt(enc)
    if "JEFATURA" in s:
        return "JEFATURA (ANC / ODANC)"
    if "-CII" in s or "INTERVENCION INMEDIATA" in s:
        return "COMISIÓN DE INTERVENCIÓN INMEDIATA"
    if "-OCCPC" in s or "CONTROL PREVENTIVO" in s:
        return "OFICINA CENTRAL DE CONTROL PREVENTIVO Y CONCURRENTE"
    if "-OCPAD" in s:
        return "OFICINA CENTRAL DE PROCEDIMIENTO ADMINISTRATIVO DISCIPLINARIO"
    if "-UDCIP" in s or "-UCIP" in s or "CALIFICACION" in s:
        return "UNIDAD DE CALIFICACIÓN E INVESTIGACIÓN PRELIMINAR (UCIP / UDCIP)"
    if "-UDPSI" in s or "-UPSI" in s or "PREVENCION" in s or "SUPERVISION" in s:
        return "UNIDAD DE PREVENCIÓN, SUPERVISIÓN E INSPECCIÓN (UPSI / UDPSI)"
    if "-UDPAD" in s or "-UPAD" in s or "PROCEDIMIENTO" in s:
        return "UNIDAD DE PROCEDIMIENTO ADMINISTRATIVO DISCIPLINARIO (UPAD / UDPAD)"
    if "-UDSAP" in s or "-USAP" in s or "SANCION" in s or "APELACION" in s:
        return "UNIDAD DE SANCIÓN Y APELACIÓN (USAP / UDSAP)"
    return "OTROS / ESPECIALIZADOS"

def map_magistrado_produccion(row):
    mag = str(row.get("MAGISTRADO", "")).strip().upper()
    if mag in ("", "NAN", "NONE", "NULL"):
        mag = "NO ESPECIFICADO"
    enc = norm_txt(row.get("ENCARGADO", ""))
    if row.get("Es_Sede_Central", False) and ("JEFATURA" in enc or "PALACIOS BRAN" in mag):
        return "JEFATURA SUPREMA"
    return mag

def estandarizar_tablas(df_c_raw, df_p_raw, df_pr_raw):
    fecha_corte_carga, fecha_corte_presc = FECHA_CORTE_CARGA, FECHA_CORTE_PRESC
    # fecha_corte_carga, fecha_corte_presc = obtener_fechas_corte()  # Tiempo real (comentado)
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
    df_c["Antiguedad_Dias"] = (fecha_corte_carga - df_c["Fecha_Creacion_Dt"]).dt.days
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

    df_p["Dias_Calc"] = (df_p["Fecha_Presc_Dt"] - fecha_corte_presc).dt.days
    df_p["Rango_Presc"] = df_p["Dias_Calc"].apply(clasificar_presc_rango)
    df_p = df_p[df_p["Rango_Presc"].notna()].copy()

    
    # C. PRODUCCIÓN MENSUAL
    df_pr = limpiar_columnas(df_pr_raw.copy())
    cols_pr_ok = [
        c for c in [
            "Distrito Actual", "EXPEDIENTE", "ENCARGADO", "MAGISTRADO",
            "USUARIO DESCARGO PROYECTO",
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
    if "USUARIO DESCARGO PROYECTO" in df_pr.columns:
        df_pr["Asistente_Prod"] = (
            df_pr["USUARIO DESCARGO PROYECTO"]
            .fillna("NO ESPECIFICADO")
            .astype(str)
            .str.strip()
            .str.upper()
            .replace({"": "NO ESPECIFICADO", "-": "NO ESPECIFICADO", "NAN": "NO ESPECIFICADO", "NONE": "NO ESPECIFICADO"})
        )
    else:
        df_pr["Asistente_Prod"] = "NO ESPECIFICADO"
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

# --- CONECTOR SYBASE ODBC EN TIEMPO REAL (COMENTADO HASTA ACTIVACIÓN UTI) ---
# @st.cache_data(ttl=300, show_spinner="Consultando base de datos Sybase en tiempo real...")
# def cargar_desde_sybase():
#     if pyodbc is None:
#         st.error("La librería pyodbc no está instalada. Ejecuta: pip install pyodbc")
#         st.stop()
#     cfg = st.secrets["sybase"] if "sybase" in st.secrets else {}
#     driver = cfg.get("driver", os.getenv("SYBASE_DRIVER", "Adaptive Server Enterprise"))
#     host = cfg.get("host", os.getenv("SYBASE_HOST", ""))
#     port = cfg.get("port", os.getenv("SYBASE_PORT", "5000"))
#     db = cfg.get("database", os.getenv("SYBASE_DB", "BD_ANC"))
#     user = cfg.get("user", os.getenv("SYBASE_USER", ""))
#     pwd = cfg.get("password", os.getenv("SYBASE_PASS", ""))
#     conn_str = (
#         f"DRIVER={{{driver}}};"
#         f"SERVER={host};"
#         f"PORT={port};"
#         f"DATABASE={db};"
#         f"UID={user};"
#         f"PWD={pwd};"
#     )
#     with pyodbc.connect(conn_str, timeout=15) as conn:
#         df_c_sql = pd.read_sql("SELECT * FROM vw_anc_carga_actual", conn)
#         df_p_sql = pd.read_sql("SELECT * FROM vw_anc_prescripciones", conn)
#         df_pr_sql = pd.read_sql("SELECT * FROM vw_anc_produccion", conn)
#     return estandarizar_tablas(df_c_sql, df_p_sql, df_pr_sql)

archivos_excel = [f for f in glob.glob("*.xlsx") if not os.path.basename(f).startswith("~$")]
if not archivos_excel:
    st.error("No se encontró ningún archivo Excel (.xlsx) en el directorio.")
    st.stop()
with open(archivos_excel[0], "rb") as f:
    df_carga, df_presc, df_prod = cargar_desde_excel(f.read())

# if MODO_DATOS == "SYBASE":
#     df_carga, df_presc, df_prod = cargar_desde_sybase()

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

    st.markdown("---")
    st.subheader("🧭 Módulos de Análisis")

    if "modulo_activo" not in st.session_state:
        st.session_state["modulo_activo"] = "CARGA"

    def cambiar_modulo(mod):
        st.session_state["modulo_activo"] = mod
        st.session_state["kpi_activo"] = None

    st.button(
        "📊 1. Carga Procesal Actual",
        key="nav_btn_carga",
        type="primary" if st.session_state["modulo_activo"] == "CARGA" else "secondary",
        on_click=cambiar_modulo,
        args=("CARGA",),
        use_container_width=True
    )
    st.button(
        "⏳ 2. Riesgo de Prescripción",
        key="nav_btn_presc",
        type="primary" if st.session_state["modulo_activo"] == "PRESC" else "secondary",
        on_click=cambiar_modulo,
        args=("PRESC",),
        use_container_width=True
    )
    st.button(
        "⚖️ 3. Producción Mensual",
        key="nav_btn_prod",
        type="primary" if st.session_state["modulo_activo"] == "PROD" else "secondary",
        on_click=cambiar_modulo,
        args=("PROD",),
        use_container_width=True
    )
    st.markdown("---")

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
df_tram_sede = df_c_sede[df_c_sede["Estado_Norm"] == "TRAMITE"].copy()
df_imp_sede = df_c_sede[df_c_sede["Estado_Norm"].str.contains("RESUELTO|IMPUGNACION", na=False)].copy()

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
def generar_cuadro_jerarquico_asistentes(df_mes_act, titulo_total="TOTAL GENERAL"):
    if df_mes_act.empty:
        return pd.DataFrame()

    total_global = len(df_mes_act)
    orden_mags = (
        df_mes_act.groupby("Magistrado_Prod")
        .size()
        .sort_values(ascending=False)
        .index.tolist()
    )

    filas = []
    tot_res_glob = 0
    tot_inf_glob = 0

    for idx_m, mag in enumerate(orden_mags, start=1):
        sub_m = df_mes_act[df_mes_act["Magistrado_Prod"] == mag]
        tot_m = len(sub_m)
        inf_m = int(sub_m["Es_Informe"].sum())
        res_m = int(tot_m - inf_m)
        pct_glob_m = (tot_m / total_global * 100.0) if total_global > 0 else 0.0

        tot_res_glob += res_m
        tot_inf_glob += inf_m

        filas.append({
            "Magistrado / Despacho → Asistente Judicial": f"🏛️ {idx_m}. {mag}",
            "Resoluciones": res_m,
            "Informes": inf_m,
            "Total": tot_m,
            "% Despacho": "100.00%",
            "% Total": f"{pct_glob_m:.2f}%"
        })

        g_asist = (
            sub_m.groupby("Asistente_Prod")
            .agg(Total=("Es_Informe", "size"), Informes=("Es_Informe", "sum"))
            .reset_index()
        )
        g_asist["Resoluciones"] = g_asist["Total"] - g_asist["Informes"]
        g_asist = g_asist.sort_values(by=["Total", "Resoluciones"], ascending=[False, False])

        for _, r_a in g_asist.iterrows():
            tot_a = int(r_a["Total"])
            res_a = int(r_a["Resoluciones"])
            inf_a = int(r_a["Informes"])
            pct_desp_a = (tot_a / tot_m * 100.0) if tot_m > 0 else 0.0
            pct_glob_a = (tot_a / total_global * 100.0) if total_global > 0 else 0.0

            filas.append({
                "Magistrado / Despacho → Asistente Judicial": f"    ↳ {r_a['Asistente_Prod']}",
                "Resoluciones": res_a,
                "Informes": inf_a,
                "Total": tot_a,
                "% Despacho": f"{pct_desp_a:.2f}%",
                "% Total": f"{pct_glob_a:.2f}%"
            })

    filas.append({
        "Magistrado / Despacho → Asistente Judicial": titulo_total,
        "Resoluciones": tot_res_glob,
        "Informes": tot_inf_glob,
        "Total": total_global,
        "% Despacho": "—",
        "% Total": "100.00%"
    })

    return pd.DataFrame(filas)
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
    # --------------------------------------------------------------------------
    # PESTAÑA 1: CARGA PROCESAL INTEGRAL (CUADRO + GRÁFICO ESPECIALIZADO)
    # --------------------------------------------------------------------------
    mod_actual = st.session_state.get("modulo_activo", "CARGA")
    if mod_actual == "CARGA":
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
    # PESTAÑA 2: RIESGO DE PRESCRIPCIÓN (CADA CUADRO CON SU GRÁFICO ASOCIADO)
    # --------------------------------------------------------------------------
    
    elif mod_actual == "PRESC":
        fp1, fp2, fp3 = st.columns([1.1, 1.2, 1.5])
        with fp1:
            opts_et_p = ["Todos"] + [e for e in ORDEN_ETAPAS if e in df_p_act["Etapa_Presc_Cat"].unique()]
            sel_et_p = st.selectbox("Filtrar por Etapa Procesal:", opts_et_p, key="filt_p_etapa")
        with fp2:
            opts_rg_p = ["Todos"] + list(ORDEN_RANGOS_PRESC)
            sel_rg_p = st.selectbox("Filtrar por Tramo de Riesgo:", opts_rg_p, key="filt_p_rango")
        with fp3:
            opts_mg_p = ["Todos"] + sorted([str(m).strip() for m in df_p_act["MAGISTRADO"].dropna().unique() if str(m).strip() not in ("", "nan", "None")])
            sel_mg_p = st.selectbox("Filtrar por Magistrado:", opts_mg_p, key="filt_p_mag")

        if sel_et_p != "Todos":
            df_p_act = df_p_act[df_p_act["Etapa_Presc_Cat"] == sel_et_p]
            df_p_vigente = df_p_vigente[df_p_vigente["Etapa_Presc_Cat"] == sel_et_p]
            df_p_atencion_pref = df_p_atencion_pref[df_p_atencion_pref["Etapa_Presc_Cat"] == sel_et_p]
        if sel_rg_p != "Todos":
            df_p_act = df_p_act[df_p_act["Rango_Presc"] == sel_rg_p]
            df_p_vigente = df_p_vigente[df_p_vigente["Rango_Presc"] == sel_rg_p]
            df_p_atencion_pref = df_p_atencion_pref[df_p_atencion_pref["Rango_Presc"] == sel_rg_p]
        if sel_mg_p != "Todos":
            df_p_act = df_p_act[df_p_act["MAGISTRADO"].astype(str).str.strip() == sel_mg_p]
            df_p_vigente = df_p_vigente[df_p_vigente["MAGISTRADO"].astype(str).str.strip() == sel_mg_p]
            df_p_atencion_pref = df_p_atencion_pref[df_p_atencion_pref["MAGISTRADO"].astype(str).str.strip() == sel_mg_p]

        # Recalcular Cuadros 2.1 y 2.2 con los filtros dinámicos aplicados
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

        # KPIs Superiores de Prescripción
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

        st.markdown("---")

        # ======================================================================
        # BLOQUE 2.1: RANGO DE RIESGO + GRÁFICO PARETO 80/20
        # ======================================================================
        st.subheader(f"2.1. Riesgo de Expedientes Prescritos y por Prescribir ({sub_titulo_sec})")
        mostrar_tabla_centrada(t_2_1_presc)

        if not df_p_act.empty:
            pareto_s = df_p_act["Distrito_Norm"].value_counts().head(12) if es_nacional else df_p_act["Rango_Presc"].value_counts().reindex(ORDEN_RANGOS_PRESC, fill_value=0)
            if not es_nacional:
                pareto_s = pareto_s[pareto_s > 0].sort_values(ascending=False)
            titulo_par = "Gráfico 2.1 — Concentración de Riesgo (Pareto 80/20): Top 12 Sedes" if es_nacional else f"Gráfico 2.1 — Concentración de Riesgo por Plazo (Pareto 80/20) — {dist_sel}"
            df_par = pd.DataFrame({"Categoria": [x.title() if es_nacional else x for x in pareto_s.index], "Expedientes": pareto_s.values})
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
                title=dict(text=titulo_par, font=dict(size=13.5, color="#f8fafc"), x=0.01),
                template="plotly_dark", paper_bgcolor="#111827", plot_bgcolor="#0f172a",
                yaxis=dict(title="Expedientes", range=[0, max(1, df_par["Expedientes"].max()) * 1.18]),
                yaxis2=dict(title="% Acumulado", overlaying="y", side="right", range=[0, 108]),
                legend=dict(orientation="h", y=1.12), height=360, margin=dict(l=15, r=15, t=55, b=35)
            )
            st.plotly_chart(fig_par, use_container_width=True)

        st.markdown("---")

        # ======================================================================
        # BLOQUE 2.2: RIESGO POR ETAPA PROCESAL + MAPA DE CALOR
        # ======================================================================
        st.subheader(f"2.2. Riesgo de Prescripción según Etapa Procesal ({sub_titulo_sec})")
        mostrar_tabla_centrada(t_2_2_et)

        if not df_p_act.empty:
            hm_data = pd.crosstab(df_p_act["Etapa_Presc_Cat"], df_p_act["Rango_Presc"]).reindex(index=ORDEN_ETAPAS, columns=ORDEN_RANGOS_PRESC, fill_value=0)
            hm_data.index = [ETAPAS_CORTAS[e] for e in ORDEN_ETAPAS]
            hm_data.columns = ["Presc (<=0d)", "Crítico (1-15d)", "Alto (16-30d)", "Moderado (31-60d)", "Prev. (61-90d)", "Monit. (91-240d)"]
            fig_hm = px.imshow(
                hm_data, text_auto=True, aspect="auto", color_continuous_scale="YlOrRd",
                title="Gráfico 2.2 — Mapa de Calor: Riesgo por Etapa Procesal y Plazo",
                labels=dict(x="Rango de Riesgo", y="", color="Exp.")
            )
            fig_hm.update_layout(
                title=dict(font=dict(size=13.5, color="#f8fafc"), x=0.01),
                template="plotly_dark", paper_bgcolor="#111827", height=360, margin=dict(l=15, r=15, t=45, b=25)
            )
            st.plotly_chart(fig_hm, use_container_width=True)

        st.markdown("---")

        # ======================================================================
        # BLOQUE 2.3: VELOCIDAD DE PRESCRIPCIÓN (CRONOGRAMA MENSUAL DE VENCIMIENTOS)
        # ======================================================================
        st.subheader(f"2.3. Velocidad de Prescripción — Cronograma Mensual de Vencimientos ({sub_titulo_sec})")
        if not df_p_act.empty:
            df_23_tl = df_p_act.dropna(subset=["Fecha_Presc_Dt"]).copy()
            if not df_23_tl.empty:
                df_23_tl["_Periodo"] = df_23_tl["Fecha_Presc_Dt"].dt.to_period("M")
                df_tl = df_23_tl.groupby("_Periodo").size().reset_index(name="Expedientes").sort_values("_Periodo")
                df_tl["Mes_Label"] = df_tl["_Periodo"].apply(lambda p: f"{MESES_ABR[p.month]} {p.year}")
                df_tl["Acumulado"] = df_tl["Expedientes"].cumsum()
                tot_tl = max(1, int(df_tl["Expedientes"].sum()))
                df_tl["Pct_Mes"] = (df_tl["Expedientes"] / tot_tl * 100.0).round(1)
                df_tl["Pct_Acum"] = (df_tl["Acumulado"] / tot_tl * 100.0).round(1)

                max_bar_tl = max(1, int(df_tl["Expedientes"].max()))
                max_acu_tl = max(1, int(df_tl["Acumulado"].max()))

                # Color dinámico por cercanía temporal (ya vencido / <= 2 meses / <= 4 meses / > 4 meses)
                per_corte = FECHA_CORTE_PRESC.to_period("M")
                colores_bar_tl = []
                for p_m in df_tl["_Periodo"]:
                    diff_m = (p_m.year - per_corte.year) * 12 + (p_m.month - per_corte.month)
                    if diff_m <= 0:
                        colores_bar_tl.append("#ef4444")   # Prescrito / Mes actual
                    elif diff_m <= 2:
                        colores_bar_tl.append("#f97316")   # Urgente (1-2 meses)
                    elif diff_m <= 4:
                        colores_bar_tl.append("#2563eb")   # Mediano plazo
                    else:
                        colores_bar_tl.append("#0284c7")   # Horizonte preventivo

                # Etiquetas dentro de la barra si tiene altura suficiente; afuera solo si es muy baja
                txt_bar_in = [
                    f"<b>{int(r['Expedientes'])}</b><br><span style='font-size:10px;'>({r['Pct_Mes']:.1f}%)</span>"
                    if r["Expedientes"] >= max_bar_tl * 0.22 else ""
                    for _, r in df_tl.iterrows()
                ]
                txt_bar_out = [
                    f"<b>{int(r['Expedientes'])}</b> ({r['Pct_Mes']:.1f}%)"
                    if r["Expedientes"] < max_bar_tl * 0.22 else ""
                    for _, r in df_tl.iterrows()
                ]

                custom_tl = np.stack([df_tl["Pct_Mes"], df_tl["Acumulado"], df_tl["Pct_Acum"]], axis=-1)

                fig_2_3_tl = go.Figure()

                # Traza 1: Barras de vencimientos mensuales con etiquetas internas limpias
                fig_2_3_tl.add_trace(go.Bar(
                    x=df_tl["Mes_Label"],
                    y=df_tl["Expedientes"],
                    name="Vencimientos del Mes",
                    width=0.46,
                    marker=dict(
                        color=colores_bar_tl,
                        line=dict(color="rgba(255,255,255,0.22)", width=1.2)
                    ),
                    text=txt_bar_in,
                    textposition="inside",
                    insidetextanchor="middle",
                    textfont=dict(color="#ffffff", size=11.5),
                    customdata=custom_tl,
                    hovertemplate=(
                        "<b>📅 Mes:</b> %{x}<br>"
                        "<b>⏳ Vencen este mes:</b> %{y:,} exp. (%{customdata[0]:.1f}%)<br>"
                        "<b>📈 Riesgo acumulado:</b> %{customdata[1]:,.0f} exp. (%{customdata[2]:.1f}%)"
                        "<extra></extra>"
                    )
                ))

                # Etiquetas superiores solo para barras muy pequeñas (< 22% de altura)
                if any(txt_bar_out):
                    fig_2_3_tl.add_trace(go.Scatter(
                        x=df_tl["Mes_Label"],
                        y=df_tl["Expedientes"],
                        mode="text",
                        text=txt_bar_out,
                        textposition="top center",
                        textfont=dict(color="#e2e8f0", size=10.5),
                        showlegend=False,
                        hoverinfo="skip"
                    ))

                # Traza 2: Curva de riesgo acumulado con área sutil
                fig_2_3_tl.add_trace(go.Scatter(
                    x=df_tl["Mes_Label"],
                    y=df_tl["Acumulado"],
                    name="Riesgo Acumulado",
                    mode="lines+markers",
                    yaxis="y2",
                    fill="tozeroy",
                    fillcolor="rgba(245, 158, 11, 0.07)",
                    line=dict(color="#f59e0b", width=3.2, shape="spline", smoothing=0.35),
                    marker=dict(size=9, color="#fbbf24", line=dict(color="#111827", width=2)),
                    customdata=custom_tl,
                    hovertemplate=(
                        "<b>📅 Hasta:</b> %{x}<br>"
                        "<b>📈 Riesgo Acumulado:</b> %{y:,} exp. (%{customdata[2]:.1f}% del total)<br>"
                        "<b>⏳ Vencimientos del mes:</b> %{customdata[0]:.1f}%"
                        "<extra></extra>"
                    )
                ))

                # Pastillas flotantes (badges) para la curva acumulada: cero cruce con barras
                anot_acum = []
                for _, r in df_tl.iterrows():
                    ratio_bar = r["Expedientes"] / max_bar_tl
                    ratio_acu = r["Acumulado"] / max_acu_tl
                    y_shift = 20 if (ratio_bar < 0.22 and ratio_acu < 0.25) else 18
                    anot_acum.append(dict(
                        x=r["Mes_Label"],
                        y=r["Acumulado"],
                        yref="y2",
                        text=f"<b>{int(r['Acumulado'])}</b> ({r['Pct_Acum']:.0f}%)",
                        showarrow=False,
                        yshift=y_shift,
                        font=dict(size=10.5, color="#fde68a"),
                        bgcolor="rgba(15, 23, 42, 0.90)",
                        bordercolor="rgba(245, 158, 11, 0.65)",
                        borderwidth=1,
                        borderpad=3
                    ))

                # Línea de referencia al 80% del acumulado
                umbral_80_tl = max_acu_tl * 0.80
                fig_2_3_tl.add_hline(
                    y=umbral_80_tl, yref="y2",
                    line_dash="dot", line_color="rgba(248, 113, 113, 0.65)", line_width=1.4,
                    annotation_text="Umbral 80% Acumulado",
                    annotation_position="top left",
                    annotation_font=dict(color="#fca5a5", size=10.5)
                )

                fig_2_3_tl.update_layout(
                    title=dict(
                        text=f"Gráfico 2.3 — Proyección Mensual de Vencimientos y Curva Acumulada ({sub_titulo_sec})",
                        font=dict(size=13.5, color="#f8fafc"), x=0.01, y=0.96
                    ),
                    template="plotly_dark",
                    paper_bgcolor="#111827",
                    plot_bgcolor="#0f172a",
                    legend=dict(
                        orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
                        font=dict(size=11.5, color="#f8fafc")
                    ),
                    xaxis=dict(
                        title="Mes de Vencimiento",
                        type="category",
                        tickfont=dict(size=11.5, color="#f8fafc"),
                        showgrid=False
                    ),
                    yaxis=dict(
                        title="Expedientes por Mes",
                        range=[0, max_bar_tl * 1.25],
                        gridcolor="rgba(148, 163, 184, 0.12)",
                        zeroline=False
                    ),
                    yaxis2=dict(
                        title="Expedientes Acumulados",
                        overlaying="y",
                        side="right",
                        range=[0, max_acu_tl * 1.24],
                        tickformat=",d",
                        nticks=6,
                        showgrid=False,
                        tickfont=dict(color="#fbbf24")
                    ),
                    annotations=anot_acum,
                    height=410,
                    margin=dict(l=20, r=30, t=68, b=42),
                    hoverlabel=dict(bgcolor="#0f172a", bordercolor="#f59e0b", font_size=12, font_color="#f8fafc")
                )
                st.plotly_chart(fig_2_3_tl, use_container_width=True)

        st.markdown("---")

        # ======================================================================
        # BLOQUE 2.4: CARGA DE PRESCRIPCIÓN POR MAGISTRADO / RESPONSABLE
        # ======================================================================
        st.subheader(f"2.4. Riesgo de Prescripción por Magistrado ({sub_titulo_sec})")
        if not df_p_act.empty:
            df_24_m = df_p_act.copy()
            df_24_m["Mag_Clean"] = (
                df_24_m["MAGISTRADO"].fillna("<SIN MAGISTRADO>").astype(str).str.strip()
                .replace({"": "<SIN MAGISTRADO>", "-": "<SIN MAGISTRADO>", "None": "<SIN MAGISTRADO>", "nan": "<SIN MAGISTRADO>"})
            )
            ct_mag_rg = pd.crosstab(df_24_m["Mag_Clean"], df_24_m["Rango_Presc"]).reindex(columns=ORDEN_RANGOS_PRESC, fill_value=0)
            t_2_4_mag = pd.DataFrame(index=ct_mag_rg.index)
            t_2_4_mag["🔴 Prescritos (<=0d)"] = ct_mag_rg["Prescritos (<= 0 días)"]
            t_2_4_mag["🟡 Crítico (1-15d)"] = ct_mag_rg["Crítico: De 1 a 15 días"]
            t_2_4_mag["🟢 Alto (16-30d)"] = ct_mag_rg["Alto: De 16 a 30 días"]
            t_2_4_mag["🔵 Moderado (31-60d)"] = ct_mag_rg["Moderado: De 31 a 60 días"]
            t_2_4_mag["🔵 Preventivo (61-90d)"] = ct_mag_rg["Preventivo: De 61 a 90 días"]
            t_2_4_mag["🟣 Monitoreo (91-240d)"] = ct_mag_rg["Monitoreo: De 91 a 240 días"]
            t_2_4_mag["Total"] = t_2_4_mag.sum(axis=1)
            t_2_4_mag = t_2_4_mag.sort_values(by=["Total", "🔴 Prescritos (<=0d)", "🟡 Crítico (1-15d)"], ascending=[False, False, False])

            tot_mag_p = max(1, int(t_2_4_mag["Total"].sum()))
            t_2_4_mag["% Part."] = (t_2_4_mag["Total"] / tot_mag_p * 100).round(2)
            t_2_4_mag = t_2_4_mag.reset_index().rename(columns={"Mag_Clean": "Magistrado"})
            t_2_4_mag.insert(0, "N°", [str(i) for i in range(1, len(t_2_4_mag) + 1)])

            fila_tot_24m = {
                "N°": "",
                "Magistrado": "TOTAL POR MAGISTRADO",
                **{c: int(t_2_4_mag[c].sum()) for c in t_2_4_mag.columns if c not in ("N°", "Magistrado", "% Part.")},
                "% Part.": 100.0 if int(t_2_4_mag["Total"].sum()) > 0 else 0.0
            }
            t_2_4_mag = pd.concat([t_2_4_mag, pd.DataFrame([fila_tot_24m])], ignore_index=True)
            mostrar_tabla_centrada(t_2_4_mag)

    # --------------------------------------------------------------------------
    # PESTAÑA 3: PRODUCCIÓN MENSUAL Y PRONUNCIAMIENTOS (CUADRO + GRÁFICO)
    # --------------------------------------------------------------------------
    elif mod_actual == "PROD":
        if df_pr_f.empty:
            st.info("No se encontraron registros en la hoja 'Producción'.")
        else:
            MESES_ES = {
                1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
                5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
                9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
            }

            # 1. Filtro base por Distrito / Sede (Sidebar)
            df_pr_base = df_pr_f.copy()
            df_pr_base["Anio_Desc"] = df_pr_base["Fecha_Desc_Dt"].dt.year
            df_pr_base["Mes_Num"] = df_pr_base["Fecha_Desc_Dt"].dt.month

            # 2. Panel de Filtros Multidimensionales en Cascada (2 filas x 3 columnas)
            with st.container(border=True):
                f_c1, f_c2, f_c3 = st.columns(3)
                anios_disp = sorted(df_pr_base["Anio_Desc"].dropna().astype(int).unique().tolist(), reverse=True)
                anio_sel = f_c1.selectbox("📅 Año:", anios_disp if anios_disp else [2026], key="pr_anio")

                df_anio = df_pr_base[df_pr_base["Anio_Desc"] == anio_sel].copy()
                meses_disp = sorted(df_anio["Mes_Num"].dropna().astype(int).unique().tolist())
                opciones_mes = ["Todos (Acumulado Anual)"] + [MESES_ES[m] for m in meses_disp]
                idx_mes_def = len(opciones_mes) - 1 if len(opciones_mes) > 1 else 0
                mes_sel_lbl = f_c2.selectbox("🗓️ Mes:", opciones_mes, index=idx_mes_def, key="pr_mes")

                col_u_pr = "Unidad_Prod_Todos" if es_nacional else "Unidad_Prod"
                unidades_disp = ["Todas"] + sorted(df_anio[col_u_pr].dropna().unique().tolist())
                unidad_sel = f_c3.selectbox("🏛️ Despacho o Unidad:", unidades_disp, key="pr_unidad")

                f_c4, f_c5, f_c6 = st.columns(3)
                df_temp_u = df_anio if unidad_sel == "Todas" else df_anio[df_anio[col_u_pr] == unidad_sel]
                mags_disp = ["Todos"] + sorted(df_temp_u["Magistrado_Prod"].dropna().unique().tolist())
                mag_sel = f_c4.selectbox("⚖️ Magistrado:", mags_disp, key="pr_mag")

                df_temp_m = df_temp_u if mag_sel == "Todos" else df_temp_u[df_temp_u["Magistrado_Prod"] == mag_sel]
                asist_disp = ["Todos"] + sorted(df_temp_m["Asistente_Prod"].dropna().unique().tolist())
                asist_sel = f_c5.selectbox("🧑‍💼 Asistente Judicial (Usuario Descarga):", asist_disp, key="pr_asist")

                col_pron = "DETALLE DEL TIPO DE PRONUNCIAMIENTO" if "DETALLE DEL TIPO DE PRONUNCIAMIENTO" in df_temp_m.columns else "TIPO DE AUTO"
                prons_disp = ["Todos"] + sorted(df_temp_m[col_pron].fillna("NO ESPECIFICADO").astype(str).str.strip().unique().tolist())
                pron_sel = f_c6.selectbox("📑 Tipo de Pronunciamiento:", prons_disp, key="pr_pron")

            # 3. Aplicar filtros transversales
            df_pr_evol = df_anio.copy()
            if unidad_sel != "Todas":
                df_pr_evol = df_pr_evol[df_pr_evol[col_u_pr] == unidad_sel]
            if mag_sel != "Todos":
                df_pr_evol = df_pr_evol[df_pr_evol["Magistrado_Prod"] == mag_sel]
            if asist_sel != "Todos":
                df_pr_evol = df_pr_evol[df_pr_evol["Asistente_Prod"] == asist_sel]
            if pron_sel != "Todos":
                df_pr_evol = df_pr_evol[df_pr_evol[col_pron].fillna("NO ESPECIFICADO").astype(str).str.strip() == pron_sel]

            if mes_sel_lbl == "Todos (Acumulado Anual)":
                df_pr_detalle = df_pr_evol.copy()
                periodo_lbl = f"Año {anio_sel} (Acumulado)"
            else:
                mes_num_sel = [k for k, v in MESES_ES.items() if v == mes_sel_lbl][0]
                df_pr_detalle = df_pr_evol[df_pr_evol["Mes_Num"] == mes_num_sel].copy()
                periodo_lbl = f"{mes_sel_lbl} {anio_sel}"

            sub_titulo_sec = "TODOS LOS DISTRITOS JUDICIALES" if es_nacional else dist_sel

            # 4. Tarjetas KPI Ejecutivas
            inf_mes_pr = int(df_pr_detalle["Es_Informe"].sum())
            res_mes_pr = len(df_pr_detalle) - inf_mes_pr
            pr_anc_mes = int(df_pr_detalle["Es_Sede_Central"].sum())
            pr_oda_mes = int((~df_pr_detalle["Es_Sede_Central"]).sum())

            prk1, prk2, prk3, prk4 = st.columns(4)
            prk1.metric(f"PRODUCCIÓN {periodo_lbl.upper()}", f"{len(df_pr_detalle):,} exp.", f"ANC Sede: {pr_anc_mes:,} | ODANC: {pr_oda_mes:,}")
            prk2.metric(f"RESOLUCIONES ({periodo_lbl.upper()})", f"{res_mes_pr:,} res.", f"{(res_mes_pr / max(1, len(df_pr_detalle)) * 100):.1f}% del periodo")
            prk3.metric(f"INFORMES ({periodo_lbl.upper()})", f"{inf_mes_pr:,} inf.", f"{(inf_mes_pr / max(1, len(df_pr_detalle)) * 100):.1f}% del periodo")
            prk4.metric(f"ACUMULADO ANUAL {anio_sel}", f"{len(df_pr_evol):,} exp.", f"Meses activos: {df_pr_evol['Mes_Num'].nunique()}")

            st.markdown("---")

            # ======================================================================
            # BLOQUE 3.1: EVOLUCIÓN MENSUAL POR NOMBRE DE MES (Enero, Febrero...)
            # ======================================================================
            st.subheader(f"3.1. Evolución Mensual de Producción — {anio_sel} ({sub_titulo_sec})")
            if not df_pr_evol.empty:
                serie_m = df_pr_evol.groupby("Mes_Num").size().sort_index()
                nombres_meses = [MESES_ES[int(m)] for m in serie_m.index]
                media_pr = float(serie_m.mean())

                fig_evol = go.Figure()
                fig_evol.add_trace(go.Scatter(
                    x=nombres_meses, y=serie_m.values, mode="lines+markers+text",
                    fill="tozeroy", fillcolor="rgba(56, 189, 248, 0.14)",
                    text=[f"<b>{v:,}</b>" for v in serie_m.values], textposition="top center",
                    textfont=dict(color="#f8fafc", size=11.5), name="Pronunciamientos",
                    line=dict(color="#38bdf8", width=3),
                    marker=dict(size=9, color="#38bdf8", line=dict(color="#0f172a", width=1.5)),
                    hovertemplate="<b>🗓️ Mes:</b> %{x}<br><b>⚖️ Pronunciamientos:</b> %{y:,}<extra></extra>"
                ))
                fig_evol.add_hline(
                    y=media_pr, line_dash="dash", line_color="#22c55e", line_width=1.8,
                    annotation_text=f"Promedio Mensual ({media_pr:,.0f})", annotation_position="top left",
                    annotation_font=dict(color="#4ade80", size=11)
                )
                fig_evol.update_layout(
                    title=dict(text=f"Gráfico 3.1 — Evolución Mensual de Producción ({anio_sel})", font=dict(size=13.5, color="#f8fafc"), x=0.01),
                    template="plotly_dark", paper_bgcolor="#111827", plot_bgcolor="#0f172a",
                    xaxis=dict(type="category", title="Mes", tickfont=dict(size=11.5, color="#f8fafc"), range=[-0.4, len(nombres_meses) - 0.6]),
                    yaxis=dict(title="Pronunciamientos Emitidos", range=[0, max(1, serie_m.max()) * 1.24], gridcolor="rgba(148, 163, 184, 0.14)"),
                    height=340, margin=dict(l=20, r=20, t=48, b=35)
                )
                st.plotly_chart(fig_evol, use_container_width=True)

            st.markdown("---")

            # ======================================================================
            # BLOQUE 3.2: TIPO DE PRONUNCIAMIENTO Y/O SANCIÓN
            # ======================================================================
            st.subheader(f"3.2. Producción por Tipo de Pronunciamiento y/o Sanción — {periodo_lbl} ({sub_titulo_sec})")
            t_3_2 = pd.DataFrame()
            if not df_pr_detalle.empty and "DETALLE DEL TIPO DE PRONUNCIAMIENTO" in df_pr_detalle.columns:
                df_pr_detalle["Pron_Limpio"] = df_pr_detalle["DETALLE DEL TIPO DE PRONUNCIAMIENTO"].fillna("NO ESPECIFICADO").astype(str).str.strip()
                tot_lbl_32 = f"TOTAL {periodo_lbl.upper()} ({'TODOS' if es_nacional else dist_sel})"
                if es_nacional:
                    t_3_2_full = pd.DataFrame({
                        "ANC Sede Central": df_pr_detalle[df_pr_detalle["Es_Sede_Central"]].groupby("Pron_Limpio").size(),
                        "ODANCs (34 Dist.)": df_pr_detalle[~df_pr_detalle["Es_Sede_Central"]].groupby("Pron_Limpio").size()
                    }).fillna(0).astype(int)
                    t_3_2_full["Total Nacional"] = t_3_2_full["ANC Sede Central"] + t_3_2_full["ODANCs (34 Dist.)"]
                    t_3_2_full = t_3_2_full.sort_values("Total Nacional", ascending=False).reset_index().rename(columns={"Pron_Limpio": "Tipo de Pronunciamiento / Sanción"})
                    tot_32_nac = int(t_3_2_full["Total Nacional"].sum())
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
                        "Tipo de Pronunciamiento / Sanción": tot_lbl_32,
                        "ANC Sede Central": int(t_3_2["ANC Sede Central"].sum()),
                        "ODANCs (34 Dist.)": int(t_3_2["ODANCs (34 Dist.)"].sum()),
                        "Total Nacional": tot_32_nac,
                        "% Nac.": 100.0 if tot_32_nac else 0.0
                    }])
                    t_3_2 = pd.concat([t_3_2, fila_tot_32], ignore_index=True)
                else:
                    g_32 = df_pr_detalle.groupby("Pron_Limpio").agg(
                        Resoluciones=("Es_Informe", lambda s: int((~s).sum())),
                        Informes=("Es_Informe", lambda s: int(s.sum())),
                        Total=("Es_Informe", "size")
                    ).sort_values("Total", ascending=False).reset_index().rename(columns={"Pron_Limpio": "Tipo de Pronunciamiento / Sanción"})
                    tot_32_loc = int(g_32["Total"].sum())
                    if len(g_32) > 15:
                        resto = g_32.iloc[15:]
                        fila_otros = pd.DataFrame([{
                            "Tipo de Pronunciamiento / Sanción": f"Otros pronunciamientos ({len(resto)} tipos)",
                            "Resoluciones": int(resto["Resoluciones"].sum()),
                            "Informes": int(resto["Informes"].sum()),
                            "Total": int(resto["Total"].sum())
                        }])
                        t_3_2 = pd.concat([g_32.head(15), fila_otros], ignore_index=True)
                    else:
                        t_3_2 = g_32.copy()
                    t_3_2["% Part."] = (t_3_2["Total"] / tot_32_loc * 100).round(2) if tot_32_loc else 0.0
                    fila_tot_32 = pd.DataFrame([{
                        "Tipo de Pronunciamiento / Sanción": tot_lbl_32,
                        "Resoluciones": int(t_3_2["Resoluciones"].sum()),
                        "Informes": int(t_3_2["Informes"].sum()),
                        "Total": tot_32_loc,
                        "% Part.": 100.0 if tot_32_loc else 0.0
                    }])
                    t_3_2 = pd.concat([t_3_2, fila_tot_32], ignore_index=True)
            if not t_3_2.empty:
                mostrar_tabla_centrada(t_3_2)
                col_val_32 = "Total Nacional" if "Total Nacional" in t_3_2.columns else "Total"
                col_pct_32 = "% Nac." if "% Nac." in t_3_2.columns else "% Part."

                df_g_32 = t_3_2[~t_3_2["Tipo de Pronunciamiento / Sanción"].str.upper().str.startswith("TOTAL")].head(12).iloc[::-1].copy()
                if not df_g_32.empty:
                    max_32 = max(1, int(df_g_32[col_val_32].max()))
                    fig_3_2 = go.Figure(go.Bar(
                        y=df_g_32["Tipo de Pronunciamiento / Sanción"], x=df_g_32[col_val_32], orientation="h",
                        marker=dict(color=df_g_32[col_val_32], colorscale=[[0.0, "#0369a1"], [0.5, "#0284c7"], [1.0, "#10b981"]], line=dict(color="#38bdf8", width=1)),
                        text=[f"<b>{int(v):,}</b> ({p:.1f}%)" for v, p in zip(df_g_32[col_val_32], df_g_32[col_pct_32])],
                        textposition="outside", textfont=dict(color="#f8fafc", size=11.5),
                        customdata=np.stack([df_g_32[col_pct_32]], axis=-1),
                        hovertemplate="<b>⚖️ Pronunciamiento:</b> %{y}<br><b>📊 Cantidad Emitida:</b> %{x:,} exp. (%{customdata[0]:.2f}%)<extra></extra>"
                    ))
                    fig_3_2.update_layout(
                        title=dict(text=f"Gráfico 3.2 — Distribución por Tipo de Pronunciamiento y/o Sanción ({periodo_lbl})", font=dict(size=13.5, color="#f8fafc"), x=0.01),
                        template="plotly_dark", paper_bgcolor="#111827", plot_bgcolor="#0f172a",
                        xaxis=dict(title="Cantidad de Pronunciamientos", range=[0, max_32 * 1.24], gridcolor="rgba(148, 163, 184, 0.14)"),
                        yaxis=dict(title="", tickfont=dict(size=11.5, color="#f8fafc")),
                        height=max(340, len(df_g_32) * 34 + 90), margin=dict(l=15, r=30, t=48, b=35)
                    )
                    st.plotly_chart(fig_3_2, use_container_width=True)

            st.markdown("---")

            # ======================================================================
            # BLOQUE 3.3: PRODUCCIÓN POR DESPACHO O UNIDAD
            # ======================================================================
            st.subheader(f"3.3. Producción por Despacho o Unidad — {periodo_lbl} ({sub_titulo_sec})")
            t_3_3 = generar_cuadro_produccion_mes_actual(
                df_pr_detalle,
                col_u_pr,
                "Unidad Orgánica / Funcional",
                f"TOTAL {periodo_lbl.upper()} ({'TODOS' if es_nacional else dist_sel})"
            )
            if not t_3_3.empty:
                mostrar_tabla_centrada(t_3_3)
                df_g_33 = t_3_3[~t_3_3["Unidad Orgánica / Funcional"].str.upper().str.startswith("TOTAL")].head(12).iloc[::-1].copy()
                if not df_g_33.empty:
                    max_33 = max(1, int(df_g_33["Total"].max()))
                    fig_3_3 = go.Figure()
                    fig_3_3.add_trace(go.Bar(
                        y=df_g_33["Unidad Orgánica / Funcional"], x=df_g_33["Resoluciones"], name="Resoluciones", orientation="h",
                        marker=dict(color="#0284c7", line=dict(color="#38bdf8", width=1)),
                        text=[f"<b>{int(v):,}</b>" if v >= max_33 * 0.06 else "" for v in df_g_33["Resoluciones"]],
                        textposition="inside", insidetextanchor="middle", textfont=dict(color="#ffffff", size=11)
                    ))
                    fig_3_3.add_trace(go.Bar(
                        y=df_g_33["Unidad Orgánica / Funcional"], x=df_g_33["Informes"], name="Informes", orientation="h",
                        marker=dict(color="#f59e0b", line=dict(color="#fbbf24", width=1)),
                        text=[f"<b>{int(v):,}</b>" if v >= max_33 * 0.06 else "" for v in df_g_33["Informes"]],
                        textposition="inside", insidetextanchor="middle", textfont=dict(color="#111827", size=11)
                    ))
                    anot_33 = [
                        dict(x=int(r["Total"]) + (max_33 * 0.018), y=r["Unidad Orgánica / Funcional"],
                            text=f"<b>{int(r['Total']):,}</b> ({r['% Part.']:.1f}%)", xanchor="left", yanchor="middle", showarrow=False,
                            font=dict(size=11.5, color="#f8fafc"))
                        for _, r in df_g_33.iterrows()
                    ]
                    fig_3_3.update_layout(
                        title=dict(text=f"Gráfico 3.3 — Desglose de Resoluciones e Informes por Despacho o Unidad ({periodo_lbl})", font=dict(size=13.5, color="#f8fafc"), x=0.01, y=0.95),
                        barmode="stack", bargap=0.30, template="plotly_dark", paper_bgcolor="#111827", plot_bgcolor="#0f172a",
                        xaxis=dict(title="Cantidad de Pronunciamientos Emitidos", range=[0, max_33 * 1.22], gridcolor="rgba(148, 163, 184, 0.14)"),
                        yaxis=dict(title="", tickfont=dict(size=11, color="#f8fafc")),
                        legend=dict(orientation="h", y=1.08, x=1.0, xanchor="right", font=dict(size=11.5, color="#f8fafc")),
                        annotations=anot_33, height=max(320, len(df_g_33) * 42 + 95), margin=dict(l=15, r=35, t=60, b=40)
                    )
                    st.plotly_chart(fig_3_3, use_container_width=True)

            st.markdown("---")

            # ======================================================================
            # BLOQUE 3.4: PRODUCCIÓN POR MAGISTRADO + GRÁFICO LOLLIPOP EJECUTIVO
            # ======================================================================
            st.subheader(f"3.4. Producción por Magistrado — {periodo_lbl} ({sub_titulo_sec})")
            t_3_4 = generar_cuadro_produccion_mes_actual(
                df_pr_detalle,
                "Magistrado_Prod",
                "Magistrado",
                f"TOTAL {periodo_lbl.upper()} ({'TODOS' if es_nacional else dist_sel})"
            )
            if not t_3_4.empty:
                mask_no_tot = ~t_3_4["Magistrado"].str.upper().str.startswith("TOTAL")
                nums_34 = [f"{i}. " for i in range(1, int(mask_no_tot.sum()) + 1)]
                t_3_4.loc[mask_no_tot, "Magistrado"] = nums_34 + t_3_4.loc[mask_no_tot, "Magistrado"].astype(str)

                mostrar_tabla_centrada(t_3_4)
                df_g_34 = t_3_4[mask_no_tot].iloc[::-1].copy()
                df_g_34["Magistrado"] = df_g_34["Magistrado"].str.replace(r"^\d+\.\s*", "", regex=True)
                if not df_g_34.empty:
                    max_34 = max(1, int(df_g_34["Total"].max()))
                    med_mag = float(t_3_4[~t_3_4["Magistrado"].str.upper().str.startswith("TOTAL")]["Total"].median())
                    fig_3_4 = go.Figure()
                    for _, r_m in df_g_34.iterrows():
                        fig_3_4.add_trace(go.Scatter(
                            x=[0, int(r_m["Total"])], y=[r_m["Magistrado"], r_m["Magistrado"]],
                            mode="lines", line=dict(color="#4f46e5", width=3), showlegend=False, hoverinfo="skip"
                        ))
                    custom_34 = np.stack([df_g_34["Resoluciones"], df_g_34["Informes"], df_g_34["% Part."]], axis=-1)
                    fig_3_4.add_trace(go.Scatter(
                        x=df_g_34["Total"], y=df_g_34["Magistrado"], mode="markers+text",
                        marker=dict(size=13, color="#818cf8", line=dict(color="#ffffff", width=1.8)),
                        text=[f"  <b>{int(r['Total']):,}</b> ({r['% Part.']:.1f}%)" for _, r in df_g_34.iterrows()],
                        textposition="middle right", textfont=dict(size=11, color="#f8fafc"), customdata=custom_34,
                        hovertemplate="<b>⚖️ Magistrado:</b> %{y}<br><b>📊 Total:</b> %{x:,} (%{customdata[2]:.2f}%)<br><b>📘 Resoluciones:</b> %{customdata[0]:,} | <b>📙 Informes:</b> %{customdata[1]:,}<extra></extra>"
                    ))
                    fig_3_4.add_vline(
                        x=med_mag, line_dash="dash", line_color="#f59e0b", line_width=1.8,
                        annotation_text=f"Mediana ({med_mag:.0f} exp.)", annotation_position="top right",
                        annotation_font=dict(color="#fbbf24", size=11)
                    )
                    fig_3_4.update_layout(
                        title=dict(text=f"Gráfico 3.4 — Ranking de Producción por Magistrado ({periodo_lbl})", font=dict(size=13.5, color="#f8fafc"), x=0.01),
                        template="plotly_dark", paper_bgcolor="#111827", plot_bgcolor="#0f172a",
                        xaxis=dict(title="Total de Pronunciamientos (Resoluciones + Informes)", range=[0, max_34 * 1.24], gridcolor="rgba(148, 163, 184, 0.14)"),
                        yaxis=dict(title="", tickfont=dict(size=11, color="#f8fafc")), showlegend=False,
                        height=max(320, len(df_g_34) * 32 + 90), margin=dict(l=15, r=30, t=48, b=40)
                    )
                    st.plotly_chart(fig_3_4, use_container_width=True)

            st.markdown("---")

            # ======================================================================
            # BLOQUE 3.5: ESTRUCTURA OPERATIVA MAGISTRADO -> ASISTENTES JUDICIALES
            # ======================================================================
            st.subheader(f"3.5. Estructura Operativa: Magistrado / Despacho y sus Asistentes — {periodo_lbl}")
            st.caption("Nota metodológica: Desglose jerárquico por Magistrado y usuario registrado en la descarga del proyecto (USUARIO DESCARGO PROYECTO).")

            t_3_5 = generar_cuadro_jerarquico_asistentes(
                df_pr_detalle,
                f"TOTAL {periodo_lbl.upper()} ({'TODOS' if es_nacional else dist_sel})"
            )
            if not t_3_5.empty:
                def resaltar_jerarquia(row):
                    txt = str(row["Magistrado / Despacho → Asistente Judicial"])
                    if txt.startswith("🏛️") or txt.startswith("TOTAL"):
                        return ["background-color: #1e293b; font-weight: bold; color: #f8fafc;"] * len(row)
                    return ["background-color: #0f172a; color: #cbd5e1;"] * len(row)

                cols_centrar_35 = ["Resoluciones", "Informes", "Total", "% Despacho", "% Total"]
                styler_35 = (
                    t_3_5.style
                    .hide(axis="index")
                    .apply(resaltar_jerarquia, axis=1)
                    .set_properties(subset=cols_centrar_35, **{"text-align": "center"})
                    .set_properties(subset=[t_3_5.columns[0]], **{"text-align": "left"})
                    .set_table_styles([
                        {"selector": "table", "props": [("width", "100%"), ("border-collapse", "collapse"), ("font-size", "13px")]},
                        {"selector": "thead th", "props": [
                            ("background-color", "#1e293b"), ("color", "#f8fafc"), ("font-weight", "600"),
                            ("padding", "10px 12px"), ("border-bottom", "2px solid #334155"),
                            ("position", "sticky"), ("top", "0"), ("z-index", "1")
                        ]},
                        {"selector": "thead th.col0", "props": [("text-align", "left")]},
                        {"selector": "thead th:not(.col0)", "props": [("text-align", "center")]},
                        {"selector": "td", "props": [("padding", "8px 12px"), ("border-bottom", "1px solid #1e293b")]},
                    ])
                )
                tabla_html_35 = styler_35.to_html(table_attributes='style="width: 100%; border-collapse: collapse;"')
                html_35 = f'<div style="width: 100%; max-height: 620px; overflow-y: auto; border: 1px solid #334155; border-radius: 8px;">{tabla_html_35}</div>'
                st.markdown(html_35, unsafe_allow_html=True)

                df_g_35 = df_pr_detalle.groupby(["Magistrado_Prod", "Asistente_Prod"]).size().reset_index(name="Cantidad")
                top_mags_35 = df_pr_detalle.groupby("Magistrado_Prod").size().sort_values(ascending=False).index.tolist()
                df_g_35 = df_g_35[df_g_35["Magistrado_Prod"].isin(top_mags_35)].copy()

                if not df_g_35.empty:
                    tot_global_35 = max(1, int(df_g_35["Cantidad"].sum()))
                    df_tot_desp = df_g_35.groupby("Magistrado_Prod", as_index=False)["Cantidad"].sum().rename(columns={"Cantidad": "Total_Despacho"})
                    df_tot_desp["Pct_Despacho"] = (df_tot_desp["Total_Despacho"] / tot_global_35 * 100.0).round(1)
                    max_x_35 = max(1, int(df_tot_desp["Total_Despacho"].max()))

                    fig_3_5 = px.bar(
                        df_g_35, x="Cantidad", y="Magistrado_Prod", color="Asistente_Prod", orientation="h",
                        category_orders={"Magistrado_Prod": top_mags_35}, text="Cantidad",
                        title=f"Gráfico 3.5 — Distribución de Producción por Asistente ({periodo_lbl})"
                    )
                    fig_3_5.update_traces(
                        textposition="inside", insidetextanchor="middle", textfont=dict(size=10.5, color="#ffffff"),
                        hovertemplate="<b>🏛️ Magistrado:</b> %{y}<br><b>🧑‍💼 Asistente:</b> %{fullData.name}<br><b>📊 Producción:</b> %{x:,}<extra></extra>"
                    )
                    fig_3_5.add_trace(go.Scatter(
                        x=df_tot_desp["Total_Despacho"],
                        y=df_tot_desp["Magistrado_Prod"],
                        mode="text",
                        text=[f"  <b>{int(r['Total_Despacho']):,}</b> ({r['Pct_Despacho']:.1f}%)" for _, r in df_tot_desp.iterrows()],
                        textposition="middle right",
                        textfont=dict(size=11, color="#f8fafc"),
                        showlegend=False,
                        hoverinfo="skip"
                    ))
                    fig_3_5.update_layout(
                        template="plotly_dark", paper_bgcolor="#111827", plot_bgcolor="#0f172a",
                        barmode="stack", showlegend=False,
                        xaxis=dict(title="Total de Producción", range=[0, max_x_35 * 1.18], gridcolor="rgba(148, 163, 184, 0.14)"),
                        yaxis=dict(title="", tickfont=dict(size=11, color="#f8fafc")),
                        height=max(340, len(top_mags_35) * 36 + 90), margin=dict(l=15, r=30, t=48, b=35)
                    )
                    st.plotly_chart(fig_3_5, use_container_width=True)
                

    # ==========================================================================
    # 11. EXPORTACIÓN DE INFORME EJECUTIVO GLOBAL (EXCEL Y PDF EN BARRA LATERAL)
    # ==========================================================================
    if mod_actual != "CARGA":
        orden_anios_cuadro = ["2021 y ant.", "2022", "2023", "2024", "2025", "2026"]
        tot_carga_cuadro = max(1, len(df_act_sede))
        if es_nacional:
            filas_1_1 = []
            for a_lbl in orden_anios_cuadro:
                sub_a = df_act_sede[df_act_sede["Anio_Grupo"] == a_lbl]
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
                    "% Part.": round(tot_a / tot_carga_cuadro * 100, 2) if len(df_act_sede) > 0 else 0.0
                })
            t_1_1_anio = pd.DataFrame(filas_1_1).sort_values(by=["Total Nacional", "Año de Creación"], ascending=[False, False]).reset_index(drop=True)
            t_1_1_anio.insert(0, "N°", [str(i) for i in range(1, len(t_1_1_anio) + 1)])
            t_1_1_anio = pd.concat([t_1_1_anio, pd.DataFrame([{
                "N°": "", "Año de Creación": "TOTAL CARGA PROCESAL",
                "ANC - En Trámite": int(t_1_1_anio["ANC - En Trámite"].sum()),
                "ANC - Res. Impug.": int(t_1_1_anio["ANC - Res. Impug."].sum()),
                "Total ANC": int(t_1_1_anio["Total ANC"].sum()),
                "ODANC - En Trámite": int(t_1_1_anio["ODANC - En Trámite"].sum()),
                "ODANC - Res. Impug.": int(t_1_1_anio["ODANC - Res. Impug."].sum()),
                "Total ODANC": int(t_1_1_anio["Total ODANC"].sum()),
                "Total Nacional": len(df_act_sede), "% Part.": 100.0 if len(df_act_sede) > 0 else 0.0
            }])], ignore_index=True)
        else:
            filas_1_1 = []
            for a_lbl in orden_anios_cuadro:
                sub_a = df_act_sede[df_act_sede["Anio_Grupo"] == a_lbl]
                tot_a = len(sub_a)
                filas_1_1.append({
                    "Año de Creación": a_lbl,
                    "En Trámite": int((sub_a["Estado_Grupo"] == "Trámite").sum()),
                    "Resuelto - Impug. Tránsito": int((sub_a["Estado_Grupo"] == "Res. Impug.").sum()),
                    "Total Expedientes": tot_a,
                    "% Part.": round(tot_a / tot_carga_cuadro * 100, 2) if len(df_act_sede) > 0 else 0.0
                })
            t_1_1_anio = pd.DataFrame(filas_1_1).sort_values(by=["Total Expedientes", "Año de Creación"], ascending=[False, False]).reset_index(drop=True)
            t_1_1_anio.insert(0, "N°", [str(i) for i in range(1, len(t_1_1_anio) + 1)])
            t_1_1_anio = pd.concat([t_1_1_anio, pd.DataFrame([{
                "N°": "", "Año de Creación": f"TOTAL CARGA {dist_sel}",
                "En Trámite": int(t_1_1_anio["En Trámite"].sum()),
                "Resuelto - Impug. Tránsito": int(t_1_1_anio["Resuelto - Impug. Tránsito"].sum()),
                "Total Expedientes": len(df_act_sede), "% Part.": 100.0 if len(df_act_sede) > 0 else 0.0
            }])], ignore_index=True)

        t_1_2_etapas, _ = construir_df_etapas_anio(df_act_sede)
        t_1_3_cal_mes, _ = construir_df_mes_anio_estado(df_act_sede[df_act_sede["Etapa_Cat"] == "Calificación"], "TOTAL CALIFICACIÓN")
        t_1_2_unidad, _ = construir_df_dimension_anio(df_act_sede, "Despacho o Unidad", "Unidad Orgánica", es_magistrado=False)
        t_1_2_magistrado, _ = construir_df_dimension_anio(df_act_sede, "Magistrado", "Magistrado de Control", es_magistrado=True)
        df_act = df_act_sede

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
        import tempfile
        import textwrap
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt

        plt.rcParams["font.sans-serif"] = "DejaVu Sans"
        plt.rcParams["axes.edgecolor"] = "#CBD5E1"
        plt.rcParams["axes.linewidth"] = 0.8

        class PDFEjecutivoANC(FPDF):
            def header(self):
                if self.page_no() > 1:
                    self.set_fill_color(127, 29, 29)
                    self.rect(10, 7, 277, 8, "F")
                    self.set_fill_color(245, 158, 11)
                    self.rect(10, 7, 2.5, 8, "F")
                    self.set_xy(15, 8.2)
                    self.set_font("Helvetica", "B", 7.8)
                    self.set_text_color(255, 255, 255)
                    self.cell(185, 5.5, limpiar_latin1(f"ANEXO ESTADÍSTICO OFICIAL - {nombre_ambito}"), align="L")
                    self.set_font("Helvetica", "", 7.5)
                    self.set_text_color(254, 243, 199)
                    self.cell(85, 5.5, limpiar_latin1(f"Corte Carga: {FECHA_CORTE_CARGA.strftime('%d/%m/%Y')} | Presc.: {FECHA_CORTE_PRESC.strftime('%d/%m/%Y')}"), align="R")
                    self.set_xy(10, 18)

            def footer(self):
                apb = self.auto_page_break
                b_m = self.b_margin
                self.set_auto_page_break(False)
                self.set_y(-10)
                self.set_draw_color(203, 213, 225)
                self.line(10, self.get_y(), 287, self.get_y())
                self.set_font("Helvetica", "I", 7.2)
                self.set_text_color(100, 116, 139)
                self.cell(195, 6, limpiar_latin1("Autoridad Nacional de Control del Poder Judicial | Unidad de Investigación y Análisis (UIA)"), align="L")
                self.set_font("Helvetica", "B", 7.2)
                self.cell(82, 6, limpiar_latin1(f"Página {self.page_no()}"), align="R")
                self.set_auto_page_break(apb, margin=b_m)

        pdf = PDFEjecutivoANC(orientation="L", unit="mm", format="A4")
        pdf.set_auto_page_break(auto=True, margin=14)
        pdf.add_page()

        # --- ENCABEZADO INSTITUCIONAL DE PORTADA ---
        pdf.set_fill_color(127, 29, 29)
        pdf.rect(10, 10, 277, 20, "F")
        pdf.set_fill_color(245, 158, 11)
        pdf.rect(10, 10, 3.5, 20, "F")

        pdf.set_xy(16, 12)
        pdf.set_font("Helvetica", "B", 13.5)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(0, 6.5, limpiar_latin1("AUTORIDAD NACIONAL DE CONTROL DEL PODER JUDICIAL"), ln=True)

        pdf.set_x(16)
        pdf.set_font("Helvetica", "B", 9.5)
        pdf.set_text_color(253, 224, 71)
        pdf.cell(0, 5, limpiar_latin1(f"ANEXO - RESUMEN DE INFORMACIÓN ESTADÍSTICA: {nombre_ambito}"), ln=True)

        pdf.set_x(16)
        pdf.set_font("Helvetica", "", 8)
        pdf.set_text_color(226, 232, 240)
        pdf.cell(0, 4.5, limpiar_latin1(f"Corte Carga Procesal: {FECHA_CORTE_CARGA.strftime('%d/%m/%Y')}   |   Corte Riesgo Prescripción: {FECHA_CORTE_PRESC.strftime('%d/%m/%Y')}   |   Elaborado por: UIA"), ln=True)
        pdf.ln(5)

        # --- TARJETAS KPI DE IMPACTO (6 TARJETAS CON DESGLOSE COMPLETO Y EXP. MÁS ANTIGUO) ---
        tot_carga_pdf = len(df_act_sede)
        tot_tram_pdf = int((df_act_sede["Estado_Grupo"] == "Trámite").sum())
        tot_res_pdf = int((df_act_sede["Estado_Grupo"] == "Res. Impug.").sum())
        df_jef_pdf = df_act_sede[df_act_sede["Despacho_Norm"].str.contains("JEFATURA", na=False)].copy()
        tot_jef_pdf = len(df_jef_pdf)
        tot_presc_pdf = len(df_p_act)
        cnt_p0 = int((df_p_act["Dias_Calc"] <= 0).sum())
        cnt_p1_90 = int(((df_p_act["Dias_Calc"] >= 1) & (df_p_act["Dias_Calc"] <= 90)).sum())
        cnt_p91_240 = int(((df_p_act["Dias_Calc"] >= 91) & (df_p_act["Dias_Calc"] <= 240)).sum())
        tot_prod_mes_pdf = len(df_pr_mes_act)

        # Cálculo dinámico de Tarjeta 5: SIN INVENTARIAR (desde Nota de Inventario en df_act_sede)
        col_nota_inv = next((c for c in df_act_sede.columns if "NOTA" in str(c).upper() and "INVENTARIO" in str(c).upper()), None)
        if col_nota_inv is not None:
            df_ni_calc = df_act_sede[df_act_sede[col_nota_inv].fillna("").astype(str).str.upper().str.contains("SIN INVENTARIAR", na=False)]
            tot_no_inv_pdf = len(df_ni_calc)
            tram_no_inv_pdf = int((df_ni_calc["Estado_Grupo"] == "Trámite").sum())
            res_no_inv_pdf = int((df_ni_calc["Estado_Grupo"] == "Res. Impug.").sum())
        else:
            tot_no_inv_pdf = len(df_no_inv)
            if not df_no_inv.empty and "Estado_Grupo" in df_no_inv.columns:
                tram_no_inv_pdf = int((df_no_inv["Estado_Grupo"] == "Trámite").sum())
                res_no_inv_pdf = int((df_no_inv["Estado_Grupo"] == "Res. Impug.").sum())
            else:
                col_est_ni = next((c for c in df_no_inv.columns if "ESTADO" in str(c).upper()), None) if not df_no_inv.empty else None
                if col_est_ni is not None:
                    est_no_inv_u = df_no_inv[col_est_ni].fillna("").astype(str).str.upper()
                    res_no_inv_pdf = int(est_no_inv_u.str.contains("RESUELTO|IMPUG", na=False).sum())
                    tram_no_inv_pdf = tot_no_inv_pdf - res_no_inv_pdf
                else:
                    tram_no_inv_pdf = tot_no_inv_pdf
                    res_no_inv_pdf = 0

        # Cálculo dinámico de Tarjeta 6: EXP. MÁS ANTIGUO
        val_antiguo_pdf = "S/D"
        sub_antiguo_pdf = "Sin registro de fecha"
        col_fecha_ant = next((c for c in ["Fecha_Creacion_Dt", "Fecha Creación", "Fecha_Creacion", "Fecha Ingreso"] if c in df_act_sede.columns), None)
        if col_fecha_ant is not None and not df_act_sede.empty:
            s_fechas_ant = pd.to_datetime(df_act_sede[col_fecha_ant], dayfirst=True, errors="coerce")
            if s_fechas_ant.notna().any():
                idx_ant = s_fechas_ant.idxmin()
                f_ant = s_fechas_ant.loc[idx_ant]
                row_ant = df_act_sede.loc[idx_ant]
                val_antiguo_pdf = f"{MESES_ABR[int(f_ant.month)]} {int(f_ant.year)}"
                desp_ant = str(row_ant.get("Despacho_Norm", row_ant.get("Despacho o Unidad", ""))).strip()[:16]
                et_ant = ETAPAS_CORTAS.get(str(row_ant.get("Etapa_Cat", "")), str(row_ant.get("Etapa_Cat", ""))[:10])
                est_ant = str(row_ant.get("Estado_Grupo", "")).strip()
                sub_antiguo_pdf = f"{desp_ant} | {et_ant} | {est_ant}"

        kpis_cards = [
            ("1. CARGA TOTAL", f"{tot_carga_pdf:,} exp.", f"Trám: {tot_tram_pdf:,} | Res.Imp: {tot_res_pdf:,}", (30, 58, 138), (239, 246, 255)),
            ("2. RIESGO PRESC. (<=240d)", f"{tot_presc_pdf:,} exp.", f"Presc: {cnt_p0:,} | 1-90d: {cnt_p1_90:,} | 91-240d: {cnt_p91_240:,}", (153, 27, 27), (254, 242, 242)),
            ("3. EN JEFATURA", f"{tot_jef_pdf:,} exp.", f"{(tot_jef_pdf/max(1,tot_carga_pdf)*100):.2f}% de la carga total", (146, 64, 14), (255, 251, 235)),
            (f"4. PROD. {mes_act_lbl.upper()[:11]}", f"{tot_prod_mes_pdf:,} exp.", f"Res: {int((~df_pr_mes_act['Es_Informe']).sum()):,} | Inf: {int(df_pr_mes_act['Es_Informe'].sum()):,}", (6, 95, 70), (236, 253, 245)),
            ("5. SIN INVENTARIAR", f"{tot_no_inv_pdf:,} exp.", f"Trám: {tram_no_inv_pdf:,} | Res.Imp: {res_no_inv_pdf:,}", (71, 85, 105), (248, 250, 252)),
            ("6. EXP. MÁS ANTIGUO", val_antiguo_pdf, sub_antiguo_pdf, (217, 119, 6), (255, 251, 235)),
        ]
        y_kpi = pdf.get_y()
        n_cards = len(kpis_cards)
        gap_card = 2.0
        w_card = (277.0 - gap_card * (n_cards - 1)) / n_cards
        for idx_c, (tit_c, val_c, sub_c, col_borde, col_bg) in enumerate(kpis_cards):
            x_c = 10 + idx_c * (w_card + gap_card)
            pdf.set_fill_color(*col_bg)
            pdf.set_draw_color(*col_borde)
            pdf.rect(x_c, y_kpi, w_card, 17, "DF")
            pdf.set_fill_color(*col_borde)
            pdf.rect(x_c, y_kpi, 1.8, 17, "F")
            pdf.set_xy(x_c + 2.6, y_kpi + 1.5)
            pdf.set_font("Helvetica", "B", 6.2)
            pdf.set_text_color(*col_borde)
            pdf.cell(w_card - 3.2, 4, limpiar_latin1(tit_c[:24]), ln=True)
            pdf.set_x(x_c + 2.6)
            pdf.set_font("Helvetica", "B", 10.5)
            pdf.set_text_color(15, 23, 42)
            pdf.cell(w_card - 3.2, 5.5, limpiar_latin1(val_c), ln=True)
            pdf.set_x(x_c + 2.6)
            pdf.set_font("Helvetica", "", 5.6)
            pdf.set_text_color(71, 85, 105)
            pdf.cell(w_card - 3.2, 4, limpiar_latin1(sub_c[:36]), ln=True)
        pdf.set_y(y_kpi + 20)

        def dibujar_banner_seccion(texto_sec):
            if pdf.get_y() > 175:
                pdf.add_page()
            pdf.set_fill_color(30, 41, 59)
            pdf.rect(10, pdf.get_y(), 277, 7, "F")
            pdf.set_fill_color(245, 158, 11)
            pdf.rect(10, pdf.get_y(), 3, 7, "F")
            pdf.set_xy(15, pdf.get_y() + 1)
            pdf.set_font("Helvetica", "B", 9.5)
            pdf.set_text_color(255, 255, 255)
            pdf.cell(0, 5, limpiar_latin1(texto_sec), ln=True)
            pdf.ln(2.5)

        MAPA_VINETA_PRESC = {
            "PRESCRITOS (<= 0": (239, 68, 68),
            "CRITICO: DE 1 A 15": (250, 204, 21),
            "ALTO: DE 16 A 30": (101, 163, 13),
            "MODERADO: DE 31 A 60": (37, 99, 235),
            "PREVENTIVO: DE 61 A 90": (6, 182, 212),
            "MONITOREO: DE 91 A 240": (139, 92, 246)
        }

        def dibujar_tabla_ejecutiva(titulo, df_tab, max_filas=50, con_vinetas_presc=False):
            if df_tab is None or df_tab.empty:
                return
            df_show = pd.concat([df_tab.head(max_filas - 1), df_tab.tail(1)], ignore_index=True) if len(df_tab) > max_filas else df_tab.copy()
            cols = [c for c in df_show.columns if not str(c).startswith("_")]
            n_cols = len(cols)

            if pdf.get_y() > 165:
                pdf.add_page()

            pdf.set_font("Helvetica", "B", 9)
            pdf.set_text_color(127, 29, 29)
            pdf.cell(0, 5.5, limpiar_latin1(titulo), ln=True)

            if cols[0] == "N°":
                w_num = 8.5
                w_nom = 54.0 if n_cols > 8 else 112.0
                w_resto = (277.0 - w_num - w_nom) / max(1, (n_cols - 2))
                anchos = [w_num, w_nom] + [w_resto] * (n_cols - 2)
            else:
                w_primera = 96.0 if n_cols > 3 else 155.0
                w_resto = (277.0 - w_primera) / max(1, (n_cols - 1))
                anchos = [w_primera] + [w_resto] * (n_cols - 1)

            font_h = 5.8 if n_cols >= 12 else (6.6 if n_cols >= 8 else 7.5)
            pdf.set_font("Helvetica", "B", font_h)
            pdf.set_fill_color(30, 41, 59)
            pdf.set_text_color(255, 255, 255)
            pdf.set_draw_color(148, 163, 184)

            for i, c in enumerate(cols):
                etiq = str(c)
                etiq = (
                    etiq.replace(" - Trámite", " Tr.")
                    .replace(" - Res. Impug.", " Res.")
                    .replace("Total - Trámite", "Tot.Tr")
                    .replace("Total - Res. Impug.", "Tot.Rs")
                    .replace("Resuelto - Impug. Tránsito", "Resuelto - Impug.")
                    .replace("Rango de Riesgo de Expedientes Prescritos y por Prescribir", "Rango de Riesgo Presc. / Por Presc.")
                    .replace("Prescritos (<=0d)", "Presc(<=0d)")
                    .replace("Crítico (1-15d)", "Crít(1-15d)")
                    .replace("Alto (16-30d)", "Alto(16-30d)")
                    .replace("Moderado (31-60d)", "Mod(31-60d)")
                    .replace("Preventivo (61-90d)", "Prev(61-90d)")
                    .replace("Monitoreo (91-240d)", "Mon(91-240d)")
                )
                max_ch = max(4, int(anchos[i] / (font_h * 0.24)))
                align_h = "C" if (i == 0 and cols[0] == "N°") else ("L" if (i == 0 or (i == 1 and cols[0] == "N°")) else "C")
                pdf.cell(anchos[i], 6.6, limpiar_latin1(etiq[:max_ch]), border=1, align=align_h, fill=True)
            pdf.ln()

            font_b = 6.2 if n_cols >= 12 else (6.8 if n_cols >= 8 else 7.4)
            for idx, row in df_show.iterrows():
                texto_check = " ".join([str(row[c]) for c in cols[:2]])
                es_tot = any(p in texto_check.upper() for p in ["TOTAL", "SUBTOTAL"])
                es_sub = str(row[cols[1] if cols[0] == "N°" else cols[0]]).startswith("   •")

                pdf.set_font("Helvetica", "B" if es_tot else ("I" if es_sub else ""), font_b)
                if es_tot:
                    pdf.set_fill_color(226, 232, 240)
                    pdf.set_text_color(15, 23, 42)
                elif es_sub:
                    pdf.set_fill_color(241, 245, 249)
                    pdf.set_text_color(51, 65, 85)
                else:
                    pdf.set_fill_color(255 if idx % 2 == 0 else 248, 255 if idx % 2 == 0 else 250, 255 if idx % 2 == 0 else 252)
                    pdf.set_text_color(30, 41, 59)

                for i, c in enumerate(cols):
                    val = row[c]
                    if isinstance(val, (int, float, np.integer, np.floating)) and pd.notna(val):
                        if "%" in str(c):
                            txt = f"{val:,.2f}%"
                        else:
                            txt = "-" if (int(val) == 0 and n_cols >= 10 and not es_tot) else f"{int(val):,}"
                    else:
                        txt = str(val) if pd.notna(val) else ""

                    color_vineta = None
                    if i == 0 and con_vinetas_presc and not es_tot:
                        txt_norm_v = norm_txt(txt)
                        for clave_v, rgb_v in MAPA_VINETA_PRESC.items():
                            if clave_v in txt_norm_v:
                                color_vineta = rgb_v
                                break
                        for em in ["🔴", "🟡", "🟢", "🔵", "🟣"]:
                            txt = txt.replace(em, "").strip()

                    max_cb = max(3, int(anchos[i] / (font_b * 0.23)))
                    align_b = "C" if (i == 0 and cols[0] == "N°") else ("L" if (i == 0 or (i == 1 and cols[0] == "N°")) else "C")

                    if color_vineta is not None:
                        x_cell = pdf.get_x()
                        y_cell = pdf.get_y()
                        pdf.cell(anchos[i], 5.5, "", border=1, fill=True)
                        pdf.set_fill_color(*color_vineta)
                        pdf.set_draw_color(100, 116, 139)
                        pdf.ellipse(x_cell + 2.2, y_cell + 1.35, 2.8, 2.8, "DF")
                        pdf.set_fill_color(255 if idx % 2 == 0 else 248, 255 if idx % 2 == 0 else 250, 255 if idx % 2 == 0 else 252)
                        pdf.set_xy(x_cell + 6.2, y_cell)
                        pdf.cell(anchos[i] - 6.2, 5.5, limpiar_latin1(txt[:max_cb]), border=0, align="L")
                        pdf.set_xy(x_cell + anchos[i], y_cell)
                        pdf.set_draw_color(148, 163, 184)
                    else:
                        pdf.cell(anchos[i], 5.5, limpiar_latin1(txt[:max_cb]), border=1, align=align_b, fill=True)
                pdf.ln()
            pdf.ln(3)

        def incrustar_figura(fig_mpl, ancho_mm=277, alto_mm=64):
            if pdf.get_y() + alto_mm > 192:
                pdf.add_page()
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
                tmp_path = tmp.name
            fig_mpl.savefig(tmp_path, dpi=220, pad_inches=0.08, facecolor="white")
            plt.close(fig_mpl)
            pdf.image(tmp_path, x=10, y=pdf.get_y(), w=ancho_mm, h=alto_mm)
            pdf.set_y(pdf.get_y() + alto_mm + 3)
            try:
                os.remove(tmp_path)
            except Exception:
                pass

        # ======================================================================
        # SECCIÓN 1: CARGA PROCESAL (NUMERACIÓN LIMPIA 1.1, 1.2, 1.3, 1.4, 1.5)
        # ======================================================================
        dibujar_banner_seccion(f"1. ANÁLISIS DE CARGA PROCESAL: {tot_carga_pdf:,} EXPEDIENTES EN TOTAL ({sub_titulo_sec})")
        dibujar_tabla_ejecutiva(f"1.1. Carga Total por Año de Creación y Estado ({sub_titulo_sec})", t_1_1_anio)

        fig1, (ax1, ax2) = plt.subplots(1, 2, figsize=(13.5, 3.2), dpi=220)
        fig1.patch.set_facecolor("white")

        orden_anios_g = ["2021 y ant.", "2022", "2023", "2024", "2025", "2026"]
        tr_anios = [int(((df_act_sede["Anio_Grupo"] == a) & (df_act_sede["Estado_Grupo"] == "Trámite")).sum()) for a in orden_anios_g]
        re_anios = [int(((df_act_sede["Anio_Grupo"] == a) & (df_act_sede["Estado_Grupo"] == "Res. Impug.")).sum()) for a in orden_anios_g]
        tot_anios = [t + r for t, r in zip(tr_anios, re_anios)]
        x_pos = np.arange(len(orden_anios_g))
        ax1.bar(x_pos, tr_anios, width=0.55, label="En Trámite", color="#1E3A8A")
        ax1.bar(x_pos, re_anios, width=0.55, bottom=tr_anios, label="Resuelto - Impug.", color="#F59E0B")
        max_y1 = max(1, max(tot_anios))
        ax1.set_ylim(0, max_y1 * 1.24)
        for i, v in enumerate(tot_anios):
            if v > 0:
                pct_v = v / max(1, tot_carga_pdf) * 100
                ax1.text(i, v + max_y1 * 0.025, f"{v:,} ({pct_v:.1f}%)", ha="center", va="bottom", fontsize=7.3, fontweight="bold", color="#0F172A")
        ax1.set_xticks(x_pos)
        ax1.set_xticklabels(orden_anios_g, fontsize=8, fontweight="bold")
        ax1.set_title("Gráfico 1.1 — Carga por Año de Creación y Estado", fontsize=9.5, fontweight="bold", color="#7F1D1D", pad=8)
        ax1.grid(axis="y", linestyle="--", alpha=0.35)
        ax1.spines["top"].set_visible(False)
        ax1.spines["right"].set_visible(False)
        ax1.legend(loc="upper left", fontsize=7.5, frameon=True)

        et_cortas_g = [ETAPAS_CORTAS[e] for e in ORDEN_ETAPAS]
        tr_et_g = [int(((df_act_sede["Etapa_Cat"] == e) & (df_act_sede["Estado_Grupo"] == "Trámite")).sum()) for e in ORDEN_ETAPAS]
        re_et_g = [int(((df_act_sede["Etapa_Cat"] == e) & (df_act_sede["Estado_Grupo"] == "Res. Impug.")).sum()) for e in ORDEN_ETAPAS]
        tot_et_g = [t + r for t, r in zip(tr_et_g, re_et_g)]
        y_pos2 = np.arange(len(et_cortas_g))[::-1]
        ax2.barh(y_pos2, tr_et_g, height=0.55, label="En Trámite", color="#0D9488")
        ax2.barh(y_pos2, re_et_g, height=0.55, left=tr_et_g, label="Resuelto - Impug.", color="#E11D48")
        max_x2 = max(1, max(tot_et_g))
        ax2.set_xlim(0, max_x2 * 1.28)
        for idx_y, v in zip(y_pos2, tot_et_g):
            pct_e = v / max(1, tot_carga_pdf) * 100
            ax2.text(v + max_x2 * 0.02, idx_y, f"{v:,} ({pct_e:.1f}%)", va="center", ha="left", fontsize=8, fontweight="bold", color="#0F172A")
        ax2.set_yticks(y_pos2)
        ax2.set_yticklabels(et_cortas_g, fontsize=8.2, fontweight="bold")
        ax2.set_title("Gráfico 1.2 — Distribución Procesal por Etapa", fontsize=9.5, fontweight="bold", color="#7F1D1D", pad=8)
        ax2.grid(axis="x", linestyle="--", alpha=0.35)
        ax2.spines["top"].set_visible(False)
        ax2.spines["right"].set_visible(False)
        ax2.legend(loc="lower right", fontsize=7.5, frameon=True)

        fig1.subplots_adjust(left=0.06, right=0.98, top=0.86, bottom=0.14, wspace=0.22)
        incrustar_figura(fig1, ancho_mm=277, alto_mm=63)

        # 1.2. Carga Procesal por Etapa y Sub-etapas de Calificación
        dibujar_tabla_ejecutiva(f"1.2. Carga Procesal por Etapa y Sub-etapas de Calificación ({sub_titulo_sec})", t_1_2_etapas)

        # 1.3. Sub-desglose de Calificación según Ubicación / Unidad Orgánica (UCIP / UPAD / USAP)
        df_cal_pdf = df_act_sede[df_act_sede["Etapa_Cat"] == "Calificación"].copy()
        if not df_cal_pdf.empty:
            df_cal_pdf["Ubicacion_Cal"] = df_cal_pdf.apply(
                lambda r: obtener_nombre_unidad_por_sigla(r.get("Despacho o Unidad", ""), bool(r.get("Es_Sede_Central", False))),
                axis=1
            )
            g_cal_u = df_cal_pdf.groupby("Ubicacion_Cal").agg(
                En_Tramite=("Estado_Grupo", lambda s: int((s == "Trámite").sum())),
                Res_Impug=("Estado_Grupo", lambda s: int((s == "Res. Impug.").sum())),
                Total=("Estado_Grupo", "size")
            ).sort_values("Total", ascending=False).reset_index()
            g_cal_u["% en Calificación"] = (g_cal_u["Total"] / max(1, len(df_cal_pdf)) * 100).round(2)
            g_cal_u = g_cal_u.rename(columns={
                "Ubicacion_Cal": "Ubicación / Unidad en Etapa de Calificación (UCIP / UPAD / USAP)",
                "En_Tramite": "En Trámite",
                "Res_Impug": "Resuelto - Impug."
            })
            fila_tot_cal_u = pd.DataFrame([{
                "Ubicación / Unidad en Etapa de Calificación (UCIP / UPAD / USAP)": "TOTAL EN ETAPA DE CALIFICACIÓN POR UBICACIÓN",
                "En Trámite": int(g_cal_u["En Trámite"].sum()),
                "Resuelto - Impug.": int(g_cal_u["Resuelto - Impug."].sum()),
                "Total": len(df_cal_pdf),
                "% en Calificación": 100.0
            }])
            t_1_2_cal_ubic = pd.concat([g_cal_u, fila_tot_cal_u], ignore_index=True)
            dibujar_tabla_ejecutiva(f"1.3. Sub-desglose de Calificación según Ubicación / Unidad Orgánica ({sub_titulo_sec})", t_1_2_cal_ubic)

        # 1.4. Carga en Calificación por Mes de Creación, Año y Estado
        dibujar_tabla_ejecutiva(f"1.4. Carga en Calificación por Mes de Creación, Año y Estado ({sub_titulo_sec})", t_1_3_cal_mes)

        # 1.5. En Jefatura ODANC / Sede Central por Etapa
        if not df_jef_pdf.empty:
            filas_jef = []
            for idx_j, et_j in enumerate(ORDEN_ETAPAS, start=1):
                sub_ej = df_jef_pdf[df_jef_pdf["Etapa_Cat"] == et_j]
                tr_j = int((sub_ej["Estado_Grupo"] == "Trámite").sum())
                re_j = int((sub_ej["Estado_Grupo"] == "Res. Impug.").sum())
                tot_j = len(sub_ej)
                filas_jef.append({
                    "N°": str(idx_j),
                    "Etapa Procesal en Jefatura": et_j,
                    "En Trámite": tr_j,
                    "Resuelto - Impug.": re_j,
                    "Total en Jefatura": tot_j,
                    "% en Jefatura": round(tot_j / max(1, tot_jef_pdf) * 100, 2)
                })
            filas_jef.append({
                "N°": "",
                "Etapa Procesal en Jefatura": f"TOTAL CARGA EN JEFATURA ({'ANC / ODANCs' if es_nacional else dist_sel})",
                "En Trámite": int((df_jef_pdf["Estado_Grupo"] == "Trámite").sum()),
                "Resuelto - Impug.": int((df_jef_pdf["Estado_Grupo"] == "Res. Impug.").sum()),
                "Total en Jefatura": tot_jef_pdf,
                "% en Jefatura": 100.0
            })
            t_1_3_jef = pd.DataFrame(filas_jef)
        else:
            t_1_3_jef = pd.DataFrame([{
                "N°": "-",
                "Etapa Procesal en Jefatura": f"NO REGISTRA EXPEDIENTES EN JEFATURA ({dist_sel})",
                "En Trámite": 0, "Resuelto - Impug.": 0, "Total en Jefatura": 0, "% en Jefatura": 0.0
            }])
        dibujar_tabla_ejecutiva(f"1.5. Carga Procesal en Jefatura por Etapa ({sub_titulo_sec})", t_1_3_jef)

        # ======================================================================
        # SECCIÓN 2: RIESGO DE PRESCRIPCIÓN DETALLADO (CON VIÑETAS DE COLOR)
        # ======================================================================
        dibujar_banner_seccion(f"2. ANÁLISIS DE RIESGO DE PRESCRIPCIÓN (<= 240 DÍAS): {tot_presc_pdf:,} EXPEDIENTES")
        dibujar_tabla_ejecutiva(
            f"2.1. Rango de Riesgo de Expedientes Prescritos y por Prescribir ({sub_titulo_sec})",
            t_2_1_presc,
            con_vinetas_presc=True
        )
        dibujar_tabla_ejecutiva(f"2.2. Matriz de Riesgo de Prescripción según Etapa Procesal ({sub_titulo_sec})", t_2_2_et)

        if not df_p_act.empty:
            fig2, (axp1, axp2) = plt.subplots(1, 2, figsize=(13.5, 3.3), dpi=220)
            fig2.patch.set_facecolor("white")

            etiq_sem_p = ["Presc.\n(<=0d)", "Crítico\n(1-15d)", "Alto\n(16-30d)", "Mod.\n(31-60d)", "Prev.\n(61-90d)", "Monit.\n(91-240d)"]
            leyenda_sem_p = [
                "Prescritos (<= 0 días)",
                "Crítico: De 1 a 15 días",
                "Alto: De 16 a 30 días",
                "Moderado: De 31 a 60 días",
                "Preventivo: De 61 a 90 días",
                "Monitoreo: De 91 a 240 días"
            ]
            vals_sem_p = [int((df_p_act["Rango_Presc"] == rg).sum()) for rg in ORDEN_RANGOS_PRESC]
            cols_sem_p = ["#EF4444", "#FACC15", "#65A30D", "#2563EB", "#06B6D4", "#8B5CF6"]
            xp = np.arange(len(etiq_sem_p))
            axp1.bar(xp, vals_sem_p, width=0.52, color=cols_sem_p, edgecolor="#334155", linewidth=0.6)
            max_yp1 = max(1, max(vals_sem_p))
            axp1.set_ylim(0, max_yp1 * 1.28)
            for i, v in enumerate(vals_sem_p):
                axp1.text(i, v + max_yp1 * 0.03, f"{v:,}\n({v/max(1,tot_presc_pdf)*100:.1f}%)", ha="center", va="bottom", fontsize=7.2, fontweight="bold", color="#0F172A")
            axp1.set_xticks(xp)
            axp1.set_xticklabels(etiq_sem_p, fontsize=7.2, fontweight="bold")
            axp1.set_title("Gráfico 2.1 — Expedientes por Nivel de Riesgo de Prescripción", fontsize=9.2, fontweight="bold", color="#7F1D1D", pad=8)
            axp1.grid(axis="y", linestyle="--", alpha=0.35)
            axp1.spines["top"].set_visible(False)
            axp1.spines["right"].set_visible(False)

            handles_vinetas = [
                plt.Line2D([0], [0], marker="o", color="w", label=lbl_s, markerfacecolor=col_s, markeredgecolor="#475569", markersize=7.5)
                for lbl_s, col_s in zip(leyenda_sem_p, cols_sem_p)
            ]
            axp1.legend(handles=handles_vinetas, loc="upper left", fontsize=6.4, frameon=True, facecolor="#F8FAFC", edgecolor="#CBD5E1", title="Semáforo de Riesgo", title_fontsize=6.8)

            vals_et_p = [int((df_p_act["Etapa_Presc_Cat"] == e).sum()) for e in ORDEN_ETAPAS]
            yp2 = np.arange(len(ORDEN_ETAPAS))[::-1]
            axp2.barh(yp2, vals_et_p, height=0.55, color="#B91C1C")
            max_xp2 = max(1, max(vals_et_p))
            axp2.set_xlim(0, max_xp2 * 1.28)
            for idx_y, v in zip(yp2, vals_et_p):
                axp2.text(v + max_xp2 * 0.02, idx_y, f"{v:,} ({v/max(1,tot_presc_pdf)*100:.1f}%)", va="center", ha="left", fontsize=8, fontweight="bold", color="#0F172A")
            axp2.set_yticks(yp2)
            axp2.set_yticklabels([ETAPAS_CORTAS[e] for e in ORDEN_ETAPAS], fontsize=8.2, fontweight="bold")
            axp2.set_title("Gráfico 2.2 — Concentración de Riesgo por Etapa Procesal", fontsize=9.2, fontweight="bold", color="#7F1D1D", pad=8)
            axp2.grid(axis="x", linestyle="--", alpha=0.35)
            axp2.spines["top"].set_visible(False)
            axp2.spines["right"].set_visible(False)

            fig2.subplots_adjust(left=0.06, right=0.98, top=0.86, bottom=0.16, wspace=0.22)
            incrustar_figura(fig2, ancho_mm=277, alto_mm=64)

        # ======================================================================
        # SECCIÓN 3: PRODUCCIÓN Y PRONUNCIAMIENTOS
        # ======================================================================
        dibujar_banner_seccion(f"3. PRODUCCIÓN Y PRONUNCIAMIENTOS ({sub_titulo_sec})")
        if not t_3_1.empty:
            dibujar_tabla_ejecutiva(f"3.1. Comparativo Mensual de Producción: Mes Anterior vs. Mes Actual ({sub_titulo_sec})", t_3_1)
        if not t_3_2.empty:
            dibujar_tabla_ejecutiva(f"3.2. Producción por Tipo de Pronunciamiento y/o Sanción - {mes_act_lbl}", t_3_2)

        if not df_pr_f.empty and not t_3_3.empty:
            fig3, (axr1, axr2) = plt.subplots(
                1, 2,
                figsize=(14.5, 3.6),
                dpi=220,
                gridspec_kw={"width_ratios": [1.0, 1.35]}
            )
            fig3.patch.set_facecolor("white")

            # --- GRÁFICO 3.1: EVOLUCIÓN MENSUAL DESGLOSADA (RESOLUCIONES + INFORMES) ---
            df_pr_anio_g = df_pr_f.dropna(subset=["Fecha_Desc_Dt"]).copy()
            anio_max_pr = int(df_pr_anio_g["Fecha_Desc_Dt"].dt.year.max())
            df_pr_anio_g = df_pr_anio_g[df_pr_anio_g["Fecha_Desc_Dt"].dt.year == anio_max_pr]
            df_pr_anio_g["Mes_Num"] = df_pr_anio_g["Fecha_Desc_Dt"].dt.month

            meses_ord = sorted(df_pr_anio_g["Mes_Num"].unique())
            x_mes_lbl = [MESES_ABR[int(m)] for m in meses_ord]
            res_mes_vals = [int(((df_pr_anio_g["Mes_Num"] == m) & (~df_pr_anio_g["Es_Informe"])).sum()) for m in meses_ord]
            inf_mes_vals = [int(((df_pr_anio_g["Mes_Num"] == m) & (df_pr_anio_g["Es_Informe"])).sum()) for m in meses_ord]
            tot_mes_vals = [r + inf for r, inf in zip(res_mes_vals, inf_mes_vals)]

            x_m = np.arange(len(x_mes_lbl))
            w_bar_m = 0.36 if len(x_m) <= 3 else 0.55

            axr1.bar(x_m, res_mes_vals, width=w_bar_m, label="Resoluciones", color="#1E3A8A", edgecolor="#0F172A", linewidth=0.4)
            axr1.bar(x_m, inf_mes_vals, width=w_bar_m, bottom=res_mes_vals, label="Informes", color="#D97706", edgecolor="#0F172A", linewidth=0.4)
            if len(x_m) >= 2:
                axr1.plot(x_m, tot_mes_vals, marker="o", markersize=5.5, linewidth=1.6, linestyle="--", color="#0284C7")

            max_yr1 = max(1, max(tot_mes_vals))
            axr1.set_ylim(0, max_yr1 * 1.28)
            axr1.set_xlim(-0.55, len(x_m) - 0.45)

            for i, (r_v, i_v, t_v) in enumerate(zip(res_mes_vals, inf_mes_vals, tot_mes_vals)):
                axr1.text(i, t_v + max_yr1 * 0.025, f"{t_v:,} exp.", ha="center", va="bottom", fontsize=8.2, fontweight="bold", color="#0F172A")
                if r_v > max_yr1 * 0.10:
                    axr1.text(i, r_v / 2, f"{r_v:,}", ha="center", va="center", fontsize=7.3, fontweight="bold", color="white")
                if i_v > max_yr1 * 0.08:
                    axr1.text(i, r_v + i_v / 2, f"{i_v:,}", ha="center", va="center", fontsize=7.0, fontweight="bold", color="white")

            axr1.set_xticks(x_m)
            axr1.set_xticklabels(x_mes_lbl, fontsize=8.5, fontweight="bold", color="#1E293B")
            axr1.set_title(f"Gráfico 3.1 — Evolución Mensual de Producción ({anio_max_pr})", fontsize=9.5, fontweight="bold", color="#7F1D1D", pad=10)
            axr1.grid(axis="y", linestyle="--", alpha=0.35)
            axr1.spines["top"].set_visible(False)
            axr1.spines["right"].set_visible(False)
            axr1.legend(loc="upper right", fontsize=7.3, frameon=True, facecolor="#F8FAFC", edgecolor="#CBD5E1")

            # --- GRÁFICO 3.2: PRODUCCIÓN POR UNIDAD SIN CORTAR NOMBRES ---
            df_u_plot = t_3_3[~t_3_3["Unidad Orgánica / Funcional"].str.upper().str.startswith("TOTAL")].head(7).iloc[::-1]
            yu = np.arange(len(df_u_plot))
            axr2.barh(yu, df_u_plot["Resoluciones"], height=0.58, label="Resoluciones", color="#1E3A8A", edgecolor="#0F172A", linewidth=0.4)
            axr2.barh(yu, df_u_plot["Informes"], height=0.58, left=df_u_plot["Resoluciones"], label="Informes", color="#D97706", edgecolor="#0F172A", linewidth=0.4)

            max_xr2 = max(1, int(df_u_plot["Total"].max()))
            axr2.set_xlim(0, max_xr2 * 1.28)
            for idx_y, (_, r_u) in zip(yu, df_u_plot.iterrows()):
                axr2.text(
                    int(r_u["Total"]) + max_xr2 * 0.02,
                    idx_y,
                    f"{int(r_u['Total']):,} ({r_u['% Part.']:.1f}%)",
                    va="center",
                    ha="left",
                    fontsize=7.6,
                    fontweight="bold",
                    color="#0F172A"
                )

            etiq_u_envueltas = [textwrap.fill(str(u), width=30) for u in df_u_plot["Unidad Orgánica / Funcional"]]
            axr2.set_yticks(yu)
            axr2.set_yticklabels(etiq_u_envueltas, fontsize=6.9, fontweight="bold", color="#1E293B")
            axr2.set_title(f"Gráfico 3.2 — Producción por Unidad ({mes_act_lbl})", fontsize=9.5, fontweight="bold", color="#7F1D1D", pad=10)
            axr2.grid(axis="x", linestyle="--", alpha=0.35)
            axr2.spines["top"].set_visible(False)
            axr2.spines["right"].set_visible(False)
            axr2.legend(loc="lower right", fontsize=7.3, frameon=True, facecolor="#F8FAFC", edgecolor="#CBD5E1")

            fig3.subplots_adjust(left=0.05, right=0.98, top=0.86, bottom=0.13, wspace=0.58)
            incrustar_figura(fig3, ancho_mm=277, alto_mm=68)

        if not t_3_3.empty:
            dibujar_tabla_ejecutiva(f"3.3. Producción por Unidad Orgánica / Funcional - {mes_act_lbl}", t_3_3, max_filas=60)

        if not t_3_4.empty:
            t_3_4_pdf = t_3_4.copy()
            if "N°" not in t_3_4_pdf.columns:
                nums_mag = []
                contador_mag = 1
                for _, r_m in t_3_4_pdf.iterrows():
                    if str(r_m["Magistrado"]).strip().upper().startswith("TOTAL"):
                        nums_mag.append("")
                    else:
                        nums_mag.append(str(contador_mag))
                        contador_mag += 1
                t_3_4_pdf.insert(0, "N°", nums_mag)

            dibujar_tabla_ejecutiva(f"3.4. Ranking de Producción por Magistrado - {mes_act_lbl}", t_3_4_pdf, max_filas=60)

            # Gráfico de Impacto: Ranking de Producción por Magistrado (Estirado y sin números en etiquetas)
            df_mag_plot = t_3_4_pdf[~t_3_4_pdf["Magistrado"].str.upper().str.startswith("TOTAL")].head(15).iloc[::-1]
            if not df_mag_plot.empty:
                fig4, axm = plt.subplots(figsize=(15.0, 4.3), dpi=220)
                fig4.patch.set_facecolor("white")

                ym = np.arange(len(df_mag_plot))
                axm.barh(ym, df_mag_plot["Resoluciones"], height=0.62, label="Resoluciones", color="#1E3A8A", edgecolor="#0F172A", linewidth=0.4)
                axm.barh(ym, df_mag_plot["Informes"], height=0.62, left=df_mag_plot["Resoluciones"], label="Informes", color="#F59E0B", edgecolor="#0F172A", linewidth=0.4)

                max_xm = max(1, int(df_mag_plot["Total"].max()))
                axm.set_xlim(0, max_xm * 1.18)

                for idx_y, (_, r_m) in zip(ym, df_mag_plot.iterrows()):
                    tot_m = int(r_m["Total"])
                    pct_m = float(r_m["% Part."])
                    axm.text(
                        tot_m + max_xm * 0.012,
                        idx_y,
                        f"{tot_m:,} exp. ({pct_m:.2f}%)",
                        va="center",
                        ha="left",
                        fontsize=7.8,
                        fontweight="bold",
                        color="#0F172A"
                    )

                prom_mag = float(t_3_4_pdf[~t_3_4_pdf["Magistrado"].str.upper().str.startswith("TOTAL")]["Total"].mean())
                axm.axvline(prom_mag, color="#DC2626", linestyle="--", linewidth=1.4, label=f"Promedio: {prom_mag:.1f} exp./mag.")

                etiq_mag = [str(r_m["Magistrado"])[:42] for _, r_m in df_mag_plot.iterrows()]
                axm.set_yticks(ym)
                axm.set_yticklabels(etiq_mag, fontsize=7.8, fontweight="bold", color="#1E293B")
                axm.set_title(f"Gráfico 3.3 — Impacto de Producción por Magistrado ({mes_act_lbl})", fontsize=10.0, fontweight="bold", color="#7F1D1D", pad=8)
                axm.set_xlabel("Cantidad de Expedientes Resueltos / Informados", fontsize=8.0, fontweight="bold", color="#334155")
                axm.grid(axis="x", linestyle="--", alpha=0.35)
                axm.spines["top"].set_visible(False)
                axm.spines["right"].set_visible(False)
                axm.legend(loc="lower right", fontsize=7.8, frameon=True, facecolor="#F8FAFC", edgecolor="#CBD5E1")

                fig4.subplots_adjust(left=0.22, right=0.99, top=0.88, bottom=0.13)
                incrustar_figura(fig4, ancho_mm=277, alto_mm=84)

        return bytes(pdf.output())

    with st.sidebar:
        st.subheader("📥 Exportar Informe Oficial")
        st.download_button("📊 Descargar Informe en Excel", buffer_anexo.getvalue(), f"ANALISIS_ESTADISTICO_{nombre_file_limpio}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        st.download_button("📄 Descargar Reporte Ejecutivo (PDF)", construir_pdf_anexo(), f"REPORTE_ESTADISTICO_{nombre_file_limpio}.pdf", "application/pdf", use_container_width=True)