import io
import os
import unicodedata
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

try:
    import pyodbc
except ImportError:
    pyodbc = None

# ==============================================================================
# CONFIGURACIÓN DE FUENTE DE DATOS (CAMBIAR A "SYBASE" EN PRODUCCIÓN CON UTI)
# ==============================================================================
MODO_DATOS = "EXCEL"  # Opciones: "EXCEL" (piloto) | "SYBASE" (tiempo real ODBC)

st.set_page_config(page_title="ANC - Análisis Estadístico Oficial (UIA)", page_icon="⚖️", layout="wide")

MESES_ABR = {
    1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Ago", 9: "Set", 10: "Oct", 11: "Nov", 12: "Dic"
}
MESES_NOM = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

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
    return df

def categorizar_etapa_auto(etapa_val):
    e = norm_txt(etapa_val)
    if e == "" or e in ("NAN", "NONE", "NULL"):
        return "Otros (Abstenciones, incidentes, etc.)"
    if "CALIFICACION" in e:
        return "Calificación"
    if "INSTRUCCION" in e:
        return "Instrucción"
    if "SANCION" in e:
        return "Sanción"
    if "APELACION" in e:
        return "Apelación"
    return "Otros (Abstenciones, incidentes, etc.)"

def subcat_calificacion(sub_val):
    s = norm_txt(sub_val)
    if s == "CALIFICACION":
        return "   • Calificación propiamente dicha"
    if s == "APELACION":
        return "   • Apelación (en Calificación)"
    return "   • Pend. Informe / Por Elevar / En Recepción / Otros"

def clasificar_presc_tramo(dias):
    if pd.isna(dias):
        return "No Registra"
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

def mostrar_tabla_centrada(df_in):
    if df_in.empty:
        st.info("Sin registros disponibles para mostrar.")
        return

    cols = df_in.columns.tolist()
    cols_num = df_in.select_dtypes(include=["number"]).columns.tolist()

    ths = "".join([
        f"<th style='padding:8px 10px; border:1px solid #334155; text-align:{'left' if i == 0 else 'center'}; color:#f8fafc;'>{c}</th>"
        for i, c in enumerate(cols)
    ])

    filas_html = ""
    for idx, row in df_in.iterrows():
        primera_celda = str(row[cols[0]])
        es_total = any(palabra in primera_celda.upper() for palabra in ["TOTAL", "SUBTOTAL"])
        bg_fila = "#1e293b" if es_total else ("#0f172a" if idx % 2 == 1 else "#111827")
        borde_fila = "1px solid #334155" if es_total else "1px solid #1f2937"
        peso = "700" if es_total else "400"

        celdas = ""
        for i, c in enumerate(cols):
            val = row[c]
            if c in cols_num and pd.notna(val):
                txt_val = f"{val:,.2f}" if "%" in str(c) else f"{int(val):,}"
            else:
                txt_val = str(val) if pd.notna(val) else ""

            alineacion = "left" if i == 0 else "center"
            if i == 0 and txt_val.startswith("   •"):
                estilo_extra = "padding-left:22px; color:#cbd5e1;"
            elif es_total and i == len(cols) - 2:
                estilo_extra = "color:#fbbf24;"
            elif es_total:
                estilo_extra = "color:#f8fafc;"
            else:
                estilo_extra = "color:#e2e8f0;"

            celdas += f"<td style='padding:7px 10px; border:{borde_fila}; text-align:{alineacion}; font-weight:{peso}; {estilo_extra}'>{txt_val}</td>"

        filas_html += f"<tr style='background:{bg_fila};'>{celdas}</tr>"

    html_tabla = f"""
    <div style="overflow-x:auto; margin-bottom:16px; border-radius:8px; border:1px solid #334155;">
        <table style="width:100%; border-collapse:collapse; font-size:13px; background:#111827;">
            <thead>
                <tr style="background:#1e293b; border-bottom:2px solid #475569;">
                    {ths}
                </tr>
            </thead>
            <tbody>
                {filas_html}
            </tbody>
        </table>
    </div>
    """
    st.markdown(html_tabla, unsafe_allow_html=True)

def generar_cuadro_carga(df_sub, col_cat, orden_filas, label_col, es_nacional, nombre_sede, nombre_total="TOTAL"):
    if es_nacional:
        anc_df = df_sub[df_sub["Es_Sede_Central"]]
        oda_df = df_sub[~df_sub["Es_Sede_Central"]]

        anc_tr = anc_df[anc_df["Estado_Grupo"] == "Trámite"].groupby(col_cat).size()
        anc_im = anc_df[anc_df["Estado_Grupo"] == "Res. Impug."].groupby(col_cat).size()
        oda_tr = oda_df[oda_df["Estado_Grupo"] == "Trámite"].groupby(col_cat).size()
        oda_im = oda_df[oda_df["Estado_Grupo"] == "Res. Impug."].groupby(col_cat).size()

        res = pd.DataFrame(index=orden_filas)
        res["ANC Trámite"] = anc_tr.reindex(orden_filas, fill_value=0)
        res["ANC Res. Impug."] = anc_im.reindex(orden_filas, fill_value=0)
        res["Total ANC"] = res["ANC Trámite"] + res["ANC Res. Impug."]
        res["ODANC Trámite"] = oda_tr.reindex(orden_filas, fill_value=0)
        res["ODANC Res. Impug."] = oda_im.reindex(orden_filas, fill_value=0)
        res["Total ODANC"] = res["ODANC Trámite"] + res["ODANC Res. Impug."]
        res["Total Nacional"] = res["Total ANC"] + res["Total ODANC"]

        tot_nac = res["Total Nacional"].sum()
        res["% Nac."] = (res["Total Nacional"] / tot_nac * 100).round(2) if tot_nac > 0 else 0.0
        res = res.reset_index().rename(columns={"index": label_col})

        fila_tot = pd.DataFrame([{
            label_col: nombre_total,
            "ANC Trámite": int(res["ANC Trámite"].sum()),
            "ANC Res. Impug.": int(res["ANC Res. Impug."].sum()),
            "Total ANC": int(res["Total ANC"].sum()),
            "ODANC Trámite": int(res["ODANC Trámite"].sum()),
            "ODANC Res. Impug.": int(res["ODANC Res. Impug."].sum()),
            "Total ODANC": int(res["Total ODANC"].sum()),
            "Total Nacional": int(tot_nac),
            "% Nac.": 100.0 if tot_nac > 0 else 0.0
        }])
        return pd.concat([res, fila_tot], ignore_index=True)
    else:
        tr_s = df_sub[df_sub["Estado_Grupo"] == "Trámite"].groupby(col_cat).size()
        im_s = df_sub[df_sub["Estado_Grupo"] == "Res. Impug."].groupby(col_cat).size()

        res = pd.DataFrame(index=orden_filas)
        res["En Trámite"] = tr_s.reindex(orden_filas, fill_value=0)
        res["Res. Impug."] = im_s.reindex(orden_filas, fill_value=0)
        res["Total"] = res["En Trámite"] + res["Res. Impug."]

        tot_sede = res["Total"].sum()
        res["% Part."] = (res["Total"] / tot_sede * 100).round(2) if tot_sede > 0 else 0.0
        res = res.reset_index().rename(columns={"index": label_col})

        fila_tot = pd.DataFrame([{
            label_col: nombre_total,
            "En Trámite": int(res["En Trámite"].sum()),
            "Res. Impug.": int(res["Res. Impug."].sum()),
            "Total": int(tot_sede),
            "% Part.": 100.0 if tot_sede > 0 else 0.0
        }])
        return pd.concat([res, fila_tot], ignore_index=True)

def estandarizar_tablas(df_c_raw, df_p_raw, df_pr_raw):
    # 1. CARGA ACTUAL
    df_c_orig = limpiar_columnas(df_c_raw.copy())
    mask_no_inv_global = df_c_orig.apply(
        lambda fila: fila.astype(str).str.upper().str.contains("SIN INVENTARIAR", na=False).any(),
        axis=1
    )
    cols_base_c = ["Distrito Actual", "Año", "Fecha creación. Exp.", "Estado Exp.", "Es Carga", "Despacho o Unidad", "Etapa", "Subetapa"]
    cols_c_ok = [c for c in cols_base_c if c in df_c_orig.columns]
    df_c = df_c_orig[cols_c_ok].copy()
    df_c["Es_No_Inventariado"] = mask_no_inv_global
    df_c["Distrito_Norm"] = df_c["Distrito Actual"].apply(norm_txt)
    df_c["Es_Sede_Central"] = df_c["Distrito_Norm"] == "A.N.C. SEDE CENTRAL"
    df_c["Estado_Norm"] = df_c["Estado Exp."].apply(norm_txt)
    df_c["Etapa_Cat"] = df_c["Etapa"].apply(categorizar_etapa_auto)
    df_c["Subetapa_Raw"] = df_c["Subetapa"].fillna("SIN SUBETAPA").astype(str).str.strip()
    df_c["Subetapa_Cat"] = df_c["Subetapa"].apply(subcat_calificacion)
    df_c["Despacho_Norm"] = df_c["Despacho o Unidad"].apply(norm_txt)
    df_c["Fecha_Creacion_Dt"] = pd.to_datetime(df_c["Fecha creación. Exp."], errors="coerce", dayfirst=True)
    df_c["Anio_Num"] = pd.to_numeric(df_c["Año"], errors="coerce").fillna(0).astype(int)

    # 2. PRESCRIPCIONES
    df_p = limpiar_columnas(df_p_raw.copy())
    cols_p_ok = [c for c in ["DISTRITO ACTUAL", "N° EXPEDIENTE", "CARGA", "ETAPA", "FECHA DE PRESCRIPCION"] if c in df_p.columns]
    df_p = df_p[cols_p_ok].copy()
    df_p["Distrito_Norm"] = df_p["DISTRITO ACTUAL"].apply(norm_txt)
    df_p["Es_Sede_Central"] = df_p["Distrito_Norm"] == "A.N.C. SEDE CENTRAL"
    df_p["Carga_Norm"] = df_p["CARGA"].apply(norm_txt) if "CARGA" in df_p.columns else "SI"
    df_p["Fecha_Presc_Dt"] = pd.to_datetime(df_p["FECHA DE PRESCRIPCION"], errors="coerce", dayfirst=True)

    fecha_corte_presc = pd.Timestamp("2026-11-13")
    df_p["Dias_Calc"] = (df_p["Fecha_Presc_Dt"] - fecha_corte_presc).dt.days
    df_p["Tramo_Presc"] = df_p["Dias_Calc"].apply(clasificar_presc_tramo)
    df_p = df_p[df_p["Tramo_Presc"].notna()].copy()

    # 3. PRODUCCIÓN
    df_pr = limpiar_columnas(df_pr_raw.copy())
    cols_pr_ok = [c for c in ["Distrito Actual", "EXPEDIENTE", "INFORME", "RESOLUCIÓN", "FECHA_DESCARGA", "DETALLE DEL TIPO DE PRONUNCIAMIENTO", "TIPO DE AUTO"] if c in df_pr.columns]
    df_pr = df_pr[cols_pr_ok].copy()
    df_pr["Distrito_Norm"] = df_pr["Distrito Actual"].apply(norm_txt)
    df_pr["Es_Sede_Central"] = df_pr["Distrito_Norm"] == "A.N.C. SEDE CENTRAL"
    df_pr["Fecha_Desc_Dt"] = pd.to_datetime(df_pr["FECHA_DESCARGA"], errors="coerce", dayfirst=True)
    df_pr["Es_Informe"] = df_pr["INFORME"].notna() & (df_pr["INFORME"].astype(str).str.strip() != "") if "INFORME" in df_pr.columns else False

    return df_c, df_p, df_pr, fecha_corte_presc

@st.cache_data(show_spinner="Procesando data cruda de Excel...")
def cargar_desde_excel_v11(file_bytes):
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
def cargar_desde_sybase_v11():
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

import glob

if MODO_DATOS == "EXCEL":
    archivos_excel = glob.glob("*.xlsx")
    if not archivos_excel:
        st.error("No se encontró ningún archivo Excel (.xlsx) en el repositorio.")
        st.stop()
    archivo_local = archivos_excel[0]
    with open(archivo_local, "rb") as f:
        df_carga, df_presc, df_prod, fecha_corte = cargar_desde_excel_v11(
            f.read(), archivo_local
        )
else:
    df_carga, df_presc, df_prod, fecha_corte = cargar_desde_sybase_v11()


distritos_unicos = sorted([d for d in df_carga["Distrito_Norm"].unique() if d != "" and d != "A.N.C. SEDE CENTRAL"])
with st.sidebar:
    st.subheader("🎯 Filtro por Sede / ODANC")
    dist_sel = st.selectbox("Visualizar:", ["NACIONAL (ANC Sede Central + 34 ODANCs)", "A.N.C. SEDE CENTRAL"] + distritos_unicos)

es_nacional = (dist_sel == "NACIONAL (ANC Sede Central + 34 ODANCs)")
nombre_ambito = "NACIONAL (ANC SEDE CENTRAL + 34 ODANCs)" if es_nacional else dist_sel

st.markdown(f"""
<div style="background: linear-gradient(90deg, #7f1d1d 0%, #991b1b 60%, #1e293b 100%); padding: 18px 24px; border-radius: 10px; border-left: 6px solid #f59e0b; margin-bottom: 18px;">
    <h2 style="color: white; margin: 0; font-size: 22px;">⚖️ Autoridad Nacional de Control del Poder Judicial</h2>
    <p style="color: #e2e8f0; margin: 4px 0 0 0; font-size: 13px;">ANÁLISIS ESTADÍSTICO DE CARGA POR ETAPAS, PRESCRIPCIONES Y PRODUCCIÓN: <b>{nombre_ambito}</b> (Resumen al {fecha_corte.strftime('%d/%m/%Y')}) — UIA</p>
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

df_no_inv = df_c_f[df_c_f["Es_No_Inventariado"]]

# ==============================================================================
# FILTRADO EXCLUSIVO POR LA COLUMNA "Estado Exp." (33,353 exp. a nivel nacional)
# ==============================================================================
df_tram = df_c_f[df_c_f["Estado_Norm"] == "TRAMITE"].copy()
df_imp = df_c_f[df_c_f["Estado_Norm"].str.contains("RESUELTO|IMPUGNACION", na=False)].copy()
df_act = pd.concat([df_tram, df_imp], ignore_index=True)
df_act["Estado_Grupo"] = np.where(df_act["Estado_Norm"] == "TRAMITE", "Trámite", "Res. Impug.")

df_jef_all = df_act[df_act["Despacho_Norm"].str.contains("JEFATURA", na=False)]
df_p_240 = df_p_f[(df_p_f["Carga_Norm"] == "SI") & (df_p_f["Dias_Calc"] <= 240)]
p_90 = int((df_p_240["Dias_Calc"] <= 90).sum())
p_0 = int((df_p_240["Dias_Calc"] <= 0).sum())

k1, k2, k3, k4, k5 = st.columns(5)
if es_nacional:
    c_anc_tot = int(df_act["Es_Sede_Central"].sum())
    c_oda_tot = int((~df_act["Es_Sede_Central"]).sum())
    j_anc = int(df_jef_all["Es_Sede_Central"].sum())
    j_oda = int((~df_jef_all["Es_Sede_Central"]).sum())
    pr_anc = int(df_pr_f["Es_Sede_Central"].sum())
    pr_oda = int((~df_pr_f["Es_Sede_Central"]).sum())
    ni_anc = int(df_no_inv["Es_Sede_Central"].sum())
    ni_oda = int((~df_no_inv["Es_Sede_Central"]).sum())

    k1.metric("CARGA PROCESAL", f"{len(df_act):,} exp.", f"ANC: {c_anc_tot:,} | ODANC: {c_oda_tot:,}")
    k2.metric("EN JEFATURA", f"{len(df_jef_all):,} exp.", f"ANC: {j_anc:,} | ODANC: {j_oda:,}")
    k3.metric("RIESGO PRESC. (<=240d)", f"{len(df_p_240):,} exp.", f"{p_90:,} <=90d | {p_0:,} Presc.")
    k4.metric("PRODUCCIÓN ACUM.", f"{len(df_pr_f):,} exp.", f"ANC: {pr_anc:,} | ODANC: {pr_oda:,}")
    k5.metric("SIN INVENTARIAR", f"{len(df_no_inv):,} exp.", f"ANC: {ni_anc:,} | ODANC: {ni_oda:,}")
else:
    inf_sede = int(df_pr_f["Es_Informe"].sum())
    res_sede = len(df_pr_f) - inf_sede
    k1.metric("CARGA PROCESAL", f"{len(df_act):,} exp.", f"Trám: {len(df_tram):,} | Impug: {len(df_imp):,}")
    k2.metric("EN JEFATURA", f"{len(df_jef_all):,} exp.", f"Despacho Jefatura ({dist_sel})")
    k3.metric("RIESGO PRESC. (<=240d)", f"{len(df_p_240):,} exp.", f"{p_90:,} <=90d | {p_0:,} Presc.")
    k4.metric("PRODUCCIÓN ACUM.", f"{len(df_pr_f):,} exp.", f"Res: {res_sede:,} | Inf: {inf_sede:,}")
    k5.metric("SIN INVENTARIAR", f"{len(df_no_inv):,} exp.", f"Sede: {dist_sel}")

sub_titulo_sec = "ANC vs. ODANC" if es_nacional else dist_sel

tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Sec. 1.1 - 1.3 (Carga, Prescripción y Etapas)",
    "📅 Sec. 1.4 (Calificación por Año y Mes)",
    f"🏛️ Sec. 1.5 (En Jefatura — {dist_sel if not es_nacional else 'ANC / ODANC'})",
    "⚖️ Sec. 2 (Producción Mensual y Pronunciamientos)"
])

with tab1:
    st.subheader(f"1.1. Según la antigüedad por AÑO de creación / ingreso ({sub_titulo_sec})")
    df_act["Anio_Grupo"] = np.where(df_act["Anio_Num"] >= 2022, df_act["Anio_Num"].astype(str), "2021 y ant.")
    ord_anios = ["2026", "2025", "2024", "2023", "2022", "2021 y ant."]
    tot_lbl_11 = "TOTAL CARGA NACIONAL" if es_nacional else f"TOTAL CARGA {dist_sel}"
    t_1_1 = generar_cuadro_carga(df_act, "Anio_Grupo", ord_anios, "Año de Creación", es_nacional, dist_sel, tot_lbl_11)
    mostrar_tabla_centrada(t_1_1)

    st.markdown("---")
    st.subheader(f"1.2. Según riesgo de prescripción ({sub_titulo_sec})")
    df_p_act = df_p_f[df_p_f["Carga_Norm"] == "SI"].copy()
    ord_tramos = [
        "Prescritos (<= 0 días)",
        "Crítico: De 1 a 15 días",
        "Alto: De 16 a 30 días",
        "Moderado: De 31 a 60 días",
        "Preventivo: De 61 a 90 días",
        "Monitoreo: De 91 a 240 días"
    ]
    if es_nacional:
        p_anc_s = df_p_act[df_p_act["Es_Sede_Central"]].groupby("Tramo_Presc").size().reindex(ord_tramos, fill_value=0)
        p_oda_s = df_p_act[~df_p_act["Es_Sede_Central"]].groupby("Tramo_Presc").size().reindex(ord_tramos, fill_value=0)
        t_1_2 = pd.DataFrame({
            "Tramo de Vencimiento / Riesgo": ord_tramos,
            "ANC Sede Central": p_anc_s.values,
            "% en ANC": (p_anc_s.values / p_anc_s.sum() * 100).round(2) if p_anc_s.sum() else 0,
            "ODANCs (34 Dist.)": p_oda_s.values,
            "% en ODANC": (p_oda_s.values / p_oda_s.sum() * 100).round(2) if p_oda_s.sum() else 0,
            "Total Nacional": p_anc_s.values + p_oda_s.values
        })
        tot_p_nac = t_1_2["Total Nacional"].sum()
        t_1_2["% Nac."] = (t_1_2["Total Nacional"] / tot_p_nac * 100).round(2) if tot_p_nac else 0
        fila_tot_p = pd.DataFrame([{
            "Tramo de Vencimiento / Riesgo": "TOTAL PRESCRIPCIONES",
            "ANC Sede Central": int(p_anc_s.sum()),
            "% en ANC": 100.0 if p_anc_s.sum() else 0.0,
            "ODANCs (34 Dist.)": int(p_oda_s.sum()),
            "% en ODANC": 100.0 if p_oda_s.sum() else 0.0,
            "Total Nacional": int(tot_p_nac),
            "% Nac.": 100.0 if tot_p_nac else 0.0
        }])
        t_1_2 = pd.concat([t_1_2, fila_tot_p], ignore_index=True)
    else:
        p_sede_s = df_p_act.groupby("Tramo_Presc").size().reindex(ord_tramos, fill_value=0)
        tot_p_s = p_sede_s.sum()
        t_1_2 = pd.DataFrame({
            "Tramo de Vencimiento / Riesgo": ord_tramos,
            "Expedientes": p_sede_s.values,
            "% Part.": (p_sede_s.values / tot_p_s * 100).round(2) if tot_p_s else 0
        })
        fila_tot_p = pd.DataFrame([{
            "Tramo de Vencimiento / Riesgo": f"TOTAL PRESCRIPCIONES {dist_sel}",
            "Expedientes": int(tot_p_s),
            "% Part.": 100.0 if tot_p_s else 0.0
        }])
        t_1_2 = pd.concat([t_1_2, fila_tot_p], ignore_index=True)

    mostrar_tabla_centrada(t_1_2)

    if es_nacional:
        top5_p = df_p_act[~df_p_act["Es_Sede_Central"]]["Distrito_Norm"].value_counts().head(5)
        if not top5_p.empty:
            txt_top5 = ", ".join([f"{d.title()} ({c:,})" for d, c in top5_p.items()])
            st.caption(f"*Mayor riesgo ODANC: {txt_top5}.")

    st.markdown("---")
    st.subheader(f"1.3. Según ETAPAS PROCESALES y Sub-etapas de Calificación ({sub_titulo_sec})")
    ord_et = ["Calificación", "Instrucción", "Sanción", "Apelación", "Otros (Abstenciones, incidentes, etc.)"]
    t_et_base = generar_cuadro_carga(df_act, "Etapa_Cat", ord_et, "Etapa / Sub-etapa Procesal", es_nacional, dist_sel, "TOTAL POR ETAPAS")

    df_cal = df_act[df_act["Etapa_Cat"] == "Calificación"].copy()
    ord_sub = [
        "   • Calificación propiamente dicha",
        "   • Apelación (en Calificación)",
        "   • Pend. Informe / Por Elevar / En Recepción / Otros"
    ]
    t_sub_cal = generar_cuadro_carga(df_cal, "Subetapa_Cat", ord_sub, "Etapa / Sub-etapa Procesal", es_nacional, dist_sel).iloc[:-1]
    col_pct = "% Nac." if es_nacional else "% Part."
    col_tot = "Total Nacional" if es_nacional else "Total"
    t_sub_cal[col_pct] = (t_sub_cal[col_tot] / len(df_act) * 100).round(2) if len(df_act) else 0

    t_1_3 = pd.concat([t_et_base.iloc[0:1], t_sub_cal, t_et_base.iloc[1:]], ignore_index=True)
    mostrar_tabla_centrada(t_1_3)

    df_cal_otros = df_cal[df_cal["Subetapa_Cat"] == "   • Pend. Informe / Por Elevar / En Recepción / Otros"]
    if es_nacional:
        det_anc = df_cal_otros[df_cal_otros["Es_Sede_Central"]]["Subetapa_Raw"].value_counts()
        det_oda = df_cal_otros[~df_cal_otros["Es_Sede_Central"]]["Subetapa_Raw"].value_counts()
        if not det_anc.empty:
            st.caption(f"*Detalle Sub-etapas Calif. Sede Central ({det_anc.sum()}): " + ", ".join([f"{k.title()} ({v})" for k, v in det_anc.items()]))
        if not det_oda.empty:
            st.caption(f"*Detalle Sub-etapas Calif. ODANC ({det_oda.sum()}): " + ", ".join([f"{k.title()} ({v})" for k, v in det_oda.items()]))
    else:
        det_s = df_cal_otros["Subetapa_Raw"].value_counts()
        if not det_s.empty:
            st.caption(f"*Detalle Sub-etapas Calif. {dist_sel} ({det_s.sum()}): " + ", ".join([f"{k.title()} ({v})" for k, v in det_s.items()]))

with tab2:
    st.subheader(f"1.4.A. Calificación por Año de Creación ({sub_titulo_sec})")
    df_cal["Anio_Cal_Grp"] = np.where(df_cal["Anio_Num"] >= 2023, df_cal["Anio_Num"].astype(str), "2022 y anteriores")
    ord_cal_a = ["2026", "2025", "2024", "2023", "2022 y anteriores"]
    t_1_4_a = generar_cuadro_carga(df_cal, "Anio_Cal_Grp", ord_cal_a, "Año de Creación", es_nacional, dist_sel, "SUBTOTAL CALIFICACIÓN POR AÑO")
    mostrar_tabla_centrada(t_1_4_a)

    def matriz_mes_anio(df_in, titulo_mat):
        st.markdown(f"**{titulo_mat}**")
        df_v = df_in.dropna(subset=["Fecha_Creacion_Dt"]).copy()
        if df_v.empty:
            st.info("Sin registros con fecha válida.")
            return pd.DataFrame()
        df_v["Mes_N"] = df_v["Fecha_Creacion_Dt"].dt.month
        df_v["Anio_S"] = df_v["Fecha_Creacion_Dt"].dt.year.astype(str)

        p_tr = pd.crosstab(df_v["Mes_N"], df_v[df_v["Estado_Grupo"] == "Trámite"]["Anio_S"]).reindex(range(1, 13), fill_value=0)
        p_tr["Total TRÁMITE"] = p_tr.sum(axis=1)

        p_im = pd.crosstab(df_v["Mes_N"], df_v[df_v["Estado_Grupo"] == "Res. Impug."]["Anio_S"]).reindex(range(1, 13), fill_value=0)
        p_im["Total Resuelto Impug."] = p_im.sum(axis=1)

        tot_gen_col = p_tr["Total TRÁMITE"] + p_im["Total Resuelto Impug."]
        th_tr_sub = "".join([f"<th style='padding:6px; border:1px solid #334155; text-align:center;'>{c}</th>" for c in p_tr.columns])
        th_im_sub = "".join([f"<th style='padding:6px; border:1px solid #334155; text-align:center;'>{c}</th>" for c in p_im.columns])

        html = f"""
        <div style="overflow-x:auto; margin-bottom:16px;">
        <table style="width:100%; border-collapse:collapse; font-size:13px; color:#f8fafc; background:#111827; text-align:center;">
            <thead>
                <tr style="background:#1e293b; border-bottom:1px solid #334155;">
                    <th rowspan="2" style="padding:8px; border:1px solid #334155; vertical-align:middle; text-align:center;">Mes de Creación</th>
                    <th colspan="{len(p_tr.columns)}" style="padding:6px; border:1px solid #334155; text-align:center; color:#38bdf8;">TRÁMITE</th>
                    <th colspan="{len(p_im.columns)}" style="padding:6px; border:1px solid #334155; text-align:center; color:#f59e0b;">RESUELTO - IMPUGNACIÓN</th>
                    <th rowspan="2" style="padding:8px; border:1px solid #334155; vertical-align:middle; text-align:center; background:#0f172a;">Total general</th>
                </tr>
                <tr style="background:#1e293b; border-bottom:2px solid #475569;">
                    {th_tr_sub}
                    {th_im_sub}
                </tr>
            </thead>
            <tbody>
        """
        for m_idx in range(1, 13):
            tds_tr = "".join([f"<td style='padding:6px; border:1px solid #1f2937;'>{int(p_tr.loc[m_idx, c]):,}</td>" for c in p_tr.columns])
            tds_im = "".join([f"<td style='padding:6px; border:1px solid #1f2937;'>{int(p_im.loc[m_idx, c]):,}</td>" for c in p_im.columns])
            html += f"<tr><td style='padding:6px; border:1px solid #1f2937; font-weight:600;'>{MESES_ABR[m_idx]}</td>{tds_tr}{tds_im}<td style='padding:6px; border:1px solid #1f2937; font-weight:700;'>{int(tot_gen_col.loc[m_idx]):,}</td></tr>"

        tds_tr_tot = "".join([f"<td style='padding:7px; border:1px solid #334155; font-weight:700;'>{int(p_tr[c].sum()):,}</td>" for c in p_tr.columns])
        tds_im_tot = "".join([f"<td style='padding:7px; border:1px solid #334155; font-weight:700;'>{int(p_im[c].sum()):,}</td>" for c in p_im.columns])
        html += f"<tr style='background:#1e293b;'><td style='padding:7px; border:1px solid #334155; font-weight:700;'>Total general</td>{tds_tr_tot}{tds_im_tot}<td style='padding:7px; border:1px solid #334155; font-weight:700; color:#fbbf24;'>{int(tot_gen_col.sum()):,}</td></tr></tbody></table></div>"
        st.markdown(html, unsafe_allow_html=True)

        p_tr_ex = p_tr.add_prefix("TRÁMITE - ")
        p_im_ex = p_im.add_prefix("RES. IMPUG. - ")
        df_ex = pd.concat([p_tr_ex, p_im_ex], axis=1)
        df_ex["Total general"] = tot_gen_col
        df_ex.insert(0, "Mes de Creación", [MESES_ABR[m] for m in df_ex.index])
        fila_tot_ex = {"Mes de Creación": "Total general", **{c: int(df_ex[c].sum()) for c in df_ex.columns if c != "Mes de Creación"}}
        return pd.concat([df_ex, pd.DataFrame([fila_tot_ex])], ignore_index=True)

    st.markdown("---")
    if es_nacional:
        mat_1 = matriz_mes_anio(df_cal[df_cal["Es_Sede_Central"]], "1.4.B. Calificación por Mes y Año — ANC Sede Central")
        mat_2 = matriz_mes_anio(df_cal[~df_cal["Es_Sede_Central"]], "1.4.B. Calificación por Mes y Año — ODANCs (34 Distritos)")
    else:
        mat_1 = matriz_mes_anio(df_cal, f"1.4.B. Calificación por Mes y Año — {dist_sel}")
        mat_2 = pd.DataFrame()

    st.markdown("---")
    st.subheader(f"1.4.C. Calificación según Ubicación / Unidad y Sub-etapa ({sub_titulo_sec})")
    df_c14 = df_cal.copy()
    df_c14["Unidad_Grp"] = df_c14["Despacho_Norm"].replace("", "NO ESPECIFICADO")
    df_c14["Sub_Grp"] = df_c14["Subetapa_Raw"].replace("", "SIN SUBETAPA").str.upper()

    unidades_ord = sorted(df_c14["Unidad_Grp"].unique().tolist())
    filas_14c = []
    tot_cal_nac = len(df_c14)

    for u_nom in unidades_ord:
        df_u = df_c14[df_c14["Unidad_Grp"] == u_nom]
        fila_u = generar_cuadro_carga(df_u, "Unidad_Grp", [u_nom], "Unidad / Sub-etapa en Calificación", es_nacional, dist_sel).iloc[0].to_dict()
        if es_nacional:
            fila_u["% Nac."] = round(fila_u["Total Nacional"] / tot_cal_nac * 100, 2) if tot_cal_nac else 0.0
        else:
            fila_u["% Part."] = round(fila_u["Total"] / tot_cal_nac * 100, 2) if tot_cal_nac else 0.0
        filas_14c.append(fila_u)

        subs_u = sorted(df_u["Sub_Grp"].unique().tolist())
        df_subs_u = generar_cuadro_carga(df_u, "Sub_Grp", subs_u, "Unidad / Sub-etapa en Calificación", es_nacional, dist_sel).iloc[:-1]
        for _, r_sub in df_subs_u.iterrows():
            d_sub = r_sub.to_dict()
            nombre_sub = d_sub["Unidad / Sub-etapa en Calificación"]
            d_sub["Unidad / Sub-etapa en Calificación"] = f"   • {nombre_sub}"
            if "% Nac." in d_sub:
                d_sub["% Nac."] = round(d_sub["Total Nacional"] / tot_cal_nac * 100, 2) if tot_cal_nac else 0.0
            else:
                d_sub["% Part."] = round(d_sub["Total"] / tot_cal_nac * 100, 2) if tot_cal_nac else 0.0
            filas_14c.append(d_sub)

    fila_tot_14c = generar_cuadro_carga(df_c14, "Unidad_Grp", unidades_ord, "Unidad / Sub-etapa en Calificación", es_nacional, dist_sel, "TOTAL CALIFICACIÓN POR UNIDAD").iloc[-1].to_dict()
    filas_14c.append(fila_tot_14c)

    t_1_4_c = pd.DataFrame(filas_14c)
    mostrar_tabla_centrada(t_1_4_c)

with tab3:
    st.subheader(f"1.5. En Jefatura ({sub_titulo_sec})")
    ord_jef = ["Apelación", "Calificación", "Otros (Abstenciones, incidentes, etc.)", "Instrucción", "Sanción"]
    t_1_5 = generar_cuadro_carga(df_jef_all, "Etapa_Cat", ord_jef, "Etapa en Jefatura", es_nacional, dist_sel, "TOTAL EN JEFATURA")
    mostrar_tabla_centrada(t_1_5)

with tab4:
    st.subheader(f"2.1. Comparativo Mensual de Producción: Mes Anterior vs. Mes Actual ({sub_titulo_sec})")
    fechas_pr = df_pr_f["Fecha_Desc_Dt"].dropna()
    mes_act_lbl = "Mes Actual"
    if not fechas_pr.empty:
        per_act = fechas_pr.max().to_period("M")
        per_ant = per_act - 1
        mes_act_lbl = f"{MESES_NOM[per_act.month]} {per_act.year}"
        df_pr_f = df_pr_f.copy()
        df_pr_f["Periodo_M"] = df_pr_f["Fecha_Desc_Dt"].dt.to_period("M")

        filas_pr = []
        periodos_eval = [
            (per_ant, f"Mes anterior ({MESES_NOM[per_ant.month]} {per_ant.year})"),
            (per_act, f"Mes actual ({mes_act_lbl})")
        ]
        for per_obj, etiq in periodos_eval:
            sub_p = df_pr_f[df_pr_f["Periodo_M"] == per_obj]
            inf_p = int(sub_p["Es_Informe"].sum())
            res_p = len(sub_p) - inf_p
            if es_nacional:
                filas_pr.append({
                    "Periodo Estadístico": etiq,
                    "ANC Sede Central": int(sub_p["Es_Sede_Central"].sum()),
                    "ODANCs (34 Dist.)": int((~sub_p["Es_Sede_Central"]).sum()),
                    "Resoluciones": res_p,
                    "Informes": inf_p,
                    "Total Nacional": len(sub_p)
                })
            else:
                filas_pr.append({
                    "Periodo Estadístico": etiq,
                    "Resoluciones": res_p,
                    "Informes": inf_p,
                    "Total": len(sub_p)
                })

        t_2_1 = pd.DataFrame(filas_pr)
        col_tot_21 = "Total Nacional" if es_nacional else "Total"
        tot_2_1 = t_2_1[col_tot_21].sum()
        t_2_1["% Acum."] = (t_2_1[col_tot_21] / tot_2_1 * 100).round(2) if tot_2_1 else 0.0

        fila_tot_21 = {"Periodo Estadístico": "TOTAL PRODUCCIÓN ACUMULADA"}
        for c_num in [c for c in t_2_1.columns if c not in ("Periodo Estadístico", "% Acum.")]:
            fila_tot_21[c_num] = int(t_2_1[c_num].sum())
        fila_tot_21["% Acum."] = 100.0 if tot_2_1 else 0.0
        t_2_1 = pd.concat([t_2_1, pd.DataFrame([fila_tot_21])], ignore_index=True)
        mostrar_tabla_centrada(t_2_1)

    st.markdown("---")
    st.subheader(f"2.2. Producción por Tipo de Pronunciamiento y/o Sanción — {mes_act_lbl} ({sub_titulo_sec})")
    if "DETALLE DEL TIPO DE PRONUNCIAMIENTO" in df_pr_f.columns:
        df_pr_mes_act = df_pr_f[df_pr_f["Periodo_M"] == per_act].copy() if not fechas_pr.empty else df_pr_f.copy()
        df_pr_mes_act["Pron_Limpio"] = df_pr_mes_act["DETALLE DEL TIPO DE PRONUNCIAMIENTO"].fillna("NO ESPECIFICADO").astype(str).str.strip()
        if es_nacional:
            pr_anc_t = df_pr_mes_act[df_pr_mes_act["Es_Sede_Central"]].groupby("Pron_Limpio").size()
            pr_oda_t = df_pr_mes_act[~df_pr_mes_act["Es_Sede_Central"]].groupby("Pron_Limpio").size()
            t_2_2_full = pd.DataFrame({"ANC Sede Central": pr_anc_t, "ODANCs (34 Dist.)": pr_oda_t}).fillna(0).astype(int)
            t_2_2_full["Total Nacional"] = t_2_2_full["ANC Sede Central"] + t_2_2_full["ODANCs (34 Dist.)"]
            t_2_2_full = t_2_2_full.sort_values("Total Nacional", ascending=False).reset_index().rename(columns={"Pron_Limpio": "Tipo de Pronunciamiento / Sanción"})
            tot_22_nac = t_2_2_full["Total Nacional"].sum()

            if len(t_2_2_full) > 15:
                top15 = t_2_2_full.head(15).copy()
                resto = t_2_2_full.iloc[15:]
                fila_otros = pd.DataFrame([{
                    "Tipo de Pronunciamiento / Sanción": f"Otros pronunciamientos ({len(resto)} tipos)",
                    "ANC Sede Central": int(resto["ANC Sede Central"].sum()),
                    "ODANCs (34 Dist.)": int(resto["ODANCs (34 Dist.)"].sum()),
                    "Total Nacional": int(resto["Total Nacional"].sum())
                }])
                t_2_2 = pd.concat([top15, fila_otros], ignore_index=True)
            else:
                t_2_2 = t_2_2_full.copy()

            t_2_2["% Nac."] = (t_2_2["Total Nacional"] / tot_22_nac * 100).round(2) if tot_22_nac else 0.0
            fila_tot_22 = pd.DataFrame([{
                "Tipo de Pronunciamiento / Sanción": f"TOTAL {mes_act_lbl.upper()}",
                "ANC Sede Central": int(t_2_2["ANC Sede Central"].sum()),
                "ODANCs (34 Dist.)": int(t_2_2["ODANCs (34 Dist.)"].sum()),
                "Total Nacional": int(tot_22_nac),
                "% Nac.": 100.0 if tot_22_nac else 0.0
            }])
            t_2_2 = pd.concat([t_2_2, fila_tot_22], ignore_index=True)
        else:
            t_2_2_full = df_pr_mes_act["Pron_Limpio"].value_counts().reset_index()
            t_2_2_full.columns = ["Tipo de Pronunciamiento / Sanción", "Total"]
            tot_22_s = t_2_2_full["Total"].sum()
            if len(t_2_2_full) > 15:
                top15 = t_2_2_full.head(15).copy()
                resto = t_2_2_full.iloc[15:]
                fila_otros = pd.DataFrame([{
                    "Tipo de Pronunciamiento / Sanción": f"Otros pronunciamientos ({len(resto)} tipos)",
                    "Total": int(resto["Total"].sum())
                }])
                t_2_2 = pd.concat([top15, fila_otros], ignore_index=True)
            else:
                t_2_2 = t_2_2_full.copy()
            t_2_2["% Part."] = (t_2_2["Total"] / tot_22_s * 100).round(2) if tot_22_s else 0.0
            fila_tot_22 = pd.DataFrame([{
                "Tipo de Pronunciamiento / Sanción": f"TOTAL {mes_act_lbl.upper()} ({dist_sel})",
                "Total": int(tot_22_s),
                "% Part.": 100.0 if tot_22_s else 0.0
            }])
            t_2_2 = pd.concat([t_2_2, fila_tot_22], ignore_index=True)
        mostrar_tabla_centrada(t_2_2)

buffer_anexo = io.BytesIO()
with pd.ExcelWriter(buffer_anexo, engine="openpyxl") as writer:
    t_1_1.to_excel(writer, sheet_name="1_1_Anio_Creacion", index=False)
    t_1_2.to_excel(writer, sheet_name="1_2_Prescripciones", index=False)
    t_1_3.to_excel(writer, sheet_name="1_3_Etapas_Subetapas", index=False)
    t_1_4_a.to_excel(writer, sheet_name="1_4A_Calif_Anio", index=False)
    if not mat_1.empty:
        mat_1.to_excel(writer, sheet_name="1_4B_Calif_Mes_1", index=False)
    if not mat_2.empty:
        mat_2.to_excel(writer, sheet_name="1_4B_Calif_Mes_2", index=False)
    t_1_4_c.to_excel(writer, sheet_name="1_4C_Calif_Unidad", index=False)
    t_1_5.to_excel(writer, sheet_name="1_5_Jefatura", index=False)
    if not fechas_pr.empty:
        t_2_1.to_excel(writer, sheet_name="2_1_Produccion_Mensual", index=False)
    if "DETALLE DEL TIPO DE PRONUNCIAMIENTO" in df_pr_f.columns:
        t_2_2.to_excel(writer, sheet_name="2_2_Prod_Pronunciamiento", index=False)

nombre_file_limpio = "NACIONAL" if es_nacional else dist_sel.replace(" ", "_").replace(".", "")

def limpiar_latin1(txt):
    s = str(txt).replace("•", "-")
    return s.encode("latin-1", "replace").decode("latin-1")

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
    pdf.cell(0, 5, limpiar_latin1(f"ANÁLISIS ESTADÍSTICO DE CARGA POR ETAPAS, PRESCRIPCIONES Y PRODUCCIÓN: {nombre_ambito} | Corte: {fecha_corte.strftime('%d/%m/%Y')} (UIA)"), ln=True)

    pdf.ln(6)

    pdf.set_font("Helvetica", "B", 8.5)
    pdf.set_text_color(30, 41, 59)
    resumen_kpis = (
        f"CARGA PROCESAL: {len(df_act):,} exp.   |   "
        f"EN JEFATURA: {len(df_jef_all):,} exp.   |   "
        f"RIESGO PRESC. (<=240d): {len(df_p_240):,} exp.   |   "
        f"PRODUCCIÓN ACUM.: {len(df_pr_f):,} exp.   |   "
        f"SIN INVENTARIAR: {len(df_no_inv):,} exp."
    )
    pdf.set_fill_color(241, 245, 249)
    pdf.cell(277, 8, limpiar_latin1(resumen_kpis), border=1, ln=True, align="C", fill=True)
    pdf.ln(4)

    def dibujar_tabla_pdf(titulo, df_tab, max_filas=45):
        if df_tab is None or df_tab.empty:
            return
        
        if len(df_tab) > max_filas:
            df_show = pd.concat([df_tab.head(max_filas - 1), df_tab.tail(1)], ignore_index=True)
        else:
            df_show = df_tab.copy()
            
        cols = df_show.columns.tolist()
        n_cols = len(cols)
        ancho_total = 277.0
        w_primera = 88.0 if n_cols > 3 else 155.0
        w_resto = (ancho_total - w_primera) / max(1, (n_cols - 1))
        anchos = [w_primera] + [w_resto] * (n_cols - 1)

        pdf.set_font("Helvetica", "B", 9.5)
        pdf.set_text_color(127, 29, 29)
        pdf.cell(0, 6, limpiar_latin1(titulo), ln=True)

        pdf.set_font("Helvetica", "B", 7.5)
        pdf.set_fill_color(30, 41, 59)
        pdf.set_text_color(255, 255, 255)
        for i, c in enumerate(cols):
            alin = "L" if i == 0 else "C"
            pdf.cell(anchos[i], 6.5, limpiar_latin1(str(c)[:34]), border=1, align=alin, fill=True)
        pdf.ln()

        pdf.set_font("Helvetica", "", 7.5)
        for idx, row in df_show.iterrows():
            prim = str(row[cols[0]])
            es_tot = any(p in prim.upper() for p in ["TOTAL", "SUBTOTAL"])
            if es_tot:
                pdf.set_font("Helvetica", "B", 7.5)
                pdf.set_fill_color(226, 232, 240)
                pdf.set_text_color(15, 23, 42)
            else:
                pdf.set_font("Helvetica", "", 7.5)
                if idx % 2 == 0:
                    pdf.set_fill_color(255, 255, 255)
                else:
                    pdf.set_fill_color(248, 250, 252)
                pdf.set_text_color(30, 41, 59)

            for i, c in enumerate(cols):
                val = row[c]
                if isinstance(val, (int, float, np.integer, np.floating)) and pd.notna(val):
                    txt = f"{val:,.2f}" if "%" in str(c) else f"{int(val):,}"
                else:
                    txt = str(val) if pd.notna(val) else ""
                alin = "L" if i == 0 else "C"
                pdf.cell(anchos[i], 5.8, limpiar_latin1(txt[:45]), border=1, align=alin, fill=True)
            pdf.ln()
        pdf.ln(3)

    dibujar_tabla_pdf(f"1.1. Carga Procesal por Año de Creación ({sub_titulo_sec})", t_1_1)
    dibujar_tabla_pdf(f"1.2. Riesgo de Prescripción ({sub_titulo_sec})", t_1_2)
    dibujar_tabla_pdf(f"1.3. Etapas Procesales y Sub-etapas de Calificación ({sub_titulo_sec})", t_1_3)
    dibujar_tabla_pdf(f"1.4.A. Calificación por Año de Creación ({sub_titulo_sec})", t_1_4_a)

    def dibujar_matriz_pdf(titulo_m, df_m):
        if df_m is None or df_m.empty:
            return
        cols_m = df_m.columns.tolist()
        n_c = len(cols_m)
        w_mes = 28.0
        w_col = (277.0 - w_mes) / max(1, (n_c - 1))
        anchos_m = [w_mes] + [w_col] * (n_c - 1)

        pdf.set_font("Helvetica", "B", 9.5)
        pdf.set_text_color(127, 29, 29)
        pdf.cell(0, 6, limpiar_latin1(titulo_m), ln=True)

        pdf.set_font("Helvetica", "B", 6.2)
        pdf.set_fill_color(30, 41, 59)
        pdf.set_text_color(255, 255, 255)
        for i, c in enumerate(cols_m):
            etiq_c = str(c).replace("TRÁMITE - ", "TR ").replace("RES. IMPUG. - ", "IMP ").replace("Total ", "Tot.")
            pdf.cell(anchos_m[i], 6.0, limpiar_latin1(etiq_c[:16]), border=1, align="C", fill=True)
        pdf.ln()

        for idx, row in df_m.iterrows():
            es_t = "TOTAL" in str(row[cols_m[0]]).upper()
            pdf.set_font("Helvetica", "B" if es_t else "", 6.5)
            if es_t:
                pdf.set_fill_color(226, 232, 240)
            elif idx % 2 == 0:
                pdf.set_fill_color(255, 255, 255)
            else:
                pdf.set_fill_color(248, 250, 252)
            
            if es_t:
                pdf.set_text_color(15, 23, 42)
            else:
                pdf.set_text_color(30, 41, 59)

            for i, c in enumerate(cols_m):
                val = row[c]
                txt = f"{int(val):,}" if isinstance(val, (int, float, np.integer, np.floating)) and pd.notna(val) else str(val)
                pdf.cell(anchos_m[i], 5.2, limpiar_latin1(txt[:16]), border=1, align="C", fill=True)
            pdf.ln()
        pdf.ln(3)

    if es_nacional:
        dibujar_matriz_pdf("1.4.B. Calificación por Mes y Año - ANC Sede Central", mat_1)
        dibujar_matriz_pdf("1.4.B. Calificación por Mes y Año - ODANCs (34 Distritos)", mat_2)
    else:
        dibujar_matriz_pdf(f"1.4.B. Calificación por Mes y Año - {dist_sel}", mat_1)

    dibujar_tabla_pdf(f"1.4.C. Calificación según Ubicación / Unidad y Sub-etapa ({sub_titulo_sec})", t_1_4_c, max_filas=100)
    dibujar_tabla_pdf(f"1.5. Expedientes en Jefatura ({sub_titulo_sec})", t_1_5)
    if not fechas_pr.empty:
        dibujar_tabla_pdf(f"2.1. Comparativo Mensual de Producción ({sub_titulo_sec})", t_2_1)
    if "DETALLE DEL TIPO DE PRONUNCIAMIENTO" in df_pr_f.columns:
        dibujar_tabla_pdf(f"2.2. Producción por Tipo de Pronunciamiento / Sanción - {mes_act_lbl}", t_2_2)

    return bytes(pdf.output())

pdf_bytes = construir_pdf_anexo()

with st.sidebar:
    st.subheader("📥 Exportar Informe Oficial")
    st.download_button(
        label="📊 Descargar Informe en Excel",
        data=buffer_anexo.getvalue(),
        file_name=f"ANALISIS_ESTADISTICO_{nombre_file_limpio}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    st.download_button(
        label="📄 Descargar Reporte Ejecutivo (PDF)",
        data=pdf_bytes,
        file_name=f"REPORTE_ESTADISTICO_{nombre_file_limpio}.pdf",
        mime="application/pdf",
        use_container_width=True
    )