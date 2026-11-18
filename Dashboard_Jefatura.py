import io
import os
import re
import glob
import unicodedata
import numpy as np
import pandas as pd
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

# Estilos CSS para tarjetas KPI clicables (Carga, Atención Preferencial, Jefatura, Producción Acumulada y 6 Semáforos)
st.markdown("""
<style>
    /* Botones KPI Carga, Atención Preferencial, Jefatura y Producción Acumulada */
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

    /* Estilos específicos para los 6 botones semáforo de Prescripción */
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
    /* Colores por columna de semáforo (1 al 6) */
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

MESES_ABR = {
    1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Ago", 9: "Set", 10: "Oct", 11: "Nov", 12: "Dic"
}
MESES_NOM = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

# ==============================================================================
# COLUMNAS EXACTAS PARA EL EXPLORADOR DE DETALLE DE CARGA (SIN REPETIDOS)
# ==============================================================================
COLUMNAS_EXACTAS_DETALLE = [
    "N°",
    "Distrito Actual",
    "Número",
    "Año",
    "Tipo",
    "Distrito Judicial",
    "Encargado Actual",
    "Fecha creación. Exp.",
    "Estado Exp.",
    "Nota de Inventario",
    "Personal Asignado Actual",
    "Tipo de Atención",
    "Despacho o Unidad",
    "Etapa",
    "Subetapa",
    "Fecha Envio/Recep. Despacho",
    "Magistrado"
]

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
    df = df.loc[:, ~df.columns.duplicated()].copy()
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

def clasificar_presc_rango(dias):
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

def formatear_detalle_exp_presc(row):
    raw_exp = str(row.get("N° EXPEDIENTE", "")).strip()
    base_exp = re.split(r'\s+de\s+Origen\b|\s+DE\s+ORIGEN\b', raw_exp)[0].strip()

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
    es_sc = row.get("Es_Sede_Central", False)
    if es_sc:
        if "JEFATURA" in s:
            return "JEFATURA ANC"
        m = re.search(r'-(UCIP|UPAD|USAP|OCPAD|UPSI|CII|OCCPC)\b', s)
        if m:
            return m.group(1)
        return s if s else "NO ESPECIFICADO"
    else:
        if "JEFATURA" in s:
            return "Jefatura ODANC"
        if "CALIFICACION" in s or "UDCIP" in s or "UCIP" in s or "CALIF" in s:
            return "Calificación e Inv. Preliminar (UDCIP/UCIP)"
        if "PROCEDIMIENTO" in s or "UDPAD" in s or "UPAD" in s or "OCPAD" in s or "PAD" in s or "DISCIPLINARIO" in s or "INSTRUCCION" in s:
            return "Procedimiento Adm. Disciplinario (UDPAD/UPAD)"
        if "SANCION" in s or "APELACION" in s or "UDSAP" in s or "USAP" in s:
            return "Sanción y Apelación (UDSAP/USAP)"
        if "PREVENCION" in s or "SUPERVISION" in s or "INSPECCION" in s or "UPSI" in s or "UDPSI" in s:
            return "Prevención, Supervisión e Inspección (UDPSI/UPSI)"
        if "INTERVENCION INMEDIATA" in s or "-CII" in s:
            return "Comisión de Intervención Inmediata (CII)"
        return "Otros / Especializados"

def map_unidad_todos(enc):
    s = norm_txt(enc)
    if "JEFATURA" in s:
        return "Jefatura (ANC / ODANC)"
    if "CALIFICACION" in s or "UDCIP" in s or "UCIP" in s or "CALIF" in s:
        return "Calificación e Inv. Preliminar (UCIP / UDCIP)"
    if "PROCEDIMIENTO" in s or "UDPAD" in s or "UPAD" in s or "OCPAD" in s or "PAD" in s or "DISCIPLINARIO" in s or "INSTRUCCION" in s:
        return "Procedimiento Adm. Disciplinario (UPAD / OCPAD / UDPAD)"
    if "SANCION" in s or "APELACION" in s or "UDSAP" in s or "USAP" in s:
        return "Sanción y Apelación (USAP / UDSAP)"
    if "PREVENCION" in s or "SUPERVISION" in s or "INSPECCION" in s or "UPSI" in s or "UDPSI" in s:
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

def limpiar_latin1(txt):
    s = (
        str(txt)
        .replace("•", "-")
        .replace("🔴", "")
        .replace("🟡", "")
        .replace("🟢", "")
        .replace("🔵", "")
        .replace("🟠", "")
        .replace("🟣", "")
    )
    return s.encode("latin-1", "replace").decode("latin-1").strip()

# ==============================================================================
# FUNCIONES DE EXPORTACIÓN DEL DETALLE DE EXPEDIENTES (EXCEL Y PDF)
# ==============================================================================
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
            align = "C" if i == 0 or "FECHA" in str(col_name).upper() or "DIAS" in str(col_name).upper() or "DÍAS" in str(col_name).upper() else "L"
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

# ==============================================================================
# TARJETA DE IMPACTO: EXPEDIENTE MÁS ANTIGUO
# ==============================================================================
def abreviar_unidad_card(raw_unidad, es_sede_central):
    u = norm_txt(raw_unidad)
    if u in ("", "-", "NONE", "NAN"):
        return "Sin unidad"
    if "JEFATURA" in u:
        return "Jefatura ANC" if es_sede_central else "Jefatura"
    if "SANCION" in u or "APELACION" in u or "UDSAP" in u or "USAP" in u:
        return "USAP" if es_sede_central else "UDSAP"
    if "CALIFICACION" in u or "UDCIP" in u or "UCIP" in u or "CALIF" in u:
        return "UCIP" if es_sede_central else "UDCIP"
    if "PROCEDIMIENTO" in u or "UDPAD" in u or "UPAD" in u or "OCPAD" in u or "PAD" in u or "DISCIPLINARIO" in u:
        return "UPAD" if es_sede_central else "UDPAD"
    if "PREVENCION" in u or "SUPERVISION" in u or "INSPECCION" in u or "UDPSI" in u or "UPSI" in u:
        return "UPSI" if es_sede_central else "UDPSI"
    if "OCCPC" in u or "CONTROL PREVENTIVO" in u:
        return "OCCPC"
    if "INTERVENCION INMEDIATA" in u or "CII" in u:
        return "CII"
    return str(raw_unidad).strip().title()

def render_tarjeta_mas_antiguo(df_origen, es_nac, titulo_card="EXP. MÁS ANTIGUO"):
    df_fechas = df_origen.dropna(subset=["Fecha_Creacion_Dt"]).sort_values("Fecha_Creacion_Dt", ascending=True)
    if df_fechas.empty:
        mes_anio = "Sin datos"
        unidad_txt = "-"
        unidad_full = "-"
        etapa_txt = "-"
        estado_txt = "-"
        sede_txt = ""
        color_est = "#94a3b8"
        bg_est = "rgba(148, 163, 184, 0.15)"
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
        if estado_txt == "Trámite":
            color_est = "#38bdf8"
            bg_est = "rgba(56, 189, 248, 0.15)"
        else:
            color_est = "#fbbf24"
            bg_est = "rgba(251, 191, 36, 0.15)"

    cabecera_sede = (
        f'<span title="{sede_txt}" style="background: rgba(245, 158, 11, 0.18); color: #fde68a; font-weight: 700; font-size: 0.70rem; padding: 1px 6px; border-radius: 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 55%;">📍 {sede_txt}</span>'
        if es_nac and sede_txt else ""
    )

    html_card = (
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
        f'</div>'
        f'</div>'
    )
    st.markdown(html_card, unsafe_allow_html=True)

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

def construir_df_matriz_jefatura_15b(df_in):
    ord_jef_et = ["Apelación", "Calificación", "Otros (Abstenciones, incidentes, etc.)", "Instrucción", "Sanción"]
    df_v = df_in.copy()
    if df_v.empty:
        return pd.DataFrame(), [], []

    df_v["Anio_Col"] = np.where(df_v["Anio_Num"] > 0, df_v["Anio_Num"].astype(str), "S/A")
    anios_disp = sorted([a for a in df_v["Anio_Col"].unique() if a != "S/A"])
    if "S/A" in df_v["Anio_Col"].unique():
        anios_disp.append("S/A")

    p_tr = pd.crosstab(df_v["Etapa_Cat"], df_v[df_v["Estado_Grupo"] == "Trámite"]["Anio_Col"]).reindex(index=ord_jef_et, columns=anios_disp, fill_value=0)
    p_tr["Total TRÁMITE"] = p_tr.sum(axis=1)

    p_im = pd.crosstab(df_v["Etapa_Cat"], df_v[df_v["Estado_Grupo"] == "Res. Impug."]["Anio_Col"]).reindex(index=ord_jef_et, columns=anios_disp, fill_value=0)
    p_im["Total Res. Impug."] = p_im.sum(axis=1)

    tot_gen_col = p_tr["Total TRÁMITE"] + p_im["Total Res. Impug."]
    cols_tr = list(p_tr.columns)
    cols_im = list(p_im.columns)

    df_ex = pd.concat([p_tr.add_prefix("TRÁMITE - "), p_im.add_prefix("RES. IMPUG. - ")], axis=1)
    df_ex["Total general"] = tot_gen_col
    df_ex.insert(0, "Etapa en Jefatura", ord_jef_et)

    fila_tot = {
        "Etapa en Jefatura": "TOTAL EN JEFATURA",
        **{c: int(df_ex[c].sum()) for c in df_ex.columns if c != "Etapa en Jefatura"}
    }
    df_final = pd.concat([df_ex, pd.DataFrame([fila_tot])], ignore_index=True)
    return df_final, cols_tr, cols_im

def estandarizar_tablas(df_c_raw, df_p_raw, df_pr_raw):
    # 1. CARGA ACTUAL
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
    df_c["Etapa_Cat"] = df_c["Etapa"].apply(categorizar_etapa_auto)
    df_c["Subetapa_Raw"] = df_c["Subetapa"].fillna("SIN SUBETAPA").astype(str).str.strip()
    df_c["Subetapa_Cat"] = df_c["Subetapa"].apply(subcat_calificacion)
    df_c["Despacho_Norm"] = df_c["Despacho o Unidad"].apply(norm_txt)
    df_c["Fecha_Creacion_Dt"] = pd.to_datetime(df_c["Fecha creación. Exp."], errors="coerce", dayfirst=True)
    df_c["Anio_Num"] = pd.to_numeric(df_c["Año"], errors="coerce").fillna(0).astype(int)

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

    # 2. PRESCRIPCIONES
    df_p = limpiar_columnas(df_p_raw.copy())
    df_p["Distrito_Norm"] = df_p["DISTRITO ACTUAL"].apply(norm_txt)
    df_p["Es_Sede_Central"] = df_p["Distrito_Norm"] == "A.N.C. SEDE CENTRAL"
    df_p["Carga_Norm"] = df_p["CARGA"].apply(norm_txt) if "CARGA" in df_p.columns else "SI"
    df_p["Etapa_Presc_Cat"] = df_p["ETAPA"].apply(categorizar_etapa_auto) if "ETAPA" in df_p.columns else "Otros (Abstenciones, incidentes, etc.)"
    df_p["Fecha_Presc_Dt"] = pd.to_datetime(df_p["FECHA DE PRESCRIPCION"], errors="coerce", dayfirst=True)

    col_ap = next((c for c in df_p.columns if "PREFERENCIAL" in norm_txt(c)), None)
    if col_ap is not None:
        ap_norm = df_p[col_ap].apply(norm_txt)
        df_p["Es_Atencion_Pref"] = (ap_norm != "") & (ap_norm != "SIN ATENCION PREFERENCIAL") & (ap_norm != "NAN") & (ap_norm != "NONE")
        df_p["Atencion_Pref_Norm"] = ap_norm
    else:
        df_p["Es_Atencion_Pref"] = False
        df_p["Atencion_Pref_Norm"] = "SIN ATENCION PREFERENCIAL"

    col_mag_p = next((c for c in df_p.columns if norm_txt(c) == "MAGISTRADO"), None)
    if col_mag_p is None and "Magistrado" in df_c.columns:
        df_c_map = df_c[["Distrito_Norm", "Número", "Año", "Magistrado"]].drop_duplicates(subset=["Distrito_Norm", "Número", "Año"])
        if "NUMERO" in df_p.columns and "AÑO" in df_p.columns:
            df_p["_num_tmp"] = df_p["NUMERO"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
            df_p["_anio_tmp"] = df_p["AÑO"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
            df_p = df_p.merge(
                df_c_map,
                left_on=["Distrito_Norm", "_num_tmp", "_anio_tmp"],
                right_on=["Distrito_Norm", "Número", "Año"],
                how="left"
            ).drop(columns=["_num_tmp", "_anio_tmp", "Número", "Año"], errors="ignore")
            df_p["MAGISTRADO"] = df_p["Magistrado"].fillna("-")
            df_p = df_p.drop(columns=["Magistrado"], errors="ignore")
    elif col_mag_p is not None and col_mag_p != "MAGISTRADO":
        df_p["MAGISTRADO"] = df_p[col_mag_p].fillna("-")

    fecha_corte_presc = pd.Timestamp("2026-11-13")
    df_p["Dias_Calc"] = (df_p["Fecha_Presc_Dt"] - fecha_corte_presc).dt.days
    df_p["Rango_Presc"] = df_p["Dias_Calc"].apply(clasificar_presc_rango)
    df_p = df_p[df_p["Rango_Presc"].notna()].copy()

    # 3. PRODUCCIÓN
    df_pr = limpiar_columnas(df_pr_raw.copy())
    cols_pr_ok = [
        c for c in [
            "Distrito Actual", "EXPEDIENTE", "ENCARGADO", "MAGISTRADO",
            "INFORME", "RESOLUCIÓN", "FECHA_DESCARGA",
            "DETALLE DEL TIPO DE PRONUNCIAMIENTO", "TIPO DE AUTO", "DESPACHO_ANC"
        ] if c in df_pr.columns
    ]
    df_pr = df_pr[cols_pr_ok].copy()
    df_pr["Distrito_Norm"] = df_pr["Distrito Actual"].apply(norm_txt)
    df_pr["Es_Sede_Central"] = df_pr["Distrito_Norm"] == "A.N.C. SEDE CENTRAL"
    df_pr["Fecha_Desc_Dt"] = pd.to_datetime(df_pr["FECHA_DESCARGA"], errors="coerce", dayfirst=True)
    df_pr["Es_Informe"] = df_pr["INFORME"].notna() & (df_pr["INFORME"].astype(str).str.strip() != "") if "INFORME" in df_pr.columns else False

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

    return df_c, df_p, df_pr, fecha_corte_presc

@st.cache_data(show_spinner="Procesando data cruda de Excel...")
def cargar_desde_excel_v27(file_bytes):
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
def cargar_desde_sybase_v27():
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
        st.error("No se encontró ningún archivo Excel (.xlsx) en el repositorio.")
        st.stop()
    archivo_local = next(iter(archivos_excel))
    with open(archivo_local, "rb") as f:
        df_carga, df_presc, df_prod, fecha_corte = cargar_desde_excel_v27(f.read())
else:
    df_carga, df_presc, df_prod, fecha_corte = cargar_desde_sybase_v27()

distritos_unicos = sorted([d for d in df_carga["Distrito_Norm"].unique() if d != "" and d != "A.N.C. SEDE CENTRAL"])
opciones_filtro = ["TODOS", "A.N.C. SEDE CENTRAL"] + distritos_unicos
with st.sidebar:
    st.subheader("🎯 Filtro por Sede / ODANC")
    dist_sel = st.selectbox("Visualizar:", opciones_filtro, index=1)

es_nacional = (dist_sel == "TODOS")
nombre_ambito = "TODOS (ANC SEDE CENTRAL + 34 ODANCs)" if es_nacional else dist_sel

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
# FILTRADO EXCLUSIVO POR LA COLUMNA "Estado Exp."
# ==============================================================================
df_tram = df_c_f[df_c_f["Estado_Norm"] == "TRAMITE"].copy()
df_imp = df_c_f[df_c_f["Estado_Norm"].str.contains("RESUELTO|IMPUGNACION", na=False)].copy()
df_act = pd.concat([df_tram, df_imp], ignore_index=True)
df_act["Estado_Grupo"] = np.where(df_act["Estado_Norm"] == "TRAMITE", "Trámite", "Res. Impug.")

df_jef_all = df_act[df_act["Despacho_Norm"].str.contains("JEFATURA", na=False)]
df_p_240 = df_p_f[(df_p_f["Carga_Norm"] == "SI") & (df_p_f["Dias_Calc"] <= 240)]
df_p_act = df_p_f[df_p_f["Carga_Norm"] == "SI"].copy()
df_p_vigente = df_p_act[df_p_act["Dias_Calc"] >= 1].copy()
df_p_atencion_pref = df_p_vigente[df_p_vigente["Es_Atencion_Pref"]].copy()

# ==============================================================================
# CONTROL DE VISTA EXCLUSIVA DE DETALLE
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
            opts_e = ["Todas"] + sorted([x for x in df_det["Etapa"].unique() if x != "-"])
            sel_e = st.selectbox("Filtrar por Etapa:", opts_e, key=f"f_etapa_{clave}")
            if sel_e != "Todas":
                df_det = df_det[df_det["Etapa"] == sel_e]
    with f2:
        if "Año" in df_det.columns:
            opts_a = ["Todos"] + sorted([x for x in df_det["Año"].unique() if x != "-"], reverse=True)
            sel_a = st.selectbox("Filtrar por Año:", opts_a, key=f"f_anio_{clave}")
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

    altura_dinamica = min(480, max(75, (len(df_mostrar) + 1) * 35 + 3))
    st.dataframe(df_mostrar, use_container_width=True, hide_index=True, height=altura_dinamica)

    b1, b2, _ = st.columns([1, 1, 3])
    nombre_arch = titulo.replace(" ", "_").replace("-", "").replace(".", "")
    with b1:
        st.download_button(
            label="📥 Exportar en Excel (.xlsx)",
            data=exportar_detalle_excel(df_mostrar),
            file_name=f"{nombre_arch}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"xlsx_{clave}",
            use_container_width=True
        )
    with b2:
        st.download_button(
            label="📄 Exportar en PDF",
            data=exportar_detalle_pdf(df_mostrar, titulo=titulo),
            file_name=f"{nombre_arch}.pdf",
            mime="application/pdf",
            key=f"pdf_{clave}",
            use_container_width=True
        )

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
        opts_et = ["Todas"] + sorted([x for x in df_det["Etapa_Presc_Cat"].unique() if pd.notna(x)])
        sel_et = st.selectbox("Filtrar por Etapa Procesal:", opts_et, key=f"f_etapa_p_{clave}")
        if sel_et != "Todas":
            df_det = df_det[df_det["Etapa_Presc_Cat"] == sel_et]
    with f2:
        col_mag_det = next((c for c in df_det.columns if norm_txt(c) == "MAGISTRADO"), None)
        if col_mag_det:
            opts_m = ["Todos"] + sorted([str(x) for x in df_det[col_mag_det].dropna().unique() if str(x).strip() not in ("", "-", "None", "nan")])
            sel_m = st.selectbox("Filtrar por Magistrado:", opts_m, key=f"f_mag_p_{clave}")
            if sel_m != "Todos":
                df_det = df_det[df_det[col_mag_det].astype(str) == sel_m]
    with f3:
        busq = st.text_input("Buscar por N° Expediente, Distrito o cualquier campo:", key=f"f_busq_p_{clave}")
        if busq:
            mask = df_det.astype(str).apply(lambda x: x.str.contains(busq, case=False, na=False, regex=False)).any(axis=1)
            df_det = df_det[mask]

    cols_internas = {
        "Distrito_Norm", "Es_Sede_Central", "Carga_Norm", "Etapa_Presc_Cat",
        "Fecha_Presc_Dt", "Dias_Calc", "Rango_Presc", "Es_Atencion_Pref", "Atencion_Pref_Norm"
    }
    cols_orig_p = [c for c in df_det.columns if c not in cols_internas and c != "N°"]

    df_mostrar = df_det[cols_orig_p].copy()
    if "FECHA DE PRESCRIPCION" in df_mostrar.columns:
        f_p_dt = pd.to_datetime(df_mostrar["FECHA DE PRESCRIPCION"], errors="coerce", dayfirst=True)
        df_mostrar["FECHA DE PRESCRIPCION"] = f_p_dt.dt.strftime("%d/%m/%Y").fillna(df_mostrar["FECHA DE PRESCRIPCION"].astype(str))

    df_mostrar["DÍAS P/ PRESCRIBIR"] = df_det["Dias_Calc"].astype(int)
    df_mostrar["RANGO DE RIESGO DE EXPEDIENTES PRESCRITOS Y POR PRESCRIBIR"] = df_det["Rango_Presc"]

    df_mostrar = df_mostrar.reset_index(drop=True)
    df_mostrar.insert(0, "N°", range(1, len(df_mostrar) + 1))

    st.caption(f"Mostrando **{len(df_mostrar):,}** expedientes encontrados (ordenados del más próximo a prescribir).")

    altura_dinamica = min(480, max(75, (len(df_mostrar) + 1) * 35 + 3))
    st.dataframe(df_mostrar, use_container_width=True, hide_index=True, height=altura_dinamica)

    b1, b2, _ = st.columns([1, 1, 3])
    nombre_arch = titulo.replace(" ", "_").replace("-", "").replace(".", "").replace("(", "").replace(")", "").replace("<=", "menor_igual_").replace(":", "")
    with b1:
        st.download_button(
            label="📥 Exportar en Excel (.xlsx)",
            data=exportar_detalle_excel(df_mostrar),
            file_name=f"{nombre_arch}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"xlsx_p_{clave}",
            use_container_width=True
        )
    with b2:
        st.download_button(
            label="📄 Exportar en PDF",
            data=exportar_detalle_pdf(df_mostrar, titulo=titulo),
            file_name=f"{nombre_arch}.pdf",
            mime="application/pdf",
            key=f"pdf_p_{clave}",
            use_container_width=True
        )
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

    # Columnas auxiliares para Año y Mes (solo nombre del mes, sin año)
    if "Fecha_Desc_Dt" in df_det.columns:
        df_det["_anio_desc"] = df_det["Fecha_Desc_Dt"].apply(
            lambda d: str(int(d.year)) if pd.notna(d) else "Sin año"
        )
        df_det["_mes_num"] = df_det["Fecha_Desc_Dt"].apply(
            lambda d: int(d.month) if pd.notna(d) else 99
        )
        df_det["_mes_lbl"] = df_det["Fecha_Desc_Dt"].apply(
            lambda d: MESES_NOM[d.month] if pd.notna(d) else "Sin fecha"
        )
    else:
        df_det["_anio_desc"] = "Sin año"
        df_det["_mes_num"] = 99
        df_det["_mes_lbl"] = "Sin fecha"

    df_det["_tipo_doc"] = np.where(df_det["Es_Informe"], "Informe", "Resolución")

    # 5 Filtros: Año (2026 por defecto), Mes (mes actual por defecto), Magistrado, Tipo y Buscador
    f1, f2, f3, f4, f5 = st.columns([0.9, 1.1, 1.4, 1.1, 1.5])
    with f1:
        anios_ord = sorted(
            [a for a in df_det["_anio_desc"].unique() if a != "Sin año"],
            reverse=True
        )
        if "Sin año" in df_det["_anio_desc"].unique():
            anios_ord.append("Sin año")
        opts_anio = ["Todos"] + anios_ord
        idx_anio_def = opts_anio.index("2026") if "2026" in opts_anio else (1 if len(opts_anio) > 1 else 0)
        sel_anio = st.selectbox("Año (Descarga):", opts_anio, index=idx_anio_def, key=f"f_anio_pr_{clave}")
        if sel_anio != "Todos":
            df_det = df_det[df_det["_anio_desc"] == sel_anio]

    with f2:
        meses_ord = (
            df_det[["_mes_num", "_mes_lbl"]]
            .drop_duplicates()
            .sort_values("_mes_num", ascending=True)["_mes_lbl"]
            .tolist()
        )
        opts_mes = ["Todos"] + meses_ord
        mes_actual_def = None
        if "Fecha_Desc_Dt" in df_det.columns and df_det["Fecha_Desc_Dt"].notna().any():
            max_f = df_det["Fecha_Desc_Dt"].dropna().max()
            mes_actual_def = MESES_NOM[int(max_f.month)]
        idx_mes_def = opts_mes.index(mes_actual_def) if mes_actual_def in opts_mes else 0
        sel_mes = st.selectbox("Mes de Resol.:", opts_mes, index=idx_mes_def, key=f"f_mes_pr_{clave}")
        if sel_mes != "Todos":
            df_det = df_det[df_det["_mes_lbl"] == sel_mes]

    with f3:
        col_mag_pr = "MAGISTRADO" if "MAGISTRADO" in df_det.columns else "Magistrado_Prod"
        opts_mag = ["Todos"] + sorted([
            str(x).strip() for x in df_det[col_mag_pr].dropna().unique()
            if str(x).strip() not in ("", "-", "None", "nan", "NO ESPECIFICADO")
        ])
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

    cols_internas_pr = {
        "Distrito_Norm", "Es_Sede_Central", "Fecha_Desc_Dt", "Es_Informe",
        "Unidad_Prod", "Unidad_Prod_Todos", "Magistrado_Prod", "Periodo_M",
        "Pron_Limpio", "_anio_desc", "_mes_num", "_mes_lbl", "_tipo_doc"
    }
    cols_orig_pr = [c for c in df_det.columns if c not in cols_internas_pr and c != "N°"]

    df_mostrar = df_det[cols_orig_pr].copy()
    if "FECHA_DESCARGA" in df_mostrar.columns:
        f_d_dt = pd.to_datetime(df_mostrar["FECHA_DESCARGA"], errors="coerce", dayfirst=True)
        df_mostrar["FECHA_DESCARGA"] = f_d_dt.dt.strftime("%d/%m/%Y").fillna("")

    for col in df_mostrar.columns:
        df_mostrar[col] = (
            df_mostrar[col]
            .fillna("")
            .astype(str)
            .replace({"None": "", "none": "", "NONE": "", "nan": "", "NaN": "", "NAN": "", "NaT": "", "NAT": "", "<NA>": ""})
            .str.strip()
        )

    df_mostrar.insert(0, "Tipo Doc.", df_det["_tipo_doc"].values)
    df_mostrar = df_mostrar.reset_index(drop=True)
    df_mostrar.insert(0, "N°", range(1, len(df_mostrar) + 1))

    st.caption(f"Mostrando **{len(df_mostrar):,}** registros de producción encontrados.")

    altura_dinamica = min(480, max(75, (len(df_mostrar) + 1) * 35 + 3))
    st.dataframe(df_mostrar, use_container_width=True, hide_index=True, height=altura_dinamica)

    b1, b2, _ = st.columns([1, 1, 3])
    nombre_arch = titulo.replace(" ", "_").replace("-", "").replace(".", "")
    with b1:
        st.download_button(
            label="📥 Exportar en Excel (.xlsx)",
            data=exportar_detalle_excel(df_mostrar),
            file_name=f"{nombre_arch}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"xlsx_pr_{clave}",
            use_container_width=True
        )
    with b2:
        st.download_button(
            label="📄 Exportar en PDF",
            data=exportar_detalle_pdf(df_mostrar, titulo=titulo),
            file_name=f"{nombre_arch}.pdf",
            mime="application/pdf",
            key=f"pdf_pr_{clave}",
            use_container_width=True
        )
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
        df_det["_anio_desc"] = df_det["Fecha_Desc_Dt"].apply(
            lambda d: str(int(d.year)) if pd.notna(d) else "Sin año"
        )
        df_det["_mes_num"] = df_det["Fecha_Desc_Dt"].apply(
            lambda d: int(d.month) if pd.notna(d) else 99
        )
        df_det["_mes_lbl"] = df_det["Fecha_Desc_Dt"].apply(
            lambda d: MESES_NOM[d.month] if pd.notna(d) else "Sin fecha"
        )
    else:
        df_det["_anio_desc"] = "Sin año"
        df_det["_mes_num"] = 99
        df_det["_mes_lbl"] = "Sin fecha"

    df_det["_tipo_doc"] = np.where(df_det["Es_Informe"], "Informe", "Resolución")

    f1, f2, f3, f4, f5 = st.columns([0.9, 1.1, 1.4, 1.1, 1.5])
    with f1:
        anios_ord = sorted(
            [a for a in df_det["_anio_desc"].unique() if a != "Sin año"],
            reverse=True
        )
        if "Sin año" in df_det["_anio_desc"].unique():
            anios_ord.append("Sin año")
        opts_anio = ["Todos"] + anios_ord
        idx_anio_def = opts_anio.index("2026") if "2026" in opts_anio else (1 if len(opts_anio) > 1 else 0)
        sel_anio = st.selectbox("Año (Descarga):", opts_anio, index=idx_anio_def, key=f"f_anio_pr_{clave}")
        if sel_anio != "Todos":
            df_det = df_det[df_det["_anio_desc"] == sel_anio]

    with f2:
        meses_ord = (
            df_det[["_mes_num", "_mes_lbl"]]
            .drop_duplicates()
            .sort_values("_mes_num", ascending=True)["_mes_lbl"]
            .tolist()
        )
        opts_mes = ["Todos"] + meses_ord
        mes_actual_def = None
        if "Fecha_Desc_Dt" in df_det.columns and df_det["Fecha_Desc_Dt"].notna().any():
            max_f = df_det["Fecha_Desc_Dt"].dropna().max()
            mes_actual_def = MESES_NOM[int(max_f.month)]
        idx_mes_def = opts_mes.index(mes_actual_def) if mes_actual_def in opts_mes else 0
        sel_mes = st.selectbox("Mes de Resol.:", opts_mes, index=idx_mes_def, key=f"f_mes_pr_{clave}")
        if sel_mes != "Todos":
            df_det = df_det[df_det["_mes_lbl"] == sel_mes]

    with f3:
        col_mag_pr = "MAGISTRADO" if "MAGISTRADO" in df_det.columns else "Magistrado_Prod"
        opts_mag = ["Todos"] + sorted([
            str(x).strip() for x in df_det[col_mag_pr].dropna().unique()
            if str(x).strip() not in ("", "-", "None", "nan", "NO ESPECIFICADO")
        ])
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

    cols_internas_pr = {
        "Distrito_Norm", "Es_Sede_Central", "Fecha_Desc_Dt", "Es_Informe",
        "Unidad_Prod", "Unidad_Prod_Todos", "Magistrado_Prod", "Periodo_M",
        "Pron_Limpio", "_anio_desc", "_mes_num", "_mes_lbl", "_tipo_doc"
    }
    cols_orig_pr = [c for c in df_det.columns if c not in cols_internas_pr and c != "N°"]

    df_mostrar = df_det[cols_orig_pr].copy()
    if "FECHA_DESCARGA" in df_mostrar.columns:
        f_d_dt = pd.to_datetime(df_mostrar["FECHA_DESCARGA"], errors="coerce", dayfirst=True)
        df_mostrar["FECHA_DESCARGA"] = f_d_dt.dt.strftime("%d/%m/%Y").fillna("")

    for col in df_mostrar.columns:
        df_mostrar[col] = (
            df_mostrar[col]
            .fillna("")
            .astype(str)
            .replace({"None": "", "none": "", "NONE": "", "nan": "", "NaN": "", "NAN": "", "NaT": "", "NAT": "", "<NA>": ""})
            .str.strip()
        )

    df_mostrar.insert(0, "Tipo Doc.", df_det["_tipo_doc"].values)
    df_mostrar = df_mostrar.reset_index(drop=True)
    df_mostrar.insert(0, "N°", range(1, len(df_mostrar) + 1))

    st.caption(f"Mostrando **{len(df_mostrar):,}** registros de producción encontrados.")

    altura_dinamica = min(480, max(75, (len(df_mostrar) + 1) * 35 + 3))
    st.dataframe(df_mostrar, use_container_width=True, hide_index=True, height=altura_dinamica)

    b1, b2, _ = st.columns([1, 1, 3])
    nombre_arch = titulo.replace(" ", "_").replace("-", "").replace(".", "")
    with b1:
        st.download_button(
            label="📥 Exportar en Excel (.xlsx)",
            data=exportar_detalle_excel(df_mostrar),
            file_name=f"{nombre_arch}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"xlsx_pr_{clave}",
            use_container_width=True
        )
    with b2:
        st.download_button(
            label="📄 Exportar en PDF",
            data=exportar_detalle_pdf(df_mostrar, titulo=titulo),
            file_name=f"{nombre_arch}.pdf",
            mime="application/pdf",
            key=f"pdf_pr_{clave}",
            use_container_width=True
        )
def render_kpis_carga():
    c_anc_tot = int(df_act["Es_Sede_Central"].sum())
    c_oda_tot = int((~df_act["Es_Sede_Central"]).sum())

    # Fila 1: 3 Tarjetas KPI clicables de Carga
    t1, t2, t3 = st.columns(3)
    with t1:
        sub_tot = (
            f"↑ ANC: {c_anc_tot:,} | ODANC: {c_oda_tot:,}"
            if es_nacional
            else f"↑ 100.00% de {dist_sel}"
        )
        st.button(
            f"CARGA PROCESAL TOTAL\n**{len(df_act):,} exp.**\n*{sub_tot}*",
            key=f"kpi_btn_total_{st.session_state.get('vista_tab', '1')}",
            on_click=alternar_kpi,
            args=("TOTAL",),
            use_container_width=True
        )
    with t2:
        pct_tr = (len(df_tram) / len(df_act) * 100) if len(df_act) > 0 else 0.0
        st.button(
            f"EN TRÁMITE\n**{len(df_tram):,} exp.**\n*↑ {pct_tr:.1f}% de la carga*",
            key=f"kpi_btn_tramite_{st.session_state.get('vista_tab', '1')}",
            on_click=alternar_kpi,
            args=("TRAMITE",),
            use_container_width=True
        )
    with t3:
        pct_im = (len(df_imp) / len(df_act) * 100) if len(df_act) > 0 else 0.0
        st.button(
            f"RESUELTO - IMPUG. TRÁNSITO\n**{len(df_imp):,} exp.**\n*↑ {pct_im:.1f}% de la carga*",
            key=f"kpi_btn_resuelto_{st.session_state.get('vista_tab', '1')}",
            on_click=alternar_kpi,
            args=("RESUELTO",),
            use_container_width=True
        )

    # Fila 2: 4 Tarjetas
    j_anc = int(df_jef_all["Es_Sede_Central"].sum())
    j_oda = int((~df_jef_all["Es_Sede_Central"]).sum())
    ni_anc = int(df_no_inv["Es_Sede_Central"].sum())
    ni_oda = int((~df_no_inv["Es_Sede_Central"]).sum())
    cal_tot = int((df_act["Etapa_Cat"] == "Calificación").sum())
    pct_cal = (cal_tot / len(df_act) * 100) if len(df_act) > 0 else 0.0

    k1, k2, k3, k4 = st.columns(4)
    if es_nacional:
        k1.metric("EN CALIFICACIÓN", f"{cal_tot:,} exp.", f"{pct_cal:.1f}% de la carga total")
        k2.metric("EN JEFATURA", f"{len(df_jef_all):,} exp.", f"ANC: {j_anc:,} | ODANC: {j_oda:,}")
        k3.metric("SIN INVENTARIAR", f"{len(df_no_inv):,} exp.", f"ANC: {ni_anc:,} | ODANC: {ni_oda:,}")
    else:
        k1.metric("EN CALIFICACIÓN", f"{cal_tot:,} exp.", f"{pct_cal:.1f}% de la carga en sede")
        k2.metric("EN JEFATURA", f"{len(df_jef_all):,} exp.", f"Despacho Jefatura ({dist_sel})")
        k3.metric("SIN INVENTARIAR", f"{len(df_no_inv):,} exp.", f"Sede: {dist_sel}")
    with k4:
        render_tarjeta_mas_antiguo(df_act, es_nacional, "EXP. MÁS ANTIGUO")
    st.markdown("<div style='margin-bottom: 10px;'></div>", unsafe_allow_html=True)

# ==============================================================================
# CÁLCULO PREVIO DE TODAS LAS TABLAS PARA EXPORTACIÓN EXCEL/PDF SIEMPRE DISPONIBLE
# ==============================================================================
sub_titulo_sec = "ANC vs. ODANC" if es_nacional else dist_sel

# Tabla 1.1 (Antigüedad por Año)
df_act["Anio_Grupo"] = np.where(df_act["Anio_Num"] >= 2022, df_act["Anio_Num"].astype(str), "2021 y ant.")
ord_anios = ["2026", "2025", "2024", "2023", "2022", "2021 y ant."]
tot_lbl_11 = "TOTAL CARGA (TODOS)" if es_nacional else f"TOTAL CARGA {dist_sel}"
t_1_1 = generar_cuadro_carga(df_act, "Anio_Grupo", ord_anios, "Año de Creación", es_nacional, dist_sel, tot_lbl_11)

# Tabla 1.2 (Etapas Procesales y Sub-etapas de Calificación)
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
t_1_2_etapas = pd.concat([t_et_base.iloc[0:1], t_sub_cal, t_et_base.iloc[1:]], ignore_index=True)

# Tabla 1.3.A (Rango de Riesgo de Expedientes Prescritos y por Prescribir) y 1.3.B (Por Etapa Procesal)
ord_rangos = [
    "Prescritos (<= 0 días)",
    "Crítico: De 1 a 15 días",
    "Alto: De 16 a 30 días",
    "Moderado: De 31 a 60 días",
    "Preventivo: De 61 a 90 días",
    "Monitoreo: De 91 a 240 días"
]
config_colores_rangos = [
    ("PRESC_RANGO_0", "Prescritos (<= 0 días)", "Prescritos", "<= 0 días", "#ef4444", "🔴"),
    ("PRESC_RANGO_1", "Crítico: De 1 a 15 días", "Crítico", "De 1 a 15 días", "#facc15", "🟡"),
    ("PRESC_RANGO_2", "Alto: De 16 a 30 días", "Alto", "De 16 a 30 días", "#65a30d", "🟢"),
    ("PRESC_RANGO_3", "Moderado: De 31 a 60 días", "Moderado", "De 31 a 60 días", "#2563eb", "🔵"),
    ("PRESC_RANGO_4", "Preventivo: De 61 a 90 días", "Preventivo", "De 61 a 90 días", "#06b6d4", "🔵"),
    ("PRESC_RANGO_5", "Monitoreo: De 91 a 240 días", "Monitoreo", "De 91 a 240 días", "#8b5cf6", "🟣"),
]
etiquetas_rangos_color = [f"{emoji} {clave_rg}" for _, clave_rg, _, _, _, emoji in config_colores_rangos]

if es_nacional:
    p_anc_s = df_p_act[df_p_act["Es_Sede_Central"]].groupby("Rango_Presc").size().reindex(ord_rangos, fill_value=0)
    p_oda_s = df_p_act[~df_p_act["Es_Sede_Central"]].groupby("Rango_Presc").size().reindex(ord_rangos, fill_value=0)
    t_1_3_presc = pd.DataFrame({
        "Rango de Riesgo de Expedientes Prescritos y por Prescribir": etiquetas_rangos_color,
        "ANC Sede Central": p_anc_s.values,
        "% en ANC": (p_anc_s.values / p_anc_s.sum() * 100).round(2) if p_anc_s.sum() else 0,
        "ODANCs (34 Dist.)": p_oda_s.values,
        "% en ODANC": (p_oda_s.values / p_oda_s.sum() * 100).round(2) if p_oda_s.sum() else 0,
        "Total Nacional": p_anc_s.values + p_oda_s.values
    })
    tot_p_nac = t_1_3_presc["Total Nacional"].sum()
    t_1_3_presc["% Nac."] = (t_1_3_presc["Total Nacional"] / tot_p_nac * 100).round(2) if tot_p_nac else 0
    fila_tot_p = pd.DataFrame([{
        "Rango de Riesgo de Expedientes Prescritos y por Prescribir": "TOTAL PRESCRIPCIONES",
        "ANC Sede Central": int(p_anc_s.sum()),
        "% en ANC": 100.0 if p_anc_s.sum() else 0.0,
        "ODANCs (34 Dist.)": int(p_oda_s.sum()),
        "% en ODANC": 100.0 if p_oda_s.sum() else 0.0,
        "Total Nacional": int(tot_p_nac),
        "% Nac.": 100.0 if tot_p_nac else 0.0
    }])
    t_1_3_presc = pd.concat([t_1_3_presc, fila_tot_p], ignore_index=True)
else:
    p_sede_s = df_p_act.groupby("Rango_Presc").size().reindex(ord_rangos, fill_value=0)
    tot_p_s = p_sede_s.sum()
    t_1_3_presc = pd.DataFrame({
        "Rango de Riesgo de Expedientes Prescritos y por Prescribir": etiquetas_rangos_color,
        "Expedientes": p_sede_s.values,
        "% Part.": (p_sede_s.values / tot_p_s * 100).round(2) if tot_p_s else 0
    })
    fila_tot_p = pd.DataFrame([{
        "Rango de Riesgo de Expedientes Prescritos y por Prescribir": f"TOTAL PRESCRIPCIONES {dist_sel}",
        "Expedientes": int(tot_p_s),
        "% Part.": 100.0 if tot_p_s else 0.0
    }])
    t_1_3_presc = pd.concat([t_1_3_presc, fila_tot_p], ignore_index=True)

# Cuadro 1.3.B: Prescripciones por Etapa Procesal con los 6 rangos exactos
ord_et_p = ["Instrucción", "Calificación", "Sanción", "Apelación", "Otros (Abstenciones, incidentes, etc.)"]
ct_et_rg = pd.crosstab(df_p_act["Etapa_Presc_Cat"], df_p_act["Rango_Presc"]).reindex(index=ord_et_p, columns=ord_rangos, fill_value=0)
t_1_3_et = pd.DataFrame(index=ord_et_p)
t_1_3_et["🔴 Prescritos (<=0d)"] = ct_et_rg["Prescritos (<= 0 días)"]
t_1_3_et["🟡 Crítico (1-15d)"] = ct_et_rg["Crítico: De 1 a 15 días"]
t_1_3_et["🟢 Alto (16-30d)"] = ct_et_rg["Alto: De 16 a 30 días"]
t_1_3_et["🔵 Moderado (31-60d)"] = ct_et_rg["Moderado: De 31 a 60 días"]
t_1_3_et["🔵 Preventivo (61-90d)"] = ct_et_rg["Preventivo: De 61 a 90 días"]
t_1_3_et["🟣 Monitoreo (91-240d)"] = ct_et_rg["Monitoreo: De 91 a 240 días"]
t_1_3_et["Total"] = t_1_3_et.sum(axis=1)
tot_et_p_gen = int(t_1_3_et["Total"].sum())
t_1_3_et["% Part."] = (t_1_3_et["Total"] / tot_et_p_gen * 100).round(2) if tot_et_p_gen > 0 else 0.0
t_1_3_et = t_1_3_et.reset_index().rename(columns={"index": "Etapa Procesal"})

fila_tot_13et = {
    "Etapa Procesal": "TOTAL POR ETAPA PROCESAL",
    "🔴 Prescritos (<=0d)": int(t_1_3_et["🔴 Prescritos (<=0d)"].sum()),
    "🟡 Crítico (1-15d)": int(t_1_3_et["🟡 Crítico (1-15d)"].sum()),
    "🟢 Alto (16-30d)": int(t_1_3_et["🟢 Alto (16-30d)"].sum()),
    "🔵 Moderado (31-60d)": int(t_1_3_et["🔵 Moderado (31-60d)"].sum()),
    "🔵 Preventivo (61-90d)": int(t_1_3_et["🔵 Preventivo (61-90d)"].sum()),
    "🟣 Monitoreo (91-240d)": int(t_1_3_et["🟣 Monitoreo (91-240d)"].sum()),
    "Total": tot_et_p_gen,
    "% Part.": 100.0 if tot_et_p_gen > 0 else 0.0
}
t_1_3_et = pd.concat([t_1_3_et, pd.DataFrame([fila_tot_13et])], ignore_index=True)

# Tabla 1.4.A y 1.4.C
df_cal["Anio_Cal_Grp"] = np.where(df_cal["Anio_Num"] >= 2023, df_cal["Anio_Num"].astype(str), "2022 y anteriores")
ord_cal_a = ["2026", "2025", "2024", "2023", "2022 y anteriores"]
t_1_4_a = generar_cuadro_carga(df_cal, "Anio_Cal_Grp", ord_cal_a, "Año de Creación", es_nacional, dist_sel, "SUBTOTAL CALIFICACIÓN POR AÑO")

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

# Tabla 1.5.A y Tabla 1.5.B (Jefatura)
ord_jef = ["Apelación", "Calificación", "Otros (Abstenciones, incidentes, etc.)", "Instrucción", "Sanción"]
t_1_5 = generar_cuadro_carga(df_jef_all, "Etapa_Cat", ord_jef, "Etapa en Jefatura", es_nacional, dist_sel, "TOTAL EN JEFATURA")
t_1_5_b, cols_15b_tr, cols_15b_im = construir_df_matriz_jefatura_15b(df_jef_all)

# ==============================================================================
# TABLAS DE PRODUCCIÓN (2.1 Comparativo, 2.2 Pronunciamientos, 2.3 Unidad, 2.4 Magistrado)
# ==============================================================================
fechas_pr = df_pr_f["Fecha_Desc_Dt"].dropna()
mes_act_lbl = "Mes Actual"
t_2_1 = pd.DataFrame()
t_2_2 = pd.DataFrame()
t_2_3 = pd.DataFrame()
t_2_4 = pd.DataFrame()
per_act = None
per_ant = None
df_pr_mes_act = df_pr_f.copy()

if not fechas_pr.empty:
    per_act = fechas_pr.max().to_period("M")
    per_ant = per_act - 1
    mes_act_lbl = f"{MESES_NOM[per_act.month]} {per_act.year}"
    df_pr_f = df_pr_f.copy()
    df_pr_f["Periodo_M"] = df_pr_f["Fecha_Desc_Dt"].dt.to_period("M")
    df_pr_mes_act = df_pr_f[df_pr_f["Periodo_M"] == per_act].copy()

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

if "DETALLE DEL TIPO DE PRONUNCIAMIENTO" in df_pr_mes_act.columns:
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

# Tabla 2.3 (Producción por Unidad) y Tabla 2.4 (Producción por Magistrado)
col_u_23 = "Unidad_Prod_Todos" if es_nacional else "Unidad_Prod"
tot_lbl_23 = f"TOTAL {mes_act_lbl.upper()} (TODOS)" if es_nacional else f"TOTAL {mes_act_lbl.upper()} ({dist_sel})"
t_2_3 = generar_cuadro_produccion_mes_actual(df_pr_mes_act, col_u_23, "Unidad Orgánica / Funcional", tot_lbl_23)

tot_lbl_24 = f"TOTAL {mes_act_lbl.upper()} (TODOS)" if es_nacional else f"TOTAL {mes_act_lbl.upper()} ({dist_sel})"
t_2_4 = generar_cuadro_produccion_mes_actual(df_pr_mes_act, "Magistrado_Prod", "Magistrado", tot_lbl_24)

def construir_df_matriz_14b(df_in):
    df_v = df_in.dropna(subset=["Fecha_Creacion_Dt"]).copy()
    if df_v.empty:
        return pd.DataFrame()
    df_v["Mes_N"] = df_v["Fecha_Creacion_Dt"].dt.month
    df_v["Anio_S"] = df_v["Fecha_Creacion_Dt"].dt.year.astype(str)
    p_tr = pd.crosstab(df_v["Mes_N"], df_v[df_v["Estado_Grupo"] == "Trámite"]["Anio_S"]).reindex(range(1, 13), fill_value=0)
    p_tr["Total TRÁMITE"] = p_tr.sum(axis=1)
    p_im = pd.crosstab(df_v["Mes_N"], df_v[df_v["Estado_Grupo"] == "Res. Impug."]["Anio_S"]).reindex(range(1, 13), fill_value=0)
    p_im["Total Resuelto Impug."] = p_im.sum(axis=1)
    tot_gen_col = p_tr["Total TRÁMITE"] + p_im["Total Resuelto Impug."]
    p_tr_ex = p_tr.add_prefix("TRÁMITE - ")
    p_im_ex = p_im.add_prefix("RES. IMPUG. - ")
    df_ex = pd.concat([p_tr_ex, p_im_ex], axis=1)
    df_ex["Total general"] = tot_gen_col
    df_ex.insert(0, "Mes de Creación", [MESES_ABR[m] for m in df_ex.index])
    fila_tot_ex = {"Mes de Creación": "Total general", **{c: int(df_ex[c].sum()) for c in df_ex.columns if c != "Mes de Creación"}}
    return pd.concat([df_ex, pd.DataFrame([fila_tot_ex])], ignore_index=True)

if es_nacional:
    mat_1 = construir_df_matriz_14b(df_cal[df_cal["Es_Sede_Central"]])
    mat_2 = construir_df_matriz_14b(df_cal[~df_cal["Es_Sede_Central"]])
else:
    mat_1 = construir_df_matriz_14b(df_cal)
    mat_2 = pd.DataFrame()

# ==============================================================================
# RENDERIZADO CONDICIONAL: SOLO DETALLE O DASHBOARD CON KPIs CLICABLES
# ==============================================================================
kpi_sel = st.session_state["kpi_activo"]

if kpi_sel is not None:
    if kpi_sel in ("TOTAL", "TRAMITE", "RESUELTO", "JEF_TOTAL", "JEF_TRAMITE", "JEF_RESUELTO"):
        mapa_kpis_carga = {
            "TOTAL": (f"Detalle de Carga Procesal Total — {nombre_ambito}", df_act),
            "TRAMITE": (f"Detalle de Expedientes En Trámite — {nombre_ambito}", df_tram),
            "RESUELTO": (f"Detalle de Resuelto - Impug. Tránsito — {nombre_ambito}", df_imp),
            "JEF_TOTAL": (f"Detalle de Expedientes en Jefatura (Total) — {nombre_ambito}", df_jef_all),
            "JEF_TRAMITE": (f"Detalle de Jefatura En Trámite — {nombre_ambito}", df_jef_all[df_jef_all["Estado_Grupo"] == "Trámite"]),
            "JEF_RESUELTO": (f"Detalle de Jefatura Resuelto - Impugnación — {nombre_ambito}", df_jef_all[df_jef_all["Estado_Grupo"] == "Res. Impug."]),
        }
        tit_kpi, df_kpi = mapa_kpis_carga[kpi_sel]
        mostrar_pantalla_detalle(df_kpi, tit_kpi, kpi_sel)
    elif kpi_sel == "PROD_ACUMULADA":
        mostrar_pantalla_detalle_prod(
            df_pr_f,
            f"Detalle de Producción Acumulada — {nombre_ambito}",
            kpi_sel
        )
    elif kpi_sel == "PRESC_ATENCION_PREF":
        mostrar_pantalla_detalle_presc(
            df_p_atencion_pref,
            f"Detalle de Riesgo de Prescripción con Atención Preferencial — {nombre_ambito}",
            kpi_sel
        )
    elif kpi_sel.startswith("PRESC_RANGO_"):
        mapa_kpis_presc = {
            cod_k: (
                f"Detalle — Rango de Riesgo de Expedientes Prescritos y por Prescribir: {clave_rg} — {nombre_ambito}",
                df_p_act[df_p_act["Rango_Presc"] == clave_rg]
            )
            for cod_k, clave_rg, _, _, _, _ in config_colores_rangos
        }
        tit_p, df_p_sub = mapa_kpis_presc[kpi_sel]
        mostrar_pantalla_detalle_presc(df_p_sub, tit_p, kpi_sel)
else:
    tab1, tab_presc, tab2, tab3, tab4 = st.tabs([
        "📊 Sec. 1.1 y 1.2 (Carga y Etapas)",
        "⏳ Sec. 1.3 (Rango de Riesgo de Prescripción)",
        "📅 Sec. 1.4 (Calificación por Año y Mes)",
        f"🏛️ Sec. 1.5 (En Jefatura — {dist_sel if not es_nacional else 'ANC / ODANC'})",
        "⚖️ Sec. 2 (Producción Mensual y Pronunciamientos)"
    ])

    with tab1:
        st.session_state["vista_tab"] = "tab1"
        render_kpis_carga()

        st.subheader(f"1.1. Según la antigüedad por AÑO de creación / ingreso ({sub_titulo_sec})")
        mostrar_tabla_centrada(t_1_1)

        st.markdown("---")
        st.subheader(f"1.2. Según ETAPAS PROCESALES y Sub-etapas de Calificación ({sub_titulo_sec})")
        mostrar_tabla_centrada(t_1_2_etapas)

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
    with tab_presc:
        tot_p_sem = len(df_p_act)
        tot_p_vig = len(df_p_vigente)
        p_anc_vig = int(df_p_vigente["Es_Sede_Central"].sum())
        p_oda_vig = int((~df_p_vigente["Es_Sede_Central"]).sum())

        tot_ap = len(df_p_atencion_pref)
        cnt_m60 = int((df_p_atencion_pref["Atencion_Pref_Norm"].str.contains("60", na=False)).sum())
        cnt_m75 = int((df_p_atencion_pref["Atencion_Pref_Norm"].str.contains("75", na=False)).sum())
        if cnt_m60 > 0 or cnt_m75 > 0:
            sub_ap = f"↑ +60 años: {cnt_m60:,} | +75 años: {cnt_m75:,}"
        else:
            pct_ap = (tot_ap / tot_p_vig * 100) if tot_p_vig > 0 else 0.0
            sub_ap = f"↑ {pct_ap:.1f}% del riesgo vigente"

        etapa_top_p = df_p_vigente["Etapa_Presc_Cat"].value_counts() if not df_p_vigente.empty else df_p_act["Etapa_Presc_Cat"].value_counts()
        et_top_nom = etapa_top_p.index[0] if not etapa_top_p.empty else "Sin datos"
        et_top_cnt = int(etapa_top_p.iloc[0]) if not etapa_top_p.empty else 0
        base_et_pct = max(1, tot_p_vig if not df_p_vigente.empty else tot_p_sem)
        pk1, pk2, pk3, pk4 = st.columns(4)
        pk1.metric(
            "TOTAL RIESGO PRESC. (1-240d)",
            f"{tot_p_vig:,} exp.",
            f"ANC: {p_anc_vig:,} | ODANC: {p_oda_vig:,}" if es_nacional else f"Sede: {dist_sel}"
        )
        with pk2:
            if tot_ap > 0:
                st.button(
                    f"ATENCIÓN PREFERENCIAL\n**{tot_ap:,} exp.**\n*{sub_ap}*",
                    key="btn_presc_atencion_pref",
                    on_click=alternar_kpi,
                    args=("PRESC_ATENCION_PREF",),
                    use_container_width=True
                )
            else:
                st.metric(
                    "ATENCIÓN PREFERENCIAL",
                    "0 exp.",
                    "Sin atención preferencial",
                    delta_color="off"
                )
        pk3.metric(
            "ETAPA CON MAYOR RIESGO",
            f"{et_top_nom}",
            f"{et_top_cnt:,} exp. ({et_top_cnt/base_et_pct*100:.1f}%)"
        )
        if es_nacional:
            top1_dist = df_p_vigente[~df_p_vigente["Es_Sede_Central"]]["Distrito_Norm"].value_counts()
            d_top_nom = top1_dist.index[0].title() if not top1_dist.empty else "-"
            d_top_cnt = int(top1_dist.iloc[0]) if not top1_dist.empty else 0
            pk4.metric(
                "ODANC CON MAYOR RIESGO",
                f"{d_top_nom}",
                f"{d_top_cnt:,} exp. en riesgo (1-240d)"
            )
        else:
            df_p_ord = df_p_vigente.sort_values("Dias_Calc", ascending=True)
            if not df_p_ord.empty:
                fila_min = df_p_ord.iloc[0]
                dias_min = int(fila_min["Dias_Calc"])
                detalle_exp_limpio = formatear_detalle_exp_presc(fila_min)
                unidad_txt = "día" if dias_min == 1 else "días"
                pk4.metric(
                    "EXP. MÁS PRÓXIMO A VENCER",
                    f"Vence ({dias_min} {unidad_txt})",
                    f"{detalle_exp_limpio}"
                )
            else:
                pk4.metric("EXP. MÁS PRÓXIMO A VENCER", "Sin exp. por vencer", "-", delta_color="off")

        st.markdown("<div style='margin-bottom: 12px;'></div>", unsafe_allow_html=True)

        # 6 Tarjetas individuales CLICABLES por rango
        conteos_rangos = df_p_act.groupby("Rango_Presc").size().reindex(ord_rangos, fill_value=0)
        sem_cols = st.columns(6)
        for col_s, (cod_k, clave_rg, titulo_rg, rango_txt, _, emoji) in zip(sem_cols, config_colores_rangos):
            val_s = int(conteos_rangos.loc[clave_rg])
            pct_s = (val_s / tot_p_sem * 100) if tot_p_sem > 0 else 0.0
            with col_s:
                st.button(
                    f"{emoji} {titulo_rg}\n**{val_s:,} exp.**\n*{rango_txt} ({pct_s:.2f}%)*",
                    key=f"btn_{cod_k}",
                    on_click=alternar_kpi,
                    args=(cod_k,),
                    use_container_width=True
                )

        st.subheader(f"1.3.A. Rango de Riesgo de Expedientes Prescritos y por Prescribir ({sub_titulo_sec})")
        mostrar_tabla_centrada(t_1_3_presc)

        st.markdown("---")
        st.subheader(f"1.3.B. Riesgo de prescripción según ETAPA PROCESAL ({sub_titulo_sec})")
        mostrar_tabla_centrada(t_1_3_et)

        if es_nacional:
            top5_p = df_p_vigente[~df_p_vigente["Es_Sede_Central"]]["Distrito_Norm"].value_counts().head(5)
            if not top5_p.empty:
                txt_top5 = ", ".join([f"{d.title()} ({c:,})" for d, c in top5_p.items()])
                st.caption(f"*Mayor riesgo vigente ODANC (1-240d): {txt_top5}.")
    with tab2:
        cal_tr = int((df_cal["Estado_Grupo"] == "Trámite").sum())
        cal_im = int((df_cal["Estado_Grupo"] == "Res. Impug.").sum())
        cal_anc = int(df_cal["Es_Sede_Central"].sum())
        cal_oda = int((~df_cal["Es_Sede_Central"]).sum())

        if es_nacional:
            ck1, ck2, ck3, ck4, ck5 = st.columns(5)
            ck1.metric("TOTAL EN CALIFICACIÓN", f"{len(df_cal):,} exp.", f"{len(df_cal)/max(1,len(df_act))*100:.1f}% de la carga total")
            ck2.metric("CALIFICACIÓN EN TRÁMITE", f"{cal_tr:,} exp.", f"{cal_tr/max(1,len(df_cal))*100:.1f}% de calificación")
            ck3.metric("CALIF. RES. IMPUGNACIÓN", f"{cal_im:,} exp.", f"{cal_im/max(1,len(df_cal))*100:.1f}% de calificación")
            ck4.metric("SEDE CENTRAL vs. ODANC", f"ANC: {cal_anc:,}", f"ODANC: {cal_oda:,} exp.")
            with ck5:
                render_tarjeta_mas_antiguo(df_cal, es_nacional, "EXP. MÁS ANTIGUO")
        else:
            ck1, ck2, ck3, ck4 = st.columns(4)
            ck1.metric("TOTAL EN CALIFICACIÓN", f"{len(df_cal):,} exp.", f"{len(df_cal)/max(1,len(df_act))*100:.1f}% de la carga total")
            ck2.metric("CALIFICACIÓN EN TRÁMITE", f"{cal_tr:,} exp.", f"{cal_tr/max(1,len(df_cal))*100:.1f}% de calificación")
            ck3.metric("CALIF. RES. IMPUGNACIÓN", f"{cal_im:,} exp.", f"{cal_im/max(1,len(df_cal))*100:.1f}% de calificación")
            with ck4:
                render_tarjeta_mas_antiguo(df_cal, es_nacional, "EXP. MÁS ANTIGUO")
        st.markdown("<div style='margin-bottom: 10px;'></div>", unsafe_allow_html=True)

        st.subheader(f"1.4.A. Calificación por Año de Creación ({sub_titulo_sec})")
        mostrar_tabla_centrada(t_1_4_a)

        def dibujar_html_matriz_mes_anio(df_in, titulo_mat):
            st.markdown(f"**{titulo_mat}**")
            df_v = df_in.dropna(subset=["Fecha_Creacion_Dt"]).copy()
            if df_v.empty:
                st.info("Sin registros con fecha válida.")
                return
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

        st.markdown("---")
        if es_nacional:
            dibujar_html_matriz_mes_anio(df_cal[df_cal["Es_Sede_Central"]], "1.4.B. Calificación por Mes y Año — ANC Sede Central")
            dibujar_html_matriz_mes_anio(df_cal[~df_cal["Es_Sede_Central"]], "1.4.B. Calificación por Mes y Año — ODANCs (34 Distritos)")
        else:
            dibujar_html_matriz_mes_anio(df_cal, f"1.4.B. Calificación por Mes y Año — {dist_sel}")

        st.markdown("---")
        st.subheader(f"1.4.C. Calificación según Ubicación / Unidad y Sub-etapa ({sub_titulo_sec})")
        mostrar_tabla_centrada(t_1_4_c)

    with tab3:
        jef_tot_cnt = len(df_jef_all)
        jef_tr = int((df_jef_all["Estado_Grupo"] == "Trámite").sum())
        jef_im = int((df_jef_all["Estado_Grupo"] == "Res. Impug.").sum())
        jef_anc = int(df_jef_all["Es_Sede_Central"].sum())
        jef_oda = int((~df_jef_all["Es_Sede_Central"]).sum())

        def render_botones_jefatura(c1, c2, c3):
            with c1:
                st.button(
                    f"TOTAL EN JEFATURA\n**{jef_tot_cnt:,} exp.**\n*↑ {jef_tot_cnt/max(1,len(df_act))*100:.1f}% de la carga total*",
                    key="kpi_btn_jef_total",
                    on_click=alternar_kpi,
                    args=("JEF_TOTAL",),
                    use_container_width=True
                )
            with c2:
                st.button(
                    f"JEFATURA EN TRÁMITE\n**{jef_tr:,} exp.**\n*↑ {jef_tr/max(1,jef_tot_cnt)*100:.1f}% en jefatura*",
                    key="kpi_btn_jef_tramite",
                    on_click=alternar_kpi,
                    args=("JEF_TRAMITE",),
                    use_container_width=True
                )
            with c3:
                st.button(
                    f"JEFATURA RES. IMPUG.\n**{jef_im:,} exp.**\n*↑ {jef_im/max(1,jef_tot_cnt)*100:.1f}% en jefatura*",
                    key="kpi_btn_jef_resuelto",
                    on_click=alternar_kpi,
                    args=("JEF_RESUELTO",),
                    use_container_width=True
                )

        if es_nacional:
            jk1, jk2, jk3, jk4, jk5 = st.columns(5)
            render_botones_jefatura(jk1, jk2, jk3)
            jk4.metric("JEFATURA ANC vs. ODANC", f"ANC: {jef_anc:,}", f"ODANC: {jef_oda:,} exp.")
            with jk5:
                render_tarjeta_mas_antiguo(df_jef_all, es_nacional, "EXP. MÁS ANTIGUO")
        else:
            jk1, jk2, jk3, jk4 = st.columns(4)
            render_botones_jefatura(jk1, jk2, jk3)
            with jk4:
                render_tarjeta_mas_antiguo(df_jef_all, es_nacional, "EXP. MÁS ANTIGUO")

        st.markdown("<div style='margin-bottom: 10px;'></div>", unsafe_allow_html=True)

        st.subheader(f"1.5.A. En Jefatura por Etapa y Estado ({sub_titulo_sec})")
        mostrar_tabla_centrada(t_1_5)

        st.markdown("---")
        st.subheader(f"1.5.B. En Jefatura según Etapa, Año de Creación y Estado ({sub_titulo_sec})")
        if t_1_5_b.empty:
            st.info("Sin registros en Jefatura para desagregar por año y estado.")
        else:
            th_tr_15 = "".join([f"<th style='padding:6px; border:1px solid #334155; text-align:center;'>{c}</th>" for c in cols_15b_tr])
            th_im_15 = "".join([f"<th style='padding:6px; border:1px solid #334155; text-align:center;'>{c}</th>" for c in cols_15b_im])

            html_15b = f"""
            <div style="overflow-x:auto; margin-bottom:16px; border-radius:8px; border:1px solid #334155;">
            <table style="width:100%; border-collapse:collapse; font-size:13px; color:#f8fafc; background:#111827; text-align:center;">
                <thead>
                    <tr style="background:#1e293b; border-bottom:1px solid #334155;">
                        <th rowspan="2" style="padding:8px 10px; border:1px solid #334155; vertical-align:middle; text-align:left;">Etapa en Jefatura</th>
                        <th colspan="{len(cols_15b_tr)}" style="padding:6px; border:1px solid #334155; text-align:center; color:#38bdf8;">TRÁMITE</th>
                        <th colspan="{len(cols_15b_im)}" style="padding:6px; border:1px solid #334155; text-align:center; color:#f59e0b;">RESUELTO - IMPUGNACIÓN</th>
                        <th rowspan="2" style="padding:8px; border:1px solid #334155; vertical-align:middle; text-align:center; background:#0f172a;">Total general</th>
                    </tr>
                    <tr style="background:#1e293b; border-bottom:2px solid #475569;">
                        {th_tr_15}
                        {th_im_15}
                    </tr>
                </thead>
                <tbody>
            """
            for idx, r in t_1_5_b.iterrows():
                et_nom = str(r["Etapa en Jefatura"])
                es_tot_r = "TOTAL" in et_nom.upper()
                bg_r = "#1e293b" if es_tot_r else ("#0f172a" if idx % 2 == 1 else "#111827")
                fw_r = "700" if es_tot_r else "400"

                tds_tr_r = "".join([f"<td style='padding:6px; border:1px solid #1f2937; font-weight:{fw_r};'>{int(r[f'TRÁMITE - {c}']):,}</td>" for c in cols_15b_tr])
                tds_im_r = "".join([f"<td style='padding:6px; border:1px solid #1f2937; font-weight:{fw_r};'>{int(r[f'RES. IMPUG. - {c}']):,}</td>" for c in cols_15b_im])
                color_tot_c = "color:#fbbf24;" if es_tot_r else ""

                html_15b += (
                    f"<tr style='background:{bg_r};'>"
                    f"<td style='padding:7px 10px; border:1px solid #1f2937; text-align:left; font-weight:{'700' if es_tot_r else '500'};'>{et_nom}</td>"
                    f"{tds_tr_r}{tds_im_r}"
                    f"<td style='padding:7px; border:1px solid #1f2937; font-weight:700; {color_tot_c}'>{int(r['Total general']):,}</td>"
                    f"</tr>"
                )
            html_15b += "</tbody></table></div>"
            st.markdown(html_15b, unsafe_allow_html=True)

    with tab4:
        inf_mes_pr = int(df_pr_mes_act["Es_Informe"].sum())
        res_mes_pr = len(df_pr_mes_act) - inf_mes_pr
        pr_anc_mes = int(df_pr_mes_act["Es_Sede_Central"].sum())
        pr_oda_mes = int((~df_pr_mes_act["Es_Sede_Central"]).sum())

        prk1, prk2, prk3, prk4 = st.columns(4)
        prk1.metric(
            f"PRODUCCIÓN {mes_act_lbl.upper()}",
            f"{len(df_pr_mes_act):,} exp.",
            f"ANC: {pr_anc_mes:,} | ODANC: {pr_oda_mes:,}" if es_nacional else f"Sede: {dist_sel}"
        )
        prk2.metric(
            f"RESOLUCIONES ({mes_act_lbl.upper()})",
            f"{res_mes_pr:,} res.",
            f"{res_mes_pr/max(1,len(df_pr_mes_act))*100:.1f}% del mes actual"
        )
        prk3.metric(
            f"INFORMES ({mes_act_lbl.upper()})",
            f"{inf_mes_pr:,} inf.",
            f"{inf_mes_pr/max(1,len(df_pr_mes_act))*100:.1f}% del mes actual"
        )
        with prk4:
            if not fechas_pr.empty and per_act is not None and per_ant is not None:
                cnt_mes_ant = int((df_pr_f["Periodo_M"] == per_ant).sum())
                sub_pr_acum = f"↑ Mes ant.: {cnt_mes_ant:,} | Mes act.: {len(df_pr_mes_act):,}"
            else:
                sub_pr_acum = "↑ Ver detalle acumulado"

            st.button(
                f"PRODUCCIÓN ACUMULADA\n**{len(df_pr_f):,} exp.**\n*{sub_pr_acum}*",
                key="kpi_btn_prod_acumulada",
                on_click=alternar_kpi,
                args=("PROD_ACUMULADA",),
                use_container_width=True
            )

        st.markdown("<div style='margin-bottom: 10px;'></div>", unsafe_allow_html=True)

        st.subheader(f"2.1. Comparativo Mensual de Producción: Mes Anterior vs. Mes Actual ({sub_titulo_sec})")
        if not t_2_1.empty:
            mostrar_tabla_centrada(t_2_1)

        st.markdown("---")
        st.subheader(f"2.2. Producción por Tipo de Pronunciamiento y/o Sanción — {mes_act_lbl} ({sub_titulo_sec})")
        if not t_2_2.empty:
            mostrar_tabla_centrada(t_2_2)

        st.markdown("---")
        st.subheader(f"2.3. Producción por Unidad — {mes_act_lbl} ({'TODOS' if es_nacional else dist_sel})")
        if not t_2_3.empty:
            mostrar_tabla_centrada(t_2_3)

        st.markdown("---")
        st.subheader(f"2.4. Producción por Magistrado — {mes_act_lbl} ({'TODOS' if es_nacional else dist_sel})")
        if not t_2_4.empty:
            mostrar_tabla_centrada(t_2_4)

    buffer_anexo = io.BytesIO()
    with pd.ExcelWriter(buffer_anexo, engine="openpyxl") as writer:
        t_1_1.to_excel(writer, sheet_name="1_1_Anio_Creacion", index=False)
        t_1_2_etapas.to_excel(writer, sheet_name="1_2_Etapas_Subetapas", index=False)
        t_1_3_presc.to_excel(writer, sheet_name="1_3A_Rango_Riesgo_Presc", index=False)
        t_1_3_et.to_excel(writer, sheet_name="1_3B_Presc_Etapas", index=False)
        t_1_4_a.to_excel(writer, sheet_name="1_4A_Calif_Anio", index=False)
        if not mat_1.empty:
            mat_1.to_excel(writer, sheet_name="1_4B_Calif_Mes_1", index=False)
        if not mat_2.empty:
            mat_2.to_excel(writer, sheet_name="1_4B_Calif_Mes_2", index=False)
        t_1_4_c.to_excel(writer, sheet_name="1_4C_Calif_Unidad", index=False)
        t_1_5.to_excel(writer, sheet_name="1_5A_Jefatura_Etapa", index=False)
        if not t_1_5_b.empty:
            t_1_5_b.to_excel(writer, sheet_name="1_5B_Jefatura_Anio_Est", index=False)
        if not t_2_1.empty:
            t_2_1.to_excel(writer, sheet_name="2_1_Produccion_Mensual", index=False)
        if not t_2_2.empty:
            t_2_2.to_excel(writer, sheet_name="2_2_Prod_Pronunciamiento", index=False)
        if not t_2_3.empty:
            t_2_3.to_excel(writer, sheet_name="2_3_Prod_Unidad", index=False)
        if not t_2_4.empty:
            t_2_4.to_excel(writer, sheet_name="2_4_Prod_Magistrado", index=False)

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
        pdf.cell(0, 5, limpiar_latin1(f"ANÁLISIS ESTADÍSTICO DE CARGA POR ETAPAS, PRESCRIPCIONES Y PRODUCCIÓN: {nombre_ambito} | Corte: {fecha_corte.strftime('%d/%m/%Y')} (UIA)"), ln=True)

        pdf.ln(6)

        pdf.set_font("Helvetica", "B", 8.5)
        pdf.set_text_color(30, 41, 59)
        resumen_kpis = (
            f"CARGA PROCESAL: {len(df_act):,} exp.   |   "
            f"EN JEFATURA: {len(df_jef_all):,} exp.   |   "
            f"RIESGO PRESC. (<=240d): {len(df_p_240):,} exp.   |   "
            f"PROD. {mes_act_lbl.upper()}: {len(df_pr_mes_act):,} exp.   |   "
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
            w_primera = 95.0 if n_cols > 3 else 160.0
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
                etiq_col = str(c)
                if len(etiq_col) > 36 and i == 0:
                    etiq_col = "Rango de Riesgo Presc. / Por Presc."
                pdf.cell(anchos[i], 6.5, limpiar_latin1(etiq_col[:38]), border=1, align=alin, fill=True)
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
        dibujar_tabla_pdf(f"1.2. Etapas Procesales y Sub-etapas de Calificación ({sub_titulo_sec})", t_1_2_etapas)
        dibujar_tabla_pdf(f"1.3.A. Rango de Riesgo de Expedientes Prescritos y por Prescribir ({sub_titulo_sec})", t_1_3_presc)
        dibujar_tabla_pdf(f"1.3.B. Riesgo de Prescripción por Etapa Procesal ({sub_titulo_sec})", t_1_3_et)
        dibujar_tabla_pdf(f"1.4.A. Calificación por Año de Creación ({sub_titulo_sec})", t_1_4_a)

        def dibujar_matriz_pdf(titulo_m, df_m):
            if df_m is None or df_m.empty:
                return
            cols_m = df_m.columns.tolist()
            n_c = len(cols_m)
            w_mes = 36.0
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
                    alin_m = "L" if i == 0 else "C"
                    pdf.cell(anchos_m[i], 5.2, limpiar_latin1(txt[:22]), border=1, align=alin_m, fill=True)
                pdf.ln()
            pdf.ln(3)

        if es_nacional:
            dibujar_matriz_pdf("1.4.B. Calificación por Mes y Año - ANC Sede Central", mat_1)
            dibujar_matriz_pdf("1.4.B. Calificación por Mes y Año - ODANCs (34 Distritos)", mat_2)
        else:
            dibujar_matriz_pdf(f"1.4.B. Calificación por Mes y Año - {dist_sel}", mat_1)

        dibujar_tabla_pdf(f"1.4.C. Calificación según Ubicación / Unidad y Sub-etapa ({sub_titulo_sec})", t_1_4_c, max_filas=100)
        dibujar_tabla_pdf(f"1.5.A. Expedientes en Jefatura por Etapa y Estado ({sub_titulo_sec})", t_1_5)
        if not t_1_5_b.empty:
            dibujar_matriz_pdf(f"1.5.B. En Jefatura según Etapa, Año de Creación y Estado ({sub_titulo_sec})", t_1_5_b)
        if not t_2_1.empty:
            dibujar_tabla_pdf(f"2.1. Comparativo Mensual de Producción ({sub_titulo_sec})", t_2_1)
        if not t_2_2.empty:
            dibujar_tabla_pdf(f"2.2. Producción por Tipo de Pronunciamiento / Sanción - {mes_act_lbl}", t_2_2)
        if not t_2_3.empty:
            dibujar_tabla_pdf(f"2.3. Producción por Unidad - {mes_act_lbl} ({'TODOS' if es_nacional else dist_sel})", t_2_3, max_filas=60)
        if not t_2_4.empty:
            dibujar_tabla_pdf(f"2.4. Producción por Magistrado - {mes_act_lbl} ({'TODOS' if es_nacional else dist_sel})", t_2_4, max_filas=60)

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