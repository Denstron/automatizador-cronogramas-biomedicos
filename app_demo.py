"""
AUTOMATIZADOR DE CRONOGRAMAS BIOMÉDICOS V5
==========================================
App web con Streamlit
- Dashboard interactivo
- Filtros por área
- Marcar equipos como completados
- Descargar reportes
"""

import streamlit as st
import pandas as pd
from datetime import datetime
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# ============================================================
# CONFIGURACIÓN
# ============================================================
st.set_page_config(
    page_title="Sistema de Mantenimiento Biomédico",
    page_icon="🏥",
    layout="wide"
)

ARCHIVO_ENTRADA = "Cronograma_DEMO.xlsx"

MESES = {"ENERO":1,"FEBRERO":2,"MARZO":3,"ABRIL":4,"MAYO":5,"JUNIO":6,
    "JULIO":7,"AGOSTO":8,"SEPTIEMBRE":9,"OCTUBRE":10,"NOVIEMBRE":11,"DICIEMBRE":12}
MESES_INVERSO = {v: k for k, v in MESES.items()}

HOY = datetime.today()
MES_ACTUAL = HOY.month
NOMBRE_MES = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
              7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}

# ============================================================
# FUNCIONES
# ============================================================

@st.cache_data
def cargar_datos():
    df = pd.read_excel(ARCHIVO_ENTRADA, header=0)
    df.columns = ["ITEM","DESCRIPCION","UBICACION","N_INVENTARIO","FABRICANTE","MARCA",
                  "MODELO","SERIE","CLASIFICACION_RIESGO","REGISTRO_INVIMA",
                  "FECHA_COMPRA","PERIODICIDAD","FECHA_MANTENIMIENTO"]
    df = df.dropna(subset=["DESCRIPCION"])
    df.replace("NR","",inplace=True)

    def extraer_meses(texto):
        if pd.isna(texto) or str(texto).strip()=="": return []
        return [MESES[p.strip()] for p in str(texto).upper().split("/") if p.strip() in MESES]

    def calcular_estado(meses):
        if not meses: return "SIN FECHA", None
        meses = sorted(meses)
        proximo = next((m for m in meses if m >= MES_ACTUAL), None)
        if proximo is None: return "🔴 VENCIDO", meses[-1]
        diff = proximo - MES_ACTUAL
        if diff == 0: return "🔴 ESTE MES", proximo
        elif diff == 1: return "🟡 PRÓXIMO MES", proximo
        elif diff <= 2: return "🟠 EN 2 MESES", proximo
        else: return "🟢 OK", proximo

    df["MESES_PARSED"] = df["FECHA_MANTENIMIENTO"].apply(extraer_meses)
    resultados = df["MESES_PARSED"].apply(calcular_estado)
    df["ESTADO"] = resultados.apply(lambda x: x[0])
    df["PROXIMO_MES"] = resultados.apply(lambda x: MESES_INVERSO.get(x[1],"") if x[1] else "")
    df["COMPLETADO"] = False

    return df

def generar_excel_descarga(df):
    """Genera Excel en memoria para descarga"""
    output = io.BytesIO()
    columnas = ["ITEM","DESCRIPCION","UBICACION","N_INVENTARIO",
                "PERIODICIDAD","FECHA_MANTENIMIENTO","PROXIMO_MES","ESTADO"]
    df[columnas].to_excel(output, index=False, sheet_name="CRONOGRAMA")
    output.seek(0)
    return output

# ============================================================
# CARGAR DATOS
# ============================================================
try:
    df = cargar_datos()
except FileNotFoundError:
    st.error(f"❌ No se encontró el archivo: {ARCHIVO_ENTRADA}")
    st.info("Asegúrate de que el Excel esté en la misma carpeta que app_v5.py")
    st.stop()

areas = sorted(df["UBICACION"].dropna().unique().tolist())

# ============================================================
# INICIALIZAR ESTADO DE SESIÓN (para marcar completados)
# ============================================================
if "completados" not in st.session_state:
    st.session_state.completados = set()

# ============================================================
# HEADER
# ============================================================
col1, col2 = st.columns([3, 1])
with col1:
    st.title("🏥 Sistema de Mantenimiento Biomédico")
    st.caption(f"Hospital Demo | {HOY.strftime('%d/%m/%Y')}")
with col2:
    st.metric("Total Equipos", len(df))

st.divider()

# ============================================================
# MÉTRICAS PRINCIPALES
# ============================================================
c1, c2, c3, c4 = st.columns(4)

este_mes    = len(df[df["ESTADO"]=="🔴 ESTE MES"])
proximo_mes = len(df[df["ESTADO"]=="🟡 PRÓXIMO MES"])
dos_meses   = len(df[df["ESTADO"]=="🟠 EN 2 MESES"])
ok          = len(df[df["ESTADO"]=="🟢 OK"])
completados_count = len(st.session_state.completados)

with c1:
    st.metric("🔴 Este mes", este_mes,
              delta=f"-{completados_count} completados" if completados_count else None,
              delta_color="normal")
with c2:
    st.metric("🟡 Próximo mes", proximo_mes)
with c3:
    st.metric("🟠 En 2 meses", dos_meses)
with c4:
    st.metric("🟢 OK", ok)

st.divider()

# ============================================================
# BARRA LATERAL — FILTROS
# ============================================================
st.sidebar.title("🔍 Filtros")

area_seleccionada = st.sidebar.selectbox(
    "Área / Ubicación",
    ["TODAS LAS ÁREAS"] + areas
)

estado_seleccionado = st.sidebar.multiselect(
    "Estado",
    ["🔴 ESTE MES", "🟡 PRÓXIMO MES", "🟠 EN 2 MESES", "🟢 OK", "SIN FECHA"],
    default=["🔴 ESTE MES", "🟡 PRÓXIMO MES"]
)

busqueda = st.sidebar.text_input("🔎 Buscar equipo", placeholder="Ej: monitor, bomba...")

st.sidebar.divider()
st.sidebar.metric("✅ Completados hoy", completados_count)

if st.sidebar.button("🔄 Reiniciar completados"):
    st.session_state.completados = set()
    st.rerun()

# ============================================================
# APLICAR FILTROS
# ============================================================
df_filtrado = df.copy()

if area_seleccionada != "TODAS LAS ÁREAS":
    df_filtrado = df_filtrado[df_filtrado["UBICACION"] == area_seleccionada]

if estado_seleccionado:
    df_filtrado = df_filtrado[df_filtrado["ESTADO"].isin(estado_seleccionado)]

if busqueda:
    df_filtrado = df_filtrado[
        df_filtrado["DESCRIPCION"].str.contains(busqueda, case=False, na=False)
    ]

# ============================================================
# TABLA PRINCIPAL CON CHECKBOXES
# ============================================================
st.subheader(f"📋 Equipos — {area_seleccionada} ({len(df_filtrado)} equipos)")

if len(df_filtrado) == 0:
    st.info("No hay equipos con los filtros seleccionados.")
else:
    # Mostrar tabla con checkbox por equipo
    for idx, row in df_filtrado.iterrows():
        col_check, col_inv, col_equipo, col_area, col_periodo, col_estado = st.columns([0.5, 1, 3, 2, 1.5, 1.5])

        inventario = str(row["N_INVENTARIO"])
        completado = inventario in st.session_state.completados

        with col_check:
            if st.checkbox("", value=completado, key=f"check_{idx}"):
                st.session_state.completados.add(inventario)
            else:
                st.session_state.completados.discard(inventario)

        with col_inv:
            st.write(f"**{row['N_INVENTARIO']}**")
        with col_equipo:
            texto = f"~~{row['DESCRIPCION']}~~" if completado else row['DESCRIPCION']
            st.write(texto)
        with col_area:
            st.write(row['UBICACION'])
        with col_periodo:
            st.write(row['PERIODICIDAD'])
        with col_estado:
            estado = row['ESTADO']
            if "ESTE MES" in estado or "VENCIDO" in estado:
                st.error(estado)
            elif "PRÓXIMO" in estado:
                st.warning(estado)
            elif "2 MESES" in estado:
                st.warning(estado)
            else:
                st.success(estado)

# ============================================================
# RESUMEN POR ÁREA
# ============================================================
st.divider()
st.subheader("📊 Resumen por Área")

resumen_data = []
for area in areas:
    df_a = df[df["UBICACION"]==area]
    resumen_data.append({
        "Área": area,
        "🔴 Este mes": int((df_a["ESTADO"]=="🔴 ESTE MES").sum()),
        "🟡 Próximo": int((df_a["ESTADO"]=="🟡 PRÓXIMO MES").sum()),
        "🟠 2 meses": int((df_a["ESTADO"]=="🟠 EN 2 MESES").sum()),
        "🟢 OK": int((df_a["ESTADO"]=="🟢 OK").sum()),
        "Total": len(df_a)
    })

df_resumen = pd.DataFrame(resumen_data)
st.dataframe(df_resumen, use_container_width=True, hide_index=True)

# ============================================================
# DESCARGAR REPORTE
# ============================================================
st.divider()
st.subheader("📥 Descargar Reporte")

excel_data = generar_excel_descarga(df)
st.download_button(
    label="⬇️ Descargar Excel completo",
    data=excel_data,
    file_name=f"Reporte_Mantenimiento_{HOY.strftime('%Y%m%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ============================================================
# FOOTER
# ============================================================
st.divider()
st.caption("Sistema de Gestión de Mantenimiento Biomédico v5.0 | Desarrollado por Daniel España Vargas — demo.streamlit.app | Ingeniería Biomédica")
