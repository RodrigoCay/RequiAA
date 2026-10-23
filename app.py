# -*- coding: utf-8 -*-
"""
Created on Thu Jul  9 17:49:16 2026

@author: rcayuleo
"""

import streamlit as st
import pandas as pd
import io
from pathlib import Path

# Archivos del proyecto
from config import *
from data_loader import *
from kpis import *

# =====================================================
# EXPORTAR DATAFRAME A EXCEL (bytes)
# =====================================================
# Usado por los botones de descarga que ofrecen .xlsx en
# vez de .csv.

def exportar_excel(df, nombre_hoja="Datos"):

    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:

        df.to_excel(writer, index=False, sheet_name=nombre_hoja)

    return buffer.getvalue()

# =====================================================
# ESTILO DE GRÁFICOS (PLOTLY) - FONDO BLANCO / MODO PRESENTACIÓN
# =====================================================
# Todos los gráficos pasan por estilo_grafico() para que se
# vean iguales: fondo blanco, texto negro y letras grandes
# para que se lean bien al proyectar.

import plotly.express as px

TEXTO_GRAFICO = "#111111"

GRILLA_GRAFICO = "#E3E7ED"

# Paleta de colores intensos: se leen bien sobre blanco y
# permiten texto blanco encima (sin tonos pastel).
PALETA = [
    "#0B4F8A",  # azul
    "#C0392B",  # rojo
    "#1E8449",  # verde
    "#D68910",  # naranjo
    "#6C3483",  # morado
    "#117A8B",  # turquesa
    "#A04000",  # café
    "#34495E",  # gris azulado
    "#B03A7A",  # fucsia
    "#7D8C10",  # oliva
]

px.defaults.template = "plotly_white"

px.defaults.color_discrete_sequence = PALETA


def estilo_grafico(fig, titulo, altura=480):

    fig.update_layout(

        title=dict(
            text=titulo,
            font=dict(size=24, color=TEXTO_GRAFICO),
            x=0,
            xanchor="left"
        ),

        font=dict(
            family="Segoe UI, Arial, sans-serif",
            size=16,
            color=TEXTO_GRAFICO
        ),

        paper_bgcolor="#FFFFFF",

        plot_bgcolor="#FFFFFF",

        height=altura,

        margin=dict(l=10, r=20, t=80, b=10),

        legend=dict(font=dict(size=16, color=TEXTO_GRAFICO)),

        hoverlabel=dict(font_size=16),

        piecolorway=PALETA

    )

    ejes = dict(
        tickfont=dict(size=15, color=TEXTO_GRAFICO),
        title_font=dict(size=17, color=TEXTO_GRAFICO),
        gridcolor=GRILLA_GRAFICO,
        linecolor="#9AA3AF",
        automargin=True
    )

    fig.update_xaxes(**ejes)

    fig.update_yaxes(**ejes)

    # Barras: etiqueta fuera de la barra, en negro (si va dentro
    # en blanco, en barras cortas no se alcanza a leer).
    fig.update_traces(
        selector=dict(type="bar"),
        textposition="outside",
        textfont=dict(size=17, color=TEXTO_GRAFICO),
        constraintext="none",
        cliponaxis=False
    )

    # Deja espacio extra sobre/al lado de la barra más grande
    # para que su etiqueta no quede cortada.
    barras = [t for t in fig.data if t.type == "bar"]

    if barras:

        horizontal = barras[0].orientation == "h"

        valores = pd.to_numeric(
            pd.Series([
                v for t in barras
                for v in ((t.x if horizontal else t.y) if (t.x if horizontal else t.y) is not None else [])
            ]),
            errors="coerce"
        )

        maximo = valores.max()

        if pd.notna(maximo) and maximo > 0:

            eje = fig.update_xaxes if horizontal else fig.update_yaxes

            eje(range=[0, maximo * 1.22])

    # Donas: porcentaje en blanco sobre colores intensos; los
    # nombres van en la leyenda.
    fig.update_traces(
        selector=dict(type="pie"),
        textinfo="percent",
        textposition="inside",
        insidetextorientation="horizontal",
        textfont=dict(size=18, color="#FFFFFF"),
        marker=dict(line=dict(color="#FFFFFF", width=2))
    )

    # Líneas: más gruesas y con marcadores grandes.
    fig.update_traces(
        selector=dict(type="scatter"),
        line=dict(width=4),
        marker=dict(size=10)
    )

    if any(t.type == "pie" for t in fig.data):

        fig.update_layout(
            uniformtext_minsize=14,
            uniformtext_mode="hide"
        )

    return fig


# =====================================================
# AYUDAS PARA GRÁFICOS
# =====================================================

import textwrap


def envolver(texto, ancho=14):

    # Parte etiquetas largas en varias líneas para no tener
    # que girarlas (el texto inclinado se lee mal al proyectar).
    return "<br>".join(textwrap.wrap(str(texto), ancho)) or str(texto)


def etiquetas_horizontales(fig, categorias, ancho=14):

    fig.update_xaxes(
        tickangle=0,
        tickvals=list(categorias),
        ticktext=[envolver(c, ancho) for c in categorias]
    )

    return fig


def barras_horizontales(datos, categoria, valor, titulo, color=None):

    # Barras acostadas: los nombres largos (zonas, contratos,
    # proveedores) se leen completos y en horizontal.
    fig = px.bar(

        datos,

        x=valor,

        y=categoria,

        orientation="h",

        text="Etiqueta",

        color_discrete_sequence=[color or PALETA[0]]

    )

    fig.update_yaxes(categoryorder="total ascending", title=None)

    # Nombres muy largos se recortan en el eje (el nombre
    # completo aparece al pasar el mouse).
    nombres = list(datos[categoria].astype(str))

    fig.update_yaxes(
        tickvals=nombres,
        ticktext=[n if len(n) <= 30 else n[:28].rstrip() + "…" for n in nombres]
    )

    altura = max(420, 150 + 44 * len(datos))

    return estilo_grafico(fig, titulo, altura=altura)


def mapa_colores(*series):

    # Mismo color para la misma empresa en todos los gráficos.
    valores = sorted(
        set().union(*[set(s.dropna().astype(str).unique()) for s in series])
    )

    return {v: PALETA[i % len(PALETA)] for i, v in enumerate(valores)}


def fmt_fecha(valor):

    if pd.isna(valor):

        return "Pendiente"

    try:

        return pd.to_datetime(valor).strftime(DATE_FORMAT)

    except Exception:

        return str(valor)

# =====================================================
# CONFIGURACIÓN STREAMLIT
# =====================================================

st.set_page_config(
    page_title="Dashboard Compras",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# =====================================================
# CARGAR CSS
# =====================================================

def cargar_css():

    css_file = Path("style.css")

    if css_file.exists():

        with open(css_file, encoding="utf-8") as f:

            st.markdown(

                f"<style>{f.read()}</style>",

                unsafe_allow_html=True

            )

cargar_css()

# =====================================================
# CARGA DATOS
# =====================================================

df = cargar_datos()

detalle_oc_df = cargar_detalle_oc()

requisiciones_df = cargar_requisiciones()

cotizaciones_df = cargar_cotizaciones()

# =====================================================
# TITULO
# =====================================================

col1,col2 = st.columns([5,1])

with col1:

    st.title("📦 Dashboard Control de Compras")

with col2:

    st.metric(

        "Registros",

        f"{len(df):,}"

    )

st.divider()

# =====================================================
# FILTROS (arriba, en desplegables)
# =====================================================
# Antes estaban en la barra lateral, que se colapsa y no se
# puede volver a abrir porque el encabezado está oculto.
# Ahora quedan siempre visibles bajo el título.

FILTROS = [

    ("zona", "Zona"),

    ("contrato", "Contrato"),

    ("empresa_sol", "Empresa Solicitante"),

    ("empresa_compra", "Empresa Compra"),

    ("estado", "Estado"),

    ("tipo_pago", "Tipo de Pago")

]


def limpiar_filtros():

    for clave, _ in FILTROS:

        st.session_state[f"filtro_{clave}"] = []


total_registros = len(df)

with st.container(border=True):

    cab1, cab2 = st.columns([5, 1], vertical_alignment="center")

    cab1.markdown("### 🔍 Filtros")

    cab2.button(
        "Limpiar filtros",
        on_click=limpiar_filtros,
        width="stretch"
    )

    columnas_filtro = st.columns(3) + st.columns(3)

    seleccion = {}

    for col_ui, (clave, etiqueta) in zip(columnas_filtro, FILTROS):

        opciones = sorted(

            df[COLUMNAS[clave]]

            .dropna()

            .unique()

        )

        seleccion[clave] = col_ui.multiselect(

            etiqueta,

            opciones,

            key=f"filtro_{clave}",

            placeholder="Todos"

        )

# =====================================================
# APLICAR FILTROS
# =====================================================

df = aplicar_filtros(

    df,

    seleccion["zona"],

    seleccion["contrato"],

    seleccion["empresa_sol"],

    seleccion["empresa_compra"],

    seleccion["estado"],

    seleccion["tipo_pago"]

)

# =====================================================
# INFORMACIÓN DEL ARCHIVO
# =====================================================

st.markdown(

    f"**Mostrando {len(df):,} de {total_registros:,} registros** · "

    f"Requisiciones: **{df[COLUMNAS['requisiciones']].nunique():,}** · "

    f"OC: **{df[COLUMNAS['oc']].nunique():,}** · "

    f"Archivo: {EXCEL_FILE.name} (hoja {SHEET_NAME})"

)

# =====================================================
# TABS
# =====================================================

tab1,tab2,tab3,tab4 = st.tabs(

    [

        "📊 Resumen",

        "🔎 Seguimiento",

        "📈 Análisis",

        "📋 Base de Datos"

    ]

)

# =====================================================
# TAB 1 - RESUMEN
# =====================================================

with tab1:

    st.subheader("Resumen Ejecutivo")

    mostrar_kpis(df)

    st.divider()

    mostrar_alertas(df)

    mostrar_resumen(df)

    st.divider()

    # ==================================================
    # GRÁFICOS
    # ==================================================

    col1,col2 = st.columns(2)

    #--------------------------------------------
    # ESTADOS
    #--------------------------------------------

    with col1:

        estado_df = (

            df

            .groupby(COLUMNAS["estado"])

            .size()

            .reset_index(name="Cantidad")

            .sort_values(

                "Cantidad",

                ascending=False

            )

        )

        fig = px.bar(

            estado_df,

            x=COLUMNAS["estado"],

            y="Cantidad",

            color=COLUMNAS["estado"],

            color_discrete_map=COLOR_ESTADOS,

            text="Cantidad"

        )

        fig.update_layout(showlegend=False)

        fig = etiquetas_horizontales(fig, estado_df[COLUMNAS["estado"]], ancho=12)

        fig = estilo_grafico(fig, "Estado de las Requisiciones")

        st.plotly_chart(

            fig,

            width="stretch"

        )

    #--------------------------------------------
    # TIPO DE PAGO
    #--------------------------------------------

    with col2:

        pago_df = (

            df

            .groupby(COLUMNAS["tipo_pago"])

            .size()

            .reset_index(name="Cantidad")

        )

        fig = px.pie(

            pago_df,

            values="Cantidad",

            names=COLUMNAS["tipo_pago"],

            hole=.5

        )

        fig = estilo_grafico(fig, "Tipo de Pago")

        st.plotly_chart(

            fig,

            width="stretch"

        )

    st.divider()

    # ==================================================
    # MONTO POR ZONA
    # ==================================================

    col1,col2 = st.columns(2)

    with col1:

        zona_df = (

            df

            .groupby(COLUMNAS["zona"])[COLUMNAS["monto"]]

            .sum()

            .reset_index()

            .sort_values(

                COLUMNAS["monto"],

                ascending=False

            )

        )

        zona_df["Etiqueta"] = zona_df[COLUMNAS["monto"]].apply(formato_moneda)

        fig = barras_horizontales(

            zona_df,

            COLUMNAS["zona"],

            COLUMNAS["monto"],

            "Monto por Zona",

            color=PALETA[5]

        )

        st.plotly_chart(

            fig,

            width="stretch"

        )

    # ==================================================
    # MONTO POR CONTRATO
    # ==================================================

    with col2:

        contrato_df = (

            df

            .groupby(COLUMNAS["contrato"])[COLUMNAS["monto"]]

            .sum()

            .reset_index()

            .sort_values(

                COLUMNAS["monto"],

                ascending=False

            )

            .head(10)

        )

        contrato_df["Etiqueta"] = contrato_df[COLUMNAS["monto"]].apply(formato_moneda)

        fig = barras_horizontales(

            contrato_df,

            COLUMNAS["contrato"],

            COLUMNAS["monto"],

            "Top 10 Contratos"

        )

        st.plotly_chart(

            fig,

            width="stretch"

        )

    st.divider()

    # ==================================================
    # EMPRESA SOLICITANTE
    # ==================================================

    colores_empresa = mapa_colores(

        df[COLUMNAS["empresa_sol"]],

        df[COLUMNAS["empresa_compra"]]

    )

    col1,col2 = st.columns(2)

    with col1:

        empresa_df = (

            df

            .groupby(COLUMNAS["empresa_sol"])[COLUMNAS["monto"]]

            .sum()

            .reset_index()

            .sort_values(

                COLUMNAS["monto"],

                ascending=False

            )

        )

        empresa_df["Etiqueta"] = empresa_df[COLUMNAS["monto"]].apply(formato_moneda)

        fig = px.bar(

            empresa_df,

            x=COLUMNAS["empresa_sol"],

            y=COLUMNAS["monto"],

            color=COLUMNAS["empresa_sol"],

            color_discrete_map=colores_empresa,

            text="Etiqueta"

        )

        fig.update_layout(showlegend=False)

        fig = estilo_grafico(fig, "Monto por Empresa Solicitante")

        st.plotly_chart(

            fig,

            width="stretch"

        )

    # ==================================================
    # TIEMPO CICLO
    # ==================================================

    with col2:

        tiempo_df = (

            df

            .groupby(COLUMNAS["empresa_sol"])[COLUMNAS["tiempo"]]

            .mean()

            .reset_index()

            .sort_values(

                COLUMNAS["tiempo"],

                ascending=False

            )

        )

        fig = px.bar(

            tiempo_df,

            x=COLUMNAS["empresa_sol"],

            y=COLUMNAS["tiempo"],

            color=COLUMNAS["empresa_sol"],

            color_discrete_map=colores_empresa,

            text_auto=".1f"

        )

        fig.update_layout(showlegend=False)

        fig.update_yaxes(title="Días promedio")

        fig = estilo_grafico(fig, "Tiempo Promedio de Ciclo")

        st.plotly_chart(

            fig,

            width="stretch"

        )

# =====================================================
# TAB 2 - SEGUIMIENTO
# =====================================================

with tab2:

    st.subheader("🔎 Seguimiento de Compras")

    tipo_busqueda = st.radio(

        "Buscar por",

        [

            "Requisición",

            "OC",

            "N° Orden"

        ],

        horizontal=True

    )

    # ==============================================
    # REQUISICIÓN
    # ==============================================

    if tipo_busqueda == "Requisición":

        lista = sorted(

            df[COLUMNAS["requisiciones"]]

            .dropna()

            .unique()

        )

        valor = st.selectbox(

            "Seleccione una Requisición",

            lista

        )

        detalle = df[

            df[COLUMNAS["requisiciones"]] == valor

        ]

    # ==============================================
    # OC
    # ==============================================

    elif tipo_busqueda == "OC":

        lista = sorted(

            df[COLUMNAS["oc"]]

            .dropna()

            .unique()

        )

        valor = st.selectbox(

            "Seleccione una OC",

            lista

        )

        detalle = df[

            df[COLUMNAS["oc"]] == valor

        ]

    # ==============================================
    # ORDEN
    # ==============================================

    else:

        lista = sorted(

            df[COLUMNAS["nro_orden"]]

            .dropna()

            .unique()

        )

        valor = st.selectbox(

            "Seleccione un Número de Orden",

            lista

        )

        detalle = df[

            df[COLUMNAS["nro_orden"]] == valor

        ]

    st.divider()

    # ==================================================
    # DETALLE DE OC (desde "Detalle solicitudes OC.xlsx")
    # ==================================================
    # Se calcula aquí (antes de las métricas) porque el monto
    # "real" de la solicitud debe salir de este documento, no
    # de la columna "Monto Orden" de la tabla principal.

    ordenes_sel = (

        detalle[COLUMNAS["ordenes"]]

        .dropna()

        .unique()

    )

    detalle_items = detalle_oc_df[

        detalle_oc_df[COLUMNAS_DETALLE_OC["llave"]]

        .isin(ordenes_sel)

    ]

    columnas_detalle = [

        v for k, v in COLUMNAS_DETALLE_OC.items()

        if k != "llave"

    ]

    # ==================================================
    # INFORMACIÓN GENERAL
    # ==================================================

    fila1 = st.columns(4)

    fila1[0].metric(

        "Estado",

        detalle.iloc[0][COLUMNAS["estado"]]

    )

    fila1[1].metric(

        "Monto",

        f"${detalle_items[COLUMNAS_DETALLE_OC['valor_neto']].sum():,.0f}"

    )

    fila1[2].metric(

        "OC Asociadas",

        detalle[COLUMNAS["oc"]].nunique()

    )

    fila1[3].metric(

        "Tiempo Ciclo",

        f"{detalle[COLUMNAS['tiempo']].mean():.1f} días"

    )

    st.divider()

    # ==================================================
    # INFORMACIÓN DEL PROCESO
    # ==================================================

    c1,c2 = st.columns(2)

    with c1:

        st.markdown("### Información General")

        st.write(

            "**Empresa Solicitante:**",

            detalle.iloc[0][COLUMNAS["empresa_sol"]]

        )

        st.write(

            "**Empresa Compra:**",

            detalle.iloc[0][COLUMNAS["empresa_compra"]]

        )

        st.write(

            "**Contrato:**",

            detalle.iloc[0][COLUMNAS["contrato"]]

        )

        st.write(

            "**Solicitante:**",

            detalle.iloc[0][COLUMNAS["solicitante"]]

        )

        st.write(

            "**Zona:**",

            detalle.iloc[0][COLUMNAS["zona"]]

        )

        st.write(

            "**Tipo Compra:**",

            detalle.iloc[0][COLUMNAS["tipo_compra"]]

        )

    with c2:

        st.markdown("### Fechas")

        st.markdown(

            f"**Requisición:** {fmt_fecha(detalle.iloc[0][COLUMNAS['fecha_req']])}"

        )

        st.markdown(

            f"**Cotización:** {fmt_fecha(detalle.iloc[0][COLUMNAS['fecha_cot']])}"

        )

        st.markdown(

            f"**OC:** {fmt_fecha(detalle.iloc[0][COLUMNAS['fecha_oc']])}"

        )

        st.markdown(

            f"**Pago:** {fmt_fecha(detalle.iloc[0][COLUMNAS['fecha_pago']])}"

        )

        st.markdown(

            f"**Entrega:** {fmt_fecha(detalle.iloc[0][COLUMNAS['fecha_entrega']])}"

        )

    st.divider()

    # ==================================================
    # TIMELINE
    # ==================================================

    st.markdown("## Estado del Proceso")

    estado = detalle.iloc[0][COLUMNAS["estado"]]

    if estado == "Proceso OC completo":

        progreso = 100

    elif estado == "Falta entrega":

        progreso = 80

    elif estado == "Falta OC":

        progreso = 45

    elif estado == "Falta cotización":

        progreso = 20

    else:

        progreso = 10

    st.progress(progreso/100)

    st.write(f"Avance del proceso: **{progreso}%**")

    col = st.columns(5)

    col[0].success("Requisición")

    if pd.notna(detalle.iloc[0][COLUMNAS["fecha_cot"]]):
        col[1].success("Cotización")
    else:
        col[1].warning("Pendiente")

    if pd.notna(detalle.iloc[0][COLUMNAS["fecha_oc"]]):
        col[2].success("OC")
    else:
        col[2].warning("Pendiente")

    if pd.notna(detalle.iloc[0][COLUMNAS["fecha_pago"]]):
        col[3].success("Pago")
    else:
        col[3].warning("Pendiente")

    if pd.notna(detalle.iloc[0][COLUMNAS["fecha_entrega"]]):
        col[4].success("Entrega")
    else:
        col[4].warning("Pendiente")

    st.divider()

    # ==================================================
    # ÍTEMS DE LA REQUISICIÓN (hoja "Requisiciones")
    # ==================================================

    st.subheader("Ítems de la Requisición")

    items_requisicion = requisiciones_df[

        requisiciones_df[COLUMNAS_REQUISICIONES["llave"]]

        .isin(ordenes_sel)

    ]

    columnas_requisicion = [

        v for k, v in COLUMNAS_REQUISICIONES.items()

        if k != "llave"

    ]

    if items_requisicion.empty:

        st.info(
            "No se encontró detalle de productos para esta requisición "
            "en la hoja 'Requisiciones'."
        )

    else:

        st.dataframe(

            items_requisicion[columnas_requisicion],

            width="stretch",

            hide_index=True

        )

    csv_requisicion = (

        items_requisicion[columnas_requisicion]

        .to_csv(index=False)

        .encode("utf-8-sig")

    )

    st.download_button(

        "📥 Descargar ítems de la requisición",

        csv_requisicion,

        "items_requisicion.csv",

        "text/csv",

        key="descargar_items_requisicion"

    )

    # ==================================================
    # COTIZACIONES (hoja "Cotizaciones")
    # ==================================================

    st.divider()

    st.subheader("Cotizaciones")

    cotizaciones_items = cotizaciones_df[

        cotizaciones_df[COLUMNAS_COTIZACIONES["llave"]]

        .isin(ordenes_sel)

    ]

    columnas_cotizacion = [

        v for k, v in COLUMNAS_COTIZACIONES.items()

        if k != "llave"

    ]

    if cotizaciones_items.empty:

        st.info(
            "No se encontró detalle de cotizaciones para esta solicitud "
            "en la hoja 'Cotizaciones'."
        )

    else:

        st.dataframe(

            cotizaciones_items[columnas_cotizacion],

            width="stretch",

            hide_index=True

        )

        total_neto_cot = cotizaciones_items[COLUMNAS_COTIZACIONES["subtotal"]].sum()

        total_iva_cot = cotizaciones_items[COLUMNAS_COTIZACIONES["totales"]].sum()

        fila_totales = st.columns(2)

        fila_totales[0].metric(

            "Total Neto",

            f"${total_neto_cot:,.0f}"

        )

        fila_totales[1].metric(

            "Total (Neto + IVA)",

            f"${total_iva_cot:,.0f}"

        )

    excel_cotizaciones = exportar_excel(

        cotizaciones_items[columnas_cotizacion],

        nombre_hoja="Cotizaciones"

    )

    st.download_button(

        "📥 Descargar cotizaciones",

        excel_cotizaciones,

        "cotizaciones.xlsx",

        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

        key="descargar_cotizaciones"

    )

    # ==================================================
    # TABLA DE ÍTEMS (Detalle solicitudes OC.xlsx)
    # ==================================================

    st.divider()

    st.subheader("Detalle de la Solicitud")

    if detalle_items.empty:

        st.info(
            "No se encontró detalle de ítems para esta solicitud "
            "en 'Detalle solicitudes OC.xlsx'."
        )

    else:

        st.dataframe(

            detalle_items[columnas_detalle],

            width="stretch",

            hide_index=True

        )

    # ==================================================
    # EXPORTAR
    # ==================================================

    csv = (

        detalle_items[columnas_detalle]

        .to_csv(index=False)

        .encode("utf-8-sig")

    )

    st.download_button(

        "📥 Descargar detalle",

        csv,

        "detalle_solicitud.csv",

        "text/csv",

        key="descargar_detalle_oc"

    )
    
    # =====================================================
# TAB 3 - ANÁLISIS
# =====================================================

with tab3:

    st.subheader("📈 Análisis de Compras")

    #---------------------------------------------------
    # Evolución mensual
    #---------------------------------------------------

    if COLUMNAS["periodo"] in df.columns:

        mensual = (

            df.groupby(COLUMNAS["periodo"])[COLUMNAS["monto"]]

            .sum()

            .reset_index()

        )

        fig = px.line(

            mensual,

            x=COLUMNAS["periodo"],

            y=COLUMNAS["monto"],

            markers=True

        )

        fig = estilo_grafico(fig, "Monto Comprado por Período")

        st.plotly_chart(

            fig,

            width="stretch"

        )

    st.divider()

    col1,col2 = st.columns(2)

    #---------------------------------------------------
    # Empresa Compra
    #---------------------------------------------------

    with col1:

        empresa = (

            df.groupby(COLUMNAS["empresa_compra"])[COLUMNAS["monto"]]

            .sum()

            .reset_index()

            .sort_values(

                COLUMNAS["monto"],

                ascending=False

            )

            .head(15)

        )

        empresa["Etiqueta"] = empresa[COLUMNAS["monto"]].apply(formato_moneda)

        fig = barras_horizontales(

            empresa,

            COLUMNAS["empresa_compra"],

            COLUMNAS["monto"],

            "Top 15 Empresas Compra (por monto)"

        )

        st.plotly_chart(

            fig,

            width="stretch"

        )

    #---------------------------------------------------
    # Tipo Compra
    #---------------------------------------------------

    with col2:

        tipo = (

            df.groupby(COLUMNAS["tipo_compra"])

            .size()

            .reset_index(name="Cantidad")

        )

        fig = px.pie(

            tipo,

            values="Cantidad",

            names=COLUMNAS["tipo_compra"],

            hole=.5

        )

        fig = estilo_grafico(fig, "Tipo de Compra")

        st.plotly_chart(

            fig,

            width="stretch"

        )

    st.divider()

    #---------------------------------------------------
    # Tabla resumen
    #---------------------------------------------------

    resumen = (

        df.groupby(

            [

                COLUMNAS["empresa_sol"],

                COLUMNAS["contrato"]

            ]

        )

        .agg(

            Requisiciones=(COLUMNAS["requisiciones"],"nunique"),

            OC=(COLUMNAS["oc"],"nunique"),

            Monto=(COLUMNAS["monto"],"sum"),

            Tiempo=(COLUMNAS["tiempo"],"mean")

        )

        .reset_index()

    )

    st.dataframe(

        resumen,

        width="stretch",

        hide_index=True

    )

# =====================================================
# TAB 4 - BASE
# =====================================================

with tab4:

    st.subheader("📋 Base Completa")

    texto = st.text_input(

        "Buscar en toda la base"

    )

    base = df.copy()

    if texto:

        texto = texto.lower()

        filtro = base.astype(str).apply(

            lambda x:

            x.str.lower().str.contains(

                texto,

                na=False

            )

        )

        base = base[

            filtro.any(axis=1)

        ]

    st.write(

        f"Registros encontrados: {len(base):,}"

    )

    st.dataframe(

        base,

        width="stretch",

        height=650,

        hide_index=True

    )

    st.divider()

    #---------------------------------------------
    # EXPORTAR EXCEL
    #---------------------------------------------

    from io import BytesIO

    archivo = BytesIO()

    with pd.ExcelWriter(

        archivo,

        engine="openpyxl"

    ) as writer:

        base.to_excel(

            writer,

            index=False,

            sheet_name="Compras"

        )

    st.download_button(

        "📥 Descargar Excel",

        archivo.getvalue(),

        file_name="Compras_Filtradas.xlsx",

        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    )

# =====================================================
# FOOTER
# =====================================================

st.divider()

c1,c2,c3 = st.columns([1,2,1])

with c2:

    st.caption(

        "Dashboard de Compras v2.0 | Desarrollado con Streamlit"

    )
