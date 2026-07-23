# -*- coding: utf-8 -*-
"""
Created on Thu Jul  9 17:49:16 2026

@author: rcayuleo
"""

"""
=========================================================
Seguimiento  de compras Apoyo activo y INGECAF
Versión 1.0

=========================================================
"""

import streamlit as st
import pandas as pd
from pathlib import Path

# Archivos del proyecto
from config import *
from data_loader import *
from kpis import *

# =====================================================
# ESTILO DE ETIQUETAS DE DATOS (PLOTLY)
# =====================================================
# Sin esto, Plotly elige automáticamente el color del texto
# (blanco o gris) según el contraste con cada barra/segmento,
# lo que hace que las etiquetas se vean inconsistentes entre
# sí. Esta función fuerza siempre el mismo color y tamaño.

def aplicar_estilo_texto(fig, tamano=14, color="#FFFFFF", min_size=12):

    fig.update_traces(
        textfont=dict(color=color, size=tamano)
    )

    fig.update_layout(
        uniformtext_minsize=min_size,
        uniformtext_mode="hide"
    )

    return fig

# =====================================================
# CONFIGURACIÓN STREAMLIT
# =====================================================

st.set_page_config(
    page_title="Dashboard Compras",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
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
# SIDEBAR
# =====================================================

st.sidebar.image(
    "https://img.icons8.com/fluency/96/purchase-order.png",
    width=80
)

st.sidebar.title("Filtros")

# -----------------------------------------------------

zona = st.sidebar.multiselect(

    "Zona",

    sorted(

        df[COLUMNAS["zona"]]

        .dropna()

        .unique()

    )

)

# -----------------------------------------------------

contrato = st.sidebar.multiselect(

    "Contrato",

    sorted(

        df[COLUMNAS["contrato"]]

        .dropna()

        .unique()

    )

)

# -----------------------------------------------------

empresa = st.sidebar.multiselect(

    "Empresa Solicitante",

    sorted(

        df[COLUMNAS["empresa_sol"]]

        .dropna()

        .unique()

    )

)

# -----------------------------------------------------

empresa_compra = st.sidebar.multiselect(

    "Empresa Compra",

    sorted(

        df[COLUMNAS["empresa_compra"]]

        .dropna()

        .unique()

    )

)

# -----------------------------------------------------

estado = st.sidebar.multiselect(

    "Estado",

    sorted(

        df[COLUMNAS["estado"]]

        .dropna()

        .unique()

    )

)

# -----------------------------------------------------

tipo_pago = st.sidebar.multiselect(

    "Tipo de Pago",

    sorted(

        df[COLUMNAS["tipo_pago"]]

        .dropna()

        .unique()

    )

)

# =====================================================
# APLICAR FILTROS
# =====================================================

df = aplicar_filtros(

    df,

    zona,

    contrato,

    empresa,

    empresa_compra,

    estado,

    tipo_pago

)

# =====================================================
# INFORMACIÓN DEL ARCHIVO
# =====================================================

with st.sidebar.expander("Información"):

    st.write(f"**Archivo:** {EXCEL_FILE.name}")

    st.write(f"**Hoja:** {SHEET_NAME}")

    st.write(f"**Registros:** {len(df):,}")

    st.write(

        f"**Requisiciones:** "

        f"{df[COLUMNAS['requisiciones']].nunique():,}"

    )

    st.write(

        f"**OC:** "

        f"{df[COLUMNAS['oc']].nunique():,}"

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

    import plotly.express as px

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

            text="Cantidad",

            template="plotly_dark"

        )

        fig.update_layout(

            title="Estado de las Requisiciones",

            showlegend=False,

            paper_bgcolor="#000000",

            plot_bgcolor="#000000"

        )

        fig = aplicar_estilo_texto(fig)

        st.plotly_chart(

            fig,

            use_container_width=True

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

            hole=.55,

            template="plotly_dark"

        )

        fig.update_layout(

            title="Tipo de Pago",

            paper_bgcolor="#000000"

        )

        fig = aplicar_estilo_texto(fig)

        st.plotly_chart(

            fig,

            use_container_width=True

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

        fig = px.bar(

            zona_df,

            x=COLUMNAS["zona"],

            y=COLUMNAS["monto"],

            color=COLUMNAS["zona"],

            template="plotly_dark",

            text_auto=".2s"

        )

        fig.update_layout(

            title="Monto por Zona",

            showlegend=False,

            paper_bgcolor="#000000",

            plot_bgcolor="#000000"

        )

        fig = aplicar_estilo_texto(fig)

        st.plotly_chart(

            fig,

            use_container_width=True

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

        fig = px.bar(

            contrato_df,

            x=COLUMNAS["monto"],

            y=COLUMNAS["contrato"],

            orientation="h",

            color=COLUMNAS["monto"],

            template="plotly_dark"

        )

        fig.update_layout(

            title="Top 10 Contratos",

            paper_bgcolor="#000000",

            plot_bgcolor="#000000",

            coloraxis_showscale=False

        )

        fig = aplicar_estilo_texto(fig)

        st.plotly_chart(

            fig,

            use_container_width=True

        )

    st.divider()

    # ==================================================
    # EMPRESA SOLICITANTE
    # ==================================================

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

        fig = px.bar(

            empresa_df,

            x=COLUMNAS["empresa_sol"],

            y=COLUMNAS["monto"],

            color=COLUMNAS["empresa_sol"],

            template="plotly_dark",

            text_auto=".2s"

        )

        fig.update_layout(

            title="Monto por Empresa Solicitante",

            showlegend=False,

            paper_bgcolor="#000000",

            plot_bgcolor="#000000"

        )

        fig = aplicar_estilo_texto(fig)

        st.plotly_chart(

            fig,

            use_container_width=True

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

            template="plotly_dark"

        )

        fig.update_layout(

            title="Tiempo Promedio de Ciclo",

            showlegend=False,

            paper_bgcolor="#000000",

            plot_bgcolor="#000000"

        )

        fig = aplicar_estilo_texto(fig)

        st.plotly_chart(

            fig,

            use_container_width=True

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
    # INFORMACIÓN GENERAL
    # ==================================================

    fila1 = st.columns(4)

    fila1[0].metric(

        "Estado",

        detalle.iloc[0][COLUMNAS["estado"]]

    )

    fila1[1].metric(

        "Monto",

        f"${detalle[COLUMNAS['monto']].sum():,.0f}"

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

        st.write(

            "**Requisición:**",

            detalle.iloc[0][COLUMNAS["fecha_req"]]

        )

        st.write(

            "**Cotización:**",

            detalle.iloc[0][COLUMNAS["fecha_cot"]]

        )

        st.write(

            "**OC:**",

            detalle.iloc[0][COLUMNAS["fecha_oc"]]

        )

        st.write(

            "**Pago:**",

            detalle.iloc[0][COLUMNAS["fecha_pago"]]

        )

        st.write(

            "**Entrega:**",

            detalle.iloc[0][COLUMNAS["fecha_entrega"]]

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
    # DETALLE DE OC (desde "Detalle solicitudes OC.xlsx")
    # ==================================================

    st.subheader("Detalle de la Solicitud")

    requisiciones_sel = (

        detalle[COLUMNAS["requisiciones"]]

        .dropna()

        .unique()

    )

    detalle_items = detalle_oc_df[

        detalle_oc_df[COLUMNAS_DETALLE_OC["requisicion"]]

        .isin(requisiciones_sel)

    ]

    columnas_detalle = list(COLUMNAS_DETALLE_OC.values())

    if detalle_items.empty:

        st.info(
            "No se encontró detalle de ítems para esta solicitud "
            "en 'Detalle solicitudes OC.xlsx'."
        )

    else:

        st.dataframe(

            detalle_items[columnas_detalle],

            use_container_width=True,

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

        "text/csv"

    )
    
    # =====================================================
# TAB 3 - ANÁLISIS
# =====================================================

with tab3:

    st.subheader("📈 Análisis de Compras")

    import plotly.express as px

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

            markers=True,

            template="plotly_dark"

        )

        fig.update_layout(

            title="Monto Comprado por Período",

            paper_bgcolor="#000000",

            plot_bgcolor="#000000"

        )

        fig = aplicar_estilo_texto(fig)

        st.plotly_chart(

            fig,

            use_container_width=True

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

        )

        fig = px.bar(

            empresa,

            x=COLUMNAS["empresa_compra"],

            y=COLUMNAS["monto"],

            color=COLUMNAS["empresa_compra"],

            template="plotly_dark"

        )

        fig.update_layout(

            title="Monto por Empresa Compra",

            showlegend=False,

            paper_bgcolor="#000000",

            plot_bgcolor="#000000"

        )

        fig = aplicar_estilo_texto(fig)

        st.plotly_chart(

            fig,

            use_container_width=True

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

            hole=.5,

            template="plotly_dark"

        )

        fig.update_layout(

            title="Tipo de Compra",

            paper_bgcolor="#000000"

        )

        fig = aplicar_estilo_texto(fig)

        st.plotly_chart(

            fig,

            use_container_width=True

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

        use_container_width=True,

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

        use_container_width=True,

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