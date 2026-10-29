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
from ui import *
from inventario import mostrar_inventario

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

TEXTO_GRAFICO = "#0F172A"

GRILLA_GRAFICO = "#EEF2F6"

# Paleta de colores intensos: se leen bien sobre blanco y
# permiten texto blanco encima (sin tonos pastel).
PALETA = [
    "#1E3A8A",  # azul marino
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
            pad=dict(l=8),
            font=dict(size=21, color=TEXTO_GRAFICO, weight=700),
            x=0,
            xanchor="left"
        ),

        font=dict(
            family="Inter, Segoe UI, Arial, sans-serif",
            size=16,
            color=TEXTO_GRAFICO
        ),

        paper_bgcolor="#FFFFFF",

        plot_bgcolor="#FFFFFF",

        height=altura,

        margin=dict(l=24, r=40, t=80, b=20),

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

            eje(range=[0, maximo * (1.3 if horizontal else 1.2)])

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

encabezado(

    "Control de Compras",

    "Requisiciones, cotizaciones y órdenes de compra",

    [

        ("Registros", f"{len(df):,}"),

        ("Requisiciones", f"{df[COLUMNAS['requisiciones']].nunique():,}"),

        ("Órdenes de compra", f"{df[COLUMNAS['oc']].nunique():,}")

    ]

)

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

    cab1.markdown("#### Filtros")

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

tab1,tab2,tab3,tab4,tab5 = st.tabs(

    [

        "Resumen",

        "Seguimiento",

        "Análisis",

        "Base de datos",

        "Inventario"

    ]

)

# =====================================================
# TAB 1 - RESUMEN
# =====================================================

with tab1:

    seccion("Resumen ejecutivo", "Indicadores calculados sobre los registros filtrados.")

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

    seccion(
        "Seguimiento de compras",
        "Busque una solicitud para ver su estado, fechas, ítems y cotizaciones."
    )

    # ==============================================
    # BÚSQUEDA
    # ==============================================

    BUSQUEDAS = {

        "Requisición": (COLUMNAS["requisiciones"], "Seleccione una requisición"),

        "OC": (COLUMNAS["oc"], "Seleccione una OC"),

        "N° Orden": (COLUMNAS["nro_orden"], "Seleccione un número de orden")

    }

    with st.container(border=True):

        b1, b2 = st.columns([1, 3])

        tipo_busqueda = b1.radio(

            "Buscar por",

            list(BUSQUEDAS),

            horizontal=False

        )

        columna_busqueda, texto_busqueda = BUSQUEDAS[tipo_busqueda]

        lista = sorted(

            df[columna_busqueda]

            .dropna()

            .unique()

        )

        valor = b2.selectbox(

            texto_busqueda,

            lista,

            placeholder="Sin resultados con los filtros actuales"

        )

    detalle = df[df[columna_busqueda] == valor]

    if detalle.empty:

        aviso(
            "No hay solicitudes que coincidan con los filtros seleccionados.",
            "info"
        )

    else:

        fila = detalle.iloc[0]

        # ==================================================
        # DETALLE DE OC (desde "Detalle solicitudes OC.xlsx")
        # ==================================================
        # El monto "real" de la solicitud sale de este documento,
        # no de la columna "Monto Orden" de la tabla principal.

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
        # INDICADORES
        # ==================================================

        tiempo_ciclo = detalle[COLUMNAS["tiempo"]].mean()

        kpis([

            {"etiqueta": "Estado", "valor": badge(fila[COLUMNAS["estado"]])},

            {"etiqueta": "Monto (neto)",
             "valor": moneda(detalle_items[COLUMNAS_DETALLE_OC["valor_neto"]].sum()),
             "tono": "info"},

            {"etiqueta": "OC asociadas",
             "valor": f"{detalle[COLUMNAS['oc']].nunique():,}"},

            {"etiqueta": "Tiempo de ciclo",
             "valor": "—" if pd.isna(tiempo_ciclo)
             else f"{tiempo_ciclo:.1f} <small>días</small>"},

        ])

        # ==================================================
        # AVANCE DEL PROCESO
        # ==================================================

        estado_actual = fila[COLUMNAS["estado"]]

        PROGRESO = {

            "Proceso OC completo": 100,

            "Falta entrega": 80,

            "Falta OC": 45,

            "Falta cotización": 20,

            "Falta cotizacion": 20

        }

        progreso = PROGRESO.get(estado_actual, 10)

        avance(

            [

                ("Requisición", fmt_fecha(fila[COLUMNAS["fecha_req"]]), True),

                ("Cotización", fmt_fecha(fila[COLUMNAS["fecha_cot"]]),
                 pd.notna(fila[COLUMNAS["fecha_cot"]])),

                ("Orden de compra", fmt_fecha(fila[COLUMNAS["fecha_oc"]]),
                 pd.notna(fila[COLUMNAS["fecha_oc"]])),

                ("Pago", fmt_fecha(fila[COLUMNAS["fecha_pago"]]),
                 pd.notna(fila[COLUMNAS["fecha_pago"]])),

                ("Entrega", fmt_fecha(fila[COLUMNAS["fecha_entrega"]]),
                 pd.notna(fila[COLUMNAS["fecha_entrega"]])),

            ],

            progreso

        )

        # ==================================================
        # INFORMACIÓN GENERAL
        # ==================================================

        c1, c2 = st.columns(2)

        with c1:

            ficha("Información general", [

                ("Empresa solicitante", fila[COLUMNAS["empresa_sol"]]),

                ("Empresa compra", fila[COLUMNAS["empresa_compra"]]),

                ("Contrato", fila[COLUMNAS["contrato"]]),

            ])

        with c2:

            ficha("Solicitud", [

                ("Solicitante", fila[COLUMNAS["solicitante"]]),

                ("Zona", fila[COLUMNAS["zona"]]),

                ("Tipo de compra", fila[COLUMNAS["tipo_compra"]]),

            ])

        # ==================================================
        # ÍTEMS DE LA REQUISICIÓN (hoja "Requisiciones")
        # ==================================================

        seccion("Ítems de la requisición")

        items_requisicion = requisiciones_df[

            requisiciones_df[COLUMNAS_REQUISICIONES["llave"]]

            .isin(ordenes_sel)

        ]

        columnas_requisicion = [

            v for k, v in COLUMNAS_REQUISICIONES.items()

            if k != "llave"

        ]

        if items_requisicion.empty:

            aviso(
                "No se encontró detalle de productos para esta requisición "
                "en la hoja 'Requisiciones'."
            )

        else:

            tabla_html(items_requisicion[columnas_requisicion])

            st.download_button(

                "Descargar ítems (.csv)",

                items_requisicion[columnas_requisicion]
                .to_csv(index=False)
                .encode("utf-8-sig"),

                "items_requisicion.csv",

                "text/csv",

                key="descargar_items_requisicion"

            )

        # ==================================================
        # COTIZACIONES (hoja "Cotizaciones")
        # ==================================================

        seccion("Cotizaciones")

        cotizaciones_items = cotizaciones_df[

            cotizaciones_df[COLUMNAS_COTIZACIONES["llave"]]

            .isin(ordenes_sel)

        ]

        columnas_cotizacion = [

            v for k, v in COLUMNAS_COTIZACIONES.items()

            if k != "llave"

        ]

        if cotizaciones_items.empty:

            aviso(
                "No se encontró detalle de cotizaciones para esta solicitud "
                "en la hoja 'Cotizaciones'."
            )

        else:

            montos_cot = [

                COLUMNAS_COTIZACIONES["valor_un_neto"],

                COLUMNAS_COTIZACIONES["subtotal"],

                COLUMNAS_COTIZACIONES["iva"],

                COLUMNAS_COTIZACIONES["totales"]

            ]

            kpis([

                {"etiqueta": "Total neto",
                 "valor": moneda(cotizaciones_items[COLUMNAS_COTIZACIONES["subtotal"]].sum())},

                {"etiqueta": "IVA",
                 "valor": moneda(cotizaciones_items[COLUMNAS_COTIZACIONES["iva"]].sum())},

                {"etiqueta": "Total (neto + IVA)",
                 "valor": moneda(cotizaciones_items[COLUMNAS_COTIZACIONES["totales"]].sum()),
                 "tono": "info"},

                {"etiqueta": "Proveedores",
                 "valor": f"{cotizaciones_items[COLUMNAS_COTIZACIONES['proveedor']].nunique():,}"},

            ])

            tabla_html(

                cotizaciones_items[columnas_cotizacion],

                moneda_cols=montos_cot,

                estado_cols=[COLUMNAS_COTIZACIONES["estado"]],

                totales=montos_cot[1:]

            )

            st.download_button(

                "Descargar cotizaciones (.xlsx)",

                exportar_excel(

                    cotizaciones_items[columnas_cotizacion],

                    nombre_hoja="Cotizaciones"

                ),

                "cotizaciones.xlsx",

                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

                key="descargar_cotizaciones"

            )

        # ==================================================
        # DETALLE DE LA SOLICITUD (Detalle solicitudes OC.xlsx)
        # ==================================================

        seccion("Detalle de la solicitud")

        if detalle_items.empty:

            aviso(
                "No se encontró detalle de ítems para esta solicitud "
                "en 'Detalle solicitudes OC.xlsx'."
            )

        else:

            tabla_html(

                detalle_items[columnas_detalle],

                moneda_cols=[

                    COLUMNAS_DETALLE_OC["valor_unitario"],

                    COLUMNAS_DETALLE_OC["valor_neto"]

                ],

                estado_cols=[COLUMNAS_DETALLE_OC["estado_gestion"]],

                totales=[COLUMNAS_DETALLE_OC["valor_neto"]]

            )

            st.download_button(

                "Descargar detalle (.csv)",

                detalle_items[columnas_detalle]
                .to_csv(index=False)
                .encode("utf-8-sig"),

                "detalle_solicitud.csv",

                "text/csv",

                key="descargar_detalle_oc"

            )

# =====================================================
# TAB 3 - ANÁLISIS
# =====================================================

with tab3:

    seccion("Análisis de compras")

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

    seccion("Resumen por empresa y contrato")

    tabla_html(

        resumen.sort_values("Monto", ascending=False),

        moneda_cols=["Monto"],

        decimales={"Tiempo": 1},

        totales=["Requisiciones", "OC", "Monto"]

    )

# =====================================================
# TAB 4 - BASE
# =====================================================

with tab4:

    seccion("Base completa", "Todos los registros con los filtros aplicados.")

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

    vista_grande = st.toggle(

        "Letra grande (para presentar)",

        value=False,

        help="Muestra la base como tabla de lectura con letra grande. "
             "La vista normal permite ordenar y ajustar columnas."

    )

    if vista_grande:

        LIMITE_FILAS = 300

        if len(base) > LIMITE_FILAS:

            aviso(
                f"Se muestran las primeras {LIMITE_FILAS} filas. "
                "Use el buscador o los filtros para acotar.",
                "info"
            )

        tabla_html(

            base.head(LIMITE_FILAS),

            moneda_cols=[COLUMNAS["monto"]],

            estado_cols=[COLUMNAS["estado"]],

            alto_max=720,

            una_linea=True

        )

    else:

        st.dataframe(

            base,

            width="stretch",

            height=650,

            row_height=40,

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

        "Descargar Excel",

        archivo.getvalue(),

        file_name="Compras_Filtradas.xlsx",

        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    )

# =====================================================
# TAB 5 - INVENTARIO (checklists de personas y móviles)
# =====================================================

with tab5:

    mostrar_inventario(estilo_grafico)

# =====================================================
# FOOTER
# =====================================================

st.divider()

c1,c2,c3 = st.columns([1,2,1])

with c2:

    st.caption(

        "Dashboard de Compras v3.0"

    )
