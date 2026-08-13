# -*- coding: utf-8 -*-
"""
Created on Thu Jul  9 12:44:39 2026

@author: rcayuleo
"""

"""
=========================================================
Dashboard de Compras
Carga y preparación de datos
=========================================================
"""

import pandas as pd
import streamlit as st
from pathlib import Path

from config import (
    EXCEL_FILE,
    SHEET_NAME,
    COLUMNAS,
    CACHE_SECONDS,
    DETALLE_OC_FILE,
    DETALLE_OC_SHEET,
    COLUMNAS_DETALLE_OC,
    REQUISICIONES_SHEET,
    COLUMNAS_REQUISICIONES,
    COTIZACIONES_SHEET,
    COLUMNAS_COTIZACIONES
)


# =====================================================
# VALIDACIÓN DE ARCHIVO
# =====================================================

def validar_archivo():

    if not Path(EXCEL_FILE).exists():

        st.error(f"No se encontró el archivo:\n\n{EXCEL_FILE}")

        st.stop()


# =====================================================
# VALIDACIÓN COLUMNAS
# =====================================================

def validar_columnas(df):

    faltantes = []

    for columna in COLUMNAS.values():

        if columna not in df.columns:

            faltantes.append(columna)

    if faltantes:

        st.error("Faltan las siguientes columnas:")

        for c in faltantes:

            st.write(f"• {c}")

        st.stop()


# =====================================================
# LIMPIAR TEXTO
# =====================================================

def limpiar_texto(df):

    for c in df.select_dtypes(include="object").columns:

        df[c] = (

            df[c]

            .astype(str)

            .str.strip()

            .replace("nan", "")

        )

    return df


# =====================================================
# CONVERTIR FECHAS
# =====================================================

def convertir_fechas(df):

    columnas_fecha = [

        COLUMNAS["fecha_req"],

        COLUMNAS["fecha_cot"],

        COLUMNAS["fecha_oc"],

        COLUMNAS["fecha_pago"],

        COLUMNAS["fecha_entrega"]

    ]

    for c in columnas_fecha:

        df[c] = pd.to_datetime(

            df[c],

            errors="coerce"

        )

    return df


# =====================================================
# CONVERTIR MONTOS
# =====================================================

# =====================================================
# LIMPIAR MONTOS EN FORMATO MONEDA (CLP)
# =====================================================
# Si la columna ya viene como número desde Excel (float),
# NO se debe tratar como texto: astype(str) sobre 99000.0
# produce "99000.0", y al limpiar el "." pensando que es
# separador de miles se borra en realidad el punto decimal,
# dejando "990000" -- de ahí el bug del monto multiplicado
# por 10. Solo se limpia como texto si realmente viene en
# formato moneda (ej. "$99.000").

def limpiar_moneda(serie):

    # 1) Intento directo: si el valor ya es numérico -aunque la
    #    columna completa sea dtype "object" por tener alguna
    #    celda vacía o de texto en otra fila-, to_numeric lo
    #    interpreta bien tal cual, sin tocarlo como texto.

    directo = pd.to_numeric(serie, errors="coerce")

    # 2) Solo para los valores que NO se pudieron convertir
    #    directo (vienen realmente como texto con formato
    #    moneda, ej. "$99.000"), se limpia el texto y se
    #    reintenta la conversión solo para esos.

    pendientes = directo.isna() & serie.notna()

    if pendientes.any():

        limpio = (

            serie[pendientes]

            .astype(str)

            .str.replace("$", "", regex=False)

            .str.replace(" ", "", regex=False)

            .str.replace(".", "", regex=False)

            .str.replace(",", ".", regex=False)

        )

        directo.loc[pendientes] = pd.to_numeric(limpio, errors="coerce")

    return directo


def convertir_montos(df):

    monto = COLUMNAS["monto"]

    df[monto] = limpiar_moneda(df[monto])

    return df


# =====================================================
# CONVERTIR TIEMPO CICLO
# =====================================================

def convertir_tiempo(df):

    tiempo = COLUMNAS["tiempo"]

    df[tiempo] = pd.to_numeric(

        df[tiempo],

        errors="coerce"

    )

    return df


# =====================================================
# COLUMNAS AUXILIARES
# =====================================================

def crear_columnas(df):

    df["Año"] = df[COLUMNAS["fecha_req"]].dt.year

    df["Mes"] = df[COLUMNAS["fecha_req"]].dt.month_name()

    df["Mes Número"] = df[COLUMNAS["fecha_req"]].dt.month

    df["Tiene Pago"] = df[COLUMNAS["fecha_pago"]].notna()

    df["Tiene Entrega"] = df[COLUMNAS["fecha_entrega"]].notna()

    df["Monto MM"] = df[COLUMNAS["monto"]] / 1_000_000

    return df


# =====================================================
# KPIs BASE
# =====================================================

def calcular_resumen(df):

    resumen = {

        "requisiciones":

            df[COLUMNAS["requisiciones"]].nunique(),

        "cotizaciones":

            df[COLUMNAS["cotizaciones"]].nunique(),

        "oc":

            df[COLUMNAS["oc"]].nunique(),

        "monto":

            df[COLUMNAS["monto"]].sum(),

        "tiempo":

            round(

                df[COLUMNAS["tiempo"]].mean(),

                1

            ),

        "pendientes":

            len(

                df[

                    df[COLUMNAS["estado"]]

                    != "Proceso OC completo"

                ]

            ),

        "completadas":

            len(

                df[

                    df[COLUMNAS["estado"]]

                    == "Proceso OC completo"

                ]

            )

    }

    return resumen


# =====================================================
# CARGA PRINCIPAL
# =====================================================

@st.cache_data(ttl=CACHE_SECONDS)

def cargar_datos():

    validar_archivo()

    df = pd.read_excel(

        EXCEL_FILE,

        sheet_name=SHEET_NAME

    )

    df.columns = (

        df.columns

        .str.strip()

    )

    validar_columnas(df)

    df = limpiar_texto(df)

    df = convertir_fechas(df)

    df = convertir_montos(df)

    df = convertir_tiempo(df)

    df = crear_columnas(df)

    return df


# =====================================================
# CARGA DETALLE DE ÍTEMS POR OC
# =====================================================
# Archivo separado en el mismo repositorio: "Detalle
# solicitudes OC.xlsx", hoja "Detalle". Se cruza con la
# tabla principal mediante la llave "N° de requisición".

@st.cache_data(ttl=CACHE_SECONDS)

def cargar_detalle_oc():

    if not Path(DETALLE_OC_FILE).exists():

        st.error(f"No se encontró el archivo:\n\n{DETALLE_OC_FILE}")

        st.stop()

    detalle = pd.read_excel(

        DETALLE_OC_FILE,

        sheet_name=DETALLE_OC_SHEET

    )

    detalle.columns = (

        detalle.columns

        .str.strip()

    )

    faltantes = [

        c for c in COLUMNAS_DETALLE_OC.values()

        if c not in detalle.columns

    ]

    if faltantes:

        st.error("Faltan las siguientes columnas en 'Detalle solicitudes OC.xlsx':")

        for c in faltantes:

            st.write(f"• {c}")

        st.stop()

    detalle = limpiar_texto(detalle)

    for col in [

        COLUMNAS_DETALLE_OC["valor_unitario"],

        COLUMNAS_DETALLE_OC["valor_neto"]

    ]:

        detalle[col] = limpiar_moneda(detalle[col])

    detalle[COLUMNAS_DETALLE_OC["cantidad"]] = pd.to_numeric(

        detalle[COLUMNAS_DETALLE_OC["cantidad"]],

        errors="coerce"

    )

    return detalle


# =====================================================
# CARGA DE ÍTEMS POR REQUISICIÓN
# =====================================================
# Mismo archivo "Detalle solicitudes OC.xlsx", hoja
# "Requisiciones". Se cruza con la tabla principal
# mediante la llave "LLave" == "Requisiciones".

@st.cache_data(ttl=CACHE_SECONDS)

def cargar_requisiciones():

    if not Path(DETALLE_OC_FILE).exists():

        st.error(f"No se encontró el archivo:\n\n{DETALLE_OC_FILE}")

        st.stop()

    requisiciones = pd.read_excel(

        DETALLE_OC_FILE,

        sheet_name=REQUISICIONES_SHEET

    )

    requisiciones.columns = (

        requisiciones.columns

        .str.strip()

    )

    faltantes = [

        c for c in COLUMNAS_REQUISICIONES.values()

        if c not in requisiciones.columns

    ]

    if faltantes:

        st.error("Faltan las siguientes columnas en la hoja 'Requisiciones':")

        for c in faltantes:

            st.write(f"• {c}")

        st.stop()

    requisiciones = requisiciones[list(COLUMNAS_REQUISICIONES.values())]

    requisiciones = limpiar_texto(requisiciones)

    requisiciones[COLUMNAS_REQUISICIONES["cantidad"]] = pd.to_numeric(

        requisiciones[COLUMNAS_REQUISICIONES["cantidad"]],

        errors="coerce"

    )

    return requisiciones


# =====================================================
# CARGA DE COTIZACIONES
# =====================================================
# Mismo archivo "Detalle solicitudes OC.xlsx", hoja
# "Cotizaciones". Se cruza con la tabla principal mediante
# la llave "LLave" == "Ordenes" (igual que el detalle de OC).

@st.cache_data(ttl=CACHE_SECONDS)

def cargar_cotizaciones():

    if not Path(DETALLE_OC_FILE).exists():

        st.error(f"No se encontró el archivo:\n\n{DETALLE_OC_FILE}")

        st.stop()

    cotizaciones = pd.read_excel(

        DETALLE_OC_FILE,

        sheet_name=COTIZACIONES_SHEET

    )

    cotizaciones.columns = (

        cotizaciones.columns

        .str.strip()

    )

    # Nota: los nombres de columna con espacios al inicio o
    # final en el Excel (ej. "tipo pago ") quedan sin ese
    # espacio tras el .str.strip() de arriba, por eso
    # COLUMNAS_COTIZACIONES en config.py los tiene sin espacio.

    faltantes = [

        c for c in COLUMNAS_COTIZACIONES.values()

        if c not in cotizaciones.columns

    ]

    if faltantes:

        st.error("Faltan las siguientes columnas en la hoja 'Cotizaciones':")

        for c in faltantes:

            st.write(f"• {c}")

        st.stop()

    cotizaciones = cotizaciones[list(COLUMNAS_COTIZACIONES.values())]

    cotizaciones = limpiar_texto(cotizaciones)

    for col in [

        COLUMNAS_COTIZACIONES["valor_un_neto"],

        COLUMNAS_COTIZACIONES["subtotal"],

        COLUMNAS_COTIZACIONES["iva"],

        COLUMNAS_COTIZACIONES["totales"]

    ]:

        cotizaciones[col] = limpiar_moneda(cotizaciones[col])

    cotizaciones[COLUMNAS_COTIZACIONES["un"]] = pd.to_numeric(

        cotizaciones[COLUMNAS_COTIZACIONES["un"]],

        errors="coerce"

    )

    return cotizaciones


# =====================================================
# FILTROS
# =====================================================

def aplicar_filtros(

        df,

        zona=None,

        contrato=None,

        empresa=None,

        empresa_compra=None,

        estado=None,

        tipo_pago=None

):

    if zona:

        df = df[df[COLUMNAS["zona"]].isin(zona)]

    if contrato:

        df = df[df[COLUMNAS["contrato"]].isin(contrato)]

    if empresa:

        df = df[df[COLUMNAS["empresa_sol"]].isin(empresa)]

    if empresa_compra:

        df = df[df[COLUMNAS["empresa_compra"]].isin(empresa_compra)]

    if estado:

        df = df[df[COLUMNAS["estado"]].isin(estado)]

    if tipo_pago:

        df = df[df[COLUMNAS["tipo_pago"]].isin(tipo_pago)]

    return df