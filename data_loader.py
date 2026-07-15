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
    CACHE_SECONDS
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

def convertir_montos(df):

    monto = COLUMNAS["monto"]

    df[monto] = (

        df[monto]

        .astype(str)

        .str.replace("$", "", regex=False)

        .str.replace(".", "", regex=False)

        .str.replace(",", ".", regex=False)

    )

    df[monto] = pd.to_numeric(

        df[monto],

        errors="coerce"

    )

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