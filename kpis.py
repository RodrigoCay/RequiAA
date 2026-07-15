# -*- coding: utf-8 -*-
"""
Created on Thu Jul  9 17:43:48 2026

@author: rcayuleo
"""

"""
=========================================================
Dashboard Compras
KPIs
=========================================================
"""

import streamlit as st
import pandas as pd

from config import COLUMNAS


# ======================================================
# FORMATO MONEDA
# ======================================================

def formato_moneda(valor):

    if pd.isna(valor):
        return "$0"

    if abs(valor) >= 1_000_000_000:
        return f"${valor/1_000_000_000:.1f} MM"

    if abs(valor) >= 1_000_000:
        return f"${valor/1_000_000:.1f} M"

    return f"${valor:,.0f}"


# ======================================================
# CALCULAR KPI
# ======================================================

def obtener_kpis(df):

    estado = COLUMNAS["estado"]

    kpis = {}

    kpis["requisiciones"] = df[COLUMNAS["requisiciones"]].nunique()

    kpis["cotizaciones"] = df[COLUMNAS["cotizaciones"]].nunique()

    kpis["oc"] = df[COLUMNAS["oc"]].nunique()

    kpis["monto"] = df[COLUMNAS["monto"]].sum()

    kpis["tiempo"] = round(
        df[COLUMNAS["tiempo"]].mean(),
        1
    )

    kpis["completadas"] = len(
        df[df[estado] == "Proceso OC completo"]
    )

    kpis["pendientes"] = len(
        df[df[estado] != "Proceso OC completo"]
    )

    kpis["sin_pago"] = len(
        df[df[COLUMNAS["fecha_pago"]].isna()]
    )

    kpis["sin_entrega"] = len(
        df[df[COLUMNAS["fecha_entrega"]].isna()]
    )

    return kpis


# ======================================================
# KPI PRINCIPALES
# ======================================================

def mostrar_kpis(df):

    k = obtener_kpis(df)

    fila1 = st.columns(4)

    fila1[0].metric(
        "📋 Requisiciones",
        f"{k['requisiciones']:,}"
    )

    fila1[1].metric(
        "🛒 Cotizaciones",
        f"{k['cotizaciones']:,}"
    )

    fila1[2].metric(
        "📦 Órdenes Compra",
        f"{k['oc']:,}"
    )

    fila1[3].metric(
        "💰 Monto Total",
        formato_moneda(k["monto"])
    )

    st.markdown("")

    fila2 = st.columns(4)

    fila2[0].metric(
        "✅ Completadas",
        f"{k['completadas']:,}"
    )

    fila2[1].metric(
        "🟡 Pendientes",
        f"{k['pendientes']:,}"
    )

    fila2[2].metric(
        "💳 Sin Pago",
        f"{k['sin_pago']:,}"
    )

    fila2[3].metric(
        "🚚 Sin Entrega",
        f"{k['sin_entrega']:,}"
    )

    st.markdown("")

    fila3 = st.columns(2)

    fila3[0].metric(
        "⏱ Tiempo Ciclo Promedio",
        f"{k['tiempo']} días"
    )

    porcentaje = 0

    if k["requisiciones"] > 0:

        porcentaje = (
            k["completadas"] /
            k["requisiciones"]
        ) * 100

    fila3[1].metric(
        "📈 Cumplimiento",
        f"{porcentaje:.1f}%"
    )


# ======================================================
# RESUMEN
# ======================================================

def mostrar_resumen(df):

    c1, c2, c3 = st.columns(3)

    c1.info(f"Registros : {len(df):,}")

    c2.info(
        f"Empresas : {df[COLUMNAS['empresa_sol']].nunique()}"
    )

    c3.info(
        f"Contratos : {df[COLUMNAS['contrato']].nunique()}"
    )


# ======================================================
# ALERTAS
# ======================================================

def mostrar_alertas(df):

    pendientes = len(
        df[
            df[COLUMNAS["estado"]]
            != "Proceso OC completo"
        ]
    )

    if pendientes > 0:

        st.warning(
            f"⚠ Existen {pendientes} procesos pendientes."
        )

    sin_entrega = len(
        df[
            df[COLUMNAS["fecha_entrega"]].isna()
        ]
    )

    if sin_entrega > 0:

        st.error(
            f"🚚 Existen {sin_entrega} órdenes sin entrega."
        )