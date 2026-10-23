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

from ui import kpis, aviso


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

    total = len(df)

    def pct(n):

        return f"{n / total * 100:.0f}% de los registros" if total else ""

    cumplimiento = 0

    if k["requisiciones"] > 0:

        # Requisiciones únicas con proceso completo / total de
        # requisiciones únicas (antes se dividían filas por
        # requisiciones y podía dar más de 100%).
        req_completas = df.loc[
            df[COLUMNAS["estado"]] == "Proceso OC completo",
            COLUMNAS["requisiciones"]
        ].nunique()

        cumplimiento = req_completas / k["requisiciones"] * 100

    tiempo = "—" if pd.isna(k["tiempo"]) else f"{k['tiempo']} <small>días</small>"

    kpis([

        {"etiqueta": "Requisiciones", "valor": f"{k['requisiciones']:,}"},

        {"etiqueta": "Cotizaciones", "valor": f"{k['cotizaciones']:,}"},

        {"etiqueta": "Órdenes de compra", "valor": f"{k['oc']:,}"},

        {"etiqueta": "Monto total", "valor": formato_moneda(k["monto"]), "tono": "info"},

    ])

    kpis([

        {"etiqueta": "Completadas", "valor": f"{k['completadas']:,}",
         "detalle": pct(k["completadas"]), "tono": "ok"},

        {"etiqueta": "Pendientes", "valor": f"{k['pendientes']:,}",
         "detalle": pct(k["pendientes"]), "tono": "warn"},

        {"etiqueta": "Sin pago", "valor": f"{k['sin_pago']:,}",
         "detalle": pct(k["sin_pago"]), "tono": "warn"},

        {"etiqueta": "Sin entrega", "valor": f"{k['sin_entrega']:,}",
         "detalle": pct(k["sin_entrega"]), "tono": "danger"},

    ])

    kpis([

        {"etiqueta": "Tiempo de ciclo promedio", "valor": tiempo},

        {"etiqueta": "Cumplimiento", "valor": f"{cumplimiento:.1f}%",
         "detalle": "Requisiciones con proceso completo"},

        {"etiqueta": "Empresas solicitantes",
         "valor": f"{df[COLUMNAS['empresa_sol']].nunique():,}"},

        {"etiqueta": "Contratos",
         "valor": f"{df[COLUMNAS['contrato']].nunique():,}"},

    ])


# ======================================================
# RESUMEN
# ======================================================
# (Empresas y contratos ahora van como tarjetas KPI.)

def mostrar_resumen(df):

    return


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

    sin_entrega = len(
        df[
            df[COLUMNAS["fecha_entrega"]].isna()
        ]
    )

    if pendientes > 0:

        aviso(f"Hay {pendientes:,} procesos pendientes de completar.", "warn")

    if sin_entrega > 0:

        aviso(f"Hay {sin_entrega:,} órdenes sin fecha de entrega.", "danger")
