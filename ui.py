# -*- coding: utf-8 -*-
"""
=========================================================
Dashboard Compras
Componentes visuales (HTML propio)
=========================================================
st.dataframe se dibuja en un "canvas" y no permite cambiar
el tamaño de letra. Para que las tablas se lean al proyectar
se generan aquí como tablas HTML con estilo propio (ver las
clases .tabla-* en style.css).

Se usa st.html (y no st.markdown) para que los signos "$"
no se interpreten como fórmulas matemáticas.
"""

import html

import pandas as pd
import streamlit as st

from config import DATE_FORMAT


# ======================================================
# UTILIDADES
# ======================================================

def _esc(valor):

    return html.escape(str(valor))


def _vacio(valor):

    if valor is None:
        return True

    try:
        if pd.isna(valor):
            return True
    except (TypeError, ValueError):
        pass

    return str(valor).strip().lower() in ("", "none", "nan", "nat")


def _titulo_columna(nombre):

    nombre = str(nombre).strip()

    # "valor un neto" -> "Valor un neto"; "UN" se mantiene
    return nombre[:1].upper() + nombre[1:]


def clase_estado(texto):

    t = str(texto).lower()

    if any(p in t for p in ("anul", "rechaz", "cancel")):
        return "badge-danger"

    if any(p in t for p in ("complet", "entregad", "pagad", "cerrad", "aprob")):
        return "badge-ok"

    if any(p in t for p in ("falta", "pend", "espera")):
        return "badge-warn"

    if any(p in t for p in ("envi", "proceso", "cotiz", "oc")):
        return "badge-info"

    return "badge-neutral"


def badge(texto):

    if _vacio(texto):
        return '<span class="celda-vacia">—</span>'

    return f'<span class="badge {clase_estado(texto)}">{_esc(texto)}</span>'


def moneda(valor):

    if _vacio(valor):
        return "—"

    return f"${valor:,.0f}"


# ======================================================
# TABLA HTML
# ======================================================

def tabla_html(

        df,

        moneda_cols=(),

        estado_cols=(),

        decimales=None,

        totales=(),

        alto_max=560,

        una_linea=False

):

    """
    Dibuja un DataFrame como tabla HTML de letra grande.

    moneda_cols : columnas que se muestran como $12,345
    estado_cols : columnas que se muestran como etiqueta de color
    decimales   : {columna: n_decimales} para columnas numéricas
    totales     : columnas que se suman en una fila final
    alto_max    : alto máximo en px (luego aparece scroll)
    una_linea   : True = cada celda en una sola línea (tablas
                  anchas con muchas columnas, con scroll lateral)
    """

    decimales = decimales or {}

    numericas = {
        c for c in df.columns
        if pd.api.types.is_numeric_dtype(df[c])
    }

    fechas = {
        c for c in df.columns
        if pd.api.types.is_datetime64_any_dtype(df[c])
    }

    def celda(col, valor):

        if col in estado_cols:
            return f"<td>{badge(valor)}</td>"

        if _vacio(valor):
            return '<td class="num"><span class="celda-vacia">—</span></td>' \
                if col in numericas else '<td><span class="celda-vacia">—</span></td>'

        if col in moneda_cols:
            return f'<td class="num">{moneda(valor)}</td>'

        if col in numericas:
            d = decimales.get(col)
            if d is None:
                d = 0 if float(valor).is_integer() else 1
            return f'<td class="num">{valor:,.{d}f}</td>'

        if col in fechas:
            return f"<td>{pd.to_datetime(valor).strftime(DATE_FORMAT)}</td>"

        texto = str(valor).strip()

        if texto.lower().startswith(("http://", "https://")):
            return (
                f'<td><a href="{_esc(texto)}" target="_blank" '
                f'rel="noopener">Abrir enlace ↗</a></td>'
            )

        return f"<td>{_esc(texto)}</td>"

    cabecera = "".join(

        f'<th class="{"num" if c in numericas or c in moneda_cols else ""}">'
        f"{_esc(_titulo_columna(c))}</th>"

        for c in df.columns

    )

    filas = "".join(

        "<tr>" + "".join(celda(c, v) for c, v in zip(df.columns, fila)) + "</tr>"

        for fila in df.itertuples(index=False, name=None)

    )

    pie = ""

    if totales and not df.empty:

        celdas_pie = []

        for i, c in enumerate(df.columns):

            if c in totales:
                total = pd.to_numeric(df[c], errors="coerce").sum()
                txt = moneda(total) if c in moneda_cols else f"{total:,.0f}"
                celdas_pie.append(f'<td class="num">{txt}</td>')

            elif i == 0:
                celdas_pie.append("<td>Total</td>")

            else:
                celdas_pie.append("<td></td>")

        pie = "<tfoot><tr>" + "".join(celdas_pie) + "</tr></tfoot>"

    st.html(

        f'<div class="tabla-wrap" style="max-height:{alto_max}px">'
        f'<table class="tabla{" tabla-una-linea" if una_linea else ""}">'
        f"<thead><tr>{cabecera}</tr></thead>"
        f"<tbody>{filas}</tbody>{pie}"
        f"</table></div>"
        f'<div class="tabla-nota">{len(df):,} filas</div>'

    )


# ======================================================
# ENCABEZADO
# ======================================================

def encabezado(titulo, subtitulo, datos):

    """datos: lista de (etiqueta, valor) que van a la derecha."""

    stats = "".join(

        f'<div class="hero-stat"><span>{_esc(e)}</span>'
        f"<strong>{_esc(v)}</strong></div>"

        for e, v in datos

    )

    st.html(

        '<div class="hero">'
        '<div class="hero-texto">'
        '<div class="hero-eyebrow">Gestión de compras</div>'
        f"<h1>{_esc(titulo)}</h1>"
        f"<p>{_esc(subtitulo)}</p>"
        "</div>"
        f'<div class="hero-stats">{stats}</div>'
        "</div>"

    )


# ======================================================
# TÍTULO DE SECCIÓN
# ======================================================

def seccion(titulo, descripcion=None):

    desc = f"<p>{_esc(descripcion)}</p>" if descripcion else ""

    st.html(

        f'<div class="seccion"><h3>{_esc(titulo)}</h3>{desc}</div>'

    )


# ======================================================
# TARJETAS KPI
# ======================================================

def kpis(tarjetas, columnas=4):

    """
    tarjetas: lista de dicts con
      etiqueta, valor, detalle (opcional), tono (opcional:
      'ok' | 'warn' | 'danger' | 'info')
    """

    items = []

    for t in tarjetas:

        tono = t.get("tono", "")

        detalle = (
            f'<div class="kpi-detalle">{_esc(t["detalle"])}</div>'
            if t.get("detalle") else ""
        )

        items.append(

            f'<div class="kpi kpi-{tono}">'
            f'<div class="kpi-etiqueta">{_esc(t["etiqueta"])}</div>'
            f'<div class="kpi-valor">{t["valor"]}</div>'
            f"{detalle}"
            "</div>"

        )

    st.html(

        f'<div class="kpi-grid" style="--cols:{columnas}">'
        + "".join(items) +
        "</div>"

    )


# ======================================================
# FICHA (pares etiqueta / valor)
# ======================================================

def ficha(titulo, pares):

    filas = "".join(

        f'<div class="ficha-fila"><dt>{_esc(e)}</dt>'
        f"<dd>{v if str(v).startswith('<span') else _esc(v)}</dd></div>"

        for e, v in pares

    )

    st.html(

        f'<div class="card"><div class="card-titulo">{_esc(titulo)}</div>'
        f'<dl class="ficha">{filas}</dl></div>'

    )


# ======================================================
# LÍNEA DE AVANCE DEL PROCESO
# ======================================================

def avance(pasos, porcentaje):

    """pasos: lista de (nombre, fecha_texto, completado)."""

    items = []

    for i, (nombre, fecha, hecho) in enumerate(pasos, start=1):

        clase = "paso hecho" if hecho else "paso"

        marca = "✓" if hecho else str(i)

        items.append(

            f'<div class="{clase}">'
            f'<div class="paso-circulo">{marca}</div>'
            f'<div class="paso-nombre">{_esc(nombre)}</div>'
            f'<div class="paso-fecha">{_esc(fecha)}</div>'
            "</div>"

        )

    st.html(

        '<div class="card">'
        '<div class="avance-cab">'
        '<div class="card-titulo">Estado del proceso</div>'
        f'<div class="avance-pct">{porcentaje}%</div>'
        "</div>"
        f'<div class="avance-barra"><div style="width:{porcentaje}%"></div></div>'
        f'<div class="pasos">{"".join(items)}</div>'
        "</div>"

    )


# ======================================================
# AVISOS
# ======================================================

def aviso(texto, tono="info"):

    st.html(f'<div class="aviso aviso-{tono}">{_esc(texto)}</div>')
