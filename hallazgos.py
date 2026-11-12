# -*- coding: utf-8 -*-
"""
=========================================================
Dashboard Compras
Pestaña HALLAZGOS (Prevención de Riesgos)
=========================================================
Hallazgos abiertos / cerrados por fecha, contrato, zona y
delegación, con los días que llevan abiertos (los hallazgos
abiertos fuera de plazo pueden generar multas).

Para actualizar: subir las planillas .xlsx (una por empresa
y zona) a la carpeta "hallazgos" del repositorio. Se leen
todas las que estén ahí; la limpieza está en
hallazgos_datos.py.
"""

import io
from datetime import date

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from data_loader import firma_archivo
from hallazgos_datos import (
    ABIERTO, CERRADO, RIS, ESTADOS_ORDEN, TRAMOS_ORDEN,
    archivos_hallazgos, cargar_hallazgos,
)
from ui import aviso, kpis, seccion, tabla_html


COLOR_ESTADO = {
    ABIERTO: "#B91C1C",
    CERRADO: "#15803D",
    RIS: "#0E7490",
    "Sin estado": "#9CA3AF",
}

COLOR_TRAMO = {
    "0 a 15 días": "#FCD34D",
    "16 a 30 días": "#F59E0B",
    "31 a 60 días": "#EA580C",
    "61 a 90 días": "#DC2626",
    "Más de 90 días": "#7F1D1D",
}

MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]


# ======================================================
# CARGA
# ======================================================

@st.cache_data(show_spinner="Leyendo planillas de hallazgos...", max_entries=2)
def _cargar(rutas, firmas, hoy):

    from pathlib import Path

    return cargar_hallazgos([Path(r) for r in rutas], hoy=hoy)


# ======================================================
# AYUDAS
# ======================================================

def _mes(ts):

    return f"{MESES[ts.month - 1]} {ts.year}"


def _filtro(col, etiqueta, serie, clave):

    opciones = sorted(serie.dropna().astype(str).unique())

    return col.multiselect(etiqueta, opciones, key=clave, placeholder="Todos")


def _aplicar(df, columna, seleccion):

    return df[df[columna].astype(str).isin(seleccion)] if seleccion else df


def _pct(a, b):

    return a / b * 100 if b else 0


def _excel(hojas):

    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl", datetime_format="DD-MM-YYYY",
                        date_format="DD-MM-YYYY") as writer:

        for nombre, df in hojas.items():

            df.to_excel(writer, index=False, sheet_name=nombre[:31])

            hoja = writer.sheets[nombre[:31]]

            hoja.freeze_panes = "A2"

            hoja.auto_filter.ref = hoja.dimensions

            for col in hoja.columns:

                largo = max((len(str(c.value)) for c in col[:200] if c.value is not None), default=8)

                hoja.column_dimensions[col[0].column_letter].width = min(max(10, largo + 2), 60)

    return buffer.getvalue()


def _resumen(df, por, plazo):

    """Tabla Total / Abiertos / Cerrados / % cierre / Fuera de plazo por grupo."""

    g = df.assign(
        _ab=df["Situación"].eq("Abierto"),
        _ce=df["Situación"].eq("Cerrado"),
        _fp=df["Días abierto"].fillna(-1) > plazo,
    ).groupby(por, dropna=False)

    t = pd.DataFrame({
        "Hallazgos": g.size(),
        "Abiertos": g["_ab"].sum(),
        "Cerrados": g["_ce"].sum(),
        "Fuera de plazo": g["_fp"].sum(),
        "Días máx. abierto": g["Días abierto"].max(),
        "Días prom. abierto": g["Días abierto"].mean(),
    }).reset_index()

    t["% cierre"] = (t["Cerrados"] / t["Hallazgos"] * 100).round(0)

    t = t.sort_values(["Abiertos", "Hallazgos"], ascending=False)

    cols = por + ["Hallazgos", "Abiertos", "Cerrados", "% cierre",
                  "Fuera de plazo", "Días prom. abierto", "Días máx. abierto"]

    return t[cols]


# ======================================================
# GRÁFICOS
# ======================================================

def _grafico_mes(df, estilo_grafico):

    tabla = (
        df.groupby(["Mes", "Estado"]).size().unstack(fill_value=0)
        .reindex(columns=[e for e in ESTADOS_ORDEN if e in df["Estado"].unique()], fill_value=0)
        .sort_index()
    )

    etiquetas = [_mes(m) for m in tabla.index]

    fig = go.Figure()

    for estado in tabla.columns:

        fig.add_bar(
            x=etiquetas, y=tabla[estado], name=estado,
            marker_color=COLOR_ESTADO.get(estado, "#6B7280"),
            hovertemplate="%{x}<br>" + estado + ": %{y}<extra></extra>",
        )

    total = tabla.sum(axis=1)

    fig.add_scatter(
        x=etiquetas, y=total, mode="text", text=[f"<b>{t}</b>" for t in total],
        textposition="top center", showlegend=False, hoverinfo="skip",
    )

    fig = estilo_grafico(fig, "Hallazgos por mes (según fecha de creación)", altura=470)

    fig.update_layout(
        barmode="stack", bargap=0.25,
        legend=dict(orientation="h", yanchor="bottom", y=1.0, x=0, title=None),
    )

    fig.update_yaxes(title="Hallazgos", range=[0, total.max() * 1.18 if len(total) else 1])

    fig.update_xaxes(title=None, type="category")

    return fig


def _grafico_barras(tabla, eje, titulo, estilo_grafico):

    """Barras horizontales apiladas Abierto / Cerrado por grupo."""

    datos = tabla.head(12).iloc[::-1]

    nombres = datos[eje].fillna("Sin informar").astype(str).tolist()

    fig = go.Figure()

    fig.add_bar(y=nombres, x=datos["Abiertos"], name="Abiertos", orientation="h",
                marker_color=COLOR_ESTADO[ABIERTO])

    if datos["Cerrados"].sum():
        fig.add_bar(y=nombres, x=datos["Cerrados"], name="Cerrados", orientation="h",
                    marker_color=COLOR_ESTADO[CERRADO])

    fig.add_scatter(
        y=nombres, x=datos["Hallazgos"], mode="text",
        text=[f"  {a} abierto{'s' if a != 1 else ''}" if a else "" for a in datos["Abiertos"]],
        textposition="middle right", showlegend=False, hoverinfo="skip",
        textfont=dict(color=COLOR_ESTADO[ABIERTO], size=15),
    )

    fig = estilo_grafico(fig, titulo, altura=max(380, 150 + 46 * len(datos)))

    fig.update_layout(barmode="stack", legend=dict(orientation="h", yanchor="bottom", y=1.0, x=0, title=None))

    fig.update_xaxes(title="Hallazgos", range=[0, datos["Hallazgos"].max() * 1.35 if len(datos) else 1])

    fig.update_yaxes(title=None)

    return fig


def _grafico_tramos(abiertos, estilo_grafico):

    cuenta = abiertos["Tramo antigüedad"].value_counts().reindex(TRAMOS_ORDEN, fill_value=0)

    fig = go.Figure(go.Bar(
        x=cuenta.index, y=cuenta.values,
        marker_color=[COLOR_TRAMO[t] for t in cuenta.index],
        text=cuenta.values, textposition="outside",
        hovertemplate="%{x}: %{y} hallazgos<extra></extra>",
    ))

    fig = estilo_grafico(fig, "Hallazgos abiertos según días abiertos", altura=420)

    fig.update_yaxes(title="Hallazgos abiertos", range=[0, max(cuenta.max(), 1) * 1.2])

    fig.update_xaxes(title=None)

    return fig


# ======================================================
# PESTAÑA
# ======================================================

def mostrar_hallazgos(estilo_grafico):

    seccion(
        "Hallazgos de Prevención de Riesgos",
        "Hallazgos abiertos y cerrados por fecha, contrato y zona, con los días que llevan "
        "abiertos. Los filtros generales de compras no aplican aquí."
    )

    rutas = archivos_hallazgos()

    if not rutas:

        aviso(
            "No hay planillas de hallazgos. Suba los archivos .xlsx a la carpeta "
            "\"hallazgos\" del repositorio (una planilla por empresa y zona).",
            "warn"
        )

        return

    hoy = date.today().isoformat()

    datos = _cargar(
        tuple(str(r) for r in rutas),
        tuple(firma_archivo(r) for r in rutas),
        hoy,
    )

    for a in datos["avisos"]:
        aviso(a, "warn")

    df = datos["datos"]

    if df.empty:

        aviso("Las planillas no tienen hallazgos para mostrar.", "info")

        return

    # ---------------- filtros ----------------
    with st.container(border=True):

        f1, f2, f3, f4 = st.columns(4)

        fmin, fmax = df["Fecha"].min().date(), df["Fecha"].max().date()

        rango = f1.date_input(
            "Fecha de creación", value=(fmin, fmax), min_value=fmin,
            max_value=max(fmax, date.today()), format="DD-MM-YYYY", key="hz_fecha",
        )

        s_emp = _filtro(f2, "Empresa", df["Empresa"], "hz_empresa")

        s_con = _filtro(f3, "Contrato", df["Contrato"], "hz_contrato")

        s_zon = _filtro(f4, "Zona", df["Zona"], "hz_zona")

        g1, g2, g3, g4 = st.columns(4)

        base_del = _aplicar(df, "Zona", s_zon)

        s_del = _filtro(g1, "Delegación", base_del["Delegación"], "hz_delegacion")

        s_est = g2.multiselect(
            "Estado", [e for e in ESTADOS_ORDEN if e in df["Estado"].unique()],
            key="hz_estado", placeholder="Todos",
        )

        s_tip = _filtro(g3, "Tipo de hallazgo", df["Tipo de hallazgo"], "hz_tipo")

        plazo = g4.number_input(
            "Plazo para cerrar (días)", min_value=1, max_value=365, value=30, step=1,
            key="hz_plazo",
            help="Un hallazgo abierto con más días que este plazo se cuenta como 'fuera de plazo'.",
        )

    f = df

    if isinstance(rango, (tuple, list)) and len(rango) == 2:

        desde, hasta = pd.Timestamp(rango[0]), pd.Timestamp(rango[1]) + pd.Timedelta(days=1)

        f = f[(f["Fecha"] >= desde) & (f["Fecha"] < hasta)]

    for col, sel in [("Empresa", s_emp), ("Contrato", s_con), ("Zona", s_zon),
                     ("Delegación", s_del), ("Estado", s_est), ("Tipo de hallazgo", s_tip)]:
        f = _aplicar(f, col, sel)

    if f.empty:

        aviso("No hay hallazgos con los filtros seleccionados.", "info")

        return

    abiertos = f[f["Situación"] == "Abierto"]

    total = len(f)

    n_ab = len(abiertos)

    n_ce = int((f["Situación"] == "Cerrado").sum())

    fuera = abiertos[abiertos["Días abierto"] > plazo]

    alta = abiertos[abiertos["Potencialidad"].eq("Alta") | abiertos["Paralización"].eq("Si")]

    inspecciones = f["N° registro"].nunique()

    # ---------------- KPI ----------------
    kpis([

        {"etiqueta": "Hallazgos", "valor": f"{total:,}",
         "detalle": f"En {inspecciones:,} inspecciones"},

        {"etiqueta": "Abiertos", "valor": f"{n_ab:,}",
         "detalle": f"{_pct(n_ab, total):.0f}% del total", "tono": "danger" if n_ab else "ok"},

        {"etiqueta": "Cerrados", "valor": f"{n_ce:,}",
         "detalle": f"{_pct(n_ce, total):.0f}% de cierre", "tono": "ok"},

        {"etiqueta": f"Fuera de plazo (> {plazo} días)", "valor": f"{len(fuera):,}",
         "detalle": f"{_pct(len(fuera), n_ab):.0f}% de los abiertos · posible multa",
         "tono": "danger" if len(fuera) else "ok"},

    ])

    prom = abiertos["Días abierto"].mean()

    mas_antiguo = abiertos.sort_values("Días abierto", ascending=False).head(1)

    kpis([

        {"etiqueta": "Días promedio abiertos",
         "valor": "—" if pd.isna(prom) else f"{prom:.0f} <small>días</small>",
         "detalle": "Desde la fecha de creación hasta hoy", "tono": "warn"},

        {"etiqueta": "Hallazgo más antiguo abierto",
         "valor": "—" if mas_antiguo.empty else f"{int(mas_antiguo['Días abierto'].iloc[0])} <small>días</small>",
         "detalle": "" if mas_antiguo.empty else
         f"Registro {mas_antiguo['N° registro'].iloc[0]} · {mas_antiguo['Zona'].iloc[0]}",
         "tono": "danger" if not mas_antiguo.empty else ""},

        {"etiqueta": "Abiertos críticos", "valor": f"{len(alta):,}",
         "detalle": "Potencialidad alta o con paralización", "tono": "danger" if len(alta) else "ok"},

        {"etiqueta": "Nuevos abiertos (7 días)",
         "valor": f"{int((abiertos['Días abierto'] <= 7).sum()):,}",
         "detalle": f"{int((f['Fecha'] >= pd.Timestamp.today().normalize() - pd.Timedelta(days=30)).sum()):,} hallazgos creados en los últimos 30 días",
         "tono": "info"},

    ])

    # ---------------- evolución mensual ----------------
    st.plotly_chart(_grafico_mes(f, estilo_grafico), width="stretch")

    # ---------------- por contrato y zona ----------------
    por_contrato = _resumen(f, ["Contrato"], plazo)

    por_zona = _resumen(f, ["Zona"], plazo)

    c1, c2 = st.columns(2)

    with c1:
        st.plotly_chart(_grafico_barras(por_contrato, "Contrato", "Por contrato", estilo_grafico), width="stretch")

    with c2:
        st.plotly_chart(_grafico_barras(por_zona, "Zona", "Por zona", estilo_grafico), width="stretch")

    # ---------------- antigüedad ----------------
    if n_ab:

        c3, c4 = st.columns(2)

        with c3:
            st.plotly_chart(_grafico_tramos(abiertos, estilo_grafico), width="stretch")

        with c4:
            por_del = _resumen(abiertos, ["Zona", "Delegación"], plazo)
            st.plotly_chart(
                _grafico_barras(
                    por_del.assign(Delegación=por_del["Delegación"].fillna("—")),
                    "Delegación", "Abiertos por delegación", estilo_grafico),
                width="stretch",
            )

    # ---------------- tabla resumen ----------------
    seccion(
        "Resumen por empresa, zona y contrato",
        f"'Fuera de plazo' = hallazgos abiertos con más de {plazo} días desde su creación."
    )

    resumen = _resumen(f, ["Empresa", "Zona", "Contrato"], plazo)

    tabla_html(
        resumen,
        decimales={"% cierre": 0, "Días prom. abierto": 0, "Días máx. abierto": 0},
        totales=("Hallazgos", "Abiertos", "Cerrados", "Fuera de plazo"),
        alto_max=520,
    )

    # ---------------- detalle de abiertos ----------------
    seccion(
        "Hallazgos abiertos",
        "Ordenados por días abiertos (los más antiguos primero)."
    )

    if abiertos.empty:

        aviso("No hay hallazgos abiertos con los filtros seleccionados.", "ok")

    else:

        cols = ["N° registro", "Fecha", "Días abierto", "Zona", "Delegación",
                "Contrato", "Hallazgo", "Responsable"]

        detalle = abiertos.sort_values("Días abierto", ascending=False)[cols].copy()

        detalle["Fecha"] = detalle["Fecha"].dt.normalize()

        detalle["N° registro"] = detalle["N° registro"].astype("string")

        detalle["Días abierto"] = detalle["Días abierto"].astype(float)

        detalle["Fecha"] = detalle["Fecha"].dt.strftime("%d\u2011%m\u2011%Y")  # guion que no corta línea

        tabla_html(detalle, alto_max=620)

    # ---------------- descarga ----------------
    seccion("Descargar", "Base de hallazgos con los filtros aplicados, en Excel.")

    base = f.drop(columns=["Mes"]).copy()

    base["Fecha"] = base["Fecha"].dt.normalize()

    por_mes = (
        f.groupby(["Mes", "Situación"]).size().unstack(fill_value=0).reset_index()
    )

    por_mes.insert(0, "Período", por_mes["Mes"].map(_mes))

    por_mes = por_mes.drop(columns=["Mes"])

    por_mes["Total"] = por_mes.drop(columns=["Período"]).sum(axis=1)

    d1, d2 = st.columns([2, 3])

    d1.download_button(
        "Descargar hallazgos (.xlsx)",
        _excel({
            "Hallazgos": base,
            "Abiertos": base[base["Situación"] == "Abierto"].sort_values("Días abierto", ascending=False),
            "Resumen": resumen,
            "Por mes": por_mes,
        }),
        file_name=f"hallazgos_{date.today():%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        key="hz_descarga",
    )

    d2.caption(
        f"{total:,} hallazgos · hojas: Hallazgos, Abiertos, Resumen y Por mes. "
        f"Días calculados al {date.today():%d-%m-%Y}."
    )

    with st.expander("Planillas leídas y calidad de los datos"):

        tabla_html(
            pd.DataFrame({"Concepto": list(datos["calidad"].keys()),
                          "Cantidad": list(datos["calidad"].values())}),
            alto_max=360,
        )

        st.caption("Archivos: " + ", ".join(r.name for r in rutas))
