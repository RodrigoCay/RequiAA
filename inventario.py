# -*- coding: utf-8 -*-
"""
=========================================================
Dashboard Compras
Pestaña INVENTARIO (personas y móviles)
=========================================================
Muestra lo que falta o requiere cambio según el último
checklist de cada persona (por RUT) y de cada móvil (por
patente). La limpieza está en inventario_datos.py.

Para actualizar: descargar las planillas de respuestas de
Google Forms como CSV y subirlas al repositorio (se pueden
dejar con el nombre que trae la descarga).
"""

import io

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from data_loader import firma_archivo
from inventario_datos import (
    PATRON_PERSONAS,
    PATRON_MOVILES,
    ESTADOS_PENDIENTES,
    buscar_archivo,
    limpiar_personas,
    limpiar_moviles,
)
from ui import aviso, kpis, seccion, tabla_html


COLORES_ESTADO = {

    "Falta": "#B91C1C",

    "Requiere cambio": "#D97706",

    "Incompleto": "#7C3AED",

}


# ======================================================
# CARGA (con caché por archivo)
# ======================================================

@st.cache_data(show_spinner="Limpiando checklist de personas...", max_entries=2)

def _personas(ruta, firma):

    return limpiar_personas(ruta)


@st.cache_data(show_spinner="Limpiando checklist de móviles...", max_entries=2)

def _moviles(ruta, firma):

    return limpiar_moviles(ruta)


# ======================================================
# AYUDAS
# ======================================================

def _fecha(valor):

    return "—" if pd.isna(valor) else pd.to_datetime(valor).strftime("%d-%m-%Y")


def _filtro(col, etiqueta, serie, clave):

    opciones = sorted(serie.dropna().astype(str).unique())

    return col.multiselect(etiqueta, opciones, key=clave, placeholder="Todos")


def _aplicar(df, columna, seleccion):

    return df[df[columna].astype(str).isin(seleccion)] if seleccion else df


def _resumen_pendientes(items, clave, con_tallas=False):

    """Tabla 'qué hay que reponer': un renglón por ítem."""

    pend = items[items["Estado"].isin(ESTADOS_PENDIENTES)]

    if pend.empty:
        return pd.DataFrame()

    tabla = (

        pend.pivot_table(
            index=["Categoría", "Ítem"],
            columns="Estado",
            values=clave,
            aggfunc="count",
            fill_value=0
        )
        .reindex(columns=ESTADOS_PENDIENTES, fill_value=0)

    )

    tabla["Total"] = tabla.sum(axis=1)

    if con_tallas:

        def tallas(g):

            vc = g["Talla"].dropna().astype(str).str.strip()
            vc = vc[vc != ""].value_counts()

            return ", ".join(f"{t} ({n})" for t, n in vc.items()) or None

        tabla["Tallas (cantidad)"] = pend.groupby(["Categoría", "Ítem"]).apply(tallas, include_groups=False)

    tabla = (

        tabla.reset_index()
        .sort_values(["Total", "Falta"], ascending=False)

    )

    tabla.columns.name = None

    # columnas sin datos (ej. Incompleto en personas) se quitan
    return tabla[[c for c in tabla.columns if not (c in ESTADOS_PENDIENTES and tabla[c].sum() == 0)]]


def _grafico_top(resumen, titulo, estilo_grafico, top=15):

    datos = resumen.head(top).copy()

    estados = [e for e in ESTADOS_PENDIENTES if e in datos.columns]

    largo = datos.melt(

        id_vars=["Ítem"],
        value_vars=estados,
        var_name="Estado",
        value_name="Cantidad"

    )

    largo = largo[largo["Cantidad"] > 0]

    fig = px.bar(

        largo,
        x="Cantidad",
        y="Ítem",
        color="Estado",
        orientation="h",
        color_discrete_map=COLORES_ESTADO,
        category_orders={"Estado": estados}

    )

    # Total al final de cada barra
    fig.add_trace(go.Scatter(

        x=datos["Total"],
        y=datos["Ítem"],
        mode="text",
        text=[f"  {t}" for t in datos["Total"]],
        textposition="middle right",
        showlegend=False,
        hoverinfo="skip"

    ))

    nombres = list(datos["Ítem"])

    # El ítem con más pendientes queda arriba
    fig.update_yaxes(

        categoryorder="array",
        categoryarray=list(reversed(nombres)),
        title=None,
        tickvals=nombres,
        ticktext=[n if len(n) <= 34 else n[:32].rstrip() + "…" for n in nombres]

    )

    altura = max(420, 160 + 40 * len(datos))

    fig = estilo_grafico(fig, titulo, altura=altura)

    fig.update_traces(selector=dict(type="bar"), text=None)

    fig.update_xaxes(range=[0, datos["Total"].max() * 1.15], title="Cantidad")

    fig.update_layout(

        barmode="stack",
        legend=dict(orientation="h", yanchor="bottom", y=1.0, x=0, title=None)

    )

    return fig


def _resumir(items, maximo=6):

    """'a, b, c, d, e, f, g, h' -> 'a, b, c, d, e, f y 2 más'"""

    items = sorted(items)

    if len(items) <= maximo:
        return ", ".join(items)

    return ", ".join(items[:maximo]) + f" y {len(items) - maximo} más"


def _excel(hojas):

    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:

        for nombre, df in hojas.items():

            df.to_excel(writer, index=False, sheet_name=nombre[:31])

    return buffer.getvalue()


def _calidad(calidad):

    filas = pd.DataFrame(

        {"Concepto": list(calidad.keys()), "Cantidad": list(calidad.values())}

    )

    tabla_html(filas, alto_max=420)


# ======================================================
# PERSONAS
# ======================================================

def _vista_personas(estilo_grafico):

    ruta = buscar_archivo(PATRON_PERSONAS)

    if ruta is None:

        aviso(
            "No se encontró el checklist por cargo. Suba al repositorio el CSV "
            "descargado de Google Sheets (su nombre debe contener 'por cargo').",
            "warn"
        )

        return

    datos = _personas(str(ruta), firma_archivo(ruta))

    fichas, items = datos["fichas"], datos["items"]

    # ---------------- filtros ----------------
    with st.container(border=True):

        f1, f2, f3, f4 = st.columns(4)

        s_emp = _filtro(f1, "Empresa", fichas["Empresa"], "inv_p_empresa")

        s_zon = _filtro(f2, "Zonal", fichas["Zonal"], "inv_p_zonal")

        s_ser = _filtro(f3, "Servicio", fichas["Servicio"], "inv_p_servicio")

        s_cat = _filtro(f4, "Categoría", items["Categoría"], "inv_p_categoria")

    fichas_f = _aplicar(_aplicar(_aplicar(fichas, "Empresa", s_emp), "Zonal", s_zon), "Servicio", s_ser)

    items_f = _aplicar(items[items["RUT"].isin(fichas_f["RUT"])], "Categoría", s_cat)

    pend = items_f[items_f["Estado"].isin(ESTADOS_PENDIENTES)]

    con_pend = pend["RUT"].nunique()

    total = len(fichas_f)

    kpis([

        {"etiqueta": "Personas revisadas", "valor": f"{total:,}",
         "detalle": f"Última revisión: {_fecha(fichas_f['Última revisión'].max())}"},

        {"etiqueta": "Con pendientes", "valor": f"{con_pend:,}",
         "detalle": f"{con_pend / total * 100:.0f}% de las personas" if total else "", "tono": "warn"},

        {"etiqueta": "Ítems que faltan", "valor": f"{(pend['Estado'] == 'Falta').sum():,}",
         "detalle": "Marcados como NT (no tiene)", "tono": "danger"},

        {"etiqueta": "Ítems a cambiar", "valor": f"{(pend['Estado'] == 'Requiere cambio').sum():,}",
         "detalle": "Marcados como NC o Malo", "tono": "warn"},

    ])

    if pend.empty:

        aviso("No hay ítems pendientes con los filtros seleccionados.", "info")

        return

    resumen = _resumen_pendientes(items_f, "RUT", con_tallas=True)

    resumen = resumen.merge(

        pend.groupby(["Categoría", "Ítem"])["RUT"].nunique().rename("Personas").reset_index(),
        on=["Categoría", "Ítem"]

    )

    st.plotly_chart(

        _grafico_top(resumen, "Ítems con más pendientes (top 15)", estilo_grafico),

        width="stretch"

    )

    # ---------------- necesidad de reposición ----------------
    seccion(
        "Necesidad de reposición",
        "Cantidad de ítems que faltan o requieren cambio. Para ropa y calzado se indica la talla."
    )

    tabla_html(resumen, alto_max=600)

    # ---------------- personas con pendientes ----------------
    seccion("Personas con pendientes")

    por_persona = (

        pend.groupby("RUT")
        .agg(
            Faltan=("Estado", lambda s: (s == "Falta").sum()),
            Cambio=("Estado", lambda s: (s == "Requiere cambio").sum()),
            Pendientes=("Ítem", _resumir)
        )
        .reset_index()

    )

    personas = (

        fichas_f.merge(por_persona, on="RUT")
        .sort_values(["Faltan", "Cambio"], ascending=False)

    )

    personas["Última revisión"] = personas["Última revisión"].map(_fecha)

    tabla_html(

        personas[["Nombre", "RUT", "Zonal", "Servicio", "Cargo", "Última revisión", "Faltan", "Cambio", "Pendientes"]],

        alto_max=600

    )

    # ---------------- detalle por persona ----------------
    seccion("Detalle por persona")

    opciones = (fichas_f["Nombre"] + " · " + fichas_f["RUT"]).tolist()

    elegido = st.selectbox("Persona", sorted(opciones), key="inv_p_persona")

    if elegido:

        rut = elegido.split(" · ")[-1]

        ficha = fichas_f[fichas_f["RUT"] == rut].iloc[0]

        kpis([

            {"etiqueta": "Cargo", "valor": f"<span style='font-size:1.2rem'>{ficha['Cargo']}</span>"},

            {"etiqueta": "Servicio / Zonal", "valor": f"<span style='font-size:1.2rem'>{ficha['Servicio']} · {ficha['Zonal']}</span>"},

            {"etiqueta": "Tallas (torso / pantalón / zapatos)",
             "valor": f"<span style='font-size:1.2rem'>{ficha['Talla Torso'] or '—'} / {ficha['Talla Pantalón'] or '—'} / {ficha['Talla zapatos'] or '—'}</span>"},

            {"etiqueta": "Última revisión", "valor": f"<span style='font-size:1.2rem'>{_fecha(ficha['Última revisión'])}</span>"},

        ])

        if pd.notna(ficha["Observaciones"]) and str(ficha["Observaciones"]).strip():

            aviso(f"Observaciones: {ficha['Observaciones']}", "info")

        detalle = items_f[items_f["RUT"] == rut].copy()

        detalle["_orden"] = detalle["Estado"].map({e: i for i, e in enumerate(ESTADOS_PENDIENTES)}).fillna(9)

        detalle = detalle.sort_values(["_orden", "Categoría", "Ítem"])

        tabla_html(

            detalle[["Categoría", "Ítem", "Estado", "Valor registrado"]],

            estado_cols=["Estado"],

            alto_max=600

        )

    # ---------------- calidad de datos ----------------
    with st.expander("Calidad de los datos (limpieza aplicada)"):

        _calidad(datos["calidad"])

        if not datos["conflictos"].empty:

            st.markdown("**Mismo RUT con nombres distintos** (revisar en el formulario):")

            tabla_html(datos["conflictos"], alto_max=300)

        if not datos["invalidos"].empty:

            st.markdown("**RUT con dígito verificador inválido** (probable error de digitación):")

            tabla_html(datos["invalidos"], alto_max=300)

    # ---------------- descarga ----------------
    detalle_largo = items_f[[

        "RUT", "Nombre", "Empresa", "Zonal", "Servicio", "Cargo",
        "Categoría", "Ítem", "Estado", "Valor registrado", "Talla"

    ]]

    fichas_xls = fichas_f.drop(columns=["RUT válido"]).copy()

    fichas_xls["Última revisión"] = fichas_xls["Última revisión"].dt.strftime("%d-%m-%Y %H:%M")

    st.download_button(

        "Descargar inventario de personas (.xlsx)",

        _excel({

            "Reposición": resumen,

            "Personas con pendientes": personas,

            "Personas (último registro)": fichas_xls,

            "Detalle por ítem": detalle_largo,

            "Descartados sin RUT": datos["descartados"],

        }),

        "inventario_personas.xlsx",

        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

        key="inv_p_descarga"

    )


# ======================================================
# MÓVILES
# ======================================================

def _vista_moviles(estilo_grafico):

    ruta = buscar_archivo(PATRON_MOVILES)

    if ruta is None:

        aviso(
            "No se encontró el checklist de móviles. Suba al repositorio el CSV "
            "descargado de Google Sheets (su nombre debe contener 'movil').",
            "warn"
        )

        return

    datos = _moviles(str(ruta), firma_archivo(ruta))

    fichas, items = datos["fichas"], datos["items"]

    with st.container(border=True):

        f1, f2, f3 = st.columns(3)

        s_emp = _filtro(f1, "Empresa", fichas["Empresa"], "inv_m_empresa")

        s_area = _filtro(f2, "Área", fichas["Área"], "inv_m_area")

        s_pat = _filtro(f3, "Patente", fichas["Patente"], "inv_m_patente")

    fichas_f = _aplicar(_aplicar(_aplicar(fichas, "Empresa", s_emp), "Área", s_area), "Patente", s_pat)

    items_f = items[items["Patente"].isin(fichas_f["Patente"])]

    pend = items_f[items_f["Estado"].isin(ESTADOS_PENDIENTES)]

    hoy = pd.Timestamp.now().normalize()

    dias = (hoy - fichas_f["Fecha inspección"]).dt.days

    atrasados = int((dias > 30).sum())

    kpis([

        {"etiqueta": "Móviles revisados", "valor": f"{len(fichas_f):,}",
         "detalle": f"Última inspección: {_fecha(fichas_f['Fecha inspección'].max())}"},

        {"etiqueta": "Con pendientes", "valor": f"{pend['Patente'].nunique():,}", "tono": "warn"},

        {"etiqueta": "Ítems que faltan", "valor": f"{(pend['Estado'] == 'Falta').sum():,}", "tono": "danger"},

        {"etiqueta": "Sin inspección hace +30 días", "valor": f"{atrasados:,}",
         "detalle": "Según la última fecha de inspección", "tono": "warn" if atrasados else "ok"},

    ])

    if not pend.empty:

        resumen = _resumen_pendientes(items_f, "Patente")

        resumen = resumen.merge(

            pend.groupby(["Categoría", "Ítem"])["Patente"].nunique().rename("Móviles").reset_index(),
            on=["Categoría", "Ítem"]

        )

        st.plotly_chart(

            _grafico_top(resumen, "Ítems con más pendientes en móviles (top 15)", estilo_grafico),

            width="stretch"

        )

        seccion("Necesidad de reposición en móviles")

        tabla_html(resumen.drop(columns="Categoría"), alto_max=600)

    else:

        resumen = pd.DataFrame()

        aviso("No hay ítems pendientes con los filtros seleccionados.", "info")

    # ---------------- estado por móvil ----------------
    seccion("Estado por móvil", "Último checklist de cada patente.")

    por_movil = (

        pend.groupby("Patente")
        .agg(
            Faltan=("Estado", lambda s: (s == "Falta").sum()),
            Cambio=("Estado", lambda s: (s == "Requiere cambio").sum()),
            Incompleto=("Estado", lambda s: (s == "Incompleto").sum()),
            Pendientes=("Ítem", lambda s: _resumir(s, 4))
        )

    )

    moviles = fichas_f.set_index("Patente").join(por_movil).reset_index()

    for c in ["Faltan", "Cambio", "Incompleto"]:

        moviles[c] = moviles[c].fillna(0).astype(int)

    moviles["Días desde inspección"] = (hoy - moviles["Fecha inspección"]).dt.days

    moviles["Fecha inspección"] = moviles["Fecha inspección"].map(_fecha)

    moviles = moviles.sort_values(["Faltan", "Cambio", "Incompleto"], ascending=False)

    columnas = [

        "Patente", "Área", "Brigada responsable", "Fecha inspección",
        "Días desde inspección", "Faltan", "Cambio", "Incompleto", "Pendientes",
        "Foto caja herramientas", "Foto interior cúpula"

    ]

    tabla_html(moviles[columnas], alto_max=650)

    # ---------------- detalle por móvil ----------------
    seccion("Detalle por móvil")

    patente = st.selectbox("Patente", sorted(fichas_f["Patente"]), key="inv_m_detalle")

    if patente:

        detalle = items_f[items_f["Patente"] == patente].copy()

        detalle["_orden"] = detalle["Estado"].map({e: i for i, e in enumerate(ESTADOS_PENDIENTES)}).fillna(9)

        tabla_html(

            detalle.sort_values(["_orden", "Ítem"])[["Categoría", "Ítem", "Estado", "Valor registrado"]],

            estado_cols=["Estado"],

            alto_max=600

        )

    with st.expander("Calidad de los datos (limpieza aplicada)"):

        _calidad(datos["calidad"])

    st.download_button(

        "Descargar inventario de móviles (.xlsx)",

        _excel({

            "Reposición": resumen,

            "Estado por móvil": moviles[columnas],

            "Detalle por ítem": items_f,

        }),

        "inventario_moviles.xlsx",

        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

        key="inv_m_descarga"

    )


# ======================================================
# PESTAÑA
# ======================================================

def mostrar_inventario(estilo_grafico):

    seccion(

        "Inventario de EPP, herramientas y móviles",

        "Último checklist por persona (RUT) y por móvil (patente). "
        "Los filtros generales de compras no aplican aquí."

    )

    vista = st.radio(

        "Ver inventario de",

        ["Personas", "Móviles"],

        horizontal=True,

        key="inv_vista"

    )

    if vista == "Personas":

        _vista_personas(estilo_grafico)

    else:

        _vista_moviles(estilo_grafico)
