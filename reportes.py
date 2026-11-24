# -*- coding: utf-8 -*-
"""
=========================================================
Dashboard de Compras
Reportes descargables (.xlsx) por requisición
=========================================================

Arma, a partir de la base ya filtrada (filtros de arriba de la app),
un reporte con UNA FILA POR REQUISICIÓN y la trazabilidad completa:

  - fechas del proceso (requisición, cotización, comités, OC)
  - primer y último comité, cuántas veces pasó y su resultado
  - días transcurridos entre etapas y antigüedad
  - etapa actual ("Aprobada, falta emitir OC", "En evaluación de comité"...)
  - observaciones (comentarios de gestión, comité y resumen)
  - alertas (fecha límite vencida, mucho tiempo sin OC, rechazada...)

El Excel trae además: Resumen (indicadores + tablas con gráfico),
historial de comité, ítems solicitados, cotizaciones y OC emitidas.
"""

import pandas as pd
import streamlit as st

from config import COLUMNAS
from excel_formato import excel_formato, MIME_XLSX
from ui import seccion, kpis, aviso

C = COLUMNAS
LLAVE = C["ordenes"]          # "Ordenes" (== LLave de las otras hojas)
NRO = C["nro_orden"]          # "Nro de Orden" (N° de requisición)

# =========================================================
# TIPOS DE REPORTE
# =========================================================
# Cada tipo es un filtro sobre la tabla por requisición.

TIPOS_REPORTE = {

    "Requisiciones sin OC emitida":
        "Aún no tienen ninguna OC (en cotización, en comité o aprobadas sin OC). Excluye anuladas.",

    "Aprobadas sin OC":
        "El comité las aprobó (total o parcial) pero todavía no se emite OC.",

    "En evaluación de comité":
        "Pasaron por comité y quedaron pendientes / no revisadas, sin OC.",

    "Falta cotización":
        "Requisiciones que aún no tienen cotización.",

    "Con OC, falta entrega":
        "Tienen OC emitida y falta la entrega de al menos una.",

    "Aprobación parcial":
        "Alguna parte de la requisición quedó aprobada parcialmente.",

    "Pendientes (todo lo no cerrado)":
        "Todo lo que no está en Proceso Completo ni Anulado.",

    "Todas las requisiciones":
        "Todas las requisiciones con los filtros aplicados.",

}

TRAMOS = [(-1, 7, "0 a 7 días"), (7, 15, "8 a 15 días"), (15, 30, "16 a 30 días"),
          (30, 60, "31 a 60 días"), (60, 10**6, "Más de 60 días")]


def _hoy():

    return pd.Timestamp.now(tz="America/Santiago").normalize().tz_localize(None)


def _unicos(serie, sep=" / "):

    vals = [str(v).strip() for v in serie.dropna() if str(v).strip() not in ("", "nan", "None")]

    return sep.join(dict.fromkeys(vals))


def _dias(a, b):

    """Días entre dos series de fechas (b - a)."""

    return (pd.to_datetime(b, errors="coerce") - pd.to_datetime(a, errors="coerce")).dt.days


def _tramo(d):

    if pd.isna(d):

        return "Sin fecha"

    for lo, hi, nombre in TRAMOS:

        if lo < d <= hi:

            return nombre

    return TRAMOS[-1][2]


# =========================================================
# TABLA POR REQUISICIÓN
# =========================================================

def _resumen_comite(comite_df):

    """Una fila por requisición con primer / último comité e historial."""

    if comite_df is None or comite_df.empty or "LLave" not in comite_df:

        return pd.DataFrame(columns=["LLave"])

    c = comite_df.copy()

    c["Fecha comité"] = pd.to_datetime(c["Fecha comité"], errors="coerce")

    c = c.sort_values(["LLave", "Fecha comité", "N° comité"], na_position="last")

    def _hist(g):

        partes = []

        for _, r in g.iterrows():

            fecha = r["Fecha comité"].strftime("%d-%m-%Y") if pd.notna(r["Fecha comité"]) else "s/f"

            txt = f"{r.get('Comité', '')} ({fecha}): {r.get('Resultado comité', '') or ''}"

            com = r.get("Comentario comité")

            if pd.notna(com) and str(com).strip():

                txt += f" – {str(com).strip()}"

            partes.append(txt)

        return " | ".join(partes)

    g = c.groupby("LLave", sort=False)

    # Lo aprobado: filas del acta con resultado "Aprobado" / "Aprobado parcial".
    # Si el mismo detalle se aprueba en varios comités (re-aprobación) se toma
    # la última vez; detalles distintos se suman. La hoja "Aprobado en comité"
    # del Excel muestra todas las filas del acta.
    ap = c[c["Resultado comité"].astype(str).str.strip().str.lower().str.startswith("aprob")].copy()

    ap["_det"] = ap["Detalle cotización"].fillna("").astype(str).str.strip()

    ap["_k"] = ap["_det"].str.lower().str.replace(r"\s+", " ", regex=True)

    ap = ap.drop_duplicates(["LLave", "_k"], keep="last")

    def _txt_aprob(gr):

        partes = []

        for _, r in gr.iterrows():

            det = r["_det"] or "(sin detalle en acta)"

            if str(r["Resultado comité"]).lower().startswith("aprobado parcial"):

                det += " [parcial]"

            if pd.notna(r["Monto neto comité"]):

                det += f" – $ {r['Monto neto comité']:,.0f}".replace(",", ".")

            partes.append(det)

        return " | ".join(partes)

    ga = ap.groupby("LLave", sort=False)

    aprob = pd.DataFrame({

        "Aprobado en comité": ga.apply(_txt_aprob, include_groups=False),

        "Comité de aprobación": ga["Comité"].apply(lambda s: _unicos(s, ", ")),

        "Fecha aprobación": ga["Fecha comité"].max(),

        "Monto aprobado (neto)": ga["Monto neto comité"].sum(min_count=1),

    })

    out = pd.DataFrame({

        "Primer comité": g["Comité"].first(),

        "Fecha primer comité": g["Fecha comité"].min(),

        "Último comité (acta)": g["Comité"].last(),

        "Fecha último comité (acta)": g["Fecha comité"].max(),

        "N° de comités": g["Comité"].nunique(),

        "Resultado último comité": g["Resultado comité"].last(),

        "Comentario último comité": g["Comentario comité"].apply(lambda s: _unicos(s.tail(1))),

        "Historial comité": g.apply(_hist, include_groups=False),

    }).join(aprob).reset_index()

    return out


def _etapa(r):

    if r["Anulada"]:

        return "Anulada"

    if r["Tiene OC"]:

        est = r["Estado"]

        if "Falta entrega" in est:

            return "OC emitida, falta entrega"

        if "Proceso Completo" in est:

            return "Proceso completo"

        return "OC emitida"

    res = r.get("Resultado último comité")

    if pd.isna(res):

        res = r.get("Resultado comité")

    res = "" if pd.isna(res) else str(res).lower()

    if "rechaz" in res:

        return "Rechazada en comité"

    if res.startswith("aprob") or "Aprobado Comité" in r["Estado"] or "Aprobación parcial" in r["Estado"]:

        return "Aprobada, falta emitir OC"

    if pd.isna(r["Fecha Cotización"]) and r["N° cotizaciones"] == 0 and r["N° de comités"] == 0:

        return "En cotización"

    if r["N° de comités"] > 0:

        return "En evaluación de comité"

    return "Cotizada, sin pasar a comité"


def tabla_requisiciones(df, comite_df, cotizaciones_df, hoy=None):

    """Agrupa la base (una fila por OC) en una fila por requisición."""

    hoy = hoy or _hoy()

    if df.empty:

        return pd.DataFrame()

    g = df.groupby(LLAVE, sort=True)

    def primero(col):

        return g[col].first() if col in df else None

    t = pd.DataFrame({

        "N° requisición": g[NRO].first(),

        "Contrato": primero(C["contrato"]),

        "Zona": primero(C["zona"]),

        "Empresa solicitante": primero(C["empresa_sol"]),

        "Solicitante": primero(C["solicitante"]),

        "Tipo de compra": primero(C["tipo_compra"]),

        "Motivo de compra": primero("Motivo de compra"),

        "Estado": g[C["estado"]].apply(_unicos),

        "Estado Gestión": g["Estado Gestión"].apply(_unicos) if "Estado Gestión" in df else "",

        "Fecha Requisición": g[C["fecha_req"]].min(),

        "Fecha Cotización": g[C["fecha_cot"]].min(),

        "Fecha primera OC": g[C["fecha_oc"]].min(),

        "Fecha última OC": g[C["fecha_oc"]].max(),

        "Fecha límite de entrega": g["Fecha límite de entrega"].min() if "Fecha límite de entrega" in df else pd.NaT,

        "Fecha de entrega": g[C["fecha_entrega"]].max(),

        "OC emitidas": g[C["oc"]].apply(_unicos),

        "N° OC emitidas": g[C["oc"]].nunique(),

        "Proveedores OC": g[C["proveedor"]].apply(_unicos),

        "Monto cotización": primero("Monto Cotización"),

        "Monto neto OC": g["Monto Neto OC"].sum(min_count=1) if "Monto Neto OC" in df else None,

        "Ítems solicitados": primero("Ítems solicitados"),

        "Resultado comité": primero("Resultado comité"),

        "Último comité": primero("Último comité"),

        "Fecha último comité": primero("Fecha último comité"),

        "Comités": primero("Comités"),

        "Estado comité (resumen)": g["Estado comité (resumen)"].apply(_unicos) if "Estado comité (resumen)" in df else "",

        "Comentario gestión": g["Comentario"].apply(_unicos) if "Comentario" in df else "",

        "Comentario comité": g["Comentario comité"].apply(_unicos) if "Comentario comité" in df else "",

        "Comentario resumen": g["Comentario resumen"].apply(_unicos) if "Comentario resumen" in df else "",

        "Archivo requisición": g[C["requisiciones"]].first(),

        "Archivo cotización": g[C["cotizaciones"]].apply(_unicos),

        "Anulada": g[C["estado"]].apply(lambda s: s.notna().any() and (s == "Anulado").all()),

    }).reset_index().rename(columns={LLAVE: "LLave"})

    t["Tiene OC"] = t["N° OC emitidas"] > 0

    # ---------- comité (acta) ----------
    t = t.merge(_resumen_comite(comite_df), on="LLave", how="left")

    # si no aparece en el acta, se usa lo que trae la hoja Control
    t["Primer comité"] = t["Primer comité"].fillna(t["Último comité"])

    t["Fecha primer comité"] = t["Fecha primer comité"].fillna(t["Fecha último comité"])

    t["Último comité (acta)"] = t["Último comité (acta)"].fillna(t["Último comité"])

    t["Fecha último comité (acta)"] = t["Fecha último comité (acta)"].fillna(t["Fecha último comité"])

    t["N° de comités"] = t["N° de comités"].fillna(t["Comités"]).fillna(0).astype(int)

    t["Resultado último comité"] = t["Resultado último comité"].fillna(t["Resultado comité"])

    t["Comentario último comité"] = t["Comentario último comité"].replace("", pd.NA).fillna(t["Comentario comité"])

    for col in ["Aprobado en comité", "Comité de aprobación", "Fecha aprobación", "Monto aprobado (neto)"]:

        if col not in t:

            t[col] = pd.NA

    t["Monto aprobado sin OC (saldo)"] = (

        t["Monto aprobado (neto)"] - t["Monto neto OC"].fillna(0)

    ).where(t["Monto aprobado (neto)"].notna())

    # ---------- cotizaciones ----------
    if cotizaciones_df is not None and not cotizaciones_df.empty and "LLave" in cotizaciones_df:

        cg = cotizaciones_df.groupby("LLave")

        cot = pd.DataFrame({

            "N° cotizaciones": cg.size(),

            "Proveedores cotizados": cg["proveedor"].apply(lambda s: _unicos(s.astype(str).str.strip().str.upper(), ", ")),

        }).reset_index()

        t = t.merge(cot, on="LLave", how="left")

    else:

        t["N° cotizaciones"] = 0

        t["Proveedores cotizados"] = ""

    t["N° cotizaciones"] = t["N° cotizaciones"].fillna(0).astype(int)

    # ---------- días ----------
    hoy_s = pd.Series(hoy, index=t.index)

    fin_eval = t["Fecha primera OC"].fillna(hoy_s)

    t["Antigüedad (días)"] = _dias(t["Fecha Requisición"], hoy_s)

    t["Días req. → cotización"] = _dias(t["Fecha Requisición"], t["Fecha Cotización"])

    t["Días cotización → 1er comité"] = _dias(t["Fecha Cotización"], t["Fecha primer comité"])

    t["Días en evaluación"] = _dias(t["Fecha primer comité"], fin_eval)

    t["Días desde último comité"] = _dias(t["Fecha último comité (acta)"], hoy_s).where(~t["Tiene OC"])

    t["Días req. → OC"] = _dias(t["Fecha Requisición"], t["Fecha primera OC"])

    t["Días para fecha límite"] = _dias(hoy_s, t["Fecha límite de entrega"])

    t["Tramo antigüedad"] = t["Antigüedad (días)"].apply(_tramo)

    t["Etapa actual"] = t.apply(_etapa, axis=1)

    sin_det = t["Aprobado en comité"].isna() & (t["Etapa actual"] == "Aprobada, falta emitir OC")

    t.loc[sin_det, "Aprobado en comité"] = "Aprobada según estado, sin detalle en el acta"

    # ---------- alertas ----------
    def alertas(r):

        a = []

        abierta = r["Etapa actual"] not in ("Anulada", "Proceso completo")

        if abierta and pd.notna(r["Días para fecha límite"]) and r["Días para fecha límite"] < 0 and pd.isna(r["Fecha de entrega"]):

            a.append(f"Fecha límite vencida hace {int(-r['Días para fecha límite'])} días")

        if not r["Tiene OC"] and abierta and pd.notna(r["Antigüedad (días)"]) and r["Antigüedad (días)"] > 30:

            a.append("Más de 30 días sin OC")

        if r["Etapa actual"] == "Cotizada, sin pasar a comité" and pd.notna(r["Fecha Cotización"]) and (hoy - r["Fecha Cotización"]).days > 7:

            a.append("Cotizada hace más de 7 días sin comité")

        if r["Etapa actual"] == "Aprobada, falta emitir OC" and pd.notna(r["Días desde último comité"]) and r["Días desde último comité"] > 7:

            a.append("Aprobada hace más de 7 días sin OC")

        if r["Etapa actual"] == "Rechazada en comité":

            a.append("Rechazada: definir si se anula o se recotiza")

        if not r["Tiene OC"] and "Proceso Completo" in r["Estado"]:

            a.append("Estado 'Proceso Completo' sin OC registrada (revisar base)")

        if not r["Tiene OC"] and r["N° de comités"] >= 3:

            a.append(f"Pasó {r['N° de comités']} veces por comité")

        return "; ".join(a)

    t["Alertas"] = t.apply(alertas, axis=1)

    return t


# Orden y selección de columnas para la hoja de detalle

COLUMNAS_REPORTE = [

    "N° requisición", "Etapa actual", "Alertas", "Estado", "Estado Gestión",

    "Contrato", "Zona", "Empresa solicitante", "Solicitante", "Tipo de compra", "Motivo de compra",

    "Fecha Requisición", "Fecha Cotización", "Primer comité", "Fecha primer comité",

    "Último comité (acta)", "Fecha último comité (acta)", "N° de comités", "Resultado último comité",

    "Aprobado en comité", "Comité de aprobación", "Fecha aprobación", "Monto aprobado (neto)",

    "Fecha primera OC", "Fecha límite de entrega", "Fecha de entrega",

    "Antigüedad (días)", "Tramo antigüedad", "Días req. → cotización", "Días cotización → 1er comité",

    "Días en evaluación", "Días desde último comité", "Días req. → OC", "Días para fecha límite",

    "Ítems solicitados", "N° cotizaciones", "Proveedores cotizados", "Monto cotización",

    "OC emitidas", "Proveedores OC", "Monto neto OC", "Monto aprobado sin OC (saldo)",

    "Comentario último comité", "Historial comité", "Estado comité (resumen)",

    "Comentario gestión", "Comentario resumen", "Archivo requisición", "Archivo cotización",

]


def filtrar_tipo(t, tipo, incluir_anuladas=False):

    if t.empty:

        return t

    if not incluir_anuladas:

        t = t[~t["Anulada"]]

    e = t["Etapa actual"]

    if tipo == "Requisiciones sin OC emitida":

        return t[~t["Tiene OC"] & (e != "Anulada")]

    if tipo == "Aprobadas sin OC":

        return t[e == "Aprobada, falta emitir OC"]

    if tipo == "En evaluación de comité":

        return t[e == "En evaluación de comité"]

    if tipo == "Falta cotización":

        return t[e == "En cotización"]

    if tipo == "Con OC, falta entrega":

        return t[e == "OC emitida, falta entrega"]

    if tipo == "Aprobación parcial":

        return t[t["Estado"].str.contains("Aprobación parcial", na=False)]

    if tipo == "Pendientes (todo lo no cerrado)":

        return t[~e.isin(["Anulada", "Proceso completo"])]

    return t


# =========================================================
# EXCEL
# =========================================================

def _conteo(t, col, nombre):

    d = (t.groupby(col, dropna=False)

         .agg(Requisiciones=("N° requisición", "count"),

              **{"Monto cotización": ("Monto cotización", "sum")},

              **{"Monto aprobado (neto)": ("Monto aprobado (neto)", "sum")},

              **{"Antigüedad promedio (días)": ("Antigüedad (días)", "mean")})

         .reset_index()

         .rename(columns={col: nombre}))

    d[nombre] = d[nombre].fillna("Sin dato")

    d["Antigüedad promedio (días)"] = d["Antigüedad promedio (días)"].round(1)

    return d.sort_values("Requisiciones", ascending=False).reset_index(drop=True)


def generar_excel(t, tipo, filtros_txt, comite_df, requisiciones_df, cotizaciones_df, detalle_oc_df, hoy=None):

    hoy = hoy or _hoy()

    llaves = set(t["LLave"])

    det = t[[c for c in COLUMNAS_REPORTE if c in t]].sort_values(

        ["Antigüedad (días)", "N° requisición"], ascending=[False, True])

    # -------- Resumen --------
    abiertas = t[~t["Etapa actual"].isin(["Anulada", "Proceso completo"])]

    def prom(col):

        v = t[col].mean()

        return "—" if pd.isna(v) else f"{v:.1f}"

    def pesos(v):

        return f"$ {v:,.0f}".replace(",", ".")

    def n(v):

        return f"{int(v):,}".replace(",", ".")

    ind = pd.DataFrame([

        ("Tipo de reporte", tipo),

        ("Fecha de corte", hoy.strftime("%d-%m-%Y")),

        ("Requisiciones en el reporte", n(len(t))),

        ("Requisiciones abiertas", n(len(abiertas))),

        ("Requisiciones con aprobación en comité", n(t["Monto aprobado (neto)"].notna().sum())),

        ("Con alguna alerta", n((t["Alertas"] != "").sum())),

        ("Monto cotizado total", pesos(t["Monto cotización"].sum())),

        ("Monto aprobado en comité (neto)", pesos(t["Monto aprobado (neto)"].sum())),

        ("Monto neto en OC emitidas", pesos(t["Monto neto OC"].sum())),

        ("Monto aprobado sin OC (saldo)", pesos(t["Monto aprobado sin OC (saldo)"].clip(lower=0).sum())),

    ], columns=["Indicador", "Valor"])

    tiempos = pd.DataFrame([

        ("Antigüedad promedio (días desde la requisición)", prom("Antigüedad (días)")),

        ("Antigüedad máxima (días)", "—" if t["Antigüedad (días)"].isna().all() else f"{int(t['Antigüedad (días)'].max())}"),

        ("Promedio días requisición → cotización", prom("Días req. → cotización")),

        ("Promedio días cotización → 1er comité", prom("Días cotización → 1er comité")),

        ("Promedio días en evaluación de comité", prom("Días en evaluación")),

        ("Promedio días desde último comité (sin OC)", prom("Días desde último comité")),

        ("Promedio días requisición → OC", prom("Días req. → OC")),

        ("Sin pasar nunca por comité", n((t["N° de comités"] == 0).sum())),

    ], columns=["Indicador de tiempo", "Valor"])

    orden_tramo = [n for *_, n in TRAMOS] + ["Sin fecha"]

    tramos = _conteo(t, "Tramo antigüedad", "Tramo de antigüedad")

    tramos["_o"] = tramos["Tramo de antigüedad"].map({k: i for i, k in enumerate(orden_tramo)})

    tramos = tramos.sort_values("_o").drop(columns="_o").reset_index(drop=True)

    alert = (t["Alertas"].str.split("; ").explode().replace("", pd.NA).dropna()

             .str.replace(r"vencida hace \d+ días", "vencida", regex=True)

             .str.replace(r"Pasó \d+ veces", "Pasó 3 o más veces", regex=True).value_counts().rename_axis("Alerta")

             .reset_index(name="Requisiciones"))

    bloques = [

        {"titulo": "Indicadores", "df": ind},

        {"titulo": "Por estado actual (etapa)", "df": _conteo(t, "Etapa actual", "Etapa actual"), "grafico": "bar", "totales": True},

        {"titulo": "Por contrato", "df": _conteo(t, "Contrato", "Contrato"), "grafico": "bar", "totales": True},

        {"titulo": "Por zona", "df": _conteo(t, "Zona", "Zona"), "grafico": "bar", "totales": True},

        {"titulo": "Tiempos del proceso", "df": tiempos},

        {"titulo": "Por antigüedad", "df": tramos, "grafico": "column"},

    ]

    if len(alert):

        bloques.append({"titulo": "Alertas", "df": alert})

    hojas = [

        {"nombre": "Resumen", "titulo": f"Reporte de requisiciones – {tipo}", "bloques": bloques,

         "nota": TIPOS_REPORTE.get(tipo, "")},

        {"nombre": "Requisiciones", "df": det,

         "titulo": f"Requisiciones – {tipo}",

         "nota": "Una fila por requisición · días calculados al " + hoy.strftime("%d-%m-%Y"),

         "semaforo": {"Antigüedad (días)": (15, 30), "Días en evaluación": (7, 15),

                      "Días desde último comité": (7, 14), "Días cotización → 1er comité": (7, 14)},

         "anchos": {"Historial comité": 60, "Aprobado en comité": 50, "Alertas": 34, "Comentario último comité": 40,

                    "Comentario gestión": 34, "Comentario resumen": 34, "Proveedores cotizados": 34},

         "congelar_col": 2},

    ]

    nro = dict(zip(t["LLave"], t["N° requisición"]))

    def sub(df, cols_quitar=()):

        if df is None or df.empty or "LLave" not in df:

            return pd.DataFrame()

        d = df[df["LLave"].isin(llaves)].copy()

        if "N° requisición" not in d and "N° de requisición" not in d:

            d.insert(0, "N° requisición", d["LLave"].map(nro))

        return d.drop(columns=["LLave", *cols_quitar], errors="ignore")

    hc = sub(comite_df)

    if not hc.empty:

        hojas.append({"nombre": "Historial comité", "df": hc.sort_values(["N° de requisición", "Fecha comité"]),

                      "nota": "Todas las pasadas por comité de estas requisiciones", "totales": False})

    if not hc.empty:

        apr = hc[hc["Resultado comité"].astype(str).str.strip().str.lower().str.startswith("aprob")]

        if not apr.empty:

            info = t.set_index("N° requisición")[["Contrato", "Zona", "Solicitante", "Etapa actual"]]

            apr = apr.join(info, on="N° de requisición")

            cols = ["N° de requisición", "Contrato", "Zona", "Solicitante", "Comité", "Fecha comité",

                    "Resultado comité", "Detalle cotización", "Monto neto comité", "Comentario comité", "Etapa actual"]

            hojas.insert(2, {"nombre": "Aprobado en comité",

                             "df": apr[[c for c in cols if c in apr]].sort_values(["N° de requisición", "Fecha comité"]),

                             "nota": "Lo aprobado (total o parcial) en el acta del comité, con su monto neto",

                             "anchos": {"Detalle cotización": 50, "Comentario comité": 40}})

    it = sub(requisiciones_df)

    if not it.empty:

        hojas.append({"nombre": "Ítems solicitados", "df": it.sort_values("N° requisición", kind="stable")})

    co = sub(cotizaciones_df)

    if not co.empty:

        hojas.append({"nombre": "Cotizaciones", "df": co.sort_values("N° requisición", kind="stable"),

                      "totales": False})

    oc = sub(detalle_oc_df)

    if not oc.empty:

        hojas.append({"nombre": "OC emitidas", "df": oc.sort_values(["N° de requisición", "OC"])})

    return excel_formato(hojas, titulo="Reporte de requisiciones", filtros=filtros_txt)


# =========================================================
# PESTAÑA "Reportes"
# =========================================================

def mostrar_reportes(df, comite_df, requisiciones_df, cotizaciones_df, detalle_oc_df):

    seccion(

        "Reporte por requisición",

        "Usa los filtros de arriba (contrato, zona, empresa...) y elige qué reporte quieres. "

        "Ej.: Contrato = Flota + 'Requisiciones sin OC emitida'."

    )

    c1, c2 = st.columns([3, 1], vertical_alignment="bottom")

    tipo = c1.selectbox("Tipo de reporte", list(TIPOS_REPORTE), key="reporte_tipo")

    incluir_anuladas = c2.checkbox("Incluir anuladas", value=False, key="reporte_anuladas")

    st.caption(TIPOS_REPORTE[tipo])

    hoy = _hoy()

    t = filtrar_tipo(tabla_requisiciones(df, comite_df, cotizaciones_df, hoy), tipo, incluir_anuladas)

    if t.empty:

        aviso("No hay requisiciones para este reporte con los filtros actuales.", "info")

        return

    con_alerta = int((t["Alertas"] != "").sum())

    kpis([

        {"etiqueta": "Requisiciones", "valor": f"{len(t):,}".replace(",", ".")},

        {"etiqueta": "Antigüedad promedio", "valor": f"{t['Antigüedad (días)'].mean():.0f} días"

         if t["Antigüedad (días)"].notna().any() else "—",

         "detalle": f"Máx. {int(t['Antigüedad (días)'].max())} días" if t["Antigüedad (días)"].notna().any() else ""},

        {"etiqueta": "Monto aprobado en comité", "valor": f"$ {t['Monto aprobado (neto)'].sum():,.0f}".replace(",", "."),

         "detalle": f"Cotizado: $ {t['Monto cotización'].sum():,.0f}".replace(",", ".")},

        {"etiqueta": "Con alertas", "valor": f"{con_alerta}", "tono": "danger" if con_alerta else "ok"},

    ])

    vista = [

        "N° requisición", "Etapa actual", "Contrato", "Zona", "Solicitante", "Fecha Requisición",

        "Fecha Cotización", "Primer comité", "Fecha primer comité", "Último comité (acta)",

        "Fecha último comité (acta)", "N° de comités", "Resultado último comité",

        "Aprobado en comité", "Monto aprobado (neto)",

        "Antigüedad (días)", "Días en evaluación", "Monto cotización", "Alertas", "Comentario último comité",

    ]

    st.dataframe(

        t.sort_values("Antigüedad (días)", ascending=False)[[c for c in vista if c in t]],

        width="stretch", hide_index=True, height=min(80 + 36 * len(t), 560),

        column_config={

            "Monto cotización": st.column_config.NumberColumn(format="$ %d"),

            "Monto aprobado (neto)": st.column_config.NumberColumn(format="$ %d"),

            "Fecha Requisición": st.column_config.DateColumn(format="DD-MM-YYYY"),

            "Fecha Cotización": st.column_config.DateColumn(format="DD-MM-YYYY"),

            "Fecha primer comité": st.column_config.DateColumn(format="DD-MM-YYYY"),

            "Fecha último comité (acta)": st.column_config.DateColumn(format="DD-MM-YYYY"),

        },

    )

    filtros_txt = st.session_state.get("_filtros_txt", "") + "  ·  Reporte: " + tipo + (

        " (incluye anuladas)" if incluir_anuladas else "")

    contrato = "_".join(st.session_state.get("filtro_contrato", []) or [])[:30]

    nombre = "Reporte_" + tipo.split(" (")[0].replace(" ", "_").replace(",", "") + (

        f"_{contrato}" if contrato else "") + f"_{hoy:%Y%m%d}.xlsx"

    st.download_button(

        "Descargar reporte (.xlsx)",

        generar_excel(t, tipo, filtros_txt, comite_df, requisiciones_df, cotizaciones_df, detalle_oc_df, hoy),

        file_name="".join(ch for ch in nombre if ch.isalnum() or ch in "._-"),

        mime=MIME_XLSX,

        type="primary",

        key="descargar_reporte_req",

    )
