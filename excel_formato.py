# -*- coding: utf-8 -*-
"""
Exportación a Excel con formato de tabla (mismo formato que la app de Resultados por empresa).

Cada hoja lleva: título, filtros aplicados y fecha de generación; encabezado azul,
filas alternadas, montos con separador de miles, fechas dd-mm-aaaa, columnas
ajustadas, panel inmovilizado, autofiltro y configuración para imprimir.

Uso:
    excel_formato({"Hoja": df, ...}, titulo="Compras", filtros="Zona: Talca")
    excel_formato([{"nombre": "Hoja", "df": df, "titulo": "...", "nota": "..."}], filtros="...")
"""
import io
from datetime import datetime

import pandas as pd

AZUL = "#1E3A8A"

# columnas que se muestran como monto ($ con separador de miles)
_PALABRAS_MONTO = ("monto", "valor", "total", "neto", "iva", "subtotal", "precio", "costo", "gasto", "pago", "presupuesto")


def _es_monto(nombre):
    n = str(nombre).lower()
    return any(p in n for p in _PALABRAS_MONTO)


def _tipo_columna(serie, nombre):
    if pd.api.types.is_datetime64_any_dtype(serie):
        return "fecha"
    if pd.api.types.is_bool_dtype(serie):
        return "txt"
    if pd.api.types.is_numeric_dtype(serie):
        if _es_monto(nombre):
            return "monto"
        no_nulos = serie.dropna()
        if len(no_nulos) and (no_nulos % 1 != 0).any():
            return "decimal"
        return "entero"
    # columnas de texto que en realidad traen fechas (Timestamp)
    muestra = serie.dropna().head(50)
    if len(muestra) and all(isinstance(v, (pd.Timestamp, datetime)) for v in muestra):
        return "fecha"
    return "txt"


def excel_formato(hojas, titulo="", filtros=""):
    """Devuelve los bytes de un .xlsx con formato de tabla."""
    if isinstance(hojas, dict):
        hojas = [{"nombre": k, "df": v} for k, v in hojas.items()]

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter", engine_kwargs={"options": {"strings_to_formulas": False,
                                                                             "strings_to_urls": False}}) as xw:
        wb = xw.book
        cache = {}
        NUMF = {"monto": '"$"#,##0;[Red]-"$"#,##0', "entero": "0", "decimal": "#,##0.0",
                "fecha": "dd-mm-yyyy", "txt": None}

        def fmt(tipo, zebra):
            k = (tipo, zebra)
            if k not in cache:
                p = {"font_name": "Calibri", "font_size": 10, "valign": "top", "text_wrap": tipo == "txt",
                     "bottom": 1, "bottom_color": "#E2E8F0"}
                if NUMF[tipo]:
                    p["num_format"] = NUMF[tipo]
                if tipo == "fecha":
                    p["align"] = "center"
                if zebra:
                    p["bg_color"] = "#F5F8FF"
                cache[k] = wb.add_format(p)
            return cache[k]

        f_tit = wb.add_format({"font_name": "Calibri", "bold": True, "font_size": 15, "font_color": AZUL})
        f_sub = wb.add_format({"font_name": "Calibri", "italic": True, "font_size": 9, "font_color": "#475569"})
        f_head = wb.add_format({"font_name": "Calibri", "font_size": 10, "bold": True, "font_color": "#FFFFFF", "bg_color": AZUL,
                                "text_wrap": True, "valign": "vcenter", "align": "center", "border": 1, "border_color": "#FFFFFF"})
        f_tot = wb.add_format({"font_name": "Calibri", "font_size": 10, "bold": True, "bg_color": "#DBEAFE",
                               "top": 2, "top_color": AZUL, "num_format": NUMF["monto"]})
        f_tot_lbl = wb.add_format({"font_name": "Calibri", "font_size": 10, "bold": True, "bg_color": "#DBEAFE",
                                   "top": 2, "top_color": AZUL})

        f_bloque = wb.add_format({"font_name": "Calibri", "bold": True, "font_size": 11, "font_color": AZUL,
                                  "bottom": 2, "bottom_color": AZUL})
        f_rojo = wb.add_format({"bg_color": "#FADBD8", "font_color": "#922B21", "bold": True})
        f_ambar = wb.add_format({"bg_color": "#FDEBD0", "font_color": "#9C640C"})

        def escribir_tabla(ws, d, r0, totales=True, semaforo=None):
            """Escribe una tabla con encabezado en la fila r0. Devuelve (última fila, tipos, anchos)."""
            tipos = [_tipo_columna(d[c], c) for c in d.columns]
            ws.set_row(r0, 32)
            for j, c in enumerate(d.columns):
                ws.write_string(r0, j, c, f_head)

            for i, fila in enumerate(d.itertuples(index=False, name=None)):
                rr = r0 + 1 + i
                zebra = i % 2 == 1
                for j, v in enumerate(fila):
                    t = tipos[j]
                    f = fmt(t, zebra)
                    if v is None or (not isinstance(v, (str, list, tuple)) and pd.isna(v)):
                        ws.write_blank(rr, j, None, f)
                    elif t == "fecha":
                        try:
                            ws.write_datetime(rr, j, pd.Timestamp(v).to_pydatetime().replace(tzinfo=None), f)
                        except (ValueError, TypeError):
                            ws.write_string(rr, j, str(v), fmt("txt", zebra))
                    elif t in ("monto", "entero", "decimal"):
                        ws.write_number(rr, j, float(v), f)
                    else:
                        ws.write_string(rr, j, str(v), f)

            # semáforo: {columna: (umbral_ámbar, umbral_rojo)}
            if semaforo and len(d):
                for col, (amb, rojo) in semaforo.items():
                    if col in d.columns:
                        j = list(d.columns).index(col)
                        ws.conditional_format(r0 + 1, j, r0 + len(d), j,
                                              {"type": "cell", "criteria": ">", "value": rojo, "format": f_rojo})
                        ws.conditional_format(r0 + 1, j, r0 + len(d), j,
                                              {"type": "cell", "criteria": "between", "minimum": amb + 1,
                                               "maximum": rojo, "format": f_ambar})

            # fila de totales para las columnas de monto
            # (no se suman precios unitarios)
            montos = [j for j, t in enumerate(tipos) if t == "monto" and "unit" not in d.columns[j].lower() and "valor un" not in d.columns[j].lower()]
            ultima = r0 + len(d)
            if montos and len(d) > 1 and totales:
                rt = ultima + 1
                for j in range(len(d.columns)):
                    if j in montos:
                        ws.write_number(rt, j, float(pd.to_numeric(d.iloc[:, j], errors="coerce").sum()), f_tot)
                    else:
                        ws.write_string(rt, j, "Total" if j == 0 else "", f_tot_lbl)
                ultima = rt

            anchos = []
            for j, c in enumerate(d.columns):
                if tipos[j] == "fecha":
                    ancho = 12
                else:
                    vals = d[c].head(300).dropna().astype(str).tolist()
                    largo = max([min(len(c), 20)] + [len(v) for v in vals]) if vals else len(c)
                    if tipos[j] == "monto":
                        largo += 4
                    ancho = min(max(largo + 2, 9), 50)
                anchos.append(ancho)
            return ultima, tipos, anchos

        def preparar(df):
            d = df.copy()
            if not isinstance(d.index, pd.RangeIndex):
                d = d.reset_index(drop=d.index.name is None and not isinstance(d.index, pd.MultiIndex))
            d.columns = [str(c) for c in d.columns]
            return d

        usados = set()
        for h in hojas:
            nombre = "".join(ch for ch in str(h["nombre"]) if ch not in '[]:*?/\\')[:31] or "Hoja"
            while nombre in usados:
                nombre = nombre[:28] + f"_{len(usados)}"
            usados.add(nombre)
            ws = wb.add_worksheet(nombre)

            tit = h.get("titulo") or (f"{titulo} – {h['nombre']}" if titulo and titulo != h["nombre"] else str(h["nombre"]))
            ws.write_string(0, 0, tit, f_tit)
            ws.write_string(1, 0, filtros or "Sin filtros", f_sub)

            # ---------------------------------------------
            # Hoja de varias tablas apiladas (p. ej. Resumen):
            # h["bloques"] = [{"titulo": ..., "df": ..., "grafico": "bar"|"pie"|None}, ...]
            # ---------------------------------------------
            if h.get("bloques"):
                ws.write_string(2, 0, (h.get("nota") + "  ·  " if h.get("nota") else "") +
                                f"Generado el {datetime.now():%d-%m-%Y %H:%M}", f_sub)
                r = 4
                anchos_max = {}
                for b in h["bloques"]:
                    d = preparar(b["df"])
                    ws.write_string(r, 0, b.get("titulo", ""), f_bloque)
                    r0 = r + 1
                    ultima, tipos, anchos = escribir_tabla(ws, d, r0, totales=b.get("totales", False))
                    for j, a in enumerate(anchos):
                        anchos_max[j] = max(anchos_max.get(j, 0), a)
                    if b.get("grafico") and len(d) and len(d.columns) >= 2:
                        ch = wb.add_chart({"type": b["grafico"]})
                        serie = {
                            "name": d.columns[1],
                            "categories": [nombre, r0 + 1, 0, r0 + len(d), 0],
                            "values": [nombre, r0 + 1, 1, r0 + len(d), 1],
                            "data_labels": {"value": True},
                        }
                        if b["grafico"] != "pie":
                            serie["fill"] = {"color": AZUL}
                        ch.add_series(serie)
                        if b["grafico"] == "bar":
                            ch.set_y_axis({"reverse": True})
                        ch.set_title({"name": b.get("titulo", ""), "name_font": {"size": 11}})
                        ch.set_legend({"none": b["grafico"] != "pie"})
                        ch.set_size({"width": 480, "height": max(220, 26 * len(d) + 90)})
                        ws.insert_chart(r0, len(d.columns) + 1, ch)
                        r = max(ultima, r0 + max(11, int(len(d) * 1.4) + 4)) + 2
                    else:
                        r = ultima + 2
                for j, a in anchos_max.items():
                    ws.set_column(j, j, min(max(a, 12), 45) if j else max(a, 30))
                ws.hide_gridlines(2)
                ws.set_landscape()
                ws.fit_to_pages(1, 0)
                continue

            d = preparar(h["df"])
            ws.write_string(2, 0, (h.get("nota") + "  ·  " if h.get("nota") else "") +
                            f"{len(d):,} filas  ·  Generado el {datetime.now():%d-%m-%Y %H:%M}".replace(",", "."), f_sub)

            r0 = 4
            ultima, tipos, anchos = escribir_tabla(ws, d, r0, totales=h.get("totales", True),
                                                   semaforo=h.get("semaforo"))

            if len(d):
                ws.autofilter(r0, 0, r0 + len(d), len(d.columns) - 1)

            for j, ancho in enumerate(h.get("anchos") and [h["anchos"].get(c, a) for c, a in zip(d.columns, anchos)] or anchos):
                ws.set_column(j, j, ancho)

            ws.freeze_panes(r0 + 1, h.get("congelar_col", 1 if len(d.columns) > 4 else 0))
            ws.set_landscape()
            ws.fit_to_pages(1, 0)
            ws.repeat_rows(r0)
            ws.hide_gridlines(2)
            ws.print_area(0, 0, max(ultima, r0), max(len(d.columns) - 1, 0))

    return buf.getvalue()


MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
