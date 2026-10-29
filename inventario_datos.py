# -*- coding: utf-8 -*-
"""
=========================================================
Dashboard Compras
Inventario (checklists de Google Forms) - LIMPIEZA
=========================================================
Lee directamente los CSV descargados de las planillas de
respuestas de Google Forms y los deja limpios:

PERSONAS ("Check list por cargo")
  - Se conservan solo los registros con RUT (los sin RUT
    se descartan).
  - El RUT se normaliza sin puntos y con guion: 17374225-8.
  - Se deja el último registro por RUT (según Marca temporal).
  - Las columnas repetidas por cargo (".1", ".2", ...) se
    juntan en una sola por ítem.

MÓVILES ("Check List Implementación de Móvil")
  - La patente se normaliza (mayúsculas, sin espacios).
  - Se deja el último registro por patente según la FECHA
    INSPECCION (y la Marca temporal para desempatar).

Cada ítem se clasifica en: OK, Falta, Requiere cambio,
Incompleto, Regular, No aplica o Revisar.
"""

import re
import unicodedata

import pandas as pd

from config import BASE_DIR


# ======================================================
# ARCHIVOS
# ======================================================
# Se busca en la carpeta del repo cualquier CSV cuyo nombre
# calce con estos patrones (así se puede subir el archivo tal
# como lo descarga Google Sheets, sin renombrarlo).

PATRON_PERSONAS = re.compile(r"por.?cargo.*\.csv$", re.I)

PATRON_MOVILES = re.compile(r"m[oó]vil.*\.csv$", re.I)


def buscar_archivo(patron):

    archivos = [

        p for p in BASE_DIR.iterdir()

        if p.is_file() and patron.search(p.name)

    ]

    if not archivos:
        return None

    return max(archivos, key=lambda p: p.stat().st_mtime)


def leer_csv(ruta):

    for codificacion in ("utf-8-sig", "latin-1"):

        try:
            return pd.read_csv(ruta, dtype=str, encoding=codificacion)
        except UnicodeDecodeError:
            continue

    return pd.read_csv(ruta, dtype=str, encoding="latin-1", errors="replace")


# ======================================================
# UTILIDADES
# ======================================================

ESTADOS_ORDEN = ["Falta", "Requiere cambio", "Incompleto", "Regular", "Revisar", "OK", "No aplica"]

ESTADOS_PENDIENTES = ["Falta", "Requiere cambio", "Incompleto"]


def parse_marca_temporal(serie):

    """'2026/07/13 9:08:00 a. m. GMT-3' -> datetime"""

    s = (

        serie.fillna("")
        .astype(str)
        .str.replace(" ", " ", regex=False)
        .str.replace("\xa0", " ", regex=False)
        .str.replace(r"\s*GMT[+-]?\d*", "", regex=True)
        .str.replace(r"a\.\s*m\.", "AM", regex=True)
        .str.replace(r"p\.\s*m\.", "PM", regex=True)
        .str.strip()

    )

    fecha = pd.to_datetime(s, format="%Y/%m/%d %I:%M:%S %p", errors="coerce")

    # Respaldo por si viene en otro formato (ej. 24 horas)
    faltan = fecha.isna() & (s != "")

    if faltan.any():
        fecha[faltan] = pd.to_datetime(s[faltan], errors="coerce", dayfirst=False)

    return fecha


def sin_tildes(texto):

    return "".join(

        c for c in unicodedata.normalize("NFD", str(texto))

        if unicodedata.category(c) != "Mn"

    )


def normalizar_rut(valor):

    """'17.374.225-8', '173742258', '17374225-8 ' -> '17374225-8'"""

    if valor is None or pd.isna(valor):
        return None

    limpio = re.sub(r"[^0-9kK]", "", str(valor)).upper()

    if len(limpio) < 2:
        return None

    cuerpo, dv = limpio[:-1].lstrip("0"), limpio[-1]

    if not cuerpo.isdigit():
        return None

    return f"{cuerpo}-{dv}"


def rut_valido(rut):

    if not rut:
        return False

    cuerpo, dv = rut.split("-")

    suma = sum(

        int(d) * f

        for d, f in zip(reversed(cuerpo), [2, 3, 4, 5, 6, 7] * 3)

    )

    resto = 11 - suma % 11

    esperado = "0" if resto == 11 else "K" if resto == 10 else str(resto)

    return dv == esperado


def nombre_base(columna):

    """'Estado de EPP [Casco].2' -> 'Estado de EPP [Casco]'"""

    return re.sub(r"\.\d+$", "", str(columna)).strip()


def primero_no_vacio(df, cols):

    """Primer valor no vacío entre varias columnas (más rápido que bfill)."""

    cols = list(cols)

    resultado = df[cols[0]].copy()

    for c in cols[1:]:

        resultado = resultado.where(resultado.notna(), df[c])

    return resultado


def juntar_columnas_repetidas(df):

    """
    Google Forms repite la misma pregunta en cada sección
    (por cargo) y el CSV la nombra 'X', 'X.1', 'X.2'... Aquí
    se juntan en una sola columna tomando el primer valor no
    vacío. Las 'Observaciones' se concatenan.
    """

    grupos = {}

    for c in df.columns:

        grupos.setdefault(nombre_base(c), []).append(c)

    salida = {}

    for base, cols in grupos.items():

        if len(cols) == 1:

            salida[base] = df[cols[0]]

        elif base.lower().startswith("observaciones"):

            salida[base] = df[cols].apply(

                lambda fila: " | ".join(

                    str(v).strip() for v in fila

                    if pd.notna(v) and str(v).strip()

                ) or None,

                axis=1

            )

        else:

            salida[base] = primero_no_vacio(df, cols)

    return pd.DataFrame(salida, index=df.index)


def limpiar_texto(valor):

    if valor is None or pd.isna(valor):
        return None

    v = re.sub(r"\s+", " ", str(valor)).strip()

    return v or None


def normalizar_servicio(valor):

    v = limpiar_texto(valor)

    if not v:
        return "Sin servicio"

    t = v.lower()

    if "perdida" in t or "pérdida" in t:
        return "Control de pérdidas"

    if "mantenimiento" in t:
        return "CYM"

    if t.startswith("multiprop"):
        return "Multipropósito"

    return v


def normalizar_item(texto):

    t = re.sub(r"\s+", " ", str(texto)).strip()

    # Variantes de escritura que aparecen en el formulario
    t = re.sub(r"di[ae]l[eé]ct?ir?c([oa])", r"dieléctric\1", t, flags=re.I)

    return t[:1].upper() + t[1:]


# ======================================================
# CLASIFICACIÓN DE RESPUESTAS
# ======================================================

def estado_persona(valor):

    v = limpiar_texto(valor)

    if not v:
        return None

    t = v.upper()

    codigos = {

        "BC": "OK", "B": "OK", "BUENO": "OK",

        "NT": "Falta",

        "NC": "Requiere cambio", "M": "Requiere cambio", "MALO": "Requiere cambio",

        "R": "Regular", "REGULAR": "Regular",

        "NO APLICA": "No aplica", "N/A": "No aplica", "N/M": "No aplica"

    }

    return codigos.get(t, "Revisar")


PALABRAS_NO_APLICA = ("NO APLICA", "NO SPLICA", "N/A", "MN/A", "NO USA", "NO SE REQUIERE", "NO  SÉ REQUIERE", "NO SÉ REQUIERE")

PALABRAS_CAMBIO = ("MAL ESTADO", "MALO", "DAÑ", "CAMBI", "REEMPLAZ", "DETERIOR", "GASTAD", "DESGAST", "ROTA", "ROTO", "DOBLAD", "PARCHAD")

PALABRAS_INCOMPLETO = ("INCOMPLETO", "FALTAN UNIDADES", "SOLO ", "FALTA SOPORTE", "SIN PELDAÑOS", "NO TIENE HOJA")

PALABRAS_FALTA = ("NO TIENE", "NO HAY", "NO TENEMOS", "NO CUENTA", "NO POSEE", "NO EXISTE", "NO TENGO", "NO CONTAMOS", "SIN ")


def estado_movil(valor, esperado=None):

    v = limpiar_texto(valor)

    if not v or v in (".", ",", "-"):
        return None

    t = v.upper()

    if t in ("NA", "N", "NI", "NN") or any(p in t for p in PALABRAS_NO_APLICA):
        return "No aplica"

    if any(p in t for p in PALABRAS_CAMBIO):
        return "Requiere cambio"

    partes = [p.strip() for p in t.split(";") if p.strip()]

    if "SI" in partes and "NO" in partes:
        return "Incompleto"

    if any(p in t for p in PALABRAS_INCOMPLETO):
        return "Incompleto"

    # Cantidades: "2", "6 CONOS", "5 EN BUEN ESTADO"
    numero = re.match(r"^(\d+)\b", t)

    if numero:

        cantidad = int(numero.group(1))

        if esperado and cantidad < esperado:
            return "Incompleto"

        return "OK"

    if t in ("NO", "NO O", "NO MI", "NO TIENEO") or any(p in t for p in PALABRAS_FALTA):
        return "Falta"

    if t.startswith("SI") or t in ("BUEN ESTADO", "COMPLETO", "APLICA", "USADA") or "BUEN ESTADO" in t:
        return "OK"

    return "Revisar"


# ======================================================
# PERSONAS
# ======================================================

CATEGORIAS_PERSONAS = {

    "estado de epp jefe de faena": "EPP jefe de faena",

    "estado de epp": "EPP",

    "herramientas": "Herramientas",

    "equipamiento": "Equipamiento",

    "descripción": "Vehículo: descripción",

    "seguridad": "Vehículo: seguridad",

    "mecánico": "Vehículo: mecánico",

    "accesorios": "Vehículo: accesorios",

    "normativa del tránsito": "Vehículo: normativa",

    "implementación": "Vehículo: implementación",

    "kit básico obligatorio por ley": "Vehículo: kit por ley",

    "equipo técnico": "Equipo técnico"

}

# Para comprar ropa/calzado sirve saber la talla
TALLA_POR_ITEM = [

    (("pantal",), "Talla Pantalón"),

    (("calzado", "bota", "zapato"), "Talla zapatos"),

    (("parka", "polar", "geólogo", "geologo", "camisa", "polera", "chaqueta", "traje de agua", "legionario"), "Talla Torso"),

]


def talla_para(item):

    t = item.lower()

    for palabras, campo in TALLA_POR_ITEM:

        if any(p in t for p in palabras):
            return campo

    return None


def limpiar_personas(ruta):

    crudo = leer_csv(ruta)

    calidad = {"Registros en el archivo": len(crudo)}

    sin_encabezado = [c for c in crudo.columns if str(c).startswith("Unnamed")]

    con_datos_sin_encabezado = [c for c in sin_encabezado if crudo[c].notna().any()]

    calidad["Columnas sin nombre de pregunta (ignoradas)"] = len(con_datos_sin_encabezado)

    crudo = crudo.drop(columns=sin_encabezado)

    # El RUT puede venir en "RUT" o en "RUT.1" (otra sección)
    cols_rut = [c for c in crudo.columns if nombre_base(c) == "RUT"]

    rut_crudo = primero_no_vacio(crudo, cols_rut)

    crudo["_ts"] = parse_marca_temporal(crudo["Marca temporal"])

    crudo["RUT_norm"] = rut_crudo.map(normalizar_rut)

    # Se deduplica ANTES de juntar columnas (mucho más rápido)
    crudo = crudo.sort_values("_ts", kind="stable")

    sin_rut_crudo = crudo[crudo["RUT_norm"].isna()]

    con_rut_crudo = crudo[crudo["RUT_norm"].notna()]

    ultimos_crudo = con_rut_crudo.drop_duplicates("RUT_norm", keep="last")

    df = juntar_columnas_repetidas(ultimos_crudo.drop(columns=["_ts", "RUT_norm"]))

    df["_ts"] = ultimos_crudo["_ts"]

    df["RUT_norm"] = ultimos_crudo["RUT_norm"]

    ident = juntar_columnas_repetidas(

        crudo[[c for c in crudo.columns if nombre_base(c) in ("Marca temporal", "Empresa", "Nombre y apellido", "Zonal", "Servicio")]]

    )

    ident["_ts"] = crudo["_ts"]

    ident["RUT_norm"] = crudo["RUT_norm"]

    calidad["Descartados sin RUT"] = int(ident["RUT_norm"].isna().sum())

    con_rut = ident[ident["RUT_norm"].notna()]

    calidad["Registros con RUT"] = len(con_rut)

    # Revisión: mismo RUT con primeros nombres distintos
    primeros = (

        con_rut.assign(_n=con_rut["Nombre y apellido"].fillna("").map(sin_tildes).str.strip().str.lower().str.split().str[0])
        .groupby("RUT_norm")["_n"]
        .agg(lambda s: sorted({x for x in s if x}))

    )

    ruts_conflicto = primeros[primeros.map(len) > 1]

    ultimo = df.copy()

    calidad["Duplicados eliminados (se dejó el último por RUT)"] = len(con_rut) - len(ultimo)

    calidad["Personas únicas"] = len(ultimo)

    ultimo["RUT válido"] = ultimo["RUT_norm"].map(rut_valido)

    calidad["RUT con dígito verificador inválido"] = int((~ultimo["RUT válido"]).sum())

    calidad["RUT con nombres distintos (revisar)"] = len(ruts_conflicto)

    # ---------- ficha por persona ----------
    cargo_cols = [c for c in ultimo.columns if c.upper().startswith("CARGO")]

    fichas = pd.DataFrame({

        "RUT": ultimo["RUT_norm"],

        "Nombre": ultimo["Nombre y apellido"].map(limpiar_texto).fillna("Sin nombre").str.title(),

        "Empresa": ultimo["Empresa"].map(limpiar_texto).fillna("Sin empresa"),

        "Zonal": ultimo["Zonal"].map(limpiar_texto).fillna("Sin zonal"),

        "Servicio": ultimo["Servicio"].map(normalizar_servicio),

        "Cargo": primero_no_vacio(ultimo, cargo_cols).map(limpiar_texto).fillna("Sin cargo") if cargo_cols else "Sin cargo",

        "Talla Torso": ultimo.get("Talla Torso").map(limpiar_texto),

        "Talla Pantalón": ultimo.get("Talla Pantalón").map(limpiar_texto),

        "Talla zapatos": ultimo.get("Talla zapatos").map(limpiar_texto),

        "Última revisión": ultimo["_ts"],

        "Observaciones": ultimo.get("Observaciones"),

        "RUT válido": ultimo["RUT válido"]

    }).reset_index(drop=True)

    for c in ["Talla Torso", "Talla Pantalón", "Talla zapatos"]:

        fichas[c] = fichas[c].str.upper().str.strip(" /.-")

    # ---------- ítems (formato largo) ----------
    items_cols = {}

    for c in ultimo.columns:

        m = re.match(r"^(.+?)\s*\[(.+)\]\s*$", c)

        if m:

            categoria = CATEGORIAS_PERSONAS.get(m.group(1).strip().lower(), m.group(1).strip().title())

            items_cols[c] = (categoria, normalizar_item(m.group(2)))

    largo = (

        ultimo[["RUT_norm"] + list(items_cols)]
        .melt(id_vars="RUT_norm", var_name="_col", value_name="Valor registrado")
        .dropna(subset=["Valor registrado"])

    )

    largo["Categoría"] = largo["_col"].map(lambda c: items_cols[c][0])

    largo["Ítem"] = largo["_col"].map(lambda c: items_cols[c][1])

    largo["Estado"] = largo["Valor registrado"].map(estado_persona)

    largo = (

        largo.dropna(subset=["Estado"])
        .drop_duplicates(["RUT_norm", "Categoría", "Ítem"], keep="first")
        .rename(columns={"RUT_norm": "RUT"})
        .drop(columns="_col")

    )

    largo["Campo talla"] = largo["Ítem"].map(talla_para)

    largo = largo.merge(fichas, on="RUT", how="left")

    largo["Talla"] = largo.apply(

        lambda f: f[f["Campo talla"]] if isinstance(f["Campo talla"], str) else None,

        axis=1

    )

    conflictos = pd.DataFrame({

        "RUT": ruts_conflicto.index,

        "Nombres registrados": [", ".join(n.title() for n in v) for v in ruts_conflicto]

    })

    invalidos = fichas.loc[~fichas["RUT válido"], ["RUT", "Nombre", "Zonal", "Servicio"]]

    descartados = ident[ident["RUT_norm"].isna()][["Marca temporal", "Empresa", "Nombre y apellido", "Zonal", "Servicio"]]

    return {

        "fichas": fichas,

        "items": largo.reset_index(drop=True),

        "calidad": calidad,

        "conflictos": conflictos,

        "invalidos": invalidos,

        "descartados": descartados

    }


# ======================================================
# MÓVILES
# ======================================================

MULTISELECCION = {

    "KIT DE RESCATE": ["CUERDA DE DESCENSO", "MOSQUETON (1)", "CINTA ANCLAJE MOVIL", "DESENDEDOR EDDY", "CORVO", "BOLSO"],

    "KIT REPUESTO MOVIL": ["RUEDA REPUESTO", "GATA", "LLAVE RUEDA", "VARILLA", "TRIANGULO", "EXTINTOR"],

    "LLAVES MEDIDORES": ["LLAVE FLOR", "LLAVE 5 PUNTAS", "LLAVE 7 PUNTAS", "LLAVE TRIANGULO", "LLAVE VIELA"],

}

COLUMNAS_NO_ITEM_MOVIL = {

    "Marca temporal", "Empresa", "Area", "BRIGADA RESPONSABLE DE INSPECCION",

    "FECHA INSPECCION", "BRIGADA QUE ENTREGA EL MOVIL", "Patente"

}


def bonito(texto):

    t = re.sub(r"\s+", " ", str(texto)).strip()

    t = t[:1].upper() + t[1:].lower()

    # Siglas que deben quedar en mayúscula
    return re.sub(r"\b(bt|mt|kg|mm|cm)\b", lambda m: m.group(1).upper(), t)


def opciones_marcadas(valor):

    if valor is None or pd.isna(valor):
        return None

    return {

        re.sub(r"\s+\d+$", "", p.strip().upper())

        for p in str(valor).split(";") if p.strip()

    }


def limpiar_moviles(ruta):

    crudo = leer_csv(ruta)

    calidad = {"Registros en el archivo": len(crudo)}

    df = crudo.copy()

    df["_ts"] = parse_marca_temporal(df["Marca temporal"])

    df["_fecha"] = pd.to_datetime(df["FECHA INSPECCION"], errors="coerce")

    # Fechas digitadas con año inválido (ej. 0026-03-28): se usa
    # el día de la marca temporal.
    malas = df["_fecha"].isna() | (df["_fecha"].dt.year < 2020)

    calidad["Fechas de inspección corregidas"] = int(malas.sum())

    df.loc[malas, "_fecha"] = df.loc[malas, "_ts"].dt.normalize()

    df["Patente_norm"] = df["Patente"].fillna("").str.upper().str.replace(r"[^A-Z0-9]", "", regex=True)

    df = df[df["Patente_norm"] != ""]

    df = df.sort_values(["_fecha", "_ts"], kind="stable")

    ultimo = df.drop_duplicates("Patente_norm", keep="last").copy()

    calidad["Duplicados eliminados (se dejó el último por patente)"] = len(df) - len(ultimo)

    calidad["Móviles únicos"] = len(ultimo)

    # Llaves de medidores: hay columnas sin nombre que son
    # versiones anteriores de la misma pregunta.
    cols_llaves = ["LLAVES MEDIDORES"] + [c for c in ultimo.columns if str(c).startswith("Unnamed")]

    def primera_llave(fila):

        valores = [fila[c] for c in cols_llaves if c in fila and pd.notna(fila[c]) and str(fila[c]).strip()]

        utiles = [v for v in valores if str(v).strip().upper() != "NO APLICA"]

        return utiles[0] if utiles else (valores[0] if valores else None)

    ultimo["LLAVES MEDIDORES"] = ultimo.apply(primera_llave, axis=1)

    fotos = [c for c in ultimo.columns if str(c).upper().startswith("IMAGEN")]

    fichas = pd.DataFrame({

        "Patente": ultimo["Patente_norm"],

        "Empresa": ultimo["Empresa"].map(limpiar_texto).fillna("Sin empresa"),

        "Área": ultimo["Area"].map(normalizar_servicio),

        "Brigada responsable": ultimo["BRIGADA RESPONSABLE DE INSPECCION"].map(limpiar_texto),

        "Brigada que entrega": ultimo["BRIGADA QUE ENTREGA EL MOVIL"].map(limpiar_texto),

        "Fecha inspección": ultimo["_fecha"],

        "Foto caja herramientas": ultimo[fotos[0]] if fotos else None,

        "Foto interior cúpula": ultimo[fotos[1]] if len(fotos) > 1 else None,

    }).reset_index(drop=True)

    # ---------- ítems ----------
    filas = []

    columnas_item = [

        c for c in ultimo.columns

        if c not in COLUMNAS_NO_ITEM_MOVIL and c not in fotos

        and not str(c).startswith(("_", "Unnamed", "Patente_norm"))

    ]

    for _, fila in ultimo.iterrows():

        for c in columnas_item:

            valor = fila[c]

            if c in MULTISELECCION:

                marcadas = opciones_marcadas(valor)

                if marcadas is None:
                    continue

                texto = str(valor).upper()

                if marcadas == {"NO APLICA"}:
                    estado_general = "No aplica"
                elif any(p in texto for p in ("NO CUENTA", "SIN LLAVE", "NINGUNA", "NO TIENE", "NO TENEMOS", "NO SE CUENTA")):
                    estado_general = "Falta"
                else:
                    estado_general = None

                for opcion in MULTISELECCION[c]:

                    if estado_general:
                        est = estado_general
                    else:
                        est = "OK" if opcion in marcadas else "Falta"

                    filas.append((fila["Patente_norm"], bonito(c), bonito(opcion), valor, est))

                continue

            esperado = re.search(r"\((\d+)\)", str(c))

            est = estado_movil(valor, int(esperado.group(1)) if esperado else None)

            if est:
                filas.append((fila["Patente_norm"], "Implementación", bonito(c), valor, est))

    largo = pd.DataFrame(filas, columns=["Patente", "Categoría", "Ítem", "Valor registrado", "Estado"])

    largo = largo.merge(fichas[["Patente", "Empresa", "Área", "Fecha inspección"]], on="Patente", how="left")

    return {

        "fichas": fichas,

        "items": largo,

        "calidad": calidad

    }
