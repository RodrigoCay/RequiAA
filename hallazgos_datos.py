# -*- coding: utf-8 -*-
"""
=========================================================
Dashboard Compras
Hallazgos de Prevención de Riesgos - lectura y limpieza
=========================================================
Lee TODAS las planillas .xlsx que estén en la carpeta
"hallazgos" del repositorio (una por empresa / zona) y las
une en una sola tabla con columnas estándar.

Cada planilla puede venir con columnas distintas (la de
Colchagua trae potencialidad y control asociado, la de Maule
Sur trae fecha de cierre, etc.). Las columnas se reconocen
por su nombre, no por su posición, y la fila de títulos
puede no ser la primera (ej. "PLANILLA DE HALLAZGOS ...").

Estados:
  Abierto / Abiero          -> Abierto
  Cerrado                   -> Cerrado
  RIS (resuelto in situ)    -> Resuelto en terreno (cuenta como cerrado)
"""

import re
import unicodedata

import pandas as pd

from config import BASE_DIR


CARPETAS_HALLAZGOS = [BASE_DIR / "hallazgos", BASE_DIR / "Hallazgos"]

ABIERTO = "Abierto"
CERRADO = "Cerrado"
RIS = "Resuelto en terreno (RIS)"
SIN_ESTADO = "Sin estado"

ESTADOS_ORDEN = [ABIERTO, CERRADO, RIS, SIN_ESTADO]

# Tipo de actividad -> contrato (mismos nombres que en compras)
CONTRATOS = {
    "construccion dx": "CYM",
    "mantenimiento dx": "CYM",
    "guardia sae": "SAE",
    "servicio tecnico": "SSTT",
    "poda, tala y roce": "Poda",
    "patio de residuos": "Patio de residuos",
}

TRAMOS = [
    (0, 15, "0 a 15 días"),
    (16, 30, "16 a 30 días"),
    (31, 60, "31 a 60 días"),
    (61, 90, "61 a 90 días"),
    (91, 100000, "Más de 90 días"),
]

TRAMOS_ORDEN = [t[2] for t in TRAMOS]


# ======================================================
# UTILIDADES
# ======================================================

def norm(texto):

    if texto is None or (isinstance(texto, float) and pd.isna(texto)):
        return ""

    t = unicodedata.normalize("NFD", str(texto))
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    t = t.lower().replace("hallazgos#", "").replace("°", " ").replace("º", " ")

    return re.sub(r"\s+", " ", t).strip()


def texto(v):

    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None

    t = re.sub(r"\s+", " ", str(v)).strip()

    return t or None


def archivos_hallazgos():

    archivos = []

    for carpeta in CARPETAS_HALLAZGOS:

        if carpeta.is_dir():

            archivos += [
                p for p in carpeta.iterdir()
                if p.is_file() and p.suffix.lower() in (".xlsx", ".xlsm")
                and not p.name.startswith("~$")
            ]

    return sorted(set(archivos))


# ======================================================
# RECONOCER COLUMNAS
# ======================================================

def _campo(nombre):

    """Nombre de columna normalizado -> campo estándar (o None)."""

    h = norm(nombre)

    if not h:
        return None

    if "fecha" in h and "cierre" in h:
        return "Fecha cierre"

    if h.startswith("fecha"):
        return "Fecha"

    if "registro" in h or h.startswith("numero de hallazgo") or h in ("n", "id"):
        return "N° registro"

    if "subcontrato 2" in h:
        return None

    if "subcontrato" in h:
        return "Subcontrato"

    if "empresa" in h:
        return "Empresa"

    if h == "region":
        return "Región"

    if h == "zona":
        return "Zona"

    if h.startswith("delegacion"):
        return "Delegación"

    if "tipo de actividad" in h:
        return "Tipo de actividad"

    if "tipo de hallazgo" in h:
        return "Tipo de hallazgo"

    if "descripcion de hallazgo" in h or h == "hallazgo":
        return "Hallazgo"

    if h == "descripcion":
        return "Actividad"

    if "responsable" in h:
        return "Responsable"

    if h.startswith("ids"):
        return "IDS con hallazgos"

    if h == "estado":
        return "Estado original"

    if h.startswith("paralizacion"):
        return "Paralización"

    if h.startswith("potencialidad"):
        return "Potencialidad"

    if h.startswith("clasificacion"):
        return "Clasificación"

    if h.startswith("control asociado"):
        return "Control asociado"

    if h.startswith("observacion"):
        return "Observaciones"

    if h.startswith("medid"):
        return "Medidas correctivas"

    if "usuario" in h:
        return "Inspector"

    return None


def _fila_titulos(crudo):

    """Primera fila (de las 15 primeras) con 'fecha' y 'estado'."""

    for i in range(min(15, len(crudo))):

        valores = [norm(v) for v in crudo.iloc[i].tolist()]

        if any(v.startswith("fecha") for v in valores) and any(v.endswith("estado") for v in valores):
            return i

    return None


def _estado(v):

    t = norm(v)

    if not t:
        return SIN_ESTADO

    if t.startswith("abie") or t.startswith("abier") or t == "abiero":
        return ABIERTO

    if t.startswith("cerr"):
        return CERRADO

    if t == "ris" or "in situ" in t or "terreno" in t:
        return RIS

    return texto(v).capitalize()


def _empresa(v):

    t = norm(v)

    if "apoyo" in t:
        return "Apoyo Activo"

    if "ingecaf" in t:
        return "INGECAF"

    return texto(v)


def _subcontrato_en_empresa(v):

    """'APOYO ACTIVO- EME III' -> 'EME III'"""

    m = re.match(r"\s*apoyo activo\s*-\s*(.+)$", str(v or ""), flags=re.I)

    return m.group(1).strip() if m else None


def _lugar(v):

    t = texto(v)

    if t is None:
        return None

    t = t.title() if t.isupper() or t.islower() else t

    return (t.replace("Araucania", "Araucanía")
             .replace("Villarica", "Villarrica")
             .replace("Pucon", "Pucón"))


def _contrato(actividad):

    return CONTRATOS.get(norm(actividad), texto(actividad) or "Sin informar")


# ======================================================
# LECTURA DE UNA PLANILLA
# ======================================================

def leer_planilla(ruta):

    """Devuelve (DataFrame estándar, aviso o None)."""

    hojas = pd.read_excel(ruta, sheet_name=None, header=None, dtype=object)

    partes = []

    for nombre_hoja, crudo in hojas.items():

        crudo = crudo.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)

        if crudo.empty:
            continue

        i = _fila_titulos(crudo)

        if i is None:
            continue

        titulos = crudo.iloc[i].tolist()

        datos = crudo.iloc[i + 1:].reset_index(drop=True)

        df = pd.DataFrame()

        for j, titulo in enumerate(titulos):

            campo = _campo(titulo)

            if campo and campo not in df:
                df[campo] = datos.iloc[:, j]

        df["Hoja"] = nombre_hoja

        partes.append(df)

    if not partes:
        return pd.DataFrame(), f"{ruta.name}: no se encontró la fila de títulos (Fecha / Estado)"

    df = pd.concat(partes, ignore_index=True)

    df["Archivo"] = ruta.name

    return df, None


# ======================================================
# UNIÓN Y LIMPIEZA
# ======================================================

COLUMNAS_SALIDA = [
    "N° registro", "Fecha", "Mes", "Empresa", "Subcontrato", "Región", "Zona",
    "Delegación", "Contrato", "Tipo de actividad", "Actividad", "Responsable",
    "Estado", "Situación", "Días abierto", "Tramo antigüedad", "Fecha cierre",
    "Días al cierre", "Tipo de hallazgo", "Hallazgo", "Potencialidad",
    "Paralización", "Clasificación", "Control asociado", "Observaciones",
    "Medidas correctivas", "Estado original", "Inspector", "Archivo",
]


def _tramo(dias):

    if pd.isna(dias):
        return None

    for desde, hasta, nombre in TRAMOS:

        if desde <= dias <= hasta:
            return nombre

    return TRAMOS[-1][2]


def cargar_hallazgos(rutas, hoy=None):

    """
    rutas: lista de Path a planillas .xlsx
    Devuelve dict con 'datos' (un hallazgo por fila), 'avisos' y 'calidad'.
    """

    hoy = pd.Timestamp(hoy or pd.Timestamp.today()).normalize()

    partes, avisos = [], []

    for ruta in rutas:

        try:
            df, aviso = leer_planilla(ruta)
        except Exception as e:  # archivo dañado o abierto
            df, aviso = pd.DataFrame(), f"{ruta.name}: no se pudo leer ({e})"

        if aviso:
            avisos.append(aviso)

        if not df.empty:
            partes.append(df)

    if not partes:
        return {"datos": pd.DataFrame(columns=COLUMNAS_SALIDA), "avisos": avisos, "calidad": {}}

    df = pd.concat(partes, ignore_index=True)

    for c in COLUMNAS_SALIDA:
        if c not in df:
            df[c] = None

    filas_leidas = len(df)

    # ---- textos
    for c in ["Subcontrato", "Región", "Tipo de actividad", "Actividad", "Responsable",
              "Tipo de hallazgo", "Hallazgo", "Potencialidad", "Paralización",
              "Clasificación", "Control asociado", "Observaciones",
              "Medidas correctivas", "Inspector", "IDS con hallazgos", "Estado original"]:
        df[c] = df[c].map(texto)

    # ---- sin hallazgo: inspecciones sin observaciones ("IDS con hallazgos" = No)
    sin_hallazgo = (
        df["Hallazgo"].isna() & (
            df["IDS con hallazgos"].map(norm).eq("no") | df["Estado original"].isna()
        )
    ) | df["Hallazgo"].map(norm).str.startswith("sin hallazgo")

    descartadas_sin_hallazgo = int(sin_hallazgo.sum())

    df = df[~sin_hallazgo].copy()

    # ---- números, fechas y nombres
    df["N° registro"] = pd.to_numeric(df["N° registro"], errors="coerce").astype("Int64")

    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce", dayfirst=True)

    df["Fecha cierre"] = pd.to_datetime(df["Fecha cierre"], errors="coerce", dayfirst=True)

    sub_en_empresa = df["Empresa"].map(_subcontrato_en_empresa)

    df["Subcontrato"] = df["Subcontrato"].fillna(sub_en_empresa)

    df["Empresa"] = df["Empresa"].map(_empresa)

    for c in ["Zona", "Delegación", "Región"]:
        df[c] = df[c].map(_lugar)

    df["Tipo de hallazgo"] = df["Tipo de hallazgo"].map(lambda v: v.strip() if isinstance(v, str) else "Sin informar")

    df["Potencialidad"] = df["Potencialidad"].map(lambda v: v.strip().capitalize() if isinstance(v, str) else v)

    df["Contrato"] = df["Tipo de actividad"].map(_contrato)

    df["Estado"] = df["Estado original"].map(_estado)

    df["Situación"] = df["Estado"].map(
        lambda e: "Abierto" if e == ABIERTO else ("Cerrado" if e in (CERRADO, RIS) else "Sin estado")
    )

    sin_fecha = int(df["Fecha"].isna().sum())

    # ---- duplicados exactos (misma fila copiada dos veces)
    antes = len(df)

    df = df.drop_duplicates(subset=["N° registro", "Fecha", "Hallazgo", "Estado", "Archivo"])

    duplicados = antes - len(df)

    # ---- días
    fecha_dia = df["Fecha"].dt.normalize()

    df["Días abierto"] = (hoy - fecha_dia).dt.days.where(df["Estado"] == ABIERTO).astype("Int64")

    df["Días al cierre"] = (df["Fecha cierre"].dt.normalize() - fecha_dia).dt.days.astype("Int64")

    df["Tramo antigüedad"] = df["Días abierto"].map(_tramo)

    df["Mes"] = df["Fecha"].dt.to_period("M").dt.to_timestamp()

    df = df[COLUMNAS_SALIDA].sort_values(["Fecha", "N° registro"], ascending=[False, False]).reset_index(drop=True)

    calidad = {
        "Planillas leídas": len(rutas),
        "Filas leídas": filas_leidas,
        "Inspecciones sin hallazgos (descartadas)": descartadas_sin_hallazgo,
        "Filas duplicadas (descartadas)": duplicados,
        "Hallazgos sin fecha": sin_fecha,
        "Hallazgos en el reporte": len(df),
    }

    return {"datos": df, "avisos": avisos, "calidad": calidad}
