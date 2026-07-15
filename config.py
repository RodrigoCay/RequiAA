# -*- coding: utf-8 -*-

"""
=========================================================
Dashboard de Compras
Archivo de configuración
=========================================================
"""

from pathlib import Path

# =========================================================
# ARCHIVO EXCEL
# =========================================================

EXCEL_FILE = Path(
    r"C:\Users\rcayuleo\OneDrive - Transformadores Tusan SA\Archivos de Claudia Muñoz Farias EXT - Proceso de compras\Prueba.xlsx"
)

SHEET_NAME = "Tabla Control"

# =========================================================
# COLORES
# =========================================================

BACKGROUND = "#000000"

CARD = "#161616"

SIDEBAR = "#111111"

TEXT = "#FFFFFF"

ACCENT = "#00B4D8"

SUCCESS = "#00C853"

WARNING = "#FFC107"

DANGER = "#F44336"

INFO = "#2196F3"

# =========================================================
# ESTADOS
# =========================================================

COLOR_ESTADOS = {

    "Proceso OC completo": SUCCESS,

    "Falta entrega": WARNING,

    "Falta OC": DANGER,

    "Falta cotización": INFO

}

# =========================================================
# COLUMNAS DEL EXCEL
# =========================================================

COLUMNAS = {

    "ordenes": "Ordenes",

    "requisiciones": "Requisiciones",

    "cotizaciones": "Cotizaciones",

    "oc": "OC",

    "nro_orden": "Nro de Orden",

    "fecha_req": "Fecha Requisición",

    "fecha_cot": "Fecha Cotización",

    "fecha_oc": "Fecha OC",

    "tipo_pago": "Tipo de Pago",

    "monto": "Monto Orden",

    "empresa_sol": "Empresa solicitante",

    "empresa_compra": "Empresa compra",

    "solicitante": "Nombre Solicitante",

    "contrato": "Contrato Solicitante",

    "zona": "Zona requisición",

    "tipo_compra": "Tipo de compra",

    "estado": "Estado",

    "periodo": "Periodo Requisición",

    "fecha_pago": "Fecha de pago",

    "fecha_entrega": "Fecha de entrega",

    "tiempo": "Tiempo Ciclo",

    "folio": "Folio Factura",

    "semana": "semana"

}

# =========================================================
# FORMATO DE FECHAS
# =========================================================

DATE_FORMAT = "%d-%m-%Y"

# =========================================================
# ALTURA TABLAS
# =========================================================

TABLE_HEIGHT = 600

# =========================================================
# CACHE
# =========================================================

CACHE_SECONDS = 60