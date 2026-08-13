# -*- coding: utf-8 -*-

"""
=========================================================
Dashboard de Compras
Archivo de configuración
=========================================================
"""

from pathlib import Path

# =========================================================
# ARCHIVO EXC
# =========================================================

BASE_DIR = Path(__file__).parent
 
EXCEL_FILE = BASE_DIR / "Prueba.xlsx"
 
SHEET_NAME = "Tabla Control"

# =========================================================
# ARCHIVO DE DETALLE DE OC (ítems por solicitud)
# =========================================================

DETALLE_OC_FILE = BASE_DIR / "Detalle solicitudes OC.xlsx"

DETALLE_OC_SHEET = "Detalle"

REQUISICIONES_SHEET = "Requisiciones"

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
# COLUMNAS DEL DETALLE DE OC ("Detalle solicitudes OC.xlsx")
# =========================================================

COLUMNAS_DETALLE_OC = {

    "llave": "LLave",

    "requisicion": "N° de requisición",

    "descripcion": "Descripción",

    "cantidad": "Cantidad",

    "valor_unitario": "Valor unitario",

    "valor_neto": "Valor Neto",

    "empresa_sol": "Empresa solicitante",

    "solicitante": "Nombre Solicitante",

    "contrato": "Contrato Solicitante",

    "zona": "Zona requisición",

    "estado_gestion": "Estado Gestión",

    "comentario": "Comentario"

}

# =========================================================
# COLUMNAS DE LA HOJA "Requisiciones"
# =========================================================
# (mismo archivo "Detalle solicitudes OC.xlsx", hoja aparte)
# No se incluyen las columnas de referencia de la planilla
# (listas desplegables tipo "Column8", "BackOffice (BBOO)",
# "04 Coquimbo", "Litros", etc.) porque no son datos de la
# requisición, sino listas auxiliares de la hoja de origen.

COLUMNAS_REQUISICIONES = {

    "llave": "LLave",

    "producto": "Producto",

    "descripcion": "Descripción producto",

    "cantidad": "Cantidad",

    "unidad": "Unidad medida",

    "especificaciones": "Especificaciones técnicas",

    "link": "Link (Imagen o Especificaciones técnicas)"

}

# =========================================================
# COLUMNAS DE LA HOJA "Cotizaciones"
# =========================================================
# (mismo archivo "Detalle solicitudes OC.xlsx", hoja aparte)

COTIZACIONES_SHEET = "Cotizaciones"

COLUMNAS_COTIZACIONES = {

    "llave": "LLave",

    "articulo": "articulo",

    "un": "UN",

    "proveedor": "proveedor",

    "rut_proveedor": "Rut Proveedor",

    "tipo_pago": "tipo pago",

    "valor_un_neto": "valor un neto",

    "subtotal": "subtotal",

    "iva": "iva",

    "totales": "totales",

    "estado": "Estado"

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
