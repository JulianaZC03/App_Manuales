"""Registro de documentos. Para agregar uno nuevo: crear su módulo y añadirlo aquí.

Cada módulo expone: NOMBRE, REQUIERE, FIRMAS_OBLIGATORIAS, formulario(prefijo),
archivo(comunes) y generar(comunes, datos) -> bytes del .docx.
"""
from . import declaracion_constructor, declaracion_disenador, manual_oym

REGISTRO = {
    "manual_oym": manual_oym,
    "decl_constructor": declaracion_constructor,
    "decl_disenador": declaracion_disenador,
}
