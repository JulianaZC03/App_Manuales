"""Procedimiento de verificación previa a la energización de tableros eléctricos."""
import io
from datetime import datetime

import streamlit as st
from docxtpl import DocxTemplate

import nucleo

NOMBRE = "Verificación de tableros"
PLANTILLA = nucleo.PLANTILLAS / "plantilla_verificacion_tableros.docx"
MARCA_AGUA = True      # cuadro de texto azul sobre cada firma conocida

REQUIERE = {"ubicacion",
            "disenador.cedula", "disenador.mp", "disenador.firma",
            "constructor.cedula", "constructor.mp", "constructor.firma"}
FIRMAS_OBLIGATORIAS = set()         # como en el manual: las firmas son opcionales

ALCANCE_BASE = ["Tableros de distribución", "Celdas de baja tensión"]
# El cuadro de texto de la marca de agua mide 68 x 48 mm: la letra se ajusta al largo del texto
WM_ANCHO_MM, WM_ALTO_MM, WM_BASE_PT = 63, 34, 18
RAYAS = {"firma": "___________________________", "cedula": "__________________________",
         "mp": "______________________________"}


def _lineas(texto):
    return [l.strip() for l in texto.splitlines() if l.strip()]


def formulario(p):
    """Campos propios de este documento. Devuelve (datos, errores)."""
    errores = []
    st.markdown("**Alcance (el documento dice «Aplica a:»)**")
    alcance = st.multiselect("Aplica a", ALCANCE_BASE, default=ALCANCE_BASE, key=f"{p}_alcance")
    otros = _lineas(st.text_area("Otros (uno por línea)", key=f"{p}_alcance_otros",
                                 placeholder="Tableros regulados (UPS)\nTableros de sistemas contraincendios\nTableros de control"))
    st.markdown("**Tableros de la instalación**")
    tableros = _lineas(st.text_area(
        "Nombre de cada tablero (uno por línea)", value="Tablero normal\nTablero regulado",
        key=f"{p}_tableros",
        help="Se genera una tabla de verificación por cada tablero, con ese nombre como título."))
    st.caption(f"Se generarán {len(tableros)} tabla(s) de verificación.")
    if not (alcance or otros):
        errores.append("Verificación de tableros: selecciona o escribe al menos un elemento del alcance.")
    if not tableros:
        errores.append("Verificación de tableros: escribe el nombre de al menos un tablero.")
    return {"alcance": alcance + otros, "tableros": tableros}, errores


def archivo(comunes):
    return nucleo.nombre_seguro(f"Verificacion_tableros_{comunes['nombre_proyecto']}")


def generar(comunes, datos):
    """Llena la plantilla y devuelve los bytes del .docx."""
    m = nucleo.mayus
    nombre, ubicacion = m(comunes["nombre_proyecto"]), m(comunes["ubicacion"])
    doc = DocxTemplate(str(PLANTILLA))
    ctx = {
        "proyecto_header": f"{nombre} {ubicacion}",
        "nombre_proyecto": nombre,
        "ubicacion": ubicacion,
        "anio": datetime.now().year,
        "alcance": datos["alcance"],
        "tableros": [m(t) for t in datos["tableros"]],
    }
    texto_wm = f"VÁLIDO ÚNICAMENTE PROYECTO {nombre} {ubicacion}"
    for rol, suf in (("disenador", "d"), ("constructor", "c")):
        persona = comunes.get(rol)
        if persona:                                       # firma conocida: datos + marca de agua
            ctx[f"firma_{rol}"] = nucleo.firma_inline(doc, persona["firma"])
            ctx[f"cedula_{rol}"] = persona["cedula"]
            ctx[f"mp_{rol}"] = m(persona["mp"])
            ctx[f"wm_{suf}"] = texto_wm
            ctx[f"wm_sz_{suf}"] = nucleo.tamano_marca_agua(texto_wm, WM_ANCHO_MM, WM_ALTO_MM, WM_BASE_PT)
        else:                                             # no se conoce: quedan las rayas y sin marca
            ctx[f"firma_{rol}"] = RAYAS["firma"]
            ctx[f"cedula_{rol}"] = RAYAS["cedula"]
            ctx[f"mp_{rol}"] = RAYAS["mp"]
            ctx[f"wm_{suf}"] = ""
            ctx[f"wm_sz_{suf}"] = WM_BASE_PT * 2
    doc.render(ctx, autoescape=True)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
