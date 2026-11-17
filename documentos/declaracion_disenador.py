"""Declaración de cumplimiento de diseño de instalaciones eléctricas (RETIE 2024) - diseñador."""
import io

import streamlit as st
from docxtpl import DocxTemplate

import nucleo

NOMBRE = "Declaración de cumplimiento - Diseñador"
PLANTILLA = nucleo.PLANTILLAS / "plantilla_declaracion_disenador.docx"
MARCA_AGUA = False

REQUIERE = {
    "alcance", "direccion", "ciudad", "departamento", "propietario", "firma_documento",
    "disenador.nombre", "disenador.cedula", "disenador.ciudad_exp", "disenador.profesion",
    "disenador.mp", "disenador.direccion", "disenador.telefono", "disenador.correo",
    "disenador.firma",
}
FIRMAS_OBLIGATORIAS = {"disenador"}
DESVIACION_BASE = "LOS 15 MM (ver memorias)"      # la que trae el documento original


def formulario(p):
    """Campos propios de este documento. Devuelve (datos, errores)."""
    numero = st.text_input("Número de declaración de diseñador:", key=f"{p}_numero")
    extra = st.text_area("Otras desviaciones del diseño (opcional, una por línea)", key=f"{p}_desv",
                         placeholder="Ej.: LA ALTURA DE MONTAJE DE LOS TOMACORRIENTES (ver memorias)")
    st.caption(f"Siempre se incluye «SE PRESENTA DESVIACIÓN REFERENTE A {DESVIACION_BASE}». "
               "Lo que escribas aquí se agrega como líneas adicionales.")
    errores = [] if numero.strip() else ["Declaración diseñador: falta el número de declaración."]
    return {"numero": numero.strip(),
            "desviaciones_extra": [l.strip() for l in extra.splitlines() if l.strip()]}, errores


def archivo(comunes):
    return nucleo.nombre_seguro(f"Declaracion_disenador_{comunes['nombre_proyecto']}")


def generar(comunes, datos):
    """Llena la plantilla y devuelve los bytes del .docx."""
    d, f = comunes["disenador"], comunes["firma_fecha"]
    m = nucleo.mayus
    doc = DocxTemplate(str(PLANTILLA))
    doc.render({
        "numero_declaracion": datos["numero"],
        "desviaciones": [DESVIACION_BASE] + [m(x) for x in datos.get("desviaciones_extra", [])],
        "disenador_nombre": m(d["nombre"]),
        "disenador_cedula": d["cedula"],
        "disenador_ciudad_exp": m(d["ciudad_exp"]),
        "disenador_profesion": m(d["profesion"]),
        "disenador_mp": m(d["mp"]),
        "alcance": m(comunes["alcance"]).rstrip("."),     # la plantilla ya pone el punto final
        "direccion": m(comunes["direccion"]),
        "ciudad": m(comunes["ciudad"]),
        "departamento": m(comunes["departamento"]),
        "propietario_nombre": m(comunes["propietario_nombre"]),
        "propietario_tipo_id": comunes["propietario_tipo_id"] + ":",
        "propietario_id": comunes["propietario_id"],
        "firma_dia": f.day,
        "firma_mes": nucleo.MESES[f.month - 1].capitalize(),
        "firma_anio": f.year,
        "firma_ciudad": comunes["firma_ciudad"],
        "disenador_direccion": d["direccion"],
        "disenador_telefono": d["telefono"],
        "disenador_correo": d["correo"],
        "firma_disenador": nucleo.firma_inline(doc, d["firma"]),
    }, autoescape=True)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
