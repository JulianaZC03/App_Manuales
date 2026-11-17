"""Declaración de cumplimiento de construcción de instalaciones eléctricas (RETIE 2024) - constructor."""
import io

import streamlit as st
from docxtpl import DocxTemplate

import nucleo

NOMBRE = "Declaración de cumplimiento - Constructor"
PLANTILLA = nucleo.PLANTILLAS / "plantilla_declaracion_constructor.docx"
MARCA_AGUA = False

REQUIERE = {
    "alcance", "direccion", "ciudad", "departamento", "propietario", "firma_documento",
    "constructor.nombre", "constructor.cedula", "constructor.ciudad_exp", "constructor.profesion",
    "constructor.mp", "constructor.consejo", "constructor.direccion", "constructor.telefono",
    "constructor.correo", "constructor.firma",
    "disenador.nombre", "disenador.cedula", "disenador.ciudad_exp", "disenador.profesion",
    "disenador.mp",
}
FIRMAS_OBLIGATORIAS = {"constructor", "disenador"}


def formulario(p):
    """Campos propios de este documento. Devuelve (datos, errores)."""
    errores = []
    c1, c2 = st.columns(2)
    with c1:
        numero = st.text_input("Número de declaración de constructor:", key=f"{p}_numero")
    with c2:
        inicio = st.date_input("Fecha de inicio de obra:", format="DD/MM/YYYY", key=f"{p}_inicio")
    if not numero.strip():
        errores.append("Declaración constructor: falta el número de declaración.")
    return {"numero": numero.strip(), "fecha_inicio": inicio}, errores


def archivo(comunes):
    return nucleo.nombre_seguro(f"Declaracion_constructor_{comunes['nombre_proyecto']}")


def _fecha_larga(f):
    return f"{f.day} de {nucleo.MESES[f.month - 1]} de {f.year}"


def generar(comunes, datos):
    """Llena la plantilla y devuelve los bytes del .docx."""
    c, d, f = comunes["constructor"], comunes["disenador"], comunes["firma_fecha"]
    m = nucleo.mayus
    doc = DocxTemplate(str(PLANTILLA))
    doc.render({
        "numero_declaracion": datos["numero"],
        "constructor_nombre": m(c["nombre"]),
        "constructor_cedula": c["cedula"],
        "constructor_ciudad_exp": m(c["ciudad_exp"]),
        "constructor_profesion": m(c["profesion"]),
        "constructor_mp": m(c["mp"]),
        "constructor_consejo": m(c["consejo"]),
        "alcance": m(comunes["alcance"]),
        "direccion": m(comunes["direccion"]),
        "ciudad": m(comunes["ciudad"]),
        "departamento": m(comunes["departamento"]),
        "fecha_inicio_obra": _fecha_larga(datos["fecha_inicio"]),
        "propietario_nombre": m(comunes["propietario_nombre"]),
        "propietario_tipo_id": comunes["propietario_tipo_id"] + ":",
        "propietario_id": comunes["propietario_id"],
        "disenador_nombre": m(d["nombre"]),
        "disenador_profesion": m(d["profesion"]),
        "disenador_cedula": d["cedula"],
        "disenador_ciudad_exp": m(d["ciudad_exp"]).rstrip("."),
        "disenador_mp": m(d["mp"]),
        "firma_dia": f.day,
        "firma_mes": nucleo.MESES[f.month - 1].capitalize(),
        "firma_anio": f.year,
        "firma_ciudad": comunes["firma_ciudad"],
        "constructor_direccion": c["direccion"],
        "constructor_telefono": c["telefono"],
        "constructor_correo": c["correo"],
        "firma_constructor": nucleo.firma_inline(doc, c["firma"]),
    }, autoescape=True)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
