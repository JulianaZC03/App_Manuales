"""Carta de responsabilidad de diseño: ítems RETIE 10.1.1 que aplican y que no aplican."""
import io

import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate

import nucleo

NOMBRE = "Carta de diseñador"
PLANTILLA = nucleo.PLANTILLAS / "plantilla_carta_disenador.docx"
MARCA_AGUA = True      # cuadro de texto azul que ya trae la plantilla

REQUIERE = {"ubicacion", "disenador.nombre", "disenador.mp", "disenador.firma"}
FIRMAS_OBLIGATORIAS = {"disenador"}

# (letra, título, aplica por defecto). Los valores por defecto corresponden a un local comercial
# dentro de un edificio.
ITEMS = [
    ('a', 'Análisis de riesgos de origen eléctrico y medidas para mitigarlos.', True),
    ('b', 'Análisis de riesgos por descargas eléctricas atmosféricas (rayos) y medidas de protección.', False),
    ('c', 'Análisis y cálculo de cargas iniciales y futuras, incluyendo factor de potencia y armónicos.', True),
    ('d', 'Coordinación de aislamiento eléctrico.', False),
    ('e', 'Análisis y cálculos de cortocircuito, arco eléctrico y falla a tierra.', False),
    ('f', 'Análisis del nivel tensión requerido.', True),
    ('g', 'Cálculos de campos electromagnéticos.', False),
    ('h', 'Cálculo de transformadores incluyendo efectos de los armónicos y factor de potencia en la carga.', False),
    ('i', 'Sistema de puesta a tierra.', False),
    ('j', 'Cálculo económico de conductores, teniendo en cuenta todos los factores de pérdidas, las cargas resultantes y los costos de la energía.', True),
    ('k', 'Especificación de los conductores, teniendo en cuenta el tiempo de disparo de los interruptores, la corriente de cortocircuito de la red y la capacidad de corriente del conductor, de acuerdo con la norma IEC 60909 u otra equivalente.', True),
    ('l', 'Cálculo mecánico de estructuras y de elementos de sujeción y soporte de redes de transmisión, de distribución, subestaciones y centrales de generación.', False),
    ('m', 'Cálculo y coordinación de protecciones contra sobrecorrientes. En baja tensión se permite la coordinación con las características de limitación de corriente de los dispositivos según IEC 60947-2 Anexo A.', True),
    ('n', 'Cálculos de canalizaciones (tubos, ductos, canales y electroductos), bandejas portacables y volumen de encerramientos (cajas, conduletas, armarios, etc.)', True),
    ('o', 'Cálculo de pérdidas de energía, teniendo en cuenta los efectos de armónicos y factor de potencia.', True),
    ('p', 'Cálculos de regulación de tensión.', True),
    ('q', 'Áreas clasificadas como peligrosas.', False),
    ('r', 'Diagramas unifilares.', True),
    ('s', 'Planos eléctricos para construcción.', True),
    ('t', 'Especificaciones de construcción complementarias a los planos, incluyendo las de tipo técnico de equipos y materiales y sus condiciones particulares.', False),
    ('u', 'Distancias de seguridad o servidumbre requeridas.', True),
    ('v', 'Justificación de desviaciones técnicas cuando sea estrictamente necesarias, siempre y cuando no comprometa la seguridad de las personas o de la instalación.', True),
    ('w', 'Los demás estudios que el tipo de instalación requiera para su correcta y segura operación, tales como condiciones sísmicas, acústicas, mecánicas o térmicas.', False),
    ('x', 'Selección, cálculo y especificación de equipos de generación de energía convencionales y no convencionales.', False),
]


def formulario(p):
    """Tabla de ítems con casillas "Aplica". Devuelve (datos, errores)."""
    st.markdown("Marca los ítems que **aplican**. Los valores iniciales son los de un local comercial dentro de un edificio.")
    base = pd.DataFrame({"Ítem": [i[0] for i in ITEMS],
                         "Título": [i[1] for i in ITEMS],
                         "Aplica": [i[2] for i in ITEMS]})
    editado = st.data_editor(base, key=f"{p}_items", hide_index=True, use_container_width=True,
                             disabled=["Ítem", "Título"],
                             column_config={"Ítem": st.column_config.TextColumn(width="small"),
                                            "Título": st.column_config.TextColumn(width="large"),
                                            "Aplica": st.column_config.CheckboxColumn(width="small")})
    aplica = {letra: bool(v) for letra, v in zip(editado["Ítem"], editado["Aplica"])}
    return {"aplica": aplica}, []


def archivo(comunes):
    return nucleo.nombre_seguro(f"Carta_disenador_{comunes['nombre_proyecto']}")


def generar(comunes, datos):
    """Llena la plantilla y devuelve los bytes del .docx."""
    d = comunes["disenador"]
    m = nucleo.mayus
    doc = DocxTemplate(str(PLANTILLA))
    ctx = {
        "disenador_nombre": m(d["nombre"]),
        "disenador_mp": m(d["mp"]),
        "proyecto_completo": m(f"{comunes['nombre_proyecto']} {comunes['ubicacion']}"),
        "firma_disenador": nucleo.firma_inline(doc, d["firma"]),
    }
    # El cuadro de la marca de agua mide 44 x 32 mm; se deja el alto que ocupaban ~4 líneas del original: la letra se ajusta al texto
    ctx["wm_sz"] = nucleo.tamano_marca_agua("VÁLIDO ÚNICAMENTE PROYECTO " + ctx["proyecto_completo"], 39, 21)
    for letra, _, _ in ITEMS:
        ctx[f"ap_{letra}"] = "SI" if datos["aplica"].get(letra) else "NO"
    doc.render(ctx, autoescape=True)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
