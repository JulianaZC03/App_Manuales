"""Manual de Operación y Mantenimiento (el documento original de la app)."""
import io
from datetime import datetime

import streamlit as st
from docxtpl import DocxTemplate

import nucleo
from marca_agua import agregar_marca_agua, crear_marca_agua

NOMBRE = "Manual de Operación y Mantenimiento"
PLANTILLA = nucleo.PLANTILLAS / "plantilla_manual.docx"
MARCA_AGUA = True

# Datos comunes que necesita este documento (ver app.py para la lista de claves)
REQUIERE = {"ubicacion",
            "disenador.cedula", "disenador.mp", "disenador.firma",
            "constructor.cedula", "constructor.mp", "constructor.firma"}
# Aquí las firmas son opcionales ("¿Se conoce el diseñador?"). Si otro documento las exige, mandan ellos.
FIRMAS_OBLIGATORIAS = set()

# Tipos de instalación y su género gramatical (para "un/una ... alimentado/a")
TIPOS = {
    "Local comercial": "m",
    "Consultorio": "m",
    "Oficina": "f",
    "Bodega": "f",
    "Restaurante": "m",
    "Vivienda": "f",
}

TDG = "Tablero general de distribución (TDG)"
UPS = "UPS y circuitos regulados"
SOLAR = "Paneles solares e inversor"
SOLAR_BAT = "Paneles solares, inversor y banco de baterías"

COMPONENTES = [
    TDG,
    UPS,
    "Circuitos de iluminación y tomacorrientes",
    "Cargas especiales (equipos electrónicos)",
    SOLAR,
    SOLAR_BAT,
]

# Condiciones técnicas especiales (punto 2g). "Existencia de UPS" no se elige: sale sola si hay UPS.
CONDICIONES = [
    "Cargas sensibles (POS, DVR, alarmas)",
    "Cargas continuas (refrigeración)",
    "Sistema depende del SPT del edificio",
    "Uso intensivo de iluminación LED",
]

NIVELES_TENSION = ["alta", "media", "baja"]
SISTEMAS_FIJOS = ["Breakers (protección principal)", "Tablero general de distribución"]


def formulario(p):
    """Dibuja los campos propios del manual. Devuelve (datos, errores)."""
    errores = []
    st.markdown("**Características de la instalación**")
    c1, c2 = st.columns(2)
    with c1:
        tipo_sel = st.selectbox("Tipo de instalación", list(TIPOS) + ["Otro…"], key=f"{p}_tipo")
    with c2:
        tension = st.text_input("Tensión nominal", value="208/120 V", key=f"{p}_tension")

    tipo, genero = tipo_sel, TIPOS.get(tipo_sel, "m")
    if tipo_sel == "Otro…":
        c3, c4 = st.columns(2)
        with c3:
            tipo = st.text_input("¿Cuál tipo?", placeholder="Clínica veterinaria", key=f"{p}_tipo_otro")
        with c4:
            opcion = st.radio("Género", ["un (masculino)", "una (femenino)"],
                              horizontal=True, key=f"{p}_genero")
            genero = "m" if opcion.startswith("un ") else "f"

    nivel = st.selectbox("Nivel de tensión", NIVELES_TENSION,
                         index=NIVELES_TENSION.index("baja"), key=f"{p}_nivel")

    st.markdown("**Componentes de la instalación**")
    componentes = st.multiselect("Componentes", COMPONENTES, default=COMPONENTES[:4], key=f"{p}_comp")
    otros_txt = st.text_area("Otros componentes (uno por línea)",
                             placeholder="Planta de emergencia\nAscensor", key=f"{p}_otros")
    otros = [l.strip() for l in otros_txt.splitlines() if l.strip()]

    st.markdown("**Condiciones técnicas especiales**")
    condiciones = st.multiselect("Condiciones", CONDICIONES, default=CONDICIONES, key=f"{p}_cond")
    st.caption("«Existencia de UPS» se agrega sola si marcaste «UPS y circuitos regulados».")
    cond_txt = st.text_area("Otros, ¿cuáles? (uno por línea)", placeholder="Equipos de rayos X",
                            key=f"{p}_cond_otras")
    cond_otras = [l.strip() for l in cond_txt.splitlines() if l.strip()]

    if not tipo.strip():
        errores.append("Manual: falta el tipo de instalación.")
    if not tension.strip():
        errores.append("Manual: falta la tensión nominal.")
    if not (componentes or otros):
        errores.append("Manual: selecciona al menos un componente.")
    if SOLAR in componentes and SOLAR_BAT in componentes:
        errores.append("Manual: elige solo una opción solar (con inversor, o con inversor y baterías).")

    return {"tipo": tipo.strip(), "genero": genero, "tension": tension.strip(), "nivel": nivel,
            "componentes": componentes, "otros": otros, "condiciones": condiciones,
            "cond_otras": cond_otras}, errores


def archivo(comunes):
    return nucleo.nombre_seguro(f"Manual_OyM_{comunes['nombre_proyecto']}")


def generar(comunes, datos):
    """Llena la plantilla y devuelve los bytes del .docx."""
    nombre, ubicacion = comunes["nombre_proyecto"], comunes["ubicacion"]
    disenador, constructor = comunes.get("disenador"), comunes.get("constructor")
    componentes, otros = list(datos["componentes"]), datos["otros"]

    articulo, alimentado = ("un", "alimentado") if datos["genero"] == "m" else ("una", "alimentada")
    tiene_ups = UPS in componentes
    tiene_solar = SOLAR in componentes or SOLAR_BAT in componentes
    tiene_baterias = SOLAR_BAT in componentes

    doc = DocxTemplate(str(PLANTILLA))
    ctx = {
        "tipo": datos["tipo"],
        "nombre_proyecto": nombre,
        "ubicacion": ubicacion,
        "anio": datetime.now().year,
        "tipo_descripcion": f"{articulo} {datos['tipo'].lower()} {alimentado}",
        "tension_nominal": datos["tension"],
        "nivel_tension": datos["nivel"],
        "componentes": componentes + otros,
        "tiene_ups": tiene_ups,
        "tiene_solar": tiene_solar,
        "equipos_solares": ("paneles, inversor y banco de baterías" if tiene_baterias
                            else "paneles e inversor"),
        "zonas_dc": ("tableros, paneles, inversores y baterías" if tiene_baterias
                     else "tableros, paneles e inversores"),
        "condiciones_tecnicas": (["Existencia de UPS"] if tiene_ups else [])
                                + list(datos["condiciones"]) + datos["cond_otras"],
        "sistemas_disponibles": SISTEMAS_FIJOS
                                + [c for c in componentes if c != TDG] + otros,
    }
    for clave, persona in (("disenador", disenador), ("constructor", constructor)):
        ctx[f"{clave}_conocido"] = persona is not None
        ctx[f"firma_{clave}"] = nucleo.firma_inline(doc, persona["firma"]) if persona else ""
        ctx[f"cedula_{clave}"] = persona["cedula"] if persona else ""
        ctx[f"mp_{clave}"] = persona["mp"] if persona else ""

    doc.render(ctx, autoescape=True)
    buf = io.BytesIO()
    doc.save(buf)
    resultado = buf.getvalue()

    # Marca de agua azul: una por firma conocida (celda 0 = diseñador, celda 1 = constructor)
    if MARCA_AGUA:
        celdas = [i for i, p in enumerate((disenador, constructor)) if p is not None]
        if celdas:
            png, ancho, alto = crear_marca_agua(nombre, ubicacion)
            resultado = agregar_marca_agua(resultado, celdas, png, ancho, alto)
    return resultado
