import io
import re

import streamlit as st

from pathlib import Path

from PIL import Image

from generar_manual import (COMPONENTES, CONDICIONES, NIVELES_TENSION, SOLAR, SOLAR_BAT,
                            TIPOS, a_png, docx_a_pdf, generar)

LOGO = Path(__file__).parent / "Imgs/logo.png"   # logo de la empresa (cuadrado, mínimo 64x64 px)

st.set_page_config(
    page_title="Generador de manuales O&M",
    page_icon=Image.open(LOGO) if LOGO.exists() else "⚡"
)

# Crear dos columnas: una estrecha para el logo y una ancha para el título
col1, col2 = st.columns([1, 4], vertical_alignment="center")

with col1:
    if LOGO.exists():
        st.image(str(LOGO), width=80)  # Ajusta el ancho (width) según necesites

with col2:
    st.title("Manual de Operación y Mantenimiento")
st.caption("Completa los datos del proyecto y descarga el documento en Word.")

# ---------- 1. Datos del proyecto ----------
st.subheader("1. Datos del proyecto")
c1, c2 = st.columns(2)
with c1:
    tipo_sel = st.selectbox("Tipo de instalación", list(TIPOS) + ["Otro…"])
with c2:
    nombre = st.text_input("Nombre del proyecto", placeholder="Ocean Blue Pets")

tipo, genero = tipo_sel, TIPOS.get(tipo_sel, "m")
if tipo_sel == "Otro…":
    c3, c4 = st.columns(2)
    with c3:
        tipo = st.text_input("¿Cuál tipo?", placeholder="Clínica veterinaria")
    with c4:
        opcion = st.radio("Género", ["un (masculino)", "una (femenino)"], horizontal=True)
        genero = "m" if opcion.startswith("un ") else "f"

ubicacion = st.text_input("Ubicación", placeholder="Plaza Central")

# ---------- 2. Características eléctricas ----------
st.subheader("2. Características eléctricas")
c5, c6 = st.columns(2)
with c5:
    tension_nominal = st.text_input("Tensión nominal", value="208/120 V")
with c6:
    nivel = st.selectbox("Nivel de tensión", NIVELES_TENSION,
                         index=NIVELES_TENSION.index("baja"))

# ---------- 3. Componentes ----------
st.subheader("3. Componentes de la instalación")
componentes = st.multiselect("Componentes", COMPONENTES, default=COMPONENTES[:4])
otros_txt = st.text_area("Otros componentes (uno por línea)",
                         placeholder="Planta de emergencia\nAscensor")
otros = [l.strip() for l in otros_txt.splitlines() if l.strip()]

# ---------- 4. Condiciones técnicas especiales ----------
st.subheader("4. Condiciones técnicas especiales")
condiciones = st.multiselect("Condiciones", CONDICIONES, default=CONDICIONES)
st.caption("«Existencia de UPS» se agrega sola si marcaste «UPS y circuitos regulados».")
cond_txt = st.text_area("Otros, ¿cuáles? (uno por línea)", placeholder="Equipos de rayos X")
cond_otras = [l.strip() for l in cond_txt.splitlines() if l.strip()]

# ---------- 5. Firmas ----------
st.subheader("5. Firmas")
st.caption("La firma solo se usa para generar este documento: la app no la guarda en archivos ni en perfiles.")
TIPOS_IMAGEN = ["png", "jpg", "jpeg", "webp", "bmp", "gif", "tif", "tiff"]


def bloque_firma(titulo, clave):
    """Devuelve (datos | None, errores)."""
    errores = []
    conocido = st.selectbox(f"¿Se conoce el {titulo}?", ["No", "Sí"], key=f"conoce_{clave}") == "Sí"
    if not conocido:
        return None, errores
    archivo = st.file_uploader(f"FIRMA del {titulo}", type=TIPOS_IMAGEN, key=f"firma_{clave}")
    cedula = st.text_input("CEDULA:", key=f"cedula_{clave}")
    mp = st.text_input("M.P:", key=f"mp_{clave}")
    if archivo is None:
        errores.append(f"Falta la imagen de la firma del {titulo}.")
    if not cedula.strip():
        errores.append(f"Falta la cédula del {titulo}.")
    if not mp.strip():
        errores.append(f"Falta la M.P. del {titulo}.")
    if errores:
        return None, errores
    try:
        a_png(archivo.getvalue())
    except Exception:
        return None, [f"No se pudo leer la imagen de la firma del {titulo}."]
    return {"firma": archivo.getvalue(), "cedula": cedula.strip(), "mp": mp.strip()}, []


cA, cB = st.columns(2)
with cA:
    st.markdown("**Diseñador**")
    disenador, err_d = bloque_firma("diseñador", "disenador")
with cB:
    st.markdown("**Constructor**")
    constructor, err_c = bloque_firma("constructor", "constructor")

# ---------- Validación ----------
errores = list(err_d) + list(err_c)
if not nombre.strip():
    errores.append("Falta el nombre del proyecto.")
if not ubicacion.strip():
    errores.append("Falta la ubicación.")
if not tipo.strip():
    errores.append("Falta el tipo de instalación.")
if not tension_nominal.strip():
    errores.append("Falta la tensión nominal.")
if not (componentes or otros):
    errores.append("Selecciona al menos un componente.")
if SOLAR in componentes and SOLAR_BAT in componentes:
    errores.append("Elige solo una opción solar: con inversor, o con inversor y baterías.")

if nombre or ubicacion:
    for e in errores:
        st.warning(e)

# ---------- Generación ----------
if st.button("Generar manual", type="primary", disabled=bool(errores)):
    TIPOS.setdefault(tipo, genero)  # para un tipo nuevo, usa el género elegido
    buf = io.BytesIO()
    generar(tipo, nombre.strip(), ubicacion.strip(), componentes, otros, nivel,
            tension_nominal.strip(), condiciones, cond_otras, disenador, constructor, buf)
    base = re.sub(r"[^\w\-]+", "_", f"Manual_OyM_{nombre.strip()}")
    docx_bytes = buf.getvalue()
    try:
        with st.spinner("Generando PDF…"):
            pdf_bytes = docx_a_pdf(docx_bytes)
    except Exception:
        pdf_bytes = None
    st.session_state["resultado"] = (docx_bytes, pdf_bytes, base)

if "resultado" in st.session_state:
    docx_bytes, pdf_bytes, base = st.session_state["resultado"]
    st.success("Documento generado.")
    d1, d2 = st.columns(2)
    with d1:
        st.download_button(
            "⬇️ Descargar Word (.docx)", docx_bytes, file_name=base + ".docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True)
    with d2:
        if pdf_bytes:
            st.download_button("⬇️ Descargar PDF", pdf_bytes, file_name=base + ".pdf",
                               mime="application/pdf", use_container_width=True)
        else:
            st.warning("No se pudo generar el PDF (¿LibreOffice instalado?). Usa el Word.")
