import io
import zipfile
from datetime import date
from pathlib import Path

import streamlit as st
from PIL import Image

import nucleo
from documentos import REGISTRO

LOGO = Path(__file__).parent / "Imgs/logo.png"   # logo de la empresa (cuadrado, mínimo 64x64 px)
TIPOS_IMAGEN = ["png", "jpg", "jpeg", "webp", "bmp", "gif", "tif", "tiff"]
PROFESIONES = {"disenador": ["Ingeniero Electricista", "Ingeniero Electrónico"],
               "constructor": ["Ingeniero Electricista", "Técnico Electricista"]}
ROLES = {"disenador": "Diseñador", "constructor": "Constructor"}

st.set_page_config(page_title="Generador de documentos",
                   page_icon=Image.open(LOGO) if LOGO.exists() else "⚡")

col1, col2 = st.columns([1, 4], vertical_alignment="center")
with col1:
    if LOGO.exists():
        st.image(str(LOGO), width=80)
with col2:
    st.title("Generador de documentos")
st.caption("Elige los documentos, completa los datos una sola vez y descárgalos en Word y PDF.")

# ---------- 1. Documentos ----------
st.subheader("1. Documentos a generar")
sel = [k for k, m in REGISTRO.items() if st.checkbox(m.NOMBRE, key=f"sel_{k}")]
if not sel:
    st.info("Marca al menos un documento para continuar.")
    st.stop()
mods = [REGISTRO[k] for k in sel]
req = set().union(*(m.REQUIERE for m in mods)) | {"nombre_proyecto"}
obligatorias = set().union(*(m.FIRMAS_OBLIGATORIAS for m in mods))

comunes, errores = {}, []


def texto(clave, etiqueta, placeholder="", area=False):
    w = st.text_area if area else st.text_input
    v = w(etiqueta, key=f"c_{clave}", placeholder=placeholder).strip()
    comunes[clave] = v
    if not v:
        errores.append(f"Falta: {etiqueta}")
    return v


# ---------- 2. Datos comunes ----------
st.subheader("2. Datos comunes")
st.markdown("**Proyecto**")
texto("nombre_proyecto", "Nombre del proyecto", "Ocean Blue Pets")
if "ubicacion" in req:
    texto("ubicacion", "Ubicación del proyecto", "Plaza Central")
if "direccion" in req:
    texto("direccion", "Dirección del proyecto", "Carrera 65 # 11-50 Plaza Central Local 3-08")
if "ciudad" in req or "departamento" in req:
    a, b = st.columns(2)
    with a:
        if "ciudad" in req:
            texto("ciudad", "Ciudad/municipio del proyecto", "Bogotá D.C.")
    with b:
        if "departamento" in req:
            comunes["departamento"] = st.selectbox("Departamento del proyecto",
                                                   nucleo.DEPARTAMENTOS,
                                                   index=nucleo.DEPARTAMENTOS.index("Cundinamarca"),
                                                   key="c_departamento")
if "alcance" in req:
    texto("alcance", "Alcance del proyecto",
          "Ejemplo: USO FINAL DE UN LOCAL OCEAN BLUE PETS, DESDE TABLERO DE DISTRIBUCIÓN, "
          "TABLERO REGULADO CON UPS, HASTA SALIDAS DE ILUMINACIÓN Y TOMAS", area=True)

if "propietario" in req:
    st.markdown("**Propietario**")
    a, b, c = st.columns([3, 1, 2])
    with a:
        texto("propietario_nombre", "Nombre/razón social del propietario")
    with b:
        comunes["propietario_tipo_id"] = st.selectbox("Tipo de ID", nucleo.TIPOS_ID, index=1,
                                                      key="c_prop_tipo")
    with c:
        texto("propietario_id", "Número de identificación del propietario")

if "firma_documento" in req:
    st.markdown("**Firma del documento**")
    a, b = st.columns(2)
    with a:
        comunes["firma_fecha"] = st.date_input("Fecha de firma de documento:", value=date.today(),
                                               format="DD/MM/YYYY", key="c_firma_fecha")
    with b:
        texto("firma_ciudad", "Ciudad donde se firma", "Bogotá D.C.")


def persona(rol):
    """Bloque del diseñador o constructor, con solo los campos que piden los documentos elegidos."""
    campos = {k.split(".", 1)[1] for k in req if k.startswith(rol + ".")}
    if not campos:
        return
    nombre_rol = ROLES[rol].lower()
    st.markdown(f"**{ROLES[rol]}**")
    if rol not in obligatorias:
        if st.selectbox(f"¿Se conoce el {nombre_rol}?", ["No", "Sí"], key=f"conoce_{rol}") == "No":
            comunes[rol] = None
            return
    d = {}

    def campo(clave, etiqueta, placeholder=""):
        v = st.text_input(etiqueta, key=f"{rol}_{clave}", placeholder=placeholder).strip()
        d[clave] = v
        if not v:
            errores.append(f"Falta: {etiqueta} ({nombre_rol})")

    if "nombre" in campos:
        campo("nombre", f"Nombre del {nombre_rol}")
    if "cedula" in campos:
        campo("cedula", "Número de documento")
    if "ciudad_exp" in campos:
        campo("ciudad_exp", "Ciudad de expedición")
    if "profesion" in campos:
        op = st.selectbox("Profesión", PROFESIONES[rol] + ["Otro, ¿cuál?"], key=f"{rol}_prof")
        if op.startswith("Otro"):
            campo("profesion", "¿Cuál profesión?")
        else:
            d["profesion"] = op
    if "mp" in campos:
        campo("mp", "Matrícula profesional", "CN205-000000")
    if "consejo" in campos:
        campo("consejo", "Consejo que expide la matrícula",
              "Consejo Profesional Nacional de Ingenierías Eléctrica, Mecánica y Profesiones Afines")
    if "direccion" in campos:
        campo("direccion", f"Dirección del {nombre_rol}")
    if "telefono" in campos:
        campo("telefono", f"Teléfono de contacto del {nombre_rol}")
    if "correo" in campos:
        campo("correo", f"Correo electrónico del {nombre_rol}")
    if "firma" in campos:
        archivo = st.file_uploader(f"FIRMA del {nombre_rol}", type=TIPOS_IMAGEN, key=f"{rol}_firma")
        d["firma"] = None
        if archivo is None:
            errores.append(f"Falta: imagen de la firma del {nombre_rol}")
        else:
            try:
                nucleo.a_png(archivo.getvalue())
                d["firma"] = archivo.getvalue()
            except Exception:
                errores.append(f"No se pudo leer la imagen de la firma del {nombre_rol}")
    comunes[rol] = d


if any(k.startswith(("disenador.", "constructor.")) for k in req):
    cA, cB = st.columns(2)
    with cA:
        persona("disenador")
    with cB:
        persona("constructor")

# ---------- 3. Datos específicos ----------
datos = {}
for n, (k, m) in enumerate(zip(sel, mods), start=3):
    st.subheader(f"{n}. {m.NOMBRE}")
    datos[k], e = m.formulario(k)
    errores += e

# ---------- Generación ----------
if errores:
    with st.expander(f"Faltan {len(errores)} dato(s) para poder generar", expanded=False):
        for e in errores:
            st.write("• " + e)

if st.button("Generar documentos", type="primary", disabled=bool(errores)):
    resultados = {}
    barra = st.progress(0.0)
    for i, (k, m) in enumerate(zip(sel, mods)):
        barra.progress(i / len(mods), text=f"Generando {m.NOMBRE}…")
        try:
            docx_bytes = m.generar(comunes, datos[k])
        except Exception as ex:
            st.error(f"No se pudo generar «{m.NOMBRE}»: {ex}")
            continue
        try:
            pdf_bytes = nucleo.docx_a_pdf(docx_bytes)
        except Exception:
            pdf_bytes = None
        resultados[k] = (m.NOMBRE, docx_bytes, pdf_bytes, m.archivo(comunes))
    barra.empty()
    st.session_state["resultados"] = resultados

resultados = st.session_state.get("resultados")
if resultados:
    st.success("Documentos generados.")
    for k, (nombre, docx_bytes, pdf_bytes, base) in resultados.items():
        st.markdown(f"**{nombre}**")
        d1, d2 = st.columns(2)
        with d1:
            st.download_button("⬇️ Word (.docx)", docx_bytes, file_name=base + ".docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               use_container_width=True, key=f"dl_docx_{k}")
        with d2:
            if pdf_bytes:
                st.download_button("⬇️ PDF", pdf_bytes, file_name=base + ".pdf",
                                   mime="application/pdf", use_container_width=True, key=f"dl_pdf_{k}")
            else:
                st.warning("No se pudo generar el PDF (¿Word o LibreOffice instalado?). Usa el Word.")
    if len(resultados) > 1:
        zbuf = io.BytesIO()
        with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
            for nombre, docx_bytes, pdf_bytes, base in resultados.values():
                z.writestr(base + ".docx", docx_bytes)
                if pdf_bytes:
                    z.writestr(base + ".pdf", pdf_bytes)
        st.download_button("📦 Descargar todo (.zip)", zbuf.getvalue(), file_name="Documentos.zip",
                           mime="application/zip", type="primary", use_container_width=True,
                           key="dl_zip")
