import io
import os
import shutil
import subprocess
import sys
import tempfile
import threading
from datetime import datetime

from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage
from PIL import Image

from marca_agua import agregar_marca_agua, crear_marca_agua

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


def a_png(datos):
    """Convierte la imagen (jpg, png, webp, bmp...) a PNG, le quita los márgenes vacíos
    y vuelve transparente el fondo blanco. Devuelve (png_bytes, (ancho, alto))."""
    img = Image.open(io.BytesIO(datos)).convert("RGBA")
    alfa = img.getchannel("A")
    if alfa.getextrema()[0] == 255:                      # sin transparencia: fondo claro -> transparente
        gris = img.convert("L")
        alfa = gris.point(lambda v: 0 if v >= 235 else min(255, (235 - v) * 6))
        img.putalpha(alfa)
    caja = alfa.point(lambda v: 255 if v > 20 else 0).getbbox()
    if caja:
        img = img.crop(caja)
    buf = io.BytesIO()
    img.save(buf, "PNG")
    return buf.getvalue(), img.size


# Caja en la que se acomoda la firma (se conserva la proporción)
FIRMA_ANCHO_MM, FIRMA_ALTO_MM = 50, 18


def _firma(doc, datos):
    png, (w, h) = a_png(datos)
    if w / h >= FIRMA_ANCHO_MM / FIRMA_ALTO_MM:         # firma muy ancha: manda el ancho
        return InlineImage(doc, io.BytesIO(png), width=Mm(FIRMA_ANCHO_MM))
    return InlineImage(doc, io.BytesIO(png), height=Mm(FIRMA_ALTO_MM))


def generar(tipo, nombre_proyecto, ubicacion, componentes, otros, nivel_tension,
            tension_nominal, condiciones, condiciones_otras, disenador, constructor,
            salida, plantilla="plantilla_manual.docx"):
    """disenador / constructor: None si no se conoce, o dict con firma (bytes), cedula y mp."""
    componentes, otros = list(componentes), [o for o in otros if o.strip()]
    condiciones_otras = [c for c in condiciones_otras if c.strip()]

    genero = TIPOS.get(tipo, "m")
    articulo, alimentado = ("un", "alimentado") if genero == "m" else ("una", "alimentada")
    tiene_ups = UPS in componentes
    tiene_solar = SOLAR in componentes or SOLAR_BAT in componentes
    tiene_baterias = SOLAR_BAT in componentes

    doc = DocxTemplate(plantilla)
    ctx = {
        "tipo": tipo,
        "nombre_proyecto": nombre_proyecto,
        "ubicacion": ubicacion,
        "anio": datetime.now().year,
        "tipo_descripcion": f"{articulo} {tipo.lower()} {alimentado}",
        "tension_nominal": tension_nominal,
        "nivel_tension": nivel_tension,
        "componentes": componentes + otros,
        "tiene_ups": tiene_ups,
        "tiene_solar": tiene_solar,
        "equipos_solares": ("paneles, inversor y banco de baterías" if tiene_baterias
                            else "paneles e inversor"),
        "zonas_dc": ("tableros, paneles, inversores y baterías" if tiene_baterias
                     else "tableros, paneles e inversores"),
        "condiciones_tecnicas": (["Existencia de UPS"] if tiene_ups else [])
                                + list(condiciones) + condiciones_otras,
        "sistemas_disponibles": SISTEMAS_FIJOS
                                + [c for c in componentes if c != TDG] + otros,
    }
    for clave, datos in (("disenador", disenador), ("constructor", constructor)):
        ctx[f"{clave}_conocido"] = datos is not None
        ctx[f"firma_{clave}"] = _firma(doc, datos["firma"]) if datos else ""
        ctx[f"cedula_{clave}"] = datos["cedula"] if datos else ""
        ctx[f"mp_{clave}"] = datos["mp"] if datos else ""

    doc.render(ctx, autoescape=True)
    buf = io.BytesIO()
    doc.save(buf)
    resultado = buf.getvalue()

    # Marca de agua azul: una por firma conocida (celda 0 = diseñador, celda 1 = constructor)
    celdas = [i for i, d in enumerate((disenador, constructor)) if d is not None]
    if celdas:
        png, ancho, alto = crear_marca_agua(nombre_proyecto, ubicacion)
        resultado = agregar_marca_agua(resultado, celdas, png, ancho, alto)

    if hasattr(salida, "write"):
        salida.write(resultado)
    else:
        with open(salida, "wb") as f:
            f.write(resultado)


def _buscar_soffice():
    for nombre in ("soffice", "libreoffice"):
        ruta = shutil.which(nombre)
        if ruta:
            return ruta
    for ruta in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                 r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                 "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if os.path.exists(ruta):
            return ruta
    return None


_BLOQUEO_WORD = threading.Lock()   # Word se maneja de a una conversión por vez


def _pdf_con_word(docx_bytes):
    """Convierte con Microsoft Word (Windows). Abre una instancia propia y oculta de Word,
    sin tocar los documentos que el usuario tenga abiertos."""
    import pythoncom
    import win32com.client
    with _BLOQUEO_WORD, tempfile.TemporaryDirectory() as tmp:
        entrada, salida = os.path.join(tmp, "manual.docx"), os.path.join(tmp, "manual.pdf")
        with open(entrada, "wb") as f:
            f.write(docx_bytes)
        pythoncom.CoInitialize()          # Streamlit corre en hilos: COM necesita inicializarse
        word = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            doc = word.Documents.Open(entrada, ReadOnly=True, AddToRecentFiles=False, Visible=False)
            doc.SaveAs2(salida, FileFormat=17)   # 17 = wdFormatPDF
            doc.Close(False)
        finally:
            if word is not None:
                word.Quit(False)
            pythoncom.CoUninitialize()
        with open(salida, "rb") as f:
            return f.read()


def _pdf_con_libreoffice(docx_bytes):
    soffice = _buscar_soffice()
    if not soffice:
        raise RuntimeError("LibreOffice no está instalado.")
    with tempfile.TemporaryDirectory() as tmp:
        entrada = os.path.join(tmp, "manual.docx")
        with open(entrada, "wb") as f:
            f.write(docx_bytes)
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp, entrada],
                       check=True, capture_output=True, timeout=120,
                       env={**os.environ, "HOME": tmp})
        with open(os.path.join(tmp, "manual.pdf"), "rb") as f:
            return f.read()


def docx_a_pdf(docx_bytes):
    """Devuelve el PDF del .docx: usa Word si está disponible (Windows) y LibreOffice como respaldo."""
    if sys.platform == "win32":
        try:
            return _pdf_con_word(docx_bytes)
        except Exception:
            pass                          # sin Word o falló: se intenta con LibreOffice
    return _pdf_con_libreoffice(docx_bytes)


if __name__ == "__main__":
    # Firma de ejemplo (garabato transparente) solo para las pruebas
    from PIL import ImageDraw
    img = Image.new("RGBA", (500, 200), (0, 0, 0, 0))
    ImageDraw.Draw(img).line([(20, 150), (90, 40), (140, 160), (200, 60), (260, 150), (330, 70), (480, 120)],
                             fill=(10, 10, 80, 255), width=6, joint="curve")
    b = io.BytesIO(); img.save(b, "PNG"); firma = b.getvalue()

    persona = {"firma": firma, "cedula": "1.000.810.095", "mp": "CN205-184878"}
    generar("Local comercial", "Ocean Blue Pets", "Plaza Central", COMPONENTES[:4], [], "baja",
            "208/120 V", CONDICIONES, [], persona, None, "prueba_original.docx")
    generar("Consultorio", "Clínica Dental Sonrisa", "Chapinero",
            [COMPONENTES[0], COMPONENTES[2], SOLAR_BAT], ["Planta de emergencia"], "media",
            "13,2 kV / 208/120 V", CONDICIONES[:2], ["Equipos de rayos X"], persona,
            {"firma": firma, "cedula": "52.123.456", "mp": "CN205-111111"}, "prueba_consultorio.docx")
    generar("Oficina", "Estudio Beta", "Usaquén", [COMPONENTES[0], COMPONENTES[2]], [], "baja",
            "208/120 V", [], [], None, None, "prueba_sin_firmas.docx")
