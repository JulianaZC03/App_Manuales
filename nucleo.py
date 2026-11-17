"""Funciones compartidas por todos los documentos: firmas, fechas, listas y conversión a PDF."""
import io
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
from pathlib import Path

from docx.shared import Mm
from docxtpl import InlineImage
from PIL import Image

PLANTILLAS = Path(__file__).parent / "Plantillas_manuales"

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]

DEPARTAMENTOS = [
    "Amazonas", "Antioquia", "Arauca", "Atlántico", "Bogotá D.C.", "Bolívar", "Boyacá",
    "Caldas", "Caquetá", "Casanare", "Cauca", "Cesar", "Chocó", "Córdoba", "Cundinamarca",
    "Guainía", "Guaviare", "Huila", "La Guajira", "Magdalena", "Meta", "Nariño",
    "Norte de Santander", "Putumayo", "Quindío", "Risaralda",
    "San Andrés y Providencia", "Santander", "Sucre", "Tolima", "Valle del Cauca",
    "Vaupés", "Vichada",
]

TIPOS_ID = ["CC", "NIT", "CE", "TI", "Pasaporte", "PPT", "Otro"]


def mayus(texto):
    return str(texto).strip().upper()


def nombre_seguro(texto):
    """Convierte un texto en un nombre de archivo válido."""
    return re.sub(r"[^\w\-]+", "_", texto.strip()).strip("_")


# ---------------------------------------------------------------- firmas
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


def firma_inline(doc, datos):
    """Devuelve la firma como InlineImage lista para usar en la plantilla."""
    png, (w, h) = a_png(datos)
    if w / h >= FIRMA_ANCHO_MM / FIRMA_ALTO_MM:         # firma muy ancha: manda el ancho
        return InlineImage(doc, io.BytesIO(png), width=Mm(FIRMA_ANCHO_MM))
    return InlineImage(doc, io.BytesIO(png), height=Mm(FIRMA_ALTO_MM))


# ---------------------------------------------------------------- PDF
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


# ---------------------------------------------------------------- marca de agua en cuadro de texto
_FUENTES_MEDIDA = ["arialbd.ttf", "Arial Bold.ttf", "LiberationSans-Bold.ttf", "DejaVuSans-Bold.ttf"]


def tamano_marca_agua(texto, ancho_mm, alto_mm, base_pt=11.0, minimo_pt=5.0):
    """Tamaño de letra (en medios puntos, como lo pide Word en w:sz) con el que `texto`
    cabe en un cuadro de texto de ancho x alto (mm, ya sin márgenes internos).
    Parte del tamaño original y lo reduce solo si el texto no cabe."""
    from PIL import ImageFont
    ancho_pt, alto_pt = ancho_mm / 25.4 * 72, alto_mm / 25.4 * 72
    pt = base_pt
    while pt > minimo_pt:
        escala = 10                                        # se mide a 10 px por punto
        fuente = None
        for nombre in _FUENTES_MEDIDA:
            try:
                fuente = ImageFont.truetype(nombre, int(pt * escala))
                break
            except OSError:
                continue
        ancho = lambda s: ((fuente.getlength(s) / escala * 0.85) if fuente else len(s) * pt * 0.56) + 0.5 * len(s)   # Calibri es ~15 % más angosta que Arial
        lineas, actual = 1, ""
        for palabra in texto.split():
            prueba = (actual + " " + palabra).strip()
            if actual and ancho(prueba) > ancho_pt:
                lineas, actual = lineas + 1, palabra
            else:
                actual = prueba
        if lineas * pt * 1.32 <= alto_pt:
            break
        pt -= 0.5
    return int(round(max(pt, minimo_pt) * 2))
