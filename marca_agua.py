"""Genera la marca de agua azul (imagen) y la coloca flotando sobre una firma."""
import io
import textwrap

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Emu
from lxml import etree
from PIL import Image, ImageDraw, ImageFont

DPI = 300
ANGULO = 58                       # grados, el texto sube de izquierda a derecha
AZUL = (23, 96, 196, 185)         # RGBA: azul semitransparente
FUENTES = ["arialbd.ttf", "Arial Bold.ttf", "LiberationSans-Bold.ttf", "DejaVuSans-Bold.ttf"]


def _fuente(px):
    for nombre in FUENTES:
        try:
            return ImageFont.truetype(nombre, px)
        except OSError:
            continue
    return ImageFont.load_default(size=px)


def crear_marca_agua(nombre, ubicacion):
    """Devuelve (png_bytes, ancho_emu, alto_emu) con el texto azul inclinado y fondo transparente."""
    lineas = ["VÁLIDO ÚNICAMENTE", "PARA EL PROYECTO"]
    lineas += textwrap.wrap(nombre.strip().upper(), 18) or [""]
    lineas += textwrap.wrap(ubicacion.strip().upper(), 18) or [""]
    fuente = _fuente(46)
    medida = ImageDraw.Draw(Image.new("RGBA", (10, 10)))
    altura_linea = int(fuente.size * 1.35)
    anchos = [medida.textlength(l, font=fuente) for l in lineas]
    w, h = int(max(anchos)) + 40, altura_linea * len(lineas) + 40
    bloque = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(bloque)
    for i, (l, a) in enumerate(zip(lineas, anchos)):
        d.text(((w - a) / 2, 20 + i * altura_linea), l, font=fuente, fill=AZUL)
    rot = bloque.rotate(ANGULO, expand=True, resample=Image.BICUBIC)
    rot = rot.crop(rot.getchannel("A").getbbox())
    buf = io.BytesIO()
    rot.save(buf, "PNG", dpi=(DPI, DPI))
    emu = lambda px: int(px / DPI * 914400)
    return buf.getvalue(), emu(rot.width), emu(rot.height)


def agregar_marca_agua(docx_bytes, celdas, png, ancho, alto):
    """Coloca la imagen flotante, sobre el texto, en el primer párrafo de cada celda indicada
    de la última tabla del documento (la tabla de firmas)."""
    doc = Document(io.BytesIO(docx_bytes))
    tabla = doc.tables[-1]
    for k, idx in enumerate(celdas):
        parrafo = tabla.rows[0].cells[idx].paragraphs[0]
        run = parrafo.add_run()
        run.add_picture(io.BytesIO(png), width=Emu(ancho), height=Emu(alto))
        drawing = run._r.find(qn("w:drawing"))
        inline = drawing.find(qn("wp:inline"))
        anchor = etree.Element(qn("wp:anchor"), {
            "distT": "0", "distB": "0", "distL": "0", "distR": "0", "simplePos": "0",
            "relativeHeight": str(251650000 + k), "behindDoc": "0", "locked": "0",
            "layoutInCell": "1", "allowOverlap": "1"})
        etree.SubElement(anchor, qn("wp:simplePos"), {"x": "0", "y": "0"})
        ph = etree.SubElement(anchor, qn("wp:positionH"), {"relativeFrom": "column"})
        etree.SubElement(ph, qn("wp:posOffset")).text = str(int(0.8 * 360000))
        pv = etree.SubElement(anchor, qn("wp:positionV"), {"relativeFrom": "paragraph"})
        etree.SubElement(pv, qn("wp:posOffset")).text = "0"
        anchor.append(inline.find(qn("wp:extent")))
        etree.SubElement(anchor, qn("wp:effectExtent"), {"l": "0", "t": "0", "r": "0", "b": "0"})
        etree.SubElement(anchor, qn("wp:wrapNone"))
        docpr = inline.find(qn("wp:docPr"))
        docpr.set("id", str(9000 + k)); docpr.set("name", "Marca de agua %d" % (k + 1))
        anchor.append(docpr)
        anchor.append(inline.find(qn("wp:cNvGraphicFramePr")))
        anchor.append(inline.find(qn("a:graphic")))
        drawing.remove(inline)
        drawing.append(anchor)
    salida = io.BytesIO()
    doc.save(salida)
    return salida.getvalue()
