"""Convierte el .docx original en plantilla con marcadores Jinja (se corre una sola vez).

Uso: python crear_plantilla.py ORIGINAL.docx plantilla_manual.docx
"""
import copy, os, re, shutil, sys, tempfile, zipfile
from lxml import etree

ORIG, OUT = sys.argv[1], sys.argv[2]
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
ns = {'w': W, 'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
      'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'}
q = lambda t: '{%s}%s' % (W, t)
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'


def txt(p):
    return ''.join(p.xpath('.//w:t/text()', namespaces=ns))


def set_text(p, text):
    """Deja el primer run con texto (sin el espaciado raro por palabra) y borra los demás."""
    runs = [r for r in p.findall(q('r')) if r.find(q('t')) is not None]
    keep = next((r for r in runs if r.find(q('t')).text and r.find(q('t')).text.strip()), runs[0])
    rpr = keep.find(q('rPr'))
    if rpr is not None:
        for s in rpr.findall(q('spacing')):
            rpr.remove(s)
    for r in runs:
        if r is not keep:
            p.remove(r)
    t = keep.find(q('t'))
    t.text = text
    t.set(XML_SPACE, 'preserve')
    return keep


def ctrl(tag):
    """Párrafo de control de docxtpl ({%p ... %}); desaparece al renderizar."""
    p = etree.Element(q('p'))
    r = etree.SubElement(p, q('r'))
    t = etree.SubElement(r, q('t'))
    t.text = tag
    t.set(XML_SPACE, 'preserve')
    return p


def clone(model, text):
    p = copy.deepcopy(model)
    for r in p.findall(q('r')):           # fuera dibujos y runs del modelo
        if r.find(q('t')) is None or r.find(q('drawing')) is not None:
            p.remove(r)
    set_text(p, text)
    return p


def wrap_if(first, last, cond):
    first.addprevious(ctrl('{%p if ' + cond + ' %}'))
    last.addnext(ctrl('{%p endif %}'))


def replace_with_loop(olds, var, item='c'):
    """Reemplaza varias viñetas por un bucle que genera una viñeta por elemento de `var`."""
    model = olds[0]
    first = olds[0]
    first.addprevious(ctrl('{%p for ' + item + ' in ' + var + ' %}'))
    bullet = clone(model, '{{ ' + item + ' }}')
    first.addprevious(bullet)
    bullet.addnext(ctrl('{%p endfor %}'))
    for o in olds:
        o.getparent().remove(o)


work = tempfile.mkdtemp()
with zipfile.ZipFile(ORIG) as z:
    z.extractall(work)

dpath = os.path.join(work, 'word/document.xml')
doc = etree.parse(dpath)
paras = list(doc.iter(q('p')))
body = doc.getroot().find(q('body'))
by = lambda start: next(p for p in paras if txt(p).startswith(start))
exact = lambda s: next(p for p in paras if txt(p) == s)

# ---- Portada: tipo + nombre, ubicación debajo (cambios 2 y 3) ----
p = by('INSTALACIÓN ELÉCTRICA – LOCAL')
run = set_text(p, 'INSTALACIÓN ELÉCTRICA – {{ tipo|upper }} {{ nombre_proyecto|upper }}')
br = etree.Element(q('r')); br.append(copy.deepcopy(run.find(q('rPr')))); etree.SubElement(br, q('br'))
run2 = copy.deepcopy(run); run2.find(q('t')).text = '{{ ubicacion|upper }}'
run.addnext(br); br.addnext(run2)

# Año automático
set_text(by('AÑO 2026'), 'AÑO {{ anio }}')

# ---- Punto 1: descripción general (cambios 4 y 5) ----
p = by('La instalación eléctrica corresponde')
model_body = copy.deepcopy(p)
set_text(p, 'La instalación eléctrica corresponde a {{ tipo_descripcion }} a {{ tension_nominal }}, con sistema de distribución que incluye:')
b1 = by('Tablero general de distribución (TDG)')
b2 = b1.getnext(); b3 = b2.getnext(); b4 = b3.getnext()
assert txt(b4).startswith('Cargas especiales')
replace_with_loop([b1, b2, b3, b4], 'componentes')
cumple = clone(model_body, 'El diseño cumple con RETIE y NTC 2050')
ppr = cumple.find(q('pPr'))
for sp in ppr.findall(q('spacing')):
    ppr.remove(sp)
sp = etree.Element(q('spacing')); sp.set(q('before'), '120'); sp.set(q('after'), '240')
ppr.insert(1 if ppr.find(q('pStyle')) is not None else 0, sp)
# se ubica justo después del bucle (antes del título 2)
heading2 = exact('OPERACIÓN Y MANTENIMIENTO DE LA INSTALACIÓN') if False else None
endfor = next(e for e in body.iter(q('p')) if txt(e) == '{%p endfor %}')
endfor.addnext(cumple)

# ---- Punto 2a: nivel de tensión (cambio 6) + UPS (cambios 7 y 8) + solar ----
set_text(by('La instalación es de baja tensión'),
         'La instalación es de {{ nivel_tension }} tensión, por lo tanto, las actividades deben enfocarse en:')
ups7 = exact('UPS (equipos electrónicos sensibles)'); wrap_if(ups7, ups7, 'tiene_ups')
ups8 = exact('Circuitos regulados (UPS)'); wrap_if(ups8, ups8, 'tiene_ups')

last_a = by('Todas las canalizaciones metálicas')
sol_l1 = clone(last_a, 'Sistema solar fotovoltaico ({{ equipos_solares }}):')
model_l2 = exact('Conexiones')
sol_a = clone(model_l2, 'Limpieza periódica, a menos que el diseño justifique lo contrario')
sol_b = clone(model_l2, 'Revisión visual periódica de conexiones y prueba de continuidad')
last_a.addnext(sol_l1); sol_l1.addnext(sol_a); sol_a.addnext(sol_b)
wrap_if(sol_l1, sol_b, 'tiene_solar')

# ---- Punto 2c: reemplazo de equipos, UPS (cambio 9) ----
en_ups = exact('En UPS:'); resp = exact('Respetar capacidad de carga y autonomía')
wrap_if(en_ups, resp, 'tiene_ups')

# ---- Tensión nominal en condiciones de diseño ----
set_text(by('Tensión: 208/120 V'), 'Tensión: {{ tension_nominal }}')

# ---- Punto 2e: prohibiciones (cambio 10 + solar) ----
ints_ups = exact('Intervenir UPS sin autorización')
sol_proh = clone(ints_ups, 'Manipular sin protecciones las zonas de carga DC ({{ zonas_dc }})')
ints_ups.addnext(sol_proh)
wrap_if(ints_ups, ints_ups, 'tiene_ups')
wrap_if(sol_proh, sol_proh, 'tiene_solar')

# ---- Punto 2g: condiciones técnicas especiales (cambio 11) ----
c1 = exact('Existencia de UPS')
cs = [c1]
for _ in range(4):
    cs.append(cs[-1].getnext())
assert txt(cs[-1]).startswith('Uso intensivo de iluminación LED')
replace_with_loop(cs, 'condiciones_tecnicas', item='k')

# ---- Punto 2i: sistemas disponibles (cambio 12) ----
s1 = exact('Breakers (protección principal)')
s2 = s1.getnext(); s3 = s2.getnext()
assert txt(s2) == 'Tablero general' and txt(s3).startswith('UPS (respaldo parcial)')
replace_with_loop([s1, s2, s3], 'sistemas_disponibles', item='s')

# ---- Texto sobrante de otra IA ----
p = by('Esto coincide con lo definido en tu unifilar')
p.getparent().remove(p)

# ---- Firmas y marca de agua (cambios 13 y 14) ----
p285 = by('Si la falla persiste, contactar personal calificado')
idx = list(body).index(p285)
for el in list(body)[idx + 1:]:
    if el.tag != q('sectPr'):
        body.remove(el)

# Dentro de ese párrafo viven el logo grande de fondo ("Image 10") y la marca de agua azul ("Group 11")
logo = next(d for d in p285.iter(q('drawing'))
            if d.xpath('.//wp:docPr/@name', namespaces=ns)[0] == 'Image 10')
logo.getparent().remove(logo)
for ac in list(p285.iter('{%s}AlternateContent' % ns['mc'])):
    ac.getparent().remove(ac)   # marca de agua con el nombre "horneado": se genera por proyecto

anchor_p = etree.fromstring(
    '<w:p xmlns:w="%s"><w:pPr><w:keepNext/><w:spacing w:before="480" w:after="0"/></w:pPr></w:p>' % W)
r_logo = etree.SubElement(anchor_p, q('r')); r_logo.append(logo)
p285.addnext(anchor_p)

TXT = 'xml:space="preserve"'
def run_x(text, underline=False):
    return ('<w:r><w:rPr><w:b/>%s<w:sz w:val="24"/></w:rPr><w:t %s>%s</w:t></w:r>'
            % ('<w:u w:val="single"/>' if underline else '', TXT, text))
def para_x(runs, border=False):
    ppr = ('<w:pPr><w:keepNext/>%s<w:spacing w:before="0" w:after="60" w:line="360" w:lineRule="auto"/>%s</w:pPr>'
           % ('<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"/></w:pBdr>' if border else '',
              '<w:ind w:right="200"/>' if border else ''))
    return '<w:p>%s%s</w:p>' % (ppr, runs)
def cell_x(titulo, v):
    # v = sufijo de variables: disenador / constructor
    return ('<w:tc><w:tcPr><w:tcW w:w="5040" w:type="dxa"/></w:tcPr>'
            + para_x(run_x(titulo))
            + para_x('<w:r><w:t>{%%p if %s_conocido %%}</w:t></w:r>' % v)
            + para_x(run_x('FIRMA: ') + '<w:r><w:t>{{ firma_%s }}</w:t></w:r>' % v, border=True)
            + para_x(run_x('CEDULA: ') + run_x('{{ cedula_%s }}' % v, underline=True))
            + para_x(run_x('M.P: ') + run_x('{{ mp_%s }}' % v, underline=True))
            + para_x('<w:r><w:t>{%p else %}</w:t></w:r>')
            + para_x(run_x('FIRMA:__________________________'))
            + para_x(run_x('CEDULA:_________________________'))
            + para_x(run_x('M.P:___________________________'))
            + para_x('<w:r><w:t>{%p endif %}</w:t></w:r>')
            + '</w:tc>')
tbl = etree.fromstring((
    '<w:tbl xmlns:w="%s"><w:tblPr><w:tblW w:w="10080" w:type="dxa"/><w:tblInd w:w="720" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
    '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>'
    '<w:tblGrid><w:gridCol w:w="5040"/><w:gridCol w:w="5040"/></w:tblGrid>'
    '<w:tr><w:trPr><w:cantSplit/></w:trPr>%s%s</w:tr></w:tbl>'
    % (W, cell_x('DISEÑADOR:', 'disenador'), cell_x('CONSTRUCTOR:', 'constructor'))).encode('utf8'))
anchor_p.addnext(tbl)
tbl.addnext(etree.fromstring(
    '<w:p xmlns:w="%s"><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:p>' % W))

# ---- Paginación: quitar los saltos de sección "fijos" del cuerpo ----
# El original traía 7 saltos de sección (siguiente página) puestos a mano para que cada página
# terminara justo donde estaba su contenido. Al cambiar la cantidad de texto, esos saltos dejaban
# páginas casi vacías o listas cortadas. Se conservan los 2 primeros (portada e índice).
con_seccion = [e for e in body.findall(q('p'))
               if e.find(q('pPr')) is not None and e.find(q('pPr')).find(q('sectPr')) is not None]
for e in con_seccion[2:]:
    e.find(q('pPr')).remove(e.find(q('pPr')).find(q('sectPr')))
    if not txt(e).strip() and not e.xpath('.//w:drawing|.//w:pict', namespaces=ns):
        body.remove(e)

# La tabla de periodicidad venía partida en dos (por uno de esos saltos): se une en una sola
tablas = body.findall(q('tbl'))
t1 = next(t for t in tablas if 'Medición del SPT' in txt(t))
t2 = next(t for t in tablas if 'Termografía' in txt(t))
if t1 is not t2:
    nodo = t1.getnext()
    while nodo is not None and nodo is not t2:      # párrafos vacíos que había entre las dos
        sig = nodo.getnext()
        if nodo.tag == q('p') and not txt(nodo).strip() and not nodo.xpath('.//w:drawing', namespaces=ns):
            body.remove(nodo)
        nodo = sig
    for tr in t2.findall(q('tr')):
        t1.append(tr)
    body.remove(t2)


def vacio(e):
    return (e.tag == q('p') and not txt(e).strip()
            and not e.xpath('.//w:drawing|.//w:pict|.//w:br|.//w:sectPr', namespaces=ns))


def con_num(e):
    return (e is not None and e.tag == q('p') and e.find(q('pPr')) is not None
            and e.find(q('pPr')).find(q('numPr')) is not None)


# Párrafos vacíos que quedaron en medio de una lista (eran los "rellenos" de fin de página del original)
for e in list(body.findall(q('p'))):
    if vacio(e):
        ant, sig = e.getprevious(), e.getnext()
        while sig is not None and vacio(sig):
            sig = sig.getnext()
        while ant is not None and vacio(ant):
            ant = ant.getprevious()
        if con_num(ant) and con_num(sig):
            body.remove(e)

def keep_next(p):
    ppr = p.find(q('pPr'))
    if ppr is None:
        ppr = etree.Element(q('pPr')); p.insert(0, ppr)
    if ppr.find(q('keepNext')) is None:
        ppr.insert(1 if ppr.find(q('pStyle')) is not None else 0, etree.Element(q('keepNext')))


# Títulos y líneas introductorias ("Se prohíbe:") no quedan solos al final de una página
for e in body.iter(q('p')):
    t = txt(e).strip()
    if not t or t.startswith('{%'):
        continue
    tiene_num = e.find(q('pPr')) is not None and e.find(q('pPr')).find(q('numPr')) is not None
    if t.endswith(':') or (tiene_num and len(t) > 8 and t == t.upper()):
        keep_next(e)

# Los vacíos que hay justo antes de una tabla se pegan a ella (así el título no queda solo)
for tbl_ in body.findall(q('tbl')):
    prev = tbl_.getprevious()
    while prev is not None and vacio(prev):
        keep_next(prev)
        prev = prev.getprevious()

# Las tablas pequeñas no se parten entre páginas
for tbl_ in body.iter(q('tbl')):
    filas = tbl_.findall(q('tr'))
    for k, tr in enumerate(filas):
        trpr = tr.find(q('trPr'))
        if trpr is None:
            trpr = etree.Element(q('trPr'))
            tr.insert(1 if tr.find(q('tblPrEx')) is not None else 0, trpr)
        if trpr.find(q('cantSplit')) is None:
            trpr.append(etree.Element(q('cantSplit')))
        if k < len(filas) - 1:
            for pp in tr.iter(q('p')):
                keep_next(pp)

doc.write(dpath, xml_declaration=True, encoding='UTF-8', standalone=True)

# ---- Encabezado (cambio 1) ----
hpath = os.path.join(work, 'word/header1.xml')
hdr = etree.parse(hpath)
n = 0
for p in hdr.iter(q('p')):
    if txt(p) == 'OCEAN BLUE PETS PLAZA CENTRAL':
        set_text(p, '{{ nombre_proyecto|upper }} {{ ubicacion|upper }}')
        n += 1
assert n == 2, n
hdr.write(hpath, xml_declaration=True, encoding='UTF-8', standalone=True)

# ---- Quitar imágenes que ya nadie usa (las piezas de la marca de agua y la firma de ejemplo) ----
dxml = open(dpath, encoding='utf8').read()
usados = set(re.findall(r'r:(?:embed|id|link)="(rId\d+)"', dxml))
rels_path = os.path.join(work, 'word/_rels/document.xml.rels')
rels = open(rels_path, encoding='utf8').read()
borrar = []
def quitar(m):
    rid, target = m.group(1), m.group(2)
    if 'media/' in target and rid not in usados:
        borrar.append('word/' + target)
        return ''
    return m.group(0)
rels = re.sub(r'<Relationship [^>]*?Id="(rId\d+)"[^>]*?Target="([^"]+)"[^>]*/>', quitar, rels)
open(rels_path, 'w', encoding='utf8').write(rels)
otros_rels = ''
for f in os.listdir(os.path.join(work, 'word/_rels')):
    if f != 'document.xml.rels':
        otros_rels += open(os.path.join(work, 'word/_rels', f), encoding='utf8').read()
borrar = [b for b in borrar if os.path.basename(b) not in otros_rels]

with zipfile.ZipFile(ORIG) as zin, zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        if item.is_dir() or item.filename in borrar:
            continue
        zout.writestr(item, open(os.path.join(work, item.filename), 'rb').read())
shutil.rmtree(work)
print('Plantilla creada:', OUT, '| imágenes eliminadas:', len(borrar))
