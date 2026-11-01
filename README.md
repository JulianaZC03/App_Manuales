# Generador de manuales de Operación y Mantenimiento (O&M)

Aplicación web local (Streamlit) que genera el **Manual de Operación y Mantenimiento** de una
instalación eléctrica a partir de un formulario. Entrega el documento en **Word (.docx)** y en
**PDF**, con las firmas del diseñador y del constructor y una marca de agua azul sobre cada firma.

La app corre en el propio computador y se usa desde el navegador. Las firmas no se guardan:
solo se usan para generar el documento.

## Estructura del proyecto

```
├── app.py                     Interfaz web (formulario y botones de descarga)
├── generar_manual.py          Llena la plantilla Word y convierte a PDF
├── marca_agua.py              Crea y coloca la marca de agua azul sobre las firmas
├── crear_plantilla.py         Script auxiliar de la plantilla
├── Plantillas_manuales/       Plantillas Word (.docx) con etiquetas de docxtpl
├── Imgs/
│   └── logo.png               Logo de la empresa (icono de la pestaña del navegador)
├── logo.ico                   Icono del acceso directo del Escritorio
├── requirements.txt           Librerías de Python necesarias
├── iniciar_app.bat            Prepara el entorno la primera vez y arranca la app
└── crear_acceso_directo.bat   Crea el acceso directo en el Escritorio
```

## Requisitos (Windows)

1. **Python 3.10 o superior**, desde <https://www.python.org/downloads/>.
   Durante la instalación marca **"Add python.exe to PATH"**.
2. **Microsoft Word** instalado, para exportar a PDF. Sin Word, la app intenta con LibreOffice
   y, si no lo encuentra, entrega solo el .docx.
3. Conexión a internet **solo la primera vez**, para instalar las librerías.

## Instalación en un computador nuevo

1. Descarga el proyecto:
   - Con **GitHub Desktop**: *File → Clone repository…*, elige este repositorio y una carpeta
     (por ejemplo `Documentos\manuales-oym`).
   - O con el botón verde **Code → Download ZIP** y descomprime en una carpeta.
2. Haz doble clic en **`iniciar_app.bat`**. La primera vez crea el entorno `.venv` e instala todo
   (puede tardar unos minutos). Después abre la app en el navegador.
3. Haz doble clic en **`crear_acceso_directo.bat`**. Crea un acceso directo en el Escritorio
   con el logo de la empresa.

A partir de ahí basta con usar el icono del Escritorio.

> **Importante:** la ventana negra (terminal) es el servidor de la app. Debe quedar abierta mientras
> se usa el programa. Para apagar la app, ciérrala.

La app se abre en <http://localhost:8501>.

## Uso

1. Completa los datos del proyecto, las características eléctricas, los componentes y las
   condiciones técnicas.
2. Si conoces al diseñador y/o al constructor, sube la imagen de la firma e indica cédula y M.P.
3. Pulsa **Generar manual** y descarga el Word o el PDF.

## Cómo funciona (resumen técnico)

- `app.py` recoge los datos y llama a `generar()`.
- `generar()` llena la plantilla con **docxtpl** (Jinja dentro del Word). Las firmas se recortan y
  se les quita el fondo blanco (`a_png`).
- `marca_agua.py` dibuja con Pillow el texto azul inclinado ("VÁLIDO ÚNICAMENTE PARA EL PROYECTO…")
  y lo inserta como imagen flotante sobre cada firma.
- `docx_a_pdf()` convierte con **Microsoft Word** (vía `pywin32`) y usa **LibreOffice** como respaldo.
- El logo de fondo vive en el **encabezado** de la plantilla, con posición fija respecto a la
  página, para que aparezca igual en todas las hojas sin importar la longitud del contenido.
- Las firmas de diseñador y constructor están en una tabla de filas (título, firma, cédula, M.P.),
  así que quedan alineadas aunque solo una de las dos esté presente.

## Editar la plantilla

Abre la plantilla de `Plantillas_manuales/` en Word. Las partes variables son etiquetas como
`{{ nombre_proyecto }}` o bloques `{%p if tiene_solar %} … {%p endif %}`. No las borres ni las
edites a mitad de palabra (Word a veces parte el texto internamente: si una etiqueta falla,
bórrala y vuelve a escribirla de una sola vez).

## Problemas frecuentes

| Síntoma | Solución |
|---|---|
| El `.bat` dice que no encuentra Python | Instala Python marcando *Add to PATH*, o abre `iniciar_app.bat` con el Bloc de notas y escribe la ruta de `python.exe` (la muestra `where python`). |
| Cambié código pero la app muestra el comportamiento viejo | Cierra la ventana negra y abre la app de nuevo. |
| No sale el PDF | Falta Word (y LibreOffice), o falta `pywin32`. Ejecuta `.venv\Scripts\python.exe -m pip install -r requirements.txt`. |
| El navegador no abre | Entra manualmente a <http://localhost:8501>. |
| `ImportError` al arrancar | Hay archivos de versiones distintas mezclados: actualiza todos con *Pull origin*. |

## Actualizar a una versión nueva

Con GitHub Desktop: *Fetch origin → Pull origin*. Cierra la ventana negra y abre la app otra vez.
Si cambió `requirements.txt`, ejecuta:

```
.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Privacidad

No subas al repositorio manuales generados, firmas ni cédulas de personas. El `.gitignore` ya
excluye los archivos más comunes (`Manual_OyM_*`, `prueba_*`, `firma*`). Se recomienda mantener
el repositorio **privado**.
