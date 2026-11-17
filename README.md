# Generador de Documentos Eléctricos (O&M, RETIE y Tableros)

Aplicación para generar de forma rápida los documentos oficiales de las instalaciones eléctricas. Permite rellenar los datos comunes **una sola vez**, seleccionar los documentos necesarios y descargarlos listos en **Word (.docx)**, **PDF** o en un archivo **.zip** con todo el paquete.

La aplicación corre de forma local en tu propio computador y se maneja desde el navegador web(no requiere internet).

---

## Documentos Disponibles

| Documento | ¿Qué incluye? |
|---|---|
| **Manual de Operación y Mantenimiento** | Componentes, condiciones técnicas, firmas opcionales de diseñador y constructor (con marca de agua). |
| **Declaración de Cumplimiento - Constructor** | Datos de obra, constructor, diseñador y propietario, con firma obligatoria del constructor. |
| **Declaración de Cumplimiento - Diseñador** | Normativa, desviaciones, datos de diseño y propiedad, con firma obligatoria del diseñador. |
| **Carta de Diseñador** | Tabla de ítems RETIE 10.1.1 y firma del diseñador. |
| **Verificación de Tableros** | Tablas de verificación por tablero y firmas de diseñador y constructor. |

---

## Requisitos Previos (Solo para la primera instalación)

1. **Python (versión 3.10 o superior):** Descárgalo desde [python.org](https://www.python.org/downloads/). 
   > ⚠️ **Importante durante la instalación:** Marca la casilla inferior que dice **"Add python.exe to PATH"**.
2. **Microsoft Word:** Necesario en segundo plano si deseas que la aplicación exporte los documentos directamente a formato PDF.
3. **Conexión a internet:** Solo se requiere la primera vez para descargar las librerías necesarias.

---

## Instalación en un Computador Nuevo

1. **Descarga el proyecto** desde GitHub (botón verde *Code → Download ZIP*) y descomprímelo en una carpeta de tu computador (por ejemplo, en `Documentos`).
2. **Prepara la aplicación:** Haz doble clic en el archivo **`iniciar_app.bat`**. La primera vez configurará todo automáticamente (puede tardar unos minutos) y abrirá la aplicación en tu navegador.
3. **Crea tu acceso directo:** Haz doble clic en **`crear_acceso_directo.bat`** para colocar un icono de acceso directo en tu Escritorio con el logo de la empresa.

---

## Instrucciones de Uso

1. Haz doble clic en el icono de la aplicación en tu **Escritorio**.
2. **Selecciona los documentos** que necesitas generar (puedes marcar uno o varios a la vez).
3. **Completa los datos comunes:** Solo te pedirá la información necesaria para los documentos que marcaste (proyecto, ubicación, propietario, fechas, etc.). No necesitas repetir la misma información.
4. **Sube las firmas:** Las imágenes de las firmas se adaptan automáticamente (recorta márgenes y vuelve transparente el fondo blanco).
5. **Genera y descarga:** Pulsa el botón de generar para obtener tus archivos en Word, PDF o todos comprimidos en un `.zip`.

> **Nota importante:** La ventana negra (terminal) que se abre al iniciar es el "motor" de la aplicación. **Debe permanecer abierta** mientras uses el programa. Para apagarla por completo, simplemente cierra esa ventana negra.

---

## Solución de Problemas Frecuentes

* **El archivo `.bat` dice que no encuentra Python:** Asegúrate de haber marcado la opción *Add Python to PATH* al instalar Python, o reinstálalo marcando dicha opción.
* **La aplicación no muestra los últimos cambios:** Cierra la ventana negra del servidor y vuelve a abrir el acceso directo.
* **No se generan los PDFs:** Asegúrate de tener Microsoft Word instalado en el equipo.
* **El navegador no abre automáticamente:** Entra manualmente a tu navegador web y escribe la dirección: `http://localhost:8501`.