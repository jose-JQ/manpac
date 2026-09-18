# ManPAC — Extractor de facturas automotrices con IA

Sistema local para leer facturas y proformas de vehículos (PDF o imagen), extraer los datos clave con modelos de IA y validarlos contra un catálogo de vehículos eléctricos. Si la extracción es confiable, el registro entra al Excel de éxitos; si no, el archivo pasa a revisión manual.

Probado en **Windows 10 + Python 3.14**.

---

## Qué hace

A partir de un documento, el sistema intenta obtener:

| Campo | Significado |
| --- | --- |
| `marca` / `modelo` | Vehículo comercial detectado en el documento |
| `marca_base` / `modelo_base` | Mejor coincidencia en `vehiculos.db` |
| `cantidad` | Unidades |
| `beneficiario` | Nombre del cliente |
| `ci` | Cédula o RUC (solo dígitos) |
| `costo_mas_alto` | Ítem vehicular más caro |
| `costo_total` | Total a pagar |
| `estado` | `ÉXITO` o `REVISIÓN (...motivo...)` |
| `paginas_pdf` | Páginas del PDF que formaron esa factura |

Un PDF de varias páginas puede contener **varias facturas**. El pipeline las segmenta y las va devolviendo en tiempo real.

---

## Arquitectura

Hay tres procesos que trabajan juntos:

```mermaid
flowchart LR
    usuario[Usuario] --> ui[Streamlit<br/>app.py :8501]
    ui -->|POST multipart + stream NDJSON| api[FastAPI<br/>api_backend.py :8000]
    api --> pipe[pipeline.py]
    pipe --> ocr[OCR: PyMuPDF / Tesseract / DocTR]
    ocr --> llm[Ollama: Qwen 2.5, MiniCPM-V, Llama 3.1]
    llm --> val[Validación fuzzy vs SQLite]
    val --> out[Excel, log y carpetas de salida]
    val --> ui
```

| Archivo | Rol |
| --- | --- |
| `app.py` | Interfaz: subir lotes, ver extracciones y administrar marcas/modelos |
| `api_backend.py` | API REST. Recibe el archivo, llama al pipeline y **transmite** cada factura (`application/x-ndjson`) |
| `pipeline.py` | Cerebro: OCR, LLMs, validación, Excel y logs |
| `base.py` | Esquema ORM de `vehiculos.db` (marca → modelo → especificación técnica) |
| `vehiculos.db` | Catálogo SQLite con el que se confirma si el vehículo existe |
| `alimentar_base.ipynb` | Notebook opcional para reconstruir/ampliar el catálogo |

La UI **no** carga los modelos de IA. Solo habla con la API. Por eso hay que tener **los dos** servidores encendidos.

---

## Cómo funciona el pipeline

1. **Lectura del archivo**
   - PDF nativo: extrae texto con PyMuPDF.
   - PDF escaneado o imagen: rasteriza con Poppler (`pdf2image`) y pasa por Tesseract. Si el texto sigue siendo pobre, usa **DocTR** (detección `db_resnet50` + reconocimiento `crnn_vgg16_bn`).
2. **Segmentación**
   - Busca encabezados tipo `FACTURA/PROFORMA` + número `000-000-000000000` para partir un PDF en documentos lógicos.
3. **Intento 1 — texto**
   - **Qwen 2.5 3B** (`qwen2.5:3b`) convierte el texto en JSON.
4. **Validación**
   - Costos (formatos `38.500,00` / `38,500.00`; corrige dígitos concatenados).
   - CI/RUC numérico.
   - Beneficiario real (no plantillas tipo “Nombre del cliente”).
   - Marca no puede ser un color.
   - Coincidencia **fuzzy** contra el catálogo, ignorando sufijos técnicos (`AC 5P 4X2 TA EV`).
5. **Intento 2 — si falla**
   - **MiniCPM-V** lee hasta 2 páginas como imagen.
   - **Llama 3.1** actúa de juez y fusiona ambos JSON.
6. **Salida**
   - Cada factura se `yield` a la API (la UI la pinta al momento).
   - Éxitos → `procesados_exito/` + `reporte_extracciones.xlsx`.
   - Fallos → `revision_manual/` + `log_revision_fallos.txt`.

---

## Qué se necesita instalar

### 1. Python (pip)

Ver `requirements.txt`. Las librerías y para qué sirven:

| Librería | Uso |
| --- | --- |
| `streamlit` | Interfaz web local |
| `fastapi` + `uvicorn` + `python-multipart` | API y carga de archivos |
| `requests` | Cliente HTTP de la UI hacia la API |
| `SQLAlchemy` | ORM y consultas a `vehiculos.db` |
| `pandas` + `openpyxl` | Excel de extracciones |
| `pymupdf` | Texto nativo de PDF |
| `pdf2image` | PDF → imágenes (requiere Poppler) |
| `pytesseract` | OCR clásico (requiere Tesseract) |
| `python-doctr` + `torch` | OCR neuronal para escaneos difíciles |
| `opencv-python` + `numpy` + `pillow` | Preprocesado de imagen |
| `ollama` | Cliente de los modelos locales |

Instalación rápida:

```bat
instalar.bat
```

O a mano:

```bat
python -m venv venv
venv\Scripts\python.exe -m pip install -r requirements.txt
```

La primera vez que DocTR corre, **descarga pesos** (hace falta internet).

### 2. Programas externos (no van por pip)

| Programa | Para qué | Dónde |
| --- | --- | --- |
| **Ollama** | Corre los LLMs en local | https://ollama.com |
| **Tesseract OCR** (idioma `spa`) | OCR de páginas escaneadas | https://github.com/UB-Mannheim/tesseract/wiki |
| **Poppler** | Convertir PDF a imagen | https://github.com/oschwartz10612/poppler-windows/releases |

Modelos de Ollama que el pipeline espera:

```bat
ollama pull qwen2.5:3b
ollama pull minicpm-v
ollama pull llama3.1
```

Rutas por defecto (se pueden cambiar con variables de entorno):

```bat
set POPPLER_PATH=C:\poppler-26.07.0\Library\bin
set TESSERACT_CMD=C:\Program Files\Tesseract-OCR\tesseract.exe
```

---

## Cómo arrancar

1. Deja **Ollama** en ejecución.
2. Ejecuta `iniciar.bat` (abre API + UI).
3. Abre http://localhost:8501
4. En **Procesamiento por Lotes**, sube PDF/PNG/JPG y pulsa *Enviar al Servidor IA*.

A mano, en dos terminales:

```bat
venv\Scripts\python.exe api_backend.py
venv\Scripts\streamlit.exe run app.py
```

Comprobación: http://127.0.0.1:8000/api/health

---

## Uso de la interfaz

1. **Procesamiento por Lotes** — sube uno o varios documentos. Cada factura aparece en un expander (verde = éxito, naranja = revisión).
2. **Reporte de Extracciones** — tabla del Excel consolidado y botón de descarga.
3. **Base de Datos de Vehículos** — alta/edición/baja de marcas y modelos. Si un vehículo no está en el catálogo, las facturas reales irán a revisión aunque la IA lea bien el papel.

---

## API

| Método | Ruta | Descripción |
| --- | --- | --- |
| `GET` | `/api/health` | Estado del servidor |
| `POST` | `/api/procesar-factura` | `multipart/form-data` campo `file`. Respuesta NDJSON (una línea JSON por factura) |

---

## Estructura al enviarla

Este paquete **no incluye** facturas de clientes, logs ni el `venv` (pesan y pueden tener datos personales). El receptor instala dependencias con `instalar.bat`.

```
pipeline.py
app.py
api_backend.py
base.py
vehiculos.db
requirements.txt
instalar.bat
iniciar.bat
README.md
alimentar_base.ipynb   (opcional, para regenerar el catálogo)
```

Carpetas que se crean solas al procesar: `temp_api_uploads/`, `procesados_exito/`, `revision_manual/`.

---

## Limitaciones

- Corre en local; no es un despliegue en la nube.
- MiniCPM-V y Llama 3.1 piden RAM/VRAM. En CPU será lento.
- Vehículos que no existan en `vehiculos.db` (o con OCR muy pobre) salen como `REVISIÓN`.
- Las rutas de Poppler/Tesseract están pensadas para Windows.
