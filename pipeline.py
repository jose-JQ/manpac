import os
import re
import json
import shutil
import cv2
import time
import numpy as np
import pymupdf
from pdf2image import convert_from_path
import pytesseract
from doctr.io import DocumentFile
from doctr.models import ocr_predictor
import ollama
from io import BytesIO
import base64
from PIL import Image
import pandas as pd
import difflib
from sqlalchemy import select
from sqlalchemy.orm import Session

from base import engine, Marca, Modelo

# ==========================================
# 1. CONFIGURACIÓN E INFRAESTRUCTURA
# ==========================================
PATH_POPPLER = os.environ.get("POPPLER_PATH", r"C:\poppler-26.07.0\Library\bin")
pytesseract.pytesseract.tesseract_cmd = os.environ.get(
    "TESSERACT_CMD",
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
)

CARPETA_REVISION = "revision_manual"
CARPETA_EXITOSOS = "procesados_exito"
ARCHIVO_EXCEL = "reporte_extracciones.xlsx"
ARCHIVO_LOG = "log_revision_fallos.txt"

os.makedirs(CARPETA_REVISION, exist_ok=True)
os.makedirs(CARPETA_EXITOSOS, exist_ok=True)

modelo_doctr = None

PLACEHOLDERS = {
    "valor", "value", "marca", "modelo", "none", "null", "n/a", "na", "-",
    "marca del vehiculo", "marca del vehículo", "modelo del vehiculo",
    "modelo del vehículo", "nombre del cliente", "nombre del cliente:",
    "c.i.", "c.i", "ci", "ruc", "cedula", "cédula",
}

COLORES_NO_MARCA = {
    "blanco", "negro", "rojo", "azul", "gris", "plata", "plateado", "verde",
    "amarillo", "cafe", "café", "beige", "naranja", "vino", "perla", "ivory",
    "marfil", "dorado", "champagne", "grafito",
}

SUFIJOS_TECNICOS = {
    "AC", "TA", "TM", "EV", "BEV", "PHEV", "HEV", "CN", "CD", "CS",
    "FWD", "AWD", "RWD", "QUATT", "QUATTRO", "4X2", "4X4", "6X4",
    "SUV", "HATCHBACK", "SEDAN", "SEDÁN", "PICKUP", "VAN", "MPV",
    "AUTOMOVIL", "AUTOMÓVIL", "VEHICULO", "VEHÍCULO", "ELECTRIC",
    "ELECTRICO", "ELÉCTRICO", "MOTORS", "MOTOR", "AUTO",
}

CAMPOS_EXTRACCION = ("marca", "modelo", "cantidad", "beneficiario", "ci", "costo_mas_alto", "costo_total")


def get_modelo_doctr():
    global modelo_doctr
    if modelo_doctr is None:
        print("Cargando modelo neuronal DocTR (esto tomará unos segundos)...")
        modelo_doctr = ocr_predictor(det_arch='db_resnet50', reco_arch='crnn_vgg16_bn', pretrained=True)
    return modelo_doctr


# ==========================================
# 2. VALIDACIONES Y BÚSQUEDA FUZZY
# ==========================================
def _es_placeholder(val):
    if val is None:
        return True
    s = str(val).strip().lower().rstrip(".:")
    s_norm = re.sub(r"\s+", " ", s)
    return (not s_norm) or s_norm in PLACEHOLDERS


def _tokens_vehiculo(texto):
    texto = str(texto or "").upper()
    texto = re.sub(r"\([^)]*\)", " ", texto)
    texto = re.sub(r"[^A-Z0-9ÁÉÍÓÚÑ#+\- ]", " ", texto)
    tokens = []
    for t in texto.split():
        if t in SUFIJOS_TECNICOS:
            continue
        if re.fullmatch(r"\d+P", t) or re.fullmatch(r"\d+KW", t) or re.fullmatch(r"\d+KM", t):
            continue
        if re.fullmatch(r"\d+%", t):
            continue
        tokens.append(t)
    return tokens


def _similitud_nombres(tokens_a, tokens_b):
    if not tokens_a or not tokens_b:
        return 0.0
    sa, sb = " ".join(tokens_a), " ".join(tokens_b)
    if sa == sb:
        return 1.0
    set_a, set_b = set(tokens_a), set(tokens_b)
    shorter, longer = (set_a, set_b) if len(set_a) <= len(set_b) else (set_b, set_a)
    if shorter and shorter.issubset(longer):
        return 0.90 + 0.10 * (len(shorter) / max(len(longer), 1))
    return difflib.SequenceMatcher(None, sa, sb).ratio()


def buscar_vehiculo_fuzzy(marca_buscada, modelo_buscado, umbral_total=0.80):
    if _es_placeholder(marca_buscada) and _es_placeholder(modelo_buscado):
        return None, None

    marca_txt = "" if _es_placeholder(marca_buscada) else str(marca_buscada).strip()
    modelo_txt = "" if _es_placeholder(modelo_buscado) else str(modelo_buscado).strip()
    tokens_marca = _tokens_vehiculo(marca_txt)
    tokens_modelo = _tokens_vehiculo(modelo_txt)
    tokens_combo = _tokens_vehiculo(f"{marca_txt} {modelo_txt}".strip())

    with Session(engine) as session:
        resultados = session.execute(select(Marca.nombre, Modelo.nombre).join(Marca.modelos)).all()

    if not resultados:
        return None, None

    mejor_similitud = 0.0
    mejor_marca, mejor_modelo = None, None

    for m_bd, mod_bd in resultados:
        tokens_marca_bd = _tokens_vehiculo(m_bd)
        tokens_mod_bd = _tokens_vehiculo(mod_bd)

        sim_marca = _similitud_nombres(tokens_marca, tokens_marca_bd) if tokens_marca else 0.55
        sim_mod = _similitud_nombres(tokens_modelo, tokens_mod_bd) if tokens_modelo else 0.0

        # A veces el modelo extraído incluye la marca, o el catálogo trae sufijos largos.
        sim_combo = _similitud_nombres(tokens_combo, tokens_marca_bd + tokens_mod_bd)
        sim_mod = max(sim_mod, _similitud_nombres(tokens_combo, tokens_mod_bd), sim_combo * 0.98)

        if sim_mod < 0.72:
            continue

        sim_total = (sim_marca * 0.4) + (sim_mod * 0.6)
        if sim_combo > sim_total:
            sim_total = sim_combo

        if sim_total > mejor_similitud:
            mejor_similitud = sim_total
            mejor_marca = m_bd
            mejor_modelo = mod_bd

    if mejor_similitud >= umbral_total:
        return mejor_marca, mejor_modelo
    return None, None


def parse_costo(val):
    if val is None or val == "":
        return 0.0

    if isinstance(val, bool):
        return 0.0

    if isinstance(val, (int, float)):
        num = float(val)
    else:
        s = str(val).strip().upper().replace("USD", "").replace("US$", "").replace("$", "")
        s = s.replace(" ", "").replace("\u00a0", "")
        if not s or _es_placeholder(s):
            return 0.0
        if re.fullmatch(r"\d{1,3}(\.\d{3})+,\d{1,2}", s):
            s = s.replace(".", "").replace(",", ".")
        elif re.fullmatch(r"\d{1,3}(,\d{3})+\.\d{1,2}", s):
            s = s.replace(",", "")
        elif re.fullmatch(r"\d+,\d{1,2}", s):
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
        try:
            num = float(s)
        except (ValueError, TypeError):
            return 0.0

    if num < 0:
        return 0.0

    # OCR/LLM a veces concatena dígitos (680336567 en vez de 68033.66).
    if num >= 1_000_000:
        tmp = num
        while tmp > 500_000:
            tmp /= 10.0
        if 500 <= tmp <= 500_000:
            num = tmp

    return round(num, 2)


def limpiar_ci(ci_original):
    raw = str(ci_original or "").strip()
    if _es_placeholder(raw):
        return ""
    compacto = re.sub(r"[.\-\s]", "", raw)
    digitos = re.sub(r"\D", "", compacto)
    if not digitos:
        grupos = re.findall(r"\d{8,13}", raw)
        digitos = max(grupos, key=len) if grupos else ""
    return digitos


def normalizar_datos_llm(datos):
    if not isinstance(datos, dict):
        return {"error": "La IA no devolvió un objeto JSON"}

    if "error" in datos and len(datos) <= 2:
        return datos

    limpio = dict(datos)
    for campo in CAMPOS_EXTRACCION:
        limpio.setdefault(campo, None)

    for campo in ("marca", "modelo", "beneficiario", "ci"):
        val = limpio.get(campo)
        if val is None:
            continue
        texto = str(val).strip()
        limpio[campo] = None if _es_placeholder(texto) else texto

    try:
        cantidad = int(float(str(limpio.get("cantidad") or 1).replace(",", ".")))
        limpio["cantidad"] = cantidad if cantidad > 0 else 1
    except (ValueError, TypeError):
        limpio["cantidad"] = 1

    limpio["costo_mas_alto"] = parse_costo(limpio.get("costo_mas_alto"))
    limpio["costo_total"] = parse_costo(limpio.get("costo_total"))
    return limpio


def evaluar_extraccion(datos):
    errores = []
    if not isinstance(datos, dict):
        return False, "La IA no devolvió un objeto JSON"

    normalizado = normalizar_datos_llm(datos)
    if normalizado is not datos:
        datos.clear()
        datos.update(normalizado)

    if "error" in datos and not any(datos.get(c) for c in CAMPOS_EXTRACCION if c != "cantidad"):
        return False, f"Error de IA: {datos['error']}"

    costo_total = parse_costo(datos.get("costo_total"))
    costo_mas_alto = parse_costo(datos.get("costo_mas_alto"))
    datos["costo_total"] = costo_total
    datos["costo_mas_alto"] = costo_mas_alto

    if costo_total == 0.0:
        errores.append("Faltan costos")

    if costo_mas_alto == 0.0:
        datos["costo_mas_alto"] = costo_total
        costo_mas_alto = costo_total

    if costo_mas_alto > costo_total > 0:
        datos["costo_mas_alto"] = costo_total

    ci_original = str(datos.get("ci") or "").strip()
    ci_limpio = limpiar_ci(ci_original)

    if not ci_limpio:
        errores.append("Falta CI/RUC")
    elif not ci_limpio.isdigit():
        errores.append(f"CI/RUC inválido (contiene letras o símbolos: '{ci_original}')")
    elif len(ci_limpio) < 8:
        errores.append(f"CI/RUC demasiado corto ('{ci_limpio}')")
    else:
        datos["ci"] = ci_limpio

    beneficiario = str(datos.get("beneficiario") or "").strip()
    if _es_placeholder(beneficiario):
        errores.append("Falta beneficiario")
        datos["beneficiario"] = None
    elif len(beneficiario) < 4 or not re.search(r"[a-zA-ZáéíóúÁÉÍÓÚñÑ]", beneficiario):
        errores.append(f"Beneficiario ilógico o mal formado ('{beneficiario}')")
    else:
        datos["beneficiario"] = beneficiario

    marca = datos.get("marca")
    modelo = datos.get("modelo")
    if _es_placeholder(marca):
        datos["marca"] = None
        marca = None
    if _es_placeholder(modelo):
        datos["modelo"] = None
        modelo = None

    if marca and str(marca).strip().lower() in COLORES_NO_MARCA:
        errores.append(f"La marca extraída parece un color ('{marca}')")

    marca_base, modelo_base = buscar_vehiculo_fuzzy(marca, modelo)
    datos["marca_base"] = marca_base if marca_base else "NO ENCONTRADO"
    datos["modelo_base"] = modelo_base if modelo_base else "NO ENCONTRADO"

    if not marca_base or not modelo_base:
        errores.append(f"Vehículo no existe en BD o similitud baja ({marca or '-'} {modelo or '-'})")

    if errores:
        return False, " | ".join(errores)

    return True, "Validación Exitosa"


# ==========================================
# 3. EXTRACCIÓN Y LLMS
# ==========================================
PROMPT_CAMPOS = """
Extrae los datos de esta factura/proforma de vehículos. Devuelve SOLO un JSON válido con exactamente estas claves:
- marca: marca comercial del vehículo (nunca un color, nunca un marcador de plantilla)
- modelo: modelo comercial del vehículo
- cantidad: entero
- beneficiario: nombre completo del cliente/comprador
- ci: cédula o RUC, solo dígitos (sin puntos, guiones ni etiquetas)
- costo_mas_alto: número float del ítem vehicular más caro
- costo_total: número float del total a pagar

REGLAS:
1. Si un dato no aparece, usa null (o 0.0 en costos). NUNCA inventes ni copies palabras como "valor".
2. Los costos van en dólares, con punto decimal y SIN separadores de miles. Ejemplo: 38500.00
3. No uses el color del vehículo como marca.
4. No incluyas markdown ni texto fuera del JSON.
""".strip()


def es_pagina_nativa(pdf_page):
    texto = pdf_page.get_text("text").strip()
    return len(texto) > 50, texto


def extraer_texto_doctr(imagen_pil):
    img_cv = cv2.cvtColor(np.array(imagen_pil), cv2.COLOR_RGB2BGR)
    gris = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
    binarizada = cv2.adaptiveThreshold(
        cv2.GaussianBlur(gris, (5, 5), 0),
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        11,
        2,
    )
    _exito, buffer = cv2.imencode('.png', binarizada)
    res = get_modelo_doctr()(DocumentFile.from_images(buffer.tobytes()))
    texto = ""
    for page in res.pages:
        for block in page.blocks:
            texto += " ".join([w.value for line in block.lines for w in line.words]) + "\n"
    return texto.strip()


def extraer_json_de_texto(txt):
    txt = (txt or "").strip()
    if txt.startswith("```"):
        txt = re.sub(r"^```(?:json)?\n?(.*?)\n?```$", r"\1", txt, flags=re.DOTALL)
    match = re.search(r"\{.*\}", txt, re.DOTALL)
    if not match:
        raise ValueError("No se encontraron llaves {} de JSON.")
    return json.loads(match.group(0))


def estructurar_con_qwen(texto, intento_nombre="Qwen2.5"):
    prompt = f"{PROMPT_CAMPOS}\n\nTEXTO:\n{texto}"
    try:
        res = ollama.chat(
            model='qwen2.5:3b',
            format='json',
            messages=[{'role': 'user', 'content': prompt}],
            options={'temperature': 0.0},
        )
        datos = extraer_json_de_texto(res['message']['content'])
        datos = normalizar_datos_llm(datos)
        print(f"\n--- RESPUESTA {intento_nombre} ---")
        print(json.dumps(datos, indent=2, ensure_ascii=False, default=str))
        return datos
    except Exception as e:
        return {"error": str(e)}


def extraer_con_minicpm(imagenes_pil, max_reintentos=3):
    img_b64 = []
    for img in imagenes_pil[:2]:
        img_resized = img.copy()
        img_resized.thumbnail((1024, 1024))
        buf = BytesIO()
        img_resized.save(buf, format="JPEG", quality=75)
        img_b64.append(base64.b64encode(buf.getvalue()).decode("utf-8"))

    prompt = PROMPT_CAMPOS + "\nAnaliza la imagen. El campo ci debe ir entre comillas y contener solo números."
    for intento in range(max_reintentos):
        txt = ""
        try:
            res = ollama.chat(
                model='minicpm-v',
                messages=[{'role': 'user', 'content': prompt, 'images': img_b64}],
                options={'temperature': 0.1},
            )
            txt = res['message']['content'].strip()
            datos = normalizar_datos_llm(extraer_json_de_texto(txt))
            print(f"\n--- RESPUESTA MINICPM-V (Intento {intento+1}) ---")
            print(json.dumps(datos, indent=2, ensure_ascii=False, default=str))
            return datos
        except Exception as e:
            print(f"\n[⚠️ Error MiniCPM-V {intento+1}/{max_reintentos}]: {str(e)}")
            print(f"Texto sin procesar recibido (para depuración):\n{txt[:300]}")
            time.sleep(3)

    return {"error": "MiniCPM-V falló en todos los reintentos."}


def consolidar_extracciones(datos_1, datos_2, motivo_fallo):
    modelo_juez = 'llama3.1'
    prompt = f"""
Eres un auditor combinando dos JSON extraídos de la misma factura.
El Intento 1 falló por: "{motivo_fallo}".

Intento 1: {json.dumps(datos_1, ensure_ascii=False, default=str)}
Intento 2: {json.dumps(datos_2, ensure_ascii=False, default=str)}

Elige los valores más lógicos. ci solo dígitos. costos numéricos sin miles.
NUNCA uses "valor" como dato. Si un campo no se puede determinar, usa null o 0.0.

Responde SOLO JSON:
{{
  "analisis": "Tu razonamiento aquí",
  "datos_consolidados": {{
    "marca": null,
    "modelo": null,
    "cantidad": 1,
    "beneficiario": null,
    "ci": null,
    "costo_mas_alto": 0.0,
    "costo_total": 0.0
  }}
}}
"""
    try:
        res = ollama.chat(
            model=modelo_juez,
            format='json',
            messages=[{'role': 'user', 'content': prompt}],
            options={'temperature': 0.1},
        )
        respuesta_completa = extraer_json_de_texto(res['message']['content'])
        print(f"\n--- 🧠 ANÁLISIS DEL JUEZ LLM ({modelo_juez}) ---")
        print(f"Razonamiento: {respuesta_completa.get('analisis', 'Sin análisis detallado')}")

        datos_finales = normalizar_datos_llm(respuesta_completa.get('datos_consolidados', datos_1.copy()))

        for campo in ("costo_total", "costo_mas_alto"):
            if parse_costo(datos_finales.get(campo)) == 0.0:
                if parse_costo(datos_1.get(campo)) > 0:
                    datos_finales[campo] = parse_costo(datos_1.get(campo))
                elif parse_costo(datos_2.get(campo)) > 0:
                    datos_finales[campo] = parse_costo(datos_2.get(campo))

        if "Vehículo" in motivo_fallo or "costo" in motivo_fallo.lower() or "Faltan costos" in motivo_fallo:
            for campo in ("beneficiario", "ci"):
                val_1 = datos_1.get(campo)
                if val_1 and not _es_placeholder(val_1):
                    if _es_placeholder(datos_finales.get(campo)):
                        datos_finales[campo] = val_1

        for campo in CAMPOS_EXTRACCION:
            if _es_placeholder(datos_finales.get(campo)) or datos_finales.get(campo) in (0, 0.0):
                for fuente in (datos_1, datos_2):
                    candidato = fuente.get(campo)
                    if campo in ("costo_total", "costo_mas_alto"):
                        if parse_costo(candidato) > 0:
                            datos_finales[campo] = parse_costo(candidato)
                            break
                    elif candidato and not _es_placeholder(candidato):
                        datos_finales[campo] = candidato
                        break

        datos_finales = normalizar_datos_llm(datos_finales)
        print("Resultado Consolidado Definitivo:")
        print(json.dumps(datos_finales, indent=2, ensure_ascii=False, default=str))
        return datos_finales
    except Exception as e:
        print(f"\n[⚠️ Error en Juez LLM]: {str(e)}")
        return normalizar_datos_llm(datos_1)


def serializar_resultado(datos):
    limpio = {}
    for k, v in (datos or {}).items():
        if k == "paginas_pdf" and isinstance(v, list):
            limpio[k] = [int(x) for x in v]
        elif isinstance(v, (np.integer,)):
            limpio[k] = int(v)
        elif isinstance(v, (np.floating,)):
            limpio[k] = float(v)
        else:
            limpio[k] = v
    return limpio


def nombre_origen_limpio(ruta_archivo):
    nombre = os.path.basename(ruta_archivo)
    return re.sub(r"^[0-9a-f]{32}_", "", nombre, flags=re.IGNORECASE)


def nombre_fragmento(ruta_archivo, idx_doc, paginas, total_docs):
    nombre = nombre_origen_limpio(ruta_archivo)
    base, ext = os.path.splitext(nombre)
    ext = ext or ".pdf"
    if not paginas:
        paginas_txt = "x"
    elif paginas[0] == paginas[-1]:
        paginas_txt = str(paginas[0] + 1)
    else:
        paginas_txt = f"{paginas[0] + 1}-{paginas[-1] + 1}"
    if total_docs <= 1:
        return f"{base}_p{paginas_txt}{ext}"
    return f"{base}_doc{idx_doc + 1}_p{paginas_txt}{ext}"


def _zona_cabecera(texto, n_lineas=14, n_chars=500):
    lineas = [ln.strip() for ln in (texto or "").splitlines() if ln.strip()]
    return "\n".join(lineas[:n_lineas])[:n_chars]


def extraer_paginacion(texto):
    m = re.search(r"(?i)p[aá]gina\s+(\d+)\s*(?:de|/)\s*(\d+)", texto or "")
    if m:
        return int(m.group(1)), int(m.group(2))
    m = re.search(r"(?i)hoja\s+(\d+)\s*(?:de|/)\s*(\d+)", texto or "")
    if m:
        return int(m.group(1)), int(m.group(2))
    return None


def extraer_id_documento(texto):
    t = (texto or "").replace("\u00a0", " ")
    patrones = [
        r"(?i)(?:factura|proforma)[^\n]{0,80}?(?:n[°ºo]|no\.?|n\.°|núm(?:ero)?)[^\dA-Z]{0,12}(\d{3,6}\s*[-–]\s*\d{5,12})",
        r"(?i)(?:factura|proforma)[^\n]{0,80}?(\d{3}-\d{3}-\d{4,9})",
        r"(?i)proforma\s*n[^\n]{0,20}?([A-Z]{1,4}[-/]?\d{4}[-/]\d{2,6}[-/]\d{2,6})",
        r"(?i)(?:factura|proforma)\s*(?:n[°ºo]|no\.?|n\.°)?[^\d]{0,12}(\d{4,8})",
    ]
    for patron in patrones:
        m = re.search(patron, t)
        if m:
            return re.sub(r"\s+", "", m.group(1).replace("–", "-"))
    return None


def parece_continuacion(texto):
    t = texto or ""
    pag = extraer_paginacion(t)
    if pag and pag[0] > 1:
        return True
    if len(t.strip()) < 160:
        return True
    cab = _zona_cabecera(t).lower()
    inicios = (
        "términos y condiciones", "terminos y condiciones",
        "especificaciones técnicas", "especificaciones tecnicas",
        "formas de pago", "observaciones", "notas importantes",
        "firma vendedor", "servicio post-venta", "homologacion",
        "homologación", "garantia", "garantía", "certificacion",
        "certificación", "incluye cargador", "contrato mantenimiento",
        "pedido especial",
    )
    return any(cab.startswith(s) or s in cab[:90] for s in inicios)


def es_portada_documento(texto):
    t = texto or ""
    pag = extraer_paginacion(t)
    if pag and pag[0] > 1:
        return False
    if pag and pag[0] == 1:
        return True
    if parece_continuacion(t) and not extraer_id_documento(_zona_cabecera(t)):
        return False

    cab = _zona_cabecera(t)
    patrones = [
        r"(?i)^\s*factura(\s+electr[oó]nica)?\b",
        r"(?i)^\s*proforma(\s+comercial)?\b",
        r"(?i)dossier y proforma",
        r"(?i)proforma\s*n[°ºo.\s:]",
        r"(?i)factura\s*n[°ºo.\s:]",
        r"(?i)venta de veh[ií]culos.{0,20}proforma",
    ]
    for linea in cab.splitlines()[:12]:
        for patron in patrones:
            if re.search(patron, linea):
                return True
    return extraer_id_documento(cab) is not None


def documento_parece_cerrado(texto):
    t = (texto or "").lower()
    pag = extraer_paginacion(texto or "")
    if pag and pag[0] == pag[1] and pag[1] >= 1:
        return True
    if "firma vendedor" in t:
        return True
    if "términos y condiciones" in t or "terminos y condiciones" in t:
        return True
    if extraer_id_documento(texto or "") and re.search(r"total a pagar|valor total|total factura", t):
        return True
    return False


def _doc_vacio():
    return {"texto": "", "paginas": [], "imagenes": [], "tiene_escaneado": False}


def _agregar_pagina(doc, pagina):
    doc["texto"] += "\n" + (pagina.get("texto") or "")
    doc["paginas"].append(pagina["idx"])
    doc["imagenes"].append(pagina["imagen"])
    if not pagina.get("es_nativo", True):
        doc["tiene_escaneado"] = True


def segmentar_documentos_logicos(paginas_data):
    """Agrupa páginas consecutivas de la misma factura/proforma."""
    docs = []
    actual = None
    id_actual = None
    cerrado = False

    for pagina in paginas_data:
        texto = pagina.get("texto") or ""
        id_pagina = extraer_id_documento(texto[:1800])
        portada = es_portada_documento(texto)
        continuacion = parece_continuacion(texto)
        mismo_id = bool(id_pagina and id_actual and id_pagina == id_actual)
        nuevo_id = bool(id_pagina and id_actual and id_pagina != id_actual)

        if actual is None:
            actual = _doc_vacio()
            _agregar_pagina(actual, pagina)
            id_actual = id_pagina
            cerrado = documento_parece_cerrado(texto)
            continue

        if mismo_id or (continuacion and not nuevo_id):
            _agregar_pagina(actual, pagina)
            if id_pagina:
                id_actual = id_pagina
            cerrado = documento_parece_cerrado(actual["texto"])
            continue

        debe_partir = nuevo_id or (portada and not continuacion) or (
            cerrado and not continuacion and len(texto.strip()) > 200
        )
        if debe_partir:
            docs.append(actual)
            actual = _doc_vacio()
            _agregar_pagina(actual, pagina)
            id_actual = id_pagina
            cerrado = documento_parece_cerrado(texto)
            continue

        _agregar_pagina(actual, pagina)
        if id_pagina and not id_actual:
            id_actual = id_pagina
        cerrado = documento_parece_cerrado(actual["texto"])

    if actual and (actual["texto"].strip() or actual["imagenes"]):
        docs.append(actual)
    return docs


def guardar_documento_logico(ruta_origen, es_pdf, indices_paginas, ruta_destino):
    os.makedirs(os.path.dirname(os.path.abspath(ruta_destino)), exist_ok=True)
    if not es_pdf:
        shutil.copy(ruta_origen, ruta_destino)
        return ruta_destino

    origen = pymupdf.open(ruta_origen)
    try:
        indices = [i for i in indices_paginas if 0 <= i < origen.page_count]
        if not indices:
            shutil.copy(ruta_origen, ruta_destino)
            return ruta_destino
        origen.select(indices)
        origen.save(ruta_destino, garbage=4, deflate=True)
    finally:
        origen.close()
    return ruta_destino


# ==========================================
# 4. PIPELINE CENTRAL
# ==========================================
def procesar_archivo_interno(ruta_archivo):
    print(f"\n[{os.path.basename(ruta_archivo)}] Iniciando análisis...")
    es_pdf = ruta_archivo.lower().endswith('.pdf')

    try:
        poppler = PATH_POPPLER if PATH_POPPLER and os.path.isdir(PATH_POPPLER) else None
        imagenes = convert_from_path(ruta_archivo, poppler_path=poppler) if es_pdf else [Image.open(ruta_archivo).convert('RGB')]
    except Exception as e:
        fallo = {
            "estado": f"REVISIÓN (No se pudo leer el archivo: {e})",
            "archivo": os.path.basename(ruta_archivo),
            "paginas_pdf": [],
            "error": str(e),
        }
        yield serializar_resultado(fallo)
        return

    paginas_data = []

    if es_pdf:
        doc = pymupdf.open(ruta_archivo)
        for i, page in enumerate(doc):
            es_nativo, texto = es_pagina_nativa(page)
            if not es_nativo:
                texto = pytesseract.image_to_string(imagenes[i], lang='spa')
            paginas_data.append({"idx": i, "texto": texto, "es_nativo": es_nativo, "imagen": imagenes[i]})
        doc.close()
    else:
        texto = pytesseract.image_to_string(imagenes[0], lang='spa')
        paginas_data.append({"idx": 0, "texto": texto, "es_nativo": False, "imagen": imagenes[0]})

    docs_logicos = segmentar_documentos_logicos(paginas_data)

    if not docs_logicos:
        docs_logicos = [{
            "texto": "",
            "paginas": [p["idx"] for p in paginas_data],
            "imagenes": [p["imagen"] for p in paginas_data],
            "tiene_escaneado": True,
        }]

    print(f"[INFO] {len(docs_logicos)} documento(s) lógico(s) detectado(s):")
    for i, d in enumerate(docs_logicos, 1):
        pags = [p + 1 for p in d["paginas"]]
        print(f"    - Doc {i}: páginas {pags}")

    resultados_finales = []

    for i, doc in enumerate(docs_logicos):
        print(f"\n--- PROCESANDO DOCUMENTO LÓGICO {i+1}/{len(docs_logicos)} ---")

        texto_analizar = doc["texto"]
        if doc["tiene_escaneado"] or len((doc["texto"] or "").strip()) < 40:
            print("[INFO] Documento sin texto nativo suficiente. Pasando OCR Matemático (DocTR)...")
            texto_ocr = "\n".join([extraer_texto_doctr(img) for img in doc["imagenes"]])
            texto_analizar = (texto_analizar + "\n" + texto_ocr).strip()

        print("[INTENTO 1] Analizando texto estructurado...")
        datos_1 = estructurar_con_qwen(texto_analizar, intento_nombre="INTENTO 1 (Texto Base)")
        es_valido, msj = evaluar_extraccion(datos_1)
        datos_finales = datos_1

        if not es_valido:
            print(f"\n[FALLO INTENTO 1: {msj}]. Desplegando IA Visual (MiniCPM-V)...")
            datos_2 = extraer_con_minicpm(doc["imagenes"])
            print("\n[JUEZ LLM] Mandando a corregir focalizado en los errores detectados...")

            datos_1_limpio = {k: v for k, v in datos_1.items() if k not in ["marca_base", "modelo_base", "error"]}
            datos_2_limpio = {k: v for k, v in datos_2.items() if k not in ["marca_base", "modelo_base", "error"]}

            datos_fusionados = consolidar_extracciones(datos_1_limpio, datos_2_limpio, msj)
            es_valido_fusion, msj_fusion = evaluar_extraccion(datos_fusionados)
            datos_finales = datos_fusionados
            es_valido = es_valido_fusion
            msj = msj_fusion

        datos_finales["estado"] = "ÉXITO" if es_valido else f"REVISIÓN ({msj})"
        datos_finales["archivo"] = nombre_origen_limpio(ruta_archivo)
        datos_finales["paginas_pdf"] = [pg + 1 for pg in doc["paginas"]]
        datos_finales.pop("error", None)

        carpeta_destino = CARPETA_EXITOSOS if es_valido else CARPETA_REVISION
        nombre_frag = nombre_fragmento(ruta_archivo, i, doc["paginas"], len(docs_logicos))
        ruta_frag = os.path.join(carpeta_destino, nombre_frag)
        if os.path.exists(ruta_frag):
            base_frag, ext_frag = os.path.splitext(ruta_frag)
            ruta_frag = f"{base_frag}_{int(time.time())}{ext_frag}"
            nombre_frag = os.path.basename(ruta_frag)
        try:
            guardar_documento_logico(ruta_archivo, es_pdf, doc["paginas"], ruta_frag)
            datos_finales["archivo_guardado"] = nombre_frag
            print(f"[+] Fragmento guardado en {carpeta_destino}/{nombre_frag}")
        except Exception as e:
            datos_finales["archivo_guardado"] = None
            print(f"[⚠️ No se pudo guardar el fragmento {nombre_frag}]: {e}")

        datos_finales = serializar_resultado(datos_finales)
        resultados_finales.append(datos_finales)
        yield datos_finales

        print(
            f"\n✅ RESULTADO FINAL DOC {i+1} (Páginas {datos_finales['paginas_pdf']}): "
            f"CI {datos_finales.get('ci')} | {datos_finales.get('marca')} {datos_finales.get('modelo')} | "
            f"ESTADO: {datos_finales['estado']}"
        )

    resultados_exitosos = [r for r in resultados_finales if "ÉXITO" in str(r.get("estado", ""))]
    resultados_fallidos = [r for r in resultados_finales if "REVISIÓN" in str(r.get("estado", ""))]

    if resultados_exitosos:
        df_exitosos_limpio = [{k: v for k, v in r.items() if k != "paginas_pdf"} for r in resultados_exitosos]
        df_nuevo = pd.DataFrame(df_exitosos_limpio)

        if os.path.exists(ARCHIVO_EXCEL):
            df_existente = pd.read_excel(ARCHIVO_EXCEL)
            df_final = pd.concat([df_existente, df_nuevo], ignore_index=True)
        else:
            df_final = df_nuevo
        df_final.to_excel(ARCHIVO_EXCEL, index=False)
        print(f"[+] {len(resultados_exitosos)} registro(s) exitoso(s) añadido(s) al Excel.")

    if resultados_fallidos:
        with open(ARCHIVO_LOG, "a", encoding="utf-8") as f:
            for r in resultados_fallidos:
                f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Archivo: {r['archivo']} | Páginas PDF: {r.get('paginas_pdf')} | Motivo: {r['estado']}\n")
                f.write(f"Datos extraídos: {json.dumps(r, ensure_ascii=False, default=str)}\n")
                f.write("-" * 80 + "\n")
        print(f"[-] {len(resultados_fallidos)} registro(s) fallido(s) registrado(s) en {ARCHIVO_LOG}.")


if __name__ == "__main__":
    ruta_ejemplo = "docs/merge_facturas.pdf"
    if not os.path.exists(ruta_ejemplo):
        print(f"ERROR: No se encuentra el archivo en {ruta_ejemplo}.")
    else:
        resultados = list(procesar_archivo_interno(ruta_ejemplo))
        print("\n=== RESULTADOS FINALES EN FORMATO JSON ===")
        print(json.dumps(resultados, indent=4, ensure_ascii=False, default=str))
