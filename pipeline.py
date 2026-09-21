import os
import re
import json
import shutil
import time
import numpy as np
import pymupdf
from pdf2image import convert_from_path
import pytesseract
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
_catalogo_vehiculos = None

PLACEHOLDERS = {
    "valor", "value", "marca", "modelo", "none", "null", "n/a", "na", "-",
    "marca del vehiculo", "marca del vehículo", "modelo del vehiculo",
    "modelo del vehículo", "nombre del cliente", "nombre del cliente:",
    "c.i.", "c.i", "ci", "ruc", "cedula", "cédula",
    "nombre/razón", "nombre/razon", "razón social", "razon social",
    "nombre/razón social", "nombre/razon social", "social", "razón", "razon",
    "cliente", "proveedor", "no especificado", "no especifica", "n/e",
    "nombre", "beneficiary", "comprador",
}

COLORES_NO_MARCA = {
    "blanco", "negro", "rojo", "azul", "gris", "plata", "plateado", "verde",
    "amarillo", "cafe", "café", "beige", "naranja", "vino", "perla", "ivory",
    "marfil", "dorado", "champagne", "grafito",
}

_MODELO_GENERICO = {
    "insignia", "referencial", "vehiculo", "vehículo", "electrico", "eléctrico",
    "auto", "camion", "camión", "suv", "sedan", "sedán",
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
        from doctr.io import DocumentFile
        from doctr.models import ocr_predictor
        modelo_doctr = ocr_predictor(det_arch='db_resnet50', reco_arch='crnn_vgg16_bn', pretrained=True)
    return modelo_doctr


def get_catalogo_vehiculos():
    global _catalogo_vehiculos
    if _catalogo_vehiculos is None:
        with Session(engine) as session:
            _catalogo_vehiculos = session.execute(select(Marca.nombre, Modelo.nombre).join(Marca.modelos)).all()
    return _catalogo_vehiculos


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

    resultados = get_catalogo_vehiculos()

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


def parse_costo(val, reparar_concatenado=True):
    if val is None or val == "":
        return 0.0
    if isinstance(val, bool):
        return 0.0

    tiene_separadores = False
    if isinstance(val, (int, float)):
        num = float(val)
    else:
        s = str(val).strip().upper().replace("USD", "").replace("US$", "").replace("CLP", "").replace("$", "")
        s = s.replace(" ", "").replace("\u00a0", "")
        if re.search(r"\d:\d{2}$", s):
            s = s.replace(":", ".")
        if not s or _es_placeholder(s):
            return 0.0
        if re.fullmatch(r"\d{1,3}(\.\d{3})+,\d{1,2}", s):
            tiene_separadores = True
            s = s.replace(".", "").replace(",", ".")
        elif re.fullmatch(r"\d{1,3}(,\d{3})+\.\d{1,2}", s):
            tiene_separadores = True
            s = s.replace(",", "")
        elif re.fullmatch(r"\d{1,3}(,\d{3})+\.\d{3}", s):
            tiene_separadores = True
            s = re.sub(r"[.,]", "", s)
        elif re.fullmatch(r"\d{1,3}(\.\d{3})+", s):
            tiene_separadores = True
            s = s.replace(".", "")
        elif re.fullmatch(r"\d{1,3}(,\d{3})+", s):
            tiene_separadores = True
            s = s.replace(",", "")
        elif re.fullmatch(r"\d+,\d{1,2}", s):
            s = s.replace(",", ".")
        try:
            num = float(s)
        except (ValueError, TypeError):
            return 0.0

    if num < 0:
        return 0.0

    # Solo reparar concatenaciones de OCR/LLM si no había separadores de miles reales.
    if reparar_concatenado and not tiene_separadores and num >= 1_000_000 and num == int(num):
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
    if re.search(r"\d+[.,]\d{2}\s*$", raw) and parse_costo(raw, reparar_concatenado=False) >= 100:
        return ""
    compacto = re.sub(r"[.\-\s]", "", raw)
    digitos = re.sub(r"\D", "", compacto)
    if not digitos:
        grupos = re.findall(r"\d{8,13}", raw)
        digitos = max(grupos, key=len) if grupos else ""
    return digitos


def _parece_identificador(texto):
    s = str(texto or "").strip()
    if not s:
        return False
    if re.fullmatch(r"[\d.\-\sKk]+", s) and len(re.sub(r"\D", "", s)) >= 7:
        return True
    return False


_RE_ETIQUETA_ID = re.compile(
    r"(?i)\b(?:rut|ruc|c\.?\s*i\.?|c[eé]dula|nit|dni)(?:\s*/\s*(?:rut|ruc|c\.?\s*i\.?|c[eé]dula|nit|dni))?\b"
)
_RE_ID_EN_LINEA = re.compile(
    r"(?i)(?:rut|ruc|c\.?\s*i\.?|c[eé]dula|nit|dni)(?:\s*/\s*(?:rut|ruc|c\.?\s*i\.?|c[eé]dula|nit|dni))?\s*:?\s*"
    r"(\d{1,2}\.?\d{3}\.?\d{3}-[\dkK]|\d{8,13})"
)
_RE_ID_SOLO_ETIQUETA = re.compile(
    r"(?i)(?:rut|ruc|c\.?\s*i\.?|c[eé]dula|nit|dni)(?:\s*/\s*(?:rut|ruc|c\.?\s*i\.?|c[eé]dula|nit|dni))?\s*:?\s*$"
)
_RE_ID_PARENTESIS = re.compile(r"(?i)\((?:c[eé]dula|ruc|c\.?i\.?|rut|nit|dni)\s*:?\s*(\d{8,13})\)")
_RE_RUT_CHILE = re.compile(r"\b(\d{1,2}\.\d{3}\.\d{3}-[\dkK])\b")
_RE_PARTES_JUNTAS = re.compile(
    r"(?i)((?:datos del\s+)?(?:vendedor|emisor|proveedor|seller|vendor)|"
    r"(?:datos del\s+)?(?:comprador|cliente|adquiriente|adquirente|receptor|buyer))"
    r"\W{0,40}"
    r"((?:datos del\s+)?(?:vendedor|emisor|proveedor|seller|vendor)|"
    r"(?:datos del\s+)?(?:comprador|cliente|adquiriente|adquirente|receptor|buyer))"
)

_PATS_VENDEDOR = (
    (r"(?i)datos del (?:vendedor|emisor|proveedor)", 5),
    (r"(?i)\bemisor\s*/\s*vendedor\b", 5),
    (r"(?i)\b(?:vendedor|proveedor|emisor)\s*:", 4),
    (r"(?i)\b(?:seller|vendor|issuer)\b", 3),
    (r"(?i)\b(?:vendedor|proveedor|emisor)\b", 2),
    (r"(?i)autorizaci[oó]n.{0,24}sri", 2),
)
_PATS_COMPRADOR = (
    (r"(?i)datos del (?:comprador|cliente|adquiriente|adquirente)", 5),
    (r"(?i)adquiriente\s*/\s*cliente", 5),
    (r"(?i)preparado para\s*:", 5),
    (r"(?i)\b(?:comprador|adquiriente|adquirente|receptor)\s*:", 4),
    (r"(?i)\bcliente\s*:", 4),
    (r"(?i)se[nñ]or(?:es|\(es\))?\s*:", 4),
    (r"(?i)facturar a|bill to|sold to", 4),
    (r"(?i)\b(?:comprador|adquiriente|adquirente|receptor|buyer)\b", 2),
)
_RE_NOMBRE_COMP = re.compile(
    r"(?i)(?:nombre(?:/raz[oó]n)?(?:\s*social)?|raz[oó]n\s*social|cliente|cient[ea]|comprador|adquiriente|"
    r"se[nñ]or(?:es|\(es\))?|preparado para|facturar a|bill to|sold to|receptor)\s*:\s*([^\n]*)"
)
_RE_ETIQUETA_CAMPO = re.compile(
    r"(?i)^(nombre|raz[oó]n|social|cliente|comprador|vendedor|se[nñ]or(?:\(es\)|es)?|"
    r"tel[eé]fono|celular|correo|email|direcci[oó]n|ciudad|giro|ruc|rut|c\.?i\.?|"
    r"c[eé]dula|fecha|validez|asesor|marca|modelo|a[nñ]o|color|cantidad|total|"
    r"subtotal|iva|p[aá]gina|observaciones|empresa|monto|concepto)s?$"
)
_RE_GIRO = re.compile(
    r"(?i)^(?:venta de\b|transporte de\b|transporte terrestre\b|servicios de\b)"
)
_RE_RUIDO_NOMBRE = re.compile(
    r"(?i)^(?:observaciones|monto|concepto|detalle|patente|color|tipo|firma)\b"
)


def _es_nombre_empresa(nombre):
    return bool(re.search(
        r"(?i)\b(?:spa|s\.?\s*a\.?|ltda|limitada|eirl|s\.?a\.?s\.?|cia\.?|compañía|compania)\b",
        str(nombre or ""),
    ))


def _es_relleno(texto):
    s = str(texto or "").strip()
    return bool(s) and bool(re.fullmatch(r"[_—\-.\s]+", s))


def _nombre_ok(nom):
    if not nom or _es_placeholder(nom) or _es_relleno(nom) or _parece_identificador(nom):
        return False
    s = re.sub(r"\s+", " ", str(nom)).strip(" :-|")
    if len(s) < 4 or len(s) > 90 or not re.search(r"[a-zA-ZáéíóúÁÉÍÓÚñÑ]", s):
        return False
    if len(s.split()) > 12:
        return False
    if _RE_ETIQUETA_CAMPO.match(s) or _RE_GIRO.search(s) or _RE_RUIDO_NOMBRE.search(s):
        return False
    if re.search(
        r"(?i)^datos del\b|^detalle\b|documento v[aá]lido|documento generado|"
        r"modelo de proforma|sin firma|los datos del|referencial",
        s,
    ):
        return False
    if s.count("_") >= 5:
        return False
    if re.search(r"(?i)\b(?:av\.?|pasaje|calle)\b", s) and re.search(r"\d", s):
        return False
    if re.search(r"(?i)^(?:motor|chasis|patente)\s*:", s):
        return False
    if re.search(r"(?i)\b(?:vendedor|proveedor|emisor|detalle|concepto|precio|total|iva|neto)\b", s):
        return False
    if "$" in s or re.search(r"\d{1,3}([.,]\d{3}){2,}", s):
        return False
    return True


def _score_nombre(nom):
    s = str(nom or "")
    letras = len(re.findall(r"[A-Za-záéíóúñÁÉÍÓÚÑ]", s))
    junk = len(re.findall(r"[^A-Za-záéíóúñÁÉÍÓÚÑ\s.\-']", s))
    return letras - junk * 4 + len(s.split())


def _mejor_nombre(a, b):
    ok_a, ok_b = _nombre_ok(a), _nombre_ok(b)
    if ok_a and not ok_b:
        return a
    if ok_b and not ok_a:
        return b
    if not ok_a:
        return None
    return a if _score_nombre(a) >= _score_nombre(b) else b


def _id_descartable(valor, linea=""):
    d = limpiar_ci(valor)
    if len(d) < 8 or len(d) > 13:
        return True
    if _es_relleno(linea):
        return True
    if re.fullmatch(r"0*9{6,}0*1?", d):
        return True
    return False


def _recoger_ids(texto):
    lineas = (texto or "").splitlines()
    hallados = []
    vistos = set()
    offset = 0

    def add(linea_idx, valor, linea_txt, pos):
        if _id_descartable(valor, linea_txt):
            return
        key = (linea_idx, limpiar_ci(valor))
        if key in vistos:
            return
        vistos.add(key)
        hallados.append({
            "linea": linea_idx,
            "valor": valor,
            "texto_linea": linea_txt,
            "pos": pos,
        })

    for i, linea in enumerate(lineas):
        for m in _RE_ID_EN_LINEA.finditer(linea):
            add(i, m.group(1), linea, offset + m.start())
        if _RE_ID_SOLO_ETIQUETA.search(linea.strip()):
            for j in range(i + 1, min(i + 3, len(lineas))):
                nxt = lineas[j].strip()
                if not nxt:
                    continue
                m = re.match(r"^(\d{1,2}\.?\d{3}\.?\d{3}-[\dkK]|\d{8,13})\b", nxt)
                if m:
                    add(j, m.group(1), lineas[j], offset + len(linea) + 1)
                break
        for m in _RE_ID_PARENTESIS.finditer(linea):
            add(i, m.group(1), linea, offset + m.start())
        for m in _RE_RUT_CHILE.finditer(linea):
            add(i, m.group(1), linea, offset + m.start())
        offset += len(linea) + 1
    return hallados


def _ventana_id(texto, item, radio=420):
    pos = item.get("pos") or 0
    ini = max(0, pos - radio)
    return ((texto or "")[ini:pos + 80] + "\n" + (item.get("texto_linea") or "")).lower()


def _puntaje_rol(ctx, linea):
    sv = sc = 0
    for pat, w in _PATS_VENDEDOR:
        if re.search(pat, ctx):
            sv += w
    for pat, w in _PATS_COMPRADOR:
        if re.search(pat, ctx):
            sc += w
    linea = linea or ""
    if re.search(r"(?i)[•·]\s*ruc\b", linea) or (_es_nombre_empresa(linea) and _RE_ETIQUETA_ID.search(linea)):
        sv += 4
    if re.search(r"(?i)p[aá]gina\s+\d+", linea) and len(limpiar_ci(linea)) >= 10:
        sv += 2
    return sv, sc


def _lado_parte(texto):
    m = _RE_PARTES_JUNTAS.search((texto or "")[:3000])
    if not m:
        return None

    def kind(s):
        s = (s or "").lower()
        if re.search(r"comprador|cliente|adquir|receptor|buyer", s):
            return "comprador"
        if re.search(r"vendedor|emisor|proveedor|seller|vendor", s):
            return "vendedor"
        return None

    a, b = kind(m.group(1)), kind(m.group(2))
    if a and b and a != b:
        return a, b
    return None


def extraer_ids_vendedor_comprador(texto):
    """Separa emisor vs comprador por rol, no por plantilla de factura."""
    t = texto or ""
    hallados = _recoger_ids(t)
    vendedor = comprador = None
    if not hallados:
        return {"vendedor": None, "comprador": None}

    lados = _lado_parte(t)
    if lados:
        por_linea = {}
        for item in hallados:
            por_linea.setdefault(item["linea"], []).append(item["valor"])
        pares = []
        doble = next((vals for vals in por_linea.values() if len(vals) >= 2), None)
        if doble:
            pares = doble[:2]
        else:
            vistos = []
            for item in hallados:
                if re.search(r"(?i)\b(?:total|iva|precio)\b", item["texto_linea"]):
                    continue
                cid = limpiar_ci(item["valor"])
                if vistos and limpiar_ci(vistos[-1]) == cid:
                    continue
                vistos.append(item["valor"])
                if len(vistos) >= 2:
                    break
            pares = vistos[:2]
        if len(pares) >= 2:
            if lados[0] == "vendedor":
                vendedor, comprador = pares[0], pares[1]
            else:
                comprador, vendedor = pares[0], pares[1]
            if limpiar_ci(vendedor) != limpiar_ci(comprador):
                return {"vendedor": vendedor, "comprador": comprador}
            comprador = None

    by_id = {}
    for item in hallados:
        ctx = _ventana_id(t, item)
        sv, sc = _puntaje_rol(ctx, item["texto_linea"])
        cid = limpiar_ci(item["valor"])
        rec = by_id.setdefault(cid, {"valor": item["valor"], "sv": 0, "sc": 0})
        rec["sv"] = max(rec["sv"], sv)
        rec["sc"] = max(rec["sc"], sc)

    if not by_id:
        return {"vendedor": vendedor, "comprador": comprador}

    items = list(by_id.items())
    best_v = max(items, key=lambda kv: kv[1]["sv"] - kv[1]["sc"])
    best_c = max(items, key=lambda kv: kv[1]["sc"] - kv[1]["sv"])
    if best_v[1]["sv"] > best_v[1]["sc"] and best_v[1]["sv"] >= 2:
        vendedor = best_v[1]["valor"]
    if best_c[1]["sc"] > best_c[1]["sv"] and best_c[1]["sc"] >= 2:
        comprador = best_c[1]["valor"]

    vend_id = limpiar_ci(vendedor)
    comp_id = limpiar_ci(comprador)
    otros = [k for k in by_id if k != vend_id]
    if vend_id and not comp_id and len(otros) == 1:
        comprador = by_id[otros[0]]["valor"]
        comp_id = otros[0]
    otros_c = [k for k in by_id if k != comp_id]
    if comp_id and not vend_id and len(otros_c) == 1:
        vendedor = by_id[otros_c[0]]["valor"]
        vend_id = otros_c[0]

    if vend_id and comp_id and vend_id == comp_id:
        comprador = None
    return {"vendedor": vendedor, "comprador": comprador}


def _limpiar_nombre_extraido(nom):
    nom = str(nom or "")
    nom = re.sub(
        r"\s*\(?\s*(?:c[eé]dula|\bc\.?\s*i\.?|\bruc\b|\brut\b)\s*:?\s*[\d.\-kK]+\)?",
        " ",
        nom,
        flags=re.I,
    )
    nom = re.split(
        r"\s{2,}|\bRUT\b|\bRUC\b|\bC\.I\b|\bDirecci|"
        r"\b(?:asesor|vendedor|proveedor|emisor|cargo)\s*:",
        nom,
        maxsplit=1,
        flags=re.I,
    )[0]
    nom = re.sub(r"\s+", " ", nom).strip(" :-|()")
    return re.sub(r"[\s:.,;|/]+$", "", nom)


def extraer_nombre_comprador(texto):
    t = texto or ""

    for m_lab in _RE_NOMBRE_COMP.finditer(t):
        etiqueta = t[max(0, m_lab.start() - 24): m_lab.start()].lower()
        if re.search(r"vendedor|emisor|proveedor", etiqueta):
            continue
        nom = _limpiar_nombre_extraido(m_lab.group(1))
        if not nom:
            after = t[m_lab.end():]
            for line in after.splitlines()[:4]:
                cand = _limpiar_nombre_extraido(line)
                if _nombre_ok(cand):
                    nom = cand
                    break
        if _nombre_ok(nom):
            return nom

    lados = _lado_parte(t)
    if lados:
        m_blk = _RE_PARTES_JUNTAS.search(t[:3000])
        bloque = t[m_blk.end(): m_blk.end() + 600] if m_blk else ""
        lineas_blk = [ln.strip() for ln in bloque.splitlines() if ln.strip()]
        linea0 = re.sub(r"\s+", " ", lineas_blk[0]) if lineas_blk else ""
        m_emp = re.search(
            r"(?i)^(.+?\b(?:spa|s\.?\s*a\.?|ltda|limitada|eirl|s\.?a\.?s\.?))\s+(.+)$",
            linea0,
        )
        if m_emp:
            lado_der = m_emp.group(2).strip(" |-")
            if _nombre_ok(lado_der):
                return lado_der
        nombres = []
        for line in bloque.splitlines():
            line = line.strip()
            if not line or _RE_ETIQUETA_ID.search(line):
                continue
            if re.search(r"(?i)^chile\s*$", line):
                continue
            if re.search(r"(?i)giro|direcci|detalle|marca\s*:|patente|pasaje|\bcalle\b|\bav\.?\b|documento v[aá]lido|motor\s*:", line):
                continue
            if re.search(r"(?i)\bchile\b", line) and re.search(r"\d{2,}", line):
                continue
            cand = _limpiar_nombre_extraido(line)
            if _nombre_ok(cand):
                nombres.append(cand)
        if len(nombres) >= 2:
            return nombres[1]

    m_sec = re.search(
        r"(?i)(?:datos del comprador|datos del cliente|adquiriente|preparado para|receptor|facturar a)[^\n]*\s*([\s\S]{0,450})",
        t,
    )
    if m_sec:
        bloque = re.split(
            r"(?i)datos del vendedor|datos del veh[ií]culo|emisor\b|detalle del|veh[ií]culo adquirido|asesor",
            m_sec.group(1),
            maxsplit=1,
        )[0]
        for line in bloque.splitlines():
            cand = _limpiar_nombre_extraido(line)
            if _nombre_ok(cand):
                return cand
    return None


def _valor_etiquetado(texto, etiqueta):
    m = re.search(rf"(?i){etiqueta}\s*:\s*([^\n]+)", texto or "")
    if m and m.group(1).strip() and not _es_relleno(m.group(1)):
        return m.group(1).strip()
    m = re.search(rf"(?i){etiqueta}\s*:\s*\n+\s*([^\n]+)", texto or "")
    if m and m.group(1).strip() and not _es_relleno(m.group(1)):
        return m.group(1).strip()
    return None


def _contexto_no_precio(texto, match):
    after = (texto or "")[match.end(): match.end() + 16].lower()
    before = (texto or "")[max(0, match.start() - 40): match.start()].lower()
    if re.search(r"^\s*(km|kwh|kw|hp|min|ah|cc)\b", after):
        return True
    if re.search(r"(autonom|potencia|bater|tel[eé]fono|p[aá]gina|a[nñ]o modelo)", before):
        return True
    return False


def _parsear_monto_en_match(texto, match, bruto, tiene_dollar):
    if _contexto_no_precio(texto, match):
        return 0.0
    val = parse_costo(bruto, reparar_concatenado=False)
    if val < 100:
        return 0.0
    if 1990 <= val <= 2035 and not tiene_dollar and not re.search(r"[.,]", bruto):
        return 0.0
    if val < 1000 and not tiene_dollar and not re.search(r"[.,]", bruto):
        return 0.0
    return val


def _monto_tras_etiqueta(texto, etiquetas, tomar="primero"):
    t = texto or ""
    hallados = []
    for etq in etiquetas:
        for mlab in re.finditer(rf"(?i){etq}", t):
            ventana = t[mlab.end(): mlab.end() + 220]
            locales = []
            for m in re.finditer(r"(?:US\$|USD|CLP|\$)\s*([\d.,:]+)", ventana, flags=re.I):
                val = _parsear_monto_en_match(ventana, m, m.group(1), True)
                if val:
                    locales.append(val)
            if not locales:
                for m in re.finditer(r"([\d.,:]+)", ventana):
                    val = _parsear_monto_en_match(ventana, m, m.group(1), False)
                    if val:
                        locales.append(val)
            if locales:
                hallados.append(locales[0])
        if hallados:
            break
    if not hallados:
        return 0.0
    return hallados[-1] if tomar == "ultimo" else hallados[0]


def extraer_hechos_del_texto(texto):
    """Anclas de rol, nombre real y montos etiquetados. El CI es el del comprador."""
    t = texto or ""
    hechos = {}

    ids = extraer_ids_vendedor_comprador(t)
    if ids.get("vendedor"):
        hechos["ci_vendedor"] = ids["vendedor"]
    if ids.get("comprador") and limpiar_ci(ids["comprador"]) != limpiar_ci(ids.get("vendedor")):
        hechos["ci"] = ids["comprador"]

    marca = _valor_etiquetado(t, r"marca")
    modelo = _valor_etiquetado(t, r"modelo")
    m_combo = re.search(r"(?i)m[aá]?rca\s*/\s*modelo\s*[:\s]*\n+\s*([^\n]+)", t)
    if not m_combo:
        m_combo = re.search(r"(?i)m[aá]?rca\s*/\s*modelo\s+([A-Za-zÁÉÍÓÚñÑ0-9][^\n]*)", t)
    if m_combo:
        combo = re.sub(r"(?i)\(modelo referencial\)", "", m_combo.group(1)).strip(" .;,-")
        partes = combo.split()
        if len(partes) >= 2:
            marca = marca or partes[0]
            modelo = modelo or " ".join(partes[1:])
        elif combo:
            marca = marca or combo
    if marca:
        marca = re.split(r"\s{2,}|\bmodelo\b|\btipo\b", marca, maxsplit=1, flags=re.I)[0].strip(" .;,-")
        if marca.lower() not in COLORES_NO_MARCA and not _es_placeholder(marca):
            hechos["marca"] = marca
    if modelo:
        modelo = re.split(r"\s{2,}|\btipo\b|\ba[nñ]o\b|\bano\b|\bcolor\b|\(", modelo, maxsplit=1, flags=re.I)[0]
        modelo = modelo.strip(" .;,-")
        if modelo and not _es_placeholder(modelo) and modelo.lower() not in _MODELO_GENERICO:
            hechos["modelo"] = modelo

    beneficiario = extraer_nombre_comprador(t)
    if beneficiario and _nombre_ok(beneficiario):
        hechos["beneficiario"] = beneficiario

    total = _monto_tras_etiqueta(
        t,
        (
            r"valor\s+total",
            r"total\s*a?\s*pagar",
            r"total\s+(?:factura|proforma|referencial|cotizaci[oó]n|presupuesto)",
            r"(?<![a-záéíóúñ])total\s*:",
        ),
        tomar="ultimo",
    )
    if total > 0:
        hechos["costo_total"] = total

    neto = _monto_tras_etiqueta(
        t,
        (
            r"precio\s+neto(?:\s+veh[ií]culo)?",
            r"p(?:recio)?\.?\s*unitario",
            r"precio\s+unitario",
        ),
    )
    if neto <= 0:
        m_veh = re.search(
            r"(?i)veh[ií]culo[^\n]{0,90}\$\s*([\d.,]+)",
            t,
        )
        if m_veh:
            neto = parse_costo(m_veh.group(1), reparar_concatenado=False)
    if total > 0 and neto > 0 and total > neto * 2.2:
        s = str(int(round(total)))
        if len(s) >= 3:
            alt = parse_costo(s[1:], reparar_concatenado=False)
            if neto * 0.95 <= alt <= neto * 1.6:
                total = alt
    if neto > total > 0:
        s = str(int(round(neto)))
        if len(s) >= 3:
            alt = parse_costo(s[1:], reparar_concatenado=False)
            if 0 < alt <= total:
                neto = alt
    if neto > 0 and total > 0 and neto < total * 0.2:
        neto = 0.0
    if total > 0:
        hechos["costo_total"] = total
    if neto > 0:
        hechos["costo_mas_alto"] = neto

    return hechos


def aplicar_hechos(datos, hechos):
    """Veta el ID del emisor; solo impone el del comprador si el rol está claro."""
    out = dict(datos or {})
    if not hechos:
        return normalizar_datos_llm(out)

    ci_llm = limpiar_ci(out.get("ci"))
    ci_vend = limpiar_ci(hechos.get("ci_vendedor"))
    ci_comp = limpiar_ci(hechos.get("ci"))
    if ci_comp and ci_vend and ci_comp == ci_vend:
        ci_comp = ""

    if ci_comp and ci_vend and ci_comp != ci_vend:
        out["ci"] = ci_comp
    elif ci_llm and ci_vend and ci_llm == ci_vend:
        out["ci"] = ci_comp or ""
    elif ci_comp and (not ci_llm or ci_llm == ci_vend):
        out["ci"] = ci_comp

    if ci_vend and limpiar_ci(out.get("ci")) == ci_vend:
        out["ci"] = ci_comp if (ci_comp and ci_comp != ci_vend) else ""

    total_h = parse_costo(hechos.get("costo_total"), reparar_concatenado=False)
    neto_h = parse_costo(hechos.get("costo_mas_alto"), reparar_concatenado=False)
    neto_llm = parse_costo(out.get("costo_mas_alto"), reparar_concatenado=False)
    if total_h > 0:
        out["costo_total"] = total_h
    total_final = parse_costo(out.get("costo_total"), reparar_concatenado=False)
    if neto_h > 0 and (neto_h != total_h or neto_llm <= 0):
        out["costo_mas_alto"] = neto_h
    elif 0 < neto_llm <= (total_final or neto_llm):
        out["costo_mas_alto"] = neto_llm
    elif neto_h > 0:
        out["costo_mas_alto"] = neto_h

    ben_llm = out.get("beneficiario")
    ben_txt = hechos.get("beneficiario")
    elegido = _mejor_nombre(ben_txt, ben_llm)
    if _nombre_ok(elegido):
        out["beneficiario"] = elegido
    elif not _nombre_ok(ben_llm):
        out["beneficiario"] = None
    if hechos.get("marca") and (
        _es_placeholder(out.get("marca")) or str(out.get("marca") or "").strip().lower() in COLORES_NO_MARCA
    ):
        out["marca"] = hechos["marca"]
    if hechos.get("modelo") and _es_placeholder(out.get("modelo")):
        out["modelo"] = hechos["modelo"]
    return normalizar_datos_llm(out)


def normalizar_datos_llm(datos):
    if not isinstance(datos, dict):
        return {"error": "La IA no devolvió un objeto JSON"}

    if "error" in datos and len(datos) <= 2:
        return datos

    limpio = dict(datos)
    alias = {
        "beneficiary": "beneficiario",
        "client": "beneficiario",
        "cliente": "beneficiario",
        "buyer": "beneficiario",
        "nombre_cliente": "beneficiario",
        "ruc": "ci",
        "cedula": "ci",
        "cédula": "ci",
        "dni": "ci",
        "brand": "marca",
        "make": "marca",
        "model": "modelo",
        "qty": "cantidad",
        "quantity": "cantidad",
        "total": "costo_total",
        "grand_total": "costo_total",
        "price": "costo_mas_alto",
    }
    for origen, destino in alias.items():
        if origen in limpio and (limpio.get(destino) in (None, "", 0, 0.0) or destino not in limpio):
            limpio[destino] = limpio.get(origen)

    for campo in CAMPOS_EXTRACCION:
        limpio.setdefault(campo, None)

    beneficiario = limpio.get("beneficiario")
    if beneficiario and not _es_placeholder(beneficiario):
        texto_ben = str(beneficiario)
        m_ci = re.search(r"(?i)(?:c[eé]dula|ci|ruc)\s*:?\s*([\d.\- ]{8,20})", texto_ben)
        if m_ci and _es_placeholder(limpio.get("ci")):
            limpio["ci"] = m_ci.group(1)
        limpio["beneficiario"] = re.sub(
            r"\s*\((?:c[eé]dula|ci|ruc)[^)]*\)\s*",
            " ",
            texto_ben,
            flags=re.I,
        ).strip()
    if limpio.get("beneficiario") and not _nombre_ok(limpio.get("beneficiario")):
        limpio["beneficiario"] = None

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

    limpio["costo_mas_alto"] = parse_costo(limpio.get("costo_mas_alto"), reparar_concatenado=False)
    limpio["costo_total"] = parse_costo(limpio.get("costo_total"), reparar_concatenado=False)
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

    costo_total = parse_costo(datos.get("costo_total"), reparar_concatenado=False)
    costo_mas_alto = parse_costo(datos.get("costo_mas_alto"), reparar_concatenado=False)
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
    if _es_placeholder(beneficiario) or not _nombre_ok(beneficiario):
        errores.append("Falta beneficiario")
        datos["beneficiario"] = None
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

    if marca and datos.get("beneficiario") and str(marca).strip().lower() == str(datos.get("beneficiario")).strip().lower():
        errores.append("Beneficiario parece la marca")
        datos["beneficiario"] = None

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
Extrae datos de una factura, proforma o cotización de vehículos.
El diseño puede ser cualquiera: electrónica, escaneada, dos columnas, SRI, SII, membrete libre, etc.
No asumas un layout. Distingue ROLES, no posiciones.

Devuelve SOLO un JSON con exactamente estas claves en español:
- marca, modelo, cantidad, beneficiario, ci, costo_mas_alto, costo_total

ROLES:
- EMISOR/VENDEDOR: quien emite el documento (membrete, pie, "vendedor", "emisor", "proveedor").
- COMPRADOR/CLIENTE: a quien se vende o cotiza ("comprador", "cliente", "adquiriente", "receptor", "preparado para", "facturar a").

REGLAS:
1. beneficiario y ci son SIEMPRE del COMPRADOR. Nunca del emisor. Nunca uses la clave beneficiary.
2. Si hay dos RUT/RUC/cédulas, descarta la del emisor. El bloque del cliente puede ir arriba, abajo o al lado.
3. ci: solo dígitos. Puede ser cédula de persona o RUT/RUC de empresa compradora. Nunca un precio.
4. beneficiario: persona natural O razón social/empresa que compra. Nunca el emisor, nunca un RUT, nunca "cliente"/"proveedor", nunca la marca.
5. costo_total = total a pagar / valor total / total factura, no un subtotal ni IVA.
6. costo_mas_alto = precio neto o unitario del vehículo, no un accesorio. Puede ser menor que el total.
7. Números con punto decimal y SIN miles. Chile: 37.344.066 → 37344066. Ecuador: 48,500.00 → 48500.00
8. Marca comercial, nunca un color. No inventes ni copies "valor" ni "No especificado".
9. Si un campo no está (plantilla vacía), null (costos 0.0).
""".strip()


def es_pagina_nativa(pdf_page):
    """Nativa si hay texto útil y no una foto/escaneo que cubra la hoja."""
    texto = pdf_page.get_text("text").strip()
    cover = 0.0
    max_pix = 0
    try:
        page_area = abs(pdf_page.rect.width * pdf_page.rect.height) or 1.0
        doc = pdf_page.parent
        seen = set()
        for img in pdf_page.get_images(full=True):
            xref = img[0]
            if xref in seen:
                continue
            seen.add(xref)
            try:
                info = doc.extract_image(xref)
                max_pix = max(max_pix, int(info.get("width") or 0) * int(info.get("height") or 0))
            except Exception:
                pass
            try:
                for r in pdf_page.get_image_rects(xref):
                    cover += abs(r.width * r.height)
            except Exception:
                pass
        cover = min(cover / page_area, 1.0)
    except Exception:
        cover = 0.0

    n = len(texto)
    blanks = texto.count("_")
    escaneo_grande = max_pix >= 400 * 400
    if cover >= 0.45 and (n < 1200 or escaneo_grande):
        return False, texto
    if escaneo_grande and n < 200:
        return False, texto
    if cover >= 0.20 and (n < 120 or blanks >= 12):
        return False, texto
    if n < 40:
        return False, texto
    return True, texto


def extraer_texto_doctr(imagen_pil):
    from doctr.io import DocumentFile

    buf = BytesIO()
    imagen_pil.convert("RGB").save(buf, format="PNG")
    res = get_modelo_doctr()(DocumentFile.from_images(buf.getvalue()))
    texto = ""
    for page in res.pages:
        for block in page.blocks:
            texto += " ".join([w.value for line in block.lines for w in line.words]) + "\n"
    return texto.strip()


TESS_CFG = "--oem 1 --psm 6"


def _calidad_ocr(texto):
    t = texto or ""
    letras = len(re.findall(r"[A-Za-záéíóúÁÉÍÓÚñÑ]", t))
    claves = len(re.findall(
        r"(?i)\b(?:factura|proforma|rut|ruc|total|marca|modelo|cliente|comprador|vendedor|cédula|cedula)\b",
        t,
    ))
    return letras + claves * 25


def _ocr_tiene_cliente(texto):
    t = texto or ""
    return bool(re.search(
        r"(?i)(?:nombre|cliente|c\.?\s*i\.?|c[eé]dula|ruc)\s*:",
        t,
    )) and bool(re.search(r"[A-ZÁÉÍÓÚÑ][a-záéíóúñ]{2,}", t))


def _parece_texto_invertido(texto):
    """Si al invertir cada palabra aparecen más anclas de factura, el OCR está al revés."""
    t = texto or ""
    claves = (
        r"(?i)\b(?:cliente|proveedor|factura|proforma|nombre|total|marca|"
        r"modelo|documento|comprador|vendedor|cotizaci[oó]n)\b"
    )
    n_ok = len(re.findall(claves, t))
    invertido = " ".join(w[::-1] for w in re.findall(r"[A-Za-záéíóúñÁÉÍÓÚÑ]{4,}", t))
    n_inv = len(re.findall(claves, invertido))
    return n_inv >= 2 and n_inv > n_ok


def _rotar_osd(img):
    try:
        osd = pytesseract.image_to_osd(img)
        m = re.search(r"Rotate:\s*(\d+)", osd)
        if m:
            ang = int(m.group(1)) % 360
            if ang:
                return img.rotate(360 - ang, expand=True, fillcolor="white")
    except Exception:
        pass
    return img


def extraer_texto_ocr(imagen_pil):
    """Tesseract con orientación; DocTR si faltan anclas. Combina 180° si el pie está invertido."""
    if imagen_pil is None:
        return ""
    img = imagen_pil.convert("RGB")
    w, h = img.size
    if max(w, h) < 1400:
        scale = 1400 / max(w, h)
        img = img.resize((int(w * scale), int(h * scale)), Image.Resampling.LANCZOS)
    img = _rotar_osd(img)

    t0 = pytesseract.image_to_string(img, lang="spa", config=TESS_CFG)
    texto = t0 or ""
    if _parece_texto_invertido(t0) or _calidad_ocr(t0) < 140:
        t180 = pytesseract.image_to_string(
            img.rotate(180, expand=True, fillcolor="white"),
            lang="spa",
            config=TESS_CFG,
        )
        q0, q180 = _calidad_ocr(t0), _calidad_ocr(t180)
        if q180 > q0 * 1.1 and q180 > 80:
            texto = t180
        elif _ocr_tiene_cliente(t180) and not _ocr_tiene_cliente(t0):
            texto = (t0 or "") + "\n" + (t180 or "")
        elif _parece_texto_invertido(t0) and q180 > 40:
            texto = (t0 or "") + "\n" + (t180 or "")

    n_montos = len(re.findall(r"\$\s*[\d.,:]{3,}", texto))
    anclas = bool(re.search(r"(?i)factura|proforma|rut|ruc|total|marca|vendedor|comprador", texto or ""))
    tess_util = anclas and _calidad_ocr(texto) >= 120
    if not tess_util or n_montos < 1:
        try:
            td = extraer_texto_doctr(img)
            if td:
                if tess_util:
                    texto = (texto or "") + "\n" + td
                elif _calidad_ocr(td) > _calidad_ocr(texto):
                    texto = td
        except Exception as e:
            print(f"[⚠️ DocTR: {e}]")
    return (texto or "").strip()


def _combinar_textos(nativo, ocr):
    a, b = (nativo or "").strip(), (ocr or "").strip()
    if not a:
        return b
    if not b:
        return a
    if _calidad_ocr(b) >= _calidad_ocr(a) and len(b) >= max(80, int(len(a) * 0.6)):
        return b
    return a + "\n" + b


def extraer_json_de_texto(txt):
    txt = (txt or "").strip()
    if txt.startswith("```"):
        txt = re.sub(r"^```(?:json)?\n?(.*?)\n?```$", r"\1", txt, flags=re.DOTALL)
    match = re.search(r"\{.*\}", txt, re.DOTALL)
    if not match:
        raise ValueError("No se encontraron llaves {} de JSON.")
    return json.loads(match.group(0))


def estructurar_con_llm(texto, modelo="qwen2.5:3b", intento_nombre=None, motivo_fallo=None, hechos=None):
    extra = ""
    if hechos and hechos.get("ci_vendedor"):
        extra += (
            f"\nEl identificador del EMISOR/VENDEDOR es {hechos['ci_vendedor']}. "
            "Prohibido usarlo como ci o mezclarlo con el comprador.\n"
        )
    if hechos and _nombre_ok(hechos.get("beneficiario")):
        extra += (
            f"\nNombre o razón social del COMPRADOR en el texto: {hechos['beneficiario']}. "
            "No lo reemplaces por un RUT, un total ni por 'cliente'.\n"
        )
    if hechos and parse_costo(hechos.get("costo_total"), reparar_concatenado=False) > 0:
        extra += f"\nTotal a pagar en el texto: {hechos['costo_total']}.\n"
    if hechos and parse_costo(hechos.get("costo_mas_alto"), reparar_concatenado=False) > 0:
        extra += f"Precio neto/unitario del vehículo en el texto: {hechos['costo_mas_alto']}.\n"
    extra += (
        "\nEl formato del documento es desconocido. Identifica emisor vs comprador por el rol "
        "(quién vende / a quién se factura), no por el orden de aparición.\n"
    )
    if motivo_fallo:
        extra += (
            f"\nEl intento anterior falló por: {motivo_fallo}\n"
            "Corrige esos campos. No pongas 0.0 si el texto tiene un total.\n"
        )
    prompt = f"{PROMPT_CAMPOS}{extra}\nTEXTO:\n{texto}"
    try:
        res = ollama.chat(
            model=modelo,
            format='json',
            messages=[{'role': 'user', 'content': prompt}],
            options={'temperature': 0.0},
        )
        datos = normalizar_datos_llm(extraer_json_de_texto(res['message']['content']))
        etiqueta = intento_nombre or modelo
        print(f"\n--- RESPUESTA {etiqueta} ---")
        print(json.dumps(datos, indent=2, ensure_ascii=False, default=str))
        return datos
    except Exception as e:
        return {"error": str(e)}


def extraer_con_minicpm(imagenes_pil, max_reintentos=2):
    img_b64 = []
    for img in imagenes_pil[:2]:
        img_resized = img.copy()
        img_resized.thumbnail((1600, 1600))
        buf = BytesIO()
        img_resized.save(buf, format="JPEG", quality=85)
        img_b64.append(base64.b64encode(buf.getvalue()).decode("utf-8"))

    prompt = (
        PROMPT_CAMPOS
        + "\nAnaliza la imagen. ci va entre comillas y solo con dígitos de cédula/RUT/RUC, nunca un precio."
        " beneficiario es el comprador: persona o razón social, nunca un RUT ni la marca del vehículo."
    )
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
            time.sleep(1)

    return {"error": "MiniCPM-V falló en todos los reintentos."}


def requiere_reextraccion(motivo_fallo):
    """MiniCPM/otro LLM no van a meter un vehículo inventado en la BD."""
    partes = [p.strip().lower() for p in (motivo_fallo or "").split("|") if p.strip()]
    if not partes:
        return True
    recuperables = (
        "faltan costos",
        "falta ci",
        "falta beneficiario",
        "inválido",
        "invalido",
        "ilógico",
        "ilogico",
        "error de ia",
        "demasiado corto",
        "mal formado",
    )
    return any(any(clave in parte for clave in recuperables) for parte in partes)


def fusionar_sin_pisar(datos_1, datos_2):
    """Prefiere valores reales del intento 1; nunca reemplaza un costo > 0 por 0."""
    base = normalizar_datos_llm(dict(datos_1 or {}))
    extra = normalizar_datos_llm(dict(datos_2 or {}))
    for campo in CAMPOS_EXTRACCION:
        val_base = base.get(campo)
        val_extra = extra.get(campo)
        if campo in ("costo_total", "costo_mas_alto"):
            vb = parse_costo(val_base, reparar_concatenado=False)
            ve = parse_costo(val_extra, reparar_concatenado=False)
            if ve > 0 and (vb <= 0 or (vb < ve / 20 and vb < 5000)):
                base[campo] = ve
            continue
        if campo == "beneficiario":
            elegido = _mejor_nombre(val_base, val_extra)
            if elegido:
                base[campo] = elegido
            continue
        if campo in ("marca", "modelo"):
            if _es_placeholder(val_base) and not _es_placeholder(val_extra):
                base[campo] = val_extra
            elif not _es_placeholder(val_extra) and not _es_placeholder(val_base):
                marca_b = val_base if campo == "marca" else base.get("marca")
                modelo_b = val_base if campo == "modelo" else base.get("modelo")
                marca_e = val_extra if campo == "marca" else (base.get("marca") or extra.get("marca"))
                modelo_e = val_extra if campo == "modelo" else (base.get("modelo") or extra.get("modelo"))
                hit_b = buscar_vehiculo_fuzzy(marca_b, modelo_b)[0]
                hit_e = buscar_vehiculo_fuzzy(marca_e, modelo_e)[0]
                if hit_e and not hit_b:
                    base[campo] = val_extra
            continue
        if _es_placeholder(val_base) and val_extra and not _es_placeholder(val_extra):
            base[campo] = val_extra
    return normalizar_datos_llm(base)


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


def _zona_cabecera(texto, n_lineas=16, n_chars=800):
    lineas = [ln.strip() for ln in (texto or "").splitlines() if ln.strip()]
    return "\n".join(lineas[:n_lineas])[:n_chars]


def extraer_paginacion(texto):
    m = re.search(r"(?i)p[aá]g(?:ina|\.)?\s*(\d+)\s*(?:de|/)\s*(\d+)", texto or "")
    if m:
        return int(m.group(1)), int(m.group(2))
    m = re.search(r"(?i)hoja\s+(\d+)\s*(?:de|/)\s*(\d+)", texto or "")
    if m:
        return int(m.group(1)), int(m.group(2))
    return None


def _id_doc_descartable(valor):
    v = re.sub(r"\s+", "", str(valor or "").replace("–", "-"))
    d = re.sub(r"\D", "", v)
    if not d:
        return True
    if re.fullmatch(r"0*9{6,}0*1?", d) or re.fullmatch(r"(\d)\1{6,}", d):
        return True
    if len(d) == 13 and d.startswith("0") and not re.search(r"\d{3}-\d{3}-\d{4,}", v):
        return True
    return False


_RE_TIPO_DOC = (
    r"(?:factura(?:\s+electr[oó]nica)?|proforma(?:\s+comercial)?|"
    r"cotizaci[oó]n|presupuesto|nota\s+de\s+venta|dossier)"
)
_RE_NUM_DOC = r"(?:n[úu]m(?:ero)?|n[°ºo0*\.]|nro\.?|no\.?|#)"


def extraer_id_documento(texto):
    """Folio / número de documento, no RUT/RUC ni teléfono."""
    t = (texto or "").replace("\u00a0", " ")
    cab = t[:2500]
    patrones = [
        rf"(?i)(?:{_RE_TIPO_DOC})?\s*(?:{_RE_NUM_DOC})[^\nA-Z0-9]{{0,24}}"
        r"([A-Z]{1,8}[-/]\d{2,8}(?:[-/]\d{2,8}){0,3})",
        rf"(?i)(?:{_RE_TIPO_DOC})[\s\S]{{0,60}}?(\d{{3}}-\d{{3}}-\d{{4,9}})",
        rf"(?i)(?:{_RE_TIPO_DOC})[\s\S]{{0,90}}?(\d{{3,6}}\s*[-–]\s*\d{{5,12}})",
        rf"(?i)(?:{_RE_TIPO_DOC})\s*(?:{_RE_NUM_DOC})\s*[:.\-]?\s*(\d{{3,10}})",
        rf"(?i)(?:{_RE_NUM_DOC})\s*[:.\-]?\s*"
        r"([A-Z]{1,8}[-/]\d{2,8}(?:[-/]\d{2,8}){0,3}|\d{3,6}\s*[-–]\s*\d{5,12}|\d{3,10})",
    ]
    for patron in patrones:
        for m in re.finditer(patron, cab):
            prev = cab[max(0, m.start() - 22): m.start()].lower()
            if re.search(r"(?:rut|ruc|c\.?\s*i\.?|c[eé]dula|nit|dni|tel|cel|fax)\s*:?\s*$", prev):
                continue
            val = re.sub(r"\s+", "", m.group(1).replace("–", "-"))
            if _id_doc_descartable(val):
                continue
            if re.fullmatch(r"\d{4,8}", val) and re.search(r"\d{3,6}-\d{5,}", cab[m.start(): m.end() + 16]):
                continue
            return val
    if re.search(r"(?i)\b(?:vendedor|comprador|emisor|adquir|cliente|proveedor)\b", cab):
        m = re.search(r"\b(\d{3,6}[-–]\d{7,12})\b", cab)
        if m and not _id_doc_descartable(m.group(1)):
            return m.group(1).replace("–", "-")
    return None


def _ids_iguales(a, b):
    if not a or not b:
        return False
    na = re.sub(r"[\s\-]", "", str(a).upper())
    nb = re.sub(r"[\s\-]", "", str(b).upper())
    return na == nb


def _cabecera_titulos(texto):
    cab = _zona_cabecera(texto, 20, 1000)
    return re.sub(
        r"(?i)(?:esta|la presente|una|presente|modelo de|copia de)\s+"
        r"(?:proforma|factura|cotizaci[oó]n|presupuesto).{0,140}",
        " ",
        cab,
    )


def _tipo_documento(texto):
    cab = _cabecera_titulos(texto).lower()
    m = re.search(
        r"\b(factura(?:\s+electr[oó]nica)?|nota\s+de\s+venta|"
        r"proforma(?:\s+comercial)?|cotizaci[oó]n|presupuesto|dossier)\b",
        cab,
    )
    if not m:
        return None
    w = m.group(1).lower()
    if w.startswith("fact") or "nota" in w:
        return "factura"
    return "proforma"


def parece_continuacion(texto):
    t = texto or ""
    pag = extraer_paginacion(t)
    if pag and pag[0] > 1:
        return True
    cab = _zona_cabecera(t).lower()
    cab_t = _cabecera_titulos(t).lower()
    if re.search(r"\b(?:factura|proforma|cotizaci[oó]n|presupuesto|dossier|nota de venta)\b", cab_t) and not (
        pag and pag[0] > 1
    ):
        return False
    if re.search(
        r"(?i)(?:nombre|raz[oó]n(?:\s*social)?|c\.?i\.?|ruc|rut|tel[eé]fono|correo|email)\s*:\s*_{3,}"
        r"|(?:recibido|aceptado)\s*/\s*conforme"
        r"|firma\s+(?:del\s+)?(?:cliente|comprador)",
        t[:1200],
    ) and not re.search(r"(?i)\b(?:factura|proforma|cotizaci[oó]n)\b.{0,24}(?:n[°ºo]|no\.?|\d)", cab_t):
        return True
    if len(t.strip()) < 120:
        return True
    inicios = (
        "términos y condiciones", "terminos y condiciones",
        "especificaciones técnicas", "especificaciones tecnicas",
        "formas de pago", "condiciones comerciales", "observaciones",
        "notas importantes", "firma vendedor", "anexo",
        "homologacion", "homologación", "garantia", "garantía",
        "certificacion", "certificación",
    )
    return any(cab.startswith(s) or s in cab[:90] for s in inicios)


def es_portada_documento(texto):
    t = texto or ""
    pag = extraer_paginacion(t)
    if pag and pag[0] > 1:
        return False
    if pag and pag[0] == 1:
        return True

    cab = _cabecera_titulos(t)
    patrones = [
        r"(?i)\bfactura(\s+electr[oó]nica)?\b",
        r"(?i)\bproforma(\s+comercial)?\b",
        r"(?i)\bcotizaci[oó]n\b",
        r"(?i)\bpresupuesto\b",
        r"(?i)\bnota\s+de\s+venta\b",
        r"(?i)\bdossier\b",
        rf"(?i){_RE_TIPO_DOC}\s*{_RE_NUM_DOC}",
    ]
    if any(re.search(p, cab) for p in patrones):
        return True
    if re.search(
        r"(?i)(?:vendedor|emisor|proveedor).{0,48}(?:comprador|cliente|adquir)"
        r"|(?:comprador|cliente|adquir).{0,48}(?:vendedor|emisor|proveedor)",
        cab,
    ):
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
    if extraer_id_documento(texto or "") and re.search(
        r"total\s*a?\s*pagar|valor total|total (?:factura|proforma|referencial|cotizaci|presupuesto)",
        t,
    ):
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
    tipo_actual = None
    cerrado = False

    for pagina in paginas_data:
        texto = pagina.get("texto") or ""
        id_pagina = extraer_id_documento(texto[:2500])
        pag = extraer_paginacion(texto)
        es_pag_sig = bool(pag and pag[0] > 1)
        portada = es_portada_documento(texto)
        tipo = _tipo_documento(texto)
        continuacion = parece_continuacion(texto)
        if portada and not es_pag_sig:
            continuacion = False

        mismo_id = _ids_iguales(id_pagina, id_actual)
        nuevo_id = bool(id_pagina and id_actual and not mismo_id)
        nuevo_tipo = bool(tipo and tipo_actual and tipo != tipo_actual and not es_pag_sig)

        if actual is None:
            actual = _doc_vacio()
            _agregar_pagina(actual, pagina)
            id_actual = id_pagina
            tipo_actual = tipo
            cerrado = documento_parece_cerrado(texto)
            continue

        if mismo_id or (continuacion and not nuevo_id and not nuevo_tipo):
            _agregar_pagina(actual, pagina)
            if id_pagina:
                id_actual = id_pagina
            if tipo:
                tipo_actual = tipo
            cerrado = documento_parece_cerrado(actual["texto"])
            continue

        debe_partir = nuevo_id or nuevo_tipo or (portada and not es_pag_sig)
        if debe_partir:
            docs.append(actual)
            actual = _doc_vacio()
            _agregar_pagina(actual, pagina)
            id_actual = id_pagina
            tipo_actual = tipo
            cerrado = documento_parece_cerrado(texto)
            continue

        _agregar_pagina(actual, pagina)
        if id_pagina and not id_actual:
            id_actual = id_pagina
        if tipo and not tipo_actual:
            tipo_actual = tipo
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


def rasterizar_paginas(ruta_archivo, indices, dpi=250):
    """Convierte solo las páginas que hacen falta (escaneadas o MiniCPM)."""
    if not indices:
        return {}
    poppler = PATH_POPPLER if PATH_POPPLER and os.path.isdir(PATH_POPPLER) else None
    primero = min(indices) + 1
    ultimo = max(indices) + 1
    imagenes = convert_from_path(
        ruta_archivo,
        poppler_path=poppler,
        first_page=primero,
        last_page=ultimo,
        dpi=dpi,
    )
    by_page = {min(indices) + i: img for i, img in enumerate(imagenes)}
    return {i: by_page[i] for i in indices if i in by_page}


def extraer_y_validar(texto, imagenes_doc, es_nativo):
    """Un LLM de texto + anclas regex. MiniCPM solo si es escaneado y faltan campos."""
    hechos = extraer_hechos_del_texto(texto)
    print(f"[INFO] Hechos anclados del texto: {json.dumps(hechos, ensure_ascii=False, default=str)}")

    datos = estructurar_con_llm(
        texto, modelo="qwen2.5:3b", intento_nombre="INTENTO 1 (Qwen texto)", hechos=hechos
    )
    datos = aplicar_hechos(datos, hechos)
    es_valido, msj = evaluar_extraccion(datos)

    if es_valido or not requiere_reextraccion(msj):
        if not es_valido:
            print(f"\n[FALLO: {msj}] Sin reintento de modelo: el texto ya cubre los campos; queda catálogo/validación.")
        return datos, es_valido, msj

    if es_nativo:
        print(f"\n[FALLO INTENTO 1: {msj}]. Reintento de texto con Qwen (sin MiniCPM-V).")
        datos_2 = estructurar_con_llm(
            texto,
            modelo="qwen2.5:3b",
            intento_nombre="INTENTO 2 (Qwen, campos faltantes)",
            motivo_fallo=msj,
            hechos=hechos,
        )
        datos = aplicar_hechos(fusionar_sin_pisar(datos, datos_2), hechos)
        es_valido, msj = evaluar_extraccion(datos)
        return datos, es_valido, msj

    print(f"\n[FALLO INTENTO 1: {msj}]. Escaneado: MiniCPM-V una sola pasada visual.")
    datos_2 = extraer_con_minicpm(imagenes_doc, max_reintentos=2)
    datos = aplicar_hechos(fusionar_sin_pisar(datos, datos_2), hechos)
    es_valido, msj = evaluar_extraccion(datos)
    return datos, es_valido, msj


def procesar_archivo_interno(ruta_archivo):
    print(f"\n[{os.path.basename(ruta_archivo)}] Iniciando análisis...")
    es_pdf = ruta_archivo.lower().endswith('.pdf')
    imagenes_por_idx = {}

    paginas_data = []
    indices_ocr = []

    try:
        if es_pdf:
            doc_pdf = pymupdf.open(ruta_archivo)
            for i, page in enumerate(doc_pdf):
                es_nativo, texto = es_pagina_nativa(page)
                if es_nativo:
                    paginas_data.append({"idx": i, "texto": texto, "es_nativo": True, "imagen": None})
                else:
                    indices_ocr.append(i)
                    paginas_data.append({"idx": i, "texto": texto, "es_nativo": False, "imagen": None})
            doc_pdf.close()
        else:
            indices_ocr = [0]
            paginas_data.append({"idx": 0, "texto": "", "es_nativo": False, "imagen": None})
    except Exception as e:
        fallo = {
            "estado": f"REVISIÓN (No se pudo leer el archivo: {e})",
            "archivo": os.path.basename(ruta_archivo),
            "paginas_pdf": [],
            "error": str(e),
        }
        yield serializar_resultado(fallo)
        return

    if indices_ocr:
        print(f"[INFO] Rasterizando {len(indices_ocr)} página(s) no nativas...")
        try:
            if es_pdf:
                imagenes_por_idx = rasterizar_paginas(ruta_archivo, indices_ocr)
            else:
                imagenes_por_idx = {0: Image.open(ruta_archivo).convert("RGB")}
        except Exception as e:
            print(f"[⚠️ Rasterizado falló: {e}]")

        for p in paginas_data:
            if p["es_nativo"]:
                continue
            img = imagenes_por_idx.get(p["idx"])
            p["imagen"] = img
            if img is None:
                continue
            texto = extraer_texto_ocr(img)
            p["texto"] = _combinar_textos(p.get("texto") or "", texto)

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

        es_nativo = not doc.get("tiene_escaneado", False)
        texto_analizar = doc["texto"] or ""

        if (not es_nativo or len(texto_analizar.strip()) < 40) and doc.get("imagenes"):
            imgs_faltantes = [img for img in doc["imagenes"] if img is not None]
            if imgs_faltantes and len(texto_analizar.strip()) < 40:
                print("[INFO] Texto insuficiente. Completando con DocTR...")
                texto_ocr = "\n".join(extraer_texto_doctr(img) for img in imgs_faltantes)
                texto_analizar = (texto_analizar + "\n" + texto_ocr).strip()

        imagenes_doc = [img for img in (doc.get("imagenes") or []) if img is not None]
        if not imagenes_doc and not es_nativo:
            faltan = [idx for idx in doc["paginas"] if idx not in imagenes_por_idx]
            if faltan and es_pdf:
                imagenes_por_idx.update(rasterizar_paginas(ruta_archivo, faltan))
            imagenes_doc = [imagenes_por_idx[idx] for idx in doc["paginas"] if idx in imagenes_por_idx]

        datos_finales, es_valido, msj = extraer_y_validar(texto_analizar, imagenes_doc, es_nativo)

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
    ruta_ejemplo = "docs/F-0012-I.pdf"
    if not os.path.exists(ruta_ejemplo):
        print(f"ERROR: No se encuentra el archivo en {ruta_ejemplo}.")
    else:
        resultados = list(procesar_archivo_interno(ruta_ejemplo))
        print("\n=== RESULTADOS FINALES EN FORMATO JSON ===")
        print(json.dumps(resultados, indent=4, ensure_ascii=False, default=str))
