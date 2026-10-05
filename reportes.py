"""Excel de la corrida reciente y del histórico. Ítems de tabla van aquí."""
import io
import json
import os
import re
import time
from datetime import datetime

import pandas as pd
from sqlalchemy import text
from sqlalchemy.orm import Session

from base import engine as engine_catalogo

ARCHIVO_HISTORICO = "historico_ejecuciones.xlsx"
DIR_RECIENTE = "reportes_recientes"
ARCHIVO_EXITOS_RECIENTE = os.path.join(DIR_RECIENTE, "reporte_exitos_reciente.xlsx")
ARCHIVO_REVISION_RECIENTE = os.path.join(DIR_RECIENTE, "reporte_revision_reciente.xlsx")
ARCHIVO_CORRIDA_JSON = os.path.join(DIR_RECIENTE, "corrida.json")

TIPOS_CARROCERIA = (
    "SUV", "CAMION", "AUTOMOVIL", "CAMIONETA", "VAN", "BUS", "SCOOTER", "OTRO",
)

COLUMNAS_RESUMEN = (
    "Estado",
    "Fecha_Ejecucion",
    "Marca",
    "Modelo",
    "cantidad",
    "beneficiario",
    "ci",
    "costo_mas_alto",
    "costo_total",
    "marca_base",
    "modelo_base",
    "Tipo",
    "archivo",
)

COLUMNAS_DETALLE = (
    "Marca",
    "Modelo",
    "cantidad",
    "beneficiario",
    "ci",
    "costo",
)

COLUMNAS_DETALLE_HIST = COLUMNAS_DETALLE + (
    "Fecha_Ejecucion",
    "archivo",
)

_MARCAS_COMUNES = (
    "BYD", "HYUNDAI", "TESLA", "RENAULT", "BMW", "AUDI", "VOLVO", "FOTON",
    "MAXUS", "MAXUSE", "GWM", "ORA", "GEELY", "XIAOMI", "SEGWAY", "NIU", "NINEBOT",
    "KIA", "NISSAN", "CHEVROLET", "JAC", "DONGFENG", "YUTONG", "CHERY",
    "VOLKSWAGEN", "MERCEDES", "PEUGEOT", "CITROEN", "TOYOTA", "HONDA",
    "FORD", "JEEP", "MAZDA", "SUZUKI", "MITSUBISHI", "PORSCHE", "POLESTAR",
    "NIO", "XPENG", "ZEEKR", "MINI", "FIAT", "CUPRA", "SEAT", "MG",
    "GET", "KING LONG", "HIGER", "GREAT WALL", "LYNK", "SMART", "WALLBOX",
    "JUICEBOX",
)

_RE_ACCESORIO = re.compile(
    r"(?i)\b(?:wall\s*box|wallbox|cargador|charger|mantenimiento|service|"
    r"matricul|tr[aá]mite|placas?|legaliz|seguro|kit\b|instalaci[oó]n|"
    r"color\s+exterior|incentivos?\s+fiscal|exoneraci[oó]n|inspecci[oó]n|"
    r"homologaci[oó]n|base\s+imponible|tratamiento\s+impositivo|"
    r"infraestructura\s+de\s+recarga|\bOVA\b|gravamen)\b"
)
_RE_NOTA_FISCAL = re.compile(
    r"(?i)(?:base\s+imponible|tratamiento\s+impositivo|\bOVA\b|"
    r"infraestructura\s+de\s+recarga|gravamen|"
    r"mantenimientos?\s+preventivos|homologaci[oó]n\s+y\s+placas)"
)
_RE_ELECTRICO = re.compile(
    r"(?i)\b(?:veh[ií]culo|el[eé]ctrico|\bev\b|suv|sed[aá]n|hatch|cami[oó]n|"
    r"bus|moto|scooter|patinete|pickup|camioneta|carro)\b"
)
_MARCAS_NO_VEHICULO = {"WALLBOX", "JUICEBOX"}


def fecha_hoy():
    return datetime.now().strftime("%d-%m-%Y")


def normalizar_tipo(texto):
    t = str(texto or "").upper()
    if re.search(r"\bSUV\b", t):
        return "SUV"
    if re.search(r"CAMIONETA|PICK-?UP", t):
        return "CAMIONETA"
    if re.search(r"\bVAN\b|FURG[OÓ]N", t):
        return "VAN"
    if re.search(r"\bBUS\b|AUTOB[UÚ]S", t):
        return "BUS"
    if re.search(r"SCOOT+ER|SCOTTER|PATINETE|\bMOTO\b", t):
        return "SCOOTER"
    if re.search(r"CAMI[OÓ]N|TRUCK", t):
        return "CAMION"
    if re.search(r"AUTO|SED[AÁ]N|HATCH|COUP[EÉ]", t):
        return "AUTOMOVIL"
    if t.strip() in TIPOS_CARROCERIA:
        return t.strip()
    return "OTRO"


def _es_nota_no_producto(texto):
    return bool(_RE_NOTA_FISCAL.search(str(texto or "")))


def clasificar_item(descripcion, etiqueta=""):
    d = f"{etiqueta} {descripcion}".strip()
    if re.search(r"(?i)color\s+exterior|incentivos?\s+fiscal|exoneraci[oó]n", d):
        return "Accesorio"
    if _es_nota_no_producto(d):
        return "Accesorio"
    if re.search(r"(?i)wall\s*box|wallbox", d):
        return "Wallbox"
    if re.search(r"(?i)\bcargador\b|\bcharger\b|juicebox", d):
        return "Cargador"
    if re.search(r"(?i)mantenimiento|service|inspecci[oó]n", d):
        return "Mantenimiento"
    if re.search(r"(?i)matricul|tr[aá]mite|placas?|legaliz|soat|seguro", d):
        return "Trámite"
    marca, _ = partir_marca_modelo(descripcion)
    marca_veh = bool(marca) and str(marca).upper() not in _MARCAS_NO_VEHICULO
    if re.search(r"(?i)scooter|patinete", d):
        return "Scooter"
    if re.search(r"(?i)\bmoto\b", d) and not re.search(r"(?i)motor", d):
        return "Moto"
    if re.search(r"(?i)\bbus\b|autob[uú]s", d):
        return "Bus"
    if re.search(r"(?i)\bvan\b|furg[oó]n", d):
        return "Van"
    if re.search(r"(?i)cami[oó]n", d):
        return "Camión"
    if re.search(r"(?i)autom[oó]vil|suv|sed[aá]n|hatch|veh[ií]culo|carro", d):
        return "Coche"
    if marca_veh:
        return "Coche"
    if _RE_ELECTRICO.search(d):
        return "Coche"
    return "Accesorio"


_RE_ETIQUETA_ITEM = re.compile(
    r"(?i)^(autom[oó]vil|scooter|patinete|moto|cami[oó]n|bus|van|suv|camioneta)$"
)
_RE_CATEGORIA_SOLA = re.compile(
    r"(?i)^(veh[ií]culo(?:\s+el[eé]ctrico)?|furg[oó]n(?:\s+carga)?|"
    r"cargador(?:\s+premium)?|suv(?:\s+el[eé]ctrico)?|"
    r"scooter(?:\s+el[eé]ctrico|\s+profesional)?|"
    r"city\s+car(?:\s+ev)?|sed[aá]n(?:\s+ejecutivo)?|hatchback(?:\s+ev)?|"
    r"gran\s+coup[eé](?:\s+ev)?|autob[uú]s(?:\s+\d+\s+metros)?|"
    r"estaci[oó]n\s+dc|telemetr[ií]a|autom[oó]vil|patinete|"
    r"cami[oó]n|van|moto|equipo)$"
)
_RE_SPEC_FRAG = re.compile(
    r"(?i)^(de\s+|con\s+|y\s+|e\s+|incluye\s+|sistema\s+|bater[ií]a\s+|"
    r"motor\s+|autonom[ií]a|capacidad\s+|conector\s+|protocolo\s+|"
    r"rines?\s+|pantalla\s+|firmware\s+|instalaci[oó]n\s+|"
    r"integrado\.?$|trif[aá]sico|bif[aá]sico|chasis\s*:|"
    r"control\s+de\s+tracci[oó]n|arquitectura\s+|plataforma\s+|"
    r"potencia\s+|acabado\s*:)"
)
_RE_FILA_TOTAL = re.compile(
    r"(?i)(?:^|\b)(?:subtotal(?:\s|$)|total(?:\s|$|:)|valor\s+total|monto\s+final|"
    r"iva\s*(?:\(\d+\s*%\))?\s*$|descuento(?:\s|$)|base\s+imponible|"
    r"costo\s+operativo|ahorro\s+operacional|valor\s+final|"
    r"precio\s+final|valor\s+a\s+cancelar|liquidaci[oó]n|"
    r"total\s+proforma|total\s+factura|total\s+contrato|"
    r"tratamiento\s+impositivo)"
)
_RE_NO_PRODUCTO = re.compile(
    r"(?i)(?:precio\s+neto|precio\s+de\s+venta|precio\s+final|p\.?\s*unit|"
    r"p\.?\s*total|^subtotal|^total\b|^iva\b|^ice\b|descuento|bono\b|"
    r"^monto\b|^cantidad|^cant\.?|^c[oó]digo|^descripci[oó]n|^concepto|"
    r"tarifa|incluido|cortes[ií]a|bonificaci[oó]n|r[eé]gimen|"
    r"impuestos?|garant[ií]as?|condiciones|pol[ií]ticas|exenciones|"
    r"ahorros?|validez|detalle\s+del|datos\s+del|forma\s+de\s+pago|"
    r"valor\s+a\s+cancelar|base\s+imponible|tratamiento\s+impositivo|"
    r"final\s*:)"
)
_RE_ENCABEZADO_TABLA = re.compile(
    r"(?i)^(c[oó]digo|cant\.?|#|ref\b|[ií]tem|descripci[oó]n|"
    r"detalle\s+del|especificaci[oó]n\s+de\s+unidad|p\.?\s*unit|"
    r"tarifa\s+base|concepto\s*/|monto\s*\(usd\)|pvp\s+oficial|"
    r"importe\s+total|valor\s+unit)"
)
_RE_PRECIO_LINEA = re.compile(
    r"^(?:[-–—•]\s*)?(?:US\$|USD|CLP|\$)\s*-?\s*[\d.,]+\s*(?:USD|US\$|CLP)?\s*$",
    re.I,
)
_RE_PRECIO_OCR = re.compile(
    r"^-?\d{1,3}(?:[.,]\d{3})+(?:[.,]\d{2})?\s*(?:USD|US\$|CLP)?$"
)
_RE_SKU = re.compile(r"(?i)^(?:ev|mc|ve|pt|eq|wb|sc)-?[A-Z0-9\-]{0,12}$")
_PREFIJOS_TIPO = re.compile(
    r"(?i)^(crossover|hatchback|sed[aá]n|suv|furg[oó]n|autob[uú]s|city\s+car|"
    r"veh[ií]culo|autom[oó]vil|carro|scooter(?:\s+profesional)?|patinete|"
    r"el[eé]ctric[oa]|urbano|compacto|ejecutivo|premium)\s+"
)
_STOP_MARCA = {
    "tipo", "marca", "modelo", "color", "año", "ano", "cliente", "item",
    "valor", "total", "cantidad", "detalle", "descripcion", "descripción",
    "vehiculo", "vehículo", "electrico", "eléctrico", "suv", "auto",
    "camion", "camión", "carro", "precio", "neto", "oficial", "pvp", "usd",
}
_CATS_VEHICULO = ("Coche", "Scooter", "Moto", "Bus", "Camión", "Van")

_cache_marcas = None


def _marcas_conocidas():
    global _cache_marcas
    if _cache_marcas is not None:
        return _cache_marcas
    nombres = {m.upper() for m in _MARCAS_COMUNES}
    try:
        with Session(engine_catalogo) as session:
            for (n,) in session.execute(text("SELECT nombre FROM marca")):
                if n and str(n).strip():
                    nombres.add(str(n).strip().upper())
    except Exception:
        pass
    _cache_marcas = sorted(nombres, key=len, reverse=True)
    return _cache_marcas


def estado_excel(datos):
    """EXITO o REVISION: MOTIVO, para Resumen e histórico."""
    est = str((datos or {}).get("estado") or "").strip()
    if re.search(r"(?i)^é?xito\b|^exito\b", est) or "ÉXITO" in est:
        return "EXITO"
    motivo = re.sub(r"(?i)^revisi[óo]n\s*[:\(]?\s*", "", est).strip(" )")
    return f"REVISION: {motivo}" if motivo else "REVISION"


def _es_linea_precio(ln):
    s = (ln or "").strip()
    if not s:
        return False
    if _RE_PRECIO_LINEA.match(s):
        return True
    if _RE_PRECIO_OCR.match(s) and _importe(s) >= 80:
        return True
    return False


def _recortar_bloque_item(bloque):
    utiles = [ln.strip() for ln in bloque if str(ln).strip()]
    start = 0
    for i, ln in enumerate(utiles):
        if _RE_ENCABEZADO_TABLA.match(ln):
            start = i + 1
    recorte = utiles[start:]
    if len(recorte) > 22:
        recorte = recorte[-22:]
    return recorte


def _cantidad_bloque(recorte):
    for ln in reversed(recorte[-8:]):
        if re.fullmatch(r"\d{1,4}", ln):
            n = int(ln)
            if n <= 0 or n > 999:
                continue
            if ln.startswith("0") and len(ln) >= 2:
                continue
            return n
    return 1


def _score_titulo_item(ln, marcas):
    t = str(ln or "").strip(" •-–—\t")
    t = _limpiar_titulo_item(t)
    if not t or len(t) < 3:
        return -100
    if _es_nota_no_producto(t) or re.search(r"(?i)ahorro\s+operacional|costo\s+operativo|proyectado a favor", t):
        return -80
    if re.fullmatch(r"[\d.,\s$USDCLP%\-]+", t, flags=re.I):
        return -100
    if re.fullmatch(r"\d{1,4}", t) or _RE_SKU.match(t):
        return -100
    if _RE_CATEGORIA_SOLA.match(t):
        return -50
    if _RE_NO_PRODUCTO.match(t) or _RE_FILA_TOTAL.match(t):
        return -80
    if re.match(r"(?i)^(precio\s+final|valor\s+a\s+cancelar|final\s*:)", t):
        return -80
    if _RE_SPEC_FRAG.match(t):
        return -40
    if _RE_ETIQUETA_ITEM.match(t):
        return -20
    if len(t) > 140:
        return -10
    sc = 2.0
    hay_marca = False
    for marca in marcas:
        if len(marca) < 2:
            continue
        if re.search(rf"(?i)\b{re.escape(marca)}\b", t) or (
            len(marca) <= 3 and re.search(rf"(?i)\b{re.escape(marca)}\d", t)
        ):
            hay_marca = True
            sc += 12
            if re.match(rf"(?i)^{re.escape(marca)}\b", t) or re.match(rf"(?i)^{re.escape(marca)}\d", t):
                sc += 4
            break
    if re.search(r"[A-Za-záéíóúñÁÉÍÓÚÑ]{2,}\s+[A-Za-z0-9][\w\-]*", t):
        sc += 3
    if re.search(
        r"(?i)\b(?:ev|bev|suv|hatch|sedan|cami[oó]n|van|scooter|kwid|ioniq|"
        r"dolphin|seal|yuan|model\s+[3sxy]|ex30|mg4|ix1|e-?tron|auman|ora|"
        r"firefly|zoe|han)\b",
        t,
    ):
        sc += 4
    if t[:1].islower() and not re.match(r"(?i)^e-", t):
        sc -= 8
    if re.search(
        r"(?i)\b(?:airbags?|wltp|kwh|c[aá]mara 360|google built|android auto)\b",
        t,
    ) and not hay_marca:
        sc -= 8
    if hay_marca and len(t.split()) <= 3 and re.search(r"(?i)\bkW\b|\(\s*\d", t) and not re.search(
        r"(?i)cargador|wallbox|charger|veh[ií]culo|autom", t
    ):
        sc -= 10
    return sc


def _separar_ocr_pegado(texto):
    """TeslaModelY / TESLAMODELY / BYDDolphinMini → marca + modelo."""
    t = str(texto or "")
    if not t:
        return t
    t = re.sub(r"(?i)\blONIQ\b", "IONIQ", t)
    t = re.sub(r"([a-záéíóúñ])([A-ZÁÉÍÓÚÑ])", r"\1 \2", t)
    t = re.sub(r"([A-ZÁÉÍÓÚÑ]{2,})([A-ZÁÉÍÓÚÑ][a-záéíóúñ])", r"\1 \2", t)
    t = re.sub(r"(?i)\bl\s+ONIQ\b", "IONIQ", t)
    t = re.sub(r"(?i)\b(MODEL)([3SXY])\b", r"\1 \2", t)
    partes = []
    for tok in t.split():
        compact = re.sub(r"[^A-Za-z0-9ÁÉÍÓÚÑáéíóúñ]", "", tok)
        cu = compact.upper()
        elegido = tok
        for marca in _marcas_conocidas():
            mu = re.sub(r"[^A-Z0-9]", "", marca.upper())
            if len(mu) < 2 or mu in _MARCAS_NO_VEHICULO:
                continue
            if cu.startswith(mu) and len(cu) > len(mu) + 1:
                rest = compact[len(mu):]
                rest = re.sub(r"([a-záéíóúñ])([A-ZÁÉÍÓÚÑ])", r"\1 \2", rest)
                rest = re.sub(r"(?i)^(MODEL)([3SXY])$", r"\1 \2", rest)
                if marca.upper() in {"BYD", "BMW", "GWM", "MG", "KIA", "JAC", "ORA", "NIO"}:
                    etiqueta = marca.upper()
                elif marca.isupper() and len(marca) > 3:
                    etiqueta = marca.title()
                else:
                    etiqueta = marca
                elegido = f"{etiqueta} {rest}".strip()
                break
        partes.append(elegido)
    return " ".join(partes)


def _limpiar_titulo_item(titulo):
    t = str(titulo or "").strip(" •-–—\t")
    t = re.sub(r"^[\+\-–—•]+\s*", "", t)
    t = re.sub(r"^\d+\s+", "", t)
    t = re.sub(r"(?i)^precio\s+de\s+venta(?:\s+proforma)?\s*[-–—:]?\s*", "", t)
    t = re.sub(r"(?i)\s*\((?:pvp|oficial)[^)]*\)", "", t)
    t = re.sub(r"(?i)\s*\((?:edici[oó]n|modelo)?\s*\d{4}\)", "", t)
    t = re.sub(r"(?i)\s*:\s*(?:US\$|USD|CLP|\$)\s*[\d.,]+.*$", "", t)
    t = _separar_ocr_pegado(t)
    t = re.sub(r"(?i)\bi\s*x\s*(\d)", r"iX\1", t)
    t = re.sub(r"\s+", " ", t).strip(" .;,-")
    return t


def _titulo_desde_etiquetas(recorte):
    marca = modelo = ""
    for i, ln in enumerate(recorte):
        m = re.match(r"(?i)^marca\s*:\s*(.+)$", ln)
        if m:
            marca = m.group(1).strip()
        elif re.match(r"(?i)^marca\s*:?\s*$", ln) and i + 1 < len(recorte):
            marca = recorte[i + 1]
        m = re.match(r"(?i)^modelo\s*:\s*(.+)$", ln)
        if m:
            modelo = m.group(1).strip()
        elif re.match(r"(?i)^modelo\s*:?\s*$", ln) and i + 1 < len(recorte):
            modelo = recorte[i + 1]
    marca = _limpiar_titulo_item(marca)
    modelo = _limpiar_titulo_item(modelo)
    if marca and modelo and modelo.lower() not in ("tipo", "año", "ano", "color"):
        return f"{marca} {modelo}"
    return ""


def _elegir_producto(recorte, valor):
    etiqueta = ""
    for ln in recorte:
        if _RE_ETIQUETA_ITEM.match(ln) or _RE_CATEGORIA_SOLA.match(ln):
            etiqueta = ln
    cantidad = _cantidad_bloque(recorte)
    etiquetado = _titulo_desde_etiquetas(recorte)
    marcas = _marcas_conocidas()
    scored = []
    for ln in recorte:
        limpio = _limpiar_titulo_item(ln)
        scored.append((_score_titulo_item(limpio, marcas), limpio, ln))
    mejor = max((s[0] for s in scored), default=-999)
    titulo = ""
    if mejor >= 1:
        umbral = mejor - 3
        cand = [s for s in scored if s[0] >= umbral and s[1]]
        best = max(cand, key=lambda s: (s[0], len(s[1])))
        titulo = best[1]
        i = recorte.index(best[2]) if best[2] in recorte else -1
        if 0 <= i + 1 < len(recorte) and re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9\- )(]{1,28}", recorte[i + 1]):
            nxt = recorte[i + 1]
            if _score_titulo_item(nxt, marcas) > -20 and not _RE_SPEC_FRAG.match(nxt):
                if len(nxt.split()) <= 4 and not _es_linea_precio(nxt):
                    titulo = f"{titulo} {nxt}".strip()
        titulo = _limpiar_titulo_item(titulo)
    if etiquetado and (not titulo or len(titulo.split()) <= 1):
        return etiquetado, etiqueta, cantidad
    if not titulo:
        return "", etiqueta, cantidad
    return titulo, etiqueta, cantidad


def _bloque_es_total(bloque):
    utiles = [ln.strip() for ln in bloque if str(ln).strip()]
    if not utiles:
        return False
    cerca = utiles[-4:]
    blob_cerca = " ".join(cerca)
    if _es_nota_no_producto(blob_cerca):
        return True
    for ln in cerca:
        if _es_nota_no_producto(ln) or _RE_FILA_TOTAL.search(ln):
            return True
        if re.match(
            r"(?i)^(subtotal\b|valor\s+total\b|monto\s+final\b|descuento\b|"
            r"base\s+imponible\b|precio\s+final\b|valor\s+a\s+cancelar\b|"
            r"liquidaci[oó]n\b|total(?:\s|$|:)|total\s+(?:a\s+pagar|proforma|"
            r"factura|final|contrato|documento|general))",
            ln,
        ):
            return True
        if re.match(r"(?i)^iva\s*(?:\(\d+\s*%\))?\s*:?\s*$", ln):
            return True
        if re.match(r"(?i)^iva\b", ln) and re.search(r"%\)?\s*:?\s*$", ln) and not re.search(r"(?i)\bice\b", ln):
            return True
        if re.search(
            r"(?i)\b(?:valor\s+total\s+facturado|subtotal\s+(?:veh[ií]culo|comercial|cargadores)|"
            r"costo\s+operativo|ahorro\s+operacional|tarifa\s+0\s*%\)?:)\b",
            ln,
        ):
            return True
    return False


def _parece_micromovilidad(texto):
    return bool(re.search(r"(?i)\b(?:scooter|patinete|kickscooter|ninebot)\b", texto or ""))


def partir_marca_modelo(descripcion):
    """Marca comercial + modelo. No usa el primer espacio a ciegas."""
    t0 = _limpiar_titulo_item(descripcion)
    if not t0:
        return "", ""
    t = t0
    for _ in range(4):
        n = _PREFIJOS_TIPO.sub("", t, count=1).strip()
        if n == t:
            break
        t = n
    marcas = _marcas_conocidas()
    mejor = None
    for marca in marcas:
        if len(marca) < 2:
            continue
        m = re.search(rf"(?i)(?:^|[\s/])({re.escape(marca)})\b", t)
        if not m:
            m = re.match(rf"(?i)^({re.escape(marca)})\b", t)
        if m:
            pos = m.start(1) if m.lastindex else m.start()
            if mejor is None or pos < mejor[0] or (pos == mejor[0] and len(marca) > len(mejor[1])):
                mejor = (pos, marca, m.end())
        m2 = re.search(rf"(?i)\b({re.escape(marca)})(\d[\w]*)", t) if len(marca) <= 3 else None
        if m2:
            pos = m2.start(1)
            if mejor is None or pos < mejor[0] or (pos == mejor[0] and len(marca) > len(mejor[1])):
                mejor = (pos, marca, m2.end())
    if mejor:
        _, marca, end = mejor
        resto = t[end:].strip(" -–—/,")
        token = re.match(rf"(?i){re.escape(marca)}(\S*)(.*)$", t[mejor[0]:])
        if token and token.group(1):
            resto = (marca + token.group(1) + " " + (token.group(2) or "")).strip()
        etiqueta = marca.title() if marca.isupper() and len(marca) > 3 else marca
        return etiqueta, resto
    partes = t.split()
    while partes and partes[0].lower().rstrip(":") in _STOP_MARCA:
        partes = partes[1:]
    if not partes:
        return "", t0
    conocidos = {m.upper() for m in _marcas_conocidas()}
    cabeza = re.sub(r"[^\wÁÉÍÓÚÑáéíóúñ\-]", "", partes[0])
    if cabeza.upper() in conocidos:
        return partes[0], " ".join(partes[1:])
    return "", t0


def _parece_producto_vehiculo(titulo, etiqueta=""):
    blob = f"{etiqueta} {titulo}"
    if _RE_ACCESORIO.search(blob):
        return False
    marca, _ = partir_marca_modelo(titulo)
    if marca:
        return True
    return clasificar_item(titulo, etiqueta) in _CATS_VEHICULO


def _item_dict(titulo, cantidad, unitario, importe, etiqueta=""):
    categoria = clasificar_item(titulo, etiqueta)
    if (
        categoria == "Accesorio"
        and (importe or unitario or 0) >= 8000
        and _parece_producto_vehiculo(titulo, etiqueta)
    ):
        categoria = "Coche"
    tipo_item = normalizar_tipo(etiqueta or titulo)
    if tipo_item == "OTRO" and categoria == "Coche":
        tipo_item = "AUTOMOVIL"
    if tipo_item == "OTRO" and categoria in ("Van", "Scooter", "Bus", "Camión"):
        tipo_item = {
            "Van": "VAN", "Scooter": "SCOOTER", "Bus": "BUS", "Camión": "CAMION",
        }[categoria]
    marca, modelo = partir_marca_modelo(titulo)
    valor = importe or unitario or 0
    return {
        "Marca": marca or "",
        "Modelo": modelo or titulo,
        "Descripcion": titulo[:180],
        "Cantidad": cantidad or 1,
        "Valor_Unitario": unitario or valor,
        "Valor": valor,
        "Categoria_Item": categoria,
        "Tipo": tipo_item,
        "Etiqueta": etiqueta,
    }


def tipo_desde_catalogo(marca, modelo, texto=""):
    """Carrocería del catálogo local. Oracle, si está configurado, tiene prioridad."""
    try:
        from oracle_db import carroceria_oracle
        remoto = carroceria_oracle(modelo)
        if remoto:
            return normalizar_tipo(remoto)
    except Exception:
        pass
    nombre = str(modelo or "").strip()
    if nombre:
        try:
            with Session(engine_catalogo) as session:
                fila = session.execute(
                    text(
                        "SELECT tipo_carroceria FROM modelo "
                        "WHERE UPPER(nombre) = UPPER(:n)"
                    ),
                    {"n": nombre},
                ).first()
                if fila and fila[0]:
                    return normalizar_tipo(fila[0])
        except Exception:
            pass
    return normalizar_tipo(" ".join(x for x in (marca, modelo, texto[:400] if texto else "") if x))


def _importe(token):
    raw = str(token or "").strip()
    neg = bool(re.match(r"^-", raw.replace(" ", ""))) or bool(re.search(r"^\s*-\s*(?:US\$|USD|\$)", raw, flags=re.I))
    s = raw.upper().replace("USD", "").replace("US$", "").replace("CLP", "").replace("$", "").strip()
    s = s.replace("−", "-").lstrip("-").strip()
    if re.fullmatch(r"\d{1,3}(?:\.\d{3})+,\d{2}", s):
        s = s.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"\d{1,3}(?:,\d{3})+\.\d{2}", s):
        s = s.replace(",", "")
    elif re.fullmatch(r"\d{1,3}(?:\.\d{3})+", s):
        s = s.replace(".", "")
    elif re.fullmatch(r"\d+,\d{2}", s):
        s = s.replace(",", ".")
    else:
        s = s.replace(",", "")
    try:
        val = float(s)
    except ValueError:
        return 0.0
    return -val if neg and val > 0 else val


def extraer_items(texto, respaldo=None):
    """Filas reales de producto: marca/modelo/cantidad/costo. Sirve nativo y OCR."""
    lineas = [ln.strip() for ln in (texto or "").splitlines()]
    precios = [i for i, ln in enumerate(lineas) if _es_linea_precio(ln)]
    items = []
    previo = 0
    k = 0
    while k < len(precios):
        a = precios[k]
        es_par = k + 1 < len(precios) and precios[k + 1] == a + 1
        b = precios[k + 1] if es_par else a
        bloque = lineas[previo:a]
        previo = b + 1
        k += 2 if es_par else 1
        unitario = _importe(lineas[a])
        importe = _importe(lineas[b]) if es_par else unitario
        if unitario < 0 or importe < 0:
            continue
        valor = abs(importe) if importe else abs(unitario)
        if valor <= 0:
            continue
        if _bloque_es_total(bloque):
            continue
        recorte = _recortar_bloque_item(bloque)
        blob = " ".join(recorte[-8:])
        if _es_nota_no_producto(blob) or _es_nota_no_producto(" ".join(recorte)):
            continue
        if valor < 80 and not _parece_micromovilidad(blob):
            continue
        if re.search(r"(?i)costo\s+de\s+carga|tarifa\s+residencial|aproximada", blob) and valor < 100:
            continue
        titulo, etiqueta, cantidad = _elegir_producto(recorte, valor)
        if not titulo:
            continue
        if _es_nota_no_producto(titulo) or _RE_FILA_TOTAL.search(titulo) or _RE_NO_PRODUCTO.match(titulo):
            continue
        if re.search(
            r"(?i)\b(?:infraestructura de recarga|se incluye|cl[aá]usulas?|"
            r"beneficios fiscales|gravamen|exonerad|homologad|tarifa\s+0\s*%|"
            r"sin restricci[oó]n|acreditar|unidades homologadas|"
            r"mantenimientos?\s+preventivos)\b",
            titulo,
        ):
            continue
        if titulo.endswith(":") and len(titulo.split()) <= 5:
            continue
        if _score_titulo_item(titulo, _marcas_conocidas()) < 1 and len(titulo.split()) < 4:
            continue
        items.append(_item_dict(titulo, cantidad, abs(unitario), abs(importe or unitario), etiqueta))

    if not items:
        items = _items_misma_linea(lineas)

    vistos = set()
    limpios = []
    for it in items:
        clave = (
            str(it.get("Marca") or "").lower(),
            str(it.get("Modelo") or it.get("Descripcion") or "").lower()[:40],
            round(float(it.get("Valor") or 0), 2),
        )
        if clave in vistos:
            continue
        vistos.add(clave)
        limpios.append(it)
    items = limpios

    if items:
        return items
    respaldo = respaldo or {}
    nombre = " ".join(x for x in (respaldo.get("marca"), respaldo.get("modelo")) if x).strip()
    if not nombre:
        return []
    return [_item_dict(
        nombre,
        respaldo.get("cantidad") or 1,
        float(respaldo.get("costo_mas_alto") or respaldo.get("costo_total") or 0),
        float(respaldo.get("costo_mas_alto") or respaldo.get("costo_total") or 0),
    )]


def _items_misma_linea(lineas):
    items = []
    vistos = set()
    for linea in lineas:
        if len(linea) < 4 or not re.search(r"(?:\$|USD|CLP)\s*-?\s*\d", linea, flags=re.I):
            continue
        if re.search(r"(?i)\b(?:subtotal|iva|ice|descuento|bono|total|monto\s+final|costo\s+de\s+carga)\b", linea):
            continue
        if _es_nota_no_producto(linea):
            continue
        montos = re.findall(r"(?:US\$|USD|CLP|\$)\s*-?\s*[\d.,]+", linea, flags=re.I)
        if not montos:
            continue
        valor = max(_importe(m) for m in montos)
        if valor < 80 and not _parece_micromovilidad(linea):
            continue
        if valor <= 0:
            continue
        desc = re.split(r"(?i)(?:US\$|USD|CLP|\$)", linea)[0]
        desc = re.sub(r"^\d+\s+", "", desc).strip(" .-:|•")
        desc = _limpiar_titulo_item(desc)
        if len(desc) < 3 or _RE_NO_PRODUCTO.match(desc) or _RE_CATEGORIA_SOLA.match(desc):
            continue
        if _es_nota_no_producto(desc) or _RE_FILA_TOTAL.search(desc):
            continue
        if _score_titulo_item(desc, _marcas_conocidas()) < 1:
            continue
        clave = (desc.lower(), round(valor, 2))
        if clave in vistos:
            continue
        vistos.add(clave)
        items.append(_item_dict(desc, 1, valor, valor))
    return items


def _folio(datos):
    folio = datos.get("id_documento") or datos.get("factura") or ""
    if not folio:
        folio = os.path.splitext(str(datos.get("archivo_guardado") or datos.get("archivo") or "factura"))[0]
    folio = re.sub(r"[^\w\-]+", "_", str(folio))[:28] or "factura"
    return folio


def item_vehicular_mas_caro(items):
    """Ítem vehicular de mayor Valor; ignora notas, cargadores y precios OCR disparatados."""
    filas = [it for it in (items or []) if float(it.get("Valor") or 0) > 0]
    if not filas:
        return None
    vehiculos = [it for it in filas if it.get("Categoria_Item") in _CATS_VEHICULO]
    if not vehiculos:
        return None
    vals = sorted(float(it.get("Valor") or 0) for it in vehiculos)
    bajos = [v for v in vals if v < 200_000]
    altos = [v for v in vals if v >= 1_000_000]
    if bajos and altos and len(bajos) >= len(altos):
        vehiculos = [it for it in vehiculos if float(it.get("Valor") or 0) < 200_000]
    if not vehiculos:
        return None
    return max(vehiculos, key=lambda it: float(it.get("Valor") or 0))


def aplicar_item_principal(datos, texto=""):
    """Marca, modelo, cantidad y costo_mas_alto salen del ítem vehicular más caro."""
    out = dict(datos or {})
    items = out.get("items") or extraer_items(texto, out)
    out["items"] = items
    principal = item_vehicular_mas_caro(items)
    mapa_tipo = {
        "Coche": "AUTOMOVIL",
        "Scooter": "SCOOTER",
        "Moto": "SCOOTER",
        "Bus": "BUS",
        "Camión": "CAMION",
        "Van": "VAN",
    }
    if not principal:
        out["item_principal"] = " ".join(x for x in (out.get("marca"), out.get("modelo")) if x)
        out["categoria_item"] = "Coche" if out.get("marca") else ""
        out["valor_item"] = out.get("costo_mas_alto") or out.get("costo_total") or 0
        out["tipo"] = tipo_desde_catalogo(out.get("marca"), out.get("modelo"), texto)
        return out
    valor = float(principal.get("Valor") or 0)
    if valor > 0:
        out["costo_mas_alto"] = valor
    marca = str(principal.get("Marca") or "").strip()
    modelo = str(principal.get("Modelo") or "").strip()
    if marca:
        out["marca"] = marca
    if modelo:
        out["modelo"] = modelo
    try:
        cant = int(float(principal.get("Cantidad") or 1)) or 1
    except (TypeError, ValueError):
        cant = 1
    out["cantidad"] = cant
    out["item_principal"] = principal.get("Descripcion") or ""
    out["categoria_item"] = principal.get("Categoria_Item")
    out["valor_item"] = valor
    tipo = tipo_desde_catalogo(
        out.get("marca"),
        out.get("modelo"),
        principal.get("Descripcion") or "",
    )
    if tipo == "OTRO":
        tipo = mapa_tipo.get(principal.get("Categoria_Item"), tipo)
    out["tipo"] = tipo
    return out


def completar_resultado(datos, texto=""):
    """Añade ítems y alinea el resumen al vehículo más caro."""
    out = aplicar_item_principal(datos, texto)
    out.setdefault("fecha_ejecucion", fecha_hoy())
    return out


def fila_resumen(datos):
    return {
        "Estado": datos.get("Estado") or estado_excel(datos),
        "Fecha_Ejecucion": datos.get("fecha_ejecucion") or fecha_hoy(),
        "Marca": datos.get("marca") or "",
        "Modelo": datos.get("modelo") or "",
        "cantidad": datos.get("cantidad") or 1,
        "beneficiario": datos.get("beneficiario") or "",
        "ci": datos.get("ci") or "",
        "costo_mas_alto": datos.get("costo_mas_alto") or 0,
        "costo_total": datos.get("costo_total") or 0,
        "marca_base": datos.get("marca_base") or "",
        "modelo_base": datos.get("modelo_base") or "",
        "Tipo": normalizar_tipo(datos.get("tipo")),
        "archivo": datos.get("archivo") or datos.get("archivo_guardado") or "",
    }


def _detalle_frame(reg):
    filas = list(_filas_detalle(reg, hist=False))
    return pd.DataFrame(filas, columns=list(COLUMNAS_DETALLE))


def _filas_detalle(reg, hist=False):
    filas = []
    extras = {}
    if hist:
        extras = {
            "Fecha_Ejecucion": reg.get("fecha_ejecucion") or reg.get("Fecha_Ejecucion") or fecha_hoy(),
            "archivo": reg.get("archivo") or reg.get("archivo_guardado") or "",
        }
    for it in reg.get("items") or []:
        marca = it.get("Marca") or ""
        modelo = it.get("Modelo") or ""
        if not marca or not modelo:
            pm, pmod = partir_marca_modelo(it.get("Descripcion"))
            marca = marca or pm
            modelo = modelo or pmod or it.get("Descripcion") or ""
        fila = {
            "Marca": marca,
            "Modelo": modelo,
            "cantidad": it.get("Cantidad") or it.get("cantidad") or 1,
            "beneficiario": it.get("beneficiario") or reg.get("beneficiario") or "",
            "ci": it.get("ci") or reg.get("ci") or "",
            "costo": it.get("Valor") or it.get("costo") or 0,
        }
        fila.update(extras)
        filas.append(fila)
    if not filas:
        fila = {
            "Marca": reg.get("marca") or reg.get("Marca") or "",
            "Modelo": reg.get("modelo") or reg.get("Modelo") or "",
            "cantidad": reg.get("cantidad") or 1,
            "beneficiario": reg.get("beneficiario") or "",
            "ci": reg.get("ci") or "",
            "costo": reg.get("costo_mas_alto") or reg.get("costo_total") or 0,
        }
        fila.update(extras)
        filas.append(fila)
    return filas


def _nombre_hoja(folio, usados):
    base = re.sub(r"[\[\]\:\*\?\/\\]", "_", str(folio))[:28] or "factura"
    nombre = base
    n = 2
    while nombre in usados:
        nombre = f"{base[:24]}_{n}"
        n += 1
    usados.add(nombre)
    return nombre


def excel_exitos(registros):
    buffer = io.BytesIO()
    resumen = pd.DataFrame([fila_resumen(r) for r in registros], columns=list(COLUMNAS_RESUMEN))
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
        usados = {"Resumen"}
        for reg in registros:
            hoja = _nombre_hoja(_folio(reg), usados)
            _detalle_frame(reg).to_excel(writer, sheet_name=hoja, index=False)
    buffer.seek(0)
    return buffer.getvalue()


def excel_revision(registros):
    buffer = io.BytesIO()
    frame = pd.DataFrame([fila_resumen(r) for r in registros], columns=list(COLUMNAS_RESUMEN))
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        frame.to_excel(writer, sheet_name="Resumen", index=False)
        usados = {"Resumen"}
        for reg in registros:
            hoja = _nombre_hoja(_folio(reg), usados)
            _detalle_frame(reg).to_excel(writer, sheet_name=hoja, index=False)
    buffer.seek(0)
    return buffer.getvalue()


def _es_exito(reg):
    return "ÉXITO" in str(reg.get("estado", "")) or str(reg.get("Estado", "")).upper().startswith("EXITO")


def escribir_reciente(registros):
    """Excel + JSON de la corrida actual, sin tocar el histórico."""
    os.makedirs(DIR_RECIENTE, exist_ok=True)
    exitos = [r for r in registros if _es_exito(r)]
    revision = [r for r in registros if not _es_exito(r)]
    with open(ARCHIVO_CORRIDA_JSON, "w", encoding="utf-8") as f:
        json.dump(registros, f, ensure_ascii=False, default=str, indent=2)
    _escribir_binario_reintento(ARCHIVO_EXITOS_RECIENTE, excel_exitos(exitos))
    _escribir_binario_reintento(ARCHIVO_REVISION_RECIENTE, excel_revision(revision))
    return ARCHIVO_EXITOS_RECIENTE, ARCHIVO_REVISION_RECIENTE


def _es_fallo_archivo(reg):
    texto = " ".join(str((reg or {}).get(k) or "") for k in ("estado", "Estado", "error"))
    if not re.search(r"(?i)permission denied", texto):
        return False
    return not (reg or {}).get("marca") and not (reg or {}).get("ci") and not (reg or {}).get("Marca")


def anexar_corrida(registros):
    """Añade documentos de un archivo. El histórico solo recibe estos nuevos."""
    actuales = cargar_corrida()
    nuevos = [r for r in (registros or []) if not _es_fallo_archivo(r)]
    escribir_reciente(actuales + nuevos)
    exitos = [r for r in nuevos if _es_exito(r)]
    revision = [r for r in nuevos if not _es_exito(r)]
    if exitos or revision:
        try:
            _acumular_historico(exitos, revision)
        except Exception as e:
            print(f"[WARN] Histórico no actualizado ({e}). La corrida JSON se conservó.")
    return ARCHIVO_EXITOS_RECIENTE, ARCHIVO_REVISION_RECIENTE


def iniciar_corrida():
    """Vacía la corrida reciente. Solo al empezar un lote nuevo."""
    os.makedirs(DIR_RECIENTE, exist_ok=True)
    escribir_reciente([])
    return []


def guardar_corrida(registros):
    """Reescribe la corrida reciente (edición). No toca el histórico."""
    return escribir_reciente(registros)


def cargar_corrida():
    if not os.path.exists(ARCHIVO_CORRIDA_JSON):
        return []
    try:
        with open(ARCHIVO_CORRIDA_JSON, encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, list):
            return []
        return [r for r in data if not _es_fallo_archivo(r)]
    except Exception:
        return []


def _escribir_binario_reintento(ruta, contenido, intentos=5):
    ultimo = None
    for i in range(intentos):
        try:
            with open(ruta, "wb") as f:
                f.write(contenido)
            return
        except PermissionError as e:
            ultimo = e
            time.sleep(0.4 * (i + 1))
        except OSError as e:
            ultimo = e
            time.sleep(0.4 * (i + 1))
    print(f"[WARN] No se pudo escribir {ruta}: {ultimo}")


def _escribir_xlsx_reintento(ruta, construir, intentos=5):
    ultimo = None
    for i in range(intentos):
        try:
            construir(ruta)
            return True
        except PermissionError as e:
            ultimo = e
            time.sleep(0.4 * (i + 1))
        except OSError as e:
            ultimo = e
            time.sleep(0.4 * (i + 1))
    print(f"[WARN] No se pudo escribir {ruta}: {ultimo}")
    return False


def _alinear_columnas(df, hoja=""):
    out = df.copy() if df is not None and not df.empty else pd.DataFrame()
    if "Estado" not in out.columns:
        if hoja.lower().startswith("exito"):
            out["Estado"] = "EXITO"
        elif hoja.lower().startswith("rev"):
            out["Estado"] = "REVISION"
        else:
            out["Estado"] = ""
    for c in COLUMNAS_RESUMEN:
        if c not in out.columns:
            out[c] = ""
    out = out[list(COLUMNAS_RESUMEN)]
    if out.empty:
        return out
    marca = out["Marca"].astype(str).str.strip().replace({"nan": ""})
    est = out["Estado"].astype(str)
    basura = est.str.contains("Permission denied", case=False, na=False) & marca.eq("")
    return out.loc[~basura].reset_index(drop=True)


def _alinear_detalle(df):
    out = df.copy() if df is not None and not df.empty else pd.DataFrame()
    for c in COLUMNAS_DETALLE_HIST:
        if c not in out.columns:
            out[c] = ""
    return out[list(COLUMNAS_DETALLE_HIST)]


def _detalle_hist_frame(registros):
    filas = []
    for reg in registros or []:
        filas.extend(_filas_detalle(reg, hist=True))
    return pd.DataFrame(filas, columns=list(COLUMNAS_DETALLE_HIST))


def _acumular_historico(exitos, revision):
    prev_ex, prev_re, prev_dx, prev_dr = _leer_historico()
    add_ex = pd.DataFrame([fila_resumen(r) for r in exitos], columns=list(COLUMNAS_RESUMEN))
    add_re = pd.DataFrame([fila_resumen(r) for r in revision], columns=list(COLUMNAS_RESUMEN))
    add_dx = _detalle_hist_frame(exitos)
    add_dr = _detalle_hist_frame(revision)
    ex = pd.concat([prev_ex, add_ex], ignore_index=True)
    re = pd.concat([prev_re, add_re], ignore_index=True)
    dx = pd.concat([prev_dx, add_dx], ignore_index=True)
    dr = pd.concat([prev_dr, add_dr], ignore_index=True)

    def _write(ruta):
        with pd.ExcelWriter(ruta, engine="openpyxl") as writer:
            _alinear_columnas(ex, "Exitos").to_excel(writer, sheet_name="Exitos", index=False)
            _alinear_columnas(re, "Revision").to_excel(writer, sheet_name="Revision", index=False)
            _alinear_detalle(dx).to_excel(writer, sheet_name="Detalle_Exitos", index=False)
            _alinear_detalle(dr).to_excel(writer, sheet_name="Detalle_Revision", index=False)

    if not _escribir_xlsx_reintento(ARCHIVO_HISTORICO, _write):
        raise PermissionError(f"No se pudo actualizar {ARCHIVO_HISTORICO}")


def _leer_historico():
    vacio = pd.DataFrame(columns=list(COLUMNAS_RESUMEN))
    vacio_d = pd.DataFrame(columns=list(COLUMNAS_DETALLE_HIST))
    if not os.path.exists(ARCHIVO_HISTORICO):
        return vacio.copy(), vacio.copy(), vacio_d.copy(), vacio_d.copy()
    try:
        ex = _alinear_columnas(pd.read_excel(ARCHIVO_HISTORICO, sheet_name="Exitos"), "Exitos")
    except Exception:
        ex = vacio.copy()
    try:
        re = _alinear_columnas(pd.read_excel(ARCHIVO_HISTORICO, sheet_name="Revision"), "Revision")
    except Exception:
        re = vacio.copy()
    try:
        dx = _alinear_detalle(pd.read_excel(ARCHIVO_HISTORICO, sheet_name="Detalle_Exitos"))
    except Exception:
        dx = vacio_d.copy()
    try:
        dr = _alinear_detalle(pd.read_excel(ARCHIVO_HISTORICO, sheet_name="Detalle_Revision"))
    except Exception:
        dr = vacio_d.copy()
    return ex, re, dx, dr


def cargar_historico():
    ex, re, _, _ = _leer_historico()
    return ex, re


def _items_de_detalle(df_det, fila):
    sub = None if df_det is None or df_det.empty else df_det
    items = []
    if sub is not None:
        arch = str(fila.get("archivo") or "")
        fecha = str(fila.get("Fecha_Ejecucion") or "")
        if arch and "archivo" in sub.columns:
            matched = sub[sub["archivo"].astype(str) == arch]
            sub = matched if not matched.empty else None
        elif not arch:
            sub = None
        if sub is not None and fecha and "Fecha_Ejecucion" in sub.columns:
            matched = sub[sub["Fecha_Ejecucion"].astype(str) == fecha]
            if not matched.empty:
                sub = matched
        if sub is not None:
            for rec in sub.fillna("").to_dict(orient="records"):
                items.append({
                    "Marca": rec.get("Marca") or "",
                    "Modelo": rec.get("Modelo") or "",
                    "Cantidad": rec.get("cantidad") or 1,
                    "cantidad": rec.get("cantidad") or 1,
                    "beneficiario": rec.get("beneficiario") or "",
                    "ci": rec.get("ci") or "",
                    "Valor": rec.get("costo") or 0,
                    "costo": rec.get("costo") or 0,
                })
    if items:
        return items
    return [{
        "Marca": fila.get("Marca") or "",
        "Modelo": fila.get("Modelo") or "",
        "Cantidad": fila.get("cantidad") or 1,
        "cantidad": fila.get("cantidad") or 1,
        "beneficiario": fila.get("beneficiario") or "",
        "ci": fila.get("ci") or "",
        "Valor": fila.get("costo_mas_alto") or fila.get("costo_total") or 0,
        "costo": fila.get("costo_mas_alto") or fila.get("costo_total") or 0,
    }]


def _fila_a_registro(fila, items, estado_def):
    rec = dict(fila)
    rec["items"] = items
    rec["archivo_guardado"] = rec.get("archivo") or ""
    rec["marca"] = rec.get("Marca") or ""
    rec["modelo"] = rec.get("Modelo") or ""
    rec["tipo"] = rec.get("Tipo") or ""
    rec["fecha_ejecucion"] = rec.get("Fecha_Ejecucion") or ""
    rec["estado"] = rec.get("Estado") or estado_def
    rec["Estado"] = rec.get("Estado") or estado_def
    return rec


def registros_historico():
    ex, re, dx, dr = _leer_historico()
    exitos = [
        _fila_a_registro(f, _items_de_detalle(dx, f), "EXITO")
        for f in ex.fillna("").to_dict(orient="records")
    ]
    revision = [
        _fila_a_registro(f, _items_de_detalle(dr, f), "REVISION")
        for f in re.fillna("").to_dict(orient="records")
    ]
    return exitos, revision


def modelo_en_catalogo(modelo):
    nombre = str(modelo or "").strip()
    if not nombre:
        return False
    try:
        from oracle_db import modelo_en_oracle
        remoto = modelo_en_oracle(nombre)
        if remoto is not None:
            return remoto
    except Exception:
        pass
    try:
        with Session(engine_catalogo) as session:
            fila = session.execute(
                text("SELECT 1 FROM modelo WHERE UPPER(nombre) = UPPER(:n)"),
                {"n": nombre},
            ).first()
            return fila is not None
    except Exception:
        return False


def registrar_modelo(marca, modelo, carroceria):
    """Alta en el catálogo local y, si hay credenciales, en Oracle."""
    tipo = normalizar_tipo(carroceria)
    marca_n = str(marca or "").strip().upper()
    modelo_n = str(modelo or "").strip().upper()
    if not marca_n or not modelo_n:
        raise ValueError("Marca y modelo son obligatorios.")
    with Session(engine_catalogo) as session:
        fila = session.execute(
            text("SELECT id_marca FROM marca WHERE UPPER(nombre) = :n"),
            {"n": marca_n},
        ).first()
        if fila:
            id_marca = fila[0]
        else:
            session.execute(
                text("INSERT INTO marca (nombre, origen) VALUES (:n, :o)"),
                {"n": marca_n, "o": ""},
            )
            session.commit()
            id_marca = session.execute(
                text("SELECT id_marca FROM marca WHERE UPPER(nombre) = :n"),
                {"n": marca_n},
            ).first()[0]
        existe = session.execute(
            text("SELECT 1 FROM modelo WHERE id_marca = :m AND UPPER(nombre) = :n"),
            {"m": id_marca, "n": modelo_n},
        ).first()
        if not existe:
            session.execute(
                text(
                    "INSERT INTO modelo (id_marca, nombre, tipo_carroceria) "
                    "VALUES (:m, :n, :t)"
                ),
                {"m": id_marca, "n": modelo_n, "t": tipo},
            )
            session.commit()
    aviso_oracle = ""
    try:
        from oracle_db import insertar_modelo_oracle
        aviso_oracle = insertar_modelo_oracle(marca_n, modelo_n, tipo)
    except Exception as exc:
        aviso_oracle = str(exc)
    return tipo, aviso_oracle
