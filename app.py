"""
Corrector de Documentos Coca-Cola (importación)
Controla: Documento de transporte, COA, Certificado de origen (NALADISA), Factura y Packing List.

Estrategia:
  1) La IA (Groq, con respaldo Anthropic) SOLO extrae datos de cada PDF a JSON.
  2) Las reglas se aplican en Python (determinístico) -> resultados repetibles.
  3) Se genera un Excel con ✔ / ✘ / — y el motivo.

Hoy implementado: modo MARÍTIMO + Certificado NALADISA.
Aéreo / Terrestre / CO Mercosur: pendientes (se agregan en TRANSPORT_RULES / CERT_RULES).
"""
import io
import json
import os
import re
import unicodedata
from datetime import date, datetime, timedelta
from difflib import SequenceMatcher

import pandas as pd
import streamlit as st

try:
    from zoneinfo import ZoneInfo
    HOY = datetime.now(ZoneInfo("America/Argentina/Buenos_Aires")).date()
except Exception:  # pragma: no cover
    HOY = date.today()

# ----------------------------------------------------------------------------
# Constantes de negocio
# ----------------------------------------------------------------------------
RAZON_SOCIAL = "SERVICIOS Y PRODUCTOS PARA BEBIDAS REFRESCANTES S.R.L"
RAZON_CLAVE = "SERVICIOS Y PRODUCTOS PARA BEBIDAS REFRESCANTES"
CUIT = "30-67773730-8"
DIRECCIONES_BL = ["LUIS VIALE 2045", "VEDIA 4090"]
INCOTERMS = ["EXW", "FCA", "CPT", "CIP", "DAP", "FOB", "CFR", "CIF"]
DIAS_MIN_VENCIMIENTO = 25

DOCS = {
    "transporte": "Documento de transporte",
    "coa": "COA",
    "certificado": "Certificado de origen",
    "factura": "Factura",
    "packing": "Packing List",
}

OK, ERR, NA = "✔", "✘", "—"

# ----------------------------------------------------------------------------
# Esquemas de extracción (lo que se le pide a la IA)
# ----------------------------------------------------------------------------
NUM_NOTE = ("IMPORTANTE: los pesos, cantidades e importes devolvelos como TEXTO exactamente como están impresos (ej. \"1 036,000\"), sin convertirlos. "
            "Las fechas en formato YYYY-MM-DD (si es ambigua, asumir día/mes/año). "
            "El texto del PDF puede venir con dos columnas intercaladas (ej. Proveedor/Cliente o Recipient/Customer) y los nombres pueden cortarse en dos renglones: "
            "reconstruí el dato completo SOLO si la continuación está realmente en el texto; no completes ni corrijas nada que no esté escrito.")

SCHEMAS = {
    "factura": {
        "desc": "Factura comercial / Invoice",
        "campos": {
            "numero_factura": "número de factura/invoice",
            "fecha_factura": "fecha de la factura",
            "po": "número de orden de compra / PO",
            "incoterm": "incoterm informado (ej. FOB, CIF) o null si no figura",
            "forma_pago": "forma de pago / payment terms o null si no figura",
            "moneda": "moneda / currency o null si no figura",
            "vendedor_razon_social": "vendedor / seller / exportador",
            "vendedor_direccion": "dirección del vendedor",
            "comprador_razon_social": "comprador / buyer / consignatario / sold to",
            "comprador_cuit": "CUIT/tax id del comprador",
            "comprador_direccion": "dirección del comprador",
            "descripcion_producto": "descripción de TODOS los productos, concatenada en un solo texto",
            "cantidad": "cantidad total en la unidad de medida de la factura (ej. KG), tal cual impresa",
            "peso_bruto_kg": "peso BRUTO total en kg si figura, tal cual impreso, o null",
            "peso_neto_kg": "peso NETO total en kg si figura, tal cual impreso, o null",
            "valor_total": "valor total de la factura, tal cual impreso",
            "codigo_item": "código de ítem del CLIENTE: número de 6-8 dígitos que acompaña a \"Coca Cola spec.\" (ej. 1003135). NO es el código de artículo/abreviatura del proveedor (ej. 006062-0003). null si no figura",
            "contrato": "número de contrato si figura o null",
        },
    },
    "packing": {
        "desc": "Packing List",
        "campos": {
            "numero_packing": "número de packing list",
            "cliente_razon_social": "cliente / consignatario / customer / consignee",
            "cliente_cuit": "CUIT/tax id del cliente",
            "po": "número de orden / PO",
            "descripcion_producto": "descripción de TODOS los productos, concatenada en un solo texto",
            "peso_bruto_kg": "peso bruto TOTAL en kg (columna Total/Gross weight), tal cual impreso",
        },
    },
    "coa": {
        "desc": "COA (Certificate of Analysis)",
        "campos": {
            "po": "número de PO / orden de compra o null si no figura",
            "descripcion_producto": "descripción / nombre del producto o null si no figura",
            "fecha_vencimiento": "fecha de vencimiento (expiry / best before / retest date / use by); si hay varias, la MÁS PRÓXIMA",
        },
    },
    "certificado": {
        "desc": "Certificado de origen NALADISA",
        "campos": {
            "exportador_razon_social": "razón social del exportador o productor",
            "exportador_direccion": "dirección del exportador o productor",
            "importador_razon_social": "razón social del importador",
            "importador_direccion": "dirección del importador",
            "denominacion_mercaderias": "denominación de las mercaderías, concatenada",
            "peso_o_cantidad": "peso o cantidad informado, tal cual impreso",
            "valor_fob": "valor FOB, tal cual impreso",
            "declaracion_factura_numero": "número de factura mencionado en la declaración de origen",
            "declaracion_factura_fecha": "fecha de factura mencionada en la declaración de origen",
        },
    },
    # Transporte según modo
    "transporte_maritimo": {
        "desc": "Bill of Lading (BL) marítimo",
        "campos": {
            "consignee_razon_social": "campo CONSIGNED TO / consignee: razón social",
            "consignee_cuit": "CUIT/tax id que figura en el consignee",
            "consignee_direccion": "dirección que figura en el consignee",
            "gross_weight_kg": "GROSS WEIGHT total en kg, tal cual impreso",
            "items": "LISTA (una entrada por cada factura mencionada en DESCRIPTION OF GOODS, incluida la hoja de continuación) de objetos con claves: numero_factura, po, codigo_item, descripcion (texto tal cual figura, sin corregir errores de tipeo)",
        },
    },
}

# ----------------------------------------------------------------------------
# Utilidades de texto / números / fechas
# ----------------------------------------------------------------------------
def norm(s):
    if s is None:
        return ""
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c)).upper()
    s = re.sub(r"[^A-Z0-9 ]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def alnum(s):
    return re.sub(r"[^A-Z0-9]", "", norm(s))


def digits(s):
    return re.sub(r"\D", "", str(s or ""))


def parse_num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d,.\-]", "", str(v))
    if not s:
        return None
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".") if re.search(r",\d{1,2}$", s) else s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def num_cands(v):
    """Interpretaciones posibles de un número impreso (decimal primero, miles después)."""
    if v is None or v == "":
        return []
    if isinstance(v, (int, float)):
        return [float(v)]
    s = re.sub(r"[^\d,.\-]", "", str(v))
    if not s:
        return []
    try:
        if "," in s and "." in s:
            return [parse_num(s)]
        sep = "," if "," in s else "." if "." in s else None
        if sep is None:
            return [float(s)]
        if s.count(sep) > 1:
            return [float(s.replace(sep, ""))]
        dec = float(s.replace(sep, "."))
        if re.search(r"\%s\d{3}$" % sep, s):
            return [dec, float(s.replace(sep, ""))]
        return [dec]
    except ValueError:
        return []


def pick_num(v, ref=None):
    c = num_cands(v)
    if not c:
        return None
    if ref is not None:
        r0 = parse_num(ref)
        if r0 is not None:
            c = c + [x / 1000 for x in c]  # por si la IA convirtió "1 036,000" como 1036000
            return min(c, key=lambda x: abs(x - r0))
    return c[0]


def eq_num(a, b, **kw):
    return any(num_equal(x, y, **kw) for x in num_cands(a) for y in num_cands(b))


def parse_date(v):
    if not v:
        return None
    s = str(v).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s[:10] if fmt.startswith("%Y") else s, fmt).date()
        except ValueError:
            continue
    return None


def num_equal(a, b, rel=0.001, abs_tol=0.05):
    return abs(a - b) <= max(abs_tol, rel * max(abs(a), abs(b)))


STOP = {"DE", "DEL", "LA", "EL", "Y", "CON", "PARA", "THE", "AND", "OF", "WITH", "FOR", "X", "EN", "A"}


def tokens(s):
    return {t for t in norm(s).split() if t not in STOP and len(t) > 1}


def desc_match(a, b, umbral=0.7):
    """Proporción de tokens del texto más corto contenidos en el más largo."""
    ta, tb = tokens(a), tokens(b)
    if not ta or not tb:
        return False, 0.0
    corto, largo = (ta, tb) if len(ta) <= len(tb) else (tb, ta)
    r = len(corto & largo) / len(corto)
    return r >= umbral, r


def name_match(a, b):
    na, nb = norm(a), norm(b)
    if not na or not nb:
        return False
    if na in nb or nb in na:
        return True
    return SequenceMatcher(None, na, nb).ratio() >= 0.9


def id_match(a, b):
    xa, xb = alnum(a), alnum(b)
    return bool(xa and xb and (xa == xb or xa in xb or xb in xa))


def show(v):
    return "" if v is None else str(v)


# ----------------------------------------------------------------------------
# Motor de controles (soporta varias facturas / COA / packing por operación)
# ----------------------------------------------------------------------------
def typos(a, b):
    """Palabras casi iguales pero distintas (posibles errores de tipeo)."""
    ta, tb = tokens(a), tokens(b)
    out = []
    for x in ta - tb:
        for y in tb - ta:
            if x.isalpha() and y.isalpha() and len(x) >= 6 and SequenceMatcher(None, x, y).ratio() >= 0.85:
                out.append(f"'{x}' vs '{y}'")
    return out


def desc_check(a, b, umbral=0.7):
    ok, ratio = desc_match(a, b, umbral)
    t = typos(a, b)
    if t:
        return False, f"Posible error de tipeo en la descripción: {', '.join(t)}."
    if not ok:
        return False, f"La descripción no coincide con la factura (coincidencia {ratio:.0%}). Revisar."
    return True, ""


class Resultados:
    def __init__(self):
        self.filas = []

    def add(self, doc, control, encontrado, esperado, estado, motivo=""):
        self.filas.append(
            {"Documento": doc, "Control": control, "En el documento": show(encontrado),
             "Esperado / Referencia": show(esperado), "Resultado": estado, "Motivo": motivo}
        )

    def check(self, doc, control, encontrado, esperado, cond, motivo_error):
        self.add(doc, control, encontrado, esperado, OK if cond else ERR, "" if cond else motivo_error)

    def err(self, doc, control, encontrado, esperado, motivo):
        self.add(doc, control, encontrado, esperado, ERR, motivo)

    def na(self, doc, control, motivo):
        self.add(doc, control, "", "", NA, motivo)


def lbl(base, num):
    return f"{base} {num}" if num else base


def buscar_factura(facs, po=None, numero=None):
    for f in facs:
        if numero and id_match(numero, f.get("numero_factura")):
            return f
        if po and id_match(po, f.get("po")):
            return f
    return None


def sin_factura(r, d, controles, hay_facturas, po=None):
    motivo = ("No verificable: falta la Factura." if not hay_facturas
              else f"No se encontró ninguna factura cargada con la PO {po}." if po else "No se encontró la factura correspondiente.")
    for c in controles:
        (r.na if not hay_facturas else lambda *a: r.err(a[0], a[1], "", "", a[2]))(d, c, motivo)


def razon_ok(nombre, texto=None, usar_texto=False):
    """Razón social completa; en packing/factura (2 columnas) se acepta si todas las palabras están en el texto."""
    if RAZON_CLAVE in norm(nombre):
        return True
    if usar_texto and texto:
        tk = tokens(texto)
        return all(w in tk for w in RAZON_CLAVE.split() if w not in STOP)
    return False


def cuit_ok(cuit, texto=None):
    return digits(CUIT) in digits(cuit) or bool(texto and digits(CUIT) in digits(texto))


def es_consignatario_ok(r, doc, nombre, cuit, direccion=None, exigir_dir=False, texto=None, usar_texto=False):
    r.check(doc, "Razón social del consignatario/comprador", nombre, RAZON_SOCIAL, razon_ok(nombre, texto, usar_texto),
            "No figura SERVICIOS Y PRODUCTOS PARA BEBIDAS REFRESCANTES S.R.L (razón social incompleta o mal escrita)." if nombre else "No se encontró razón social.")
    r.check(doc, "CUIT", cuit, CUIT, cuit_ok(cuit, texto),
            "No figura el CUIT 30-67773730-8." if cuit else "No figura CUIT.")
    if exigir_dir:
        ok = any(d in norm(direccion) for d in DIRECCIONES_BL)
        r.check(doc, "Dirección del consignatario", direccion, "Luis Viale 2045 o Vedia 4090", ok,
                "La dirección no es Luis Viale 2045 ni Vedia 4090." if direccion else "No se encontró dirección.")


def reglas_transporte_maritimo(r, bl, facs, packs):
    d = DOCS["transporte"] + " (BL)"
    es_consignatario_ok(r, d, bl.get("consignee_razon_social"), bl.get("consignee_cuit"),
                        bl.get("consignee_direccion"), exigir_dir=True, texto=bl.get("_texto"))
    items = bl.get("items") or []
    usados = set()
    for f in facs:
        nf = f.get("numero_factura")
        dd = f"{d} / Factura {nf}"
        it = next((i for i, x in enumerate(items) if id_match(x.get("numero_factura"), nf)), None)
        if it is None:
            r.err(dd, "Factura mencionada en la descripción", "", nf, "El número de factura no figura en la descripción de mercadería del BL.")
            continue
        usados.add(it)
        x = items[it]
        r.check(dd, "Número de factura", x.get("numero_factura"), nf, True, "")
        r.check(dd, "PO vs Factura", x.get("po"), f.get("po"), bool(x.get("po")) and id_match(x.get("po"), f.get("po")),
                "No figura PO en el BL." if not x.get("po") else "La PO no coincide con la factura.")
        ok, mot = desc_check(x.get("descripcion"), f.get("descripcion_producto"))
        r.check(dd, "Descripción de la mercadería vs Factura", x.get("descripcion"), f.get("descripcion_producto"), ok, mot)
        if x.get("codigo_item"):
            cod = alnum(x.get("codigo_item"))
            en_factura = (re.search(r"(?<![0-9A-Z])" + re.escape(cod) + r"(?![0-9A-Z])", re.sub(r"[^A-Z0-9]+", " ", norm(f.get("_texto"))))
                          is not None) if f.get("_texto") else cod == alnum(f.get("codigo_item"))
            r.check(dd, "Código de ítem vs Factura (adicional)", x.get("codigo_item"), f.get("codigo_item") or "(buscado en el texto de la factura)",
                    en_factura, "El código de ítem del BL no figura en la factura (posible error de tipeo).")
    for i, x in enumerate(items):
        if i not in usados:
            r.na(d, f"Factura {x.get('numero_factura')} mencionada en el BL", "El BL menciona una factura que no está cargada.")
    # Peso
    bl_kg = pick_num(bl.get("gross_weight_kg"), sum(parse_num(pick_num(f.get("peso_bruto_kg"))) or 0 for f in facs) or None)
    if not packs:
        r.na(d, "Gross weight vs Packing List", "No verificable: falta el Packing List.")
    elif bl_kg is None:
        r.err(d, "Gross weight vs Packing List", "", "", "No se pudo leer el gross weight del BL. Revisar.")
    else:
        suma = sum(pick_num(p.get("peso_bruto_kg"), (buscar_factura(facs, po=p.get("po")) or {}).get("peso_bruto_kg") and pick_num(buscar_factura(facs, po=p.get("po")).get("peso_bruto_kg"))) or 0 for p in packs)
        sin_pack = [f.get("numero_factura") for f in facs if not any(id_match(p.get("po"), f.get("po")) for p in packs)]
        if sin_pack:
            r.na(d, "Gross weight vs Packing List", f"Parcial: faltan Packing de las facturas {', '.join(map(str, sin_pack))}. "
                 f"Suma de los packings cargados {suma:,.2f} kg vs BL {bl_kg:,.2f} kg.")
        else:
            r.check(d, "Gross weight vs Packing List", bl_kg, suma, num_equal(bl_kg, suma),
                    f"Los kilos no coinciden (BL {bl_kg:,.2f} vs suma de packings {suma:,.2f}).")
    pesos = [pick_num(f.get("peso_bruto_kg")) for f in facs]
    if facs and all(p is not None for p in pesos) and bl_kg is not None:
        t = sum(pesos)
        r.check(d, "Gross weight vs suma de facturas (adicional)", bl_kg, t, num_equal(bl_kg, t),
                f"El gross weight del BL ({bl_kg:,.2f}) no coincide con la suma del peso bruto de las facturas ({t:,.2f}).")


TRANSPORT_RULES = {
    "Marítimo": ("transporte_maritimo", reglas_transporte_maritimo),
    # "Aéreo": pendiente (AWB)
    # "Terrestre": pendiente (CRT)
}


def reglas_factura(r, f):
    d = lbl(DOCS["factura"], f.get("numero_factura"))
    inc = norm(f.get("incoterm"))
    if not inc:
        r.err(d, "Incoterm", "", ", ".join(INCOTERMS), "No detalla incoterm.")
    else:
        r.check(d, "Incoterm", f.get("incoterm"), ", ".join(INCOTERMS), any(t in inc.split() for t in INCOTERMS),
                "El incoterm no es uno de los permitidos.")
    for key, ctrl, msg in (("forma_pago", "Forma de pago / payment term", "No detalla forma de pago."),
                           ("moneda", "Moneda / currency", "No detalla moneda."),
                           ("vendedor_razon_social", "Vendedor", "No detalla vendedor.")):
        r.check(d, ctrl, f.get(key), "Debe figurar", bool(str(f.get(key) or "").strip()), msg)
    r.check(d, "Comprador: razón social", f.get("comprador_razon_social"), RAZON_SOCIAL,
            razon_ok(f.get("comprador_razon_social"), f.get("_texto"), True),
            "El comprador no es SERVICIOS Y PRODUCTOS PARA BEBIDAS REFRESCANTES S.R.L (revisar ortografía).")
    r.check(d, "Comprador: CUIT", f.get("comprador_cuit"), CUIT, cuit_ok(f.get("comprador_cuit"), f.get("_texto")),
            "No figura el CUIT 30-67773730-8 del comprador.")


def reglas_packing(r, p, fac, hay):
    d = lbl(DOCS["packing"], p.get("numero_packing") or p.get("po"))
    es_consignatario_ok(r, d, p.get("cliente_razon_social"), p.get("cliente_cuit"), texto=p.get("_texto"), usar_texto=True)
    if fac is None:
        sin_factura(r, d, ["PO vs Factura", "Descripción vs Factura"], hay, p.get("po"))
        return
    r.check(d, "PO vs Factura " + show(fac.get("numero_factura")), p.get("po"), fac.get("po"),
            bool(p.get("po")) and id_match(p.get("po"), fac.get("po")), "No figura PO." if not p.get("po") else "La PO no coincide con la factura.")
    ok, mot = desc_check(p.get("descripcion_producto"), fac.get("descripcion_producto"))
    r.check(d, "Descripción vs Factura", p.get("descripcion_producto"), fac.get("descripcion_producto"), ok, mot)


def reglas_coa(r, c, fac, hay):
    d = lbl(DOCS["coa"], c.get("po"))
    tiene_po, tiene_desc = bool(c.get("po")), bool(c.get("descripcion_producto"))
    if not (tiene_po or tiene_desc):
        r.err(d, "PO o descripción vs Factura", "", "", "El COA no contiene ni PO ni descripción del producto.")
    elif fac is None:
        sin_factura(r, d, ["PO o descripción vs Factura"], hay, c.get("po"))
    else:
        ok_po = tiene_po and id_match(c.get("po"), fac.get("po"))
        ok_desc, ratio = desc_match(c.get("descripcion_producto"), fac.get("descripcion_producto"), 0.6) if tiene_desc else (False, 0)
        r.check(d, "PO o descripción vs Factura " + show(fac.get("numero_factura")),
                f"PO: {show(c.get('po'))} | Desc: {show(c.get('descripcion_producto'))}",
                f"PO: {show(fac.get('po'))} | Desc: {show(fac.get('descripcion_producto'))}",
                ok_po or ok_desc, "Ni la PO ni la descripción del COA coinciden con la factura.")
    venc = parse_date(c.get("fecha_vencimiento"))
    limite = HOY + timedelta(days=DIAS_MIN_VENCIMIENTO)
    if venc is None:
        r.err(d, "Fecha de vencimiento", c.get("fecha_vencimiento"), f"> {limite:%d/%m/%Y}", "No se pudo leer la fecha de vencimiento. Revisar.")
    else:
        r.check(d, "Fecha de vencimiento (o retest date)", f"{venc:%d/%m/%Y}", f"posterior a {limite:%d/%m/%Y} (hoy + {DIAS_MIN_VENCIMIENTO} días)",
                venc > limite, f"Vence el {venc:%d/%m/%Y}: dentro de los próximos {DIAS_MIN_VENCIMIENTO} días (o ya vencido).")


def reglas_certificado_naladisa(r, c, fac, hay):
    d = lbl(DOCS["certificado"] + " (NALADISA)", c.get("declaracion_factura_numero"))
    ctrls = ["Exportador: razón social", "Exportador: dirección", "Importador: razón social", "Importador: dirección",
             "Denominación vs Factura", "Peso o cantidad vs Factura", "Valor FOB vs Factura",
             "Declaración: número de factura", "Declaración: fecha de factura"]
    if fac is None:
        sin_factura(r, d, ctrls, hay, None)
        return
    r.check(d, ctrls[0], c.get("exportador_razon_social"), fac.get("vendedor_razon_social"),
            name_match(c.get("exportador_razon_social"), fac.get("vendedor_razon_social")),
            "La razón social del exportador no coincide con el vendedor de la factura.")
    ok, _ = desc_match(c.get("exportador_direccion"), fac.get("vendedor_direccion"), 0.6)
    r.check(d, ctrls[1], c.get("exportador_direccion"), fac.get("vendedor_direccion"), ok,
            "La dirección del exportador no coincide con la de la factura. Revisar.")
    r.check(d, ctrls[2], c.get("importador_razon_social"), fac.get("comprador_razon_social"),
            name_match(c.get("importador_razon_social"), fac.get("comprador_razon_social")),
            "La razón social del importador no coincide con el comprador de la factura.")
    ok, _ = desc_match(c.get("importador_direccion"), fac.get("comprador_direccion"), 0.6)
    r.check(d, ctrls[3], c.get("importador_direccion"), fac.get("comprador_direccion"), ok,
            "La dirección del importador no coincide con la de la factura. Revisar.")
    ok, mot = desc_check(c.get("denominacion_mercaderias"), fac.get("descripcion_producto"))
    r.check(d, ctrls[4], c.get("denominacion_mercaderias"), fac.get("descripcion_producto"), ok, mot)
    pc = pick_num(c.get("peso_o_cantidad"))
    cands = [x for k in ("cantidad", "peso_bruto_kg", "peso_neto_kg") for x in num_cands(fac.get(k))]
    if pc is None or not cands:
        r.err(d, ctrls[5], c.get("peso_o_cantidad"), "", "No se pudo comparar peso/cantidad (dato ilegible o ausente). Revisar.")
    else:
        r.check(d, ctrls[5], pc, " / ".join(str(x) for x in cands), eq_num(c.get("peso_o_cantidad"), cands[0]) or any(eq_num(c.get("peso_o_cantidad"), x) for x in cands),
                "El peso/cantidad no coincide con la factura.")
    fob, tot = pick_num(c.get("valor_fob")), pick_num(fac.get("valor_total"))
    if fob is None or tot is None:
        r.err(d, ctrls[6], c.get("valor_fob"), fac.get("valor_total"), "No se pudo comparar el valor FOB (dato ilegible o ausente). Revisar.")
    else:
        r.check(d, ctrls[6], fob, tot, num_equal(fob, tot, rel=0.0005, abs_tol=0.01),
                f"El valor FOB ({fob:,.2f}) no coincide con el total de la factura ({tot:,.2f}).")
    if not c.get("declaracion_factura_numero"):
        r.err(d, ctrls[7], "", fac.get("numero_factura"), "La declaración de origen no detalla número de factura.")
    else:
        r.check(d, ctrls[7], c.get("declaracion_factura_numero"), fac.get("numero_factura"),
                id_match(c.get("declaracion_factura_numero"), fac.get("numero_factura")),
                "El número de factura de la declaración no coincide con la factura.")
    fc, ff = parse_date(c.get("declaracion_factura_fecha")), parse_date(fac.get("fecha_factura"))
    if fc is None or ff is None:
        r.err(d, ctrls[8], c.get("declaracion_factura_fecha"), fac.get("fecha_factura"),
              "No se pudo leer la fecha de factura en la declaración o en la factura. Revisar.")
    else:
        r.check(d, ctrls[8], f"{fc:%d/%m/%Y}", f"{ff:%d/%m/%Y}", fc == ff, "La fecha de factura de la declaración no coincide con la factura.")


def ejecutar_controles(modo, ext):
    """ext: {'transporte': dict|None, 'factura': [..], 'packing': [..], 'coa': [..], 'certificado': [..]}"""
    r = Resultados()
    facs, packs = ext.get("factura") or [], ext.get("packing") or []
    hay = bool(facs)

    t = ext.get("transporte")
    if t is None:
        r.na(DOCS["transporte"], "Documento", "Documento no cargado.")
    elif modo in TRANSPORT_RULES:
        TRANSPORT_RULES[modo][1](r, t, facs, packs)
    else:
        r.na(DOCS["transporte"], "Controles de transporte", f"Lógica pendiente para modo {modo}.")

    if not facs:
        r.na(DOCS["factura"], "Documento", "Documento no cargado.")
    for f in facs:
        reglas_factura(r, f)

    for key, fn in (("packing", reglas_packing), ("coa", reglas_coa), ("certificado", reglas_certificado_naladisa)):
        docs = ext.get(key) or []
        if not docs:
            r.na(DOCS[key], "Documento", "Documento no cargado.")
            continue
        for x in docs:
            if key == "certificado":
                fac = buscar_factura(facs, numero=x.get("declaracion_factura_numero")) or (facs[0] if len(facs) == 1 else None)
            else:
                fac = buscar_factura(facs, po=x.get("po"))
                if fac is None and key == "coa" and x.get("descripcion_producto"):
                    fac = next((f for f in facs if desc_match(x["descripcion_producto"], f.get("descripcion_producto"), 0.8)[0]), None)
            fn(r, x, fac, hay)
        # facturas sin su documento
        for f in facs:
            if key == "certificado":
                cubierta = any(buscar_factura([f], numero=x.get("declaracion_factura_numero")) or len(facs) == 1 for x in docs)
            else:
                cubierta = any(id_match(x.get("po"), f.get("po")) for x in docs) or (
                    key == "coa" and any(x.get("descripcion_producto") and desc_match(x["descripcion_producto"], f.get("descripcion_producto"), 0.8)[0] for x in docs))
            if not cubierta:
                r.na(lbl(DOCS[key], f"de factura {f.get('numero_factura')}"), "Documento", f"No cargado: falta {DOCS[key]} de la factura {f.get('numero_factura')} (PO {f.get('po')}).")
    return r.filas


# ----------------------------------------------------------------------------
# Lectura de PDF e IA
# ----------------------------------------------------------------------------
def pdf_a_texto(archivo):
    import pdfplumber
    partes = []
    with pdfplumber.open(io.BytesIO(archivo.getvalue())) as pdf:
        for pg in pdf.pages:
            partes.append(pg.extract_text() or "")
    txt = "\n".join(partes).strip()
    if len(txt) > 15000:
        txt = txt[:9000] + "\n[...]\n" + txt[-6000:]
    return txt


def secreto(nombre):
    try:
        v = st.secrets.get(nombre)
    except Exception:
        v = None
    return v or os.environ.get(nombre)


def construir_prompt(schema):
    campos = "\n".join(f'- "{k}": {v}' for k, v in schema["campos"].items())
    return (
        f"Sos un asistente de comercio exterior. Extraé datos de un documento: {schema['desc']}.\n"
        f"Devolvé SOLO un objeto JSON con exactamente estas claves (null si el dato no figura; NO inventes):\n{campos}\n"
        f"{NUM_NOTE}"
    )


def parse_json(txt):
    try:
        return json.loads(txt)
    except Exception:
        m = re.search(r"\{.*\}", txt, re.S)
        if m:
            return json.loads(m.group(0))
        raise


GROQ_PREFERIDOS = ["openai/gpt-oss-120b", "llama-3.3-70b-versatile", "meta-llama/llama-4-scout-17b-16e-instruct",
                   "openai/gpt-oss-20b", "qwen/qwen3-32b"]
_MODELO_GROQ = {}


def modelos_groq(cliente):
    """Modelos a probar, en orden: GROQ_MODEL (secret) o los preferidos que la cuenta tenga disponibles."""
    if secreto("GROQ_MODEL"):
        return [secreto("GROQ_MODEL")]
    if "lista" not in _MODELO_GROQ:
        ids = [m.id for m in cliente.models.list().data]
        lista = [m for m in GROQ_PREFERIDOS if m in ids]
        extra = [i for i in ids if i not in lista and not any(x in i.lower() for x in ("whisper", "guard", "tts", "playai", "orpheus", "compound"))]
        lista += extra[:2]
        if not lista:
            raise RuntimeError(f"La cuenta de Groq no tiene modelos de texto disponibles: {ids}")
        _MODELO_GROQ["lista"] = lista
    return _MODELO_GROQ["lista"]


def es_temporal(e):
    t = str(e).lower()
    return any(x in t for x in ("429", "503", "502", "500", "over capacity", "rate limit", "overloaded", "timeout"))


def llamar_ia(system, texto):
    error_groq = None
    if secreto("GROQ_API_KEY"):
        import time
        from groq import Groq
        try:
            cliente = Groq(api_key=secreto("GROQ_API_KEY"))
            for modelo in modelos_groq(cliente):
                for intento in range(3):
                    try:
                        resp = cliente.chat.completions.create(
                            model=modelo, temperature=0, response_format={"type": "json_object"},
                            messages=[{"role": "system", "content": system},
                                      {"role": "user", "content": f"TEXTO DEL DOCUMENTO:\n{texto}"}],
                        )
                        return parse_json(resp.choices[0].message.content), f"Groq ({modelo})"
                    except Exception as e:
                        error_groq = e
                        if es_temporal(e) and intento < 2:
                            time.sleep(5 * (intento + 1))
                            continue
                        break  # error no temporal o reintentos agotados: probar el siguiente modelo
        except Exception as e:
            error_groq = e
    if secreto("ANTHROPIC_API_KEY"):
        import anthropic
        msg = anthropic.Anthropic(api_key=secreto("ANTHROPIC_API_KEY")).messages.create(
            model="claude-sonnet-5-5", max_tokens=2000, temperature=0, system=system,
            messages=[{"role": "user", "content": f"TEXTO DEL DOCUMENTO:\n{texto}\n\nRespondé solo con el JSON."}],
        )
        return parse_json(msg.content[0].text), "Anthropic"
    raise RuntimeError(f"Groq falló: {error_groq}" if error_groq else "Falta GROQ_API_KEY (o ANTHROPIC_API_KEY) en secrets.")


# ----------------------------------------------------------------------------
# Excel
# ----------------------------------------------------------------------------
def generar_excel(filas, extraidos, referencia, modo):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    wb = Workbook()
    ws = wb.active
    ws.title = "Resultado"
    ws["A1"] = f"Corrector de documentos Coca-Cola — Referencia: {referencia} — Modo: {modo} — Fecha: {HOY:%d/%m/%Y}"
    ws["A1"].font = Font(bold=True, size=12)
    enc = ["Documento", "Control", "En el documento", "Esperado / Referencia", "Resultado", "Motivo"]
    ws.append([])
    ws.append(enc)
    for c in ws[3]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="C00000")
    colores = {OK: "C6EFCE", ERR: "FFC7CE", NA: "EDEDED"}
    for f in filas:
        ws.append([f[k] for k in enc])
        fila = ws.max_row
        ws.cell(fila, 5).fill = PatternFill("solid", fgColor=colores[f["Resultado"]])
        ws.cell(fila, 5).alignment = Alignment(horizontal="center")
        ws.cell(fila, 5).font = Font(bold=True, size=13)
    for col, ancho in zip("ABCDEF", (24, 38, 45, 45, 11, 60)):
        ws.column_dimensions[col].width = ancho
    for row in ws.iter_rows(min_row=4):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top", horizontal=c.alignment.horizontal)
    ws.freeze_panes = "A4"

    ws2 = wb.create_sheet("Datos extraídos")
    ws2.append(["Documento", "Campo", "Valor"])
    for c in ws2[1]:
        c.font = Font(bold=True)
    for key, data in extraidos.items():
        lista = data if isinstance(data, list) else [data]
        for n, dato in enumerate(lista, 1):
            for k, v in (dato or {}).items():
                if k.startswith("_"):
                    continue
                ws2.append([f"{DOCS[key]} #{n}", k, json.dumps(v, ensure_ascii=False) if isinstance(v, (list, dict)) else show(v)])
    ws2.column_dimensions["A"].width, ws2.column_dimensions["B"].width, ws2.column_dimensions["C"].width = 26, 32, 90
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ----------------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------------
def main():
    st.set_page_config(page_title="Corrector Docs Coca", page_icon="📑", layout="wide")
    st.title("📑 Corrector de documentos Coca-Cola")
    st.caption("Subí los documentos de la operación (no hace falta que estén los cinco). La IA lee los PDF y las reglas se aplican en Python.")

    c1, c2 = st.columns(2)
    referencia = c1.text_input("Referencia de la operación *", placeholder="Ej: IM-2026-123")
    modo = c2.selectbox("Modo de transporte *", ["Marítimo", "Aéreo", "Terrestre"])
    if modo not in TRANSPORT_RULES:
        st.warning(f"Lógica de transporte para **{modo}** pendiente: se controlan los demás documentos y el de transporte queda como 'Lógica pendiente'.")
    st.info("Certificado de origen: hoy solo **NALADISA**. (Mercosur: pendiente.)")

    archivos = {}
    cols = st.columns(5)
    for col, (key, nombre) in zip(cols, DOCS.items()):
        multiple = key != "transporte"
        archivos[key] = col.file_uploader(nombre + (" (varios)" if multiple else ""), type=["pdf"], key=key,
                                          accept_multiple_files=multiple)

    if st.button("🔍 Analizar", type="primary"):
        if not referencia.strip():
            st.error("Ingresá la referencia.")
            return
        if not any(archivos.values()):
            st.error("Subí al menos un documento.")
            return
        extraidos, avisos, fallidos = {}, [], []
        with st.spinner("Leyendo documentos..."):
            for key, subidos in archivos.items():
                lista = subidos if isinstance(subidos, list) else ([subidos] if subidos else [])
                datos_doc = []
                for arch in lista:
                    skey = key
                    if key == "transporte":
                        if modo not in TRANSPORT_RULES:
                            datos_doc.append({})
                            continue
                        skey = TRANSPORT_RULES[modo][0]
                    try:
                        texto = pdf_a_texto(arch)
                        if len(texto) < 30:
                            raise ValueError("El PDF no tiene texto legible (¿escaneado?).")
                        datos, motor = llamar_ia(construir_prompt(SCHEMAS[skey]), texto)
                        datos["_texto"] = texto
                        datos_doc.append(datos)
                        avisos.append(f"{DOCS[key]} ({arch.name}): leído con {motor}")
                    except Exception as e:
                        st.error(f"No se pudo leer {arch.name}: {e}")
                        fallidos.append({"Documento": f"{DOCS[key]}: {arch.name}", "Control": "Lectura del archivo", "En el documento": "",
                                         "Esperado / Referencia": "", "Resultado": ERR,
                                         "Motivo": f"NO SE PUDO LEER, volver a analizar. Los controles que dependen de este archivo pueden aparecer mal. ({str(e)[:120]})"})
                extraidos[key] = (datos_doc[0] if datos_doc else None) if key == "transporte" else datos_doc
        filas = fallidos + ejecutar_controles(modo, extraidos)
        st.session_state["res"] = (filas, extraidos, referencia.strip(), modo, avisos)

    if "res" in st.session_state:
        filas, extraidos, ref, modo_r, avisos = st.session_state["res"]
        df = pd.DataFrame(filas)
        n_ok, n_err = (df["Resultado"] == OK).sum(), (df["Resultado"] == ERR).sum()
        m1, m2, m3 = st.columns(3)
        m1.metric("✔ Correctos", int(n_ok))
        m2.metric("✘ A corregir / revisar", int(n_err))
        m3.metric("— No verificables", int((df["Resultado"] == NA).sum()))
        st.dataframe(
            df.style.map(lambda v: {OK: "background-color:#C6EFCE", ERR: "background-color:#FFC7CE", NA: "background-color:#EDEDED"}.get(v, ""),
                         subset=["Resultado"]),
            use_container_width=True, hide_index=True,
        )
        nombre = re.sub(r"[^\w\-]+", "_", ref)
        st.download_button("⬇️ Descargar Excel", generar_excel(filas, extraidos, ref, modo_r),
                           file_name=f"Correccion_{nombre}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with st.expander("Detalle de lectura"):
            st.write(avisos)
            st.json({DOCS[k]: [{a: b for a, b in d.items() if not a.startswith("_")} for d in (v if isinstance(v, list) else [v or {}])]
                     for k, v in extraidos.items()})


if not os.environ.get("CORRECTOR_TEST"):
    main()
