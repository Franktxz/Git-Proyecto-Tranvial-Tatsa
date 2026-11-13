"""
Pendientes de compra (OC por entregar y RQ sin OC) · TATSA y TVL
================================================================

Genera `Herramienta_Pendientes.html`: una página interactiva (sin internet) para hacer seguimiento de:
  1. OC POR ENTREGAR: OC_PENDIENTE_ENTREGA del Excel (repuestos e insumos) + OC de servicio (SERV-) pendientes de
     facturar leídas de Odoo. Vista por proveedor (con contacto y texto listo para enviarle sus pendientes), por comprador
     o en lista, con atraso frente a la fecha prevista.
  2. RQ SIN OC: RQ_SIN_OC del Excel, con días desde la FECHA DE APROBACIÓN y la cotización en curso (comprador y proveedor).
Cada línea se clasifica como Repuesto, Insumo / herramienta o Servicio.

Alcance (periodo 2026): TATSA [TA01432] Taller Operativo; TVL [TR01432] Taller Operativo y [TR01421] Taller Administrativo.

Cómo actualizar:
  1. Ejecutar `Saldo_Disponible_Real.ipynb` (genera el Excel con OC_PENDIENTE_ENTREGA y RQ_SIN_OC).
  2. `python Herramienta_Pendientes.py` (lee el Excel y consulta Odoo en modo solo lectura con el usuario del .env).
  3. Abrir `Herramienta_Pendientes.html`.
"""

import json
import os
import re
from datetime import datetime

import numpy as np
import pandas as pd

ARCHIVO_SALDO = "Saldo_Disponible_Real_2024_2026.xlsx"
PLANTILLA = "Herramienta_Pendientes_plantilla.html"
ARCHIVO_SALIDA = "Herramienta_Pendientes.html"
USUARIO_ODOO = "JIM"
PERIODO = 2026
ALCANCE = {"TATSA": ["TA01432"], "TVL": ["TR01432", "TR01421"]}   # empresa -> centros de costo
EMPRESA_ID = {"TVL": 1, "TATSA": 2}
HOY = pd.Timestamp.today().normalize()


def texto(v):
    return None if v is None or v is False or (isinstance(v, float) and np.isnan(v)) or str(v).strip() in ("", "nan", "None") else str(v).strip()


def num(v, d=2):
    try:
        v = float(v)
    except (TypeError, ValueError):
        return None
    return None if np.isnan(v) or np.isinf(v) else round(v, d)


def fecha(v):
    if v is None or v is False or (not isinstance(v, str) and pd.isna(v)):
        return None
    return pd.Timestamp(v).strftime("%Y-%m-%d")


def en_alcance(df):
    cc = df["Cuenta Analítica"].astype(str).str.extract(r"\[(\w+)\]", expand=False)
    ok = pd.Series(False, index=df.index)
    for emp, centros in ALCANCE.items():
        ok |= (df["Empresa"] == emp) & cc.isin(centros)
    return df[ok & (df["Periodo"] == PERIODO)].assign(CC=cc)


def tipo_de(partida, servicio=False):
    if servicio:
        return "Servicio"
    return "Repuesto" if "REPUESTOS" in str(partida).upper() else "Insumo / herramienta"


# ----------------------------------------------------------------------------------------------------------------------
# 1. EXCEL (Saldo_Disponible_Real)
# ----------------------------------------------------------------------------------------------------------------------
oc = en_alcance(pd.read_excel(ARCHIVO_SALDO, sheet_name="OC_PENDIENTE_ENTREGA", engine="calamine"))
rq = en_alcance(pd.read_excel(ARCHIVO_SALDO, sheet_name="RQ_SIN_OC", engine="calamine"))
# Llave (grupo de SKUs equivalentes) de Dim_Producto, para buscar por llave
_dim = pd.read_parquet("Información_Proyecto_BI/Dim_Producto.parquet")
LLAVE = _dim.dropna(subset=["Cod_Grupo_Sku"]).drop_duplicates("Cod_Producto").set_index("Cod_Producto")["Cod_Grupo_Sku"].to_dict()
print(f"Excel: {len(oc)} líneas de OC por entregar | {len(rq)} líneas de RQ sin OC")

# ----------------------------------------------------------------------------------------------------------------------
# 2. ODOO (solo lectura): contacto de proveedores, OC de servicio pendientes, aprobación y cotizaciones de los RQ
# ----------------------------------------------------------------------------------------------------------------------
import xmlrpc.client
from dotenv import load_dotenv

load_dotenv()
url, db = os.getenv("URL_ODOO"), os.getenv("BD_ODOO")
usr, pwd = os.getenv(f"USER_{USUARIO_ODOO}"), os.getenv(f"PASSWORD_{USUARIO_ODOO}")
uid = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common").authenticate(db, usr, pwd, {})
mdl = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object", allow_none=True)


def odoo(modelo, metodo, args, kw=None, empresa=None):
    kw = dict(kw or {})
    kw["context"] = {"lang": "es_PE", "allowed_company_ids": [empresa] if empresa else [1, 2], "active_test": False}
    return mdl.execute_kw(db, uid, pwd, modelo, metodo, args, kw)


def leer_todo(modelo, dominio, campos, empresa, lote=2000):
    filas, off = [], 0
    while True:
        b = odoo(modelo, "search_read", [dominio], {"fields": campos, "limit": lote, "offset": off, "order": "id"}, empresa)
        filas += b
        if len(b) < lote:
            return filas
        off += lote


m2o = lambda v: v[1] if isinstance(v, list) and len(v) == 2 else None
m2o_id = lambda v: v[0] if isinstance(v, list) and v else None

# 2.1 OC de servicio (SERV-) abiertas y pendientes de facturar en los centros del alcance
serv = []
for emp, centros in ALCANCE.items():
    e = EMPRESA_ID[emp]
    lineas = leer_todo("purchase.order.line", [
        ["company_id", "=", e], ["order_id.name", "=like", "SERV-%"],
        ["order_id.state_new", "in", ["purchase", "done", "parcialmente_cerrado"]],
        ["account_analytic_id.code", "in", centros]],
        ["order_id", "product_id", "name", "product_qty", "qty_invoiced", "price_total", "account_analytic_id", "date_planned"], e)
    if not lineas:
        continue
    ords = {o["id"]: o for o in odoo("purchase.order", "read", [sorted({m2o_id(l["order_id"]) for l in lineas})],
                                     {"fields": ["name", "partner_id", "user_id", "date_approve", "currency_rate", "requisition_id", "state_new"]}, e)}
    rq_ids = sorted({m2o_id(o["requisition_id"]) for o in ords.values() if o["requisition_id"]})
    rq_info = {r["id"]: r for r in odoo("purchase.requisition", "read", [rq_ids], {"fields": ["name", "periodo_id"]}, e)} if rq_ids else {}
    for l in lineas:
        o = ords[m2o_id(l["order_id"])]
        r = rq_info.get(m2o_id(o["requisition_id"]))
        periodo = int(m2o(r["periodo_id"])) if r and r["periodo_id"] and str(m2o(r["periodo_id"])).isdigit() else int((o["date_approve"] or "2000")[:4])
        pend = (l["product_qty"] or 0) - (l["qty_invoiced"] or 0)
        if periodo != PERIODO or pend <= 1e-6 or not l["product_qty"]:
            continue
        valor = (l["price_total"] or 0) / (o["currency_rate"] or 1.0) * pend / l["product_qty"]
        cod = re.match(r"\[(\w+)\]", m2o(l["product_id"]) or "")
        serv.append({"Empresa": emp, "Cod_RQ": r["name"] if r else None, "Cod_OC": o["name"], "Fecha_OC": (o["date_approve"] or "")[:10] or None,
                     "Comprador": m2o(o["user_id"]), "Proveedor": m2o(o["partner_id"]), "partner_id": m2o_id(o["partner_id"]),
                     "Cod_Producto": cod.group(1) if cod else None, "Producto": re.sub(r"^\[\w+\]\s*", "", m2o(l["product_id"]) or l["name"] or ""),
                     "CC": (m2o(l["account_analytic_id"]) or "")[1:8], "Cant_OC": l["product_qty"], "Cant_Recibida": l["qty_invoiced"],
                     "Cant_Pendiente": pend, "Valor_Pendiente": valor, "Fecha_Entrega_Prevista": (l["date_planned"] or "")[:10] or None,
                     "Estado_OC": o["state_new"], "Servicio": True})
print(f"Odoo: {len(serv)} líneas de servicio pendientes de facturar")

# 2.2 Proveedor y comprador de cada OC del Excel (para el contacto y la referencia del proveedor)
oc_ids = {}
for emp in ALCANCE:
    nombres = sorted(oc.loc[oc["Empresa"] == emp, "Cod_OC"].unique())
    for o in (odoo("purchase.order", "search_read", [[["name", "in", nombres]]], {"fields": ["name", "partner_id", "partner_ref"]}, EMPRESA_ID[emp]) if nombres else []):
        oc_ids[(emp, o["name"])] = (m2o_id(o["partner_id"]), texto(o["partner_ref"]))
oc["partner_id"] = [oc_ids.get((e, c), (None, None))[0] for e, c in zip(oc["Empresa"], oc["Cod_OC"])]
oc["Ref_Proveedor"] = [oc_ids.get((e, c), (None, None))[1] for e, c in zip(oc["Empresa"], oc["Cod_OC"])]

socios = sorted({int(p) for p in list(oc["partner_id"].dropna()) + [s["partner_id"] for s in serv if s["partner_id"]]})
CONTACTO = {}
if socios:
    for p in odoo("res.partner", "read", [socios], {"fields": ["name", "vat", "email", "phone", "mobile", "child_ids"]}):
        CONTACTO[p["id"]] = {"ruc": texto(p["vat"]), "mail": texto(p["email"]), "tel": texto(p["phone"]) or texto(p["mobile"])}

# 2.3 RQ: fecha de aprobación (si el Excel aún no la trae) y cotizaciones en curso
rq_head = {}
for emp in ALCANCE:
    nombres = sorted(rq.loc[rq["Empresa"] == emp, "Cod_RQ"].unique())
    if not nombres:
        continue
    for r in odoo("purchase.requisition", "search_read", [[["name", "in", nombres]]],
                  {"fields": ["name", "create_date", "fecha_aprobacion", "purchase_date_apro", "prior"]}, EMPRESA_ID[emp]):
        rq_head[(emp, r["name"])] = r
    cots = odoo("purchase.order", "search_read", [[["requisition_id.name", "in", nombres], ["state", "in", ["draft", "sent", "to approve"]]]],
                {"fields": ["name", "requisition_id", "partner_id", "user_id", "state", "date_order", "amount_total", "currency_id",
                            "aprobado_char", "valida", "fue_desaprobado", "user_tecnica", "fecha_tecnica", "message_attachment_count"]}, EMPRESA_ID[emp])
    # Chat de cada cotización: el comprador publica "ADJUNTO COTIZACIÓN … PARA VALIDAR" con el archivo; sin adjunto no se puede validar
    chat = {}
    if cots:
        for mm in odoo("mail.message", "search_read", [[["model", "=", "purchase.order"], ["res_id", "in", [c["id"] for c in cots]], ["message_type", "in", ["comment", "email"]]]],
                       {"fields": ["res_id", "date", "author_id", "body", "attachment_ids"], "order": "date asc"}, EMPRESA_ID[emp]):
            d = chat.setdefault(mm["res_id"], {"adj": 0, "ult": None})
            d["adj"] += len(mm["attachment_ids"])
            cuerpo = re.sub(r"\s+", " ", re.sub("<[^>]+>", " ", mm["body"] or "")).strip()
            if cuerpo or mm["attachment_ids"]:
                d["ult"] = {"f": mm["date"][:16], "a": m2o(mm["author_id"]), "t": cuerpo[:160], "n": len(mm["attachment_ids"])}
    for c in cots:
        c["chat"] = chat.get(c["id"], {"adj": 0, "ult": None})
        rq_head.setdefault((emp, m2o(c["requisition_id"])), {}).setdefault("cots", []).append(c)


def estado_cot(c):
    """Dónde está la cotización en el flujo: proveedor → cotización adjunta en el chat → Aprobar (selección técnica) → confirmar OC."""
    if c.get("fue_desaprobado") or c.get("aprobado_char") == "Desaprobado" or c.get("valida") == "no_val":
        return "Desaprobada"
    if c.get("aprobado_char") == "Aprobado" or c.get("user_tecnica"):
        return "Aprobada · falta confirmar OC"
    adjunto = (c.get("message_attachment_count") or 0) > 0 or c["chat"]["adj"] > 0
    if adjunto:
        return "Por aprobar"
    if m2o(c["partner_id"]) in RELLENO:
        return "Sin proveedor"
    return "Sin cotización adjunta"


RELLENO = (None, "OdooBot", "ODOO INC")  # proveedores de relleno en cotizaciones recién creadas

# 2.4 Tipo de producto en Odoo: un RQ "compra" puede pedir un servicio (p. ej. mantenimiento de una herramienta)
codigos = sorted(set(oc["Cod_Producto"].dropna()) | set(rq["Cod_Producto"].dropna()))
ES_SERVICIO = {p["default_code"]: p["detailed_type"] == "service"
               for p in odoo("product.product", "search_read", [[["default_code", "in", codigos]]], {"fields": ["default_code", "detailed_type"]})} if codigos else {}

if "Fecha_Aprobacion" not in rq.columns:
    rq["Fecha_Aprobacion"] = [((rq_head.get((e, c)) or {}).get("fecha_aprobacion") or ((rq_head.get((e, c)) or {}).get("purchase_date_apro") or "")[:10] or f)
                              for e, c, f in zip(rq["Empresa"], rq["Cod_RQ"], rq["Fecha_RQ"].astype(str).str[:10])]
rq["Fecha_Aprobacion"] = rq["Fecha_Aprobacion"].fillna(rq["Fecha_RQ"])

# ----------------------------------------------------------------------------------------------------------------------
# 3. FILAS PARA LA PÁGINA
# ----------------------------------------------------------------------------------------------------------------------
def dias(desde):
    return None if desde is None else int((HOY - pd.Timestamp(desde)).days)


filas_oc = []
for r in oc.itertuples():
    prev = fecha(r.Fecha_Entrega_Prevista)
    filas_oc.append({
        "e": r.Empresa, "cc": r.CC, "t": tipo_de(r.Partida, ES_SERVICIO.get(r.Cod_Producto, False)), "rq": texto(r.Cod_RQ), "oc": r.Cod_OC, "fo": fecha(r.Fecha_OC),
        "cm": texto(r.Comprador), "pv": texto(r.Proveedor), "pid": int(r.partner_id) if pd.notna(r.partner_id) else None,
        "ref": texto(r.Ref_Proveedor), "c": texto(r.Cod_Producto), "k": LLAVE.get(r.Cod_Producto), "n": texto(r.Producto), "q": num(r.Cant_OC), "qr": num(r.Cant_Recibida),
        "qp": num(r.Cant_Pendiente), "v": num(r.Valor_Pendiente), "fp": prev, "at": max(0, dias(prev)) if prev else None,
        "dv": -dias(prev) if prev and dias(prev) < 0 else None, "en": texto(r.Estado_Entrega), "st": texto(r.Estado_OC),
    })
ESTADO_OC = {"purchase": "Orden de Compra", "done": "Bloqueado", "parcialmente_cerrado": "Cerrado Parcialmente"}
for s in serv:
    prev = s["Fecha_Entrega_Prevista"]
    filas_oc.append({
        "e": s["Empresa"], "cc": s["CC"], "t": "Servicio", "rq": s["Cod_RQ"], "oc": s["Cod_OC"], "fo": s["Fecha_OC"], "cm": s["Comprador"],
        "pv": s["Proveedor"], "pid": s["partner_id"], "ref": None, "c": s["Cod_Producto"], "k": LLAVE.get(s["Cod_Producto"]), "n": s["Producto"], "q": num(s["Cant_OC"]),
        "qr": num(s["Cant_Recibida"]), "qp": num(s["Cant_Pendiente"]), "v": num(s["Valor_Pendiente"]), "fp": prev,
        "at": max(0, dias(prev)) if prev else None, "dv": -dias(prev) if prev and dias(prev) < 0 else None,
        "en": "Pendiente de facturar" if not s["Cant_Recibida"] else "Facturado en parte", "st": ESTADO_OC.get(s["Estado_OC"], s["Estado_OC"]),
    })

filas_rq = []
for r in rq.itertuples():
    h = rq_head.get((r.Empresa, r.Cod_RQ)) or {}
    cots = [{"oc": c["name"], "pv": None if m2o(c["partner_id"]) in RELLENO else m2o(c["partner_id"]), "cm": m2o(c["user_id"]),
             "f": (c["date_order"] or "")[:10], "es": estado_cot(c), "adj": int(max(c.get("message_attachment_count") or 0, c["chat"]["adj"])),
             "ft": (c.get("fecha_tecnica") or "")[:10] or None, "ut": m2o(c.get("user_tecnica")), "ult": c["chat"]["ult"]} for c in h.get("cots", [])]
    fa = fecha(r.Fecha_Aprobacion)
    filas_rq.append({
        "e": r.Empresa, "cc": r.CC, "t": tipo_de(r.Partida, str(r.Tipo_RQ).lower() == "servicio" or ES_SERVICIO.get(r.Cod_Producto, False)), "rq": r.Cod_RQ,
        "fc": fecha(r.Fecha_RQ), "fa": fa, "d": dias(fa), "so": texto(r.Solicitante), "pr": texto(h.get("prior")),
        "c": texto(r.Cod_Producto), "k": LLAVE.get(r.Cod_Producto), "n": texto(r.Producto), "q": num(r.Cant_RQ), "qo": num(r.Cant_Pedida_OC), "qp": num(r.Cant_Pendiente),
        "v": num(r.Valor_Pendiente), "obs": texto(r.Observaciones), "cot": cots,
    })

meta = {
    "generado": datetime.now().strftime("%d/%m/%Y %H:%M"), "hoy": HOY.strftime("%Y-%m-%d"), "periodo": PERIODO,
    "excel": datetime.fromtimestamp(os.path.getmtime(ARCHIVO_SALDO)).strftime("%d/%m/%Y %H:%M"),
    "alcance": " · ".join(f"{e} {', '.join(c)}" for e, c in ALCANCE.items()),
}
contactos = {str(k): v for k, v in CONTACTO.items()}
plantilla = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), PLANTILLA), encoding="utf-8").read()
datos = json.dumps({"meta": meta, "oc": filas_oc, "rq": filas_rq, "ct": contactos}, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
with open(ARCHIVO_SALIDA, "w", encoding="utf-8") as fh:
    fh.write(plantilla.replace("/*__DATOS__*/null", datos))
print(f"Listo: {ARCHIVO_SALIDA} | {len(filas_oc)} líneas de OC ({len(serv)} de servicio) | {len(filas_rq)} líneas de RQ")
