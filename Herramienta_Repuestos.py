"""
Herramienta de repuestos por llave (reemplaza Analisis_Repuestos.xlsx)
=======================================================================

Genera `Herramienta_Repuestos.html`, una página interactiva que funciona sin internet, con tres vistas:

1. ELEGIR REPUESTOS: una fila por LLAVE (repuesto único = grupo de SKUs equivalentes de Dim_Producto) con su consumo a
   buses, stock, tránsito, compras y estado (COMPRAR / EVALUAR / APURAR TRÁNSITO / SOBRESTOCK / SIN ROTACIÓN…) con la
   explicación del cálculo. Se marcan las llaves que van al RQ.
2. ARMAR RQ: las llaves marcadas, agrupadas como se arman los requerimientos (proveedor, sistema, reparación o familia),
   con código a pedir, cantidad, precio y nota editables, y el saldo/déficit de presupuesto de cada línea al instante.
3. FONDOS: lo que pide el RQ frente al saldo de [TA01432] REPUESTOS / REPUESTOS (TATSA); se eligen los SKU de donde
   sacar el dinero (o se sugieren) y sale la redistribución origen → destino para cargar en Odoo.
El RQ en armado y los saldos elegidos quedan guardados en el navegador de quien usa la página.

Alcance (definido con el usuario):
  - Operación de buses = almacenes ATATS/Mantenimiento, PREST/Stock y TRANS/Stock; centros TA01432 (y TR01432/TR01421 de TVL
    solo para el tránsito físico). Los almacenes virtuales (Chasis, Conversión, Virtual, Buses) y los centros TA03432, TA04432,
    TA05432, TA06432 son PROYECTOS: se muestran aparte y no cuentan por defecto.
  - Presupuesto = solo [TA01432] situación REPUESTOS / REPUESTOS. El 1432 de TVL es mantenimiento de taller y herramientas.

Cómo actualizar:
  1. Ejecutar `Tablas_Actual_BI.ipynb` (kardex, Dim_Producto) y `Saldo_Disponible_Real.ipynb` (tránsito y presupuesto).
  2. `python Herramienta_Repuestos.py` (unos 30 s).
  3. Abrir `Herramienta_Repuestos.html` en el navegador.
Odoo solo se lee (usuario del .env) para el lead time real de cada compra (RQ → recepción, sin RQ de regularización).
"""

import json
import os
from datetime import datetime

import numpy as np
import pandas as pd

# ----------------------------------------------------------------------------------------------------------------------
# PARÁMETROS
# ----------------------------------------------------------------------------------------------------------------------
CARPETA_BI = "Información_Proyecto_BI"
ARCHIVO_KARDEX = f"{CARPETA_BI}/Fact_Kardex.parquet"
ARCHIVO_PRODUCTO = f"{CARPETA_BI}/Dim_Producto.parquet"
ARCHIVO_FLOTA = f"{CARPETA_BI}/Dim_Flota.parquet"
ARCHIVO_SALDO = "Saldo_Disponible_Real_2024_2026.xlsx"
PLANTILLA = "Herramienta_Repuestos_plantilla.html"
ARCHIVO_SALIDA = "Herramienta_Repuestos.html"

USAR_ODOO = True
USUARIO_ODOO = "JIM"
ANIO_PPTO = 2026
CENTRO_PPTO = "[TA01432]"                          # Presupuesto de repuestos de buses
SITUACION_PPTO = "REPUESTOS / REPUESTOS"
CENTROS_OPERACION = {"TA01432", "TR01432", "TR01421"}  # Tránsito que llega a la operación; el resto de centros = proyectos
MESES_SERIE = 24
PRODUCTOS_EXCLUIDOS = {"G3370009"}
FAMILIAS_EXCLUIDAS = {"BUSES"}

ALMACENES = {  # almacén del kardex -> grupo de stock en la página
    "ATATS/Mantenimiento": "TATSA",
    "PREST/Stock": "Préstamos",
    "TRANS/Stock": "TVL",
    "AVCH/Stock": "Proyectos",
    "AVCON/Stock": "Proyectos",
    "AVIRT/Stock": "Proyectos",
    "ABUSE/Stock": "Proyectos",
}
OPERACION = {a for a, g in ALMACENES.items() if g != "Proyectos"}
GRUPOS_STOCK = ["TATSA", "Préstamos", "TVL", "Proyectos"]


def texto(v):
    return None if v is None or (isinstance(v, float) and np.isnan(v)) or str(v).strip() in ("", "nan", "None") else str(v).strip()


def num(v, d=2):
    if v is None or (isinstance(v, float) and (np.isnan(v) or np.isinf(v))):
        return None
    return round(float(v), d)


def fecha(v):
    return None if v is None or pd.isna(v) else pd.Timestamp(v).strftime("%Y-%m-%d")


# ----------------------------------------------------------------------------------------------------------------------
# 1. PRODUCTOS Y LLAVES
# ----------------------------------------------------------------------------------------------------------------------
dim = pd.read_parquet(ARCHIVO_PRODUCTO)
dim = dim[~dim["Cod_Producto"].isin(PRODUCTOS_EXCLUIDOS)].copy()
for c in ["Familia", "Familia_2", "Familia_3", "Uso"]:
    dim[c] = dim[c].astype("string").str.strip()
dim["Llave"] = dim["Cod_Grupo_Sku"].fillna(dim["Cod_Producto"])
dim["Producto_Grupo"] = dim["Producto_Grupo"].fillna(dim["Producto"])
LLAVE_DE = dim.drop_duplicates("Cod_Producto").set_index("Cod_Producto")["Llave"].to_dict()


def moda(s):
    s = s.dropna()
    s = s[~s.astype(str).isin(["", "S/M", "S/NP", "Otros", "OTROS"])]
    return s.mode().iat[0] if len(s) else None


info = dim.groupby("Llave").agg(
    Producto=("Producto_Grupo", "first"), Familia=("Familia", moda), Familia_2=("Familia_2", moda), Familia_3=("Familia_3", moda),
    Uso=("Uso", moda),
    UdM=("Unidad_Medida", moda))
CODIGOS = {ll: [{"c": r.Cod_Producto, "n": texto(r.Producto), "p": texto(r.Nro_Parte), "m": texto(r.Marca)}
                for r in g.sort_values("Cod_Producto").itertuples()] for ll, g in dim.groupby("Llave")}

# ----------------------------------------------------------------------------------------------------------------------
# 2. KARDEX: CONSUMO A BUSES, STOCK Y COMPRAS
# ----------------------------------------------------------------------------------------------------------------------
k = pd.read_parquet(ARCHIVO_KARDEX)
k = k[~k["Cod_Producto"].isin(PRODUCTOS_EXCLUIDOS)].copy()
k["Llave"] = k["Cod_Producto"].map(LLAVE_DE).fillna(k["Cod_Producto"])
k["Mes"] = k["Fecha"].dt.to_period("M")
CORTE = k["Fecha"].max().normalize()
ULTIMO_CERRADO = CORTE.to_period("M") - 1 if CORTE.day < 28 else CORTE.to_period("M")
MESES = pd.period_range(ULTIMO_CERRADO - MESES_SERIE + 1, ULTIMO_CERRADO, freq="M")
M12, M6, M3 = MESES[-12:], MESES[-6:], MESES[-3:]
MESES_RESTANTES = max(0.0, (pd.Timestamp(f"{ANIO_PPTO}-12-31") - CORTE).days / 30.44)
print(f"Kardex al {CORTE.date()} | último mes cerrado {ULTIMO_CERRADO} | meses que quedan del año {MESES_RESTANTES:.1f}")

# 2.1 Consumo a buses (91 salida − 92 devolución): la demanda es solo la de los almacenes de la operación
bus = k[k["Cod_Movimiento"].isin([91, 92])].copy()
bus["Und"] = bus["Cant_Salida"] - bus["Cant_Ingreso"]
bus["Val"] = bus["Valor_Salida"] - bus["Valor_Ingreso"]
flota = pd.read_parquet(ARCHIVO_FLOTA).drop_duplicates("Cod_Bus").set_index("Cod_Bus")["Tamaño_Bus"]
bus["Tam"] = bus["Cod_Bus"].map(flota).fillna("OTROS")
bus_proy = bus[~bus["Almacen"].isin(OPERACION) & bus["Mes"].isin(M12)]
bus = bus[bus["Almacen"].isin(OPERACION)]

und_mes = bus[bus["Mes"].isin(MESES)].pivot_table(index="Llave", columns="Mes", values="Und", aggfunc="sum").reindex(columns=MESES).fillna(0.0)
val_mes = bus[bus["Mes"].isin(MESES)].pivot_table(index="Llave", columns="Mes", values="Val", aggfunc="sum").reindex(columns=MESES).fillna(0.0)
und_pos = und_mes.clip(lower=0)
cons = pd.DataFrame(index=und_mes.index)
cons["Und12"] = und_pos[M12].sum(axis=1)
cons["Val12"] = val_mes[M12].clip(lower=0).sum(axis=1)
cons["Prom12"] = und_pos[M12].mean(axis=1)
cons["Prom6"] = und_pos[M6].mean(axis=1)
cons["Prom3"] = und_pos[M3].mean(axis=1)
cons["Desv12"] = und_pos[M12].std(axis=1, ddof=0)
cons["Meses12"] = (und_pos[M12] > 0).sum(axis=1)
cons["CV"] = (cons["Desv12"] / cons["Prom12"]).replace([np.inf, -np.inf], np.nan)
ult_cons = bus[bus["Und"] > 0].groupby("Llave")["Fecha"].max()
mes_actual = bus[bus["Mes"] == CORTE.to_period("M")].groupby("Llave")["Und"].sum()
b12 = bus[bus["Mes"].isin(M12)]
buses12 = b12[b12["Und"] > 0].groupby("Llave")["Cod_Bus"].nunique()
tam12 = b12.pivot_table(index="Llave", columns="Tam", values="Und", aggfunc="sum").fillna(0.0)
proy12 = bus_proy.groupby("Llave")["Und"].sum()

con_valor = cons[cons["Val12"] > 0].sort_values("Val12", ascending=False)
acum = con_valor["Val12"].cumsum() / con_valor["Val12"].sum()
cons["ABC"] = "-"
cons.loc[con_valor.index, "ABC"] = np.where(acum <= 0.80, "A", np.where(acum <= 0.95, "B", "C"))
cons["XYZ"] = np.where(cons["Val12"] <= 0, "-", np.where(cons["CV"] < 0.5, "X", np.where(cons["CV"] <= 1.0, "Y", "Z")))

# 2.2 Stock hoy: último saldo de cada producto en cada almacén
ult = k.sort_values("Fecha_Completa", kind="stable").groupby(["Empresa", "Almacen", "Cod_Producto"]).tail(1)
ult = ult[ult["Almacen"].isin(ALMACENES) & (ult["Cant_Saldo"].abs() > 1e-9)].copy()
ult["Grupo"] = ult["Almacen"].map(ALMACENES)
stock_und = ult.pivot_table(index="Llave", columns="Grupo", values="Cant_Saldo", aggfunc="sum").reindex(columns=GRUPOS_STOCK).fillna(0.0)
stock_val = ult.pivot_table(index="Llave", columns="Grupo", values="Valor_Saldo", aggfunc="sum").reindex(columns=GRUPOS_STOCK).fillna(0.0)

# 2.3 Compras (ingresos desde proveedor). Nacional = RUC peruano (10/20 + 9 dígitos)
comp = k[(k["Ubi_Origen"] == "Partner Locations/Vendors") & (k["Cant_Ingreso"] > 0) & k["Almacen"].isin(ALMACENES)].copy()
comp["Nacional"] = comp["Ruc"].astype(str).str.match(r"^(10|20)\d{9}$")
# Compras entre empresas del grupo (TVL a TATSA PERÚ, con margen): quedan en el historial pero no sirven como proveedor sugerido
comp["Intercompania"] = comp["Proveedor"].astype(str).str.upper().str.contains("TATSA PERU|TRANSVIAL LIMA", regex=True)
comp = comp.sort_values("Fecha_Completa")
oc_lin = comp.groupby(["Llave", "Cod_Orden_Compra"], dropna=False).agg(
    Fecha=("Fecha", "min"), Empresa=("Empresa", "first"), Proveedor=("Proveedor", "first"), Nacional=("Nacional", "first"),
    Inter=("Intercompania", "first"), Cant=("Cant_Ingreso", "sum"), Valor=("Valor_Ingreso", "sum"),
    Codigo=("Cod_Producto", "first")).reset_index().sort_values("Fecha")
oc_lin["Precio"] = oc_lin["Valor"] / oc_lin["Cant"]
ext_lin = oc_lin[~oc_lin["Inter"]]  # compras a terceros: base para proveedor sugerido, precios y bandera de importación
desde12, desde24 = M12[0].start_time, MESES[0].start_time
c12 = ext_lin[ext_lin["Fecha"] >= desde12]
precio_prom12 = c12.groupby("Llave")["Valor"].sum() / c12.groupby("Llave")["Cant"].sum()
compras_n12 = c12.groupby("Llave")["Cod_Orden_Compra"].nunique()
ultima = ext_lin.groupby("Llave").tail(1).set_index("Llave")
ultima_nac = ext_lin[ext_lin["Nacional"]].groupby("Llave").tail(1).set_index("Llave")
c24n = ext_lin[(ext_lin["Fecha"] >= desde24) & ext_lin["Nacional"]]
principal = c24n.groupby(["Llave", "Proveedor"])["Valor"].sum().reset_index().sort_values("Valor", ascending=False).drop_duplicates("Llave").set_index("Llave")["Proveedor"]
con_ext = set(ext_lin.loc[~ext_lin["Nacional"], "Llave"])
con_nac = set(ext_lin.loc[ext_lin["Nacional"], "Llave"])

# ----------------------------------------------------------------------------------------------------------------------
# 3. LEAD TIME REAL (ODOO, OPCIONAL): RQ (o aprobación de la OC) → primera recepción. Se excluyen RQ de regularización
# ----------------------------------------------------------------------------------------------------------------------
LEAD = pd.Series(dtype=float)
if USAR_ODOO:
    try:
        import xmlrpc.client
        from dotenv import load_dotenv
        load_dotenv()
        url, db = os.getenv("URL_ODOO"), os.getenv("BD_ODOO")
        usr, pwd = os.getenv(f"USER_{USUARIO_ODOO}"), os.getenv(f"PASSWORD_{USUARIO_ODOO}")
        uid = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common").authenticate(db, usr, pwd, {})
        mdl = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object", allow_none=True)

        def odoo(modelo, metodo, args, kw=None):
            kw = dict(kw or {})
            kw["context"] = {"lang": "es_PE", "allowed_company_ids": [1, 2], "active_test": False}
            return mdl.execute_kw(db, uid, pwd, modelo, metodo, args, kw)

        nombres = sorted({n for n in oc_lin["Cod_Orden_Compra"].dropna().unique() if str(n).startswith(("COMPR", "SERV"))})
        ocs = []
        for i in range(0, len(nombres), 1000):
            ocs += odoo("purchase.order", "search_read", [[["name", "in", nombres[i:i + 1000]]]],
                        {"fields": ["name", "company_id", "date_approve", "date_order", "requisition_id"]})
        rq_ids = sorted({o["requisition_id"][0] for o in ocs if o["requisition_id"]})
        rqs = odoo("purchase.requisition", "read", [rq_ids], {"fields": ["create_date", "prior"]}) if rq_ids else []
        rq_fecha = {r["id"]: (None if r["prior"] == "regilarizacion" else r["create_date"][:10]) for r in rqs}
        ocs = [o for o in ocs if not (o["requisition_id"] and rq_fecha.get(o["requisition_id"][0]) is None)]
        empresa = {1: "Tvl", 2: "Tatsa"}
        info_oc = pd.DataFrame([{
            "Cod_Orden_Compra": o["name"], "Empresa": empresa.get(o["company_id"][0]),
            "Inicio": rq_fecha.get(o["requisition_id"][0]) if o["requisition_id"] else (o["date_approve"] or o["date_order"])[:10],
        } for o in ocs])
        lt = oc_lin.merge(info_oc, on=["Cod_Orden_Compra", "Empresa"], how="inner")
        lt["Dias"] = (lt["Fecha"] - pd.to_datetime(lt["Inicio"])).dt.days
        lt = lt[(lt["Dias"] >= 0) & (lt["Dias"] <= 365) & (lt["Fecha"] >= desde24)]
        LEAD = lt.groupby("Llave")["Dias"].median()
        print(f"Odoo: {len(ocs):,} OC | lead time real para {len(LEAD):,} llaves (mediana {lt['Dias'].median():.0f} días)")
    except Exception as e:
        print("Odoo no disponible, se usa el lead time por defecto:", str(e)[:120])

# ----------------------------------------------------------------------------------------------------------------------
# 4. TRÁNSITO Y PRESUPUESTO (Saldo_Disponible_Real)
# ----------------------------------------------------------------------------------------------------------------------
oc_p = pd.read_excel(ARCHIVO_SALDO, sheet_name="OC_PENDIENTE_ENTREGA", engine="calamine")
rq_p = pd.read_excel(ARCHIVO_SALDO, sheet_name="RQ_SIN_OC", engine="calamine")
sal = pd.read_excel(ARCHIVO_SALDO, sheet_name="SALDO", engine="calamine")
for df in (oc_p, rq_p, sal):
    df["Llave"] = df["Cod_Producto"].map(LLAVE_DE).fillna(df["Cod_Producto"])
    df["CC"] = df["Cuenta Analítica"].astype(str).str.extract(r"\[(\w+)\]", expand=False)
    df["Oper"] = df["CC"].isin(CENTROS_OPERACION)
oc_op = oc_p[oc_p["Oper"]]
rq_op = rq_p[rq_p["Oper"]]
oc_und, oc_val = oc_op.groupby("Llave")["Cant_Pendiente"].sum(), oc_op.groupby("Llave")["Valor_Pendiente"].sum()
rq_und, rq_val = rq_op.groupby("Llave")["Cant_Pendiente"].sum(), rq_op.groupby("Llave")["Valor_Pendiente"].sum()
ocp_und = oc_p[~oc_p["Oper"]].groupby("Llave")["Cant_Pendiente"].sum()
rqp_und = rq_p[~rq_p["Oper"]].groupby("Llave")["Cant_Pendiente"].sum()

ppto = sal[(sal["Periodo"] == ANIO_PPTO) & (sal["Cuenta Analítica"].astype(str).str.startswith(CENTRO_PPTO))
           & sal["Situación presupuestaria"].astype(str).str.contains(SITUACION_PPTO, regex=False)].copy()
ppto["Ejec"] = ppto["Compras"] + ppto["Gasto_Sin_OC"]
ppto["Comp"] = ppto["RQ_COC"] + ppto["RQ_SOC"]
PPTO = {ll: [{"c": r.Cod_Producto, "pp": num(r.Presupuesto), "ej": num(r.Ejec), "co": num(r.Comp), "sa": num(r.Saldo),
              "do": num(r.Disponible_Odoo),
              "cu": num(r.RQ_COC_und), "su": num(r.RQ_SOC_und)} for r in g.sort_values("Presupuesto", ascending=False).itertuples()]
        for ll, g in ppto.groupby("Llave")}
print(f"Presupuesto {CENTRO_PPTO} {SITUACION_PPTO}: {len(ppto):,} SKUs | S/ {ppto['Presupuesto'].sum():,.0f} | saldo real S/ {ppto['Saldo'].sum():,.0f}")

# ----------------------------------------------------------------------------------------------------------------------
# 5. FILAS PARA LA PÁGINA
# ----------------------------------------------------------------------------------------------------------------------
activas = (set(und_pos.index[und_pos.sum(axis=1) > 0]) | set(stock_und.index[stock_und.sum(axis=1) > 0])
           | set(oc_und.index) | set(rq_und.index) | set(ocp_und.index) | set(rqp_und.index) | set(PPTO))
activas = [ll for ll in activas if ll in info.index and info.at[ll, "Familia"] not in FAMILIAS_EXCLUIDAS]
print(f"Llaves en la herramienta: {len(activas):,}")

VACIO = pd.DataFrame()
HIST = {ll: d.tail(6) for ll, d in oc_lin.groupby("Llave")}
TR_OC = {ll: d for ll, d in oc_p.groupby("Llave")}
TR_RQ = {ll: d for ll, d in rq_p.groupby("Llave")}


def v(serie, ll, d=0.0):
    return serie.get(ll, d)


filas = []
for ll in activas:
    i = info.loc[ll]
    c = cons.loc[ll] if ll in cons.index else None
    st = stock_und.loc[ll] if ll in stock_und.index else pd.Series(0.0, index=GRUPOS_STOCK)
    sv = stock_val.loc[ll] if ll in stock_val.index else pd.Series(0.0, index=GRUPOS_STOCK)
    u = ultima.loc[ll] if ll in ultima.index else None
    un = ultima_nac.loc[ll] if ll in ultima_nac.index else None
    hist, tr_oc, tr_rq = HIST.get(ll, VACIO), TR_OC.get(ll, VACIO), TR_RQ.get(ll, VACIO)
    filas.append({
        "k": ll, "n": i["Producto"], "f": texto(i["Familia"]), "f2": texto(i["Familia_2"]), "f3": texto(i["Familia_3"]),
        "u": texto(i["Uso"]), "um": texto(i["UdM"]), "cod": CODIGOS.get(ll, []),
        "abc": c["ABC"] if c is not None else "-", "xyz": c["XYZ"] if c is not None else "-",
        "p12": num(c["Prom12"] if c is not None else 0, 3), "p6": num(c["Prom6"] if c is not None else 0, 3),
        "p3": num(c["Prom3"] if c is not None else 0, 3), "sd": num(c["Desv12"] if c is not None else 0, 3),
        "u12": num(c["Und12"] if c is not None else 0), "v12": num(c["Val12"] if c is not None else 0),
        "m12": int(c["Meses12"]) if c is not None else 0, "cv": num(c["CV"], 2) if c is not None else None,
        "uc": fecha(ult_cons.get(ll)), "ma": num(mes_actual.get(ll, 0)), "bu": int(buses12.get(ll, 0)),
        "pu12": num(proy12.get(ll, 0)),
        "tam": {t: num(x) for t, x in (tam12.loc[ll].items() if ll in tam12.index else []) if x},
        "s": [num(x) for x in (und_mes.loc[ll].values if ll in und_mes.index else np.zeros(len(MESES)))],
        "st": {g: num(st.get(g, 0.0)) for g in GRUPOS_STOCK}, "sv": {g: num(sv.get(g, 0.0)) for g in GRUPOS_STOCK},
        "oc": num(v(oc_und, ll)), "ocv": num(v(oc_val, ll)), "rq": num(v(rq_und, ll)), "rqv": num(v(rq_val, ll)),
        "ocp": num(v(ocp_und, ll)), "rqp": num(v(rqp_und, ll)),
        "toc": [{"oc": r.Cod_OC, "rq": texto(r.Cod_RQ), "e": r.Empresa, "pv": texto(r.Proveedor), "cm": texto(r.Comprador),
                 "q": num(r.Cant_Pendiente), "v": num(r.Valor_Pendiente), "fp": fecha(r.Fecha_Entrega_Prevista),
                 "at": int(r.Dias_Atraso) if pd.notna(r.Dias_Atraso) else None, "cc": texto(r.CC), "p": not r.Oper}
                for r in tr_oc.itertuples()],
        "trq": [{"rq": r.Cod_RQ, "e": r.Empresa, "f": fecha(r.Fecha_RQ), "d": int(r.Dias_RQ) if pd.notna(r.Dias_RQ) else None,
                 "q": num(r.Cant_Pendiente), "v": num(r.Valor_Pendiente), "cc": texto(r.CC), "o": texto(r.Observaciones), "p": not r.Oper}
                for r in tr_rq.itertuples()],
        "uf": fecha(u["Fecha"]) if u is not None else None, "uoc": texto(u["Cod_Orden_Compra"]) if u is not None else None,
        "upv": texto(u["Proveedor"]) if u is not None else None, "upr": num(u["Precio"]) if u is not None else None,
        "ext": bool(u["Nacional"] == False) if u is not None else False,  # noqa: E712  última compra al exterior
        "sx": ll in con_ext and ll not in con_nac,                        # solo se ha comprado al exterior
        "pvn": texto(un["Proveedor"]) if un is not None else None, "prn": num(un["Precio"]) if un is not None else None,
        "fn": fecha(un["Fecha"]) if un is not None else None, "pvp": texto(principal.get(ll)),
        "pp12": num(precio_prom12.get(ll)), "nc12": int(compras_n12.get(ll, 0)), "lt": num(LEAD.get(ll), 0),
        "hc": [{"f": fecha(r.Fecha), "oc": texto(r.Cod_Orden_Compra), "e": r.Empresa, "pv": texto(r.Proveedor), "c": r.Codigo,
                "q": num(r.Cant), "pr": num(r.Precio), "x": not r.Nacional, "ic": bool(r.Inter)} for r in hist.iloc[::-1].itertuples()],
        "pr": PPTO.get(ll),
    })

meta = {
    "generado": datetime.now().strftime("%d/%m/%Y %H:%M"), "corte": CORTE.strftime("%d/%m/%Y"),
    "meses": [m.strftime("%b-%y") for m in MESES], "ultimo": ULTIMO_CERRADO.strftime("%b-%Y"), "anio": ANIO_PPTO,
    "lt_general": num(LEAD.median(), 0) if len(LEAD) else None, "odoo": bool(len(LEAD)), "grupos": GRUPOS_STOCK,
    "meses_rest": num(MESES_RESTANTES, 2), "centro": CENTRO_PPTO, "situacion": SITUACION_PPTO,
    "ppto_total": num(ppto["Presupuesto"].sum()), "ppto_saldo": num(ppto["Saldo"].sum()), "ppto_skus": int(len(ppto)),
}

plantilla = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), PLANTILLA), encoding="utf-8").read()
datos = json.dumps({"meta": meta, "filas": filas}, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
with open(ARCHIVO_SALIDA, "w", encoding="utf-8") as fh:
    fh.write(plantilla.replace("/*__DATOS__*/null", datos))
print(f"Listo: {ARCHIVO_SALIDA} ({os.path.getsize(ARCHIVO_SALIDA) / 1e6:.1f} MB, {len(filas):,} llaves)")
