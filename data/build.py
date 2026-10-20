"""Capa central de datos — Análisis Comercial Integra Cimentaciones.

Una sola fuente para la presentación (index.html) y el reporte PDF (reporte/reporte.html):

    python3 data/build.py

Lee:   data/ANALISIS_DE_DATOS_2.0.xlsx (ventas; gitignored)
       data/geo.json                             (coordenadas municipales)
       data/forecast_publicado_2026-07.json       (forecast del trimestre anterior)
       data/ghl_mastermetrics_2026-10-14.json     (Google Ads vía Master Metrics; gitignored)
       ghl/raw/<fecha>/summary.json               (GHL vía API; ghl/summarize.py --write; gitignored)
       data/baseline_clientify_2026-07.json       (actividades y funnel Clientify)
Escribe:
       data/data.js                 → window.DATA, público (sin cobranza por cliente)
       reporte/data_interno.js      → window.INTERNO, solo PDF (gitignored)

Todas las reglas de normalización aprobadas por Bernardo en la Fase 0 (2026-10-19)
viven aquí, en NORMALIZACION, y se documentan en el anexo de calidad de datos del PDF.
"""
import collections as C
import datetime as dt
import json
import re
import statistics
import sys
from pathlib import Path

sys.path.append(str(Path.home() / "Library/Python/3.9/lib/python/site-packages"))
import openpyxl  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
XLSX = DATA / "ANALISIS_DE_DATOS_2.0.xlsx"
CORTE = "2026-09-30"
MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO",
         "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
MES_CORTO = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

# ---------------------------------------------------------------------------
# Normalización (aprobada en Fase 0)
# ---------------------------------------------------------------------------
SERVICIO = {
    "MECANICA DE SUELOS": "Mecánica de suelos", "MECANCIA DE SUELOS": "Mecánica de suelos",
    "MECANICCA DE SUELOS": "Mecánica de suelos",
    "MECANICA Y TOPOGRAFIA": "Mecánica + topografía",
    "SERVICIOS INGENIERIA": "Servicios de ingeniería", "SERVICIOS DE INGENIERIA": "Servicios de ingeniería",
    "COMPACTACION": "Compactación", "COMPACTACIONES": "Compactación", "PRUEBAS DE COMPACTACION": "Compactación",
    "CONCRETO Y COMPACTACIONES": "Concreto + compactación",
    "CONCRETO": "Concreto", "DISEÑO DE CONCRETO": "Diseño de concreto", "EXTRACCION DE NUCLEOS": "Extracción de núcleos",
    "LABORATORIO": "Laboratorio", "LABORATORIOS": "Laboratorio", "LABORATORIO CALIDAD": "Laboratorio",
    "VISITA": "Visita técnica",
    "GEORRADAR": "Georradar (GPR)", "TRE Y GPR": "TRE + georradar",
    "TOMOGRAFIA DE RESISTIVIDAD": "Tomografía de resistividad (TRE)",
    "SEV": "SEV", "GEOFISICA": "Geofísica",
    "ANALISIS DE AGREGADOS": "Análisis de agregados", "GABINETE": "Gabinete",
    "SUBCONTRATADO": "Subcontratado", "TOPOGRAFIA": "Topografía", "EXPLORACION": "Exploración",
}
CANAL = {"WHATSAPP": "WhatsApp", "INBOX": "WhatsApp", "LLAMADA": "Llamada", "WEB": "Web", "CORREO": "Correo",
         "FACEBOOK": "Facebook", "RECURRENTE": "Recurrente", "RECOMENDADO": "Recomendado"}
RELACION = {"Recurrente", "Recomendado"}
TIPO = {"NUEVO": "Nuevo", "RECURRENTE": "Recurrente", "RECOMENDADO": "Recomendado",
        # mayo-2026: la columna Tipo trae la ubicación de Tiendas 3B (canal = RECURRENTE)
        "PANALES": "Recurrente", "HUICHAPAN": "Recurrente", "SAN MIGUEL": "Recurrente"}
CLIENTE_ALIAS = {"TIENDAS 3 B": "TIENDAS 3B", "TIENDAS TRES B": "TIENDAS 3B", "INTELISITE": "INTELLISITE"}
CLIENTE_NOMBRE = {"TIENDAS 3B": "Tiendas 3B", "INTELLISITE": "Intellisite"}
TITULOS = r"^(ING|ARQ|SR|SRA|SRTA|LIC|DR|DRA|C)\.?\s+"

# Ubicación original (limpia) → sitios [(municipio, estado, localidad, peso)]; None = no geográfico
EDO = {"CDMX": "Ciudad de México", "EDOMEX": "Estado de México"}
UBIC = {
    "ACAPULCO, GUERRERO": ("Acapulco de Juárez", "Guerrero"),
    "ACOCUL DE CEBOLLETAS, HIDALGO": ("Metepec", "Hidalgo", "Acocul Cebolletas"),
    "ACOLAPA, MORELOS": ("Jiutepec", "Morelos", "Acolapan"),
    "ACTOPAN, HIDALGO": ("Actopan", "Hidalgo"),
    "AICM, CDMX": ("Venustiano Carranza", "Ciudad de México", "AICM"),
    "ALDAMA, GUANAJUATO": ("Irapuato", "Guanajuato", "Aldama"),
    "ALMOLOYA, EDOMEX": ("Almoloya de Juárez", "Estado de México"),
    "ALVARO OBREGON, CDMX": ("Álvaro Obregón", "Ciudad de México"),
    "APAXCO, EDOMEX": ("Apaxco", "Estado de México"),
    "ARANDAS, JALISCO": ("Arandas", "Jalisco"),
    "ATITALAQUI, HIDALGO": ("Atitalaquia", "Hidalgo"), "ATITALAQUIA, HIDALGO": ("Atitalaquia", "Hidalgo"),
    "ATIZAPAN, EDOMEX": ("Atizapán de Zaragoza", "Estado de México"),
    "ATLACOMULCO, EDOMEX": ("Atlacomulco", "Estado de México"),
    "ATOTONILCO DE TULA, HIDALGO": ("Atotonilco de Tula", "Hidalgo"),
    "ATOTONILCO, HIDALGO": ("Atotonilco de Tula", "Hidalgo"),
    "AZCAPOTZALCO, CDMX": ("Azcapotzalco", "Ciudad de México"),
    "BAJIO": ("Silao de la Victoria", "Guanajuato", "Bajío"),
    "BARRANCA PIÑA, GUERRERO": ("Acatepec", "Guerrero", "Barranca Piña"),
    "BENITO JUAREZ, CDMX": ("Benito Juárez", "Ciudad de México"),
    "CANTA EL AGUA, MICHOACAN": (None, "Michoacán", "Canta el Agua"),
    "CAPULIN, EDOMEX": (None, "Estado de México", "Capulín"),
    "CD SAHAGUN, HIDALGO": ("Tepeapulco", "Hidalgo", "Cd. Sahagún"),
    "CDMX": (None, "Ciudad de México"),
    "CELAYA, GUANAJUATO": ("Celaya", "Guanajuato"),
    "CHALCO, EDOMEX": ("Chalco", "Estado de México"),
    "CHALMA, EDOMEX": ("Malinalco", "Estado de México", "Chalma"),
    "CHILPANCINGO GUERRERO": ("Chilpancingo de los Bravo", "Guerrero"),
    "CHILPANCINGO, GUERRERO": ("Chilpancingo de los Bravo", "Guerrero"),
    "CHIMALHUACAN, EDOMEX": ("Chimalhuacán", "Estado de México"),
    "COACALCO, EDOMEX": ("Coacalco de Berriozábal", "Estado de México"),
    "CUAJIMALPA, CDMX": ("Cuajimalpa de Morelos", "Ciudad de México"),
    "CUATITLAN, EDOMEX": ("Cuautitlán", "Estado de México"), "CUAUTITLAN, EDOMEX": ("Cuautitlán", "Estado de México"),
    "CUAUHTEMOC, CDMX": ("Cuauhtémoc", "Ciudad de México"),
    "CULHUACAN, CDMX": ("Iztapalapa", "Ciudad de México", "Culhuacán"),
    "DIAZ ORDAZ, TAMPAULIPAS": ("Gustavo Díaz Ordaz", "Tamaulipas"),
    "ECATEPEC, EDOMEX": ("Ecatepec de Morelos", "Estado de México"),
    "EDOMEX": (None, "Estado de México"),
    "EL LLANO, HIDALGO (EL LLANO)": ("Tula de Allende", "Hidalgo", "El Llano"),
    "EL MARQUEZ, QUERETARO": ("El Marqués", "Querétaro"),
    "EL SALTO, DURANGO": ("Pueblo Nuevo", "Durango", "El Salto"),
    "GABINETE": None,
    "GAM, CDMX": ("Gustavo A. Madero", "Ciudad de México"),
    "GUSTAVO A. MADERO, CDMX": ("Gustavo A. Madero", "Ciudad de México"),
    "GUADALUPE CENTRO": ("Chapantongo", "Hidalgo", "Guadalupe Centro"),
    "GUADALUPE CENTRO, HIDALGO": ("Chapantongo", "Hidalgo", "Guadalupe Centro"),
    "GUERRERO": (None, "Guerrero"),
    "HUEHUETOCA, EDOMEX (CITARA)": ("Huehuetoca", "Estado de México", "Citara"),
    "HUICHAPAN, HIDALGO (ROJO GOMEZ)": ("Huichapan", "Hidalgo", "Rojo Gómez"),
    "HUICHAPAN, HIDALGO": ("Huichapan", "Hidalgo"),
    "HUITEL, HIDALGO": ("Tezontepec de Aldama", "Hidalgo", "Huitel"),
    "HUIXQUILUCAN, EDOMEX": ("Huixquilucan", "Estado de México"),
    "IXMIQUILPAN, HIDALGO (HGO MADERO)": ("Ixmiquilpan", "Hidalgo", "Hgo. Madero"),
    "IXMIQUILPAN, HIDALGO": ("Ixmiquilpan", "Hidalgo"),
    "IXTAPAN DE LA SAL, EDOMEX": ("Ixtapan de la Sal", "Estado de México"),
    "IZTACALCO, CDMX": ("Iztacalco", "Ciudad de México"),
    "IZTAPALAPA": ("Iztapalapa", "Ciudad de México"), "IZTAPALAPA, CDMX": ("Iztapalapa", "Ciudad de México"),
    "JALMOLONGA, EDOMEX": ("Malinalco", "Estado de México", "Jalmolonga"),
    "JILOTZINGO, EDOMEX": ("Jilotzingo", "Estado de México"),
    "LABORATORIO CENTRAL": None, "LABORATORO": None,
    "LAZARO CARDENA, MICHACAN": ("Lázaro Cárdenas", "Michoacán"),
    "LOS REYES ACAQUILPAN, EDOMEX": ("La Paz", "Estado de México", "Los Reyes Acaquilpan"),
    "MELCHOR OCAMPO, EDOMEX": ("Melchor Ocampo", "Estado de México"),
    "MEXICALI, BAJA CALIFORNIA": ("Mexicali", "Baja California"),
    "MIGUEL HIDALGO, CDMX": ("Miguel Hidalgo", "Ciudad de México"),
    "MIXQUIAHUALA, HIDALGO": ("Mixquiahuala de Juárez", "Hidalgo"),
    "NAUCALPAN, EDOMEX": ("Naucalpan de Juárez", "Estado de México"),
    "NEZAHUALCOYOTL, CDMX": ("Nezahualcóyotl", "Estado de México"),
    "NICOLAS ROMERO, EDOMEX": ("Nicolás Romero", "Estado de México"),
    "NOPALA, HIDALGO": ("Nopala de Villagrán", "Hidalgo"),
    "NUEVO LEON, MONTERREY": ("Monterrey", "Nuevo León"),
    "OAXACA, OAXACA": ("Oaxaca de Juárez", "Oaxaca"),
    "OJO DE AGUA, EDOMEX": ("Tecámac", "Estado de México", "Ojo de Agua"),
    "PACHUCA DE SOTO, HIDALGO": ("Pachuca de Soto", "Hidalgo"), "PACHUCA, HIDALGO": ("Pachuca de Soto", "Hidalgo"),
    "SAN FERNANDO, PACHUCA": ("Pachuca de Soto", "Hidalgo", "San Fernando"),
    "PANALES, HIDALGO": ("Ixmiquilpan", "Hidalgo", "Panales"),
    "POPOCATEPETL, PUEBLA": (None, "Puebla", "Popocatépetl"),
    "PTO ESCONDIDO, OAXACA": ("San Pedro Mixtepec", "Oaxaca", "Puerto Escondido"),
    "PUERTO ESCONDIDO, OAXACA": ("San Pedro Mixtepec", "Oaxaca", "Puerto Escondido"),
    "PUEBLA, PUEBLA": ("Puebla", "Puebla"),
    "QUERETARO, QUERETARO": ("Querétaro", "Querétaro"), "QUERETATO,QUERETARO": ("Querétaro", "Querétaro"),
    "REAL DE CATORCE": ("Catorce", "San Luis Potosí", "Real de Catorce"),
    "REAL DE CATORCE, SAN LUIS POTOSI": ("Catorce", "San Luis Potosí", "Real de Catorce"),
    "REAL DEL MONTE, HIDALGO": ("Mineral del Monte", "Hidalgo", "Real del Monte"),
    "SAN JOSE ATLAN, HIDALGO": ("Huichapan", "Hidalgo", "San José Atlán"),
    "SAN JOSE ITRUBIDE, GUANAJUATO": ("San José Iturbide", "Guanajuato"),
    "SAN JOSE ITURBIDE, GTO": ("San José Iturbide", "Guanajuato"),
    "SAN JOSE ITURBIDE, GUANAJUATO": ("San José Iturbide", "Guanajuato"),
    "SAN MIGUEL VINDHO, HIDALGO": ("Tula de Allende", "Hidalgo", "San Miguel Vindhó"),
    "SAN SEBASTIAN TENOCHTITLAN, HIDALGO": ("Nopala de Villagrán", "Hidalgo", "San Sebastián Tenochtitlán"),
    "SANTA ANA BATHA, HIDALGO (HGO. SANTA ANA BATHA #4880)": ("Chilcuautla", "Hidalgo", "Santa Ana Batha"),
    "SANTA ELENA, MICHOACAN": (None, "Michoacán", "Santa Elena"),
    "SANTA FE, CDMX": ("Álvaro Obregón", "Ciudad de México", "Santa Fe"),
    "SANTA LUCIA, SINALOA": ("Concordia", "Sinaloa", "Santa Lucía"),
    "TAMIAHUA, VERACRUZ": ("Tamiahua", "Veracruz"),
    "TAMPICO, TAMAULIPAS": ("Tampico", "Tamaulipas"),
    "TASQUILLO 2": ("Tasquillo", "Hidalgo"), "TASQUILLO 2, HIDALGO": ("Tasquillo", "Hidalgo"),
    "TASQUILLO, HIDALGO (HGO. TASQUILLO #4885)": ("Tasquillo", "Hidalgo"),
    "TECOMAN, COLIMA": ("Tecomán", "Colima"),
    "TECOZAUTLA CRUCERO, HIDALGO": ("Tecozautla", "Hidalgo", "Crucero"),
    "TECOZAUTLA LA CURVA, HIDALGO": ("Tecozautla", "Hidalgo", "La Curva"),
    "TEMASCALTEPEC, EDOMEX": ("Temascaltepec", "Estado de México"),
    "TEOTIHUACAN, EDOMEX": ("Teotihuacán", "Estado de México"),
    "TEPEJI, HIDALGO": ("Tepeji del Río de Ocampo", "Hidalgo"),
    "TEPETLAOXTOC, EDOMEX": ("Tepetlaoxtoc", "Estado de México"),
    "TEXCOCO, EDOMEX": ("Texcoco", "Estado de México"),
    "TEZONTEPEC, HIDALGO": ("Villa de Tezontepec", "Hidalgo"),
    "VILLA TEZONTEPEC, HIDALGO": ("Villa de Tezontepec", "Hidalgo"),
    "TEZOYUCA, EDOMEX": ("Tezoyuca", "Estado de México"),
    "TIZAYUCA, HIDALGO": ("Tizayuca", "Hidalgo"),
    "TLAHUAC, EDOMEX": ("Tláhuac", "Ciudad de México"),
    "TLALNEPANTLA, EDOMEX": ("Tlalnepantla de Baz", "Estado de México"),
    "TLALPAN, CDMX": ("Tlalpan", "Ciudad de México"),
    "TLALTIZAPAN, MORELOS": ("Tlaltizapán de Zapata", "Morelos"),
    "TLAPACOYAN, VERACRUZ": ("Tlapacoyan", "Veracruz"),
    "TOLUCA, EDOMEX": ("Toluca", "Estado de México"),
    "TULA DE ALLENDE HIDALGO (BOJAY)": ("Tula de Allende", "Hidalgo", "Bojay"),
    "TULA, HIDALGO": ("Tula de Allende", "Hidalgo"),
    "TULTITLAN, EDOMEX": ("Tultitlán", "Estado de México"),
    "TUXPAN, VERACRUZ": ("Tuxpan", "Veracruz"),
    "VALLE DE CHALCO, EDOMEX": ("Valle de Chalco Solidaridad", "Estado de México"),
    "VERACRUZ, VERACRUZ": ("Veracruz", "Veracruz"),
    "VILLA DE REYES, SAN LUIS POTOSI": ("Villa de Reyes", "San Luis Potosí"),
    "XALATLACO, EDOMEX": ("Xalatlaco", "Estado de México"),
    "XALOSTOC, EDOMEX": ("Ecatepec de Morelos", "Estado de México", "Xalostoc"),
    "XICOTEPEC, PUEBLA": ("Xicotepec", "Puebla"),
    "XOCHITLAN, HIDALGO": ("Progreso de Obregón", "Hidalgo", "Xochitlán"),
    "ZACATECAS, ZACATECAS": ("Zacatecas", "Zacatecas"),
    "ZACATEPEC, MORELOS": ("Zacatepec", "Morelos"),
    "ZUMPANGO, EDOMEX": ("Zumpango", "Estado de México"),
}
# Ubicaciones que son dos sitios de una misma venta: el monto se divide entre ambos
UBIC_MULTI = {"LEON - FRESNILLO": [("León", "Guanajuato", None), ("Fresnillo", "Zacatecas", None)]}


def clean(v):
    return re.sub(r"\s+", " ", str(v)).strip() if v not in (None, "") else ""


def num(v):
    return float(v) if isinstance(v, (int, float)) else 0.0


# ---------------------------------------------------------------------------
# 1. Lectura del Excel (bloques mensuales con encabezados repetidos)
# ---------------------------------------------------------------------------
def leer_excel():
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    filas, subtotales = [], {}
    for ws in wb.worksheets:
        year, mes, hdr = int(ws.title), None, None
        for i, r in enumerate(ws.iter_rows(values_only=True), 1):
            r = list(r)
            if isinstance(r[1], str) and r[1].strip() in MESES:
                mes, hdr = MESES.index(r[1].strip()) + 1, None
                continue
            if r[1] == "# De venta":
                hdr = r
                continue
            if mes and hdr is None and (year, mes) not in subtotales:
                nums = [v for v in r if isinstance(v, (int, float))]
                if len(nums) >= 2:
                    subtotales[(year, mes)] = (int(nums[-2]), float(nums[-1]))
            if hdr is None or not isinstance(r[1], (int, float)) or r[9] in (None, ""):
                continue
            d = {h: v for h, v in zip(hdr, r) if h}
            d.update(_year=year, _mes=mes, _fila=i)
            filas.append(d)
    return filas, subtotales


# ---------------------------------------------------------------------------
# 2. Normalización → registros limpios
# ---------------------------------------------------------------------------
def normalizar(filas):
    log, regs = [], []
    for f in filas:
        y, m, n = f["_year"], f["_mes"], int(f["# De venta"])
        rid = f"{y}-{m:02d}-{n:02d}"
        canal_o, tipo_o = clean(f.get("Canal de llegada")).upper(), clean(f.get("Tipo de Cliente")).upper()
        if canal_o == "NUEVO" and tipo_o in CANAL:  # marzo-2026: columnas intercambiadas
            log.append((rid, "Canal/Tipo intercambiados", f"{canal_o}/{tipo_o}", f"{tipo_o}/NUEVO"))
            canal_o, tipo_o = tipo_o, "NUEVO"
        if tipo_o in ("PANALES", "HUICHAPAN", "SAN MIGUEL"):
            log.append((rid, "Tipo de cliente con ubicación", tipo_o, "RECURRENTE"))

        ubic = clean(f.get("Ubicación")).upper()
        cliente_raw = clean(f.get("Nombre del cliente"))
        if not ubic and tipo_o == "SAN MIGUEL":
            log.append((rid, "Ubicación vacía", "(vacío)", "SAN MIGUEL VINDHO, HIDALGO"))
            ubic = "SAN MIGUEL VINDHO, HIDALGO"
        if ubic in UBIC_MULTI:
            sitios = [dict(mun=a, edo=b, loc=c, peso=1 / len(UBIC_MULTI[ubic])) for a, b, c in UBIC_MULTI[ubic]]
        else:
            if ubic not in UBIC:
                raise SystemExit(f"Ubicación sin regla de normalización: {ubic!r} ({rid})")
            t = UBIC[ubic]
            sitios = [] if t is None else [dict(mun=t[0], edo=t[1], loc=t[2] if len(t) > 2 else None, peso=1.0)]

        serv_o = clean(f.get("Servicio Contratado")).upper()
        key = re.sub(TITULOS, "", cliente_raw.upper())
        key = CLIENTE_ALIAS.get(key, key)

        costo = num(f["Costo del servicio"])
        est = clean(f.get("Estatus Venta")).upper()
        liq, cred = clean(f.get("Liquidado")), clean(f.get("Credito"))
        if y == 2025:
            if est.startswith("SE PERDIERON"):
                perdida = float(re.sub(r"[^\d.]", "", est.split("PERDIERON")[1]))
                pago, cobrado = "Pérdida parcial", costo - perdida
            elif est.startswith("SE PERDIO"):
                perdida, pago, cobrado = costo, "Perdida", 0.0
            else:
                perdida, pago, cobrado = 0.0, "Liquidado", costo
            restante = 0.0
        else:
            perdida = 0.0
            if liq == "Liquidado":
                pago, cobrado, restante = "Liquidado", costo, 0.0
            elif cred == "Credito" and not isinstance(f.get("Abonado"), (int, float)):
                pago, cobrado, restante = "Crédito", 0.0, costo
            else:
                pago, cobrado, restante = "Parcial", num(f.get("Abonado")), num(f.get("Restante"))

        regs.append(dict(
            id=rid, year=y, month=m, num=n,
            fecha_cot=f["Fecha de Cotizacion"].date().isoformat() if isinstance(f.get("Fecha de Cotizacion"), dt.datetime) else None,
            fecha_cierre=f["Fecha de Cierre"].date().isoformat() if isinstance(f.get("Fecha de Cierre"), dt.datetime) else None,
            dias=f.get("Dias Transcurridos") if isinstance(f.get("Dias Transcurridos"), (int, float)) and f.get("Fecha de Cierre") else None,
            ubicacion=ubic, sitios=sitios,
            servicio=SERVICIO[serv_o], servicio_orig=serv_o,
            canal=CANAL[canal_o], canal_orig=canal_o,
            tipo=TIPO[tipo_o],
            costo=round(costo, 2), cliente=cliente_raw, cliente_key=key,
            pago=pago, cobrado=round(cobrado, 2), restante=round(restante, 2), perdida=round(perdida, 2),
            nota="",
        ))
    # WattsnBytes ago-2026 #14 + #15: una sola venta con 2 servicios (Santa Lucía, Sin. y El Salto, Dgo.)
    a = next(r for r in regs if r["id"] == "2026-08-14")
    b = next(r for r in regs if r["id"] == "2026-08-15")
    a["costo"] = round(a["costo"] + b["costo"], 2)
    a["cobrado"] = round(a["cobrado"] + b["cobrado"], 2)
    a["sitios"] = [dict(s, peso=0.5) for s in a["sitios"] + b["sitios"]]
    a["ubicacion"] = f'{a["ubicacion"]} + {b["ubicacion"]}'
    a["nota"] = "Una venta con 2 servicios (filas #14 y #15 del Excel)"
    regs.remove(b)
    log.append(("2026-08-14/15", "Dos filas = una venta (WattsnBytes)", "2 ventas × $45,146.50", "1 venta · $90,293"))
    return regs, log


# ---------------------------------------------------------------------------
# 3. Métricas
# ---------------------------------------------------------------------------
def tot(rs):
    return dict(ventas=len(rs), monto=round(sum(r["costo"] for r in rs), 2))


def sel(regs, y, meses):
    return [r for r in regs if r["year"] == y and r["month"] in meses]


def mensual(regs):
    out = []
    for (y, m) in sorted({(r["year"], r["month"]) for r in regs}):
        rs = sel(regs, y, [m])
        rel = [r for r in rs if r["canal"] in RELACION]
        t, tr = tot(rs), tot(rel)
        out.append(dict(ym=f"{y}-{m:02d}", label=f"{MES_CORTO[m - 1]}-{str(y)[2:]}", **t,
                        ventas_rel=tr["ventas"], monto_rel=tr["monto"],
                        pct_rel_ventas=round(tr["ventas"] / t["ventas"], 4), pct_rel_monto=round(tr["monto"] / t["monto"], 4)))
    return out


def clientes(regs):
    g = C.defaultdict(list)
    for r in regs:
        g[r["cliente_key"]].append(r)
    out = []
    for k, rs in g.items():
        rs.sort(key=lambda r: (r["fecha_cierre"] or f'{r["year"]}-{r["month"]:02d}-15'))
        last = rs[-1]["fecha_cierre"] or f'{rs[-1]["year"]}-{rs[-1]["month"]:02d}-15'
        canales = C.Counter(r["canal"] for r in rs)
        estados = C.Counter(s["edo"] for r in rs for s in r["sitios"] if s["edo"])
        out.append(dict(
            key=k, nombre=CLIENTE_NOMBRE.get(k, rs[-1]["cliente"]), compras=len(rs), monto=round(sum(r["costo"] for r in rs), 2),
            monto_2025=round(sum(r["costo"] for r in rs if r["year"] == 2025), 2),
            monto_2026=round(sum(r["costo"] for r in rs if r["year"] == 2026), 2),
            compras_2025=sum(r["year"] == 2025 for r in rs), compras_2026=sum(r["year"] == 2026 for r in rs),
            ultima=last, dias_sin_compra=(dt.date.fromisoformat(CORTE) - dt.date.fromisoformat(last)).days,
            estados=[e for e, _ in estados.most_common(2)], canal_principal=canales.most_common(1)[0][0],
        ))
    out.sort(key=lambda c: -c["monto"])
    return out


def agrupa(regs, campo):
    g = C.defaultdict(list)
    for r in regs:
        g[r[campo]].append(r)
    out = [dict(nombre=k, **tot(v), ticket=round(sum(r["costo"] for r in v) / len(v), 2)) for k, v in g.items()]
    return sorted(out, key=lambda x: -x["monto"])


def por_estado(regs):
    g = C.defaultdict(lambda: dict(ventas=0.0, monto=0.0, ventas_2025=0.0, ventas_2026=0.0))
    for r in regs:
        for s in r["sitios"]:
            e = g[s["edo"]]
            e["ventas"] += s["peso"]
            e["monto"] += r["costo"] * s["peso"]
            e[f'ventas_{r["year"]}'] += s["peso"]
    out = [dict(estado=k, **{kk: round(vv, 2) for kk, vv in v.items()}) for k, v in g.items()]
    return sorted(out, key=lambda x: -x["monto"])


def por_municipio(regs, coords):
    g = C.defaultdict(lambda: dict(ventas=0.0, monto=0.0, monto_2025=0.0, monto_2026=0.0, locs=set()))
    for r in regs:
        for s in r["sitios"]:
            if not s["mun"]:
                continue
            e = g[(s["mun"], s["edo"])]
            e["ventas"] += s["peso"]
            e["monto"] += r["costo"] * s["peso"]
            e[f'monto_{r["year"]}'] += r["costo"] * s["peso"]
            if s["loc"]:
                e["locs"].add(s["loc"])
    out = []
    for (mun, edo), v in g.items():
        c = coords.get(f"{mun}, {edo}")
        if not c:
            raise SystemExit(f"Sin coordenadas para {mun}, {edo}: agrégalas a data/geo.json")
        out.append(dict(municipio=mun, estado=edo, lat=c["lat"], lon=c["lon"], ventas=round(v["ventas"], 2),
                        monto=round(v["monto"], 2), monto_2025=round(v["monto_2025"], 2), monto_2026=round(v["monto_2026"], 2),
                        localidades=sorted(v["locs"])))
    return sorted(out, key=lambda x: -x["monto"])


def ciclo_por_servicio(regs):
    g = C.defaultdict(list)
    for r in regs:
        if r["dias"] is not None:
            g[r["servicio"]].append(r["dias"])
    out = [dict(servicio=k, ventas=len(v), prom=round(statistics.mean(v), 1), mediana=statistics.median(v))
           for k, v in g.items() if len(v) >= 2]
    return sorted(out, key=lambda x: -x["ventas"])


def forecast_q4(serie_2026):
    """Tendencia lineal sobre los montos mensuales de 2026 (decisión: solo tendencia de este año)."""
    y = [p["monto"] for p in serie_2026]
    x = list(range(1, len(y) + 1))
    mx, my = statistics.mean(x), statistics.mean(y)
    b = sum((a - mx) * (c - my) for a, c in zip(x, y)) / sum((a - mx) ** 2 for a in x)
    a0 = my - b * mx
    res = [c - (a0 + b * a) for a, c in zip(x, y)]
    sd = statistics.stdev(res)
    r2 = 1 - sum(e * e for e in res) / sum((c - my) ** 2 for c in y)
    meses = ["2026-10", "2026-11", "2026-12"]
    base = [a0 + b * m for m in (10, 11, 12)]
    esc = {"pesimista": [v - sd for v in base], "base": base, "optimista": [v + sd for v in base]}
    return dict(
        metodo="Regresión lineal sobre los montos mensuales ene–sep 2026; escenarios = tendencia ± 1 desviación estándar de los residuales por mes",
        pendiente_mes=round(b, 2), intercepto=round(a0, 2), sd_residual=round(sd, 2), r2=round(r2, 3),
        tendencia_2026=[round(a0 + b * m, 2) for m in x],
        meses=meses, escenarios={k: [round(v, 2) for v in vs] for k, vs in esc.items()},
        q4={k: round(sum(vs), 2) for k, vs in esc.items()},
    )


def main():
    filas, subtot = leer_excel()
    regs, log = normalizar(filas)
    coords = json.loads((DATA / "geo.json").read_text())["coords"]
    fpub = json.loads((DATA / "forecast_publicado_2026-07.json").read_text())
    mm = json.loads((DATA / "ghl_mastermetrics_2026-10-14.json").read_text())
    base_cf = json.loads((DATA / "baseline_clientify_2026-07.json").read_text())
    previo = json.loads((DATA / "corte_anterior_2026-07.json").read_text())
    # GHL vía API: resumen completo de la extracción más reciente (ghl/summarize.py --write; gitignored)
    ghl_dirs = sorted(p for p in (ROOT / "ghl" / "raw").glob("20*") if (p / "summary.json").exists())
    if not ghl_dirs:
        raise SystemExit("Falta ghl/raw/<fecha>/summary.json: corre ghl/extract.py y ghl/summarize.py --write")
    gs = json.loads((ghl_dirs[-1] / "summary.json").read_text())
    gs_corte = ghl_dirs[-1].name

    Q3 = [7, 8, 9]
    q3_26, q3_25 = sel(regs, 2026, Q3), sel(regs, 2025, Q3)
    y26, y25 = sel(regs, 2026, range(1, 13)), sel(regs, 2025, range(1, 13))
    serie = mensual(regs)
    s26 = [p for p in serie if p["ym"].startswith("2026")]
    rel = lambda rs: [r for r in rs if r["canal"] in RELACION]
    mkt = lambda rs: [r for r in rs if r["canal"] not in RELACION]

    # --- Forecast publicado vs real Q3
    fq3 = []
    for m in Q3:
        ym = f"2026-{m:02d}"
        fq3.append(dict(ym=ym, mes=MES_CORTO[m - 1], real=tot(sel(regs, 2026, [m]))["monto"],
                        **{k: round(fpub["mensual"][k][ym], 2) for k in ("pesimista", "base", "optimista")}))
    t = lambda k: round(sum(p[k] for p in fq3), 2)
    proy = dict(meses=fq3, real=t("real"), optimista=t("optimista"), base=t("base"), pesimista=t("pesimista"))
    proy.update(dif_opt=round(proy["real"] - proy["optimista"], 2), pct_opt=round(proy["real"] / proy["optimista"] - 1, 4),
                dif_base=round(proy["real"] - proy["base"], 2), pct_base=round(proy["real"] / proy["base"] - 1, 4),
                q4_publicado={k: round(sum(fpub["mensual"][k][f"2026-{m}"] for m in (10, 11, 12)), 2)
                              for k in ("pesimista", "base", "optimista")},
                q4_publicado_mensual={k: [round(fpub["mensual"][k][f"2026-{m}"], 2) for m in (10, 11, 12)]
                                      for k in ("pesimista", "base", "optimista")})

    # --- Cobranza 2026
    cob = dict(total=tot(y26)["monto"], cobrado=round(sum(r["cobrado"] for r in y26), 2),
               pendiente=round(sum(r["restante"] for r in y26), 2),
               parcial=tot([r for r in y26 if r["pago"] == "Parcial"]),
               saldo_parcial=round(sum(r["restante"] for r in y26 if r["pago"] == "Parcial"), 2),
               credito=tot([r for r in y26 if r["pago"] == "Crédito"]))
    cob["pct_cobrado"] = round(cob["cobrado"] / cob["total"], 4)
    perdidas_2025 = [dict(id=r["id"], cliente=r["cliente"], costo=r["costo"], perdida=r["perdida"], estatus=r["pago"])
                     for r in y25 if r["perdida"]]

    # --- Clientes
    cli = clientes(regs)
    cross = [c for c in cli if c["compras_2025"] and c["compras_2026"]]
    cli_q3 = clientes(q3_26)

    # --- GHL: cohorte Q3 (opps nativas creadas jul–sep, estado al corte de la extracción) + Google Ads (Master Metrics)
    ga, dem, ges = mm["google_ads"], mm["demanda"], mm["gestion"]
    coh = [gs["cohorte_por_mes"].get(f"2026-{m:02d}", {}) for m in Q3]
    q3c = {k: round(sum(c.get(k, 0) for c in coh), 2) for k in ("opps", "valor", "open", "won", "lost", "valor_won")}
    q3c.update(por_mes={f"2026-{m:02d}": c.get("opps", 0) for m, c in zip(Q3, coh)})
    act_keys = ("whatsapp", "whatsapp_salientes", "llamadas", "llamadas_salientes", "email", "email_salientes",
                "tareas_creadas", "tareas_completadas")
    api = dict(corte=gs_corte, fuente=f"API Go High Level (extracción {gs_corte})", q3=q3c,
               pipeline_abierto=dict(opps=gs["opps_abiertas"], valor=gs["pipeline_abierto_bruto"]),
               actividades={m: {k: v.get(k, 0) for k in act_keys} for m, v in gs["actividades_por_mes"].items() if m >= "2026-07"},
               llamadas_completadas=gs["llamadas_por_estado"].get("completed", 0),
               nota="Opps migradas de Clientify excluidas (migración parcial: fecha y estado no confiables)")
    mkt_q3 = tot([r for r in q3_26 if r["canal"] in ("WhatsApp", "Web")])
    crm = dict(
        corte=gs_corte, estatus="API GHL",
        tasa_cierre=round(q3c["won"] / q3c["opps"], 4),
        tasa_cierre_resueltas=round(q3c["won"] / (q3c["won"] + q3c["lost"]), 4),
        pct_abiertas=round(q3c["open"] / q3c["opps"], 4),
        mm_ref=dict(corte=mm["corte"], ganadas=ges["ganadas"], abiertas=ges["abiertas"], tasa=round(ges["ganadas"] / dem["opps"], 4)),
        cpl=round(ga["inversion"] / dem["opps"], 2),
        cpl_mes={m: round(ga["inversion_por_mes"][m] / dem["opps_por_mes"][m], 2) for m in ga["inversion_por_mes"]},
        costo_contacto=round(ga["inversion"] / dem["contactos"], 2),
        costo_venta_crm=round(ga["inversion"] / q3c["won"], 2),
        roas_excel=dict(valor=round(mkt_q3["monto"] / ga["inversion"], 2), ventas_whatsapp_web=mkt_q3,
                        nota="ventas WhatsApp (INBOX) + Web del Excel, Q3 2026 ÷ inversión Google Ads Q3"),
        flujo=dict(h1_clientify=dict(opps=base_cf["funnel_2026_h1"]["opps"], meses=6,
                                     ganadas=base_cf["funnel_2026_h1"]["ganadas"], abiertas=base_cf["funnel_2026_h1"]["abiertas"]),
                   q3_ghl=dict(opps=q3c["opps"], meses=3, ganadas=q3c["won"], abiertas=q3c["open"])),
    )
    for k in ("h1_clientify", "q3_ghl"):
        f = crm["flujo"][k]
        f.update(opps_mes=round(f["opps"] / f["meses"], 1), pct_ganadas=round(f["ganadas"] / f["opps"], 4),
                 pct_abiertas=round(f["abiertas"] / f["opps"], 4))

    tq3_26, tq3_25 = tot(q3_26), tot(q3_25)
    capa = dict(relacion=tot(rel(q3_26)), marketing=tot(mkt(q3_26)), por_canal=agrupa(q3_26, "canal"))
    metrics = {
        "q3_26": dict(**tq3_26, ticket=round(tq3_26["monto"] / tq3_26["ventas"], 2), periodo="jul–sep 2026"),
        "q3_25": dict(**tq3_25, ticket=round(tq3_25["monto"] / tq3_25["ventas"], 2), periodo="jul–sep 2025",
                      nota="incluye venta de sep-2025 con pérdida parcial de $6,800"),
        "yoy_monto": round(tq3_26["monto"] / tq3_25["monto"] - 1, 4),
        "yoy_ventas": round(tq3_26["ventas"] / tq3_25["ventas"] - 1, 4),
        "multiplo_monto": round(tq3_26["monto"] / tq3_25["monto"], 2),
        "rel_q3_26": dict(**tot(rel(q3_26)), pct_ventas=round(len(rel(q3_26)) / len(q3_26), 4),
                          pct_monto=round(tot(rel(q3_26))["monto"] / tq3_26["monto"], 4)),
        "rel_q3_25": dict(**tot(rel(q3_25)), pct_ventas=round(len(rel(q3_25)) / len(q3_25), 4),
                          pct_monto=round(tot(rel(q3_25))["monto"] / tq3_25["monto"], 4)),
        "y2025": dict(**tot(y25), perdida=round(sum(r["perdida"] for r in y25), 2), periodo="jun–dic 2025"),
        "y2026": dict(**tot(y26), ticket=round(tot(y26)["monto"] / len(y26), 2), periodo="ene–sep 2026",
                      prom_mensual=round(tot(y26)["monto"] / 9, 2), ventas_mes=round(len(y26) / 9, 1)),
        "total": dict(**tot(regs), periodo="jun-2025 – sep-2026"),
        "clientes_unicos": len(cli), "clientes_recurrentes": sum(c["compras"] >= 2 for c in cli),
        "clientes_fieles": sum(c["compras"] >= 3 for c in cli),
        "estados": len({s["edo"] for r in regs for s in r["sitios"] if s["edo"]}),
        "ventas_mes_q3": round(len(q3_26) / 3, 1),
        "ventas_mes_max_2026": max(p["ventas"] for p in s26),
        "ventas_mes_h1_26": round(len(sel(regs, 2026, range(1, 7))) / 6, 1),
        "valor_cliente": dict(recurrente=round(statistics.mean(c["monto"] for c in cli if c["compras"] >= 2), 2),
                              una_compra=round(statistics.mean(c["monto"] for c in cli if c["compras"] == 1), 2)),
    }
    metrics["valor_cliente"]["multiplo"] = round(metrics["valor_cliente"]["recurrente"] / metrics["valor_cliente"]["una_compra"], 2)
    trimestres = []
    for y, q in [(2025, 3), (2025, 4), (2026, 1), (2026, 2), (2026, 3)]:
        rs = sel(regs, y, range(3 * q - 2, 3 * q + 1))
        trimestres.append(dict(label=f"Q{q}-{str(y)[2:]}", **tot(rs)))
    metrics["trimestres"] = trimestres
    rel_t = [r for r in q3_26 if r["tipo"] in RELACION]
    metrics["rel_q3_26_por_tipo"] = dict(**tot(rel_t), pct_ventas=round(len(rel_t) / len(q3_26), 4))
    metrics["recomendados_nuevos_q3"] = tot([r for r in q3_26 if r["canal"] == "Recomendado" and r["tipo"] == "Nuevo"])

    public_rows = [{k: r[k] for k in ("id", "year", "month", "fecha_cierre", "dias", "ubicacion", "servicio", "canal",
                                      "tipo", "costo", "cliente", "nota")} |
                   {"estado": " / ".join(dict.fromkeys(s["edo"] for s in r["sitios"] if s["edo"])) or "—",
                    "municipio": " / ".join(s["mun"] for s in r["sitios"] if s["mun"]) or "—"} for r in regs]

    D = dict(
        generado=dt.datetime.now().isoformat(timespec="seconds"), corte=CORTE,
        fuentes=dict(ventas="Excel ANALISIS DE DATOS 2.0 (hojas 2025: jun–dic; 2026: ene–sep)",
                     ghl=f'API Go High Level (extracción {gs_corte}); Google Ads: {mm["fuente"]} · corte {mm["corte"]}',
                     clientify=base_cf["fuente"], forecast=fpub["fuente"]),
        metrics=metrics, serie=serie, proyeccion=proy, forecast=forecast_q4(s26),
        capas_q3=capa,
        clientes=dict(top=cli[:20], cross=sorted(cross, key=lambda c: -c["monto_2026"]),
                      fieles=[c for c in cli if c["compras"] >= 3], top_q3=cli_q3[:10]),
        servicios=dict(total=agrupa(regs, "servicio"), q3_26=agrupa(q3_26, "servicio"), q3_25=agrupa(q3_25, "servicio")),
        canales=dict(total=agrupa(regs, "canal"), q3_26=agrupa(q3_26, "canal"), y2026=agrupa(y26, "canal")),
        geo=dict(estados=por_estado(regs), municipios=por_municipio(regs, coords),
                 fuera_mapa=dict(**tot([r for r in regs if not any(s["mun"] for s in r["sitios"])]),
                                 detalle="Laboratorio Central, Gabinete, ubicaciones sin municipio (EDOMEX, CDMX, Guerrero) y localidades por confirmar")),
        ciclo=ciclo_por_servicio(regs),
        # Público: solo lo que usa la presentación. El tablero completo (etapas, pérdidas, ingresos CRM) va a INTERNO.
        ghl=dict(mm={k: mm[k] for k in ("fuente", "corte", "estatus", "demanda", "google_ads")}, api=api, calc=crm),
        clientify=base_cf,
        equipo=dict(asistente=dict(nombre="Daniela Torres", puesto="Asistente comercial", ingreso="2026-08-17",
                                   funciones=["Seguimiento", "Cotizaciones", "Agenda"])),
        rows=public_rows,
    )
    I = dict(
        cobranza=cob, corte_anterior=previo, ghl_mm=mm, ghl=gs,
        snapshot_anterior=json.loads((ROOT / "snapshots" / "pipeline_2026-07.json").read_text()),
        pendientes=[dict(id=r["id"], cliente=r["cliente"], costo=r["costo"], cobrado=r["cobrado"], restante=r["restante"],
                         pago=r["pago"], fecha_cierre=r["fecha_cierre"],
                         dias=(dt.date.fromisoformat(CORTE) - dt.date.fromisoformat(r["fecha_cierre"])).days if r["fecha_cierre"] else None)
                    for r in y26 if r["restante"] > 0],
        perdidas_2025=perdidas_2025,
        normalizacion_log=[dict(id=a, regla=b, original=c, limpio=d) for a, b, c, d in log],
        servicios_map=[dict(original=k, limpio=v, filas=sum(r["servicio_orig"] == k for r in regs)) for k, v in SERVICIO.items()],
        ubicaciones_map=sorted({(r["ubicacion"], " / ".join(f'{s["mun"] or "—"}, {s["edo"]}' for s in r["sitios"]) or "Fuera del mapa")
                                for r in regs}),
        subtotales_excel={f"{y}-{m:02d}": v for (y, m), v in sorted(subtot.items())},
        rows=regs,
    )
    I["ubicaciones_map"] = [dict(original=a, limpio=b) for a, b in I["ubicaciones_map"]]

    (DATA / "data.js").write_text("// Generado por data/build.py — no editar a mano\nwindow.DATA = "
                                  + json.dumps(D, ensure_ascii=False, separators=(",", ":")) + ";\n")
    (ROOT / "reporte").mkdir(exist_ok=True)
    (ROOT / "reporte" / "data_interno.js").write_text("// Generado por data/build.py — INTERNO, no publicar\nwindow.INTERNO = "
                                                      + json.dumps(I, ensure_ascii=False, separators=(",", ":")) + ";\n")

    # El snapshot de pipeline lo escribe ghl/summarize.py --write (fuente: API GHL)

    # Auditoría en consola
    print(f"Filas Excel: {len(filas)} → ventas normalizadas: {len(regs)}")
    for k in ("q3_26", "q3_25", "y2025", "y2026", "total"):
        print(f"  {k:7} {metrics[k]['ventas']:>4} ventas  ${metrics[k]['monto']:>14,.2f}")
    print(f"  Q3 vs optimista {proy['dif_opt']:+,.2f} ({proy['pct_opt']:+.1%}) · vs Q3-25 {metrics['yoy_monto']:+.1%} monto, {metrics['yoy_ventas']:+.1%} ventas")
    print(f"  Relación Q3-26 {metrics['rel_q3_26']['ventas']}/{tq3_26['ventas']} ({metrics['rel_q3_26']['pct_ventas']:.1%}) · ${metrics['rel_q3_26']['monto']:,.2f} ({metrics['rel_q3_26']['pct_monto']:.1%})")
    print(f"  Cobranza 2026 {cob['pct_cobrado']:.1%} · pendiente ${cob['pendiente']:,.2f}")
    print(f"  Q4 forecast {D['forecast']['q4']}")
    for (y, m), (n, mo) in sorted(subtot.items()):
        t2 = tot(sel(regs, y, [m]))
        flag = "" if (n, round(mo, 2)) == (t2["ventas"], t2["monto"]) else f"  ← Excel dice {n} / ${mo:,.2f}"
        print(f"  {y}-{m:02d}: {t2['ventas']:>3} / ${t2['monto']:>13,.2f}{flag}")


if __name__ == "__main__":
    main()
