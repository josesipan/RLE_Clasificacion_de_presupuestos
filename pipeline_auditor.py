"""
pipeline_auditor.py
====================
Script único para:

  1) LEER un presupuesto de Excel nuevo (detecta automáticamente las
     columnas de descripción, metrado, P.U., parcial, en soles y en
     dólares -- no importa si el layout cambia de proyecto a proyecto).
  2) CLASIFICAR cada partida según el catálogo KIMSAV (classify_rules.py).
  3) CUADRAR la suma (S/ + USD×TC) contra el total de la hoja RESUMEN del
     mismo archivo, para confirmar que no se perdió ni se duplicó nada.
  4) ITERAR: comparar los ratios S//m² resultantes contra el histórico de
     proyectos ya cargados, detectar valores fuera de rango (outliers),
     volver a probar la descripción de esas partidas contra TODAS las
     especialidades (no solo la asignada), reclasificar si hay mejor
     encaje, recalcular, y repetir -- reportando en cada vuelta el R² de
     la regresión Ratio vs. Área Techada como medida objetiva de qué tan
     "limpia" quedó la clasificación.

Requiere: openpyxl, numpy  (pip install openpyxl numpy)
Usa:      classify_rules.py (debe estar en la misma carpeta)

------------------------------------------------------------------------
USO BÁSICO
------------------------------------------------------------------------
from pipeline_auditor import procesar_presupuesto, cargar_historico

# 1. Cargar el histórico ya consolidado (tu BD_Historico_presupuestos.xlsx)
historico = cargar_historico("BD_Historico_presupuestos.xlsx")

# 2. Procesar un presupuesto nuevo
resultado = procesar_presupuesto(
    "Presupuesto_Nuevo_Proyecto.xlsx",
    nombre_proyecto="Proyecto Nuevo",
    area_techada=12345.0,          # si lo sabes; si no, lo intenta detectar
    historico=historico,
    max_iter=6,
)

# 3. Ver el reporte
resultado.imprimir_resumen()

# 4. Si todo cuadra y el R² final te convence, exportar las partidas ya
#    clasificadas y auditadas para pegarlas en BD_HISTORICO PPTO:
resultado.items_df.to_excel("partidas_para_agregar.xlsx", index=False)

------------------------------------------------------------------------
También se puede correr directo desde la terminal:

    python3 pipeline_auditor.py Presupuesto_Nuevo.xlsx --historico BD_Historico_presupuestos.xlsx --area 12345

------------------------------------------------------------------------
"""

import re
import sys
import argparse
import datetime
import statistics
from dataclasses import dataclass, field
from collections import defaultdict, Counter

import openpyxl
import numpy as np

from classify_rules import score_all, best_match, NON_ITEM_SPECIALTIES, U, SUBCODES


# ===========================================================================
# PARTE 1: LECTURA GENÉRICA DE UN PRESUPUESTO EXCEL
# ===========================================================================

HEADER_KEYWORDS = {
    # ORDEN IMPORTA: los campos más específicos van primero para que no
    # se los "robe" un patrón más genérico (p.ej. "codigo" debe resolverse
    # antes de que nada confunda la columna ITEM con la de texto).
    "codigo":      [r"^[IÍ]TEM$", r"^C[OÓ]DIGO$", r"^N[°º]?$"],
    "descripcion": [r"DESCRIPCI[OÓ]N", r"\bPARTIDA\b"],
    "unidad":      [r"\bUND\b", r"UNIDAD", r"\bUM\b"],
    "metrado":     [r"METRADO", r"CANTIDAD", r"\bCANT\b"],
    # nota: SIN \b final -- "S/." o "S/)" no cumplen un límite de palabra
    # porque "/" y "." son ambos no-alfanuméricos, así que \b nunca se
    # cumplía ahí y la columna quedaba sin detectar.
    # ojo: "P.U. TOTAL (PEN)" es precio UNITARIO, no el parcial de la
    # partida -- por eso parcial_soles exige la palabra "PARCIAL" y NO
    # acepta un "TOTAL" suelto aquí (eso se maneja aparte, como último
    # recurso, en el fallback de find_header_row para hojas sin ninguna
    # columna que diga "PARCIAL").
    "pu_soles":    [r"P\.?\s*U\.?.*(S/|PEN)", r"PRECIO\s*UNIT.*(S/|PEN)"],
    "parcial_soles": [r"PARCIAL.*(S/|PEN)", r"PRECIO\s*PARCIAL.*(S/|PEN)"],
    "pu_usd":      [r"P\.?\s*U\.?.*(US\$|USD|D[OÓ]LAR|\$)"],
    "parcial_usd": [r"PARCIAL.*(US\$|USD|D[OÓ]LAR|\$)", r"TOTAL.*(US\$|USD|\$)"],
}


def _match_header(text, patterns):
    if not isinstance(text, str):
        return False
    t = text.upper().strip()
    return any(re.search(p, t) for p in patterns)


def find_header_row(ws, max_scan_rows=60):
    """Busca la fila que más parece un encabezado de tabla de presupuesto."""
    best_row, best_hits, best_map = None, 0, {}
    for r in range(1, min(max_scan_rows, ws.max_row) + 1):
        col_map = {}
        used_cols = set()
        hits = 0
        for c in range(1, ws.max_column + 1):
            val = ws.cell(row=r, column=c).value
            if val is None or c in used_cols:
                continue
            for field_name, patterns in HEADER_KEYWORDS.items():
                if field_name in col_map:
                    continue
                if _match_header(val, patterns):
                    col_map[field_name] = c
                    used_cols.add(c)
                    hits += 1
                    break
        # Respaldo: algunas hojas (tablas chicas de equipamiento, "OBRA
        # CIVIL ASCENSOR", SEG, GG) no rotulan la moneda ("PARCIAL" o
        # "TOTAL" a secas, sin "S/" ni "PEN") porque ahí todo es soles y
        # no hay columna en USD. Si no se encontró parcial_soles pero sí
        # hay una columna llamada exactamente "PARCIAL" (o si no, "TOTAL")
        # sin usar, se toma como el parcial en soles.
        if "descripcion" in col_map and "parcial_soles" not in col_map and "parcial_usd" not in col_map:
            for candidato in ("PARCIAL", "TOTAL"):
                for c in range(1, ws.max_column + 1):
                    if c in used_cols:
                        continue
                    val = ws.cell(row=r, column=c).value
                    if isinstance(val, str) and val.strip().upper() == candidato:
                        col_map["parcial_soles"] = c
                        used_cols.add(c)
                        hits += 1
                        break
                if "parcial_soles" in col_map:
                    break
        # un encabezado real de presupuesto SIEMPRE trae al menos
        # descripción + metrado + algún parcial
        if "descripcion" in col_map and "metrado" in col_map and hits > best_hits:
            best_row, best_hits, best_map = r, hits, col_map
    return best_row, best_map


def find_tc(ws_or_wb):
    """Busca una celda tipo 'TC' / 'Tipo de Cambio' y devuelve su valor numérico."""
    sheets = ws_or_wb.worksheets if hasattr(ws_or_wb, "worksheets") else [ws_or_wb]
    for ws in sheets:
        for r in range(1, min(ws.max_row, 400) + 1):
            for c in range(1, min(ws.max_column, 30) + 1):
                val = ws.cell(row=r, column=c).value
                if isinstance(val, str) and re.search(r"TIPO\s*DE\s*CAMBIO|^\s*T\.?C\.?\s*$", val.upper()):
                    # el valor numérico suele estar en la celda vecina
                    for dc in (1, 2, -1):
                        v2 = ws.cell(row=r, column=c + dc).value
                        if isinstance(v2, (int, float)) and 2.5 < v2 < 5.0:
                            return float(v2)
    return 3.7  # fallback razonable (TC histórico usado en los otros proyectos)


def find_resumen_total(wb):
    """Busca en todas las hojas una fila tipo 'TOTAL GENERAL' / 'TOTAL PRESUPUESTO'
    (o una celda que diga exactamente 'TOTAL') y devuelve el valor numérico
    más a la derecha de esa fila (en las hojas tipo RESUMEN de KIMSAV esa es
    la columna 'TOTAL (S/)', que ya incluye el USD convertido)."""
    candidatos = []
    patrones_total = re.compile(
        r"TOTAL\s*(GENERAL|PRESUPUESTO|DE\s*OBRA|DEL\s*PROYECTO)|COSTO\s*TOTAL|PRESUPUESTO\s*TOTAL", re.I
    )
    for ws in wb.worksheets:
        for r in range(1, min(ws.max_row, 400) + 1):
            row_vals = [ws.cell(row=r, column=c).value for c in range(1, min(ws.max_column, 30) + 1)]
            textos = [str(v).strip() for v in row_vals if isinstance(v, str)]
            texto = " ".join(textos)
            es_total_exacto = any(t.upper() == "TOTAL" for t in textos)
            if patrones_total.search(texto) or es_total_exacto:
                numeros = [v for v in row_vals if isinstance(v, (int, float)) and v > 1000]
                if numeros:
                    # el total consolidado (S/ + USD×TC) suele ser el mayor
                    # valor numérico de la fila -- columnas más a la derecha
                    # pueden tener ratios/otros valores menores que no son el total.
                    peso = 2 if es_total_exacto else 1
                    candidatos.append((ws.title, r, max(numeros), peso))
    # ordenar para que los matches exactos de "TOTAL" aparezcan primero
    candidatos.sort(key=lambda c: -c[3])
    return [(t, r, v) for (t, r, v, _p) in candidatos]


def find_costo_directo(wb):
    """Busca la fila 'COSTO DIRECTO' del RESUMEN. Es la referencia correcta
    para cuadrar contra las partidas de las hojas de detalle, porque GG/
    Utilidad/IGV casi nunca están desglosados en ninguna hoja -- son
    porcentajes aplicados sobre el Costo Directo directamente en el
    RESUMEN (misma convención que se usó para los otros proyectos de la
    base: GG/Utilidad/IGV se agregan como una fila global aparte, no como
    partidas clasificadas)."""
    patron = re.compile(r"COSTO\s*DIRECTO", re.I)
    for ws in wb.worksheets:
        for r in range(1, min(ws.max_row, 400) + 1):
            for c in range(1, min(ws.max_column, 10) + 1):
                val = ws.cell(row=r, column=c).value
                if isinstance(val, str) and patron.search(val):
                    row_vals = [ws.cell(row=r, column=cc).value for cc in range(1, min(ws.max_column, 30) + 1)]
                    numeros = [v for v in row_vals if isinstance(v, (int, float)) and v > 1000]
                    if numeros:
                        return (ws.title, r, max(numeros))
    return None


def find_area_techada(wb):
    """Busca una celda tipo 'Área techada total' y devuelve su valor numérico."""
    patron = re.compile(r"[ÁA]REA\s*TECHADA\s*TOTAL", re.I)
    for ws in wb.worksheets:
        for r in range(1, min(ws.max_row, 400) + 1):
            for c in range(1, min(ws.max_column, 30) + 1):
                val = ws.cell(row=r, column=c).value
                if isinstance(val, str) and patron.search(val):
                    for dc in (1, 2, -1):
                        v2 = ws.cell(row=r, column=c + dc).value
                        if isinstance(v2, (int, float)) and v2 > 100:
                            return float(v2)
    return None


def extract_items(path, column_overrides=None, hojas_excluir=None, reporte_hojas=None):
    """Recorre TODAS las hojas del archivo (menos la de resumen) y extrae
    las filas que parezcan partidas de presupuesto reales.

    column_overrides: {"NOMBRE_HOJA": {"descripcion": 2, "metrado": 4, ...}}
        -- para las hojas donde la detección automática de encabezado
        falle (layouts raros, múltiples tablas, etc.), se puede forzar el
        mapeo de columnas manualmente en lugar de escribir un script
        completo desde cero como antes.
    hojas_excluir: lista adicional de nombres de hoja a saltar.
    reporte_hojas: si se pasa una lista, se le agregan (hoja, header_row,
        col_map, n_items) por cada hoja procesada -- útil para depurar
        rápido qué hoja no se está leyendo bien.
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    tc = find_tc(wb)
    items = []
    column_overrides = column_overrides or {}
    # La hoja de Gastos Generales se excluye por defecto: GG (y Utilidad e
    # IGV) no son partidas reales, son porcentajes que el RESUMEN aplica
    # sobre el Costo Directo -- se agregan aparte como una fila global
    # (misma convención usada en el resto de la base histórica), no como
    # ítems clasificados uno por uno.
    excluir = re.compile(r"RESUMEN|CARATULA|PORTADA|INDICE|^GG$|GASTOS\s*GENERALES" +
                          (("|" + "|".join(hojas_excluir)) if hojas_excluir else ""))

    for ws in wb.worksheets:
        if excluir.search(ws.title.upper()):
            continue
        if ws.title in column_overrides:
            header_row = column_overrides[ws.title].get("_header_row")
            col_map = {k: v for k, v in column_overrides[ws.title].items() if k != "_header_row"}
            if header_row is None:
                header_row, _ = find_header_row(ws)
        else:
            header_row, col_map = find_header_row(ws)
        if reporte_hojas is not None:
            n_antes = len(items)
        if header_row is None:
            if reporte_hojas is not None:
                reporte_hojas.append((ws.title, None, {}, 0))
            continue

        if "descripcion" not in col_map:
            if reporte_hojas is not None:
                reporte_hojas.append((ws.title, header_row, col_map, 0))
            continue

        last_codigo = None
        for r in range(header_row + 1, ws.max_row + 1):
            desc = ws.cell(row=r, column=col_map["descripcion"]).value
            metrado = ws.cell(row=r, column=col_map.get("metrado", 0)).value if "metrado" in col_map else None
            parcial_s = ws.cell(row=r, column=col_map.get("parcial_soles", 0)).value if "parcial_soles" in col_map else None
            parcial_d = ws.cell(row=r, column=col_map.get("parcial_usd", 0)).value if "parcial_usd" in col_map else None

            # fila vacía -> probablemente terminó la tabla de esta sección,
            # pero no cortamos de inmediato (puede haber subtítulos en medio)
            if desc is None and metrado is None and parcial_s is None and parcial_d is None:
                continue
            if not isinstance(desc, str) or len(desc.strip()) < 3:
                continue
            # filtra filas de subtítulo/sección (sin metrado numérico y sin parcial)
            metrado_ok = isinstance(metrado, (int, float)) and metrado != 0
            parcial_ok = isinstance(parcial_s, (int, float)) and parcial_s != 0
            parcial_usd_ok = isinstance(parcial_d, (int, float)) and parcial_d != 0
            if not (parcial_ok or parcial_usd_ok):
                continue
            # IMPORTANTE: varias hojas (IISS, IIEE, ...) traen una fila de
            # "subtotal de capítulo" (ej. '01  INSTALACIONES SANITARIAS')
            # con el monto YA sumado de todas las partidas de abajo, pero
            # sin metrado. Si se cuenta esa fila Y las partidas reales que
            # vienen después, se duplica el total (visto en Pardo 669: IISS
            # e IIEE salieron exactamente el doble). Por eso, cuando la hoja
            # sí tiene columna de metrado, se exige que la fila tenga un
            # metrado numérico real para contarla como partida.
            if "metrado" in col_map and not metrado_ok:
                continue

            codigo = ws.cell(row=r, column=col_map.get("codigo", 0)).value if "codigo" in col_map else None
            if codigo:
                last_codigo = codigo

            items.append({
                "hoja_origen": ws.title,
                "fila_origen": r,
                "codigo_origen": last_codigo,
                "descripcion": desc.strip(),
                "unidad": ws.cell(row=r, column=col_map.get("unidad", 0)).value if "unidad" in col_map else None,
                "metrado": metrado if metrado_ok else 1,
                "parcial_soles": parcial_s if parcial_ok else 0,
                "parcial_usd": parcial_d if parcial_usd_ok else 0,
                "tc": tc,
            })

        if reporte_hojas is not None:
            reporte_hojas.append((ws.title, header_row, col_map, len(items) - n_antes))

    items = _eliminar_hoja_maestra_duplicada(items, reporte_hojas=reporte_hojas)
    return wb, items


def _eliminar_hoja_maestra_duplicada(items, reporte_hojas=None, tolerancia=0.005):
    """Algunos presupuestos (visto en Éccolo) tienen una hoja "maestra"
    tipo "PPTO" que consolida en una sola hoja TODAS las partidas que YA
    están desglosadas, idénticas, en las hojas por especialidad (ESTRUCTURAS,
    ARQUITECTURA, IIEE, IISS, ACI, IIGG, ...). Si se cuentan ambas, el total
    del presupuesto sale exactamente el doble.

    No se detecta por nombre de hoja (eso sería específico de este
    proyecto) sino por el hecho real y generalizable que delata la
    duplicación: el monto total de esa hoja coincide casi exacto (±0.5%)
    con la SUMA de el resto de hojas combinadas. Cuando eso pasa, se
    descarta la hoja "maestra" (la de mayor cantidad de partidas de las
    dos) y se deja la versión ya desglosada por especialidad, que es la
    que permite clasificar bien.
    """
    por_hoja = {}
    for it in items:
        h = it.get("hoja_origen")
        por_hoja.setdefault(h, []).append(it)
    if len(por_hoja) < 2:
        return items

    montos = {h: sum((it["parcial_soles"] or 0) + (it["parcial_usd"] or 0) * it["tc"] for it in lst)
              for h, lst in por_hoja.items()}
    total_general = sum(montos.values())
    if total_general == 0:
        return items

    for h, monto_h in montos.items():
        resto = total_general - monto_h
        if monto_h == 0:
            continue
        diff_rel = abs(monto_h - resto) / monto_h
        if diff_rel <= tolerancia and len(por_hoja[h]) >= max(len(v) for k, v in por_hoja.items() if k != h):
            if reporte_hojas is not None:
                reporte_hojas.append((h, "DESCARTADA (duplica al resto de hojas combinadas)", {}, 0))
            items = [it for it in items if it.get("hoja_origen") != h]
            break  # solo se espera una hoja maestra; evita falsos positivos en cascada
    return items


# ===========================================================================
# PARTE 2: CLASIFICACIÓN
# ===========================================================================

def clasificar_items(items, historico=None, min_similitud_hist=0.45, min_consenso_hist=0.60):
    """Clasifica cada partida.
    Si se proporciona `historico` con proyectos previos:
      1. Intenta clasificar por similitud con histórico.
      2. Si no hay suficiente similitud/evidencia, usa reglas KIMSAV (classify_rules.py).
    Si NO hay histórico (o no hubo coincidencia):
      1. Usa reglas KIMSAV (classify_rules.py).
      2. Usa información auxiliar de hoja de origen (fallback).
      3. Hereda para partidas de desglose continuo.
    """
    corpus = construir_corpus_similitud(historico) if historico else []

    for it in items:
        it["subtotal_soles"] = (it["parcial_soles"] or 0) + (it["parcial_usd"] or 0) * it["tc"]
        it["ref_proyecto"] = None
        it["ref_descripcion"] = None
        it["observaciones"] = None

        clasificado_hist = False
        if corpus:
            esp_h, sub_h, cod_h, conf_h, ref_proy, ref_desc = clasificar_por_similitud(
                it["descripcion"], corpus, min_similitud=min_similitud_hist,
                min_consenso=min_consenso_hist, min_palabras_comunes=2, min_vecinos=1
            )
            if esp_h:
                it["especialidad"] = esp_h
                it["subespecialidad"] = sub_h
                it["codigo"] = cod_h
                it["score_clasificacion"] = round(conf_h * 10, 2)
                it["metodo_clasificacion"] = "historico"
                it["confianza_similitud"] = round(conf_h, 3)
                it["ref_proyecto"] = ref_proy
                it["ref_descripcion"] = ref_desc
                clasificado_hist = True

        if not clasificado_hist:
            esp, sub, codigo, score = best_match(it["descripcion"])
            it["especialidad"] = esp or "SIN_CLASIFICAR"
            it["subespecialidad"] = sub or "Revisar manualmente"
            it["codigo"] = codigo
            it["score_clasificacion"] = score
            it["metodo_clasificacion"] = "reglas" if esp else "sin_resolver"

    # Fallback auxiliar por hoja de origen para los que queden sin clasificar
    aplicar_fallback_por_hoja(items)
    _heredar_clasificacion_de_item_anterior(items)

    for it in items:
        if it["especialidad"] == "SIN_CLASIFICAR" or it["subespecialidad"] == "Revisar manualmente" or not it.get("codigo"):
            if not it.get("observaciones"):
                it["observaciones"] = "Revisar manualmente"

    return items


# patrón de una partida que NO tiene descripción propia -- solo lista rangos
# de departamentos/ambientes a los que se les aplica la partida DE ARRIBA
# (ej. "DPTO: (310 - 2010), (2109 - 2309)", "(304 - 2704),(305 - 2005)").
# Es una convención real de cómo arman el detalle por unidad en estos
# presupuestos: la descripción real está en la fila anterior, esta fila es
# solo el desglose/metrado por departamento de esa misma partida.
_PATRON_RANGO_DPTOS = re.compile(
    r"^\s*(DPTO:?\s*)?\(?\s*\d{2,4}\s*[-–]\s*\d{2,4}\s*\)?"
    r"(\s*[,]?\s*\(?\s*\d{2,4}(\s*[-–]\s*\d{2,4})?\)?)*\s*$"
)


def _heredar_clasificacion_de_item_anterior(items):
    """Para partidas sin descripción propia (solo rangos de departamentos),
    hereda especialidad/subespecialidad/código de la partida clasificada más
    cercana ANTES que ella en la misma hoja -- no inventa una clasificación
    nueva, usa la que ya se calculó para la partida real de la que es
    desglose."""
    por_hoja = defaultdict(list)
    for it in items:
        por_hoja[it.get("hoja_origen")].append(it)

    for hoja, lst in por_hoja.items():
        ultimo_valido = None
        for it in lst:
            si_es_rango = bool(_PATRON_RANGO_DPTOS.match((it.get("descripcion") or "").strip()))
            tiene_clasif_real = (it["subespecialidad"] != "Revisar manualmente"
                                  and it["especialidad"] != "SIN_CLASIFICAR")
            if tiene_clasif_real:
                ultimo_valido = it
            elif si_es_rango and ultimo_valido is not None:
                it["especialidad"] = ultimo_valido["especialidad"]
                it["subespecialidad"] = ultimo_valido["subespecialidad"]
                it["codigo"] = ultimo_valido["codigo"]
                it["metodo_clasificacion"] = "heredado_item_anterior"


# ---------------------------------------------------------------------------
# Clasificador de RESPALDO por HOJA DE ORIGEN: los presupuestos de este tipo
# de obra se arman casi siempre agrupando las partidas en una hoja de Excel
# por especialidad ("IISS", "IIEE", "ARQ TORRE", "EST SOTANO", "ACI", "GAS",
# "CCTV - DACI", "GG_E1", ...). Esta es una convención del RUBRO, no una
# particularidad de un proyecto, así que es un patrón legítimo para
# generalizar: cuando la descripción no da ninguna pista (las reglas de
# palabras clave no encuentran nada), el NOMBRE DE LA HOJA donde está la
# partida sigue siendo información real del propio presupuesto -- no es
# inventar un dato, es usar uno que ya estaba ahí.
#
# Esto solo decide la ESPECIALIDAD (con confianza media, método "hoja"),
# nunca una especialidad mixta/ambigua ("IIEE Y MECANICAS", "ACIySPC") --
# ahí se prefiere dejarlo para el clasificador de similitud en vez de
# adivinar. La subespecialidad exacta solo se asigna cuando la especialidad
# tiene una sola subespecialidad posible en el catálogo (IIEE, IIMM, GAS,
# ASCENSORES, GG); si tiene varias (ARQ, ESTT, OP, IISS, CCDÉBILES, EQUIP)
# se deja "Revisar manualmente" -- ya es una ganancia real tener bien la
# especialidad (y por lo tanto el costo S/m² de ese rubro), aunque la
# subespecialidad fina quede pendiente de revisión.
# ---------------------------------------------------------------------------

_HOJA_PATTERNS = [
    (r"^GG[_\s]|GASTOS\s*GENERALES", "GG"),
    # "IIGG" (distinto de "GG" -- de hecho en los 2 proyectos donde aparece,
    # Atria 66 y Éccolo, el proyecto tiene AMBAS hojas "GG" e "IIGG" por
    # separado) significa "Instalaciones Internas de GAS", no Gastos
    # Generales ni Sanitarias: comprobado sumando el total de la hoja IIGG
    # contra la línea "SISTEMA DE GAS" del RESUMEN en ambos proyectos --
    # cuadra exacto en los dos (118,454.46 en Atria 66 y 165,192 en Éccolo).
    (r"^IIGG\b", "GAS"),
    (r"^ARQ(UITECTURA)?\b|PAISAJ", "ARQ"),
    (r"^EST(R|RUCTURAS?)?\b|^ESTR_", "ESTT"),
    (r"^MEJ\s*SUELO", "ESTT"),
    (r"^IISS\b|SANITARI", "IISS"),
    (r"^IIEE\b(?!.*MEC)", "IIEE"),            # "IIEE Y MECANICAS" se excluye (mixta)
    (r"^IIMM\b|^MEC[AÁ]NICAS?\b", "IIMM"),  # "AC" se probó: en la práctica no
                                             # siempre es Aire Acondicionado
                                             # (en un proyecto eran Acabados) --
                                             # muy ambiguo, no se usa.
    (r"^GAS\b", "GAS"),
    (r"^ACI\b(?!.*SPC)|CTO\s*BOMBAS[-\s]*ACI\b(?!.*SPC)", "ACI"),  # "ACIySPC" se excluye (mixta)
    (r"CCD[EÉ]BILES|CCTV|^DACI\b", "CCDÉBILES"),
    (r"^OP\b|^OOPP\b|^OPROV|^SEG\b|^VECINOS\b", "OP"),
    (r"^ASCENSOR", "ASCENSORES"),
    (r"^EQUIP\b|^EQ$", "EQUIP"),
]


def _especialidad_por_hoja(hoja_origen):
    if not hoja_origen:
        return None
    h = U(hoja_origen)
    for pat, esp in _HOJA_PATTERNS:
        if re.search(pat, h):
            return esp
    return None


def aplicar_fallback_por_hoja(items):
    """Solo toca las partidas que las reglas de palabras clave no pudieron
    clasificar. No pisa nada que ya tenga especialidad asignada."""
    resueltos = 0
    for it in items:
        if it["especialidad"] != "SIN_CLASIFICAR":
            continue
        esp = _especialidad_por_hoja(it.get("hoja_origen"))
        if not esp:
            continue
        # algunas hojas "ACI" en la práctica mezclan bombas de Presión
        # Constante (SPC) junto con las de ACI -- si la descripción de ESTA
        # partida en particular menciona presión constante/SPC, no hay que
        # creerle ciegamente al nombre de la hoja.
        if esp == "ACI" and re.search(r"PRESI[OÓ]N\s*CONSTANTE|\bSPC\b", U(it["descripcion"])):
            continue
        subs_posibles = SUBCODES.get(esp, {})
        if len(subs_posibles) == 1:
            sub, codigo = next(iter(subs_posibles.items()))
        else:
            sub, codigo = "Revisar manualmente", None
        it["especialidad"] = esp
        it["subespecialidad"] = sub
        it["codigo"] = codigo
        it["metodo_clasificacion"] = "hoja"
        resueltos += 1
    return resueltos


# ---------------------------------------------------------------------------
# Clasificador de RESPALDO por similitud de texto, para lo que las reglas de
# palabras clave no reconocen. No se apoya en un solo "vecino parecido" de la
# base histórica (porque si ESE vecino está mal clasificado, el error se
# contagia -- el mismo tipo de problema que pasó con "control remoto"). En
# vez de eso, busca las K partidas más parecidas y solo acepta el resultado
# si hay CONSENSO fuerte entre ellas (por defecto, al menos 60% de acuerdo);
# si están divididas o no hay suficiente parecido, lo deja sin clasificar
# para revisión manual en vez de arriesgar una reclasificación mala.
# ---------------------------------------------------------------------------

_STOPWORDS = {
    "DE", "DEL", "LA", "EL", "LOS", "LAS", "EN", "Y", "A", "CON", "PARA",
    "POR", "SEGUN", "SEGÚN", "SIN", "AL", "UN", "UNA", "UNO", "SE", "SU",
    "O", "E", "QUE", "ES", "SOBRE", "ENTRE", "COMO",
}


def _tokens(desc):
    d = U(desc)
    d = re.sub(r"[^A-ZÁÉÍÓÚÑ0-9 ]", " ", d)
    return {w for w in d.split() if len(w) > 2 and w not in _STOPWORDS and not w.isdigit()}


def construir_corpus_similitud(historico):
    """Arma el corpus de referencia: (tokens, especialidad, subespecialidad,
    codigo, proyecto, descripcion) de cada partida YA clasificada (por reglas o ya confirmada) en
    el histórico. Excluye GG/Utilidad/IGV y cualquier cosa sin clasificar."""
    corpus = []
    for proy, data in (historico or {}).items():
        for it in data["items"]:
            esp = it.get("especialidad")
            if not esp or esp in NON_ITEM_SPECIALTIES or esp == "SIN_CLASIFICAR":
                continue
            toks = _tokens(it["descripcion"])
            if toks:
                corpus.append((toks, esp, it.get("subespecialidad"), it.get("codigo"), proy, it.get("descripcion", "")))
    return corpus


def clasificar_por_similitud(desc, corpus, k=7, min_similitud=0.40, min_consenso=0.70,
                              min_palabras_comunes=2, min_vecinos=1):
    """Devuelve (especialidad, subespecialidad, codigo, confianza, ref_proyecto, ref_descripcion) o
    (None, None, None, 0.0, None, None) si no hay consenso suficiente."""
    toks = _tokens(desc)
    if not toks or not corpus:
        return None, None, None, 0.0, None, None

    candidatos = []
    for item in corpus:
        c_toks = item[0]
        esp = item[1]
        sub = item[2]
        cod = item[3]
        proy = item[4] if len(item) > 4 else None
        c_desc = item[5] if len(item) > 5 else None

        inter = toks & c_toks
        if len(inter) < min_palabras_comunes:
            continue
        sim = len(inter) / len(toks | c_toks)
        if sim >= min_similitud:
            candidatos.append((sim, esp, sub, cod, proy, c_desc))

    if not candidatos:
        return None, None, None, 0.0, None, None

    candidatos.sort(key=lambda x: -x[0])
    top = candidatos[:k]

    if len(top) < min_vecinos:
        return None, None, None, 0.0, None, None

    votos = Counter(c[1] for c in top)
    especialidad, n_votos = votos.most_common(1)[0]
    consenso = n_votos / len(top)
    if consenso < min_consenso or n_votos < min_vecinos:
        return None, None, None, 0.0, None, None

    subs = Counter(c[2] for c in top if c[1] == especialidad)
    subespecialidad, _ = subs.most_common(1)[0]
    codigos = [c[3] for c in top if c[1] == especialidad and c[2] == subespecialidad]
    codigo = codigos[0] if codigos else None

    # Tomar la mejor coincidencia que vote por esa especialidad/subespecialidad
    mejores_cands = [c for c in top if c[1] == especialidad and c[2] == subespecialidad]
    ref_cand = mejores_cands[0] if mejores_cands else top[0]
    ref_proyecto = ref_cand[4]
    ref_desc = ref_cand[5]

    confianza = ref_cand[0] * consenso
    return especialidad, subespecialidad, codigo, confianza, ref_proyecto, ref_desc


def aplicar_fallback_similitud(items, historico, min_similitud=0.45, min_consenso=0.60,
                                min_palabras_comunes=2, min_vecinos=1):
    """Solo toca las partidas que quedaron 'SIN_CLASIFICAR' tras las reglas
    de palabras clave. Marca claramente el método usado, la confianza y la referencia."""
    corpus = construir_corpus_similitud(historico)
    resueltos, sin_resolver = 0, 0
    for it in items:
        if it["especialidad"] != "SIN_CLASIFICAR":
            continue
        esp, sub, cod, conf, ref_proy, ref_desc = clasificar_por_similitud(
            it["descripcion"], corpus, min_similitud=min_similitud, min_consenso=min_consenso,
            min_palabras_comunes=min_palabras_comunes, min_vecinos=min_vecinos)
        if esp:
            it["especialidad"] = esp
            it["subespecialidad"] = sub
            it["codigo"] = cod
            it["metodo_clasificacion"] = "historico"
            it["score_clasificacion"] = round(conf * 10, 2)
            it["confianza_similitud"] = round(conf, 3)
            it["ref_proyecto"] = ref_proy
            it["ref_descripcion"] = ref_desc
            resueltos += 1
        else:
            sin_resolver += 1
    return resueltos, sin_resolver


# ===========================================================================
# PARTE 3: CUADRE CONTRA RESUMEN
# ===========================================================================

def cuadrar_contra_resumen(items, wb, tolerancia_pct=0.5):
    suma_calculada = sum(it["subtotal_soles"] for it in items)

    reporte = {
        "suma_calculada": suma_calculada,
        "cuadra": None,
        "diferencia": None,
        "diferencia_pct": None,
    }

    # Referencia #1 (la correcta para comparar partidas): COSTO DIRECTO.
    # GG/Utilidad/IGV casi nunca están desglosados por partida -- son
    # porcentajes sobre el Costo Directo que se agregan como una fila
    # global aparte (misma convención usada en el resto de la base).
    costo_directo = find_costo_directo(wb)
    if costo_directo:
        diferencia = suma_calculada - costo_directo[2]
        diferencia_pct = (diferencia / costo_directo[2] * 100) if costo_directo[2] else None
        reporte["costo_directo"] = costo_directo
        reporte["diferencia"] = diferencia
        reporte["diferencia_pct"] = diferencia_pct
        reporte["cuadra"] = abs(diferencia_pct) <= tolerancia_pct if diferencia_pct is not None else "sin_referencia"

    # Referencia #2 (solo informativa): TOTAL general del presupuesto
    # (Costo Directo + GG + Utilidad + IGV), para que quede a la vista
    # cuánto habría que sumar aparte una vez agregadas esas 3 filas globales.
    candidatos_total = find_resumen_total(wb)
    if candidatos_total:
        mejor_total = max(candidatos_total, key=lambda c: c[2])
        reporte["total_general_resumen"] = mejor_total

    if reporte["cuadra"] is None:
        reporte["cuadra"] = "sin_referencia"
    return reporte


# ===========================================================================
# PARTE 4: AUDITORÍA ITERATIVA + CALIDAD (R²)
# ===========================================================================

def _mad(values, med):
    if not values:
        return 0
    devs = [abs(v - med) for v in values]
    m = statistics.median(devs)
    return m if m > 0 else (statistics.pstdev(values) or 1.0)


def ratios_por_rubro(todos_los_items_por_proyecto):
    """todos_los_items_por_proyecto: {proyecto: {"items": [...], "area": float}}"""
    sums = defaultdict(lambda: defaultdict(float))
    for proy, data in todos_los_items_por_proyecto.items():
        for it in data["items"]:
            if it["especialidad"] in NON_ITEM_SPECIALTIES:
                continue
            key = (it["especialidad"], it["subespecialidad"])
            sums[key][proy] += it["subtotal_soles"]
    out = {}
    for key, by_proy in sums.items():
        out[key] = {}
        for proy, total in by_proy.items():
            area = todos_los_items_por_proyecto[proy]["area"] or 0
            out[key][proy] = (total / area) if area else 0
    return out


def detectar_outliers(ratios, z_threshold=2.5):
    flags = []
    for (esp, sub), by_proy in ratios.items():
        valores = list(by_proy.values())
        if len(valores) < 3:
            continue
        med = statistics.median(valores)
        m = _mad(valores, med)
        if m == 0:
            continue
        for proy, val in by_proy.items():
            z = abs(val - med) / m
            if z >= z_threshold:
                flags.append({"especialidad": esp, "subespecialidad": sub,
                               "proyecto": proy, "ratio": val, "mediana": med, "z": z})
    flags.sort(key=lambda f: -f["z"])
    return flags


def reclasificar_outliers(todos_los_items_por_proyecto, flags, min_gap=6):
    flagged = {(f["especialidad"], f["subespecialidad"], f["proyecto"]) for f in flags}
    cambios = []
    for proy, data in todos_los_items_por_proyecto.items():
        for it in data["items"]:
            key = (it["especialidad"], it["subespecialidad"], proy)
            if key not in flagged or it["especialidad"] in NON_ITEM_SPECIALTIES:
                continue
            scores = score_all(it["descripcion"])
            if not scores:
                continue
            actual = scores.get(it["especialidad"], {}).get("score", 0)
            mejor_esp = max(scores, key=lambda k: scores[k]["score"])
            mejor_score = scores[mejor_esp]["score"]
            if mejor_esp != it["especialidad"] and (mejor_score - actual) >= min_gap:
                cambios.append({
                    "proyecto": proy, "descripcion": it["descripcion"],
                    "de": (it["especialidad"], it["subespecialidad"]),
                    "a": (mejor_esp, scores[mejor_esp]["subespecialidad"]),
                })
                it["especialidad"] = mejor_esp
                it["subespecialidad"] = scores[mejor_esp]["subespecialidad"]
                it["codigo"] = scores[mejor_esp]["codigo"]
    return cambios


def calcular_r2_area(todos_los_items_por_proyecto):
    """R² de Ratio General (TOTAL S// m²) vs Área Techada entre proyectos.
    Es el indicador objetivo de qué tan 'limpia'/predecible quedó la base
    después de cada vuelta de reclasificación."""
    areas, ratios = [], []
    for proy, data in todos_los_items_por_proyecto.items():
        area = data["area"]
        if not area:
            continue
        total = sum(it["subtotal_soles"] for it in data["items"])
        areas.append(area)
        ratios.append(total / area)
    if len(areas) < 3:
        return None
    areas = np.array(areas)
    ratios = np.array(ratios)
    if np.std(areas) == 0 or np.std(ratios) == 0:
        return None
    r = np.corrcoef(areas, ratios)[0, 1]
    return r ** 2


def iterar_auditoria(todos_los_items_por_proyecto, max_iter=6, z_threshold=2.5, min_gap=6):
    historial = []
    for i in range(1, max_iter + 1):
        ratios = ratios_por_rubro(todos_los_items_por_proyecto)
        flags = detectar_outliers(ratios, z_threshold)
        cambios = reclasificar_outliers(todos_los_items_por_proyecto, flags, min_gap)
        r2 = calcular_r2_area(todos_los_items_por_proyecto)
        historial.append({"iteracion": i, "outliers": len(flags), "cambios": len(cambios), "r2": r2})
        if not cambios:
            break
    return historial


# ===========================================================================
# RESULTADO / REPORTE
# ===========================================================================

@dataclass
class ResultadoAuditoria:
    proyecto: str
    items: list
    cuadre: dict
    historial_iteraciones: list
    r2_final: float = None
    resueltos_similitud: int = 0
    sin_resolver_similitud: int = 0
    area_techada: float = None
    proyectos_hist_antes: int = 0
    proyectos_hist_despues: int = 0

    def imprimir_resumen(self):
        print(f"\n{'='*70}\nAUDITORÍA Y REPORTE: {self.proyecto}\n{'='*70}")
        print(f"Cantidad total de partidas: {len(self.items)}")

        por_hist = sum(1 for it in self.items if it.get("metodo_clasificacion") == "historico")
        por_reglas = sum(1 for it in self.items if it.get("metodo_clasificacion") == "reglas")
        por_hoja = sum(1 for it in self.items if it.get("metodo_clasificacion") == "hoja")
        por_heredado = sum(1 for it in self.items if it.get("metodo_clasificacion") == "heredado_item_anterior")
        sin_resolver = sum(1 for it in self.items if it.get("metodo_clasificacion") == "sin_resolver" or it.get("especialidad") == "SIN_CLASIFICAR")
        para_revision = sum(1 for it in self.items if it.get("observaciones") == "Revisar manualmente" or it.get("subespecialidad") == "Revisar manualmente" or it.get("especialidad") == "SIN_CLASIFICAR")

        scores = [it.get("score_clasificacion") for it in self.items if isinstance(it.get("score_clasificacion"), (int, float))]
        score_prom = statistics.mean(scores) if scores else 0.0
        score_min = min(scores) if scores else 0.0

        print(f"\n--- DESGLOSE POR MÉTODO DE CLASIFICACIÓN ---")
        print(f"  • Clasificadas por histórico: {por_hist}")
        print(f"  • Clasificadas por reglas (KIMSAV): {por_reglas}")
        print(f"  • Clasificadas por hoja de origen (auxiliar): {por_hoja}")
        if por_heredado:
            print(f"  • Clasificadas por desglose de fila anterior: {por_heredado}")
        print(f"  • Sin resolver: {sin_resolver}")
        print(f"  • Cantidad para revisión manual: {para_revision}")
        print(f"  • Score de clasificación promedio: {score_prom:.2f}")
        print(f"  • Score de clasificación mínimo: {score_min:.2f}")

        print(f"\n--- DISTRIBUCIÓN POR ESPECIALIDAD ---")
        conteo_esp = Counter(it.get("especialidad", "SIN_CLASIFICAR") for it in self.items)
        monto_esp = defaultdict(float)
        for it in self.items:
            monto_esp[it.get("especialidad", "SIN_CLASIFICAR")] += it.get("subtotal_soles", 0)

        for esp, cnt in sorted(conteo_esp.items(), key=lambda x: -x[1]):
            monto = monto_esp[esp]
            print(f"  {esp:15s}: {cnt:4d} partidas  |  S/ {monto:>14,.2f}")

        print(f"\n--- CUADRE ECONÓMICO CONTRA RESUMEN ---")
        print(f"Suma calculada de partidas (S/ + USD×TC): {self.cuadre['suma_calculada']:,.2f}")
        if self.cuadre["cuadra"] == "sin_referencia":
            print("No se encontró 'COSTO DIRECTO' en el RESUMEN para comparar automáticamente.")
        else:
            cd = self.cuadre["costo_directo"]
            print(f"Costo Directo (RESUMEN): {cd[2]:,.2f}  (hoja '{cd[0]}', fila {cd[1]})")
            print(f"Diferencia: {self.cuadre['diferencia']:,.2f}  ({self.cuadre['diferencia_pct']:.3f}%)")
            print("Estado: CUADRA" if self.cuadre["cuadra"] else "Estado: NO CUADRA -- revisar")

        if self.cuadre.get("total_general_resumen"):
            tg = self.cuadre["total_general_resumen"]
            print(f"(informativo) TOTAL general del presupuesto (con GG+Utilidad+IGV): {tg[2]:,.2f}")

        print(f"\n--- ESTADO DE LA BASE HISTÓRICA ACUMULATIVA ---")
        print(f"Proyectos históricos disponibles ANTES de procesar: {self.proyectos_hist_antes}")
        print(f"Proyectos históricos disponibles DESPUÉS de procesar: {self.proyectos_hist_despues}")

        if self.historial_iteraciones:
            print(f"\n--- ITERACIONES DE AUDITORÍA (OUTLIERS / R²) ---")
            for h in self.historial_iteraciones:
                r2_txt = f"{h['r2']:.3f}" if h["r2"] is not None else "N/D"
                print(f"  Iter {h['iteracion']}: outliers={h['outliers']:3d}  "
                      f"reclasificaciones={h['cambios']:3d}  R² general vs área={r2_txt}")


_ENCABEZADOS_BD = {
    "proyecto": ["Nombre_Proyecto"],
    "especialidad": ["Especialidad"],
    "subespecialidad": ["Subespecialidad"],
    "codigo": ["Código", "Codigo"],
    "descripcion": ["Descripcion_Partida", "Descripción_Partida"],
    "unidad": ["Unidad"],
    "metrado": ["Metrado"],
    "pu_soles": ["PU (S/)", "PU (S//)"],
    "parcial_soles": ["Parcial (S/)", "Parcial (S//)"],
    "pu_usd": ["PU (USD)"],
    "parcial_usd": ["Parcial (USD)"],
    "subtotal_soles": ["Subtotal (S/)", "Subtotal (S//)"],
    "tc": ["TC"],
    "area": ["Área Techada", "Area Techada"],
    "ratio": ["Ratio (S/./m2)", "Ratio (S//m2)", "Ratio (S/m2)"],
    "origen_hoja": ["Origen_Hoja"],
    "origen_item": ["Origen_Item"],
    "fecha": ["Fecha"],
    "id_proyecto": ["ID_Proyecto"],
    "tipo": ["Tipo"],
    "departamento": ["Departamento"],
}


def _mapear_columnas_bd(ws, max_fila_busqueda=5):
    """Encuentra la fila de encabezados de BD_HISTORICO PPTO buscando la
    columna 'Especialidad' o 'Nombre_Proyecto' y arma el diccionario columna-por-nombre."""
    fila_header = None
    for r in range(1, max_fila_busqueda + 1):
        valores = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if any(v in ("Especialidad", "Nombre_Proyecto") for v in valores):
            fila_header = r
            break
    if fila_header is None:
        raise ValueError(
            "No se encontró la fila de encabezados en la hoja BD_HISTORICO PPTO -- "
            "revisa que la base tenga encabezados como 'Nombre_Proyecto' o 'Especialidad'."
        )
    valores = {c: ws.cell(row=fila_header, column=c).value for c in range(1, ws.max_column + 1)}
    COL = {}
    for campo, variantes in _ENCABEZADOS_BD.items():
        for c, v in valores.items():
            if v in variantes:
                COL[campo] = c
                break
    return fila_header, COL


def cargar_historico(path_bd_historico):
    """Carga BD_HISTORICO PPTO ya consolidada para servir de referencia de
    similitud y ratios."""
    wb = openpyxl.load_workbook(path_bd_historico, data_only=True)
    if "BD_HISTORICO PPTO" not in wb.sheetnames:
        return {}
    ws = wb["BD_HISTORICO PPTO"]
    fila_header, COL = _mapear_columnas_bd(ws)
    out = {}
    for r in range(fila_header + 1, ws.max_row + 1):
        proy = ws.cell(row=r, column=COL.get("proyecto", 1)).value
        esp = ws.cell(row=r, column=COL.get("especialidad", 4)).value
        if proy is None or str(proy).strip() == "" or esp is None:
            continue
        proy_str = str(proy).strip()
        out.setdefault(proy_str, {"items": [], "area": None})
        if "area" in COL:
            area = ws.cell(row=r, column=COL["area"]).value
            if area:
                out[proy_str]["area"] = area
        out[proy_str]["items"].append({
            "descripcion": ws.cell(row=r, column=COL.get("descripcion", 7)).value or "",
            "unidad": ws.cell(row=r, column=COL.get("unidad", 8)).value,
            "especialidad": esp,
            "subespecialidad": ws.cell(row=r, column=COL.get("subespecialidad", 5)).value,
            "codigo": ws.cell(row=r, column=COL.get("codigo", 6)).value,
            "subtotal_soles": ws.cell(row=r, column=COL.get("subtotal_soles", 14)).value
                               or ws.cell(row=r, column=COL.get("parcial_soles", 11)).value or 0,
        })
    return out


def procesar_presupuesto(path_excel, nombre_proyecto=None, area_techada=None,
                          historico=None, max_iter=6, z_threshold=2.5, min_gap=6,
                          min_similitud_hist=0.45, min_consenso_hist=0.60):
    """Procesa y clasifica un presupuesto de Excel, utilizando el histórico disponible."""
    if nombre_proyecto is None:
        base_filename = path_excel.replace("\\", "/").rsplit("/", 1)[-1]
        nombre_proyecto = base_filename.rsplit(".", 1)[0]

    wb, items = extract_items(path_excel)
    items = clasificar_items(items, historico=historico,
                             min_similitud_hist=min_similitud_hist,
                             min_consenso_hist=min_consenso_hist)

    cuadre = cuadrar_contra_resumen(items, wb)

    if area_techada is None:
        area_techada = find_area_techada(wb)

    # Construir el universo de proyectos para la auditoría de ratios
    universo = dict(historico) if historico else {}
    universo[nombre_proyecto] = {"items": items, "area": area_techada}

    historial = iterar_auditoria(universo, max_iter=max_iter,
                                  z_threshold=z_threshold, min_gap=min_gap)
    r2_final = historial[-1]["r2"] if historial else None

    proy_antes = len(historico) if historico else 0

    return ResultadoAuditoria(
        proyecto=nombre_proyecto,
        items=items,
        cuadre=cuadre,
        historial_iteraciones=historial,
        r2_final=r2_final,
        resueltos_similitud=sum(1 for it in items if it.get("metodo_clasificacion") == "historico"),
        sin_resolver_similitud=sum(1 for it in items if it.get("metodo_clasificacion") == "sin_resolver"),
        area_techada=area_techada,
        proyectos_hist_antes=proy_antes,
        proyectos_hist_despues=proy_antes,
    )


def tabla_ratios_proyecto(resultado):
    """Ratio S//m2 por especialidad y subespecialidad de UN SOLO proyecto ya
    procesado con procesar_presupuesto()."""
    if not resultado.area_techada:
        return []

    sums = defaultdict(float)
    for it in resultado.items:
        if it["especialidad"] in NON_ITEM_SPECIALTIES:
            continue
        key = (it["especialidad"], it["subespecialidad"])
        sums[key] += it["subtotal_soles"]

    filas = []
    for (esp, sub), monto in sorted(sums.items(), key=lambda kv: -kv[1]):
        filas.append({
            "Especialidad": esp,
            "Subespecialidad": sub,
            "Monto (S/)": round(monto, 2),
            "Ratio (S//m2)": round(monto / resultado.area_techada, 2),
        })
    return filas


def imprimir_tabla_ratios(resultado):
    filas = tabla_ratios_proyecto(resultado)
    if not filas:
        return
    print(f"\n--- RATIOS S//m2 -- {resultado.proyecto} (área techada: {resultado.area_techada:,.2f} m2) ---")
    print(f"{'Especialidad':14s} {'Subespecialidad':38s} {'Monto (S/)':>16s} {'Ratio (S//m2)':>14s}")
    for f in filas:
        print(f"{f['Especialidad']:14s} {f['Subespecialidad'][:38]:38s} {f['Monto (S/)']:>16,.2f} {f['Ratio (S//m2)']:>14,.2f}")


def exportar_excel_con_clasificacion(path_original, resultado, out_path):
    """Toma el Excel ORIGINAL del presupuesto y le agrega al costado las columnas KIMSAV."""
    wb = openpyxl.load_workbook(path_original)
    area = resultado.area_techada

    por_hoja = defaultdict(list)
    for it in resultado.items:
        por_hoja[it["hoja_origen"]].append(it)

    for hoja, items_hoja in por_hoja.items():
        if hoja not in wb.sheetnames:
            continue
        ws = wb[hoja]
        col_base = ws.max_column + 2
        ws.cell(row=1, column=col_base, value="Especialidad")
        ws.cell(row=1, column=col_base + 1, value="Subespecialidad")
        ws.cell(row=1, column=col_base + 2, value="Código")
        ws.cell(row=1, column=col_base + 3, value="Ratio (S//m2)")
        for it in items_hoja:
            fila = it["fila_origen"]
            subtotal = it.get("subtotal_soles") or 0
            ratio = round(subtotal / area, 2) if area else None
            ws.cell(row=fila, column=col_base, value=it["especialidad"])
            ws.cell(row=fila, column=col_base + 1, value=it["subespecialidad"])
            ws.cell(row=fila, column=col_base + 2, value=it["codigo"])
            ws.cell(row=fila, column=col_base + 3, value=ratio)

    wb.save(out_path)
    print(f"Excel original + clasificación al costado, guardado en: {out_path}")
    return out_path


def exportar_items_excel(resultado, out_path):
    """Exporta TODAS las partidas de un proyecto clasificado a un archivo Excel <nombre>_clasificado.xlsx
    con todas las columnas mínimas requeridas por auditoría."""
    area = resultado.area_techada
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Partidas clasificadas"
    headers = [
        "Hoja origen", "Código origen", "Descripción", "Unidad", "Metrado",
        "Parcial (S/)", "Parcial (USD)", "Subtotal (S/)", "Ratio (S/./m2)",
        "Especialidad", "Subespecialidad", "Código", "Score",
        "Método clasificación", "Proyecto ref. histórico", "Descripción ref. histórico", "Observaciones"
    ]
    ws.append(headers)
    for it in resultado.items:
        subtotal = it.get("subtotal_soles") or 0
        ratio = round(subtotal / area, 2) if area else None
        ws.append([
            it.get("hoja_origen"),
            it.get("codigo_origen"),
            it.get("descripcion"),
            it.get("unidad"),
            it.get("metrado"),
            it.get("parcial_soles"),
            it.get("parcial_usd"),
            subtotal,
            ratio,
            it.get("especialidad"),
            it.get("subespecialidad"),
            it.get("codigo"),
            it.get("score_clasificacion"),
            it.get("metodo_clasificacion"),
            it.get("ref_proyecto"),
            it.get("ref_descripcion"),
            it.get("observaciones"),
        ])
    wb.save(out_path)
    print(f"Archivo clasificado exportado exitosamente a: {out_path}")
    return out_path


def exportar_ratios_excel(resultado, out_path):
    filas = tabla_ratios_proyecto(resultado)
    if not filas:
        return None
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Ratios"
    headers = list(filas[0].keys())
    ws.append(headers)
    for f in filas:
        ws.append([f[h] for h in headers])
    wb.save(out_path)
    print(f"Ratios exportados a: {out_path}")
    return out_path


# ===========================================================================
# PARTE 5: GUARDAR EN LA BASE HISTÓRICA ACUMULATIVA (APRENDIZAJE INCREMENTAL)
# ===========================================================================

def guardar_en_base(path_bd_historico, resultado, tipo=None, departamento=None,
                     id_proyecto=None, sobrescribir_si_existe=False):
    """Agrega todas las partidas clasificadas de `resultado` a la hoja 'BD_HISTORICO PPTO'
    dentro del archivo único acumulativo `path_bd_historico`.
    
    - Verifica si el proyecto ya existe para no duplicarlo accidentalmente.
    - Si sobrescribir_si_existe es True y el proyecto ya existe, elimina sus filas antiguas antes de guardar.
    - No modifica la hoja 'Clasificación'.
    """
    wb = openpyxl.load_workbook(path_bd_historico)
    if "BD_HISTORICO PPTO" not in wb.sheetnames:
        raise ValueError(f"No se encontró la hoja 'BD_HISTORICO PPTO' en {path_bd_historico}")

    ws = wb["BD_HISTORICO PPTO"]
    fila_header, COL = _mapear_columnas_bd(ws)

    # Comprobar proyectos existentes y sus filas
    filas_existentes_proy = []
    proyectos_en_base = set()
    col_proy = COL.get("proyecto", 1)

    for r in range(fila_header + 1, ws.max_row + 1):
        val_proy = ws.cell(row=r, column=col_proy).value
        if val_proy is not None and str(val_proy).strip():
            p_str = str(val_proy).strip()
            proyectos_en_base.add(p_str)
            if p_str.upper() == resultado.proyecto.strip().upper():
                filas_existentes_proy.append(r)

    if filas_existentes_proy:
        if not sobrescribir_si_existe:
            print(f"\n[AVISO] El proyecto '{resultado.proyecto}' YA existe en la base histórica ({len(filas_existentes_proy)} partidas registradas).")
            print("Para reprocesarlo y reemplazarlo, ejecuta con 'sobrescribir_si_existe=True' o flag '--forzar'. No se agregaron duplicados.")
            return False
        else:
            print(f"\n[ACTUALIZACIÓN] Eliminando registros anteriores del proyecto '{resultado.proyecto}' ({len(filas_existentes_proy)} partidas)...")
            for r in reversed(filas_existentes_proy):
                ws.delete_rows(r, 1)

    # Determinar la fila de inserción (después del último registro real o del header)
    ultima_fila_con_datos = fila_header
    for r in range(ws.max_row, fila_header, -1):
        if any(ws.cell(row=r, column=c).value is not None for c in range(1, ws.max_column + 1)):
            ultima_fila_con_datos = r
            break

    fila_inicio = ultima_fila_con_datos + 1
    hoy = datetime.date.today()
    id_proyecto = id_proyecto or f"PROY-{hoy.year}-{resultado.proyecto.upper().replace(' ', '')[:12]}"
    area = resultado.area_techada

    n_escritas = 0
    fila_actual = fila_inicio
    for it in resultado.items:
        esp = it.get("especialidad") or "SIN_CLASIFICAR"
        metrado = it.get("metrado") or 1
        parcial_s = it.get("parcial_soles") or 0
        parcial_d = it.get("parcial_usd") or 0
        subtotal = it.get("subtotal_soles") or (parcial_s + parcial_d * (it.get("tc") or 0))
        ratio = round(subtotal / area, 2) if area else None

        valores = {
            "fecha": hoy,
            "id_proyecto": id_proyecto,
            "proyecto": resultado.proyecto,
            "tipo": tipo,
            "departamento": departamento,
            "especialidad": esp,
            "subespecialidad": it.get("subespecialidad"),
            "codigo": it.get("codigo"),
            "descripcion": it.get("descripcion"),
            "unidad": it.get("unidad"),
            "metrado": metrado,
            "pu_soles": round(parcial_s / metrado, 2) if metrado else None,
            "parcial_soles": parcial_s,
            "pu_usd": round(parcial_d / metrado, 2) if metrado else None,
            "parcial_usd": parcial_d,
            "subtotal_soles": subtotal,
            "tc": it.get("tc"),
            "area": area,
            "ratio": ratio,
            "origen_hoja": it.get("hoja_origen"),
            "origen_item": it.get("codigo_origen"),
        }

        for campo, valor in valores.items():
            col = COL.get(campo)
            if col:
                ws.cell(row=fila_actual, column=col).value = valor

        fila_actual += 1
        n_escritas += 1

    wb.save(path_bd_historico)
    print(f"Se agregaron exitosamente {n_escritas} partidas de '{resultado.proyecto}' a 'BD_HISTORICO PPTO' en '{path_bd_historico}'.")
    return True


# ===========================================================================
# PARTE 6: SISTEMA INCREMENTAL PRINCIPAL (COLAB / PYTHON / CLI)
# ===========================================================================

def procesar_presupuesto_incremental(
    path_excel,
    path_bd_historico="BD_HISTORICO_BASE_VACIA.xlsx",
    nombre_proyecto=None,
    area_techada=None,
    tipo=None,
    departamento=None,
    sobrescribir_si_existe=False,
    guardar_en_historico=True,
    exportar_clasificado=True,
    out_excel_clasificado=None,
):
    """Función principal para el flujo incremental de clasificación de presupuestos.

    Flujo:
      1. Carga la base histórica actual desde `path_bd_historico`.
      2. Si está vacía (primer presupuesto), clasifica usando reglas KIMSAV + hoja auxiliar.
      3. Si contiene proyectos anteriores, utiliza su información para clasificar por similitud.
      4. Genera el reporte completo y el Excel clasificado `<nombre>_clasificado.xlsx`.
      5. Actualiza automáticamente la hoja 'BD_HISTORICO PPTO' en `path_bd_historico`.
      6. Muestra el estado del aprendizaje incremental (proyectos antes vs después).
    """
    if nombre_proyecto is None:
        base_filename = path_excel.replace("\\", "/").rsplit("/", 1)[-1]
        nombre_proyecto = base_filename.rsplit(".", 1)[0]

    # 1. Cargar histórico acumulado
    historico = cargar_historico(path_bd_historico)
    n_proyectos_antes = len(historico)

    print(f"\n=======================================================")
    print(f"PROCESANDO PRESUPUESTO INCREMENTAL: {nombre_proyecto}")
    print(f"Proyectos históricos acumulados como referencia: {n_proyectos_antes}")
    if n_proyectos_antes > 0:
        print(f"Referencias disponibles: {', '.join(historico.keys())}")
    else:
        print("Base histórica vacía. Se clasifica usando reglas KIMSAV y hoja de origen.")
    print(f"=======================================================\n")

    # 2. Procesar y clasificar
    resultado = procesar_presupuesto(
        path_excel=path_excel,
        nombre_proyecto=nombre_proyecto,
        area_techada=area_techada,
        historico=historico,
    )

    resultado.proyectos_hist_antes = n_proyectos_antes

    # 3. Exportar Excel clasificado
    if exportar_clasificado:
        if out_excel_clasificado is None:
            out_excel_clasificado = f"{nombre_proyecto}_clasificado.xlsx"
        exportar_items_excel(resultado, out_excel_clasificado)

    # 4. Guardar en base histórica
    if guardar_en_historico:
        guardado_ok = guardar_en_base(
            path_bd_historico=path_bd_historico,
            resultado=resultado,
            tipo=tipo,
            departamento=departamento,
            sobrescribir_si_existe=sobrescribir_si_existe,
        )
        historico_post = cargar_historico(path_bd_historico)
        resultado.proyectos_hist_despues = len(historico_post)
    else:
        resultado.proyectos_hist_despues = n_proyectos_antes

    # 5. Imprimir auditoría completa
    resultado.imprimir_resumen()

    return resultado


# ===========================================================================
# CLI
# ===========================================================================

if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Sistema incremental de clasificación de presupuestos KIMSAV")
    ap.add_argument("excel", help="Ruta al archivo Excel de presupuesto a procesar")
    ap.add_argument("--bd", default="BD_HISTORICO_BASE_VACIA.xlsx", help="Ruta al archivo de base histórica")
    ap.add_argument("--proyecto", default=None, help="Nombre del proyecto")
    ap.add_argument("--area", type=float, default=None, help="Área techada en m2")
    ap.add_argument("--tipo", default=None, help="Tipo de proyecto (e.g. Edificio Multifamiliar)")
    ap.add_argument("--departamento", default=None, help="Departamento/Ubicación")
    ap.add_argument("--forzar", action="store_true", help="Sobrescribir si el proyecto ya existe en el histórico")
    ap.add_argument("--no-guardar", action="store_true", help="No guardar en el histórico")
    args = ap.parse_args()

    procesar_presupuesto_incremental(
        path_excel=args.excel,
        path_bd_historico=args.bd,
        nombre_proyecto=args.proyecto,
        area_techada=args.area,
        tipo=args.tipo,
        departamento=args.departamento,
        sobrescribir_si_existe=args.forzar,
        guardar_en_historico=not args.no_guardar,
    )
