"""
classify_rules.py
------------------
Motor de clasificación por palabras clave para partidas de presupuestos de
construcción (KIMSAV), calibrado contra el catálogo REAL de la pestaña
"Clasificación" (14 especialidades, cada una con sus subespecialidades y
códigos numéricos oficiales -- ver SUBCODES más abajo).

Cada especialidad tiene su propia función classify_XXX(desc) que devuelve
(subespecialidad, codigo, score) -- el score mide qué tan fuerte es la
coincidencia, para poder comparar una misma descripción contra TODAS las
especialidades (ver score_all / best_match) y saber si la que tiene
asignada hoy sigue siendo la mejor opción o si hay otra más clara.

Uso:
    from classify_rules import score_all, best_match

    scores = score_all("CONTROL REMOTO PARA PUERTA DE GARAGE")
    # -> {'IIEE': {'score': 8, 'subespecialidad': 'Instalaciones Eléctricas',
    #              'codigo': 501}, ...}

    especialidad, subespecialidad, codigo, score = best_match(desc)
"""

import re


def U(s):
    if s is None:
        return ""
    return str(s).upper().strip()


# ---------------------------------------------------------------------------
# Catálogo oficial (especialidad -> {subespecialidad: código}), tomado
# directamente de la pestaña "Clasificación" de BD_Historico_presupuestos.xlsx
# ---------------------------------------------------------------------------

SUBCODES = {
    "OP": {
        "Instalaciones provisionales": 101,
        "Transporte vertical y horizontal": 102,
        "Seguridad de obra": 103,
        "Andamios": 104,
        "Topografía": 105,
        "Servicios": 106,
        "Obras exteriores": 107,
        "Limpieza": 108,
    },
    "ESTT": {
        "Movimiento de tierras": 201,
        "Concreto premezclado MP": 202,
        "Concreto premezclado Cimentaciones": 203,
        "Concreto premezclado": 204,
        "Prelosa": 205,
        "Prelosa instalación": 206,
        "Acero MP": 207,
        "Acero Cimentaciones": 208,
        "Acero": 209,
        "Encofrado MP": 210,
        "Encofrado Cimentaciones": 211,
        "Encofrado simple": 212,
        "Encofrado doble altura": 213,
        "Anclajes": 214,
    },
    "ARQ": {
        "Tabiquería": 301,
        "Drywall": 302,
        "Tarrajeo": 303,
        "Solaqueo": 304,
        "Impermeabilización": 305,
        "Contrapisos": 306,
        "Enchapes y revestimiento (piso y pared)": 307,
        "Carpintería de madera": 308,
        "Carpintería metálica": 309,
        "Carpintería de melamine": 310,
        "Vidrios y cristales": 311,
        "Tableros": 312,
        "Aparatos sanitarios y griferías": 313,
        "Luminarias": 314,
        "Pintura": 315,
        "Papel mural": 316,
        "Fachada": 317,
        "Equipamiento depas": 318,
        "Paisajismo": 319,
        "Señalización": 320,
        "Entrega de departamentos": 321,
        "Equipamiento de áreas comunes": 322,
        "Sellos cortafuego": 323,
        "Estacionamiento de bicicletas": 324,
        "Puertas, ventanas y mamparas": 325,
    },
    "IISS": {
        "Instalaciones Sanitarias": 401,
        "Bombas SPC": 402,
    },
    "IIEE": {"Instalaciones Eléctricas": 501},
    "IIMM": {"Instalaciones Mecánicas": 601},
    "GAS": {"Instalaciones de Gas": 701},
    "ACI": {
        "Red ACI": 801,
        "Bomba ACI": 802,
    },
    "CCDÉBILES": {
        "DACI": 901,
        "CCTV": 902,
        "Intercomunicadores": 903,
    },
    "EQUIP": {
        "Grupo Electrógeno": 1001,
        "Acelerógrafo": 1002,
        "Paneles Solares": 1003,
        "Equipamiento Piscina": 1004,
    },
    "ASCENSORES": {"Ascensores": 1101},
    "GG": {"Gastos Generales": 1201},
    "UTIL": {"Utilidad": 1301},
    "IGV": {"IGV": 1401},
}

NON_ITEM_SPECIALTIES = {"GG", "UTIL", "IGV"}


def _sub(especialidad, nombre):
    return nombre, SUBCODES[especialidad][nombre]


# ---------------------------------------------------------------------------
# Clasificadores por especialidad.
# Cada uno recibe la descripción en MAYÚSCULAS y devuelve una lista de
# (score, subespecialidad, codigo) -- se usa el de mayor score si hay
# varias coincidencias dentro de la misma especialidad.
# ---------------------------------------------------------------------------

def classify_OP(d):
    hits = []
    if re.search(r"OFICINA\s*(DE\s*OBRA|RESIDENCIA|SALA\s*DE\s*REUNIONES)|ALMAC[EÉ]N\s*(DE\s*OBRA|CON\s*ESTANTER[IÍ]A)|"
                 r"COMEDOR\s*(DE\s*)?(PERSONAL|OBRERO)|"
                 r"VESTUARIO\s*(PARA\s*)?OBRERO|CAMPAMENTO\s*PROVISIONAL|BA[ÑN]OS?\s*PORT[AÁ]TILES|"
                 r"SSHH\s*PORT[AÁ]TILES|DUCHAS?\s*PORT[AÁ]TILES|LAVAMANOS|CASETA\s*(PARA|DE)\s*GUARDIAN[IÍ]A|"
                 r"CARTEL\s*DE\s*OBRA|CERCO\s*PERIM[EÉ]TRIC|CONSTRUCCI[OÓ]N\s*DE\s*SSHH", d):
        hits.append((9, *_sub("OP", "Instalaciones provisionales")))
    # "grúa" suelta dentro de una frase de excavación/demolición no cuenta
    # -- solo cuando la grúa/el transporte ES la partida (alquiler,
    # movilización, montacargas de obra, etc.)
    if re.search(r"ALQUILER\s*(DE\s*)?GR[UÚ]A|GR[UÚ]A\s*TORRE|MOVILIZACI[OÓ]N\s*DE\s*GR[UÚ]A|"
                 r"MONTACARGA\s*DE\s*OBRA|\bWINCHE\b|\bMANLIFT\b|ASCENSOR\s*DE\s*OBRA|"
                 r"(CIMENTACI[OÓ]N|ZAPATA|PATAS\s*DE\s*EMPOTRAMIENTO|ALQUILER\s*MENSUAL)\s*(PARA\s*)?(DE\s*)?TORRE\s*GR[UÚ]A|"
                 r"TRANSPORTE\s*VERTICAL\s*CON\s*GR[UÚ]A|OPERADOR\s*DE\s*(GR[UÚ]A|ELEVADOR)|"
                 r"\bRIGGER\b|PULPO\s*DE\s*CADENA|ALQUILER\s*(DE\s*)?ELEVADOR\s*DE\s*PERSONAL|"
                 r"COMBUSTIBLE\s*PARA\s*ELEVADOR|REVISI[OÓ]N\s*(MENSUAL\s*)?DE\s*ELEVADOR|"
                 r"CHUTE\s*(PARA\s*)?ELIMINACI[OÓ]N\s*VERTICAL", d):
        hits.append((8, *_sub("OP", "Transporte vertical y horizontal")))
    if re.search(r"SEGURIDAD\s*(Y\s*SALUD|INDUSTRIAL)|\bEPP\b|EQU[IÍ]?POS?\s*DE\s*PROTECCI[OÓ]N\s*(INDIVIDUAL|COLECTIVA)|"
                 r"SE[ÑN]ALIZACI[OÓ]N\s*(DE\s*)?SEGURIDAD|MALLAS?\s*(DE\s*SEGURIDAD|ANTICA[IÍ]DAS)|"
                 r"EXTINTOR\s*DE\s*OBRA|SERVICIO\s*DE\s*GUARDIAN[IÍ]A|EXAMEN\s*M[EÉ]DICO|"
                 r"DESV[IÍ]O\s*DE\s*TR[AÁ]NSITO|SE[ÑN]ALERO|MONITOR(ES)?\s*DE\s*SEGURIDAD|"
                 r"\bLLAVEROS?\b\s*(\(CONTROL\))?", d):
        hits.append((7, *_sub("OP", "Seguridad de obra")))
    if re.search(r"\bANDAMIO", d):
        hits.append((8, *_sub("OP", "Andamios")))
    if re.search(r"TRAZO\s*Y\s*REPLANTEO|TOPOGRAF[IÍ]A|NIVELACI[OÓ]N\s*TOPOGR[AÁ]FIC|TRAZO(S)?.{0,20}REPLANTEO(S)?", d):
        hits.append((8, *_sub("OP", "Topografía")))
    if re.search(r"CONSUMO\s*DE\s*(AGUA|ENERG[IÍ]A|INTERNET)\b|INSTALACI[OÓ]N(ES)?\s*PROVISIONAL(ES)?\s*(DE\s*)?"
                 r"(AGUA|ENERG[IÍ]A|DESAG[UÜ]E|TELEFON[IÍ]A)?|MOVILIZACI[OÓ]N\s*Y\s*DESMOVILIZACI[OÓ]N|FLETE|"
                 r"LIMPIEZA\s*(PERMANENTE|FINAL|FINA)\s*(DE\s*)?(OBRA|EDIFICIO|DEPARTAMENTOS|ZONAS\s*COMUNES)|"
                 r"ENSAYOS?\s*(DE\s*)?(CALIDAD|ESTANDAR|"
                 r"DE\s*DENSIDAD|DE\s*CLASIFICACI[OÓ]N|DE\s*LABORATORIO\s*DE\s*SUELOS)|PROBETAS\s*DE\s*CONCRETO|"
                 r"AMOLADORA(S)?|HIDROLAVADORA", d):
        hits.append((7, *_sub("OP", "Servicios")))
    if re.search(r"OBRAS\s*EXTERIORES|VEREDA(S)?\s*(EXTERIOR)?|PAVIMENTO\s*EXTERIOR|REPOSICI[OÓ]N\s*DE\s*VEREDA", d):
        hits.append((7, *_sub("OP", "Obras exteriores")))

    # transporte vertical temporal de obra (elevador/montacargas de obra) --
    # distinto del Ascensor definitivo del edificio (ese lleva la palabra
    # "ASCENSOR" y lo captura classify_ASCENSORES).
    if re.search(r"TRANSPORTE\s*VERTICAL\s*(CON\s*)?ELEVADOR|(MONTAJE|DESMONTAJE)\s*(Y\s*(MONTAJE|DESMONTAJE)\s*)?"
                 r"DE\s*ELEVADOR\b|CARGA\s*DE\s*TORRE\s*GR[UÚ]A|"
                 r"(MONTAJE|DESMONTAJE)\s*(Y\s*(MONTAJE|DESMONTAJE)\s*)?DE\s*TORRE\s*GR[UÚ]A", d):
        hits.append((9, *_sub("OP", "Transporte vertical y horizontal")))

    if re.search(r"MOVILIZACI[OÓ]N\s*(DE\s*)?CAMPAMENTO|DESMOVILIZACI[OÓ]N\s*(DE\s*)?CAMPAMENTO|"
                 r"CERCO\s*(DE\s*)?OBRA|CERCO\s*MET[AÁ]LICO|CASETA\s*DE\s*VENTAS|"
                 r"ALQUILER\s*(DE\s*)?CONTENEDOR|ESCUADRAS\s*MET[AÁ]LICAS|"
                 r"MOVILIZACI[OÓ]N\s*Y\s*DESMOVIL[IZ]{1,3}ACI[OÓ]N\s*DE\s*EQUIPOS|"
                 r"ALQUILER\s*(Y\s*ARMADO\s*)?DE\s*ESCALERA\b|MALLA\s*RA[CS]HEL", d):
        hits.append((8, *_sub("OP", "Instalaciones provisionales")))
    # monitoreos y controles ambientales/ocupacionales del propio proceso de
    # obra (no son una instalación técnica, son un servicio de control).
    if re.search(r"MONITOREO\s*(DE\s*)?(SALUD\s*OCUPACIONAL|AMBIENTAL|OCUPACIONAL|CALIDAD\s*DE\s*AIRE|AIRE\s*Y\s*RUIDO)", d):
        hits.append((7, *_sub("OP", "Seguridad de obra")))
    # protección y trabajos preliminares frente a predios/vecinos colindantes,
    # y demolición de estructuras/elementos existentes antes de empezar obra
    # nueva -- trabajo preliminar de obra, no estructura propia del edificio.
    if re.search(r"PROTECCI[OÓ]N\s*(PERIMETRAL\s*)?(INICIAL\s*)?(CONTRA|DE)\s*VECINOS|"
                 r"DEMOLICI[OÓ]N\s*DE\s*ESTRUCTURAS\s*EXISTENTES|"
                 r"DEMOLICI[OÓ]N\s*DE\s*CACHIMBAS|PROTECCI[OÓ]N\s*CON\s*CART[OÓ]N\s*Y\s*PL[AÁ]STICO|"
                 r"VENTILACI[OÓ]N\s*PROVISIONAL\s*EN\s*S[OÓ]TANOS", d):
        hits.append((8, *_sub("OP", "Instalaciones provisionales")))
    # logística interna de obra: traslado de materiales con parihuelas/stockas
    if re.search(r"PARIHUELAS|\bSTOCKAS?\b", d):
        hits.append((6, *_sub("OP", "Servicios")))
    return hits


def classify_ESTT(d):
    hits = []
    es_mp = bool(re.search(r"MURO\s*PANTALLA|MURO\s*ANCLADO", d))
    es_cim = bool(re.search(r"ZAPATA|CIMENTACI[OÓ]N|PLATEA\s*DE\s*CIMENTACI[OÓ]N", d))
    es_doble_altura = bool(re.search(r"DOBLE\s*ALTURA", d))

    if re.search(r"EXCAVACI[OÓ]N\s*(MASIVA|ESTRUCTURAL)|RELLENO\s*ESTRUCTURAL|CORTE\s*DE\s*TERRENO|"
                 r"CORTE\s*Y\s*ELIMINACI[OÓ]N|ELIMINACI[OÓ]N\s*(DE\s*)?MATERIAL(\s*DE)?\s*(DESMONTE)?|"
                 r"ELIMINACI[OÓ]N\s*(DE\s*)?DESMONTE|"
                 r"EXTRACCI[OÓ]N\s*DE\s*(MATERIAL|DESMONTE|EQUIPOS?\s*(DE\s*)?PERFORACI[OÓ]N)|"
                 r"ACARREO\s*(DE\s*)?(DESMONTE|MATERIAL|HORIZONTAL)|TRASLADO\s*Y\s*ACARREO|"
                 r"RELLENO\s*(COMPACTADO|CON\s*MATERIAL)|BASE\s*(DE\s*)?AFIRMADO|PROCTOR\s*MODIFICADO|"
                 r"NIVELACI[OÓ]N\s*Y\s*COMPACTACI[OÓ]N|COMPACTACI[OÓ]N\s*DE\s*TERRENO|"
                 r"EXCAVACI[OÓ]N\s*LOCALIZADA|EXCAVACI[OÓ]N\s*DE\s*ZANJA", d):
        hits.append((9, *_sub("ESTT", "Movimiento de tierras")))

    if re.search(r"CONCRETO\s*(F'?C|PREMEZCLADO|ARMADO)|\bF'C\b", d):
        if es_mp:
            hits.append((10, *_sub("ESTT", "Concreto premezclado MP")))
        elif es_cim:
            hits.append((10, *_sub("ESTT", "Concreto premezclado Cimentaciones")))
        else:
            hits.append((8, *_sub("ESTT", "Concreto premezclado")))

    if re.search(r"PRELOSA", d):
        if re.search(r"INSTALACI[OÓ]N|MONTAJE|COLOCACI[OÓ]N", d):
            hits.append((9, *_sub("ESTT", "Prelosa instalación")))
        else:
            hits.append((9, *_sub("ESTT", "Prelosa")))

    # ojo: NO se usa "\bACERO\b" suelto -- "acero inoxidable" aparece en
    # accesorios de baño (toallero, papelera) y en bombas, que no son
    # estructura. Solo cuenta el acero de refuerzo/estructural.
    if re.search(r"ACERO\s*DE\s*REFUERZO|FIERRO\s*CORRUGADO|ACERO\s*ESTRUCTURAL|ACERO\s*CORRUGADO|"
                 r"ACERO\s*(PARA\s*)?(MUROS?|COLUMNAS?|VIGAS?|LOSAS?|PLACAS?)\b|ACERO\s*F.?Y\s*=", d) \
            and not re.search(r"INOX", d):
        if es_mp:
            hits.append((10, *_sub("ESTT", "Acero MP")))
        elif es_cim:
            hits.append((10, *_sub("ESTT", "Acero Cimentaciones")))
        else:
            hits.append((8, *_sub("ESTT", "Acero")))

    if re.search(r"ENCOFRADO|DESENCOFRADO", d):
        if es_mp:
            hits.append((10, *_sub("ESTT", "Encofrado MP")))
        elif es_cim:
            hits.append((10, *_sub("ESTT", "Encofrado Cimentaciones")))
        elif es_doble_altura:
            hits.append((10, *_sub("ESTT", "Encofrado doble altura")))
        else:
            hits.append((8, *_sub("ESTT", "Encofrado simple")))

    if re.search(r"\bANCLAJE|MURO\s*ANCLADO\b|PASES\s*PARA\s*ANCLAJE", d) and not es_mp:
        hits.append((7, *_sub("ESTT", "Anclajes")))

    # curado de concreto -- siempre es un paso del vaciado estructural,
    # sin importar si el elemento curado está expuesto o no.
    if re.search(r"CURADO\s*(DE\s*)?(ELEMENTOS?|CONCRETO)", d):
        hits.append((9, *_sub("ESTT", "Concreto premezclado")))

    # sardinel / forjado de pasos y contrapasos de escalera: son vaciados de
    # concreto menores (no parte de la estructura principal) -- distintos de
    # un sardinel ya cubierto por otra especialidad más específica si la hay.
    if re.search(r"\bSARDINEL\b|FORJADO\s*(DE\s*)?(PASO|CONTRAPASO|DESCANSO)", d):
        hits.append((6, *_sub("ESTT", "Concreto premezclado")))

    # parapetos de concreto, juntas de vaciado/dilatación -- trabajo de
    # concreto menor, no estructura principal.
    if re.search(r"\bPARAPETO\b|JUNTA\s*(ASERRADA|DE\s*CONSTRUCCI[OÓ]N|DE\s*DILATACI[OÓ]N|DE\s*BORDE|"
                 r"DE\s*CONTRACCI[OÓ]N|DE\s*VACIADO)\b|\bJC\b|\bJV\b|"
                 r"DINTEL(ES)?\s*DE\s*CONCRETO", d):
        hits.append((6, *_sub("ESTT", "Concreto premezclado")))

    # elementos e insumos menores del vaciado de concreto estructural:
    # waterstop (junta de impermeabilización entre vaciados), dowells
    # (pasadores de transferencia de carga), perfilado y pañeteo de
    # superficies, epóxico de unión entre vaciados nuevos y antiguos.
    if re.search(r"\bWATERSTOP\b|\bDOWELLS?\b|PERFILADO\s*(GRUESO|FINO)|PA[NÑ]ETEO\s*DE\s*SUELO|"
                 r"EP[OÓ]XICO\s*DE\s*UNI[OÓ]N\s*(DE\s*)?CONCRETO", d):
        hits.append((7, *_sub("ESTT", "Concreto premezclado")))

    # bovedilla de arcilla / ladrillo de techo: aligerado de losas, parte del
    # vaciado estructural de la losa (no es tabiquería/albañilería de muros).
    if re.search(r"TECNOPOR\s*PARA\s*CAJUELAS", d):
        hits.append((7, *_sub("ESTT", "Concreto premezclado")))

    if re.search(r"BOVEDILLA\s*DE\s*ARCILLA|LADRILLO\s*DE\s*TECHO", d):
        hits.append((8, *_sub("ESTT", "Concreto premezclado")))

    # trabajos sobre un muro pantalla ya identificado (picado, excavación
    # puntual, tratamiento de encuentros) que no mencionan explícitamente
    # concreto/acero/encofrado pero igual son parte de esa partida estructural.
    if es_mp and not hits:
        hits.append((5, *_sub("ESTT", "Concreto premezclado MP")))

    if not hits and re.search(r"\bVIGA(S)?\b|\bCOLUMNA(S)?\b|\bPLACA(S)?\s*(DE\s*)?CONCRETO\b|"
                               r"LOSA\s*(ALIGERADA|MACIZA)|MURO\s*DE\s*CONTENCI[OÓ]N", d):
        hits.append((7, *_sub("ESTT", "Concreto premezclado")))

    return hits


def classify_ARQ(d):
    hits = []
    # instalación/mano de obra de aparatos sanitarios es IISS (ver classify_IISS);
    # aquí solo el ARTEFACTO/modelo en sí (sin "instalación de" explícito).
    if re.search(r"(INODORO|LAVATORIO|GRIFER[IÍ]A|\bDUCHA\b|\bTINA\b|URINARIO|LAVADERO|MEZCLADORA|ONE\s*PIECE)", d) \
            and not re.search(r"INSTALACI[OÓ]N\s*DE|SUMINISTRO\s*E\s*INSTALACI[OÓ]N", d):
        hits.append((11, *_sub("ARQ", "Aparatos sanitarios y griferías")))

    # mobiliario/accesorios de baño (no son griferías/aparatos pero van en el
    # mismo rubro del catálogo): vanitorio, toallero, papelera, espejo de SSHH.
    if re.search(r"\bVANITORIO\b|TOALLERO|PAPELERA|BARRA\s*(PARA\s*)?COLGADOR|"
                 r"ESPEJO\s*(EN\s*)?SSHH|BARRAS?\s*DE\s*SEGURIDAD\s*EN\s*SSHH", d):
        hits.append((10, *_sub("ARQ", "Aparatos sanitarios y griferías")))

    # luminaria como artefacto/modelo (no como salida/instalación eléctrica)
    if re.search(r"LUMINARIA|FLUORESCENTE|REFLECTOR|SPOT\s*LIGHT|DICROICO", d) \
            and not re.search(r"^SALIDA\b|SALIDA\s*PARA|INSTALACI[OÓ]N\s*DE\s*LUMINARIA", d):
        hits.append((10, *_sub("ARQ", "Luminarias")))

    if re.search(r"SOLAQUEO|EMPASTAD|RESANE|\bDERRAME(S)?\b|VESTIDURA\s*DE\s*DERRAME|"
                 r"FROTACHAD|CAL\s*\+?\s*CEMENTO|SOBREMURO", d):
        hits.append((12, *_sub("ARQ", "Solaqueo")))

    if re.search(r"\bCZ\b|ZOCALO|Z[OÓ]CALO|PORCELANATO|CER[AÁ]MIC[OA]\b|ENCHAPE", d):
        hits.append((11, *_sub("ARQ", "Enchapes y revestimiento (piso y pared)")))

    if re.search(r"\bTARRAJEO\b", d):
        hits.append((10, *_sub("ARQ", "Tarrajeo")))

    if re.search(r"CONTRAPISO|NIVELACI[OÓ]N\s*DE\s*PISO", d):
        hits.append((9, *_sub("ARQ", "Contrapisos")))

    if re.search(r"PISO\s*(PORCEL|VIN[IÍ]LIC|LAMINAD)|\bSPC\s*CLICK\b|\bPISOPAK\b", d):
        hits.append((9, *_sub("ARQ", "Enchapes y revestimiento (piso y pared)")))

    if re.search(r"PISOS?\s*(DE\s*)?CEMENTO\s*(SEMI)?PULIDO", d):
        hits.append((9, *_sub("ARQ", "Enchapes y revestimiento (piso y pared)")))

    if re.search(r"IMPERMEABILIZACI[OÓ]N", d):
        hits.append((9, *_sub("ARQ", "Impermeabilización")))

    if re.search(r"MURO\s*(DE\s*)?LADRILLO|TABIQUER[IÍ]A|TABIQUE\s*(DE\s*)?LADRILLO|"
                 r"DINTEL(ES)?\s*(DE\s*)?LADRILLO|DINTEL(ES)?\s*S[IÍ]LICO\s*CALC[AÁ]REO", d):
        hits.append((8, *_sub("ARQ", "Tabiquería")))

    if re.search(r"CIELO\s*RASO|DRYWALL|MOLDURA(S)?\s*(DE\s*)?POLIESTIRENO", d):
        hits.append((8, *_sub("ARQ", "Drywall")))

    if re.search(r"MELAMINE|\bMDF\b", d):
        hits.append((9, *_sub("ARQ", "Carpintería de melamine")))
    elif re.search(r"PUERTA\s*(CONTRAPLACAD|DE\s*MADERA)|MARCO\s*DE\s*PUERTA\s*DE\s*MADERA|"
                   r"CONTRAPLACAD[AO]|CONTRAPL\.|\bVAIV[EÉ]N\b|CERRADURA|BISAGRA|MANIJA\s*(PARA\s*)?PUERTA|"
                   r"PUERTA(S)?\s*BATIENTE|\bBATIENTE\b", d):
        hits.append((7, *_sub("ARQ", "Carpintería de madera")))

    if re.search(r"\bCARPINTER[IÍ]A\s*MET[AÁ]LICA\b|BARANDAS?\s*(MET[AÁ]LICA|DE\s*(FIERRO|ACERO))|"
                 r"REJA\s*MET[AÁ]LICA|\bBARANDAS?\b|\bPASAMANOS\b|PORT[OÓ]N\s*VEHICULAR|"
                 r"PUERTA\s*VEHICULAR|ESTRUCTURA\s*MET[AÁ]LICA\s*(DE\s*)?SOL\s*Y\s*SOMBRA|"
                 r"CELOS[IÍ]A", d):
        hits.append((8, *_sub("ARQ", "Carpintería metálica")))

    if re.search(r"VENTANA\s*(DE\s*ALUMINIO|CORREDIZA)|MAMPARA|BARANDA\s*DE\s*VIDRIO|\bVIDRIO\b|CRISTAL", d):
        hits.append((7, *_sub("ARQ", "Vidrios y cristales")))

    # Tableros ARQ = mesones (granito/cuarzo/melamine de cocina), NO tablero eléctrico
    if re.search(r"TABLERO\s*(DE\s*)?(GRANITO|M[AÁ]RMOL|CUARZO|COCINA)", d):
        hits.append((12, *_sub("ARQ", "Tableros")))

    if re.search(r"PINTURA\s*(L[AÁ]TEX|ESMALTE)|^PINTURA\b|^ACABADO\b", d):
        hits.append((7, *_sub("ARQ", "Pintura")))

    if re.search(r"PAPEL\s*MURAL|PREPARACI[OÓ]N\s*DE\s*PAPEL", d):
        hits.append((9, *_sub("ARQ", "Papel mural")))

    if re.search(r"\bFACHADA\b|PARAPETO(S)?\s*(EN\s*VENTANAS)?", d):
        hits.append((8, *_sub("ARQ", "Fachada")))

    if re.search(r"COCINA\s*(INTEGRAL|AMOBLADA)|MUEBLE\s*DE\s*COCINA|CLOSET\b|COCINA\s*[-–]\s*\d{1,3}\b", d):
        hits.append((8, *_sub("ARQ", "Equipamiento depas")))

    if re.search(r"PAISAJISMO|[AÁ]REAS?\s*VERDES|JARDINER[IÍ]A", d):
        hits.append((8, *_sub("ARQ", "Paisajismo")))

    if re.search(r"SE[ÑN]ALIZACI[OÓ]N\s*(DE\s*)?(EVACUACI[OÓ]N|AMBIENTES|N[UÚ]MERO)|"
                 r"SE[ÑN]AL\s*(TIPO\s*BANDERA\s*)?DIRECCIONAL|SE[ÑN]AL\s*DE\s*RUTA\s*DE\s*EVACUACI[OÓ]N|"
                 r"SE[ÑN]ALIZACI[OÓ]N\s*FOTOLUMINISCENTE|SE[ÑN]AL[EÉ]TICA|"
                 r"NUMERACI[OÓ]N\s*DE\s*(DPTOS?|DEPARTAMENTOS?|DEP[OÓ]SITOS?|PISOS?|ESTACIONAMIENTOS?|S[OÓ]TANOS?)", d):
        hits.append((7, *_sub("ARQ", "Señalización")))

    if re.search(r"ENTREGA\s*DE\s*DEPARTAMENTOS?", d):
        hits.append((8, *_sub("ARQ", "Entrega de departamentos")))

    if re.search(r"EQUIPAMIENTO\s*(DE\s*)?[AÁ]REAS?\s*COMUNES|EQUIPAMIENTO\s*DE\s*GIMNASIO|"
                 r"EQUIPAMIENTO\s*DE\s*(SUM|SALA)|BANCA\s*(DE\s*CONCRETO|EN\s*PATIO)|"
                 r"MESA\s*DE\s*PARRILLA", d):
        hits.append((9, *_sub("ARQ", "Equipamiento de áreas comunes")))

    if re.search(r"SELLOS?\s*CORTAFUEGO", d):
        hits.append((9, *_sub("ARQ", "Sellos cortafuego")))

    if re.search(r"ESTACIONAMIENTO(S)?\s*(DE\s*)?BICICLETAS?", d):
        hits.append((9, *_sub("ARQ", "Estacionamiento de bicicletas")))

    # grifería de marca/modelo (artefacto, no instalación -- eso es IISS)
    if re.search(r"LLAVE\s*(TR[EÉ]BOL|ECO\b)|GRIFO\s*(DE\s*)?LAVANDER[IÍ]A|LAVARROPA|MONOMANDO", d) \
            and not re.search(r"INSTALACI[OÓ]N\s*DE|SUMINISTRO\s*E\s*INSTALACI[OÓ]N", d):
        hits.append((10, *_sub("ARQ", "Aparatos sanitarios y griferías")))

    # acabados de piso continuo (microcemento, terrazo) -- junto con
    # porcelanato/cerámico en el mismo rubro de revestimientos.
    if re.search(r"MICROCEMENTO|\bTERRAZO\b", d):
        hits.append((9, *_sub("ARQ", "Enchapes y revestimiento (piso y pared)")))

    # bruñas / cemento bruñado -- detalle de acabado de piso o escalera
    if re.search(r"BRU[ÑN]A(S|DO)?", d):
        hits.append((6, *_sub("ARQ", "Pintura")))

    if re.search(r"KITCHENETTE", d):
        hits.append((8, *_sub("ARQ", "Equipamiento depas")))

    # escalera de gato (acceso metálico fijo, no la escalera de evacuación
    # de concreto) -- carpintería metálica.
    if re.search(r"ESCALERA\s*DE\s*GATO", d):
        hits.append((8, *_sub("ARQ", "Carpintería metálica")))

    # códigos de vano de puerta/ventana/mampara de un cuadro de vanos
    # (ej. "PV.03 - (0.70X2.10) BAÑOS", "MV.01 (2.8x2.1)", "VV.06b - (1.50X1.00)",
    # "PCF. 01(1x2.2)" de puerta cortafuego) -- el separador real entre el
    # código y el número es un PUNTO, no un espacio/guión como asumía antes
    # esta regla, por eso nunca hacía match y esas partidas (S/+2M en un solo
    # proyecto) quedaban sin clasificar.
    if re.search(r"\bPV[.\s-]*\d|\bMV[.\s-]*\d|\bVV[.\s-]*\d|\bPCF[.\s-]*\d|\bMF[.\s-]*\d|\bPS[.\s-]*\d", d):
        hits.append((7, *_sub("ARQ", "Puertas, ventanas y mamparas")))

    # anclas débiles (código de vano) -> aproximación a carpintería de madera
    if not hits and re.search(r"\bVANO\b|\bV[- ]?\d|\bVA[- ]?\d|\bM[- ]?\d|\bMA[- ]?\d|\bMV[- ]?\d|"
                               r"\bPV[- ]?\d|\bPCF[- ]?\d|\bPA[- ]?\d|\bPD[- ]?\d|\bP[- ]?\d|\bCL[- ]?\d|"
                               r"\bC[- ]?\d|\bL\d|\bEMG\b|\bWCL\b|\bB[- ]?\d|\bPC\.|\bW\.?C\.?\b", d):
        hits.append((3, *_sub("ARQ", "Carpintería de madera")))

    return hits


def classify_IISS(d):
    hits = []
    if re.search(r"(SUMINISTRO\s*E\s*)?INSTALACI[OÓ]N\s*DE\s*(INODORO|LAVATORIO|GRIFER[IÍ]A|LAVADERO|"
                 r"URINARIO|DUCHA|OVALINES?|MEZCLADORA)", d):
        hits.append((20, *_sub("IISS", "Instalaciones Sanitarias")))
    if re.search(r"SOMBRERO\s*(DE\s*)?VENTILACI[OÓ]N", d):
        hits.append((10, *_sub("IISS", "Instalaciones Sanitarias")))
    if re.search(r"SISTEMA\s*DE\s*PISCINA", d):
        hits.append((10, *_sub("IISS", "Instalaciones Sanitarias")))
    # rejilla de piso (sumidero/drenaje) -- distinta de la rejilla metálica
    # de inyección/extracción de aire (esa es IIMM, ver classify_IIMM).
    if re.search(r"REJILLA(S)?\s*\d*\s*\([\d.]+\s*[Xx]\s*[\d.]+\)\s*EN\s*PISO|REJILLA(S)?\s*(DE\s*)?PISO\b", d):
        hits.append((9, *_sub("IISS", "Instalaciones Sanitarias")))
    # "SPC" también es una sigla comercial de piso (Stone Plastic Composite,
    # p.ej. "PISO SPC ANTIDESLIZANTE") -- solo cuenta como bomba si aparece
    # junto a BOMBA o PRESIÓN CONSTANTE, nunca suelto.
    if re.search(r"(ELECTRO)?BOMBAS?\s*(DE\s*)?PRESI[OÓ]N\s*CONSTANTE|SISTEMA\s*(DE\s*)?PRESI[OÓ]N\s*CONSTANTE|"
                 r"(ELECTRO)?BOMBAS?.*\bSPC\b|\bSPC\b.*(ELECTRO)?BOMBAS?", d):
        hits.append((11, *_sub("IISS", "Bombas SPC")))
    if re.search(r"TUBER[IÍ]A\s*(PVC\s*)?(SAL|SAP|AGUA|DESAG[UÜ]E)|TUBER[IÍ]A\s*CPVC|TUBER[IÍ]A\s*PPR", d):
        hits.append((10, *_sub("IISS", "Instalaciones Sanitarias")))
    # tubería de agua/desagüe con material intermedio (PP, colgante/colgada,
    # SAL) entre "TUBERÍA" y el destino -- misma familia que el patrón de
    # arriba, generalizado para que no dependa de que el material venga
    # pegado a la palabra TUBERÍA.
    if re.search(r"TUBER[IÍ]A\s*(COLGAN?TE|COLGADA)?\s*(PVC|PP|PPR|CPVC)?\s*(SAL|SAP)?\s*"
                 r"(PARA\s*)?(IMPULSI[OÓ]N\s*)?(AGUA|DESAG[UÜ]E)\b", d):
        hits.append((9, *_sub("IISS", "Instalaciones Sanitarias")))
    # válvulas PPR y de bola en redes de agua (ya existe el cluster de
    # válvulas flotadora/compuerta/check/paso -- esto cubre la redacción con
    # material comercial PPR o "tipo bola", muy común en redes interiores).
    if re.search(r"V[AÁ]LVULA\s*(DE\s*)?\d.{0,15}\bPPR\b|V[AÁ]LVULA\s*TIPO\s*BOLA|GRIFO\s*DE\s*RIEGO", d):
        hits.append((9, *_sub("IISS", "Instalaciones Sanitarias")))
    # registros de inspección colgados / cajas y cajuelas de registro y
    # drenaje -- variantes de redacción del mismo elemento sanitario ya
    # cubierto arriba solo para el caso "...DE DESAGÜE/ALCANTARILLADO".
    if re.search(r"REGISTRO\s*COLGADO|CAJA\s*DE\s*REGISTRO\b|CAJUELA\s*DE\s*DRENAJE", d):
        hits.append((8, *_sub("IISS", "Instalaciones Sanitarias")))
    # pruebas de hermeticidad/estanqueidad y limpieza-desinfección de redes
    # de agua: puesta en marcha del sistema sanitario.
    if re.search(r"PRUEBA\s*(DE\s*)?ESTANQUEIDAD|LIMPIEZA\s*Y\s*DESINFECCI[OÓ]N\s*DE\s*TUBER[IÍ]AS", d):
        hits.append((9, *_sub("IISS", "Instalaciones Sanitarias")))
    if re.search(r"^SALIDA\s+(AGUA|DESAG[UÜ]E|VENTILACI[OÓ]N)\b|SALIDA\s*(DE\s*)?(AGUA|DESAG[UÜ]E)", d):
        hits.append((10, *_sub("IISS", "Instalaciones Sanitarias")))
    if re.search(r"CISTERNA|TANQUE\s*ELEVADO|TANQUE\s*DE\s*AGUA", d):
        hits.append((8, *_sub("IISS", "Instalaciones Sanitarias")))
    if re.search(r"MONTANTE\s*(DE\s*)?(AGUA|DESAG[UÜ]E)|VENTILACI[OÓ]N\s*SANITARIA|TRAMPA\s*GRASA|"
                 r"CAJA\s*DE\s*REGISTRO\s*(DE\s*)?(DESAG[UÜ]E|ALCANTARILL)|BUZ[OÓ]N\s*DE\s*DESAG[UÜ]E", d):
        hits.append((8, *_sub("IISS", "Instalaciones Sanitarias")))
    # tubería/montante DE VENTILACIÓN (ventilación de columnas de desagüe) es
    # sanitaria -- no confundir con ventilación MECÁNICA de sótanos/ductos de
    # monóxido, que es IIMM (ver classify_IIMM).
    if re.search(r"TUBER[IÍ]A\s*(DE\s*)?VENTILACI[OÓ]N\s*PVC|MONTANTE\s*(DE\s*)?VENTILACI[OÓ]N\b", d):
        hits.append((10, *_sub("IISS", "Instalaciones Sanitarias")))
    # válvulas y accesorios de redes de agua/desagüe
    if re.search(r"V[AÁ]LVULA\s*(FLOTADORA|COMPUERTA|ANTIRRETORNO|ANTIRETORNO|CHECK|DE\s*PIE|DE\s*PASO)|"
                 r"BRIDA\s*ROMPEAGUA|\bSUMIDERO\b|ESTACI[OÓ]N(ES)?\s*REDUCTORA(S)?\s*DE\s*PRESI[OÓ]N", d):
        hits.append((9, *_sub("IISS", "Instalaciones Sanitarias")))
    if re.search(r"\bSANITARI", d):
        hits.append((4, *_sub("IISS", "Instalaciones Sanitarias")))
    # registro roscado (caja de inspección de tubería), llave de paso
    # (válvula de corte, no la grifería terminal -- esa es ARQ), medidor de
    # agua, y las pruebas hidráulicas de puesta en marcha del sistema.
    if re.search(r"REGISTRO\s*ROSCADO|LLAVE\s*DE\s*PASO|MEDIDOR\s*DE\s*AGUA|"
                 r"PRUEBA(S)?\s*HIDR[AÁ]ULICA", d):
        hits.append((9, *_sub("IISS", "Instalaciones Sanitarias")))
    return hits


def classify_IIEE(d):
    # única subespecialidad real en el catálogo -> cualquier ancla eléctrica
    # fuerte basta para asignarla.
    hits = []
    if re.search(r"^SALIDAS?\s+(?!AGUA\b|DESAG[UÜ]E\b|VENTILACI[OÓ]N\b)|SALIDAS?\s*(DE\s*FUERZA\s*)?PARA\b|"
                 r"SALIDAS?\s*DE\s*FUERZA", d):
        hits.append((25, *_sub("IIEE", "Instalaciones Eléctricas")))
    if re.search(r"^TABLERO\b(?!.*(GRANITO|M[AÁ]RMOL|CUARZO|COCINA))|TABLERO\s*(GENERAL|DE\s*DISTRIBUCI[OÓ]N|"
                 r"DE\s*TRANSFERENCIA|DE\s*CONTROL|EL[EÉ]CTRIC)", d):
        hits.append((20, *_sub("IIEE", "Instalaciones Eléctricas")))
    if re.search(r"ACOMETIDA\s*EL[EÉ]CTRIC|PUNTO\s*DE\s*(LUZ|TOMACORRIENTE)|CAJA\s*DE\s*PASE", d):
        hits.append((18, *_sub("IIEE", "Instalaciones Eléctricas")))
    # tableros identificados solo por su código comercial (convención: el
    # código empieza con T -- TG=general, TN=normal, TD=distribución,
    # T-ASC/T-SUM/T-MB=tablero por uso específico) cuando la palabra
    # "TABLERO" no viene seguida de uno de los calificativos ya cubiertos
    # arriba (general, eléctrico, etc.) sino directo del código.
    if re.search(r"TABLERO\s+T[A-ZÑ]", d):
        hits.append((18, *_sub("IIEE", "Instalaciones Eléctricas")))
    # placas de interruptores/tomacorrientes y timbres -- mismo cluster de
    # accesorios eléctricos terminales que tomacorriente/interruptor.
    if re.search(r"PLACA\s*(INTERRUPTOR|PARA\s*CAMPANA|TV\b|TE\b)|\bTIMBRE\b|PULSADOR\s*TIMBRE", d):
        hits.append((8, *_sub("IIEE", "Instalaciones Eléctricas")))
    if re.search(r"\bTD[- ]?\d|POZO\s*(A|DE)\s*TIERRA|PUESTA\s*A\s*TIERRA|EQUIPOTENCIAL\s*DE\s*TIERRA|"
                 r"SOLDADURA\s*EXOT[EÉ]RMICA|BANCO\s*DE\s*MEDIDORES", d):
        hits.append((9, *_sub("IIEE", "Instalaciones Eléctricas")))
    if re.search(r"CABLE\s*(ELECTRIC|THHN|THW|NYY|AWG)|CONDUCTOR\s*EL[EÉ]CTRIC", d):
        hits.append((8, *_sub("IIEE", "Instalaciones Eléctricas")))
    if re.search(r"TOMACORRIENTE|INTERRUPTOR\s*(SIMPLE|DOBLE|CONMUTACI[OÓ]N)|CONTROL\s*REMOTO", d):
        hits.append((8, *_sub("IIEE", "Instalaciones Eléctricas")))
    if re.search(r"BREAKER|INTERRUPTOR\s*TERMOMAGN[EÉ]TIC|ITM\b|MEDIA\s*TENSI[OÓ]N|SUBESTACI[OÓ]N\s*EL[EÉ]CTRIC|"
                 r"TRANSFORMADOR", d):
        hits.append((7, *_sub("IIEE", "Instalaciones Eléctricas")))
    if re.search(r"CANALIZACI[OÓ]N\s*EL[EÉ]CTRIC|DUCTO\s*EL[EÉ]CTRIC|TUBER[IÍ]A\s*(PVC[- ]?SEL|EMT)|"
                 r"BANDEJA\s*(PORTA)?CABLE|\bEMT\b|\bCONDUIT\b", d):
        hits.append((7, *_sub("IIEE", "Instalaciones Eléctricas")))
    if re.search(r"\bEL[EÉ]CTRIC", d):
        hits.append((5, *_sub("IIEE", "Instalaciones Eléctricas")))
    # códigos de cable de potencia/control (sin la palabra "eléctrico" en la
    # descripción, típico en líneas de alimentadores "De X hasta Y: ...")
    if re.search(r"\bN2XOH\b|\bN2X0H\b|\bLSOHX\b|\bLS0HX\b|\bLSOH\b|\bTHHN\b|\bXLPE\b|"
                 r"(CABLE|CONDUCTOR)\s*DESNUDO|ATERRAMIENTO\s*DE\s*BANDEJA", d):
        hits.append((6, *_sub("IIEE", "Instalaciones Eléctricas")))
    # líneas de circuitos de alumbrado/alimentadores escritas como código técnico
    # de cable (frecuente en presupuestos que detallan cada circuito línea por
    # línea, ej. "ALUMBRADO 2-1X4MM2NH+1X4MM2NH/T-20MM∅PVC-P"): se reconoce por
    # la palabra ALUMBRADO como marcador de circuito, o por el patrón numérico
    # de especificación de cable (N-NXNMM2...) que acompaña a estas líneas,
    # incluyendo el código de aislamiento NH (no detectado por el cluster de
    # arriba porque ahí solo están N2XOH/LSOHX/LSOH/THHN/XLPE).
    if re.search(r"^ALUMBRADO\b|LUZ\s*DE\s*EMERGENCIA|\bNH\+|MM2\s*NH\b|\bNH\b.{0,15}MM2|"
                 r"\d-\d?X?\d+MM2.{0,10}(PVC|SAP)[- ]?P\b", d):
        hits.append((7, *_sub("IIEE", "Instalaciones Eléctricas")))
    # cajas de paso / pases para instalaciones eléctricas (variante de redacción
    # distinta a "CAJA DE PASE" ya cubierta arriba).
    if re.search(r"CAJA\s*DE\s*PASO\b|TIPO\s*ESCALERILLA.{0,30}ALIMENTADOR", d):
        hits.append((7, *_sub("IIEE", "Instalaciones Eléctricas")))
    return hits


def classify_IIMM(d):
    hits = []
    if re.search(r"(ELECTRO)?BOMBAS?\s*(DE\s*)?(SUMIDERO|ACHIQUE)\b|(ELECTRO)?BOMBAS?\s*(SUMERGIBLES?\s*)?"
                 r"(PARA\s*)?AGUAS\s*RESIDUALES", d):  # bombas de agua simple sin ser SPC/ACI
        hits.append((9, *_sub("IIMM", "Instalaciones Mecánicas")))
    if re.search(r"AIRE\s*ACONDICIONADO|SPLIT|VRF|CHILLER|UMA\b|UNIDAD\s*MANEJADORA", d):
        hits.append((10, *_sub("IIMM", "Instalaciones Mecánicas")))
    if re.search(r"VENTILACI[OÓ]N\s*MEC[AÁ]NICA|EXTRACTOR|INYECTOR\s*DE\s*AIRE|VENTILADOR\s*CENTR[IÍ]FUGO|"
                 r"DUCTOS?\s*(MET[AÁ]LICOS?|DE\s*VENTILACI[OÓ]N)|LIMPIEZA\s*DE\s*DUCTOS", d):
        hits.append((9, *_sub("IIMM", "Instalaciones Mecánicas")))
    # equipos/accesorios de ventilación mecánica forzada (inyección/extracción
    # de sótanos, extracción de monóxido) -- no confundir con la tubería de
    # ventilación SANITARIA de columnas de desagüe (eso es IISS, ver arriba).
    if re.search(r"\bDAMPER\b|REJILLA(S)?\s*(MET[AÁ]LICA\s*)?(DE\s*)?(INYECCI[OÓ]N|EXTRACCI[OÓ]N)|"
                 r"REJILLA(S)?\s*MET\.?\s*P/?\s*DESMONTABLE|INYECTOR(ES)?\s*(AXIAL|CENTR[IÍ]FUGO|HELICOCENTR[IÍ]FUGO)|"
                 r"EXTRACTOR(ES)?\s*CENTR[IÍ]FUGO|EXTRACCI[OÓ]N\s*DE\s*MON[OÓ]XIDO|"
                 r"VENTILADOR(ES)?\s*(AXIAL|HELICOCENTR[IÍ]FUGO)|JET\s*FAN|"
                 r"DUCTO(S)?\s*MET[AÁ]LICO(S)?\s*DE\s*FIERRO\s*GALVANIZADO|FILTRO\s*DE\s*MALLA|"
                 r"ACCESORIOS\s*DE\s*VENTILACI[OÓ]N", d):
        hits.append((11, *_sub("IIMM", "Instalaciones Mecánicas")))
    # equipos de ventilación mecánica identificados solo por su código de
    # tablero (VA, EA, EC, VC, JF...) y su especificación técnica en CFM --
    # el caudal de aire en CFM es el marcador universal de este tipo de
    # equipo en cualquier proyecto, incluso sin una palabra en español que
    # lo acompañe.
    if re.search(r"\bCFM\b", d):
        hits.append((9, *_sub("IIMM", "Instalaciones Mecánicas")))
    return hits


def classify_GAS(d):
    hits = []
    if re.search(r"RED\s*DE\s*GAS|TUBER[IÍ]A\s*DE\s*GAS|GAS\s*NATURAL|GLP\b|MEDIDOR\s*DE\s*GAS|V[AÁ]LVULA\s*DE\s*GAS|"
                 r"SISTEMA\s*DE\s*GAS|NICHOS?\s*PARA\s*V[AÁ]LVULAS?\s*DE\s*GAS|"
                 r"REGULADOR.{0,10}MAN[OÓ]METRO|MAN[OÓ]METRO.{0,10}REGULADOR", d):
        hits.append((9, *_sub("GAS", "Instalaciones de Gas")))
    # tubería de cobre (Cu): en el catálogo real de esta empresa, la tubería
    # de cobre se usa exclusivamente para redes de GAS (el agua usa PVC/PPR/
    # CPVC) -- comprobado comparando el total de la hoja "IIGG" contra la
    # línea "SISTEMA DE GAS" del RESUMEN en dos proyectos distintos, cuadra
    # exacto. Red interna de cobre/PEALPE y accesorios asociados (nichos
    # para cocina/therma -- ambos artefactos a gas, válvulas con acabado
    # cromado/mariposa de corte de gas) van aquí, no en IISS.
    if re.search(r"TUBER[IÍ]A\s*CU\b|TUBER[IÍ]A\s*CU\.|RED\s*(INTERNA\s*)?DE\s*COBRE|\bPEALPE\b|"
                 r"NICHO.{0,15}(COCINA|THERMA)|V[AÁ]LVULA.{0,15}(CROMADA|MARIPOSA)", d):
        hits.append((7, *_sub("GAS", "Instalaciones de Gas")))
    return hits


def classify_ACI(d):
    hits = []
    if re.search(r"BOMBAS?\s*(CONTRA\s*INCENDIO|JOCKEY|ACI\b)", d):
        hits.append((10, *_sub("ACI", "Bomba ACI")))
    if re.search(r"ROCIADOR|SPRINKLER|RED\s*H[UÚ]MEDA|RED\s*SECA|GABINETE\s*CONTRA\s*INCENDIO|"
                 r"CONTRAINCENDIO|CONTRA\s*INCENDIO|EXTINTOR", d):
        hits.append((9, *_sub("ACI", "Red ACI")))
    return hits


def classify_CCDEBILES(d):
    hits = []
    if re.search(r"DETECTOR\s*DE\s*HUMO|ALARMA\s*CONTRA\s*INCENDIO\b(?!.*ROCIADOR)|PANEL\s*CONTRA\s*INCENDIO|"
                 r"SENSOR\s*DE\s*(HUMO|ANIEGO)|SIRENA\s*(CON\s*LUZ\s*)?ESTROBOSC[OÓ]PICA", d):
        hits.append((9, *_sub("CCDÉBILES", "DACI")))
    if re.search(r"C[AÁ]MARA(S)?\s*(DE\s*)?(SEGURIDAD|VIGILANCIA|CCTV|VIDEO)|CIRCUITO\s*CERRADO|"
                 r"\bDVR\b|VIDEOVIGILANCIA", d):
        hits.append((9, *_sub("CCDÉBILES", "CCTV")))
    if re.search(r"CITOFON|INTERCOMUNICADOR|VIDEOPORTERO", d):
        hits.append((9, *_sub("CCDÉBILES", "Intercomunicadores")))
    # sistema de portero/conserjería (audio-video) y sus componentes de
    # control -- vocabulario repetido en varios proyectos dentro de una hoja
    # genérica "INSTALACIONES" (por eso no lo agarra el fallback por hoja):
    # portero alfanumérico, central de conserjería, placa de calle,
    # interfaz/tarjeta expansora, configuradores -- todos accesorios del
    # mismo sistema de comunicación del edificio.
    if re.search(r"PORTERO\s*ALFANUM[EÉ]RICO|AUDIOPORTERO|CENTRAL(ITA)?\s*DE\s*CONSERJER[IÍ]A|PLACA\s*DE\s*CALLE|"
                 r"INTERFAZ\s*DE\s*EXPANSI[OÓ]N|TARJETA\s*EXPANSORA|CONFIGURADOR(ES)?\s*(DE\s*TEL[EÉ]FONOS)?\b|"
                 r"\bNEWSFERA\b|M[OÓ]DULO\s*(DE\s*)?(MONITOREO|SE[ÑN]AL|EXPANSOR|AISLAMIENTO|CIEGO)|"
                 r"ALIMENTADOR(ES)?\s*(ADICIONAL\s*)?\d*\s*HILOS|PLACA\s*TECNOPOL[IÍ]MERO|"
                 r"SOPORTE\s*DE\s*RESINA\s*PARA\s*PLACAS", d):
        hits.append((9, *_sub("CCDÉBILES", "Intercomunicadores")))

    # electrónica/accesorios de un sistema de seguridad/CCTV genérico, dentro
    # de una hoja "INSTALACIONES" sin más contexto -- cámaras por tipo de
    # cuerpo, grabación (disco duro), monitor de visualización, cableado y
    # alimentación propios de ese sistema.
    if re.search(r"C[AÁ]MARAS?\s*(TUBO|MINI\s*DOMO|DOMO)\b|DISCO\s*DURO\s*(DE\s*)?\d+\s*TB|"
                 r"MONITOR\s*(A\s*COLOR\s*)?LED|CABLE\s*UTP|CABLE\s*8H\b|"
                 r"BATER[IÍ]A\s*(DE\s*)?\d+\s*V\s*/?\s*\d*\s*AH|FUENTES?\s*DE\s*ALIMENTACI[OÓ]N\s*DE\s*\d+\s*V|"
                 r"GABINETE\s*MET[AÁ]LICO\s*DE\s*\d+\s*RU|TOMA\s*\d+\s*CONTACTOS", d):
        hits.append((9, *_sub("CCDÉBILES", "CCTV")))
    # detector de temperatura (variante del detector de humo/aniego, mismo
    # sistema de detección y alarma contra incendio) y cableado de datos
    # dedicado a estos sistemas de comunicación/seguridad.
    if re.search(r"DETECTOR\s*DE\s*TEMPERATURA|CABLE\s*DE\s*COMUNICACIONES\s*UTP|"
                 r"ESTACI[OÓ]N\s*MANUAL(\s*(DE\s*)?ACCI[OÓ]N\s*DOBLE)?|BOOSTER\s*NACS?|"
                 r"FUENTE\s*(DE\s*)?ALIMENTACI[OÓ]N.{0,20}SIRENAS?", d):
        hits.append((8, *_sub("CCDÉBILES", "DACI")))
    # bandeja/escalerilla de cableado para RED DE COMUNICACIONES (voz, datos,
    # intercomunicador) -- misma convención de "TIPO ESCALERILLA" ya usada en
    # IIEE para alimentadores eléctricos, aquí es la variante de datos.
    if re.search(r"TIPO\s*ESCALERILLA.{0,30}COMUNICACIONES", d):
        hits.append((7, *_sub("CCDÉBILES", "Intercomunicadores")))
    # automatización/domótica de departamentos -- sistema de control
    # inteligente del edificio, misma familia de instalaciones débiles. No
    # hay subespecialidad exacta en el catálogo para esto -- se deja para
    # revisión manual del código, pero ya cae en la especialidad correcta.
    if re.search(r"AUTOMATIZACI[OÓ]N\s*DE\s*(DEPARTAMENTOS|VIVIENDAS)|\bDOM[OÓ]TICA\b", d):
        hits.append((9, "Revisar manualmente", None))
    return hits


def classify_GG(d):
    """Gastos Generales: planilla de staff de obra, seguros, oficina técnica,
    equipamiento administrativo, trámites y gastos financieros del propio
    contrato -- en la mayoría de proyectos esto es una sola línea global de
    % en el RESUMEN (por eso GG está en NON_ITEM_SPECIALTIES y sus hojas se
    excluyen por defecto), pero algunos presupuestos SÍ lo desglosan partida
    por partida en hojas propias (ej. "GG_E1", "GG_PILOTAJE") -- cuando eso
    pasa, estas son las partidas reales que hay que reconocer ahí."""
    hits = []
    if re.search(r"\b(JEFE|GERENTE|INGENIERO|CAPATAZ|ASISTENTE|SUPERVISOR|RESIDENTE|SUPERINTENDENTE|"
                 r"PREVENCIONISTA|ADMINISTRADOR)\s*DE\s*(OFICINA\s*T[EÉ]CNICA|OBRA|CAMPO|SEGURIDAD|"
                 r"CALIDAD|ALMAC[EÉ]N|PRODUCCI[OÓ]N)|"
                 r"\bINGENIERO\s*RESIDENTE\b|\bSUPERINTENDENTE\s*DE\s*OBRAS?\b|"
                 r"\bVIGILANTE\b|\bCAPATAZ\b|\bGUARDI[AÁ]N\s*DE\s*OBRA\b|\bPREVENCIONISTA\b", d):
        hits.append((9, *_sub("GG", "Gastos Generales")))
    if re.search(r"SCTR\s*\(|SEGURO\s*DE\s*VIDA\s*LEY|PERSONAL\s*(DE\s*)?(LIMPIEZA|OBRERO)\b|"
                 r"P[OÓ]LIZA\s*(CAR|3D|DE\s*)|PAZ\s*SOCIAL|SENCICO", d):
        hits.append((9, *_sub("GG", "Gastos Generales")))
    if re.search(r"[UÚ]TILES\s*DE\s*OFICINA|ECONOMATO|CONTRIBUCI[OÓ]N\s*OFICINA\s*CENTRAL|"
                 r"LICENCIA\s*(DE\s*)?(SOFTWARE|S10|MS\s*OFFICE|AUTOCAD)|GASTOS?\s*ADMINISTRATIVOS?|"
                 r"FOTOCOPIADO|COPIA\s*DE\s*PLANOS|COMUNICACI[OÓ]N\s*TEL[EÉ]FONO|BOT[IÍ]QU[IÍ]N\s*DE\s*OBRA|"
                 r"GASTOS?\s*FINANCIEROS?|COSTO\s*FINANCIERO\s*DE\s*CAPITAL\s*DE\s*TRABAJO|\bITF\b", d):
        hits.append((8, *_sub("GG", "Gastos Generales")))
    return hits


def classify_EQUIP(d):
    # EQUIP en el catálogo real son SOLO estos 4 sistemas mayores.
    hits = []
    if re.search(r"GRUPO\s*ELECTR[OÓ]GENO", d):
        hits.append((23, *_sub("EQUIP", "Grupo Electrógeno")))
    if re.search(r"ACELER[OÓ]GRAFO", d):
        hits.append((23, *_sub("EQUIP", "Acelerógrafo")))
    if re.search(r"PANEL(ES)?\s*(FOTOVOLTAICO|SOLAR)", d):
        hits.append((23, *_sub("EQUIP", "Paneles Solares")))
    if re.search(r"EQUIPAMIENTO\s*(DE\s*)?PISCINA|CLORINADOR", d):
        hits.append((23, *_sub("EQUIP", "Equipamiento Piscina")))
    return hits


def classify_ASCENSORES(d):
    hits = []
    if re.search(r"ASCENSOR(?!.*\b(VIGA|COLUMNA|OBRA\s*CIVIL|IMPERMEABILIZACI[OÓ]N|ENCOFRADO|CONCRETO|"
                 r"SOLAQUEO|EMPASTAD|PINTURA|PISO|CZ|ZOCALO|Z[OÓ]CALO|PORCELANATO|CER[AÁ]MIC|"
                 r"CEMENTO\s*PULIDO|TABLERO|SALIDA|TABIQUE|LADRILLO|CORTAFUEGO)\b)", d):
        hits.append((10, *_sub("ASCENSORES", "Ascensores")))
    if re.search(r"CABINA\s*DE\s*ASCENSOR|M[AÁ]QUINA\s*DE\s*ASCENSOR|MONTACARGA", d):
        hits.append((8, *_sub("ASCENSORES", "Ascensores")))
    return hits


CLASIFICADORES = {
    "OP": classify_OP,
    "ESTT": classify_ESTT,
    "ARQ": classify_ARQ,
    "IISS": classify_IISS,
    "IIEE": classify_IIEE,
    "IIMM": classify_IIMM,
    "GAS": classify_GAS,
    "ACI": classify_ACI,
    "CCDÉBILES": classify_CCDEBILES,
    "EQUIP": classify_EQUIP,
    "ASCENSORES": classify_ASCENSORES,
    "GG": classify_GG,
}


def score_all(desc):
    """{especialidad: {"score":, "subespecialidad":, "codigo":}} para una
    descripción dada, evaluando TODAS las especialidades."""
    d = U(desc)
    out = {}
    for especialidad, fn in CLASIFICADORES.items():
        hits = fn(d)
        if not hits:
            continue
        total = sum(h[0] for h in hits)
        mejor = max(hits, key=lambda h: h[0])
        out[especialidad] = {"score": total, "subespecialidad": mejor[1], "codigo": mejor[2]}
    return out


def best_match(desc):
    """(especialidad, subespecialidad, codigo, score) de mejor coincidencia,
    o (None, None, None, 0) si no hay ningún match."""
    scores = score_all(desc)
    if not scores:
        return None, None, None, 0
    especialidad = max(scores, key=lambda k: scores[k]["score"])
    info = scores[especialidad]
    return especialidad, info["subespecialidad"], info["codigo"], info["score"]
