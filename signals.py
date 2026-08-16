"""
signals.py
==========
Replica EXACTA (celda por celda) de las formulas del Excel
"AmbosMarcan-BTTS - Python.xlsx" (hoja "Partidos", columnas AB-AT).

Cada funcion de aqui corresponde a una columna del Excel original. Los
nombres de columna de entrada usan las mismas letras que el Excel para que
sea trivial auditar contra el archivo fuente:

  E  = xG Local            F  = xGA Local
  G  = xG Visitante         H  = xGA Visitante
  I  = GF Local             J  = GC Local
  K  = GF Visitante         L  = GC Visitante
  M  = Failed Local %       N  = CS Local %
  O  = Failed Visit %       P  = CS Visit %
  Q  = SoT Local            R  = SoT Visitante
  S  = Cuota BTTS           T  = Cuota O1.5        U  = Cuota O2.5
  V  = BTTS% Local          W  = BTTS% Visitante
  X  = Cuota Local          Y  = Cuota Visitante

NO se modifico ningun umbral ni condicion: si crees que una senal esta
"mal", el problema esta en el Excel original, no aqui. Cualquier cambio de
regla debe hacerse conscientemente y documentarse.
"""
from __future__ import annotations
from dataclasses import dataclass
from typing import Optional
import math
import config
import reglas_activas

# Version de la formula - subela cada vez que cambies un umbral/condicion en
# este archivo. Se guarda en cada partido nuevo (historial_predicciones.csv,
# columna formula_version) para poder despues aislar EXACTAMENTE que
# partidos tienen todos los ajustes de una fecha dada aplicados, en vez de
# adivinar por fecha de deteccion. Sube este numero en cada sesion de
# cambios (no hace falta un esquema complejo, un entero simple basta).
# Calibracion empirica de la probabilidad de BTTS, ajustada con 1705
# partidos reales (8 ago 2026). El modelo resulto MUY sobreconfiado: abria
# un abanico de 47 puntos (dijimos entre 27% y 74%) mientras la realidad
# solo se movia 15 puntos (44% a 59%). La recta ajustada fue:
#     real = 32.9 + 0.355 * dicha
# La pendiente de 0.355 significa que subir 10 puntos nuestra probabilidad
# solo corresponde a 3.5 puntos reales.
#
# IMPORTANTE: esta calibracion se aplica SOLO al numero que se muestra
# (prob_btts_calibrada), NO a la probabilidad interna que usan los
# umbrales de las formulas (AD>=0.62 etc.) - si se calibrara AD, ningun
# partido volveria a superar 0.62 nunca (el maximo calibrado es ~0.59) y
# todas las señales dejarian de dispararse.
#
# Advertencia: la recta se ajusto sobre los mismos datos con los que se
# midio, asi que el ajuste real hacia adelante sera algo peor. La
# DIRECCION (sobreconfianza fuerte en los extremos) si es robusta: el
# patron es monotono y consistente en los 1705 partidos.
# Ajuste por porteria en cero (ver comentario extenso donde se aplica).
# CS_MAX_PROMEDIO: promedio historico real del maximo de clean sheet% entre
# los dos equipos (32.3% con n=1293). El ajuste se centra en este valor
# para no mover el promedio general de la probabilidad.
# Castigo por correlacion en la combinada BTTS + Over 2.5.
#
# PROBLEMA que corrige: multiplicar cuota_btts x cuota_o25 asume que los
# dos eventos son INDEPENDIENTES, y no lo son ni de cerca. Medido con
# datos reales propios: P(Over 2.5) es 52.9% en general, pero sube a
# 79.8% cuando ya se dio el BTTS (logico: si ambos marcan ya hay 2 goles).
# Las casas de apuestas lo saben y pagan bastante menos por la combinada
# de lo que da la multiplicacion - a esto se le llama "correlation
# discount" o descuento de bet builder / same game multi.
#
# El rango documentado para BTTS + Over/Under especificamente es 20-35%
# (fuente: gamblingcalc.com, tabla de correlacion por tipo de combinada).
# Se usa el punto medio 27.5% por defecto.
#
# Formula (el descuento aplica a la GANANCIA, no a la cuota completa):
#     cuota_real = 1 + (cuota_producto - 1) * (1 - descuento)
#
# Impacto medido en el historico propio (n=1461): el ROI de esta señal
# pasa de un ficticio +3.7% a un reali -11.1%. Ese +3.7% que veniamos
# midiendo era enteramente un artefacto de usar cuotas inexistentes.
DESCUENTO_CORRELACION_COMBO = 0.275

CS_MAX_PROMEDIO = 32.3     # promedio historico real (n=1293)
XG_DEBIL_PROMEDIO = 1.012  # xG del equipo mas flojo / no favorito
SOT_DEBIL_PROMEDIO = 3.530 # tiros a puerta del visitante / no favorito

# Coeficientes del modelo logistico, en escala logit por unidad real de
# cada variable (ya des-normalizados). Signo negativo = baja la
# probabilidad de BTTS; positivo = la sube.
COEF_CS_MAX = -0.011285    # mas porteria en cero -> menos BTTS
COEF_XG_DEBIL = -0.065143  # (ver comentario donde se aplica)
COEF_SOT_DEBIL = +0.087058 # mas tiros del que necesita marcar -> mas BTTS

CALIBRACION_BTTS_INTERCEPTO = 0.329
CALIBRACION_BTTS_PENDIENTE = 0.355

FORMULA_VERSION = 4  # v1 = original; v2 = sesion 6 ago (muchos cambios distintos
# mezclados sin volver a subir la version - leccion aprendida); v3 = 13 ago,
# modelo multivariable de BTTS aplicado; v4 = 13 ago (limpieza final) - se
# elimino Paridad Ofensiva y los filtros de favorito en VIP/Fuerte tras
# descubrir que el analisis que los respaldaba tenia un bug de cruce de
# datos (perdida de filas por emparejamiento de nombres); VIP y Fuerte
# vuelven a su forma validada (solo el umbral de P).
# (P thresholds VIP/Fuerte, Stake10, mezcla con mercado, favorito/no-favorito
# en VIP y Fuerte, formula intermedia VIP, minimo partidos ORO/PLATA, edge
# positivo en rescates y VIP/Fuerte)


def _num(v, default=0.0):
    """Equivalente a IFERROR(v, default): si v es None/NaN, devuelve default."""
    if v is None:
        return default
    try:
        f = float(v)
        if math.isnan(f):
            return default
        return f
    except (TypeError, ValueError):
        return default


def _pct(v, default):
    """
    Equivalente a IF(v>1, v/100, IFERROR(v, default)).
    Los porcentajes en el Excel a veces vienen como 29 (=29%) y a veces
    como 0.29 - esta funcion normaliza ambos casos a fraccion (0-1).
    """
    if v is None:
        return default
    try:
        f = float(v)
        if math.isnan(f):
            return default
    except (TypeError, ValueError):
        return default
    return f / 100.0 if f > 1 else f


def _blank(v) -> bool:
    """Equivalente a v="" en Excel: None, NaN o cadena vacia."""
    if v is None:
        return True
    try:
        return math.isnan(v)
    except TypeError:
        return v == ""


@dataclass
class MatchInput:
    """Todas las columnas de entrada manuales/automaticas para UN partido."""
    xg_local: Optional[float]           # E
    xga_local: Optional[float]          # F
    xg_visitante: Optional[float]       # G
    xga_visitante: Optional[float]      # H
    gf_local: Optional[float]           # I
    gc_local: Optional[float]           # J
    gf_visitante: Optional[float]       # K
    gc_visitante: Optional[float]       # L
    failed_local_pct: Optional[float]   # M
    cs_local_pct: Optional[float]       # N
    failed_visit_pct: Optional[float]   # O
    cs_visit_pct: Optional[float]       # P
    sot_local: Optional[float]          # Q
    sot_visitante: Optional[float]      # R
    cuota_btts: Optional[float]         # S
    cuota_o15: Optional[float]          # T
    cuota_o25: Optional[float]          # U
    btts_pct_local: Optional[float]     # V
    btts_pct_visitante: Optional[float] # W
    cuota_local: Optional[float] = None       # X
    cuota_visitante: Optional[float] = None   # Y
    cuota_btts_no: Optional[float] = None     # cuota de "No BTTS" - solo para quitarle el margen
    # --- Primera mitad (Over 0.5 goles HT) - mercado nuevo, agosto 2026 ---
    scored_ht_local: Optional[float] = None
    conceded_ht_local: Optional[float] = None
    scored_ht_visitante: Optional[float] = None
    conceded_ht_visitante: Optional[float] = None
    over05_ht_pct_local: Optional[float] = None
    over05_ht_pct_visitante: Optional[float] = None
    cuota_ht_over05: Optional[float] = None
    # Partidos jugados (para exigir minimo de muestra en ORO_BTTS/PLATA)
    pj_local: Optional[int] = None
    pj_visitante: Optional[int] = None
    reglas_activas: Optional[dict] = None  # que reglas pueden generar JUGAR (dinamico, ver reglas_activas.py)
    # --- Candidatas nuevas SIN VALIDAR todavia (13 ago 2026) - solo se
    # capturan y se guardan en el historial para poder medir su ROI real
    # en unas semanas, cuando haya suficiente muestra. NO se usan en
    # ninguna formula de decision (JUGAR/EVITAR) por ahora.
    over25_and_btts_pct_local: Optional[float] = None
    over25_and_btts_pct_visitante: Optional[float] = None
    ppg_local: Optional[float] = None
    ppg_visitante: Optional[float] = None
    shots_local: Optional[float] = None
    shots_visitante: Optional[float] = None
    avght_local: Optional[float] = None
    avght_visitante: Optional[float] = None
    avg2h_local: Optional[float] = None
    avg2h_visitante: Optional[float] = None


@dataclass
class MatchSignals:
    lambda_local: Optional[float]        # AB
    lambda_visitante: Optional[float]    # AC
    prob_btts: Optional[float]           # AD (cruda, la que usan los umbrales internos)
    prob_btts_calibrada: Optional[float]  # AD ajustada a la realidad - es la que hay que MOSTRAR
    mu_total: Optional[float]            # AE
    p_over15: Optional[float]            # AF
    p_over25: Optional[float]            # AG
    senal_btts_optimizada: str           # AH
    nivel_final: str                     # AI
    nivel_final_motivo: str              # cual regla exacta activo el JUGAR/EVITAR de AI
    edge_btts: Optional[float]           # AJ
    senal_o15: str                       # AK
    sug_o15: str                         # AL
    edge_o15: Optional[float]            # AM
    senal_o25: str                       # AN
    sug_o25: str                         # AO
    motivo_o25: str                      # por que regla se activo (o no) el JUGAR de O2.5
    edge_o25: Optional[float]            # AP
    vip_elite: str                       # AQ
    senal_btts_o25: str                  # AR
    oro: str                             # AS
    nivel_confianza: str                 # AT  (ELITE / PLATA / BRONCE)
    stake10_prob_estimada: Optional[float] = None  # diagnostico: am_sin_externos
    stake10_cumple: Optional[bool] = None           # diagnostico: si cumplio la condicion
    # --- Primera mitad (Over 0.5 goles HT) ---
    prob_over05_ht: Optional[float] = None
    prob_over15_ht: Optional[float] = None  # nota informativa, no es señal propia
    senal_ht05: str = ""
    sug_ht05: str = ""
    motivo_ht05: str = ""
    edge_ht05: Optional[float] = None
    formula_version: int = FORMULA_VERSION
    prob_combo_calibrada: Optional[float] = None
    btts_min_alto: Optional[bool] = None
    motivo_combo: str = "ninguno_evitar_combo"

def _excel_round(x: float) -> int:
    """
    Replica ROUND(x,0) de Excel: redondeo 'half away from zero' (0.5 sube),
    distinto del round() nativo de Python que usa 'banker's rounding'
    (redondea al par mas cercano en los .5 exactos).
    """
    return math.floor(x + 0.5) if x >= 0 else math.ceil(x - 0.5)


def calculate_marcador_exacto(m: MatchInput) -> dict | None:
    """
    Replica EXACTA de la formula de la columna AW ("Marcador Exacto
    Sugerido") de la hoja 'AmbosMarcan-BTTS-Mejorado' del Excel. Usa un
    modelo de lambda distinto ('v2', regresion lineal propia) al que usa
    calculate_signals() para BTTS/Over - son dos modelos separados en el
    Excel original, y se replican por separado aqui.

    Formulas originales (columnas K,L,M,N,O,P,Q,R de esa hoja):
      lambda_local_v2     = 0.3055*xg_local + 0.388*xga_visitante
                             + 0.2395*gf_local + 0.1956*gc_visitante - 0.2848
      lambda_visitante_v2 = 0.3055*xg_visitante + 0.388*xga_local
                             + 0.2395*gf_visitante + 0.1956*gc_local - 0.3231
      marcador_sugerido   = clamp(0,5, ROUND(lambda_local_v2)) & "-" &
                             clamp(0,5, ROUND(lambda_visitante_v2))

    Devuelve None si falta algun dato de entrada (igual que el Excel
    quedaria en blanco).
    """
    campos = [m.xg_local, m.xga_local, m.xg_visitante, m.xga_visitante,
              m.gf_local, m.gc_local, m.gf_visitante, m.gc_visitante]
    if any(c is None for c in campos):
        return None

    lambda_local_v2 = (0.3055 * m.xg_local + 0.388 * m.xga_visitante
                        + 0.2395 * m.gf_local + 0.1956 * m.gc_visitante - 0.2848)
    lambda_visitante_v2 = (0.3055 * m.xg_visitante + 0.388 * m.xga_local
                            + 0.2395 * m.gf_visitante + 0.1956 * m.gc_local - 0.3231)

    goles_local_sug = max(0, min(5, _excel_round(lambda_local_v2)))
    goles_visit_sug = max(0, min(5, _excel_round(lambda_visitante_v2)))

    return {
        "lambda_local_v2": lambda_local_v2,
        "lambda_visitante_v2": lambda_visitante_v2,
        "marcador_sugerido": f"{goles_local_sug}-{goles_visit_sug}",
    }


def calculate_signals(m: MatchInput) -> MatchSignals:
    E, F, G, H = m.xg_local, m.xga_local, m.xg_visitante, m.xga_visitante
    I, J, K, L = m.gf_local, m.gc_local, m.gf_visitante, m.gc_visitante
    M, N, O, P = m.failed_local_pct, m.cs_local_pct, m.failed_visit_pct, m.cs_visit_pct
    Q, R = m.sot_local, m.sot_visitante
    S, T, U = m.cuota_btts, m.cuota_o15, m.cuota_o25
    S_no = m.cuota_btts_no
    V, W = m.btts_pct_local, m.btts_pct_visitante
    X, Y = m.cuota_local, m.cuota_visitante

    # --- AB: lambda Local ---
    if _blank(E) or _blank(H) or _blank(I) or _blank(L):
        AB = None
    else:
        AB = 0.62 * ((_num(E, 0) + _num(H, 0)) / 2) + 0.45 * ((_num(I, 0) + _num(L, 0)) / 2)

    # --- AC: lambda Visitante ---
    if _blank(G) or _blank(F) or _blank(K) or _blank(J):
        AC = None
    else:
        AC = 0.62 * ((_num(G, 0) + _num(F, 0)) / 2) + 0.45 * ((_num(K, 0) + _num(J, 0)) / 2)

    # --- AD: Probabilidad BTTS (Poisson, P(local>=1) * P(visit>=1)) ---
    if AB is None or AC is None:
        AD = None
    else:
        AD = (1 - math.exp(-AB)) * (1 - math.exp(-AC))

        # --- Ajuste MULTIVARIABLE de la probabilidad de BTTS ---
        # Modelo de regresion logistica ajustado con 1293 partidos reales
        # de produccion (8 ago 2026), validado con validacion cruzada de 5
        # pliegues (no sobre los mismos datos con que se ajusto).
        #
        # Resultados de la busqueda (AUC = poder de separar ganados de
        # perdidos; 0.50 = no sirve de nada):
        #   Formula Poisson sola (lo que habia)      -> 0.550
        #   cs_max sola                               -> 0.559
        #   3 variables nuevas juntas                 -> 0.567
        #   Poisson + las 3 nuevas (esto)             -> 0.571
        #   18 variables (probado y DESCARTADO)       -> 0.534  <- sobreajuste
        #
        # Se probaron todas las combinaciones de 1 a 4 variables; el optimo
        # fueron 3 (con 4 el AUC ya empieza a BAJAR por sobreajuste). Las
        # elegidas y por que aportan:
        #   cs_max        (-) el % de porteria en cero del equipo que MAS la
        #                     mantiene: si cualquiera de los dos suele dejar
        #                     el arco en cero, es menos probable el BTTS
        #   xg_min        (-) el xG del equipo mas flojo en ataque
        #   sot_visitante (+) tiros a puerta del visitante - normalmente el
        #                     no-favorito, y es justo el que necesita marcar
        #                     para que se de el BTTS
        #
        # Cuando hay cuota 1X2 disponible se usan el xG y los tiros del
        # NO FAVORITO real en vez del visitante (confirmado que mejora:
        # AUC 0.579 vs 0.572 en los 518 partidos donde hay cuota).
        #
        # El ajuste se aplica CENTRADO en los promedios historicos, para no
        # mover el promedio general de la probabilidad y no romper los
        # umbrales existentes (AD>=0.62 etc.).
        cs_max_v = max(_pct(N, 0), _pct(P, 0)) * 100 if not (_blank(N) and _blank(P)) else CS_MAX_PROMEDIO

        # Si hay cuota 1X2, usar los datos del NO favorito (el que necesita
        # marcar); si no, usar el visitante y el xG mas bajo como respaldo.
        if not (_blank(X) or _blank(Y) or _num(X, 0) == _num(Y, 0)):
            if _num(X, 0) < _num(Y, 0):   # local favorito -> no favorito = visitante
                xg_debil_v, sot_debil_v = _num(G, 0), _num(R, 0)
            else:                          # visitante favorito -> no favorito = local
                xg_debil_v, sot_debil_v = _num(E, 0), _num(Q, 0)
        else:
            xg_debil_v = min(_num(E, 0), _num(G, 0))
            sot_debil_v = _num(R, 0)

        ajuste_logit = (
            COEF_CS_MAX * (cs_max_v - CS_MAX_PROMEDIO)
            + COEF_XG_DEBIL * (xg_debil_v - XG_DEBIL_PROMEDIO)
            + COEF_SOT_DEBIL * (sot_debil_v - SOT_DEBIL_PROMEDIO)
        )
        # Se aplica en escala logit (no multiplicando) para que el ajuste
        # respete los limites 0-1 de forma natural y sea mas fuerte en el
        # medio que en los extremos, como corresponde a una probabilidad.
        AD_clamp = min(max(AD, 1e-6), 1 - 1e-6)
        logit = math.log(AD_clamp / (1 - AD_clamp)) + ajuste_logit
        AD = 1 / (1 + math.exp(-logit))

    # Calibracion con el mercado: se confirmo con datos reales (backtest
    # propio Y el Excel de 1000+ partidos) que nuestra probabilidad de BTTS
    # sale mal calibrada con datos automatizados (plana, 43-57% sin importar
    # que digamos), mientras que la cuota del mercado predice mejor. En vez
    # de reemplazar nuestra formula, se mezcla 50/50 con la probabilidad del
    # mercado YA SIN el margen de la casa de apuestas (se quita usando la
    # cuota de "Si" y "No" juntas - practica estandar en la industria,
    # conocida como "quitar el vig"). Respaldado en literatura academica:
    # optimizar por calibracion (no solo por acierto) genero 69.86% mas
    # retorno en un estudio (Walsh & Joshi, 2024).
    if config.BTTS_MEZCLAR_CON_MERCADO and AD is not None and S and S_no and S > 0 and S_no > 0:
        prob_implicita_si = 1 / S
        prob_implicita_no = 1 / S_no
        margen_total = prob_implicita_si + prob_implicita_no  # >1.0 por el margen de la casa
        prob_mercado_limpia = prob_implicita_si / margen_total  # ya sin margen, suma 1.0 con la de "No"
        AD = 0.5 * AD + 0.5 * prob_mercado_limpia

    # NOTA (13 ago 2026): aqui existia una calibracion lineal continua
    # (real = 32.9 + 0.355*dicha) que se reemplazo por una version anclada
    # a los grupos reales validados (ver mas abajo, donde se calcula
    # AD_calibrada justo despues de AI_motivo) - esa version SI garantiza
    # coherencia con la cuota de equilibrio para cada grupo especifico,
    # cosa que la formula lineal continua no podia prometer con honestidad.

    # --- AE: mu Total ---
    AE = None if (AB is None or AC is None) else AB + AC

    # --- AF: P(Over 1.5) = 1 - P(0) - P(1) ---
    AF = None if AE is None else 1 - math.exp(-AE) * (1 + AE)

    # --- AG: P(Over 2.5) = 1 - P(0) - P(1) - P(2) ---
    AG = None if AE is None else 1 - math.exp(-AE) * (1 + AE + (AE ** 2) / 2)

    # --- AH: SEÑAL BTTS Optimizada ---
    if AD is None:
        AH = ""
    else:
        suma_xg = _num(E, 0) + _num(G, 0)
        umbral_fallo = 0.25 if suma_xg >= 3.1 else 0.2

        cond_m = True if _blank(M) else (_pct(M, 0) <= umbral_fallo)
        cond_o = True if _blank(O) else (_pct(O, 0) <= umbral_fallo)
        cond_n = True if _blank(N) else (_pct(N, 0) <= 0.4)
        cond_tiros = True if (_blank(Q) or _blank(R)) else (_num(Q, 0) + _num(R, 0) >= 10)

        # --- Stake10: de la hoja "AmbosMarcan-BTTS-Mejorado" ---
        # Formula original combinaba variables de FootyStats CON 3 sitios
        # externos (Betmines, Forebet, Statarea) que no tenemos disponibles
        # en el sistema automatizado. Se quitaron esos 3 terminos y se
        # revalido contra el historico real: 79.6% de acierto (n=93) SIN
        # ellos, prueba de que las variables propias ya bastan.
        # LIMITACION conocida: la formula original usa BTTS "ultimos 10
        # partidos" especificamente (V/W aqui son el promedio de TODA la
        # temporada, la unica granularidad que FootyStats nos da) - es la
        # mejor aproximacion posible, pero no es identica al original.
        AG_n = min(1, max(0, _num(H, 0) / 2.5))   # xga_visitante normalizado
        AH_n = min(1, max(0, _num(G, 0) / 2.5))   # xg_visitante normalizado
        AI_n = min(1, max(0, _num(I, 0) / 2.5))   # gf_local normalizado
        AJ_n = min(1, max(0, _num(K, 0) / 2.5))   # gf_visitante normalizado
        AK_n = _pct(W, 0)  # btts_pct_visitante (proxy de "ultimos 10")
        AL_n = _pct(V, 0)  # btts_pct_local (proxy de "historico")

        am_sin_externos = max(0, min(0.73, 1 / S if S else 0)) - 0.1
        am_sin_externos += (0.09 if _num(F, 0) >= 1.9 else 0.07 if _num(F, 0) >= 1.6
                             else 0.05 if _num(F, 0) >= 1.4 else 0.03 if _num(F, 0) >= 1.2
                             else (-0.03 if _num(F, 0) < 1 else 0))
        # BUG CORREGIDO (13 ago 2026): este termino usa la columna K del
        # Excel ("xG Local (anota)" = xg_local, nuestra E) - estaba usando
        # por error nuestra propia letra K (gf_visitante), por confundir
        # el sistema de columnas de la hoja Mejorado (C-V) con el de la
        # hoja Partidos (E-Y), que usan las mismas letras para cosas
        # distintas.
        am_sin_externos += (0.03 if _num(E, 0) >= 1.6 else 0.02 if _num(E, 0) >= 1.4
                             else (-0.02 if _num(E, 0) < 1.2 else 0))
        am_sin_externos += 0.05 if _num(G, 0) >= 1.9 else (0.02 if _num(G, 0) >= 1.6 else 0)
        am_sin_externos += 0.03 if _num(H, 0) >= 1.9 else (0.03 if _num(H, 0) < 1 else 0)
        am_sin_externos += (0.04 if _num(L, 0) >= 2.3 else 0.03 if _num(L, 0) >= 1.9
                             else 0.02 if _num(L, 0) >= 1.3 else 0)
        am_sin_externos += 0.02 if _num(J, 0) >= 1.5 else (0.01 if _num(J, 0) >= 1.2 else 0)
        am_sin_externos += (-0.03 if _pct(N, 1) > 0.4 else 0.02 if _pct(N, 1) > 0.3
                             else (0.01 if _pct(N, 1) > 0.1 else 0))
        am_sin_externos += (-0.03 if _pct(P, 1) > 0.4 else 0.02 if _pct(P, 1) > 0.3
                             else (0.01 if _pct(P, 1) > 0.1 else 0))
        am_sin_externos += 0.02 if _pct(V, 0) >= 0.65 else (-0.02 if _pct(V, 0) < 0.5 else 0.01)
        am_sin_externos += (0.03 if _pct(W, 0) <= 0.5 else 0.02 if _pct(W, 0) <= 0.65
                             else 0.01 if _pct(W, 0) <= 0.7 else -0.01)

        stake10_cond = (
            ((AG_n + AH_n) / 2 >= 0.6 and (AI_n + AJ_n) / 2 >= 0.77
             and AK_n >= 0.84 and AL_n >= 0.85
             and min(_num(F, 0), _num(H, 0)) >= 1.5)
            or (am_sin_externos >= 0.76 and min(_num(F, 0), _num(H, 0)) >= 1.3)
        )

        # P (cs_visit_pct) resulto ser, con el historico real del Excel
        # (1000+ partidos, CORREGIDO el 13 ago 2026 - ver nota abajo), la
        # variable que mas separa ganados de perdidos. Se aprieta de forma
        # DISTINTA segun la señal, cada una con su propio umbral optimo:
        #   VIP+  -> se deja INTACTA (P<=0.4, sin cambios) - funciona muy bien.
        #   VIP   -> P<=0.25 (61%->69% de acierto, n=113->74, base CORREGIDA)
        #   Fuerte-> P<=0.20 (75%->77% de acierto, n=202->129, base CORREGIDA -
        #            mejora mas chica de lo que se penso al principio, pero real)
        cond_p_vip_mas = True if _blank(P) else (_pct(P, 0) <= 0.4)
        cond_p_vip = True if _blank(P) else (_pct(P, 0) <= 0.25)
        cond_p_fuerte = True if _blank(P) else (_pct(P, 0) <= 0.20)

        vip_cond_base = (
            _num(AD, 0) >= 0.62
            and min(_num(AB, 0), _num(AC, 0)) >= 1.3
            and cond_m and cond_o and cond_n and cond_tiros
            and _num(L, 0) >= 1
        )

        # NOTA IMPORTANTE (13 ago 2026): en una sesion anterior se probaron
        # aqui una "formula intermedia" para VIP (umbrales mas flojos) y un
        # filtro de favorito/no-favorito para VIP y Fuerte (SOT del no
        # favorito, etc.), respaldados en un cruce de datos del Excel que
        # resulto tener un BUG: emparejaba partidos por nombre de equipo en
        # texto entre dos hojas distintas, y perdia cientos de filas por
        # diferencias de tildes/formato - perdida que NO era pareja entre
        # categorias (se llevaba desproporcionadamente Fuerte y rescates).
        # Al corregir el cruce (1494 de 1494 partidos, cero perdidas) esos
        # dos hallazgos NO se sostuvieron:
        #   - Formula intermedia VIP: se penso que daba 66% (n=281) vs 76%
        #     de VIP+ estricta: con datos correctos, VIP normal (P<=0.25) ya
        #     da 69% por si solo, la intermedia no aportaba nada real.
        #   - Filtro SOT favorito/no-favorito en Fuerte: se penso que subia
        #     66%->69%; con datos correctos, Fuerte YA rinde 70-73% sin el
        #     filtro, asi que el filtro no aportaba nada.
        #   - Filtro SOT+BTTS en VIP: unico que broadly se mantuvo (61%->68%)
        #     pero con muestra demasiado chica (n=19-22) para confiar.
        # Se revierte a la forma ORIGINAL (solo el umbral de P) para las
        # tres señales, que es lo unico validado con certeza en ambos
        # cruces (el corregido, Y directamente en produccion).
        if stake10_cond:
            AH = "Stake10"
        elif vip_cond_base and _num(L, 0) >= 1.6 and cond_p_vip_mas:
            AH = "VIP+"
        elif vip_cond_base and _num(L, 0) < 1.6 and cond_p_vip:
            AH = "VIP"
        else:
            fuerte_cond1 = (
                _num(F, 0) >= 1.3 and cond_p_fuerte
                and (True if (_blank(Q) or _blank(R)) else (_num(Q, 0) + _num(R, 0) >= 10))
            )
            if fuerte_cond1:
                AH = "Fuerte"
            else:
                fuerte_cond2 = (
                    AE is not None and AE >= 3.1 and AE < 3.5
                    and _num(F, 0) >= 1.5
                    and _num(W, 0) <= 80
                    and (S not in (None, 0) and 1 / S >= 0.6)
                    and cond_p_fuerte
                )
                AH = "Fuerte" if fuerte_cond2 else "Evitar"

    # --- AJ: Edge BTTS (calculado antes porque AI lo usa) ---
    AJ = None if (_blank(S) or _num(S, 0) == 0 or AD is None) else AD - 1 / S

    # --- AI: NIVEL FINAL +Externas ---
    if AD is None:
        AI = ""
        AI_motivo = ""
        AD_calibrada = None
    else:
        if AH == "Evitar":
            base = "EVITAR"
        elif AH in ("VIP+", "Stake10"):
            base = "JUGAR"
        elif AH in ("VIP", "Fuerte") and _num(AJ, -1) > 0:
            # Umbral subido de -0.027 (el original del Excel) a 0 (edge
            # realmente positivo) - confirmado con datos reales del
            # backtest que las señales con edge negativo o cercano a cero
            # rendian MEJOR que las de edge positivo (91%/75% vs 57%/40%
            # en dos muestras distintas), lo contrario de lo que deberia
            # pasar si el edge fuera una señal de valor real. Ajuste
            # deliberado, no una fidelidad ciega al Excel original.
            base = "JUGAR"
        else:
            w_frac = _pct(W, 1)
            v_frac = _pct(V, 1)
            cond_a = (_blank(X) and _num(F, 0) >= 1.5 and _num(F, 0) < 1.9 and w_frac <= 0.72)
            cond_b = (not _blank(X) and not _blank(Y) and _num(X) < _num(Y)
                      and _num(F, 0) >= 1.5 and _num(F, 0) < 1.9 and w_frac <= 0.72)
            cond_c = (not _blank(X) and not _blank(Y) and _num(Y) < _num(X)
                      and _num(H, 0) >= 1.5 and _num(H, 0) < 1.9 and v_frac <= 0.72)
            base = "JUGAR" if (cond_a or cond_b or cond_c) else "EVITAR"

        extra_1 = (not _blank(E) and not _blank(G) and E >= 1.7 and G >= 1.6
                   and (_pct(V, V if V else 0) + _pct(W, W if W else 0)) / 2 >= 0.7)
        # AS y AT se calculan mas abajo; los adelantamos aqui porque AI depende de ellos
        AS = _calc_oro(V, W, AH, m.pj_local, m.pj_visitante)
        AT = _calc_nivel_confianza(V, W, AH, J, L, M, O, m.pj_local, m.pj_visitante)

        extra_oro_btts = (AS == "ORO_BTTS")
        extra_plata = (AT == "PLATA")
        # Edge positivo agregado a los dos rescates - confirmado con datos
        # reales que dentro de "Evitar" (la categoria de mayor volumen),
        # edge positivo daba +8.9% de ROI vs -15.3% con edge negativo/cero
        # (n=54 vs n=90). Si no hay cuota real todavia para calcular AJ, no
        # se bloquea (mismo criterio de siempre para datos faltantes).
        cond_edge_positivo = True if AJ is None else (AJ > 0)
        extra_evitar_1 = (AH == "Evitar" and AB is not None and AC is not None
                           and 2.5 <= (AB + AC) <= 3.1 and _num(AD, 0) >= 0.55
                           and (not _blank(Q)) and (not _blank(R)) and (_num(Q, 0) + _num(R, 0) >= 10)
                           and cond_edge_positivo)
        # Rango angostado de 0.55-0.6 (antes solo <0.6, sin piso) - confirmado
        # con datos reales del backtest que dentro de este camino, los casos
        # con probabilidad 0.55-0.60 acertaban 67% vs solo 50% los de
        # 0.50-0.55 (n=36 vs n=16). El piso de 0.55 filtra el segmento mas
        # debil sin perder los casos que si funcionan.
        # Se agrega cond_p_vip (P<=0.25) - confirmado con el historico real
        # del Excel: sube el acierto de esta regla de 68% a 76% (n=37).
        extra_evitar_2 = (AH == "Evitar" and _num(F, 0) >= 1.3 and 0.55 <= _num(AD, 1) < 0.6 and cond_p_vip
                           and cond_edge_positivo)

        # --- Ataque Balanceado: señal nueva de 2 niveles (Media/Alta) ---
        # Construida el 13 ago 2026 con el cruce de Excel YA CORREGIDO
        # (Backtest x Partidos, 1494/1494 partidos, cero perdidas - a
        # diferencia de "Paridad Ofensiva", que se elimino por estar
        # construida sobre un cruce con bug). Ajustada por favorito real
        # (cuota 1X2), y validada probando estabilidad partiendo el
        # historico en dos mitades por fecha - las dos mitades quedan
        # POSITIVAS en ambos niveles, señal de que no es una racha:
        #   Media: SOT_favorito>=5, SOT_no_favorito>=4
        #          -> 66% acierto, ROI +1.9% (n=449), mitades +2.4%/+1.4%
        #   Alta : lo de Media + xG_no_favorito>=1.6
        #          -> 69% acierto, ROI +6.6% (n=193), mitades +6.2%/+6.9%
        # La logica: un partido donde el favorito Y el no-favorito generan
        # tiros reales, y ademas el no-favorito tiene xG serio (no es solo
        # un equipo que se cierra atras), predice bien que ambos anoten.
        if _blank(X) or _blank(Y) or _num(X, 0) == _num(Y, 0):
            ataque_bal_tier = None
        else:
            if _num(X, 0) < _num(Y, 0):
                sot_fav_ab, sot_nofav_ab, xg_nofav_ab = _num(Q, 0), _num(R, 0), _num(G, 0)
            else:
                sot_fav_ab, sot_nofav_ab, xg_nofav_ab = _num(R, 0), _num(Q, 0), _num(E, 0)

            media_ab = (sot_fav_ab >= 5 and sot_nofav_ab >= 4)
            if media_ab and xg_nofav_ab >= 1.6:
                ataque_bal_tier = "alta"
            elif media_ab:
                ataque_bal_tier = "media"
            else:
                ataque_bal_tier = None

        extra_ataque_balanceado = (AH == "Evitar" and ataque_bal_tier is not None)

        # NOTA (13 ago 2026): aqui existia una señal nueva "Paridad
        # Ofensiva" (3 niveles), construida el mismo dia con el cruce de
        # Excel que resulto tener el bug de emparejamiento por nombre
        # descrito arriba. Con el cruce corregido (1494/1494 partidos, sin
        # perdidas), los 3 niveles dieron 47-59% de acierto - IGUAL o PEOR
        # que no aplicar ningun filtro dentro de "Evitar" (57% base). Se
        # confirmo ademas en produccion real: su primer dia en vivo, el
        # nivel "Alta" (el que se pensaba mejor, 72% en el analisis con
        # bug) dio 0 aciertos de 4. Se elimina por completo - no hay
        # evidencia real de que aporte nada.

        # Motivo diagnostico: cual regla exacta se activo (se calcula ANTES
        # de decidir el JUGAR, porque ahora la decision depende de cual
        # regla fue). Se sigue registrando el motivo de TODAS las reglas,
        # incluso las que ya no generan JUGAR - asi se les puede seguir
        # midiendo el %acierto y ROI en el backtest para saber si algun dia
        # mejoran y vale la pena reactivarlas.
        if base == "JUGAR":
            if AH == "Stake10":
                AI_motivo = "camino_normal_Stake10"
            elif AH == "VIP+":
                AI_motivo = "camino_normal_VIP+"
            elif AH in ("VIP", "Fuerte"):
                AI_motivo = "camino_normal_VIP_Fuerte_edge_ok"
            else:
                AI_motivo = "rescate_cuota_favorable_1X2"
        elif extra_1:
            AI_motivo = "rescate_xg_alto_btts_historico_alto"
        elif extra_oro_btts:
            AI_motivo = "rescate_oro_btts"
        elif extra_plata:
            AI_motivo = "rescate_plata"
        elif extra_evitar_1:
            AI_motivo = "rescate_evitar_lambda_moderado"
        elif extra_evitar_2:
            AI_motivo = "rescate_evitar_edge_bajo_prob_baja"
        elif extra_ataque_balanceado:
            AI_motivo = f"rescate_ataque_balanceado_{ataque_bal_tier}"
        else:
            AI_motivo = "ninguno_evitar_final"

        # --- Probabilidad MOSTRADA, anclada a grupos reales validados ---
        # A diferencia de AD (la probabilidad Poisson interna, que sigue
        # usandose para los umbrales de las formulas), esta es la que se
        # MUESTRA al usuario - y en vez de un modelo estadistico continuo
        # (que hoy no logra separar mas alla de 55-65% con honestidad, ver
        # sesion del 13 ago), usa el %acierto REAL medido de cada grupo
        # validado. Esto garantiza coherencia con la cuota por diseño: si
        # dice 69%, es porque ESE grupo especifico realmente acierta 69% en
        # el historico (Excel corregido y/o produccion), y la cuota de
        # equilibrio (1/0.69=1.45) es matematicamente consistente con eso.
        # RECONSTRUIDO 13 ago 2026 (sesion final): en vez de una tabla fija
        # por categoria, se construyo un modelo CONTINUO real, con pesos
        # ajustados por regresion logistica (no a mano) sobre 1174 partidos
        # de produccion, y validado con prediccion FUERA de muestra
        # (validacion cruzada de 5 pliegues - la unica forma honesta de
        # medir esto). Combina las DOS variables mas fuertes encontradas
        # hoy con evidencia real:
        #   1. La formula de Stake10 (am_sin_externos, arriba) - AUC 0.559
        #      ella sola, la mejor formula individual de toda la sesion.
        #      Se corrigio ademas un bug real: el termino de "xG Local
        #      anota" (columna K del Excel) estaba usando por error la
        #      variable de goles a favor del visitante (nuestra propia
        #      letra K en otra hoja), en vez de xg_local (nuestra E).
        #   2. cs_max (el mayor % de porteria en cero entre los dos
        #      equipos) - variable ya validada varias veces hoy.
        # Resultado combinado: AUC 0.556 (n=1174, el mejor que sostiene
        # semejante volumen), con calibracion HONESTA por rango:
        #   dijimos 40-50% -> real 50%  (n=400)
        #   dijimos 50-60% -> real 52%  (n=609)
        #   dijimos 60-70% -> real 68%  (n=96)
        # Coeficientes exactos de la regresion (variable, coef, media):
        cs_max_calib = max(_num(N, 0), _num(P, 0))
        logit_calibrado = (
            0.047695
            + 1.537510 * (am_sin_externos - 0.588835)
            + (-0.007913) * (cs_max_calib - 33.163543)
        )
        AD_calibrada = 1 / (1 + math.exp(-logit_calibrado))
        AD_calibrada = max(0.0, min(1.0, AD_calibrada))

        # SOLO estas 3 reglas generan JUGAR - confirmado con datos reales
        # de produccion (n=380, 8 ago 2026) que son las unicas con ROI
        # positivo: +29.7%, +11.3% y +6.4% respectivamente. Las demas
        # (VIP_Fuerte_edge_ok -23%, oro_btts -44.6%, Stake10 -15.2%,
        # paridad_ofensiva -13.7%/-53.7%, evitar_edge_bajo -8.2%) siguen
        # calculandose y registrando su motivo, pero ya no generan apuesta.
        # Para reactivar alguna, basta con agregar su motivo a esta lista.
        # Que reglas pueden generar JUGAR: se decide DINAMICAMENTE segun el
        # ROI real acumulado de cada una (ver reglas_activas.py). Una regla
        # apagada se REACTIVA sola si empieza a rendir bien, y una activa se
        # APAGA sola si empieza a perder - con protecciones para que no se
        # prenda/apague por casualidad (muestra minima de 25 casos, y
        # umbrales distintos para prender que para apagar).
        # El set se inyecta desde afuera via m.reglas_activas (main.py lo
        # calcula una vez por corrida leyendo el historial). Si no viene,
        # se usa el estado inicial conocido como respaldo seguro.
        reglas = m.reglas_activas if m.reglas_activas is not None else reglas_activas.ESTADO_INICIAL

        # Banda de cuota rentable (camino ALTERNATIVO, se suma a las reglas
        # de arriba). Confirmado con datos reales de produccion
        # (n=383, 8 ago 2026) comparando contra apostar A CIEGAS en el
        # MISMO rango de cuota - que es la comparacion justa:
        #   cuota 1.75-2.00 -> 58% acierto, ROI +7.6%, aporte +13.5pp sobre el azar
        # Fuera de esa banda el modelo no aporta nada real: en cuotas mas
        # cortas (<1.7) pierde por el margen de la casa, y en cuotas mas
        # largas (>=2.0) el acierto se desploma a 35%.
        # Banda ampliada de 1.75-2.00 a 1.75-2.05: agrega 9 partidos mas
        # (n=72 -> 81) manteniendo ROI positivo (+5.6% vs +7.6%) Y, lo mas
        # importante, manteniendo la ESTABILIDAD - al partir el historico
        # en dos mitades por fecha, ambas siguen en positivo (+11.2% y
        # +0.1%). Se probo ampliar mas (hasta 2.20, n=92) pero ahi la
        # segunda mitad se va a -4.9%, o sea deja de ser confiable.
        # El PISO de 1.75 es innegociable: bajarlo a 1.70 tira todo a
        # negativo (-1.4%), y a 1.65 o 1.60 tambien (-2.5%, -2.6%).
        cuota_en_banda_rentable_base = (not _blank(S)) and 1.75 <= _num(S, 0) < 2.05

        # REFINAMIENTO 13 ago 2026: dentro de esa misma banda de cuota, se
        # encontraron 2 variables adicionales que suben el ROI de forma
        # marcada y con progresion ordenada (senal de que es real, no
        # ruido) - validado con datos de produccion, estable partiendo el
        # historico en dos mitades por fecha:
        #   banda sola (lo que habia)              -> n=345, ROI -7.0%
        #   + BTTS% minimo de ambos equipos >=42%  -> n=149, ROI +4.9%
        #   + SOT total (ambos equipos) >=9        -> n=52,  ROI +13.6%
        #     (mitades +5.9%/+21.3%, ambas positivas)
        # Se confirmo ademas que SOT_total>=9 NO sirve solo (sin el filtro
        # de BTTS%) - da ROI -9.2% por si solo en la misma banda - la
        # combinacion especifica es lo que aporta, no cada variable por
        # separado.
        btts_min_banda = (not (_blank(V) or _blank(W))) and min(_pct(V, 0), _pct(W, 0)) >= 0.42
        sot_total_banda = (not (_blank(Q) or _blank(R))) and (_num(Q, 0) + _num(R, 0)) >= 9
        cuota_en_banda_rentable = cuota_en_banda_rentable_base and btts_min_banda and sot_total_banda

        # IMPORTANTE: la banda de cuota NO convierte cualquier partido en
        # JUGAR - solo aplica a partidos donde YA se activo alguna regla.
        alguna_regla_activada = AI_motivo != "ninguno_evitar_final"

        jugar = alguna_regla_activada and (
            reglas.get(AI_motivo, False) or cuota_en_banda_rentable
        )

        # Candado maestro: si no hay cuota real de BTTS (None o 0 - FootyStats
        # devuelve 0 cuando el mercado simplemente no esta disponible, no es
        # una cuota real), NUNCA se puede decir JUGAR, sin importar que regla
        # se haya disparado. Sin cuota real no hay forma de calcular edge de
        # verdad ni de colocar la apuesta.
        if _blank(S) or _num(S, 0) == 0:
            jugar = False

        AI = "JUGAR" if jugar else "EVITAR"

    # --- AK: Señal O1.5 ---
    if AF is None:
        AK = ""
    else:
        cond_tiros_8 = True if (_blank(Q) or _blank(R)) else (_num(Q, 0) + _num(R, 0) >= 8)
        cond_tiros_7 = True if (_blank(Q) or _blank(R)) else (_num(Q, 0) + _num(R, 0) >= 7)
        if AF >= 0.78 and AE >= 2.2 and cond_tiros_8:
            AK = "VIP"
        elif AF >= 0.72 and AE >= 2 and cond_tiros_7:
            AK = "Fuerte"
        else:
            AK = "Evitar"

    # --- AL: Sug O1.5 ---
    if AK == "":
        AL = ""
    elif AK == "VIP":
        cond = (AF >= 0.78 and AE >= 2.2
                and max((_num(E, 0) + _num(I, 0)) / 2, (_num(G, 0) + _num(K, 0)) / 2) >= 1.45
                and max((_num(F, 0) + _num(J, 0)) / 2, (_num(H, 0) + _num(L, 0)) / 2) >= 1.1)
        AL = "JUGAR" if cond else "EVITAR"
    elif AK == "Fuerte":
        cond = (AF >= 0.72 and AE >= 2
                and max((_num(E, 0) + _num(I, 0)) / 2, (_num(G, 0) + _num(K, 0)) / 2) >= 1.3
                and max((_num(F, 0) + _num(J, 0)) / 2, (_num(H, 0) + _num(L, 0)) / 2) >= 1)
        AL = "JUGAR" if cond else "EVITAR"
    else:
        AL = "EVITAR"

    # Candado maestro (igual que en BTTS): sin cuota real de Over 1.5, nunca JUGAR.
    if AL == "JUGAR" and (_blank(T) or _num(T, 0) == 0):
        AL = "EVITAR"

    # Franja de cuota rentable: se confirmo con datos reales del backtest
    # que Over 1.5 acertaba MUCHO (75%) pero perdia -54.6u en total, porque
    # las cuotas eran demasiado cortas en casi todo el rango (por ejemplo
    # cuota 1.05-1.10 necesita 94% de acierto para no perder, pero solo
    # llegaba a 84% real). Rango ampliado a 1.20-1.40 por decision del
    # usuario (mas volumen a cambio de un poco mas de perdida que el 1.20-
    # 1.30 original) - y se le agrego ademas una exigencia de edge>0 (la
    # probabilidad propia debe superar la que ya implica la cuota), que
    # confirmado con el historico bajo la perdida de -3.6u a -0.9u dentro
    # de ese mismo rango (n=186 -> n=110), sin bajar el %acierto.
    edge_o15_actual = None if (_blank(T) or _num(T, 0) == 0 or AF is None) else AF - 1 / T
    if AL == "JUGAR" and not (_blank(T) or _num(T, 0) == 0):
        if not (1.20 <= _num(T, 0) < 1.40):
            AL = "EVITAR"
        elif edge_o15_actual is None or edge_o15_actual <= 0:
            AL = "EVITAR"

    # --- AM: Edge O1.5 ---
    AM = None if (_blank(T) or _num(T, 0) == 0 or AF is None) else AF - 1 / T

    # --- AN: Señal O2.5 ---
    if AG is None:
        AN = ""
    else:
        gf_total = _num(I, 0) + _num(K, 0)
        gc_total = _num(J, 0) + _num(L, 0)
        if gf_total >= 3.5 and gc_total >= 2.7 and AG >= 0.62:
            AN = "VIP"
        elif gf_total >= 3.6 and 2.6 <= gc_total < 2.8 and AG >= 0.68:
            AN = "VIP"
        elif gf_total >= 3.5 and gc_total >= 2.8:
            AN = "Fuerte"
        else:
            AN = "Evitar"

    # --- AO: Sug O2.5 ---
    AO = "" if AN == "" else ("JUGAR" if AN in ("VIP", "Fuerte") else "EVITAR")
    motivo_o25 = "camino_normal_o25" if AO == "JUGAR" else ("" if AO == "" else "ninguno_evitar_o25")

    # Rescate v2 (de la hoja "AmbosMarcan-BTTS-Mejorado", columnas AO/AR):
    # una segunda formula independiente para Over 2.5, con otros pesos
    # (0.4/0.4/0.1/0.1 en vez del camino normal). Validada contra 440 casos
    # reales del historico -> 70% de acierto. Solo actua como RESCATE: si
    # el camino normal ya dice JUGAR no cambia nada, solo puede convertir
    # un EVITAR en JUGAR si esta segunda formula esta de acuerdo.
    mu_v2 = (
        0.4 * _num(E, 0) + 0.4 * _num(F, 0) + 0.1 * _num(I, 0) + 0.1 * _num(L, 0)
        + 0.4 * _num(G, 0) + 0.4 * _num(H, 0) + 0.1 * _num(K, 0) + 0.1 * _num(J, 0)
    )
    prob_o25_v2 = max(0.0, min(1.0, 1 - math.exp(-mu_v2) * (1 + mu_v2 + (mu_v2 ** 2) / 2)))
    condicion_v2 = (
        prob_o25_v2 >= 0.62
        or (_num(F, 0) + _num(H, 0) >= 3 and _num(I, 0) + _num(K, 0) >= 3.5)
    )
    if AO == "EVITAR" and condicion_v2:
        AO = "JUGAR"
        motivo_o25 = "rescate_v2_over25"

    # Candado maestro (igual que en BTTS/O1.5): sin cuota real de Over 2.5, nunca JUGAR.
    if AO == "JUGAR" and (_blank(U) or _num(U, 0) == 0):
        AO = "EVITAR"
        motivo_o25 = "candado_sin_cuota"

    # Piso de cuota rentable para Over 2.5 - confirmado con datos reales de
    # produccion (n=251, 8 ago 2026) comparando contra apostar A CIEGAS en
    # el mismo rango de cuota (la comparacion justa):
    #   cuota >= 1.50 -> 60% acierto, ROI +0.1%, aporte +10.0pp sobre el azar
    #   cuota <  1.50 -> pierde fuerte (-10% a -13%), el modelo no supera
    #                     el margen de la casa en cuotas cortas
    # El acierto BAJA al subir el piso (75% en cuotas <1.3 vs 60% aqui),
    # pero el ROI SUBE - porque en cuotas cortas se necesita muchisimo mas
    # acierto para no perder (77% en cuota 1.3 vs 67% en cuota 1.5).
    # ADVERTENCIA sobre este piso: a diferencia de la banda de BTTS (que
    # SI es estable en las dos mitades del historico), este piso de Over
    # 2.5 resulto INESTABLE al revisarlo - primera mitad +14.8%, segunda
    # -14.3%. El +0.1% total es el promedio de una buena racha y una mala,
    # no una ventaja consistente. Se sube a 1.60, que es el unico piso que
    # no se desploma en la segunda mitad (+6.8% / -2.9%), aunque tampoco
    # es solido. Revisar con mas muestra antes de confiar en este mercado.
    if AO == "JUGAR" and _num(U, 0) < 1.60:
        AO = "EVITAR"
        motivo_o25 = "cuota_muy_corta_o25"

    # --- AP: Edge O2.5 ---
    AP = None if (_blank(U) or _num(U, 0) == 0 or AG is None) else AG - 1 / U

    # --- AQ: VIP ELITE ---
    if (AI == "JUGAR" and AH == "VIP"
            and _num(AD, 0) >= 0.62
            and min(_num(AB, 0), _num(AC, 0)) >= 1.12
            and (_num(E, 0) + _num(G, 0)) >= 3.1
            and (_num(Q, 0) + _num(R, 0)) >= 9
            and 1.55 <= _num(S, 0) <= 1.88
            and _pct(M, 0) <= 0.23
            and _pct(O, 0) <= 0.23
            and _pct(N, 0) <= 0.31
            and _pct(P, 0) <= 0.31):
        AQ = "VIP ELITE"
    else:
        AQ = ""

    # --- AR: Señal BTTS+O2.5 ---
    # Se trackean 3 reglas de forma INDEPENDIENTE (a peticion del usuario,
    # 13 ago 2026), para medir cada una por separado con el tiempo en vez
    # de suponer cual es mejor:
    #   1. combo_btts_vipmas: el lado BTTS especificamente es VIP+ (la
    #      señal BTTS mas fuerte que tenemos) - sin importar el rango de
    #      cuota del producto.
    #   2. combo_original_revisar: producto de cuotas entre 2.3-2.6 (el
    #      rango original "REVISAR") - en el historico mostro +45.9% ROI
    #      (n=16, MUESTRA CHICA - vigilar antes de confiar del todo).
    #   3. combo_original_vip: producto de cuotas fuera de 2.6-3.2 (el
    #      rango original "VIP") - en el historico mostro -17.6% ROI
    #      (n=81) - se mantiene activo solo para seguir midiendolo, no
    #      porque ya se confirme rentable.
    # Prioridad si un partido cumple varias a la vez: BTTS VIP+ primero
    # (la mas especifica), luego REVISAR, luego VIP.
    if AI == "JUGAR" and AO == "JUGAR" and _num(S, 0) > 0 and _num(U, 0) > 0:
        # Cuota REAL de la combinada, con el castigo por correlacion (ver
        # comentario en DESCUENTO_CORRELACION_COMBO). Antes se usaba el
        # producto crudo, que es una cuota que ninguna casa paga.
        producto = 1 + (S * U - 1) * (1 - DESCUENTO_CORRELACION_COMBO)
        if AH == "VIP+":
            AR = "VIP"
            motivo_combo = "combo_btts_vipmas"
        elif 2.3 <= producto <= 2.6:
            AR = "REVISAR"
            motivo_combo = "combo_original_revisar"
        elif 2.6 < producto <= 3.2:
            AR = "—"
            motivo_combo = "ninguno_evitar_combo"
        else:
            AR = "VIP"
            motivo_combo = "combo_original_vip"
    else:
        AR = "—"
        motivo_combo = "ninguno_evitar_combo"

    # --- Probabilidad CALIBRADA de la combinada (BTTS y Over 2.5 A LA VEZ) ---
    # Mismo metodo que para BTTS solo: regresion logistica sobre 1174
    # partidos de produccion, validada con prediccion fuera de muestra.
    # Combina Stake10 (corregida) + goles concedidos totales (ambos
    # equipos) - AUC 0.567, calibracion honesta:
    #   dijimos 30-40% -> real 37% (n=460)
    #   dijimos 40-50% -> real 44% (n=552)
    #   dijimos 50-60% -> real 50% (n=104)
    if AD is not None and 'am_sin_externos' in dir() and not _blank(J) and not _blank(L):
        gc_total_calib = _num(J, 0) + _num(L, 0)
        logit_combo = (
            -0.375923
            + (0.335196 / 0.117370) * (am_sin_externos - 0.588835)
            + (-0.134549 / 1.083475) * (gc_total_calib - 2.474097)
        )
        prob_combo_calibrada = 1 / (1 + math.exp(-logit_combo))
        prob_combo_calibrada = max(0.0, min(1.0, prob_combo_calibrada))
    else:
        prob_combo_calibrada = None

    # --- AS: ORO (recalculado con AH final, ya lo teniamos preliminar) ---
    AS = _calc_oro(V, W, AH, m.pj_local, m.pj_visitante)

    # --- AT: nivel de confianza ---
    AT = _calc_nivel_confianza(V, W, AH, J, L, M, O, m.pj_local, m.pj_visitante)

    # --- Primera mitad: Over 0.5 goles antes del descanso (mercado nuevo) ---
    # FootyStats no da xG separado por mitad (solo existe para el partido
    # completo), asi que aqui se usa directamente el promedio de goles
    # anotados/recibidos en el PRIMER TIEMPO que si nos da la API, con la
    # misma logica Poisson que ya usamos en O1.5/O2.5 (lambda = lo que
    # anota un equipo + lo que le conceden al rival, /2 cada lado).
    #
    # LIMITACION IMPORTANTE: esta es la PRIMERA version de este mercado
    # (agosto 2026) - no hay historial todavia para validar en que franja
    # de cuota da ROI positivo (recordar el caso de Over 1.5, que acertaba
    # 75% pero perdia plata por cuotas muy cortas). El umbral de
    # probabilidad usado aqui (VIP>=0.75, Fuerte>=0.68) es un punto de
    # partida razonable, NO esta validado con datos reales todavia. Revisar
    # con el historico en 1-2 semanas, igual que se hizo con Stake10.
    if m.scored_ht_local is None or m.conceded_ht_visitante is None or m.scored_ht_visitante is None or m.conceded_ht_local is None:
        prob_over05_ht = None
        prob_over15_ht = None
        AH_ht = ""
    else:
        lambda_local_ht = (_num(m.scored_ht_local, 0) + _num(m.conceded_ht_visitante, 0)) / 2
        lambda_visit_ht = (_num(m.scored_ht_visitante, 0) + _num(m.conceded_ht_local, 0)) / 2
        mu_total_ht = lambda_local_ht + lambda_visit_ht
        prob_over05_ht = 1 - math.exp(-mu_total_ht)
        # Nota informativa (no es una señal propia con JUGAR/EVITAR, solo un
        # dato extra para destacar los partidos con mas probabilidad de que
        # sean 2+ goles en el primer tiempo, no solo 1). Reutiliza el mismo
        # mu_total_ht ya calculado, con la formula de "al menos 2" en vez de
        # "al menos 1" (misma logica que O1.5 del partido completo).
        prob_over15_ht = 1 - math.exp(-mu_total_ht) * (1 + mu_total_ht)

        pct_local_ht = _pct(m.over05_ht_pct_local, 0)
        pct_visit_ht = _pct(m.over05_ht_pct_visitante, 0)
        confirmacion_historica = min(pct_local_ht, pct_visit_ht) >= 0.55

        if prob_over05_ht >= 0.75 and confirmacion_historica:
            AH_ht = "VIP"
        elif prob_over05_ht >= 0.68:
            AH_ht = "Fuerte"
        else:
            AH_ht = "Evitar"

    if AH_ht == "":
        AI_ht, AI_ht_motivo, AM_ht = "", "", None
    else:
        AI_ht = "JUGAR" if AH_ht in ("VIP", "Fuerte") else "EVITAR"
        AI_ht_motivo = "camino_normal_ht05" if AI_ht == "JUGAR" else "ninguno_evitar_ht05"
        # Candado maestro: sin cuota real, nunca JUGAR (mismo patron que BTTS/O1.5/O2.5).
        if AI_ht == "JUGAR" and (_blank(m.cuota_ht_over05) or _num(m.cuota_ht_over05, 0) == 0):
            AI_ht = "EVITAR"
            AI_ht_motivo = "candado_sin_cuota"
        AM_ht = (None if (_blank(m.cuota_ht_over05) or _num(m.cuota_ht_over05, 0) == 0 or prob_over05_ht is None)
                 else prob_over05_ht - 1 / m.cuota_ht_over05)

    # --- Nota informativa: btts_min_alto ---
    # NO genera JUGAR por si sola (a peticion del usuario, 13 ago 2026) -
    # solo se guarda como nota, para medir su ROI real con el tiempo antes
    # de decidir si activarla. Hallazgo: dentro de la banda de cuota
    # rentable (1.75-2.05, ya activa), exigir ademas que el MENOR de los
    # dos BTTS% de temporada sea >=42% subio el ROI de -7.0% a +4.9%
    # (n=345->149), estable en las dos mitades del historico partido por
    # fecha (+3.6%/+6.1%). La logica: dentro de la banda ya rentable,
    # descarta los partidos donde solo UN equipo suele anotar (favorito
    # ofensivo contra rival muy cerrado), que no favorecen el BTTS aunque
    # la cuota se vea atractiva.
    if not (_blank(V) or _blank(W)):
        btts_min_alto = min(_pct(V, 0), _pct(W, 0)) >= 0.42
    else:
        btts_min_alto = None

    return MatchSignals(
        lambda_local=AB, lambda_visitante=AC, prob_btts=AD, prob_btts_calibrada=AD_calibrada, mu_total=AE,
        p_over15=AF, p_over25=AG,
        senal_btts_optimizada=AH, nivel_final=AI, nivel_final_motivo=AI_motivo, edge_btts=AJ,
        senal_o15=AK, sug_o15=AL, edge_o15=AM,
        senal_o25=AN, sug_o25=AO, motivo_o25=motivo_o25, edge_o25=AP,
        vip_elite=AQ, senal_btts_o25=AR, oro=AS, nivel_confianza=AT,
        stake10_prob_estimada=locals().get("am_sin_externos"),
        stake10_cumple=locals().get("stake10_cond"),
        prob_over05_ht=prob_over05_ht, prob_over15_ht=prob_over15_ht, senal_ht05=AH_ht, sug_ht05=AI_ht,
        motivo_ht05=AI_ht_motivo, edge_ht05=AM_ht,
        prob_combo_calibrada=prob_combo_calibrada,
        btts_min_alto=btts_min_alto,
        motivo_combo=motivo_combo,
    )


def _calc_oro(V, W, AH, pj_local=None, pj_visitante=None) -> str:
    v_frac = _pct(V, 0)
    w_frac = _pct(W, 0)
    minimo = min(v_frac, w_frac)
    if minimo >= 0.65 and AH == "Fuerte":
        return "ORO_FUERTE"
    # ORO_BTTS es la UNICA via de esta funcion sin ningun otro filtro de
    # calidad (no exige tier Fuerte, no usa mercado, no usa edge) - por eso
    # es la mas expuesta a que un 80-100% de BTTS sea solo una racha corta
    # sin muestra suficiente para confiar en ella. Se confirmo con datos
    # reales que 6 de 12 casos recientes mostraban un 100% literal con
    # apenas 5-8 partidos jugados, y rendian muy mal (30-43% de acierto).
    # Se exige el doble de partidos que el minimo general (10 en vez de 5)
    # para reducir el riesgo de "racha corta disfrazada de elite real".
    min_pj = min(_num(pj_local, 0), _num(pj_visitante, 0)) if (pj_local is not None and pj_visitante is not None) else None
    if minimo >= 0.8 and (min_pj is None or min_pj >= 10):
        return "ORO_BTTS"
    return ""


def _calc_nivel_confianza(V, W, AH, J, L, M, O, pj_local=None, pj_visitante=None) -> str:
    v_frac = _pct(V, 0)
    w_frac = _pct(W, 0)
    minimo = min(v_frac, w_frac)
    if AH == "Fuerte" and minimo >= 0.65 and max(_num(J, 0), _num(L, 0)) >= 1.4:
        return "ÉLITE"
    # Misma proteccion que en ORO_BTTS (ver comentario ahi) - PLATA usa
    # exactamente la misma condicion base (minimo>=0.8, sin tier Fuerte),
    # asi que le aplica el mismo riesgo de rachas cortas sin muestra.
    min_pj = min(_num(pj_local, 0), _num(pj_visitante, 0)) if (pj_local is not None and pj_visitante is not None) else None
    if minimo >= 0.8 and (min_pj is None or min_pj >= 10):
        return "PLATA"
    if AH == "Fuerte" and max(_pct(M, 1), _pct(O, 1)) <= 0.25:
        return "BRONCE"
    return ""
