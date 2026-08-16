"""
backtest.py
===========
Replica el comportamiento de la pestaña "Backtest" del Excel, pero
automatizado: en vez de que tu ingreses los goles a mano en las columnas
P/Q, este modulo los descarga de FootyStats una vez el partido termina.
"""
from __future__ import annotations
import os
import json
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

# Descuento por correlacion en la combinada BTTS + Over 2.5 (ver el mismo
# valor y explicacion completa en signals.py, DESCUENTO_CORRELACION_COMBO).
# Multiplicar cuota_btts x cuota_o25 sin mas da un precio que ninguna casa
# paga, porque los dos eventos estan correlacionados (P(Over2.5) sube de
# 52.9% a 79.8% cuando ya se dio el BTTS). Se duplica aqui la constante
# (en vez de importar signals.py) para no crear una dependencia circular -
# si se cambia un valor, hay que cambiar el otro tambien.
DESCUENTO_CORRELACION_COMBO = 0.275
import pandas as pd

import config

HISTORIAL_PATH = os.path.join(config.OUTPUT_DIR, "historial_predicciones.csv")

COLUMNS = [
    "fecha", "fecha_partido", "match_id", "season_id", "liga", "local", "visitante",
    "cuota_btts", "cuota_btts_actual", "senal_btts",
    "prob_btts",
    "motivo_btts",
    "cuota_o15", "cuota_o15_actual", "senal_o15", "prob_o15", "tier_o15",
    "cuota_o25", "cuota_o25_actual", "senal_o25", "prob_o25", "tier_o25", "motivo_o25",
    "senal_btts_optimizada",
    "oro",
    "nivel_confianza",
    "vip_elite",
    "senal_btts_o25", "motivo_combo", "btts_min_alto",
    "marcador_exacto_sugerido", "marcador_exacto_real", "marcador_exacto_acierto",
    "usa_temporada_anterior", "usa_promedio_general", "es_equipo_filial",
    "xg_local", "xga_local", "xg_visitante", "xga_visitante",
    "gf_local", "gc_local", "gf_visitante", "gc_visitante",
    "sot_local", "sot_visitante",
    "failed_local_pct", "cs_local_pct", "failed_visit_pct", "cs_visit_pct",
    "btts_pct_local", "btts_pct_visitante",
    "over25_and_btts_pct_local", "over25_and_btts_pct_visitante",
    "ppg_local", "ppg_visitante", "shots_local", "shots_visitante",
    "avght_local", "avght_visitante", "avg2h_local", "avg2h_visitante",
    "resultado_combo", "profit_combo",
    "goles_local", "goles_visitante",
    "resultado_btts", "resultado_o15", "resultado_o25",
    "profit_btts", "profit_o15", "profit_o25",
    "scored_ht_local", "conceded_ht_local", "scored_ht_visitante", "conceded_ht_visitante",
    "over05_ht_pct_local", "over05_ht_pct_visitante",
    "cuota_ht_over05", "cuota_ht_over05_actual",
    "prob_over05_ht", "prob_over15_ht", "senal_ht05", "sug_ht05", "motivo_ht05",
    "ht_goal_count", "resultado_ht05", "profit_ht05",
    "formula_version",
    "cuota_local", "cuota_visitante",
    "profit_btts_hipotetico",
]


def _load() -> pd.DataFrame:
    if os.path.exists(HISTORIAL_PATH):
        df = pd.read_csv(HISTORIAL_PATH)
    else:
        df = pd.DataFrame(columns=COLUMNS)

    for col in COLUMNS:
        if col not in df.columns:
            df[col] = None

    bool_cols = ["resultado_btts", "resultado_o15", "resultado_o25", "resultado_combo",
                 "marcador_exacto_acierto", "usa_temporada_anterior", "usa_promedio_general",
                 "es_equipo_filial", "resultado_ht05"]
    bool_map = {"True": True, "False": False, True: True, False: False}
    for col in bool_cols:
        if col not in df.columns:
            df[col] = None
        df[col] = df[col].map(lambda v: bool_map.get(v, None) if pd.notna(v) else None).astype(object)

    numeric_object_cols = ["goles_local", "goles_visitante", "profit_btts", "profit_o15",
                            "profit_o25", "profit_combo", "profit_ht05", "profit_btts_hipotetico"]
    for col in numeric_object_cols:
        if col not in df.columns:
            df[col] = None
        df[col] = df[col].astype(object).where(df[col].notna(), None)

    return df


def _save(df: pd.DataFrame) -> None:
    os.makedirs(config.OUTPUT_DIR, exist_ok=True)
    df.to_csv(HISTORIAL_PATH, index=False, encoding="utf-8-sig")


def limpiar_filas_incompletas(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    if df.empty:
        return df, 0
    sin_jugar = df["goles_local"].isna()
    todas_las_prob_vacias = (
        df["prob_btts"].isna() & df["prob_o15"].isna() & df["prob_o25"].isna()
    )
    sin_fecha_partido = df["fecha_partido"].isna() | (df["fecha_partido"] == "") if "fecha_partido" in df.columns else True
    a_borrar = sin_jugar & (todas_las_prob_vacias | sin_fecha_partido)
    n = int(a_borrar.sum())
    if n:
        df = df[~a_borrar].reset_index(drop=True)
    return df, n


def record_predictions(filas: list[dict], fecha: str) -> set:
    historial = _load()
    historial, n_limpiadas = limpiar_filas_incompletas(historial)
    if n_limpiadas:
        print(f"[backtest] {n_limpiadas} partidos incompletos eliminados del historial")
        _save(historial)

    nuevas = []
    nuevos_ids = set()
    ya_existentes = set(historial.get("match_id", []))

    pendientes_idx = {}
    if not historial.empty:
        sin_jugar = historial[historial["goles_local"].isna()]
        pendientes_idx = {mid: idx for idx, mid in sin_jugar["match_id"].items()}

    cambios_cuota = 0
    for f in filas:
        mid = f.get("match_id")

        if mid in pendientes_idx:
            idx = pendientes_idx[mid]
            if f.get("cuota_btts") is not None:
                historial.at[idx, "cuota_btts_actual"] = f.get("cuota_btts")
            if f.get("cuota_o15") is not None:
                historial.at[idx, "cuota_o15_actual"] = f.get("cuota_o15")
            if f.get("cuota_o25") is not None:
                historial.at[idx, "cuota_o25_actual"] = f.get("cuota_o25")
            if f.get("cuota_ht_over05") is not None:
                historial.at[idx, "cuota_ht_over05_actual"] = f.get("cuota_ht_over05")
            cambios_cuota += 1

        if mid in ya_existentes or mid in nuevos_ids:
            continue
        nuevos_ids.add(mid)
        nuevas.append({
            "fecha": fecha,
            "fecha_partido": f.get("fecha_partido") or fecha,
            "match_id": f.get("match_id"),
            "season_id": f.get("season_id"),
            "liga": f.get("liga"),
            "local": f.get("local"),
            "visitante": f.get("visitante"),
            "cuota_btts": f.get("cuota_btts"),
            "cuota_btts_actual": f.get("cuota_btts"),
            "senal_btts": f.get("nivel_final"),
            "prob_btts": f.get("prob_btts"),
            "motivo_btts": f.get("nivel_final_motivo"),
            "cuota_o15": f.get("cuota_o15"),
            "cuota_o15_actual": f.get("cuota_o15"),
            "senal_o15": f.get("sug_o15"),
            "prob_o15": f.get("prob_o15"),
            "tier_o15": f.get("senal_o15"),
            "cuota_o25": f.get("cuota_o25"),
            "cuota_o25_actual": f.get("cuota_o25"),
            "senal_o25": f.get("sug_o25"),
            "prob_o25": f.get("prob_o25"),
            "tier_o25": f.get("senal_o25"),
            "motivo_o25": f.get("motivo_o25"),
            "senal_btts_optimizada": f.get("senal_btts_optimizada"),
            "oro": f.get("oro"),
            "nivel_confianza": f.get("nivel_confianza"),
            "vip_elite": f.get("vip_elite"),
            "senal_btts_o25": f.get("senal_btts_o25"),
            "motivo_combo": f.get("motivo_combo"),
            "btts_min_alto": f.get("btts_min_alto"),
            "marcador_exacto_sugerido": f.get("marcador_exacto_sugerido"),
            "usa_temporada_anterior": f.get("usa_temporada_anterior"),
            "usa_promedio_general": f.get("usa_promedio_general"),
            "es_equipo_filial": f.get("es_equipo_filial"),
            "xg_local": f.get("xg_local"), "xga_local": f.get("xga_local"),
            "xg_visitante": f.get("xg_visitante"), "xga_visitante": f.get("xga_visitante"),
            "gf_local": f.get("gf_local"), "gc_local": f.get("gc_local"),
            "gf_visitante": f.get("gf_visitante"), "gc_visitante": f.get("gc_visitante"),
            "sot_local": f.get("sot_local"), "sot_visitante": f.get("sot_visitante"),
            "failed_local_pct": f.get("failed_local_pct"), "cs_local_pct": f.get("cs_local_pct"),
            "failed_visit_pct": f.get("failed_visit_pct"), "cs_visit_pct": f.get("cs_visit_pct"),
            "btts_pct_local": f.get("btts_pct_local"), "btts_pct_visitante": f.get("btts_pct_visitante"),
            "over25_and_btts_pct_local": f.get("over25_and_btts_pct_local"),
            "over25_and_btts_pct_visitante": f.get("over25_and_btts_pct_visitante"),
            "ppg_local": f.get("ppg_local"), "ppg_visitante": f.get("ppg_visitante"),
            "shots_local": f.get("shots_local"), "shots_visitante": f.get("shots_visitante"),
            "avght_local": f.get("avght_local"), "avght_visitante": f.get("avght_visitante"),
            "avg2h_local": f.get("avg2h_local"), "avg2h_visitante": f.get("avg2h_visitante"),
            "marcador_exacto_real": None, "marcador_exacto_acierto": None,
            "resultado_combo": None, "profit_combo": None,
            "goles_local": None, "goles_visitante": None,
            "resultado_btts": None, "resultado_o15": None, "resultado_o25": None,
            "profit_btts": None, "profit_o15": None, "profit_o25": None,
            "scored_ht_local": f.get("scored_ht_local"), "conceded_ht_local": f.get("conceded_ht_local"),
            "scored_ht_visitante": f.get("scored_ht_visitante"), "conceded_ht_visitante": f.get("conceded_ht_visitante"),
            "over05_ht_pct_local": f.get("over05_ht_pct_local"), "over05_ht_pct_visitante": f.get("over05_ht_pct_visitante"),
            "cuota_ht_over05": f.get("cuota_ht_over05"), "cuota_ht_over05_actual": f.get("cuota_ht_over05"),
            "prob_over05_ht": f.get("prob_over05_ht"),
            "prob_over15_ht": f.get("prob_over15_ht"),
            "senal_ht05": f.get("senal_ht05"),
            "sug_ht05": f.get("sug_ht05"),
            "motivo_ht05": f.get("motivo_ht05"),
            "ht_goal_count": None, "resultado_ht05": None, "profit_ht05": None,
            "formula_version": f.get("formula_version"),
            "cuota_local": f.get("cuota_local"), "cuota_visitante": f.get("cuota_visitante"),
            "profit_btts_hipotetico": None,
        })

    if nuevas:
        historial = pd.concat([historial, pd.DataFrame(nuevas)], ignore_index=True)

    if nuevas or cambios_cuota:
        _save(historial)

    return nuevos_ids


def _profit(stake: float, señal: str, resultado, cuota) -> float:
    if señal != "JUGAR":
        return 0.0
    if resultado is None or pd.isna(resultado) or cuota is None or pd.isna(cuota):
        return None
    return stake * (cuota - 1) if bool(resultado) else -stake


def _profit_hipotetico(stake: float, resultado, cuota) -> float:
    """
    Igual que _profit pero SIN mirar si la señal fue JUGAR - calcula cuanto
    se habria ganado/perdido si se hubiera apostado. Se usa para seguir
    midiendo el rendimiento de reglas desactivadas (las que hoy no generan
    JUGAR por venir con ROI negativo), y poder detectar si mejoran con el
    tiempo y vale la pena reactivarlas.
    """
    if resultado is None or pd.isna(resultado) or cuota is None or pd.isna(cuota):
        return None
    return stake * (cuota - 1) if bool(resultado) else -stake


def update_results(client, stake: float | None = None) -> int:
    stake = stake if stake is not None else config.STAKE_UNIDADES
    historial = _load()
    if historial.empty:
        return 0

    pendientes = historial[historial["goles_local"].isna()]
    if pendientes.empty:
        return 0

    actualizados = 0
    matches_cache: dict[int, dict] = {}

    for season_id in pendientes["season_id"].dropna().unique():
        try:
            matches = client.league_matches(int(season_id))
        except Exception as e:
            print(f"[backtest] no se pudo consultar season_id {season_id}: {e}")
            continue
        for m in matches:
            matches_cache[m.get("id")] = m

    for idx, row in pendientes.iterrows():
        m = matches_cache.get(row["match_id"])
        if not m or m.get("status") != "complete":
            continue

        gl = m.get("homeGoalCount")
        gv = m.get("awayGoalCount")
        if gl is None or gv is None:
            continue

        res_btts = (gl > 0 and gv > 0)
        res_o15 = ((gl + gv) >= 2)
        res_o25 = ((gl + gv) >= 3)

        historial.at[idx, "goles_local"] = gl
        historial.at[idx, "goles_visitante"] = gv
        historial.at[idx, "resultado_btts"] = res_btts
        historial.at[idx, "resultado_o15"] = res_o15
        historial.at[idx, "resultado_o25"] = res_o25
        historial.at[idx, "profit_btts"] = _profit(stake, row["senal_btts"], res_btts, row["cuota_btts"])
        historial.at[idx, "profit_btts_hipotetico"] = _profit_hipotetico(stake, res_btts, row["cuota_btts"])
        historial.at[idx, "profit_o15"] = _profit(stake, row["senal_o15"], res_o15, row["cuota_o15"])
        historial.at[idx, "profit_o25"] = _profit(stake, row["senal_o25"], res_o25, row["cuota_o25"])

        res_combo = res_btts and res_o25
        historial.at[idx, "resultado_combo"] = res_combo
        señal_combo = row["senal_btts_o25"]
        if señal_combo in ("VIP", "REVISAR") and pd.notna(row["cuota_btts"]) and pd.notna(row["cuota_o25"]):
            # Cuota REAL, con el descuento por correlacion - antes se
            # calculaba la ganancia con el producto crudo (una cuota que
            # ninguna casa paga), inflando todas las unidades de este
            # mercado de forma ficticia.
            cuota_producto = row["cuota_btts"] * row["cuota_o25"]
            cuota_combo = 1 + (cuota_producto - 1) * (1 - DESCUENTO_CORRELACION_COMBO)
            historial.at[idx, "profit_combo"] = stake * (cuota_combo - 1) if res_combo else -stake
        else:
            historial.at[idx, "profit_combo"] = 0.0

        ht_goal_count = m.get("HTGoalCount")
        if ht_goal_count is None:
            ht_a = m.get("ht_goals_team_a")
            ht_b = m.get("ht_goals_team_b")
            if ht_a is not None and ht_b is not None:
                ht_goal_count = ht_a + ht_b
        if ht_goal_count is not None:
            res_ht05 = (ht_goal_count >= 1)
            historial.at[idx, "ht_goal_count"] = ht_goal_count
            historial.at[idx, "resultado_ht05"] = res_ht05
            historial.at[idx, "profit_ht05"] = _profit(stake, row["sug_ht05"], res_ht05, row["cuota_ht_over05"])

        marcador_real = f"{int(gl)}-{int(gv)}"
        historial.at[idx, "marcador_exacto_real"] = marcador_real
        sugerido = row.get("marcador_exacto_sugerido")
        if pd.notna(sugerido):
            historial.at[idx, "marcador_exacto_acierto"] = (sugerido == marcador_real)

        actualizados += 1

    if actualizados:
        _save(historial)
    return actualizados


def _market_stats(df: pd.DataFrame, señal_col: str, resultado_col: str, profit_col: str, cuota_col: str, stake: float) -> dict:
    jugados = df[df[señal_col] == "JUGAR"]
    con_resultado = jugados[jugados[resultado_col].notna()]
    bets = len(con_resultado)
    wins = int(con_resultado[resultado_col].sum()) if bets else 0
    acierto = (wins / bets * 100) if bets else None
    profit = con_resultado[profit_col].sum() if bets else 0.0
    roi = (profit / (bets * stake) * 100) if bets else None
    cuota_prom = con_resultado[cuota_col].mean() if bets else None
    return {"bets": bets, "wins": wins, "acierto_pct": acierto, "profit_u": profit, "roi_pct": roi, "cuota_prom": cuota_prom}


def _tier_stats(df: pd.DataFrame, tier_col: str, tier_value: str, resultado_col: str, profit_col: str = "profit_btts") -> dict:
    subset = df[(df[tier_col] == tier_value) & (df[resultado_col].notna())]
    total = len(subset)
    ganados = int(subset[resultado_col].sum()) if total else 0
    acierto = (ganados / total * 100) if total else None
    profit_u = subset[profit_col].sum() if total else 0.0
    total_detectado = int((df[tier_col] == tier_value).sum())
    return {"total": total, "ganados": ganados, "acierto_pct": acierto, "profit_u": profit_u, "total_detectado": total_detectado}


def _motivo_stats(df: pd.DataFrame, señal_col: str = "senal_btts", motivo_col: str = "motivo_btts", resultado_col: str = "resultado_btts", profit_col: str = "profit_btts", sin_motivo_valor: str = "ninguno_evitar_final") -> dict:
    """
    %acierto y ROI desglosado por CUAL regla especifica se activo.

    IMPORTANTE: incluye TODAS las reglas que se activaron, no solo las que
    generaron JUGAR - esto permite seguir midiendo el rendimiento de reglas
    que estan desactivadas (porque venian con ROI negativo) y detectar si
    en algun momento empiezan a mejorar y vale la pena reactivarlas. Para
    las reglas desactivadas, el "profit" se calcula igual como SI se
    hubieran jugado (apuesta hipotetica), con la cuota que tenian.
    """
    con_motivo = df[df[motivo_col].notna() & (df[motivo_col] != "") & (df[motivo_col] != sin_motivo_valor)]
    out = {}
    for motivo in con_motivo[motivo_col].dropna().unique():
        total_detectado = int((con_motivo[motivo_col] == motivo).sum())
        subset = con_motivo[(con_motivo[motivo_col] == motivo) & (con_motivo[resultado_col].notna())]
        total = len(subset)
        activa = bool((subset[señal_col] == "JUGAR").any()) if total else None
        if total == 0:
            out[motivo] = {"total": 0, "ganados": 0, "acierto_pct": None, "profit_u": 0.0,
                            "total_detectado": total_detectado, "activa": activa}
            continue
        ganados = int(subset[resultado_col].sum())
        out[motivo] = {
            "total": total,
            "ganados": ganados,
            "acierto_pct": ganados / total * 100,
            "profit_u": subset[profit_col].sum(),
            "total_detectado": total_detectado,
            "activa": activa,
        }
    return out


def _nota_btts_min_alto_stats(df: pd.DataFrame) -> dict:
    """
    Estadisticas de la nota informativa "btts_min_alto" (el menor de los 2
    BTTS% de temporada >= 42%), SOLO dentro de los partidos que ya son
    JUGAR en BTTS - compara el rendimiento real de los que cumplen la nota
    contra los que no, para decidir con datos si algun dia vale la pena
    activarla como filtro real. No genera JUGAR por si sola (13 ago 2026).
    """
    jugar = df[(df["senal_btts"] == "JUGAR") & df["btts_min_alto"].notna()]
    out = {}
    for etiqueta, valor in [("cumple_nota", True), ("no_cumple_nota", False)]:
        sub = jugar[(jugar["btts_min_alto"] == valor) & jugar["resultado_btts"].notna()]
        if len(sub) == 0:
            out[etiqueta] = {"total": 0, "ganados": 0, "acierto_pct": None, "profit_u": 0.0}
            continue
        ganados = int(sub["resultado_btts"].sum())
        out[etiqueta] = {
            "total": len(sub),
            "ganados": ganados,
            "acierto_pct": ganados / len(sub) * 100,
            "profit_u": float(sub["profit_btts"].sum()),
        }
    return out


def _combo_stats(df: pd.DataFrame, tier_value: str, stake: float) -> dict:
    jugados = df[df["senal_btts_o25"] == tier_value]
    con_resultado = jugados[jugados["resultado_combo"].notna()]
    bets = len(con_resultado)
    wins = int(con_resultado["resultado_combo"].sum()) if bets else 0
    acierto = (wins / bets * 100) if bets else None
    profit = con_resultado["profit_combo"].sum() if bets else 0.0
    roi = (profit / (bets * stake) * 100) if bets else None
    cuota_prom = (1 + (con_resultado["cuota_btts"] * con_resultado["cuota_o25"] - 1) * (1 - DESCUENTO_CORRELACION_COMBO)).mean() if bets else None
    return {"bets": bets, "wins": wins, "acierto_pct": acierto, "profit_u": profit, "roi_pct": roi, "cuota_prom": cuota_prom}


def _marcador_exacto_stats(df: pd.DataFrame) -> dict:
    con_resultado = df[df["marcador_exacto_acierto"].notna()]
    total = len(con_resultado)
    aciertos = int(con_resultado["marcador_exacto_acierto"].sum()) if total else 0
    acierto_pct = (aciertos / total * 100) if total else None
    return {"total": total, "aciertos": aciertos, "acierto_pct": acierto_pct}


def _calibracion_stats(df: pd.DataFrame, prob_col: str, resultado_col: str) -> list[dict]:
    con_datos = df[df[prob_col].notna() & df[resultado_col].notna()]
    if con_datos.empty:
        return []

    rangos = [(0.5, 0.6), (0.6, 0.7), (0.7, 0.8), (0.8, 0.9), (0.9, 1.01)]
    out = []
    for lo, hi in rangos:
        subset = con_datos[(con_datos[prob_col] >= lo) & (con_datos[prob_col] < hi)]
        total = len(subset)
        if total == 0:
            continue
        aciertos = int(subset[resultado_col].sum())
        acierto_pct = aciertos / total * 100
        prob_dicha_prom = subset[prob_col].mean() * 100
        out.append({
            "rango": f"{int(lo*100)}-{int(min(hi,1)*100)}%",
            "total": total,
            "aciertos": aciertos,
            "acierto_pct": acierto_pct,
            "prob_dicha_prom": prob_dicha_prom,
            "diferencia": acierto_pct - prob_dicha_prom,
        })
    return out


def _liga_stats(df: pd.DataFrame, min_muestra: int = 3) -> list[dict]:
    """
    %acierto, ROI y clasificacion de confianza de BTTS, desglosado POR LIGA.
    La clasificacion de confianza ("alta"/"baja") es COMPLETAMENTE dinamica
    - se recalcula cada vez que corre este archivo, en base al ROI real
    acumulado hasta el momento, con un minimo de muestra mas estricto que
    la tabla general (15 partidos) para evitar marcar una liga como buena o
    mala con muy pocos casos. No es una lista fija - conforme se acumulen
    mas partidos, una liga puede pasar de "baja" a "alta" (o al reves) sola,
    sin que haga falta tocar el codigo.
    """
    jugados = df[(df["senal_btts"] == "JUGAR") & (df["resultado_btts"].notna())]
    if jugados.empty:
        return []
    out = []
    for liga in jugados["liga"].dropna().unique():
        subset = jugados[jugados["liga"] == liga]
        total = len(subset)
        if total < min_muestra:
            continue
        ganados = int(subset["resultado_btts"].sum())
        profit_u = subset["profit_btts"].sum()
        roi_pct = profit_u / total * 100

        # Umbral de confianza: minimo 15 partidos (distinto del min_muestra
        # general de la tabla) - +10% ROI para "alta", -20% ROI para "baja".
        # Estos numeros de corte se basaron en el primer barrido real hecho
        # hoy (7 ligas con +10% a +40%, 8 ligas con -27% a -57%) - se pueden
        # revisar mas adelante con mas muestra acumulada.
        if total >= 15 and roi_pct >= 10:
            confianza = "alta"
        elif total >= 15 and roi_pct <= -20:
            confianza = "baja"
        else:
            confianza = None

        out.append({
            "liga": liga,
            "total": total,
            "ganados": ganados,
            "acierto_pct": ganados / total * 100,
            "profit_u": profit_u,
            "roi_pct": roi_pct,
            "confianza": confianza,
        })
    out.sort(key=lambda x: x["acierto_pct"])
    return out


def _movimiento_stats(df: pd.DataFrame, umbral_pp: float = 4.0) -> dict:
    combos = [
        ("senal_btts", "resultado_btts", "cuota_btts", "cuota_btts_actual"),
        ("senal_o15", "resultado_o15", "cuota_o15", "cuota_o15_actual"),
        ("senal_o25", "resultado_o25", "cuota_o25", "cuota_o25_actual"),
    ]
    con_movimiento, sin_movimiento = [], []
    for señal_col, resultado_col, cuota_col, actual_col in combos:
        if actual_col not in df.columns:
            continue
        jugados = df[(df[señal_col] == "JUGAR") & (df[resultado_col].notna())]
        for _, row in jugados.iterrows():
            orig, act = row.get(cuota_col), row.get(actual_col)
            if pd.isna(orig) or pd.isna(act) or orig in (0, None) or act in (0, None):
                continue
            mov = (1 / act - 1 / orig) * 100
            resultado = bool(row[resultado_col])
            (con_movimiento if mov >= umbral_pp else sin_movimiento).append(resultado)

    def _resumen(lista):
        n = len(lista)
        if n == 0:
            return {"total": 0, "aciertos": 0, "acierto_pct": None}
        ac = sum(lista)
        return {"total": n, "aciertos": ac, "acierto_pct": ac / n * 100}

    return {"con_movimiento": _resumen(con_movimiento), "sin_movimiento": _resumen(sin_movimiento)}


def build_summary(stake: float | None = None, desde_fecha: str | None = None, hasta_fecha: str | None = None) -> dict:
    stake = stake if stake is not None else config.STAKE_UNIDADES
    vacio = {"mercados": {}, "categorias": {}, "motivos": {}, "motivos_o25": {}, "motivos_ht05": {}, "motivos_combo": {}, "nota_btts_min_alto": {},
             "marcador_exacto": {"total": 0, "aciertos": 0, "acierto_pct": None}, "calibracion": {},
             "por_liga": [],
             "movimiento_validacion": {"con_movimiento": {"total": 0, "aciertos": 0, "acierto_pct": None},
                                        "sin_movimiento": {"total": 0, "aciertos": 0, "acierto_pct": None}},
             "total_partidos_evaluados": 0}
    historial = _load()
    if historial.empty:
        return vacio

    # IMPORTANTE: filtrar por fecha_partido (el dia REAL en que se juega el
    # partido), no por "fecha" (el dia en que se DETECTO). Un partido que
    # se juega hoy pero se detecto ayer (parte del pre-fetch de "mañana")
    # tiene fecha=ayer y fecha_partido=hoy - filtrar por "fecha" lo dejaba
    # afuera del resumen "HOY" aunque ya tuviera resultado y apareciera
    # correctamente en las pestañas individuales de cada mercado (que si
    # usan fecha_partido). Con respaldo a "fecha" para filas muy viejas
    # que no tengan fecha_partido poblado.
    fecha_filtro = historial["fecha_partido"].fillna(historial["fecha"]) if "fecha_partido" in historial.columns else historial["fecha"]
    if desde_fecha:
        historial = historial[fecha_filtro >= desde_fecha]
        fecha_filtro = fecha_filtro[fecha_filtro >= desde_fecha]
    if hasta_fecha:
        historial = historial[fecha_filtro <= hasta_fecha]

    if historial.empty:
        return vacio

    mercados = {
        "BTTS": _market_stats(historial, "senal_btts", "resultado_btts", "profit_btts", "cuota_btts", stake),
        "Over 1.5": _market_stats(historial, "senal_o15", "resultado_o15", "profit_o15", "cuota_o15", stake),
        "Over 2.5": _market_stats(historial, "senal_o25", "resultado_o25", "profit_o25", "cuota_o25", stake),
        "BTTS + Over 2.5 (VIP)": _combo_stats(historial, "VIP", stake),
        "Over 0.5 1er tiempo": _market_stats(historial, "sug_ht05", "resultado_ht05", "profit_ht05", "cuota_ht_over05", stake),
    }

    categorias = {
        "Stake10": _tier_stats(historial, "senal_btts_optimizada", "Stake10", "resultado_btts"),
        "VIP+": _tier_stats(historial, "senal_btts_optimizada", "VIP+", "resultado_btts"),
        "VIP": _tier_stats(historial, "senal_btts_optimizada", "VIP", "resultado_btts"),
        "Fuerte": _tier_stats(historial, "senal_btts_optimizada", "Fuerte", "resultado_btts"),
        "ORO_FUERTE": _tier_stats(historial, "oro", "ORO_FUERTE", "resultado_btts"),
        "ORO_BTTS": _tier_stats(historial, "oro", "ORO_BTTS", "resultado_btts"),
        "ÉLITE": _tier_stats(historial, "nivel_confianza", "ÉLITE", "resultado_btts"),
        "PLATA": _tier_stats(historial, "nivel_confianza", "PLATA", "resultado_btts"),
        "BRONCE": _tier_stats(historial, "nivel_confianza", "BRONCE", "resultado_btts"),
        "VIP ELITE": _tier_stats(historial, "vip_elite", "VIP ELITE", "resultado_btts"),
    }

    return {
        "mercados": mercados,
        "categorias": categorias,
        "motivos": _motivo_stats(historial, "senal_btts", "motivo_btts", "resultado_btts", "profit_btts_hipotetico"),
        "motivos_o25": _motivo_stats(historial, "senal_o25", "motivo_o25", "resultado_o25", "profit_o25"),
        "motivos_ht05": _motivo_stats(historial, "sug_ht05", "motivo_ht05", "resultado_ht05", "profit_ht05"),
        "motivos_combo": _motivo_stats(historial, "senal_btts_o25", "motivo_combo", "resultado_combo", "profit_combo", sin_motivo_valor="ninguno_evitar_combo"),
        "nota_btts_min_alto": _nota_btts_min_alto_stats(historial),
        "marcador_exacto": _marcador_exacto_stats(historial),
        "calibracion": {
            "BTTS": _calibracion_stats(historial, "prob_btts", "resultado_btts"),
            "Over 1.5": _calibracion_stats(historial, "prob_o15", "resultado_o15"),
            "Over 2.5": _calibracion_stats(historial, "prob_o25", "resultado_o25"),
            "Over 0.5 1er tiempo": _calibracion_stats(historial, "prob_over05_ht", "resultado_ht05"),
        },
        "por_liga": _liga_stats(historial),
        "movimiento_validacion": _movimiento_stats(historial),
        "total_partidos_evaluados": int(historial["resultado_btts"].notna().sum()),
    }


def _hoy() -> datetime:
    return datetime.now(ZoneInfo(config.TIMEZONE))


def build_summary_comparativo(stake: float | None = None) -> dict:
    hoy = _hoy().strftime("%Y-%m-%d")
    hace_7_dias = (_hoy() - timedelta(days=7)).strftime("%Y-%m-%d")
    return {
        "hoy": build_summary(stake, desde_fecha=hoy),
        "recientes": build_summary(stake, desde_fecha=hace_7_dias),
        "historico": build_summary(stake),
    }


def get_category_accuracy(tier_col: str, tier_value: str) -> dict:
    historial = _load()
    if historial.empty or tier_col not in historial.columns:
        return {"total": 0, "ganados": 0, "acierto_pct": None}
    return _tier_stats(historial, tier_col, tier_value, "resultado_btts")


def format_summary_comparativo_message(comparativo: dict) -> str:
    hoy, recientes, historico = comparativo["hoy"], comparativo["recientes"], comparativo["historico"]

    if historico["total_partidos_evaluados"] == 0:
        return ""

    def pct_txt(s: dict, con_n: bool = False) -> str:
        if not s or s.get("bets", s.get("total", 0)) == 0:
            return "—"
        n = s.get("bets", s.get("total", 0))
        wins = s.get("wins", s.get("ganados", 0))
        if con_n:
            return f"{s['acierto_pct']:.0f}% ({wins}/{n})"
        return f"{s['acierto_pct']:.0f}%"

    def semaforo(s_historico: dict) -> str:
        if not s_historico or s_historico.get("bets", s_historico.get("total", 0)) == 0:
            return "⚪"
        return "🟢" if s_historico.get("profit_u", 0) >= 0 else "🔴"

    def alerta(s_reciente: dict, s_historico: dict) -> str:
        ar = s_reciente.get("acierto_pct") if s_reciente else None
        ah = s_historico.get("acierto_pct") if s_historico else None
        n_reciente = s_reciente.get("bets", s_reciente.get("total", 0)) if s_reciente else 0
        if ar is not None and ah is not None and n_reciente >= 5 and (ah - ar) >= 10:
            return " ⚠️"
        return ""

    def unidades_txt(s: dict) -> str:
        if not s or s.get("bets", s.get("total", 0)) == 0:
            return ""
        return f"  ({s.get('profit_u', 0):+.1f}u)"

    lineas = [
        "📊 <b>RESUMEN DE SEÑALES</b>",
        f"<i>{historico['total_partidos_evaluados']} partidos evaluados en el histórico total</i>",
        "━━━━━━━━━━━━━━━━━━",
        "<b>Por mercado</b>  (hoy · 7 días · histórico)",
    ]
    for nombre in historico["mercados"]:
        sh, sr, st = hoy["mercados"].get(nombre, {}), recientes["mercados"].get(nombre, {}), historico["mercados"][nombre]
        lineas.append(f"{semaforo(st)} <b>{nombre}</b>{alerta(sr, st)}: {pct_txt(sh)} · {pct_txt(sr)} · <b>{pct_txt(st, con_n=True)}</b>{unidades_txt(st)}")

    categorias_con_datos = {k: v for k, v in historico["categorias"].items() if v.get("total_detectado", v["total"]) > 0}
    if categorias_con_datos:
        lineas.append("━━━━━━━━━━━━━━━━━━")
        lineas.append("<b>Por categoría BTTS</b>  (hoy · 7 días · histórico)")
        for nombre in categorias_con_datos:
            sh = hoy["categorias"].get(nombre, {})
            sr = recientes["categorias"].get(nombre, {})
            st = historico["categorias"][nombre]
            pendientes = st.get("total_detectado", st["total"]) - st["total"]
            nota_pendientes = f" (+{pendientes} pendiente{'s' if pendientes != 1 else ''})" if pendientes > 0 else ""
            if st["total"] == 0:
                lineas.append(f"⚪ {nombre}: sin resultado todavía{nota_pendientes}")
            else:
                lineas.append(f"{semaforo(st)} {nombre}{alerta(sr, st)}: {pct_txt(sh)} · {pct_txt(sr)} · <b>{pct_txt(st, con_n=True)}</b>{unidades_txt(st)}{nota_pendientes}")

    def _seccion_motivos(titulo, key):
        motivos_hist = historico.get(key, {})
        if not motivos_hist:
            return
        lineas.append("━━━━━━━━━━━━━━━━━━")
        lineas.append(f"<b>{titulo}</b>  (7 días · histórico)")
        for motivo in sorted(motivos_hist, key=lambda m: (motivos_hist[m]["acierto_pct"] is None, motivos_hist[m]["acierto_pct"] or 0)):
            sr = recientes.get(key, {}).get(motivo, {})
            st = motivos_hist[motivo]
            pendientes = st.get("total_detectado", st["total"]) - st["total"]
            nota_pendientes = f" (+{pendientes} pendiente{'s' if pendientes != 1 else ''})" if pendientes > 0 else ""
            if st["total"] == 0:
                lineas.append(f"⚪ {motivo}: sin resultado todavía{nota_pendientes}")
            else:
                lineas.append(f"{semaforo(st)} {motivo}{alerta(sr, st)}: {pct_txt(sr)} · <b>{pct_txt(st, con_n=True)}</b>{unidades_txt(st)}{nota_pendientes}")

    _seccion_motivos("Por qué se activó JUGAR (BTTS)", "motivos")
    _seccion_motivos("Por qué se activó JUGAR (Over 2.5)", "motivos_o25")
    _seccion_motivos("Por qué se activó JUGAR (1er tiempo O0.5)", "motivos_ht05")
    _seccion_motivos("Por qué se activó JUGAR (Combinada BTTS+O2.5)", "motivos_combo")

    me = historico.get("marcador_exacto", {})
    if me.get("total", 0) > 0:
        lineas.append("━━━━━━━━━━━━━━━━━━")
        lineas.append(f"<b>🎯 Marcador exacto</b> (sin cuota confirmada aún — solo %acierto)")
        lineas.append(f"▫️ {me['aciertos']}/{me['total']} (<b>{me['acierto_pct']:.1f}%</b>)")

    calibracion = historico.get("calibracion", {})
    tiene_calibracion = any(len(v) > 0 for v in calibracion.values())
    if tiene_calibracion:
        lineas.append("━━━━━━━━━━━━━━━━━━")
        lineas.append("<b>📐 Calibración</b> (probabilidad dicha vs acierto real)")
        for mercado, buckets in calibracion.items():
            if not buckets:
                continue
            lineas.append(f"<i>{mercado}:</i>")
            for b in buckets:
                diff = b["diferencia"]
                marca = "✅" if abs(diff) <= 5 else ("⚠️" if diff < -5 else "💪")
                lineas.append(f"  {marca} dijimos {b['rango']} → acertó {b['acierto_pct']:.0f}% ({b['aciertos']}/{b['total']})")

    por_liga = historico.get("por_liga", [])
    if por_liga:
        lineas.append("━━━━━━━━━━━━━━━━━━")
        lineas.append("<b>🌍 Por liga (BTTS)</b> — peor a mejor")
        for l in por_liga[:8]:
            cls = "🔴" if l["profit_u"] < 0 else "🟢"
            lineas.append(f"  {cls} {l['liga']}: {l['acierto_pct']:.0f}% ({l['ganados']}/{l['total']})")
        if len(por_liga) > 8:
            lineas.append(f"  <i>... y {len(por_liga)-8} ligas más con suficiente muestra</i>")

    mv = historico.get("movimiento_validacion", {})
    con_mov, sin_mov = mv.get("con_movimiento", {}), mv.get("sin_movimiento", {})
    if con_mov.get("total", 0) > 0 or sin_mov.get("total", 0) > 0:
        lineas.append("━━━━━━━━━━━━━━━━━━")
        lineas.append("<b>📉 ¿Sirve la señal de movimiento?</b> (BTTS+O1.5+O2.5)")
        if con_mov.get("total", 0) > 0:
            lineas.append(f"  Con +4pp de movimiento: {con_mov['acierto_pct']:.0f}% ({con_mov['aciertos']}/{con_mov['total']})")
        else:
            lineas.append(f"  Con +4pp de movimiento: sin casos todavía")
        if sin_mov.get("total", 0) > 0:
            lineas.append(f"  Sin ese movimiento: {sin_mov['acierto_pct']:.0f}% ({sin_mov['aciertos']}/{sin_mov['total']})")
        else:
            lineas.append(f"  Sin ese movimiento: sin casos todavía")

    lineas.append("━━━━━━━━━━━━━━━━━━")
    lineas.append("<i>🟢 = viene ganando unidades · 🔴 = viene perdiendo · ⚠️ = caída fuerte reciente (5+ casos). Con pocos partidos (n bajo) un solo resultado mueve el % mucho — no saques conclusiones todavía con menos de ~15-20 casos.</i>")

    return "\n".join(lineas)


def _serie_diaria(df: pd.DataFrame) -> list[dict]:
    jugados = df[(df["senal_btts"] == "JUGAR") & (df["resultado_btts"].notna())]
    if jugados.empty:
        return []
    agrupado = jugados.groupby("fecha").agg(
        bets=("resultado_btts", "count"),
        wins=("resultado_btts", "sum"),
        profit_u=("profit_btts", "sum"),
    ).reset_index()
    agrupado["acierto_pct"] = (agrupado["wins"] / agrupado["bets"] * 100).round(1)
    return agrupado.sort_values("fecha").to_dict(orient="records")


def _v(x):
    if x is None:
        return None
    try:
        if pd.isna(x):
            return None
    except (TypeError, ValueError):
        pass
    return x


def _sanitizar_nan(obj):
    if isinstance(obj, dict):
        return {k: _sanitizar_nan(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_sanitizar_nan(v) for v in obj]
    try:
        if pd.isna(obj):
            return None
    except (TypeError, ValueError):
        pass
    return obj


def export_dashboard_json(path: str, stake: float | None = None) -> None:
    historial = _load()
    comparativo = build_summary_comparativo(stake)

    # Mapa liga -> "alta"/"baja" (o ausente si es neutral), calculado en
    # vivo a partir del ROI historico real de esa liga. Se usa para
    # marcar cada partido individual con la nota de confianza dinamica.
    confianza_por_liga = {
        l["liga"]: l["confianza"]
        for l in comparativo["historico"].get("por_liga", [])
        if l.get("confianza")
    }

    hoy_str = _hoy().strftime("%Y-%m-%d")
    manana_str = (_hoy() + timedelta(days=1)).strftime("%Y-%m-%d")
    UMBRAL_MOVIMIENTO_PP = 4.0

    def _movimiento_prob(original, actual, señal) -> float | None:
        if señal != "JUGAR":
            return None
        if original is None or actual is None:
            return None
        try:
            if pd.isna(original) or pd.isna(actual) or original == 0 or actual == 0:
                return None
        except TypeError:
            return None
        return (1 / actual - 1 / original) * 100

    fecha_efectiva = historial["fecha_partido"].fillna(historial["fecha"]) if "fecha_partido" in historial.columns else historial["fecha"]
    de_hoy = historial[fecha_efectiva.isin([hoy_str, manana_str])]
    picks_hoy = []
    for _, row in de_hoy.sort_values("match_id").iterrows():
        mov_btts = _movimiento_prob(row.get("cuota_btts"), row.get("cuota_btts_actual"), row.get("senal_btts"))
        mov_o15 = _movimiento_prob(row.get("cuota_o15"), row.get("cuota_o15_actual"), row.get("senal_o15"))
        mov_o25 = _movimiento_prob(row.get("cuota_o25"), row.get("cuota_o25_actual"), row.get("senal_o25"))
        picks_hoy.append({
            "fecha": _v(row.get("fecha")),
            "fecha_partido": _v(row.get("fecha_partido")) or _v(row.get("fecha")),
            "liga": _v(row.get("liga")),
            "liga_confianza": confianza_por_liga.get(row.get("liga")),
            "local": _v(row.get("local")),
            "visitante": _v(row.get("visitante")),
            "senal_btts_optimizada": _v(row.get("senal_btts_optimizada")),
            "nivel_final": _v(row.get("senal_btts")),
            "motivo_btts": _v(row.get("motivo_btts")),
            "prob_btts": _v(row.get("prob_btts")),
            "btts_min_alto": _v(row.get("btts_min_alto")),
            # Preferimos cuota_btts_actual (se actualiza en cada corrida)
            # sobre cuota_btts (congelada desde la primera deteccion). Si
            # un partido se detecto ANTES de que FootyStats publicara su
            # cuota de BTTS (comun en partidos de mañana o ligas chicas),
            # cuota_btts queda en 0 para siempre y el dashboard mostraba
            # "—" aunque la cuota ya estuviera disponible en corridas
            # posteriores - esto lo corrige, sin tocar el calculo de
            # ganancia/ROI (que sigue usando la cuota original real).
            "cuota_btts": _v(row.get("cuota_btts_actual")) or _v(row.get("cuota_btts")),
            "cuota_btts_actual": _v(row.get("cuota_btts_actual")),
            "movimiento_btts_pp": _v(mov_btts),
            "resultado_btts": _v(row.get("resultado_btts")),
            "senal_o15_tier": _v(row.get("tier_o15")),
            "sug_o15": _v(row.get("senal_o15")),
            "prob_o15": _v(row.get("prob_o15")),
            "cuota_o15": _v(row.get("cuota_o15")),
            "cuota_o15_actual": _v(row.get("cuota_o15_actual")),
            "movimiento_o15_pp": _v(mov_o15),
            "resultado_o15": _v(row.get("resultado_o15")),
            "senal_o25": _v(row.get("tier_o25")),
            "sug_o25": _v(row.get("senal_o25")),
            "prob_o25": _v(row.get("prob_o25")),
            "cuota_o25": _v(row.get("cuota_o25")),
            "cuota_o25_actual": _v(row.get("cuota_o25_actual")),
            "movimiento_o25_pp": _v(mov_o25),
            "resultado_o25": _v(row.get("resultado_o25")),
            "senal_btts_o25": _v(row.get("senal_btts_o25")),
            "motivo_combo": _v(row.get("motivo_combo")),
            "btts_min_alto": _v(row.get("btts_min_alto")),
            "resultado_combo": _v(row.get("resultado_combo")),
            "vip_elite": _v(row.get("vip_elite")),
            "oro": _v(row.get("oro")),
            "nivel_confianza": _v(row.get("nivel_confianza")),
            "marcador_exacto_sugerido": _v(row.get("marcador_exacto_sugerido")),
            "marcador_exacto_real": _v(row.get("marcador_exacto_real")),
            "usa_temporada_anterior": _v(row.get("usa_temporada_anterior")),
            "usa_promedio_general": _v(row.get("usa_promedio_general")),
            "es_equipo_filial": _v(row.get("es_equipo_filial")),
            "senal_ht05_tier": _v(row.get("senal_ht05")),
            "sug_ht05": _v(row.get("sug_ht05")),
            "prob_over05_ht": _v(row.get("prob_over05_ht")),
            "prob_over15_ht": _v(row.get("prob_over15_ht")),
            "cuota_ht_over05": _v(row.get("cuota_ht_over05")),
            "cuota_ht_over05_actual": _v(row.get("cuota_ht_over05_actual")),
            "movimiento_ht05_pp": _v(_movimiento_prob(row.get("cuota_ht_over05"), row.get("cuota_ht_over05_actual"), row.get("sug_ht05"))),
            "resultado_ht05": _v(row.get("resultado_ht05")),
        })

    fecha_orden = historial["fecha_partido"].fillna(historial["fecha"]) if "fecha_partido" in historial.columns else historial["fecha"]
    ultimos = historial.loc[fecha_orden.sort_values(ascending=False).index].head(200)
    picks_recientes = []
    for _, row in ultimos.iterrows():
        cuota_combo = None
        if pd.notna(row.get("cuota_btts")) and pd.notna(row.get("cuota_o25")):
            cuota_producto = row.get("cuota_btts") * row.get("cuota_o25")
            cuota_combo = 1 + (cuota_producto - 1) * (1 - DESCUENTO_CORRELACION_COMBO)
        picks_recientes.append({
            "fecha": _v(row.get("fecha")),
            "fecha_partido": _v(row.get("fecha_partido")) or _v(row.get("fecha")),
            "liga": _v(row.get("liga")),
            "liga_confianza": confianza_por_liga.get(row.get("liga")),
            "local": _v(row.get("local")),
            "visitante": _v(row.get("visitante")),
            "senal_btts_optimizada": _v(row.get("senal_btts_optimizada")),
            "nivel_final": _v(row.get("senal_btts")),
            "motivo_btts": _v(row.get("motivo_btts")),
            # Igual que en picks_hoy, PERO solo mientras el partido sigue
            # pendiente (sin resultado) - si ya se jugo, se muestra la
            # cuota ORIGINAL (la misma que se uso para calcular la
            # ganancia real), para que no se vea distinta a la que
            # realmente genero el profit_btts guardado.
            "cuota_btts": (
                (_v(row.get("cuota_btts_actual")) or _v(row.get("cuota_btts")))
                if pd.isna(row.get("resultado_btts"))
                else _v(row.get("cuota_btts"))
            ),
            "resultado_btts": _v(row.get("resultado_btts")),
            "movimiento_btts_pp": _v(_movimiento_prob(row.get("cuota_btts"), row.get("cuota_btts_actual"), row.get("senal_btts"))),
            "senal_o15_tier": _v(row.get("tier_o15")),
            "sug_o15": _v(row.get("senal_o15")),
            "cuota_o15": _v(row.get("cuota_o15")),
            "resultado_o15": _v(row.get("resultado_o15")),
            "movimiento_o15_pp": _v(_movimiento_prob(row.get("cuota_o15"), row.get("cuota_o15_actual"), row.get("senal_o15"))),
            "senal_o25": _v(row.get("tier_o25")),
            "sug_o25": _v(row.get("senal_o25")),
            "cuota_o25": _v(row.get("cuota_o25")),
            "resultado_o25": _v(row.get("resultado_o25")),
            "movimiento_o25_pp": _v(_movimiento_prob(row.get("cuota_o25"), row.get("cuota_o25_actual"), row.get("senal_o25"))),
            "senal_btts_o25": _v(row.get("senal_btts_o25")),
            "motivo_combo": _v(row.get("motivo_combo")),
            "btts_min_alto": _v(row.get("btts_min_alto")),
            "cuota_combo": _v(cuota_combo),
            "resultado_combo": _v(row.get("resultado_combo")),
            "senal_ht05_tier": _v(row.get("senal_ht05")),
            "sug_ht05": _v(row.get("sug_ht05")),
            "cuota_ht_over05": _v(row.get("cuota_ht_over05")),
            "resultado_ht05": _v(row.get("resultado_ht05")),
        })

    data = {
        "actualizado": _hoy().isoformat(),
        "hoy_referencia": hoy_str,
        "umbral_movimiento_pp": UMBRAL_MOVIMIENTO_PP,
        "confianza_por_liga": confianza_por_liga,
        "comparativo": comparativo,
        "serie_diaria": _serie_diaria(historial),
        "picks_hoy": picks_hoy,
        "picks_recientes": picks_recientes,
    }
    data = _sanitizar_nan(data)

    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str, allow_nan=False)


def format_daily_summary_message(summary: dict) -> str:
    lineas = [
        "📊 <b>RESUMEN DE SEÑALES</b>",
        f"<i>{summary['total_partidos_evaluados']} partidos evaluados en total</i>",
        "━━━━━━━━━━━━━━━━━━",
        "<b>Por mercado</b>",
    ]
    for nombre, s in summary["mercados"].items():
        if s["bets"] == 0:
            lineas.append(f"▫️ {nombre}: <i>sin datos aún</i>")
            continue
        signo = "🟢" if s["profit_u"] >= 0 else "🔴"
        lineas.append(
            f"▫️ <b>{nombre}</b>  ·  {s['wins']}/{s['bets']} aciertos (<b>{s['acierto_pct']:.1f}%</b>)\n"
            f"    {signo} {s['profit_u']:+.2f}u  ·  ROI {s['roi_pct']:.1f}%  ·  cuota prom. {s['cuota_prom']:.2f}"
        )

    categorias_con_datos = {k: v for k, v in summary["categorias"].items() if v["total"] > 0}
    if categorias_con_datos:
        lineas.append("━━━━━━━━━━━━━━━━━━")
        lineas.append("<b>Por categoría (BTTS)</b>")
        for nombre, s in categorias_con_datos.items():
            lineas.append(f"▫️ {nombre}: {s['ganados']}/{s['total']} (<b>{s['acierto_pct']:.1f}%</b>)")

    motivos = summary.get("motivos", {})
    if motivos:
        lineas.append("━━━━━━━━━━━━━━━━━━")
        lineas.append("<b>Por qué se activó JUGAR (BTTS)</b>")
        for motivo, s in sorted(motivos.items(), key=lambda kv: kv[1]["acierto_pct"]):
            alerta = "⚠️ " if s["acierto_pct"] < 50 else ""
            lineas.append(f"▫️ {alerta}{motivo}: {s['ganados']}/{s['total']} (<b>{s['acierto_pct']:.1f}%</b>)")

    return "\n".join(lineas)
