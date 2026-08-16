"""
footystats_client.py
=====================
Cliente minimo para la API de FootyStats (https://footystats.org/api).

Usa 3 endpoints:
  - league-teams   -> estadisticas de temporada por equipo (xG, xGA, %fallar
                       en anotar, %valla invicta, tiros a puerta, BTTS%)
  - todays-matches -> partidos del dia (o de una fecha dada) con cuotas
  - league-list     -> util una sola vez para encontrar los season_id de tus ligas

IMPORTANTE: los nombres de campo de FootyStats para promedios por
local/visitante (xg_for_avg_home, seasonFTSPercentage_away, etc.) estan
tomados de la documentacion oficial, pero FootyStats a veces ajusta
nombres entre paquetes/planes. Antes de la primera corrida real, ejecuta
`python footystats_client.py --debug-team <team_id>` y confirma que las
claves existen en la respuesta cruda; si algun nombre cambio, ajustalo en
TEAM_FIELD_MAP abajo (es la unica seccion que deberias tocar).
"""
from __future__ import annotations
import requests
from typing import Optional

BASE_URL = "https://api.football-data-api.com"

# Mapeo "columna del Excel" -> "campo de la API para el equipo LOCAL / VISITANTE / TOTAL TEMPORADA"
# Ajusta aqui si FootyStats cambia algun nombre de campo.
TEAM_FIELD_MAP = {
    "xg_home":            "xg_for_avg_home",
    "xg_away":             "xg_for_avg_away",
    "xg_overall":          "xg_for_avg_overall",
    "xga_home":           "xg_against_avg_home",
    "xga_away":            "xg_against_avg_away",
    "xga_overall":         "xg_against_avg_overall",
    "gf_home":             "seasonScoredAVG_home",
    "gf_away":              "seasonScoredAVG_away",
    "gf_overall":           "seasonScoredAVG_overall",
    "gc_home":             "seasonConcededAVG_home",
    "gc_away":              "seasonConcededAVG_away",
    "gc_overall":           "seasonConcededAVG_overall",
    "failed_home":        "seasonFTSPercentage_home",
    "failed_away":         "seasonFTSPercentage_away",
    "failed_overall":      "seasonFTSPercentage_overall",
    "cs_home":            "seasonCSPercentage_home",
    "cs_away":             "seasonCSPercentage_away",
    "cs_overall":          "seasonCSPercentage_overall",
    "sot_home":           "shotsOnTargetAVG_home",
    "sot_away":            "shotsOnTargetAVG_away",
    "sot_overall":         "shotsOnTargetAVG_overall",
    "btts_home":          "seasonBTTSPercentage_home",
    "btts_away":           "seasonBTTSPercentage_away",
    "btts_overall":        "seasonBTTSPercentage_overall",
    "pj_home":            "seasonMatchesPlayed_home",
    "pj_away":             "seasonMatchesPlayed_away",
    "pj_overall":          "seasonMatchesPlayed_overall",
    # --- Candidatas nuevas (13 ago 2026) para BTTS, aun SIN validar con
    # datos reales - se agregan solo para EMPEZAR A CAPTURARLAS desde
    # ahora. No se usan todavia en ninguna formula; hace falta acumular
    # 2-3 semanas de partidos reales antes de poder medir si aportan algo
    # con el mismo rigor (estabilidad partiendo por fecha) que el resto
    # del sistema. Ver comentario en signals.py donde se guardan.
    "shots_home":         "shotsAVG_home",          # tiros TOTALES (no solo a puerta)
    "shots_away":          "shotsAVG_away",
    "shots_overall":       "shotsAVG_overall",
    "ppg_home":           "seasonPPG_home",          # puntos por partido (fuerza del equipo en la tabla)
    "ppg_away":            "seasonPPG_away",
    "ppg_overall":         "seasonPPG_overall",
    "avght_home":         "AVGHT_home",              # goles promedio en el 1er tiempo (equipo)
    "avght_away":          "AVGHT_away",
    "avght_overall":       "AVGHT_overall",
    "avg2h_home":         "AVG_2hg_home",            # goles promedio en el 2do tiempo (equipo)
    "avg2h_away":          "AVG_2hg_away",
    "avg2h_overall":       "AVG_2hg_overall",
}

# Campos de primera mitad (HT), para el mercado "Over 0.5 goles 1er tiempo".
# Distinto de TEAM_FIELD_MAP porque FootyStats no da xG separado por mitad
# (solo existe para el partido completo) - aqui solo se usan promedios
# directos de goles anotados/recibidos en el primer tiempo.
TEAM_FIELD_MAP_HT = {
    "scored_ht_home":    "scoredAVGHT_home",
    "scored_ht_away":     "scoredAVGHT_away",
    "conceded_ht_home":  "concededAVGHT_home",
    "conceded_ht_away":   "concededAVGHT_away",
    "over05_ht_pct_home": "seasonOver05PercentageHT_home",
    "over05_ht_pct_away":  "seasonOver05PercentageHT_away",
}


class FootyStatsClient:
    def __init__(self, api_key: str, timeout: int = 20):
        if not api_key:
            raise ValueError("Falta FOOTYSTATS_API_KEY (revisa tu archivo .env)")
        self.api_key = api_key
        self.timeout = timeout

    def _get(self, endpoint: str, params: dict) -> dict:
        params = {**params, "key": self.api_key}
        resp = requests.get(f"{BASE_URL}/{endpoint}", params=params, timeout=self.timeout, headers={"Expect": ""})
        resp.raise_for_status()
        data = resp.json()
        if not data.get("success", True):
            raise RuntimeError(f"FootyStats devolvio error en {endpoint}: {data}")
        return data

    def league_list(self) -> list[dict]:
        return self._get("league-list", {}).get("data", [])

    def league_teams(self, season_id: int) -> list[dict]:
        """Estadisticas de temporada de todos los equipos de una liga."""
        return self._get("league-teams", {"season_id": season_id, "include": "stats"}).get("data", [])

    def league_matches(self, season_id: int, max_per_page: int = 1000) -> list[dict]:
        """
        Calendario completo de una liga (pasados y futuros), con resultado
        final (homeGoalCount/awayGoalCount/status) para los que ya se jugaron.
        Usado por backtest.py para completar resultados.
        """
        return self._get("league-matches", {"season_id": season_id, "max_per_page": max_per_page}).get("data", [])

    def todays_matches(self, date: Optional[str] = None, timezone: str = "Etc/UTC") -> list[dict]:
        """
        Partidos de una fecha (YYYY-MM-DD). Si date es None, usa el dia
        actual en UTC. Requiere que las ligas esten activadas en tu cuenta
        de FootyStats (Ajustes de API).

        IMPORTANTE: este endpoint de FootyStats esta paginado - devuelve
        MAXIMO 200 partidos por pagina (confirmado en su documentacion
        oficial). Con muchas ligas activas a la vez, un solo dia puede
        facilmente superar los 200 partidos totales entre todas las ligas.
        Antes esta funcion solo pedia la primera pagina, asi que cualquier
        partido que cayera despues del puesto 200 (en el orden que
        FootyStats decida, no necesariamente por liga o importancia) se
        perdia en silencio - esto causaba que ligas completas parecieran
        "no estar activas" cuando en realidad si tenian partidos ese dia,
        solo que en una pagina que nunca se pedia. Ahora se piden todas
        las paginas hasta que no queden mas resultados.
        """
        params = {"timezone": timezone}
        if date:
            params["date"] = date

        todos = []
        pagina = 1
        while True:
            params_pagina = {**params, "page": pagina}
            data = self._get("todays-matches", params_pagina)
            lote = data.get("data", [])
            if not lote:
                break
            todos.extend(lote)
            # Si esta pagina vino con menos de 200 (el maximo documentado),
            # ya no hay mas paginas que pedir - evita una llamada de mas.
            if len(lote) < 200:
                break
            pagina += 1
            if pagina > 20:  # limite de seguridad, no deberia hacer falta nunca
                break
        return todos

    def team_stats_row(self, team: dict, is_home: bool, forzar_fuente: str | None = None) -> dict:
        """
        Extrae del objeto 'team' (tal como viene de league-teams) los
        valores que necesita el modelo. Los datos reales vienen anidados
        dentro de team['stats'].

        forzar_fuente:
          None (por defecto) -> automatico: usa el lado (home/away) si
              tiene al menos 1 partido jugado ahi, si no usa 'overall'
              (evita el 0 falso que da FootyStats cuando el lado
              especifico no tiene partidos todavia).
          "lado"    -> fuerza usar SIEMPRE el lado especifico (home/away),
              sin importar cuantos partidos tenga (puede venir en 0/None).
          "overall" -> fuerza usar SIEMPRE el promedio general de toda la
              temporada, sin importar el lado.

        Devuelve un dict con: xg, xga, gf, gc, failed, cs, sot, btts, pj
        (partidos jugados de la fuente finalmente usada), pj_lado
        (partidos jugados de ESE lado especifico, crudo), pj_overall
        (partidos jugados en toda la temporada, crudo), fuente_lado.
        """
        stats = team.get("stats", {}) or {}
        side = "home" if is_home else "away"
        bases = {"xg", "xga", "gf", "gc", "failed", "cs", "sot", "btts",
                 "shots", "ppg", "avght", "avg2h"}

        pj_lado = stats.get(TEAM_FIELD_MAP[f"pj_{side}"])
        pj_overall = stats.get(TEAM_FIELD_MAP["pj_overall"])

        if forzar_fuente == "overall":
            fuente = "overall"
        elif forzar_fuente == "lado":
            fuente = side
        else:
            fuente = "overall" if (pj_lado is None or pj_lado == 0) else side

        out = {}
        for base in bases:
            api_field = TEAM_FIELD_MAP[f"{base}_{fuente}"]
            out[base] = stats.get(api_field)
        out["pj"] = pj_overall if fuente == "overall" else pj_lado
        out["pj_lado"] = pj_lado
        out["pj_overall"] = pj_overall
        out["fuente_lado"] = fuente

        # over25_and_btts_percentage viene ANIDADO dentro de
        # stats['additional_info'], no como campo directo (a diferencia de
        # todos los demas campos de este mapa) - se extrae aparte, usando
        # la misma "fuente" (lado/overall) ya decidida arriba para que sea
        # consistente con el resto de los datos del equipo.
        info_extra = stats.get("additional_info", {}) or {}
        out["over25_and_btts_pct"] = info_extra.get(f"over25_and_btts_percentage_{fuente}")

        return out

    def team_stats_ht_row(self, team: dict, is_home: bool) -> dict:
        """
        Extrae las estadisticas de PRIMERA MITAD del equipo (goles anotados
        y recibidos en promedio antes del descanso, y % de partidos con
        Over 0.5 goles en el primer tiempo). Version simplificada (sin el
        respaldo de 3 niveles que usa team_stats_row) porque es un mercado
        nuevo que todavia no tiene suficiente historial para saber si ese
        mismo nivel de respaldo tiene sentido aqui tambien - se revisa mas
        adelante con datos reales.
        """
        stats = team.get("stats", {}) or {}
        side = "home" if is_home else "away"
        return {
            "scored_ht": stats.get(TEAM_FIELD_MAP_HT[f"scored_ht_{side}"]),
            "conceded_ht": stats.get(TEAM_FIELD_MAP_HT[f"conceded_ht_{side}"]),
            "over05_ht_pct": stats.get(TEAM_FIELD_MAP_HT[f"over05_ht_pct_{side}"]),
        }
