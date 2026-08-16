"""
main.py
=======
Orquesta todo el pipeline:

  1. Descarga los partidos (de un dia especifico, o de una ventana movil de
     las proximas 24 horas) desde FootyStats.
  2. Para cada liga involucrada, descarga las estadisticas de temporada de sus
     equipos (con cache en memoria para no repetir llamadas).
  3. Arma el MatchInput de cada partido y calcula las señales (signals.py).
  4. Guarda todo en un CSV/Excel dentro de OUTPUT_DIR.
  5. Envia una alerta de Telegram por cada partido NUEVO cuya señal final sea
     "JUGAR" (columnas configurables en config.NOTIFY_ON_COLUMNS). "Nuevo"
     quiere decir que nunca se habia notificado antes en el historial, asi
     que correr esto varias veces al dia (modo --ventana-24h) no reenvia
     alertas duplicadas del mismo partido.

Uso:
    python main.py                        # partidos de HOY (una sola vez)
    python main.py --date 2026-07-20      # partidos de una fecha especifica
    python main.py --dias-adelante 1      # partidos de mañana
    python main.py --ventana-24h          # partidos entre AHORA y +24h (pensado para correr cada 1-2h)
    python main.py --ventana-24h --sin-resumen   # igual, pero sin mandar el resumen de backtest (para no repetirlo cada hora)
"""
from __future__ import annotations
import argparse
import os
import re
from datetime import datetime, timezone, timedelta
from zoneinfo import ZoneInfo

import pandas as pd

import config
from footystats_client import FootyStatsClient
from signals import MatchInput, calculate_signals, calculate_marcador_exacto
from telegram_bot import send_telegram_message, format_match_alert
import backtest
import reglas_activas

PATRON_EQUIPO_FILIAL = re.compile(
    r"(\bII\b$| III$|Reserve|Reserves|Next Pro|\bU2[0-9]\b|\bU1[0-9]\b|Youth|Sub-2[0-9]|Sub 2[0-9])",
    re.IGNORECASE,
)


def es_equipo_filial(nombre: str | None) -> bool:
    """
    Detecta equipos filiales/reserva/desarrollo (ej: 'Sporting Kansas City II',
    'New England Revolution II', equipos de 'Next Pro', sub-21/23, etc.) por
    el nombre. Estos equipos tienen planteles que rotan mucho semana a semana,
    asi que el promedio de temporada de FootyStats para ellos predice mucho
    peor que para un equipo de primera plantilla estable - confirmado con
    datos reales del backtest (78% de acierto en equipos normales vs 20% en
    filiales, en la misma muestra). Por eso se excluyen de generar JUGAR,
    sin importar que diga el resto del modelo.
    """
    if not nombre:
        return False
    return bool(PATRON_EQUIPO_FILIAL.search(nombre))


def build_match_input(home_stats: dict, away_stats: dict, match: dict) -> MatchInput:
    """Combina stats de equipo (temporada) + cuotas del partido (match) en un MatchInput."""
    return MatchInput(
        xg_local=home_stats.get("xg"),
        xga_local=home_stats.get("xga"),
        xg_visitante=away_stats.get("xg"),
        xga_visitante=away_stats.get("xga"),
        gf_local=home_stats.get("gf"),
        gc_local=home_stats.get("gc"),
        gf_visitante=away_stats.get("gf"),
        gc_visitante=away_stats.get("gc"),
        failed_local_pct=home_stats.get("failed"),
        cs_local_pct=home_stats.get("cs"),
        failed_visit_pct=away_stats.get("failed"),
        cs_visit_pct=away_stats.get("cs"),
        sot_local=home_stats.get("sot"),
        sot_visitante=away_stats.get("sot"),
        cuota_btts=match.get("odds_btts_yes"),
        cuota_btts_no=match.get("odds_btts_no"),
        cuota_o15=match.get("odds_ft_over15"),
        cuota_o25=match.get("odds_ft_over25"),
        btts_pct_local=home_stats.get("btts"),
        btts_pct_visitante=away_stats.get("btts"),
        cuota_local=match.get("odds_ft_1"),
        cuota_visitante=match.get("odds_ft_2"),
        scored_ht_local=home_stats.get("scored_ht"),
        conceded_ht_local=home_stats.get("conceded_ht"),
        scored_ht_visitante=away_stats.get("scored_ht"),
        conceded_ht_visitante=away_stats.get("conceded_ht"),
        over05_ht_pct_local=home_stats.get("over05_ht_pct"),
        over05_ht_pct_visitante=away_stats.get("over05_ht_pct"),
        cuota_ht_over05=match.get("odds_1st_half_over05"),
        pj_local=home_stats.get("pj"),
        pj_visitante=away_stats.get("pj"),
        over25_and_btts_pct_local=home_stats.get("over25_and_btts_pct"),
        over25_and_btts_pct_visitante=away_stats.get("over25_and_btts_pct"),
        ppg_local=home_stats.get("ppg"),
        ppg_visitante=away_stats.get("ppg"),
        shots_local=home_stats.get("shots"),
        shots_visitante=away_stats.get("shots"),
        avght_local=home_stats.get("avght"),
        avght_visitante=away_stats.get("avght"),
        avg2h_local=home_stats.get("avg2h"),
        avg2h_visitante=away_stats.get("avg2h"),
    )


def fetch_matches_ventana_24h(client: FootyStatsClient) -> list[dict]:
    """
    Trae los partidos entre AHORA y +24 horas, combinando 'hoy' y 'mañana'
    (porque la ventana casi siempre cruza la medianoche) y filtrando por
    date_unix real. Pensado para correr varias veces al dia sin perder
    partidos que caen justo despues de medianoche.
    """
    tz = ZoneInfo(config.TIMEZONE)
    ahora = datetime.now(tz)
    hoy_str = ahora.strftime("%Y-%m-%d")
    manana_str = (ahora + timedelta(days=1)).strftime("%Y-%m-%d")

    crudos = client.todays_matches(date=hoy_str, timezone=config.TIMEZONE)
    crudos += client.todays_matches(date=manana_str, timezone=config.TIMEZONE)

    ahora_unix = ahora.timestamp()
    limite_unix = (ahora + timedelta(hours=24)).timestamp()

    vistos = set()
    resultado = []
    for m in crudos:
        mid = m.get("id")
        du = m.get("date_unix")
        if mid in vistos or du is None:
            continue
        if ahora_unix <= du <= limite_unix:
            vistos.add(mid)
            resultado.append(m)
    return resultado


def run(date: str | None = None, notify: bool = True, ventana_24h: bool = False, enviar_resumen: bool = True) -> pd.DataFrame:
    client = FootyStatsClient(config.FOOTYSTATS_API_KEY)

    print("Cargando nombres de ligas...")
    ligas = client.league_list()
    nombre_liga_por_season = {}
    for liga in ligas:
        nombre = liga.get("name", "")
        for temporada in liga.get("season", []):
            sid = temporada.get("id")
            if sid is not None:
                nombre_liga_por_season[sid] = nombre

    def _year_key(year_value):
        m = re.search(r"\d{4}", str(year_value))
        return int(m.group()) if m else 0

    temporada_anterior_por_season = {}
    for liga in ligas:
        temporadas = sorted(liga.get("season", []), key=lambda t: _year_key(t.get("year")))
        for i in range(1, len(temporadas)):
            actual_id = temporadas[i].get("id")
            anterior_id = temporadas[i - 1].get("id")
            if actual_id is not None and anterior_id is not None:
                temporada_anterior_por_season[actual_id] = anterior_id

    if ventana_24h:
        print("Descargando partidos entre ahora y +24h...")
        matches = fetch_matches_ventana_24h(client)
    else:
        print(f"Descargando partidos de {date or 'hoy'}...")
        matches = client.todays_matches(date=date, timezone=config.TIMEZONE)
    print(f"  -> {len(matches)} partidos encontrados")

    teams_cache: dict[int, dict[int, dict]] = {}

    def get_team(season_id: int, team_id: int) -> dict | None:
        if season_id not in teams_cache:
            teams = client.league_teams(season_id)
            teams_cache[season_id] = {t["id"]: t for t in teams}
        return teams_cache[season_id].get(team_id)

    def get_team_stats_con_respaldo(season_id: int, team_id: int, is_home: bool, team_obj: dict) -> dict:
        """
        Jerarquia de 3 niveles para decidir que datos usar (ver docstring
        original mas abajo en el historial del proyecto). Ademas, siempre
        agrega las estadisticas de PRIMERA MITAD (sin respaldo de 3 niveles
        todavia - se revisa mas adelante con datos reales si hace falta).
        """
        stats_lado = client.team_stats_row(team_obj, is_home=is_home, forzar_fuente="lado")
        pj_lado = stats_lado.get("pj_lado")

        stats_ht = client.team_stats_ht_row(team_obj, is_home=is_home)

        if pj_lado is not None and pj_lado >= config.MIN_PARTIDOS_JUGADOS:
            stats_lado["fuente_datos"] = "actual"
            stats_lado.update(stats_ht)
            return stats_lado

        pj_overall = stats_lado.get("pj_overall")
        if pj_overall is not None and pj_overall >= config.MIN_PARTIDOS_JUGADOS:
            stats_general = client.team_stats_row(team_obj, is_home=is_home, forzar_fuente="overall")
            stats_general["fuente_datos"] = "actual"
            stats_general.update(stats_ht)
            return stats_general

        prev_sid = temporada_anterior_por_season.get(season_id)
        if prev_sid:
            prev_team = get_team(prev_sid, team_id)
            if prev_team:
                prev_stats_lado = client.team_stats_row(prev_team, is_home=is_home, forzar_fuente="lado")
                prev_pj_lado = prev_stats_lado.get("pj_lado")
                if prev_pj_lado is not None and prev_pj_lado >= config.MIN_PARTIDOS_TEMPORADA_ANTERIOR:
                    prev_stats_lado["fuente_datos"] = "anterior"
                    prev_stats_lado.update(stats_ht)
                    return prev_stats_lado

        stats_lado["fuente_datos"] = "actual"
        stats_lado.update(stats_ht)
        return stats_lado

    # Que reglas de BTTS pueden generar JUGAR en esta corrida - se calcula
    # UNA sola vez (no por partido) leyendo el ROI real acumulado de cada
    # regla en el historial. Una regla desactivada se reactiva sola si
    # empieza a rendir bien, y viceversa (ver reglas_activas.py).
    try:
        _resumen_motivos = backtest.build_summary()["motivos"]
        REGLAS_ACTIVAS = reglas_activas.calcular_reglas_activas(_resumen_motivos)
        _activas = [k for k, v in REGLAS_ACTIVAS.items() if v]
        print(f"Reglas BTTS habilitadas esta corrida ({len(_activas)}): {', '.join(sorted(_activas))}")
    except Exception as e:
        print(f"[reglas] no se pudo calcular el estado dinamico ({e}), usando el estado inicial conocido")
        REGLAS_ACTIVAS = dict(reglas_activas.ESTADO_INICIAL)

    filas = []
    for match in matches:
        season_id = match.get("competition_id")
        home_id = match.get("homeID")
        away_id = match.get("awayID")
        if not season_id or not home_id or not away_id:
            continue

        home_team = get_team(season_id, home_id)
        away_team = get_team(season_id, away_id)
        if not home_team or not away_team:
            continue

        home_stats = get_team_stats_con_respaldo(season_id, home_id, is_home=True, team_obj=home_team)
        away_stats = get_team_stats_con_respaldo(season_id, away_id, is_home=False, team_obj=away_team)

        m_input = build_match_input(home_stats, away_stats, match)
        m_input.reglas_activas = REGLAS_ACTIVAS
        sig = calculate_signals(m_input)
        marcador = calculate_marcador_exacto(m_input)

        es_filial = es_equipo_filial(home_team.get("name")) or es_equipo_filial(away_team.get("name"))

        fecha_unix = match.get("date_unix")
        fecha_partido_dt = (datetime.fromtimestamp(fecha_unix, tz=timezone.utc).astimezone(ZoneInfo(config.TIMEZONE))
                             if fecha_unix else None)
        fecha = fecha_partido_dt.strftime("%Y-%m-%d %H:%M") if fecha_partido_dt else ""
        fecha_partido = fecha_partido_dt.strftime("%Y-%m-%d") if fecha_partido_dt else ""

        fila = {
            "fecha": fecha,
            "fecha_partido": fecha_partido,
            "match_id": match.get("id"),
            "season_id": season_id,
            "liga": nombre_liga_por_season.get(season_id, f"Liga {season_id}"),
            "local": home_team.get("name", home_id),
            "visitante": away_team.get("name", away_id),
            "xg_local": m_input.xg_local, "xga_local": m_input.xga_local,
            "xg_visitante": m_input.xg_visitante, "xga_visitante": m_input.xga_visitante,
            "gf_local": m_input.gf_local, "gc_local": m_input.gc_local,
            "gf_visitante": m_input.gf_visitante, "gc_visitante": m_input.gc_visitante,
            "sot_local": m_input.sot_local, "sot_visitante": m_input.sot_visitante,
            "failed_local_pct": m_input.failed_local_pct, "cs_local_pct": m_input.cs_local_pct,
            "failed_visit_pct": m_input.failed_visit_pct, "cs_visit_pct": m_input.cs_visit_pct,
            "btts_pct_local": m_input.btts_pct_local, "btts_pct_visitante": m_input.btts_pct_visitante,
            # Candidatas nuevas SIN VALIDAR (13 ago 2026) - se guardan para
            # medir su ROI en unas semanas, no se usan en ninguna formula
            # todavia.
            "over25_and_btts_pct_local": m_input.over25_and_btts_pct_local,
            "over25_and_btts_pct_visitante": m_input.over25_and_btts_pct_visitante,
            "ppg_local": m_input.ppg_local, "ppg_visitante": m_input.ppg_visitante,
            "shots_local": m_input.shots_local, "shots_visitante": m_input.shots_visitante,
            "avght_local": m_input.avght_local, "avght_visitante": m_input.avght_visitante,
            "avg2h_local": m_input.avg2h_local, "avg2h_visitante": m_input.avg2h_visitante,
            "cuota_btts": m_input.cuota_btts, "cuota_btts_no": match.get("odds_btts_no"),
            "cuota_o15": m_input.cuota_o15, "cuota_u15": match.get("odds_ft_under15"),
            "cuota_o25": m_input.cuota_o25, "cuota_u25": match.get("odds_ft_under25"),
            "cuota_local": m_input.cuota_local, "cuota_empate": match.get("odds_ft_x"),
            "cuota_visitante": m_input.cuota_visitante,
            "posicion_local": home_team.get("table_position"),
            "posicion_visitante": away_team.get("table_position"),
            "prob_btts": sig.prob_btts,
            "prob_btts_calibrada": sig.prob_btts_calibrada,
            "prob_o15": sig.p_over15,
            "prob_o25": sig.p_over25,
            "senal_btts_optimizada": sig.senal_btts_optimizada,
            "nivel_final": sig.nivel_final,
            "nivel_final_motivo": sig.nivel_final_motivo,
            "motivo_o25": sig.motivo_o25,
            "edge_btts": sig.edge_btts,
            "senal_o15": sig.senal_o15, "sug_o15": sig.sug_o15, "edge_o15": sig.edge_o15,
            "senal_o25": sig.senal_o25, "sug_o25": sig.sug_o25, "edge_o25": sig.edge_o25,
            "vip_elite": sig.vip_elite,
            "senal_btts_o25": sig.senal_btts_o25,
            "motivo_combo": sig.motivo_combo,
            "btts_min_alto": sig.btts_min_alto,
            "oro": sig.oro,
            "nivel_confianza": sig.nivel_confianza,
            "marcador_exacto_sugerido": marcador["marcador_sugerido"] if marcador else None,
            "usa_temporada_anterior": (home_stats.get("fuente_datos") == "anterior"
                                        or away_stats.get("fuente_datos") == "anterior"),
            "usa_promedio_general": (home_stats.get("fuente_datos") == "actual" and home_stats.get("fuente_lado") == "overall"
                                      or away_stats.get("fuente_datos") == "actual" and away_stats.get("fuente_lado") == "overall"),
            "es_equipo_filial": es_filial,
            "scored_ht_local": m_input.scored_ht_local, "conceded_ht_local": m_input.conceded_ht_local,
            "scored_ht_visitante": m_input.scored_ht_visitante, "conceded_ht_visitante": m_input.conceded_ht_visitante,
            "over05_ht_pct_local": m_input.over05_ht_pct_local, "over05_ht_pct_visitante": m_input.over05_ht_pct_visitante,
            "cuota_ht_over05": m_input.cuota_ht_over05,
            "prob_over05_ht": sig.prob_over05_ht,
            "prob_over15_ht": sig.prob_over15_ht,
            "senal_ht05": sig.senal_ht05, "sug_ht05": sig.sug_ht05,
            "motivo_ht05": sig.motivo_ht05, "edge_ht05": sig.edge_ht05,
            "formula_version": sig.formula_version,
        }
        filas.append(fila)

    df = pd.DataFrame(filas)

    os.makedirs(config.OUTPUT_DIR, exist_ok=True)
    stamp = datetime.now(ZoneInfo(config.TIMEZONE)).strftime("%Y-%m-%d") if ventana_24h else (date or datetime.now().strftime("%Y-%m-%d"))
    csv_path = os.path.join(config.OUTPUT_DIR, f"senales_{stamp}.csv")
    xlsx_path = os.path.join(config.OUTPUT_DIR, f"senales_{stamp}.xlsx")
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    df.to_excel(xlsx_path, index=False)
    print(f"Guardado: {csv_path}")
    print(f"Guardado: {xlsx_path}")

    nuevos_ids = backtest.record_predictions(filas, fecha=stamp)
    n_actualizados = backtest.update_results(client)
    if n_actualizados:
        print(f"[backtest] {n_actualizados} partidos anteriores completados con resultado real")

    if notify:
        n_alertas = 0
        for fila in filas:
            si_es_nuevo = fila["match_id"] in nuevos_ids
            if si_es_nuevo and any(fila.get(col) == "JUGAR" for col in config.NOTIFY_ON_COLUMNS):
                label = f"{fila['local']} vs {fila['visitante']}"
                precision = backtest.get_category_accuracy("senal_btts_optimizada", fila["senal_btts_optimizada"])
                texto = format_match_alert(label, str(fila["liga"]), fila, historial_accuracy=precision)
                send_telegram_message(config.TELEGRAM_BOT_TOKEN, config.TELEGRAM_CHAT_ID, texto)
                n_alertas += 1
        print(f"Alertas nuevas enviadas: {n_alertas}")

        if enviar_resumen:
            comparativo = backtest.build_summary_comparativo()
            if comparativo["historico"]["total_partidos_evaluados"] > 0:
                texto_resumen = backtest.format_summary_comparativo_message(comparativo)
                send_telegram_message(config.TELEGRAM_BOT_TOKEN, config.TELEGRAM_CHAT_ID, texto_resumen)

    backtest.export_dashboard_json(os.path.join(config.OUTPUT_DIR, "dashboard_data.json"))

    return df


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Calcula señales BTTS y notifica por Telegram")
    parser.add_argument("--date", default=None, help="Fecha YYYY-MM-DD exacta (ignora --dias-adelante si se usa)")
    parser.add_argument("--dias-adelante", type=int, default=0,
                         help="0 = partidos de hoy, 1 = partidos de mañana (calculado segun TIMEZONE), etc.")
    parser.add_argument("--ventana-24h", action="store_true",
                         help="Ignora --date/--dias-adelante: busca partidos entre AHORA y +24h. Pensado para correr cada 1-2 horas.")
    parser.add_argument("--sin-resumen", action="store_true",
                         help="No enviar el resumen diario de backtest (util si corres --ventana-24h varias veces al dia)")
    parser.add_argument("--no-notify", action="store_true", help="No enviar alertas de Telegram")
    args = parser.parse_args()

    if args.ventana_24h:
        run(notify=not args.no_notify, ventana_24h=True, enviar_resumen=not args.sin_resumen)
    else:
        if args.date:
            fecha_objetivo = args.date
        else:
            hoy_local = datetime.now(ZoneInfo(config.TIMEZONE)).date()
            fecha_objetivo = (hoy_local + timedelta(days=args.dias_adelante)).strftime("%Y-%m-%d")
        run(date=fecha_objetivo, notify=not args.no_notify, enviar_resumen=not args.sin_resumen)
