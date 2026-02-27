import random
import re
import sys
import json
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup


ROOT_DIR = Path(__file__).resolve().parent
UPDATE_DIR = ROOT_DIR / "RPL_V2" / "Update_Jul_2025"
TARGET_EXCEL = ROOT_DIR / "RPL_V2" / "xlsx" / "Матч Тура 25-26.xlsx"
RESULTS_URL = "https://soccer365.ru/competitions/13/results/"
SCHEDULE_URL = "https://soccer365.ru/competitions/13/shedule/"
COMPETITION_URL = "https://soccer365.ru/competitions/13/"
PLAYERS_URL = "https://soccer365.ru/competitions/13/players/"
BASE_URL = "https://soccer365.ru"
QUEST_IMAGE_BASE_URL = "https://amaranth64.github.io/RPL_V2/matchday/images"
REQUEST_HEADERS = {"User-Agent": "Mozilla/5.0"}
H2H_CACHE_FILE = ROOT_DIR / "RPL_V2" / "matchday" / "h2h_cache.json"


@dataclass(frozen=True)
class ClubInfo:
    slug: str
    display_name: str
    logo_slug: str
    aliases: tuple[str, ...]


@dataclass
class StageGame:
    home: str
    away: str
    status: str
    score_home: str
    score_away: str
    url: str
    round_num: int
    round_title: str
    dt_id: str


LOGO_BY_SLUG = {
    "dinmos": "logo_dinamo",
    "dinmah": "logo_dinamo_mah",
    "loko": "logo_lokomotiv",
    "pari": "logo_pari",
    "ks": "logo_ks",
}


EXTRA_ALIASES_BY_SLUG = {
    "cska": ["цска", "цска москва"],
    "dinmos": ["динамо", "динамо москва"],
    "dinmah": ["динамо махачкала", "махачкалинское динамо"],
    "baltika": ["балтика", "балтика калининград"],
    "krasnodar": ["краснодар"],
    "ks": ["крылья советов"],
    "loko": ["локомотив", "локомотив москва"],
    "pari": ["пари", "пари нн", "пари нижний новгород"],
    "rostov": ["ростов"],
    "rubin": ["рубин"],
    "sochi": ["сочи", "пфк сочи"],
    "spartak": ["спартак", "спартак москва"],
    "zenit": ["зенит"],
    "akron": ["акрон", "акрон тольятти"],
    "akhmat": ["ахмат"],
    "orenburg": ["оренбург"],
}


def log(message: str):
    print(f"[create_quiz] {message}")


def normalize_name(name: str) -> str:
    value = (name or "").lower().replace("ё", "е")
    value = re.sub(r"[^a-zа-я0-9]+", " ", value)
    return " ".join(value.split())


def build_club_catalog() -> dict[str, ClubInfo]:
    clubs: dict[str, ClubInfo] = {}
    for club_dir in sorted(path for path in UPDATE_DIR.iterdir() if path.is_dir()):
        club_files = sorted(club_dir.glob("club_*.xlsx"))
        if not club_files:
            continue

        slug = club_files[0].stem.replace("club_", "", 1)
        display_name = club_dir.name
        logo_slug = LOGO_BY_SLUG.get(slug, f"logo_{slug}")

        aliases = {display_name, slug, slug.replace("_", " ")}
        aliases.update(EXTRA_ALIASES_BY_SLUG.get(slug, []))

        clubs[slug] = ClubInfo(
            slug=slug,
            display_name=display_name,
            logo_slug=logo_slug,
            aliases=tuple(aliases),
        )

    return clubs


def build_alias_map(clubs: dict[str, ClubInfo]) -> dict[str, ClubInfo]:
    alias_map: dict[str, ClubInfo] = {}
    for club in clubs.values():
        for alias in club.aliases:
            alias_map[normalize_name(alias)] = club
    return alias_map


def resolve_club(name: str, alias_map: dict[str, ClubInfo]) -> ClubInfo:
    normalized = normalize_name(name)
    if normalized in alias_map:
        return alias_map[normalized]
    raise ValueError(f"Не удалось сопоставить клуб: '{name}'")


def fetch_soup(url: str) -> BeautifulSoup:
    response = requests.get(url, headers=REQUEST_HEADERS, timeout=25)
    response.raise_for_status()
    encoding = response.apparent_encoding or response.encoding or "utf-8"
    html = response.content.decode(encoding, errors="ignore")
    return BeautifulSoup(html, "html.parser")


def parse_round_number(round_title: str) -> int:
    if not round_title:
        return -1
    match = re.search(r"(\d+)", round_title)
    return int(match.group(1)) if match else -1


def add_one_hour(time_text: str) -> str:
    if not time_text:
        return time_text

    patterns = [
        (r"(\d{2}\.\d{2}\.\d{4})\.\s*(\d{1,2}:\d{2})", "%d.%m.%Y"),
        (r"(\d{2}\.\d{2}\.\d{2}),\s*(\d{1,2}:\d{2})", "%d.%m.%y"),
        (r"(\d{2}\.\d{2}),\s*(\d{1,2}:\d{2})", "%d.%m"),
    ]

    for pattern, date_fmt in patterns:
        found = re.search(pattern, time_text)
        if not found:
            continue

        date_part = found.group(1)
        time_part = found.group(2)

        if date_fmt == "%d.%m":
            dt = datetime.strptime(f"{date_part}.{datetime.now().year} {time_part}", "%d.%m.%Y %H:%M")
            dt += timedelta(hours=1)
            replacement = f"{dt.strftime('%d.%m')}, {dt.strftime('%H:%M')}"
            return re.sub(pattern, replacement, time_text, count=1)

        dt = datetime.strptime(f"{date_part} {time_part}", f"{date_fmt} %H:%M") + timedelta(hours=1)
        if date_fmt == "%d.%m.%Y":
            replacement = f"{dt.strftime('%d.%m.%Y')}. {dt.strftime('%H:%M')}"
        else:
            replacement = f"{dt.strftime('%d.%m.%y')}, {dt.strftime('%H:%M')}"
        return re.sub(pattern, replacement, time_text, count=1)

    only_time = re.search(r"\b(\d{1,2}):(\d{2})\b", time_text)
    if only_time:
        hour = int(only_time.group(1))
        minute = int(only_time.group(2))
        new_hour = (hour + 1) % 24
        replacement = f"{new_hour:02d}:{minute:02d}"
        return re.sub(r"\b\d{1,2}:\d{2}\b", replacement, time_text, count=1)

    return time_text


def get_club_excel_path(russian_club_name: str) -> Path:
    club_dir = UPDATE_DIR / russian_club_name
    if not club_dir.exists() or not club_dir.is_dir():
        raise FileNotFoundError(f"Не найдена папка клуба: {club_dir}")

    club_files = sorted(club_dir.glob("club_*.xlsx"))
    if not club_files:
        raise FileNotFoundError(f"В папке {club_dir} не найден club_*.xlsx")

    return club_files[0]


def get_club_slug(club_excel_path: Path) -> str:
    stem = club_excel_path.stem
    if not stem.startswith("club_"):
        raise ValueError(f"Некорректное имя файла клуба: {club_excel_path.name}")
    return stem.replace("club_", "", 1)


def split_question_pools(df: pd.DataFrame):
    text_series = (
        df["textQue"]
        .fillna("")
        .astype(str)
        .str.lower()
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )

    who_mask = text_series.str.match(r"^кто\s+эт[ао]т\s+футболист\??$", na=False)
    if int(who_mask.sum()) < 4:
        who_mask = text_series.str.contains(r"кто.*футболист", regex=True, na=False)

    coach_mask = text_series.str.contains(r"как\s+зовут.*главн(?:ого|ый)\s+тренер", regex=True, na=False)

    who_df = df[who_mask]
    other_df = df[~who_mask]
    coach_df = df[coach_mask]
    return who_df, other_df, coach_df


def sample_questions(df: pd.DataFrame, amount: int, label: str) -> pd.DataFrame:
    if len(df) < amount:
        raise ValueError(f"Недостаточно вопросов в категории '{label}': нужно {amount}, найдено {len(df)}")
    return df.sample(n=amount, replace=False)


def build_quiz_rows(club_excel_path: Path, columns: list[str]) -> pd.DataFrame:
    club_df = pd.read_excel(club_excel_path)
    missing = [col for col in columns if col not in club_df.columns]
    if missing:
        raise ValueError(f"В {club_excel_path.name} нет колонок: {missing}")

    who_df, other_df, coach_df = split_question_pools(club_df)
    log(f"{club_excel_path.name}: кто это футболист = {len(who_df)}, прочие = {len(other_df)}, тренер = {len(coach_df)}")

    who_part = sample_questions(who_df, 4, "Кто это футболист?")

    if len(coach_df) > 0:
        coach_row = coach_df.head(1)
        other_without_coach = other_df[~other_df.index.isin(coach_row.index)]
        other_part = sample_questions(other_without_coach, 3, "не 'Кто это футболист?' (без тренера)")
        picked = pd.concat([who_part, other_part, coach_row], ignore_index=True)
    else:
        other_part = sample_questions(other_df, 4, "не 'Кто это футболист?'")
        picked = pd.concat([who_part, other_part], ignore_index=True)

    return picked[columns]


def parse_competition_stages(url: str) -> list[dict]:
    soup = fetch_soup(url)
    stages: list[dict] = []

    for stage in soup.select("#result_data .live_comptt_bd"):
        round_title_node = stage.select_one(".cmp_stg_ttl")
        round_title = round_title_node.get_text(" ", strip=True) if round_title_node else ""
        round_num = parse_round_number(round_title)
        games: list[StageGame] = []

        for block in stage.select(".game_block"):
            link_node = block.select_one("a.game_link")
            if not link_node:
                continue

            home_node = block.select_one(".result .ht .name span")
            away_node = block.select_one(".result .at .name span")
            if not home_node or not away_node:
                continue

            status_node = block.select_one(".status span")
            score_home_node = block.select_one(".result .ht .gls")
            score_away_node = block.select_one(".result .at .gls")

            games.append(
                StageGame(
                    home=home_node.get_text(" ", strip=True),
                    away=away_node.get_text(" ", strip=True),
                    status=status_node.get_text(" ", strip=True) if status_node else "",
                    score_home=score_home_node.get_text(" ", strip=True) if score_home_node else "-",
                    score_away=score_away_node.get_text(" ", strip=True) if score_away_node else "-",
                    url=urljoin(BASE_URL, str(link_node.get("href", "") or "")),
                    round_num=round_num,
                    round_title=round_title,
                    dt_id=str(link_node.get("dt-id", "") or ""),
                )
            )

        if games:
            stages.append({"round": round_title, "round_num": round_num, "games": games})

    return stages


def find_pair_round(
    club_a: ClubInfo,
    club_b: ClubInfo,
    results_stages: list[dict],
    schedule_stages: list[dict],
    alias_map: dict[str, ClubInfo],
) -> tuple[int, StageGame, str]:
    def is_pair(game: StageGame) -> bool:
        try:
            home_club = resolve_club(game.home, alias_map)
            away_club = resolve_club(game.away, alias_map)
        except ValueError:
            return False

        slugs = {home_club.slug, away_club.slug}
        return slugs == {club_a.slug, club_b.slug}

    for stage in schedule_stages:
        for game in stage["games"]:
            if is_pair(game):
                return stage["round_num"], game, "schedule"

    for stage in results_stages:
        for game in stage["games"]:
            if is_pair(game):
                return stage["round_num"], game, "results"

    raise ValueError("Не найден очный матч введённых клубов в results/shedule")


def find_game_for_club_round(
    stages: list[dict],
    round_num: int,
    club: ClubInfo,
    alias_map: dict[str, ClubInfo],
) -> StageGame | None:
    for stage in stages:
        if stage.get("round_num") != round_num:
            continue
        for game in stage["games"]:
            try:
                home_club = resolve_club(game.home, alias_map)
                away_club = resolve_club(game.away, alias_map)
            except ValueError:
                continue

            if home_club.slug == club.slug or away_club.slug == club.slug:
                return game

    return None


def goal_separator(parts_by_scorer: list[tuple[str, list[str]]]) -> str:
    has_multi = any(len(minutes) > 1 for _, minutes in parts_by_scorer)
    joiner = "; " if has_multi else ", "
    chunks = []
    for scorer, minutes in parts_by_scorer:
        if len(minutes) > 1:
            chunks.append(f"{scorer} {', '.join(minutes)}")
        else:
            chunks.append(f"{scorer} {minutes[0]}")
    return joiner.join(chunks)


def build_match_comment(match_url: str) -> str:
    soup = fetch_soup(match_url)
    events_root = soup.select_one("#game_events")
    if not events_root:
        return ""

    status_node = events_root.select_one(".live_game_status b")
    status_text = status_node.get_text(" ", strip=True).lower() if status_node else ""

    home_name_node = events_root.select_one(".live_game.left .live_game_ht a")
    away_name_node = events_root.select_one(".live_game.right .live_game_at a")
    home_score_node = events_root.select_one(".live_game.left .live_game_goal span")
    away_score_node = events_root.select_one(".live_game.right .live_game_goal span")

    home_name = home_name_node.get_text(" ", strip=True) if home_name_node else ""
    away_name = away_name_node.get_text(" ", strip=True) if away_name_node else ""
    home_score = home_score_node.get_text(" ", strip=True) if home_score_node else "-"
    away_score = away_score_node.get_text(" ", strip=True) if away_score_node else "-"

    if home_score == "-" or away_score == "-":
        return ""

    if "заверш" not in status_text and "окончен" not in status_text:
        return ""

    home_events: list[tuple[str, str]] = []
    away_events: list[tuple[str, str]] = []

    for minute_node in events_root.select(".event_min"):
        row = minute_node.parent
        if row is None:
            continue
        minute = minute_node.get_text(" ", strip=True).replace(" +", "+").replace("+ ", "+")
        row_text = row.get_text(" ", strip=True).lower()
        is_own_goal = "(аг" in row_text or "автог" in row_text

        if row.select_one(".event_ht .live_goal"):
            scorer_node = row.select_one(".event_ht .img16 span a") or row.select_one(".event_ht .img16 span")
            scorer = scorer_node.get_text(" ", strip=True) if scorer_node else "Гол"
            if is_own_goal:
                scorer = f"{scorer}"
                minute = f"{minute}(аг)"
            home_events.append((scorer, minute))

        if row.select_one(".event_at .live_goal"):
            scorer_node = row.select_one(".event_at .img16 span a") or row.select_one(".event_at .img16 span")
            scorer = scorer_node.get_text(" ", strip=True) if scorer_node else "Гол"
            if is_own_goal:
                scorer = f"{scorer}"
                minute = f"{minute}(аг)"
            away_events.append((scorer, minute))

    def group_events(events: list[tuple[str, str]]) -> str:
        if not events:
            return ""
        grouped: dict[str, list[str]] = {}
        order: list[str] = []
        for scorer, minute in events:
            if scorer not in grouped:
                grouped[scorer] = []
                order.append(scorer)
            grouped[scorer].append(minute)
        pairs = [(scorer, grouped[scorer]) for scorer in order]
        return goal_separator(pairs)

    home_part = group_events(home_events)
    away_part = group_events(away_events)

    base = f"{home_name} {home_score}:{away_score} {away_name}".strip()
    if home_part and away_part:
        return f"{base} ({home_part} - {away_part})"
    if home_part:
        return f"{base} ({home_part})"
    if away_part:
        return f"{base} ({away_part})"
    return base


def random_wrong_clubs(clubs: dict[str, ClubInfo], exclude_slugs: set[str], amount: int) -> list[ClubInfo]:
    pool = [club for club in clubs.values() if club.slug not in exclude_slugs]
    if len(pool) < amount:
        raise ValueError("Недостаточно клубов для генерации неправильных ответов")
    return random.sample(pool, amount)


def choose_disjoint_wrong_clubs(
    clubs: dict[str, ClubInfo],
    common_exclude: set[str],
    extra_exclude_first: set[str],
    extra_exclude_second: set[str],
    amount: int,
) -> tuple[list[ClubInfo], list[ClubInfo]]:
    first_exclude = set(common_exclude) | set(extra_exclude_first)
    first_pool = [club for club in clubs.values() if club.slug not in first_exclude]
    if len(first_pool) < amount:
        raise ValueError("Недостаточно клубов для уникальных неправильных ответов (первый вопрос)")

    first_pick = random.sample(first_pool, amount)
    first_slugs = {club.slug for club in first_pick}

    second_exclude = set(common_exclude) | set(extra_exclude_second) | first_slugs
    second_pool = [club for club in clubs.values() if club.slug not in second_exclude]

    if len(second_pool) >= amount:
        second_pick = random.sample(second_pool, amount)
    else:
        # жёстко соблюдаем уникальность где возможно; если вдруг клубов мало — добираем без пересечений с правильными
        strict = second_pool
        relaxed_pool = [
            club
            for club in clubs.values()
            if club.slug not in (set(common_exclude) | set(extra_exclude_second)) and club.slug not in {c.slug for c in strict}
        ]
        need = amount - len(strict)
        if len(relaxed_pool) < need:
            raise ValueError("Недостаточно клубов для неправильных ответов второго вопроса")
        second_pick = strict + random.sample(relaxed_pool, need)

    return first_pick, second_pick


def make_team_question_row(
    columns: list[str],
    text_que: str,
    question_club: ClubInfo,
    correct_club: ClubInfo,
    comment: str,
    clubs: dict[str, ClubInfo],
    wrong_clubs: list[ClubInfo] | None = None,
) -> dict:
    row: dict[str, object] = {column: "" for column in columns}
    row["type"] = 2
    row["textQue"] = text_que
    row["slugQue"] = question_club.logo_slug
    row["comment"] = comment
    row["slugDlg"] = correct_club.logo_slug
    row["slugCorAns"] = correct_club.logo_slug
    row["txtCorAns"] = correct_club.display_name

    wrong_club_list = wrong_clubs or random_wrong_clubs(clubs, {question_club.slug, correct_club.slug}, 5)
    for index, wrong_club in enumerate(wrong_club_list, start=1):
        row[f"wrSlgAns{index}"] = wrong_club.logo_slug
        row[f"wrTxtAns{index}"] = wrong_club.display_name

    return row


def build_match_questions(
    clubs_pair: list[ClubInfo],
    columns: list[str],
    all_clubs: dict[str, ClubInfo],
    alias_map: dict[str, ClubInfo],
    sheet_name: str,
) -> tuple[pd.DataFrame, StageGame, int]:
    results_stages = parse_competition_stages(RESULTS_URL)
    schedule_stages = parse_competition_stages(SCHEDULE_URL)

    active_round, pair_game, source = find_pair_round(
        clubs_pair[0], clubs_pair[1], results_stages, schedule_stages, alias_map
    )
    log(f"Найден очный матч пары в {source}: раунд {active_round}, game_id={pair_game.dt_id}")

    previous_round = active_round - 1
    next_round = active_round + 1
    log(f"Для вопросов используем: прошлый тур={previous_round}, следующий тур={next_round}")

    rows: list[dict] = []
    for club in clubs_pair:
        prev_game = find_game_for_club_round(results_stages, previous_round, club, alias_map)
        prev_game_from_schedule = False
        if prev_game is None:
            prev_game = find_game_for_club_round(schedule_stages, previous_round, club, alias_map)
            prev_game_from_schedule = prev_game is not None
        next_game = find_game_for_club_round(schedule_stages, next_round, club, alias_map)

        prev_opponent = None
        next_opponent = None

        if prev_game:
            home_club = resolve_club(prev_game.home, alias_map)
            away_club = resolve_club(prev_game.away, alias_map)
            opponent = away_club if home_club.slug == club.slug else home_club
            prev_opponent = opponent
            if prev_game_from_schedule or prev_game.score_home == "-" or prev_game.score_away == "-":
                prev_comment = ""
            else:
                prev_comment = build_match_comment(prev_game.url)
                if not prev_comment:
                    prev_comment = f"{prev_game.home} {prev_game.score_home}:{prev_game.score_away} {prev_game.away}"
        else:
            log(f"Предупреждение: не найден матч прошлого тура (ни в results, ни в shedule) для {club.display_name}")

        if next_game:
            home_club = resolve_club(next_game.home, alias_map)
            away_club = resolve_club(next_game.away, alias_map)
            opponent = away_club if home_club.slug == club.slug else home_club
            next_opponent = opponent
            adjusted_status = add_one_hour(next_game.status)
            next_comment = f"{next_round} тур. {adjusted_status}".strip()
        else:
            log(f"Предупреждение: не найден матч следующего тура для {club.display_name}")

        if prev_opponent and next_opponent:
            wrong_prev, wrong_next = choose_disjoint_wrong_clubs(
                clubs=all_clubs,
                common_exclude={club.slug},
                extra_exclude_first={prev_opponent.slug},
                extra_exclude_second={next_opponent.slug},
                amount=5,
            )

            rows.append(
                make_team_question_row(
                    columns=columns,
                    text_que=f"С какой командой играл {club.display_name} в прошлом туре?",
                    question_club=club,
                    correct_club=prev_opponent,
                    comment=prev_comment,
                    clubs=all_clubs,
                    wrong_clubs=wrong_prev,
                )
            )
            rows.append(
                make_team_question_row(
                    columns=columns,
                    text_que=f"С кем сыграет {club.display_name} в следующем туре?",
                    question_club=club,
                    correct_club=next_opponent,
                    comment=next_comment,
                    clubs=all_clubs,
                    wrong_clubs=wrong_next,
                )
            )
        elif prev_opponent:
            rows.append(
                make_team_question_row(
                    columns=columns,
                    text_que=f"С какой командой играл {club.display_name} в прошлом туре?",
                    question_club=club,
                    correct_club=prev_opponent,
                    comment=prev_comment,
                    clubs=all_clubs,
                )
            )
        elif next_opponent:
            rows.append(
                make_team_question_row(
                    columns=columns,
                    text_que=f"С кем сыграет {club.display_name} в следующем туре?",
                    question_club=club,
                    correct_club=next_opponent,
                    comment=next_comment,
                    clubs=all_clubs,
                )
            )

    return pd.DataFrame(rows, columns=columns), pair_game, active_round


def parse_standings(
    alias_map: dict[str, ClubInfo],
) -> tuple[list[dict], int]:
    soup = fetch_soup(COMPETITION_URL)
    table = soup.select_one("#competition_table table") or soup.select_one("#competition_table")
    if not table:
        raise ValueError("Не удалось найти таблицу РПЛ на странице competitions/13")

    standings: list[dict] = []
    rounds = 0

    for tr in table.select("tr"):
        tds = tr.select("td")
        if len(tds) < 3:
            continue

        team_anchor = tr.select_one("a[href*='/clubs/']")
        if not team_anchor:
            continue

        team_name_raw = team_anchor.get_text(" ", strip=True)
        try:
            club = resolve_club(team_name_raw, alias_map)
        except ValueError:
            continue

        position_text = tds[0].get_text(" ", strip=True)
        games_text = tds[2].get_text(" ", strip=True)
        points_text = tds[-1].get_text(" ", strip=True)

        if not position_text.isdigit() or not games_text.isdigit() or not points_text.isdigit():
            continue

        position = int(position_text)
        games = int(games_text)
        points = int(points_text)
        rounds = max(rounds, games)

        standings.append(
            {
                "club": club,
                "position": position,
                "games": games,
                "points": points,
            }
        )

    standings.sort(key=lambda item: item["position"])
    if not standings:
        raise ValueError("Не удалось распарсить турнирную таблицу РПЛ")

    return standings, rounds


def build_player_base() -> dict[str, dict]:
    player_map: dict[str, dict] = {}

    def store(name: str, slug: str):
        clean_name = (name or "").strip()
        if not clean_name:
            return

        normalized = normalize_name(clean_name)
        if not normalized:
            return

        normalized_slug = (slug or "").strip()
        if normalized_slug.startswith("logo_") or normalized_slug.startswith("num_"):
            return
        if normalized_slug.startswith("http"):
            return

        if normalized not in player_map:
            player_map[normalized] = {"name": clean_name, "slug": normalized_slug or "no_pic"}
            return

        if player_map[normalized]["slug"] in {"", "no_pic"} and normalized_slug and normalized_slug != "no_pic":
            player_map[normalized]["slug"] = normalized_slug

    for club_dir in sorted(path for path in UPDATE_DIR.iterdir() if path.is_dir()):
        excel_files = sorted(club_dir.glob("club_*.xlsx"))
        if not excel_files:
            continue

        df = pd.read_excel(excel_files[0])
        for _, row in df.iterrows():
            store(str(row.get("txtCorAns", "")), str(row.get("slugCorAns", "")))
            for index in range(1, 6):
                store(str(row.get(f"wrTxtAns{index}", "")), str(row.get(f"wrSlgAns{index}", "")))

    return player_map


def parse_top_scorers(player_base: dict[str, dict]) -> list[dict]:
    soup = fetch_soup(PLAYERS_URL)
    table = soup.select_one("table")
    if not table:
        raise ValueError("Не удалось найти таблицу игроков на странице players")

    scorers: list[dict] = []
    for tr in table.select("tr"):
        tds = tr.select("td")
        if len(tds) < 4:
            continue

        player_anchor = tr.select_one("a[href*='/players/']")
        if not player_anchor:
            continue

        name = player_anchor.get_text(" ", strip=True)
        normalized = normalize_name(name)
        if normalized not in player_base:
            continue

        goals_text = tds[1].get_text(" ", strip=True)
        matches_text = tds[3].get_text(" ", strip=True)
        if not goals_text.isdigit() or not matches_text.isdigit():
            continue

        base_data = player_base[normalized]
        scorers.append(
            {
                "name": base_data["name"],
                "slug": base_data["slug"] if base_data["slug"] else "no_pic",
                "goals": int(goals_text),
                "matches": int(matches_text),
            }
        )

    if len(scorers) < 6:
        raise ValueError(
            "Не удалось собрать 6 бомбардиров с совпадением имён в базе клубов. "
            f"Найдено: {len(scorers)}"
        )

    return scorers


def make_place_question(columns: list[str], club: ClubInfo, place: int, points: int, rounds: int) -> dict:
    row: dict[str, object] = {column: "" for column in columns}
    row["type"] = 4
    row["textQue"] = f"На каком месте находится {club.display_name} после {rounds} туров?"
    row["slugQue"] = club.logo_slug
    row["comment"] = f"{points} очков"
    row["slugDlg"] = club.logo_slug
    row["slugCorAns"] = f"num_{place}"
    row["txtCorAns"] = f"{place} место"

    min_place = max(1, place - 4)
    max_place = min(16, place + 4)
    wrong_places = [value for value in range(min_place, max_place + 1) if value != place]
    if len(wrong_places) < 5:
        wrong_places = [value for value in range(1, 17) if value != place]
    wrong_pick = random.sample(wrong_places, 5)
    for index, wrong_place in enumerate(wrong_pick, start=1):
        row[f"wrSlgAns{index}"] = f"num_{wrong_place}"
        row[f"wrTxtAns{index}"] = f"{wrong_place} место"

    return row


def build_competition_questions(
    columns: list[str],
    clubs_pair: list[ClubInfo],
    clubs_catalog: dict[str, ClubInfo],
    alias_map: dict[str, ClubInfo],
) -> pd.DataFrame:
    standings, rounds_from_table = parse_standings(alias_map)
    results_stages = parse_competition_stages(RESULTS_URL)
    completed_rounds = [
        stage["round_num"]
        for stage in results_stages
        if stage.get("round_num", -1) > 0 and all(game.score_home.isdigit() and game.score_away.isdigit() for game in stage["games"])
    ]
    rounds = max(completed_rounds) if completed_rounds else rounds_from_table
    log(f"Таблица РПЛ: туров={rounds}, команд={len(standings)}")

    by_slug = {item["club"].slug: item for item in standings}
    leader = standings[0]

    rows: list[dict] = []

    # Кто лидирует
    leader_row: dict[str, object] = {column: "" for column in columns}
    leader_row["type"] = 2
    leader_row["textQue"] = f"Кто лидирует в РПЛ после {rounds} туров?"
    leader_row["slugQue"] = "logo_rpl"
    leader_row["comment"] = f"{leader['points']} очков"
    leader_row["slugDlg"] = leader["club"].logo_slug
    leader_row["slugCorAns"] = leader["club"].logo_slug
    leader_row["txtCorAns"] = leader["club"].display_name

    wrong_leader_clubs = random_wrong_clubs(clubs_catalog, {leader["club"].slug}, 5)
    for index, wrong_club in enumerate(wrong_leader_clubs, start=1):
        leader_row[f"wrSlgAns{index}"] = wrong_club.logo_slug
        leader_row[f"wrTxtAns{index}"] = wrong_club.display_name

    rows.append(leader_row)

    # На каком месте 2 выбранных клуба
    for club in clubs_pair:
        if club.slug not in by_slug:
            raise ValueError(f"Клуб '{club.display_name}' не найден в турнирной таблице")
        info = by_slug[club.slug]
        rows.append(make_place_question(columns, club, info["position"], info["points"], rounds))

    # Лучший бомбардир
    player_base = build_player_base()
    scorers = parse_top_scorers(player_base)
    log(f"Топ-бомбардиры с совпадением в базе: {len(scorers)}")

    if len(scorers) > 1 and scorers[0]["goals"] == scorers[1]["goals"]:
        log("Вопрос о лучшем бомбардире пропущен: у лидеров одинаковое число голов")
    else:
        scorer_row: dict[str, object] = {column: "" for column in columns}
        scorer_row["type"] = 2
        scorer_row["textQue"] = f"Кто является лучшим бомбардиром РПЛ после {rounds} туров?"
        scorer_row["slugQue"] = "logo_rpl"
        scorer_row["comment"] = f"Забил {scorers[0]['goals']} голов в {scorers[0]['matches']} матчах"
        scorer_row["slugDlg"] = scorers[0]["slug"]
        scorer_row["slugCorAns"] = scorers[0]["slug"]
        scorer_row["txtCorAns"] = scorers[0]["name"]

        wrong_scorers = scorers[1:6]
        for index, player in enumerate(wrong_scorers, start=1):
            scorer_row[f"wrSlgAns{index}"] = player["slug"]
            scorer_row[f"wrTxtAns{index}"] = player["name"]

        rows.append(scorer_row)

    # Последние 5 матчей по очному матчу текущей пары
    return pd.DataFrame(rows, columns=columns)


def parse_last5_points_for_pair_match(pair_game: StageGame) -> tuple[int, int] | None:
    soup = fetch_soup(pair_game.url)
    form_blocks = soup.select(".game_form_team")
    if len(form_blocks) < 2:
        return None

    def calc_points(block) -> int:
        result = 0
        for span in block.select("span")[:5]:
            cls = " ".join(span.get("class", [])).lower()
            if "wins" in cls:
                result += 3
            elif "draw" in cls:
                result += 1
        return result

    home_points = calc_points(form_blocks[0])
    away_points = calc_points(form_blocks[1])
    return home_points, away_points


def make_form_question(
    columns: list[str],
    pair_game: StageGame,
    alias_map: dict[str, ClubInfo],
) -> pd.DataFrame:
    points = parse_last5_points_for_pair_match(pair_game)
    if not points:
        log("Вопрос по форме за 5 матчей пропущен: не найден блок формы")
        return pd.DataFrame([], columns=columns)

    home_points, away_points = points
    if home_points == away_points:
        log("Вопрос по форме за 5 матчей пропущен: очки одинаковые")
        return pd.DataFrame([], columns=columns)

    home_club = resolve_club(pair_game.home, alias_map)
    away_club = resolve_club(pair_game.away, alias_map)
    winner = home_club if home_points > away_points else away_club

    row: dict[str, object] = {column: "" for column in columns}
    row["type"] = 2
    row["textQue"] = "Какая из команд набрала больше очков в последних 5 матчах?"
    row["slugQue"] = "logo_rpl"
    row["comment"] = (
        f"{home_club.display_name}: {home_points} очков, "
        f"{away_club.display_name}: {away_points} очков"
    )
    row["slugDlg"] = winner.logo_slug
    row["slugCorAns"] = winner.logo_slug
    row["txtCorAns"] = winner.display_name

    wrong_pool = [home_club, away_club]
    wrong_pool = [club for club in wrong_pool if club.slug != winner.slug]
    # добираем ещё клубами РПЛ, чтобы не было очевидно
    filler = [club for club in build_club_catalog().values() if club.slug not in {home_club.slug, away_club.slug, winner.slug}]
    random.shuffle(filler)
    wrong_clubs = wrong_pool + filler[:4]

    for index, club in enumerate(wrong_clubs[:5], start=1):
        row[f"wrSlgAns{index}"] = club.logo_slug
        row[f"wrTxtAns{index}"] = club.display_name

    return pd.DataFrame([row], columns=columns)

    return pd.DataFrame(rows, columns=columns)


def parse_h2h_games(stats_url: str) -> list[dict]:
    soup = fetch_soup(stats_url)
    games: list[dict] = []

    for link in soup.select("a.game_link"):
        home_node = link.select_one(".result .ht .name span")
        away_node = link.select_one(".result .at .name span")
        hg_node = link.select_one(".result .ht .gls")
        ag_node = link.select_one(".result .at .gls")
        status_node = link.select_one(".status span")
        comp_node = link.select_one(".cmp span")

        if not home_node or not away_node or not hg_node or not ag_node:
            continue

        games.append(
            {
                "home": home_node.get_text(" ", strip=True),
                "away": away_node.get_text(" ", strip=True),
                "score_home": hg_node.get_text(" ", strip=True),
                "score_away": ag_node.get_text(" ", strip=True),
                "status": status_node.get_text(" ", strip=True) if status_node else "",
                "competition": comp_node.get_text(" ", strip=True) if comp_node else "",
                "url": urljoin(BASE_URL, str(link.get("href", "") or "")),
                "dt_id": str(link.get("dt-id", "") or ""),
            }
        )

    return games


def load_h2h_cache() -> dict:
    if not H2H_CACHE_FILE.exists():
        return {}
    try:
        with open(H2H_CACHE_FILE, "r", encoding="utf-8") as file:
            data = json.load(file)
            return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_h2h_cache(cache: dict):
    with open(H2H_CACHE_FILE, "w", encoding="utf-8") as file:
        json.dump(cache, file, ensure_ascii=False, indent=2)


def get_match_goal_analytics(match_url: str) -> dict:
    soup = fetch_soup(match_url)
    events_root = soup.select_one("#game_events")
    if not events_root:
        return {"comment": "", "home_scorers": [], "away_scorers": []}

    status_node = events_root.select_one(".live_game_status b")
    status_text = status_node.get_text(" ", strip=True).lower() if status_node else ""

    home_name_node = events_root.select_one(".live_game.left .live_game_ht a")
    away_name_node = events_root.select_one(".live_game.right .live_game_at a")
    home_score_node = events_root.select_one(".live_game.left .live_game_goal span")
    away_score_node = events_root.select_one(".live_game.right .live_game_goal span")

    home_name = home_name_node.get_text(" ", strip=True) if home_name_node else ""
    away_name = away_name_node.get_text(" ", strip=True) if away_name_node else ""
    home_score = home_score_node.get_text(" ", strip=True) if home_score_node else "-"
    away_score = away_score_node.get_text(" ", strip=True) if away_score_node else "-"

    if home_score == "-" or away_score == "-":
        return {"comment": "", "home_scorers": [], "away_scorers": []}

    if "заверш" not in status_text and "окончен" not in status_text:
        return {"comment": "", "home_scorers": [], "away_scorers": []}

    home_events: list[tuple[str, str, bool]] = []
    away_events: list[tuple[str, str, bool]] = []

    for minute_node in events_root.select(".event_min"):
        row = minute_node.parent
        if row is None:
            continue
        minute = minute_node.get_text(" ", strip=True).replace(" +", "+").replace("+ ", "+")
        row_text = row.get_text(" ", strip=True).lower()
        is_own_goal = "(аг" in row_text or "автог" in row_text

        if row.select_one(".event_ht .live_goal"):
            scorer_node = row.select_one(".event_ht .img16 span a") or row.select_one(".event_ht .img16 span")
            scorer = scorer_node.get_text(" ", strip=True) if scorer_node else "Гол"
            minute_formatted = f"{minute}(аг)" if is_own_goal else minute
            home_events.append((scorer, minute_formatted, is_own_goal))

        if row.select_one(".event_at .live_goal"):
            scorer_node = row.select_one(".event_at .img16 span a") or row.select_one(".event_at .img16 span")
            scorer = scorer_node.get_text(" ", strip=True) if scorer_node else "Гол"
            minute_formatted = f"{minute}(аг)" if is_own_goal else minute
            away_events.append((scorer, minute_formatted, is_own_goal))

    def group_events(events: list[tuple[str, str, bool]]) -> str:
        if not events:
            return ""
        grouped: dict[str, list[str]] = {}
        order: list[str] = []
        for scorer, minute, _ in events:
            if scorer not in grouped:
                grouped[scorer] = []
                order.append(scorer)
            grouped[scorer].append(minute)
        pairs = [(scorer, grouped[scorer]) for scorer in order]
        return goal_separator(pairs)

    home_part = group_events(home_events)
    away_part = group_events(away_events)

    base = f"{home_name} {home_score}:{away_score} {away_name}".strip()
    if home_part and away_part:
        comment = f"{base} ({home_part} - {away_part})"
    elif home_part:
        comment = f"{base} ({home_part})"
    elif away_part:
        comment = f"{base} ({away_part})"
    else:
        comment = base

    # Для статистики игроков не учитываем автоголы
    home_scorers = [scorer for scorer, _, is_own in home_events if not is_own]
    away_scorers = [scorer for scorer, _, is_own in away_events if not is_own]

    return {
        "comment": comment,
        "home_scorers": home_scorers,
        "away_scorers": away_scorers,
    }


def build_h2h_questions(
    columns: list[str],
    club_a: ClubInfo,
    club_b: ClubInfo,
    pair_game: StageGame,
    sheet_name: str,
    alias_map: dict[str, ClubInfo],
) -> pd.DataFrame:
    if not pair_game.dt_id:
        log("Предупреждение: у очного матча пары нет dt-id, вопросы личных встреч пропущены")
        return pd.DataFrame([], columns=columns)

    stats_url = f"{BASE_URL}/games/{pair_game.dt_id}/&tab=stats_games"
    h2h_games = parse_h2h_games(stats_url)
    log(f"Личные встречи: получено матчей {len(h2h_games)}")

    official_games = [game for game in h2h_games if "товарищ" not in normalize_name(game["competition"])]
    if not official_games:
        log("Предупреждение: не найдено официальных матчей в личных встречах")
        return pd.DataFrame([], columns=columns)

    league_games = [
        game
        for game in official_games
        if "премьер" in normalize_name(game["competition"]) and "кубок" not in normalize_name(game["competition"])
    ]

    wins = {club_a.slug: 0, club_b.slug: 0}
    for game in official_games:
        if not game["score_home"].isdigit() or not game["score_away"].isdigit():
            continue

        home_score = int(game["score_home"])
        away_score = int(game["score_away"])

        try:
            home_club = resolve_club(game["home"], alias_map)
            away_club = resolve_club(game["away"], alias_map)
        except ValueError:
            continue

        if home_club.slug not in wins or away_club.slug not in wins:
            continue

        if home_score > away_score:
            wins[home_club.slug] += 1
        elif away_score > home_score:
            wins[away_club.slug] += 1

    if wins[club_a.slug] > wins[club_b.slug]:
        winner_text = club_a.display_name
    elif wins[club_b.slug] > wins[club_a.slug]:
        winner_text = club_b.display_name
    else:
        winner_text = "Поровну"

    quest_image = f"{QUEST_IMAGE_BASE_URL}/{sheet_name}_quest.jpg"

    q1: dict[str, object] = {column: "" for column in columns}
    q1["type"] = 1
    q1["textQue"] = f"Кто чаще побеждал в очных противостояниях, {club_a.display_name} или {club_b.display_name}?"
    q1["slugQue"] = quest_image
    q1["comment"] = f"Официальные: победы {club_a.display_name} — {wins[club_a.slug]}, {club_b.display_name} — {wins[club_b.slug]}"
    q1["slugDlg"] = quest_image
    q1["slugCorAns"] = "no_pic"
    q1["txtCorAns"] = winner_text

    wrong_pool = [club_a.display_name, club_b.display_name, "Поровну", "Ничья", "Никто", "Затрудняюсь ответить"]
    wrong_pool = [item for item in wrong_pool if item != winner_text]
    wrong_pick = wrong_pool[:5]
    while len(wrong_pick) < 5:
        wrong_pick.append(f"Вариант {len(wrong_pick) + 1}")

    for index, value in enumerate(wrong_pick, start=1):
        q1[f"wrSlgAns{index}"] = "no_pic"
        q1[f"wrTxtAns{index}"] = value

    # Последняя встреча (без товарищеских и без кубка)
    last_pool = league_games if league_games else official_games
    last_game = last_pool[0]
    score_text = f"{last_game['home']} {last_game['score_home']}:{last_game['score_away']} {last_game['away']}"
    last_comment = ""
    cache = load_h2h_cache()

    if last_game["score_home"] != "-" and last_game["score_away"] != "-":
        cache_key = last_game.get("dt_id", "")
        cached = cache.get(cache_key, {}) if cache_key else {}
        if (
            cached
            and cached.get("score_home") == last_game["score_home"]
            and cached.get("score_away") == last_game["score_away"]
            and cached.get("status") == last_game["status"]
        ):
            last_comment = cached.get("comment", "")
        else:
            analytics = get_match_goal_analytics(last_game["url"])
            last_comment = analytics.get("comment", "")
            if cache_key:
                cache[cache_key] = {
                    "url": last_game["url"],
                    "status": last_game["status"],
                    "score_home": last_game["score_home"],
                    "score_away": last_game["score_away"],
                    "home": last_game["home"],
                    "away": last_game["away"],
                    **analytics,
                }

    q2: dict[str, object] = {column: "" for column in columns}
    q2["type"] = 1
    q2["textQue"] = "С каким счетом завершилась их последняя встреча?"
    q2["slugQue"] = quest_image
    q2["comment"] = last_comment
    q2["slugDlg"] = quest_image
    q2["slugCorAns"] = "no_pic"
    q2["txtCorAns"] = score_text

    if last_game["score_home"].isdigit() and last_game["score_away"].isdigit():
        home_score = int(last_game["score_home"])
        away_score = int(last_game["score_away"])
        variants = [
            f"{last_game['home']} {home_score + 1}:{away_score} {last_game['away']}",
            f"{last_game['home']} {home_score}:{away_score + 1} {last_game['away']}",
            f"{last_game['home']} {max(home_score - 1, 0)}:{away_score} {last_game['away']}",
            f"{last_game['home']} {home_score}:{max(away_score - 1, 0)} {last_game['away']}",
            f"{last_game['home']} {away_score}:{home_score} {last_game['away']}",
        ]
    else:
        variants = [
            f"{last_game['home']} 1:0 {last_game['away']}",
            f"{last_game['home']} 0:1 {last_game['away']}",
            f"{last_game['home']} 2:1 {last_game['away']}",
            f"{last_game['home']} 1:2 {last_game['away']}",
            f"{last_game['home']} 0:0 {last_game['away']}",
        ]

    for index, value in enumerate(variants[:5], start=1):
        q2[f"wrSlgAns{index}"] = "no_pic"
        q2[f"wrTxtAns{index}"] = value

    rows = [q1, q2]

    # Новый вопрос по лучшему бомбардиру очных встреч для каждой стороны (без товарищеских)
    scorer_totals = {
        club_a.slug: {},
        club_b.slug: {},
    }

    for game in official_games:
        if not game.get("dt_id"):
            continue

        cache_key = game["dt_id"]
        cached = cache.get(cache_key, {})
        if (
            not cached
            or cached.get("score_home") != game["score_home"]
            or cached.get("score_away") != game["score_away"]
            or cached.get("status") != game["status"]
        ):
            analytics = get_match_goal_analytics(game["url"])
            cache[cache_key] = {
                "url": game["url"],
                "status": game["status"],
                "score_home": game["score_home"],
                "score_away": game["score_away"],
                "home": game["home"],
                "away": game["away"],
                **analytics,
            }

        details = cache.get(cache_key, {})
        try:
            home_club = resolve_club(game["home"], alias_map)
            away_club = resolve_club(game["away"], alias_map)
        except ValueError:
            continue

        for scorer in details.get("home_scorers", []):
            if home_club.slug in scorer_totals:
                scorer_totals[home_club.slug][scorer] = scorer_totals[home_club.slug].get(scorer, 0) + 1
        for scorer in details.get("away_scorers", []):
            if away_club.slug in scorer_totals:
                scorer_totals[away_club.slug][scorer] = scorer_totals[away_club.slug].get(scorer, 0) + 1

    for focus, opponent in [(club_a, club_b), (club_b, club_a)]:
        totals = scorer_totals.get(focus.slug, {})
        if not totals:
            continue
        sorted_players = sorted(totals.items(), key=lambda item: (-item[1], item[0]))
        best_name, best_goals = sorted_players[0]

        q: dict[str, object] = {column: "" for column in columns}
        q["type"] = 1
        q["textQue"] = (
            f"Какой игрок забил большее число голов в личных противостояниях с "
            f"{opponent.display_name} за {focus.display_name}?"
        )
        q["slugQue"] = quest_image
        q["comment"] = f"{best_name}: {best_goals} голов"
        q["slugDlg"] = quest_image
        q["slugCorAns"] = "no_pic"
        q["txtCorAns"] = best_name

        wrong_names = [name for name, _ in sorted_players[1:6]]
        while len(wrong_names) < 5:
            wrong_names.append(f"Вариант {len(wrong_names) + 1}")
        for index, name in enumerate(wrong_names[:5], start=1):
            q[f"wrSlgAns{index}"] = "no_pic"
            q[f"wrTxtAns{index}"] = name

        rows.append(q)

    save_h2h_cache(cache)

    return pd.DataFrame(rows, columns=columns)


def main():
    if len(sys.argv) != 3:
        print('Использование: py create_quiz.py "Зенит" "Балтика"')
        sys.exit(1)

    club_home_ru = sys.argv[1].strip()
    club_guest_ru = sys.argv[2].strip()

    log(f"Старт: клубы '{club_home_ru}' и '{club_guest_ru}'")

    if not club_home_ru or not club_guest_ru:
        print("Ошибка: названия клубов не должны быть пустыми.")
        sys.exit(1)

    if not TARGET_EXCEL.exists():
        print(f"Ошибка: не найден файл {TARGET_EXCEL}")
        sys.exit(1)

    clubs_catalog = build_club_catalog()
    alias_map = build_alias_map(clubs_catalog)
    log(f"Каталог клубов собран: {len(clubs_catalog)}")

    if len(clubs_catalog) < 6:
        print("Ошибка: недостаточно клубов в каталоге для генерации вариантов ответа.")
        sys.exit(1)

    home_excel = get_club_excel_path(club_home_ru)
    guest_excel = get_club_excel_path(club_guest_ru)
    home_slug = get_club_slug(home_excel)
    guest_slug = get_club_slug(guest_excel)

    if home_slug not in clubs_catalog or guest_slug not in clubs_catalog:
        print("Ошибка: один из клубов не найден в каталоге Update_Jul_2025.")
        sys.exit(1)

    home_club = clubs_catalog[home_slug]
    guest_club = clubs_catalog[guest_slug]
    new_sheet_name = f"mt_{home_slug}_{guest_slug}"
    log(f"Лист назначения: {new_sheet_name}")

    xl = pd.ExcelFile(TARGET_EXCEL)
    sheet_exists = new_sheet_name in xl.sheet_names

    template_df = xl.parse(xl.sheet_names[0])
    columns = list(template_df.columns)
    log(f"Колонки шаблона: {len(columns)}")

    home_rows = build_quiz_rows(home_excel, columns)
    guest_rows = build_quiz_rows(guest_excel, columns)
    log(f"Клубные вопросы: {len(home_rows)} + {len(guest_rows)}")

    match_rows, pair_game, active_round = build_match_questions(
        [home_club, guest_club], columns, clubs_catalog, alias_map, new_sheet_name
    )
    log(f"Матчевые вопросы (прошлый/следующий тур): {len(match_rows)}")

    competition_rows = build_competition_questions(columns, [home_club, guest_club], clubs_catalog, alias_map)
    log(f"Турнирные вопросы: {len(competition_rows)}")

    form_rows = make_form_question(columns, pair_game, alias_map)
    log(f"Вопросы по форме 5 матчей: {len(form_rows)}")

    h2h_rows = build_h2h_questions(columns, home_club, guest_club, pair_game, new_sheet_name, alias_map)
    log(f"Вопросы по личным встречам: {len(h2h_rows)}")

    result_df = pd.concat([home_rows, guest_rows, match_rows, competition_rows, form_rows, h2h_rows], ignore_index=True)
    log(f"Итого вопросов в листе: {len(result_df)}")

    if sheet_exists:
        with pd.ExcelWriter(TARGET_EXCEL, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            result_df.to_excel(writer, sheet_name=new_sheet_name, index=False)
    else:
        with pd.ExcelWriter(TARGET_EXCEL, engine="openpyxl", mode="a") as writer:
            result_df.to_excel(writer, sheet_name=new_sheet_name, index=False)

    if sheet_exists:
        print(f"Лист '{new_sheet_name}' существовал и был пересоздан")
    else:
        print(f"Создан новый лист: {new_sheet_name}")

    print(f"Активный очный тур пары: {active_round}")
    print(
        "Добавлены блоки: 8+8 клубных (с обязательным тренером), "
        "прошлый/следующий тур, таблица РПЛ, бомбардир, очные встречи"
    )
    print(f"Файл: {TARGET_EXCEL}")


if __name__ == "__main__":
    main()
