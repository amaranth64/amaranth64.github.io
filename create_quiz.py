import random
import re
import sys
from dataclasses import dataclass
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
BASE_URL = "https://soccer365.ru"
REQUEST_HEADERS = {"User-Agent": "Mozilla/5.0"}


@dataclass(frozen=True)
class ClubInfo:
    slug: str
    display_name: str
    logo_slug: str
    aliases: tuple[str, ...]


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
        clubs[slug] = ClubInfo(slug=slug, display_name=display_name, logo_slug=logo_slug, aliases=tuple(aliases))

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

    who_df = df[who_mask]
    other_df = df[~who_mask]
    return who_df, other_df


def sample_questions(df: pd.DataFrame, amount: int, label: str) -> pd.DataFrame:
    if len(df) < amount:
        raise ValueError(f"Недостаточно вопросов в категории '{label}': нужно {amount}, найдено {len(df)}")
    return df.sample(n=amount, replace=False)


def build_quiz_rows(club_excel_path: Path, columns: list[str]) -> pd.DataFrame:
    club_df = pd.read_excel(club_excel_path)
    missing = [col for col in columns if col not in club_df.columns]
    if missing:
        raise ValueError(f"В {club_excel_path.name} нет колонок: {missing}")

    who_df, other_df = split_question_pools(club_df)
    picked = pd.concat(
        [
            sample_questions(who_df, 4, "Кто это футболист?"),
            sample_questions(other_df, 4, "не 'Кто это футболист?'"),
        ],
        ignore_index=True,
    )
    return picked[columns]


def parse_competition_stages(url: str) -> list[dict]:
    soup = fetch_soup(url)
    stages: list[dict] = []
    for stage in soup.select("#result_data .live_comptt_bd"):
        round_title_node = stage.select_one(".cmp_stg_ttl")
        round_title = round_title_node.get_text(" ", strip=True) if round_title_node else ""
        games: list[dict] = []

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
                {
                    "home": home_node.get_text(" ", strip=True),
                    "away": away_node.get_text(" ", strip=True),
                    "status": status_node.get_text(" ", strip=True) if status_node else "",
                    "score_home": score_home_node.get_text(" ", strip=True) if score_home_node else "-",
                    "score_away": score_away_node.get_text(" ", strip=True) if score_away_node else "-",
                    "url": urljoin(BASE_URL, link_node.get("href", "")),
                }
            )

        if games:
            stages.append({"round": round_title, "games": games})

    return stages


def find_club_game(stages: list[dict], club: ClubInfo, alias_map: dict[str, ClubInfo]) -> tuple[dict, dict]:
    for stage in stages:
        for game in stage["games"]:
            home_club = resolve_club(game["home"], alias_map)
            away_club = resolve_club(game["away"], alias_map)

            if home_club.slug == club.slug:
                game_data = dict(game)
                game_data["opponent"] = away_club
                return stage, game_data

            if away_club.slug == club.slug:
                game_data = dict(game)
                game_data["opponent"] = home_club
                return stage, game_data

    raise ValueError(f"Не найден матч для клуба '{club.display_name}' в турнирной выборке")


def build_match_comment(match_url: str) -> str:
    soup = fetch_soup(match_url)
    events_root = soup.select_one("#game_events")
    if not events_root:
        return ""

    home_name_node = events_root.select_one(".live_game.left .live_game_ht a")
    away_name_node = events_root.select_one(".live_game.right .live_game_at a")
    home_score_node = events_root.select_one(".live_game.left .live_game_goal span")
    away_score_node = events_root.select_one(".live_game.right .live_game_goal span")

    home_name = home_name_node.get_text(" ", strip=True) if home_name_node else ""
    away_name = away_name_node.get_text(" ", strip=True) if away_name_node else ""
    home_score = home_score_node.get_text(" ", strip=True) if home_score_node else "-"
    away_score = away_score_node.get_text(" ", strip=True) if away_score_node else "-"

    home_goals: list[str] = []
    away_goals: list[str] = []

    for minute_node in events_root.select(".event_min"):
        row = minute_node.parent
        minute = minute_node.get_text(" ", strip=True)

        if row.select_one(".event_ht .live_goal"):
            scorer_node = row.select_one(".event_ht .img16 span a") or row.select_one(".event_ht .img16 span")
            scorer = scorer_node.get_text(" ", strip=True) if scorer_node else "Гол"
            home_goals.append(f"{scorer} {minute}")

        if row.select_one(".event_at .live_goal"):
            scorer_node = row.select_one(".event_at .img16 span a") or row.select_one(".event_at .img16 span")
            scorer = scorer_node.get_text(" ", strip=True) if scorer_node else "Гол"
            away_goals.append(f"{scorer} {minute}")

    comment = f"{home_name} {home_score}:{away_score} {away_name}".strip()
    if home_goals or away_goals:
        home_part = "; ".join(home_goals) if home_goals else "-"
        away_part = "; ".join(away_goals) if away_goals else "-"
        comment = f"{comment} ({home_part} - {away_part})"

    return comment


def random_wrong_clubs(clubs: dict[str, ClubInfo], exclude_slugs: set[str], amount: int) -> list[ClubInfo]:
    pool = [club for club in clubs.values() if club.slug not in exclude_slugs]
    if len(pool) < amount:
        raise ValueError("Недостаточно клубов для генерации неправильных ответов")
    return random.sample(pool, amount)


def make_team_question_row(
    columns: list[str],
    text_que: str,
    question_club: ClubInfo,
    correct_club: ClubInfo,
    comment: str,
    clubs: dict[str, ClubInfo],
) -> dict:
    row = {column: "" for column in columns}
    row["type"] = 2
    row["textQue"] = text_que
    row["slugQue"] = question_club.logo_slug
    row["comment"] = comment
    row["slugDlg"] = correct_club.logo_slug
    row["slugCorAns"] = correct_club.logo_slug
    row["txtCorAns"] = correct_club.display_name

    wrong_clubs = random_wrong_clubs(clubs, {question_club.slug, correct_club.slug}, 5)
    for index, wrong_club in enumerate(wrong_clubs, start=1):
        row[f"wrSlgAns{index}"] = wrong_club.logo_slug
        row[f"wrTxtAns{index}"] = wrong_club.display_name

    return row


def build_match_questions(
    clubs_pair: list[ClubInfo],
    columns: list[str],
    all_clubs: dict[str, ClubInfo],
    alias_map: dict[str, ClubInfo],
) -> pd.DataFrame:
    result_stages = parse_competition_stages(RESULTS_URL)
    schedule_stages = parse_competition_stages(SCHEDULE_URL)

    rows: list[dict] = []
    for club in clubs_pair:
        result_stage, last_game = find_club_game(result_stages, club, alias_map)
        next_stage, next_game = find_club_game(schedule_stages, club, alias_map)

        last_comment = build_match_comment(last_game["url"])
        if not last_comment:
            score = f"{last_game['score_home']}:{last_game['score_away']}"
            last_comment = f"{last_game['home']} {score} {last_game['away']}"

        next_comment = f"{next_stage['round']}. {next_game['status']}".strip(" .")

        rows.append(
            make_team_question_row(
                columns=columns,
                text_que=f"С какой командой играл {club.display_name} в прошлом туре?",
                question_club=club,
                correct_club=last_game["opponent"],
                comment=last_comment,
                clubs=all_clubs,
            )
        )
        rows.append(
            make_team_question_row(
                columns=columns,
                text_que=f"С кем сыграет {club.display_name} в следующем туре?",
                question_club=club,
                correct_club=next_game["opponent"],
                comment=next_comment,
                clubs=all_clubs,
            )
        )

    return pd.DataFrame(rows, columns=columns)


def main():
    if len(sys.argv) != 3:
        print('Использование: py create_quiz.py "Зенит" "Балтика"')
        sys.exit(1)

    club_home_ru = sys.argv[1].strip()
    club_guest_ru = sys.argv[2].strip()

    if not club_home_ru or not club_guest_ru:
        print("Ошибка: названия клубов не должны быть пустыми.")
        sys.exit(1)

    if not TARGET_EXCEL.exists():
        print(f"Ошибка: не найден файл {TARGET_EXCEL}")
        sys.exit(1)

    clubs_catalog = build_club_catalog()
    alias_map = build_alias_map(clubs_catalog)

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

    xl = pd.ExcelFile(TARGET_EXCEL)
    sheet_exists = new_sheet_name in xl.sheet_names

    template_df = xl.parse(xl.sheet_names[0])
    columns = list(template_df.columns)

    home_rows = build_quiz_rows(home_excel, columns)
    guest_rows = build_quiz_rows(guest_excel, columns)
    match_rows = build_match_questions([home_club, guest_club], columns, clubs_catalog, alias_map)
    result_df = pd.concat([home_rows, guest_rows, match_rows], ignore_index=True)

    writer_kwargs = {"engine": "openpyxl", "mode": "a"}
    if sheet_exists:
        writer_kwargs["if_sheet_exists"] = "replace"

    with pd.ExcelWriter(TARGET_EXCEL, **writer_kwargs) as writer:
        result_df.to_excel(writer, sheet_name=new_sheet_name, index=False)

    if sheet_exists:
        print(f"Лист '{new_sheet_name}' существовал и был пересоздан")
    else:
        print(f"Создан новый лист: {new_sheet_name}")
    print("Добавлены вопросы: 8+8 клубных и 4 матчевых (прошлый/следующий тур)")
    print(f"Файл: {TARGET_EXCEL}")


if __name__ == "__main__":
    main()
