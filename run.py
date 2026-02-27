import json
import subprocess
import sys
from pathlib import Path

import pandas as pd


ROOT_DIR = Path(__file__).resolve().parent
EXCEL_SCRIPT = ROOT_DIR / "RPL_V2" / "xlsx" / "ExcelToJson.py"
EXCEL_FILE = ROOT_DIR / "RPL_V2" / "xlsx" / "Матч Тура 25-26.xlsx"
MATCHDAY_DIR = ROOT_DIR / "RPL_V2" / "matchday"
MD_TEST_FILE = MATCHDAY_DIR / "md_test.json"


def main():
    if len(sys.argv) < 2:
        print('Использование: py run.py "Зенит - Балтика"')
        sys.exit(1)

    match_name = sys.argv[1].strip()
    if not match_name:
        print("Ошибка: название матча пустое.")
        sys.exit(1)

    if not EXCEL_SCRIPT.exists():
        print(f"Ошибка: не найден скрипт {EXCEL_SCRIPT}")
        sys.exit(1)

    if not EXCEL_FILE.exists():
        print(f"Ошибка: не найден Excel файл {EXCEL_FILE}")
        sys.exit(1)

    subprocess.run([sys.executable, str(EXCEL_SCRIPT)], cwd=ROOT_DIR, check=True)

    xl = pd.ExcelFile(EXCEL_FILE)
    if not xl.sheet_names:
        print("Ошибка: в Excel нет листов.")
        sys.exit(1)

    last_sheet = xl.sheet_names[-1]
    json_file_name = f"{last_sheet}.json"
    json_file_path = MATCHDAY_DIR / json_file_name

    if not json_file_path.exists():
        print(f"Ошибка: после конвертации не найден файл {json_file_path}")
        sys.exit(1)

    if not MD_TEST_FILE.exists():
        print(f"Ошибка: не найден файл {MD_TEST_FILE}")
        sys.exit(1)

    with open(MD_TEST_FILE, "r", encoding="utf-8") as file:
        md_test = json.load(file)

    md_test["name"] = match_name
    md_test["url"] = f"https://amaranth64.github.io/RPL_V2/matchday/{json_file_name}"
    md_test["picture"] = f"https://amaranth64.github.io/RPL_V2/matchday/images/{last_sheet}.webp"

    with open(MD_TEST_FILE, "w", encoding="utf-8") as file:
        json.dump(md_test, file, ensure_ascii=False, indent=3)

    print(f"Готово: {MD_TEST_FILE}")
    print(f"name = {match_name}")
    print(f"url = {md_test['url']}")
    print(f"picture = {md_test['picture']}")


if __name__ == "__main__":
    main()
