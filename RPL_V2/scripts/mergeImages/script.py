#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
from PIL import Image

def merge_images(path1: str, path2: str, output_path: str = "merged.png"):
    # Открываем изображения
    img1 = Image.open(path1).convert("RGBA")
    img2 = Image.open(path2).convert("RGBA")

    # Приводим к размеру 800×800
    img1 = img1.resize((800, 800), Image.LANCZOS)
    img2 = img2.resize((800, 800), Image.LANCZOS)

    # Создаём новое полотно 1600×800 с прозрачным фоном
    merged = Image.new("RGBA", (1600, 800), (0, 0, 0, 0))

    # Вставляем по бокам с учётом альфа-канала
    merged.paste(img1, (0, 0), img1)
    merged.paste(img2, (800, 0), img2)

    # Сохраняем в PNG, чтобы сохранить прозрачность
    merged.save(output_path, format="PNG")
    print(f"Сохранено: {output_path}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Использование: python merge.py <путь_к_изображению1> <путь_к_изображению2> [<путь_к_результату>]")
        sys.exit(1)

    path1 = sys.argv[1]
    path2 = sys.argv[2]
    output = sys.argv[3] if len(sys.argv) >= 4 else "merged.png"

    merge_images(path1, path2, output)
