import os
from PIL import Image


for pic in os.listdir():
    if pic[-3:] != ".py" and pic != "out":
        imagePath = "./" + pic
        outputPath = "./out/" + pic.split('.')[0].lower () + ".webp"
        quality = "85"

        im = Image.open(imagePath)
        im.save(outputPath, 'webp', quality=quality)
