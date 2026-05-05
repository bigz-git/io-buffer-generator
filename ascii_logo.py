# import subprocess

# result = subprocess.run(
#     ['jp2a', '--width=30', '--colors', 'Quad Plus Brand Logo.png'],
#     capture_output=True,
#     text=True
# )
# print(result.stdout)

import os

from PIL import Image

def image_to_ascii(image_path, width=30):
    chars = '@#S%?*+;:,. '
    # Resolve relative paths against this module's directory so the logo works
    # both when run from source and when installed as a package.
    if not os.path.isabs(image_path):
        candidate = os.path.join(os.path.dirname(os.path.abspath(__file__)), image_path)
        if os.path.exists(candidate):
            image_path = candidate
    img = Image.open(image_path).convert('L')  # grayscale
    aspect = img.height / img.width
    img = img.resize((width, int(width * aspect * 0.55)))
    
    result = ''
    for y in range(img.height):
        for x in range(img.width):
            brightness = img.getpixel((x, y))
            result += chars[int(brightness / 255 * (len(chars) - 1))]
        result += '\n'
    print(result)

# image_to_ascii('Quad Plus Brand Logo.png')