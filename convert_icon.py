from PIL import Image
import os

img_path = 'assets/xtreme.png'
ico_path = 'assets/xtreme.ico'

if os.path.exists(img_path):
    img = Image.open(img_path)
    # Resize and save as ico. ico files can contain multiple sizes.
    img.save(ico_path, format='ICO', sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    print(f"Successfully created {ico_path}")
else:
    print(f"Source image {img_path} not found.")
