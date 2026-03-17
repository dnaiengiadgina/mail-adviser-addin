#!/usr/bin/env python3
"""
アイコン生成スクリプト
assetsフォルダに必要なアイコンPNGを生成します
実行: python3 generate_icons.py
依存: pip install Pillow
"""
from PIL import Image, ImageDraw, ImageFont
import os

def make_icon(size, path):
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    # 背景円（青）
    margin = int(size * 0.05)
    draw.ellipse([margin, margin, size-margin, size-margin], fill=(0, 120, 212, 255))
    # メールアイコン（白い封筒）
    p = int(size * 0.2)
    draw.rectangle([p, p+int(size*0.1), size-p, size-p], fill='white')
    draw.polygon([p, p+int(size*0.1), size//2, size//2, size-p, p+int(size*0.1)], fill=(0, 120, 212, 255))
    img.save(path)
    print(f"Saved: {path}")

os.makedirs('assets', exist_ok=True)
for sz in [16, 32, 64, 80, 128]:
    make_icon(sz, f'assets/icon-{sz}.png')

print("アイコン生成完了!")
