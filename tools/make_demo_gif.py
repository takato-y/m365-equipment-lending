#!/usr/bin/env python3
"""docs/images の画面（氏名などは塗りつぶし済み）をつないで、申請から承認・通知までの流れの GIF を作る。

静止画をつないだもので、画面の操作を録画したものではない。
使い方:
  python3 tools/make_demo_gif.py            # docs/images/demo.gif を作る
"""
import pathlib

from PIL import Image, ImageDraw, ImageFont

ROOT = pathlib.Path(__file__).resolve().parent.parent
IMG = ROOT / "docs" / "images"
FONT = "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"
W, H, BAR = 960, 600, 64
BG, BAR_BG, BAR_FG = (245, 246, 248), (0, 78, 120), (255, 255, 255)

# (画像, 手順の説明, 表示する時間[ミリ秒])
FRAMES = [
    ("app-equipment.png", "1. 社員がアプリで空いている備品を探す", 2600),
    ("app-apply.png", "2. 期間と利用目的を入れて申請する", 2600),
    ("approval.png", "3. 承認者に承認依頼が届く（Power Automate の承認）", 2600),
    ("teams-approved.png", "4. 承認すると Teams に知らせる（氏名は載せない。画面は別の申請）", 2400),
    ("outlook-mail.png", "5. 申請者には結果をメールで知らせる", 2400),
    ("app-mine.png", "6. アプリの「自分の貸出」で予約済みになる", 2600),
    ("forms.png", "7. 故障は Forms で報告する", 2400),
    ("planner.png", "8. Planner に確認のタスクができる", 2600),
]


def frame(path, caption, font):
    canvas = Image.new("RGB", (W, H), BG)
    d = ImageDraw.Draw(canvas)
    d.rectangle((0, 0, W, BAR), fill=BAR_BG)
    d.text((24, BAR // 2), caption, font=font, fill=BAR_FG, anchor="lm")
    shot = Image.open(path).convert("RGB")
    box_w, box_h = W - 48, H - BAR - 40
    scale = min(box_w / shot.width, box_h / shot.height, 1.6)
    shot = shot.resize((round(shot.width * scale), round(shot.height * scale)), Image.LANCZOS)
    x = (W - shot.width) // 2
    y = BAR + 20 + (box_h - shot.height) // 2
    d.rectangle((x - 1, y - 1, x + shot.width, y + shot.height), outline=(200, 200, 205))
    canvas.paste(shot, (x, y))
    return canvas


def main():
    font = ImageFont.truetype(FONT, 24)
    frames = [frame(IMG / name, caption, font) for name, caption, _ in FRAMES]
    durations = [ms for _, _, ms in FRAMES]
    pal = [f.quantize(colors=256, method=Image.MEDIANCUT, dither=Image.NONE) for f in frames]
    out = IMG / "demo.gif"
    pal[0].save(out, save_all=True, append_images=pal[1:], duration=durations, loop=0, optimize=True)
    print(out.relative_to(ROOT), out.stat().st_size, "bytes")


if __name__ == "__main__":
    main()
