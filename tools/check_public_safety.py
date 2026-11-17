#!/usr/bin/env python3
"""公開してよいかの検査（標準ライブラリだけで動く）。

リポジトリに入るファイルに、テナント・環境を特定できる値が残っていないかを調べる。
- SharePoint・onmicrosoft.com のテナントのドメイン（OneDrive のパスの形 user_xxx_onmicrosoft_com を含む）
- GUID（ハイフンなし・%2D 区切り・_ に接する形を含む）
- メールアドレス（example.com / example.jp / example.org の架空のアドレスだけを許す）
各行は、そのままの形・URL デコード・JSON の \\uXXXX・HTML の文字参照を戻した形を調べる（Base64 で埋め込まれた値は対象外）。

使い方:
  python3 tools/check_public_safety.py             # git が管理するファイル（追加予定を含む）の作業ツリーを検査
  python3 tools/check_public_safety.py --staged    # コミットに入る内容（ステージ）を検査（pre-commit 用）
  python3 tools/check_public_safety.py FILE ...    # 指定したファイルを検査
終了コード: 0＝見つからない／1＝見つかった／2＝検査できない物がある（合格にしない）
"""
import html
import pathlib
import re
import subprocess
import sys
import urllib.parse

HEX = "0-9a-fA-F"
SEP = r"(?:-|_|%2[dD])?"
PATTERNS = {
    "テナントのドメイン（SharePoint）": re.compile(r"[A-Za-z0-9-]+(?:-my|-admin)?\.sharepoint(?:-mil)?\.(?:com|us|de|cn)", re.I),
    "テナントのドメイン（onmicrosoft）": re.compile(r"[A-Za-z0-9-]+[._]onmicrosoft[._](?:com|us|de)", re.I),
    "GUID": re.compile(
        rf"(?<![{HEX}])[{HEX}]{{8}}{SEP}[{HEX}]{{4}}{SEP}[{HEX}]{{4}}{SEP}[{HEX}]{{4}}{SEP}[{HEX}]{{12}}(?![{HEX}])"
    ),
    # ハイフン区切りの形は、前に16進数の文字が付いていても検出する（例: id3f2504e0-...）
    "GUID（ハイフン区切り）": re.compile(rf"[{HEX}]{{8}}-[{HEX}]{{4}}-[{HEX}]{{4}}-[{HEX}]{{4}}-[{HEX}]{{12}}(?![{HEX}])"),
    "メールアドレス": re.compile(r"[A-Za-z0-9._%+-]+@([A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,})"),
}
EXAMPLE_DOMAINS = {"example.com", "example.jp", "example.org"}
# メールアドレスの形に見えるが、アドレスではないもの（OData の注釈・高解像度の画像名）
NOT_EMAIL = re.compile(r"(?:odata\.[A-Za-z]+|\d+x\.(?:png|jpe?g|gif|webp|svg))", re.I)  # @ の後ろ全体と照合する
ALLOW_FILE = pathlib.Path(__file__).with_name("public_safety_allow.txt")
SKIP = {"tools/check_public_safety.py", "tools/public_safety_allow.txt"}
IMAGE_SUFFIX = {".png", ".jpg", ".jpeg", ".gif", ".ico", ".webp"}  # SVG は文字なので検査する


class CannotCheck(Exception):
    pass


def git(*args):
    try:
        out = subprocess.run(["git", *args], capture_output=True, check=True)
    except (OSError, subprocess.CalledProcessError) as err:
        raise CannotCheck(f"git を実行できません: {' '.join(args)}") from err
    return out.stdout


def load_allow():
    """完全一致で許す値（公開の仕様にある固定の ID など）。1行に1つ、# から後はコメント。"""
    if not ALLOW_FILE.exists():
        return set()
    items = set()
    for line in ALLOW_FILE.read_text(encoding="utf-8").splitlines():
        value = line.split("#", 1)[0].strip()
        if value:
            items.add(value.lower())
    return items


def targets(argv):
    """(表示名, 中身のバイト列を返す関数) の一覧と、検査範囲の説明を返す。"""
    if argv and argv[0] == "--staged":
        root = git("rev-parse", "--show-toplevel").decode().strip()
        names = [n for n in git("-C", root, "diff", "--cached", "--name-only", "--diff-filter=ACMR", "-z").decode("utf-8").split("\0") if n]
        return [(n, (lambda n=n: git("-C", root, "show", f":{n}"))) for n in names], "ステージした内容"
    if argv:
        items = []
        for a in argv:
            p = pathlib.Path(a)
            if not p.is_file():
                raise CannotCheck(f"ファイルがありません: {a}")
            items.append((a, p.read_bytes))
        return items, "指定したファイル"
    root = pathlib.Path(git("rev-parse", "--show-toplevel").decode().strip())
    names = [n for n in git("-C", str(root), "ls-files", "-co", "--exclude-standard", "-z").decode("utf-8").split("\0") if n]
    return [(n, (root / n).read_bytes) for n in names if (root / n).is_file()], "git が管理するファイル（追加予定を含む）の作業ツリー全体"


def decode(data):
    if b"\0" in data:
        return None  # UTF-16 などは UTF-8 として正しく読めないので、検査できない扱いにする
    try:
        return data.decode("utf-8-sig")
    except UnicodeDecodeError:
        return None


def findings(text, allow):
    for no, line in enumerate(text.splitlines(), 1):
        # JSON の \uXXXX を文字に戻した形
        unescaped = re.sub(r"\\u([0-9a-fA-F]{4})", lambda m: chr(int(m.group(1), 16)), line)
        variants = {
            line,
            urllib.parse.unquote(line),
            unescaped,
            html.unescape(line),
            html.unescape(urllib.parse.unquote(unescaped)),
        }
        seen = set()
        for v in variants:
            for label, pat in PATTERNS.items():
                for m in pat.finditer(v):
                    token = m.group(0)
                    if token.lower() in allow:
                        continue
                    if label == "メールアドレス":
                        if NOT_EMAIL.fullmatch(m.group(1)) or m.group(1).lower() in EXAMPLE_DOMAINS:
                            continue
                    # GUID の2つの形の両方に当たっても1件として数える
                    key = ("GUID" if label.startswith("GUID") else label, token.lower())
                    if key not in seen:
                        seen.add(key)
                        yield no, label, token


def main(argv):
    try:
        items, scope = targets(argv)
    except CannotCheck as err:
        print(f"検査できません: {err}")
        return 2
    allow = load_allow()
    problems, unreadable, images = [], [], []
    for name, read in items:
        if name in SKIP:
            continue
        if pathlib.PurePosixPath(name).suffix.lower() in IMAGE_SUFFIX:
            images.append(name)
            continue
        try:
            text = decode(read())
        except (OSError, CannotCheck):
            text = None
        if text is None:
            unreadable.append(name)
            continue
        for no, label, token in findings(text, allow):
            problems.append(f"{name}:{no}: {label}: {token}")
    for p in problems:
        print(p)
    for u in unreadable:
        print(f"{u}: 検査できません（UTF-8 の文字として読めない。zip・PDF は展開・文字に直して検査する）")
    note = f"検査範囲: {scope}。画像 {len(images)} 件は中身を検査していない（公開前に目で確かめる）"
    if problems:
        print(f"NG: {len(problems)} 件（{note}）")
        return 1
    if unreadable:
        print(f"検査できないファイルが {len(unreadable)} 件あります（{note}）")
        return 2
    print(f"OK: 見つかりませんでした（{note}）")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
