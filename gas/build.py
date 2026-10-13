# -*- coding: utf-8 -*-
"""
Apps Script に配信させる index.html を用意する。

ダッシュボードの本体は **リポジトリ直下の index.html の1つだけ**。
Google版（HtmlService）に配るには、同じものが Apps Script プロジェクトの
中にも要るので、ここでコピーを作る。

    python gas/build.py          … gas/index.html を作り直す
    python gas/build.py --check  … ずれていないか見るだけ（作らない）

使い方：
    python gas/build.py && clasp push

gas/index.html は**手で編集しない**こと。直すのは直下の index.html。
"""
import io
import os
import sys
import hashlib

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC = os.path.join(ROOT, "index.html")
DST = os.path.join(HERE, "index.html")

HEADER = (
    "<!-- このファイルは gas/build.py が作っています。手で編集しないでください。\n"
    "     直すのはリポジトリ直下の index.html です。 -->\n"
)


def body(text):
    """先頭の自動生成コメントを外した中身。"""
    return text[len(HEADER):] if text.startswith(HEADER) else text


def main():
    check = "--check" in sys.argv
    if not os.path.exists(SRC):
        print("× %s がありません" % SRC)
        return 1
    src = io.open(SRC, encoding="utf-8").read()
    want = HEADER + src

    cur = io.open(DST, encoding="utf-8").read() if os.path.exists(DST) else ""
    same = body(cur) == src

    h = hashlib.sha1(src.encode("utf-8")).hexdigest()[:10]
    if same:
        print("一致しています（index.html %d KB / sha %s）" % (len(src) / 1024, h))
        return 0
    if check:
        print("★ ずれています。`python gas/build.py` で作り直してください")
        return 2
    io.open(DST, "w", encoding="utf-8", newline="").write(want)
    print("gas/index.html を作り直しました（%d KB / sha %s）" % (len(src) / 1024, h))
    return 0


if __name__ == "__main__":
    sys.exit(main())
