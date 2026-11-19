# -*- coding: utf-8 -*-
u"""開発フォルダのアプリを、この置き場（kotei-test または kotei）へ写す（2026-11-16 開発環境の整理 2.4：公開の道具を1つに）

    python tools\\公開する.py                 → 共通の部品へ移したアプリを全部写す（common・nippou・shousai・meibo・spnippou）
    python tools\\公開する.py nippou shousai  → 指定したものだけ
    python tools\\公開する.py --push          → 写したあと、コミットして GitHub へ送る（公開される）

決まり
  ・写す元は 開発\\工事総合管理 の、いちばん新しい版（VrNNN。OLD の中は見ない）の web フォルダ。
  ・アドレスに版の番号は入れない（nippou/ はいつも最新の工事日報）。版は画面の右上（__VER__）に入れる。
  ・common/config.js は写さない（テストと本番で違うのはこのファイルだけ。それぞれの置き場のものを残す）。
  ・足す・置きかえるだけで、消さない。元に無くなったファイルは「元に無い」と知らせるだけ（消すときはその都度確かめる）。
  ・テストの置き場（kotei-test）では、版の表示に「（テスト）」を付ける。
"""
import argparse
import datetime
import glob
import os
import re
import shutil
import subprocess
import sys

sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEST = os.path.basename(REPO) == "kotei-test"
DEV = os.path.join(os.path.dirname(REPO), u"工事総合管理")
ICONS = {"manifest.json", "icon-192.png", "icon-512.png", "apple-touch-icon.png"}
# id: (元のフォルダ（* は版）, 置き場のフォルダ, 写さないもの)
APPS = {
    "common": (u"07_共通部品/web", "common", {"config.js"}),
    "nippou": (u"02_日報アプリ/日報Vr*/web", "nippou", {"config.js", "lib"}),
    "shousai": (u"14_常用日報/常用日報Vr*/web", "nippou/shousai", {"config.js", "lib"} | ICONS),
    "meibo": (u"16_名簿/web", "meibo", {"config.js"}),
    "spnippou": (u"02_日報アプリ/スマホ試作/web", "sp-nippou", {"config.js"}),   # 日報スマホ（2026-11-19 本番へ）
}
TEXT = (".js", ".html", ".css", ".json", ".md", ".txt")


def newest(pattern):
    hits = [h for h in glob.glob(os.path.join(DEV, pattern)) if os.sep + "OLD" + os.sep not in h]
    if not hits:
        return None, ""
    def num(p):
        m = re.search(r"Vr(\d+)", p)
        return int(m.group(1)) if m else -1
    h = max(hits, key=num)
    m = re.search(r"(Vr\d+)", h)
    return h, (m.group(1) if m else "")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("apps", nargs="*")
    ap.add_argument("--push", action="store_true")
    a = ap.parse_args()
    ids = a.apps or list(APPS)
    bad = [i for i in ids if i not in APPS]
    if bad:
        sys.exit(u"知らないアプリ：%s（%s）" % (bad, "・".join(APPS)))
    if not os.path.exists(os.path.join(REPO, "common", "config.js")):
        sys.exit(u"common\\config.js がありません。この置き場の設定を先に手で作ってください（テストと本番で違う。写さない）")
    now = datetime.datetime.now()
    V = now.strftime("%Y%m%d%H%M")
    changed, notes = [], []
    for i in ids:
        pat, dest, skip = APPS[i]
        src, vr = newest(pat)
        if not src:
            sys.exit(u"元が見つかりません：" + pat)
        label = u"%s %s%s" % (vr or {"common": u"共通", "meibo": u"名簿", "spnippou": u"スマホ"}.get(i, i), now.strftime("%m/%d %H:%M"), u"（テスト）" if TEST else "")
        dd = os.path.join(REPO, dest)
        seen = set()
        for root, dirs, files in os.walk(src):
            rel_root = os.path.relpath(root, src)
            top = rel_root.split(os.sep)[0]
            if top in skip:
                continue
            dirs[:] = [d for d in dirs if not (rel_root == "." and d in skip) and d != "__pycache__"]
            for f in files:
                rel = os.path.normpath(os.path.join(rel_root, f))
                if rel in skip:
                    continue
                seen.add(rel.replace(os.sep, "/"))
                sp, dp = os.path.join(root, f), os.path.join(dd, rel)
                if f.endswith(TEXT):
                    new = open(sp, encoding="utf-8", newline="").read().replace("__VER__", label).replace("__V__", V)
                    old = open(dp, encoding="utf-8", newline="").read() if os.path.exists(dp) else None
                    if old is not None and re.sub(r"\d{12}|\d\d/\d\d \d\d:\d\d", "", old).replace("\r\n", "\n") == re.sub(r"\d{12}|\d\d/\d\d \d\d:\d\d", "", new).replace("\r\n", "\n"):
                        continue   # 版の印のほかは同じ
                    os.makedirs(os.path.dirname(dp), exist_ok=True)
                    open(dp, "w", encoding="utf-8", newline="").write(new)
                else:
                    if os.path.exists(dp) and open(sp, "rb").read() == open(dp, "rb").read():
                        continue
                    os.makedirs(os.path.dirname(dp), exist_ok=True)
                    shutil.copyfile(sp, dp)
                changed.append(u"%s/%s" % (dest, rel.replace(os.sep, "/")))
        # 版の印が変わったファイルも入れ直す（画面が新しい部品を読むように）：index.html は毎回書く
        for f in ("index.html",):
            sp = os.path.join(src, f)
            if os.path.exists(sp) and u"%s/%s" % (dest, f) not in changed:
                open(os.path.join(dd, f), "w", encoding="utf-8", newline="").write(open(sp, encoding="utf-8", newline="").read().replace("__VER__", label).replace("__V__", V))
                changed.append(u"%s/%s（版の印）" % (dest, f))
        if os.path.isdir(dd):
            for root, dirs, files in os.walk(dd):
                rel_root = os.path.relpath(root, dd)
                # 中に別のアプリの置き場があれば見ない（nippou の中の shousai）
                dirs[:] = [d for d in dirs if os.path.normpath(os.path.join(dest, rel_root, d)).replace(os.sep, "/") not in [x[1] for x in APPS.values()]]
                for f in files:
                    rel = os.path.normpath(os.path.join(rel_root, f)).replace(os.sep, "/")
                    if rel not in seen and rel.split("/")[0] not in skip and rel not in skip:
                        notes.append(u"%s/%s" % (dest, rel))
        print(u"%-8s %s → %s/（%s）" % (i, os.path.relpath(src, DEV), dest, label))
    print(u"写した：%d件" % len(changed))
    for c in changed:
        print(u"  " + c)
    if notes:
        print(u"元に無いファイル（消していない。要らなければ、確かめてから消す）：")
        for n in notes:
            print(u"  " + n)
    if a.push and changed:
        dests = [APPS[i][1] for i in ids]
        subprocess.run(["git", "-C", REPO, "add", "--"] + dests, check=True)
        msg = u"公開：%s（%s）" % ("・".join(ids), V)
        subprocess.run(["git", "-C", REPO, "commit", "-q", "-m", msg], check=True)
        subprocess.run(["git", "-C", REPO, "push", "-q"], check=True)
        print(u"送りました：" + msg)


if __name__ == "__main__":
    main()
