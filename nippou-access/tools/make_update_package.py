# -*- coding: utf-8 -*-
"""前に納品した一式と見比べて、「変わったファイルだけ」の zip を作る。

一式は 4 MB を超えるので、直しが出るたびに全部を送り直すのは重い。
前のコミットの状態で納品物をもう一度組み立て、いまの納品物と 1 バイトずつ
比べて、変わったものと増えたものだけを同じフォルダ構成で詰める。
受け取った側は、その形のまま上書きコピーすれば更新できる。

    python3 tools/make_update_package.py <前のコミット> [出力先.zip]
"""
import filecmp
import io
import os
import shutil
import subprocess
import sys
import zipfile
from datetime import date

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
NAME = "電話応対日報_更新分"


def build_package(repo_dir, out_dir):
    """指定のリポジトリで make_package.py を動かし、組み上がった納品物の場所を返す。"""
    os.makedirs(out_dir, exist_ok=True)
    subprocess.check_call([sys.executable, os.path.join(repo_dir, "tools", "make_package.py"),
                           out_dir],
                          cwd=repo_dir, stdout=subprocess.DEVNULL)
    return os.path.join(out_dir, "電話応対日報_納品一式")


def walk(base):
    out = {}
    for root, dirs, files in os.walk(base):
        dirs.sort()
        for f in sorted(files):
            p = os.path.join(root, f)
            out[os.path.relpath(p, base).replace(os.sep, "/")] = p
    return out


def main():
    if len(sys.argv) < 2:
        sys.exit("使い方: python3 tools/make_update_package.py <前のコミット> [出力先.zip]")
    prev_ref = sys.argv[1]
    out_zip = sys.argv[2] if len(sys.argv) > 2 else "/tmp/%s.zip" % NAME

    work = "/tmp/_update_pkg"
    shutil.rmtree(work, ignore_errors=True)
    os.makedirs(work)
    prev_repo = os.path.join(work, "prev_repo")

    # 前の状態を取り出す (作業ツリーは触らない)
    top = subprocess.check_output(["git", "rev-parse", "--show-toplevel"],
                                  cwd=HERE, text=True).strip()
    sub = os.path.relpath(HERE, top)
    subprocess.check_call(["git", "worktree", "add", "--detach", prev_repo, prev_ref],
                          cwd=top, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        old = build_package(os.path.join(prev_repo, sub), os.path.join(work, "old"))
        new = build_package(HERE, os.path.join(work, "new"))

        oldf, newf = walk(old), walk(new)
        changed, added = [], []
        for rel, p in sorted(newf.items()):
            if rel not in oldf:
                added.append(rel)
            elif not filecmp.cmp(oldf[rel], p, shallow=False):
                changed.append(rel)
        removed = sorted(set(oldf) - set(newf))

        note = make_note(changed, added, removed)
        base = os.path.join(work, NAME)
        for rel in changed + added:
            dst = os.path.join(base, rel.replace("/", os.sep))
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            shutil.copy2(newf[rel], dst)
        with open(os.path.join(base, "0_更新のしかた.txt"), "wb") as f:
            f.write(note.replace("\n", "\r\n").encode("cp932"))

        zip_dir(base, out_zip)
        print("変更 %d / 追加 %d / 削除 %d" % (len(changed), len(added), len(removed)))
        for rel in changed + added:
            print("   ", rel)
        for rel in removed:
            print("  (消えた)", rel)
        print("作成:", out_zip, "(%.2f MB)" % (os.path.getsize(out_zip) / 1048576.0))
    finally:
        subprocess.call(["git", "worktree", "remove", "--force", prev_repo],
                        cwd=top, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return 0


def make_note(changed, added, removed):
    lines = []
    a = lines.append
    a("=" * 60)
    a(" 電話応対日報 集計システム   更新分 (%s)" % date.today().strftime("%Y/%m/%d"))
    a("=" * 60)
    a("")
    a(" 前にお渡しした「電話応対日報_納品一式」のうち、")
    a(" 変わったファイルだけを、同じフォルダの形で入れてあります。")
    a("")
    a(" この zip の中身を、納品一式のフォルダに")
    a(" 「そのまま上書きコピー」してください。")
    a(" フォルダの形が同じなので、まとめてドラッグして上書きで構いません。")
    a(" （このファイル 0_更新のしかた.txt はコピーしなくて結構です）")
    a("")
    a("-" * 60)
    a(" すでにサーバーに置いて使い始めている場合")
    a("-" * 60)
    a("")
    a(" ・03_Webサイト_ASP\\wwwroot の中身")
    a("     → C:\\inetpub\\wwwroot\\nippou\\ にも同じように上書きしてください。")
    a("       ただし include\\db.asp と include\\pdf.asp は、設置のときに")
    a("       書き換えた設定が入っています。上書き前に中身を控えてください")
    a("       (この更新分に入っていなければ、触る必要はありません)。")
    a("")
    a(" ・02_データベース_Access\\日報集計_be.accdb")
    a("     → まだ入力を始めていなければ、共有フォルダのものと差し替えてください。")
    a("       すでに入力を始めている場合は差し替えないでください。")
    a("       (差し替えると入力済みのデータが消えます)")
    a("       その場合は 02_データベース_Access\\任意_すでに使い始めている場合\\ の")
    a("       .vbs を実行すると、足りない欄だけを足せます (データは消えません)。")
    a("")
    a("-" * 60)
    a(" 入っているファイル")
    a("-" * 60)
    a("")
    if changed:
        a(" ● 変わったもの (%d 件)" % len(changed))
        for rel in changed:
            a("     " + rel.replace("/", "\\"))
        a("")
    if added:
        a(" ● 増えたもの (%d 件)" % len(added))
        for rel in added:
            a("     " + rel.replace("/", "\\"))
        a("")
    if removed:
        a(" ● 要らなくなったもの (%d 件) ― 消していただいて構いません" % len(removed))
        for rel in removed:
            a("     " + rel.replace("/", "\\"))
        a("")
    return "\n".join(lines) + "\n"


def zip_dir(base, zip_path):
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for root, dirs, files in os.walk(base):
            dirs.sort()
            for f in sorted(files):
                p = os.path.join(root, f)
                arc = os.path.relpath(p, os.path.dirname(base))
                info = zipfile.ZipInfo(arc.replace(os.sep, "/"))
                info.compress_type = zipfile.ZIP_DEFLATED
                info.flag_bits |= 0x800          # 名前は UTF-8
                info.date_time = (2026, 1, 1, 0, 0, 0)
                info.external_attr = 0o644 << 16
                with open(p, "rb") as fp:
                    z.writestr(info, fp.read())


if __name__ == "__main__":
    sys.exit(main())
