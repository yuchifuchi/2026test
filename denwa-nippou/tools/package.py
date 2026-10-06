"""納品物（電話応対日報_納品一式.zip）を作る。作る前に全ファイルを検査し、1 つでも引っかかったら zip を作らない。

  python3 tools/package.py                         … 初回：全部入りの zip
  python3 tools/package.py --since release/manifest_v1.json
                                                   … 2 回目以降：前回から変わったファイルだけの zip と「置き場所の表」

検査すること（手順書 9-5・7-7）:
  .asp .html .css .md   … UTF-8（BOM なし）
  .txt .vbs .bas        … CP932（Shift_JIS）・改行は CRLF
  web.config            … UTF-8（BOM なし）・改行は CRLF・XML として読める・IIS が禁止している区画が無い
  すべて                … 制御文字（タブ・改行以外）が混ざっていない
  .accdb                … 先頭 0x14 バイト目が 0x02（Access 2007-2016 形式）
  part / staff          … part に職員用の画面が無い。include と css は part と staff で同じ（config.asp の ROLE の 1 行だけ違う）
  zip                   … 日本語のファイル名に UTF-8 フラグ（0x800）を立てる
"""
import datetime
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
from build import DIST, STAFF_PAGES, COMMON_PAGES, INCLUDES, DB_NAME  # noqa: E402
from vbsim.sim import LOCKED_SECTIONS  # noqa: E402

PKG = os.path.join(ROOT, "build", "package")
STAGE = os.path.join(PKG, "電話応対日報_納品一式")
ZIP_NAME = "電話応対日報_納品一式.zip"
DOCS = os.path.join(ROOT, "docs")
RELEASE = os.path.join(ROOT, "release")

STAMP = datetime.datetime.now().timetuple()[:6]
UTF8_EXT = {".asp", ".html", ".css", ".md"}
CP932_EXT = {".txt", ".vbs", ".bas"}
BINARY_EXT = {".accdb", ".pdf", ".png"}


def to_cp932_crlf(text):
    text = text.replace("\r\n", "\n").replace("\n", "\r\n")
    return text.encode("cp932")


def stage():
    if os.path.exists(PKG):
        shutil.rmtree(PKG)
    os.makedirs(STAGE)
    # 0_最初にお読みください.txt（リポジトリには UTF-8 で置き、ここで CP932・CRLF にする）
    src = open(os.path.join(DOCS, "0_最初にお読みください.txt"), encoding="utf-8").read()
    open(os.path.join(STAGE, "0_最初にお読みください.txt"), "wb").write(to_cp932_crlf(src))
    shutil.copytree(DIST, os.path.join(STAGE, "nippou"))
    os.makedirs(os.path.join(STAGE, "手順書"))
    for fn in sorted(os.listdir(os.path.join(DOCS, "手順書"))):
        shutil.copyfile(os.path.join(DOCS, "手順書", fn), os.path.join(STAGE, "手順書", fn))
    os.makedirs(os.path.join(STAGE, "参考"))
    for fn in sorted(os.listdir(os.path.join(DOCS, "参考"))):
        shutil.copyfile(os.path.join(DOCS, "参考", fn), os.path.join(STAGE, "参考", fn))
    sample = os.path.join(ROOT, "build", "render", "normal.pdf")
    if os.path.exists(sample):
        shutil.copyfile(sample, os.path.join(STAGE, "参考", "帳票の見本.pdf"))
    mock = os.path.join(ROOT, "build", "mockup", "画面の見本.html")      # tools/mockup.py が作る
    if os.path.exists(mock):
        shutil.copyfile(mock, os.path.join(STAGE, "参考", "画面の見本.html"))


def check_file(path, rel):
    probs = []
    ext = os.path.splitext(path)[1].lower()
    name = os.path.basename(path).lower()
    raw = open(path, "rb").read()
    if ext in BINARY_EXT:
        if ext == ".accdb":
            if raw[4:19] != b"Standard ACE DB" or raw[0x14] != 0x02:
                probs.append(f"{rel}: .accdb の形式が Access 2007-2016（0x02）ではありません（0x{raw[0x14]:02X}）")
        return probs
    if raw.startswith(b"\xef\xbb\xbf"):
        probs.append(f"{rel}: 先頭に BOM があります")
    if name == "web.config" or ext in UTF8_EXT:
        try:
            text = raw.decode("utf-8")
        except UnicodeDecodeError as e:
            return [f"{rel}: UTF-8 として読めません（{e}）"]
    elif ext in CP932_EXT:
        try:
            text = raw.decode("cp932")
        except UnicodeDecodeError as e:
            return [f"{rel}: CP932（Shift_JIS）として読めません（{e}）"]
        if re.search(rb"(?<!\r)\n", raw) or re.search(rb"\r(?!\n)", raw):
            probs.append(f"{rel}: 改行が CRLF になっていません（メモ帳で崩れます）")
    else:
        return [f"{rel}: 想定していない種類のファイルです"]
    bad = [c for c in text if (ord(c) < 0x20 and c not in "\t\r\n") or ord(c) == 0x7F or c == "\ufeff" or 0x80 <= ord(c) < 0xA0]
    if bad:
        probs.append(f"{rel}: 制御文字が混ざっています（{[hex(ord(c)) for c in bad[:5]]}）")
    if "\r" in text.replace("\r\n", ""):
        probs.append(f"{rel}: CR だけの改行があります")
    if name == "web.config":
        if re.search(rb"(?<!\r)\n", raw):
            probs.append(f"{rel}: 改行が CRLF になっていません")
        try:
            root = ET.fromstring(raw)
            for parent, sec in LOCKED_SECTIONS:
                node = root
                for part in parent.split("/"):
                    node = node.find(part) if node is not None else None
                if node is not None and node.find(sec) is not None:
                    probs.append(f"{rel}: IIS の既定で書けない <{sec}> があります（サイト全体が 500.19 になる）")
        except ET.ParseError as e:
            probs.append(f"{rel}: XML として読めません（{e}）")
    if ext == ".asp" and "\r\n" in text and "\n" in text.replace("\r\n", ""):
        probs.append(f"{rel}: 改行が LF と CRLF で混ざっています")
    return probs


def check_layout():
    probs = []
    n = os.path.join(STAGE, "nippou")
    for p in STAFF_PAGES:
        if os.path.exists(os.path.join(n, "part", p)):
            probs.append(f"part に職員用の画面 {p} があります")
        if not os.path.exists(os.path.join(n, "staff", p)):
            probs.append(f"staff に {p} がありません")
    for role in ("part", "staff"):
        for p in COMMON_PAGES + ["web.config", "css/style.css"] + [f"include/{x}" for x in INCLUDES]:
            if not os.path.exists(os.path.join(n, role, p)):
                probs.append(f"{role} に {p} がありません")
    for x in INCLUDES:
        a = open(os.path.join(n, "part", "include", x), "rb").read()
        b = open(os.path.join(n, "staff", "include", x), "rb").read()
        if x == "config.asp":
            da = [l for l in a.splitlines() if l not in b.splitlines()]
            if [l.strip() for l in da] != [b'Const ROLE = "part"']:
                probs.append("part と staff の config.asp が ROLE の 1 行以外でも違います")
        elif a != b:
            probs.append(f"part と staff の include/{x} が違います（同じものを置く決まり）")
    for p in COMMON_PAGES + ["css/style.css"]:
        if open(os.path.join(n, "part", p), "rb").read() != open(os.path.join(n, "staff", p), "rb").read():
            probs.append(f"part と staff の {p} が違います")
    if not os.path.exists(os.path.join(n, "data", DB_NAME)) or not os.path.exists(os.path.join(n, "data", "web.config")):
        probs.append("data に .accdb か web.config がありません")
    extra = set(os.listdir(os.path.join(n, "data"))) - {DB_NAME, "web.config"}
    if extra:
        probs.append(f"data に余計なファイルがあります: {extra}")
    for x in ("帳票の見本.pdf", "画面の見本.html"):
        if not os.path.exists(os.path.join(STAGE, "参考", x)):
            probs.append(f"参考 に {x} がありません（tools/render_sheet.py・tools/mockup.py を先に流す）")
    return probs


def manifest(base):
    out = {}
    for dp, _, fns in os.walk(base):
        for fn in fns:
            p = os.path.join(dp, fn)
            rel = os.path.relpath(p, base).replace(os.sep, "\\")
            out[rel] = hashlib.sha256(open(p, "rb").read()).hexdigest()
    return dict(sorted(out.items()))


def write_zip(base, files, zpath):
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        dirs = set()
        for rel in files:
            parts = rel.split("\\")[:-1]
            for i in range(1, len(parts) + 1):
                dirs.add("/".join(parts[:i]) + "/")
        for d in sorted(dirs):
            zi = zipfile.ZipInfo(d, date_time=STAMP)
            zi.flag_bits |= 0x800
            zi.external_attr = 0x10
            z.writestr(zi, b"")
        for rel in files:
            zi = zipfile.ZipInfo(rel.replace("\\", "/"), date_time=STAMP)
            zi.flag_bits |= 0x800      # 日本語のファイル名を UTF-8 と明示（エクスプローラーで化けないように）
            zi.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(zi, open(os.path.join(base, *rel.split("\\")), "rb").read())
    # 書いたものを読み直して確かめる
    with zipfile.ZipFile(zpath) as z:
        jp = 0
        for info in z.infolist():
            if not info.filename.isascii():
                jp += 1
                if not info.flag_bits & 0x800:
                    raise SystemExit(f"NG: zip の {info.filename} に UTF-8 フラグがありません（エクスプローラーで名前が化けます）")
        if jp == 0:
            raise SystemExit("NG: 日本語の名前のファイルが zip にありません（組み立てが違う）")
        if z.testzip() is not None:
            raise SystemExit("NG: zip が壊れています")
        return len(z.infolist())


def run_checks():
    probs = []
    for dp, _, fns in os.walk(STAGE):
        for fn in fns:
            p = os.path.join(dp, fn)
            probs += check_file(p, os.path.relpath(p, STAGE))
    probs += check_layout()
    r = subprocess.run([sys.executable, os.path.join(HERE, "lint_all.py"), os.path.join(STAGE, "nippou")], capture_output=True, text=True)
    if r.returncode != 0:
        probs.append("静的検査に通りません:\n" + r.stdout)
    return probs


def never_update(rel):
    """更新に決して入れないもの。
      include\\config.asp … 現地で書き換えた設定（DB_PATH など）が消える
      data\\*.accdb       … 現地の実データが、空のデータベースで上書きされて消える"""
    return rel.endswith("include\\config.asp") or (rel.startswith("nippou\\data\\") and rel.lower().endswith(".accdb"))


def db_shape(path=os.path.join(ROOT, "db", "schema.sql")):
    """データベースの形（表・列・索引・関連の定義）の照合値。初期データ（INSERT）は含めない。
    .accdb のファイルそのものは作るたびに中の日時などが変わるので、形の比べには使えない。"""
    text = "\n".join(l for l in open(path, encoding="utf-8").read().splitlines() if not l.strip().startswith("--"))
    stmts = [re.sub(r"\s+", " ", x).strip() for x in text.split(";")]
    ddl = [x for x in stmts if x and not x.upper().startswith("INSERT")]
    return hashlib.sha256("\n".join(ddl).encode("utf-8")).hexdigest()


def read_manifest(path):
    """前回の納品の一覧。最初の版（manifest_v1）はファイルの一覧だけで、形の照合値が無い。"""
    raw = json.load(open(path, encoding="utf-8"))
    if "files" in raw and isinstance(raw["files"], dict):
        return raw["files"], raw.get("db_shape")
    return raw, None


def update_blocker(old_shape, new_shape):
    """変わったファイルだけの更新では済まないとき True。
    データベースの形が変わると、新しい画面は新しい表を読むが、現地の .accdb は古い形のまま
    （.accdb は上書きしないので）。画面だけ置くと、その表を読む画面がエラーになる。
    前回の形が分からない（記録が無い）ときも、安全のため止める。"""
    return old_shape is None or old_shape != new_shape


def update_files(old, new):
    return [r for r, h in new.items() if old.get(r) != h and not never_update(r)]


def placement_table(changed, old, new):
    lines = ["電話応対日報 集計システム　更新ファイルの置き場所", "",
             "下の表のファイルを、表の「置く場所」に上書きしてください。", "表に無いファイルは、今のままにしてください。",
             "（include\\config.asp は、置き場所の設定が消えないよう、更新には入れていません）", ""]
    lines.append("置く場所（C:\\inetpub\\wwwroot\\ の中）")
    for rel in changed:
        tag = "（新しいファイル）" if rel not in old else ""
        lines.append(f"  {rel}{tag}")
    lines.append("")
    lines += ["■ 上書きしないもの",
              "  nippou\\data\\日報集計_be.accdb（今までのデータが入っています。更新には入っていません）",
              "  nippou\\part\\include\\config.asp と nippou\\staff\\include\\config.asp（置き場所の設定）", ""]
    diff_pairs = [r for r in changed if r.endswith("web.config") and ("\\part\\" in r or "\\staff\\" in r)]
    if diff_pairs:
        lines += ["■ part 用と staff 用で中身が違うファイル",
                  "  nippou\\part\\web.config と nippou\\staff\\web.config は、エラーの説明画面の場所だけが違います。",
                  "  取り違えると：エラーのときに、もう片方のフォルダの説明画面が出ます（使えはしますが、",
                  "  設置チェックで「web.config」が黄色になります）。必ず同じ名前のフォルダに置いてください。", ""]
    removed = [r for r in old if r not in new and not r.endswith("config.asp")]
    if removed:
        lines += ["■ 要らなくなったファイル（消してください）"] + [f"  {r}" for r in removed] + [""]
    lines += ["置いたあと：職員用の「設置チェック」（.../nippou/staff/setup_check.asp）を開き、",
              "赤い行が無いことを確かめてください。"]
    return "\r\n".join(lines) + "\r\n"


def main():
    since = None
    if len(sys.argv) > 2 and sys.argv[1] == "--since":
        since = sys.argv[2]
    stage()
    probs = run_checks()
    if probs:
        for p in probs:
            print("NG:", p)
        raise SystemExit(f"NG: 納品前の検査で {len(probs)} 件。zip は作りません")
    new = manifest(STAGE)
    record = {"db_shape": db_shape(), "files": new}
    os.makedirs(RELEASE, exist_ok=True)
    if since is None:
        n = write_zip(STAGE, list(new), os.path.join(PKG, ZIP_NAME))
        json.dump(record, open(os.path.join(PKG, "manifest.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        print(f"OK: 納品前の検査（{len(new)} ファイル）")
        print(f"OK: {os.path.join(PKG, ZIP_NAME)}（{n} 項目・UTF-8 フラグ付き）")
        return
    old, old_shape = read_manifest(since)
    changed = update_files(old, new)
    if update_blocker(old_shape, record["db_shape"]):
        why = "前回の一覧に形の記録がありません" if old_shape is None else "表・列の定義が前回と違います"
        raise SystemExit("NG: データベースの形（db/schema.sql）が前回と同じと言えないので、更新用の zip は作りません（" + why + "）。"
                         "画面だけを置くと、新しい表を読む画面がエラーになります。全部入りの zip で設置し直すか、"
                         "現地のデータを残したまま形を変える手順（表・列の追加）を別に用意してください。")
    if not changed:
        print("OK: 前回から変わったファイルはありません（更新用の zip は作りません）")
        return
    cfg_changed = [r for r, h in new.items() if r.endswith("include\\config.asp") and old.get(r) != h]
    if cfg_changed:
        print("注意: config.asp の中身（設定の項目）が変わっています。更新には入れないので、置き場所の表に書き足す手順を手順書に用意してください:", cfg_changed)
    table = placement_table(changed, old, new)
    tpath = os.path.join(STAGE, "置き場所の表.txt")
    open(tpath, "wb").write(to_cp932_crlf(table))
    p = check_file(tpath, "置き場所の表.txt")
    if p:
        raise SystemExit("NG: " + str(p))
    stamp = datetime.date.today().strftime("%Y%m%d")
    zpath = os.path.join(PKG, f"電話応対日報_更新_{stamp}.zip")
    n = write_zip(STAGE, ["置き場所の表.txt"] + changed, zpath)
    json.dump(record, open(os.path.join(PKG, "manifest.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"OK: 変わったファイル {len(changed)} 個の更新用 zip: {zpath}")


if __name__ == "__main__":
    main()
