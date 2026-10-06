"""検査用の Java 部品を Maven Central から取得する（開発者の作業環境用。納品物には入らない）。

なぜ 2 系統あるか:
  - .accdb の生成は Jackcess 4.0.5 を使う（指定どおり）。
  - SQL の実行確認は UCanAccess 5.0.1 を使うが、これは Jackcess 3.0.1 に依存して作られている。
    4.0.5 と同じクラスパスに混ぜると API の差で落ちることがあるため、置き場所を分ける。
取得したファイルは Maven Central の .sha1 と照合し、一致しなければ使わない。
"""
import hashlib
import os
import sys
import urllib.request

BASE = "https://repo1.maven.org/maven2"
HERE = os.path.dirname(os.path.abspath(__file__))
LIB = os.path.join(HERE, "lib")

SETS = {
    "builder": [
        ("com.healthmarketscience.jackcess", "jackcess", "4.0.5"),
        ("org.apache.commons", "commons-lang3", "3.10"),
        ("commons-logging", "commons-logging", "1.2"),
    ],
    "ucanaccess": [
        ("net.sf.ucanaccess", "ucanaccess", "5.0.1"),
        ("com.healthmarketscience.jackcess", "jackcess", "3.0.1"),
        ("org.hsqldb", "hsqldb", "2.5.0"),
        ("org.apache.commons", "commons-lang3", "3.8.1"),
        ("commons-logging", "commons-logging", "1.2"),
    ],
}


def fetch(url):
    with urllib.request.urlopen(url, timeout=60) as r:
        return r.read()


def main():
    for name, arts in SETS.items():
        d = os.path.join(LIB, name)
        os.makedirs(d, exist_ok=True)
        for g, a, v in arts:
            fn = f"{a}-{v}.jar"
            path = os.path.join(d, fn)
            url = f"{BASE}/{g.replace('.', '/')}/{a}/{v}/{fn}"
            want = fetch(url + ".sha1").decode().split()[0].strip()
            if os.path.exists(path):
                if hashlib.sha1(open(path, "rb").read()).hexdigest() == want:
                    continue
            data = fetch(url)
            got = hashlib.sha1(data).hexdigest()
            if got != want:
                print(f"NG: {fn} の照合値が一致しません", file=sys.stderr)
                sys.exit(1)
            with open(path, "wb") as f:
                f.write(data)
            print(f"取得: {name}/{fn}")
    print("OK: 検査用の部品はそろっています")


if __name__ == "__main__":
    main()
