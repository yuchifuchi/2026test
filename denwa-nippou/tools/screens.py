"""画面の見本を作る（手順書を書くとき、画面の文言と見た目を確かめるため）。

  python3 tools/screens.py   → build/screens/*.png
"""
import os
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "tests"))
from harness import get, bridge, ROOT, DIST  # noqa: E402
import test_flow  # noqa: E402

OUT = os.path.join(ROOT, "build", "screens")
CHROME = "/opt/pw-browsers/chromium"


def shot(sim, folder, url, name, height=1100):
    r = sim.request("GET", url)
    d = os.path.join(OUT, folder)
    os.makedirs(os.path.join(d, "css"), exist_ok=True)
    shutil.copyfile(os.path.join(DIST, folder, "css", "style.css"), os.path.join(d, "css", "style.css"))
    html = os.path.join(d, name + ".html")
    open(html, "wb").write(r.body)
    png = os.path.join(OUT, name + ".png")
    subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--no-sandbox", "--hide-scrollbars",
                    f"--window-size=1280,{height}", f"--screenshot={png}", "file://" + html],
                   capture_output=True, timeout=120)
    return png


def main():
    if os.path.exists(OUT):
        shutil.rmtree(OUT)
    sim = test_flow.run()
    pid = test_flow.ids(sim)
    shot(sim, "part", "/nippou/part/", "01_part_menu", 600)
    shot(sim, "part", f"/nippou/part/entry.asp?d=2026-08-25&t={pid['佐藤']}", "02_entry", 1500)
    shot(sim, "part", f"/nippou/part/tasks.asp?d=2026-08-25&t={pid['山田']}", "03_tasks", 1000)
    shot(sim, "staff", "/nippou/staff/staff.asp", "04_staff_menu", 900)
    shot(sim, "staff", "/nippou/staff/daily.asp?d=2026-08-25", "05_daily", 1500)
    shot(sim, "staff", "/nippou/staff/summary.asp?d=2026-08-25", "06_summary", 1600)
    shot(sim, "staff", "/nippou/staff/check.asp?d=2026-08-25", "07_check", 1100)
    shot(sim, "staff", "/nippou/staff/master.asp?t=tanto", "08_master", 800)
    shot(sim, "staff", "/nippou/staff/master.asp?t=tanto&id=new", "09_master_new", 900)
    shot(sim, "staff", "/nippou/staff/setup_check.asp", "10_setup_check", 1600)
    shot(sim, "staff", "/nippou/staff/report.asp?d=2026-08-25", "11_report", 1500)
    bridge().shutdown()
    print("OK:", OUT)


if __name__ == "__main__":
    main()
