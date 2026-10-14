"""Браузер для аудита в облачном контейнере: одна вкладка Chromium, живущая между командами.

Chromium запускается в фоне с CDP-портом; каждая команда подключается к той же вкладке,
поэтому состояние (вход, открытая симуляция, введённые значения) сохраняется между вызовами.

    python3 istok-audit/tools/browser.py start [--width 1920 --height 1080]
    python3 istok-audit/tools/browser.py login                  # ISTOK_LOGIN / ISTOK_PASSWORD
    python3 istok-audit/tools/browser.py login --student        # ISTOK_STUDENT_LOGIN / ..._PASSWORD
    python3 istok-audit/tools/browser.py open URL --sim 3
    python3 istok-audit/tools/browser.py size 1366 768
    python3 istok-audit/tools/browser.py shot SIM-03-07 [--selector CSS] [--zoom 2] [--full] [--mark]
    python3 istok-audit/tools/browser.py stop

Шаги проверки пишутся отдельными скриптами:

    sys.path.insert(0, "istok-audit/tools"); import browser
    with browser.page(sim=3) as p:
        p.fill("#mass", "1e9"); p.click("text=Пуск"); p.wait_for_timeout(1000)
        browser.shot(p, "SIM-03-07", selector="#result", zoom=2)

Консоль и сеть (ошибки и предупреждения JS, ответы 4xx/5xx, запросы дольше 3 с) пишутся
в STATE/console.jsonl с номером симуляции; сводку в console.md даёт команда `log --sim N`.

Логин и пароль берутся только из переменных окружения, не печатаются и не пишутся на диск.
Профиль браузера с cookie сессии лежит в STATE вне репозитория.
"""

import argparse
import base64
import contextlib
import json
import os
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

from playwright.sync_api import sync_playwright

AUDIT = Path(__file__).resolve().parent.parent
SHOTS = AUDIT / "screenshots"
STATE = Path(os.environ.get("ISTOK_STATE", Path.home() / ".istok-audit-state"))
BASE = os.environ.get("ISTOK_BASE", "https://edu-dev.istok.systems").rstrip("/")
CHROME = os.environ.get("ISTOK_CHROME", "/opt/pw-browsers/chromium")
PORT = int(os.environ.get("ISTOK_CDP_PORT", "9222"))
SLOW_MS = 3000


def cdp_up():
    try:
        urllib.request.urlopen("http://127.0.0.1:%d/json/version" % PORT, timeout=1)
        return True
    except OSError:
        return False


def start(width, height):
    if cdp_up():
        print("Chromium уже запущен")
        return
    STATE.mkdir(parents=True, exist_ok=True)
    args = [CHROME, "--headless=new", "--remote-debugging-port=%d" % PORT,
            "--user-data-dir=%s" % (STATE / "profile"), "--no-first-run",
            "--no-default-browser-check", "--window-size=%d,%d" % (width, height),
            "--lang=ru-RU", "--hide-scrollbars", "about:blank"]
    if os.geteuid() == 0:
        args.insert(1, "--no-sandbox")  # контейнер работает от root
    proxy = os.environ.get("HTTPS_PROXY") or os.environ.get("https_proxy")
    if proxy:
        args.append("--proxy-server=%s" % proxy)
        no_proxy = os.environ.get("NO_PROXY") or os.environ.get("no_proxy")
        if no_proxy:
            args.append("--proxy-bypass-list=%s" % no_proxy.replace(",", ";"))
    log = open(STATE / "chromium.log", "ab")
    subprocess.Popen(args, stdout=log, stderr=log, start_new_session=True)
    for _ in range(50):
        if cdp_up():
            with page() as pg:
                got = set_size(pg, width, height)
            print("Chromium запущен, порт %d, область страницы %dx%d" % (PORT, *got))
            return
        time.sleep(0.2)
    sys.exit("Chromium не поднялся, см. %s" % (STATE / "chromium.log"))


def stop():
    subprocess.run(["pkill", "-f", "remote-debugging-port=%d" % PORT])
    print("Chromium остановлен")


def _log(sim, kind, **data):
    STATE.mkdir(parents=True, exist_ok=True)
    rec = dict(t=time.strftime("%H:%M:%S"), sim=sim, kind=kind, **data)
    with open(STATE / "console.jsonl", "a", encoding="utf-8") as f:
        f.write(json.dumps(rec, ensure_ascii=False) + "\n")


def _attach(pg, sim):
    def on_console(msg):
        if msg.type in ("error", "warning"):
            loc = msg.location or {}
            _log(sim, "console." + msg.type, text=msg.text, url=pg.url,
                 src="%s:%s" % (loc.get("url", ""), loc.get("lineNumber", "")))

    def on_response(resp):
        timing = resp.request.timing
        ms = timing.get("responseEnd", -1)
        if resp.status >= 400:
            _log(sim, "http.%d" % resp.status, url=resp.url, method=resp.request.method)
        if ms > SLOW_MS:
            _log(sim, "slow", url=resp.url, ms=round(ms))

    pg.on("console", on_console)
    pg.on("pageerror", lambda e: _log(sim, "pageerror", text=str(e), url=pg.url))
    pg.on("requestfailed", lambda r: _log(sim, "requestfailed", url=r.url,
                                          error=r.failure or ""))
    pg.on("response", on_response)


@contextlib.contextmanager
def page(sim=0):
    """Та же вкладка, что и в прошлых командах, с журналом консоли и сети."""
    if not cdp_up():
        sys.exit("Chromium не запущен: python3 istok-audit/tools/browser.py start")
    with sync_playwright() as pw:
        browser = pw.chromium.connect_over_cdp("http://127.0.0.1:%d" % PORT)
        ctx = browser.contexts[0]
        pg = ctx.pages[0] if ctx.pages else ctx.new_page()
        _attach(pg, sim)
        want = _wanted_size()
        if want and pg.url != "about:blank" and \
                pg.evaluate("[innerWidth, innerHeight]") != want:
            set_size(pg, *want)
        try:
            yield pg
        finally:
            pg.wait_for_timeout(300)  # дать долететь событиям консоли


def shot(pg, sid, selector=None, zoom=1.0, full=False, mark=False, margin=40):
    """Скриншот screenshots/<sid>.png. selector — область вокруг элемента, zoom — увеличение."""
    SHOTS.mkdir(parents=True, exist_ok=True)
    out = SHOTS / (sid + ".png")
    metrics = "({w: innerWidth, h: innerHeight, sx: scrollX, sy: scrollY," \
              " dw: document.documentElement.scrollWidth, dh: document.documentElement.scrollHeight})"
    vp = pg.evaluate(metrics)
    clip = dict(x=0, y=0, width=vp["w"], height=vp["h"])
    el = pg.locator(selector).first if selector else None
    if el:
        el.scroll_into_view_if_needed()
        vp = pg.evaluate(metrics)
        if mark:
            el.evaluate("e => { e.dataset.auditOutline = e.style.outline;"
                        " e.style.outline = '3px solid red'; }")
        box = el.bounding_box()
        x, y = max(box["x"] - margin, 0), max(box["y"] - margin, 0)
        clip = dict(x=x, y=y, width=min(box["width"] + 2 * margin, vp["w"] - x),
                    height=min(box["height"] + 2 * margin, vp["h"] - y))
    if full and not el:
        clip = dict(x=-vp["sx"], y=-vp["sy"], width=vp["dw"], height=vp["dh"])
    # Захват через CDP: координаты документа, область перерисовывается в масштабе zoom
    clip.update(x=clip["x"] + vp["sx"], y=clip["y"] + vp["sy"], scale=zoom)
    cdp = pg.context.new_cdp_session(pg)
    try:
        data = cdp.send("Page.captureScreenshot", dict(
            format="png", clip=clip, captureBeyondViewport=bool(full)))["data"]
    finally:
        cdp.detach()
        if el and mark:
            el.evaluate("e => { e.style.outline = e.dataset.auditOutline || ''; }")
    out.write_bytes(base64.b64decode(data))
    print("Сохранён", out.relative_to(AUDIT.parent))
    return out


def _wanted_size():
    path = STATE / "viewport.json"
    return json.loads(path.read_text()) if path.exists() else None


def set_size(pg, width, height):
    """Размер окна так, чтобы область страницы (innerWidth × innerHeight) была width × height.

    Размер запоминается: каждое подключение сверяет его и поправляет, если окно уехало."""
    (STATE / "viewport.json").write_text(json.dumps([width, height]))
    cdp = pg.context.new_cdp_session(pg)
    try:
        win = cdp.send("Browser.getWindowForTarget")["windowId"]
        for _ in range(3):
            d = pg.evaluate("({w: outerWidth - innerWidth, h: outerHeight - innerHeight})")
            cdp.send("Browser.setWindowBounds", dict(windowId=win, bounds=dict(
                width=width + d["w"], height=height + d["h"], windowState="normal")))
            pg.wait_for_timeout(300)
            got = pg.evaluate("[innerWidth, innerHeight]")
            if got == [width, height]:
                break
    finally:
        cdp.detach()
    return got


def login(student=False):
    prefix = "ISTOK_STUDENT_" if student else "ISTOK_"
    user, pwd = os.environ.get(prefix + "LOGIN"), os.environ.get(prefix + "PASSWORD")
    if not user or not pwd:
        sys.exit("Нет %sLOGIN / %sPASSWORD в окружении. Их задаёт человек в настройках "
                 "окружения; в чат и в файлы они не попадают." % (prefix, prefix))
    with page() as pg:
        if "/login" not in pg.url:
            pg.goto(BASE + "/login", wait_until="networkidle")
        pw_field = pg.locator("input[type=password]:visible").first
        try:
            pw_field.wait_for(timeout=15000)
        except Exception:
            sys.exit("Поле пароля не найдено на %s — нужен скриншот и ручной разбор формы" % pg.url)
        form = pg.locator("form:has(input[type=password])").first
        scope = form if form.count() else pg
        user_field = scope.locator(
            "input:visible:not([type=password]):not([type=hidden])"
            ":not([type=checkbox]):not([type=radio]):not([type=submit])").first
        user_field.fill(user)
        pw_field.fill(pwd)
        submit = scope.locator("button[type=submit]:visible, input[type=submit]:visible").first
        if submit.count():
            submit.click()
        else:
            pw_field.press("Enter")
        pg.wait_for_load_state("networkidle")
        pg.wait_for_timeout(1000)
        still = pg.locator("input[type=password]:visible").count() > 0
        print("Адрес после входа:", pg.url)
        print("Форма входа всё ещё на странице — вход не удался?" if still else "Вход выполнен")


def summarize(sim):
    path = STATE / "console.jsonl"
    if not path.exists():
        return
    seen = set()
    for line in path.read_text(encoding="utf-8").splitlines():
        rec = json.loads(line)
        if rec["sim"] != sim:
            continue
        key = (rec["kind"], rec.get("text") or rec.get("url"))
        if key in seen:
            continue
        seen.add(key)
        detail = rec.get("text") or rec.get("url")
        extra = rec.get("ms") or rec.get("error") or rec.get("src") or ""
        print("- [%s] %s %s" % (rec["kind"], detail, extra))


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("start")
    s.add_argument("--width", type=int, default=1920)
    s.add_argument("--height", type=int, default=1080)
    sub.add_parser("stop")
    s = sub.add_parser("login")
    s.add_argument("--student", action="store_true")
    s = sub.add_parser("open")
    s.add_argument("url")
    s.add_argument("--sim", type=int, default=0)
    s = sub.add_parser("size")
    s.add_argument("width", type=int)
    s.add_argument("height", type=int)
    s = sub.add_parser("shot")
    s.add_argument("id")
    s.add_argument("--selector")
    s.add_argument("--zoom", type=float, default=1.0)
    s.add_argument("--full", action="store_true")
    s.add_argument("--mark", action="store_true")
    s.add_argument("--sim", type=int, default=0)
    s = sub.add_parser("log")
    s.add_argument("--sim", type=int, default=0)
    a = ap.parse_args()

    if a.cmd == "start":
        start(a.width, a.height)
    elif a.cmd == "stop":
        stop()
    elif a.cmd == "login":
        login(a.student)
    elif a.cmd == "open":
        url = a.url if a.url.startswith("http") else BASE + "/" + a.url.lstrip("/")
        with page(a.sim) as pg:
            t0 = time.time()
            pg.goto(url, wait_until="networkidle")
            print("%s — %.1f с, «%s»" % (pg.url, time.time() - t0, pg.title()))
    elif a.cmd == "size":
        with page() as pg:
            got = set_size(pg, a.width, a.height)
            print("Область страницы %dx%d, горизонтальный скролл: %s" % (
                *got,
                pg.evaluate("document.documentElement.scrollWidth > innerWidth")))
    elif a.cmd == "shot":
        with page(a.sim) as pg:
            shot(pg, a.id, a.selector, a.zoom, a.full, a.mark)
    elif a.cmd == "log":
        summarize(a.sim)


if __name__ == "__main__":
    main()
