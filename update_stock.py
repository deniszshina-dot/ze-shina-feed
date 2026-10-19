#!/usr/bin/env python3
"""Update price / stock in feed.yml from the WheelShop B2B CSV (windows-1251, tab separated).
Env: WS_EMAIL, WS_PASSWD. Only <price>, <quantity_in_stock>, available= and the catalog date change."""
import os, re, sys, urllib.parse, urllib.request
from datetime import datetime, timezone

MIN_SET = 4          # Prom sells only sets of 4
MAX_DROP = 0.4       # abort if more than 40% of offers would become unavailable
FEED = os.environ.get("FEED_PATH", "feed.yml")

def load_csv():
    path = os.environ.get("CSV_PATH")
    if path:
        raw = open(path, "rb").read()
    else:
        q = urllib.parse.urlencode({"email": os.environ["WS_EMAIL"], "passwd": os.environ["WS_PASSWD"]})
        with urllib.request.urlopen("https://www.wheelshop.com.ua/api/csv/wheels/1.1?" + q, timeout=90) as r:
            if r.status != 200:
                sys.exit(f"CSV http {r.status}")
            raw = r.read()
    text = raw.decode("windows-1251")
    lines = [l.rstrip("\r") for l in text.split("\n") if l.strip()]
    head = lines[0].split("\t")
    if len(head) < 12 or head[0] != "Виробник" or head[8] != "Кількість":
        sys.exit("unexpected CSV header: " + lines[0][:120])
    rows = [l.split("\t") for l in lines[1:]]
    rows = [r for r in rows if len(r) >= 12]
    if len(rows) < 500:
        sys.exit(f"CSV too short: {len(rows)} rows")
    return rows

def num(s):
    return float(str(s).replace(",", ".").strip())

def build_index(rows):
    idx = {}
    for c in rows:
        try:
            key = (c[1].strip().lower(), c[2].strip().lower(), round(num(c[3])), round(num(c[4]), 1), c[5].strip().lower())
            idx.setdefault(key, []).append((num(c[6]), int(float(c[8])), num(c[10])))
        except ValueError:
            continue
    return idx

def main():
    idx = build_index(load_csv())
    feed = open(FEED, encoding="utf-8").read()
    blocks = list(re.finditer(r'<offer id="(WS-\d+)"( selling_type="\w")? available="(true|false)">(.*?)</offer>', feed, re.S))
    if len(blocks) < 100:
        sys.exit("feed looks wrong: %d offers" % len(blocks))
    out, pos, changed, to_false, missing = [], 0, 0, 0, []
    for m in blocks:
        body = m.group(4)
        def par(n):
            x = re.search(r'<param name="%s">(.*?)</param>' % n, body)
            return x.group(1).strip() if x else ""
        name = re.search(r"<name>(.*?)</name>", body).group(1)
        model = re.sub(r"^Диск\s+", "", name).strip().lower()
        try:
            key = (model, par("Колір").lower(), round(num(par("Діаметр"))), round(num(par("Ширина")), 1), par("PCD").lower())
            et = num(par("ET"))
        except ValueError:
            out.append(feed[pos:m.end()]); pos = m.end(); continue
        cands = [c for c in idx.get(key, []) if abs(c[0] - et) <= 1.0]
        if cands:
            _, qty, price_disk = min(cands, key=lambda c: abs(c[0] - et))
            price = int(round(price_disk))   # price per disc (wholesale tier, min order 4)
        else:
            qty, price = 0, None
            missing.append(m.group(1))
        sellable = qty >= MIN_SET
        qty_out = qty if sellable else 0
        new = body
        new = re.sub(r"<quantity_in_stock>\d+</quantity_in_stock>", f"<quantity_in_stock>{qty_out}</quantity_in_stock>", new)
        if price:
            new = re.sub(r"<value>\d+</value>", f"<value>{price}</value>", new, count=1)
        avail = "true" if sellable else "false"
        if avail == "false" and m.group(3) == "true":
            to_false += 1
        if new != body or avail != m.group(3):
            changed += 1
        out.append(feed[pos:m.start()] + f'<offer id="{m.group(1)}"{m.group(2) or ""} available="{avail}">{new}</offer>')
        pos = m.end()
    out.append(feed[pos:])
    res = "".join(out)
    unavailable = len(re.findall(r'available="false"', res))
    if unavailable > MAX_DROP * len(blocks):
        sys.exit(f"abort: {unavailable} of {len(blocks)} offers would be unavailable")
    if changed:
        now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M")
        res = re.sub(r'<yml_catalog date="[^"]*">', f'<yml_catalog date="{now}">', res)
        open(FEED, "w", encoding="utf-8").write(res)
    print(f"offers={len(blocks)} changed={changed} newly_unavailable={to_false} unavailable_total={unavailable} not_in_csv={len(missing)}")

main()
