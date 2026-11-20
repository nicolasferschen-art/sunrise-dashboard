#!/usr/bin/env python3
"""
Fondspreise Updater
Liest aktuelle NAV-Preise von der Wienerbörse und lädt sie bei Parqet hoch.
Fehlende Werktage werden automatisch nachgeholt.
"""

import base64
import json
import os
import re
import subprocess
import sys
import tempfile
from datetime import date, timedelta
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

# ── Konfiguration ─────────────────────────────────────────────────────────────

ISINS = {
    "AT0000A1QA38": {"notation": "185698392", "chash": "52b3aaafcb4ed4e3afa4f68e41d8e499"},
    "AT0000A1Z882":  {"notation": "209255837", "chash": "4ad7bfd6e094679d47634d8211b9d5dd"},
    "AT0000A3EAW0":  {"notation": "479337582", "chash": "62ec12c80252f4cfd0b2732c23c394b7"},
}

PARQET_USER     = os.environ["PARQET_USER"]
PARQET_PASS     = os.environ["PARQET_PASS"]
PARQET_QUOTES   = "https://quotes-worker.parqet.com/vendors/quotes"
WIENERBORSE_URL = "https://www.wienerborse.at/marktdaten/fondsdaten-der-oekb/preisdaten/"

# Für GitHub-Issue-Benachrichtigungen (optional, gesetzt via Actions-Env)
GITHUB_TOKEN    = os.environ.get("GITHUB_TOKEN", "")
GITHUB_REPO     = os.environ.get("GITHUB_REPOSITORY", "")  # z.B. "user/repo"

# ── Hilfsfunktionen ───────────────────────────────────────────────────────────

def is_weekday(d: date) -> bool:
    return d.weekday() < 5

def missing_weekdays(last: date, today: date) -> list:
    days, d = [], last + timedelta(days=1)
    while d <= today:
        if is_weekday(d):
            days.append(d)
        d += timedelta(days=1)
    return days

def parqet_last_date(isin: str):
    """Letztes Kursdatum bei Parqet – None, wenn die Abfrage scheitert (z.B. HTTP 403)."""
    token = base64.b64encode(f"{PARQET_USER}:{PARQET_PASS}".encode()).decode()
    req = Request(f"{PARQET_QUOTES}/{isin}", headers={"Authorization": f"Basic {token}"})
    try:
        with urlopen(req, timeout=15) as resp:
            data = json.loads(resp.read())
    except HTTPError as e:
        body = ""
        try:
            body = e.read().decode("utf-8", "replace")[:300]
        except Exception:
            pass
        print(f"  ⚠️  Parqet-Abfrage {isin}: HTTP {e.code} {body}")
        return None
    except URLError as e:
        print(f"  ⚠️  Parqet nicht erreichbar ({isin}): {e}")
        return None
    dates = [q["date"] for q in data.get("quotes", [])]
    return date.fromisoformat(max(dates)) if dates else None

_WB_CACHE: dict = {}

def wienerborse_latest(isin: str):
    """Neuester NAV laut Wiener Börse als (Datum, Rücknahmepreis) – oder None.

    Das Datum kommt aus dem NAV-Block („Datum, Zeit … TT.MM.JJJJ“), nicht aus
    irgendeiner Stelle im HTML (die Fußzeile enthält immer das heutige Datum).
    """
    if isin in _WB_CACHE:
        return _WB_CACHE[isin]
    cfg = ISINS[isin]
    url = f"{WIENERBORSE_URL}?ISIN={isin}&ID_NOTATION={cfg['notation']}&cHash={cfg['chash']}"
    headers = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36",
        "Accept": "text/html,application/xhtml+xml",
        "Accept-Language": "de-AT,de;q=0.9",
    }
    result = None
    try:
        req = Request(url, headers=headers)
        with urlopen(req, timeout=15) as resp:
            raw = resp.read().decode("utf-8", errors="replace")
    except URLError as e:
        print(f"  ⚠️  Wienerbörse nicht erreichbar: {e}")
        _WB_CACHE[isin] = None
        return None

    plain = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", raw))
    m_date = re.search(r"Datum, Zeit\s+[\d.,]+\s+\S+\s*%\s+\S+\s+(\d{2})\.(\d{2})\.(\d{4})", plain)
    m_price = re.search(r"Rücknahmepreis\s+([\d.]*\d,\d+)", plain)
    if not m_date or not m_price:
        print(f"  ⚠️  NAV-Datum oder Rücknahmepreis auf Wienerbörse nicht gefunden")
    else:
        d, mth, y = (int(x) for x in m_date.groups())
        price = float(m_price.group(1).replace(".", "").replace(",", "."))
        result = (date(y, mth, d), price)
    _WB_CACHE[isin] = result
    return result

def fetch_nav_wienerborse(isin: str, target_date: date):
    latest = wienerborse_latest(isin)
    if latest is None:
        return None
    nav_date, price = latest
    if nav_date != target_date:
        print(f"  ⏭️  Wienerbörse hat Kurs vom {nav_date.strftime('%d.%m.%Y')}, nicht vom {target_date.strftime('%d.%m.%Y')}")
        return None
    return price

def upload_to_parqet(rows: list) -> int:
    csv_content = "Isin,Date,Price\n" + "".join(
        f"{r['isin']},{r['date']},{r['price']}\n" for r in rows
    )
    print(f"\nCSV:\n{csv_content}")
    with tempfile.NamedTemporaryFile(mode="w", suffix=".csv", delete=False) as f:
        f.write(csv_content)
        tmp = f.name
    body_file = tmp + ".resp"
    result = subprocess.run(
        ["curl", "-s", "-o", body_file, "-w", "%{http_code}",
         "-X", "POST", PARQET_QUOTES,
         "-u", f"{PARQET_USER}:{PARQET_PASS}",
         "-F", f"file=@{tmp}"],
        capture_output=True, text=True
    )
    os.unlink(tmp)
    code = int(result.stdout.strip() or 0)
    try:
        resp_body = open(body_file, encoding="utf-8", errors="replace").read()[:300]
        os.unlink(body_file)
    except OSError:
        resp_body = ""
    if code != 204 and resp_body:
        print(f"  Antwort von Parqet: {resp_body}")
    return code

def create_github_issue(title: str, body: str):
    """Erstellt ein GitHub Issue als Benachrichtigung (sendet automatisch E-Mail)."""
    if not GITHUB_TOKEN or not GITHUB_REPO:
        return
    payload = json.dumps({"title": title, "body": body, "labels": ["data-missing"]}).encode()
    req = Request(
        f"https://api.github.com/repos/{GITHUB_REPO}/issues",
        data=payload,
        headers={
            "Authorization": f"Bearer {GITHUB_TOKEN}",
            "Accept": "application/vnd.github+json",
            "Content-Type": "application/json",
        },
        method="POST"
    )
    try:
        with urlopen(req, timeout=10) as resp:
            issue = json.loads(resp.read())
            print(f"  📧  Issue erstellt: {issue.get('html_url')}")
    except Exception as e:
        print(f"  ⚠️  Issue konnte nicht erstellt werden: {e}")

# ── Hauptlogik ────────────────────────────────────────────────────────────────

def main():
    today = date.today()
    print(f"🗓  {today}  ({today.strftime('%A')})\n")

    if not is_weekday(today):
        print("⏭️  Wochenende – nichts zu tun.")
        return

    # Schritt 1: Fehlende Werktage ermitteln
    print("── Schritt 1: Fehlende Tage prüfen ──")
    missing: dict = {}
    for isin in ISINS:
        last = parqet_last_date(isin)
        if last is None:
            # Abfrage bei Parqet gescheitert → mindestens den neuesten Börsenkurs hochladen
            latest = wienerborse_latest(isin)
            days = [latest[0]] if latest and latest[0] <= today else []
            missing[isin] = days
            print(f"  {isin}: Parqet-Stand unbekannt → lade neuesten Kurs hoch: {days}")
            continue
        days = missing_weekdays(last, today)
        missing[isin] = days
        print(f"  {isin}: letzter Eintrag {last} → {len(days)} fehlende(r) Tag(e)")

    all_missing = sorted(set(d for days in missing.values() for d in days))
    if not all_missing:
        print("\n✅  Alles aktuell.")
        return

    # Schritt 2: NAV holen
    print("\n── Schritt 2: NAV von Wienerbörse holen ──")
    rows = []
    results = {}
    skipped_days = []

    for target_date in all_missing:
        print(f"\n📅  {target_date}:")
        day_uploaded = False
        for isin in ISINS:
            if target_date not in missing[isin]:
                continue
            print(f"  {isin}:", end=" ")
            price = fetch_nav_wienerborse(isin, target_date)
            if price is not None:
                print(f"✅  {price} EUR")
                rows.append({"isin": isin, "date": target_date.isoformat(), "price": price})
                results[isin] = {"status": "ok", "price": price, "date": target_date, "source": "Wienerbörse"}
                day_uploaded = True
            else:
                results[isin] = {"status": "skipped", "date": target_date}
        if not day_uploaded:
            skipped_days.append(target_date)

    # Schritt 3: Upload
    upload_code = None
    if rows:
        print("\n── Schritt 3: Upload bei Parqet ──")
        upload_code = upload_to_parqet(rows)
        if upload_code == 204:
            print(f"✅  Upload erfolgreich (HTTP 204)")
        else:
            print(f"❌  Upload fehlgeschlagen (HTTP {upload_code})")
    else:
        print("\n⏭️  Keine Kurse verfügbar – Upload übersprungen.")

    # Schritt 4: Benachrichtigung bei fehlenden Daten
    if skipped_days:
        today_skipped = today in skipped_days
        past_skipped = [d for d in skipped_days if d < today]

        if today_skipped:
            title = f"⏭️ Fondspreise {today} nicht verfügbar – werden morgen nachgeholt"
            body = (
                f"Die NAV-Preise für **{today}** waren um 17:00 Uhr noch nicht auf der Wienerbörse verfügbar.\n\n"
                f"➡️ Sie werden beim nächsten automatischen Run morgen nachgeholt.\n\n"
                + (f"Bereits nachgeholte Tage heute: {', '.join(str(d) for d in [r['date'] for r in rows])}\n" if rows else "")
            )
        else:
            title = f"⚠️ Fondspreise für {', '.join(str(d) for d in past_skipped)} fehlen noch"
            body = (
                f"Für folgende Werktage konnten keine NAV-Preise gefunden werden:\n\n"
                + "\n".join(f"- {d}" for d in past_skipped)
                + "\n\nBitte manuell prüfen."
            )

        print(f"\n📧  Sende Benachrichtigung: {title}")
        create_github_issue(title, body)

    # Zusammenfassung
    print("\n══════════════════════════════════════")
    any_ok = any(r.get("status") == "ok" for r in results.values())
    if upload_code == 204:
        print("Gesamtstatus: ✅ Erfolgreich")
    elif not rows:
        print("Gesamtstatus: ⏭️  Kein aktueller Kurs verfügbar – morgen Nachholen")
    else:
        print("Gesamtstatus: ⚠️  Teilweise erfolgreich")

    uploaded = [r for r in results.values() if r.get("status") == "ok"]
    print(f"Nachgeholte Tage: {len(uploaded)}")
    for isin, r in results.items():
        if r.get("status") == "ok":
            print(f"  ✅  {isin}: {r['date']} → {r['price']} EUR ({r['source']})")
        else:
            print(f"  ⏭️  {isin}: kein Kurs für {r.get('date')}")

    if upload_code == 204:
        print("Upload: ✅ HTTP 204")
    elif upload_code:
        print(f"Upload: ❌ HTTP {upload_code}")
        sys.exit(1)
    else:
        print("Upload: ⏭️  Übersprungen")

if __name__ == "__main__":
    main()
