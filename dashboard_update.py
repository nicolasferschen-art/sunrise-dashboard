#!/usr/bin/env python3
"""
IQAM Dashboard – tägliche Aktualisierung
Läuft in GitHub Actions, authentifiziert via Microsoft Graph (Refresh Token).
Liest INVENTARBLATT + INVENTARLISTE aus den letzten Mails,
generiert dashboard_data.json + IQAM_Dashboard.html,
committed die Dateien in den docs/ Ordner.
"""

import json
import os
import sys
import tempfile
import base64
import re
import copy
from datetime import datetime, date
from urllib.request import Request, urlopen
from urllib.error import HTTPError
from urllib.parse import urlencode
import traceback

try:
    import requests as _requests
    _HAS_REQUESTS = True
except ImportError:
    _HAS_REQUESTS = False

# ─── Konfiguration ────────────────────────────────────────────────────────────
SENDER_EMAIL = "rbi-fondsreporting@rbinternational.com"
FUNDS = [
    {"id": "3411", "isin": "AT0000A1QA38", "color": "#F97316",
     "name": "Standortfonds AT"},
    {"id": "3431", "isin": "AT0000A1Z882", "color": "#3B7DD8",
     "name": "Standortfonds DE"},
    {"id": "3581", "isin": "AT0000A3EAW0", "color": "#16A34A",
     "name": "Dividends and Interest"},
]

# ─── Microsoft Graph: Token holen ─────────────────────────────────────────────
def get_access_token():
    client_id     = os.environ["MS_CLIENT_ID"]
    tenant_id     = os.environ["MS_TENANT_ID"]
    refresh_token = os.environ["MS_REFRESH_TOKEN"]

    data = urlencode({
        "grant_type":    "refresh_token",
        "client_id":     client_id,
        "refresh_token": refresh_token,
        "scope":         "https://graph.microsoft.com/Mail.Read offline_access",
    }).encode()

    req = Request(
        f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
        data=data,
        method="POST",
    )
    with urlopen(req) as resp:
        result = json.loads(resp.read())

    if "access_token" not in result:
        print(f"❌ Token-Fehler: {result.get('error_description', result)}")
        sys.exit(1)

    print("✅ Access token erhalten")
    return result["access_token"]


# ─── Graph API Helper ─────────────────────────────────────────────────────────
def graph_get(access_token, path):
    url = f"https://graph.microsoft.com/v1.0{path}"
    headers = {"Authorization": f"Bearer {access_token}", "Accept": "application/json"}
    if _HAS_REQUESTS:
        resp = _requests.get(url, headers=headers)
        resp.raise_for_status()
        return resp.json()
    else:
        from urllib.parse import quote
        safe_url = quote(url, safe="/:?&=.$,@'_-+%")
        req = Request(safe_url, headers=headers)
        with urlopen(req) as r:
            return json.loads(r.read())


def graph_get_bytes(access_token, path):
    url = f"https://graph.microsoft.com/v1.0{path}"
    headers = {"Authorization": f"Bearer {access_token}"}
    if _HAS_REQUESTS:
        resp = _requests.get(url, headers=headers)
        resp.raise_for_status()
        return resp.content
    else:
        from urllib.parse import quote
        safe_url = quote(url, safe="/:?&=.$,@'_-+%")
        req = Request(safe_url, headers=headers)
        with urlopen(req) as r:
            return r.read()


# ─── Mail + Attachment suchen ─────────────────────────────────────────────────
def graph_get_url(access_token, url):
    """Holt eine vollständige URL (für Pagination mit @odata.nextLink)."""
    headers = {"Authorization": f"Bearer {access_token}", "Accept": "application/json"}
    if _HAS_REQUESTS:
        resp = _requests.get(url, headers=headers)
        resp.raise_for_status()
        return resp.json()
    else:
        from urllib.parse import quote
        safe_url = quote(url, safe="/:?&=.$,@'_-+%")
        req = Request(safe_url, headers=headers)
        with urlopen(req) as r:
            return json.loads(r.read())


def find_all_historical_emails(access_token):
    """Holt ALLE historischen Mails vom Fonds-Sender (mit Pagination, manueller Filterung)."""
    url = (
        "https://graph.microsoft.com/v1.0/me/messages"
        "?$top=100"
        "&$orderby=receivedDateTime desc"
        "&$select=id,subject,receivedDateTime,hasAttachments,from"
    )
    all_messages = []
    page = 0
    while url:
        page += 1
        print(f"  📧 Seite {page}: {len(all_messages)} Sender-Mails bisher…")
        try:
            data = graph_get_url(access_token, url)
        except Exception as e:
            print(f"  ⚠️  Fehler bei Seite {page}: {e}")
            break
        msgs = data.get("value", [])
        # Manuell nach Sender filtern
        for msg in msgs:
            sender = msg.get("from", {}).get("emailAddress", {}).get("address", "").lower()
            if sender == SENDER_EMAIL.lower():
                all_messages.append(msg)
        url = data.get("@odata.nextLink")
    print(f"  ✅ {len(all_messages)} Mails vom Fonds-Sender gefunden")
    return all_messages


def backfill_nav_history_from_emails(access_token, existing_nav_history):
    """Liest alle historischen INVENTARBLATT-Mails und befüllt nav_history mit täglichen Preisen."""
    messages = find_all_historical_emails(access_token)
    nav_history = {k: list(v) for k, v in existing_nav_history.items()}  # shallow copy

    # Gruppiere nach Fund + Eingangsdatum
    mail_map = {}  # {(fid, recv_date): msg_id}
    for msg in messages:
        subj_up = msg.get("subject", "").upper()
        if "INVENTARBLATT" not in subj_up:
            continue
        recv_date = msg.get("receivedDateTime", "")[:10]
        for fund in FUNDS:
            fid = fund["id"]
            if fid not in msg.get("subject", ""):
                continue
            key = (fid, recv_date)
            if key not in mail_map:
                mail_map[key] = msg["id"]

    total = len(mail_map)
    print(f"\n📋 {total} historische INVENTARBLATT-Mails gefunden")

    # Alle Emails chronologisch verarbeiten — keine Skip-Logik.
    # Die finale Deduplizierung (letzter Eintrag pro Datum gewinnt) sorgt dafür,
    # dass spätere Emails (mit den offiziellen Closing-Werten) frühere überschreiben.
    # Beispiel: 01.07-Email enthält den offiziellen 30.06-Closing (164.16M) und
    # überschreibt den vorläufigen Wert der 30.06-Email (163.86M).
    done = 0
    new_entries = []  # alle neu geparseten Einträge (inkl. Überschreibungen)

    for (fid, recv_date), msg_id in sorted(mail_map.items()):
        done += 1
        print(f"  [{done}/{total}] Lade {fid} {recv_date}…")
        try:
            xlsx_bytes, filename = download_attachment(access_token, msg_id, ".xlsx")
            if not xlsx_bytes:
                print(f"    ⚠️  Kein Anhang")
                continue

            blatt_data = parse_excel(xlsx_bytes, fid)
            price       = blatt_data.get("nav_per_share")
            nav         = blatt_data.get("nav")
            shares      = blatt_data.get("shares")
            perf_ytd    = blatt_data.get("perf_ytd")
            asset_date  = blatt_data.get("asset_date")   # Datum aus "Asset (by DD.MM.YYYY)"-Zeile
            report_date = blatt_data.get("report_date", recv_date)

            if not price:
                print(f"    ⚠️  Kein Preis/NAV gefunden")
                continue

            # Alle (Datum, NAV)-Paare aus dem Excel — ein Email kann mehrere Stichtage enthalten
            # (z.B. 01.07-Email: "Net Asset Value by 01.07.2026" UND "by 30.06.2026")
            nav_data_points = blatt_data.get("nav_data_points", [])

            if nav_data_points:
                # Für jedes Datum-NAV-Paar einen separaten nav_history-Eintrag anlegen.
                # perf_ytd auf alle Punkte schreiben — wird später bei Duplikaten gemergt.
                pt_ytd = round(float(perf_ytd), 4) if perf_ytd is not None else None
                for point in nav_data_points:
                    pt_nav  = point["nav"]
                    pt_date = point["date"]
                    print(f"    ✅ NAV-Paar: {pt_date} = {pt_nav/1e6:.2f} Mio. € | YTD={pt_ytd}")
                    new_entries.append((fid, {
                        "date": pt_date,
                        "price": round(float(price), 4),
                        "nav": round(float(pt_nav), 2),
                        "perf_ytd": pt_ytd,
                        "source": "measured",
                    }))
            else:
                # Fallback: einzelner Eintrag mit asset_date oder report_date
                if nav is None and price and shares:
                    nav = float(price) * float(shares)
                    print(f"    ℹ️  Nettoverm. berechnet: {price:.4f} × {shares:,.0f} = {nav/1e6:.2f} Mio. €")
                entry_date = asset_date or report_date or recv_date
                print(f"    ✅ {entry_date} | Preis={price:.4f} | NAV={nav/1e6:.2f}M" if nav else f"    ✅ {entry_date} | Preis={price:.4f}")
                new_entries.append((fid, {
                    "date": entry_date,
                    "price": round(float(price), 4),
                    "nav": round(float(nav), 2) if nav else None,
                    "perf_ytd": round(float(perf_ytd), 4) if perf_ytd is not None else None,
                    "source": "measured",
                }))
        except Exception as e:
            print(f"    ❌ Fehler: {e}")
            traceback.print_exc()

    # Bestehende + neue Einträge zusammenführen, deduplizieren.
    # Chronologische Sortierung der new_entries stellt sicher, dass spätere Emails
    # (mit offiziellen Closing-Werten) frühere vorläufige Werte überschreiben.
    for fid, entry in new_entries:
        if fid not in nav_history:
            nav_history[fid] = []
        nav_history[fid].append(entry)

    for fid in nav_history:
        by_date = {}
        for e in sorted(nav_history[fid], key=lambda x: x["date"]):
            existing = by_date.get(e["date"])
            if existing:
                # Merge: neue Werte überschreiben, aber nie einen vorhandenen Wert mit null ersetzen
                merged = {**existing, **e}
                for k in ("nav", "price", "perf_ytd"):
                    if merged.get(k) is None and existing.get(k) is not None:
                        merged[k] = existing[k]
                by_date[e["date"]] = merged
            else:
                by_date[e["date"]] = e
        nav_history[fid] = sorted(by_date.values(), key=lambda x: x["date"])

    total_entries = sum(len(v) for v in nav_history.values())
    print(f"\n✅ NAV-History: {total_entries} Einträge über alle Fonds")
    return nav_history


def backfill_holdings_history(access_token, existing_history):
    """Liest alle historischen INVENTARLISTE-Mails und befüllt holdings_history."""
    messages = find_all_historical_emails(access_token)
    holdings_history = {k: dict(v) for k, v in existing_history.items()}  # deep copy

    # Gruppiere Mails nach Fund + Datum
    mail_map = {}  # {(fid, date_str): msg_id}
    for msg in messages:
        subj   = msg.get("subject", "")
        subj_up = subj.upper()
        if "INVENTARLISTE" not in subj_up:
            continue
        recv_date = msg.get("receivedDateTime", "")[:10]  # YYYY-MM-DD
        for fund in FUNDS:
            fid = fund["id"]
            if fid not in subj:
                continue
            key = (fid, recv_date)
            if key not in mail_map:
                mail_map[key] = msg["id"]

    total = len(mail_map)
    print(f"\n📋 {total} historische INVENTARLISTE-Mails gefunden")

    done = 0
    for (fid, recv_date), msg_id in sorted(mail_map.items(), key=lambda x: x[0][1]):
        # Überspringe bereits vorhandene Snapshots
        if fid in holdings_history and recv_date in holdings_history[fid]:
            done += 1
            print(f"  ♻️  {fid} {recv_date} bereits vorhanden, überspringe")
            continue

        done += 1
        print(f"  [{done}/{total}] Lade {fid} {recv_date}…")
        try:
            xlsx_bytes, filename = download_attachment(access_token, msg_id, ".xlsx")
            if not xlsx_bytes:
                print(f"    ⚠️  Kein Anhang")
                continue
            parsed = parse_excel(xlsx_bytes, fid)
            holdings = parsed.get("holdings", [])
            if not holdings:
                print(f"    ⚠️  Keine Holdings geparst")
                continue
            snap = [
                {"isin": h.get("isin",""), "name": h.get("name",""),
                 "qty": h.get("qty"), "mv_eur": h.get("mv_eur"), "weight": h.get("weight")}
                for h in holdings if h.get("isin") and h["isin"] not in ("None","")
            ]
            if fid not in holdings_history:
                holdings_history[fid] = {}
            holdings_history[fid][recv_date] = snap
            print(f"    ✅ {len(snap)} Positionen gespeichert")
        except Exception as e:
            print(f"    ❌ Fehler: {e}")

    return holdings_history


def find_latest_emails(access_token):
    """Holt die neuesten Mails je Fond — INVENTARBLATT (NAV) + INVENTARLISTE (Holdings)."""
    path = (
        "/me/messages"
        "?$top=100"
        "&$orderby=receivedDateTime desc"
        "&$select=id,subject,receivedDateTime,hasAttachments,from"
    )
    data = graph_get(access_token, path)
    messages = data.get("value", [])
    print(f"📧 {len(messages)} Mails geladen, filtere nach {SENDER_EMAIL}")

    # Pro Fund beide Mail-Typen separat merken
    fund_mails = {}  # {fid: {"blatt": msg, "liste": msg}}
    for msg in messages:
        sender = msg.get("from", {}).get("emailAddress", {}).get("address", "").lower()
        subj   = msg.get("subject", "")
        if sender != SENDER_EMAIL.lower():
            continue
        subj_up = subj.upper()
        for fund in FUNDS:
            fid = fund["id"]
            if fid not in subj:
                continue
            if fid not in fund_mails:
                fund_mails[fid] = {"blatt": None, "liste": None}
            if "INVENTARBLATT" in subj_up and fund_mails[fid]["blatt"] is None:
                fund_mails[fid]["blatt"] = msg
                print(f"  📋 BLATT {fid}: {subj[:55]} ({msg['receivedDateTime'][:10]})")
            elif "INVENTARLISTE" in subj_up and fund_mails[fid]["liste"] is None:
                fund_mails[fid]["liste"] = msg
                print(f"  📊 LISTE {fid}: {subj[:55]} ({msg['receivedDateTime'][:10]})")

    print(f"  Gefunden: { {k: {t: bool(v) for t,v in d.items()} for k,d in fund_mails.items()} }")
    return fund_mails


def download_attachment(access_token, message_id, filename_contains):
    """Lädt den ersten Attachment herunter, dessen Name den String enthält."""
    path = f"/me/messages/{message_id}/attachments"
    data = graph_get(access_token, path)
    attachments = data.get("value", [])

    for att in attachments:
        name = att.get("name", "")
        if filename_contains.lower() in name.lower() or name.lower().endswith(".xlsx"):
            att_id = att["id"]
            print(f"  📎 Lade Anhang: {name}")
            # Binärer Download
            content = graph_get_bytes(access_token, f"/me/messages/{message_id}/attachments/{att_id}/$value")
            return content, name

    # Fallback: contentBytes aus der Metadaten-Antwort
    for att in attachments:
        if att.get("contentBytes"):
            name = att.get("name", "attachment.xlsx")
            print(f"  📎 Lade Anhang (base64): {name}")
            return base64.b64decode(att["contentBytes"]), name

    print(f"  ⚠️  Kein .xlsx Anhang gefunden in Nachricht {message_id}")
    return None, None


# ─── Excel Parser ─────────────────────────────────────────────────────────────
def parse_excel(xlsx_bytes, fund_id):
    """Parst INVENTARBLATT + INVENTARLISTE aus den xlsx-Bytes."""
    try:
        import openpyxl
    except ImportError:
        print("pip install openpyxl")
        sys.exit(1)

    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp.write(xlsx_bytes)
    tmp.close()

    wb = openpyxl.load_workbook(tmp.name, data_only=True)
    os.unlink(tmp.name)

    sheet_names = wb.sheetnames
    print(f"  📊 Sheets: {sheet_names}")

    result = {}

    # ── INVENTARBLATT ──────────────────────────────────────────────────────
    blatt = None
    for name in sheet_names:
        if "INVENTARBLATT" in name.upper() or "BLATT" in name.upper():
            blatt = wb[name]
            break
    if blatt is None and sheet_names:
        # Versuche erstes Sheet
        blatt = wb[sheet_names[0]]

    if blatt:
        result.update(_parse_inventarblatt(blatt))

    # ── INVENTARLISTE ──────────────────────────────────────────────────────
    for sheet_name in sheet_names:
        sn = sheet_name.upper()
        if "EQUIT" in sn or "AKTIE" in sn:
            result["holdings"] = _parse_positions(wb[sheet_name])
        elif "ACCOUNT" in sn or "KONTO" in sn:
            accs = _parse_positions(wb[sheet_name])
            result.setdefault("holdings", []).extend(accs)
        elif "COUNTRY" in sn or "LAND" in sn:
            result["countries"] = _parse_allocation(wb[sheet_name])
        elif "CURRENCY" in sn or "WÄHR" in sn:
            result["currencies"] = _parse_allocation(wb[sheet_name])
        elif "SEC. TYPE" in sn or "SEKTOR" in sn or "SECTOR" in sn:
            result["sectors"] = _parse_allocation(wb[sheet_name])

    return result


def _parse_inventarblatt(ws):
    """Liest NAV, Preis/Anteil, BVI-Performance aus dem Inventarblatt."""
    data = {}
    rows = list(ws.iter_rows(values_only=True))

    # Debug: alle Zeilen mit Inhalt (max 80)
    print(f"    [BLATT] Sheet '{ws.title}':")
    for i, row in enumerate(rows[:80]):
        non_empty = [(j, str(c)) for j, c in enumerate(row) if c is not None]
        if non_empty:
            print(f"      Row {i+1}: {non_empty}")

    nav_data_points = []  # alle (Datum, NAV)-Paare aus "Net Asset Value by..."-Zeilen

    def _extract_date_from_row(row, row_str):
        """Extrahiert Datum aus einer Zeile — unterstützt DD.MM.YYYY, YYYY-MM-DD und datetime-Objekte."""
        # 1. DD.MM.YYYY in row_str (Zelle als Text)
        m = re.search(r'(\d{1,2})\.(\d{1,2})\.(\d{4})', row_str)
        if m:
            day, mon, yr = m.groups()
            return f"{yr}-{mon.zfill(2)}-{day.zfill(2)}"
        # 2. YYYY-MM-DD in row_str (openpyxl datetime als String)
        m2 = re.search(r'(\d{4})-(\d{2})-(\d{2})', row_str)
        if m2:
            return m2.group(0)[:10]
        # 3. datetime-Objekt direkt in einer Zelle
        for cell in row:
            if isinstance(cell, (datetime, date)) and not isinstance(cell, bool):
                return str(cell)[:10]
        return None

    for i, row in enumerate(rows):
        row_str = " ".join(str(c) for c in row if c is not None).upper()

        # Gesamtvermögen / NAV — viele mögliche Labels.
        # Enthält die Zeile ein Datum (z.B. "Net Asset Value by 30.06.2026"),
        # wird das Datum-NAV-Paar separat gesammelt (für korrekte Monatsend-Zuordnung).
        if any(k in row_str for k in ["GESAMTVERM", "FONDSVERM", "NETTOVERM", "TOTAL NET ASSET",
                                       "TOTAL ASSETS", "FUND VOLUME", "INVENTARWERT GESAMT",
                                       "NET ASSET VALUE", "GESAMT"]):
            nav_val = None
            for cell in row:
                if isinstance(cell, (int, float)) and not isinstance(cell, bool) and cell > 1_000_000:
                    nav_val = float(cell)
                    break
            if nav_val:
                row_date = _extract_date_from_row(row, row_str)
                if row_date:
                    nav_data_points.append({"date": row_date, "nav": nav_val})
                    print(f"    → NAV-Paar Zeile {i+1}: {row_date} = {nav_val/1e6:.2f} Mio. €")
                else:
                    data["nav"] = nav_val  # kein Datum → als generischer NAV speichern
                    print(f"    → NAV Zeile {i+1}: {nav_val/1e6:.2f} Mio. € (kein Datum)")

        # Rücknahmepreis / NAV per share
        # Direkter Match auf "Asset (by DD.MM.YYYY)" — Spalte 2 = Anteile, Spalte 7 = Redemption price
        if "ASSET (BY" in row_str or "ASSET(BY" in row_str:
            # Bewertungsdatum: erste Fundstelle = Vortags-Closing-Datum
            if not data.get("asset_date"):
                d = _extract_date_from_row(row, row_str)
                if d:
                    data["asset_date"] = d
                    print(f"    → Bewertungsdatum (Asset-Zeile {i+1}): {d}")
            # Spalte 2 = Issued/Anteile
            if len(row) > 2 and row[2] is not None:
                try:
                    v = float(row[2])
                    if v > 1000:
                        data["shares"] = v
                        print(f"    → Anteile (Asset-Zeile {i+1}): {v:,.0f}")
                except (TypeError, ValueError):
                    pass
            # Spalte 7 = Redemption price, Spalte 5 = Unit Price
            for col_idx in [7, 6, 5]:
                if col_idx < len(row) and row[col_idx] is not None:
                    try:
                        v = float(row[col_idx])
                        if 10 < v < 5000:
                            data["nav_per_share"] = v
                            print(f"    → Preis (Asset-Zeile {i+1}, col {col_idx}): {v}")
                            break
                    except (TypeError, ValueError):
                        pass
        elif any(k in row_str for k in ["RÜCKNAHME", "ANTEILSWERT", "INVENTARWERT JE",
                                       "REDEMPTION PRICE", "NET ASSET VALUE PER",
                                       "PRICE PER UNIT", "VALUE PER SHARE", "UNIT VALUE",
                                       "ANTEILSPR", "RECHENWERT", "FONDSKURS", "KURS JE",
                                       "PREIS JE", "AUSGABE", "PRICE PER"]):
            for cell in row:
                try:
                    v = float(cell)
                    if 1 < v < 100_000:
                        data["nav_per_share"] = v
                        print(f"    → Preis gefunden Zeile {i+1}: {v}")
                        break
                except (TypeError, ValueError):
                    pass

        # Anzahl Anteile
        if any(k in row_str for k in ["ANTEILE", "UNITS", "SHARES OUTSTANDING", "AUSGEGEBEN"]):
            if any(k in row_str for k in ["UMLAUF", "AUSST", "OUTSTANDING", "ISSUED", "GESAMT"]):
                for cell in row:
                    if isinstance(cell, (int, float)) and cell > 100:
                        data["shares"] = float(cell)
                        print(f"    → Anteile gefunden Zeile {i+1}: {cell}")
                        break

        # BVI Performance
        if "BVI" in row_str or "PERFORMANCE" in row_str or "RENDITE" in row_str:
            if any(k in row_str for k in ["01.01", "JAHRESBEG", "YTD", "YEAR TO DATE", "SEIT 01.01"]):
                for cell in row:
                    if isinstance(cell, (int, float)) and -50 < cell < 200:
                        data["perf_ytd"] = float(cell)
                        print(f"    → YTD gefunden Zeile {i+1}: {cell}")
                        break
            if any(k in row_str for k in ["01.10", "GESCHÄFTSJ", "GJ", "FISCAL", "FISKAL", "SEIT 01.10"]):
                for cell in row:
                    if isinstance(cell, (int, float)) and -50 < cell < 200:
                        data["perf_fy"] = float(cell)
                        print(f"    → FY gefunden Zeile {i+1}: {cell}")
                        break

    # NAV-Daten-Paare auswerten:
    # Falls mehrere (Datum, NAV)-Paare gefunden (z.B. "01.07" und "30.06"),
    # alle speichern. Der neueste Eintrag wird als Primär-NAV verwendet (für tagesaktuelle Anzeige).
    # Das ältere Datum (= Monatsultimo) wird als asset_date gesetzt falls noch nicht vorhanden.
    if nav_data_points:
        sorted_points = sorted(nav_data_points, key=lambda x: x["date"])
        data["nav_data_points"] = sorted_points
        # Primär-NAV = neuester Wert (für tagesaktuelle Verwendung im Full-Run)
        data["nav"] = sorted_points[-1]["nav"]
        # asset_date = ältester Wert (= Vortags-Closing, für korrekten Backfill-Eintrag)
        if not data.get("asset_date"):
            data["asset_date"] = sorted_points[0]["date"]
        print(f"    → NAV-Paare gesamt: {[(p['date'], round(p['nav']/1e6,2)) for p in sorted_points]}")
        print(f"    → asset_date={data.get('asset_date')}, NAV(primär)={data['nav']/1e6:.2f} Mio. €")

    # Datum aus Zellen (für report_date)
    for row in rows[:10]:
        for cell in row:
            if isinstance(cell, (datetime, date)):
                data["report_date"] = str(cell)[:10]
                break

    # Fallback Rücknahmepreis: Suche Zeilen 30–65 nach plausibler Zahl (10–5000 EUR)
    # Plausibilitätsprüfung: NAV / Preis muss > 1000 Anteile ergeben
    if not data.get("nav_per_share") and data.get("nav"):
        nav_val = data["nav"]
        for i, row in enumerate(rows[29:65], start=30):
            for cell in row:
                try:
                    cell_f = float(cell)
                except (TypeError, ValueError):
                    continue
                if 10 < cell_f < 5000:
                    implied_shares = nav_val / cell_f
                    if implied_shares > 1000:
                        data["nav_per_share"] = cell_f
                        print(f"    → Preis (Fallback Zeile {i}): {cell_f} → {implied_shares:,.0f} Anteile impl.")
                        break
            if data.get("nav_per_share"):
                break

    print(f"    [BLATT] Ergebnis: {data}")
    return data


def _parse_positions(ws):
    """Liest Holdings-Positionen."""
    rows = list(ws.iter_rows(values_only=True))
    holdings = []

    # Header-Zeile finden
    header_idx = None
    for i, row in enumerate(rows):
        row_str = " ".join(str(c) for c in row if c is not None).upper()
        if ("ISIN" in row_str or "WKN" in row_str) and ("NAME" in row_str or "BEZEICH" in row_str):
            header_idx = i
            break

    if header_idx is None:
        # Fallback: ab Zeile 3 lesen
        header_idx = 2

    headers = [str(c).strip() if c else "" for c in rows[header_idx]]
    print(f"    [LISTE] Header-Zeile {header_idx+1}: {[(j,h) for j,h in enumerate(headers) if h]}")

    # Spalten-Indices bestimmen
    col_map = {}
    for j, h in enumerate(headers):
        hu = h.upper()
        if "ISIN" in hu:                       col_map["isin"] = j
        elif "NAME" in hu or "BEZEICH" in hu:  col_map["name"] = j
        elif "COUNTRY" in hu or "LAND" in hu:  col_map["country"] = j
        elif "SECTOR" in hu or "BRANCHE" in hu or "SEC.TYPE" in hu: col_map["sector"] = j
        elif "CURRENCY" in hu or "WÄHRUNG" in hu: col_map["currency"] = j
        # Marktwert in Fondswährung (EUR)
        elif "MKT VAL" in hu and ("EUR" in hu or "FNDCCY" in hu or "FUND" in hu):
            col_map["mv_eur"] = j
        elif any(k in hu for k in [
                "P&L", "P/L", "G&V", "G/V", "GEWINN", "UNREALIZED", "UNREAL",
                "GAIN/LOSS", "GAIN LOSS", "BOOK PROFIT", "BUCHGEWINN", "BUCHGEWINN/-VERLUST",
                "BUCHGEWINN/-V", "PROFIT", "VERLUST", "BWG", "BW-GEWINN",
                "UNREALISED", "UNREALIZ", "GAIN", "ACCRU"]):
            col_map.setdefault("pl", j)
        elif "WEIGHT" in hu or "ANTEIL" in hu or "%" in hu:
            col_map.setdefault("weight", j)
        elif "COST" in hu or "EINSTAND" in hu or "KAUFPREIS" in hu or "BOOK VALUE" in hu or "BUCHWERT" in hu:
            col_map.setdefault("cost", j)
        elif "PRICE" in hu or "KURS" in hu:
            col_map.setdefault("price", j)
        elif any(k in hu for k in ["QUANTITY", "STÜCK", "STUCKZAHL", "NOMINAL", "QTY",
                                    "ANZAHL", "BESTAND", "ANTEILE", "NENN", "UNITS",
                                    "SHARES", "VOLUME", "VOLUMEN", "AMOUNT"]):
            col_map.setdefault("qty", j)

    print(f"    [LISTE] col_map: {col_map}")

    # Fallback für Marktwert: Spalte 15 (aus Analyse der echten Daten)
    if "mv_eur" not in col_map:
        col_map["mv_eur"] = 15
        print(f"    [LISTE] mv_eur Fallback: Spalte 15")

    # Fallback für P&L: Spalte nach mv_eur suchen die pos+neg Werte hat
    if "pl" not in col_map and "mv_eur" in col_map:
        mv_idx = col_map["mv_eur"]
        data_rows = [r for r in rows[header_idx + 1:] if any(r)][:30]
        for try_idx in range(mv_idx + 1, min(mv_idx + 8, len(headers))):
            vals = []
            for r in data_rows:
                if try_idx < len(r) and r[try_idx] is not None:
                    try:
                        v = float(r[try_idx])
                        vals.append(v)
                    except (TypeError, ValueError):
                        pass
            has_pos = any(v > 0 for v in vals)
            has_neg = any(v < 0 for v in vals)
            if has_pos and has_neg and len(vals) >= 3:
                col_map["pl"] = try_idx
                print(f"    [LISTE] P&L Fallback: Spalte {try_idx} (pos+neg Werte)")
                break

    for row in rows[header_idx + 1:]:
        if not any(row):
            continue
        # Name-Spalte muss gefüllt sein
        name_val = row[col_map.get("name", 1)] if len(row) > col_map.get("name", 1) else None
        if not name_val or str(name_val).strip() in ("", "None", "Total", "Gesamt"):
            continue

        def get_col(key, default=None):
            idx = col_map.get(key)
            if idx is None or idx >= len(row):
                return default
            v = row[idx]
            if v is None:
                return default
            return v

        mv_raw = get_col("mv_eur")
        try:
            mv = float(mv_raw) if mv_raw is not None else 0.0
        except (TypeError, ValueError):
            mv = 0.0

        pl_raw = get_col("pl")
        try:
            pl = float(pl_raw) if pl_raw is not None else None
        except (TypeError, ValueError):
            pl = None

        w_raw = get_col("weight")
        try:
            w = float(w_raw) if w_raw is not None else None
        except (TypeError, ValueError):
            w = None

        cost_raw = get_col("cost")
        try:
            cost = float(cost_raw) if cost_raw is not None else None
        except (TypeError, ValueError):
            cost = None

        price_raw = get_col("price")
        try:
            price = float(price_raw) if price_raw is not None else None
        except (TypeError, ValueError):
            price = None

        h = {
            "isin":     str(get_col("isin", "")),
            "name":     str(name_val).strip(),
            "country":  str(get_col("country", "Unbekannt")).strip(),
            "sector":   str(get_col("sector", "Sonstiges")).strip(),
            "currency": str(get_col("currency", "EUR")).strip(),
            "mv_eur":   mv,
            "pl":       pl,
            "weight":   w,
            "cost":     cost,
            "price":    price,
        }
        if mv != 0.0 or pl is not None:
            holdings.append(h)

    return holdings


def _parse_allocation(ws):
    """Liest Allokations-Tabellen (Country / Currency / Sector).
    Versucht Prozent-Spalte zu finden; fällt auf erste numerische Spalte zurück.
    """
    rows = list(ws.iter_rows(values_only=True))

    # Header-Zeile finden und Prozent-Spalte identifizieren
    pct_col = None
    data_start = 0
    for i, row in enumerate(rows[:10]):
        row_str = " ".join(str(c) for c in row if c is not None).upper()
        if any(k in row_str for k in ["%", "WEIGHT", "ANTEIL", "PERCENT", "GEWICHT"]):
            # Finde die Prozent-Spalte
            for j, cell in enumerate(row):
                if cell and any(k in str(cell).upper() for k in ["%", "WEIGHT", "ANTEIL", "PERCENT", "GEWICHT"]):
                    pct_col = j
                    break
            data_start = i + 1
            break

    print(f"    [ALLOC] Sheet '{ws.title}': pct_col={pct_col}, data_start={data_start}")

    result = []
    skip_labels = {"", "None", "Total", "Gesamt", "Land", "Country", "Sektor", "Sector",
                   "Währung", "Currency", "Name", "Bezeichnung"}
    for row in rows[data_start:]:
        if len(row) < 2:
            continue
        label = row[0]
        if not label or str(label).strip() in skip_labels:
            continue
        label_str = str(label).strip()
        # Abbruch bei Summenzeilen
        if any(k in label_str.upper() for k in ["TOTAL", "SUMME", "GESAMT", "SUM"]):
            continue

        val = None
        if pct_col is not None and pct_col < len(row):
            cell = row[pct_col]
            if isinstance(cell, (int, float)) and cell != 0:
                val = float(cell)

        if val is None:
            # Fallback: bevorzuge kleine Zahlen (0-100) als %, vermeide Millionenbeträge
            for cell in row[1:]:
                if isinstance(cell, (int, float)) and cell != 0:
                    if 0 < abs(cell) <= 100:
                        val = float(cell)
                        break
            if val is None:
                # letzter Fallback: irgendeine Zahl
                for cell in row[1:]:
                    if isinstance(cell, (int, float)) and cell != 0:
                        val = float(cell)
                        break

        if val is not None:
            result.append({"label": label_str, "value": val})

    print(f"    [ALLOC] {len(result)} Einträge: {result[:5]}")
    return result


# ─── Kennzahlen berechnen ─────────────────────────────────────────────────────
def compute_kpis(fund_data):
    holdings = fund_data.get("holdings", [])
    nav = fund_data.get("nav", 0)

    total_mv   = sum(h["mv_eur"] for h in holdings)
    total_pl   = sum(h["pl"] for h in holdings if h["pl"] is not None)
    pos_pl     = sum(h["pl"] for h in holdings if h["pl"] and h["pl"] > 0)
    neg_pl     = sum(h["pl"] for h in holdings if h["pl"] and h["pl"] < 0)
    equities_mv = total_mv

    # HHI (Herfindahl-Hirschman Index) – basierend auf Gewichtung
    weights = []
    for h in holdings:
        w = h.get("weight")
        if w is None and nav and nav > 0:
            w = h["mv_eur"] / nav * 100
        if w:
            weights.append(w / 100)
    hhi = sum(w**2 for w in weights) * 10000 if weights else 0

    # Top 10 nach Gewicht
    sorted_h = sorted(holdings, key=lambda x: x["mv_eur"], reverse=True)
    top10_mv  = sum(h["mv_eur"] for h in sorted_h[:10])
    top10_pct = top10_mv / nav * 100 if nav else 0

    # Gewinn-/Verlust-Positionen
    win_positions  = len([h for h in holdings if h.get("pl") and h["pl"] > 0])
    total_with_pl  = len([h for h in holdings if h.get("pl") is not None])
    win_rate = win_positions / total_with_pl * 100 if total_with_pl else 0

    fund_data.update({
        "total_pl":   total_pl,
        "pos_pl":     pos_pl,
        "neg_pl":     neg_pl,
        "equities_mv": equities_mv,
        "hhi":        round(hhi, 1),
        "top10_weight": round(top10_pct, 1),
        "win_rate":   round(win_rate, 1),
    })

    # Länder- und Sektor-Allokation IMMER aus Holdings berechnen (zuverlässiger als dedizierte Sheets)
    if holdings:
        total_mv_h = sum(h["mv_eur"] for h in holdings if h["mv_eur"])
        if total_mv_h > 0:
            ctry_mv = {}
            for h in holdings:
                c = h.get("country") or "Unbekannt"
                if c not in ("None", "", "Unbekannt"):
                    ctry_mv[c] = ctry_mv.get(c, 0) + (h["mv_eur"] or 0)
            if ctry_mv:
                fund_data["countries"] = [
                    {"label": k, "value": round(v / total_mv_h * 100, 2)}
                    for k, v in sorted(ctry_mv.items(), key=lambda x: x[1], reverse=True)
                    if v > 0
                ]
                print(f"    [KPI] Länder aus Holdings: {len(fund_data['countries'])} Einträge")

            sec_mv = {}
            for h in holdings:
                s = h.get("sector") or "Sonstiges"
                if s not in ("None", ""):
                    sec_mv[s] = sec_mv.get(s, 0) + (h["mv_eur"] or 0)
            if sec_mv:
                fund_data["sectors"] = [
                    {"label": k, "value": round(v / total_mv_h * 100, 2)}
                    for k, v in sorted(sec_mv.items(), key=lambda x: x[1], reverse=True)
                    if v > 0
                ]
                print(f"    [KPI] Sektoren aus Holdings: {len(fund_data['sectors'])} Einträge")

    return fund_data


# ─── Änderungen erkennen ──────────────────────────────────────────────────────
def detect_changes(current_holdings, prev_holdings):
    """Vergleicht aktuelle vs. gestrige Holdings."""
    if not prev_holdings:
        return {"added": [], "removed": [], "increased": [], "decreased": [], "date_prev": None}

    curr_map = {h["isin"]: h for h in current_holdings if h.get("isin") and h["isin"] != "None"}
    prev_map = {h["isin"]: h for h in prev_holdings  if h.get("isin") and h["isin"] != "None"}

    added   = [curr_map[i] for i in curr_map if i not in prev_map]
    removed = [prev_map[i] for i in prev_map if i not in curr_map]

    increased = []
    decreased = []
    for isin in curr_map:
        if isin in prev_map:
            # Stückzahl vergleichen (Kauf/Verkauf durch Fondsmanager)
            c_qty = curr_map[isin].get("qty") or 0
            p_qty = prev_map[isin].get("qty") or 0
            if c_qty and p_qty and abs(c_qty - p_qty) / max(abs(p_qty), 1) > 0.005:
                diff_pct = (c_qty - p_qty) / abs(p_qty) * 100
                entry = {**curr_map[isin], "change_pct": round(diff_pct, 2),
                         "prev_qty": p_qty, "curr_qty": c_qty}
                if c_qty > p_qty:
                    increased.append(entry)
                else:
                    decreased.append(entry)

    increased.sort(key=lambda x: abs(x["change_pct"]), reverse=True)
    decreased.sort(key=lambda x: abs(x["change_pct"]), reverse=True)

    return {
        "added":     added[:20],
        "removed":   removed[:20],
        "increased": increased[:10],
        "decreased": decreased[:10],
        "date_prev": None,  # wird vom Aufrufer gesetzt
    }


# ─── Preishistorie (BVI back-calculation) ────────────────────────────────────
def build_price_history(nav_per_share, perf_ytd, perf_fy, nav_per_share_prev=None):
    today = date.today()
    hist = []

    # GJ-Start (01.10. des Vorjahres)
    fy_start_year = today.year - 1 if today.month < 10 else today.year
    fy_start = date(fy_start_year, 10, 1)
    if perf_fy is not None and nav_per_share:
        fy_price = nav_per_share / (1 + perf_fy / 100)
        hist.append({
            "date":  fy_start.isoformat(),
            "label": f"01.10.{fy_start_year}",
            "price": round(fy_price, 4),
            "note":  "GJ-Start (berechnet aus BVI)",
        })

    # Jahresstart (01.01.)
    ytd_start = date(today.year, 1, 1)
    if perf_ytd is not None and nav_per_share:
        ytd_price = nav_per_share / (1 + perf_ytd / 100)
        hist.append({
            "date":  ytd_start.isoformat(),
            "label": f"01.01.{today.year}",
            "price": round(ytd_price, 4),
            "note":  "Jahresstart (berechnet aus BVI)",
        })

    # Vortag (aus prev_data)
    if nav_per_share_prev is not None:
        from datetime import timedelta
        prev_date = today - timedelta(days=1)
        # Wochenenden überspringen
        while prev_date.weekday() >= 5:
            prev_date -= timedelta(days=1)
        hist.append({
            "date":  prev_date.isoformat(),
            "label": prev_date.strftime("%d.%m.%Y"),
            "price": round(nav_per_share_prev, 4),
            "note":  "Vortag (Inventarblatt)",
        })

    # Heute
    if nav_per_share:
        hist.append({
            "date":  today.isoformat(),
            "label": today.strftime("%d.%m.%Y"),
            "price": round(nav_per_share, 4),
            "note":  "Heute (Inventarblatt)",
        })

    return hist


# ─── News fetchen ─────────────────────────────────────────────────────────────
def summarize_news(company_name, articles, anthropic_key):
    """Fasst News-Schlagzeilen via Claude Haiku zu Überschrift + Fließtext zusammen."""
    if not anthropic_key or not articles:
        return None
    articles_text = ""
    for a in articles[:5]:  # max 5 Artikel pro Unternehmen
        title = (a.get("title") or "")[:120]
        desc  = (a.get("desc") or "")[:150]
        content = f"{title} — {desc}" if desc and title not in desc else title
        src = f" ({a['source']})" if a.get("source") else ""
        articles_text += f"- {content}{src}\n"

    prompt = (
        f"Du fasst Nachrichtenartikel für Investoren zusammen.\n"
        f"Unternehmen: {company_name}\n\n"
        f"Artikel:\n{articles_text}\n\n"
        f"STRENGE REGELN – lies sie sorgfältig:\n"
        f"1. Verwende AUSSCHLIESSLICH Informationen die WÖRTLICH im obigen Text stehen.\n"
        f"2. VERBOTEN ohne expliziten Beleg im Text:\n"
        f"   - Indexänderungen (ATX, DAX, MSCI etc. Aufnahme oder Ausschluss)\n"
        f"   - Übernahmen, Fusionen, M&A\n"
        f"   - CEO- oder Managementwechsel\n"
        f"   - Konkrete Zahlen (Umsatz, Gewinn, Kurs) die nicht im Text stehen\n"
        f"   - Zukunftsprognosen die nicht zitiert werden\n"
        f"3. Im Zweifel: weglassen oder IRRELEVANT antworten.\n"
        f"4. Wenn kein Artikel wirklich zu {company_name} passt (Geschäft, Zahlen, Strategie): antworte nur mit IRRELEVANT\n\n"
        f"Antworte in diesem Format:\n"
        f"HEADLINE: [präzise Überschrift, nur was belegt ist]\n"
        f"TEXT: [2–3 Sätze, direkt, nur belegte Fakten – lieber kürzer als spekulativ]"
    )
    body = json.dumps({
        "model": "claude-haiku-4-5-20251001",
        "max_tokens": 200,
        "messages": [{"role": "user", "content": prompt}],
    }, ensure_ascii=False).encode("utf-8")
    req = Request(
        "https://api.anthropic.com/v1/messages", data=body, method="POST",
        headers={
            "x-api-key": anthropic_key,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json; charset=utf-8",
        },
    )
    try:
        with urlopen(req, timeout=20) as resp:
            result = json.loads(resp.read())
            raw = result["content"][0]["text"].strip()
        if raw.startswith("IRRELEVANT"):
            return None
        headline, text = "", ""
        for line in raw.splitlines():
            if line.startswith("HEADLINE:"):
                headline = line[9:].strip()
            elif line.startswith("TEXT:"):
                text = line[5:].strip()
        return {"headline": headline, "text": text} if (headline or text) else None
    except HTTPError as e:
        body_err = ""
        try: body_err = e.read().decode("utf-8", errors="replace")[:300]
        except Exception: pass
        print(f"    ⚠️  Haiku-Fehler für {company_name}: HTTP {e.code} – {body_err}")
        return None
    except Exception as e:
        print(f"    ⚠️  Haiku-Fehler für {company_name}: {e}")
        return None


def _clean_news_name(name):
    """Bereinigt Firmennamen für Google News Suche."""
    clean = re.sub(
        r'\s+(PLC|AG|SE|INC\.?|CORP\.?|LTD\.?|S\.A\.|SA|NV|BV|SPA|CO\.?|GRP|GROUP|'
        r'HOLDINGS?|INH\.?|O\.N\.|DL-?[\d,\.]+|EO[\s\-]?[\d,\.]+|LS-?[\d,\.]+|'
        r'SF\s*[\d,\.]+|CL\.[A-Z]|BNK)(\s.*)?$',
        '', name.strip(), flags=re.IGNORECASE
    ).strip()
    return ' '.join(clean.split()[:3])


    # Domains die keine Finanz-News liefern → werden herausgefiltert
_BLOCKLIST_DOMAINS = {
    "xboxdynasty", "filmstarts", "photografix", "connect.de", "computerbild",
    "chip.de", "heise.de", "golem.de", "computerbase", "ign.com", "gamestar",
    "gamesradar", "eurogamer", "kotaku", "polygon.com", "pcgamer", "hardwareluxx",
    "notebookcheck", "techradar", "theverge", "engadget", "wired.com",
    "solidbau", "immobilien", "realestate", "architekt", "bau.de",
}

def _is_finance_relevant(title, source):
    """Prüft ob ein Artikel finanziell relevant ist."""
    src_low = source.lower()
    if any(bl in src_low for bl in _BLOCKLIST_DOMAINS):
        return False
    # Titel-Filter: mindestens ein Finanz-Keyword oder kein offensichtliches Off-Topic
    off_topic = ["rezept", "urlaub", "reise", "gaming", "spiel ", "film ", "serie ",
                 "musik", "mode ", "beauty", "gesundheit", "sport ", "fußball",
                 "küche", "wohnen", "garten"]
    title_low = title.lower()
    if any(kw in title_low for kw in off_topic):
        return False
    return True


def fetch_all_news(companies, max_per_company=8, request_timeout=5, anthropic_key=None, max_summaries=80, prev_news_data=None, max_wall_seconds=300):
    """Fetcht Finanz-News via Google News RSS für alle Unternehmen, optional mit Haiku-Zusammenfassung."""
    import xml.etree.ElementTree as ET
    import time as _time
    from urllib.parse import quote as _quote

    news_data = {}
    items_list = list(companies.items())
    total = len(items_list)
    summarize = bool(anthropic_key)
    summary_count = 0
    deadline = _time.monotonic() + max_wall_seconds
    print(f"\n📰 Fetche News für {total} Unternehmen{f' (KI-Summary für Top {max_summaries})' if summarize else ''} (max {max_wall_seconds}s)…")

    for i, (key, co) in enumerate(items_list):
        if _time.monotonic() > deadline:
            print(f"  ⏱️  Zeitlimit erreicht nach {i} Unternehmen – stoppe News-Fetch.")
            break
        clean = _clean_news_name(co["name"])
        if not clean or len(clean) < 3:
            continue
        q = _quote(clean[:50])
        url = f"https://news.google.com/rss/search?q={q}&hl=de&gl=AT&ceid=AT:de"
        try:
            req = Request(url, headers={
                "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36"
            })
            with urlopen(req, timeout=request_timeout) as resp:
                content = resp.read()
            root = ET.fromstring(content)
            arts = []
            for el in root.findall(".//item"):
                if len(arts) >= max_per_company:
                    break
                title = re.sub(r"<[^>]+>|<!\[CDATA\[|\]\]>", "", el.findtext("title") or "").strip()
                link  = (el.findtext("link") or "#").strip()
                pub   = (el.findtext("pubDate") or "").strip()
                src   = getattr(el.find("source"), "text", "") or ""
                desc  = re.sub(r"<[^>]+>|<!\[CDATA\[|\]\]>", "", el.findtext("description") or "").strip()
                if title and _is_finance_relevant(title, src):
                    arts.append({"title": title, "link": link, "pubDate": pub, "source": src, "desc": desc})
            if arts:
                # Article-Caching: Summary wiederverwenden wenn Top-Artikel unverändert
                prev = (prev_news_data or {}).get(key, {})
                prev_top = ((prev.get("articles") or [{}])[0]).get("title", "")
                curr_top = arts[0].get("title", "")
                if prev_top and curr_top == prev_top and prev.get("summary"):
                    summary = prev["summary"]
                    print(f"    ♻️  Cache für {co['name'][:35]}")
                else:
                    do_summary = summarize and summary_count < max_summaries
                    summary = None
                    if do_summary:
                        _time.sleep(1.5)  # Rate-limit
                        summary = summarize_news(co["name"], arts, anthropic_key)
                    if summary:
                        summary_count += 1
                news_data[key] = {
                    "company": co["name"],
                    "funds": co["funds"],
                    "articles": arts,
                    "summary": summary,
                }
        except Exception as _ne:
            if i < 3 or (i + 1) % 20 == 0:
                print(f"  ⚠️  News-Fehler {co['name'][:30]}: {type(_ne).__name__}: {_ne}")
        if (i + 1) % 20 == 0:
            print(f"  {i+1}/{total}…")
        _time.sleep(0.2)

    print(f"  ✅ {len(news_data)}/{total} Unternehmen mit Artikeln")
    return news_data


# ─── Dashboard HTML generieren ────────────────────────────────────────────────
def generate_html(funds_data, updated_at, nav_history=None, news_data=None, run_log=None, changes_history=None):
    """Generiert das vollständige Dashboard-HTML."""
    data_json = json.dumps(funds_data, ensure_ascii=False, separators=(',', ':'))
    nav_history_json = json.dumps(nav_history or {}, ensure_ascii=False, separators=(',', ':'))

    # Zahlenformatierung
    def fmt_eur(n, dec=2):
        if n is None:
            return "—"
        try:
            n = float(n)
        except (TypeError, ValueError):
            return "—"
        if abs(n) >= 1_000_000_000:
            return f"{n/1_000_000_000:+.2f} Mrd."
        if abs(n) >= 1_000_000:
            return f"{n/1_000_000:.2f} Mio."
        return f"{n:,.2f}"

    total_aum = sum(f.get("nav", 0) for f in funds_data)
    total_pl  = sum(f.get("total_pl", 0) for f in funds_data)
    avg_ytd   = sum(f.get("perf_ytd", 0) for f in funds_data) / len(funds_data) if funds_data else 0

    def pl_sign(v):
        return "+" if v > 0 else ""

    html = f"""<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Sunrise.app Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/canvas-confetti@1.9.2/dist/confetti.browser.min.js"></script>
<style>
:root {{
  --bg: #FAFAF8;
  --surface: #FFFFFF;
  --surface2: #FFF6EE;
  --border: #EDE8E1;
  --text: #1C1917;
  --muted: #78716C;
  --accent: #F97316;
  --accent-light: #FFF0E6;
  --blue: #3B7DD8;
  --green: #16A34A;
  --red: #DC2626;
  --orange: #EA580C;
}}
body.dark {{
  --bg: #1C1917;
  --surface: #292524;
  --surface2: #3B2A18;
  --border: #44403C;
  --text: #F5F5F4;
  --muted: #A8A29E;
  --accent-light: #431407;
}}
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{ background: var(--bg); color: var(--text); font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; font-size: 14px; }}
a {{ color: var(--accent); text-decoration: none; }}

/* Header */
.sticky-header {{ position: sticky; top: 0; z-index: 100; background: var(--surface); border-bottom: 1px solid var(--border); padding: 14px 32px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 1px 4px rgba(0,0,0,0.06); }}
.sticky-header h1 {{ font-size: 17px; font-weight: 700; color: var(--text); }}
.header-kpis {{ display: flex; gap: 36px; }}
.hkpi {{ text-align: right; }}
.hkpi .val {{ font-size: 17px; font-weight: 700; }}
.hkpi .lbl {{ font-size: 11px; color: var(--muted); margin-top: 1px; }}

/* Tabs */
.tabs {{ display: flex; gap: 2px; padding: 0 32px; border-bottom: 1px solid var(--border); overflow-x: auto; background: var(--surface); }}
.tab {{ padding: 12px 18px; cursor: pointer; font-size: 13px; font-weight: 500; color: var(--muted); border-bottom: 2px solid transparent; white-space: nowrap; transition: color .15s; }}
.tab.active {{ color: var(--accent); border-bottom-color: var(--accent); }}
.tab:hover:not(.active) {{ color: var(--text); }}

/* Panels */
.panel {{ display: none; padding: 28px 32px; max-width: 1400px; margin: 0 auto; }}
.panel.active {{ display: block; }}
.grid-2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }}
.grid-3 {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; }}
.grid-4 {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; }}

/* Cards */
.card {{ background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 20px; }}
.card h3 {{ font-size: 11px; font-weight: 600; color: var(--muted); text-transform: uppercase; letter-spacing: .6px; margin-bottom: 10px; }}
.kpi-val {{ font-size: 28px; font-weight: 700; letter-spacing: -.5px; }}
.kpi-sub {{ font-size: 12px; color: var(--muted); margin-top: 5px; }}

/* Colors */
.pos {{ color: var(--green); }}
.neg {{ color: var(--red); }}

/* Section titles */
.section-title {{ font-size: 15px; font-weight: 600; margin: 28px 0 14px; color: var(--text); }}

/* Charts */
.chart-wrap {{ position: relative; height: 280px; }}
.chart-wrap.tall {{ height: 360px; }}

/* Changes */
.changes-grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; }}
.change-card {{ background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 16px; }}
.change-card h4 {{ font-size: 11px; font-weight: 600; color: var(--muted); text-transform: uppercase; letter-spacing: .5px; margin-bottom: 10px; }}
.change-item {{ padding: 7px 0; border-bottom: 1px solid var(--border); font-size: 12px; }}
.change-item:last-child {{ border-bottom: none; }}
.change-item .ci-name {{ font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 100%; color: var(--text); }}
.change-item .ci-sub {{ color: var(--muted); font-size: 11px; margin-top: 2px; }}
.badge-new  {{ background: #DCFCE7; color: #16A34A; padding: 2px 7px; border-radius: 4px; font-size: 10px; font-weight: 700; }}
.badge-out  {{ background: #FEE2E2; color: #DC2626; padding: 2px 7px; border-radius: 4px; font-size: 10px; font-weight: 700; }}
.badge-up   {{ background: #DCFCE7; color: #16A34A; padding: 2px 7px; border-radius: 4px; font-size: 10px; font-weight: 700; }}
.badge-down {{ background: #FFF7ED; color: #EA580C; padding: 2px 7px; border-radius: 4px; font-size: 10px; font-weight: 700; }}
.placeholder {{ padding: 24px; text-align: center; color: var(--muted); font-size: 13px; font-style: italic; }}

/* Tables */
table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
th {{ text-align: left; padding: 9px 12px; border-bottom: 1px solid var(--border); color: var(--muted); font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: .4px; cursor: pointer; user-select: none; white-space: nowrap; }}
th:hover {{ color: var(--text); }}
td {{ padding: 9px 12px; border-bottom: 1px solid var(--border); color: var(--text); }}
tr:last-child td {{ border-bottom: none; }}
tr:hover td {{ background: var(--surface2); }}
.tbl-wrap {{ overflow-x: auto; max-height: 520px; overflow-y: auto; }}
.tbl-controls {{ display: flex; gap: 10px; margin-bottom: 14px; flex-wrap: wrap; align-items: center; }}
.tbl-controls input, .tbl-controls select {{ background: var(--bg); border: 1px solid var(--border); color: var(--text); padding: 7px 12px; border-radius: 8px; font-size: 13px; outline: none; }}
.tbl-controls input:focus, .tbl-controls select:focus {{ border-color: var(--accent); }}
.tbl-controls input {{ flex: 1; min-width: 200px; }}
.pagination {{ display: flex; gap: 5px; justify-content: center; margin-top: 14px; flex-wrap: wrap; align-items: center; }}
.page-btn {{ background: var(--surface); border: 1px solid var(--border); color: var(--muted); padding: 5px 11px; border-radius: 6px; cursor: pointer; font-size: 12px; transition: all .15s; }}
.page-btn:hover {{ border-color: var(--accent); color: var(--accent); }}
.page-btn.active {{ background: var(--accent); color: white; border-color: var(--accent); font-weight: 600; }}

/* Modal */
.modal-overlay {{ display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.3); z-index: 999; align-items: center; justify-content: center; backdrop-filter: blur(2px); }}
.modal-overlay.open {{ display: flex; }}
.modal {{ background: var(--surface); border: 1px solid var(--border); border-radius: 16px; padding: 28px; max-width: 680px; width: 95%; max-height: 85vh; overflow-y: auto; box-shadow: 0 20px 60px rgba(0,0,0,0.15); }}
.modal h2 {{ font-size: 18px; font-weight: 700; margin-bottom: 16px; }}
.modal-kpis {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; margin-bottom: 16px; }}
.modal-kpi {{ background: var(--surface2); border-radius: 10px; padding: 12px; }}
.modal-kpi .lbl {{ font-size: 11px; color: var(--muted); margin-bottom: 4px; }}
.modal-kpi .val {{ font-size: 17px; font-weight: 700; }}
.modal-links {{ display: flex; gap: 8px; flex-wrap: wrap; margin-top: 14px; }}
.modal-links a {{ background: var(--bg); border: 1px solid var(--border); padding: 6px 14px; border-radius: 8px; font-size: 12px; color: var(--text); transition: border-color .15s; }}
.modal-links a:hover {{ border-color: var(--accent); color: var(--accent); }}
.modal-close {{ float: right; cursor: pointer; color: var(--muted); font-size: 22px; line-height: 1; }}
.modal-close:hover {{ color: var(--text); }}

/* Fund cards */
.fund-card-link {{ cursor: pointer; transition: transform .15s, box-shadow .15s; }}
.fund-card-link:hover {{ transform: translateY(-2px); box-shadow: 0 4px 16px rgba(249,115,22,0.12); border-color: var(--accent); }}

/* Bars */
.bar-row {{ display: flex; align-items: center; gap: 10px; margin-bottom: 7px; }}
.bar-label {{ width: 140px; font-size: 12px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; flex-shrink: 0; color: var(--text); }}
.bar-track {{ flex: 1; height: 16px; background: var(--surface2); border-radius: 6px; overflow: hidden; }}
.bar-fill {{ height: 100%; border-radius: 6px; opacity: 0.85; }}
.bar-val {{ width: 60px; text-align: right; font-size: 12px; color: var(--muted); flex-shrink: 0; }}

.range-btn {{ background: var(--bg); border: 1px solid var(--border); color: var(--muted); padding: 5px 12px; border-radius: 6px; cursor: pointer; font-size: 12px; font-weight: 500; transition: all .15s; }}
.range-btn:hover, .range-btn.active {{ background: var(--accent); color: white; border-color: var(--accent); }}
.updated {{ font-size: 11px; color: var(--muted); }}

/* Tooltips */
.tip {{ border-bottom: 1px dashed var(--border); cursor: help; position: relative; display: inline; }}
.tip::after {{ content: attr(data-tip); position: absolute; bottom: calc(100% + 8px); left: 50%; transform: translateX(-50%); background: var(--text); border-radius: 8px; padding: 8px 12px; font-size: 12px; color: white; white-space: normal; min-width: 220px; max-width: 300px; line-height: 1.5; z-index: 500; opacity: 0; pointer-events: none; transition: opacity .2s; box-shadow: 0 4px 20px rgba(0,0,0,0.15); }}
.tip:hover::after {{ opacity: 1; }}

@media (max-width: 768px) {{
  .grid-2, .grid-3, .grid-4, .changes-grid {{ grid-template-columns: 1fr; }}
  .header-kpis {{ display: none; }}
  .modal-kpis {{ grid-template-columns: 1fr 1fr; }}
  .panel {{ padding: 16px; }}
  .sticky-header, .tabs {{ padding-left: 16px; padding-right: 16px; }}
}}
</style>
</head>
<body>

<div class="sticky-header">
  <div>
    <h1>☀️ Sunrise.app Dashboard</h1>
    <div class="updated" style="display:flex;align-items:center;gap:8px">
      <span>Stand: {updated_at}</span>
      <button onclick="toggleRunLog()" title="Run-Historie anzeigen"
        style="background:none;border:1px solid var(--border);border-radius:4px;padding:1px 7px;font-size:11px;color:var(--muted);cursor:pointer;line-height:1.6"
        onmouseover="this.style.borderColor='var(--accent)';this.style.color='var(--accent)'"
        onmouseout="this.style.borderColor='var(--border)';this.style.color='var(--muted)'">Runs</button>
      <button id="dark-toggle" onclick="toggleDark()" title="Dark Mode"
        style="background:none;border:1px solid var(--border);border-radius:4px;padding:1px 7px;font-size:13px;cursor:pointer;line-height:1.6"
        onmouseover="this.style.borderColor='var(--accent)'"
        onmouseout="this.style.borderColor='var(--border)'">🌙</button>
    </div>
  </div>
  <div class="header-kpis">
    <div class="hkpi">
      <div class="val">{fmt_eur(total_aum)} €</div>
      <div class="lbl"><span class="tip" data-tip="Nettovermögen aller verwalteten Fonds zusammen (Summe der Fondsvermögen)">Gesamt-AuM</span></div>
    </div>
    <div class="hkpi">
      <div class="val {'pos' if total_pl >= 0 else 'neg'}">{pl_sign(total_pl)}{fmt_eur(total_pl)} €</div>
      <div class="lbl"><span class="tip" data-tip="Unrealisiertes P&L: Buchgewinne und -verluste aller offenen Positionen (noch nicht realisiert)">Unreal. P&amp;L</span></div>
    </div>
    <div class="hkpi">
      <div class="val {'pos' if avg_ytd >= 0 else 'neg'}">{pl_sign(avg_ytd)}{avg_ytd:.2f}%</div>
      <div class="lbl"><span class="tip" data-tip="Durchschnittliche BVI-Performance aller Fonds seit 01.01. des laufenden Jahres">Ø YTD</span></div>
    </div>
  </div>
</div>

<div class="tabs" id="tabs">
  <div class="tab active" data-tab="overview">📊 Übersicht</div>
"""

    for f in funds_data:
        html += f'  <div class="tab" data-tab="fund-{f["id"]}">{f["name"]}</div>\n'

    html += '  <div class="tab" data-tab="news">📰 News</div>\n'
    html += '  <div class="tab" data-tab="calc">🧮 Rechner</div>\n'
    html += '</div>\n\n'

    # ── Overview Panel ──────────────────────────────────────────────────────
    html += '<div class="panel active" id="panel-overview">\n'
    html += '<div class="grid-3">\n'

    for f in funds_data:
        ytd = f.get("perf_ytd", 0) or 0
        nav_ps = f.get("nav_per_share") or 0
        nav_ps_prev = f.get("nav_per_share_prev")
        day_chg = ""
        if nav_ps and nav_ps_prev and abs((nav_ps - nav_ps_prev) / nav_ps_prev) < 0.15:
            d = (nav_ps - nav_ps_prev) / nav_ps_prev * 100
            d_cls = "pos" if d >= 0 else "neg"
            day_chg = f'<div class="kpi-sub {d_cls}">{pl_sign(d)}{d:.2f}% heute</div>'

        html += f"""<div class="card fund-card-link" onclick="switchTab('fund-{f['id']}')">
  <h3>{f['name']}</h3>
  <div class="kpi-val">{fmt_eur(f.get('nav', 0))} €</div>
  <div class="kpi-sub" style="color:var(--muted)"><span class="tip" data-tip="Nettoinventarwert pro Anteilschein (Rücknahmepreis)">Rücknahmepreis</span>: <strong>{nav_ps:.4f} €</strong></div>
  <div class="kpi-sub {'pos' if ytd>=0 else 'neg'}"><span class="tip" data-tip="BVI-Performance seit 01.01. des laufenden Jahres">YTD</span>: {pl_sign(ytd)}{ytd:.2f}%</div>
  {day_chg}
  <div class="kpi-sub" style="color:var(--muted)">{len(f.get('holdings', []))} Positionen</div>
</div>
"""

    html += '</div>\n'

    # Performance-Vergleich Chart
    html += '<div class="section-title">Performance-Vergleich</div>\n'
    html += '<div class="grid-2">\n'
    html += '<div class="card"><h3>YTD Performance</h3><div class="chart-wrap"><canvas id="cov-perf"></canvas></div></div>\n'
    html += '<div class="card"><h3>Fondsvermögen (AuM)</h3><div class="chart-wrap"><canvas id="cov-aum"></canvas></div></div>\n'
    html += '</div>\n'

    # Monatsreporting
    html += '<div class="section-title">Monatsreporting</div>\n'
    html += '<div class="card" style="overflow-x:auto">\n'
    html += '<div id="monthly-table-wrap"><div style="color:var(--muted);font-size:13px;padding:8px">Lade Monatsdaten…</div></div>\n'
    html += '</div>\n'

    # Alle Positionen aller Fonds
    html += '<div class="section-title">Alle Positionen</div>\n'
    html += '<div class="card">\n'
    html += _build_all_holdings_table(funds_data)
    html += '</div>\n'
    html += '</div>\n\n'  # end overview panel

    # ── Per-Fund Panels ──────────────────────────────────────────────────────
    for f in funds_data:
        fid = f["id"]
        holdings = f.get("holdings", [])
        countries = f.get("countries", [])
        currencies = f.get("currencies", [])
        sectors = f.get("sectors", [])
        changes = f.get("changes", {})
        ph = f.get("price_history", [])
        nav = f.get("nav", 0)
        nav_ps = float(f.get("nav_per_share") or 0)
        nav_ps_prev = f.get("nav_per_share_prev")
        day_chg_pct = ((nav_ps - nav_ps_prev) / nav_ps_prev * 100) if nav_ps and nav_ps_prev and abs((nav_ps - nav_ps_prev) / nav_ps_prev) < 0.15 else None
        ytd = f.get("perf_ytd", 0) or 0
        fy  = f.get("perf_fy", 0) or 0
        total_pl_f  = f.get("total_pl", 0) or 0
        hhi   = f.get("hhi", 0) or 0
        win_r = f.get("win_rate", 0) or 0
        top10 = f.get("top10_weight", 0) or 0

        html += f'<div class="panel" id="panel-fund-{fid}">\n'

        # KPI Cards
        html += '<div class="grid-4">\n'
        html += f'''<div class="card">
  <h3><span class="tip" data-tip="Nettovermögen des Fonds (Marktwert aller Assets minus Verbindlichkeiten)">Fondsvermögen (NAV)</span></h3>
  <div class="kpi-val">{fmt_eur(nav)} €</div>
  <div class="kpi-sub">{f.get("shares", 0):,.0f} Anteile</div>
</div>
<div class="card">
  <h3><span class="tip" data-tip="Nettoinventarwert pro Anteilschein – der Preis zu dem Anteile zurückgegeben werden können">Rücknahmepreis</span></h3>
  <div class="kpi-val">{nav_ps:.4f} €</div>
  {f'<div class="kpi-sub {chr(39)}pos{chr(39) if (day_chg_pct or 0)>=0 else chr(39)}neg{chr(39)}">{pl_sign(day_chg_pct or 0)}{(day_chg_pct or 0):.2f}% heute</div>' if day_chg_pct is not None else ''}
</div>
<div class="card">
  <h3><span class="tip" data-tip="BVI-Performance seit 01.01. des laufenden Jahres (Jahresbeginn bis heute)">YTD Performance</span></h3>
  <div class="kpi-val {'pos' if ytd>=0 else 'neg'}">{pl_sign(ytd)}{ytd:.2f}%</div>
  <div class="kpi-sub">GJ: <span class="{'pos' if fy>=0 else 'neg'}">{pl_sign(fy)}{fy:.2f}%</span></div>
</div>
<div class="card">
  <h3><span class="tip" data-tip="Unrealisiertes Gewinn/Verlust: Buchgewinne und -verluste aller offenen Positionen (noch nicht durch Verkauf realisiert)">Unrealisiertes P&amp;L</span></h3>
  <div class="kpi-val {'pos' if total_pl_f>=0 else 'neg'}">{pl_sign(total_pl_f)}{fmt_eur(total_pl_f)} €</div>
  <div class="kpi-sub pos">+{fmt_eur(f.get("pos_pl",0))} € / <span class="neg">{fmt_eur(f.get("neg_pl",0))} €</span></div>
</div>
'''
        html += '</div>\n'

        # NAV Chart with date range
        html += '<div class="section-title">NAV Entwicklung</div>\n'
        today_iso = date.today().isoformat()
        six_months_ago = date(date.today().year - (1 if date.today().month <= 6 else 0),
                              ((date.today().month - 6 - 1) % 12) + 1, 1).isoformat() if True else ""
        html += f'''<div class="card">
  <div style="display:flex;gap:12px;align-items:center;flex-wrap:wrap;margin-bottom:16px">
    <div style="display:flex;gap:8px;align-items:center">
      <label style="font-size:12px;color:var(--muted);font-weight:600">Von</label>
      <input type="date" id="from-{fid}" value="2025-10-01" style="background:var(--bg);border:1px solid var(--border);color:var(--text);padding:5px 10px;border-radius:8px;font-size:13px">
    </div>
    <div style="display:flex;gap:8px;align-items:center">
      <label style="font-size:12px;color:var(--muted);font-weight:600">Bis</label>
      <input type="date" id="to-{fid}" value="{today_iso}" style="background:var(--bg);border:1px solid var(--border);color:var(--text);padding:5px 10px;border-radius:8px;font-size:13px">
    </div>
    <div style="display:flex;gap:6px;flex-wrap:wrap">
      <button onclick="setDateRange('{fid}','gy')" class="range-btn" id="btn-gy-{fid}">GJ</button>
      <button onclick="setDateRange('{fid}','ytd')" class="range-btn" id="btn-ytd-{fid}">YTD</button>
      <button onclick="setDateRange('{fid}','3m')" class="range-btn">3M</button>
      <button onclick="setDateRange('{fid}','1m')" class="range-btn">1M</button>
      <button onclick="setDateRange('{fid}','all')" class="range-btn">Gesamt</button>
    </div>
  </div>
  <div class="chart-wrap" style="height:260px"><canvas id="chart-spark-{fid}"></canvas></div>
</div>
'''

        # Änderungen
        html += '<div class="section-title">Positionsänderungen (ggü. Vortag)</div>\n'
        html += '<div class="changes-grid">\n'
        for badge_type, badge_class, badge_label, items_key in [
            ("Neu aufgenommen", "badge-new", "NEU", "added"),
            ("Verkauft", "badge-out", "RAUS", "removed"),
            ("Aufgestockt", "badge-up", "↑", "increased"),
            ("Reduziert", "badge-down", "↓", "decreased"),
        ]:
            items = changes.get(items_key, [])
            html += f'<div class="change-card"><h4>{badge_type} <span class="{badge_class}">{badge_label}</span></h4>\n'
            if items:
                for item in items[:8]:
                    sub = ""
                    if items_key in ("increased", "decreased") and item.get("change_pct"):
                        sub = f'{pl_sign(item["change_pct"])}{item["change_pct"]:.1f}% Stückzahl'
                    elif item.get("mv_eur"):
                        sub = f'{fmt_eur(item["mv_eur"])} €'
                    html += f'<div class="change-item"><div class="ci-name">{item.get("name","—")}</div><div class="ci-sub">{sub}</div></div>\n'
            else:
                html += '<div class="placeholder">Noch keine Daten für heute<br>(wird ab dem 2. Tag befüllt)</div>\n'
            html += '</div>\n'
        html += '</div>\n'

        # Transaktionshistorie (kumulativ)
        ch = (changes_history or {}).get(fid, [])
        html += '<div class="section-title">Transaktionshistorie</div>\n'
        html += '<div class="card">\n'
        if ch:
            TYPE_CONFIG = {
                "added":     ("Neukauf",         "#16a34a", "▲"),
                "removed":   ("Komplettverkauf",  "#dc2626", "✕"),
                "increased": ("Aufstockung",      "#2563eb", "↑"),
                "decreased": ("Teilverkauf",      "#ea580c", "↓"),
            }
            def fmt_mv_tx(v):
                if v is None: return "—"
                v = float(v)
                if abs(v) >= 1e6: return f"{v/1e6:.2f} Mio. €"
                return f"{fmt_eur(v)} €"
            html += '<div class="tbl-wrap"><table>\n'
            html += '<thead><tr><th>Datum</th><th>Typ</th><th>Unternehmen</th><th>ISIN</th><th style="text-align:right">Gehandelt</th><th style="text-align:right">Nach Trade</th></tr></thead>\n'
            html += '<tbody>\n'
            for entry in sorted(ch, key=lambda x: x.get("date",""), reverse=True):
                typ   = entry.get("type","")
                label, color, icon = TYPE_CONFIG.get(typ, (typ, "#6b7280", "•"))
                mv    = entry.get("mv_eur")
                chg   = entry.get("change_pct")
                badge = f'<span style="display:inline-flex;align-items:center;gap:4px;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:600;color:{color};background:{color}18">{icon} {label}</span>'
                # Gehandelter Betrag und Restposition ableiten
                if typ == "added":
                    traded, after = mv, mv
                elif typ == "removed":
                    traded, after = (-(mv or 0)), 0
                elif chg is not None and mv:
                    prev_mv = mv / (1 + chg / 100)
                    traded  = mv - prev_mv           # positiv = gekauft, negativ = verkauft
                    after   = mv
                else:
                    traded, after = None, mv
                # Formatierung Gehandelt
                if traded is not None:
                    sign_s = "+" if traded > 0 else ""
                    tc     = "#16a34a" if traded >= 0 else "#dc2626"
                    traded_html = f'<span style="color:{tc};font-weight:600">{sign_s}{fmt_mv_tx(traded)}</span>'
                else:
                    traded_html = "—"
                # Formatierung Nach Trade
                if typ == "removed":
                    after_html = '<span style="color:var(--muted)">Position geschlossen</span>'
                else:
                    after_html = fmt_mv_tx(after)
                html += (f'<tr>'
                         f'<td style="white-space:nowrap;color:var(--muted);font-size:12px">{entry.get("date","—")}</td>'
                         f'<td>{badge}</td>'
                         f'<td style="font-weight:500">{entry.get("name","—")[:38]}</td>'
                         f'<td style="font-family:monospace;font-size:11px;color:var(--muted)">{entry.get("isin","—")}</td>'
                         f'<td style="text-align:right;font-size:12px">{traded_html}</td>'
                         f'<td style="text-align:right;font-size:12px;color:var(--muted)">{after_html}</td>'
                         f'</tr>\n')
            html += '</tbody></table></div>\n'
        else:
            html += '<div class="placeholder">Noch keine Transaktionshistorie — wird ab dem 2. Handelstag befüllt.</div>\n'
        html += '</div>\n'

        # Top Gewinner / Verlierer
        sorted_by_pl = sorted([h for h in holdings if h.get("pl") is not None], key=lambda x: x["pl"], reverse=True)
        top_gainers = sorted_by_pl[:10]
        top_losers  = sorted_by_pl[-10:][::-1]

        html += '<div class="section-title">Top Gewinner / Verlierer</div>\n'
        html += '<div class="grid-2">\n'
        html += '<div class="card"><h3>🏆 Top 10 Gewinner</h3><div class="tbl-wrap"><table>\n'
        html += '<tr><th>Name</th><th>P&L EUR</th><th>% NAV</th></tr>\n'
        for h in top_gainers:
            nav_pct = h["mv_eur"] / nav * 100 if nav else 0
            html += f'<tr><td>{h["name"][:35]}</td><td class="pos">+{fmt_eur(h["pl"])} €</td><td>{nav_pct:.2f}%</td></tr>\n'
        html += '</table></div></div>\n'

        html += '<div class="card"><h3>📉 Top 10 Verlierer</h3><div class="tbl-wrap"><table>\n'
        html += '<tr><th>Name</th><th>P&L EUR</th><th>% NAV</th></tr>\n'
        for h in top_losers:
            nav_pct = h["mv_eur"] / nav * 100 if nav else 0
            html += f'<tr><td>{h["name"][:35]}</td><td class="neg">{fmt_eur(h["pl"])} €</td><td>{nav_pct:.2f}%</td></tr>\n'
        html += '</table></div></div>\n'
        html += '</div>\n'

        # Konzentrations-Panel
        html += '<div class="section-title">Portfolio-Analyse</div>\n'
        html += '<div class="grid-4">\n'
        html += f'''<div class="card">
  <h3><span class="tip" data-tip="Herfindahl-Hirschman Index: Maß für Konzentration. 0 = perfekt diversifiziert, 10.000 = ein einziger Titel. Unter 1.500 gilt als gut diversifiziert.">HHI Konzentration</span></h3>
  <div class="kpi-val">{hhi:.0f}</div>
  <div class="kpi-sub">{"⚠️ Hoch" if hhi > 1500 else "✅ Diversifiziert"}</div>
</div>
<div class="card">
  <h3><span class="tip" data-tip="Anteil der Positionen mit positivem P&L an allen Positionen mit bekanntem P&L">Win-Rate</span></h3>
  <div class="kpi-val">{win_r:.1f}%</div>
  <div class="kpi-sub">Positionen mit Gewinn</div>
</div>
<div class="card">
  <h3>Top-10 Gewicht</h3>
  <div class="kpi-val">{top10:.1f}%</div>
  <div class="kpi-sub">der 10 größten Positionen</div>
</div>
<div class="card">
  <h3><span class="tip" data-tip="Aktienquote: Anteil der Aktien am gesamten Fondsvermögen (Marktwert Aktien / NAV)">Aktienquote</span></h3>
  <div class="kpi-val">{f.get("equities_mv", 0)/nav*100:.1f}%</div>
  <div class="kpi-sub">{fmt_eur(f.get("equities_mv",0))} € MV</div>
</div>
'''
        html += '</div>\n'

        # Allokations-Charts
        html += '<div class="section-title">Allokation</div>\n'
        html += '<div class="grid-3">\n'
        # Länder
        html += f'<div class="card"><h3>Länder</h3><div id="bars-country-{fid}">\n'
        sorted_countries = sorted(countries, key=lambda x: x["value"], reverse=True)[:15]
        max_v_c = sorted_countries[0]["value"] if sorted_countries else 1
        for item in sorted_countries:
            max_v = max_v_c
            pct = item["value"] / max_v * 100
            html += f'<div class="bar-row"><div class="bar-label" title="{item["label"]}">{item["label"]}</div><div class="bar-track"><div class="bar-fill" style="width:{pct:.1f}%;background:{f["color"]}"></div></div><div class="bar-val">{item["value"]:,.1f}%</div></div>\n'
        html += '</div></div>\n'
        # Währungen
        html += f'<div class="card"><h3>Währungen</h3><div class="chart-wrap"><canvas id="chart-ccy-{fid}"></canvas></div></div>\n'
        # Sektoren
        html += f'<div class="card"><h3>Sektoren</h3><div id="bars-sector-{fid}">\n'
        sorted_sectors = sorted(sectors, key=lambda x: x["value"], reverse=True)[:10]
        max_v_s = sorted_sectors[0]["value"] if sorted_sectors else 1
        for item in sorted_sectors:
            max_v = max_v_s
            pct = item["value"] / max_v * 100
            html += f'<div class="bar-row"><div class="bar-label" title="{item["label"]}">{item["label"]}</div><div class="bar-track"><div class="bar-fill" style="width:{pct:.1f}%;background:{f["color"]}"></div></div><div class="bar-val">{item["value"]:,.1f}%</div></div>\n'
        html += '</div></div>\n'
        html += '</div>\n'

        # P&L by Country + Sector
        country_pl = {}
        sector_pl  = {}
        for h in holdings:
            if h.get("pl") is not None:
                country_pl[h.get("country","Unbekannt")] = country_pl.get(h.get("country","Unbekannt"), 0) + h["pl"]
                sector_pl[h.get("sector","Sonstiges")]   = sector_pl.get(h.get("sector","Sonstiges"), 0) + h["pl"]

        top_c = sorted(country_pl.items(), key=lambda x: x[1])[::-1][:12]
        top_s = sorted(sector_pl.items(),  key=lambda x: x[1])[::-1][:10]

        html += '<div class="grid-2">\n'
        html += f'<div class="card"><h3>P&amp;L nach Ländern</h3><div class="chart-wrap tall"><canvas id="chart-plc-{fid}"></canvas></div></div>\n'
        html += f'<div class="card"><h3>P&amp;L nach Sektoren</h3><div class="chart-wrap tall"><canvas id="chart-pls-{fid}"></canvas></div></div>\n'
        html += '</div>\n'

        # Holdings Table
        html += '<div class="section-title">Alle Positionen</div>\n'
        html += f'<div class="card">\n'
        html += f'<div class="tbl-controls">\n'
        html += f'<input type="text" id="search-{fid}" placeholder="Suche nach Name, ISIN, Land …" oninput="filterTable(\'{fid}\')">\n'
        html += f'<select id="filter-country-{fid}" onchange="filterTable(\'{fid}\')">\n'
        countries_list = sorted(set(h.get("country","") for h in holdings if h.get("country")))
        html += '<option value="">Alle Länder</option>\n'
        for c in countries_list:
            html += f'<option value="{c}">{c}</option>\n'
        html += '</select>\n'
        sectors_list = sorted(set(h.get("sector","") for h in holdings if h.get("sector")))
        html += f'<select id="filter-sector-{fid}" onchange="filterTable(\'{fid}\')">\n'
        html += '<option value="">Alle Sektoren</option>\n'
        for s in sectors_list:
            html += f'<option value="{s}">{s}</option>\n'
        html += '</select>\n'
        html += f'<button onclick="exportCSV(\'{fid}\')" style="background:var(--surface2);border:1px solid var(--border);color:var(--text);padding:6px 12px;border-radius:6px;cursor:pointer">⬇ CSV</button>\n'
        html += '</div>\n'

        # Table
        html += f'<div class="tbl-wrap"><table id="tbl-{fid}">\n'
        html += '<thead><tr>'
        cols = [("Name","name"),("ISIN","isin"),("Land","country"),("Sektor","sector"),
                ("Währung","currency"),("Marktwert €","mv_eur"),("P&L €","pl"),
                ("% NAV","nav_pct"),("Einstand","cost"),("Kurs","price")]
        for lbl, key in cols:
            tip_map = {
                "Marktwert €": ' class="tip" data-tip="Marktwert in EUR (Fondswährung)"',
                "P&L €": ' class="tip" data-tip="Unrealisiertes Gewinn/Verlust dieser Position in EUR"',
                "% NAV": ' class="tip" data-tip="Anteil dieser Position am Gesamtvermögen des Fonds"',
                "Einstand": ' class="tip" data-tip="Durchschnittlicher Einstandskurs (Kaufpreis)"',
            }
            tip = tip_map.get(lbl, "")
            html += f'<th onclick="sortTable(\'{fid}\',\'{key}\')" data-key="{key}"><span{tip}>{lbl}</span> ↕</th>'
        html += '</tr></thead>\n'
        html += f'<tbody id="tbody-{fid}">\n'

        for h in holdings:
            nav_pct = h["mv_eur"] / nav * 100 if nav and h["mv_eur"] else 0
            pl_val  = h.get("pl") or 0
            pl_cls  = "pos" if pl_val > 0 else ("neg" if pl_val < 0 else "")
            pl_sign_s = "+" if pl_val > 0 else ""
            isin_s  = h.get("isin","") or ""
            yf_url  = f"https://finance.yahoo.com/quote/{isin_s}.DE" if isin_s else "#"
            html += f'''<tr onclick="showModal(this)" data-fid="{fid}" data-isin="{isin_s}" data-name="{h['name']}" data-country="{h.get('country','')}" data-sector="{h.get('sector','')}" data-currency="{h.get('currency','')}" data-mv="{h['mv_eur']:.2f}" data-pl="{pl_val:.2f}" data-nav-pct="{nav_pct:.3f}" data-cost="{h.get('cost') or ''}" data-price="{h.get('price') or ''}">
<td>{h['name'][:40]}</td>
<td style="font-size:11px;color:var(--muted)">{isin_s}</td>
<td>{h.get('country','')}</td>
<td>{h.get('sector','')}</td>
<td>{h.get('currency','')}</td>
<td>{fmt_eur(h['mv_eur'])} €</td>
<td class="{pl_cls}">{pl_sign_s}{fmt_eur(pl_val)} €</td>
<td>{nav_pct:.2f}%</td>
<td>{fmt_eur(h.get('cost')) if h.get('cost') else '—'}</td>
<td>{fmt_eur(h.get('price')) if h.get('price') else '—'}</td>
</tr>
'''
        html += '</tbody></table></div>\n'
        html += f'<div class="pagination" id="pages-{fid}"></div>\n'
        html += '</div>\n'  # card

        html += '</div>\n\n'  # end fund panel

    # ── News Panel ───────────────────────────────────────────────────────────
    html += '<div class="panel" id="panel-news">\n'
    news_fund_btns = '<button class="range-btn active" id="newsf-all" onclick="setNewsFund(\'all\')">Alle Fonds</button>\n'
    for f in funds_data:
        news_fund_btns += f'<button class="range-btn" id="newsf-{f["id"]}" onclick="setNewsFund(\'{f["id"]}\')" style="border-color:{f["color"]}40;color:{f["color"]}">{f["name"]}</button>\n'

    html += f'''<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;flex-wrap:wrap;gap:12px">
  <div>
    <div class="section-title" style="margin:0">📰 News — Alle Positionen</div>
    <div style="font-size:12px;color:var(--muted);margin-top:4px" id="news-status">Lade…</div>
  </div>
  <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
    <input type="text" id="news-search" placeholder="🔍 Suche…"
      oninput="renderNewsPanel(_newsFundFilter)"
      style="background:var(--bg);border:1px solid var(--border);color:var(--text);padding:6px 12px;border-radius:8px;font-size:13px;outline:none;width:180px">
    <label style="display:flex;align-items:center;gap:6px;font-size:12px;color:var(--muted);cursor:pointer">
      <input type="checkbox" id="news-filter-summary" onchange="renderNewsPanel(_newsFundFilter)"
        style="accent-color:var(--accent);cursor:pointer">
      Nur ausformulierte
    </label>
    <button onclick="try{{sessionStorage.setItem('activeTab','news')}}catch(e){{}}; location.reload()"
      style="background:none;border:1px solid var(--border);border-radius:6px;padding:6px 14px;font-size:12px;color:var(--muted);cursor:pointer"
      onmouseover="this.style.borderColor='var(--accent)';this.style.color='var(--accent)'"
      onmouseout="this.style.borderColor='var(--border)';this.style.color='var(--muted)'">↻ Aktualisieren</button>
  </div>
</div>
<div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:20px">
  {news_fund_btns}
</div>
<div id="news-feed" style="display:grid;gap:10px"></div>
'''
    html += '</div>\n\n'  # end news panel

    # ── Calculator Panel ────────────────────────────────────────────────────
    html += '<div class="panel" id="panel-calc">\n'
    html += '<div class="section-title">🧮 Performance-Rechner</div>\n'
    html += f'''<div class="card" style="max-width:900px;margin-bottom:24px">
  <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:16px;align-items:end">
    <div>
      <label style="display:block;font-size:12px;color:var(--muted);font-weight:600;margin-bottom:6px">Investitionsbetrag (€)</label>
      <input type="number" id="calc-amount" value="10000" min="1" oninput="calcPerf()" style="width:100%;background:var(--bg);border:1px solid var(--border);color:var(--text);padding:8px 12px;border-radius:8px;font-size:16px;font-weight:600">
    </div>
    <div>
      <label style="display:block;font-size:12px;color:var(--muted);font-weight:600;margin-bottom:6px">Von Datum</label>
      <input type="date" id="calc-from" oninput="calcPerf()" style="width:100%;background:var(--bg);border:1px solid var(--border);color:var(--text);padding:8px 12px;border-radius:8px;font-size:14px">
    </div>
    <div>
      <label style="display:block;font-size:12px;color:var(--muted);font-weight:600;margin-bottom:6px">Bis Datum</label>
      <input type="date" id="calc-to" oninput="calcPerf()" style="width:100%;background:var(--bg);border:1px solid var(--border);color:var(--text);padding:8px 12px;border-radius:8px;font-size:14px">
    </div>
    <div>
      <label style="display:block;font-size:12px;color:var(--muted);font-weight:600;margin-bottom:6px">Schnellauswahl</label>
      <div style="display:flex;gap:6px;flex-wrap:wrap">
        <button onclick="setCalcRange('ytd')" class="range-btn" id="calc-btn-ytd">YTD</button>
        <button onclick="setCalcRange('gy')" class="range-btn" id="calc-btn-gy">GJ</button>
        <button onclick="setCalcRange('1y')" class="range-btn">1J</button>
        <button onclick="setCalcRange('all')" class="range-btn">Gesamt</button>
      </div>
    </div>
  </div>
</div>
<div class="grid-3" id="calc-results">
'''
    for f in funds_data:
        fid = f["id"]
        html += f'''<div class="card" id="calc-card-{fid}">
  <h3 style="margin-bottom:12px">{f['name']}</h3>
  <div id="calc-detail-{fid}" style="color:var(--muted);font-size:13px">→ Zeitraum wählen</div>
</div>
'''
    html += '</div>\n'
    html += '<div class="section-title" style="margin-top:32px">Normalisierter Verlauf</div>\n'
    html += '<div class="card"><div class="chart-wrap tall"><canvas id="chart-calc-perf"></canvas></div></div>\n'
    html += '</div>\n\n'  # end calc panel

    # ── Modal ───────────────────────────────────────────────────────────────
    html += '''<div class="modal-overlay" id="modal-overlay" onclick="if(event.target===this)closeModal()">
<div class="modal" id="modal">
  <div><span class="modal-close" onclick="closeModal()">✕</span><h2 id="modal-title"></h2></div>
  <div id="modal-isin" style="font-size:12px;color:var(--muted);margin-bottom:12px"></div>
  <div class="modal-kpis" id="modal-kpis"></div>
  <div id="modal-pl-section"></div>
  <div id="modal-transactions" style="margin-top:16px"></div>
  <div class="modal-links" id="modal-links"></div>
</div>
</div>

'''

    # ── Scripts ─────────────────────────────────────────────────────────────
    changes_history_json = json.dumps(changes_history or {}, ensure_ascii=False, separators=(',', ':'))
    html += f'<script>\nconst FUNDS_DATA = {data_json};\n'
    html += f'const NAV_HISTORY = {nav_history_json};\n'
    news_data_json = json.dumps(news_data or {}, ensure_ascii=False, separators=(',', ':'))
    html += f'const NEWS_DATA = {news_data_json};\n'
    html += f'const CHANGES_HISTORY = {changes_history_json};\n'
    run_log_json = json.dumps(run_log or [], ensure_ascii=False, separators=(',', ':'))
    html += f'const RUN_LOG = {run_log_json};\n'
    html += '''
const PAGE_SIZE = 25;
const tableState = {};

// ── Tab switching ──────────────────────────────────────────────────────────
document.querySelectorAll('.tab').forEach(tab => {
  tab.addEventListener('click', () => switchTab(tab.dataset.tab));
});
function switchTab(id) {
  document.querySelectorAll('.tab').forEach(t => t.classList.toggle('active', t.dataset.tab === id));
  document.querySelectorAll('.panel').forEach(p => p.classList.toggle('active', p.id === 'panel-' + id));
  try { sessionStorage.setItem('activeTab', id); } catch(e) {}
}
// Tab nach Reload wiederherstellen
(function() {
  try {
    const saved = sessionStorage.getItem('activeTab');
    if (saved) { sessionStorage.removeItem('activeTab'); switchTab(saved); }
  } catch(e) {}
})();

// ── Number formatting ──────────────────────────────────────────────────────
function fmtEur(n, dec=2) {
  if (n === null || n === undefined || n === '') return '—';
  n = parseFloat(n);
  if (isNaN(n)) return '—';
  if (Math.abs(n) >= 1e9) return (n/1e9).toFixed(2) + ' Mrd.';
  if (Math.abs(n) >= 1e6) return (n/1e6).toFixed(2) + ' Mio.';
  return n.toLocaleString('de-AT', {minimumFractionDigits: dec, maximumFractionDigits: dec});
}
function fmtPct(n) { return n >= 0 ? '+'+n.toFixed(2)+'%' : n.toFixed(2)+'%'; }

// ── Chart defaults ─────────────────────────────────────────────────────────
Chart.defaults.color = '#78716C';
Chart.defaults.borderColor = '#EDE8E1';
Chart.defaults.backgroundColor = '#FFFFFF';
Chart.defaults.font.family = "-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif";
Chart.defaults.font.size = 12;

// ── Overview Charts ────────────────────────────────────────────────────────
const labels   = FUNDS_DATA.map(f => f.id);
const ytdVals  = FUNDS_DATA.map(f => f.perf_ytd || 0);
const aumVals  = FUNDS_DATA.map(f => (f.nav || 0) / 1e6);
const colors   = FUNDS_DATA.map(f => f.color);

new Chart(document.getElementById('cov-perf'), {
  type: 'bar',
  data: {
    labels,
    datasets: [{
      label: 'YTD %',
      data: ytdVals,
      backgroundColor: colors.map(c => c + 'cc'),
      borderColor: colors,
      borderWidth: 1,
      borderRadius: 4,
    }]
  },
  options: {responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}},
    scales:{y:{ticks:{callback:v=>v+'%'}}}}
});

new Chart(document.getElementById('cov-aum'), {
  type: 'doughnut',
  data: {
    labels: FUNDS_DATA.map(f => f.name),
    datasets: [{data: aumVals, backgroundColor: colors, borderWidth: 1}]
  },
  options: {responsive:true, maintainAspectRatio:false, plugins:{legend:{position:'bottom'},
    tooltip:{callbacks:{label:ctx=>ctx.label.split('–')[0]+': '+ctx.raw.toFixed(2)+' Mio. €'}}}}
});


// ── Monatsreporting ────────────────────────────────────────────────────────
(function renderMonthlyTable() {
  const FUND_DEFS = [
    {id: 3411, short: 'Standortfonds AT',      color: '#F97316'},
    {id: 3431, short: 'Standortfonds DE',      color: '#3B7DD8'},
    {id: 3581, short: 'Dividends & Interest',  color: '#16A34A'},
  ];

  // Letzter Eintrag je Monat aus NAV_HISTORY (enthält price = Fondspreis, nav = Nettovermögen)
  function lastEntryPerMonth(fid) {
    const hist = NAV_HISTORY[fid] || [];
    const byMonth = {};
    hist.forEach(h => {
      const m = h.date.substring(0, 7);
      if (!byMonth[m] || h.date > byMonth[m].date) byMonth[m] = h;
    });
    return byMonth; // {YYYY-MM: {date, price, nav}}
  }

  const allMonths = new Set();
  const fundMonthly = {};
  FUND_DEFS.forEach(f => {
    fundMonthly[f.id] = lastEntryPerMonth(f.id);
    Object.keys(fundMonthly[f.id]).forEach(m => allMonths.add(m));
  });

  const today = new Date();
  const currentMonth = today.toISOString().substring(0, 7);
  const months = [...allMonths].filter(m => m < currentMonth).sort().reverse(); // neueste zuerst

  if (months.length === 0) {
    document.getElementById('monthly-table-wrap').innerHTML =
      '<div style="color:var(--muted);font-size:13px;padding:8px">Noch keine abgeschlossenen Monatsdaten vorhanden.</div>';
    return;
  }

  // YTD-Basis: letzter FONDSPREIS (price) des Vorjahres — für Performance-Berechnung
  function getYtdPriceBase(fid, year) {
    const m = fundMonthly[fid];
    const prevYearMonths = Object.keys(m).filter(mo => mo.startsWith((year-1)+'-')).sort();
    if (prevYearMonths.length === 0) {
      const curYearMonths = Object.keys(m).filter(mo => mo.startsWith(year+'-')).sort();
      return curYearMonths.length > 0 ? m[curYearMonths[0]]?.price : null;
    }
    return m[prevYearMonths[prevYearMonths.length - 1]]?.price || null;
  }

  const _fmt = new Intl.NumberFormat('de-AT', {minimumFractionDigits: 2, maximumFractionDigits: 2});
  const fmtMio  = v => v == null ? '—' : _fmt.format(v) + ' €';
  const fmtDMio = v => v == null ? '—' : (v >= 0 ? '+' : '') + _fmt.format(v) + ' €';
  const fmtPct  = v => v == null ? '—' : (v >= 0 ? '+' : '') + v.toFixed(2) + '%';
  const clrPct  = v => v == null ? '' : (v >= 0 ? 'color:#16a34a' : 'color:#dc2626');

  // ── Tabelle ──
  let t = '<div style="overflow-x:auto">';
  t += '<table style="width:100%;border-collapse:collapse;font-size:13px;min-width:720px">';
  t += '<thead>';
  t += '<tr style="border-bottom:2px solid var(--border)">';
  t += '<th style="text-align:left;padding:8px 10px;font-size:11px;color:var(--muted);font-weight:600;white-space:nowrap">Monat</th>';
  FUND_DEFS.forEach(f => {
    t += `<th colspan="4" style="text-align:center;padding:8px 10px;font-size:11px;color:var(--muted);font-weight:700;border-left:2px solid var(--border)">${f.short.toUpperCase()}</th>`;
  });
  t += '</tr>';
  t += '<tr style="border-bottom:1px solid var(--border)">';
  t += '<th></th>';
  FUND_DEFS.forEach(() => {
    t += '<th style="text-align:right;padding:4px 8px;font-size:10px;color:var(--muted);font-weight:600;border-left:2px solid var(--border)" title="Gesamtes Nettovermögen des Fonds (in Mio. €)">Nettovermögen</th>';
    t += '<th style="text-align:right;padding:4px 8px;font-size:10px;color:var(--muted);font-weight:600" title="Veränderung des Nettovermögens zum Vormonat (absolut)">Δ (Mio. €)</th>';
    t += '<th style="text-align:right;padding:4px 8px;font-size:10px;color:var(--muted);font-weight:600" title="Veränderung des Nettovermögens zum Vormonat (prozentual)">Δ (%)</th>';
    t += '<th style="text-align:right;padding:4px 8px;font-size:10px;color:var(--muted);font-weight:600" title="YTD-Performance auf Basis des Fondspreises (Rücknahmepreis) seit Jahresbeginn">YTD *</th>';
  });
  t += '</tr></thead><tbody>';

  months.forEach((month, idx) => {
    const [yr, mo] = month.split('-').map(Number);
    const label = new Date(yr, mo-1, 1).toLocaleString('de-AT', {month: 'long', year: 'numeric'});
    const rowBg = idx % 2 === 0 ? '' : 'background:var(--bg)';
    t += `<tr style="border-bottom:1px solid var(--border);${rowBg}">`;
    t += `<td style="padding:8px 10px;white-space:nowrap;font-weight:600">${label}</td>`;

    FUND_DEFS.forEach(f => {
      const entry    = fundMonthly[f.id][month] ?? null;
      const curNav   = entry?.nav   ?? null;   // Nettovermögen (gesamt)
      const curPrice = entry?.price ?? null;   // Fondspreis (Rücknahmepreis)

      // Vorherigen VORHANDENEN Monat (nicht zwingend Kalendervormonat)
      const fundMonthKeys = Object.keys(fundMonthly[f.id]).filter(m => m < month).sort();
      const prevKey  = fundMonthKeys.length > 0 ? fundMonthKeys[fundMonthKeys.length - 1] : null;
      const prevNav  = prevKey ? (fundMonthly[f.id][prevKey]?.nav   ?? null) : null;

      const deltaAbs = (curNav != null && prevNav != null)            ? curNav - prevNav : null;
      const deltaPct = (curNav != null && prevNav != null && prevNav) ? (curNav - prevNav) / prevNav * 100 : null;

      // YTD: direkt aus dem Inventarblatt (perf_ytd), Fallback auf Fondspreis-Berechnung
      let ytd = entry?.perf_ytd ?? null;
      if (ytd == null) {
        const ytdBase = getYtdPriceBase(f.id, yr);
        ytd = (curPrice != null && ytdBase != null && ytdBase) ? (curPrice - ytdBase) / ytdBase * 100 : null;
      }

      t += `<td style="text-align:right;padding:8px;border-left:2px solid var(--border);white-space:nowrap;font-weight:500">${fmtMio(curNav)}</td>`;
      t += `<td style="text-align:right;padding:8px;white-space:nowrap;${clrPct(deltaAbs)}">${fmtDMio(deltaAbs)}</td>`;
      t += `<td style="text-align:right;padding:8px;white-space:nowrap;font-weight:600;${clrPct(deltaPct)}">${fmtPct(deltaPct)}</td>`;
      t += `<td style="text-align:right;padding:8px;white-space:nowrap;font-weight:600;${clrPct(ytd)}">${fmtPct(ytd)}</td>`;
    });
    t += '</tr>';
  });

  t += '</tbody></table>';
  t += '<div style="font-size:11px;color:var(--muted);padding:6px 10px 2px">* YTD laut Inventarblatt; falls nicht vorhanden: Berechnung auf Basis Fondspreis Monatsultimo vs. letzter Dezember-Preis</div>';
  t += '</div>';

  // ── Balkendiagramm: Nettovermögen je Monat ──
  t += '<div style="margin-top:20px"><canvas id="monthly-nav-chart" height="180"></canvas></div>';

  document.getElementById('monthly-table-wrap').innerHTML = t;

  // Chart aufbauen (chronologisch)
  const chartMonths = [...months].reverse();
  const chartLabels = chartMonths.map(m => {
    const [yr, mo] = m.split('-').map(Number);
    return new Date(yr, mo-1, 1).toLocaleString('de-AT', {month: 'short', year: '2-digit'});
  });
  const ctx = document.getElementById('monthly-nav-chart');
  if (ctx) {
    new Chart(ctx, {
      type: 'bar',
      data: {
        labels: chartLabels,
        datasets: FUND_DEFS.map(f => ({
          label: f.short,
          data: chartMonths.map(m => {
            const v = fundMonthly[f.id][m]?.nav;
            return v != null ? parseFloat((v / 1e6).toFixed(2)) : null;
          }),
          backgroundColor: f.color + 'bb',
          borderColor: f.color,
          borderWidth: 1,
          borderRadius: 3,
        })),
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {position: 'bottom', labels: {font: {size: 11}}},
          tooltip: {
            callbacks: {
              label: ctx => `${ctx.dataset.label}: ${ctx.raw != null ? new Intl.NumberFormat('de-AT',{minimumFractionDigits:2,maximumFractionDigits:2}).format(ctx.raw * 1e6) + ' €' : '—'}`,
            }
          }
        },
        scales: {
          x: {grid: {display: false}},
          y: {
            title: {display: true, text: 'Nettovermögen (Mio. €)', font: {size: 11}},
            ticks: {callback: v => v.toFixed(0) + ' Mio.'},
          }
        }
      }
    });
  }
})();

// ── Per-Fund Charts ────────────────────────────────────────────────────────
FUNDS_DATA.forEach(fund => {
  const fid = fund.id;

  // NAV Chart with history + date range
  const sparkCharts = {};
  function renderSparkChart(fid, fromDate, toDate) {
    const fund = FUNDS_DATA.find(f => f.id === fid);
    if (!fund) return;
    let hist = [...(NAV_HISTORY[fid] || [])];
    // Merge with price_history BVI points as seeds
    (fund.price_history || []).forEach(p => {
      if (!hist.find(h => h.date === p.date)) hist.push({date: p.date, price: p.price});
    });
    hist.sort((a,b) => a.date < b.date ? -1 : 1);
    if (fromDate) hist = hist.filter(p => p.date >= fromDate);
    if (toDate)   hist = hist.filter(p => p.date <= toDate);
    if (!hist.length) return;
    const ctx = document.getElementById('chart-spark-' + fid);
    if (!ctx) return;
    if (sparkCharts[fid]) sparkCharts[fid].destroy();
    sparkCharts[fid] = new Chart(ctx, {
      type: 'line',
      data: {
        labels: hist.map(p => p.date),
        datasets: [{
          label: 'Rücknahmepreis €',
          data: hist.map(p => p.price),
          borderColor: fund.color,
          backgroundColor: fund.color + '18',
          fill: true, tension: 0.3,
          pointRadius: hist.length > 30 ? 0 : 4,
          pointBackgroundColor: fund.color,
        }]
      },
      options: {
        responsive:true, maintainAspectRatio:false,
        plugins:{legend:{display:false}},
        scales:{x:{ticks:{maxTicksLimit:8, maxRotation:0}}, y:{ticks:{callback:v=>v.toFixed(2)+' €'}}}
      }
    });
  }
  function setDateRange(fid, range) {
    const today = new Date();
    let from;
    if (range==='ytd')     from = new Date(today.getFullYear(),0,1);
    else if (range==='gy') from = new Date(today.getMonth()<9 ? today.getFullYear()-1 : today.getFullYear(),9,1);
    else if (range==='3m') from = new Date(today.getFullYear(), today.getMonth()-3, today.getDate());
    else if (range==='1m') from = new Date(today.getFullYear(), today.getMonth()-1, today.getDate());
    else from = null;
    const fromStr = from ? from.toISOString().slice(0,10) : '2000-01-01';
    const toStr = today.toISOString().slice(0,10);
    const fi = document.getElementById('from-'+fid);
    const ti = document.getElementById('to-'+fid);
    if (fi) fi.value = fromStr;
    if (ti) ti.value = toStr;
    renderSparkChart(fid, fromStr, toStr);
  }
  document.getElementById('from-'+fid)?.addEventListener('change', () => {
    renderSparkChart(fid, document.getElementById('from-'+fid).value, document.getElementById('to-'+fid).value);
  });
  document.getElementById('to-'+fid)?.addEventListener('change', () => {
    renderSparkChart(fid, document.getElementById('from-'+fid).value, document.getElementById('to-'+fid).value);
  });
  // Init with GJ range
  setDateRange(fid, 'gy');

  // Currency donut
  if (fund.currencies && fund.currencies.length) {
    const ctx = document.getElementById('chart-ccy-' + fid);
    if (ctx) new Chart(ctx, {
      type: 'doughnut',
      data: {
        labels: fund.currencies.map(c => c.label),
        datasets: [{
          data: fund.currencies.map(c => c.value),
          backgroundColor: ['#1a56db','#d03801','#046c4e','#d29922','#bc8cff','#58a6ff','#3fb950','#f85149'],
          borderWidth: 1,
        }]
      },
      options: {responsive:true, maintainAspectRatio:false, plugins:{legend:{position:'right'}}}
    });
  }

  // P&L by Country
  const plcCtx = document.getElementById('chart-plc-' + fid);
  if (plcCtx) {
    const byC = {};
    (fund.holdings || []).forEach(h => { if (h.pl != null) byC[h.country||'Unbekannt'] = (byC[h.country||'Unbekannt']||0) + h.pl; });
    const cItems = Object.entries(byC).sort((a,b) => b[1]-a[1]).slice(0, 12);
    new Chart(plcCtx, {
      type: 'bar',
      data: {
        labels: cItems.map(x => x[0]),
        datasets: [{
          label: 'P&L €',
          data: cItems.map(x => x[1]),
          backgroundColor: cItems.map(x => x[1] >= 0 ? '#3fb95099' : '#f8514999'),
          borderColor: cItems.map(x => x[1] >= 0 ? '#3fb950' : '#f85149'),
          borderWidth: 1,
          borderRadius: 3,
        }]
      },
      options: {
        indexAxis: 'y',
        responsive: true, maintainAspectRatio: false,
        plugins: {legend: {display: false}},
        scales: {x: {ticks: {callback: v => fmtEur(v) + ' €'}}}
      }
    });
  }

  // P&L by Sector
  const plsCtx = document.getElementById('chart-pls-' + fid);
  if (plsCtx) {
    const byS = {};
    (fund.holdings || []).forEach(h => { if (h.pl != null) byS[h.sector||'Sonstiges'] = (byS[h.sector||'Sonstiges']||0) + h.pl; });
    const sItems = Object.entries(byS).sort((a,b) => b[1]-a[1]).slice(0, 10);
    new Chart(plsCtx, {
      type: 'bar',
      data: {
        labels: sItems.map(x => x[0]),
        datasets: [{
          label: 'P&L €',
          data: sItems.map(x => x[1]),
          backgroundColor: sItems.map(x => x[1] >= 0 ? '#3fb95099' : '#f8514999'),
          borderColor: sItems.map(x => x[1] >= 0 ? '#3fb950' : '#f85149'),
          borderWidth: 1,
          borderRadius: 3,
        }]
      },
      options: {
        indexAxis: 'y',
        responsive: true, maintainAspectRatio: false,
        plugins: {legend: {display: false}},
        scales: {x: {ticks: {callback: v => fmtEur(v) + ' €'}}}
      }
    });
  }

  // Table pagination init
  tableState[fid] = { page: 1, sortKey: 'mv_eur', sortDir: -1, filter: '', filterCountry: '', filterSector: '' };
  renderTable(fid);
});

// ── Table ──────────────────────────────────────────────────────────────────
function getHoldingsForFund(fid) {
  const f = FUNDS_DATA.find(f => f.id === fid);
  return (f && f.holdings) || [];
}

function filterTable(fid) {
  const q = (document.getElementById('search-' + fid)?.value || '').toLowerCase();
  const fc = (document.getElementById('filter-country-' + fid)?.value || '');
  const fs = (document.getElementById('filter-sector-' + fid)?.value || '');
  tableState[fid].filter = q;
  tableState[fid].filterCountry = fc;
  tableState[fid].filterSector  = fs;
  tableState[fid].page = 1;
  renderTable(fid);
}

function sortTable(fid, key) {
  const st = tableState[fid];
  if (st.sortKey === key) st.sortDir *= -1;
  else { st.sortKey = key; st.sortDir = -1; }
  st.page = 1;
  renderTable(fid);
}

function renderTable(fid) {
  const tbody = document.getElementById('tbody-' + fid);
  if (!tbody) return;
  const rows = Array.from(tbody.querySelectorAll('tr'));
  const st = tableState[fid];
  const nav = (FUNDS_DATA.find(f=>f.id===fid)||{}).nav || 1;

  // Filter
  const visible = rows.filter(row => {
    const name = row.dataset.name?.toLowerCase() || '';
    const isin = row.dataset.isin?.toLowerCase() || '';
    const ctry = row.dataset.country || '';
    const sec  = row.dataset.sector || '';
    if (st.filter && !name.includes(st.filter) && !isin.includes(st.filter) && !ctry.toLowerCase().includes(st.filter) && !sec.toLowerCase().includes(st.filter)) return false;
    if (st.filterCountry && ctry !== st.filterCountry) return false;
    if (st.filterSector  && sec  !== st.filterSector)  return false;
    return true;
  });

  // Sort
  const keyMap = {name:'name',isin:'isin',country:'country',sector:'sector',currency:'currency',
    mv_eur:'mv',pl:'pl',nav_pct:'navPct',cost:'cost',price:'price'};
  const dataKey = {name:'name',isin:'isin',country:'country',sector:'sector',currency:'currency',
    mv_eur:'mv',pl:'pl',nav_pct:'navPct',cost:'cost',price:'price'};
  const dsKey = {name:'name',isin:'isin',country:'country',sector:'sector',currency:'currency',
    mv_eur:'mv',pl:'pl',nav_pct:'navPct',cost:'cost',price:'price'};

  visible.sort((a, b) => {
    let av, bv;
    if (st.sortKey === 'name')    { av = a.dataset.name; bv = b.dataset.name; return av < bv ? -st.sortDir : st.sortDir; }
    if (st.sortKey === 'isin')    { av = a.dataset.isin; bv = b.dataset.isin; return av < bv ? -st.sortDir : st.sortDir; }
    if (st.sortKey === 'country') { av = a.dataset.country; bv = b.dataset.country; return av < bv ? -st.sortDir : st.sortDir; }
    if (st.sortKey === 'sector')  { av = a.dataset.sector; bv = b.dataset.sector; return av < bv ? -st.sortDir : st.sortDir; }
    if (st.sortKey === 'currency'){ av = a.dataset.currency; bv = b.dataset.currency; return av < bv ? -st.sortDir : st.sortDir; }
    if (st.sortKey === 'mv_eur')  { av = parseFloat(a.dataset.mv||0); bv = parseFloat(b.dataset.mv||0); }
    else if (st.sortKey === 'pl') { av = parseFloat(a.dataset.pl||0); bv = parseFloat(b.dataset.pl||0); }
    else if (st.sortKey === 'nav_pct') { av = parseFloat(a.dataset.navPct||0); bv = parseFloat(b.dataset.navPct||0); }
    else if (st.sortKey === 'cost')  { av = parseFloat(a.dataset.cost||0); bv = parseFloat(b.dataset.cost||0); }
    else if (st.sortKey === 'price') { av = parseFloat(a.dataset.price||0); bv = parseFloat(b.dataset.price||0); }
    else { av = 0; bv = 0; }
    return (av - bv) * st.sortDir;
  });

  // Pagination
  const total = visible.length;
  const pages = Math.ceil(total / PAGE_SIZE);
  const start = (st.page - 1) * PAGE_SIZE;

  // Reorder DOM so sort actually takes effect
  const visSet = new Set(visible);
  const hiddenRows = rows.filter(r => !visSet.has(r));
  const frag = document.createDocumentFragment();
  visible.forEach(r => frag.appendChild(r));
  hiddenRows.forEach(r => frag.appendChild(r));
  tbody.appendChild(frag);

  rows.forEach(r => r.style.display = 'none');
  visible.slice(start, start + PAGE_SIZE).forEach(r => r.style.display = '');

  // Page buttons
  const pagesEl = document.getElementById('pages-' + fid);
  if (pagesEl) {
    pagesEl.innerHTML = '';
    for (let i = 1; i <= pages; i++) {
      const btn = document.createElement('button');
      btn.className = 'page-btn' + (i === st.page ? ' active' : '');
      btn.textContent = i;
      btn.onclick = () => { tableState[fid].page = i; renderTable(fid); };
      pagesEl.appendChild(btn);
    }
    const info = document.createElement('span');
    info.style = 'font-size:12px;color:var(--muted);align-self:center;margin-left:8px';
    info.textContent = `${total} Positionen`;
    pagesEl.appendChild(info);
  }
}

// ── Modal ──────────────────────────────────────────────────────────────────
function showModal(row) {
  const d = row.dataset;
  document.getElementById('modal-title').textContent = d.name;
  document.getElementById('modal-isin').textContent = d.isin || '';

  const pl = parseFloat(d.pl || 0);
  const mv = parseFloat(d.mv || 0);
  const plCls = pl >= 0 ? 'pos' : 'neg';
  const plSign = pl >= 0 ? '+' : '';

  document.getElementById('modal-kpis').innerHTML = `
    <div class="modal-kpi"><div class="lbl">Marktwert EUR</div><div class="val">${fmtEur(mv)} €</div></div>
    <div class="modal-kpi"><div class="lbl">P&L EUR</div><div class="val ${plCls}">${plSign}${fmtEur(pl)} €</div></div>
    <div class="modal-kpi"><div class="lbl">% NAV</div><div class="val">${parseFloat(d.navPct||0).toFixed(2)}%</div></div>
    <div class="modal-kpi"><div class="lbl">Einstandskurs</div><div class="val">${d.cost ? fmtEur(d.cost) : '—'}</div></div>
    <div class="modal-kpi"><div class="lbl">Aktueller Kurs</div><div class="val">${d.price ? fmtEur(d.price) : '—'}</div></div>
    <div class="modal-kpi"><div class="lbl">Währung</div><div class="val">${d.currency || '—'}</div></div>
  `;

  const isin = d.isin || '';
  document.getElementById('modal-links').innerHTML = isin ? `
    <a href="https://finance.yahoo.com/quote/${isin}" target="_blank">Yahoo Finance</a>
    <a href="https://www.boerse.de/aktien/kurs/${isin}" target="_blank">Boerse.de</a>
    <a href="https://www.wienerborse.at/en/market-data/securities/share-detail/?ISIN=${isin}" target="_blank">Wiener Börse</a>
    <a href="https://www.finanzen.net/suche/?_search=${isin}" target="_blank">Finanzen.net</a>
    <a href="https://www.onvista.de/suche/?searchValue=${isin}" target="_blank">Onvista</a>
  ` : '';

  // Transaktionshistorie für diese ISIN
  const isinSearch = (d.isin || '').trim();
  const txEl = document.getElementById('modal-transactions');
  const TX_CONFIG = {
    added:     {label:'Neukauf',        color:'#16a34a', icon:'▲'},
    removed:   {label:'Komplettverkauf',color:'#dc2626', icon:'✕'},
    increased: {label:'Aufstockung',    color:'#2563eb', icon:'↑'},
    decreased: {label:'Teilverkauf',    color:'#ea580c', icon:'↓'},
  };
  const rowFid = d.fid || '';
  const allTx = [];
  if (isinSearch && rowFid) {
    // Nur Transaktionen aus dem Fonds des angeklickten Unternehmens
    const fundEntries = CHANGES_HISTORY[rowFid] || [];
    fundEntries.forEach(e => { if (e.isin === isinSearch) allTx.push({...e, fid: rowFid}); });
  }
  allTx.sort((a,b) => b.date.localeCompare(a.date));
  if (allTx.length > 0) {
    let html = '<div style="font-size:12px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px">Transaktionshistorie</div>';
    html += '<table style="width:100%;border-collapse:collapse;font-size:13px">';
    html += '<thead><tr style="border-bottom:1px solid var(--border)">'
          + '<th style="text-align:left;padding:6px 8px;font-size:11px;color:var(--muted);font-weight:600">Datum</th>'
          + '<th style="text-align:left;padding:6px 8px;font-size:11px;color:var(--muted);font-weight:600">Typ</th>'
          + '<th style="text-align:right;padding:6px 8px;font-size:11px;color:var(--muted);font-weight:600">Gehandelt</th>'
          + '<th style="text-align:right;padding:6px 8px;font-size:11px;color:var(--muted);font-weight:600">Nach Trade</th>'
          + '</tr></thead><tbody>';
    // Hilfsfunktion: Marktwert formatieren
    const fmtMv = v => {
      if (v == null) return '—';
      const abs = Math.abs(v);
      if (abs >= 1e6) return (v/1e6).toFixed(2) + ' Mio. €';
      return fmtEur(Math.abs(v)) + ' €';
    };
    allTx.forEach(e => {
      const cfg = TX_CONFIG[e.type] || {label:e.type, color:'#6b7280', icon:'•'};
      const badge = `<span style="display:inline-flex;align-items:center;gap:3px;padding:2px 7px;border-radius:999px;font-size:11px;font-weight:600;color:${cfg.color};background:${cfg.color}18">${cfg.icon} ${cfg.label}</span>`;
      // Gehandelten Betrag und Restposition aus mv_eur + change_pct ableiten
      let traded = null, afterMv = null;
      if (e.type === 'added') {
        traded = e.mv_eur; afterMv = e.mv_eur;
      } else if (e.type === 'removed') {
        traded = e.mv_eur ? -e.mv_eur : null; afterMv = 0;
      } else if (e.change_pct != null && e.mv_eur) {
        const prevMv = e.mv_eur / (1 + e.change_pct / 100);
        traded = e.mv_eur - prevMv;   // positiv = Kauf, negativ = Verkauf
        afterMv = e.mv_eur;
      }
      // Gehandelt-Zelle
      let tradedHtml = '—';
      if (traded != null) {
        const sign = traded >= 0 ? '+' : '';
        const tc = traded >= 0 ? '#16a34a' : '#dc2626';
        tradedHtml = `<span style="color:${tc};font-weight:600">${sign}${fmtMv(traded)}</span>`;
      }
      // Nach-Trade-Zelle
      let afterHtml = '—';
      if (e.type === 'removed') {
        afterHtml = '<span style="color:var(--muted);font-size:12px">Position geschlossen</span>';
      } else if (afterMv != null) {
        afterHtml = `<span style="color:var(--muted)">${fmtMv(afterMv)}</span>`;
      }
      html += `<tr style="border-bottom:1px solid var(--border)">
        <td style="padding:8px;white-space:nowrap;color:var(--muted)">${e.date}</td>
        <td style="padding:8px">${badge}</td>
        <td style="padding:8px;text-align:right;white-space:nowrap">${tradedHtml}</td>
        <td style="padding:8px;text-align:right;white-space:nowrap;font-size:12px">${afterHtml}</td>
      </tr>`;
    });
    html += '</tbody></table>';
    txEl.innerHTML = html;
  } else {
    txEl.innerHTML = isinSearch ? '<div style="font-size:12px;color:var(--muted);margin-top:4px">Keine Transaktionen erfasst</div>' : '';
  }

  document.getElementById('modal-overlay').classList.add('open');
}
function closeModal() { document.getElementById('modal-overlay').classList.remove('open'); }
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); });

// ── CSV Export ─────────────────────────────────────────────────────────────
function exportCSV(fid) {
  const fund = FUNDS_DATA.find(f => f.id === fid);
  if (!fund) return;
  const rows = [['Name','ISIN','Land','Sektor','Währung','Marktwert EUR','P&L EUR','% NAV','Einstand','Kurs']];
  const nav = fund.nav || 1;
  (fund.holdings || []).forEach(h => {
    rows.push([h.name, h.isin, h.country, h.sector, h.currency,
      h.mv_eur?.toFixed(2) || '', h.pl?.toFixed(2) || '',
      ((h.mv_eur || 0) / nav * 100).toFixed(3),
      h.cost?.toFixed(4) || '', h.price?.toFixed(4) || '']);
  });
  const csv = rows.map(r => r.map(v => '"' + String(v||'').replace(/"/g, '""') + '"').join(',')).join('\\n');
  const a = document.createElement('a');
  a.href = 'data:text/csv;charset=utf-8,' + encodeURIComponent(csv);
  a.download = `IQAM_${fid}_${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
}

// ── Common Holdings Table ──────────────────────────────────────────────────
// (already rendered server-side)

// ── Performance Calculator ─────────────────────────────────────────────────
let calcChart = null;
function setCalcRange(range) {
  const today = new Date();
  let from;
  if (range==='ytd')     from = new Date(today.getFullYear(),0,1);
  else if (range==='gy') from = new Date(today.getMonth()<9 ? today.getFullYear()-1 : today.getFullYear(),9,1);
  else if (range==='1y') from = new Date(today.getFullYear()-1, today.getMonth(), today.getDate());
  else from = null;
  const fi = document.getElementById('calc-from');
  const ti = document.getElementById('calc-to');
  if (fi) fi.value = from ? from.toISOString().slice(0,10) : '2020-01-01';
  if (ti) ti.value = today.toISOString().slice(0,10);
  calcPerf();
}
function calcPerf() {
  const amount  = parseFloat(document.getElementById('calc-amount')?.value) || 10000;
  const fromDate = document.getElementById('calc-from')?.value || '';
  const toDate   = document.getElementById('calc-to')?.value   || '';
  const datasets = [];

  FUNDS_DATA.forEach(fund => {
    const el = document.getElementById('calc-detail-' + fund.id);
    if (!el) return;

    // Merge NAV_HISTORY with price_history seeds
    let hist = [...(NAV_HISTORY[fund.id] || [])];
    (fund.price_history || []).forEach(p => {
      if (!hist.find(h => h.date === p.date)) hist.push({date: p.date, price: p.price});
    });
    hist.sort((a,b) => a.date < b.date ? -1 : 1);

    // Find closest prices
    const afterFrom = hist.filter(p => !fromDate || p.date >= fromDate);
    const beforeTo  = hist.filter(p => !toDate   || p.date <= toDate);
    if (!afterFrom.length || !beforeTo.length) {
      el.innerHTML = '<span style="color:var(--muted)">Keine Daten für diesen Zeitraum</span>';
      return;
    }
    const startEntry = afterFrom[0];
    const endEntry   = beforeTo[beforeTo.length-1];
    if (startEntry.date === endEntry.date) {
      el.innerHTML = '<span style="color:var(--muted)">Start = Ende, bitte anderen Zeitraum wählen</span>';
      return;
    }

    const startPrice = startEntry.price, endPrice = endEntry.price;
    const units   = amount / startPrice;
    const current = units * endPrice;
    const gain    = current - amount;
    const retPct  = (endPrice - startPrice) / startPrice * 100;
    // Annualized return
    const days    = (new Date(endEntry.date) - new Date(startEntry.date)) / 86400000;
    const annRet  = days > 0 ? (Math.pow(endPrice/startPrice, 365/days) - 1) * 100 : 0;
    const cls  = gain >= 0 ? 'pos' : 'neg';
    const s    = gain >= 0 ? '+' : '';

    el.innerHTML = `
      <div style="margin-bottom:12px;padding:12px;background:var(--surface2);border-radius:8px">
        <div style="font-size:11px;color:var(--muted);margin-bottom:4px">${startEntry.date} → ${endEntry.date}</div>
        <div style="font-size:24px;font-weight:700" class="${cls}">${s}${fmtEur(gain)} €</div>
        <div style="font-size:13px;color:var(--muted)">aus ${fmtEur(amount)} € → <strong style="color:var(--text)">${fmtEur(current)} €</strong></div>
      </div>
      <table style="width:100%;font-size:13px">
        <tr><td style="color:var(--muted);padding:4px 0">Einstiegskurs</td><td style="text-align:right;font-weight:600">${startPrice.toFixed(4)} €</td></tr>
        <tr><td style="color:var(--muted);padding:4px 0">Endkurs</td><td style="text-align:right;font-weight:600">${endPrice.toFixed(4)} €</td></tr>
        <tr><td style="color:var(--muted);padding:4px 0">Anteile</td><td style="text-align:right">${units.toFixed(3)}</td></tr>
        <tr><td style="color:var(--muted);padding:4px 0">Gesamtrendite</td><td style="text-align:right;font-weight:700" class="${cls}">${s}${retPct.toFixed(2)}%</td></tr>
        <tr><td style="color:var(--muted);padding:4px 0">Annualisiert</td><td style="text-align:right" class="${cls}">${s}${annRet.toFixed(2)}% p.a.</td></tr>
        <tr><td style="color:var(--muted);padding:4px 0">Zeitraum</td><td style="text-align:right">${Math.round(days)} Tage</td></tr>
      </table>
    `;

    if (hist.length >= 2) {
      const base = startPrice;
      datasets.push({
        label: fund.name,
        data: hist.filter(p => (!fromDate||p.date>=fromDate)&&(!toDate||p.date<=toDate))
                  .map(p => ({x: p.date, y: (p.price-base)/base*100})),
        borderColor: fund.color,
        backgroundColor: 'transparent',
        tension: 0.3,
        pointRadius: 0,
        borderWidth: 2,
      });
    }
  });

  const ctx = document.getElementById('chart-calc-perf');
  if (ctx) {
    if (calcChart) calcChart.destroy();
    calcChart = new Chart(ctx, {
      type: 'line',
      data: { datasets },
      options: {
        responsive: true, maintainAspectRatio: false,
        parsing: false,
        plugins: { legend: { position: 'top' }, tooltip: { mode: 'index' } },
        scales: {
          x: { type: 'time', time: { unit: 'month' }, ticks: { maxTicksLimit: 10 } },
          y: { ticks: { callback: v => v.toFixed(1)+'%' } }
        }
      }
    });
  }
}
// Init calculator with YTD
setTimeout(() => setCalcRange('ytd'), 600);
</script>
</body>
</html>'''

    return html


def _build_all_holdings_table(funds_data):
    """Alle Positionen aller Fonds mit Haken-Spalten pro Fonds."""
    # Sammle alle Positionen, dedupliziert per ISIN (oder Name als Fallback)
    items = {}  # key -> {isin, name, country, sector, funds: {fid: mv_eur}}
    for f in funds_data:
        fid = f["id"]
        for h in f.get("holdings", []):
            isin = (h.get("isin") or "").strip()
            name = (h.get("name") or "").strip()
            if not name or name in ("None",):
                continue
            key = isin if (isin and isin not in ("None", "")) else name
            if key not in items:
                items[key] = {
                    "isin": isin if isin not in ("None", "") else "",
                    "name": name,
                    "country": h.get("country", "") or "",
                    "sector": h.get("sector", "") or "",
                    "funds": {}
                }
            items[key]["funds"][fid] = h.get("mv_eur", 0) or 0

    # Sortierung: zuerst Positionen in mehreren Fonds, dann nach Gesamt-Marktwert
    sorted_items = sorted(items.values(),
                          key=lambda x: (-len(x["funds"]), -sum(x["funds"].values())))

    if not sorted_items:
        return '<div class="placeholder">Keine Positionen gefunden</div>'

    fund_ids   = [f["id"]   for f in funds_data]
    fund_short = [f["name"].split("–")[-1].strip()[:18] for f in funds_data]

    html = '<div class="tbl-controls">\n'
    html += '<input type="text" id="search-all" placeholder="Name, ISIN, Land …" oninput="_allState.page=1;renderAllTable()">\n'
    html += '<select id="filter-all-funds" onchange="_allState.page=1;renderAllTable()">\n'
    html += '<option value="">Alle Fonds</option>\n'
    html += '<option value="multi">In mehreren Fonds</option>\n'
    for fid, fname in zip(fund_ids, fund_short):
        html += f'<option value="{fid}">{fid} – {fname}</option>\n'
    html += '</select>\n</div>\n'

    html += '<div class="tbl-wrap"><table id="tbl-all">\n'
    html += '<thead><tr>'
    html += '<th onclick="sortAllTable(\'name\')">Name ↕</th>'
    html += '<th onclick="sortAllTable(\'isin\')" style="width:110px">ISIN ↕</th>'
    html += '<th onclick="sortAllTable(\'country\')">Land ↕</th>'
    html += '<th onclick="sortAllTable(\'sector\')">Sektor ↕</th>'
    for fname in fund_short:
        html += f'<th style="text-align:center;width:90px">{fname}</th>'
    html += '</tr></thead>\n<tbody id="tbody-all">\n'

    for item in sorted_items:
        html += f'<tr data-name="{item["name"].lower()}" data-isin="{item["isin"].lower()}" data-country="{(item["country"] or "").lower()}" data-funds="{",".join(item["funds"].keys())}">'
        html += f'<td>{item["name"][:45]}</td>'
        html += f'<td style="font-size:11px;color:var(--muted)">{item["isin"]}</td>'
        html += f'<td>{item["country"]}</td>'
        html += f'<td>{item["sector"]}</td>'
        for fid in fund_ids:
            if fid in item["funds"]:
                mv = item["funds"][fid]
                html += f'<td style="text-align:center" title="{mv/1e6:.2f} Mio. €"><span style="color:var(--green);font-size:15px;font-weight:700">✓</span></td>'
            else:
                html += '<td style="text-align:center;color:var(--border)">—</td>'
        html += '</tr>\n'

    html += '</tbody></table></div>\n'
    html += '<div class="pagination" id="pages-all"></div>\n'

    html += '''<script>
const _allState = {page:1, sortKey:'name', sortDir:1};
function renderAllTable() {
  const tbody = document.getElementById('tbody-all');
  if (!tbody) return;
  const q    = (document.getElementById('search-all')?.value || '').toLowerCase();
  const ff   = document.getElementById('filter-all-funds')?.value || '';
  const rows = Array.from(tbody.querySelectorAll('tr'));

  const visible = rows.filter(r => {
    if (q && !r.dataset.name?.includes(q) && !r.dataset.isin?.includes(q) && !r.dataset.country?.includes(q)) return false;
    if (ff === 'multi' && (r.dataset.funds||'').split(',').filter(Boolean).length < 2) return false;
    if (ff && ff !== 'multi' && !(r.dataset.funds||'').split(',').includes(ff)) return false;
    return true;
  });

  visible.sort((a,b) => {
    const av = a.dataset[_allState.sortKey] || '', bv = b.dataset[_allState.sortKey] || '';
    return av < bv ? -_allState.sortDir : av > bv ? _allState.sortDir : 0;
  });

  const total = visible.length, pages = Math.ceil(total/25), start = (_allState.page-1)*25;

  // Reorder DOM so sort takes effect
  const _visSet = new Set(visible);
  const _hidden = rows.filter(r => !_visSet.has(r));
  const _frag = document.createDocumentFragment();
  visible.forEach(r => _frag.appendChild(r));
  _hidden.forEach(r => _frag.appendChild(r));
  tbody.appendChild(_frag);

  rows.forEach(r => r.style.display='none');
  visible.slice(start, start+25).forEach(r => r.style.display='');

  const pel = document.getElementById('pages-all');
  if (pel) {
    pel.innerHTML = '';
    for (let i=1; i<=Math.min(pages,20); i++) {
      const b = document.createElement('button');
      b.className = 'page-btn' + (i===_allState.page ? ' active' : '');
      b.textContent = i;
      b.onclick = () => { _allState.page=i; renderAllTable(); };
      pel.appendChild(b);
    }
    const sp = document.createElement('span');
    sp.style = 'font-size:12px;color:var(--muted);align-self:center;margin-left:8px';
    sp.textContent = total + ' Positionen';
    pel.appendChild(sp);
  }
}
function sortAllTable(key) {
  if (_allState.sortKey===key) _allState.sortDir*=-1;
  else { _allState.sortKey=key; _allState.sortDir=1; }
  _allState.page=1; renderAllTable();
}
setTimeout(renderAllTable, 200);

// ── News ──────────────────────────────────────────────────────────────────
function _escH(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function _fmtD(s){
  if (!s) return '';
  try { const d=new Date(s); return d.toLocaleDateString('de-AT',{day:'2-digit',month:'2-digit'})+' '+d.toLocaleTimeString('de-AT',{hour:'2-digit',minute:'2-digit'}); }
  catch(e){ return ''; }
}

function renderNewsPanel(fundFilter) {
  const feed = document.getElementById('news-feed');
  const status = document.getElementById('news-status');
  if (!feed) return;

  // Build fund badge lookup
  const fundMap = {};
  FUNDS_DATA.forEach(f => {
    fundMap[f.id] = { color: f.color, short: f.name.replace('Standortfonds ','SF ').replace('Dividends and Interest','D&I') };
  });

  // Filter companies by fund + optional summary-only filter + keyword search
  const onlySummary = document.getElementById('news-filter-summary')?.checked;
  const searchQ = (document.getElementById('news-search')?.value || '').toLowerCase().trim();
  const companies = Object.values(NEWS_DATA).filter(co => {
    if (fundFilter && fundFilter !== 'all' && !co.funds.includes(fundFilter)) return false;
    if (onlySummary && !co.summary) return false;
    if (searchQ) {
      const haystack = [
        co.company,
        co.summary?.headline || '',
        co.summary?.text || '',
        ...(co.articles || []).map(a => a.title || '')
      ].join(' ').toLowerCase();
      if (!haystack.includes(searchQ)) return false;
    }
    return (co.articles || []).length > 0;
  });

  // Sort by newest article date descending
  companies.sort((a, b) => {
    const da = new Date((a.articles[0]?.pubDate) || 0);
    const db = new Date((b.articles[0]?.pubDate) || 0);
    return db - da;
  });

  if (!companies.length) {
    feed.innerHTML = '<div class="placeholder" style="padding:40px">Keine News verfügbar</div>';
    if (status) status.textContent = 'Keine Daten';
    return;
  }

  feed.innerHTML = companies.map(co => {
    const badges = (co.funds||[]).map(fid => {
      const f = fundMap[fid];
      if (!f) return '';
      return `<span style="background:${f.color}20;color:${f.color};border:1px solid ${f.color}40;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:700">${_escH(f.short)}</span>`;
    }).join('');

    const newestDate = _fmtD((co.articles[0]?.pubDate) || '');

    const s = co.summary;
    const summaryHtml = s
      ? `${s.headline ? `<div style="margin:10px 0 4px;font-size:14px;font-weight:700;color:var(--text);line-height:1.4">${_escH(s.headline)}</div>` : ''}
         <p style="margin:0;font-size:13px;line-height:1.65;color:var(--text)">${_escH(s.text || '')}</p>`
      : (co.articles[0] ? `<p style="margin:10px 0 0;font-size:13px;line-height:1.65;color:var(--muted);font-style:italic">${_escH(co.articles[0].title)}</p>` : '');

    const sources = (co.articles||[]).map(a => {
      let label = a.source || '';
      if (!label && a.link && a.link !== '#') {
        try { label = new URL(a.link).hostname.replace(/^www\./,''); } catch(e) {}
      }
      label = label || 'Quelle';
      const dateStr = a.pubDate ? ' <span style="color:var(--muted);font-size:11px">(' + _fmtD(a.pubDate) + ')</span>' : '';
      const href = _escH(a.link || '#');
      return `<a href="${href}" target="_blank" rel="noopener noreferrer"
        title="${_escH(a.title)}"
        style="color:var(--accent);font-size:12px;text-decoration:none"
        onmouseover="this.style.textDecoration='underline'"
        onmouseout="this.style.textDecoration='none'">${_escH(label)}</a>${dateStr}`;
    }).join('<span style="color:var(--border);margin:0 6px">·</span>');

    return `<div class="card" style="padding:16px 18px">
  <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
    <span style="font-size:13px;font-weight:700;color:var(--text)">${_escH(co.company)}</span>
    ${badges}
    <span style="font-size:11px;color:var(--muted);margin-left:auto;white-space:nowrap">${newestDate}</span>
  </div>
  ${summaryHtml}
  <div style="margin-top:10px;padding-top:8px;border-top:1px solid var(--border);display:flex;align-items:center;gap:4px;flex-wrap:wrap">
    <span style="font-size:11px;color:var(--muted);margin-right:2px">Quellen:</span>${sources}
  </div>
</div>`;
  }).join('');

  const lastRun = RUN_LOG[0]?.ts ? new Date(RUN_LOG[0].ts).toLocaleString('de-AT',{day:'2-digit',month:'2-digit',hour:'2-digit',minute:'2-digit'}) + ' Uhr' : '—';
  if (status) status.textContent = `${companies.length} Unternehmen · Stand: ${lastRun}`;
}

// Fund filter for news
let _newsFundFilter = 'all';
function setNewsFund(fid) {
  _newsFundFilter = fid;
  document.querySelectorAll('[id^="newsf-"]').forEach(b => {
    const active = b.id === 'newsf-' + fid;
    b.classList.toggle('active', active);
  });
  renderNewsPanel(_newsFundFilter);
}

document.querySelectorAll('.tab').forEach(t => {
  if (t.dataset.tab === 'news') t.addEventListener('click', () => renderNewsPanel(_newsFundFilter));
});
setTimeout(() => { if (document.getElementById('panel-news')?.classList.contains('active')) renderNewsPanel('all'); }, 300);

// ── Run-Log Modal ─────────────────────────────────────────────────────────────
function toggleRunLog() {
  let el = document.getElementById('run-log-modal');
  if (el) { el.remove(); return; }

  const fmt = ts => {
    try {
      const d = new Date(ts);
      const today = new Date().toDateString() === d.toDateString();
      const dateStr = today ? 'Heute' : d.toLocaleDateString('de-AT',{day:'2-digit',month:'2-digit'});
      return dateStr + ' · ' + d.toLocaleTimeString('de-AT',{hour:'2-digit',minute:'2-digit'}) + ' Uhr';
    } catch(e) { return ts || '—'; }
  };

  const rows = RUN_LOG.length ? RUN_LOG.map((r,i) => {
    const ok = r.status === 'success';
    const isLatest = i === 0;
    return `<tr style="border-bottom:1px solid var(--border);${isLatest?'background:var(--surface2)':''}">
      <td style="padding:10px 16px;font-size:13px;font-weight:${isLatest?700:400};white-space:nowrap">${fmt(r.ts)}</td>
      <td style="padding:10px 16px;text-align:center;font-size:15px">${ok ? '✅' : '❌'}</td>
    </tr>`;
  }).join('') : '<tr><td colspan="2" style="padding:24px;text-align:center;color:var(--muted)">Noch keine Runs</td></tr>';

  el = document.createElement('div');
  el.id = 'run-log-modal';
  el.style.cssText = 'position:fixed;top:58px;left:16px;z-index:9999;background:var(--surface);border:1px solid var(--border);border-radius:10px;box-shadow:0 8px 32px #0002;min-width:380px;max-width:95vw;max-height:70vh;overflow:auto';
  el.innerHTML = `
    <div style="display:flex;align-items:center;justify-content:space-between;padding:12px 16px;border-bottom:1px solid var(--border)">
      <span style="font-weight:700;font-size:13px">Tägliche Daten-Updates</span>
      <button onclick="document.getElementById('run-log-modal').remove()"
        style="background:none;border:none;font-size:18px;cursor:pointer;color:var(--muted);line-height:1;padding:0 4px">×</button>
    </div>
    <table style="width:100%;border-collapse:collapse">
      <thead><tr style="border-bottom:1px solid var(--border)">
        <th style="padding:8px 16px;font-size:11px;color:var(--muted);text-align:left;font-weight:600;text-transform:uppercase;letter-spacing:.5px">Datum</th>
        <th style="padding:8px 16px;font-size:11px;color:var(--muted);text-align:center;font-weight:600;text-transform:uppercase;letter-spacing:.5px">Status</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>`;
  document.body.appendChild(el);

  setTimeout(() => {
    document.addEventListener('click', function handler(e) {
      if (!el.contains(e.target) && !e.target.closest('button[onclick="toggleRunLog()"]')) {
        el.remove();
        document.removeEventListener('click', handler);
      }
    });
  }, 100);
}

// ── Dark Mode ──────────────────────────────────────────────────────────────────
(function() {
  if (localStorage.getItem('darkMode') === '1') {
    document.body.classList.add('dark');
    const btn = document.getElementById('dark-toggle');
    if (btn) btn.textContent = '☀️';
  }
})();
function toggleDark() {
  const dark = document.body.classList.toggle('dark');
  localStorage.setItem('darkMode', dark ? '1' : '0');
  const btn = document.getElementById('dark-toggle');
  if (btn) btn.textContent = dark ? '☀️' : '🌙';
  // Redraw all charts so colors update
  Object.values(Chart.instances).forEach(c => {
    const style = getComputedStyle(document.body);
    if (c.options.plugins?.legend?.labels) {
      c.options.plugins.legend.labels.color = style.getPropertyValue('--text').trim();
    }
    c.update();
  });
}

// ── Donut Charts (Sektor, Länder, Währung) ─────────────────────────────────────
const DONUT_PALETTE = ['#F97316','#3B7DD8','#16A34A','#DC2626','#8B5CF6','#06B6D4','#D97706','#EC4899','#6B7280','#14B8A6','#F59E0B','#10B981'];

function _aggregate(holdings, key, topN) {
  const map = {};
  const total = holdings.reduce((s, h) => s + (h.mv_eur || 0), 0);
  holdings.forEach(h => {
    const k = h[key] || 'Sonstige';
    map[k] = (map[k] || 0) + (h.mv_eur || 0);
  });
  let entries = Object.entries(map).sort((a,b) => b[1]-a[1]);
  if (topN && entries.length > topN) {
    const rest = entries.slice(topN).reduce((s,e) => s+e[1], 0);
    entries = entries.slice(0, topN);
    if (rest > 0) entries.push(['Sonstige', rest]);
  }
  return { labels: entries.map(e => e[0]), data: entries.map(e => e[1]), total };
}

function _makeDonut(canvasId, agg) {
  const ctx = document.getElementById(canvasId);
  if (!ctx) return;
  const style = getComputedStyle(document.body);
  new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: agg.labels,
      datasets: [{
        data: agg.data,
        backgroundColor: DONUT_PALETTE.slice(0, agg.labels.length),
        borderWidth: 2,
        borderColor: style.getPropertyValue('--surface').trim() || '#fff'
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'bottom',
          labels: { font: { size: 10 }, padding: 6, color: style.getPropertyValue('--text').trim(), boxWidth: 10 }
        },
        tooltip: {
          callbacks: {
            label: ctx => {
              const pct = agg.total > 0 ? (ctx.raw / agg.total * 100).toFixed(1) : '0';
              return ` ${ctx.label}: ${pct}%`;
            }
          }
        }
      },
      cutout: '60%'
    }
  });
}

function initDonutCharts() {
  FUNDS_DATA.forEach(fund => {
    const panel = document.getElementById('panel-' + fund.id);
    if (!panel || !fund.holdings?.length) return;

    const wrap = document.createElement('div');
    wrap.innerHTML = `
      <div class="section-title">📊 Analyse</div>
      <div class="grid-3" style="margin-bottom:28px">
        <div class="card"><h3>Sektoren</h3><div class="chart-wrap" style="height:240px"><canvas id="cs-${fund.id}"></canvas></div></div>
        <div class="card"><h3>Länder (Top 8)</h3><div class="chart-wrap" style="height:240px"><canvas id="cl-${fund.id}"></canvas></div></div>
        <div class="card"><h3>Währungsrisiko</h3><div class="chart-wrap" style="height:240px"><canvas id="cw-${fund.id}"></canvas></div></div>
      </div>`;
    panel.appendChild(wrap);

    _makeDonut('cs-' + fund.id, _aggregate(fund.holdings, 'sector', null));
    _makeDonut('cl-' + fund.id, _aggregate(fund.holdings, 'country', 8));
    _makeDonut('cw-' + fund.id, _aggregate(fund.holdings, 'currency', null));
  });
}
setTimeout(initDonutCharts, 150);

// ── Konfetti 🎉 ────────────────────────────────────────────────────────────────
(function() {
  if (typeof confetti === 'undefined') return;
  const allPositive = FUNDS_DATA.every(f => {
    return (f.holdings || []).reduce((s, h) => s + (h.pl || 0), 0) > 0;
  });
  if (allPositive) {
    setTimeout(() => {
      confetti({ particleCount: 120, spread: 80, origin: { y: 0.5 }, colors: ['#F97316','#16A34A','#3B7DD8','#FBBF24'] });
      setTimeout(() => confetti({ particleCount: 80, spread: 100, origin: { x: 0.2, y: 0.6 }, colors: ['#F97316','#16A34A'] }), 400);
      setTimeout(() => confetti({ particleCount: 80, spread: 100, origin: { x: 0.8, y: 0.6 }, colors: ['#3B7DD8','#FBBF24'] }), 700);
    }, 1200);
  }
})();
</script>'''
    return html


# ─── GitHub Commit ────────────────────────────────────────────────────────────
def git_push_file(token, repo, path, content_bytes, message, branch="main"):
    """Committed eine Datei direkt per GitHub API. Retry bei 5xx-Fehlern."""
    import base64
    import time as _time
    b64_content = base64.b64encode(content_bytes).decode()

    def _api(req, retries=3):
        for attempt in range(retries):
            try:
                with urlopen(req, timeout=30) as resp:
                    return json.loads(resp.read())
            except HTTPError as e:
                if e.code in (502, 503, 504) and attempt < retries - 1:
                    _time.sleep(3 * (attempt + 1))
                    continue
                raise
        raise RuntimeError("Unreachable")

    # Bestehende SHA holen (falls Datei existiert) — 5xx → Datei als neu behandeln
    sha = None
    try:
        sha_req = Request(
            f"https://api.github.com/repos/{repo}/contents/{path}?ref={branch}",
            headers={"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"},
        )
        existing = _api(sha_req)
        sha = existing.get("sha")
    except HTTPError as e:
        if e.code != 404:
            print(f"  ⚠️  SHA-Abfrage {path}: HTTP {e.code} – fahre ohne SHA fort")
    except Exception as e:
        print(f"  ⚠️  SHA-Abfrage {path}: {e} – fahre ohne SHA fort")

    # Commit — bei 422 ohne SHA nochmal versuchen (SHA-Konflikt)
    for attempt in range(2):
        body = {"message": message, "content": b64_content, "branch": branch}
        if sha:
            body["sha"] = sha
        put_req = Request(
            f"https://api.github.com/repos/{repo}/contents/{path}",
            data=json.dumps(body).encode(),
            method="PUT",
            headers={
                "Authorization": f"token {token}",
                "Accept":        "application/vnd.github.v3+json",
                "Content-Type":  "application/json",
            },
        )
        try:
            result = _api(put_req)
            break
        except HTTPError as e:
            if e.code == 422 and sha and attempt == 0:
                # SHA veraltet → SHA weglassen (Datei neu anlegen)
                print(f"  ⚠️  SHA-Konflikt {path} – versuche ohne SHA")
                sha = None
                continue
            raise
    else:
        raise RuntimeError(f"Push fehlgeschlagen: {path}")

    commit_sha = result.get("commit", {}).get("sha", "")[:8]
    print(f"  ✅ Committed {path} → {commit_sha}")
    return result


def get_or_create_gh_pages(token, repo):
    """Aktiviert GitHub Pages (docs/ Ordner auf main)."""
    # Pages API
    req = Request(
        f"https://api.github.com/repos/{repo}/pages",
        data=json.dumps({"source": {"branch": "main", "path": "/docs"}}).encode(),
        method="POST",
        headers={
            "Authorization": f"token {token}",
            "Accept": "application/vnd.github.v3+json",
            "Content-Type": "application/json",
        },
    )
    try:
        with urlopen(req) as resp:
            pages = json.loads(resp.read())
            url = pages.get("html_url", "")
            print(f"  🌐 GitHub Pages aktiviert: {url}")
            return url
    except HTTPError as e:
        if e.code == 409:
            # Bereits aktiviert – URL lesen
            req2 = Request(
                f"https://api.github.com/repos/{repo}/pages",
                headers={"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"},
            )
            with urlopen(req2) as resp:
                pages = json.loads(resp.read())
                url = pages.get("html_url", "")
                print(f"  🌐 GitHub Pages bereits aktiv: {url}")
                return url
        else:
            print(f"  ⚠️  Pages-Fehler {e.code}: {e.read()[:200]}")
            return ""


# ─── Prev Data laden/speichern ────────────────────────────────────────────────
def load_prev_data(token, repo, branch="main"):
    """Liest prev_data.json aus dem Repo (für Änderungserkennung)."""
    try:
        req = Request(
            f"https://api.github.com/repos/{repo}/contents/docs/prev_data.json?ref={branch}",
            headers={"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"},
        )
        with urlopen(req) as resp:
            meta = json.loads(resp.read())
            content = base64.b64decode(meta["content"]).decode()
            return json.loads(content)
    except Exception as e:
        print(f"  ℹ️  Keine prev_data.json gefunden ({e}), starte frisch")
        return {}


def load_run_log(token, repo, branch="main"):
    """Liest den Run-Log aus docs/run_log.json (letzte N Runs)."""
    try:
        req = Request(
            f"https://api.github.com/repos/{repo}/contents/docs/run_log.json?ref={branch}",
            headers={"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"},
        )
        with urlopen(req) as resp:
            meta = json.loads(resp.read())
            content = base64.b64decode(meta["content"]).decode()
            return json.loads(content)
    except Exception:
        return []


def load_nav_history(token, repo, branch="main"):
    """Liest akkumulierte NAV-Historie aus docs/nav_history.json."""
    try:
        req = Request(
            f"https://api.github.com/repos/{repo}/contents/docs/nav_history.json?ref={branch}",
            headers={"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"},
        )
        with urlopen(req) as resp:
            meta = json.loads(resp.read())
            content = base64.b64decode(meta["content"]).decode()
            return json.loads(content)
    except Exception as e:
        print(f"  ℹ️  Keine nav_history.json ({e}), starte frisch")
        return {}


def load_json_from_github(token, repo, path, branch="main"):
    """Lädt eine beliebige JSON-Datei aus dem Repo."""
    try:
        req = Request(
            f"https://api.github.com/repos/{repo}/contents/{path}?ref={branch}",
            headers={"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"},
        )
        with urlopen(req) as resp:
            meta = json.loads(resp.read())
            content = base64.b64decode(meta["content"]).decode()
            return json.loads(content)
    except Exception:
        return None


# ─── Main ─────────────────────────────────────────────────────────────────────
def main():
    github_token = os.environ.get("GITHUB_TOKEN", "")
    github_repo  = os.environ.get("GITHUB_REPOSITORY", "")
    RUN_MODE     = os.environ.get("RUN_MODE", "full").lower()  # "full" oder "news"

    print("=" * 60)
    print(f"🚀 IQAM Dashboard Update – {date.today()} [{RUN_MODE.upper()}]")
    print("=" * 60)

    # 2. Cached data laden
    prev_data = {}
    nav_history = {}
    run_log = []
    changes_history = {}
    holdings_prev = {}  # {fid: {"date": "...", "isins": [...]}}
    if github_token and github_repo:
        nav_history = load_nav_history(github_token, github_repo)
        run_log = load_run_log(github_token, github_repo)
        changes_history = load_json_from_github(github_token, github_repo, "docs/changes_history.json") or {}
        holdings_prev = load_json_from_github(github_token, github_repo, "docs/holdings_prev.json") or {}

    if RUN_MODE == "backfill":
        # ── Backfill: Alle historischen Holdings aus Outlook-Mails laden ───────
        print("\n🔄 BACKFILL-Modus: Lade alle historischen INVENTARLISTE-Mails…")
        access_token = get_access_token()
        holdings_history = backfill_holdings_history(access_token, {})

        # Changes-History aus vollständiger holdings_history neu aufbauen (inkl. Teilkäufe/-verkäufe)
        print("\n🔁 Baue Transaktionshistorie aus Holdings-History auf…")
        changes_history = {}
        for fid, snaps in holdings_history.items():
            changes_history[fid] = []
            existing_keys = set()
            sorted_dates = sorted(snaps.keys())
            for i in range(1, len(sorted_dates)):
                d_curr = sorted_dates[i]
                d_prev = sorted_dates[i - 1]
                curr_snap = {h["isin"]: h for h in snaps[d_curr] if h.get("isin")}
                prev_snap = {h["isin"]: h for h in snaps[d_prev] if h.get("isin")}
                for isin, h in curr_snap.items():
                    prev_h = prev_snap.get(isin)
                    if prev_h is None:
                        key = (isin, d_curr, "added")
                        if key not in existing_keys:
                            changes_history[fid].append({"date": d_curr, "type": "added", "isin": isin, "name": h.get("name",""), "mv_eur": h.get("mv_eur"), "qty": h.get("qty")})
                            existing_keys.add(key)
                    else:
                        prev_qty = prev_h.get("qty") or 0
                        curr_qty = h.get("qty") or 0
                        if prev_qty and curr_qty:
                            # Qty-basierte Erkennung (präzise)
                            if abs(curr_qty - prev_qty) / max(abs(prev_qty), 1) > 0.005:
                                diff_pct = (curr_qty - prev_qty) / abs(prev_qty) * 100
                                typ = "increased" if curr_qty > prev_qty else "decreased"
                                key = (isin, d_curr, typ)
                                if key not in existing_keys:
                                    price_per_share = round(h.get("mv_eur") / curr_qty, 4) if (h.get("mv_eur") and curr_qty) else None
                                    changes_history[fid].append({"date": d_curr, "type": typ, "isin": isin, "name": h.get("name",""), "mv_eur": h.get("mv_eur"), "qty": curr_qty, "prev_qty": prev_qty, "change_pct": round(diff_pct, 1), "price_per_share": price_per_share})
                                    existing_keys.add(key)
                        else:
                            # Fallback: mv_eur-Proxy wenn qty nicht im Excel (>10% Schwelle)
                            prev_mv = prev_h.get("mv_eur") or 0
                            curr_mv = h.get("mv_eur") or 0
                            if prev_mv and curr_mv and abs(curr_mv - prev_mv) / max(abs(prev_mv), 1) > 0.10:
                                diff_pct = (curr_mv - prev_mv) / abs(prev_mv) * 100
                                typ = "increased" if curr_mv > prev_mv else "decreased"
                                key = (isin, d_curr, typ)
                                if key not in existing_keys:
                                    changes_history[fid].append({"date": d_curr, "type": typ, "isin": isin, "name": h.get("name",""), "mv_eur": curr_mv, "change_pct": round(diff_pct, 1), "mv_proxy": True})
                                    existing_keys.add(key)
                for isin, h in prev_snap.items():
                    if isin not in curr_snap:
                        key = (isin, d_curr, "removed")
                        if key not in existing_keys:
                            changes_history[fid].append({"date": d_curr, "type": "removed", "isin": isin, "name": h.get("name",""), "mv_eur": h.get("mv_eur"), "qty": h.get("qty")})
                            existing_keys.add(key)
            total_ch = len(changes_history[fid])
            print(f"  📊 {fid}: {len(sorted_dates)} Tage, {total_ch} Transaktionen erkannt")

        # Pushen — changes_history zuerst (wichtig!), holdings_history nicht pushen (zu groß)
        # Stattdessen: letzten Snapshot als holdings_prev speichern (für zukünftige tägliche Vergleiche)
        today_str = date.today().isoformat()
        holdings_prev_new = {}
        for fid, snaps in holdings_history.items():
            if snaps:
                last_date = sorted(snaps.keys())[-1]
                holdings_prev_new[fid] = {
                    "date": last_date,
                    "isins": {h["isin"]: {"name": h.get("name",""), "mv_eur": h.get("mv_eur")}
                              for h in snaps[last_date] if h.get("isin")}
                }
        # NAV-History aus INVENTARBLATT-Mails backfillen
        print("\n📈 Backfille NAV-History aus INVENTARBLATT-Mails…")
        nav_history = backfill_nav_history_from_emails(access_token, nav_history)

        if github_token and github_repo:
            print("\n📤 Pushe changes_history.json, holdings_prev.json und nav_history.json…")
            git_push_file(github_token, github_repo, "docs/changes_history.json",
                         json.dumps(changes_history, ensure_ascii=False).encode("utf-8"),
                         f"Backfill changes history {today_str}")
            git_push_file(github_token, github_repo, "docs/holdings_prev.json",
                         json.dumps(holdings_prev_new, ensure_ascii=False).encode("utf-8"),
                         f"Backfill holdings prev {today_str}")
            git_push_file(github_token, github_repo, "docs/nav_history.json",
                         json.dumps(nav_history, ensure_ascii=False).encode("utf-8"),
                         f"Backfill nav history {today_str}")
        print("\n✅ Backfill abgeschlossen!")
        return

    if RUN_MODE == "news":
        # ── News-only Run: Fondsdaten aus Cache laden ──────────────────────────
        print("\n📂 News-only Run: Lade gecachte Fondsdaten…")
        cached = load_json_from_github(github_token, github_repo, "docs/prev_data.json")
        if not cached:
            print("❌ Keine gecachten Fondsdaten gefunden. Bitte zuerst Full-Run ausführen.")
            sys.exit(1)
        funds_data = list(cached.values())
        print(f"  ✅ {len(funds_data)} Fonds aus Cache geladen")
    else:
        # ── Full Run: Fondsdaten frisch aus Email holen ────────────────────────
        # 1. MS Token
        access_token = get_access_token()

        # 2. Prev data für Änderungserkennung
        if github_token and github_repo:
            prev_data = load_prev_data(github_token, github_repo)

        # 3. Mails finden
        fund_mails = find_latest_emails(access_token)

    # 4. Pro Fund Excel laden + parsen (nur Full-Run)
    if RUN_MODE != "news":
        funds_data = []
        for fund_meta in FUNDS:
            fid        = fund_meta["id"]
            mail_entry = fund_mails.get(fid, {})
            mail_blatt = mail_entry.get("blatt")
            mail_liste = mail_entry.get("liste")
    
            if not mail_blatt and not mail_liste:
                print(f"⚠️  Keine Mail für Fund {fid} gefunden!")
                prev_fund = prev_data.get(fid, {})
                if prev_fund:
                    print(f"   → Verwende gestrige Daten für {fid}")
                    funds_data.append({**fund_meta, **prev_fund, "changes": {}})
                continue
    
            fund_parsed = {}
    
            # INVENTARBLATT → NAV, Preis, YTD, FY
            if mail_blatt:
                print(f"\n📋 Fund {fid} BLATT: {mail_blatt['subject'][:55]}")
                xlsx_bytes, filename = download_attachment(access_token, mail_blatt["id"], ".xlsx")
                if xlsx_bytes:
                    print(f"   Parsing {filename}…")
                    try:
                        blatt_data = parse_excel(xlsx_bytes, fid)
                        fund_parsed.update(blatt_data)
                    except Exception as e:
                        print(f"   ❌ BLATT Parse-Fehler: {e}")
                        traceback.print_exc()
    
            # INVENTARLISTE → Holdings, Länder, Währungen, Sektoren
            mail_for_holdings = mail_liste or mail_blatt  # Fallback auf BLATT wenn keine LISTE
            if mail_for_holdings:
                print(f"\n📊 Fund {fid} LISTE: {mail_for_holdings['subject'][:55]}")
                xlsx_bytes, filename = download_attachment(access_token, mail_for_holdings["id"], ".xlsx")
                if xlsx_bytes:
                    print(f"   Parsing {filename}…")
                    try:
                        liste_data = parse_excel(xlsx_bytes, fid)
                        # Merge: LISTE-Daten überschreiben nur Holdings/Allokation, nicht NAV
                        for key in ["holdings", "countries", "currencies", "sectors"]:
                            if key in liste_data:
                                fund_parsed[key] = liste_data[key]
                        # NAV aus BLATT bevorzugen — nur übernehmen wenn noch nicht gesetzt
                        for key in ["nav", "nav_per_share", "shares", "perf_ytd", "perf_fy", "report_date"]:
                            if key not in fund_parsed and key in liste_data:
                                fund_parsed[key] = liste_data[key]
                    except Exception as e:
                        print(f"   ❌ LISTE Parse-Fehler: {e}")
                        traceback.print_exc()
    
            # Vortags-Preis: aus prev_data wenn run_date vor heute gesetzt, sonst nav_history
            # run_date wird von uns gesetzt (nicht aus Excel) → zuverlässig unabhängig von NAV-Datum
            prev_fund = prev_data.get(fid, {})
            today_iso = date.today().isoformat()
            prev_run_date = prev_fund.get("run_date", "")
            if prev_fund and prev_run_date and prev_run_date < today_iso:
                # Vorheriger Run war gestern (oder früher) → direkt als Baseline
                nav_ps_prev = prev_fund.get("nav_per_share")
            else:
                # Kein valides prev_data → nav_history als Fallback (nur real gemessene Punkte)
                hist_entries = [h for h in nav_history.get(fid, []) if h["date"] < today_iso and h.get("source") == "measured"]
                if not hist_entries:
                    # Alle Punkte (auch Seed-Punkte) falls keine gemessenen vorhanden
                    hist_entries = [h for h in nav_history.get(fid, []) if h["date"] < today_iso]
                nav_ps_prev = hist_entries[-1]["price"] if hist_entries else None
            fund_parsed["nav_per_share_prev"] = nav_ps_prev
            fund_parsed["nav_prev"] = prev_fund.get("nav") if prev_fund else None
    
            # Price history aufbauen
            fund_parsed["price_history"] = build_price_history(
                fund_parsed.get("nav_per_share"),
                fund_parsed.get("perf_ytd"),
                fund_parsed.get("perf_fy"),
                nav_ps_prev,
            )
    
            # Änderungen erkennen
            prev_holdings = prev_fund.get("holdings", []) if prev_fund else []
            changes = detect_changes(fund_parsed.get("holdings", []), prev_holdings)
            changes["date_prev"] = prev_fund.get("report_date") if prev_fund else None
            fund_parsed["changes"] = changes

            # Tägliche Änderungen vs. letztem Snapshot erkennen und changes_history aktualisieren
            today_str = date.today().isoformat()
            curr_map = {h.get("isin"): h for h in fund_parsed.get("holdings", [])
                        if h.get("isin") and h["isin"] not in ("None","")}
            prev_snap_entry = holdings_prev.get(fid, {})
            prev_isin_map = prev_snap_entry.get("isins", {})
            prev_date = prev_snap_entry.get("date", "")

            if fid not in changes_history:
                changes_history[fid] = []
            existing_keys = {(e["isin"], e["date"], e["type"]) for e in changes_history[fid]}

            if prev_isin_map and prev_date and prev_date < today_str:
                for isin, h in curr_map.items():
                    prev_info = prev_isin_map.get(isin)
                    if prev_info is None:
                        # Neukauf (komplett neu)
                        key = (isin, today_str, "added")
                        if key not in existing_keys:
                            changes_history[fid].append({"date": today_str, "type": "added", "isin": isin, "name": h.get("name",""), "mv_eur": h.get("mv_eur"), "qty": h.get("qty")})
                            existing_keys.add(key)
                    else:
                        # Teilkauf / Teilverkauf (Position bleibt, Menge ändert sich)
                        prev_qty = prev_info.get("qty") or 0
                        curr_qty = h.get("qty") or 0
                        if prev_qty and curr_qty:
                            # Qty-basierte Erkennung (präzise)
                            if abs(curr_qty - prev_qty) / max(abs(prev_qty), 1) > 0.005:
                                diff_pct = (curr_qty - prev_qty) / abs(prev_qty) * 100
                                typ = "increased" if curr_qty > prev_qty else "decreased"
                                key = (isin, today_str, typ)
                                if key not in existing_keys:
                                    price_per_share = round(h.get("mv_eur") / curr_qty, 4) if (h.get("mv_eur") and curr_qty) else None
                                    changes_history[fid].append({"date": today_str, "type": typ, "isin": isin, "name": h.get("name",""), "mv_eur": h.get("mv_eur"), "qty": curr_qty, "prev_qty": prev_qty, "change_pct": round(diff_pct, 1), "price_per_share": price_per_share})
                                    existing_keys.add(key)
                        else:
                            # Fallback: mv_eur-Proxy wenn qty nicht im Excel vorhanden (>10% Schwelle)
                            prev_mv = prev_info.get("mv_eur") or 0
                            curr_mv = h.get("mv_eur") or 0
                            if prev_mv and curr_mv and abs(curr_mv - prev_mv) / max(abs(prev_mv), 1) > 0.10:
                                diff_pct = (curr_mv - prev_mv) / abs(prev_mv) * 100
                                typ = "increased" if curr_mv > prev_mv else "decreased"
                                key = (isin, today_str, typ)
                                if key not in existing_keys:
                                    changes_history[fid].append({"date": today_str, "type": typ, "isin": isin, "name": h.get("name",""), "mv_eur": curr_mv, "change_pct": round(diff_pct, 1), "mv_proxy": True})
                                    existing_keys.add(key)
                for isin, info in prev_isin_map.items():
                    if isin not in curr_map:
                        # Komplettverkauf
                        key = (isin, today_str, "removed")
                        if key not in existing_keys:
                            changes_history[fid].append({"date": today_str, "type": "removed", "isin": isin, "name": info.get("name",""), "mv_eur": info.get("mv_eur"), "qty": info.get("qty")})
                            existing_keys.add(key)

            # Aktuellen Snapshot als neues holdings_prev speichern (inkl. qty für Teilkauf/-verkauf-Erkennung)
            holdings_prev[fid] = {
                "date": today_str,
                "isins": {isin: {"name": h.get("name",""), "mv_eur": h.get("mv_eur"), "qty": h.get("qty")} for isin, h in curr_map.items()}
            }
    
            # Vortags-Holdings für Tagesvergleich (lean – nur nötige Felder)
            fund_parsed["prev_holdings"] = [
                {"isin": h.get("isin"), "name": h.get("name"), "mv_eur": h.get("mv_eur"),
                 "pl": h.get("pl"), "qty": h.get("qty")}
                for h in prev_holdings if h.get("isin") and h["isin"] not in ("None", "")
            ]
    
            # KPIs berechnen
            fund_parsed = compute_kpis(fund_parsed)
    
            # Merge mit Fund-Metadaten
            funds_data.append({**fund_meta, **fund_parsed})
    
            print(f"   ✅ {fid}: NAV {fund_parsed.get('nav',0)/1e6:.2f} Mio., "
                  f"Preis {fund_parsed.get('nav_per_share',0):.4f}, "
                  f"YTD {fund_parsed.get('perf_ytd',0):.2f}%")

    if not funds_data:
        print("❌ Keine Daten gefunden. Abbruch.")
        sys.exit(1)

    # 4b. NAV-Historie aktualisieren (nur Full-Run)
    today_str = date.today().isoformat()
    if RUN_MODE != "news":
        for fund in funds_data:
            fid      = fund["id"]
            price    = fund.get("nav_per_share")
            nav      = fund.get("nav")
            shares   = fund.get("shares")
            perf_ytd = fund.get("perf_ytd")
            # Fallback: Nettovermögen = Preis × Anteile wenn direkt nicht gefunden
            if nav is None and price and shares:
                nav = float(price) * float(shares)
            if price and price > 0:
                if fid not in nav_history:
                    nav_history[fid] = []
                # BVI-Seed-Punkte hinzufügen (GJ-Start, YTD-Start)
                for ph_point in fund.get("price_history", []):
                    if not any(h["date"] == ph_point["date"] for h in nav_history[fid]):
                        nav_history[fid].append({"date": ph_point["date"], "price": ph_point["price"]})
                # Heutigen Datenpunkt hinzufügen (als "measured" markiert, für zuverlässige Baseline)
                if not any(h["date"] == today_str for h in nav_history[fid]):
                    nav_history[fid].append({
                        "date": today_str,
                        "price": round(price, 4),
                        "nav": round(nav, 2) if nav else None,
                        "perf_ytd": round(float(perf_ytd), 4) if perf_ytd is not None else None,
                        "source": "measured",
                    })
                nav_history[fid].sort(key=lambda x: x["date"])
                print(f"  📈 {fid} NAV-Historie: {len(nav_history[fid])} Punkte")

    # 5. News fetchen (max 2× pro Tag)
    today_str = date.today().isoformat()
    news_runs_today = sum(
        1 for e in run_log
        if e.get("ts", "").startswith(today_str) and e.get("news", 0) > 0
    )
    prev_news_data = {}
    if github_token and github_repo:
        prev_news_data = load_json_from_github(github_token, github_repo, "docs/news_data.json") or {}

    if news_runs_today >= 2 and prev_news_data:
        print(f"\n⏭️  News bereits {news_runs_today}× heute aktualisiert – überspringe News-Fetch.")
        news_data = prev_news_data
    else:
        companies_for_news = {}
        for fund in funds_data:
            fid = fund["id"]
            for h in fund.get("holdings", []):
                isin = (h.get("isin") or "").strip()
                name = (h.get("name") or "").strip()
                if not name or name in ("None", ""):
                    continue
                key = isin if (isin and isin not in ("None", "")) else name
                if key not in companies_for_news:
                    companies_for_news[key] = {"name": name, "funds": [], "mv": 0}
                companies_for_news[key]["mv"] += h.get("mv_eur") or 0
                if fid not in companies_for_news[key]["funds"]:
                    companies_for_news[key]["funds"].append(fid)
        companies_for_news = dict(sorted(companies_for_news.items(), key=lambda x: x[1]["mv"], reverse=True))
        anthropic_key = os.environ.get("ANTHROPIC_API_KEY", "")
        news_data = fetch_all_news(companies_for_news, anthropic_key=anthropic_key, prev_news_data=prev_news_data, max_summaries=10, max_wall_seconds=240)
        # Fallback: wenn Fetch nichts liefert (z.B. Google blockiert GitHub IPs) → Altdaten behalten
        if not news_data and prev_news_data:
            print("  ↩️  News-Fetch lieferte nichts – verwende gecachte Altdaten.")
            news_data = prev_news_data

    # 5b. Run-Log Eintrag erstellen
    run_entry = {
        "ts": datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
        "status": "success",
        "funds": len(funds_data),
        "holdings": sum(len(f.get("holdings", [])) for f in funds_data),
        "news": len(news_data),
        "summaries": sum(1 for v in news_data.values() if v.get("summary")),
        "aum": sum(f.get("nav", 0) for f in funds_data),
    }
    run_log.insert(0, run_entry)
    run_log = run_log[:30]  # max 30 Einträge

    # 6. Dashboard generieren
    updated_at = datetime.now().strftime("%d.%m.%Y %H:%M UTC")
    print(f"\n🔨 Generiere Dashboard ({updated_at})…")
    html = generate_html(funds_data, updated_at, nav_history=nav_history, news_data=news_data, run_log=run_log, changes_history=changes_history)
    data_json = json.dumps(
        [{k: v for k, v in f.items() if k != "holdings"} | {"holdings": f.get("holdings", [])}
         for f in funds_data],
        ensure_ascii=False, indent=2
    )

    # 7. In GitHub pushen
    if github_token and github_repo:
        print(f"\n📤 Push zu {github_repo}…")
        today_str = date.today().isoformat()
        git_push_file(github_token, github_repo, "docs/index.html",
                     html.encode("utf-8"),
                     f"Dashboard update {today_str}")
        # Jekyll-Verarbeitung deaktivieren (verhindert Build-Fehler durch {{ }} in JSON-Daten)
        git_push_file(github_token, github_repo, "docs/.nojekyll",
                     b"",
                     "Disable Jekyll")
        git_push_file(github_token, github_repo, "docs/dashboard_data.json",
                     data_json.encode("utf-8"),
                     f"Data update {today_str}")
        # News-Daten speichern — nur wenn vorhanden (verhindert Überschreiben mit leerem Dict)
        if news_data:
            git_push_file(github_token, github_repo, "docs/news_data.json",
                         json.dumps(news_data, ensure_ascii=False).encode("utf-8"),
                         f"News update {today_str}")
        # Nur beim Full-Run: Prev-Data, NAV-Historie speichern
        if RUN_MODE != "news":
            prev_save = {f["id"]: {**{k: v for k, v in f.items() if k not in ("changes",)}, "run_date": today_str}
                         for f in funds_data}
            git_push_file(github_token, github_repo, "docs/prev_data.json",
                         json.dumps(prev_save, ensure_ascii=False).encode("utf-8"),
                         f"Prev data {today_str}")
            git_push_file(github_token, github_repo, "docs/nav_history.json",
                         json.dumps(nav_history, ensure_ascii=False).encode("utf-8"),
                         f"NAV history {today_str}")
            git_push_file(github_token, github_repo, "docs/changes_history.json",
                         json.dumps(changes_history, ensure_ascii=False).encode("utf-8"),
                         f"Changes history {today_str}")
            git_push_file(github_token, github_repo, "docs/holdings_prev.json",
                         json.dumps(holdings_prev, ensure_ascii=False).encode("utf-8"),
                         f"Holdings prev {today_str}")
        git_push_file(github_token, github_repo, "docs/run_log.json",
                     json.dumps(run_log, ensure_ascii=False).encode("utf-8"),
                     f"Run log {today_str}")
        # GitHub Pages aktivieren (idempotent)
        pages_url = get_or_create_gh_pages(github_token, github_repo)
        if pages_url:
            print(f"\n🌐 Dashboard URL: {pages_url}")
    else:
        # Lokaler Modus: Dateien schreiben
        print("\n💾 Lokaler Modus (kein GITHUB_TOKEN gesetzt)")
        with open("docs/index.html", "w", encoding="utf-8") as f:
            f.write(html)
        with open("docs/dashboard_data.json", "w", encoding="utf-8") as f:
            f.write(data_json)
        print("   → docs/index.html")
        print("   → docs/dashboard_data.json")

    print("\n✅ Dashboard-Update abgeschlossen!")


if __name__ == "__main__":
    main()
