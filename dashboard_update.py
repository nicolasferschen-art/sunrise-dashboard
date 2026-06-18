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
    {"id": "3411", "isin": "AT0000A1QA38", "color": "#1a56db",
     "name": "3411 – IQAM Standortfonds Österreich"},
    {"id": "3431", "isin": "AT0000A1Z882", "color": "#d03801",
     "name": "3431 – IQAM Standortfonds Deutschland"},
    {"id": "3581", "isin": "AT0000A3EAW0", "color": "#046c4e",
     "name": "3581 – IQAM Sunrise Dividends and Interest"},
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

    for i, row in enumerate(rows):
        row_str = " ".join(str(c) for c in row if c is not None).upper()

        # Gesamtvermögen / NAV — viele mögliche Labels
        if any(k in row_str for k in ["GESAMTVERM", "FONDSVERM", "NETTOVERM", "TOTAL NET ASSET",
                                       "TOTAL ASSETS", "FUND VOLUME", "INVENTARWERT GESAMT",
                                       "NET ASSET VALUE", "GESAMT"]):
            for cell in row:
                if isinstance(cell, (int, float)) and cell > 1_000_000:
                    data["nav"] = float(cell)
                    print(f"    → NAV gefunden Zeile {i+1}: {cell}")
                    break

        # Rücknahmepreis / NAV per share
        if any(k in row_str for k in ["RÜCKNAHME", "ANTEILSWERT", "INVENTARWERT JE",
                                       "REDEMPTION PRICE", "NET ASSET VALUE PER",
                                       "PRICE PER UNIT", "VALUE PER SHARE", "UNIT VALUE",
                                       "ANTEILSPR"]):
            for cell in row:
                if isinstance(cell, (int, float)) and 1 < cell < 100_000:
                    data["nav_per_share"] = float(cell)
                    print(f"    → Preis gefunden Zeile {i+1}: {cell}")
                    break

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

    # Datum aus Zellen
    for row in rows[:10]:
        for cell in row:
            if isinstance(cell, (datetime, date)):
                data["report_date"] = str(cell)[:10]
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

    # Spalten-Indices bestimmen
    col_map = {}
    for j, h in enumerate(headers):
        hu = h.upper()
        if "ISIN" in hu:                       col_map["isin"] = j
        elif "NAME" in hu or "BEZEICH" in hu:  col_map["name"] = j
        elif "COUNTRY" in hu or "LAND" in hu:  col_map["country"] = j
        elif "SECTOR" in hu or "BRANCHE" in hu or "SEC.TYPE" in hu: col_map["sector"] = j
        elif "CURRENCY" in hu or "WÄHRUNG" in hu: col_map["currency"] = j
        # Marktwert in Fondswährung (EUR) – Index 15 aus Erfahrung
        elif "MKT VAL" in hu and ("EUR" in hu or "FNDCCY" in hu or "FUND" in hu):
            col_map["mv_eur"] = j
        elif "P&L" in hu or "G&V" in hu or "GEWINN" in hu:
            col_map["pl"] = j
        elif "WEIGHT" in hu or "ANTEIL" in hu or "%" in hu:
            col_map.setdefault("weight", j)
        elif "COST" in hu or "EINSTAND" in hu or "KAUFPREIS" in hu:
            col_map["cost"] = j
        elif "PRICE" in hu or "KURS" in hu:
            col_map.setdefault("price", j)
        elif "QUANTITY" in hu or "STÜCK" in hu or "NOMINAL" in hu:
            col_map["qty"] = j

    # Fallback für Marktwert: Spalte 15 (aus Analyse der echten Daten)
    if "mv_eur" not in col_map:
        col_map["mv_eur"] = 15

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
    """Liest Allokations-Tabellen (Country / Currency / Sector)."""
    rows = list(ws.iter_rows(values_only=True))
    result = []
    for row in rows:
        if len(row) < 2:
            continue
        label = row[0]
        if not label or str(label).strip() in ("", "None", "Total", "Gesamt", "Land", "Country"):
            continue
        # Wert: erste numerische Spalte
        for cell in row[1:]:
            if isinstance(cell, (int, float)) and cell != 0:
                try:
                    result.append({"label": str(label).strip(), "value": float(cell)})
                except (TypeError, ValueError):
                    pass
                break
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
            c_w = curr_map[isin].get("mv_eur", 0)
            p_w = prev_map[isin].get("mv_eur", 0)
            if c_w and p_w and abs(c_w - p_w) / max(abs(p_w), 1) > 0.005:
                diff_pct = (c_w - p_w) / abs(p_w) * 100
                entry = {**curr_map[isin], "change_pct": round(diff_pct, 2),
                         "prev_mv": p_w, "curr_mv": c_w}
                if c_w > p_w:
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


# ─── Dashboard HTML generieren ────────────────────────────────────────────────
def generate_html(funds_data, updated_at):
    """Generiert das vollständige Dashboard-HTML."""
    data_json = json.dumps(funds_data, ensure_ascii=False, separators=(',', ':'))

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
<style>
:root {{
  --bg: #0d1117; --surface: #161b22; --surface2: #21262d;
  --border: #30363d; --text: #e6edf3; --muted: #8b949e;
  --blue: #58a6ff; --green: #3fb950; --red: #f85149;
  --orange: #d29922; --purple: #bc8cff;
}}
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{ background: var(--bg); color: var(--text); font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; font-size: 14px; }}
a {{ color: var(--blue); text-decoration: none; }}
.sticky-header {{ position: sticky; top: 0; z-index: 100; background: var(--surface); border-bottom: 1px solid var(--border); padding: 12px 24px; display: flex; align-items: center; justify-content: space-between; }}
.sticky-header h1 {{ font-size: 18px; font-weight: 700; color: var(--text); }}
.header-kpis {{ display: flex; gap: 32px; }}
.hkpi {{ text-align: right; }}
.hkpi .val {{ font-size: 16px; font-weight: 700; }}
.hkpi .lbl {{ font-size: 11px; color: var(--muted); }}
.tabs {{ display: flex; gap: 4px; padding: 12px 24px 0; border-bottom: 1px solid var(--border); overflow-x: auto; background: var(--surface); }}
.tab {{ padding: 8px 16px; border-radius: 6px 6px 0 0; cursor: pointer; font-size: 13px; color: var(--muted); border: 1px solid transparent; border-bottom: none; white-space: nowrap; }}
.tab.active {{ background: var(--bg); color: var(--text); border-color: var(--border); }}
.tab:hover:not(.active) {{ color: var(--text); background: var(--surface2); }}
.panel {{ display: none; padding: 24px; max-width: 1400px; margin: 0 auto; }}
.panel.active {{ display: block; }}
.grid-2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }}
.grid-3 {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; }}
.grid-4 {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; }}
.card {{ background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 16px; }}
.card h3 {{ font-size: 12px; color: var(--muted); text-transform: uppercase; letter-spacing: .5px; margin-bottom: 8px; }}
.kpi-val {{ font-size: 26px; font-weight: 700; }}
.kpi-sub {{ font-size: 12px; color: var(--muted); margin-top: 4px; }}
.pos {{ color: var(--green); }} .neg {{ color: var(--red); }}
.section-title {{ font-size: 16px; font-weight: 600; margin: 24px 0 12px; padding-bottom: 8px; border-bottom: 1px solid var(--border); }}
.chart-wrap {{ position: relative; height: 280px; }}
.chart-wrap.tall {{ height: 360px; }}
.changes-grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; }}
.change-card {{ background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 14px; }}
.change-card h4 {{ font-size: 12px; color: var(--muted); text-transform: uppercase; margin-bottom: 8px; }}
.change-item {{ padding: 6px 0; border-bottom: 1px solid var(--border); font-size: 12px; }}
.change-item:last-child {{ border-bottom: none; }}
.change-item .ci-name {{ font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 100%; }}
.change-item .ci-sub {{ color: var(--muted); font-size: 11px; }}
.badge-new  {{ background: #1f6031; color: #3fb950; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: 700; }}
.badge-out  {{ background: #3d1a19; color: #f85149; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: 700; }}
.badge-up   {{ background: #1a3d26; color: #3fb950; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: 700; }}
.badge-down {{ background: #3d2e1a; color: #d29922; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: 700; }}
.placeholder {{ padding: 20px; text-align: center; color: var(--muted); font-size: 12px; font-style: italic; }}
table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
th {{ text-align: left; padding: 8px 10px; border-bottom: 2px solid var(--border); color: var(--muted); font-size: 11px; text-transform: uppercase; cursor: pointer; user-select: none; white-space: nowrap; }}
th:hover {{ color: var(--text); }}
td {{ padding: 7px 10px; border-bottom: 1px solid var(--border); }}
tr:hover td {{ background: var(--surface2); }}
.tbl-wrap {{ overflow-x: auto; max-height: 520px; overflow-y: auto; }}
.tbl-controls {{ display: flex; gap: 10px; margin-bottom: 12px; flex-wrap: wrap; align-items: center; }}
.tbl-controls input, .tbl-controls select {{ background: var(--surface2); border: 1px solid var(--border); color: var(--text); padding: 6px 10px; border-radius: 6px; font-size: 13px; }}
.tbl-controls input {{ flex: 1; min-width: 200px; }}
.pagination {{ display: flex; gap: 6px; justify-content: center; margin-top: 12px; flex-wrap: wrap; }}
.page-btn {{ background: var(--surface2); border: 1px solid var(--border); color: var(--text); padding: 4px 10px; border-radius: 4px; cursor: pointer; font-size: 12px; }}
.page-btn.active {{ background: var(--blue); color: #000; border-color: var(--blue); }}
.modal-overlay {{ display: none; position: fixed; inset: 0; background: #000c; z-index: 999; align-items: center; justify-content: center; }}
.modal-overlay.open {{ display: flex; }}
.modal {{ background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 24px; max-width: 680px; width: 95%; max-height: 85vh; overflow-y: auto; }}
.modal h2 {{ font-size: 18px; margin-bottom: 16px; }}
.modal-kpis {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-bottom: 16px; }}
.modal-kpi {{ background: var(--surface2); border-radius: 8px; padding: 12px; }}
.modal-kpi .lbl {{ font-size: 11px; color: var(--muted); margin-bottom: 4px; }}
.modal-kpi .val {{ font-size: 18px; font-weight: 700; }}
.modal-links {{ display: flex; gap: 8px; flex-wrap: wrap; margin-top: 12px; }}
.modal-links a {{ background: var(--surface2); border: 1px solid var(--border); padding: 6px 14px; border-radius: 6px; font-size: 12px; color: var(--text); }}
.modal-links a:hover {{ border-color: var(--blue); color: var(--blue); }}
.modal-close {{ float: right; cursor: pointer; color: var(--muted); font-size: 20px; line-height: 1; }}
.modal-close:hover {{ color: var(--text); }}
.fund-card-link {{ cursor: pointer; transition: transform .15s; }}
.fund-card-link:hover {{ transform: translateY(-2px); border-color: #58a6ff66; }}
.bar-row {{ display: flex; align-items: center; gap: 8px; margin-bottom: 6px; }}
.bar-label {{ width: 140px; font-size: 12px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; flex-shrink: 0; }}
.bar-track {{ flex: 1; height: 18px; background: var(--surface2); border-radius: 4px; overflow: hidden; }}
.bar-fill {{ height: 100%; border-radius: 4px; }}
.bar-val {{ width: 80px; text-align: right; font-size: 12px; color: var(--muted); flex-shrink: 0; }}
.updated {{ font-size: 11px; color: var(--muted); }}
/* Tooltips */
.tip {{ border-bottom: 1px dashed var(--muted); cursor: help; position: relative; display: inline; }}
.tip::after {{ content: attr(data-tip); position: absolute; bottom: calc(100% + 8px); left: 50%; transform: translateX(-50%); background: #1a1d2e; border: 1px solid var(--border); border-radius: 8px; padding: 8px 12px; font-size: 12px; color: var(--text); white-space: normal; min-width: 220px; max-width: 300px; line-height: 1.5; z-index: 500; opacity: 0; pointer-events: none; transition: opacity .2s; box-shadow: 0 4px 20px #00000066; }}
.tip:hover::after {{ opacity: 1; }}
@media (max-width: 768px) {{
  .grid-2, .grid-3, .grid-4, .changes-grid {{ grid-template-columns: 1fr; }}
  .header-kpis {{ display: none; }}
  .modal-kpis {{ grid-template-columns: 1fr 1fr; }}
}}
</style>
</head>
<body>

<div class="sticky-header">
  <div>
    <h1>☀️ Sunrise.app Dashboard</h1>
    <div class="updated">Stand: {updated_at}</div>
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
        if nav_ps and nav_ps_prev:
            d = (nav_ps - nav_ps_prev) / nav_ps_prev * 100
            day_chg = f'<div class="kpi-sub {'pos' if d>=0 else 'neg'}">{pl_sign(d)}{d:.2f}% heute</div>'

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

    # Gemeinsame Positionen
    html += '<div class="section-title">Gemeinsame Positionen (≥ 2 Fonds)</div>\n'
    html += '<div class="card">\n'
    html += _build_common_holdings_table(funds_data)
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
        day_chg_pct = ((nav_ps - nav_ps_prev) / nav_ps_prev * 100) if nav_ps and nav_ps_prev else None
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

        # NAV Sparkline
        if ph:
            html += '<div class="section-title">NAV Entwicklung</div>\n'
            html += f'<div class="card"><div class="chart-wrap"><canvas id="chart-spark-{fid}"></canvas></div></div>\n'

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
                        sub = f'{pl_sign(item["change_pct"])}{item["change_pct"]:.1f}% MV-Änderung'
                    elif item.get("mv_eur"):
                        sub = f'{fmt_eur(item["mv_eur"])} €'
                    html += f'<div class="change-item"><div class="ci-name">{item.get("name","—")}</div><div class="ci-sub">{sub}</div></div>\n'
            else:
                html += '<div class="placeholder">Noch keine Daten für heute<br>(wird ab dem 2. Tag befüllt)</div>\n'
            html += '</div>\n'
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
        for item in sorted(countries, key=lambda x: x["value"], reverse=True)[:15]:
            max_v = countries[0]["value"] if countries else 1
            pct = item["value"] / max_v * 100
            html += f'<div class="bar-row"><div class="bar-label" title="{item["label"]}">{item["label"]}</div><div class="bar-track"><div class="bar-fill" style="width:{pct:.1f}%;background:{f["color"]}"></div></div><div class="bar-val">{item["value"]:,.1f}%</div></div>\n'
        html += '</div></div>\n'
        # Währungen
        html += f'<div class="card"><h3>Währungen</h3><div class="chart-wrap"><canvas id="chart-ccy-{fid}"></canvas></div></div>\n'
        # Sektoren
        html += f'<div class="card"><h3>Sektoren</h3><div id="bars-sector-{fid}">\n'
        for item in sorted(sectors, key=lambda x: x["value"], reverse=True)[:10]:
            max_v = sectors[0]["value"] if sectors else 1
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
            html += f'''<tr onclick="showModal(this)" data-isin="{isin_s}" data-name="{h['name']}" data-country="{h.get('country','')}" data-sector="{h.get('sector','')}" data-currency="{h.get('currency','')}" data-mv="{h['mv_eur']:.2f}" data-pl="{pl_val:.2f}" data-nav-pct="{nav_pct:.3f}" data-cost="{h.get('cost') or ''}" data-price="{h.get('price') or ''}">
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

    # ── Calculator Panel ────────────────────────────────────────────────────
    html += '<div class="panel" id="panel-calc">\n'
    html += '<div class="section-title">🧮 Performance-Rechner</div>\n'
    html += '<div class="card" style="max-width:700px;margin-bottom:24px">\n'
    html += '''<div style="display:flex;gap:16px;align-items:flex-end;flex-wrap:wrap">
  <div>
    <label style="display:block;font-size:12px;color:var(--muted);margin-bottom:4px">Investitionsbetrag (€)</label>
    <input type="number" id="calc-amount" value="10000" min="1" style="background:var(--surface2);border:1px solid var(--border);color:var(--text);padding:8px 12px;border-radius:6px;font-size:16px;width:160px">
  </div>
  <div>
    <label style="display:block;font-size:12px;color:var(--muted);margin-bottom:4px">Zeitraum</label>
    <select id="calc-period" style="background:var(--surface2);border:1px solid var(--border);color:var(--text);padding:8px 12px;border-radius:6px;font-size:14px">
      <option value="ytd">YTD (01.01.–heute)</option>
      <option value="gy">Geschäftsjahr (01.10.–heute)</option>
      <option value="day">Gestern→Heute</option>
    </select>
  </div>
  <button onclick="calcPerf()" style="background:var(--blue);color:#000;border:none;padding:9px 20px;border-radius:6px;cursor:pointer;font-weight:700;font-size:14px">Berechnen</button>
</div>
'''
    html += '</div>\n'
    html += '<div class="grid-3" id="calc-results">\n'

    for f in funds_data:
        fid = f["id"]
        html += f'''<div class="card" id="calc-card-{fid}">
  <h3 style="margin-bottom:12px">{f['name']}</h3>
  <div id="calc-detail-{fid}" style="color:var(--muted);font-size:13px">→ Betrag eingeben und berechnen</div>
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
  <div class="modal-links" id="modal-links"></div>
</div>
</div>

'''

    # ── Scripts ─────────────────────────────────────────────────────────────
    html += f'<script>\nconst FUNDS_DATA = {data_json};\n'
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
}

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
Chart.defaults.color = '#8b949e';
Chart.defaults.borderColor = '#30363d';
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

// ── Per-Fund Charts ────────────────────────────────────────────────────────
FUNDS_DATA.forEach(fund => {
  const fid = fund.id;

  // Sparkline
  if (fund.price_history && fund.price_history.length) {
    const ctx = document.getElementById('chart-spark-' + fid);
    if (ctx) new Chart(ctx, {
      type: 'line',
      data: {
        labels: fund.price_history.map(p => p.label),
        datasets: [{
          label: 'Rücknahmepreis €',
          data: fund.price_history.map(p => p.price),
          borderColor: fund.color,
          backgroundColor: fund.color + '22',
          fill: true,
          tension: 0.3,
          pointRadius: 4,
          pointBackgroundColor: fund.color,
        }]
      },
      options: {responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}},
        scales:{y:{ticks:{callback:v=>v.toFixed(2)+' €'}}}}
    });
  }

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
function calcPerf() {
  const amount = parseFloat(document.getElementById('calc-amount').value) || 10000;
  const period = document.getElementById('calc-period').value;
  const datasets = [];

  FUNDS_DATA.forEach(fund => {
    const el = document.getElementById('calc-detail-' + fund.id);
    if (!el) return;

    const ph = fund.price_history || [];
    let startPrice = null, endPrice = null, startLabel = '';

    if (period === 'ytd') {
      const s = ph.find(p => p.date && p.date.startsWith(new Date().getFullYear() + '-01'));
      const e = ph[ph.length - 1];
      startPrice = s?.price; endPrice = e?.price; startLabel = s?.label || '01.01.';
    } else if (period === 'gy') {
      const fy = new Date().getFullYear();
      const fyY = new Date().getMonth() < 9 ? fy - 1 : fy;
      const s = ph.find(p => p.date && p.date.startsWith(fyY + '-10'));
      const e = ph[ph.length - 1];
      startPrice = s?.price; endPrice = e?.price; startLabel = s?.label || '01.10.';
    } else { // day
      if (ph.length >= 2) {
        startPrice = ph[ph.length-2].price;
        endPrice   = ph[ph.length-1].price;
        startLabel = ph[ph.length-2].label;
      }
    }

    if (!startPrice || !endPrice) {
      el.innerHTML = '<span style="color:var(--muted)">Keine Daten für diesen Zeitraum</span>';
      return;
    }

    const units   = amount / startPrice;
    const current = units * endPrice;
    const gain    = current - amount;
    const retPct  = (endPrice - startPrice) / startPrice * 100;
    const cls     = gain >= 0 ? 'pos' : 'neg';
    const s       = gain >= 0 ? '+' : '';

    el.innerHTML = `
      <table style="width:100%;font-size:13px">
        <tr><td style="color:var(--muted)">Einstiegskurs (${startLabel})</td><td style="text-align:right">${startPrice.toFixed(4)} €</td></tr>
        <tr><td style="color:var(--muted)">Heutiger Kurs</td><td style="text-align:right">${endPrice.toFixed(4)} €</td></tr>
        <tr><td style="color:var(--muted)">Anteile</td><td style="text-align:right">${units.toFixed(4)}</td></tr>
        <tr><td style="color:var(--muted)">Rendite</td><td style="text-align:right;font-weight:700" class="${cls}">${s}${retPct.toFixed(2)}%</td></tr>
        <tr><td style="color:var(--muted)">Aktueller Wert</td><td style="text-align:right;font-weight:700">${fmtEur(current)} €</td></tr>
        <tr><td style="color:var(--muted)">Gewinn/Verlust</td><td style="text-align:right;font-weight:700" class="${cls}">${s}${fmtEur(gain)} €</td></tr>
      </table>
    `;

    // Normalisierter Verlauf (nur für YTD/GJ sinnvoll)
    if (ph.length >= 2) {
      const base = ph[0].price;
      datasets.push({
        label: fund.name,
        data: ph.map(p => ({ x: p.label, y: (p.price - base) / base * 100 })),
        borderColor: fund.color,
        backgroundColor: 'transparent',
        tension: 0.3,
        pointRadius: 4,
      });
    }
  });

  // Update comparison chart
  const ctx = document.getElementById('chart-calc-perf');
  if (ctx) {
    if (calcChart) calcChart.destroy();
    const allLabels = [...new Set(FUNDS_DATA.flatMap(f => (f.price_history||[]).map(p => p.label)))];
    calcChart = new Chart(ctx, {
      type: 'line',
      data: { labels: allLabels, datasets },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { position: 'top' }, tooltip: { mode: 'index' } },
        scales: { y: { ticks: { callback: v => v.toFixed(2) + '%' } } }
      }
    });
  }
}
// Auto-calculate on load
setTimeout(calcPerf, 500);
</script>
</body>
</html>'''

    return html


def _build_common_holdings_table(funds_data):
    """Findet Positionen die in ≥2 Fonds sind."""
    from collections import defaultdict
    isin_funds = defaultdict(list)
    isin_info  = {}
    for f in funds_data:
        for h in f.get("holdings", []):
            isin = h.get("isin","")
            if isin and isin != "None":
                isin_funds[isin].append(f["id"])
                isin_info[isin] = h

    common = [(isin, funds, isin_info[isin]) for isin, funds in isin_funds.items() if len(funds) >= 2]
    common.sort(key=lambda x: sum(
        (f.get("holdings") or [{}])[-1].get("mv_eur", 0)
        for f in funds_data if f["id"] in x[1]
    ), reverse=True)

    if not common:
        return '<div class="placeholder">Keine gemeinsamen Positionen gefunden</div>'

    html = '<div class="tbl-wrap"><table><thead><tr><th>Name</th><th>ISIN</th><th>Land</th><th>Sektor</th>'
    for f in funds_data:
        html += f'<th>{f["id"]} MV</th>'
    html += '</tr></thead><tbody>\n'

    for isin, fund_ids, info in common[:30]:
        html += f'<tr><td>{info["name"][:40]}</td><td style="font-size:11px;color:var(--muted)">{isin}</td>'
        html += f'<td>{info.get("country","")}</td><td>{info.get("sector","")}</td>'
        for f in funds_data:
            h_match = next((h for h in f.get("holdings",[]) if h.get("isin")==isin), None)
            if h_match:
                mv = h_match["mv_eur"]
                pl = h_match.get("pl", 0) or 0
                cls = "pos" if pl > 0 else ("neg" if pl < 0 else "")
                html += f'<td><span class="{cls}">{mv/1e6:.2f} Mio. €</span></td>'
            else:
                html += '<td>—</td>'
        html += '</tr>\n'

    html += '</tbody></table></div>'
    return html


# ─── GitHub Commit ────────────────────────────────────────────────────────────
def git_push_file(token, repo, path, content_bytes, message, branch="main"):
    """Committed eine Datei direkt per GitHub API."""
    import base64
    b64_content = base64.b64encode(content_bytes).decode()

    # Bestehende SHA holen (falls Datei existiert)
    sha = None
    try:
        req = Request(
            f"https://api.github.com/repos/{repo}/contents/{path}?ref={branch}",
            headers={"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"},
        )
        with urlopen(req) as resp:
            existing = json.loads(resp.read())
            sha = existing.get("sha")
    except HTTPError as e:
        if e.code != 404:
            print(f"  ⚠️  Konnte SHA für {path} nicht lesen: {e}")

    # Commit vorbereiten
    body = {"message": message, "content": b64_content, "branch": branch}
    if sha:
        body["sha"] = sha

    req = Request(
        f"https://api.github.com/repos/{repo}/contents/{path}",
        data=json.dumps(body).encode(),
        method="PUT",
        headers={
            "Authorization": f"token {token}",
            "Accept":        "application/vnd.github.v3+json",
            "Content-Type":  "application/json",
        },
    )
    with urlopen(req) as resp:
        result = json.loads(resp.read())

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


# ─── Main ─────────────────────────────────────────────────────────────────────
def main():
    github_token = os.environ.get("GITHUB_TOKEN", "")
    github_repo  = os.environ.get("GITHUB_REPOSITORY", "")  # z.B. "user/repo"

    print("=" * 60)
    print(f"🚀 IQAM Dashboard Update – {date.today()}")
    print("=" * 60)

    # 1. Token holen
    access_token = get_access_token()

    # 2. Prev data laden
    prev_data = {}
    if github_token and github_repo:
        prev_data = load_prev_data(github_token, github_repo)

    # 3. Mails finden
    fund_mails = find_latest_emails(access_token)

    # 4. Pro Fund Excel laden + parsen
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

        # Vortags-Preis aus prev_data
        prev_fund = prev_data.get(fid, {})
        nav_ps_prev = prev_fund.get("nav_per_share") if prev_fund else None
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

    # 5. Dashboard generieren
    updated_at = datetime.now().strftime("%d.%m.%Y %H:%M UTC")
    print(f"\n🔨 Generiere Dashboard ({updated_at})…")
    html = generate_html(funds_data, updated_at)
    data_json = json.dumps(
        [{k: v for k, v in f.items() if k != "holdings"} | {"holdings": f.get("holdings", [])}
         for f in funds_data],
        ensure_ascii=False, indent=2
    )

    # 6. In GitHub pushen
    if github_token and github_repo:
        print(f"\n📤 Push zu {github_repo}…")
        today_str = date.today().isoformat()
        git_push_file(github_token, github_repo, "docs/index.html",
                     html.encode("utf-8"),
                     f"Dashboard update {today_str}")
        git_push_file(github_token, github_repo, "docs/dashboard_data.json",
                     data_json.encode("utf-8"),
                     f"Data update {today_str}")
        # Prev data für morgen speichern
        prev_save = {f["id"]: {k: v for k, v in f.items() if k not in ("changes",)}
                     for f in funds_data}
        git_push_file(github_token, github_repo, "docs/prev_data.json",
                     json.dumps(prev_save, ensure_ascii=False).encode("utf-8"),
                     f"Prev data {today_str}")
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
