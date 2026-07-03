#!/usr/bin/env python3
import imaplib, email, os, re, json, sys, subprocess, math, smtplib, zipfile
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from pathlib import Path
try:
    import openpyxl
except ImportError:
    sys.exit("Missing: pip install openpyxl")

GMAIL_USER         = "thamthorn.su@gmail.com"
GMAIL_APP_PASSWORD = os.environ.get("HH_GMAIL_APP_PASSWORD", "")
NETLIFY_TOKEN      = os.environ.get("NETLIFY_TOKEN", "")
NETLIFY_SITE_DOMAIN = "visionary-donut-6607b6"  # public *.netlify.app subdomain
NETLIFY_SITE_ID    = "09b3bf40-a3a9-48c7-921c-6e762b27bec8"  # API site UUID — the deploys endpoint 404s on the name
NOTIFY_EMAIL       = "thamthorn.suksawat@minor.com"
HUNGRYHUB_SENDER   = "bookings@hungryhub.net"
HUNGRYHUB_SUBJECT  = "Booking Export Ready"
DASHBOARD_HTML     = Path("/tmp/dashboard.html")
EXCEL_CACHE        = Path("/tmp/hh_export.xlsx")
MONTH_ABBR = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

def log(msg):
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}", flush=True)

def fetch_dashboard():
    log("Fetching current dashboard from Netlify...")
    r = subprocess.run(["curl","-sL",f"https://{NETLIFY_SITE_DOMAIN}.netlify.app","-o",str(DASHBOARD_HTML)], capture_output=True)
    if r.returncode != 0 or not DASHBOARD_HTML.exists() or DASHBOARD_HTML.stat().st_size < 1000:
        sys.exit("Failed to fetch dashboard HTML from Netlify")
    log(f"Dashboard fetched ({DASHBOARD_HTML.stat().st_size:,} bytes)")

def connect_gmail():
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(GMAIL_USER, GMAIL_APP_PASSWORD)
    return mail

def find_download_link(mail):
    ids = []
    for mailbox in ["inbox", '"[Gmail]/All Mail"']:
        mail.select(mailbox)
        _, data = mail.search(None, f'FROM "{HUNGRYHUB_SENDER}"')
        ids = data[0].split()
        if ids:
            break
    if not ids:
        log("No emails found from " + HUNGRYHUB_SENDER)
        return None, None
    for msg_id in reversed(ids[-10:]):
        _, msg_data = mail.fetch(msg_id, "(RFC822)")
        msg = email.message_from_bytes(msg_data[0][1])
        subject = msg.get("Subject", "")
        if HUNGRYHUB_SUBJECT.lower() not in subject.lower():
            continue
        email_date = msg.get("Date", "")
        log(f"Found email: '{subject}' [{email_date}]")
        for part in msg.walk():
            if part.get_content_type() in ("text/plain", "text/html"):
                body = part.get_payload(decode=True).decode("utf-8", errors="replace")
                match = re.search(r'https://images\.hungryhub\.com[^\s<>"\']+\.xlsx', body)
                if match:
                    url = match.group(0)
                    log(f"Download URL: {url}")
                    return url, email_date
    log("No download URL found")
    return None, None

def download_excel(url):
    r = subprocess.run(["curl","-L","-s","-A","Mozilla/5.0",url,"-o",str(EXCEL_CACHE),"-w","%{http_code}"], capture_output=True, text=True)
    http_code = r.stdout.strip()
    if http_code == "200" and EXCEL_CACHE.exists() and EXCEL_CACHE.stat().st_size > 1000:
        size = EXCEL_CACHE.stat().st_size
        log(f"Excel downloaded ({size:,} bytes)")
        return True, size
    log(f"Download failed — HTTP {http_code}")
    return False, 0

def _clean(val):
    if val is None: return ""
    s = str(val).strip()
    return "" if s in ("nan","None","") else s.replace("&amp;","&")

def _safe_float(val):
    if val is None: return None
    try:
        if isinstance(val, float) and math.isnan(val): return None
        return float(val)
    except: return None

def _safe_int(val):
    try:
        if isinstance(val, float) and math.isnan(val): return None
        return int(float(val))
    except: return None

def _format_date(val):
    if val is None: return None
    if isinstance(val, datetime): return val.strftime("%Y-%m-%d")
    try:
        from datetime import date
        if isinstance(val, date): return val.strftime("%Y-%m-%d")
    except: pass
    s = str(val).strip()
    if s in ("","nan","None"): return None
    for fmt in ("%Y-%m-%d","%d/%m/%Y","%m/%d/%Y","%Y/%m/%d"):
        try: return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except: pass
    return s

def _meal_period(t):
    if not t: return "Other"
    try:
        h = int(str(t).split(":")[0])
        return "Lunch" if h < 15 else "Dinner"
    except: return "Other"

def _package_type(pkg_type_raw, pkg_name):
    if pkg_type_raw:
        v = _clean(pkg_type_raw)
        if v: return v
    if not pkg_name: return None
    n = pkg_name.lower()
    if any(k in n for k in ["buffet","all you can","ayce"]): return "All You Can Eat"
    if any(k in n for k in ["party","premium set","set for"]): return "Party Pack"
    return None

def parse_excel(fp):
    log(f"Parsing {fp.name} …")
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        log("Excel is empty"); return []
    headers = [str(h).strip() if h else "" for h in rows[0]]
    log(f"  Columns: {headers}")
    records = []
    for raw in rows[1:]:
        row = dict(zip(headers, raw))
        pkg_name = _clean(row.get("Package-Name")) or None
        if pkg_name: pkg_name = pkg_name.strip()
        time_str = _clean(row.get("Dining-Time")) or None
        records.append({
            "id": _safe_int(row.get("ID")) or 0,
            "restaurant": _clean(row.get("Restaurant-Name")),
            "customer_name": _clean(row.get("Customer-Name")),
            "dining_date": _format_date(row.get("Dining-Date")),
            "dining_time": time_str,
            "meal_period": _meal_period(time_str),
            "party_size": _safe_int(row.get("Party-Size")) or 0,
            "status": _clean(row.get("Status")),
            "package_type": _package_type(row.get("Package-Type"), pkg_name),
            "package_name": pkg_name,
            "package_price": _safe_float(row.get("Package-Price")),
            "restaurant_revenue": _safe_float(row.get("Restaurant-Revenue")) or 0.0,
            "commission": _safe_float(row.get("Commision")) or 0.0,
            "created_date": _format_date(row.get("Created-Date")),
            "service_type": _clean(row.get("Service-Type")) or "Dine In",
            "special_request": _clean(row.get("Special-Request")) or None,
            "pre_payment": _safe_float(row.get("Pre-Payment")) or 0.0,
            "payment_type": _clean(row.get("Payment-Type")) or None,
        })
    wb.close()
    log(f"  Parsed {len(records)} records")
    return records

ZERO_STATUSES_GUARD_KPI = "a+(!ZERO_STATUSES.has(r.status)?(r.restaurant_revenue||0):0),0)"
ZERO_STATUSES_GUARD_MONTHLY = "(!ZERO_STATUSES.has(r.status)?(r.restaurant_revenue||0):0)"

def update_dashboard(records):
    content = DASHBOARD_HTML.read_text(encoding="utf-8")
    json_str = json.dumps(records, ensure_ascii=False)
    new_block = f"const RAW_DATA = {json_str};"
    start = content.find("const RAW_DATA = [")
    if start == -1:
        log("ERROR: Could not find 'const RAW_DATA' in dashboard HTML"); return None, None
    end = content.find("];", start) + 2
    content = content[:start] + new_block + content[end:]
    data_months = sorted({r["dining_date"][:7] for r in records if r.get("dining_date") and len(r["dining_date"]) >= 7})
    if data_months:
        months_js = json.dumps(data_months)
        mlabels_js = json.dumps([MONTH_ABBR[int(m[5:7])-1] for m in data_months])
        content = re.sub(r'const MONTHS\s*=\s*\[.*?\];', f'const MONTHS = {months_js};', content)
        content = re.sub(r'const MLABELS\s*=\s*\[.*?\];', f'const MLABELS = {mlabels_js};', content)
        log(f"  MONTHS updated: {data_months[0]} → {data_months[-1]}")
    if ZERO_STATUSES_GUARD_KPI not in content: log("WARNING: KPI revenue guard missing")
    if ZERO_STATUSES_GUARD_MONTHLY not in content: log("WARNING: Monthly revenue guard missing")
    total = len(records)
    # NOTE: pattern must match literal "</div>" (single slash) — a stray extra
    # escaped slash here silently breaks the badge-count replacement.
    content = re.sub(r'(<div class="badge">)\d+( bookings</div>)', rf'\g<1>{total}\2', content)
    now = datetime.now().strftime("%d %b %Y %H:%M")
    content = re.sub(r'(<div class="subtitle">).*?(<\/div>)', rf'\g<1>Last updated: {now}\2', content, count=1)
    assert content.count('"id":') == total
    DASHBOARD_HTML.write_text(content, encoding="utf-8")
    log(f"Dashboard updated — {total} bookings | {now}")
    rev_all = sum(r["restaurant_revenue"] for r in records)
    rev_excl = sum(r["restaurant_revenue"] for r in records if r["status"] not in ("Cancel","No Show","Cancel (modified)","Rejected"))
    log(f"  Revenue: ฿{rev_excl:,.0f} (excl. cancelled) vs ฿{rev_all:,.0f} (raw export)")
    return rev_all, rev_excl

def deploy_to_netlify():
    zip_path = Path("/tmp/deploy.zip")
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.write(DASHBOARD_HTML, "index.html")
    log(f"Uploading deploy.zip ({zip_path.stat().st_size:,} bytes)...")
    r = subprocess.run([
        "curl","-s","-X","POST",
        f"https://api.netlify.com/api/v1/sites/{NETLIFY_SITE_ID}/deploys",
        "-H", f"Authorization: Bearer {NETLIFY_TOKEN}",
        "-H", "Content-Type: application/zip",
        "--data-binary", f"@{zip_path}",
        "-w", "\n%{http_code}"
    ], capture_output=True, text=True)
    if r.returncode != 0:
        log(f"Netlify deploy curl failed (exit {r.returncode}): {r.stderr[:300]}")
        return None
    body, _, http_code = r.stdout.rpartition("\n")
    try:
        resp = json.loads(body)
    except Exception:
        log(f"Netlify deploy: could not parse response (HTTP {http_code}): {body[:500]}")
        return None
    url = resp.get("ssl_url") or resp.get("url")
    if not http_code.startswith("2") or not url:
        log(f"Netlify deploy FAILED (HTTP {http_code}): {body[:500]}")
        return None
    log(f"Deployed to Netlify → {url} (HTTP {http_code}, deploy id {resp.get('id')})")
    return url

def send_email(total, rev_excl, rev_all, dashboard_url, date_range, email_date, file_size):
    try:
        now_str = datetime.now().strftime("%d %b %Y %H:%M")
        rows_data = [
            ("Export email", email_date or "—"),
            ("Excel downloaded", f"{file_size:,} bytes" if file_size else "—"),
            ("Records parsed", f"{total} bookings"),
            ("Date range", date_range),
            ("Revenue (excl. cancelled)", f"฿{rev_excl:,.0f}"),
            ("Revenue (raw export)", f"฿{rev_all:,.0f}"),
            ("Deployed to Netlify", dashboard_url),
            ("Updated at", now_str),
        ]
        msg = MIMEMultipart("alternative")
        msg["Subject"] = f"HungryHub Dashboard Updated — {total} bookings | {datetime.now().strftime('%d %b %Y')}"
        msg["From"] = GMAIL_USER
        msg["To"] = NOTIFY_EMAIL
        text = "\n".join(f"{r[0]}: {r[1]}" for r in rows_data)
        msg.attach(MIMEText(text, "plain"))
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
            server.sendmail(GMAIL_USER, NOTIFY_EMAIL, msg.as_string())
        log(f"Notification email sent to {NOTIFY_EMAIL}")
    except Exception as e:
        log(f"Email failed: {e}")

def main():
    log("=" * 60)
    log("HungryHub Dashboard Updater (GitHub Actions) starting")
    if not GMAIL_APP_PASSWORD:
        sys.exit("ERROR: HH_GMAIL_APP_PASSWORD not set")
    if not NETLIFY_TOKEN:
        sys.exit("ERROR: NETLIFY_TOKEN not set")
    fetch_dashboard()
    log("Connecting to Gmail …")
    mail = connect_gmail()
    url, email_date = find_download_link(mail)
    mail.logout()
    if not url:
        log("No export link found — exiting"); sys.exit(0)
    ok, file_size = download_excel(url)
    if not ok:
        log("Download failed — exiting"); sys.exit(1)
    records = parse_excel(EXCEL_CACHE)
    if not records:
        log("No records — exiting"); sys.exit(0)
    rev_all, rev_excl = update_dashboard(records)
    deploy_url = deploy_to_netlify()
    if not deploy_url:
        log("Deploy failed — not sending success email"); sys.exit(1)
    dates = sorted(r["dining_date"] for r in records if r.get("dining_date"))
    date_range = f"{dates[0]} → {dates[-1]}" if dates else "N/A"
    send_email(len(records), rev_excl, rev_all, deploy_url, date_range, email_date, file_size)
    log("Done.")

if __name__ == "__main__":
    main()
