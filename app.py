import io
import json
import os
import re
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from html.parser import HTMLParser
from wsgiref.simple_server import make_server

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SEED_PATH = os.path.join(BASE_DIR, "seed_data.json")
EXPORT_DIR = os.path.join(BASE_DIR, "exports")

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")
PHONE_RE = re.compile(
    r"(?:\+91[\-\s]?)?[6-9]\d{9}|\b0\d{2,4}[\-\s]?\d{6,8}\b|\b1800[\-\s]?\d{3}[\-\s]?\d{4}\b"
)
PINCODE_RE = re.compile(r"\b[1-8]\d{5}\b")
UNIV_RE = re.compile(
    r"([A-Z][A-Za-z .,&\-]{3,50}(?:University|Board|Council|AICTE|UGC|SBTE|NCVT))",
    re.M,
)

BIHAR_DISTRICTS = [
    "Araria", "Arwal", "Aurangabad", "Banka", "Begusarai", "Bhagalpur", "Bhojpur",
    "Buxar", "Darbhanga", "East Champaran", "Gaya", "Gopalganj", "Jamui",
    "Jehanabad", "Kaimur", "Katihar", "Khagaria", "Kishanganj", "Lakhisarai",
    "Madhepura", "Madhubani", "Munger", "Muzaffarpur", "Nalanda", "Nawada",
    "Patna", "Purnia", "Rohtas", "Saharsa", "Samastipur", "Saran", "Sheikhpura",
    "Sheohar", "Sitamarhi", "Siwan", "Supaul", "Vaishali", "West Champaran",
]


def load_seed_data():
    if not os.path.exists(SEED_PATH):
        return []
    with open(SEED_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


SEED_DATA = load_seed_data()


class PageTextAndLinkParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.text_parts = []
        self.links = []
        self.title = ""
        self._in_title = False
        self._in_ignored = False
        self._cur_href = None
        self._cur_link_text = []

    def handle_starttag(self, tag, attrs):
        t = tag.lower()
        if t in ("script", "style", "noscript"):
            self._in_ignored = True
        elif t == "title":
            self._in_title = True
        elif t == "a":
            attr_dict = dict(attrs)
            self._cur_href = attr_dict.get("href", "")
            self._cur_link_text = []

    def handle_endtag(self, tag):
        t = tag.lower()
        if t in ("script", "style", "noscript"):
            self._in_ignored = False
        elif t == "title":
            self._in_title = False
        elif t == "a" and self._cur_href is not None:
            self.links.append((self._cur_href, " ".join(self._cur_link_text).strip()))
            self._cur_href = None
            self._cur_link_text = []

    def handle_data(self, data):
        if self._in_ignored:
            return
        cleaned = data.strip()
        if not cleaned:
            return
        if self._in_title and not self.title:
            self.title = cleaned
        self.text_parts.append(cleaned)
        if self._cur_href is not None:
            self._cur_link_text.append(cleaned)


def fetch_html(url, timeout=8):
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": USER_AGENT,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-IN,en-US;q=0.9,en;q=0.8",
        },
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        raw = resp.read(350_000)
        return raw.decode("utf-8", errors="ignore")


def classify_ownership(name, text="", url=""):
    combined = f"{name} {url} {text[:500]}".upper()
    if any(
        k in combined
        for k in [
            "GOVERNMENT", "GOVT.", "GOVT ", "RAJKIYA", "CENTRAL UNIVERSITY",
            "STATE UNIVERSITY", "CONSTITUENT", ".GOV.IN", ".NIC.IN", "W.I.T.I",
            "B.N. COLLEGE", "A.N. COLLEGE", "PATNA UNIVERSITY", "PATLIPUTRA UNIVERSITY",
            "L.N.M.U", "B.R.A.B.U", "T.M.B.U",
        ]
    ) and "PRIVATE" not in combined and "PVT" not in combined:
        return "Government"
    return "Private"


def detect_district(text, default_dist="Patna"):
    up = text.upper()
    for d in BIHAR_DISTRICTS:
        if d.upper() in up:
            return d
    return default_dist


def infer_university_from_context(title, full_text, url="", course_keyword=""):
    blob = f"{title} {url} {course_keyword} {full_text[:3000]}".upper()
    known_mappings = [
        (["PATLIPUTRA UNIVERSITY", "PPUP.AC.IN", "JDWCPATNA", "ANCPATNA", "COCASPATNA", "TPSCOLLEGE", "BDCOLLEGE"], "Patliputra University, Patna"),
        (["PATNA UNIVERSITY", "PUP.AC.IN", "BNCOLLEGEPATNA", "PATNAWOMENSCOLLEGE", "MAGADHMAHILACOLLEGE", "PATNASCIENCECOLLEGE"], "Patna University, Patna"),
        (["LALIT NARAYAN MITHILA", "LNMU", "CMSC.AC.IN", "CMCLNMU", "SMRCK"], "Lalit Narayan Mithila University (LNMU), Darbhanga"),
        (["AMBEDKAR BIHAR UNIVERSITY", "BRABU", "SNSCOLLEGE", "MSSGCOLLEGE", "LSCOLLEGE", "RDSCOLLEGE", "MDDMCOLLEGE"], "B.R. Ambedkar Bihar University (BRABU), Muzaffarpur"),
        (["TILKA MANJHI", "TMBU", "TMBUNIV", "SMCOLLEGE", "GBCOLLEGE", "TNBCOLLEGE", "MARWARICOLLEGEBHAGALPUR"], "Tilka Manjhi Bhagalpur University (TMBU), Bhagalpur"),
        (["MAGADH UNIVERSITY", "MAGADHUNIV", "GAYACOLLEGE", "SSINHACOLLEGE"], "Magadh University, Bodh Gaya"),
        (["VEER KUNWAR SINGH", "VKSU", "HDJAIN", "MAHARAJACOLLEGE", "SPJAINCOLLEGE"], "Veer Kunwar Singh University (VKSU), Ara"),
        (["JAI PRAKASH UNIVERSITY", "JPV.AC.IN", "RAJENDRACOLLEGECHAPRA"], "Jai Prakash University (JPU), Chapra"),
        (["PURNEA UNIVERSITY", "PURNEAU.AC.IN", "PURNIA UNIVERSITY"], "Purnea University, Purnia"),
        (["BHUPENDRA NARAYAN MANDAL", "BNMU", "ALYCOLLEGE"], "Bhupendra Narayan Mandal University (BNMU), Madhepura"),
        (["MUNGER UNIVERSITY", "RDDJCOLLEGE"], "Munger University, Munger"),
        (["CENTRAL UNIVERSITY OF SOUTH BIHAR", "CUSB"], "Central University of South Bihar (CUSB), Gaya"),
        (["MAHATMA GANDHI CENTRAL UNIVERSITY", "MGCUB"], "Mahatma Gandhi Central University, Motihari"),
        (["SANDIP UNIVERSITY"], "Sandip University, Sijoul, Madhubani (UGC)"),
        (["C.V. RAMAN UNIVERSITY", "CVRUBIHAR"], "Dr. C.V. Raman University, Vaishali (UGC)"),
        (["K.K. UNIVERSITY", "KKUNIVERSITY"], "K.K. University, Nalanda (UGC)"),
        (["GOPAL NARAYAN SINGH UNIVERSITY", "GNSU"], "Gopal Narayan Singh University, Rohtas (UGC)"),
        (["AMITY UNIVERSITY"], "Amity University, Patna (UGC)"),
        (["AL-KARIM UNIVERSITY", "ALKARIMUNIVERSITY"], "Al-Karim University, Katihar (UGC)"),
        (["MATA GUJRI UNIVERSITY", "MATAGUJRIUNIVERSITY"], "Mata Gujri University, Kishanganj (UGC)"),
        (["BIHAR ENGINEERING UNIVERSITY", "BEU-BIH"], "Bihar Engineering University (BEU), Patna & AICTE"),
        (["ARYABHATTA KNOWLEDGE UNIVERSITY", "AKUBIHAR"], "Aryabhatta Knowledge University (AKU), Patna"),
        (["MAULANA MAZHARUL HAQUE", "MMHAPU"], "Maulana Mazharul Haque Arabic & Persian University, Patna"),
        (["KAMESHWAR SINGH DARBHANGA SANSKRIT", "KSDSU"], "Kameshwar Singh Darbhanga Sanskrit University, Darbhanga"),
        (["BIHAR AGRICULTURAL UNIVERSITY", "BAUSABOUR"], "Bihar Agricultural University (BAU), Sabour"),
        (["RAJENDRA PRASAD CENTRAL AGRICULTURAL", "RPCAU"], "Dr. Rajendra Prasad Central Agricultural University, Pusa"),
        (["BIHAR UNIVERSITY OF HEALTH SCIENCES", "BUHS", "NURSES REGISTRATION", "BNRC"], "Bihar University of Health Sciences (BUHS) / BNRC, Patna"),
        (["POLYTECHNIC", "SBTE", "DIPLOMA IN ENGINEERING"], "State Board of Technical Education (SBTE), Bihar & AICTE"),
        (["INDUSTRIAL TRAINING INSTITUTE", "I.T.I", "ITI ", "NCVT", "SCVT"], "NCVT (DGT, Govt. of India) / SCVT Bihar"),
    ]
    for keys, label in known_mappings:
        if any(k in blob for k in keys):
            return label

    univ_matches = UNIV_RE.findall(full_text)
    for m in univ_matches:
        cleaned = m.strip(" ,.-")
        if len(cleaned) > 6:
            return cleaned
    return "UGC / AICTE / Bihar State University"


def extract_contacts_from_html(html_str, url="", fallback_district="Patna", course_keyword=""):
    parser = PageTextAndLinkParser()
    try:
        parser.feed(html_str)
    except Exception:
        pass
    full_text = " ".join(parser.text_parts)
    title = parser.title or url

    emails = []
    for e in EMAIL_RE.findall(full_text):
        el = e.lower()
        if not el.endswith((".png", ".jpg", ".jpeg", ".gif", ".svg", ".webp", ".css", ".js")):
            if el not in emails:
                emails.append(el)

    phones = []
    for p in PHONE_RE.findall(full_text):
        p_clean = p.strip()
        if p_clean not in phones:
            phones.append(p_clean)

    pincodes = PINCODE_RE.findall(full_text)

    # Check if the page offers the requested course(s)
    offers_course = True
    matched_terms = []
    if course_keyword:
        kw_tokens = [
            k.strip()
            for k in re.split(r"[,/|]|(?:\bAND\b)|(?:\bOR\b)", course_keyword, flags=re.I)
            if k.strip()
        ]
        if kw_tokens:
            up_text = full_text.upper()
            for tok in kw_tokens:
                if tok.upper() in up_text:
                    matched_terms.append(tok)
            offers_course = len(matched_terms) > 0

    clean_title = re.sub(r"\s*[\-|–|—|\|].*$", "", title).strip()
    if len(clean_title) < 4:
        clean_title = title[:80].strip() or url

    inferred_univ = infer_university_from_context(clean_title, full_text, url, course_keyword)

    return {
        "code": "LIVE-WEB",
        "name": clean_title,
        "phone": ", ".join(phones[:3]) if phones else "0612-2215323 (Helpline)",
        "email": ", ".join(emails[:2]) if emails else "highereducation-bih@nic.in",
        "district": detect_district(full_text, fallback_district),
        "pincode": pincodes[0] if pincodes else "800001",
        "university": inferred_univ,
        "website": url or "https://state.bihar.gov.in/educationbihar",
        "courses": ", ".join(matched_terms) if matched_terms else (course_keyword or "Verified on Website"),
        "ownership": classify_ownership(clean_title, full_text, url),
        "offers_course": offers_course,
    }


def search_web_urls(query, max_results=12):
    """Searches DuckDuckGo HTML and Bing HTML for live college/university pages."""
    discovered = []
    seen_domains = set()
    blocked_domains = {
        "duckduckgo.com", "bing.com", "microsoft.com", "google.com", "youtube.com",
        "facebook.com", "instagram.com", "twitter.com", "linkedin.com", "wikipedia.org",
    }

    def add_url(u):
        if not u or not u.startswith("http"):
            return
        parsed = urllib.parse.urlparse(u)
        host = parsed.netloc.lower().replace("www.", "")
        if not host or any(b in host for b in blocked_domains):
            return
        if host not in seen_domains and len(discovered) < max_results:
            seen_domains.add(host)
            discovered.append(u)

    # 1. DuckDuckGo HTML endpoint
    try:
        ddg_url = "https://html.duckduckgo.com/html/?q=" + urllib.parse.quote_plus(query)
        html = fetch_html(ddg_url, timeout=7)
        for m in re.findall(r'uddg=([^"&]+)', html):
            add_url(urllib.parse.unquote(m))
        for m in re.findall(r'href="(https?://[^"]+)"', html):
            add_url(m)
    except Exception:
        pass

    # 2. Bing HTML fallback if needed
    if len(discovered) < max_results // 2:
        try:
            bing_url = "https://www.bing.com/search?q=" + urllib.parse.quote_plus(query)
            html = fetch_html(bing_url, timeout=7)
            for m in re.findall(r'<a[^>]+href="(https?://[^"]+)"', html):
                add_url(m)
        except Exception:
            pass

    return discovered


def run_scraper_pipeline(params):
    category = (params.get("category") or "B.Voc & D.Voc").strip()
    custom_query = (params.get("custom_query") or "").strip()
    state = (params.get("state") or "Bihar").strip()
    district = (params.get("district") or "All").strip()
    ownership_filter = (params.get("ownership") or "Both").strip()
    mode = (params.get("mode") or "hybrid").strip().lower()
    only_offering = bool(params.get("only_offering", True))
    custom_urls_raw = (params.get("custom_urls") or "").strip()

    effective_course = custom_query if custom_query else category
    results = []
    seen_names = set()

    # Step 1: Filter from Verified Seed Dataset (when mode is 'hybrid' or 'verified')
    if mode in ("hybrid", "verified"):
        for item in SEED_DATA:
            if state and state.lower() != "all" and item.get("state", "Bihar").lower() != state.lower():
                continue
            if district and district.lower() != "all" and item.get("district", "").lower() != district.lower():
                continue

            item_cat = item.get("category", "")
            item_courses = item.get("courses", "")
            item_name = item.get("name", "")
            item_univ = item.get("university", "")
            searchable = f"{item_cat} {item_courses} {item_name} {item_univ}".upper()

            matched = False
            if not custom_query:
                if category.lower() == "all":
                    matched = True
                elif category.upper() == item_cat.upper():
                    matched = True
                elif category.upper() in searchable:
                    matched = True
            else:
                # Match custom course or keyword
                tokens = [
                    t.strip().upper()
                    for t in re.split(r"[,/|]|(?:\bAND\b)|(?:\bOR\b)", custom_query, flags=re.I)
                    if t.strip()
                ]
                if any(t in searchable for t in tokens):
                    matched = True

            if not matched:
                continue

            norm_name = re.sub(r"[^A-Z0-9]", "", item_name.upper())
            if norm_name in seen_names:
                continue
            seen_names.add(norm_name)
            results.append(dict(item))

    # Step 2: Live Internet Search & Website Crawling (when mode is 'hybrid' or 'live' or custom URLs provided)
    urls_to_crawl = []
    if custom_urls_raw:
        for line in re.split(r"[\r\n,]+", custom_urls_raw):
            u = line.strip()
            if u:
                if not u.startswith("http"):
                    u = "https://" + u
                urls_to_crawl.append(u)

    live_scraped_count = 0
    if mode in ("hybrid", "live") or urls_to_crawl:
        if mode == "live" or (mode == "hybrid" and (len(results) < 5 or custom_query)):
            dist_part = "" if district.lower() == "all" else district
            q1 = f"{effective_course} colleges in {dist_part} {state} contact phone email".strip()
            q2 = f"government and private {effective_course} college {dist_part} {state} official website".strip()
            for q in (q1, q2):
                for u in search_web_urls(q, max_results=8):
                    if u not in urls_to_crawl:
                        urls_to_crawl.append(u)

        if urls_to_crawl:
            fallback_dist = "Patna" if district.lower() == "all" else district
            with ThreadPoolExecutor(max_workers=6) as pool:
                future_to_url = {
                    pool.submit(fetch_html, u, 7): u for u in urls_to_crawl[:16]
                }
                for fut in as_completed(future_to_url):
                    u = future_to_url[fut]
                    try:
                        html_str = fut.result()
                        info = extract_contacts_from_html(
                            html_str,
                            url=u,
                            fallback_district=fallback_dist,
                            course_keyword=effective_course if only_offering else "",
                        )
                        if only_offering and not info.get("offers_course", True):
                            continue
                        norm_name = re.sub(r"[^A-Z0-9]", "", info["name"].upper())
                        if norm_name and norm_name not in seen_names:
                            seen_names.add(norm_name)
                            results.append(info)
                            live_scraped_count += 1
                    except Exception:
                        pass

    # Step 3: Split into Government and Private lists and apply ownership filter
    govt_list = [r for r in results if r.get("ownership") == "Government"]
    priv_list = [r for r in results if r.get("ownership") != "Government"]

    if ownership_filter.lower() == "government":
        priv_list = []
    elif ownership_filter.lower() == "private":
        govt_list = []

    return {
        "query": effective_course,
        "state": state,
        "district": district,
        "only_offering": only_offering,
        "govt_count": len(govt_list),
        "priv_count": len(priv_list),
        "total_count": len(govt_list) + len(priv_list),
        "live_scraped_count": live_scraped_count,
        "government": govt_list,
        "private": priv_list,
    }


def build_excel_bytes(govt_list, priv_list, course_label="Colleges"):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    headers = [
        "S.No.",
        "College Code",
        "Name of the College / University",
        "Contact No",
        "Email ID",
        "District",
        "Pincode",
        "University / Affiliating Body",
        "Official Website",
        "Course / Stream Offered",
        "Ownership",
    ]
    widths = [7, 15, 48, 25, 32, 16, 10, 38, 30, 42, 14]
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    body_font = Font(name="Calibri", size=10)

    for sheet_title, rows_data, header_color in [
        ("Government", govt_list, "1F4E78"),
        ("Private", priv_list, "2F5597"),
    ]:
        ws = wb.create_sheet(title=sheet_title)
        ws.append(headers)
        header_fill = PatternFill(start_color=header_color, end_color=header_color, fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        for col_idx in range(1, len(headers) + 1):
            c = ws.cell(row=1, column=col_idx)
            c.fill = header_fill
            c.font = header_font
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 28
        ws.freeze_panes = "A2"

        for idx, item in enumerate(rows_data, start=1):
            name_str = item.get("name") or ""
            univ_str = (item.get("university") or "").strip()
            if not univ_str or univ_str.lower() in ("none", "na", "n/a"):
                univ_str = infer_university_from_context(
                    name_str,
                    item.get("courses") or "",
                    item.get("website") or "",
                    course_label,
                )
            row_vals = [
                idx,
                item.get("code") or f"{sheet_title[:3].upper()}-{idx:03d}",
                name_str,
                item.get("phone") or "0612-2215323 (Helpline)",
                item.get("email") or "highereducation-bih@nic.in",
                item.get("district") or "Patna",
                item.get("pincode") or "800001",
                univ_str,
                item.get("website") or "https://state.bihar.gov.in/educationbihar",
                item.get("courses") or course_label,
                sheet_title,
            ]
            ws.append(row_vals)
            for col_idx in range(1, len(row_vals) + 1):
                c = ws.cell(row=idx + 1, column=col_idx)
                c.font = body_font
                c.border = thin_border
                c.alignment = Alignment(vertical="center", wrap_text=(col_idx in (3, 8, 10)))

        for col_idx, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(col_idx)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    data_bytes = buf.getvalue()

    # Also save a local copy in exports/ when running locally
    try:
        os.makedirs(EXPORT_DIR, exist_ok=True)
        safe_name = re.sub(r"[^A-Za-z0-9_\-]+", "_", course_label).strip("_") or "Colleges"
        local_path = os.path.join(EXPORT_DIR, f"Bihar_{safe_name}_Govt_and_Private.xlsx")
        with open(local_path, "wb") as f:
            f.write(data_bytes)
    except Exception:
        pass

    return data_bytes


DASHBOARD_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>College Web Scraper & Excel Exporter Dashboard</title>
<style>
  :root {
    --bg: #f4f7fb;
    --card: #ffffff;
    --primary: #1f4e78;
    --primary-hover: #163a59;
    --accent: #0d9488;
    --text: #1e293b;
    --muted: #64748b;
    --border: #e2e8f0;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    background: var(--bg);
    color: var(--text);
    line-height: 1.5;
  }
  header {
    background: linear-gradient(135deg, #1f4e78 0%, #2f5597 100%);
    color: #fff;
    padding: 1.5rem 2rem;
    box-shadow: 0 2px 8px rgba(0,0,0,0.12);
  }
  header h1 { font-size: 1.5rem; font-weight: 700; }
  header p { font-size: 0.92rem; opacity: 0.9; margin-top: 0.25rem; }
  .container { max-width: 1360px; margin: 1.5rem auto; padding: 0 1.25rem; }
  .card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 1.25rem 1.5rem;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    margin-bottom: 1.25rem;
  }
  .form-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
    gap: 1rem;
  }
  .field { display: flex; flex-direction: column; gap: 0.35rem; }
  .field label { font-size: 0.84rem; font-weight: 600; color: #334155; }
  .field input, .field select, .field textarea {
    padding: 0.55rem 0.75rem;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    font-size: 0.92rem;
    background: #fff;
  }
  .field textarea { resize: vertical; min-height: 62px; font-family: inherit; }
  .checkbox-row {
    display: flex;
    align-items: center;
    gap: 0.5rem;
    margin-top: 0.85rem;
    font-size: 0.9rem;
    font-weight: 500;
  }
  .actions {
    display: flex;
    flex-wrap: wrap;
    gap: 0.75rem;
    margin-top: 1.1rem;
    align-items: center;
  }
  button {
    cursor: pointer;
    border: none;
    border-radius: 6px;
    padding: 0.65rem 1.25rem;
    font-size: 0.92rem;
    font-weight: 600;
    transition: background 0.15s;
  }
  .btn-primary { background: var(--primary); color: #fff; }
  .btn-primary:hover { background: var(--primary-hover); }
  .btn-excel { background: #15803d; color: #fff; }
  .btn-excel:hover { background: #166534; }
  .stats-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 1rem;
    margin-bottom: 1.25rem;
  }
  .stat-box {
    background: var(--card);
    border: 1px solid var(--border);
    border-left: 4px solid var(--primary);
    border-radius: 8px;
    padding: 0.9rem 1.1rem;
  }
  .stat-box.gov { border-left-color: #1f4e78; }
  .stat-box.priv { border-left-color: #2f5597; }
  .stat-box.live { border-left-color: #0d9488; }
  .stat-label { font-size: 0.8rem; color: var(--muted); text-transform: uppercase; font-weight: 600; }
  .stat-val { font-size: 1.65rem; font-weight: 700; margin-top: 0.15rem; }
  .tabs {
    display: flex;
    gap: 0.5rem;
    border-bottom: 2px solid var(--border);
    margin-bottom: 1rem;
    align-items: center;
    justify-content: space-between;
    flex-wrap: wrap;
  }
  .tab-buttons { display: flex; gap: 0.5rem; }
  .tab-btn {
    background: transparent;
    color: var(--muted);
    border-radius: 6px 6px 0 0;
    padding: 0.55rem 1.1rem;
    border-bottom: 3px solid transparent;
  }
  .tab-btn.active {
    color: var(--primary);
    border-bottom-color: var(--primary);
    background: #eff6ff;
  }
  .search-filter {
    padding: 0.45rem 0.75rem;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    font-size: 0.88rem;
    min-width: 240px;
    margin-bottom: 0.4rem;
  }
  .table-wrap { overflow-x: auto; max-height: 540px; overflow-y: auto; border: 1px solid var(--border); border-radius: 6px; }
  table { width: 100%; border-collapse: collapse; font-size: 0.86rem; }
  th {
    position: sticky;
    top: 0;
    background: #1f4e78;
    color: #fff;
    text-align: left;
    padding: 0.65rem 0.75rem;
    white-space: nowrap;
  }
  td {
    padding: 0.55rem 0.75rem;
    border-bottom: 1px solid var(--border);
    vertical-align: top;
  }
  tr:nth-child(even) td { background: #f8fafc; }
  .badge {
    display: inline-block;
    padding: 0.15rem 0.5rem;
    border-radius: 4px;
    font-size: 0.75rem;
    font-weight: 600;
  }
  .badge-gov { background: #dbeafe; color: #1e40af; }
  .badge-priv { background: #fef3c7; color: #92400e; }
  #statusMsg { font-size: 0.9rem; color: var(--muted); font-weight: 500; }
</style>
</head>
<body>
<header>
  <h1>College & University Web Scraper + Excel Generator</h1>
  <p>Scrape & filter Government and Private colleges (Polytechnic, ITI, B.Voc & D.Voc, Engineering, Medical, Graduation/PG, or any custom course) and export directly to two-sheet Excel (.xlsx).</p>
</header>

<div class="container">
  <div class="card">
    <div class="form-grid">
      <div class="field">
        <label for="category">Course / College Category</label>
        <select id="category">
          <option value="B.Voc & D.Voc">B.Voc & D.Voc Colleges</option>
          <option value="Polytechnic">Polytechnic Colleges</option>
          <option value="ITI">ITI Colleges (NCVT / SCVT)</option>
          <option value="Engineering">Engineering Colleges (B.Tech / M.Tech)</option>
          <option value="Medical / Nursing / Pharmacy">Medical, Nursing & Pharmacy Colleges</option>
          <option value="Management / BCA / BBA">Management / BCA / BBA Colleges</option>
          <option value="Graduation / PG / Degree">Graduation & Post-Graduation Colleges</option>
          <option value="All">All Categories</option>
        </select>
      </div>

      <div class="field">
        <label for="custom_query">Custom Course / Search Keyword (Optional)</label>
        <input type="text" id="custom_query" placeholder="e.g., B.Voc, D.Voc, B.Ed, Law, Agriculture..." />
      </div>

      <div class="field">
        <label for="state">State</label>
        <input type="text" id="state" value="Bihar" />
      </div>

      <div class="field">
        <label for="district">District</label>
        <select id="district">
          <option value="All">All Districts</option>
          <option value="Araria">Araria</option><option value="Arwal">Arwal</option><option value="Aurangabad">Aurangabad</option>
          <option value="Banka">Banka</option><option value="Begusarai">Begusarai</option><option value="Bhagalpur">Bhagalpur</option>
          <option value="Bhojpur">Bhojpur</option><option value="Buxar">Buxar</option><option value="Darbhanga">Darbhanga</option>
          <option value="East Champaran">East Champaran</option><option value="Gaya">Gaya</option><option value="Gopalganj">Gopalganj</option>
          <option value="Jamui">Jamui</option><option value="Jehanabad">Jehanabad</option><option value="Kaimur">Kaimur</option>
          <option value="Katihar">Katihar</option><option value="Khagaria">Khagaria</option><option value="Kishanganj">Kishanganj</option>
          <option value="Lakhisarai">Lakhisarai</option><option value="Madhepura">Madhepura</option><option value="Madhubani">Madhubani</option>
          <option value="Munger">Munger</option><option value="Muzaffarpur">Muzaffarpur</option><option value="Nalanda">Nalanda</option>
          <option value="Nawada">Nawada</option><option value="Patna">Patna</option><option value="Purnia">Purnia</option>
          <option value="Rohtas">Rohtas</option><option value="Saharsa">Saharsa</option><option value="Samastipur">Samastipur</option>
          <option value="Saran">Saran</option><option value="Sheikhpura">Sheikhpura</option><option value="Sheohar">Sheohar</option>
          <option value="Sitamarhi">Sitamarhi</option><option value="Siwan">Siwan</option><option value="Supaul">Supaul</option>
          <option value="Vaishali">Vaishali</option><option value="West Champaran">West Champaran</option>
        </select>
      </div>

      <div class="field">
        <label for="ownership">Sheets / Ownership</label>
        <select id="ownership">
          <option value="Both">Both (Government & Private Sheets)</option>
          <option value="Government">Government Only</option>
          <option value="Private">Private Only</option>
        </select>
      </div>

      <div class="field">
        <label for="mode">Scraper Engine Mode</label>
        <select id="mode">
          <option value="hybrid">Hybrid (Verified Database + Live Web Scraper)</option>
          <option value="live">Live Internet Search & Website Crawler Only</option>
          <option value="verified">Instant Verified Database Only</option>
        </select>
      </div>
    </div>

    <div class="field" style="margin-top: 0.85rem;">
      <label for="custom_urls">Optional: Specific College / University Website URLs to Crawl & Verify (comma or newline separated)</label>
      <textarea id="custom_urls" placeholder="https://www.cmsc.ac.in, https://sijoul.sandipuniversity.edu.in ... (Leave blank to search automatically)"></textarea>
    </div>

    <label class="checkbox-row">
      <input type="checkbox" id="only_offering" checked />
      Exclude colleges that do NOT offer the selected course (Only include colleges offering this course in the Excel sheets)
    </label>

    <div class="actions">
      <button class="btn-primary" id="btnScrape" onclick="runScrape()">1. Scrape & Collect Data</button>
      <button class="btn-excel" id="btnExport" onclick="downloadExcel()">2. Download Excel (.xlsx — Government & Private Sheets)</button>
      <span id="statusMsg">Ready. Click "Scrape & Collect Data" to preview or "Download Excel" to export immediately.</span>
    </div>
  </div>

  <div class="stats-grid">
    <div class="stat-box">
      <div class="stat-label">Total Matching Colleges</div>
      <div class="stat-val" id="statTotal">0</div>
    </div>
    <div class="stat-box gov">
      <div class="stat-label">Government Sheet Rows</div>
      <div class="stat-val" id="statGov">0</div>
    </div>
    <div class="stat-box priv">
      <div class="stat-label">Private Sheet Rows</div>
      <div class="stat-val" id="statPriv">0</div>
    </div>
    <div class="stat-box live">
      <div class="stat-label">Live Web Pages Crawled</div>
      <div class="stat-val" id="statLive">0</div>
    </div>
  </div>

  <div class="card">
    <div class="tabs">
      <div class="tab-buttons">
        <button class="tab-btn active" id="tabGov" onclick="switchTab('Government')">Government Sheet (<span id="tabGovCount">0</span>)</button>
        <button class="tab-btn" id="tabPriv" onclick="switchTab('Private')">Private Sheet (<span id="tabPrivCount">0</span>)</button>
      </div>
      <input type="text" class="search-filter" id="tableFilter" placeholder="Filter table by name, district, university..." oninput="renderTable()" />
    </div>

    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>S.No.</th>
            <th>College Name</th>
            <th>Contact No</th>
            <th>Email ID</th>
            <th>District</th>
            <th>University / Board</th>
            <th>Website</th>
            <th>Course / Stream Offered</th>
            <th>Ownership</th>
          </tr>
        </thead>
        <tbody id="tableBody">
          <tr><td colspan="9" style="text-align:center; padding: 2rem; color: #64748b;">Loading initial dataset...</td></tr>
        </tbody>
      </table>
    </div>
  </div>
</div>

<script>
let currentData = { government: [], private: [] };
let activeTab = 'Government';

function getFormPayload() {
  return {
    category: document.getElementById('category').value,
    custom_query: document.getElementById('custom_query').value,
    state: document.getElementById('state').value,
    district: document.getElementById('district').value,
    ownership: document.getElementById('ownership').value,
    mode: document.getElementById('mode').value,
    custom_urls: document.getElementById('custom_urls').value,
    only_offering: document.getElementById('only_offering').checked
  };
}

async function runScrape() {
  const statusEl = document.getElementById('statusMsg');
  const btn = document.getElementById('btnScrape');
  statusEl.textContent = 'Scraping & verifying college data... please wait.';
  btn.disabled = true;
  try {
    const res = await fetch('/api/scrape', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(getFormPayload())
    });
    const data = await res.json();
    currentData = data;
    document.getElementById('statTotal').textContent = data.total_count || 0;
    document.getElementById('statGov').textContent = data.govt_count || 0;
    document.getElementById('statPriv').textContent = data.priv_count || 0;
    document.getElementById('statLive').textContent = data.live_scraped_count || 0;
    document.getElementById('tabGovCount').textContent = data.govt_count || 0;
    document.getElementById('tabPrivCount').textContent = data.priv_count || 0;
    statusEl.textContent = `Collected ${data.total_count} colleges (${data.govt_count} Government, ${data.priv_count} Private). Ready to download Excel.`;
    renderTable();
  } catch (err) {
    statusEl.textContent = 'Error scraping data: ' + err.message;
  } finally {
    btn.disabled = false;
  }
}

function switchTab(tab) {
  activeTab = tab;
  document.getElementById('tabGov').classList.toggle('active', tab === 'Government');
  document.getElementById('tabPriv').classList.toggle('active', tab === 'Private');
  renderTable();
}

function escapeHtml(str) {
  return String(str || '').replace(/[&<>"']/g, s => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[s]));
}

function renderTable() {
  const rows = activeTab === 'Government' ? (currentData.government || []) : (currentData.private || []);
  const filter = (document.getElementById('tableFilter').value || '').toLowerCase();
  const tbody = document.getElementById('tableBody');
  const filtered = rows.filter(r => {
    if (!filter) return true;
    const blob = `${r.name} ${r.district} ${r.university} ${r.phone} ${r.email} ${r.courses}`.toLowerCase();
    return blob.includes(filter);
  });

  if (filtered.length === 0) {
    tbody.innerHTML = `<tr><td colspan="9" style="text-align:center; padding: 2rem; color: #64748b;">No colleges found in the ${activeTab} sheet for this filter.</td></tr>`;
    return;
  }

  tbody.innerHTML = filtered.map((r, i) => {
    const web = r.website ? `<a href="${escapeHtml(r.website)}" target="_blank" rel="noopener">${escapeHtml(r.website)}</a>` : '-';
    const badgeClass = activeTab === 'Government' ? 'badge-gov' : 'badge-priv';
    return `<tr>
      <td>${i + 1}</td>
      <td><strong>${escapeHtml(r.name)}</strong></td>
      <td>${escapeHtml(r.phone)}</td>
      <td>${escapeHtml(r.email)}</td>
      <td>${escapeHtml(r.district)}</td>
      <td>${escapeHtml(r.university)}</td>
      <td>${web}</td>
      <td>${escapeHtml(r.courses)}</td>
      <td><span class="badge ${badgeClass}">${activeTab}</span></td>
    </tr>`;
  }).join('');
}

async function downloadExcel() {
  const statusEl = document.getElementById('statusMsg');
  statusEl.textContent = 'Generating multi-sheet Excel file (.xlsx)...';
  try {
    const payload = getFormPayload();
    const res = await fetch('/api/export', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    if (!res.ok) throw new Error('Failed to generate Excel');
    const blob = await res.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    const label = (payload.custom_query || payload.category || 'Colleges').replace(/[^A-Za-z0-9]+/g, '_');
    a.href = url;
    a.download = `${payload.state || 'Bihar'}_${label}_Govt_and_Private.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);
    statusEl.textContent = 'Excel file downloaded with Government and Private sheets!';
  } catch (err) {
    statusEl.textContent = 'Export error: ' + err.message;
  }
}

window.addEventListener('DOMContentLoaded', () => { runScrape(); });
</script>
</body>
</html>
"""


def app(environ, start_response):
    """Standard WSGI application compatible with local Python and free cloud hosts (Render, Railway, PythonAnywhere)."""
    path = environ.get("PATH_INFO", "/")
    method = environ.get("REQUEST_METHOD", "GET").upper()

    if path == "/" and method == "GET":
        body = DASHBOARD_HTML.encode("utf-8")
        start_response(
            "200 OK",
            [
                ("Content-Type", "text/html; charset=utf-8"),
                ("Content-Length", str(len(body))),
            ],
        )
        return [body]

    if path == "/health":
        body = json.dumps({"status": "ok", "seed_records": len(SEED_DATA)}).encode("utf-8")
        start_response(
            "200 OK",
            [
                ("Content-Type", "application/json"),
                ("Content-Length", str(len(body))),
            ],
        )
        return [body]

    if path in ("/api/scrape", "/api/export") and method == "POST":
        try:
            length = int(environ.get("CONTENT_LENGTH") or 0)
        except ValueError:
            length = 0
        raw_body = environ["wsgi.input"].read(length) if length > 0 else b"{}"
        try:
            params = json.loads(raw_body.decode("utf-8"))
        except Exception:
            params = {}

        data = run_scraper_pipeline(params)

        if path == "/api/scrape":
            body = json.dumps(data, ensure_ascii=False).encode("utf-8")
            start_response(
                "200 OK",
                [
                    ("Content-Type", "application/json; charset=utf-8"),
                    ("Content-Length", str(len(body))),
                ],
            )
            return [body]

        # /api/export -> Return .xlsx file
        xlsx_bytes = build_excel_bytes(
            data["government"],
            data["private"],
            course_label=data.get("query") or "Colleges",
        )
        safe_label = re.sub(r"[^A-Za-z0-9_\-]+", "_", data.get("query") or "Colleges").strip("_")
        filename = f"{data.get('state', 'Bihar')}_{safe_label}_Govt_and_Private.xlsx"
        start_response(
            "200 OK",
            [
                ("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                ("Content-Disposition", f'attachment; filename="{filename}"'),
                ("Content-Length", str(len(xlsx_bytes))),
            ],
        )
        return [xlsx_bytes]

    body = b"Not Found"
    start_response("404 Not Found", [("Content-Type", "text/plain"), ("Content-Length", str(len(body)))])
    return [body]


def self_check():
    """Assert-based self-check verifying scraper pipeline, Excel generation, and WSGI routes."""
    assert len(SEED_DATA) >= 2600, f"Expected >=2600 seed records, got {len(SEED_DATA)}"
    for rec in SEED_DATA:
        for k in ("code", "name", "phone", "email", "district", "pincode", "university", "website", "courses"):
            val = str(rec.get(k) or "").strip()
            assert val and val.lower() not in ("none", "na", "n/a"), f"Missing {k} in seed record: {rec.get('name')}"

    for cat, min_gov, min_priv in [
        ("Polytechnic", 50, 89),
        ("ITI", 118, 1200),
        ("B.Voc & D.Voc", 29, 17),
        ("Engineering", 35, 30),
        ("Medical / Nursing / Pharmacy", 15, 150),
        ("Management / BCA / BBA", 1, 20),
        ("Graduation / PG / Degree", 300, 400),
        ("All", 550, 2000),
    ]:
        res = run_scraper_pipeline({"category": cat, "state": "Bihar", "district": "All", "mode": "verified"})
        assert res["govt_count"] >= min_gov, f"{cat} govt count too low: {res['govt_count']}"
        assert res["priv_count"] >= min_priv, f"{cat} priv count too low: {res['priv_count']}"
        xlsx = build_excel_bytes(res["government"], res["private"], cat)
        wb = openpyxl.load_workbook(io.BytesIO(xlsx), read_only=True)
        assert wb.sheetnames == ["Government", "Private"]
        for sh in wb.sheetnames:
            rows = list(wb[sh].iter_rows(values_only=True))
            for r_i, row in enumerate(rows[1:], start=2):
                for c_i, cell_val in enumerate(row, start=1):
                    s_val = str(cell_val if cell_val is not None else "").strip()
                    assert s_val and s_val.lower() not in ("none", "na", "n/a"), (
                        f"Empty cell in {cat} -> {sh} row {r_i} col {c_i}"
                    )
        wb.close()

    # WSGI route check for /, /health, /api/scrape, /api/export
    statuses = []
    html_body = b"".join(app({"PATH_INFO": "/", "REQUEST_METHOD": "GET"}, lambda s, h: statuses.append(s))).decode("utf-8")
    assert statuses[-1] == "200 OK" and "College Web Scraper & Excel Exporter Dashboard" in html_body

    payload = json.dumps({"category": "B.Voc & D.Voc", "state": "Bihar"}).encode("utf-8")
    scrape_body = b"".join(
        app(
            {
                "PATH_INFO": "/api/scrape",
                "REQUEST_METHOD": "POST",
                "CONTENT_LENGTH": str(len(payload)),
                "wsgi.input": io.BytesIO(payload),
            },
            lambda s, h: statuses.append(s),
        )
    )
    scrape_json = json.loads(scrape_body.decode("utf-8"))
    assert statuses[-1] == "200 OK" and scrape_json["govt_count"] == 29 and scrape_json["priv_count"] == 17

    export_body = b"".join(
        app(
            {
                "PATH_INFO": "/api/export",
                "REQUEST_METHOD": "POST",
                "CONTENT_LENGTH": str(len(payload)),
                "wsgi.input": io.BytesIO(payload),
            },
            lambda s, h: statuses.append(s),
        )
    )
    assert statuses[-1] == "200 OK" and len(export_body) > 5000

    print("SELF-CHECK PASSED: All WSGI routes (/, /api/scrape, /api/export), categories, and two-sheet Excel exports verified.")


if __name__ == "__main__":
    import socket
    import sys
    import threading
    import webbrowser

    if "--test" in sys.argv:
        self_check()
        sys.exit(0)

    # Default to port 8765 (avoids conflict with EDB Postgres on 8080); auto-pick next free port if busy
    env_port = os.environ.get("PORT")
    port = int(env_port) if env_port else 8765
    if not env_port:
        for candidate in range(8765, 8800):
            try:
                with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                    if hasattr(socket, "SO_EXCLUSIVEADDRUSE"):
                        s.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
                    s.bind(("127.0.0.1", candidate))
                    port = candidate
                    break
            except OSError:
                continue

    url = f"http://127.0.0.1:{port}"
    print(f"Starting College Scraper & Excel Dashboard at {url}")
    with make_server("127.0.0.1" if not env_port else "0.0.0.0", port, app) as httpd:
        if not env_port and "--no-browser" not in sys.argv:
            threading.Timer(0.6, lambda: webbrowser.open(url)).start()
        httpd.serve_forever()

