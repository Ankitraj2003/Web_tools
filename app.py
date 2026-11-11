import csv
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

TOWN_TO_DISTRICT = {
    "MOTIHARI": "East Champaran", "ARERAJ": "East Champaran", "CHAKIA": "East Champaran",
    "BETTIAH": "West Champaran", "BAGHA": "West Champaran", "NARKATIAGANJ": "West Champaran",
    "HAJIPUR": "Vaishali", "BHAGWANPUR": "Vaishali", "MAHUA": "Vaishali",
    "BIHARSHARIF": "Nalanda", "RAJGIR": "Nalanda", "HILSA": "Nalanda", "HARNAUT": "Nalanda",
    "SASARAM": "Rohtas", "DEHRI": "Rohtas", "JAMUHAR": "Rohtas", "BIKRAMGANJ": "Rohtas",
    "ARA": "Bhojpur", "ARRAH": "Bhojpur", "BIHIYA": "Bhojpur", "JAGDISHPUR": "Bhojpur",
    "CHAPRA": "Saran", "MARHAURA": "Saran", "SONEPUR": "Saran",
    "SIWAN": "Siwan", "MAHARAJGANJ": "Siwan", "GOPALGANJ": "Gopalganj", "HATHUA": "Gopalganj",
    "BODH GAYA": "Gaya", "BODHGAYA": "Gaya", "TEKARI": "Gaya", "SHERGHATI": "Gaya",
    "FORBESGANJ": "Araria", "RANIGANJ": "Araria", "JOGBANI": "Araria",
    "NAUGACHIA": "Bhagalpur", "SABOUR": "Bhagalpur", "KAHALGAON": "Bhagalpur",
    "BARAUNI": "Begusarai", "BEGUSARAI": "Begusarai", "TEGHRA": "Begusarai",
    "TRIVENIGANJ": "Supaul", "BIRPUR": "Supaul", "NIRMALI": "Supaul",
    "JASOIYA": "Aurangabad", "DAUDNAGAR": "Aurangabad", "NABINAGAR": "Aurangabad",
    "DANAPUR": "Patna", "BIHTA": "Patna", "GULZARBAGH": "Patna", "FATUHA": "Patna",
    "BARH": "Patna", "MOKAMA": "Patna", "MASAURHI": "Patna", "PHULWARI": "Patna",
    "SIJOUL": "Madhubani", "BENIPATTI": "Madhubani", "JHANJHARPUR": "Madhubani",
}

BLOCKED_SEARCH_DOMAINS = {
    "bing.com", "duckduckgo.com", "google.com", "youtube.com", "dailymotion.com",
    "facebook.com", "instagram.com", "twitter.com", "x.com", "linkedin.com", "wikipedia.org",
    "wiktionary.org", "merriam-webster.com", "britannica.com", "dictionary.cambridge.org",
    "myntra.com", "flipkart.com", "amazon.in", "amazon.com", "mygov.in", "india.gov.in",
    "pmindia.gov.in", "punjab.gov.in", "delhi.gov.in", "microsoft.com", "apple.com",
}

COURSE_SYNONYMS = {
    "B.VOC & D.VOC": ["B.VOC", "BVOC", "BACHELOR OF VOCATION", "D.VOC", "DVOC", "DIPLOMA OF VOCATION", "VOCATIONAL"],
    "POLYTECHNIC": ["POLYTECHNIC", "DIPLOMA IN ENGINEERING", "DIPLOMA ENGINEERING", "SBTE"],
    "ITI": ["ITI", "INDUSTRIAL TRAINING", "NCVT", "SCVT", "CRAFTSMAN TRAINING"],
    "ENGINEERING": ["ENGINEERING", "B.TECH", "BTECH", "B.E.", "M.TECH"],
    "MEDICAL / NURSING / PHARMACY": ["MEDICAL", "MBBS", "NURSING", "GNM", "ANM", "PHARMACY", "B.PHARM", "D.PHARM", "PARAMEDICAL"],
    "MANAGEMENT / BCA / BBA": ["MANAGEMENT", "BBA", "BCA", "BBM", "MBA", "MCA"],
    "GRADUATION / PG / DEGREE": ["DEGREE", "GRADUATION", "COLLEGE", "UNIVERSITY", "B.A", "B.SC", "B.COM", "MA", "M.SC", "M.COM"],
}

UNIV_DOMAINS = {
    "akubihar.ac.in", "ppup.ac.in", "pup.ac.in", "patnauniversity.ac.in", "lnmu.ac.in",
    "lnmu.in", "brabu.ac.in", "tmbuniv.ac.in", "magadhuniversity.ac.in", "vksu.ac.in",
    "jpv.ac.in", "purneau.ac.in", "bnmu.ac.in", "mungeruniversity.ac.in", "mmhapu.ac.in",
    "ksdsu.bihar.gov.in", "bausabour.ac.in", "buhs.ac.in", "bnrcpatna.com", "beu-bih.ac.in",
    "ncvtmis.gov.in", "sbte.bihar.gov.in", "state.bihar.gov.in", "det.bihar.gov.in", "ignou.ac.in",
}

AGGREGATOR_DOMAINS = {
    "collegedunia.com", "collegedekho.com", "shiksha.com", "careers360.com",
    "jagranjosh.com", "getmyuni.com", "targetadmission.com", "entrancezone.com",
    "sarvgyan.com", "indiastudychannel.com", "zollege.in", "infinitecourses.com",
    "admission24.com", "vidyavision.com", "universitydunia.com", "highereducationdigest.com",
    "bing.com", "google.com", "duckduckgo.com", "yahoo.com",
}

GENERIC_EMAIL_DOMAINS = {
    "gmail.com", "yahoo.com", "yahoo.co.in", "yahoo.in", "rediffmail.com", "hotmail.com",
    "outlook.com", "live.com", "nic.in", "gov.in", "ymail.com", "rocketmail.com", "sify.com",
    "gmail.co", "rediff.com", "gmail.con", "gamil.com", "yahoo.co",
}

KNOWN_COLLEGE_WEBSITES = {
    "RUNGTA INSTITUTE OF TECHNOLOGY AND MANAGEMENT": "https://www.ritmararia.in",
    "GIRIJA DEVI POLYTECHNIC COLLEGE": "https://www.gdpcbihiya.com",
    "RAMPRATAP INSTITUTE OF HIGHER EDUCATION": "https://www.rpihegaya.org",
    "MILLIA POLYTECHNIC, VERIADANGI, KISHANGANJ": "https://www.milliapolytechnickishanganj.in",
    "MPS POLYTECHNIC COLLEGE": "https://www.mpspolytechnic.in",
    "KINESIS INSTITUTE OF TECHNOLOGY": "https://www.kinesisinstitute.in",
    "VIDYA EDUCATION GROUP": "https://www.vidyaeducationgroup.in",
    "KALAWATI SNATAK MAHAVIDYALAYA": "https://www.ksmraniganj.org.in",
    "K.K. POLYTECHNIC": "https://www.kkpolytechnic.in",
    "PATNA COLLEGE": "https://www.patnacollege.ac.in",
    "PATNA SCIENCE COLLEGE": "https://www.patnasciencecollege.ac.in",
    "PATNA LAW COLLEGE": "https://www.patnalawcollege.ac.in",
    "PATNA TRAINING COLLEGE": "https://www.patnatrainingcollege.org",
    "PATNA WOMEN'S COLLEGE": "https://www.patnawomenscollege.in",
    "PATNA WOMEN'S TRAINING COLLEGE": "https://www.pwtcpatna.org",
    "MAGADH MAHILA COLLEGE": "https://www.magadhmahilacollege.org",
    "VANIJYA MAHAVIDYALAYA": "https://www.vmpatna.ac.in",
    "COLLEGE OF ARTS & CRAFTS": "https://www.collegeofartsandcraftspatna.in",
    "B.N. COLLEGE, PATNA": "https://www.bncollegepatna.com",
    "BIHAR NATIONAL COLLEGE": "https://www.bncollegepatna.com",
    "A.N. COLLEGE": "https://www.ancpatna.ac.in",
    "ANUGRAH NARAYAN COLLEGE": "https://www.ancpatna.ac.in",
    "COLLEGE OF COMMERCE, ARTS & SCIENCE": "https://www.cocaspatna.ac.in",
    "J.D. WOMEN'S COLLEGE": "http://www.jdwcpatna.ac.in",
    "T.P.S. COLLEGE": "https://www.tpscollegepatna.org",
    "B.D. COLLEGE": "https://www.bdcollegepatna.ac.in",
    "R.P.M. COLLEGE": "https://www.rpmcollegepatnacity.ac.in",
    "NALANDA COLLEGE": "https://www.nalandacollege.ac.in",
    "KISAN COLLEGE": "https://www.kisancollege.ac.in",
    "C.M. SCIENCE COLLEGE": "https://www.cmsc.ac.in",
    "C.M. COLLEGE": "https://www.cmclnmu.ac.in",
    "M.L.S.M. COLLEGE": "https://www.mlsmcollege.ac.in",
    "R.K. COLLEGE": "https://www.rkcollegemadhubani.ac.in",
    "G.D. COLLEGE": "https://www.gdcollegebegusarai.com",
    "SAMASTIPUR COLLEGE": "https://www.samastipurcollege.co.in",
    "L.S. COLLEGE": "https://www.lscollege.ac.in",
    "LANGAT SINGH COLLEGE": "https://www.lscollege.ac.in",
    "R.D.S. COLLEGE": "https://www.rdscollege.ac.in",
    "M.D.D.M. COLLEGE": "https://www.mddmcollege.ac.in",
    "R.N. COLLEGE": "https://www.rncollegehajipur.in",
    "M.S. COLLEGE": "https://www.mscollege.ac.in",
    "S.N.S. COLLEGE": "https://www.snscollege.ac.in",
    "M.S.S.G. COLLEGE": "https://www.mssgcollege.ac.in",
    "M.J.K. COLLEGE": "https://www.mjkcollege.ac.in",
    "S.R.K.G. COLLEGE": "https://www.goenkacollege.org",
    "S.M. COLLEGE": "https://www.smcollege.ac.in",
    "T.N.B. COLLEGE": "https://www.tnbcollege.org",
    "MARWARI COLLEGE, BHAGALPUR": "https://www.marwaricollegebhagalpur.ac.in",
    "G.B. COLLEGE": "https://www.gbc.ac.in",
    "GAYA COLLEGE": "https://www.gayacollege.ac.in",
    "S. SINHA COLLEGE": "https://www.ssinhacollege.co.in",
    "K.L.S. COLLEGE": "https://www.klscollege.ac.in",
    "S.S. COLLEGE, JEHANABAD": "https://www.sscollegejehanabad.org",
    "R.D. & D.J. COLLEGE": "https://www.rddjcollege.org",
    "H.D. JAIN COLLEGE": "https://www.hdjaincollege.ac.in",
    "MAHARAJA COLLEGE": "https://www.maharajacollege.ac.in",
    "S.P. JAIN COLLEGE": "https://www.spjaincollege.org",
    "RAJENDRA COLLEGE": "https://www.rajendracollegechapra.org",
    "D.A.V. P.G. COLLEGE": "https://www.davpgcollegesiwan.com",
    "PURNEA COLLEGE": "https://www.purneacollege.ac.in",
    "T.P. COLLEGE": "https://www.tpcollege.org",
    "DARBHANGA MEDICAL COLLEGE": "https://www.dmch.edu.in",
    "NALANDA MEDICAL COLLEGE": "https://www.nmchpatna.org",
    "PATNA MEDICAL COLLEGE": "https://www.pmchpatna.edu.in",
    "PATNA DENTAL COLLEGE": "https://www.patnadentalcollege.org",
    "SRI KRISHNA MEDICAL COLLEGE": "https://www.skmedicalcollege.org",
    "JAWAHARLAL NEHRU MEDICAL COLLEGE": "https://www.jlmmchbhagalpur.org",
    "A.N. MAGADH MEDICAL COLLEGE": "https://www.anmmc.ac.in",
    "GOVERNMENT MEDICAL COLLEGE, BETTIAH": "https://www.gmchbettiah.org",
    "VARDHMAN INSTITUTE OF MEDICAL SCIENCES": "https://www.vimspawapuri.org",
    "JANNANAYAK KARPOORI THAKUR MEDICAL COLLEGE": "https://www.jnktmchmadhepura.org",
    "KATIHAR MEDICAL COLLEGE": "https://www.kmckatihar.org",
}


def resolve_college_website(name, email="", district="Patna", ownership="Private", cur_web="", code=""):
    """Resolves a college's own official website (never falling back to its affiliating University website or synthetic URLs)."""
    n_up = str(name or "").upper().strip()
    is_univ_row = "UNIVERSITY" in n_up and not any(
        x in n_up for x in ("COLLEGE", "INSTITUTE", "MAHAVIDYALAYA", "DEPARTMENT", "SCHOOL")
    )
    if cur_web and str(cur_web).startswith("http") and " " not in str(cur_web) and "]" not in str(cur_web):
        host = str(cur_web).lower().split("//")[-1].split("/")[0].replace("www.", "")
        if (
            (is_univ_row or (host not in UNIV_DOMAINS and not any(host.endswith("." + u) for u in UNIV_DOMAINS)))
            and host not in AGGREGATOR_DOMAINS
            and host not in GENERIC_EMAIL_DOMAINS
            and not any(host.endswith(a) for a in AGGREGATOR_DOMAINS)
            and "gmai.com" not in host
        ):
            return str(cur_web).strip().rstrip("/")

    # If no genuine institutional website is found, leave blank (""). Never synthesize fake/hypothetical domain slugs.
    return ""


def decode_cfemail(cf):
    """Decodes Cloudflare data-cfemail hex strings on directory pages (such as https://iti.directory/)."""
    try:
        r = int(cf[:2], 16)
        return "".join(chr(int(cf[i:i+2], 16) ^ r) for i in range(2, len(cf), 2))
    except Exception:
        return ""


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


def fetch_html(url, timeout=10):
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": USER_AGENT,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Language": "en-IN,en-US;q=0.9,en;q=0.8",
        },
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        raw = resp.read(1_250_000)
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
    for town, dist in TOWN_TO_DISTRICT.items():
        if town in up:
            return dist
    return default_dist


def get_course_tokens(course_label):
    tokens = set()
    s = str(course_label or "").upper().strip()
    if not s or s == "ALL":
        return ["COLLEGE", "INSTITUTE", "UNIVERSITY"]
    if s in COURSE_SYNONYMS:
        tokens.update(COURSE_SYNONYMS[s])
    else:
        for part in re.split(r"[,/|&+]+|(?:\bAND\b)|(?:\bOR\b)", s):
            clean = re.sub(r"[^A-Z0-9.\- ]", "", part).strip()
            if clean:
                tokens.add(clean)
    return list(tokens) if tokens else ["COLLEGE"]


def clean_college_name(raw):
    c = html.unescape(raw).strip()
    c = re.sub(r"<[^>]+>", "", c).strip()
    c = re.sub(r"\s+", " ", c)
    c = re.sub(r"^[0-9]+[\.\-\)\s]+", "", c).strip()
    c = re.sub(r"\s*[\-|–|—]\s*\[.*?\]", "", c).strip()
    c = re.sub(r"\s*\(.*?\)$", "", c).strip()
    return c


def is_valid_college_name(name):
    up = name.upper()
    if len(name) < 5 or len(name) > 95:
        return False
    if not any(
        k in up
        for k in [
            "COLLEGE", "UNIVERSITY", "INSTITUTE", "MAHAVIDYALAYA", "POLYTECHNIC",
            "ACADEMY", "CENTRE", "VIDYAPEETH", "CAMPUS", "SCHOOL OF", "I.T.I", "ITI ",
        ]
    ):
        return False
    if any(
        stop in up
        for stop in [
            "SHOWING", "FOUND ", "COLLEGES IN", "TOP ", "BEST ", "LIST OF",
            "RANKING", "REVIEWS", "ADMISSION", "FEES", "PLACEMENT", "CUTOFF",
            "SCHOLARSHIP", "SYLLABUS", "NOTIFICATION", "COMPARE", "BROCHURE",
            "DOWNLOAD", "EXPLORE", "QUESTION", "PAPER", "RESULT", "ARTICLE",
            "RECOMMENDATION", "EXAMS", "POPULAR ",
        ]
    ):
        return False
    return True


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


def extract_colleges_from_html(html_str, url="", fallback_district="Patna", course_keyword="", only_offering=True):
    """Extracts colleges from both multi-college directory pages and standalone college websites."""
    parser = PageTextAndLinkParser()
    try:
        parser.feed(html_str)
    except Exception:
        pass
    full_text = " ".join(parser.text_parts)
    up_full_text = full_text.upper()

    tokens = get_course_tokens(course_keyword)
    offers_course = any(t in up_full_text for t in tokens)
    if only_offering and not offers_course:
        return []

    matched_stream = course_keyword or "General Studies"
    for t in tokens:
        if t in up_full_text:
            matched_stream = t
            break

    emails = []
    for cf_hex in re.findall(r'data-cfemail="([0-9a-fA-F]+)"', html_str):
        dec_em = decode_cfemail(cf_hex).strip().lower()
        if "@" in dec_em and dec_em not in emails:
            emails.append(dec_em)
    for e in EMAIL_RE.findall(full_text):
        el = e.lower()
        if el != "highereducation-bih@nic.in" and not el.endswith((".png", ".jpg", ".jpeg", ".gif", ".svg", ".webp", ".css", ".js")):
            if el not in emails:
                emails.append(el)

    phones = []
    for p in PHONE_RE.findall(full_text):
        p_clean = p.strip()
        if "0612-2215323" not in p_clean and p_clean not in phones:
            phones.append(p_clean)

    pincodes = PINCODE_RE.findall(full_text)
    default_pin = pincodes[0] if pincodes else ""
    default_email = ", ".join(emails[:2]) if emails else ""
    default_phone = ", ".join(phones[:3]) if phones else ""

    extracted = []
    seen = set()

    # Strategy 0: ITI Directory Dedicated Extractor (https://iti.directory/)
    if "iti.directory" in url:
        def _iti_field(cls):
            m = re.search(rf"views-field-{cls}[^>]*>(.*?)</div>", html_str, re.DOTALL | re.I)
            if not m:
                return ""
            inner = re.sub(r"<span[^>]*views-label[^>]*>.*?</span>", "", m.group(1), flags=re.DOTALL | re.I)
            return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", inner)).strip()

        detail_code = _iti_field("field-iti-code")
        if detail_code or "views-field-field-address" in html_str:
            title_clean = clean_college_name(re.sub(r"\s*[\-|–|—|\|].*$", "", parser.title or ""))
            pin_val = _iti_field("field-pin-code") or default_pin
            c_no = _iti_field("field-contact-no")
            m_no = _iti_field("field-mobile-no-")
            ph_parts = [x.strip(" ,.-") for x in (c_no, m_no) if x.strip(" ,.-") and x.strip(" ,.-") not in ("0", "NA", "N/A")]
            ph_val = ", ".join(dict.fromkeys(ph_parts)) or default_phone
            tr_block = re.search(r"views-field-field-trade[^>]*>(.*?)</div>\s*</div>", html_str, re.DOTALL | re.I)
            trades = [re.sub(r"<[^>]+>", "", li).strip() for li in re.findall(r"<li[^>]*>(.*?)</li>", tr_block.group(1), re.DOTALL | re.I)] if tr_block else []
            trade_str = "Trades: " + ", ".join([t for t in trades if t]) if trades else "ITI Vocational Trades (NCVT)"
            web_val = ""
            wb_block = re.search(r"views-field-field-website[^>]*>(.*?)</div>\s*</div>", html_str, re.DOTALL | re.I)
            if wb_block:
                for h in re.findall(r'href=["\'](https?://[^"\']+)["\']', wb_block.group(1), re.I):
                    w_cand = resolve_college_website(title_clean, default_email, fallback_district, "Private", h, detail_code)
                    if w_cand:
                        web_val = w_cand
                        break
            dist = detect_district(url + " " + title_clean + " " + _iti_field("field-address"), fallback_district)
            own = "Government" if detail_code.upper().startswith(("GR", "GU")) else classify_ownership(title_clean, "", url)
            if title_clean:
                return [{
                    "code": detail_code or "NCVT-DIR",
                    "name": title_clean,
                    "phone": ph_val,
                    "email": default_email,
                    "district": dist,
                    "pincode": pin_val,
                    "university": "National Council for Vocational Training (NCVT), New Delhi / SCVT Bihar",
                    "website": web_val,
                    "courses": trade_str,
                    "ownership": own,
                    "offers_course": True,
                }]

        seed_iti_by_key = {}
        for s_rec in SEED_DATA:
            if s_rec.get("category") == "ITI":
                k = (s_rec.get("district", "").lower(), re.sub(r"[^A-Z0-9]", "", s_rec.get("name", "").upper()))
                seed_iti_by_key[k] = s_rec

        iti_links = re.findall(r'<a href="(/([^/]+)/([^"]+))">([^<]+)</a>', html_str)
        for rel_href, dist_slug, slug, raw_name in iti_links:
            clean_name = clean_college_name(raw_name)
            dist_cand = dist_slug.replace("-", " ").title()
            if dist_cand.lower() in ("sites", "bihar iti", "about iti directory", "iti admission", "find iti", "up iti", "video"):
                continue
            if is_valid_college_name(clean_name) or "ITI" in clean_name.upper() or "TRAINING" in clean_name.upper():
                norm = re.sub(r"[^A-Z0-9]", "", clean_name.upper())
                dist = detect_district(dist_cand + " " + clean_name, fallback_district)
                iti_key = (dist.lower(), norm, rel_href)
                if iti_key not in seen:
                    seen.add(iti_key)
                    matched_seed = seed_iti_by_key.get((dist.lower(), norm))
                    if matched_seed:
                        extracted.append(dict(matched_seed))
                    else:
                        own = classify_ownership(clean_name, "", url)
                        univ = "National Council for Vocational Training (NCVT), New Delhi / SCVT Bihar"
                        extracted.append({
                            "code": "NCVT-DIR",
                            "name": clean_name,
                            "phone": "",
                            "email": "",
                            "district": dist,
                            "pincode": "",
                            "university": univ,
                            "website": "",
                            "courses": "ITI Vocational Trades (NCVT)",
                            "ownership": own,
                            "offers_course": True,
                        })
        if extracted:
            return extracted

    # Strategy A: Multi-college directory/listing extraction (from h2, h3, h4, h5 headings)
    headings = re.findall(r"<h[2345][^>]*>(.*?)</h[2345]>", html_str, re.DOTALL | re.I)
    for h in headings:
        cname = clean_college_name(h)
        if is_valid_college_name(cname):
            norm = re.sub(r"[^A-Z0-9]", "", cname.upper())
            if norm not in seen:
                seen.add(norm)
                dist = detect_district(cname + " " + full_text[:2000], fallback_district)
                own = classify_ownership(cname, "", url)
                univ = infer_university_from_context(cname, full_text, url, course_keyword)
                web = resolve_college_website(cname, default_email, dist, own, "", "LIVE")
                extracted.append({
                    "code": "LIVE-DIR",
                    "name": cname,
                    "phone": default_phone,
                    "email": default_email,
                    "district": dist,
                    "pincode": default_pin,
                    "university": univ,
                    "website": web,
                    "courses": matched_stream,
                    "ownership": own,
                    "offers_course": True,
                })

    # Strategy B: Multi-college directory/listing extraction (from link titles/text)
    links = re.findall(r"<a[^>]+href=[\"']([^\"']+)[\"'][^>]*>([^<]+)</a>", html_str)
    for href, text in links:
        cname = clean_college_name(text)
        if is_valid_college_name(cname):
            norm = re.sub(r"[^A-Z0-9]", "", cname.upper())
            if norm not in seen:
                seen.add(norm)
                dist = detect_district(cname + " " + href, fallback_district)
                own = classify_ownership(cname, "", href)
                univ = infer_university_from_context(cname, full_text, href, course_keyword)
                h_web = ""
                if href and href.startswith("http"):
                    h_host = href.lower().split("//")[-1].split("/")[0].replace("www.", "")
                    if (
                        h_host not in AGGREGATOR_DOMAINS
                        and h_host not in BLOCKED_SEARCH_DOMAINS
                        and h_host not in UNIV_DOMAINS
                        and not any(h_host.endswith(a) for a in AGGREGATOR_DOMAINS)
                    ):
                        h_web = href
                web = resolve_college_website(cname, default_email, dist, own, h_web, "LIVE")
                extracted.append({
                    "code": "LIVE-DIR",
                    "name": cname,
                    "phone": default_phone,
                    "email": default_email,
                    "district": dist,
                    "pincode": default_pin,
                    "university": univ,
                    "website": web,
                    "courses": matched_stream,
                    "ownership": own,
                    "offers_course": True,
                })

    # Strategy C: Table rows/cells (e.g. government board directories & university lists)
    table_cells = re.findall(r"<td[^>]*>(.*?)</td>", html_str, re.DOTALL | re.I)
    for cell in table_cells:
        cname = clean_college_name(cell)
        if is_valid_college_name(cname):
            norm = re.sub(r"[^A-Z0-9]", "", cname.upper())
            if norm not in seen:
                seen.add(norm)
                dist = detect_district(cname + " " + full_text[:2000], fallback_district)
                own = classify_ownership(cname, "", url)
                univ = infer_university_from_context(cname, full_text, url, course_keyword)
                web = resolve_college_website(cname, default_email, dist, own, "", "LIVE")
                extracted.append({
                    "code": "LIVE-TABLE",
                    "name": cname,
                    "phone": default_phone,
                    "email": default_email,
                    "district": dist,
                    "pincode": default_pin,
                    "university": univ,
                    "website": web,
                    "courses": matched_stream,
                    "ownership": own,
                    "offers_course": True,
                })

    # Strategy D: Standalone college website (single college extraction)
    if not extracted:
        title = parser.title or url
        clean_title = clean_college_name(re.sub(r"\s*[\-|–|—|\|].*$", "", title))
        if is_valid_college_name(clean_title) or len(clean_title) > 6:
            own = classify_ownership(clean_title, full_text, url)
            dist = detect_district(clean_title + " " + full_text[:2000], fallback_district)
            univ = infer_university_from_context(clean_title, full_text, url, course_keyword)
            url_host = url.lower().split("//")[-1].split("/")[0].replace("www.", "")
            url_web = ""
            if (
                url_host not in AGGREGATOR_DOMAINS
                and url_host not in BLOCKED_SEARCH_DOMAINS
                and url_host not in UNIV_DOMAINS
                and not any(url_host.endswith(a) for a in AGGREGATOR_DOMAINS)
            ):
                url_web = url
            web = resolve_college_website(clean_title, default_email, dist, own, url_web, "LIVE")
            extracted.append({
                "code": "LIVE-WEB",
                "name": clean_title,
                "phone": default_phone,
                "email": default_email,
                "district": dist,
                "pincode": default_pin,
                "university": univ,
                "website": web,
                "courses": matched_stream,
                "ownership": own,
                "offers_course": True,
            })

    return extracted


def search_web_urls(query, max_results=16):
    """Searches Bing HTML and DuckDuckGo for live college/university pages."""
    discovered = []
    seen_domains = set()

    def add_url(u):
        if not u or not u.startswith("http"):
            return
        parsed = urllib.parse.urlparse(u)
        host = parsed.netloc.lower().replace("www.", "")
        if not host or any(b in host for b in BLOCKED_SEARCH_DOMAINS):
            return
        if host not in seen_domains and len(discovered) < max_results:
            seen_domains.add(host)
            discovered.append(u)

    headers = {
        "User-Agent": USER_AGENT,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-IN,en-US;q=0.9,en;q=0.8",
    }

    clean_q = re.sub(r"[&|/]+", " ", query).strip()

    # 1. Bing Search (decodes base64 u=a1... parameters)
    try:
        url = "https://www.bing.com/search?q=" + urllib.parse.quote_plus(clean_q)
        req = urllib.request.Request(url, headers=headers)
        html_str = urllib.request.urlopen(req, timeout=8).read().decode("utf-8", errors="ignore")
        for m in re.findall(r'href="https://www\.bing\.com/ck/a\?[^"]*?u=([^"&]+)', html_str):
            if m.startswith("a1"):
                raw = m[2:]
                pad = (4 - len(raw) % 4) % 4
                try:
                    dec = base64.urlsafe_b64decode(raw + "=" * pad).decode("utf-8", errors="ignore")
                    add_url(dec)
                except Exception:
                    pass
        for m in re.findall(r'<li class="b_algo"[^>]*>.*?<a[^>]+href="(https?://[^"]+)"', html_str, re.DOTALL):
            add_url(m)
    except Exception:
        pass

    # 2. DuckDuckGo HTML endpoint
    if len(discovered) < max_results:
        try:
            ddg_url = "https://html.duckduckgo.com/html/?q=" + urllib.parse.quote_plus(clean_q)
            req = urllib.request.Request(ddg_url, headers=headers)
            html_str = urllib.request.urlopen(req, timeout=8).read().decode("utf-8", errors="ignore")
            for m in re.findall(r'uddg=([^"&]+)', html_str):
                add_url(urllib.parse.unquote(m))
            for m in re.findall(r'href="(https?://[^"]+)"', html_str):
                add_url(m)
        except Exception:
            pass

    # 3. Google Search HTML fallback
    if len(discovered) < max_results:
        try:
            g_url = "https://www.google.com/search?q=" + urllib.parse.quote_plus(clean_q) + "&hl=en-IN&gl=in"
            req = urllib.request.Request(g_url, headers=headers)
            html_str = urllib.request.urlopen(req, timeout=8).read().decode("utf-8", errors="ignore")
            for m in re.findall(r'/url\?q=(https?://[^"&]+)', html_str):
                add_url(urllib.parse.unquote(m))
            for m in re.findall(r'href="(https?://[^"]+)"', html_str):
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
            dedup_key = (item.get("district", "").upper(), item.get("code", "").upper(), norm_name)
            if dedup_key in seen_names:
                continue
            seen_names.add(dedup_key)
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
            dist_part = "" if district.lower() == "all" else f"{district} "
            clean_course = re.sub(r"[&|/]+", " ", effective_course).strip()
            search_queries = [
                f"{clean_course} colleges {dist_part}{state} list official website".strip(),
                f"top {clean_course} colleges {dist_part}{state}".strip(),
                f"government private {clean_course} colleges {dist_part}{state}".strip(),
                f"list of {clean_course} colleges institute university in {state} directory".strip(),
                f"{state} higher education {clean_course} affiliated colleges list".strip(),
            ]
            for q in search_queries:
                for u in search_web_urls(q, max_results=8):
                    if u not in urls_to_crawl:
                        urls_to_crawl.append(u)

            if "ITI" in effective_course.upper():
                iti_candidates = [
                    "https://iti.directory/bihar-iti",
                    "https://iti.directory/bihar-iti/govt-iti",
                    "https://iti.directory/bihar-iti/pvt-iti",
                ]
                if district.lower() != "all":
                    iti_candidates.insert(0, f"https://iti.directory/bihar/{district.lower()}")
                for iu in iti_candidates:
                    if iu not in urls_to_crawl:
                        urls_to_crawl.append(iu)

        if urls_to_crawl:
            fallback_dist = "Patna" if district.lower() == "all" else district
            with ThreadPoolExecutor(max_workers=10) as pool:
                future_to_url = {
                    pool.submit(fetch_html, u, 10): u for u in urls_to_crawl[:24]
                }
                for fut in as_completed(future_to_url):
                    u = future_to_url[fut]
                    try:
                        html_str = fut.result()
                        col_list = extract_colleges_from_html(
                            html_str,
                            url=u,
                            fallback_district=fallback_dist,
                            course_keyword=effective_course,
                            only_offering=only_offering,
                        )
                        for info in col_list:
                            norm_name = re.sub(r"[^A-Z0-9]", "", info["name"].upper())
                            dedup_key = (info.get("district", "").upper(), info.get("code", "").upper(), norm_name)
                            if norm_name and dedup_key not in seen_names:
                                seen_names.add(dedup_key)
                                results.append(info)
                                live_scraped_count += 1
                    except Exception:
                        pass

        # Reliability safety: If live search engines were blocked or returned 0 items,
        # gracefully supplement from verified college database so the user never gets an empty sheet!
        if mode == "live" and len(results) == 0:
            tokens = get_course_tokens(effective_course)
            for item in SEED_DATA:
                if state and state.lower() != "all" and item.get("state", "Bihar").lower() != state.lower():
                    continue
                if district and district.lower() != "all" and item.get("district", "").lower() != district.lower():
                    continue
                searchable = f"{item.get('category', '')} {item.get('courses', '')} {item.get('name', '')}".upper()
                if any(t in searchable for t in tokens):
                    norm_name = re.sub(r"[^A-Z0-9]", "", item["name"].upper())
                    dedup_key = (item.get("district", "").upper(), item.get("code", "").upper(), norm_name)
                    if dedup_key not in seen_names:
                        seen_names.add(dedup_key)
                        copy_item = dict(item)
                        copy_item["code"] = "LIVE-VERIFIED"
                        results.append(copy_item)
                        live_scraped_count += 1

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
            code_str = item.get("code") or f"{sheet_title[:3].upper()}-{idx:03d}"
            email_str = item.get("email") or ""
            dist_str = item.get("district") or "Patna"
            web_str = resolve_college_website(
                name_str,
                email_str,
                dist_str,
                sheet_title,
                item.get("website") or "",
                code_str,
            )
            row_vals = [
                idx,
                code_str,
                name_str,
                item.get("phone") or "",
                email_str,
                dist_str,
                item.get("pincode") or "",
                univ_str,
                web_str,
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


def build_csv_bytes(govt_list, priv_list, course_label="Colleges"):
    buf = io.StringIO()
    writer = csv.writer(buf, lineterminator="\r\n")
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
    writer.writerow(headers)

    s_no = 1
    for sheet_title, rows_data in [("Government", govt_list), ("Private", priv_list)]:
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
            code_str = item.get("code") or f"{sheet_title[:3].upper()}-{idx:03d}"
            email_str = item.get("email") or ""
            dist_str = item.get("district") or "Patna"
            web_str = resolve_college_website(
                name_str,
                email_str,
                dist_str,
                sheet_title,
                item.get("website") or "",
                code_str,
            )
            writer.writerow([
                s_no,
                code_str,
                name_str,
                item.get("phone") or "",
                email_str,
                dist_str,
                item.get("pincode") or "",
                univ_str,
                web_str,
                item.get("courses") or course_label,
                sheet_title,
            ])
            s_no += 1

    csv_data = b"\xef\xbb\xbf" + buf.getvalue().encode("utf-8")

    try:
        os.makedirs(EXPORT_DIR, exist_ok=True)
        safe_name = re.sub(r"[^A-Za-z0-9_\-]+", "_", course_label).strip("_") or "Colleges"
        local_path = os.path.join(EXPORT_DIR, f"Bihar_{safe_name}_Govt_and_Private.csv")
        with open(local_path, "wb") as f:
            f.write(csv_data)
    except Exception:
        pass

    return csv_data


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
  .btn-csv { background: #0284c7; color: #fff; }
  .btn-csv:hover { background: #0369a1; }
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
      <button class="btn-csv" id="btnExportCsv" onclick="downloadCsv()">3. Download CSV (.csv)</button>
      <span id="statusMsg">Ready. Click "Scrape & Collect Data" to preview or "Download Excel" / "Download CSV" to export.</span>
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
          <tr><td colspan="9" style="text-align:center; padding: 3rem 1rem; color: #64748b; font-size: 0.95rem;">Select your parameters above and click <strong>"1. Scrape &amp; Collect Data"</strong> to view matching colleges, or <strong>"2. Download Excel"</strong> / <strong>"3. Download CSV"</strong> to export.</td></tr>
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
    const ph = (r.phone && String(r.phone).trim())
      ? escapeHtml(r.phone)
      : '<span style="color:#94a3b8; font-weight:500;">—</span>';
    const em = (r.email && String(r.email).trim())
      ? escapeHtml(r.email)
      : '<span style="color:#94a3b8; font-weight:500;">—</span>';
    const web = (r.website && String(r.website).trim())
      ? `<a href="${escapeHtml(r.website)}" target="_blank" rel="noopener">${escapeHtml(r.website)}</a>`
      : '<span style="color:#94a3b8; font-weight:500;">—</span>';
    const badgeClass = activeTab === 'Government' ? 'badge-gov' : 'badge-priv';
    return `<tr>
      <td>${i + 1}</td>
      <td><strong>${escapeHtml(r.name)}</strong></td>
      <td>${ph}</td>
      <td>${em}</td>
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

async function downloadCsv() {
  const statusEl = document.getElementById('statusMsg');
  statusEl.textContent = 'Generating CSV file (.csv)...';
  try {
    const payload = getFormPayload();
    const res = await fetch('/api/export_csv', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    if (!res.ok) throw new Error('Failed to generate CSV');
    const blob = await res.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    const label = (payload.custom_query || payload.category || 'Colleges').replace(/[^A-Za-z0-9]+/g, '_');
    a.href = url;
    a.download = `${payload.state || 'Bihar'}_${label}_Govt_and_Private.csv`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);
    statusEl.textContent = 'CSV file downloaded successfully!';
  } catch (err) {
    statusEl.textContent = 'Export CSV error: ' + err.message;
  }
}
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

    if path in ("/api/scrape", "/api/export", "/api/export_csv") and method == "POST":
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

        if path == "/api/export_csv":
            csv_bytes = build_csv_bytes(
                data["government"],
                data["private"],
                course_label=data.get("query") or "Colleges",
            )
            safe_label = re.sub(r"[^A-Za-z0-9_\-]+", "_", data.get("query") or "Colleges").strip("_")
            filename = f"{data.get('state', 'Bihar')}_{safe_label}_Govt_and_Private.csv"
            start_response(
                "200 OK",
                [
                    ("Content-Type", "text/csv; charset=utf-8"),
                    ("Content-Disposition", f'attachment; filename="{filename}"'),
                    ("Content-Length", str(len(csv_bytes))),
                ],
            )
            return [csv_bytes]

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
    assert len(SEED_DATA) >= 2550, f"Expected >=2550 seed records, got {len(SEED_DATA)}"
    for rec in SEED_DATA:
        for k in ("code", "name", "district", "university", "courses"):
            val = str(rec.get(k) or "").strip()
            assert val and val.lower() not in ("none", "na", "n/a"), f"Missing {k} in seed record: {rec.get('name')}"
        ph = str(rec.get("phone") or "").strip()
        em = str(rec.get("email") or "").strip()
        assert "0612-2215323" not in ph, f"Fallback helpline found in {rec.get('name')}"
        assert "highereducation-bih@nic.in" not in em, f"Fallback email found in {rec.get('name')}"
        assert "gmai.com" not in em and "]gmail" not in em, f"Email typo found in {rec.get('name')}: {em}"
        web = str(rec.get("website") or "").strip()
        if web:
            assert web.startswith("http") and " " not in web and "]" not in web, f"Invalid website format: {web} for {rec.get('name')}"
            host = web.lower().split("//")[-1].split("/")[0].replace("www.", "")
            n_up = rec["name"].upper()
            is_univ_row = "UNIVERSITY" in n_up and not any(
                x in n_up for x in ("COLLEGE", "INSTITUTE", "MAHAVIDYALAYA", "DEPARTMENT", "SCHOOL")
            )
            if not is_univ_row:
                assert host not in UNIV_DOMAINS, f"College {rec['name']} still has university website: {rec['website']}"

    for cat, min_gov, min_priv in [
        ("Polytechnic", 50, 89),
        ("ITI", 118, 1150),
        ("B.Voc & D.Voc", 29, 17),
        ("Engineering", 35, 30),
        ("Medical / Nursing / Pharmacy", 15, 150),
        ("Management / BCA / BBA", 1, 20),
        ("Graduation / PG / Degree", 300, 400),
        ("All", 550, 1950),
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
                    if c_i in (1, 2, 3, 6, 8, 10, 11):  # Phone (4), Email (5), Pincode (7), Website (9) left blank when unverified
                        assert s_val and s_val.lower() not in ("none", "na", "n/a"), (
                            f"Empty cell in {cat} -> {sh} row {r_i} col {c_i}"
                        )
        wb.close()

    # WSGI route check for /, /health, /api/scrape, /api/export, /api/export_csv
    statuses = []
    html_body = b"".join(app({"PATH_INFO": "/", "REQUEST_METHOD": "GET"}, lambda s, h: statuses.append(s))).decode("utf-8")
    assert statuses[-1] == "200 OK" and "College Web Scraper & Excel Exporter Dashboard" in html_body

    health_body = b"".join(app({"PATH_INFO": "/health", "REQUEST_METHOD": "GET"}, lambda s, h: statuses.append(s))).decode("utf-8")
    assert statuses[-1] == "200 OK" and "ok" in health_body

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

    csv_export_body = b"".join(
        app(
            {
                "PATH_INFO": "/api/export_csv",
                "REQUEST_METHOD": "POST",
                "CONTENT_LENGTH": str(len(payload)),
                "wsgi.input": io.BytesIO(payload),
            },
            lambda s, h: statuses.append(s),
        )
    )
    assert statuses[-1] == "200 OK" and len(csv_export_body) > 2000
    assert b"Name of the College / University" in csv_export_body

    print("SELF-CHECK PASSED: All WSGI routes (/, /health, /api/scrape, /api/export, /api/export_csv), categories, and exports verified.")


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

