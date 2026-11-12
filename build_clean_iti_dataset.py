import json
import pdfplumber
import re
from bs4 import BeautifulSoup
import os

DIST_PINS = {
    'Araria': '854311', 'Arwal': '804428', 'Aurangabad': '824101', 'Banka': '813102',
    'Begusarai': '851101', 'Bhagalpur': '812001', 'Bhojpur': '802301', 'Buxar': '802101',
    'Darbhanga': '846004', 'East Champaran': '845401', 'Gaya': '823001', 'Gopalganj': '841428',
    'Jamui': '811307', 'Jehanabad': '804408', 'Kaimur': '821101', 'Katihar': '854105',
    'Khagaria': '851204', 'Kishanganj': '855107', 'Lakhisarai': '811311', 'Madhepura': '852113',
    'Madhubani': '847211', 'Munger': '811201', 'Muzaffarpur': '842001', 'Nalanda': '803101',
    'Nawada': '805110', 'Patna': '800001', 'Purnia': '854301', 'Rohtas': '821115',
    'Saharsa': '852201', 'Samastipur': '848101', 'Saran': '841301', 'Sheikhpura': '811105',
    'Sheohar': '843329', 'Sitamarhi': '843302', 'Siwan': '841226', 'Supaul': '852131',
    'Vaishali': '844101', 'West Champaran': '845438'
}

def clean_dist(d):
    s = d.upper().replace(' ', '')
    if 'EAST' in s or 'PURBI' in s: return 'East Champaran'
    if 'WEST' in s or 'PASCHIM' in s or 'BETTIAH' in s: return 'West Champaran'
    if 'AURANG' in s: return 'Aurangabad'
    if 'MADHEP' in s: return 'Madhepura'
    if 'MUZAFF' in s: return 'Muzaffarpur'
    if 'SAHARS' in s: return 'Saharsa'
    if 'SHEIKH' in s or 'SHEKH' in s: return 'Sheikhpura'
    return d.strip().title()

def clean_name(raw_n, subdiv, dist):
    n = raw_n.strip().replace('\n', ' ')
    n = re.sub(r'\s+', ' ', n)
    is_women = 'witi' in n.lower() or 'women' in n.lower()
    is_lwe = 'lwe' in n.lower()
    core = re.sub(r'^(GOVT\.?\s*WOMEN\s*ITI|GOVT\.?\s*WITI|WOMEN\s*ITI|WITI|GOVT\.?\s*ITI|ITI)\s*[\(\-]?\s*', '', n, flags=re.I)
    core = re.sub(r'^\(LWE\)\s*', '', core, flags=re.I)
    core = core.strip(' ,-()')
    if not core:
        core = subdiv or dist
    core = core.title()
    if is_women:
        prefix = 'Government Women ITI'
    elif is_lwe:
        prefix = 'Government ITI (LWE)'
    else:
        prefix = 'Government ITI'
    final_name = f'{prefix}, {core}'
    if dist.lower() not in final_name.lower():
        final_name += f', {dist}'
    return final_name

print("1. Parsing DET Bihar officers...")
with open('web tool/raw_data/det_officers.html', 'r', encoding='utf-8', errors='ignore') as f:
    soup = BeautifulSoup(f.read(), 'html.parser')
officers = []
for row in soup.find_all('table')[1].find_all('tr')[1:]:
    cols = [c.get_text(strip=True) for c in row.find_all(['td', 'th'])]
    if len(cols) >= 6:
        raw_email = cols[5]
        split_raw = re.sub(r'(\.com|\.in|\.org|\.net|\.nic\.in|\.gov\.in)', r'\1 ', raw_email)
        emails = re.findall(r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+', split_raw)
        officers.append({
            'name': cols[1],
            'desig': cols[2],
            'dist': clean_dist(cols[3]),
            'mobile': cols[4],
            'emails': emails
        })

print("2. Parsing betastate Govt ITI PDF...")
with pdfplumber.open('web tool/raw_data/betastate_iti.pdf') as pdf:
    govt_pdf = []
    for p in range(10):
        for tbl in pdf.pages[p].extract_tables():
            for row in tbl:
                sr = (row[0] or '').strip().replace('\n', '')
                if sr.isdigit() and int(sr) > 0 and row[1] != '2':
                    govt_pdf.append({
                        'sr': int(sr),
                        'raw_name': (row[1] or '').strip().replace('\n', ' '),
                        'code': (row[2] or '').strip().replace('\n', ''),
                        'district': clean_dist(row[4] or ''),
                        'subdiv': (row[5] or '').strip().replace('\n', ' ').title(),
                        'address': (row[6] or '').strip().replace('\n', ' '),
                        'principal_pdf': (row[12] or '').strip().replace('\n', ' ').title() if len(row)>12 else '',
                    })

built_govt = []
for it in govt_pdf:
    dist = it['district'].lower()
    subdiv = re.sub(r'[^a-zA-Z0-9]', '', it['subdiv'].lower())
    it_name_clean = re.sub(r'[^a-zA-Z0-9]', '', it['raw_name'].lower())
    is_women = 'witi' in it_name_clean or 'women' in it_name_clean
    is_lwe = 'lwe' in it_name_clean or 'lwe' in it['address'].lower()
    
    cand = [o for o in officers if clean_dist(o['dist']).lower() == dist]
    best_cand = None
    best_score = -1
    for o in cand:
        score = 0
        desig_norm = re.sub(r'[^a-zA-Z0-9]', '', o['desig'].lower())
        o_name_norm = re.sub(r'[^a-zA-Z0-9]', '', o['name'].lower())
        p_name_norm = re.sub(r'[^a-zA-Z0-9]', '', it['principal_pdf'].lower())
        if p_name_norm and (p_name_norm in o_name_norm or o_name_norm in p_name_norm):
            score += 10
        o_women = 'women' in desig_norm or 'woman' in desig_norm or 'witi' in desig_norm
        if is_women == o_women:
            score += 5
        else:
            score -= 5
        o_lwe = 'lwe' in desig_norm
        if is_lwe == o_lwe:
            score += 3
        if subdiv and subdiv in desig_norm:
            score += 8
        for token in re.findall(r'[a-zA-Z]{4,}', it['raw_name'].lower()):
            if token in ('govt', 'witi'): continue
            if token in desig_norm:
                score += 5
        if score > best_score:
            best_score = score
            best_cand = o
            
    phone = ''
    email = ''
    if best_cand:
        raw_mob = best_cand['mobile'].strip()
        mobs = re.findall(r'\b[6-9]\d{9}\b', raw_mob)
        phone = mobs[0] if mobs else raw_mob
        em_list = best_cand['emails']
        iti_emails = [e for e in em_list if 'iti' in e.lower() or 'gov.in' in e.lower()]
        email = iti_emails[0] if iti_emails else (em_list[0] if em_list else 'det-bih@nic.in')
    else:
        email = 'det-bih@nic.in'
        
    pin_m = re.search(r'(?:PIN|\b)(8[0-5]\d{4})\b', it['address'])
    pin = pin_m.group(1) if pin_m else DIST_PINS.get(it['district'], '800001')
    
    c_name = clean_name(it['raw_name'], it['subdiv'], it['district'])
    
    if is_women:
        courses = "Trades: Electronic Mechanic, Computer Operator and Programming Assistant (COPA), Sewing Technology, Dress Making, Cosmetology, Electrician"
    elif is_lwe:
        courses = "Trades: Electrician, Fitter, Welder, Wireman, Mechanic (Diesel), Plumber, COPA"
    else:
        courses = "Trades: Electrician, Fitter, Welder, Wireman, Turner, Machinist, Mechanic (Motor Vehicle), COPA"
        
    built_govt.append({
        'code': it['code'],
        'name': c_name,
        'phone': phone,
        'email': email,
        'district': it['district'],
        'pincode': pin,
        'state': 'Bihar',
        'university': 'NCVT (DGT, Govt. of India) / SCVT Bihar, Directorate of Employment & Training, Patna',
        'website': '',
        'courses': courses,
        'category': 'ITI',
        'ownership': 'Government',
        'offers_course': True,
        'address': it['address']
    })

print(f"Total Government ITIs built: {len(built_govt)}")

print("3. Parsing PVTITIPSB members...")
with open('web tool/raw_data/pvtitipsb_members.html', 'r', encoding='utf-8', errors='ignore') as f:
    soup = BeautifulSoup(f.read(), 'html.parser')
pvtitipsb = {}
for row in soup.find_all('table')[0].find_all('tr')[1:]:
    cols = [c.get_text(strip=True) for c in row.find_all(['td', 'th'])]
    if len(cols) >= 5 and cols[0].isdigit():
        code = cols[1].strip()
        pvtitipsb[code] = {
            'contact': cols[5] if len(cols) > 5 else '',
            'trade': cols[6] if len(cols) > 6 else '',
            'website': cols[8] if len(cols) > 8 else ''
        }

print("4. Enriching & cleaning Private ITIs from seed_data.json...")
with open('web tool/seed_data.json', 'r', encoding='utf-8') as f:
    seed_all = json.load(f)

seed_pvt = [c for c in seed_all if c.get('category') == 'ITI' and c.get('ownership') == 'Private']
cleaned_pvt = []
seen_pvt_codes = set()

for c in seed_pvt:
    code = c.get('code')
    if code in seen_pvt_codes:
        continue
    seen_pvt_codes.add(code)
    
    # Enrich from PVTITIPSB if available
    if code in pvtitipsb:
        m = pvtitipsb[code]
        if m.get('contact') and (not c.get('phone') or len(str(c.get('phone'))) < 10):
            c['phone'] = m['contact'].split(',')[0].strip()
        if m.get('website') and not c.get('website'):
            c['website'] = m['website']
            
    # Clean up placeholder emails
    em = str(c.get('email') or '').strip()
    if em.lower() == 'highereducation-bih@nic.in':
        # replace with genuine email or empty
        c['email'] = ''
        
    # Ensure university is standard NCVT
    if not c.get('university') or 'BSDM' in str(c.get('university')):
        c['university'] = 'National Council for Vocational Training (NCVT), New Delhi / SCVT Bihar'
        
    # Ensure courses
    if not c.get('courses'):
        c['courses'] = 'Trades: Electrician, Fitter, Welder, Wireman, COPA'
        
    cleaned_pvt.append(c)

print(f"Total Private ITIs cleaned: {len(cleaned_pvt)}")

# Combine all: keep all non-ITI records from seed_all, plus built_govt, plus cleaned_pvt
non_iti = [c for c in seed_all if c.get('category') != 'ITI']
print(f"Non-ITI records: {len(non_iti)}")
final_seed = non_iti + built_govt + cleaned_pvt
print(f"Total updated seed records: {len(final_seed)}")

# Save to seed_data.json
with open('web tool/seed_data.json', 'w', encoding='utf-8') as f:
    json.dump(final_seed, f, indent=2, ensure_ascii=False)
print("Saved updated web tool/seed_data.json successfully!")
