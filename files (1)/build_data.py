#!/usr/bin/env python3
"""
build_data.py — wire The Open Jobs Map to live ONS data
========================================================
Downloads the latest ONS "Labour demand volumes by SOC 2020" release (online job
advert counts by occupation and area, monthly back to 2017, Open Government Licence),
aggregates detailed occupations into the ~21 fields of work the app uses, and writes
data.json. Pay comes from a bundled ASHE-derived median table (optionally overridable);
the "what the work is like" / "routes in" text comes from National Careers Service-style
profiles bundled below.

WHY A PIPELINE (not in the browser): the ONS data is a ~45 MB spreadsheet on a release
page; browsers can't fetch/parse that directly (format + cross-origin). So you run this
on your machine or a scheduled job; it emits data.json, which index.html reads.

QUICK START
-----------
    python3 -m pip install requests pandas openpyxl
    python3 build_data.py --inspect      # prints the workbook's sheets + columns
    # set CONFIG below to match what --inspect shows, then:
    python3 build_data.py                 # writes ./data.json

IMPORTANT: ONS occasionally changes the workbook's sheet name / column layout. This
script auto-detects where it can, but you should run --inspect once and confirm the
CONFIG block matches. The download step and field grouping are robust; the only thing
that may need a tweak is which sheet/columns hold the occupation × month grid.
"""

import argparse, io, json, re, sys, datetime
from urllib.parse import urljoin

DATASET_PAGE = ("https://www.ons.gov.uk/employmentandlabourmarket/peopleinwork/"
                "employmentandemployeetypes/datasets/"
                "labourdemandvolumesbystandardoccupationclassificationsoc2020uk")

# ASHE Table 15 = earnings by region x four-digit SOC (annual; Open Government Licence).
# Gives a real median per role per region. Published each autumn (provisional, then revised).
ASHE_PAGE = ("https://www.ons.gov.uk/employmentandlabourmarket/peopleinwork/"
             "earningsandworkinghours/datasets/regionbyoccupation4digitsoc2010ashetable15")

# ---- CONFIG: confirm against `--inspect` output, then adjust if needed -------------
CONFIG = {
    # Sheet that holds the occupation × time grid. None = auto-pick the largest sheet.
    "sheet": None,
    # Column header (case-insensitive substring) holding the occupation NAME.
    "occupation_name_col": "occupation",
    # Column header substring holding the SOC code (optional; used only as a hint).
    "soc_code_col": "soc",
    # Column header substring holding the geography NAME (region/LA). Optional.
    "geography_col": "geography",
    # Month columns are auto-detected: any header that parses as a date / "YYYY MMM".

    # ---- ASHE Table 15 (pay) workbook selection inside the downloaded zip ----
    # Confirm with: python3 build_data.py --inspect-ashe
    "ashe_file_match": ["annual","gross"],   # pick the 'All employees, Annual pay - Gross' workbook
    "ashe_file_exclude": ["male","female","part","full","incentive","weekly","hourly","basic","overtime"],
    "ashe_stat_sheet": "Median",             # worksheet holding the Median statistic
}

# ---- 21 fields: keyword rules to assign each ONS occupation to a field -------------
# First field whose keyword matches the occupation name wins, so order = priority.
FIELD_KEYWORDS = [
 ("health",   ["nurse","nursing","midwif","health profession","paramedic","physiotherap","radiograph","clinical","doctor","medical practitioner","healthcare assistant","nursing auxiliar"]),
 ("care",     ["care worker","carer","care assistant","senior care","social care","home care","support worker","residential"]),
 ("it",       ["programmer","software","web design","web develop","it ","information technology","cyber","data analyst","data scientist","database","systems admin","devops","network","it user","it support","it technician"]),
 ("eng",      ["engineer","engineering technician","draughts"]),
 ("trade",    ["electrician","plumber","carpenter","joiner","bricklay","plasterer","construction","scaffold","roofer","glazier","painter and decorat","steel erect","groundwork"]),
 ("edu",      ["teacher","teaching","lecturer","education","tutor","classroom","nursery","early years","teaching assistant","school"]),
 ("hosp",     ["chef","cook","waiter","waitress","bar staff","catering","kitchen","hospitality","publican","barista","restaurant"]),
 ("log",      ["driver","hgv","lgv","delivery","warehouse","forklift","logistics","transport","postal","courier","van "]),
 ("fin",      ["account","finance","financial","auditor","bookkeep","payroll","actuar","insurance underwrit","tax "]),
 ("legal",    ["solicitor","barrister","legal","paralegal","conveyanc","law "]),
 ("sci",      ["scientist","laborator","lab technician","research","chemist","biologist","physicist","pharmacolog","microbiolog","quality assurance scientist"]),
 ("mkt",      ["marketing","advertising","public relations","pr ","seo","brand","social media","content"]),
 ("hr",       ["human resource","hr ","recruitment","personnel","talent","learning and develop"]),
 ("creative", ["graphic design","ux","ui design","artist","photograph","video","animator","designer","creative","media","journalist","editor"]),
 ("energy",   ["energy","electricity","gas ","utilities","renewable","power plant","water "]),
 ("mfg",      ["production","manufactur","assembl","machine operat","factory","plant operat","fabricat","process operat"]),
 ("maint",    ["maintenance","facilities","cleaner","caretaker","janitor","groundskeep","gardener","handyman","repair"]),
 ("cust",     ["customer service","customer support","call centre","contact centre","claims handler","complaints"]),
 ("sales",    ["sales","account manager","business development","retail manager","sales assistant","buyer","merchand"]),
 ("admin",    ["administ","secretar","receptionist","clerk","data entry","office manager","personal assistant","typist","records"]),
 ("retail",   ["retail","shop","cashier","checkout","store ","shelf"]),
]
# fallback by SOC major group (first digit) if no keyword matched
SOC1_FALLBACK = {"1":"admin","2":"sci","3":"admin","4":"admin","5":"trade","6":"care","7":"sales","8":"mfg","9":"maint"}

# ---- field metadata: group label, ASHE-derived median pay, careers text ------------
# Pay medians are starting values you can refresh from ASHE; text mirrors NCS profiles.
FIELD_META = {
 "health":("Health professionals & associate professionals",35000,
   "Hands-on work caring for patients and supporting clinical teams in hospitals, GP surgeries, care settings and the community — varied, people-facing and often shift-based.",
   "Enter as a healthcare assistant or support worker with no formal qualifications and train on the job; nursing needs a degree, increasingly available as a paid apprenticeship.",
   "One of the largest and most resilient fields of UK hiring."),
 "it":("Information technology professionals",62000,
   "Building, running and supporting the software and systems organisations rely on — from writing code to fixing laptops to analysing data, much of it remote or hybrid.",
   "Helpdesk and junior data roles take entry-level starters; bootcamps, free certifications and degree apprenticeships get people in without a computer-science degree.",
   "Highest-paying broad field and the most remote-friendly."),
 "sales":("Sales & customer service",32000,
   "Persuading customers and businesses to buy, and managing those relationships — fast-paced, target-driven and sociable, with earnings often tied to results.",
   "Most roles hire on attitude and communication rather than qualifications; you typically start as an executive or in retail sales and progress on performance.",
   "Hires constantly and rewards results over credentials."),
 "eng":("Science, engineering & technology professionals",45000,
   "Designing, building and maintaining physical things and systems — machines, buildings, vehicles, power — combining problem-solving with practical application.",
   "Engineering apprenticeships (level 3 up to degree level) are a strong debt-free route; technician roles also open up via BTEC and HNC qualifications.",
   "Spread across the country and tied to manufacturing and infrastructure."),
 "trade":("Skilled construction & building trades",36000,
   "Skilled hands-on building and installation work on sites and in homes — practical, physical and varied, often self-employed or on contract.",
   "Apprenticeships are the classic route; many also start labouring and gain trade qualifications (NVQs, CSCS card) while working.",
   "Skilled trades are in chronic short supply."),
 "edu":("Teaching & educational professionals",33000,
   "Helping people learn — from supporting children in classrooms to teaching subjects and running lessons — structured around the school year.",
   "Teaching-assistant roles need few formal qualifications; qualified teaching needs a degree plus PGCE or a salaried teaching apprenticeship.",
   "Strongly seasonal, peaking before the September term."),
 "hosp":("Caring, leisure & other service",26000,
   "Preparing and serving food and drink and looking after guests in restaurants, pubs, hotels and cafés — energetic, social, flexible-hours work.",
   "Most roles offer immediate starts with on-the-job training; chefs often progress through kitchen apprenticeships or college catering courses.",
   "Always hiring, with minimal barriers to entry."),
 "log":("Process, plant & machine operatives",28000,
   "Moving and storing goods — driving, picking and packing in warehouses, and coordinating deliveries — practical work that keeps supply chains running.",
   "Warehouse roles start immediately with no qualifications; HGV driving needs a licence (often employer-funded), with strong demand and pay.",
   "Concentrated around the Midlands distribution hubs."),
 "fin":("Business, finance & associate professionals",42000,
   "Recording, analysing and managing money for organisations — invoices, reporting, auditing, pricing risk — structured, detail-focused office work.",
   "Start as an accounts assistant and study AAT then ACCA or CIMA while working; finance and actuarial apprenticeships are widely available.",
   "Stable and structured, with clear qualification ladders."),
 "admin":("Administrative & secretarial",26000,
   "Keeping organisations running — scheduling, records, correspondence and coordination — broad, transferable office work that touches every department.",
   "Open to entry-level starters with good organisation and IT skills; business-administration apprenticeships are a common formal route.",
   "A broad first rung into many industries."),
 "retail":("Sales & customer service",23000,
   "Selling to and serving customers in shops, plus stock, displays and store running — accessible, people-facing work with lots of flexible hours.",
   "Sales-assistant roles need no qualifications and offer immediate starts; progression into supervisor and store manager is performance-based.",
   "Highly accessible and everywhere."),
 "cust":("Sales & customer service",25000,
   "Helping customers by phone, chat and email — answering questions, solving problems and handling claims and complaints — increasingly home-based.",
   "Hired on communication skills rather than qualifications, with full training provided; a strong stepping stone into sales or operations.",
   "A reliable, increasingly remote entry route."),
 "care":("Caring, leisure & other service",24000,
   "Supporting people who need help with daily life — older people, those with disabilities or illness — in their homes and in care settings.",
   "Very open entry valuing reliability and empathy; the Care Certificate and NVQs are gained on the job, with routes into management.",
   "Deep, growing demand and very open entry."),
 "mkt":("Business, finance & associate professionals",38000,
   "Promoting products and brands across digital channels, content, events and PR — creative, analytical and fast-moving work.",
   "Junior digital and content roles are reachable via short courses, platform certifications and a portfolio; marketing apprenticeships exist too.",
   "Lots of junior digital roles."),
 "mfg":("Process, plant & machine operatives",30000,
   "Making and assembling products on production lines and operating machinery — practical, often shift-based work with structured progression.",
   "Operative roles start immediately with on-the-job training; routes into supervision and skilled maintenance via NVQs and apprenticeships.",
   "Strong in the Midlands and North, with shift premiums."),
 "maint":("Skilled trades",32000,
   "Keeping buildings and grounds working and clean — repairs, facilities, caretaking and grounds work — practical and steady, needed everywhere.",
   "Cleaning and caretaking roles offer easy entry; technical facilities roles build on trade skills and short certifications.",
   "Practical, steady work needed everywhere."),
 "hr":("Business, finance & associate professionals",36000,
   "Finding, supporting and developing an organisation's people — recruiting, advising on policy, handling pay and training.",
   "Recruitment hires on drive and people skills with no set qualifications; HR roles build via CIPD qualifications, often started as an administrator.",
   "A well-trodden way into a professional office career."),
 "legal":("Professional occupations",44000,
   "Advising on and handling the law — drafting documents, managing cases, ensuring compliance — detail-heavy, research-led professional work.",
   "Paralegal and legal-secretary roles are accessible entry points; qualifying as a solicitor now has a degree route and a paid apprenticeship.",
   "Concentrated in London, but paralegal roles are realistic entry points."),
 "sci":("Science, engineering & technology professionals",38000,
   "Investigating, testing and developing through lab and field research — methodical, evidence-driven work across health, pharma, environment and industry.",
   "Lab-technician roles suit science graduates and apprentices; research roles typically build on a relevant degree, sometimes postgraduate.",
   "Clustered around research hubs."),
 "creative":("Professional occupations",34000,
   "Designing and producing visual and digital content — graphics, interfaces, video and imagery — blending creativity with software skills.",
   "Portfolio over paper: self-taught skills, short courses and bootcamps open doors, especially in UX/UI.",
   "Portfolio matters more than formal qualifications."),
 "energy":("Science, engineering & technology professionals",44000,
   "Generating and distributing power and running utilities — increasingly renewable — combining field engineering, technical and safety work.",
   "Engineering and technician apprenticeships are widely funded by employers; trade and electrical backgrounds transfer well into renewables.",
   "Fast-growing as renewables expand, strong in Scotland and the North East."),
}
FIELD_NAME = {
 "health":"Health & Nursing","it":"IT & Technology","sales":"Sales","eng":"Engineering",
 "trade":"Construction & Trades","edu":"Teaching & Education","hosp":"Hospitality & Catering",
 "log":"Transport & Logistics","fin":"Accounting & Finance","admin":"Admin & Secretarial",
 "retail":"Retail","cust":"Customer Service","care":"Social Care","mkt":"Marketing & PR",
 "mfg":"Manufacturing","maint":"Skilled Trades & Maintenance","hr":"HR & Recruitment",
 "legal":"Legal","sci":"Science & Research","creative":"Creative & Media","energy":"Energy & Utilities",
}
# regions + pay multiplier; region_share falls back to these population weights if the
# workbook has no usable geography breakdown.
REGIONS = [("all","All of the UK",1.0,1.00),("london","London",0.160,1.18),("southeast","South East",0.140,1.08),
 ("east","East of England",0.090,1.00),("southwest","South West",0.080,0.96),("westmidlands","West Midlands",0.085,0.95),
 ("eastmidlands","East Midlands",0.070,0.93),("yorkshire","Yorkshire & Humber",0.080,0.93),("northwest","North West",0.110,0.96),
 ("northeast","North East",0.040,0.90),("wales","Wales",0.045,0.91),("scotland","Scotland",0.085,0.98),("ni","Northern Ireland",0.025,0.90)]
REGION_NAME_MATCH = {  # match an ONS geography name to a region key
 "london":"london","south east":"southeast","east of england":"east","south west":"southwest",
 "west midlands":"westmidlands","east midlands":"eastmidlands","yorkshire":"yorkshire","north west":"northwest",
 "north east":"northeast","wales":"wales","scotland":"scotland","northern ireland":"ni"}
REG_KEYS=[r[0] for r in REGIONS if r[0]!="all"]
ASHE_UK_NAMES=("united kingdom","great britain","uk")
SYNONYMS = json.loads("""{"insurance":["insurance","underwriter","actuary","claims"],"coding":["developer","software","devops","engineer"],"software":["developer","software","qa"],"driving":["driver","hgv","delivery"],"driver":["driver","hgv","delivery"],"hgv":["driver","hgv"],"teaching":["teacher","teaching","tutor","lecturer","sen"],"nhs":["nurse","healthcare","clinical","health"],"nurse":["nurse","nursing"],"finance":["account","finance","financial","payroll","audit","bookkeep","actuary"],"accounting":["account","bookkeep","audit"],"law":["solicitor","legal","paralegal","conveyanc","compliance"],"legal":["solicitor","legal","paralegal"],"marketing":["marketing","seo","ppc","social","content"],"chef":["chef","kitchen","catering"],"catering":["chef","kitchen","barista","waiting","bar"],"warehouse":["warehouse","forklift","picker"],"logistics":["warehouse","logistics","driver","forklift"],"design":["design","ux","ui","graphic","artwork","motion"],"care":["care","carer","support worker"],"remote":["developer","data","customer","marketing","support"]}""")

def log(*a): print(*a, file=sys.stderr)

def latest_file_url():
    import requests
    log("Fetching dataset page…")
    html = requests.get(DATASET_PAGE, timeout=60, headers={"User-Agent":"open-jobs-map"}).text
    # editions are listed newest-first; grab the first .xlsx download link
    m = re.findall(r'href="(/file\?uri=[^"]+?\.xlsx)"', html)
    if not m:
        raise SystemExit("Could not find an .xlsx link on the dataset page — the layout may have changed.")
    url = urljoin("https://www.ons.gov.uk", m[0].replace("&amp;","&"))
    log("Latest file:", url)
    return url

def load_workbook_bytes(url):
    import requests
    log("Downloading (~45 MB)…")
    return requests.get(url, timeout=600, headers={"User-Agent":"open-jobs-map"}).content

def inspect(xbytes):
    import pandas as pd
    xl = pd.ExcelFile(io.BytesIO(xbytes))
    log("Sheets:", xl.sheet_names)
    for sh in xl.sheet_names:
        df = xl.parse(sh, nrows=6)
        log(f"\n--- {sh} ({df.shape[1]} cols) ---")
        log(list(df.columns)[:18])

def parse_month(col):
    """Return month index from Jan 2017, or None. Handles datetimes and 'YYYY MMM'/'MMM-YY' text."""
    import pandas as pd
    if isinstance(col, (datetime.datetime, datetime.date)):
        return (col.year-2017)*12 + (col.month-1)
    s = str(col).strip()
    try:
        ts = pd.to_datetime(s, errors="raise", dayfirst=True)
        return (ts.year-2017)*12 + (ts.month-1)
    except Exception:
        return None

def assign_field(name):
    n = (name or "").lower()
    for fk, kws in FIELD_KEYWORDS:
        if any(k in n for k in kws):
            return fk
    return None

def find_col(cols, needle):
    for c in cols:
        if needle.lower() in str(c).lower():
            return c
    return None

def build(xbytes, ashe=None):
    import pandas as pd
    xl = pd.ExcelFile(io.BytesIO(xbytes))
    sheet = CONFIG["sheet"] or max(xl.sheet_names, key=lambda s: xl.parse(s, nrows=0).shape[1])
    log("Using sheet:", sheet)
    df = xl.parse(sheet)
    df.columns = [str(c).strip() for c in df.columns]

    occ_col = find_col(df.columns, CONFIG["occupation_name_col"])
    geo_col = find_col(df.columns, CONFIG["geography_col"])
    if occ_col is None:
        raise SystemExit("Couldn't find the occupation-name column. Run --inspect and set CONFIG['occupation_name_col'].")

    # month columns = any header that parses as a date
    month_cols = [(c, parse_month(c)) for c in df.columns]
    month_cols = [(c,mi) for c,mi in month_cols if mi is not None and 0 <= mi]
    if not month_cols:
        raise SystemExit("Couldn't detect month columns. Run --inspect; the months may be in a long 'date' column instead of wide columns.")
    max_idx = max(mi for _,mi in month_cols)
    N = max_idx + 1
    log(f"Detected {len(month_cols)} month columns, span {N} months to index {max_idx}.")

    MONTH_NAMES=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    months=[f"{MONTH_NAMES[i%12]} {2017+i//12}" for i in range(N)]

    # accumulators
    field_series = {fk:[0.0]*N for fk in FIELD_NAME}
    field_region = {fk:{r[0]:0.0 for r in REGIONS if r[0]!="all"} for fk in FIELD_NAME}
    occ_latest = {fk:{} for fk in FIELD_NAME}      # occ name -> latest count
    occ_series = {fk:{} for fk in FIELD_NAME}      # occ name -> [N]
    occ_soc    = {fk:{} for fk in FIELD_NAME}      # occ name -> 4-digit SOC code
    soc_col = find_col(df.columns, CONFIG["soc_code_col"])

    for _, row in df.iterrows():
        fk = assign_field(row[occ_col])
        if fk is None:
            soc = str(row.get(find_col(df.columns, CONFIG["soc_code_col"]) or "", "")).strip()
            fk = SOC1_FALLBACK.get(soc[:1], None)
            if fk is None:
                continue
        # region (optional)
        rk = None
        if geo_col is not None:
            gname = str(row[geo_col]).lower()
            for frag,key in REGION_NAME_MATCH.items():
                if frag in gname: rk=key; break
        # add monthly values
        vals=[0.0]*N
        for c,mi in month_cols:
            v=row[c]
            try: v=float(v)
            except Exception: v=0.0
            if v!=v: v=0.0
            vals[mi]+=v
        for i in range(N): field_series[fk][i]+=vals[i]
        if rk: field_region[fk][rk]+=vals[max_idx]
        nm=str(row[occ_col]).strip()
        occ_series[fk][nm]=[occ_series[fk].get(nm,[0.0]*N)[i]+vals[i] for i in range(N)] if nm in occ_series[fk] else vals
        occ_latest[fk][nm]=occ_latest[fk].get(nm,0)+vals[max_idx]
        if soc_col is not None and nm not in occ_soc[fk]:
            m4=re.search(r"\d{4}",str(row[soc_col]))
            if m4: occ_soc[fk][nm]=m4.group(0)

    # assemble fields
    fields=[]
    have_region = any(sum(field_region[fk].values())>0 for fk in FIELD_NAME)
    def role_pay(soc4, field_median):
        a=ashe.get(soc4,{}) if ashe else {}
        nat=a.get("all") or field_median
        reg={rk:(a.get(rk) if a.get(rk) else nat) for rk in REG_KEYS}
        return round(nat), {rk:round(reg[rk]) for rk in REG_KEYS}
    for fk in FIELD_NAME:
        s=[round(x,1) for x in field_series[fk]]
        if sum(s)==0: continue
        group,salary,about,routes,note = FIELD_META[fk]
        # region share: real if available, else population fallback
        if have_region and sum(field_region[fk].values())>0:
            tot=sum(field_region[fk].values())
            rshare={k:round(v/tot,4) for k,v in field_region[fk].items()}
        else:
            pop={r[0]:r[2] for r in REGIONS if r[0]!="all"}; tp=sum(pop.values())
            rshare={k:round(v/tp,4) for k,v in pop.items()}
        # top roles within the field, each with its own ASHE median per region
        top=sorted(occ_latest[fk].items(), key=lambda kv: kv[1], reverse=True)[:8]
        roles=[]; wsum=0.0; fnat=0.0; freg={rk:0.0 for rk in REG_KEYS}
        for nm,_ in top:
            w=occ_latest[fk][nm]
            if w<=0: continue
            nat,reg=role_pay(occ_soc[fk].get(nm,""), salary)
            wsum+=w; fnat+=nat*w
            for rk in REG_KEYS: freg[rk]+=reg[rk]*w
            roles.append({"name":nm,"salary":nat,"salary_region":reg,"series":[round(x,1) for x in occ_series[fk][nm]]})
        field_nat = round(fnat/wsum) if wsum else salary
        field_reg = {rk:(round(freg[rk]/wsum) if wsum else salary) for rk in REG_KEYS}
        fields.append({"k":fk,"name":FIELD_NAME[fk],"group":group,"salary":field_nat,"salary_region":field_reg,
                       "region_share":rshare,"series":s,"about":about,"routes":routes,"note":note,"roles":roles})

    fields.sort(key=lambda f: f["series"][-1], reverse=True)
    pay_src = (f"live ASHE Table 15 medians for {len(ashe)} occupations (region x 4-digit SOC)"
               if ashe else "bundled ASHE-derived field medians (run with --ashe to use live ASHE)")
    data={"meta":{"latest_month":months[-1],"months":months,
                  "source":f"ONS Labour demand volumes by SOC 2020 (online job adverts, Textkernel); pay: {pay_src}; profiles from National Careers Service",
                  "licence":"Open Government Licence v3.0 — source: Office for National Statistics",
                  "generated":datetime.date.today().isoformat(),"is_sample":False},
          "regions":[{"k":r[0],"name":r[1],"sal":r[3]} for r in REGIONS],
          "fields":fields,"synonyms":SYNONYMS}
    if not have_region:
        log("NOTE: no usable region breakdown found in the workbook; region_share fell back to population weights. "
            "If the file has local-authority rows, add a local-authority→region lookup to make 'Where' fully live.")
    return data

def ashe_latest_zip_url():
    import requests
    html=requests.get(ASHE_PAGE,timeout=60,headers={"User-Agent":"open-jobs-map"}).text
    m=re.findall(r'href="(/file\?uri=[^"]+?\.zip)"',html)
    if not m: raise SystemExit("Could not find an ASHE .zip link — the page layout may have changed.")
    url=urljoin("https://www.ons.gov.uk",m[0].replace("&amp;","&"))
    log("Latest ASHE file:",url); return url

def _region_key_from(text):
    t=str(text).lower()
    if any(u in t for u in ASHE_UK_NAMES): return "all"
    for frag,key in REGION_NAME_MATCH.items():
        if frag in t: return key
    return None

def load_ashe_csv(path):
    """RECOMMENDED, most reliable route. CSV with a SOC-code column, a region column
       (region name or 'United Kingdom'), and a median annual-pay column. Easy to export
       from NOMIS (dataset 'ASHE - occupation (4 digit SOC)') or the ONS filter tool."""
    import pandas as pd
    df=pd.read_csv(path); df.columns=[str(c).strip().lower() for c in df.columns]
    soc_c=next((c for c in df.columns if "soc" in c or "occupation code" in c),None)
    reg_c=next((c for c in df.columns if "region" in c or "geograph" in c or "area" in c),None)
    val_c=next((c for c in df.columns if "median" in c),None) or next((c for c in df.columns if "value" in c or "pay" in c),None)
    if not(soc_c and val_c): raise SystemExit("ASHE CSV needs at least a SOC column and a median column.")
    out={}
    for _,r in df.iterrows():
        m4=re.search(r"\d{4}",str(r[soc_c]))
        if not m4: continue
        try: v=float(str(r[val_c]).replace(",","").replace("£",""))
        except Exception: continue
        if v<=0: continue
        rk=_region_key_from(r[reg_c]) if reg_c else "all"
        if rk is None: continue
        out.setdefault(m4.group(0),{})[rk]=v
    log(f"ASHE (csv): medians for {len(out)} occupations across regions.")
    return out

def _ashe_pick_workbook(names):
    cand=[n for n in names if n.lower().endswith((".xlsx",".xls"))]
    def ok(n):
        nl=n.lower()
        return all(k in nl for k in CONFIG["ashe_file_match"]) and not any(x in nl for x in CONFIG["ashe_file_exclude"])
    hits=[n for n in cand if ok(n)]
    return hits[0] if hits else (cand[0] if cand else None)

def load_ashe_zip(zbytes):
    """Best-effort parse of the ASHE Table 15 zip -> {soc4: {'all': median, region: median}}.
       ASHE workbook layouts vary year to year; this reliably gets NATIONAL medians. For full
       per-region medians the robust path is --ashe-csv (see README). Falls back silently."""
    import zipfile,io,pandas as pd
    try: z=zipfile.ZipFile(io.BytesIO(zbytes))
    except Exception as e: log("ASHE: not a readable zip:",e); return {}
    wb=_ashe_pick_workbook(z.namelist())
    if not wb: log("ASHE: no workbook found in zip."); return {}
    log("ASHE workbook:",wb)
    try: xl=pd.ExcelFile(io.BytesIO(z.read(wb)))
    except Exception as e: log("ASHE: cannot open workbook:",e); return {}
    sheet=next((s for s in xl.sheet_names if CONFIG["ashe_stat_sheet"].lower() in s.lower()), xl.sheet_names[0])
    raw=xl.parse(sheet,header=None)
    hdr=None
    for i in range(min(12,len(raw))):
        joined=" ".join(str(x).lower() for x in raw.iloc[i].tolist())
        if "description" in joined or "soc" in joined or ("code" in joined and "median" in joined):
            hdr=i; break
    df=xl.parse(sheet,header=hdr if hdr is not None else 0); df.columns=[str(c).strip() for c in df.columns]
    code_c=find_col(df.columns,"code") or find_col(df.columns,"soc")
    med_c =find_col(df.columns,"median") or find_col(df.columns,"value")
    out={}
    if code_c and med_c:
        for _,r in df.iterrows():
            m4=re.search(r"\d{4}",str(r[code_c]))
            if not m4: continue
            try: v=float(str(r[med_c]).replace(",","").replace("£",""))
            except Exception: continue
            if v>0: out.setdefault(m4.group(0),{})["all"]=v
    if out: log(f"ASHE (zip): national medians for {len(out)} occupations. "
                "(Per-region medians are not auto-parsed from the zip — use --ashe-csv for those.)")
    else:   log("ASHE: automatic parse found nothing; run --inspect-ashe or use --ashe-csv.")
    return out

def inspect_ashe(zbytes):
    import zipfile,io,pandas as pd
    z=zipfile.ZipFile(io.BytesIO(zbytes))
    log("Files in ASHE zip:");  [log("  ",n) for n in z.namelist()]
    wb=_ashe_pick_workbook(z.namelist())
    if wb:
        log("\nPicked workbook:",wb)
        xl=pd.ExcelFile(io.BytesIO(z.read(wb))); log("Sheets:",xl.sheet_names)
        s=next((x for x in xl.sheet_names if "median" in x.lower()),xl.sheet_names[0])
        log(f"\nFirst rows of '{s}':"); log(xl.parse(s,header=None,nrows=8).to_string())

def load_ashe(args):
    if args.no_ashe: return None
    if args.ashe_csv: return load_ashe_csv(args.ashe_csv)
    import requests
    if args.ashe_file: zb=open(args.ashe_file,"rb").read()
    else:
        log("Downloading ASHE Table 15 (~80 MB)…")
        zb=requests.get(ashe_latest_zip_url(),timeout=900,headers={"User-Agent":"open-jobs-map"}).content
    return load_ashe_zip(zb)

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("--inspect",action="store_true",help="print the labour-demand workbook's sheets/columns and exit")
    ap.add_argument("--inspect-ashe",dest="inspect_ashe",action="store_true",help="print the ASHE zip's files/sheets and exit")
    ap.add_argument("--file",help="use a local labour-demand .xlsx instead of downloading")
    ap.add_argument("--ashe-csv",dest="ashe_csv",help="path to a SOC,region,median CSV (recommended for per-region pay)")
    ap.add_argument("--ashe-file",dest="ashe_file",help="use a local ASHE Table 15 .zip instead of downloading")
    ap.add_argument("--no-ashe",dest="no_ashe",action="store_true",help="skip ASHE; use bundled field medians")
    ap.add_argument("--out",default="data.json")
    a=ap.parse_args()
    if a.inspect_ashe:
        import requests
        zb=open(a.ashe_file,"rb").read() if a.ashe_file else requests.get(ashe_latest_zip_url(),timeout=900,headers={"User-Agent":"open-jobs-map"}).content
        inspect_ashe(zb); return
    xbytes = open(a.file,"rb").read() if a.file else load_workbook_bytes(latest_file_url())
    if a.inspect:
        inspect(xbytes); return
    ashe=load_ashe(a)
    data=build(xbytes, ashe)
    with open(a.out,"w") as f: json.dump(data,f,separators=(",",":"))
    log(f"Wrote {a.out}: {len(data['fields'])} fields, {sum(len(x['roles']) for x in data['fields'])} roles, "
        f"{len(data['meta']['months'])} months to {data['meta']['latest_month']}.")

if __name__=="__main__":
    main()
