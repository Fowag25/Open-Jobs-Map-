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

    # ---- ASHE Table 15 (pay) selection inside the downloaded zip ----
    # Confirm with: python3 build_data.py --inspect-ashe
    "ashe_sheet": "All",                     # worksheet for All employees (not Male/Female/Full/Part)
}

# ---- 23 fields: ordered rules to assign each ONS occupation to a field -------------
# First field whose keyword matches the occupation name wins, so order = priority.
# Specific overrides are placed FIRST so they beat the broad rules below them.
FIELD_KEYWORDS = [
 ("data",     ["data analyst","data scientist","data engineer","data science","business intelligence","market research","statistician","statistical","insight analyst"]),
 ("consulting",["management consultant","business analyst","regulatory affair","regulatory profession","regulatory adviser","quality assurance and regulatory","compliance officer","compliance manager","compliance profession","change manager","business adviser"]),
 ("care",     ["social worker","social work","care worker","carer","care assistant","senior care","home care","support worker","residential care","care home","nursing auxiliar"]),
 ("hosp",     ["chef","cook","waiter","waitress","waiting staff","bar staff","bar person","catering","kitchen","hospitality","publican","barista","restaurant","hairdress","barber","beautician","beauty","nail technician","fitness","sports coach","leisure","spa "]),
 ("edu",      ["teacher","teaching assistant","teaching","lecturer","education","tutor","classroom","nursery","early years","vocational and industrial trainer","industrial trainer","vocational trainer","further education"]),
 ("maint",    ["servicer","repairer","tv, video","television, video","aerial","security guard","security officer","door supervisor","maintenance","facilities","cleaner","caretaker","janitor","groundskeep","gardener","handyman","grounds"]),
 ("trade",    ["electrician","plumber","carpenter","joiner","bricklay","plasterer","construction","scaffold","roofer","glazier","painter and decorat","steel erect","groundwork","estimator","valuer","valuers and assessors","quantity surveyor","site manager","labourer"]),
 ("fin",      ["accountant","accounting","accounts assistant","accounts clerk","bookkeep","payroll","audit","actuar","taxation","tax adviser","tax expert","financial analyst","finance manager","finance officer","financial adviser","insurance","underwrit"]),
 ("it",       ["programmer","software","web design","web develop","information technology","it support","it user","it technician","it operations","cyber security","information security","devops","systems admin","network engineer","network admin","it project","database admin","cloud engineer","it security"]),
 ("retail",   ["retail","sales assistant","retail cashier","cashier","checkout","shelf","store assistant"]),
 ("sales",    ["sales","business development","account manager","key account","telesales","estate agent"]),
 ("eng",      ["engineer","engineering technician","draughts","cad technician"]),
 ("health",   ["nurse","nursing","midwif","health profession","paramedic","physiotherap","radiograph","occupational therap","clinical","doctor","medical practitioner","healthcare assistant","dental","pharmacist","health associate"]),
 ("legal",    ["solicitor","barrister","legal","paralegal","conveyanc"]),
 ("sci",      ["research scientist","scientist","laborator","lab technician","research and development","chemist","biologist","physicist","biochemist","pharmacolog","microbiolog","ecologist","geologist"]),
 ("mkt",      ["marketing","advertising","public relations","seo","brand manager","social media","copywriter"]),
 ("hr",       ["human resource","recruitment","personnel","talent acquisition","learning and develop"]),
 ("creative", ["graphic design","ux","ui design","interior design","artist","photograph","video editor","videograph","motion","animator","designer","creative","journalist","editor","broadcast"]),
 ("energy",   ["energy","renewable","wind turbine","wind farm","solar","nuclear","power station","power plant","power network","electricity","national grid","gas network","gas distribution","natural gas","oil and gas","petroleum","water treatment","water and sewerage","sewerage","utilities","drilling","pipeline"]),
 ("mfg",      ["production","manufactur","assembl","machine operat","factory","plant operat","fabricat","process operat","quality control","quality inspector","quality assurance technician"]),
 ("log",      ["driver","hgv","lgv","delivery","warehouse","forklift","logistics","transport","postal","courier"]),
 ("cust",     ["customer service","customer support","call centre","contact centre","claims handler","complaints"]),
 ("admin",    ["administ","secretar","receptionist","clerk","data entry","office manager","personal assistant","typist","records"]),
]
# fallbacks if no keyword matched: SOC2020 sub-major (2-digit) then major (1-digit)
SOC2_FALLBACK = {"11":"consulting","12":"admin","13":"consulting","21":"sci","22":"health","23":"edu","24":"fin",
 "25":"it","26":"care","31":"sci","32":"health","33":"care","34":"creative","35":"fin","41":"admin","42":"admin",
 "51":"trade","52":"maint","53":"trade","54":"mfg","55":"hosp","61":"care","62":"hosp","71":"retail","72":"cust",
 "81":"mfg","82":"log","91":"maint","92":"admin"}
SOC1_FALLBACK = {"1":"admin","2":"sci","3":"admin","4":"admin","5":"trade","6":"care","7":"sales","8":"mfg","9":"maint"}

# ---- field metadata: group label, fallback median pay, careers text (mirrors NCS) ----
FIELD_META = {
 "health":("Health professionals and associate professionals",35000,
   "Hands-on work caring for patients and supporting clinical teams in hospitals, GP surgeries, care settings and the community.",
   "Enter as a healthcare assistant or support worker with no formal qualifications and train on the job; nursing needs a degree, increasingly available as a paid apprenticeship.",
   "Among the largest and most resilient fields."),
 "care":("Caring, leisure and other service",24000,
   "Supporting people who need help with daily life — older people, those with disabilities or illness — in their homes and in care settings.",
   "Very open entry valuing reliability and empathy; the Care Certificate and NVQs are gained on the job. Social work requires a degree.",
   "Deep, growing demand and very open entry."),
 "it":("Information technology professionals",60000,
   "Building, running and supporting the software and systems organisations rely on — much of it remote or hybrid.",
   "Helpdesk roles take entry-level starters; bootcamps, certifications and degree apprenticeships get people in without a computer-science degree.",
   "Highly remote-friendly, with a sharp 2021-22 hiring surge that has since cooled."),
 "data":("Information technology professionals",48000,
   "Finding patterns in data and turning them into decisions — from dashboards to forecasting and research.",
   "Junior analyst roles are reachable through short courses, certifications and a portfolio of projects; many come from numerate degrees.",
   "One of the fastest-growing fields, spread across every industry."),
 "eng":("Science, engineering and technology professionals",45000,
   "Designing, building and maintaining physical things and systems — machines, buildings, vehicles, power.",
   "Engineering apprenticeships (level 3 up to degree level) are a strong debt-free route; technician roles open up via BTEC and HNC qualifications.",
   "Spread across the country and tied to manufacturing and infrastructure."),
 "trade":("Skilled construction and building trades",36000,
   "Skilled hands-on building and installation work on sites and in homes — practical, physical and varied.",
   "Apprenticeships are the classic route; many also start labouring and gain trade qualifications (NVQs, CSCS card) while working.",
   "Skilled trades are in chronic short supply, so qualified tradespeople have real bargaining power."),
 "edu":("Teaching and educational professionals",33000,
   "Helping people learn — from supporting children in classrooms to teaching subjects and training adults.",
   "Teaching-assistant roles need few formal qualifications; qualified teaching needs a degree plus PGCE or a salaried teaching apprenticeship.",
   "Strongly seasonal, peaking before the September term."),
 "hosp":("Caring, leisure and other service",25000,
   "Preparing and serving food and drink, running venues, and personal-care and leisure services — social, flexible-hours work.",
   "Most roles offer immediate starts with on-the-job training; chefs, hairdressers and beauticians often train through apprenticeships or college courses.",
   "Crashed hardest in 2020, then rebounded sharply as venues reopened."),
 "log":("Process, plant and machine operatives",28000,
   "Moving and storing goods — driving, picking and packing, and coordinating deliveries.",
   "Warehouse roles start immediately with no qualifications; HGV driving needs a licence, often employer-funded.",
   "Surged through the pandemic as online retail boomed, especially driving roles."),
 "fin":("Business, finance and associate professionals",42000,
   "Recording, analysing and managing money and risk for organisations — including insurance.",
   "Start as an accounts assistant and study AAT then ACCA or CIMA while working; underwriting and actuarial work have their own qualification routes.",
   "Stable and structured, with clear qualification ladders."),
 "admin":("Administrative and secretarial",26000,
   "Keeping organisations running — scheduling, records, correspondence and coordination.",
   "Open to entry-level starters with good organisation and IT skills; business-administration apprenticeships are a common formal route.",
   "A broad first rung that touches every department, in slow long-term decline as tasks automate."),
 "retail":("Sales and customer service",23000,
   "Selling to and serving customers in shops, plus stock, displays and store running.",
   "Sales-assistant roles need no qualifications and offer immediate starts; progression into supervisor and store manager is performance-based.",
   "Highly accessible and everywhere, with lots of part-time hours."),
 "cust":("Sales and customer service",25000,
   "Helping customers by phone, chat and email — answering questions, solving problems and handling complaints.",
   "Hired on communication skills rather than qualifications, with full training provided; a strong stepping stone into sales or operations.",
   "A reliable, increasingly home-based entry route."),
 "mkt":("Business, finance and associate professionals",38000,
   "Promoting products and brands across digital channels, content, events and PR.",
   "Junior digital and content roles are reachable via short courses, platform certifications and a portfolio; marketing apprenticeships exist too.",
   "Creative and fast-moving, with lots of junior digital roles."),
 "mfg":("Process, plant and machine operatives",30000,
   "Making and assembling products on production lines and operating machinery — often shift-based.",
   "Operative roles start immediately with on-the-job training; routes into supervision and skilled maintenance via NVQs and apprenticeships.",
   "Strong in the Midlands and North, with shift premiums."),
 "maint":("Skilled trades",30000,
   "Keeping buildings, grounds and equipment working, clean and secure — practical, steady work needed everywhere.",
   "Cleaning, caretaking and security roles offer easy entry; technical and repair roles build on trade skills and short certifications.",
   "Practical, steady work needed everywhere."),
 "hr":("Business, finance and associate professionals",36000,
   "Finding, supporting and developing an organisation's people — recruiting, advising on policy, handling pay and training.",
   "Recruitment hires on drive and people skills with no set qualifications; HR roles build via CIPD qualifications, often started as an administrator.",
   "A well-trodden way into a professional office career."),
 "legal":("Professional occupations",44000,
   "Advising on and handling the law — drafting documents, managing cases and ensuring compliance.",
   "Paralegal and legal-secretary roles are accessible entry points; qualifying as a solicitor now has a degree route and a paid apprenticeship.",
   "Concentrated in London, but paralegal roles are realistic entry points."),
 "sci":("Science, engineering and technology professionals",38000,
   "Investigating, testing and developing through lab and field research — methodical, evidence-driven work.",
   "Laboratory-technician roles suit science graduates and apprentices; research roles typically build on a relevant degree, sometimes postgraduate.",
   "Clustered around research hubs like Cambridge, Oxford and the North West."),
 "creative":("Professional occupations",34000,
   "Designing and producing visual and digital content — graphics, interfaces, interiors, video and imagery.",
   "Portfolio over paper: self-taught skills, short courses and bootcamps open doors, especially in UX/UI and interior design.",
   "Portfolio matters more than formal qualifications."),
 "consulting":("Business, finance and associate professionals",52000,
   "Advising organisations on how to work better, stay compliant and manage change — analytical, client-facing work.",
   "Most consultants come from a relevant degree or industry experience; graduate schemes and analyst roles are common entry points.",
   "London-heavy and steadily growing across regulation and transformation work."),
 "sales":("Sales and customer service",34000,
   "Persuading customers and businesses to buy, and managing those relationships — target-driven and sociable.",
   "Most roles hire on attitude and communication rather than qualifications; you typically start as an executive and progress on performance.",
   "Hires constantly and rewards results over credentials."),
 "energy":("Science, engineering and technology professionals",44000,
   "Generating and distributing power and running utilities — increasingly renewable.",
   "Engineering and technician apprenticeships are widely funded by employers; trade and electrical backgrounds transfer well into renewables.",
   "Fast-growing as renewables expand, strong in Scotland and the North East."),
}
FIELD_NAME = {
 "health":"Health and Nursing","care":"Social Care","it":"IT and Technology","data":"Data and Analytics",
 "eng":"Engineering","trade":"Construction and Trades","edu":"Teaching and Education",
 "hosp":"Hospitality, Leisure and Personal Care","log":"Transport and Logistics",
 "fin":"Finance, Accounting and Insurance","admin":"Admin and Secretarial","retail":"Retail",
 "cust":"Customer Service","mkt":"Marketing and PR","mfg":"Manufacturing",
 "maint":"Skilled Trades & Maintenance","hr":"HR and Recruitment","legal":"Legal",
 "sci":"Science and Research","creative":"Creative and Media","consulting":"Consulting",
 "sales":"Sales","energy":"Energy and Utilities",
}
# ---- per-role descriptions (Option A): attach a 1-2 sentence line by keyword ---------
DESC_RULES = [
 (["programmer","software develop","software engineer"],"Designs, writes and tests the code behind applications, websites and systems."),
 (["it support","it user","help desk","helpdesk","it technician"],"First port of call for technical problems, fixing hardware, software and access issues for users."),
 (["devops"],"Automates and maintains the infrastructure that lets software be built, released and run reliably."),
 (["cyber","information security","it security"],"Protects systems and data from attack, monitoring threats and responding to incidents."),
 (["data analyst"],"Turns raw data into clear insights and dashboards that help an organisation make decisions."),
 (["data scientist"],"Uses statistics and machine learning to model data and predict outcomes."),
 (["business intelligence"],"Builds reports and metrics that track how a business is performing."),
 (["market research"],"Gathers and interprets data on customers and markets to guide strategy and marketing."),
 (["data engineer"],"Builds and maintains the pipelines that collect, clean and store data for analysis."),
 (["statistic"],"Designs studies and applies statistical methods to draw reliable conclusions from data."),
 (["mechanical engineer"],"Designs, develops and maintains mechanical systems, machines and products."),
 (["electrical engineer"],"Designs and oversees electrical systems and equipment."),
 (["civil engineer"],"Plans and supervises construction of infrastructure such as roads and bridges."),
 (["maintenance engineer"],"Keeps plant and machinery running, diagnosing faults and carrying out repairs."),
 (["renewable","wind turbine","solar","power station","power plant","national grid","substation","gas network","water treatment","sewerage","energy plant","pipeline"],"Helps generate, distribute or maintain power, water or other utilities."),
 (["estate agent"],"Helps people buy, sell and rent property, valuing homes and arranging viewings."),
 (["conveyanc"],"Handles the legal side of buying and selling property."),
 (["registered nurse","nurse","nursing"],"Plans and delivers clinical care to patients, monitoring conditions and giving treatment."),
 (["healthcare assistant","nursing auxiliar"],"Supports nurses and doctors with day-to-day patient care such as washing, feeding and observations."),
 (["care worker","carer","care assistant"],"Supports older or disabled people with everyday tasks like washing, dressing, meals and medication."),
 (["support worker"],"Helps people live as independently as possible, providing practical and emotional support."),
 (["social worker"],"Assesses and safeguards vulnerable children or adults and arranges the support they need."),
 (["bookkeep","payroll"],"Keeps financial records accurate and makes sure people and bills are paid correctly."),
 (["accountant","accounts","accounting"],"Records and reports on an organisation's finances, from bookkeeping to management accounts."),
 (["underwrit"],"Assesses risk and decides the terms and price of insurance cover."),
 (["insurance"],"Works in insurance — assessing risk, pricing cover or handling claims."),
 (["actuar"],"Uses statistics to measure and price long-term financial risk."),
 (["tax"],"Advises on tax, prepares returns and helps clients stay compliant and efficient."),
 (["audit"],"Checks that financial records are accurate and meet the rules."),
 (["electrician"],"Installs, tests and maintains electrical wiring and equipment in buildings."),
 (["plumber"],"Installs and repairs water, heating and drainage systems."),
 (["carpenter","joiner"],"Makes and fits wooden structures and fittings, from frames and floors to staircases."),
 (["estimator","valuer"],"Works out the likely cost or value of building work, materials or property."),
 (["labourer"],"Provides general support on site, moving materials and assisting skilled trades."),
 (["site manager"],"Runs a construction site day to day, overseeing safety, progress and trades."),
 (["bricklay","plasterer"],"Builds and finishes walls and surfaces using brick, block and plaster."),
 (["chef","cook"],"Prepares and cooks food to order in a professional kitchen, managing quality and timing."),
 (["waiter","waitress","waiting"],"Greets and serves customers, taking orders and looking after their experience."),
 (["bar staff","bar person","bartender"],"Serves drinks, handles payments and keeps the bar running and stocked."),
 (["hairdress","barber"],"Cuts, colours and styles hair and advises clients on care and looks."),
 (["beautician","beauty","nail technician"],"Provides beauty treatments such as facials, nails and waxing."),
 (["fitness","sports coach"],"Leads exercise or coaching sessions and helps people train safely towards their goals."),
 (["teaching assistant"],"Supports teachers in the classroom and helps individual pupils with their learning."),
 (["primary"],"Plans and teaches across the curriculum to children in primary school."),
 (["secondary"],"Teaches a specialist subject to older pupils and prepares them for exams."),
 (["vocational","industrial trainer","trainer"],"Trains people in practical, job-specific skills in workplaces or colleges."),
 (["lecturer"],"Teaches and assesses students in further or higher education."),
 (["teacher"],"Plans and delivers lessons, assessing progress and supporting pupils' learning."),
 (["warehouse"],"Picks, packs and moves stock in a warehouse, keeping orders flowing."),
 (["hgv","lgv","lorry"],"Drives large goods vehicles to deliver freight across the country."),
 (["delivery","courier","van driver"],"Delivers parcels and goods to homes and businesses on a set route."),
 (["forklift"],"Operates a forklift to load, unload and move palletised goods safely."),
 (["administ"],"Keeps an office running with scheduling, records, correspondence and data."),
 (["receptionist"],"The first point of contact for visitors and callers, handling enquiries and bookings."),
 (["personal assistant","executive assistant"],"Provides high-level support to a senior leader, managing their diary and priorities."),
 (["research scientist","research and development"],"Designs and runs experiments to develop new knowledge or products."),
 (["laborator","lab technician"],"Prepares, runs and records lab tests and keeps equipment and samples in order."),
 (["chemist"],"Studies substances and their reactions to develop or test materials and products."),
 (["biologist"],"Studies living organisms in the lab or field to support health, environment or industry."),
 (["graphic design"],"Creates visual designs for print and screen — logos, layouts, branding and marketing."),
 (["ux","ui design"],"Designs how digital products look and feel so they are easy and pleasant to use."),
 (["interior design"],"Plans and designs indoor spaces, balancing function, style and budget."),
 (["photograph"],"Takes and edits images for clients, publications, products or events."),
 (["video","motion","animator"],"Produces and edits video or animation for film, marketing and social media."),
 (["journalist","editor"],"Researches, writes and edits content for publication or broadcast."),
 (["management consultant"],"Advises organisations on how to improve performance, strategy and operations."),
 (["regulatory"],"Guides organisations through laws and regulations so products and practices stay compliant."),
 (["quality assurance"],"Sets and checks standards so products and services consistently meet requirements."),
 (["business analyst"],"Bridges business needs and solutions, mapping processes and defining requirements."),
 (["compliance"],"Helps a business follow the rules that apply to it and manage regulatory risk."),
 (["change manager"],"Plans and guides organisations through major changes so they actually stick."),
 (["business development","account manager","sales"],"Wins and manages customer relationships, working to sales targets."),
 (["solicitor"],"Advises clients on the law and represents them in legal matters."),
 (["paralegal","legal assistant","legal secretary"],"Supports lawyers by preparing documents, researching and managing cases."),
 (["recruitment","talent"],"Finds and places candidates into jobs, matching people to employers."),
 (["human resource"],"Supports an organisation's people — pay, policy, hiring and wellbeing."),
 (["marketing"],"Promotes products and brands across channels, content and campaigns."),
 (["production","machine operat","assembl"],"Operates machinery and assembles products on a production line."),
 (["engineer"],"Designs, builds or maintains technical systems and equipment."),
 (["quality"],"Checks that products meet the required standard, finding and flagging faults."),
 (["security"],"Protects people, premises and property, monitoring and deterring risks."),
 (["servicer","repairer"],"Installs, services and repairs equipment, diagnosing faults and putting them right."),
 (["cleaner"],"Keeps premises clean, hygienic and presentable."),
 (["maintenance","facilities","caretaker"],"Keeps buildings and equipment working, carrying out repairs and upkeep."),
 (["customer service","customer support","call cent","contact cent"],"Helps customers by phone, chat and email, answering questions and solving problems."),
 (["sales assistant","retail","cashier","checkout"],"Serves customers in store, handles payments and keeps the shop running."),
]
def clean_name(name):
    n=re.sub(r"\s*\bn\.?e\.?c\.?\b\.?","",str(name),flags=re.I)      # drop 'n.e.c.'
    n=n.replace(" & "," and ").replace("&","and")
    return re.sub(r"\s{2,}"," ",n).strip(" ,;:-")
def desc_for(name):
    n=str(name).lower()
    for kws,d in DESC_RULES:
        if any(k in n for k in kws): return d
    return None
FIELD_DESC_FALLBACK={
 "health":"A healthcare role caring for patients and supporting clinical teams.",
 "care":"A social-care role supporting people who need help with daily life.",
 "it":"A technology role building, running or supporting software and systems.",
 "data":"A data role turning information into insight and decisions.",
 "eng":"An engineering role designing, building or maintaining systems and equipment.",
 "trade":"A construction or skilled-trade role doing hands-on building and installation work.",
 "edu":"An education role helping people learn and develop.",
 "hosp":"A hospitality, leisure or personal-care role serving and looking after customers.",
 "log":"A transport or logistics role moving and storing goods.",
 "fin":"A finance, accounting or insurance role managing money and risk.",
 "admin":"An administrative role keeping an organisation organised and running.",
 "retail":"A retail role selling to and serving customers in store.",
 "cust":"A customer-service role helping customers and resolving issues.",
 "mkt":"A marketing or PR role promoting products, brands and content.",
 "mfg":"A manufacturing role making, assembling or checking products.",
 "maint":"A maintenance or facilities role keeping buildings and equipment working and secure.",
 "hr":"An HR or recruitment role finding, supporting and developing people.",
 "legal":"A legal role advising on and handling the law.",
 "sci":"A science or research role investigating, testing and developing.",
 "creative":"A creative or media role designing and producing content.",
 "consulting":"A consulting role advising organisations on how to work better.",
 "sales":"A sales role winning and managing customer relationships.",
 "energy":"An energy or utilities role generating, distributing or maintaining power and services.",
}
# regions + pay multiplier; region_share falls back to these population weights if the
# workbook has no usable geography breakdown.
REGIONS = [("all","All of the UK",1.0,1.00),("london","London",0.160,1.18),("southeast","South East",0.140,1.08),
 ("east","East of England",0.090,1.00),("southwest","South West",0.080,0.96),("westmidlands","West Midlands",0.085,0.95),
 ("eastmidlands","East Midlands",0.070,0.93),("yorkshire","Yorkshire and The Humber",0.080,0.93),("northwest","North West",0.110,0.96),
 ("northeast","North East",0.040,0.90),("wales","Wales",0.045,0.91),("scotland","Scotland",0.085,0.98),("ni","Northern Ireland",0.025,0.90)]
REGION_NAME_MATCH = {  # match an ONS geography name to a region key
 "london":"london","south east":"southeast","east of england":"east","south west":"southwest",
 "west midlands":"westmidlands","east midlands":"eastmidlands","yorkshire":"yorkshire","north west":"northwest",
 "north east":"northeast","wales":"wales","scotland":"scotland","northern ireland":"ni"}
REG_KEYS=[r[0] for r in REGIONS if r[0]!="all"]
ASHE_UK_NAMES=("united kingdom","great britain","uk")
# ASHE Table 15 prefixes each row's Description with the region; match those spellings exactly.
ASHE_REGION={"north east":"northeast","north west":"northwest","yorkshire and the humber":"yorkshire",
 "east midlands":"eastmidlands","west midlands":"westmidlands","east":"east","east of england":"east",
 "london":"london","south east":"southeast","south west":"southwest","wales":"wales","scotland":"scotland",
 "northern ireland":"ni","united kingdom":"all","great britain":"all"}
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
    """Return month index from Jan 2017, or None. Handles real datetimes and ONS text like 'Jan-17'."""
    import pandas as pd
    def idx(y,m):
        if y<2017 or y>2035: return None
        return (y-2017)*12 + (m-1)
    if isinstance(col, (datetime.datetime, datetime.date)):
        return idx(col.year, col.month)
    s = str(col).strip()
    if not s or s.lower()=="nan": return None
    for fmt in ("%b-%y","%b-%Y","%B-%y","%B-%Y","%b %y","%b %Y","%B %Y","%Y-%m","%Y %b","%Y %B","%b%y"):
        try:
            d=datetime.datetime.strptime(s,fmt); return idx(d.year,d.month)
        except Exception: pass
    try:
        ts = pd.to_datetime(s, errors="raise")
        if pd.isna(ts): return None
        return idx(ts.year, ts.month)
    except Exception:
        return None

def assign_field(name, soc=""):
    n = (name or "").lower()
    for fk, kws in FIELD_KEYWORDS:
        if any(k in n for k in kws):
            return fk
    soc=str(soc)
    return SOC2_FALLBACK.get(soc[:2]) or SOC1_FALLBACK.get(soc[:1])

def find_col(cols, needle):
    for c in cols:
        if needle.lower() in str(c).lower():
            return c
    return None

def pick_sheet(xl):
    """Choose the 'region x 4-digit SOC' table (Table 3) by its title row."""
    import pandas as pd
    if CONFIG["sheet"]: return CONFIG["sheet"]
    for s in xl.sheet_names:
        try: title=str(xl.parse(s,header=None,nrows=1).iloc[0,0]).lower()
        except Exception: title=""
        if "region" in title and ("4-digit" in title or "4 digit" in title or "four-digit" in title or "soc2020" in title and "region" in title):
            if "region" in title and "4" in title: return s
    # fallback: widest sheet
    return max(xl.sheet_names, key=lambda s: xl.parse(s,header=None,nrows=1).shape[1])

def build(xbytes, ashe=None):
    import pandas as pd
    xl = pd.ExcelFile(io.BytesIO(xbytes))
    sheet = pick_sheet(xl)
    log("Using sheet:", sheet)
    raw = xl.parse(sheet, header=None)

    # The real header row is the one with the most month-parseable cells (titles sit above it).
    best_row, best_count = 0, 0
    for i in range(min(25, len(raw))):
        cnt = sum(1 for x in raw.iloc[i].tolist() if parse_month(x) is not None)
        if cnt > best_count: best_count, best_row = cnt, i
    if best_count < 6:
        raise SystemExit("Couldn't find the month header row. Run --inspect.")
    hdr = raw.iloc[best_row].tolist()
    month_cols = [(j, parse_month(hdr[j])) for j in range(len(hdr)) if parse_month(hdr[j]) is not None]
    first_month_col = min(j for j,_ in month_cols)
    max_idx = max(mi for _,mi in month_cols); N = max_idx+1
    log(f"Header on row {best_row}; {len(month_cols)} month columns; {N} months to index {max_idx}.")

    MONTH_NAMES=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    months=[f"{MONTH_NAMES[i%12]} {2017+i//12}" for i in range(N)]

    data = raw.iloc[best_row+1:].reset_index(drop=True)   # rows below the header
    label_cols = list(range(first_month_col))             # columns left of the months
    def col_vals(j): return [str(v) for v in data.iloc[:,j].tolist()]

    # identify region / SOC-code / occupation-name columns among the label columns
    region_col=soc_col=None
    for j in label_cols:
        vals=col_vals(j); n=max(1,len(vals))
        if region_col is None and sum(1 for v in vals if _region_key_from(v)) > n*0.3: region_col=j; continue
        if soc_col   is None and sum(1 for v in vals if re.fullmatch(r"\s*\d{4}\s*",v)) > n*0.3: soc_col=j; continue
    cand=[j for j in label_cols if j not in (region_col,soc_col)]
    name_col = max(cand, key=lambda j: sum(len(v) for v in col_vals(j))/max(1,len(data))) if cand else None
    if name_col is None: raise SystemExit("Couldn't find the occupation-name column. Run --inspect.")
    log(f"Columns -> region:{region_col}  soc:{soc_col}  name:{name_col}")

    field_series = {fk:[0.0]*N for fk in FIELD_NAME}
    field_region = {fk:{r[0]:0.0 for r in REGIONS if r[0]!="all"} for fk in FIELD_NAME}
    occ_latest = {fk:{} for fk in FIELD_NAME}
    occ_series = {fk:{} for fk in FIELD_NAME}
    occ_soc    = {fk:{} for fk in FIELD_NAME}

    for _, row in data.iterrows():
        nm=str(row.iloc[name_col]).strip()
        low=nm.lower()
        if not nm or low=="nan" or low.startswith("total") or low.startswith("all "): continue
        soc4=""
        if soc_col is not None:
            m4=re.search(r"\d{4}",str(row.iloc[soc_col])); soc4=m4.group(0) if m4 else ""
        fk = assign_field(nm, soc4)
        if fk is None: continue
        rk=None
        if region_col is not None:
            rk=_region_key_from(row.iloc[region_col])
            if rk=="all": continue        # skip UK/aggregate rows so summing regions doesn't double-count
        vals=[0.0]*N
        for j,mi in month_cols:
            v=row.iloc[j]
            try: v=float(v)
            except Exception: v=0.0
            if v!=v: v=0.0
            vals[mi]+=v
        for i in range(N): field_series[fk][i]+=vals[i]
        if rk: field_region[fk][rk]+=vals[max_idx]
        occ_series[fk][nm]=[occ_series[fk][nm][i]+vals[i] for i in range(N)] if nm in occ_series[fk] else vals
        occ_latest[fk][nm]=occ_latest[fk].get(nm,0)+vals[max_idx]
        if soc4 and nm not in occ_soc[fk]: occ_soc[fk][nm]=soc4

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
            dn=clean_name(nm); rd={"name":dn,"salary":nat,"salary_region":reg,"series":[round(x,1) for x in occ_series[fk][nm]]}
            de=desc_for(dn) or FIELD_DESC_FALLBACK.get(fk)
            if de: rd["desc"]=de
            roles.append(rd)
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
    """Prefer the 4-digit SOC, All-employees, Annual pay - Gross workbook (…(4)…7a…), not the CV file."""
    cand=[n for n in names if n.lower().endswith((".xlsx",".xls"))]
    def score(n):
        nl=n.lower(); s=0
        if "(4)" in nl: s+=4                      # 4-digit SOC (matches our roles)
        if "annual" in nl and "gross" in nl: s+=4
        if ".7a" in nl: s+=2
        if ".7b" in nl or " cv" in nl or "cv.xlsx" in nl: s-=8   # CV = coefficient of variation file
        if any(x in nl for x in ["incentive","weekly","hourly","basic","overtime","paid hours"]): s-=4
        if any(x in nl for x in ["male","female","part-time","full-time"]): s-=3
        return s
    cand.sort(key=score,reverse=True)
    return cand[0] if cand else None

def _ashe_region(region_part):
    return ASHE_REGION.get(str(region_part).strip().lower())

def load_ashe_zip(zbytes):
    """Parse ASHE Table 15 zip -> {soc4: {regionkey: median, 'all': median}} from the All sheet.
       Region + occupation are combined in the 'Description' column ('North East, <job>')."""
    import zipfile,io,pandas as pd
    try: z=zipfile.ZipFile(io.BytesIO(zbytes))
    except Exception as e: log("ASHE: not a readable zip:",e); return {}
    wb=_ashe_pick_workbook(z.namelist())
    if not wb: log("ASHE: no workbook found in zip."); return {}
    log("ASHE workbook:",wb.split("/")[-1])
    try: raw=pd.read_excel(io.BytesIO(z.read(wb)), sheet_name=CONFIG.get("ashe_sheet","All"), header=None)
    except Exception as e: log("ASHE: cannot open the All sheet:",e); return {}
    hdr=None
    for i in range(min(14,len(raw))):
        joined=" ".join(str(x).lower() for x in raw.iloc[i].tolist())
        if "description" in joined and "code" in joined and "median" in joined: hdr=i; break
    if hdr is None: log("ASHE: couldn't find the Description/Code/Median header row."); return {}
    cols=[str(x).strip().lower() for x in raw.iloc[hdr].tolist()]
    def colidx(name):
        for j,c in enumerate(cols):
            if c==name: return j
        for j,c in enumerate(cols):
            if name in c: return j
        return None
    dj,cj,mj=colidx("description"),colidx("code"),colidx("median")
    if None in (dj,cj,mj): log("ASHE: missing Description/Code/Median columns."); return {}
    out={}
    for _,row in raw.iloc[hdr+1:].iterrows():
        desc=str(row.iloc[dj]).strip()
        if "," not in desc or desc.lower()=="nan": continue
        region_part,_=desc.split(",",1)
        rk=_ashe_region(region_part)
        if rk is None: continue
        m4=re.search(r"\d{4}",str(row.iloc[cj]))
        if not m4: continue
        val=str(row.iloc[mj]).replace(",","").replace("£","").strip()
        if val.lower() in (":","x","nan","",".."): continue   # not available / suppressed
        try: v=float(val)
        except Exception: continue
        if v!=v or v<=0: continue                              # v!=v catches NaN
        out.setdefault(m4.group(0),{})[rk]=v
    for soc,d in out.items():                 # if no UK row for a SOC, approximate national from regions
        vals=[x for x in d.values() if x==x]
        if "all" not in d and vals: d["all"]=round(sum(vals)/len(vals))
    n_nat=sum(1 for d in out.values() if "all" in d)
    log(f"ASHE: medians for {len(out)} occupations across regions ({n_nat} with a national figure).")
    return out

def inspect_ashe(zbytes):
    import zipfile,io,pandas as pd
    z=zipfile.ZipFile(io.BytesIO(zbytes))
    log("Files in ASHE zip:");  [log("  ",n) for n in z.namelist()]
    wb=_ashe_pick_workbook(z.namelist())
    if wb:
        log("\nPicked workbook:",wb)
        xl=pd.ExcelFile(io.BytesIO(z.read(wb))); log("Sheets:",xl.sheet_names)
        log(f"\nFirst rows of '{CONFIG.get('ashe_sheet','All')}':")
        log(xl.parse(CONFIG.get("ashe_sheet","All"),header=None,nrows=8).iloc[:8,:8].to_string())

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
