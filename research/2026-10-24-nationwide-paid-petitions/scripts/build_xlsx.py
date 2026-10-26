import json, glob, os, re, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import FormulaRule

R = "data/raw"
OUT = "Petition_Campaign_Database_2026-10-24.xlsx"
TODAY = dt.date(2026, 10, 24)

STATES = {"Alabama":"AL","Alaska":"AK","Arizona":"AZ","Arkansas":"AR","California":"CA","Colorado":"CO","Connecticut":"CT","Delaware":"DE","District of Columbia":"DC","Florida":"FL","Georgia":"GA","Hawaii":"HI","Idaho":"ID","Illinois":"IL","Indiana":"IN","Iowa":"IA","Kansas":"KS","Kentucky":"KY","Louisiana":"LA","Maine":"ME","Maryland":"MD","Massachusetts":"MA","Michigan":"MI","Minnesota":"MN","Mississippi":"MS","Missouri":"MO","Montana":"MT","Nebraska":"NE","Nevada":"NV","New Hampshire":"NH","New Jersey":"NJ","New Mexico":"NM","New York":"NY","North Carolina":"NC","North Dakota":"ND","Ohio":"OH","Oklahoma":"OK","Oregon":"OR","Pennsylvania":"PA","Rhode Island":"RI","South Carolina":"SC","South Dakota":"SD","Tennessee":"TN","Texas":"TX","Utah":"UT","Vermont":"VT","Virginia":"VA","Washington":"WA","West Virginia":"WV","Wisconsin":"WI","Wyoming":"WY"}
ABBR = {v: k for k, v in STATES.items()}
ALIASES = {"washington dc": "District of Columbia", "washington, d.c.": "District of Columbia", "dc": "District of Columbia", "d.c.": "District of Columbia", "washington d.c.": "District of Columbia"}
REGION_FILES = {"west":"West","mountain":"Mountain/Southwest","plains":"Plains/Texas","greatlakes":"Great Lakes","northeast":"Northeast","southeast":"Southeast",
                "gap_west":"West","gap_central":"Central","gap_northeast":"Northeast","gap_south":"South"}
STATUS_ORDER = ["Confirmed active paid","Strong evidence active paid","Active – paid unconfirmed","Upcoming","Completed/Expired","No verified project"]

def norm_state(s):
    s = (s or "").strip()
    if s.upper() in ABBR: return ABBR[s.upper()]
    if s.lower() in ALIASES: return ALIASES[s.lower()]
    for k in STATES:
        if s.lower() == k.lower(): return k
    return s

def txt(v):
    if v is None: return "unknown"
    if isinstance(v, list): return "; ".join(str(x) for x in v)
    v = str(v).strip()
    return v if v else "unknown"

def parse_deadline(s):
    s = txt(s)
    m = re.search(r"(20\d\d)-(\d\d)-(\d\d)", s)
    if m:
        try: return dt.date(*map(int, m.groups()))
        except ValueError: return None
    m = re.search(r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+(\d{1,2})(?:\s*[-–]\s*\d{1,2})?,?\s+(20\d\d)", s)
    if m:
        mon = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"].index(m.group(1).lower()) + 1
        try: return dt.date(int(m.group(3)), mon, int(m.group(2)))
        except ValueError: return None
    return None

recs = []
for f in sorted(glob.glob(f"{R}/*.json")):
    key = os.path.basename(f)[:-5]
    if key == "industry": continue
    data = json.load(open(f))
    if isinstance(data, dict): data = data.get("projects", [])
    batch = "Pass 2 (gap-fill)" if key.startswith("gap_") else "Pass 1"
    expanded = []
    for r in data:
        parts = [p.strip() for p in re.split(r"[/,]", str(r.get("state", "")))] if re.search(r"[/,]", str(r.get("state", ""))) else [r.get("state")]
        for p_ in parts:
            rr = dict(r); rr["state"] = p_; expanded.append(rr)
    for r in expanded:
        r = dict(r)
        r["state"] = norm_state(r.get("state"))
        name = txt(r.get("campaign_name"))
        st = txt(r.get("status"))
        if "NO VERIFIED" in name.upper() or st.lower().startswith("no verified"):
            st = "No verified project"
        if st not in STATUS_ORDER:
            st = next((s for s in STATUS_ORDER if s.split()[0].lower() in st.lower()), "Active – paid unconfirmed")
        r["status"] = st
        r["region"] = REGION_FILES.get(key, key)
        r["batch"] = batch
        recs.append(r)

# de-duplicate: same state + same campaign name -> keep the later pass (more facts), merge sources
STOP = set("the a an of and to for in on 2026 2027 2028 initiative petition petitions measure act cycle state statewide amendment constitutional proposed referendum campaign drive new".split())
def ids(r):
    n = txt(r.get("campaign_name"))
    return set(re.findall(r"\b2[56]-\d\d\b|#\s?\d{2,4}\b|\bIP\s?#?\d+\b|\bI-\d+\b|\bIL2\d-\d+\b|Initiative \d{2,3}(?![-\d])", n))
def words(r):
    return {w for w in re.findall(r"[a-z]{3,}", txt(r.get("campaign_name")).lower()) if w not in STOP}
def same(a, b):
    if a["state"] != b["state"]: return False
    if "No verified project" in (a["status"], b["status"]): return a["status"] == b["status"]
    ia, ib = ids(a), ids(b)
    if ia and ib and len(ia) <= 2 and len(ib) <= 2 and ia & ib: return True
    if len(ia) > 2 and len(ib) > 2 and len(ia & ib) >= 3: return True
    na, nb = txt(a.get("campaign_name")).lower(), txt(b.get("campaign_name")).lower()
    for kw in ("hubschman", "signaturedrive", "top two", "top-two", "compassionate care", "landry"):
        if kw in na and kw in nb: return True
    wa, wb = words(a), words(b)
    if not wa or not wb: return False
    return len(wa & wb) / len(wa | wb) >= 0.4
merged = []
for r in recs:
    hit = next((m for m in merged if same(m, r)), None)
    if hit:
        srcs = list(dict.fromkeys((hit.get("sources") or []) + (r.get("sources") or [])))
        for fld, v in r.items():
            if txt(v) != "unknown" and fld not in ("batch", "campaign_name"): hit[fld] = v
        hit["sources"] = srcs
        hit["batch"] = "Pass 1 + Pass 2"
    else:
        merged.append(r)
recs = merged

# A state with a real project shouldn't also show a "No verified project" placeholder
real_states = {r["state"] for r in recs if r["status"] != "No verified project"}
recs = [r for r in recs if not (r["status"] == "No verified project" and r["state"] in real_states)]
recs.sort(key=lambda r: (r["state"], STATUS_ORDER.index(r["status"]), txt(r.get("deadline"))))

# ---------- styles ----------
F = "Arial"
H = Font(name=F, bold=True, color="FFFFFF", size=10)
HF = PatternFill("solid", fgColor="1F3A5F")
B = Font(name=F, size=10)
BB = Font(name=F, size=10, bold=True)
LINK = Font(name=F, size=10, color="0563C1", underline="single")
TITLE = Font(name=F, size=16, bold=True, color="1F3A5F")
WRAP = Alignment(wrap_text=True, vertical="top")
thin = Side(style="thin", color="D0D7DE")
STATUS_FILL = {"Confirmed active paid":"C6EFCE","Strong evidence active paid":"E2F0D9","Active – paid unconfirmed":"FFF2CC","Upcoming":"DDEBF7","Completed/Expired":"EDEDED","No verified project":"F8CBAD"}

def header(ws, row, cols, widths):
    for i, (c, w) in enumerate(zip(cols, widths), 1):
        cell = ws.cell(row=row, column=i, value=c)
        cell.font, cell.fill = H, HF
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[row].height = 30

def add_table(ws, ref, name):
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    ws.add_table(t)

wb = Workbook()

# ---------- README ----------
ws = wb.active; ws.title = "Read Me"
ws["A1"] = "Nationwide Paid Petition Campaign Database"; ws["A1"].font = TITLE
lines = [
 ("Last verified", "October 24, 2026 (10 days before the Nov 3, 2026 general election)"),
 ("What this is", "Every petition project our research could find in all 50 states + DC, with status, deadlines, pay, firms and sources."),
 ("Big picture", "Most 2026 ballot drives are finished. Live money is in the next wave: Massachusetts 2028 (first round due Nov 18 / Dec 2, 2026), Maine and Washington petitions to their legislatures, California's Jan 11, 2027 deadline, and the 2027 local / 2028 statewide cycle."),
 ("", ""),
 ("STATUS KEY", ""),
 ("Confirmed active paid", "A current source (job ad naming the petition, finance filing, or official) shows paid circulators working it now."),
 ("Strong evidence active paid", "Recent, credible reports of paid circulation, but not a current posting or filing."),
 ("Active – paid unconfirmed", "The petition is collecting now, but we could not prove circulators are paid."),
 ("Upcoming", "Announced or filed; collection hasn't started, or starts soon."),
 ("Completed/Expired", "Drive is over. Kept because it shows who spends money and which firms they hire."),
 ("No verified project", "We searched and found nothing verifiable. This means 'not found yet', NOT 'none exist'."),
 ("", ""),
 ("TABS", ""),
 ("Database", "Every record. Use the filter arrows in the header row to sort by state, status, deadline, firm."),
 ("Active Opportunities", "Live or near-live projects ranked by 5 scores (hiring, pay, deadline urgency, scale, evidence). Total = formula."),
 ("State Coverage", "One row per state with live counts (formulas) so you can see where activity is and where gaps remain."),
 ("Firms", "Circulation companies found, their states, contacts and pay."),
 ("Job Postings", "Recent circulator job ads with stated pay."),
 ("Trends & Laws", "Industry notes: pay rates, new laws, where the market is moving."),
 ("New Discoveries", "Monitoring log. Add each newly found project here first, then move it to Database once verified."),
 ("", ""),
 ("RULES WE FOLLOWED", "Nothing invented. 'unknown' means no reliable public source. A registered petition is NOT counted as paid. Contacts are official public business contacts only."),
 ("Known gaps", "Several state election sites and campaign-finance systems blocked automated access (e.g. Florida initiative database, Nevada SOS, Missouri SOS 2028 page, Arizona filings). Those are the first places to check by hand."),
]
for i, (a, b) in enumerate(lines, 3):
    ws.cell(row=i, column=1, value=a).font = BB
    c = ws.cell(row=i, column=2, value=b); c.font = B; c.alignment = WRAP
    if a in STATUS_FILL:
        ws.cell(row=i, column=1).fill = PatternFill("solid", fgColor=STATUS_FILL[a])
ws.column_dimensions["A"].width = 28; ws.column_dimensions["B"].width = 110

# ---------- Database ----------
db = wb.create_sheet("Database")
cols = ["ID","State","Region","Area","Campaign / Petition","Purpose","Type","Sponsor / Committee","Circulation Firm","Status",
        "Signatures Required","Deadline","Deadline (date)","Days Left","Signatures Collected","Circulator Pay","Hiring Link","Budget / Spend",
        "Decision Maker","Public Contact","Website","Evidence Note","Source 1","Source 2","Source 3","More Sources","Discovery Batch","Last Verified"]
widths = [7,14,14,18,34,40,14,28,24,22,14,20,12,9,14,22,30,24,22,26,28,50,34,34,34,30,14,12]
header(db, 1, cols, widths)
for i, r in enumerate(recs, 2):
    srcs = r.get("sources") or []
    if isinstance(srcs, str): srcs = [srcs]
    d = parse_deadline(r.get("deadline"))
    row = [f"P{i-1:03d}", r["state"], r["region"], txt(r.get("area")), txt(r.get("campaign_name")), txt(r.get("purpose")), txt(r.get("type")),
           txt(r.get("sponsor")), txt(r.get("circulation_firm")), r["status"], txt(r.get("signatures_required")), txt(r.get("deadline")),
           d, None, txt(r.get("signatures_collected")), txt(r.get("pay_rate")), txt(r.get("hiring_link")), txt(r.get("budget_or_spend")),
           txt(r.get("decision_maker")), txt(r.get("contact")), txt(r.get("website")), txt(r.get("evidence_note")),
           srcs[0] if len(srcs) > 0 else "", srcs[1] if len(srcs) > 1 else "", srcs[2] if len(srcs) > 2 else "",
           "; ".join(srcs[3:]), r["batch"], dt.date(2026, 10, 24)]
    for j, v in enumerate(row, 1):
        c = db.cell(row=i, column=j, value=v); c.font = B; c.alignment = WRAP
    db.cell(row=i, column=13).number_format = "yyyy-mm-dd"
    db.cell(row=i, column=28).number_format = "yyyy-mm-dd"
    db.cell(row=i, column=14, value=f'=IF(M{i}="","",M{i}-\'Read Me\'!$B$30)').number_format = "0"
    for j in (17, 21, 23, 24, 25):
        v = db.cell(row=i, column=j).value
        if isinstance(v, str) and v.startswith("http"):
            db.cell(row=i, column=j).hyperlink = v; db.cell(row=i, column=j).font = LINK
    db.cell(row=i, column=10).fill = PatternFill("solid", fgColor=STATUS_FILL[r["status"]])
last = len(recs) + 1
add_table(db, f"A1:{get_column_letter(len(cols))}{last}", "Projects")
db.freeze_panes = "F2"
# "today" cell used by Days Left
wb["Read Me"]["A30"] = "Today (edit to refresh 'Days Left')"; wb["Read Me"]["A30"].font = BB
wb["Read Me"]["B30"] = TODAY; wb["Read Me"]["B30"].number_format = "yyyy-mm-dd"
wb["Read Me"]["B30"].fill = PatternFill("solid", fgColor="FFFF00"); wb["Read Me"]["B30"].font = Font(name=F, color="0000FF")

# ---------- Active Opportunities ----------
def scores(r):
    st = r["status"]
    ev = {"Confirmed active paid":5,"Strong evidence active paid":4,"Active – paid unconfirmed":2,"Upcoming":1}.get(st, 0)
    hire = 5 if str(r.get("hiring_link","")).startswith("http") else (2 if txt(r.get("circulation_firm")) != "unknown" else 0)
    pay = 5 if re.search(r"\$\s?\d", txt(r.get("pay_rate"))) else 0
    d = parse_deadline(r.get("deadline"))
    if d:
        days = (d - TODAY).days
        urg = 5 if 0 <= days <= 45 else 4 if days <= 90 else 3 if days <= 180 else 2 if days <= 400 else 1
        if days < 0: urg = 0
    else: urg = 1
    area = txt(r.get("area")).lower()
    sig = re.sub(r"[^\d]", "", txt(r.get("signatures_required")))
    n = int(sig) if sig and len(sig) < 9 else 0
    scale = 5 if n >= 300000 else 4 if n >= 70000 else 3 if ("statewide" in area or n >= 20000) else 2 if n else 1
    return hire, pay, urg, scale, ev

act = [r for r in recs if r["status"] in STATUS_ORDER[:4]]
scored = sorted(((scores(r), r) for r in act), key=lambda x: (-sum(x[0]), STATUS_ORDER.index(x[1]["status"])))
ao = wb.create_sheet("Active Opportunities")
ao["A1"] = "Ranked Opportunities — live and near-live projects"; ao["A1"].font = TITLE
ao["A2"] = ("Each project gets 0–5 points on the five ranking rules. Total is a formula. Scoring: Hiring = 5 with a live job link, 2 if a firm is named. "
            "Pay = 5 if a dollar rate is published. Urgency = 5 if deadline ≤45 days, 4 ≤90, 3 ≤180, 2 ≤400, 1 unknown. Scale = signatures required "
            "(5 ≥300k, 4 ≥70k, 3 statewide/≥20k, 2 local, 1 unknown). Evidence = 5 confirmed, 4 strong, 2 unconfirmed, 1 upcoming.")
ao["A2"].font = Font(name=F, size=9, italic=True); ao["A2"].alignment = WRAP; ao.merge_cells("A2:P2"); ao.row_dimensions[2].height = 42
c2 = ["Rank","State","Campaign / Petition","Status","Circulation Firm","Circulator Pay","Deadline","Signatures Required","Hiring Link",
      "Hiring","Pay","Urgency","Scale","Evidence","TOTAL (0–25)","Why it matters"]
header(ao, 4, c2, [6,14,36,22,24,22,16,14,30,8,8,8,8,9,11,60])
for i, (sc, r) in enumerate(scored, 5):
    vals = [i-4, r["state"], txt(r.get("campaign_name")), r["status"], txt(r.get("circulation_firm")), txt(r.get("pay_rate")),
            txt(r.get("deadline")), txt(r.get("signatures_required")), txt(r.get("hiring_link")), *sc, f"=SUM(J{i}:N{i})", txt(r.get("evidence_note"))]
    for j, v in enumerate(vals, 1):
        c = ao.cell(row=i, column=j, value=v); c.font = B; c.alignment = WRAP
    ao.cell(row=i, column=15).font = BB
    ao.cell(row=i, column=4).fill = PatternFill("solid", fgColor=STATUS_FILL[r["status"]])
    for j in range(10, 15): ao.cell(row=i, column=j).font = Font(name=F, size=10, color="0000FF")
    hl = ao.cell(row=i, column=9)
    if str(hl.value).startswith("http"): hl.hyperlink = hl.value; hl.font = LINK
alast = 4 + len(scored)
if scored: add_table(ao, f"A4:P{alast}", "Opportunities")
ao.freeze_panes = "D5"

# ---------- State Coverage ----------
sc_ = wb.create_sheet("State Coverage")
c3 = ["State","Abbr"] + STATUS_ORDER + ["Total Records","Live (first 4)","Coverage Note"]
header(sc_, 1, c3, [22,6,12,12,12,10,12,12,10,10,50])
allst = sorted(STATES)
for i, s in enumerate(allst, 2):
    sc_.cell(row=i, column=1, value=s).font = B
    sc_.cell(row=i, column=2, value=STATES[s]).font = B
    for j, st in enumerate(STATUS_ORDER, 3):
        sc_.cell(row=i, column=j, value=f'=COUNTIFS(Database!$B:$B,$A{i},Database!$J:$J,"{st}")').font = B
    sc_.cell(row=i, column=9, value=f"=SUM(C{i}:H{i})").font = BB
    sc_.cell(row=i, column=10, value=f"=SUM(C{i}:F{i})").font = BB
    sc_.cell(row=i, column=11, value=f'=IF(I{i}=0,"Not yet researched",IF(H{i}>0,"Searched — nothing verified (gap)",""))').font = B
n = len(allst) + 1
sc_.cell(row=n+1, column=1, value="TOTAL").font = BB
for j in range(3, 11):
    col = get_column_letter(j)
    sc_.cell(row=n+1, column=j, value=f"=SUM({col}2:{col}{n})").font = BB
add_table(sc_, f"A1:K{n}", "Coverage")
sc_.conditional_formatting.add(f"J2:J{n}", FormulaRule(formula=["J2>0"], fill=PatternFill("solid", fgColor="C6EFCE")))
sc_.freeze_panes = "C2"

# ---------- Firms ----------
ind = json.load(open(f"{R}/industry.json"))
def fkey(n): return re.sub(r"[^a-z]", "", re.split(r" LLC| Inc|,| [(]|/", n)[0].lower())[:14]
firms = {fkey(f["name"]): dict(f) for f in ind.get("firms", []) if f.get("name")}
gs = f"{R}/gap_south.json"
if os.path.exists(gs):
    g = json.load(open(gs))
    if isinstance(g, dict):
        for f in g.get("firms", []):
            if not f.get("name"): continue
            k = fkey(f["name"])
            base = firms.setdefault(k, {"name": f["name"]})
            for fld, v in f.items():
                if fld == "sources": base["sources"] = list(dict.fromkeys((base.get("sources") or []) + (v or [])))
                elif txt(v) != "unknown": base[fld] = v
# count database projects per firm (text match)
fm = wb.create_sheet("Firms")
c4 = ["Firm","HQ","Website","Public Contact","States Seen","Known Campaigns (2026)","Projects in Database","Recruiting Link","Stated Pay","Notes","Sources"]
header(fm, 1, c4, [28,20,28,30,24,40,12,30,22,50,50])
for i, f in enumerate(sorted(firms.values(), key=lambda x: x["name"].lower()), 2):
    short = re.split(r"[\s(]", f["name"].strip())[0]
    vals = [f["name"], txt(f.get("hq")), txt(f.get("website")), txt(f.get("contact")), txt(f.get("active_states")),
            txt(f.get("known_campaigns_2026")), f'=COUNTIF(Database!$I:$I,"*"&"{re.split(r" LLC| Inc|,| [(]|/", f["name"])[0].strip().replace(chr(34),"")}"&"*")',
            txt(f.get("recruiting_link")), txt(f.get("stated_pay")), txt(f.get("notes")), txt(f.get("sources"))]
    for j, v in enumerate(vals, 1):
        c = fm.cell(row=i, column=j, value=v); c.font = B; c.alignment = WRAP
    for j in (3, 8):
        v = fm.cell(row=i, column=j).value
        if isinstance(v, str) and v.startswith("http"): fm.cell(row=i, column=j).hyperlink = v; fm.cell(row=i, column=j).font = LINK
add_table(fm, f"A1:K{len(firms)+1}", "FirmsTbl")
fm.freeze_panes = "B2"

# ---------- Job Postings ----------
jp = wb.create_sheet("Job Postings")
c5 = ["State","City","Campaign","Company","Stated Pay","Date Posted","Platform","Link"]
header(jp, 1, c5, [14,16,30,24,26,14,14,50])
jobs = ind.get("job_postings", [])
for i, j_ in enumerate(jobs, 2):
    vals = [norm_state(j_.get("state")), txt(j_.get("city")), txt(j_.get("campaign")), txt(j_.get("company")), txt(j_.get("pay")),
            txt(j_.get("date_posted")), txt(j_.get("source_platform")), txt(j_.get("link"))]
    for k, v in enumerate(vals, 1):
        c = jp.cell(row=i, column=k, value=v); c.font = B; c.alignment = WRAP
    l = jp.cell(row=i, column=8)
    if str(l.value).startswith("http"): l.hyperlink = l.value; l.font = LINK
if jobs: add_table(jp, f"A1:H{len(jobs)+1}", "Jobs")
jp.freeze_panes = "A2"

# ---------- Trends ----------
tr = wb.create_sheet("Trends & Laws")
header(tr, 1, ["Topic","Finding","Sources"], [28,90,60])
for i, t in enumerate(ind.get("trends", []), 2):
    for k, v in enumerate([txt(t.get("topic")), txt(t.get("finding")), txt(t.get("sources"))], 1):
        c = tr.cell(row=i, column=k, value=v); c.font = B; c.alignment = WRAP

# ---------- New Discoveries (monitoring log) ----------
nd = wb.create_sheet("New Discoveries")
nd["A1"] = "Monitoring log — add newly found projects here first"; nd["A1"].font = TITLE
nd["A2"] = "Yellow row = example of the expected format; overwrite it. Once a project is verified, copy it to Database and mark 'Moved to Database' = Yes."
nd["A2"].font = Font(name=F, size=9, italic=True)
c6 = ["Date Found","State","Campaign / Petition","How Found (source type)","Status","Firm","Pay","Deadline","Source Link","Moved to Database?"]
header(nd, 4, c6, [12,14,34,24,22,22,18,14,40,14])
ex = ["2026-11-02","Massachusetts","(example) Gas tax repeal 2028","OCPF expenditure report","Strong evidence active paid","unknown","unknown","2026-11-18","https://www.ocpf.us","No"]
for k, v in enumerate(ex, 1):
    c = nd.cell(row=5, column=k, value=v); c.font = Font(name=F, size=10, color="0000FF"); c.fill = PatternFill("solid", fgColor="FFFF00")

wb.save(OUT)
print("records", len(recs), "active", len(act), "firms", len(firms), "jobs", len(jobs))
