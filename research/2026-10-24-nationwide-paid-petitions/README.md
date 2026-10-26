# Nationwide paid petition campaign intelligence (verified Oct 24, 2026)

Research covers all 50 states + DC: who is circulating what, where, for how much, under which company, and until when.

## Files

| File | What it is |
|---|---|
| `Petition_Campaign_Database_2026-10-24.xlsx` | Deliverables 1, 2, 3, 5. Tabs: Read Me, Database (every record, with filters and sources), Active Opportunities (ranked, score is a formula), State Coverage (live counts per state), Firms, Job Postings, Trends & Laws, New Discoveries (monitoring log). |
| `monitoring-playbook.md` | Deliverable 5: the weekly routine for finding new campaigns, rate changes, job ads and deadlines. |
| `data/*.json` | Clean data behind the dashboard (projects, ranked, firms, jobs, trends). |
| `data/raw/*.json` | Raw research output from each regional team, pass 1 and pass 2 (gap-fill). |
| `scripts/build_xlsx.py` | Rebuilds the spreadsheet from `data/raw/`. Run from this folder. |
| `scripts/export.py` | Exports the spreadsheet into `data/*.json` for the dashboard. |

The live dashboard with the map (Deliverable 4) is the claude.ai artifact "Petition Market Intelligence".

## Headline findings

- **Timing:** The Nov 3, 2026 election is days away, so most 2026 drives are finished. Paid demand is moving to the next wave.
- **Hottest window:** Massachusetts 2028. Fourteen AG-certified petitions each need 74,574 signatures, due to local clerks Nov 18, 2026 and to the Secretary Dec 2, 2026. About 10 were filed by Harold Hubschman (SignatureDrive.com) with no paying backer reported as of Sept 22.
- **Strongest proof of paid circulation:** the Maine cannabis rollback (out-of-state paid circulators reported; deadline June 8, 2027).
- **Big open statewide drives:** California's top-two primary repeal (874,641 signatures due Jan 11, 2027), and Nebraska's 15-petition "Make Nebraska Even Better" slate launching after Nov 3 with a new Nebraska-only signature firm.
- **Direct-client leads:** Ohio 2027 drives (Equal Rights, property tax abolition) with no confirmed vendor; the NC Forward Party, whose chair said they will need to "write checks and hire vendors"; the California food sales-tax exemption, whose proponent has no funding; Idaho medical cannabis (two paid vendors and still short for 2026); and the Montgomery County, MD spending cap (missed by 302 signatures, retrying in 2028).
- **Laws:** Florida HB 1205 was upheld (Apr 30, 2026): circulators must be registered Florida residents and U.S. citizens, with fines up to $50K per ineligible circulator. Arkansas banned per-signature pay in 2025.

## Limits

- Nothing was invented. "unknown" means no reliable public source.
- Several official sites blocked automated access (the Florida initiative database, Nevada SOS, Missouri SOS 2028 page, Arizona filings, MA OCPF). Check those by hand first.
- "No verified project" means searched and found nothing, not that none exist.
