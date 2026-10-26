# Monitoring playbook: keeping the petition database fresh

**Goal:** find new paid petition campaigns, rate changes, job ads and deadlines before competitors do.

**Where finds go:** log every new find in the spreadsheet's **New Discoveries** tab first. Move it to **Database** only after it is verified, and set "Moved to Database?" to Yes. That keeps new finds separate from the existing records.

## Weekly rhythm (about 3 hours a week)

### Monday: official filings ("who filed what")
1. Check the Secretary of State or elections pages in the active states: MA, ME, WA, CA, OR, OH, MI, MO, NE, FL, AZ, CO, and the DC Board of Elections.
2. Look for three changes: newly approved petitions, new signature counts, new deadlines.
3. Check the paid-circulator registries: Colorado licensed petition firms, Arizona and Oregon paid-circulator registration, and Florida circulator registration. **A new registration is the earliest strong sign a drive is going paid.**

### Wednesday: follow the money
1. Search campaign finance for payments to petition firms: Cal-Access (CA), OCPF (MA), Maine Ethics Commission, WA PDC, ORESTAR (OR), Missouri MEC, Nebraska NADC, Colorado TRACER, Arizona SeeTheMoney, Florida DOE.
2. Search these payee names: FieldWorks, Advanced Micro Targeting, Petition Partners, PCI Consultants, Landslide Political, Blitz Canvassing, Let the Voters Decide, The Outreach Team, SignatureDrive.com, Taylor Petition Management, Campaign & Petition Management, Arno Petition Consultants, Allstate Petition Management.
3. A payment to a firm upgrades a record to **Confirmed active paid**.

### Friday: job boards (where pay rates show up first)
1. Search Indeed, ZipRecruiter, Craigslist (jobs and gigs in every large metro), petitioner Facebook groups and LinkedIn.
2. Use these search terms: "petition circulator", "signature gatherer", "petitioner", "petition field manager", "signature collector".
3. Log every posting with its pay, its date and the petition it names. Put rate changes in the **Trends & Laws** tab.

### Daily: news alerts
- Set up Google Alerts for: "signature gathering", "petition drive", "ballot initiative" signatures, "recall petition", "paid circulators", "petition circulators".
- Skim Ballot Access News and Ballotpedia News for national roundups.

## Status upgrade rules
| Status | What it takes |
|---|---|
| Confirmed active paid | A current job ad naming the petition, a finance filing paying a firm, or an official statement |
| Strong evidence active paid | Credible recent reporting of paid crews, but no filing or posting yet |
| Active – paid unconfirmed | Collecting, with no proof of pay. **These are the best pitch targets for a full-service campaign.** |
| Upcoming | Filed or announced, not yet collecting |
| Completed/Expired | Over. Keep the record, because it shows who spends and which firms they hire |
| No verified project | Searched and found nothing. Re-check monthly |

Never upgrade a record on a registration alone.

## Calendar to watch (as of Oct 24, 2026)
- **Nov 18 / Dec 2, 2026:** Massachusetts 2028 first-round signatures (local clerks, then the Secretary).
- **Dec 31, 2026:** Washington initiatives to the Legislature.
- **Jan 11, 2027:** California top-two primary repeal.
- **Mar 24, 2027:** California food sales-tax exemption.
- **Apr 24 / Jun 8, 2027:** Maine health care and cannabis initiatives.
- **Jul 1, 2027 (likely):** Ohio November 2027 amendments.
- **After Nov 3, 2026:** watch for 2028 launches in NE, MO, MI, FL, OH, ID and OR.
- **Feb 1, 2028:** Florida 2028 initiative deadline.

## Refresh the files
1. Add new research to `data/raw/` as JSON, in the same fields as the existing files.
2. Run `python3 scripts/build_xlsx.py`, then recalculate the formulas (open the file in Excel or LibreOffice).
3. Run `python3 scripts/export.py` to regenerate `data/*.json` for the dashboard.
