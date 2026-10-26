import pandas as pd, json
X='Petition_Campaign_Database_2026-10-24.xlsx'
db=pd.read_excel(X,sheet_name='Database')
ao=pd.read_excel(X,sheet_name='Active Opportunities',header=3)
fm=pd.read_excel(X,sheet_name='Firms')
jp=pd.read_excel(X,sheet_name='Job Postings')
tr=pd.read_excel(X,sheet_name='Trends & Laws')
ab={'Alabama':'AL','Alaska':'AK','Arizona':'AZ','Arkansas':'AR','California':'CA','Colorado':'CO','Connecticut':'CT','Delaware':'DE','District of Columbia':'DC','Florida':'FL','Georgia':'GA','Hawaii':'HI','Idaho':'ID','Illinois':'IL','Indiana':'IN','Iowa':'IA','Kansas':'KS','Kentucky':'KY','Louisiana':'LA','Maine':'ME','Maryland':'MD','Massachusetts':'MA','Michigan':'MI','Minnesota':'MN','Mississippi':'MS','Missouri':'MO','Montana':'MT','Nebraska':'NE','Nevada':'NV','New Hampshire':'NH','New Jersey':'NJ','New Mexico':'NM','New York':'NY','North Carolina':'NC','North Dakota':'ND','Ohio':'OH','Oklahoma':'OK','Oregon':'OR','Pennsylvania':'PA','Rhode Island':'RI','South Carolina':'SC','South Dakota':'SD','Tennessee':'TN','Texas':'TX','Utah':'UT','Vermont':'VT','Virginia':'VA','Washington':'WA','West Virginia':'WV','Wisconsin':'WI','Wyoming':'WY'}
def s(v):
    if v is None or (isinstance(v,float) and pd.isna(v)) or v is pd.NaT: return ''
    return str(v)
proj=[]
for _,r in db.iterrows():
    d=r['Deadline (date)']
    dd = d.strftime('%Y-%m-%d') if (hasattr(d,'strftime') and not pd.isna(d)) else ''
    proj.append({'id':r['ID'],'state':r['State'],'abbr':ab.get(r['State'],''),'campaign':s(r['Campaign / Petition']),'type':s(r['Type']),
      'sponsor':s(r['Sponsor / Committee']),'firm':s(r['Circulation Firm']),'status':r['Status'],'signatures_required':s(r['Signatures Required']),
      'deadline':s(r['Deadline']),'deadline_date':dd,'pay':s(r['Circulator Pay']),'contact':s(r['Public Contact']),'source':s(r['Source 1']),
      'evidence':s(r['Evidence Note'])})
rank=[]
for _,r in ao.iterrows():
    if not s(r['State']): continue
    sc=[int(r[c]) for c in ['Hiring','Pay','Urgency','Scale','Evidence']]
    rank.append({'rank':int(r['Rank']),'state':r['State'],'campaign':s(r['Campaign / Petition']),'status':r['Status'],'firm':s(r['Circulation Firm']),
      'pay':s(r['Circulator Pay']),'deadline':s(r['Deadline']),'hiring':sc[0],'pay_score':sc[1],'urgency':sc[2],'scale':sc[3],'evidence':sc[4],'total':sum(sc),'why':s(r['Why it matters'])})
firms=[{'firm':s(r['Firm']),'hq':s(r['HQ']),'contact':s(r['Public Contact']),'states':s(r['States Seen']),'projects':int(r['Projects in Database'] if not pd.isna(r['Projects in Database']) else 0),'pay':s(r['Stated Pay']),'website':s(r['Website'])} for _,r in fm.iterrows() if not s(r['Firm']).startswith('Other firms')]
jobs=[{'job':f'J{i+1:02d}','state':s(r['State']),'city':s(r['City']),'campaign':s(r['Campaign']),'company':s(r['Company']),'pay':s(r['Stated Pay']),'posted':s(r['Date Posted']),'link':s(r['Link'])} for i,(_,r) in enumerate(jp.iterrows())]
trends=[{'topic':s(r['Topic']),'finding':s(r['Finding'])} for _,r in tr.iterrows()]
for n,v in [('projects',proj),('ranked',rank),('firms',firms),('jobs',jobs),('trends',trends)]:
    json.dump(v,open(f'data/{n}.json','w'),indent=0); print(n,len(v))
print(len(set(p['id'] for p in proj)), len(set(j['job'] for j in jobs)), len(set(f['firm'] for f in firms)), len(set(r['rank'] for r in rank)))
print([(r['rank'],r['state'],r['total']) for r in rank[:6]])
print(sorted(p['deadline_date'] for p in proj if p['status'] not in ('Completed/Expired','No verified project') and p['deadline_date']>='2026-10-24'))
