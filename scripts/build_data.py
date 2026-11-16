import openpyxl,re,json,sys
wb=openpyxl.load_workbook(sys.argv[1] if len(sys.argv)>1 else 'New_PDP_2H2026.xlsx',data_only=True)
S=lambda v:('' if v is None else str(v).replace('\n',' ').strip())
def num(v):
    try: return float(str(v).replace(',','').split()[0])
    except: return None
nodes=[];pics={};
def pic(name,dept):
    k=name.strip().lower()
    if k not in pics: pics[k]={'id':'p%d'%len(pics),'n':name.strip(),'d':[]}
    if dept and dept not in pics[k]['d']: pics[k]['d'].append(dept)
    return pics[k]['id']
def resources(ws,r,c0,c1,hr,dr):
    P,Sx=[],[]
    for c in range(c0,c1+1):
        v=ws.cell(r,c+1).value
        if v in('l','m'):
            nm=S(ws.cell(hr,c+1).value)
            if nm: (P if v=='l' else Sx).append(pic(nm,S(ws.cell(dr,c+1).value)))
    return P,Sx
s1=wb['Sheet1']
def monthly(r):
    m=[num(s1.cell(r,8+i+1).value) for i in range(12)]
    return m if any(x is not None for x in m) else None
OBJ=[]  # objectives
# --- L1 (matrix gives codes/pics, Sheet1 gives target+monthly)
mx=wb['Top Level Matrix']
s1rows={}
obj=None
for r in range(4,26):
    if s1.cell(r,2).value: OBJ.append({'c':S(s1.cell(r,2).value),'n':S(s1.cell(r,3).value)})
    s1rows[S(s1.cell(r,5).value).lower()]=(r,len(OBJ)-1)
alias={'mini line':'mini line & s/f collocation'}
l1={}
for r in range(4,26):
    n=S(mx.cell(r,9).value);k=alias.get(n.lower(),n.lower())
    sr,oi=s1rows[k]
    P,Sx=resources(mx,r,34,46,26,27)
    nd={'id':'n%d'%len(nodes),'l':1,'c':S(mx.cell(r,8).value),'n':n,'o':oi,'p':None,'t':S(s1.cell(sr,8).value)+' poin','tn':num(s1.cell(sr,8).value),'d':'ge','m':monthly(sr),'x':S(s1.cell(sr,7).value),'pp':P,'sp':Sx}
    nodes.append(nd);l1[n.lower()]=nd
# --- L2
m2=wb['2nd Level Matrix']
colname={c:S(m2.cell(54,c+1).value).lower() for c in range(2,23)}
amap={'mini line & s/f collocation':'mini line'}
l2=[]
for i,r in enumerate(range(4,54)):
    code=S(m2.cell(r,24).value) and S(s1.cell(31+i,4).value)
    sr=31+i
    left=[c for c in range(2,23) if m2.cell(r,c+1).value=='l'][0]
    pn=colname[left]
    cand=[v for k,v in l1.items() if k==pn or k==amap.get(pn,pn)]
    if not cand: cand=[v for k,v in l1.items() if k.startswith(pn[:12])]
    if pn.startswith('development quality'):
        cand=[l1['development quality team survey' if 'survey' in S(s1.cell(sr,5).value).lower()+S(s1.cell(sr,7).value).lower() else 'development quality assessment']]
    par=cand[0]
    t=S(s1.cell(sr,8).value);tn=num(re.sub(r'[<≤≥>]','',t)) ;d='le' if (t[:1] in '<≤' or tn==0) else 'ge'
    m=monthly(sr)
    if m and d=='le' and tn and tn>1000: m=[(x*1000 if x is not None and x<100 else x) for x in m]
    P,Sx=resources(m2,r,78,92,54,55)
    nd={'id':'n%d'%len(nodes),'l':2,'c':S(s1.cell(sr,4).value),'n':S(s1.cell(sr,5).value),'o':par['o'],'p':par['id'],'t':t,'tn':tn,'d':d,'m':m,'x':S(s1.cell(sr,7).value),'pp':P,'sp':Sx}
    nodes.append(nd);l2.append(nd)
# --- L3
m3=wb['3rd Level Matrix']
for i,r in enumerate(range(4,60)):
    left=[c for c in range(2,52) if m3.cell(r,c+1).value=='l'][0]
    bn=S(m3.cell(60,left+1).value).lower()
    sm=[x for x in l2 if S(m2.cell(4+l2.index(x),25).value).lower()==bn or x['n'].lower()==bn]
    par=min(sm,key=lambda x:abs(l2.index(x)-(left-2))) if sm else l2[left-2]
    if not sm: print('fallback',r,bn)
    tg=[[S(m3.cell(61,c+1).value),S(m3.cell(60,c+1).value)] for c in range(53,113) if m3.cell(r,c+1).value=='l']
    P,Sx=resources(m3,r,116,130,60,61)
    nodes.append({'id':'n%d'%len(nodes),'l':3,'c':S(m3.cell(r,53).value),'n':S(m3.cell(r,54).value),'o':par['o'],'p':par['id'],'t':' · '.join(q[1] for q in tg),'tn':None,'d':'ge','m':None,'x':', '.join(q[0] for q in tg),'pp':P,'sp':Sx,'tg':tg})
def hdr(ws,c0,c1,hr,dr): return [[pic(S(ws.cell(hr,c+1).value),S(ws.cell(dr,c+1).value)),S(ws.cell(dr,c+1).value)] for c in range(c0,c1+1) if S(ws.cell(hr,c+1).value)]
bw=wb['Bowler Chart']
for r in range(4,81):
    it=S(bw.cell(r,3).value).lower()
    if not it or num(bw.cell(r,4).value) is None: continue
    L=[x for x in nodes if x['l']==1 and 'h' not in x]
    nd=next((x for x in L if x['n'].lower()==it),None) or next((x for x in L if x['n'].lower().startswith(it[:12]) or it.startswith(x['n'].lower()[:12])),None)
    if nd:
        a=[num(bw.cell(r,5+i).value) for i in range(12)]
        if any(v is not None for v in a): nd['h']=a
cols={'1':hdr(mx,34,46,26,27),'2':hdr(m2,78,92,54,55),'3':hdr(m3,116,130,60,61)}
json.dump({'obj':OBJ,'nodes':nodes,'pics':list(pics.values()),'cols':cols},open('scripts/data.json','w'),ensure_ascii=False,separators=(',',':'))
print(len(nodes),len(pics),OBJ,[ (n['c'],n['n'][:20],n['p']) for n in nodes if n['l']==2][:4])
print('scripts/data.json ditulis')
