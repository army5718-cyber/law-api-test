"""VBA(Module1/2/4) 핵심 로직을 파이썬으로 옮겨 실제 데이터로 돌려 보는 테스트용 포트.
 엑셀 객체 모델은 최소한의 인메모리 모델로 대체한다 (시트 셀 사전, 병합, 원데이터 표)."""
import datetime, pickle, re, collections, openpyxl

KIND_REG,KIND_PRE,KIND_MEMO='정기검사','사용전검사','메모'
ZONE_A,ZONE_B='부천지역','김포지역'
RW=dict(DATE=0,KIND=1,ZONE=2,NAME=3,CAP=4,REG=5,TIME=6,FEE=7,INSP=8,NOTE=9,ID=10)  # 0-based
LEDGER_MIN_ROWS=24; SCR_FIRST_ROW=7; NOTE_L=13; NOTE_R=14

def vstr(x):
    if x is None: return ''
    return str(x).strip()
def is_task_kind(k): return k[:2] in ('정기','사용')
def is_regular(k): return k[:2]=='정기'
def day_key(v):
    if isinstance(v,datetime.datetime): return v.toordinal()
    if isinstance(v,datetime.date): return v.toordinal()
    return -1
def cap_num(x):
    if x is None: return 0.0
    if isinstance(x,(int,float)): return float(x)
    s=str(x).strip(); t=''
    for ch in s:
        if ch in '0123456789.': t+=ch
        else: break
    try: return float(t) if t else 0.0
    except: return 0.0
def cell_time_str(v):
    if isinstance(v,datetime.datetime): return v.strftime('%H:%M')
    if isinstance(v,datetime.time): return v.strftime('%H:%M')
    if isinstance(v,float) and 0<v<1:
        m=round(v*1440); return f'{m//60:02d}:{m%60:02d}'
    return vstr(v)
def promise_time(v):
    if v is None: return ''
    s=''
    if isinstance(v,(datetime.datetime,datetime.time)) or (isinstance(v,float) and 0<v<1): s=cell_time_str(v)
    elif isinstance(v,str):
        try: s=datetime.datetime.strptime(v.strip(),'%H:%M').strftime('%H:%M')
        except: s=''
    return s if re.fullmatch(r'\d\d:\d\d',s) else ''
FULL=('휴가','병가','교육','건강','체련','연가'); PART=('반차','오후','퇴근','조퇴','(')
def has_any(t,k): return any(x in t for x in k)
def is_full_leave(t): return has_any(t,FULL) and not has_any(t,PART)

class Env:
    def __init__(self,raw,erp='stage1.xlsm'):
        wb=openpyxl.load_workbook(erp,keep_vba=False)
        self.roster=[(r[0],str(r[2]).upper()=='TRUE') for r in wb['01_검사자명단'].iter_rows(min_row=2,max_row=30,max_col=3,values_only=True) if r[0] and not str(r[0]).startswith('※')]
        self.alias={r[0]:r[1] for r in wb['01_검사자명단'].iter_rows(min_row=2,min_col=8,max_col=9,values_only=True) if r[0] and r[1]}
        self.std={}; self.zone={}; self.cl={}; self.zo={}
        for a,b,c,dd,e in wb['04_지역표'].iter_rows(min_row=2,max_col=5,values_only=True):
            if a is None: continue
            a=str(a).strip(); b=vstr(b); c=vstr(c); dd=vstr(dd)
            if b: self.std[a]=b
            if c: self.zone[a]=c
            if b and not b.startswith('('):
                if dd and dd!='확인 필요' and not dd.startswith('데이터'): self.cl.setdefault(b,dd)
                if c: self.zo.setdefault(b,c)
        st={r[1]:r[2] for r in wb['02_설정'].iter_rows(min_row=4,max_row=15,min_col=1,max_col=3,values_only=True) if r[1]}
        self.set={k:v for k,v in st.items()}
        self.raw=[list(r) for r in raw]            # 원데이터 표 (행 리스트, 11열)
        self.screen={}                            # (row,col)->value
        self.merges=[]                            # 열 A 병합 (top,bottom,label)
        self.results=[]                           # 09 (행 리스트, 16열)
        self.logs=[]
        self.date=datetime.datetime(2026,10,14)
        self.dirty=False
    # ---- 설정
    def slot_min(self): return int(self.set['슬롯단위_분'])
    def disp_slots(self): return 1440//self.slot_min()
    def work_lo(self): h,m=map(int,str(self.set['운영시작시간']).split(':')); return (h*60+m)//self.slot_min()+1
    def work_hi(self): 
        h,m=map(int,str(self.set['운영종료시간']).split(':')); n=(h*60+m)-(int(str(self.set['운영시작시간']).split(':')[0])*60+int(str(self.set['운영시작시간']).split(':')[1]))
        return self.work_lo()+n//self.slot_min()-1
    def log(self,t,c,rel): self.logs.append((t,c,rel))
    def norm(self,s):
        t=vstr(s); return self.alias.get(t,t)
    def lead(self,s):
        t=vstr(s); p=t.find('-')
        if p<0: p=t.find(',')
        if p>0: t=t[:p].strip()
        return self.norm(t)

# ------------------------------------------------------------ DrawLedger (Module4)
def draw_ledger(env,d):
    env.screen={}; env.merges=[]
    key=d.toordinal()
    day=[]
    for i,r in enumerate(env.raw):
        if day_key(r[RW['DATE']])==key:
            s=vstr(r[RW['ZONE']])
            if not s:
                reg=vstr(r[RW['REG']])
                s=env.zone.get(reg,'') if reg else ''
            if not s: s='미분류'
            day.append((i,s))
    has_other=any(s not in (ZONE_A,ZONE_B) for _,s in day)
    zones=[ZONE_A,ZONE_B]+(['미분류'] if has_other else [])
    cur=SCR_FIRST_ROW
    for z,zn in enumerate(zones,1):
        reg=[];pre=[]
        for i,s in day:
            if s==zn or (z==3 and s not in (ZONE_A,ZONE_B)):
                (reg if is_regular(vstr(env.raw[i][RW['KIND']])) else pre).append(i)
        nrows=max(LEDGER_MIN_ROWS,len(reg)+2,len(pre)+2)
        for r,i in enumerate(reg):
            row=cur+r; R=env.raw[i]
            env.screen[(row,2)]=R[RW['NAME']]; env.screen[(row,3)]=R[RW['CAP']]; env.screen[(row,4)]=R[RW['REG']]
            env.screen[(row,5)]=R[RW['TIME']]; env.screen[(row,6)]=R[RW['FEE']]; env.screen[(row,7)]=R[RW['INSP']]; env.screen[(row,NOTE_L)]=R[RW['NOTE']]
        for r,i in enumerate(pre):
            row=cur+r; R=env.raw[i]
            env.screen[(row,8)]=R[RW['NAME']]; env.screen[(row,9)]=R[RW['CAP']]; env.screen[(row,10)]=R[RW['REG']]
            env.screen[(row,11)]=R[RW['TIME']]; env.screen[(row,12)]=R[RW['INSP']]; env.screen[(row,NOTE_R)]=R[RW['NOTE']]
        env.merges.append((cur,cur+nrows-1,'기타지역' if z==3 else zn))
        cur+=nrows
    env.note_row=cur
    return cur

def scan_sections(env):
    secs=[]; r=SCR_FIRST_ROW
    for top,bot,lab in env.merges:
        if top==r and lab and lab!='비고': secs.append((top,bot,lab)); r=bot+1
    return secs,r

# ------------------------------------------------------------ ReadScreen
def read_screen(env):
    secs,noteRow=scan_sections(env)
    rec=[]
    big=float(env.set['정식설비기준용량_kW']); bigm=int(env.set['정식설비소요시간_분']); smallm=int(env.set['약식설비소요시간_분']); onm=int(env.set['온라인업무소요시간_분'])
    for top,bot,lab in secs:
        zone='미분류' if lab=='기타지역' else lab
        for r in range(top,bot+1):
            for side in (0,1):
                c0=2 if side==0 else 8
                g=lambda c: env.screen.get((r,c))
                nm=vstr(g(c0)); capV=g(c0+1); reg=vstr(g(c0+2)); tmV=g(c0+3)
                if side==0: feeV=g(6); insp=vstr(g(7)); note=vstr(g(NOTE_L))
                else: feeV=None; insp=vstr(g(12)); note=vstr(g(NOTE_R))
                if nm or vstr(capV) or reg or vstr(tmV) or vstr(feeV) or insp:
                    t=dict(Key=r*10+side,ScrRow=r,Side=side,Zone=zone,TaskName=nm,CapV=capV,CapText=vstr(capV),Cap=cap_num(capV),
                           RegionRaw=reg,RegionStd='',TimeV=tmV,TimeStr=promise_time(tmV),FeeV=feeV,Fee=vstr(feeV),
                           InspRaw=insp,Insp=env.lead(insp),InspFinal=env.norm(insp),Note=note,Online='온라인' in note)
                    sd=env.std.get(reg,'')
                    t['RegionStd']='' if sd.startswith('(') else sd
                    if nm and t['CapText']=='' and reg=='' and vstr(tmV)=='' and t['Fee']=='' and insp=='': t['Kind']=KIND_MEMO
                    elif side==0: t['Kind']=KIND_REG
                    else: t['Kind']=KIND_PRE
                    t['DurMin']=onm if t['Online'] else (bigm if t['Cap']>=big else smallm)
                    t['Slots']=max(1,-(-t['DurMin']//env.slot_min()))
                    rec.append(t)
    return rec

def leave_from_recs(env,rec):
    res=set(); names=[n for n,_ in env.roster]
    for t in rec:
        if t['Kind']==KIND_MEMO and is_full_leave(t['TaskName']):
            for nm in names:
                if nm in t['TaskName']: res.add(nm)
    return res
def active_inspectors(env,leave):
    return [n for n,a in env.roster if a and n not in leave]

# ------------------------------------------------------------ 자동배분 (Module2)
def area_distance(env,a,b):
    if not a or not b: return 500
    if a==b: return 0
    if a in env.cl and b in env.cl and env.cl[a]==env.cl[b]: return 1
    if a in env.zo and b in env.zo and env.zo[a]==env.zo[b]: return 2
    return 3
def clear_results(env,d): env.results=[r for r in env.results if r[0]!=d]
def auto_assign(env,d,quiet=True):
    rec=read_screen(env); leave=leave_from_recs(env,rec)
    insp=active_inspectors(env,leave)
    if not insp: return None
    nslots=env.disp_slots(); lo,hi=env.work_lo(),env.work_hi()
    clear_results(env,d)
    occ={n:[0]*(nslots+2) for n in insp}; last={n:'' for n in insp}; cnt={n:0 for n in insp}
    tasks=[t for t in rec if is_task_kind(t['Kind']) and t['TaskName']]
    out=[]; placed=[0]; fail=[0]
    def free(n,s,dur): return not(s<1 or s+dur-1>nslots) and all(occ[n][k]==0 for k in range(s,s+dur))
    def first_free(n,dur):
        for s in range(lo,hi-dur+2):
            if free(n,s,dur): return s
        return 0
    def rank(reg): return sorted(insp,key=lambda n:(area_distance(env,reg,last[n])*1000+cnt[n]))  # 안정 정렬 = 버블정렬과 동일
    def slot_of(ts):
        h,m=map(int,ts.split(':')); return (h*60+m)//env.slot_min()+1
    def place(t,n,s,typ):
        for k in range(s,s+t['Slots']): occ[n][k]=t['Key']
        last[n]=t['RegionStd']; cnt[n]+=1; placed[0]+=1
        out.append([d,t['Key'],t['TaskName'],t['CapText'],t['RegionStd'],str(t['DurMin']),t['TimeStr'],n,s,'',t['Slots'],typ,'',t['Fee'],None,None])
    def failt(t,why):
        fail[0]+=1; env.log('미배정',why,t['TaskName'])
        out.append([d,t['Key'],t['TaskName'],t['CapText'],t['RegionStd'],str(t['DurMin']),t['TimeStr'],'',0,'',t['Slots'],0,'',t['Fee'],None,None])
    for t in tasks:
        if t['Insp'] and t['TimeStr']:
            s=slot_of(t['TimeStr'])
            if t['Insp'] not in insp: failt(t,'근무자없음')
            elif s<1 or s+t['Slots']-1>nslots: failt(t,'시간축')
            elif not free(t['Insp'],s,t['Slots']): failt(t,'충돌')
            else: place(t,t['Insp'],s,1)
    for t in tasks:
        if not t['Insp'] and t['TimeStr']:
            s=slot_of(t['TimeStr'])
            if s<1 or s+t['Slots']-1>nslots: failt(t,'시간축'); continue
            for n in rank(t['RegionStd']):
                if free(n,s,t['Slots']): place(t,n,s,2); break
            else: failt(t,'빈검사자없음')
    for t in tasks:
        if t['Insp'] and not t['TimeStr']:
            if t['Insp'] not in insp: failt(t,'근무자없음'); continue
            fs=first_free(t['Insp'],t['Slots'])
            if fs==0: failt(t,'근무시간내 빈시간없음')
            else: place(t,t['Insp'],fs,3)
    lunch=[x.strip() for x in str(env.set['점심탐색순서']).split(',')]; ls=max(1,int(env.set['점심시간길이_분'])//env.slot_min())
    for n in insp:
        for lt in lunch:
            s=slot_of(lt)
            if s>=1 and free(n,s,ls):
                for k in range(s,s+ls): occ[n][k]=-1
                out.append([d,0,'점심시간','','',str(ls*env.slot_min()),'',n,s,'',ls,5,'','',None,None]); break
        else: env.log('충돌','점심 못 넣음',n)
    for t in tasks:
        if not t['Insp'] and not t['TimeStr']:
            for n in rank(t['RegionStd']):
                fs=first_free(n,t['Slots'])
                if fs: place(t,n,fs,4); break
            else: failt(t,'마감')
    env.results+=out
    return dict(insp=insp,leave=leave,tasks=len(tasks),placed=placed[0],fail=fail[0],out=out,occ=occ)

def results_stale(env,d):
    rec=read_screen(env)
    seen={r[1] for r in env.results if r[0]==d and r[1]>0}
    nres=sum(1 for r in env.results if r[0]==d and r[1]>0)
    if nres==0: return False
    nT=0
    for t in rec:
        if is_task_kind(t['Kind']) and t['TaskName']:
            nT+=1
            if t['Key'] not in seen: return True
    return nT!=nres

# ------------------------------------------------------------ Day_Commit / CommitToRaw (Module4)
def check_recs(env,rec):
    known={n for n,_ in env.roster}; msg=[]; unknown=[]
    for t in rec:
        if t['TaskName']=='': msg.append(f"{t['Zone']} {t['ScrRow']}행 상호 없음")
        if is_task_kind(t['Kind']) and t['InspFinal']:
            for p in re.split('-',t['InspFinal'].replace(',','-')):
                p=env.norm(p)
                if p and p not in known: unknown.append((t['Zone'],t['ScrRow'],p))
    return msg,unknown

def commit_to_raw(env,d,rec):
    old=env.raw; key=d.toordinal(); maxid=0; keep=[]; lastLE=0
    for r in old:
        if isinstance(r[RW['ID']],(int,float)) and r[RW['ID']]>maxid: maxid=r[RW['ID']]
        dk=day_key(r[RW['DATE']])
        if dk==-1 and r[RW['NAME']] is None and r[RW['KIND']] is None: continue
        if dk!=key:
            keep.append(r)
            if 0<=dk<=key: lastLE=len(keep)
    new=[]
    for t in rec:
        maxid+=1
        row=[None]*11
        row[RW['DATE']]=d; row[RW['KIND']]=t['Kind']
        if t['Zone']: row[RW['ZONE']]=t['Zone']
        if t['TaskName']: row[RW['NAME']]=t['TaskName']
        if t['CapText']: row[RW['CAP']]=t['CapV']
        if t['RegionRaw']: row[RW['REG']]=t['RegionRaw']
        if vstr(t['TimeV']): row[RW['TIME']]=t['TimeV']
        if t['Fee']: row[RW['FEE']]=t['FeeV']
        if t['InspFinal']: row[RW['INSP']]=t['InspFinal']
        if t['Note']: row[RW['NOTE']]=t['Note']
        row[RW['ID']]=maxid
        new.append(row)
    env.raw=keep[:lastLE]+new+keep[lastLE:]

def commit(env,d,ask_unknown=True,apply_results=True,ask_noname=True):
    rec=read_screen(env)
    msg,unknown=check_recs(env,rec)
    if msg and not ask_noname: return ('ERR',msg)
    if unknown and not ask_unknown: return ('CANCEL',unknown)
    if not rec and not any(day_key(r[RW['DATE']])==d.toordinal() for r in env.raw): return ('EMPTY',)
    res={r[1]:r for r in env.results if r[0]==d and r[1]>0}
    hasRes=bool(res)
    useRes=hasRes and not results_stale(env,d)
    nfill=nchg=nblank=0
    for t in rec:
        if is_task_kind(t['Kind']) and t['TaskName']:
            k=t['Key']
            if apply_results and useRes and k in res and res[k][7] and res[k][8]>0 and res[k][2]==t['TaskName']:
                if t['InspFinal']=='': t['InspFinal']=res[k][7]; nfill+=1
                elif env.lead(t['InspFinal'])!=res[k][7] and res[k][12]=='Y': t['InspFinal']=res[k][7]; nchg+=1
            if t['InspFinal']=='': nblank+=1
    commit_to_raw(env,d,rec)
    for t in rec:
        if t['InspFinal']!=t['InspRaw']:
            env.screen[(t['ScrRow'],7 if t['Side']==0 else 12)]=t['InspFinal']
    env.dirty=False
    return ('OK',dict(n=len(rec),fill=nfill,chg=nchg,blank=nblank,stale=(hasRes and not useRes)))
