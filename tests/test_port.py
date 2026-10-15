import datetime, collections, copy, sys
import openpyxl
import vbaport as V
from vbaport import *
XLSM=sys.argv[1]
_wb=openpyxl.load_workbook(XLSM,keep_vba=False)
RAW=[list(r) for r in _wb['10_원데이터'].iter_rows(min_row=2,max_col=11,values_only=True) if r[0] is not None]
V.XLSM=XLSM
FAILS=[]; OKS=[0]
def check(cond,msg):
    if cond: OKS[0]+=1
    else: FAILS.append(msg); print('FAIL:',msg)
def norm_row(r):  # 비교용: ID 제외
    return tuple(r[:10])
E_=lambda: Env(RAW,XLSM)

# ---------- T1: 그리기 -> 읽기 왕복 (모든 날짜)
env=E_()
dates=sorted({r[0] for r in RAW})
bad=0; tot=0; examples=[]
for d in dates:
    draw_ledger(env,d)
    rec=read_screen(env)
    day=[r for r in RAW if r[0]==d]
    check(len(rec)==len(day),f'T1 건수 {d.date()} {len(rec)}!={len(day)}')
    # 비교: 이름/종류/권역/용량/지역/시간/수수료/검사자/비고 (집합 비교)
    def key_raw(r):
        side=0 if r[1]=='정기검사' else 1
        kind=r[1]
        nm=vstr(r[3]); cap=vstr(r[4]); reg=vstr(r[5]); tm=vstr(r[6]); fee=vstr(r[7]) if side==0 else ''; insp=vstr(r[8])
        if nm and not cap and not reg and not tm and not fee and not insp: kind='메모'
        return (kind,r[2] if r[2] in('부천지역','김포지역') else '미분류',nm,cap,reg,tm,fee,env.norm(insp),vstr(r[9]))
    def key_rec(t):
        return (t['Kind'],t['Zone'],t['TaskName'],t['CapText'],t['RegionRaw'],vstr(t['TimeV']),t['Fee'] if t['Side']==0 else '',t['InspFinal'],t['Note'])
    a=collections.Counter(key_raw(r) for r in day); b=collections.Counter(key_rec(t) for t in rec)
    tot+=1
    if a!=b:
        bad+=1
        if len(examples)<3: examples.append((d.date(),list((a-b).items())[:2],list((b-a).items())[:2]))
check(bad==0,f'T1 왕복 불일치 {bad}/{tot} {examples}')
print('T1 dates',tot,'bad',bad)

# ---------- T2: 변경 없이 입력 완료 -> 원데이터 동일(ID 제외), 다른 날짜 불변
for d in [datetime.datetime(2026,10,14),datetime.datetime(2026,6,17),datetime.datetime(2024,1,12),datetime.datetime(2026,11,4),datetime.datetime(2026,3,31)]:
    env=E_(); before=copy.deepcopy(env.raw); draw_ledger(env,d)
    r=commit(env,d)
    if r[0]=='ERR': print('  (상호 없는 줄이 있어 차단:',d.date(),r[1],')'); continue
    check(r[0]=='OK',f'T2 commit {d.date()} {r}')
    after=env.raw
    check(len(after)==len(before),f'T2 건수 변화 {d.date()} {len(before)}->{len(after)}')
    ob=[norm_row(x) for x in before if x[0]!=d]; oa=[norm_row(x) for x in after if x[0]!=d]
    check(ob==oa,f'T2 다른 날짜 변경 {d.date()}')
    # 같은 날짜 행 집합 동일 (정렬 순서 외)
    sb=collections.Counter(norm_row(x) for x in before if x[0]==d); sa=collections.Counter(norm_row(x) for x in after if x[0]==d)
    # 메모 cap 이동 등은 이미 데이터에서 정리됨 -> 동일해야 함
    check(sb==sa,f'T2 같은 날짜 행 불일치 {d.date()} {(sb-sa).most_common(2)} {(sa-sb).most_common(2)}')
    dts=[x[0] for x in after]
    check(dts==sorted(dts),f'T2 날짜 정렬 깨짐 {d.date()}')
    ids=[x[10] for x in after]; check(len(set(ids))==len(ids),f'T2 ID 중복 {d.date()}')

# ---------- T3: 새 날짜에 입력/삽입 위치/삭제
env=E_(); d=datetime.datetime(2027,1,5)
draw_ledger(env,d); rec=read_screen(env); check(len(rec)==0,'T3 빈 날짜')
secs,_=scan_sections(env); check(len(secs)==2 and secs[0][1]-secs[0][0]+1==24,'T3 구역 24행')
env.screen[(7,2)]='테스트상호'; env.screen[(7,3)]=300; env.screen[(7,4)]='중동'; env.screen[(7,5)]=datetime.time(10,0); env.screen[(7,6)]=500000; env.screen[(7,7)]='허창수'
env.screen[(secs[1][0],8)]='한태희휴가'                               # 김포 구역 사용전칸 메모
env.screen[(secs[1][0]+1,2)]='김포업체'; env.screen[(secs[1][0]+1,4)]='학운리'
r=commit(env,d); check(r[0]=='OK' and r[1]['n']==3,f'T3 commit {r}')
new=[x for x in env.raw if x[0]==d]; check(len(new)==3,'T3 새 행 3')
check([x[0] for x in env.raw]==sorted(x[0] for x in env.raw),'T3 날짜 정렬')
kinds={x[3]:(x[1],x[2]) for x in new}
check(kinds['테스트상호']==('정기검사','부천지역') and kinds['한태희휴가']==('메모','김포지역') and kinds['김포업체']==('정기검사','김포지역'),f'T3 분류 {kinds}')
check(max(x[10] for x in env.raw)==len(RAW)+3,'T3 ID')
# 삭제 (화면 비우고 저장)
for k in list(env.screen): env.screen[k]=None
r=commit(env,d); check(r[0]=='OK' and not any(x[0]==d for x in env.raw) and len(env.raw)==len(RAW),f'T3 삭제 {r} {len(env.raw)}')
# 과거 날짜 중간 삽입
env=E_(); d=datetime.datetime(2026,5,31)
draw_ledger(env,d); secs,_=scan_sections(env); env.screen[(7,2)]='중간삽입'
r=commit(env,d); idx=[i for i,x in enumerate(env.raw) if x[3]=='중간삽입']
check(r[0]=='OK' and len(idx)==1 and env.raw[idx[0]-1][0]<=d and env.raw[idx[0]+1][0]>=d,f'T3 중간삽입 {r}')

# ---------- T4: 모든 미래 날짜 자동배분
env=E_(); stats=collections.Counter(); problems=[]
fut=[d for d in dates if d>=datetime.datetime(2026,10,14)]
for d in fut:
    draw_ledger(env,d); r=auto_assign(env,d)
    if r is None: stats['no inspectors']+=1; continue
    insp=r['insp']; lo,hi=env.work_lo(),env.work_hi()
    # 검사자 중복 없음, 점심 포함 겹침 없음
    per=collections.defaultdict(list)
    for row in r['out']:
        if row[7] and row[8]>0: per[row[7]].append((row[8],row[8]+row[10]-1,row[2]))
    for n,iv in per.items():
        iv.sort()
        for a,b in zip(iv,iv[1:]):
            if a[1]>=b[0]: problems.append((d.date(),n,a,b))
    # 휴가자 배치 없음
    for n in r['leave']:
        if n in per: problems.append((d.date(),'휴가자 배치',n))
    # 시간 미지정 업무는 근무시간 안
    for row in r['out']:
        if row[11] in (3,4,5) and row[8]>0 and not(lo<=row[8] and row[8]+row[10]-1<=hi): problems.append((d.date(),'근무시간밖',row[2],row[8]))
    check(r['placed']+r['fail']==r['tasks'],f"T4 건수 {d.date()} {r['placed']}+{r['fail']}!={r['tasks']}")
    lunches=sum(1 for x in r['out'] if x[11]==5); check(lunches<=len(insp),f'T4 점심 {d.date()}')
    nolunch=len(insp)-lunches; stats['점심 못 넣음']+=nolunch
    keys=[x[1] for x in r['out'] if x[1]>0]; check(len(keys)==len(set(keys)),f'T4 키중복 {d.date()}')
    check(not results_stale(env,d),f'T4 stale {d.date()}')
    stats['days']+=1; stats['tasks']+=r['tasks']; stats['fail']+=r['fail']; stats['placed']+=r['placed']
check(not problems,f'T4 문제 {problems[:5]}')
print('T4',dict(stats))

# ---------- T5/T6: 배분 -> 입력 완료 시 검사자 채움, 수동조정 반영, 이름 불일치 방어
env=E_(); d=datetime.datetime(2026,10,19); draw_ledger(env,d); r=auto_assign(env,d)
rec=read_screen(env); blanks=[t for t in rec if is_task_kind(t['Kind']) and t['InspFinal']=='']
check(len(blanks)>0,'T6 미지정 업무 존재')
res=commit(env,d); check(res[0]=='OK' and res[1]['fill']==len(blanks),f'T6 채움 {res}')
rec2=read_screen(env); check(all(t['InspFinal'] for t in rec2 if is_task_kind(t['Kind']) and t['TaskName']),'T6 화면 검사자 채워짐')
d2=[x for x in env.raw if x[0]==d and x[1]!='메모']; check(all(x[8] for x in d2),'T6 원데이터 검사자 채워짐')
# 수동조정: 한 건을 다른 검사자로 이동한 상태 시뮬레이션
env=E_(); draw_ledger(env,d); auto_assign(env,d)
tgt=[r for r in env.results if r[0]==d and r[1]>0 and r[11]==4][0]
other=[n for n in active_inspectors(env,set()) if n!=tgt[7]][0]; tgt[7]=other; tgt[12]='Y'
res=commit(env,d); check(res[0]=='OK',f'T6 수동 {res}')
row=[x for x in env.raw if x[0]==d and x[3]==tgt[2]][0]; check(row[8]==other,f'T6 수동조정 반영 {row[8]} vs {other}')
# 화면 변경 후(stale) 저장 -> 배분 결과 미반영
env=E_(); draw_ledger(env,d); auto_assign(env,d)
secs,_=scan_sections(env); env.screen[(secs[0][1],2)]='추가업무'; env.screen[(secs[0][1],4)]='중동'
check(results_stale(env,d),'T5 stale 감지')
res=commit(env,d); check(res[0]=='OK' and res[1]['stale'] and res[1]['fill']==0,f'T5 stale 저장 {res}')
# 이름 바꿔치기 (같은 키, 다른 상호) -> 결과 적용 안 됨 (stale 아님)
env=E_(); draw_ledger(env,d); auto_assign(env,d)
r0=[r for r in env.results if r[0]==d and r[1]>0 and r[11]==4][0]; kk=r0[1]
env.screen[(kk//10,2 if kk%10==0 else 8)]='다른상호'
res=commit(env,d); check(res[0]=='OK',f'T5 이름불일치 {res}')
row=[x for x in env.raw if x[0]==d and x[3]=='다른상호'][0]; check(row[8] is None,f'T5 이름 다르면 검사자 미반영 {row[8]}')

# ---------- T7: 입력 점검
env=E_(); d=datetime.datetime(2027,2,2); draw_ledger(env,d)
env.screen[(7,3)]=100                                   # 상호 없음
r=commit(env,d,ask_noname=False); check(r[0]=='ERR',f'T7 상호없음 확인창 {r}')
env.screen[(7,3)]=None; env.screen[(7,2)]='X'; env.screen[(7,7)]='허창소'
r=commit(env,d,ask_unknown=False); check(r[0]=='CANCEL',f'T7 미등록 검사자 {r}')
env.screen[(7,7)]='옹승환'; r=commit(env,d); check(r[0]=='OK' and [x for x in env.raw if x[3]=='X'][0][8]=='옹승완',f'T7 별칭 정정 {r}')
env.screen[(8,2)]='공동'; env.screen[(8,7)]='이덕기-김영권'; r=commit(env,d,ask_unknown=False); check(r[0]=='OK',f'T7 공동검사 허용 {r}')
r=auto_assign(env,d); check(r is not None,'T7 공동검사 배분')
# 빈 화면 + 저장된 데이터 없음
env=E_(); draw_ledger(env,datetime.datetime(2027,3,3)); check(commit(env,datetime.datetime(2027,3,3))[0]=='EMPTY','T7 빈 저장')

# ---------- T8: 휴가 메모로 열 제외 / 전원 휴가 / 오늘근무 FALSE
env=E_(); d=datetime.datetime(2027,4,1); draw_ledger(env,d); secs,_=scan_sections(env)
for i,n in enumerate([n for n,a in env.roster if a]): env.screen[(secs[0][0]+i,8)]=n+'휴가'
r=auto_assign(env,d); check(r is None,'T8 전원 휴가 -> 배분 불가')
env.screen={k:v for k,v in env.screen.items() if k[1]!=8}; env.screen[(7,8)]='한태희 오후반차'; 
r=auto_assign(env,d); check(r is not None and '한태희' in r['insp'],'T8 반차는 제외 안 함')

# ---------- T9: 자정 넘는 / 근무시간 밖 / 겹침 충돌
env=E_(); d=datetime.datetime(2027,5,3); draw_ledger(env,d)
env.screen[(7,2)]='새벽'; env.screen[(7,3)]=100; env.screen[(7,5)]=datetime.time(6,0); env.screen[(7,7)]='허창수'
env.screen[(8,2)]='야간'; env.screen[(8,3)]=100; env.screen[(8,5)]=datetime.time(23,30); env.screen[(8,7)]='허창수'
env.screen[(9,2)]='겹침1'; env.screen[(9,3)]=100; env.screen[(9,5)]=datetime.time(10,0); env.screen[(9,7)]='이덕기'
env.screen[(10,2)]='겹침2'; env.screen[(10,3)]=100; env.screen[(10,5)]=datetime.time(10,30); env.screen[(10,7)]='이덕기'
r=auto_assign(env,d); o={x[2]:x for x in r['out']}
check(o['새벽'][8]==13 and o['새벽'][11]==1,f'T9 06:00 배치 {o["새벽"]}')
check(o['야간'][11]==0,f'T9 23:30 + 90분은 자정 넘어 실패 {o["야간"]}')
check(o['겹침1'][11]==1 and o['겹침2'][11]==0,'T9 겹침 실패 기록')

print('OK checks',OKS[0],'FAILS',len(FAILS))
for f in FAILS[:20]: print(' -',f)
