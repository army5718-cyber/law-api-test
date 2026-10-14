"""11_일일대장 미리 그려진 레이아웃 점검: 구역 병합 구간, 건당 2행 병합, 아래 행 비어 있음, 건수 = 원데이터 해당일 건수, 틀 고정."""
import sys, openpyxl
wb=openpyxl.load_workbook(sys.argv[1],keep_vba=True); ws=wb['11_일일대장']
bad=0
def chk(c,m):
    global bad
    if not c: bad+=1; print('FAIL',m)
chk(ws.freeze_panes=='A7','틀 고정 A7')
r=7; n=0; secs=0
while ws.cell(r,1).value not in (None,'비고'):
    m=[x for x in ws.merged_cells.ranges if x.min_row==r and x.min_col==1][0]
    chk((m.max_row-m.min_row+1)%2==0,'구역 행 수 짝수'); secs+=1
    for rr in range(r,m.max_row+1,2):
        for c in range(2,13):
            chk(any(x.min_row==rr and x.max_row==rr+1 and x.min_col==c==x.max_col for x in ws.merged_cells.ranges),f'병합 {rr},{c}')
            chk(ws.cell(rr+1,c).value is None,f'아래 행 비어 있음 {rr+1},{c}')
        n+=(ws.cell(rr,2).value is not None)+(ws.cell(rr,8).value is not None)
    r=m.max_row+1
d=ws['H4'].value
import re
mo,da=map(int,re.findall(r'\d+',d)[:2])
raw=wb['10_원데이터']
k=sum(1 for row in raw.iter_rows(min_row=2,values_only=True) if row[0] and row[0].month==mo and row[0].day==da and row[0].year==2026)
chk(n==k,f'건수 {n} != 원데이터 {k}')
print('sections',secs,'entries',n,'raw',k,'OK' if not bad else 'FAILS',bad)
sys.exit(1 if bad else 0)
