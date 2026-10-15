"""엑셀이 '컴파일 오류'로 막는 선언 규칙 점검: (1) 모듈 수준 선언은 첫 프로시저보다 위, (2) 한 프로시저 안 변수 중복 선언,
   (3) 같은 모듈 안 프로시저/상수/변수 이름 중복, (4) 모듈 간 Public 프로시저 이름 중복."""
import re,glob,sys,collections
d=sys.argv[1]; bad=0
procs=collections.defaultdict(list)
PROC=re.compile(r'^(Public |Private |Friend )?(Static )?(Sub|Function|Property (?:Get|Let|Set))\s+(\w+)',re.I)
for p in sorted(glob.glob(d+'/*.bas')+glob.glob(d+'/*.cls')+glob.glob(d+'/*.frm')):
    lines=open(p,encoding='utf8').read().split('\n'); inproc=False; seen=False; dims=None; names=collections.Counter(); modvars=set()
    for i,l in enumerate(lines,1):
        t=re.sub(r"'.*$","",l).strip() if '"' not in l else l.strip()
        m=PROC.match(t)
        if m and not inproc:
            inproc=True; seen=True; dims={}; nm=m.group(4); names[nm.lower()]+=1
            if (m.group(1) or 'Public ').strip().lower()!='private': procs[nm.lower()].append(p)
            # 매개변수도 이름 목록에 포함
            for a in re.findall(r'(?:ByVal |ByRef |Optional |ParamArray )*(\w+)(?:\(\))?\s+As\b',t.split('(',1)[1] if '(' in t else ''):
                dims[a.lower()]=i
            continue
        if re.match(r'End (Sub|Function|Property)\b',t): inproc=False; continue
        if not inproc:
            if seen and re.match(r'(Public |Private |Dim |Const |Type |Enum |Declare )',t) and not l.lstrip().startswith("'"):
                print('FAIL 프로시저 뒤 모듈 선언',p,i,t[:60]); bad+=1
            m2=re.match(r'(?:Public |Private )?(?:Const )?(\w+)\s+As\b',t)
            continue
        m=re.match(r'(?:Dim|Static)\s+(.*)$',t,re.I)
        if m and 'ReDim' not in t:
            for part in re.split(r',\s*(?=\w+(?:\([^)]*\))?\s+As\b|\w+\s*(?:,|$))',m.group(1)):
                mm=re.match(r'(\w+)',part.strip())
                if mm:
                    k=mm.group(1).lower()
                    if k in dims: print('FAIL 중복 선언',p,i,mm.group(1),'(앞선 선언:',dims[k],')'); bad+=1
                    dims[k]=i
    for k,v in names.items():
        if v>1: print('FAIL 프로시저 이름 중복',p,k); bad+=1
for k,v in procs.items():
    if len(v)>1 and not all(x.endswith('.cls') for x in v): print('FAIL 모듈 간 Public 프로시저 이름 중복',k,v); bad+=1
print('decl FAILS',bad); sys.exit(1 if bad else 0)
