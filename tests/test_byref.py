import re,glob,sys,collections
SRC=sys.argv[1]
def strip(line):
    out=[];ins=False;i=0
    while i<len(line):
        ch=line[i]
        if ch=='"':
            if ins and i+1<len(line) and line[i+1]=='"': out.append('  '); i+=2; continue
            ins=not ins; out.append('"'); i+=1; continue
        if ch=="'" and not ins: break
        out.append(ch); i+=1
    return ''.join(out)
def logical(txt):
    res=[];buf=''
    for l in txt.splitlines():
        s=strip(l)
        if s.rstrip().endswith(' _'): buf+=s.rstrip()[:-1]+' '; continue
        res.append(buf+s); buf=''
    return res
def split_top(s):
    a=[];d=0;c='';ins=False
    for ch in s:
        if ch=='"': ins=not ins
        if not ins:
            if ch=='(': d+=1
            if ch==')': d-=1
        if ch==',' and d==0 and not ins: a.append(c); c=''
        else: c+=ch
    if c.strip(): a.append(c)
    return a
files={f:open(f,encoding='utf8').read() for f in [x for e in ('bas','cls','frm') for x in glob.glob(SRC+'/*.'+e)]}
sigs={}
modvars=collections.defaultdict(dict)
procvars={}
for f,txt in files.items():
    mod=re.search(r'VB_Name = "([^"]+)"',txt).group(1)
    cur=None
    for ln in logical(txt):
        t=ln.strip()
        m=re.match(r'(Public |Private |Friend )?(Static )?(Sub|Function)\s+([\w가-힣]+)\s*(\((.*)\))?',t)
        if m:
            cur=(mod,m.group(4)); params=[]
            for p in split_top(m.group(6) or ''):
                p=p.strip(); byval=bool(re.match(r'(Optional\s+)?ByVal',p))
                pm=re.match(r'(?:Optional\s+)?(?:ByVal\s+|ByRef\s+)?([\w가-힣]+)(\(\))?\s*(?:As\s+([\w.]+))?',p)
                params.append((pm.group(1),byval,(pm.group(3) or 'Variant')+('()' if pm.group(2) else '')))
            sigs[m.group(4)]=sigs.get(m.group(4),[])+[(mod,params)]
            procvars[cur]={n:ty for n,bv,ty in params}
            continue
        if t in ('End Sub','End Function'): cur=None; continue
        for dm in re.finditer(r'\b(?:Dim|Private|Public|Static)\s+(.*)',t):
            if re.match(r'(Sub|Function|Const|Type)\b',dm.group(1)): continue
            for part in split_top(dm.group(1)):
                pm=re.match(r'\s*([\w가-힣]+)\s*(\([^)]*\))?\s*(?:As\s+(New\s+)?([\w.]+))?',part)
                if pm and pm.group(1):
                    ty=(pm.group(4) or 'Variant')+('()' if pm.group(2) is not None else '')
                    (procvars[cur] if cur else modvars[mod])[pm.group(1)]=ty
problems=0; checked=0
for f,txt in files.items():
    mod=re.search(r'VB_Name = "([^"]+)"',txt).group(1); cur=None
    for ln in logical(txt):
        t=ln.strip()
        m=re.match(r'(Public |Private |Friend )?(Static )?(Sub|Function)\s+([\w가-힣]+)',t)
        if m: cur=(mod,m.group(4)); continue
        if t in('End Sub','End Function'): cur=None; continue
        if not cur: continue
        vars_={**modvars[mod],**procvars.get(cur,{})}
        u=re.sub(r'"[^"]*"','""',t)
        for name,defs in sigs.items():
            if len(name)<3: continue
            mm=re.match(r'^(?:Call\s+)?'+re.escape(name)+r'(?![\w가-힣])\s*(.*)$',u) or None
            hits=[]
            if mm and not u.startswith(name+' =') : hits.append(mm.group(1))
            for m2 in re.finditer(r'(?<![\w가-힣.])'+re.escape(name)+r'\(',u):
                rest=u[m2.end()-1:]; d=0
                for i,ch in enumerate(rest):
                    if ch=='(': d+=1
                    if ch==')':
                        d-=1
                        if d==0: hits.append(rest[1:i]); break
            for argstr in hits:
                if argstr.startswith('=') : continue
                args=split_top(argstr)
                cand=[d for d in defs if len(d[1])>=len(args)]
                if not cand: continue
                params=cand[0][1]
                for (pn,byval,pty),a in zip(params,args):
                    a=a.strip()
                    if byval or pty in('Variant',) or a.startswith('ByVal'): continue
                    if re.fullmatch(r'[\w가-힣]+',a) and a in vars_:
                        checked+=1
                        vty=vars_[a]
                        # 배열 파라미터 표기 정규화
                        if vty!=pty and not (pty=='Object' and vty not in ('Long','String','Double','Boolean','Date')):
                            print('BYREF 불일치',mod,cur[1],name,pn,pty,'<-',a,vty); problems+=1
print('ByRef 인자 점검',checked,'불일치',problems)
