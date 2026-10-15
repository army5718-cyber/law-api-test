import re,glob,sys,collections
files=sorted([x for e in ('bas','cls','frm') for x in glob.glob(sys.argv[1]+'/*.'+e)])
KW=set('''Option Explicit Public Private Const Dim As Long String Boolean Double Date Variant Object Integer Byte Sub Function End If Then Else ElseIf For To Next Each In Do While Loop Until Select Case Exit Set New Not And Or Xor Mod ByVal ByRef Optional Type With GoTo Resume On Error Call Step True False Nothing Empty Null Is Like ReDim Preserve Static Property Get Let WithEvents Friend Enum Declare Attribute VB_Name VB_Base VB_GlobalNameSpace VB_Creatable VB_PredeclaredId VB_Exposed VB_TemplateDerived VB_Customizable Me Then ParamArray
Err Number Description Raise Application Worksheet Range Shape Shapes ListObject Collection Scripting Dictionary
CStr CLng CDbl CDate CBool Int Abs Round Len Left Right Mid Trim InStr Format Chr ChrW Choose IIf IsNumeric IsDate IsEmpty IsError UCase LCase Split Join Array UBound LBound TimeValue TimeSerial DateSerial DateValue DateAdd Year Month Day Weekday Now Date Time Val Replace Environ TypeName VarType MsgBox CreateObject Rnd Fix Sgn Sqr
vbInformation vbExclamation vbCritical vbYesNo vbYes vbQuestion vbCrLf vbCr vbLf vbDate vbDouble vbString vbModal vbMonday vbSunday vbFormControlMenu vbObjectError VbMsgBoxStyle
xlUp xlToLeft xlCenter xlLeft xlRight xlTop xlContinuous xlThin xlMedium xlDot xlInsideVertical xlEdgeLeft xlEdgeRight xlEdgeTop xlEdgeBottom xlVertical xlMove xlNone xlPaperA4 xlPaperA3 xlLandscape xlTypePDF xlQualityStandard xlCalculationManual XlCalculation
msoShapeRectangle msoShapeRoundedRectangle msoFalse msoTrue msoAnchorMiddle msoAnchorTop msoAlignCenter msoAlignLeft msoLineDash
Worksheet Workbook ThisWorkbook ActiveSheet Selection Target Cancel CloseMode Sheets rows Rows Columns Cells Names names
'''.split())
def strip(line):
    out=[];ins=False;i=0
    while i<len(line):
        ch=line[i]
        if ch=='"':
            if ins and i+1<len(line) and line[i+1]=='"': i+=2; continue
            ins=not ins; out.append(' '); i+=1; continue
        if ch=="'" and not ins: break
        out.append(' ' if ins else ch); i+=1
    return ''.join(out)
allproc={}
modvars={}
for fn in files:
    src=open(fn,encoding='utf8').read()
    for m in re.finditer(r'^\s*(?:Public |Private |Friend )?(?:Static )?(?:Sub|Function)\s+([\w가-힣]+)',src,re.M): allproc[m.group(1)]=fn
    for m in re.finditer(r'^\s*(?:Public|Private|Global)\s+(?:Const\s+)?([\w가-힣]+)',src,re.M):
        if m.group(1) not in ('Sub','Function','Type','Const'): allproc[m.group(1)]=fn
    for m in re.finditer(r'^\s*Public Type\s+(\w+)',src,re.M): allproc[m.group(1)]=fn
glob_names=set(allproc)
members=set()
for fn in files:
    src=open(fn,encoding='utf8').read()
    # UDT members
    for m in re.finditer(r'Type TaskRec(.*?)End Type',src,re.S):
        for l in m.group(1).splitlines():
            mm=re.match(r'\s*(\w+)\s+As',l)
            if mm: members.add(mm.group(1))
problems=collections.defaultdict(set)
for fn in files:
    lines=open(fn,encoding='utf8').read().splitlines()
    # join continuation
    joined=[];buf=''
    for l in lines:
        s=strip(l)
        if s.rstrip().endswith('_'): buf+=s.rstrip()[:-1]+' '; continue
        joined.append(buf+s); buf=''
    declared_mod=set()
    proc_decl=None
    for ln in joined:
        t=ln.strip()
        if re.match(r'(Public |Private |Friend )?(Static )?(Sub|Function|Property)\b',t):
            proc_decl=set()
            # params
            m=re.search(r'\((.*)\)',t)
            if m:
                for p in m.group(1).split(','):
                    pm=re.search(r'([\w가-힣]+)\s*(\(\))?\s*(As|$)',p.strip())
                    pm2=re.match(r'(?:Optional\s+)?(?:ByVal\s+|ByRef\s+)?([\w가-힣]+)',p.strip())
                    if pm2: proc_decl.add(pm2.group(1))
            nm=re.search(r'(?:Sub|Function|Property \w+)\s+([\w가-힣]+)',t)
            if nm: proc_decl.add(nm.group(1))
        # declarations
        for m in re.finditer(r'\b(?:Dim|Static|ReDim|Private|Public|Const)\s+(?:Preserve\s+)?(.*)',t):
            decl=m.group(1)
            if re.match(r'(Sub|Function|Const|Type)\b',decl) and False: pass
            for part in re.split(r',(?![^(]*\))',decl):
                pm=re.match(r'\s*(?:WithEvents\s+)?([\w가-힣]+)',part)
                if pm and (proc_decl is not None): proc_decl.add(pm.group(1))
                elif pm: declared_mod.add(pm.group(1))
        for m in re.finditer(r'\bFor(?: Each)?\s+([\w가-힣]+)\s+(?:In|=)',t):
            if proc_decl is not None: proc_decl.add(m.group(1))
        if re.match(r'End (Sub|Function|Property)',t): 
            proc_decl=None; continue
        # check tokens
        body=re.sub(r'^\s*\d+\s*','',t)
        if re.match(r'(Attribute|Option)\b',body): continue
        for m in re.finditer(r'(?<![.\w가-힣!])([A-Za-z_가-힣][\w가-힣]*)',body):
            w=m.group(1)
            if w in KW or w in glob_names or w in members: continue
            if proc_decl is not None and w in proc_decl: continue
            if w in declared_mod: continue
            # label?
            if re.match(r'^[\w가-힣]+:\s*$',body) and body.rstrip(': ')==w: continue
            problems[fn].add(w)
for fn,s in problems.items(): print(fn,sorted(s))
