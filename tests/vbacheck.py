import re,glob,sys
d=sys.argv[1]
def strip(line):
    out=[];ins=False;i=0
    while i<len(line):
        ch=line[i]
        if ch=='"':
            if ins and i+1<len(line) and line[i+1]=='"': i+=2; continue
            ins=not ins; out.append('"'); i+=1; continue
        if ch=="'" and not ins: break
        out.append(' ' if ins else ch); i+=1
    return ''.join(out), ins
bad=0
for fn in sorted([x for e in ('bas','cls','frm') for x in glob.glob(d+'/*.'+e)]):
    txt=open(fn,encoding='utf8').read()
    try: txt.encode('cp949')
    except UnicodeEncodeError as e: print(fn,'cp949',txt[e.start-15:e.end+15]); bad+=1
    lines=txt.splitlines(); joined=[]; buf=''
    for n,l in enumerate(lines,1):
        s,ins=strip(l)
        if ins: print(fn,n,'unterminated string'); bad+=1
        if s.rstrip().endswith(' _') or s.rstrip()=='_': buf+=s.rstrip()[:-1]+' '; continue
        joined.append((n,buf+s)); buf=''
    stack=[]
    def pop(k,n):
        global bad
        if not stack or stack[-1][0]!=k: print(fn.split('/')[-1],n,'mismatch end',k,stack[-2:]); bad+=1
        else: stack.pop()
    for n,ln in joined:
        t=ln.strip()
        if not t: continue
        if t.count('(')!=t.count(')'): print(fn,n,'paren',t[:90]); bad+=1
        if re.match(r'(If|ElseIf)\b',t) and not re.search(r'\bThen\b',t): print(fn,n,'noThen',t[:90]); bad+=1
        if re.search(r'(\+=|-=|!=|==|&&|\|\|)',t.replace('"','')): print(fn,n,'cism',t[:90]); bad+=1
        if re.match(r'(Public |Private |Friend )?(Static )?Sub\b',t): stack.append(('Sub',n))
        elif re.match(r'(Public |Private |Friend )?(Static )?Function\b',t): stack.append(('Function',n))
        elif re.match(r'(Public )?Type\b',t): stack.append(('Type',n))
        elif t=='End Sub': pop('Sub',n)
        elif t=='End Function': pop('Function',n)
        elif t=='End Type': pop('Type',n)
        elif t=='End If': pop('If',n)
        elif t=='End With': pop('With',n)
        elif t=='End Select': pop('Select',n)
        elif re.match(r'With\b',t): stack.append(('With',n))
        elif re.match(r'Select Case\b',t): stack.append(('Select',n))
        elif re.match(r'For\b',t): stack.append(('For',n))
        elif re.match(r'Next\b',t): pop('For',n)
        elif re.match(r'Do\b',t): stack.append(('Do',n))
        elif re.match(r'Loop\b',t): pop('Do',n)
        elif re.match(r'If\b.*\bThen\s*$',t): stack.append(('If',n))
    if stack: print(fn,'unclosed',stack); bad+=1
print('problems',bad)
