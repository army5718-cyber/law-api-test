"""vbaProject.bin 생성기 (MS-OVBA / MS-CFB).
 - 소스 텍스트만 담은 프로젝트를 만든다(컴파일 캐시 없음 -> Excel 이 열 때 소스에서 다시 컴파일).
 - UserForm 디자이너 저장소(frmCalendar)는 원본 vbaProject.bin 에서 그대로 복사한다."""
import struct, olefile, re, glob, os

CP='cp949'

# ------------------------------------------------------------------ MS-OVBA 압축
def _chunk_compress(chunk):
    out=bytearray(); n=len(chunk); i=0
    # 3바이트 해시 -> 위치 목록
    idx={}
    def add(pos):
        if pos+3<=n: idx.setdefault(chunk[pos:pos+3],[]).append(pos)
    flag_pos=None; flag=0; cnt=0
    added=0
    while i<n:
        if cnt==0:
            flag_pos=len(out); out.append(0); flag=0
        # 현재 위치까지의 해시를 등록
        while added<i: add(added); added+=1
        best_len=0; best_off=0
        if i>=1:
            bit_count=max((i-1).bit_length(),4)
            max_len=(0xFFFF>>bit_count)+3
            for j in reversed(idx.get(chunk[i:i+3],[])):
                if i-j>(1<<bit_count): break
                l=0
                while i+l<n and l<max_len and chunk[j+l]==chunk[i+l]: l+=1
                if l>best_len: best_len=l; best_off=i-j
                if l>=max_len: break
        if best_len>=3:
            tok=((best_off-1)<<(16-bit_count))|(best_len-3)
            out+=struct.pack('<H',tok); flag|=(1<<cnt); i+=best_len
        else:
            out.append(chunk[i]); i+=1
        cnt+=1
        if cnt==8:
            out[flag_pos]=flag; cnt=0
    if cnt: out[flag_pos]=flag
    return bytes(out)

def compress(data):
    data=bytes(data); out=bytearray([1])
    for p in range(0,max(len(data),1),4096):
        chunk=data[p:p+4096]
        c=_chunk_compress(chunk)
        if len(c)>4096 or len(chunk)<1:
            raw=chunk+b'\x00'*(4096-len(chunk))
            out+=struct.pack('<H',(4096+2-3)|0x3000)+raw
        else:
            out+=struct.pack('<H',(len(c)+2-3)|0xB000)+c
    return bytes(out)

def decompress(buf):
    buf=bytes(buf); assert buf[0]==1
    out=bytearray(); p=1
    while p<len(buf):
        h=struct.unpack_from('<H',buf,p)[0]; size=(h&0x0FFF)+3; comp=h&0x8000
        end=p+size; p+=2; start=len(out)
        if not comp:
            out+=buf[p:p+4096]; p=end; continue
        while p<end:
            fl=buf[p]; p+=1
            for k in range(8):
                if p>=end: break
                if not (fl>>k)&1: out.append(buf[p]); p+=1
                else:
                    tok=struct.unpack_from('<H',buf,p)[0]; p+=2
                    diff=len(out)-start
                    bc=max((diff-1).bit_length(),4)
                    lm=0xFFFF>>bc
                    length=(tok&lm)+3; off=(tok>>(16-bc))+1
                    for _ in range(length): out.append(out[-off])
        p=end
    return bytes(out)

# ------------------------------------------------------------------ CFB 쓰기
def _dir_entry(name,etype,color,left,right,child,start,size):
    nm=name.encode('utf-16-le')
    assert len(nm)<=62
    e=nm+b'\x00'*(64-len(nm))
    e+=struct.pack('<H',len(nm)+2 if name else 0)
    e+=struct.pack('<BB',etype,color)
    e+=struct.pack('<III',left,right,child)
    e+=b'\x00'*16+struct.pack('<I',0)+b'\x00'*16
    e+=struct.pack('<I',start)+struct.pack('<Q',size)
    return e

NOSTREAM=0xFFFFFFFF
class Node:
    def __init__(self,name,data=None,children=None):
        self.name=name; self.data=data; self.children=children or []
        self.is_storage=data is None
        self.id=None; self.left=self.right=self.child=NOSTREAM; self.color=1; self.start=0; self.size=0

def _sortkey(n): return (len(n.name),n.name.upper())

def _balance(nodes,depth,maxdepth,perfect):
    """정렬된 형제 목록으로 균형 이진 트리를 만들고 루트를 돌려준다(말단 층은 적색)."""
    if not nodes: return None
    mid=len(nodes)//2
    root=nodes[mid]
    l=_balance(nodes[:mid],depth+1,maxdepth,perfect); r=_balance(nodes[mid+1:],depth+1,maxdepth,perfect)
    root.left=l.id if l else NOSTREAM; root.right=r.id if r else NOSTREAM
    root.color=0 if (not perfect and depth==maxdepth and depth>0) else 1
    return root

def write_cfb(root_children):
    root=Node('Root Entry',children=root_children)
    flat=[root]
    def assign(n):
        kids=sorted(n.children,key=_sortkey)
        for k in kids:
            k.id=len(flat); flat.append(k)
        for k in kids:
            if k.is_storage: assign(k)
        n._kids=kids
    root.id=0
    assign(root)
    def link(n):
        kids=n._kids
        if kids:
            m=len(kids); maxd=m.bit_length()-1; perfect=(m==(1<<(maxd+1))-1)
            r=_balance(kids,0,maxd,perfect)
            n.child=r.id
        for k in kids:
            if k.is_storage: link(k)
    link(root)
    # --- 섹터 배치
    SEC=512; MINI=64
    streams=[n for n in flat if not n.is_storage and n is not root]
    mini=[n for n in streams if len(n.data)<4096]
    big=[n for n in streams if len(n.data)>=4096]
    # 미니 스트림
    ministream=bytearray(); minifat=[]
    for n in mini:
        n.start=len(ministream)//MINI
        cnt=(len(n.data)+MINI-1)//MINI
        for k in range(cnt): minifat.append(n.start+k+1 if k<cnt-1 else 0xFFFFFFFE)
        ministream+=n.data+b'\x00'*(cnt*MINI-len(n.data))
        n.size=len(n.data)
    nminifat_sec=(len(minifat)*4+SEC-1)//SEC if minifat else 0
    nmini_sec=(len(ministream)+SEC-1)//SEC
    ndir_sec=(len(flat)*128+SEC-1)//SEC
    big_secs=[(n,(len(n.data)+SEC-1)//SEC) for n in big]
    nbig=sum(c for _,c in big_secs)
    nonfat=ndir_sec+nminifat_sec+nmini_sec+nbig
    nfat=1
    while nfat*128<nonfat+nfat: nfat+=1
    assert nfat<=109
    # 섹터 번호: FAT들 0.., dir, minifat, ministream, big
    pos=nfat
    dir_start=pos; pos+=ndir_sec
    minifat_start=pos if nminifat_sec else 0xFFFFFFFE; pos+=nminifat_sec
    mini_start=pos if nmini_sec else 0xFFFFFFFE; pos+=nmini_sec
    fat=[0xFFFFFFFF]*(nfat*128)
    for i in range(nfat): fat[i]=0xFFFFFFFD
    def chain(start,count):
        for k in range(count): fat[start+k]=start+k+1 if k<count-1 else 0xFFFFFFFE
    chain(dir_start,ndir_sec)
    if nminifat_sec: chain(minifat_start,nminifat_sec)
    if nmini_sec: chain(mini_start,nmini_sec)
    body={}
    for n,c in big_secs:
        n.start=pos; n.size=len(n.data); chain(pos,c); pos+=c
    root.start=mini_start if nmini_sec else 0xFFFFFFFE; root.size=len(ministream)
    for n in flat:
        if n.is_storage and n is not root: n.start=0; n.size=0
    # 디렉터리 바이트
    d=bytearray()
    for n in flat:
        etype=5 if n is root else (1 if n.is_storage else 2)
        if n is root: color=1
        else: color=n.color
        d+=_dir_entry(n.name,etype,color,n.left,n.right,n.child,n.start if (not n.is_storage or n is root) else 0,n.size)
    while len(d)<ndir_sec*SEC:
        d+=_dir_entry('',0,0,NOSTREAM,NOSTREAM,NOSTREAM,0,0)[:128]
    # 빈 엔트리는 모두 0 이어야 하지만 sibling 필드는 NOSTREAM
    # --- 헤더
    hdr=bytearray()
    hdr+=bytes.fromhex('D0CF11E0A1B11AE1')+b'\x00'*16
    hdr+=struct.pack('<HHHHH',0x3E,3,0xFFFE,9,6)+b'\x00'*6
    hdr+=struct.pack('<IIIIIIIII',0,nfat,dir_start,0,4096,minifat_start,nminifat_sec,0xFFFFFFFE,0)
    difat=list(range(nfat))+[0xFFFFFFFF]*(109-nfat)
    hdr+=struct.pack('<109I',*difat)
    assert len(hdr)==512
    out=bytearray(hdr)
    out+=struct.pack('<%dI'%len(fat),*fat)
    out+=d
    if nminifat_sec:
        mf=struct.pack('<%dI'%len(minifat),*minifat); mf+=b'\xff'*(nminifat_sec*SEC-len(mf)); out+=mf
    if nmini_sec: out+=bytes(ministream)+b'\x00'*(nmini_sec*SEC-len(ministream))
    for n,c in big_secs: out+=n.data+b'\x00'*(c*SEC-len(n.data))
    return bytes(out)

# ------------------------------------------------------------------ dir 스트림
def rec(rid,data): return struct.pack('<HI',rid,len(data))+data

def read_original_header(orig_dir):
    """원본 dir 에서 PROJECTMODULES(0x0F) 직전까지의 바이트를 돌려준다."""
    p=0
    while p<len(orig_dir):
        rid=struct.unpack_from('<H',orig_dir,p)[0]
        if rid==0x000F: return orig_dir[:p], orig_dir[p:]
        if rid==0x0009: p+=12; continue
        sz=struct.unpack_from('<I',orig_dir,p+2)[0]; p+=6+sz
    raise ValueError

def module_records(name,kind,private=False):
    mb=name.encode(CP); mu=name.encode('utf-16-le')
    b=rec(0x19,mb)+rec(0x47,mu)+rec(0x1A,mb)+rec(0x32,mu)+rec(0x1C,b'')+rec(0x48,b'')
    b+=rec(0x31,struct.pack('<I',0))+rec(0x1E,struct.pack('<I',0))+rec(0x2C,struct.pack('<H',0xFFFF))
    b+=rec(0x22 if kind in ('doc','form') else 0x21,b'')
    if private: b+=rec(0x28,b'')
    b+=rec(0x2B,b'')
    return b

def build(orig_bin,modules,out_path):
    """modules: [(name, kind('doc'|'std'|'form'), source_text)] 순서 = 프로젝트 순서"""
    o=olefile.OleFileIO(orig_bin)
    orig_dir=decompress(o.openstream('VBA/dir').read())
    head,tail=read_original_header(orig_dir)
    project_cookie=struct.unpack_from('<H',tail,6)[0] if False else None
    # tail = 0F record(8 bytes) + 13 record(8) + modules ... + 10
    cookie_rec=tail[8:16]
    d=head+rec(0x0F,struct.pack('<H',len(modules)))+cookie_rec
    for name,kind,src in modules:
        d+=module_records(name,kind,private=(kind=='form'))
    d+=rec(0x10,b'')
    vba_children=[Node('dir',compress(d)),Node('_VBA_PROJECT',bytes.fromhex('CC61FFFF000000'))]
    for name,kind,src in modules:
        data=src.replace('\r\n','\n').replace('\n','\r\n').encode(CP)
        vba_children.append(Node(name,compress(data)))
    # PROJECT 스트림
    orig_proj=o.openstream('PROJECT').read().decode(CP)
    keep={k:re.search(r'^%s=.*$'%k,orig_proj,re.M).group(0) for k in ('ID','CMG','DPB','GC','VersionCompatible32','HelpContextID')}
    lines=[keep['ID']]
    for name,kind,src in modules:
        if kind=='doc': lines.append(f'Document={name}/&H00000000')
        elif kind=='std': lines.append(f'Module={name}')
        elif kind=='form':
            lines.append('Package={AC9F2F90-E877-11CE-9F68-00AA00574A4F}'); lines.append(f'BaseClass={name}')
    lines+=['Name="VBAProject"',keep['HelpContextID'],keep['VersionCompatible32'],keep['CMG'],keep['DPB'],keep['GC'],'',
            '[Host Extender Info]','&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000','','[Workspace]']
    for name,kind,src in modules:
        lines.append(f'{name}=0, 0, 0, 0, C'+(', 32, 32, 1328, 691, C' if kind=='form' else ''))
    project='\r\n'.join(lines)+'\r\n'
    wm=b''
    for name,kind,src in modules: wm+=name.encode(CP)+b'\x00'+name.encode('utf-16-le')+b'\x00\x00'
    wm+=b'\x00\x00'
    root=[Node('PROJECT',project.encode(CP)),Node('PROJECTwm',wm),Node('VBA',children=vba_children)]
    for name,kind,src in modules:
        if kind=='form':
            kids=[Node(s,o.openstream(f'{name}/{s}').read()) for s in ('\x01CompObj','\x03VBFrame','f','o')]
            root.append(Node(name,children=kids))
    blob=write_cfb(root)
    open(out_path,'wb').write(blob)
    return blob
