"""엑셀 없이 할 수 있는 전체 점검 모음.
   사용: python tests/run_all.py KESCO_검사ERP_2026.xlsm vba
   1) vbacheck  : 문법 구조(If/For/With 짝, 괄호, 문자열, CP949 변환)
   2) vbalint   : Option Explicit 위반(선언 안 한 변수) 탐지
   3) test_byref: ByRef 인자 형식 불일치 탐지
   4) test_contract: 시트/표/열/이름정의/매크로/문서모듈/호출 인자 수 일치
   5) test_port : 핵심 로직(화면<->원데이터, 자동배분, 입력 완료)을 파이썬 포트로 실제 데이터에 실행"""
import subprocess, sys, os
here=os.path.dirname(os.path.abspath(__file__)); x,v=sys.argv[1:3]
steps=[('vbacheck',[sys.executable,f'{here}/vbacheck.py',v]),('vbalint',[sys.executable,f'{here}/vbalint.py',v]),
       ('byref',[sys.executable,f'{here}/test_byref.py',v]),('contract',[sys.executable,f'{here}/test_contract.py',x,v]),
       ('port',[sys.executable,f'{here}/test_port.py',x]),('layout',[sys.executable,f'{here}/test_layout.py',x])]
bad=0
for n,cmd in steps:
    p=subprocess.run(cmd,capture_output=True,text=True); out=(p.stdout+p.stderr).strip()
    print(f'=== {n} (exit {p.returncode})'); print(out[-1500:])
    if p.returncode or 'FAIL' in out and 'FAILS 0' not in out and 'FAIL 0' not in out: bad+=1
print('RESULT','PASS' if not bad else 'CHECK')
