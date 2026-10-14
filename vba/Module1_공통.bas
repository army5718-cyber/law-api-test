Attribute VB_Name = "Module1_공통"
Option Explicit

' ============================================================
' KESCO 부천김포지사 검사업무 ERP
' Module1_공통 : 시트명/컬럼 상수, 원데이터 읽기(캐시), 설정/명단/지역표 읽기, 로그
'
' 구조
'   11_일일대장 = 입력/작업 화면 (날짜별 접수대장 + 배분현황판)
'   10_원데이터 = 확정 저장소. 화면에서 [입력 완료]를 눌러야 저장됩니다.
'   배분(자동배분)은 10_원데이터가 아니라 "화면에 입력된 내용"을 기준으로 계산합니다.
' ============================================================

Public Const SH_ROSTER As String = "01_검사자명단"
Public Const SH_SETTINGS As String = "02_설정"
Public Const SH_MAP As String = "04_지역표"
Public Const SH_LOG As String = "07_배분로그"
Public Const SH_RESULT As String = "09_배분결과"
Public Const SH_RAW As String = "10_원데이터"
Public Const SH_DAY As String = "11_일일대장"
Public Const SH_MON As String = "12_월대장"
Public Const SH_STAT As String = "13_통계"

' 11_일일대장 배치
Public Const SCR_FIRST_ROW As Long = 7          ' 접수대장 본문 첫 행
Public Const SCR_NOTE_L As Long = 13            ' M열(숨김): 정기검사 행의 비고
Public Const SCR_NOTE_R As Long = 14            ' N열(숨김): 사용전검사 행의 비고
Public Const GRID_COL0 As Long = 15             ' O열 = 현황판 시간 라벨, P열부터 검사자, 마지막 = 미지정
Public Const GRID_HEADER_ROW As Long = 6
Public Const GRID_START_ROW As Long = 7
Public Const BODY_ROW_H As Long = 14            ' 한 칸(30분)의 행 높이
Public Const ENTRY_ROWS As Long = 2             ' 접수 1건 = 위아래 2행을 합친 칸 (글자를 크게 쓰기 위함)
Public Const LEDGER_MIN_ENTRIES As Long = 10    ' 구역(부천/김포)별 최소 입력 줄 수 (줄당 2행 -> 2구역 = 40행)
Public Const BOARD_WIN_ROWS As Long = 40        ' 현황판 시간대 창: 한 번에 보이는 칸 수 (접수대장 2구역과 같은 높이)

' 10_원데이터 컬럼
Public Const RW_DATE As Long = 1       ' 접수일자
Public Const RW_KIND As Long = 2       ' 구분 (정기검사 / 사용전검사 / 메모)
Public Const RW_ZONE As Long = 3       ' 권역
Public Const RW_NAME As Long = 4       ' 상호
Public Const RW_CAP As Long = 5        ' 용량
Public Const RW_REG As Long = 6        ' 지역
Public Const RW_TIME As Long = 7       ' 시간
Public Const RW_FEE As Long = 8        ' 수수료
Public Const RW_INSP As Long = 9       ' 검사자
Public Const RW_NOTE As Long = 10      ' 비고
Public Const RW_ID As Long = 11        ' 관리번호(자동)
Public Const RW_COLS As Long = 11

Public Const KIND_REG As String = "정기검사"
Public Const KIND_PRE As String = "사용전검사"
Public Const KIND_MEMO As String = "메모"
Public Const ZONE_A As String = "부천지역"
Public Const ZONE_B As String = "김포지역"

' 09_배분결과 컬럼 (1행 머리글, 2행부터 데이터)
Public Const R_DATE As Long = 1         ' 날짜
Public Const R_ID As Long = 2           ' 화면 행 키 (화면 행번호 x 10 + 0=정기/1=사용전, 점심=0)
Public Const R_NAME As Long = 3         ' 상호
Public Const R_CAP As Long = 4          ' 용량(kW)
Public Const R_REGION As Long = 5       ' 표준지역
Public Const R_DUR As Long = 6          ' 소요(분)
Public Const R_TIME As Long = 7         ' 약속시간
Public Const R_INSP As Long = 8         ' 검사자 (빈칸 = 미지정 열)
Public Const R_SLOT As Long = 9         ' 시작 슬롯 (0 = 미배치)
Public Const R_START As Long = 10       ' 시작 시각 (표시용)
Public Const R_NSLOT As Long = 11       ' 소요 칸 수
Public Const R_TYPE As Long = 12        ' 유형 1~3=순위, 4=미지정업무, 5=점심, 7=수동배정, 0=배치실패
Public Const R_MANUAL As Long = 13      ' 수동조정 여부 (Y)
Public Const R_FEE As Long = 14         ' 수수료(참고)
Public Const R_DCOL As Long = 15        ' 현황판 표시 열(참고)
Public Const R_DROW As Long = 16        ' 현황판 표시 행(참고)
Public Const R_COLS As Long = 16

' 화면(또는 원데이터) 한 줄 = 업무 1건
Public Type TaskRec
    Key As Long             ' 화면 행 키 (행번호 x 10 + 0/1)
    ScrRow As Long
    Side As Long            ' 0 = 정기검사(왼쪽), 1 = 사용전검사(오른쪽)
    Kind As String          ' 정기검사 / 사용전검사 / 메모
    Zone As String          ' 권역
    TaskName As String
    CapV As Variant
    CapText As String
    Cap As Double
    RegionRaw As String
    RegionStd As String
    TimeV As Variant
    TimeStr As String       ' "hh:mm" (약속시간 없으면 "")
    FeeV As Variant
    Fee As String
    InspRaw As String       ' 화면에 적힌 검사자
    Insp As String          ' 배분 기준 검사자 (공동검사는 첫 번째, 별칭 정정)
    InspFinal As String     ' 저장할 검사자 (별칭 정정 + 배분 결과 반영)
    Note As String
    Online As Boolean
    DurMin As Long
    Slots As Long
End Type

' ------------------------------------------------------------
' 화면 상태 (Day_Render / 입력 완료 / 월대장이 함께 사용)
' ------------------------------------------------------------
Public gDirty As Boolean            ' 화면을 고쳤고 아직 [입력 완료]를 누르지 않음
Public gRawVer As Long              ' 원데이터가 바뀔 때마다 증가
Public gDayDate As Date             ' 화면에 그려진 날짜
Public gDayVer As Long              ' 화면을 그릴 때의 gRawVer
Public gDayDrawn As Boolean
Public gBoardTop As Long            ' 현황판 시간대 창의 첫 칸(슬롯 번호, 1 = 표시시작시간). 0 = 아직 정하지 않음

Public gLogCount As Long
Public gSilent As Boolean
Public gLastMsg As String

Public Sub SetSilent(ByVal v As Boolean)
    gSilent = v
End Sub

Public Function LastMsg() As String
    LastMsg = gLastMsg
End Function

Public Sub Say(ByVal msg As String, Optional ByVal style As VbMsgBoxStyle = vbInformation)
    gLastMsg = msg
    If Not gSilent Then MsgBox msg, style
End Sub

Public Function Ask(ByVal msg As String) As Boolean
    gLastMsg = msg
    If gSilent Then
        Ask = True
    Else
        Ask = (MsgBox(msg, vbYesNo + vbQuestion) = vbYes)
    End If
End Function

' ------------------------------------------------------------
' 일일대장에서 선택된 날짜 (02_설정!C13, 이름정의 '보드날짜')
' ------------------------------------------------------------
Public Function BoardDate() As Date
    Dim v As Variant
    v = ThisWorkbook.names("보드날짜").RefersToRange.Value
    If IsDate(v) Then
        BoardDate = DateValue(v)
    Else
        BoardDate = Date
    End If
End Function

Public Sub SetBoardDate(ByVal d As Date)
    ThisWorkbook.names("보드날짜").RefersToRange.Value = DateValue(d)
End Sub

Public Function WeekdayKo(ByVal d As Date) As String
    WeekdayKo = Choose(Weekday(d), "일", "월", "화", "수", "목", "금", "토")
End Function

' ------------------------------------------------------------
' 02_설정 읽기
' ------------------------------------------------------------
Private Function SettingRaw(ByVal nm As String) As Variant
    SettingRaw = ThisWorkbook.names(nm).RefersToRange.Value
End Function

Public Function SettingText(ByVal nm As String) As String
    SettingText = Trim(CStr(SettingRaw(nm)))
End Function

Public Function SettingLong(ByVal nm As String) As Long
    SettingLong = CLng(SettingRaw(nm))
End Function

Public Function SettingTime(ByVal nm As String) As Date
    Dim s As String
    s = CellTimeStr(SettingRaw(nm))
    If s = "24:00" Then
        SettingTime = 1                  ' 하루의 끝 (24:00)
    Else
        SettingTime = TimeValue(s)
    End If
End Function

Private Function NameExists(ByVal nm As String) As Boolean
    Dim r As Range
    On Error Resume Next
    Set r = ThisWorkbook.names(nm).RefersToRange
    On Error GoTo 0
    NameExists = Not r Is Nothing
End Function

' 현황판 시간축(표시시간) = 02_설정 표시시작시간 ~ 표시종료시간 (기본 00:00 ~ 24:00)
' 업무 자동배치(미지정 업무/점심)는 운영시작시간 ~ 운영종료시간(근무시간) 안에서만 합니다.
Public Function DispStartTime() As Date
    If NameExists("표시시작시간") Then
        DispStartTime = SettingTime("표시시작시간")
    Else
        DispStartTime = SettingTime("운영시작시간")
    End If
End Function

Public Function DispEndTime() As Date
    If NameExists("표시종료시간") Then
        DispEndTime = SettingTime("표시종료시간")
    Else
        DispEndTime = SettingTime("운영종료시간")
    End If
End Function

Public Function DispSlotCount() As Long
    DispSlotCount = CalcSlotCount(DispStartTime(), DispEndTime(), SettingLong("슬롯단위_분"))
End Function

' 현황판 시간대 창 크기(칸). 시간축이 짧으면 시간축 전체(마감 행 포함)
Public Function BoardWinRows() As Long
    Dim n As Long
    n = DispSlotCount() + 1
    If n > BOARD_WIN_ROWS Then n = BOARD_WIN_ROWS
    BoardWinRows = n
End Function

' 창의 첫 칸을 허용 범위(1 ~ 시간축 끝이 창 맨 아래에 오는 위치)로 맞춰 돌려줌. 처음에는 근무시작 약 6시간 전부터
Public Function BoardTopSlot() As Long
    Dim mx As Long, t As Long
    mx = DispSlotCount() + 2 - BoardWinRows()
    If mx < 1 Then mx = 1
    t = gBoardTop
    If t < 1 Then t = WorkFirstSlot() - 12
    If t < 1 Then t = 1
    If t > mx Then t = mx
    gBoardTop = t
    BoardTopSlot = t
End Function

Public Function WorkFirstSlot() As Long
    Dim n As Long
    n = TimeToSlotIndex(CellTimeStr(SettingRaw("운영시작시간")), DispStartTime(), SettingLong("슬롯단위_분"))
    If n < 1 Then n = 1
    WorkFirstSlot = n
End Function

Public Function WorkLastSlot() As Long
    Dim n As Long
    n = WorkFirstSlot() + CalcSlotCount(SettingTime("운영시작시간"), SettingTime("운영종료시간"), SettingLong("슬롯단위_분")) - 1
    If n > DispSlotCount() Then n = DispSlotCount()
    WorkLastSlot = n
End Function

Public Function CalcSlotCount(ByVal startTime As Date, ByVal endTime As Date, ByVal slotMin As Long) As Long
    CalcSlotCount = CLng(Round((endTime - startTime) * 1440, 0)) \ slotMin
End Function

Public Function TimeLabel(ByVal t As Date) As String
    Dim m As Long
    m = CLng(Round(CDbl(t) * 1440, 0))
    If m >= 1440 Then
        TimeLabel = "24:00"
    Else
        TimeLabel = Format(m \ 60, "00") & ":" & Format(m Mod 60, "00")
    End If
End Function

' 셀 값을 "hh:mm" 문자열로 정규화 (시간값/문자열 모두 처리). 시간이 아니면 원문(또는 "")
Public Function CellTimeStr(ByVal v As Variant) As String
    If IsError(v) Then
        CellTimeStr = ""
    ElseIf IsEmpty(v) Then
        CellTimeStr = ""
    ElseIf VarType(v) = vbDate Then
        CellTimeStr = Format(v, "hh:nn")
    ElseIf VarType(v) = vbDouble Then
        If v > 0 And v < 1 Then
            CellTimeStr = Format(v, "hh:nn")
        Else
            CellTimeStr = Trim(CStr(v))
        End If
    Else
        CellTimeStr = Trim(CStr(v))
    End If
End Function

' 약속시간 칸 값 -> "hh:mm" (시간으로 읽을 수 없으면 "")
Public Function PromiseTime(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Or VarType(v) = vbDouble Then
        s = CellTimeStr(v)
    ElseIf VarType(v) = vbString Then
        If IsDate(v) Then s = Format(TimeValue(CStr(v)), "hh:nn")
    End If
    If s Like "##:##" Then PromiseTime = s
End Function

Public Function TimeToSlotIndex(ByVal tstr As String, ByVal startTime As Date, ByVal slotMin As Long) As Long
    On Error GoTo fail
    Dim tval As Date
    tval = TimeValue(tstr)
    Dim diffMin As Long
    diffMin = CLng(Round((tval - startTime) * 1440, 0))
    If diffMin < 0 Then GoTo fail
    TimeToSlotIndex = (diffMin \ slotMin) + 1
    Exit Function
fail:
    TimeToSlotIndex = 0
End Function

Public Function SlotIndexToTimeStr(ByVal slotIdx As Long, ByVal startTime As Date, ByVal slotMin As Long) As String
    Dim m As Long
    m = CLng(Round(CDbl(startTime) * 1440, 0)) + (slotIdx - 1) * slotMin
    If m >= 1440 Then
        SlotIndexToTimeStr = "24:00"
    Else
        SlotIndexToTimeStr = Format(m \ 60, "00") & ":" & Format(m Mod 60, "00")
    End If
End Function

' ------------------------------------------------------------
' 10_원데이터 (표 tblRaw) - 값 전체를 한 번만 읽어 캐시합니다.
' 원데이터가 바뀌면 RawInvalidate 를 불러야 합니다 (Worksheet_Change / 입력 완료가 호출).
' ------------------------------------------------------------
Private rawCache As Variant
Private rawCacheOK As Boolean

Public Function RawTable() As ListObject
    Set RawTable = ThisWorkbook.Sheets(SH_RAW).ListObjects("tblRaw")
End Function

Public Sub RawInvalidate()
    rawCacheOK = False
    gRawVer = gRawVer + 1
End Sub

Public Function RawData() As Variant
    If Not rawCacheOK Then
        Dim lo As ListObject
        Set lo = RawTable()
        If lo.DataBodyRange Is Nothing Then
            rawCache = Empty
        Else
            rawCache = lo.DataBodyRange.Value
        End If
        rawCacheOK = True
    End If
    RawData = rawCache
End Function

Public Function DayKey(ByVal v As Variant) As Long
    DayKey = -1
    If IsError(v) Then Exit Function
    If IsDate(v) Then DayKey = CLng(Int(CDbl(CDate(v))))
End Function

Public Function IsRegular(ByVal kind As String) As Boolean
    IsRegular = (Left$(Trim$(kind), 2) = "정기")
End Function

Public Function IsTaskKind(ByVal kind As String) As Boolean
    Dim k As String
    k = Left$(Trim$(kind), 2)
    IsTaskKind = (k = "정기" Or k = "사용")
End Function

' 용량 칸 값에서 숫자만 추출 (예: "800/236" -> 800, "전기차" -> 0)
Public Function CapNum(ByVal x As Variant) As Double
    Dim s As String, i As Long, t As String
    If IsError(x) Then Exit Function
    If IsEmpty(x) Then Exit Function
    If IsNumeric(x) Then
        CapNum = CDbl(x)
        Exit Function
    End If
    s = Trim$(CStr(x))
    For i = 1 To Len(s)
        If Mid$(s, i, 1) Like "[0-9.]" Then t = t & Mid$(s, i, 1) Else Exit For
    Next i
    If t <> "" Then CapNum = Val(t)
End Function

' ------------------------------------------------------------
' 검사자 이름 정규화 (01_검사자명단 H:I 별칭표 기준)
' ------------------------------------------------------------
Public Function LoadAliasMap() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim rg As Range
    On Error Resume Next
    Set rg = ThisWorkbook.names("검사자별칭").RefersToRange
    On Error GoTo 0
    If Not rg Is Nothing Then
        Dim v As Variant, r As Long, a As String, b As String
        v = rg.Value
        For r = 1 To UBound(v, 1)
            a = Trim$(CStr(v(r, 1)))
            b = Trim$(CStr(v(r, 2)))
            If a <> "" And b <> "" Then d(a) = b
        Next r
    End If
    Set LoadAliasMap = d
End Function

Public Function NormInsp(ByVal s As String, aliasMap As Object) As String
    Dim t As String
    t = Trim$(s)
    If t <> "" Then
        If aliasMap.Exists(t) Then t = aliasMap(t)
    End If
    NormInsp = t
End Function

' 공동검사("이덕기-김영권")이면 첫 번째 이름만 배분 기준으로 사용
Public Function LeadInsp(ByVal s As String, aliasMap As Object) As String
    Dim t As String, p As Long
    t = Trim$(s)
    p = InStr(t, "-")
    If p = 0 Then p = InStr(t, ",")
    If p > 1 Then t = Trim$(Left$(t, p - 1))
    LeadInsp = NormInsp(t, aliasMap)
End Function

' ------------------------------------------------------------
' 01_검사자명단
' ------------------------------------------------------------
Public Function RosterAllNames() As Collection
    Dim c As New Collection
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_ROSTER)
    Dim v As Variant, r As Long, nm As String
    v = ws.Range("A2:A30").Value
    For r = 1 To UBound(v, 1)
        nm = Trim(CStr(v(r, 1)))
        If nm <> "" And Left$(nm, 1) <> ChrW(8251) Then c.Add nm
    Next r
    Set RosterAllNames = c
End Function

Private Function HasAny(ByVal t As String, kws As Variant) As Boolean
    Dim i As Long
    For i = LBound(kws) To UBound(kws)
        If InStr(t, kws(i)) > 0 Then
            HasAny = True
            Exit Function
        End If
    Next i
End Function

' "한태희휴가" / "이종석 건강검진" 처럼 하루 종일 자리를 비우는 메모인지
Public Function IsFullDayLeave(ByVal t As String) As Boolean
    If HasAny(t, Array("휴가", "병가", "교육", "건강", "체련", "연가")) Then
        IsFullDayLeave = Not HasAny(t, Array("반차", "오후", "퇴근", "조퇴", "("))
    End If
End Function

' 메모 문구(근태/행사)인지
Public Function IsMemoText(ByVal t As String) As Boolean
    IsMemoText = HasAny(t, Array("휴가", "병가", "교육", "건강", "체련", "연가", "반차", "퇴근", "조퇴", "휴일", "출장"))
End Function

' 메모 행 목록에서 하루 종일 빠지는 검사자 이름 사전
Public Function LeaveFromRecs(rec() As TaskRec, ByVal n As Long) As Object
    Dim res As Object
    Set res = CreateObject("Scripting.Dictionary")
    Dim i As Long, nm As Variant, names As Collection
    Set names = RosterAllNames()
    For i = 1 To n
        If rec(i).Kind = KIND_MEMO Then
            If IsFullDayLeave(rec(i).TaskName) Then
                For Each nm In names
                    If InStr(rec(i).TaskName, CStr(nm)) > 0 Then res(CStr(nm)) = 1
                Next nm
            End If
        End If
    Next i
    Set LeaveFromRecs = res
End Function

' 오늘근무=TRUE 인 검사자 중 leave(휴가/교육 메모)에 없는 사람 (명단 순서 유지)
Public Function GetActiveInspectorsFor(leave As Object) As String()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_ROSTER)
    Dim v As Variant
    v = ws.Range("A2:C30").Value
    Dim result() As String
    Dim n As Long
    ReDim result(1 To UBound(v, 1))
    Dim r As Long, nm As String
    For r = 1 To UBound(v, 1)
        nm = Trim(CStr(v(r, 1)))
        If nm <> "" And Left$(nm, 1) <> ChrW(8251) Then
            If UCase(Trim(CStr(v(r, 3)))) = "TRUE" Then
                If Not leave.Exists(nm) Then
                    n = n + 1
                    result(n) = nm
                End If
            End If
        End If
    Next r
    If n = 0 Then
        ReDim GetActiveInspectorsFor(0)
        Exit Function
    End If
    ReDim Preserve result(1 To n)
    GetActiveInspectorsFor = result
End Function

Public Function HasInspectors(arr() As String) As Boolean
    On Error GoTo none
    HasInspectors = (UBound(arr) >= LBound(arr)) And (arr(LBound(arr)) <> "")
    Exit Function
none:
    HasInspectors = False
End Function

Public Function FindInspectorIndex(inspectors() As String, ByVal name As String) As Long
    Dim i As Long
    For i = LBound(inspectors) To UBound(inspectors)
        If inspectors(i) = name Then
            FindInspectorIndex = i
            Exit Function
        End If
    Next i
    FindInspectorIndex = 0
End Function

' ------------------------------------------------------------
' 04_지역표: 지역 -> 표준지역 / 권역, 표준지역 -> 클러스터 / 권역 (근접도 판정용)
' ------------------------------------------------------------
Public Sub LoadRegionMaps(stdMap As Object, zoneMap As Object)
    Set stdMap = CreateObject("Scripting.Dictionary")
    Set zoneMap = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_MAP)
    Dim last As Long
    last = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    If last < 2 Then Exit Sub
    Dim v As Variant, r As Long, a As String, b As String, c As String
    v = ws.Range(ws.Cells(2, 1), ws.Cells(last, 5)).Value
    For r = 1 To UBound(v, 1)
        a = Trim$(CStr(v(r, 1)))
        If a <> "" Then
            b = Trim$(CStr(v(r, 2)))
            c = Trim$(CStr(v(r, 3)))
            If b <> "" Then stdMap(a) = b
            If c <> "" Then zoneMap(a) = c
        End If
    Next r
End Sub

Public Sub LoadRegionTopology(clusterOfStd As Object, zoneOfStd As Object)
    Set clusterOfStd = CreateObject("Scripting.Dictionary")
    Set zoneOfStd = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_MAP)
    Dim last As Long
    last = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    If last < 2 Then Exit Sub
    Dim v As Variant, r As Long, s As String, c As String, z As String
    v = ws.Range(ws.Cells(2, 1), ws.Cells(last, 5)).Value
    For r = 1 To UBound(v, 1)
        s = Trim$(CStr(v(r, 2)))
        If s <> "" And Left$(s, 1) <> "(" Then
            c = Trim$(CStr(v(r, 4)))
            z = Trim$(CStr(v(r, 3)))
            If c <> "" And c <> "확인 필요" And Left$(c, 3) <> "데이터" Then
                If Not clusterOfStd.Exists(s) Then clusterOfStd(s) = c
            End If
            If z <> "" Then
                If Not zoneOfStd.Exists(s) Then zoneOfStd(s) = z
            End If
        End If
    Next r
End Sub

' ------------------------------------------------------------
' 07_배분로그에 한 줄 기록 (조용히 고치지 않고 항상 로그로 남김)
' ------------------------------------------------------------
Public Sub AddLog(ByVal logType As String, ByVal content As String, ByVal relatedTask As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_LOG)
    Dim r As Long
    r = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Value = Array(Now, logType, content, relatedTask, "확인필요")
    ws.Cells(r, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    gLogCount = gLogCount + 1
End Sub

' ------------------------------------------------------------
' 작업 시작/끝에서 화면 갱신·이벤트·자동계산을 잠시 끄고 되돌리는 도우미 (속도)
' ------------------------------------------------------------
Private sPrevSU As Boolean, sPrevEv As Boolean, sPrevCalc As Long, sDepth As Long

Public Sub FastBegin()
    If sDepth = 0 Then
        sPrevSU = Application.ScreenUpdating
        sPrevEv = Application.EnableEvents
        sPrevCalc = Application.Calculation
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
    End If
    sDepth = sDepth + 1
End Sub

Public Sub FastEnd()
    If sDepth > 0 Then sDepth = sDepth - 1
    If sDepth = 0 Then
        Application.Calculation = sPrevCalc
        Application.EnableEvents = sPrevEv
        Application.ScreenUpdating = sPrevSU
    End If
End Sub

' 오류로 중단되어 FastBegin 만 남았을 때 강제로 원상복구
Public Sub FastReset()
    sDepth = 0
    gBoardTop = 0
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
