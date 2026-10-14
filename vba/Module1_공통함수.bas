Attribute VB_Name = "Module1_공통함수"
Option Explicit

' ============================================================
' KESCO 부천김포지사 검사업무 ERP
' Module1_공통함수 : 시트명/컬럼 상수 + 공용 유틸리티
'
' 입력은 10_원데이터(tblRaw) 한 곳뿐입니다.
' 일일대장 / 월대장 / 통계 / 자동배분(현황판)이 모두 이 표를 직접 읽습니다.
' ============================================================

Public Const SH_ROSTER As String = "01_검사자명단"
Public Const SH_SETTINGS As String = "02_설정"
Public Const SH_MAP As String = "04_지역표"
Public Const SH_PRI As String = "05_우선순위매트릭스"
Public Const SH_LOG As String = "07_배분로그"
Public Const SH_RESULT As String = "09_배분결과"
Public Const SH_RAW As String = "10_원데이터"
Public Const SH_DAY As String = "11_일일대장"
Public Const SH_MON As String = "12_월대장"
Public Const SH_STAT As String = "13_통계"
Public Const SH_GRID As String = "11_일일대장"      ' 배분현황판은 일일대장 시트 오른쪽에 있습니다

' 현황판 위치 (11_일일대장): N열 = 시간 라벨, O열부터 검사자, 마지막 = 미지정
Public Const GRID_COL0 As Long = 14
Public Const GRID_HEADER_ROW As Long = 6
Public Const GRID_START_ROW As Long = 7
Public Const BODY_ROW_H As Long = 16          ' 접수대장 본문/현황판 한 칸(30분)의 행 높이
Public Const LEDGER_MIN_ROWS As Long = 24     ' 구역(부천/김포)별 접수대장 최소 행 수 (2구역 x 24 = 현황판 48칸과 같은 높이)

' 10_원데이터 컬럼
Public Const RW_DATE As Long = 1       ' 접수일자
Public Const RW_KIND As Long = 2       ' 구분 (정기검사 / 사용전검사 / 메모)
Public Const RW_ZONE As Long = 3       ' 권역 (비우면 지역으로 자동 판정)
Public Const RW_NAME As Long = 4       ' 상호
Public Const RW_CAP As Long = 5        ' 용량
Public Const RW_REG As Long = 6        ' 지역
Public Const RW_TIME As Long = 7       ' 시간
Public Const RW_FEE As Long = 8        ' 수수료
Public Const RW_INSP As Long = 9       ' 검사자
Public Const RW_NOTE As Long = 10      ' 비고
Public Const RW_ID As Long = 11        ' 관리번호(자동)
Public Const RW_CHECK As Long = 12     ' 점검(수식)

Public Const KIND_REG As String = "정기검사"
Public Const KIND_PRE As String = "사용전검사"
Public Const KIND_MEMO As String = "메모"
Public Const ZONE_A As String = "부천지역"
Public Const ZONE_B As String = "김포지역"

' 09_배분결과 컬럼 (1행 머리글, 2행부터 데이터)
Public Const R_DATE As Long = 1         ' 날짜
Public Const R_ID As Long = 2           ' 10_원데이터 관리번호 (점심=0)
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

' 배분 대상 업무 1건 (10_원데이터 1행을 계산용으로 정리한 것)
Public Type TaskRec
    ID As Long
    Kind As String
    TaskName As String
    CapText As String
    Cap As Double
    RegionRaw As String
    RegionStd As String
    TimeStr As String       ' "hh:mm" (약속시간 없으면 "")
    Insp As String          ' 명단 기준 정규 이름 ("" = 미지정)
    InspRaw As String       ' 원데이터에 적힌 값
    Online As Boolean
    Fee As String
    DurMin As Long
    Slots As Long
End Type

' 이번 실행에서 07_배분로그에 기록된 건수 (실행 요약용)
Public gLogCount As Long

' 안내창 대신 마지막 메시지만 저장 (자동 테스트용)
Public gSilent As Boolean
Public gLastMsg As String

' ------------------------------------------------------------
' 메시지 / 확인 (gSilent=True 이면 창을 띄우지 않음)
' ------------------------------------------------------------
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
' 02_설정의 이름정의 값 읽기 (활성 시트와 무관하게 동작)
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

' ------------------------------------------------------------
' 현황판 시간축(표시시간) = 02_설정 표시시작시간 ~ 표시종료시간 (기본 00:00 ~ 24:00, 24시간)
' 업무 자동배치(미지정 업무/점심)는 운영시작시간 ~ 운영종료시간(근무시간) 안에서만 합니다.
' 설정이 없으면 예전처럼 운영시간을 그대로 시간축으로 씁니다.
' ------------------------------------------------------------
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

' 시간축 전체 칸 수 (24시간 / 30분 = 48)
Public Function DispSlotCount() As Long
    DispSlotCount = CalcSlotCount(DispStartTime(), DispEndTime(), SettingLong("슬롯단위_분"))
End Function

' 근무시간(운영시작~운영종료)의 첫 칸 / 마지막 칸 번호 (시간축 기준 1-based)
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

' 시각 -> "hh:mm" (하루의 끝은 "24:00")
Public Function TimeLabel(ByVal t As Date) As String
    Dim m As Long
    m = CLng(Round(CDbl(t) * 1440, 0))
    If m >= 1440 Then
        TimeLabel = "24:00"
    Else
        TimeLabel = Format(m \ 60, "00") & ":" & Format(m Mod 60, "00")
    End If
End Function

' ------------------------------------------------------------
' 셀 값을 "hh:mm" 문자열로 정규화
' (셀에 09:30 을 시간값으로 입력했든 문자열로 입력했든 동일하게 처리)
' ------------------------------------------------------------
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

' 시간 범위 내 슬롯 개수 (24시간 / 30분 = 48칸, 근무시간 09:30~18:00 = 17칸)
Public Function CalcSlotCount(ByVal startTime As Date, ByVal endTime As Date, ByVal slotMin As Long) As Long
    CalcSlotCount = CLng(Round((endTime - startTime) * 1440, 0)) \ slotMin
End Function

' ------------------------------------------------------------
' 10_원데이터 표 접근
' ------------------------------------------------------------
Public Function RawTable() As ListObject
    Set RawTable = ThisWorkbook.Sheets(SH_RAW).ListObjects("tblRaw")
End Function

' 표 본문 값을 2차원 배열로 (비어 있으면 Empty)
Public Function RawData() As Variant
    Dim lo As ListObject
    Set lo = RawTable()
    If lo.DataBodyRange Is Nothing Then
        RawData = Empty
    Else
        RawData = lo.DataBodyRange.Value
    End If
End Function

' 날짜 값 -> 일련번호(Long). 날짜가 아니면 -1
Public Function DayKey(ByVal v As Variant) As Long
    DayKey = -1
    If IsError(v) Then Exit Function
    If IsDate(v) Then DayKey = CLng(Int(CDbl(CDate(v))))
End Function

Public Function IsRegular(ByVal kind As String) As Boolean
    IsRegular = (Left$(Trim$(kind), 2) = "정기")
End Function

' 정기검사/사용전검사 행인지 (메모 행 제외)
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

' 원데이터 관리번호 -> 표 행 번호(1-based) 사전
Public Function RawIdIndex() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim lo As ListObject
    Set lo = RawTable()
    If Not lo.DataBodyRange Is Nothing Then
        Dim v As Variant, i As Long
        v = lo.ListColumns(RW_ID).DataBodyRange.Value
        For i = 1 To UBound(v, 1)
            If Trim$(CStr(v(i, 1))) <> "" Then d(CStr(v(i, 1))) = i
        Next i
    End If
    Set RawIdIndex = d
End Function

' ------------------------------------------------------------
' 검사자 이름 정규화 (01_검사자명단 H:I 별칭표 기준)
' ------------------------------------------------------------
Public Function LoadAliasMap() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim rg As Range, r As Long, a As String, b As String
    On Error Resume Next
    Set rg = ThisWorkbook.names("검사자별칭").RefersToRange
    On Error GoTo 0
    If Not rg Is Nothing Then
        For r = 1 To rg.rows.Count
            a = Trim$(CStr(rg.Cells(r, 1).Value))
            b = Trim$(CStr(rg.Cells(r, 2).Value))
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
' 01_검사자명단 읽기
'   오늘근무=TRUE 인 검사자만, 해당 날짜에 휴가/교육 메모가 있는 사람은 제외
' ------------------------------------------------------------
Public Function RosterAllNames() As Collection
    Dim c As New Collection
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_ROSTER)
    Dim lastRow As Long, r As Long, nm As String
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        nm = Trim(CStr(ws.Cells(r, 1).Value))
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

' 해당 날짜에 휴가/교육 등으로 빠지는 검사자 이름 사전
Public Function LeaveSet(ByVal d As Date) As Object
    Dim res As Object
    Set res = CreateObject("Scripting.Dictionary")
    Dim v As Variant
    v = RawData()
    If Not IsEmpty(v) Then
        Dim names As Collection, nm As Variant, i As Long, t As String, key As Long
        Set names = RosterAllNames()
        key = CLng(Int(CDbl(d)))
        For i = 1 To UBound(v, 1)
            If DayKey(v(i, RW_DATE)) = key Then
                If Trim$(CStr(v(i, RW_KIND))) = KIND_MEMO Then
                    t = Trim$(CStr(v(i, RW_NAME)))
                    If t <> "" Then
                        If IsFullDayLeave(t) Then
                            For Each nm In names
                                If InStr(t, CStr(nm)) > 0 Then res(CStr(nm)) = 1
                            Next nm
                        End If
                    End If
                End If
            End If
        Next i
    End If
    Set LeaveSet = res
End Function

Public Function GetActiveInspectorsFor(ByVal d As Date) As String()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_ROSTER)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    Dim leave As Object
    Set leave = LeaveSet(d)

    Dim result() As String
    Dim n As Long: n = 0
    ReDim result(1 To lastRow)

    Dim r As Long, nm As String
    For r = 2 To lastRow
        nm = Trim(CStr(ws.Cells(r, 1).Value))
        If nm <> "" And Left$(nm, 1) <> ChrW(8251) Then
            If UCase(Trim(CStr(ws.Cells(r, 3).Value))) = "TRUE" Then
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

' 검사자 배열에 실제 이름이 1명 이상 들어있는지
Public Function HasInspectors(arr() As String) As Boolean
    On Error GoTo none
    HasInspectors = (UBound(arr) >= LBound(arr)) And (arr(LBound(arr)) <> "")
    Exit Function
none:
    HasInspectors = False
End Function

' ------------------------------------------------------------
' 검사자 이름 배열에서 특정 이름의 인덱스(1-based)를 찾음. 없으면 0.
' ------------------------------------------------------------
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
' 04_지역표: 지역 -> 표준지역 / 권역 사전
' ------------------------------------------------------------
Public Sub LoadRegionMaps(stdMap As Object, zoneMap As Object)
    Set stdMap = CreateObject("Scripting.Dictionary")
    Set zoneMap = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_MAP)
    Dim last As Long, r As Long, a As String, b As String, c As String
    last = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    For r = 2 To last
        a = Trim$(CStr(ws.Cells(r, 1).Value))
        If a <> "" Then
            b = Trim$(CStr(ws.Cells(r, 2).Value))
            c = Trim$(CStr(ws.Cells(r, 3).Value))
            If b <> "" Then stdMap(a) = b
            If c <> "" Then zoneMap(a) = c
        End If
    Next r
End Sub

' ------------------------------------------------------------
' "HH:MM" 문자열을 운영시작시간 기준 슬롯 인덱스(1-based)로 변환
' 운영시간 이전이거나 해석 불가이면 0
' ------------------------------------------------------------
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

' ------------------------------------------------------------
' 슬롯 인덱스를 "HH:MM" 문자열로 변환 (로그/디버깅용)
' ------------------------------------------------------------
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
' 07_배분로그에 한 줄 기록 (자동 침묵 수정 금지 원칙 - 항상 로그로 남김)
' ------------------------------------------------------------
Public Sub AddLog(ByVal logType As String, ByVal content As String, ByVal relatedTask As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_LOG)
    Dim r As Long
    r = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Cells(r, 2).Value = logType
    ws.Cells(r, 3).Value = content
    ws.Cells(r, 4).Value = relatedTask
    ws.Cells(r, 5).Value = "확인필요"
    gLogCount = gLogCount + 1
End Sub

' ------------------------------------------------------------
' 05_우선순위매트릭스를 Dictionary(표준지역 -> 순위배열)로 로드
' ------------------------------------------------------------
Public Sub BuildPriorityMap(ByRef priMap As Object)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_PRI)
    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim r As Long, c As Long, n As Long
    Dim baseRegion As String, v As String
    Dim arr() As String
    For r = 2 To lastRow
        baseRegion = Trim(CStr(ws.Cells(r, 1).Value))
        If baseRegion <> "" Then
            n = 0
            ReDim arr(1 To lastCol)
            For c = 2 To lastCol
                v = Trim(CStr(ws.Cells(r, c).Value))
                If v <> "" Then
                    n = n + 1
                    arr(n) = v
                End If
            Next c
            If n > 0 Then
                ReDim Preserve arr(1 To n)
                priMap(baseRegion) = arr
            End If
        End If
    Next r
End Sub

' ------------------------------------------------------------
' 두 표준지역 간 "거리"(순위 매트릭스 상 위치, 낮을수록 가까움)
' 같은 지역이면 0, 정보 없음 500, 매트릭스에 없으면 999(먼 것으로 간주)
' ------------------------------------------------------------
Public Function RegionDistance(ByVal priMap As Object, ByVal fromRegion As String, ByVal toRegion As String) As Long
    If fromRegion = "" Or toRegion = "" Then
        RegionDistance = 500
        Exit Function
    End If
    If fromRegion = toRegion Then
        RegionDistance = 0
        Exit Function
    End If
    If Not priMap.Exists(fromRegion) Then
        RegionDistance = 999
        Exit Function
    End If
    Dim arr() As String
    arr = priMap(fromRegion)
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = toRegion Then
            RegionDistance = i
            Exit Function
        End If
    Next i
    RegionDistance = 999
End Function
