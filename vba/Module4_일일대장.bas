Attribute VB_Name = "Module4_일일대장"
Option Explicit

' ============================================================
' KESCO 부천김포지사 검사업무 ERP
' Module4_일일대장 : 일일 접수대장 = 입력/작업 화면
'
'   흐름
'     1) 날짜를 고르면 10_원데이터의 그날 데이터가 화면에 나타납니다 (Day_Render).
'        접수가 없는 날짜는 빈 칸이므로 바로 입력하면 됩니다.
'     2) 화면에서 상호/용량/지역/시간/수수료/검사자를 직접 입력·수정합니다.
'        (왼쪽 = 정기검사, 오른쪽 = 사용전검사. 오른쪽 칸에 "한태희휴가"처럼 상호만 적으면 메모)
'     3) [자동배분]은 화면에 입력된 내용을 기준으로 현황판을 계산합니다 (저장하지 않음).
'     4) [입력 완료]를 누르면 화면 내용(+현황판에서 배정된 검사자)이 10_원데이터에 확정 저장됩니다 (Day_Commit).
'   저장하지 않고 날짜를 옮기면 입력한 내용이 사라집니다 (확인창이 뜹니다).
' ============================================================

Private sViewDone As Boolean
Private Const TITLE_TEXT As String = "입력 후 [입력 완료]를 눌러야 10_원데이터에 저장됩니다"

' ------------------------------------------------------------
' 화면 읽기
' ------------------------------------------------------------
Private Function VStr(ByVal x As Variant) As String
    If IsError(x) Then Exit Function
    If IsEmpty(x) Then Exit Function
    VStr = Trim$(CStr(x))
End Function

' A열의 병합 라벨(부천지역/김포지역/...)로 구역 구간을 찾음. 비고 행 번호를 noteRow 로 돌려줌
Private Function ScanSections(ws As Worksheet, secTop() As Long, secBot() As Long, secLab() As String, ByRef noteRow As Long) As Long
    Dim r As Long, n As Long, lab As String, ma As Range
    r = SCR_FIRST_ROW
    Do While r < 400
        lab = Trim$(CStr(ws.Cells(r, 1).Value))
        If lab = "" Or lab = "비고" Then Exit Do
        If n >= 6 Then Exit Do
        Set ma = ws.Cells(r, 1).MergeArea
        n = n + 1
        secTop(n) = r
        secBot(n) = r + ma.rows.Count - 1
        secLab(n) = lab
        r = secBot(n) + 1
    Loop
    noteRow = r
    ScanSections = n
End Function

' 화면의 모든 입력 줄(업무/메모)을 TaskRec 배열로 읽음. 반환 = 건수
Public Function ReadScreen(rec() As TaskRec) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_DAY)
    Dim secTop(1 To 6) As Long, secBot(1 To 6) As Long, secLab(1 To 6) As String
    Dim nSec As Long, noteRow As Long
    nSec = ScanSections(ws, secTop, secBot, secLab, noteRow)
    If nSec = 0 Then
        ReDim rec(1 To 1)
        ReadScreen = 0
        Exit Function
    End If
    Dim top1 As Long
    top1 = secTop(1)
    Dim v As Variant
    v = ws.Range(ws.Cells(top1, 1), ws.Cells(secBot(nSec), SCR_NOTE_R)).Value

    Dim stdMap As Object, zoneMap As Object, aliasMap As Object
    LoadRegionMaps stdMap, zoneMap
    Set aliasMap = LoadAliasMap()
    Dim slotMin As Long, bigKw As Double, bigMin As Long, smallMin As Long, onMin As Long
    slotMin = SettingLong("슬롯단위_분")
    bigKw = CDbl(SettingLong("정식설비기준용량_kW"))
    bigMin = SettingLong("정식설비소요시간_분")
    smallMin = SettingLong("약식설비소요시간_분")
    onMin = SettingLong("온라인업무소요시간_분")

    ReDim rec(1 To (secBot(nSec) - top1 + 1) * 2)
    Dim n As Long, s As Long, r As Long, ri As Long, side As Long, c0 As Long
    Dim nm As String, reg As String, insp As String, note As String, zone As String
    Dim capV As Variant, tmV As Variant, feeV As Variant
    For s = 1 To nSec
        zone = secLab(s)
        If zone = "기타지역" Then zone = "미분류"
        For r = secTop(s) To secBot(s) Step ENTRY_ROWS      ' 접수 1건 = 위아래 2행(병합) -> 윗 행이 대표
            ri = r - top1 + 1
            For side = 0 To 1
                If side = 0 Then c0 = 2 Else c0 = 8
                nm = VStr(v(ri, c0))
                capV = v(ri, c0 + 1)
                reg = VStr(v(ri, c0 + 2))
                tmV = v(ri, c0 + 3)
                If side = 0 Then
                    feeV = v(ri, 6)
                    insp = VStr(v(ri, 7))
                    note = VStr(v(ri, SCR_NOTE_L))
                Else
                    feeV = Empty
                    insp = VStr(v(ri, 12))
                    note = VStr(v(ri, SCR_NOTE_R))
                End If
                If nm <> "" Or VStr(capV) <> "" Or reg <> "" Or VStr(tmV) <> "" Or VStr(feeV) <> "" Or insp <> "" Then
                    n = n + 1
                    With rec(n)
                        .Key = r * 10 + side
                        .ScrRow = r
                        .Side = side
                        .Zone = zone
                        .TaskName = nm
                        .CapV = capV
                        .CapText = VStr(capV)
                        .Cap = CapNum(capV)
                        .RegionRaw = reg
                        .RegionStd = ""
                        If stdMap.Exists(reg) Then .RegionStd = stdMap(reg)
                        If Left$(.RegionStd, 1) = "(" Then .RegionStd = ""
                        .TimeV = tmV
                        .TimeStr = PromiseTime(tmV)
                        .FeeV = feeV
                        .Fee = VStr(feeV)
                        .InspRaw = insp
                        .Insp = LeadInsp(insp, aliasMap)
                        .InspFinal = NormInsp(insp, aliasMap)
                        .Note = note
                        .Online = (InStr(note, "온라인") > 0)
                        If nm <> "" And .CapText = "" And reg = "" And VStr(tmV) = "" And .Fee = "" And insp = "" Then
                            .Kind = KIND_MEMO            ' 상호만 있는 줄 = 메모(휴가/교육 등)
                        ElseIf side = 0 Then
                            .Kind = KIND_REG
                        Else
                            .Kind = KIND_PRE
                        End If
                        If .Online Then
                            .DurMin = onMin
                        ElseIf .Cap >= bigKw Then
                            .DurMin = bigMin
                        Else
                            .DurMin = smallMin
                        End If
                        .Slots = -Int(-.DurMin / slotMin)        ' 올림
                        If .Slots < 1 Then .Slots = 1
                    End With
                End If
            Next side
        Next r
    Next s
    If n = 0 Then
        ReDim rec(1 To 1)
    Else
        ReDim Preserve rec(1 To n)
    End If
    ReadScreen = n
End Function

' ------------------------------------------------------------
' 상태 표시 / 변경 감지
' ------------------------------------------------------------
Public Sub SetStatus()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_DAY)
    Dim prevEv As Boolean
    prevEv = Application.EnableEvents
    Application.EnableEvents = False
    With ws.Cells(2, 1)
        .Font.name = "맑은 고딕"
        .Font.Size = 10
        .Font.Bold = True
        .VerticalAlignment = xlCenter
        If gDirty Then
            .Value = "● 저장 안 됨 - [자동배분]으로 현황판을 갱신하고 [입력 완료]로 저장하세요"
            .Font.Color = RGB(192, 0, 0)
        Else
            .Value = TITLE_TEXT
            .Font.Color = RGB(127, 127, 127)
        End If
    End With
    Application.EnableEvents = prevEv
End Sub

' 시트 이벤트(Worksheet_Change)에서 호출: 접수대장 칸을 고치면 '저장 안 됨' 표시
Public Sub Day_OnChange(ByVal Target As Range)
    If gDirty Then Exit Sub
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_DAY)
    If Intersect(Target, ws.Range(ws.Cells(SCR_FIRST_ROW, 2), ws.Cells(400, SCR_NOTE_R))) Is Nothing Then Exit Sub
    gDirty = True
    SetStatus
End Sub

' 시트 이벤트(Worksheet_Activate)에서 호출: 날짜/원데이터가 바뀐 경우에만 다시 그림
Public Sub Day_Activate()
    If gDayDrawn Then
        If gDayDate = BoardDate() Then
            If gDirty Or gDayVer = gRawVer Then Exit Sub
        End If
    End If
    Day_Render
End Sub

Private Function ConfirmLeave() As Boolean
    If gDirty Then
        ConfirmLeave = Ask("저장하지 않은 입력이 있습니다." & vbCrLf & "[입력 완료]를 누르지 않고 이동하면 입력한 내용이 사라집니다. 이동할까요?")
    Else
        ConfirmLeave = True
    End If
End Function

' ============================================================
' 날짜 이동 / 되돌리기
' ============================================================
Sub Day_Prev()
    If Not ConfirmLeave() Then Exit Sub
    SetBoardDate BoardDate() - 1
    Day_Render
End Sub

Sub Day_Next()
    If Not ConfirmLeave() Then Exit Sub
    SetBoardDate BoardDate() + 1
    Day_Render
End Sub

Sub Day_Today()
    If Not ConfirmLeave() Then Exit Sub
    SetBoardDate Date
    Day_Render
End Sub

Sub Day_Pick()
    Dim d As Date
    If PickDate(BoardDate(), d) Then
        If Not ConfirmLeave() Then Exit Sub
        SetBoardDate d
        Day_Render
    End If
End Sub

' 저장하지 않은 입력을 버리고 10_원데이터의 내용으로 화면을 다시 그림
Sub Day_Revert()
    If gDirty Then
        If Not Ask("저장하지 않은 입력을 버리고 10_원데이터에 저장된 내용으로 되돌릴까요?") Then Exit Sub
    End If
    Day_Render
End Sub

' ------------------------------------------------------------
' 상단 조작 버튼 (인쇄 영역 밖: 1~3행) - 없을 때만 만듭니다
' ------------------------------------------------------------
Private Function ShapeThere(ws As Worksheet, ByVal nm As String) As Boolean
    Dim s As Shape
    On Error Resume Next
    Set s = ws.Shapes(nm)
    On Error GoTo 0
    ShapeThere = Not s Is Nothing
End Function

Public Sub Day_BuildChrome()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_DAY)
    Dim k As Long
    For k = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(k).name, 4) = "DBtn" Then ws.Shapes(k).Delete
    Next k
    Dim names As Variant, caps As Variant, macros As Variant, cols As Variant, i As Long
    names = Array("DBtnPrev", "DBtnToday", "DBtnNext", "DBtnPick", "DBtnRun", "DBtnCommit", "DBtnRevert", "DBtnPrint", "DBtnPdf")
    caps = Array(ChrW(9664) & " 이전일", "오늘", "다음일 " & ChrW(9654), ChrW(9638) & " 달력", "자동배분", "입력 완료", "되돌리기", "인쇄", "PDF 저장")
    macros = Array("Day_Prev", "Day_Today", "Day_Next", "Day_Pick", "RunAutoAssign", "Day_Commit", "Day_Revert", "Day_PrintForm", "Day_SavePdf")
    cols = Array(RGB(47, 117, 181), RGB(47, 117, 181), RGB(47, 117, 181), RGB(47, 117, 181), _
                 RGB(31, 56, 100), RGB(0, 112, 192), RGB(127, 127, 127), RGB(89, 89, 89), RGB(89, 89, 89))
    Dim L As Double, shp As Shape
    L = ws.Cells(1, 1).Left + 3
    For i = 0 To 8
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, L, ws.Cells(1, 1).Top + 3, 83, 24)
        shp.name = names(i)
        shp.Fill.ForeColor.RGB = cols(i)
        shp.Line.Visible = msoFalse
        shp.Placement = xlMove
        shp.OnAction = macros(i)
        With shp.TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 1: .MarginRight = 1
            .TextRange.Text = caps(i)
            .TextRange.Font.Size = 11
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
        L = L + 87.5
    Next i
End Sub

' ============================================================
' 화면 그리기: 10_원데이터 -> 접수대장(왼쪽) + 배분현황판(오른쪽)
' ============================================================
Public Sub Day_Render()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_DAY)
    Dim d As Date
    d = BoardDate()

    FastBegin
    On Error GoTo fail

    If Not ShapeThere(ws, "DBtnCommit") Then Day_BuildChrome

    Dim noteRow As Long
    noteRow = DrawLedger(ws, d)                 ' 왼쪽: 접수대장 (A~L열, 숨김 M~N열 = 행별 비고)
    gDirty = False
    gDayDate = d
    gDayVer = gRawVer
    gDayDrawn = True
    EnsureResultsForDate d                      ' 오늘 이후 날짜에 결과가 없으면 조용히 자동배분
    RebuildGrid                                 ' 오른쪽: 배분현황판 (O열~)
    SetupDayPrintIfNeeded ws, noteRow
    SetStatus

    If ActiveSheet.name <> SH_DAY Then ws.Activate
    ApplyDayView
fail:
    Dim em As String
    If Err.Number <> 0 Then em = Err.Description
    FastEnd
    If em <> "" Then Say "일일대장을 만드는 중 오류가 발생했습니다: " & em, vbCritical
End Sub

' 화면 고정: 1~6행(버튼/머리글)을 고정하고 처음 한 번만 90% 로 맞춤. 접수대장은 제자리, 현황판은 시간대 창만 스크롤([▲][▼])
Private Sub ApplyDayView()
    On Error Resume Next
    If ActiveSheet.name <> SH_DAY Then Exit Sub
    With ActiveWindow
        If Not (.FreezePanes And .SplitRow = GRID_HEADER_ROW And .SplitColumn = 0) Then
            .FreezePanes = False
            .ScrollRow = 1
            .ScrollColumn = 1
            .SplitColumn = 0
            .SplitRow = GRID_HEADER_ROW
            .FreezePanes = True
        End If
        If Not sViewDone Then
            .Zoom = 90
            sViewDone = True
        End If
    End With
End Sub

' 접수대장(A~N열)을 그리고 비고 행 번호를 돌려줌
Private Function DrawLedger(ws As Worksheet, ByVal d As Date) As Long
    With ws.Range("A4:N400")
        .UnMerge
        .Clear
        .Font.name = "맑은 고딕"
        .Font.Size = 10
        .VerticalAlignment = xlCenter
    End With

    Dim w As Variant, c As Long
    w = Array(5, 19, 8, 8.5, 7, 9.5, 8, 19, 8, 8.5, 7, 8)
    For c = 0 To 11
        ws.Columns(c + 1).ColumnWidth = w(c)
    Next c
    ' 숨김 속성은 환경에 따라 거부될 수 있어(1004) 실패하면 너비 0 으로 대신함
    On Error Resume Next
    ws.Columns(SCR_NOTE_L).EntireColumn.Hidden = True
    If Err.Number <> 0 Then Err.Clear: ws.Columns(SCR_NOTE_L).ColumnWidth = 0
    ws.Columns(SCR_NOTE_R).EntireColumn.Hidden = True
    If Err.Number <> 0 Then Err.Clear: ws.Columns(SCR_NOTE_R).ColumnWidth = 0
    On Error GoTo 0

    ' 1행 = 버튼 줄, 2행 = 상태 문구 + 현황판 날짜, 3행 숨김, 4행 = 날짜/범례, 5행 = 제목 막대, 6행 = 머리글
    ws.rows(1).RowHeight = 30
    ws.rows(2).RowHeight = 22
    On Error Resume Next
    ws.rows(3).EntireRow.Hidden = True
    If Err.Number <> 0 Then Err.Clear: ws.rows(3).RowHeight = 0.5
    On Error GoTo 0
    ws.rows(4).RowHeight = 20
    ws.rows(5).RowHeight = 24
    ws.rows(6).RowHeight = 24
    With ws.Range("H4:L4")
        .Merge
        .Value = Format(d, "m월 d일") & "  (" & WeekdayKo(d) & ")"
        .HorizontalAlignment = xlRight
        .Font.Bold = True
        .Font.Size = 13
    End With
    ws.Range("A5:G5").Merge
    ws.Range("A5").Value = "정기검사"
    ws.Range("H5:L5").Merge
    ws.Range("H5").Value = "사용전검사"
    ws.Range("A5:L5").HorizontalAlignment = xlCenter
    ws.Range("A5:L5").Font.Bold = True
    ws.Range("A5:L5").Font.Size = 14
    ws.Range("A5:L5").Font.Color = RGB(255, 255, 255)
    ws.Range("A5:G5").Interior.Color = RGB(31, 56, 100)
    ws.Range("H5:L5").Interior.Color = RGB(91, 143, 214)
    ws.Range("A6:L6").Value = Array("구분", "상호", "용량", "지역", "시간", "수수료", "검사자", "상호", "용량", "지역", "시간", "검사자")
    ws.Range("A6:L6").HorizontalAlignment = xlCenter
    ws.Range("A6:L6").Font.Bold = True
    ws.Range("A6:L6").Font.Size = 11

    Dim v As Variant, nTot As Long, key As Long
    v = RawData()
    If IsEmpty(v) Then nTot = 0 Else nTot = UBound(v, 1)
    key = CLng(Int(CDbl(d)))
    Dim stdMap As Object, zoneMap As Object
    LoadRegionMaps stdMap, zoneMap

    ' 그날의 행을 모아 구역(권역)별로 나눔
    Dim dayIdx() As Long, nDay As Long, i As Long
    ReDim dayIdx(1 To IIf(nTot < 1, 1, nTot))
    Dim zt As Variant
    Dim zoneOf() As String
    ReDim zoneOf(1 To IIf(nTot < 1, 1, nTot))
    Dim hasOther As Boolean, s As String
    For i = 1 To nTot
        If DayKey(v(i, RW_DATE)) = key Then
            nDay = nDay + 1
            dayIdx(nDay) = i
            s = Trim$(CStr(v(i, RW_ZONE)))
            If s = "" Then
                s = Trim$(CStr(v(i, RW_REG)))
                If s <> "" Then
                    If zoneMap.Exists(s) Then s = zoneMap(s) Else s = ""
                End If
            End If
            If s = "" Then s = "미분류"
            zoneOf(i) = s
            If s <> ZONE_A And s <> ZONE_B Then hasOther = True
        End If
    Next i

    Dim zones(1 To 3) As String, lastZone As Long
    zones(1) = ZONE_A
    zones(2) = ZONE_B
    zones(3) = "미분류"
    lastZone = 2
    If hasOther Then lastZone = 3

    Dim curRow As Long, z As Long, nR As Long, nP As Long, nRows As Long, k As Long, r As Long, nEnt As Long, e As Long
    Dim regIdx() As Long, preIdx() As Long
    Dim out() As Variant
    curRow = SCR_FIRST_ROW
    For z = 1 To lastZone
        ReDim regIdx(1 To IIf(nDay < 1, 1, nDay))
        ReDim preIdx(1 To IIf(nDay < 1, 1, nDay))
        nR = 0
        nP = 0
        For k = 1 To nDay
            i = dayIdx(k)
            If zoneOf(i) = zones(z) Or (z = 3 And zoneOf(i) <> ZONE_A And zoneOf(i) <> ZONE_B) Then
                If IsRegular(CStr(v(i, RW_KIND))) Then
                    nR = nR + 1
                    regIdx(nR) = i
                Else
                    nP = nP + 1
                    preIdx(nP) = i
                End If
            End If
        Next k
        nEnt = LEDGER_MIN_ENTRIES
        If nR + 1 > nEnt Then nEnt = nR + 1
        If nP + 1 > nEnt Then nEnt = nP + 1
        nRows = nEnt * ENTRY_ROWS

        ReDim out(1 To nRows, 1 To SCR_NOTE_R)
        For r = 1 To nR
            i = regIdx(r)
            e = (r - 1) * ENTRY_ROWS + 1               ' 건마다 윗 행에만 값을 넣음 (아래 행은 병합됨)
            out(e, 2) = v(i, RW_NAME)
            out(e, 3) = v(i, RW_CAP)
            out(e, 4) = v(i, RW_REG)
            out(e, 5) = v(i, RW_TIME)
            out(e, 6) = v(i, RW_FEE)
            out(e, 7) = v(i, RW_INSP)
            out(e, SCR_NOTE_L) = v(i, RW_NOTE)
        Next r
        For r = 1 To nP
            i = preIdx(r)
            e = (r - 1) * ENTRY_ROWS + 1
            out(e, 8) = v(i, RW_NAME)
            out(e, 9) = v(i, RW_CAP)
            out(e, 10) = v(i, RW_REG)
            out(e, 11) = v(i, RW_TIME)
            out(e, 12) = v(i, RW_INSP)
            out(e, SCR_NOTE_R) = v(i, RW_NOTE)
        Next r
        ws.Range(ws.Cells(curRow, 1), ws.Cells(curRow + nRows - 1, SCR_NOTE_R)).Value = out
        FormatDaySection ws, curRow, nRows, IIf(z = 3, "기타지역", zones(z))
        curRow = curRow + nRows
    Next z

    ' 검사자 칸: 명단에서 고르거나 직접 입력 (목록에 없으면 경고만)
    On Error Resume Next
    Dim vr As Variant
    For Each vr In Array(7, 12)
        With ws.Range(ws.Cells(SCR_FIRST_ROW, CLng(vr)), ws.Cells(curRow - 1, CLng(vr))).Validation
            .Delete
            .Add Type:=3, AlertStyle:=2, Operator:=1, Formula1:="=검사자명단"
            .IgnoreBlank = True
            .ShowError = False
        End With
    Next vr
    On Error GoTo 0

    ' 비고 행 (행별 비고 모음 - 화면 표시용)
    Dim notes As String
    For k = 1 To nDay
        i = dayIdx(k)
        If Trim$(CStr(v(i, RW_NOTE))) <> "" Then
            If notes <> "" Then notes = notes & " / "
            notes = notes & Trim$(CStr(v(i, RW_NAME))) & ": " & Trim$(CStr(v(i, RW_NOTE)))
        End If
    Next k
    ws.Cells(curRow, 1).Value = "비고"
    ws.Range(ws.Cells(curRow, 2), ws.Cells(curRow, 12)).Merge
    ws.Cells(curRow, 2).Value = notes
    ws.rows(curRow).RowHeight = 22
    With ws.Range(ws.Cells(curRow, 1), ws.Cells(curRow, 12))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    ws.Cells(curRow, 1).HorizontalAlignment = xlCenter
    ws.Cells(curRow, 1).Font.Bold = True
    ws.Cells(curRow, 2).Font.Size = 9
    ws.Cells(curRow, 2).WrapText = True
    DrawLedger = curRow
End Function

Private Sub FormatDaySection(ws As Worksheet, ByVal sRow As Long, ByVal nRows As Long, ByVal label As String)
    Dim eRow As Long
    eRow = sRow + nRows - 1
    Dim rg As Range
    Set rg = ws.Range(ws.Cells(sRow, 1), ws.Cells(eRow, 1))
    rg.Merge
    rg.Value = label
    rg.Orientation = xlVertical
    rg.HorizontalAlignment = xlCenter
    rg.VerticalAlignment = xlCenter
    rg.Font.Bold = True
    rg.Font.Size = 12

    ws.rows(sRow & ":" & eRow).RowHeight = BODY_ROW_H

    ' 본문 글자: 한 건이 위아래 2행이므로 크게 (12pt). 서식은 구역 전체에 한 번에 준 뒤 건마다 병합
    With ws.Range(ws.Cells(sRow, 2), ws.Cells(eRow, 12))
        .Font.Size = 12
        .VerticalAlignment = xlCenter
    End With
    ws.Range(ws.Cells(sRow, 8), ws.Cells(eRow, 12)).Font.Color = RGB(255, 0, 0)
    ws.Range(ws.Cells(sRow, 4), ws.Cells(eRow, 5)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(sRow, 10), ws.Cells(eRow, 11)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(sRow, 7), ws.Cells(eRow, 7)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(sRow, 12), ws.Cells(eRow, 12)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(sRow, 5), ws.Cells(eRow, 5)).NumberFormat = "h:mm"
    ws.Range(ws.Cells(sRow, 11), ws.Cells(eRow, 11)).NumberFormat = "h:mm"
    ws.Range(ws.Cells(sRow, 6), ws.Cells(eRow, 6)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(sRow, 6), ws.Cells(eRow, 6)).Font.Size = 10
    ws.Range(ws.Cells(sRow, 6), ws.Cells(eRow, 6)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(sRow, 3), ws.Cells(eRow, 3)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(sRow, 9), ws.Cells(eRow, 9)).HorizontalAlignment = xlRight
    ' 긴 상호/지역/검사자 이름은 줄이지 않고 칸에 맞춰 축소 표시
    ws.Range(ws.Cells(sRow, 2), ws.Cells(eRow, 2)).ShrinkToFit = True
    ws.Range(ws.Cells(sRow, 8), ws.Cells(eRow, 8)).ShrinkToFit = True
    ws.Range(ws.Cells(sRow, 4), ws.Cells(eRow, 4)).ShrinkToFit = True
    ws.Range(ws.Cells(sRow, 10), ws.Cells(eRow, 10)).ShrinkToFit = True
    ws.Range(ws.Cells(sRow, 7), ws.Cells(eRow, 7)).ShrinkToFit = True
    ws.Range(ws.Cells(sRow, 12), ws.Cells(eRow, 12)).ShrinkToFit = True

    ' 건마다 B~L열을 위아래 2행씩 병합
    Dim r As Long, c As Long
    For r = sRow To eRow Step ENTRY_ROWS
        For c = 2 To 12
            ws.Range(ws.Cells(r, c), ws.Cells(r + ENTRY_ROWS - 1, c)).Merge
        Next c
    Next r

    Set rg = ws.Range(ws.Cells(sRow, 1), ws.Cells(eRow, 12))
    rg.Borders.LineStyle = xlContinuous
    rg.Borders.Weight = xlThin
    rg.Borders(xlEdgeTop).Weight = xlMedium
    rg.Borders(xlEdgeBottom).Weight = xlMedium
    rg.Borders(xlEdgeLeft).Weight = xlMedium
    rg.Borders(xlEdgeRight).Weight = xlMedium
    ws.Range(ws.Cells(sRow, 8), ws.Cells(eRow, 8)).Borders(xlEdgeLeft).Weight = xlMedium
    ws.Range(ws.Cells(sRow, 2), ws.Cells(eRow, 2)).Borders(xlEdgeLeft).Weight = xlMedium
End Sub

' ============================================================
' 인쇄: A4 가로 1쪽 (접수대장 + 현황판 전체). 페이지 설정은 느리므로 인쇄 영역이 바뀔 때만 다시 설정
' ============================================================
Private Function DesiredPrintArea(ws As Worksheet, ByVal noteRow As Long) As String
    Dim lastRow As Long
    lastRow = noteRow
    If lastRow < GRID_START_ROW + BoardWinRows() Then lastRow = GRID_START_ROW + BoardWinRows()
    DesiredPrintArea = ws.Range(ws.Cells(4, 1), ws.Cells(lastRow, BoardLastCol())).Address
End Function

Private Sub SetupDayPrintIfNeeded(ws As Worksheet, ByVal noteRow As Long)
    Dim want As String, cur As String
    want = DesiredPrintArea(ws, noteRow)
    On Error Resume Next
    cur = ws.names("Print_Area").RefersToRange.Address
    On Error GoTo 0
    If cur <> want Then SetupDayPrint ws, want
End Sub

Private Sub SetupDayPrint(ws As Worksheet, ByVal area As String)
    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = area
        .PaperSize = xlPaperA4
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
        .LeftMargin = Application.CentimetersToPoints(0.8)
        .RightMargin = Application.CentimetersToPoints(0.8)
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(1)
        .HeaderMargin = Application.CentimetersToPoints(0.4)
        .FooterMargin = Application.CentimetersToPoints(0.4)
        .PrintGridlines = False
        .CenterFooter = "&8&P / &N"
    End With
    Application.PrintCommunication = True
    On Error GoTo 0
End Sub

Private Function CurrentNoteRow(ws As Worksheet) As Long
    Dim secTop(1 To 6) As Long, secBot(1 To 6) As Long, secLab(1 To 6) As String, noteRow As Long
    ScanSections ws, secTop, secBot, secLab, noteRow
    CurrentNoteRow = noteRow
End Function

Sub Day_PrintForm()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_DAY)
    SetupDayPrintIfNeeded ws, CurrentNoteRow(ws)
    ws.PrintOut
End Sub

Sub Day_SavePdf()
    Dim p As String, ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_DAY)
    SetupDayPrintIfNeeded ws, CurrentNoteRow(ws)
    p = ThisWorkbook.Path
    If p = "" Then p = Environ$("USERPROFILE") & "\Desktop"
    p = p & "\접수대장_" & Format(BoardDate(), "yyyymmdd") & ".pdf"
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=p, Quality:=xlQualityStandard, OpenAfterPublish:=False
    Say "PDF로 저장했습니다." & vbCrLf & p, vbInformation
End Sub

' ============================================================
' [입력 완료] : 화면 내용(+현황판 배정 검사자)을 10_원데이터에 확정 저장
' ============================================================
Private Function ReadResultMaps(ByVal d As Date, mInsp As Object, mManual As Object, mName As Object) As Boolean
    Set mInsp = CreateObject("Scripting.Dictionary")
    Set mManual = CreateObject("Scripting.Dictionary")
    Set mName = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_RESULT)
    Dim last As Long
    last = ws.Cells(ws.rows.Count, R_NAME).End(xlUp).Row
    If last < 2 Then Exit Function
    Dim v As Variant, r As Long, k As String
    v = ws.Range(ws.Cells(2, 1), ws.Cells(last, R_COLS)).Value
    For r = 1 To UBound(v, 1)
        If IsDate(v(r, R_DATE)) Then
            If DateValue(v(r, R_DATE)) = d Then
                If CLng(v(r, R_ID)) > 0 Then
                    k = CStr(CLng(v(r, R_ID)))
                    mName(k) = CStr(v(r, R_NAME))
                    If Trim$(CStr(v(r, R_INSP))) <> "" And CLng(v(r, R_SLOT)) > 0 Then
                        mInsp(k) = Trim$(CStr(v(r, R_INSP)))
                        mManual(k) = CStr(v(r, R_MANUAL))
                    End If
                    ReadResultMaps = True
                End If
            End If
        End If
    Next r
End Function

Private Function RowLabel(rc As TaskRec) As String
    RowLabel = rc.Zone & " " & rc.ScrRow & "행 " & IIf(rc.Side = 0, "(정기검사 칸)", "(사용전검사 칸)")
End Function

Sub Day_Commit()
    Dim d As Date
    d = BoardDate()
    Dim rec() As TaskRec, n As Long
    n = ReadScreen(rec)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_DAY)
    Dim i As Long, msg As String

    ' ---------- 1) 입력 점검 ----------
    Dim known As Object, nm As Variant, aliasMap As Object
    Set known = CreateObject("Scripting.Dictionary")
    For Each nm In RosterAllNames()
        known(CStr(nm)) = 1
    Next nm
    Set aliasMap = LoadAliasMap()
    Dim unknown As String, parts() As String, j As Long, p As String
    For i = 1 To n
        If rec(i).TaskName = "" Then
            msg = msg & vbCrLf & "· " & RowLabel(rec(i)) & ": 상호가 없습니다."
        End If
        If IsTaskKind(rec(i).Kind) And rec(i).InspFinal <> "" Then
            parts = Split(Replace(rec(i).InspFinal, ",", "-"), "-")
            For j = LBound(parts) To UBound(parts)
                p = NormInsp(CStr(parts(j)), aliasMap)
                If p <> "" Then
                    If Not known.Exists(p) Then unknown = unknown & vbCrLf & "· " & RowLabel(rec(i)) & ": " & p
                End If
            Next j
        End If
    Next i
    If msg <> "" Then
        If Not Ask("상호가 비어 있는 줄이 있습니다." & msg & vbCrLf & vbCrLf & "그대로 저장할까요? (아니오 = 돌아가서 고치기)") Then Exit Sub
    End If
    If unknown <> "" Then
        If Not Ask("01_검사자명단에 없는 검사자 이름이 있습니다." & unknown & vbCrLf & vbCrLf & "그대로 저장할까요?") Then Exit Sub
    End If

    ' ---------- 2) 기존 저장 건수 확인 ----------
    Dim oldCnt As Long, v As Variant, key As Long
    v = RawData()
    key = CLng(Int(CDbl(d)))
    If Not IsEmpty(v) Then
        For i = 1 To UBound(v, 1)
            If DayKey(v(i, RW_DATE)) = key Then oldCnt = oldCnt + 1
        Next i
    End If
    If n = 0 Then
        If oldCnt = 0 Then
            Say "저장할 내용이 없습니다.", vbInformation
            Exit Sub
        End If
        If Not Ask(Format(d, "yyyy-mm-dd") & " 의 저장된 데이터 " & oldCnt & "건이 모두 삭제됩니다. 계속할까요?") Then Exit Sub
    End If

    FastBegin
    On Error GoTo fail

    ' ---------- 3) 현황판에서 배정된 검사자를 빈 칸에 채움 ----------
    Dim mInsp As Object, mManual As Object, mName As Object, hasRes As Boolean
    hasRes = ReadResultMaps(d, mInsp, mManual, mName)
    Dim nFill As Long, nChg As Long, nBlank As Long, nSkip As Long, k As String
    Dim useRes As Boolean
    If hasRes Then useRes = Not ResultsStale(d)
    If hasRes And Not useRes Then nSkip = 1
    For i = 1 To n
        If IsTaskKind(rec(i).Kind) And rec(i).TaskName <> "" Then
            k = CStr(rec(i).Key)
            If useRes Then
                If mInsp.Exists(k) Then
                    If mName(k) = rec(i).TaskName Then
                        If rec(i).InspFinal = "" Then
                            rec(i).InspFinal = mInsp(k)
                            nFill = nFill + 1
                        ElseIf LeadInsp(rec(i).InspFinal, aliasMap) <> mInsp(k) And mManual(k) = "Y" Then
                            AddLog "수동조정", "검사자(" & rec(i).InspFinal & ")를 현황판 수동 조정 결과(" & mInsp(k) & ")로 바꿔 저장했습니다.", rec(i).TaskName
                            rec(i).InspFinal = mInsp(k)
                            nChg = nChg + 1
                        End If
                    End If
                End If
            End If
            If rec(i).InspFinal = "" Then nBlank = nBlank + 1
        End If
    Next i

    ' ---------- 4) 10_원데이터에 반영 (그 날짜 행을 통째로 교체) ----------
    CommitToRaw d, rec, n

    ' ---------- 5) 화면의 검사자 칸을 최종값으로 갱신 ----------
    For i = 1 To n
        If rec(i).InspFinal <> rec(i).InspRaw Then
            If rec(i).Side = 0 Then
                ws.Cells(rec(i).ScrRow, 7).Value = rec(i).InspFinal
            Else
                ws.Cells(rec(i).ScrRow, 12).Value = rec(i).InspFinal
            End If
        End If
    Next i
    gDirty = False
    gDayVer = gRawVer
    SetStatus
    FastEnd
    Board_Refresh

    Say Format(d, "yyyy-mm-dd") & " 입력 " & n & "건을 10_원데이터에 저장했습니다." & _
        IIf(nFill > 0, vbCrLf & "현황판 배정 검사자 " & nFill & "건을 채웠습니다.", "") & _
        IIf(nChg > 0, vbCrLf & "수동 조정으로 바뀐 검사자 " & nChg & "건", "") & _
        IIf(nBlank > 0, vbCrLf & "검사자가 비어 있는 업무 " & nBlank & "건 (자동배분 후 다시 [입력 완료])", "") & _
        IIf(nSkip > 0, vbCrLf & "※ 배분 이후 화면이 바뀌어 배분 결과는 반영하지 않았습니다. [자동배분] 후 다시 눌러 주세요.", ""), vbInformation
    Exit Sub
fail:
    Dim em As String
    em = Err.Description
    FastEnd
    Say "저장 중 오류가 발생했습니다: " & em, vbCritical
End Sub

' 그 날짜의 기존 행을 빼고, 화면 행을 날짜 순서에 맞는 위치에 넣어 표 전체를 한 번에 다시 씀
Private Sub CommitToRaw(ByVal d As Date, rec() As TaskRec, ByVal n As Long)
    Dim lo As ListObject, ws As Worksheet
    Set lo = RawTable()
    Set ws = lo.Parent
    RawInvalidate
    Dim old As Variant, oldN As Long
    old = RawData()
    If IsEmpty(old) Then oldN = 0 Else oldN = UBound(old, 1)

    Dim key As Long, i As Long, dk As Long, keep As Long, lastLE As Long, maxId As Double
    key = CLng(Int(CDbl(d)))
    For i = 1 To oldN
        If IsNumeric(old(i, RW_ID)) And Not IsEmpty(old(i, RW_ID)) Then
            If CDbl(old(i, RW_ID)) > maxId Then maxId = CDbl(old(i, RW_ID))
        End If
        dk = DayKey(old(i, RW_DATE))
        If dk = -1 And IsEmpty(old(i, RW_NAME)) And IsEmpty(old(i, RW_KIND)) Then
            ' 완전히 빈 줄은 버림
        ElseIf dk <> key Then
            keep = keep + 1
            If dk >= 0 And dk <= key Then lastLE = keep
        End If
    Next i

    Dim newN As Long
    newN = keep + n
    Dim arr() As Variant, j As Long, c As Long, k As Long, inserted As Boolean
    If newN > 0 Then ReDim arr(1 To newN, 1 To RW_COLS)
    j = 0
    If lastLE = 0 Then AppendNew arr, j, d, rec, n, maxId: inserted = True
    For i = 1 To oldN
        dk = DayKey(old(i, RW_DATE))
        If dk = -1 And IsEmpty(old(i, RW_NAME)) And IsEmpty(old(i, RW_KIND)) Then
            ' 빈 줄
        ElseIf dk <> key Then
            j = j + 1
            For c = 1 To RW_COLS
                arr(j, c) = old(i, c)
            Next c
            If Not inserted And j = lastLE Then
                AppendNew arr, j, d, rec, n, maxId
                inserted = True
            End If
        End If
    Next i

    ' 표 크기 조정 + 한 번에 쓰기
    Dim minRows As Long
    minRows = newN
    If minRows < 1 Then minRows = 1
    If minRows > oldN Then
        lo.Resize ws.Range(ws.Cells(1, 1), ws.Cells(minRows + 1, RW_COLS))
        ws.Range(ws.Cells(oldN + 2, RW_DATE), ws.Cells(minRows + 1, RW_DATE)).NumberFormat = "yyyy-mm-dd"
        ws.Range(ws.Cells(oldN + 2, RW_TIME), ws.Cells(minRows + 1, RW_TIME)).NumberFormat = "h:mm"
        ws.Range(ws.Cells(oldN + 2, RW_FEE), ws.Cells(minRows + 1, RW_FEE)).NumberFormat = "#,##0"
    End If
    If newN > 0 Then
        ws.Cells(2, 1).Resize(newN, RW_COLS).Value = arr
    End If
    If oldN > newN Then
        ws.Cells(newN + 2, 1).Resize(oldN - newN, RW_COLS).ClearContents
        If minRows < oldN Then lo.Resize ws.Range(ws.Cells(1, 1), ws.Cells(minRows + 1, RW_COLS))
    End If
    RawInvalidate
End Sub

' 화면 행 n건을 arr 의 j+1 행부터 채움 (관리번호는 최대값 다음부터 부여)
Private Sub AppendNew(arr() As Variant, ByRef j As Long, ByVal d As Date, rec() As TaskRec, ByVal n As Long, ByRef maxId As Double)
    Dim i As Long
    For i = 1 To n
        j = j + 1
        arr(j, RW_DATE) = d
        arr(j, RW_KIND) = rec(i).Kind
        If rec(i).Zone <> "" Then arr(j, RW_ZONE) = rec(i).Zone
        If rec(i).TaskName <> "" Then arr(j, RW_NAME) = rec(i).TaskName
        If rec(i).CapText <> "" Then arr(j, RW_CAP) = rec(i).CapV
        If rec(i).RegionRaw <> "" Then arr(j, RW_REG) = rec(i).RegionRaw
        If VStr(rec(i).TimeV) <> "" Then arr(j, RW_TIME) = rec(i).TimeV
        If rec(i).Fee <> "" Then arr(j, RW_FEE) = rec(i).FeeV
        If rec(i).InspFinal <> "" Then arr(j, RW_INSP) = rec(i).InspFinal
        If rec(i).Note <> "" Then arr(j, RW_NOTE) = rec(i).Note
        maxId = maxId + 1
        arr(j, RW_ID) = maxId
    Next i
End Sub

' ============================================================
' 10_원데이터를 직접 고쳤을 때(보조 경로): 캐시 무효화 + 관리번호/권역/검사자 별칭 자동 정리
' (일상 입력은 11_일일대장 + [입력 완료]를 쓰세요)
' ============================================================
Public Sub Raw_OnChange(ByVal Target As Range)
    RawInvalidate
    Dim lo As ListObject
    Set lo = RawTable()
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Dim hit As Range
    Set hit = Intersect(Target, lo.DataBodyRange)
    If hit Is Nothing Then Exit Sub
    If hit.CountLarge > 20000 Then Exit Sub
    Dim prevEv As Boolean
    prevEv = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo bail

    Dim aliasMap As Object, stdMap As Object, zoneMap As Object
    Set aliasMap = LoadAliasMap()
    LoadRegionMaps stdMap, zoneMap
    Dim mx As Double
    mx = Application.WorksheetFunction.Max(lo.ListColumns(RW_ID).DataBodyRange)

    Dim ar As Range, rw As Range, cel As Range, s As String, t As String
    For Each ar In Intersect(hit.EntireRow, lo.DataBodyRange).Areas
        For Each rw In ar.rows
            Set cel = rw.Cells(1, RW_INSP)
            s = Trim$(CStr(cel.Value))
            If s <> "" Then
                t = NormInsp(s, aliasMap)
                If t <> CStr(cel.Value) Then cel.Value = t
            End If
            If Trim$(CStr(rw.Cells(1, RW_ZONE).Value)) = "" Then
                s = Trim$(CStr(rw.Cells(1, RW_REG).Value))
                If s <> "" Then
                    If zoneMap.Exists(s) Then rw.Cells(1, RW_ZONE).Value = zoneMap(s)
                End If
            End If
            Set cel = rw.Cells(1, RW_ID)
            If Trim$(CStr(cel.Value)) = "" Then
                If Trim$(CStr(rw.Cells(1, RW_NAME).Value)) <> "" Or IsDate(rw.Cells(1, RW_DATE).Value) Then
                    mx = mx + 1
                    cel.Value = mx
                End If
            End If
        Next rw
    Next ar
bail:
    Application.EnableEvents = prevEv
End Sub
