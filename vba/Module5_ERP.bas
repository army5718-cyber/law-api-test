Attribute VB_Name = "Module5_ERP"
Option Explicit

' ============================================================
' KESCO 부천김포지사 검사업무 ERP
' Module5_ERP : 원데이터(10_원데이터) -> 일일대장(+현황판) / 월대장 자동 생성, 배분 결과 반영
'
'   Day_Render      : 선택일의 접수대장(왼쪽)과 배분현황판(오른쪽)을 11_일일대장에 그림
'   Month_Render    : 선택월의 월 접수대장(요일 x 주차 배치)을 12_월대장에 그림
'   Erp_ApplyAssignment : 배분 결과의 검사자를 원데이터에 기록
'   Erp_OnRawChange : 원데이터에 입력할 때 관리번호/권역/검사자 이름을 자동 정리
' 입력은 10_원데이터 표(tblRaw) 한 곳뿐이며, 대장과 배분은 모두 여기서 만들어집니다.
' ============================================================

' 대장에 표시할 구분 텍스트 / 구역
Private Function ZoneText(v As Variant, ByVal i As Long, zoneMap As Object) As String
    Dim s As String
    s = Trim$(CStr(v(i, RW_ZONE)))
    If s = "" Then
        Dim reg As String
        reg = Trim$(CStr(v(i, RW_REG)))
        If reg <> "" Then
            If zoneMap.Exists(reg) Then s = zoneMap(reg)
        End If
    End If
    If s = "" Then s = "미분류"
    ZoneText = s
End Function

' 시간 칸 값: 시간값이면 그대로, 문자면 문자 그대로
Private Function TimeCell(ByVal x As Variant) As Variant
    If IsError(x) Then
        TimeCell = ""
    ElseIf IsEmpty(x) Then
        TimeCell = ""
    Else
        TimeCell = x
    End If
End Function

' 선택일의 배분 결과 중 검사자가 정해진 업무: 원데이터ID -> 검사자
Private Function ResultInspMap(ByVal d As Date) As Object
    Dim m As Object
    Set m = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_RESULT)
    Dim last As Long, r As Long, tid As Long
    last = ws.Cells(ws.rows.Count, R_NAME).End(xlUp).Row
    For r = 2 To last
        If IsDate(ws.Cells(r, R_DATE).Value) Then
            If DateValue(ws.Cells(r, R_DATE).Value) = d Then
                tid = CLng(ws.Cells(r, R_ID).Value)
                If tid > 0 And Trim$(CStr(ws.Cells(r, R_INSP).Value)) <> "" And CLng(ws.Cells(r, R_SLOT).Value) > 0 Then
                    m(CStr(tid)) = Trim$(CStr(ws.Cells(r, R_INSP).Value))
                End If
            End If
        End If
    Next r
    Set ResultInspMap = m
End Function

' ============================================================
' 1. 일일 접수대장 + 배분현황판 (실무 출력폼)
' ============================================================
Sub Day_Prev()
    SetBoardDate BoardDate() - 1
    Day_Render
End Sub

Sub Day_Next()
    SetBoardDate BoardDate() + 1
    Day_Render
End Sub

Sub Day_Today()
    SetBoardDate Date
    Day_Render
End Sub

Sub Day_Pick()
    Dim d As Date
    If PickDate(BoardDate(), d) Then
        SetBoardDate d
        Day_Render
    End If
End Sub

Sub Day_PrintForm()
    Day_Render
    ThisWorkbook.Sheets(SH_DAY).PrintOut
End Sub

Sub Day_SavePdf()
    Dim p As String, ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_DAY)
    Day_Render
    p = ThisWorkbook.Path
    If p = "" Then p = Environ$("USERPROFILE") & "\Desktop"
    p = p & "\접수대장_" & Format(BoardDate(), "yyyymmdd") & ".pdf"
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=p, Quality:=xlQualityStandard, OpenAfterPublish:=False
    Say "PDF로 저장했습니다." & vbCrLf & p, vbInformation
End Sub

' 상단 조작 버튼 (인쇄 영역 밖: 1~3행)
Public Sub Day_BuildChrome()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_DAY)
    Dim k As Long
    For k = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(k).name, 4) = "DBtn" Then ws.Shapes(k).Delete
    Next k
    ws.rows(1).RowHeight = 24
    ws.rows(2).RowHeight = 26
    ws.rows(3).RowHeight = 6
    ws.Cells(1, 1).Value = "일일 접수대장 + 배분현황판 (10_원데이터 기준 자동 생성 / 인쇄: 4행부터, A4 가로)"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 12
    ws.Cells(1, 1).Font.Color = RGB(31, 56, 100)
    Dim names As Variant, caps As Variant, macros As Variant, cols As Variant, i As Long
    names = Array("DBtnPrev", "DBtnToday", "DBtnNext", "DBtnPick", "DBtnRun", "DBtnApply", "DBtnPrint", "DBtnPdf")
    caps = Array(ChrW(9664) & " 이전일", "오늘", "다음일 " & ChrW(9654), ChrW(9638) & " 달력", "자동배분", "결과 반영", "인쇄", "PDF 저장")
    macros = Array("Day_Prev", "Day_Today", "Day_Next", "Day_Pick", "RunAutoAssign", "Erp_ApplyAssignment", "Day_PrintForm", "Day_SavePdf")
    cols = Array(RGB(47, 117, 181), RGB(47, 117, 181), RGB(47, 117, 181), RGB(47, 117, 181), _
                 RGB(31, 56, 100), RGB(0, 112, 192), RGB(89, 89, 89), RGB(89, 89, 89))
    Dim L As Double, shp As Shape
    L = ws.Cells(2, 1).Left + 4
    For i = 0 To 7
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, L, ws.Cells(2, 1).Top + 2, 72, 22)
        shp.name = names(i)
        shp.Fill.ForeColor.RGB = cols(i)
        shp.Line.Visible = msoFalse
        shp.Placement = xlMove
        shp.OnAction = macros(i)
        With shp.TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 2: .MarginRight = 2
            .TextRange.Text = caps(i)
            .TextRange.Font.Size = 10
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
        L = L + 76
    Next i
End Sub

Public Sub Day_Render()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_DAY)
    Dim d As Date
    d = BoardDate()

    Dim prevSU As Boolean, prevEv As Boolean
    prevSU = Application.ScreenUpdating
    prevEv = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo fail

    ' 오늘 이후 날짜인데 배분 결과가 없으면 조용히 자동배분 (원데이터 -> 배분 연계)
    EnsureResultsForDate d

    Dim ledgerLast As Long
    ledgerLast = DrawLedger(ws, d)          ' 왼쪽: 접수대장 (A~L열)
    RebuildGrid True                        ' 오른쪽: 배분현황판 (N열~)
    SetupDayPrint ws, ledgerLast

    ws.Activate
    ws.Range("A1").Select
fail:
    Dim em As String
    If Err.Number <> 0 Then em = Err.Description
    Application.EnableEvents = prevEv
    Application.ScreenUpdating = prevSU
    If em <> "" Then Say "일일대장을 만드는 중 오류가 발생했습니다: " & em, vbCritical
End Sub

' 접수대장(A~L열)을 그리고 마지막 행 번호(비고 행)를 돌려줌
Private Function DrawLedger(ws As Worksheet, ByVal d As Date) As Long
    With ws.Range("A4:L400")
        .UnMerge
        .Clear
        .Font.name = "맑은 고딕"
        .Font.Size = 10
        .VerticalAlignment = xlCenter
    End With

    ' 열 너비 (왼쪽 접수대장 + 오른쪽 현황판이 A4 가로 1쪽에 들어가도록 좁게)
    Dim w As Variant, c As Long
    w = Array(5, 18, 7.5, 8, 6.5, 9.5, 7.5, 18, 7.5, 8, 6.5, 7.5)
    For c = 0 To 11
        ws.Columns(c + 1).ColumnWidth = w(c)
    Next c
    ws.Columns(13).ColumnWidth = 1.5
    Day_BuildChrome          ' 열 너비가 바뀐 뒤 버튼 위치를 다시 맞춤

    Dim v As Variant
    v = RawData()
    Dim stdMap As Object, zoneMap As Object
    LoadRegionMaps stdMap, zoneMap
    Dim resMap As Object
    Set resMap = ResultInspMap(d)

    ' 제목 줄: 날짜 / 구분 / 머리글
    ws.rows(4).RowHeight = 26
    ws.rows(5).RowHeight = 24
    ws.rows(6).RowHeight = 24
    With ws.Range("H4:L4")
        .Merge
        .Value = Format(d, "m월 d일") & "  (" & WeekdayKo(d) & ")"
        .HorizontalAlignment = xlRight
        .Font.Bold = True
        .Font.Size = 14
    End With
    ws.Range("A5:G5").Merge
    ws.Range("A5").Value = "정기검사"
    ws.Range("H5:L5").Merge
    ws.Range("H5").Value = "사용전검사"
    ws.Range("A5:L5").HorizontalAlignment = xlCenter
    ws.Range("A5:L5").Font.Bold = True
    ws.Range("A5:L5").Font.Size = 14
    ws.Range("H5").Font.Color = RGB(255, 0, 0)

    Dim hdr As Variant
    hdr = Array("구분", "상호", "용량", "지역", "시간", "수수료", "검사자", "상호", "용량", "지역", "시간", "검사자")
    For c = 0 To 11
        ws.Cells(6, c + 1).Value = hdr(c)
    Next c
    ws.Range("A6:L6").HorizontalAlignment = xlCenter
    ws.Range("A6:L6").Font.Bold = True
    ws.Range("A6:L6").Font.Size = 11

    ' 구역별 본문
    Dim zones(1 To 3) As String
    zones(1) = ZONE_A
    zones(2) = ZONE_B
    zones(3) = "미분류"
    Dim curRow As Long
    curRow = 7

    Dim z As Long, i As Long, nR As Long, nP As Long, rows As Long
    Dim regIdx() As Long, preIdx() As Long, nTot As Long
    Dim hasOther As Boolean, hasDerived As Boolean
    Dim sRow As Long, key As Long, zt As String
    Dim notes As String

    If IsEmpty(v) Then nTot = 0 Else nTot = UBound(v, 1)
    key = CLng(Int(CDbl(d)))

    ' 부천/김포 외 구역이 있는지, 비고 모으기
    For i = 1 To nTot
        If DayKey(v(i, RW_DATE)) = key Then
            zt = ZoneText(v, i, zoneMap)
            If zt <> ZONE_A And zt <> ZONE_B Then hasOther = True
            If Trim$(CStr(v(i, RW_NOTE))) <> "" Then
                If notes <> "" Then notes = notes & " / "
                notes = notes & Trim$(CStr(v(i, RW_NAME))) & ": " & Trim$(CStr(v(i, RW_NOTE)))
            End If
        End If
    Next i

    Dim lastZone As Long
    lastZone = 2
    If hasOther Then lastZone = 3

    Dim out() As Variant, derA() As Boolean, derP() As Boolean
    Dim r As Long, tid As String, insp As String
    For z = 1 To lastZone
        ReDim regIdx(1 To IIf(nTot < 1, 1, nTot))
        ReDim preIdx(1 To IIf(nTot < 1, 1, nTot))
        nR = 0
        nP = 0
        For i = 1 To nTot
            If DayKey(v(i, RW_DATE)) = key Then
                zt = ZoneText(v, i, zoneMap)
                If zt = zones(z) Or (z = 3 And zt <> ZONE_A And zt <> ZONE_B) Then
                    If IsRegular(CStr(v(i, RW_KIND))) Then
                        nR = nR + 1
                        regIdx(nR) = i
                    Else
                        nP = nP + 1
                        preIdx(nP) = i
                    End If
                End If
            End If
        Next i
        rows = 10
        If nR > rows Then rows = nR
        If nP > rows Then rows = nP

        ReDim out(1 To rows, 1 To 12)
        ReDim derA(1 To rows)
        ReDim derP(1 To rows)
        For r = 1 To nR
            i = regIdx(r)
            out(r, 2) = v(i, RW_NAME)
            out(r, 3) = v(i, RW_CAP)
            out(r, 4) = v(i, RW_REG)
            out(r, 5) = TimeCell(v(i, RW_TIME))
            out(r, 6) = v(i, RW_FEE)
            insp = Trim$(CStr(v(i, RW_INSP)))
            If insp = "" Then
                tid = Trim$(CStr(v(i, RW_ID)))
                If resMap.Exists(tid) Then insp = resMap(tid): derA(r) = True: hasDerived = True
            End If
            out(r, 7) = insp
        Next r
        For r = 1 To nP
            i = preIdx(r)
            out(r, 8) = v(i, RW_NAME)
            out(r, 9) = v(i, RW_CAP)
            out(r, 10) = v(i, RW_REG)
            out(r, 11) = TimeCell(v(i, RW_TIME))
            insp = Trim$(CStr(v(i, RW_INSP)))
            If insp = "" And IsTaskKind(CStr(v(i, RW_KIND))) Then
                tid = Trim$(CStr(v(i, RW_ID)))
                If resMap.Exists(tid) Then insp = resMap(tid): derP(r) = True: hasDerived = True
            End If
            out(r, 12) = insp
        Next r

        sRow = curRow
        ws.Range(ws.Cells(sRow, 1), ws.Cells(sRow + rows - 1, 12)).Value = out
        FormatDaySection ws, sRow, rows, IIf(z = 3, "기타지역", zones(z))
        ' 배분 결과로 채운 검사자(원데이터 미반영)는 파란 기울임
        For r = 1 To rows
            If derA(r) Then
                ws.Cells(sRow + r - 1, 7).Font.Color = RGB(0, 70, 200)
                ws.Cells(sRow + r - 1, 7).Font.Italic = True
            End If
            If derP(r) Then
                ws.Cells(sRow + r - 1, 12).Font.Color = RGB(0, 70, 200)
                ws.Cells(sRow + r - 1, 12).Font.Italic = True
            End If
        Next r
        curRow = curRow + rows
    Next z

    ' 비고
    If hasDerived Then
        If notes <> "" Then notes = notes & "   "
        notes = notes & "[파란 기울임 = 자동배분 결과 (결과 반영 전)]"
    End If
    ws.Cells(curRow, 1).Value = "비고"
    ws.Range(ws.Cells(curRow, 2), ws.Cells(curRow, 12)).Merge
    ws.Cells(curRow, 2).Value = notes
    ws.rows(curRow).RowHeight = 26
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

Private Sub FormatDaySection(ws As Worksheet, ByVal sRow As Long, ByVal rows As Long, ByVal label As String)
    Dim eRow As Long
    eRow = sRow + rows - 1
    Dim rg As Range
    ' 구분 칸: 세로 병합 + 세로쓰기
    Set rg = ws.Range(ws.Cells(sRow, 1), ws.Cells(eRow, 1))
    rg.Merge
    rg.Value = label
    rg.Orientation = xlVertical
    rg.HorizontalAlignment = xlCenter
    rg.VerticalAlignment = xlCenter
    rg.Font.Bold = True
    rg.Font.Size = 11

    Set rg = ws.Range(ws.Cells(sRow, 1), ws.Cells(eRow, 12))
    rg.Borders.LineStyle = xlContinuous
    rg.Borders.Weight = xlThin
    rg.Borders(xlEdgeTop).Weight = xlMedium
    rg.Borders(xlEdgeBottom).Weight = xlMedium
    rg.Borders(xlEdgeLeft).Weight = xlMedium
    rg.Borders(xlEdgeRight).Weight = xlMedium
    ws.Range(ws.Cells(sRow, 8), ws.Cells(eRow, 8)).Borders(xlEdgeLeft).Weight = xlMedium
    ws.Range(ws.Cells(sRow, 2), ws.Cells(eRow, 2)).Borders(xlEdgeLeft).Weight = xlMedium

    ws.rows(sRow & ":" & eRow).RowHeight = 24
    ' 사용전검사 영역은 붉은 글씨
    ws.Range(ws.Cells(sRow, 8), ws.Cells(eRow, 12)).Font.Color = RGB(255, 0, 0)
    ' 정렬/표시 형식
    ws.Range(ws.Cells(sRow, 4), ws.Cells(eRow, 5)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(sRow, 10), ws.Cells(eRow, 11)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(sRow, 7), ws.Cells(eRow, 7)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(sRow, 12), ws.Cells(eRow, 12)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(sRow, 5), ws.Cells(eRow, 5)).NumberFormat = "h:mm"
    ws.Range(ws.Cells(sRow, 11), ws.Cells(eRow, 11)).NumberFormat = "h:mm"
    ws.Range(ws.Cells(sRow, 6), ws.Cells(eRow, 6)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(sRow, 6), ws.Cells(eRow, 6)).Font.Size = 9
    ws.Range(ws.Cells(sRow, 6), ws.Cells(eRow, 6)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(sRow, 3), ws.Cells(eRow, 3)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(sRow, 9), ws.Cells(eRow, 9)).HorizontalAlignment = xlRight
End Sub

' A4 가로 1쪽: 왼쪽 접수대장 + 오른쪽 현황판 전체가 인쇄 영역
Private Sub SetupDayPrint(ws As Worksheet, ByVal ledgerLast As Long)
    Dim lastRow As Long, lastCol As Long
    lastRow = ledgerLast
    If lastRow < GRID_START_ROW + 18 Then lastRow = GRID_START_ROW + 18
    lastCol = BoardLastCol()
    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(4, 1), ws.Cells(lastRow, lastCol)).Address
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
        .CenterHeader = ""
        .CenterFooter = "&8&P / &N"
    End With
    Application.PrintCommunication = True
    On Error GoTo 0
End Sub

' ============================================================
' 2. 월 접수대장 (원본 월별 시트와 같은 배치: 요일 = 행 블록, 주차 = 열 블록)
' ============================================================
Private Function MonthStart() As Date
    Dim v As Variant
    v = ThisWorkbook.names("월대장월").RefersToRange.Value
    If IsDate(v) Then
        MonthStart = DateSerial(Year(CDate(v)), Month(CDate(v)), 1)
    Else
        MonthStart = DateSerial(Year(BoardDate()), Month(BoardDate()), 1)
    End If
End Function

Sub Month_Prev()
    ThisWorkbook.names("월대장월").RefersToRange.Value = DateAdd("m", -1, MonthStart())
    Month_Render
End Sub

Sub Month_Next()
    ThisWorkbook.names("월대장월").RefersToRange.Value = DateAdd("m", 1, MonthStart())
    Month_Render
End Sub

Sub Month_ThisMonth()
    ThisWorkbook.names("월대장월").RefersToRange.Value = DateSerial(Year(Date), Month(Date), 1)
    Month_Render
End Sub

Sub Month_Pick()
    Dim d As Date
    If PickDate(MonthStart(), d) Then
        ThisWorkbook.names("월대장월").RefersToRange.Value = DateSerial(Year(d), Month(d), 1)
        Month_Render
    End If
End Sub

Public Sub Month_BuildChrome()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_MON)
    Dim k As Long
    For k = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(k).name, 4) = "MBtn" Then ws.Shapes(k).Delete
    Next k
    Dim names As Variant, caps As Variant, macros As Variant, i As Long
    names = Array("MBtnPrev", "MBtnThis", "MBtnNext", "MBtnPick")
    caps = Array(ChrW(9664) & " 이전달", "이번달", "다음달 " & ChrW(9654), ChrW(9638) & " 달력")
    macros = Array("Month_Prev", "Month_ThisMonth", "Month_Next", "Month_Pick")
    Dim L As Double, shp As Shape
    L = ws.Cells(2, 2).Left + 2
    For i = 0 To 3
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, L, ws.Cells(2, 2).Top + 2, 80, 22)
        shp.name = names(i)
        shp.Fill.ForeColor.RGB = RGB(47, 117, 181)
        shp.Line.Visible = msoFalse
        shp.Placement = xlMove
        shp.OnAction = macros(i)
        With shp.TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.Text = caps(i)
            .TextRange.Font.Size = 10
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
        L = L + 84
    Next i
End Sub

Public Sub Month_Render()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_MON)
    Dim prevSU As Boolean, prevEv As Boolean
    prevSU = Application.ScreenUpdating
    prevEv = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo fail

    Dim mFirst As Date, mLast As Date, mon0 As Date
    mFirst = MonthStart()
    mLast = DateAdd("d", -1, DateAdd("m", 1, mFirst))
    mon0 = mFirst - (Weekday(mFirst, vbMonday) - 1)
    Dim nWeeks As Long
    nWeeks = ((mLast - mon0) \ 7) + 1
    If nWeeks > 6 Then nWeeks = 6

    ' 영역 초기화
    With ws.Range(ws.Cells(5, 1), ws.Cells(5 + 24 * 7 + 2, 13 * 6 + 3))
        .UnMerge
        .Clear
        .Font.name = "맑은 고딕"
        .Font.Size = 10
        .VerticalAlignment = xlCenter
    End With
    ws.Cells(1, 2).Value = Format(mFirst, "yyyy년 m월") & " 접수대장 (원데이터 기준 자동 생성)"
    ws.Cells(1, 2).Font.Bold = True
    ws.Cells(1, 2).Font.Size = 13
    ws.Cells(1, 2).Font.Color = RGB(31, 56, 100)

    ' 열 너비
    Dim w As Variant, j As Long, c As Long
    w = Array(9, 20, 8, 9, 6.5, 10, 8, 20, 8, 9, 6.5, 8, 2)
    For j = 0 To 5
        For c = 0 To 12
            ws.Columns(2 + 13 * j + c).ColumnWidth = w(c)
        Next c
    Next j
    Month_BuildChrome        ' 열 너비가 바뀐 뒤 버튼 위치를 다시 맞춤

    ' 월 내 데이터를 일자별 버킷에 모음
    Dim v As Variant, nTot As Long, i As Long
    v = RawData()
    If IsEmpty(v) Then nTot = 0 Else nTot = UBound(v, 1)
    Dim stdMap As Object, zoneMap As Object
    LoadRegionMaps stdMap, zoneMap
    Dim bucket(1 To 31) As Collection
    For i = 1 To 31
        Set bucket(i) = New Collection
    Next i
    Dim dk As Long
    For i = 1 To nTot
        dk = DayKey(v(i, RW_DATE))
        If dk > 0 Then
            If dk >= CLng(mFirst) And dk <= CLng(mLast) Then
                bucket(Day(CDate(dk))).Add i
            End If
        End If
    Next i

    Dim k As Long, dt As Date, r0 As Long, c0 As Long
    For j = 0 To nWeeks - 1
        For k = 0 To 6
            dt = mon0 + 7 * j + k
            r0 = 5 + 24 * k
            c0 = 2 + 13 * j
            If dt >= mFirst And dt <= mLast Then
                MonthBlock ws, v, bucket(Day(dt)), dt, r0, c0, zoneMap
            End If
        Next k
    Next j

    SetupMonthPrint ws, nWeeks
fail:
    Dim em As String
    If Err.Number <> 0 Then em = Err.Description
    Application.EnableEvents = prevEv
    Application.ScreenUpdating = prevSU
    If em <> "" Then Say "월대장을 만드는 중 오류가 발생했습니다: " & em, vbCritical
End Sub

Private Sub MonthBlock(ws As Worksheet, v As Variant, bk As Collection, ByVal dt As Date, _
        ByVal r0 As Long, ByVal c0 As Long, zoneMap As Object)
    Dim out() As Variant
    ReDim out(1 To 20, 1 To 12)
    Dim nA(1 To 2) As Long, nP(1 To 2) As Long, ovf(1 To 2) As Long, x As Variant, i As Long
    Dim zi As Long, rr As Long, z As String
    For Each x In bk
        i = CLng(x)
        z = ZoneText(v, i, zoneMap)
        If z = ZONE_B Then zi = 2 Else zi = 1          ' 미분류는 부천 칸 아래쪽에 표시
        If IsRegular(CStr(v(i, RW_KIND))) Then
            nA(zi) = nA(zi) + 1
            If nA(zi) <= 10 Then
                rr = (zi - 1) * 10 + nA(zi)
                out(rr, 2) = v(i, RW_NAME)
                out(rr, 3) = v(i, RW_CAP)
                out(rr, 4) = v(i, RW_REG)
                out(rr, 5) = TimeCell(v(i, RW_TIME))
                out(rr, 6) = v(i, RW_FEE)
                out(rr, 7) = v(i, RW_INSP)
            Else
                ovf(zi) = ovf(zi) + 1
            End If
        Else
            nP(zi) = nP(zi) + 1
            If nP(zi) <= 10 Then
                rr = (zi - 1) * 10 + nP(zi)
                out(rr, 8) = v(i, RW_NAME)
                out(rr, 9) = v(i, RW_CAP)
                out(rr, 10) = v(i, RW_REG)
                out(rr, 11) = TimeCell(v(i, RW_TIME))
                out(rr, 12) = v(i, RW_INSP)
            Else
                ovf(zi) = ovf(zi) + 1
            End If
        End If
    Next x
    out(1, 1) = ZONE_A
    out(11, 1) = ZONE_B

    ' 날짜 / 요일
    ws.Range(ws.Cells(r0, c0), ws.Cells(r0, c0 + 1)).Merge
    With ws.Cells(r0, c0)
        .Value = dt
        .NumberFormat = "m""월"" d""일"""
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
        .Font.Size = 12
    End With
    With ws.Cells(r0, c0 + 11)
        .Value = "(" & WeekdayKo(dt) & ")"
        .HorizontalAlignment = xlRight
        .Font.Bold = True
        .Font.Size = 12
    End With

    ' 정기검사 / 사용전검사 표제
    ws.Range(ws.Cells(r0 + 1, c0), ws.Cells(r0 + 1, c0 + 6)).Merge
    ws.Cells(r0 + 1, c0).Value = "정기검사"
    ws.Range(ws.Cells(r0 + 1, c0 + 7), ws.Cells(r0 + 1, c0 + 11)).Merge
    ws.Cells(r0 + 1, c0 + 7).Value = "사용전검사"
    ws.Cells(r0 + 1, c0 + 7).Font.Color = RGB(255, 0, 0)
    ws.Range(ws.Cells(r0 + 1, c0), ws.Cells(r0 + 1, c0 + 11)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(r0 + 1, c0), ws.Cells(r0 + 1, c0 + 11)).Font.Bold = True

    ' 머리글
    Dim hdr As Variant, c As Long
    hdr = Array("구분", "상호", "용량", "지역", "시간", "수수료", "검사자", "상호", "용량", "지역", "시간", "검사자")
    For c = 0 To 11
        ws.Cells(r0 + 2, c0 + c).Value = hdr(c)
    Next c
    ws.Range(ws.Cells(r0 + 2, c0), ws.Cells(r0 + 2, c0 + 11)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(r0 + 2, c0), ws.Cells(r0 + 2, c0 + 11)).Font.Bold = True

    ' 본문 20행
    Dim body As Range
    Set body = ws.Range(ws.Cells(r0 + 3, c0), ws.Cells(r0 + 22, c0 + 11))
    body.Value = out
    ws.Range(ws.Cells(r0 + 3, c0 + 7), ws.Cells(r0 + 22, c0 + 11)).Font.Color = RGB(255, 0, 0)
    ws.Range(ws.Cells(r0 + 3, c0 + 4), ws.Cells(r0 + 22, c0 + 4)).NumberFormat = "h:mm"
    ws.Range(ws.Cells(r0 + 3, c0 + 10), ws.Cells(r0 + 22, c0 + 10)).NumberFormat = "h:mm"
    ws.Range(ws.Cells(r0 + 3, c0 + 5), ws.Cells(r0 + 22, c0 + 5)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(r0 + 3, c0 + 5), ws.Cells(r0 + 22, c0 + 5)).Font.Size = 8
    ws.Range(ws.Cells(r0 + 3, c0 + 3), ws.Cells(r0 + 22, c0 + 4)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(r0 + 3, c0 + 9), ws.Cells(r0 + 22, c0 + 10)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(r0 + 3, c0 + 6), ws.Cells(r0 + 22, c0 + 6)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(r0 + 3, c0 + 11), ws.Cells(r0 + 22, c0 + 11)).HorizontalAlignment = xlCenter
    ws.Cells(r0 + 3, c0).Font.Bold = True
    ws.Cells(r0 + 13, c0).Font.Bold = True

    ' 비고 행
    ws.Cells(r0 + 23, c0).Value = "비고"
    ws.Cells(r0 + 23, c0).Font.Bold = True
    If ovf(1) + ovf(2) > 0 Then
        ws.Cells(r0 + 23, c0 + 1).Value = "※ 표시 한도(구역당 10건) 초과 " & (ovf(1) + ovf(2)) & "건 - 일일대장에서 확인"
        ws.Cells(r0 + 23, c0 + 1).Font.Color = RGB(192, 0, 0)
    End If

    ' 테두리
    Dim rg As Range
    Set rg = ws.Range(ws.Cells(r0 + 1, c0), ws.Cells(r0 + 23, c0 + 11))
    rg.Borders.LineStyle = xlContinuous
    rg.Borders.Weight = xlThin
    rg.Borders.Color = RGB(0, 0, 0)
    rg.Borders(xlEdgeLeft).Weight = xlMedium
    rg.Borders(xlEdgeRight).Weight = xlMedium
    rg.Borders(xlEdgeTop).Weight = xlMedium
    rg.Borders(xlEdgeBottom).Weight = xlMedium
    ws.Range(ws.Cells(r0 + 13, c0), ws.Cells(r0 + 13, c0 + 11)).Borders(xlEdgeTop).Weight = xlMedium
    ws.Range(ws.Cells(r0 + 1, c0 + 7), ws.Cells(r0 + 23, c0 + 7)).Borders(xlEdgeLeft).Weight = xlMedium
End Sub

Private Sub SetupMonthPrint(ws As Worksheet, ByVal nWeeks As Long)
    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(5 + 24 * 7 - 1, 1 + 13 * nWeeks)).Address
        .PaperSize = xlPaperA3
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintGridlines = False
    End With
    Application.PrintCommunication = True
    On Error GoTo 0
End Sub

' ============================================================
' 3. 원데이터 <-> 자동배분 연동
' ============================================================

' 배분 결과(선택일)의 검사자를 10_원데이터에 기록
'   - 원데이터 검사자가 비어 있으면 채움
'   - 이미 있는 값과 다르면: 현황판에서 수동으로 옮긴 건만 바꾸고, 나머지는 건드리지 않고 07_배분로그에 기록
Sub Erp_ApplyAssignment()
    Dim wsRes As Worksheet, lo As ListObject
    Set wsRes = ThisWorkbook.Sheets(SH_RESULT)
    Set lo = RawTable()
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Dim d As Date
    d = BoardDate()

    Dim ids As Object, aliasMap As Object
    Set ids = RawIdIndex()
    Set aliasMap = LoadAliasMap()

    Dim last As Long, r As Long, tid As Long, key As String, insp As String, cur As String
    Dim cnt As Long, chg As Long, skip As Long
    last = wsRes.Cells(wsRes.rows.Count, R_NAME).End(xlUp).Row
    Dim prevEv As Boolean
    prevEv = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo done
    For r = 2 To last
        If IsDate(wsRes.Cells(r, R_DATE).Value) Then
            If DateValue(wsRes.Cells(r, R_DATE).Value) = d Then
                tid = CLng(wsRes.Cells(r, R_ID).Value)
                insp = Trim$(CStr(wsRes.Cells(r, R_INSP).Value))
                If tid > 0 And insp <> "" And CLng(wsRes.Cells(r, R_SLOT).Value) > 0 Then
                    key = CStr(tid)
                    If ids.Exists(key) Then
                        With lo.DataBodyRange.Cells(ids(key), RW_INSP)
                            cur = Trim$(CStr(.Value))
                            If cur = "" Then
                                .Value = insp
                                cnt = cnt + 1
                            ElseIf LeadInsp(cur, aliasMap) <> insp Then
                                If CStr(wsRes.Cells(r, R_MANUAL).Value) = "Y" Then
                                    AddLog "수동조정", "원데이터 검사자(" & cur & ")를 수동 조정 결과(" & insp & ")로 바꿨습니다.", CStr(wsRes.Cells(r, R_NAME).Value)
                                    .Value = insp
                                    chg = chg + 1
                                Else
                                    skip = skip + 1
                                    AddLog "수동조정", "원데이터 검사자(" & cur & ")와 배분 결과(" & insp & ")가 달라 반영하지 않았습니다.", CStr(wsRes.Cells(r, R_NAME).Value)
                                End If
                            End If
                        End With
                    End If
                End If
            End If
        End If
    Next r
done:
    Dim em As String
    If Err.Number <> 0 Then em = Err.Description
    Application.EnableEvents = prevEv
    Day_Render
    If em <> "" Then Say "결과를 반영하는 중 오류가 발생했습니다: " & em, vbCritical
    Say "검사자 " & cnt & "건을 10_원데이터에 기록했습니다." & IIf(chg > 0, vbCrLf & "수동 조정으로 바뀐 건 " & chg, "") & _
        IIf(skip > 0, vbCrLf & "기존 값과 달라 반영하지 않은 건 " & skip & " (07_배분로그 확인)", ""), vbInformation
End Sub

' 원데이터 표에 입력/수정이 생기면 자동 정리:
'   관리번호 부여 / 권역 자동 판정(04_지역표) / 검사자 별칭을 정규 이름으로
Public Sub Erp_OnRawChange(ByVal Target As Range)
    Dim lo As ListObject
    Set lo = RawTable()
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Dim hit As Range
    Set hit = Intersect(Target, lo.DataBodyRange)
    If hit Is Nothing Then Exit Sub
    If hit.CountLarge > 20000 Then Exit Sub          ' 정렬 등 대량 변경은 건너뜀
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
            ' 검사자 이름 정규화
            Set cel = rw.Cells(1, RW_INSP)
            s = Trim$(CStr(cel.Value))
            If s <> "" Then
                t = NormInsp(s, aliasMap)
                If t <> CStr(cel.Value) Then cel.Value = t
            End If
            ' 권역 자동 판정
            If Trim$(CStr(rw.Cells(1, RW_ZONE).Value)) = "" Then
                s = Trim$(CStr(rw.Cells(1, RW_REG).Value))
                If s <> "" Then
                    If zoneMap.Exists(s) Then rw.Cells(1, RW_ZONE).Value = zoneMap(s)
                End If
            End If
            ' 관리번호 부여
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
