Attribute VB_Name = "Module5_월대장"
Option Explicit

' ============================================================
' KESCO 부천김포지사 검사업무 ERP
' Module5_월대장 : 월 접수대장 (원본 월별 시트와 같은 배치: 요일 = 행 블록, 주차 = 열 블록)
'   원데이터(10_원데이터)에서 자동 생성 (읽기 전용 화면). 월/원데이터가 바뀐 경우에만 다시 그립니다.
' ============================================================

Private monKey As Long
Private monVer As Long
Private monDrawn As Boolean

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

Private Function TimeCell(ByVal x As Variant) As Variant
    If IsError(x) Then
        TimeCell = ""
    ElseIf IsEmpty(x) Then
        TimeCell = ""
    Else
        TimeCell = x
    End If
End Function

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

' 시트 이벤트(Worksheet_Activate)에서 호출
Public Sub Month_Activate()
    If monDrawn Then
        If monKey = CLng(MonthStart()) And monVer = gRawVer Then Exit Sub
    End If
    Month_Render
End Sub

Private Function ShapeThere(ws As Worksheet, ByVal nm As String) As Boolean
    Dim s As Shape
    On Error Resume Next
    Set s = ws.Shapes(nm)
    On Error GoTo 0
    ShapeThere = Not s Is Nothing
End Function

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
    FastBegin
    On Error GoTo fail

    Dim mFirst As Date, mLast As Date, mon0 As Date
    mFirst = MonthStart()
    mLast = DateAdd("d", -1, DateAdd("m", 1, mFirst))
    mon0 = mFirst - (Weekday(mFirst, vbMonday) - 1)
    Dim nWeeks As Long
    nWeeks = ((mLast - mon0) \ 7) + 1
    If nWeeks > 6 Then nWeeks = 6

    With ws.Range(ws.Cells(5, 1), ws.Cells(5 + 24 * 7 + 2, 13 * 6 + 3))
        .UnMerge
        .Clear
        .Font.name = "맑은 고딕"
        .Font.Size = 10
        .VerticalAlignment = xlCenter
    End With
    ws.Cells(1, 2).Value = Format(mFirst, "yyyy년 m월") & " 접수대장 (원데이터 기준 자동 생성 / 입력은 11_일일대장)"
    ws.Cells(1, 2).Font.Bold = True
    ws.Cells(1, 2).Font.Size = 13
    ws.Cells(1, 2).Font.Color = RGB(31, 56, 100)

    Dim w As Variant, j As Long, c As Long
    w = Array(9, 20, 8, 9, 6.5, 10, 8, 20, 8, 9, 6.5, 8, 2)
    For j = 0 To 5
        For c = 0 To 12
            ws.Columns(2 + 13 * j + c).ColumnWidth = w(c)
        Next c
    Next j
    If Not ShapeThere(ws, "MBtnPick") Then Month_BuildChrome

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
    monKey = CLng(mFirst)
    monVer = gRawVer
    monDrawn = True
fail:
    Dim em As String
    If Err.Number <> 0 Then em = Err.Description
    FastEnd
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
        If z = ZONE_B Then zi = 2 Else zi = 1
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

    ws.Range(ws.Cells(r0 + 1, c0), ws.Cells(r0 + 1, c0 + 6)).Merge
    ws.Cells(r0 + 1, c0).Value = "정기검사"
    ws.Range(ws.Cells(r0 + 1, c0 + 7), ws.Cells(r0 + 1, c0 + 11)).Merge
    ws.Cells(r0 + 1, c0 + 7).Value = "사용전검사"
    ws.Cells(r0 + 1, c0 + 7).Font.Color = RGB(255, 0, 0)
    ws.Range(ws.Cells(r0 + 1, c0), ws.Cells(r0 + 1, c0 + 11)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(r0 + 1, c0), ws.Cells(r0 + 1, c0 + 11)).Font.Bold = True

    ws.Range(ws.Cells(r0 + 2, c0), ws.Cells(r0 + 2, c0 + 11)).Value = Array("구분", "상호", "용량", "지역", "시간", "수수료", "검사자", "상호", "용량", "지역", "시간", "검사자")
    ws.Range(ws.Cells(r0 + 2, c0), ws.Cells(r0 + 2, c0 + 11)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(r0 + 2, c0), ws.Cells(r0 + 2, c0 + 11)).Font.Bold = True

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

    ws.Cells(r0 + 23, c0).Value = "비고"
    ws.Cells(r0 + 23, c0).Font.Bold = True
    If ovf(1) + ovf(2) > 0 Then
        ws.Cells(r0 + 23, c0 + 1).Value = "※ 표시 한도(구역당 10건) 초과 " & (ovf(1) + ovf(2)) & "건 - 일일대장에서 확인"
        ws.Cells(r0 + 23, c0 + 1).Font.Color = RGB(192, 0, 0)
    End If

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
    Dim want As String, cur As String
    want = ws.Range(ws.Cells(1, 1), ws.Cells(5 + 24 * 7 - 1, 1 + 13 * nWeeks)).Address
    On Error Resume Next
    cur = ws.names("Print_Area").RefersToRange.Address
    If cur = want Then Exit Sub
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = want
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
