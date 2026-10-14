Attribute VB_Name = "Module3_현황판"
Option Explicit

' ============================================================
' KESCO 부천김포지사 검사업무 ERP
' Module3_현황판 : 11_일일대장 시트 오른쪽(O열~)에 놓인 24시간 배분현황판
'   - 09_배분결과 표를 읽어 현황판 위에 "업무 블록(도형)"을 그림
'   - 블록 클릭 -> 상단 상세 패널 / 블록 드래그 -> 검사자 열/시간으로 환산해 저장
'     (겹치거나 시간축을 벗어나면 자동으로 고치지 않고 원위치 + 07_배분로그 '충돌' 기록)
' 블록 이름 규칙: Blk_<09_배분결과 행번호>
' ============================================================

Private Const SCROLL_STEP As Long = 4      ' 한 번에 4칸(2시간)씩 이동
Private pollOn As Boolean
Private pollNext As Date
Private busy As Boolean

' ------------------------------------------------------------
' 달력 팝업(frmCalendar)을 띄워 날짜를 고름. 선택하면 True.
' ------------------------------------------------------------
Public Function PickDate(ByVal initial As Date, ByRef res As Date, _
        Optional ByVal allowClear As Boolean = False, Optional ByRef cleared As Boolean = False) As Boolean
    Dim f As frmCalendar
    Set f = New frmCalendar
    f.Init initial, allowClear
    f.Show vbModal
    If f.picked Then
        cleared = f.cleared
        res = f.PickedDate
        PickDate = True
    End If
    Unload f
End Function

' 현황판 머리글(6행)에서 검사자 이름과 '미지정' 열 위치를 구함
Private Function GetBoardCols(names() As String, ByRef unCol As Long) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_DAY)
    Dim c As Long, n As Long, v As String
    ReDim names(1 To 60)
    c = GRID_COL0 + 1
    Do While n < 59
        v = Trim(CStr(ws.Cells(GRID_HEADER_ROW, c).Value))
        If v = "" Or v = "미지정" Then Exit Do
        n = n + 1
        names(n) = v
        c = c + 1
    Loop
    unCol = c
    GetBoardCols = n
End Function

Public Function BoardLastCol() As Long
    Dim names() As String, unCol As Long
    GetBoardCols names, unCol
    BoardLastCol = unCol
End Function

Private Function ShapeExists(ws As Worksheet, ByVal nm As String) As Boolean
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes(nm)
    On Error GoTo 0
    ShapeExists = Not shp Is Nothing
End Function

Private Function LegendIdx(ByVal typ As Long) As Long
    Select Case typ
        Case 1: LegendIdx = 1
        Case 2: LegendIdx = 2
        Case 3: LegendIdx = 3
        Case 4: LegendIdx = 4
        Case 5: LegendIdx = 5
        Case 7: LegendIdx = 6
        Case Else: LegendIdx = 7     ' 0 = 배치 실패
    End Select
End Function

Private Function TypeLabel(ByVal typ As Long) As String
    Select Case typ
        Case 1: TypeLabel = "검사자+시간 지정 (1순위)"
        Case 2: TypeLabel = "시간만 지정 (2순위)"
        Case 3: TypeLabel = "검사자만 지정 (3순위)"
        Case 4: TypeLabel = "미지정 업무 (5순위)"
        Case 5: TypeLabel = "점심시간 (4순위)"
        Case 7: TypeLabel = "수동 배정"
        Case Else: TypeLabel = "배치 실패 (미지정 열)"
    End Select
End Function

' ------------------------------------------------------------
' 현황판 격자를 화면 기준 근무자(휴가/교육 메모 제외)로 다시 그리고 블록을 표시
' 11_일일대장의 A~N열(접수대장)은 건드리지 않습니다.
' ------------------------------------------------------------
Public Sub RebuildGrid()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_DAY)
    FastBegin
    On Error GoTo bail

    Dim rec() As TaskRec, nRec As Long, leave As Object
    nRec = ReadScreen(rec)
    Set leave = LeaveFromRecs(rec, nRec)
    Dim inspectors() As String
    inspectors = GetActiveInspectorsFor(leave)
    Dim n As Long
    If HasInspectors(inspectors) Then n = UBound(inspectors) - LBound(inspectors) + 1 Else n = 0

    Dim unCol As Long, c As Long
    unCol = GRID_COL0 + n + 1

    ' 현황판 영역 초기화 (O열 이후만)
    With ws.Range(ws.Cells(4, GRID_COL0), ws.Cells(GRID_START_ROW + 60, GRID_COL0 + 40))
        .UnMerge
        .Clear
        .Font.name = "맑은 고딕"
        .Font.Size = 9
        .VerticalAlignment = xlTop
    End With

    Dim navy As Long, paleBlue As Long
    navy = RGB(31, 56, 100)
    paleBlue = RGB(221, 235, 247)

    ' 열 너비 / 머리글 (한 번에)
    ws.Columns(GRID_COL0).ColumnWidth = 5.5
    ws.Range(ws.Columns(GRID_COL0 + 1), ws.Columns(unCol)).ColumnWidth = 8.6
    ws.Columns(unCol + 1).ColumnWidth = 2
    Dim hdr() As Variant
    ReDim hdr(1 To 1, 1 To n + 2)
    hdr(1, 1) = "시간"
    For c = 1 To n
        hdr(1, c + 1) = inspectors(c)
    Next c
    hdr(1, n + 2) = "미지정"
    With ws.Range(ws.Cells(GRID_HEADER_ROW, GRID_COL0), ws.Cells(GRID_HEADER_ROW, unCol))
        .Value = hdr
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = navy
        .Interior.Color = paleBlue
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ShrinkToFit = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeBottom).Color = navy
    End With

    DrawTimeWindow ws, unCol

    Board_BuildChrome
    Board_Refresh
bail:
    Dim em As String
    If Err.Number <> 0 Then em = Err.Description
    FastEnd
    If em <> "" Then Say "현황판을 그리는 중 오류가 발생했습니다: " & em, vbCritical
End Sub

' ------------------------------------------------------------
' 시간대 창: 7행부터 BoardWinRows 칸만 보이고, [▲][▼]로 창을 움직임 (접수대장은 제자리)
'   라벨/격자/근무시간 밖 음영만 다시 그림 (검사자 열은 건드리지 않음)
' ------------------------------------------------------------
Private Sub DrawTimeWindow(ws As Worksheet, ByVal unCol As Long)
    Dim navy As Long, paleBlue As Long, dotBlue As Long, gridBlue As Long
    navy = RGB(31, 56, 100)
    paleBlue = RGB(221, 235, 247)
    dotBlue = RGB(91, 155, 213)
    gridBlue = RGB(191, 207, 224)

    Dim startTime As Date, endTime As Date, slotMin As Long, nSlots As Long, workLo As Long, workHi As Long
    startTime = DispStartTime()
    endTime = DispEndTime()
    slotMin = SettingLong("슬롯단위_분")
    nSlots = CalcSlotCount(startTime, endTime, slotMin)
    workLo = WorkFirstSlot()
    workHi = WorkLastSlot()
    Dim winN As Long, topS As Long, i As Long, sl As Long
    winN = BoardWinRows()
    topS = BoardTopSlot()

    ' 본문 영역 초기화 (O열 이후, 창 높이 전체)
    With ws.Range(ws.Cells(GRID_START_ROW, GRID_COL0), ws.Cells(GRID_START_ROW + BOARD_WIN_ROWS - 1, GRID_COL0 + 40))
        .Clear
        .Font.name = "맑은 고딕"
        .Font.Size = 9
        .VerticalAlignment = xlTop
    End With

    ' 시간 라벨 (창에 보이는 칸만; 시간축 끝의 다음 칸 = 마감 행)
    Dim lab() As Variant
    ReDim lab(1 To winN, 1 To 1)
    For i = 1 To winN
        sl = topS + i - 1
        If sl <= nSlots Then
            lab(i, 1) = SlotIndexToTimeStr(sl, startTime, slotMin)
        ElseIf sl = nSlots + 1 Then
            lab(i, 1) = TimeLabel(endTime)
        End If
    Next i
    Dim lastRow As Long, lastSlotRow As Long
    lastRow = GRID_START_ROW + winN - 1
    With ws.Range(ws.Cells(GRID_START_ROW, GRID_COL0), ws.Cells(lastRow, GRID_COL0))
        .NumberFormat = "@"
        .Value = lab
        .Font.Size = 8
        .Font.Bold = True
        .Font.Color = navy
        .Interior.Color = paleBlue
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlTop
    End With
    ws.Range(ws.rows(GRID_START_ROW), ws.rows(lastRow)).RowHeight = BODY_ROW_H

    ' 슬롯 칸이 있는 마지막 행 (마감 행이 창에 들어오면 그 바로 위)
    lastSlotRow = lastRow
    If topS + winN - 1 >= nSlots + 1 Then lastSlotRow = GRID_START_ROW + (nSlots - topS)
    If lastSlotRow >= GRID_START_ROW Then
        With ws.Range(ws.Cells(GRID_START_ROW, GRID_COL0 + 1), ws.Cells(lastSlotRow, unCol))
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideVertical).Color = gridBlue
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Color = gridBlue
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Color = gridBlue
            .Borders(xlInsideHorizontal).LineStyle = xlDot
            .Borders(xlInsideHorizontal).Color = dotBlue
            .Borders(xlEdgeBottom).LineStyle = xlDot
            .Borders(xlEdgeBottom).Color = dotBlue
        End With
        ' 근무시간 밖은 연한 회색 (창 안의 해당 구간만)
        Dim r1 As Long, r2 As Long
        If workLo > topS Then
            r1 = GRID_START_ROW
            r2 = GRID_START_ROW + workLo - topS - 1
            If r2 > lastSlotRow Then r2 = lastSlotRow
            ws.Range(ws.Cells(r1, GRID_COL0 + 1), ws.Cells(r2, unCol)).Interior.Color = RGB(244, 244, 244)
        End If
        If workHi < nSlots Then
            r1 = GRID_START_ROW + workHi + 1 - topS
            If r1 < GRID_START_ROW Then r1 = GRID_START_ROW
            If r1 <= lastSlotRow Then ws.Range(ws.Cells(r1, GRID_COL0 + 1), ws.Cells(lastSlotRow, unCol)).Interior.Color = RGB(244, 244, 244)
        End If
    End If
    ' 마감 행 (하루 끝 선)
    If lastSlotRow < lastRow Then
        With ws.Range(ws.Cells(lastRow, GRID_COL0 + 1), ws.Cells(lastRow, unCol))
            .Interior.Color = RGB(242, 242, 242)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = navy
        End With
        ws.Cells(lastRow, GRID_COL0).Font.Color = RGB(128, 128, 128)
        ws.Cells(lastRow, GRID_COL0 + 1).Value = "하루 끝"
        ws.Cells(lastRow, GRID_COL0 + 1).Font.Color = RGB(128, 128, 128)
        ws.Cells(lastRow, GRID_COL0 + 1).Font.Size = 8
    End If
End Sub

' 슬롯(시간 칸) -> 화면 행. 창 밖이면 창 위/아래 바깥 행번호가 나옴
Private Function SlotRow(ByVal slot As Long) As Long
    SlotRow = GRID_START_ROW + slot - BoardTopSlot()
End Function

' 슬롯 slot ~ slot+nS-1 중 창에 보이는 행 범위. 하나도 안 보이면 False
Private Function VisibleRows(ByVal slot As Long, ByVal nS As Long, ByRef dr1 As Long, ByRef dr2 As Long) As Boolean
    Dim winBot As Long
    winBot = GRID_START_ROW + BoardWinRows() - 1
    dr1 = SlotRow(slot)
    dr2 = dr1 + nS - 1
    If dr2 < GRID_START_ROW Or dr1 > winBot Then Exit Function
    If dr1 < GRID_START_ROW Then dr1 = GRID_START_ROW
    If dr2 > winBot Then dr2 = winBot
    VisibleRows = True
End Function

' ------------------------------------------------------------
' 시간대 창 스크롤 ([▲] [▼] [근무시간] 버튼). 블록을 끌어 놓은 직후라면 그 이동부터 처리
' ------------------------------------------------------------
Private Sub Board_Scroll(ByVal delta As Long)
    Board_CheckMoves
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_DAY)
    Dim before As Long
    before = BoardTopSlot()
    If delta = 0 Then
        gBoardTop = 0
    Else
        gBoardTop = before + delta
    End If
    If BoardTopSlot() = before And delta <> 0 Then Exit Sub
    Dim names() As String, unCol As Long
    GetBoardCols names, unCol
    FastBegin
    On Error GoTo bail
    DrawTimeWindow ws, unCol
    Board_Refresh
bail:
    FastEnd
End Sub

Public Sub Board_ScrollUp()
    Board_Scroll -SCROLL_STEP
End Sub

Public Sub Board_ScrollDown()
    Board_Scroll SCROLL_STEP
End Sub

Public Sub Board_ScrollHome()
    Board_Scroll 0
End Sub

' ------------------------------------------------------------
' 제목줄/범례/상세 패널 (여러 번 호출해도 안전)
' ------------------------------------------------------------
Public Sub Board_BuildChrome()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_DAY)
    Dim names() As String, unCol As Long, n As Long
    n = GetBoardCols(names, unCol)

    Dim navy As Long
    navy = RGB(31, 56, 100)

    ' 제목 막대 (5행, 남색)
    With ws.Range(ws.Cells(5, GRID_COL0), ws.Cells(5, unCol))
        .Merge
        .Interior.Color = navy
        .Font.name = "맑은 고딕"
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    Board_UpdateHeader
    Board_BuildScrollButtons ws, unCol

    ' 범례 (4행, 현황판 오른쪽 끝에 맞춰 7칸)
    Dim labels As Variant, i As Long, startCol As Long
    labels = Array("검사자+시간", "시간만", "검사자만", "미지정", "점심", "수동", "배치실패")
    startCol = unCol - 6
    If startCol < GRID_COL0 + 1 Then startCol = GRID_COL0 + 1
    For i = 0 To 6
        If startCol + i <= unCol Then
            With ws.Cells(4, startCol + i)
                .Value = labels(i)
                .Interior.Color = LegendColor(i + 1)
                .Font.name = "맑은 고딕"
                .Font.Size = 8
                .Font.Bold = True
                .Font.Color = RGB(255, 255, 255)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .ShrinkToFit = True
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Color = RGB(255, 255, 255)
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Color = RGB(255, 255, 255)
            End With
        End If
    Next i

    ' 상세 패널: 버튼 줄 오른쪽 ~ 현황판 끝, 1~2행 (인쇄 영역 밖)
    Dim shp As Shape
    Dim pLeft As Double, pTop As Double, pW As Double, pH As Double
    pLeft = ws.Cells(1, GRID_COL0).Left
    If ShapeExists(ws, "DBtnPdf") Then
        If ws.Shapes("DBtnPdf").Left + ws.Shapes("DBtnPdf").Width + 6 > pLeft Then pLeft = ws.Shapes("DBtnPdf").Left + ws.Shapes("DBtnPdf").Width + 6
    End If
    pTop = ws.Cells(1, 1).Top + 1
    pW = ws.Cells(1, unCol + 1).Left - pLeft
    If pW < 160 Then pW = 160
    pH = ws.Cells(4, 1).Top - pTop - 2
    If Not ShapeExists(ws, "Panel") Then
        Set shp = ws.Shapes.AddShape(msoShapeRectangle, pLeft, pTop, pW, pH)
        shp.name = "Panel"
        shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
        shp.Line.ForeColor.RGB = navy
        shp.Line.Weight = 1
        shp.Placement = xlMove
        With shp.TextFrame2
            .VerticalAnchor = msoAnchorTop
            .MarginLeft = 8: .MarginTop = 3: .MarginRight = 6: .MarginBottom = 1
            .WordWrap = msoTrue
            .TextRange.ParagraphFormat.Alignment = msoAlignLeft
            .TextRange.Font.Size = 8
            .TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        End With
    Else
        Set shp = ws.Shapes("Panel")
        shp.Left = pLeft
        shp.Top = pTop
        shp.Width = pW
        shp.Height = pH
    End If
End Sub

' 제목 막대(5행) 오른쪽 끝의 시간대 스크롤 버튼
Private Sub Board_BuildScrollButtons(ws As Worksheet, ByVal unCol As Long)
    Dim names As Variant, caps As Variant, macros As Variant, wd As Variant, i As Long
    names = Array("BBtnUp", "BBtnDn", "BBtnHome")
    caps = Array(ChrW(9650), ChrW(9660), "근무시간")
    macros = Array("Board_ScrollUp", "Board_ScrollDown", "Board_ScrollHome")
    wd = Array(26, 26, 56)
    Dim rEdge As Double, L As Double, shp As Shape, tp As Double
    rEdge = ws.Cells(5, unCol).Left + ws.Cells(5, unCol).Width - 4
    tp = ws.Cells(5, 1).Top + 3
    L = rEdge
    For i = 2 To 0 Step -1
        L = L - wd(i)
        If Not ShapeExists(ws, CStr(names(i))) Then
            Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, L, tp, wd(i), 18)
            shp.name = names(i)
            shp.Fill.ForeColor.RGB = RGB(91, 143, 214)
            shp.Line.ForeColor.RGB = RGB(255, 255, 255)
            shp.Line.Weight = 0.75
            shp.OnAction = macros(i)
            With shp.TextFrame2
                .VerticalAnchor = msoAnchorMiddle
                .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
                .TextRange.Text = caps(i)
                .TextRange.Font.Size = 9
                .TextRange.Font.Bold = msoTrue
                .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            End With
        Else
            Set shp = ws.Shapes(CStr(names(i)))
        End If
        shp.Left = L
        shp.Top = tp
        shp.Placement = xlMove
        On Error Resume Next
        shp.DrawingObject.PrintObject = False
        On Error GoTo 0
        L = L - 3
    Next i
End Sub

' ------------------------------------------------------------
' 09_배분결과(선택일) 기준으로 블록을 모두 다시 그림
' ------------------------------------------------------------
Public Sub Board_Refresh()
    Dim ws As Worksheet, wsR As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_DAY)
    Set wsR = ThisWorkbook.Sheets(SH_RESULT)
    FastBegin
    On Error GoTo bail

    If Not ShapeExists(ws, "Panel") Then Board_BuildChrome

    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).name, 4) = "Blk_" Then ws.Shapes(i).Delete
    Next i

    Dim names() As String, unCol As Long, nInsp As Long
    nInsp = GetBoardCols(names, unCol)
    Dim d As Date
    d = BoardDate()
    Dim last As Long
    last = wsR.Cells(wsR.rows.Count, R_NAME).End(xlUp).Row
    Dim nTask As Long, nLunch As Long, nFail As Long, nManual As Long, nHidden As Long
    Dim winBot As Long
    winBot = GRID_START_ROW + BoardWinRows() - 1

    If last >= 2 Then
        Dim v As Variant, dc() As Variant
        v = wsR.Range(wsR.Cells(2, 1), wsR.Cells(last, R_COLS)).Value
        ReDim dc(1 To UBound(v, 1), 1 To 2)
        Dim insp As String, col As Long, row1 As Long, nS As Long, slot As Long, typ As Long
        Dim unOffset As Long, rg As Range, shp As Shape, k As Long, r As Long, dr1 As Long, dr2 As Long, show As Boolean
        For r = 1 To UBound(v, 1)
            dc(r, 1) = v(r, R_DCOL)
            dc(r, 2) = v(r, R_DROW)
            If IsDate(v(r, R_DATE)) Then
                If DateValue(v(r, R_DATE)) = d Then
                    insp = CStr(v(r, R_INSP))
                    nS = CLng(v(r, R_NSLOT))
                    If nS < 1 Then nS = 1
                    slot = CLng(v(r, R_SLOT))
                    typ = CLng(v(r, R_TYPE))
                    Select Case typ
                        Case 5: nLunch = nLunch + 1
                        Case 0: nFail = nFail + 1
                        Case Else: nTask = nTask + 1
                    End Select
                    If CStr(v(r, R_MANUAL)) = "Y" Then nManual = nManual + 1

                    col = 0
                    If insp <> "" And slot > 0 Then
                        For k = 1 To nInsp
                            If names(k) = insp Then col = GRID_COL0 + k: Exit For
                        Next k
                    End If
                    If col = 0 Then
                        col = unCol                       ' 미지정 열: 창 맨 위부터 쌓기 (스크롤과 무관)
                        dr1 = GRID_START_ROW + unOffset
                        dr2 = dr1 + nS - 1
                        unOffset = unOffset + nS
                        show = (dr1 <= winBot)
                        If dr2 > winBot Then dr2 = winBot
                    Else
                        show = VisibleRows(slot, nS, dr1, dr2)   ' 창 밖 구간은 잘라서 표시
                    End If
                    dc(r, 1) = 0
                    dc(r, 2) = 0
                    If Not show Then
                        nHidden = nHidden + 1
                        GoTo nextRow
                    End If
                    row1 = dr1
                    Set rg = ws.Range(ws.Cells(dr1, col), ws.Cells(dr2, col))
                    Set shp = ws.Shapes.AddShape(msoShapeRectangle, rg.Left + 1, rg.Top + 1, rg.Width - 2, rg.Height - 2)
                    shp.name = "Blk_" & (r + 1)
                    shp.Fill.ForeColor.RGB = LegendColor(LegendIdx(typ))
                    If typ = 0 Then
                        shp.Line.ForeColor.RGB = RGB(255, 255, 255)
                        shp.Line.DashStyle = msoLineDash
                        shp.Line.Weight = 1.25
                    Else
                        shp.Line.Visible = msoFalse
                    End If
                    shp.Placement = xlMove
                    shp.OnAction = "Board_BlockClick"
                    With shp.TextFrame2
                        .WordWrap = msoTrue
                        .VerticalAnchor = msoAnchorTop
                        .MarginLeft = 3: .MarginTop = 1: .MarginRight = 1: .MarginBottom = 0
                        .TextRange.Text = BlockText(v, r, nS, typ)
                        .TextRange.Font.name = "맑은 고딕"
                        .TextRange.Font.Size = 7.5
                        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                        .TextRange.ParagraphFormat.Alignment = msoAlignLeft
                        .TextRange.Paragraphs(1).Font.Bold = msoTrue
                    End With
                    dc(r, 1) = col        ' 표시 열/행(참고) - 드래그 이동 감지에 사용
                    dc(r, 2) = row1
                End If
            End If
nextRow:
        Next r
        wsR.Range(wsR.Cells(2, R_DCOL), wsR.Cells(last, R_DROW)).Value = dc
    End If

    ' 제목/요약
    Board_UpdateHeader
    Dim body As String
    If nTask + nLunch + nFail = 0 Then
        body = "이 날짜의 배분 결과가 없습니다. 오늘 이후 날짜는 열 때 자동으로 배분되고, 지난 날짜는 [자동배분] 버튼을 누르세요." & vbCr & _
               "휴가/교육 메모(사용전검사 칸에 '한태희휴가')가 있는 검사자는 열에서 자동으로 빠집니다."
    Else
        body = "배치 " & nTask & "건 · 점심 " & nLunch & "건 · 배치 실패 " & nFail & "건 · 수동조정 " & nManual & "건" & vbCr & _
               "블록을 클릭하면 상세, 끌어서 다른 검사자/시간으로 옮길 수 있습니다. 조정이 끝나면 [입력 완료]를 눌러 10_원데이터에 저장하세요." & vbCr & _
               "겹치거나 하루(" & TimeLabel(DispStartTime()) & "~" & TimeLabel(DispEndTime()) & ")를 벗어나면 원위치되고 07_배분로그에 기록됩니다. 연한 회색 칸 = 근무시간 밖."
        If nHidden > 0 Then body = body & vbCr & "※ 지금 보이는 시간 구간 밖에 업무 " & nHidden & "건이 있습니다. 현황판 제목줄의 [▲][▼] 버튼으로 시간대를 옮겨 보세요."
        If ResultsStale(d) Then body = body & vbCr & "※ 배분 이후 화면의 업무가 바뀌었습니다. [자동배분]으로 다시 배분하세요."
    End If
    SetPanel "선택일 요약 - " & Format(d, "yyyy-mm-dd"), body
bail:
    Dim errMsg As String
    If Err.Number <> 0 Then errMsg = Err.Description
    FastEnd
    If errMsg <> "" Then Say "현황판을 그리는 중 오류가 발생했습니다: " & errMsg, vbCritical
End Sub

' 현황판 제목 막대(5행) + 날짜(2행 O열)
Private Sub Board_UpdateHeader()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_DAY)
    Dim d As Date
    d = BoardDate()
    ws.Cells(5, GRID_COL0).Value = "KESCO 검사 업무 배분 현황판 - " & Format(d, "yyyy년 m월 d일") & " (" & WeekdayKo(d) & ")"
    With ws.Range(ws.Cells(2, GRID_COL0), ws.Cells(2, GRID_COL0 + 3))
        .UnMerge
        .Merge
        .Value = Format(d, "m월 d일") & "  (" & WeekdayKo(d) & ")"
        .Font.name = "맑은 고딕"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(31, 56, 100)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    ' 4행 왼쪽: 지금 보이는 시간 구간 (범례 왼쪽 빈 칸)
    Dim names() As String, unCol As Long, sc As Long
    GetBoardCols names, unCol
    sc = unCol - 6
    If sc >= GRID_COL0 + 3 Then
        Dim topS As Long, nS As Long, a As String, b As String
        topS = BoardTopSlot()
        nS = DispSlotCount()
        a = SlotIndexToTimeStr(topS, DispStartTime(), SettingLong("슬롯단위_분"))
        If topS + BoardWinRows() - 1 >= nS + 1 Then
            b = TimeLabel(DispEndTime())
        Else
            b = SlotIndexToTimeStr(topS + BoardWinRows(), DispStartTime(), SettingLong("슬롯단위_분"))
        End If
        With ws.Range(ws.Cells(4, GRID_COL0), ws.Cells(4, sc - 1))
            .UnMerge
            .Merge
            .Value = "표시 시간 " & a & " ~ " & b & "   (제목줄 ▲▼로 시간대 이동)"
            .Font.name = "맑은 고딕"
            .Font.Size = 9
            .Font.Bold = False
            .Font.Color = RGB(127, 127, 127)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
    End If
End Sub

Private Function LegendColor(ByVal idx As Long) As Long
    Select Case idx
        Case 1: LegendColor = RGB(95, 99, 185)
        Case 2: LegendColor = RGB(0, 176, 80)
        Case 3: LegendColor = RGB(17, 85, 204)
        Case 4: LegendColor = RGB(207, 158, 0)
        Case 5: LegendColor = RGB(63, 127, 133)
        Case 6: LegendColor = RGB(112, 48, 160)
        Case Else: LegendColor = RGB(160, 40, 40)
    End Select
End Function

Private Function BlockText(v As Variant, ByVal r As Long, ByVal nS As Long, ByVal typ As Long) As String
    Dim nm As String, mark As String
    nm = CStr(v(r, R_NAME))
    If CStr(v(r, R_MANUAL)) = "Y" Then mark = ChrW(9998) & " "
    If typ = 5 Or nS <= 1 Then
        BlockText = mark & nm
        Exit Function
    End If
    Dim tstr As String
    tstr = CStr(v(r, R_TIME))
    BlockText = mark & nm & vbCr & CStr(v(r, R_CAP)) & "kW" & ChrW(183) & CStr(v(r, R_DUR)) & "분"
    If tstr <> "" Then BlockText = BlockText & vbCr & "약속 " & tstr
    If typ = 0 Then BlockText = BlockText & vbCr & "※ 배치 실패 - 07_배분로그 확인"
End Function

Private Sub SetPanel(ByVal title As String, ByVal body As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_DAY)
    If Not ShapeExists(ws, "Panel") Then Exit Sub
    With ws.Shapes("Panel").TextFrame2.TextRange
        .Text = title & vbCr & body
        .Font.Size = 8
        .Font.Bold = msoFalse
        .Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        With .Paragraphs(1).Font
            .Size = 9
            .Bold = msoTrue
            .Fill.ForeColor.RGB = RGB(31, 56, 100)
        End With
    End With
End Sub

Public Sub Board_ShowDetail(ByVal r As Long)
    Dim wsR As Worksheet, ws As Worksheet
    Set wsR = ThisWorkbook.Sheets(SH_RESULT)
    Set ws = ThisWorkbook.Sheets(SH_DAY)
    Dim typ As Long, nm As String, insp As String, slot As Long, nS As Long
    nm = CStr(wsR.Cells(r, R_NAME).Value)
    typ = CLng(wsR.Cells(r, R_TYPE).Value)
    insp = CStr(wsR.Cells(r, R_INSP).Value)
    slot = CLng(wsR.Cells(r, R_SLOT).Value)
    nS = CLng(wsR.Cells(r, R_NSLOT).Value)

    Dim whenTxt As String
    If slot > 0 And insp <> "" Then
        whenTxt = insp & "  " & CStr(wsR.Cells(r, R_START).Value) & "~" & _
                  SlotIndexToTimeStr(slot + nS, DispStartTime(), SettingLong("슬롯단위_분"))
    Else
        whenTxt = "미지정 열 (배치되지 않음)"
    End If

    Dim body As String
    If typ = 5 Then
        body = "배정: " & whenTxt & vbCr & "유형: " & TypeLabel(typ)
    Else
        Dim key As Long, scrRow As Long, jung As String, online As String
        key = CLng(wsR.Cells(r, R_ID).Value)
        scrRow = key \ 10
        If scrRow >= SCR_FIRST_ROW Then
            If key Mod 10 = 0 Then
                jung = Trim(CStr(ws.Cells(scrRow, 7).Value))
                If InStr(CStr(ws.Cells(scrRow, SCR_NOTE_L).Value), "온라인") > 0 Then online = "Y"
            Else
                jung = Trim(CStr(ws.Cells(scrRow, 12).Value))
                If InStr(CStr(ws.Cells(scrRow, SCR_NOTE_R).Value), "온라인") > 0 Then online = "Y"
            End If
        End If
        body = "용량 " & CStr(wsR.Cells(r, R_CAP).Value) & "kW · 지역 " & CStr(wsR.Cells(r, R_REGION).Value) & _
               " · 소요 " & CStr(wsR.Cells(r, R_DUR).Value) & "분 (" & nS & "칸)" & vbCr & _
               "배정: " & whenTxt & vbCr & _
               "화면 검사자: " & IIf(jung = "", "-", jung) & " · 약속시간: " & IIf(CStr(wsR.Cells(r, R_TIME).Value) = "", "-", CStr(wsR.Cells(r, R_TIME).Value)) & _
               " · 온라인: " & IIf(online = "", "N", online) & vbCr & _
               "유형: " & TypeLabel(typ) & " · 수수료(참고): " & CStr(wsR.Cells(r, R_FEE).Value) & _
               IIf(CStr(wsR.Cells(r, R_MANUAL).Value) = "Y", " · 수동조정됨", "")
    End If
    SetPanel "배분 상세 - " & nm, body
End Sub

' ------------------------------------------------------------
' 블록 클릭 (OnAction)
' ------------------------------------------------------------
Public Sub Board_BlockClick()
    Dim nm As String
    On Error GoTo bail
    nm = CStr(Application.Caller)
    If Left$(nm, 4) = "Blk_" Then Board_ShowDetail CLng(Mid$(nm, 5))
    Board_StartPoll
bail:
End Sub

' ------------------------------------------------------------
' 이동 감지: 블록이 표시 위치에서 벗어났는지 확인하고 처리
' ------------------------------------------------------------
Public Sub Board_CheckMoves()
    If busy Then Exit Sub
    Dim ws As Worksheet, wsR As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_DAY)
    Set wsR = ThisWorkbook.Sheets(SH_RESULT)
    busy = True
    On Error GoTo bail

    Dim shp As Shape, r As Long, c As Long, row1 As Long, nS As Long
    Dim rg As Range
    For Each shp In ws.Shapes
        If Left$(shp.name, 4) = "Blk_" Then
            r = CLng(Mid$(shp.name, 5))
            c = CLng(wsR.Cells(r, R_DCOL).Value)
            row1 = CLng(wsR.Cells(r, R_DROW).Value)
            nS = CLng(wsR.Cells(r, R_NSLOT).Value)
            If nS < 1 Then nS = 1
            If c > 0 And row1 > 0 Then
                Set rg = ws.Range(ws.Cells(row1, c), ws.Cells(ShownLastRow(r, row1, nS), c))
                If Abs(shp.Left - (rg.Left + 1)) > 2 Or Abs(shp.Top - (rg.Top + 1)) > 2 Then
                    HandleMove shp, r
                    Exit For        ' 처리 후 블록이 다시 그려졌으므로 중단
                End If
            End If
        End If
    Next shp
bail:
    busy = False
End Sub

' 표시된 블록의 마지막 행 (창 밖으로 잘린 블록은 창 맨 아래까지)
Private Function ShownLastRow(ByVal r As Long, ByVal row1 As Long, ByVal nS As Long) As Long
    Dim winBot As Long, e As Long
    winBot = GRID_START_ROW + BoardWinRows() - 1
    e = row1 + nS - 1
    If CLng(ThisWorkbook.Sheets(SH_RESULT).Cells(r, R_SLOT).Value) > 0 Then
        e = SlotRow(CLng(ThisWorkbook.Sheets(SH_RESULT).Cells(r, R_SLOT).Value)) + nS - 1
    End If
    If e > winBot Then e = winBot
    If e < row1 Then e = row1
    ShownLastRow = e
End Function

Private Sub HandleMove(shp As Shape, ByVal r As Long)
    Dim ws As Worksheet, wsR As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_DAY)
    Set wsR = ThisWorkbook.Sheets(SH_RESULT)

    Dim names() As String, unCol As Long, nInsp As Long
    nInsp = GetBoardCols(names, unCol)

    Dim startTime As Date, endTime As Date, slotMin As Long, nSlots As Long
    startTime = DispStartTime()
    endTime = DispEndTime()
    slotMin = SettingLong("슬롯단위_분")
    nSlots = CalcSlotCount(startTime, endTime, slotMin)

    Dim nm As String, oldInsp As String, oldSlot As Long, oldTyp As Long, nS As Long
    nm = CStr(wsR.Cells(r, R_NAME).Value)
    oldInsp = CStr(wsR.Cells(r, R_INSP).Value)
    oldSlot = CLng(wsR.Cells(r, R_SLOT).Value)
    oldTyp = CLng(wsR.Cells(r, R_TYPE).Value)
    nS = CLng(wsR.Cells(r, R_NSLOT).Value)
    If nS < 1 Then nS = 1

    ' 놓은 위치의 열(검사자) 찾기: 블록 가운데 x 기준
    Dim cx As Double, c As Long, tc As Long
    cx = shp.Left + shp.Width / 2
    For c = GRID_COL0 + 1 To unCol
        With ws.Cells(GRID_START_ROW, c)
            If cx >= .Left And cx < .Left + .Width Then
                tc = c
                Exit For
            End If
        End With
    Next c

    Dim reason As String, newInsp As String, newSlot As Long
    Dim d As Date: d = BoardDate()

    If tc = 0 Then
        reason = "검사자 열 밖에 놓았습니다."
    ElseIf tc = unCol Then
        If oldTyp = 5 Then
            reason = "점심시간은 '미지정' 열로 옮길 수 없습니다."
        Else
            newInsp = ""
            newSlot = 0
        End If
    Else
        newInsp = names(tc - GRID_COL0)
        Dim winTopY As Double, winBotY As Double, cy As Double
        winTopY = ws.Cells(GRID_START_ROW, GRID_COL0).Top
        winBotY = ws.Cells(GRID_START_ROW + BoardWinRows(), GRID_COL0).Top
        cy = shp.Top + shp.Height / 2
        With ws.Cells(GRID_START_ROW, GRID_COL0)
            If oldSlot > 0 And CLng(wsR.Cells(r, R_DROW).Value) > 0 Then
                ' 보이는 위치에서 옮긴 칸 수만큼 시작 슬롯을 옮김 (창 위/아래로 잘린 블록도 정확)
                newSlot = oldSlot + CLng(Round((shp.Top - (ws.Cells(CLng(wsR.Cells(r, R_DROW).Value), GRID_COL0).Top + 1)) / .Height, 0))
            Else
                newSlot = CLng(Round((shp.Top - .Top) / .Height, 0)) + BoardTopSlot()
            End If
        End With
        If cy < winTopY Or cy > winBotY Then
            reason = "현황판의 보이는 시간 구간 밖에 놓았습니다."
        ElseIf newSlot < 1 Or newSlot + nS - 1 > nSlots Then
            reason = "시간축(" & TimeLabel(startTime) & "~" & TimeLabel(endTime) & ") 밖입니다."
        Else
            ' 같은 날짜·같은 검사자의 다른 블록(점심 포함)과 겹침 검사
            Dim last As Long, rr As Long, s2 As Long, n2 As Long, v As Variant
            last = wsR.Cells(wsR.rows.Count, R_NAME).End(xlUp).Row
            v = wsR.Range(wsR.Cells(2, 1), wsR.Cells(last, R_COLS)).Value
            For rr = 2 To last
                If rr <> r And CStr(v(rr - 1, R_INSP)) = newInsp Then
                    If IsDate(v(rr - 1, R_DATE)) Then
                        If DateValue(v(rr - 1, R_DATE)) = d Then
                            s2 = CLng(v(rr - 1, R_SLOT))
                            n2 = CLng(v(rr - 1, R_NSLOT))
                            If s2 > 0 Then
                                If newSlot <= s2 + n2 - 1 And s2 <= newSlot + nS - 1 Then
                                    reason = "'" & newInsp & "' 검사자의 " & SlotIndexToTimeStr(newSlot, startTime, slotMin) & _
                                             " 시간대가 '" & CStr(v(rr - 1, R_NAME)) & "' 업무와 겹칩니다."
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                End If
            Next rr
        End If
    End If

    If reason <> "" Then
        AddLog "충돌", "수동 이동 취소: " & reason, nm
        Board_Refresh                       ' 원위치
        Say "이동할 수 없습니다. " & reason & vbCrLf & "원래 위치로 되돌렸습니다.", vbExclamation
        Exit Sub
    End If

    If newInsp = oldInsp And newSlot = oldSlot Then
        Board_Refresh
        Exit Sub
    End If

    Dim fromTxt As String, toTxt As String
    If oldInsp = "" Then fromTxt = "미지정 열" Else fromTxt = oldInsp & " " & CStr(wsR.Cells(r, R_START).Value)
    If newInsp = "" Then
        toTxt = "미지정 열"
    Else
        toTxt = newInsp & " " & SlotIndexToTimeStr(newSlot, startTime, slotMin)
    End If

    wsR.Cells(r, R_INSP).Value = newInsp
    wsR.Cells(r, R_SLOT).Value = newSlot
    wsR.Cells(r, R_START).NumberFormat = "@"
    If newSlot > 0 Then wsR.Cells(r, R_START).Value = SlotIndexToTimeStr(newSlot, startTime, slotMin) Else wsR.Cells(r, R_START).Value = ""
    If newInsp = "" Then
        wsR.Cells(r, R_TYPE).Value = 0
    ElseIf oldTyp = 0 Then
        wsR.Cells(r, R_TYPE).Value = 7      ' 배치 실패 -> 수동 배정
    End If
    wsR.Cells(r, R_MANUAL).Value = "Y"

    AddLog "수동조정", "블록 이동: " & fromTxt & " -> " & toTxt, nm
    Board_Refresh
    Board_ShowDetail r
End Sub

' 블록이 선택돼 있는 동안 1초마다 이동 여부 확인 (드래그 직후 즉시 반영)
Public Sub Board_StartPoll()
    If pollOn Then Exit Sub
    pollOn = True
    pollNext = Now + TimeSerial(0, 0, 1)
    Application.OnTime pollNext, "Board_Poll"
End Sub

Public Sub Board_Poll()
    pollOn = False
    On Error Resume Next
    If ActiveSheet.name <> SH_DAY Then Exit Sub
    On Error GoTo 0
    Board_CheckMoves
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Board_StartPoll
End Sub

Public Sub Board_StopPoll()
    On Error Resume Next
    If pollOn Then Application.OnTime pollNext, "Board_Poll", , False
    pollOn = False
End Sub
