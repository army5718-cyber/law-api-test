Attribute VB_Name = "Module4_현황판"
Option Explicit

' ============================================================
' KESCO 부천김포지사 검사업무 ERP
' Module4_현황판 : 11_일일대장 시트 오른쪽(N열~)에 놓인 배분현황판
'   - 09_배분결과 표를 읽어 현황판 위에 "업무 블록(도형)"을 그림
'   - 블록 클릭 -> 상단 상세 패널 표시
'   - 블록 드래그 -> 검사자 열/시간으로 환산, 겹침·운영시간 검사 후 저장
'     (실패하면 자동으로 고치지 않고 원위치 + 07_배분로그 '충돌' 기록)
'   - 날짜 이동 / 달력은 일일대장 버튼(Module5_ERP)과 공용입니다.
' 블록 이름 규칙: Blk_<09_배분결과 행번호>
' ============================================================

Private pollOn As Boolean
Private pollNext As Date
Private busy As Boolean

' ------------------------------------------------------------
' 달력 팝업(frmCalendar)을 띄워 날짜를 고름. 선택하면 True.
' allowClear=True 이면 '지움' 버튼이 보이고, 눌렀을 때 cleared=True 로 반환.
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

' ------------------------------------------------------------
' 현황판 열 구성 읽기: 머리글(6행)에서 검사자 이름과 '미지정' 열 위치를 구함
' ------------------------------------------------------------
Private Function GetBoardCols(names() As String, ByRef unCol As Long) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_GRID)
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

' 일일대장이 그려진 현황판의 마지막(미지정) 열 번호 (인쇄 영역 계산용)
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

' 유형 -> 범례 칸 번호(1~7, 5행 O열부터)
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
' 현황판 격자를 선택일 근무자 기준으로 다시 그림
' (머리글 / 시간 라벨 / 점선 격자 / 제목줄·범례·상세패널) 후 블록을 다시 표시.
' 11_일일대장의 A~L열(접수대장)은 건드리지 않습니다.
' ------------------------------------------------------------
Public Sub RebuildGrid(ByVal silent As Boolean)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_GRID)

    Dim prevEvents As Boolean, prevSU As Boolean
    prevEvents = Application.EnableEvents
    prevSU = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim inspectors() As String
    inspectors = GetActiveInspectorsFor(BoardDate())
    Dim n As Long
    If HasInspectors(inspectors) Then n = UBound(inspectors) - LBound(inspectors) + 1 Else n = 0

    ' 현황판 영역 초기화 (N열 이후만)
    With ws.Range(ws.Cells(4, GRID_COL0), ws.Cells(GRID_START_ROW + 30, GRID_COL0 + 40))
        .UnMerge
        .Clear
        .Font.name = "맑은 고딕"
        .Font.Size = 9
        .VerticalAlignment = xlTop
    End With

    Dim navy As Long, paleBlue As Long, dotBlue As Long
    navy = RGB(31, 56, 100)
    paleBlue = RGB(221, 235, 247)
    dotBlue = RGB(91, 155, 213)

    Dim c As Long
    ws.Columns(GRID_COL0).ColumnWidth = 5.5
    FormatHeader ws.Cells(GRID_HEADER_ROW, GRID_COL0), "시간", paleBlue, navy
    For c = 1 To n
        FormatHeader ws.Cells(GRID_HEADER_ROW, GRID_COL0 + c), inspectors(c), paleBlue, navy
        ws.Columns(GRID_COL0 + c).ColumnWidth = 8.6
    Next c
    Dim unCol As Long
    unCol = GRID_COL0 + n + 1
    FormatHeader ws.Cells(GRID_HEADER_ROW, unCol), "미지정", paleBlue, navy
    ws.Columns(unCol).ColumnWidth = 8.6
    ws.Columns(unCol + 1).ColumnWidth = 2
    ws.Cells(GRID_HEADER_ROW, GRID_COL0).Font.Size = 9

    ' 시간 슬롯 행 : 02_설정 기준으로 시간 라벨과 점선 격자를 그림
    Dim startTime As Date, endTime As Date, slotMin As Long, nSlots As Long
    startTime = SettingTime("운영시작시간")
    endTime = SettingTime("운영종료시간")
    slotMin = SettingLong("슬롯단위_분")
    nSlots = CalcSlotCount(startTime, endTime, slotMin)

    Dim s As Long, rowIdx As Long
    For s = 1 To nSlots
        rowIdx = GRID_START_ROW + s - 1
        ws.rows(rowIdx).RowHeight = 24
        With ws.Cells(rowIdx, GRID_COL0)
            .NumberFormat = "@"
            .Value = SlotIndexToTimeStr(s, startTime, slotMin)
            .Font.name = "맑은 고딕": .Font.Size = 9: .Font.Bold = True
            .Font.Color = navy
            .Interior.Color = paleBlue
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlTop
        End With
        With ws.Range(ws.Cells(rowIdx, GRID_COL0 + 1), ws.Cells(rowIdx, unCol))
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideVertical).Color = RGB(191, 207, 224)
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Color = RGB(191, 207, 224)
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Color = RGB(191, 207, 224)
            .Borders(xlEdgeBottom).LineStyle = xlDot            ' 30분 점선
            .Borders(xlEdgeBottom).Color = dotBlue
        End With
    Next s

    ' 마감 행 (배정 불가, 운영 종료선 표시)
    rowIdx = GRID_START_ROW + nSlots
    ws.rows(rowIdx).RowHeight = 24
    With ws.Cells(rowIdx, GRID_COL0)
        .NumberFormat = "@"
        .Value = Format(endTime, "hh:nn")
        .Font.name = "맑은 고딕": .Font.Size = 9: .Font.Bold = True
        .Font.Color = RGB(128, 128, 128)
        .Interior.Color = paleBlue
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlTop
    End With
    With ws.Range(ws.Cells(rowIdx, GRID_COL0 + 1), ws.Cells(rowIdx, unCol))
        .Interior.Color = RGB(242, 242, 242)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = navy
    End With
    ws.Cells(rowIdx, GRID_COL0 + 1).Value = "운영 종료"
    ws.Cells(rowIdx, GRID_COL0 + 1).Font.Color = RGB(128, 128, 128)
    ws.Cells(rowIdx, GRID_COL0 + 1).Font.Size = 8

    Board_BuildChrome      ' 제목줄/범례/상세 패널
    Board_Refresh          ' 09_배분결과 기준으로 블록 표시

    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevSU
End Sub

Private Sub FormatHeader(cell As Range, ByVal txt As String, ByVal fillColor As Long, ByVal fontColor As Long)
    With cell
        .Value = txt
        .Font.name = "맑은 고딕"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = fontColor
        .Interior.Color = fillColor
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ShrinkToFit = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeBottom).Color = fontColor
    End With
End Sub

' ------------------------------------------------------------
' 화면 구성요소(제목줄/범례/상세 패널) 생성 - 여러 번 호출해도 안전
' ------------------------------------------------------------
Public Sub Board_BuildChrome()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_GRID)
    Dim names() As String, unCol As Long, n As Long
    n = GetBoardCols(names, unCol)

    Dim navy As Long
    navy = RGB(31, 56, 100)

    ' 제목줄 (4행)
    With ws.Range(ws.Cells(4, GRID_COL0), ws.Cells(4, unCol))
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
    ws.Cells(4, GRID_COL0).Value = "KESCO 검사 업무 배분 현황판"

    ' 범례 (5행, O열부터 한 칸씩)
    Dim labels As Variant, colors As Variant, i As Long
    labels = Array("검사자+시간", "시간만", "검사자만", "미지정", "점심", "수동", "배치실패")
    colors = Array(RGB(95, 99, 185), RGB(0, 176, 80), RGB(17, 85, 204), RGB(207, 158, 0), _
                   RGB(63, 127, 133), RGB(112, 48, 160), RGB(160, 40, 40))
    For i = 0 To 6
        With ws.Cells(5, GRID_COL0 + 1 + i)
            .Value = labels(i)
            .Interior.Color = colors(i)
            .Font.name = "맑은 고딕"
            .Font.Size = 8
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Color = RGB(255, 255, 255)
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Color = RGB(255, 255, 255)
        End With
    Next i

    ' 상세 패널 (1~3행, 현황판 위쪽 - 인쇄 영역 밖)
    Dim shp As Shape
    Dim pLeft As Double, pTop As Double, pW As Double, pH As Double
    pLeft = ws.Cells(1, GRID_COL0).Left
    pTop = ws.Cells(1, 1).Top + 1
    pW = ws.Cells(1, unCol + 1).Left - pLeft
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

' ------------------------------------------------------------
' 09_배분결과(선택일) 기준으로 블록을 모두 다시 그림
' ------------------------------------------------------------
Public Sub Board_Refresh()
    Dim ws As Worksheet, wsR As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_GRID)
    Set wsR = ThisWorkbook.Sheets(SH_RESULT)

    Dim prevSU As Boolean, prevEv As Boolean
    prevSU = Application.ScreenUpdating
    prevEv = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
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
    Dim lastRow As Long, r As Long
    lastRow = wsR.Cells(wsR.rows.Count, R_NAME).End(xlUp).Row

    Dim insp As String, col As Long, row1 As Long, nS As Long, slot As Long, typ As Long
    Dim unOffset As Long, rg As Range, shp As Shape, txt As String, k As Long

    For r = 2 To lastRow
        If IsDate(wsR.Cells(r, R_DATE).Value) Then
            If DateValue(wsR.Cells(r, R_DATE).Value) = d Then
                insp = CStr(wsR.Cells(r, R_INSP).Value)
                nS = CLng(wsR.Cells(r, R_NSLOT).Value)
                If nS < 1 Then nS = 1
                slot = CLng(wsR.Cells(r, R_SLOT).Value)
                typ = CLng(wsR.Cells(r, R_TYPE).Value)

                col = 0
                If insp <> "" And slot > 0 Then
                    For k = 1 To nInsp
                        If names(k) = insp Then col = GRID_COL0 + k: Exit For
                    Next k
                End If
                If col = 0 Then
                    col = unCol                       ' 미지정 열에 위에서부터 쌓기
                    row1 = GRID_START_ROW + unOffset
                    unOffset = unOffset + nS
                Else
                    row1 = GRID_START_ROW + slot - 1
                End If

                Set rg = ws.Range(ws.Cells(row1, col), ws.Cells(row1 + nS - 1, col))
                Set shp = ws.Shapes.AddShape(msoShapeRectangle, rg.Left + 1, rg.Top + 1, rg.Width - 2, rg.Height - 2)
                shp.name = "Blk_" & r
                shp.Fill.ForeColor.RGB = ws.Cells(5, GRID_COL0 + LegendIdx(typ)).Interior.Color
                If typ = 0 Then
                    shp.Line.Visible = msoTrue
                    shp.Line.ForeColor.RGB = RGB(255, 255, 255)
                    shp.Line.DashStyle = msoLineDash
                    shp.Line.Weight = 1.25
                Else
                    shp.Line.Visible = msoFalse
                End If
                shp.Shadow.Visible = msoTrue
                shp.Shadow.OffsetX = 1
                shp.Shadow.OffsetY = 1.5
                shp.Shadow.Transparency = 0.65
                shp.Shadow.Blur = 2
                shp.Placement = xlMove
                shp.OnAction = "Board_BlockClick"

                txt = BlockText(wsR, r, nS, insp, typ)
                With shp.TextFrame2
                    .WordWrap = msoTrue
                    .VerticalAnchor = msoAnchorTop
                    .MarginLeft = 3: .MarginTop = 1: .MarginRight = 1: .MarginBottom = 0
                    .TextRange.Text = txt
                    .TextRange.Font.name = "맑은 고딕"
                    .TextRange.Font.Size = 7.5
                    .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                    .TextRange.ParagraphFormat.Alignment = msoAlignLeft
                    If Len(.TextRange.Paragraphs(1).Text) > 0 Then .TextRange.Paragraphs(1).Font.Bold = msoTrue
                End With
                wsR.Cells(r, R_NAME + 12).Value = col        ' 15열: 표시 열(참고)
                wsR.Cells(r, R_NAME + 13).Value = row1       ' 16열: 표시 행(참고)
            End If
        End If
    Next r

    Board_UpdateTitle
    Board_ShowSummary
bail:
    Dim errMsg As String
    If Err.Number <> 0 Then errMsg = Err.Description
    Application.EnableEvents = prevEv
    Application.ScreenUpdating = prevSU
    If errMsg <> "" Then Say "현황판을 그리는 중 오류가 발생했습니다: " & errMsg, vbCritical
End Sub

Private Function BlockText(wsR As Worksheet, ByVal r As Long, ByVal nS As Long, ByVal insp As String, ByVal typ As Long) As String
    Dim nm As String, mark As String
    nm = CStr(wsR.Cells(r, R_NAME).Value)
    If CStr(wsR.Cells(r, R_MANUAL).Value) = "Y" Then mark = ChrW(9998) & " "
    If typ = 5 Or nS <= 1 Then
        BlockText = mark & nm
        Exit Function
    End If
    Dim line2 As String, tstr As String
    line2 = CStr(wsR.Cells(r, R_CAP).Value) & "kW" & ChrW(183) & CStr(wsR.Cells(r, R_DUR).Value) & "분"
    tstr = CStr(wsR.Cells(r, R_TIME).Value)
    BlockText = mark & nm & vbCr & line2
    If tstr <> "" Then BlockText = BlockText & vbCr & "약속 " & tstr
    If typ = 0 Then BlockText = BlockText & vbCr & "※ 배치 실패 - 07_배분로그 확인"
End Function

Private Sub Board_UpdateTitle()
    Dim ws As Worksheet, d As Date
    Set ws = ThisWorkbook.Sheets(SH_GRID)
    d = BoardDate()
    On Error Resume Next
    ws.Cells(4, GRID_COL0).Value = "KESCO 검사 업무 배분 현황판 - " & Format(d, "yyyy년 m월 d일") & " (" & WeekdayKo(d) & ")"
    On Error GoTo 0
End Sub

' ------------------------------------------------------------
' 상세 패널
' ------------------------------------------------------------
Private Sub SetPanel(ByVal title As String, ByVal body As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_GRID)
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

Public Sub Board_ShowSummary()
    Dim wsR As Worksheet
    Set wsR = ThisWorkbook.Sheets(SH_RESULT)
    Dim d As Date: d = BoardDate()
    Dim lastRow As Long, r As Long
    Dim nTask As Long, nLunch As Long, nFail As Long, nManual As Long
    lastRow = wsR.Cells(wsR.rows.Count, R_NAME).End(xlUp).Row
    For r = 2 To lastRow
        If IsDate(wsR.Cells(r, R_DATE).Value) Then
            If DateValue(wsR.Cells(r, R_DATE).Value) = d Then
                Select Case CLng(wsR.Cells(r, R_TYPE).Value)
                    Case 5: nLunch = nLunch + 1
                    Case 0: nFail = nFail + 1
                    Case Else: nTask = nTask + 1
                End Select
                If CStr(wsR.Cells(r, R_MANUAL).Value) = "Y" Then nManual = nManual + 1
            End If
        End If
    Next r
    Dim body As String
    If nTask + nLunch + nFail = 0 Then
        body = "이 날짜의 배분 결과가 없습니다. 오늘 이후 날짜는 열 때 자동으로 배분되고, 지난 날짜는 [자동배분] 버튼을 누르세요." & vbCr & _
               "휴가/교육 메모가 있는 검사자는 열에서 자동으로 빠집니다."
    Else
        body = "배치 " & nTask & "건 · 점심 " & nLunch & "건 · 배치 실패 " & nFail & "건 · 수동조정 " & nManual & "건" & vbCr & _
               "블록을 클릭하면 상세가 표시되고, 끌어서 다른 검사자/시간으로 옮길 수 있습니다. 조정 후 [결과 반영]을 누르면 원데이터 검사자에 기록됩니다." & vbCr & _
               "겹치거나 운영시간(~" & Format(SettingTime("운영종료시간"), "hh:nn") & ")을 넘으면 원위치되고 07_배분로그에 기록됩니다."
        If ResultsStale(d) Then body = body & vbCr & "※ 이 날짜의 원데이터가 배분 이후 바뀌었습니다. [자동배분]으로 다시 배분하세요."
    End If
    SetPanel "선택일 요약 - " & Format(d, "yyyy-mm-dd"), body
End Sub

Public Sub Board_ShowDetail(ByVal r As Long)
    Dim wsR As Worksheet
    Set wsR = ThisWorkbook.Sheets(SH_RESULT)
    Dim typ As Long, nm As String, insp As String, slot As Long, nS As Long
    nm = CStr(wsR.Cells(r, R_NAME).Value)
    typ = CLng(wsR.Cells(r, R_TYPE).Value)
    insp = CStr(wsR.Cells(r, R_INSP).Value)
    slot = CLng(wsR.Cells(r, R_SLOT).Value)
    nS = CLng(wsR.Cells(r, R_NSLOT).Value)

    Dim whenTxt As String
    If slot > 0 And insp <> "" Then
        whenTxt = insp & "  " & CStr(wsR.Cells(r, R_START).Value) & "~" & _
                  SlotIndexToTimeStr(slot + nS, SettingTime("운영시작시간"), SettingLong("슬롯단위_분"))
    Else
        whenTxt = "미지정 열 (배치되지 않음)"
    End If

    Dim body As String
    If typ = 5 Then
        body = "배정: " & whenTxt & vbCr & "유형: " & TypeLabel(typ)
    Else
        Dim tid As Long, rawReg As String, jung As String, online As String, idx As Object, lo As ListObject
        tid = CLng(wsR.Cells(r, R_ID).Value)
        If tid > 0 Then
            Set idx = RawIdIndex()
            If idx.Exists(CStr(tid)) Then
                Set lo = RawTable()
                With lo.DataBodyRange
                    rawReg = CStr(.Cells(idx(CStr(tid)), RW_REG).Value)
                    jung = Trim(CStr(.Cells(idx(CStr(tid)), RW_INSP).Value))
                    If InStr(CStr(.Cells(idx(CStr(tid)), RW_NOTE).Value), "온라인") > 0 Then online = "Y"
                End With
            End If
        End If
        body = "용량 " & CStr(wsR.Cells(r, R_CAP).Value) & "kW · 지역 " & CStr(wsR.Cells(r, R_REGION).Value) & _
               IIf(rawReg <> "" And rawReg <> CStr(wsR.Cells(r, R_REGION).Value), " (입력: " & rawReg & ")", "") & _
               " · 소요 " & CStr(wsR.Cells(r, R_DUR).Value) & "분 (" & nS & "칸)" & vbCr & _
               "배정: " & whenTxt & vbCr & _
               "원데이터 검사자: " & IIf(jung = "", "-", jung) & " · 약속시간: " & IIf(CStr(wsR.Cells(r, R_TIME).Value) = "", "-", CStr(wsR.Cells(r, R_TIME).Value)) & _
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
' (시트 선택 변경 이벤트 + 블록 선택 중 1초 간격 확인에서 호출)
' ------------------------------------------------------------
Public Sub Board_CheckMoves()
    If busy Then Exit Sub
    Dim ws As Worksheet, wsR As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_GRID)
    Set wsR = ThisWorkbook.Sheets(SH_RESULT)
    busy = True
    On Error GoTo bail

    Dim shp As Shape, r As Long, c As Long, row1 As Long, nS As Long
    Dim rg As Range
    For Each shp In ws.Shapes
        If Left$(shp.name, 4) = "Blk_" Then
            r = CLng(Mid$(shp.name, 5))
            c = CLng(wsR.Cells(r, R_NAME + 12).Value)
            row1 = CLng(wsR.Cells(r, R_NAME + 13).Value)
            nS = CLng(wsR.Cells(r, R_NSLOT).Value)
            If nS < 1 Then nS = 1
            If c > 0 And row1 > 0 Then
                Set rg = ws.Range(ws.Cells(row1, c), ws.Cells(row1 + nS - 1, c))
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

Private Sub HandleMove(shp As Shape, ByVal r As Long)
    Dim ws As Worksheet, wsR As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_GRID)
    Set wsR = ThisWorkbook.Sheets(SH_RESULT)

    Dim names() As String, unCol As Long, nInsp As Long
    nInsp = GetBoardCols(names, unCol)

    Dim startTime As Date, endTime As Date, slotMin As Long, nSlots As Long
    startTime = SettingTime("운영시작시간")
    endTime = SettingTime("운영종료시간")
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
        With ws.Cells(GRID_START_ROW, GRID_COL0)
            newSlot = CLng(Round((shp.Top - .Top) / .Height, 0)) + 1
        End With
        If newSlot < 1 Or newSlot + nS - 1 > nSlots Then
            reason = "운영시간(" & Format(startTime, "hh:nn") & "~" & Format(endTime, "hh:nn") & ") 밖입니다."
        Else
            ' 같은 날짜·같은 검사자의 다른 블록(점심 포함)과 겹침 검사
            Dim lastRow As Long, rr As Long, s2 As Long, n2 As Long
            lastRow = wsR.Cells(wsR.rows.Count, R_NAME).End(xlUp).Row
            For rr = 2 To lastRow
                If rr <> r And CStr(wsR.Cells(rr, R_INSP).Value) = newInsp Then
                    If IsDate(wsR.Cells(rr, R_DATE).Value) Then
                        If DateValue(wsR.Cells(rr, R_DATE).Value) = d Then
                            s2 = CLng(wsR.Cells(rr, R_SLOT).Value)
                            n2 = CLng(wsR.Cells(rr, R_NSLOT).Value)
                            If s2 > 0 Then
                                If newSlot <= s2 + n2 - 1 And s2 <= newSlot + nS - 1 Then
                                    reason = "'" & newInsp & "' 검사자의 " & SlotIndexToTimeStr(newSlot, startTime, slotMin) & _
                                             " 시간대가 '" & CStr(wsR.Cells(rr, R_NAME).Value) & "' 업무와 겹칩니다."
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

    ' 제자리(스냅)이면 변경 없음
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

' ------------------------------------------------------------
' 블록이 선택돼 있는 동안 1초마다 이동 여부 확인 (드래그 직후 즉시 반영)
' ------------------------------------------------------------
Public Sub Board_StartPoll()
    If pollOn Then Exit Sub
    pollOn = True
    pollNext = Now + TimeSerial(0, 0, 1)
    Application.OnTime pollNext, "Board_Poll"
End Sub

Public Sub Board_Poll()
    pollOn = False
    On Error Resume Next
    If ActiveSheet.name <> SH_GRID Then Exit Sub
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
