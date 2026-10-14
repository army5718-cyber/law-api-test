Attribute VB_Name = "Module2_자동배분"
Option Explicit

' ============================================================
' KESCO 부천김포지사 검사업무 ERP
' Module2_자동배분 : 5단계 배분 알고리즘 본체 (계산만 담당)
'
' 배분 대상 = 11_일일대장에서 선택한 날짜의 10_원데이터 행 (구분이 정기검사/사용전검사인 행)
'   - 03_업무입력 시트는 없어졌습니다. 원데이터가 곧 업무 목록입니다.
'   - 원데이터의 검사자/시간 칸이 채워져 있으면 그것을 "지정"으로 보고 우선 배치합니다.
'
' 순서 고정 (원본 요구사항과 동일):
'   0. 사전준비  1순위: 검사자+시간 모두 지정  2순위: 시간만 지정
'   3순위: 검사자만 지정  4순위: 점심시간  5순위: 미지정 업무
'
' 결과는 09_배분결과 표에 기록하고, 화면 표시는 Module4_현황판(Board_Refresh)이 담당합니다.
'
' ※ 2/5순위의 "인근지역 우선순위가 가장 높은 검사자" 판정 기준:
'    각 검사자의 "가장 최근에 배정된 업무의 지역"과 신규 업무 지역 사이의
'    05_우선순위매트릭스 순위를 거리로 사용합니다 (근사치).
'    팀 운영 방식과 다르면 RankInspectorsByProximity 함수만 조정하세요.
'
' ※ 자동 침묵 수정 금지: 배치하지 못한 업무는 07_배분로그에 기록하고
'    '미지정' 열에 배치 실패 블록으로 표시합니다.
' ============================================================

Private occupied() As Long      ' occupied(검사자, 슬롯) = 업무 번호 (0=빈칸, -1=점심)
Private lastRegion() As String  ' lastRegion(검사자) = 최근 배정 업무의 표준지역
Private taskCount() As Long     ' taskCount(검사자) = 배정 건수 (동일 거리일 때 공정배분용)
Private placedTotal As Long
Private unassignedTotal As Long

Private mInsp() As String       ' 이번 실행의 검사자 이름
Private mDate As Date           ' 이번 실행의 배분 날짜
Private mStart As Date
Private mSlotMin As Long
Private wsRes As Worksheet
Private mTasks() As TaskRec

' ------------------------------------------------------------
' 버튼용: 선택한 날짜를 (다시) 자동배분하고 화면을 새로 그림
' ------------------------------------------------------------
Sub RunAutoAssign()
    Dim d As Date
    d = BoardDate()
    If CountManual(d) > 0 Then
        If Not Ask(Format(d, "yyyy-mm-dd") & " 에 수동으로 옮긴 블록이 " & CountManual(d) & "건 있습니다." & vbCrLf & _
                   "다시 자동배분하면 수동 조정은 사라집니다. 계속할까요?") Then Exit Sub
    End If
    Dim ok As Boolean
    ok = AutoAssignDate(d, False)
    Day_Render
    If ok Then
        Say "자동 배분이 완료되었습니다 (" & Format(d, "yyyy-mm-dd") & ")." & vbCrLf & _
            "배치 " & placedTotal & "건 / 미배정 " & unassignedTotal & "건 / 로그 " & gLogCount & "줄" & vbCrLf & _
            "충돌·미배정 항목은 07_배분로그에서 확인하세요." & vbCrLf & _
            "블록을 끌어 조정한 뒤 [결과 반영]을 누르면 검사자가 10_원데이터에 기록됩니다.", _
            IIf(unassignedTotal > 0, vbExclamation, vbInformation)
    End If
End Sub

' 선택일의 09_배분결과 중 수동 조정된 건수
Public Function CountManual(ByVal d As Date) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_RESULT)
    Dim last As Long, r As Long, n As Long
    last = ws.Cells(ws.rows.Count, R_NAME).End(xlUp).Row
    For r = 2 To last
        If IsDate(ws.Cells(r, R_DATE).Value) Then
            If DateValue(ws.Cells(r, R_DATE).Value) = d Then
                If CStr(ws.Cells(r, R_MANUAL).Value) = "Y" Then n = n + 1
            End If
        End If
    Next r
    CountManual = n
End Function

' 선택일의 09_배분결과 행 수
Public Function CountResults(ByVal d As Date) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_RESULT)
    Dim last As Long, r As Long, n As Long
    last = ws.Cells(ws.rows.Count, R_NAME).End(xlUp).Row
    For r = 2 To last
        If IsDate(ws.Cells(r, R_DATE).Value) Then
            If DateValue(ws.Cells(r, R_DATE).Value) = d Then n = n + 1
        End If
    Next r
    CountResults = n
End Function

' ------------------------------------------------------------
' 10_원데이터에서 선택일의 배분 대상 업무를 읽어 TaskRec 배열로 만든다.
' 반환값 = 업무 건수 (배열은 1..건수)
' ------------------------------------------------------------
Public Function LoadTasksForDate(ByVal d As Date, tk() As TaskRec) As Long
    Dim v As Variant
    v = RawData()
    If IsEmpty(v) Then
        LoadTasksForDate = 0
        Exit Function
    End If

    Dim stdMap As Object, zoneMap As Object, aliasMap As Object
    LoadRegionMaps stdMap, zoneMap
    Set aliasMap = LoadAliasMap()

    Dim slotMin As Long, bigKw As Double, bigMin As Long, smallMin As Long, onMin As Long
    slotMin = SettingLong("슬롯단위_분")
    bigKw = CDbl(SettingLong("정식설비기준용량_kW"))
    bigMin = SettingLong("정식설비소요시간_분")
    smallMin = SettingLong("약식설비소요시간_분")
    onMin = SettingLong("온라인업무소요시간_분")

    Dim n As Long, i As Long, key As Long, reg As String, tm As Variant
    key = CLng(Int(CDbl(d)))
    ReDim tk(1 To UBound(v, 1))
    For i = 1 To UBound(v, 1)
        If DayKey(v(i, RW_DATE)) = key Then
            If IsTaskKind(CStr(v(i, RW_KIND))) Then
                n = n + 1
                With tk(n)
                    .ID = 0
                    If IsNumeric(v(i, RW_ID)) And Not IsEmpty(v(i, RW_ID)) Then .ID = CLng(v(i, RW_ID))
                    .Kind = Trim$(CStr(v(i, RW_KIND)))
                    .TaskName = Trim$(CStr(v(i, RW_NAME)))
                    .CapText = Trim$(CStr(v(i, RW_CAP)))
                    .Cap = CapNum(v(i, RW_CAP))
                    reg = Trim$(CStr(v(i, RW_REG)))
                    .RegionRaw = reg
                    If stdMap.Exists(reg) Then .RegionStd = stdMap(reg) Else .RegionStd = ""
                    If Left$(.RegionStd, 1) = "(" Then .RegionStd = ""
                    tm = v(i, RW_TIME)
                    .TimeStr = ""
                    If Not IsError(tm) Then
                        If VarType(tm) = vbDate Or VarType(tm) = vbDouble Then
                            .TimeStr = CellTimeStr(tm)
                        ElseIf VarType(tm) = vbString Then
                            If IsDate(tm) Then .TimeStr = Format(TimeValue(CStr(tm)), "hh:nn")
                        End If
                    End If
                    If Not (.TimeStr Like "##:##") Then .TimeStr = ""
                    .InspRaw = Trim$(CStr(v(i, RW_INSP)))
                    .Insp = LeadInsp(.InspRaw, aliasMap)
                    .Online = (InStr(CStr(v(i, RW_NOTE)), "온라인") > 0)
                    If IsNumeric(v(i, RW_FEE)) And Not IsEmpty(v(i, RW_FEE)) Then .Fee = CStr(v(i, RW_FEE)) Else .Fee = ""
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
        End If
    Next i
    If n = 0 Then
        ReDim tk(1 To 1)
    Else
        ReDim Preserve tk(1 To n)
    End If
    LoadTasksForDate = n
End Function

' ------------------------------------------------------------
' 선택일을 자동배분해 09_배분결과에 저장 (화면은 그리지 않음)
' quiet=True 이면 안내/요약 로그를 남기지 않음 (일일대장을 열 때 자동 실행하는 경우)
' ------------------------------------------------------------
Public Function AutoAssignDate(ByVal d As Date, ByVal quiet As Boolean) As Boolean
    Dim prevEvents As Boolean, prevCalc As XlCalculation, prevSU As Boolean
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation
    prevSU = Application.ScreenUpdating
    Dim msgText As String

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    On Error GoTo fail

    ' ---------- 0. 사전준비 ----------
    Dim inspectors() As String
    inspectors = GetActiveInspectorsFor(d)
    If Not HasInspectors(inspectors) Then
        msgText = "배분할 검사자가 없습니다. 01_검사자명단의 오늘근무 값을 확인하세요."
        GoTo done
    End If
    Dim nInsp As Long
    nInsp = UBound(inspectors) - LBound(inspectors) + 1
    mInsp = inspectors

    ' 시간축은 24시간(표시시간), 약속시간이 없는 업무/점심은 근무시간(workLo~workHi 칸) 안에만 배치
    Dim startTime As Date, endTime As Date, slotMin As Long, nSlots As Long, workLo As Long, workHi As Long
    startTime = DispStartTime()
    endTime = DispEndTime()
    slotMin = SettingLong("슬롯단위_분")
    If slotMin <= 0 Then Err.Raise vbObjectError + 1, , "02_설정의 슬롯단위_분이 올바르지 않습니다."
    nSlots = CalcSlotCount(startTime, endTime, slotMin)
    If nSlots < 1 Then Err.Raise vbObjectError + 2, , "02_설정의 표시시간(또는 운영시간)이 올바르지 않습니다."
    workLo = WorkFirstSlot()
    workHi = WorkLastSlot()
    mStart = startTime
    mSlotMin = slotMin
    mDate = d
    Set wsRes = ThisWorkbook.Sheets(SH_RESULT)

    ClearResultsForDate mDate  ' 선택일의 이전 결과만 삭제 (다른 날짜는 보존)

    ReDim occupied(1 To nInsp, 1 To nSlots)
    ReDim lastRegion(1 To nInsp)
    ReDim taskCount(1 To nInsp)
    placedTotal = 0
    unassignedTotal = 0
    gLogCount = 0

    Dim nTasks As Long
    nTasks = LoadTasksForDate(d, mTasks)

    Dim priMap As Object
    Set priMap = CreateObject("Scripting.Dictionary")
    BuildPriorityMap priMap

    Dim r As Long, k As Long
    Dim insp As String, tstr As String, region As String
    Dim idx As Long, slotIdx As Long, dur As Long, freeSlot As Long, cand As Long
    Dim placed As Boolean
    Dim order() As Long

    ' 입력 점검: 지역 매핑 실패는 조용히 넘어가지 않고 로그
    For r = 1 To nTasks
        If mTasks(r).RegionStd = "" Then
            AddLog "이름표기불일치", "지역 '" & mTasks(r).RegionRaw & "'이(가) 04_지역표에 없거나 제외 대상입니다. 근접도 판정에서 제외됩니다.", mTasks(r).TaskName
        End If
        If mTasks(r).InspRaw <> "" And mTasks(r).Insp <> mTasks(r).InspRaw Then
            If InStr(mTasks(r).InspRaw, "-") > 0 Or InStr(mTasks(r).InspRaw, ",") > 0 Then
                AddLog "안내", "공동검사 '" & mTasks(r).InspRaw & "' - 첫 번째 검사자 열에 배치합니다.", mTasks(r).TaskName
            End If
        End If
    Next r
    If Not quiet Then
        AddLog "안내", "자동배분 시작 (" & Format(mDate, "yyyy-mm-dd") & "): 검사자 " & nInsp & "명, 업무 " & nTasks & "건", ""
    End If

    ' ---------- 1순위 : 검사자 + 시간 모두 지정 ----------
    For r = 1 To nTasks
        insp = mTasks(r).Insp
        tstr = mTasks(r).TimeStr
        If insp <> "" And tstr <> "" Then
            idx = FindInspectorIndex(inspectors, insp)
            slotIdx = TimeToSlotIndex(tstr, startTime, slotMin)
            dur = mTasks(r).Slots
            If idx = 0 Then
                AddLog "미배정", "지정검사자 '" & insp & "'가 이 날짜 근무자 명단에 없습니다 (휴가/비활성 포함).", mTasks(r).TaskName
                FailToUnassigned r
            ElseIf slotIdx < 1 Or slotIdx + dur - 1 > nSlots Then
                AddLog "충돌", "시간 '" & tstr & "' 업무가 시간축(" & TimeLabel(startTime) & "~" & TimeLabel(endTime) & ") 안에 들어가지 않습니다(자정을 넘는지 확인).", mTasks(r).TaskName
                FailToUnassigned r
            ElseIf Not SlotRangeFree(idx, slotIdx, dur, nSlots) Then
                AddLog "충돌", "검사자 '" & insp & "'의 " & tstr & " 시간대가 이미 다른 업무로 차 있습니다.", mTasks(r).TaskName
                FailToUnassigned r
            Else
                PlaceAtSlot r, idx, slotIdx, dur, 1
            End If
        End If
    Next r

    ' ---------- 2순위 : 시간만 지정 ----------
    For r = 1 To nTasks
        insp = mTasks(r).Insp
        tstr = mTasks(r).TimeStr
        If insp = "" And tstr <> "" Then
            slotIdx = TimeToSlotIndex(tstr, startTime, slotMin)
            dur = mTasks(r).Slots
            If slotIdx < 1 Or slotIdx + dur - 1 > nSlots Then
                AddLog "충돌", "시간 '" & tstr & "' 업무가 시간축(" & TimeLabel(startTime) & "~" & TimeLabel(endTime) & ") 안에 들어가지 않습니다(자정을 넘는지 확인).", mTasks(r).TaskName
                FailToUnassigned r
            Else
                region = mTasks(r).RegionStd
                order = RankInspectorsByProximity(region, nInsp, priMap)
                placed = False
                For k = LBound(order) To UBound(order)
                    cand = order(k)
                    If SlotRangeFree(cand, slotIdx, dur, nSlots) Then
                        PlaceAtSlot r, cand, slotIdx, dur, 2
                        placed = True
                        Exit For
                    End If
                Next k
                If Not placed Then
                    AddLog "충돌", "시간(" & tstr & ") 지정 업무를 배치할 빈 검사자를 찾지 못했습니다.", mTasks(r).TaskName
                    FailToUnassigned r
                End If
            End If
        End If
    Next r

    ' ---------- 3순위 : 검사자만 지정 ----------
    For r = 1 To nTasks
        insp = mTasks(r).Insp
        tstr = mTasks(r).TimeStr
        If insp <> "" And tstr = "" Then
            idx = FindInspectorIndex(inspectors, insp)
            If idx = 0 Then
                AddLog "미배정", "지정검사자 '" & insp & "'가 이 날짜 근무자 명단에 없습니다 (휴가/비활성 포함).", mTasks(r).TaskName
                FailToUnassigned r
            Else
                dur = mTasks(r).Slots
                freeSlot = FindFirstFreeSlot(idx, workLo, workHi, dur)
                If freeSlot = 0 Then
                    AddLog "충돌", "검사자 '" & insp & "'의 근무시간 내 빈 시간이 없어 배치하지 못했습니다.", mTasks(r).TaskName
                    FailToUnassigned r
                Else
                    PlaceAtSlot r, idx, freeSlot, dur, 3
                End If
            End If
        End If
    Next r

    ' ---------- 4순위 : 점심시간 ----------
    AssignLunch nInsp, inspectors, startTime, slotMin, nSlots

    ' ---------- 5순위 : 미지정 업무 ----------
    For r = 1 To nTasks
        insp = mTasks(r).Insp
        tstr = mTasks(r).TimeStr
        If insp = "" And tstr = "" Then
            region = mTasks(r).RegionStd
            order = RankInspectorsByProximity(region, nInsp, priMap)
            dur = mTasks(r).Slots
            placed = False
            For k = LBound(order) To UBound(order)
                cand = order(k)
                freeSlot = FindFirstFreeSlot(cand, workLo, workHi, dur)
                If freeSlot > 0 Then
                    PlaceAtSlot r, cand, freeSlot, dur, 4
                    placed = True
                    Exit For
                End If
            Next k
            If Not placed Then
                AddLog "미배정", "모든 검사자가 마감되어 배치하지 못했습니다.", mTasks(r).TaskName
                FailToUnassigned r
            End If
        End If
    Next r

    If Not quiet Then
        AddLog "안내", "자동배분 종료: 배치 " & placedTotal & "건, 미배정 " & unassignedTotal & "건", ""
    End If
    AutoAssignDate = True

done:
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevSU
    If msgText <> "" Then Say msgText, vbExclamation
    Exit Function

fail:
    msgText = "오류가 발생했습니다: " & Err.Description & vbCrLf & "07_배분로그에 진행 상황이 남아있을 수 있습니다."
    Resume done
End Function

' ------------------------------------------------------------
' 선택일에 결과가 없고 배분할 업무가 있으면 조용히 자동배분 (일일대장을 열 때 호출)
' 오늘 이후 날짜만 자동으로 배분합니다. 지난 날짜는 [자동배분 실행]을 눌러야 합니다.
' ------------------------------------------------------------
Public Sub EnsureResultsForDate(ByVal d As Date)
    If d < Date Then Exit Sub
    If CountResults(d) > 0 Then Exit Sub
    Dim tk() As TaskRec
    If LoadTasksForDate(d, tk) = 0 Then Exit Sub
    AutoAssignDate d, True
End Sub

' 결과가 만들어진 뒤 원데이터가 바뀌었는지 (업무 번호 집합 비교)
Public Function ResultsStale(ByVal d As Date) As Boolean
    Dim tk() As TaskRec
    Dim n As Long
    n = LoadTasksForDate(d, tk)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_RESULT)
    Dim last As Long, r As Long, nRes As Long, id As Long
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    last = ws.Cells(ws.rows.Count, R_NAME).End(xlUp).Row
    For r = 2 To last
        If IsDate(ws.Cells(r, R_DATE).Value) Then
            If DateValue(ws.Cells(r, R_DATE).Value) = d Then
                id = CLng(ws.Cells(r, R_ID).Value)
                If id > 0 Then
                    nRes = nRes + 1
                    seen(CStr(id)) = 1
                End If
            End If
        End If
    Next r
    If nRes = 0 Then Exit Function
    If nRes <> n Then
        ResultsStale = True
        Exit Function
    End If
    Dim i As Long
    For i = 1 To n
        If Not seen.Exists(CStr(tk(i).ID)) Then
            ResultsStale = True
            Exit Function
        End If
    Next i
End Function

' ------------------------------------------------------------
' 09_배분결과에 한 줄 추가
' ------------------------------------------------------------
Private Sub AppendResult(ByVal taskId As Long, ByVal nm As String, ByVal cap As String, ByVal reg As String, _
        ByVal durMin As String, ByVal tstr As String, ByVal insp As String, ByVal slotIdx As Long, _
        ByVal nS As Long, ByVal typ As Long, ByVal fee As String)
    Dim rr As Long
    rr = wsRes.Cells(wsRes.rows.Count, R_NAME).End(xlUp).Row + 1
    If rr < 2 Then rr = 2
    wsRes.Cells(rr, R_DATE).Value = mDate
    wsRes.Cells(rr, R_ID).Value = taskId
    wsRes.Cells(rr, R_NAME).Value = nm
    wsRes.Cells(rr, R_CAP).Value = cap
    wsRes.Cells(rr, R_REGION).Value = reg
    wsRes.Cells(rr, R_DUR).Value = durMin
    wsRes.Cells(rr, R_TIME).Value = tstr
    wsRes.Cells(rr, R_INSP).Value = insp
    wsRes.Cells(rr, R_SLOT).Value = slotIdx
    If slotIdx > 0 Then wsRes.Cells(rr, R_START).Value = SlotIndexToTimeStr(slotIdx, mStart, mSlotMin) Else wsRes.Cells(rr, R_START).Value = ""
    wsRes.Cells(rr, R_NSLOT).Value = nS
    wsRes.Cells(rr, R_TYPE).Value = typ
    wsRes.Cells(rr, R_FEE).Value = fee
End Sub

' 선택일의 기존 결과 행만 삭제
Public Sub ClearResultsForDate(ByVal d As Date)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_RESULT)
    Dim last As Long, r As Long
    last = ws.Cells(ws.rows.Count, R_NAME).End(xlUp).Row
    For r = last To 2 Step -1
        If IsDate(ws.Cells(r, R_DATE).Value) Then
            If DateValue(ws.Cells(r, R_DATE).Value) = d Then ws.rows(r).Delete
        End If
    Next r
End Sub

' ------------------------------------------------------------
' 배치 실패 업무를 '미지정' 열(배치 실패 블록)에 기록 (조용히 사라지지 않도록)
' ------------------------------------------------------------
Private Sub FailToUnassigned(ByVal r As Long)
    unassignedTotal = unassignedTotal + 1
    AppendResult mTasks(r).ID, mTasks(r).TaskName, mTasks(r).CapText, mTasks(r).RegionStd, CStr(mTasks(r).DurMin), _
        mTasks(r).TimeStr, "", 0, mTasks(r).Slots, 0, mTasks(r).Fee
End Sub

' ------------------------------------------------------------
' slotIdx부터 dur칸이 모두 비어있고 범위 내인지 확인
' ------------------------------------------------------------
Private Function SlotRangeFree(ByVal idx As Long, ByVal slotIdx As Long, ByVal dur As Long, ByVal nSlots As Long) As Boolean
    If slotIdx < 1 Or slotIdx + dur - 1 > nSlots Then
        SlotRangeFree = False
        Exit Function
    End If
    Dim s As Long
    For s = slotIdx To slotIdx + dur - 1
        If occupied(idx, s) <> 0 Then
            SlotRangeFree = False
            Exit Function
        End If
    Next s
    SlotRangeFree = True
End Function

' ------------------------------------------------------------
' 해당 검사자의 근무시간(lo~hi칸) 안에서 첫 번째로 dur칸 연속 비어있는 슬롯 위치를 반환 (없으면 0)
' ------------------------------------------------------------
Private Function FindFirstFreeSlot(ByVal idx As Long, ByVal lo As Long, ByVal hi As Long, ByVal dur As Long) As Long
    Dim s As Long, nSlots As Long
    nSlots = UBound(occupied, 2)
    For s = lo To hi - dur + 1
        If SlotRangeFree(idx, s, dur, nSlots) Then
            FindFirstFreeSlot = s
            Exit Function
        End If
    Next s
    FindFirstFreeSlot = 0
End Function

' ------------------------------------------------------------
' 신규 업무 지역(region)을 기준으로, 각 검사자와의 "거리"가 가까운 순으로
' 검사자 인덱스를 정렬해 반환 (거리가 같으면 배정건수가 적은 검사자 우선 = 공정배분)
' ------------------------------------------------------------
Private Function RankInspectorsByProximity(ByVal region As String, ByVal nInsp As Long, ByVal priMap As Object) As Long()
    Dim idxArr() As Long, distArr() As Long
    ReDim idxArr(1 To nInsp)
    ReDim distArr(1 To nInsp)
    Dim i As Long
    For i = 1 To nInsp
        idxArr(i) = i
        distArr(i) = RegionDistance(priMap, region, lastRegion(i)) * 1000 + taskCount(i)
    Next i
    ' 단순 버블정렬 (검사자 수가 20명 내외이므로 충분, 안정 정렬)
    Dim a As Long, b As Long, tmpI As Long, tmpD As Long
    For a = 1 To nInsp - 1
        For b = 1 To nInsp - a
            If distArr(b) > distArr(b + 1) Then
                tmpD = distArr(b): distArr(b) = distArr(b + 1): distArr(b + 1) = tmpD
                tmpI = idxArr(b): idxArr(b) = idxArr(b + 1): idxArr(b + 1) = tmpI
            End If
        Next b
    Next a
    RankInspectorsByProximity = idxArr
End Function

' ------------------------------------------------------------
' occupied 배열을 채우고 09_배분결과에 기록 (taskType: 1~3 순위, 4=미지정 업무)
' ------------------------------------------------------------
Private Sub PlaceAtSlot(ByVal r As Long, ByVal idx As Long, ByVal slotIdx As Long, _
        ByVal dur As Long, ByVal taskType As Long)
    Dim s As Long
    For s = slotIdx To slotIdx + dur - 1
        occupied(idx, s) = r
    Next s
    lastRegion(idx) = mTasks(r).RegionStd
    taskCount(idx) = taskCount(idx) + 1
    placedTotal = placedTotal + 1

    AppendResult mTasks(r).ID, mTasks(r).TaskName, mTasks(r).CapText, mTasks(r).RegionStd, CStr(mTasks(r).DurMin), _
        mTasks(r).TimeStr, mInsp(idx), slotIdx, dur, taskType, mTasks(r).Fee
End Sub

' ------------------------------------------------------------
' 4순위: 점심시간 배정 - 검사자별 1시간(2칸)을 설정된 탐색순서로 배치
' ------------------------------------------------------------
Private Sub AssignLunch(ByVal nInsp As Long, inspectors() As String, ByVal startTime As Date, _
        ByVal slotMin As Long, ByVal nSlots As Long)
    Dim lunchSlots As Long
    lunchSlots = SettingLong("점심시간길이_분") \ slotMin
    If lunchSlots < 1 Then lunchSlots = 1

    Dim orderList() As String
    orderList = Split(SettingText("점심탐색순서"), ",")

    Dim i As Long, o As Long, s As Long, slotIdx As Long
    Dim placed As Boolean
    For i = 1 To nInsp
        placed = False
        For o = LBound(orderList) To UBound(orderList)
            slotIdx = TimeToSlotIndex(Trim(orderList(o)), startTime, slotMin)
            If slotIdx >= 1 Then
                If SlotRangeFree(i, slotIdx, lunchSlots, nSlots) Then
                    For s = slotIdx To slotIdx + lunchSlots - 1
                        occupied(i, s) = -1
                    Next s
                    AppendResult 0, "점심시간", "", "", CStr(lunchSlots * slotMin), "", mInsp(i), slotIdx, lunchSlots, 5, ""
                    placed = True
                    Exit For
                End If
            End If
        Next o
        If Not placed Then
            AddLog "충돌", "검사자 '" & inspectors(i) & "'의 점심시간을 배정할 빈 시간을 찾지 못했습니다.", "(점심시간)"
        End If
    Next i
End Sub
