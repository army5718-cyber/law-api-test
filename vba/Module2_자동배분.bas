Attribute VB_Name = "Module2_자동배분"
Option Explicit

' ============================================================
' KESCO 부천김포지사 검사업무 ERP
' Module2_자동배분 : 5단계 배분 알고리즘 (계산만 담당)
'
' 배분 대상 = 11_일일대장 화면에 입력된 업무 (정기검사/사용전검사 칸, 메모 제외)
'   - 화면의 검사자/약속시간 칸이 채워져 있으면 그것을 "지정"으로 보고 우선 배치합니다.
'   - 시간축은 24시간이고, 약속시간 없는 업무와 점심은 근무시간(02_설정 운영시작~운영종료) 안에만 배치합니다.
'
' 순서 고정:
'   1순위: 검사자+시간 모두 지정  2순위: 시간만 지정  3순위: 검사자만 지정
'   4순위: 점심시간              5순위: 미지정 업무
'
' 2/5순위의 "가까운 검사자" 판정 (예전 05_우선순위매트릭스 시트를 대체):
'   각 검사자의 가장 최근 배정 업무 지역과 신규 업무 지역을 04_지역표 기준으로 비교
'     같은 표준지역 0 < 같은 클러스터(읍/면/동/구) 1 < 같은 권역 2 < 그 밖 3 (정보 없음 500)
'   같은 거리이면 배정 건수가 적은 검사자를 먼저 씁니다(공정 배분).
'   기준을 바꾸려면 AreaDistance 함수만 고치세요.
'
' ※ 조용히 고치지 않기: 배치하지 못한 업무는 07_배분로그에 기록하고 '미지정' 열에 배치 실패 블록으로 표시합니다.
' ============================================================

Private occupied() As Long      ' occupied(검사자, 슬롯) = 업무 번호 (0=빈칸, -1=점심)
Private lastRegion() As String  ' lastRegion(검사자) = 최근 배정 업무의 표준지역
Private taskCount() As Long     ' taskCount(검사자) = 배정 건수
Private placedTotal As Long
Private unassignedTotal As Long

Private mInsp() As String
Private mDate As Date
Private mStart As Date
Private mSlotMin As Long
Private mTasks() As TaskRec
Private clusterOf As Object
Private zoneOf As Object

' 결과 버퍼 (마지막에 한 번에 시트에 씀)
Private buf() As Variant
Private bufN As Long

' ------------------------------------------------------------
' 버튼용: 화면 내용을 기준으로 (다시) 자동배분하고 현황판을 새로 그림
' ------------------------------------------------------------
Sub RunAutoAssign()
    Dim d As Date
    d = BoardDate()
    Dim nm As Long
    nm = CountManual(d)
    If nm > 0 Then
        If Not Ask(Format(d, "yyyy-mm-dd") & " 에 수동으로 옮긴 블록이 " & nm & "건 있습니다." & vbCrLf & _
                   "다시 자동배분하면 수동 조정은 사라집니다. 계속할까요?") Then Exit Sub
    End If
    Dim ok As Boolean
    FastBegin
    On Error GoTo fail
    ok = AutoAssignDate(d, False)
    RebuildGrid
    FastEnd
    If ok Then
        Say "자동 배분이 완료되었습니다 (" & Format(d, "yyyy-mm-dd") & ")." & vbCrLf & _
            "배치 " & placedTotal & "건 / 미배정 " & unassignedTotal & "건" & vbCrLf & _
            "충돌·미배정 항목은 07_배분로그에서 확인하세요." & vbCrLf & _
            "블록을 끌어 조정한 뒤 [입력 완료]를 누르면 검사자가 10_원데이터에 저장됩니다.", _
            IIf(unassignedTotal > 0, vbExclamation, vbInformation)
    End If
    Exit Sub
fail:
    Dim em As String
    em = Err.Description
    FastEnd
    Say "자동배분 중 오류가 발생했습니다: " & em, vbCritical
End Sub

' ------------------------------------------------------------
' 09_배분결과 읽기/쓰기 (한 번에)
' ------------------------------------------------------------
Private Function ResLast(ws As Worksheet) As Long
    ResLast = ws.Cells(ws.rows.Count, R_NAME).End(xlUp).Row
End Function

Public Function CountManual(ByVal d As Date) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_RESULT)
    Dim last As Long
    last = ResLast(ws)
    If last < 2 Then Exit Function
    Dim v As Variant, r As Long, n As Long
    v = ws.Range(ws.Cells(2, 1), ws.Cells(last, R_COLS)).Value
    For r = 1 To UBound(v, 1)
        If IsDate(v(r, R_DATE)) Then
            If DateValue(v(r, R_DATE)) = d Then
                If CStr(v(r, R_MANUAL)) = "Y" Then n = n + 1
            End If
        End If
    Next r
    CountManual = n
End Function

Public Function CountResults(ByVal d As Date) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_RESULT)
    Dim last As Long
    last = ResLast(ws)
    If last < 2 Then Exit Function
    Dim v As Variant, r As Long, n As Long
    v = ws.Range(ws.Cells(2, 1), ws.Cells(last, R_ID)).Value
    For r = 1 To UBound(v, 1)
        If IsDate(v(r, 1)) Then
            If DateValue(v(r, 1)) = d Then n = n + 1
        End If
    Next r
    CountResults = n
End Function

' 선택일의 기존 결과 행만 삭제 (다른 날짜는 보존) - 한 번에 다시 씀
Public Sub ClearResultsForDate(ByVal d As Date)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_RESULT)
    Dim last As Long
    last = ResLast(ws)
    If last < 2 Then Exit Sub
    Dim v As Variant, w() As Variant, r As Long, c As Long, keep As Long, hit As Boolean
    v = ws.Range(ws.Cells(2, 1), ws.Cells(last, R_COLS)).Value
    ReDim w(1 To UBound(v, 1), 1 To R_COLS)
    For r = 1 To UBound(v, 1)
        hit = False
        If IsDate(v(r, R_DATE)) Then
            If DateValue(v(r, R_DATE)) = d Then hit = True
        End If
        If Not hit Then
            keep = keep + 1
            For c = 1 To R_COLS
                w(keep, c) = v(r, c)
            Next c
        End If
    Next r
    If keep = UBound(v, 1) Then Exit Sub
    ws.Range(ws.Cells(2, 1), ws.Cells(last, R_COLS)).ClearContents
    If keep > 0 Then
        FormatResultColumns ws, keep + 1
        ws.Range(ws.Cells(2, 1), ws.Cells(keep + 1, R_COLS)).Value = w
    End If
End Sub

' 문자로 보관할 열(용량/지역/소요/약속시간/시작시각)은 텍스트 서식으로 (시간이 숫자로 바뀌지 않게)
Public Sub FormatResultColumns(ws As Worksheet, ByVal lastRow As Long)
    Dim c As Variant
    For Each c In Array(R_CAP, R_REGION, R_DUR, R_TIME, R_START)
        ws.Range(ws.Cells(2, CLng(c)), ws.Cells(lastRow + 60, CLng(c))).NumberFormat = "@"
    Next c
End Sub

Private Sub BufAdd(ByVal taskKey As Long, ByVal nm As String, ByVal cap As String, ByVal reg As String, _
        ByVal durMin As String, ByVal tstr As String, ByVal insp As String, ByVal slotIdx As Long, _
        ByVal nS As Long, ByVal typ As Long, ByVal fee As String)
    bufN = bufN + 1
    If bufN > UBound(buf, 2) Then ReDim Preserve buf(1 To R_COLS, 1 To UBound(buf, 2) + 64)
    buf(R_DATE, bufN) = mDate
    buf(R_ID, bufN) = taskKey
    buf(R_NAME, bufN) = nm
    buf(R_CAP, bufN) = cap
    buf(R_REGION, bufN) = reg
    buf(R_DUR, bufN) = durMin
    buf(R_TIME, bufN) = tstr
    buf(R_INSP, bufN) = insp
    buf(R_SLOT, bufN) = slotIdx
    If slotIdx > 0 Then buf(R_START, bufN) = SlotIndexToTimeStr(slotIdx, mStart, mSlotMin) Else buf(R_START, bufN) = ""
    buf(R_NSLOT, bufN) = nS
    buf(R_TYPE, bufN) = typ
    buf(R_FEE, bufN) = fee
End Sub

Private Sub BufFlush()
    If bufN = 0 Then Exit Sub
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_RESULT)
    Dim first As Long
    first = ResLast(ws) + 1
    If first < 2 Then first = 2
    FormatResultColumns ws, first + bufN
    Dim w() As Variant, i As Long, c As Long
    ReDim w(1 To bufN, 1 To R_COLS)
    For i = 1 To bufN
        For c = 1 To R_COLS
            w(i, c) = buf(c, i)
        Next c
    Next i
    ws.Range(ws.Cells(first, 1), ws.Cells(first + bufN - 1, R_COLS)).Value = w
End Sub

' ------------------------------------------------------------
' 화면에서 읽은 줄 중 배분할 업무(정기/사용전, 상호 있음)만 mTasks 로 모음
' ------------------------------------------------------------
Private Function CollectTasks(rec() As TaskRec, ByVal n As Long) As Long
    Dim i As Long, k As Long
    ReDim mTasks(1 To IIf(n < 1, 1, n))
    For i = 1 To n
        If IsTaskKind(rec(i).Kind) And rec(i).TaskName <> "" Then
            k = k + 1
            mTasks(k) = rec(i)
        End If
    Next i
    If k = 0 Then ReDim mTasks(1 To 1) Else ReDim Preserve mTasks(1 To k)
    CollectTasks = k
End Function

' ------------------------------------------------------------
' 선택일을 자동배분해 09_배분결과에 저장 (화면은 그리지 않음)
' quiet=True 이면 안내/요약 로그를 남기지 않음 (일일대장을 열 때 자동 실행하는 경우)
' ------------------------------------------------------------
Public Function AutoAssignDate(ByVal d As Date, ByVal quiet As Boolean) As Boolean
    Dim msgText As String
    FastBegin
    On Error GoTo fail

    Dim rec() As TaskRec, nRec As Long, leave As Object
    nRec = ReadScreen(rec)
    Set leave = LeaveFromRecs(rec, nRec)

    ' ---------- 0. 사전준비 ----------
    Dim inspectors() As String
    inspectors = GetActiveInspectorsFor(leave)
    If Not HasInspectors(inspectors) Then
        msgText = "배분할 검사자가 없습니다. 01_검사자명단의 오늘근무 값을 확인하세요."
        GoTo done
    End If
    Dim nInsp As Long
    nInsp = UBound(inspectors) - LBound(inspectors) + 1
    mInsp = inspectors

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

    ClearResultsForDate d

    ReDim occupied(1 To nInsp, 1 To nSlots)
    ReDim lastRegion(1 To nInsp)
    ReDim taskCount(1 To nInsp)
    ReDim buf(1 To R_COLS, 1 To 64)
    bufN = 0
    placedTotal = 0
    unassignedTotal = 0
    gLogCount = 0

    Dim nTasks As Long
    nTasks = CollectTasks(rec, nRec)
    LoadRegionTopology clusterOf, zoneOf

    Dim r As Long, k As Long
    Dim insp As String, tstr As String, region As String
    Dim idx As Long, slotIdx As Long, dur As Long, freeSlot As Long, cand As Long
    Dim placed As Boolean
    Dim order() As Long

    For r = 1 To nTasks
        If mTasks(r).RegionStd = "" Then
            AddLog "이름표기불일치", "지역 '" & mTasks(r).RegionRaw & "'이(가) 04_지역표에 없거나 제외 대상입니다. 근접도 판정에서 제외됩니다.", mTasks(r).TaskName
        End If
        If InStr(mTasks(r).InspRaw, "-") > 0 Or InStr(mTasks(r).InspRaw, ",") > 0 Then
            AddLog "안내", "공동검사 '" & mTasks(r).InspRaw & "' - 첫 번째 검사자 열에 배치합니다.", mTasks(r).TaskName
        End If
    Next r
    If Not quiet Then
        AddLog "안내", "자동배분 시작 (" & Format(d, "yyyy-mm-dd") & "): 검사자 " & nInsp & "명, 업무 " & nTasks & "건", ""
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
                order = RankInspectorsByProximity(mTasks(r).RegionStd, nInsp)
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
            order = RankInspectorsByProximity(mTasks(r).RegionStd, nInsp)
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

    BufFlush
    If Not quiet Then
        AddLog "안내", "자동배분 종료: 배치 " & placedTotal & "건, 미배정 " & unassignedTotal & "건", ""
    End If
    AutoAssignDate = True

done:
    FastEnd
    If msgText <> "" Then Say msgText, vbExclamation
    Exit Function

fail:
    msgText = "오류가 발생했습니다: " & Err.Description
    Resume done
End Function

' ------------------------------------------------------------
' 선택일에 결과가 없거나(오늘 이후만) 화면과 어긋나 있으면 조용히 자동배분
' 지난 날짜는 어긋난 결과만 지우고, 배분은 [자동배분]을 눌렀을 때만 합니다.
' ------------------------------------------------------------
Public Sub EnsureResultsForDate(ByVal d As Date)
    Dim nRes As Long
    nRes = CountResults(d)
    If nRes > 0 Then
        If Not ResultsStale(d) Then Exit Sub
        If d < Date Then
            ClearResultsForDate d
            Exit Sub
        End If
    ElseIf d < Date Then
        Exit Sub
    End If
    Dim rec() As TaskRec, n As Long, i As Long, hasTask As Boolean
    n = ReadScreen(rec)
    For i = 1 To n
        If IsTaskKind(rec(i).Kind) Then hasTask = True: Exit For
    Next i
    If hasTask Then AutoAssignDate d, True
End Sub

' 결과가 만들어진 뒤 화면의 업무가 바뀌었는지 (화면 행 키 집합 비교)
Public Function ResultsStale(ByVal d As Date) As Boolean
    Dim rec() As TaskRec, n As Long
    n = ReadScreen(rec)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_RESULT)
    Dim last As Long
    last = ResLast(ws)
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim nRes As Long, r As Long, id As Long
    If last >= 2 Then
        Dim v As Variant
        v = ws.Range(ws.Cells(2, 1), ws.Cells(last, R_ID)).Value
        For r = 1 To UBound(v, 1)
            If IsDate(v(r, R_DATE)) Then
                If DateValue(v(r, R_DATE)) = d Then
                    id = CLng(v(r, R_ID))
                    If id > 0 Then
                        nRes = nRes + 1
                        seen(CStr(id)) = 1
                    End If
                End If
            End If
        Next r
    End If
    If nRes = 0 Then Exit Function
    Dim i As Long, nT As Long
    For i = 1 To n
        If IsTaskKind(rec(i).Kind) And rec(i).TaskName <> "" Then
            nT = nT + 1
            If Not seen.Exists(CStr(rec(i).Key)) Then
                ResultsStale = True
                Exit Function
            End If
        End If
    Next i
    If nT <> nRes Then ResultsStale = True
End Function

' ------------------------------------------------------------
' 배치 실패 업무를 '미지정' 열(배치 실패 블록)에 기록 (조용히 사라지지 않도록)
' ------------------------------------------------------------
Private Sub FailToUnassigned(ByVal r As Long)
    unassignedTotal = unassignedTotal + 1
    BufAdd mTasks(r).Key, mTasks(r).TaskName, mTasks(r).CapText, mTasks(r).RegionStd, CStr(mTasks(r).DurMin), _
        mTasks(r).TimeStr, "", 0, mTasks(r).Slots, 0, mTasks(r).Fee
End Sub

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

' 해당 검사자의 근무시간(lo~hi칸) 안에서 첫 번째로 dur칸 연속 비어있는 슬롯 위치 (없으면 0)
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

' 두 표준지역의 거리: 같은 지역 0 / 같은 클러스터 1 / 같은 권역 2 / 그 밖 3 / 정보 없음 500
Private Function AreaDistance(ByVal a As String, ByVal b As String) As Long
    If a = "" Or b = "" Then
        AreaDistance = 500
    ElseIf a = b Then
        AreaDistance = 0
    ElseIf clusterOf.Exists(a) And clusterOf.Exists(b) And clusterOf(a) = clusterOf(b) Then
        AreaDistance = 1
    ElseIf zoneOf.Exists(a) And zoneOf.Exists(b) And zoneOf(a) = zoneOf(b) Then
        AreaDistance = 2
    Else
        AreaDistance = 3
    End If
End Function

' 신규 업무 지역 기준으로 거리가 가까운 검사자 순서 (같으면 배정건수 적은 검사자 우선)
Private Function RankInspectorsByProximity(ByVal region As String, ByVal nInsp As Long) As Long()
    Dim idxArr() As Long, distArr() As Long
    ReDim idxArr(1 To nInsp)
    ReDim distArr(1 To nInsp)
    Dim i As Long
    For i = 1 To nInsp
        idxArr(i) = i
        distArr(i) = AreaDistance(region, lastRegion(i)) * 1000 + taskCount(i)
    Next i
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

Private Sub PlaceAtSlot(ByVal r As Long, ByVal idx As Long, ByVal slotIdx As Long, _
        ByVal dur As Long, ByVal taskType As Long)
    Dim s As Long
    For s = slotIdx To slotIdx + dur - 1
        occupied(idx, s) = r
    Next s
    lastRegion(idx) = mTasks(r).RegionStd
    taskCount(idx) = taskCount(idx) + 1
    placedTotal = placedTotal + 1
    BufAdd mTasks(r).Key, mTasks(r).TaskName, mTasks(r).CapText, mTasks(r).RegionStd, CStr(mTasks(r).DurMin), _
        mTasks(r).TimeStr, mInsp(idx), slotIdx, dur, taskType, mTasks(r).Fee
End Sub

' 4순위: 점심시간 - 검사자별 1시간(2칸)을 설정된 탐색순서로 배치
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
                    BufAdd 0, "점심시간", "", "", CStr(lunchSlots * slotMin), "", mInsp(i), slotIdx, lunchSlots, 5, ""
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
