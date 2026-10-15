Attribute VB_Name = "frmCalendar"
Attribute VB_Base = "0{F9E5C028-C3A6-499F-AA30-0ADF349F60E1}{779E5A79-4B04-46F4-BA6C-87FC83F31CF8}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' 달력 팝업: 월 이동, 날짜 선택, 오늘, (선택) 지움
' 배분 결과가 있는 날짜는 파란 배경, 원데이터에 업무가 있는 날짜는 노란 배경으로 표시합니다.

Public picked As Boolean
Public cleared As Boolean
Public PickedDate As Date

Private vY As Long, vM As Long
Private selD As Date
Private dts(1 To 42) As Date
Private markA As Object
Private markI As Object

Public Sub Init(ByVal initial As Date, ByVal allowClear As Boolean)
    picked = False
    cleared = False
    selD = initial
    vY = Year(initial)
    vM = Month(initial)
    btnClear.Visible = allowClear
    LoadMarks
    Render
End Sub

Private Sub LoadMarks()
    Set markA = CreateObject("Scripting.Dictionary")
    Set markI = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet, last As Long, r As Long, v As Variant
    Set ws = ThisWorkbook.Sheets(SH_RESULT)
    last = ws.Cells(ws.rows.Count, R_NAME).End(xlUp).Row
    For r = 2 To last
        v = ws.Cells(r, R_DATE).Value
        If IsDate(v) Then markA(CLng(DateValue(v))) = 1
    Next r
    ' 업무가 있는 날짜 = 10_원데이터의 접수일자
    Dim vr As Variant, i As Long
    vr = RawData()
    If Not IsEmpty(vr) Then
        For i = 1 To UBound(vr, 1)
            If IsDate(vr(i, RW_DATE)) Then markI(CLng(DateValue(vr(i, RW_DATE)))) = 1
        Next i
    End If
End Sub

Private Sub Render()
    Dim first As Date, offs As Long, i As Long, d As Date, b As Object
    first = DateSerial(vY, vM, 1)
    offs = Weekday(first, vbSunday) - 1
    lblTitle.caption = vY & "년 " & vM & "월"
    For i = 1 To 42
        d = first + (i - 1 - offs)
        dts(i) = d
        Set b = Me.Controls("dayBtn" & Format(i, "00"))
        b.caption = CStr(Day(d))
        If Month(d) <> vM Then
            b.ForeColor = RGB(170, 170, 170)
        ElseIf Weekday(d, vbSunday) = 1 Then
            b.ForeColor = RGB(192, 0, 0)
        ElseIf Weekday(d, vbSunday) = 7 Then
            b.ForeColor = RGB(0, 90, 180)
        Else
            b.ForeColor = RGB(0, 0, 0)
        End If
        If markA.Exists(CLng(d)) Then
            b.BackColor = RGB(189, 215, 238)
        ElseIf markI.Exists(CLng(d)) Then
            b.BackColor = RGB(255, 242, 204)
        Else
            b.BackColor = RGB(255, 255, 255)
        End If
        b.Font.Bold = (d = Date)
        If d = selD Then
            b.BackColor = RGB(31, 56, 100)
            b.ForeColor = RGB(255, 255, 255)
        End If
    Next i
End Sub

Private Sub PickDay(ByVal i As Long)
    PickedDate = dts(i)
    picked = True
    Me.Hide
End Sub

Private Sub btnPrevM_Click()
    vM = vM - 1
    If vM < 1 Then vM = 12: vY = vY - 1
    Render
End Sub

Private Sub btnNextM_Click()
    vM = vM + 1
    If vM > 12 Then vM = 1: vY = vY + 1
    Render
End Sub

Private Sub btnPrevY_Click()
    vY = vY - 1
    Render
End Sub

Private Sub btnNextY_Click()
    vY = vY + 1
    Render
End Sub

Private Sub btnToday_Click()
    PickedDate = Date
    picked = True
    Me.Hide
End Sub

Private Sub btnClear_Click()
    cleared = True
    picked = True
    Me.Hide
End Sub

Private Sub btnClose_Click()
    picked = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        picked = False
        Me.Hide
    End If
End Sub
Private Sub dayBtn01_Click()
    PickDay 1
End Sub

Private Sub dayBtn02_Click()
    PickDay 2
End Sub

Private Sub dayBtn03_Click()
    PickDay 3
End Sub

Private Sub dayBtn04_Click()
    PickDay 4
End Sub

Private Sub dayBtn05_Click()
    PickDay 5
End Sub

Private Sub dayBtn06_Click()
    PickDay 6
End Sub

Private Sub dayBtn07_Click()
    PickDay 7
End Sub

Private Sub dayBtn08_Click()
    PickDay 8
End Sub

Private Sub dayBtn09_Click()
    PickDay 9
End Sub

Private Sub dayBtn10_Click()
    PickDay 10
End Sub

Private Sub dayBtn11_Click()
    PickDay 11
End Sub

Private Sub dayBtn12_Click()
    PickDay 12
End Sub

Private Sub dayBtn13_Click()
    PickDay 13
End Sub

Private Sub dayBtn14_Click()
    PickDay 14
End Sub

Private Sub dayBtn15_Click()
    PickDay 15
End Sub

Private Sub dayBtn16_Click()
    PickDay 16
End Sub

Private Sub dayBtn17_Click()
    PickDay 17
End Sub

Private Sub dayBtn18_Click()
    PickDay 18
End Sub

Private Sub dayBtn19_Click()
    PickDay 19
End Sub

Private Sub dayBtn20_Click()
    PickDay 20
End Sub

Private Sub dayBtn21_Click()
    PickDay 21
End Sub

Private Sub dayBtn22_Click()
    PickDay 22
End Sub

Private Sub dayBtn23_Click()
    PickDay 23
End Sub

Private Sub dayBtn24_Click()
    PickDay 24
End Sub

Private Sub dayBtn25_Click()
    PickDay 25
End Sub

Private Sub dayBtn26_Click()
    PickDay 26
End Sub

Private Sub dayBtn27_Click()
    PickDay 27
End Sub

Private Sub dayBtn28_Click()
    PickDay 28
End Sub

Private Sub dayBtn29_Click()
    PickDay 29
End Sub

Private Sub dayBtn30_Click()
    PickDay 30
End Sub

Private Sub dayBtn31_Click()
    PickDay 31
End Sub

Private Sub dayBtn32_Click()
    PickDay 32
End Sub

Private Sub dayBtn33_Click()
    PickDay 33
End Sub

Private Sub dayBtn34_Click()
    PickDay 34
End Sub

Private Sub dayBtn35_Click()
    PickDay 35
End Sub

Private Sub dayBtn36_Click()
    PickDay 36
End Sub

Private Sub dayBtn37_Click()
    PickDay 37
End Sub

Private Sub dayBtn38_Click()
    PickDay 38
End Sub

Private Sub dayBtn39_Click()
    PickDay 39
End Sub

Private Sub dayBtn40_Click()
    PickDay 40
End Sub

Private Sub dayBtn41_Click()
    PickDay 41
End Sub

Private Sub dayBtn42_Click()
    PickDay 42
End Sub


