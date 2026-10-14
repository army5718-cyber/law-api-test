Attribute VB_Name = "현재_통합_문서"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Workbook_Open()
    ' 파일을 열면 오늘 날짜의 일일대장(입력 화면) + 배분현황판을 보여줍니다.
    On Error Resume Next
    FastReset
    SetBoardDate Date
    Day_Render
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' 예약된 이동 확인 타이머가 파일을 다시 열지 않도록 해제
    Board_StopPoll
End Sub
