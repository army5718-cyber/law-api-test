Attribute VB_Name = "Sheet12"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Activate()
    ' 시트를 열 때마다 선택일 기준으로 접수대장과 배분현황판을 다시 만듭니다.
    Day_Render
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' 블록을 끌어다 놓은 뒤 다른 곳을 클릭하면 이동을 확인합니다.
    Board_CheckMoves
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' 날짜(H4:L4)를 더블클릭하면 달력으로 날짜를 고릅니다.
    If Not Intersect(Target, Me.Range("H4:L4")) Is Nothing Then
        Cancel = True
        Day_Pick
    End If
End Sub
