VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AddRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents cmdEvent As MSForms.CommandButton
Attribute cmdEvent.VB_VarHelpID = -1
Public column_cmd As New Collection
Public Sub cmdEvent_Click()
    'usf_test.Controls("frm_").SetFocus
    Dim frm_test As MSForms.Frame
    Dim lb_test As MSForms.Label
    Dim cmd_test As MSForms.CommandButton
    Set frm_test = usf_test.Controls(1).Add("Forms.Frame.1", "frm_" & usf_test.id, True)
    With frm_test
        .Height = 15
        .Width = 150
        .top = 6 + usf_test.x
        .Left = 6
    End With
    
    
    Set lb_test = frm_test.Controls.Add("Forms.Label.1", "lb_" & usf_test.id, True)
    With lb_test
        .Width = 10
        .Height = 5
        .top = 2
        .Left = 5
        .BackColor = RGB(155, 0, 125)
    End With
    
    Set cmd_test = frm_test.Controls.Add("Forms.commandbutton.1", "cmd_" & usf_test.id, True)
    With cmd_test
        .Width = 10
        .Height = 5
        .top = 2
        .Left = 18
        .BackColor = RGB(155, 0, 125)
        .Caption = "test"
    End With
    Dim cmd As DeleteRow
    Set cmd = New DeleteRow
    Set cmd.cmdDel = cmd_test
    column_cmd.Add cmd
    
    Call check_length
    'usf_test.Controls("frm_").Height = usf_test.Controls("frm_").Height + 15
    
    
    usf_test.id = usf_test.id + 1
    usf_test.x = usf_test.x + 15
End Sub

Private Sub check_length()
    If usf_test.x >= usf_test.Controls("frm_").Height Then
        With usf_test.Controls("frm_")
            .ScrollBars = fmScrollBarsHorizontal
            .ScrollWidth = 200
        End With
    Else
        With usf_test.Controls("frm_")
            .ScrollBars = fmScrollBarsNone
            '.ScrollHeight = 200
        End With
    End If
End Sub

