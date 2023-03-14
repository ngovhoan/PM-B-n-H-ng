VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MouseMoveFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents frmEvent As MSForms.Frame
Attribute frmEvent.VB_VarHelpID = -1

Private Sub frmEvent_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    frmEvent.BackColor = RGB(255, 255, 255)
End Sub

Private Sub frmEvent_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    
    Dim i As Integer
    i = CInt(Mid(frmEvent.name, 10, 3))
    Dim ctr As MSForms.Control
    For Each ctr In usf_order.Controls
        If StrComp(ctr.name, "frm_item_" & i) = 0 Then
            Call Animation.hightlight(usf_order, ctr.name)
        Else
            Call Animation.normal(usf_order)
        End If
    Next
    
    
End Sub

Private Sub test()
    Dim i_top As Integer, i_bottom As Integer, i_left As Integer, i_right As Integer
    i_top = frmEvent.top
    i_bottom = i_top + frmEvent.Height
    i_left = frmEvent.Left
    i_right = i_left + frmEvent.Width
    If x > i_left And x < i_right And y > i_top And y < i_bottom Then
        Call Animation.hightlight(usf_order, frmEvent.name)
        Call Animation.normal(usf_order)
        'MsgBox " v" & i_top
    Else
        Call Animation.normal(usf_order)
    End If
End Sub
