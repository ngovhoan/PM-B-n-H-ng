VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EditItemButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents cmdEdit As MSForms.CommandButton
Attribute cmdEdit.VB_VarHelpID = -1

Private Sub cmdEdit_Click()
    usf_order.index_edit = CInt(Mid(cmdEdit.name, 15, 3))
    Dim ctrl As Control
    'Set ctrl = usf_order.Controls("frm_item_" & usf_order.index_edit)
    For Each ctrl In usf_order.Controls
        If InStr(1, ctrl.name, "frm_item", vbTextCompare) > 0 Then
            ctrl.BackColor = RGB(255, 255, 255)
        End If
    Next
    usf_order.Controls("frm_item_" & usf_order.index_edit).BackColor = RGB(190, 190, 190)
    If usf_edit_item.Visible Then
        Unload usf_edit_item
        usf_order.Controls("frm_item_" & usf_order.index_edit).BackColor = RGB(255, 255, 255)
        Exit Sub
    Else
        usf_edit_item.Show (vbModeless)
    End If
End Sub
