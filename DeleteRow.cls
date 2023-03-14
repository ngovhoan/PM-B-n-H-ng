VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeleteRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents cmdDel As MSForms.CommandButton
Attribute cmdDel.VB_VarHelpID = -1

Private Sub cmdDel_Click()
    Dim e As AddRow
    Set e = New AddRow
    Call e.cmdEvent_Click
End Sub
