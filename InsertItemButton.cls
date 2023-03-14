VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InsertItemButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents cmdInsert As MSForms.CommandButton
Attribute cmdInsert.VB_VarHelpID = -1

Public Sub cmdInsert_Click()
    
    Dim str() As String
    
    Dim i As Integer
    i = CInt(Mid(cmdInsert.name, 17, 3))
    
    Dim flt As Integer
    flt = usf_order.Controls("frm_item_" & i).top
    'Call usf_add_item.add_new_line_in_order(flt + 26)
    Call list_frm_update_insert(i, flt)
    If usf_order.isPre = False Then
        str = Algorithm.acess_an_item_to_string_arr(usf_order.arr_list, i, 15)
        Call usf_add_item.add_new_line_in_order(flt + 26, str, usf_order.arr_list)
    Else
        str = Algorithm.acess_an_item_to_string_arr(usf_order.arr_list_previous, i, 16)
        Call usf_add_item.add_new_line_in_order(flt + 26, str, usf_order.arr_list_previous)
    End If
End Sub

Public Sub list_frm_update_insert(ByVal i As Integer, ByVal flt As Integer)
    Dim ctrl As MSForms.Control
    Dim j As Integer
    For Each ctrl In usf_order.Controls
        For j = 1 To usf_order.collection_cmd.Count
            If StrComp(ctrl.name, "frm_item_" & j, vbTextCompare) = 0 And ctrl.top > flt Then
                ctrl.top = ctrl.top + 26
            End If
        Next
    Next
    usf_order.x = usf_order.x + 26
End Sub
