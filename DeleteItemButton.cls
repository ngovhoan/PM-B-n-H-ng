VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeleteItemButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents cmdDelete As MSForms.CommandButton
Attribute cmdDelete.VB_VarHelpID = -1

Private Sub cmdDelete_Click()

    Dim i As Integer
    i = CInt(Mid(cmdDelete.name, 17, 3))
    
    Dim flt As Integer
    flt = usf_order.Controls("frm_item_" & i).top
    
    'usf_order.Controls.Remove ("lb_count_item_" & i)
    usf_order.Controls.Remove ("frm_item_" & i)
    
    Call list_frm_update_delete(i, flt)
    Call usf_add_item.check_scrollbar_of_list
    Call usf_order.to_total_price(usf_order)
    If usf_order.isPre = False Then
        '------- Xoa item trong array_list-------------
        Call Algorithm.delete_an_item(usf_order.arr_list, i)
        '------- In lai du lieu tron array_list vao excel -------
        'Call Algorithm.print_array_list(usf_order.arr_list, ThisWorkbook.Worksheets("Order"))
    Else
        '------- Xoa item trong array_list-------------
        Call Algorithm.delete_an_item(usf_order.arr_list_previous, i)
        '------- In lai du lieu tron array_list vao excel -------
        'Call Algorithm.print_array_list(usf_order.arr_list_previous, ThisWorkbook.Worksheets("Order"))
    End If
End Sub

Public Sub list_frm_update_delete(ByVal i As Integer, ByVal flt As Integer)
    Dim ctrl As MSForms.Control
    Dim j As Integer
    For Each ctrl In usf_order.Controls
        For j = 1 To usf_order.collection_cmd.Count
            If ctrl.name = "frm_item_" & j And ctrl.top > flt Then
                ctrl.top = ctrl.top - 26
            End If
        Next
    Next
    usf_order.x = usf_order.x - 26
End Sub
