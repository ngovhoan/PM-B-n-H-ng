VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usf_add_item_u 
   Caption         =   "Add item"
   ClientHeight    =   1668
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3648
   OleObjectBlob   =   "usf_add_item_u.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usf_add_item_u"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_ok_Click()
    Dim i As Integer
    i = usf_order.lst_inventory.ListIndex
    Dim p_code As String
    p_code = usf_order.lst_inventory.List([i], [0])
    Dim tmp_v As Variant
    'Dim tmp_v_ As Variant
    Dim exc_s As String
    exc_s = "Select [Code],[Type],[Size],[Name],[Brand],[Unit],[Price_1],[Price_2],[Price_o],[Class],[ID] from Products where [Code] = '" & p_code & "'"
    tmp_v = DB.read_data(exc_s)
    'exc_s = "Select Distinct [Type] from Customers"
    'tmp_v_ = DB.read_data(exc_s)
    Dim s(15) As String
    s(0) = tmp_v(10, 0)
    s(1) = tmp_v(0, 0)
    s(2) = tmp_v(1, 0)
    s(3) = tmp_v(2, 0)
    s(4) = tmp_v(3, 0)
    s(5) = tmp_v(4, 0)
    s(6) = tmp_v(5, 0)
    If IsNumeric(txt_quantify) Then
        s(7) = txt_quantify.Value
    Else
        s(7) = 0
    End If
    If usf_order.cmb_customer_type.ListIndex = 0 Then
        s(8) = tmp_v(8, 0)
        
    ElseIf usf_order.cmb_customer_type.ListIndex = 1 Then
        s(8) = tmp_v(6, 0)
    Else
        s(8) = tmp_v(7, 0)
    End If
    s(9) = CStr(CDbl(s(7)) * CDbl(s(8)))
    '------------- Thanh tien ------
    If usf_order.cmb_customer_type.Value = Null Then
        s(10) = ""
    Else
        s(10) = usf_order.cmb_customer_type.Value
    End If
    s(11) = tmp_v(9, 0)
    s(12) = tmp_v(10, 0)
    If CDbl(tmp_v(8, 0)) = 0 Then
        s(13) = 0
    Else
        s(13) = CDbl(s(8)) - CDbl(tmp_v(8, 0))
    End If
    s(14) = CDbl(s(13)) * CDbl(s(7))
    Call usf_add_item.add_new_line_in_order(usf_order.x, s, usf_order.arr_list)
    usf_order.x = usf_order.x + 26
    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub
