VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usf_edit_item 
   Caption         =   "Edit"
   ClientHeight    =   2736
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   16080
   OleObjectBlob   =   "usf_edit_item.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usf_edit_item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim s() As String

Private Sub cmb_brand_Change()
    Dim s As String
    cmb_unit.Clear
    cmb_unit.Enabled = True
    s = "select distinct [Unit] from Products where [Class]='" & cmb_class.Value & "' and [Type]='" & cmb_type.Value & "' and [Size]='" & cmb_size.Value & "' and [Brand]='" & cmb_brand.Value & "'"
    Call DB.fill_combobox(cmb_unit, s, 4)
    If cmb_unit.ListCount > 0 And cmb_unit.ListIndex = -1 Then
        cmb_unit.ListIndex = 0
    End If
    Call set_name_code_id
End Sub

Private Sub cmb_class_Change()
    Dim s As String
    cmb_type.Clear
    cmb_size.Clear
    cmb_brand.Clear
    cmb_unit.Clear
    cmb_type.Enabled = True
    s = "Select distinct [Type] from Products where [Class]='" & cmb_class.Value & "'"
    Call DB.fill_combobox(cmb_type, s, 1)
    Call set_name_code_id
End Sub

Private Sub cmb_size_Change()
    Dim s As String
    cmb_brand.Clear
    cmb_unit.Clear
    cmb_brand.Enabled = True
    s = "Select distinct [Brand] from Products where [Class] = '" & cmb_class.Value & "' and [Type] = '" & cmb_type.Value & "' and [Size] = '" & cmb_size.Value & "'"
    Call DB.fill_combobox(cmb_brand, s, 3)
    Call set_name_code_id
End Sub

Private Sub cmb_type_Change()
    Dim s As String
    cmb_size.Clear
    cmb_brand.Clear
    cmb_unit.Clear
    cmb_size.Enabled = True
    s = "Select distinct [Size] from Products where [Class]='" & cmb_class.Value & "' and [Type]='" & cmb_type.Value & "'"
    Call DB.fill_combobox(cmb_size, s, 2)
    Call set_name_code_id
End Sub

Private Sub cmb_type_price_Change()
    Call set_price
End Sub

Private Sub cmb_unit_Change()
    Dim s As String
    cmb_type_price.Enabled = True
    s = "Select [Price_1] from Products where [Class]='" & cmb_class.Value & "' and [Type]='" & cmb_type.Value & "' and [Size]='" & cmb_size.Value & "' and [Brand]='" & cmb_brand.Value & "' and [Unit]='" & cmb_unit.Value & "'"
    Call DB.fill_text(txt_price, s, 1)
    Call set_name_code_id
    Call set_price
End Sub

Private Sub cmd_cancel_Click()
    Unload Me
End Sub

Private Sub cmd_clear_Click()
    cmb_type.ListIndex = -1
    cmb_size.ListIndex = -1
    cmb_brand.ListIndex = -1
    cmb_unit.ListIndex = -1
    txt_quantify.Value = 0
    txt_price.Value = 0
    lb_code.Caption = ""
    lb_name.Caption = ""
End Sub

Private Sub cmd_edit_Click()
    '------- lay du lieu thay doi cap nhat vao mang string ----------
    'ReDim s(10) As String
    's(0) = usf_order.id
    Call Algorithm.tranfer_data_an_item_to_string_array(s, Me)
    'Call Algorithm.print_string_arr(s)
    If usf_order.isPre = False Then
        '-------- cap nhat mang string vao array list -----------------
        Call Algorithm.edit_an_item(usf_order.arr_list, usf_order.index_edit, s)
        '--------- in array list vao sheet order ----------------------
        'Call Algorithm.print_array_list(usf_order.arr_list, ThisWorkbook.Worksheets("Order"))
        '--------- cap nhat order list trong user form --------------------------------
        Call Algorithm.update_item_to_order(usf_order.arr_list, usf_order.index_edit, usf_order)
    Else
        '-------- cap nhat mang string vao array list -----------------
        Call Algorithm.edit_an_item(usf_order.arr_list_previous, usf_order.index_edit, s)
        '--------- in array list vao sheet order ----------------------
        'Call Algorithm.print_array_list(usf_order.arr_list_previous, ThisWorkbook.Worksheets("Order"))
        '--------- cap nhat order list trong user form --------------------------------
        Call Algorithm.update_item_to_order(usf_order.arr_list_previous, usf_order.index_edit, usf_order)
    End If
    '----------cap nhat thanh tien trong list -------------------
    Call usf_order.to_total_price(usf_order)
    Unload Me
End Sub

Private Sub cmd_more_Click()
    Dim l As Integer
    l = Me.frm_input.Width
    If l = 700 Then
        Me.frm_input.Width = 535
        Me.Width = Me.Width - 165
        Me.cmd_edit.Left = Me.cmd_edit.Left - 165
        Me.cmd_clear.Left = Me.cmd_clear.Left - 165
        Me.cmd_cancel.Left = Me.cmd_cancel.Left - 165
    Else
        Me.frm_input.Width = 700
        Me.Width = Me.Width + 165
        Me.cmd_edit.Left = Me.cmd_edit.Left + 165
        Me.cmd_clear.Left = Me.cmd_clear.Left + 165
        Me.cmd_cancel.Left = Me.cmd_cancel.Left + 165
    End If
End Sub




Private Sub txt_price_Change()
    Dim s As String
    s = "Select [Price_o] from Products where [Class]='" & cmb_class.Value & "' and [Type]='" & cmb_type.Value & "' and [Size]='" & cmb_size.Value & "' and [Brand]='" & cmb_brand.Value & "' and [Unit]='" & cmb_unit.Value & "'"
    Dim tmp_v As Variant
    tmp_v = DB.read_data(s)
    If IsEmpty(tmp_v) Then
        lb_profit_loss_price.Caption = 0
    Else
        If IsNumeric(txt_price.Value) Then lb_profit_loss_price.Caption = CDbl(txt_price.Value) - CDbl(tmp_v(0, 0))
    End If
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    If usf_order.isPre = False Then
        s = Algorithm.acess_an_item_to_string_arr(usf_order.arr_list, usf_order.index_edit, 15)
    Else
        s = Algorithm.acess_an_item_to_string_arr(usf_order.arr_list_previous, usf_order.index_edit, 16)
    End If
    Me.frm_input.Width = 535
    Me.Width = 815 - 165
    Me.cmd_edit.Left = 726 - 165
    Me.cmd_clear.Left = 726 - 165
    Me.cmd_cancel.Left = 726 - 165
    
    Call set_combobox_default
    
    Call acess_product
End Sub
    
    
Private Sub acess_product()
    
    'Dim s() As String
    lb_id.Caption = s(12)
    lb_code.Caption = s(1)
    lb_name.Caption = s(4)
    txt_quantify.Value = s(7)
    'MsgBox " so luong: " & s(7)
    txt_price.Value = s(8)
End Sub

Private Sub set_combobox_default()
    Dim tmp_v As Variant
    Dim i As Integer
    Dim l As Integer
    Dim s1() As String
    Dim exc_s As String
    
    '--------- default combobox class to unit ------------
    tmp_v = DB.read_data("Select distinct [Class] from Products")
    Call Algorithm.set_combobox(cmb_class, tmp_v)
    cmb_class.Value = s(11)
    'Debug.Print s(11)
    
    If StrComp(cmb_class.Value, "", vbTextCompare) = 1 Then
        exc_s = "Select distinct [Type] from Products where [Class]='" & cmb_class.Value & "'"
        Call DB.fill_combobox(cmb_type, exc_s, 1)
    End If
    cmb_type.Value = s(2)
    
    If StrComp(cmb_type.Value, "", vbTextCompare) = 1 Then
        exc_s = "Select distinct [Size] from Products where [Class]='" & cmb_class.Value & "' and [Type]='" & cmb_type.Value & "'"
        Call DB.fill_combobox(cmb_size, exc_s, 2)
    End If
    cmb_size.Value = s(3)
    
    If StrComp(cmb_size.Value, "", vbTextCompare) = 1 Then
        exc_s = "Select distinct [Brand] from Products where [Class]='" & cmb_class.Value & "' and [Type]='" & cmb_type.Value & "' and [Size]='" & cmb_size.Value & "'"
        Call DB.fill_combobox(cmb_brand, exc_s, 3)
    End If
    cmb_brand.Value = s(5)
    
    If StrComp(cmb_brand.Value, "", vbTextCompare) = 1 Then
        exc_s = "Select distinct [Unit] from Products where [Class]='" & cmb_class.Value & "' and [Type]='" & cmb_type.Value & "' and [Size]='" & cmb_size.Value & "' and [Brand]='" & cmb_brand.Value & "'"
        Call DB.fill_combobox(cmb_unit, exc_s, 4)
    End If
    cmb_unit.Value = s(6)
    
    '--------- default combox box type price-----------
    exc_s = "Select distinct [Type] from Customers"
    Call DB.fill_combobox_customer(cmb_type_price, exc_s, 3)
    cmb_type_price.Value = s(10)
End Sub


Private Sub set_name_code_id()
    Dim exc_s As String
    Dim tmp_v As Variant
    If StrComp(cmb_unit.Value, "", vbTextCompare) > 0 Then
        exc_s = "Select [ID],[Code],[Name] from Products where [Class]='" & cmb_class.Value & "' and [Type]='" & cmb_type.Value & "' and [Size]='" & cmb_size.Value & "' and [Brand]='" & cmb_brand.Value & "' and [Unit]='" & cmb_unit.Value & "'"
        tmp_v = DB.read_data(exc_s)
        lb_id.Caption = tmp_v(0, 0)
        lb_code.Caption = tmp_v(1, 0)
        lb_name.Caption = tmp_v(2, 0)
    End If
    
End Sub

Private Sub set_price()
    'Dim exc_s As String
    'Dim tmp_v As Variant
    'If StrComp(cmb_unit.Value, "", vbTextCompare) > 0 And StrComp(cmb_type_price, "", vbTextCompare) > 0 Then
     '   exc_s = "Select [Price_1],[Price_2] from Products where [Class]='" & cmb_class.Value & "' and [Type]='" & cmb_type.Value & "' and [Size]='" & cmb_size.Value & "' and [Brand]='" & cmb_brand.Value & "' and [Unit]='" & cmb_unit.Value & "'"
      '  tmp_v = DB.read_data(exc_s)
     '   If cmb_type_price.ListIndex = 0 Then
      '      txt_price.Value = tmp_v(0, 0)
      '  ElseIf cmb_type_price.ListIndex = 1 Then
     '       txt_price.Value = tmp_v(1, 0)
     '   End If
   ' End If
    
    Dim exc_s As String
    Dim tmp_v As Variant
    If StrComp(cmb_unit.Value, "", vbTextCompare) > 0 And StrComp(cmb_type_price, "", vbTextCompare) > 0 Then
        exc_s = "Select [Price_o],[Price_1],[Price_2] from Products where [Class]='" & cmb_class.Value & "' and [Type]='" & cmb_type.Value & "' and [Size]='" & cmb_size.Value & "' and [Brand]='" & cmb_brand.Value & "' and [Unit]='" & cmb_unit.Value & "'"
        tmp_v = DB.read_data(exc_s)
        If cmb_type_price.ListIndex = 0 Then
            txt_price.Value = tmp_v(0, 0)
        ElseIf cmb_type_price.ListIndex = 1 Then
            txt_price.Value = tmp_v(1, 0)
        ElseIf cmb_type_price.ListIndex = 2 Then
            txt_price.Value = tmp_v(2, 0)
        End If
    End If
End Sub
