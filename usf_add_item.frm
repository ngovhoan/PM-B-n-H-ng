VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usf_add_item 
   Caption         =   "Add"
   ClientHeight    =   2736
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   16080
   OleObjectBlob   =   "usf_add_item.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usf_add_item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim str() As String

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


Public Sub hello_()
    MsgBox "hello"
End Sub

Private Sub cmb_type_price_Change()
    Call set_name_code_id
    Call set_price
End Sub

Private Sub cmb_unit_Change()
    Dim s As String
    cmb_type_price.Enabled = True
    If usf_order.cmb_customer_type.ListIndex = 1 Then
        s = "Select [Price_1] from Products where [Class]='" & cmb_class.Value & "' and [Type]='" & cmb_type.Value & "' and [Size]='" & cmb_size.Value & "' and [Brand]='" & cmb_brand.Value & "' and [Unit]='" & cmb_unit.Value & "'"
        cmb_type_price.ListIndex = 1
        Call DB.fill_text(txt_price, s, 1)
    ElseIf usf_order.cmb_customer_type.ListIndex = 2 Then
        s = "Select [Price_2] from Products where [Class]='" & cmb_class.Value & "' and [Type]='" & cmb_type.Value & "' and [Size]='" & cmb_size.Value & "' and [Brand]='" & cmb_brand.Value & "' and [Unit]='" & cmb_unit.Value & "'"
        cmb_type_price.ListIndex = 2
        Call DB.fill_text(txt_price, s, 2)
    Else
        s = "Select [Price_o] from Products where [Class]='" & cmb_class.Value & "' and [Type]='" & cmb_type.Value & "' and [Size]='" & cmb_size.Value & "' and [Brand]='" & cmb_brand.Value & "' and [Unit]='" & cmb_unit.Value & "'"
        cmb_type_price.ListIndex = 0
        Call DB.fill_text(txt_price, s, 3)
    End If
    
    
    Call set_name_code_id
    Call set_price
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

Public Sub cmd_add_Click()
    Dim str(15) As String
    str(0) = usf_order.id
    '------- Thuc hien lay du lieu tu user chuyen vao mang chuoi ------------
    Call Algorithm.tranfer_data_an_item_to_string_array(str, Me)
    '------- tao dong moi trong order---------------------------
    Call add_new_line_in_order(usf_order.x, str, usf_order.arr_list)
    
    '-------- Bien chay ------------------
    usf_order.x = usf_order.x + 26
    'usf_order.id = usf_order.id + 1
    
    Unload Me
End Sub

Public Sub add_new_line_in_order(pos As Integer, str() As String, a_l As ArrayList)
    Dim frm As MSForms.Frame
    Dim ctrl As MSForms.Control
    Dim lb As MSForms.Label
    Dim cmd As MSForms.CommandButton
    Dim cl As Long
    cl = RGB(236, 236, 236)
    
    '************************************* FRAME **************************
    'usf_order.Controls("frm_list_").SetFocus
    
    Set frm = usf_order.Controls("frm_list_").Add("Forms.Frame.1", "frm_item_" & usf_order.id, True)
    With frm
        .top = pos
        .Left = 6
        .Height = 26
        .Width = 585
        .BorderStyle = 1
        .BorderStyle = 0
    End With
    Dim frm_event As MouseMoveFrame
    Set frm_event = New MouseMoveFrame
    Set frm_event.frmEvent = frm
    usf_order.collection_frm.Add frm_event
    
    '********************************** LABEL *********************************
    '---------- Them Label thu tu sp-------------
    Set lb = usf_order.Controls("frm_item_" & usf_order.id).Add("Forms.Label.1", "lb_count_item_" & usf_order.id, True)
    Dim s As String
    s = CStr(usf_order.id)
    With lb
        .top = 3
        .Height = 20
        .Width = 24
        .Left = 6
        .BackColor = cl
        .TextAlign = fmTextAlignCenter
        .FontSize = 9
        .Caption = usf_order.id
    End With
    usf_order.col_count.Add lb
    '---------- Them Label ma sp -------------
    Set lb = usf_order.Controls("frm_item_" & usf_order.id).Add("Forms.Label.1", "lb_code_item_" & usf_order.id, True)
    With lb
        .top = 3
        .Height = 20
        .Width = 54
        .Left = 36
        .BackColor = cl
        .TextAlign = fmTextAlignCenter
        .FontSize = 9
        .Caption = str(1)
    End With
    usf_order.collection_lb.Add lb
    
    '---------- Them Label ten sp -------------
    Set lb = usf_order.Controls("frm_item_" & usf_order.id).Add("Forms.Label.1", "lb_name_item_" & usf_order.id, True)
    With lb
        .top = 3
        .Height = 20
        .Width = 120
        .Left = 96
        .BackColor = cl
        .TextAlign = fmTextAlignLeft
        .FontSize = 9
        .Caption = str(4)
    End With
    usf_order.collection_lb.Add lb
    '---------- Them Label xuat xu sp -------------
    Set lb = usf_order.Controls("frm_item_" & usf_order.id).Add("Forms.Label.1", "lb_brand_item_" & usf_order.id, True)
    With lb
        .top = 3
        .Height = 20
        .Width = 54
        .Left = 222
        .BackColor = cl
        .TextAlign = fmTextAlignCenter
        .FontSize = 9
        .Caption = str(5)
    End With
    usf_order.collection_lb.Add lb
    
    '---------- Them Label don vi sp -------------
    Set lb = usf_order.Controls("frm_item_" & usf_order.id).Add("Forms.Label.1", "lb_unit_item_" & usf_order.id, True)
    With lb
        .top = 3
        .Height = 20
        .Width = 42
        .Left = 282
        .BackColor = cl
        .TextAlign = fmTextAlignCenter
        .FontSize = 9
        .Caption = str(6)
    End With
    usf_order.collection_lb.Add lb
    
    '---------- Them Label so luong sp -------------
    Set lb = usf_order.Controls("frm_item_" & usf_order.id).Add("Forms.Label.1", "lb_quantify_item_" & usf_order.id, True)
    With lb
        .top = 3
        .Height = 20
        .Width = 48
        .Left = 330
        .BackColor = cl
        .TextAlign = fmTextAlignCenter
        .FontSize = 9
        .Caption = Format(CDbl(str(7)), "#.#0")
    End With
    usf_order.collection_lb.Add lb
    
    '---------- Them Label hien thi lai lo don gia -------------
    Set lb = usf_order.Controls("frm_item_" & usf_order.id).Add("Forms.Label.1", "lb_price_profit_loss_item_" & usf_order.id, True)
    With lb
        .top = 3
        .Height = 20
        .Width = 60
        .Left = 384
        
        .BackColor = cl
        .TextAlign = fmTextAlignRight
        .FontSize = 9
        .ForeColor = RGB(255, 0, 255)
        .Caption = Format(CDbl(str(13)), "#,##0")

    End With
    usf_order.collection_lb.Add lb
    
    
    '---------- Them Label don gia sp -------------
    Set lb = usf_order.Controls("frm_item_" & usf_order.id).Add("Forms.Label.1", "lb_price_item_" & usf_order.id, True)
    With lb
        .top = 3
        .Height = 20
        .Width = 60
        .Left = 384
        
        .BackColor = cl
        .TextAlign = fmTextAlignRight
        .FontSize = 9
        
        .Caption = Format(CDbl(str(8)), "#,##0")

    End With
    usf_order.collection_lb.Add lb
    
    '-------- Them Label hien thi lai lo sau khi thanh tien----------------
    Set lb = usf_order.Controls("frm_item_" & usf_order.id).Add("Forms.Label.1", "lb_total_profit_loss_item_" & usf_order.id, True)
    With lb
        .top = 3
        .Height = 20
        .Width = 65
        .Left = 450
        .BackColor = cl
        .TextAlign = fmTextAlignRight
        .FontSize = 9
        .ForeColor = RGB(255, 0, 255)
        .Caption = Format(CDbl(str(14)), "#,##0")

    End With
    usf_order.collection_lb.Add lb
    
    
    '---------- Them Label thanh tien sp -------------
    Set lb = usf_order.Controls("frm_item_" & usf_order.id).Add("Forms.Label.1", "lb_total_item_" & usf_order.id, True)
    With lb
        .top = 3
        .Height = 20
        .Width = 65
        .Left = 450
        .BackColor = cl
        .TextAlign = fmTextAlignRight
        .FontSize = 9
        
        .Caption = Format(CDbl(str(9)), "#,##0")

    End With
    usf_order.collection_lb.Add lb
    
    
    '***************************************CMD*******************************
    '---------- Them command sua ----------
    Set cmd = usf_order.Controls("frm_item_" & usf_order.id).Add("Forms.commandbutton.1", "cmd_edit_item_" & usf_order.id, True)
    With cmd
        .top = 2
        .Height = 20 '14
        .Width = 20 '25
        .Left = 520
        .BackStyle = fmBackStyleTransparent
        .Caption = "Edit"
        .Picture = LoadPicture("" & ThisWorkbook.Path & "\Images\icon_edit.jpg")
    End With
    Dim cmd_edit_item As EditItemButton
    Set cmd_edit_item = New EditItemButton
    Set cmd_edit_item.cmdEdit = cmd
    usf_order.collection_cmd.Add cmd_edit_item
    
    '---------- Them command xoa ----------
    Set cmd = usf_order.Controls("frm_item_" & usf_order.id).Add("Forms.commandbutton.1", "cmd_delete_item_" & usf_order.id, True)
    With cmd
        .top = 2
        .Height = 20
        .Width = 20
        .Left = 540
        .BackStyle = fmBackStyleTransparent
        .Caption = "Delete"
        .Picture = LoadPicture("" & ThisWorkbook.Path & "\Images\icon_delete.jpg")
    End With
    Dim cmd_delete_item As DeleteItemButton
    Set cmd_delete_item = New DeleteItemButton
    Set cmd_delete_item.cmdDelete = cmd
    usf_order.collection_cmd.Add cmd_delete_item
    
    '---------- them command insert ------------
    Set cmd = usf_order.Controls("frm_item_" & usf_order.id).Add("Forms.commandbutton.1", "cmd_insert_item_" & usf_order.id, True)
    With cmd
        .top = 2
        .Height = 20
        .Width = 20
        .Left = 560
        .BackStyle = fmBackStyleTransparent
        .Caption = "Delete"
        .Picture = LoadPicture("" & ThisWorkbook.Path & "\Images\icon_insert.jpg")
    End With
    Dim cmd_insert_item As InsertItemButton
    Set cmd_insert_item = New InsertItemButton
    Set cmd_insert_item.cmdInsert = cmd
    usf_order.collection_cmd.Add cmd_insert_item
    
    ' -------- Truyen tai du lieu vao array list -----------
    'Debug.Print "Gia tri 15 (id io): " + str(15)
    Call Algorithm.tranfer_data_to_array_list_(Me, a_l, str, usf_order.id)
    
    '-------- Bien chay ------------------
    'usf_order.X = usf_order.X + 26
    usf_order.id = usf_order.id + 1
    
    Call check_scrollbar_of_list
    Call usf_order.to_total_price(usf_order)
End Sub

Public Sub check_scrollbar_of_list()
    If usf_order.x > usf_order.Controls("frm_list_").Height Then
        With usf_order.Controls("frm_list_")
            .ScrollBars = fmScrollBarsVertical
            .ScrollHeight = usf_order.x + 10
        End With
    Else
        With usf_order.Controls("frm_list_")
            
            '.ScrollBars.Value = 0
            '.ScrollBars = fmScrollBarsNone
            '.ScrollHeight = usf_order.Controls("frm_list_").Height
            .ScrollHeight = usf_order.x
            .ScrollBars = fmScrollBarsNone
        End With
    End If
    
End Sub

Private Function to_amount(price As Double, quantify As Double) As Double
    to_amount = price * quantify
End Function

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

Private Sub cmd_more_Click()
    Dim l As Integer
    l = Me.frm_input.Width
    If l = 700 Then
        Me.frm_input.Width = 535
        Me.Width = Me.Width - 165
        Me.cmd_add.Left = Me.cmd_add.Left - 165
        Me.cmd_clear.Left = Me.cmd_clear.Left - 165
        Me.cmd_cancel.Left = Me.cmd_cancel.Left - 165
    Else
        Me.frm_input.Width = 700
        Me.Width = Me.Width + 165
        Me.cmd_add.Left = Me.cmd_add.Left + 165
        Me.cmd_clear.Left = Me.cmd_clear.Left + 165
        Me.cmd_cancel.Left = Me.cmd_cancel.Left + 165
    End If
End Sub

Private Sub frm_input_Click()

End Sub

Private Sub Label8_Click()

End Sub



Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

    Me.frm_input.Width = 535
    Me.Width = 815 - 165
    Me.cmd_add.Left = 726 - 165
    Me.cmd_clear.Left = 726 - 165
    Me.cmd_cancel.Left = 726 - 165
    Call set_combobox_default
    
End Sub


Private Sub set_combobox_default()
    
    Dim i As Integer
    Dim l As Integer
    Dim s1() As String
    Dim tmp_v As Variant
    '--------- default combobox class---------
    tmp_v = DB.read_data("Select distinct [Class] from Products")
    Call Algorithm.set_combobox(cmb_class, tmp_v)
    'cmb_class.ListIndex = 0
    '--------- default combox box type-----------
    cmb_type.Enabled = False
    '--------- default combox box size-----------
    cmb_size.Enabled = False
    '--------- default combox box brand-----------
    cmb_brand.Enabled = False
    '--------- default combox box unit-----------
    cmb_unit.Enabled = False
    '--------- default combox box type price-----------
    Dim exc_s As String
    exc_s = "Select distinct [Type] from Customers order by [Type] ASC"
    Call DB.fill_combobox_customer(cmb_type_price, exc_s, 3)
    cmb_type_price.Value = usf_order.cmb_customer_type.Value
    cmb_type_price.Enabled = False
    
    txt_quantify.Value = 0
    txt_price.Value = 0
End Sub


'============= set name and code and id =============
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
