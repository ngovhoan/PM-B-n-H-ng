VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usf_order 
   Caption         =   "Order"
   ClientHeight    =   10320
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   17544
   OleObjectBlob   =   "usf_order.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usf_order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public collection_frm As New Collection
Public collection_lb As New Collection
Public collection_cmd As New Collection
Public col_count As New Collection

Public arr_list As ArrayList
Public arr_list_previous As ArrayList
Public arr_list_tmp As ArrayList
    
Public x As Integer
Public id As Integer
Public index_edit As Integer
Public isPre As Boolean
Public isSeeingPL As Boolean ' bien kiem tra xem co hien thi lai lo

Dim cl_new As Long
Dim cl_old As Long
Dim i_more As Integer
Dim vat As Double

Dim arr_list_inventory As ArrayList
Dim arr_list_history As ArrayList

Private Sub cb_logistic_Change()
    If cb_logistic.Value = True Then
        txt_logistic.Visible = True
        lb_vnd.Visible = True
    Else
        txt_logistic.Visible = False
        lb_vnd.Visible = False
    End If
End Sub

Private Sub cb_old_debt_Click()
    If cb_old_debt.Value = True Then
        Call to_debt_old
    Else
        lb_old_debt.Caption = 0
        lb_total_amount.Caption = Format(CDbl(lb_after_tax.Caption) + CDbl(lb_old_debt.Caption), "#,##0")
        If txt_pay_prepay.Value = "" Then
            lb_pay.Caption = Format(0, "#,##0")
        Else
            lb_pay.Caption = Format(CDbl(txt_pay_prepay.Value), "#,##0")
        End If
        lb_refund.Caption = Format(CDbl(lb_total_amount.Caption) - CDbl(lb_pay.Caption), "#,##0")
    End If
End Sub

Private Sub cb_vat_Change()
    If cb_vat.Value = True Then
        cmb_vat.Visible = True
        vat = cmb_vat.Value / 100
    Else
        cmb_vat.Visible = False
        vat = 0
    End If
    Call update_widget
End Sub

Private Sub cmb_customer_Change()
    Dim tmp_v As Variant
    Dim exc_s As String
    cmb_customer_id.Enabled = True
    exc_s = "select [ID_c] from Customers where [Name_c] = '" & cmb_customer.Value & "'"
    Call DB.fill_combobox_customer(cmb_customer_id, exc_s, 4)
    'cmb_customer_id.ListIndex = 0
    If cmb_customer_id.ListIndex = -1 And cmb_customer_id.ListCount > 0 Then
        cmb_customer_id.ListIndex = 0
    End If
    If cmb_customer.ListIndex > -1 Then
        exc_s = "Select [Name_c],[Number],[Address],[Type] from Customers where [ID_c]=" & CInt(cmb_customer_id.Value) & " and [Name_c] = '" & cmb_customer.Value & "'"
        tmp_v = DB.read_data(exc_s)
        txt_customer_name.Value = tmp_v(0, 0)
        txt_customer_phone = tmp_v(1, 0)
        txt_customer_address = tmp_v(2, 0)
        cmb_customer_type = tmp_v(3, 0)
        
        exc_s = "Select [Project] from Orders where [Customer_ID]=" & CInt(cmb_customer_id.Value)
        Call DB.fill_combobox_customer(cmb_order_project, exc_s, 5)
    End If
    
    'Call cmd_all_history_Click
End Sub

Private Sub cmb_customer_id_Change()
    Dim exc_s As String
    If isPre = True Then
        'lst_customer_history.ListIndex = -1
        lst_customer_history.Clear
        arr_list_history.Clear
        Call cmd_back_to_present_order_Click
        cmd_show_old_order.Enabled = False
        Call clear_label_info_history
        isPre = False
    End If
    If StrComp(cmb_customer_id.Value, "", vbTextCompare) > 0 Then
        exc_s = "Select [Project] from Orders where [Customer_ID]=" & CInt(cmb_customer_id.Value) & ""
        Call DB.fill_combobox_customer(cmb_order_project, exc_s, 5)
        
        '**** Cap nhat lis history****
        Set arr_list_history = New ArrayList
        lst_customer_history.Clear
        If check_is_order(CInt(cmb_customer_id.Value)) Then
            Dim s() As String
            exc_s = "Select [Code], [Project], [Time],[Total],[Pay], [Note],[ID] from Orders where [Customer_ID]=" & CInt(cmb_customer_id.Value)
            Call DB.readData(arr_list_history, exc_s)
            s = covert_array_list_to_string_array(arr_list_history)
            lst_customer_history.List = s
        End If
        
    End If
    If StrComp(cmb_customer_id.Value, "", vbTextCompare) Then
        Call cmd_all_history_Click
    End If
End Sub

Private Function check_is_order(id_ As Integer) As Boolean
    Dim exc_s As String
    Dim tmp_v As Variant
    Dim vl As Boolean
    vl = False
    exc_s = "select [Customer_ID] from Orders"
    tmp_v = DB.read_data(exc_s)
    Dim i_ As Integer
    For i_ = 0 To UBound(tmp_v, 2)
        If CInt(tmp_v(0, i_)) = id_ Then
            vl = True
        End If
    Next
    check_is_order = vl
End Function


Private Sub clear_label_info_history()
    lb_order_id_history.Caption = ""
    lb_project_history.Caption = ""
    lb_note_history.Caption = ""
    lb_total_amount_order_history.Caption = ""
    lb_pay_order_history.Caption = ""
    lb_time_order_history.Caption = ""
    lb_code_order_history.Caption = ""
End Sub

Private Sub cmb_ex_im_Change()
    If cmb_ex_im.ListIndex = -1 Then
        cmd_save.Enabled = False
    Else
        cmd_save.Enabled = True
    End If
End Sub

Private Sub cmb_new_Click()
    cmb_customer.ListIndex = -1
    cmb_customer_id.ListIndex = -1
    cmb_customer_type.ListIndex = -1
    txt_customer_name.Value = ""
    txt_customer_phone.Value = ""
    txt_customer_address.Value = ""
    cmb_customer_id.Enabled = False
    
    lst_customer_history.Clear
    arr_list_history.Clear
    If isPre = True Then
        Call cmd_back_to_present_order_Click
        Call clear_label_info_history
        isPre = False
    End If
End Sub

Private Sub cmb_order_code_Change()

End Sub

Private Sub cmb_order_project_Change()
    Dim tmp_v As Variant
    Dim exc_s As String
    If StrComp(cmb_order_project.Value, "", vbTextCompare) > 0 Then
        exc_s = "Select [Project],[Address] from Orders where [Customer_ID]=" & CInt(cmb_customer_id.Value) & " and [Project]='" & cmb_order_project.Value & "'"
        tmp_v = DB.read_data(exc_s)
        txt_order_project.Value = tmp_v(0, 0)
        txt_order_project_address.Value = tmp_v(1, 0)
    End If
End Sub

Private Sub cmb_search_delete_Click()
    txt_search_inventory.Value = ""
End Sub

Private Sub cmb_type_inventory_Change()
    lst_inventory.ListIndex = -1
    arr_list_inventory.Clear
    Dim exc_s As String
    Dim tmp_v As Variant, tmp_v_ As Variant
    If StrComp(cmb_type_inventory.Value, "", vbTextCompare) = 0 Then
        Exit Sub
    End If
    exc_s = "Select [Code],[Name], [Unit] from Products where [Type]='" & cmb_type_inventory.Value & "'"
    'Call DB.readData(arr_list_inventory, "Select [Code],[Name] from Inventory_Item where [Type]='" & cmb_type_inventory.Value & "'")
    tmp_v = DB.read_data(exc_s)
    
    Dim tmp_s(3) As String
    Dim i_ As Integer
    For i_ = 0 To UBound(tmp_v, 2)
        'exc_s = "Select Sum(IO_Products.[Quantify]) FROM Products INNER JOIN IO_Products ON Products.[ID] = IO_Products.[Product_ID] where IO_Products.[Product_ID]=" & CInt(tmp_v(2, i_)) & " and Products.[Type]='" & cmb_type_inventory.Value & "'"
        'tmp_v_ = DB.read_data(exc_s)
        tmp_s(0) = tmp_v(0, i_)
        tmp_s(1) = tmp_v(1, i_)
        'tmp_s(2) = tmp_v_(0, 0)
        tmp_s(2) = tmp_v(2, i_)
        arr_list_inventory.Add tmp_s
    Next
    
    lst_inventory.Clear
    Dim s() As String
    s = covert_array_list_to_string_array(arr_list_inventory)
    lst_inventory.List = s
End Sub


Private Sub cmb_vat_Change()
    vat = cmb_vat.Value / 100
    Call update_widget
End Sub

Private Sub cmd_add_item_Click()
    'usf_order.Controls("frm_list_").TabIndex = 100
    usf_add_item.Show (vbModeless)
End Sub


Private Sub cmd_all_history_Click()
    cmd_all_history.BackColor = cl_new
    cmd_loading_history.BackColor = cl_old
    cmd_finished_history.BackColor = cl_old
    Call set_arr_list_history(1)
End Sub

Private Sub cmd_change_tab_Click()
    Dim i As Integer
    i = mlt_page.Value
    If i < 2 Then
        i = i + 1
    Else
        i = 0
    End If
    mlt_page.Value = i
    Select Case i
        Case 0
            lb_flat_1.BackColor = cl_new
            lb_flat_2.BackColor = cl_old
            lb_flat_3.BackColor = cl_old
        Case 1
            lb_flat_1.BackColor = cl_old
            lb_flat_2.BackColor = cl_new
            lb_flat_3.BackColor = cl_old
        Case Else
            lb_flat_1.BackColor = cl_old
            lb_flat_2.BackColor = cl_old
            lb_flat_3.BackColor = cl_new
    End Select
    'lb_notification.Caption = ""
End Sub

Private Sub cmd_clean_list_Click()
    Dim ans As Integer
    
    ans = MsgBox("Ban co muon xoa danh sach?", vbQuestion + vbYesNo)
    If ans = vbYes Then
        Dim ctrl As MSForms.Control
        For Each ctrl In Me.Controls
            If InStr(1, ctrl.name, "frm_item", 1) > 0 Then
                
                usf_order.Controls.Remove (ctrl.name)
                arr_list.Clear
                'Call Algorithm.clear_store(ThisWorkbook.Worksheets("Order"))
                index_edit = 0
                x = 6
                Call usf_add_item.check_scrollbar_of_list
            End If
        Next
        If isPre Then
            Call cmd_back_to_present_order_Click
        End If
        txt_order_note.Value = ""
        txt_pay_prepay.Value = 0
    End If
End Sub

Private Sub cmd_finished_history_Click()
    cmd_all_history.BackColor = cl_old
    cmd_loading_history.BackColor = cl_old
    cmd_finished_history.BackColor = cl_new
    Call set_arr_list_history(3)
End Sub

Private Sub cmd_loading_history_Click()
    cmd_all_history.BackColor = cl_old
    cmd_loading_history.BackColor = cl_new
    cmd_finished_history.BackColor = cl_old
    Call set_arr_list_history(2)
End Sub

Private Sub cmd_print_Click()
    Dim l As Integer
    Dim i As Integer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Order")
    Call Algorithm.clean_row_in_excel(ws)
    If isPre = False Then
        l = arr_list.Count - 1
        For i = 0 To l
            Call Algorithm.insert_new_row_in_excel(ws, arr_list, 6 + i, i)
        Next
    Else
        l = arr_list_previous.Count - 1
        For i = 0 To l
            Call Algorithm.insert_new_row_in_excel(ws, arr_list_previous, 6 + i, i)
        Next
    End If
    For i = 0 To l
        ws.Range("A" & i + 6).Value = i + 1
    Next
    ws.Range("G" & l + 7).Formula = "=SUM(G6:G" & l + 6 & ")"
    ws.Range("G" & l + 8).Value = lb_old_debt.Caption
    ws.Range("G" & l + 9).Formula = "=G" & l + 7 & " + G" & l + 8
    ws.Range("G" & l + 10).Value = lb_pay.Caption
    ws.Range("G" & l + 11).Formula = "=G" & l + 9 & " - G" & l + 10
    ws.Range("G" & l + 7 & ":G" & l + 11).NumberFormat = "#,##0"
    ws.Range("B6:B" & l + 6).HorizontalAlignment = xlLeft
    ws.Range("E6:E" & l + 6).NumberFormat = "0.0"
    ws.Range("F6:G" & l + 6).HorizontalAlignment = xlRight
    ws.Range("F6:G" & l + 6).NumberFormat = "#,##0"
    With ws.Range("A6:G" & l + 6)
        .Borders.LineStyle = XlLineStyle.xlContinuous
        .Borders.Color = RGB(255, 0, 0)
        .Font.FontStyle = "Regular"
        .Font.Size = 12
        .WrapText = True
    End With
    ws.Shapes("Rectangle 6").TextFrame.Characters.Text = cmb_customer_id.Value
    ws.Shapes("Rectangle 8").TextFrame.Characters.Text = txt_customer_name.Value
    ws.Shapes("Rectangle 9").TextFrame.Characters.Text = txt_customer_phone.Value
    ws.Shapes("Rectangle 12").TextFrame.Characters.Text = txt_customer_address.Value
    If isPre Then
        ws.Shapes("Rectangle 11").TextFrame.Characters.Text = lb_code_order_history.Caption
    Else
        ws.Shapes("Rectangle 11").TextFrame.Characters.Text = lb_order_code.Caption
    End If
End Sub

Private Sub set_count_of_order()
    Dim ctrl As MSForms.Control
    Dim tmp As Integer
    Dim top_ As Integer
    
    For Each ctrl In Me.Controls
        If InStr(1, CStr(ctrl.name), "frm_item", 1) > 0 Then
            top_ = ctrl.top
            'MsgBox (ctrl.Name)
            tmp = CInt(Mid(CStr(ctrl.name), 10, 3))
            Me.Controls("lb_count_item_" & tmp).Caption = ((top_ - 6) / 26 + 1)
        End If
    Next
End Sub

Private Sub cmd_print_pdf_Click()

    Call cmd_print_Click

    Dim l As Integer
    'Dim i As Integer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Order")
    l = ws.Range("B" & Application.Rows.Count).End(xlUp).Row
    Dim path_s As String
    If isPre Then
        path_s = ThisWorkbook.Path & "\PDF\" & lb_code_order_history.Caption & ".pdf"
    Else
        path_s = ThisWorkbook.Path & "\PDF\" & lb_order_code.Caption & "(nhap).pdf"
    End If
    ws.PageSetup.PrintArea = "A1:G" & l + 1
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path_s, Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    
End Sub

Private Sub cmd_profit_loss_Click()
    Dim PL_total As Double
    PL_total = 0
    Dim ctrl As MSForms.Control
    If isSeeingPL Then
        For Each ctrl In Me.Controls
            If InStr(ctrl.name, "lb_price_item_") > 0 Or InStr(ctrl.name, "lb_total_item_") > 0 Then
                ctrl.Visible = True
            End If
        Next
        lb_PL_total.Visible = False
        isSeeingPL = False
    Else
        For Each ctrl In Me.Controls
            If InStr(ctrl.name, "lb_price_item_") > 0 Or InStr(ctrl.name, "lb_total_item_") > 0 Then
                ctrl.Visible = False
            End If
            If InStr(ctrl.name, "lb_total_profit_loss_item_") > 0 Then PL_total = PL_total + CDbl(Me.Controls(ctrl.name).Caption)
        Next
        lb_PL_total.Caption = Format(PL_total, "#,##0")
        lb_PL_total.Visible = True
        isSeeingPL = True
    End If
End Sub

Private Sub cmd_save_Click()
    'If IsEmpty(arr_list) And IsEmpty(arr_list_history) Then
    If arr_list.Count < 0 And arr_list_history.Count < 0 Then
        MsgBox "Ban chua nhap san pham"
    Else
        Dim i_ As Integer
        Dim dtToday As Date
        dtToday = Now()
    
        Dim exc_s As String
        
        '---- Neu don moi -----
        If isPre = False Then
            Call create_update_customer
            Call create_update_order
            For i_ = 0 To arr_list.Count - 1
                exc_s = "Insert into IO_Products ([Product_ID],[Count],[Price],[Quantify],[Type],[Date],[Order_ID],[Note]) values (" & arr_list.Item(i_)(12) & "," & arr_list.Item(i_)(0) & "," & arr_list.Item(i_)(8) & "," & arr_list.Item(i_)(7) & ",'" & cmb_ex_im.Value & "','" & CDate(CDbl(dtToday)) & "'," & lb_order_id.Caption & ",'" & txt_customer_name.Value & "')"
                Call DB.insert_row(exc_s)
            Next
            'lb_notification.Caption = "New Order has been Added!"
        '----- Neu don cu -------
        Else
            For i_ = 0 To arr_list_previous.Count - 1
                exc_s = "Update IO_Products set [Product_ID]=" & CInt(arr_list_previous.Item(i_)(12)) & ",[Price]=" & CDbl(arr_list_previous.Item(i_)(8)) & ",[Quantify]=" & CDbl(arr_list_previous.Item(i_)(7)) & ",[Date]='" & CDate(dtToday) & "',[Note]='" & txt_customer_name.Value & "' where [ID]=" & CInt(arr_list_previous.Item(i_)(15)) & ""
                Call DB.insert_row(exc_s)
            Next
            'Call to_total_price(usf_order)
            exc_s = "Update Orders set [Total]=" & CLng(lb_after_tax.Caption) & ",[Pay]=" & CLng(lb_pay.Caption) & " where [ID]=" & CInt(lb_order_id_history.Caption)
            Call DB.insert_row(exc_s)
            Dim pt As Integer
            pt = lst_customer_history.ListIndex
            If cmd_all_history.BackColor = cl_new Then
                Call cmd_all_history_Click
                'arr_list_history.ListIndex = pivot
                lst_customer_history.ListIndex = pt
            ElseIf cmd_loading_history.BackColor = cl_new Then
                Call cmd_loading_history_Click
                Call cmd_back_to_present_order_Click
                lst_customer_history.ListIndex = pt - 1
            Else
                Call cmd_finished_history_Click
                Call cmd_back_to_present_order_Click
                lst_customer_history.ListIndex = pt - 1
            End If
            
            lb_notification.Caption = "Order History has been Updated!"
        End If
    End If
    Call set_new_order_id_code
    'Call cmd_clean_list_Click
End Sub

Private Sub create_update_customer()
    Dim exc_s As String
    If StrComp(txt_customer_name.Value, "", vbTextCompare) > 0 And StrComp(cmb_customer.Value, "", vbTextCompare) = 0 Then
        exc_s = "insert into Customers ([Name_c],[Number],[Address],[Type],[Note]) values ('" & txt_customer_name & "','" & txt_customer_phone & "','" & txt_customer_address & "','" & cmb_customer_type & "','')"
        Call DB.insert_row(exc_s)
    End If
End Sub

Private Sub create_update_order()
    Dim exc_s As String
    Dim tmp_v As Variant
    Dim id_c As Integer
    exc_s = "select [ID_c] from Customers"
    tmp_v = DB.read_data(exc_s)
    'MsgBox UBound(tmp_v, 2)
    If StrComp(cmb_customer_id.Value, "", vbTextCompare) = 0 Then
        id_c = UBound(tmp_v, 2) + 1
    Else
        id_c = CInt(cmb_customer_id.Value)
    End If
    exc_s = "insert into Orders ([Code],[Customer_ID],[Project],[Address],[Time],[Total],[Pay],[Type],[Note]) values (" & lb_order_code.Caption & "," & id_c & ",'" & txt_order_project.Value & "','" & txt_order_project_address.Value & "','" & txt_order_time.Value & "'," & CLng(lb_after_tax.Caption) & "," & CLng(lb_pay) & ",'" & cmb_ex_im.Value & "','" & txt_order_note & "')"
    Call DB.insert_row(exc_s)
    cmb_customer_id.Enabled = True
    cmb_customer_id.AddItem UBound(tmp_v, 2) + 1
    cmb_customer_id.ListIndex = 0
    exc_s = "Select [ID] from Orders"
    tmp_v = DB.read_data(exc_s)
    lb_notification.Caption = "New Order " & CInt(lb_order_id.Caption) & " has been Added!"
End Sub

Private Sub cmd_search_delete_Click()
    txt_search_inventory.Value = " "
End Sub

Private Sub cmd_show_full_list_Click()
    If i_more = 0 Then
        i_more = 1
        cmd_show_full_list.Picture = LoadPicture("" & ThisWorkbook.Path & "\Images\icon_arrow_up.jpg")
        frm_order.Height = frm_order.Height + 143
        Me.Controls("frm_list_").Height = Me.Controls("frm_list_").Height + 145
    Else
        i_more = 0
        cmd_show_full_list.Picture = LoadPicture("" & ThisWorkbook.Path & "\Images\icon_arrow_down.jpg")
        frm_order.Height = frm_order.Height - 143
        Me.Controls("frm_list_").Height = Me.Controls("frm_list_").Height - 145
    End If
    Call usf_add_item.check_scrollbar_of_list
End Sub


Public Sub to_total_price(usf As MSForms.UserForm)
    
    Dim ctrl As MSForms.Control
    Dim tt As Double
    tt = 0
    For Each ctrl In usf.Controls
        If InStr(1, ctrl.name, "lb_total_item", 1) > 0 Then
            tt = tt + CDbl(ctrl.Caption)
        End If
    Next
    'Me.lb_total_price.Caption = tt
    Me.lb_total_price.Caption = Format(tt, "#,##0")
    
    Call to_debt_old
    
    Call update_widget
End Sub

Private Sub to_debt_old()
    Dim exc_s  As String
    Dim total_exp As Double, total_imp As Double, pay_exp As Double, pay_imp As Double
    total_exp = 0
    total_imp = 0
    pay_exp = 0
    pay_imp = 0
    Dim tmp_v As Variant
    If StrComp(lb_order_id_history.Caption, "", vbTextCompare) = 0 Then
        exc_s = "Select sum (Total) from Orders Where [Customer_ID]=" & CInt(cmb_customer_id.Value) & " and [Type]='EXP'"
        tmp_v = DB.read_data(exc_s)
        If IsNull(tmp_v(0, 0)) = False Then total_exp = CDbl(tmp_v(0, 0))
        exc_s = "Select sum (Total) from Orders Where [Customer_ID]=" & CInt(cmb_customer_id.Value) & " and [Type]='IMP'"
        tmp_v = DB.read_data(exc_s)
        If IsNull(tmp_v(0, 0)) = False Then total_imp = CDbl(tmp_v(0, 0))
        exc_s = "Select sum (Pay) from Orders Where [Customer_ID]=" & CInt(cmb_customer_id.Value) & " and [Type]='IMP'"
        tmp_v = DB.read_data(exc_s)
        If IsNull(tmp_v(0, 0)) = False Then pay_imp = CDbl(tmp_v(0, 0))
        exc_s = "Select sum (Pay) from Orders Where [Customer_ID]=" & CInt(cmb_customer_id.Value) & " and [Type]='IMP'"
        tmp_v = DB.read_data(exc_s)
        If IsNull(tmp_v(0, 0)) = False Then pay_imp = CDbl(tmp_v(0, 0))
    Else
        exc_s = "Select sum (Total) from Orders Where [Customer_ID]=" & CInt(cmb_customer_id.Value) & " and [Type]='EXP' and [ID] <" & CInt(lb_order_id_history.Caption)
        tmp_v = DB.read_data(exc_s)
        If IsNull(tmp_v(0, 0)) = False Then total_exp = CDbl(tmp_v(0, 0))
        exc_s = "Select sum (Total) from Orders Where [Customer_ID]=" & CInt(cmb_customer_id.Value) & " and [Type]='IMP' and [ID] <" & CInt(lb_order_id_history.Caption)
        tmp_v = DB.read_data(exc_s)
        If IsNull(tmp_v(0, 0)) = False Then total_imp = CDbl(tmp_v(0, 0))
        exc_s = "Select sum (Pay) from Orders Where [Customer_ID]=" & CInt(cmb_customer_id.Value) & " and [Type]='IMP' and [ID] <" & CInt(lb_order_id_history.Caption)
        tmp_v = DB.read_data(exc_s)
        If IsNull(tmp_v(0, 0)) = False Then pay_imp = CDbl(tmp_v(0, 0))
        exc_s = "Select sum (Pay) from Orders Where [Customer_ID]=" & CInt(cmb_customer_id.Value) & " and [Type]='IMP' and [ID] <" & CInt(lb_order_id_history.Caption)
        tmp_v = DB.read_data(exc_s)
        If IsNull(tmp_v(0, 0)) = False Then pay_imp = CDbl(tmp_v(0, 0))
    End If
    'lb_old_debt.Caption = total_exp - pay_exp + pay_imp - total_imp
    lb_old_debt.Caption = Format(total_exp - pay_exp + pay_imp - total_imp, "#,##0")
    lb_total_amount.Caption = Format(CDbl(lb_after_tax.Caption) + CDbl(lb_old_debt.Caption), "#,##0")
    If txt_pay_prepay.Value = "" Then
        lb_pay.Caption = Format(0, "#,##0")
    Else
        lb_pay.Caption = Format(CDbl(txt_pay_prepay.Value), "#,##0")
    End If
    lb_refund.Caption = Format(CDbl(lb_total_amount.Caption) - CDbl(lb_pay.Caption), "#,##0")
End Sub

Private Sub set_icon_cmd()
    cmd_add_item.Picture = LoadPicture("" & ThisWorkbook.Path & "\Images\icon_add.jpg")
    cmd_show_full_list.Picture = LoadPicture("" & ThisWorkbook.Path & "\Images\icon_arrow_down.jpg")
    cmd_clean_list.Picture = LoadPicture("" & ThisWorkbook.Path & "\Images\icon_clean.jpg")
    cmd_change_tab.Picture = LoadPicture("" & ThisWorkbook.Path & "\Images\icon_change_tab.jpg")
    cmd_show_old_order.Picture = LoadPicture("" & ThisWorkbook.Path & "\Images\icon_details.jpg")
    cmd_back_to_present_order.Picture = LoadPicture("" & ThisWorkbook.Path & "\Images\icon_close_details.jpg")
    cmb_new.Picture = LoadPicture("" & ThisWorkbook.Path & "\Images\icon_blank_page.jpg")
    cmd_search_delete.Picture = LoadPicture("" & ThisWorkbook.Path & "\Images\icon_delete.jpg")
End Sub

Private Sub set_widget_default()
    Dim i As Integer
    For i = 0 To 10
        cmb_vat.AddItem (i)
    Next
    cb_vat.Value = False
    cmb_vat.ListIndex = 0
    cmb_vat.Visible = False
    cb_logistic.Value = False
    txt_logistic.Visible = False
    lb_vnd.Visible = False
    
    '-------- thanh toan -------------
    
    Call update_widget
    Call cb_vat_Change
    Call cb_logistic_Change
End Sub

Private Sub update_widget()
    
    lb_taxes.Caption = Format(CDbl(lb_total_price.Caption) * vat, "#,##0")
    lb_after_tax.Caption = Format(CDbl(lb_total_price.Caption) + CDbl(lb_taxes.Caption), "#,##0")
    lb_old_debt.Caption = Format(0, "#,##0")
    lb_total_amount.Caption = Format(CDbl(lb_after_tax.Caption) + CDbl(lb_old_debt.Caption), "#,##0")
    If txt_pay_prepay.Value = "" Then
        lb_pay.Caption = Format(0, "#,##0")
    Else
        lb_pay.Caption = Format(CDbl(txt_pay_prepay.Value), "#,##0")
    End If
    lb_refund.Caption = Format(CDbl(lb_total_amount.Caption) - CDbl(lb_pay.Caption), "#,##0")
End Sub

Private Sub set_combobox_customer_default()
    Dim exc_s As String
    exc_s = "Select [Name_c] from Customers"
    Call DB.fill_combobox_customer(cmb_customer, exc_s, 1)
    exc_s = "Select [Code] from Orders"
    Call DB.fill_combobox_customer(cmb_order_code, exc_s, 2)
    exc_s = "Select distinct [Type] from Customers order by [Type] ASC"
    Call DB.fill_combobox_customer(cmb_customer_type, exc_s, 3)
    cmb_customer_type.ListIndex = 1
    exc_s = "Select [ID_c] from Customers order by [ID_c] ASC"
    Call DB.fill_combobox_customer(cmb_customer_id, exc_s, 4)
    cmb_customer_id.Value = "1"
End Sub

Private Sub set_combobox_order_default()
    'Dim exc_s As String
    'exc_s = "Select [Project] from Orders where [Customer_ID]=" & cmb_customer_id.Value
    'Call DB.fill_combobox_customer(cmb_order_project, exc_s, 5)
    cmb_ex_im.AddItem "EXP"
    cmb_ex_im.AddItem "IMP"
    cmb_ex_im.ListIndex = -1
End Sub

Private Sub cmd_back_to_present_order_Click()
    isPre = False
    cmd_add_item.Enabled = True
    cmd_clean_list.Enabled = True
    cmd_show_old_order.Enabled = True
    cmd_back_to_present_order.Enabled = False
    Dim ctrl As MSForms.Control
    For Each ctrl In Me.Controls
        If InStr(1, ctrl.name, "frm_item", 1) > 0 Then
            
            usf_order.Controls.Remove (ctrl.name)
            arr_list_previous.Clear
            'Call Algorithm.clear_store(ThisWorkbook.Worksheets("Order"))
            index_edit = 0
            x = 6
            Call usf_add_item.check_scrollbar_of_list
        End If
    Next
    cmb_ex_im.ListIndex = -1
    txt_pay_prepay.Value = 0
    Dim i_ As Integer
    Dim s_(18) As String
    'Debug.Print arr_list_tmp.Count
    For i_ = 0 To arr_list_tmp.Count - 1
        s_(0) = arr_list_tmp.Item(i_)(0) ' tt
        s_(1) = arr_list_tmp.Item(i_)(1) ' code
        s_(2) = arr_list_tmp.Item(i_)(2) ' type
        s_(3) = arr_list_tmp.Item(i_)(3) ' size
        s_(4) = arr_list_tmp.Item(i_)(4) ' name
        s_(5) = arr_list_tmp.Item(i_)(5) ' brand
        s_(6) = arr_list_tmp.Item(i_)(6) ' unit
        s_(7) = arr_list_tmp.Item(i_)(7) ' quantify
        s_(8) = arr_list_tmp.Item(i_)(8) ' price
        s_(9) = CStr(CDbl(s_(7)) * CDbl(s_(8))) ' total price
        s_(10) = "" ' type price
        s_(11) = arr_list_tmp.Item(i_)(10) ' class
        s_(12) = arr_list_tmp.Item(i_)(11) ' id
        s_(13) = arr_list_tmp.Item(i_)(13) ' profit - lot price
        s_(14) = arr_list_tmp.Item(i_)(14) ' profit - lot total
        Call usf_add_item.add_new_line_in_order(x, s_, arr_list)
        x = x + 26
    Next
    lb_notification.Caption = ""
End Sub


Private Sub cmd_show_old_order_Click()
    isPre = True
    
    cmd_add_item.Enabled = False
    cmd_clean_list.Enabled = False
    cmd_show_old_order.Enabled = False
    cmd_back_to_present_order.Enabled = True
    Set arr_list_tmp = Algorithm.coppy_array_list(arr_list)
    arr_list.Clear
    '**** Xoa hien thi du lieu cu **********
    Dim ctrl As MSForms.Control
    For Each ctrl In Me.Controls
        If InStr(1, ctrl.name, "frm_item", 1) > 0 Then
            usf_order.Controls.Remove (ctrl.name)
            index_edit = 0
            x = 6
            'Call usf_add_item.check_scrollbar_of_list
        End If
    Next
    Call usf_add_item.check_scrollbar_of_list
    
    cmb_ex_im.Value = lb_ex_im_order.Caption
    txt_pay_prepay.Value = lb_pay_order_history.Caption
    Dim exc_s As String
    Dim tmp_v As Variant
    exc_s = "SELECT IO_Products.[Count], Products.[Code],Products.[Type],Products.[Size], Products.[Name], Products.[Brand], Products.[Unit], IO_Products.[Quantify], IO_Products.[Price], Products.[Class], Products.[ID],IO_Products.[ID], Products.[Price_o] FROM Products INNER JOIN (Orders INNER JOIN IO_Products ON Orders.[ID] = IO_Products.[Order_Id]) ON Products.[ID] = IO_Products.[Product_ID] WHERE (((Orders.[ID])=" & CInt(lb_order_id_history.Caption) & "))"
    tmp_v = DB.read_data(exc_s)
    Dim i_ As Integer
    Dim s_(16) As String
    Set arr_list_previous = New ArrayList
    If IsEmpty(tmp_v) = False Then
        For i_ = 0 To UBound(tmp_v, 2)
            s_(0) = tmp_v(0, i_) ' tt
            s_(1) = tmp_v(1, i_) ' code
            s_(2) = tmp_v(2, i_) ' type
            s_(3) = tmp_v(3, i_) ' size
            s_(4) = tmp_v(4, i_) ' name
            s_(5) = tmp_v(5, i_) ' brand
            s_(6) = tmp_v(6, i_) ' unit
            s_(7) = tmp_v(7, i_) ' quantify
            s_(8) = tmp_v(8, i_) ' price
            s_(9) = CStr(CDbl(s_(7)) * CDbl(s_(8))) ' total price
            s_(10) = "" ' type price
            s_(11) = tmp_v(9, i_) ' class
            s_(12) = tmp_v(10, i_) ' id
            s_(13) = s_(8) - tmp_v(12, i_)   ' profit - loss price
            s_(14) = s_(13) * s_(7)     ' profit - loss total
            s_(15) = tmp_v(11, i_)  'id io
            
            Call usf_add_item.add_new_line_in_order(x, s_, arr_list_previous)
            x = x + 26
        Next
    End If
End Sub

Private Sub lb_a_all_Click()
    lb_a_all.BackColor = cl_new
    lb_p_provider.BackColor = cl_old
    lb_c_customer.BackColor = cl_old
    lb_w_worker.BackColor = cl_old
    Dim exc_s As String
    exc_s = "Select [Name_c] from Customers"
    Call DB.fill_combobox_customer(cmb_customer, exc_s, 1)
End Sub

Private Sub lb_add_Click()
    lb_add.BackColor = cl_new
    lb_nhua.BackColor = cl_old
    lb_nhiet.BackColor = cl_old
    lb_dien.BackColor = cl_old
    lb_van.BackColor = cl_old
    lb_thietbi.BackColor = cl_old
    lb_kimkhi.BackColor = cl_old
    lb_nuocsach.BackColor = cl_old
    lb_kem.BackColor = cl_old
    Call set_arr_list_product(1)
End Sub

Private Sub lb_c_customer_Click()
    lb_a_all.BackColor = cl_old
    lb_p_provider.BackColor = cl_old
    lb_c_customer.BackColor = cl_new
    lb_w_worker.BackColor = cl_old
    Dim exc_s As String
    exc_s = "Select [Name_c] from Customers where [Type]='Kh" & ChrW(225) & "ch l" & ChrW(7867) & "'"
    Call DB.fill_combobox_customer(cmb_customer, exc_s, 1)
End Sub

Private Sub lb_dien_Click()
    lb_add.BackColor = cl_old
    lb_nhua.BackColor = cl_old
    lb_nhiet.BackColor = cl_old
    lb_dien.BackColor = cl_new
    lb_van.BackColor = cl_old
    lb_thietbi.BackColor = cl_old
    lb_kimkhi.BackColor = cl_old
    lb_nuocsach.BackColor = cl_old
    lb_kem.BackColor = cl_old
    Call set_arr_list_product(2)
End Sub

Private Sub lb_kem_Click()
    lb_add.BackColor = cl_old
    lb_nhua.BackColor = cl_old
    lb_nhiet.BackColor = cl_old
    lb_dien.BackColor = cl_old
    lb_van.BackColor = cl_old
    lb_thietbi.BackColor = cl_old
    lb_kimkhi.BackColor = cl_old
    lb_nuocsach.BackColor = cl_old
    lb_kem.BackColor = cl_new
    Call set_arr_list_product(3)
End Sub

Private Sub lb_kimkhi_Click()
    lb_add.BackColor = cl_old
    lb_nhua.BackColor = cl_old
    lb_nhiet.BackColor = cl_old
    lb_dien.BackColor = cl_old
    lb_van.BackColor = cl_old
    lb_thietbi.BackColor = cl_old
    lb_kimkhi.BackColor = cl_new
    lb_nuocsach.BackColor = cl_old
    lb_kem.BackColor = cl_old
    Call set_arr_list_product(4)
End Sub

Private Sub lb_nhiet_Click()
    lb_add.BackColor = cl_old
    lb_nhua.BackColor = cl_old
    lb_nhiet.BackColor = cl_new
    lb_dien.BackColor = cl_old
    lb_van.BackColor = cl_old
    lb_thietbi.BackColor = cl_old
    lb_kimkhi.BackColor = cl_old
    lb_nuocsach.BackColor = cl_old
    lb_kem.BackColor = cl_old
    Call set_arr_list_product(5)
End Sub

Private Sub lb_nhua_Click()
    lb_add.BackColor = cl_old
    lb_nhua.BackColor = cl_new
    lb_nhiet.BackColor = cl_old
    lb_dien.BackColor = cl_old
    lb_van.BackColor = cl_old
    lb_thietbi.BackColor = cl_old
    lb_kimkhi.BackColor = cl_old
    lb_nuocsach.BackColor = cl_old
    lb_kem.BackColor = cl_old
    Call set_arr_list_product(6)
End Sub

Private Sub lb_nuocsach_Click()
    lb_add.BackColor = cl_old
    lb_nhua.BackColor = cl_old
    lb_nhiet.BackColor = cl_old
    lb_dien.BackColor = cl_old
    lb_van.BackColor = cl_old
    lb_thietbi.BackColor = cl_old
    lb_kimkhi.BackColor = cl_old
    lb_nuocsach.BackColor = cl_new
    lb_kem.BackColor = cl_old
    Call set_arr_list_product(7)
End Sub

Private Sub lb_p_provider_Click()
    lb_a_all.BackColor = cl_old
    lb_p_provider.BackColor = cl_new
    lb_c_customer.BackColor = cl_old
    lb_w_worker.BackColor = cl_old
    Dim exc_s As String
    exc_s = "Select [Name_c] from Customers where [Type]='CC'"
    Call DB.fill_combobox_customer(cmb_customer, exc_s, 1)
End Sub

Private Sub lb_thietbi_Click()
    lb_add.BackColor = cl_old
    lb_nhua.BackColor = cl_old
    lb_nhiet.BackColor = cl_old
    lb_dien.BackColor = cl_old
    lb_van.BackColor = cl_old
    lb_thietbi.BackColor = cl_new
    lb_kimkhi.BackColor = cl_old
    lb_nuocsach.BackColor = cl_old
    lb_kem.BackColor = cl_old
    Call set_arr_list_product(8)
End Sub

Private Sub lb_van_Click()
    lb_add.BackColor = cl_old
    lb_nhua.BackColor = cl_old
    lb_nhiet.BackColor = cl_old
    lb_dien.BackColor = cl_old
    lb_van.BackColor = cl_new
    lb_thietbi.BackColor = cl_old
    lb_kimkhi.BackColor = cl_old
    lb_nuocsach.BackColor = cl_old
    lb_kem.BackColor = cl_old
    Call set_arr_list_product(9)
End Sub

Private Sub lb_w_worker_Click()
    lb_a_all.BackColor = cl_old
    lb_p_provider.BackColor = cl_old
    lb_c_customer.BackColor = cl_old
    lb_w_worker.BackColor = cl_new
    Dim exc_s As String
    exc_s = "Select [Name_c] from Customers where [Type]='Th" & ChrW(7907) & "'"
    Call DB.fill_combobox_customer(cmb_customer, exc_s, 1)
End Sub

Private Sub lst_customer_history_Change()
    If isSeeingPL Then Call cmd_profit_loss_Click

    If lst_customer_history.ListIndex = -1 Then
        Call clear_label_info_history
        cmd_show_old_order.Enabled = False
        Exit Sub
    End If
    Dim i As Integer
    i = lst_customer_history.ListIndex
    If lst_customer_history.ListCount > 0 And i < arr_list_history.Count Then
        lb_order_id_history.Caption = arr_list_history.Item(i)(6)
        lb_code_order_history.Caption = arr_list_history.Item(i)(0)
        lb_project_history.Caption = arr_list_history.Item(i)(1)
        lb_time_order_history.Caption = arr_list_history.Item(i)(2)
        lb_total_amount_order_history.Caption = arr_list_history.Item(i)(3)
        lb_pay_order_history.Caption = arr_list_history.Item(i)(4)
        lb_note_history.Caption = arr_list_history.Item(i)(5)
        lb_ex_im_order.Caption = arr_list_history.Item(i)(7)
        If StrComp(arr_list_history.Item(i)(7), "EXP", vbTextCompare) = 0 Then
            lb_ex_im_order.ForeColor = RGB(0, 128, 255)
        Else
            lb_ex_im_order.ForeColor = RGB(255, 0, 0)
        End If
        cmd_show_old_order.Enabled = True
    ElseIf i < 0 Then
        cmd_show_old_order.Enabled = False
    End If
    
    
    If isPre = True Then
        Call cmd_back_to_present_order_Click
        Call cmd_show_old_order_Click
    End If
    
End Sub


Private Sub lst_inventory_Change()
    Dim exc_s As String
    Dim tmp_v As Variant, tmp_v_ As Variant
    Dim id_ As Long
    Dim imp_ As Double
    Dim exp_ As Double
    Dim i As Integer
    i = lst_inventory.ListIndex
    If i < arr_list_inventory.Count And i > -1 Then
        exc_s = "Select [Code],[Name],[Unit], [ID] from Products Where [Code]='" & lst_inventory.List(i, 0) & "'"
        tmp_v = DB.read_data(exc_s)
        If IsEmpty(tmp_v) Then Exit Sub
        lb_name_inventory.Caption = tmp_v(1, 0) 'arr_list_inventory.Item(i)(1)
        lb_code_inventory.Caption = tmp_v(0, 0) 'arr_list_inventory.Item(i)(0)
        lb_unit_inventory.Caption = tmp_v(2, 0) 'arr_list_inventory.Item(i)(2)
        id_ = CLng(tmp_v(3, 0))
        
        ' ------- lay so luong nhap ------------
        exc_s = "select sum (Quantify) from IO_Products Where [Product_ID]=" & id_ & " and [Type]='IMP'"
        tmp_v_ = DB.read_data(exc_s)
        If IsNull(tmp_v_(0, 0)) Then
            imp_ = 0
        Else
            imp_ = CDbl(tmp_v_(0, 0))
        End If
        
        '-------- lay so luong ban -------------
        exc_s = "select sum (Quantify) from IO_Products Where [Product_ID]=" & id_ & " and [Type]='EXP'"
        tmp_v_ = DB.read_data(exc_s)
        If IsNull(tmp_v_(0, 0)) Then
            exp_ = 0
        Else
            exp_ = CDbl(tmp_v_(0, 0))
        End If
        '-------- hien thi so luong ton kho --------
        lb_quantify_inventory.Caption = imp_ - exp_
    End If
End Sub

Private Sub lst_inventory_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Integer
    i = lst_inventory.ListIndex
    If StrComp(lst_inventory.List(i, 0), "", vbTextCompare) > 0 Then
        usf_add_item_u.Show (vbModeless)
    End If
End Sub

Private Sub txt_pay_prepay_Change()
    Call update_widget
End Sub

Private Sub creatte_widget_frame_list()
    Dim frm As Control
    Set frm = usf_order.Controls.Add("Forms.Frame.1", "frm_list_", True)
    With frm
        .Height = 300
        .Width = 610
        .Left = 257
        .top = 57 + 4
        .BackColor = RGB(255, 255, 255)
        .ScrollBars = fmScrollBarsNone
    End With
    collection_frm.Add frm
    
End Sub

Private Sub set_new_order_id_code()
    Dim td As Date
    td = Date
    Dim d As Integer, m As Integer, y As Integer, i As Integer, j As Integer
    Dim tmp_v As Variant
    d = Day(td)
    m = Month(td)
    y = Year(td)
    Dim date_s As String
    If d < 10 And m < 10 Then date_s = CStr(y) & "0" & CStr(m) & "0" & CStr(d)
    If d < 10 And m >= 10 Then date_s = CStr(y) & "" & CStr(m) & "0" & CStr(d)
    If d >= 10 And m < 10 Then date_s = CStr(y) & "0" & CStr(m) & "" & CStr(d)
    If d >= 10 And m >= 10 Then date_s = CStr(y) & "" & CStr(m) & "" & CStr(d)
    
    Dim exc_s As String
    exc_s = "Select [Code] from Orders"
    tmp_v = DB.read_data(exc_s)
    Dim l As Integer
    l = UBound(tmp_v, 2)
    Dim max_j As Integer
    max_j = 0
    'MsgBox "size: " & tmp_v(0, 5)
    For i = 0 To l
        If InStr(1, tmp_v(0, i), date_s, vbTextCompare) > 0 Then
            j = 0
            Do While StrComp(tmp_v(0, i), date_s & "" & j, vbTextCompare) <> 0
                j = j + 1
            Loop
            max_j = WorksheetFunction.Max(max_j, j)
            'lb_order_code.Caption = date_s & "" & j + 1
        End If
    Next
    lb_order_code.Caption = date_s & "" & max_j + 1
    lb_order_id.Caption = CStr(l + 3)
End Sub

Private Sub txt_search_inventory_Change()
    If StrComp(txt_search_inventory.Value, "", vbTextCompare) = 0 Then
        txt_search_inventory.Value = " "
        Exit Sub
    End If
    Dim arr_s() As String
    Dim j As Long
    j = 0
    Dim result As Boolean
    arr_s = Split(txt_search_inventory.Value)
    Dim i As Integer, i_ As Integer
    Dim arr() As String
    ReDim arr(lst_inventory.ListCount, 3)
    lst_inventory.Clear
    lst_inventory.ColumnCount = 3
    lst_inventory.ColumnWidths = "54;108;45"
    For i = 0 To arr_list_inventory.Count - 1
        For i_ = 0 To UBound(arr_s) - 1
            If InStr(1, arr_list_inventory.Item(i)(1), arr_s(i_), vbTextCompare) > 0 Then
                result = True
            Else
                result = False
                Exit For
            End If
        Next
        
        If result Then
            arr(j, 0) = arr_list_inventory.Item(i)(0)
            arr(j, 1) = arr_list_inventory.Item(i)(1)
            arr(j, 2) = arr_list_inventory.Item(i)(2)
            j = j + 1
        End If
    Next
    lst_inventory.List = arr
End Sub

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Order")
    Dim exc_s As String
    Dim s() As String
    
    x = 6
    id = 1
    index_edit = 0
    i_more = 0
    vat = 0
    cl_old = RGB(236, 236, 236)
    cl_new = RGB(160, 160, 160)
    frm_total.BackColor = RGB(255, 255, 255)
    lb_flat_1.BackColor = cl_new
    cmd_save.Enabled = False
    mlt_page.Value = 0
    lb_PL_total.Visible = False
    
    txt_order_time.Value = CStr(Date)
    lb_add.BackColor = cl_new
    cmd_all_history.BackColor = cl_new
    
    isPre = False
    cmd_show_old_order.Enabled = False
    cmd_back_to_present_order.Enabled = False
    
    Set arr_list = New ArrayList
    
    Call creatte_widget_frame_list
    
    txt_pay_prepay.Value = Format(0, "#,##0")
    Call set_icon_cmd
    Call set_widget_default
    
    lb_refund.Caption = 0
    
    Call set_new_order_id_code
    
    '======================================= CUSTOMER & ORDER INFOMATION ===================
    Call set_combobox_customer_default
    Call set_combobox_order_default
    
    '======================================= LIST INVENTORY ==================================
    '------------- Set list inventory---------------
    Set arr_list_inventory = New ArrayList
    
    ' tao danh sach
    lst_inventory.Clear
    lst_inventory.ColumnCount = 3
    lst_inventory.ColumnWidths = "54;108;45"
    
    Call set_arr_list_product(1)
    'Call set_arr_list_inventory
    
    '======================================= LIST HISTORY =================================
    'l = ws.Range("R" & Application.Rows.Count).End(xlUp).Row
    Set arr_list_history = New ArrayList
    'Call Algorithm.tranfer_data_from_excel_to_array_list(arr_list_history, ws, 18, 2, 23, l)
    
    lst_customer_history.Clear
    lst_customer_history.ColumnCount = 3
    lst_customer_history.ColumnWidths = "60;96;54"
    
    Call set_arr_list_history(1)
    
End Sub


Private Sub set_arr_list_history(i As Integer)
    Dim tmp_v As Variant
    Dim exc_s As String
    Dim s() As String
    lst_customer_history.ListIndex = -1
    If isPre = True Then Call cmd_back_to_present_order_Click
    lst_customer_history.Clear
    exc_s = "select [ID] from Orders where [Customer_ID] = " & CInt(cmb_customer_id.Value)
    tmp_v = DB.read_data(exc_s)
    If IsEmpty(tmp_v) Then
        Exit Sub
    End If
    
    Select Case i:
        Case 1:
            exc_s = "Select [Code], [Project], [Time],[Total],[Pay], [Note],[ID],[Type] from Orders where [Customer_ID]=" & CInt(cmb_customer_id.Value) & " order by [ID] ASC"
        Case 2:
            exc_s = "Select [Code], [Project], [Time],[Total],[Pay], [Note],[ID],[Type] from Orders where [Customer_ID]=" & CInt(cmb_customer_id.Value) & " and [Pay] < [Total] order by [ID] ASC"
        Case 3:
            exc_s = "Select [Code], [Project], [Time],[Total],[Pay], [Note],[ID],[Type] from Orders where [Customer_ID]=" & CInt(cmb_customer_id.Value) & " and [Pay] >= [Total] order by [ID] ASC"
        Case Else
    End Select
    tmp_v = DB.read_data(exc_s)
    If IsEmpty(tmp_v) Then Exit Sub
    Call DB.readData(arr_list_history, exc_s)
    s = covert_array_list_to_string_array(arr_list_history)
    lst_customer_history.List = s
End Sub

Private Sub set_arr_list_inventory()
    Dim tmp_v As Variant, tmp_v_ As Variant, tmp_v__ As Variant
    Dim exc_s As String
    exc_s = "SELECT distinct IO_Products.[Product_ID],Products.[Code], Products.[Name] FROM Products INNER JOIN IO_Products ON Products.[ID] = IO_Products.[Product_ID];"
    tmp_v = DB.read_data(exc_s)
    
    Dim tmp_s(3) As String
    Dim i_ As Integer
    For i_ = 0 To UBound(tmp_v, 2)
        exc_s = " Select Sum(Quantify) from IO_Products where [Product_ID]=" & CInt(tmp_v(0, i_)) & " and [Type]='IMP'"
        tmp_v_ = DB.read_data(exc_s)
        exc_s = " Select Sum(Quantify) from IO_Products where [Product_ID]=" & CInt(tmp_v(0, i_)) & " and [Type]='EXP'"
        tmp_v__ = DB.read_data(exc_s)
        tmp_s(0) = tmp_v(1, i_)
        tmp_s(1) = tmp_v(2, i_)
        If IsNull(tmp_v_(0, 0)) Then
            tmp_v_(0, 0) = 0
        End If
        If IsNull(tmp_v__(0, 0)) Then
            tmp_v__(0, 0) = 0
        End If
        tmp_s(2) = CDbl(tmp_v_(0, 0)) - CDbl(tmp_v__(0, 0))
        arr_list_inventory.Add tmp_s
    Next
End Sub

Private Sub set_arr_list_product(i As Integer)
    Dim tmp_v As Variant, tmp_v_ As Variant
    Dim exc_s As String, exc_s_ As String
    Dim s() As String
    lst_inventory.ListIndex = -1
    lst_inventory.Clear
    cmb_type_inventory.Clear
    'Dim s() As String
    exc_s = "Select Distinct [Class] from Products order by [Class] ASC"
    tmp_v = DB.read_data(exc_s)
    Select Case i
        Case 1:
            exc_s = "Select [Code],[Name],[Unit] from Products"
            exc_s_ = "Select Distinct [Type] from Products"
        Case 2:
            exc_s = "Select [Code],[Name],[Unit] from Products Where [Class]='" & tmp_v(0, 0) & "'"
            exc_s_ = "Select Distinct [Type] from Products Where [Class]='" & tmp_v(0, 0) & "'"
        Case 3:
            exc_s = "Select [Code],[Name],[Unit] from Products Where [Class]='" & tmp_v(0, 1) & "'"
            exc_s_ = "Select Distinct [Type] from Products Where [Class]='" & tmp_v(0, 1) & "'"
        Case 4:
            exc_s = "Select [Code],[Name],[Unit] from Products Where [Class]='" & tmp_v(0, 2) & "'"
            exc_s_ = "Select Distinct [Type] from Products Where [Class]='" & tmp_v(0, 2) & "'"
        Case 5:
            exc_s = "Select [Code],[Name],[Unit] from Products Where [Class]='" & tmp_v(0, 3) & "'"
            exc_s_ = "Select Distinct [Type] from Products Where [Class]='" & tmp_v(0, 3) & "'"
        Case 6
            exc_s = "Select [Code],[Name],[Unit] from Products Where [Class]='" & tmp_v(0, 4) & "'"
            exc_s_ = "Select Distinct [Type] from Products Where [Class]='" & tmp_v(0, 4) & "'"
        Case 7
            exc_s = "Select [Code],[Name],[Unit] from Products Where [Class]='" & tmp_v(0, 5) & "'"
            exc_s_ = "Select Distinct [Type] from Products Where [Class]='" & tmp_v(0, 5) & "'"
        Case 8
            exc_s = "Select [Code],[Name],[Unit] from Products Where [Class]='" & tmp_v(0, 6) & "'"
            exc_s_ = "Select Distinct [Type] from Products Where [Class]='" & tmp_v(0, 6) & "'"
        Case 9
            exc_s = "Select [Code],[Name],[Unit] from Products Where [Class]='" & tmp_v(0, 7) & "'"
            exc_s_ = "Select Distinct [Type] from Products Where [Class]='" & tmp_v(0, 7) & "'"
        Case Else
    End Select
    '----------------- Set type inventory--------------------------------------
    tmp_v_ = DB.read_data(exc_s_)
    Call Algorithm.set_combobox(cmb_type_inventory, tmp_v_)
    Call DB.readData(arr_list_inventory, exc_s)
    's = covert_array_list_to_string_array(arr_list_inventory)
    s = covert_array_list_to_string_array(arr_list_inventory)
    lst_inventory.List = s
End Sub

Private Sub print_array_list_(arr As ArrayList, ws As Worksheet)
    Dim i As Integer
    Dim j As Integer
    Call clear_store(ws)
    For i = 0 To arr.Count - 1
        For j = 0 To 2
            ws.Cells(2 + i, 18 + j).Value = arr.Item(i)(j)
            'Debug.Print arr.Item(i)(j)
        Next
    Next
End Sub

Private Function create_arr(ws As Worksheet, l As Integer) As String()
    Dim s() As String
    ReDim s(l - 1, 3)
    Dim i As Integer
    Dim j As Integer
    For i = 0 To l - 1
        For j = 0 To 2
            s(i, j) = CStr(ws.Cells(2 + i, 14 + j).Value)
        Next
    Next
    create_arr = s
End Function

Private Function covert_array_list_to_string_array(arr As ArrayList) As String()
    Dim s() As String
    ReDim s(arr.Count, 3)
    Dim i As Integer, j As Integer
    For i = 0 To arr.Count - 1
        For j = 0 To 2
            s(i, j) = arr.Item(i)(j)
        Next
    Next
    covert_array_list_to_string_array = s
End Function
