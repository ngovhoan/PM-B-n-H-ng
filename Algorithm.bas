Attribute VB_Name = "Algorithm"
Option Explicit

' --------------- Lay du lieu vao mang chuoi ky tu -----------------
Public Sub create_arr(ws As Worksheet, col As Integer, start As Integer, finish As Integer, s() As String)
    ReDim s(finish - start + 1) As String
    Dim i As Integer
    For i = 0 To (finish - start)
        s(i) = CStr(ws.Cells(start + i, col).Value)
    Next
End Sub

Public Sub tranfer_data_to_array_list(usf As MSForms.UserForm, i As Integer)
    Dim s(11) As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Order")
    Dim n As Integer
    n = ws.Range("B" & Application.Rows.Count).End(xlUp).Row
    
    '------------- thu tu-------------
    s(0) = CStr(i)
    '------------- ma----------------
    s(1) = usf.lb_code.Caption
    '------------- Chung loai---------
    s(2) = usf.cmb_type.Value
    '------------- Kich thuoc --------
    s(3) = usf.cmb_size.Value
    '------------- Ten --------
    s(4) = usf.lb_name.Caption
    '------------- Thuong hieu -------
    s(5) = usf.cmb_brand.Value
    '------------- Don vi ------------
    s(6) = usf.cmb_unit.Value
    '------------- So luong ---------
    s(7) = usf.txt_quantify.Value
    '------------- Don gia ----------
    s(8) = usf.txt_price.Value
    '------------- Thanh tien ---------
    s(9) = usf_order.Controls("lb_total_item_" & i).Caption
    '------------- Loai gia -------
    s(10) = usf.cmb_type_price.Value
    
    usf_order.arr_list.Add s
    
    '----- to Excel order -----------------
    Call print_array_list(usf_order.arr_list, ws)
    
End Sub

Public Sub tranfer_data_to_array_list_(usf As MSForms.UserForm, a_l As ArrayList, s() As String, i As Integer)
    
    '------------- thu tu-------------
    s(0) = CStr(i)
    
    a_l.Add s
    
    'Call print_string_arr(s)
    
    '----- to Excel order -----------------
    'Call print_array_list(a_l, ws)
    
End Sub


Public Sub tranfer_data_an_item_to_string_array(s() As String, usf As MSForms.UserForm)
    'ReDim s(11) As String
    ' ------------ tt ----------
    '------------- ma----------------
    If usf.lb_code.Caption = Null Then
        s(1) = ""
    Else
        s(1) = usf.lb_code.Caption
    End If
    '------------- Chung loai---------
    If usf.cmb_type.Value = Null Then
        s(2) = ""
    Else
        s(2) = usf.cmb_type.Value
    End If
    '------------- Kich thuoc --------
    If usf.cmb_size.Value = Null Then
        s(3) = ""
    Else
        s(3) = usf.cmb_size.Value
    End If
    '------------- Ten --------
    If usf.lb_name.Caption = Null Then
        s(4) = ""
    Else
        s(4) = usf.lb_name.Caption
    End If
    '------------- Thuong hieu -------
    If usf.cmb_brand.Value = Null Then
        s(5) = ""
    Else
        s(5) = usf.cmb_brand.Value
    End If
    '------------- Don vi ------------
    If usf.cmb_unit.Value = Null Then
        s(6) = ""
    Else
        s(6) = usf.cmb_unit.Value
    End If
    '------------- So luong ---------
    If usf.txt_quantify.Value = Null Then
        s(7) = 0
    Else
        s(7) = usf.txt_quantify.Value
    End If
    '------------- Don gia ----------
    If usf.txt_price.Value = Null Then
        s(8) = 0
    Else
        s(8) = usf.txt_price.Value
    End If
    
    s(9) = CStr(CDbl(s(7)) * CDbl(s(8)))
    '------------- Thanh tien ------
    If usf.cmb_type_price.Value = Null Then
        s(10) = ""
    Else
        s(10) = usf.cmb_type_price.Value
    End If
    '------------- Lop--------------
    If usf.cmb_class.Value = Null Then
        s(11) = ""
    Else
        s(11) = usf.cmb_class.Value
    End If
    
    s(12) = usf.lb_id.Caption
    
    s(13) = usf.lb_profit_loss_price.Caption
    
    s(14) = CStr(s(7) * s(13))
End Sub

Public Sub update_item_to_order(arr As ArrayList, n As Integer, usf As MSForms.UserForm)
    Dim i As Integer
    For i = 0 To arr.Count - 1
        If arr.Item(i)(0) = n Then
            usf.Controls("lb_count_item_" & n).Caption = arr.Item(i)(0)
            usf.Controls("lb_code_item_" & n).Caption = arr.Item(i)(1)
            usf.Controls("lb_name_item_" & n).Caption = arr.Item(i)(4)
            usf.Controls("lb_brand_item_" & n).Caption = arr.Item(i)(5)
            usf.Controls("lb_unit_item_" & n).Caption = arr.Item(i)(6)
            usf.Controls("lb_quantify_item_" & n).Caption = Format(arr.Item(i)(7), "#.0")
            usf.Controls("lb_price_item_" & n).Caption = Format(arr.Item(i)(8), "#,##0")
            usf.Controls("lb_total_item_" & n).Caption = Format(arr.Item(i)(9), "#,##0")
            usf.Controls("lb_price_profit_loss_item_" & n).Caption = Format(arr.Item(i)(13), "#,##0")
            usf.Controls("lb_total_profit_loss_item_" & n).Caption = Format(arr.Item(i)(14), "#,##0")
            Exit For
        End If
    Next
End Sub

'----------- Hien thi order trong excel-------------------
Public Sub print_array_list(arr As ArrayList, ws As Worksheet)
    Dim i As Integer
    Dim j As Integer
    Call clear_store(ws)
    For i = 0 To arr.Count - 1
        For j = 0 To 11
            ws.Cells(2 + i, 2 + j).Value = arr.Item(i)(j)
            'Debug.Print arr.Item(i)(j)
        Next
    Next
End Sub

Public Sub print_string_arr(s() As String)
    Dim i As Integer
    For i = 0 To UBound(s) - 1
        Debug.Print " S " & s(i)
    Next
End Sub

Public Sub tranfer_data_from_excel_to_array_list(arr As ArrayList, ws As Worksheet, x As Integer, y As Integer, lx As Integer, ly As Integer)
    Dim i As Integer, j As Integer
    'Dim s() As String
    'ReDim s(lx - x + 1)
    'Set arr = New ArrayList
    For i = 0 To ly - y
        arr.Add excel_to_string_array(ws, x, y + i, lx)
    Next
    
End Sub

Public Function excel_to_string_array(ws As Worksheet, x As Integer, y As Integer, lx As Integer) As String()
    Dim j As Integer
    Dim s() As String
    ReDim s(lx - x + 1)
    For j = 0 To lx - x
        s(j) = CStr(ws.Cells(y, x + j).Value)
    Next
    excel_to_string_array = s
End Function


Public Sub edit_an_item(arr As ArrayList, n As Integer, s() As String)
    Dim i As Integer, j As Integer
    For i = 0 To arr.Count - 1
        If CInt(arr.Item(i)(0)) = n Then
            arr.RemoveAt i
            arr.Insert i, s
            Exit For
        End If
    Next
End Sub

Public Sub insert_an_item(arr As ArrayList, n As Integer, s() As String)
    Dim i As Integer, j As Integer
    For i = 0 To arr.Count - 1
        If CInt(arr.Item(i)(0)) = n Then
            arr.Insert i + 1, s
            Exit For
        End If
    Next
End Sub

Public Sub coppy_paste_an_item(arr As ArrayList, n As Integer)
    Dim i As Integer, j As Integer
    For i = 0 To arr.Count - 1
        If CInt(arr.Item(i)(0)) = n Then
            arr.Insert i + 1, arr.Item(i)
            Exit For
        End If
    Next
End Sub


Public Sub delete_an_item(arr As ArrayList, n As Integer)
    Dim i As Integer
    For i = 0 To arr.Count - 1
        If CInt(arr.Item(i)(0)) = n Then
            arr.RemoveAt (i)
            Exit For
        End If
    Next
End Sub

Public Function acess_an_item_to_string_arr(arr As ArrayList, n As Integer, l As Integer) As String()
    Dim i As Integer, j As Integer
    Dim s() As String
    ReDim s(l)
    For i = 0 To arr.Count - 1
        If CInt(arr.Item(i)(0)) = n Then
            For j = 0 To l - 1
                s(j) = arr.Item(i)(j)
            Next
            Exit For
            
        End If
    Next
    acess_an_item_to_string_arr = s
End Function

Public Function tranfer_variant_to_string_array(tmp_v As Variant) As String()
    Dim str() As String
    ReDim str(UBound(tmp_v, 2) + 1, UBound(tmp_v) + 1)
    'Debug.Print UBound(tmp_v, 2) + 1
    'Debug.Print UBound(tmp_v) + 1
    Dim i As Long, j As Long
    For i = 0 To UBound(tmp_v, 2)
        'Debug.Print i
        For j = 0 To UBound(tmp_v)
            str(i, j) = tmp_v(j, i)
            'Debug.Print str(i, j)
        Next
    Next
    tranfer_variant_to_string_array = str
End Function

Public Function coppy_array_list(arr As ArrayList) As ArrayList
    Dim a_l As ArrayList
    Set a_l = New ArrayList
    Dim i As Integer
    For i = 0 To arr.Count - 1
        a_l.Add arr.Item(i)
    Next
    Set coppy_array_list = a_l
End Function


Public Function covert_array_list_to_string_array(arr As ArrayList, col As Integer) As String()
    Dim s() As String
    ReDim s(arr.Count, col)
    Dim i As Integer, j As Integer
    For i = 0 To arr.Count - 1
        For j = 0 To col - 1
            s(i, j) = arr.Item(i)(j)
        Next
    Next
    covert_array_list_to_string_array = s
End Function

Public Sub clear_store(ws As Worksheet)
    Dim i As Integer
    i = ws.Range("B" & Application.Rows.Count).End(xlUp).Row
    If i > 1 Then
        ws.Range("B2:L" & i).ClearContents
    End If
End Sub


'---------------Set combobox by msacess-----------
Public Sub set_combobox(ctr As MSForms.ComboBox, v As Variant)
    Dim i As Integer
    ctr.Clear
    For i = 0 To UBound(v, 2)
        'Debug.Print v(i, 0)
        ctr.AddItem v(0, i)
    Next
End Sub

'------------ excel order: insert row ---------
Public Sub insert_new_row_in_excel(ws As Worksheet, a_l As ArrayList, i_row As Integer, i_arr As Integer)
    Dim i As Integer
    ws.Rows(i_row).EntireRow.Insert
    ws.Cells(i_row, 1).Value = a_l.Item(i_arr)(0)
    ws.Cells(i_row, 2).Value = a_l.Item(i_arr)(4)
    ws.Cells(i_row, 3).Value = a_l.Item(i_arr)(5)
    ws.Cells(i_row, 4).Value = a_l.Item(i_arr)(6)
    ws.Cells(i_row, 5).Value = a_l.Item(i_arr)(7)
    ws.Cells(i_row, 6).Value = a_l.Item(i_arr)(8)
    ws.Cells(i_row, 7).Value = a_l.Item(i_arr)(9)
End Sub

'---------- excel order: clean row------------
Public Sub clean_row_in_excel(ws As Worksheet)
    Dim l As Integer
    l = ws.Range("B" & Application.Rows.Count).End(xlUp).Row
    Do While l > 13
        ws.Rows(6).EntireRow.Delete
        l = l - 1
    Loop
End Sub
