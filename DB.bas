Attribute VB_Name = "DB"

Option Explicit

Public Sub readData(arr As ArrayList, c_n As String)
    Dim cn As Object
    Dim rs As Object
    
    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    With cn
        .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\du_lieu.accdb;Persist Security Info=False;"
        .Open
    End With
    
    Set rs = cn.Execute(c_n)
    
    Dim s As Variant
    s = rs.GetRows()
    Set arr = convert_arr_to_arr_list(s)
    'Call printArrList(arr)
    
    'Call printArr(s)
    'Call printArrList(convert_arr_to_arr_list(s))
End Sub

Public Function read_data(exc_s As String) As Variant
    Dim cn As Object
    Dim rs As Object
    
    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    With cn
        .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\du_lieu.accdb;Persist Security Info=False;"
        .Open
    End With
    
    Set rs = cn.Execute(exc_s)

    On Error GoTo ex
    Dim s As Variant
    s = rs.GetRows()
    'cn.Close
ex:         Debug.Print Error
    read_data = s
End Function

Public Function read_data_(exc_s As String) As Variant
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim con_str As String
    con_str = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\du_lieu.accdb;Persist Security Info=False;"
    rs.Open exc_s, con_str
    Dim s As Variant
    s = rs.GetRows()
    read_data_ = s
End Function

Public Sub fill_combobox(ctrl As MSForms.ComboBox, exc_s As String, i As Integer)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    ctrl.Clear
    Dim con_str As String
    con_str = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\du_lieu.accdb;Persist Security Info=False;"
    rs.Open exc_s, con_str
    Do While Not rs.EOF
        Select Case i:
            Case 1:
                ctrl.AddItem rs!Type
            Case 2:
                ctrl.AddItem rs!Size
            Case 3:
                ctrl.AddItem rs!brand
            Case 4:
                ctrl.AddItem rs!Unit
            Case 5:
                ctrl.AddItem rs!Class
        End Select
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub

Public Sub fill_combobox_customer(ctrl As MSForms.ComboBox, exc_s As String, i As Integer)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    ctrl.Clear
    Dim con_str As String
    con_str = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\du_lieu.accdb;Persist Security Info=False;"
    rs.Open exc_s, con_str
    Do While Not rs.EOF
        Select Case i:
            Case 1:
                ctrl.AddItem rs!Name_c
            Case 2:
                ctrl.AddItem rs!Code
            Case 3:
                ctrl.AddItem rs!Type
            Case 4:
                ctrl.AddItem rs!id_c
            Case 5:
                ctrl.AddItem rs!Project
        End Select
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub



Public Sub fill_text(txt As MSForms.TextBox, exc_s As String, i As Integer)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim con_str As String
    con_str = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\du_lieu.accdb;Persist Security Info=False;"
    rs.Open exc_s, con_str
    Do While Not rs.EOF
        Select Case i:
            Case 1:
                txt.Value = rs!Price_1
                'Debug.Print TypeName(rs!Price_1)
            Case 2:
                txt.Value = rs!Price_2
                'TypeName (rs!Price_2)
            Case 3:
                txt.Value = rs!Price_o
        End Select
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub

Public Sub insert_row(exc_s As String)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim con_str As String
    con_str = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\du_lieu.accdb;Persist Security Info=False;"
    rs.Open exc_s, con_str
    Set rs = Nothing
End Sub

Private Sub printArr(arr As Variant)
    Dim i As Integer
    Dim j As Integer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    For i = 0 To UBound(arr, 2)
        For j = 0 To UBound(arr)
            'Debug.Print arr(j, i)
            ws.Cells(1 + i, 1 + j).Value = arr(j, i)
        Next
    Next
End Sub

Private Sub printArrList(arr As ArrayList)
    Dim i As Integer, j As Integer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    For i = 0 To arr.Count - 1
        For j = 0 To UBound(arr.Item(0))
            'Debug.Print arr(j, i)
            ws.Cells(1 + i, 1 + j).Value = arr.Item(i)(j)
        Next
    Next
End Sub

Private Function convert_arr_to_arr_list(arr As Variant) As ArrayList
    Dim i As Integer, j As Integer
    Dim arr_lst As ArrayList
    Set arr_lst = New ArrayList
    Dim s() As String
    ReDim s(UBound(arr) + 1)
    For i = 0 To UBound(arr, 2)
        For j = 0 To UBound(arr)
            On Error GoTo ex:
            s(j) = arr(j, i)
ex:             Debug.Print Error
        Next
        arr_lst.Add s
    Next
    Set convert_arr_to_arr_list = arr_lst
End Function
