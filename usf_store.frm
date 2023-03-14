VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usf_store 
   Caption         =   "Store"
   ClientHeight    =   8580.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   16908
   OleObjectBlob   =   "usf_store.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usf_store"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public al_import As ArrayList
Public al_product As ArrayList

Private Sub cmb_brand_Change()
    If StrComp(cmb_brand.Value, "", vbTextCompare) > 0 And StrComp(cmb_size.Value, "", vbTextCompare) > 0 And StrComp(cmb_type.Value, "", vbTextCompare) > 0 And StrComp(cmb_class.Value, "", vbTextCompare) > 0 Then
        cmb_unit.Enabled = True
        Dim exc_s As String
        exc_s = "Select distinct [Unit] from Products where [Class]='" & cmb_class.Value & "' and [Type]='" & cmb_type.Value & "' and [Size]='" & cmb_size.Value & "' and [Brand]='" & cmb_brand.Value & "'"
        Call DB.fill_combobox(cmb_unit, exc_s, 4)
        cmb_unit.ListIndex = 0
    End If
End Sub

Private Sub cmb_class_Change()
    cmb_size.Clear
    cmb_size.Enabled = False
    cmb_brand.Clear
    cmb_brand.Enabled = False
    cmb_unit.Clear
    cmb_unit.Enabled = False
    If StrComp(cmb_class.Value, "", vbTextCompare) > 0 Then
        cmb_type.Enabled = True
        Dim exc_s As String
        exc_s = "Select distinct [Type] from Products where [Class]='" & cmb_class.Value & "'"
        Call DB.fill_combobox(cmb_type, exc_s, 1)
    End If
End Sub

Private Sub cmb_provider_Change()
    Dim exc_s As String
    Dim tmp_v As Variant
    
    exc_s = "Select [ID_c],[Name_c],[Number],[Address] from Customers where [Type]='CC' and [Name_c]='" & cmb_provider.Value & "'"
    tmp_v = DB.read_data(exc_s)
    
    lb_id_customer.Caption = tmp_v(0, 0)
    lb_name_customer.Caption = tmp_v(1, 0)
    lb_number_customer.Caption = tmp_v(2, 0)
    lb_address_customer.Caption = tmp_v(3, 0)
    
    If StrComp(cmb_provider.Value, "", vbTextCompare) > 0 Then
        frm_add_product.Enabled = True
    End If
End Sub

Private Sub cmb_size_Change()
    cmb_unit.Clear
    cmb_unit.Enabled = False
    If StrComp(cmb_size.Value, "", vbTextCompare) > 0 And StrComp(cmb_type.Value, "", vbTextCompare) > 0 And StrComp(cmb_class.Value, "", vbTextCompare) > 0 Then
        cmb_brand.Enabled = True
        Dim exc_s As String
        exc_s = "Select distinct [Brand] from Products where [Class]='" & cmb_class.Value & "' and [Type]='" & cmb_type.Value & "' and [Size]='" & cmb_size.Value & "'"
        Call DB.fill_combobox(cmb_brand, exc_s, 3)
    End If
End Sub

Private Sub cmb_type_Change()
    cmb_brand.Clear
    cmb_brand.Enabled = False
    cmb_unit.Clear
    cmb_unit.Enabled = False
    If StrComp(cmb_type.Value, "", vbTextCompare) > 0 And StrComp(cmb_class, "", vbTextCompare) > 0 Then
        cmb_size.Enabled = True
        Dim exc_s As String
        exc_s = "Select distinct [Size] from Products where [Class]='" & cmb_class.Value & "' and [Type]='" & cmb_type.Value & "'"
        Call DB.fill_combobox(cmb_size, exc_s, 2)
    End If
End Sub

Private Sub cmb_unit_Change()

    Dim exc_s As String
    Dim tmp_v As Variant
    
    If StrComp(cmb_brand.Value, "", vbTextCompare) > 0 And StrComp(cmb_unit.Value, "", vbTextCompare) > 0 And StrComp(cmb_size.Value, "", vbTextCompare) > 0 And StrComp(cmb_type.Value, "", vbTextCompare) > 0 And StrComp(cmb_class.Value, "", vbTextCompare) > 0 Then
        exc_s = "Select [ID],[Code],[Name] from Products where [Class]='" & cmb_class.Value & "' and [Type]='" & cmb_type.Value & "' and [Size]='" & cmb_size.Value & "' and [Brand]='" & cmb_brand.Value & "' and [Unit]='" & cmb_unit.Value & "'"
        tmp_v = DB.read_data(exc_s)
        
        lb_product_id.Caption = tmp_v(0, 0)
        lb_product_code.Caption = tmp_v(1, 0)
        lb_product_name.Caption = tmp_v(2, 0)
    End If
    
End Sub

Private Sub cmd_import_Click()
    Dim td As Date
    td = Now()
    
    Dim exc_s As String
    'exc_s = "Insert into IO_Products([Product_ID],[Count],[Price],[Quantify],[Type],[Date],[Order_ID],[Note]) values (" & CInt(lb_product_id.Caption) & ",1000," & CDbl(txt_price.Value) & "," & CDbl(txt_quantify.Value) & ",'" & cmb_ex_im.Value & "','" & CDate(td) & "',,'" & txt_note.Value & "')"
    'Call DB.insert_row(exc_s)
End Sub

Private Sub cmd_provider_info_Click()
    If frm_provider_info.Height = 48 Then
        frm_provider_info.Height = 120
        frm_provider_info.ZOrder (0)
        frm_add_product.ZOrder (1)
    Else
        frm_provider_info.Height = 48
    End If
End Sub



Private Sub frm_add_product_Click()

End Sub

Private Sub txt_price_Change()
    If StrComp(txt_quantify.Value, "", vbTextCompare) > 0 And IsNumeric(txt_quantify.Value) And IsNumeric(txt_price) And StrComp(cmb_unit.Value, "", vbTextCompare) > 0 Then
        cmd_import.Enabled = True
    Else
        cmd_import.Enabled = False
    End If
End Sub

Private Sub txt_quantify_Change()
    If StrComp(txt_quantify.Value, "", vbTextCompare) > 0 And IsNumeric(txt_quantify.Value) And IsNumeric(txt_price) And StrComp(cmb_unit.Value, "", vbTextCompare) > 0 Then
        cmd_import.Enabled = True
    Else
        cmd_import.Enabled = False
    End If
End Sub

Private Sub UserForm_Initialize()
    Set al_product = New ArrayList
    Set al_import = New ArrayList
    frm_provider_info.Height = 48
    frm_add_product.Enabled = False
    
    Call set_icon_cmd
    Call set_default_textbox
    Call set_default_combobox
    Call set_default_listbox
End Sub

Private Sub set_icon_cmd()
    cmd_new_provider.Picture = LoadPicture("" & ThisWorkbook.Path & "\Images\icon_add.jpg")
    cmd_add_new_product.Picture = LoadPicture("" & ThisWorkbook.Path & "\Images\icon_add.jpg")
End Sub

Private Sub set_default_textbox()
    txt_price.Value = 0
    txt_quantify.Value = 0
End Sub

Private Sub set_default_combobox()
    Dim exc_s As String
    Dim tmp_v As Variant
    
    exc_s = "Select [Name_c] from Customers where [Type]='CC'"
    Call DB.fill_combobox_customer(cmb_provider, exc_s, 1)
    
    exc_s = "Select distinct [Class] from Products"
    Call DB.fill_combobox(cmb_class, exc_s, 5)
    
    cmb_type.Enabled = False
    cmb_size.Enabled = False
    cmb_brand.Enabled = False
    cmb_unit.Enabled = False
    cmd_import.Enabled = False
    
    cmb_ex_im.AddItem "IMP"
    cmb_ex_im.AddItem "EXP"
    cmb_ex_im.ListIndex = 0
End Sub


Private Sub set_default_listbox()
    Dim exc_s As String
    Dim tmp_v As Variant
    Dim tmp_s(10) As String
    Dim tmp_s_() As String
    
    '*** List Import *****
    exc_s = "Select distinct IO_Products.[Product_ID], IO_Products.[ID],Products.[Code], Products.[Name],Products.[Brand],Products.[Unit],IO_Products.[Quantify],Products.[Price_o],IO_Products.[Date],IO_Products.[Note] FROM Products INNER JOIN IO_Products ON Products.[ID] = IO_Products.[Product_ID] where IO_Products.[Type]='IMP' Order by IO_Products.[ID] ASC"
    tmp_v = DB.read_data(exc_s)
    Dim i_ As Integer
    For i_ = 0 To UBound(tmp_v, 2)
        'exc_s = " Select Sum(Quantify) from IO_Products where [Product_ID]=" & CInt(tmp_v(0, i_))
        'tmp_v_ = DB.read_data(exc_s)
        tmp_s(0) = tmp_v(1, i_) ' id
        tmp_s(1) = tmp_v(2, i_) ' code
        tmp_s(2) = tmp_v(3, i_) ' name
        tmp_s(3) = tmp_v(4, i_) ' brand
        tmp_s(4) = tmp_v(5, i_) ' unit
        tmp_s(5) = tmp_v(6, i_) ' quantify
        tmp_s(6) = tmp_v(7, i_) ' price
        tmp_s(7) = CDbl(tmp_s(5)) * CDbl(tmp_s(6)) ' total
        tmp_s(8) = tmp_v(8, i_) ' date
        tmp_s(9) = tmp_v(9, i_) ' note
        'tmp_s(2) = tmp_v_(0, 0)
        al_import.Add tmp_s
    Next
    
    lst_import.Clear
    lst_import.ColumnCount = 10
    lst_import.ColumnWidths = "24;54;144;72;36;36;48;60;60;90"
    lst_import.List = Algorithm.covert_array_list_to_string_array(al_import, 10)
    
    '*** List Product ****
    exc_s = "Select [ID],[Code],[Class],[Name],[Brand],[Unit],[Note] from Products order by [ID] ASC"
    'tmp_v = DB.read_data(exc_s)
    Call DB.readData(al_product, exc_s)
    'tmp_s = Algorithm.tranfer_variant_to_string_array(tmp_v)
    tmp_s_ = Algorithm.covert_array_list_to_string_array(al_product, 7)
    lst_product.Clear
    lst_product.ColumnCount = 7
    lst_product.ColumnWidths = "24;54;48;144;72;42;126"
    lst_product.List = tmp_s_
    
    
End Sub
