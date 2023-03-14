Attribute VB_Name = "Animation"
Option Explicit
Dim old_name As String

Public Sub hightlight(frm As MSForms.UserForm, name)
    On Error Resume Next
    With frm(old_name)
        '.BackColor = RGB(255, 255, 255)
        .BorderStyle = 0
        '.Left = 6
    End With
    old_name = name
    With frm(name)
        '.BackColor = RGB(211, 211, 211)
        .BorderStyle = 1
        '.Left = 5
    End With
End Sub

Public Sub normal(frm As MSForms.UserForm)
    On Error Resume Next
    With frm(old_name)
       '.BackColor = RGB(211, 211, 211)
       .BorderStyle = 1
       '.Left = 5
    End With
End Sub
