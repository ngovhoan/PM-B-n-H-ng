VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usf_test 
   Caption         =   "UserForm1"
   ClientHeight    =   9036.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   18120
   OleObjectBlob   =   "usf_test.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usf_test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public id As Integer
Public x As Integer
Public col_cmd As New Collection
Public col_frm As New Collection
Dim col_img As New Collection
Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    id = 1
    x = 0
    Dim frm As MSForms.Frame
    Dim cmd As MSForms.CommandButton
    Dim img As MSForms.Image
    Set frm = Me.Controls.Add("Forms.Frame.1", "frm_", True)
    
    With frm
        .Width = 200
        .Height = 100
        .Left = 10
        .top = 25
        '.ScrollBars = fmScrollBarsVertical
        '.ScrollHeight = 200
        .ScrollBars = fmScrollBarsNone
        '.TabIndex = 5
    End With
    col_frm.Add frm
    
    Set frm = Me.Controls.Add("Forms.Frame.1", "frm_h", True)
    With frm
        .Width = 200
        .Height = 100
        .Left = 230
        .top = 200
        '.ScrollBars = fmScrollBarsVertical
        '.ScrollHeight = 200
        .ScrollBars = fmScrollBarsNone
        '.TabIndex = 4
    End With
    col_frm.Add frm
    'Me.Controls("frm_").Repaint
    
    Me.Controls("frm_").ZOrder (0)
    Me.Controls("frm_").ZOrder (1)
    
    Set cmd = Me.Controls.Add("Forms.CommandButton.1", "cmd_", True)
    With cmd
        .Width = 200
        .Height = 200
        .top = 6
        .Left = 250
        
        '.Picture = LoadPicture("W:\BH\test_img.jpg")
    End With
    'Debug.Print Me.Controls("cmd_").TabIndex

    Dim event_ As AddRow
    Set event_ = New AddRow
    Set event_.cmdEvent = cmd
    col_cmd.Add event_
    
    Dim ctr As MSForms.Control
    For Each ctr In col_frm
        Debug.Print ctr.name
    Next
    
    Debug.Print Me.Controls(1).name
End Sub




