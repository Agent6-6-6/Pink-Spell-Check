VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Spelling_form 
   Caption         =   "Highlight Spelling For Review"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   OleObjectBlob   =   "Spelling_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Spelling_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'LAST UPDATED WITH REVISION B OF STANDARD TEMPLATE
'____________________________________________________________________________________________________________
'B - updated form to automate and toggle types of checks
'C.4 - REMOVED Check_all_capitals_BeforeUpdate MODULE AS NOT USED
'____________________________________________________________________________________________________________
Option Explicit
Public check_all As Boolean


Private Sub Check_all_capitals_Click()
    If Check_all_capitals.Value = False Then caps_border.Visible = False
    If Check_all_capitals.Value = True Then caps_border.Visible = True

End Sub

Private Sub Check_filenames_Click()
    If Check_filenames.Value = False Then filename_border.Visible = False
    If Check_filenames.Value = True Then filename_border.Visible = True

End Sub

Private Sub Check_mixed_digits_Click()
    If Check_mixed_digits.Value = False Then mixed_digits_border.Visible = False
    If Check_mixed_digits.Value = True Then mixed_digits_border.Visible = True

End Sub

Private Sub Check_Selection_Click()

    check_all = False
    Me.Hide
End Sub

Private Sub Check_Used_range_Click()
    check_all = True

    Me.Hide

End Sub

Private Sub UserForm_Initialize()
'hide labels representing shading within toggle buttons
    caps_border.Visible = False
    mixed_digits_border.Visible = False
    filename_border.Visible = False
End Sub
