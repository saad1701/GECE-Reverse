VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSplash 
   Caption         =   "GECE"
   ClientHeight    =   5120
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   6760
   OleObjectBlob   =   "frmSplash.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdOpen_Click()
On Error GoTo ErrHandler

frmComplete.Show vbModal
Call Unload(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub UserForm_Activate()
On Error GoTo ErrHandler

Me.Controls("lblGECEVersion").Caption = "GECE " & GECEXLSVERSION
ThisWorkbook.Worksheets("CoverSheet").Select
ThisWorkbook.Worksheets("CoverSheet").Activate

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

