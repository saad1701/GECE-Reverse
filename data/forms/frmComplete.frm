VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmComplete 
   Caption         =   "Detailed Data Entry Form"
   ClientHeight    =   8230.001
   ClientLeft      =   50
   ClientTop       =   440
   ClientWidth     =   10770
   OleObjectBlob   =   "frmComplete.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmComplete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub btnExport_Click()
    Dim fName As Variant
    Dim defName As String

    ' propose same folder/name as workbook, but .xml
    defName = Replace(Replace(Replace(ThisWorkbook.FullName, ".xlsm", ".xml"), ".xlsb", ".xml"), ".xls", ".xml")

    fName = Application.GetSaveAsFilename( _
                InitialFileName:=defName, _
                FileFilter:="XML Files (*.xml), *.xml", _
                Title:="Save Exported XML File As")

    ' user pressed Cancel
    If VarType(fName) = vbBoolean And fName = False Then Exit Sub

    CreateMSXMLFile ThisWorkbook.Name, CStr(fName)

    MsgBox "XML file exported successfully." & vbCrLf & _
           "Saved as: " & CStr(fName), vbInformation, "Export Complete"
End Sub


Private Sub cboCustomerType_Change()
Call UpdateSummaryPage(Me)
End Sub


Public Sub YesNo(Answer As Integer)
    Answer = MsgBox("Do you want to reset all remote countries? This will also overwrite all remote countries on the Application Based sheet.", vbExclamation + vbYesNo, "Reset Remote Country")
End Sub

Private Sub cboDefaultRemoteCountry_Change()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(gstrGECEAssumptionsProposalSheet)

    'Update the default remote country from combo selection
    ws.Range("DEFAULT_REM_COUNTRY").Value = Me.cboDefaultRemoteCountry.Value

    'Keep the unique actions (not duplicated in Worksheet_Change)
    UpdateSummaryPage Me
    ThisWorkbook.Worksheets("Proposal Summary").Calculate
    MsgBox "Remote Country values have been updated successfully.", _
           vbInformation, "Update Complete"
End Sub

Private Sub cboIndustry_Change()
Call UpdateSummaryPage(Me)
End Sub

Private Sub cboLocalCuntry_AfterUpdate()
Call SetWPA(Me)
Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("WPA") = ""
Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("WPA_TYPE") = ""
Call UpdateSummaryPage(Me)
End Sub

Private Sub cboToolkit_Change()
Call UpdateSummaryPage(Me)
End Sub

Private Sub cbTypeCompressor_AfterUpdate()
    Call UpdateTMCTab
    Call UpdateSummaryPage(Me)
End Sub



Private Sub ckbAPP_BUS_STS_Click()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub ckbDOC_BOM_REQ_Click()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub ckbDOC_CAB_ELEC_REQ_Click()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub ckbDOC_CAB_MECH_REQ_Click()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub ckbDOC_LOOP_REQ_Click()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub ckbDOC_PWR_GND_REQ_Click()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub ckbDOC_PWR_HEAT_REQ_Click()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub ckbDOC_QA_REQ_Click()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub ckbDOC_SYS_ARCH_REQ_Click()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub ckbDOC_SYS_IND_REQ_Click()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub ckbDOC_TAGLIST_REQ_Click()
    Call UpdateSummaryPage(Me)
End Sub



Private Sub cmdAdvanced_Click()
On Error Resume Next
    ThisWorkbook.Worksheets(gstrGECEApplicationBasedSheet).Select
    ThisWorkbook.Worksheets(gstrGECEApplicationBasedSheet).Activate
    Me.Hide
    
End Sub
Private Sub cmdAdvancedUserGuide_Click()
    On Error GoTo ErrHandler
    Dim w As Object, s As String
    Set w = CreateObject("WScript.Shell")
    s = AddSlash(GetGECEPath_Universal()) & "GECEAdvancedUserGuide.doc"
    w.Run """" & s & """"
CleanUp: Set w = Nothing: Exit Sub
ErrHandler: MsgBox Err.Number & "; " & Err.Description: Resume CleanUp
End Sub

Private Sub cmdCompass_Click()
On Error GoTo ErrHandler
    Dim mstrIndustryType
    Application.Cursor = xlWait
'
'    Call MergeAndPrintRTF
    Dim dblTimer As Double
    Dim myCompass As Object
    Dim strCSE As String
'    Dim strDescription As String
'    Dim strIndustryType As String
'    Dim strOpportunityID As String
'    Dim strProjectMGR As String
    
    
    Set myCompass = CreateObject("CompassLookup.CompassLookup")
    Application.Cursor = xlDefault
    Call myCompass.LoadCompassUI
    
    If myCompass.cancellookup Then
        GoTo CleanUp
    End If
    
    dblTimer = Timer
    Do Until (Timer - dblTimer) > 1
        DoEvents
        strCSE = myCompass.CSE
        txtPROJECT_NAME.Text = myCompass.Description
        'The industries you have do not match up
        mstrIndustryType = myCompass.IndustryType
        'The industries you have do not match up
        
        txtGSP_ID.Text = myCompass.OpportunityID
        txtPROJECT_MANAGER.Text = myCompass.CSE
        txtCUSTOMER_NAME.Text = myCompass.Account
    Loop
    
    strCSE = myCompass.CSE
    txtPROJECT_NAME.Text = myCompass.Description
    'The industries you have do not match up
    mstrIndustryType = myCompass.IndustryType
    'The industries you have do not match up
    
    txtGSP_ID.Text = myCompass.OpportunityID
    txtPROJECT_MANAGER.Text = myCompass.CSE
    txtCUSTOMER_NAME.Text = myCompass.Account
    
    
CleanUp:
On Error Resume Next
    Set myCompass = Nothing
    DoEvents
    txtGSP_ID.Enabled = True
    txtPROJECT_MANAGER.Enabled = True
    txtCUSTOMER_NAME.Enabled = True
    'MsgBox "Compass data accepted!"

    Application.Cursor = xlDefault
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
    
End Sub

Private Sub cmdExportToERP_Click()
    ExportToERP
    Application.Cursor = xlDefault
End Sub


'Private Sub cmdImportFast_Click()
'On Error Resume Next
'Dim strFileToOpen As String
'strFileToOpen = Application.GetOpenFilename("Excel Files (*.xls), *.xls, All Files (*.*),*.*", , "Select the workbook to import from", , False) ', , "Select the workbook to import from", "Import", False
'If strFileToOpen <> "False" Then
'    ImportWorkbookFast (strFileToOpen)
'
'    'reset form
'    Call UserForm_Activate
'
'End If
'End Sub

Private Sub cmdImportFull_Click()
    On Error Resume Next


Dim strFileToOpen As String
Dim strMSG As String
strMSG = "You can import an Excel file (.xls or .xlsx) or XML file" & Chr(13) & Chr(13)
strMSG = strMSG & "Select 'Yes' for Excel, 'No' for XML, or 'Cancel' to exit."

Select Case MsgBox(strMSG, vbYesNoCancel + vbQuestion)
    Case vbYes
       strFileToOpen = Application.GetOpenFilename("Excel Files (*.xls), *.xls, All Files (*.*),*.*", , "Select the workbook to import from", , False) ', , "Select the workbook to import from", "Import", False
        If strFileToOpen <> "False" Then
            ImportWorkbookFull (strFileToOpen)
            'Reset Form
            Call UserForm_Activate
            'reset the tab to the about tab
            Me.Controls("MultiPage1").Value = 16
        End If
    Case vbNo
        strFileToOpen = Application.GetOpenFilename("XML Files (*.xml), *.xml, All Files (*.*),*.*", , "Select the xml file to import", , False)
        If strFileToOpen <> "False" Then
            Call ReadGECEXML(strFileToOpen, Application.Workbooks(1).Name)
            'reset form
             'Call UserForm_Activate
            'reset the tab to the about tab
            Me.Controls("MultiPage1").Value = 16
        End If
    Case vbCancel

End Select

End Sub

Private Sub cmdMergeFile_Click()
    MergeFile
    Application.Cursor = xlDefault
End Sub

Private Sub cmdPrintProposal_Click()
On Error Resume Next
    Application.Cursor = xlWait
    
    Call MergeAndPrintRTF
    Application.Cursor = xlDefault
End Sub

Private Sub cmdReleaseNotes_Click()
    On Error GoTo ErrHandler
    Dim w As Object, s As String
    Set w = CreateObject("WScript.Shell")
    s = AddSlash(GetGECEPath_Universal()) & "GECEReleaseNotes.doc"
    w.Run """" & s & """"
CleanUp: Set w = Nothing: Exit Sub
ErrHandler: MsgBox Err.Number & "; " & Err.Description: Resume CleanUp
End Sub

Private Sub cmdSchedule_Click()
On Error GoTo ErrHandler

    ThisWorkbook.Worksheets("coverSheet").Select
    ThisWorkbook.Worksheets("coverSheet").Activate
        If Not Gantt_by_Phases() Then
        MsgBox "Gantt export failed.", vbExclamation
    Else
        MsgBox "Gantt created.", vbInformation
    End If

    Exit Sub
ErrHandler:
    MsgBox "Err " & Err.Number & ": " & Err.Description, vbCritical, "GECE Gantt Export"
    Resume Next
End Sub

Private Sub cmdScopeChange_Click()
    ScopeChange
End Sub

Private Sub cmdUserGuide_Click()
    On Error GoTo ErrHandler
    Application.Cursor = xlWait

    'Optional: if the form isn't open from elsewhere, ensure it is visible
    Me.txtOutput.Text = "Starting reset..." & vbCrLf
    Me.Repaint

    ResetWorkbookFields

Clean:
    Application.Cursor = xlDefault
    Exit Sub
ErrHandler:
    Application.Cursor = xlDefault
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical, "GECE"
    Resume Clean
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrHandler

    'may not save the first time so try ThisWorkbook.Save
    'Application.ActiveWorkbook.Save
    'Call CreateMSXMLFile(ThisWorkbook.Application.Workbooks(1).Name, ThisWorkbook.Application.Workbooks(1).FullName)

    ThisWorkbook.Save

Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume Next
End Sub

'################################################
'set the option buttons on the documentation tab
Private Sub setCUSTOMER_SPEC_REQ()
On Error GoTo ErrHandler

    'no is set as default so it has a value of true
    Me.Controls("obCUSTOMER_SPEC_REQ_no").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("CUSTOMER_SPEC_REQ").Value
    'set to the opposite of soft i/o
    Me.Controls("obCUSTOMER_SPEC_REQ_yes").Value = Not (Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("CUSTOMER_SPEC_REQ").Value)
    
CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub cmdXML_Click()
    'Call CreateMSXMLFile(ThisWorkbook.Application.Workbooks(1).Name, ThisWorkbook.Application.Workbooks(1).FullName)
End Sub

Private Sub obCUSTOMER_SPEC_REQ_no_Click()
On Error GoTo ErrHandler
    
    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("CUSTOMER_SPEC_REQ").Value = False
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub obCUSTOMER_SPEC_REQ_yes_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("CUSTOMER_SPEC_REQ").Value = True
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

'################################################
'hardwired and soft i/o settings
Private Sub setDI_IOTYPE_STS()
On Error GoTo ErrHandler

    Me.Controls("obDI_IOTYPE_STS_softio").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("DI_IOTYPE_STS").Value
    'set to the opposite of soft i/o
    Me.Controls("obDI_IOTYPE_STS_Hardwired").Value = Not (Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("DI_IOTYPE_STS").Value)
    
CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub obDI_IOTYPE_STS_Hardwired_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("DI_IOTYPE_STS").Value = False
    Call UpdateDeviceIntegrationTab
    Call UpdateSummaryPage(Me)
    
CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub obDI_IOTYPE_STS_softio_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("DI_IOTYPE_STS").Value = True
    Call UpdateDeviceIntegrationTab
    Call UpdateSummaryPage(Me)
    
CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
'#####################################################
'Testing
Public Sub SetSW_SIMULATOR_REQ()
On Error GoTo ErrHandler

    Me.Controls("obSW_SIMULATOR_REQ_yes").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("SW_SIMULATOR_REQ").Value
    'set to the opposite of soft i/o
    Me.Controls("obSW_SIMULATOR_REQ_no").Value = Not (Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("SW_SIMULATOR_REQ").Value)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub obINCLUDE_ESCALATION_no_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("INCLUDE_ESCALATION").Value = False
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub




Private Sub obESD_MARSH_CAB_REQ2_no_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("ESD_MARSH_CAB_REQ").Value = False


CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub obESD_MARSH_CAB_REQ2_yes_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("ESD_MARSH_CAB_REQ").Value = True


CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub obESD_MARSH_CAB_REQ_no_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("ESD_MARSH_CAB_REQ").Value = False


CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub obESD_MARSH_CAB_REQ_yes_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("ESD_MARSH_CAB_REQ").Value = True


CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub obHIGH_RISK_SITE_no_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("HIGH_RISK_SITE").Value = False
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

'################################################
'hardwired and soft i/o settings
Private Sub setHIGH_RISK_SITE()
On Error GoTo ErrHandler

    Me.Controls("obHIGH_RISK_SITE_yes").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("HIGH_RISK_SITE").Value
    'set to the opposite of soft i/o
    Me.Controls("obHIGH_RISK_SITE_no").Value = Not (Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("HIGH_RISK_SITE").Value)
    
CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub obHIGH_RISK_SITE_yes_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("HIGH_RISK_SITE").Value = True
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub obMARSH_CAB_REQ_no_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("MARSH_CAB_REQ").Value = False


CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub obMARSH_CAB_REQ_yes_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("MARSH_CAB_REQ").Value = True


CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub obSW_SIMULATOR_REQ_no_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("SW_SIMULATOR_REQ").Value = False
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub obSW_SIMULATOR_REQ_yes_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("SW_SIMULATOR_REQ").Value = True
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
'#####################################################
' Added by AB 12/09/2006
'ESD
Public Sub SetESD_DEFAULT()
On Error GoTo ErrHandler

    Call SetESD_HMI_REQ
    Call SetESD_PROG_REQ
    Call SetESD_SYSTEM_REQ
    Call SetESD_MARSH_CAB_REQ
    
CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Public Sub SetESD_HMI_REQ()
On Error GoTo ErrHandler

    Me.Controls("obESD_HMI_REQ_yes").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("ESD_HMI_REQ").Value
    'set to the opposite of soft i/o
    Me.Controls("obESD_HMI_REQ_no").Value = Not (Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("ESD_HMI_REQ").Value)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub obESD_HMI_REQ_yes_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("ESD_HMI_REQ").Value = True
    Call UpdateSummaryPage(Me)
    Call UpdateESDTab

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub obESD_HMI_REQ_no_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("ESD_HMI_REQ").Value = False
    Call UpdateSummaryPage(Me)
    Call UpdateESDTab

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Public Sub SetESD_PROG_REQ()
On Error GoTo ErrHandler

    Me.Controls("obESD_PROG_REQ_yes").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("ESD_PROG_REQ").Value
    'set to the opposite of soft i/o
    Me.Controls("obESD_PROG_REQ_no").Value = Not (Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("ESD_PROG_REQ").Value)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub obESD_PROG_REQ_yes_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("ESD_PROG_REQ").Value = True
    Call UpdateSummaryPage(Me)
    Call UpdateESDTab

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub obESD_PROG_REQ_no_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("ESD_PROG_REQ").Value = False
    Call UpdateSummaryPage(Me)
    Call UpdateESDTab

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Public Sub SetESD_MARSH_CAB_REQ()
On Error GoTo ErrHandler

    Me.Controls("obESD_MARSH_CAB_REQ_yes").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("ESD_MARSH_CAB_REQ").Value
    'set to the opposite of soft i/o
    Me.Controls("obESD_MARSH_CAB_REQ_no").Value = Not (Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("ESD_MARSH_CAB_REQ").Value)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Public Sub SetESD_SYSTEM_REQ()
On Error GoTo ErrHandler
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    If (.Range("ESD_SYSTEM_REQ").Value = 1) Then
        Me.Controls("obESD_SYSTEM_ESD_FG").Value = True
        Me.Controls("obESD_SYSTEM_BMS").Value = False
        Me.Controls("obESD_SYSTEM_TMC").Value = False
    ElseIf (.Range("ESD_SYSTEM_REQ").Value = 2) Then
        Me.Controls("obESD_SYSTEM_ESD_FG").Value = False
        Me.Controls("obESD_SYSTEM_BMS").Value = True
        Me.Controls("obESD_SYSTEM_TMC").Value = False
    Else
        Me.Controls("obESD_SYSTEM_ESD_FG").Value = False
        Me.Controls("obESD_SYSTEM_BMS").Value = False
        Me.Controls("obESD_SYSTEM_TMC").Value = True
    End If
End With

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub obESD_SYSTEM_ESD_FG_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("SystemType").Value = "ESD"
    Call UpdateSummaryPage(Me)
    Call UpdateESDTab

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub obESD_SYSTEM_BMS_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("SystemType").Value = "BMS"
    Call UpdateSummaryPage(Me)
    Call UpdateESDTab

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub obESD_SYSTEM_TMC_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("SystemType").Value = "TMC"
    Call UpdateSummaryPage(Me)
    Call UpdateESDTab

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

'#####################################################
' Added by JFR 14/09/2007
'ESD
Public Sub SetTMC_DEFAULT()
On Error GoTo ErrHandler

    Call Set_PRVControls_REQ
    Call Set_SurgeControl_REQ
    Call Set_LoadSharing_REQ
    Call Set_Autosynchronization_REQ
    Call Set_PowerSystemStabilizer_REQ
    Call Set_LoadSharingGen_REQ
    
CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Public Sub Set_PRVControls_REQ()
On Error GoTo ErrHandler

    Me.Controls("ob_PRVControls_yes").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("PRV_Controls_REQ").Value
    'set to the opposite of soft i/o
    Me.Controls("ob_PRVControls_no").Value = Not (Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("PRV_Controls_REQ").Value)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub ob_PRVControls_yes_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("PRV_Controls_REQ").Value = True
    Call UpdateSummaryPage(Me)
    Call UpdateTMCTab

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub ob_PRVControls_no_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("PRV_Controls_REQ").Value = False
    Call UpdateSummaryPage(Me)
    Call UpdateTMCTab

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Public Sub Set_SurgeControl_REQ()
On Error GoTo ErrHandler

    Me.Controls("ob_SurgeControl_yes").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("SurgeControl_REQ").Value
    'set to the opposite of soft i/o
    Me.Controls("ob_SurgeControl_no").Value = Not (Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("SurgeControl_REQ").Value)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub ob_SurgeControl_yes_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("SurgeControl_REQ").Value = True
    Call UpdateSummaryPage(Me)
    Call UpdateTMCTab

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub ob_SurgeControl_no_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("SurgeControl_REQ").Value = False
    Call UpdateSummaryPage(Me)
    Call UpdateTMCTab

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Public Sub Set_LoadSharing_REQ()
On Error GoTo ErrHandler

    Me.Controls("ob_LoadSharing_yes").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("LoadSharing_REQ").Value
    'set to the opposite of soft i/o
    Me.Controls("ob_LoadSharing_no").Value = Not (Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("LoadSharing_REQ").Value)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub ob_LoadSharing_yes_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("LoadSharing_REQ").Value = True
    Call UpdateSummaryPage(Me)
    Call UpdateTMCTab

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub ob_LoadSharing_no_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("LoadSharing_REQ").Value = False
    Call UpdateSummaryPage(Me)
    Call UpdateTMCTab

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Public Sub Set_Autosynchronization_REQ()
On Error GoTo ErrHandler

    Me.Controls("ob_Autosynchronization_yes").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("Autosynchronization_REQ").Value
    'set to the opposite of soft i/o
    Me.Controls("ob_Autosynchronization_no").Value = Not (Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("Autosynchronization_REQ").Value)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub ob_Autosynchronization_yes_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("Autosynchronization_REQ").Value = True
    Call UpdateSummaryPage(Me)
    Call UpdateTMCTab

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub ob_Autosynchronization_no_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("Autosynchronization_REQ").Value = False
    Call UpdateSummaryPage(Me)
    Call UpdateTMCTab

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Public Sub Set_PowerSystemStabilizer_REQ()
On Error GoTo ErrHandler

    Me.Controls("ob_PowerSystemStabilizer_yes").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("PowerSystemStabilizer_REQ").Value
    'set to the opposite of soft i/o
    Me.Controls("ob_PowerSystemStabilizer_no").Value = Not (Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("PowerSystemStabilizer_REQ").Value)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub ob_PowerSystemStabilizer_yes_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("PowerSystemStabilizer_REQ").Value = True
    Call UpdateSummaryPage(Me)
    Call UpdateTMCTab

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub ob_PowerSystemStabilizer_no_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("PowerSystemStabilizer_REQ").Value = False
    Call UpdateSummaryPage(Me)
    Call UpdateTMCTab

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Public Sub Set_LoadSharingGen_REQ()
On Error GoTo ErrHandler

    Me.Controls("ob_LoadSharingGen_yes").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("LoadSharingGen_REQ").Value
    'set to the opposite of soft i/o
    Me.Controls("ob_LoadSharingGen_no").Value = Not (Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("LoadSharingGen_REQ").Value)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub ob_LoadSharingGen_yes_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("LoadSharingGen_REQ").Value = True
    Call UpdateSummaryPage(Me)
    Call UpdateTMCTab

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub ob_LoadSharingGen_no_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("LoadSharingGen_REQ").Value = False
    Call UpdateSummaryPage(Me)
    Call UpdateTMCTab

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub ob_BoilerFeedwaterPump_yes_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("BoilerFeedwaterPump_REQ").Value = True
    Call UpdateSummaryPage(Me)
    Call UpdateTMCTab

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub ob_BoilerFeedwaterPump_no_Click()
On Error GoTo ErrHandler

    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("BoilerFeedwaterPump_REQ").Value = False
    Call UpdateSummaryPage(Me)
    Call UpdateTMCTab

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub






Private Sub txtAeroderivative_QTY_AfterUpdate()
    Call UpdateTMCTab
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtAPP_1_HOURS_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtAPP_2_HOURS_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtAPP_3_HOURS_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtAPP_4_HOURS_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtAPP_5_HOURS_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtAPP_6_HOURS_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtAPP_7_HOURS_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtAPP_8_HOURS_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtESDNotes_Change()

End Sub

Private Sub txTrainingNotes_Change()

End Sub



Private Sub txtAssumptions_Change()
On Error GoTo ErrHandler

    If Me.Controls("txtAssumptions").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("ASSUMPTIONS").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("ASSUMPTIONS").Value = Me.Controls("txtAssumptions").Value
    End If
    

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp

End Sub

Private Sub txtCAB_CONSOLES_QTY_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtCAB_IO_QTY_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtCAB_MARSH_QTY_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtCAB_PROC_QTY_Change()
    Call UpdateSummaryPage(Me)
End Sub






Private Sub txtCompressor_QTY_AfterUpdate()
    Call UpdateTMCTab
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtDoubleExtraction_QTY_AfterUpdate()
    Call UpdateTMCTab
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtESD_CAB_IO_QTY_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtESD_CAB_MARSH_QTY_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtESD_CAB_PROC_QTY_Change()
    Call UpdateSummaryPage(Me)
End Sub
Private Sub txtControlProcessorNotes_Change()
On Error GoTo ErrHandler

    If Me.Controls("txtControlProcessorNotes").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("CP_NOTES").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("CP_NOTES").Value = Me.Controls("txtControlProcessorNotes").Value
    End If
    

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtCOST_BUYOUT_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtCOST_IA_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtCOURSE_1_HOURS_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtCOURSE_2_HOURS_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtCOURSE_3_HOURS_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtCOURSE_4_HOURS_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtCOURSE_5_HOURS_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtDeviceIntegratorNotes_Change()
On Error GoTo ErrHandler

    If Me.Controls("txtDeviceIntegratorNotes").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DI_NOTES").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DI_NOTES").Value = Me.Controls("txtDeviceIntegratorNotes").Value
    End If
    

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtDocumentationNotes_Change()
On Error GoTo ErrHandler

    If Me.Controls("txtDocumentationNotes").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DOC_NOTES").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DOC_NOTES").Value = Me.Controls("txtDocumentationNotes").Value
    End If
    

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtDURATION_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtDATE_START_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtDATE_END_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtESD_AI_Change()
    Call UpdateSummaryPage(Me)
End Sub
Private Sub txtESD_AI_DESIRED_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtESD_AO_Change()
    Call UpdateSummaryPage(Me)
End Sub
Private Sub txtESD_AO_DESIRED_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtESD_DI_Change()
    Call UpdateSummaryPage(Me)
End Sub
Private Sub txtESD_DI_DESIRED_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtESD_DIGITAL_CTRL_DI_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtESD_DIGITAL_CTRL_DO_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtESD_DO_Change()
    Call UpdateSummaryPage(Me)
End Sub
Private Sub txtESD_DO_DESIRED_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtFreeFormatNotes_Change()
On Error GoTo ErrHandler
    
    If Me.Controls("txtFreeFormatNotes").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("APP_NOTES").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("APP_NOTES").Value = Me.Controls("txtFreeFormatNotes").Value
    End If
    

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtGasTurbine_QTY_AfterUpdate()
    Call UpdateTMCTab
    Call UpdateSummaryPage(Me)
End Sub


Private Sub txtGenerator_QTY_AfterUpdate()
    Call UpdateTMCTab
    Call UpdateSummaryPage(Me)
End Sub


Private Sub txtMEETING_CLOSE_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtMEETING_DESIGN_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtMEETING_KICKOFF_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtMEETING_OTHER_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtMEETING_PROGRESS_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtMeetingNotes_Change()
On Error GoTo ErrHandler

    If Me.Controls("txtMeetingNotes").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("MEETING_NOTES").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("MEETING_NOTES").Value = Me.Controls("txtMeetingNotes").Value
    End If
    

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtMotorDriven_QTY_AfterUpdate()
    Call UpdateTMCTab
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtMultiShaft_QTY_AfterUpdate()
    Call UpdateTMCTab
    Call UpdateSummaryPage(Me)
End Sub



Private Sub txtNO_OF_UNITS_Change()
    If IsNumeric(Me.txtNO_OF_UNITS.Value) Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("NO_OF_UNITS").Value = Me.txtNO_OF_UNITS.Value
        Call UpdateSummaryPage(Me)
        Me.Repaint
    End If
End Sub

Private Sub txtPROPSAL_DATE_AfterUpdate()
    If Not IsDate(Me.Controls("txtPROPSAL_DATE").Value) Then
        Me.Controls("txtPROPSAL_DATE").Value = Year(Now) & "-" & Month(Now) & "-" & Day(Now)
    End If
    Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("PROPSAL_DATE").Value = Year(Me.Controls("txtPROPSAL_DATE").Value) & "-" & Month(Me.Controls("txtPROPSAL_DATE").Value) & "-" & Day(Me.Controls("txtPROPSAL_DATE").Value)
    Me.Controls("txtPROPSAL_DATE").Value = Format(Me.Controls("txtPROPSAL_DATE").Value, "yyyy-mm-dd")
End Sub



Private Sub txtSITE_TASK1_REM_PCT_AfterUpdate()
    If Me.Controls("txtSITE_TASK1_REM_PCT").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEApplicationBasedSheet).Range("SITE_TASK1_REM_PCT").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEApplicationBasedSheet).Range("SITE_TASK1_REM_PCT").Value = FormatPercent(Me.Controls("txtSITE_TASK1_REM_PCT").Value)
    End If
    Call UpdateSiteServices
    Call UpdateSummaryPage(Me)
End Sub



Private Sub txtSITE_TASK2_REM_PCT_AfterUpdate()
    If Me.Controls("txtSITE_TASK2_REM_PCT").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEApplicationBasedSheet).Range("SITE_TASK2_REM_PCT").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEApplicationBasedSheet).Range("SITE_TASK2_REM_PCT").Value = FormatPercent(Me.Controls("txtSITE_TASK2_REM_PCT").Value)
    End If
    Call UpdateSiteServices
    Call UpdateSummaryPage(Me)
End Sub



Private Sub txtSITE_TASK3_REM_PCT_AfterUpdate()
    If Me.Controls("txtSITE_TASK3_REM_PCT").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEApplicationBasedSheet).Range("SITE_TASK3_REM_PCT").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEApplicationBasedSheet).Range("SITE_TASK3_REM_PCT").Value = FormatPercent(Me.Controls("txtSITE_TASK3_REM_PCT").Value)
    End If
    Call UpdateSiteServices
    Call UpdateSummaryPage(Me)
End Sub


Private Sub txtSITE_TASK4_REM_PCT_AfterUpdate()
    If Me.Controls("txtSITE_TASK4_REM_PCT").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEApplicationBasedSheet).Range("SITE_TASK4_REM_PCT").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEApplicationBasedSheet).Range("SITE_TASK4_REM_PCT").Value = FormatPercent(Me.Controls("txtSITE_TASK4_REM_PCT").Value)
    End If
    Call UpdateSiteServices
    Call UpdateSummaryPage(Me)
End Sub


Private Sub txtTEST_CUSTOMER_FAT_AfterUpdate()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtTEST_FAT_PCT_AfterUpdate()
On Error GoTo ErrHandler

    If Me.Controls("txtTEST_FAT_PCT").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("TEST_FAT_PCT").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("TEST_FAT_PCT").Value = FormatPercent(Me.Controls("txtTEST_FAT_PCT").Value)
    End If
    Call UpdateTestingTab
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtTEST_PRE_FAT_PCT_AfterUpdate()
On Error GoTo ErrHandler

    If Me.Controls("txtTEST_PRE_FAT_PCT").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("TEST_PRE_FAT_PCT").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("TEST_PRE_FAT_PCT").Value = FormatPercent(Me.Controls("txtTEST_PRE_FAT_PCT").Value)
    End If
    Call UpdateTestingTab
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtTIME_PREPARE_HW_EST_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtTIME_PREPARE_HW_QTY").Value = Me.Controls("txtTIME_PREPARE_HW_EST").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtTIME_PREPARE_HW_QTY_AfterUpdate()
On Error GoTo ErrHandler
    'TIME_PREPARE_HW_QTY
    If Me.Controls("txtTIME_PREPARE_HW_QTY").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("TIME_PREPARE_HW_QTY").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("TIME_PREPARE_HW_QTY").Value = Format(Me.Controls("txtTIME_PREPARE_HW_QTY").Value, "#,#0.0")
    End If
    Call UpdateScheduleTab
CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp

End Sub

Private Sub txtPRVValves_QTY_AfterUpdate()
    Call UpdateTMCTab
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtTIME_RECEIVE_HW_EST_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtTIME_RECEIVE_HW_QTY").Value = Me.Controls("txtTIME_RECEIVE_HW_EST").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtTIME_RECEIVE_HW_QTY_AfterUpdate()
On Error GoTo ErrHandler
    'TIME_RECEIVE_HW_QTY
    If Me.Controls("txtTIME_RECEIVE_HW_QTY").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("TIME_RECEIVE_HW_QTY").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("TIME_RECEIVE_HW_QTY").Value = Format(Me.Controls("txtTIME_RECEIVE_HW_QTY").Value, "#,#0.0")
    End If
    Call UpdateScheduleTab
CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtRecycleValves_QTY_AfterUpdate()
    Call UpdateTMCTab
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtReheat_QTY_AfterUpdate()
    Call UpdateTMCTab
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtRENTAL_COST_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtREP_CUSTOM_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtREP_MASS_HEAT_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtREP_STD_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtReportNotes_Change()
On Error GoTo ErrHandler

    If Me.Controls("txtReportNotes").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("REPORT_NOTES").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("REPORT_NOTES").Value = Me.Controls("txtReportNotes").Value
    End If
    

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtSingleExtraction_QTY_AfterUpdate()
    Call UpdateTMCTab
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtSingleShaft_QTY_AfterUpdate()
    Call UpdateTMCTab
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtSITE_COMM_HOURS_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtSITE_PWRUP_HOURS_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtSITE_SAT_HOURS_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtSITE_SURVEY_HOURS_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtSiteServicesNotes_Change()
On Error GoTo ErrHandler
    
    If Me.Controls("txtSiteServicesNotes").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("SITE_NOTES").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("SITE_NOTES").Value = Me.Controls("txtSiteServicesNotes").Value
    End If
    

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtSteamTurbine_QTY_AfterUpdate()
    Call UpdateTMCTab
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtSYS_WORKSTATIONS_QTY_Change()
    Call UpdateSummaryPage(Me)
End Sub
Private Sub txtSYS_CONTROLLERS_QTY_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtSYS_FBM_QTY_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtSYS_FDSI_QTY_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtESD_SYSTEMS_QTY_Change()
    Call UpdateSummaryPage(Me)
End Sub
Private Sub txtESD_CHASSIS_QTY_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtESD_IO_CARD_QTY_Change()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtESD_COMM_QTY_Change()
    Call UpdateSummaryPage(Me)
End Sub


Private Sub txtSystemNotes_Change()
On Error GoTo ErrHandler
    
    If Me.Controls("txtSystemNotes").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("SYSTEM_NOTES").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("SYSTEM_NOTES").Value = Me.Controls("txtSystemNotes").Value
    End If
    

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub



'+JFR 2007-03-01
Private Sub UpdateAssumptionsTab()
On Error GoTo ErrHandler
'set all but the bound controls on this tab
Me.Controls("txtToolkitFactor").Value = Format(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("TOOLKIT_FACTOR").Value, "0%")

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
'-JFR
'################################################
'CP tab stuff
Private Sub UpdateControlProcessorTab()
On Error GoTo ErrHandler
'set all but the bound controls on this tab
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    Me.Controls("txtCP_ANA_COMPLEX_QTY").Value = Format(.Range("CP_ANA_COMPLEX_QTY").Value, "0%")
    Me.Controls("txtCP_DIGITAL_COMPLEX_QTY").Value = Format(.Range("CP_DIGITAL_COMPLEX_QTY").Value, "0%")
    Me.Controls("txtCP_FIELDBUS_IO_QTY").Value = Format(.Range("CP_FIELDBUS_IO_QTY").Value, "0%")
    Me.Controls("txtCP_GRP_START_COMPLEX_QTY").Value = Format(.Range("CP_GRP_START_COMPLEX_QTY").Value, "0%")
    Me.Controls("txtCP_GRP_START_LOOP_QTY").Value = Format(.Range("CP_GRP_START_LOOP_QTY").Value, "0%")
    Me.Controls("txtCP_SEQ_COMPLEX_QTY").Value = Format(.Range("CP_SEQ_COMPLEX_QTY").Value, "0%")
    Me.Controls("txtCP_SEQ_LOOP_QTY").Value = Format(.Range("CP_SEQ_LOOP_QTY").Value, "0%")
   
    Me.Controls("txtCP_DIGITAL_CTRL_DI").Value = .Range("CP_DIGITAL_CTRL_DI").Value
    Me.Controls("txtCP_DIGITAL_CTRL_DO").Value = .Range("CP_DIGITAL_CTRL_DO").Value

  End With

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
'bound controls
Private Sub txtCP_AI_AfterUpdate()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtCP_AO_AfterUpdate()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtCP_DI_AfterUpdate()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtCP_DO_AfterUpdate()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtTIME_KICKOFF_EST_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtTIME_KICKOFF_QTY").Value = Me.Controls("txtTIME_KICKOFF_EST").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtTIME_KICKOFF_QTY_AfterUpdate()
On Error GoTo ErrHandler
    'TIME_KICKOFF_QTY
    If Me.Controls("txtTIME_KICKOFF_QTY").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("TIME_KICKOFF_QTY").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("TIME_KICKOFF_QTY").Value = Format(Me.Controls("txtTIME_KICKOFF_QTY").Value, "#,#0.0")
    End If
    'Call UpdateScheduleTab(Me)
CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

'+JFR 2007-03-01
Private Sub txtToolkitFactor_AfterUpdate()
On Error GoTo ErrHandler
    'TOOLKIT FACTOR
    If Me.Controls("txtToolkitFactor").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("TOOLKIT_FACTOR").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("TOOLKIT_FACTOR").Value = FormatPercent(Me.Controls("txtToolkitFactor").Value)
    End If
    Call UpdateSummaryPage(Me)
CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub


Private Sub txtCP_DIGITAL_CTRL_DI_AfterUpdate()

On Error GoTo ErrHandler
    'CP_DIGITAL_CTRL_DI
    If Me.Controls("txtCP_DIGITAL_CTRL_DI").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("CP_DIGITAL_CTRL_DI").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("CP_DIGITAL_CTRL_DI").Value = Format(Me.Controls("txtCP_DIGITAL_CTRL_DI").Value)
    End If
    
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtCP_DIGITAL_CTRL_DO_AfterUpdate()
On Error GoTo ErrHandler
    'CP_DIGITAL_CTRL_DO
    If Me.Controls("txtCP_DIGITAL_CTRL_DO").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("CP_DIGITAL_CTRL_DO").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("CP_DIGITAL_CTRL_DO").Value = Format(Me.Controls("txtCP_DIGITAL_CTRL_DO").Value)
    End If
    
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

'the unbound controls are set through code
Private Sub txtCP_ANA_COMPLEX_QTY_AfterUpdate()
On Error GoTo ErrHandler
    'CP_ANA_COMPLEX_QTY
    If Me.Controls("txtCP_ANA_COMPLEX_QTY").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("CP_ANA_COMPLEX_QTY").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("CP_ANA_COMPLEX_QTY").Value = FormatPercent(Me.Controls("txtCP_ANA_COMPLEX_QTY").Value)
    End If
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtCP_DIGITAL_COMPLEX_QTY_AfterUpdate()
On Error GoTo ErrHandler
    'CP_DIGITAL_COMPLEX_QTY
    If Me.Controls("txtCP_DIGITAL_COMPLEX_QTY").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("CP_DIGITAL_COMPLEX_QTY").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("CP_DIGITAL_COMPLEX_QTY").Value = FormatPercent(Me.Controls("txtCP_DIGITAL_COMPLEX_QTY").Value)
    End If
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtCP_FIELDBUS_IO_QTY_AfterUpdate()
On Error GoTo ErrHandler
    'CP_FIELDBUS_IO_QTY
    If Me.Controls("txtCP_FIELDBUS_IO_QTY").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("CP_FIELDBUS_IO_QTY").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("CP_FIELDBUS_IO_QTY").Value = FormatPercent(Me.Controls("txtCP_FIELDBUS_IO_QTY").Value)
    End If
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtCP_GRP_START_COMPLEX_QTY_AfterUpdate()
On Error GoTo ErrHandler
    'CP_GRP_START_COMPLEX_QTY
    If Me.Controls("txtCP_GRP_START_COMPLEX_QTY").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("CP_GRP_START_COMPLEX_QTY").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("CP_GRP_START_COMPLEX_QTY").Value = FormatPercent(Me.Controls("txtCP_GRP_START_COMPLEX_QTY").Value)
    End If
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtCP_GRP_START_LOOP_QTY_AfterUpdate()
On Error GoTo ErrHandler
    'CP_GRP_START_LOOP_QTY
    If Me.Controls("txtCP_GRP_START_LOOP_QTY").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("CP_GRP_START_LOOP_QTY").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("CP_GRP_START_LOOP_QTY").Value = FormatPercent(Me.Controls("txtCP_GRP_START_LOOP_QTY").Value)
    End If
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtCP_SEQ_COMPLEX_QTY_AfterUpdate()
On Error GoTo ErrHandler
    'CP_SEQ_COMPLEX_QTY
    If Me.Controls("txtCP_SEQ_COMPLEX_QTY").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("CP_SEQ_COMPLEX_QTY").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("CP_SEQ_COMPLEX_QTY").Value = FormatPercent(Me.Controls("txtCP_SEQ_COMPLEX_QTY").Value)
    End If
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtCP_SEQ_LOOP_QTY_AfterUpdate()
On Error GoTo ErrHandler
    'CP_SEQ_LOOP_QTY
    If Me.Controls("txtCP_SEQ_LOOP_QTY").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("CP_SEQ_LOOP_QTY").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("CP_SEQ_LOOP_QTY").Value = FormatPercent(Me.Controls("txtCP_SEQ_LOOP_QTY").Value)
    End If
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

'################################################
'DI tab stuff
'recalc after each update so user will see realtime calc
Private Sub txtDI_AI_AfterUpdate()
    Call UpdateDeviceIntegrationTab
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtDI_AO_AfterUpdate()
    Call UpdateDeviceIntegrationTab
    Call UpdateSummaryPage(Me)
End Sub
Private Sub txtDI_DIGITAL_CTRL_DI_AfterUpdate()
On Error GoTo ErrHandler
    'DI_DIGITAL_CTRL_DI
    If Me.Controls("txtDI_DIGITAL_CTRL_DI").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("DI_DIGITAL_CTRL_DI").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("DI_DIGITAL_CTRL_DI").Value = Format(Me.Controls("txtDI_DIGITAL_CTRL_DI").Value)
    End If
    
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtDI_DIGITAL_CTRL_DO_AfterUpdate()
On Error GoTo ErrHandler
    'DI_DIGITAL_CTRL_DO
    If Me.Controls("txtDI_DIGITAL_CTRL_DO").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("DI_DIGITAL_CTRL_DO").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("DI_DIGITAL_CTRL_DO").Value = Format(Me.Controls("txtDI_DIGITAL_CTRL_DO").Value)
    End If
    
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub txtDI_COMPLEX_QTY_AfterUpdate()
On Error GoTo ErrHandler

    If Me.Controls("txtDI_COMPLEX_QTY").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("DI_COMPLEX_QTY").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("DI_COMPLEX_QTY").Value = FormatPercent(Me.Controls("txtDI_COMPLEX_QTY").Value)
    End If
    Call UpdateDeviceIntegrationTab
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtDI_DEVICES_AfterUpdate()
    Call UpdateDeviceIntegrationTab
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtDI_DI_AfterUpdate()
    Call UpdateDeviceIntegrationTab
    Call UpdateSummaryPage(Me)
End Sub


Private Sub txtDI_DO_AfterUpdate()
    Call UpdateDeviceIntegrationTab
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtDI_GRP_START_COMPLEX_QTY_AfterUpdate()
On Error GoTo ErrHandler

    If Me.Controls("txtDI_GRP_START_COMPLEX_QTY").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("DI_GRP_START_COMPLEX_QTY").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("DI_GRP_START_COMPLEX_QTY").Value = FormatPercent(Me.Controls("txtDI_GRP_START_COMPLEX_QTY").Value)
    End If
    Call UpdateDeviceIntegrationTab
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtDI_GRP_START_LOOP_QTY_AfterUpdate()
On Error GoTo ErrHandler

    If Me.Controls("txtDI_GRP_START_LOOP_QTY").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("DI_GRP_START_LOOP_QTY").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("DI_GRP_START_LOOP_QTY").Value = FormatPercent(Me.Controls("txtDI_GRP_START_LOOP_QTY").Value)
    End If
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtDI_INTERFACES_AfterUpdate()
    Call UpdateDeviceIntegrationTab
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtDI_SEQ_COMPLEX_QTY_AfterUpdate()
On Error GoTo ErrHandler

    If Me.Controls("txtDI_SEQ_COMPLEX_QTY").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("DI_SEQ_COMPLEX_QTY").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("DI_SEQ_COMPLEX_QTY").Value = FormatPercent(Me.Controls("txtDI_SEQ_COMPLEX_QTY").Value)
    End If
    Call UpdateDeviceIntegrationTab
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtDI_SEQ_LOOP_QTY_AfterUpdate()
On Error GoTo ErrHandler

    If Me.Controls("txtDI_SEQ_LOOP_QTY").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("DI_SEQ_LOOP_QTY").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("DI_SEQ_LOOP_QTY").Value = FormatPercent(Me.Controls("txtDI_SEQ_LOOP_QTY").Value)
    End If
    Call UpdateDeviceIntegrationTab
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub txtESD_AI_AfterUpdate()
    Call UpdateESDTab
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtESD_AO_AfterUpdate()
    Call UpdateESDTab
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtESD_DI_AfterUpdate()
    Call UpdateESDTab
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtESD_DO_AfterUpdate()
    Call UpdateESDTab
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtESD_COMPLEX_QTY_AfterUpdate()
On Error GoTo ErrHandler

    If Me.Controls("txtESD_COMPLEX_QTY").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("ESD_COMPLEX_QTY").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("ESD_COMPLEX_QTY").Value = FormatPercent(Me.Controls("txtESD_COMPLEX_QTY").Value)
    End If
    Call UpdateESDTab
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub txtESD_MISC_CAB_QTY_AfterUpdate()
On Error GoTo ErrHandler

    If Me.Controls("txtESD_MISC_CAB_QTY").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("ESD_MISC_CAB_QTY").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("ESD_MISC_CAB_QTY").Value = Me.Controls("txtESD_MISC_CAB_QTY").Value
    End If
    Call UpdateESDTab
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtESD_GRP_START_COMPLEX_QTY_AfterUpdate()
On Error GoTo ErrHandler

    If Me.Controls("txtESD_GRP_START_COMPLEX_QTY").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("ESD_GRP_START_COMPLEX_QTY").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("ESD_GRP_START_COMPLEX_QTY").Value = FormatPercent(Me.Controls("txtESD_GRP_START_COMPLEX_QTY").Value)
    End If
    Call UpdateESDTab
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtESD_GRP_START_LOOP_QTY_AfterUpdate()
On Error GoTo ErrHandler

    If Me.Controls("txtESD_GRP_START_LOOP_QTY").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("ESD_GRP_START_LOOP_QTY").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("ESD_GRP_START_LOOP_QTY").Value = FormatPercent(Me.Controls("txtESD_GRP_START_LOOP_QTY").Value)
    End If
    Call UpdateESDTab
    Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub txtTestingNotes_Change()
On Error GoTo ErrHandler

    If Me.Controls("txtTestingNotes").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("TEST_NOTES").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("TEST_NOTES").Value = Me.Controls("txtTestingNotes").Value
    End If
    

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtTL_AIRFARE_AfterUpdate()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtTL_DAILY_ALLOW_AfterUpdate()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtTL_DAYS_QTY_AfterUpdate()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtTL_PERS_QTY_AfterUpdate()
    Call UpdateSummaryPage(Me)
End Sub

Private Sub txtTL_TRIPS_REQ_AfterUpdate()
    Call UpdateSummaryPage(Me)
End Sub














Private Sub txtTrainingNotes_Change()
On Error GoTo ErrHandler
    
    If Me.Controls("txtTrainingNotes").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("TRAIN_NOTES").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("TRAIN_NOTES").Value = Me.Controls("txtTrainingNotes").Value
    End If
    

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtTIME_TRANSPORT_HW_EST_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtTIME_TRANSPORT_HW_QTY").Value = Me.Controls("txtTIME_TRANSPORT_HW_EST").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtTIME_TRANSPORT_HW_QTY_AfterUpdate()
On Error GoTo ErrHandler
    'TIME_TRANSPORT_HW_QTY
    If Me.Controls("txtTIME_TRANSPORT_HW_QTY").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("TIME_TRANSPORT_HW_QTY").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("TIME_TRANSPORT_HW_QTY").Value = Format(Me.Controls("txtTIME_TRANSPORT_HW_QTY").Value, "#,#0.0")
    End If
    Call UpdateScheduleTab
CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub



Private Sub txtTravelLivingNotes_Change()
On Error GoTo ErrHandler
    
    If Me.Controls("txtTravelLivingNotes").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("TL_NOTES").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("TL_NOTES").Value = Me.Controls("txtTravelLivingNotes").Value
    End If
    

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtTriconNotes_Change()
On Error GoTo ErrHandler

    If Me.Controls("txtTriconNotes").Text = "" Then
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("TRICON_NOTES").Value = ""
    Else
        Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("TRICON_NOTES").Value = Me.Controls("txtTriconNotes").Value
    End If
    

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

'############################################
Private Sub UserForm_Activate()
On Error GoTo ErrHandler
'set the workbook name and close if we cannot set it
If getWorkBookName = True Then
    'When opening, only update price one time
    OpenDataEntryForm = True

    'fill in the calculated fields on summary tab
    Call ShowLoadingThenContinue
    Call CPIOTotal
    Call DIIOTotal
    Call EDSIOTotal
    'set unit labels from data entry sheet
    Call SetUnitLabels(Me)
    'set the tooltips from the comments in the sheet
    Call GetToolTips(Me)
    'set the default values from the sheet
    Call SetDefaults(Me)
    'set the firts tab as the visible one on open
    Me.Controls("MultiPage1").Value = 0
     'set the rowsource and controlsource for the Scope language dropdowns
    Me.Controls("cboScopeLanguage").RowSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECECurrencySheet).Range("LANGUAGE_LIST").Address(External:=True)
    Me.Controls("cboScopeLanguage").ControlSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("SCOPE_LANGUAGE").Address(External:=True)
    'set the rowsource and controlsource for the two country dropdowns
    Me.Controls("cboLocalCuntry").RowSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECECurrencySheet).Range("COUNTRY_LIST").Address(External:=True)
    Me.Controls("cboLocalCuntry").ControlSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("LOCAL_COUNTRY").Address(External:=True)
    Me.Controls("cboDefaultRemoteCountry").RowSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECECurrencySheet).Range("COUNTRY_LIST").Address(External:=True)
    Me.Controls("cboDefaultRemoteCountry").ControlSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Address(External:=True)
     'set the rowsource and controlsource for the industry dropdowns
    Me.Controls("cboIndustry").RowSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEIndustrySheet).Range("INDUSTRY_LIST").Address(External:=True)
    Me.Controls("cboIndustry").ControlSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("INDUSTRY").Address(External:=True)
     'set the rowsource and controlsource for the toolkit dropdowns
    Me.Controls("cboToolkit").RowSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEToolkitSheet).Range("TOOLKIT_LIST_ROW").Address(External:=True)
    Me.Controls("cboToolkit").ControlSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("TOOLKIT").Address(External:=True)
     'set the rowsource and controlsource for the customer type dropdowns
    Me.Controls("cboCustomerType").RowSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEIndustrySheet).Range("CUSTOMER_TYPE_LIST").Address(External:=True)
    Me.Controls("cboCustomerType").ControlSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("CUSTOMER_TYPE").Address(External:=True)
    
    Call SetWPA(Me)

        
    '2005/10/07
    'BUG line feed in notes field cause excel to continuously recalculate so need to unbind these fields
    '    Me.Controls("txtAssumptions").ControlSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("ASSUMPTIONS").Address(external:=True)
    '    Me.Controls("txtCPNotes").ControlSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("CP_NOTES").Address(external:=True)
    '    Me.Controls("txtDINotes").ControlSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DI_NOTES").Address(external:=True)
    '    Me.Controls("txtTriconNotes").ControlSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("TRICON_NOTES").Address(external:=True)
    '    Me.Controls("txtReportNotes").ControlSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("REPORT_NOTES").Address(external:=True)
    '    Me.Controls("txtFreeFormatNotes").ControlSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("APP_NOTES").Address(external:=True)
    '    Me.Controls("txtTrainNotes").ControlSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("TRAIN_NOTES").Address(external:=True)
    '    Me.Controls("txtSiteNotes").ControlSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("SITE_NOTES").Address(external:=True)
    '    Me.Controls("txtTestNotes").ControlSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("TEST_NOTES").Address(external:=True)
    '    Me.Controls("txtDocNotes").ControlSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DOC_NOTES").Address(external:=True)
    '    Me.Controls("txtTLNotes").ControlSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("TL_NOTES").Address(external:=True)
    '    Me.Controls("txtMeetingNotes").ControlSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("MEETING_NOTES").Address(external:=True)
    '    Me.Controls("txtSystemNotes").ControlSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("SYSTEM_NOTES").Address(external:=True)
    
    Me.Controls("txtAssumptions").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("ASSUMPTIONS").Value
    Me.Controls("txtControlProcessorNotes").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("CP_NOTES").Value
    Me.Controls("txtDeviceIntegratorNotes").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DI_NOTES").Value
    Me.Controls("txtTriconNotes").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("TRICON_NOTES").Value
    Me.Controls("txtTestingNotes").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("TEST_NOTES").Value
    Me.Controls("txtDocumentationNotes").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DOC_NOTES").Value
    Me.Controls("txtMeetingNotes").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("MEETING_NOTES").Value
    Me.Controls("txtReportNotes").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("REPORT_NOTES").Value
    Me.Controls("txtFreeFormatNotes").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("APP_NOTES").Value
    Me.Controls("txtTrainingNotes").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("TRAIN_NOTES").Value
    Me.Controls("txtSiteServicesNotes").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("SITE_NOTES").Value
    Me.Controls("txtTravelLivingNotes").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("TL_NOTES").Value
    Me.Controls("txtSystemNotes").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("SYSTEM_NOTES").Value

    
    
    'set the version on the about tab
    Me.Controls("lblAboutVersion").Caption = "GECE Rev " & GECEXLSVERSION
    
    'set the default value for group Existing PFS on documentation tab
    If Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("CUSTOMER_SPEC_REQ").Value Then
        Me.Controls("obCUSTOMER_SPEC_REQ_yes").Value = True
    Else
        Me.Controls("obCUSTOMER_SPEC_REQ_no").Value = True
    End If
    
    '#### TABS
    'update the recomended values on the Free Format Tab
    Call UpdateFreeFormatTab
    'update default values for Device Integration Tab
    Call UpdateDeviceIntegrationTab
    'set cp tab
    Call UpdateControlProcessorTab
    'testing tab
    Call UpdateTestingTab
    'Assumptions tag
    Call UpdateAssumptionsTab
    Call UpdateScheduleTab
    
    'set hard or soft io
    Call setDI_IOTYPE_STS
    Call SetESD_DEFAULT
    Call SetTMC_DEFAULT
    Call setHIGH_RISK_SITE
    
    OpenDataEntryForm = False
    Call UpdateSummaryPage(Me)
    
Else
    Call Unload(Me)
End If

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
    Resume Next
End Sub

Private Sub MultiPage1_Change()
On Error GoTo ErrHandler
'all
If OpenDataEntryForm = False Then
    Call CPIOTotal
    Call DIIOTotal
    Call EDSIOTotal
    Call UpdateSummaryPage(Me)
    Call SetDefaults(Me)
    Call SetUnitLabels(Me)
    
    'set hard or soft io
    Call setDI_IOTYPE_STS
    Call SetESD_DEFAULT
    Call SetTMC_DEFAULT
    
    'these are just for when an individual tab needs to be updated
    Select Case Me.Controls("MultiPage1").SelectedItem.Name
        Case "PageControlProcessor"
            Call UpdateControlProcessorTab
        Case "PageDeviceIntegrator"
            Call UpdateDeviceIntegrationTab
        Case "PageTesting"
            Call UpdateTestingTab
        Case "PageFreeFormat"
            Call UpdateFreeFormatTab
        Case "PageESD" 'ESD
             Call UpdateESDTab
        Case "PageTMC" 'TMC
             Call UpdateTMCTab
        'Case "PageTravelLiving"
        'Case "PageCustomTraining"
        'Case "PageDocumentation"
        'Case "PageMeetings"
        'Case "PageReport"
        Case "PageSiteServices"
            Call UpdateSiteServices
        'Case "PageSummary"
        'Case "PageSystem"
        '+JFR
        Case "PageAssumptions" 'header
            Call UpdateAssumptionsTab
        '-JFR
        'Case "PageAbout"
        Case Else
    End Select
End If

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub


Private Sub cmdClose_Click()
    'Me.Hide
    Call SaveAndExit
End Sub

'reset default, copies the value from the rcomended value col to the data entry col
'CP
Private Sub txtCP_DIGITAL_CTRL_DI_REC_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    .Range("CP_DIGITAL_CTRL_DI").Value = .Range("CP_DIGITAL_CTRL_DI_REC").Value
    Me.Controls("txtCP_DIGITAL_CTRL_DI").Value = .Range("CP_DIGITAL_CTRL_DI_REC").Value
End With

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub txtCP_DIGITAL_CTRL_DO_REC_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    .Range("CP_DIGITAL_CTRL_DO").Value = .Range("CP_DIGITAL_CTRL_DO_REC").Value
    Me.Controls("txtCP_DIGITAL_CTRL_DO").Value = .Range("CP_DIGITAL_CTRL_DO_REC").Value
End With

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub txtCP_ANA_COMPLEX_REC_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    .Range("CP_ANA_COMPLEX_QTY").Value = .Range("CP_ANA_COMPLEX_REC").Value
    Me.Controls("txtCP_ANA_COMPLEX_QTY").Value = Format(.Range("CP_ANA_COMPLEX_REC").Value, "0%")
End With

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtCP_DIGITAL_COMPLEX_REC_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    .Range("CP_DIGITAL_COMPLEX_QTY").Value = .Range("CP_DIGITAL_COMPLEX_REC").Value
    Me.Controls("txtCP_DIGITAL_COMPLEX_QTY").Value = Format(.Range("CP_DIGITAL_COMPLEX_REC").Value, "0%")
End With

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtCP_FIELDBUS_IO_REC_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    .Range("CP_FIELDBUS_IO_QTY").Value = .Range("CP_FIELDBUS_IO_REC").Value
    Me.Controls("txtCP_FIELDBUS_IO_QTY").Value = Format(.Range("CP_FIELDBUS_IO_REC").Value, "0%")
End With

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtCP_GRP_START_COMPLEX_REC_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    .Range("CP_GRP_START_COMPLEX_QTY").Value = .Range("CP_GRP_START_COMPLEX_REC").Value
    Me.Controls("txtCP_GRP_START_COMPLEX_QTY").Value = Format(.Range("CP_GRP_START_COMPLEX_REC").Value, "0%")
End With

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtCP_GRP_START_LOOP_REC_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    .Range("CP_GRP_START_LOOP_QTY").Value = .Range("CP_GRP_START_LOOP_REC").Value
    Me.Controls("txtCP_GRP_START_LOOP_QTY").Value = Format(.Range("CP_GRP_START_LOOP_REC").Value, "0%")
End With

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtCP_SEQ_COMPLEX_REC_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    .Range("CP_SEQ_COMPLEX_QTY").Value = .Range("CP_SEQ_COMPLEX_REC").Value
    Me.Controls("txtCP_SEQ_COMPLEX_QTY").Value = Format(.Range("CP_SEQ_COMPLEX_REC").Value, "0%")
End With

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtCP_SEQ_LOOP_REC_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    .Range("CP_SEQ_LOOP_QTY").Value = .Range("CP_SEQ_LOOP_REC").Value
    Me.Controls("txtCP_SEQ_LOOP_QTY").Value = Format(.Range("CP_SEQ_LOOP_REC").Value, "0%")
End With

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

'DI
Private Sub txtDI_DIGITAL_CTRL_DI_REC_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    .Range("DI_DIGITAL_CTRL_DI").Value = .Range("DI_DIGITAL_CTRL_DI_REC").Value
    Me.Controls("txtDI_DIGITAL_CTRL_DI").Value = Range("DI_DIGITAL_CTRL_DI_REC").Value
End With

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub txtDI_DIGITAL_CTRL_DO_REC_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    .Range("DI_DIGITAL_CTRL_DO").Value = .Range("DI_DIGITAL_CTRL_DO_REC").Value
    Me.Controls("txtDI_DIGITAL_CTRL_DO").Value = Range("DI_DIGITAL_CTRL_DO_REC").Value
End With

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub txtDI_COMPLEX_REC_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    .Range("DI_COMPLEX_QTY").Value = .Range("DI_COMPLEX_REC").Value
    Me.Controls("txtDI_COMPLEX_QTY").Value = Format(.Range("DI_COMPLEX_REC").Value, "0%")
End With

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtDI_SEQ_LOOP_REC_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    .Range("DI_SEQ_LOOP_QTY").Value = .Range("DI_SEQ_LOOP_REC").Value
    Me.Controls("txtDI_SEQ_LOOP_QTY").Value = Format(.Range("DI_SEQ_LOOP_REC").Value, "0%")
End With

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtDI_SEQ_COMPLEX_REC_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    .Range("DI_SEQ_COMPLEX_QTY").Value = .Range("DI_SEQ_COMPLEX_REC").Value
    Me.Controls("txtDI_SEQ_COMPLEX_QTY").Value = Format(.Range("DI_SEQ_COMPLEX_REC").Value, "0%")
End With

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtDI_GRP_START_LOOP_REC_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    .Range("DI_GRP_START_LOOP_QTY").Value = .Range("DI_GRP_START_LOOP_REC").Value
    Me.Controls("txtDI_GRP_START_LOOP_QTY").Value = Format(.Range("DI_GRP_START_LOOP_REC").Value, "0%")
End With

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtDI_GRP_START_COMPLEX_REC_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    .Range("DI_GRP_START_COMPLEX_QTY").Value = .Range("DI_GRP_START_COMPLEX_REC").Value
    Me.Controls("txtDI_GRP_START_COMPLEX_QTY").Value = Format(.Range("DI_GRP_START_COMPLEX_REC").Value, "0%")
End With

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
'ESD
Private Sub txtESD_COMPLEX_REC_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    .Range("ESD_COMPLEX_QTY").Value = .Range("ESD_COMPLEX_REC").Value
    Me.Controls("txtESD_COMPLEX_QTY").Value = Format(.Range("ESD_COMPLEX_REC").Value, "0%")
End With

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub txtESD_MISC_CAB_REC_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    .Range("ESD_MISC_CAB_QTY").Value = .Range("ESD_MISC_CAB_REC").Value
    Me.Controls("txtESD_MISC_CAB_QTY").Value = Format(.Range("ESD_MISC_CAB_REC").Value, "#,##0.0%")
End With

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub txtESD_GRP_START_LOOP_REC_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    .Range("ESD_GRP_START_LOOP_QTY").Value = .Range("ESD_GRP_START_LOOP_REC").Value
    Me.Controls("txtESD_GRP_START_LOOP_QTY").Value = Format(.Range("ESD_GRP_START_LOOP_REC").Value, "0%")
End With

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtESD_GRP_START_COMPLEX_REC_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    .Range("ESD_GRP_START_COMPLEX_QTY").Value = .Range("ESD_GRP_START_COMPLEX_REC").Value
    Me.Controls("txtESD_GRP_START_COMPLEX_QTY").Value = Format(.Range("ESD_GRP_START_COMPLEX_REC").Value, "0%")
End With

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
'report
Private Sub txtREP_STD_EST_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtREP_STD").Value = Me.Controls("txtREP_STD_EST").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
'site services
Private Sub txtSITE_SAT_EST_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtSITE_SAT_HOURS").Value = Me.Controls("txtSITE_SAT_EST").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

'testing
Private Sub txtTEST_PRE_FAT_REC_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    .Range("TEST_PRE_FAT_PCT").Value = .Range("TEST_PRE_FAT_REC").Value
    Me.Controls("txtTEST_PRE_FAT_PCT").Value = Format(.Range("TEST_PRE_FAT_REC").Value, "###0%")
End With

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtTEST_FAT_REC_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    .Range("TEST_FAT_PCT").Value = .Range("TEST_FAT_REC").Value
    Me.Controls("txtTEST_FAT_PCT").Value = Format(.Range("TEST_FAT_REC").Value, "0%")
End With

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

'this one got its default from a range and not the default cell see set defaulst sub
Private Sub txtTEST_FAT_QTY_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtTEST_CUSTOMER_FAT").Value = Me.Controls("txtTEST_FAT_QTY").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

'meeting
Private Sub txtMEETING_KICKOFF_EST_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtMEETING_KICKOFF").Value = Me.Controls("txtMEETING_KICKOFF_EST").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub txtMEETING_DESIGN_EST_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtMEETING_DESIGN").Value = Me.Controls("txtMEETING_DESIGN_EST").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub txtMEETING_PROGRESS_EST_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtMEETING_PROGRESS").Value = Me.Controls("txtMEETING_PROGRESS_EST").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub txtMEETING_CLOSE_EST_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtMEETING_CLOSE").Value = Me.Controls("txtMEETING_CLOSE_EST").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
'system summary
Private Sub txtSYS_WORKSTATIONS_EST_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtSYS_WORKSTATIONS_QTY").Value = Me.Controls("txtSYS_WORKSTATIONS_EST").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtSYS_CONTROLLERS_EST_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtSYS_CONTROLLERS_QTY").Value = Me.Controls("txtSYS_CONTROLLERS_EST").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtSYS_FBM_EST_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtSYS_FBM_QTY").Value = Me.Controls("txtSYS_FBM_EST").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtSYS_FDSI_EST_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtSYS_FDSI_QTY").Value = Me.Controls("txtSYS_FDSI_EST").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub txtESD_SYSTEMS_EST_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtESD_SYSTEMS_QTY").Value = Me.Controls("txtESD_SYSTEMS_EST").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtESD_CHASSIS_EST_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtESD_CHASSIS_QTY").Value = Me.Controls("txtESD_CHASSIS_EST").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtESD_IO_CARD_EST_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtESD_IO_CARD_QTY").Value = Me.Controls("txtESD_IO_CARD_EST").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtESD_COMM_EST_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtESD_COMM_QTY").Value = Me.Controls("txtESD_COMM_EST").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub txtCAB_PROC_EST_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtCAB_PROC_QTY").Value = Me.Controls("txtCAB_PROC_EST").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtCAB_IO_EST_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtCAB_IO_QTY").Value = Me.Controls("txtCAB_IO_EST").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtCAB_MARSH_EST_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'remed out 9/30/2005 to avoid adding N/A as a default
'On Error GoTo ErrHandler
'    Me.Controls("txtCAB_MARSH_QTY").Value = Me.Controls("txtCAB_MARSH_EST").Value
'
'CleanUp:
'Exit Sub
'ErrHandler:
'    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
'    Resume CleanUp
End Sub

Private Sub txtESD_CAB_CONSOLES_EST_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtESD_CAB_CONSOLES_QTY").Value = Me.Controls("txtESD_CAB_CONSOLES_EST").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub txtESD_CAB_PROC_EST_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtESD_CAB_PROC_QTY").Value = Me.Controls("txtESD_CAB_PROC_EST").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtESD_CAB_IO_EST_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtESD_CAB_IO_QTY").Value = Me.Controls("txtESD_CAB_IO_EST").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtESD_CAB_MARSH_EST_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtESD_CAB_MARSH_QTY").Value = Me.Controls("txtESD_CAB_MARSH_EST").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub txtDURATION_RECOMMENDED_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtDURATION").Value = Me.Controls("txtDURATION_RECOMMENDED").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtDATE_START_RECOMMENDED_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtDATE_START").Value = Me.Controls("txtDATE_START_RECOMMENDED").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub txtDATE_END_RECOMMENDED_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo ErrHandler
    Me.Controls("txtDATE_END").Value = Me.Controls("txtDATE_END_RECOMMENDED").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub CPIOTotal()
On Error GoTo ErrHandler
    Me.Controls("txtCP_TOT_IO").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("CP_TOT_IO").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub DIIOTotal()
On Error GoTo ErrHandler
    Me.Controls("txtDI_TOT_IO").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("DI_TOT_IO").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub EDSIOTotal()
On Error GoTo ErrHandler
    Me.Controls("txtESD_TOT_IO").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range("ESD_TOT_IO").Value

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub
Private Sub UpdateFreeFormatTab()
On Error GoTo ErrHandler
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet)
    Me.Controls("txtAPP_1_AVG_RATE").Value = Round(.Range("APP_1_AVG_RATE").Value)
    Me.Controls("txtAPP_2_AVG_RATE").Value = Round(.Range("APP_2_AVG_RATE").Value)
    Me.Controls("txtAPP_3_AVG_RATE").Value = Round(.Range("APP_3_AVG_RATE").Value)
    Me.Controls("txtAPP_4_AVG_RATE").Value = Round(.Range("APP_4_AVG_RATE").Value)
    Me.Controls("txtAPP_5_AVG_RATE").Value = Round(.Range("APP_5_AVG_RATE").Value)
    Me.Controls("txtAPP_6_AVG_RATE").Value = Round(.Range("APP_6_AVG_RATE").Value)
    Me.Controls("txtAPP_7_AVG_RATE").Value = Round(.Range("APP_7_AVG_RATE").Value)
    Me.Controls("txtAPP_8_AVG_RATE").Value = Round(.Range("APP_8_AVG_RATE").Value)
End With

'set currency for label
Me.Controls("lblFreeFormatCurrency").Caption = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECECostSheet).Range("LOC_CURRENCY").Text


CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub UpdateDeviceIntegrationTab()
On Error GoTo ErrHandler
'update the calc fields in the ditab
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet)
    Me.Controls("txtDI_AI_QTY").Value = .Range("DI_TASK2_QTY").Value
    Me.Controls("txtDI_AO_QTY").Value = .Range("DI_TASK3_QTY").Value
    Me.Controls("txtDI_DI_QTY").Value = .Range("DI_TASK4_QTY").Value
    Me.Controls("txtDI_DO_QTY").Value = .Range("DI_TASK5_QTY").Value
 
    
End With

'set all but the bound controls on this tab
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    'DI_COMPLEX_QTY
    Me.Controls("txtDI_COMPLEX_QTY").Value = Format(.Range("DI_COMPLEX_QTY").Value, "0%")
    'DI_SEQ_LOOP_QTY
    Me.Controls("txtDI_SEQ_LOOP_QTY").Value = Format(.Range("DI_SEQ_LOOP_QTY").Value, "0%")
    'DI_SEQ_COMPLEX_QTY
    Me.Controls("txtDI_SEQ_COMPLEX_QTY").Value = Format(.Range("DI_SEQ_COMPLEX_QTY").Value, "0%")
    'DI_GRP_START_LOOP_QTY
    Me.Controls("txtDI_GRP_START_LOOP_QTY").Value = Format(.Range("DI_GRP_START_LOOP_QTY").Value, "0%")
    'DI_GRP_START_COMPLEX_QTY
    Me.Controls("txtDI_GRP_START_COMPLEX_QTY").Value = Format(.Range("DI_GRP_START_COMPLEX_QTY").Value, "0%")
    Me.Controls("txtDI_DIGITAL_CTRL_DI").Value = .Range("DI_DIGITAL_CTRL_DI").Value
    Me.Controls("txtDI_DIGITAL_CTRL_DO").Value = .Range("DI_DIGITAL_CTRL_DO").Value
    
End With
Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub UpdateSiteServices()
On Error GoTo ErrHandler
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEApplicationBasedSheet)
    Me.Controls("txtSITE_TASK1_REM_PCT").Value = Format(.Range("SITE_TASK1_REM_PCT").Value, "0%")
    Me.Controls("txtSITE_TASK2_REM_PCT").Value = Format(.Range("SITE_TASK2_REM_PCT").Value, "0%")
    Me.Controls("txtSITE_TASK3_REM_PCT").Value = Format(.Range("SITE_TASK3_REM_PCT").Value, "0%")
    Me.Controls("txtSITE_TASK4_REM_PCT").Value = Format(.Range("SITE_TASK4_REM_PCT").Value, "0%")
End With
Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub


Private Sub UpdateESDTab()
' Added by AB 12/09/2006
On Error GoTo ErrHandler
'update the calc fields in the ESDtab
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet)
    Me.Controls("txtESD_AI_QTY").Value = Format(.Range("ESD_TASK2_QTY").Value, "0")
    Me.Controls("txtESD_AO_QTY").Value = Format(.Range("ESD_TASK3_QTY").Value, "0")
    Me.Controls("txtESD_DI_QTY").Value = Format(.Range("ESD_TASK4_QTY").Value, "0")
    Me.Controls("txtESD_DO_QTY").Value = Format(.Range("ESD_TASK5_QTY").Value, "0")
End With

'set all but the bound controls on this tab
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    'ESD_COMPLEX_QTY
    Me.Controls("txtESD_COMPLEX_QTY").Value = Format(.Range("ESD_COMPLEX_QTY").Value, "0%")
    Me.Controls("txtESD_COMPLEX_REC").Value = Format(.Range("ESD_COMPLEX_REC").Value, "0%")
    'ESD_MISC_CAB_QTY
    Me.Controls("txtESD_MISC_CAB_QTY").Value = Format(.Range("ESD_MISC_CAB_QTY").Value, "#,##0.00%")
    Me.Controls("txtESD_MISC_CAB_REC").Value = Format(.Range("ESD_MISC_CAB_REC").Value, "#,##0.00%")
    'ESD_GRP_START_LOOP_QTY
    Me.Controls("txtESD_GRP_START_LOOP_QTY").Value = Format(.Range("ESD_GRP_START_LOOP_QTY").Value, "0%")
    Me.Controls("txtESD_GRP_START_LOOP_REC").Value = Format(.Range("ESD_GRP_START_LOOP_REC").Value, "0%")
    'ESD_GRP_START_COMPLEX_QTY
    Me.Controls("txtESD_GRP_START_COMPLEX_QTY").Value = Format(.Range("ESD_GRP_START_COMPLEX_QTY").Value, "0%")
    Me.Controls("txtESD_GRP_START_COMPLEX_REC").Value = Format(.Range("ESD_GRP_START_COMPLEX_REC").Value, "0%")
    'EST_SYSTEM_REQ
    If (.Range("ESD_SYSTEM_REQ").Value = 1) Then
        Me.Controls("obESD_SYSTEM_ESD_FG").Value = True
        Me.Controls("obESD_SYSTEM_BMS").Value = False
        Me.Controls("obESD_SYSTEM_TMC").Value = False
    ElseIf (.Range("ESD_SYSTEM_REQ").Value = 2) Then
        Me.Controls("obESD_SYSTEM_ESD_FG").Value = False
        Me.Controls("obESD_SYSTEM_BMS").Value = True
        Me.Controls("obESD_SYSTEM_TMC").Value = False
    Else
        Me.Controls("obESD_SYSTEM_ESD_FG").Value = False
        Me.Controls("obESD_SYSTEM_BMS").Value = False
        Me.Controls("obESD_SYSTEM_TMC").Value = True
    End If
End With
Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub UpdateScheduleTab()
' Added by JFR 19/09/2007
On Error GoTo ErrHandler

'set all but the bound controls on this tab
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    'TIME_KICKOFF_QTY
    Me.Controls("txtTIME_KICKOFF_QTY").Value = Format(.Range("TIME_KICKOFF_QTY").Value, "0")
    Me.Controls("txtTIME_KICKOFF_EST").Value = Format(.Range("TIME_KICKOFF_EST").Value, "#,#0.0")
    'TIME_RECEIVE_HW_QTY
    Me.Controls("txtTIME_RECEIVE_HW_QTY").Value = Format(.Range("TIME_RECEIVE_HW_QTY").Value, "0")
    Me.Controls("txtTIME_RECEIVE_HW_EST").Value = Format(.Range("TIME_RECEIVE_HW_EST").Value, "#,#0.0")
    'TIME_PREPARE_HW_QTY
    Me.Controls("txtTIME_PREPARE_HW_QTY").Value = Format(.Range("TIME_PREPARE_HW_QTY").Value, "0")
    Me.Controls("txtTIME_PREPARE_HW_EST").Value = Format(.Range("TIME_PREPARE_HW_EST").Value, "#,#0.0")
    'TIME_TRANSPORT_HW_QTY
    Me.Controls("txtTIME_TRANSPORT_HW_QTY").Value = Format(.Range("TIME_TRANSPORT_HW_QTY").Value, "0")
    Me.Controls("txtTIME_TRANSPORT_HW_EST").Value = Format(.Range("TIME_TRANSPORT_HW_EST").Value, "#,#0.0")
    
End With

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub UpdateTMCTab()
' Added by JFR 15/08/2007
On Error GoTo ErrHandler

'set all but the bound controls on this tab
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    'GasTurbine_QTY
    Me.Controls("txtGasTurbine_QTY").Value = Format(.Range("GasTurbine_QTY").Value, "0")
    'SingleShaft_QTY
    Me.Controls("txtSingleShaft_QTY").Value = Format(.Range("SingleShaft_QTY").Value, "0")
    'MultiShaft_QTY
    Me.Controls("txtMultiShaft_QTY").Value = Format(.Range("MultiShaft_QTY").Value, "0")
    'Aeroderivative_QTY
    Me.Controls("txtAeroderivative_QTY").Value = Format(.Range("Aeroderivative_QTY").Value, "0")
    'SteamTurbine_QTY
    Me.Controls("txtSteamTurbine_QTY").Value = Format(.Range("SteamTurbine_QTY").Value, "0")
    'SingleExtraction_QTY
    Me.Controls("txtSingleExtraction_QTY").Value = Format(.Range("SingleExtraction_QTY").Value, "0")
    'DoubleExtraction_QTY
    Me.Controls("txtDoubleExtraction_QTY").Value = Format(.Range("DoubleExtraction_QTY").Value, "0")
    'Reheat_QTY
    Me.Controls("txtReheat_QTY").Value = Format(.Range("Reheat_QTY").Value, "0")
    'PRVValves_QTY
    Me.Controls("txtPRVValves_QTY").Value = Format(.Range("PRVValves_QTY").Value, "0")
    'MotorDriven_QTY
    Me.Controls("txtMotorDriven_QTY").Value = Format(.Range("MotorDriven_QTY").Value, "0")
    'Compressor_QTY
    Me.Controls("txtCompressor_QTY").Value = Format(.Range("Compressor_QTY").Value, "0")
    'RecycleValves_QTY
    Me.Controls("txtRecycleValves_QTY").Value = Format(.Range("RecycleValves_QTY").Value, "0")
    'Generator_QTY
    Me.Controls("txtGenerator_QTY").Value = Format(.Range("Generator_QTY").Value, "0")

    
End With
Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub UpdateTestingTab()
On Error GoTo ErrHandler
'just the unbound fields
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    'TEST_PRE_FAT_PCT
    Me.Controls("txtTEST_PRE_FAT_PCT").Value = Format(.Range("TEST_PRE_FAT_PCT").Value, "0%")
    'TEST_FAT_PCT
    Me.Controls("txtTEST_FAT_PCT").Value = Format(.Range("TEST_FAT_PCT").Value, "0%")
    'set units for label rental by babs
    Me.Controls("lblTestRentalUnits").Caption = .Range("RENTAL_COST_UNITS").Text

End With
Call UpdateSummaryPage(Me)

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Private Sub SaveAndExit()
'check to see if the saved propert is set and prompt to save or cancel then close the workbook
On Error GoTo ErrHandler
Dim intReturn As Integer
If ThisWorkbook.Saved Then
    'MsgBox "saved"
    ThisWorkbook.Close
Else

    intReturn = MsgBox("Do you want to save before closing?", vbYesNoCancel, "Save")
    Select Case intReturn
    Case vbYes
        'close with save 'vbYes = 6
        ThisWorkbook.Close (True)
    Case vbNo
        'close don't save 'vbNo = 7
        ThisWorkbook.Close (False)
    Case Else
        'cancel 'vbCancel = 2
        
    End Select
    
End If

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Sub ExportToERP()
On Error GoTo ErrHandler

Dim strExportFile As String, strPath As String
Dim rwIndex As Integer, colIndex As Integer
Dim strTemp As String, strPipe As String
Dim fso As Object, tf As Object
Dim strHeader As String, strHeaderData As String

strExportFile = "ExportToERP.csv"
strPipe = "|"

Set fso = CreateObject("Scripting.FileSystemObject")

strHeader = "UserName" & strPipe & "Date" & strPipe & "GECE Version" & strPipe & "Workbook Name" & strPipe & "Excel Version"
strHeaderData = Application.UserName & strPipe & Now() & strPipe & GECEXLSVERSION & strPipe & Application.ActiveWorkbook.Name & strPipe & Application.Version

strPath = Application.GetSaveAsFilename(strExportFile, "CSV Files (*.csv), *.csv", , "Save ERP Export File")
If strPath = "False" Then Exit Sub

Set tf = fso.CreateTextFile(strPath, True)
tf.WriteLine strHeader
tf.WriteLine strHeaderData

For rwIndex = 1 To 92
    strTemp = ""
    For colIndex = 1 To 8
        With Worksheets(gstrGECEExportToERPSheet).Cells(rwIndex, colIndex)
            If colIndex = 8 Then
                strTemp = strTemp & .Value
            ElseIf colIndex = 7 And rwIndex <> 1 Then
                strTemp = strTemp & Round(.Value, 2) & strPipe
            ElseIf (colIndex = 5 Or colIndex = 6) And rwIndex <> 1 Then
                strTemp = strTemp & Round(.Value, 1) & strPipe
            Else
                strTemp = strTemp & .Value & strPipe
            End If
        End With
    Next colIndex
    tf.WriteLine strTemp
Next rwIndex

MsgBox "File created: " & strPath, vbOKOnly, "GECE Export To ERP"

CleanUp:
On Error Resume Next
tf.Close
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description
    Resume CleanUp
End Sub

Sub MergeFile()
On Error GoTo ErrHandler

Dim vntPath As Variant, vntPathObj As Variant
Dim strExportFile As String, strPipe As String
Dim rwIndex As Integer, colIndex As Integer
Dim strTemp As String, strHeader As String, strHeaderData As String
Dim tf As Integer, tf1 As Integer
Dim strPath As String, strPath1 As String
Dim ii As Long, jj As Long, fileNumber As Long
Dim ws As Worksheet
Const sheetName As String = "ExportToERP"

Set ws = Nothing
On Error Resume Next
Set ws = Worksheets(sheetName)
On Error GoTo 0
If ws Is Nothing Then
    frmComplete.Controls("txtOutput").Text = "Sheet not found: " & sheetName & vbCrLf
    Exit Sub
End If

frmComplete.Controls("txtOutput").Text = "Merging function started" & vbCrLf

strPipe = ","
strExportFile = "ExportToERP_Merged.csv"

strHeader = "UserName" & strPipe & "Date" & strPipe & "GECE Version" & strPipe & "Workbook Name" & strPipe & "Excel Version"
strHeaderData = Application.UserName & strPipe & Now() & strPipe & GECEXLSVERSION & strPipe & Application.ActiveWorkbook.Name & strPipe & Application.Version

vntPath = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , "Select files to merge", , True)
If IsArray(vntPath) = False Then
    frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & "Cancelled: No input files" & vbCrLf
    Exit Sub
End If
If UBound(vntPath) < 1 Then
    frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & "Cancelled: At least 2 files required" & vbCrLf
    Exit Sub
End If

fileNumber = 0
ws.Range("K1:R92").ClearContents

For Each vntPathObj In vntPath
    fileNumber = fileNumber + 1
    strPath1 = CStr(vntPathObj)
    tf1 = FreeFile
    Open strPath1 For Input As tf1
    frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & "Reading: " & strPath1 & vbCrLf

    Line Input #tf1, strHeader
    Line Input #tf1, strHeaderData

    For rwIndex = 1 To 92
        strTemp = ""
        jj = 0
        Line Input #tf1, strTemp
        For colIndex = 1 To 8
            If colIndex < 5 And fileNumber <> 1 And rwIndex <> 1 Then
                ii = jj + 1
                jj = InStr(ii, strTemp, strPipe)
            ElseIf fileNumber <> 1 And rwIndex = 1 Then
                Exit For
            ElseIf colIndex = 8 Or (colIndex = 8 And fileNumber <> 1 And rwIndex = 2) Then
                ii = jj + 1
                Dim newCurrency As String: newCurrency = mid(strTemp, ii, Len(strTemp) - ii + 1)
                If ws.Cells(rwIndex, colIndex + 10).Value = "" Then
                    ws.Cells(rwIndex, colIndex + 10).Value = newCurrency
                ElseIf ws.Cells(rwIndex, colIndex + 10).Value <> newCurrency Then
                    frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & _
                        "Currency mismatch at row " & rwIndex & ": " & ws.Cells(rwIndex, colIndex + 10).Value & " ? " & newCurrency & vbCrLf
                    GoTo CleanUp
                End If
            Else
                ii = jj + 1
                jj = InStr(ii, strTemp, strPipe)
                If ws.Cells(rwIndex, colIndex + 10).Value = "" Then
                    ws.Cells(rwIndex, colIndex + 10).Value = mid(strTemp, ii, jj - ii)
                Else
                    ws.Cells(rwIndex, colIndex + 10).Value = CDbl(ws.Cells(rwIndex, colIndex + 10).Value) + CDbl(mid(strTemp, ii, jj - ii))
                End If
            End If
        Next colIndex
    Next rwIndex
    Close tf1
Next vntPathObj

strPath = Application.GetSaveAsFilename(InitialFileName:=strExportFile, FileFilter:="CSV Files (*.csv), *.csv", Title:="Save Merged Export File")
If strPath = "False" Then Exit Sub

tf = FreeFile
Open strPath For Output As tf

frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & "Writing merged file: " & strPath & vbCrLf
Print #tf, strHeader
Print #tf, strHeaderData

For rwIndex = 1 To 92
    strTemp = ""
    For colIndex = 1 To 8
        With ws.Cells(rwIndex, colIndex + 10)
            If colIndex = 8 Then
                strTemp = strTemp & .Value
            ElseIf colIndex = 7 And rwIndex <> 1 Then
                strTemp = strTemp & Round(.Value, 2) & strPipe
            ElseIf (colIndex = 5 Or colIndex = 6) And rwIndex <> 1 Then
                strTemp = strTemp & Round(.Value, 1) & strPipe
            Else
                strTemp = strTemp & .Value & strPipe
            End If
        End With
    Next colIndex
    Print #tf, strTemp
Next rwIndex

frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & "Merge complete." & vbCrLf
MsgBox "Merged file created: " & strPath, vbOKOnly, "GECE Export"

CleanUp:
On Error Resume Next
Close tf
Exit Sub

ErrHandler:
MsgBox Err.Number & " - " & Err.Description, vbCritical, "MergeFile Error"
Resume CleanUp

End Sub

Sub ScopeChange()
On Error GoTo ErrHandler

Dim vntPath As Variant, vntPathObj As Variant
Dim strPath1 As String
Dim rwIndex As Integer, colIndex As Integer
Dim strTemp As String, strPipe As String
Dim tf1 As Integer
Dim strHeader As String, strHeaderData As String
Dim ii As Long, jj As Long, fileNumber As Long
Dim ws As Worksheet
Dim NewFile As Long

Set ws = Worksheets(gstrGECEExportToERPSheet)
strPipe = "|"
NewFile = 0
frmComplete.Controls("txtOutput").Text = "Scope change function started" & vbCrLf

vntPath = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , , , True)
If IsArray(vntPath) = False Then
    frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & "Cancelled by the user: No input files" & vbCrLf
    Exit Sub
End If

If UBound(vntPath) + 1 <> 2 Then
    MsgBox "Select 2 files.", vbExclamation
    frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & "Cancelled: Need 2 files" & vbCrLf
    Exit Sub
End If

fileNumber = 0
ws.Range("K1:R92").ClearContents

For Each vntPathObj In vntPath
    fileNumber = fileNumber + 1
    strPath1 = CStr(vntPathObj)
    tf1 = FreeFile
    Open strPath1 For Input As tf1
    frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & "Reading file: " & strPath1 & vbCrLf

    Line Input #tf1, strHeader
    Line Input #tf1, strHeaderData

    For rwIndex = 1 To 92
        strTemp = ""
        jj = 0
        Line Input #tf1, strTemp
        For colIndex = 1 To 8
            With ws
                If colIndex < 5 And fileNumber <> 1 And rwIndex <> 1 Then
                    ii = jj + 1
                    jj = InStr(ii, strTemp, strPipe)
                ElseIf fileNumber <> 1 And rwIndex = 1 Then
                    Exit For
                ElseIf colIndex = 8 Or (colIndex = 8 And fileNumber <> 1 And rwIndex = 2) Then
                    ii = jj + 1
                    If .Cells(rwIndex, colIndex + 10).Value = "" Then
                        .Cells(rwIndex, colIndex + 10).Value = mid(strTemp, ii, Len(strTemp) - ii + 1)
                    Else
                        If .Cells(rwIndex, colIndex + 10).Value <> mid(strTemp, ii, Len(strTemp) - ii + 1) Then
                            MsgBox "You need to have the same currency across your file.  " & .Cells(rwIndex, colIndex + 10).Value & " ? " & mid(strTemp, ii, Len(strTemp) - ii + 1)
                            frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & "Currency mismatch at row " & rwIndex & vbCrLf
                            Close tf1
                            Exit Sub
                        End If
                    End If
                Else
                    ii = jj + 1
                    jj = InStr(ii, strTemp, strPipe)
                    If .Cells(rwIndex, colIndex + 10).Value = "" Then
                        .Cells(rwIndex, colIndex + 10).Value = mid(strTemp, ii, jj - ii)
                    Else
                        If NewFile = 1 Then
                            .Cells(rwIndex, colIndex + 10).Value = CDbl(.Cells(rwIndex, colIndex + 10).Value) - CDbl(mid(strTemp, ii, jj - ii))
                        ElseIf NewFile = 2 Then
                            .Cells(rwIndex, colIndex + 10).Value = CDbl(mid(strTemp, ii, jj - ii)) - CDbl(.Cells(rwIndex, colIndex + 10).Value)
                        Else
                            If CDbl(.Cells(rwIndex, colIndex + 10).Value) > CDbl(mid(strTemp, ii, jj - ii)) Then
                                NewFile = 1
                                .Cells(rwIndex, colIndex + 10).Value = CDbl(.Cells(rwIndex, colIndex + 10).Value) - CDbl(mid(strTemp, ii, jj - ii))
                            ElseIf CDbl(.Cells(rwIndex, colIndex + 10).Value) < CDbl(mid(strTemp, ii, jj - ii)) Then
                                NewFile = 2
                                .Cells(rwIndex, colIndex + 10).Value = CDbl(mid(strTemp, ii, jj - ii)) - CDbl(.Cells(rwIndex, colIndex + 10).Value)
                            Else
                                .Cells(rwIndex, colIndex + 10).Value = 0
                            End If
                        End If
                    End If
                End If
            End With
        Next colIndex
    Next rwIndex
    Close tf1
Next vntPathObj

Dim savePath As String, fileNameOut As String
fileNameOut = "ExportToERP_scopechange_" & Format(Now, "yyyymmdd_HHMMSS") & ".xlsx"
savePath = Application.GetSaveAsFilename(fileNameOut, "Excel Files (*.xlsx), *.xlsx", , "Save Scope Change File")

If savePath = "False" Then
    frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & "User cancelled scope change save." & vbCrLf
    Exit Sub
End If

Dim wbOut As Workbook
Set wbOut = Workbooks.Add
Dim wsOut As Worksheet
Set wsOut = wbOut.Worksheets(1)

wsOut.Cells(1, 1).Value = "Brand"
wsOut.Cells(1, 2).Value = "Tracking Activity Id"
wsOut.Cells(1, 3).Value = "Tracking Activity"
wsOut.Cells(1, 4).Value = "Activity Type"
wsOut.Cells(1, 5).Value = "Local Hours ?"
wsOut.Cells(1, 6).Value = "Remote Hours ?"
wsOut.Cells(1, 7).Value = "Cost ?"
wsOut.Cells(1, 8).Value = "Currency"

ws.Range("K1:R92").Copy Destination:=wsOut.Range("A2")

Application.DisplayAlerts = False
wbOut.SaveAs filename:=savePath, FileFormat:=xlOpenXMLWorkbook
Application.DisplayAlerts = True
wbOut.Close SaveChanges:=False

frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & "Scope change file saved to: " & savePath & vbCrLf
MsgBox "Scope Change saved: " & savePath, vbOKOnly, "GECE Scope Change"

CleanUp:
On Error Resume Next
Close tf1
Exit Sub

ErrHandler:
MsgBox Err.Number & ": " & Err.Description
Resume CleanUp
End Sub


Sub SetWPA(oForm As UserForm)
    If Workbooks(gstrGECEWorkBook).Worksheets(gstrGECECostSheet).Range("LOC_REGION") = "NA" Then
        oForm.Controls("cboWPA").RowSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEWPASheet).Range("WPA_NA").Address(External:=True)
        oForm.Controls("cboWPAType").RowSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEWPASheet).Range("WPA_TYPE_LIST").Address(External:=True)
    ElseIf Workbooks(gstrGECEWorkBook).Worksheets(gstrGECECostSheet).Range("LOC_REGION") = "LAO" Then
        oForm.Controls("cboWPA").RowSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEWPASheet).Range("WPA_LAO").Address(External:=True)
        oForm.Controls("cboWPAType").RowSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEWPASheet).Range("WPA_TYPE_LIST").Address(External:=True)
    ElseIf Workbooks(gstrGECEWorkBook).Worksheets(gstrGECECostSheet).Range("LOC_REGION") = "EURA" Then
        oForm.Controls("cboWPA").RowSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEWPASheet).Range("WPA_EURA").Address(External:=True)
        oForm.Controls("cboWPAType").RowSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEWPASheet).Range("WPA_TYPE_EURA").Address(External:=True)
    ElseIf Workbooks(gstrGECEWorkBook).Worksheets(gstrGECECostSheet).Range("LOC_REGION") = "MENA" Then
        oForm.Controls("cboWPA").RowSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEWPASheet).Range("WPA_MENA").Address(External:=True)
        oForm.Controls("cboWPAType").RowSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEWPASheet).Range("WPA_TYPE_LIST").Address(External:=True)
    ElseIf Workbooks(gstrGECEWorkBook).Worksheets(gstrGECECostSheet).Range("LOC_REGION") = "APAC" Then
        oForm.Controls("cboWPA").RowSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEWPASheet).Range("WPA_APAC").Address(External:=True)
        oForm.Controls("cboWPAType").RowSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEWPASheet).Range("WPA_TYPE_LIST").Address(External:=True)
    Else
        oForm.Controls("cboWPA").RowSource = "No WPA"
    End If
    oForm.Controls("cboWPA").ControlSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("WPA").Address(External:=True)
    oForm.Controls("cboWPAType").ControlSource = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("WPA_TYPE").Address(External:=True)
End Sub
Public Sub ShowLoadingThenContinue()
    On Error GoTo ErrHandler

    Me.lblLoading.Visible = True
    DoEvents
    Me.Repaint

CleanExit:
    Me.lblLoading.Visible = False
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Load Error"
    Resume CleanExit
End Sub
Function PercentText(v As Variant) As String
    PercentText = Format(PercentAsDouble(v), "0%")
    
End Function
