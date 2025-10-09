Attribute VB_Name = "modImport"
Option Explicit
'version 1.0


'get the worksheet version we are importing
Dim strCoverSheetVersion As String
Public gbolImporting As Boolean


Public Sub ImportWorkbookFull(strFileToOpen As String)
On Error Resume Next
Dim strDataEntrySheet As String
Dim i As Integer
Dim wbSource As Workbook
Dim strCountry As String
Dim strActiveWorkbook As String
strActiveWorkbook = ActiveWorkbook.Name

'set flag for importing so default remote country function will not update
gbolImporting = True


Dim strP As String
strP = Chr(116) & Chr(68) & Chr(107) & Chr(49) & Chr(52) & Chr(52) & Chr(77) & Chr(98)

Set wbSource = Workbooks.Open(strFileToOpen, , True, , strP)


strCoverSheetVersion = wbSource.Worksheets("CoverSheet").Range("GECE_XLS_VERSION").Value

Application.Cursor = xlWait

'2006-02-17 updated 'DataEntry
frmComplete.Controls("txtOutput").Text = "Source Version: " & strCoverSheetVersion & vbCrLf & "Importing from: " & gstrGECEDataEntrySheet
DoEvents
Call ImportDataEntrySheet(wbSource.Worksheets(gstrGECEDataEntrySheet), Workbooks(strActiveWorkbook).Worksheets(gstrGECEDataEntrySheet))
DoEvents
Application.Cursor = xlWait

'2006-02-17 updated 'AssumptionsProposal
frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & vbCrLf & "Importing from: " & gstrGECEAssumptionsProposalSheet
DoEvents
Call ImportAssumptionsProposalSheet(wbSource.Worksheets(gstrGECEAssumptionsProposalSheet), Workbooks(strActiveWorkbook).Worksheets(gstrGECEAssumptionsProposalSheet))
DoEvents
Application.Cursor = xlWait


'ImportApplicationBasedSheet
'2006-02-17 updated '
frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & vbCrLf & "Importing from: " & gstrGECEApplicationBasedSheet
DoEvents

    For i = 1 To wbSource.Worksheets.Count
        Select Case wbSource.Worksheets(i).Name
        Case "Summary"
            'if  its an old workbook import from summary sheet
            If val(strCoverSheetVersion) >= val("1.0") Then
                Call ImportApplicationBasedSheet_428(wbSource.Worksheets("Summary"), Workbooks(strActiveWorkbook).Worksheets(gstrGECEApplicationBasedSheet))
            Else
                Call ImportApplicationBasedSheet(wbSource.Worksheets("Summary"), Workbooks(strActiveWorkbook).Worksheets(gstrGECEApplicationBasedSheet))
            End If
            Exit For
            DoEvents
            Application.Cursor = xlWait
        Case "Application Based"
            
            'Call ImportSheet(gstrGECEApplicationBasedSheet, "A1:U320", wbSource.Worksheets(gstrGECEApplicationBasedSheet), Workbooks(strActiveWorkbook).Worksheets(gstrGECEApplicationBasedSheet))
            If val(strCoverSheetVersion) >= val("1.0") Then
                Call ImportApplicationBasedSheet_428(wbSource.Worksheets(gstrGECEApplicationBasedSheet), Workbooks(strActiveWorkbook).Worksheets(gstrGECEApplicationBasedSheet))
            Else
                Call ImportApplicationBasedSheet(wbSource.Worksheets(gstrGECEApplicationBasedSheet), Workbooks(strActiveWorkbook).Worksheets(gstrGECEApplicationBasedSheet))
            End If
            Exit For
            DoEvents
            Application.Cursor = xlWait
        Case Else
            'do nothing
        End Select

    Next
    
'ImportPriceMakeupSheet
frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & vbCrLf & "Importing from: " & gstrGECEPriceMakeUpSheet
DoEvents
If val(strCoverSheetVersion) >= val("1.0") Then
    Call ImportPriceMakeupSheet_428(wbSource.Worksheets(gstrGECEPriceMakeUpSheet), Workbooks(strActiveWorkbook).Worksheets(gstrGECEPriceMakeUpSheet))
Else
    Call ImportPriceMakeupSheet(wbSource.Worksheets(gstrGECEPriceMakeUpSheet), Workbooks(strActiveWorkbook).Worksheets(gstrGECEPriceMakeUpSheet))
End If
DoEvents


'ONLY import duration and task sheets if version is greater or equal to 1.0
If val(strCoverSheetVersion) >= val("1.0") Then
    'ImportDurationBasedSheet
    
    frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & vbCrLf & "Importing from: " & gstrGECEDurationBasedSheet
    DoEvents
    Call ImportDurationBasedSheet(wbSource.Worksheets(gstrGECEDurationBasedSheet), Workbooks(strActiveWorkbook).Worksheets(gstrGECEDurationBasedSheet))
    DoEvents
    Application.Cursor = xlWait
    ''These will loop through cells and copy them
    ''could break if cells shift
    
    '' USE ABOVE HARD CODED FUNCTION''
  
    'frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & vbCrLf & "Importing from: " & gstrGECEDurationBasedSheet
    'DoEvents
    'Call ImportSheet(gstrGECEDurationBasedSheet, "A1:O108", wbSource.Worksheets(gstrGECEDurationBasedSheet), Workbooks(strActiveWorkbook).Worksheets(gstrGECEDurationBasedSheet))
    'DoEvents
    
    
    'ImportTaskBasedSheet
    frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & vbCrLf & "Importing from: " & gstrGECETaskBasedSheet
    DoEvents
    Call ImportSheet(gstrGECETaskBasedSheet, "A1:BR386", wbSource.Worksheets(gstrGECETaskBasedSheet), Workbooks(strActiveWorkbook).Worksheets(gstrGECETaskBasedSheet))
    DoEvents
    Application.Cursor = xlWait
Else
frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & vbCrLf & "Skipping - pre 1.0 Workbook: " & gstrGECEDurationBasedSheet
frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & vbCrLf & "Skipping - pre 1.0 Workbook: " & gstrGECETaskBasedSheet

End If

Application.Cursor = xlDefault
frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & vbCrLf & "Finished."
DoEvents
MsgBox "Finished importing data from " & strFileToOpen
wbSource.Close (False)

Set wbSource = Nothing

gbolImporting = False
End Sub

Private Sub ImportPriceMakeupSheet_428(wsSource As Worksheet, wsTarget As Worksheet)
On Error Resume Next
'GECE Version = 1.0

With wsTarget
    .Range("SPEC_TASK1_OVD_STS").Value = wsSource.Range("SPEC_TASK1_OVD_STS").Value ' $E$116
    .Range("SPEC_TASK2_OVD_STS").Value = wsSource.Range("SPEC_TASK2_OVD_STS").Value ' $E$117
    .Range("SPEC_TASK3_OVD_STS").Value = wsSource.Range("SPEC_TASK3_OVD_STS").Value ' $E$118
    .Range("SPEC_TASK4_OVD_STS").Value = wsSource.Range("SPEC_TASK4_OVD_STS").Value ' $E$119
    .Range("SPEC_TASK5_OVD_STS").Value = wsSource.Range("SPEC_TASK5_OVD_STS").Value ' $E$120
    .Range("SPEC_TASK6_OVD_STS").Value = wsSource.Range("SPEC_TASK6_OVD_STS").Value ' $E$121
    .Range("SPEC_TASK7_OVD_STS").Value = wsSource.Range("SPEC_TASK7_OVD_STS").Value ' $E$122
    .Range("SPEC_TASK8_OVD_STS").Value = wsSource.Range("SPEC_TASK8_OVD_STS").Value ' $E$123
    .Range("SPEC_TASK9_OVD_STS").Value = wsSource.Range("SPEC_TASK9_OVD_STS").Value ' $E$124
    .Range("SPEC_TASK10_OVD_STS").Value = wsSource.Range("SPEC_TASK10_OVD_STS").Value ' $E$125
    
    .Range("SYSENG_TASK1_OVD_STS").Value = wsSource.Range("SYSENG_TASK1_OVD_STS").Value ' $E$116
    .Range("SYSENG_TASK2_OVD_STS").Value = wsSource.Range("SYSENG_TASK2_OVD_STS").Value ' $E$117
    .Range("SYSENG_TASK3_OVD_STS").Value = wsSource.Range("SYSENG_TASK3_OVD_STS").Value ' $E$118
    .Range("SYSENG_TASK4_OVD_STS").Value = wsSource.Range("SYSENG_TASK4_OVD_STS").Value ' $E$119
    .Range("SYSENG_TASK5_OVD_STS").Value = wsSource.Range("SYSENG_TASK5_OVD_STS").Value ' $E$120
    .Range("SYSENG_TASK6_OVD_STS").Value = wsSource.Range("SYSENG_TASK6_OVD_STS").Value ' $E$121
    .Range("SYSENG_TASK7_OVD_STS").Value = wsSource.Range("SYSENG_TASK7_OVD_STS").Value ' $E$122
    .Range("SYSENG_TASK8_OVD_STS").Value = wsSource.Range("SYSENG_TASK8_OVD_STS").Value ' $E$123
    .Range("SYSENG_TASK9_OVD_STS").Value = wsSource.Range("SYSENG_TASK9_OVD_STS").Value ' $E$124
    .Range("SYSENG_TASK10_OVD_STS").Value = wsSource.Range("SYSENG_TASK10_OVD_STS").Value ' $E$125
    
    .Range("HMI_TASK1_OVD_STS").Value = wsSource.Range("HMI_TASK1_OVD_STS").Value ' $E$116
    .Range("HMI_TASK2_OVD_STS").Value = wsSource.Range("HMI_TASK2_OVD_STS").Value ' $E$117
    .Range("HMI_TASK3_OVD_STS").Value = wsSource.Range("HMI_TASK3_OVD_STS").Value ' $E$118
    .Range("HMI_TASK4_OVD_STS").Value = wsSource.Range("HMI_TASK4_OVD_STS").Value ' $E$119
    .Range("HMI_TASK5_OVD_STS").Value = wsSource.Range("HMI_TASK5_OVD_STS").Value ' $E$120
    .Range("HMI_TASK6_OVD_STS").Value = wsSource.Range("HMI_TASK6_OVD_STS").Value ' $E$121
    .Range("HMI_TASK7_OVD_STS").Value = wsSource.Range("HMI_TASK7_OVD_STS").Value ' $E$122
    .Range("HMI_TASK8_OVD_STS").Value = wsSource.Range("HMI_TASK8_OVD_STS").Value ' $E$123
    .Range("HMI_TASK9_OVD_STS").Value = wsSource.Range("HMI_TASK9_OVD_STS").Value ' $E$124
    .Range("HMI_TASK10_OVD_STS").Value = wsSource.Range("HMI_TASK10_OVD_STS").Value ' $E$125
    
    .Range("CP_TASK1_OVD_STS").Value = wsSource.Range("CP_TASK1_OVD_STS").Value ' $E$116
    .Range("CP_TASK2_OVD_STS").Value = wsSource.Range("CP_TASK2_OVD_STS").Value ' $E$117
    .Range("CP_TASK3_OVD_STS").Value = wsSource.Range("CP_TASK3_OVD_STS").Value ' $E$118
    .Range("CP_TASK4_OVD_STS").Value = wsSource.Range("CP_TASK4_OVD_STS").Value ' $E$119
    .Range("CP_TASK5_OVD_STS").Value = wsSource.Range("CP_TASK5_OVD_STS").Value ' $E$120
    .Range("CP_TASK6_OVD_STS").Value = wsSource.Range("CP_TASK6_OVD_STS").Value ' $E$121
    .Range("CP_TASK7_OVD_STS").Value = wsSource.Range("CP_TASK7_OVD_STS").Value ' $E$122
    .Range("CP_TASK8_OVD_STS").Value = wsSource.Range("CP_TASK8_OVD_STS").Value ' $E$123
    .Range("CP_TASK9_OVD_STS").Value = wsSource.Range("CP_TASK9_OVD_STS").Value ' $E$124
    .Range("CP_TASK10_OVD_STS").Value = wsSource.Range("CP_TASK10_OVD_STS").Value ' $E$125
    
    .Range("DI_TASK1_OVD_STS").Value = wsSource.Range("DI_TASK1_OVD_STS").Value ' $E$116
    .Range("DI_TASK2_OVD_STS").Value = wsSource.Range("DI_TASK2_OVD_STS").Value ' $E$117
    .Range("DI_TASK3_OVD_STS").Value = wsSource.Range("DI_TASK3_OVD_STS").Value ' $E$118
    .Range("DI_TASK4_OVD_STS").Value = wsSource.Range("DI_TASK4_OVD_STS").Value ' $E$119
    .Range("DI_TASK5_OVD_STS").Value = wsSource.Range("DI_TASK5_OVD_STS").Value ' $E$120
    .Range("DI_TASK6_OVD_STS").Value = wsSource.Range("DI_TASK6_OVD_STS").Value ' $E$121
    .Range("DI_TASK7_OVD_STS").Value = wsSource.Range("DI_TASK7_OVD_STS").Value ' $E$122
    .Range("DI_TASK8_OVD_STS").Value = wsSource.Range("DI_TASK8_OVD_STS").Value ' $E$123
    .Range("DI_TASK9_OVD_STS").Value = wsSource.Range("DI_TASK9_OVD_STS").Value ' $E$124
    .Range("DI_TASK10_OVD_STS").Value = wsSource.Range("DI_TASK10_OVD_STS").Value ' $E$125
    
    .Range("ESD_TASK1_OVD_STS").Value = wsSource.Range("ESD_TASK1_OVD_STS").Value ' $E$116
    .Range("ESD_TASK2_OVD_STS").Value = wsSource.Range("ESD_TASK2_OVD_STS").Value ' $E$117
    .Range("ESD_TASK3_OVD_STS").Value = wsSource.Range("ESD_TASK3_OVD_STS").Value ' $E$118
    .Range("ESD_TASK4_OVD_STS").Value = wsSource.Range("ESD_TASK4_OVD_STS").Value ' $E$119
    .Range("ESD_TASK5_OVD_STS").Value = wsSource.Range("ESD_TASK5_OVD_STS").Value ' $E$120
    .Range("ESD_TASK6_OVD_STS").Value = wsSource.Range("ESD_TASK6_OVD_STS").Value ' $E$121
    .Range("ESD_TASK7_OVD_STS").Value = wsSource.Range("ESD_TASK7_OVD_STS").Value ' $E$122
    .Range("ESD_TASK8_OVD_STS").Value = wsSource.Range("ESD_TASK8_OVD_STS").Value ' $E$123
    .Range("ESD_TASK9_OVD_STS").Value = wsSource.Range("ESD_TASK9_OVD_STS").Value ' $E$124
    .Range("ESD_TASK10_OVD_STS").Value = wsSource.Range("ESD_TASK10_OVD_STS").Value ' $E$125
    
    .Range("REP_TASK1_OVD_STS").Value = wsSource.Range("REP_TASK1_OVD_STS").Value ' $E$116
    .Range("REP_TASK2_OVD_STS").Value = wsSource.Range("REP_TASK2_OVD_STS").Value ' $E$117
    .Range("REP_TASK3_OVD_STS").Value = wsSource.Range("REP_TASK3_OVD_STS").Value ' $E$118
    .Range("REP_TASK4_OVD_STS").Value = wsSource.Range("REP_TASK4_OVD_STS").Value ' $E$119
    .Range("REP_TASK5_OVD_STS").Value = wsSource.Range("REP_TASK5_OVD_STS").Value ' $E$120
    .Range("REP_TASK6_OVD_STS").Value = wsSource.Range("REP_TASK6_OVD_STS").Value ' $E$121
    .Range("REP_TASK7_OVD_STS").Value = wsSource.Range("REP_TASK7_OVD_STS").Value ' $E$122
    .Range("REP_TASK8_OVD_STS").Value = wsSource.Range("REP_TASK8_OVD_STS").Value ' $E$123
    .Range("REP_TASK9_OVD_STS").Value = wsSource.Range("REP_TASK9_OVD_STS").Value ' $E$124
    .Range("REP_TASK10_OVD_STS").Value = wsSource.Range("REP_TASK10_OVD_STS").Value ' $E$125
    
    .Range("APP_TASK1_OVD_STS").Value = wsSource.Range("APP_TASK1_OVD_STS").Value ' $E$116
    .Range("APP_TASK2_OVD_STS").Value = wsSource.Range("APP_TASK2_OVD_STS").Value ' $E$117
    .Range("APP_TASK3_OVD_STS").Value = wsSource.Range("APP_TASK3_OVD_STS").Value ' $E$118
    .Range("APP_TASK4_OVD_STS").Value = wsSource.Range("APP_TASK4_OVD_STS").Value ' $E$119
    .Range("APP_TASK5_OVD_STS").Value = wsSource.Range("APP_TASK5_OVD_STS").Value ' $E$120
    .Range("APP_TASK6_OVD_STS").Value = wsSource.Range("APP_TASK6_OVD_STS").Value ' $E$121
    .Range("APP_TASK7_OVD_STS").Value = wsSource.Range("APP_TASK7_OVD_STS").Value ' $E$122
    .Range("APP_TASK8_OVD_STS").Value = wsSource.Range("APP_TASK8_OVD_STS").Value ' $E$123
    .Range("APP_TASK9_OVD_STS").Value = wsSource.Range("APP_TASK9_OVD_STS").Value ' $E$124
    .Range("APP_TASK10_OVD_STS").Value = wsSource.Range("APP_TASK10_OVD_STS").Value ' $E$125
    
    .Range("TEST_TASK1_OVD_STS").Value = wsSource.Range("TEST_TASK1_OVD_STS").Value ' $E$116
    .Range("TEST_TASK2_OVD_STS").Value = wsSource.Range("TEST_TASK2_OVD_STS").Value ' $E$117
    .Range("TEST_TASK3_OVD_STS").Value = wsSource.Range("TEST_TASK3_OVD_STS").Value ' $E$118
    .Range("TEST_TASK4_OVD_STS").Value = wsSource.Range("TEST_TASK4_OVD_STS").Value ' $E$119
    .Range("TEST_TASK5_OVD_STS").Value = wsSource.Range("TEST_TASK5_OVD_STS").Value ' $E$120
    .Range("TEST_TASK6_OVD_STS").Value = wsSource.Range("TEST_TASK6_OVD_STS").Value ' $E$121
    .Range("TEST_TASK7_OVD_STS").Value = wsSource.Range("TEST_TASK7_OVD_STS").Value ' $E$122
    .Range("TEST_TASK8_OVD_STS").Value = wsSource.Range("TEST_TASK8_OVD_STS").Value ' $E$123
    .Range("TEST_TASK9_OVD_STS").Value = wsSource.Range("TEST_TASK9_OVD_STS").Value ' $E$124
    .Range("TEST_TASK10_OVD_STS").Value = wsSource.Range("TEST_TASK10_OVD_STS").Value ' $E$125
    
    .Range("DOC_TASK1_OVD_STS").Value = wsSource.Range("DOC_TASK1_OVD_STS").Value ' $E$116
    .Range("DOC_TASK2_OVD_STS").Value = wsSource.Range("DOC_TASK2_OVD_STS").Value ' $E$117
    .Range("DOC_TASK3_OVD_STS").Value = wsSource.Range("DOC_TASK3_OVD_STS").Value ' $E$118
    .Range("DOC_TASK4_OVD_STS").Value = wsSource.Range("DOC_TASK4_OVD_STS").Value ' $E$119
    .Range("DOC_TASK5_OVD_STS").Value = wsSource.Range("DOC_TASK5_OVD_STS").Value ' $E$120
    .Range("DOC_TASK6_OVD_STS").Value = wsSource.Range("DOC_TASK6_OVD_STS").Value ' $E$121
    .Range("DOC_TASK7_OVD_STS").Value = wsSource.Range("DOC_TASK7_OVD_STS").Value ' $E$122
    .Range("DOC_TASK8_OVD_STS").Value = wsSource.Range("DOC_TASK8_OVD_STS").Value ' $E$123
    .Range("DOC_TASK9_OVD_STS").Value = wsSource.Range("DOC_TASK9_OVD_STS").Value ' $E$124
    .Range("DOC_TASK10_OVD_STS").Value = wsSource.Range("DOC_TASK10_OVD_STS").Value ' $E$125
    
    .Range("COURSE_TASK1_OVD_STS").Value = wsSource.Range("COURSE_TASK1_OVD_STS").Value ' $E$116
    .Range("COURSE_TASK2_OVD_STS").Value = wsSource.Range("COURSE_TASK2_OVD_STS").Value ' $E$117
    .Range("COURSE_TASK3_OVD_STS").Value = wsSource.Range("COURSE_TASK3_OVD_STS").Value ' $E$118
    .Range("COURSE_TASK4_OVD_STS").Value = wsSource.Range("COURSE_TASK4_OVD_STS").Value ' $E$119
    .Range("COURSE_TASK5_OVD_STS").Value = wsSource.Range("COURSE_TASK5_OVD_STS").Value ' $E$120
    .Range("COURSE_TASK6_OVD_STS").Value = wsSource.Range("COURSE_TASK6_OVD_STS").Value ' $E$121
    .Range("COURSE_TASK7_OVD_STS").Value = wsSource.Range("COURSE_TASK7_OVD_STS").Value ' $E$122
    .Range("COURSE_TASK8_OVD_STS").Value = wsSource.Range("COURSE_TASK8_OVD_STS").Value ' $E$123
    .Range("COURSE_TASK9_OVD_STS").Value = wsSource.Range("COURSE_TASK9_OVD_STS").Value ' $E$124
    .Range("COURSE_TASK10_OVD_STS").Value = wsSource.Range("COURSE_TASK10_OVD_STS").Value ' $E$125
    
    .Range("PM_TASK1_OVD_STS").Value = wsSource.Range("PM_TASK1_OVD_STS").Value ' $E$116
    .Range("PM_TASK2_OVD_STS").Value = wsSource.Range("PM_TASK2_OVD_STS").Value ' $E$117
    .Range("PM_TASK3_OVD_STS").Value = wsSource.Range("PM_TASK3_OVD_STS").Value ' $E$118
    .Range("PM_TASK4_OVD_STS").Value = wsSource.Range("PM_TASK4_OVD_STS").Value ' $E$119
    .Range("PM_TASK5_OVD_STS").Value = wsSource.Range("PM_TASK5_OVD_STS").Value ' $E$120
    .Range("PM_TASK6_OVD_STS").Value = wsSource.Range("PM_TASK6_OVD_STS").Value ' $E$121
    .Range("PM_TASK7_OVD_STS").Value = wsSource.Range("PM_TASK7_OVD_STS").Value ' $E$122
    .Range("PM_TASK8_OVD_STS").Value = wsSource.Range("PM_TASK8_OVD_STS").Value ' $E$123
    .Range("PM_TASK9_OVD_STS").Value = wsSource.Range("PM_TASK9_OVD_STS").Value ' $E$124
    .Range("PM_TASK10_OVD_STS").Value = wsSource.Range("PM_TASK10_OVD_STS").Value ' $E$125
    
    .Range("MEETING_TASK1_OVD_STS").Value = wsSource.Range("MEETING_TASK1_OVD_STS").Value ' $E$116
    .Range("MEETING_TASK2_OVD_STS").Value = wsSource.Range("MEETING_TASK2_OVD_STS").Value ' $E$117
    .Range("MEETING_TASK3_OVD_STS").Value = wsSource.Range("MEETING_TASK3_OVD_STS").Value ' $E$118
    .Range("MEETING_TASK4_OVD_STS").Value = wsSource.Range("MEETING_TASK4_OVD_STS").Value ' $E$119
    .Range("MEETING_TASK5_OVD_STS").Value = wsSource.Range("MEETING_TASK5_OVD_STS").Value ' $E$120
    .Range("MEETING_TASK6_OVD_STS").Value = wsSource.Range("MEETING_TASK6_OVD_STS").Value ' $E$121
    .Range("MEETING_TASK7_OVD_STS").Value = wsSource.Range("MEETING_TASK7_OVD_STS").Value ' $E$122
    .Range("MEETING_TASK8_OVD_STS").Value = wsSource.Range("MEETING_TASK8_OVD_STS").Value ' $E$123
    .Range("MEETING_TASK9_OVD_STS").Value = wsSource.Range("MEETING_TASK9_OVD_STS").Value ' $E$124
    .Range("MEETING_TASK10_OVD_STS").Value = wsSource.Range("MEETING_TASK10_OVD_STS").Value ' $E$125
    
    .Range("SITE_TASK1_OVD_STS").Value = wsSource.Range("SITE_TASK1_OVD_STS").Value ' $E$116
    .Range("SITE_TASK2_OVD_STS").Value = wsSource.Range("SITE_TASK2_OVD_STS").Value ' $E$117
    .Range("SITE_TASK3_OVD_STS").Value = wsSource.Range("SITE_TASK3_OVD_STS").Value ' $E$118
    .Range("SITE_TASK4_OVD_STS").Value = wsSource.Range("SITE_TASK4_OVD_STS").Value ' $E$119
    .Range("SITE_TASK5_OVD_STS").Value = wsSource.Range("SITE_TASK5_OVD_STS").Value ' $E$120
    .Range("SITE_TASK6_OVD_STS").Value = wsSource.Range("SITE_TASK6_OVD_STS").Value ' $E$121
    .Range("SITE_TASK7_OVD_STS").Value = wsSource.Range("SITE_TASK7_OVD_STS").Value ' $E$122
    .Range("SITE_TASK8_OVD_STS").Value = wsSource.Range("SITE_TASK8_OVD_STS").Value ' $E$123
    .Range("SITE_TASK9_OVD_STS").Value = wsSource.Range("SITE_TASK9_OVD_STS").Value ' $E$124
    .Range("SITE_TASK10_OVD_STS").Value = wsSource.Range("SITE_TASK10_OVD_STS").Value ' $E$125
    
    .Range("TL_TASK1_OVD_STS").Value = wsSource.Range("TL_TASK1_OVD_STS").Value ' $E$116
    .Range("TL_TASK2_OVD_STS").Value = wsSource.Range("TL_TASK2_OVD_STS").Value ' $E$117
    .Range("TL_TASK3_OVD_STS").Value = wsSource.Range("TL_TASK3_OVD_STS").Value ' $E$118
    .Range("TL_TASK4_OVD_STS").Value = wsSource.Range("TL_TASK4_OVD_STS").Value ' $E$119
    .Range("TL_TASK5_OVD_STS").Value = wsSource.Range("TL_TASK5_OVD_STS").Value ' $E$120
    .Range("TL_TASK6_OVD_STS").Value = wsSource.Range("TL_TASK6_OVD_STS").Value ' $E$121
    .Range("TL_TASK7_OVD_STS").Value = wsSource.Range("TL_TASK7_OVD_STS").Value ' $E$122
    .Range("TL_TASK8_OVD_STS").Value = wsSource.Range("TL_TASK8_OVD_STS").Value ' $E$123
    .Range("TL_TASK9_OVD_STS").Value = wsSource.Range("TL_TASK9_OVD_STS").Value ' $E$124
    .Range("TL_TASK10_OVD_STS").Value = wsSource.Range("TL_TASK10_OVD_STS").Value ' $E$125
End With


End Sub

Private Sub ImportPriceMakeupSheet(wsSource As Worksheet, wsTarget As Worksheet)
On Error Resume Next
'GECE Version = 1.0

With wsTarget
    .Range("SPEC_TASK1_OVD_STS").Value = wsSource.Range("SPEC_PFS_OVD_STS").Value ' GREEN $E$4    SPEC_PFS_OVD_STS
    .Range("SPEC_TASK2_OVD_STS").Value = wsSource.Range("SPEC_FAT_OVD_STS").Value ' GREEN $E$5    SPEC_FAT_OVD_STS
    .Range("SPEC_TASK3_OVD_STS").Value = wsSource.Range("SPEC_SAT_OVD_STS").Value ' GREEN $E$6    SPEC_SAT_OVD_STS
    .Range("SPEC_TASK4_OVD_STS").Value = wsSource.Range("SPEC_QA_OVD_STS").Value ' GREEN $E$7  SPEC_QA_OVD_STS
    .Range("SYSENG_TASK1_OVD_STS").Value = wsSource.Range("SYSENG_PANEL_OVD_STS").Value ' GREEN $E$14   SYSENG_PANEL_OVD_STS
    .Range("SYSENG_TASK2_OVD_STS").Value = wsSource.Range("SYSENG_SYS_CFG_OVD_STS").Value ' GREEN $E$15   SYSENG_SYS_CFG_OVD_STS
    .Range("HMI_TASK1_OVD_STS").Value = wsSource.Range("HMI_ELEMENTS_OVD_STS").Value ' GREEN $E$22   HMI_ELEMENTS_OVD_STS
    .Range("HMI_TASK2_OVD_STS").Value = wsSource.Range("HMI_OVR_OVD_STS").Value ' GREEN $E$23 HMI_OVR_OVD_STS
    .Range("HMI_TASK3_OVD_STS").Value = wsSource.Range("HMI_PROC_OVD_STS").Value ' GREEN $E$24   HMI_PROC_OVD_STS
    .Range("HMI_TASK4_OVD_STS").Value = wsSource.Range("HMI_RPT_OVD_STS").Value ' GREEN $E$25 HMI_RPT_OVD_STS
    .Range("HMI_TASK5_OVD_STS").Value = wsSource.Range("HMI_ESD_OVD_STS").Value ' GREEN $E$26 HMI_ESD_OVD_STS
    .Range("HMI_TASK6_OVD_STS").Value = wsSource.Range("HMI_TRD_OVD_STS").Value ' GREEN $E$27 HMI_TRD_OVD_STS
    .Range("HMI_TASK7_OVD_STS").Value = wsSource.Range("HMI_GRP_OVD_STS").Value ' GREEN $E$28 HMI_GRP_OVD_STS
    .Range("HMI_TASK8_OVD_STS").Value = wsSource.Range("HMI_OVL_OVD_STS").Value ' GREEN $E$29 HMI_OVL_OVD_STS
    .Range("HMI_TASK9_OVD_STS").Value = wsSource.Range("HMI_ENV_OVD_STS").Value ' GREEN $E$30 HMI_ENV_OVD_STS
    .Range("HMI_TASK10_OVD_STS").Value = wsSource.Range("HMI_ALARM_OVD_STS").Value ' GREEN $E$31 HMI_ALARM_OVD_STS
    .Range("CP_TASK1_OVD_STS").Value = wsSource.Range("CP_DESIGN_OVD_STS").Value ' GREEN $E$38 CP_DESIGN_OVD_STS
    .Range("CP_TASK2_OVD_STS").Value = wsSource.Range("CP_AI_OVD_STS").Value ' GREEN $E$39 CP_AI_OVD_STS
    .Range("CP_TASK3_OVD_STS").Value = wsSource.Range("CP_AO_OVD_STS").Value ' GREEN $E$40 CP_AO_OVD_STS
    .Range("CP_TASK4_OVD_STS").Value = wsSource.Range("CP_DI_OVD_STS").Value ' GREEN $E$41 CP_DI_OVD_STS
    .Range("CP_TASK5_OVD_STS").Value = wsSource.Range("CP_DO_OVD_STS").Value ' GREEN $E$42 CP_DO_OVD_STS
    .Range("CP_TASK6_OVD_STS").Value = wsSource.Range("CP_LOGIC_OVD_STS").Value ' GREEN $E$43   CP_LOGIC_OVD_STS
    .Range("CP_TASK7_OVD_STS").Value = wsSource.Range("CP_GRP_START_OVD_STS").Value ' GREEN $E$44   CP_GRP_START_OVD_STS
    .Range("CP_TASK8_OVD_STS").Value = wsSource.Range("CP_SEQ_OVD_STS").Value ' GREEN $E$45   CP_SEQ_OVD_STS
    .Range("DI_TASK1_OVD_STS").Value = wsSource.Range("DI_DESIGN_OVD_STS").Value ' GREEN $E$52 DI_DESIGN_OVD_STS
    .Range("DI_TASK2_OVD_STS").Value = wsSource.Range("DI_AI_OVD_STS").Value ' GREEN $E$53 DI_AI_OVD_STS
    .Range("DI_TASK3_OVD_STS").Value = wsSource.Range("DI_AO_OVD_STS").Value ' GREEN $E$54 DI_AO_OVD_STS
    .Range("DI_TASK4_OVD_STS").Value = wsSource.Range("DI_DI_OVD_STS").Value ' GREEN $E$55 DI_DI_OVD_STS
    .Range("DI_TASK5_OVD_STS").Value = wsSource.Range("DI_DO_OVD_STS").Value ' GREEN $E$56 DI_DO_OVD_STS
    .Range("DI_TASK6_OVD_STS").Value = wsSource.Range("DI_LOGIC_OVD_STS").Value ' GREEN $E$57   DI_LOGIC_OVD_STS
    .Range("DI_TASK7_OVD_STS").Value = wsSource.Range("DI_GRP_START_OVD_STS").Value ' GREEN $E$58   DI_GRP_START_OVD_STS
    .Range("DI_TASK8_OVD_STS").Value = wsSource.Range("DI_SEQ_OVD_STS").Value ' GREEN $E$59   DI_SEQ_OVD_STS
    .Range("ESD_TASK1_OVD_STS").Value = wsSource.Range("ESD_TRICON_DESIGN_OVD_STS").Value ' GREEN $E$66 ESD_TRICON_DESIGN_OVD_STS
    .Range("ESD_TASK2_OVD_STS").Value = wsSource.Range("ESD_AI_OVD_STS").Value ' GREEN $E$67   ESD_AI_OVD_STS
    .Range("ESD_TASK3_OVD_STS").Value = wsSource.Range("ESD_AO_OVD_STS").Value ' GREEN $E$68   ESD_AO_OVD_STS
    .Range("ESD_TASK4_OVD_STS").Value = wsSource.Range("ESD_DI_OVD_STS").Value ' GREEN $E$69   ESD_DI_OVD_STS
    .Range("ESD_TASK5_OVD_STS").Value = wsSource.Range("ESD_DO_OVD_STS").Value ' GREEN $E$70   ESD_DO_OVD_STS
    .Range("ESD_TASK6_OVD_STS").Value = wsSource.Range("ESD_TRICON_LOGIC_OVD_STS").Value ' GREEN $E$71   ESD_TRICON_LOGIC_OVD_STS
    .Range("ESD_TASK7_OVD_STS").Value = wsSource.Range("ESD_GRP_START_OVD_STS").Value ' GREEN $E$72 ESD_GRP_START_OVD_STS
    .Range("REP_TASK1_OVD_STS").Value = wsSource.Range("REP_POINTS_OVD_STS").Value ' GREEN $E$79   REP_POINTS_OVD_STS
    .Range("REP_TASK2_OVD_STS").Value = wsSource.Range("REP_STD_OVD_STS").Value ' GREEN $E$80 REP_STD_OVD_STS
    .Range("REP_TASK3_OVD_STS").Value = wsSource.Range("REP_CUSTOM_OVD_STS").Value ' GREEN $E$81   REP_CUSTOM_OVD_STS
    .Range("APP_TASK1_OVD_STS").Value = wsSource.Range("APP_1_OVD_STS").Value ' GREEN $E$88 APP_1_OVD_STS
    .Range("APP_TASK2_OVD_STS").Value = wsSource.Range("APP_2_OVD_STS").Value ' GREEN $E$89 APP_2_OVD_STS
    .Range("APP_TASK3_OVD_STS").Value = wsSource.Range("APP_3_OVD_STS").Value ' GREEN $E$90 APP_3_OVD_STS
    .Range("APP_TASK4_OVD_STS").Value = wsSource.Range("APP_4_OVD_STS").Value ' GREEN $E$91 APP_4_OVD_STS
    .Range("APP_TASK5_OVD_STS").Value = wsSource.Range("APP_5_OVD_STS").Value ' GREEN $E$92 APP_5_OVD_STS
    .Range("APP_TASK6_OVD_STS").Value = wsSource.Range("APP_6_OVD_STS").Value ' GREEN $E$93 APP_6_OVD_STS
    .Range("APP_TASK7_OVD_STS").Value = wsSource.Range("APP_7_OVD_STS").Value ' GREEN $E$94 APP_7_OVD_STS
    .Range("APP_TASK8_OVD_STS").Value = wsSource.Range("APP_8_OVD_STS").Value ' GREEN $E$95 APP_8_OVD_STS
    .Range("APP_TASK9_OVD_STS").Value = wsSource.Range("APP_BUS_OVD_STS").Value ' GREEN $E$96 APP_BUS_OVD_STS
    .Range("TEST_TASK1_OVD_STS").Value = wsSource.Range("TEST_FAT_OVD_STS").Value
    .Range("TEST_TASK2_OVD_STS").Value = wsSource.Range("TEST_IO_OVD_STS").Value
    .Range("TEST_TASK3_OVD_STS").Value = wsSource.Range("TEST_SI_OVD_STS").Value
    .Range("TEST_TASK4_OVD_STS").Value = wsSource.Range("TEST_PACK_OVD_STS").Value
    .Range("TEST_TASK5_OVD_STS").Value = wsSource.Range("TEST_SIM_OVD_STS").Value
    .Range("TEST_TASK6_OVD_STS").Value = wsSource.Range("TEST_RENT_OVD_STS").Value
    .Range("DOC_TASK1_OVD_STS").Value = wsSource.Range("DOC_BOM_OVD_STS").Value
    .Range("DOC_TASK2_OVD_STS").Value = wsSource.Range("DOC_SYS_ARCH_OVD_STS").Value
    .Range("DOC_TASK3_OVD_STS").Value = wsSource.Range("DOC_SYS_INT_OVD_STS").Value
    .Range("DOC_TASK4_OVD_STS").Value = wsSource.Range("DOC_PWR_GND_OVD_STS").Value
    .Range("DOC_TASK5_OVD_STS").Value = wsSource.Range("DOC_CAB_MECH_OVD_STS").Value
    .Range("DOC_TASK6_OVD_STS").Value = wsSource.Range("DOC_CAB_ELEC_OVD_STS").Value
    .Range("DOC_TASK7_OVD_STS").Value = wsSource.Range("DOC_PWR_HEAT_OVD_STS").Value
    .Range("DOC_TASK8_OVD_STS").Value = wsSource.Range("DOC_LOOP_OVD_STS").Value
    .Range("DOC_TASK9_OVD_STS").Value = wsSource.Range("DOC_CUSTOM_OVD_STS").Value
    .Range("DOC_TASK10_OVD_STS").Value = wsSource.Range("DOC_TAGLIST_OVD_STS").Value
    .Range("COURSE_TASK1_OVD_STS").Value = wsSource.Range("COURSE_1_OVD_STS").Value
    .Range("COURSE_TASK2_OVD_STS").Value = wsSource.Range("COURSE_2_OVD_STS").Value
    .Range("COURSE_TASK3_OVD_STS").Value = wsSource.Range("COURSE_3_OVD_STS").Value
    .Range("COURSE_TASK4_OVD_STS").Value = wsSource.Range("COURSE_4_OVD_STS").Value
    .Range("COURSE_TASK5_OVD_STS").Value = wsSource.Range("COURSE_5_OVD_STS").Value
    .Range("PM_TASK1_OVD_STS").Value = wsSource.Range("PM_IA_OVD_STS").Value
    .Range("PM_TASK2_OVD_STS").Value = wsSource.Range("PM_BUYOUT_OVD_STS").Value
    .Range("MEETING_TASK1_OVD_STS").Value = wsSource.Range("MEETING_KICKOFF_OVD_STS").Value
    .Range("MEETING_TASK2_OVD_STS").Value = wsSource.Range("MEETING_DESIGN_OVD_STS").Value
    .Range("MEETING_TASK3_OVD_STS").Value = wsSource.Range("MEETING_PROGRESS_OVD_STS").Value
    .Range("MEETING_TASK4_OVD_STS").Value = wsSource.Range("MEETING_OTHER_OVD_STS").Value
    .Range("MEETING_TASK5_OVD_STS").Value = wsSource.Range("MEETING_CLOSE_OVD_STS").Value
    .Range("SITE_TASK1_OVD_STS").Value = wsSource.Range("SITE_SURVEY_OVD_STS").Value
    .Range("SITE_TASK2_OVD_STS").Value = wsSource.Range("SITE_PWRUP_OVD_STS").Value
    .Range("SITE_TASK3_OVD_STS").Value = wsSource.Range("SITE_COMM_OVD_STS").Value
    .Range("SITE_TASK4_OVD_STS").Value = wsSource.Range("SITE_SAT_OVD_STS").Value
    .Range("TL_TASK1_OVD_STS").Value = wsSource.Range("TL_ENTER_OVD_STS").Value
    .Range("TL_TASK2_OVD_STS").Value = wsSource.Range("TL_TL_OVD_STS").Value
    .Range("TL_TASK3_OVD_STS").Value = wsSource.Range("TL_REMOTE_OVD_STS").Value
    .Range("TL_TASK4_OVD_STS").Value = wsSource.Range("TL_SITE_OVD_STS").Value
End With


End Sub

Sub ImportApplicationBasedSheet_428(wsSource As Worksheet, wsTarget As Worksheet)
'GECE Version = 1.0
On Error Resume Next
With wsTarget
    If .Range("APP_TASK1_OVD_JUST").Value <> wsSource.Range("APP_TASK1_OVD_JUST").Value Then
        .Range("APP_TASK1_OVD_JUST").Value = wsSource.Range("APP_TASK1_OVD_JUST").Value ' $I$280
    End If
    If .Range("APP_TASK1_OVD_QTY").Value <> wsSource.Range("APP_TASK1_OVD_QTY").Value Then
        .Range("APP_TASK1_OVD_QTY").Value = wsSource.Range("APP_TASK1_OVD_QTY").Value ' $H$280
    End If
    If .Range("APP_TASK1_REM_COUNTRY").Value <> wsSource.Range("APP_TASK1_REM_COUNTRY").Value Then
        .Range("APP_TASK1_REM_COUNTRY").Value = wsSource.Range("APP_TASK1_REM_COUNTRY").Value ' $M$280
    End If
    If .Range("APP_TASK1_REM_PCT").Value <> wsSource.Range("APP_TASK1_REM_PCT").Value Then
        .Range("APP_TASK1_REM_PCT").Value = wsSource.Range("APP_TASK1_REM_PCT").Value ' $N$280
    End If
    If .Range("APP_TASK10_OVD_JUST").Value <> wsSource.Range("APP_TASK10_OVD_JUST").Value Then
        .Range("APP_TASK10_OVD_JUST").Value = wsSource.Range("APP_TASK10_OVD_JUST").Value ' $I$289
    End If
    If .Range("APP_TASK10_OVD_QTY").Value <> wsSource.Range("APP_TASK10_OVD_QTY").Value Then
        .Range("APP_TASK10_OVD_QTY").Value = wsSource.Range("APP_TASK10_OVD_QTY").Value ' $H$289
    End If
    If .Range("APP_TASK10_REM_COUNTRY").Value <> wsSource.Range("APP_TASK10_REM_COUNTRY").Value Then
        .Range("APP_TASK10_REM_COUNTRY").Value = wsSource.Range("APP_TASK10_REM_COUNTRY").Value ' $M$289
    End If
    If .Range("APP_TASK10_REM_PCT").Value <> wsSource.Range("APP_TASK10_REM_PCT").Value Then
        .Range("APP_TASK10_REM_PCT").Value = wsSource.Range("APP_TASK10_REM_PCT").Value ' $N$289
    End If
    If .Range("APP_TASK2_OVD_JUST").Value <> wsSource.Range("APP_TASK2_OVD_JUST").Value Then
        .Range("APP_TASK2_OVD_JUST").Value = wsSource.Range("APP_TASK2_OVD_JUST").Value ' $I$281
    End If
    If .Range("APP_TASK2_OVD_QTY").Value <> wsSource.Range("APP_TASK2_OVD_QTY").Value Then
        .Range("APP_TASK2_OVD_QTY").Value = wsSource.Range("APP_TASK2_OVD_QTY").Value ' $H$281
    End If
    If .Range("APP_TASK2_REM_COUNTRY").Value <> wsSource.Range("APP_TASK2_REM_COUNTRY").Value Then
        .Range("APP_TASK2_REM_COUNTRY").Value = wsSource.Range("APP_TASK2_REM_COUNTRY").Value ' $M$281
    End If
    If .Range("APP_TASK2_REM_PCT").Value <> wsSource.Range("APP_TASK2_REM_PCT").Value Then
        .Range("APP_TASK2_REM_PCT").Value = wsSource.Range("APP_TASK2_REM_PCT").Value ' $N$281
    End If
    If .Range("APP_TASK3_OVD_JUST").Value <> wsSource.Range("APP_TASK3_OVD_JUST").Value Then
        .Range("APP_TASK3_OVD_JUST").Value = wsSource.Range("APP_TASK3_OVD_JUST").Value ' $I$282
    End If
    If .Range("APP_TASK3_OVD_QTY").Value <> wsSource.Range("APP_TASK3_OVD_QTY").Value Then
        .Range("APP_TASK3_OVD_QTY").Value = wsSource.Range("APP_TASK3_OVD_QTY").Value ' $H$282
    End If
    If .Range("APP_TASK3_REM_COUNTRY").Value <> wsSource.Range("APP_TASK3_REM_COUNTRY").Value Then
        .Range("APP_TASK3_REM_COUNTRY").Value = wsSource.Range("APP_TASK3_REM_COUNTRY").Value ' $M$282
    End If
    If .Range("APP_TASK3_REM_PCT").Value <> wsSource.Range("APP_TASK3_REM_PCT").Value Then
        .Range("APP_TASK3_REM_PCT").Value = wsSource.Range("APP_TASK3_REM_PCT").Value ' $N$282
    End If
    If .Range("APP_TASK4_OVD_JUST").Value <> wsSource.Range("APP_TASK4_OVD_JUST").Value Then
        .Range("APP_TASK4_OVD_JUST").Value = wsSource.Range("APP_TASK4_OVD_JUST").Value ' $I$283
    End If
    If .Range("APP_TASK4_OVD_QTY").Value <> wsSource.Range("APP_TASK4_OVD_QTY").Value Then
        .Range("APP_TASK4_OVD_QTY").Value = wsSource.Range("APP_TASK4_OVD_QTY").Value ' $H$283
    End If
    If .Range("APP_TASK4_REM_COUNTRY").Value <> wsSource.Range("APP_TASK4_REM_COUNTRY").Value Then
        .Range("APP_TASK4_REM_COUNTRY").Value = wsSource.Range("APP_TASK4_REM_COUNTRY").Value ' $M$283
    End If
    If .Range("APP_TASK4_REM_PCT").Value <> wsSource.Range("APP_TASK4_REM_PCT").Value Then
        .Range("APP_TASK4_REM_PCT").Value = wsSource.Range("APP_TASK4_REM_PCT").Value ' $N$283
    End If
    If .Range("APP_TASK5_OVD_JUST").Value <> wsSource.Range("APP_TASK5_OVD_JUST").Value Then
        .Range("APP_TASK5_OVD_JUST").Value = wsSource.Range("APP_TASK5_OVD_JUST").Value ' $I$284
    End If
    If .Range("APP_TASK5_OVD_QTY").Value <> wsSource.Range("APP_TASK5_OVD_QTY").Value Then
        .Range("APP_TASK5_OVD_QTY").Value = wsSource.Range("APP_TASK5_OVD_QTY").Value ' $H$284
    End If
    If .Range("APP_TASK5_REM_COUNTRY").Value <> wsSource.Range("APP_TASK5_REM_COUNTRY").Value Then
        .Range("APP_TASK5_REM_COUNTRY").Value = wsSource.Range("APP_TASK5_REM_COUNTRY").Value ' $M$284
    End If
    If .Range("APP_TASK5_REM_PCT").Value <> wsSource.Range("APP_TASK5_REM_PCT").Value Then
        .Range("APP_TASK5_REM_PCT").Value = wsSource.Range("APP_TASK5_REM_PCT").Value ' $N$284
    End If
    If .Range("APP_TASK6_OVD_JUST").Value <> wsSource.Range("APP_TASK6_OVD_JUST").Value Then
        .Range("APP_TASK6_OVD_JUST").Value = wsSource.Range("APP_TASK6_OVD_JUST").Value ' $I$285
    End If
    If .Range("APP_TASK6_OVD_QTY").Value <> wsSource.Range("APP_TASK6_OVD_QTY").Value Then
        .Range("APP_TASK6_OVD_QTY").Value = wsSource.Range("APP_TASK6_OVD_QTY").Value ' $H$285
    End If
    If .Range("APP_TASK6_REM_COUNTRY").Value <> wsSource.Range("APP_TASK6_REM_COUNTRY").Value Then
        .Range("APP_TASK6_REM_COUNTRY").Value = wsSource.Range("APP_TASK6_REM_COUNTRY").Value ' $M$285
    End If
    If .Range("APP_TASK6_REM_PCT").Value <> wsSource.Range("APP_TASK6_REM_PCT").Value Then
        .Range("APP_TASK6_REM_PCT").Value = wsSource.Range("APP_TASK6_REM_PCT").Value ' $N$285
    End If
    If .Range("APP_TASK7_OVD_JUST").Value <> wsSource.Range("APP_TASK7_OVD_JUST").Value Then
        .Range("APP_TASK7_OVD_JUST").Value = wsSource.Range("APP_TASK7_OVD_JUST").Value ' $I$286
    End If
    If .Range("APP_TASK7_OVD_QTY").Value <> wsSource.Range("APP_TASK7_OVD_QTY").Value Then
        .Range("APP_TASK7_OVD_QTY").Value = wsSource.Range("APP_TASK7_OVD_QTY").Value ' $H$286
    End If
    If .Range("APP_TASK7_REM_COUNTRY").Value <> wsSource.Range("APP_TASK7_REM_COUNTRY").Value Then
        .Range("APP_TASK7_REM_COUNTRY").Value = wsSource.Range("APP_TASK7_REM_COUNTRY").Value ' $M$286
    End If
    If .Range("APP_TASK7_REM_PCT").Value <> wsSource.Range("APP_TASK7_REM_PCT").Value Then
        .Range("APP_TASK7_REM_PCT").Value = wsSource.Range("APP_TASK7_REM_PCT").Value ' $N$286
    End If
    If .Range("APP_TASK8_OVD_JUST").Value <> wsSource.Range("APP_TASK8_OVD_JUST").Value Then
        .Range("APP_TASK8_OVD_JUST").Value = wsSource.Range("APP_TASK8_OVD_JUST").Value ' $I$287
    End If
    If .Range("APP_TASK8_OVD_QTY").Value <> wsSource.Range("APP_TASK8_OVD_QTY").Value Then
        .Range("APP_TASK8_OVD_QTY").Value = wsSource.Range("APP_TASK8_OVD_QTY").Value ' $H$287
    End If
    If .Range("APP_TASK8_REM_COUNTRY").Value <> wsSource.Range("APP_TASK8_REM_COUNTRY").Value Then
        .Range("APP_TASK8_REM_COUNTRY").Value = wsSource.Range("APP_TASK8_REM_COUNTRY").Value ' $M$287
    End If
    If .Range("APP_TASK8_REM_PCT").Value <> wsSource.Range("APP_TASK8_REM_PCT").Value Then
        .Range("APP_TASK8_REM_PCT").Value = wsSource.Range("APP_TASK8_REM_PCT").Value ' $N$287
    End If
    If .Range("APP_TASK9_OVD_JUST").Value <> wsSource.Range("APP_TASK9_OVD_JUST").Value Then
        .Range("APP_TASK9_OVD_JUST").Value = wsSource.Range("APP_TASK9_OVD_JUST").Value ' $I$288
    End If
    If .Range("APP_TASK9_OVD_QTY").Value <> wsSource.Range("APP_TASK9_OVD_QTY").Value Then
        .Range("APP_TASK9_OVD_QTY").Value = wsSource.Range("APP_TASK9_OVD_QTY").Value ' $H$288
    End If
    If .Range("APP_TASK9_REM_COUNTRY").Value <> wsSource.Range("APP_TASK9_REM_COUNTRY").Value Then
        .Range("APP_TASK9_REM_COUNTRY").Value = wsSource.Range("APP_TASK9_REM_COUNTRY").Value ' $M$288
    End If
    If .Range("APP_TASK9_REM_PCT").Value <> wsSource.Range("APP_TASK9_REM_PCT").Value Then
        .Range("APP_TASK9_REM_PCT").Value = wsSource.Range("APP_TASK9_REM_PCT").Value ' $N$288
    End If
    If .Range("COURSE_REM_COST").Value <> wsSource.Range("COURSE_REM_COST").Value Then
        .Range("COURSE_REM_COST").Value = wsSource.Range("APP_TASK9_REM_PCT").Value ' $N$288
    End If
    If .Range("COURSE_TASK1_OVD_JUST").Value <> wsSource.Range("COURSE_TASK1_OVD_JUST").Value Then
        .Range("COURSE_TASK1_OVD_JUST").Value = wsSource.Range("COURSE_TASK1_OVD_JUST").Value ' $I$343
    End If
    If .Range("COURSE_TASK1_OVD_QTY").Value <> wsSource.Range("COURSE_TASK1_OVD_QTY").Value Then
        .Range("COURSE_TASK1_OVD_QTY").Value = wsSource.Range("COURSE_TASK1_OVD_QTY").Value ' $H$343
    End If
    If .Range("COURSE_TASK1_REM_COUNTRY").Value <> wsSource.Range("COURSE_TASK1_REM_COUNTRY").Value Then
        .Range("COURSE_TASK1_REM_COUNTRY").Value = wsSource.Range("COURSE_TASK1_REM_COUNTRY").Value ' $M$343
    End If
    If .Range("COURSE_TASK1_REM_PCT").Value <> wsSource.Range("COURSE_TASK1_REM_PCT").Value Then
        .Range("COURSE_TASK1_REM_PCT").Value = wsSource.Range("COURSE_TASK1_REM_PCT").Value ' $N$343
    End If
    If .Range("COURSE_TASK10_OVD_JUST").Value <> wsSource.Range("COURSE_TASK10_OVD_JUST").Value Then
        .Range("COURSE_TASK10_OVD_JUST").Value = wsSource.Range("COURSE_TASK10_OVD_JUST").Value ' $I$352
    End If
    If .Range("COURSE_TASK10_OVD_QTY").Value <> wsSource.Range("COURSE_TASK10_OVD_QTY").Value Then
        .Range("COURSE_TASK10_OVD_QTY").Value = wsSource.Range("COURSE_TASK10_OVD_QTY").Value ' $H$352
    End If
    If .Range("COURSE_TASK10_REM_COUNTRY").Value <> wsSource.Range("COURSE_TASK10_REM_COUNTRY").Value Then
        .Range("COURSE_TASK10_REM_COUNTRY").Value = wsSource.Range("COURSE_TASK10_REM_COUNTRY").Value ' $M$352
    End If
    If .Range("COURSE_TASK10_REM_PCT").Value <> wsSource.Range("COURSE_TASK10_REM_PCT").Value Then
        .Range("COURSE_TASK10_REM_PCT").Value = wsSource.Range("COURSE_TASK10_REM_PCT").Value ' $N$352
    End If
    If .Range("COURSE_TASK2_OVD_JUST").Value <> wsSource.Range("COURSE_TASK2_OVD_JUST").Value Then
        .Range("COURSE_TASK2_OVD_JUST").Value = wsSource.Range("COURSE_TASK2_OVD_JUST").Value ' $I$344
    End If
    If .Range("COURSE_TASK2_OVD_QTY").Value <> wsSource.Range("COURSE_TASK2_OVD_QTY").Value Then
        .Range("COURSE_TASK2_OVD_QTY").Value = wsSource.Range("COURSE_TASK2_OVD_QTY").Value ' $H$344
    End If
    If .Range("COURSE_TASK2_REM_COUNTRY").Value <> wsSource.Range("COURSE_TASK2_REM_COUNTRY").Value Then
        .Range("COURSE_TASK2_REM_COUNTRY").Value = wsSource.Range("COURSE_TASK2_REM_COUNTRY").Value ' $M$344
    End If
    If .Range("COURSE_TASK2_REM_PCT").Value <> wsSource.Range("COURSE_TASK2_REM_PCT").Value Then
        .Range("COURSE_TASK2_REM_PCT").Value = wsSource.Range("COURSE_TASK2_REM_PCT").Value ' $N$344
    End If
    If .Range("COURSE_TASK3_OVD_JUST").Value <> wsSource.Range("COURSE_TASK3_OVD_JUST").Value Then
        .Range("COURSE_TASK3_OVD_JUST").Value = wsSource.Range("COURSE_TASK3_OVD_JUST").Value ' $I$345
    End If
    If .Range("COURSE_TASK3_OVD_QTY").Value <> wsSource.Range("COURSE_TASK3_OVD_QTY").Value Then
        .Range("COURSE_TASK3_OVD_QTY").Value = wsSource.Range("COURSE_TASK3_OVD_QTY").Value ' $H$345
    End If
    If .Range("COURSE_TASK3_REM_COUNTRY").Value <> wsSource.Range("COURSE_TASK3_REM_COUNTRY").Value Then
        .Range("COURSE_TASK3_REM_COUNTRY").Value = wsSource.Range("COURSE_TASK3_REM_COUNTRY").Value ' $M$345
    End If
    If .Range("COURSE_TASK3_REM_PCT").Value <> wsSource.Range("COURSE_TASK3_REM_PCT").Value Then
        .Range("COURSE_TASK3_REM_PCT").Value = wsSource.Range("COURSE_TASK3_REM_PCT").Value ' $N$345
    End If
    If .Range("COURSE_TASK4_OVD_JUST").Value <> wsSource.Range("COURSE_TASK4_OVD_JUST").Value Then
        .Range("COURSE_TASK4_OVD_JUST").Value = wsSource.Range("COURSE_TASK4_OVD_JUST").Value ' $I$346
    End If
    If .Range("COURSE_TASK4_OVD_QTY").Value <> wsSource.Range("COURSE_TASK4_OVD_QTY").Value Then
        .Range("COURSE_TASK4_OVD_QTY").Value = wsSource.Range("COURSE_TASK4_OVD_QTY").Value ' $H$346
    End If
    If .Range("COURSE_TASK4_REM_COUNTRY").Value <> wsSource.Range("COURSE_TASK4_REM_COUNTRY").Value Then
        .Range("COURSE_TASK4_REM_COUNTRY").Value = wsSource.Range("COURSE_TASK4_REM_COUNTRY").Value ' $M$346
    End If
    If .Range("COURSE_TASK4_REM_PCT").Value <> wsSource.Range("COURSE_TASK4_REM_PCT").Value Then
        .Range("COURSE_TASK4_REM_PCT").Value = wsSource.Range("COURSE_TASK4_REM_PCT").Value ' $N$346
    End If
    If .Range("COURSE_TASK5_OVD_JUST").Value <> wsSource.Range("COURSE_TASK5_OVD_JUST").Value Then
        .Range("COURSE_TASK5_OVD_JUST").Value = wsSource.Range("COURSE_TASK5_OVD_JUST").Value ' $I$347
    End If
    If .Range("COURSE_TASK5_OVD_QTY").Value <> wsSource.Range("COURSE_TASK5_OVD_QTY").Value Then
        .Range("COURSE_TASK5_OVD_QTY").Value = wsSource.Range("COURSE_TASK5_OVD_QTY").Value ' $H$347
    End If
    If .Range("COURSE_TASK5_REM_COUNTRY").Value <> wsSource.Range("COURSE_TASK5_REM_COUNTRY").Value Then
        .Range("COURSE_TASK5_REM_COUNTRY").Value = wsSource.Range("COURSE_TASK5_REM_COUNTRY").Value ' $M$347
    End If
    If .Range("COURSE_TASK5_REM_PCT").Value <> wsSource.Range("COURSE_TASK5_REM_PCT").Value Then
        .Range("COURSE_TASK5_REM_PCT").Value = wsSource.Range("COURSE_TASK5_REM_PCT").Value ' $N$347
    End If
    If .Range("COURSE_TASK6_OVD_JUST").Value <> wsSource.Range("COURSE_TASK6_OVD_JUST").Value Then
        .Range("COURSE_TASK6_OVD_JUST").Value = wsSource.Range("COURSE_TASK6_OVD_JUST").Value ' $I$348
    End If
    If .Range("COURSE_TASK6_OVD_QTY").Value <> wsSource.Range("COURSE_TASK6_OVD_QTY").Value Then
        .Range("COURSE_TASK6_OVD_QTY").Value = wsSource.Range("COURSE_TASK6_OVD_QTY").Value ' $H$348
    End If
    If .Range("COURSE_TASK6_REM_COUNTRY").Value <> wsSource.Range("COURSE_TASK6_REM_COUNTRY").Value Then
        .Range("COURSE_TASK6_REM_COUNTRY").Value = wsSource.Range("COURSE_TASK6_REM_COUNTRY").Value ' $M$348
    End If
    If .Range("COURSE_TASK6_REM_PCT").Value <> wsSource.Range("COURSE_TASK6_REM_PCT").Value Then
        .Range("COURSE_TASK6_REM_PCT").Value = wsSource.Range("COURSE_TASK6_REM_PCT").Value ' $N$348
    End If
    If .Range("COURSE_TASK7_OVD_JUST").Value <> wsSource.Range("COURSE_TASK7_OVD_JUST").Value Then
        .Range("COURSE_TASK7_OVD_JUST").Value = wsSource.Range("COURSE_TASK7_OVD_JUST").Value ' $I$349
    End If
    If .Range("COURSE_TASK7_OVD_QTY").Value <> wsSource.Range("COURSE_TASK7_OVD_QTY").Value Then
        .Range("COURSE_TASK7_OVD_QTY").Value = wsSource.Range("COURSE_TASK7_OVD_QTY").Value ' $H$349
    End If
    If .Range("COURSE_TASK7_REM_COUNTRY").Value <> wsSource.Range("COURSE_TASK7_REM_COUNTRY").Value Then
        .Range("COURSE_TASK7_REM_COUNTRY").Value = wsSource.Range("COURSE_TASK7_REM_COUNTRY").Value ' $M$349
    End If
    If .Range("COURSE_TASK7_REM_PCT").Value <> wsSource.Range("COURSE_TASK7_REM_PCT").Value Then
        .Range("COURSE_TASK7_REM_PCT").Value = wsSource.Range("COURSE_TASK7_REM_PCT").Value ' $N$349
    End If
    If .Range("COURSE_TASK8_OVD_JUST").Value <> wsSource.Range("COURSE_TASK8_OVD_JUST").Value Then
        .Range("COURSE_TASK8_OVD_JUST").Value = wsSource.Range("COURSE_TASK8_OVD_JUST").Value ' $I$350
    End If
    If .Range("COURSE_TASK8_OVD_QTY").Value <> wsSource.Range("COURSE_TASK8_OVD_QTY").Value Then
        .Range("COURSE_TASK8_OVD_QTY").Value = wsSource.Range("COURSE_TASK8_OVD_QTY").Value ' $H$350
    End If
    If .Range("COURSE_TASK8_REM_COUNTRY").Value <> wsSource.Range("COURSE_TASK8_REM_COUNTRY").Value Then
        .Range("COURSE_TASK8_REM_COUNTRY").Value = wsSource.Range("COURSE_TASK8_REM_COUNTRY").Value ' $M$350
    End If
    If .Range("COURSE_TASK8_REM_PCT").Value <> wsSource.Range("COURSE_TASK8_REM_PCT").Value Then
        .Range("COURSE_TASK8_REM_PCT").Value = wsSource.Range("COURSE_TASK8_REM_PCT").Value ' $N$350
    End If
    If .Range("COURSE_TASK9_OVD_JUST").Value <> wsSource.Range("COURSE_TASK9_OVD_JUST").Value Then
        .Range("COURSE_TASK9_OVD_JUST").Value = wsSource.Range("COURSE_TASK9_OVD_JUST").Value ' $I$351
    End If
    If .Range("COURSE_TASK9_OVD_QTY").Value <> wsSource.Range("COURSE_TASK9_OVD_QTY").Value Then
        .Range("COURSE_TASK9_OVD_QTY").Value = wsSource.Range("COURSE_TASK9_OVD_QTY").Value ' $H$351
    End If
    If .Range("COURSE_TASK9_REM_COUNTRY").Value <> wsSource.Range("COURSE_TASK9_REM_COUNTRY").Value Then
        .Range("COURSE_TASK9_REM_COUNTRY").Value = wsSource.Range("COURSE_TASK9_REM_COUNTRY").Value ' $M$351
    End If
    If .Range("COURSE_TASK9_REM_PCT").Value <> wsSource.Range("COURSE_TASK9_REM_PCT").Value Then
        .Range("COURSE_TASK9_REM_PCT").Value = wsSource.Range("COURSE_TASK9_REM_PCT").Value ' $N$351
    End If
    If .Range("CP_REM_COST").Value <> wsSource.Range("CP_REM_COST").Value Then
        .Range("CP_REM_COST").Value = wsSource.Range("COURSE_TASK9_REM_PCT").Value ' $N$351
    End If
    If .Range("CP_TASK1_OVD_JUST").Value <> wsSource.Range("CP_TASK1_OVD_JUST").Value Then
        .Range("CP_TASK1_OVD_JUST").Value = wsSource.Range("CP_TASK1_OVD_JUST").Value ' $I$193
    End If
    If .Range("CP_TASK1_OVD_QTY").Value <> wsSource.Range("CP_TASK1_OVD_QTY").Value Then
        .Range("CP_TASK1_OVD_QTY").Value = wsSource.Range("CP_TASK1_OVD_QTY").Value ' $H$193
    End If
    If .Range("CP_TASK1_REM_COUNTRY").Value <> wsSource.Range("CP_TASK1_REM_COUNTRY").Value Then
        .Range("CP_TASK1_REM_COUNTRY").Value = wsSource.Range("CP_TASK1_REM_COUNTRY").Value ' $M$193
    End If
    If .Range("CP_TASK1_REM_PCT").Value <> wsSource.Range("CP_TASK1_REM_PCT").Value Then
        .Range("CP_TASK1_REM_PCT").Value = wsSource.Range("CP_TASK1_REM_PCT").Value ' $N$193
    End If
    If .Range("CP_TASK10_OVD_JUST").Value <> wsSource.Range("CP_TASK10_OVD_JUST").Value Then
        .Range("CP_TASK10_OVD_JUST").Value = wsSource.Range("CP_TASK10_OVD_JUST").Value ' $I$202
    End If
    If .Range("CP_TASK10_OVD_QTY").Value <> wsSource.Range("CP_TASK10_OVD_QTY").Value Then
        .Range("CP_TASK10_OVD_QTY").Value = wsSource.Range("CP_TASK10_OVD_QTY").Value ' $H$202
    End If
    If .Range("CP_TASK10_REM_COUNTRY").Value <> wsSource.Range("CP_TASK10_REM_COUNTRY").Value Then
        .Range("CP_TASK10_REM_COUNTRY").Value = wsSource.Range("CP_TASK10_REM_COUNTRY").Value ' $M$202
    End If
    If .Range("CP_TASK10_REM_PCT").Value <> wsSource.Range("CP_TASK10_REM_PCT").Value Then
        .Range("CP_TASK10_REM_PCT").Value = wsSource.Range("CP_TASK10_REM_PCT").Value ' $N$202
    End If
    If .Range("CP_TASK2_OVD_JUST").Value <> wsSource.Range("CP_TASK2_OVD_JUST").Value Then
        .Range("CP_TASK2_OVD_JUST").Value = wsSource.Range("CP_TASK2_OVD_JUST").Value ' $I$194
    End If
    If .Range("CP_TASK2_OVD_QTY").Value <> wsSource.Range("CP_TASK2_OVD_QTY").Value Then
        .Range("CP_TASK2_OVD_QTY").Value = wsSource.Range("CP_TASK2_OVD_QTY").Value ' $H$194
    End If
    If .Range("CP_TASK2_REM_COUNTRY").Value <> wsSource.Range("CP_TASK2_REM_COUNTRY").Value Then
        .Range("CP_TASK2_REM_COUNTRY").Value = wsSource.Range("CP_TASK2_REM_COUNTRY").Value ' $M$194
    End If
    If .Range("CP_TASK2_REM_PCT").Value <> wsSource.Range("CP_TASK2_REM_PCT").Value Then
        .Range("CP_TASK2_REM_PCT").Value = wsSource.Range("CP_TASK2_REM_PCT").Value ' $N$194
    End If
    If .Range("CP_TASK3_OVD_JUST").Value <> wsSource.Range("CP_TASK3_OVD_JUST").Value Then
        .Range("CP_TASK3_OVD_JUST").Value = wsSource.Range("CP_TASK3_OVD_JUST").Value ' $I$195
    End If
    If .Range("CP_TASK3_OVD_QTY").Value <> wsSource.Range("CP_TASK3_OVD_QTY").Value Then
        .Range("CP_TASK3_OVD_QTY").Value = wsSource.Range("CP_TASK3_OVD_QTY").Value ' $H$195
    End If
    If .Range("CP_TASK3_REM_COUNTRY").Value <> wsSource.Range("CP_TASK3_REM_COUNTRY").Value Then
        .Range("CP_TASK3_REM_COUNTRY").Value = wsSource.Range("CP_TASK3_REM_COUNTRY").Value ' $M$195
    End If
    If .Range("CP_TASK3_REM_PCT").Value <> wsSource.Range("CP_TASK3_REM_PCT").Value Then
        .Range("CP_TASK3_REM_PCT").Value = wsSource.Range("CP_TASK3_REM_PCT").Value ' $N$195
    End If
    If .Range("CP_TASK4_OVD_JUST").Value <> wsSource.Range("CP_TASK4_OVD_JUST").Value Then
        .Range("CP_TASK4_OVD_JUST").Value = wsSource.Range("CP_TASK4_OVD_JUST").Value ' $I$196
    End If
    If .Range("CP_TASK4_OVD_QTY").Value <> wsSource.Range("CP_TASK4_OVD_QTY").Value Then
        .Range("CP_TASK4_OVD_QTY").Value = wsSource.Range("CP_TASK4_OVD_QTY").Value ' $H$196
    End If
    If .Range("CP_TASK4_REM_COUNTRY").Value <> wsSource.Range("CP_TASK4_REM_COUNTRY").Value Then
        .Range("CP_TASK4_REM_COUNTRY").Value = wsSource.Range("CP_TASK4_REM_COUNTRY").Value ' $M$196
    End If
    If .Range("CP_TASK4_REM_PCT").Value <> wsSource.Range("CP_TASK4_REM_PCT").Value Then
        .Range("CP_TASK4_REM_PCT").Value = wsSource.Range("CP_TASK4_REM_PCT").Value ' $N$196
    End If
    If .Range("CP_TASK5_OVD_JUST").Value <> wsSource.Range("CP_TASK5_OVD_JUST").Value Then
        .Range("CP_TASK5_OVD_JUST").Value = wsSource.Range("CP_TASK5_OVD_JUST").Value ' $I$197
    End If
    If .Range("CP_TASK5_OVD_QTY").Value <> wsSource.Range("CP_TASK5_OVD_QTY").Value Then
        .Range("CP_TASK5_OVD_QTY").Value = wsSource.Range("CP_TASK5_OVD_QTY").Value ' $H$197
    End If
    If .Range("CP_TASK5_REM_COUNTRY").Value <> wsSource.Range("CP_TASK5_REM_COUNTRY").Value Then
        .Range("CP_TASK5_REM_COUNTRY").Value = wsSource.Range("CP_TASK5_REM_COUNTRY").Value ' $M$197
    End If
    If .Range("CP_TASK5_REM_PCT").Value <> wsSource.Range("CP_TASK5_REM_PCT").Value Then
        .Range("CP_TASK5_REM_PCT").Value = wsSource.Range("CP_TASK5_REM_PCT").Value ' $N$197
    End If
    If .Range("CP_TASK6_OVD_JUST").Value <> wsSource.Range("CP_TASK6_OVD_JUST").Value Then
        .Range("CP_TASK6_OVD_JUST").Value = wsSource.Range("CP_TASK6_OVD_JUST").Value ' $I$198
    End If
    If .Range("CP_TASK6_OVD_QTY").Value <> wsSource.Range("CP_TASK6_OVD_QTY").Value Then
        .Range("CP_TASK6_OVD_QTY").Value = wsSource.Range("CP_TASK6_OVD_QTY").Value ' $H$198
    End If
    If .Range("CP_TASK6_REM_COUNTRY").Value <> wsSource.Range("CP_TASK6_REM_COUNTRY").Value Then
        .Range("CP_TASK6_REM_COUNTRY").Value = wsSource.Range("CP_TASK6_REM_COUNTRY").Value ' $M$198
    End If
    If .Range("CP_TASK6_REM_PCT").Value <> wsSource.Range("CP_TASK6_REM_PCT").Value Then
        .Range("CP_TASK6_REM_PCT").Value = wsSource.Range("CP_TASK6_REM_PCT").Value ' $N$198
    End If
    If .Range("CP_TASK7_OVD_JUST").Value <> wsSource.Range("CP_TASK7_OVD_JUST").Value Then
        .Range("CP_TASK7_OVD_JUST").Value = wsSource.Range("CP_TASK7_OVD_JUST").Value ' $I$199
    End If
    If .Range("CP_TASK7_OVD_QTY").Value <> wsSource.Range("CP_TASK7_OVD_QTY").Value Then
        .Range("CP_TASK7_OVD_QTY").Value = wsSource.Range("CP_TASK7_OVD_QTY").Value ' $H$199
    End If
    If .Range("CP_TASK7_REM_COUNTRY").Value <> wsSource.Range("CP_TASK7_REM_COUNTRY").Value Then
        .Range("CP_TASK7_REM_COUNTRY").Value = wsSource.Range("CP_TASK7_REM_COUNTRY").Value ' $M$199
    End If
    If .Range("CP_TASK7_REM_PCT").Value <> wsSource.Range("CP_TASK7_REM_PCT").Value Then
        .Range("CP_TASK7_REM_PCT").Value = wsSource.Range("CP_TASK7_REM_PCT").Value ' $N$199
    End If
    If .Range("CP_TASK8_OVD_JUST").Value <> wsSource.Range("CP_TASK8_OVD_JUST").Value Then
        .Range("CP_TASK8_OVD_JUST").Value = wsSource.Range("CP_TASK8_OVD_JUST").Value ' $I$200
    End If
    If .Range("CP_TASK8_OVD_QTY").Value <> wsSource.Range("CP_TASK8_OVD_QTY").Value Then
        .Range("CP_TASK8_OVD_QTY").Value = wsSource.Range("CP_TASK8_OVD_QTY").Value ' $H$200
    End If
    If .Range("CP_TASK8_REM_COUNTRY").Value <> wsSource.Range("CP_TASK8_REM_COUNTRY").Value Then
        .Range("CP_TASK8_REM_COUNTRY").Value = wsSource.Range("CP_TASK8_REM_COUNTRY").Value ' $M$200
    End If
    If .Range("CP_TASK8_REM_PCT").Value <> wsSource.Range("CP_TASK8_REM_PCT").Value Then
        .Range("CP_TASK8_REM_PCT").Value = wsSource.Range("CP_TASK8_REM_PCT").Value ' $N$200
    End If
    If .Range("CP_TASK9_OVD_JUST").Value <> wsSource.Range("CP_TASK9_OVD_JUST").Value Then
        .Range("CP_TASK9_OVD_JUST").Value = wsSource.Range("CP_TASK9_OVD_JUST").Value ' $I$201
    End If
    If .Range("CP_TASK9_OVD_QTY").Value <> wsSource.Range("CP_TASK9_OVD_QTY").Value Then
        .Range("CP_TASK9_OVD_QTY").Value = wsSource.Range("CP_TASK9_OVD_QTY").Value ' $H$201
    End If
    If .Range("CP_TASK9_REM_COUNTRY").Value <> wsSource.Range("CP_TASK9_REM_COUNTRY").Value Then
        .Range("CP_TASK9_REM_COUNTRY").Value = wsSource.Range("CP_TASK9_REM_COUNTRY").Value ' $M$201
    End If
    If .Range("CP_TASK9_REM_PCT").Value <> wsSource.Range("CP_TASK9_REM_PCT").Value Then
        .Range("CP_TASK9_REM_PCT").Value = wsSource.Range("CP_TASK9_REM_PCT").Value ' $N$201
    End If
    If .Range("DI_REM_COST").Value <> wsSource.Range("DI_REM_COST").Value Then
        .Range("DI_REM_COST").Value = wsSource.Range("CP_TASK9_REM_PCT").Value ' $N$201
    End If
    If .Range("DI_TASK1_OVD_JUST").Value <> wsSource.Range("DI_TASK1_OVD_JUST").Value Then
        .Range("DI_TASK1_OVD_JUST").Value = wsSource.Range("DI_TASK1_OVD_JUST").Value ' $I$215
    End If
    If .Range("DI_TASK1_OVD_QTY").Value <> wsSource.Range("DI_TASK1_OVD_QTY").Value Then
        .Range("DI_TASK1_OVD_QTY").Value = wsSource.Range("DI_TASK1_OVD_QTY").Value ' $H$215
    End If
    If .Range("DI_TASK1_REM_COUNTRY").Value <> wsSource.Range("DI_TASK1_REM_COUNTRY").Value Then
        .Range("DI_TASK1_REM_COUNTRY").Value = wsSource.Range("DI_TASK1_REM_COUNTRY").Value ' $M$215
    End If
    If .Range("DI_TASK1_REM_PCT").Value <> wsSource.Range("DI_TASK1_REM_PCT").Value Then
        .Range("DI_TASK1_REM_PCT").Value = wsSource.Range("DI_TASK1_REM_PCT").Value ' $N$215
    End If
    If .Range("DI_TASK10_OVD_JUST").Value <> wsSource.Range("DI_TASK10_OVD_JUST").Value Then
        .Range("DI_TASK10_OVD_JUST").Value = wsSource.Range("DI_TASK10_OVD_JUST").Value ' $I$224
    End If
    If .Range("DI_TASK10_OVD_QTY").Value <> wsSource.Range("DI_TASK10_OVD_QTY").Value Then
        .Range("DI_TASK10_OVD_QTY").Value = wsSource.Range("DI_TASK10_OVD_QTY").Value ' $H$224
    End If
    If .Range("DI_TASK10_REM_COUNTRY").Value <> wsSource.Range("DI_TASK10_REM_COUNTRY").Value Then
        .Range("DI_TASK10_REM_COUNTRY").Value = wsSource.Range("DI_TASK10_REM_COUNTRY").Value ' $M$224
    End If
    If .Range("DI_TASK10_REM_PCT").Value <> wsSource.Range("DI_TASK10_REM_PCT").Value Then
        .Range("DI_TASK10_REM_PCT").Value = wsSource.Range("DI_TASK10_REM_PCT").Value ' $N$224
    End If
    If .Range("DI_TASK2_OVD_JUST").Value <> wsSource.Range("DI_TASK2_OVD_JUST").Value Then
        .Range("DI_TASK2_OVD_JUST").Value = wsSource.Range("DI_TASK2_OVD_JUST").Value ' $I$216
    End If
    If .Range("DI_TASK2_OVD_QTY").Value <> wsSource.Range("DI_TASK2_OVD_QTY").Value Then
        .Range("DI_TASK2_OVD_QTY").Value = wsSource.Range("DI_TASK2_OVD_QTY").Value ' $H$216
    End If
    If .Range("DI_TASK2_REM_COUNTRY").Value <> wsSource.Range("DI_TASK2_REM_COUNTRY").Value Then
        .Range("DI_TASK2_REM_COUNTRY").Value = wsSource.Range("DI_TASK2_REM_COUNTRY").Value ' $M$216
    End If
    If .Range("DI_TASK2_REM_PCT").Value <> wsSource.Range("DI_TASK2_REM_PCT").Value Then
        .Range("DI_TASK2_REM_PCT").Value = wsSource.Range("DI_TASK2_REM_PCT").Value ' $N$216
    End If
    If .Range("DI_TASK3_OVD_JUST").Value <> wsSource.Range("DI_TASK3_OVD_JUST").Value Then
        .Range("DI_TASK3_OVD_JUST").Value = wsSource.Range("DI_TASK3_OVD_JUST").Value ' $I$217
    End If
    If .Range("DI_TASK3_OVD_QTY").Value <> wsSource.Range("DI_TASK3_OVD_QTY").Value Then
        .Range("DI_TASK3_OVD_QTY").Value = wsSource.Range("DI_TASK3_OVD_QTY").Value ' $H$217
    End If
    If .Range("DI_TASK3_REM_COUNTRY").Value <> wsSource.Range("DI_TASK3_REM_COUNTRY").Value Then
        .Range("DI_TASK3_REM_COUNTRY").Value = wsSource.Range("DI_TASK3_REM_COUNTRY").Value ' $M$217
    End If
    If .Range("DI_TASK3_REM_PCT").Value <> wsSource.Range("DI_TASK3_REM_PCT").Value Then
        .Range("DI_TASK3_REM_PCT").Value = wsSource.Range("DI_TASK3_REM_PCT").Value ' $N$217
    End If
    If .Range("DI_TASK4_OVD_JUST").Value <> wsSource.Range("DI_TASK4_OVD_JUST").Value Then
        .Range("DI_TASK4_OVD_JUST").Value = wsSource.Range("DI_TASK4_OVD_JUST").Value ' $I$218
    End If
    If .Range("DI_TASK4_OVD_QTY").Value <> wsSource.Range("DI_TASK4_OVD_QTY").Value Then
        .Range("DI_TASK4_OVD_QTY").Value = wsSource.Range("DI_TASK4_OVD_QTY").Value ' $H$218
    End If
    If .Range("DI_TASK4_REM_COUNTRY").Value <> wsSource.Range("DI_TASK4_REM_COUNTRY").Value Then
        .Range("DI_TASK4_REM_COUNTRY").Value = wsSource.Range("DI_TASK4_REM_COUNTRY").Value ' $M$218
    End If
    If .Range("DI_TASK4_REM_PCT").Value <> wsSource.Range("DI_TASK4_REM_PCT").Value Then
        .Range("DI_TASK4_REM_PCT").Value = wsSource.Range("DI_TASK4_REM_PCT").Value ' $N$218
    End If
    If .Range("DI_TASK5_OVD_JUST").Value <> wsSource.Range("DI_TASK5_OVD_JUST").Value Then
        .Range("DI_TASK5_OVD_JUST").Value = wsSource.Range("DI_TASK5_OVD_JUST").Value ' $I$219
    End If
    If .Range("DI_TASK5_OVD_QTY").Value <> wsSource.Range("DI_TASK5_OVD_QTY").Value Then
        .Range("DI_TASK5_OVD_QTY").Value = wsSource.Range("DI_TASK5_OVD_QTY").Value ' $H$219
    End If
    If .Range("DI_TASK5_REM_COUNTRY").Value <> wsSource.Range("DI_TASK5_REM_COUNTRY").Value Then
        .Range("DI_TASK5_REM_COUNTRY").Value = wsSource.Range("DI_TASK5_REM_COUNTRY").Value ' $M$219
    End If
    If .Range("DI_TASK5_REM_PCT").Value <> wsSource.Range("DI_TASK5_REM_PCT").Value Then
        .Range("DI_TASK5_REM_PCT").Value = wsSource.Range("DI_TASK5_REM_PCT").Value ' $N$219
    End If
    If .Range("DI_TASK6_OVD_JUST").Value <> wsSource.Range("DI_TASK6_OVD_JUST").Value Then
        .Range("DI_TASK6_OVD_JUST").Value = wsSource.Range("DI_TASK6_OVD_JUST").Value ' $I$220
    End If
    If .Range("DI_TASK6_OVD_QTY").Value <> wsSource.Range("DI_TASK6_OVD_QTY").Value Then
        .Range("DI_TASK6_OVD_QTY").Value = wsSource.Range("DI_TASK6_OVD_QTY").Value ' $H$220
    End If
    If .Range("DI_TASK6_REM_COUNTRY").Value <> wsSource.Range("DI_TASK6_REM_COUNTRY").Value Then
        .Range("DI_TASK6_REM_COUNTRY").Value = wsSource.Range("DI_TASK6_REM_COUNTRY").Value ' $M$220
    End If
    If .Range("DI_TASK6_REM_PCT").Value <> wsSource.Range("DI_TASK6_REM_PCT").Value Then
        .Range("DI_TASK6_REM_PCT").Value = wsSource.Range("DI_TASK6_REM_PCT").Value ' $N$220
    End If
    If .Range("DI_TASK7_OVD_JUST").Value <> wsSource.Range("DI_TASK7_OVD_JUST").Value Then
        .Range("DI_TASK7_OVD_JUST").Value = wsSource.Range("DI_TASK7_OVD_JUST").Value ' $I$221
    End If
    If .Range("DI_TASK7_OVD_QTY").Value <> wsSource.Range("DI_TASK7_OVD_QTY").Value Then
        .Range("DI_TASK7_OVD_QTY").Value = wsSource.Range("DI_TASK7_OVD_QTY").Value ' $H$221
    End If
    If .Range("DI_TASK7_REM_COUNTRY").Value <> wsSource.Range("DI_TASK7_REM_COUNTRY").Value Then
        .Range("DI_TASK7_REM_COUNTRY").Value = wsSource.Range("DI_TASK7_REM_COUNTRY").Value ' $M$221
    End If
    If .Range("DI_TASK7_REM_PCT").Value <> wsSource.Range("DI_TASK7_REM_PCT").Value Then
        .Range("DI_TASK7_REM_PCT").Value = wsSource.Range("DI_TASK7_REM_PCT").Value ' $N$221
    End If
    If .Range("DI_TASK8_OVD_JUST").Value <> wsSource.Range("DI_TASK8_OVD_JUST").Value Then
        .Range("DI_TASK8_OVD_JUST").Value = wsSource.Range("DI_TASK8_OVD_JUST").Value ' $I$222
    End If
    If .Range("DI_TASK8_OVD_QTY").Value <> wsSource.Range("DI_TASK8_OVD_QTY").Value Then
        .Range("DI_TASK8_OVD_QTY").Value = wsSource.Range("DI_TASK8_OVD_QTY").Value ' $H$222
    End If
    If .Range("DI_TASK8_REM_COUNTRY").Value <> wsSource.Range("DI_TASK8_REM_COUNTRY").Value Then
        .Range("DI_TASK8_REM_COUNTRY").Value = wsSource.Range("DI_TASK8_REM_COUNTRY").Value ' $M$222
    End If
    If .Range("DI_TASK8_REM_PCT").Value <> wsSource.Range("DI_TASK8_REM_PCT").Value Then
        .Range("DI_TASK8_REM_PCT").Value = wsSource.Range("DI_TASK8_REM_PCT").Value ' $N$222
    End If
    If .Range("DI_TASK9_OVD_JUST").Value <> wsSource.Range("DI_TASK9_OVD_JUST").Value Then
        .Range("DI_TASK9_OVD_JUST").Value = wsSource.Range("DI_TASK9_OVD_JUST").Value ' $I$223
    End If
    If .Range("DI_TASK9_OVD_QTY").Value <> wsSource.Range("DI_TASK9_OVD_QTY").Value Then
        .Range("DI_TASK9_OVD_QTY").Value = wsSource.Range("DI_TASK9_OVD_QTY").Value ' $H$223
    End If
    If .Range("DI_TASK9_REM_COUNTRY").Value <> wsSource.Range("DI_TASK9_REM_COUNTRY").Value Then
        .Range("DI_TASK9_REM_COUNTRY").Value = wsSource.Range("DI_TASK9_REM_COUNTRY").Value ' $M$223
    End If
    If .Range("DI_TASK9_REM_PCT").Value <> wsSource.Range("DI_TASK9_REM_PCT").Value Then
        .Range("DI_TASK9_REM_PCT").Value = wsSource.Range("DI_TASK9_REM_PCT").Value ' $N$223
    End If
    If .Range("DOC_REM_COST").Value <> wsSource.Range("DOC_REM_COST").Value Then
        .Range("DOC_REM_COST").Value = wsSource.Range("DI_TASK9_REM_PCT").Value ' $N$223
    End If
    If .Range("DOC_TASK1_OVD_JUST").Value <> wsSource.Range("DOC_TASK1_OVD_JUST").Value Then
        .Range("DOC_TASK1_OVD_JUST").Value = wsSource.Range("DOC_TASK1_OVD_JUST").Value ' $I$322
    End If
    If .Range("DOC_TASK1_OVD_QTY").Value <> wsSource.Range("DOC_TASK1_OVD_QTY").Value Then
        .Range("DOC_TASK1_OVD_QTY").Value = wsSource.Range("DOC_TASK1_OVD_QTY").Value ' $H$322
    End If
    If .Range("DOC_TASK1_REM_COUNTRY").Value <> wsSource.Range("DOC_TASK1_REM_COUNTRY").Value Then
        .Range("DOC_TASK1_REM_COUNTRY").Value = wsSource.Range("DOC_TASK1_REM_COUNTRY").Value ' $M$322
    End If
    If .Range("DOC_TASK1_REM_PCT").Value <> wsSource.Range("DOC_TASK1_REM_PCT").Value Then
        .Range("DOC_TASK1_REM_PCT").Value = wsSource.Range("DOC_TASK1_REM_PCT").Value ' $N$322
    End If
    If .Range("DOC_TASK10_OVD_JUST").Value <> wsSource.Range("DOC_TASK10_OVD_JUST").Value Then
        .Range("DOC_TASK10_OVD_JUST").Value = wsSource.Range("DOC_TASK10_OVD_JUST").Value ' $I$331
    End If
    If .Range("DOC_TASK10_OVD_QTY").Value <> wsSource.Range("DOC_TASK10_OVD_QTY").Value Then
        .Range("DOC_TASK10_OVD_QTY").Value = wsSource.Range("DOC_TASK10_OVD_QTY").Value ' $H$331
    End If
    If .Range("DOC_TASK10_REM_COUNTRY").Value <> wsSource.Range("DOC_TASK10_REM_COUNTRY").Value Then
        .Range("DOC_TASK10_REM_COUNTRY").Value = wsSource.Range("DOC_TASK10_REM_COUNTRY").Value ' $M$331
    End If
    If .Range("DOC_TASK10_REM_PCT").Value <> wsSource.Range("DOC_TASK10_REM_PCT").Value Then
        .Range("DOC_TASK10_REM_PCT").Value = wsSource.Range("DOC_TASK10_REM_PCT").Value ' $N$331
    End If
    If .Range("DOC_TASK2_OVD_JUST").Value <> wsSource.Range("DOC_TASK2_OVD_JUST").Value Then
        .Range("DOC_TASK2_OVD_JUST").Value = wsSource.Range("DOC_TASK2_OVD_JUST").Value ' $I$323
    End If
    If .Range("DOC_TASK2_OVD_QTY").Value <> wsSource.Range("DOC_TASK2_OVD_QTY").Value Then
        .Range("DOC_TASK2_OVD_QTY").Value = wsSource.Range("DOC_TASK2_OVD_QTY").Value ' $H$323
    End If
    If .Range("DOC_TASK2_REM_COUNTRY").Value <> wsSource.Range("DOC_TASK2_REM_COUNTRY").Value Then
        .Range("DOC_TASK2_REM_COUNTRY").Value = wsSource.Range("DOC_TASK2_REM_COUNTRY").Value ' $M$323
    End If
    If .Range("DOC_TASK2_REM_PCT").Value <> wsSource.Range("DOC_TASK2_REM_PCT").Value Then
        .Range("DOC_TASK2_REM_PCT").Value = wsSource.Range("DOC_TASK2_REM_PCT").Value ' $N$323
    End If
    If .Range("DOC_TASK3_OVD_JUST").Value <> wsSource.Range("DOC_TASK3_OVD_JUST").Value Then
        .Range("DOC_TASK3_OVD_JUST").Value = wsSource.Range("DOC_TASK3_OVD_JUST").Value ' $I$324
    End If
    If .Range("DOC_TASK3_OVD_QTY").Value <> wsSource.Range("DOC_TASK3_OVD_QTY").Value Then
        .Range("DOC_TASK3_OVD_QTY").Value = wsSource.Range("DOC_TASK3_OVD_QTY").Value ' $H$324
    End If
    If .Range("DOC_TASK3_REM_COUNTRY").Value <> wsSource.Range("DOC_TASK3_REM_COUNTRY").Value Then
        .Range("DOC_TASK3_REM_COUNTRY").Value = wsSource.Range("DOC_TASK3_REM_COUNTRY").Value ' $M$324
    End If
    If .Range("DOC_TASK3_REM_PCT").Value <> wsSource.Range("DOC_TASK3_REM_PCT").Value Then
        .Range("DOC_TASK3_REM_PCT").Value = wsSource.Range("DOC_TASK3_REM_PCT").Value ' $N$324
    End If
    If .Range("DOC_TASK4_OVD_JUST").Value <> wsSource.Range("DOC_TASK4_OVD_JUST").Value Then
        .Range("DOC_TASK4_OVD_JUST").Value = wsSource.Range("DOC_TASK4_OVD_JUST").Value ' $I$325
    End If
    If .Range("DOC_TASK4_OVD_QTY").Value <> wsSource.Range("DOC_TASK4_OVD_QTY").Value Then
        .Range("DOC_TASK4_OVD_QTY").Value = wsSource.Range("DOC_TASK4_OVD_QTY").Value ' $H$325
    End If
    If .Range("DOC_TASK4_REM_COUNTRY").Value <> wsSource.Range("DOC_TASK4_REM_COUNTRY").Value Then
        .Range("DOC_TASK4_REM_COUNTRY").Value = wsSource.Range("DOC_TASK4_REM_COUNTRY").Value ' $M$325
    End If
    If .Range("DOC_TASK4_REM_PCT").Value <> wsSource.Range("DOC_TASK4_REM_PCT").Value Then
        .Range("DOC_TASK4_REM_PCT").Value = wsSource.Range("DOC_TASK4_REM_PCT").Value ' $N$325
    End If
    If .Range("DOC_TASK5_OVD_JUST").Value <> wsSource.Range("DOC_TASK5_OVD_JUST").Value Then
        .Range("DOC_TASK5_OVD_JUST").Value = wsSource.Range("DOC_TASK5_OVD_JUST").Value ' $I$326
    End If
    If .Range("DOC_TASK5_OVD_QTY").Value <> wsSource.Range("DOC_TASK5_OVD_QTY").Value Then
        .Range("DOC_TASK5_OVD_QTY").Value = wsSource.Range("DOC_TASK5_OVD_QTY").Value ' $H$326
    End If
    If .Range("DOC_TASK5_REM_COUNTRY").Value <> wsSource.Range("DOC_TASK5_REM_COUNTRY").Value Then
        .Range("DOC_TASK5_REM_COUNTRY").Value = wsSource.Range("DOC_TASK5_REM_COUNTRY").Value ' $M$326
    End If
    If .Range("DOC_TASK5_REM_PCT").Value <> wsSource.Range("DOC_TASK5_REM_PCT").Value Then
        .Range("DOC_TASK5_REM_PCT").Value = wsSource.Range("DOC_TASK5_REM_PCT").Value ' $N$326
    End If
    If .Range("DOC_TASK6_OVD_JUST").Value <> wsSource.Range("DOC_TASK6_OVD_JUST").Value Then
        .Range("DOC_TASK6_OVD_JUST").Value = wsSource.Range("DOC_TASK6_OVD_JUST").Value ' $I$327
    End If
    If .Range("DOC_TASK6_OVD_QTY").Value <> wsSource.Range("DOC_TASK6_OVD_QTY").Value Then
        .Range("DOC_TASK6_OVD_QTY").Value = wsSource.Range("DOC_TASK6_OVD_QTY").Value ' $H$327
    End If
    If .Range("DOC_TASK6_REM_COUNTRY").Value <> wsSource.Range("DOC_TASK6_REM_COUNTRY").Value Then
        .Range("DOC_TASK6_REM_COUNTRY").Value = wsSource.Range("DOC_TASK6_REM_COUNTRY").Value ' $M$327
    End If
    If .Range("DOC_TASK6_REM_PCT").Value <> wsSource.Range("DOC_TASK6_REM_PCT").Value Then
        .Range("DOC_TASK6_REM_PCT").Value = wsSource.Range("DOC_TASK6_REM_PCT").Value ' $N$327
    End If
    If .Range("DOC_TASK7_OVD_JUST").Value <> wsSource.Range("DOC_TASK7_OVD_JUST").Value Then
        .Range("DOC_TASK7_OVD_JUST").Value = wsSource.Range("DOC_TASK7_OVD_JUST").Value ' $I$328
    End If
    If .Range("DOC_TASK7_OVD_QTY").Value <> wsSource.Range("DOC_TASK7_OVD_QTY").Value Then
        .Range("DOC_TASK7_OVD_QTY").Value = wsSource.Range("DOC_TASK7_OVD_QTY").Value ' $H$328
    End If
    If .Range("DOC_TASK7_REM_COUNTRY").Value <> wsSource.Range("DOC_TASK7_REM_COUNTRY").Value Then
        .Range("DOC_TASK7_REM_COUNTRY").Value = wsSource.Range("DOC_TASK7_REM_COUNTRY").Value ' $M$328
    End If
    If .Range("DOC_TASK7_REM_PCT").Value <> wsSource.Range("DOC_TASK7_REM_PCT").Value Then
        .Range("DOC_TASK7_REM_PCT").Value = wsSource.Range("DOC_TASK7_REM_PCT").Value ' $N$328
    End If
    If .Range("DOC_TASK8_OVD_JUST").Value <> wsSource.Range("DOC_TASK8_OVD_JUST").Value Then
        .Range("DOC_TASK8_OVD_JUST").Value = wsSource.Range("DOC_TASK8_OVD_JUST").Value ' $I$329
    End If
    If .Range("DOC_TASK8_OVD_QTY").Value <> wsSource.Range("DOC_TASK8_OVD_QTY").Value Then
        .Range("DOC_TASK8_OVD_QTY").Value = wsSource.Range("DOC_TASK8_OVD_QTY").Value ' $H$329
    End If
    If .Range("DOC_TASK8_REM_COUNTRY").Value <> wsSource.Range("DOC_TASK8_REM_COUNTRY").Value Then
        .Range("DOC_TASK8_REM_COUNTRY").Value = wsSource.Range("DOC_TASK8_REM_COUNTRY").Value ' $M$329
    End If
    If .Range("DOC_TASK8_REM_PCT").Value <> wsSource.Range("DOC_TASK8_REM_PCT").Value Then
        .Range("DOC_TASK8_REM_PCT").Value = wsSource.Range("DOC_TASK8_REM_PCT").Value ' $N$329
    End If
    If .Range("DOC_TASK9_OVD_JUST").Value <> wsSource.Range("DOC_TASK9_OVD_JUST").Value Then
        .Range("DOC_TASK9_OVD_JUST").Value = wsSource.Range("DOC_TASK9_OVD_JUST").Value ' $I$330
    End If
    If .Range("DOC_TASK9_OVD_QTY").Value <> wsSource.Range("DOC_TASK9_OVD_QTY").Value Then
        .Range("DOC_TASK9_OVD_QTY").Value = wsSource.Range("DOC_TASK9_OVD_QTY").Value ' $H$330
    End If
    If .Range("DOC_TASK9_REM_COUNTRY").Value <> wsSource.Range("DOC_TASK9_REM_COUNTRY").Value Then
        .Range("DOC_TASK9_REM_COUNTRY").Value = wsSource.Range("DOC_TASK9_REM_COUNTRY").Value ' $M$330
    End If
    If .Range("DOC_TASK9_REM_PCT").Value <> wsSource.Range("DOC_TASK9_REM_PCT").Value Then
        .Range("DOC_TASK9_REM_PCT").Value = wsSource.Range("DOC_TASK9_REM_PCT").Value ' $N$330
    End If
    If .Range("ESD_TASK1_OVD_JUST").Value <> wsSource.Range("ESD_TASK1_OVD_JUST").Value Then
        .Range("ESD_TASK1_OVD_JUST").Value = wsSource.Range("ESD_TASK1_OVD_JUST").Value ' $I$237
    End If
    If .Range("ESD_TASK1_OVD_QTY").Value <> wsSource.Range("ESD_TASK1_OVD_QTY").Value Then
        .Range("ESD_TASK1_OVD_QTY").Value = wsSource.Range("ESD_TASK1_OVD_QTY").Value ' $H$237
    End If
    If .Range("ESD_TASK1_REM_COUNTRY").Value <> wsSource.Range("ESD_TASK1_REM_COUNTRY").Value Then
        .Range("ESD_TASK1_REM_COUNTRY").Value = wsSource.Range("ESD_TASK1_REM_COUNTRY").Value ' $M$237
    End If
    If .Range("ESD_TASK1_REM_PCT").Value <> wsSource.Range("ESD_TASK1_REM_PCT").Value Then
        .Range("ESD_TASK1_REM_PCT").Value = wsSource.Range("ESD_TASK1_REM_PCT").Value ' $N$237
    End If
    If .Range("ESD_TASK10_OVD_JUST").Value <> wsSource.Range("ESD_TASK10_OVD_JUST").Value Then
        .Range("ESD_TASK10_OVD_JUST").Value = wsSource.Range("ESD_TASK10_OVD_JUST").Value ' $I$246
    End If
    If .Range("ESD_TASK10_OVD_QTY").Value <> wsSource.Range("ESD_TASK10_OVD_QTY").Value Then
        .Range("ESD_TASK10_OVD_QTY").Value = wsSource.Range("ESD_TASK10_OVD_QTY").Value ' $H$246
    End If
    If .Range("ESD_TASK10_REM_COUNTRY").Value <> wsSource.Range("ESD_TASK10_REM_COUNTRY").Value Then
        .Range("ESD_TASK10_REM_COUNTRY").Value = wsSource.Range("ESD_TASK10_REM_COUNTRY").Value ' $M$246
    End If
    If .Range("ESD_TASK10_REM_PCT").Value <> wsSource.Range("ESD_TASK10_REM_PCT").Value Then
        .Range("ESD_TASK10_REM_PCT").Value = wsSource.Range("ESD_TASK10_REM_PCT").Value ' $N$246
    End If
    If .Range("ESD_TASK2_OVD_JUST").Value <> wsSource.Range("ESD_TASK2_OVD_JUST").Value Then
        .Range("ESD_TASK2_OVD_JUST").Value = wsSource.Range("ESD_TASK2_OVD_JUST").Value ' $I$238
    End If
    If .Range("ESD_TASK2_OVD_QTY").Value <> wsSource.Range("ESD_TASK2_OVD_QTY").Value Then
        .Range("ESD_TASK2_OVD_QTY").Value = wsSource.Range("ESD_TASK2_OVD_QTY").Value ' $H$238
    End If
    If .Range("ESD_TASK2_REM_COUNTRY").Value <> wsSource.Range("ESD_TASK2_REM_COUNTRY").Value Then
        .Range("ESD_TASK2_REM_COUNTRY").Value = wsSource.Range("ESD_TASK2_REM_COUNTRY").Value ' $M$238
    End If
    If .Range("ESD_TASK2_REM_PCT").Value <> wsSource.Range("ESD_TASK2_REM_PCT").Value Then
        .Range("ESD_TASK2_REM_PCT").Value = wsSource.Range("ESD_TASK2_REM_PCT").Value ' $N$238
    End If
    If .Range("ESD_TASK3_OVD_JUST").Value <> wsSource.Range("ESD_TASK3_OVD_JUST").Value Then
        .Range("ESD_TASK3_OVD_JUST").Value = wsSource.Range("ESD_TASK3_OVD_JUST").Value ' $I$239
    End If
    If .Range("ESD_TASK3_OVD_QTY").Value <> wsSource.Range("ESD_TASK3_OVD_QTY").Value Then
        .Range("ESD_TASK3_OVD_QTY").Value = wsSource.Range("ESD_TASK3_OVD_QTY").Value ' $H$239
    End If
    If .Range("ESD_TASK3_REM_COUNTRY").Value <> wsSource.Range("ESD_TASK3_REM_COUNTRY").Value Then
        .Range("ESD_TASK3_REM_COUNTRY").Value = wsSource.Range("ESD_TASK3_REM_COUNTRY").Value ' $M$239
    End If
    If .Range("ESD_TASK3_REM_PCT").Value <> wsSource.Range("ESD_TASK3_REM_PCT").Value Then
        .Range("ESD_TASK3_REM_PCT").Value = wsSource.Range("ESD_TASK3_REM_PCT").Value ' $N$239
    End If
    If .Range("ESD_TASK4_OVD_JUST").Value <> wsSource.Range("ESD_TASK4_OVD_JUST").Value Then
        .Range("ESD_TASK4_OVD_JUST").Value = wsSource.Range("ESD_TASK4_OVD_JUST").Value ' $I$240
    End If
    If .Range("ESD_TASK4_OVD_QTY").Value <> wsSource.Range("ESD_TASK4_OVD_QTY").Value Then
        .Range("ESD_TASK4_OVD_QTY").Value = wsSource.Range("ESD_TASK4_OVD_QTY").Value ' $H$240
    End If
    If .Range("ESD_TASK4_REM_COUNTRY").Value <> wsSource.Range("ESD_TASK4_REM_COUNTRY").Value Then
        .Range("ESD_TASK4_REM_COUNTRY").Value = wsSource.Range("ESD_TASK4_REM_COUNTRY").Value ' $M$240
    End If
    If .Range("ESD_TASK4_REM_PCT").Value <> wsSource.Range("ESD_TASK4_REM_PCT").Value Then
        .Range("ESD_TASK4_REM_PCT").Value = wsSource.Range("ESD_TASK4_REM_PCT").Value ' $N$240
    End If
    If .Range("ESD_TASK5_OVD_JUST").Value <> wsSource.Range("ESD_TASK5_OVD_JUST").Value Then
        .Range("ESD_TASK5_OVD_JUST").Value = wsSource.Range("ESD_TASK5_OVD_JUST").Value ' $I$241
    End If
    If .Range("ESD_TASK5_OVD_QTY").Value <> wsSource.Range("ESD_TASK5_OVD_QTY").Value Then
        .Range("ESD_TASK5_OVD_QTY").Value = wsSource.Range("ESD_TASK5_OVD_QTY").Value ' $H$241
    End If
    If .Range("ESD_TASK5_REM_COUNTRY").Value <> wsSource.Range("ESD_TASK5_REM_COUNTRY").Value Then
        .Range("ESD_TASK5_REM_COUNTRY").Value = wsSource.Range("ESD_TASK5_REM_COUNTRY").Value ' $M$241
    End If
    If .Range("ESD_TASK5_REM_PCT").Value <> wsSource.Range("ESD_TASK5_REM_PCT").Value Then
        .Range("ESD_TASK5_REM_PCT").Value = wsSource.Range("ESD_TASK5_REM_PCT").Value ' $N$241
    End If
    If .Range("ESD_TASK6_OVD_JUST").Value <> wsSource.Range("ESD_TASK6_OVD_JUST").Value Then
        .Range("ESD_TASK6_OVD_JUST").Value = wsSource.Range("ESD_TASK6_OVD_JUST").Value ' $I$242
    End If
    If .Range("ESD_TASK6_OVD_QTY").Value <> wsSource.Range("ESD_TASK6_OVD_QTY").Value Then
        .Range("ESD_TASK6_OVD_QTY").Value = wsSource.Range("ESD_TASK6_OVD_QTY").Value ' $H$242
    End If
    If .Range("ESD_TASK6_REM_COUNTRY").Value <> wsSource.Range("ESD_TASK6_REM_COUNTRY").Value Then
        .Range("ESD_TASK6_REM_COUNTRY").Value = wsSource.Range("ESD_TASK6_REM_COUNTRY").Value ' $M$242
    End If
    If .Range("ESD_TASK6_REM_PCT").Value <> wsSource.Range("ESD_TASK6_REM_PCT").Value Then
        .Range("ESD_TASK6_REM_PCT").Value = wsSource.Range("ESD_TASK6_REM_PCT").Value ' $N$242
    End If
    If .Range("ESD_TASK7_OVD_JUST").Value <> wsSource.Range("ESD_TASK7_OVD_JUST").Value Then
        .Range("ESD_TASK7_OVD_JUST").Value = wsSource.Range("ESD_TASK7_OVD_JUST").Value ' $I$243
    End If
    If .Range("ESD_TASK7_OVD_QTY").Value <> wsSource.Range("ESD_TASK7_OVD_QTY").Value Then
        .Range("ESD_TASK7_OVD_QTY").Value = wsSource.Range("ESD_TASK7_OVD_QTY").Value ' $H$243
    End If
    If .Range("ESD_TASK7_REM_COUNTRY").Value <> wsSource.Range("ESD_TASK7_REM_COUNTRY").Value Then
        .Range("ESD_TASK7_REM_COUNTRY").Value = wsSource.Range("ESD_TASK7_REM_COUNTRY").Value ' $M$243
    End If
    If .Range("ESD_TASK7_REM_PCT").Value <> wsSource.Range("ESD_TASK7_REM_PCT").Value Then
        .Range("ESD_TASK7_REM_PCT").Value = wsSource.Range("ESD_TASK7_REM_PCT").Value ' $N$243
    End If
    If .Range("ESD_TASK8_OVD_JUST").Value <> wsSource.Range("ESD_TASK8_OVD_JUST").Value Then
        .Range("ESD_TASK8_OVD_JUST").Value = wsSource.Range("ESD_TASK8_OVD_JUST").Value ' $I$244
    End If
    If .Range("ESD_TASK8_OVD_QTY").Value <> wsSource.Range("ESD_TASK8_OVD_QTY").Value Then
        .Range("ESD_TASK8_OVD_QTY").Value = wsSource.Range("ESD_TASK8_OVD_QTY").Value ' $H$244
    End If
    If .Range("ESD_TASK8_REM_COUNTRY").Value <> wsSource.Range("ESD_TASK8_REM_COUNTRY").Value Then
        .Range("ESD_TASK8_REM_COUNTRY").Value = wsSource.Range("ESD_TASK8_REM_COUNTRY").Value ' $M$244
    End If
    If .Range("ESD_TASK8_REM_PCT").Value <> wsSource.Range("ESD_TASK8_REM_PCT").Value Then
        .Range("ESD_TASK8_REM_PCT").Value = wsSource.Range("ESD_TASK8_REM_PCT").Value ' $N$244
    End If
    If .Range("ESD_TASK9_OVD_JUST").Value <> wsSource.Range("ESD_TASK9_OVD_JUST").Value Then
        .Range("ESD_TASK9_OVD_JUST").Value = wsSource.Range("ESD_TASK9_OVD_JUST").Value ' $I$245
    End If
    If .Range("ESD_TASK9_OVD_QTY").Value <> wsSource.Range("ESD_TASK9_OVD_QTY").Value Then
        .Range("ESD_TASK9_OVD_QTY").Value = wsSource.Range("ESD_TASK9_OVD_QTY").Value ' $H$245
    End If
    If .Range("ESD_TASK9_REM_COUNTRY").Value <> wsSource.Range("ESD_TASK9_REM_COUNTRY").Value Then
        .Range("ESD_TASK9_REM_COUNTRY").Value = wsSource.Range("ESD_TASK9_REM_COUNTRY").Value ' $M$245
    End If
    If .Range("ESD_TASK9_REM_PCT").Value <> wsSource.Range("ESD_TASK9_REM_PCT").Value Then
        .Range("ESD_TASK9_REM_PCT").Value = wsSource.Range("ESD_TASK9_REM_PCT").Value ' $N$245
    End If
    If .Range("HMI_TASK1_OVD_JUST").Value <> wsSource.Range("HMI_TASK1_OVD_JUST").Value Then
        .Range("HMI_TASK1_OVD_JUST").Value = wsSource.Range("HMI_TASK1_OVD_JUST").Value ' $I$171
    End If
    If .Range("HMI_TASK1_OVD_QTY").Value <> wsSource.Range("HMI_TASK1_OVD_QTY").Value Then
        .Range("HMI_TASK1_OVD_QTY").Value = wsSource.Range("HMI_TASK1_OVD_QTY").Value ' $H$171
    End If
    If .Range("HMI_TASK1_REM_COUNTRY").Value <> wsSource.Range("HMI_TASK1_REM_COUNTRY").Value Then
        .Range("HMI_TASK1_REM_COUNTRY").Value = wsSource.Range("HMI_TASK1_REM_COUNTRY").Value ' $M$171
    End If
    If .Range("HMI_TASK1_REM_PCT").Value <> wsSource.Range("HMI_TASK1_REM_PCT").Value Then
        .Range("HMI_TASK1_REM_PCT").Value = wsSource.Range("HMI_TASK1_REM_PCT").Value ' $N$171
    End If
    If .Range("HMI_TASK10_OVD_JUST").Value <> wsSource.Range("HMI_TASK10_OVD_JUST").Value Then
        .Range("HMI_TASK10_OVD_JUST").Value = wsSource.Range("HMI_TASK10_OVD_JUST").Value ' $I$180
    End If
    If .Range("HMI_TASK10_OVD_QTY").Value <> wsSource.Range("HMI_TASK10_OVD_QTY").Value Then
        .Range("HMI_TASK10_OVD_QTY").Value = wsSource.Range("HMI_TASK10_OVD_QTY").Value ' $H$180
    End If
    If .Range("HMI_TASK10_REM_COUNTRY").Value <> wsSource.Range("HMI_TASK10_REM_COUNTRY").Value Then
        .Range("HMI_TASK10_REM_COUNTRY").Value = wsSource.Range("HMI_TASK10_REM_COUNTRY").Value ' $M$180
    End If
    If .Range("HMI_TASK10_REM_PCT").Value <> wsSource.Range("HMI_TASK10_REM_PCT").Value Then
        .Range("HMI_TASK10_REM_PCT").Value = wsSource.Range("HMI_TASK10_REM_PCT").Value ' $N$180
    End If
    If .Range("HMI_TASK2_OVD_JUST").Value <> wsSource.Range("HMI_TASK2_OVD_JUST").Value Then
        .Range("HMI_TASK2_OVD_JUST").Value = wsSource.Range("HMI_TASK2_OVD_JUST").Value ' $I$172
    End If
    If .Range("HMI_TASK2_OVD_QTY").Value <> wsSource.Range("HMI_TASK2_OVD_QTY").Value Then
        .Range("HMI_TASK2_OVD_QTY").Value = wsSource.Range("HMI_TASK2_OVD_QTY").Value ' $H$172
    End If
    If .Range("HMI_TASK2_REM_COUNTRY").Value <> wsSource.Range("HMI_TASK2_REM_COUNTRY").Value Then
        .Range("HMI_TASK2_REM_COUNTRY").Value = wsSource.Range("HMI_TASK2_REM_COUNTRY").Value ' $M$172
    End If
    If .Range("HMI_TASK2_REM_PCT").Value <> wsSource.Range("HMI_TASK2_REM_PCT").Value Then
        .Range("HMI_TASK2_REM_PCT").Value = wsSource.Range("HMI_TASK2_REM_PCT").Value ' $N$172
    End If
    If .Range("HMI_TASK3_OVD_JUST").Value <> wsSource.Range("HMI_TASK3_OVD_JUST").Value Then
        .Range("HMI_TASK3_OVD_JUST").Value = wsSource.Range("HMI_TASK3_OVD_JUST").Value ' $I$173
    End If
    If .Range("HMI_TASK3_OVD_QTY").Value <> wsSource.Range("HMI_TASK3_OVD_QTY").Value Then
        .Range("HMI_TASK3_OVD_QTY").Value = wsSource.Range("HMI_TASK3_OVD_QTY").Value ' $H$173
    End If
    If .Range("HMI_TASK3_REM_COUNTRY").Value <> wsSource.Range("HMI_TASK3_REM_COUNTRY").Value Then
        .Range("HMI_TASK3_REM_COUNTRY").Value = wsSource.Range("HMI_TASK3_REM_COUNTRY").Value ' $M$173
    End If
    If .Range("HMI_TASK3_REM_PCT").Value <> wsSource.Range("HMI_TASK3_REM_PCT").Value Then
        .Range("HMI_TASK3_REM_PCT").Value = wsSource.Range("HMI_TASK3_REM_PCT").Value ' $N$173
    End If
    If .Range("HMI_TASK4_OVD_JUST").Value <> wsSource.Range("HMI_TASK4_OVD_JUST").Value Then
        .Range("HMI_TASK4_OVD_JUST").Value = wsSource.Range("HMI_TASK4_OVD_JUST").Value ' $I$174
    End If
    If .Range("HMI_TASK4_OVD_QTY").Value <> wsSource.Range("HMI_TASK4_OVD_QTY").Value Then
        .Range("HMI_TASK4_OVD_QTY").Value = wsSource.Range("HMI_TASK4_OVD_QTY").Value ' $H$174
    End If
    If .Range("HMI_TASK4_REM_COUNTRY").Value <> wsSource.Range("HMI_TASK4_REM_COUNTRY").Value Then
        .Range("HMI_TASK4_REM_COUNTRY").Value = wsSource.Range("HMI_TASK4_REM_COUNTRY").Value ' $M$174
    End If
    If .Range("HMI_TASK4_REM_PCT").Value <> wsSource.Range("HMI_TASK4_REM_PCT").Value Then
        .Range("HMI_TASK4_REM_PCT").Value = wsSource.Range("HMI_TASK4_REM_PCT").Value ' $N$174
    End If
    If .Range("HMI_TASK5_OVD_JUST").Value <> wsSource.Range("HMI_TASK5_OVD_JUST").Value Then
        .Range("HMI_TASK5_OVD_JUST").Value = wsSource.Range("HMI_TASK5_OVD_JUST").Value ' $I$175
    End If
    If .Range("HMI_TASK5_OVD_QTY").Value <> wsSource.Range("HMI_TASK5_OVD_QTY").Value Then
        .Range("HMI_TASK5_OVD_QTY").Value = wsSource.Range("HMI_TASK5_OVD_QTY").Value ' $H$175
    End If
    If .Range("HMI_TASK5_REM_COUNTRY").Value <> wsSource.Range("HMI_TASK5_REM_COUNTRY").Value Then
        .Range("HMI_TASK5_REM_COUNTRY").Value = wsSource.Range("HMI_TASK5_REM_COUNTRY").Value ' $M$175
    End If
    If .Range("HMI_TASK5_REM_PCT").Value <> wsSource.Range("HMI_TASK5_REM_PCT").Value Then
        .Range("HMI_TASK5_REM_PCT").Value = wsSource.Range("HMI_TASK5_REM_PCT").Value ' $N$175
    End If
    If .Range("HMI_TASK6_OVD_JUST").Value <> wsSource.Range("HMI_TASK6_OVD_JUST").Value Then
        .Range("HMI_TASK6_OVD_JUST").Value = wsSource.Range("HMI_TASK6_OVD_JUST").Value ' $I$176
    End If
    If .Range("HMI_TASK6_OVD_QTY").Value <> wsSource.Range("HMI_TASK6_OVD_QTY").Value Then
        .Range("HMI_TASK6_OVD_QTY").Value = wsSource.Range("HMI_TASK6_OVD_QTY").Value ' $H$176
    End If
    If .Range("HMI_TASK6_REM_COUNTRY").Value <> wsSource.Range("HMI_TASK6_REM_COUNTRY").Value Then
        .Range("HMI_TASK6_REM_COUNTRY").Value = wsSource.Range("HMI_TASK6_REM_COUNTRY").Value ' $M$176
    End If
    If .Range("HMI_TASK6_REM_PCT").Value <> wsSource.Range("HMI_TASK6_REM_PCT").Value Then
        .Range("HMI_TASK6_REM_PCT").Value = wsSource.Range("HMI_TASK6_REM_PCT").Value ' $N$176
    End If
    If .Range("HMI_TASK7_OVD_JUST").Value <> wsSource.Range("HMI_TASK7_OVD_JUST").Value Then
        .Range("HMI_TASK7_OVD_JUST").Value = wsSource.Range("HMI_TASK7_OVD_JUST").Value ' $I$177
    End If
    If .Range("HMI_TASK7_OVD_QTY").Value <> wsSource.Range("HMI_TASK7_OVD_QTY").Value Then
        .Range("HMI_TASK7_OVD_QTY").Value = wsSource.Range("HMI_TASK7_OVD_QTY").Value ' $H$177
    End If
    If .Range("HMI_TASK7_REM_COUNTRY").Value <> wsSource.Range("HMI_TASK7_REM_COUNTRY").Value Then
        .Range("HMI_TASK7_REM_COUNTRY").Value = wsSource.Range("HMI_TASK7_REM_COUNTRY").Value ' $M$177
    End If
    If .Range("HMI_TASK7_REM_PCT").Value <> wsSource.Range("HMI_TASK7_REM_PCT").Value Then
        .Range("HMI_TASK7_REM_PCT").Value = wsSource.Range("HMI_TASK7_REM_PCT").Value ' $N$177
    End If
    If .Range("HMI_TASK8_OVD_JUST").Value <> wsSource.Range("HMI_TASK8_OVD_JUST").Value Then
        .Range("HMI_TASK8_OVD_JUST").Value = wsSource.Range("HMI_TASK8_OVD_JUST").Value ' $I$178
    End If
    If .Range("HMI_TASK8_OVD_QTY").Value <> wsSource.Range("HMI_TASK8_OVD_QTY").Value Then
        .Range("HMI_TASK8_OVD_QTY").Value = wsSource.Range("HMI_TASK8_OVD_QTY").Value ' $H$178
    End If
    If .Range("HMI_TASK8_REM_COUNTRY").Value <> wsSource.Range("HMI_TASK8_REM_COUNTRY").Value Then
        .Range("HMI_TASK8_REM_COUNTRY").Value = wsSource.Range("HMI_TASK8_REM_COUNTRY").Value ' $M$178
    End If
    If .Range("HMI_TASK8_REM_PCT").Value <> wsSource.Range("HMI_TASK8_REM_PCT").Value Then
        .Range("HMI_TASK8_REM_PCT").Value = wsSource.Range("HMI_TASK8_REM_PCT").Value ' $N$178
    End If
    If .Range("HMI_TASK9_OVD_JUST").Value <> wsSource.Range("HMI_TASK9_OVD_JUST").Value Then
        .Range("HMI_TASK9_OVD_JUST").Value = wsSource.Range("HMI_TASK9_OVD_JUST").Value ' $I$179
    End If
    If .Range("HMI_TASK9_OVD_QTY").Value <> wsSource.Range("HMI_TASK9_OVD_QTY").Value Then
        .Range("HMI_TASK9_OVD_QTY").Value = wsSource.Range("HMI_TASK9_OVD_QTY").Value ' $H$179
    End If
    If .Range("HMI_TASK9_REM_COUNTRY").Value <> wsSource.Range("HMI_TASK9_REM_COUNTRY").Value Then
        .Range("HMI_TASK9_REM_COUNTRY").Value = wsSource.Range("HMI_TASK9_REM_COUNTRY").Value ' $M$179
    End If
    If .Range("HMI_TASK9_REM_PCT").Value <> wsSource.Range("HMI_TASK9_REM_PCT").Value Then
        .Range("HMI_TASK9_REM_PCT").Value = wsSource.Range("HMI_TASK9_REM_PCT").Value ' $N$179
    End If
    If .Range("MEETING_TASK1_OVD_JUST").Value <> wsSource.Range("MEETING_TASK1_OVD_JUST").Value Then
        .Range("MEETING_TASK1_OVD_JUST").Value = wsSource.Range("MEETING_TASK1_OVD_JUST").Value ' $I$385
    End If
    If .Range("MEETING_TASK1_OVD_QTY").Value <> wsSource.Range("MEETING_TASK1_OVD_QTY").Value Then
        .Range("MEETING_TASK1_OVD_QTY").Value = wsSource.Range("MEETING_TASK1_OVD_QTY").Value ' $H$385
    End If
    If .Range("MEETING_TASK1_REM_COUNTRY").Value <> wsSource.Range("MEETING_TASK1_REM_COUNTRY").Value Then
        .Range("MEETING_TASK1_REM_COUNTRY").Value = wsSource.Range("MEETING_TASK1_REM_COUNTRY").Value ' $M$385
    End If
    If .Range("MEETING_TASK1_REM_PCT").Value <> wsSource.Range("MEETING_TASK1_REM_PCT").Value Then
        .Range("MEETING_TASK1_REM_PCT").Value = wsSource.Range("MEETING_TASK1_REM_PCT").Value ' $N$385
    End If
    If .Range("MEETING_TASK10_OVD_JUST").Value <> wsSource.Range("MEETING_TASK10_OVD_JUST").Value Then
        .Range("MEETING_TASK10_OVD_JUST").Value = wsSource.Range("MEETING_TASK10_OVD_JUST").Value ' $I$394
    End If
    If .Range("MEETING_TASK10_OVD_QTY").Value <> wsSource.Range("MEETING_TASK10_OVD_QTY").Value Then
        .Range("MEETING_TASK10_OVD_QTY").Value = wsSource.Range("MEETING_TASK10_OVD_QTY").Value ' $H$394
    End If
    If .Range("MEETING_TASK10_REM_COUNTRY").Value <> wsSource.Range("MEETING_TASK10_REM_COUNTRY").Value Then
        .Range("MEETING_TASK10_REM_COUNTRY").Value = wsSource.Range("MEETING_TASK10_REM_COUNTRY").Value ' $M$394
    End If
    If .Range("MEETING_TASK10_REM_PCT").Value <> wsSource.Range("MEETING_TASK10_REM_PCT").Value Then
        .Range("MEETING_TASK10_REM_PCT").Value = wsSource.Range("MEETING_TASK10_REM_PCT").Value ' $N$394
    End If
    If .Range("MEETING_TASK2_OVD_JUST").Value <> wsSource.Range("MEETING_TASK2_OVD_JUST").Value Then
        .Range("MEETING_TASK2_OVD_JUST").Value = wsSource.Range("MEETING_TASK2_OVD_JUST").Value ' $I$386
    End If
    If .Range("MEETING_TASK2_OVD_QTY").Value <> wsSource.Range("MEETING_TASK2_OVD_QTY").Value Then
        .Range("MEETING_TASK2_OVD_QTY").Value = wsSource.Range("MEETING_TASK2_OVD_QTY").Value ' $H$386
    End If
    If .Range("MEETING_TASK2_REM_COUNTRY").Value <> wsSource.Range("MEETING_TASK2_REM_COUNTRY").Value Then
        .Range("MEETING_TASK2_REM_COUNTRY").Value = wsSource.Range("MEETING_TASK2_REM_COUNTRY").Value ' $M$386
    End If
    If .Range("MEETING_TASK2_REM_PCT").Value <> wsSource.Range("MEETING_TASK2_REM_PCT").Value Then
        .Range("MEETING_TASK2_REM_PCT").Value = wsSource.Range("MEETING_TASK2_REM_PCT").Value ' $N$386
    End If
    If .Range("MEETING_TASK3_OVD_JUST").Value <> wsSource.Range("MEETING_TASK3_OVD_JUST").Value Then
        .Range("MEETING_TASK3_OVD_JUST").Value = wsSource.Range("MEETING_TASK3_OVD_JUST").Value ' $I$387
    End If
    If .Range("MEETING_TASK3_OVD_QTY").Value <> wsSource.Range("MEETING_TASK3_OVD_QTY").Value Then
        .Range("MEETING_TASK3_OVD_QTY").Value = wsSource.Range("MEETING_TASK3_OVD_QTY").Value ' $H$387
    End If
    If .Range("MEETING_TASK3_REM_COUNTRY").Value <> wsSource.Range("MEETING_TASK3_REM_COUNTRY").Value Then
        .Range("MEETING_TASK3_REM_COUNTRY").Value = wsSource.Range("MEETING_TASK3_REM_COUNTRY").Value ' $M$387
    End If
    If .Range("MEETING_TASK3_REM_PCT").Value <> wsSource.Range("MEETING_TASK3_REM_PCT").Value Then
        .Range("MEETING_TASK3_REM_PCT").Value = wsSource.Range("MEETING_TASK3_REM_PCT").Value ' $N$387
    End If
    If .Range("MEETING_TASK4_OVD_JUST").Value <> wsSource.Range("MEETING_TASK4_OVD_JUST").Value Then
        .Range("MEETING_TASK4_OVD_JUST").Value = wsSource.Range("MEETING_TASK4_OVD_JUST").Value ' $I$388
    End If
    If .Range("MEETING_TASK4_OVD_QTY").Value <> wsSource.Range("MEETING_TASK4_OVD_QTY").Value Then
        .Range("MEETING_TASK4_OVD_QTY").Value = wsSource.Range("MEETING_TASK4_OVD_QTY").Value ' $H$388
    End If
    If .Range("MEETING_TASK4_REM_COUNTRY").Value <> wsSource.Range("MEETING_TASK4_REM_COUNTRY").Value Then
        .Range("MEETING_TASK4_REM_COUNTRY").Value = wsSource.Range("MEETING_TASK4_REM_COUNTRY").Value ' $M$388
    End If
    If .Range("MEETING_TASK4_REM_PCT").Value <> wsSource.Range("MEETING_TASK4_REM_PCT").Value Then
        .Range("MEETING_TASK4_REM_PCT").Value = wsSource.Range("MEETING_TASK4_REM_PCT").Value ' $N$388
    End If
    If .Range("MEETING_TASK5_OVD_JUST").Value <> wsSource.Range("MEETING_TASK5_OVD_JUST").Value Then
        .Range("MEETING_TASK5_OVD_JUST").Value = wsSource.Range("MEETING_TASK5_OVD_JUST").Value ' $I$389
    End If
    If .Range("MEETING_TASK5_OVD_QTY").Value <> wsSource.Range("MEETING_TASK5_OVD_QTY").Value Then
        .Range("MEETING_TASK5_OVD_QTY").Value = wsSource.Range("MEETING_TASK5_OVD_QTY").Value ' $H$389
    End If
    If .Range("MEETING_TASK5_REM_COUNTRY").Value <> wsSource.Range("MEETING_TASK5_REM_COUNTRY").Value Then
        .Range("MEETING_TASK5_REM_COUNTRY").Value = wsSource.Range("MEETING_TASK5_REM_COUNTRY").Value ' $M$389
    End If
    If .Range("MEETING_TASK5_REM_PCT").Value <> wsSource.Range("MEETING_TASK5_REM_PCT").Value Then
        .Range("MEETING_TASK5_REM_PCT").Value = wsSource.Range("MEETING_TASK5_REM_PCT").Value ' $N$389
    End If
    If .Range("MEETING_TASK6_OVD_JUST").Value <> wsSource.Range("MEETING_TASK6_OVD_JUST").Value Then
        .Range("MEETING_TASK6_OVD_JUST").Value = wsSource.Range("MEETING_TASK6_OVD_JUST").Value ' $I$390
    End If
    If .Range("MEETING_TASK6_OVD_QTY").Value <> wsSource.Range("MEETING_TASK6_OVD_QTY").Value Then
        .Range("MEETING_TASK6_OVD_QTY").Value = wsSource.Range("MEETING_TASK6_OVD_QTY").Value ' $H$390
    End If
    If .Range("MEETING_TASK6_REM_COUNTRY").Value <> wsSource.Range("MEETING_TASK6_REM_COUNTRY").Value Then
        .Range("MEETING_TASK6_REM_COUNTRY").Value = wsSource.Range("MEETING_TASK6_REM_COUNTRY").Value ' $M$390
    End If
    If .Range("MEETING_TASK6_REM_PCT").Value <> wsSource.Range("MEETING_TASK6_REM_PCT").Value Then
        .Range("MEETING_TASK6_REM_PCT").Value = wsSource.Range("MEETING_TASK6_REM_PCT").Value ' $N$390
    End If
    If .Range("MEETING_TASK7_OVD_JUST").Value <> wsSource.Range("MEETING_TASK7_OVD_JUST").Value Then
        .Range("MEETING_TASK7_OVD_JUST").Value = wsSource.Range("MEETING_TASK7_OVD_JUST").Value ' $I$391
    End If
    If .Range("MEETING_TASK7_OVD_QTY").Value <> wsSource.Range("MEETING_TASK7_OVD_QTY").Value Then
        .Range("MEETING_TASK7_OVD_QTY").Value = wsSource.Range("MEETING_TASK7_OVD_QTY").Value ' $H$391
    End If
    If .Range("MEETING_TASK7_REM_COUNTRY").Value <> wsSource.Range("MEETING_TASK7_REM_COUNTRY").Value Then
        .Range("MEETING_TASK7_REM_COUNTRY").Value = wsSource.Range("MEETING_TASK7_REM_COUNTRY").Value ' $M$391
    End If
    If .Range("MEETING_TASK7_REM_PCT").Value <> wsSource.Range("MEETING_TASK7_REM_PCT").Value Then
        .Range("MEETING_TASK7_REM_PCT").Value = wsSource.Range("MEETING_TASK7_REM_PCT").Value ' $N$391
    End If
    If .Range("MEETING_TASK8_OVD_JUST").Value <> wsSource.Range("MEETING_TASK8_OVD_JUST").Value Then
        .Range("MEETING_TASK8_OVD_JUST").Value = wsSource.Range("MEETING_TASK8_OVD_JUST").Value ' $I$392
    End If
    If .Range("MEETING_TASK8_OVD_QTY").Value <> wsSource.Range("MEETING_TASK8_OVD_QTY").Value Then
        .Range("MEETING_TASK8_OVD_QTY").Value = wsSource.Range("MEETING_TASK8_OVD_QTY").Value ' $H$392
    End If
    If .Range("MEETING_TASK8_REM_COUNTRY").Value <> wsSource.Range("MEETING_TASK8_REM_COUNTRY").Value Then
        .Range("MEETING_TASK8_REM_COUNTRY").Value = wsSource.Range("MEETING_TASK8_REM_COUNTRY").Value ' $M$392
    End If
    If .Range("MEETING_TASK8_REM_PCT").Value <> wsSource.Range("MEETING_TASK8_REM_PCT").Value Then
        .Range("MEETING_TASK8_REM_PCT").Value = wsSource.Range("MEETING_TASK8_REM_PCT").Value ' $N$392
    End If
    If .Range("MEETING_TASK9_OVD_JUST").Value <> wsSource.Range("MEETING_TASK9_OVD_JUST").Value Then
        .Range("MEETING_TASK9_OVD_JUST").Value = wsSource.Range("MEETING_TASK9_OVD_JUST").Value ' $I$393
    End If
    If .Range("MEETING_TASK9_OVD_QTY").Value <> wsSource.Range("MEETING_TASK9_OVD_QTY").Value Then
        .Range("MEETING_TASK9_OVD_QTY").Value = wsSource.Range("MEETING_TASK9_OVD_QTY").Value ' $H$393
    End If
    If .Range("MEETING_TASK9_REM_COUNTRY").Value <> wsSource.Range("MEETING_TASK9_REM_COUNTRY").Value Then
        .Range("MEETING_TASK9_REM_COUNTRY").Value = wsSource.Range("MEETING_TASK9_REM_COUNTRY").Value ' $M$393
    End If
    If .Range("MEETING_TASK9_REM_PCT").Value <> wsSource.Range("MEETING_TASK9_REM_PCT").Value Then
        .Range("MEETING_TASK9_REM_PCT").Value = wsSource.Range("MEETING_TASK9_REM_PCT").Value ' $N$393
    End If

    
    
End With

    Call ImportApplicationBasedSheet_428_part2(wsSource, wsTarget)
 '387

End Sub

Sub ImportApplicationBasedSheet_428_part2(wsSource As Worksheet, wsTarget As Worksheet)
'GECE Version = 1.0
On Error Resume Next
With wsTarget

    If .Range("PM_TASK1_OVD_JUST").Value <> wsSource.Range("PM_TASK1_OVD_JUST").Value Then
        .Range("PM_TASK1_OVD_JUST").Value = wsSource.Range("PM_TASK1_OVD_JUST").Value ' $I$364
    End If
    If .Range("PM_TASK1_OVD_QTY").Value <> wsSource.Range("PM_TASK1_OVD_QTY").Value Then
        .Range("PM_TASK1_OVD_QTY").Value = wsSource.Range("PM_TASK1_OVD_QTY").Value ' $H$364
    End If
    If .Range("PM_TASK1_REM_COUNTRY").Value <> wsSource.Range("PM_TASK1_REM_COUNTRY").Value Then
        .Range("PM_TASK1_REM_COUNTRY").Value = wsSource.Range("PM_TASK1_REM_COUNTRY").Value ' $M$364
    End If
    If .Range("PM_TASK1_REM_PCT").Value <> wsSource.Range("PM_TASK1_REM_PCT").Value Then
        .Range("PM_TASK1_REM_PCT").Value = wsSource.Range("PM_TASK1_REM_PCT").Value ' $N$364
    End If
    If .Range("PM_TASK10_OVD_JUST").Value <> wsSource.Range("PM_TASK10_OVD_JUST").Value Then
        .Range("PM_TASK10_OVD_JUST").Value = wsSource.Range("PM_TASK10_OVD_JUST").Value ' $I$373
    End If
    If .Range("PM_TASK10_OVD_QTY").Value <> wsSource.Range("PM_TASK10_OVD_QTY").Value Then
        .Range("PM_TASK10_OVD_QTY").Value = wsSource.Range("PM_TASK10_OVD_QTY").Value ' $H$373
    End If
    If .Range("PM_TASK10_REM_COUNTRY").Value <> wsSource.Range("PM_TASK10_REM_COUNTRY").Value Then
        .Range("PM_TASK10_REM_COUNTRY").Value = wsSource.Range("PM_TASK10_REM_COUNTRY").Value ' $M$373
    End If
    If .Range("PM_TASK10_REM_PCT").Value <> wsSource.Range("PM_TASK10_REM_PCT").Value Then
        .Range("PM_TASK10_REM_PCT").Value = wsSource.Range("PM_TASK10_REM_PCT").Value ' $N$373
    End If
    If .Range("PM_TASK2_OVD_JUST").Value <> wsSource.Range("PM_TASK2_OVD_JUST").Value Then
        .Range("PM_TASK2_OVD_JUST").Value = wsSource.Range("PM_TASK2_OVD_JUST").Value ' $I$365
    End If
    If .Range("PM_TASK2_OVD_QTY").Value <> wsSource.Range("PM_TASK2_OVD_QTY").Value Then
        .Range("PM_TASK2_OVD_QTY").Value = wsSource.Range("PM_TASK2_OVD_QTY").Value ' $H$365
    End If
    If .Range("PM_TASK2_REM_COUNTRY").Value <> wsSource.Range("PM_TASK2_REM_COUNTRY").Value Then
        .Range("PM_TASK2_REM_COUNTRY").Value = wsSource.Range("PM_TASK2_REM_COUNTRY").Value ' $M$365
    End If
    If .Range("PM_TASK2_REM_PCT").Value <> wsSource.Range("PM_TASK2_REM_PCT").Value Then
        .Range("PM_TASK2_REM_PCT").Value = wsSource.Range("PM_TASK2_REM_PCT").Value ' $N$365
    End If
    If .Range("PM_TASK3_REM_COUNTRY").Value <> wsSource.Range("PM_TASK3_REM_COUNTRY").Value Then
        .Range("PM_TASK3_REM_COUNTRY").Value = wsSource.Range("PM_TASK3_REM_COUNTRY").Value ' $M$366
    End If
    If .Range("PM_TASK3_REM_PCT").Value <> wsSource.Range("PM_TASK3_REM_PCT").Value Then
        .Range("PM_TASK3_REM_PCT").Value = wsSource.Range("PM_TASK3_REM_PCT").Value ' $N$366
    End If
    If .Range("PM_TASK4_OVD_JUST").Value <> wsSource.Range("PM_TASK4_OVD_JUST").Value Then
        .Range("PM_TASK4_OVD_JUST").Value = wsSource.Range("PM_TASK4_OVD_JUST").Value ' $I$367
    End If
    If .Range("PM_TASK4_OVD_QTY").Value <> wsSource.Range("PM_TASK4_OVD_QTY").Value Then
        .Range("PM_TASK4_OVD_QTY").Value = wsSource.Range("PM_TASK4_OVD_QTY").Value ' $H$367
    End If
    If .Range("PM_TASK4_REM_COUNTRY").Value <> wsSource.Range("PM_TASK4_REM_COUNTRY").Value Then
        .Range("PM_TASK4_REM_COUNTRY").Value = wsSource.Range("PM_TASK4_REM_COUNTRY").Value ' $M$367
    End If
    If .Range("PM_TASK4_REM_PCT").Value <> wsSource.Range("PM_TASK4_REM_PCT").Value Then
        .Range("PM_TASK4_REM_PCT").Value = wsSource.Range("PM_TASK4_REM_PCT").Value ' $N$367
    End If
    If .Range("PM_TASK5_OVD_JUST").Value <> wsSource.Range("PM_TASK5_OVD_JUST").Value Then
        .Range("PM_TASK5_OVD_JUST").Value = wsSource.Range("PM_TASK5_OVD_JUST").Value ' $I$368
    End If
    If .Range("PM_TASK5_OVD_QTY").Value <> wsSource.Range("PM_TASK5_OVD_QTY").Value Then
        .Range("PM_TASK5_OVD_QTY").Value = wsSource.Range("PM_TASK5_OVD_QTY").Value ' $H$368
    End If
    If .Range("PM_TASK5_REM_COUNTRY").Value <> wsSource.Range("PM_TASK5_REM_COUNTRY").Value Then
        .Range("PM_TASK5_REM_COUNTRY").Value = wsSource.Range("PM_TASK5_REM_COUNTRY").Value ' $M$368
    End If
    If .Range("PM_TASK5_REM_PCT").Value <> wsSource.Range("PM_TASK5_REM_PCT").Value Then
        .Range("PM_TASK5_REM_PCT").Value = wsSource.Range("PM_TASK5_REM_PCT").Value ' $N$368
    End If
    If .Range("PM_TASK6_OVD_JUST").Value <> wsSource.Range("PM_TASK6_OVD_JUST").Value Then
        .Range("PM_TASK6_OVD_JUST").Value = wsSource.Range("PM_TASK6_OVD_JUST").Value ' $I$369
    End If
    If .Range("PM_TASK6_OVD_QTY").Value <> wsSource.Range("PM_TASK6_OVD_QTY").Value Then
        .Range("PM_TASK6_OVD_QTY").Value = wsSource.Range("PM_TASK6_OVD_QTY").Value ' $H$369
    End If
    If .Range("PM_TASK6_REM_COUNTRY").Value <> wsSource.Range("PM_TASK6_REM_COUNTRY").Value Then
        .Range("PM_TASK6_REM_COUNTRY").Value = wsSource.Range("PM_TASK6_REM_COUNTRY").Value ' $M$369
    End If
    If .Range("PM_TASK6_REM_PCT").Value <> wsSource.Range("PM_TASK6_REM_PCT").Value Then
        .Range("PM_TASK6_REM_PCT").Value = wsSource.Range("PM_TASK6_REM_PCT").Value ' $N$369
    End If
    If .Range("PM_TASK7_OVD_JUST").Value <> wsSource.Range("PM_TASK7_OVD_JUST").Value Then
        .Range("PM_TASK7_OVD_JUST").Value = wsSource.Range("PM_TASK7_OVD_JUST").Value ' $I$370
    End If
    If .Range("PM_TASK7_OVD_QTY").Value <> wsSource.Range("PM_TASK7_OVD_QTY").Value Then
        .Range("PM_TASK7_OVD_QTY").Value = wsSource.Range("PM_TASK7_OVD_QTY").Value ' $H$370
    End If
    If .Range("PM_TASK7_REM_COUNTRY").Value <> wsSource.Range("PM_TASK7_REM_COUNTRY").Value Then
        .Range("PM_TASK7_REM_COUNTRY").Value = wsSource.Range("PM_TASK7_REM_COUNTRY").Value ' $M$370
    End If
    If .Range("PM_TASK7_REM_PCT").Value <> wsSource.Range("PM_TASK7_REM_PCT").Value Then
        .Range("PM_TASK7_REM_PCT").Value = wsSource.Range("PM_TASK7_REM_PCT").Value ' $N$370
    End If
    If .Range("PM_TASK8_OVD_JUST").Value <> wsSource.Range("PM_TASK8_OVD_JUST").Value Then
        .Range("PM_TASK8_OVD_JUST").Value = wsSource.Range("PM_TASK8_OVD_JUST").Value ' $I$371
    End If
    If .Range("PM_TASK8_OVD_QTY").Value <> wsSource.Range("PM_TASK8_OVD_QTY").Value Then
        .Range("PM_TASK8_OVD_QTY").Value = wsSource.Range("PM_TASK8_OVD_QTY").Value ' $H$371
    End If
    If .Range("PM_TASK8_REM_COUNTRY").Value <> wsSource.Range("PM_TASK8_REM_COUNTRY").Value Then
        .Range("PM_TASK8_REM_COUNTRY").Value = wsSource.Range("PM_TASK8_REM_COUNTRY").Value ' $M$371
    End If
    If .Range("PM_TASK8_REM_PCT").Value <> wsSource.Range("PM_TASK8_REM_PCT").Value Then
        .Range("PM_TASK8_REM_PCT").Value = wsSource.Range("PM_TASK8_REM_PCT").Value ' $N$371
    End If
    If .Range("PM_TASK9_OVD_JUST").Value <> wsSource.Range("PM_TASK9_OVD_JUST").Value Then
        .Range("PM_TASK9_OVD_JUST").Value = wsSource.Range("PM_TASK9_OVD_JUST").Value ' $I$372
    End If
    If .Range("PM_TASK9_OVD_QTY").Value <> wsSource.Range("PM_TASK9_OVD_QTY").Value Then
        .Range("PM_TASK9_OVD_QTY").Value = wsSource.Range("PM_TASK9_OVD_QTY").Value ' $H$372
    End If
    If .Range("PM_TASK9_REM_COUNTRY").Value <> wsSource.Range("PM_TASK9_REM_COUNTRY").Value Then
        .Range("PM_TASK9_REM_COUNTRY").Value = wsSource.Range("PM_TASK9_REM_COUNTRY").Value ' $M$372
    End If
    If .Range("PM_TASK9_REM_PCT").Value <> wsSource.Range("PM_TASK9_REM_PCT").Value Then
        .Range("PM_TASK9_REM_PCT").Value = wsSource.Range("PM_TASK9_REM_PCT").Value ' $N$372
    End If
    If .Range("REP_TASK1_OVD_JUST").Value <> wsSource.Range("REP_TASK1_OVD_JUST").Value Then
        .Range("REP_TASK1_OVD_JUST").Value = wsSource.Range("REP_TASK1_OVD_JUST").Value ' $I$259
    End If
    If .Range("REP_TASK1_OVD_QTY").Value <> wsSource.Range("REP_TASK1_OVD_QTY").Value Then
        .Range("REP_TASK1_OVD_QTY").Value = wsSource.Range("REP_TASK1_OVD_QTY").Value ' $H$259
    End If
    If .Range("REP_TASK1_REM_COUNTRY").Value <> wsSource.Range("REP_TASK1_REM_COUNTRY").Value Then
        .Range("REP_TASK1_REM_COUNTRY").Value = wsSource.Range("REP_TASK1_REM_COUNTRY").Value ' $M$259
    End If
    If .Range("REP_TASK1_REM_PCT").Value <> wsSource.Range("REP_TASK1_REM_PCT").Value Then
        .Range("REP_TASK1_REM_PCT").Value = wsSource.Range("REP_TASK1_REM_PCT").Value ' $N$259
    End If
    If .Range("REP_TASK10_OVD_JUST").Value <> wsSource.Range("REP_TASK10_OVD_JUST").Value Then
        .Range("REP_TASK10_OVD_JUST").Value = wsSource.Range("REP_TASK10_OVD_JUST").Value ' $I$268
    End If
    If .Range("REP_TASK10_OVD_QTY").Value <> wsSource.Range("REP_TASK10_OVD_QTY").Value Then
        .Range("REP_TASK10_OVD_QTY").Value = wsSource.Range("REP_TASK10_OVD_QTY").Value ' $H$268
    End If
    If .Range("REP_TASK10_REM_COUNTRY").Value <> wsSource.Range("REP_TASK10_REM_COUNTRY").Value Then
        .Range("REP_TASK10_REM_COUNTRY").Value = wsSource.Range("REP_TASK10_REM_COUNTRY").Value ' $M$268
    End If
    If .Range("REP_TASK10_REM_PCT").Value <> wsSource.Range("REP_TASK10_REM_PCT").Value Then
        .Range("REP_TASK10_REM_PCT").Value = wsSource.Range("REP_TASK10_REM_PCT").Value ' $N$268
    End If
    If .Range("REP_TASK2_OVD_JUST").Value <> wsSource.Range("REP_TASK2_OVD_JUST").Value Then
        .Range("REP_TASK2_OVD_JUST").Value = wsSource.Range("REP_TASK2_OVD_JUST").Value ' $I$260
    End If
    If .Range("REP_TASK2_OVD_QTY").Value <> wsSource.Range("REP_TASK2_OVD_QTY").Value Then
        .Range("REP_TASK2_OVD_QTY").Value = wsSource.Range("REP_TASK2_OVD_QTY").Value ' $H$260
    End If
    If .Range("REP_TASK2_REM_COUNTRY").Value <> wsSource.Range("REP_TASK2_REM_COUNTRY").Value Then
        .Range("REP_TASK2_REM_COUNTRY").Value = wsSource.Range("REP_TASK2_REM_COUNTRY").Value ' $M$260
    End If
    If .Range("REP_TASK2_REM_PCT").Value <> wsSource.Range("REP_TASK2_REM_PCT").Value Then
        .Range("REP_TASK2_REM_PCT").Value = wsSource.Range("REP_TASK2_REM_PCT").Value ' $N$260
    End If
    If .Range("REP_TASK3_OVD_JUST").Value <> wsSource.Range("REP_TASK3_OVD_JUST").Value Then
        .Range("REP_TASK3_OVD_JUST").Value = wsSource.Range("REP_TASK3_OVD_JUST").Value ' $I$261
    End If
    If .Range("REP_TASK3_OVD_QTY").Value <> wsSource.Range("REP_TASK3_OVD_QTY").Value Then
        .Range("REP_TASK3_OVD_QTY").Value = wsSource.Range("REP_TASK3_OVD_QTY").Value ' $H$261
    End If
    If .Range("REP_TASK3_REM_COUNTRY").Value <> wsSource.Range("REP_TASK3_REM_COUNTRY").Value Then
        .Range("REP_TASK3_REM_COUNTRY").Value = wsSource.Range("REP_TASK3_REM_COUNTRY").Value ' $M$261
    End If
    If .Range("REP_TASK3_REM_PCT").Value <> wsSource.Range("REP_TASK3_REM_PCT").Value Then
        .Range("REP_TASK3_REM_PCT").Value = wsSource.Range("REP_TASK3_REM_PCT").Value ' $N$261
    End If
    If .Range("REP_TASK4_OVD_JUST").Value <> wsSource.Range("REP_TASK4_OVD_JUST").Value Then
        .Range("REP_TASK4_OVD_JUST").Value = wsSource.Range("REP_TASK4_OVD_JUST").Value ' $I$262
    End If
    If .Range("REP_TASK4_OVD_QTY").Value <> wsSource.Range("REP_TASK4_OVD_QTY").Value Then
        .Range("REP_TASK4_OVD_QTY").Value = wsSource.Range("REP_TASK4_OVD_QTY").Value ' $H$262
    End If
    If .Range("REP_TASK4_REM_COUNTRY").Value <> wsSource.Range("REP_TASK4_REM_COUNTRY").Value Then
        .Range("REP_TASK4_REM_COUNTRY").Value = wsSource.Range("REP_TASK4_REM_COUNTRY").Value ' $M$262
    End If
    If .Range("REP_TASK4_REM_PCT").Value <> wsSource.Range("REP_TASK4_REM_PCT").Value Then
        .Range("REP_TASK4_REM_PCT").Value = wsSource.Range("REP_TASK4_REM_PCT").Value ' $N$262
    End If
    If .Range("REP_TASK5_OVD_JUST").Value <> wsSource.Range("REP_TASK5_OVD_JUST").Value Then
        .Range("REP_TASK5_OVD_JUST").Value = wsSource.Range("REP_TASK5_OVD_JUST").Value ' $I$263
    End If
    If .Range("REP_TASK5_OVD_QTY").Value <> wsSource.Range("REP_TASK5_OVD_QTY").Value Then
        .Range("REP_TASK5_OVD_QTY").Value = wsSource.Range("REP_TASK5_OVD_QTY").Value ' $H$263
    End If
    If .Range("REP_TASK5_REM_COUNTRY").Value <> wsSource.Range("REP_TASK5_REM_COUNTRY").Value Then
        .Range("REP_TASK5_REM_COUNTRY").Value = wsSource.Range("REP_TASK5_REM_COUNTRY").Value ' $M$263
    End If
    If .Range("REP_TASK5_REM_PCT").Value <> wsSource.Range("REP_TASK5_REM_PCT").Value Then
        .Range("REP_TASK5_REM_PCT").Value = wsSource.Range("REP_TASK5_REM_PCT").Value ' $N$263
    End If
    If .Range("REP_TASK6_OVD_JUST").Value <> wsSource.Range("REP_TASK6_OVD_JUST").Value Then
        .Range("REP_TASK6_OVD_JUST").Value = wsSource.Range("REP_TASK6_OVD_JUST").Value ' $I$264
    End If
    If .Range("REP_TASK6_OVD_QTY").Value <> wsSource.Range("REP_TASK6_OVD_QTY").Value Then
        .Range("REP_TASK6_OVD_QTY").Value = wsSource.Range("REP_TASK6_OVD_QTY").Value ' $H$264
    End If
    If .Range("REP_TASK6_REM_COUNTRY").Value <> wsSource.Range("REP_TASK6_REM_COUNTRY").Value Then
        .Range("REP_TASK6_REM_COUNTRY").Value = wsSource.Range("REP_TASK6_REM_COUNTRY").Value ' $M$264
    End If
    If .Range("REP_TASK6_REM_PCT").Value <> wsSource.Range("REP_TASK6_REM_PCT").Value Then
        .Range("REP_TASK6_REM_PCT").Value = wsSource.Range("REP_TASK6_REM_PCT").Value ' $N$264
    End If
    If .Range("REP_TASK7_OVD_JUST").Value <> wsSource.Range("REP_TASK7_OVD_JUST").Value Then
        .Range("REP_TASK7_OVD_JUST").Value = wsSource.Range("REP_TASK7_OVD_JUST").Value ' $I$265
    End If
    If .Range("REP_TASK7_OVD_QTY").Value <> wsSource.Range("REP_TASK7_OVD_QTY").Value Then
        .Range("REP_TASK7_OVD_QTY").Value = wsSource.Range("REP_TASK7_OVD_QTY").Value ' $H$265
    End If
    If .Range("REP_TASK7_REM_COUNTRY").Value <> wsSource.Range("REP_TASK7_REM_COUNTRY").Value Then
        .Range("REP_TASK7_REM_COUNTRY").Value = wsSource.Range("REP_TASK7_REM_COUNTRY").Value ' $M$265
    End If
    If .Range("REP_TASK7_REM_PCT").Value <> wsSource.Range("REP_TASK7_REM_PCT").Value Then
        .Range("REP_TASK7_REM_PCT").Value = wsSource.Range("REP_TASK7_REM_PCT").Value ' $N$265
    End If
    If .Range("REP_TASK8_OVD_JUST").Value <> wsSource.Range("REP_TASK8_OVD_JUST").Value Then
        .Range("REP_TASK8_OVD_JUST").Value = wsSource.Range("REP_TASK8_OVD_JUST").Value ' $I$266
    End If
    If .Range("REP_TASK8_OVD_QTY").Value <> wsSource.Range("REP_TASK8_OVD_QTY").Value Then
        .Range("REP_TASK8_OVD_QTY").Value = wsSource.Range("REP_TASK8_OVD_QTY").Value ' $H$266
    End If
    If .Range("REP_TASK8_REM_COUNTRY").Value <> wsSource.Range("REP_TASK8_REM_COUNTRY").Value Then
        .Range("REP_TASK8_REM_COUNTRY").Value = wsSource.Range("REP_TASK8_REM_COUNTRY").Value ' $M$266
    End If
    If .Range("REP_TASK8_REM_PCT").Value <> wsSource.Range("REP_TASK8_REM_PCT").Value Then
        .Range("REP_TASK8_REM_PCT").Value = wsSource.Range("REP_TASK8_REM_PCT").Value ' $N$266
    End If
    If .Range("REP_TASK9_OVD_JUST").Value <> wsSource.Range("REP_TASK9_OVD_JUST").Value Then
        .Range("REP_TASK9_OVD_JUST").Value = wsSource.Range("REP_TASK9_OVD_JUST").Value ' $I$267
    End If
    If .Range("REP_TASK9_OVD_QTY").Value <> wsSource.Range("REP_TASK9_OVD_QTY").Value Then
        .Range("REP_TASK9_OVD_QTY").Value = wsSource.Range("REP_TASK9_OVD_QTY").Value ' $H$267
    End If
    If .Range("REP_TASK9_REM_COUNTRY").Value <> wsSource.Range("REP_TASK9_REM_COUNTRY").Value Then
        .Range("REP_TASK9_REM_COUNTRY").Value = wsSource.Range("REP_TASK9_REM_COUNTRY").Value ' $M$267
    End If
    If .Range("REP_TASK9_REM_PCT").Value <> wsSource.Range("REP_TASK9_REM_PCT").Value Then
        .Range("REP_TASK9_REM_PCT").Value = wsSource.Range("REP_TASK9_REM_PCT").Value ' $N$267
    End If
    If .Range("SITE_TASK1_OVD_JUST").Value <> wsSource.Range("SITE_TASK1_OVD_JUST").Value Then
        .Range("SITE_TASK1_OVD_JUST").Value = wsSource.Range("SITE_TASK1_OVD_JUST").Value ' $I$406
    End If
    If .Range("SITE_TASK1_OVD_QTY").Value <> wsSource.Range("SITE_TASK1_OVD_QTY").Value Then
        .Range("SITE_TASK1_OVD_QTY").Value = wsSource.Range("SITE_TASK1_OVD_QTY").Value ' $H$406
    End If
    If .Range("SITE_TASK1_REM_COUNTRY").Value <> wsSource.Range("SITE_TASK1_REM_COUNTRY").Value Then
        .Range("SITE_TASK1_REM_COUNTRY").Value = wsSource.Range("SITE_TASK1_REM_COUNTRY").Value ' $M$406
    End If
    If .Range("SITE_TASK1_REM_PCT").Value <> wsSource.Range("SITE_TASK1_REM_PCT").Value Then
        .Range("SITE_TASK1_REM_PCT").Value = wsSource.Range("SITE_TASK1_REM_PCT").Value ' $N$406
    End If
    If .Range("SITE_TASK10_OVD_JUST").Value <> wsSource.Range("SITE_TASK10_OVD_JUST").Value Then
        .Range("SITE_TASK10_OVD_JUST").Value = wsSource.Range("SITE_TASK10_OVD_JUST").Value ' $I$415
    End If
    If .Range("SITE_TASK10_OVD_QTY").Value <> wsSource.Range("SITE_TASK10_OVD_QTY").Value Then
        .Range("SITE_TASK10_OVD_QTY").Value = wsSource.Range("SITE_TASK10_OVD_QTY").Value ' $H$415
    End If
    If .Range("SITE_TASK10_REM_COUNTRY").Value <> wsSource.Range("SITE_TASK10_REM_COUNTRY").Value Then
        .Range("SITE_TASK10_REM_COUNTRY").Value = wsSource.Range("SITE_TASK10_REM_COUNTRY").Value ' $M$415
    End If
    If .Range("SITE_TASK10_REM_PCT").Value <> wsSource.Range("SITE_TASK10_REM_PCT").Value Then
        .Range("SITE_TASK10_REM_PCT").Value = wsSource.Range("SITE_TASK10_REM_PCT").Value ' $N$415
    End If
    If .Range("SITE_TASK2_OVD_JUST").Value <> wsSource.Range("SITE_TASK2_OVD_JUST").Value Then
        .Range("SITE_TASK2_OVD_JUST").Value = wsSource.Range("SITE_TASK2_OVD_JUST").Value ' $I$407
    End If
    If .Range("SITE_TASK2_OVD_QTY").Value <> wsSource.Range("SITE_TASK2_OVD_QTY").Value Then
        .Range("SITE_TASK2_OVD_QTY").Value = wsSource.Range("SITE_TASK2_OVD_QTY").Value ' $H$407
    End If
    If .Range("SITE_TASK2_REM_COUNTRY").Value <> wsSource.Range("SITE_TASK2_REM_COUNTRY").Value Then
        .Range("SITE_TASK2_REM_COUNTRY").Value = wsSource.Range("SITE_TASK2_REM_COUNTRY").Value ' $M$407
    End If
    If .Range("SITE_TASK2_REM_PCT").Value <> wsSource.Range("SITE_TASK2_REM_PCT").Value Then
        .Range("SITE_TASK2_REM_PCT").Value = wsSource.Range("SITE_TASK2_REM_PCT").Value ' $N$407
    End If
    If .Range("SITE_TASK3_OVD_JUST").Value <> wsSource.Range("SITE_TASK3_OVD_JUST").Value Then
        .Range("SITE_TASK3_OVD_JUST").Value = wsSource.Range("SITE_TASK3_OVD_JUST").Value ' $I$408
    End If
    If .Range("SITE_TASK3_OVD_QTY").Value <> wsSource.Range("SITE_TASK3_OVD_QTY").Value Then
        .Range("SITE_TASK3_OVD_QTY").Value = wsSource.Range("SITE_TASK3_OVD_QTY").Value ' $H$408
    End If
    If .Range("SITE_TASK3_REM_COUNTRY").Value <> wsSource.Range("SITE_TASK3_REM_COUNTRY").Value Then
        .Range("SITE_TASK3_REM_COUNTRY").Value = wsSource.Range("SITE_TASK3_REM_COUNTRY").Value ' $M$408
    End If
    If .Range("SITE_TASK3_REM_PCT").Value <> wsSource.Range("SITE_TASK3_REM_PCT").Value Then
        .Range("SITE_TASK3_REM_PCT").Value = wsSource.Range("SITE_TASK3_REM_PCT").Value ' $N$408
    End If
    If .Range("SITE_TASK4_OVD_JUST").Value <> wsSource.Range("SITE_TASK4_OVD_JUST").Value Then
        .Range("SITE_TASK4_OVD_JUST").Value = wsSource.Range("SITE_TASK4_OVD_JUST").Value ' $I$409
    End If
    If .Range("SITE_TASK4_OVD_QTY").Value <> wsSource.Range("SITE_TASK4_OVD_QTY").Value Then
        .Range("SITE_TASK4_OVD_QTY").Value = wsSource.Range("SITE_TASK4_OVD_QTY").Value ' $H$409
    End If
    If .Range("SITE_TASK4_REM_COUNTRY").Value <> wsSource.Range("SITE_TASK4_REM_COUNTRY").Value Then
        .Range("SITE_TASK4_REM_COUNTRY").Value = wsSource.Range("SITE_TASK4_REM_COUNTRY").Value ' $M$409
    End If
    If .Range("SITE_TASK4_REM_PCT").Value <> wsSource.Range("SITE_TASK4_REM_PCT").Value Then
        .Range("SITE_TASK4_REM_PCT").Value = wsSource.Range("SITE_TASK4_REM_PCT").Value ' $N$409
    End If
    If .Range("SITE_TASK5_OVD_JUST").Value <> wsSource.Range("SITE_TASK5_OVD_JUST").Value Then
        .Range("SITE_TASK5_OVD_JUST").Value = wsSource.Range("SITE_TASK5_OVD_JUST").Value ' $I$410
    End If
    If .Range("SITE_TASK5_OVD_QTY").Value <> wsSource.Range("SITE_TASK5_OVD_QTY").Value Then
        .Range("SITE_TASK5_OVD_QTY").Value = wsSource.Range("SITE_TASK5_OVD_QTY").Value ' $H$410
    End If
    If .Range("SITE_TASK5_REM_COUNTRY").Value <> wsSource.Range("SITE_TASK5_REM_COUNTRY").Value Then
        .Range("SITE_TASK5_REM_COUNTRY").Value = wsSource.Range("SITE_TASK5_REM_COUNTRY").Value ' $M$410
    End If
    If .Range("SITE_TASK5_REM_PCT").Value <> wsSource.Range("SITE_TASK5_REM_PCT").Value Then
        .Range("SITE_TASK5_REM_PCT").Value = wsSource.Range("SITE_TASK5_REM_PCT").Value ' $N$410
    End If
    If .Range("SITE_TASK6_OVD_JUST").Value <> wsSource.Range("SITE_TASK6_OVD_JUST").Value Then
        .Range("SITE_TASK6_OVD_JUST").Value = wsSource.Range("SITE_TASK6_OVD_JUST").Value ' $I$411
    End If
    If .Range("SITE_TASK6_OVD_QTY").Value <> wsSource.Range("SITE_TASK6_OVD_QTY").Value Then
        .Range("SITE_TASK6_OVD_QTY").Value = wsSource.Range("SITE_TASK6_OVD_QTY").Value ' $H$411
    End If
    If .Range("SITE_TASK6_REM_COUNTRY").Value <> wsSource.Range("SITE_TASK6_REM_COUNTRY").Value Then
        .Range("SITE_TASK6_REM_COUNTRY").Value = wsSource.Range("SITE_TASK6_REM_COUNTRY").Value ' $M$411
    End If
    If .Range("SITE_TASK6_REM_PCT").Value <> wsSource.Range("SITE_TASK6_REM_PCT").Value Then
        .Range("SITE_TASK6_REM_PCT").Value = wsSource.Range("SITE_TASK6_REM_PCT").Value ' $N$411
    End If
    If .Range("SITE_TASK7_OVD_JUST").Value <> wsSource.Range("SITE_TASK7_OVD_JUST").Value Then
        .Range("SITE_TASK7_OVD_JUST").Value = wsSource.Range("SITE_TASK7_OVD_JUST").Value ' $I$412
    End If
    If .Range("SITE_TASK7_OVD_QTY").Value <> wsSource.Range("SITE_TASK7_OVD_QTY").Value Then
        .Range("SITE_TASK7_OVD_QTY").Value = wsSource.Range("SITE_TASK7_OVD_QTY").Value ' $H$412
    End If
    If .Range("SITE_TASK7_REM_COUNTRY").Value <> wsSource.Range("SITE_TASK7_REM_COUNTRY").Value Then
        .Range("SITE_TASK7_REM_COUNTRY").Value = wsSource.Range("SITE_TASK7_REM_COUNTRY").Value ' $M$412
    End If
    If .Range("SITE_TASK7_REM_PCT").Value <> wsSource.Range("SITE_TASK7_REM_PCT").Value Then
        .Range("SITE_TASK7_REM_PCT").Value = wsSource.Range("SITE_TASK7_REM_PCT").Value ' $N$412
    End If
    If .Range("SITE_TASK8_OVD_JUST").Value <> wsSource.Range("SITE_TASK8_OVD_JUST").Value Then
        .Range("SITE_TASK8_OVD_JUST").Value = wsSource.Range("SITE_TASK8_OVD_JUST").Value ' $I$413
    End If
    If .Range("SITE_TASK8_OVD_QTY").Value <> wsSource.Range("SITE_TASK8_OVD_QTY").Value Then
        .Range("SITE_TASK8_OVD_QTY").Value = wsSource.Range("SITE_TASK8_OVD_QTY").Value ' $H$413
    End If
    If .Range("SITE_TASK8_REM_COUNTRY").Value <> wsSource.Range("SITE_TASK8_REM_COUNTRY").Value Then
        .Range("SITE_TASK8_REM_COUNTRY").Value = wsSource.Range("SITE_TASK8_REM_COUNTRY").Value ' $M$413
    End If
    If .Range("SITE_TASK8_REM_PCT").Value <> wsSource.Range("SITE_TASK8_REM_PCT").Value Then
        .Range("SITE_TASK8_REM_PCT").Value = wsSource.Range("SITE_TASK8_REM_PCT").Value ' $N$413
    End If
    If .Range("SITE_TASK9_OVD_JUST").Value <> wsSource.Range("SITE_TASK9_OVD_JUST").Value Then
        .Range("SITE_TASK9_OVD_JUST").Value = wsSource.Range("SITE_TASK9_OVD_JUST").Value ' $I$414
    End If
    If .Range("SITE_TASK9_OVD_QTY").Value <> wsSource.Range("SITE_TASK9_OVD_QTY").Value Then
        .Range("SITE_TASK9_OVD_QTY").Value = wsSource.Range("SITE_TASK9_OVD_QTY").Value ' $H$414
    End If
    If .Range("SITE_TASK9_REM_COUNTRY").Value <> wsSource.Range("SITE_TASK9_REM_COUNTRY").Value Then
        .Range("SITE_TASK9_REM_COUNTRY").Value = wsSource.Range("SITE_TASK9_REM_COUNTRY").Value ' $M$414
    End If
    If .Range("SITE_TASK9_REM_PCT").Value <> wsSource.Range("SITE_TASK9_REM_PCT").Value Then
        .Range("SITE_TASK9_REM_PCT").Value = wsSource.Range("SITE_TASK9_REM_PCT").Value ' $N$414
    End If
    If .Range("SPEC_TASK1_OVD_JUST").Value <> wsSource.Range("SPEC_TASK1_OVD_JUST").Value Then
        .Range("SPEC_TASK1_OVD_JUST").Value = wsSource.Range("SPEC_TASK1_OVD_JUST").Value ' $I$129
    End If
    If .Range("SPEC_TASK1_OVD_QTY").Value <> wsSource.Range("SPEC_TASK1_OVD_QTY").Value Then
        .Range("SPEC_TASK1_OVD_QTY").Value = wsSource.Range("SPEC_TASK1_OVD_QTY").Value ' $H$129
    End If
    If .Range("SPEC_TASK1_REM_COUNTRY").Value <> wsSource.Range("SPEC_TASK1_REM_COUNTRY").Value Then
        .Range("SPEC_TASK1_REM_COUNTRY").Value = wsSource.Range("SPEC_TASK1_REM_COUNTRY").Value ' $M$129
    End If
    If .Range("SPEC_TASK1_REM_PCT").Value <> wsSource.Range("SPEC_TASK1_REM_PCT").Value Then
        .Range("SPEC_TASK1_REM_PCT").Value = wsSource.Range("SPEC_TASK1_REM_PCT").Value ' $N$129
    End If
    If .Range("SPEC_TASK10_OVD_JUST").Value <> wsSource.Range("SPEC_TASK10_OVD_JUST").Value Then
        .Range("SPEC_TASK10_OVD_JUST").Value = wsSource.Range("SPEC_TASK10_OVD_JUST").Value ' $I$138
    End If
    If .Range("SPEC_TASK10_OVD_QTY").Value <> wsSource.Range("SPEC_TASK10_OVD_QTY").Value Then
        .Range("SPEC_TASK10_OVD_QTY").Value = wsSource.Range("SPEC_TASK10_OVD_QTY").Value ' $H$138
    End If
    If .Range("SPEC_TASK10_REM_COUNTRY").Value <> wsSource.Range("SPEC_TASK10_REM_COUNTRY").Value Then
        .Range("SPEC_TASK10_REM_COUNTRY").Value = wsSource.Range("SPEC_TASK10_REM_COUNTRY").Value ' $M$138
    End If
    If .Range("SPEC_TASK10_REM_PCT").Value <> wsSource.Range("SPEC_TASK10_REM_PCT").Value Then
        .Range("SPEC_TASK10_REM_PCT").Value = wsSource.Range("SPEC_TASK10_REM_PCT").Value ' $N$138
    End If
    If .Range("SPEC_TASK2_OVD_JUST").Value <> wsSource.Range("SPEC_TASK2_OVD_JUST").Value Then
        .Range("SPEC_TASK2_OVD_JUST").Value = wsSource.Range("SPEC_TASK2_OVD_JUST").Value ' $I$130
    End If
    If .Range("SPEC_TASK2_OVD_QTY").Value <> wsSource.Range("SPEC_TASK2_OVD_QTY").Value Then
        .Range("SPEC_TASK2_OVD_QTY").Value = wsSource.Range("SPEC_TASK2_OVD_QTY").Value ' $H$130
    End If
    If .Range("SPEC_TASK2_REM_COUNTRY").Value <> wsSource.Range("SPEC_TASK2_REM_COUNTRY").Value Then
        .Range("SPEC_TASK2_REM_COUNTRY").Value = wsSource.Range("SPEC_TASK2_REM_COUNTRY").Value ' $M$130
    End If
    If .Range("SPEC_TASK2_REM_PCT").Value <> wsSource.Range("SPEC_TASK2_REM_PCT").Value Then
        .Range("SPEC_TASK2_REM_PCT").Value = wsSource.Range("SPEC_TASK2_REM_PCT").Value ' $N$130
    End If
    If .Range("SPEC_TASK3_OVD_JUST").Value <> wsSource.Range("SPEC_TASK3_OVD_JUST").Value Then
        .Range("SPEC_TASK3_OVD_JUST").Value = wsSource.Range("SPEC_TASK3_OVD_JUST").Value ' $I$131
    End If
    If .Range("SPEC_TASK3_OVD_QTY").Value <> wsSource.Range("SPEC_TASK3_OVD_QTY").Value Then
        .Range("SPEC_TASK3_OVD_QTY").Value = wsSource.Range("SPEC_TASK3_OVD_QTY").Value ' $H$131
    End If
    If .Range("SPEC_TASK3_REM_COUNTRY").Value <> wsSource.Range("SPEC_TASK3_REM_COUNTRY").Value Then
        .Range("SPEC_TASK3_REM_COUNTRY").Value = wsSource.Range("SPEC_TASK3_REM_COUNTRY").Value ' $M$131
    End If
    If .Range("SPEC_TASK3_REM_PCT").Value <> wsSource.Range("SPEC_TASK3_REM_PCT").Value Then
        .Range("SPEC_TASK3_REM_PCT").Value = wsSource.Range("SPEC_TASK3_REM_PCT").Value ' $N$131
    End If
    If .Range("SPEC_TASK4_OVD_JUST").Value <> wsSource.Range("SPEC_TASK4_OVD_JUST").Value Then
        .Range("SPEC_TASK4_OVD_JUST").Value = wsSource.Range("SPEC_TASK4_OVD_JUST").Value ' $I$132
    End If
    If .Range("SPEC_TASK4_OVD_QTY").Value <> wsSource.Range("SPEC_TASK4_OVD_QTY").Value Then
        .Range("SPEC_TASK4_OVD_QTY").Value = wsSource.Range("SPEC_TASK4_OVD_QTY").Value ' $H$132
    End If
    If .Range("SPEC_TASK4_REM_COUNTRY").Value <> wsSource.Range("SPEC_TASK4_REM_COUNTRY").Value Then
        .Range("SPEC_TASK4_REM_COUNTRY").Value = wsSource.Range("SPEC_TASK4_REM_COUNTRY").Value ' $M$132
    End If
    If .Range("SPEC_TASK4_REM_PCT").Value <> wsSource.Range("SPEC_TASK4_REM_PCT").Value Then
        .Range("SPEC_TASK4_REM_PCT").Value = wsSource.Range("SPEC_TASK4_REM_PCT").Value ' $N$132
    End If
    If .Range("SPEC_TASK5_OVD_JUST").Value <> wsSource.Range("SPEC_TASK5_OVD_JUST").Value Then
        .Range("SPEC_TASK5_OVD_JUST").Value = wsSource.Range("SPEC_TASK5_OVD_JUST").Value ' $I$133
    End If
    If .Range("SPEC_TASK5_OVD_QTY").Value <> wsSource.Range("SPEC_TASK5_OVD_QTY").Value Then
        .Range("SPEC_TASK5_OVD_QTY").Value = wsSource.Range("SPEC_TASK5_OVD_QTY").Value ' $H$133
    End If
    If .Range("SPEC_TASK5_REM_COUNTRY").Value <> wsSource.Range("SPEC_TASK5_REM_COUNTRY").Value Then
        .Range("SPEC_TASK5_REM_COUNTRY").Value = wsSource.Range("SPEC_TASK5_REM_COUNTRY").Value ' $M$133
    End If
    If .Range("SPEC_TASK5_REM_PCT").Value <> wsSource.Range("SPEC_TASK5_REM_PCT").Value Then
        .Range("SPEC_TASK5_REM_PCT").Value = wsSource.Range("SPEC_TASK5_REM_PCT").Value ' $N$133
    End If
    If .Range("SPEC_TASK6_OVD_JUST").Value <> wsSource.Range("SPEC_TASK6_OVD_JUST").Value Then
        .Range("SPEC_TASK6_OVD_JUST").Value = wsSource.Range("SPEC_TASK6_OVD_JUST").Value ' $I$134
    End If
    If .Range("SPEC_TASK6_OVD_QTY").Value <> wsSource.Range("SPEC_TASK6_OVD_QTY").Value Then
        .Range("SPEC_TASK6_OVD_QTY").Value = wsSource.Range("SPEC_TASK6_OVD_QTY").Value ' $H$134
    End If
    If .Range("SPEC_TASK6_REM_COUNTRY").Value <> wsSource.Range("SPEC_TASK6_REM_COUNTRY").Value Then
        .Range("SPEC_TASK6_REM_COUNTRY").Value = wsSource.Range("SPEC_TASK6_REM_COUNTRY").Value ' $M$134
    End If
    If .Range("SPEC_TASK6_REM_PCT").Value <> wsSource.Range("SPEC_TASK6_REM_PCT").Value Then
        .Range("SPEC_TASK6_REM_PCT").Value = wsSource.Range("SPEC_TASK6_REM_PCT").Value ' $N$134
    End If
    If .Range("SPEC_TASK7_OVD_JUST").Value <> wsSource.Range("SPEC_TASK7_OVD_JUST").Value Then
        .Range("SPEC_TASK7_OVD_JUST").Value = wsSource.Range("SPEC_TASK7_OVD_JUST").Value ' $I$135
    End If
    If .Range("SPEC_TASK7_OVD_QTY").Value <> wsSource.Range("SPEC_TASK7_OVD_QTY").Value Then
        .Range("SPEC_TASK7_OVD_QTY").Value = wsSource.Range("SPEC_TASK7_OVD_QTY").Value ' $H$135
    End If
    If .Range("SPEC_TASK7_REM_COUNTRY").Value <> wsSource.Range("SPEC_TASK7_REM_COUNTRY").Value Then
        .Range("SPEC_TASK7_REM_COUNTRY").Value = wsSource.Range("SPEC_TASK7_REM_COUNTRY").Value ' $M$135
    End If
    If .Range("SPEC_TASK7_REM_PCT").Value <> wsSource.Range("SPEC_TASK7_REM_PCT").Value Then
        .Range("SPEC_TASK7_REM_PCT").Value = wsSource.Range("SPEC_TASK7_REM_PCT").Value ' $N$135
    End If
    If .Range("SPEC_TASK8_OVD_JUST").Value <> wsSource.Range("SPEC_TASK8_OVD_JUST").Value Then
        .Range("SPEC_TASK8_OVD_JUST").Value = wsSource.Range("SPEC_TASK8_OVD_JUST").Value ' $I$136
    End If
    If .Range("SPEC_TASK8_OVD_QTY").Value <> wsSource.Range("SPEC_TASK8_OVD_QTY").Value Then
        .Range("SPEC_TASK8_OVD_QTY").Value = wsSource.Range("SPEC_TASK8_OVD_QTY").Value ' $H$136
    End If
    If .Range("SPEC_TASK8_REM_COUNTRY").Value <> wsSource.Range("SPEC_TASK8_REM_COUNTRY").Value Then
        .Range("SPEC_TASK8_REM_COUNTRY").Value = wsSource.Range("SPEC_TASK8_REM_COUNTRY").Value ' $M$136
    End If
    If .Range("SPEC_TASK8_REM_PCT").Value <> wsSource.Range("SPEC_TASK8_REM_PCT").Value Then
        .Range("SPEC_TASK8_REM_PCT").Value = wsSource.Range("SPEC_TASK8_REM_PCT").Value ' $N$136
    End If
    If .Range("SPEC_TASK9_OVD_JUST").Value <> wsSource.Range("SPEC_TASK9_OVD_JUST").Value Then
        .Range("SPEC_TASK9_OVD_JUST").Value = wsSource.Range("SPEC_TASK9_OVD_JUST").Value ' $I$137
    End If
    If .Range("SPEC_TASK9_OVD_QTY").Value <> wsSource.Range("SPEC_TASK9_OVD_QTY").Value Then
        .Range("SPEC_TASK9_OVD_QTY").Value = wsSource.Range("SPEC_TASK9_OVD_QTY").Value ' $H$137
    End If
    If .Range("SPEC_TASK9_REM_COUNTRY").Value <> wsSource.Range("SPEC_TASK9_REM_COUNTRY").Value Then
        .Range("SPEC_TASK9_REM_COUNTRY").Value = wsSource.Range("SPEC_TASK9_REM_COUNTRY").Value ' $M$137
    End If
    If .Range("SPEC_TASK9_REM_PCT").Value <> wsSource.Range("SPEC_TASK9_REM_PCT").Value Then
        .Range("SPEC_TASK9_REM_PCT").Value = wsSource.Range("SPEC_TASK9_REM_PCT").Value ' $N$137
    End If
    If .Range("SYSENG_TASK1_OVD_JUST").Value <> wsSource.Range("SYSENG_TASK1_OVD_JUST").Value Then
        .Range("SYSENG_TASK1_OVD_JUST").Value = wsSource.Range("SYSENG_TASK1_OVD_JUST").Value ' $I$150
    End If
    If .Range("SYSENG_TASK1_OVD_QTY").Value <> wsSource.Range("SYSENG_TASK1_OVD_QTY").Value Then
        .Range("SYSENG_TASK1_OVD_QTY").Value = wsSource.Range("SYSENG_TASK1_OVD_QTY").Value ' $H$150
    End If
    If .Range("SYSENG_TASK1_REM_COUNTRY").Value <> wsSource.Range("SYSENG_TASK1_REM_COUNTRY").Value Then
        .Range("SYSENG_TASK1_REM_COUNTRY").Value = wsSource.Range("SYSENG_TASK1_REM_COUNTRY").Value ' $M$150
    End If
    If .Range("SYSENG_TASK1_REM_PCT").Value <> wsSource.Range("SYSENG_TASK1_REM_PCT").Value Then
        .Range("SYSENG_TASK1_REM_PCT").Value = wsSource.Range("SYSENG_TASK1_REM_PCT").Value ' $N$150
    End If
    If .Range("SYSENG_TASK10_OVD_JUST").Value <> wsSource.Range("SYSENG_TASK10_OVD_JUST").Value Then
        .Range("SYSENG_TASK10_OVD_JUST").Value = wsSource.Range("SYSENG_TASK10_OVD_JUST").Value ' $I$159
    End If
    If .Range("SYSENG_TASK10_OVD_QTY").Value <> wsSource.Range("SYSENG_TASK10_OVD_QTY").Value Then
        .Range("SYSENG_TASK10_OVD_QTY").Value = wsSource.Range("SYSENG_TASK10_OVD_QTY").Value ' $H$159
    End If
    If .Range("SYSENG_TASK10_REM_COUNTRY").Value <> wsSource.Range("SYSENG_TASK10_REM_COUNTRY").Value Then
        .Range("SYSENG_TASK10_REM_COUNTRY").Value = wsSource.Range("SYSENG_TASK10_REM_COUNTRY").Value ' $M$159
    End If
    If .Range("SYSENG_TASK10_REM_PCT").Value <> wsSource.Range("SYSENG_TASK10_REM_PCT").Value Then
        .Range("SYSENG_TASK10_REM_PCT").Value = wsSource.Range("SYSENG_TASK10_REM_PCT").Value ' $N$159
    End If
    If .Range("SYSENG_TASK2_OVD_JUST").Value <> wsSource.Range("SYSENG_TASK2_OVD_JUST").Value Then
        .Range("SYSENG_TASK2_OVD_JUST").Value = wsSource.Range("SYSENG_TASK2_OVD_JUST").Value ' $I$151
    End If
    If .Range("SYSENG_TASK2_OVD_QTY").Value <> wsSource.Range("SYSENG_TASK2_OVD_QTY").Value Then
        .Range("SYSENG_TASK2_OVD_QTY").Value = wsSource.Range("SYSENG_TASK2_OVD_QTY").Value ' $H$151
    End If
    If .Range("SYSENG_TASK2_REM_COUNTRY").Value <> wsSource.Range("SYSENG_TASK2_REM_COUNTRY").Value Then
        .Range("SYSENG_TASK2_REM_COUNTRY").Value = wsSource.Range("SYSENG_TASK2_REM_COUNTRY").Value ' $M$151
    End If
    If .Range("SYSENG_TASK2_REM_PCT").Value <> wsSource.Range("SYSENG_TASK2_REM_PCT").Value Then
        .Range("SYSENG_TASK2_REM_PCT").Value = wsSource.Range("SYSENG_TASK2_REM_PCT").Value ' $N$151
    End If
    If .Range("SYSENG_TASK3_OVD_JUST").Value <> wsSource.Range("SYSENG_TASK3_OVD_JUST").Value Then
        .Range("SYSENG_TASK3_OVD_JUST").Value = wsSource.Range("SYSENG_TASK3_OVD_JUST").Value ' $I$152
    End If
    If .Range("SYSENG_TASK3_OVD_QTY").Value <> wsSource.Range("SYSENG_TASK3_OVD_QTY").Value Then
        .Range("SYSENG_TASK3_OVD_QTY").Value = wsSource.Range("SYSENG_TASK3_OVD_QTY").Value ' $H$152
    End If
    If .Range("SYSENG_TASK3_REM_COUNTRY").Value <> wsSource.Range("SYSENG_TASK3_REM_COUNTRY").Value Then
        .Range("SYSENG_TASK3_REM_COUNTRY").Value = wsSource.Range("SYSENG_TASK3_REM_COUNTRY").Value ' $M$152
    End If
    If .Range("SYSENG_TASK3_REM_PCT").Value <> wsSource.Range("SYSENG_TASK3_REM_PCT").Value Then
        .Range("SYSENG_TASK3_REM_PCT").Value = wsSource.Range("SYSENG_TASK3_REM_PCT").Value ' $N$152
    End If
    If .Range("SYSENG_TASK4_OVD_JUST").Value <> wsSource.Range("SYSENG_TASK4_OVD_JUST").Value Then
        .Range("SYSENG_TASK4_OVD_JUST").Value = wsSource.Range("SYSENG_TASK4_OVD_JUST").Value ' $I$153
    End If
    If .Range("SYSENG_TASK4_OVD_QTY").Value <> wsSource.Range("SYSENG_TASK4_OVD_QTY").Value Then
        .Range("SYSENG_TASK4_OVD_QTY").Value = wsSource.Range("SYSENG_TASK4_OVD_QTY").Value ' $H$153
    End If
    If .Range("SYSENG_TASK4_REM_COUNTRY").Value <> wsSource.Range("SYSENG_TASK4_REM_COUNTRY").Value Then
        .Range("SYSENG_TASK4_REM_COUNTRY").Value = wsSource.Range("SYSENG_TASK4_REM_COUNTRY").Value ' $M$153
    End If
    If .Range("SYSENG_TASK4_REM_PCT").Value <> wsSource.Range("SYSENG_TASK4_REM_PCT").Value Then
        .Range("SYSENG_TASK4_REM_PCT").Value = wsSource.Range("SYSENG_TASK4_REM_PCT").Value ' $N$153
    End If
    If .Range("SYSENG_TASK5_OVD_JUST").Value <> wsSource.Range("SYSENG_TASK5_OVD_JUST").Value Then
        .Range("SYSENG_TASK5_OVD_JUST").Value = wsSource.Range("SYSENG_TASK5_OVD_JUST").Value ' $I$154
    End If
    If .Range("SYSENG_TASK5_OVD_QTY").Value <> wsSource.Range("SYSENG_TASK5_OVD_QTY").Value Then
        .Range("SYSENG_TASK5_OVD_QTY").Value = wsSource.Range("SYSENG_TASK5_OVD_QTY").Value ' $H$154
    End If
    If .Range("SYSENG_TASK5_REM_COUNTRY").Value <> wsSource.Range("SYSENG_TASK5_REM_COUNTRY").Value Then
        .Range("SYSENG_TASK5_REM_COUNTRY").Value = wsSource.Range("SYSENG_TASK5_REM_COUNTRY").Value ' $M$154
    End If
    If .Range("SYSENG_TASK5_REM_PCT").Value <> wsSource.Range("SYSENG_TASK5_REM_PCT").Value Then
        .Range("SYSENG_TASK5_REM_PCT").Value = wsSource.Range("SYSENG_TASK5_REM_PCT").Value ' $N$154
    End If
    If .Range("SYSENG_TASK6_OVD_JUST").Value <> wsSource.Range("SYSENG_TASK6_OVD_JUST").Value Then
        .Range("SYSENG_TASK6_OVD_JUST").Value = wsSource.Range("SYSENG_TASK6_OVD_JUST").Value ' $I$155
    End If
    If .Range("SYSENG_TASK6_OVD_QTY").Value <> wsSource.Range("SYSENG_TASK6_OVD_QTY").Value Then
        .Range("SYSENG_TASK6_OVD_QTY").Value = wsSource.Range("SYSENG_TASK6_OVD_QTY").Value ' $H$155
    End If
    If .Range("SYSENG_TASK6_REM_COUNTRY").Value <> wsSource.Range("SYSENG_TASK6_REM_COUNTRY").Value Then
        .Range("SYSENG_TASK6_REM_COUNTRY").Value = wsSource.Range("SYSENG_TASK6_REM_COUNTRY").Value ' $M$155
    End If
    If .Range("SYSENG_TASK6_REM_PCT").Value <> wsSource.Range("SYSENG_TASK6_REM_PCT").Value Then
        .Range("SYSENG_TASK6_REM_PCT").Value = wsSource.Range("SYSENG_TASK6_REM_PCT").Value ' $N$155
    End If
    If .Range("SYSENG_TASK7_OVD_JUST").Value <> wsSource.Range("SYSENG_TASK7_OVD_JUST").Value Then
        .Range("SYSENG_TASK7_OVD_JUST").Value = wsSource.Range("SYSENG_TASK7_OVD_JUST").Value ' $I$156
    End If
    If .Range("SYSENG_TASK7_OVD_QTY").Value <> wsSource.Range("SYSENG_TASK7_OVD_QTY").Value Then
        .Range("SYSENG_TASK7_OVD_QTY").Value = wsSource.Range("SYSENG_TASK7_OVD_QTY").Value ' $H$156
    End If
    If .Range("SYSENG_TASK7_REM_COUNTRY").Value <> wsSource.Range("SYSENG_TASK7_REM_COUNTRY").Value Then
        .Range("SYSENG_TASK7_REM_COUNTRY").Value = wsSource.Range("SYSENG_TASK7_REM_COUNTRY").Value ' $M$156
    End If
    If .Range("SYSENG_TASK7_REM_PCT").Value <> wsSource.Range("SYSENG_TASK7_REM_PCT").Value Then
        .Range("SYSENG_TASK7_REM_PCT").Value = wsSource.Range("SYSENG_TASK7_REM_PCT").Value ' $N$156
    End If
    If .Range("SYSENG_TASK8_OVD_JUST").Value <> wsSource.Range("SYSENG_TASK8_OVD_JUST").Value Then
        .Range("SYSENG_TASK8_OVD_JUST").Value = wsSource.Range("SYSENG_TASK8_OVD_JUST").Value ' $I$157
    End If
    If .Range("SYSENG_TASK8_OVD_QTY").Value <> wsSource.Range("SYSENG_TASK8_OVD_QTY").Value Then
        .Range("SYSENG_TASK8_OVD_QTY").Value = wsSource.Range("SYSENG_TASK8_OVD_QTY").Value ' $H$157
    End If
    If .Range("SYSENG_TASK8_REM_COUNTRY").Value <> wsSource.Range("SYSENG_TASK8_REM_COUNTRY").Value Then
        .Range("SYSENG_TASK8_REM_COUNTRY").Value = wsSource.Range("SYSENG_TASK8_REM_COUNTRY").Value ' $M$157
    End If
    If .Range("SYSENG_TASK8_REM_PCT").Value <> wsSource.Range("SYSENG_TASK8_REM_PCT").Value Then
        .Range("SYSENG_TASK8_REM_PCT").Value = wsSource.Range("SYSENG_TASK8_REM_PCT").Value ' $N$157
    End If
    If .Range("SYSENG_TASK9_OVD_JUST").Value <> wsSource.Range("SYSENG_TASK9_OVD_JUST").Value Then
        .Range("SYSENG_TASK9_OVD_JUST").Value = wsSource.Range("SYSENG_TASK9_OVD_JUST").Value ' $I$158
    End If
    If .Range("SYSENG_TASK9_OVD_QTY").Value <> wsSource.Range("SYSENG_TASK9_OVD_QTY").Value Then
        .Range("SYSENG_TASK9_OVD_QTY").Value = wsSource.Range("SYSENG_TASK9_OVD_QTY").Value ' $H$158
    End If
    If .Range("SYSENG_TASK9_REM_COUNTRY").Value <> wsSource.Range("SYSENG_TASK9_REM_COUNTRY").Value Then
        .Range("SYSENG_TASK9_REM_COUNTRY").Value = wsSource.Range("SYSENG_TASK9_REM_COUNTRY").Value ' $M$158
    End If
    If .Range("SYSENG_TASK9_REM_PCT").Value <> wsSource.Range("SYSENG_TASK9_REM_PCT").Value Then
        .Range("SYSENG_TASK9_REM_PCT").Value = wsSource.Range("SYSENG_TASK9_REM_PCT").Value ' $N$158
    End If
    If .Range("TEST_TASK1_OVD_JUST").Value <> wsSource.Range("TEST_TASK1_OVD_JUST").Value Then
        .Range("TEST_TASK1_OVD_JUST").Value = wsSource.Range("TEST_TASK1_OVD_JUST").Value ' $I$301
    End If
    If .Range("TEST_TASK1_OVD_QTY").Value <> wsSource.Range("TEST_TASK1_OVD_QTY").Value Then
        .Range("TEST_TASK1_OVD_QTY").Value = wsSource.Range("TEST_TASK1_OVD_QTY").Value ' $H$301
    End If
    If .Range("TEST_TASK1_REM_COUNTRY").Value <> wsSource.Range("TEST_TASK1_REM_COUNTRY").Value Then
        .Range("TEST_TASK1_REM_COUNTRY").Value = wsSource.Range("TEST_TASK1_REM_COUNTRY").Value ' $M$301
    End If
    If .Range("TEST_TASK1_REM_PCT").Value <> wsSource.Range("TEST_TASK1_REM_PCT").Value Then
        .Range("TEST_TASK1_REM_PCT").Value = wsSource.Range("TEST_TASK1_REM_PCT").Value ' $N$301
    End If
    If .Range("TEST_TASK10_OVD_JUST").Value <> wsSource.Range("TEST_TASK10_OVD_JUST").Value Then
        .Range("TEST_TASK10_OVD_JUST").Value = wsSource.Range("TEST_TASK10_OVD_JUST").Value ' $I$310
    End If
    If .Range("TEST_TASK10_OVD_QTY").Value <> wsSource.Range("TEST_TASK10_OVD_QTY").Value Then
        .Range("TEST_TASK10_OVD_QTY").Value = wsSource.Range("TEST_TASK10_OVD_QTY").Value ' $H$310
    End If
    If .Range("TEST_TASK10_REM_COUNTRY").Value <> wsSource.Range("TEST_TASK10_REM_COUNTRY").Value Then
        .Range("TEST_TASK10_REM_COUNTRY").Value = wsSource.Range("TEST_TASK10_REM_COUNTRY").Value ' $M$310
    End If
    If .Range("TEST_TASK10_REM_PCT").Value <> wsSource.Range("TEST_TASK10_REM_PCT").Value Then
        .Range("TEST_TASK10_REM_PCT").Value = wsSource.Range("TEST_TASK10_REM_PCT").Value ' $N$310
    End If
    If .Range("TEST_TASK2_OVD_JUST").Value <> wsSource.Range("TEST_TASK2_OVD_JUST").Value Then
        .Range("TEST_TASK2_OVD_JUST").Value = wsSource.Range("TEST_TASK2_OVD_JUST").Value ' $I$302
    End If
    If .Range("TEST_TASK2_OVD_QTY").Value <> wsSource.Range("TEST_TASK2_OVD_QTY").Value Then
        .Range("TEST_TASK2_OVD_QTY").Value = wsSource.Range("TEST_TASK2_OVD_QTY").Value ' $H$302
    End If
    If .Range("TEST_TASK2_REM_COUNTRY").Value <> wsSource.Range("TEST_TASK2_REM_COUNTRY").Value Then
        .Range("TEST_TASK2_REM_COUNTRY").Value = wsSource.Range("TEST_TASK2_REM_COUNTRY").Value ' $M$302
    End If
    If .Range("TEST_TASK2_REM_PCT").Value <> wsSource.Range("TEST_TASK2_REM_PCT").Value Then
        .Range("TEST_TASK2_REM_PCT").Value = wsSource.Range("TEST_TASK2_REM_PCT").Value ' $N$302
    End If
    If .Range("TEST_TASK3_OVD_JUST").Value <> wsSource.Range("TEST_TASK3_OVD_JUST").Value Then
        .Range("TEST_TASK3_OVD_JUST").Value = wsSource.Range("TEST_TASK3_OVD_JUST").Value ' $I$303
    End If
    If .Range("TEST_TASK3_OVD_QTY").Value <> wsSource.Range("TEST_TASK3_OVD_QTY").Value Then
        .Range("TEST_TASK3_OVD_QTY").Value = wsSource.Range("TEST_TASK3_OVD_QTY").Value ' $H$303
    End If
    If .Range("TEST_TASK3_REM_COUNTRY").Value <> wsSource.Range("TEST_TASK3_REM_COUNTRY").Value Then
        .Range("TEST_TASK3_REM_COUNTRY").Value = wsSource.Range("TEST_TASK3_REM_COUNTRY").Value ' $M$303
    End If
    If .Range("TEST_TASK3_REM_PCT").Value <> wsSource.Range("TEST_TASK3_REM_PCT").Value Then
        .Range("TEST_TASK3_REM_PCT").Value = wsSource.Range("TEST_TASK3_REM_PCT").Value ' $N$303
    End If
    If .Range("TEST_TASK4_OVD_JUST").Value <> wsSource.Range("TEST_TASK4_OVD_JUST").Value Then
        .Range("TEST_TASK4_OVD_JUST").Value = wsSource.Range("TEST_TASK4_OVD_JUST").Value ' $I$304
    End If
    If .Range("TEST_TASK4_OVD_QTY").Value <> wsSource.Range("TEST_TASK4_OVD_QTY").Value Then
        .Range("TEST_TASK4_OVD_QTY").Value = wsSource.Range("TEST_TASK4_OVD_QTY").Value ' $H$304
    End If
    If .Range("TEST_TASK4_REM_COUNTRY").Value <> wsSource.Range("TEST_TASK4_REM_COUNTRY").Value Then
        .Range("TEST_TASK4_REM_COUNTRY").Value = wsSource.Range("TEST_TASK4_REM_COUNTRY").Value ' $M$304
    End If
    If .Range("TEST_TASK4_REM_PCT").Value <> wsSource.Range("TEST_TASK4_REM_PCT").Value Then
        .Range("TEST_TASK4_REM_PCT").Value = wsSource.Range("TEST_TASK4_REM_PCT").Value ' $N$304
    End If
    If .Range("TEST_TASK5_OVD_JUST").Value <> wsSource.Range("TEST_TASK5_OVD_JUST").Value Then
        .Range("TEST_TASK5_OVD_JUST").Value = wsSource.Range("TEST_TASK5_OVD_JUST").Value ' $I$305
    End If
    If .Range("TEST_TASK5_OVD_QTY").Value <> wsSource.Range("TEST_TASK5_OVD_QTY").Value Then
        .Range("TEST_TASK5_OVD_QTY").Value = wsSource.Range("TEST_TASK5_OVD_QTY").Value ' $H$305
    End If
    If .Range("TEST_TASK5_REM_COUNTRY").Value <> wsSource.Range("TEST_TASK5_REM_COUNTRY").Value Then
        .Range("TEST_TASK5_REM_COUNTRY").Value = wsSource.Range("TEST_TASK5_REM_COUNTRY").Value ' $M$305
    End If
    If .Range("TEST_TASK5_REM_PCT").Value <> wsSource.Range("TEST_TASK5_REM_PCT").Value Then
        .Range("TEST_TASK5_REM_PCT").Value = wsSource.Range("TEST_TASK5_REM_PCT").Value ' $N$305
    End If
    If .Range("TEST_TASK6_OVD_JUST").Value <> wsSource.Range("TEST_TASK6_OVD_JUST").Value Then
        .Range("TEST_TASK6_OVD_JUST").Value = wsSource.Range("TEST_TASK6_OVD_JUST").Value ' $I$306
    End If
    If .Range("TEST_TASK6_OVD_QTY").Value <> wsSource.Range("TEST_TASK6_OVD_QTY").Value Then
        .Range("TEST_TASK6_OVD_QTY").Value = wsSource.Range("TEST_TASK6_OVD_QTY").Value ' $H$306
    End If
    If .Range("TEST_TASK6_REM_COUNTRY").Value <> wsSource.Range("TEST_TASK6_REM_COUNTRY").Value Then
        .Range("TEST_TASK6_REM_COUNTRY").Value = wsSource.Range("TEST_TASK6_REM_COUNTRY").Value ' $M$306
    End If
    If .Range("TEST_TASK6_REM_PCT").Value <> wsSource.Range("TEST_TASK6_REM_PCT").Value Then
        .Range("TEST_TASK6_REM_PCT").Value = wsSource.Range("TEST_TASK6_REM_PCT").Value ' $N$306
    End If
    If .Range("TEST_TASK7_OVD_JUST").Value <> wsSource.Range("TEST_TASK7_OVD_JUST").Value Then
        .Range("TEST_TASK7_OVD_JUST").Value = wsSource.Range("TEST_TASK7_OVD_JUST").Value ' $I$307
    End If
    If .Range("TEST_TASK7_OVD_QTY").Value <> wsSource.Range("TEST_TASK7_OVD_QTY").Value Then
        .Range("TEST_TASK7_OVD_QTY").Value = wsSource.Range("TEST_TASK7_OVD_QTY").Value ' $H$307
    End If
    If .Range("TEST_TASK7_REM_COUNTRY").Value <> wsSource.Range("TEST_TASK7_REM_COUNTRY").Value Then
        .Range("TEST_TASK7_REM_COUNTRY").Value = wsSource.Range("TEST_TASK7_REM_COUNTRY").Value ' $M$307
    End If
    If .Range("TEST_TASK7_REM_PCT").Value <> wsSource.Range("TEST_TASK7_REM_PCT").Value Then
        .Range("TEST_TASK7_REM_PCT").Value = wsSource.Range("TEST_TASK7_REM_PCT").Value ' $N$307
    End If
    If .Range("TEST_TASK8_OVD_JUST").Value <> wsSource.Range("TEST_TASK8_OVD_JUST").Value Then
        .Range("TEST_TASK8_OVD_JUST").Value = wsSource.Range("TEST_TASK8_OVD_JUST").Value ' $I$308
    End If
    If .Range("TEST_TASK8_OVD_QTY").Value <> wsSource.Range("TEST_TASK8_OVD_QTY").Value Then
        .Range("TEST_TASK8_OVD_QTY").Value = wsSource.Range("TEST_TASK8_OVD_QTY").Value ' $H$308
    End If
    If .Range("TEST_TASK8_REM_COUNTRY").Value <> wsSource.Range("TEST_TASK8_REM_COUNTRY").Value Then
        .Range("TEST_TASK8_REM_COUNTRY").Value = wsSource.Range("TEST_TASK8_REM_COUNTRY").Value ' $M$308
    End If
    If .Range("TEST_TASK8_REM_PCT").Value <> wsSource.Range("TEST_TASK8_REM_PCT").Value Then
        .Range("TEST_TASK8_REM_PCT").Value = wsSource.Range("TEST_TASK8_REM_PCT").Value ' $N$308
    End If
    If .Range("TEST_TASK9_OVD_JUST").Value <> wsSource.Range("TEST_TASK9_OVD_JUST").Value Then
        .Range("TEST_TASK9_OVD_JUST").Value = wsSource.Range("TEST_TASK9_OVD_JUST").Value ' $I$309
    End If
    If .Range("TEST_TASK9_OVD_QTY").Value <> wsSource.Range("TEST_TASK9_OVD_QTY").Value Then
        .Range("TEST_TASK9_OVD_QTY").Value = wsSource.Range("TEST_TASK9_OVD_QTY").Value ' $H$309
    End If
    If .Range("TEST_TASK9_REM_COUNTRY").Value <> wsSource.Range("TEST_TASK9_REM_COUNTRY").Value Then
        .Range("TEST_TASK9_REM_COUNTRY").Value = wsSource.Range("TEST_TASK9_REM_COUNTRY").Value ' $M$309
    End If
    If .Range("TEST_TASK9_REM_PCT").Value <> wsSource.Range("TEST_TASK9_REM_PCT").Value Then
        .Range("TEST_TASK9_REM_PCT").Value = wsSource.Range("TEST_TASK9_REM_PCT").Value ' $N$309
    End If
    If .Range("TL_TASK1_OVD_JUST").Value <> wsSource.Range("TL_TASK1_OVD_JUST").Value Then
        .Range("TL_TASK1_OVD_JUST").Value = wsSource.Range("TL_TASK1_OVD_JUST").Value ' $I$426
    End If
    If .Range("TL_TASK1_OVD_QTY").Value <> wsSource.Range("TL_TASK1_OVD_QTY").Value Then
        .Range("TL_TASK1_OVD_QTY").Value = wsSource.Range("TL_TASK1_OVD_QTY").Value ' $H$426
    End If
    If .Range("TL_TASK10_OVD_JUST").Value <> wsSource.Range("TL_TASK10_OVD_JUST").Value Then
        .Range("TL_TASK10_OVD_JUST").Value = wsSource.Range("TL_TASK10_OVD_JUST").Value ' $I$435
    End If
    If .Range("TL_TASK10_OVD_QTY").Value <> wsSource.Range("TL_TASK10_OVD_QTY").Value Then
        .Range("TL_TASK10_OVD_QTY").Value = wsSource.Range("TL_TASK10_OVD_QTY").Value ' $H$435
    End If
    If .Range("TL_TASK2_OVD_JUST").Value <> wsSource.Range("TL_TASK2_OVD_JUST").Value Then
        .Range("TL_TASK2_OVD_JUST").Value = wsSource.Range("TL_TASK2_OVD_JUST").Value ' $I$427
    End If
    If .Range("TL_TASK2_OVD_QTY").Value <> wsSource.Range("TL_TASK2_OVD_QTY").Value Then
        .Range("TL_TASK2_OVD_QTY").Value = wsSource.Range("TL_TASK2_OVD_QTY").Value ' $H$427
    End If
    If .Range("TL_TASK3_OVD_JUST").Value <> wsSource.Range("TL_TASK3_OVD_JUST").Value Then
        .Range("TL_TASK3_OVD_JUST").Value = wsSource.Range("TL_TASK3_OVD_JUST").Value ' $I$428
    End If
    If .Range("TL_TASK3_OVD_QTY").Value <> wsSource.Range("TL_TASK3_OVD_QTY").Value Then
        .Range("TL_TASK3_OVD_QTY").Value = wsSource.Range("TL_TASK3_OVD_QTY").Value ' $H$428
    End If
    If .Range("TL_TASK4_OVD_JUST").Value <> wsSource.Range("TL_TASK4_OVD_JUST").Value Then
        .Range("TL_TASK4_OVD_JUST").Value = wsSource.Range("TL_TASK4_OVD_JUST").Value ' $I$429
    End If
    If .Range("TL_TASK4_OVD_QTY").Value <> wsSource.Range("TL_TASK4_OVD_QTY").Value Then
        .Range("TL_TASK4_OVD_QTY").Value = wsSource.Range("TL_TASK4_OVD_QTY").Value ' $H$429
    End If
    If .Range("TL_TASK5_OVD_JUST").Value <> wsSource.Range("TL_TASK5_OVD_JUST").Value Then
        .Range("TL_TASK5_OVD_JUST").Value = wsSource.Range("TL_TASK5_OVD_JUST").Value ' $I$430
    End If
    If .Range("TL_TASK5_OVD_QTY").Value <> wsSource.Range("TL_TASK5_OVD_QTY").Value Then
        .Range("TL_TASK5_OVD_QTY").Value = wsSource.Range("TL_TASK5_OVD_QTY").Value ' $H$430
    End If
    If .Range("TL_TASK6_OVD_JUST").Value <> wsSource.Range("TL_TASK6_OVD_JUST").Value Then
        .Range("TL_TASK6_OVD_JUST").Value = wsSource.Range("TL_TASK6_OVD_JUST").Value ' $I$431
    End If
    If .Range("TL_TASK6_OVD_QTY").Value <> wsSource.Range("TL_TASK6_OVD_QTY").Value Then
        .Range("TL_TASK6_OVD_QTY").Value = wsSource.Range("TL_TASK6_OVD_QTY").Value ' $H$431
    End If
    If .Range("TL_TASK7_OVD_JUST").Value <> wsSource.Range("TL_TASK7_OVD_JUST").Value Then
        .Range("TL_TASK7_OVD_JUST").Value = wsSource.Range("TL_TASK7_OVD_JUST").Value ' $I$432
    End If
    If .Range("TL_TASK7_OVD_QTY").Value <> wsSource.Range("TL_TASK7_OVD_QTY").Value Then
        .Range("TL_TASK7_OVD_QTY").Value = wsSource.Range("TL_TASK7_OVD_QTY").Value ' $H$432
    End If
    If .Range("TL_TASK8_OVD_JUST").Value <> wsSource.Range("TL_TASK8_OVD_JUST").Value Then
        .Range("TL_TASK8_OVD_JUST").Value = wsSource.Range("TL_TASK8_OVD_JUST").Value ' $I$433
    End If
    If .Range("TL_TASK8_OVD_QTY").Value <> wsSource.Range("TL_TASK8_OVD_QTY").Value Then
        .Range("TL_TASK8_OVD_QTY").Value = wsSource.Range("TL_TASK8_OVD_QTY").Value ' $H$433
    End If
    If .Range("TL_TASK9_OVD_JUST").Value <> wsSource.Range("TL_TASK9_OVD_JUST").Value Then
        .Range("TL_TASK9_OVD_JUST").Value = wsSource.Range("TL_TASK9_OVD_JUST").Value ' $I$434
    End If
    If .Range("TL_TASK9_OVD_QTY").Value <> wsSource.Range("TL_TASK9_OVD_QTY").Value Then
        .Range("TL_TASK9_OVD_QTY").Value = wsSource.Range("TL_TASK9_OVD_QTY").Value ' $H$434
    End If
    
    
    
    
    
 'Formula: =G242
    .Range("DUP_FACT_APP_BASED_APP").Value = wsSource.Range("DUP_FACT_APP_BASED_APP").Value ' $F$242
    .Range("DUP_FACT_APP_BASED_COURSE").Value = wsSource.Range("DUP_FACT_APP_BASED_COURSE").Value ' $F$300
    .Range("DUP_FACT_APP_BASED_CP").Value = wsSource.Range("DUP_FACT_APP_BASED_CP").Value ' $F$169
    .Range("DUP_FACT_APP_BASED_DI").Value = wsSource.Range("DUP_FACT_APP_BASED_DI").Value ' $F$189
    .Range("DUP_FACT_APP_BASED_DOC").Value = wsSource.Range("DUP_FACT_APP_BASED_DOC").Value ' $F$279
    .Range("DUP_FACT_APP_BASED_ESD").Value = wsSource.Range("DUP_FACT_APP_BASED_ESD").Value ' $F$209
    .Range("DUP_FACT_APP_BASED_HMI").Value = wsSource.Range("DUP_FACT_APP_BASED_HMI").Value ' $F$147
    .Range("DUP_FACT_APP_BASED_MEETING").Value = wsSource.Range("DUP_FACT_APP_BASED_MEETING").Value ' $F$330
    .Range("DUP_FACT_APP_BASED_PM").Value = wsSource.Range("DUP_FACT_APP_BASED_PM").Value ' $F$316
    .Range("DUP_FACT_APP_BASED_REP").Value = wsSource.Range("DUP_FACT_APP_BASED_REP").Value ' $F$228
    .Range("DUP_FACT_APP_BASED_SITE").Value = wsSource.Range("DUP_FACT_APP_BASED_SITE").Value ' $F$346
    .Range("DUP_FACT_APP_BASED_SPEC").Value = wsSource.Range("DUP_FACT_APP_BASED_SPEC").Value ' $F$119
    .Range("DUP_FACT_APP_BASED_SYS_ENG").Value = wsSource.Range("DUP_FACT_APP_BASED_SYS_ENG").Value ' $F$134
    .Range("DUP_FACT_APP_BASED_SYSENG").Value = wsSource.Range("DUP_FACT_APP_BASED_SYSENG").Value ' $F$134
    .Range("DUP_FACT_APP_BASED_TEST").Value = wsSource.Range("DUP_FACT_APP_BASED_TEST").Value ' $F$262
    
    
    
 'Formula: =HMI_RPT_PCT_QTY_TUNED
    .Range("HMI_RPT_PCT_QTY").Value = wsSource.Range("HMI_RPT_PCT_QTY").Value ' $E$106
    .Range("IO_OPTIMISATION_EXP_RATE").Value = wsSource.Range("IO_OPTIMISATION_EXP_RATE").Value ' $E$106
    .Range("IO_OPTIMISATION_MAX_REDUCTION").Value = wsSource.Range("IO_OPTIMISATION_MAX_REDUCTION").Value ' $E$106
    .Range("PM_PCT_TOTAL").Value = wsSource.Range("PM_PCT_TOTAL").Value ' $E$106
    .Range("PM_IA_TOTAL").Value = wsSource.Range("PM_IA_TOTAL").Value ' $E$106
    .Range("PM_BUYOUT_TOTAL").Value = wsSource.Range("PM_BUYOUT_TOTAL").Value ' $E$106
    
End With
 '387

End Sub

Sub ImportApplicationBasedSheet(wsSource As Worksheet, wsTarget As Worksheet)
'GECE Version = 1.0
On Error Resume Next
With wsTarget
    If .Range("APP_TASK1_OVD_JUST").Value <> wsSource.Range("APP_1_OVD_JUST").Value Then
        .Range("APP_TASK1_OVD_JUST").Value = wsSource.Range("APP_1_OVD_JUST").Value ' $I$251
    End If
    If .Range("APP_TASK1_OVD_QTY").Value <> wsSource.Range("APP_1_OVD_QTY").Value Then
        .Range("APP_TASK1_OVD_QTY").Value = wsSource.Range("APP_1_OVD_QTY").Value ' $H$251
    End If
    If .Range("APP_TASK1_REM_COUNTRY").Value <> wsSource.Range("APP_1_REM_COUNTRY").Value Then
        .Range("APP_TASK1_REM_COUNTRY").Value = wsSource.Range("APP_1_REM_COUNTRY").Value ' $M$251
    End If
    If .Range("APP_TASK1_REM_PCT").Value <> wsSource.Range("APP_1_REM_PCT").Value Then
        .Range("APP_TASK1_REM_PCT").Value = wsSource.Range("APP_1_REM_PCT").Value ' $N$251
    End If
    If .Range("APP_TASK2_OVD_JUST").Value <> wsSource.Range("APP_2_OVD_JUST").Value Then
        .Range("APP_TASK2_OVD_JUST").Value = wsSource.Range("APP_2_OVD_JUST").Value ' $I$252
    End If
    If .Range("APP_TASK2_OVD_QTY").Value <> wsSource.Range("APP_2_OVD_QTY").Value Then
        .Range("APP_TASK2_OVD_QTY").Value = wsSource.Range("APP_2_OVD_QTY").Value ' $H$252
    End If
    If .Range("APP_TASK2_REM_COUNTRY").Value <> wsSource.Range("APP_2_REM_COUNTRY").Value Then
        .Range("APP_TASK2_REM_COUNTRY").Value = wsSource.Range("APP_2_REM_COUNTRY").Value ' $M$252
    End If
    If .Range("APP_TASK2_REM_PCT").Value <> wsSource.Range("APP_2_REM_PCT").Value Then
        .Range("APP_TASK2_REM_PCT").Value = wsSource.Range("APP_2_REM_PCT").Value ' $N$252
    End If
    If .Range("APP_TASK3_OVD_JUST").Value <> wsSource.Range("APP_3_OVD_JUST").Value Then
        .Range("APP_TASK3_OVD_JUST").Value = wsSource.Range("APP_3_OVD_JUST").Value ' $I$253
    End If
    If .Range("APP_TASK3_OVD_QTY").Value <> wsSource.Range("APP_3_OVD_QTY").Value Then
        .Range("APP_TASK3_OVD_QTY").Value = wsSource.Range("APP_3_OVD_QTY").Value ' $H$253
    End If
    If .Range("APP_TASK3_REM_COUNTRY").Value <> wsSource.Range("APP_3_REM_COUNTRY").Value Then
        .Range("APP_TASK3_REM_COUNTRY").Value = wsSource.Range("APP_3_REM_COUNTRY").Value ' $M$253
    End If
    If .Range("APP_TASK3_REM_PCT").Value <> wsSource.Range("APP_3_REM_PCT").Value Then
        .Range("APP_TASK3_REM_PCT").Value = wsSource.Range("APP_3_REM_PCT").Value ' $N$253
    End If
    If .Range("APP_TASK4_OVD_JUST").Value <> wsSource.Range("APP_4_OVD_JUST").Value Then
        .Range("APP_TASK4_OVD_JUST").Value = wsSource.Range("APP_4_OVD_JUST").Value ' $I$254
    End If
    If .Range("APP_TASK4_OVD_QTY").Value <> wsSource.Range("APP_4_OVD_QTY").Value Then
        .Range("APP_TASK4_OVD_QTY").Value = wsSource.Range("APP_4_OVD_QTY").Value ' $H$254
    End If
    If .Range("APP_TASK4_REM_COUNTRY").Value <> wsSource.Range("APP_4_REM_COUNTRY").Value Then
        .Range("APP_TASK4_REM_COUNTRY").Value = wsSource.Range("APP_4_REM_COUNTRY").Value ' $M$254
    End If
    If .Range("APP_TASK4_REM_PCT").Value <> wsSource.Range("APP_4_REM_PCT").Value Then
        .Range("APP_TASK4_REM_PCT").Value = wsSource.Range("APP_4_REM_PCT").Value ' $N$254
    End If
    If .Range("APP_TASK5_OVD_JUST").Value <> wsSource.Range("APP_5_OVD_JUST").Value Then
        .Range("APP_TASK5_OVD_JUST").Value = wsSource.Range("APP_5_OVD_JUST").Value ' $I$255
    End If
    If .Range("APP_TASK5_OVD_QTY").Value <> wsSource.Range("APP_5_OVD_QTY").Value Then
        .Range("APP_TASK5_OVD_QTY").Value = wsSource.Range("APP_5_OVD_QTY").Value ' $H$255
    End If
    If .Range("APP_TASK5_REM_COUNTRY").Value <> wsSource.Range("APP_5_REM_COUNTRY").Value Then
        .Range("APP_TASK5_REM_COUNTRY").Value = wsSource.Range("APP_5_REM_COUNTRY").Value ' $M$255
    End If
    If .Range("APP_TASK5_REM_PCT").Value <> wsSource.Range("APP_5_REM_PCT").Value Then
        .Range("APP_TASK5_REM_PCT").Value = wsSource.Range("APP_5_REM_PCT").Value ' $N$255
    End If
    If .Range("APP_TASK6_OVD_JUST").Value <> wsSource.Range("APP_6_OVD_JUST").Value Then
        .Range("APP_TASK6_OVD_JUST").Value = wsSource.Range("APP_6_OVD_JUST").Value ' $I$256
    End If
    If .Range("APP_TASK6_OVD_QTY").Value <> wsSource.Range("APP_6_OVD_QTY").Value Then
        .Range("APP_TASK6_OVD_QTY").Value = wsSource.Range("APP_6_OVD_QTY").Value ' $H$256
    End If
    If .Range("APP_TASK6_REM_COUNTRY").Value <> wsSource.Range("APP_6_REM_COUNTRY").Value Then
        .Range("APP_TASK6_REM_COUNTRY").Value = wsSource.Range("APP_6_REM_COUNTRY").Value ' $M$256
    End If
    If .Range("APP_TASK6_REM_PCT").Value <> wsSource.Range("APP_6_REM_PCT").Value Then
        .Range("APP_TASK6_REM_PCT").Value = wsSource.Range("APP_6_REM_PCT").Value ' $N$256
    End If
    If .Range("APP_TASK7_OVD_JUST").Value <> wsSource.Range("APP_7_OVD_JUST").Value Then
        .Range("APP_TASK7_OVD_JUST").Value = wsSource.Range("APP_7_OVD_JUST").Value ' $I$257
    End If
    If .Range("APP_TASK7_OVD_QTY").Value <> wsSource.Range("APP_7_OVD_QTY").Value Then
        .Range("APP_TASK7_OVD_QTY").Value = wsSource.Range("APP_7_OVD_QTY").Value ' $H$257
    End If
    If .Range("APP_TASK7_REM_COUNTRY").Value <> wsSource.Range("APP_7_REM_COUNTRY").Value Then
        .Range("APP_TASK7_REM_COUNTRY").Value = wsSource.Range("APP_7_REM_COUNTRY").Value ' $M$257
    End If
    If .Range("APP_TASK7_REM_PCT").Value <> wsSource.Range("APP_7_REM_PCT").Value Then
        .Range("APP_TASK7_REM_PCT").Value = wsSource.Range("APP_7_REM_PCT").Value ' $N$257
    End If
    If .Range("APP_TASK8_OVD_JUST").Value <> wsSource.Range("APP_8_OVD_JUST").Value Then
        .Range("APP_TASK8_OVD_JUST").Value = wsSource.Range("APP_8_OVD_JUST").Value ' $I$258
    End If
    If .Range("APP_TASK8_OVD_QTY").Value <> wsSource.Range("APP_8_OVD_QTY").Value Then
        .Range("APP_TASK8_OVD_QTY").Value = wsSource.Range("APP_8_OVD_QTY").Value ' $H$258
    End If
    If .Range("APP_TASK8_REM_COUNTRY").Value <> wsSource.Range("APP_8_REM_COUNTRY").Value Then
        .Range("APP_TASK8_REM_COUNTRY").Value = wsSource.Range("APP_8_REM_COUNTRY").Value ' $M$258
    End If
    If .Range("APP_TASK8_REM_PCT").Value <> wsSource.Range("APP_8_REM_PCT").Value Then
        .Range("APP_TASK8_REM_PCT").Value = wsSource.Range("APP_8_REM_PCT").Value ' $N$258
    End If
    If .Range("APP_TASK9_OVD_JUST").Value <> wsSource.Range("APP_BUS_OVD_JUST").Value Then
        .Range("APP_TASK9_OVD_JUST").Value = wsSource.Range("APP_BUS_OVD_JUST").Value ' $I$259
    End If
    If .Range("APP_TASK9_OVD_QTY").Value <> wsSource.Range("APP_BUS_OVD_QTY").Value Then
        .Range("APP_TASK9_OVD_QTY").Value = wsSource.Range("APP_BUS_OVD_QTY").Value ' $H$259
    End If
    If .Range("APP_TASK9_REM_COUNTRY").Value <> wsSource.Range("APP_BUS_REM_COUNTRY").Value Then
        .Range("APP_TASK9_REM_COUNTRY").Value = wsSource.Range("APP_BUS_REM_COUNTRY").Value ' $M$259
    End If
    If .Range("APP_TASK9_REM_PCT").Value <> wsSource.Range("APP_BUS_REM_PCT").Value Then
        .Range("APP_TASK9_REM_PCT").Value = wsSource.Range("APP_BUS_REM_PCT").Value ' $N$259
    End If
    If .Range("COURSE_TASK1_OVD_JUST").Value <> wsSource.Range("COURSE_1_OVD_JUST").Value Then
        .Range("COURSE_TASK1_OVD_JUST").Value = wsSource.Range("COURSE_1_OVD_JUST").Value ' $I$309
    End If
    If .Range("COURSE_TASK1_OVD_QTY").Value <> wsSource.Range("COURSE_1_OVD_QTY").Value Then
        .Range("COURSE_TASK1_OVD_QTY").Value = wsSource.Range("COURSE_1_OVD_QTY").Value ' $H$309
    End If
    If .Range("COURSE_TASK1_REM_COUNTRY").Value <> wsSource.Range("COURSE_1_REM_COUNTRY").Value Then
        .Range("COURSE_TASK1_REM_COUNTRY").Value = wsSource.Range("COURSE_1_REM_COUNTRY").Value ' $M$309
    End If
    If .Range("COURSE_TASK1_REM_PCT").Value <> wsSource.Range("COURSE_1_REM_PCT").Value Then
        .Range("COURSE_TASK1_REM_PCT").Value = wsSource.Range("COURSE_1_REM_PCT").Value ' $N$309
    End If
    If .Range("COURSE_TASK2_OVD_JUST").Value <> wsSource.Range("COURSE_2_OVD_JUST").Value Then
        .Range("COURSE_TASK2_OVD_JUST").Value = wsSource.Range("COURSE_2_OVD_JUST").Value ' $I$310
    End If
    If .Range("COURSE_TASK2_OVD_QTY").Value <> wsSource.Range("COURSE_2_OVD_QTY").Value Then
        .Range("COURSE_TASK2_OVD_QTY").Value = wsSource.Range("COURSE_2_OVD_QTY").Value ' $H$310
    End If
    If .Range("COURSE_TASK2_REM_COUNTRY").Value <> wsSource.Range("COURSE_2_REM_COUNTRY").Value Then
        .Range("COURSE_TASK2_REM_COUNTRY").Value = wsSource.Range("COURSE_2_REM_COUNTRY").Value ' $M$310
    End If
    If .Range("COURSE_TASK2_REM_PCT").Value <> wsSource.Range("COURSE_2_REM_PCT").Value Then
        .Range("COURSE_TASK2_REM_PCT").Value = wsSource.Range("COURSE_2_REM_PCT").Value ' $N$310
    End If
    If .Range("COURSE_TASK3_OVD_JUST").Value <> wsSource.Range("COURSE_3_OVD_JUST").Value Then
        .Range("COURSE_TASK3_OVD_JUST").Value = wsSource.Range("COURSE_3_OVD_JUST").Value ' $I$311
    End If
    If .Range("COURSE_TASK3_OVD_QTY").Value <> wsSource.Range("COURSE_3_OVD_QTY").Value Then
        .Range("COURSE_TASK3_OVD_QTY").Value = wsSource.Range("COURSE_3_OVD_QTY").Value ' $H$311
    End If
    If .Range("COURSE_TASK3_REM_COUNTRY").Value <> wsSource.Range("COURSE_3_REM_COUNTRY").Value Then
        .Range("COURSE_TASK3_REM_COUNTRY").Value = wsSource.Range("COURSE_3_REM_COUNTRY").Value ' $M$311
    End If
    If .Range("COURSE_TASK3_REM_PCT").Value <> wsSource.Range("COURSE_3_REM_PCT").Value Then
        .Range("COURSE_TASK3_REM_PCT").Value = wsSource.Range("COURSE_3_REM_PCT").Value ' $N$311
    End If
    If .Range("COURSE_TASK4_OVD_JUST").Value <> wsSource.Range("COURSE_4_OVD_JUST").Value Then
        .Range("COURSE_TASK4_OVD_JUST").Value = wsSource.Range("COURSE_4_OVD_JUST").Value ' $I$312
    End If
    If .Range("COURSE_TASK4_OVD_QTY").Value <> wsSource.Range("COURSE_4_OVD_QTY").Value Then
        .Range("COURSE_TASK4_OVD_QTY").Value = wsSource.Range("COURSE_4_OVD_QTY").Value ' $H$312
    End If
    If .Range("COURSE_TASK4_REM_COUNTRY").Value <> wsSource.Range("COURSE_4_REM_COUNTRY").Value Then
        .Range("COURSE_TASK4_REM_COUNTRY").Value = wsSource.Range("COURSE_4_REM_COUNTRY").Value ' $M$312
    End If
    If .Range("COURSE_TASK4_REM_PCT").Value <> wsSource.Range("COURSE_4_REM_PCT").Value Then
        .Range("COURSE_TASK4_REM_PCT").Value = wsSource.Range("COURSE_4_REM_PCT").Value ' $N$312
    End If
    If .Range("COURSE_TASK5_OVD_JUST").Value <> wsSource.Range("COURSE_5_OVD_JUST").Value Then
        .Range("COURSE_TASK5_OVD_JUST").Value = wsSource.Range("COURSE_5_OVD_JUST").Value ' $I$313
    End If
    If .Range("COURSE_TASK5_OVD_QTY").Value <> wsSource.Range("COURSE_5_OVD_QTY").Value Then
        .Range("COURSE_TASK5_OVD_QTY").Value = wsSource.Range("COURSE_5_OVD_QTY").Value ' $H$313
    End If
    If .Range("COURSE_TASK5_REM_COUNTRY").Value <> wsSource.Range("COURSE_5_REM_COUNTRY").Value Then
        .Range("COURSE_TASK5_REM_COUNTRY").Value = wsSource.Range("COURSE_5_REM_COUNTRY").Value ' $M$313
    End If
    If .Range("COURSE_TASK5_REM_PCT").Value <> wsSource.Range("COURSE_5_REM_PCT").Value Then
        .Range("COURSE_TASK5_REM_PCT").Value = wsSource.Range("COURSE_5_REM_PCT").Value ' $N$313
    End If
    If .Range("CP_TASK2_OVD_JUST").Value <> wsSource.Range("CP_AI_OVD_JUST").Value Then
        .Range("CP_TASK2_OVD_JUST").Value = wsSource.Range("CP_AI_OVD_JUST").Value ' $I$179
    End If
    If .Range("CP_TASK2_OVD_QTY").Value <> wsSource.Range("CP_AI_OVD_QTY").Value Then
        .Range("CP_TASK2_OVD_QTY").Value = wsSource.Range("CP_AI_OVD_QTY").Value ' $H$179
    End If
    If .Range("CP_TASK2_REM_COUNTRY").Value <> wsSource.Range("CP_AI_REM_COUNTRY").Value Then
        .Range("CP_TASK2_REM_COUNTRY").Value = wsSource.Range("CP_AI_REM_COUNTRY").Value ' $M$179
    End If
    If .Range("CP_TASK2_REM_PCT").Value <> wsSource.Range("CP_AI_REM_PCT").Value Then
        .Range("CP_TASK2_REM_PCT").Value = wsSource.Range("CP_AI_REM_PCT").Value ' $N$179
    End If
    If .Range("CP_TASK3_OVD_JUST").Value <> wsSource.Range("CP_AO_OVD_JUST").Value Then
        .Range("CP_TASK3_OVD_JUST").Value = wsSource.Range("CP_AO_OVD_JUST").Value ' $I$180
    End If
    If .Range("CP_TASK3_OVD_QTY").Value <> wsSource.Range("CP_AO_OVD_QTY").Value Then
        .Range("CP_TASK3_OVD_QTY").Value = wsSource.Range("CP_AO_OVD_QTY").Value ' $H$180
    End If
    If .Range("CP_TASK3_REM_COUNTRY").Value <> wsSource.Range("CP_AO_REM_COUNTRY").Value Then
        .Range("CP_TASK3_REM_COUNTRY").Value = wsSource.Range("CP_AO_REM_COUNTRY").Value ' $M$180
    End If
    If .Range("CP_TASK3_REM_PCT").Value <> wsSource.Range("CP_AO_REM_PCT").Value Then
        .Range("CP_TASK3_REM_PCT").Value = wsSource.Range("CP_AO_REM_PCT").Value ' $N$180
    End If
    If .Range("CP_TASK1_OVD_JUST").Value <> wsSource.Range("CP_DESIGN_OVD_JUST").Value Then
        .Range("CP_TASK1_OVD_JUST").Value = wsSource.Range("CP_DESIGN_OVD_JUST").Value ' $I$178
    End If
    If .Range("CP_TASK1_OVD_QTY").Value <> wsSource.Range("CP_DESIGN_OVD_QTY").Value Then
        .Range("CP_TASK1_OVD_QTY").Value = wsSource.Range("CP_DESIGN_OVD_QTY").Value ' $H$178
    End If
    If .Range("CP_TASK1_REM_COUNTRY").Value <> wsSource.Range("CP_DESIGN_REM_COUNTRY").Value Then
        .Range("CP_TASK1_REM_COUNTRY").Value = wsSource.Range("CP_DESIGN_REM_COUNTRY").Value ' $M$178
    End If
    If .Range("CP_TASK1_REM_PCT").Value <> wsSource.Range("CP_DESIGN_REM_PCT").Value Then
        .Range("CP_TASK1_REM_PCT").Value = wsSource.Range("CP_DESIGN_REM_PCT").Value ' $N$178
    End If
    If .Range("CP_TASK4_OVD_JUST").Value <> wsSource.Range("CP_DI_OVD_JUST").Value Then
        .Range("CP_TASK4_OVD_JUST").Value = wsSource.Range("CP_DI_OVD_JUST").Value ' $I$181
    End If
    If .Range("CP_TASK4_OVD_QTY").Value <> wsSource.Range("CP_DI_OVD_QTY").Value Then
        .Range("CP_TASK4_OVD_QTY").Value = wsSource.Range("CP_DI_OVD_QTY").Value ' $H$181
    End If
    If .Range("CP_TASK4_REM_COUNTRY").Value <> wsSource.Range("CP_DI_REM_COUNTRY").Value Then
        .Range("CP_TASK4_REM_COUNTRY").Value = wsSource.Range("CP_DI_REM_COUNTRY").Value ' $M$181
    End If
    If .Range("CP_TASK4_REM_PCT").Value <> wsSource.Range("CP_DI_REM_PCT").Value Then
        .Range("CP_TASK4_REM_PCT").Value = wsSource.Range("CP_DI_REM_PCT").Value ' $N$181
    End If
    If .Range("CP_TASK5_OVD_JUST").Value <> wsSource.Range("CP_DO_OVD_JUST").Value Then
        .Range("CP_TASK5_OVD_JUST").Value = wsSource.Range("CP_DO_OVD_JUST").Value ' $I$182
    End If
    If .Range("CP_TASK5_OVD_QTY").Value <> wsSource.Range("CP_DO_OVD_QTY").Value Then
        .Range("CP_TASK5_OVD_QTY").Value = wsSource.Range("CP_DO_OVD_QTY").Value ' $H$182
    End If
    If .Range("CP_TASK5_REM_COUNTRY").Value <> wsSource.Range("CP_DO_REM_COUNTRY").Value Then
        .Range("CP_TASK5_REM_COUNTRY").Value = wsSource.Range("CP_DO_REM_COUNTRY").Value ' $M$182
    End If
    If .Range("CP_TASK5_REM_PCT").Value <> wsSource.Range("CP_DO_REM_PCT").Value Then
        .Range("CP_TASK5_REM_PCT").Value = wsSource.Range("CP_DO_REM_PCT").Value ' $N$182
    End If
    If .Range("CP_TASK7_OVD_JUST").Value <> wsSource.Range("CP_GRP_START_OVD_JUST").Value Then
        .Range("CP_TASK7_OVD_JUST").Value = wsSource.Range("CP_GRP_START_OVD_JUST").Value ' $I$184
    End If
    If .Range("CP_TASK7_OVD_QTY").Value <> wsSource.Range("CP_GRP_START_OVD_QTY").Value Then
        .Range("CP_TASK7_OVD_QTY").Value = wsSource.Range("CP_GRP_START_OVD_QTY").Value ' $H$184
    End If
    If .Range("CP_TASK7_REM_COUNTRY").Value <> wsSource.Range("CP_GRP_START_REM_COUNTRY").Value Then
        .Range("CP_TASK7_REM_COUNTRY").Value = wsSource.Range("CP_GRP_START_REM_COUNTRY").Value ' $M$184
    End If
    If .Range("CP_TASK7_REM_PCT").Value <> wsSource.Range("CP_GRP_START_REM_PCT").Value Then
        .Range("CP_TASK7_REM_PCT").Value = wsSource.Range("CP_GRP_START_REM_PCT").Value ' $N$184
    End If
    If .Range("CP_TASK6_OVD_JUST").Value <> wsSource.Range("CP_LOGIC_OVD_JUST").Value Then
        .Range("CP_TASK6_OVD_JUST").Value = wsSource.Range("CP_LOGIC_OVD_JUST").Value ' $I$183
    End If
    If .Range("CP_TASK6_OVD_QTY").Value <> wsSource.Range("CP_LOGIC_OVD_QTY").Value Then
        .Range("CP_TASK6_OVD_QTY").Value = wsSource.Range("CP_LOGIC_OVD_QTY").Value ' $H$183
    End If
    If .Range("CP_TASK6_REM_COUNTRY").Value <> wsSource.Range("CP_LOGIC_REM_COUNTRY").Value Then
        .Range("CP_TASK6_REM_COUNTRY").Value = wsSource.Range("CP_LOGIC_REM_COUNTRY").Value ' $M$183
    End If
    If .Range("CP_TASK6_REM_PCT").Value <> wsSource.Range("CP_LOGIC_REM_PCT").Value Then
        .Range("CP_TASK6_REM_PCT").Value = wsSource.Range("CP_LOGIC_REM_PCT").Value ' $N$183
    End If
    If .Range("CP_TASK8_OVD_JUST").Value <> wsSource.Range("CP_SEQ_OVD_JUST").Value Then
        .Range("CP_TASK8_OVD_JUST").Value = wsSource.Range("CP_SEQ_OVD_JUST").Value ' $I$185
    End If
    If .Range("CP_TASK8_OVD_QTY").Value <> wsSource.Range("CP_SEQ_OVD_QTY").Value Then
        .Range("CP_TASK8_OVD_QTY").Value = wsSource.Range("CP_SEQ_OVD_QTY").Value ' $H$185
    End If
    If .Range("CP_TASK8_REM_COUNTRY").Value <> wsSource.Range("CP_SEQ_REM_COUNTRY").Value Then
        .Range("CP_TASK8_REM_COUNTRY").Value = wsSource.Range("CP_SEQ_REM_COUNTRY").Value ' $M$185
    End If
    If .Range("CP_TASK8_REM_PCT").Value <> wsSource.Range("CP_SEQ_REM_PCT").Value Then
        .Range("CP_TASK8_REM_PCT").Value = wsSource.Range("CP_SEQ_REM_PCT").Value ' $N$185
    End If
    
    'DI
    If .Range("DI_TASK2_OVD_JUST").Value <> wsSource.Range("DI_AI_OVD_JUST").Value Then
        .Range("DI_TASK2_OVD_JUST").Value = wsSource.Range("DI_AI_OVD_JUST").Value ' $I$199
    End If
    If .Range("DI_TASK2_OVD_QTY").Value <> wsSource.Range("DI_AI_OVD_QTY").Value Then
        .Range("DI_TASK2_OVD_QTY").Value = wsSource.Range("DI_AI_OVD_QTY").Value ' $H$199
    End If
    If .Range("DI_TASK2_REM_COUNTRY").Value <> wsSource.Range("DI_AI_REM_COUNTRY").Value Then
        .Range("DI_TASK2_REM_COUNTRY").Value = wsSource.Range("DI_AI_REM_COUNTRY").Value ' $M$199
    End If
    If .Range("DI_TASK2_REM_PCT").Value <> wsSource.Range("DI_AI_REM_PCT").Value Then
        .Range("DI_TASK2_REM_PCT").Value = wsSource.Range("DI_AI_REM_PCT").Value ' $N$199
    End If
    If .Range("DI_TASK3_OVD_JUST").Value <> wsSource.Range("DI_AO_OVD_JUST").Value Then
        .Range("DI_TASK3_OVD_JUST").Value = wsSource.Range("DI_AO_OVD_JUST").Value ' $I$200
    End If
    If .Range("DI_TASK3_OVD_QTY").Value <> wsSource.Range("DI_AO_OVD_QTY").Value Then
        .Range("DI_TASK3_OVD_QTY").Value = wsSource.Range("DI_AO_OVD_QTY").Value ' $H$200
    End If
    If .Range("DI_TASK3_REM_COUNTRY").Value <> wsSource.Range("DI_AO_REM_COUNTRY").Value Then
        .Range("DI_TASK3_REM_COUNTRY").Value = wsSource.Range("DI_AO_REM_COUNTRY").Value ' $M$200
    End If
    If .Range("DI_TASK3_REM_PCT").Value <> wsSource.Range("DI_AO_REM_PCT").Value Then
        .Range("DI_TASK3_REM_PCT").Value = wsSource.Range("DI_AO_REM_PCT").Value ' $N$200
    End If
    If .Range("DI_TASK1_OVD_JUST").Value <> wsSource.Range("DI_DESIGN_OVD_JUST").Value Then
        .Range("DI_TASK1_OVD_JUST").Value = wsSource.Range("DI_DESIGN_OVD_JUST").Value ' $I$198
    End If
    If .Range("DI_TASK1_OVD_QTY").Value <> wsSource.Range("DI_DESIGN_OVD_QTY").Value Then
        .Range("DI_TASK1_OVD_QTY").Value = wsSource.Range("DI_DESIGN_OVD_QTY").Value ' $H$198
    End If
    If .Range("DI_TASK1_REM_COUNTRY").Value <> wsSource.Range("DI_DESIGN_REM_COUNTRY").Value Then
        .Range("DI_TASK1_REM_COUNTRY").Value = wsSource.Range("DI_DESIGN_REM_COUNTRY").Value ' $M$198
    End If
    If .Range("DI_TASK1_REM_PCT").Value <> wsSource.Range("DI_DESIGN_REM_PCT").Value Then
        .Range("DI_TASK1_REM_PCT").Value = wsSource.Range("DI_DESIGN_REM_PCT").Value ' $N$198
    End If
    If .Range("DI_TASK4_OVD_JUST").Value <> wsSource.Range("DI_DI_OVD_JUST").Value Then
        .Range("DI_TASK4_OVD_JUST").Value = wsSource.Range("DI_DI_OVD_JUST").Value ' $I$201
    End If
    If .Range("DI_TASK4_OVD_QTY").Value <> wsSource.Range("DI_DI_OVD_QTY").Value Then
        .Range("DI_TASK4_OVD_QTY").Value = wsSource.Range("DI_DI_OVD_QTY").Value ' $H$201
    End If
    If .Range("DI_TASK4_REM_COUNTRY").Value <> wsSource.Range("DI_DI_REM_COUNTRY").Value Then
        .Range("DI_TASK4_REM_COUNTRY").Value = wsSource.Range("DI_DI_REM_COUNTRY").Value ' $M$201
    End If
    If .Range("DI_TASK4_REM_PCT").Value <> wsSource.Range("DI_DI_REM_PCT").Value Then
        .Range("DI_TASK4_REM_PCT").Value = wsSource.Range("DI_DI_REM_PCT").Value ' $N$201
    End If
    If .Range("DI_TASK5_OVD_JUST").Value <> wsSource.Range("DI_DO_OVD_JUST").Value Then
        .Range("DI_TASK5_OVD_JUST").Value = wsSource.Range("DI_DO_OVD_JUST").Value ' $I$202
    End If
    If .Range("DI_TASK5_OVD_QTY").Value <> wsSource.Range("DI_DO_OVD_QTY").Value Then
        .Range("DI_TASK5_OVD_QTY").Value = wsSource.Range("DI_DO_OVD_QTY").Value ' $H$202
    End If
    If .Range("DI_TASK5_REM_COUNTRY").Value <> wsSource.Range("DI_DO_REM_COUNTRY").Value Then
        .Range("DI_TASK5_REM_COUNTRY").Value = wsSource.Range("DI_DO_REM_COUNTRY").Value ' $M$202
    End If
    If .Range("DI_TASK5_REM_PCT").Value <> wsSource.Range("DI_DO_REM_PCT").Value Then
        .Range("DI_TASK5_REM_PCT").Value = wsSource.Range("DI_DO_REM_PCT").Value ' $N$202
    End If
    If .Range("DI_TASK7_OVD_JUST").Value <> wsSource.Range("DI_GRP_START_OVD_JUST").Value Then
        .Range("DI_TASK7_OVD_JUST").Value = wsSource.Range("DI_GRP_START_OVD_JUST").Value ' $I$204
    End If
    If .Range("DI_TASK7_OVD_QTY").Value <> wsSource.Range("DI_GRP_START_OVD_QTY").Value Then
        .Range("DI_TASK7_OVD_QTY").Value = wsSource.Range("DI_GRP_START_OVD_QTY").Value ' $H$204
    End If
    If .Range("DI_TASK7_REM_COUNTRY").Value <> wsSource.Range("DI_GRP_START_REM_COUNTRY").Value Then
        .Range("DI_TASK7_REM_COUNTRY").Value = wsSource.Range("DI_GRP_START_REM_COUNTRY").Value ' $M$204
    End If
    If .Range("DI_TASK7_REM_PCT").Value <> wsSource.Range("DI_GRP_START_REM_PCT").Value Then
        .Range("DI_TASK7_REM_PCT").Value = wsSource.Range("DI_GRP_START_REM_PCT").Value ' $N$204
    End If
    If .Range("DI_TASK6_OVD_JUST").Value <> wsSource.Range("DI_LOGIC_OVD_JUST").Value Then
        .Range("DI_TASK6_OVD_JUST").Value = wsSource.Range("DI_LOGIC_OVD_JUST").Value ' $I$203
    End If
    If .Range("DI_TASK6_OVD_QTY").Value <> wsSource.Range("DI_LOGIC_OVD_QTY").Value Then
        .Range("DI_TASK6_OVD_QTY").Value = wsSource.Range("DI_LOGIC_OVD_QTY").Value ' $H$203
    End If
    If .Range("DI_TASK6_REM_COUNTRY").Value <> wsSource.Range("DI_LOGIC_REM_COUNTRY").Value Then
        .Range("DI_TASK6_REM_COUNTRY").Value = wsSource.Range("DI_LOGIC_REM_COUNTRY").Value ' $M$203
    End If
    If .Range("DI_TASK6_REM_PCT").Value <> wsSource.Range("DI_LOGIC_REM_PCT").Value Then
        .Range("DI_TASK6_REM_PCT").Value = wsSource.Range("DI_LOGIC_REM_PCT").Value ' $N$203
    End If
    'If .Range("DI_REM_COST").Value <> wsSource.Range("DI_REM_COST").Value Then
    '    .Range("DI_REM_COST").Value = wsSource.Range("DI_LOGIC_REM_PCT").Value ' $N$203
    'End If
    If .Range("DI_TASK8_OVD_JUST").Value <> wsSource.Range("DI_SEQ_OVD_JUST").Value Then
        .Range("DI_TASK8_OVD_JUST").Value = wsSource.Range("DI_SEQ_OVD_JUST").Value ' $I$205
    End If
    If .Range("DI_TASK8_OVD_QTY").Value <> wsSource.Range("DI_SEQ_OVD_QTY").Value Then
        .Range("DI_TASK8_OVD_QTY").Value = wsSource.Range("DI_SEQ_OVD_QTY").Value ' $H$205
    End If
    If .Range("DI_TASK8_REM_COUNTRY").Value <> wsSource.Range("DI_SEQ_REM_COUNTRY").Value Then
        .Range("DI_TASK8_REM_COUNTRY").Value = wsSource.Range("DI_SEQ_REM_COUNTRY").Value ' $M$205
    End If
    If .Range("DI_TASK8_REM_PCT").Value <> wsSource.Range("DI_SEQ_REM_PCT").Value Then
        .Range("DI_TASK8_REM_PCT").Value = wsSource.Range("DI_SEQ_REM_PCT").Value ' $N$205
    End If
    
    'DOC
    If .Range("DOC_TASK1_OVD_JUST").Value <> wsSource.Range("DOC_BOM_OVD_JUST").Value Then
        .Range("DOC_TASK1_OVD_JUST").Value = wsSource.Range("DOC_BOM_OVD_JUST").Value ' $I$288
    End If
    If .Range("DOC_TASK1_OVD_QTY").Value <> wsSource.Range("DOC_BOM_OVD_QTY").Value Then
        .Range("DOC_TASK1_OVD_QTY").Value = wsSource.Range("DOC_BOM_OVD_QTY").Value ' $H$288
    End If
    If .Range("DOC_TASK1_REM_COUNTRY").Value <> wsSource.Range("DOC_BOM_REM_COUNTRY").Value Then
        .Range("DOC_TASK1_REM_COUNTRY").Value = wsSource.Range("DOC_BOM_REM_COUNTRY").Value ' $M$288
    End If
    If .Range("DOC_TASK1_REM_PCT").Value <> wsSource.Range("DOC_BOM_REM_PCT").Value Then
        .Range("DOC_TASK1_REM_PCT").Value = wsSource.Range("DOC_BOM_REM_PCT").Value ' $N$288
    End If
    If .Range("DOC_TASK6_OVD_JUST").Value <> wsSource.Range("DOC_CAB_ELEC_OVD_JUST").Value Then
        .Range("DOC_TASK6_OVD_JUST").Value = wsSource.Range("DOC_CAB_ELEC_OVD_JUST").Value ' $I$293
    End If
    If .Range("DOC_TASK6_OVD_QTY").Value <> wsSource.Range("DOC_CAB_ELEC_OVD_QTY").Value Then
        .Range("DOC_TASK6_OVD_QTY").Value = wsSource.Range("DOC_CAB_ELEC_OVD_QTY").Value ' $H$293
    End If
    If .Range("DOC_TASK6_REM_COUNTRY").Value <> wsSource.Range("DOC_CAB_ELEC_REM_COUNTRY").Value Then
        .Range("DOC_TASK6_REM_COUNTRY").Value = wsSource.Range("DOC_CAB_ELEC_REM_COUNTRY").Value ' $M$293
    End If
    If .Range("DOC_TASK6_REM_PCT").Value <> wsSource.Range("DOC_CAB_ELEC_REM_PCT").Value Then
        .Range("DOC_TASK6_REM_PCT").Value = wsSource.Range("DOC_CAB_ELEC_REM_PCT").Value ' $N$293
    End If
    If .Range("DOC_TASK5_OVD_JUST").Value <> wsSource.Range("DOC_CAB_MECH_OVD_JUST").Value Then
        .Range("DOC_TASK5_OVD_JUST").Value = wsSource.Range("DOC_CAB_MECH_OVD_JUST").Value ' $I$292
    End If
    If .Range("DOC_TASK5_OVD_QTY").Value <> wsSource.Range("DOC_CAB_MECH_OVD_QTY").Value Then
        .Range("DOC_TASK5_OVD_QTY").Value = wsSource.Range("DOC_CAB_MECH_OVD_QTY").Value ' $H$292
    End If
    If .Range("DOC_TASK5_REM_COUNTRY").Value <> wsSource.Range("DOC_CAB_MECH_REM_COUNTRY").Value Then
        .Range("DOC_TASK5_REM_COUNTRY").Value = wsSource.Range("DOC_CAB_MECH_REM_COUNTRY").Value ' $M$292
    End If
    If .Range("DOC_TASK5_REM_PCT").Value <> wsSource.Range("DOC_CAB_MECH_REM_PCT").Value Then
        .Range("DOC_TASK5_REM_PCT").Value = wsSource.Range("DOC_CAB_MECH_REM_PCT").Value ' $N$292
    End If
    If .Range("DOC_TASK9_OVD_JUST").Value <> wsSource.Range("DOC_CUSTOM_OVD_JUST").Value Then
        .Range("DOC_TASK9_OVD_JUST").Value = wsSource.Range("DOC_CUSTOM_OVD_JUST").Value ' $I$296
    End If
    If .Range("DOC_TASK9_OVD_QTY").Value <> wsSource.Range("DOC_CUSTOM_OVD_QTY").Value Then
        .Range("DOC_TASK9_OVD_QTY").Value = wsSource.Range("DOC_CUSTOM_OVD_QTY").Value ' $H$296
    End If
    If .Range("DOC_TASK9_REM_COUNTRY").Value <> wsSource.Range("DOC_CUSTOM_REM_COUNTRY").Value Then
        .Range("DOC_TASK9_REM_COUNTRY").Value = wsSource.Range("DOC_CUSTOM_REM_COUNTRY").Value ' $M$296
    End If
    If .Range("DOC_TASK9_REM_PCT").Value <> wsSource.Range("DOC_CUSTOM_REM_PCT").Value Then
        .Range("DOC_TASK9_REM_PCT").Value = wsSource.Range("DOC_CUSTOM_REM_PCT").Value ' $N$296
    End If
    If .Range("DOC_TASK8_OVD_JUST").Value <> wsSource.Range("DOC_LOOP_OVD_JUST").Value Then
        .Range("DOC_TASK8_OVD_JUST").Value = wsSource.Range("DOC_LOOP_OVD_JUST").Value ' $I$295
    End If
    If .Range("DOC_TASK8_OVD_QTY").Value <> wsSource.Range("DOC_LOOP_OVD_QTY").Value Then
        .Range("DOC_TASK8_OVD_QTY").Value = wsSource.Range("DOC_LOOP_OVD_QTY").Value ' $H$295
    End If
    If .Range("DOC_TASK8_REM_COUNTRY").Value <> wsSource.Range("DOC_LOOP_REM_COUNTRY").Value Then
        .Range("DOC_TASK8_REM_COUNTRY").Value = wsSource.Range("DOC_LOOP_REM_COUNTRY").Value ' $M$295
    End If
    If .Range("DOC_TASK8_REM_PCT").Value <> wsSource.Range("DOC_LOOP_REM_PCT").Value Then
        .Range("DOC_TASK8_REM_PCT").Value = wsSource.Range("DOC_LOOP_REM_PCT").Value ' $N$295
    End If
    If .Range("DOC_TASK4_OVD_JUST").Value <> wsSource.Range("DOC_PWR_GND_OVD_JUST").Value Then
        .Range("DOC_TASK4_OVD_JUST").Value = wsSource.Range("DOC_PWR_GND_OVD_JUST").Value ' $I$291
    End If
    If .Range("DOC_TASK4_OVD_QTY").Value <> wsSource.Range("DOC_PWR_GND_OVD_QTY").Value Then
        .Range("DOC_TASK4_OVD_QTY").Value = wsSource.Range("DOC_PWR_GND_OVD_QTY").Value ' $H$291
    End If
    If .Range("DOC_TASK4_REM_COUNTRY").Value <> wsSource.Range("DOC_PWR_GND_REM_COUNTRY").Value Then
        .Range("DOC_TASK4_REM_COUNTRY").Value = wsSource.Range("DOC_PWR_GND_REM_COUNTRY").Value ' $M$291
    End If
    If .Range("DOC_TASK4_REM_PCT").Value <> wsSource.Range("DOC_PWR_GND_REM_PCT").Value Then
        .Range("DOC_TASK4_REM_PCT").Value = wsSource.Range("DOC_PWR_GND_REM_PCT").Value ' $N$291
    End If
    If .Range("DOC_TASK7_OVD_JUST").Value <> wsSource.Range("DOC_PWR_HEAT_OVD_JUST").Value Then
        .Range("DOC_TASK7_OVD_JUST").Value = wsSource.Range("DOC_PWR_HEAT_OVD_JUST").Value ' $I$294
    End If
    If .Range("DOC_TASK7_OVD_QTY").Value <> wsSource.Range("DOC_PWR_HEAT_OVD_QTY").Value Then
        .Range("DOC_TASK7_OVD_QTY").Value = wsSource.Range("DOC_PWR_HEAT_OVD_QTY").Value ' $H$294
    End If
    If .Range("DOC_TASK7_REM_COUNTRY").Value <> wsSource.Range("DOC_PWR_HEAT_REM_COUNTRY").Value Then
        .Range("DOC_TASK7_REM_COUNTRY").Value = wsSource.Range("DOC_PWR_HEAT_REM_COUNTRY").Value ' $M$294
    End If
    If .Range("DOC_TASK7_REM_PCT").Value <> wsSource.Range("DOC_PWR_HEAT_REM_PCT").Value Then
        .Range("DOC_TASK7_REM_PCT").Value = wsSource.Range("DOC_PWR_HEAT_REM_PCT").Value ' $N$294
    End If
    'If .Range("DOC_REM_COST").Value <> wsSource.Range("DOC_REM_COST").Value Then
    '    .Range("DOC_REM_COST").Value = wsSource.Range("DOC_PWR_HEAT_REM_PCT").Value ' $N$294
    'End If
    If .Range("DOC_TASK2_OVD_JUST").Value <> wsSource.Range("DOC_SYS_ARCH_OVD_JUST").Value Then
        .Range("DOC_TASK2_OVD_JUST").Value = wsSource.Range("DOC_SYS_ARCH_OVD_JUST").Value ' $I$289
    End If
    If .Range("DOC_TASK2_OVD_QTY").Value <> wsSource.Range("DOC_SYS_ARCH_OVD_QTY").Value Then
        .Range("DOC_TASK2_OVD_QTY").Value = wsSource.Range("DOC_SYS_ARCH_OVD_QTY").Value ' $H$289
    End If
    If .Range("DOC_TASK2_REM_COUNTRY").Value <> wsSource.Range("DOC_SYS_ARCH_REM_COUNTRY").Value Then
        .Range("DOC_TASK2_REM_COUNTRY").Value = wsSource.Range("DOC_SYS_ARCH_REM_COUNTRY").Value ' $M$289
    End If
    If .Range("DOC_TASK2_REM_PCT").Value <> wsSource.Range("DOC_SYS_ARCH_REM_PCT").Value Then
        .Range("DOC_TASK2_REM_PCT").Value = wsSource.Range("DOC_SYS_ARCH_REM_PCT").Value ' $N$289
    End If
    If .Range("DOC_TASK3_OVD_JUST").Value <> wsSource.Range("DOC_SYS_INT_OVD_JUST").Value Then
        .Range("DOC_TASK3_OVD_JUST").Value = wsSource.Range("DOC_SYS_INT_OVD_JUST").Value ' $I$290
    End If
    If .Range("DOC_TASK3_OVD_QTY").Value <> wsSource.Range("DOC_SYS_INT_OVD_QTY").Value Then
        .Range("DOC_TASK3_OVD_QTY").Value = wsSource.Range("DOC_SYS_INT_OVD_QTY").Value ' $H$290
    End If
    If .Range("DOC_TASK3_REM_COUNTRY").Value <> wsSource.Range("DOC_SYS_INT_REM_COUNTRY").Value Then
        .Range("DOC_TASK3_REM_COUNTRY").Value = wsSource.Range("DOC_SYS_INT_REM_COUNTRY").Value ' $M$290
    End If
    If .Range("DOC_TASK3_REM_PCT").Value <> wsSource.Range("DOC_SYS_INT_REM_PCT").Value Then
        .Range("DOC_TASK3_REM_PCT").Value = wsSource.Range("DOC_SYS_INT_REM_PCT").Value ' $N$290
    End If
    If .Range("DOC_TASK10_OVD_JUST").Value <> wsSource.Range("DOC_TAGLIST_OVD_JUST").Value Then
        .Range("DOC_TASK10_OVD_JUST").Value = wsSource.Range("DOC_TAGLIST_OVD_JUST").Value ' $I$297
    End If
    If .Range("DOC_TASK10_OVD_QTY").Value <> wsSource.Range("DOC_TAGLIST_OVD_QTY").Value Then
        .Range("DOC_TASK10_OVD_QTY").Value = wsSource.Range("DOC_TAGLIST_OVD_QTY").Value ' $H$297
    End If
    If .Range("DOC_TASK10_REM_COUNTRY").Value <> wsSource.Range("DOC_TAGLIST_REM_COUNTRY").Value Then
        .Range("DOC_TASK10_REM_COUNTRY").Value = wsSource.Range("DOC_TAGLIST_REM_COUNTRY").Value ' $M$297
    End If
    If .Range("DOC_TASK10_REM_PCT").Value <> wsSource.Range("DOC_TAGLIST_REM_PCT").Value Then
        .Range("DOC_TASK10_REM_PCT").Value = wsSource.Range("DOC_TAGLIST_REM_PCT").Value ' $N$297
    End If
 'Formula: =G242
    If .Range("DUP_FACT_APP_BASED_APP").Value <> wsSource.Range("DUP_FACT_APP_BASED_APP").Value Then
        .Range("DUP_FACT_APP_BASED_APP").Value = wsSource.Range("DUP_FACT_APP_BASED_APP").Value ' $F$242
    End If
 'Formula: =G300
    If .Range("DUP_FACT_APP_BASED_COURSE").Value <> wsSource.Range("DUP_FACT_APP_BASED_COURSE").Value Then
        .Range("DUP_FACT_APP_BASED_COURSE").Value = wsSource.Range("DUP_FACT_APP_BASED_COURSE").Value ' $F$300
    End If
 'Formula: =G169
    If .Range("DUP_FACT_APP_BASED_CP").Value <> wsSource.Range("DUP_FACT_APP_BASED_CP").Value Then
        .Range("DUP_FACT_APP_BASED_CP").Value = wsSource.Range("DUP_FACT_APP_BASED_CP").Value ' $F$169
    End If
 'Formula: =G189
    If .Range("DUP_FACT_APP_BASED_DI").Value <> wsSource.Range("DUP_FACT_APP_BASED_DI").Value Then
        .Range("DUP_FACT_APP_BASED_DI").Value = wsSource.Range("DUP_FACT_APP_BASED_DI").Value ' $F$189
    End If
 'Formula: =G279
    If .Range("DUP_FACT_APP_BASED_DOC").Value <> wsSource.Range("DUP_FACT_APP_BASED_DOC").Value Then
        .Range("DUP_FACT_APP_BASED_DOC").Value = wsSource.Range("DUP_FACT_APP_BASED_DOC").Value ' $F$279
    End If
 'Formula: =G209
    If .Range("DUP_FACT_APP_BASED_ESD").Value <> wsSource.Range("DUP_FACT_APP_BASED_ESD").Value Then
        .Range("DUP_FACT_APP_BASED_ESD").Value = wsSource.Range("DUP_FACT_APP_BASED_ESD").Value ' $F$209
    End If
 'Formula: =G147
    If .Range("DUP_FACT_APP_BASED_HMI").Value <> wsSource.Range("DUP_FACT_APP_BASED_HMI").Value Then
        .Range("DUP_FACT_APP_BASED_HMI").Value = wsSource.Range("DUP_FACT_APP_BASED_HMI").Value ' $F$147
    End If
 'Formula: =G330
    If .Range("DUP_FACT_APP_BASED_MEETING").Value <> wsSource.Range("DUP_FACT_APP_BASED_MEETING").Value Then
        .Range("DUP_FACT_APP_BASED_MEETING").Value = wsSource.Range("DUP_FACT_APP_BASED_MEETING").Value ' $F$330
    End If
 'Formula: =G316
    If .Range("DUP_FACT_APP_BASED_PM").Value <> wsSource.Range("DUP_FACT_APP_BASED_PM").Value Then
        .Range("DUP_FACT_APP_BASED_PM").Value = wsSource.Range("DUP_FACT_APP_BASED_PM").Value ' $F$316
    End If
 'Formula: =G228
    If .Range("DUP_FACT_APP_BASED_REP").Value <> wsSource.Range("DUP_FACT_APP_BASED_REP").Value Then
        .Range("DUP_FACT_APP_BASED_REP").Value = wsSource.Range("DUP_FACT_APP_BASED_REP").Value ' $F$228
    End If
 'Formula: =G346
    If .Range("DUP_FACT_APP_BASED_SITE").Value <> wsSource.Range("DUP_FACT_APP_BASED_SITE").Value Then
        .Range("DUP_FACT_APP_BASED_SITE").Value = wsSource.Range("DUP_FACT_APP_BASED_SITE").Value ' $F$346
    End If
 'Formula: =G119
    If .Range("DUP_FACT_APP_BASED_SPEC").Value <> wsSource.Range("DUP_FACT_APP_BASED_SPEC").Value Then
        .Range("DUP_FACT_APP_BASED_SPEC").Value = wsSource.Range("DUP_FACT_APP_BASED_SPEC").Value ' $F$119
    End If
 'Formula: =G134
    If .Range("DUP_FACT_APP_BASED_SYS_ENG").Value <> wsSource.Range("DUP_FACT_APP_BASED_SYS_ENG").Value Then
        .Range("DUP_FACT_APP_BASED_SYS_ENG").Value = wsSource.Range("DUP_FACT_APP_BASED_SYS_ENG").Value ' $F$134
    End If
 'Formula: =G134
    If .Range("DUP_FACT_APP_BASED_SYSENG").Value <> wsSource.Range("DUP_FACT_APP_BASED_SYSENG").Value Then
        .Range("DUP_FACT_APP_BASED_SYSENG").Value = wsSource.Range("DUP_FACT_APP_BASED_SYSENG").Value ' $F$134
    End If
 'Formula: =G262
    If .Range("DUP_FACT_APP_BASED_TEST").Value <> wsSource.Range("DUP_FACT_APP_BASED_TEST").Value Then
        .Range("DUP_FACT_APP_BASED_TEST").Value = wsSource.Range("DUP_FACT_APP_BASED_TEST").Value ' $F$262
    End If
    
    'ESD
    If .Range("ESD_TASK2_OVD_JUST").Value <> wsSource.Range("ESD_AI_OVD_JUST").Value Then
        .Range("ESD_TASK2_OVD_JUST").Value = wsSource.Range("ESD_AI_OVD_JUST").Value ' $I$219
    End If
    If .Range("ESD_TASK2_OVD_QTY").Value <> wsSource.Range("ESD_AI_OVD_QTY").Value Then
        .Range("ESD_TASK2_OVD_QTY").Value = wsSource.Range("ESD_AI_OVD_QTY").Value ' $H$219
    End If
    If .Range("ESD_TASK2_REM_COUNTRY").Value <> wsSource.Range("ESD_AI_REM_COUNTRY").Value Then
        .Range("ESD_TASK2_REM_COUNTRY").Value = wsSource.Range("ESD_AI_REM_COUNTRY").Value ' $M$219
    End If
    If .Range("ESD_TASK2_REM_PCT").Value <> wsSource.Range("ESD_AI_REM_PCT").Value Then
        .Range("ESD_TASK2_REM_PCT").Value = wsSource.Range("ESD_AI_REM_PCT").Value ' $N$219
    End If
    If .Range("ESD_TASK3_OVD_JUST").Value <> wsSource.Range("ESD_AO_OVD_JUST").Value Then
        .Range("ESD_TASK3_OVD_JUST").Value = wsSource.Range("ESD_AO_OVD_JUST").Value ' $I$220
    End If
    If .Range("ESD_TASK3_OVD_QTY").Value <> wsSource.Range("ESD_AO_OVD_QTY").Value Then
        .Range("ESD_TASK3_OVD_QTY").Value = wsSource.Range("ESD_AO_OVD_QTY").Value ' $H$220
    End If
    If .Range("ESD_TASK3_REM_COUNTRY").Value <> wsSource.Range("ESD_AO_REM_COUNTRY").Value Then
        .Range("ESD_TASK3_REM_COUNTRY").Value = wsSource.Range("ESD_AO_REM_COUNTRY").Value ' $M$220
    End If
    If .Range("ESD_TASK3_REM_PCT").Value <> wsSource.Range("ESD_AO_REM_PCT").Value Then
        .Range("ESD_TASK3_REM_PCT").Value = wsSource.Range("ESD_AO_REM_PCT").Value ' $N$220
    End If
    If .Range("ESD_TASK4_OVD_JUST").Value <> wsSource.Range("ESD_DI_OVD_JUST").Value Then
        .Range("ESD_TASK4_OVD_JUST").Value = wsSource.Range("ESD_DI_OVD_JUST").Value ' $I$221
    End If
    If .Range("ESD_TASK4_OVD_QTY").Value <> wsSource.Range("ESD_DI_OVD_QTY").Value Then
        .Range("ESD_TASK4_OVD_QTY").Value = wsSource.Range("ESD_DI_OVD_QTY").Value ' $H$221
    End If
    If .Range("ESD_TASK4_REM_COUNTRY").Value <> wsSource.Range("ESD_DI_REM_COUNTRY").Value Then
        .Range("ESD_TASK4_REM_COUNTRY").Value = wsSource.Range("ESD_DI_REM_COUNTRY").Value ' $M$221
    End If
    If .Range("ESD_TASK4_REM_PCT").Value <> wsSource.Range("ESD_DI_REM_PCT").Value Then
        .Range("ESD_TASK4_REM_PCT").Value = wsSource.Range("ESD_DI_REM_PCT").Value ' $N$221
    End If
    If .Range("ESD_TASK5_OVD_JUST").Value <> wsSource.Range("ESD_DO_OVD_JUST").Value Then
        .Range("ESD_TASK5_OVD_JUST").Value = wsSource.Range("ESD_DO_OVD_JUST").Value ' $I$222
    End If
    If .Range("ESD_TASK5_OVD_QTY").Value <> wsSource.Range("ESD_DO_OVD_QTY").Value Then
        .Range("ESD_TASK5_OVD_QTY").Value = wsSource.Range("ESD_DO_OVD_QTY").Value ' $H$222
    End If
    If .Range("ESD_TASK5_REM_COUNTRY").Value <> wsSource.Range("ESD_DO_REM_COUNTRY").Value Then
        .Range("ESD_TASK5_REM_COUNTRY").Value = wsSource.Range("ESD_DO_REM_COUNTRY").Value ' $M$222
    End If
    If .Range("ESD_TASK5_REM_PCT").Value <> wsSource.Range("ESD_DO_REM_PCT").Value Then
        .Range("ESD_TASK5_REM_PCT").Value = wsSource.Range("ESD_DO_REM_PCT").Value ' $N$222
    End If
    If .Range("ESD_TASK7_OVD_JUST").Value <> wsSource.Range("ESD_GRP_START_OVD_JUST").Value Then
        .Range("ESD_TASK7_OVD_JUST").Value = wsSource.Range("ESD_GRP_START_OVD_JUST").Value ' $I$224
    End If
    If .Range("ESD_TASK7_OVD_QTY").Value <> wsSource.Range("ESD_GRP_START_OVD_QTY").Value Then
        .Range("ESD_TASK7_OVD_QTY").Value = wsSource.Range("ESD_GRP_START_OVD_QTY").Value ' $H$224
    End If
    If .Range("ESD_TASK7_REM_COUNTRY").Value <> wsSource.Range("ESD_GRP_START_REM_COUNTRY").Value Then
        .Range("ESD_TASK7_REM_COUNTRY").Value = wsSource.Range("ESD_GRP_START_REM_COUNTRY").Value ' $M$224
    End If
    If .Range("ESD_TASK7_REM_PCT").Value <> wsSource.Range("ESD_GRP_START_REM_PCT").Value Then
        .Range("ESD_TASK7_REM_PCT").Value = wsSource.Range("ESD_GRP_START_REM_PCT").Value ' $N$224
    End If
    'If .Range("ESD_IA_DESIGN_OVD_JUST").Value <> wsSource.Range("ESD_IA_DESIGN_OVD_JUST").Value Then
    '    .Range("ESD_IA_DESIGN_OVD_JUST").Value = wsSource.Range("ESD_GRP_START_REM_PCT").Value ' $N$224
    'End If
    'If .Range("ESD_IA_DESIGN_OVD_QTY").Value <> wsSource.Range("ESD_IA_DESIGN_OVD_QTY").Value Then
    '    .Range("ESD_IA_DESIGN_OVD_QTY").Value = wsSource.Range("ESD_GRP_START_REM_PCT").Value ' $N$224
    'End If
    'If .Range("ESD_IA_DESIGN_REM_COUNTRY").Value <> wsSource.Range("ESD_IA_DESIGN_REM_COUNTRY").Value Then
    '    .Range("ESD_IA_DESIGN_REM_COUNTRY").Value = wsSource.Range("ESD_GRP_START_REM_PCT").Value ' $N$224
    'End If
    'If .Range("ESD_IA_DESIGN_REM_CURRENCY").Value <> wsSource.Range("ESD_IA_DESIGN_REM_CURRENCY").Value Then
    '    .Range("ESD_IA_DESIGN_REM_CURRENCY").Value = wsSource.Range("ESD_GRP_START_REM_PCT").Value ' $N$224
    'End If
    'If .Range("ESD_IA_DESIGN_REM_PCT").Value <> wsSource.Range("ESD_IA_DESIGN_REM_PCT").Value Then
    '    .Range("ESD_IA_DESIGN_REM_PCT").Value = wsSource.Range("ESD_GRP_START_REM_PCT").Value ' $N$224
    'End If
    'If .Range("ESD_IA_IMPL_OVD_JUST").Value <> wsSource.Range("ESD_IA_IMPL_OVD_JUST").Value Then
    '    .Range("ESD_IA_IMPL_OVD_JUST").Value = wsSource.Range("ESD_GRP_START_REM_PCT").Value ' $N$224
    'End If
    'If .Range("ESD_IA_IMPL_OVD_QTY").Value <> wsSource.Range("ESD_IA_IMPL_OVD_QTY").Value Then
    '    .Range("ESD_IA_IMPL_OVD_QTY").Value = wsSource.Range("ESD_GRP_START_REM_PCT").Value ' $N$224
    'End If
    'If .Range("ESD_IA_IMPL_REM_COUNTRY").Value <> wsSource.Range("ESD_IA_IMPL_REM_COUNTRY").Value Then
    '    .Range("ESD_IA_IMPL_REM_COUNTRY").Value = wsSource.Range("ESD_GRP_START_REM_PCT").Value ' $N$224
    'End If
    'If .Range("ESD_IA_IMPL_REM_CURRENCY").Value <> wsSource.Range("ESD_IA_IMPL_REM_CURRENCY").Value Then
    '    .Range("ESD_IA_IMPL_REM_CURRENCY").Value = wsSource.Range("ESD_GRP_START_REM_PCT").Value ' $N$224
    'End If
    'If .Range("ESD_IA_IMPL_REM_PCT").Value <> wsSource.Range("ESD_IA_IMPL_REM_PCT").Value Then
    '    .Range("ESD_IA_IMPL_REM_PCT").Value = wsSource.Range("ESD_GRP_START_REM_PCT").Value ' $N$224
    'End If
    If .Range("ESD_TASK1_OVD_JUST").Value <> wsSource.Range("ESD_TRICON_DESIGN_OVD_JUST").Value Then
        .Range("ESD_TASK1_OVD_JUST").Value = wsSource.Range("ESD_TRICON_DESIGN_OVD_JUST").Value ' $I$218
    End If
    If .Range("ESD_TASK1_OVD_QTY").Value <> wsSource.Range("ESD_TRICON_DESIGN_OVD_QTY").Value Then
        .Range("ESD_TASK1_OVD_QTY").Value = wsSource.Range("ESD_TRICON_DESIGN_OVD_QTY").Value ' $H$218
    End If
    If .Range("ESD_TASK1_REM_COUNTRY").Value <> wsSource.Range("ESD_TRICON_DESIGN_REM_COUNTRY").Value Then
        .Range("ESD_TASK1_REM_COUNTRY").Value = wsSource.Range("ESD_TRICON_DESIGN_REM_COUNTRY").Value ' $M$218
    End If
    If .Range("ESD_TASK1_REM_PCT").Value <> wsSource.Range("ESD_TRICON_DESIGN_REM_PCT").Value Then
        .Range("ESD_TASK1_REM_PCT").Value = wsSource.Range("ESD_TRICON_DESIGN_REM_PCT").Value ' $N$218
    End If
    If .Range("ESD_TASK6_OVD_JUST").Value <> wsSource.Range("ESD_TRICON_LOGIC_OVD_JUST").Value Then
        .Range("ESD_TASK6_OVD_JUST").Value = wsSource.Range("ESD_TRICON_LOGIC_OVD_JUST").Value ' $I$223
    End If
    If .Range("ESD_TASK6_OVD_QTY").Value <> wsSource.Range("ESD_TRICON_LOGIC_OVD_QTY").Value Then
        .Range("ESD_TASK6_OVD_QTY").Value = wsSource.Range("ESD_TRICON_LOGIC_OVD_QTY").Value ' $H$223
    End If
    If .Range("ESD_TASK6_REM_COUNTRY").Value <> wsSource.Range("ESD_TRICON_LOGIC_REM_COUNTRY").Value Then
        .Range("ESD_TASK6_REM_COUNTRY").Value = wsSource.Range("ESD_TRICON_LOGIC_REM_COUNTRY").Value ' $M$223
    End If
    If .Range("ESD_TASK6_REM_PCT").Value <> wsSource.Range("ESD_TRICON_LOGIC_REM_PCT").Value Then
        .Range("ESD_TASK6_REM_PCT").Value = wsSource.Range("ESD_TRICON_LOGIC_REM_PCT").Value ' $N$223
    End If
    
    'HMI
    If .Range("HMI_TASK10_OVD_JUST").Value <> wsSource.Range("HMI_ALARM_OVD_JUST").Value Then
        .Range("HMI_TASK10_OVD_JUST").Value = wsSource.Range("HMI_ALARM_OVD_JUST").Value ' $I$165
    End If
    If .Range("HMI_TASK10_OVD_QTY").Value <> wsSource.Range("HMI_ALARM_OVD_QTY").Value Then
        .Range("HMI_TASK10_OVD_QTY").Value = wsSource.Range("HMI_ALARM_OVD_QTY").Value ' $H$165
    End If
    If .Range("HMI_TASK10_REM_COUNTRY").Value <> wsSource.Range("HMI_ALARM_REM_COUNTRY").Value Then
        .Range("HMI_TASK10_REM_COUNTRY").Value = wsSource.Range("HMI_ALARM_REM_COUNTRY").Value ' $M$165
    End If
    If .Range("HMI_TASK10_REM_PCT").Value <> wsSource.Range("HMI_ALARM_REM_PCT").Value Then
        .Range("HMI_TASK10_REM_PCT").Value = wsSource.Range("HMI_ALARM_REM_PCT").Value ' $N$165
    End If
    If .Range("HMI_TASK1_OVD_JUST").Value <> wsSource.Range("HMI_ELEMENTS_OVD_JUST").Value Then
        .Range("HMI_TASK1_OVD_JUST").Value = wsSource.Range("HMI_ELEMENTS_OVD_JUST").Value ' $I$156
    End If
    If .Range("HMI_TASK1_OVD_QTY").Value <> wsSource.Range("HMI_ELEMENTS_OVD_QTY").Value Then
        .Range("HMI_TASK1_OVD_QTY").Value = wsSource.Range("HMI_ELEMENTS_OVD_QTY").Value ' $H$156
    End If
    If .Range("HMI_TASK1_REM_COUNTRY").Value <> wsSource.Range("HMI_ELEMENTS_REM_COUNTRY").Value Then
        .Range("HMI_TASK1_REM_COUNTRY").Value = wsSource.Range("HMI_ELEMENTS_REM_COUNTRY").Value ' $M$156
    End If
    If .Range("HMI_TASK1_REM_PCT").Value <> wsSource.Range("HMI_ELEMENTS_REM_PCT").Value Then
        .Range("HMI_TASK1_REM_PCT").Value = wsSource.Range("HMI_ELEMENTS_REM_PCT").Value ' $N$156
    End If
    If .Range("HMI_TASK9_OVD_JUST").Value <> wsSource.Range("HMI_ENV_OVD_JUST").Value Then
        .Range("HMI_TASK9_OVD_JUST").Value = wsSource.Range("HMI_ENV_OVD_JUST").Value ' $I$164
    End If
    If .Range("HMI_TASK9_OVD_QTY").Value <> wsSource.Range("HMI_ENV_OVD_QTY").Value Then
        .Range("HMI_TASK9_OVD_QTY").Value = wsSource.Range("HMI_ENV_OVD_QTY").Value ' $H$164
    End If
    If .Range("HMI_TASK9_REM_COUNTRY").Value <> wsSource.Range("HMI_ENV_REM_COUNTRY").Value Then
        .Range("HMI_TASK9_REM_COUNTRY").Value = wsSource.Range("HMI_ENV_REM_COUNTRY").Value ' $M$164
    End If
    If .Range("HMI_TASK9_REM_PCT").Value <> wsSource.Range("HMI_ENV_REM_PCT").Value Then
        .Range("HMI_TASK9_REM_PCT").Value = wsSource.Range("HMI_ENV_REM_PCT").Value ' $N$164
    End If
    If .Range("HMI_TASK5_OVD_JUST").Value <> wsSource.Range("HMI_ESD_OVD_JUST").Value Then
        .Range("HMI_TASK5_OVD_JUST").Value = wsSource.Range("HMI_ESD_OVD_JUST").Value ' $I$160
    End If
    If .Range("HMI_TASK5_OVD_QTY").Value <> wsSource.Range("HMI_ESD_OVD_QTY").Value Then
        .Range("HMI_TASK5_OVD_QTY").Value = wsSource.Range("HMI_ESD_OVD_QTY").Value ' $H$160
    End If
    If .Range("HMI_TASK5_REM_COUNTRY").Value <> wsSource.Range("HMI_ESD_REM_COUNTRY").Value Then
        .Range("HMI_TASK5_REM_COUNTRY").Value = wsSource.Range("HMI_ESD_REM_COUNTRY").Value ' $M$160
    End If
    If .Range("HMI_TASK5_REM_PCT").Value <> wsSource.Range("HMI_ESD_REM_PCT").Value Then
        .Range("HMI_TASK5_REM_PCT").Value = wsSource.Range("HMI_ESD_REM_PCT").Value ' $N$160
    End If
    If .Range("HMI_TASK7_OVD_JUST").Value <> wsSource.Range("HMI_GRP_OVD_JUST").Value Then
        .Range("HMI_TASK7_OVD_JUST").Value = wsSource.Range("HMI_GRP_OVD_JUST").Value ' $I$162
    End If
    If .Range("HMI_TASK7_OVD_QTY").Value <> wsSource.Range("HMI_GRP_OVD_QTY").Value Then
        .Range("HMI_TASK7_OVD_QTY").Value = wsSource.Range("HMI_GRP_OVD_QTY").Value ' $H$162
    End If
    If .Range("HMI_TASK7_REM_COUNTRY").Value <> wsSource.Range("HMI_GRP_REM_COUNTRY").Value Then
        .Range("HMI_TASK7_REM_COUNTRY").Value = wsSource.Range("HMI_GRP_REM_COUNTRY").Value ' $M$162
    End If
    If .Range("HMI_TASK7_REM_PCT").Value <> wsSource.Range("HMI_GRP_REM_PCT").Value Then
        .Range("HMI_TASK7_REM_PCT").Value = wsSource.Range("HMI_GRP_REM_PCT").Value ' $N$162
    End If
    If .Range("HMI_TASK8_OVD_JUST").Value <> wsSource.Range("HMI_OVL_OVD_JUST").Value Then
        .Range("HMI_TASK8_OVD_JUST").Value = wsSource.Range("HMI_OVL_OVD_JUST").Value ' $I$163
    End If
    If .Range("HMI_TASK8_OVD_QTY").Value <> wsSource.Range("HMI_OVL_OVD_QTY").Value Then
        .Range("HMI_TASK8_OVD_QTY").Value = wsSource.Range("HMI_OVL_OVD_QTY").Value ' $H$163
    End If
    If .Range("HMI_TASK8_REM_COUNTRY").Value <> wsSource.Range("HMI_OVL_REM_COUNTRY").Value Then
        .Range("HMI_TASK8_REM_COUNTRY").Value = wsSource.Range("HMI_OVL_REM_COUNTRY").Value ' $M$163
    End If
    If .Range("HMI_TASK8_REM_PCT").Value <> wsSource.Range("HMI_OVL_REM_PCT").Value Then
        .Range("HMI_TASK8_REM_PCT").Value = wsSource.Range("HMI_OVL_REM_PCT").Value ' $N$163
    End If
    If .Range("HMI_TASK2_OVD_JUST").Value <> wsSource.Range("HMI_OVR_OVD_JUST").Value Then
        .Range("HMI_TASK2_OVD_JUST").Value = wsSource.Range("HMI_OVR_OVD_JUST").Value ' $I$157
    End If
    If .Range("HMI_TASK2_OVD_QTY").Value <> wsSource.Range("HMI_OVR_OVD_QTY").Value Then
        .Range("HMI_TASK2_OVD_QTY").Value = wsSource.Range("HMI_OVR_OVD_QTY").Value ' $H$157
    End If
    If .Range("HMI_TASK2_REM_COUNTRY").Value <> wsSource.Range("HMI_OVR_REM_COUNTRY").Value Then
        .Range("HMI_TASK2_REM_COUNTRY").Value = wsSource.Range("HMI_OVR_REM_COUNTRY").Value ' $M$157
    End If
    If .Range("HMI_TASK2_REM_PCT").Value <> wsSource.Range("HMI_OVR_REM_PCT").Value Then
        .Range("HMI_TASK2_REM_PCT").Value = wsSource.Range("HMI_OVR_REM_PCT").Value ' $N$157
    End If
    If .Range("HMI_TASK3_OVD_JUST").Value <> wsSource.Range("HMI_PROC_OVD_JUST").Value Then
        .Range("HMI_TASK3_OVD_JUST").Value = wsSource.Range("HMI_PROC_OVD_JUST").Value ' $I$158
    End If
    If .Range("HMI_TASK3_OVD_QTY").Value <> wsSource.Range("HMI_PROC_OVD_QTY").Value Then
        .Range("HMI_TASK3_OVD_QTY").Value = wsSource.Range("HMI_PROC_OVD_QTY").Value ' $H$158
    End If
    If .Range("HMI_TASK3_REM_COUNTRY").Value <> wsSource.Range("HMI_PROC_REM_COUNTRY").Value Then
        .Range("HMI_TASK3_REM_COUNTRY").Value = wsSource.Range("HMI_PROC_REM_COUNTRY").Value ' $M$158
    End If
    If .Range("HMI_TASK3_REM_PCT").Value <> wsSource.Range("HMI_PROC_REM_PCT").Value Then
        .Range("HMI_TASK3_REM_PCT").Value = wsSource.Range("HMI_PROC_REM_PCT").Value ' $N$158
    End If
    'If .Range("HMI_REM_COST").Value <> wsSource.Range("HMI_REM_COST").Value Then
    '    .Range("HMI_REM_COST").Value = wsSource.Range("HMI_PROC_REM_PCT").Value ' $N$158
    'End If
    If .Range("HMI_TASK4_OVD_JUST").Value <> wsSource.Range("HMI_RPT_OVD_JUST").Value Then
        .Range("HMI_TASK4_OVD_JUST").Value = wsSource.Range("HMI_RPT_OVD_JUST").Value ' $I$159
    End If
    If .Range("HMI_TASK4_OVD_QTY").Value <> wsSource.Range("HMI_RPT_OVD_QTY").Value Then
        .Range("HMI_TASK4_OVD_QTY").Value = wsSource.Range("HMI_RPT_OVD_QTY").Value ' $H$159
    End If
 'Formula: =HMI_RPT_PCT_QTY_TUNED
    If .Range("HMI_RPT_PCT_QTY").Value <> wsSource.Range("HMI_RPT_PCT_QTY").Value Then
        .Range("HMI_RPT_PCT_QTY").Value = wsSource.Range("HMI_RPT_PCT_QTY").Value ' $E$106
    End If
    If .Range("HMI_TASK4_REM_COUNTRY").Value <> wsSource.Range("HMI_RPT_REM_COUNTRY").Value Then
        .Range("HMI_TASK4_REM_COUNTRY").Value = wsSource.Range("HMI_RPT_REM_COUNTRY").Value ' $M$159
    End If
    If .Range("HMI_TASK4_REM_PCT").Value <> wsSource.Range("HMI_RPT_REM_PCT").Value Then
        .Range("HMI_TASK4_REM_PCT").Value = wsSource.Range("HMI_RPT_REM_PCT").Value ' $N$159
    End If
    If .Range("HMI_TASK6_OVD_JUST").Value <> wsSource.Range("HMI_TRD_OVD_JUST").Value Then
        .Range("HMI_TASK6_OVD_JUST").Value = wsSource.Range("HMI_TRD_OVD_JUST").Value ' $I$161
    End If
    If .Range("HMI_TASK6_OVD_QTY").Value <> wsSource.Range("HMI_TRD_OVD_QTY").Value Then
        .Range("HMI_TASK6_OVD_QTY").Value = wsSource.Range("HMI_TRD_OVD_QTY").Value ' $H$161
    End If
    If .Range("HMI_TASK6_REM_COUNTRY").Value <> wsSource.Range("HMI_TRD_REM_COUNTRY").Value Then
        .Range("HMI_TASK6_REM_COUNTRY").Value = wsSource.Range("HMI_TRD_REM_COUNTRY").Value ' $M$161
    End If
    If .Range("HMI_TASK6_REM_PCT").Value <> wsSource.Range("HMI_TRD_REM_PCT").Value Then
        .Range("HMI_TASK6_REM_PCT").Value = wsSource.Range("HMI_TRD_REM_PCT").Value ' $N$161
    End If
    If .Range("IO_OPTIMISATION_EXP_RATE").Value <> wsSource.Range("IO_OPTIMISATION_EXP_RATE").Value Then
        .Range("IO_OPTIMISATION_EXP_RATE").Value = wsSource.Range("IO_OPTIMISATION_EXP_RATE").Value ' $E$108
    End If
    If .Range("IO_OPTIMISATION_MAX_REDUCTION").Value <> wsSource.Range("IO_OPTIMISATION_MAX_REDUCTION").Value Then
        .Range("IO_OPTIMISATION_MAX_REDUCTION").Value = wsSource.Range("IO_OPTIMISATION_MAX_REDUCTION").Value ' $E$109
    End If
    If .Range("MEETING_TASK5_OVD_JUST").Value <> wsSource.Range("MEETING_CLOSE_OVD_JUST").Value Then
        .Range("MEETING_TASK5_OVD_JUST").Value = wsSource.Range("MEETING_CLOSE_OVD_JUST").Value ' $I$343
    End If
    If .Range("MEETING_TASK5_OVD_QTY").Value <> wsSource.Range("MEETING_CLOSE_OVD_QTY").Value Then
        .Range("MEETING_TASK5_OVD_QTY").Value = wsSource.Range("MEETING_CLOSE_OVD_QTY").Value ' $H$343
    End If
    If .Range("MEETING_TASK5_REM_COUNTRY").Value <> wsSource.Range("MEETING_CLOSE_REM_COUNTRY").Value Then
        .Range("MEETING_TASK5_REM_COUNTRY").Value = wsSource.Range("MEETING_CLOSE_REM_COUNTRY").Value ' $M$343
    End If
    If .Range("MEETING_TASK5_REM_PCT").Value <> wsSource.Range("MEETING_CLOSE_REM_PCT").Value Then
        .Range("MEETING_TASK5_REM_PCT").Value = wsSource.Range("MEETING_CLOSE_REM_PCT").Value ' $N$343
    End If
    If .Range("MEETING_TASK2_OVD_JUST").Value <> wsSource.Range("MEETING_DESIGN_OVD_JUST").Value Then
        .Range("MEETING_TASK2_OVD_JUST").Value = wsSource.Range("MEETING_DESIGN_OVD_JUST").Value ' $I$340
    End If
    If .Range("MEETING_TASK2_OVD_QTY").Value <> wsSource.Range("MEETING_DESIGN_OVD_QTY").Value Then
        .Range("MEETING_TASK2_OVD_QTY").Value = wsSource.Range("MEETING_DESIGN_OVD_QTY").Value ' $H$340
    End If
    If .Range("MEETING_TASK2_REM_COUNTRY").Value <> wsSource.Range("MEETING_DESIGN_REM_COUNTRY").Value Then
        .Range("MEETING_TASK2_REM_COUNTRY").Value = wsSource.Range("MEETING_DESIGN_REM_COUNTRY").Value ' $M$340
    End If
    If .Range("MEETING_TASK2_REM_PCT").Value <> wsSource.Range("MEETING_DESIGN_REM_PCT").Value Then
        .Range("MEETING_TASK2_REM_PCT").Value = wsSource.Range("MEETING_DESIGN_REM_PCT").Value ' $N$340
    End If
    'If .Range("MEETING_DESIGNREM_PCT").Value <> wsSource.Range("MEETING_DESIGNREM_PCT").Value Then
    '    .Range("MEETING_DESIGNREM_PCT").Value = wsSource.Range("MEETING_DESIGNREM_PCT").Value ' $N$340
    'End If
    If .Range("MEETING_TASK1_OVD_JUST").Value <> wsSource.Range("MEETING_KICKOFF_OVD_JUST").Value Then
        .Range("MEETING_TASK1_OVD_JUST").Value = wsSource.Range("MEETING_KICKOFF_OVD_JUST").Value ' $I$339
    End If
    If .Range("MEETING_TASK1_OVD_QTY").Value <> wsSource.Range("MEETING_KICKOFF_OVD_QTY").Value Then
        .Range("MEETING_TASK1_OVD_QTY").Value = wsSource.Range("MEETING_KICKOFF_OVD_QTY").Value ' $H$339
    End If
    If .Range("MEETING_TASK1_REM_COUNTRY").Value <> wsSource.Range("MEETING_KICKOFF_REM_COUNTRY").Value Then
        .Range("MEETING_TASK1_REM_COUNTRY").Value = wsSource.Range("MEETING_KICKOFF_REM_COUNTRY").Value ' $M$339
    End If
    If .Range("MEETING_TASK1_REM_PCT").Value <> wsSource.Range("MEETING_KICKOFF_REM_PCT").Value Then
        .Range("MEETING_TASK1_REM_PCT").Value = wsSource.Range("MEETING_KICKOFF_REM_PCT").Value ' $N$339
    End If
    If .Range("MEETING_TASK4_OVD_JUST").Value <> wsSource.Range("MEETING_OTHER_OVD_JUST").Value Then
        .Range("MEETING_TASK4_OVD_JUST").Value = wsSource.Range("MEETING_OTHER_OVD_JUST").Value ' $I$342
    End If
    If .Range("MEETING_TASK4_OVD_QTY").Value <> wsSource.Range("MEETING_OTHER_OVD_QTY").Value Then
        .Range("MEETING_TASK4_OVD_QTY").Value = wsSource.Range("MEETING_OTHER_OVD_QTY").Value ' $H$342
    End If
    If .Range("MEETING_TASK4_REM_COUNTRY").Value <> wsSource.Range("MEETING_OTHER_REM_COUNTRY").Value Then
        .Range("MEETING_TASK4_REM_COUNTRY").Value = wsSource.Range("MEETING_OTHER_REM_COUNTRY").Value ' $M$342
    End If
    If .Range("MEETING_TASK4_REM_PCT").Value <> wsSource.Range("MEETING_OTHER_REM_PCT").Value Then
        .Range("MEETING_TASK4_REM_PCT").Value = wsSource.Range("MEETING_OTHER_REM_PCT").Value ' $N$342
    End If
    'If .Range("MEETING_OVD_QTY").Value <> wsSource.Range("MEETING_OVD_QTY").Value Then
    '    .Range("MEETING_OVD_QTY").Value = wsSource.Range("MEETING_OVD_QTY").Value ' $H$342
    'End If
    If .Range("MEETING_TASK3_OVD_JUST").Value <> wsSource.Range("MEETING_PROGRESS_OVD_JUST").Value Then
        .Range("MEETING_TASK3_OVD_JUST").Value = wsSource.Range("MEETING_PROGRESS_OVD_JUST").Value ' $I$341
    End If
    If .Range("MEETING_TASK3_OVD_QTY").Value <> wsSource.Range("MEETING_PROGRESS_OVD_QTY").Value Then
        .Range("MEETING_TASK3_OVD_QTY").Value = wsSource.Range("MEETING_PROGRESS_OVD_QTY").Value ' $H$341
    End If
    If .Range("MEETING_TASK3_REM_COUNTRY").Value <> wsSource.Range("MEETING_PROGRESS_REM_COUNTRY").Value Then
        .Range("MEETING_TASK3_REM_COUNTRY").Value = wsSource.Range("MEETING_PROGRESS_REM_COUNTRY").Value ' $M$341
    End If
    If .Range("MEETING_TASK3_REM_PCT").Value <> wsSource.Range("MEETING_PROGRESS_REM_PCT").Value Then
        .Range("MEETING_TASK3_REM_PCT").Value = wsSource.Range("MEETING_PROGRESS_REM_PCT").Value ' $N$341
    End If
    'If .Range("MEETING_REM_COST").Value <> wsSource.Range("MEETING_REM_COST").Value Then
    '    .Range("MEETING_REM_COST").Value = wsSource.Range("MEETING_PROGRESS_REM_PCT").Value ' $N$341
    'End If
    If .Range("PM_TASK2_OVD_JUST").Value <> wsSource.Range("PM_BUYOUT_OVD_JUST").Value Then
        .Range("PM_TASK2_OVD_JUST").Value = wsSource.Range("PM_BUYOUT_OVD_JUST").Value ' $I$326
    End If
    If .Range("PM_TASK2_OVD_QTY").Value <> wsSource.Range("PM_BUYOUT_OVD_QTY").Value Then
        .Range("PM_TASK2_OVD_QTY").Value = wsSource.Range("PM_BUYOUT_OVD_QTY").Value ' $H$326
    End If
    If .Range("PM_TASK2_REM_COUNTRY").Value <> wsSource.Range("PM_BUYOUT_REM_COUNTRY").Value Then
        .Range("PM_TASK2_REM_COUNTRY").Value = wsSource.Range("PM_BUYOUT_REM_COUNTRY").Value ' $M$326
    End If
    If .Range("PM_TASK2_REM_PCT").Value <> wsSource.Range("PM_BUYOUT_REM_PCT").Value Then
        .Range("PM_TASK2_REM_PCT").Value = wsSource.Range("PM_BUYOUT_REM_PCT").Value ' $N$326
    End If
    If .Range("PM_TASK2_TOTAL").Value <> wsSource.Range("PM_BUYOUT_TOTAL").Value Then
        .Range("PM_TASK2_TOTAL").Value = wsSource.Range("PM_BUYOUT_TOTAL").Value ' $E$116
    End If
    If .Range("PM_TASK1_OVD_JUST").Value <> wsSource.Range("PM_IA_OVD_JUST").Value Then
        .Range("PM_TASK1_OVD_JUST").Value = wsSource.Range("PM_IA_OVD_JUST").Value ' $I$325
    End If
    If .Range("PM_TASK1_OVD_QTY").Value <> wsSource.Range("PM_IA_OVD_QTY").Value Then
        .Range("PM_TASK1_OVD_QTY").Value = wsSource.Range("PM_IA_OVD_QTY").Value ' $H$325
    End If
    If .Range("PM_TASK1_REM_COUNTRY").Value <> wsSource.Range("PM_IA_REM_COUNTRY").Value Then
        .Range("PM_TASK1_REM_COUNTRY").Value = wsSource.Range("PM_IA_REM_COUNTRY").Value ' $M$325
    End If
    If .Range("PM_TASK1_REM_PCT").Value <> wsSource.Range("PM_IA_REM_PCT").Value Then
        .Range("PM_TASK1_REM_PCT").Value = wsSource.Range("PM_IA_REM_PCT").Value ' $N$325
    End If
    If .Range("PM_IA_TOTAL").Value <> wsSource.Range("PM_IA_TOTAL").Value Then
        .Range("PM_IA_TOTAL").Value = wsSource.Range("PM_IA_TOTAL").Value ' $E$115
    End If
    If .Range("PM_PCT_TOTAL").Value <> wsSource.Range("PM_PCT_TOTAL").Value Then
        .Range("PM_PCT_TOTAL").Value = wsSource.Range("PM_PCT_TOTAL").Value ' $E$114
    End If
    'If .Range("PM_REM_COST").Value <> wsSource.Range("PM_REM_COST").Value Then
    '    .Range("PM_REM_COST").Value = wsSource.Range("PM_PCT_TOTAL").Value ' $E$114
    'End If
    If .Range("REP_TASK3_OVD_JUST").Value <> wsSource.Range("REP_CUSTOM_OVD_JUST").Value Then
        .Range("REP_TASK3_OVD_JUST").Value = wsSource.Range("REP_CUSTOM_OVD_JUST").Value ' $I$239
    End If
    If .Range("REP_TASK3_OVD_QTY").Value <> wsSource.Range("REP_CUSTOM_OVD_QTY").Value Then
        .Range("REP_TASK3_OVD_QTY").Value = wsSource.Range("REP_CUSTOM_OVD_QTY").Value ' $H$239
    End If
    If .Range("REP_TASK3_REM_COUNTRY").Value <> wsSource.Range("REP_CUSTOM_REM_COUNTRY").Value Then
        .Range("REP_TASK3_REM_COUNTRY").Value = wsSource.Range("REP_CUSTOM_REM_COUNTRY").Value ' $M$239
    End If
    If .Range("REP_TASK3_REM_PCT").Value <> wsSource.Range("REP_CUSTOM_REM_PCT").Value Then
        .Range("REP_TASK3_REM_PCT").Value = wsSource.Range("REP_CUSTOM_REM_PCT").Value ' $N$239
    End If
    If .Range("REP_TASK1_OVD_JUST").Value <> wsSource.Range("REP_POINTS_OVD_JUST").Value Then
        .Range("REP_TASK1_OVD_JUST").Value = wsSource.Range("REP_POINTS_OVD_JUST").Value ' $I$237
    End If
    If .Range("REP_TASK1_OVD_QTY").Value <> wsSource.Range("REP_POINTS_OVD_QTY").Value Then
        .Range("REP_TASK1_OVD_QTY").Value = wsSource.Range("REP_POINTS_OVD_QTY").Value ' $H$237
    End If
    If .Range("REP_TASK1_REM_COUNTRY").Value <> wsSource.Range("REP_POINTS_REM_COUNTRY").Value Then
        .Range("REP_TASK1_REM_COUNTRY").Value = wsSource.Range("REP_POINTS_REM_COUNTRY").Value ' $M$237
    End If
    If .Range("REP_TASK1_REM_PCT").Value <> wsSource.Range("REP_POINTS_REM_PCT").Value Then
        .Range("REP_TASK1_REM_PCT").Value = wsSource.Range("REP_POINTS_REM_PCT").Value ' $N$237
    End If
    'If .Range("REP_REM_COST").Value <> wsSource.Range("REP_REM_COST").Value Then
    '    .Range("REP_REM_COST").Value = wsSource.Range("REP_POINTS_REM_PCT").Value ' $N$237
    'End If
    If .Range("REP_TASK2_OVD_JUST").Value <> wsSource.Range("REP_STD_OVD_JUST").Value Then
        .Range("REP_TASK2_OVD_JUST").Value = wsSource.Range("REP_STD_OVD_JUST").Value ' $I$238
    End If
    If .Range("REP_TASK2_OVD_QTY").Value <> wsSource.Range("REP_STD_OVD_QTY").Value Then
        .Range("REP_TASK2_OVD_QTY").Value = wsSource.Range("REP_STD_OVD_QTY").Value ' $H$238
    End If
    If .Range("REP_TASK2_REM_COUNTRY").Value <> wsSource.Range("REP_STD_REM_COUNTRY").Value Then
        .Range("REP_TASK2_REM_COUNTRY").Value = wsSource.Range("REP_STD_REM_COUNTRY").Value ' $M$238
    End If
    If .Range("REP_TASK2_REM_PCT").Value <> wsSource.Range("REP_STD_REM_PCT").Value Then
        .Range("REP_TASK2_REM_PCT").Value = wsSource.Range("REP_STD_REM_PCT").Value ' $N$238
    End If
    If .Range("SITE_TASK3_OVD_JUST").Value <> wsSource.Range("SITE_COMM_OVD_JUST").Value Then
        .Range("SITE_TASK3_OVD_JUST").Value = wsSource.Range("SITE_COMM_OVD_JUST").Value ' $I$357
    End If
    If .Range("SITE_TASK3_OVD_QTY").Value <> wsSource.Range("SITE_COMM_OVD_QTY").Value Then
        .Range("SITE_TASK3_OVD_QTY").Value = wsSource.Range("SITE_COMM_OVD_QTY").Value ' $H$357
    End If
    If .Range("SITE_TASK3_REM_COUNTRY").Value <> wsSource.Range("SITE_COMM_REM_COUNTRY").Value Then
        .Range("SITE_TASK3_REM_COUNTRY").Value = wsSource.Range("SITE_COMM_REM_COUNTRY").Value ' $M$357
    End If
    If .Range("SITE_TASK3_REM_PCT").Value <> wsSource.Range("SITE_COMM_REM_PCT").Value Then
        .Range("SITE_TASK3_REM_PCT").Value = wsSource.Range("SITE_COMM_REM_PCT").Value ' $N$357
    End If
    If .Range("SITE_TASK2_OVD_JUST").Value <> wsSource.Range("SITE_PWRUP_OVD_JUST").Value Then
        .Range("SITE_TASK2_OVD_JUST").Value = wsSource.Range("SITE_PWRUP_OVD_JUST").Value ' $I$356
    End If
    If .Range("SITE_TASK2_OVD_QTY").Value <> wsSource.Range("SITE_PWRUP_OVD_QTY").Value Then
        .Range("SITE_TASK2_OVD_QTY").Value = wsSource.Range("SITE_PWRUP_OVD_QTY").Value ' $H$356
    End If
    If .Range("SITE_TASK2_REM_COUNTRY").Value <> wsSource.Range("SITE_PWRUP_REM_COUNTRY").Value Then
        .Range("SITE_TASK2_REM_COUNTRY").Value = wsSource.Range("SITE_PWRUP_REM_COUNTRY").Value ' $M$356
    End If
    If .Range("SITE_TASK2_REM_PCT").Value <> wsSource.Range("SITE_PWRUP_REM_PCT").Value Then
        .Range("SITE_TASK2_REM_PCT").Value = wsSource.Range("SITE_PWRUP_REM_PCT").Value ' $N$356
    End If
    'If .Range("SITE_REM_COST").Value <> wsSource.Range("SITE_REM_COST").Value Then
    '    .Range("SITE_REM_COST").Value = wsSource.Range("SITE_PWRUP_REM_PCT").Value ' $N$356
    'End If
    If .Range("SITE_TASK4_OVD_JUST").Value <> wsSource.Range("SITE_SAT_OVD_JUST").Value Then
        .Range("SITE_TASK4_OVD_JUST").Value = wsSource.Range("SITE_SAT_OVD_JUST").Value ' $I$358
    End If
    If .Range("SITE_TASK4_OVD_QTY").Value <> wsSource.Range("SITE_SAT_OVD_QTY").Value Then
        .Range("SITE_TASK4_OVD_QTY").Value = wsSource.Range("SITE_SAT_OVD_QTY").Value ' $H$358
    End If
    If .Range("SITE_TASK4_REM_COUNTRY").Value <> wsSource.Range("SITE_SAT_REM_COUNTRY").Value Then
        .Range("SITE_TASK4_REM_COUNTRY").Value = wsSource.Range("SITE_SAT_REM_COUNTRY").Value ' $M$358
    End If
    If .Range("SITE_TASK4_REM_PCT").Value <> wsSource.Range("SITE_SAT_REM_PCT").Value Then
        .Range("SITE_TASK4_REM_PCT").Value = wsSource.Range("SITE_SAT_REM_PCT").Value ' $N$358
    End If
    If .Range("SITE_TASK1_OVD_JUST").Value <> wsSource.Range("SITE_SURVEY_OVD_JUST").Value Then
        .Range("SITE_TASK1_OVD_JUST").Value = wsSource.Range("SITE_SURVEY_OVD_JUST").Value ' $I$355
    End If
    If .Range("SITE_TASK1_OVD_QTY").Value <> wsSource.Range("SITE_SURVEY_OVD_QTY").Value Then
        .Range("SITE_TASK1_OVD_QTY").Value = wsSource.Range("SITE_SURVEY_OVD_QTY").Value ' $H$355
    End If
    If .Range("SITE_TASK1_REM_COUNTRY").Value <> wsSource.Range("SITE_SURVEY_REM_COUNTRY").Value Then
        .Range("SITE_TASK1_REM_COUNTRY").Value = wsSource.Range("SITE_SURVEY_REM_COUNTRY").Value ' $M$355
    End If
    If .Range("SITE_TASK1_REM_PCT").Value <> wsSource.Range("SITE_SURVEY_REM_PCT").Value Then
        .Range("SITE_TASK1_REM_PCT").Value = wsSource.Range("SITE_SURVEY_REM_PCT").Value ' $N$355
    End If
    If .Range("SPEC_TASK2_OVD_JUST").Value <> wsSource.Range("SPEC_FAT_OVD_JUST").Value Then
        .Range("SPEC_TASK2_OVD_JUST").Value = wsSource.Range("SPEC_FAT_OVD_JUST").Value ' $I$129
    End If
    If .Range("SPEC_TASK2_OVD_QTY").Value <> wsSource.Range("SPEC_FAT_OVD_QTY").Value Then
        .Range("SPEC_TASK2_OVD_QTY").Value = wsSource.Range("SPEC_FAT_OVD_QTY").Value ' $H$129
    End If
    If .Range("SPEC_TASK2_REM_COUNTRY").Value <> wsSource.Range("SPEC_FAT_REM_COUNTRY").Value Then
        .Range("SPEC_TASK2_REM_COUNTRY").Value = wsSource.Range("SPEC_FAT_REM_COUNTRY").Value ' $M$129
    End If
    If .Range("SPEC_TASK2_REM_PCT").Value <> wsSource.Range("SPEC_FAT_REM_PCT").Value Then
        .Range("SPEC_TASK2_REM_PCT").Value = wsSource.Range("SPEC_FAT_REM_PCT").Value ' $N$129
    End If
    If .Range("SPEC_TASK1_OVD_JST").Value <> wsSource.Range("SPEC_PFS_OVD_JST").Value Then
        .Range("SPEC_TASK1_OVD_JST").Value = wsSource.Range("SPEC_PFS_OVD_JST").Value ' $I$128
    End If
    If .Range("SPEC_TASK1_OVD_JUST").Value <> wsSource.Range("SPEC_PFS_OVD_JUST").Value Then
        .Range("SPEC_TASK1_OVD_JUST").Value = wsSource.Range("SPEC_PFS_OVD_JUST").Value ' $I$128
    End If
    If .Range("SPEC_TASK1_OVD_QTY").Value <> wsSource.Range("SPEC_PFS_OVD_QTY").Value Then
        .Range("SPEC_TASK1_OVD_QTY").Value = wsSource.Range("SPEC_PFS_OVD_QTY").Value ' $H$128
    End If
    If .Range("SPEC_TASK1_REM_COUNTRY").Value <> wsSource.Range("SPEC_PFS_REM_COUNTRY").Value Then
        .Range("SPEC_TASK1_REM_COUNTRY").Value = wsSource.Range("SPEC_PFS_REM_COUNTRY").Value ' $M$128
    End If
    If .Range("SPEC_TASK1_REM_PCT").Value <> wsSource.Range("SPEC_PFS_REM_PCT").Value Then
        .Range("SPEC_TASK1_REM_PCT").Value = wsSource.Range("SPEC_PFS_REM_PCT").Value ' $N$128
    End If
    If .Range("SPEC_TASK4_OVD_JUST").Value <> wsSource.Range("SPEC_QA_OVD_JUST").Value Then
        .Range("SPEC_TASK4_OVD_JUST").Value = wsSource.Range("SPEC_QA_OVD_JUST").Value ' $I$131
    End If
    If .Range("SPEC_TASK4_OVD_QTY").Value <> wsSource.Range("SPEC_QA_OVD_QTY").Value Then
        .Range("SPEC_TASK4_OVD_QTY").Value = wsSource.Range("SPEC_QA_OVD_QTY").Value ' $H$131
    End If
    If .Range("SPEC_TASK4_REM_COUNTRY").Value <> wsSource.Range("SPEC_QA_REM_COUNTRY").Value Then
        .Range("SPEC_TASK4_REM_COUNTRY").Value = wsSource.Range("SPEC_QA_REM_COUNTRY").Value ' $M$131
    End If
    If .Range("SPEC_TASK4_REM_PCT").Value <> wsSource.Range("SPEC_QA_REM_PCT").Value Then
        .Range("SPEC_TASK4_REM_PCT").Value = wsSource.Range("SPEC_QA_REM_PCT").Value ' $N$131
    End If
    If .Range("SPEC_TASK3_OVD_JUST").Value <> wsSource.Range("SPEC_SAT_OVD_JUST").Value Then
        .Range("SPEC_TASK3_OVD_JUST").Value = wsSource.Range("SPEC_SAT_OVD_JUST").Value ' $I$130
    End If
    If .Range("SPEC_TASK3_OVD_QTY").Value <> wsSource.Range("SPEC_SAT_OVD_QTY").Value Then
        .Range("SPEC_TASK3_OVD_QTY").Value = wsSource.Range("SPEC_SAT_OVD_QTY").Value ' $H$130
    End If
    If .Range("SPEC_TASK3_REM_COUNTRY").Value <> wsSource.Range("SPEC_SAT_REM_COUNTRY").Value Then
        .Range("SPEC_TASK3_REM_COUNTRY").Value = wsSource.Range("SPEC_SAT_REM_COUNTRY").Value ' $M$130
    End If
    If .Range("SPEC_TASK3_REM_PCT").Value <> wsSource.Range("SPEC_SAT_REM_PCT").Value Then
        .Range("SPEC_TASK3_REM_PCT").Value = wsSource.Range("SPEC_SAT_REM_PCT").Value ' $N$130
    End If
    If .Range("SYSENG_TASK1_OVD_JUST").Value <> wsSource.Range("SYSENG_PANEL_OVD_JUST").Value Then
        .Range("SYSENG_TASK1_OVD_JUST").Value = wsSource.Range("SYSENG_PANEL_OVD_JUST").Value ' $I$143
    End If
    If .Range("SYSENG_TASK1_OVD_QTY").Value <> wsSource.Range("SYSENG_PANEL_OVD_QTY").Value Then
        .Range("SYSENG_TASK1_OVD_QTY").Value = wsSource.Range("SYSENG_PANEL_OVD_QTY").Value ' $H$143
    End If
    If .Range("SYSENG_TASK1_REM_COUNTRY").Value <> wsSource.Range("SYSENG_PANEL_REM_COUNTRY").Value Then
        .Range("SYSENG_TASK1_REM_COUNTRY").Value = wsSource.Range("SYSENG_PANEL_REM_COUNTRY").Value ' $M$143
    End If
    If .Range("SYSENG_TASK1_REM_PCT").Value <> wsSource.Range("SYSENG_PANEL_REM_PCT").Value Then
        .Range("SYSENG_TASK1_REM_PCT").Value = wsSource.Range("SYSENG_PANEL_REM_PCT").Value ' $N$143
    End If
    'If .Range("SYSENG_REM_COST").Value <> wsSource.Range("SYSENG_REM_COST").Value Then
    '    .Range("SYSENG_REM_COST").Value = wsSource.Range("SYSENG_PANEL_REM_PCT").Value ' $N$143
    'End If
    If .Range("SYSENG_TASK2_OVD_JUST").Value <> wsSource.Range("SYSENG_SYS_CFG_OVD_JUST").Value Then
        .Range("SYSENG_TASK2_OVD_JUST").Value = wsSource.Range("SYSENG_SYS_CFG_OVD_JUST").Value ' $I$144
    End If
    If .Range("SYSENG_TASK2_OVD_QTY").Value <> wsSource.Range("SYSENG_SYS_CFG_OVD_QTY").Value Then
        .Range("SYSENG_TASK2_OVD_QTY").Value = wsSource.Range("SYSENG_SYS_CFG_OVD_QTY").Value ' $H$144
    End If
    If .Range("SYSENG_TASK2_REM_COUNTRY").Value <> wsSource.Range("SYSENG_SYS_CFG_REM_COUNTRY").Value Then
        .Range("SYSENG_TASK2_REM_COUNTRY").Value = wsSource.Range("SYSENG_SYS_CFG_REM_COUNTRY").Value ' $M$144
    End If
    If .Range("SYSENG_TASK2_REM_PCT").Value <> wsSource.Range("SYSENG_SYS_CFG_REM_PCT").Value Then
        .Range("SYSENG_TASK2_REM_PCT").Value = wsSource.Range("SYSENG_SYS_CFG_REM_PCT").Value ' $N$144
    End If
    
    'TEST
    If .Range("TEST_TASK1_OVD_JUST").Value <> wsSource.Range("TEST_FAT_OVD_JUST").Value Then
        .Range("TEST_TASK1_OVD_JUST").Value = wsSource.Range("TEST_FAT_OVD_JUST").Value ' $I$271
    End If
    If .Range("TEST_TASK1_OVD_QTY").Value <> wsSource.Range("TEST_FAT_OVD_QTY").Value Then
        .Range("TEST_TASK1_OVD_QTY").Value = wsSource.Range("TEST_FAT_OVD_QTY").Value ' $H$271
    End If
    If .Range("TEST_TASK1_REM_COUNTRY").Value <> wsSource.Range("TEST_FAT_REM_COUNTRY").Value Then
        .Range("TEST_TASK1_REM_COUNTRY").Value = wsSource.Range("TEST_FAT_REM_COUNTRY").Value ' $M$271
    End If
    If .Range("TEST_TASK1_REM_PCT").Value <> wsSource.Range("TEST_FAT_REM_PCT").Value Then
        .Range("TEST_TASK1_REM_PCT").Value = wsSource.Range("TEST_FAT_REM_PCT").Value ' $N$271
    End If
    If .Range("TEST_TASK2_OVD_JUST").Value <> wsSource.Range("TEST_IO_OVD_JUST").Value Then
        .Range("TEST_TASK2_OVD_JUST").Value = wsSource.Range("TEST_IO_OVD_JUST").Value ' $I$272
    End If
    If .Range("TEST_TASK2_OVD_QTY").Value <> wsSource.Range("TEST_IO_OVD_QTY").Value Then
        .Range("TEST_TASK2_OVD_QTY").Value = wsSource.Range("TEST_IO_OVD_QTY").Value ' $H$272
    End If
    If .Range("TEST_TASK2_REM_COUNTRY").Value <> wsSource.Range("TEST_IO_REM_COUNTRY").Value Then
        .Range("TEST_TASK2_REM_COUNTRY").Value = wsSource.Range("TEST_IO_REM_COUNTRY").Value ' $M$272
    End If
    If .Range("TEST_TASK2_REM_PCT").Value <> wsSource.Range("TEST_IO_REM_PCT").Value Then
        .Range("TEST_TASK2_REM_PCT").Value = wsSource.Range("TEST_IO_REM_PCT").Value ' $N$272
    End If
    If .Range("TEST_TASK4_OVD_JUST").Value <> wsSource.Range("TEST_PACK_OVD_JUST").Value Then
        .Range("TEST_TASK4_OVD_JUST").Value = wsSource.Range("TEST_PACK_OVD_JUST").Value ' $I$274
    End If
    If .Range("TEST_TASK4_OVD_QTY").Value <> wsSource.Range("TEST_PACK_OVD_QTY").Value Then
        .Range("TEST_TASK4_OVD_QTY").Value = wsSource.Range("TEST_PACK_OVD_QTY").Value ' $H$274
    End If
    If .Range("TEST_TASK4_REM_COUNTRY").Value <> wsSource.Range("TEST_PACK_REM_COUNTRY").Value Then
        .Range("TEST_TASK4_REM_COUNTRY").Value = wsSource.Range("TEST_PACK_REM_COUNTRY").Value ' $M$274
    End If
    If .Range("TEST_TASK4_REM_PCT").Value <> wsSource.Range("TEST_PACK_REM_PCT").Value Then
        .Range("TEST_TASK4_REM_PCT").Value = wsSource.Range("TEST_PACK_REM_PCT").Value ' $N$274
    End If
    'If .Range("TEST_REM_COST").Value <> wsSource.Range("TEST_REM_COST").Value Then
    '    .Range("TEST_REM_COST").Value = wsSource.Range("TEST_PACK_REM_PCT").Value ' $N$274
    'End If
    If .Range("TEST_TASK6_OVD_JUST").Value <> wsSource.Range("TEST_RENT_OVD_JUST").Value Then
        .Range("TEST_TASK6_OVD_JUST").Value = wsSource.Range("TEST_RENT_OVD_JUST").Value ' $I$276
    End If
    If .Range("TEST_TASK6_OVD_QTY").Value <> wsSource.Range("TEST_RENT_OVD_QTY").Value Then
        .Range("TEST_TASK6_OVD_QTY").Value = wsSource.Range("TEST_RENT_OVD_QTY").Value ' $H$276
    End If
    If .Range("TEST_TASK6_REM_COUNTRY").Value <> wsSource.Range("TEST_RENT_REM_COUNTRY").Value Then
        .Range("TEST_TASK6_REM_COUNTRY").Value = wsSource.Range("TEST_RENT_REM_COUNTRY").Value ' $M$276
    End If
    If .Range("TEST_TASK6_REM_PCT").Value <> wsSource.Range("TEST_RENT_REM_PCT").Value Then
        .Range("TEST_TASK6_REM_PCT").Value = wsSource.Range("TEST_RENT_REM_PCT").Value ' $N$276
    End If
    If .Range("TEST_TASK3_OVD_JUST").Value <> wsSource.Range("TEST_SI_OVD_JUST").Value Then
        .Range("TEST_TASK3_OVD_JUST").Value = wsSource.Range("TEST_SI_OVD_JUST").Value ' $I$273
    End If
    If .Range("TEST_TASK3_OVD_QTY").Value <> wsSource.Range("TEST_SI_OVD_QTY").Value Then
        .Range("TEST_TASK3_OVD_QTY").Value = wsSource.Range("TEST_SI_OVD_QTY").Value ' $H$273
    End If
    If .Range("TEST_TASK3_REM_COUNTRY").Value <> wsSource.Range("TEST_SI_REM_COUNTRY").Value Then
        .Range("TEST_TASK3_REM_COUNTRY").Value = wsSource.Range("TEST_SI_REM_COUNTRY").Value ' $M$273
    End If
    If .Range("TEST_TASK3_REM_PCT").Value <> wsSource.Range("TEST_SI_REM_PCT").Value Then
        .Range("TEST_TASK3_REM_PCT").Value = wsSource.Range("TEST_SI_REM_PCT").Value ' $N$273
    End If
    If .Range("TEST_TASK5_OVD_JUST").Value <> wsSource.Range("TEST_SIM_OVD_JUST").Value Then
        .Range("TEST_TASK5_OVD_JUST").Value = wsSource.Range("TEST_SIM_OVD_JUST").Value ' $I$275
    End If
    If .Range("TEST_TASK5_OVD_QTY").Value <> wsSource.Range("TEST_SIM_OVD_QTY").Value Then
        .Range("TEST_TASK5_OVD_QTY").Value = wsSource.Range("TEST_SIM_OVD_QTY").Value ' $H$275
    End If
    If .Range("TEST_TASK5_REM_COUNTRY").Value <> wsSource.Range("TEST_SIM_REM_COUNTRY").Value Then
        .Range("TEST_TASK5_REM_COUNTRY").Value = wsSource.Range("TEST_SIM_REM_COUNTRY").Value ' $M$275
    End If
    If .Range("TEST_TASK5_REM_PCT").Value <> wsSource.Range("TEST_SIM_REM_PCT").Value Then
        .Range("TEST_TASK5_REM_PCT").Value = wsSource.Range("TEST_SIM_REM_PCT").Value ' $N$275
    End If
    
    'TL
    If .Range("TL_TASK1_OVD_JUST").Value <> wsSource.Range("TL_ENTER_OVD_JUST").Value Then
        .Range("TL_TASK1_OVD_JUST").Value = wsSource.Range("TL_ENTER_OVD_JUST").Value ' $I$369
    End If
    If .Range("TL_TASK1_OVD_QTY").Value <> wsSource.Range("TL_ENTER_OVD_QTY").Value Then
        .Range("TL_TASK1_OVD_QTY").Value = wsSource.Range("TL_ENTER_OVD_QTY").Value ' $H$369
    End If
    If .Range("TL_TASK3_OVD_QTY").Value <> wsSource.Range("TL_OVD_QTY").Value Then
        .Range("TL_TASK3_OVD_QTY").Value = wsSource.Range("TL_OVD_QTY").Value ' $H$371
    End If
    If .Range("TL_TASK3_OVD_JUST").Value <> wsSource.Range("TL_REM_OVD_JUST").Value Then
        .Range("TL_TASK3_OVD_JUST").Value = wsSource.Range("TL_REM_OVD_JUST").Value ' $I$371
    End If
    If .Range("TL_TASK4_OVD_JUST").Value <> wsSource.Range("TL_SERVICES_OVD_JUST").Value Then
        .Range("TL_TASK4_OVD_JUST").Value = wsSource.Range("TL_SERVICES_OVD_JUST").Value ' $I$372
    End If
    If .Range("TL_TASK4_OVD_QTY").Value <> wsSource.Range("TL_SITE_OVD_QTY").Value Then
        .Range("TL_TASK4_OVD_QTY").Value = wsSource.Range("TL_SITE_OVD_QTY").Value ' $H$372
    End If
    If .Range("TL_TASK1_OVD_JUST").Value <> wsSource.Range("TL_TL_OVD_JUST").Value Then
        .Range("TL_TASK1_OVD_JUST").Value = wsSource.Range("TL_TL_OVD_JUST").Value ' $I$370
    End If
    If .Range("TL_TASK1_OVD_QTY").Value <> wsSource.Range("TL_TL_OVD_QTY").Value Then
        .Range("TL_TASK1_OVD_QTY").Value = wsSource.Range("TL_TL_OVD_QTY").Value ' $H$370
    End If
End With
 '387

End Sub




'assumption and data entry sheet are done with hard coded script because of exceptions
Sub ImportAssumptionsProposalSheet(wsSource As Worksheet, wsTarget As Worksheet)
'GECE Version = 1.0
On Error Resume Next
With wsTarget
    'If .Range("SCOPE_LANGUAGE").Value <> wsSource.Range("SCOPE_LANGUAGE").Value Then
    '    .Range("SCOPE_LANGUAGE").Value = wsSource.Range("SCOPE_LANGUAGE").Value ' $D$63
    'End If
        
    'Version 1.0
    If .Range("TOOLKIT_1_REQ").Value <> wsSource.Range("TOOLKIT_1_REQ").Value Then
        .Range("TOOLKIT_1_REQ").Value = wsSource.Range("TOOLKIT_1_REQ").Value
    End If
    If .Range("TOOLKIT_2_REQ").Value <> wsSource.Range("TOOLKIT_2_REQ").Value Then
        .Range("TOOLKIT_2_REQ").Value = wsSource.Range("TOOLKIT_2_REQ").Value
    End If
    If .Range("TOOLKIT_3_REQ").Value <> wsSource.Range("TOOLKIT_3_REQ").Value Then
        .Range("TOOLKIT_3_REQ").Value = wsSource.Range("TOOLKIT_3_REQ").Value
    End If
    If .Range("TOOLKIT_4_REQ").Value <> wsSource.Range("TOOLKIT_4_REQ").Value Then
        .Range("TOOLKIT_4_REQ").Value = wsSource.Range("TOOLKIT_4_REQ").Value
    End If
    If .Range("TOOLKIT_5_REQ").Value <> wsSource.Range("TOOLKIT_5_REQ").Value Then
        .Range("TOOLKIT_5_REQ").Value = wsSource.Range("TOOLKIT_5_REQ").Value
    End If
    If .Range("TOOLKIT_6_REQ").Value <> wsSource.Range("TOOLKIT_6_REQ").Value Then
        .Range("TOOLKIT_6_REQ").Value = wsSource.Range("TOOLKIT_6_REQ").Value
    End If
    If .Range("TOOLKIT_7_REQ").Value <> wsSource.Range("TOOLKIT_7_REQ").Value Then
        .Range("TOOLKIT_7_REQ").Value = wsSource.Range("TOOLKIT_7_REQ").Value
    End If
    If .Range("TOOLKIT_8_REQ").Value <> wsSource.Range("TOOLKIT_8_REQ").Value Then
        .Range("TOOLKIT_8_REQ").Value = wsSource.Range("TOOLKIT_8_REQ").Value
    End If
    If .Range("TOOLKIT_9_REQ").Value <> wsSource.Range("TOOLKIT_9_REQ").Value Then
        .Range("TOOLKIT_9_REQ").Value = wsSource.Range("TOOLKIT_9_REQ").Value
    End If
    If .Range("TOOLKIT_10_REQ").Value <> wsSource.Range("TOOLKIT_10_REQ").Value Then
        .Range("TOOLKIT_10_REQ").Value = wsSource.Range("TOOLKIT_10_REQ").Value
    End If
    If .Range("TOOLKIT_11_REQ").Value <> wsSource.Range("TOOLKIT_11_REQ").Value Then
        .Range("TOOLKIT_11_REQ").Value = wsSource.Range("TOOLKIT_11_REQ").Value
    End If
    If .Range("TOOLKIT_12_REQ").Value <> wsSource.Range("TOOLKIT_12_REQ").Value Then
        .Range("TOOLKIT_12_REQ").Value = wsSource.Range("TOOLKIT_12_REQ").Value
    End If
    If .Range("TOOLKIT_13_REQ").Value <> wsSource.Range("TOOLKIT_13_REQ").Value Then
        .Range("TOOLKIT_13_REQ").Value = wsSource.Range("TOOLKIT_13_REQ").Value
    End If
    If .Range("TOOLKIT_14_REQ").Value <> wsSource.Range("TOOLKIT_14_REQ").Value Then
        .Range("TOOLKIT_14_REQ").Value = wsSource.Range("TOOLKIT_14_REQ").Value
    End If
    If .Range("TOOLKIT_15_REQ").Value <> wsSource.Range("TOOLKIT_15_REQ").Value Then
        .Range("TOOLKIT_15_REQ").Value = wsSource.Range("TOOLKIT_15_REQ").Value
    End If
    If .Range("TOOLKIT_16_REQ").Value <> wsSource.Range("TOOLKIT_16_REQ").Value Then
        .Range("TOOLKIT_16_REQ").Value = wsSource.Range("TOOLKIT_16_REQ").Value
    End If
    If .Range("TOOLKIT_17_REQ").Value <> wsSource.Range("TOOLKIT_17_REQ").Value Then
        .Range("TOOLKIT_17_REQ").Value = wsSource.Range("TOOLKIT_17_REQ").Value
    End If
    If .Range("TOOLKIT_18_REQ").Value <> wsSource.Range("TOOLKIT_18_REQ").Value Then
        .Range("TOOLKIT_18_REQ").Value = wsSource.Range("TOOLKIT_18_REQ").Value
    End If
    If .Range("TOOLKIT_20_REQ").Value <> wsSource.Range("TOOLKIT_20_REQ").Value Then
        .Range("TOOLKIT_20_REQ").Value = wsSource.Range("TOOLKIT_20_REQ").Value
    End If
    If .Range("TOOLKIT_21_REQ").Value <> wsSource.Range("TOOLKIT_21_REQ").Value Then
        .Range("TOOLKIT_21_REQ").Value = wsSource.Range("TOOLKIT_21_REQ").Value
    End If
    If .Range("TOOLKIT_22_REQ").Value <> wsSource.Range("TOOLKIT_22_REQ").Value Then
        .Range("TOOLKIT_22_REQ").Value = wsSource.Range("TOOLKIT_22_REQ").Value
    End If
    If .Range("TOOLKIT_23_REQ").Value <> wsSource.Range("TOOLKIT_23_REQ").Value Then
        .Range("TOOLKIT_23_REQ").Value = wsSource.Range("TOOLKIT_23_REQ").Value
    End If
    
    
    If .Range("APP_NOTES").Value <> wsSource.Range("APP_NOTES").Value Then
        .Range("APP_NOTES").Value = wsSource.Range("APP_NOTES").Value ' $D$63
    End If
    If .Range("ASSUMPTIONS").Value <> wsSource.Range("ASSUMPTIONS").Value Then
        .Range("ASSUMPTIONS").Value = wsSource.Range("ASSUMPTIONS").Value ' $G$25
    End If
    If .Range("CP_NOTES").Value <> wsSource.Range("CP_NOTES").Value Then
        .Range("CP_NOTES").Value = wsSource.Range("CP_NOTES").Value ' $D$47
    End If
    If .Range("CUSTOMER_NAME").Value <> wsSource.Range("CUSTOMER_NAME").Value Then
        .Range("CUSTOMER_NAME").Value = wsSource.Range("CUSTOMER_NAME").Value ' $E$11
    End If
    If .Range("CUSTOMER_TYPE").Value <> wsSource.Range("CUSTOMER_TYPE").Value Then
        .Range("CUSTOMER_TYPE").Value = wsSource.Range("CUSTOMER_TYPE").Value ' $H$20
    End If
    If .Range("DEFAULT_REM_COUNTRY").Value <> wsSource.Range("DEFAULT_REM_COUNTRY").Value Then
        .Range("DEFAULT_REM_COUNTRY").Value = wsSource.Range("DEFAULT_REM_COUNTRY").Value ' $H$17
    End If
    If .Range("DI_NOTES").Value <> wsSource.Range("DI_NOTES").Value Then
        .Range("DI_NOTES").Value = wsSource.Range("DI_NOTES").Value ' $D$51
    End If
    If .Range("DOC_NOTES").Value <> wsSource.Range("DOC_NOTES").Value Then
        .Range("DOC_NOTES").Value = wsSource.Range("DOC_NOTES").Value ' $D$79
    End If
    If .Range("GSP_ID").Value <> wsSource.Range("GSP_ID").Value Then
        .Range("GSP_ID").Value = wsSource.Range("GSP_ID").Value ' $I$12
    End If
    If .Range("INDUSTRY").Value <> wsSource.Range("INDUSTRY").Value Then
        .Range("INDUSTRY").Value = wsSource.Range("INDUSTRY").Value ' $D$20
    End If
    If .Range("LOCAL_COUNTRY").Value <> wsSource.Range("LOCAL_COUNTRY").Value Then
        .Range("LOCAL_COUNTRY").Value = wsSource.Range("LOCAL_COUNTRY").Value ' $D$17
    End If
    If .Range("MEETING_NOTES").Value <> wsSource.Range("MEETING_NOTES").Value Then
        .Range("MEETING_NOTES").Value = wsSource.Range("MEETING_NOTES").Value ' $D$83
    End If
    If .Range("PROJECT_MANAGER").Value <> wsSource.Range("PROJECT_MANAGER").Value Then
        .Range("PROJECT_MANAGER").Value = wsSource.Range("PROJECT_MANAGER").Value ' $I$11
    End If
    If .Range("PROJECT_NAME").Value <> wsSource.Range("PROJECT_NAME").Value Then
        .Range("PROJECT_NAME").Value = wsSource.Range("PROJECT_NAME").Value ' $E$12
    End If
    If .Range("PROPOSAL_NUMBER").Value <> wsSource.Range("PROPOSAL_NUMBER").Value Then
        .Range("PROPOSAL_NUMBER").Value = wsSource.Range("PROPOSAL_NUMBER").Value ' $E$10
    End If
    If .Range("PROPOSAL_REVISION").Value <> wsSource.Range("PROPOSAL_REVISION").Value Then
        .Range("PROPOSAL_REVISION").Value = wsSource.Range("PROPOSAL_REVISION").Value ' $E$13
    End If
    If .Range("PROPSAL_DATE").Value <> wsSource.Range("PROPSAL_DATE").Value Then
        .Range("PROPSAL_DATE").Value = wsSource.Range("PROPSAL_DATE").Value ' $E$14
    End If
    If .Range("PROPSAL_ENGINEER").Value <> wsSource.Range("PROPSAL_ENGINEER").Value Then
        .Range("PROPSAL_ENGINEER").Value = wsSource.Range("PROPSAL_ENGINEER").Value ' $I$10
    End If
    If .Range("REPORT_NOTES").Value <> wsSource.Range("REPORT_NOTES").Value Then
        .Range("REPORT_NOTES").Value = wsSource.Range("REPORT_NOTES").Value ' $D$59
    End If
    If .Range("SITE_NOTES").Value <> wsSource.Range("SITE_NOTES").Value Then
        .Range("SITE_NOTES").Value = wsSource.Range("SITE_NOTES").Value ' $D$71
    End If
    If .Range("SYSTEM_NOTES").Value <> wsSource.Range("SYSTEM_NOTES").Value Then
        .Range("SYSTEM_NOTES").Value = wsSource.Range("SYSTEM_NOTES").Value ' $D$91
    End If
    If .Range("TEST_NOTES").Value <> wsSource.Range("TEST_NOTES").Value Then
        .Range("TEST_NOTES").Value = wsSource.Range("TEST_NOTES").Value ' $D$75
    End If
    If .Range("TL_NOTES").Value <> wsSource.Range("TL_NOTES").Value Then
        .Range("TL_NOTES").Value = wsSource.Range("TL_NOTES").Value ' $D$87
    End If
    If .Range("TRAIN_NOTES").Value <> wsSource.Range("TRAIN_NOTES").Value Then
        .Range("TRAIN_NOTES").Value = wsSource.Range("TRAIN_NOTES").Value ' $D$67
    End If
    If .Range("TRICON_NOTES").Value <> wsSource.Range("TRICON_NOTES").Value Then
        .Range("TRICON_NOTES").Value = wsSource.Range("TRICON_NOTES").Value ' $D$55
    End If
End With
 '25
End Sub

Sub ImportDataEntrySheet(wsSource As Worksheet, wsTarget As Worksheet)
On Error Resume Next
'GECE Version = 1.0

With wsTarget
    
    If .Range("HIGH_RISK_SITE").Value <> wsSource.Range("HIGH_RISK_SITE").Value Then
        .Range("HIGH_RISK_SITE").Value = wsSource.Range("HIGH_RISK_SITE").Value
    End If
    If .Range("TL_TRIPS_REQ_FAT").Value <> wsSource.Range("TL_TRIPS_REQ_FAT").Value Then
        .Range("TL_TRIPS_REQ_FAT").Value = wsSource.Range("TL_TRIPS_REQ_FAT").Value
    End If
    If .Range("TL_PERS_QTY_FAT").Value <> wsSource.Range("TL_PERS_QTY_FAT").Value Then
        .Range("TL_PERS_QTY_FAT").Value = wsSource.Range("TL_PERS_QTY_FAT").Value
    End If
    If .Range("TL_DAYS_QTY_FAT").Value <> wsSource.Range("TL_DAYS_QTY_FAT").Value Then
        .Range("TL_DAYS_QTY_FAT").Value = wsSource.Range("TL_DAYS_QTY_FAT").Value
    End If
    If .Range("TL_AIRFARE_FAT").Value <> wsSource.Range("TL_AIRFARE_FAT").Value Then
        .Range("TL_AIRFARE_FAT").Value = wsSource.Range("TL_AIRFARE_FAT").Value
    End If
    If .Range("TL_DAILY_ALLOW_SITE").Value <> wsSource.Range("TL_DAILY_ALLOW_SITE").Value Then
        .Range("TL_DAILY_ALLOW_SITE").Value = wsSource.Range("TL_DAILY_ALLOW_SITE").Value
    End If
    If .Range("TL_TRIPS_REQ_SITE").Value <> wsSource.Range("TL_TRIPS_REQ_SITE").Value Then
        .Range("TL_TRIPS_REQ_SITE").Value = wsSource.Range("TL_TRIPS_REQ_SITE").Value
    End If
    If .Range("TL_PERS_QTY_SITE").Value <> wsSource.Range("TL_PERS_QTY_SITE").Value Then
        .Range("TL_PERS_QTY_SITE").Value = wsSource.Range("TL_PERS_QTY_SITE").Value
    End If
    If .Range("TL_DAYS_QTY_SITE").Value <> wsSource.Range("TL_DAYS_QTY_SITE").Value Then
        .Range("TL_DAYS_QTY_SITE").Value = wsSource.Range("TL_DAYS_QTY_SITE").Value
    End If
    If .Range("TL_AIRFARE_SITE").Value <> wsSource.Range("TL_AIRFARE_SITE").Value Then
        .Range("TL_AIRFARE_SITE").Value = wsSource.Range("TL_AIRFARE_SITE").Value
    End If
    If .Range("LongPtr_PCT").Value <> wsSource.Range("LongPtr_PCT").Value Then
        .Range("LongPtr_PCT").Value = wsSource.Range("LongPtr_PCT").Value
    End If
    If .Range("LongPtr_PCT_ESD").Value <> wsSource.Range("LongPtr_PCT_ESD").Value Then
        .Range("LongPtr_PCT_ESD").Value = wsSource.Range("LongPtr_PCT_ESD").Value
    End If
    If .Range("MARSH_CAB_REQ").Value <> wsSource.Range("MARSH_CAB_REQ").Value Then
        .Range("MARSH_CAB_REQ").Value = wsSource.Range("MARSH_CAB_REQ").Value
    End If
    If .Range("MarshWiring").Value <> wsSource.Range("MarshWiring").Value Then
        .Range("MarshWiring").Value = wsSource.Range("MarshWiring").Value
    End If
    If .Range("MarshWiring_ESD").Value <> wsSource.Range("MarshWiring_ESD").Value Then
        .Range("MarshWiring_ESD").Value = wsSource.Range("MarshWiring_ESD").Value
    End If
    If .Range("STAGEB_CUSTOM").Value <> wsSource.Range("STAGEB_CUSTOM").Value Then
        .Range("STAGEB_CUSTOM").Value = wsSource.Range("STAGEB_CUSTOM").Value
    End If
    If .Range("STAGEB_CUSTOM_ESD").Value <> wsSource.Range("STAGEB_CUSTOM_ESD").Value Then
        .Range("STAGEB_CUSTOM_ESD").Value = wsSource.Range("STAGEB_CUSTOM_ESD").Value
    End If
    If .Range("STAGEC_CUSTOM").Value <> wsSource.Range("STAGEC_CUSTOM").Value Then
        .Range("STAGEC_CUSTOM").Value = wsSource.Range("STAGEC_CUSTOM").Value
    End If
    If .Range("STAGEC_CUSTOM_ESD").Value <> wsSource.Range("STAGEC_CUSTOM_ESD").Value Then
        .Range("STAGEC_CUSTOM_ESD").Value = wsSource.Range("STAGEC_CUSTOM_ESD").Value
    End If
    If .Range("CABINET_VENDOR").Value <> wsSource.Range("CABINET_VENDOR").Value Then
        .Range("CABINET_VENDOR").Value = wsSource.Range("CABINET_VENDOR").Value
    End If

    '+JFR+ 4.34 - May 29th, 2008
    If .Range("LongPtr_PCT").Value <> wsSource.Range("LongPtr_PCT").Value Then
        .Range("LongPtr_PCT").Value = wsSource.Range("LongPtr_PCT").Value '
    End If
    If .Range("STAGEB_CUSTOM").Value <> wsSource.Range("STAGEB_CUSTOM").Value Then
        .Range("STAGEB_CUSTOM").Value = wsSource.Range("STAGEB_CUSTOM").Value '
    End If
    If .Range("STAGEC_CUSTOM").Value <> wsSource.Range("STAGEC_CUSTOM").Value Then
        .Range("STAGEC_CUSTOM").Value = wsSource.Range("STAGEC_CUSTOM").Value '
    End If
    '-JFR- 4.34 - May 29th, 2008
    
    If .Range("INCLUDE_ESCALATION").Value <> wsSource.Range("INCLUDE_ESCALATION").Value Then
        .Range("INCLUDE_ESCALATION").Value = wsSource.Range("INCLUDE_ESCALATION").Value '
    End If
    If .Range("APP_1_HOURS").Value <> wsSource.Range("APP_1_HOURS").Value Then
        .Range("APP_1_HOURS").Value = wsSource.Range("APP_1_HOURS").Value ' $E$93
    End If
    If .Range("APP_1_NAME").Value <> wsSource.Range("APP_1_NAME").Value Then
        .Range("APP_1_NAME").Value = wsSource.Range("APP_1_NAME").Value ' $D$93
    End If
    If .Range("APP_2_HOURS").Value <> wsSource.Range("APP_2_HOURS").Value Then
        .Range("APP_2_HOURS").Value = wsSource.Range("APP_2_HOURS").Value ' $E$94
    End If
    If .Range("APP_2_NAME").Value <> wsSource.Range("APP_2_NAME").Value Then
        .Range("APP_2_NAME").Value = wsSource.Range("APP_2_NAME").Value ' $D$94
    End If
    If .Range("APP_3_HOURS").Value <> wsSource.Range("APP_3_HOURS").Value Then
        .Range("APP_3_HOURS").Value = wsSource.Range("APP_3_HOURS").Value ' $E$95
    End If
    If .Range("APP_3_NAME").Value <> wsSource.Range("APP_3_NAME").Value Then
        .Range("APP_3_NAME").Value = wsSource.Range("APP_3_NAME").Value ' $D$95
    End If
    If .Range("APP_4_HOURS").Value <> wsSource.Range("APP_4_HOURS").Value Then
        .Range("APP_4_HOURS").Value = wsSource.Range("APP_4_HOURS").Value ' $E$96
    End If
    If .Range("APP_4_NAME").Value <> wsSource.Range("APP_4_NAME").Value Then
        .Range("APP_4_NAME").Value = wsSource.Range("APP_4_NAME").Value ' $D$96
    End If
    If .Range("APP_5_HOURS").Value <> wsSource.Range("APP_5_HOURS").Value Then
        .Range("APP_5_HOURS").Value = wsSource.Range("APP_5_HOURS").Value ' $E$97
    End If
    If .Range("APP_5_NAME").Value <> wsSource.Range("APP_5_NAME").Value Then
        .Range("APP_5_NAME").Value = wsSource.Range("APP_5_NAME").Value ' $D$97
    End If
    If .Range("APP_6_HOURS").Value <> wsSource.Range("APP_6_HOURS").Value Then
        .Range("APP_6_HOURS").Value = wsSource.Range("APP_6_HOURS").Value ' $E$98
    End If
    If .Range("APP_6_NAME").Value <> wsSource.Range("APP_6_NAME").Value Then
        .Range("APP_6_NAME").Value = wsSource.Range("APP_6_NAME").Value ' $D$98
    End If
    If .Range("APP_7_HOURS").Value <> wsSource.Range("APP_7_HOURS").Value Then
        .Range("APP_7_HOURS").Value = wsSource.Range("APP_7_HOURS").Value ' $E$99
    End If
    If .Range("APP_7_NAME").Value <> wsSource.Range("APP_7_NAME").Value Then
        .Range("APP_7_NAME").Value = wsSource.Range("APP_7_NAME").Value ' $D$99
    End If
    If .Range("APP_8_HOURS").Value <> wsSource.Range("APP_8_HOURS").Value Then
        .Range("APP_8_HOURS").Value = wsSource.Range("APP_8_HOURS").Value ' $E$100
    End If
    If .Range("APP_8_NAME").Value <> wsSource.Range("APP_8_NAME").Value Then
        .Range("APP_8_NAME").Value = wsSource.Range("APP_8_NAME").Value ' $D$100
    End If
    If .Range("APP_BUS_STS").Value <> wsSource.Range("APP_BUS_STS").Value Then
        .Range("APP_BUS_STS").Value = wsSource.Range("APP_BUS_STS").Value ' $H$87
    End If
    If .Range("CAB_CONSOLES_QTY").Value <> wsSource.Range("CAB_CONSOLES_QTY").Value Then
        .Range("CAB_CONSOLES_QTY").Value = wsSource.Range("CAB_CONSOLES_QTY").Value ' $E$180
    End If
    If .Range("CAB_IO_QTY").Value <> wsSource.Range("CAB_IO_QTY").Value Then
        .Range("CAB_IO_QTY").Value = wsSource.Range("CAB_IO_QTY").Value ' $E$178
    End If
    If .Range("CAB_MARSH_QTY").Value <> wsSource.Range("CAB_MARSH_QTY").Value Then
        .Range("CAB_MARSH_QTY").Value = wsSource.Range("CAB_MARSH_QTY").Value ' $E$179
    End If
    If .Range("CAB_PROC_QTY").Value <> wsSource.Range("CAB_PROC_QTY").Value Then
        .Range("CAB_PROC_QTY").Value = wsSource.Range("CAB_PROC_QTY").Value ' $E$177
    End If
    If .Range("COST_BUYOUT").Value <> wsSource.Range("COST_BUYOUT").Value Then
        .Range("COST_BUYOUT").Value = wsSource.Range("COST_BUYOUT").Value ' $E$184
    End If
    If .Range("COST_IA").Value <> wsSource.Range("COST_IA").Value Then
        .Range("COST_IA").Value = wsSource.Range("COST_IA").Value ' $E$183
    End If
    If .Range("COURSE_1_HOURS").Value <> wsSource.Range("COURSE_1_HOURS").Value Then
        .Range("COURSE_1_HOURS").Value = wsSource.Range("COURSE_1_HOURS").Value ' $E$106
    End If
    If .Range("COURSE_1_NAME").Value <> wsSource.Range("COURSE_1_NAME").Value Then
        .Range("COURSE_1_NAME").Value = wsSource.Range("COURSE_1_NAME").Value ' $D$106
    End If
    If .Range("COURSE_2_HOURS").Value <> wsSource.Range("COURSE_2_HOURS").Value Then
        .Range("COURSE_2_HOURS").Value = wsSource.Range("COURSE_2_HOURS").Value ' $E$107
    End If
    If .Range("COURSE_2_NAME").Value <> wsSource.Range("COURSE_2_NAME").Value Then
        .Range("COURSE_2_NAME").Value = wsSource.Range("COURSE_2_NAME").Value ' $D$107
    End If
    If .Range("COURSE_3_HOURS").Value <> wsSource.Range("COURSE_3_HOURS").Value Then
        .Range("COURSE_3_HOURS").Value = wsSource.Range("COURSE_3_HOURS").Value ' $E$108
    End If
    If .Range("COURSE_3_NAME").Value <> wsSource.Range("COURSE_3_NAME").Value Then
        .Range("COURSE_3_NAME").Value = wsSource.Range("COURSE_3_NAME").Value ' $D$108
    End If
    If .Range("COURSE_4_HOURS").Value <> wsSource.Range("COURSE_4_HOURS").Value Then
        .Range("COURSE_4_HOURS").Value = wsSource.Range("COURSE_4_HOURS").Value ' $E$109
    End If
    If .Range("COURSE_4_NAME").Value <> wsSource.Range("COURSE_4_NAME").Value Then
        .Range("COURSE_4_NAME").Value = wsSource.Range("COURSE_4_NAME").Value ' $D$109
    End If
    If .Range("COURSE_5_HOURS").Value <> wsSource.Range("COURSE_5_HOURS").Value Then
        .Range("COURSE_5_HOURS").Value = wsSource.Range("COURSE_5_HOURS").Value ' $E$110
    End If
    If .Range("COURSE_5_NAME").Value <> wsSource.Range("COURSE_5_NAME").Value Then
        .Range("COURSE_5_NAME").Value = wsSource.Range("COURSE_5_NAME").Value ' $D$110
    End If
    If .Range("COURSE1_HOURS").Value <> wsSource.Range("COURSE1_HOURS").Value Then
        .Range("COURSE1_HOURS").Value = wsSource.Range("COURSE1_HOURS").Value ' $E$106
    End If
    If .Range("COURSE1_NAME").Value <> wsSource.Range("COURSE1_NAME").Value Then
        .Range("COURSE1_NAME").Value = wsSource.Range("COURSE1_NAME").Value ' $D$106
    End If
    If .Range("COURSE2_HOURS").Value <> wsSource.Range("COURSE2_HOURS").Value Then
        .Range("COURSE2_HOURS").Value = wsSource.Range("COURSE2_HOURS").Value ' $E$107
    End If
    If .Range("COURSE2_NAME").Value <> wsSource.Range("COURSE2_NAME").Value Then
        .Range("COURSE2_NAME").Value = wsSource.Range("COURSE2_NAME").Value ' $D$107
    End If
    If .Range("COURSE3_HOURS").Value <> wsSource.Range("COURSE3_HOURS").Value Then
        .Range("COURSE3_HOURS").Value = wsSource.Range("COURSE3_HOURS").Value ' $E$108
    End If
    If .Range("COURSE3_NAME").Value <> wsSource.Range("COURSE3_NAME").Value Then
        .Range("COURSE3_NAME").Value = wsSource.Range("COURSE3_NAME").Value ' $D$108
    End If
    If .Range("COURSE4_HOURS").Value <> wsSource.Range("COURSE4_HOURS").Value Then
        .Range("COURSE4_HOURS").Value = wsSource.Range("COURSE4_HOURS").Value ' $E$109
    End If
    If .Range("COURSE4_NAME").Value <> wsSource.Range("COURSE4_NAME").Value Then
        .Range("COURSE4_NAME").Value = wsSource.Range("COURSE4_NAME").Value ' $D$109
    End If
    If .Range("COURSE5_HOURS").Value <> wsSource.Range("COURSE5_HOURS").Value Then
        .Range("COURSE5_HOURS").Value = wsSource.Range("COURSE5_HOURS").Value ' $E$110
    End If
    If .Range("COURSE5_NAME").Value <> wsSource.Range("COURSE5_NAME").Value Then
        .Range("COURSE5_NAME").Value = wsSource.Range("COURSE5_NAME").Value ' $D$110
    End If
    If .Range("CP_AI").Value <> wsSource.Range("CP_AI").Value Then
        .Range("CP_AI").Value = wsSource.Range("CP_AI").Value ' $E$14
    End If
    If .Range("CP_ANA_COMPLEX_QTY").Value <> wsSource.Range("CP_ANA_COMPLEX_QTY").Value Then
        .Range("CP_ANA_COMPLEX_QTY").Value = wsSource.Range("CP_ANA_COMPLEX_QTY").Value ' $E$21
    End If
    If .Range("CP_AO").Value <> wsSource.Range("CP_AO").Value Then
        .Range("CP_AO").Value = wsSource.Range("CP_AO").Value ' $E$15
    End If
    If .Range("CP_DI").Value <> wsSource.Range("CP_DI").Value Then
        .Range("CP_DI").Value = wsSource.Range("CP_DI").Value ' $E$16
    End If
    .Range("CP_DIGITAL_COMPLEX_QTY").Value = wsSource.Range("CP_DIGITAL_COMPLEX_QTY").Value ' $E$22
    .Range("CP_DIGITAL_CTRL_DI").Value = wsSource.Range("CP_DIGITAL_CTRL_DI").Value ' $I$15
    .Range("CP_DIGITAL_CTRL_DO").Value = wsSource.Range("CP_DIGITAL_CTRL_DO").Value ' $J$15
    If .Range("CP_DO").Value <> wsSource.Range("CP_DO").Value Then
        .Range("CP_DO").Value = wsSource.Range("CP_DO").Value ' $E$17
    End If
    .Range("CP_FIELDBUS_IO_QTY").Value = wsSource.Range("CP_FIELDBUS_IO_QTY").Value ' $E$23
    .Range("CP_FIELDBUS_IO_RATIO").Value = wsSource.Range("CP_FIELDBUS_IO_RATIO").Value ' $E$23
    
 'Formula: =CP_GRP_START_COMPLEX_REC_TUNED
    .Range("CP_GRP_START_COMPLEX_QTY").Value = wsSource.Range("CP_GRP_START_COMPLEX_QTY").Value ' $E$31
    .Range("CP_GRP_START_LOOP_QTY").Value = wsSource.Range("CP_GRP_START_LOOP_QTY").Value ' $E$30
    .Range("CP_SEQ_COMPLEX_QTY").Value = wsSource.Range("CP_SEQ_COMPLEX_QTY").Value ' $E$27
    .Range("CP_SEQ_LOOP_QTY").Value = wsSource.Range("CP_SEQ_LOOP_QTY").Value ' $E$26
    .Range("CUSTOMER_SPEC_REQ").Value = wsSource.Range("CUSTOMER_SPEC_REQ").Value ' $H$145
    .Range("CUSTOMER_SPECS_EXIST").Value = wsSource.Range("CUSTOMER_SPECS_EXIST").Value ' $H$145
    .Range("DATE_END").Value = wsSource.Range("DATE_END").Value ' $E$191
    .Range("DATE_START").Value = wsSource.Range("DATE_START").Value ' $E$190
    If .Range("DI_AI").Value <> wsSource.Range("DI_AI").Value Then
        .Range("DI_AI").Value = wsSource.Range("DI_AI").Value ' $E$38
    End If
    If .Range("DI_AO").Value <> wsSource.Range("DI_AO").Value Then
        .Range("DI_AO").Value = wsSource.Range("DI_AO").Value ' $E$39
    End If
    .Range("DI_COMPLEX_QTY").Value = wsSource.Range("DI_COMPLEX_QTY").Value ' $E$47
    If .Range("DI_DEVICES").Value <> wsSource.Range("DI_DEVICES").Value Then
        .Range("DI_DEVICES").Value = wsSource.Range("DI_DEVICES").Value ' $E$45
    End If
    If .Range("DI_DI").Value <> wsSource.Range("DI_DI").Value Then
        .Range("DI_DI").Value = wsSource.Range("DI_DI").Value ' $E$40
    End If
    .Range("DI_DIGITAL_CTRL_DI").Value = wsSource.Range("DI_DIGITAL_CTRL_DI").Value ' $I$39
    .Range("DI_DIGITAL_CTRL_DO").Value = wsSource.Range("DI_DIGITAL_CTRL_DO").Value ' $J$39
    If .Range("DI_DO").Value <> wsSource.Range("DI_DO").Value Then
        .Range("DI_DO").Value = wsSource.Range("DI_DO").Value ' $E$41
    End If
 'Formula: =DI_GRP_START_COMPLEX_REC_TUNED
    .Range("DI_GRP_START_COMPLEX_QTY").Value = wsSource.Range("DI_GRP_START_COMPLEX_QTY").Value ' $E$55
    .Range("DI_GRP_START_LOOP_QTY").Value = wsSource.Range("DI_GRP_START_LOOP_QTY").Value ' $E$54
    If .Range("DI_INTERFACES").Value <> wsSource.Range("DI_INTERFACES").Value Then
        .Range("DI_INTERFACES").Value = wsSource.Range("DI_INTERFACES").Value ' $E$46
    End If
    If .Range("DI_IOTYPE_STS").Value <> wsSource.Range("DI_IOTYPE_STS").Value Then
        .Range("DI_IOTYPE_STS").Value = wsSource.Range("DI_IOTYPE_STS").Value ' $H$36
    End If
    .Range("DI_SEQ_COMPLEX_QTY").Value = wsSource.Range("DI_SEQ_COMPLEX_QTY").Value ' $E$51
    .Range("DI_SEQ_LOOP_QTY").Value = wsSource.Range("DI_SEQ_LOOP_QTY").Value ' $E$50
    If .Range("DOC_BOM_REQ").Value <> wsSource.Range("DOC_BOM_REQ").Value Then
        .Range("DOC_BOM_REQ").Value = wsSource.Range("DOC_BOM_REQ").Value ' $H$135
    End If
    If .Range("DOC_CAB_ELEC_REQ").Value <> wsSource.Range("DOC_CAB_ELEC_REQ").Value Then
        .Range("DOC_CAB_ELEC_REQ").Value = wsSource.Range("DOC_CAB_ELEC_REQ").Value ' $H$140
    End If
    If .Range("DOC_CAB_MECH_REQ").Value <> wsSource.Range("DOC_CAB_MECH_REQ").Value Then
        .Range("DOC_CAB_MECH_REQ").Value = wsSource.Range("DOC_CAB_MECH_REQ").Value ' $H$139
    End If
    If .Range("DOC_LOOP_REQ").Value <> wsSource.Range("DOC_LOOP_REQ").Value Then
        .Range("DOC_LOOP_REQ").Value = wsSource.Range("DOC_LOOP_REQ").Value ' $H$142
    End If
    If .Range("DOC_PWR_GND_REQ").Value <> wsSource.Range("DOC_PWR_GND_REQ").Value Then
        .Range("DOC_PWR_GND_REQ").Value = wsSource.Range("DOC_PWR_GND_REQ").Value ' $H$138
    End If
    If .Range("DOC_PWR_HEAT_REQ").Value <> wsSource.Range("DOC_PWR_HEAT_REQ").Value Then
        .Range("DOC_PWR_HEAT_REQ").Value = wsSource.Range("DOC_PWR_HEAT_REQ").Value ' $H$141
    End If
    If .Range("DOC_QA_REQ").Value <> wsSource.Range("DOC_QA_REQ").Value Then
        .Range("DOC_QA_REQ").Value = wsSource.Range("DOC_QA_REQ").Value ' $H$143
    End If
    If .Range("DOC_SYS_ARCH_REQ").Value <> wsSource.Range("DOC_SYS_ARCH_REQ").Value Then
        .Range("DOC_SYS_ARCH_REQ").Value = wsSource.Range("DOC_SYS_ARCH_REQ").Value ' $H$136
    End If
    If .Range("DOC_SYS_IND_REQ").Value <> wsSource.Range("DOC_SYS_IND_REQ").Value Then
        .Range("DOC_SYS_IND_REQ").Value = wsSource.Range("DOC_SYS_IND_REQ").Value ' $H$137
    End If
    If .Range("DOC_TAGLIST_REQ").Value <> wsSource.Range("DOC_TAGLIST_REQ").Value Then
        .Range("DOC_TAGLIST_REQ").Value = wsSource.Range("DOC_TAGLIST_REQ").Value ' $H$144
    End If
    .Range("DURATION").Value = wsSource.Range("DURATION").Value ' $E$188
    If .Range("ESD_AI").Value <> wsSource.Range("ESD_AI").Value Then
        .Range("ESD_AI").Value = wsSource.Range("ESD_AI").Value ' $E$62
    End If
    If .Range("ESD_AI_DESIRED").Value <> wsSource.Range("ESD_AI_DESIRED").Value Then
        .Range("ESD_AI_DESIRED").Value = wsSource.Range("ESD_AI").Value ' $E$62
    End If
    If .Range("ESD_AO").Value <> wsSource.Range("ESD_AO").Value Then
        .Range("ESD_AO").Value = wsSource.Range("ESD_AO").Value ' $E$63
    End If
    If .Range("ESD_AO_DESIRED").Value <> wsSource.Range("ESD_AO_DESIRED").Value Then
        .Range("ESD_AO_DESIRED").Value = wsSource.Range("ESD_AO").Value ' $E$63
    End If
    .Range("ESD_CAB_IO_QTY").Value = wsSource.Range("ESD_CAB_IO_QTY").Value ' $H$178
    .Range("ESD_CAB_MARSH_QTY").Value = wsSource.Range("ESD_CAB_MARSH_QTY").Value ' $H$179
    .Range("ESD_CAB_PROC_QTY").Value = wsSource.Range("ESD_CAB_PROC_QTY").Value ' $H$177
    .Range("ESD_CHASSIS_QTY").Value = wsSource.Range("ESD_CHASSIS_QTY").Value ' $H$172
    .Range("ESD_COMM_QTY").Value = wsSource.Range("ESD_COMM_QTY").Value ' $H$174
    .Range("ESD_COMPLEX_QTY").Value = wsSource.Range("ESD_COMPLEX_QTY").Value ' $E$73
    If .Range("ESD_DI").Value <> wsSource.Range("ESD_DI").Value Then
        .Range("ESD_DI").Value = wsSource.Range("ESD_DI").Value ' $E$64
    End If
    If .Range("ESD_DI_DESIRED").Value <> wsSource.Range("ESD_DI_DESIRED").Value Then
        .Range("ESD_DI_DESIRED").Value = wsSource.Range("ESD_DI").Value ' $E$64
    End If
    If .Range("ESD_DO").Value <> wsSource.Range("ESD_DO").Value Then
        .Range("ESD_DO").Value = wsSource.Range("ESD_DO").Value ' $E$65
    End If
    If .Range("ESD_DO_DESIRED").Value <> wsSource.Range("ESD_DO_DESIRED").Value Then
        .Range("ESD_DO_DESIRED").Value = wsSource.Range("ESD_DO").Value ' $E$65
    End If
    .Range("ESD_GRP_START_COMPLEX_QTY").Value = wsSource.Range("ESD_GRP_START_COMPLEX_QTY").Value ' $E$78
    .Range("ESD_GRP_START_LOOP_QTY").Value = wsSource.Range("ESD_GRP_START_LOOP_QTY").Value ' $E$77
    If .Range("ESD_HMI_REQ").Value <> wsSource.Range("ESD_HMI_REQ").Value Then
        .Range("ESD_HMI_REQ").Value = wsSource.Range("ESD_HMI_REQ").Value ' $H$70
    End If
    .Range("ESD_IO_CARD_QTY").Value = wsSource.Range("ESD_IO_CARD_QTY").Value ' $H$173
    If .Range("ESD_MARSH_CAB_REQ").Value <> wsSource.Range("ESD_MARSH_CAB_REQ").Value Then
        .Range("ESD_MARSH_CAB_REQ").Value = wsSource.Range("ESD_MARSH_CAB_REQ").Value ' $H$71
    End If
    .Range("ESD_MISC_CAB_QTY").Value = wsSource.Range("ESD_MISC_CAB_QTY").Value ' $E$74
    If .Range("ESD_PROG_REQ").Value <> wsSource.Range("ESD_PROG_REQ").Value Then
        .Range("ESD_PROG_REQ").Value = wsSource.Range("ESD_PROG_REQ").Value ' $H$72
    End If
    
    
    ' TMC tab
    .Range("Aeroderivative_QTY").Value = wsSource.Range("Aeroderivative_QTY").Value ' $E$85
    .Range("AirBlower_REQ").Value = wsSource.Range("AirBlower_REQ").Value ' $H$110
    .Range("Autosynchronization_REQ").Value = wsSource.Range("Autosynchronization_REQ").Value ' $H$101
    .Range("B35A_QTY").Value = wsSource.Range("B35A_QTY").Value ' $E$105
    .Range("BN_REQ").Value = wsSource.Range("BN_REQ").Value ' $H$114
    .Range("BoilerFeedwaterPump_REQ").Value = wsSource.Range("BoilerFeedwaterPump_REQ").Value ' $H$108
    .Range("Compressor_QTY").Value = wsSource.Range("Compressor_QTY").Value ' $E$95
    .Range("DoubleExtraction_QTY").Value = wsSource.Range("DoubleExtraction_QTY").Value ' $E$88
    .Range("FanDrive_REQ").Value = wsSource.Range("FanDrive_REQ").Value ' $H$109
    .Range("GasTurbine_QTY").Value = wsSource.Range("GasTurbine_QTY").Value ' $E$82
    .Range("Generator_QTY").Value = wsSource.Range("Generator_QTY").Value ' $E$100
    .Range("LoadSharing_REQ").Value = wsSource.Range("LoadSharing_REQ").Value ' $H$99
    .Range("LoadSharingGen_REQ").Value = wsSource.Range("LoadSharingGen_REQ").Value ' $H$103
    .Range("MechanicalRetrofit_REQ").Value = wsSource.Range("MechanicalRetrofit_REQ").Value ' $H$113
    .Range("MotorDriven_QTY").Value = wsSource.Range("MotorDriven_QTY").Value ' $E$92
    .Range("MultiShaft_QTY").Value = wsSource.Range("MultiShaft_QTY").Value ' $E$84
    .Range("PowerSystemStabilizer_REQ").Value = wsSource.Range("PowerSystemStabilizer_REQ").Value ' $H$102
    .Range("PRV_Controls_REQ").Value = wsSource.Range("PRV_Controls_REQ").Value ' $H$90
    .Range("PRVValves_QTY").Value = wsSource.Range("PRVValves_QTY").Value ' $E$91
    .Range("RecycleValves_QTY").Value = wsSource.Range("RecycleValves_QTY").Value ' $E$98
    .Range("Reheat_QTY").Value = wsSource.Range("Reheat_QTY").Value ' $E$89
    .Range("S506A_QTY").Value = wsSource.Range("S506A_QTY").Value ' $E$106
    .Range("S720A_QTY").Value = wsSource.Range("S720A_QTY").Value ' $E$107
    .Range("SingleExtraction_QTY").Value = wsSource.Range("SingleExtraction_QTY").Value ' $E$87
    .Range("SingleShaft_QTY").Value = wsSource.Range("SingleShaft_QTY").Value ' $E$83
    .Range("SteamTurbine_QTY").Value = wsSource.Range("SteamTurbine_QTY").Value ' $E$86
    .Range("SurgeControl_REQ").Value = wsSource.Range("SurgeControl_REQ").Value ' $H$96
    .Range("TurboSentry_REQ").Value = wsSource.Range("TurboSentry_REQ").Value ' $H$115
    .Range("TypeCompressor").Value = wsSource.Range("TypeCompressor").Value ' $E$97
        
    'ESD_SYSTEM_REQ
    
    If val(strCoverSheetVersion) >= val("1.0") Then
        If wsSource.Range("ESD_SYSTEM_REQ").Value = 1 Then
            .Range("SystemType").Value = "ESD" ' $H$69
        ElseIf wsSource.Range("ESD_SYSTEM_REQ").Value = 2 Then
            .Range("SystemType").Value = "BMS" ' $H$69
        Else
            .Range("SystemType").Value = "TMC" ' $H$69
        End If
    Else
        If wsSource.Range("ESD_SYSTEM_REQ").Value = True Then
            .Range("SystemType").Value = "ESD" ' $H$69
        Else
            .Range("SystemType").Value = "BMS" ' $H$69
        End If
    End If
    .Range("ESD_SYSTEMS_QTY").Value = wsSource.Range("ESD_SYSTEMS_QTY").Value ' $H$171
    .Range("MEETING_CLOSE").Value = wsSource.Range("MEETING_CLOSE").Value ' $E$155
    .Range("MEETING_DESIGN").Value = wsSource.Range("MEETING_DESIGN").Value ' $E$152
    .Range("MEETING_KICKOFF").Value = wsSource.Range("MEETING_KICKOFF").Value ' $E$151
    .Range("MEETING_OTHER").Value = wsSource.Range("MEETING_OTHER").Value ' $E$154
    .Range("MEETING_PROGRESS").Value = wsSource.Range("MEETING_PROGRESS").Value ' $E$153
    If .Range("NO_OF_UNITS").Value <> wsSource.Range("NO_OF_UNITS").Value Then
        .Range("NO_OF_UNITS").Value = wsSource.Range("NO_OF_UNITS").Value ' $E$187
    End If
    .Range("RENTAL_COST").Value = wsSource.Range("RENTAL_COST").Value ' $E$129
    .Range("REP_CUSTOM").Value = wsSource.Range("REP_CUSTOM").Value ' $E$85
    .Range("REP_MASS_HEAT").Value = wsSource.Range("REP_MASS_HEAT").Value ' $E$86
    .Range("REP_STD").Value = wsSource.Range("REP_STD").Value ' $E$84
    .Range("REPORT_STD").Value = wsSource.Range("REPORT_STD").Value ' $E$84
    .Range("SITE_COMM_HOURS").Value = wsSource.Range("SITE_COMM_HOURS").Value ' $E$118
    .Range("SITE_PWRUP_HOURS").Value = wsSource.Range("SITE_PWRUP_HOURS").Value ' $E$117
    .Range("SITE_SAT_HOURS").Value = wsSource.Range("SITE_SAT_HOURS").Value ' $E$119
    .Range("SITE_SURVEY_HOURS").Value = wsSource.Range("SITE_SURVEY_HOURS").Value ' $E$116
    .Range("SW_SIMULATOR_REQ").Value = wsSource.Range("SW_SIMULATOR_REQ").Value ' $H$127
    .Range("SYS_CONTROLLERS_QTY").Value = wsSource.Range("SYS_CONTROLLERS_QTY").Value ' $E$172
    .Range("SYS_FBM_QTY").Value = wsSource.Range("SYS_FBM_QTY").Value ' $E$173
    .Range("SYS_FDSI_QTY").Value = wsSource.Range("SYS_FDSI_QTY").Value ' $E$174
    .Range("SYS_WORKSTATIONS_QTY").Value = wsSource.Range("SYS_WORKSTATIONS_QTY").Value ' $E$171
    .Range("TEST_CUSTOMER_FAT").Value = wsSource.Range("TEST_CUSTOMER_FAT").Value ' $E$128
    .Range("TEST_FAT_PCT").Value = wsSource.Range("TEST_FAT_PCT").Value ' $E$126
    .Range("TEST_PRE_FAT_PCT").Value = wsSource.Range("TEST_PRE_FAT_PCT").Value ' $E$125
    .Range("TL_AIRFARE").Value = wsSource.Range("TL_AIRFARE").Value ' $E$164
    .Range("TL_DAILY_ALLOW").Value = wsSource.Range("TL_DAILY_ALLOW").Value ' $E$165
    .Range("TL_DAYS_QTY").Value = wsSource.Range("TL_DAYS_QTY").Value ' $E$163
    .Range("TL_PERS_QTY").Value = wsSource.Range("TL_PERS_QTY").Value ' $E$162
    .Range("TL_TRIPS_REQ").Value = wsSource.Range("TL_TRIPS_REQ").Value ' $E$161
End With
 '138
End Sub


Private Sub ImportDurationBasedSheet(wsSource As Worksheet, wsTarget As Worksheet)
On Error Resume Next

With wsTarget

    
    '### HACK to get these cells from workbooks that had no named ranges.
    Select Case strCoverSheetVersion
    Case "1.0"
        '89
        If .Range("DURATION_LOC_SE_QTY").Value <> wsSource.Range("$C$89").Value Then
            .Range("DURATION_LOC_SE_QTY").Value = wsSource.Range("$C$89").Value ' $C$89
        End If
        If .Range("DURATION_LOC_SE_WEEK").Value <> wsSource.Range("$F$89").Value Then
            .Range("DURATION_LOC_SE_WEEK").Value = wsSource.Range("$F$89").Value ' $F$89
        End If
            If .Range("DURATION_LOC_SE_DUP").Value <> wsSource.Range("$H$89").Value Then
            .Range("DURATION_LOC_SE_DUP").Value = wsSource.Range("$H$89").Value ' $H$89
        End If
        
        '90
        If .Range("DURATION_LOC_SR_QTY").Value <> wsSource.Range("$C$90").Value Then
            .Range("DURATION_LOC_SR_QTY").Value = wsSource.Range("$C$90").Value ' $C$90
        End If
        If .Range("DURATION_LOC_SR_WEEK").Value <> wsSource.Range("$F$90").Value Then
            .Range("DURATION_LOC_SR_WEEK").Value = wsSource.Range("$F$90").Value ' $F$90
        End If
        If .Range("DURATION_LOC_SSE_DUP").Value <> wsSource.Range("$H$90").Value Then
            .Range("DURATION_LOC_SSE_DUP").Value = wsSource.Range("$H$90").Value ' $H$90
        End If
        
        '109
        If .Range("DURATION_REM_SE_QTY").Value <> wsSource.Range("$C$109").Value Then
            .Range("DURATION_REM_SE_QTY").Value = wsSource.Range("$C$109").Value ' $C$109
        End If
        If .Range("DURATION_REM_SE_WEEK").Value <> wsSource.Range("$F$109").Value Then
            .Range("DURATION_REM_SE_WEEK").Value = wsSource.Range("$F$109").Value ' $F$109
        End If
        If .Range("DURATION_REM_SE_DUP").Value <> wsSource.Range("$H$109").Value Then
            .Range("DURATION_REM_SE_DUP").Value = wsSource.Range("$H$109").Value ' $H$109
        End If
    
        '110
        If .Range("DURATION_REM_SR_QTY").Value <> wsSource.Range("$C$110").Value Then
            .Range("DURATION_REM_SR_QTY").Value = wsSource.Range("$C$110").Value ' $C$110
        End If
        If .Range("DURATION_REM_SR_WEEK").Value <> wsSource.Range("$F$110").Value Then
            .Range("DURATION_REM_SR_WEEK").Value = wsSource.Range("$F$110").Value ' $F$110
        End If
        If .Range("DURATION_REM_SR_DUP").Value <> wsSource.Range("$H$110").Value Then
            .Range("DURATION_REM_SR_DUP").Value = wsSource.Range("$H$110").Value ' $H$110
        End If
        
    Case Else
    
        '89
        If .Range("DURATION_LOC_SE_QTY").Value <> wsSource.Range("DURATION_LOC_SE_QTY").Value Then
            .Range("DURATION_LOC_SE_QTY").Value = wsSource.Range("DURATION_LOC_SE_QTY").Value ' $C$89
        End If
        If .Range("DURATION_LOC_SE_WEEK").Value <> wsSource.Range("DURATION_LOC_SE_WEEK").Value Then
            .Range("DURATION_LOC_SE_WEEK").Value = wsSource.Range("DURATION_LOC_SE_WEEK").Value ' $F$89
        End If
        If .Range("DURATION_LOC_SE_DUP").Value <> wsSource.Range("DURATION_LOC_SE_DUP").Value Then
            .Range("DURATION_LOC_SE_DUP").Value = wsSource.Range("DURATION_LOC_SE_DUP").Value ' $H$89
        End If
        If .Range("DURATION_LOC_SE_UTIL").Value <> wsSource.Range("DURATION_LOC_SE_UTIL").Value Then
            .Range("DURATION_LOC_SE_UTIL").Value = wsSource.Range("DURATION_LOC_SE_UTIL").Value ' $J$89
        End If
        
        '90
        If .Range("DURATION_LOC_SR_QTY").Value <> wsSource.Range("DURATION_LOC_SR_QTY").Value Then
            .Range("DURATION_LOC_SR_QTY").Value = wsSource.Range("DURATION_LOC_SR_QTY").Value ' $C$90
        End If
        If .Range("DURATION_LOC_SR_WEEK").Value <> wsSource.Range("DURATION_LOC_SR_WEEK").Value Then
            .Range("DURATION_LOC_SR_WEEK").Value = wsSource.Range("DURATION_LOC_SR_WEEK").Value ' $F$90
        End If
        If .Range("DURATION_LOC_SSE_DUP").Value <> wsSource.Range("DURATION_LOC_SSE_DUP").Value Then
            .Range("DURATION_LOC_SSE_DUP").Value = wsSource.Range("DURATION_LOC_SSE_DUP").Value ' $H$90
        End If
        If .Range("DURATION_LOC_SR_UTIL").Value <> wsSource.Range("DURATION_LOC_SR_UTIL").Value Then
            .Range("DURATION_LOC_SR_UTIL").Value = wsSource.Range("DURATION_LOC_SR_UTIL").Value ' $J$90
        End If
        
        '109
        If .Range("DURATION_REM_SE_QTY").Value <> wsSource.Range("DURATION_REM_SE_QTY").Value Then
            .Range("DURATION_REM_SE_QTY").Value = wsSource.Range("DURATION_REM_SE_QTY").Value ' $C$109
        End If
        If .Range("DURATION_REM_SE_WEEK").Value <> wsSource.Range("DURATION_REM_SE_WEEK").Value Then
            .Range("DURATION_REM_SE_WEEK").Value = wsSource.Range("DURATION_REM_SE_WEEK").Value ' $F$109
        End If
        If .Range("DURATION_SE_REM_COUNTRY").Value <> wsSource.Range("DURATION_SE_REM_COUNTRY").Value Then
            .Range("DURATION_SE_REM_COUNTRY").Value = wsSource.Range("DURATION_SE_REM_COUNTRY").Value ' $G$109
        End If
        If .Range("DURATION_REM_SE_DUP").Value <> wsSource.Range("DURATION_REM_SE_DUP").Value Then
            .Range("DURATION_REM_SE_DUP").Value = wsSource.Range("DURATION_REM_SE_DUP").Value ' $H$109
        End If
        If .Range("DURATION_REM_SE_UTIL").Value <> wsSource.Range("DURATION_REM_SE_UTIL").Value Then
            .Range("DURATION_REM_SE_UTIL").Value = wsSource.Range("DURATION_REM_SE_UTIL").Value ' $J$109
        End If
        
        '110
        If .Range("DURATION_REM_SR_QTY").Value <> wsSource.Range("DURATION_REM_SR_QTY").Value Then
            .Range("DURATION_REM_SR_QTY").Value = wsSource.Range("DURATION_REM_SR_QTY").Value ' $C$110
        End If
        If .Range("DURATION_REM_SR_WEEK").Value <> wsSource.Range("DURATION_REM_SR_WEEK").Value Then
            .Range("DURATION_REM_SR_WEEK").Value = wsSource.Range("DURATION_REM_SR_WEEK").Value ' $F$110
        End If
        If .Range("DURATION_SR_REM_COUNTRY").Value <> wsSource.Range("DURATION_SR_REM_COUNTRY").Value Then
            .Range("DURATION_SR_REM_COUNTRY").Value = wsSource.Range("DURATION_SR_REM_COUNTRY").Value ' $G$110
        End If
        If .Range("DURATION_REM_SR_DUP").Value <> wsSource.Range("DURATION_REM_SR_DUP").Value Then
            .Range("DURATION_REM_SR_DUP").Value = wsSource.Range("DURATION_REM_SR_DUP").Value ' $H$110
        End If
        If .Range("DURATION_REM_SR_UTIL").Value <> wsSource.Range("DURATION_REM_SR_UTIL").Value Then
            .Range("DURATION_REM_SR_UTIL").Value = wsSource.Range("DURATION_REM_SR_UTIL").Value ' $J$110
        End If
    
    
    End Select

'now finish by just using named ranges
    If .Range("DURATION_AC_REM_COUNTRY").Value <> wsSource.Range("DURATION_AC_REM_COUNTRY").Value Then
        .Range("DURATION_AC_REM_COUNTRY").Value = wsSource.Range("DURATION_AC_REM_COUNTRY").Value ' $G$105
    End If
    If .Range("DURATION_AE_REM_COUNTRY").Value <> wsSource.Range("DURATION_AE_REM_COUNTRY").Value Then
        .Range("DURATION_AE_REM_COUNTRY").Value = wsSource.Range("DURATION_AE_REM_COUNTRY").Value ' $G$107
    End If
 'Formula: =DURATION_BASED_APP_PCT_TUNED
    If .Range("DURATION_BASED_APP_PCT_QTY").Value <> wsSource.Range("DURATION_BASED_APP_PCT_QTY").Value Then
        .Range("DURATION_BASED_APP_PCT_QTY").Value = wsSource.Range("DURATION_BASED_APP_PCT_QTY").Value ' $E$46
    End If
 'Formula: =DURATION_BASED_CONTROLS_PCT_TUNED
    If .Range("DURATION_BASED_CONTROLS_PCT_QTY").Value <> wsSource.Range("DURATION_BASED_CONTROLS_PCT_QTY").Value Then
        .Range("DURATION_BASED_CONTROLS_PCT_QTY").Value = wsSource.Range("DURATION_BASED_CONTROLS_PCT_QTY").Value ' $E$37
    End If
 'Formula: =DURATION_BASED_COURSE_PCT_TUNED
    If .Range("DURATION_BASED_COURSE_PCT_QTY").Value <> wsSource.Range("DURATION_BASED_COURSE_PCT_QTY").Value Then
        .Range("DURATION_BASED_COURSE_PCT_QTY").Value = wsSource.Range("DURATION_BASED_COURSE_PCT_QTY").Value ' $E$55
    End If
 'Formula: =DURATION_BASED_DESIGN_PCT_TUNED
    If .Range("DURATION_BASED_DESIGN_PCT_QTY").Value <> wsSource.Range("DURATION_BASED_DESIGN_PCT_QTY").Value Then
        .Range("DURATION_BASED_DESIGN_PCT_QTY").Value = wsSource.Range("DURATION_BASED_DESIGN_PCT_QTY").Value ' $G$22
    End If
 'Formula: =DURATION_BASED_DI_PCT_TUNED
    If .Range("DURATION_BASED_DI_PCT_QTY").Value <> wsSource.Range("DURATION_BASED_DI_PCT_QTY").Value Then
        .Range("DURATION_BASED_DI_PCT_QTY").Value = wsSource.Range("DURATION_BASED_DI_PCT_QTY").Value ' $E$40
    End If
 'Formula: =DURATION_BASED_DOC_PCT_TUNED
    If .Range("DURATION_BASED_DOC_PCT_QTY").Value <> wsSource.Range("DURATION_BASED_DOC_PCT_QTY").Value Then
        .Range("DURATION_BASED_DOC_PCT_QTY").Value = wsSource.Range("DURATION_BASED_DOC_PCT_QTY").Value ' $E$52
    End If
 'Formula: =DURATION_BASED_HMI_PCT_TUNED
    If .Range("DURATION_BASED_HMI_PCT_QTY").Value <> wsSource.Range("DURATION_BASED_HMI_PCT_QTY").Value Then
        .Range("DURATION_BASED_HMI_PCT_QTY").Value = wsSource.Range("DURATION_BASED_HMI_PCT_QTY").Value ' $E$34
    End If
 'Formula: =DURATION_BASED_IMPLEMENT_PCT_TUNED
    If .Range("DURATION_BASED_IMPLEMENT_PCT_QTY").Value <> wsSource.Range("DURATION_BASED_IMPLEMENT_PCT_QTY").Value Then
        .Range("DURATION_BASED_IMPLEMENT_PCT_QTY").Value = wsSource.Range("DURATION_BASED_IMPLEMENT_PCT_QTY").Value ' $H$22
    End If
 'Formula: =DURATION_BASED_MEETINGS_PCT_TUNED
    If .Range("DURATION_BASED_MEETINGS_PCT_QTY").Value <> wsSource.Range("DURATION_BASED_MEETINGS_PCT_QTY").Value Then
        .Range("DURATION_BASED_MEETINGS_PCT_QTY").Value = wsSource.Range("DURATION_BASED_MEETINGS_PCT_QTY").Value ' $E$61
    End If
 'Formula: =DURATION_BASED_PM_PCT_TUNED
    If .Range("DURATION_BASED_PM_PCT_QTY").Value <> wsSource.Range("DURATION_BASED_PM_PCT_QTY").Value Then
        .Range("DURATION_BASED_PM_PCT_QTY").Value = wsSource.Range("DURATION_BASED_PM_PCT_QTY").Value ' $E$58
    End If
 'Formula: =DURATION_BASED_REP_PCT_TUNED
    If .Range("DURATION_BASED_REP_PCT_QTY").Value <> wsSource.Range("DURATION_BASED_REP_PCT_QTY").Value Then
        .Range("DURATION_BASED_REP_PCT_QTY").Value = wsSource.Range("DURATION_BASED_REP_PCT_QTY").Value ' $E$43
    End If
 'Formula: =DURATION_BASED_REVIEW_PCT_TUNED
    If .Range("DURATION_BASED_REVIEW_PCT_QTY").Value <> wsSource.Range("DURATION_BASED_REVIEW_PCT_QTY").Value Then
        .Range("DURATION_BASED_REVIEW_PCT_QTY").Value = wsSource.Range("DURATION_BASED_REVIEW_PCT_QTY").Value ' $F$22
    End If
 'Formula: =DURATION_BASED_SITE_PCT_TUNED
    If .Range("DURATION_BASED_SITE_PCT_QTY").Value <> wsSource.Range("DURATION_BASED_SITE_PCT_QTY").Value Then
        .Range("DURATION_BASED_SITE_PCT_QTY").Value = wsSource.Range("DURATION_BASED_SITE_PCT_QTY").Value ' $E$64
    End If
 'Formula: =DURATION_BASED_SITE_PCT_TUNED
    If .Range("DURATION_BASED_SITE_SITE_PCT_QTY").Value <> wsSource.Range("DURATION_BASED_SITE_SITE_PCT_QTY").Value Then
        .Range("DURATION_BASED_SITE_SITE_PCT_QTY").Value = wsSource.Range("DURATION_BASED_SITE_SITE_PCT_QTY").Value ' $E$64
    End If
 'Formula: =DURATION_BASED_SPEC_PCT_TUNED
    If .Range("DURATION_BASED_SPEC_PCT_QTY").Value <> wsSource.Range("DURATION_BASED_SPEC_PCT_QTY").Value Then
        .Range("DURATION_BASED_SPEC_PCT_QTY").Value = wsSource.Range("DURATION_BASED_SPEC_PCT_QTY").Value ' $E$28
    End If
 'Formula: =DURATION_BASED_SYS_ENG_PCT_TUNED
    If .Range("DURATION_BASED_SYS_ENG_PCT_QTY").Value <> wsSource.Range("DURATION_BASED_SYS_ENG_PCT_QTY").Value Then
        .Range("DURATION_BASED_SYS_ENG_PCT_QTY").Value = wsSource.Range("DURATION_BASED_SYS_ENG_PCT_QTY").Value ' $E$31
    End If
 'Formula: =DURATION_BASED_TEST_PCT_TUNED
    If .Range("DURATION_BASED_TEST_PCT_QTY").Value <> wsSource.Range("DURATION_BASED_TEST_PCT_QTY").Value Then
        .Range("DURATION_BASED_TEST_PCT_QTY").Value = wsSource.Range("DURATION_BASED_TEST_PCT_QTY").Value ' $E$49
    End If
 'Formula: =DURATION_BASED_TEST_PCT_TUNED
    If .Range("DURATION_BASED_TEST_TEST_PCT_QTY").Value <> wsSource.Range("DURATION_BASED_TEST_TEST_PCT_QTY").Value Then
        .Range("DURATION_BASED_TEST_TEST_PCT_QTY").Value = wsSource.Range("DURATION_BASED_TEST_TEST_PCT_QTY").Value ' $E$49
    End If
 'Formula: =DURATION_BASED_TEST_PCT_TUNED
    If .Range("DURATION_ESD_AC_LOC_HOURS").Value <> wsSource.Range("DURATION_ESD_AC_LOC_HOURS").Value Then
        .Range("DURATION_ESD_AC_LOC_HOURS").Value = wsSource.Range("DURATION_BASED_TEST_TEST_PCT_QTY").Value ' $E$49
    End If
 'Formula: =DURATION_BASED_TEST_PCT_TUNED
    If .Range("DURATION_ESD_AC_REM_HOURS").Value <> wsSource.Range("DURATION_ESD_AC_REM_HOURS").Value Then
        .Range("DURATION_ESD_AC_REM_HOURS").Value = wsSource.Range("DURATION_BASED_TEST_TEST_PCT_QTY").Value ' $E$49
    End If
 'Formula: =DURATION_BASED_TEST_PCT_TUNED
    If .Range("DURATION_ESD_AE_LOC_HOURS").Value <> wsSource.Range("DURATION_ESD_AE_LOC_HOURS").Value Then
        .Range("DURATION_ESD_AE_LOC_HOURS").Value = wsSource.Range("DURATION_BASED_TEST_TEST_PCT_QTY").Value ' $E$49
    End If
 'Formula: =DURATION_BASED_TEST_PCT_TUNED
    If .Range("DURATION_ESD_AE_REM_HOURS").Value <> wsSource.Range("DURATION_ESD_AE_REM_HOURS").Value Then
        .Range("DURATION_ESD_AE_REM_HOURS").Value = wsSource.Range("DURATION_BASED_TEST_TEST_PCT_QTY").Value ' $E$49
    End If
 'Formula: =DURATION_BASED_TEST_PCT_TUNED
    If .Range("DURATION_ESD_LE_LOC_HOURS").Value <> wsSource.Range("DURATION_ESD_LE_LOC_HOURS").Value Then
        .Range("DURATION_ESD_LE_LOC_HOURS").Value = wsSource.Range("DURATION_BASED_TEST_TEST_PCT_QTY").Value ' $E$49
    End If
 'Formula: =DURATION_BASED_TEST_PCT_TUNED
    If .Range("DURATION_ESD_LE_REM_HOURS").Value <> wsSource.Range("DURATION_ESD_LE_REM_HOURS").Value Then
        .Range("DURATION_ESD_LE_REM_HOURS").Value = wsSource.Range("DURATION_BASED_TEST_TEST_PCT_QTY").Value ' $E$49
    End If
 'Formula: =DURATION_BASED_TEST_PCT_TUNED
    If .Range("DURATION_ESD_T_LOC_HOURS").Value <> wsSource.Range("DURATION_ESD_T_LOC_HOURS").Value Then
        .Range("DURATION_ESD_T_LOC_HOURS").Value = wsSource.Range("DURATION_BASED_TEST_TEST_PCT_QTY").Value ' $E$49
    End If
 'Formula: =DURATION_BASED_TEST_PCT_TUNED
    If .Range("DURATION_ESD_T_REM_HOURS").Value <> wsSource.Range("DURATION_ESD_T_REM_HOURS").Value Then
        .Range("DURATION_ESD_T_REM_HOURS").Value = wsSource.Range("DURATION_BASED_TEST_TEST_PCT_QTY").Value ' $E$49
    End If
    If .Range("DURATION_FREE_1_CAT").Value <> wsSource.Range("DURATION_FREE_1_CAT").Value Then
        .Range("DURATION_FREE_1_CAT").Value = wsSource.Range("DURATION_FREE_1_CAT").Value ' $D$125
    End If
    If .Range("DURATION_FREE_1_COST").Value <> wsSource.Range("DURATION_FREE_1_COST").Value Then
        .Range("DURATION_FREE_1_COST").Value = wsSource.Range("DURATION_FREE_1_COST").Value ' $G$125
    End If
    If .Range("DURATION_FREE_1_DUP").Value <> wsSource.Range("DURATION_FREE_1_DUP").Value Then
        .Range("DURATION_FREE_1_DUP").Value = wsSource.Range("DURATION_FREE_1_DUP").Value ' $H$125
    End If
    If .Range("DURATION_FREE_1_QTY").Value <> wsSource.Range("DURATION_FREE_1_QTY").Value Then
        .Range("DURATION_FREE_1_QTY").Value = wsSource.Range("DURATION_FREE_1_QTY").Value ' $C$125
    End If
    If .Range("DURATION_FREE_1_UTIL").Value <> wsSource.Range("DURATION_FREE_1_UTIL").Value Then
        .Range("DURATION_FREE_1_UTIL").Value = wsSource.Range("DURATION_FREE_1_UTIL").Value ' $J$125
    End If
    If .Range("DURATION_FREE_1_WEEK").Value <> wsSource.Range("DURATION_FREE_1_WEEK").Value Then
        .Range("DURATION_FREE_1_WEEK").Value = wsSource.Range("DURATION_FREE_1_WEEK").Value ' $F$125
    End If
    If .Range("DURATION_FREE_10_CAT").Value <> wsSource.Range("DURATION_FREE_10_CAT").Value Then
        .Range("DURATION_FREE_10_CAT").Value = wsSource.Range("DURATION_FREE_10_CAT").Value ' $D$134
    End If
    If .Range("DURATION_FREE_10_COST").Value <> wsSource.Range("DURATION_FREE_10_COST").Value Then
        .Range("DURATION_FREE_10_COST").Value = wsSource.Range("DURATION_FREE_10_COST").Value ' $G$134
    End If
    If .Range("DURATION_FREE_10_QTY").Value <> wsSource.Range("DURATION_FREE_10_QTY").Value Then
        .Range("DURATION_FREE_10_QTY").Value = wsSource.Range("DURATION_FREE_10_QTY").Value ' $C$134
    End If
    If .Range("DURATION_FREE_10_UTIL").Value <> wsSource.Range("DURATION_FREE_10_UTIL").Value Then
        .Range("DURATION_FREE_10_UTIL").Value = wsSource.Range("DURATION_FREE_10_UTIL").Value ' $J$134
    End If
    If .Range("DURATION_FREE_10_WEEK").Value <> wsSource.Range("DURATION_FREE_10_WEEK").Value Then
        .Range("DURATION_FREE_10_WEEK").Value = wsSource.Range("DURATION_FREE_10_WEEK").Value ' $F$134
    End If
    If .Range("DURATION_FREE_11_CAT").Value <> wsSource.Range("DURATION_FREE_11_CAT").Value Then
        .Range("DURATION_FREE_11_CAT").Value = wsSource.Range("DURATION_FREE_11_CAT").Value ' $D$135
    End If
    If .Range("DURATION_FREE_11_COST").Value <> wsSource.Range("DURATION_FREE_11_COST").Value Then
        .Range("DURATION_FREE_11_COST").Value = wsSource.Range("DURATION_FREE_11_COST").Value ' $G$135
    End If
    If .Range("DURATION_FREE_11_DUP").Value <> wsSource.Range("DURATION_FREE_11_DUP").Value Then
        .Range("DURATION_FREE_11_DUP").Value = wsSource.Range("DURATION_FREE_11_DUP").Value ' $H$135
    End If
    If .Range("DURATION_FREE_11_QTY").Value <> wsSource.Range("DURATION_FREE_11_QTY").Value Then
        .Range("DURATION_FREE_11_QTY").Value = wsSource.Range("DURATION_FREE_11_QTY").Value ' $C$135
    End If
    If .Range("DURATION_FREE_11_UTIL").Value <> wsSource.Range("DURATION_FREE_11_UTIL").Value Then
        .Range("DURATION_FREE_11_UTIL").Value = wsSource.Range("DURATION_FREE_11_UTIL").Value ' $J$135
    End If
    If .Range("DURATION_FREE_11_WEEK").Value <> wsSource.Range("DURATION_FREE_11_WEEK").Value Then
        .Range("DURATION_FREE_11_WEEK").Value = wsSource.Range("DURATION_FREE_11_WEEK").Value ' $F$135
    End If
    If .Range("DURATION_FREE_2_CAT").Value <> wsSource.Range("DURATION_FREE_2_CAT").Value Then
        .Range("DURATION_FREE_2_CAT").Value = wsSource.Range("DURATION_FREE_2_CAT").Value ' $D$126
    End If
    If .Range("DURATION_FREE_2_COST").Value <> wsSource.Range("DURATION_FREE_2_COST").Value Then
        .Range("DURATION_FREE_2_COST").Value = wsSource.Range("DURATION_FREE_2_COST").Value ' $G$126
    End If
    If .Range("DURATION_FREE_2_DUP").Value <> wsSource.Range("DURATION_FREE_2_DUP").Value Then
        .Range("DURATION_FREE_2_DUP").Value = wsSource.Range("DURATION_FREE_2_DUP").Value ' $H$126
    End If
    If .Range("DURATION_FREE_2_QTY").Value <> wsSource.Range("DURATION_FREE_2_QTY").Value Then
        .Range("DURATION_FREE_2_QTY").Value = wsSource.Range("DURATION_FREE_2_QTY").Value ' $C$126
    End If
    If .Range("DURATION_FREE_2_UTIL").Value <> wsSource.Range("DURATION_FREE_2_UTIL").Value Then
        .Range("DURATION_FREE_2_UTIL").Value = wsSource.Range("DURATION_FREE_2_UTIL").Value ' $J$126
    End If
    If .Range("DURATION_FREE_2_WEEK").Value <> wsSource.Range("DURATION_FREE_2_WEEK").Value Then
        .Range("DURATION_FREE_2_WEEK").Value = wsSource.Range("DURATION_FREE_2_WEEK").Value ' $F$126
    End If
    If .Range("DURATION_FREE_3_CAT").Value <> wsSource.Range("DURATION_FREE_3_CAT").Value Then
        .Range("DURATION_FREE_3_CAT").Value = wsSource.Range("DURATION_FREE_3_CAT").Value ' $D$127
    End If
    If .Range("DURATION_FREE_3_COST").Value <> wsSource.Range("DURATION_FREE_3_COST").Value Then
        .Range("DURATION_FREE_3_COST").Value = wsSource.Range("DURATION_FREE_3_COST").Value ' $G$127
    End If
    If .Range("DURATION_FREE_3_DUP").Value <> wsSource.Range("DURATION_FREE_3_DUP").Value Then
        .Range("DURATION_FREE_3_DUP").Value = wsSource.Range("DURATION_FREE_3_DUP").Value ' $H$127
    End If
    If .Range("DURATION_FREE_3_QTY").Value <> wsSource.Range("DURATION_FREE_3_QTY").Value Then
        .Range("DURATION_FREE_3_QTY").Value = wsSource.Range("DURATION_FREE_3_QTY").Value ' $C$127
    End If
    If .Range("DURATION_FREE_3_UTIL").Value <> wsSource.Range("DURATION_FREE_3_UTIL").Value Then
        .Range("DURATION_FREE_3_UTIL").Value = wsSource.Range("DURATION_FREE_3_UTIL").Value ' $J$127
    End If
    If .Range("DURATION_FREE_3_WEEK").Value <> wsSource.Range("DURATION_FREE_3_WEEK").Value Then
        .Range("DURATION_FREE_3_WEEK").Value = wsSource.Range("DURATION_FREE_3_WEEK").Value ' $F$127
    End If
    If .Range("DURATION_FREE_4_CAT").Value <> wsSource.Range("DURATION_FREE_4_CAT").Value Then
        .Range("DURATION_FREE_4_CAT").Value = wsSource.Range("DURATION_FREE_4_CAT").Value ' $D$128
    End If
    If .Range("DURATION_FREE_4_COST").Value <> wsSource.Range("DURATION_FREE_4_COST").Value Then
        .Range("DURATION_FREE_4_COST").Value = wsSource.Range("DURATION_FREE_4_COST").Value ' $G$128
    End If
    If .Range("DURATION_FREE_4_DUP").Value <> wsSource.Range("DURATION_FREE_4_DUP").Value Then
        .Range("DURATION_FREE_4_DUP").Value = wsSource.Range("DURATION_FREE_4_DUP").Value ' $H$128
    End If
    If .Range("DURATION_FREE_4_QTY").Value <> wsSource.Range("DURATION_FREE_4_QTY").Value Then
        .Range("DURATION_FREE_4_QTY").Value = wsSource.Range("DURATION_FREE_4_QTY").Value ' $C$128
    End If
    If .Range("DURATION_FREE_4_UTIL").Value <> wsSource.Range("DURATION_FREE_4_UTIL").Value Then
        .Range("DURATION_FREE_4_UTIL").Value = wsSource.Range("DURATION_FREE_4_UTIL").Value ' $J$128
    End If
    If .Range("DURATION_FREE_4_WEEK").Value <> wsSource.Range("DURATION_FREE_4_WEEK").Value Then
        .Range("DURATION_FREE_4_WEEK").Value = wsSource.Range("DURATION_FREE_4_WEEK").Value ' $F$128
    End If
    If .Range("DURATION_FREE_5_CAT").Value <> wsSource.Range("DURATION_FREE_5_CAT").Value Then
        .Range("DURATION_FREE_5_CAT").Value = wsSource.Range("DURATION_FREE_5_CAT").Value ' $D$129
    End If
    If .Range("DURATION_FREE_5_COST").Value <> wsSource.Range("DURATION_FREE_5_COST").Value Then
        .Range("DURATION_FREE_5_COST").Value = wsSource.Range("DURATION_FREE_5_COST").Value ' $G$129
    End If
    If .Range("DURATION_FREE_5_DUP").Value <> wsSource.Range("DURATION_FREE_5_DUP").Value Then
        .Range("DURATION_FREE_5_DUP").Value = wsSource.Range("DURATION_FREE_5_DUP").Value ' $H$129
    End If
    If .Range("DURATION_FREE_5_QTY").Value <> wsSource.Range("DURATION_FREE_5_QTY").Value Then
        .Range("DURATION_FREE_5_QTY").Value = wsSource.Range("DURATION_FREE_5_QTY").Value ' $C$129
    End If
    If .Range("DURATION_FREE_5_UTIL").Value <> wsSource.Range("DURATION_FREE_5_UTIL").Value Then
        .Range("DURATION_FREE_5_UTIL").Value = wsSource.Range("DURATION_FREE_5_UTIL").Value ' $J$129
    End If
    If .Range("DURATION_FREE_5_WEEK").Value <> wsSource.Range("DURATION_FREE_5_WEEK").Value Then
        .Range("DURATION_FREE_5_WEEK").Value = wsSource.Range("DURATION_FREE_5_WEEK").Value ' $F$129
    End If
    If .Range("DURATION_FREE_6_CAT").Value <> wsSource.Range("DURATION_FREE_6_CAT").Value Then
        .Range("DURATION_FREE_6_CAT").Value = wsSource.Range("DURATION_FREE_6_CAT").Value ' $D$130
    End If
    If .Range("DURATION_FREE_6_COST").Value <> wsSource.Range("DURATION_FREE_6_COST").Value Then
        .Range("DURATION_FREE_6_COST").Value = wsSource.Range("DURATION_FREE_6_COST").Value ' $G$130
    End If
    If .Range("DURATION_FREE_6_QTY").Value <> wsSource.Range("DURATION_FREE_6_QTY").Value Then
        .Range("DURATION_FREE_6_QTY").Value = wsSource.Range("DURATION_FREE_6_QTY").Value ' $C$130
    End If
    If .Range("DURATION_FREE_6_UTIL").Value <> wsSource.Range("DURATION_FREE_6_UTIL").Value Then
        .Range("DURATION_FREE_6_UTIL").Value = wsSource.Range("DURATION_FREE_6_UTIL").Value ' $J$130
    End If
    If .Range("DURATION_FREE_6_WEEK").Value <> wsSource.Range("DURATION_FREE_6_WEEK").Value Then
        .Range("DURATION_FREE_6_WEEK").Value = wsSource.Range("DURATION_FREE_6_WEEK").Value ' $F$130
    End If
    If .Range("DURATION_FREE_7_CAT").Value <> wsSource.Range("DURATION_FREE_7_CAT").Value Then
        .Range("DURATION_FREE_7_CAT").Value = wsSource.Range("DURATION_FREE_7_CAT").Value ' $D$131
    End If
    If .Range("DURATION_FREE_7_COST").Value <> wsSource.Range("DURATION_FREE_7_COST").Value Then
        .Range("DURATION_FREE_7_COST").Value = wsSource.Range("DURATION_FREE_7_COST").Value ' $G$131
    End If
    If .Range("DURATION_FREE_7_DUP").Value <> wsSource.Range("DURATION_FREE_7_DUP").Value Then
        .Range("DURATION_FREE_7_DUP").Value = wsSource.Range("DURATION_FREE_7_DUP").Value ' $H$131
    End If
    If .Range("DURATION_FREE_7_QTY").Value <> wsSource.Range("DURATION_FREE_7_QTY").Value Then
        .Range("DURATION_FREE_7_QTY").Value = wsSource.Range("DURATION_FREE_7_QTY").Value ' $C$131
    End If
    If .Range("DURATION_FREE_7_UTIL").Value <> wsSource.Range("DURATION_FREE_7_UTIL").Value Then
        .Range("DURATION_FREE_7_UTIL").Value = wsSource.Range("DURATION_FREE_7_UTIL").Value ' $J$131
    End If
    If .Range("DURATION_FREE_7_WEEK").Value <> wsSource.Range("DURATION_FREE_7_WEEK").Value Then
        .Range("DURATION_FREE_7_WEEK").Value = wsSource.Range("DURATION_FREE_7_WEEK").Value ' $F$131
    End If
    If .Range("DURATION_FREE_8_CAT").Value <> wsSource.Range("DURATION_FREE_8_CAT").Value Then
        .Range("DURATION_FREE_8_CAT").Value = wsSource.Range("DURATION_FREE_8_CAT").Value ' $D$132
    End If
    If .Range("DURATION_FREE_8_COST").Value <> wsSource.Range("DURATION_FREE_8_COST").Value Then
        .Range("DURATION_FREE_8_COST").Value = wsSource.Range("DURATION_FREE_8_COST").Value ' $G$132
    End If
    If .Range("DURATION_FREE_8_QTY").Value <> wsSource.Range("DURATION_FREE_8_QTY").Value Then
        .Range("DURATION_FREE_8_QTY").Value = wsSource.Range("DURATION_FREE_8_QTY").Value ' $C$132
    End If
    If .Range("DURATION_FREE_8_UTIL").Value <> wsSource.Range("DURATION_FREE_8_UTIL").Value Then
        .Range("DURATION_FREE_8_UTIL").Value = wsSource.Range("DURATION_FREE_8_UTIL").Value ' $J$132
    End If
    If .Range("DURATION_FREE_8_WEEK").Value <> wsSource.Range("DURATION_FREE_8_WEEK").Value Then
        .Range("DURATION_FREE_8_WEEK").Value = wsSource.Range("DURATION_FREE_8_WEEK").Value ' $F$132
    End If
    If .Range("DURATION_FREE_9_CAT").Value <> wsSource.Range("DURATION_FREE_9_CAT").Value Then
        .Range("DURATION_FREE_9_CAT").Value = wsSource.Range("DURATION_FREE_9_CAT").Value ' $D$133
    End If
    If .Range("DURATION_FREE_9_COST").Value <> wsSource.Range("DURATION_FREE_9_COST").Value Then
        .Range("DURATION_FREE_9_COST").Value = wsSource.Range("DURATION_FREE_9_COST").Value ' $G$133
    End If
    If .Range("DURATION_FREE_9_DUP").Value <> wsSource.Range("DURATION_FREE_9_DUP").Value Then
        .Range("DURATION_FREE_9_DUP").Value = wsSource.Range("DURATION_FREE_9_DUP").Value ' $H$133
    End If
    If .Range("DURATION_FREE_9_QTY").Value <> wsSource.Range("DURATION_FREE_9_QTY").Value Then
        .Range("DURATION_FREE_9_QTY").Value = wsSource.Range("DURATION_FREE_9_QTY").Value ' $C$133
    End If
    If .Range("DURATION_FREE_9_UTIL").Value <> wsSource.Range("DURATION_FREE_9_UTIL").Value Then
        .Range("DURATION_FREE_9_UTIL").Value = wsSource.Range("DURATION_FREE_9_UTIL").Value ' $J$133
    End If
    If .Range("DURATION_FREE_9_WEEK").Value <> wsSource.Range("DURATION_FREE_9_WEEK").Value Then
        .Range("DURATION_FREE_9_WEEK").Value = wsSource.Range("DURATION_FREE_9_WEEK").Value ' $F$133
    End If
    If .Range("DURATION_LE_REM_COUNTRY").Value <> wsSource.Range("DURATION_LE_REM_COUNTRY").Value Then
        .Range("DURATION_LE_REM_COUNTRY").Value = wsSource.Range("DURATION_LE_REM_COUNTRY").Value ' $G$106
    End If
    If .Range("DURATION_LOC_AC_DUP").Value <> wsSource.Range("DURATION_LOC_AC_DUP").Value Then
        .Range("DURATION_LOC_AC_DUP").Value = wsSource.Range("DURATION_LOC_AC_DUP").Value ' $H$85
    End If
    If .Range("DURATION_LOC_AC_QTY").Value <> wsSource.Range("DURATION_LOC_AC_QTY").Value Then
        .Range("DURATION_LOC_AC_QTY").Value = wsSource.Range("DURATION_LOC_AC_QTY").Value ' $C$85
    End If
    If .Range("DURATION_LOC_AC_UTIL").Value <> wsSource.Range("DURATION_LOC_AC_UTIL").Value Then
        .Range("DURATION_LOC_AC_UTIL").Value = wsSource.Range("DURATION_LOC_AC_UTIL").Value ' $J$85
    End If
    If .Range("DURATION_LOC_AC_WEEK").Value <> wsSource.Range("DURATION_LOC_AC_WEEK").Value Then
        .Range("DURATION_LOC_AC_WEEK").Value = wsSource.Range("DURATION_LOC_AC_WEEK").Value ' $F$85
    End If
    If .Range("DURATION_LOC_AE_DUP").Value <> wsSource.Range("DURATION_LOC_AE_DUP").Value Then
        .Range("DURATION_LOC_AE_DUP").Value = wsSource.Range("DURATION_LOC_AE_DUP").Value ' $H$87
    End If
    If .Range("DURATION_LOC_AE_QTY").Value <> wsSource.Range("DURATION_LOC_AE_QTY").Value Then
        .Range("DURATION_LOC_AE_QTY").Value = wsSource.Range("DURATION_LOC_AE_QTY").Value ' $C$87
    End If
    If .Range("DURATION_LOC_AE_UTIL").Value <> wsSource.Range("DURATION_LOC_AE_UTIL").Value Then
        .Range("DURATION_LOC_AE_UTIL").Value = wsSource.Range("DURATION_LOC_AE_UTIL").Value ' $J$87
    End If
    If .Range("DURATION_LOC_AE_WEEK").Value <> wsSource.Range("DURATION_LOC_AE_WEEK").Value Then
        .Range("DURATION_LOC_AE_WEEK").Value = wsSource.Range("DURATION_LOC_AE_WEEK").Value ' $F$87
    End If
    If .Range("DURATION_LOC_BA_DUP").Value <> wsSource.Range("DURATION_LOC_BA_DUP").Value Then
        .Range("DURATION_LOC_BA_DUP").Value = wsSource.Range("DURATION_LOC_BA_DUP").Value ' $H$92
    End If
    If .Range("DURATION_LOC_BA_QTY").Value <> wsSource.Range("DURATION_LOC_BA_QTY").Value Then
        .Range("DURATION_LOC_BA_QTY").Value = wsSource.Range("DURATION_LOC_BA_QTY").Value ' $C$92
    End If
    If .Range("DURATION_LOC_BA_UTIL").Value <> wsSource.Range("DURATION_LOC_BA_UTIL").Value Then
        .Range("DURATION_LOC_BA_UTIL").Value = wsSource.Range("DURATION_LOC_BA_UTIL").Value ' $J$92
    End If
    If .Range("DURATION_LOC_BA_WEEK").Value <> wsSource.Range("DURATION_LOC_BA_WEEK").Value Then
        .Range("DURATION_LOC_BA_WEEK").Value = wsSource.Range("DURATION_LOC_BA_WEEK").Value ' $F$92
    End If
    If .Range("DURATION_LOC_LE_DUP").Value <> wsSource.Range("DURATION_LOC_LE_DUP").Value Then
        .Range("DURATION_LOC_LE_DUP").Value = wsSource.Range("DURATION_LOC_LE_DUP").Value ' $H$86
    End If
    If .Range("DURATION_LOC_LE_QTY").Value <> wsSource.Range("DURATION_LOC_LE_QTY").Value Then
        .Range("DURATION_LOC_LE_QTY").Value = wsSource.Range("DURATION_LOC_LE_QTY").Value ' $C$86
    End If
    If .Range("DURATION_LOC_LE_UTIL").Value <> wsSource.Range("DURATION_LOC_LE_UTIL").Value Then
        .Range("DURATION_LOC_LE_UTIL").Value = wsSource.Range("DURATION_LOC_LE_UTIL").Value ' $J$86
    End If
    If .Range("DURATION_LOC_LE_WEEK").Value <> wsSource.Range("DURATION_LOC_LE_WEEK").Value Then
        .Range("DURATION_LOC_LE_WEEK").Value = wsSource.Range("DURATION_LOC_LE_WEEK").Value ' $F$86
    End If
    If .Range("DURATION_LOC_PM_DUP").Value <> wsSource.Range("DURATION_LOC_PM_DUP").Value Then
        .Range("DURATION_LOC_PM_DUP").Value = wsSource.Range("DURATION_LOC_PM_DUP").Value ' $H$84
    End If
    If .Range("DURATION_LOC_PM_QTY").Value <> wsSource.Range("DURATION_LOC_PM_QTY").Value Then
        .Range("DURATION_LOC_PM_QTY").Value = wsSource.Range("DURATION_LOC_PM_QTY").Value ' $C$84
    End If
    If .Range("DURATION_LOC_PM_UTIL").Value <> wsSource.Range("DURATION_LOC_PM_UTIL").Value Then
        .Range("DURATION_LOC_PM_UTIL").Value = wsSource.Range("DURATION_LOC_PM_UTIL").Value ' $J$84
    End If
    If .Range("DURATION_LOC_PM_WEEK").Value <> wsSource.Range("DURATION_LOC_PM_WEEK").Value Then
        .Range("DURATION_LOC_PM_WEEK").Value = wsSource.Range("DURATION_LOC_PM_WEEK").Value ' $F$84
    End If
    If .Range("DURATION_LOC_SITE_DUP").Value <> wsSource.Range("DURATION_LOC_SITE_DUP").Value Then
        .Range("DURATION_LOC_SITE_DUP").Value = wsSource.Range("DURATION_LOC_SITE_DUP").Value ' $H$96
    End If
    If .Range("DURATION_LOC_SITE_QTY").Value <> wsSource.Range("DURATION_LOC_SITE_QTY").Value Then
        .Range("DURATION_LOC_SITE_QTY").Value = wsSource.Range("DURATION_LOC_SITE_QTY").Value ' $C$96
    End If
    If .Range("DURATION_LOC_SITE_UTIL").Value <> wsSource.Range("DURATION_LOC_SITE_UTIL").Value Then
        .Range("DURATION_LOC_SITE_UTIL").Value = wsSource.Range("DURATION_LOC_SITE_UTIL").Value ' $J$96
    End If
    If .Range("DURATION_LOC_SITE_WEEK").Value <> wsSource.Range("DURATION_LOC_SITE_WEEK").Value Then
        .Range("DURATION_LOC_SITE_WEEK").Value = wsSource.Range("DURATION_LOC_SITE_WEEK").Value ' $F$96
    End If
    If .Range("DURATION_LOC_STAGING_DUP").Value <> wsSource.Range("DURATION_LOC_STAGING_DUP").Value Then
        .Range("DURATION_LOC_STAGING_DUP").Value = wsSource.Range("DURATION_LOC_STAGING_DUP").Value ' $H$94
    End If
    If .Range("DURATION_LOC_STAGING_QTY").Value <> wsSource.Range("DURATION_LOC_STAGING_QTY").Value Then
        .Range("DURATION_LOC_STAGING_QTY").Value = wsSource.Range("DURATION_LOC_STAGING_QTY").Value ' $C$94
    End If
    If .Range("DURATION_LOC_STAGING_UTIL").Value <> wsSource.Range("DURATION_LOC_STAGING_UTIL").Value Then
        .Range("DURATION_LOC_STAGING_UTIL").Value = wsSource.Range("DURATION_LOC_STAGING_UTIL").Value ' $J$94
    End If
    If .Range("DURATION_LOC_STAGING_WEEK").Value <> wsSource.Range("DURATION_LOC_STAGING_WEEK").Value Then
        .Range("DURATION_LOC_STAGING_WEEK").Value = wsSource.Range("DURATION_LOC_STAGING_WEEK").Value ' $F$94
    End If
    If .Range("DURATION_LOC_T_DUP").Value <> wsSource.Range("DURATION_LOC_T_DUP").Value Then
        .Range("DURATION_LOC_T_DUP").Value = wsSource.Range("DURATION_LOC_T_DUP").Value ' $H$88
    End If
    If .Range("DURATION_LOC_T_QTY").Value <> wsSource.Range("DURATION_LOC_T_QTY").Value Then
        .Range("DURATION_LOC_T_QTY").Value = wsSource.Range("DURATION_LOC_T_QTY").Value ' $C$88
    End If
    If .Range("DURATION_LOC_T_UTIL").Value <> wsSource.Range("DURATION_LOC_T_UTIL").Value Then
        .Range("DURATION_LOC_T_UTIL").Value = wsSource.Range("DURATION_LOC_T_UTIL").Value ' $J$88
    End If
    If .Range("DURATION_LOC_T_WEEK").Value <> wsSource.Range("DURATION_LOC_T_WEEK").Value Then
        .Range("DURATION_LOC_T_WEEK").Value = wsSource.Range("DURATION_LOC_T_WEEK").Value ' $F$88
    End If
    If .Range("DURATION_PM_REM_COUNTRY").Value <> wsSource.Range("DURATION_PM_REM_COUNTRY").Value Then
        .Range("DURATION_PM_REM_COUNTRY").Value = wsSource.Range("DURATION_PM_REM_COUNTRY").Value ' $G$104
    End If
    If .Range("DURATION_REM_AC_DUP").Value <> wsSource.Range("DURATION_REM_AC_DUP").Value Then
        .Range("DURATION_REM_AC_DUP").Value = wsSource.Range("DURATION_REM_AC_DUP").Value ' $H$105
    End If
    If .Range("DURATION_REM_AC_QTY").Value <> wsSource.Range("DURATION_REM_AC_QTY").Value Then
        .Range("DURATION_REM_AC_QTY").Value = wsSource.Range("DURATION_REM_AC_QTY").Value ' $C$105
    End If
    If .Range("DURATION_REM_AC_UTIL").Value <> wsSource.Range("DURATION_REM_AC_UTIL").Value Then
        .Range("DURATION_REM_AC_UTIL").Value = wsSource.Range("DURATION_REM_AC_UTIL").Value ' $J$105
    End If
    If .Range("DURATION_REM_AC_WEEK").Value <> wsSource.Range("DURATION_REM_AC_WEEK").Value Then
        .Range("DURATION_REM_AC_WEEK").Value = wsSource.Range("DURATION_REM_AC_WEEK").Value ' $F$105
    End If
    If .Range("DURATION_REM_AE_DUP").Value <> wsSource.Range("DURATION_REM_AE_DUP").Value Then
        .Range("DURATION_REM_AE_DUP").Value = wsSource.Range("DURATION_REM_AE_DUP").Value ' $H$107
    End If
    If .Range("DURATION_REM_AE_QTY").Value <> wsSource.Range("DURATION_REM_AE_QTY").Value Then
        .Range("DURATION_REM_AE_QTY").Value = wsSource.Range("DURATION_REM_AE_QTY").Value ' $C$107
    End If
    If .Range("DURATION_REM_AE_UTIL").Value <> wsSource.Range("DURATION_REM_AE_UTIL").Value Then
        .Range("DURATION_REM_AE_UTIL").Value = wsSource.Range("DURATION_REM_AE_UTIL").Value ' $J$107
    End If
    If .Range("DURATION_REM_AE_WEEK").Value <> wsSource.Range("DURATION_REM_AE_WEEK").Value Then
        .Range("DURATION_REM_AE_WEEK").Value = wsSource.Range("DURATION_REM_AE_WEEK").Value ' $F$107
    End If
    If .Range("DURATION_REM_BA_DUP").Value <> wsSource.Range("DURATION_REM_BA_DUP").Value Then
        .Range("DURATION_REM_BA_DUP").Value = wsSource.Range("DURATION_REM_BA_DUP").Value ' $H$112
    End If
    If .Range("DURATION_REM_BA_QTY").Value <> wsSource.Range("DURATION_REM_BA_QTY").Value Then
        .Range("DURATION_REM_BA_QTY").Value = wsSource.Range("DURATION_REM_BA_QTY").Value ' $C$112
    End If
    If .Range("DURATION_REM_BA_UTIL").Value <> wsSource.Range("DURATION_REM_BA_UTIL").Value Then
        .Range("DURATION_REM_BA_UTIL").Value = wsSource.Range("DURATION_REM_BA_UTIL").Value ' $J$112
    End If
    If .Range("DURATION_REM_BA_WEEK").Value <> wsSource.Range("DURATION_REM_BA_WEEK").Value Then
        .Range("DURATION_REM_BA_WEEK").Value = wsSource.Range("DURATION_REM_BA_WEEK").Value ' $F$112
    End If
    If .Range("DURATION_REM_LE_DUP").Value <> wsSource.Range("DURATION_REM_LE_DUP").Value Then
        .Range("DURATION_REM_LE_DUP").Value = wsSource.Range("DURATION_REM_LE_DUP").Value ' $H$106
    End If
    If .Range("DURATION_REM_LE_QTY").Value <> wsSource.Range("DURATION_REM_LE_QTY").Value Then
        .Range("DURATION_REM_LE_QTY").Value = wsSource.Range("DURATION_REM_LE_QTY").Value ' $C$106
    End If
    If .Range("DURATION_REM_LE_UTIL").Value <> wsSource.Range("DURATION_REM_LE_UTIL").Value Then
        .Range("DURATION_REM_LE_UTIL").Value = wsSource.Range("DURATION_REM_LE_UTIL").Value ' $J$106
    End If
    If .Range("DURATION_REM_LE_WEEK").Value <> wsSource.Range("DURATION_REM_LE_WEEK").Value Then
        .Range("DURATION_REM_LE_WEEK").Value = wsSource.Range("DURATION_REM_LE_WEEK").Value ' $F$106
    End If
    If .Range("DURATION_REM_PM_DUP").Value <> wsSource.Range("DURATION_REM_PM_DUP").Value Then
        .Range("DURATION_REM_PM_DUP").Value = wsSource.Range("DURATION_REM_PM_DUP").Value ' $H$104
    End If
    If .Range("DURATION_REM_PM_QTY").Value <> wsSource.Range("DURATION_REM_PM_QTY").Value Then
        .Range("DURATION_REM_PM_QTY").Value = wsSource.Range("DURATION_REM_PM_QTY").Value ' $C$104
    End If
    If .Range("DURATION_REM_PM_UTIL").Value <> wsSource.Range("DURATION_REM_PM_UTIL").Value Then
        .Range("DURATION_REM_PM_UTIL").Value = wsSource.Range("DURATION_REM_PM_UTIL").Value ' $J$104
    End If
    If .Range("DURATION_REM_PM_WEEK").Value <> wsSource.Range("DURATION_REM_PM_WEEK").Value Then
        .Range("DURATION_REM_PM_WEEK").Value = wsSource.Range("DURATION_REM_PM_WEEK").Value ' $F$104
    End If
    If .Range("DURATION_REM_SITE_DUP").Value <> wsSource.Range("DURATION_REM_SITE_DUP").Value Then
        .Range("DURATION_REM_SITE_DUP").Value = wsSource.Range("DURATION_REM_SITE_DUP").Value ' $H$116
    End If
    If .Range("DURATION_REM_SITE_QTY").Value <> wsSource.Range("DURATION_REM_SITE_QTY").Value Then
        .Range("DURATION_REM_SITE_QTY").Value = wsSource.Range("DURATION_REM_SITE_QTY").Value ' $C$116
    End If
    If .Range("DURATION_REM_SITE_UTIL").Value <> wsSource.Range("DURATION_REM_SITE_UTIL").Value Then
        .Range("DURATION_REM_SITE_UTIL").Value = wsSource.Range("DURATION_REM_SITE_UTIL").Value ' $J$116
    End If
    If .Range("DURATION_REM_SITE_WEEK").Value <> wsSource.Range("DURATION_REM_SITE_WEEK").Value Then
        .Range("DURATION_REM_SITE_WEEK").Value = wsSource.Range("DURATION_REM_SITE_WEEK").Value ' $F$116
    End If
    If .Range("DURATION_REM_STAGING_DUP").Value <> wsSource.Range("DURATION_REM_STAGING_DUP").Value Then
        .Range("DURATION_REM_STAGING_DUP").Value = wsSource.Range("DURATION_REM_STAGING_DUP").Value ' $H$114
    End If
    If .Range("DURATION_REM_STAGING_QTY").Value <> wsSource.Range("DURATION_REM_STAGING_QTY").Value Then
        .Range("DURATION_REM_STAGING_QTY").Value = wsSource.Range("DURATION_REM_STAGING_QTY").Value ' $C$114
    End If
    If .Range("DURATION_REM_STAGING_UTIL").Value <> wsSource.Range("DURATION_REM_STAGING_UTIL").Value Then
        .Range("DURATION_REM_STAGING_UTIL").Value = wsSource.Range("DURATION_REM_STAGING_UTIL").Value ' $J$114
    End If
    If .Range("DURATION_REM_STAGING_WEEK").Value <> wsSource.Range("DURATION_REM_STAGING_WEEK").Value Then
        .Range("DURATION_REM_STAGING_WEEK").Value = wsSource.Range("DURATION_REM_STAGING_WEEK").Value ' $F$114
    End If
    If .Range("DURATION_REM_T_DUP").Value <> wsSource.Range("DURATION_REM_T_DUP").Value Then
        .Range("DURATION_REM_T_DUP").Value = wsSource.Range("DURATION_REM_T_DUP").Value ' $H$108
    End If
    If .Range("DURATION_REM_T_QTY").Value <> wsSource.Range("DURATION_REM_T_QTY").Value Then
        .Range("DURATION_REM_T_QTY").Value = wsSource.Range("DURATION_REM_T_QTY").Value ' $C$108
    End If
    If .Range("DURATION_REM_T_UTIL").Value <> wsSource.Range("DURATION_REM_T_UTIL").Value Then
        .Range("DURATION_REM_T_UTIL").Value = wsSource.Range("DURATION_REM_T_UTIL").Value ' $J$108
    End If
    If .Range("DURATION_REM_T_WEEK").Value <> wsSource.Range("DURATION_REM_T_WEEK").Value Then
        .Range("DURATION_REM_T_WEEK").Value = wsSource.Range("DURATION_REM_T_WEEK").Value ' $F$108
    End If
    If .Range("DURATION_T_REM_COUNTRY").Value <> wsSource.Range("DURATION_T_REM_COUNTRY").Value Then
        .Range("DURATION_T_REM_COUNTRY").Value = wsSource.Range("DURATION_T_REM_COUNTRY").Value ' $G$108
    End If
    If .Range("DURATION_TOTAL_BUSINESS_REM_COUNTRY").Value <> wsSource.Range("DURATION_TOTAL_BUSINESS_REM_COUNTRY").Value Then
        .Range("DURATION_TOTAL_BUSINESS_REM_COUNTRY").Value = wsSource.Range("DURATION_TOTAL_BUSINESS_REM_COUNTRY").Value ' $G$112
    End If
    If .Range("DURATION_TOTAL_SITE_REM_COUNTRY").Value <> wsSource.Range("DURATION_TOTAL_SITE_REM_COUNTRY").Value Then
        .Range("DURATION_TOTAL_SITE_REM_COUNTRY").Value = wsSource.Range("DURATION_TOTAL_SITE_REM_COUNTRY").Value ' $G$116
    End If
    If .Range("DURATION_TOTAL_STAGING_REM_COUNTRY").Value <> wsSource.Range("DURATION_TOTAL_STAGING_REM_COUNTRY").Value Then
        .Range("DURATION_TOTAL_STAGING_REM_COUNTRY").Value = wsSource.Range("DURATION_TOTAL_STAGING_REM_COUNTRY").Value ' $G$114
    End If
End With

End Sub

'strMaxRange is the cell reference for the lowest right cell where data could be
'so we do not have to loop through all the cells of the sheet
'currently this is the only sheet imported with this function since most of the cells do not have named ranges
'    gstrGECETaskBasedSheet , "A1:BR386"
Private Sub ImportSheet(strSheetName As String, strMaxRange As String, wsSource As Worksheet, wsTarget As Worksheet)
' this will loop through the cells

Dim c As Range
Dim strStatus As String

strStatus = frmComplete.Controls("txtOutput").Text

On Error Resume Next

With wsTarget
    For Each c In Worksheets(strSheetName).Range(strMaxRange).Cells
        'if its a green cell, has no formula and is not locked then import it
        'old sheets had incorrect color for column F ad N so look for it as well
        'c.Interior.ColorIndex = 8
        
        'If (c.Interior.ColorIndex = 4 Or c.Interior.ColorIndex = 44 Or c.Interior.ColorIndex = 8) And c.HasFormula = False And c.Locked = False Then
        If (c.Interior.ColorIndex = 4 Or c.Interior.ColorIndex = 44 Or c.Interior.ColorIndex = 8) And c.Locked = False Then
            
            'the current workbook range we want to set
            If .Range("" & c.Address & "").Value = wsSource.Range("" & c.Address & "").Value Then
                'do nothing
            Else
                
                .Range("" & c.Address & "").Value = wsSource.Range("" & c.Address & "").Value
                'only need to give feedback for task based since it can take a real LongPtr time if sheet is full
                'column 5 is tracking activity we know it will have a value
                If strSheetName = gstrGECETaskBasedSheet And c.Column = 5 Then
                    frmComplete.Controls("txtOutput").Text = strStatus & vbCrLf & "---- Importing Row: " & c.row & "  Task number: " & .Cells(c.row, 3).Value
                    DoEvents
                End If
            End If
            
        End If
    Next
End With

End Sub


''''

'utility function to script out import functions for a given sheet
'this one does not check if the cells are equal
Private Sub callScriptSheet()
    ScriptSheet gstrGECEPriceMakeUpSheet
'    ScriptSheet gstrGECEDataEntrySheet
'    ScriptSheet gstrGECEAssumptionsProposalSheet
'
'    ScriptSheet gstrGECEApplicationBasedSheet
'    ScriptSheet gstrGECETaskBasedSheet
'    ScriptSheet gstrGECEDurationBasedSheet
End Sub

Private Sub ScriptSheet(strSheetName As String)
Dim wb As Workbook
Dim ns As Names
Dim N As Name
Dim wss As Worksheets
Dim ws As Worksheet

Set wb = ActiveWorkbook
Set ns = wb.Names
Dim strTemp As String
Dim strTemp1 As String
Dim strTemp2 As String
Dim strStart As String
Dim strEnd As String

On Error Resume Next

Dim fso As New FileSystemObject
Dim s As Scripting.TextStream

Set s = fso.CreateTextFile("D:\GPO\GECE\LatestVersion\1.0\" & strSheetName & "_rangenames.txt", True, False)
s.WriteLine "'GECE Version = " & GECEXLSVERSION & " " & Now()
s.WriteBlankLines 1



strStart = "With wsTarget" & vbCrLf
strEnd = "End With" & vbCrLf
Dim intCount As Integer

For Each N In ns
    If InStr(1, N.Value, strSheetName) Then

            If ActiveWorkbook.Worksheets(strSheetName).Range(N.Name).HasFormula = False And ActiveWorkbook.Worksheets(strSheetName).Range(N.Name).Locked = False Then
            'If ActiveWorkbook.Worksheets(strSheetName).Range(n.Name).HasFormula = False Then
            intCount = intCount + 1
                'the current workbook range we want to set
                strTemp1 = ".Range(""" & N.Name & """).Value = "
                strTemp2 = "wsSource.Range(""" & N.Name & """).Value" & " ' " & ActiveWorkbook.Worksheets(strSheetName).Range(N.Name).Address
                strTemp = strTemp & vbTab & strTemp1 & strTemp2 & vbCrLf
            End If
        'End If
    End If
Next


'Debug.Print strStart & strTemp & strEnd & " '" & intCount
s.Write strStart & strTemp & strEnd & " '" & intCount
s.WriteBlankLines 1
'frmImport.Controls("txtOutput").Text = strStart & strTemp & strEnd & " '" & intCount
s.Close

Set s = Nothing
Set fso = Nothing
Set N = Nothing
Set ns = Nothing
Set ws = Nothing

End Sub

'this one will script with a check to see if the cells ar equal and then dont bother importing the cell
'this makes the import function faster sice we are not wasing time importing the same value
Private Sub callScriptSheetIfThen()
'    ScriptSheetIfThen gstrGECEPriceMakeUpSheet
'    ScriptSheetIfThen gstrGECEDataEntrySheet
'    ScriptSheetIfThen gstrGECEAssumptionsProposalSheet
  ScriptSheetIfThen gstrGECEApplicationBasedSheet
'    ScriptSheetIfThen gstrGECETaskBasedSheet
'    ScriptSheetIfThen gstrGECEDurationBasedSheet
End Sub
Private Sub ScriptSheetIfThen(strSheetName As String)
Dim wb As Workbook
Dim ns As Names
Dim N As Name
Dim wss As Worksheets
Dim ws As Worksheet

Set wb = ActiveWorkbook
Set ns = wb.Names
Dim strTemp As String
Dim strTemp1 As String
Dim strTemp2 As String
Dim strTempIfThen As String
Dim strTempEndIf As String
Dim strStart As String
Dim strEnd As String

On Error Resume Next

Dim fso As New FileSystemObject
Dim s As Scripting.TextStream

Set s = fso.CreateTextFile("D:\GPO\GECE\LatestVersion\1.0\" & strSheetName & "_rangenames.txt", True, False)
s.WriteLine "'GECE Version = " & GECEXLSVERSION & " " & Now()
s.WriteBlankLines 1



strStart = "With wsTarget" & vbCrLf
strEnd = "End With" & vbCrLf
Dim intCount As Integer

For Each N In ns
    If InStr(1, N.Value, strSheetName) Then

            If ActiveWorkbook.Worksheets(strSheetName).Range(N.Name).HasFormula = False And ActiveWorkbook.Worksheets(strSheetName).Range(N.Name).Locked = False Then
            'If ActiveWorkbook.Worksheets(strSheetName).Range(n.Name).HasFormula = False Then
            intCount = intCount + 1
                'the current workbook range we want to set
                strTempIfThen = vbTab & "If .Range(""" & N.Name & """).Value <> wsSource.Range(""" & N.Name & """).Value Then" & vbCrLf
                strTemp1 = ".Range(""" & N.Name & """).Value = "
                strTemp2 = "wsSource.Range(""" & N.Name & """).Value" & " ' " & ActiveWorkbook.Worksheets(strSheetName).Range(N.Name).Address
                strTempEndIf = vbTab & "End If" & vbCrLf
                strTemp = strTemp & strTempIfThen & vbTab & vbTab & strTemp1 & strTemp2 & vbCrLf & strTempEndIf
                'wrap an if statement
                
                
                
                
                'strTemp = strTempIfThen & strTemp & strTempEndIf
                
            End If
        'End If
    End If
Next


'Debug.Print strStart & strTemp & strEnd & " '" & intCount
s.Write strStart & strTemp & strEnd & " '" & intCount
s.WriteBlankLines 1
'frmImport.Controls("txtOutput").Text = strStart & strTemp & strEnd & " '" & intCount
s.Close

Set s = Nothing
Set fso = Nothing
Set N = Nothing
Set ns = Nothing
Set ws = Nothing

End Sub

Private Sub callScriptSheetIfThenNoFormula()
'ScriptSheetIfThenNoFormula gstrGECEPriceMakeUpSheet
    'ScriptSheetIfThenNoFormula gstrGECEDataEntrySheet
    'ScriptSheetIfThenNoFormula gstrGECEAssumptionsProposalSheet
  'ScriptSheetIfThenNoFormula gstrGECEApplicationBasedSheet
'    ScriptSheetIfThenNoFormula gstrGECETaskBasedSheet
    ScriptSheetIfThenNoFormula gstrGECEDurationBasedSheet
End Sub
Private Sub ScriptSheetIfThenNoFormula(strSheetName As String)
Dim wb As Workbook
Dim ns As Names
Dim N As Name
Dim wss As Worksheets
Dim ws As Worksheet

Set wb = ActiveWorkbook
Set ns = wb.Names
Dim strTemp As String
Dim strTemp1 As String
Dim strTemp2 As String
Dim strTempIfThen As String
Dim strTempEndIf As String
Dim strStart As String
Dim strEnd As String

On Error Resume Next

Dim fso As New FileSystemObject
Dim s As Scripting.TextStream

Set s = fso.CreateTextFile("D:\project\GECE\template\1.0\" & strSheetName & "_NoFormula.txt", True, False)
s.WriteLine "'GECE Version = " & GECEXLSVERSION & " " & Now()
s.WriteBlankLines 1



Dim strFormula As String

strStart = "With wsTarget" & vbCrLf
strEnd = "End With" & vbCrLf
Dim intCount As Integer

For Each N In ns
    If InStr(1, N.Value, strSheetName) Then


            If (ActiveWorkbook.Worksheets(strSheetName).Range(N.Name).Interior.ColorIndex = 4 Or ActiveWorkbook.Worksheets(strSheetName).Range(N.Name).Interior.ColorIndex = 44) And ActiveWorkbook.Worksheets(strSheetName).Range(N.Name).Locked = False Then
            'If ActiveWorkbook.Worksheets(strSheetName).Range(n.Name).HasFormula = False Then
            intCount = intCount + 1
                
                If ActiveWorkbook.Worksheets(strSheetName).Range(N.Name).HasFormula = True Then
                    strFormula = " 'Formula: " & ActiveWorkbook.Worksheets(strSheetName).Range(N.Name).Formula
                Else
                    strFormula = ""
                End If
                'add <strFormula & vbCrLf &> to begining of block
                If strFormula = "" Then
                    strTempIfThen = vbTab & "If .Range(""" & N.Name & """).Value <> wsSource.Range(""" & N.Name & """).Value Then" & vbCrLf
                Else
                    strTempIfThen = strFormula & vbCrLf & vbTab & "If .Range(""" & N.Name & """).Value <> wsSource.Range(""" & N.Name & """).Value Then" & vbCrLf
                End If
                'the current workbook range we want to set
                
                strTemp1 = ".Range(""" & N.Name & """).Value = "
                strTemp2 = "wsSource.Range(""" & N.Name & """).Value" & " ' " & ActiveWorkbook.Worksheets(strSheetName).Range(N.Name).Address
                strTempEndIf = vbTab & "End If" & vbCrLf
                strTemp = strTemp & strTempIfThen & vbTab & vbTab & strTemp1 & strTemp2 & vbCrLf & strTempEndIf
                'wrap an if statement
                
                
                
                
                'strTemp = strTempIfThen & strTemp & strTempEndIf
                
            End If
        'End If
    End If
Next


'Debug.Print strStart & strTemp & strEnd & " '" & intCount
s.Write strStart & strTemp & strEnd & " '" & intCount
s.WriteBlankLines 1
'frmImport.Controls("txtOutput").Text = strStart & strTemp & strEnd & " '" & intCount
s.Close

Set s = Nothing
Set fso = Nothing
Set N = Nothing
Set ns = Nothing
Set ws = Nothing

End Sub



'scripts out with cell range name not a defined named range
Private Sub callScriptSheet1()
    'ScriptSheet1 gstrGECEPriceMakeUpSheet
    'ScriptSheet1 gstrGECEDataEntrySheet, "A1:N193"

    'ScriptSheet1 gstrGECEAssumptionsProposalSheet, "A1:Q94"
   
    'ScriptSheet1 gstrGECEApplicationBasedSheet, "A1:V374"
    'ScriptSheet1 gstrGECETaskBasedSheet, "A1:BH437"
    ScriptSheet1 gstrGECEDurationBasedSheet, "A1:W138"
End Sub


Private Sub ScriptSheet1(strSheetName As String, strMaxRange As String)
' this will loop through the cells
'These are the rules for scripting data entry cells for the import routine.
'Green cell are user editable (Dark and Light Green) 4, 44
'Green cell should not contain formulas
'Green cells should be unlocked
'Green cell should have a Named Range assigned to them
'All other cells should be locked
'All other cells should have non Green cell color
Dim wb As Workbook
Dim ns As Names
Dim N As Name
Dim wss As Worksheets
Dim ws As Worksheet


Set wb = ActiveWorkbook
Set ns = wb.Names

Dim strTemp As String
Dim strTemp1 As String
Dim strTemp2 As String
Dim strStart As String
Dim strEnd As String

On Error Resume Next

Dim fso As New FileSystemObject
Dim s As Scripting.TextStream

Set s = fso.CreateTextFile("D:\project\GECE\template\1.0\" & strSheetName & "1.txt", True, False)
s.WriteLine "'GECE Version = " & GECEXLSVERSION & " " & Now()
s.WriteBlankLines 1



strStart = "With wsTarget" & vbCrLf
strEnd = "End With" & vbCrLf
Dim intCount As Integer

Dim c As Range


    For Each c In Worksheets(strSheetName).Range(strMaxRange).Cells
        If (c.Interior.ColorIndex = 4 Or c.Interior.ColorIndex = 44) And c.Locked = False Then
            'Debug.Print c.Address & vbTab & c.Row & vbTab & c.Column & vbTab & " color: " & c.Interior.ColorIndex
            'the current workbook range we want to set
            intCount = intCount + 1
            ' use thes two lines for cell address
            strTemp1 = ".Range(""" & c.Address & """).Value = "
            strTemp2 = "wsSource.Range(""" & c.Address & """).Value" & " ' GREEN " & c.Address '& vbTab & c.Name.Name
            
'            ' use thes two lines for named ranges
'            strTemp1 = ".Range(""" & c.Name.Name & """).Value = "
'            strTemp2 = "wsSource.Range(""" & c.Name.Name & """).Value" & " ' GREEN " & c.Address & vbTab & c.Name.Name
            
            strTemp = strTemp & vbTab & strTemp1 & strTemp2 & vbCrLf
        End If
'        If c.HasFormula = False And c.Locked = False Then
'            'Debug.Print c.Address & vbTab & c.Row & vbTab & c.Column & vbTab & " color: " & c.Interior.ColorIndex
'            'the current workbook range we want to set
'            strTemp1 = ".Range(""" & c.Address & """).Value = "
'            strTemp2 = "wsSource.Range(""" & c.Address & """).Value" & " ' UNLOCKED "
'            strTemp = strTemp & vbTab & strTemp1 & strTemp2 & vbCrLf
'        End If
        
        
        
    Next

'.Interior.ColorIndex
'4   light green
'15  dark grey
'34  light blue
'44  dark green
'46  orange
'47  light grey

'For Each n In ns
'    If InStr(1, n.Value, strSheetName) Then
'
'            If ActiveWorkbook.Worksheets(strSheetName).Range(n.Name).HasFormula = False And ActiveWorkbook.Worksheets(strSheetName).Range(n.Name).Locked = False Then
'            intCount = intCount + 1
'                'the current workbook range we want to set
'                strTemp1 = ".Range(""" & n.Name & """).Value = "
'                strTemp2 = "wsSource.Range(""" & n.Name & """).Value" & " ' " & ActiveWorkbook.Worksheets(strSheetName).Range(n.Name).Address
'                strTemp = strTemp & vbTab & strTemp1 & strTemp2 & vbCrLf
'            End If
'        'End If
'    End If
'Next


'Debug.Print strStart & strTemp & strEnd & " '" & intCount
s.Write strStart & strTemp & strEnd & " '" & intCount
s.WriteBlankLines 1
'frmImport.Controls("txtOutput").Text = strStart & strTemp & strEnd & " '" & intCount
s.Close

Set s = Nothing
Set fso = Nothing
Set N = Nothing
Set ns = Nothing
Set ws = Nothing

End Sub

Private Sub SafeSet(ws As Worksheet, ByVal nm As String, ByVal v As Variant)
    On Error Resume Next: ws.Range(nm).Value2 = v: On Error GoTo 0
End Sub
Private Sub SafeSetBool(ws As Worksheet, ByVal nm As String, ByVal v As Boolean)
    On Error Resume Next: ws.Range(nm).Value2 = v: On Error GoTo 0
End Sub
Private Sub SafeClearText(ws As Worksheet, ByVal nm As String)
    On Error Resume Next: ws.Range(nm).Value2 = vbNullString: On Error GoTo 0
End Sub

Private Sub ZeroOverridePairs(ws As Worksheet, ByVal prefix As String, ByVal a As Long, ByVal b As Long)
    Dim i As Long
    For i = a To b
        SafeClearText ws, prefix & i & "_OVD_JUST"
        SafeSet ws, prefix & i & "_OVD_QTY", 0
    Next i
End Sub
Private Sub ZeroOverrideStatuses(ws As Worksheet, ByVal prefix As String, ByVal a As Long, ByVal b As Long, _
                                 Optional headName As String = vbNullString)
    Dim i As Long
    If Len(headName) > 0 Then SafeSetBool ws, headName, False
    For i = a To b
        SafeSetBool ws, prefix & i & "_OVD_STS", False
    Next i
End Sub

'=== entry point ============================================================
Public Sub ResetWorkbookFields()
    On Error GoTo Fail

    Dim scr As Boolean, ev As Boolean, calc As XlCalculation
    scr = Application.ScreenUpdating: ev = Application.EnableEvents: calc = Application.Calculation
    Application.ScreenUpdating = False: Application.EnableEvents = False: Application.Calculation = xlCalculationManual
    Application.Cursor = xlWait

    Dim wb As Workbook: Set wb = ThisWorkbook

    'if frmComplete is loaded, show progress
    On Error Resume Next
    frmComplete.txtOutput.Text = "Resetting fields..." & vbCrLf
    On Error GoTo 0

    gbolImporting = True
    ResetDataEntrySheet wb.Worksheets(gstrGECEDataEntrySheet)
    ResetAssumptionsProposalSheet wb.Worksheets(gstrGECEAssumptionsProposalSheet)
    ResetApplicationBasedSheet wb.Worksheets(gstrGECEApplicationBasedSheet)
    ResetPriceMakeupSheet wb.Worksheets(gstrGECEPriceMakeUpSheet)

    Application.CalculateFullRebuild
    DoEvents

    On Error Resume Next
    frmComplete.txtOutput.Text = frmComplete.txtOutput.Text & "Reset completed."
    On Error GoTo 0

    gbolImporting = False
    GoTo CleanUp

Fail:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical, "GECE Reset"
CleanUp:
    Application.Calculation = calc
    Application.EnableEvents = ev
    Application.ScreenUpdating = scr
    Application.Cursor = xlDefault
    If Err.Number = 0 Then MsgBox "GECE Tool has been reset to factory defaults.", vbInformation, "GECE Reset"
End Sub

'=== sheet resets ===========================================================
Private Sub ResetApplicationBasedSheet(ws As Worksheet)
    Dim prefixes As Variant, p As Variant, i As Long
    prefixes = Array("APP_TASK", "COURSE_TASK", "CP_TASK", "DI_TASK", "DOC_TASK", "ESD_TASK", _
                     "HMI_TASK", "MEETING_TASK", "PM_TASK", "REP_TASK", "SITE_TASK", _
                     "SPEC_TASK", "SYSENG_TASK", "TEST_TASK", "CLEAN_TASK", "FF_TASK")
    For Each p In prefixes: ZeroOverridePairs ws, CStr(p), 1, 10: Next p

    If Not gbolImporting Then
        For Each p In prefixes
            For i = 1 To 10
                On Error Resume Next
                With ws.Range(CStr(p) & i & "_REM_COUNTRY")
                    .Value2 = .Value2
                End With
                On Error GoTo 0
            Next i
        Next p
    End If
End Sub

Private Sub ResetPriceMakeupSheet(ws As Worksheet)
    ZeroOverrideStatuses ws, "APP_TASK", 1, 10, "APP_OVD_STS"
    ZeroOverrideStatuses ws, "CLEAN_TASK", 1, 10, "CLEAN_OVD_STS"
    ZeroOverrideStatuses ws, "COURSE_TASK", 1, 10, "COURSE_OVD_STS"
    ZeroOverrideStatuses ws, "CP_TASK", 1, 10, "CP_OVD_STS"
    ZeroOverrideStatuses ws, "DI_TASK", 1, 10, "DI_OVD_STS"
    ZeroOverrideStatuses ws, "DOC_TASK", 1, 10, "DOC_OVD_STS"
    ZeroOverrideStatuses ws, "ESD_TASK", 1, 10, "ESD_OVD_STS"
    ZeroOverrideStatuses ws, "FF_TASK", 1, 10
    ZeroOverrideStatuses ws, "HMI_TASK", 1, 10, "HMI_OVD_STS"
    ZeroOverrideStatuses ws, "MEETING_TASK", 1, 10, "MEETING_OVD_STS"
    ZeroOverrideStatuses ws, "PM_TASK", 1, 10, "PM_OVD_STS"
    ZeroOverrideStatuses ws, "REP_TASK", 1, 10, "REP_OVD_STS"
    ZeroOverrideStatuses ws, "SITE_TASK", 1, 10, "SITE_OVD_STS"
    ZeroOverrideStatuses ws, "SPEC_TASK", 1, 10, "SPEC_OVD_STS"
    ZeroOverrideStatuses ws, "SYSENG_TASK", 1, 10, "SYS_ENG_OVD_STS"
    ZeroOverrideStatuses ws, "TEST_TASK", 1, 10, "TEST_OVD_STS"
    ZeroOverrideStatuses ws, "TL_TASK", 1, 10, "TL_OVD_STS"
End Sub

Private Sub ResetAssumptionsProposalSheet(ws As Worksheet)
    SafeClearText ws, "CUSTOMER_NAME"
    SafeClearText ws, "GSP_ID"
    SafeClearText ws, "PROJECT_MANAGER"
    SafeClearText ws, "PROJECT_NAME"
    SafeClearText ws, "PROPOSAL_NUMBER"
    SafeClearText ws, "PROPOSAL_REVISION"
    SafeClearText ws, "PROPSAL_DATE"
    SafeClearText ws, "PROPSAL_ENGINEER"

    SafeSet ws, "TOOLKIT_FACTOR", 0
    Dim i As Long
    For i = 2 To 22: SafeSetBool ws, "TOOLKIT_" & i & "_REQ", False: Next i

    SafeSet ws, "LOCAL_COUNTRY", "UK"
    SafeSet ws, "WPA", "Green Field"
    SafeSet ws, "WPA_TYPE", "Base model - Export"
End Sub

Private Sub ResetDataEntrySheet(ws As Worksheet)
    'counts
    SafeSet ws, "CP_AI", 0: SafeSet ws, "CP_AO", 0: SafeSet ws, "CP_DI", 0: SafeSet ws, "CP_DO", 0
    SafeSet ws, "DI_AI", 0: SafeSet ws, "DI_AO", 0: SafeSet ws, "DI_DI", 0: SafeSet ws, "DI_DO", 0
    SafeSet ws, "DI_DEVICES", 0: SafeSet ws, "DI_INTERFACES", 0: SafeSetBool ws, "DI_IOTYPE_STS", True
    SafeSet ws, "ESD_AI", 0: SafeSet ws, "ESD_AO", 0: SafeSet ws, "ESD_DI", 0: SafeSet ws, "ESD_DO", 0

    'percents (numeric 0)
    SafeSet ws, "CP_ANA_COMPLEX_QTY", 0
    SafeSet ws, "CP_DIGITAL_COMPLEX_QTY", 0
    SafeSet ws, "CP_FIELDBUS_IO_QTY", 0
    SafeSet ws, "CP_GRP_START_COMPLEX_QTY", 0
    SafeSet ws, "CP_GRP_START_LOOP_QTY", 0
    SafeSet ws, "CP_SEQ_LOOP_QTY", 0
    SafeSet ws, "DI_COMPLEX_QTY", 0
    SafeSet ws, "DI_GRP_START_COMPLEX_QTY", 0
    SafeSet ws, "DI_GRP_START_LOOP_QTY", 0
    SafeSet ws, "DI_SEQ_COMPLEX_QTY", 0
    SafeSet ws, "DI_SEQ_LOOP_QTY", 0
    SafeSet ws, "ESD_COMPLEX_QTY", 0
    SafeSet ws, "ESD_GRP_START_COMPLEX_QTY", 0
    SafeSet ws, "ESD_GRP_START_LOOP_QTY", 0
    SafeSet ws, "ESD_MISC_CAB_QTY", 0
    SafeSet ws, "TEST_FAT_PCT", 0
    SafeSet ws, "TEST_PRE_FAT_PCT", 0

    'flags
    SafeSetBool ws, "APP_BUS_STS", False
    SafeSetBool ws, "CUSTOMER_SPEC_REQ", False
    SafeSetBool ws, "CUSTOMER_SPECS_EXIST", False
    SafeSetBool ws, "DOC_BOM_REQ", False
    SafeSetBool ws, "DOC_CAB_ELEC_REQ", False
    SafeSetBool ws, "DOC_CAB_MECH_REQ", False
    SafeSetBool ws, "DOC_LOOP_REQ", False
    SafeSetBool ws, "DOC_PWR_GND_REQ", False
    SafeSetBool ws, "DOC_PWR_HEAT_REQ", False
    SafeSetBool ws, "DOC_QA_REQ", False
    SafeSetBool ws, "DOC_SYS_ARCH_REQ", False
    SafeSetBool ws, "DOC_SYS_IND_REQ", False
    SafeSetBool ws, "DOC_TAGLIST_REQ", False
    SafeSetBool ws, "ESD_HMI_REQ", False
    SafeSetBool ws, "ESD_MARSH_CAB_REQ", False
    SafeSetBool ws, "ESD_PROG_REQ", False
    SafeSetBool ws, "HIGH_RISK_SITE", False
    SafeSetBool ws, "LoadSharing_REQ", False
    SafeSetBool ws, "LoadSharingGen_REQ", False
    SafeSetBool ws, "PowerSystemStabilizer_REQ", False
    SafeSetBool ws, "PRV_Controls_REQ", False
    SafeSetBool ws, "SurgeControl_REQ", False
    SafeSetBool ws, "SW_SIMULATOR_REQ", False

    'clear texts
    Dim clears As Variant, c As Variant
    clears = Array("APP_1_HOURS", "APP_2_HOURS", "APP_3_HOURS", "APP_4_HOURS", "APP_5_HOURS", "APP_6_HOURS", "APP_7_HOURS", "APP_8_HOURS", _
                   "CAB_CONSOLES_QTY", "CAB_IO_QTY", "CAB_MARSH_QTY", "CAB_PROC_QTY", _
                   "DATE_START", "DATE_END", "DURATION", _
                   "ESD_CAB_IO_QTY", "ESD_CAB_MARSH_QTY", "ESD_CHASSIS_QTY", "ESD_COMM_QTY", "ESD_IO_CARD_QTY", "ESD_SYSTEMS_QTY", _
                   "SYS_CONTROLLERS_QTY", "SYS_FBM_QTY", "SYS_FDSI_QTY", "SYS_WORKSTATIONS_QTY", _
                   "TIME_KICKOFF_QTY", "TIME_PREPARE_HW_QTY", "TIME_RECEIVE_HW_QTY", "TIME_TRANSPORT_HW_QTY", _
                   "TL_PERS_QTY_SITE", "TL_TRIPS_REQ_SITE")
    For Each c In clears: SafeClearText ws, CStr(c): Next c

    'zeros
    SafeSet ws, "Aeroderivative_QTY", 0
    SafeSet ws, "Compressor_QTY", 0
    SafeSet ws, "COST_BUYOUT", 0
    SafeSet ws, "COST_IA", 0
    SafeSet ws, "COURSE_1_HOURS", 0
    SafeSet ws, "COURSE_2_HOURS", 0
    SafeSet ws, "COURSE_3_HOURS", 0
    SafeSet ws, "COURSE_4_HOURS", 0
    SafeSet ws, "COURSE_5_HOURS", 0
    SafeSet ws, "DoubleExtraction_QTY", 0
    SafeSet ws, "GasTurbine_QTY", 0
    SafeSet ws, "Generator_QTY", 0
    SafeSet ws, "MEETING_CLOSE", 0
    SafeSet ws, "MEETING_DESIGN", 0
    SafeSet ws, "MEETING_KICKOFF", 0
    SafeSet ws, "MEETING_OTHER", 0
    SafeSet ws, "MEETING_PROGRESS", 0
    SafeSet ws, "MotorDriven_QTY", 0
    SafeSet ws, "MultiShaft_QTY", 0
    SafeSet ws, "NO_OF_UNITS", 1
    SafeSet ws, "PRVValves_QTY", 0
    SafeSet ws, "RecycleValves_QTY", 0
    SafeSet ws, "Reheat_QTY", 0
    SafeSet ws, "RENTAL_COST", 0
    SafeSet ws, "REP_CUSTOM", 0
    SafeSet ws, "REP_MASS_HEAT", 0
    SafeSet ws, "REP_STD", 0
    SafeSet ws, "SingleExtraction_QTY", 0
    SafeSet ws, "SingleShaft_QTY", 0
    SafeSet ws, "SITE_COMM_HOURS", 0
    SafeSet ws, "SITE_PWRUP_HOURS", 0
    SafeSet ws, "SITE_SAT_HOURS", 0
    SafeSet ws, "SITE_SURVEY_HOURS", 0
    SafeSet ws, "SteamTurbine_QTY", 0
    SafeSet ws, "TEST_CUSTOMER_FAT", 0
    SafeSet ws, "TL_AIRFARE", 0
    SafeSet ws, "TL_AIRFARE_FAT", 0
    SafeSet ws, "TL_AIRFARE_SITE", 0
    SafeSet ws, "TL_DAILY_ALLOW", 0
    SafeSet ws, "TL_DAILY_ALLOW_FAT", 0
    SafeSet ws, "TL_DAILY_ALLOW_SITE", 0
    SafeSet ws, "TL_DAYS_QTY", 0
    SafeSet ws, "TL_DAYS_QTY_FAT", 0
    SafeSet ws, "TL_DAYS_QTY_SITE", 0
    SafeSet ws, "TL_PERS_QTY", 0
    SafeSet ws, "TL_PERS_QTY_FAT", 0
    SafeSet ws, "TL_TRIPS_REQ", 0
    SafeSet ws, "TL_TRIPS_REQ_FAT", 0

    'defaults
    SafeSet ws, "SystemType", "ESD"
    SafeSet ws, "TypeCompressor", "NA"
End Sub
