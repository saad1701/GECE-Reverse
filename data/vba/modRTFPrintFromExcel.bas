Attribute VB_Name = "modRTFPrintFromExcel"

Option Explicit
Private Const MODULE_NAME = "modRTFPrintFromExcel"
Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

Public aryValues(12) As String
Public arrFile(1) As String
Public Reply As Long
Public mstrPath As String
Public gblnBypassMode As Boolean
    'Constants
Public Const REGISTRY_APPNAME As String = "Software\GECE\GECE"
Public Const REG_OPTION_NON_VOLATILE = 0
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_DWORD = 4                      ' 32-bit number
Public Const HKEY_CURRENT_USER = &H80000001
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const SYNCHRONIZE = &H100000
Public Const KEY_NOTIFY = &H10
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const ERROR_NONE = 0

' Registry functions declarations

#If VBA7 Then
    ' --- Core registry functions for 64-bit Office ---
    Public Declare PtrSafe Function RegOpenKeyExA Lib "advapi32.dll" ( _
        ByVal hKey As LongPtr, ByVal lpSubKey As String, ByVal ulOptions As Long, _
        ByVal samDesired As Long, phkResult As LongPtr) As Long

''    Public Declare PtrSafe Function RegQueryValueExA Lib "advapi32.dll" ( _
        ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal lpReserved As Long, _
        lpType As Long, lpData As Any, lpcbData As Long) As Long

    Public Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" ( _
        ByVal hKey As LongPtr) As Long

    ' --- Aliased variants for specific data types ---
''    Public Declare PtrSafe Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
        ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal lpReserved As LongPtr, _
        lpType As LongPtr, ByVal lpData As LongPtr, lpcbData As LongPtr) As Long

''    Public Declare PtrSafe Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
        ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal lpReserved As LongPtr, _
        lpType As LongPtr, ByVal lpData As String, lpcbData As LongPtr) As Long

''    Public Declare PtrSafe Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
        ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal lpReserved As LongPtr, _
        lpType As LongPtr, lpData As LongPtr, lpcbData As LongPtr) As Long

    Public Declare PtrSafe Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" ( _
        ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal Reserved As LongPtr, _
        ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
#Else
    ' --- Core registry functions for 32-bit Office ---
    Public Declare Function RegOpenKeyExA Lib "advapi32.dll" ( _
        ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
        ByVal samDesired As Long, phkResult As Long) As Long

''    Public Declare Function RegQueryValueExA Lib "advapi32.dll" ( _
        ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
        lpType As Long, lpData As Any, lpcbData As Long) As Long

    Public Declare Function RegCloseKey Lib "advapi32.dll" ( _
        ByVal hKey As Long) As Long
#End If
Public Const RTF_ALIGNLEFT = 1      ' vbdefault left aligned
Public Const RTF_ALIGNRIGHT = 2     ' right aligned
Public Const RTF_ALIGNCENTER = 3    ' center aligned

' Clipboard Manager Functions
Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hWnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr

Private Declare PtrSafe Function SHGetSpecialFolderLocation Lib "Shell32" ( _
    ByVal hWnd As LongPtr, ByVal nFolder As Long, ppidl As LongPtr) As Long

Private Declare PtrSafe Function SHGetPathFromIDList Lib "Shell32" ( _
    ByVal Pidl As LongPtr, ByVal pszPath As String) As Long


'RTF Errors
'problems with code that is calling the class - non severe
Private Const ERR_FILEEXISTS = vbObjectError + 100
Private Const ERR_BOOKMARKNOTFOUND = vbObjectError + 101
Private Const ERR_TABLEALREADYEXISTS = vbObjectError + 102
Private Const ERR_TABLEINDEXNOTFOUND = vbObjectError + 103
Private Const ERR_BOOKMARKINDEXNOTFOUND = vbObjectError + 104
Private Const ERR_NOTABLES = vbObjectError + 105
Private Const ERR_TABLENOTFOUND = vbObjectError + 106
Private Const ERR_COPYTABLEFAILED = vbObjectError + 107
Private Const ERR_INVALIDCOLUMNINDEX = vbObjectError + 108

'problems internal to the class or template - severe
Private Const ERR_DISJOINTEDBOOKMARKS = vbObjectError + 200
Private Const ERR_SETCOLWIDTHFAILED = vbObjectError + 201
Private Const ERR_FORMATADDRESSFAILED = vbObjectError + 202


Public Const CAT_Discount = "discount"
Public Const ENCRYPTION_KEY = "BUTTHEAD"
Public Const BIZARRO = "BIZARRO"

'********************************************************************************
'*      Function Name: Encrypt                                                  *
'*
'*          Parameters:                                                         *
'*                  rstrSecret      -   The string to encrypt                   *
'*                  rstrPassword    -   An Encrypting Password                  *
'*            Comments:                                                         *
'*  This function will encrypt the passed string using the passed password in   *
'*  an XOR algorithm.                                                           *
'*  To decrypt simply call this function again with the same password.          *
'********************************************************************************
Public Function Encrypt(ByRef rstrText$, ByRef rstrPassword$) As String

On Error GoTo ErrorHandler

Dim intLength%, i%, intChar%
Dim strEncrypt$

intLength = Len(rstrPassword$)
strEncrypt = rstrText

For i = 1 To Len(strEncrypt)
    intChar = Asc(mid$(rstrPassword, (i Mod intLength) - intLength * ((i Mod intLength) = 0), 1))
    Mid$(strEncrypt, i, 1) = Chr$(Asc(mid$(rstrText, i, 1)) Xor intChar)
Next

Encrypt = strEncrypt

Exit Function
ErrorHandler:
    WriteLogFile Err, "modNG_VB.Encrypt:" & Error, False
    Encrypt = ""
    Exit Function
End Function

'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Procedure Name: DownloadDiscount
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DownloadDiscount()

On Error GoTo ErrorHandler
Dim strFile$
Dim i                       As Long
Dim lngRow                  As Long
Dim strTemp                 As String
Dim strFileName             As String
Dim strPath                 As String
Dim invRope                 As Object

Dim rsServer                As Object
Dim rsDetail                As Object

Dim MyExcel As Excel.Application
Dim MyWKB As Excel.Workbook
Dim MySheet As Worksheet
'
    strPath = ""
''1   Call GetKeyValue(HKEY_LOCAL_MACHINE, "Software\GECE\GECE\Install", "GECEPath", strPath, "1")
    strFileName = FixPath(strPath) & "Template\GECECurrencyfromiastore.XLS"
2  Kill (strFileName)

'Create ROPE objects
3   Set invRope = CreateObject("ROPE_20050101.Services")
    Call invRope.Authorization("", Encrypt(" $1:2!0", ENCRYPTION_KEY), Encrypt("rddex", ENCRYPTION_KEY), "0915", "PRD")
    
'refCurrency
    'Call Status("Loading Currencies...", 60, "Red")
    Set rsServer = invRope.GetData(CAT_Discount, "refCurrency_get", "@Status_Cd", "ACT")
    
    lngRow = 1
    Set MyExcel = New Excel.Application
    With MyExcel
        .Visible = False
        Set MyWKB = MyExcel.Workbooks.Add()
        With MyWKB
        Set MySheet = .Worksheets(1)
            With MySheet
              .Name = "Downloaded Currencies"
                  Do Until rsServer.EOF
                       Set rsDetail = invRope.GetData(CAT_Discount, "refCurrencyConvert_get", "@Source_Cd", "0915", "@Client_Cd", "PRD", "@Foreign_Currency_Cd", rsServer.fields("Currency_Cd"))
                       Do Until rsDetail.EOF
                           .Cells(lngRow, 1) = rsDetail.fields("Foreign_Currency_Cd")
                           .Cells(lngRow, 2) = rsDetail.fields("Currency_Cd")
                           .Cells(lngRow, 3) = rsDetail.fields("Conversion_Factor")
                           rsDetail.movenext
                           lngRow = lngRow + 1
                       Loop
                       i = i + 1
                       rsServer.movenext
                   Loop
             End With
        End With
        .Workbooks(1).SaveAs (strFileName)
    End With
CleanUp:
    On Error Resume Next
    Set MyExcel = Nothing
    Set MySheet = Nothing
    Set MyWKB = Nothing
    rsDetail.Clear
    Set rsDetail = Nothing
    rsServer.Clear
    Set rsServer = Nothing
    Set invRope = Nothing
    
Exit Sub
ErrorHandler:
    If Erl = 1 Then
        strFileName = "c:\GECECurrencyfromiastore.XLS"
        Resume 2
    End If
    If Erl = 2 Then Resume 3
    WriteLogFile Err, MODULE_NAME & ".DownloadDiscount: " & Error, strFileName, True
    Resume CleanUp
    Resume 0
End Sub



'********************************************************************************
'''*      Function Name: GetKeyValue                                              *
'*                                       *
'********************************************************************************
''Public Function GetKeyValue(ByVal PredefinedKey As Variant, ByVal KeyName As Variant, ByVal KeyValueName As Variant, ByRef KeyValueData As Variant, Optional ByRef rvntDefault As Variant, Optional ByRef rblnFixPath As Boolean) As Long

'Dim hKey        As LongPtr
'Dim cch         As LongPtr
'Dim lrc         As LongPtr
'Dim lngKeyType  As LongPtr
'Dim lValue      As LongPtr
'Dim strValue    As String

'KeyValueData = ""
'Reply = RegOpenKeyExA(PredefinedKey, KeyName, 0&, KEY_QUERY_VALUE, hKey) 'KEY_ALL_ACCESS
'If Not IsMissing(rvntDefault) Then KeyValueData = rvntDefault

'If Reply = 0 Then
'Determine the size and type of data to be read
    'Reply = RegQueryValueExNULL(hKey, KeyValueName, 0&, lngKeyType, 0&, cch)
    'If Reply <> ERROR_NONE Then
        'GoTo CleanUp
    'End If

'Determine the Type of data
    'KeyValueData = Empty
    'Select Case lngKeyType
        'Case REG_SZ:
            'strValue = String(CInt(cch), Chr(0))
            'lrc = RegQueryValueExString(hKey, KeyValueName, 0&, lngKeyType, strValue, cch)
            'If lrc = ERROR_NONE Then
                'KeyValueData = Left$(strValue, CLng(cch) - 1)
            'End If
            'If rblnFixPath Then KeyValueData = FixPath(KeyValueData)
        'Case REG_DWORD:
            'lrc = RegQueryValueExLong(hKey, KeyValueName, 0&, lngKeyType, lValue, cch)
            'If lrc = ERROR_NONE Then
                'KeyValueData = lValue
            'End If
    'End Select
'End If

'CleanUp:
    'On Error Resume Next
'Close the Key
    'Call RegCloseKey(hKey)
''    GetKeyValue = Reply

'End Function


Public Function FixPath(ByVal rstrAnyPath As String, Optional ByVal NoSlash As Boolean) As String
On Error GoTo ErrorHandler
'Description: Take in any path and if there is not a \ at the end, put it there
'
'Parameters:
' vstrAnyPath - the path in question
'
'Return: A fixed Path
'
'Philip Walsh 29 April 1997
'
Dim strTemp As String

    strTemp = Trim(rstrAnyPath)
  
    If Right(strTemp, 1) <> "\" Then
        strTemp = strTemp & "\"
    Else
        If Right(strTemp, 2) = "\\" Then
            strTemp = Left(strTemp, Len(strTemp) - 1)
        End If
    End If

    If Not IsMissing(NoSlash) Then
        If NoSlash Then
            strTemp = Left(strTemp, Len(strTemp) - 1)
        End If
    End If

FixPath = strTemp

Exit Function
ErrorHandler:
    WriteLogFile Err, MODULE_NAME & ".FixPath:" & Error, mstrPath
    FixPath = rstrAnyPath
End Function

Public Function DeleteTempFiles(ByVal vstrTempdir, Optional DeleteAll As Variant) As Long
On Error GoTo DeleteTempFile_Error
Dim blnDeleteAll As Boolean

    DeleteTempFiles = vbOK
    
    If Not IsMissing(DeleteAll) Then
        blnDeleteAll = DeleteAll
    Else
        blnDeleteAll = False
    End If
  
    If blnDeleteAll Then
        Kill FixPath(vstrTempdir) & "~*.*"
        Kill FixPath(vstrTempdir) & "*.tmp"
        Kill FixPath(vstrTempdir) & "*.rtf"
    Else
        Kill FixPath(vstrTempdir) & "*.tmp"
    End If

Exit Function
DeleteTempFile_Error:
    Select Case Err
        Case 53:
            Resume Next
        Case Else:
            WriteLogFile Err, MODULE_NAME & ".DeleteTempFiles:" & Error, mstrPath
            DeleteTempFiles = vbError
            Exit Function
            Resume
    End Select
End Function

Function NamedRangeExists(rngName As String, Optional ws As Worksheet = Nothing) As Boolean
    Dim N As Name
    NamedRangeExists = False
    For Each N In ThisWorkbook.Names
        If N.Name = rngName Then NamedRangeExists = True: Exit Function
    Next N
    If Not ws Is Nothing Then
        For Each N In ws.Names
            If N.Name = rngName Then NamedRangeExists = True: Exit Function
        Next N
    End If
End Function

Function GetField(rngName As String, hardcodedDefault As Variant, Optional wsName As String = "Assumptions - Proposal infos") As Variant
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(wsName)
    On Error GoTo 0
    If Not ws Is Nothing Then
        If NamedRangeExists(rngName, ws) Then
            On Error Resume Next
            GetField = ws.Range(rngName).Value
            If IsError(GetField) Or IsEmpty(GetField) Then GetField = hardcodedDefault
            On Error GoTo 0
            Exit Function
        End If
    End If
    GetField = hardcodedDefault
End Function

Public Sub MergeAndPrintRTF()
    On Error GoTo ErrHandler

    Dim objWord As Object
    Dim objDoc As Object
    Dim objTable As Object
    Dim objRange As Object
    Dim strMyDocPath As String
    Dim strFileName As String
    Dim strProposalNumber As String

    Dim CUSTOMER_NAME As String: CUSTOMER_NAME = GetField("CUSTOMER_NAME", "Invensys UK")
    Dim INDUSTRY As String: INDUSTRY = GetField("INDUSTRY", "Default")
    Dim PROJECT_MANAGER As String: PROJECT_MANAGER = GetField("PROJECT_MANAGER", "")
    Dim PROJECT_NAME As String: PROJECT_NAME = GetField("PROJECT_NAME", "WRPC")
    Dim DATE_START As String: DATE_START = GetField("DATE_START", "")

    strMyDocPath = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\GECE\"
    If Dir(strMyDocPath, vbDirectory) = "" Then MkDir strMyDocPath

    strProposalNumber = PROJECT_NAME
    If Len(strProposalNumber) = 0 Then strProposalNumber = "NoProjectNumber"
    strFileName = strMyDocPath & strProposalNumber & "_ScopeSummary.docx"

    Set objWord = CreateObject("Word.Application")
    objWord.Visible = True
    Set objDoc = objWord.Documents.Add

    With objDoc.content
        .InsertAfter "GECE Scope Summary" & vbCrLf
        .Paragraphs(1).Range.Font.Bold = True
        .Paragraphs(1).Range.Font.Size = 14
        .InsertParagraphAfter
        .InsertAfter "Proposal Number: " & strProposalNumber & vbCrLf
        .Paragraphs(2).Range.Font.Size = 12
        .InsertAfter "Generated by GECE Tool" & vbCrLf
        .InsertParagraphAfter
    End With

    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 5, 2)
    With objTable
        .cell(1, 1).Range.Text = "Parameter"
        .cell(1, 2).Range.Text = "Value"
        .Rows(1).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(1).Range.Font.Bold = True
        .Rows(1).Range.Font.Size = 12
        .cell(2, 1).Range.Text = "Customer Name"
        .cell(2, 2).Range.Text = CUSTOMER_NAME
        .cell(3, 1).Range.Text = "Industry"
        .cell(3, 2).Range.Text = INDUSTRY
        .cell(4, 1).Range.Text = "Project Manager"
        .cell(4, 2).Range.Text = PROJECT_MANAGER
        .cell(5, 1).Range.Text = "Start Date"
        .cell(5, 2).Range.Text = DATE_START
        .Range.ParagraphFormat.Alignment = 0
        .Columns(1).Width = 120
        .Columns(2).Width = 120
    End With

    objDoc.content.InsertParagraphAfter

    Call Table_DataEntry(objDoc)
    Call Table_AssumptionsProposal(objDoc)
    Call Table_ApplicationBased(objDoc)
    Call Table_DurationBased(objDoc)
    Call Table_TaskBased(objDoc)
    Call Table_PriceMakeUp(objDoc)
    Call Table_APPOverRides(objDoc)
    Call Table_CourseOverRides(objDoc)
    Call Table_CPOverRides(objDoc)
    Call Table_DIOverRides(objDoc)
    Call Table_DOCOverRides(objDoc)
    Call Table_ESDOverRides(objDoc)
    Call Table_FATOverRides(objDoc)
    Call Table_HMIOverRides(objDoc)
    Call Table_PMOverRides(objDoc)
    Call Table_PROPOSAL_SUMMARY_APP(objDoc)
    Call Table_PROPOSAL_SUMMARY_COURSE(objDoc)
    Call Table_PROPOSAL_SUMMARY_CP(objDoc)
    Call Table_PROPOSAL_SUMMARY_DI(objDoc)
    Call Table_PROPOSAL_SUMMARY_DOCS(objDoc)
    Call Table_PROPOSAL_SUMMARY_ESD(objDoc)
    Call Table_PROPOSAL_SUMMARY_HIST_POINTS(objDoc)
    Call Table_PROPOSAL_SUMMARY_HMI(objDoc)
    Call Table_PROPOSAL_SUMMARY_MEETING(objDoc)
    Call Table_PROPOSAL_SUMMARY_REPORTS(objDoc)
    Call Table_PROPOSAL_SUMMARY_TL(objDoc)
    Call Table_Report_Override(objDoc)
    Call Table_SITEOverRides(objDoc)
    Call Table_Spec_OverRide(objDoc)
    Call Table_SystemOverRides(objDoc)
    Call Table_TLOverRides(objDoc)

    objDoc.SaveAs2 strFileName
    MsgBox "GECE Scope Summary created successfully at:" & vbCrLf & strFileName, vbInformation, "GECE"

CleanUp:
    Set objDoc = Nothing
    Set objWord = Nothing
    Exit Sub

ErrHandler:
    MsgBox "MergeAndPrintRTF Error: " & Err.Number & vbCrLf & Err.Description, vbCritical, "GECE Error"
    Resume CleanUp
End Sub

Sub Table_DataEntry(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 61, 2)
    With objTable
        .cell(row, 1).Range.Text = "Field"
        .cell(row, 2).Range.Text = "Value"
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "High Risk Site"
        .cell(row, 2).Range.Text = IIf(GetField("HIGH_RISK_SITE", False, "Data Entry"), "Yes", "No"): row = row + 1
        .cell(row, 1).Range.Text = "Trips Required (FAT)"
        .cell(row, 2).Range.Text = GetField("TL_TRIPS_REQ_FAT", "", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Personnel Quantity (FAT)"
        .cell(row, 2).Range.Text = GetField("TL_PERS_QTY_FAT", "", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Days Quantity (FAT)"
        .cell(row, 2).Range.Text = GetField("TL_DAYS_QTY_FAT", "", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Airfare (FAT)"
        .cell(row, 2).Range.Text = GetField("TL_AIRFARE_FAT", "", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Daily Allowance (Site)"
        .cell(row, 2).Range.Text = GetField("TL_DAILY_ALLOW_SITE", "", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Trips Required (Site)"
        .cell(row, 2).Range.Text = GetField("TL_TRIPS_REQ_SITE", "", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Personnel Quantity (Site)"
        .cell(row, 2).Range.Text = GetField("TL_PERS_QTY_SITE", "", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Days Quantity (Site)"
        .cell(row, 2).Range.Text = GetField("TL_DAYS_QTY_SITE", "", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Airfare (Site)"
        .cell(row, 2).Range.Text = GetField("TL_AIRFARE_SITE", "", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "APP 1 Hours"
        .cell(row, 2).Range.Text = GetField("APP_1_HOURS", "", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "APP 2 Hours"
        .cell(row, 2).Range.Text = GetField("APP_2_HOURS", "", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "APP 3 Hours"
        .cell(row, 2).Range.Text = GetField("APP_3_HOURS", "", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "APP 4 Hours"
        .cell(row, 2).Range.Text = GetField("APP_4_HOURS", "", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "APP 5 Hours"
        .cell(row, 2).Range.Text = GetField("APP_5_HOURS", "", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "APP 6 Hours"
        .cell(row, 2).Range.Text = GetField("APP_6_HOURS", "", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "APP 7 Hours"
        .cell(row, 2).Range.Text = GetField("APP_7_HOURS", "", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "APP 8 Hours"
        .cell(row, 2).Range.Text = GetField("APP_8_HOURS", "", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "APP Bus STS"
        .cell(row, 2).Range.Text = IIf(GetField("APP_BUS_STS", False, "Data Entry"), "Yes", "No"): row = row + 1
        .cell(row, 1).Range.Text = "Cab Consoles Qty"
        .cell(row, 2).Range.Text = GetField("CAB_CONSOLES_QTY", "", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Cab IO Qty"
        .cell(row, 2).Range.Text = GetField("CAB_IO_QTY", "", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Cab Marsh Qty"
        .cell(row, 2).Range.Text = GetField("CAB_MARSH_QTY", "", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Cab Proc Qty"
        .cell(row, 2).Range.Text = GetField("CAB_PROC_QTY", "", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Cost Buyout"
        .cell(row, 2).Range.Text = GetField("COST_BUYOUT", 0, "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Cost IA"
        .cell(row, 2).Range.Text = GetField("COST_IA", 0, "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Course 1 Hours"
        .cell(row, 2).Range.Text = GetField("COURSE_1_HOURS", 0, "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Course 2 Hours"
        .cell(row, 2).Range.Text = GetField("COURSE_2_HOURS", 0, "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Course 3 Hours"
        .cell(row, 2).Range.Text = GetField("COURSE_3_HOURS", 0, "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Course 4 Hours"
        .cell(row, 2).Range.Text = GetField("COURSE_4_HOURS", 0, "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Course 5 Hours"
        .cell(row, 2).Range.Text = GetField("COURSE_5_HOURS", 0, "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Course1 Hours"
        .cell(row, 2).Range.Text = GetField("COURSE1_HOURS", 0, "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Course1 Name"
        .cell(row, 2).Range.Text = GetField("COURSE1_NAME", "Course 1", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Course2 Hours"
        .cell(row, 2).Range.Text = GetField("COURSE2_HOURS", 0, "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Course2 Name"
        .cell(row, 2).Range.Text = GetField("COURSE2_NAME", "Course 2", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Course3 Hours"
        .cell(row, 2).Range.Text = GetField("COURSE3_HOURS", 0, "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Course3 Name"
        .cell(row, 2).Range.Text = GetField("COURSE3_NAME", "Course 3", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Course4 Hours"
        .cell(row, 2).Range.Text = GetField("COURSE4_HOURS", 0, "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Course4 Name"
        .cell(row, 2).Range.Text = GetField("COURSE4_NAME", "Course 4", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Course5 Hours"
        .cell(row, 2).Range.Text = GetField("COURSE5_HOURS", 0, "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Course5 Name"
        .cell(row, 2).Range.Text = GetField("COURSE5_NAME", "Course 5", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "CP AI"
        .cell(row, 2).Range.Text = GetField("CP_AI", 0, "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "CP ANA COMPLEX Qty"
        .cell(row, 2).Range.Text = GetField("CP_ANA_COMPLEX_QTY", 0, "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "CP AO"
        .cell(row, 2).Range.Text = GetField("CP_AO", 0, "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "CP DI"
        .cell(row, 2).Range.Text = GetField("CP_DI", 0, "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "CP DIGITAL COMPLEX Qty"
        .cell(row, 2).Range.Text = GetField("CP_DIGITAL_COMPLEX_QTY", "", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "CP DIGITAL CTRL DI"
        .cell(row, 2).Range.Text = GetField("CP_DIGITAL_CTRL_DI", 2, "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "CP DIGITAL CTRL DO"
        .cell(row, 2).Range.Text = GetField("CP_DIGITAL_CTRL_DO", 1, "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "CP DO"
        .cell(row, 2).Range.Text = GetField("CP_DO", 0, "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "CP FIELDBUS IO Qty"
        .cell(row, 2).Range.Text = GetField("CP_FIELDBUS_IO_QTY", "", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "CP FIELDBUS IO Ratio"
        .cell(row, 2).Range.Text = GetField("CP_FIELDBUS_IO_RATIO", "", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "CP GRP START COMPLEX Qty"
        .cell(row, 2).Range.Text = GetField("CP_GRP_START_COMPLEX_QTY", "", "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "CP GRP START LOOP Qty"
        .cell(row, 2).Range.Text = GetField("CP_GRP_START_LOOP_QTY", "", "Data Entry")
    End With
    objDoc.content.InsertParagraphAfter
End Sub
Sub Table_AssumptionsProposal(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 29, 2)
    With objTable
        .cell(row, 1).Range.Text = "Field"
        .cell(row, 2).Range.Text = "Value"
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "Toolkit 1 Req"
        .cell(row, 2).Range.Text = IIf(GetField("TOOLKIT_1_REQ", False), "Yes", "No"): row = row + 1
        .cell(row, 1).Range.Text = "Toolkit 2 Req"
        .cell(row, 2).Range.Text = IIf(GetField("TOOLKIT_2_REQ", False), "Yes", "No"): row = row + 1
        .cell(row, 1).Range.Text = "Toolkit 3 Req"
        .cell(row, 2).Range.Text = IIf(GetField("TOOLKIT_3_REQ", False), "Yes", "No"): row = row + 1
        .cell(row, 1).Range.Text = "Toolkit 4 Req"
        .cell(row, 2).Range.Text = IIf(GetField("TOOLKIT_4_REQ", False), "Yes", "No"): row = row + 1
        .cell(row, 1).Range.Text = "Toolkit 5 Req"
        .cell(row, 2).Range.Text = IIf(GetField("TOOLKIT_5_REQ", False), "Yes", "No"): row = row + 1
        .cell(row, 1).Range.Text = "Toolkit 6 Req"
        .cell(row, 2).Range.Text = IIf(GetField("TOOLKIT_6_REQ", False), "Yes", "No"): row = row + 1
        .cell(row, 1).Range.Text = "Toolkit 7 Req"
        .cell(row, 2).Range.Text = IIf(GetField("TOOLKIT_7_REQ", False), "Yes", "No"): row = row + 1
        .cell(row, 1).Range.Text = "Toolkit 8 Req"
        .cell(row, 2).Range.Text = IIf(GetField("TOOLKIT_8_REQ", False), "Yes", "No"): row = row + 1
        .cell(row, 1).Range.Text = "Toolkit 9 Req"
        .cell(row, 2).Range.Text = IIf(GetField("TOOLKIT_9_REQ", False), "Yes", "No"): row = row + 1
        .cell(row, 1).Range.Text = "Toolkit 10 Req"
        .cell(row, 2).Range.Text = IIf(GetField("TOOLKIT_10_REQ", False), "Yes", "No"): row = row + 1
        .cell(row, 1).Range.Text = "Toolkit 11 Req"
        .cell(row, 2).Range.Text = IIf(GetField("TOOLKIT_11_REQ", False), "Yes", "No"): row = row + 1
        .cell(row, 1).Range.Text = "Toolkit 12 Req"
        .cell(row, 2).Range.Text = IIf(GetField("TOOLKIT_12_REQ", False), "Yes", "No"): row = row + 1
        .cell(row, 1).Range.Text = "Toolkit 13 Req"
        .cell(row, 2).Range.Text = IIf(GetField("TOOLKIT_13_REQ", False), "Yes", "No"): row = row + 1
        .cell(row, 1).Range.Text = "Toolkit 14 Req"
        .cell(row, 2).Range.Text = IIf(GetField("TOOLKIT_14_REQ", False), "Yes", "No"): row = row + 1
        .cell(row, 1).Range.Text = "Toolkit 15 Req"
        .cell(row, 2).Range.Text = IIf(GetField("TOOLKIT_15_REQ", False), "Yes", "No"): row = row + 1
        .cell(row, 1).Range.Text = "Toolkit 16 Req"
        .cell(row, 2).Range.Text = IIf(GetField("TOOLKIT_16_REQ", False), "Yes", "No"): row = row + 1
        .cell(row, 1).Range.Text = "Toolkit 17 Req"
        .cell(row, 2).Range.Text = IIf(GetField("TOOLKIT_17_REQ", False), "Yes", "No"): row = row + 1
        .cell(row, 1).Range.Text = "Toolkit 18 Req"
        .cell(row, 2).Range.Text = IIf(GetField("TOOLKIT_18_REQ", False), "Yes", "No"): row = row + 1
        .cell(row, 1).Range.Text = "Toolkit 19 Req"
        .cell(row, 2).Range.Text = IIf(GetField("TOOLKIT_19_REQ", False), "Yes", "No"): row = row + 1
        .cell(row, 1).Range.Text = "Toolkit 20 Req"
        .cell(row, 2).Range.Text = IIf(GetField("TOOLKIT_20_REQ", False), "Yes", "No"): row = row + 1
        .cell(row, 1).Range.Text = "Toolkit 21 Req"
        .cell(row, 2).Range.Text = IIf(GetField("TOOLKIT_21_REQ", False), "Yes", "No"): row = row + 1
        .cell(row, 1).Range.Text = "Toolkit 22 Req"
        .cell(row, 2).Range.Text = IIf(GetField("TOOLKIT_22_REQ", False), "Yes", "No"): row = row + 1
        .cell(row, 1).Range.Text = "Customer Name"
        .cell(row, 2).Range.Text = GetField("CUSTOMER_NAME", "Invensys UK"): row = row + 1
        .cell(row, 1).Range.Text = "Industry"
        .cell(row, 2).Range.Text = GetField("INDUSTRY", "Default"): row = row + 1
        .cell(row, 1).Range.Text = "Project Manager"
        .cell(row, 2).Range.Text = GetField("PROJECT_MANAGER", ""): row = row + 1
        .cell(row, 1).Range.Text = "Local Country"
        .cell(row, 2).Range.Text = GetField("LOCAL_COUNTRY", "UK"): row = row + 1
        .cell(row, 1).Range.Text = "Default Remote Country"
        .cell(row, 2).Range.Text = GetField("DEFAULT_REM_COUNTRY", "Egypt (EEC)")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_ApplicationBased(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 23, 2)
    With objTable
        .cell(row, 1).Range.Text = "Field"
        .cell(row, 2).Range.Text = "Value"
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "APP Task 1 Override Justification"
        .cell(row, 2).Range.Text = GetField("APP_TASK1_OVD_JUST", "", "Application Based"): row = row + 1
        .cell(row, 1).Range.Text = "APP Task 1 Override Quantity"
        .cell(row, 2).Range.Text = GetField("APP_TASK1_OVD_QTY", 0, "Application Based"): row = row + 1
        .cell(row, 1).Range.Text = "APP Task 1 Remote Country"
        .cell(row, 2).Range.Text = GetField("APP_TASK1_REM_COUNTRY", "Egypt (EEC)", "Application Based"): row = row + 1
        .cell(row, 1).Range.Text = "APP Task 10 Override Justification"
        .cell(row, 2).Range.Text = GetField("APP_TASK10_OVD_JUST", "", "Application Based"): row = row + 1
        .cell(row, 1).Range.Text = "APP Task 10 Override Quantity"
        .cell(row, 2).Range.Text = GetField("APP_TASK10_OVD_QTY", 0, "Application Based"): row = row + 1
        .cell(row, 1).Range.Text = "APP Task 10 Remote Country"
        .cell(row, 2).Range.Text = GetField("APP_TASK10_REM_COUNTRY", "Egypt (EEC)", "Application Based"): row = row + 1
        .cell(row, 1).Range.Text = "APP Task 10 Remote Percentage"
        .cell(row, 2).Range.Text = GetField("APP_TASK10_REM_PCT", 0, "Application Based"): row = row + 1
        .cell(row, 1).Range.Text = "DUP Fact APP Based APP"
        .cell(row, 2).Range.Text = GetField("DUP_FACT_APP_BASED_APP", 0.4, "Application Based"): row = row + 1
        .cell(row, 1).Range.Text = "DUP Fact APP Based Course"
        .cell(row, 2).Range.Text = GetField("DUP_FACT_APP_BASED_COURSE", 1, "Application Based"): row = row + 1
        .cell(row, 1).Range.Text = "DUP Fact APP Based CP"
        .cell(row, 2).Range.Text = GetField("DUP_FACT_APP_BASED_CP", 0.4, "Application Based"): row = row + 1
        .cell(row, 1).Range.Text = "DUP Fact APP Based DI"
        .cell(row, 2).Range.Text = GetField("DUP_FACT_APP_BASED_DI", 0.4, "Application Based"): row = row + 1
        .cell(row, 1).Range.Text = "DUP Fact APP Based DOC"
        .cell(row, 2).Range.Text = GetField("DUP_FACT_APP_BASED_DOC", 0.4, "Application Based"): row = row + 1
        .cell(row, 1).Range.Text = "DUP Fact APP Based ESD"
        .cell(row, 2).Range.Text = GetField("DUP_FACT_APP_BASED_ESD", 0.4, "Application Based"): row = row + 1
        .cell(row, 1).Range.Text = "DUP Fact APP Based HMI"
        .cell(row, 2).Range.Text = GetField("DUP_FACT_APP_BASED_HMI", 0.4, "Application Based"): row = row + 1
        .cell(row, 1).Range.Text = "DUP Fact APP Based MEETING"
        .cell(row, 2).Range.Text = GetField("DUP_FACT_APP_BASED_MEETING", 0.2, "Application Based"): row = row + 1
        .cell(row, 1).Range.Text = "DUP Fact APP Based PM"
        .cell(row, 2).Range.Text = GetField("DUP_FACT_APP_BASED_PM", 0.5, "Application Based"): row = row + 1
        .cell(row, 1).Range.Text = "DUP Fact APP Based REP"
        .cell(row, 2).Range.Text = GetField("DUP_FACT_APP_BASED_REP", 0.4, "Application Based"): row = row + 1
        .cell(row, 1).Range.Text = "DUP Fact APP Based SITE"
        .cell(row, 2).Range.Text = GetField("DUP_FACT_APP_BASED_SITE", 1, "Application Based"): row = row + 1
        .cell(row, 1).Range.Text = "DUP Fact APP Based SPEC"
        .cell(row, 2).Range.Text = GetField("DUP_FACT_APP_BASED_SPEC", 0.4, "Application Based"): row = row + 1
        .cell(row, 1).Range.Text = "DUP Fact APP Based SYS ENG"
        .cell(row, 2).Range.Text = GetField("DUP_FACT_APP_BASED_SYS_ENG", 0.4, "Application Based"): row = row + 1
        .cell(row, 1).Range.Text = "DUP Fact APP Based SYSENG"
        .cell(row, 2).Range.Text = GetField("DUP_FACT_APP_BASED_SYSENG", 0.4, "Application Based"): row = row + 1
        .cell(row, 1).Range.Text = "DUP Fact APP Based TEST"
        .cell(row, 2).Range.Text = GetField("DUP_FACT_APP_BASED_TEST", 1, "Application Based")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_DurationBased(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 17, 2)
    With objTable
        .cell(row, 1).Range.Text = "Field"
        .cell(row, 2).Range.Text = "Value"
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "Duration Loc SE Qty"
        .cell(row, 2).Range.Text = GetField("DURATION_LOC_SE_QTY", 0, "Duration Based"): row = row + 1
        .cell(row, 1).Range.Text = "Duration Loc SE Week"
        .cell(row, 2).Range.Text = GetField("DURATION_LOC_SE_WEEK", 0, "Duration Based"): row = row + 1
        .cell(row, 1).Range.Text = "Duration Loc SE Dup"
        .cell(row, 2).Range.Text = GetField("DURATION_LOC_SE_DUP", 1, "Duration Based"): row = row + 1
        .cell(row, 1).Range.Text = "Duration Loc SE Util"
        .cell(row, 2).Range.Text = GetField("DURATION_LOC_SE_UTIL", 1, "Duration Based"): row = row + 1
        .cell(row, 1).Range.Text = "Duration Loc SR Qty"
        .cell(row, 2).Range.Text = GetField("DURATION_LOC_SR_QTY", 0, "Duration Based"): row = row + 1
        .cell(row, 1).Range.Text = "Duration Loc SR Week"
        .cell(row, 2).Range.Text = GetField("DURATION_LOC_SR_WEEK", 0, "Duration Based"): row = row + 1
        .cell(row, 1).Range.Text = "Duration Loc SSE Dup"
        .cell(row, 2).Range.Text = GetField("DURATION_LOC_SSE_DUP", 1, "Duration Based"): row = row + 1
        .cell(row, 1).Range.Text = "Duration Loc SR Util"
        .cell(row, 2).Range.Text = GetField("DURATION_LOC_SR_UTIL", 1, "Duration Based"): row = row + 1
        .cell(row, 1).Range.Text = "Duration Rem SE Qty"
        .cell(row, 2).Range.Text = GetField("DURATION_REM_SE_QTY", 0, "Duration Based"): row = row + 1
        .cell(row, 1).Range.Text = "Duration Rem SE Week"
        .cell(row, 2).Range.Text = GetField("DURATION_REM_SE_WEEK", 0, "Duration Based"): row = row + 1
        .cell(row, 1).Range.Text = "Duration SE Rem Country"
        .cell(row, 2).Range.Text = GetField("DURATION_SE_REM_COUNTRY", "Egypt (EEC)", "Duration Based"): row = row + 1
        .cell(row, 1).Range.Text = "Duration Rem SE Dup"
        .cell(row, 2).Range.Text = GetField("DURATION_REM_SE_DUP", 1, "Duration Based"): row = row + 1
        .cell(row, 1).Range.Text = "Duration Rem SE Util"
        .cell(row, 2).Range.Text = GetField("DURATION_REM_SE_UTIL", 1, "Duration Based"): row = row + 1
        .cell(row, 1).Range.Text = "Duration Rem SR Qty"
        .cell(row, 2).Range.Text = GetField("DURATION_REM_SR_QTY", 0, "Duration Based"): row = row + 1
        .cell(row, 1).Range.Text = "Duration Rem SR Week"
        .cell(row, 2).Range.Text = GetField("DURATION_REM_SR_WEEK", 0, "Duration Based"): row = row + 1
        .cell(row, 1).Range.Text = "Duration SR Rem Country"
        .cell(row, 2).Range.Text = GetField("DURATION_SR_REM_COUNTRY", "Egypt (EEC)", "Duration Based"): row = row + 1
        .cell(row, 1).Range.Text = "Duration Rem SR Dup"
        .cell(row, 2).Range.Text = GetField("DURATION_REM_SR_DUP", 1, "Duration Based"): row = row + 1
        .cell(row, 1).Range.Text = "Duration Rem SR Util"
        .cell(row, 2).Range.Text = GetField("DURATION_REM_SR_UTIL", 1, "Duration Based")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_TaskBased(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 11, 2)
    With objTable
        .cell(row, 1).Range.Text = "Field"
        .cell(row, 2).Range.Text = "Value"
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "CP Task 1 Override Just"
        .cell(row, 2).Range.Text = GetField("CP_TASK1_OVD_JUST", "", "Task Based"): row = row + 1
        .cell(row, 1).Range.Text = "CP Task 1 Override Qty"
        .cell(row, 2).Range.Text = GetField("CP_TASK1_OVD_QTY", 0, "Task Based"): row = row + 1
        .cell(row, 1).Range.Text = "CP Task 2 Override Just"
        .cell(row, 2).Range.Text = GetField("CP_TASK2_OVD_JUST", "", "Task Based"): row = row + 1
        .cell(row, 1).Range.Text = "CP Task 2 Override Qty"
        .cell(row, 2).Range.Text = GetField("CP_TASK2_OVD_QTY", 0, "Task Based"): row = row + 1
        .cell(row, 1).Range.Text = "DI Task 1 Override Just"
        .cell(row, 2).Range.Text = GetField("DI_TASK1_OVD_JUST", "", "Task Based"): row = row + 1
        .cell(row, 1).Range.Text = "DI Task 1 Override Qty"
        .cell(row, 2).Range.Text = GetField("DI_TASK1_OVD_QTY", 0, "Task Based"): row = row + 1
        .cell(row, 1).Range.Text = "DI Task 2 Override Just"
        .cell(row, 2).Range.Text = GetField("DI_TASK2_OVD_JUST", "", "Task Based"): row = row + 1
        .cell(row, 1).Range.Text = "DI Task 2 Override Qty"
        .cell(row, 2).Range.Text = GetField("DI_TASK2_OVD_QTY", 0, "Task Based"): row = row + 1
        .cell(row, 1).Range.Text = "DOC Task 1 Override Just"
        .cell(row, 2).Range.Text = GetField("DOC_TASK1_OVD_JUST", "", "Task Based"): row = row + 1
        .cell(row, 1).Range.Text = "DOC Task 1 Override Qty"
        .cell(row, 2).Range.Text = GetField("DOC_TASK1_OVD_QTY", 0, "Task Based")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_PriceMakeUp(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 7, 2)
    With objTable
        .cell(row, 1).Range.Text = "Field"
        .cell(row, 2).Range.Text = "Value"
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "Makeup Engineering"
        .cell(row, 2).Range.Text = GetField("MAKEUP_ENGINEERING", 0, "Price Make-up"): row = row + 1
        .cell(row, 1).Range.Text = "Makeup Commissioning"
        .cell(row, 2).Range.Text = GetField("MAKEUP_COMMISSIONING", 0, "Price Make-up"): row = row + 1
        .cell(row, 1).Range.Text = "Makeup Training"
        .cell(row, 2).Range.Text = GetField("MAKEUP_TRAINING", 0, "Price Make-up"): row = row + 1
        .cell(row, 1).Range.Text = "Makeup Hardware"
        .cell(row, 2).Range.Text = GetField("MAKEUP_HARDWARE", 0, "Price Make-up"): row = row + 1
        .cell(row, 1).Range.Text = "Makeup Software"
        .cell(row, 2).Range.Text = GetField("MAKEUP_SOFTWARE", 0, "Price Make-up"): row = row + 1
        .cell(row, 1).Range.Text = "Makeup Buyout"
        .cell(row, 2).Range.Text = GetField("MAKEUP_BUYOUT", 0, "Price Make-up"): row = row + 1
        .cell(row, 1).Range.Text = "Makeup Project Management"
        .cell(row, 2).Range.Text = GetField("MAKEUP_PM", 0, "Price Make-up")
    End With
    objDoc.content.InsertParagraphAfter
End Sub
Sub Table_APPOverRides(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 4, 2)
    With objTable
        .cell(row, 1).Range.Text = "Override"
        .cell(row, 2).Range.Text = "Value"
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "APP Task 1"
        .cell(row, 2).Range.Text = GetField("APP_TASK1_OVD_QTY", 0, "Application Based"): row = row + 1
        .cell(row, 1).Range.Text = "APP Task 10"
        .cell(row, 2).Range.Text = GetField("APP_TASK10_OVD_QTY", 0, "Application Based"): row = row + 1
        .cell(row, 1).Range.Text = "APP Task 10 Remote Percentage"
        .cell(row, 2).Range.Text = GetField("APP_TASK10_REM_PCT", 0, "Application Based")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_CourseOverRides(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 6, 2)
    With objTable
        .cell(row, 1).Range.Text = "Override"
        .cell(row, 2).Range.Text = "Value"
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "Course1 Hours"
        .cell(row, 2).Range.Text = GetField("COURSE1_HOURS", 0, "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Course2 Hours"
        .cell(row, 2).Range.Text = GetField("COURSE2_HOURS", 0, "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Course3 Hours"
        .cell(row, 2).Range.Text = GetField("COURSE3_HOURS", 0, "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Course4 Hours"
        .cell(row, 2).Range.Text = GetField("COURSE4_HOURS", 0, "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Course5 Hours"
        .cell(row, 2).Range.Text = GetField("COURSE5_HOURS", 0, "Data Entry")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_CPOverRides(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 5, 2)
    With objTable
        .cell(row, 1).Range.Text = "Override"
        .cell(row, 2).Range.Text = "Value"
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "CP Task 1 Override Qty"
        .cell(row, 2).Range.Text = GetField("CP_TASK1_OVD_QTY", 0, "Task Based"): row = row + 1
        .cell(row, 1).Range.Text = "CP Task 2 Override Qty"
        .cell(row, 2).Range.Text = GetField("CP_TASK2_OVD_QTY", 0, "Task Based"): row = row + 1
        .cell(row, 1).Range.Text = "CP ANA COMPLEX Qty"
        .cell(row, 2).Range.Text = GetField("CP_ANA_COMPLEX_QTY", 0, "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "CP DIGITAL COMPLEX Qty"
        .cell(row, 2).Range.Text = GetField("CP_DIGITAL_COMPLEX_QTY", 0, "Data Entry")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_DIOverRides(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 3, 2)
    With objTable
        .cell(row, 1).Range.Text = "Override"
        .cell(row, 2).Range.Text = "Value"
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "DI Task 1 Override Qty"
        .cell(row, 2).Range.Text = GetField("DI_TASK1_OVD_QTY", 0, "Task Based"): row = row + 1
        .cell(row, 1).Range.Text = "DI Task 2 Override Qty"
        .cell(row, 2).Range.Text = GetField("DI_TASK2_OVD_QTY", 0, "Task Based")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_DOCOverRides(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 3, 2)
    With objTable
        .cell(row, 1).Range.Text = "Override"
        .cell(row, 2).Range.Text = "Value"
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "DOC Task 1 Override Qty"
        .cell(row, 2).Range.Text = GetField("DOC_TASK1_OVD_QTY", 0, "Task Based"): row = row + 1
        .cell(row, 1).Range.Text = "DOC Task 2 Override Qty"
        .cell(row, 2).Range.Text = GetField("DOC_TASK2_OVD_QTY", 0, "Task Based")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_ESDOverRides(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 2, 2)
    With objTable
        .cell(row, 1).Range.Text = "Override"
        .cell(row, 2).Range.Text = "Value"
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "ESD Override"
        .cell(row, 2).Range.Text = GetField("ESD_OVERRIDE", 0, "Data Entry")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_FATOverRides(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 2, 2)
    With objTable
        .cell(row, 1).Range.Text = "Override"
        .cell(row, 2).Range.Text = "Value"
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "FAT Override"
        .cell(row, 2).Range.Text = GetField("FAT_OVERRIDE", 0, "Data Entry")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_HMIOverRides(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 2, 2)
    With objTable
        .cell(row, 1).Range.Text = "Override"
        .cell(row, 2).Range.Text = "Value"
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "HMI Override"
        .cell(row, 2).Range.Text = GetField("HMI_OVERRIDE", 0, "Data Entry")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_PMOverRides(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 2, 2)
    With objTable
        .cell(row, 1).Range.Text = "Override"
        .cell(row, 2).Range.Text = "Value"
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "PM Override"
        .cell(row, 2).Range.Text = GetField("PM_OVERRIDE", 0, "Data Entry")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_SITEOverRides(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 2, 2)
    With objTable
        .cell(row, 1).Range.Text = "Override"
        .cell(row, 2).Range.Text = "Value"
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "Site Override"
        .cell(row, 2).Range.Text = GetField("SITE_OVERRIDE", 0, "Data Entry")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_Spec_OverRide(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 2, 2)
    With objTable
        .cell(row, 1).Range.Text = "Override"
        .cell(row, 2).Range.Text = "Value"
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "Spec Override"
        .cell(row, 2).Range.Text = GetField("SPEC_OVERRIDE", 0, "Data Entry")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_SystemOverRides(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 2, 2)
    With objTable
        .cell(row, 1).Range.Text = "Override"
        .cell(row, 2).Range.Text = "Value"
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "System Override"
        .cell(row, 2).Range.Text = GetField("SYSTEM_OVERRIDE", 0, "Data Entry")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_TLOverRides(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 2, 2)
    With objTable
        .cell(row, 1).Range.Text = "Override"
        .cell(row, 2).Range.Text = "Value"
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "TL Override"
        .cell(row, 2).Range.Text = GetField("TL_OVERRIDE", 0, "Data Entry")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_PROPOSAL_SUMMARY_APP(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 3, 2)
    With objTable
        .cell(row, 1).Range.Text = "Proposal Summary (APP)"
        .cell(row, 2).Range.Text = ""
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "APP Total Hours"
        .cell(row, 2).Range.Text = GetField("APP_TOTAL_HOURS", 0, "Application Based"): row = row + 1
        .cell(row, 1).Range.Text = "APP Remote Country"
        .cell(row, 2).Range.Text = GetField("APP_TASK10_REM_COUNTRY", "Egypt (EEC)", "Application Based")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_PROPOSAL_SUMMARY_COURSE(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 4, 2)
    With objTable
        .cell(row, 1).Range.Text = "Proposal Summary (COURSE)"
        .cell(row, 2).Range.Text = ""
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "Course 1 Hours"
        .cell(row, 2).Range.Text = GetField("COURSE1_HOURS", 0, "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Course 2 Hours"
        .cell(row, 2).Range.Text = GetField("COURSE2_HOURS", 0, "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "Course 3 Hours"
        .cell(row, 2).Range.Text = GetField("COURSE3_HOURS", 0, "Data Entry")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_PROPOSAL_SUMMARY_CP(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 3, 2)
    With objTable
        .cell(row, 1).Range.Text = "Proposal Summary (CP)"
        .cell(row, 2).Range.Text = ""
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "CP ANA COMPLEX Qty"
        .cell(row, 2).Range.Text = GetField("CP_ANA_COMPLEX_QTY", 0, "Data Entry"): row = row + 1
        .cell(row, 1).Range.Text = "CP DIGITAL COMPLEX Qty"
        .cell(row, 2).Range.Text = GetField("CP_DIGITAL_COMPLEX_QTY", 0, "Data Entry")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_PROPOSAL_SUMMARY_DI(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 3, 2)
    With objTable
        .cell(row, 1).Range.Text = "Proposal Summary (DI)"
        .cell(row, 2).Range.Text = ""
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "DI Task 1 Override Qty"
        .cell(row, 2).Range.Text = GetField("DI_TASK1_OVD_QTY", 0, "Task Based"): row = row + 1
        .cell(row, 1).Range.Text = "DI Task 2 Override Qty"
        .cell(row, 2).Range.Text = GetField("DI_TASK2_OVD_QTY", 0, "Task Based")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_PROPOSAL_SUMMARY_DOCS(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 3, 2)
    With objTable
        .cell(row, 1).Range.Text = "Proposal Summary (DOCS)"
        .cell(row, 2).Range.Text = ""
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "DOC Task 1 Override Qty"
        .cell(row, 2).Range.Text = GetField("DOC_TASK1_OVD_QTY", 0, "Task Based"): row = row + 1
        .cell(row, 1).Range.Text = "DOC Task 2 Override Qty"
        .cell(row, 2).Range.Text = GetField("DOC_TASK2_OVD_QTY", 0, "Task Based")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_PROPOSAL_SUMMARY_ESD(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 2, 2)
    With objTable
        .cell(row, 1).Range.Text = "Proposal Summary (ESD)"
        .cell(row, 2).Range.Text = ""
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "ESD Override"
        .cell(row, 2).Range.Text = GetField("ESD_OVERRIDE", 0, "Data Entry")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_PROPOSAL_SUMMARY_HIST_POINTS(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 2, 2)
    With objTable
        .cell(row, 1).Range.Text = "Proposal Summary (HIST_POINTS)"
        .cell(row, 2).Range.Text = ""
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "HIST POINTS"
        .cell(row, 2).Range.Text = GetField("HIST_POINTS", 0, "Data Entry")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_PROPOSAL_SUMMARY_HMI(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 2, 2)
    With objTable
        .cell(row, 1).Range.Text = "Proposal Summary (HMI)"
        .cell(row, 2).Range.Text = ""
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "HMI Override"
        .cell(row, 2).Range.Text = GetField("HMI_OVERRIDE", 0, "Data Entry")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_PROPOSAL_SUMMARY_MEETING(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 2, 2)
    With objTable
        .cell(row, 1).Range.Text = "Proposal Summary (MEETING)"
        .cell(row, 2).Range.Text = ""
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "Meeting Points"
        .cell(row, 2).Range.Text = GetField("MEETING_POINTS", 0, "Data Entry")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_PROPOSAL_SUMMARY_REPORTS(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 2, 2)
    With objTable
        .cell(row, 1).Range.Text = "Proposal Summary (REPORTS)"
        .cell(row, 2).Range.Text = ""
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "Reports Points"
        .cell(row, 2).Range.Text = GetField("REPORTS_POINTS", 0, "Data Entry")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_PROPOSAL_SUMMARY_TL(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 2, 2)
    With objTable
        .cell(row, 1).Range.Text = "Proposal Summary (TL)"
        .cell(row, 2).Range.Text = ""
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "TL Override"
        .cell(row, 2).Range.Text = GetField("TL_OVERRIDE", 0, "Data Entry")
    End With
    objDoc.content.InsertParagraphAfter
End Sub

Sub Table_Report_Override(objDoc As Object)
    Dim objTable As Object, objRange As Object
    Dim row As Integer: row = 1
    Set objRange = objDoc.content.Paragraphs.Last.Range
    Set objTable = objDoc.Tables.Add(objRange, 2, 2)
    With objTable
        .cell(row, 1).Range.Text = "Report Override"
        .cell(row, 2).Range.Text = "Value"
        .Rows(row).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
        .Rows(row).Range.Font.Bold = True
        .Rows(row).Range.Font.Size = 12
        .Columns(1).Width = 180
        .Columns(2).Width = 80
        row = row + 1
        .cell(row, 1).Range.Text = "Report Override"
        .cell(row, 2).Range.Text = GetField("REPORT_OVERRIDE", 0, "Data Entry")
    End With
    objDoc.content.InsertParagraphAfter
End Sub
Function CreatePATfile(path As String, filename As String)

    Dim filePath As String
    Dim MyConfigFile As Integer
    Dim strBuffer As String
    
    MyConfigFile = FreeFile
    filePath = Left(path, Len(path) - 4) & ".dai"

    'Opens the file
    Open filePath For Output As MyConfigFile
    
    'Print the header
    Print #MyConfigFile, "|GECE Scope Document|1|False"
    Print #MyConfigFile, "DAC_3939159442|GECE Scope – Heading Title|2|False"
    Print #MyConfigFile, "|Overview|2|False"
    Print #MyConfigFile, "DAC_3939159444|Overview – Heading Title|3|False"
    Print #MyConfigFile, "|Specifications|2|False"
    Print #MyConfigFile, "DAC_3939159446|Specifications – Heading Title|3|False"
    Print #MyConfigFile, "DAC_3939159447|Specifications - Engineers resumes|3|False"
    Print #MyConfigFile, "|System Engineering|2|False"
    Print #MyConfigFile, "DAC_3939159449|System Engineering– Heading Title|3|False"
    Print #MyConfigFile, "|Human Interface Configuration|2|False"
    Print #MyConfigFile, "DAC_39391594411|Human Interface Configuration– Heading Title|3|False"
    Print #MyConfigFile, "|Controls Configuration|2|False"
    Print #MyConfigFile, "DAC_39391594514|Controls Configuration– Heading Title|3|False"
    Print #MyConfigFile, "|Device Interface Configuration|2|False"
    Print #MyConfigFile, "DAC_39391594516|Device Interface Configuration– Heading Title|3|False"
    Print #MyConfigFile, "|Safety System Configuration|2|False"
    Print #MyConfigFile, "DAC_39391594518|Safety System Configuration– Heading Title|3|False"
    Print #MyConfigFile, "|Reports Configuration|2|False"
    Print #MyConfigFile, "DAC_39391594520|Reports Configuration– Heading Title|3|False"
    Print #MyConfigFile, "|Applications |2|False"
    Print #MyConfigFile, "DAC_39391594522|Applications– Heading Title|3|False"
    Print #MyConfigFile, "|Testing |2|False"
    Print #MyConfigFile, "DAC_39391594524|Testing– Heading Title|3|False"
    Print #MyConfigFile, "|Documentation |2|False"
    Print #MyConfigFile, "DAC_39391594526|Documentation– Heading Title|3|False"
    Print #MyConfigFile, "|Training |2|False"
    Print #MyConfigFile, "DAC_39391594528|Training– Heading Title|3|False"
    Print #MyConfigFile, "|Project Management |2|False"
    Print #MyConfigFile, "DAC_39391594530|Project Management– Heading Title|3|False"
    Print #MyConfigFile, "|Site Services |2|False"
    Print #MyConfigFile, "DAC_39391594532|Site Services– Heading Title|3|False"
    Print #MyConfigFile, "DAC_39391594533|Installation Supervision– Heading Title|3|False"
    Print #MyConfigFile, "DAC_39391594534|Commissioning– Heading Title|3|False"
    Print #MyConfigFile, "DAC_39391594535|Site Acceptance Test– Heading Title|3|False"
    Print #MyConfigFile, "DAC_39391594536|Site Survey– Heading Title|3|False"
    Print #MyConfigFile, "|Travel & Living|2|False"
    Print #MyConfigFile, "DAC_39391594538|Travel & Living– Heading Title|3|False"
    Print #MyConfigFile, "|Assumptions|2|False"
    Print #MyConfigFile, "DAC_39391594540|Assumptions– Heading Title|3|False"

    Close MyConfigFile

End Function

Public Function Table_Report_Overide(ByRef rstrStrings() As String, ByVal vstrFileName As String, ByRef robjReport As Object) As Long
'Table_Report_Overide
On Error GoTo ErrorHandler

Dim i As Integer
Dim intRowCount As Integer


Call robjReport.CreateTable("Table_Report_Overide", 1, UBound(rstrStrings) + 1)
Call robjReport.SetColumnWidth("Table_Report_Overide", 0, 6)
For i = 0 To UBound(rstrStrings)
    Call robjReport.InsertData("Table_Report_Overide", 0, (i), rstrStrings(i), False, FontSize:=11, Alignment:=RTF_ALIGNLEFT, Bold:=False, vvntColor:=True)
Next
    
Call robjReport.GetTableText("Table_Report_Overide", vstrFileName)

CleanUp:
Exit Function

ErrorHandler:
    WriteLogFile Err, MODULE_NAME & ".Table_Report_Overide:" & Error, mstrPath, True
    Resume CleanUp
    Resume 0
End Function

Private Sub CreateGECEFolder(strPath As String)
On Error GoTo CreateGECEFolder_Error
Dim oFSO As New Scripting.FileSystemObject
Dim oFolder As Scripting.Folder

If Not oFSO.FolderExists(strPath) Then
    'create folder
   Set oFolder = oFSO.CreateFolder(strPath)
    If Not oFSO.FolderExists(oFolder.path) Then
        MsgBox "Could not create the folder: " & strPath, vbExclamation
    End If
    
End If

CleanUp:
On Error Resume Next
Set oFolder = Nothing
Set oFSO = Nothing

Exit Sub
CreateGECEFolder_Error:
    WriteLogFile Err, MODULE_NAME & ".CreateGECEFolder:" & Error, mstrPath
    Resume CleanUp
    Resume
End Sub
Public Function QPGetTempFile(ByVal vstrTempdir As String, ByVal vlngFileNum As Long) As String

On Error GoTo QPGetTempFile_Error
Dim strTempFileName As String

    QPGetTempFile = vstrTempdir & "~" & Format(vlngFileNum, "0000000") & ".tmp"

Exit Function
QPGetTempFile_Error:
    WriteLogFile Err, MODULE_NAME & ".QPGetTempFile:" & Error, mstrPath
    QPGetTempFile = ""
    Exit Function
    Resume
End Function
'rows and columns in the table start with 0 when stuffing data into them.
'Therefore when you initialze the table the number of columns/rows is always 1 minus the number of columns and rows
Public Function Table_ESA_Assumptions(ByRef robjRange As Excel.Range, ByVal vstrFileName As String, ByRef robjReport As Object) As Long
On Error GoTo ErrorHandler

Dim i As Integer
Dim intRowCount As Integer
For i = 1 To robjRange.Rows.Count
    If IsEmpty(robjRange(i, 1)) Then
        Exit For
    End If
Next
intRowCount = i - 1
If intRowCount > 0 Then
    Table_ESA_Assumptions = 0
Else
    Table_ESA_Assumptions = 1
    Exit Function
End If

Call robjReport.CreateTable("ESA_Assumptions", 1, intRowCount)
Call robjReport.SetColumnWidth("ESA_Assumptions", 0, 6)
For i = 1 To robjRange.Rows.Count
    If IsEmpty(robjRange(i, 1)) Then
        Exit For
    Else
        Call robjReport.InsertData("ESA_Assumptions", 0, (i - 1), robjRange(i, 1), False, FontSize:=11, Alignment:=RTF_ALIGNLEFT, Bold:=False)
    End If
Next
    
Call robjReport.GetTableText("ESA_Assumptions", vstrFileName)

CleanUp:
Exit Function

ErrorHandler:
    WriteLogFile Err, MODULE_NAME & ".Table_ESA_Assumptions:" & Error, mstrPath, True
    Resume CleanUp
End Function

'rows and columns in the table start with 0 when stuffing data into them.
'Therefore when you initialze the table the number of columns/rows is always 1 minus the number of columns and rows
Public Function Table_ESA_RevisionHistory(ByRef robjRange As Excel.Range, ByVal vstrFileName As String, ByRef robjReport As Object) As Long
On Error GoTo ErrorHandler
Dim i As Integer
Dim intRowCount As Integer
For i = 1 To robjRange.Rows.Count
    If IsEmpty(robjRange(i, 1)) Then
        Exit For
    End If
Next
intRowCount = i - 1
If intRowCount > 0 Then
    Table_ESA_RevisionHistory = 0
Else
    Table_ESA_RevisionHistory = 1
    Exit Function
End If

Call robjReport.CreateTable("ESA_RevisionHistory", 1, intRowCount)
Call robjReport.SetColumnWidth("ESA_RevisionHistory", 0, 6)
For i = 1 To robjRange.Rows.Count
    If IsEmpty(robjRange(i, 1)) Then
        Exit For
    Else
        Call robjReport.InsertData("ESA_RevisionHistory", 0, (i - 1), robjRange(i, 1), False, FontSize:=11, Alignment:=RTF_ALIGNLEFT, Bold:=False)
    End If
Next
    
Call robjReport.GetTableText("ESA_RevisionHistory", vstrFileName)
CleanUp:
Exit Function

ErrorHandler:
    WriteLogFile Err, MODULE_NAME & ".Table_ESA_RevisionHistory:" & Error, mstrPath, True
    Resume CleanUp
End Function

Public Sub WriteLogFile(ByVal vintErr&, ByRef rstrLogError$, ByVal vstrPath As String, Optional ByRef rvntDisplay As Variant)
'Must retrieve the error's source before resetting to new errorhandler
    Dim strSource$
    strSource = Err.Source

On Error Resume Next

Dim blnDisplay As Boolean
Dim intNextFile%, intTry%
Dim strFileName$, strDBErrors$, strError$, strLogErrors$, strTemp$
Dim strPath$

If vintErr = 999 Then Exit Sub

'Retrieve optional paramaters
    If IsMissing(rvntDisplay) Then
        blnDisplay = True
    Else
        blnDisplay = CInt(rvntDisplay)
    End If

'Reset screen paramaters
'    If blnDisplay Then
'        Screen.MousePointer = 0
'        LockWindowUpdate 0
'    End If

'Get the log filename and filenumber, make the Logs sub-directory
    strPath = FixPath(vstrPath)
'    'strFileName = strPath & Replace("GECE.exe", ".exe", "") & ".log"
'    strFileName = strPath & SaveReplace("GECE.exe", ".exe", "") & ".log"
    intNextFile = FreeFile

'Replace any User Defined Error message with predefined text and retrieve any database errors
    strLogErrors = rstrLogError
    strLogErrors = strLogErrors & vbCrLf & CLng(vintErr)
    strDBErrors = ""

'Log the error, If the log file is too big then kill it first
    Open strFileName For Append As intNextFile
        Print #intNextFile, Now & ":1" & "." & ".0" & ":" & Trim(Str(vintErr)) & ":";
        Print #intNextFile, strSource;
        Print #intNextFile, rstrLogError
        If strDBErrors <> "" And (Trim(strDBErrors) <> Trim(rstrLogError)) Then Print #intNextFile, strDBErrors
    Close intNextFile%

'Display the message if appropriate
    If blnDisplay Then
        Reply = MsgBox(strLogErrors, vbCritical Or vbSystemModal, "GECE Sheet")
    End If

End Sub




'*-------------
'*  This function will replace all occurences of a string
'*  by a replacement string.
'*  The replacement string may have a different length
'*  from the string to be replaced.
'*
'*  Input    : sS, Source string
'*             sR, String to replace
'*             sB, String to replace by
'*  Modifies : None
'*  Return   : Modified string
'*-------------
Function SaveReplace(ByVal sS As String, ByVal sR As String, ByVal sB As String) As String

  Dim iNdx As Integer, iR As Integer, iB As Integer
  Dim sLeft As String, sRight As String

  iR = Len(sR)                         'Length of string to be replaced
  iB = Len(sB)                         'Length of replacing string

  iNdx = InStr(sS, sR)                 'Find string to be replaced
  Do While iNdx <> 0                   'While more to be replaced

    sLeft = Left$(sS, iNdx - 1)        'Isolate left part
    sRight = mid$(sS, iNdx + iR)       'Isolate right part
    sS = sLeft + sB + sRight           'Replace by requested string
    iNdx = iNdx + iB                   'Skip over replaced string
    iNdx = InStr(iNdx, sS, sR)         'Find string to be replaced

  Loop

  SaveReplace = sS                      'Return replaced string

End Function

Function MyDocumentsFolder() As String
  Dim Pidl As LongPtr
  MyDocumentsFolder = Space$(260)
  SHGetSpecialFolderLocation 0, 5, Pidl
  SHGetPathFromIDList Pidl, MyDocumentsFolder
  MyDocumentsFolder = Left$(MyDocumentsFolder, _
    InStr(1, MyDocumentsFolder, vbNullChar) - 1)
End Function


