Attribute VB_Name = "modGeceExport"
#If VBA7 Then
    Private Declare PtrSafe Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" ( _
        ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As LongPtr

    Private Declare PtrSafe Function FindClose Lib "kernel32" (ByVal hFindFile As LongPtr) As Long
    Private Declare PtrSafe Function GetTempPath Lib "kernel32" Alias "GetTempPathA" ( _
        ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

#End If


Private Const TEMP_FOLDER As String = "C:\Temp\"
Private Const LOG_FILE As String = "ExportLog.txt"
Private Const MAX_RETRIES As Integer = 3
Private Const WAIT_TIME As Integer = 1000

Private Const MAX_PATH As Long = 260
Private Const ERROR_SUCCESS As Long = 0
Private Const INVALID_HANDLE_VALUE As Long = -1

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternateFileName As String * 14
End Type

Private Sub LogError(ByVal procedureName As String, ByVal errorNumber As Long, ByVal errorDescription As String)
    On Error Resume Next
    Dim logFile As String, fileNum As Integer, logMessage As String
    logFile = GetLogFilePath()
    fileNum = FreeFile
    logMessage = Format(Now, "yyyy-mm-dd hh:mm:ss") & " - " & procedureName & " - Error " & errorNumber & ": " & errorDescription
    Open logFile For Append As #fileNum
    Print #fileNum, logMessage
    Close #fileNum
End Sub

Private Function GetLogFilePath() As String
    On Error GoTo ErrorHandler
    Dim tempPath As String
    tempPath = GetTempFolder()
    If Right(tempPath, 1) <> "\" Then tempPath = tempPath & "\"
    GetLogFilePath = tempPath & LOG_FILE
    Exit Function
ErrorHandler:
    GetLogFilePath = TEMP_FOLDER & LOG_FILE
End Function

Private Function GetTempFolder() As String
    On Error GoTo ErrorHandler
    Dim buffer As String * 512
    Dim length As Long
    length = GetTempPath(Len(buffer), buffer)
    If length > 0 Then
        GetTempFolder = Left(buffer, length)
    Else
        GetTempFolder = TEMP_FOLDER
    End If
    Exit Function
ErrorHandler:
    GetTempFolder = TEMP_FOLDER
End Function

Public Function ExportGECEData(ByVal ws As Worksheet, Optional ByVal exportPath As String = "") As Boolean
    On Error GoTo ErrorHandler
    Dim exportContent As String, retryCount As Integer, success As Boolean
    If ws Is Nothing Then
        LogError "ExportGECEData", vbObjectError + 1, "Invalid worksheet"
        Exit Function
    End If
    exportContent = GenerateExportContent(ws)
    If Len(exportContent) = 0 Then
        LogError "ExportGECEData", vbObjectError + 2, "No content to export"
        Exit Function
    End If
    If Len(Trim(exportPath)) = 0 Then
        exportPath = GetExportFilePath(ws.Name)
        If Len(exportPath) = 0 Then Exit Function
    End If
    Do While retryCount < MAX_RETRIES And Not success
        success = SaveExportFile(exportPath, exportContent)
        If Not success Then
            retryCount = retryCount + 1
            Sleep WAIT_TIME
        End If
    Loop
    ExportGECEData = success
    Exit Function
ErrorHandler:
    LogError "ExportGECEData", Err.Number, Err.Description
    ExportGECEData = False
End Function

Private Function GenerateExportContent(ByVal ws As Worksheet) As String
    On Error GoTo ErrorHandler
    Dim usedRange As Range, cell As Range, row As Range
    Dim exportText As String, rowText As String
    Set usedRange = ws.usedRange
    For Each row In usedRange.Rows
        rowText = ""
        For Each cell In row.Cells
            rowText = rowText & ProcessCellContent(cell) & vbTab
        Next cell
        exportText = exportText & Left(rowText, Len(rowText) - 1) & vbCrLf
    Next row
    GenerateExportContent = exportText
    Exit Function
ErrorHandler:
    LogError "GenerateExportContent", Err.Number, Err.Description
    GenerateExportContent = ""
End Function

Private Function ProcessCellContent(ByVal cell As Range) As String
    On Error GoTo ErrorHandler
    Dim cellValue As String
    Select Case cell.NumberFormat
        Case "@": cellValue = cell.Text
        Case "General"
            If IsError(cell.Value) Then
                cellValue = ""
            Else
                cellValue = CStr(cell.Value)
            End If
        Case Else
            cellValue = cell.Text
    End Select
    cellValue = Replace(cellValue, vbCr, "")
    cellValue = Replace(cellValue, vbLf, "")
    cellValue = Replace(cellValue, vbTab, " ")
    ProcessCellContent = cellValue
    Exit Function
ErrorHandler:
    ProcessCellContent = ""
End Function

Private Function GetExportFilePath(ByVal wsName As String) As String
    On Error GoTo ErrorHandler
    Dim defaultName As String, result As Variant
    defaultName = wsName & "_Export_" & Format(Now, "yyyymmdd_hhmmss") & ".txt"
    result = Application.GetSaveAsFilename( _
        InitialFileName:=defaultName, _
        FileFilter:="Text Files (*.txt), *.txt", _
        Title:="Select Export Location")
    If result <> False Then
        GetExportFilePath = result
    End If
    Exit Function
ErrorHandler:
    LogError "GetExportFilePath", Err.Number, Err.Description
    GetExportFilePath = ""
End Function

Private Function SaveExportFile(ByVal filePath As String, ByVal content As String) As Boolean
    On Error GoTo ErrorHandler
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, content
    Close #fileNum
    SaveExportFile = True
    Exit Function
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    LogError "SaveExportFile", Err.Number, Err.Description
    SaveExportFile = False
End Function

Private Function FileExists(ByVal filePath As String) As Boolean
    On Error GoTo ErrorHandler
    Dim findData As WIN32_FIND_DATA
    Dim handle As LongPtr
    handle = FindFirstFile(filePath, findData)
    FileExists = (handle <> INVALID_HANDLE_VALUE)
    If handle <> INVALID_HANDLE_VALUE Then
        FindClose handle
    End If
    Exit Function
ErrorHandler:
    FileExists = False
End Function

Public Sub ExportWorksheet()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveSheet

    MsgBox "ExportWorksheet triggered", vbInformation ' Debug message

    If ExportGECEData(ws) Then
        MsgBox "Export completed successfully!", vbInformation
    Else
        MsgBox "Export failed. Check the log file for details.", vbCritical
    End If
    Exit Sub

ErrorHandler:
    LogError "ExportWorksheet", Err.Number, Err.Description
End Sub

