Attribute VB_Name = "modSched"
Option Explicit
Private Const LOG_ENABLED As Boolean = False   ' set True for dev only

'--- logging (disabled for users) ------------------------------------------------
Private Sub LogMsg(ByVal msg As String)
    If Not LOG_ENABLED Then Exit Sub
    On Error Resume Next
    Dim f As Integer: f = FreeFile
    Open ThisWorkbook.path & "\GECE_debug.log" For Append As #f
    Print #f, Format(Now, "yyyy-mm-dd hh:nn:ss"); " | "; msg
    Close #f
End Sub

'--- EXE-safe Named Range resolver (workbook, then sheet scope) ------------------
Public Function NR(ByVal nm As String, Optional ByVal ws As Worksheet) As Range
    Dim N As Name
    On Error Resume Next
    Set N = ThisWorkbook.Names(nm)
    If Not N Is Nothing Then
        If InStr(1, N.RefersTo, "#REF", vbTextCompare) = 0 Then Set NR = N.RefersToRange: Exit Function
    End If
    If Not ws Is Nothing Then
        Set N = ws.Names(nm)
        If Not N Is Nothing Then
            If InStr(1, N.RefersTo, "#REF", vbTextCompare) = 0 Then Set NR = N.RefersToRange: Exit Function
        End If
    End If
    On Error GoTo 0
End Function

'--- (DEV ONLY) wrapper with log+message; do not wire users to this --------------
Public Sub Run_Gantt_Debug()
    On Error GoTo ErrH
    LogMsg "=== RUN START ==="
    If Gantt_by_Phases Then
        LogMsg "SUCCESS | Gantt created"
        MsgBox "Gantt created.", vbInformation
    Else
        LogMsg "FAILED | Gantt_by_Phases returned False"
        MsgBox "Export did not complete. See GECE_debug.log.", vbExclamation
    End If
    Exit Sub
ErrH:
    LogMsg "ERR " & Err.Number & " at line " & Erl & " | " & Err.Description
    MsgBox "Err " & Err.Number & " at line " & Erl & ": " & Err.Description, vbCritical
End Sub

'--- PRODUCTION entry (wire the user button to THIS function) --------------------
Public Function Gantt_by_Phases() As Boolean
10  On Error GoTo ErrH
    Gantt_by_Phases = False
    If CheckForProject() = 0 Then Exit Function

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Schedule")
    LogMsg "Start Gantt_by_Phases"

    ' 0) Data present?
20  Dim init_hours As Double
    init_hours = ThisWorkbook.Worksheets(gstrGECECQAOutputSheet).Range("TOTAL_HOURS").Value
    If init_hours = 0 Then Exit Function   ' silent for users
    LogMsg "TOTAL_HOURS=" & init_hours

    ' 1) GoalSeek hardened ? fallback solver if native fails (silent)
30  Dim rFinal As Range, rReq As Range, rFactor As Range
    Set rFinal = NR("SCHED_FINAL_DURATION", ws)
    Set rReq = NR("SCHED_REQUIRED_DURATION", ws)
    Set rFactor = NR("SCHED_FACTOR", ws)

    If Not (rFinal Is Nothing Or rReq Is Nothing Or rFactor Is Nothing) Then
        If rFinal.CountLarge = 1 And rReq.CountLarge = 1 And rFactor.CountLarge = 1 And IsNumeric(rReq.Value) Then
            If Not (rFactor.Locked And ws.ProtectContents) Then
                On Error Resume Next
                rFinal.GoalSeek CDbl(rReq.Value), rFactor     ' positional args (late-binding safe)
                If Err.Number <> 0 Then
                    LogMsg "GoalSeek ERR " & Err.Number & " | " & Err.Description & " -> fallback"
                    Err.Clear
                    Call SolveFactorBinary(rFinal, rReq.Value, rFactor)
                End If
                On Error GoTo ErrH
            Else
                LogMsg "GoalSeek skipped (factor locked, protected sheet)"
            End If
        Else
            LogMsg "GoalSeek skipped (inputs invalid)"
        End If
    Else
        LogMsg "GoalSeek skipped (names missing)"
    End If
    If Not rFactor Is Nothing Then If val(rFactor.Value) < 0 Then Exit Function

    ' 2) Inputs
40  Dim rTasks As Range, rDur As Range, rPred As Range, rRes As Range
    Set rTasks = NR("SCHED_TASK", ws)
    Set rDur = NR("SCHED_DURATION", ws)
    Set rPred = NR("SCHED_PREDECESSORS", ws)
    Set rRes = NR("SCHED_RESOURCE_NAMES", ws)
    If rTasks Is Nothing Or rDur Is Nothing Or rPred Is Nothing Then Exit Function

60  Dim taskArr As Variant, durArr As Variant, predArr As Variant, resArr As Variant
    taskArr = rTasks.Value: durArr = rDur.Value: predArr = rPred.Value
    If Not rRes Is Nothing Then resArr = rRes.Value
    If Not IsArray(taskArr) Then Exit Function
    Dim nRows As Long: nRows = UBound(taskArr, 1)

    ' 3) Project/per-task starts
70  Dim projStart As Variant
    Dim rStart As Range: Set rStart = NR("SCHED_START_DATE", ws)
    projStart = IIf(rStart Is Nothing Or IsEmpty(rStart.Value), Date, rStart.Value)

80  Dim taskStarts() As Variant, i As Long
    ReDim taskStarts(1 To nRows, 1 To 1)
    taskStarts(1, 1) = projStart
    For i = 2 To nRows
        Dim nm As String: nm = "Start_Date_Task" & CStr(i)   ' legacy names on Schedule
        On Error Resume Next
        taskStarts(i, 1) = ws.Range(nm).Value
        On Error GoTo ErrH
    Next i

    ' 4) MS Project
90  Dim pj As Object, proj As Object
    On Error Resume Next
    Set pj = GetObject(, "MSProject.Application")
    If pj Is Nothing Then Set pj = CreateObject("MSProject.Application")
    On Error GoTo ErrH
    If pj Is Nothing Then Exit Function

    pj.Visible = True
    pj.DisplayAlerts = False
    Set proj = pj.Projects.Add
    If proj Is Nothing Then GoTo CleanUp
    If IsDate(projStart) Then proj.ProjectStart = CDate(projStart)

    ' set default to Auto if supported (no named args; ignore failures)
    On Error Resume Next
    pj.Application.NewTasksAreManual = False
    pj.Application.NewTasksAreManuallyScheduled = False
    On Error GoTo ErrH

    ' 5) PASS 1 — tasks, duration (hours?minutes), resources, starts (force Auto)
100 Dim minutesPerHour As Long: minutesPerHour = 60
    Dim t As Object, taskName As String, resStr As String, durH As Double
    Dim uidMap() As Long: ReDim uidMap(1 To nRows)

    For i = 1 To nRows
        taskName = Trim$(CStr(taskArr(i, 1)))
        If Len(taskName) > 0 Then
            Set t = proj.tasks.Add(taskName)
            If Not t Is Nothing Then
                On Error Resume Next
                t.TaskMode = 0      ' 0 = Auto (newer versions)
                t.Manual = False    ' older versions
                On Error GoTo ErrH

                uidMap(i) = t.uniqueID
                durH = val(durArr(i, 1))
                If durH > 0 Then t.Duration = CLng(durH * minutesPerHour)
                If IsArray(resArr) Then
                    resStr = Trim$(CStr(resArr(i, 1)))
                    If Len(resStr) > 0 Then t.ResourceNames = resStr
                End If
                If IsDate(taskStarts(i, 1)) Then t.Start = CDate(taskStarts(i, 1))
            End If
        End If
    Next i

    ' 6) PASS 2 — predecessors via UniqueID (drop self/invalid)
110 Dim predStr As String, norm As String
    For i = 1 To nRows
        If uidMap(i) <> 0 Then
            predStr = Trim$(CStr(predArr(i, 1)))
            If Len(predStr) > 0 Then
                norm = NormalizePredUID(predStr, i, uidMap)
                If Len(norm) > 0 Then
                    Set t = TaskByUID(proj, uidMap(i))
                    If Not t Is Nothing Then
                        On Error Resume Next
                        t.UniqueIDPredecessors = norm
                        Err.Clear
                        On Error GoTo ErrH
                    End If
                End If
            End If
        End If
    Next i

    ' 7) final sweep: ensure Auto & remove “?” (Estimated)
115 Dim tt As Object
    On Error Resume Next
    For Each tt In proj.tasks
        If Not tt Is Nothing Then
            tt.TaskMode = 0
            tt.Manual = False
            tt.Estimated = False
        End If
    Next tt
    On Error GoTo ErrH

    ' 8) basic cosmetics (safe positional args)
120 On Error Resume Next
    pj.Application.ViewApply "Gantt Chart"
    pj.Application.TableApply "Entry"
    pj.Application.TimescaleEdit 0, 2, 0, 10, True, True, True, True, 2
    On Error GoTo ErrH

    ' 9) save silently beside workbook
130 Dim savePath As String
    savePath = ThisWorkbook.path & "\" & Format(Date, "yyyymmdd") & "_Schedule_ProjectName_Rev01.mpp"
    On Error Resume Next
    proj.SaveAs savePath
    pj.DisplayAlerts = True
    On Error GoTo ErrH

    Gantt_by_Phases = True
CleanUp:
    Exit Function

ErrH:
    LogMsg "ERR " & Err.Number & " @ line " & Erl & " | " & Err.Description
    Resume CleanUp
End Function

'--- Helpers ----------------------------------------------------------------
Private Function SolveFactorBinary(ByVal rFinal As Range, ByVal Target As Double, ByVal rFactor As Range) As Boolean
    Dim lo As Double, hi As Double, mid As Double, f As Double, i As Long
    Dim cur As Double: cur = val(rFactor.Value)
    lo = IIf(cur > 0, 0, cur - 10): hi = IIf(cur > 0, cur * 2 + 10, 10)
    If hi <= lo Then hi = lo + 10

    rFactor.Value = lo: Application.Calculate
    Dim fLo As Double: fLo = val(rFinal.Value)
    rFactor.Value = hi: Application.Calculate
    Dim fHi As Double: fHi = val(rFinal.Value)
    If Abs(fLo - fHi) < 0.0000001 Then SolveFactorBinary = False: Exit Function

    For i = 1 To 40
        mid = (lo + hi) / 2
        rFactor.Value = mid: Application.Calculate
        f = val(rFinal.Value)
        If (f < Target) = (fLo < fHi) Then
            lo = mid: fLo = f
        Else
            hi = mid: fHi = f
        End If
        If Abs(f - Target) <= 0.0001 Then Exit For
    Next i
    SolveFactorBinary = True
End Function

Private Function NormalizePredUID(ByVal s As String, ByVal curRow As Long, ByRef uidMap() As Long) As String
    Dim parts() As String, p As String, out As String, num As Long, suffix As String, mapped As Long, i As Long
    s = Trim$(Replace(s, ";", ",")): If Len(s) = 0 Then Exit Function
    parts = Split(s, ",")
    For i = LBound(parts) To UBound(parts)
        p = Trim$(parts(i)): If Len(p) = 0 Then GoTo nxt
        num = val(p): suffix = mid$(p, Len(CStr(num)) + 1)
        If num > 0 And num <> curRow And num <= UBound(uidMap) Then
            mapped = uidMap(num)
            If mapped > 0 Then
                If Len(out) > 0 Then out = out & ","
                out = out & CStr(mapped) & suffix
            End If
        End If
nxt: Next i
    NormalizePredUID = out
End Function

Private Function TaskByUID(ByVal proj As Object, ByVal uid As Long) As Object
    Dim t As Object
    For Each t In proj.tasks
        If Not t Is Nothing Then If t.uniqueID = uid Then Set TaskByUID = t: Exit Function
    Next t
End Function

'--- MS Project presence check ----------------------------------------------
Public Function CheckForProject() As Integer
    On Error GoTo ErrH
    Dim p As Object
    On Error Resume Next
    Set p = GetObject(, "MSProject.Application")
    If p Is Nothing Then Set p = CreateObject("MSProject.Application")
    On Error GoTo ErrH
    CheckForProject = IIf(p Is Nothing, 0, 1)
CleanUp:
    On Error Resume Next
    Set p = Nothing
    Exit Function
ErrH:
    CheckForProject = 0
    Resume CleanUp
End Function


' ===== (Your existing function, kept; already guarded & DoneEx-safe) =====
Public Function CashInflowOutflow() As Boolean
    On Error GoTo ErrHandler
    CashInflowOutflow = False

    Dim currentDate As Date, proposalDate As Date, projectstartDate As Date, projectfinishDate As Date
    Dim proposalDate_tmp As Date
    Dim startDate As Date, finishDate As Date
    Dim ii As Long, jj As Long, kk As Long, escalation_increment As Long
    Dim WritecellRow As Long, WritecellColumn As Long
    Dim HourcellRow As Long, HourcellColumn As Long
    Dim StartDatecellRow As Long, StartDatecellColumn As Long
    Dim FinishDatecellRow As Long, FinishDatecellColumn As Long
    Dim RessourcecellRow As Long, RessourcecellColumn As Long
    Dim RemotePCTcellRow As Long, RemotePCTcellColumn As Long
    Dim DurationcellRow As Long, DurationcellColumn As Long

    Dim oldCalc As XlCalculation, oldEv As Boolean, oldUpd As Boolean
    oldCalc = Application.Calculation
    oldEv = Application.EnableEvents
    oldUpd = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    If gstrGECEWorkBook = "" Then getWorkBookName

    With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEScheduleSheet)
        WritecellRow = .Range("SCHED_LOCAL_START").row
        WritecellColumn = .Range("SCHED_LOCAL_START").Column
        HourcellRow = .Range("SCHED_HOUR_START").row
        HourcellColumn = .Range("SCHED_HOUR_START").Column
        StartDatecellRow = .Range("SCHED_START_DATE").row
        StartDatecellColumn = .Range("SCHED_START_DATE").Column
        FinishDatecellRow = .Range("SCHED_FINISH_DATE_START").row
        FinishDatecellColumn = .Range("SCHED_FINISH_DATE_START").Column
        RessourcecellRow = .Range("SCHED_RESSOURCES_START").row
        RessourcecellColumn = .Range("SCHED_RESSOURCES_START").Column
        RemotePCTcellRow = .Range("SCHED_REM_PCT_START").row
        RemotePCTcellColumn = .Range("SCHED_REM_PCT_START").Column
        DurationcellRow = .Range("SCHED_DURATION_START").row
        DurationcellColumn = .Range("SCHED_DURATION_START").Column

        .Range("I320:CZ320").ClearContents
        .Range("I323:CZ329").ClearContents
        .Range("I333:CZ339").ClearContents

        projectstartDate = .Cells(StartDatecellRow + 2, StartDatecellColumn).Value
        projectfinishDate = .Cells(FinishDatecellRow + 2, FinishDatecellColumn).Value

        proposalDate_tmp = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("PROPSAL_DATE").Value
        If proposalDate_tmp = "00:00:00" Then
            proposalDate = Now()
        Else
            proposalDate = proposalDate_tmp
        End If

        If Month(proposalDate) - 4 > 0 And Month(projectstartDate) - 4 < 0 Then
            escalation_increment = escalation_increment - 1
        ElseIf Month(proposalDate) - 4 < 0 And Month(projectstartDate) - 4 > 0 Then
            escalation_increment = escalation_increment + 1
        End If
        escalation_increment = escalation_increment + Year(projectstartDate) - Year(proposalDate)

        For kk = 0 To 95
            currentDate = .Cells(.Range("SCHED_CALENDAR_START").row, .Range("SCHED_CALENDAR_START").Column + kk).Value

            If Month(currentDate) = 4 Then
                escalation_increment = escalation_increment + 1
            End If
            .Cells(WritecellRow - 3, WritecellColumn + kk).Value = escalation_increment

            For ii = 0 To 50
                If .Cells(HourcellRow + ii, HourcellColumn).Value <> 0 Then
                    startDate = .Cells(StartDatecellRow + ii, StartDatecellColumn).Value
                    finishDate = .Cells(FinishDatecellRow + ii, FinishDatecellColumn).Value

                    If (Month(currentDate) - Month(startDate) + (Year(currentDate) - Year(startDate)) * 12 >= 0) And _
                       (Month(finishDate) - Month(currentDate) + (Year(finishDate) - Year(currentDate)) * 12 >= 0) Then

                        For jj = 0 To 6
                            If .Cells(RessourcecellRow + ii, RessourcecellColumn + jj).Value <> 0 Then
                                .Cells(WritecellRow + jj, WritecellColumn + kk).Value = _
                                    .Cells(WritecellRow + jj, WritecellColumn + kk).Value + _
                                    (.Cells(RessourcecellRow + ii, RessourcecellColumn + jj).Value * (1 - .Cells(RemotePCTcellRow + ii, RemotePCTcellColumn).Value))
                            End If
                        Next jj

                        For jj = 0 To 6
                            If .Cells(RessourcecellRow + ii, RessourcecellColumn + jj).Value <> 0 Then
                                .Cells(WritecellRow + 10 + jj, WritecellColumn + kk).Value = _
                                    .Cells(WritecellRow + 10 + jj, WritecellColumn + kk).Value + _
                                    (.Cells(RessourcecellRow + ii, RessourcecellColumn + jj).Value * .Cells(RemotePCTcellRow + ii, RemotePCTcellColumn).Value)
                            End If
                        Next jj
                    End If
                End If
            Next ii
        Next kk
    End With

    CashInflowOutflow = True
CleanUp:
    With Application
        .Calculation = oldCalc
        .EnableEvents = oldEv
        .ScreenUpdating = oldUpd
    End With
    Exit Function

ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source, vbCritical
    Resume CleanUp
End Function

' ===== Stubs (unchanged) =====
Public Function ShowRessourcePlanningForm()
End Function

Public Function ShowForm()
    ReplaceCellName.Show vbModal
End Function
