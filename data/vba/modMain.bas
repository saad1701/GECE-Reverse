Attribute VB_Name = "modMain"
Option Explicit

'constants
Public Const GECEXLSVERSION As String = "1.0"
Public gstrGECEWorkBook As String
Public OpenDataEntryForm As Boolean
Public Const gstrGECEDataEntrySheet = "Data Entry"
Public Const gstrGECECurrencySheet = "World Currency Table"
Public Const gstrGECECostSheet = "World Cost tables"
Public Const gstrGECEApplicationBasedSheet = "Application Based"
Public Const gstrGECEPriceMakeUpSheet = "Price Make-up"
Public Const gstrGECEAssumptionsProposalSheet = "Assumptions - Proposal infos"
Public Const gstrGECEProposalSummarySheet = "Proposal Summary"
Public Const gstrGECETaskBasedSheet = "Task Based"
Public Const gstrGECEDurationBasedSheet = "Duration Based"
Public Const gstrGECEIndustrySheet = "Industry"
Public Const gstrGECEToolkitSheet = "Toolkit"
Public Const gstrGECEScheduleSheet = "Schedule"
Public Const gstrGECEWPASheet = "WPA"

Public Const gstrGECECQAOutputSheet = "GECE TO CQA Output"
Public Const gstrGECETotalsSheet = "Totals"
Public Const gstrGECEExportToERPSheet = "ExportToERP"



'Added by AB 23/01/2006 for task based estimator
'Modified by JFR 22/10/2007 to add project phases by tracking activities

Public Function TaskBasedSummaryUpdate() As Boolean

On Error GoTo ErrHandler
'gets the name of the opened GECE sheet workbook
Dim activity As String
Dim i As Integer
Dim tmp
Dim wb As Workbooks
Dim s As Worksheet

'clear variable
' ***** need to initialise the variables in the summary *****
activity = ""
i = 0
Application.Cursor = xlWait

With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECETaskBasedSheet)
    

    ' Specifications
    .Range("TASK_BASED_SPEC_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_SPEC_1st_UNIT_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_SPEC_2nd_UNIT_TOTAL_HOURS").Value = 0
    
    .Range("TASK_BASED_SPEC_REM_HOURS").Value = 0
    .Range("TASK_BASED_SPEC_1st_UNIT_REM_HOURS").Value = 0
    .Range("TASK_BASED_SPEC_2nd_UNIT_REM_HOURS").Value = 0
    
    .Range("TASK_BASED_SPEC_TOTAL_COST").Value = 0
    .Range("TASK_BASED_SPEC_1st_UNIT_TOTAL_COST").Value = 0
    .Range("TASK_BASED_SPEC_2nd_UNIT_TOTAL_COST").Value = 0
    
    .Range("TASK_BASED_SPEC_LOC_COST").Value = 0
    .Range("TASK_BASED_SPEC_1st_UNIT_LOC_COST").Value = 0
    .Range("TASK_BASED_SPEC_2nd_UNIT_LOC_COST").Value = 0
    
    .Range("TASK_BASED_SPEC_LOC_AC_HOURS").Value = 0
    .Range("TASK_BASED_SPEC_LOC_LE_HOURS").Value = 0
    .Range("TASK_BASED_SPEC_LOC_AE_HOURS").Value = 0
    .Range("TASK_BASED_SPEC_LOC_T_HOURS").Value = 0
    .Range("TASK_BASED_SPEC_LOC_SE_HOURS").Value = 0
    .Range("TASK_BASED_SPEC_LOC_SR_HOURS").Value = 0

    .Range("TASK_BASED_SPEC_REM_AC_HOURS").Value = 0
    .Range("TASK_BASED_SPEC_REM_LE_HOURS").Value = 0
    .Range("TASK_BASED_SPEC_REM_AE_HOURS").Value = 0
    .Range("TASK_BASED_SPEC_REM_T_HOURS").Value = 0
    .Range("TASK_BASED_SPEC_REM_SE_HOURS").Value = 0
    .Range("TASK_BASED_SPEC_REM_SR_HOURS").Value = 0

    .Range("TASK_BASED_SPEC_REVIEW").Value = 0
    .Range("TASK_BASED_SPEC_DESIGN").Value = 0
    .Range("TASK_BASED_SPEC_IMPLEMENT").Value = 0
    .Range("TASK_BASED_SPEC_TEST").Value = 0
    .Range("TASK_BASED_SPEC_SITE").Value = 0
    
    ' System Engineeering
    .Range("TASK_BASED_SYS_ENG_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_SYS_ENG_1st_UNIT_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_SYS_ENG_2nd_UNIT_TOTAL_HOURS").Value = 0
    
    .Range("TASK_BASED_SYS_ENG_REM_HOURS").Value = 0
    .Range("TASK_BASED_SYS_ENG_1st_UNIT_REM_HOURS").Value = 0
    .Range("TASK_BASED_SYS_ENG_2nd_UNIT_REM_HOURS").Value = 0
    
    .Range("TASK_BASED_SYS_ENG_TOTAL_COST").Value = 0
    .Range("TASK_BASED_SYS_ENG_1st_UNIT_TOTAL_COST").Value = 0
    .Range("TASK_BASED_SYS_ENG_2nd_UNIT_TOTAL_COST").Value = 0
    
    .Range("TASK_BASED_SYS_ENG_LOC_COST").Value = 0
    .Range("TASK_BASED_SYS_ENG_1st_UNIT_LOC_COST").Value = 0
    .Range("TASK_BASED_SYS_ENG_2nd_UNIT_LOC_COST").Value = 0
    
    .Range("TASK_BASED_SYS_ENG_LOC_AC_HOURS").Value = 0
    .Range("TASK_BASED_SYS_ENG_LOC_LE_HOURS").Value = 0
    .Range("TASK_BASED_SYS_ENG_LOC_AE_HOURS").Value = 0
    .Range("TASK_BASED_SYS_ENG_LOC_T_HOURS").Value = 0
    .Range("TASK_BASED_SYS_ENG_LOC_SE_HOURS").Value = 0
    .Range("TASK_BASED_SYS_ENG_LOC_SR_HOURS").Value = 0
        
    .Range("TASK_BASED_SYS_ENG_REM_AC_HOURS").Value = 0
    .Range("TASK_BASED_SYS_ENG_REM_LE_HOURS").Value = 0
    .Range("TASK_BASED_SYS_ENG_REM_AE_HOURS").Value = 0
    .Range("TASK_BASED_SYS_ENG_REM_T_HOURS").Value = 0
    .Range("TASK_BASED_SYS_ENG_REM_SE_HOURS").Value = 0
    .Range("TASK_BASED_SYS_ENG_REM_SR_HOURS").Value = 0

    .Range("TASK_BASED_SYSENG_REVIEW").Value = 0
    .Range("TASK_BASED_SYSENG_DESIGN").Value = 0
    .Range("TASK_BASED_SYSENG_IMPLEMENT").Value = 0
    .Range("TASK_BASED_SYSENG_TEST").Value = 0
    .Range("TASK_BASED_SYSENG_SITE").Value = 0
    
    
    ' HMI
    .Range("TASK_BASED_HMI_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_HMI_1st_UNIT_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_HMI_2nd_UNIT_TOTAL_HOURS").Value = 0
    
    .Range("TASK_BASED_HMI_REM_HOURS").Value = 0
    .Range("TASK_BASED_HMI_1st_UNIT_REM_HOURS").Value = 0
    .Range("TASK_BASED_HMI_2nd_UNIT_REM_HOURS").Value = 0
    
    .Range("TASK_BASED_HMI_TOTAL_COST").Value = 0
    .Range("TASK_BASED_HMI_1st_UNIT_TOTAL_COST").Value = 0
    .Range("TASK_BASED_HMI_2nd_UNIT_TOTAL_COST").Value = 0
    
    .Range("TASK_BASED_HMI_LOC_COST").Value = 0
    .Range("TASK_BASED_HMI_1st_UNIT_LOC_COST").Value = 0
    .Range("TASK_BASED_HMI_2nd_UNIT_LOC_COST").Value = 0
    
    .Range("TASK_BASED_HMI_LOC_AC_HOURS").Value = 0
    .Range("TASK_BASED_HMI_LOC_LE_HOURS").Value = 0
    .Range("TASK_BASED_HMI_LOC_AE_HOURS").Value = 0
    .Range("TASK_BASED_HMI_LOC_T_HOURS").Value = 0
    .Range("TASK_BASED_HMI_LOC_SE_HOURS").Value = 0
    .Range("TASK_BASED_HMI_LOC_SR_HOURS").Value = 0
        
    .Range("TASK_BASED_HMI_REM_AC_HOURS").Value = 0
    .Range("TASK_BASED_HMI_REM_LE_HOURS").Value = 0
    .Range("TASK_BASED_HMI_REM_AE_HOURS").Value = 0
    .Range("TASK_BASED_HMI_REM_T_HOURS").Value = 0
    .Range("TASK_BASED_HMI_REM_SE_HOURS").Value = 0
    .Range("TASK_BASED_HMI_REM_SR_HOURS").Value = 0
    
    .Range("TASK_BASED_HMI_REVIEW").Value = 0
    .Range("TASK_BASED_HMI_DESIGN").Value = 0
    .Range("TASK_BASED_HMI_IMPLEMENT").Value = 0
    .Range("TASK_BASED_HMI_TEST").Value = 0
    .Range("TASK_BASED_HMI_SITE").Value = 0
    
        ' Configuration / Programming
    .Range("TASK_BASED_CP_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_CP_1st_UNIT_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_CP_2nd_UNIT_TOTAL_HOURS").Value = 0
    
    .Range("TASK_BASED_CP_REM_HOURS").Value = 0
    .Range("TASK_BASED_CP_1st_UNIT_REM_HOURS").Value = 0
    .Range("TASK_BASED_CP_2nd_UNIT_REM_HOURS").Value = 0
    
    .Range("TASK_BASED_CP_TOTAL_COST").Value = 0
    .Range("TASK_BASED_CP_1st_UNIT_TOTAL_COST").Value = 0
    .Range("TASK_BASED_CP_2nd_UNIT_TOTAL_COST").Value = 0
    
    .Range("TASK_BASED_CP_LOC_COST").Value = 0
    .Range("TASK_BASED_CP_1st_UNIT_LOC_COST").Value = 0
    .Range("TASK_BASED_CP_2nd_UNIT_LOC_COST").Value = 0
     
    .Range("TASK_BASED_CP_LOC_AC_HOURS").Value = 0
    .Range("TASK_BASED_CP_LOC_LE_HOURS").Value = 0
    .Range("TASK_BASED_CP_LOC_AE_HOURS").Value = 0
    .Range("TASK_BASED_CP_LOC_T_HOURS").Value = 0
    .Range("TASK_BASED_CP_LOC_SE_HOURS").Value = 0
    .Range("TASK_BASED_CP_LOC_SR_HOURS").Value = 0
        
    .Range("TASK_BASED_CP_REM_AC_HOURS").Value = 0
    .Range("TASK_BASED_CP_REM_LE_HOURS").Value = 0
    .Range("TASK_BASED_CP_REM_AE_HOURS").Value = 0
    .Range("TASK_BASED_CP_REM_T_HOURS").Value = 0
    .Range("TASK_BASED_CP_REM_SE_HOURS").Value = 0
    .Range("TASK_BASED_CP_REM_SR_HOURS").Value = 0
    
    .Range("TASK_BASED_CP_REVIEW").Value = 0
    .Range("TASK_BASED_CP_DESIGN").Value = 0
    .Range("TASK_BASED_CP_IMPLEMENT").Value = 0
    .Range("TASK_BASED_CP_TEST").Value = 0
    .Range("TASK_BASED_CP_SITE").Value = 0
    
    
    ' Data Interface
    .Range("TASK_BASED_DI_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_DI_1st_UNIT_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_DI_2nd_UNIT_TOTAL_HOURS").Value = 0
    
    .Range("TASK_BASED_DI_REM_HOURS").Value = 0
    .Range("TASK_BASED_DI_1st_UNIT_REM_HOURS").Value = 0
    .Range("TASK_BASED_DI_2nd_UNIT_REM_HOURS").Value = 0
    
    .Range("TASK_BASED_DI_TOTAL_COST").Value = 0
    .Range("TASK_BASED_DI_1st_UNIT_TOTAL_COST").Value = 0
    .Range("TASK_BASED_DI_2nd_UNIT_TOTAL_COST").Value = 0
    
    .Range("TASK_BASED_DI_LOC_COST").Value = 0
    .Range("TASK_BASED_DI_1st_UNIT_LOC_COST").Value = 0
    .Range("TASK_BASED_DI_2nd_UNIT_LOC_COST").Value = 0
    
    .Range("TASK_BASED_DI_LOC_AC_HOURS").Value = 0
    .Range("TASK_BASED_DI_LOC_LE_HOURS").Value = 0
    .Range("TASK_BASED_DI_LOC_AE_HOURS").Value = 0
    .Range("TASK_BASED_DI_LOC_T_HOURS").Value = 0
    .Range("TASK_BASED_DI_LOC_SE_HOURS").Value = 0
    .Range("TASK_BASED_DI_LOC_SR_HOURS").Value = 0
        
    .Range("TASK_BASED_DI_REM_AC_HOURS").Value = 0
    .Range("TASK_BASED_DI_REM_LE_HOURS").Value = 0
    .Range("TASK_BASED_DI_REM_AE_HOURS").Value = 0
    .Range("TASK_BASED_DI_REM_T_HOURS").Value = 0
    .Range("TASK_BASED_DI_REM_SE_HOURS").Value = 0
    .Range("TASK_BASED_DI_REM_SR_HOURS").Value = 0
    
    .Range("TASK_BASED_DI_REVIEW").Value = 0
    .Range("TASK_BASED_DI_DESIGN").Value = 0
    .Range("TASK_BASED_DI_IMPLEMENT").Value = 0
    .Range("TASK_BASED_DI_TEST").Value = 0
    .Range("TASK_BASED_DI_SITE").Value = 0
    
    
    ' Reports
    .Range("TASK_BASED_REP_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_REP_1st_UNIT_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_REP_2nd_UNIT_TOTAL_HOURS").Value = 0
    
    .Range("TASK_BASED_REP_REM_HOURS").Value = 0
    .Range("TASK_BASED_REP_1st_UNIT_REM_HOURS").Value = 0
    .Range("TASK_BASED_REP_2nd_UNIT_REM_HOURS").Value = 0
    
    .Range("TASK_BASED_REP_TOTAL_COST").Value = 0
    .Range("TASK_BASED_REP_1st_UNIT_TOTAL_COST").Value = 0
    .Range("TASK_BASED_REP_2nd_UNIT_TOTAL_COST").Value = 0
    
    .Range("TASK_BASED_REP_LOC_COST").Value = 0
    .Range("TASK_BASED_REP_1st_UNIT_LOC_COST").Value = 0
    .Range("TASK_BASED_REP_2nd_UNIT_LOC_COST").Value = 0
    
    .Range("TASK_BASED_REP_LOC_AC_HOURS").Value = 0
    .Range("TASK_BASED_REP_LOC_LE_HOURS").Value = 0
    .Range("TASK_BASED_REP_LOC_AE_HOURS").Value = 0
    .Range("TASK_BASED_REP_LOC_T_HOURS").Value = 0
    .Range("TASK_BASED_REP_LOC_SE_HOURS").Value = 0
    .Range("TASK_BASED_REP_LOC_SR_HOURS").Value = 0
    
    .Range("TASK_BASED_REP_REM_AC_HOURS").Value = 0
    .Range("TASK_BASED_REP_REM_LE_HOURS").Value = 0
    .Range("TASK_BASED_REP_REM_AE_HOURS").Value = 0
    .Range("TASK_BASED_REP_REM_T_HOURS").Value = 0
    .Range("TASK_BASED_REP_REM_SE_HOURS").Value = 0
    .Range("TASK_BASED_REP_REM_SR_HOURS").Value = 0

    .Range("TASK_BASED_REP_REVIEW").Value = 0
    .Range("TASK_BASED_REP_DESIGN").Value = 0
    .Range("TASK_BASED_REP_IMPLEMENT").Value = 0
    .Range("TASK_BASED_REP_TEST").Value = 0
    .Range("TASK_BASED_REP_SITE").Value = 0

    ' Applications
    .Range("TASK_BASED_APP_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_APP_1st_UNIT_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_APP_2nd_UNIT_TOTAL_HOURS").Value = 0
    
    .Range("TASK_BASED_APP_REM_HOURS").Value = 0
    .Range("TASK_BASED_APP_1st_UNIT_REM_HOURS").Value = 0
    .Range("TASK_BASED_APP_2nd_UNIT_REM_HOURS").Value = 0
    
    .Range("TASK_BASED_APP_TOTAL_COST").Value = 0
    .Range("TASK_BASED_APP_1st_UNIT_TOTAL_COST").Value = 0
    .Range("TASK_BASED_APP_2nd_UNIT_TOTAL_COST").Value = 0
    
    .Range("TASK_BASED_APP_LOC_COST").Value = 0
    .Range("TASK_BASED_APP_1st_UNIT_LOC_COST").Value = 0
    .Range("TASK_BASED_APP_2nd_UNIT_LOC_COST").Value = 0
    
    .Range("TASK_BASED_APP_LOC_AC_HOURS").Value = 0
    .Range("TASK_BASED_APP_LOC_LE_HOURS").Value = 0
    .Range("TASK_BASED_APP_LOC_AE_HOURS").Value = 0
    .Range("TASK_BASED_APP_LOC_T_HOURS").Value = 0
    .Range("TASK_BASED_APP_LOC_SE_HOURS").Value = 0
    .Range("TASK_BASED_APP_LOC_SR_HOURS").Value = 0
        
    .Range("TASK_BASED_APP_REM_AC_HOURS").Value = 0
    .Range("TASK_BASED_APP_REM_LE_HOURS").Value = 0
    .Range("TASK_BASED_APP_REM_AE_HOURS").Value = 0
    .Range("TASK_BASED_APP_REM_T_HOURS").Value = 0
    .Range("TASK_BASED_APP_REM_SE_HOURS").Value = 0
    .Range("TASK_BASED_APP_REM_SR_HOURS").Value = 0
    
    .Range("TASK_BASED_APP_REVIEW").Value = 0
    .Range("TASK_BASED_APP_DESIGN").Value = 0
    .Range("TASK_BASED_APP_IMPLEMENT").Value = 0
    .Range("TASK_BASED_APP_TEST").Value = 0
    .Range("TASK_BASED_APP_SITE").Value = 0
    
    
    ' Testing
    .Range("TASK_BASED_TEST_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_TEST_1st_UNIT_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_TEST_2nd_UNIT_TOTAL_HOURS").Value = 0
    
    .Range("TASK_BASED_TEST_REM_HOURS").Value = 0
    .Range("TASK_BASED_TEST_1st_UNIT_REM_HOURS").Value = 0
    .Range("TASK_BASED_TEST_2nd_UNIT_REM_HOURS").Value = 0
    
    .Range("TASK_BASED_TEST_TOTAL_COST").Value = 0
    .Range("TASK_BASED_TEST_1st_UNIT_TOTAL_COST").Value = 0
    .Range("TASK_BASED_TEST_2nd_UNIT_TOTAL_COST").Value = 0
    
    .Range("TASK_BASED_TEST_LOC_COST").Value = 0
    .Range("TASK_BASED_TEST_1st_UNIT_LOC_COST").Value = 0
    .Range("TASK_BASED_TEST_2nd_UNIT_LOC_COST").Value = 0
    
    .Range("TASK_BASED_TEST_LOC_AC_HOURS").Value = 0
    .Range("TASK_BASED_TEST_LOC_LE_HOURS").Value = 0
    .Range("TASK_BASED_TEST_LOC_AE_HOURS").Value = 0
    .Range("TASK_BASED_TEST_LOC_T_HOURS").Value = 0
    .Range("TASK_BASED_TEST_LOC_SE_HOURS").Value = 0
    .Range("TASK_BASED_TEST_LOC_SR_HOURS").Value = 0
        
    .Range("TASK_BASED_TEST_REM_AC_HOURS").Value = 0
    .Range("TASK_BASED_TEST_REM_LE_HOURS").Value = 0
    .Range("TASK_BASED_TEST_REM_AE_HOURS").Value = 0
    .Range("TASK_BASED_TEST_REM_T_HOURS").Value = 0
    .Range("TASK_BASED_TEST_REM_SE_HOURS").Value = 0
    .Range("TASK_BASED_TEST_REM_SR_HOURS").Value = 0
    
    .Range("TASK_BASED_TEST_REVIEW").Value = 0
    .Range("TASK_BASED_TEST_DESIGN").Value = 0
    .Range("TASK_BASED_TEST_IMPLEMENT").Value = 0
    .Range("TASK_BASED_TEST_TEST").Value = 0
    .Range("TASK_BASED_TEST_SITE").Value = 0
     
    ' Documentation
    .Range("TASK_BASED_DOC_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_DOC_1st_UNIT_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_DOC_2nd_UNIT_TOTAL_HOURS").Value = 0
    
    .Range("TASK_BASED_DOC_REM_HOURS").Value = 0
    .Range("TASK_BASED_DOC_1st_UNIT_REM_HOURS").Value = 0
    .Range("TASK_BASED_DOC_2nd_UNIT_REM_HOURS").Value = 0
    
    .Range("TASK_BASED_DOC_TOTAL_COST").Value = 0
    .Range("TASK_BASED_DOC_1st_UNIT_TOTAL_COST").Value = 0
    .Range("TASK_BASED_DOC_2nd_UNIT_TOTAL_COST").Value = 0
    
    .Range("TASK_BASED_DOC_LOC_COST").Value = 0
    .Range("TASK_BASED_DOC_1st_UNIT_LOC_COST").Value = 0
    .Range("TASK_BASED_DOC_2nd_UNIT_LOC_COST").Value = 0
    
    .Range("TASK_BASED_DOC_LOC_AC_HOURS").Value = 0
    .Range("TASK_BASED_DOC_LOC_LE_HOURS").Value = 0
    .Range("TASK_BASED_DOC_LOC_AE_HOURS").Value = 0
    .Range("TASK_BASED_DOC_LOC_T_HOURS").Value = 0
    .Range("TASK_BASED_DOC_LOC_SE_HOURS").Value = 0
    .Range("TASK_BASED_DOC_LOC_SR_HOURS").Value = 0
        
    .Range("TASK_BASED_DOC_REM_AC_HOURS").Value = 0
    .Range("TASK_BASED_DOC_REM_LE_HOURS").Value = 0
    .Range("TASK_BASED_DOC_REM_AE_HOURS").Value = 0
    .Range("TASK_BASED_DOC_REM_T_HOURS").Value = 0
    .Range("TASK_BASED_DOC_REM_SE_HOURS").Value = 0
    .Range("TASK_BASED_DOC_REM_SR_HOURS").Value = 0
    
    .Range("TASK_BASED_DOC_REVIEW").Value = 0
    .Range("TASK_BASED_DOC_DESIGN").Value = 0
    .Range("TASK_BASED_DOC_IMPLEMENT").Value = 0
    .Range("TASK_BASED_DOC_TEST").Value = 0
    .Range("TASK_BASED_DOC_SITE").Value = 0
 
    ' Training
    .Range("TASK_BASED_COURSE_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_COURSE_1st_UNIT_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_COURSE_2nd_UNIT_TOTAL_HOURS").Value = 0
    
    .Range("TASK_BASED_COURSE_REM_HOURS").Value = 0
    .Range("TASK_BASED_COURSE_1st_UNIT_REM_HOURS").Value = 0
    .Range("TASK_BASED_COURSE_2nd_UNIT_REM_HOURS").Value = 0
    
    .Range("TASK_BASED_COURSE_TOTAL_COST").Value = 0
    .Range("TASK_BASED_COURSE_1st_UNIT_TOTAL_COST").Value = 0
    .Range("TASK_BASED_COURSE_2nd_UNIT_TOTAL_COST").Value = 0
    
    .Range("TASK_BASED_COURSE_LOC_COST").Value = 0
    .Range("TASK_BASED_COURSE_1st_UNIT_LOC_COST").Value = 0
    .Range("TASK_BASED_COURSE_2nd_UNIT_LOC_COST").Value = 0
    
    .Range("TASK_BASED_COURSE_LOC_AC_HOURS").Value = 0
    .Range("TASK_BASED_COURSE_LOC_LE_HOURS").Value = 0
    .Range("TASK_BASED_COURSE_LOC_AE_HOURS").Value = 0
    .Range("TASK_BASED_COURSE_LOC_T_HOURS").Value = 0
    .Range("TASK_BASED_COURSE_LOC_SE_HOURS").Value = 0
    .Range("TASK_BASED_COURSE_LOC_SR_HOURS").Value = 0
        
    .Range("TASK_BASED_COURSE_REM_AC_HOURS").Value = 0
    .Range("TASK_BASED_COURSE_REM_LE_HOURS").Value = 0
    .Range("TASK_BASED_COURSE_REM_AE_HOURS").Value = 0
    .Range("TASK_BASED_COURSE_REM_T_HOURS").Value = 0
    .Range("TASK_BASED_COURSE_REM_SE_HOURS").Value = 0
    .Range("TASK_BASED_COURSE_REM_SR_HOURS").Value = 0
    
    .Range("TASK_BASED_COURSE_REVIEW").Value = 0
    .Range("TASK_BASED_COURSE_DESIGN").Value = 0
    .Range("TASK_BASED_COURSE_IMPLEMENT").Value = 0
    .Range("TASK_BASED_COURSE_TEST").Value = 0
    .Range("TASK_BASED_COURSE_SITE").Value = 0
     
    ' Project Management
    .Range("TASK_BASED_PM_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_PM_1st_UNIT_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_PM_2nd_UNIT_TOTAL_HOURS").Value = 0
    
    .Range("TASK_BASED_PM_REM_HOURS").Value = 0
    .Range("TASK_BASED_PM_1st_UNIT_REM_HOURS").Value = 0
    .Range("TASK_BASED_PM_2nd_UNIT_REM_HOURS").Value = 0
    
    .Range("TASK_BASED_PM_TOTAL_COST").Value = 0
    .Range("TASK_BASED_PM_1st_UNIT_TOTAL_COST").Value = 0
    .Range("TASK_BASED_PM_2nd_UNIT_TOTAL_COST").Value = 0
    
    .Range("TASK_BASED_PM_LOC_COST").Value = 0
    .Range("TASK_BASED_PM_1st_UNIT_LOC_COST").Value = 0
    .Range("TASK_BASED_PM_2nd_UNIT_LOC_COST").Value = 0
    
    .Range("TASK_BASED_PM_REVIEW").Value = 0
    .Range("TASK_BASED_PM_DESIGN").Value = 0
    .Range("TASK_BASED_PM_IMPLEMENT").Value = 0
    .Range("TASK_BASED_PM_TEST").Value = 0
    .Range("TASK_BASED_PM_SITE").Value = 0
           
    ' Meetings
    .Range("TASK_BASED_MEETINGS_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_MEETINGS_1st_UNIT_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_MEETINGS_2nd_UNIT_TOTAL_HOURS").Value = 0
    
    .Range("TASK_BASED_MEETINGS_REM_HOURS").Value = 0
    .Range("TASK_BASED_MEETINGS_1st_UNIT_REM_HOURS").Value = 0
    .Range("TASK_BASED_MEETINGS_2nd_UNIT_REM_HOURS").Value = 0
    
    .Range("TASK_BASED_MEETINGS_TOTAL_COST").Value = 0
    .Range("TASK_BASED_MEETINGS_1st_UNIT_TOTAL_COST").Value = 0
    .Range("TASK_BASED_MEETINGS_2nd_UNIT_TOTAL_COST").Value = 0
    
    .Range("TASK_BASED_MEETINGS_LOC_COST").Value = 0
    .Range("TASK_BASED_MEETINGS_1st_UNIT_LOC_COST").Value = 0
    .Range("TASK_BASED_MEETINGS_2nd_UNIT_LOC_COST").Value = 0
    
    .Range("TASK_BASED_MEETING_LOC_AC_HOURS").Value = 0
    .Range("TASK_BASED_MEETING_LOC_LE_HOURS").Value = 0
    .Range("TASK_BASED_MEETING_LOC_AE_HOURS").Value = 0
    .Range("TASK_BASED_MEETING_LOC_T_HOURS").Value = 0
    .Range("TASK_BASED_MEETING_LOC_SE_HOURS").Value = 0
    .Range("TASK_BASED_MEETING_LOC_SR_HOURS").Value = 0
        
    .Range("TASK_BASED_MEETING_REM_AC_HOURS").Value = 0
    .Range("TASK_BASED_MEETING_REM_LE_HOURS").Value = 0
    .Range("TASK_BASED_MEETING_REM_AE_HOURS").Value = 0
    .Range("TASK_BASED_MEETING_REM_T_HOURS").Value = 0
    .Range("TASK_BASED_MEETING_REM_SE_HOURS").Value = 0
    .Range("TASK_BASED_MEETING_REM_SR_HOURS").Value = 0
    
    .Range("TASK_BASED_MEETING_REVIEW").Value = 0
    .Range("TASK_BASED_MEETING_DESIGN").Value = 0
    .Range("TASK_BASED_MEETING_IMPLEMENT").Value = 0
    .Range("TASK_BASED_MEETING_TEST").Value = 0
    .Range("TASK_BASED_MEETING_SITE").Value = 0
    
    ' CLEAN ORDER PROCESSING
    .Range("TASK_BASED_CLEAN_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_CLEAN_1st_UNIT_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_CLEAN_2nd_UNIT_TOTAL_HOURS").Value = 0
    
    .Range("TASK_BASED_CLEAN_REM_HOURS").Value = 0
    .Range("TASK_BASED_CLEAN_1st_UNIT_REM_HOURS").Value = 0
    .Range("TASK_BASED_CLEAN_2nd_UNIT_REM_HOURS").Value = 0
    
    .Range("TASK_BASED_CLEAN_TOTAL_COST").Value = 0
    .Range("TASK_BASED_CLEAN_1st_UNIT_TOTAL_COST").Value = 0
    .Range("TASK_BASED_CLEAN_2nd_UNIT_TOTAL_COST").Value = 0
    
    .Range("TASK_BASED_CLEAN_LOC_COST").Value = 0
    .Range("TASK_BASED_CLEAN_1st_UNIT_LOC_COST").Value = 0
    .Range("TASK_BASED_CLEAN_2nd_UNIT_LOC_COST").Value = 0
    
    .Range("TASK_BASED_CLEAN_LOC_AC_HOURS").Value = 0
    .Range("TASK_BASED_CLEAN_LOC_LE_HOURS").Value = 0
    .Range("TASK_BASED_CLEAN_LOC_AE_HOURS").Value = 0
    .Range("TASK_BASED_CLEAN_LOC_T_HOURS").Value = 0
    .Range("TASK_BASED_CLEAN_LOC_SE_HOURS").Value = 0
    .Range("TASK_BASED_CLEAN_LOC_SR_HOURS").Value = 0
        
    .Range("TASK_BASED_CLEAN_REM_AC_HOURS").Value = 0
    .Range("TASK_BASED_CLEAN_REM_LE_HOURS").Value = 0
    .Range("TASK_BASED_CLEAN_REM_AE_HOURS").Value = 0
    .Range("TASK_BASED_CLEAN_REM_T_HOURS").Value = 0
    .Range("TASK_BASED_CLEAN_REM_SE_HOURS").Value = 0
    .Range("TASK_BASED_CLEAN_REM_SR_HOURS").Value = 0
    
    .Range("TASK_BASED_CLEAN_REVIEW").Value = 0
    .Range("TASK_BASED_CLEAN_DESIGN").Value = 0
    .Range("TASK_BASED_CLEAN_IMPLEMENT").Value = 0
    .Range("TASK_BASED_CLEAN_TEST").Value = 0
    .Range("TASK_BASED_CLEAN_SITE").Value = 0
    
     ' Site
    .Range("TASK_BASED_SITE_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_SITE_1st_UNIT_TOTAL_HOURS").Value = 0
    .Range("TASK_BASED_SITE_2nd_UNIT_TOTAL_HOURS").Value = 0
    
    .Range("TASK_BASED_SITE_REM_HOURS").Value = 0
    .Range("TASK_BASED_SITE_1st_UNIT_REM_HOURS").Value = 0
    .Range("TASK_BASED_SITE_2nd_UNIT_REM_HOURS").Value = 0
    
    .Range("TASK_BASED_SITE_TOTAL_COST").Value = 0
    .Range("TASK_BASED_SITE_1st_UNIT_TOTAL_COST").Value = 0
    .Range("TASK_BASED_SITE_2nd_UNIT_TOTAL_COST").Value = 0
    
    .Range("TASK_BASED_SITE_LOC_COST").Value = 0
    .Range("TASK_BASED_SITE_1st_UNIT_LOC_COST").Value = 0
    .Range("TASK_BASED_SITE_2nd_UNIT_LOC_COST").Value = 0
    
    .Range("TASK_BASED_SITE_LOC_AC_HOURS").Value = 0
    .Range("TASK_BASED_SITE_LOC_LE_HOURS").Value = 0
    .Range("TASK_BASED_SITE_LOC_AE_HOURS").Value = 0
    .Range("TASK_BASED_SITE_LOC_T_HOURS").Value = 0
    .Range("TASK_BASED_SITE_LOC_SE_HOURS").Value = 0
    .Range("TASK_BASED_SITE_LOC_SR_HOURS").Value = 0
        
    .Range("TASK_BASED_SITE_REM_AC_HOURS").Value = 0
    .Range("TASK_BASED_SITE_REM_LE_HOURS").Value = 0
    .Range("TASK_BASED_SITE_REM_AE_HOURS").Value = 0
    .Range("TASK_BASED_SITE_REM_T_HOURS").Value = 0
    .Range("TASK_BASED_SITE_REM_SE_HOURS").Value = 0
    .Range("TASK_BASED_SITE_REM_SR_HOURS").Value = 0
    
    .Range("TASK_BASED_SITE_REVIEW").Value = 0
    .Range("TASK_BASED_SITE_DESIGN").Value = 0
    .Range("TASK_BASED_SITE_IMPLEMENT").Value = 0
    .Range("TASK_BASED_SITE_TEST").Value = 0
    .Range("TASK_BASED_SITE_SITE").Value = 0
    
    Call SETTaskBasedSummaryUpdate
End With


CleanUp:
Application.Cursor = xlDefault
On Error Resume Next
    Set tmp = Nothing
    Set s = Nothing
    Set wb = Nothing
Exit Function

ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Function

Public Function SETTaskBasedSummaryUpdate() As Boolean

On Error GoTo ErrHandler
'gets the name of the opened GECE sheet workbook
Dim activity As String
Dim i As Integer
Dim tmp
Dim wb As Workbooks
Dim s As Worksheet

'clear variable
' ***** need to initialise the variables in the summary *****
activity = ""
i = 0
Application.Cursor = xlWait

With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECETaskBasedSheet)
    For i = 0 To 297
        activity = .Range("TRACK_ACTIVITY_1").Offset(i, 0).Value
    
         Select Case activity
            Case "SPECIFICATIONS"
                tmp = .Range("TASK_BASED_SPEC_TOTAL_HOURS").Value
                .Range("TASK_BASED_SPEC_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_SPEC_1st_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_SPEC_1st_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_SPEC_2nd_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_SPEC_2nd_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_SPEC_REM_HOURS").Value
                .Range("TASK_BASED_SPEC_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_SPEC_1st_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_SPEC_1st_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -2).Value + tmp
                tmp = .Range("TASK_BASED_SPEC_2nd_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_SPEC_2nd_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_SPEC_TOTAL_COST").Value
                .Range("TASK_BASED_SPEC_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_SPEC_1st_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_SPEC_1st_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_SPEC_2nd_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_SPEC_2nd_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_SPEC_LOC_COST").Value
                .Range("TASK_BASED_SPEC_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_SPEC_1st_UNIT_LOC_COST").Value
                .Range("TASK_BASED_SPEC_1st_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -2).Value + tmp
                tmp = .Range("TASK_BASED_SPEC_2nd_UNIT_LOC_COST").Value
                .Range("TASK_BASED_SPEC_2nd_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_SPEC_LOC_AC_HOURS").Value
                .Range("TASK_BASED_SPEC_LOC_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_SPEC_LOC_LE_HOURS").Value
                .Range("TASK_BASED_SPEC_LOC_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 1).Value + tmp
                tmp = .Range("TASK_BASED_SPEC_LOC_AE_HOURS").Value
                .Range("TASK_BASED_SPEC_LOC_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 2).Value + tmp
                tmp = .Range("TASK_BASED_SPEC_LOC_T_HOURS").Value
                .Range("TASK_BASED_SPEC_LOC_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 3).Value + tmp
                tmp = .Range("TASK_BASED_SPEC_LOC_SE_HOURS").Value
                .Range("TASK_BASED_SPEC_LOC_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 4).Value + tmp
                tmp = .Range("TASK_BASED_SPEC_LOC_SR_HOURS").Value
                .Range("TASK_BASED_SPEC_LOC_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 5).Value + tmp
                
                 tmp = .Range("TASK_BASED_SPEC_REM_AC_HOURS").Value
                .Range("TASK_BASED_SPEC_REM_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_SPEC_REM_LE_HOURS").Value
                .Range("TASK_BASED_SPEC_REM_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 1).Value + tmp
                tmp = .Range("TASK_BASED_SPEC_REM_AE_HOURS").Value
                .Range("TASK_BASED_SPEC_REM_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 2).Value + tmp
                tmp = .Range("TASK_BASED_SPEC_REM_T_HOURS").Value
                .Range("TASK_BASED_SPEC_REM_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 3).Value + tmp
                tmp = .Range("TASK_BASED_SPEC_REM_SE_HOURS").Value
                .Range("TASK_BASED_SPEC_REM_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 4).Value + tmp
                tmp = .Range("TASK_BASED_SPEC_REM_SR_HOURS").Value
                .Range("TASK_BASED_SPEC_REM_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 5).Value + tmp
                
                ' PHASES
                tmp = .Range("TASK_BASED_SPEC_REVIEW").Value
                .Range("TASK_BASED_SPEC_REVIEW").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_SPEC_DESIGN").Value
                .Range("TASK_BASED_SPEC_DESIGN").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 1).Value + tmp
                tmp = .Range("TASK_BASED_SPEC_IMPLEMENT").Value
                .Range("TASK_BASED_SPEC_IMPLEMENT").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 2).Value + tmp
                tmp = .Range("TASK_BASED_SPEC_TEST").Value
                .Range("TASK_BASED_SPEC_TEST").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 3).Value + tmp
                tmp = .Range("TASK_BASED_SPEC_SITE").Value
                .Range("TASK_BASED_SPEC_SITE").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 4).Value + tmp
                
                ' PM PHASES
                tmp = .Range("TASK_BASED_PM_REVIEW").Value
                .Range("TASK_BASED_PM_REVIEW").Value = .Range("TASK_BASED_PM_REVIEW_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_DESIGN").Value
                .Range("TASK_BASED_PM_DESIGN").Value = .Range("TASK_BASED_PM_DESIGN_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_IMPLEMENT").Value
                .Range("TASK_BASED_PM_IMPLEMENT").Value = .Range("TASK_BASED_PM_IMPLEMENT_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_TEST").Value
                .Range("TASK_BASED_PM_TEST").Value = .Range("TASK_BASED_PM_TEST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_SITE").Value
                .Range("TASK_BASED_PM_SITE").Value = .Range("TASK_BASED_PM_SITE_START").Offset(i, 0).Value + tmp
                
                
            Case "SYSTEM ENGINEERING"
                tmp = .Range("TASK_BASED_SYS_ENG_TOTAL_HOURS").Value
                .Range("TASK_BASED_SYS_ENG_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_SYS_ENG_1st_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_SYS_ENG_1st_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_SYS_ENG_2nd_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_SYS_ENG_2nd_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_SYS_ENG_REM_HOURS").Value
                .Range("TASK_BASED_SYS_ENG_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_SYS_ENG_1st_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_SYS_ENG_1st_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -2).Value + tmp
                tmp = .Range("TASK_BASED_SYS_ENG_2nd_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_SYS_ENG_2nd_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_SYS_ENG_TOTAL_COST").Value
                .Range("TASK_BASED_SYS_ENG_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_SYS_ENG_1st_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_SYS_ENG_1st_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_SYS_ENG_2nd_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_SYS_ENG_2nd_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_SYS_ENG_LOC_COST").Value
                .Range("TASK_BASED_SYS_ENG_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_SYS_ENG_1st_UNIT_LOC_COST").Value
                .Range("TASK_BASED_SYS_ENG_1st_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -2).Value + tmp
                tmp = .Range("TASK_BASED_SYS_ENG_2nd_UNIT_LOC_COST").Value
                .Range("TASK_BASED_SYS_ENG_2nd_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_SYS_ENG_LOC_AC_HOURS").Value
                .Range("TASK_BASED_SYS_ENG_LOC_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_SYS_ENG_LOC_LE_HOURS").Value
                .Range("TASK_BASED_SYS_ENG_LOC_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 1).Value + tmp
                tmp = .Range("TASK_BASED_SYS_ENG_LOC_AE_HOURS").Value
                .Range("TASK_BASED_SYS_ENG_LOC_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 2).Value + tmp
                tmp = .Range("TASK_BASED_SYS_ENG_LOC_T_HOURS").Value
                .Range("TASK_BASED_SYS_ENG_LOC_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 3).Value + tmp
                tmp = .Range("TASK_BASED_SYS_ENG_LOC_SE_HOURS").Value
                .Range("TASK_BASED_SYS_ENG_LOC_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 4).Value + tmp
                tmp = .Range("TASK_BASED_SYS_ENG_LOC_SR_HOURS").Value
                .Range("TASK_BASED_SYS_ENG_LOC_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 5).Value + tmp
                
                 tmp = .Range("TASK_BASED_SYS_ENG_REM_AC_HOURS").Value
                .Range("TASK_BASED_SYS_ENG_REM_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_SYS_ENG_REM_LE_HOURS").Value
                .Range("TASK_BASED_SYS_ENG_REM_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 1).Value + tmp
                tmp = .Range("TASK_BASED_SYS_ENG_REM_AE_HOURS").Value
                .Range("TASK_BASED_SYS_ENG_REM_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 2).Value + tmp
                tmp = .Range("TASK_BASED_SYS_ENG_REM_T_HOURS").Value
                .Range("TASK_BASED_SYS_ENG_REM_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 3).Value + tmp
                tmp = .Range("TASK_BASED_SYS_ENG_REM_SE_HOURS").Value
                .Range("TASK_BASED_SYS_ENG_REM_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 4).Value + tmp
                tmp = .Range("TASK_BASED_SYS_ENG_REM_SR_HOURS").Value
                .Range("TASK_BASED_SYS_ENG_REM_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 5).Value + tmp
                
                tmp = .Range("TASK_BASED_SYSENG_REVIEW").Value
                .Range("TASK_BASED_SYSENG_REVIEW").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_SYSENG_DESIGN").Value
                .Range("TASK_BASED_SYSENG_DESIGN").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 1).Value + tmp
                tmp = .Range("TASK_BASED_SYSENG_IMPLEMENT").Value
                .Range("TASK_BASED_SYSENG_IMPLEMENT").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 2).Value + tmp
                tmp = .Range("TASK_BASED_SYSENG_TEST").Value
                .Range("TASK_BASED_SYSENG_TEST").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 3).Value + tmp
                tmp = .Range("TASK_BASED_SYSENG_SITE").Value
                .Range("TASK_BASED_SYSENG_SITE").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 4).Value + tmp
                
                ' PM PHASES
                tmp = .Range("TASK_BASED_PM_REVIEW").Value
                .Range("TASK_BASED_PM_REVIEW").Value = .Range("TASK_BASED_PM_REVIEW_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_DESIGN").Value
                .Range("TASK_BASED_PM_DESIGN").Value = .Range("TASK_BASED_PM_DESIGN_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_IMPLEMENT").Value
                .Range("TASK_BASED_PM_IMPLEMENT").Value = .Range("TASK_BASED_PM_IMPLEMENT_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_TEST").Value
                .Range("TASK_BASED_PM_TEST").Value = .Range("TASK_BASED_PM_TEST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_SITE").Value
                .Range("TASK_BASED_PM_SITE").Value = .Range("TASK_BASED_PM_SITE_START").Offset(i, 0).Value + tmp
        
            Case "HUMAN MACHINE INTERFACE"
                tmp = .Range("TASK_BASED_HMI_TOTAL_HOURS").Value
                .Range("TASK_BASED_HMI_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_HMI_1st_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_HMI_1st_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_HMI_2nd_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_HMI_2nd_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_HMI_REM_HOURS").Value
                .Range("TASK_BASED_HMI_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_HMI_1st_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_HMI_1st_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -2).Value + tmp
                tmp = .Range("TASK_BASED_HMI_2nd_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_HMI_2nd_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_HMI_TOTAL_COST").Value
                .Range("TASK_BASED_HMI_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_HMI_1st_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_HMI_1st_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_HMI_2nd_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_HMI_2nd_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_HMI_LOC_COST").Value
                .Range("TASK_BASED_HMI_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_HMI_1st_UNIT_LOC_COST").Value
                .Range("TASK_BASED_HMI_1st_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -2).Value + tmp
                tmp = .Range("TASK_BASED_HMI_2nd_UNIT_LOC_COST").Value
                .Range("TASK_BASED_HMI_2nd_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_HMI_LOC_AC_HOURS").Value
                .Range("TASK_BASED_HMI_LOC_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_HMI_LOC_LE_HOURS").Value
                .Range("TASK_BASED_HMI_LOC_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 1).Value + tmp
                tmp = .Range("TASK_BASED_HMI_LOC_AE_HOURS").Value
                .Range("TASK_BASED_HMI_LOC_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 2).Value + tmp
                tmp = .Range("TASK_BASED_HMI_LOC_T_HOURS").Value
                .Range("TASK_BASED_HMI_LOC_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 3).Value + tmp
                tmp = .Range("TASK_BASED_HMI_LOC_SE_HOURS").Value
                .Range("TASK_BASED_HMI_LOC_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 4).Value + tmp
                tmp = .Range("TASK_BASED_HMI_LOC_SR_HOURS").Value
                .Range("TASK_BASED_HMI_LOC_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 5).Value + tmp
                
                 tmp = .Range("TASK_BASED_HMI_REM_AC_HOURS").Value
                .Range("TASK_BASED_HMI_REM_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_HMI_REM_LE_HOURS").Value
                .Range("TASK_BASED_HMI_REM_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 1).Value + tmp
                tmp = .Range("TASK_BASED_HMI_REM_AE_HOURS").Value
                .Range("TASK_BASED_HMI_REM_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 2).Value + tmp
                tmp = .Range("TASK_BASED_HMI_REM_T_HOURS").Value
                .Range("TASK_BASED_HMI_REM_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 3).Value + tmp
                tmp = .Range("TASK_BASED_HMI_REM_SE_HOURS").Value
                .Range("TASK_BASED_HMI_REM_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 4).Value + tmp
                tmp = .Range("TASK_BASED_HMI_REM_SR_HOURS").Value
                .Range("TASK_BASED_HMI_REM_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 5).Value + tmp
                 
                tmp = .Range("TASK_BASED_HMI_REVIEW").Value
                .Range("TASK_BASED_HMI_REVIEW").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_HMI_DESIGN").Value
                .Range("TASK_BASED_HMI_DESIGN").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 1).Value + tmp
                tmp = .Range("TASK_BASED_HMI_IMPLEMENT").Value
                .Range("TASK_BASED_HMI_IMPLEMENT").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 2).Value + tmp
                tmp = .Range("TASK_BASED_HMI_TEST").Value
                .Range("TASK_BASED_HMI_TEST").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 3).Value + tmp
                tmp = .Range("TASK_BASED_HMI_SITE").Value
                .Range("TASK_BASED_HMI_SITE").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 4).Value + tmp
                
                ' PM PHASES
                tmp = .Range("TASK_BASED_PM_REVIEW").Value
                .Range("TASK_BASED_PM_REVIEW").Value = .Range("TASK_BASED_PM_REVIEW_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_DESIGN").Value
                .Range("TASK_BASED_PM_DESIGN").Value = .Range("TASK_BASED_PM_DESIGN_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_IMPLEMENT").Value
                .Range("TASK_BASED_PM_IMPLEMENT").Value = .Range("TASK_BASED_PM_IMPLEMENT_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_TEST").Value
                .Range("TASK_BASED_PM_TEST").Value = .Range("TASK_BASED_PM_TEST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_SITE").Value
                .Range("TASK_BASED_PM_SITE").Value = .Range("TASK_BASED_PM_SITE_START").Offset(i, 0).Value + tmp
                
             Case "CONFIGURATION / PROGRAMMING"
                tmp = .Range("TASK_BASED_CP_TOTAL_HOURS").Value
                .Range("TASK_BASED_CP_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_CP_1st_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_CP_1st_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_CP_2nd_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_CP_2nd_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_CP_REM_HOURS").Value
                .Range("TASK_BASED_CP_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_CP_1st_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_CP_1st_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -2).Value + tmp
                tmp = .Range("TASK_BASED_CP_2nd_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_CP_2nd_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_CP_TOTAL_COST").Value
                .Range("TASK_BASED_CP_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_CP_1st_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_CP_1st_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_CP_2nd_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_CP_2nd_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_CP_LOC_COST").Value
                .Range("TASK_BASED_CP_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_CP_1st_UNIT_LOC_COST").Value
                .Range("TASK_BASED_CP_1st_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -2).Value + tmp
                tmp = .Range("TASK_BASED_CP_2nd_UNIT_LOC_COST").Value
                .Range("TASK_BASED_CP_2nd_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_CP_LOC_AC_HOURS").Value
                .Range("TASK_BASED_CP_LOC_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_CP_LOC_LE_HOURS").Value
                .Range("TASK_BASED_CP_LOC_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 1).Value + tmp
                tmp = .Range("TASK_BASED_CP_LOC_AE_HOURS").Value
                .Range("TASK_BASED_CP_LOC_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 2).Value + tmp
                tmp = .Range("TASK_BASED_CP_LOC_T_HOURS").Value
                .Range("TASK_BASED_CP_LOC_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 3).Value + tmp
                tmp = .Range("TASK_BASED_CP_LOC_SE_HOURS").Value
                .Range("TASK_BASED_CP_LOC_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 4).Value + tmp
                tmp = .Range("TASK_BASED_CP_LOC_SR_HOURS").Value
                .Range("TASK_BASED_CP_LOC_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 5).Value + tmp
                
                 tmp = .Range("TASK_BASED_CP_REM_AC_HOURS").Value
                .Range("TASK_BASED_CP_REM_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_CP_REM_LE_HOURS").Value
                .Range("TASK_BASED_CP_REM_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 1).Value + tmp
                tmp = .Range("TASK_BASED_CP_REM_AE_HOURS").Value
                .Range("TASK_BASED_CP_REM_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 2).Value + tmp
                tmp = .Range("TASK_BASED_CP_REM_T_HOURS").Value
                .Range("TASK_BASED_CP_REM_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 3).Value + tmp
                tmp = .Range("TASK_BASED_CP_REM_SE_HOURS").Value
                .Range("TASK_BASED_CP_REM_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 4).Value + tmp
                tmp = .Range("TASK_BASED_CP_REM_SR_HOURS").Value
                .Range("TASK_BASED_CP_REM_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 5).Value + tmp
                
                tmp = .Range("TASK_BASED_CP_REVIEW").Value
                .Range("TASK_BASED_CP_REVIEW").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_CP_DESIGN").Value
                .Range("TASK_BASED_CP_DESIGN").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 1).Value + tmp
                tmp = .Range("TASK_BASED_CP_IMPLEMENT").Value
                .Range("TASK_BASED_CP_IMPLEMENT").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 2).Value + tmp
                tmp = .Range("TASK_BASED_CP_TEST").Value
                .Range("TASK_BASED_CP_TEST").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 3).Value + tmp
                tmp = .Range("TASK_BASED_CP_SITE").Value
                .Range("TASK_BASED_CP_SITE").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 4).Value + tmp
                
                ' PM PHASES
                tmp = .Range("TASK_BASED_PM_REVIEW").Value
                .Range("TASK_BASED_PM_REVIEW").Value = .Range("TASK_BASED_PM_REVIEW_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_DESIGN").Value
                .Range("TASK_BASED_PM_DESIGN").Value = .Range("TASK_BASED_PM_DESIGN_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_IMPLEMENT").Value
                .Range("TASK_BASED_PM_IMPLEMENT").Value = .Range("TASK_BASED_PM_IMPLEMENT_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_TEST").Value
                .Range("TASK_BASED_PM_TEST").Value = .Range("TASK_BASED_PM_TEST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_SITE").Value
                .Range("TASK_BASED_PM_SITE").Value = .Range("TASK_BASED_PM_SITE_START").Offset(i, 0).Value + tmp
                
            Case "DATA INTERFACE"
                tmp = .Range("TASK_BASED_DI_TOTAL_HOURS").Value
                .Range("TASK_BASED_DI_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_DI_1st_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_DI_1st_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_DI_2nd_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_DI_2nd_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_DI_REM_HOURS").Value
                .Range("TASK_BASED_DI_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_DI_1st_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_DI_1st_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -2).Value + tmp
                tmp = .Range("TASK_BASED_DI_2nd_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_DI_2nd_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_DI_TOTAL_COST").Value
                .Range("TASK_BASED_DI_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_DI_1st_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_DI_1st_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_DI_2nd_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_DI_2nd_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_DI_LOC_COST").Value
                .Range("TASK_BASED_DI_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_DI_1st_UNIT_LOC_COST").Value
                .Range("TASK_BASED_DI_1st_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -2).Value + tmp
                tmp = .Range("TASK_BASED_DI_2nd_UNIT_LOC_COST").Value
                .Range("TASK_BASED_DI_2nd_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_DI_LOC_AC_HOURS").Value
                .Range("TASK_BASED_DI_LOC_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_DI_LOC_LE_HOURS").Value
                .Range("TASK_BASED_DI_LOC_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 1).Value + tmp
                tmp = .Range("TASK_BASED_DI_LOC_AE_HOURS").Value
                .Range("TASK_BASED_DI_LOC_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 2).Value + tmp
                tmp = .Range("TASK_BASED_DI_LOC_T_HOURS").Value
                .Range("TASK_BASED_DI_LOC_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 3).Value + tmp
                tmp = .Range("TASK_BASED_DI_LOC_SE_HOURS").Value
                .Range("TASK_BASED_DI_LOC_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 4).Value + tmp
                tmp = .Range("TASK_BASED_DI_LOC_SR_HOURS").Value
                .Range("TASK_BASED_DI_LOC_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 5).Value + tmp
                
                 tmp = .Range("TASK_BASED_DI_REM_AC_HOURS").Value
                .Range("TASK_BASED_DI_REM_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_DI_REM_LE_HOURS").Value
                .Range("TASK_BASED_DI_REM_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 1).Value + tmp
                tmp = .Range("TASK_BASED_DI_REM_AE_HOURS").Value
                .Range("TASK_BASED_DI_REM_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 2).Value + tmp
                tmp = .Range("TASK_BASED_DI_REM_T_HOURS").Value
                .Range("TASK_BASED_DI_REM_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 3).Value + tmp
                tmp = .Range("TASK_BASED_DI_REM_SE_HOURS").Value
                .Range("TASK_BASED_DI_REM_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 4).Value + tmp
                tmp = .Range("TASK_BASED_DI_REM_SR_HOURS").Value
                .Range("TASK_BASED_DI_REM_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 5).Value + tmp
                
                tmp = .Range("TASK_BASED_DI_REVIEW").Value
                .Range("TASK_BASED_DI_REVIEW").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_DI_DESIGN").Value
                .Range("TASK_BASED_DI_DESIGN").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 1).Value + tmp
                tmp = .Range("TASK_BASED_DI_IMPLEMENT").Value
                .Range("TASK_BASED_DI_IMPLEMENT").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 2).Value + tmp
                tmp = .Range("TASK_BASED_DI_TEST").Value
                .Range("TASK_BASED_DI_TEST").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 3).Value + tmp
                tmp = .Range("TASK_BASED_DI_SITE").Value
                .Range("TASK_BASED_DI_SITE").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 4).Value + tmp
                
                ' PM PHASES
                tmp = .Range("TASK_BASED_PM_REVIEW").Value
                .Range("TASK_BASED_PM_REVIEW").Value = .Range("TASK_BASED_PM_REVIEW_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_DESIGN").Value
                .Range("TASK_BASED_PM_DESIGN").Value = .Range("TASK_BASED_PM_DESIGN_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_IMPLEMENT").Value
                .Range("TASK_BASED_PM_IMPLEMENT").Value = .Range("TASK_BASED_PM_IMPLEMENT_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_TEST").Value
                .Range("TASK_BASED_PM_TEST").Value = .Range("TASK_BASED_PM_TEST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_SITE").Value
                .Range("TASK_BASED_PM_SITE").Value = .Range("TASK_BASED_PM_SITE_START").Offset(i, 0).Value + tmp
                
            Case "ESD ESD"
                tmp = .Range("TASK_BASED_ESD_TOTAL_HOURS").Value
                .Range("TASK_BASED_ESD_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_ESD_1st_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_ESD_1st_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_ESD_2nd_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_ESD_2nd_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_ESD_REM_HOURS").Value
                .Range("TASK_BASED_ESD_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_ESD_1st_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_ESD_1st_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -2).Value + tmp
                tmp = .Range("TASK_BASED_ESD_2nd_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_ESD_2nd_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_ESD_TOTAL_COST").Value
                .Range("TASK_BASED_ESD_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_ESD_1st_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_ESD_1st_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_ESD_2nd_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_ESD_2nd_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_ESD_LOC_COST").Value
                .Range("TASK_BASED_ESD_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_ESD_1st_UNIT_LOC_COST").Value
                .Range("TASK_BASED_ESD_1st_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -2).Value + tmp
                tmp = .Range("TASK_BASED_ESD_2nd_UNIT_LOC_COST").Value
                .Range("TASK_BASED_ESD_2nd_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -1).Value + tmp
                                
                tmp = .Range("TASK_BASED_ESD_LOC_AC_HOURS").Value
                .Range("TASK_BASED_ESD_LOC_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_ESD_LOC_LE_HOURS").Value
                .Range("TASK_BASED_ESD_LOC_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 1).Value + tmp
                tmp = .Range("TASK_BASED_ESD_LOC_AE_HOURS").Value
                .Range("TASK_BASED_ESD_LOC_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 2).Value + tmp
                tmp = .Range("TASK_BASED_ESD_LOC_T_HOURS").Value
                .Range("TASK_BASED_ESD_LOC_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 3).Value + tmp
                tmp = .Range("TASK_BASED_ESD_LOC_SE_HOURS").Value
                .Range("TASK_BASED_ESD_LOC_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 4).Value + tmp
                tmp = .Range("TASK_BASED_ESD_LOC_SR_HOURS").Value
                .Range("TASK_BASED_ESD_LOC_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 5).Value + tmp
                
                 tmp = .Range("TASK_BASED_ESD_REM_AC_HOURS").Value
                .Range("TASK_BASED_ESD_REM_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_ESD_REM_LE_HOURS").Value
                .Range("TASK_BASED_ESD_REM_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 1).Value + tmp
                tmp = .Range("TASK_BASED_ESD_REM_AE_HOURS").Value
                .Range("TASK_BASED_ESD_REM_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 2).Value + tmp
                tmp = .Range("TASK_BASED_ESD_REM_T_HOURS").Value
                .Range("TASK_BASED_ESD_REM_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 3).Value + tmp
                tmp = .Range("TASK_BASED_ESD_REM_SE_HOURS").Value
                .Range("TASK_BASED_ESD_REM_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 4).Value + tmp
                tmp = .Range("TASK_BASED_ESD_REM_SR_HOURS").Value
                .Range("TASK_BASED_ESD_REM_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 5).Value + tmp
                    
                ' PM PHASES
                tmp = .Range("TASK_BASED_PM_REVIEW").Value
                .Range("TASK_BASED_PM_REVIEW").Value = .Range("TASK_BASED_PM_REVIEW_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_DESIGN").Value
                .Range("TASK_BASED_PM_DESIGN").Value = .Range("TASK_BASED_PM_DESIGN_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_IMPLEMENT").Value
                .Range("TASK_BASED_PM_IMPLEMENT").Value = .Range("TASK_BASED_PM_IMPLEMENT_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_TEST").Value
                .Range("TASK_BASED_PM_TEST").Value = .Range("TASK_BASED_PM_TEST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_SITE").Value
                .Range("TASK_BASED_PM_SITE").Value = .Range("TASK_BASED_PM_SITE_START").Offset(i, 0).Value + tmp
                
            Case "REPORTS"
                tmp = .Range("TASK_BASED_REP_TOTAL_HOURS").Value
                .Range("TASK_BASED_REP_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_REP_1st_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_REP_1st_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_REP_2nd_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_REP_2nd_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_REP_REM_HOURS").Value
                .Range("TASK_BASED_REP_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_REP_1st_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_REP_1st_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -2).Value + tmp
                tmp = .Range("TASK_BASED_REP_2nd_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_REP_2nd_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_REP_TOTAL_COST").Value
                .Range("TASK_BASED_REP_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_REP_1st_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_REP_1st_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_REP_2nd_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_REP_2nd_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_REP_LOC_COST").Value
                .Range("TASK_BASED_REP_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_REP_1st_UNIT_LOC_COST").Value
                .Range("TASK_BASED_REP_1st_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -2).Value + tmp
                tmp = .Range("TASK_BASED_REP_2nd_UNIT_LOC_COST").Value
                .Range("TASK_BASED_REP_2nd_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -1).Value + tmp
                                   
                tmp = .Range("TASK_BASED_REP_LOC_AC_HOURS").Value
                .Range("TASK_BASED_REP_LOC_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_REP_LOC_LE_HOURS").Value
                .Range("TASK_BASED_REP_LOC_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 1).Value + tmp
                tmp = .Range("TASK_BASED_REP_LOC_AE_HOURS").Value
                .Range("TASK_BASED_REP_LOC_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 2).Value + tmp
                tmp = .Range("TASK_BASED_REP_LOC_T_HOURS").Value
                .Range("TASK_BASED_REP_LOC_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 3).Value + tmp
                tmp = .Range("TASK_BASED_REP_LOC_SE_HOURS").Value
                .Range("TASK_BASED_REP_LOC_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 4).Value + tmp
                tmp = .Range("TASK_BASED_REP_LOC_SR_HOURS").Value
                .Range("TASK_BASED_REP_LOC_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 5).Value + tmp
                
                 tmp = .Range("TASK_BASED_REP_REM_AC_HOURS").Value
                .Range("TASK_BASED_REP_REM_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_REP_REM_LE_HOURS").Value
                .Range("TASK_BASED_REP_REM_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 1).Value + tmp
                tmp = .Range("TASK_BASED_REP_REM_AE_HOURS").Value
                .Range("TASK_BASED_REP_REM_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 2).Value + tmp
                tmp = .Range("TASK_BASED_REP_REM_T_HOURS").Value
                .Range("TASK_BASED_REP_REM_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 3).Value + tmp
                tmp = .Range("TASK_BASED_REP_REM_SE_HOURS").Value
                .Range("TASK_BASED_REP_REM_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 4).Value + tmp
                tmp = .Range("TASK_BASED_REP_REM_SR_HOURS").Value
                .Range("TASK_BASED_REP_REM_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 5).Value + tmp
                   
                tmp = .Range("TASK_BASED_REP_REVIEW").Value
                .Range("TASK_BASED_REP_REVIEW").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_REP_DESIGN").Value
                .Range("TASK_BASED_REP_DESIGN").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 1).Value + tmp
                tmp = .Range("TASK_BASED_REP_IMPLEMENT").Value
                .Range("TASK_BASED_REP_IMPLEMENT").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 2).Value + tmp
                tmp = .Range("TASK_BASED_REP_TEST").Value
                .Range("TASK_BASED_REP_TEST").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 3).Value + tmp
                tmp = .Range("TASK_BASED_REP_SITE").Value
                .Range("TASK_BASED_REP_SITE").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 4).Value + tmp
                
                ' PM PHASES
                tmp = .Range("TASK_BASED_PM_REVIEW").Value
                .Range("TASK_BASED_PM_REVIEW").Value = .Range("TASK_BASED_PM_REVIEW_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_DESIGN").Value
                .Range("TASK_BASED_PM_DESIGN").Value = .Range("TASK_BASED_PM_DESIGN_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_IMPLEMENT").Value
                .Range("TASK_BASED_PM_IMPLEMENT").Value = .Range("TASK_BASED_PM_IMPLEMENT_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_TEST").Value
                .Range("TASK_BASED_PM_TEST").Value = .Range("TASK_BASED_PM_TEST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_SITE").Value
                .Range("TASK_BASED_PM_SITE").Value = .Range("TASK_BASED_PM_SITE_START").Offset(i, 0).Value + tmp
                
            Case "APPLICATIONS"
                tmp = .Range("TASK_BASED_APP_TOTAL_HOURS").Value
                .Range("TASK_BASED_APP_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_APP_1st_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_APP_1st_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_APP_2nd_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_APP_2nd_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_APP_REM_HOURS").Value
                .Range("TASK_BASED_APP_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_APP_1st_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_APP_1st_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -2).Value + tmp
                tmp = .Range("TASK_BASED_APP_2nd_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_APP_2nd_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_APP_TOTAL_COST").Value
                .Range("TASK_BASED_APP_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_APP_1st_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_APP_1st_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_APP_2nd_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_APP_2nd_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_APP_LOC_COST").Value
                .Range("TASK_BASED_APP_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_APP_1st_UNIT_LOC_COST").Value
                .Range("TASK_BASED_APP_1st_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -2).Value + tmp
                tmp = .Range("TASK_BASED_APP_2nd_UNIT_LOC_COST").Value
                .Range("TASK_BASED_APP_2nd_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -1).Value + tmp
                     
                tmp = .Range("TASK_BASED_APP_LOC_AC_HOURS").Value
                .Range("TASK_BASED_APP_LOC_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_APP_LOC_LE_HOURS").Value
                .Range("TASK_BASED_APP_LOC_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 1).Value + tmp
                tmp = .Range("TASK_BASED_APP_LOC_AE_HOURS").Value
                .Range("TASK_BASED_APP_LOC_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 2).Value + tmp
                tmp = .Range("TASK_BASED_APP_LOC_T_HOURS").Value
                .Range("TASK_BASED_APP_LOC_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 3).Value + tmp
                tmp = .Range("TASK_BASED_APP_LOC_SE_HOURS").Value
                .Range("TASK_BASED_APP_LOC_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 4).Value + tmp
                tmp = .Range("TASK_BASED_APP_LOC_SR_HOURS").Value
                .Range("TASK_BASED_APP_LOC_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 5).Value + tmp
                
                 tmp = .Range("TASK_BASED_APP_REM_AC_HOURS").Value
                .Range("TASK_BASED_APP_REM_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_APP_REM_LE_HOURS").Value
                .Range("TASK_BASED_APP_REM_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 1).Value + tmp
                tmp = .Range("TASK_BASED_APP_REM_AE_HOURS").Value
                .Range("TASK_BASED_APP_REM_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 2).Value + tmp
                tmp = .Range("TASK_BASED_APP_REM_T_HOURS").Value
                .Range("TASK_BASED_APP_REM_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 3).Value + tmp
                tmp = .Range("TASK_BASED_APP_REM_SE_HOURS").Value
                .Range("TASK_BASED_APP_REM_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 4).Value + tmp
                tmp = .Range("TASK_BASED_APP_REM_SR_HOURS").Value
                .Range("TASK_BASED_APP_REM_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 5).Value + tmp
                   
                tmp = .Range("TASK_BASED_APP_REVIEW").Value
                .Range("TASK_BASED_APP_REVIEW").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_APP_DESIGN").Value
                .Range("TASK_BASED_APP_DESIGN").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 1).Value + tmp
                tmp = .Range("TASK_BASED_APP_IMPLEMENT").Value
                .Range("TASK_BASED_APP_IMPLEMENT").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 2).Value + tmp
                tmp = .Range("TASK_BASED_APP_TEST").Value
                .Range("TASK_BASED_APP_TEST").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 3).Value + tmp
                tmp = .Range("TASK_BASED_APP_SITE").Value
                .Range("TASK_BASED_APP_SITE").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 4).Value + tmp
                
                ' PM PHASES
                tmp = .Range("TASK_BASED_PM_REVIEW").Value
                .Range("TASK_BASED_PM_REVIEW").Value = .Range("TASK_BASED_PM_REVIEW_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_DESIGN").Value
                .Range("TASK_BASED_PM_DESIGN").Value = .Range("TASK_BASED_PM_DESIGN_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_IMPLEMENT").Value
                .Range("TASK_BASED_PM_IMPLEMENT").Value = .Range("TASK_BASED_PM_IMPLEMENT_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_TEST").Value
                .Range("TASK_BASED_PM_TEST").Value = .Range("TASK_BASED_PM_TEST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_SITE").Value
                .Range("TASK_BASED_PM_SITE").Value = .Range("TASK_BASED_PM_SITE_START").Offset(i, 0).Value + tmp
                
            Case "TESTING"
                tmp = .Range("TASK_BASED_TEST_TOTAL_HOURS").Value
                .Range("TASK_BASED_TEST_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_TEST_1st_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_TEST_1st_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_TEST_2nd_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_TEST_2nd_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_TEST_REM_HOURS").Value
                .Range("TASK_BASED_TEST_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_TEST_1st_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_TEST_1st_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -2).Value + tmp
                tmp = .Range("TASK_BASED_TEST_2nd_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_TEST_2nd_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_TEST_TOTAL_COST").Value
                .Range("TASK_BASED_TEST_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_TEST_1st_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_TEST_1st_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_TEST_2nd_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_TEST_2nd_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_TEST_LOC_COST").Value
                .Range("TASK_BASED_TEST_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_TEST_1st_UNIT_LOC_COST").Value
                .Range("TASK_BASED_TEST_1st_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -2).Value + tmp
                tmp = .Range("TASK_BASED_TEST_2nd_UNIT_LOC_COST").Value
                .Range("TASK_BASED_TEST_2nd_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -1).Value + tmp
                     
                tmp = .Range("TASK_BASED_TEST_LOC_AC_HOURS").Value
                .Range("TASK_BASED_TEST_LOC_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_TEST_LOC_LE_HOURS").Value
                .Range("TASK_BASED_TEST_LOC_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 1).Value + tmp
                tmp = .Range("TASK_BASED_TEST_LOC_AE_HOURS").Value
                .Range("TASK_BASED_TEST_LOC_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 2).Value + tmp
                tmp = .Range("TASK_BASED_TEST_LOC_T_HOURS").Value
                .Range("TASK_BASED_TEST_LOC_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 3).Value + tmp
                tmp = .Range("TASK_BASED_TEST_LOC_SE_HOURS").Value
                .Range("TASK_BASED_TEST_LOC_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 4).Value + tmp
                tmp = .Range("TASK_BASED_TEST_LOC_SR_HOURS").Value
                .Range("TASK_BASED_TEST_LOC_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 5).Value + tmp
                
                 tmp = .Range("TASK_BASED_TEST_REM_AC_HOURS").Value
                .Range("TASK_BASED_TEST_REM_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_TEST_REM_LE_HOURS").Value
                .Range("TASK_BASED_TEST_REM_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 1).Value + tmp
                tmp = .Range("TASK_BASED_TEST_REM_AE_HOURS").Value
                .Range("TASK_BASED_TEST_REM_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 2).Value + tmp
                tmp = .Range("TASK_BASED_TEST_REM_T_HOURS").Value
                .Range("TASK_BASED_TEST_REM_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 3).Value + tmp
                tmp = .Range("TASK_BASED_TEST_REM_SE_HOURS").Value
                .Range("TASK_BASED_TEST_REM_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 4).Value + tmp
                tmp = .Range("TASK_BASED_TEST_REM_SR_HOURS").Value
                .Range("TASK_BASED_TEST_REM_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 5).Value + tmp
                    
                tmp = .Range("TASK_BASED_TEST_REVIEW").Value
                .Range("TASK_BASED_TEST_REVIEW").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_TEST_DESIGN").Value
                .Range("TASK_BASED_TEST_DESIGN").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 1).Value + tmp
                tmp = .Range("TASK_BASED_TEST_IMPLEMENT").Value
                .Range("TASK_BASED_TEST_IMPLEMENT").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 2).Value + tmp
                tmp = .Range("TASK_BASED_TEST_TEST").Value
                .Range("TASK_BASED_TEST_TEST").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 3).Value + tmp
                tmp = .Range("TASK_BASED_TEST_SITE").Value
                .Range("TASK_BASED_TEST_SITE").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 4).Value + tmp
                
                ' PM PHASES
                tmp = .Range("TASK_BASED_PM_REVIEW").Value
                .Range("TASK_BASED_PM_REVIEW").Value = .Range("TASK_BASED_PM_REVIEW_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_DESIGN").Value
                .Range("TASK_BASED_PM_DESIGN").Value = .Range("TASK_BASED_PM_DESIGN_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_IMPLEMENT").Value
                .Range("TASK_BASED_PM_IMPLEMENT").Value = .Range("TASK_BASED_PM_IMPLEMENT_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_TEST").Value
                .Range("TASK_BASED_PM_TEST").Value = .Range("TASK_BASED_PM_TEST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_SITE").Value
                .Range("TASK_BASED_PM_SITE").Value = .Range("TASK_BASED_PM_SITE_START").Offset(i, 0).Value + tmp
                
           Case "DOCUMENTATION"
                tmp = .Range("TASK_BASED_DOC_TOTAL_HOURS").Value
                .Range("TASK_BASED_DOC_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_DOC_1st_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_DOC_1st_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_DOC_2nd_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_DOC_2nd_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_DOC_REM_HOURS").Value
                .Range("TASK_BASED_DOC_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_DOC_1st_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_DOC_1st_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -2).Value + tmp
                tmp = .Range("TASK_BASED_DOC_2nd_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_DOC_2nd_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_DOC_TOTAL_COST").Value
                .Range("TASK_BASED_DOC_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_DOC_1st_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_DOC_1st_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_DOC_2nd_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_DOC_2nd_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_DOC_LOC_COST").Value
                .Range("TASK_BASED_DOC_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_DOC_1st_UNIT_LOC_COST").Value
                .Range("TASK_BASED_DOC_1st_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -2).Value + tmp
                tmp = .Range("TASK_BASED_DOC_2nd_UNIT_LOC_COST").Value
                .Range("TASK_BASED_DOC_2nd_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -1).Value + tmp
                     
                tmp = .Range("TASK_BASED_DOC_LOC_AC_HOURS").Value
                .Range("TASK_BASED_DOC_LOC_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_DOC_LOC_LE_HOURS").Value
                .Range("TASK_BASED_DOC_LOC_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 1).Value + tmp
                tmp = .Range("TASK_BASED_DOC_LOC_AE_HOURS").Value
                .Range("TASK_BASED_DOC_LOC_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 2).Value + tmp
                tmp = .Range("TASK_BASED_DOC_LOC_T_HOURS").Value
                .Range("TASK_BASED_DOC_LOC_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 3).Value + tmp
                tmp = .Range("TASK_BASED_DOC_LOC_SE_HOURS").Value
                .Range("TASK_BASED_DOC_LOC_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 4).Value + tmp
                tmp = .Range("TASK_BASED_DOC_LOC_SR_HOURS").Value
                .Range("TASK_BASED_DOC_LOC_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 5).Value + tmp
                
                 tmp = .Range("TASK_BASED_DOC_REM_AC_HOURS").Value
                .Range("TASK_BASED_DOC_REM_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_DOC_REM_LE_HOURS").Value
                .Range("TASK_BASED_DOC_REM_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 1).Value + tmp
                tmp = .Range("TASK_BASED_DOC_REM_AE_HOURS").Value
                .Range("TASK_BASED_DOC_REM_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 2).Value + tmp
                tmp = .Range("TASK_BASED_DOC_REM_T_HOURS").Value
                .Range("TASK_BASED_DOC_REM_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 3).Value + tmp
                tmp = .Range("TASK_BASED_DOC_REM_SE_HOURS").Value
                .Range("TASK_BASED_DOC_REM_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 4).Value + tmp
                tmp = .Range("TASK_BASED_DOC_REM_SR_HOURS").Value
                .Range("TASK_BASED_DOC_REM_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 5).Value + tmp
                   
                tmp = .Range("TASK_BASED_DOC_REVIEW").Value
                .Range("TASK_BASED_DOC_REVIEW").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_DOC_DESIGN").Value
                .Range("TASK_BASED_DOC_DESIGN").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 1).Value + tmp
                tmp = .Range("TASK_BASED_DOC_IMPLEMENT").Value
                .Range("TASK_BASED_DOC_IMPLEMENT").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 2).Value + tmp
                tmp = .Range("TASK_BASED_DOC_TEST").Value
                .Range("TASK_BASED_DOC_TEST").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 3).Value + tmp
                tmp = .Range("TASK_BASED_DOC_SITE").Value
                .Range("TASK_BASED_DOC_SITE").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 4).Value + tmp
                
                ' PM PHASES
                tmp = .Range("TASK_BASED_PM_REVIEW").Value
                .Range("TASK_BASED_PM_REVIEW").Value = .Range("TASK_BASED_PM_REVIEW_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_DESIGN").Value
                .Range("TASK_BASED_PM_DESIGN").Value = .Range("TASK_BASED_PM_DESIGN_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_IMPLEMENT").Value
                .Range("TASK_BASED_PM_IMPLEMENT").Value = .Range("TASK_BASED_PM_IMPLEMENT_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_TEST").Value
                .Range("TASK_BASED_PM_TEST").Value = .Range("TASK_BASED_PM_TEST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_SITE").Value
                .Range("TASK_BASED_PM_SITE").Value = .Range("TASK_BASED_PM_SITE_START").Offset(i, 0).Value + tmp
                
            Case "CUSTOM TRAINING"
                tmp = .Range("TASK_BASED_COURSE_TOTAL_HOURS").Value
                .Range("TASK_BASED_COURSE_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_COURSE_1st_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_COURSE_1st_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_COURSE_2nd_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_COURSE_2nd_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_COURSE_REM_HOURS").Value
                .Range("TASK_BASED_COURSE_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_COURSE_1st_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_COURSE_1st_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -2).Value + tmp
                tmp = .Range("TASK_BASED_COURSE_2nd_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_COURSE_2nd_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_COURSE_TOTAL_COST").Value
                .Range("TASK_BASED_COURSE_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_COURSE_1st_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_COURSE_1st_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_COURSE_2nd_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_COURSE_2nd_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_COURSE_LOC_COST").Value
                .Range("TASK_BASED_COURSE_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_COURSE_1st_UNIT_LOC_COST").Value
                .Range("TASK_BASED_COURSE_1st_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -2).Value + tmp
                tmp = .Range("TASK_BASED_COURSE_2nd_UNIT_LOC_COST").Value
                .Range("TASK_BASED_COURSE_2nd_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -1).Value + tmp
                     
                tmp = .Range("TASK_BASED_COURSE_LOC_AC_HOURS").Value
                .Range("TASK_BASED_COURSE_LOC_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_COURSE_LOC_LE_HOURS").Value
                .Range("TASK_BASED_COURSE_LOC_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 1).Value + tmp
                tmp = .Range("TASK_BASED_COURSE_LOC_AE_HOURS").Value
                .Range("TASK_BASED_COURSE_LOC_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 2).Value + tmp
                tmp = .Range("TASK_BASED_COURSE_LOC_T_HOURS").Value
                .Range("TASK_BASED_COURSE_LOC_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 3).Value + tmp
                tmp = .Range("TASK_BASED_COURSE_LOC_SE_HOURS").Value
                .Range("TASK_BASED_COURSE_LOC_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 4).Value + tmp
                tmp = .Range("TASK_BASED_COURSE_LOC_SR_HOURS").Value
                .Range("TASK_BASED_COURSE_LOC_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 5).Value + tmp
                
                 tmp = .Range("TASK_BASED_COURSE_REM_AC_HOURS").Value
                .Range("TASK_BASED_COURSE_REM_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_COURSE_REM_LE_HOURS").Value
                .Range("TASK_BASED_COURSE_REM_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 1).Value + tmp
                tmp = .Range("TASK_BASED_COURSE_REM_AE_HOURS").Value
                .Range("TASK_BASED_COURSE_REM_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 2).Value + tmp
                tmp = .Range("TASK_BASED_COURSE_REM_T_HOURS").Value
                .Range("TASK_BASED_COURSE_REM_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 3).Value + tmp
                tmp = .Range("TASK_BASED_COURSE_REM_SE_HOURS").Value
                .Range("TASK_BASED_COURSE_REM_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 4).Value + tmp
                tmp = .Range("TASK_BASED_COURSE_REM_SR_HOURS").Value
                .Range("TASK_BASED_COURSE_REM_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 5).Value + tmp
                         
                tmp = .Range("TASK_BASED_COURSE_REVIEW").Value
                .Range("TASK_BASED_COURSE_REVIEW").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_COURSE_DESIGN").Value
                .Range("TASK_BASED_COURSE_DESIGN").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 1).Value + tmp
                tmp = .Range("TASK_BASED_COURSE_IMPLEMENT").Value
                .Range("TASK_BASED_COURSE_IMPLEMENT").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 2).Value + tmp
                tmp = .Range("TASK_BASED_COURSE_TEST").Value
                .Range("TASK_BASED_COURSE_TEST").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 3).Value + tmp
                tmp = .Range("TASK_BASED_COURSE_SITE").Value
                .Range("TASK_BASED_COURSE_SITE").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 4).Value + tmp
                
                ' PM PHASES
                tmp = .Range("TASK_BASED_PM_REVIEW").Value
                .Range("TASK_BASED_PM_REVIEW").Value = .Range("TASK_BASED_PM_REVIEW_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_DESIGN").Value
                .Range("TASK_BASED_PM_DESIGN").Value = .Range("TASK_BASED_PM_DESIGN_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_IMPLEMENT").Value
                .Range("TASK_BASED_PM_IMPLEMENT").Value = .Range("TASK_BASED_PM_IMPLEMENT_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_TEST").Value
                .Range("TASK_BASED_PM_TEST").Value = .Range("TASK_BASED_PM_TEST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_SITE").Value
                .Range("TASK_BASED_PM_SITE").Value = .Range("TASK_BASED_PM_SITE_START").Offset(i, 0).Value + tmp
                
            Case "PROJECT MANAGEMENT"
                tmp = .Range("TASK_BASED_PM_TOTAL_HOURS").Value
                .Range("TASK_BASED_PM_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_1st_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_PM_1st_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_PM_2nd_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_PM_2nd_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_PM_REM_HOURS").Value
                .Range("TASK_BASED_PM_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_1st_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_PM_1st_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -2).Value + tmp
                tmp = .Range("TASK_BASED_PM_2nd_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_PM_2nd_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_PM_TOTAL_COST").Value
                .Range("TASK_BASED_PM_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_1st_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_PM_1st_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_PM_2nd_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_PM_2nd_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_PM_LOC_COST").Value
                .Range("TASK_BASED_PM_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_1st_UNIT_LOC_COST").Value
                .Range("TASK_BASED_PM_1st_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -2).Value + tmp
                tmp = .Range("TASK_BASED_PM_2nd_UNIT_LOC_COST").Value
                .Range("TASK_BASED_PM_2nd_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -1).Value + tmp
                     
                ' PM PHASES
                tmp = .Range("TASK_BASED_PM_REVIEW").Value
                .Range("TASK_BASED_PM_REVIEW").Value = .Range("TASK_BASED_PM_REVIEW_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_DESIGN").Value
                .Range("TASK_BASED_PM_DESIGN").Value = .Range("TASK_BASED_PM_DESIGN_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_IMPLEMENT").Value
                .Range("TASK_BASED_PM_IMPLEMENT").Value = .Range("TASK_BASED_PM_IMPLEMENT_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_TEST").Value
                .Range("TASK_BASED_PM_TEST").Value = .Range("TASK_BASED_PM_TEST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_SITE").Value
                .Range("TASK_BASED_PM_SITE").Value = .Range("TASK_BASED_PM_SITE_START").Offset(i, 0).Value + tmp
                
            Case "MEETINGS"
                tmp = .Range("TASK_BASED_MEETINGS_TOTAL_HOURS").Value
                .Range("TASK_BASED_MEETINGS_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_MEETINGS_1st_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_MEETINGS_1st_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_MEETINGS_2nd_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_MEETINGS_2nd_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_MEETINGS_REM_HOURS").Value
                .Range("TASK_BASED_MEETINGS_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_MEETINGS_1st_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_MEETINGS_1st_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -2).Value + tmp
                tmp = .Range("TASK_BASED_MEETINGS_2nd_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_MEETINGS_2nd_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_MEETINGS_TOTAL_COST").Value
                .Range("TASK_BASED_MEETINGS_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_MEETINGS_1st_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_MEETINGS_1st_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_MEETINGS_2nd_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_MEETINGS_2nd_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_MEETINGS_LOC_COST").Value
                .Range("TASK_BASED_MEETINGS_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_MEETINGS_1st_UNIT_LOC_COST").Value
                .Range("TASK_BASED_MEETINGS_1st_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -2).Value + tmp
                tmp = .Range("TASK_BASED_MEETINGS_2nd_UNIT_LOC_COST").Value
                .Range("TASK_BASED_MEETINGS_2nd_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -1).Value + tmp
                   
                tmp = .Range("TASK_BASED_MEETING_LOC_AC_HOURS").Value
                .Range("TASK_BASED_MEETING_LOC_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_MEETING_LOC_LE_HOURS").Value
                .Range("TASK_BASED_MEETING_LOC_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 1).Value + tmp
                tmp = .Range("TASK_BASED_MEETING_LOC_AE_HOURS").Value
                .Range("TASK_BASED_MEETING_LOC_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 2).Value + tmp
                tmp = .Range("TASK_BASED_MEETING_LOC_T_HOURS").Value
                .Range("TASK_BASED_MEETING_LOC_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 3).Value + tmp
                tmp = .Range("TASK_BASED_MEETING_LOC_SE_HOURS").Value
                .Range("TASK_BASED_MEETING_LOC_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 4).Value + tmp
                tmp = .Range("TASK_BASED_MEETING_LOC_SR_HOURS").Value
                .Range("TASK_BASED_MEETING_LOC_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 5).Value + tmp
                
                 tmp = .Range("TASK_BASED_MEETING_REM_AC_HOURS").Value
                .Range("TASK_BASED_MEETING_REM_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_MEETING_REM_LE_HOURS").Value
                .Range("TASK_BASED_MEETING_REM_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 1).Value + tmp
                tmp = .Range("TASK_BASED_MEETING_REM_AE_HOURS").Value
                .Range("TASK_BASED_MEETING_REM_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 2).Value + tmp
                tmp = .Range("TASK_BASED_MEETING_REM_T_HOURS").Value
                .Range("TASK_BASED_MEETING_REM_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 3).Value + tmp
                tmp = .Range("TASK_BASED_MEETING_REM_SE_HOURS").Value
                .Range("TASK_BASED_MEETING_REM_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 4).Value + tmp
                tmp = .Range("TASK_BASED_MEETING_REM_SR_HOURS").Value
                .Range("TASK_BASED_MEETING_REM_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 5).Value + tmp
                     
                tmp = .Range("TASK_BASED_MEETING_REVIEW").Value
                .Range("TASK_BASED_MEETING_REVIEW").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_MEETING_DESIGN").Value
                .Range("TASK_BASED_MEETING_DESIGN").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 1).Value + tmp
                tmp = .Range("TASK_BASED_MEETING_IMPLEMENT").Value
                .Range("TASK_BASED_MEETING_IMPLEMENT").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 2).Value + tmp
                tmp = .Range("TASK_BASED_MEETING_TEST").Value
                .Range("TASK_BASED_MEETING_TEST").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 3).Value + tmp
                tmp = .Range("TASK_BASED_MEETING_SITE").Value
                .Range("TASK_BASED_MEETING_SITE").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 4).Value + tmp
                
                ' PM PHASES
                tmp = .Range("TASK_BASED_PM_REVIEW").Value
                .Range("TASK_BASED_PM_REVIEW").Value = .Range("TASK_BASED_PM_REVIEW_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_DESIGN").Value
                .Range("TASK_BASED_PM_DESIGN").Value = .Range("TASK_BASED_PM_DESIGN_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_IMPLEMENT").Value
                .Range("TASK_BASED_PM_IMPLEMENT").Value = .Range("TASK_BASED_PM_IMPLEMENT_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_TEST").Value
                .Range("TASK_BASED_PM_TEST").Value = .Range("TASK_BASED_PM_TEST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_SITE").Value
                .Range("TASK_BASED_PM_SITE").Value = .Range("TASK_BASED_PM_SITE_START").Offset(i, 0).Value + tmp
                
            Case "CABINET DESIGN"
                tmp = .Range("TASK_BASED_CLEAN_TOTAL_HOURS").Value
                .Range("TASK_BASED_CLEAN_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_CLEAN_1st_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_CLEAN_1st_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_CLEAN_2nd_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_CLEAN_2nd_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_CLEAN_REM_HOURS").Value
                .Range("TASK_BASED_CLEAN_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_CLEAN_1st_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_CLEAN_1st_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -2).Value + tmp
                tmp = .Range("TASK_BASED_CLEAN_2nd_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_CLEAN_2nd_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_CLEAN_TOTAL_COST").Value
                .Range("TASK_BASED_CLEAN_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_CLEAN_1st_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_CLEAN_1st_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_CLEAN_2nd_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_CLEAN_2nd_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_CLEAN_LOC_COST").Value
                .Range("TASK_BASED_CLEAN_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_CLEAN_1st_UNIT_LOC_COST").Value
                .Range("TASK_BASED_CLEAN_1st_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -2).Value + tmp
                tmp = .Range("TASK_BASED_CLEAN_2nd_UNIT_LOC_COST").Value
                .Range("TASK_BASED_CLEAN_2nd_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -1).Value + tmp
                   
                tmp = .Range("TASK_BASED_CLEAN_LOC_AC_HOURS").Value
                .Range("TASK_BASED_CLEAN_LOC_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_CLEAN_LOC_LE_HOURS").Value
                .Range("TASK_BASED_CLEAN_LOC_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 1).Value + tmp
                tmp = .Range("TASK_BASED_CLEAN_LOC_AE_HOURS").Value
                .Range("TASK_BASED_CLEAN_LOC_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 2).Value + tmp
                tmp = .Range("TASK_BASED_CLEAN_LOC_T_HOURS").Value
                .Range("TASK_BASED_CLEAN_LOC_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 3).Value + tmp
                tmp = .Range("TASK_BASED_CLEAN_LOC_SE_HOURS").Value
                .Range("TASK_BASED_CLEAN_LOC_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 4).Value + tmp
                tmp = .Range("TASK_BASED_CLEAN_LOC_SR_HOURS").Value
                .Range("TASK_BASED_CLEAN_LOC_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 5).Value + tmp
                
                 tmp = .Range("TASK_BASED_CLEAN_REM_AC_HOURS").Value
                .Range("TASK_BASED_CLEAN_REM_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_CLEAN_REM_LE_HOURS").Value
                .Range("TASK_BASED_CLEAN_REM_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 1).Value + tmp
                tmp = .Range("TASK_BASED_CLEAN_REM_AE_HOURS").Value
                .Range("TASK_BASED_CLEAN_REM_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 2).Value + tmp
                tmp = .Range("TASK_BASED_CLEAN_REM_T_HOURS").Value
                .Range("TASK_BASED_CLEAN_REM_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 3).Value + tmp
                tmp = .Range("TASK_BASED_CLEAN_REM_SE_HOURS").Value
                .Range("TASK_BASED_CLEAN_REM_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 4).Value + tmp
                tmp = .Range("TASK_BASED_CLEAN_REM_SR_HOURS").Value
                .Range("TASK_BASED_CLEAN_REM_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 5).Value + tmp
                     
                tmp = .Range("TASK_BASED_CLEAN_REVIEW").Value
                .Range("TASK_BASED_CLEAN_REVIEW").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_CLEAN_DESIGN").Value
                .Range("TASK_BASED_CLEAN_DESIGN").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 1).Value + tmp
                tmp = .Range("TASK_BASED_CLEAN_IMPLEMENT").Value
                .Range("TASK_BASED_CLEAN_IMPLEMENT").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 2).Value + tmp
                tmp = .Range("TASK_BASED_CLEAN_TEST").Value
                .Range("TASK_BASED_CLEAN_TEST").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 3).Value + tmp
                tmp = .Range("TASK_BASED_CLEAN_SITE").Value
                .Range("TASK_BASED_CLEAN_SITE").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 4).Value + tmp
                
                ' PM PHASES
                tmp = .Range("TASK_BASED_PM_REVIEW").Value
                .Range("TASK_BASED_PM_REVIEW").Value = .Range("TASK_BASED_PM_REVIEW_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_DESIGN").Value
                .Range("TASK_BASED_PM_DESIGN").Value = .Range("TASK_BASED_PM_DESIGN_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_IMPLEMENT").Value
                .Range("TASK_BASED_PM_IMPLEMENT").Value = .Range("TASK_BASED_PM_IMPLEMENT_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_TEST").Value
                .Range("TASK_BASED_PM_TEST").Value = .Range("TASK_BASED_PM_TEST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_SITE").Value
                .Range("TASK_BASED_PM_SITE").Value = .Range("TASK_BASED_PM_SITE_START").Offset(i, 0).Value + tmp
                
            Case "SITE"
                tmp = .Range("TASK_BASED_SITE_TOTAL_HOURS").Value
                .Range("TASK_BASED_SITE_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_SITE_1st_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_SITE_1st_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_SITE_2nd_UNIT_TOTAL_HOURS").Value
                .Range("TASK_BASED_SITE_2nd_UNIT_TOTAL_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_SITE_REM_HOURS").Value
                .Range("TASK_BASED_SITE_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_SITE_1st_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_SITE_1st_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -2).Value + tmp
                tmp = .Range("TASK_BASED_SITE_2nd_UNIT_REM_HOURS").Value
                .Range("TASK_BASED_SITE_2nd_UNIT_REM_HOURS").Value = .Range("TASK_BASED_TOTAL_HOURS_START").Offset(i + 2, -1).Value + tmp
                
                tmp = .Range("TASK_BASED_SITE_TOTAL_COST").Value
                .Range("TASK_BASED_SITE_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_SITE_1st_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_SITE_1st_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -2).Value + tmp
                tmp = .Range("TASK_BASED_SITE_2nd_UNIT_TOTAL_COST").Value
                .Range("TASK_BASED_SITE_2nd_UNIT_TOTAL_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i, -1).Value + tmp
                 
                tmp = .Range("TASK_BASED_SITE_LOC_COST").Value
                .Range("TASK_BASED_SITE_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_SITE_1st_UNIT_LOC_COST").Value
                .Range("TASK_BASED_SITE_1st_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -2).Value + tmp
                tmp = .Range("TASK_BASED_SITE_2nd_UNIT_LOC_COST").Value
                .Range("TASK_BASED_SITE_2nd_UNIT_LOC_COST").Value = .Range("TASK_BASED_TOTAL_COST_START").Offset(i + 1, -1).Value + tmp
                   
                tmp = .Range("TASK_BASED_SITE_LOC_AC_HOURS").Value
                .Range("TASK_BASED_SITE_LOC_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 0).Value + tmp
                tmp = .Range("TASK_BASED_SITE_LOC_LE_HOURS").Value
                .Range("TASK_BASED_SITE_LOC_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 1).Value + tmp
                tmp = .Range("TASK_BASED_SITE_LOC_AE_HOURS").Value
                .Range("TASK_BASED_SITE_LOC_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 2).Value + tmp
                tmp = .Range("TASK_BASED_SITE_LOC_T_HOURS").Value
                .Range("TASK_BASED_SITE_LOC_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 3).Value + tmp
                tmp = .Range("TASK_BASED_SITE_LOC_SE_HOURS").Value
                .Range("TASK_BASED_SITE_LOC_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 4).Value + tmp
                tmp = .Range("TASK_BASED_SITE_LOC_SR_HOURS").Value
                .Range("TASK_BASED_SITE_LOC_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 1, 5).Value + tmp
                
                 tmp = .Range("TASK_BASED_SITE_REM_AC_HOURS").Value
                .Range("TASK_BASED_SITE_REM_AC_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 0).Value + tmp
                tmp = .Range("TASK_BASED_SITE_REM_LE_HOURS").Value
                .Range("TASK_BASED_SITE_REM_LE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 1).Value + tmp
                tmp = .Range("TASK_BASED_SITE_REM_AE_HOURS").Value
                .Range("TASK_BASED_SITE_REM_AE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 2).Value + tmp
                tmp = .Range("TASK_BASED_SITE_REM_T_HOURS").Value
                .Range("TASK_BASED_SITE_REM_T_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 3).Value + tmp
                tmp = .Range("TASK_BASED_SITE_REM_SE_HOURS").Value
                .Range("TASK_BASED_SITE_REM_SE_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 4).Value + tmp
                tmp = .Range("TASK_BASED_SITE_REM_SR_HOURS").Value
                .Range("TASK_BASED_SITE_REM_SR_HOURS").Value = .Range("TASK_BASED_TOTAL_ENG_CAT_START").Offset(i + 2, 5).Value + tmp
                 
                tmp = .Range("TASK_BASED_SITE_REVIEW").Value
                .Range("TASK_BASED_SITE_REVIEW").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_SITE_DESIGN").Value
                .Range("TASK_BASED_SITE_DESIGN").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 1).Value + tmp
                tmp = .Range("TASK_BASED_SITE_IMPLEMENT").Value
                .Range("TASK_BASED_SITE_IMPLEMENT").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 2).Value + tmp
                tmp = .Range("TASK_BASED_SITE_TEST").Value
                .Range("TASK_BASED_SITE_TEST").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 3).Value + tmp
                tmp = .Range("TASK_BASED_SITE_SITE").Value
                .Range("TASK_BASED_SITE_SITE").Value = .Range("TASK_BASED_TOTAL_PHASES_START").Offset(i, 4).Value + tmp
                
                ' PM PHASES
                tmp = .Range("TASK_BASED_PM_REVIEW").Value
                .Range("TASK_BASED_PM_REVIEW").Value = .Range("TASK_BASED_PM_REVIEW_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_DESIGN").Value
                .Range("TASK_BASED_PM_DESIGN").Value = .Range("TASK_BASED_PM_DESIGN_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_IMPLEMENT").Value
                .Range("TASK_BASED_PM_IMPLEMENT").Value = .Range("TASK_BASED_PM_IMPLEMENT_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_TEST").Value
                .Range("TASK_BASED_PM_TEST").Value = .Range("TASK_BASED_PM_TEST_START").Offset(i, 0).Value + tmp
                tmp = .Range("TASK_BASED_PM_SITE").Value
                .Range("TASK_BASED_PM_SITE").Value = .Range("TASK_BASED_PM_SITE_START").Offset(i, 0).Value + tmp
    
            Case Else
             
        End Select
        i = i + 2
    Next
   
    
End With


CleanUp:
Application.Cursor = xlDefault
On Error Resume Next
    Set tmp = Nothing
    Set s = Nothing
    Set wb = Nothing
Exit Function

ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Function



Public Function getWorkBookName() As Boolean
On Error GoTo ErrHandler
'gets the name of the opened GECE sheet workbook
Dim strTemp As String
Dim w As Workbook
Dim wb As Workbooks
Dim s As Worksheet

'clear variable
gstrGECEWorkBook = ""
Set wb = Excel.Workbooks

For Each w In wb
'Debug.Print w.Name
    
    For Each s In w.Worksheets
        'look for the data entry sheet so we know this is a valid GECE sheet.
        If s.Name = gstrGECEDataEntrySheet Then
            'set the workbook name variable
            gstrGECEWorkBook = w.Name
            getWorkBookName = True
            Exit For
        End If
        'Debug.Print s.Name
    Next
Next

If gstrGECEDataEntrySheet = "" Then
    'we did not find the data entry sheet so we must assume the GECE sheet has not been opened
    MsgBox "GECE workbook is not open. Please open a GECE workbook."
    getWorkBookName = False
End If


CleanUp:
On Error Resume Next
    Set s = Nothing
    Set w = Nothing
    Set wb = Nothing
Exit Function
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Function

Sub UpdateSummaryPage(oForm As UserForm)
On Error Resume Next
    'update the summary page
If OpenDataEntryForm = False Then

    With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECECQAOutputSheet)
        '###################
        'set total cost on status bar
        '2005/05/16
        'fix a bug where documentaton tab check boxes are not recalculating total on status fields at bottom of form
        DoEvents
        oForm.txtStatusTotalCost.Value = .Range("TOTAL_COST").Text 'TOTAL_COST
        'set total hours on status bar
        oForm.txtStatusHours.Value = .Range("TOTAL_HOURS").Text 'TOTAL_HOURS
        '#################
        oForm.txttt.Value = Format(.Range("TOTAL_PCT").Value, "0%")
            
        'Total Engineering
        'TOTAL_HOURS
        oForm.txtTOTAL_HOURS.Value = .Range("TOTAL_HOURS").Text
        'TOTAL_ENG_COST
        oForm.txtTOTAL_COST.Value = .Range("TOTAL_COST").Text
        'TOTAL_ENG_LIST
        oForm.txtTOTAL_LIST.Value = .Range("TOTAL_LIST").Text
        'TOTAL_ENG_PCT
        oForm.txtTOTAL_PCT.Value = Format(.Range("TOTAL_PCT").Value, "0%")
        
        'Local Engineering
        oForm.txtTOTAL_LOC_HOURS.Value = .Range("TOTAL_LOC_HOURS").Text
        'LOC_TOTAL_COST
        oForm.txtTOTAL_LOC_COST.Value = .Range("TOTAL_LOC_COST").Text
        'LOC_TOTAL_LIST
        oForm.txtTOTAL_LOC_LIST.Value = .Range("TOTAL_LOC_LIST").Text
        'TOTAL_ENG_LOC_PCT
        oForm.txtTOTAL_LOC_PCT.Value = Format(.Range("TOTAL_LOC_PCT").Value, "0%")
        
        'Remote Engineering
        oForm.txtTOTAL_REM_HOURS.Value = Format(Round(.Range("TOTAL_REM_HOURS").Text), "#,##0")
        'REM_TOTAL_COST
        oForm.txtTOTAL_REM_COST.Value = .Range("TOTAL_REM_COST").Text
        'REM_TOTAL_LIST
        oForm.txtTOTAL_REM_LIST.Value = .Range("TOTAL_REM_LIST").Text
        'TOTAL_ENG_REM_PCT
        oForm.txtTOTAL_REM_PCT.Value = Format(.Range("TOTAL_REM_PCT").Value, "0%")
        
        't&l
        oForm.txtTL_COST.Value = .Range("TOTAL_TL_COST").Text
        oForm.txtTL_LIST.Value = .Range("TOTAL_TL_LIST").Text
      
      'List values
        'APP_BASED_LIST
        oForm.txtAPP_BASED_TOTAL_LIST.Value = .Range("APP_BASED_TOTAL_LIST").Text
        oForm.txtAPP_BASED_LOC_LIST.Value = .Range("APP_BASED_LOC_LIST").Text
        oForm.txtAPP_BASED_REM_LIST.Value = .Range("APP_BASED_REM_LIST").Text
        oForm.txtAPP_BASED_TOTAL_CURRENCY.Value = .Range("APP_BASED_REM_LIST").Text
       'TASK_BASED_LIST
        oForm.txtTASK_BASED_TOTAL_LIST.Value = .Range("TASK_BASED_TOTAL_LIST").Text
        oForm.txtTASK_BASED_LOC_LIST.Value = .Range("TASK_BASED_LOC_LIST").Text
        oForm.txtTASK_BASED_REM_LIST.Value = .Range("TASK_BASED_REM_LIST").Text
        'DURATION_BASED_LIST
        oForm.txtDURATION_BASED_TOTAL_LIST.Value = .Range("DURATION_BASED_TOTAL_LIST").Text
        oForm.txtDURATION_BASED_LOC_LIST.Value = .Range("DURATION_BASED_LOC_LIST").Text
        oForm.txtDURATION_BASED_REM_LIST.Value = .Range("DURATION_BASED_REM_LIST").Text
        oForm.txtDURATION_BASED_FREE_LIST.Value = .Range("DURATION_BASED_FREE_LIST").Text
        
     'Unit Values
        'First Unit
        oForm.txt1st_UNIT_TOTAL_HOURS.Value = .Range("TOTAL_1st_UNIT_TOTAL_HOURS").Text
        oForm.txt1st_UNIT_TOTAL_COST.Value = .Range("TOTAL_1st_UNIT_TOTAL_COST").Text
        oForm.txt1st_UNIT_TOTAL_LIST.Value = .Range("TOTAL_1st_UNIT_TOTAL_LIST").Text
        oForm.txt1st_UNIT_TOTAL_PCT.Value = .Range("TOTAL_1st_UNIT_TOTAL_PCT").Text
        
        'Second Unit
        oForm.txt2nd_UNIT_TOTAL_HOURS.Value = .Range("TOTAL_2nd_UNIT_TOTAL_HOURS").Text
        oForm.txt2nd_UNIT_TOTAL_COST.Value = .Range("TOTAL_2nd_UNIT_TOTAL_COST").Text
        oForm.txt2nd_UNIT_TOTAL_LIST.Value = .Range("TOTAL_2nd_UNIT_TOTAL_LIST").Text
        oForm.txt2nd_UNIT_TOTAL_PCT.Value = .Range("TOTAL_2nd_UNIT_TOTAL_PCT").Text
       
    
    End With
    
    With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEApplicationBasedSheet)
        'Total APP based Engineering
        oForm.txtAPP_BASED_TOTAL_HOURS.Value = Round(.Range("APP_BASED_TOTAL_HOURS").Text)
        'TOTAL_APP_COST
        oForm.txtAPP_BASED_TOTAL_COST.Value = .Range("APP_BASED_TOTAL_COST").Text
        'TOTAL_APP_PCT
        oForm.txtAPP_BASED_TOTAL_PCT.Value = Format(.Range("APP_BASED_TOTAL_PCT").Value, "0%")
    
        'LOC APP based Engineering
        oForm.txtAPP_BASED_LOC_HOURS.Value = Round(.Range("APP_BASED_LOC_HOURS").Text)
        'LOC_APP_COST
        oForm.txtAPP_BASED_LOC_COST.Value = .Range("APP_BASED_LOC_COST").Text
        'LOC_APP_PCT
        oForm.txtAPP_BASED_LOC_PCT.Value = Format(.Range("APP_BASED_LOC_PCT").Value, "0%")
    
        'REM APP based Engineering
        oForm.txtAPP_BASED_REM_HOURS.Value = Round(.Range("APP_BASED_REM_HOURS").Text)
        'REM_APP_COST
        oForm.txtAPP_BASED_REM_COST.Value = .Range("APP_BASED_REM_COST").Text
        'REM_APP_PCT
        oForm.txtAPP_BASED_REM_PCT.Value = Format(.Range("APP_BASED_REM_PCT").Value, "0%")
    End With
    
    With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECETaskBasedSheet)
        'Total TASK based Engineering
        oForm.txtTASK_BASED_TOTAL_HOURS.Value = Round(.Range("TASK_BASED_TOTAL_HOURS").Text)
        'TOTAL_TASK_COST
        oForm.txtTASK_BASED_TOTAL_COST.Value = .Range("TASK_BASED_TOTAL_COST").Text
        'TOTAL_TASK_PCT
        oForm.txtTASK_BASED_TOTAL_PCT.Value = Format(.Range("TASK_BASED_TOTAL_PCT").Value, "0%")
    
        'LOC TASK based Engineering
        oForm.txtTASK_BASED_LOC_HOURS.Value = Round(.Range("TASK_BASED_LOC_HOURS").Text)
        'LOC_TASK_COST
        oForm.txtTASK_BASED_LOC_COST.Value = .Range("TASK_BASED_LOC_COST").Text
        'LOC_TASK_PCT
        oForm.txtTASK_BASED_LOC_PCT.Value = Format(.Range("TASK_BASED_LOC_PCT").Value, "0%")
    
        'REM TASK based Engineering
        oForm.txtTASK_BASED_REM_HOURS.Value = Round(.Range("TASK_BASED_REM_HOURS").Text)
        'REM_TASK_COST
        oForm.txtTASK_BASED_REM_COST.Value = .Range("TASK_BASED_REM_COST").Text
        'REM_TASK_PCT
        oForm.txtTASK_BASED_REM_PCT.Value = Format(.Range("TASK_BASED_REM_PCT").Value, "0%")
    End With
    
    With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDurationBasedSheet)
        'Total DURATION based Engineering txtDURATION_BASED_TOTAL_HOURS
         oForm.txtDURATION_BASED_TOTAL_HOURS.Value = Round(.Range("DURATION_TOTAL_HOURS").Text)
        'TOTAL_DURATION_COST
        oForm.txtDURATION_BASED_TOTAL_COST.Value = .Range("DURATION_TOTAL_COST").Text
        'TOTAL_DURATION_PCT
        oForm.txtDURATION_BASED_TOTAL_PCT.Value = Format(.Range("DURATION_BASED_TOTAL_PCT").Value, "0%")
    
        'LOC DURATION based Engineering
        oForm.txtDURATION_BASED_LOC_HOURS.Value = Round(.Range("DURATION_TOTAL_LOC_HOURS").Text)
        'LOC_DURATION_COST
        oForm.txtDURATION_BASED_LOC_COST.Value = .Range("DURATION_TOTAL_LOC_COST").Text
        'LOC_DURATION_PCT
        oForm.txtDURATION_BASED_LOC_PCT.Value = Format(.Range("DURATION_BASED_LOC_PCT").Value, "0%")
    
        'REM DURATION based Engineering
        oForm.txtDURATION_BASED_REM_HOURS.Value = Round(.Range("DURATION_TOTAL_REM_HOURS").Text)
        'REM_DURATION_COST
        oForm.txtDURATION_BASED_REM_COST.Value = .Range("DURATION_TOTAL_REM_COST").Text
        'REM_DURATION_PCT
        oForm.txtDURATION_BASED_REM_PCT.Value = Format(.Range("DURATION_BASED_REM_PCT").Value, "0%")
        
    End With
    
    With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEApplicationBasedSheet)
        '''''''
        'metrics section
        'Total Number of Equivalent IO
        oForm.txtTOT_IO_COUNT.Value = .Range("TOT_IO_COUNT").Text
        'Total Hours per point
        oForm.txtTOTAL_HOURS_PER_POINT.Value = .Range("TOTAL_HOURS_PER_POINT").Text
        oForm.txtHrs.Value = "hrs"
        
        'RATIO of CP  points
        oForm.txtRATIO_CP.Value = .Range("RATIO_CP").Text
        'RATIO of DI  points
        oForm.txtRATIO_DI.Value = .Range("RATIO_DI").Text
        'RATIO of ESD  points
        oForm.txtRATIO_ESD.Value = .Range("RATIO_ESD").Text
    
        'TOTAL_COST_PER_POINT
        oForm.txtTOTAL_COST_PER_POINT.Value = Format(Round(.Range("TOTAL_COST_PER_POINT").Text), "#,##0")
        'TOTAL_APP_HARDWARE_COST_PER_POINT
        oForm.txtTOTAL_APP_HARDWARE_COST_PER_POINT.Value = Format(Round(.Range("TOTAL_APP_HARDWARE_COST_PER_POINT").Text), "#,##0")
        
         'TOTAL_PCT_HW_ENG
        oForm.txtTOTAL_PCT_HW_ENG.Value = .Range("TOTAL_PCT_HW_ENG").Text
       
    End With
    
    With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECECostSheet)
        'summary
        oForm.txtAPP_BASED_TOTAL_CURRENCY.Value = .Range("LOC_CURRENCY").Text
        oForm.txtTASK_BASED_TOTAL_CURRENCY.Value = .Range("LOC_CURRENCY").Text
        oForm.txtDURATION_BASED_TOTAL_CURRENCY.Value = .Range("LOC_CURRENCY").Text
        oForm.txtAPP_BASED_LOC_CURRENCY.Value = .Range("LOC_CURRENCY").Text
        oForm.txtTASK_BASED_LOC_CURRENCY.Value = .Range("LOC_CURRENCY").Text
        oForm.txtDURATION_BASED_LOC_CURRENCY.Value = .Range("LOC_CURRENCY").Text
        oForm.txtAPP_BASED_REM_CURRENCY.Value = .Range("LOC_CURRENCY").Text
        oForm.txtTASK_BASED_REM_CURRENCY.Value = .Range("LOC_CURRENCY").Text
        oForm.txtDURATION_BASED_REM_CURRENCY.Value = .Range("LOC_CURRENCY").Text
        oForm.txtDURATION_BASED_FREE_CURRENCY.Value = .Range("LOC_CURRENCY").Text
        oForm.txtTOTAL_CURRENCY.Value = .Range("LOC_CURRENCY").Text
        oForm.txtTOTAL_LOC_CURRENCY.Value = .Range("LOC_CURRENCY").Text
        oForm.txtTOTAL_REM_CURRENCY.Value = .Range("LOC_CURRENCY").Text
        oForm.txtTL_CURRENCY.Value = .Range("LOC_CURRENCY").Text
        oForm.txt1st_UNIT_TOTAL_CURRENCY.Value = .Range("LOC_CURRENCY").Text
        oForm.txt2nd_UNIT_TOTAL_CURRENCY.Value = .Range("LOC_CURRENCY").Text
        oForm.txtCOST_PER_POINT_CURRENCY.Value = .Range("LOC_CURRENCY").Text
        
        'metrics
        oForm.txtMetricsCurrency1.Value = .Range("LOC_CURRENCY").Text
        oForm.txtMetricsCurrency2.Value = .Range("LOC_CURRENCY").Text
        
        'set status bar currency
        'txtStatusCurrency
        oForm.txtStatusCurrency.Value = .Range("LOC_CURRENCY").Text
    End With
End If

End Sub

Public Function UpdateRemoteCountry()
On Error Resume Next
Dim i As Integer

' Modified to allow setting directly from the worksheet
'loop through all country fields and set default
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEApplicationBasedSheet)
    'SPEC
    .Range("SPEC_TASK1_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("SPEC_TASK2_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("SPEC_TASK3_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("SPEC_TASK4_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("SPEC_TASK5_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("SPEC_TASK6_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("SPEC_TASK7_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("SPEC_TASK8_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("SPEC_TASK9_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("SPEC_TASK10_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    
    'SYSENG
    .Range("SYSENG_TASK1_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("SYSENG_TASK2_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("SYSENG_TASK3_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("SYSENG_TASK4_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("SYSENG_TASK5_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("SYSENG_TASK6_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("SYSENG_TASK7_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("SYSENG_TASK8_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("SYSENG_TASK9_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("SYSENG_TASK10_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    
    'HMI
    .Range("HMI_TASK1_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("HMI_TASK2_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("HMI_TASK3_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("HMI_TASK4_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("HMI_TASK5_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("HMI_TASK6_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("HMI_TASK7_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("HMI_TASK8_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("HMI_TASK9_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("HMI_TASK10_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    
    'CP
    .Range("CP_TASK1_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("CP_TASK2_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("CP_TASK3_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("CP_TASK4_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("CP_TASK5_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("CP_TASK6_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("CP_TASK7_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("CP_TASK8_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("CP_TASK9_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("CP_TASK10_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    
    'DI
    .Range("DI_TASK1_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("DI_TASK2_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("DI_TASK3_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("DI_TASK4_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("DI_TASK5_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("DI_TASK6_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("DI_TASK7_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("DI_TASK8_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("DI_TASK9_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("DI_TASK10_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    
    'ESD
    .Range("ESD_TASK1_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("ESD_TASK2_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("ESD_TASK3_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("ESD_TASK4_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("ESD_TASK5_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("ESD_TASK6_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("ESD_TASK7_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("ESD_TASK8_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("ESD_TASK9_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("ESD_TASK10_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    
    'REP
    .Range("REP_TASK1_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("REP_TASK2_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("REP_TASK3_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("REP_TASK4_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("REP_TASK5_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("REP_TASK6_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("REP_TASK7_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("REP_TASK8_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("REP_TASK9_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("REP_TASK10_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    
    'APP
    .Range("APP_TASK1_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("APP_TASK2_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("APP_TASK3_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("APP_TASK4_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("APP_TASK5_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("APP_TASK6_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("APP_TASK7_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("APP_TASK8_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("APP_TASK9_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("APP_TASK10_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    
    'TEST
    .Range("TEST_TASK1_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("TEST_TASK2_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("TEST_TASK3_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("TEST_TASK4_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("TEST_TASK5_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("TEST_TASK6_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("TEST_TASK7_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("TEST_TASK8_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("TEST_TASK9_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("TEST_TASK10_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    
    'DOC
    .Range("DOC_TASK1_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("DOC_TASK2_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("DOC_TASK3_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("DOC_TASK4_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("DOC_TASK5_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("DOC_TASK6_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("DOC_TASK7_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("DOC_TASK8_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("DOC_TASK9_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("DOC_TASK10_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    
    'COURSE
    .Range("COURSE_TASK1_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("COURSE_TASK2_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("COURSE_TASK3_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("COURSE_TASK4_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("COURSE_TASK5_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("COURSE_TASK6_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("COURSE_TASK7_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("COURSE_TASK8_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("COURSE_TASK9_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("COURSE_TASK10_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    
    'PM
    .Range("PM_TASK1_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("PM_TASK2_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    '.Range("PM_TASK3_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("PM_TASK4_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("PM_TASK5_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("PM_TASK6_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("PM_TASK7_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("PM_TASK8_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("PM_TASK9_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("PM_TASK10_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    
    'MEETING
    .Range("MEETING_TASK1_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("MEETING_TASK2_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("MEETING_TASK3_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("MEETING_TASK4_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("MEETING_TASK5_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("MEETING_TASK6_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("MEETING_TASK7_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("MEETING_TASK8_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("MEETING_TASK9_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("MEETING_TASK10_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    
    'SITE
    .Range("SITE_TASK1_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("SITE_TASK2_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("SITE_TASK3_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("SITE_TASK4_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("SITE_TASK5_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("SITE_TASK6_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("SITE_TASK7_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("SITE_TASK8_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("SITE_TASK9_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("SITE_TASK10_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    
    'MEETING
    .Range("CLEAN_TASK1_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("CLEAN_TASK2_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("CLEAN_TASK3_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("CLEAN_TASK4_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("CLEAN_TASK5_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("CLEAN_TASK6_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("CLEAN_TASK7_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("CLEAN_TASK8_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("CLEAN_TASK9_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("CLEAN_TASK10_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    
    'TL
    .Range("TL_TASK1_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("TL_TASK2_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("TL_TASK3_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("TL_TASK4_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("TL_TASK5_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("TL_TASK6_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("TL_TASK7_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("TL_TASK8_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("TL_TASK9_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    .Range("TL_TASK10_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
End With


'some of the named ranges changed????
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDurationBasedSheet)
    .Range("DURATION_PM_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value 'DURATION_PM_REM_CURRENCY
    .Range("DURATION_AC_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value 'DURATION_AC_REM_CURRENCY
    .Range("DURATION_LE_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value 'DURATION_LE_REM_CURRENCY
    .Range("DURATION_AE_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value 'DURATION_AE_REM_CURRENCY
    .Range("DURATION_T_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value 'DURATION_T_REM_CURRENCY '
    .Range("DURATION_SE_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value 'DURATION_SE_REM_CURRENCY '
    .Range("DURATION_SR_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value 'DURATION_SR_REM_CURRENCY '
    .Range("DURATION_TOTAL_BUSINESS_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value 'DURATION_TOTAL_BUSINESS_REM_COUNTRY
    .Range("DURATION_TOTAL_STAGING_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value 'DURATION_TOTAL_STAGING_REM_COUNTRY
    .Range("DURATION_TOTAL_SITE_REM_COUNTRY").Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value 'DURATION_TOTAL_SITE_REM_COUNTRY
End With

With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECETaskBasedSheet)
    For i = 1 To 100
    .Range("TASK_BASED_REM_COUNTRY_" & i).Value = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("DEFAULT_REM_COUNTRY").Value
    Next
End With


End Function

'all ranges with _UNITS in them are for setting labels on the form to equal the ones in the data entry sheet.
'so if the unit cell on the sheet changes they will be reflected on the form.
Sub SetUnitLabels(oForm As UserForm)
On Error Resume Next
'data entry
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    'example using offset from dataentry field
    'oForm.lblTL_DAYS_QTY_UNITS.Caption = .Range("TL_DAYS_QTY").Offset(0, 1).Text
    'system
    
    'Dan 2006-11-05 labels are now used by both ia and ESD so don't set
    'Andre will need to decide on the lable names.
    'oForm.lblSYS_WORKSTATIONS_QTY_UNITS.Caption = .Range("SYS_WORKSTATIONS_QTY_UNITS").Text
    'oForm.lblSYS_CONTROLLERS_QTY_UNITS.Caption = .Range("SYS_CONTROLLERS_QTY_UNITS").Text
    'oForm.lblSYS_FBM_QTY_UNITS.Caption = .Range("SYS_FBM_QTY_UNITS").Text
    'oForm.lblSYS_FDSI_QTY_UNITS.Caption = .Range("SYS_FDSI_QTY_UNITS").Text
    
    oForm.lblCAB_PROC_QTY_UNITS.Caption = .Range("CAB_PROC_QTY_UNITS").Text
    oForm.lblCAB_IO_QTY_UNITS.Caption = .Range("CAB_IO_QTY_UNITS").Text
    oForm.lblCAB_MARSH_QTY_UNITS.Caption = .Range("CAB_MARSH_QTY_UNITS").Text
    oForm.lblCAB_CONSOLES_QTY_UNITS.Caption = .Range("CAB_CONSOLES_QTY_UNITS").Text
    oForm.lblCOST_IA_UNITS.Caption = .Range("COST_IA_UNITS").Text
    oForm.lblCOST_BUYOUT_UNI.Caption = .Range("COST_BUYOUT_UNI").Text
    oForm.lblDURATION_UNITS.Caption = .Range("DURATION_UNITS").Text
    'Travel & Living
    oForm.lblTL_TRIPS_REQ_UNITS.Caption = .Range("TL_TRIPS_REQ_UNITS").Text
    oForm.lblTL_PERS_QTY_UNITS.Caption = .Range("TL_PERS_QTY_UNITS").Text
    oForm.lblTL_DAYS_QTY_UNITS.Caption = .Range("TL_DAYS_QTY_UNITS").Text
    oForm.lblTL_AIRFARE_UNITS.Caption = .Range("TL_AIRFARE_UNITS").Text
    oForm.lblTL_DAILY_ALLOW_UNITS.Caption = .Range("TL_DAILY_ALLOW_UNITS").Text
    oForm.lblTL_AIRFARE_UNITS2.Caption = .Range("TL_AIRFARE_UNITS").Text
    oForm.lblTL_DAILY_ALLOW_UNITS2.Caption = .Range("TL_DAILY_ALLOW_UNITS").Text
    'meetings
    oForm.lblMEETING_KICKOFF_UNITS.Caption = .Range("MEETING_KICKOFF_UNITS").Text
    oForm.lblMEETING_DESIGN_UNITS.Caption = .Range("MEETING_DESIGN_UNITS").Text
    oForm.lblMEETING_PROGRESS_UNITS.Caption = .Range("MEETING_PROGRESS_UNITS").Text
    oForm.lblMEETING_OTHER_UNITS.Caption = .Range("MEETING_OTHER_UNITS").Text
    oForm.lblMEETING_CLOSE_UNITS.Caption = .Range("MEETING_CLOSE_UNITS").Text
    'Site Services
    oForm.lblSITE_SURVEY_HOURS_UNITS.Caption = .Range("SITE_SURVEY_HOURS_UNITS").Text
    oForm.lblSITE_PWRUP_HOURS_UNITS.Caption = .Range("SITE_PWRUP_HOURS_UNITS").Text
    oForm.lblSITE_COMM_HOURS_UNITS.Caption = .Range("SITE_COMM_HOURS_UNITS").Text
    oForm.lblSITE_SAT_HOURS_UNITS.Caption = .Range("SITE_SAT_HOURS_UNITS").Text
    'report
    oForm.lblREP_STD_UNITS.Caption = .Range("REP_STD_UNITS").Text
    oForm.lblREP_CUSTOM_UNITS.Caption = .Range("REP_CUSTOM_UNITS").Text
    oForm.lblREP_MASS_HEAT_UNITS.Caption = .Range("REP_MASS_HEAT_UNITS").Text
    'testing
    oForm.lblTEST_PRE_FAT_PCT_UNITS.Caption = .Range("TEST_PRE_FAT_PCT_UNITS").Text
    oForm.lblTEST_FAT_PCT_UNITS.Caption = .Range("TEST_FAT_PCT_UNITS").Text
    oForm.lblTEST_CUSTOMER_FAT_UNITS.Caption = .Range("TEST_CUSTOMER_FAT_UNITS").Text
    oForm.lblRENTAL_COST_UNITS.Caption = .Range("RENTAL_COST_UNITS").Text
    'di
    oForm.lblDI_DEVICES_UNITS.Caption = .Range("DI_DEVICES_UNITS").Text
    oForm.lblDI_INTERFACES_UNITS.Caption = .Range("DI_INTERFACES_UNITS").Text
    oForm.lblDI_COMPLEX_UNITS.Caption = .Range("DI_COMPLEX_UNITS").Text
    oForm.lblDI_SEQ_LOOP_UNITS.Caption = .Range("DI_SEQ_LOOP_UNITS").Text
    oForm.lblDI_SEQ_COMPLEX_UNITS.Caption = .Range("DI_SEQ_COMPLEX_UNITS").Text
    oForm.lblDI_GRP_START_LOOP_UNITS.Caption = .Range("DI_GRP_START_LOOP_UNITS").Text
    oForm.lblDI_GRP_START_COMPLEX_UNITS.Caption = .Range("DI_GRP_START_COMPLEX_UNITS").Text
    'cp
    oForm.lblCP_ANA_COMPLEX_UNITS.Caption = .Range("CP_ANA_COMPLEX_UNITS").Text
    oForm.lblCP_DIGITAL_COMPLEX_UNITS.Caption = .Range("CP_DIGITAL_COMPLEX_UNITS").Text
    oForm.lblCP_FIELDBUS_IO_UNITS.Caption = .Range("CP_FIELDBUS_IO_UNITS").Text
    oForm.lblCP_SEQ_LOOP_UNITS.Caption = .Range("CP_SEQ_LOOP_UNITS").Text
    oForm.lblCP_SEQ_COMPLEX_UNITS.Caption = .Range("CP_SEQ_COMPLEX_UNITS").Text
    oForm.lblCP_GRP_START_LOOP_UNITS.Caption = .Range("CP_GRP_START_LOOP_UNITS").Text
    oForm.lblCP_GRP_START_COMPLEX_UNITS.Caption = .Range("CP_GRP_START_COMPLEX_UNITS").Text
End With
End Sub

Public Sub SetDefaults(oForm As UserForm)
On Error Resume Next
With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet)
    'cp
    oForm.txtCP_DIGITAL_CTRL_DI_REC.Value = .Range("CP_DIGITAL_CTRL_DI_REC").Value
    oForm.txtCP_DIGITAL_CTRL_DO_REC.Value = .Range("CP_DIGITAL_CTRL_DO_REC").Value
    oForm.txtCP_ANA_COMPLEX_REC.Value = Format(.Range("CP_ANA_COMPLEX_REC").Value, "0%")
    oForm.txtCP_DIGITAL_COMPLEX_REC.Value = Format(.Range("CP_DIGITAL_COMPLEX_REC").Value, "0%")
    oForm.txtCP_FIELDBUS_IO_REC.Value = Format(.Range("CP_FIELDBUS_IO_REC").Value, "0%")
    oForm.txtCP_SEQ_LOOP_REC.Value = Format(.Range("CP_SEQ_LOOP_REC").Value, "0%")
    oForm.txtCP_SEQ_COMPLEX_REC.Value = Format(.Range("CP_SEQ_COMPLEX_REC").Value, "0%")
    oForm.txtCP_GRP_START_LOOP_REC.Value = Format(.Range("CP_GRP_START_LOOP_REC").Value, "0%")
    oForm.txtCP_GRP_START_COMPLEX_REC.Value = Format(.Range("CP_GRP_START_COMPLEX_REC").Value, "0%")
    'di
    oForm.txtDI_DIGITAL_CTRL_DI_REC.Value = .Range("DI_DIGITAL_CTRL_DI_REC").Value
    oForm.txtDI_DIGITAL_CTRL_DO_REC.Value = .Range("DI_DIGITAL_CTRL_DO_REC").Value
    oForm.txtDI_COMPLEX_REC.Value = Format(.Range("DI_COMPLEX_REC").Value, "0%")
    oForm.txtDI_SEQ_LOOP_REC.Value = Format(.Range("DI_SEQ_LOOP_REC").Value, "0%")
    oForm.txtDI_SEQ_COMPLEX_REC.Value = Format(.Range("DI_SEQ_COMPLEX_REC").Value, "0%")
    oForm.txtDI_GRP_START_LOOP_REC.Value = Format(.Range("DI_GRP_START_LOOP_REC").Value, "0%")
    oForm.txtDI_GRP_START_COMPLEX_REC.Value = Format(.Range("DI_GRP_START_COMPLEX_REC").Value, "0%")
    'esd
    oForm.txtESD_COMPLEX_REC.Value = Format(.Range("ESD_COMPLEX_REC").Value, "0%")
    oForm.txtESD_MISC_CAB_REC.Value = Format(.Range("ESD_MISC_CAB_REC").Value, "0.00%")
    oForm.txtESD_GRP_START_LOOP_REC.Value = Format(.Range("ESD_GRP_START_LOOP_REC").Value, "0%")
    oForm.txtESD_GRP_START_COMPLEX_REC.Value = Format(.Range("ESD_GRP_START_COMPLEX_REC").Value, "0%")
    'frmComplete.SetESD_HMI_REQ
    'frmComplete.SetESD_MARSH_CAB_REQ
    'frmComplete.SetESD_SYSTEM_REQ
    'Test
    oForm.txtTEST_PRE_FAT_REC.Value = Format(.Range("TEST_PRE_FAT_REC").Value, "0%")
    oForm.txtTEST_FAT_REC.Value = Format(.Range("TEST_FAT_REC").Value, "0%")
    
   
   'cell refs a named range on another sheet
    If Not IsError(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("TEST_TASK1_QTY").Value) Then
        oForm.txtTEST_FAT_QTY.Value = Round(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("TEST_TASK1_QTY").Value)
    End If

'meeting
    oForm.txtMEETING_KICKOFF_EST.Value = .Range("MEETING_KICKOFF_EST").Value
    oForm.txtMEETING_DESIGN_EST.Value = .Range("MEETING_DESIGN_EST").Value
    oForm.txtMEETING_PROGRESS_EST.Value = .Range("MEETING_PROGRESS_EST").Value
    oForm.txtMEETING_CLOSE_EST.Value = .Range("MEETING_CLOSE_EST").Value
    'system summary
    oForm.txtSYS_WORKSTATIONS_EST.Value = .Range("SYS_WORKSTATIONS_EST").Value
    oForm.txtSYS_CONTROLLERS_EST.Value = .Range("SYS_CONTROLLERS_EST").Value
    oForm.txtSYS_FBM_EST.Value = .Range("SYS_FBM_EST").Value
    oForm.txtSYS_FDSI_EST.Value = .Range("SYS_FDSI_EST").Value
    oForm.txtESD_SYSTEMS_EST.Value = .Range("ESD_SYSTEMS_EST").Value
    oForm.txtESD_CHASSIS_EST.Value = .Range("ESD_CHASSIS_EST").Value
    oForm.txtESD_IO_CARD_EST.Value = .Range("ESD_IO_CARD_EST").Value
    oForm.txtESD_COMM_EST.Value = .Range("ESD_COMM_EST").Value
    oForm.txtCAB_PROC_EST.Value = .Range("CAB_PROC_EST").Value
    oForm.txtCAB_IO_EST.Value = .Range("CAB_IO_EST").Value
    oForm.txtCAB_MARSH_EST.Value = .Range("CAB_MARSH_EST").Value
    oForm.txtCAB_CONSOLES_EST.Value = .Range("CAB_CONSOLES_EST").Value
    oForm.txtESD_CAB_PROC_EST.Value = .Range("ESD_CAB_PROC_EST").Value
    oForm.txtESD_CAB_IO_EST.Value = .Range("ESD_CAB_IO_EST").Value
    oForm.txtESD_CAB_MARSH_EST.Value = .Range("ESD_CAB_MARSH_EST").Value
    oForm.txtDURATION_RECOMMENDED.Value = Format(.Range("DURATION_REC").Value, "0,0")
    oForm.txtDATE_START_RECOMMENDED.Value = .Range("DATE_START_RECOMMENDED").Value
    oForm.txtDATE_END_RECOMMENDED.Value = .Range("DATE_END_RECOMMENDED").Value
    'report
    oForm.txtREP_STD_EST.Value = .Range("REP_STD_EST").Value
    'site services
    oForm.txtSITE_SAT_EST.Value = .Range("SITE_SAT_EST").Value
    oForm.txtTL_PERS_QTY_FAT_REC.Value = .Range("TL_PERS_QTY_FAT_REC").Value
    oForm.txtFAT_NB_DAY_REC.Value = .Range("FAT_NB_DAY_REC").Value
    oForm.txtTL_PERS_QTY_SITE_REC.Value = .Range("TL_PERS_QTY_SITE_REC").Value
    oForm.txtSITE_NB_DAY_REC.Value = .Range("SITE_NB_DAY_REC").Value
    oForm.txtTL_TRIPS_REQ_SITE_REC.Value = .Range("TL_TRIPS_REQ_SITE_REC").Value
    oForm.txtTL_TRIPS_REQ_FAT_REC.Value = .Range("TL_TRIPS_REQ_FAT_REC").Value
    
    oForm.txtLONG_PCT_EST.Value = Format(.Range("LONG_PCT_EST").Value, "0%")
    oForm.txtLONG_PCT_EST_ESD.Value = Format(.Range("LONG_PCT_EST_ESD").Value, "0%")
    oForm.txtMarshWiring.Value = Format(.Range("MarshWiring_REQ").Value, "0")
    oForm.txtMarshWiring_ESD.Value = Format(.Range("MarshWiring_ESD_REQ").Value, "0")
    
    oForm.txtCABINET_VENDOR_REC.Value = .Range("CABINET_VENDOR_REC").Value
    
End With

'testing
 oForm.txtTEST_FAT_QTY.Value = Round(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("TEST_TASK1_QTY").Value)
 frmComplete.SetSW_SIMULATOR_REQ

End Sub

Public Sub GetToolTips(oForm As UserForm)
On Error Resume Next
Dim MyControl As Control
Dim strTemp As String

    For Each MyControl In oForm.Controls
        If Left(MyControl.Name, 3) = "txt" Or Left(MyControl.Name, 3) = "ckb" Then
            'MyControl.ControlTipText = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range(MyControl.ControlSource).Comment.Text
            
            strTemp = mid(MyControl.Name, 4, Len(MyControl.Name) - 3)
            'don't know which sheet it is from so try all, on error will skip it if not found
            MyControl.ControlTipText = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEDataEntrySheet).Range(strTemp).Comment.Text
            MyControl.ControlTipText = Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range(strTemp).Comment.Text
        End If
    Next
End Sub

Private Sub DocumentFields()
'loops through all fields on the form and gets the label txtbox names and what tab they are on also name ranges
On Error Resume Next
Dim MyControl As Control
'Dim strTemp As String
Dim oForm As UserForm
Set oForm = frmComplete
    For Each MyControl In oForm.Controls
        'Tab Name|Label|Control Name|Range Name
        'just get the dark green fields
        If MyControl.BackColor = &H55DB4B Then
            Debug.Print MyControl.Parent.Parent.Caption & "|" & oForm.Controls("lbl" & MyControl.ControlSource).Caption & "|" & MyControl.Name & "|" & MyControl.ControlSource
        End If
        'get the checkboxes
        If Left(MyControl.Name, 3) = "ckb" Then
            Debug.Print MyControl.Parent.Parent.Caption & "|" & MyControl.Caption & "|" & MyControl.Name & "|" & MyControl.ControlSource
        End If
    Next
End Sub



Private Sub Auto_Open()
On Error GoTo ErrHandler

    ThisWorkbook.Sheets("CoverSheet").Select
    ThisWorkbook.Sheets("CoverSheet").Activate
    frmSplash.Show vbModal

CleanUp:
Exit Sub
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
    Resume CleanUp
End Sub

Public Function GetGECEPath() As String

    On Error GoTo Fail
    Dim sh As Object, v, k1 As String, k2 As String
    Set sh = CreateObject("WScript.Shell")

    k1 = "HKLM\SOFTWARE\GECE\GECE\Install\GECEPath"
    v = sh.RegRead(k1)                          'try native hive first
    GetGECEPath = CStr(v)
    Exit Function

Fail:
    Err.Clear
    '32-bit Office on 64-bit Windows stores under WOW6432Node
    k2 = "HKLM\SOFTWARE\WOW6432Node\GECE\GECE\Install\GECEPath"
    On Error Resume Next
    v = CreateObject("WScript.Shell").RegRead(k2)
    If Err.Number = 0 Then
        GetGECEPath = CStr(v)
    Else
        GetGECEPath = ""                        'not found
    End If
End Function

Public Function FormatPercent(vntValue) As Variant
On Error GoTo ErrHandler

If IsNull(vntValue) Or vntValue = 0 Then
    FormatPercent = 0
Else
    FormatPercent = Format(val(vntValue) / 100, "00%")
End If
Exit Function
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
End Function

Public Function ResetRemotePercentage() As Variant
On Error GoTo ErrHandler
    Application.Cursor = xlWait
    With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEApplicationBasedSheet)
'        If Workbooks(gstrGECEWorkBook).Worksheets(gstrGECECostSheet).Range("LOC_REGION") = "EMEA" Then
'            If Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet).Range("SCOPE_LANGUAGE") = "English" Then
                'SPEC
                .Range("SPEC_TASK1_REM_PCT").Value = 0
                .Range("SPEC_TASK2_REM_PCT").Value = 0
                .Range("SPEC_TASK3_REM_PCT").Value = 0
                .Range("SPEC_TASK4_REM_PCT").Value = 0
                .Range("SPEC_TASK5_REM_PCT").Value = 0
                .Range("SPEC_TASK6_REM_PCT").Value = 0
                .Range("SPEC_TASK7_REM_PCT").Value = 0
                .Range("SPEC_TASK8_REM_PCT").Value = 0
                .Range("SPEC_TASK9_REM_PCT").Value = 0
                .Range("SPEC_TASK10_REM_PCT").Value = 0
                
                'SYSENG
                .Range("SYSENG_TASK1_REM_PCT").Value = 0
                .Range("SYSENG_TASK2_REM_PCT").Value = 0
                .Range("SYSENG_TASK3_REM_PCT").Value = 0
                .Range("SYSENG_TASK4_REM_PCT").Value = 0
                .Range("SYSENG_TASK5_REM_PCT").Value = 0
                .Range("SYSENG_TASK6_REM_PCT").Value = 0
                .Range("SYSENG_TASK7_REM_PCT").Value = 0
                .Range("SYSENG_TASK8_REM_PCT").Value = 0
                .Range("SYSENG_TASK9_REM_PCT").Value = 0
                .Range("SYSENG_TASK10_REM_PCT").Value = 0
                
                'HMI
                .Range("HMI_TASK1_REM_PCT").Value = 0.2
                .Range("HMI_TASK2_REM_PCT").Value = 0.8
                .Range("HMI_TASK3_REM_PCT").Value = 0.8
                .Range("HMI_TASK4_REM_PCT").Value = 0.8
                .Range("HMI_TASK5_REM_PCT").Value = 0.8
                .Range("HMI_TASK6_REM_PCT").Value = 0.8
                .Range("HMI_TASK7_REM_PCT").Value = 0.8
                .Range("HMI_TASK8_REM_PCT").Value = 0.8
                .Range("HMI_TASK9_REM_PCT").Value = 0.75
                .Range("HMI_TASK10_REM_PCT").Value = 0.75
                
                'CP
                .Range("CP_TASK1_REM_PCT").Value = 0.2
                .Range("CP_TASK2_REM_PCT").Value = 0.8
                .Range("CP_TASK3_REM_PCT").Value = 0.8
                .Range("CP_TASK4_REM_PCT").Value = 0.8
                .Range("CP_TASK5_REM_PCT").Value = 0.8
                .Range("CP_TASK6_REM_PCT").Value = 0.8
                .Range("CP_TASK7_REM_PCT").Value = 0.7
                .Range("CP_TASK8_REM_PCT").Value = 0.7
                .Range("CP_TASK9_REM_PCT").Value = 0
                .Range("CP_TASK10_REM_PCT").Value = 0
                
                'DI
                .Range("DI_TASK1_REM_PCT").Value = 0.2
                .Range("DI_TASK2_REM_PCT").Value = 0.8
                .Range("DI_TASK3_REM_PCT").Value = 0.8
                .Range("DI_TASK4_REM_PCT").Value = 0.8
                .Range("DI_TASK5_REM_PCT").Value = 0.8
                .Range("DI_TASK6_REM_PCT").Value = 0.8
                .Range("DI_TASK7_REM_PCT").Value = 0.7
                .Range("DI_TASK8_REM_PCT").Value = 0.7
                .Range("DI_TASK9_REM_PCT").Value = 0
                .Range("DI_TASK10_REM_PCT").Value = 0
                
                'ESD
                .Range("ESD_TASK1_REM_PCT").Value = 0.2
                .Range("ESD_TASK2_REM_PCT").Value = 0.8
                .Range("ESD_TASK3_REM_PCT").Value = 0.8
                .Range("ESD_TASK4_REM_PCT").Value = 0.8
                .Range("ESD_TASK5_REM_PCT").Value = 0.8
                .Range("ESD_TASK6_REM_PCT").Value = 0.8
                .Range("ESD_TASK7_REM_PCT").Value = 0.5
                .Range("ESD_TASK8_REM_PCT").Value = 0
                .Range("ESD_TASK9_REM_PCT").Value = 0
                .Range("ESD_TASK10_REM_PCT").Value = 0
                
                'REP
                .Range("REP_TASK1_REM_PCT").Value = 0
                .Range("REP_TASK2_REM_PCT").Value = 0
                .Range("REP_TASK3_REM_PCT").Value = 0
                .Range("REP_TASK4_REM_PCT").Value = 0
                .Range("REP_TASK5_REM_PCT").Value = 0
                .Range("REP_TASK6_REM_PCT").Value = 0
                .Range("REP_TASK7_REM_PCT").Value = 0
                .Range("REP_TASK8_REM_PCT").Value = 0
                .Range("REP_TASK9_REM_PCT").Value = 0
                .Range("REP_TASK10_REM_PCT").Value = 0
                
                'TEST
                .Range("TEST_TASK1_REM_PCT").Value = 0
                .Range("TEST_TASK2_REM_PCT").Value = 0
                .Range("TEST_TASK3_REM_PCT").Value = 0
                .Range("TEST_TASK4_REM_PCT").Value = 0
                .Range("TEST_TASK5_REM_PCT").Value = 0
                .Range("TEST_TASK6_REM_PCT").Value = 0
                .Range("TEST_TASK7_REM_PCT").Value = 0
                .Range("TEST_TASK8_REM_PCT").Value = 0
                .Range("TEST_TASK9_REM_PCT").Value = 0
                .Range("TEST_TASK10_REM_PCT").Value = 0
                
                'DOC
                .Range("DOC_TASK1_REM_PCT").Value = 0
                .Range("DOC_TASK2_REM_PCT").Value = 0
                .Range("DOC_TASK3_REM_PCT").Value = 0.75
                .Range("DOC_TASK4_REM_PCT").Value = 0.75
                .Range("DOC_TASK5_REM_PCT").Value = 0.75
                .Range("DOC_TASK6_REM_PCT").Value = 0.75
                .Range("DOC_TASK7_REM_PCT").Value = 0.75
                .Range("DOC_TASK8_REM_PCT").Value = 0.65
                .Range("DOC_TASK9_REM_PCT").Value = 0
                .Range("DOC_TASK10_REM_PCT").Value = 1
                
                'COURSE
                .Range("COURSE_TASK1_REM_PCT").Value = 0
                .Range("COURSE_TASK2_REM_PCT").Value = 0
                .Range("COURSE_TASK3_REM_PCT").Value = 0
                .Range("COURSE_TASK4_REM_PCT").Value = 0
                .Range("COURSE_TASK5_REM_PCT").Value = 0
                .Range("COURSE_TASK6_REM_PCT").Value = 0
                .Range("COURSE_TASK7_REM_PCT").Value = 0
                .Range("COURSE_TASK8_REM_PCT").Value = 0
                .Range("COURSE_TASK9_REM_PCT").Value = 0
                .Range("COURSE_TASK10_REM_PCT").Value = 0
                
                'PM
                .Range("PM_TASK1_REM_PCT").Value = 0
                .Range("PM_TASK2_REM_PCT").Value = 0
                .Range("PM_TASK3_REM_PCT").Value = 0
                .Range("PM_TASK4_REM_PCT").Value = 0
                .Range("PM_TASK5_REM_PCT").Value = 0
                .Range("PM_TASK6_REM_PCT").Value = 0
                .Range("PM_TASK7_REM_PCT").Value = 0
                .Range("PM_TASK8_REM_PCT").Value = 0
                .Range("PM_TASK9_REM_PCT").Value = 0
                .Range("PM_TASK10_REM_PCT").Value = 0
                
                'MEETING
                .Range("MEETING_TASK1_REM_PCT").Value = 0
                .Range("MEETING_TASK2_REM_PCT").Value = 0
                .Range("MEETING_TASK3_REM_PCT").Value = 0
                .Range("MEETING_TASK4_REM_PCT").Value = 0
                .Range("MEETING_TASK5_REM_PCT").Value = 0
                .Range("MEETING_TASK6_REM_PCT").Value = 0
                .Range("MEETING_TASK7_REM_PCT").Value = 0
                .Range("MEETING_TASK8_REM_PCT").Value = 0
                .Range("MEETING_TASK9_REM_PCT").Value = 0
                .Range("MEETING_TASK10_REM_PCT").Value = 0
                
                'SITE
                .Range("SITE_TASK1_REM_PCT").Value = 0
                .Range("SITE_TASK2_REM_PCT").Value = 0
                .Range("SITE_TASK3_REM_PCT").Value = 0
                .Range("SITE_TASK4_REM_PCT").Value = 0
                .Range("SITE_TASK5_REM_PCT").Value = 0
                .Range("SITE_TASK6_REM_PCT").Value = 0
                .Range("SITE_TASK7_REM_PCT").Value = 0
                .Range("SITE_TASK8_REM_PCT").Value = 0
                .Range("SITE_TASK9_REM_PCT").Value = 0
                .Range("SITE_TASK10_REM_PCT").Value = 0
            
        
    End With
    Application.Cursor = xlDefault
Exit Function
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
End Function

Public Function UpdateRemotePercentage() As Variant
On Error GoTo ErrHandler
    Dim pct_okay As Boolean
    Dim pct_task_tmp As Double
    Dim hours_task_tmp As Double
    Dim pct_change As Double
    Dim ii As Long
    
    If getWorkBookName = True Then
        With Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEApplicationBasedSheet)
            
            If (.Range("TOTAL_APP_BASED_REM_PCT").Value < 0.05) Then
                Call ResetRemotePercentage
            End If
            Application.Cursor = xlWait
            For ii = 1 To 10
                pct_okay = True
                'If .Range("REM_PCT_REC").Value <> 0 Then
                    If .Range("REM_PCT_REC").Value = 1 Then
                        pct_change = 1
                    Else
                        pct_change = .Range("REM_PCT_REC").Value / .Range("TOTAL_APP_BASED_REM_PCT").Value
                    End If
                    
                    'SPEC
                    .Range("SPEC_TASK1_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SPEC_TASK1_TOTAL_HOURS").Value, .Range("SPEC_TASK1_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SPEC_TASK1_REM_PCT").Value, pct_okay)
                    .Range("SPEC_TASK2_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SPEC_TASK2_TOTAL_HOURS").Value, .Range("SPEC_TASK2_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SPEC_TASK2_REM_PCT").Value, pct_okay)
                    .Range("SPEC_TASK3_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SPEC_TASK3_TOTAL_HOURS").Value, .Range("SPEC_TASK3_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SPEC_TASK3_REM_PCT").Value, pct_okay)
                    .Range("SPEC_TASK4_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SPEC_TASK4_TOTAL_HOURS").Value, .Range("SPEC_TASK4_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SPEC_TASK4_REM_PCT").Value, pct_okay)
                    .Range("SPEC_TASK5_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SPEC_TASK5_TOTAL_HOURS").Value, .Range("SPEC_TASK5_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SPEC_TASK5_REM_PCT").Value, pct_okay)
                    .Range("SPEC_TASK6_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SPEC_TASK6_TOTAL_HOURS").Value, .Range("SPEC_TASK6_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SPEC_TASK6_REM_PCT").Value, pct_okay)
                    .Range("SPEC_TASK7_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SPEC_TASK7_TOTAL_HOURS").Value, .Range("SPEC_TASK7_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SPEC_TASK7_REM_PCT").Value, pct_okay)
                    .Range("SPEC_TASK8_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SPEC_TASK8_TOTAL_HOURS").Value, .Range("SPEC_TASK8_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SPEC_TASK8_REM_PCT").Value, pct_okay)
                    .Range("SPEC_TASK9_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SPEC_TASK9_TOTAL_HOURS").Value, .Range("SPEC_TASK9_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SPEC_TASK9_REM_PCT").Value, pct_okay)
                    .Range("SPEC_TASK10_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SPEC_TASK10_TOTAL_HOURS").Value, .Range("SPEC_TASK10_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SPEC_TASK10_REM_PCT").Value, pct_okay)
                    
                    'SYSENG
                    .Range("SYSENG_TASK1_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SYSENG_TASK1_TOTAL_HOURS").Value, .Range("SYSENG_TASK1_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SYSENG_TASK1_REM_PCT").Value, pct_okay)
                    .Range("SYSENG_TASK2_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SYSENG_TASK2_TOTAL_HOURS").Value, .Range("SYSENG_TASK2_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SYSENG_TASK2_REM_PCT").Value, pct_okay)
                    .Range("SYSENG_TASK3_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SYSENG_TASK3_TOTAL_HOURS").Value, .Range("SYSENG_TASK3_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SYSENG_TASK3_REM_PCT").Value, pct_okay)
                    .Range("SYSENG_TASK4_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SYSENG_TASK4_TOTAL_HOURS").Value, .Range("SYSENG_TASK4_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SYSENG_TASK4_REM_PCT").Value, pct_okay)
                    .Range("SYSENG_TASK5_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SYSENG_TASK5_TOTAL_HOURS").Value, .Range("SYSENG_TASK5_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SYSENG_TASK5_REM_PCT").Value, pct_okay)
                    .Range("SYSENG_TASK6_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SYSENG_TASK6_TOTAL_HOURS").Value, .Range("SYSENG_TASK6_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SYSENG_TASK6_REM_PCT").Value, pct_okay)
                    .Range("SYSENG_TASK7_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SYSENG_TASK7_TOTAL_HOURS").Value, .Range("SYSENG_TASK7_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SYSENG_TASK7_REM_PCT").Value, pct_okay)
                    .Range("SYSENG_TASK8_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SYSENG_TASK8_TOTAL_HOURS").Value, .Range("SYSENG_TASK8_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SYSENG_TASK8_REM_PCT").Value, pct_okay)
                    .Range("SYSENG_TASK9_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SYSENG_TASK9_TOTAL_HOURS").Value, .Range("SYSENG_TASK9_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SYSENG_TASK9_REM_PCT").Value, pct_okay)
                    .Range("SYSENG_TASK10_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SYSENG_TASK10_TOTAL_HOURS").Value, .Range("SYSENG_TASK10_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SYSENG_TASK10_REM_PCT").Value, pct_okay)
                    
                    'HMI
                    .Range("HMI_TASK1_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("HMI_TASK1_TOTAL_HOURS").Value, .Range("HMI_TASK1_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("HMI_TASK1_REM_PCT").Value, pct_okay)
                    .Range("HMI_TASK2_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("HMI_TASK2_TOTAL_HOURS").Value, .Range("HMI_TASK2_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("HMI_TASK2_REM_PCT").Value, pct_okay)
                    .Range("HMI_TASK3_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("HMI_TASK3_TOTAL_HOURS").Value, .Range("HMI_TASK3_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("HMI_TASK3_REM_PCT").Value, pct_okay)
                    .Range("HMI_TASK4_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("HMI_TASK4_TOTAL_HOURS").Value, .Range("HMI_TASK4_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("HMI_TASK4_REM_PCT").Value, pct_okay)
                    .Range("HMI_TASK5_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("HMI_TASK5_TOTAL_HOURS").Value, .Range("HMI_TASK5_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("HMI_TASK5_REM_PCT").Value, pct_okay)
                    .Range("HMI_TASK6_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("HMI_TASK6_TOTAL_HOURS").Value, .Range("HMI_TASK6_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("HMI_TASK6_REM_PCT").Value, pct_okay)
                    .Range("HMI_TASK7_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("HMI_TASK7_TOTAL_HOURS").Value, .Range("HMI_TASK7_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("HMI_TASK7_REM_PCT").Value, pct_okay)
                    .Range("HMI_TASK8_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("HMI_TASK8_TOTAL_HOURS").Value, .Range("HMI_TASK8_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("HMI_TASK8_REM_PCT").Value, pct_okay)
                    .Range("HMI_TASK9_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("HMI_TASK9_TOTAL_HOURS").Value, .Range("HMI_TASK9_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("HMI_TASK9_REM_PCT").Value, pct_okay)
                    .Range("HMI_TASK10_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("HMI_TASK10_TOTAL_HOURS").Value, .Range("HMI_TASK10_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("HMI_TASK10_REM_PCT").Value, pct_okay)
                    
                    'CP
                    .Range("CP_TASK1_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("CP_TASK1_TOTAL_HOURS").Value, .Range("CP_TASK1_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("CP_TASK1_REM_PCT").Value, pct_okay)
                    .Range("CP_TASK2_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("CP_TASK2_TOTAL_HOURS").Value, .Range("CP_TASK2_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("CP_TASK2_REM_PCT").Value, pct_okay)
                    .Range("CP_TASK3_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("CP_TASK3_TOTAL_HOURS").Value, .Range("CP_TASK3_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("CP_TASK3_REM_PCT").Value, pct_okay)
                    .Range("CP_TASK4_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("CP_TASK4_TOTAL_HOURS").Value, .Range("CP_TASK4_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("CP_TASK4_REM_PCT").Value, pct_okay)
                    .Range("CP_TASK5_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("CP_TASK5_TOTAL_HOURS").Value, .Range("CP_TASK5_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("CP_TASK5_REM_PCT").Value, pct_okay)
                    .Range("CP_TASK6_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("CP_TASK6_TOTAL_HOURS").Value, .Range("CP_TASK6_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("CP_TASK6_REM_PCT").Value, pct_okay)
                    .Range("CP_TASK7_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("CP_TASK7_TOTAL_HOURS").Value, .Range("CP_TASK7_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("CP_TASK7_REM_PCT").Value, pct_okay)
                    .Range("CP_TASK8_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("CP_TASK8_TOTAL_HOURS").Value, .Range("CP_TASK8_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("CP_TASK8_REM_PCT").Value, pct_okay)
                    .Range("CP_TASK9_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("CP_TASK9_TOTAL_HOURS").Value, .Range("CP_TASK9_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("CP_TASK9_REM_PCT").Value, pct_okay)
                    .Range("CP_TASK10_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("CP_TASK10_TOTAL_HOURS").Value, .Range("CP_TASK10_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("CP_TASK10_REM_PCT").Value, pct_okay)
                    
                    'DI
                    .Range("DI_TASK1_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("DI_TASK1_TOTAL_HOURS").Value, .Range("DI_TASK1_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("DI_TASK1_REM_PCT").Value, pct_okay)
                    .Range("DI_TASK2_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("DI_TASK2_TOTAL_HOURS").Value, .Range("DI_TASK2_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("DI_TASK2_REM_PCT").Value, pct_okay)
                    .Range("DI_TASK3_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("DI_TASK3_TOTAL_HOURS").Value, .Range("DI_TASK3_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("DI_TASK3_REM_PCT").Value, pct_okay)
                    .Range("DI_TASK4_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("DI_TASK4_TOTAL_HOURS").Value, .Range("DI_TASK4_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("DI_TASK4_REM_PCT").Value, pct_okay)
                    .Range("DI_TASK5_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("DI_TASK5_TOTAL_HOURS").Value, .Range("DI_TASK5_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("DI_TASK5_REM_PCT").Value, pct_okay)
                    .Range("DI_TASK6_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("DI_TASK6_TOTAL_HOURS").Value, .Range("DI_TASK6_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("DI_TASK6_REM_PCT").Value, pct_okay)
                    .Range("DI_TASK7_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("DI_TASK7_TOTAL_HOURS").Value, .Range("DI_TASK7_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("DI_TASK7_REM_PCT").Value, pct_okay)
                    .Range("DI_TASK8_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("DI_TASK8_TOTAL_HOURS").Value, .Range("DI_TASK8_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("DI_TASK8_REM_PCT").Value, pct_okay)
                    .Range("DI_TASK9_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("DI_TASK9_TOTAL_HOURS").Value, .Range("DI_TASK9_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("DI_TASK9_REM_PCT").Value, pct_okay)
                    .Range("DI_TASK10_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("DI_TASK10_TOTAL_HOURS").Value, .Range("DI_TASK10_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("DI_TASK10_REM_PCT").Value, pct_okay)
                    
                    'ESD
                    .Range("ESD_TASK1_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("ESD_TASK1_TOTAL_HOURS").Value, .Range("ESD_TASK1_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("ESD_TASK1_REM_PCT").Value, pct_okay)
                    .Range("ESD_TASK2_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("ESD_TASK2_TOTAL_HOURS").Value, .Range("ESD_TASK2_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("ESD_TASK2_REM_PCT").Value, pct_okay)
                    .Range("ESD_TASK3_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("ESD_TASK3_TOTAL_HOURS").Value, .Range("ESD_TASK3_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("ESD_TASK3_REM_PCT").Value, pct_okay)
                    .Range("ESD_TASK4_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("ESD_TASK4_TOTAL_HOURS").Value, .Range("ESD_TASK4_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("ESD_TASK4_REM_PCT").Value, pct_okay)
                    .Range("ESD_TASK5_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("ESD_TASK5_TOTAL_HOURS").Value, .Range("ESD_TASK5_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("ESD_TASK5_REM_PCT").Value, pct_okay)
                    .Range("ESD_TASK6_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("ESD_TASK6_TOTAL_HOURS").Value, .Range("ESD_TASK6_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("ESD_TASK6_REM_PCT").Value, pct_okay)
                    .Range("ESD_TASK7_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("ESD_TASK7_TOTAL_HOURS").Value, .Range("ESD_TASK7_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("ESD_TASK7_REM_PCT").Value, pct_okay)
                    .Range("ESD_TASK8_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("ESD_TASK8_TOTAL_HOURS").Value, .Range("ESD_TASK8_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("ESD_TASK8_REM_PCT").Value, pct_okay)
                    .Range("ESD_TASK9_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("ESD_TASK9_TOTAL_HOURS").Value, .Range("ESD_TASK9_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("ESD_TASK9_REM_PCT").Value, pct_okay)
                    .Range("ESD_TASK10_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("ESD_TASK10_TOTAL_HOURS").Value, .Range("ESD_TASK10_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("ESD_TASK10_REM_PCT").Value, pct_okay)
                    
                    'REP
                    .Range("REP_TASK1_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("REP_TASK1_TOTAL_HOURS").Value, .Range("REP_TASK1_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("REP_TASK1_REM_PCT").Value, pct_okay)
                    .Range("REP_TASK2_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("REP_TASK2_TOTAL_HOURS").Value, .Range("REP_TASK2_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("REP_TASK2_REM_PCT").Value, pct_okay)
                    .Range("REP_TASK3_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("REP_TASK3_TOTAL_HOURS").Value, .Range("REP_TASK3_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("REP_TASK3_REM_PCT").Value, pct_okay)
                    .Range("REP_TASK4_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("REP_TASK4_TOTAL_HOURS").Value, .Range("REP_TASK4_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("REP_TASK4_REM_PCT").Value, pct_okay)
                    .Range("REP_TASK5_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("REP_TASK5_TOTAL_HOURS").Value, .Range("REP_TASK5_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("REP_TASK5_REM_PCT").Value, pct_okay)
                    .Range("REP_TASK6_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("REP_TASK6_TOTAL_HOURS").Value, .Range("REP_TASK6_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("REP_TASK6_REM_PCT").Value, pct_okay)
                    .Range("REP_TASK7_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("REP_TASK7_TOTAL_HOURS").Value, .Range("REP_TASK7_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("REP_TASK7_REM_PCT").Value, pct_okay)
                    .Range("REP_TASK8_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("REP_TASK8_TOTAL_HOURS").Value, .Range("REP_TASK8_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("REP_TASK8_REM_PCT").Value, pct_okay)
                    .Range("REP_TASK9_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("REP_TASK9_TOTAL_HOURS").Value, .Range("REP_TASK9_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("REP_TASK9_REM_PCT").Value, pct_okay)
                    .Range("REP_TASK10_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("REP_TASK10_TOTAL_HOURS").Value, .Range("REP_TASK10_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("REP_TASK10_REM_PCT").Value, pct_okay)
                    
                    'TEST
                    .Range("TEST_TASK1_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("TEST_TASK1_TOTAL_HOURS").Value, .Range("TEST_TASK1_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("TEST_TASK1_REM_PCT").Value, pct_okay)
                    .Range("TEST_TASK2_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("TEST_TASK2_TOTAL_HOURS").Value, .Range("TEST_TASK2_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("TEST_TASK2_REM_PCT").Value, pct_okay)
                    .Range("TEST_TASK3_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("TEST_TASK3_TOTAL_HOURS").Value, .Range("TEST_TASK3_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("TEST_TASK3_REM_PCT").Value, pct_okay)
                    .Range("TEST_TASK4_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("TEST_TASK4_TOTAL_HOURS").Value, .Range("TEST_TASK4_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("TEST_TASK4_REM_PCT").Value, pct_okay)
                    .Range("TEST_TASK5_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("TEST_TASK5_TOTAL_HOURS").Value, .Range("TEST_TASK5_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("TEST_TASK5_REM_PCT").Value, pct_okay)
                    .Range("TEST_TASK6_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("TEST_TASK6_TOTAL_HOURS").Value, .Range("TEST_TASK6_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("TEST_TASK6_REM_PCT").Value, pct_okay)
                    .Range("TEST_TASK7_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("TEST_TASK7_TOTAL_HOURS").Value, .Range("TEST_TASK7_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("TEST_TASK7_REM_PCT").Value, pct_okay)
                    .Range("TEST_TASK8_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("TEST_TASK8_TOTAL_HOURS").Value, .Range("TEST_TASK8_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("TEST_TASK8_REM_PCT").Value, pct_okay)
                    .Range("TEST_TASK9_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("TEST_TASK9_TOTAL_HOURS").Value, .Range("TEST_TASK9_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("TEST_TASK9_REM_PCT").Value, pct_okay)
                    .Range("TEST_TASK10_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("TEST_TASK10_TOTAL_HOURS").Value, .Range("TEST_TASK10_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("TEST_TASK10_REM_PCT").Value, pct_okay)
                    
                    'DOC
                    .Range("DOC_TASK1_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("DOC_TASK1_TOTAL_HOURS").Value, .Range("DOC_TASK1_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("DOC_TASK1_REM_PCT").Value, pct_okay)
                    .Range("DOC_TASK2_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("DOC_TASK2_TOTAL_HOURS").Value, .Range("DOC_TASK2_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("DOC_TASK2_REM_PCT").Value, pct_okay)
                    .Range("DOC_TASK3_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("DOC_TASK3_TOTAL_HOURS").Value, .Range("DOC_TASK3_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("DOC_TASK3_REM_PCT").Value, pct_okay)
                    .Range("DOC_TASK4_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("DOC_TASK4_TOTAL_HOURS").Value, .Range("DOC_TASK4_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("DOC_TASK4_REM_PCT").Value, pct_okay)
                    .Range("DOC_TASK5_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("DOC_TASK5_TOTAL_HOURS").Value, .Range("DOC_TASK5_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("DOC_TASK5_REM_PCT").Value, pct_okay)
                    .Range("DOC_TASK6_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("DOC_TASK6_TOTAL_HOURS").Value, .Range("DOC_TASK6_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("DOC_TASK6_REM_PCT").Value, pct_okay)
                    .Range("DOC_TASK7_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("DOC_TASK7_TOTAL_HOURS").Value, .Range("DOC_TASK7_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("DOC_TASK7_REM_PCT").Value, pct_okay)
                    .Range("DOC_TASK8_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("DOC_TASK8_TOTAL_HOURS").Value, .Range("DOC_TASK8_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("DOC_TASK8_REM_PCT").Value, pct_okay)
                    .Range("DOC_TASK9_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("DOC_TASK9_TOTAL_HOURS").Value, .Range("DOC_TASK9_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("DOC_TASK9_REM_PCT").Value, pct_okay)
                    .Range("DOC_TASK10_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("DOC_TASK10_TOTAL_HOURS").Value, .Range("DOC_TASK10_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("DOC_TASK10_REM_PCT").Value, pct_okay)
                    
                    'COURSE
                    .Range("COURSE_TASK1_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("COURSE_TASK1_TOTAL_HOURS").Value, .Range("COURSE_TASK1_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("COURSE_TASK1_REM_PCT").Value, pct_okay)
                    .Range("COURSE_TASK2_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("COURSE_TASK2_TOTAL_HOURS").Value, .Range("COURSE_TASK2_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("COURSE_TASK2_REM_PCT").Value, pct_okay)
                    .Range("COURSE_TASK3_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("COURSE_TASK3_TOTAL_HOURS").Value, .Range("COURSE_TASK3_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("COURSE_TASK3_REM_PCT").Value, pct_okay)
                    .Range("COURSE_TASK4_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("COURSE_TASK4_TOTAL_HOURS").Value, .Range("COURSE_TASK4_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("COURSE_TASK4_REM_PCT").Value, pct_okay)
                    .Range("COURSE_TASK5_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("COURSE_TASK5_TOTAL_HOURS").Value, .Range("COURSE_TASK5_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("COURSE_TASK5_REM_PCT").Value, pct_okay)
                    .Range("COURSE_TASK6_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("COURSE_TASK6_TOTAL_HOURS").Value, .Range("COURSE_TASK6_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("COURSE_TASK6_REM_PCT").Value, pct_okay)
                    .Range("COURSE_TASK7_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("COURSE_TASK7_TOTAL_HOURS").Value, .Range("COURSE_TASK7_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("COURSE_TASK7_REM_PCT").Value, pct_okay)
                    .Range("COURSE_TASK8_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("COURSE_TASK8_TOTAL_HOURS").Value, .Range("COURSE_TASK8_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("COURSE_TASK8_REM_PCT").Value, pct_okay)
                    .Range("COURSE_TASK9_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("COURSE_TASK9_TOTAL_HOURS").Value, .Range("COURSE_TASK9_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("COURSE_TASK9_REM_PCT").Value, pct_okay)
                    .Range("COURSE_TASK10_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("COURSE_TASK10_TOTAL_HOURS").Value, .Range("COURSE_TASK10_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("COURSE_TASK10_REM_PCT").Value, pct_okay)
                    
                    'PM
                    .Range("PM_TASK1_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("PM_TASK1_TOTAL_HOURS").Value, .Range("PM_TASK1_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("PM_TASK1_REM_PCT").Value, pct_okay)
                    .Range("PM_TASK2_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("PM_TASK2_TOTAL_HOURS").Value, .Range("PM_TASK2_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("PM_TASK2_REM_PCT").Value, pct_okay)
                    .Range("PM_TASK3_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("PM_TASK3_TOTAL_HOURS").Value, .Range("PM_TASK3_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("PM_TASK3_REM_PCT").Value, pct_okay)
                    .Range("PM_TASK4_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("PM_TASK4_TOTAL_HOURS").Value, .Range("PM_TASK4_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("PM_TASK4_REM_PCT").Value, pct_okay)
                    .Range("PM_TASK5_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("PM_TASK5_TOTAL_HOURS").Value, .Range("PM_TASK5_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("PM_TASK5_REM_PCT").Value, pct_okay)
                    .Range("PM_TASK6_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("PM_TASK6_TOTAL_HOURS").Value, .Range("PM_TASK6_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("PM_TASK6_REM_PCT").Value, pct_okay)
                    .Range("PM_TASK7_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("PM_TASK7_TOTAL_HOURS").Value, .Range("PM_TASK7_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("PM_TASK7_REM_PCT").Value, pct_okay)
                    .Range("PM_TASK8_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("PM_TASK8_TOTAL_HOURS").Value, .Range("PM_TASK8_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("PM_TASK8_REM_PCT").Value, pct_okay)
                    .Range("PM_TASK9_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("PM_TASK9_TOTAL_HOURS").Value, .Range("PM_TASK9_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("PM_TASK9_REM_PCT").Value, pct_okay)
                    .Range("PM_TASK10_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("PM_TASK10_TOTAL_HOURS").Value, .Range("PM_TASK10_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("PM_TASK10_REM_PCT").Value, pct_okay)
                    
                    'MEETING
                    .Range("MEETING_TASK1_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("MEETING_TASK1_TOTAL_HOURS").Value, .Range("MEETING_TASK1_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("MEETING_TASK1_REM_PCT").Value, pct_okay)
                    .Range("MEETING_TASK2_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("MEETING_TASK2_TOTAL_HOURS").Value, .Range("MEETING_TASK2_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("MEETING_TASK2_REM_PCT").Value, pct_okay)
                    .Range("MEETING_TASK3_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("MEETING_TASK3_TOTAL_HOURS").Value, .Range("MEETING_TASK3_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("MEETING_TASK3_REM_PCT").Value, pct_okay)
                    .Range("MEETING_TASK4_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("MEETING_TASK4_TOTAL_HOURS").Value, .Range("MEETING_TASK4_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("MEETING_TASK4_REM_PCT").Value, pct_okay)
                    .Range("MEETING_TASK5_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("MEETING_TASK5_TOTAL_HOURS").Value, .Range("MEETING_TASK5_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("MEETING_TASK5_REM_PCT").Value, pct_okay)
                    .Range("MEETING_TASK6_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("MEETING_TASK6_TOTAL_HOURS").Value, .Range("MEETING_TASK6_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("MEETING_TASK6_REM_PCT").Value, pct_okay)
                    .Range("MEETING_TASK7_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("MEETING_TASK7_TOTAL_HOURS").Value, .Range("MEETING_TASK7_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("MEETING_TASK7_REM_PCT").Value, pct_okay)
                    .Range("MEETING_TASK8_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("MEETING_TASK8_TOTAL_HOURS").Value, .Range("MEETING_TASK8_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("MEETING_TASK8_REM_PCT").Value, pct_okay)
                    .Range("MEETING_TASK9_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("MEETING_TASK9_TOTAL_HOURS").Value, .Range("MEETING_TASK9_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("MEETING_TASK9_REM_PCT").Value, pct_okay)
                    .Range("MEETING_TASK10_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("MEETING_TASK10_TOTAL_HOURS").Value, .Range("MEETING_TASK10_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("MEETING_TASK10_REM_PCT").Value, pct_okay)
                    
                    'SITE
                    .Range("SITE_TASK1_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SITE_TASK1_TOTAL_HOURS").Value, .Range("SITE_TASK1_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SITE_TASK1_REM_PCT").Value, pct_okay)
                    .Range("SITE_TASK2_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SITE_TASK2_TOTAL_HOURS").Value, .Range("SITE_TASK2_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SITE_TASK2_REM_PCT").Value, pct_okay)
                    .Range("SITE_TASK3_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SITE_TASK3_TOTAL_HOURS").Value, .Range("SITE_TASK3_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SITE_TASK3_REM_PCT").Value, pct_okay)
                    .Range("SITE_TASK4_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SITE_TASK4_TOTAL_HOURS").Value, .Range("SITE_TASK4_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SITE_TASK4_REM_PCT").Value, pct_okay)
                    .Range("SITE_TASK5_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SITE_TASK5_TOTAL_HOURS").Value, .Range("SITE_TASK5_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SITE_TASK5_REM_PCT").Value, pct_okay)
                    .Range("SITE_TASK6_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SITE_TASK6_TOTAL_HOURS").Value, .Range("SITE_TASK6_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SITE_TASK6_REM_PCT").Value, pct_okay)
                    .Range("SITE_TASK7_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SITE_TASK7_TOTAL_HOURS").Value, .Range("SITE_TASK7_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SITE_TASK7_REM_PCT").Value, pct_okay)
                    .Range("SITE_TASK8_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SITE_TASK8_TOTAL_HOURS").Value, .Range("SITE_TASK8_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SITE_TASK8_REM_PCT").Value, pct_okay)
                    .Range("SITE_TASK9_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SITE_TASK9_TOTAL_HOURS").Value, .Range("SITE_TASK9_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SITE_TASK9_REM_PCT").Value, pct_okay)
                    .Range("SITE_TASK10_REM_PCT").Value = NewRemotePercentage(Workbooks(gstrGECEWorkBook).Worksheets(gstrGECEPriceMakeUpSheet).Range("SITE_TASK10_TOTAL_HOURS").Value, .Range("SITE_TASK10_REM_PCT").Value, pct_change)
                    pct_okay = ModPctPossible(.Range("SITE_TASK10_REM_PCT").Value, pct_okay)
                    
                'End If
                
                If (Abs(.Range("TOTAL_APP_BASED_REM_PCT").Value - .Range("REM_PCT_REC").Value) < 0.01) Or (pct_okay = True) Then
                    Exit For
                End If
            Next ii
            Application.Cursor = xlDefault
        End With
    End If
    
Exit Function
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
End Function


''

Function ModPctPossible(task_pct As Double, pct_okay As Boolean) As Boolean
On Error GoTo ErrHandler
    If task_pct <> 0 And task_pct <> 1 Then
        ModPctPossible = False
    Else
        ModPctPossible = pct_okay
    End If
Exit Function
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
End Function

Function NewRemotePercentage(task_hours As Double, task_pct As Double, pct_change As Double) As Double
On Error GoTo ErrHandler
    Dim pct_task_tmp As Double
    
    If task_hours <> 0 Then
        If pct_change = 1 Then
            pct_task_tmp = 1
        Else
            pct_task_tmp = pct_change * task_pct
        End If
    Else
        pct_task_tmp = 0
    End If
    
    If pct_task_tmp > 1 Then
        pct_task_tmp = 1
    ElseIf pct_task_tmp < 0 Then
        pct_task_tmp = 0
    End If
    
    NewRemotePercentage = pct_task_tmp
Exit Function
ErrHandler:
    MsgBox Err.Number & "; " & Err.Description & "; " & Err.Source
End Function


Private Sub unhideandunprotect()

'
'   'gets a list of all sheets and their visible property
'    For Each Sheet In Sheets
'        Debug.Print Sheet.Name & " '" & Sheet.Visible
'    Next
    
Dim strPwd As String
strPwd = InputBox("Enter Protection password")
If strPwd <> "" Then

ThisWorkbook.Worksheets("coverSheet").Select
ThisWorkbook.Worksheets("coverSheet").Activate

'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
'workbook
ActiveWorkbook.Protect Password:=strPwd, Structure:=False, Windows:=False

'Visible Sheets (set just in case they were marked hidden)
Sheets("CoverSheet").Visible = True
Sheets("Assumptions - Proposal infos").Visible = True
Sheets("Data Entry").Visible = True
Sheets("Application Based").Visible = True
Sheets("Task Based").Visible = True
Sheets("Duration Based").Visible = True
Sheets("GECE TO CQA Output").Visible = True
Sheets("Unit Price").Visible = True
Sheets("Free Form").Visible = True
Sheets("Global Factor").Visible = True
Sheets("Proposal Summary").Visible = True

'Hidden Sheets (unhide them)
Sheets("Estimator Revision").Visible = True
Sheets("Unit Pricing Calculations").Visible = True
Sheets("Price Make-up").Visible = True
Sheets("Industry").Visible = True
Sheets("ToolKit").Visible = True
Sheets("World Cost tables").Visible = True
Sheets("World Currency Table").Visible = True
Sheets("CurrencyUpdate").Visible = True
Sheets("Schedule").Visible = True
Sheets("ExportToERP").Visible = True
Sheets("WPA").Visible = True

'protect
'sheets that are protected
'unprotect them
Sheets("CoverSheet").Protect Password:=strPwd, DrawingObjects:=False, Contents:=False, Scenarios:=False
Sheets("Assumptions - Proposal infos").Protect Password:=strPwd, DrawingObjects:=False, Contents:=False, Scenarios:=False
Sheets("Data Entry").Protect Password:=strPwd, DrawingObjects:=False, Contents:=False, Scenarios:=False
Sheets("Application Based").Protect Password:=strPwd, DrawingObjects:=False, Contents:=False, Scenarios:=False
Sheets("Task Based").Protect Password:=strPwd, DrawingObjects:=False, Contents:=False, Scenarios:=False
Sheets("Duration Based").Protect Password:=strPwd, DrawingObjects:=False, Contents:=False, Scenarios:=False
Sheets("Global Factor").Protect Password:=strPwd, DrawingObjects:=False, Contents:=False, Scenarios:=False
Sheets("Estimator Revision").Protect Password:=strPwd, DrawingObjects:=False, Contents:=False, Scenarios:=False
Sheets("Unit Price").Protect Password:=strPwd, DrawingObjects:=False, Contents:=False, Scenarios:=False
Sheets("Unit Pricing Calculations").Protect Password:=strPwd, DrawingObjects:=False, Contents:=False, Scenarios:=False
Sheets("Proposal Summary").Protect Password:=strPwd, DrawingObjects:=False, Contents:=False, Scenarios:=False
Sheets("Price Make-up").Protect Password:=strPwd, DrawingObjects:=False, Contents:=False, Scenarios:=False
Sheets("Industry").Protect Password:=strPwd, DrawingObjects:=False, Contents:=False, Scenarios:=False
Sheets("ToolKit").Protect Password:=strPwd, DrawingObjects:=False, Contents:=False, Scenarios:=False
Sheets("World Cost tables").Protect Password:=strPwd, DrawingObjects:=False, Contents:=False, Scenarios:=False
Sheets("GECE TO CQA Output").Protect Password:=strPwd, DrawingObjects:=False, Contents:=False, Scenarios:=False
Sheets("WPA").Protect Password:=strPwd, DrawingObjects:=False, Contents:=False, Scenarios:=False

Sheets("Schedule").Protect Password:=strPwd, DrawingObjects:=False, Contents:=False, Scenarios:=False
Sheets("ExportToERP").Protect Password:=strPwd, DrawingObjects:=False, Contents:=False, Scenarios:=False



'sheets that are NOT protected (reset them)
Sheets("Free Form").Protect Password:=strPwd, DrawingObjects:=False, Contents:=False, Scenarios:=False
Sheets("CurrencyUpdate").Protect Password:=strPwd, DrawingObjects:=False, Contents:=False, Scenarios:=False
Sheets("World Currency Table").Protect Password:=strPwd, DrawingObjects:=False, Contents:=False, Scenarios:=False


Else
    MsgBox ("Password missing!")
End If

End Sub
Private Sub hideandprotect()
'dph 6.8
Dim strPwd As String
strPwd = InputBox("Enter Protection password")
If strPwd <> "" Then

ThisWorkbook.Worksheets("coverSheet").Select
ThisWorkbook.Worksheets("coverSheet").Activate

'Visible Sheets
Sheets("CoverSheet").Visible = True
Sheets("Assumptions - Proposal infos").Visible = True
Sheets("Data Entry").Visible = True
Sheets("Application Based").Visible = True
Sheets("Task Based").Visible = True
Sheets("Duration Based").Visible = True
Sheets("Free Form").Visible = True
Sheets("Global Factor").Visible = True
Sheets("GECE TO CQA Output").Visible = True
Sheets("Unit Price").Visible = True
Sheets("Proposal Summary").Visible = True

'Hidden Sheets
Sheets("Estimator Revision").Visible = False
Sheets("Unit Pricing Calculations").Visible = False
Sheets("Price Make-up").Visible = False
Sheets("Industry").Visible = False
Sheets("ToolKit").Visible = False
Sheets("World Cost tables").Visible = False
Sheets("World Currency Table").Visible = False
Sheets("CurrencyUpdate").Visible = False
Sheets("ExportToERP").Visible = False
Sheets("Schedule").Visible = False
Sheets("WPA").Visible = False




'protect
'sheets that are protected

Sheets("CoverSheet").Protect Password:=strPwd, DrawingObjects:=True, Contents:=True, Scenarios:=True
Sheets("Assumptions - Proposal infos").Protect Password:=strPwd, DrawingObjects:=True, Contents:=True, Scenarios:=True
Sheets("Data Entry").Protect Password:=strPwd, DrawingObjects:=True, Contents:=True, Scenarios:=True
Sheets("Application Based").Protect Password:=strPwd, DrawingObjects:=True, Contents:=True, Scenarios:=True
Sheets("Task Based").Protect Password:=strPwd, DrawingObjects:=True, Contents:=True, Scenarios:=True
Sheets("Duration Based").Protect Password:=strPwd, DrawingObjects:=True, Contents:=True, Scenarios:=True
Sheets("Global Factor").Protect Password:=strPwd, DrawingObjects:=True, Contents:=True, Scenarios:=True
Sheets("Estimator Revision").Protect Password:=strPwd, DrawingObjects:=True, Contents:=True, Scenarios:=True
Sheets("Unit Price").Protect Password:=strPwd, DrawingObjects:=True, Contents:=True, Scenarios:=True
Sheets("Unit Pricing Calculations").Protect Password:=strPwd, DrawingObjects:=True, Contents:=True, Scenarios:=True
Sheets("Proposal Summary").Protect Password:=strPwd, DrawingObjects:=True, Contents:=True, Scenarios:=True
Sheets("Price Make-up").Protect Password:=strPwd, DrawingObjects:=True, Contents:=True, Scenarios:=True
Sheets("Industry").Protect Password:=strPwd, DrawingObjects:=True, Contents:=True, Scenarios:=True
Sheets("ToolKit").Protect Password:=strPwd, DrawingObjects:=True, Contents:=True, Scenarios:=True
Sheets("World Cost tables").Protect Password:=strPwd, DrawingObjects:=True, Contents:=True, Scenarios:=True
Sheets("GECE TO CQA Output").Protect Password:=strPwd, DrawingObjects:=True, Contents:=True, Scenarios:=True
Sheets("Schedule").Protect Password:=strPwd, DrawingObjects:=True, Contents:=True, Scenarios:=True
Sheets("ExportToERP").Protect Password:=strPwd, DrawingObjects:=True, Contents:=True, Scenarios:=True
Sheets("WPA").Protect Password:=strPwd, DrawingObjects:=True, Contents:=True, Scenarios:=True

'sheets that are NOT protected
Sheets("Free Form").Protect Password:=strPwd, DrawingObjects:=False, Contents:=False, Scenarios:=False
Sheets("CurrencyUpdate").Protect Password:=strPwd, DrawingObjects:=False, Contents:=False, Scenarios:=False
Sheets("World Currency Table").Protect Password:=strPwd, DrawingObjects:=False, Contents:=False, Scenarios:=False


'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
'workbook
ActiveWorkbook.Protect Password:=strPwd, Structure:=True, Windows:=False

Else
    MsgBox ("Password missing!")
End If

End Sub

Private Sub UnlockProjectCLASS()

'
'   'gets a list of all sheets and their visible property
'    For Each Sheet In Sheets
'        Debug.Print Sheet.Name & " '" & Sheet.Visible
'    Next
    
Dim strPwd As String
strPwd = InputBox("Enter Protection password")
If strPwd <> "" Then

ThisWorkbook.Worksheets("coverSheet").Select
ThisWorkbook.Worksheets("coverSheet").Activate

'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
'workbook
ActiveWorkbook.Protect Password:=strPwd, Structure:=False, Windows:=False

'Visible Sheets (set just in case they were marked hidden)
Sheets("CoverSheet").Visible = True
Sheets("Assumptions - Proposal infos").Visible = True
Sheets("Data Entry").Visible = True
Sheets("Application Based").Visible = True
Sheets("Task Based").Visible = True
Sheets("Duration Based").Visible = True
Sheets("GECE TO CQA Output").Visible = True
Sheets("Unit Price").Visible = True
Sheets("Free Form").Visible = True
Sheets("Global Factor").Visible = True
Sheets("Proposal Summary").Visible = True
Sheets("Project CLASS").Visible = True

'protect
'sheets that are protected
'unprotect them
Sheets("Assumptions - Proposal infos").Protect Password:=strPwd, DrawingObjects:=False, Contents:=False, Scenarios:=False
Else
    MsgBox ("Password missing!")
End If

End Sub
Private Sub LockProjectCLASS()
    Call hideandprotect
End Sub

Private Sub ReportProtectedSheets()
Dim ws As Worksheet
Dim wbs As Sheets
Set wbs = ActiveWorkbook.Sheets

   'gets a list of all sheets and their protected property
    For Each ws In wbs
        Debug.Print " '" & ws.Name & vbTab & vbTab & " DrawingObjects:=" & ws.ProtectDrawingObjects & " Contents:=" & ws.ProtectContents & " Scenarios:=" & ws.ProtectScenarios
    Next
End Sub

Private Sub ShowGECETool()
    ReplaceCellName.Show
End Sub

Public Function AddSlash(p As String) As String
    If Len(p) > 0 And Right$(p, 1) <> "\" Then
        p = p & "\"
    End If
    AddSlash = p
End Function
Public Function GetGECEPath_Universal() As String
    GetGECEPath_Universal = GetGECEPath()
End Function
Public Function PercentAsDouble(v As Variant) As Double
    On Error GoTo Z
    If IsError(v) Or IsEmpty(v) Then GoTo Z
    Dim s$: s$ = Trim$(CStr(v))
    If s$ = "" Then GoTo Z
    If Right$(s$, 1) = "%" Then
        PercentAsDouble = CDbl(Left$(s$, Len(s$) - 1)) / 100#
    ElseIf CDbl(s$) > 1# Then
        PercentAsDouble = CDbl(s$) / 100#
    Else
        PercentAsDouble = CDbl(s$)
    End If
    If PercentAsDouble < 0# Then PercentAsDouble = 0#
    If PercentAsDouble > 1# Then PercentAsDouble = 1#
    Exit Function
Z:  PercentAsDouble = 0#
End Function

Public Function PercentText(v As Variant) As String
    PercentText = Format(PercentAsDouble(v), "0%")
End Function
