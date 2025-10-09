Attribute VB_Name = "xmlMod"
'Option Explicit
Public mstrDecimalSeparator As String
Public mstrCPDecimalSeparator As String
Public Sub ExportPriceMarkupSheet(wsTarget As Worksheet, ByRef rinvGECE As MSXML2.DOMDocument30, ByRef robjRoot As Object)
On Error GoTo ErrorHandler
Dim objDataEntry                 As Object
With wsTarget
    Set objDataEntry = xmlCreateSubRoot(rinvGECE, robjRoot, "PriceMarkupSheet")
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK1_OVD_STS", .Range("SPEC_TASK1_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK2_OVD_STS", .Range("SPEC_TASK2_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK3_OVD_STS", .Range("SPEC_TASK3_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK4_OVD_STS", .Range("SPEC_TASK4_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK5_OVD_STS", .Range("SPEC_TASK5_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK6_OVD_STS", .Range("SPEC_TASK6_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK7_OVD_STS", .Range("SPEC_TASK7_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK8_OVD_STS", .Range("SPEC_TASK8_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK9_OVD_STS", .Range("SPEC_TASK9_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK10_OVD_STS", .Range("SPEC_TASK10_OVD_STS").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK1_OVD_STS", .Range("SYSENG_TASK1_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK2_OVD_STS", .Range("SYSENG_TASK2_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK3_OVD_STS", .Range("SYSENG_TASK3_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK4_OVD_STS", .Range("SYSENG_TASK4_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK5_OVD_STS", .Range("SYSENG_TASK5_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK6_OVD_STS", .Range("SYSENG_TASK6_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK7_OVD_STS", .Range("SYSENG_TASK7_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK8_OVD_STS", .Range("SYSENG_TASK8_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK9_OVD_STS", .Range("SYSENG_TASK9_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK10_OVD_STS", .Range("SYSENG_TASK10_OVD_STS").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK1_OVD_STS", .Range("HMI_TASK1_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK2_OVD_STS", .Range("HMI_TASK2_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK3_OVD_STS", .Range("HMI_TASK3_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK4_OVD_STS", .Range("HMI_TASK4_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK5_OVD_STS", .Range("HMI_TASK5_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK6_OVD_STS", .Range("HMI_TASK6_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK7_OVD_STS", .Range("HMI_TASK7_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK8_OVD_STS", .Range("HMI_TASK8_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK9_OVD_STS", .Range("HMI_TASK9_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK10_OVD_STS", .Range("HMI_TASK10_OVD_STS").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK1_OVD_STS", .Range("CP_TASK1_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK2_OVD_STS", .Range("CP_TASK2_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK3_OVD_STS", .Range("CP_TASK3_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK4_OVD_STS", .Range("CP_TASK4_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK5_OVD_STS", .Range("CP_TASK5_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK6_OVD_STS", .Range("CP_TASK6_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK7_OVD_STS", .Range("CP_TASK7_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK8_OVD_STS", .Range("CP_TASK8_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK9_OVD_STS", .Range("CP_TASK9_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK10_OVD_STS", .Range("CP_TASK10_OVD_STS").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK1_OVD_STS", .Range("DI_TASK1_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK2_OVD_STS", .Range("DI_TASK2_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK3_OVD_STS", .Range("DI_TASK3_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK4_OVD_STS", .Range("DI_TASK4_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK5_OVD_STS", .Range("DI_TASK5_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK6_OVD_STS", .Range("DI_TASK6_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK7_OVD_STS", .Range("DI_TASK7_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK8_OVD_STS", .Range("DI_TASK8_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK9_OVD_STS", .Range("DI_TASK9_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK10_OVD_STS", .Range("DI_TASK10_OVD_STS").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK1_OVD_STS", .Range("ESD_TASK1_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK2_OVD_STS", .Range("ESD_TASK2_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK3_OVD_STS", .Range("ESD_TASK3_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK4_OVD_STS", .Range("ESD_TASK4_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK5_OVD_STS", .Range("ESD_TASK5_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK6_OVD_STS", .Range("ESD_TASK6_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK7_OVD_STS", .Range("ESD_TASK7_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK8_OVD_STS", .Range("ESD_TASK8_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK9_OVD_STS", .Range("ESD_TASK9_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK10_OVD_STS", .Range("ESD_TASK10_OVD_STS").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK1_OVD_STS", .Range("REP_TASK1_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK2_OVD_STS", .Range("REP_TASK2_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK3_OVD_STS", .Range("REP_TASK3_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK4_OVD_STS", .Range("REP_TASK4_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK5_OVD_STS", .Range("REP_TASK5_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK6_OVD_STS", .Range("REP_TASK6_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK7_OVD_STS", .Range("REP_TASK7_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK8_OVD_STS", .Range("REP_TASK8_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK9_OVD_STS", .Range("REP_TASK9_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK10_OVD_STS", .Range("REP_TASK10_OVD_STS").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK1_OVD_STS", .Range("APP_TASK1_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK2_OVD_STS", .Range("APP_TASK2_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK3_OVD_STS", .Range("APP_TASK3_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK4_OVD_STS", .Range("APP_TASK4_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK5_OVD_STS", .Range("APP_TASK5_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK6_OVD_STS", .Range("APP_TASK6_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK7_OVD_STS", .Range("APP_TASK7_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK8_OVD_STS", .Range("APP_TASK8_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK9_OVD_STS", .Range("APP_TASK9_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK10_OVD_STS", .Range("APP_TASK10_OVD_STS").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK1_OVD_STS", .Range("TEST_TASK1_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK2_OVD_STS", .Range("TEST_TASK2_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK3_OVD_STS", .Range("TEST_TASK3_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK4_OVD_STS", .Range("TEST_TASK4_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK5_OVD_STS", .Range("TEST_TASK5_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK6_OVD_STS", .Range("TEST_TASK6_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK7_OVD_STS", .Range("TEST_TASK7_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK8_OVD_STS", .Range("TEST_TASK8_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK9_OVD_STS", .Range("TEST_TASK9_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK10_OVD_STS", .Range("TEST_TASK10_OVD_STS").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK1_OVD_STS", .Range("DOC_TASK1_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK2_OVD_STS", .Range("DOC_TASK2_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK3_OVD_STS", .Range("DOC_TASK3_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK4_OVD_STS", .Range("DOC_TASK4_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK5_OVD_STS", .Range("DOC_TASK5_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK6_OVD_STS", .Range("DOC_TASK6_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK7_OVD_STS", .Range("DOC_TASK7_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK8_OVD_STS", .Range("DOC_TASK8_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK9_OVD_STS", .Range("DOC_TASK9_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK10_OVD_STS", .Range("DOC_TASK10_OVD_STS").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK1_OVD_STS", .Range("COURSE_TASK1_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK2_OVD_STS", .Range("COURSE_TASK2_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK3_OVD_STS", .Range("COURSE_TASK3_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK4_OVD_STS", .Range("COURSE_TASK4_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK5_OVD_STS", .Range("COURSE_TASK5_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK6_OVD_STS", .Range("COURSE_TASK6_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK7_OVD_STS", .Range("COURSE_TASK7_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK8_OVD_STS", .Range("COURSE_TASK8_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK9_OVD_STS", .Range("COURSE_TASK9_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK10_OVD_STS", .Range("COURSE_TASK10_OVD_STS").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK1_OVD_STS", .Range("PM_TASK1_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK2_OVD_STS", .Range("PM_TASK2_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK3_OVD_STS", .Range("PM_TASK3_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK4_OVD_STS", .Range("PM_TASK4_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK5_OVD_STS", .Range("PM_TASK5_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK6_OVD_STS", .Range("PM_TASK6_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK7_OVD_STS", .Range("PM_TASK7_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK8_OVD_STS", .Range("PM_TASK8_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK9_OVD_STS", .Range("PM_TASK9_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK10_OVD_STS", .Range("PM_TASK10_OVD_STS").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK1_OVD_STS", .Range("MEETING_TASK1_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK2_OVD_STS", .Range("MEETING_TASK2_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK3_OVD_STS", .Range("MEETING_TASK3_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK4_OVD_STS", .Range("MEETING_TASK4_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK5_OVD_STS", .Range("MEETING_TASK5_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK6_OVD_STS", .Range("MEETING_TASK6_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK7_OVD_STS", .Range("MEETING_TASK7_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK8_OVD_STS", .Range("MEETING_TASK8_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK9_OVD_STS", .Range("MEETING_TASK9_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK10_OVD_STS", .Range("MEETING_TASK10_OVD_STS").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK1_OVD_STS", .Range("SITE_TASK1_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK2_OVD_STS", .Range("SITE_TASK2_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK3_OVD_STS", .Range("SITE_TASK3_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK4_OVD_STS", .Range("SITE_TASK4_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK5_OVD_STS", .Range("SITE_TASK5_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK6_OVD_STS", .Range("SITE_TASK6_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK7_OVD_STS", .Range("SITE_TASK7_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK8_OVD_STS", .Range("SITE_TASK8_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK9_OVD_STS", .Range("SITE_TASK9_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK10_OVD_STS", .Range("SITE_TASK10_OVD_STS").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK1_OVD_STS", .Range("TL_TASK1_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK2_OVD_STS", .Range("TL_TASK2_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK3_OVD_STS", .Range("TL_TASK3_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK4_OVD_STS", .Range("TL_TASK4_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK5_OVD_STS", .Range("TL_TASK5_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK6_OVD_STS", .Range("TL_TASK6_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK7_OVD_STS", .Range("TL_TASK7_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK8_OVD_STS", .Range("TL_TASK8_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK9_OVD_STS", .Range("TL_TASK9_OVD_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK10_OVD_STS", .Range("TL_TASK10_OVD_STS").Value)
    

    
    
    
End With
CleanUp:
    On Error Resume Next
    Set objDataEntry = Nothing


Exit Sub
ErrorHandler:
    MsgBox "error in ExportPriceMarkupSheet " & Err.Description, vbCritical
    Resume CleanUp
End Sub

Public Sub ExportDataEntrySheet(wsTarget As Worksheet, ByRef rinvGECE As MSXML2.DOMDocument30, ByRef robjRoot As Object)
On Error GoTo ErrorHandler
Dim objDataEntry                 As Object


With wsTarget
    Set objDataEntry = xmlCreateSubRoot(rinvGECE, robjRoot, "DataEntry")
    Call xmlCreateElement(rinvGECE, objDataEntry, "HIGH_RISK_SITE", .Range("HIGH_RISK_SITE").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TRIPS_REQ_FAT", .Range("TL_TRIPS_REQ_FAT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_PERS_QTY_FAT", .Range("TL_PERS_QTY_FAT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_DAYS_QTY_FAT", .Range("TL_DAYS_QTY_FAT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_AIRFARE_FAT", .Range("TL_AIRFARE_FAT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_DAILY_ALLOW_SITE", .Range("TL_DAILY_ALLOW_SITE").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TRIPS_REQ_SITE", .Range("TL_TRIPS_REQ_SITE").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_PERS_QTY_SITE", .Range("TL_PERS_QTY_SITE").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_DAYS_QTY_SITE", .Range("TL_DAYS_QTY_SITE").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_AIRFARE_SITE", .Range("TL_AIRFARE_SITE").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_1_HOURS", .Range("APP_1_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_2_HOURS", .Range("APP_2_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_3_HOURS", .Range("APP_3_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_4_HOURS", .Range("APP_4_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_5_HOURS", .Range("APP_5_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_6_HOURS", .Range("APP_6_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_7_HOURS", .Range("APP_7_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_8_HOURS", .Range("APP_8_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_BUS_STS", .Range("APP_BUS_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_BUS_STS", .Range("APP_BUS_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CAB_CONSOLES_QTY", .Range("CAB_CONSOLES_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CAB_IO_QTY", .Range("CAB_IO_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CAB_MARSH_QTY", .Range("CAB_MARSH_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CAB_PROC_QTY", .Range("CAB_PROC_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COST_BUYOUT", .Range("COST_BUYOUT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COST_IA", .Range("COST_IA").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_1_HOURS", .Range("COURSE_1_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_2_HOURS", .Range("COURSE_2_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_3_HOURS", .Range("COURSE_3_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_4_HOURS", .Range("COURSE_4_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_5_HOURS", .Range("COURSE_5_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE1_HOURS", .Range("COURSE1_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE1_NAME", .Range("COURSE1_NAME").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE2_HOURS", .Range("COURSE2_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE2_NAME", .Range("COURSE2_NAME").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE3_HOURS", .Range("COURSE3_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE3_NAME", .Range("COURSE3_NAME").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE4_HOURS", .Range("COURSE4_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE4_NAME", .Range("COURSE4_NAME").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE5_HOURS", .Range("COURSE5_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE5_NAME", .Range("COURSE5_NAME").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_AI", .Range("CP_AI").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_ANA_COMPLEX_QTY", .Range("CP_ANA_COMPLEX_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_AO", .Range("CP_AO").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_DI", .Range("CP_DI").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_DIGITAL_COMPLEX_QTY", .Range("CP_DIGITAL_COMPLEX_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_DIGITAL_CTRL_DI", .Range("CP_DIGITAL_CTRL_DI").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_DIGITAL_CTRL_DO", .Range("CP_DIGITAL_CTRL_DO").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_DO", .Range("CP_DO").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_FIELDBUS_IO_QTY", .Range("CP_FIELDBUS_IO_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_FIELDBUS_IO_RATIO", .Range("CP_FIELDBUS_IO_RATIO").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_GRP_START_COMPLEX_QTY", .Range("CP_GRP_START_COMPLEX_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_GRP_START_LOOP_QTY", .Range("CP_GRP_START_LOOP_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_SEQ_COMPLEX_QTY", .Range("CP_SEQ_COMPLEX_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_SEQ_LOOP_QTY", .Range("CP_SEQ_LOOP_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CUSTOMER_SPEC_REQ", .Range("CUSTOMER_SPEC_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CUSTOMER_SPECS_EXIST", .Range("CUSTOMER_SPECS_EXIST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DATE_END", .Range("DATE_END").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DATE_START", .Range("DATE_START").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_AI", .Range("DI_AI").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_AO", .Range("DI_AO").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_COMPLEX_QTY", .Range("DI_COMPLEX_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_DEVICES", .Range("DI_DEVICES").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_DI", .Range("DI_DI").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_DIGITAL_CTRL_DI", .Range("DI_DIGITAL_CTRL_DI").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_DIGITAL_CTRL_DO", .Range("DI_DIGITAL_CTRL_DO").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_DO", .Range("DI_DO").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_GRP_START_COMPLEX_QTY", .Range("DI_GRP_START_COMPLEX_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_GRP_START_LOOP_QTY", .Range("DI_GRP_START_LOOP_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_INTERFACES", .Range("DI_INTERFACES").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_IOTYPE_STS", .Range("DI_IOTYPE_STS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_SEQ_COMPLEX_QTY", .Range("DI_SEQ_COMPLEX_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_SEQ_LOOP_QTY", .Range("DI_SEQ_LOOP_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_BOM_REQ", .Range("DOC_BOM_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_CAB_ELEC_REQ", .Range("DOC_CAB_ELEC_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_CAB_MECH_REQ", .Range("DOC_CAB_MECH_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_LOOP_REQ", .Range("DOC_LOOP_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_PWR_GND_REQ", .Range("DOC_PWR_GND_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_PWR_HEAT_REQ", .Range("DOC_PWR_HEAT_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_QA_REQ", .Range("DOC_QA_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_SYS_ARCH_REQ", .Range("DOC_SYS_ARCH_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_SYS_IND_REQ", .Range("DOC_SYS_IND_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TAGLIST_REQ", .Range("DOC_TAGLIST_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION", .Range("DURATION").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_AI", .Range("ESD_AI").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_AO", .Range("ESD_AO").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_CAB_IO_QTY", .Range("ESD_CAB_IO_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_CAB_MARSH_QTY", .Range("ESD_CAB_MARSH_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_CAB_PROC_QTY", .Range("ESD_CAB_PROC_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_CHASSIS_QTY", .Range("ESD_CHASSIS_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_COMM_QTY", .Range("ESD_COMM_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_COMPLEX_QTY", .Range("ESD_COMPLEX_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_DI", .Range("ESD_DI").Value)
    'Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_DI_DESIRED", .Range("ESD_DI_DESIRED").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_DO", .Range("ESD_DO").Value)
    'Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_DO_DESIRED", .Range("ESD_DO_DESIRED").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_GRP_START_COMPLEX_QTY", .Range("ESD_GRP_START_COMPLEX_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_GRP_START_LOOP_QTY", .Range("ESD_GRP_START_LOOP_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_HMI_REQ", .Range("ESD_HMI_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_IO_CARD_QTY", .Range("ESD_IO_CARD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_MARSH_CAB_REQ", .Range("ESD_MARSH_CAB_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_MISC_CAB_QTY", .Range("ESD_MISC_CAB_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_PROG_REQ", .Range("ESD_PROG_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "Aeroderivative_QTY", .Range("Aeroderivative_QTY").Value)
    'Call xmlCreateElement(rinvGECE, objDataEntry, "AirBlower_REQ", .Range("AirBlower_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "Autosynchronization_REQ", .Range("Autosynchronization_REQ").Value)
    'Call xmlCreateElement(rinvGECE, objDataEntry, "B35A_QTY", .Range("B35A_QTY").Value)
    'Call xmlCreateElement(rinvGECE, objDataEntry, "BN_REQ", .Range("BN_REQ").Value)
    'Call xmlCreateElement(rinvGECE, objDataEntry, "BoilerFeedwaterPump_REQ", .Range("BoilerFeedwaterPump_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "Compressor_QTY", .Range("Compressor_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DoubleExtraction_QTY", .Range("DoubleExtraction_QTY").Value)
    'Call xmlCreateElement(rinvGECE, objDataEntry, "FanDrive_REQ", .Range("FanDrive_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "GasTurbine_QTY", .Range("GasTurbine_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "Generator_QTY", .Range("Generator_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "LoadSharing_REQ", .Range("LoadSharing_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "LoadSharingGen_REQ", .Range("LoadSharingGen_REQ").Value)
    'Call xmlCreateElement(rinvGECE, objDataEntry, "MechanicalRetrofit_REQ", .Range("MechanicalRetrofit_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MotorDriven_QTY", .Range("MotorDriven_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MultiShaft_QTY", .Range("MultiShaft_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PowerSystemStabilizer_REQ", .Range("PowerSystemStabilizer_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PRV_Controls_REQ", .Range("PRV_Controls_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PRVValves_QTY", .Range("PRVValves_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "RecycleValves_QTY", .Range("RecycleValves_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "Reheat_QTY", .Range("Reheat_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_SYSTEM_REQ", .Range("ESD_SYSTEM_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CompressorType_QTY", .Range("CompressorType_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SingleExtraction_QTY", .Range("SingleExtraction_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SingleShaft_QTY", .Range("SingleShaft_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SteamTurbine_QTY", .Range("SteamTurbine_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SurgeControl_REQ", .Range("SurgeControl_REQ").Value)
    'Call xmlCreateElement(rinvGECE, objDataEntry, "TurboSentry_REQ", .Range("TurboSentry_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TypeCompressor", .Range("TypeCompressor").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_SYSTEMS_QTY", .Range("ESD_SYSTEMS_QTY").Value)


'    If Val(strCoverSheetVersion) >= Val("1.0") Then
'        If wsSource.Range("ESD_SYSTEM_REQ").Value = 1 Then
'            .Range("SystemType").Value = "ESD" ' $H$69
'        ElseIf wsSource.Range("ESD_SYSTEM_REQ").Value = 2 Then
'            .Range("SystemType").Value = "BMS" ' $H$69
'        Else
'            .Range("SystemType").Value = "TMC" ' $H$69
'        End If
'    Else
'        If wsSource.Range("ESD_SYSTEM_REQ").Value = True Then
'            .Range("SystemType").Value = "ESD" ' $H$69
'        Else
'            .Range("SystemType").Value = "BMS" ' $H$69
'        End If
'    End If
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_CLOSE", .Range("MEETING_CLOSE").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_DESIGN", .Range("MEETING_DESIGN").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_KICKOFF", .Range("MEETING_KICKOFF").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_OTHER", .Range("MEETING_OTHER").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_PROGRESS", .Range("MEETING_PROGRESS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "NO_OF_UNITS", .Range("NO_OF_UNITS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "RENTAL_COST", .Range("RENTAL_COST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_CUSTOM", .Range("REP_CUSTOM").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_MASS_HEAT", .Range("REP_MASS_HEAT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_STD", .Range("REP_STD").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REPORT_STD", .Range("REPORT_STD").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_COMM_HOURS", .Range("SITE_COMM_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_PWRUP_HOURS", .Range("SITE_PWRUP_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_SAT_HOURS", .Range("SITE_SAT_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_SURVEY_HOURS", .Range("SITE_SURVEY_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SW_SIMULATOR_REQ", .Range("SW_SIMULATOR_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYS_CONTROLLERS_QTY", .Range("SYS_CONTROLLERS_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYS_FBM_QTY", .Range("SYS_FBM_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYS_FDSI_QTY", .Range("SYS_FDSI_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYS_WORKSTATIONS_QTY", .Range("SYS_WORKSTATIONS_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_CUSTOMER_FAT", .Range("TEST_CUSTOMER_FAT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_FAT_PCT", .Range("TEST_FAT_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_PRE_FAT_PCT", .Range("TEST_PRE_FAT_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_AIRFARE", .Range("TL_AIRFARE").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_DAILY_ALLOW", .Range("TL_DAILY_ALLOW").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_DAYS_QTY", .Range("TL_DAYS_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_PERS_QTY", .Range("TL_PERS_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TRIPS_REQ", .Range("TL_TRIPS_REQ").Value)

End With

CleanUp:
    On Error Resume Next
    Set objDataEntry = Nothing


Exit Sub
ErrorHandler:
    MsgBox "error in ExportDataEntrySheet " & Err.Description, vbCritical
    Resume CleanUp
    Resume 0
    
End Sub
Sub ExportAssumptionsProposalSheet(wsTarget As Worksheet, ByRef rinvGECE As MSXML2.DOMDocument30, ByRef robjRoot As Object)
On Error Resume Next
Dim objDataEntry                 As Object
With wsTarget
    Set objDataEntry = xmlCreateSubRoot(rinvGECE, robjRoot, "AssumptionsProposal")
    Call xmlCreateElement(rinvGECE, objDataEntry, "TOOLKIT_1_REQ", .Range("TOOLKIT_1_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TOOLKIT_2_REQ", .Range("TOOLKIT_2_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TOOLKIT_3_REQ", .Range("TOOLKIT_3_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TOOLKIT_4_REQ", .Range("TOOLKIT_4_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TOOLKIT_5_REQ", .Range("TOOLKIT_5_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TOOLKIT_6_REQ", .Range("TOOLKIT_6_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TOOLKIT_7_REQ", .Range("TOOLKIT_7_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TOOLKIT_8_REQ", .Range("TOOLKIT_8_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TOOLKIT_9_REQ", .Range("TOOLKIT_9_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TOOLKIT_10_REQ", .Range("TOOLKIT_10_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TOOLKIT_11_REQ", .Range("TOOLKIT_11_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TOOLKIT_12_REQ", .Range("TOOLKIT_12_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TOOLKIT_13_REQ", .Range("TOOLKIT_13_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TOOLKIT_14_REQ", .Range("TOOLKIT_14_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TOOLKIT_15_REQ", .Range("TOOLKIT_15_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TOOLKIT_16_REQ", .Range("TOOLKIT_16_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TOOLKIT_17_REQ", .Range("TOOLKIT_17_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TOOLKIT_18_REQ", .Range("TOOLKIT_18_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TOOLKIT_19_REQ", .Range("TOOLKIT_19_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TOOLKIT_20_REQ", .Range("TOOLKIT_20_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TOOLKIT_21_REQ", .Range("TOOLKIT_21_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TOOLKIT_22_REQ", .Range("TOOLKIT_22_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TOOLKIT_23_REQ", .Range("TOOLKIT_23_REQ").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_NOTES", .Range("APP_NOTES").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ASSUMPTIONS", .Range("ASSUMPTIONS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_NOTES", .Range("CP_NOTES").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CUSTOMER_NAME", .Range("CUSTOMER_NAME").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CUSTOMER_TYPE", .Range("CUSTOMER_TYPE").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DEFAULT_REM_COUNTRY", .Range("DEFAULT_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_NOTES", .Range("DI_NOTES").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_NOTES", .Range("DOC_NOTES").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "GSP_ID", .Range("GSP_ID").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "INDUSTRY", .Range("INDUSTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "LOCAL_COUNTRY", .Range("LOCAL_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_NOTES", .Range("MEETING_NOTES").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PROJECT_MANAGER", .Range("PROJECT_MANAGER").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PROJECT_NAME", .Range("PROJECT_NAME").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PROPOSAL_NUMBER", .Range("PROPOSAL_NUMBER").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PROPOSAL_REVISION", .Range("PROPOSAL_REVISION").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PROPSAL_DATE", .Range("PROPSAL_DATE").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PROPSAL_ENGINEER", .Range("PROPSAL_ENGINEER").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REPORT_NOTES", .Range("REPORT_NOTES").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_NOTES", .Range("SITE_NOTES").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSTEM_NOTES", .Range("SYSTEM_NOTES").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_NOTES", .Range("TEST_NOTES").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_NOTES", .Range("TL_NOTES").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TRAIN_NOTES", .Range("TRAIN_NOTES").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TRICON_NOTES", .Range("TRICON_NOTES").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "WPA", .Range("WPA").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "WPA_TYPE", .Range("WPA_TYPE").Value)
    
End With
CleanUp:
    On Error Resume Next
    Set objDataEntry = Nothing


Exit Sub
ErrorHandler:
    MsgBox "error in ExportAssumptionsProposalSheet " & Err.Description, vbCritical
    Resume CleanUp
End Sub
Sub ExportDurationBasedSheet(wsTarget As Worksheet, ByRef rinvGECE As MSXML2.DOMDocument30, ByRef robjRoot As Object)

Dim objDataEntry                 As Object

On Error Resume Next
With wsTarget
    Set objDataEntry = xmlCreateSubRoot(rinvGECE, robjRoot, "DurationBasedSheet")
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_SE_QTY", .Range("DURATION_LOC_SE_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_SE_WEEK", .Range("DURATION_LOC_SE_WEEK").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_SE_DUP", .Range("DURATION_LOC_SE_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_SE_UTIL", .Range("DURATION_LOC_SE_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_SR_QTY", .Range("DURATION_LOC_SR_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_SR_WEEK", .Range("DURATION_LOC_SR_WEEK").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_SSE_DUP", .Range("DURATION_LOC_SSE_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_SR_UTIL", .Range("DURATION_LOC_SR_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_SE_QTY", .Range("DURATION_REM_SE_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_SE_WEEK", .Range("DURATION_REM_SE_WEEK").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_SE_REM_COUNTRY", .Range("DURATION_SE_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_SE_DUP", .Range("DURATION_REM_SE_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_SE_UTIL", .Range("DURATION_REM_SE_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_SR_QTY", .Range("DURATION_REM_SR_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_SR_WEEK", .Range("DURATION_REM_SR_WEEK").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_SR_REM_COUNTRY", .Range("DURATION_SR_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_SR_DUP", .Range("DURATION_REM_SR_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_SR_UTIL", .Range("DURATION_REM_SR_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_AC_REM_COUNTRY", .Range("DURATION_AC_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_AE_REM_COUNTRY", .Range("DURATION_AE_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_BASED_APP_PCT_QTY", .Range("DURATION_BASED_APP_PCT_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_BASED_CONTROLS_PCT_QTY", .Range("DURATION_BASED_CONTROLS_PCT_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_BASED_DESIGN_PCT_QTY", .Range("DURATION_BASED_DESIGN_PCT_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_BASED_DI_PCT_QTY", .Range("DURATION_BASED_DI_PCT_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_BASED_COURSE_PCT_QTY", .Range("DURATION_BASED_COURSE_PCT_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_BASED_DOC_PCT_QTY", .Range("DURATION_BASED_DOC_PCT_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_BASED_HMI_PCT_QTY", .Range("DURATION_BASED_HMI_PCT_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_BASED_IMPLEMENT_PCT_QTY", .Range("DURATION_BASED_IMPLEMENT_PCT_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_BASED_MEETINGS_PCT_QTY", .Range("DURATION_BASED_MEETINGS_PCT_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_BASED_PM_PCT_QTY", .Range("DURATION_BASED_PM_PCT_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_BASED_REP_PCT_QTY", .Range("DURATION_BASED_REP_PCT_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_BASED_REVIEW_PCT_QTY", .Range("DURATION_BASED_REVIEW_PCT_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_BASED_SITE_PCT_QTY", .Range("DURATION_BASED_SITE_PCT_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_BASED_SITE_SITE_PCT_QTY", .Range("DURATION_BASED_SITE_SITE_PCT_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_BASED_SPEC_PCT_QTY", .Range("DURATION_BASED_SPEC_PCT_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_BASED_SYS_ENG_PCT_QTY", .Range("DURATION_BASED_SYS_ENG_PCT_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_BASED_TEST_PCT_QTY", .Range("DURATION_BASED_TEST_PCT_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_BASED_TEST_TEST_PCT_QTY", .Range("DURATION_BASED_TEST_TEST_PCT_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_ESD_AC_LOC_HOURS", .Range("DURATION_ESD_AC_LOC_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_ESD_AC_REM_HOURS", .Range("DURATION_ESD_AC_REM_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_ESD_AE_LOC_HOURS", .Range("DURATION_ESD_AE_LOC_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_ESD_AE_REM_HOURS", .Range("DURATION_ESD_AE_REM_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_ESD_LE_LOC_HOURS", .Range("DURATION_ESD_LE_LOC_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_ESD_LE_REM_HOURS", .Range("DURATION_ESD_LE_REM_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_ESD_T_LOC_HOURS", .Range("DURATION_ESD_T_LOC_HOURS").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_ESD_T_REM_HOURS", .Range("DURATION_ESD_T_REM_HOURS").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_1_CAT", .Range("DURATION_FREE_1_CAT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_1_COST", .Range("DURATION_FREE_1_COST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_1_DUP", .Range("DURATION_FREE_1_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_1_QTY", .Range("DURATION_FREE_1_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_1_UTIL", .Range("DURATION_FREE_1_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_1_WEEK", .Range("DURATION_FREE_1_WEEK").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_2_CAT", .Range("DURATION_FREE_2_CAT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_2_COST", .Range("DURATION_FREE_2_COST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_2_DUP", .Range("DURATION_FREE_2_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_2_QTY", .Range("DURATION_FREE_2_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_2_UTIL", .Range("DURATION_FREE_2_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_2_WEEK", .Range("DURATION_FREE_2_WEEK").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_3_CAT", .Range("DURATION_FREE_3_CAT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_3_COST", .Range("DURATION_FREE_3_COST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_3_DUP", .Range("DURATION_FREE_3_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_3_QTY", .Range("DURATION_FREE_3_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_3_UTIL", .Range("DURATION_FREE_3_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_3_WEEK", .Range("DURATION_FREE_3_WEEK").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_4_CAT", .Range("DURATION_FREE_4_CAT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_4_COST", .Range("DURATION_FREE_4_COST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_4_DUP", .Range("DURATION_FREE_4_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_4_QTY", .Range("DURATION_FREE_4_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_4_UTIL", .Range("DURATION_FREE_4_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_4_WEEK", .Range("DURATION_FREE_4_WEEK").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_5_CAT", .Range("DURATION_FREE_5_CAT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_5_COST", .Range("DURATION_FREE_5_COST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_5_DUP", .Range("DURATION_FREE_5_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_5_QTY", .Range("DURATION_FREE_5_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_5_UTIL", .Range("DURATION_FREE_5_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_5_WEEK", .Range("DURATION_FREE_5_WEEK").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_6_CAT", .Range("DURATION_FREE_6_CAT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_6_COST", .Range("DURATION_FREE_6_COST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_6_DUP", .Range("DURATION_FREE_6_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_6_QTY", .Range("DURATION_FREE_6_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_6_UTIL", .Range("DURATION_FREE_6_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_6_WEEK", .Range("DURATION_FREE_6_WEEK").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_7_CAT", .Range("DURATION_FREE_7_CAT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_7_COST", .Range("DURATION_FREE_7_COST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_7_DUP", .Range("DURATION_FREE_7_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_7_QTY", .Range("DURATION_FREE_7_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_7_UTIL", .Range("DURATION_FREE_7_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_7_WEEK", .Range("DURATION_FREE_7_WEEK").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_8_CAT", .Range("DURATION_FREE_8_CAT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_8_COST", .Range("DURATION_FREE_8_COST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_8_DUP", .Range("DURATION_FREE_8_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_8_QTY", .Range("DURATION_FREE_8_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_8_UTIL", .Range("DURATION_FREE_8_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_8_WEEK", .Range("DURATION_FREE_8_WEEK").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_9_CAT", .Range("DURATION_FREE_9_CAT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_9_COST", .Range("DURATION_FREE_9_COST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_9_DUP", .Range("DURATION_FREE_9_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_9_QTY", .Range("DURATION_FREE_9_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_9_UTIL", .Range("DURATION_FREE_9_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_9_WEEK", .Range("DURATION_FREE_9_WEEK").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_10_CAT", .Range("DURATION_FREE_10_CAT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_10_COST", .Range("DURATION_FREE_10_COST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_10_DUP", .Range("DURATION_FREE_10_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_10_QTY", .Range("DURATION_FREE_10_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_10_UTIL", .Range("DURATION_FREE_10_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_10_WEEK", .Range("DURATION_FREE_10_WEEK").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_11_CAT", .Range("DURATION_FREE_11_CAT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_11_COST", .Range("DURATION_FREE_11_COST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_11_DUP", .Range("DURATION_FREE_11_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_11_QTY", .Range("DURATION_FREE_11_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_11_UTIL", .Range("DURATION_FREE_11_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_FREE_11_WEEK", .Range("DURATION_FREE_11_WEEK").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LE_REM_COUNTRY", .Range("DURATION_LE_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_AC_DUP", .Range("DURATION_LOC_AC_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_AC_QTY", .Range("DURATION_LOC_AC_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_AC_UTIL", .Range("DURATION_LOC_AC_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_AC_WEEK", .Range("DURATION_LOC_AC_WEEK").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_AE_DUP", .Range("DURATION_LOC_AE_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_AE_QTY", .Range("DURATION_LOC_AE_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_AE_QTY", .Range("DURATION_LOC_AE_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_AE_UTIL", .Range("DURATION_LOC_AE_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_AE_WEEK", .Range("DURATION_LOC_AE_WEEK").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_BA_DUP", .Range("DURATION_LOC_BA_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_BA_QTY", .Range("DURATION_LOC_BA_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_BA_UTIL", .Range("DURATION_LOC_BA_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_BA_WEEK", .Range("DURATION_LOC_BA_WEEK").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_LE_DUP", .Range("DURATION_LOC_LE_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_LE_QTY", .Range("DURATION_LOC_LE_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_LE_UTIL", .Range("DURATION_LOC_LE_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_LE_WEEK", .Range("DURATION_LOC_LE_WEEK").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_PM_DUP", .Range("DURATION_LOC_PM_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_PM_QTY", .Range("DURATION_LOC_PM_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_PM_UTIL", .Range("DURATION_LOC_PM_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_PM_WEEK", .Range("DURATION_LOC_PM_WEEK").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_LE_DUP", .Range("DURATION_LOC_LE_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_LE_QTY", .Range("DURATION_LOC_LE_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_LE_UTIL", .Range("DURATION_LOC_LE_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_LE_WEEK", .Range("DURATION_LOC_LE_WEEK").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_SITE_DUP", .Range("DURATION_LOC_SITE_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_SITE_QTY", .Range("DURATION_LOC_SITE_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_SITE_UTIL", .Range("DURATION_LOC_SITE_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_SITE_WEEK", .Range("DURATION_LOC_SITE_WEEK").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_STAGING_DUP", .Range("DURATION_LOC_STAGING_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_STAGING_QTY", .Range("DURATION_LOC_STAGING_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_STAGING_UTIL", .Range("DURATION_LOC_STAGING_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_STAGING_WEEK", .Range("DURATION_LOC_STAGING_WEEK").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_T_DUP", .Range("DURATION_LOC_T_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_T_QTY", .Range("DURATION_LOC_T_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_T_UTIL", .Range("DURATION_LOC_T_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_LOC_T_WEEK", .Range("DURATION_LOC_T_WEEK").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_PM_REM_COUNTRY", .Range("DURATION_PM_REM_COUNTRY").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_AC_DUP", .Range("DURATION_REM_AC_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_AC_QTY", .Range("DURATION_REM_AC_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_AC_UTIL", .Range("DURATION_REM_AC_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_AC_WEEK", .Range("DURATION_REM_AC_WEEK").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_AE_DUP", .Range("DURATION_REM_AE_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_AE_QTY", .Range("DURATION_REM_AE_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_AE_UTIL", .Range("DURATION_REM_AE_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_AE_WEEK", .Range("DURATION_REM_AE_WEEK").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_BA_DUP", .Range("DURATION_REM_BA_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_BA_QTY", .Range("DURATION_REM_BA_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_BA_UTIL", .Range("DURATION_REM_BA_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_BA_WEEK", .Range("DURATION_REM_BA_WEEK").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_LE_DUP", .Range("DURATION_REM_LE_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_LE_QTY", .Range("DURATION_REM_LE_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_LE_UTIL", .Range("DURATION_REM_LE_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_LE_WEEK", .Range("DURATION_REM_LE_WEEK").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_PM_DUP", .Range("DURATION_REM_PM_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_PM_QTY", .Range("DURATION_REM_PM_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_PM_UTIL", .Range("DURATION_REM_PM_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_PM_WEEK", .Range("DURATION_REM_PM_WEEK").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_SITE_DUP", .Range("DURATION_REM_SITE_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_SITE_QTY", .Range("DURATION_REM_SITE_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_SITE_UTIL", .Range("DURATION_REM_SITE_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_SITE_WEEK", .Range("DURATION_REM_SITE_WEEK").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_STAGING_DUP", .Range("DURATION_REM_STAGING_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_STAGING_QTY", .Range("DURATION_REM_STAGING_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_STAGING_UTIL", .Range("DURATION_REM_STAGING_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_STAGING_WEEK", .Range("DURATION_REM_STAGING_WEEK").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_T_DUP", .Range("DURATION_REM_T_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_T_QTY", .Range("DURATION_REM_T_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_T_UTIL", .Range("DURATION_REM_T_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_T_WEEK", .Range("DURATION_REM_T_WEEK").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_T_DUP", .Range("DURATION_REM_T_DUP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_T_QTY", .Range("DURATION_REM_T_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_T_UTIL", .Range("DURATION_REM_T_UTIL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_REM_T_WEEK", .Range("DURATION_REM_T_WEEK").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_T_REM_COUNTRY", .Range("DURATION_T_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_TOTAL_BUSINESS_REM_COUNTRY", .Range("DURATION_TOTAL_BUSINESS_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_TOTAL_SITE_REM_COUNTRY", .Range("DURATION_TOTAL_SITE_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DURATION_TOTAL_STAGING_REM_COUNTRY", .Range("DURATION_TOTAL_STAGING_REM_COUNTRY").Value)
    
End With

CleanUp:
    On Error Resume Next
    Set objDataEntry = Nothing


Exit Sub
ErrorHandler:
    MsgBox "error in ExportDurationBasedSheet " & Err.Description, vbCritical
    Resume CleanUp

End Sub
Sub ExportTaskBasedSheet(wsTarget As Worksheet, ByRef rinvGECE As MSXML2.DOMDocument30, ByRef robjRoot As Object)

Dim objDataEntry                 As Object
Dim strAddress                   As String
Dim strMaxRange                  As String
Dim c                            As Range


On Error Resume Next
strMaxRange = "A1:BR386"
With wsTarget
    Set objDataEntry = xmlCreateSubRoot(rinvGECE, robjRoot, "TaskBased")
    For Each c In .Range(strMaxRange).Cells
    'For Each c In Worksheets(strSheetName).Range(strMaxRange).Cells
        'if its a green cell, has no formula and is not locked then import it
        'old sheets had incorrect color for column F ad N so look for it as well
        'c.Interior.ColorIndex = 8
        
        If (c.Interior.ColorIndex = 4 Or c.Interior.ColorIndex = 44 Or c.Interior.ColorIndex = 8) And c.Locked = False Then
            
                strAddress = Replace(c.Address, "$", "_")
                Call xmlCreateElement(rinvGECE, objDataEntry, strAddress, .Range("" & c.Address & "").Value)
                '.Range("" & c.Address & "").Value = wsSource.Range("" & c.Address & "").Value
            
        End If
    Next
    
End With
CleanUp:
    On Error Resume Next
    Set objDataEntry = Nothing


Exit Sub
ErrorHandler:
    MsgBox "error in ExportTaskBasedSheet " & Err.Description, vbCritical
    Resume CleanUp
End Sub


Sub ExportApplicationBasedSheet(wsTarget As Worksheet, ByRef rinvGECE As MSXML2.DOMDocument30, ByRef robjRoot As Object)

Dim objDataEntry                 As Object

On Error Resume Next
With wsTarget
    Set objDataEntry = xmlCreateSubRoot(rinvGECE, robjRoot, "ApplicationBased")
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK1_OVD_JUST", .Range("APP_TASK1_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK1_OVD_QTY", .Range("APP_TASK1_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK1_REM_COUNTRY", .Range("APP_TASK1_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK10_OVD_JUST", .Range("APP_TASK10_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK10_OVD_QTY", .Range("APP_TASK10_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK10_REM_COUNTRY", .Range("APP_TASK10_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK10_REM_PCT", .Range("APP_TASK10_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK2_OVD_JUST", .Range("APP_TASK2_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK2_OVD_QTY", .Range("APP_TASK2_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK2_REM_COUNTRY", .Range("APP_TASK2_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK2_REM_PCT", .Range("APP_TASK2_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK3_OVD_JUST", .Range("APP_TASK3_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK3_OVD_QTY", .Range("APP_TASK3_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK3_REM_COUNTRY", .Range("APP_TASK3_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK3_REM_PCT", .Range("APP_TASK3_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK4_OVD_JUST", .Range("APP_TASK4_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK4_OVD_QTY", .Range("APP_TASK4_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK4_REM_COUNTRY", .Range("APP_TASK4_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK4_REM_PCT", .Range("APP_TASK4_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK5_OVD_JUST", .Range("APP_TASK5_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK5_OVD_QTY", .Range("APP_TASK5_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK5_REM_COUNTRY", .Range("APP_TASK5_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK5_REM_PCT", .Range("APP_TASK5_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK6_OVD_JUST", .Range("APP_TASK6_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK6_OVD_QTY", .Range("APP_TASK6_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK6_REM_COUNTRY", .Range("APP_TASK6_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK6_REM_PCT", .Range("APP_TASK6_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK7_OVD_JUST", .Range("APP_TASK7_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK7_OVD_QTY", .Range("APP_TASK7_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK7_REM_COUNTRY", .Range("APP_TASK7_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK7_REM_PCT", .Range("APP_TASK7_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK8_OVD_JUST", .Range("APP_TASK8_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK8_OVD_QTY", .Range("APP_TASK8_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK8_REM_COUNTRY", .Range("APP_TASK8_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK8_REM_PCT", .Range("APP_TASK8_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK9_OVD_JUST", .Range("APP_TASK9_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK9_OVD_QTY", .Range("APP_TASK9_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK9_REM_COUNTRY", .Range("APP_TASK9_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "APP_TASK9_REM_PCT", .Range("APP_TASK9_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_REM_COST", .Range("COURSE_REM_COST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK1_OVD_JUST", .Range("COURSE_TASK1_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK1_REM_COUNTRY", .Range("COURSE_TASK1_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK1_REM_PCT", .Range("COURSE_TASK1_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK10_OVD_JUST", .Range("COURSE_TASK10_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK10_OVD_QTY", .Range("COURSE_TASK10_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK10_REM_COUNTRY", .Range("COURSE_TASK10_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK1_REM_PCT", .Range("COURSE_TASK1_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK2_OVD_JUST", .Range("COURSE_TASK2_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK2_OVD_QTY", .Range("COURSE_TASK2_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK2_REM_COUNTRY", .Range("COURSE_TASK2_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK2_REM_PCT", .Range("COURSE_TASK2_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK3_OVD_JUST", .Range("COURSE_TASK3_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK3_OVD_QTY", .Range("COURSE_TASK3_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK3_REM_COUNTRY", .Range("COURSE_TASK3_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK3_REM_PCT", .Range("COURSE_TASK3_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK4_OVD_JUST", .Range("COURSE_TASK4_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK4_OVD_QTY", .Range("COURSE_TASK4_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK4_REM_COUNTRY", .Range("COURSE_TASK4_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK4_REM_PCT", .Range("COURSE_TASK4_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK5_OVD_JUST", .Range("COURSE_TASK5_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK5_OVD_QTY", .Range("COURSE_TASK5_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK5_REM_COUNTRY", .Range("COURSE_TASK5_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK5_REM_PCT", .Range("COURSE_TASK5_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK6_OVD_JUST", .Range("COURSE_TASK6_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK6_OVD_QTY", .Range("COURSE_TASK6_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK6_REM_COUNTRY", .Range("COURSE_TASK6_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK6_REM_PCT", .Range("COURSE_TASK6_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK7_OVD_JUST", .Range("COURSE_TASK7_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK7_OVD_QTY", .Range("COURSE_TASK7_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK7_REM_COUNTRY", .Range("COURSE_TASK7_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK7_REM_PCT", .Range("COURSE_TASK7_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK8_OVD_JUST", .Range("COURSE_TASK8_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK8_OVD_QTY", .Range("COURSE_TASK8_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK8_REM_COUNTRY", .Range("COURSE_TASK8_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK8_REM_PCT", .Range("COURSE_TASK8_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK9_OVD_JUST", .Range("COURSE_TASK9_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK9_OVD_QTY", .Range("COURSE_TASK9_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK9_REM_COUNTRY", .Range("COURSE_TASK9_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "COURSE_TASK9_REM_PCT", .Range("COURSE_TASK9_REM_PCT").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK1_OVD_JUST", .Range("CP_TASK1_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK1_OVD_QTY", .Range("CP_TASK1_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK1_REM_COUNTRY", .Range("CP_TASK1_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK1_REM_PCT", .Range("CP_TASK1_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK2_OVD_JUST", .Range("CP_TASK2_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK2_OVD_QTY", .Range("CP_TASK2_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK2_REM_COUNTRY", .Range("CP_TASK2_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK2_REM_PCT", .Range("CP_TASK2_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK3_OVD_JUST", .Range("CP_TASK3_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK3_OVD_QTY", .Range("CP_TASK3_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK3_REM_COUNTRY", .Range("CP_TASK3_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK3_REM_PCT", .Range("CP_TASK3_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK4_OVD_JUST", .Range("CP_TASK4_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK4_OVD_QTY", .Range("CP_TASK4_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK4_REM_COUNTRY", .Range("CP_TASK4_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK4_REM_PCT", .Range("CP_TASK4_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK5_OVD_JUST", .Range("CP_TASK5_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK5_OVD_QTY", .Range("CP_TASK5_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK5_REM_COUNTRY", .Range("CP_TASK5_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK5_REM_PCT", .Range("CP_TASK5_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK6_OVD_JUST", .Range("CP_TASK6_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK6_OVD_QTY", .Range("CP_TASK6_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK6_REM_COUNTRY", .Range("CP_TASK6_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK6_REM_PCT", .Range("CP_TASK6_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK7_OVD_JUST", .Range("CP_TASK7_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK7_OVD_QTY", .Range("CP_TASK7_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK7_REM_COUNTRY", .Range("CP_TASK7_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK7_REM_PCT", .Range("CP_TASK7_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK8_OVD_JUST", .Range("CP_TASK8_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK8_OVD_QTY", .Range("CP_TASK8_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK8_REM_COUNTRY", .Range("CP_TASK8_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK8_REM_PCT", .Range("CP_TASK8_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK9_OVD_JUST", .Range("CP_TASK9_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK9_OVD_QTY", .Range("CP_TASK9_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK9_REM_COUNTRY", .Range("CP_TASK9_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK9_REM_PCT", .Range("CP_TASK9_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK10_OVD_JUST", .Range("CP_TASK10_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK10_OVD_QTY", .Range("CP_TASK10_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK10_REM_COUNTRY", .Range("CP_TASK10_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "CP_TASK10_REM_PCT", .Range("CP_TASK10_REM_PCT").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_REM_COST", .Range("DI_REM_COST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK1_OVD_JUST", .Range("DI_TASK1_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK1_OVD_QTY", .Range("DI_TASK1_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK1_REM_COUNTRY", .Range("DI_TASK1_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK1_REM_PCT", .Range("DI_TASK1_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK2_OVD_JUST", .Range("DI_TASK2_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK2_OVD_QTY", .Range("DI_TASK2_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK2_REM_COUNTRY", .Range("DI_TASK2_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK2_REM_PCT", .Range("DI_TASK2_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK3_OVD_JUST", .Range("DI_TASK3_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK3_OVD_QTY", .Range("DI_TASK3_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK3_REM_COUNTRY", .Range("DI_TASK3_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK3_REM_PCT", .Range("DI_TASK3_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK4_OVD_JUST", .Range("DI_TASK4_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK4_OVD_QTY", .Range("DI_TASK4_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK4_REM_COUNTRY", .Range("DI_TASK4_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK4_REM_PCT", .Range("DI_TASK4_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK5_OVD_JUST", .Range("DI_TASK5_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK5_OVD_QTY", .Range("DI_TASK5_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK5_REM_COUNTRY", .Range("DI_TASK5_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK5_REM_PCT", .Range("DI_TASK5_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK6_OVD_JUST", .Range("DI_TASK6_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK6_OVD_QTY", .Range("DI_TASK6_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK6_REM_COUNTRY", .Range("DI_TASK6_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK6_REM_PCT", .Range("DI_TASK6_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK7_OVD_JUST", .Range("DI_TASK7_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK7_OVD_QTY", .Range("DI_TASK7_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK7_REM_COUNTRY", .Range("DI_TASK7_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK7_REM_PCT", .Range("DI_TASK7_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK8_OVD_JUST", .Range("DI_TASK8_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK8_OVD_QTY", .Range("DI_TASK8_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK8_REM_COUNTRY", .Range("DI_TASK8_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK8_REM_PCT", .Range("DI_TASK8_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK9_OVD_JUST", .Range("DI_TASK9_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK9_OVD_QTY", .Range("DI_TASK9_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK9_REM_COUNTRY", .Range("DI_TASK9_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK9_REM_PCT", .Range("DI_TASK9_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK10_OVD_JUST", .Range("DI_TASK10_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK10_OVD_QTY", .Range("DI_TASK10_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK10_REM_COUNTRY", .Range("DI_TASK10_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DI_TASK10_REM_PCT", .Range("DI_TASK10_REM_PCT").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_REM_COST", .Range("DOC_REM_COST").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK1_OVD_JUST", .Range("DOC_TASK1_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK1_OVD_QTY", .Range("DOC_TASK1_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK1_REM_COUNTRY", .Range("DOC_TASK1_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK1_REM_PCT", .Range("DOC_TASK1_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK2_OVD_JUST", .Range("DOC_TASK2_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK2_OVD_QTY", .Range("DOC_TASK2_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK2_REM_COUNTRY", .Range("DOC_TASK2_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK2_REM_PCT", .Range("DOC_TASK2_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK3_OVD_JUST", .Range("DOC_TASK3_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK3_OVD_QTY", .Range("DOC_TASK3_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK3_REM_COUNTRY", .Range("DOC_TASK3_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK3_REM_PCT", .Range("DOC_TASK3_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK4_OVD_JUST", .Range("DOC_TASK4_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK4_OVD_QTY", .Range("DOC_TASK4_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK4_REM_COUNTRY", .Range("DOC_TASK4_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK4_REM_PCT", .Range("DOC_TASK4_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK5_OVD_JUST", .Range("DOC_TASK5_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK5_OVD_QTY", .Range("DOC_TASK5_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK5_REM_COUNTRY", .Range("DOC_TASK5_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK5_REM_PCT", .Range("DOC_TASK5_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK6_OVD_JUST", .Range("DOC_TASK6_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK6_OVD_QTY", .Range("DOC_TASK6_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK6_REM_COUNTRY", .Range("DOC_TASK6_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK6_REM_PCT", .Range("DOC_TASK6_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK7_OVD_JUST", .Range("DOC_TASK7_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK7_OVD_QTY", .Range("DOC_TASK7_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK7_REM_COUNTRY", .Range("DOC_TASK7_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK7_REM_PCT", .Range("DOC_TASK7_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK8_OVD_JUST", .Range("DOC_TASK8_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK8_OVD_QTY", .Range("DOC_TASK8_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK8_REM_COUNTRY", .Range("DOC_TASK8_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK8_REM_PCT", .Range("DOC_TASK8_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK9_OVD_JUST", .Range("DOC_TASK9_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK9_OVD_QTY", .Range("DOC_TASK9_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK9_REM_COUNTRY", .Range("DOC_TASK9_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK9_REM_PCT", .Range("DOC_TASK9_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK10_OVD_JUST", .Range("DOC_TASK10_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK10_OVD_QTY", .Range("DOC_TASK10_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK10_REM_COUNTRY", .Range("DOC_TASK10_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DOC_TASK10_REM_PCT", .Range("DOC_TASK10_REM_PCT").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK1_OVD_JUST", .Range("ESD_TASK1_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK1_OVD_QTY", .Range("ESD_TASK1_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK1_REM_COUNTRY", .Range("ESD_TASK1_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK1_REM_PCT", .Range("ESD_TASK1_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK2_OVD_JUST", .Range("ESD_TASK2_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK2_OVD_QTY", .Range("ESD_TASK2_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK2_REM_COUNTRY", .Range("ESD_TASK2_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK2_REM_PCT", .Range("ESD_TASK2_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK3_OVD_JUST", .Range("ESD_TASK3_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK3_OVD_QTY", .Range("ESD_TASK3_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK3_REM_COUNTRY", .Range("ESD_TASK3_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK3_REM_PCT", .Range("ESD_TASK3_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK4_OVD_JUST", .Range("ESD_TASK4_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK4_OVD_QTY", .Range("ESD_TASK4_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK4_REM_COUNTRY", .Range("ESD_TASK4_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK4_REM_PCT", .Range("ESD_TASK4_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK5_OVD_JUST", .Range("ESD_TASK5_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK5_OVD_QTY", .Range("ESD_TASK5_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK5_REM_COUNTRY", .Range("ESD_TASK5_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK5_REM_PCT", .Range("ESD_TASK5_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK6_OVD_JUST", .Range("ESD_TASK6_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK6_OVD_QTY", .Range("ESD_TASK6_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK6_REM_COUNTRY", .Range("ESD_TASK6_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK6_REM_PCT", .Range("ESD_TASK6_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK7_OVD_JUST", .Range("ESD_TASK7_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK7_OVD_QTY", .Range("ESD_TASK7_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK7_REM_COUNTRY", .Range("ESD_TASK7_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK7_REM_PCT", .Range("ESD_TASK7_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK8_OVD_JUST", .Range("ESD_TASK8_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK8_OVD_QTY", .Range("ESD_TASK8_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK8_REM_COUNTRY", .Range("ESD_TASK8_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK8_REM_PCT", .Range("ESD_TASK8_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK9_OVD_JUST", .Range("ESD_TASK9_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK9_OVD_QTY", .Range("ESD_TASK9_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK9_REM_COUNTRY", .Range("ESD_TASK9_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK9_REM_PCT", .Range("ESD_TASK9_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK10_OVD_JUST", .Range("ESD_TASK10_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK10_OVD_QTY", .Range("ESD_TASK10_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK10_REM_COUNTRY", .Range("ESD_TASK10_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "ESD_TASK10_REM_PCT", .Range("ESD_TASK10_REM_PCT").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK1_OVD_JUST", .Range("HMI_TASK1_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK1_OVD_QTY", .Range("HMI_TASK1_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK1_REM_COUNTRY", .Range("HMI_TASK1_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK1_REM_PCT", .Range("HMI_TASK1_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK2_OVD_JUST", .Range("HMI_TASK2_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK2_OVD_QTY", .Range("HMI_TASK2_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK2_REM_COUNTRY", .Range("HMI_TASK2_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK2_REM_PCT", .Range("HMI_TASK2_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK3_OVD_JUST", .Range("HMI_TASK3_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK3_OVD_QTY", .Range("HMI_TASK3_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK3_REM_COUNTRY", .Range("HMI_TASK3_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK3_REM_PCT", .Range("HMI_TASK3_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK4_OVD_JUST", .Range("HMI_TASK4_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK4_OVD_QTY", .Range("HMI_TASK4_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK4_REM_COUNTRY", .Range("HMI_TASK4_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK4_REM_PCT", .Range("HMI_TASK4_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK5_OVD_JUST", .Range("HMI_TASK5_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK5_OVD_QTY", .Range("HMI_TASK5_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK5_REM_COUNTRY", .Range("HMI_TASK5_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK5_REM_PCT", .Range("HMI_TASK5_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK6_OVD_JUST", .Range("HMI_TASK6_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK6_OVD_QTY", .Range("HMI_TASK6_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK6_REM_COUNTRY", .Range("HMI_TASK6_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK6_REM_PCT", .Range("HMI_TASK6_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK7_OVD_JUST", .Range("HMI_TASK7_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK7_OVD_QTY", .Range("HMI_TASK7_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK7_REM_COUNTRY", .Range("HMI_TASK7_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK7_REM_PCT", .Range("HMI_TASK7_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK8_OVD_JUST", .Range("HMI_TASK8_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK8_OVD_QTY", .Range("HMI_TASK8_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK8_REM_COUNTRY", .Range("HMI_TASK8_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK8_REM_PCT", .Range("HMI_TASK8_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK9_OVD_JUST", .Range("HMI_TASK9_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK9_OVD_QTY", .Range("HMI_TASK9_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK9_REM_COUNTRY", .Range("HMI_TASK9_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK9_REM_PCT", .Range("HMI_TASK9_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK10_OVD_JUST", .Range("HMI_TASK10_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK10_OVD_QTY", .Range("HMI_TASK10_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK10_REM_COUNTRY", .Range("HMI_TASK10_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_TASK10_REM_PCT", .Range("HMI_TASK10_REM_PCT").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK1_OVD_JUST", .Range("MEETING_TASK1_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK1_OVD_QTY", .Range("MEETING_TASK1_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK1_REM_COUNTRY", .Range("MEETING_TASK1_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK1_REM_PCT", .Range("MEETING_TASK1_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK2_OVD_JUST", .Range("MEETING_TASK2_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK2_OVD_QTY", .Range("MEETING_TASK2_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK2_REM_COUNTRY", .Range("MEETING_TASK2_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK2_REM_PCT", .Range("MEETING_TASK2_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK3_OVD_JUST", .Range("MEETING_TASK3_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK3_OVD_QTY", .Range("MEETING_TASK3_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK3_REM_COUNTRY", .Range("MEETING_TASK3_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK3_REM_PCT", .Range("MEETING_TASK3_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK4_OVD_JUST", .Range("MEETING_TASK4_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK4_OVD_QTY", .Range("MEETING_TASK4_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK4_REM_COUNTRY", .Range("MEETING_TASK4_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK4_REM_PCT", .Range("MEETING_TASK4_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK5_OVD_JUST", .Range("MEETING_TASK5_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK5_OVD_QTY", .Range("MEETING_TASK5_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK5_REM_COUNTRY", .Range("MEETING_TASK5_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK5_REM_PCT", .Range("MEETING_TASK5_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK6_OVD_JUST", .Range("MEETING_TASK6_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK6_OVD_QTY", .Range("MEETING_TASK6_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK6_REM_COUNTRY", .Range("MEETING_TASK6_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK6_REM_PCT", .Range("MEETING_TASK6_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK7_OVD_JUST", .Range("MEETING_TASK7_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK7_OVD_QTY", .Range("MEETING_TASK7_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK7_REM_COUNTRY", .Range("MEETING_TASK7_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK7_REM_PCT", .Range("MEETING_TASK7_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK8_OVD_JUST", .Range("MEETING_TASK8_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK8_OVD_QTY", .Range("MEETING_TASK8_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK8_REM_COUNTRY", .Range("MEETING_TASK8_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK8_REM_PCT", .Range("MEETING_TASK8_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK8_OVD_JUST", .Range("MEETING_TASK8_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK8_OVD_QTY", .Range("MEETING_TASK8_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK8_REM_COUNTRY", .Range("MEETING_TASK8_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK8_REM_PCT", .Range("MEETING_TASK8_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK9_OVD_JUST", .Range("MEETING_TASK9_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK9_OVD_QTY", .Range("MEETING_TASK9_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK9_REM_COUNTRY", .Range("MEETING_TASK9_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK9_REM_PCT", .Range("MEETING_TASK9_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK10_OVD_JUST", .Range("MEETING_TASK10_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK10_OVD_QTY", .Range("MEETING_TASK10_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK10_REM_COUNTRY", .Range("MEETING_TASK10_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "MEETING_TASK10_REM_PCT", .Range("MEETING_TASK10_REM_PCT").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK1_OVD_JUST", .Range("PM_TASK1_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK1_OVD_QTY", .Range("PM_TASK1_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK1_REM_COUNTRY", .Range("PM_TASK1_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK1_REM_PCT", .Range("PM_TASK1_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK2_OVD_JUST", .Range("PM_TASK2_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK2_OVD_QTY", .Range("PM_TASK2_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK2_REM_COUNTRY", .Range("PM_TASK2_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK2_REM_PCT", .Range("PM_TASK2_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK3_OVD_JUST", .Range("PM_TASK3_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK3_OVD_QTY", .Range("PM_TASK3_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK3_REM_COUNTRY", .Range("PM_TASK3_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK3_REM_PCT", .Range("PM_TASK3_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK4_OVD_JUST", .Range("PM_TASK4_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK4_OVD_QTY", .Range("PM_TASK4_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK4_REM_COUNTRY", .Range("PM_TASK4_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK4_REM_PCT", .Range("PM_TASK4_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK5_OVD_JUST", .Range("PM_TASK5_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK5_OVD_QTY", .Range("PM_TASK5_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK5_REM_COUNTRY", .Range("PM_TASK5_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK5_REM_PCT", .Range("PM_TASK5_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK6_OVD_JUST", .Range("PM_TASK6_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK6_OVD_QTY", .Range("PM_TASK6_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK6_REM_COUNTRY", .Range("PM_TASK6_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK6_REM_PCT", .Range("PM_TASK6_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK7_OVD_JUST", .Range("PM_TASK7_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK7_OVD_QTY", .Range("PM_TASK7_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK7_REM_COUNTRY", .Range("PM_TASK7_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK7_REM_PCT", .Range("PM_TASK7_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK8_OVD_JUST", .Range("PM_TASK8_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK8_OVD_QTY", .Range("PM_TASK8_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK8_REM_COUNTRY", .Range("PM_TASK8_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK8_REM_PCT", .Range("PM_TASK8_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK9_OVD_JUST", .Range("PM_TASK9_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK9_OVD_QTY", .Range("PM_TASK9_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK9_REM_COUNTRY", .Range("PM_TASK9_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK9_REM_PCT", .Range("PM_TASK9_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK10_OVD_JUST", .Range("PM_TASK10_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK10_OVD_QTY", .Range("PM_TASK10_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK10_REM_COUNTRY", .Range("PM_TASK10_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_TASK10_REM_PCT", .Range("PM_TASK10_REM_PCT").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK1_OVD_JUST", .Range("REP_TASK1_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK1_OVD_QTY", .Range("REP_TASK1_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK1_REM_COUNTRY", .Range("REP_TASK1_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK1_REM_PCT", .Range("REP_TASK1_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK2_OVD_JUST", .Range("REP_TASK2_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK2_OVD_QTY", .Range("REP_TASK2_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK2_REM_COUNTRY", .Range("REP_TASK2_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK2_REM_PCT", .Range("REP_TASK2_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK3_OVD_JUST", .Range("REP_TASK3_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK3_OVD_QTY", .Range("REP_TASK3_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK3_REM_COUNTRY", .Range("REP_TASK3_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK3_REM_PCT", .Range("REP_TASK3_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK4_OVD_JUST", .Range("REP_TASK4_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK4_OVD_QTY", .Range("REP_TASK4_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK4_REM_COUNTRY", .Range("REP_TASK4_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK4_REM_PCT", .Range("REP_TASK4_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK5_OVD_JUST", .Range("REP_TASK5_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK5_OVD_QTY", .Range("REP_TASK5_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK5_REM_COUNTRY", .Range("REP_TASK5_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK5_REM_PCT", .Range("REP_TASK5_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK6_OVD_JUST", .Range("REP_TASK6_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK6_OVD_QTY", .Range("REP_TASK6_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK6_REM_COUNTRY", .Range("REP_TASK6_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK6_REM_PCT", .Range("REP_TASK6_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK7_OVD_JUST", .Range("REP_TASK7_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK7_OVD_QTY", .Range("REP_TASK7_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK7_REM_COUNTRY", .Range("REP_TASK7_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK7_REM_PCT", .Range("REP_TASK7_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK8_OVD_JUST", .Range("REP_TASK8_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK8_OVD_QTY", .Range("REP_TASK8_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK8_REM_COUNTRY", .Range("REP_TASK8_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK8_REM_PCT", .Range("REP_TASK8_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK9_OVD_JUST", .Range("REP_TASK9_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK9_OVD_QTY", .Range("REP_TASK9_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK9_REM_COUNTRY", .Range("REP_TASK9_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK9_REM_PCT", .Range("REP_TASK9_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK10_OVD_JUST", .Range("REP_TASK10_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK10_OVD_QTY", .Range("REP_TASK10_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK10_REM_COUNTRY", .Range("REP_TASK10_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "REP_TASK10_REM_PCT", .Range("REP_TASK10_REM_PCT").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK1_OVD_JUST", .Range("SITE_TASK1_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK1_OVD_QTY", .Range("SITE_TASK1_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK1_REM_COUNTRY", .Range("SITE_TASK1_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK1_REM_PCT", .Range("SITE_TASK1_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK2_OVD_JUST", .Range("SITE_TASK2_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK2_OVD_QTY", .Range("SITE_TASK2_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK2_REM_COUNTRY", .Range("SITE_TASK2_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK2_REM_PCT", .Range("SITE_TASK2_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK3_OVD_JUST", .Range("SITE_TASK3_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK3_OVD_QTY", .Range("SITE_TASK3_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK3_REM_COUNTRY", .Range("SITE_TASK3_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK3_REM_PCT", .Range("SITE_TASK3_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK4_OVD_JUST", .Range("SITE_TASK4_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK4_OVD_QTY", .Range("SITE_TASK4_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK4_REM_COUNTRY", .Range("SITE_TASK4_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK4_REM_PCT", .Range("SITE_TASK4_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK5_OVD_JUST", .Range("SITE_TASK5_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK5_OVD_QTY", .Range("SITE_TASK5_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK5_REM_COUNTRY", .Range("SITE_TASK5_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK5_REM_PCT", .Range("SITE_TASK5_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK6_OVD_JUST", .Range("SITE_TASK6_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK6_OVD_QTY", .Range("SITE_TASK6_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK6_REM_COUNTRY", .Range("SITE_TASK6_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK6_REM_PCT", .Range("SITE_TASK6_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK7_OVD_JUST", .Range("SITE_TASK7_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK7_OVD_QTY", .Range("SITE_TASK7_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK7_REM_COUNTRY", .Range("SITE_TASK7_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK7_REM_PCT", .Range("SITE_TASK7_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK8_OVD_JUST", .Range("SITE_TASK8_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK8_OVD_QTY", .Range("SITE_TASK8_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK8_REM_COUNTRY", .Range("SITE_TASK8_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK8_REM_PCT", .Range("SITE_TASK8_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK9_OVD_JUST", .Range("SITE_TASK9_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK9_OVD_QTY", .Range("SITE_TASK9_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK9_REM_COUNTRY", .Range("SITE_TASK9_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK9_REM_PCT", .Range("SITE_TASK9_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK10_OVD_JUST", .Range("SITE_TASK10_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK10_OVD_QTY", .Range("SITE_TASK10_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK10_REM_COUNTRY", .Range("SITE_TASK10_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SITE_TASK10_REM_PCT", .Range("SITE_TASK10_REM_PCT").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK1_OVD_JUST", .Range("SPEC_TASK1_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK1_OVD_QTY", .Range("SPEC_TASK1_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK1_REM_COUNTRY", .Range("SPEC_TASK1_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK1_REM_PCT", .Range("SPEC_TASK1_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK2_OVD_JUST", .Range("SPEC_TASK2_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK2_OVD_QTY", .Range("SPEC_TASK2_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK2_REM_COUNTRY", .Range("SPEC_TASK2_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK2_REM_PCT", .Range("SPEC_TASK2_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK3_OVD_JUST", .Range("SPEC_TASK3_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK3_OVD_QTY", .Range("SPEC_TASK3_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK3_REM_COUNTRY", .Range("SPEC_TASK3_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK3_REM_PCT", .Range("SPEC_TASK3_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK4_OVD_JUST", .Range("SPEC_TASK4_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK4_OVD_QTY", .Range("SPEC_TASK4_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK4_REM_COUNTRY", .Range("SPEC_TASK4_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK4_REM_PCT", .Range("SPEC_TASK4_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK5_OVD_JUST", .Range("SPEC_TASK5_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK5_OVD_QTY", .Range("SPEC_TASK5_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK5_REM_COUNTRY", .Range("SPEC_TASK5_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK5_REM_PCT", .Range("SPEC_TASK5_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK6_OVD_JUST", .Range("SPEC_TASK6_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK6_OVD_QTY", .Range("SPEC_TASK6_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK6_REM_COUNTRY", .Range("SPEC_TASK6_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK6_REM_PCT", .Range("SPEC_TASK6_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK7_OVD_JUST", .Range("SPEC_TASK7_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK7_OVD_QTY", .Range("SPEC_TASK7_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK7_REM_COUNTRY", .Range("SPEC_TASK7_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK7_REM_PCT", .Range("SPEC_TASK7_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK8_OVD_JUST", .Range("SPEC_TASK8_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK8_OVD_QTY", .Range("SPEC_TASK8_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK8_REM_COUNTRY", .Range("SPEC_TASK8_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK8_REM_PCT", .Range("SPEC_TASK8_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK9_OVD_JUST", .Range("SPEC_TASK9_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK9_OVD_QTY", .Range("SPEC_TASK9_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK9_REM_COUNTRY", .Range("SPEC_TASK9_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK9_REM_PCT", .Range("SPEC_TASK9_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK10_OVD_JUST", .Range("SPEC_TASK10_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK10_OVD_QTY", .Range("SPEC_TASK10_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK10_REM_COUNTRY", .Range("SPEC_TASK10_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SPEC_TASK10_REM_PCT", .Range("SPEC_TASK10_REM_PCT").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK1_OVD_JUST", .Range("SYSENG_TASK1_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK1_OVD_QTY", .Range("SYSENG_TASK1_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK1_REM_COUNTRY", .Range("SYSENG_TASK1_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK1_REM_PCT", .Range("SYSENG_TASK1_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK2_OVD_JUST", .Range("SYSENG_TASK2_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK2_OVD_QTY", .Range("SYSENG_TASK2_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK2_REM_COUNTRY", .Range("SYSENG_TASK2_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK2_REM_PCT", .Range("SYSENG_TASK2_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK3_OVD_JUST", .Range("SYSENG_TASK3_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK3_OVD_QTY", .Range("SYSENG_TASK3_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK3_REM_COUNTRY", .Range("SYSENG_TASK3_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK3_REM_PCT", .Range("SYSENG_TASK3_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK4_OVD_JUST", .Range("SYSENG_TASK4_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK4_OVD_QTY", .Range("SYSENG_TASK4_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK4_REM_COUNTRY", .Range("SYSENG_TASK4_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK4_REM_PCT", .Range("SYSENG_TASK4_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK5_OVD_JUST", .Range("SYSENG_TASK5_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK5_OVD_QTY", .Range("SYSENG_TASK5_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK5_REM_COUNTRY", .Range("SYSENG_TASK5_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK5_REM_PCT", .Range("SYSENG_TASK5_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK6_OVD_JUST", .Range("SYSENG_TASK6_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK6_OVD_QTY", .Range("SYSENG_TASK6_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK6_REM_COUNTRY", .Range("SYSENG_TASK6_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK6_REM_PCT", .Range("SYSENG_TASK6_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK7_OVD_JUST", .Range("SYSENG_TASK7_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK7_OVD_QTY", .Range("SYSENG_TASK7_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK7_REM_COUNTRY", .Range("SYSENG_TASK7_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK7_REM_PCT", .Range("SYSENG_TASK7_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK8_OVD_JUST", .Range("SYSENG_TASK8_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK8_OVD_QTY", .Range("SYSENG_TASK8_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK8_REM_COUNTRY", .Range("SYSENG_TASK8_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK8_REM_PCT", .Range("SYSENG_TASK8_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK9_OVD_JUST", .Range("SYSENG_TASK9_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK9_OVD_QTY", .Range("SYSENG_TASK9_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK9_REM_COUNTRY", .Range("SYSENG_TASK9_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK9_REM_PCT", .Range("SYSENG_TASK9_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK10_OVD_JUST", .Range("SYSENG_TASK10_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK10_OVD_QTY", .Range("SYSENG_TASK10_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK10_REM_COUNTRY", .Range("SYSENG_TASK10_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "SYSENG_TASK10_REM_PCT", .Range("SYSENG_TASK10_REM_PCT").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK1_OVD_JUST", .Range("TEST_TASK1_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK1_OVD_QTY", .Range("TEST_TASK1_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK1_REM_COUNTRY", .Range("TEST_TASK1_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK1_REM_PCT", .Range("TEST_TASK1_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK2_OVD_JUST", .Range("TEST_TASK2_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK2_OVD_QTY", .Range("TEST_TASK2_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK2_REM_COUNTRY", .Range("TEST_TASK2_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK2_REM_PCT", .Range("TEST_TASK2_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK3_OVD_JUST", .Range("TEST_TASK3_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK3_OVD_QTY", .Range("TEST_TASK3_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK3_REM_COUNTRY", .Range("TEST_TASK3_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK3_REM_PCT", .Range("TEST_TASK3_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK4_OVD_JUST", .Range("TEST_TASK4_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK4_OVD_QTY", .Range("TEST_TASK4_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK4_REM_COUNTRY", .Range("TEST_TASK4_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK4_REM_PCT", .Range("TEST_TASK4_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK5_OVD_JUST", .Range("TEST_TASK5_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK5_OVD_QTY", .Range("TEST_TASK5_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK5_REM_COUNTRY", .Range("TEST_TASK5_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK5_REM_PCT", .Range("TEST_TASK5_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK6_OVD_JUST", .Range("TEST_TASK6_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK6_OVD_QTY", .Range("TEST_TASK6_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK6_REM_COUNTRY", .Range("TEST_TASK6_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK6_REM_PCT", .Range("TEST_TASK6_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK7_OVD_JUST", .Range("TEST_TASK7_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK7_OVD_QTY", .Range("TEST_TASK7_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK7_REM_COUNTRY", .Range("TEST_TASK7_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK7_REM_PCT", .Range("TEST_TASK7_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK8_OVD_JUST", .Range("TEST_TASK8_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK8_OVD_QTY", .Range("TEST_TASK8_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK8_REM_COUNTRY", .Range("TEST_TASK8_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK8_REM_PCT", .Range("TEST_TASK8_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK9_OVD_JUST", .Range("TEST_TASK9_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK9_OVD_QTY", .Range("TEST_TASK9_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK9_REM_COUNTRY", .Range("TEST_TASK9_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK9_REM_PCT", .Range("TEST_TASK9_REM_PCT").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK10_OVD_JUST", .Range("TEST_TASK10_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK10_OVD_QTY", .Range("TEST_TASK10_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK10_REM_COUNTRY", .Range("TEST_TASK10_REM_COUNTRY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TEST_TASK10_REM_PCT", .Range("TEST_TASK10_REM_PCT").Value)
    
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK1_OVD_JUST", .Range("TL_TASK1_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK1_OVD_QTY", .Range("TL_TASK1_OVD_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK2_OVD_JUST", .Range("TL_TASK2_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK2_OVD_JUST", .Range("TL_TASK2_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK3_OVD_JUST", .Range("TL_TASK3_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK3_OVD_JUST", .Range("TL_TASK3_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK4_OVD_JUST", .Range("TL_TASK4_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK4_OVD_JUST", .Range("TL_TASK4_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK5_OVD_JUST", .Range("TL_TASK5_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK5_OVD_JUST", .Range("TL_TASK5_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK6_OVD_JUST", .Range("TL_TASK6_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK6_OVD_JUST", .Range("TL_TASK6_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK7_OVD_JUST", .Range("TL_TASK7_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK7_OVD_JUST", .Range("TL_TASK7_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK8_OVD_JUST", .Range("TL_TASK8_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK8_OVD_JUST", .Range("TL_TASK8_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK9_OVD_JUST", .Range("TL_TASK9_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK9_OVD_JUST", .Range("TL_TASK9_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK10_OVD_JUST", .Range("TL_TASK10_OVD_JUST").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "TL_TASK10_OVD_JUST", .Range("TL_TASK10_OVD_JUST").Value)
    
 'Formula: =G242
    Call xmlCreateElement(rinvGECE, objDataEntry, "DUP_FACT_APP_BASED_APP", .Range("DUP_FACT_APP_BASED_APP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DUP_FACT_APP_BASED_COURSE", .Range("DUP_FACT_APP_BASED_COURSE").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DUP_FACT_APP_BASED_CP", .Range("DUP_FACT_APP_BASED_CP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DUP_FACT_APP_BASED_DI", .Range("DUP_FACT_APP_BASED_DI").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DUP_FACT_APP_BASED_DOC", .Range("DUP_FACT_APP_BASED_DOC").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DUP_FACT_APP_BASED_ESD", .Range("DUP_FACT_APP_BASED_ESD").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DUP_FACT_APP_BASED_HMI", .Range("DUP_FACT_APP_BASED_HMI").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DUP_FACT_APP_BASED_MEETING", .Range("DUP_FACT_APP_BASED_MEETING").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DUP_FACT_APP_BASED_PM", .Range("DUP_FACT_APP_BASED_PM").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DUP_FACT_APP_BASED_REP", .Range("DUP_FACT_APP_BASED_REP").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DUP_FACT_APP_BASED_SITE", .Range("DUP_FACT_APP_BASED_SITE").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DUP_FACT_APP_BASED_SPEC", .Range("DUP_FACT_APP_BASED_SPEC").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DUP_FACT_APP_BASED_SYS_ENG", .Range("DUP_FACT_APP_BASED_SYS_ENG").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DUP_FACT_APP_BASED_SYSENG", .Range("DUP_FACT_APP_BASED_SYSENG").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "DUP_FACT_APP_BASED_TEST", .Range("DUP_FACT_APP_BASED_TEST").Value)
    
 'Formula: =HMI_RPT_PCT_QTY_TUNED
    Call xmlCreateElement(rinvGECE, objDataEntry, "HMI_RPT_PCT_QTY", .Range("HMI_RPT_PCT_QTY").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "IO_OPTIMISATION_EXP_RATE", .Range("IO_OPTIMISATION_EXP_RATE").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "IO_OPTIMISATION_MAX_REDUCTION", .Range("IO_OPTIMISATION_MAX_REDUCTION").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_PCT_TOTAL", .Range("PM_PCT_TOTAL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_IA_TOTAL", .Range("PM_IA_TOTAL").Value)
    Call xmlCreateElement(rinvGECE, objDataEntry, "PM_BUYOUT_TOTAL", .Range("PM_BUYOUT_TOTAL").Value)
    
End With


CleanUp:
    On Error Resume Next
    Set objDataEntry = Nothing


Exit Sub
ErrorHandler:
    MsgBox "error in ExportApplicationBasedSheet " & Err.Description, vbCritical
    Resume CleanUp
End Sub
Public Function ReadGECEXML(ByVal vstrFileToOpen As String, ByVal vstrWorkBook As String)
On Error GoTo ErrorHandler
Dim invGECE                  As MSXML2.DOMDocument30
Dim invGECEElement           As IXMLDOMElement
Dim invGECENode              As IXMLDOMNode
'Dim newRoot                 As IXMLDOMElement
'Dim newAtt                  As IXMLDOMAttribute
'Dim newNode                 As IXMLDOMNode
'Dim newElement              As IXMLDOMElement
'Dim pi                      As IXMLDOMProcessingInstruction
Set invGECE = New MSXML2.DOMDocument30

invGECE.Load (vstrFileToOpen)
invGECE.async = False
Set invGECEElement = invGECE.documentElement
Application.Cursor = xlWait
frmComplete.Controls("txtOutput").Text = "Start Import: " & Time
Select Case mid$(1 / 2, 2, 1)
    Case "."
        mstrCPDecimalSeparator = "."
    Case ","
        mstrCPDecimalSeparator = ","
    Case " "
        mstrCPDecimalSeparator = " "
End Select


For Each invGECENode In invGECEElement.childNodes
   
   Select Case invGECENode.nodeName
        Case "DecimalSeparator"
            mstrDecimalSeparator = invGECENode.nodeTypedValue
        Case "DataEntry"
            DoEvents
            Application.Cursor = xlWait
            frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & vbCrLf & Time & "  Importing: " & gstrGECEDataEntrySheet
            Call LoadXML(Workbooks(vstrWorkBook).Worksheets(gstrGECEDataEntrySheet), invGECENode)
        Case "AssumptionsProposal"
            DoEvents
            Application.Cursor = xlWait
            frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & vbCrLf & Time & " Importing: " & gstrGECEAssumptionsProposalSheet
            Call LoadXML(Workbooks(vstrWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet), invGECENode)
        Case "ApplicationBased"
            DoEvents
            Application.Cursor = xlWait
            frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & vbCrLf & Time & "I mporting: " & gstrGECEApplicationBasedSheet
            Call LoadXML(Workbooks(vstrWorkBook).Worksheets(gstrGECEApplicationBasedSheet), invGECENode)
        Case "PriceMarkupSheet"
            DoEvents
            Application.Cursor = xlWait
            frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & vbCrLf & Time & " Importing: " & gstrGECEPriceMakeUpSheet
            Call LoadXML(Workbooks(vstrWorkBook).Worksheets(gstrGECEPriceMakeUpSheet), invGECENode)
        Case "DurationBasedSheet" '
            DoEvents
            Application.Cursor = xlWait
            frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & vbCrLf & Time & " Importing: " & gstrGECEDurationBasedSheet
            Call LoadXML(Workbooks(vstrWorkBook).Worksheets(gstrGECEDurationBasedSheet), invGECENode)
        Case "TaskBased"
            DoEvents
            Application.Cursor = xlWait
            frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & vbCrLf & Time & " Importing: " & gstrGECETaskBasedSheet
            Call LoadTaskBased(Workbooks(vstrWorkBook).Worksheets(gstrGECETaskBasedSheet), invGECENode)
        Case Else
        
        
    
   End Select
   
   
Next

'Update task based
DoEvents
Application.Cursor = xlWait
frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & vbCrLf & Time & " Updating Task Based sheet"
Call TaskBasedSummaryUpdate

CleanUp:
    frmComplete.Controls("txtOutput").Text = frmComplete.Controls("txtOutput").Text & vbCrLf & "Finished Import: " & Time
    On Error Resume Next
    Set invGECE = Nothing
    Application.Cursor = xlDefault

Exit Function
ErrorHandler:
    MsgBox "error in ReadGECEXML " & Err.Description, vbCritical
    
    Resume CleanUp
    Resume 0
End Function
Public Sub LoadTaskBased(wsTarget As Worksheet, ByVal vinvGECENode As IXMLDOMNode)
On Error GoTo ErrorHandler
Dim invGECEElement           As IXMLDOMElement
Dim invGECENode              As IXMLDOMNode
Dim strRange                 As String
Dim strTemp                  As String
Dim rngTest                  As Range

With wsTarget

    For Each invGECENode In vinvGECENode.childNodes
        strRange = Replace(invGECENode.nodeName, "_", "$")
1       Set rngTest = .Range(strRange)
        If Err = 0 Then
        End If
        If Len(invGECENode.nodeTypedValue) > 0 Then
            If mstrDecimalSeparator = mstrCPDecimalSeparator Then
                If (.Range(strRange).Value & "") <> invGECENode.nodeTypedValue Then
2                .Range(strRange).Value = invGECENode.nodeTypedValue
            End If
            ElseIf InStr(1, .Range(strRange).NumberFormat, "0") Then
                    Select Case mstrDecimalSeparator
                        Case "comma"
                            strTemp = Replace(invGECENode.nodeTypedValue, ",", "*")
                            strTemp = Replace(strTemp, ".", "")
                            strTemp = Replace(strTemp, " ", "")
                            strTemp = Replace(strTemp, "*", mstrCPDecimalSeparator)
                        Case "period"
                            strTemp = Replace(invGECENode.nodeTypedValue, ".", "*")
                            strTemp = Replace(strTemp, ".", "")
                            strTemp = Replace(strTemp, " ", "")
                            strTemp = Replace(strTemp, "*", mstrCPDecimalSeparator)
                        Case "blank"
                            strTemp = Replace(invGECENode.nodeTypedValue, " ", "*")
                            strTemp = Replace(strTemp, ".", "")
                            strTemp = Replace(strTemp, " ", "")
                            strTemp = Replace(strTemp, "*", mstrCPDecimalSeparator)
                    End Select
                    If (.Range(strRange).Value & "") <> strTemp Then
3                       .Range(strRange).Value = strTemp
                    End If
            ElseIf (.Range(strRange).Value & "") <> invGECENode.nodeTypedValue Then
4               .Range(strRange).Value = invGECENode.nodeTypedValue
            End If
            
        End If
5
    Next
    
    
End With

CleanUp:
    On Error Resume Next
   
Exit Sub
ErrorHandler:
    If Erl = 1 Then
        Debug.Print "Named Range does not exist " & invGECENode.nodeName
        Err.Clear
        Resume 5
        Resume 0
    End If
    If Erl = 2 And wsTarget.Range(invGECENode.nodeName).Locked = True Then
        Debug.Print invGECENode.nodeName
        Resume 5
        Resume 0
    End If
    If Erl = 3 And wsTarget.Range(invGECENode.nodeName).Locked = True Then
        Debug.Print invGECENode.nodeName
        Resume 5
    End If
       If Erl = 4 And wsTarget.Range(invGECENode.nodeName).Locked = True Then
        Debug.Print invGECENode.nodeName
        Resume 5
    End If
    MsgBox "error in LoadTaskBased " & Err.Description, vbCritical
    Resume CleanUp
    Resume 0
End Sub
Public Sub LoadXML(wsTarget As Worksheet, ByVal vinvGECENode As IXMLDOMNode)
On Error GoTo ErrorHandler
Dim rngTest                  As Range
Dim invGECENode              As IXMLDOMNode
Dim strTemp                  As String

With wsTarget

    For Each invGECENode In vinvGECENode.childNodes
1       Set rngTest = .Range(invGECENode.nodeName)
        If Err = 0 Then
        End If

        If Len(invGECENode.nodeTypedValue) > 0 Then
            If mstrDecimalSeparator = mstrCPDecimalSeparator Then
                If (.Range(invGECENode.nodeName).Value & "") <> invGECENode.nodeTypedValue Then
2               .Range(invGECENode.nodeName).Value = invGECENode.nodeTypedValue
                End If
            ElseIf InStr(1, .Range(invGECENode.nodeName).NumberFormat, "0") Then
                    Select Case mstrDecimalSeparator
                        Case "comma"
                            strTemp = Replace(invGECENode.nodeTypedValue, ",", "*")
                            strTemp = Replace(strTemp, ".", "")
                            strTemp = Replace(strTemp, " ", "")
                            strTemp = Replace(strTemp, "*", mstrCPDecimalSeparator)
                        Case "period"
                            strTemp = Replace(invGECENode.nodeTypedValue, ".", "*")
                            strTemp = Replace(strTemp, ".", "")
                            strTemp = Replace(strTemp, " ", "")
                            strTemp = Replace(strTemp, "*", mstrCPDecimalSeparator)
                        Case "blank"
                            strTemp = Replace(invGECENode.nodeTypedValue, " ", "*")
                            strTemp = Replace(strTemp, ".", "")
                            strTemp = Replace(strTemp, " ", "")
                            strTemp = Replace(strTemp, "*", mstrCPDecimalSeparator)
                    End Select
                    If (.Range(invGECENode.nodeName).Value & "") <> strTemp Then
3                       .Range(invGECENode.nodeName).Value = strTemp
                    End If
            ElseIf (.Range(invGECENode.nodeName).Value & "") <> invGECENode.nodeTypedValue Then
4               .Range(invGECENode.nodeName).Value = invGECENode.nodeTypedValue
            End If

        End If
5
    Next
    
    
End With
CleanUp:
    On Error Resume Next
   
Exit Sub
ErrorHandler:
    If Erl = 1 Then
        Debug.Print "Named Range does not exist " & invGECENode.nodeName
        Err.Clear
        Resume 5
        Resume 0
    End If
   If Erl = 2 And wsTarget.Range(invGECENode.nodeName).Locked = True Then
        Debug.Print invGECENode.nodeName
        Resume 5
        Resume 0
    End If
    If Erl = 3 And wsTarget.Range(invGECENode.nodeName).Locked = True Then
        Debug.Print invGECENode.nodeName
        Resume 5
    End If
       If Erl = 4 And wsTarget.Range(invGECENode.nodeName).Locked = True Then
        Debug.Print invGECENode.nodeName
        Resume 5
    End If
    MsgBox "error in LoadXML " & invGECENode.nodeName & " " & Err.Description, vbCritical
    
    Resume CleanUp
    Resume 0
End Sub
Public Sub CreateMSXMLFile(ByVal vstrWorkBook As String, ByVal vstrFileName As String)
    On Error GoTo ErrorHandler

    Dim invGECE As Object   ' late bind for DoneEx safety
    Dim objRoot As Object
    Dim xmlFileName As String

    Set invGECE = CreateObject("MSXML2.DOMDocument.3.0")  ' robust ProgID
    invGECE.async = False

    ' decimal separator (replace the old Mid$ hack)
    Select Case Application.International(xlDecimalSeparator)
        Case ".": mstrCPDecimalSeparator = "period"
        Case ",": mstrCPDecimalSeparator = "comma"
        Case " ": mstrCPDecimalSeparator = "blank"
        Case Else: mstrCPDecimalSeparator = "period"
    End Select

    Set objRoot = xmlCreateRoot(invGECE, "GECEData", True)
    Call xmlCreateElement(invGECE, objRoot, "DecimalSeparator", mstrCPDecimalSeparator)

    Call ExportDataEntrySheet(Workbooks(vstrWorkBook).Worksheets(gstrGECEDataEntrySheet), invGECE, objRoot)
    Call ExportAssumptionsProposalSheet(Workbooks(vstrWorkBook).Worksheets(gstrGECEAssumptionsProposalSheet), invGECE, objRoot)
    Call ExportApplicationBasedSheet(Workbooks(vstrWorkBook).Worksheets(gstrGECEApplicationBasedSheet), invGECE, objRoot)
    Call ExportDurationBasedSheet(Workbooks(vstrWorkBook).Worksheets(gstrGECEDurationBasedSheet), invGECE, objRoot)
    Call ExportPriceMarkupSheet(Workbooks(vstrWorkBook).Worksheets(gstrGECEPriceMakeUpSheet), invGECE, objRoot)
    Call ExportTaskBasedSheet(Workbooks(vstrWorkBook).Worksheets(gstrGECETaskBasedSheet), invGECE, objRoot)

CleanUp:
    On Error Resume Next
    xmlFileName = vstrFileName
    If LCase$(Right$(xmlFileName, 4)) <> ".xml" Then xmlFileName = xmlFileName & ".xml"
    invGECE.Save xmlFileName
    Set objRoot = Nothing
    Set invGECE = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error in CreateMSXML: " & Err.Description, vbCritical
    Resume CleanUp
End Sub

Public Function xmlCreateRoot(ByRef robjXML As MSXML2.DOMDocument30, ByVal vstrRootName As String, Optional blnMainRoot As Boolean) As Object

On Error GoTo ErrorHandler

Dim strError                As String
Dim strSource               As String
Dim lngErr                  As Long
Dim objElement              As MSXML2.IXMLDOMElement
Dim objAttrib               As MSXML2.IXMLDOMAttribute
Dim pi                      As MSXML2.IXMLDOMProcessingInstruction
Dim newRoot                 As MSXML2.IXMLDOMElement
    
    If IsMissing(blnMainRoot) Then
        blnMainRoot = False
    End If

'Set xmlDoc = New MSXML.DOMDocument
Set newRoot = robjXML.createElement(vstrRootName)
Set robjXML.documentElement = newRoot

'processing instruction
    Set pi = robjXML.createProcessingInstruction("xml", "version=""1.0""")
    robjXML.insertBefore pi, robjXML.childNodes.Item(0)
    
    Set objAttrib = robjXML.createAttribute("SchemaVersion")
    objAttrib.Value = "1.0"
    
    'newRoot.setAttributeNode objAttrib
    
    Set xmlCreateRoot = newRoot

CleanUp:
On Error Resume Next
    Set newRoot = Nothing

  If lngErr <> 0 Then Err.Raise lngErr, "xmlMod" & ".xmlCreateRoot|" & strSource, strError
Exit Function

ErrorHandler:
    lngErr = Err
    strError = Error
    strSource = Err.Source
    Resume CleanUp
    Resume 0
End Function

Public Function xmlCreateElement(ByRef robjXML As MSXML2.DOMDocument30, ByRef objRoot As Object, ByVal vstrElementName As String, ByVal vData As Variant, Optional ByVal vstrAttributeName As String, Optional ByVal vstrAttributeValue As String) As Object
On Error GoTo ErrorHandler

Dim strError                As String
Dim strSource               As String
Dim lngErr                  As Long
Dim lngType                 As Long
Dim lngLength               As Long
Dim blnSkipAttribs          As Boolean
Dim objXMLElem              As Object
Dim objXMLAttrib            As Object
Dim strTemp                 As String

    Set objXMLElem = robjXML.createElement(vstrElementName)
    
    'set the text
    Select Case VarType(vData)
'        Case vbEmpty:
'        Case vbNull:
'        Case vbInteger:     If vData = 0 Then Nullify = Null
'        Case vbLong:        If vData = 0 Then Nullify = Null
'        Case vbSingle:      If vData = 0 Then Nullify = Null
'        Case vbDouble:      If vData = 0 Then Nullify = Null
'        Case vbCurrency:    If vData = 0 Then Nullify = Null
        Case vbDate:
            strTemp = vData
            objXMLElem.Text = strTemp
            'Debug.Print vstrElementName & " " & vData
        'Case vbString:
        Case vbBoolean:
            strTemp = vData
            objXMLElem.Text = strTemp
            'Debug.Print vstrElementName & " " & vData
        Case Else
            objXMLElem.Text = vData
    End Select
    'objXMLElem.Text = vData
    
    objRoot.appendChild objXMLElem

CleanUp:
On Error Resume Next
    Set objXMLElem = Nothing
If lngErr <> 0 Then Err.Raise lngErr, "xmlMod" & ".xmlCreateElement|" & strSource, strError
Exit Function

ErrorHandler:
    lngErr = Err
    strError = Error
    strSource = Err.Source
    Resume CleanUp
End Function

Public Function xmlCreateRootNode(ByRef robjXML As MSXML2.DOMDocument30, ByVal vstrRootName As String, Optional blnMainRoot As Boolean) As Object
On Error GoTo ErrorHandler
Dim strError                As String
Dim strSource               As String
Dim lngErr                  As Long

Dim newRoot                 As IXMLDOMElement
Dim newAtt                  As IXMLDOMAttribute
Dim newNode                 As IXMLDOMNode
Dim newElement              As IXMLDOMElement
Dim pi                      As IXMLDOMProcessingInstruction

Set newRoot = robjXML.createElement("root")
Set robjXML.documentElement = newRoot

'processing instruction
Set pi = robjXML.createProcessingInstruction("xml", "version=""1.0""")
robjXML.insertBefore pi, robjXML.childNodes.Item(0)

'attribute
Set newAtt = robjXML.createAttribute("SchemaVersion")
newAtt.Value = "1.0"

newRoot.setAttributeNode newAtt

'element by node
Set newNode = robjXML.createNode(NODE_ELEMENT, "OrderHeader", "")
Set newElement = newNode
newElement.Text = "XML here..."
newRoot.appendChild newNode

'element by element
Set newElement = robjXML.createElement("OrderLineItem")
newElement.Text = "Line Items XML here..."
newNode.appendChild newElement

Set xmlCreateRootNode = robjXML

CleanUp:
On Error Resume Next
    Set newRoot = Nothing

  If lngErr <> 0 Then Err.Raise lngErr, "xmlMod" & ".xmlCreateRoot|" & strSource, strError
Exit Function

ErrorHandler:
    lngErr = Err
    strError = Error
    strSource = Err.Source
    Resume CleanUp
    Resume 0
End Function

Public Function xmlCreateSubRoot(ByRef robjXML As MSXML2.DOMDocument30, ByRef robjRoot As Object, ByVal vstrRootName As String) As Object
On Error GoTo ErrorHandler

Dim strError                As String
Dim strSource               As String
Dim lngErr                  As Long
Dim objXMLElem              As Object

    Set objXMLElem = robjXML.createElement(vstrRootName)
    Set xmlCreateSubRoot = robjRoot.appendChild(objXMLElem)


CleanUp:
On Error Resume Next
    Set objXMLElem = Nothing

If lngErr <> 0 Then Err.Raise lngErr, "xmlMod" & ".xmlCreateSubRoot|" & strSource, strError
Exit Function

ErrorHandler:
    lngErr = Err
    strError = Error
    strSource = Err.Source
    Resume CleanUp
End Function




