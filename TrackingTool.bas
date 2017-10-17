Attribute VB_Name = "TrackingTool"
'/**
'* This is the VBA module for the DCSO Incident Tracking Tool.
'* Import this module to the Excel file using the Visual Basic Editor's import function.
'* Updated: August 09, 2017
'* Author: IRT - Incident Response Team DCSO - Daniel Nguyen
'*
'*/


'#############################################
' PUBLIC VARIABLE LISTING
'#############################################

Dim businessAreas() As String
Dim businessAreasID() As String
Dim businessAreasScope() As Boolean
Dim businessAreasGroup() As Integer

Dim workPackages() As String
Dim workPackagesID() As String
Dim workPackagesShortID() As String
Dim workPackagesLink() As String
Dim workPackagesActionsID() As Integer
Dim workPackagesActions() As String
Dim workPackages_i_Actions() As String
 
Dim actions() As String
Dim actionsID() As String
Dim actionsColor() As String
Dim actionsSeeIfActive() As String
Dim actionsSetBy() As String

Dim wsTracking As Worksheet
Dim wsIDs As Worksheet
Dim wsDashboard As Worksheet
Dim wsStartPackages As Worksheet
Dim wsOutlookImport As Worksheet
Dim wsIssues As Worksheet
Dim wsConfiguration As Worksheet

Dim ReceivedTime As String

Dim number_of_workpackages As Integer
Dim number_of_states As Integer
Dim number_of_businessareas As Integer

Dim data_already_read As Boolean
Dim automaticImport As Boolean
Dim Timestring As String


'#############################################
' SET PUBLIC VARIABLES
'#############################################
'/**
'* Set worksheets.
'*/
Sub SetWorksheets()
    
    Set wsTracking = ThisWorkbook.Sheets("Tracking")
    Set wsIDs = ThisWorkbook.Sheets("IDs")
    Set wsDashboard = ThisWorkbook.Sheets("Dashboard")
    Set wsStartPackages = ThisWorkbook.Sheets("Start packages")
    Set wsOutlookImport = ThisWorkbook.Sheets("OutlookImport")
    Set wsIssues = ThisWorkbook.Sheets("Issues")
    Set wsConfiguration = ThisWorkbook.Sheets("Configuration")

End Sub

'/**
'* Set array sizes of public variables.
'*/
Sub SetArraySize()
    
    number_of_workpackages = wsIDs.Range("WorkpackageTable").Rows.Count
    number_of_states = wsIDs.Range("StateTable").Rows.Count
    number_of_businessareas = wsIDs.Range("BusinessAreaTable").Rows.Count
    
    ReDim businessAreas(number_of_businessareas)
    ReDim businessAreasID(number_of_businessareas)
    ReDim businessAreasScope(number_of_businessareas)
    ReDim businessAreasGroup(number_of_businessareas)
    
    ReDim workPackages(number_of_workpackages)
    ReDim workPackagesID(number_of_workpackages)
    ReDim workPackagesLink(number_of_workpackages)
    ReDim workPackagesShortID(number_of_workpackages) 'OPTIONAL Set a short Value for WPs - Got to rearrange ranges of the matrix
    ReDim workPackages_i_Actions(number_of_workpackages) 'Matrix
    
    ReDim actions(number_of_states)
    ReDim actionsID(number_of_states)
    ReDim actionsColor(number_of_states)
    ReDim actionsSeeIfActive(number_of_states)
    ReDim actionsSetBy(number_of_states)
        
    ReDim workPackagesActionsID(number_of_workpackages, number_of_states) As Integer
    ReDim workPackagesActions(number_of_workpackages, number_of_states) As String
    
End Sub

'#############################################
' READ IN ALL DATA
'#############################################
'/**
'* Read in data from the "IDs" sheet and store in public variables.
'*/

Sub Read_data()
    data_already_read = True
    SetWorksheets
    SetArraySize

    Dim cell As Range
    Dim i_of_BusinessAreas As Integer
    Dim i_of_Workpackages As Integer
    Dim i_of_Actions As Integer
    Dim rng_BusinessAreas As Range
    Dim rng_WorkPackages As Range
    Dim rng_Actions As Range
    
    Set rng_BusinessAreas = wsIDs.Range("BusinessAreaTable[Business Area]")
    Set rng_WorkPackages = wsIDs.Range("WorkpackageTable[Workpackages]")
    Set rng_Actions = wsIDs.Range("StateTable[State]")

    '//Create Business Area Array and set scope
    i_of_BusinessAreas = 0
    For Each cell In rng_BusinessAreas
        
        i_of_BusinessAreas = i_of_BusinessAreas + 1                     'Counting all cells in range
        businessAreas(i_of_BusinessAreas) = cell.Text                   'Set the BA name into the array
        businessAreasID(i_of_BusinessAreas) = cell.offset(0, 1).Text    'Set the ID into the array
        
        If cell.offset(0, 2).Text = "Yes" Then businessAreasScope(i_of_BusinessAreas) = True    'Setting the Scope with boolean
        If cell.offset(0, 2).Text = "No" Then businessAreasScope(i_of_BusinessAreas) = False
        '//OPTIONAL Setting of a group
        If cell.offset(0, 3).Value <> "" Then
            businessAreasGroup(i_of_BusinessAreas) = cell.offset(0, 3).Value
        Else
            businessAreasGroup(i_of_BusinessAreas) = 0
        End If
    Next cell

    
    '//Create Work Packages Array
    i_of_Workpackages = 0
    For Each cell In rng_WorkPackages
        i_of_Workpackages = i_of_Workpackages + 1
        If cell.offset(0, number_of_states + 2).Value <> "" Then
            workPackagesShortID(i_of_Workpackages) = cell.offset(0, number_of_states + 2).Text
        End If
        'OPTIONAL 3rd Value
'        ReDim Preserve KHeaders(i_of_Workpackages)
'        KHeaders(i_of_Workpackages) = cell.Offset(0, 3).Text
        workPackages(i_of_Workpackages) = cell.Text                 'Set the WP name into the array
        workPackagesID(i_of_Workpackages) = cell.offset(0, 1).Text  'Set the IDS
        ' Set the Links if exist
        If cell.offset(0, number_of_states + 3).Value <> "" Then
            workPackagesLink(i_of_Workpackages) = cell.offset(0, number_of_states + 3).Text
        End If
        'Fill the matrix with actions
        For j = 1 To number_of_states
            workPackagesActions(i_of_Workpackages, j) = cell.offset(0, j + 1).Text
        Next j
        workPackages_i_Actions(i_of_Workpackages) = number_of_states    'Set Array Elements to number of Actions found (or columns found in matrix)
    Next cell
    
    '//Create Actions Arrays
    i_of_Actions = 0
    For Each cell In rng_Actions
        i_of_Actions = i_of_Actions + 1
        'Set the Action name into the array
        actions(i_of_Actions) = cell.Text
        actionsID(i_of_Actions) = cell.offset(0, 1).Text
        actionsColor(i_of_Actions) = cell.offset(0, 2).Text
        actionsSetBy(i_of_Actions) = cell.offset(0, 3).Text
        actionsSeeIfActive(i_of_Actions) = cell.offset(0, 4).Text
    Next cell


    For j = 1 To number_of_workpackages
        For k = 1 To workPackages_i_Actions(j)
            For i = 1 To number_of_states
                If actions(i) = workPackagesActions(j, k) Then
                    workPackagesActionsID(j, k) = actionsID(i)
                    Exit For
                End If
            Next i
        Next k
    Next j

        
    With Application
        .EnableEvents = True
        .ScreenUpdating = False
    End With
    
End Sub

'#############################################
' RESET BUTTONS
'#############################################
'/**
'* Spam a message box to accpet or decline a reset of all data. If accepted, delete all data and re-create the "Tracking" sheet.
'*/
Sub Button_Tool_Reset()

    Dim answer As Integer
    
    answer = MsgBox("Are you sure you want to reset All Data?", vbYesNo + vbQuestion, "Reset Tracking Tool and Dashboard")
    If answer = vbYes Then
        
        Read_data
        
        Sheets("Tracking").Activate
        Create_Tracking_Sheet
        
        Sheets("Dashboard").Activate
        Sheets("Dashboard").Cells(1, 10).Value = "<Your next call date>"
        
        Sheets("Configuration").Activate
        answer = MsgBox("Reset configurations?", vbYesNo + vbQuestion, "Reset Configuration's Sheet")
        If answer = vbYes Then
            Sheets("Configuration").Range("E4:H60").ClearContents
            Sheets("Configuration").Range("C16:C60").ClearContents
            
            Sheets("Configuration").Cells(4, 5).Value = "<Your Email>"
            Sheets("Configuration").Cells(5, 5).Value = "Inbox"
            Sheets("Configuration").Cells(5, 6).Value = "Tracking"
            Sheets("Configuration").Cells(6, 5).Value = "[<Identifier>]"
            Sheets("Configuration").Cells(9, 5).Value = "<Your Email>"
            Sheets("Configuration").Cells(16, 5).Value = "<Email>"
            
            For i = 1 To number_of_businessareas
                j = i * 2
                Sheets("Configuration").Cells(14 + j, 3).Value = "=INDEX(BusinessAreaTable[Business Area]," & i & ")"
                'Sheets("Configuration").Cells(16, 5).Value = "INDEX(BusinessAreaTable[Business Area],i)"
                'Sheets("Configuration").Cells(16, 7).Value = "INDEX(BusinessAreaTable[Business Area],i)"
                'Sheets("Configuration").Cells(16, 3).Value = [BusinessAreaTable].Cells(1, 1)
            Next i
        End If
        
        Sheets("OutlookImport").Activate
        answer = MsgBox("Reset Outlook imported items list?", vbYesNo + vbQuestion, "Reset 'OutlookImport' Sheet")
        If answer = vbYes Then
            Sheets("OutlookImport").Rows("2:" & Rows.Count).ClearContents
        End If
        
        Sheets("Issues").Activate
        answer = MsgBox("Reset Outlook issues' list?", vbYesNo + vbQuestion, "Reset 'Issues' Sheet")
        If answer = vbYes Then
            Sheets("Issues").Range("A5:H" & Rows.Count).ClearContents
            Sheets("Issues").Range("B2:B3").ClearContents
            Sheets("Issues").Cells(2, 2).Value = "0"
        End If
        
        Sheets("Changelog").Activate
        answer = MsgBox("Reset changelog?", vbYesNo + vbQuestion, "Reset 'ChangeLog' Sheet")
        If answer = vbYes Then
            Sheets("ChangeLog").Rows("2:" & Rows.Count).ClearContents
        End If
        
        Sheets("Deadlines").Activate
        answer = MsgBox("Reset deadlines?", vbYesNo + vbQuestion, "Reset 'Deadlines' Sheet")
        If answer = vbYes Then
            Sheets("Deadlines").Range("B3:B" & Rows.Count).ClearContents
            Sheets("Deadlines").Range("D3:D" & Rows.Count).ClearContents
            Sheets("Deadlines").Cells(3, 2).Value = "1/1/2017  12:00:00 AM"
            Sheets("Deadlines").Cells(4, 2).Value = "1/1/2019  12:00:00 AM"
        End If
        
        Sheets("IDs").Range("WorkpackageTableWP Short ID

        Sheets("Configuration").Activate
        MsgBox "Delete Mail body manually."
        
    Else
    'do nothing
    End If

End Sub


'/**
'* Spam a message box to accpet or decline a reset.
'*/
Sub Button_Tracking_Reset()
    
    Dim answer As Integer
    
    answer = MsgBox("Are you sure you want to reset the Tracking Tool and Dashboard?", vbYesNo + vbQuestion, "Reset Tracking Tool and Dashboard")
    Worksheets("Tracking").Activate
    
    If answer = vbYes Then
        Create_Tracking_Sheet
    Else
    'do nothing
    End If

End Sub


'#############################################
' DELETE AND RE-CREATE THE TRACKING SHEET
'#############################################
'/**
'* Delete and re-create the "Tracking" sheet. Read in data from tables in "IDs" sheet.
'*/
Sub Create_Tracking_Sheet()
    With Application
        .EnableEvents = True
        .ScreenUpdating = False
    End With
    
    If Not data_already_read Then
        MsgBox "Read Data in Create Tracking"
        Read_data
    End If
    
    Dim line As Integer
    Dim i, j, k As Integer
  
    wsTracking.Cells.Clear  'Clear Everything
    
    'Now create header row
    wsTracking.Cells(1, 1).Value = "Location"
    wsTracking.Cells(1, 2).Value = "Work Package"
    wsTracking.Cells(1, 3).Value = "Action"
    wsTracking.Cells(1, 4).Value = "ID"
    wsTracking.Cells(1, 5).Value = "Status"
    wsTracking.Cells(1, 6).Value = "WP ShortID"
    wsTracking.Cells(1, 7).Value = "Color"
    'wsTracking.Cells(1, 8).Value = "R"
    'wsTracking.Cells(1, 9).Value = "G"
    'wsTracking.Cells(1, 10).Value = "B"
    wsTracking.Cells(1, 11).Value = "Date"
    wsTracking.Cells(1, 12).Value = "Manual/Reported"
    
    For i = 1 To 12
        wsTracking.Cells(1, i).Font.Bold = True
    Next i
    
    
    line = 2    'start after header line
    For i = 1 To number_of_businessareas
        For j = 1 To number_of_workpackages
            For k = 1 To workPackages_i_Actions(j)
                wsTracking.Cells(line, 1).Value = businessAreas(i)
                wsTracking.Cells(line, 2).Value = workPackages(j)
                wsTracking.Cells(line, 3).Value = actions(workPackagesActionsID(j, k))
                wsTracking.Cells(line, 4).NumberFormat = "@"
                'Concatenate Sub ID s
                wsTracking.Cells(line, 4).Value = businessAreasID(i) & workPackagesID(j) & actionsID(workPackagesActionsID(j, k))
                'Always set first item to Active
                If k = 1 Then
                    wsTracking.Cells(line, 5).Value = "Active"
                Else
                    wsTracking.Cells(line, 5).Value = "Inactive"
                End If
                
                wsTracking.Cells(line, 6).Value = workPackagesShortID(j)
                'Set Colors
                wsTracking.Cells(line, 7).Value = actionsColor(workPackagesActionsID(j, k))
                wsTracking.Cells(line, 8).FormulaR1C1 = "=VLOOKUP(R[0]C[-1],ColorTable,2,FALSE)"
                wsTracking.Cells(line, 9).FormulaR1C1 = "=VLOOKUP(R[0]C[-2],ColorTable,3,FALSE)"
                wsTracking.Cells(line, 10).FormulaR1C1 = "=VLOOKUP(R[0]C[-3],ColorTable,4,FALSE)"
                wsTracking.Cells(line, 8).Font.Color = RGB(255, 255, 255)
                wsTracking.Cells(line, 9).Font.Color = RGB(255, 255, 255)
                wsTracking.Cells(line, 10).Font.Color = RGB(255, 255, 255)
                line = line + 1
            Next k
        Next j
    Next i

    
    '//COLORS
    Columns("E:E").Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="Inactive", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:="Active", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    
    Selection.FormatConditions(1).StopIfTrue = False
    Cells.FormatConditions.Delete
    Columns("E:E").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""Inactive"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Columns("E:E").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""Active"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    'END COLORS
    
    '//Borders
    wsTracking.Range("A1:G" & Cells(1, 1).End(xlDown).Row).BorderAround xlDouble
    wsTracking.Range("A1:L" & Cells(1, 1).End(xlDown).Row).BorderAround xlDouble
    wsTracking.Range("A1:L" & Cells(1, 1).End(xlDown).Row).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    wsTracking.Range("A1:G" & Cells(1, 1).End(xlDown).Row).Borders(xlInsideVertical).LineStyle = xlContinuous
    wsTracking.Range("K1:L" & Cells(1, 1).End(xlDown).Row).Borders(xlInsideVertical).LineStyle = xlContinuous

    
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    
    Application.Wait (Now + TimeValue("00:00:02"))
    wsDashboard.Activate
    Update_Dashboard

End Sub



'#############################################
' UPDATE THE DASHBOARD
'#############################################
'/**
'* Read in data from "Tracking" sheet and update the dashboard matrix.
'*/
Sub Update_Dashboard()

    Dim i, j, k As Integer
    Dim R, G, B As Integer
    Dim cell As Range
    Dim code As String
    Dim rngMeasures As Range
    Dim rngTracking As Range
    Dim wsDashboard As Worksheet
    Dim wsTracking As Worksheet
    Dim wsIDs As Worksheet
    
    Set wsDashboard = ThisWorkbook.Sheets("Dashboard")
    Set wsTracking = ThisWorkbook.Sheets("Tracking")
    Set wsIDs = ThisWorkbook.Sheets("IDs")
    Set rngMeasures = wsDashboard.Range("DashboardMatrix")
    
    number_of_states = wsIDs.Range("StateTable").Rows.Count
    
    For j = 2 To rngMeasures.Columns.Count
        For i = 1 To rngMeasures.Rows.Count
            Set cell = wsDashboard.Range("DashboardMatrix").Cells(i, j)
            code = rngMeasures.Cells(i, j).Value & "01"
            Set rngTracking = wsTracking.Range("D:D").Find(code, LookIn:=xlValues) 'Look for start Position
            '//Fill cells with color
            For k = 0 To number_of_states - 1
                If wsTracking.Cells(rngTracking.Row + k, 5) = "Active" Then
                    R = wsTracking.Cells(rngTracking.Row + k, 8)
                    G = wsTracking.Cells(rngTracking.Row + k, 9)
                    B = wsTracking.Cells(rngTracking.Row + k, 10)
                    
                    cell.Interior.Color = RGB(R, G, B)
                    cell.Font.Color = RGB(R, G, B)
                    Exit For
                    'Stop the for loop
                End If
            Next k
        Next i
    Next j
        
    wsDashboard.Cells(2, 2) = "Last Update: " & Now
            
    ThisWorkbook.RefreshAll
    
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With
        
End Sub


'#############################################
' CHANGE STATUS FOR ALL LOCATION
'#############################################
'/**
'* Change status of a package for all location and update "Tracking" sheet and dashboard.
'*/
Sub Change_in_all()
    
    If Not data_already_read Then
        MsgBox "Read Data in Change_in_all"
        Read_data
    End If
    
    Dim domain As String
    Dim package As String
    Dim action As String
    Dim rngChangeInAll As Range
    Dim rngTracking As Range
    Set rngChangeInAll = ActiveWorkbook.Names("ChangeinAll").RefersToRange
    
    
    package = rngChangeInAll.offset(0, 1).Value
    action = rngChangeInAll.offset(0, 2).Value
    
    For i = 1 To number_of_businessareas        'Loop through all location
        domain = businessAreas(i)
        Set rngTracking = ChangeStatusWithID(businessAreasID(Domaintoi(domain)), workPackagesID(Packagetoi(package)), actionsID(Actiontoi(action)))
        LogTime rngTracking, 0
    Next i
        
    Change_log 1, rngChangeInAll, 0

    wsDashboard.Activate
    Update_Dashboard
    Application.Wait (Now + TimeValue("00:00:02"))
    wsStartPackages.Activate
    
    With Application
        .EnableEvents = True
        .ScreenUpdating = False
    End With
    
End Sub


'#############################################
' CHANGE STATUS FOR ONE LOCATION
'#############################################
'/**
'* Change status of a package for one location and update "Tracking" sheet and dashboard.
'*/
Sub Change_in_one()

    If Not data_already_read Then
        MsgBox "Read Data in Change_in_one"
        Read_data
    End If
    
    Dim rngChangeInOne As Range
    Dim domain As String
    Dim package As String
    Dim action As String
    Dim rngTracking As Range
    
    Set rngChangeInOne = ActiveWorkbook.Names("ChangeInOne").RefersToRange
    domain = rngChangeInOne.offset(0, 1).Value
    package = rngChangeInOne.offset(0, 2).Value
    action = rngChangeInOne.offset(0, 3).Value
    
    'Calls function
    Set rngTracking = ChangeStatusWithID(businessAreasID(Domaintoi(domain)), workPackagesID(Packagetoi(package)), actionsID(Actiontoi(action)))
    
    LogTime rngTracking, 0
    Change_log 0, rngChangeInOne, 0

    wsDashboard.Activate
    Update_Dashboard
    Application.Wait (Now + TimeValue("00:00:02"))
    wsStartPackages.Activate
    
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    
End Sub

'#############################################
' WRITE TIMESTAMP TO TRACKING SHEET
'#############################################
'/**
'* Write timestamp and "Reported" or "Manual" to "Tracking" sheet.
'* @param: all As Boolean, rngChange As Range, imported As Boolean
'*/
Sub LogTime(rngTracking As Range, imported As Boolean)
    Dim wsTracking As Worksheet
    Set wsTracking = ThisWorkbook.Sheets("Tracking")
    
    If imported Then
        wsTracking.Cells(rngTracking.Row, 12).Value = "Reported"
        wsTracking.Cells(rngTracking.Row, 11).Value = ReceivedTime
        Change_log 0, rngTracking, 1
    Else
        wsTracking.Cells(rngTracking.Row, 12).Value = "Manual"
        wsTracking.Cells(rngTracking.Row, 11).Value = Now
    End If

End Sub


'#############################################
' WRITE DATA TO CHANGELOG
'#############################################
'/**
'* Write changes to the "Changelog" sheet
'* @param: all As Boolean, rngChange As Range, imported As Boolean
'*/
Sub Change_log(all As Boolean, rngChange As Range, imported As Boolean)
    Dim wsChangeLog As Worksheet
    Dim actualLine As Integer
    Dim wsTracking As Worksheet
    Set wsTracking = ThisWorkbook.Sheets("Tracking")
    
    Set wsChangeLog = ThisWorkbook.Sheets("ChangeLog")
    actualLine = 2
    
    While wsChangeLog.Cells(actualLine, 1).Text <> ""
        actualLine = actualLine + 1
    Wend
    
    If imported Then
        wsChangeLog.Cells(actualLine, 1).Value = wsTracking.Cells(rngChange.Row, 2).Value
        wsChangeLog.Cells(actualLine, 2).Value = wsTracking.Cells(rngChange.Row, 3).Value
        wsChangeLog.Cells(actualLine, 3).Value = wsTracking.Cells(rngChange.Row, 1).Value
        wsChangeLog.Cells(actualLine, 4).Value = ReceivedTime
        wsChangeLog.Cells(actualLine, 5).Value = "Reported"
    ElseIf all Then
        wsChangeLog.Cells(actualLine, 1).Value = rngChange.offset(0, 1).Value
        wsChangeLog.Cells(actualLine, 2).Value = rngChange.offset(0, 2).Value
        wsChangeLog.Cells(actualLine, 3).Value = "All locations"
        wsChangeLog.Cells(actualLine, 4).Value = Now
        wsChangeLog.Cells(actualLine, 5).Value = "Manual"
    Else
        wsChangeLog.Cells(actualLine, 1).Value = rngChange.offset(0, 2).Value
        wsChangeLog.Cells(actualLine, 2).Value = rngChange.offset(0, 3).Value
        wsChangeLog.Cells(actualLine, 3).Value = rngChange.offset(0, 1).Value
        wsChangeLog.Cells(actualLine, 4).Value = Now
        wsChangeLog.Cells(actualLine, 5).Value = "Manual"
    End If
    
    
End Sub

'#############################################
' GET TRACKING FOLDER IN OUTLOOK
'#############################################
'/**
'* Read in the tracking folder in Outlook
'* @return: tracking folder as object
'*/
Function GetTrackingFolder() As Object
    
    Dim wsConfiguration As Worksheet
    Dim outlookObj As Object
    Dim outlookNameSpace As Object
    Dim trackingFolder As Object
    Dim j As Integer
    
    Set wsConfiguration = ThisWorkbook.Sheets("Configuration")
    Set outlookObj = CreateObject("Outlook.Application")
    Set outlookNameSpace = outlookObj.GetNamespace("MAPI")
    Set trackingFolder = outlookNameSpace.Folders(wsConfiguration.Range("E4").Value)
    j = 0
    
    While (wsConfiguration.Range("E5").offset(0, j).Value) <> ""
        Set trackingFolder = trackingFolder.Folders(wsConfiguration.Range("E5").offset(0, j).Value)
        j = j + 1
        'MsgBox trackingFolder 'DEBUG
    Wend
    
    Set GetTrackingFolder = trackingFolder
    
End Function


'#############################################
'IMPORT MAIL-OBJECTS FROM OUTLOOK
'#############################################
'/**
'* Read in the imported ID codes in "OutlookImport" and update tracking and timestamps.
'*/
Sub importFromOutlook()
    
    ImportedFlag = 1
    
    Dim wsOutlookImport As Worksheet
    Dim wsIssues As Worksheet
    Dim wsConfiguration As Worksheet

    Dim code As String
    Dim trackingFolder As Object
    Dim subjectIDStr As String
    Dim importLine As Integer
    Dim latestEmailDate As Date
    
    Set wsOutlookImport = ThisWorkbook.Sheets("OutlookImport")
    Set wsIssues = ThisWorkbook.Sheets("Issues")
    Set wsConfiguration = ThisWorkbook.Sheets("Configuration")
    Set trackingFolder = GetTrackingFolder()
    
    subjectIDStr = wsConfiguration.Range("E6").Value
    latestEmailDate = #1/1/1900#
    importLine = 2
    
    While wsOutlookImport.Cells(importLine, 1).Text <> ""          'Compare latestEmailDate
        If latestEmailDate < wsOutlookImport.Cells(importLine, 3) Then
            latestEmailDate = wsOutlookImport.Cells(importLine, 3)
        End If
        importLine = importLine + 1
    Wend

    For Each outlookMail In trackingFolder.Items     'Iterate through mails received
        'Matching Subject ID String
        If outlookMail.ReceivedTime > latestEmailDate And Left(outlookMail.subject, Len(subjectIDStr)) = subjectIDStr Then
            
            code = Mid(outlookMail.subject, Len(subjectIDStr) + 2, 6) 'Read in the 6 digit code at index +2 instead of +1 because of [ before code
        
            If Mid(outlookMail.subject, Len(subjectIDStr) + 9, 7) = "[ISSUE]" Then   'Matching for String "[ISSUE]" in Subject
                If outlookMail.ReceivedTime > wsIssues.Cells(6, 6) Then
                    ImportIssues code, outlookMail, wsIssues
                End If
            Else
                wsOutlookImport.Cells(importLine, 1).Value = outlookMail.SenderName
                wsOutlookImport.Cells(importLine, 2).Value = outlookMail.subject
                wsOutlookImport.Cells(importLine, 3).Value = outlookMail.ReceivedTime
                wsOutlookImport.Cells(importLine, 4).Value = code
            
                ReceivedTime = outlookMail.ReceivedTime 'For the time tracking
                ParseID (code)
                importLine = importLine + 1
                
                If outlookMail.UnRead = True Then
                    outlookMail.UnRead = False
                End If
            End If
        End If
    Next outlookMail
    
    ImportedFlag = 0

    
    If automaticImport Then
        Update_Dashboard
        Application.OnTime Now + TimeValue("00:01:00"), "importFromOutlook"
    Else
        Worksheets("Dashboard").Activate
        Update_Dashboard
        Application.Wait (Now + TimeValue("00:00:02"))
        Worksheets("OutlookImport").Activate
    End If
    
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With

End Sub


'/**
'* Read in the issue mail. Write data to "Issues" sheet.
'* @param: String code; outlookMail; wsIssues As Worksheet
'*/
Sub ImportIssues(code As String, outlookMail, wsIssues As Worksheet)
    
    If Not data_already_read Then
        MsgBox "Read Data to ImportIssues"
        Read_data
    End If

    Dim code_locationID As Integer
    Dim code_packageID As Integer
    
    code_locationID = CInt(Left(code, 2))     'Extract the location code
    code_packageID = Mid(code, 3, 2)          'Extract the package code
    
    '// Write to "Issues" sheet
    wsIssues.Cells(2, 2).Value = wsIssues.Cells(2, 2).Value + 1     'Number of received issues + 1
    wsIssues.Cells(3, 2).Value = outlookMail.ReceivedTime
    wsIssues.Rows("6:6").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove  'Add a new Row
    
    '// Look for the index of the package
    For j = 1 To number_of_workpackages
        If workPackagesID(j) = code_packageID Then
            Exit For
        End If
    Next j
    
    wsIssues.Cells(6, 1).Value = businessAreas(code_locationID)
    wsIssues.Cells(6, 2).Value = workPackages(j)
    wsIssues.Cells(6, 3).Value = outlookMail.SenderName
    wsIssues.Cells(6, 4).Value = outlookMail.subject
    wsIssues.Cells(6, 5).Value = outlookMail.body
    wsIssues.Cells(6, 6).Value = outlookMail.ReceivedTime
    wsIssues.Cells(6, 7).Value = code
    wsIssues.Cells(6, 8).Value = "Unsolved"

End Sub


'#############################################
' RE-APPLY IMPORTED IDS
'#############################################
'/**
'* Read in the imported ID codes in "OutlookImport" and update tracking and timestamps.
'*/
Sub ApplyOutlookImportedIDs()
    Dim i As Integer
    Dim rngOutlookImportIDs As Range
    
    Set rngOutlookImportIDs = ActiveWorkbook.Names("OutlookImportIDs").RefersToRange
    ImportedFlag = 1

    i = 0
    Do While rngOutlookImportIDs.offset(i, 0).Value <> ""   ' Look for empty line
        ReceivedTime = rngOutlookImportIDs.offset(i, -1).Value
        ParseID rngOutlookImportIDs.offset(i, 0).Value
        i = i + 1
    Loop
    
    ImportedFlag = 0
    
    Worksheets("Dashboard").Activate
    Update_Dashboard
    Application.Wait (Now + TimeValue("00:00:02"))
    Worksheets("OutlookImport").Activate
End Sub

'#############################################
' AUTOMATIC IMPORT
'#############################################
'/**
'* Start automatic import. Set automaticImport flag to True.
'*/
Sub startAutomaticImport()
    If Not IS_DATA_READ Then
        Read_data
    End If
    
    MsgBox "Start Automatic Import"
    automaticImport = True
    
    Application.OnTime Now + TimeValue("00:01:00"), "importFromOutlook"
End Sub
'/**
'* Stop automatic import. Set automaticImport flag to false.
'*/
Sub stopAutomaticImport()
    MsgBox "Stop Automatic Import"
    automaticImport = False
End Sub


'#############################################
' SEND STATUS REQUEST MAILS
'#############################################
'/**
'* Generate and send status request mail to a location/business area.
'* Replace placeholder variables in HTML code.
'* Create Outlook Object.
'*/
Sub Send_to_location()
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With
    
    If Not data_already_read Then
        MsgBox "Read Data to Send Status Mail"
        Read_data
    End If

    Dim rngMailCompany As Range
    Dim rngMailSubject As Range
    Dim rngMailHead As Range
    Dim rngMailBody As Range
    Dim rngMailStatusReply As Range
    Dim rngMailIssueReply As Range
    Dim rngMailSignature As Range
    Dim rngContacts As Range
    
    Dim MAILTO As String
    Dim MAILBCC As String
    Dim MAILONBEHALF As String
    Dim SUBJECTID As String
    Dim REPLYTO As String
    Dim DOMAINNAME As String
    Dim CURRENTDATE As String
    Dim PACKAGENAME As String
    Dim PACKAGELINK As String
    Dim PACKAGESTATUS As String
    Dim IDCODE As String
    Dim subject As String
    Dim body As String
    Dim actionsString As String
    Dim domain As String
    
    Dim OutApp As Object
    Dim OutMail As Object

    Dim iDomain As Integer
    Dim activeaction As Integer
    Dim i As Integer
    
    ReDim contactsTO(number_of_businessareas)
    ReDim contactsCC(number_of_businessareas)

        
    Set rngMailCompany = ActiveWorkbook.Names("MailCompany").RefersToRange
    Set rngMailSubject = ActiveWorkbook.Names("MailSubject").RefersToRange
    Set rngMailHead = ActiveWorkbook.Names("MailHead").RefersToRange
    Set rngMailBody = ActiveWorkbook.Names("MailBody").RefersToRange
    Set rngMailStatusReply = ActiveWorkbook.Names("MailStatusReply").RefersToRange
    Set rngMailIssueReply = ActiveWorkbook.Names("MailIssueReply").RefersToRange
    Set rngMailSignature = ActiveWorkbook.Names("MailSignature").RefersToRange
    Set rngContacts = ActiveWorkbook.Names("Contacts").RefersToRange
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    domain = rngMailCompany.Value
    
    SUBJECTID = wsConfiguration.Range("E6").Value
    MAILONBEHALF = wsConfiguration.Range("E7").Value
    MAILBCC = wsConfiguration.Range("E8").Value
    REPLYTO = wsConfiguration.Range("E9").Value
    DOMAINNAME = domain
    CURRENTDATE = Date
    
    iDomain = Domaintoi(domain)

    ' Read in Replyto mail addresses
    i = 1
    While wsConfiguration.Range("E9").offset(0, i).Value <> ""
        REPLYTO = REPLYTO + "; " + wsConfiguration.Range("E9").offset(0, i).Value
        i = i + 1
    Wend
    
    'Replace SUBJECTID placeholder
    subject = rngMailSubject.Value
    subject = Replace(subject, "SUBJECTID", SUBJECTID)
    subject = Replace(subject, "CURRENTDATE", CURRENTDATE & " at " & Time)
    subject = Replace(subject, "BUSINESSAREA", DOMAINNAME)
    'DEBUG
    'MsgBox "Email Head" & rngMailHead.offset(0, 0).Value
    
    body = rngMailHead.offset(0, 0).Value
    body = Replace(body, "CURRENTDATE", CURRENTDATE & " at " & Time)
    body = Replace(body, "BUSINESSAREA", DOMAINNAME)

    'Read in Contacts
    For i = 1 To number_of_businessareas
        k = 2 * i
        j = 2
        While rngContacts.offset(k - 1, j).Text <> ""
        contactsTO(i) = contactsTO(i) & rngContacts.offset(k - 1, j).Text + "; "
        j = j + 1
        Wend
        
        j = 2
        While rngContacts.offset(k, j).Text <> ""
        contactsCC(i) = contactsCC(i) & rngContacts.offset(k, j).Text + "; "
        j = j + 1
        Wend
    Next i
    
    For i = 1 To number_of_workpackages
        PACKAGENAME = workPackages(i)
        PACKAGESHORTNAME = workPackagesShortID(i)
        
        If workPackagesLink(i) <> "" Then
            PACKAGELINK = "<a href=""" & workPackagesLink(i) & """>" & workPackages(i) & "</a>"
        Else
            PACKAGELINK = PACKAGENAME
        End If
        
        PACKAGESTATUS = getActiveAction(domain, workPackages(i))
        
        actionsString = ""
        activeaction = Actiontoi(PACKAGESTATUS)
        
        'IF is not released
        If actionsSeeIfActive(Actiontoi(PACKAGESTATUS)) = "No" Then
            actionsString = "None"
        Else
            For j = 1 To workPackages_i_Actions(i)
                If actionsSetBy(workPackagesActionsID(i, j)) = "Company" Then
                    If (workPackagesActionsID(i, j) > activeaction) Then
                        'Generating ID Code
                        IDCODE = "[" & businessAreasID(iDomain) & workPackagesID(i) & actionsID(workPackagesActionsID(i, j)) & "]"
                        MAILTO = rngMailStatusReply.offset(0, 0).Value
                        MAILTO = Replace(MAILTO, "BUSINESSAREA", DOMAINNAME)
                        MAILTO = Replace(MAILTO, "PACKAGESHORTNAME", PACKAGESHORTNAME)
                        MAILTO = Replace(MAILTO, "ACTIONNAME", actions(workPackagesActionsID(i, j)))
                        MAILTO = Replace(MAILTO, "REPLYTO", REPLYTO)
                        MAILTO = Replace(MAILTO, "IDCODE", IDCODE)
                        actionsString = actionsString & "<a href=""" & MAILTO & """>" & actions(workPackagesActionsID(i, j)) & "</a> "
                    End If
                End If
            Next j
        End If
        
        If actionsString = "" Then
            actionsString = "None"
        End If
        
        If PACKAGESTATUS = "released" Then
            PACKAGESTATUS = "<strong>" + PACKAGESTATUS + "</strong>"
            Timestring = "<strong>" + Timestring + "</strong>"
        End If
        
        'Never forget to reinitialise for the replacement
        ISSUEMAILTO = rngMailIssueReply.offset(0, 0).Value
        ISSUEIDCODE = "[" & businessAreasID(iDomain) & workPackagesID(i) & "00" & "]"
        ISSUEMAILTO = Replace(ISSUEMAILTO, "PACKAGESHORTNAME", PACKAGESHORTNAME)
        ISSUEMAILTO = Replace(ISSUEMAILTO, "REPLYTO", REPLYTO)
        ISSUEMAILTO = Replace(ISSUEMAILTO, "ISSUEIDCODE", ISSUEIDCODE)
        issueString = "<a href=""" & ISSUEMAILTO & """>" & "Issue" & "</a> "
        
        body = body & rngMailBody.offset(0, 0).Value
        body = Replace(body, "PACKAGELINK", PACKAGELINK)
        body = Replace(body, "PACKAGESTATUS", PACKAGESTATUS)
        body = Replace(body, "ACTIONS", actionsString)
        body = Replace(body, "UPDATETIME", Timestring)
        body = Replace(body, "ISSUELINK", issueString)
                        
    Next i

    body = body & rngMailSignature
    
    With OutMail
        .To = contactsTO(iDomain)
        .cc = contactsCC(iDomain)
        .bcc = MAILBCC
        .subject = subject
        .HTMLBody = body
        .Sentonbehalfofname = MAILONBEHALF
        .Display
        'Or use .Display .Send
        '.GetInspector.CommandBars.FindControl(, 718).Execute
        '.GetInspector.CommandBars.FindControl(, 719).Execute
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
    
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    
End Sub



'/**
'* Summary: Look in the "Tracking" sheet for the passed doamin and package and returns their corresponding action/status.
'* @param: String domain, package
'* @return: Active action as String
'*/
Function getActiveAction(ByVal domain As String, ByVal package As String) As String

    Dim rngTracking As Range
    Dim code As String
    Dim toBeStartedID As String
    
    ' 01 is ID for "to be started"
    toBeStartedID = businessAreasID(Domaintoi(domain)) & workPackagesID(Packagetoi(package)) & "01"
    
    'Looking for the first entry of a package in Tracking
    Set wsTracking = Sheets("Tracking")
    Set rngTracking = wsTracking.Range("D:D").Find(toBeStartedID, LookIn:=xlValues)
        
    i_inferior = rngTracking.Row
    i_superior = rngTracking.Row + 10
    
    For i = i_inferior To i_superior
        If wsTracking.Cells(i, 5).Value = "Active" Then
        
            getActiveAction = wsTracking.Cells(i, 3).Value
            
            If wsTracking.Cells(i, 11).Value = "" Then
                wsTracking.Cells(i, 11).Value = Now  'Sets timestamp in tracking sheet
            End If
            Timestring = Now
            Exit For
        End If
    Next i
    
End Function
    

'/**
'* Summary: Split the "importedCode" String and pass substrings to "ChangeStatusWithID"
'*          Calls LogTime with import flag.
'* @param: String importedCode
'*/
Sub ParseID(importedCode As String)
    
    Dim rngTracking As Range
    Set rngTracking = ChangeStatusWithID(Left(importedCode, 2), Mid(importedCode, 3, 2), Mid(importedCode, 5, 2))
    
    LogTime rngTracking, 1

End Sub

'/**
'* Summary: Look in the "Tracking" sheet for the passed code and set the corresponding status to "Active"
'*          and everything else to "Inactive".
'* @param: String domain, package, action
'* @return: corresponding line as Range Object
'*/
Function ChangeStatusWithID(domain As String, package As String, action As String) As Range

    Dim subcode As String
    Dim code As String
    Dim wsTracking As Worksheet
    Dim number_of_states As Integer
    Dim i_inferior, i_superior As Integer
    Dim rngTracking As Range
    
    Set wsTracking = Sheets("Tracking")
    number_of_states = Sheets("IDs").Range("StateTable").Rows.Count
    
    subcode = domain & package
    code = domain & package & action
    
    Set rngTracking = wsTracking.Range("D:D").Find(code, LookIn:=xlValues) 'Find code in Tracking
    
    i_inferior = rngTracking.Row - number_of_states
    i_superior = rngTracking.Row + number_of_states
    
    If i_inferior < 2 Then
        i_inferior = 2
    End If
     
    For i = i_inferior To i_superior
        'Looking for first 4 numbers of ID
        If Left(wsTracking.Cells(i, 4).Text, 4) = subcode Then
            wsTracking.Cells(i, 5).Value = "Inactive"
        End If
    Next i
    
    'Set the row to active
    wsTracking.Cells(rngTracking.Row, 5).Value = "Active"
    
    Set ChangeStatusWithID = rngTracking

End Function



'/**
'* Summary: Get the 'businessAreas()' array index for the passed 'domain' string
'* @param: String domain
'* @return: index as Integer
'*/
Function Domaintoi(ByVal domain As String) As Integer
    
    For i = 1 To number_of_businessareas
        If businessAreas(i) = domain Then
            Domaintoi = i
            Exit For
        End If
    Next i
    
End Function

'/**
'* Summary: Get the 'workPackages()' array index for the passed 'package' string
'* @param: String package
'* @return: index as Integer
'*/
Function Packagetoi(ByVal package As String) As Integer
    
    For i = 1 To number_of_workpackages
        If workPackages(i) = package Then
            Packagetoi = i
            Exit For
        End If
    Next i
    
End Function

'/**
'* Summary: Get the 'actions()' array index for the passed action string
'* @param: String action
'* @return: index as Integer
'*/
Function Actiontoi(ByVal action As String) As Integer
    
    For i = 1 To number_of_states
        If actions(i) = action Then
            Actiontoi = i
            Exit For
        End If
    Next i
    
End Function


Sub Create()
    
    Read_data
    Create_Tracking_Sheet
    
    Dim range_matrix As Range
    Dim wsDashboard As Worksheet
    Dim lastCol As String
    Dim sourceCol As Range
    Dim fillCol As Range
    Dim sourceRange As Range
    Dim fillRange As Range
    
    
    Set wsDashboard = ThisWorkbook.Sheets("Dashboard")
    Set range_matrix = wsDashboard.Range("DashboardMatrix")
    
    ThisWorkbook.Sheets("Configuration").Range("C16:C60").ClearContents
    ThisWorkbook.Sheets("Configuration").Range("E16:E60").ClearContents
    ThisWorkbook.Sheets("Configuration").Cells(16, 5).Value = "<Email>"
    For i = 1 To number_of_businessareas
        j = i * 2
        Sheets("Configuration").Cells(14 + j, 3).Value = "=INDEX(BusinessAreaTable[Business Area]," & i & ")"
        'Sheets("Configuration").Cells(16, 5).Value = "INDEX(BusinessAreaTable[Business Area],i)"
        'Sheets("Configuration").Cells(16, 7).Value = "INDEX(BusinessAreaTable[Business Area],i)"
        'Sheets("Configuration").Cells(16, 3).Value = [BusinessAreaTable].Cells(1, 1)
    Next i
        
    
    ThisWorkbook.Sheets("Deadlines").Range("DeadlineTable").Rows("2:" & ThisWorkbook.Sheets("Deadlines").Range("DeadlineTable").Rows.Count).Delete
    For i = 2 To number_of_workpackages
        ThisWorkbook.Sheets("Deadlines").Range("DeadlineTable").ListObject.ListRows.Add
    Next i
    ThisWorkbook.Sheets("Deadlines").Cells(3, 2).Value = "1/1/2017  12:00:00 AM"
    ThisWorkbook.Sheets("Deadlines").Cells(4, 2).Value = "1/1/2019  12:00:00 AM"
    
    
    'https://www.thespreadsheetguru.com/blog/2014/6/20/the-vba-guide-to-listobject-excel-tables
    
    If range_matrix.Rows.Count > 1 Then
        range_matrix.Rows("2:" & range_matrix.Rows.Count).Delete
    End If
    
    For i = 2 To number_of_workpackages
        wsDashboard.Range("DashboardMatrix").ListObject.ListRows.Add
    Next i

    
    If wsDashboard.Range("DashboardMatrix").Columns.Count > 3 Then
        wsDashboard.Range("DashboardMatrix").offset(, 2).Resize(, wsDashboard.Range("DashboardMatrix").Columns.Count - 3).Columns.Delete
        For i = 2 To number_of_businessareas - 1
            wsDashboard.Range("DashboardMatrix").ListObject.ListColumns.Add
        Next i
    End If
    
    
    wsDashboard.Range("DashboardMatrix[[#All],[Column2]]").Rows(1).Value = "=CONCATENATE(VLOOKUP(C$2,BusinessAreaTable,2,FALSE),VLOOKUP($B3,WorkpackageTable,2,FALSE))"
    Set sourceCol = wsDashboard.Range("DashboardMatrix[[#All],[Column2]]").Rows(1)
    Set fillCol = wsDashboard.Range("DashboardMatrix[[#All],[Column2]]")
    sourceCol.AutoFill Destination:=fillCol
    
    Set sourceRange = wsDashboard.Range("DashboardMatrix[[#All],[Column2]]")
    Set fillRange = wsDashboard.Range("DashboardMatrix").offset(0, 1).Resize(number_of_workpackages, number_of_businessareas)
    sourceRange.AutoFill Destination:=fillRange
    
    
    For i = 1 To number_of_businessareas
        wsDashboard.Cells(2, i + 2).Value = "=INDEX(BusinessAreaTable[Business Area]," & i & ")"
    Next i

    
    Set range_matrix = wsDashboard.Range("DashboardMatrix")
    'ThisWorkbook.Sheets("Dashboard").Range("C3:E" & range_matrix.Rows.Count + 2).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    'ThisWorkbook.Sheets("Dashboard").Range("B3:E" & range_matrix.Rows.Count + 2).Borders(xlInsideVertical).LineStyle = xlContinuous
    ThisWorkbook.Sheets("Dashboard").Range(range_matrix.Address()).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    ThisWorkbook.Sheets("Dashboard").Range(range_matrix.Address()).Borders(xlInsideVertical).LineStyle = xlContinuous

End Sub






