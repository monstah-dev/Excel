    
'///------------------------------------------------------------------
'///   Class:          clsProject
'///   Description:    Wrapper Class for MSProject
'///   Author:         Monstah Developers               Date: 11/04/16
'///   Notes:
'///            Filter -> Upto 3 Options
'///            Clear Filter
'///            OpenFile
'///            CloseFile
'///            TimePhaseDate
'///            Count -> number of tasks in the active view
'///            UpdateProgress -> Update the Status Date
'///            SelectData (Unlimited Fields)
'///            SelectSubData (ParentID, clearDict, Unlimited Fields)
'///            GetData -> returns SelectData and SelectSubData result
'///            GetTimePhaseData -> returns timephasedata result
'///            SelectView -> either Gantt or Resource Charts
'///
'///   Revision History:
'///   Date:        Description:
'///   11/04/16            Initial Release
'///   26/04/17            Commented out quit on save for this project
'///   03/05/17            Added Function for timephasing %complete
'///                       Added Enums for Timephasedata, Clause Operator, Views, and timescale
'///                       Updated Filter to include created enums and extend filtering options
'///                       Updated Change view to include enums
'///                       added Method to set prject file save option
'///                       added Method to cancel close of project
'///                       added methods to set project and file if class is used within MSProject
'///                       Extended SelectData to allow for fields with two layers e.g. OutlineParent.UniqueID
'///------------------------------------------------------------------


'///------------------------------------------------------------------
'///    SAMPLE Usage
'///------------------------------------------------------------------
'///    Dim prjData as Object
'///    Dim prjApp as New clsMSProject
'///    Dim strKey as Variant                                   'stores the Dictionary Key
'///    Dim col as Object
'///    prjApp.OpenFile "File Location"                         'Open Project File
'///    prjApp.Save prjDoNotSave                                'Do not save on Exit
'///    prjApp.CloseOnExit false                                'Close on Applcation/File on Exit
'///    prjApp.SelectView PRJ_VIEW_GANTTCHART                   'Select Gantt Chart
'///    prjApp.ClearFilter                                      'Clear any existing Filter
'///    prjApp.Filter "Text1", "Somevalue", PRJ_CLAUSE_EQUALS   'Set any filters
'///    prjApp.SelectData Field1, Field2, Field3                'Select required fields
'///    prjData = prjApp.GetData                                'retrieve the data
'///    for each strKey in prjData                              'loop through returned data dictionary
'///            set col = prjData(strKey)                       'Set the collection
'///            Do something with collection data here
'///    Next strKey
'///    Set prjApp = Nothing
'///------------------------------------------------------------------

Option Explicit

'///------------------------------------------------------------------
'///
'/// Global Variables
'///
'///------------------------------------------------------------------

'///    enum of Project Save values.
Public Enum pj
    prjDoNotSave = 0
    prjPromptSave = 2
    prjSave = 1
End Enum

Private prjApp As Object        'Holds the MSProject Application Object
Private prjFile As Object       'Holds the MSProject Applcation File Object
Private tpDic_ As Object        'Dictionary for TimePhase Data
Private dataDic_ As Object      'Dictionary for Select Data
Private filterCount As Long     'Keeps track of the number of filters created
Private isProject As Boolean    'Set with CloseOnExit property; default is False which closes MSProject File and Application on termination of class
Private pjSave As pj            'Varaiable to store whether to save the file on exit. Default is Do Not Save.

Private Enum PjProjectUpdate
   pj0or100Percent = 0  'Sets only the Actual Start and Actual Finish dates.
   pj0to100Percent = 1  'Sets the percent complete to reflect the update date.
   pjReschedule = 2             'Schedules the remainder of the work to start on the update date.
End Enum

'///    Enumerator of MSProject TimeScale

Public Enum pjTimescale
    pjTimescaleDays = 4
    pjTimescaleHalfYears = 8
    pjTimescaleHours = 5
    pjTimescaleMinutes = 6
    pjTimescaleMonths = 2
    pjTimescaleNone = 255
    pjTimescaleQuarters = 1
    pjTimescaleThirdsOfMonths = 7
    pjTimescaleWeeks = 3
    pjTimescaleYears = 0
End Enum

'///    Custom Enumerator for Filtering of Data. Used in Conjuction with TransformOp Function
Public Enum PRJ_CLAUSE_OPERATOR
    PRJ_CLAUSE_EQUALS
    PRJ_CLAUSE_GREATERTHAN
    PRJ_CLAUSE_LESSTHAN
    PRJ_CLAUSE_GREATERTHANOREQUAL
    PRJ_CLAUSE_LESSTHANOREQUAL
    PRJ_CLAUSE_DOESNOTEQUAL
    PRJ_CLAUSE_CONTAINS
    PRJ_CLAUSE_DOESNOTCONTAIN
End Enum
'///    Custom Enumerator for changing of MSProject views. Used in Conjuction with TransformView Function
Public Enum PRJ_VIEWS
    PRJ_VIEW_BARROLLUP
    PRJ_VIEW_DETAILGANTT
    PRJ_VIEW_GANTTCHART
    PRJ_VIEW_LEVELINGGANTT
    PRJ_VIEW_MILESTONE
    PRJ_VIEW_TRACKINGGANTT
    PRJ_VIEW_TASKUSAGE
    PRJ_VIEW_RESOURCEUSAGE
End Enum

'///    Enumerator of MSProject TimeScaledData for Office 2010 -2014
Public Enum PjTaskTimescaledData
    pjTaskTimescaledRegularWork = 166
    pjTaskTimescaledActualWork = 2
    pjTaskTimescaledBaselineWork = 1
    pjTaskTimescaledCumulativeWork = 176
    pjTaskTimescaledPercentComplete = 32
End Enum

Public Enum prjTaskOutline
    prjOutlineShowLevel1 = 1
    prjOutlineShowLevel2 = 2
    prjOutlineShowLevel3 = 3
    prjOutlineShowLevel4 = 4
    prjOutlineShowLevel5 = 5
    prjOutlineShowLevel6 = 6
    prjOutlineShowLevel7 = 7
    prjOutlineShowLevel8 = 8
    prjOutlineShowLevel9 = 9
    prjOutlineShowLevelAll = 65535
End Enum

Private Sub Class_initialize()

    '///
    '///    Initilization Class
    '///
    '///     Set Global Objects and sets initial settings
    '///

    'Create Objects
    Set prjApp = createObject("MSProject.Application")
    Set tpDic_ = createObject("Scripting.Dictionary")
    Set dataDic_ = createObject("Scripting.Dictionary")
    pjSave = prjDoNotSave
    'Set MS Project Properties
    prjApp.Application.DisplayAlerts = False
    prjApp.Application.ScreenUpdating = False
    prjApp.Application.DisplayStatusBar = False
    prjApp.Application.Calculation = -1
    prjApp.Visible = False
    filterCount = 0
End Sub


Private Sub Class_Terminate()

    '///
    '///     Terminiation Class
    '///
    '///     Destroys Global Objects
    '///

    On Error Resume Next
    'Close MS Project - Do not save changes
    
    prjApp.Application.DisplayAlerts = True
    prjApp.Application.ScreenUpdating = True
    prjApp.Application.DisplayStatusBar = True
    
    'close if no longer required; default is to close
    If isProject = False Then
        closeFile
        prjApp.Quit
    End If
        Set prjApp = Nothing
        Set prjFile = Nothing
    Set tpDic_ = Nothing
    Set dataDic_ = Nothing
End Sub

Sub openFile(ByVal FilePath As String)
        
    '///
    '///    FilePath - String containing path to file location including FileName
    '///
    '///    Sets the prjApp and prjFile Global Variables
    '///
    
    If prjApp.fileOpenEx(Name:=FilePath, ReadOnly:=False) Then
        Set prjFile = prjApp.ActiveProject
    Else
        MsgBox "Failed to open " & FilePath
        End 'end all further VBA Calls
    End If
End Sub

Sub UpdateProject(uDate As Date)
    '///
    '///        update MSProject file status date and sets the Update date to the passed date value
    '///
    prjApp.UpdateProject all:=True, UpdateDate:=uDate, Action:=pj0to100Percent
End Sub

Sub Filter(ByVal field As String, ByVal criteria As String, Optional ByVal match As PRJ_CLAUSE_OPERATOR = PRJ_CLAUSE_EQUALS, Optional ByVal Field2 As String, Optional ByVal criteria2 As String, Optional ByVal match2 As PRJ_CLAUSE_OPERATOR = PRJ_CLAUSE_EQUALS, Optional ByVal Operation As String, Optional ByVal field3 As String, Optional ByVal criteria3 As String, Optional ByVal match3 As PRJ_CLAUSE_OPERATOR = PRJ_CLAUSE_EQUALS, Optional ByVal Operation2 As String, Optional ByVal view As PRJ_VIEWS = PRJ_VIEW_GANTTCHART, Optional ByVal showSummary As Boolean = False)
                
    '///
    '///    Creates and applies filters to the selected View. Filters do not show in Menu
    '///
    '///    Required Fields:
    '///            field - The MSProject field to filter
    '///            criteria - the criteria the filter is to apply
    '///            match - the condition the field and criteria need to meet.
    '///
    '///    Optional Fields:
    '///            field2, field3 criteria2, criteria3, match2, match3 - same as above
    '///            Operation, Operation2 - Either AND / OR used to join filter requiments
    '///            view - Select the required view. Default is Gantt Chart
    '///            showSummary - Show Summary Tasks in Filter. Default is False
    '///      
                
    'clear any existing filters and select the appropriate view
    prjApp.FilterClear
    'Default view is Gantt Chart
    prjApp.ViewApply TransformView(view)
    
    'build first part of filter based on initial options
    prjApp.FilterEdit "Filter" & filterCount, TaskFilter:=True, Create:=True, OverwriteExisting:=True, FieldName:=field, Test:=TransformOp(match), Value:=criteria, ShowInMenu:=False, ShowSummaryTasks:=showSummary
    
    'check to see if a second criteria has been entered
    If Field2 <> "" And criteria2 <> "" Then
            
        'set default operation to AND if not selected
        If Operation <> "AND" And Operation <> "Or" Then
                Operation = "AND"
        End If
        
        'extend filter if required
        prjApp.FilterEdit Name:="Filter" & filterCount, TaskFilter:=True, FieldName:="", NewFieldName:=Field2, Test:=TransformOp(match2), Value:=criteria2, Operation:=Operation, ShowSummaryTasks:=showSummary
    End If
    
    'check to see if a second criteria has been entered
    If field3 <> "" And criteria3 <> "" Then
            
        'set default operation to AND if not selected
        If Operation2 <> "AND" And Operation <> "Or" Then
                Operation2 = "AND"
        End If
        
        'extend filter if required
        prjApp.FilterEdit Name:="Filter" & filterCount, TaskFilter:=True, FieldName:="", NewFieldName:=field3, Test:=TransformOp(match3), Value:=criteria3, Operation:=Operation2, ShowSummaryTasks:=showSummary
    End If
    
    'Apply the filter
    prjApp.FilterApply "Filter" & filterCount
    
    '!important: Select the sheet for use
    prjApp.SelectSheet
    
    'Increase the number of filters
    filterCount = filterCount + 1
        
End Sub

Sub closeFile()
    '///
    '///    Closes Current Project File
    '///
                
    prjApp.FileCloseEx pjSave
    prjFile = Nothing
End Sub

Sub ClearFilter(Optional ByVal view As PRJ_VIEWS = PRJ_VIEW_GANTTCHART)
    '///
    '///    Clears all filters from Selected View
    '///
    '///    view - Select the required MSProject view. Default is Gantt Chart
    '///
      
    'Default clears Filter from current view
    Dim strTempView As String
    
    'store current view in case view changes
    strTempView = prjFile.CurrentView
    
    'CLEAR FILTER FROM VIEW
    prjApp.ViewApply TransformView(view)
    prjApp.FilterClear
    
    'reslect current view and select sheet
    prjApp.ViewApply strTempView
    prjApp.SelectSheet
End Sub

Sub TimePhasePercentage(ByVal scales As Integer, Optional ByVal timescale As pjTimescale = pjTimescaleHours, Optional ByVal selection As PjTaskTimescaledData = pjTaskTimescaledRegularWork, Optional ByVal startDate As Date = "1/1/1901", Optional ByVal endDate As Date = "1/1/1901")
    '///
    '///    Default Usage: TimePhase Data as a percentage of Task WorkHours
    '///
    '///    Each Task.Name, Task.WorkHours and TimePhaseData information is added to a Collection
    '///    Collection is added to tpDic_ Dictionary
    '///    Use GetTimePhaseData to return tpDic_
    '///
    '///    Required Fields:
    '///            Scales - Integer representing the TimeScale Multiplyer
    '///
    '///    Optional Fields:
    '///            timescale - enum pjTimescale. Determines the base scale of the Time Phase Data e.g. Hours, Days, Weeks, etc. Default is Hours
    '///            selection - enum PjTaskTimescaledData. The Type of data to use; Default is RegularWork.
    '///            startDate - Start of TimePhaseData. Default is Baseline Start
    '///            endDate - End of TimePhaseData. Default is Baseline Finish
    '///
        
    Dim prjTSV As Object
    Dim prjPTSV As Object
    Dim prjTask As Object
    Dim colTask As Collection
    Dim intTimePhase As Integer
    Dim tmpHours As Double
    Dim i As Long
    
    tpDic_.RemoveAll
    'set time in Hours with minimum of one hour
    If scales < 1 Then
            scales = 1
    End If
    If prjApp.ActiveSelection.Tasks.Count > 0 Then
        If startDate = "1/1/1901" Then
                startDate = prjFile.ProjectSummaryTask.Start ' default to Baselinbe start date if no date is propvided
        End If
        If endDate = "1/1/1901" Then
                endDate = prjFile.ProjectSummaryTask.Finish  ' default to Baselinbe Finish date if no date is propvided
        End If
        
        Set prjPTSV = prjFile.ProjectSummaryTask.TimeScaleData(startDate, endDate, , timescale, scales) 'pjTimescaleHours returns timephase data in hours
        For Each prjTask In prjApp.ActiveSelection.Tasks ' Loop through filtered tasks
            Set colTask = New Collection
            colTask.add prjTask.Name
            tmpHours = prjTask.Work / 60 'total work hours. divide by 60 to turn minutes into hours
            colTask.add tmpHours
            Set prjPTSV = prjTask.TimeScaleData(prjTask.Start, prjTask.Finish, selection, timescale, scales)
                            
            For i = 1 To prjPTSV.Count 'Loop through aggregated time phase data and add to collection
                colTask.add (prjPTSV(i) / 60) / tmpHours
            Next i
            tpDic_.add prjTask.UniqueID, colTask 'add collection to dictionary ussing UniqueID
            Set colTask = Nothing
        Next prjTask
    End If
End Sub


Sub TimePhaseHoursData(ByVal scales As Integer, Optional ByVal timescale As pjTimescale = pjTimescaleHours, Optional ByVal startDate As Date, Optional ByVal endDate As Date)
    '///
    '///    Default Usage: TimePhase for each task is current filtered view
    '///    Returns 4 timephasedata lines per task based:
    '///            1) Work
    '///            2) Actual Work
    '///            3) Baseline Work
    '///            4) Cummlative Work
    '///
    '///    Each Task.Name, Task.WorkHours and TimePhaseData information is added to a Collection
    '///    Collection is added to tpDic_ Dictionary
    '///    Use GetTimePhaseData to return tpDic_
    '///
    '///    Required Fields:
    '///            Scales - Integer representing the TimeScale Multiplyer
    '///
    '///    Optional Fields:
    '///            timescale - enum pjTimescale. Determines the base scale of the Time Phase Data e.g. Hours, Days, Weeks, etc. Default is Hours
    '///            startDate - Start of TimePhaseData. Default is Baseline Start
    '///            endDate - End of TimePhaseData. Default is Baseline Finish
    '///
    
    Dim prjTSV As Object
    Dim prjPTSV As Object
    Dim prjTask As Object
    Dim colTask As Collection
    Dim intTimePhase As Integer
    Dim i As Long
    Dim j As Integer
    
    'set time in Hours with minimum of one hour
    If scales < 1 Then
            scales = 1
    End If
    
    'clear existing dictionary
    tpDic_.RemoveAll
    
    If Not prjFile Is Nothing Then
        If IsNull(startDate) Then
                startDate = prjFile.ProjectSummaryTask.Start ' default to Baselinbe start date if no date is propvided
        End If
        If IsNull(endDate) Then
                endDate = prjFile.ProjectSummaryTask.Finish  ' default to Baselinbe Finish date if no date is propvided
        End If
        
        If prjApp.ActiveSelection.Tasks.Count > 0 Then
            For Each prjTask In prjApp.ActiveSelection.Tasks ' Loop through filtered tasks
                For j = 1 To 4
                    Set colTask = New Collection ' Create new collection
                    colTask.add prjTask.Name
                    colTask.add prjTask.Work / 60 'divide by 60 to turn minutes into hours
                    
                    Select Case j
                        Case 1
                            colTask.add "Work" ' Work Line
                            intTimePhase = pjTaskTimescaledRegularWork
                        Case 2
                            colTask.add "act. Work" ' Actual Work Line
                            intTimePhase = pjTaskTimescaledActualWork
                        Case 3
                            colTask.add "Base. Work" ' Baseline Work
                            intTimePhase = pjTaskTimescaledBaselineWork
                        Case 4
                            colTask.add "cuml. Work" ' Cummlative Work
                            intTimePhase = pjTaskTimescaledCumulativeWork
                    End Select
                    
                    Set prjPTSV = prjTask.TimeScaleData(prjTask.Start, prjTask.Finish, intTimePhase, timescale, scales)
                    
                    For i = 1 To prjPTSV.Count 'Loop through aggregated time phase data and add to collection
                            If isError(prjPTSV(i) / 60) Then
                                colTask.add 0
                            Else
                                colTask.add prjPTSV(i) / 60
                            End If
                    Next i
                    tpDic_.add key:=prjTask.UniqueID & j, Value:=colTask 'add collection to dictionary ussing UniqueID & the value of J
                    Set colTask = Nothing ' Destroy Collection
                Next j
            Next prjTask
        End If
    End If
End Sub

Sub SelectData(ParamArray Fields() As Variant)

    '///
    '///    For each task within current filtered data set, selected Field(s) are added to a collection, which is then added to the dictionary dataDic_
    '///
    '///    Required Fields:
    '///
    '///            Fields - ParamArray to allow for infinite number of fields to be selected
    '///
    '///    Data to be returned through function GetData
    '///
            
    On Error Resume Next
    Dim prjTask As Object
    Dim field As Variant
    Dim colTask As Collection
    Dim i As Integer
    Dim arrTemp() As String
    Dim strTemp As String
    Dim varTemp As Object
    
    'clear dictionary
    dataDic_.RemoveAll
    
    If Not prjFile Is Nothing Then
        If prjApp.ActiveSelection.Tasks.Count > 0 Then 'Loop through all filtered tasks
            For Each prjTask In prjApp.ActiveSelection.Tasks
                Set colTask = New Collection
                For Each field In Fields
                    If InStr(field, ".") Then
                        'if . is found in string assume special field i.e. OulineParent.UniqueID
                        arrTemp = Split(field, ".") 'split field
                        Set varTemp = CallByName(prjTask, arrTemp(0), VbGet) 'get Object of first part i.e. OulineParent
                        strTemp = CallByName(varTemp, arrTemp(1), VbGet) ' get child value of special field i.e. UniqueID of OulineParent.UniqueID
                        Set varTemp = Nothing ' Memory Management
                    Else
                        strTemp = CallByName(prjTask, field, VbGet) 'Get task Field Value
                    End If
                    colTask.add strTemp ' add field value to collection
                Next field
                dataDic_.add prjTask.UniqueID, colTask 'add collection to dictionary using UniqueID as key
                Set colTask = Nothing
            Next prjTask
        End If
    End If
End Sub

Public Sub SelectSubDataSummary(ParentID As String, clearDic As Boolean, ParamArray Fields() As Variant)
                
    '///
    '///    For each task within current filtered data set, Find Task with UniqueID equal to ParentID, and then the selected Field(s) for the Children Tasks are added to a collection recursivley, which is then added to the dictionary dataDic_
    '///
    '///    Required Fields:
    '///
    '///    Fields - ParamArray to allow for infinite number of fields to be selected
    '///
    '///    Data to be returned through function GetData
    '///
    
    On Error Resume Next
    Dim prjTask As Object
    Dim field As Variant
    Dim colTask As Collection
    Dim strTemp As String
    Dim parentTask As Object
    Dim Sibling As Object
    
    'clear dictionary
    If clearDic Then
        dataDic_.RemoveAll
    End If
    
    'find task associated with ParentID
    prjApp.Find field:="UniqueID", Test:="equals", Value:=ParentID
    Set parentTask = prjApp.ActiveCell.task
    
    'Loop through Task Siblings
    For Each Sibling In parentTask.outlineChildren
        If Sibling.UniqueID <> CLng(ParentID) Then
            If Sibling.outlineChildren.Count > 0 Then
                Set colTask = New Collection
                For Each field In Fields
                        strTemp = CallByName(Sibling, field, VbGet) 'Get Sibling Field Value
                        colTask.add strTemp 'add field value to collection
                Next field
                dataDic_.add Sibling.UniqueID, colTask 'add collection to dictionary using UniqueID as key
                Set colTask = Nothing
                SelectSubDataSummary Sibling.UniqueID, False, Fields 'recursivley check childs for their own siblings
            End If
        End If
    Next Sibling

End Sub

Sub SelectSubDataDetail(ParentID As Variant, clearDic As Boolean, ParamArray Fields() As Variant)
                
    '///
    '///    For each task within current filtered data set, Find Task with UniqueID equal to ParentID, and then the selected Field(s) for the Children Tasks are added to a collection recursivley, which is then added to the dictionary dataDic_
    '///
    '///    Required Fields:
    '///
    '///    Fields - ParamArray to allow for infinite number of fields to be selected
    '///
    '///    Data to be returned through function GetData
    '///
    
    On Error Resume Next
    Dim prjTask As Object
    Dim field As Variant
    Dim colTask As Collection
    Dim strTemp As String
    Dim parentTask As Object
    Dim Sibling As Object
    
    'clear dictionary
    If clearDic Then
        dataDic_.RemoveAll
    End If
    
    'find task associated with ParentID
    prjApp.Find field:="UniqueID", Test:="equals", Value:=ParentID
    Set parentTask = prjApp.ActiveCell.task
    
    'Loop through Task Siblings
    For Each Sibling In parentTask.outlineChildren
        If Sibling.UniqueID <> CLng(ParentID) Then
            Set colTask = New Collection
            For Each field In Fields
                    strTemp = CallByName(Sibling, field, VbGet) 'Get Sibling Field Value
                    colTask.add strTemp 'add field value to collection
            Next field
            dataDic_.add Sibling.UniqueID, colTask 'add collection to dictionary using UniqueID as key
            Set colTask = Nothing
            SelectSubDataDetail Sibling.UniqueID, False, Fields 'recursivley check childs for their own siblings
        End If
    Next Sibling

End Sub

Sub SelectView(Optional ByVal view As PRJ_VIEWS = PRJ_VIEW_GANTTCHART)
    '///
    '/// Change the view of MS Project. Default is Gantt Chart
    '///

    prjApp.ViewApply TransformView(view)
        
End Sub

Private Function TransformOp(ByVal op As PRJ_CLAUSE_OPERATOR) As String
    '///
    '/// Change the view of MS Project. Default is Gantt Chart
    '///    required Field:
    '///
    '/// op as a PRJ_CLAUSE_OPERATOR Default is equals
    '///
    '///    returns the MSProject Filter string equivalent
    '///

    Dim strOperation As String
    
    Select Case op
        Case PRJ_CLAUSE_EQUALS
            strOperation = "equals"
        Case PRJ_CLAUSE_GREATERTHAN
            strOperation = "is greater than"
        Case PRJ_CLAUSE_LESSTHAN
            strOperation = "is less than"
        Case PRJ_CLAUSE_GREATERTHANOREQUAL
            strOperation = "is greater than or equal to"
        Case PRJ_CLAUSE_LESSTHANOREQUAL
            strOperation = "is less than or equal to"
        Case PRJ_CLAUSE_DOESNOTEQUAL
            strOperation = "does not equal"
        Case PRJ_CLAUSE_CONTAINS
            strOperation = "contains"
        Case PRJ_CLAUSE_DOESNOTCONTAIN
            strOperation = "does not contain"
        Case Else
            strOperation = "equals"
    End Select
    TransformOp = strOperation
End Function

Private Function TransformView(ByVal view As PRJ_VIEWS) As String
    '///
    '/// Change the view of MS Project. Default is Gantt Chart
    '///    required Field:
    '///
    '///    view as a PRJ_VIEWS.
    '///
    '///    returns the MSProject View string equivalent. Default is Gantt Chart
    '///
        
    Dim strView As String
    Select Case view
    Case PRJ_VIEW_BARROLLUP
        strView = "Bar Rollup"
    Case PRJ_VIEW_DETAILGANTT
        strView = "Detail Gantt"
    Case PRJ_VIEW_GANTTCHART
        strView = "Gantt Chart"
    Case PRJ_VIEW_LEVELINGGANTT
        strView = "Leveling Gantt"
    Case PRJ_VIEW_MILESTONE
        strView = "Milestone Rollup"
    Case PRJ_VIEW_TRACKINGGANTT
        strView = "Tracking Gantt"
    Case PRJ_VIEW_TASKUSAGE
        strView = "Task Usage"
    Case PRJ_VIEW_RESOURCEUSAGE
        strView = "Resource Usage"
    Case Else
        strView = "Gantt Chart"
    End Select
    TransformView = strView
End Function

Public Property Get Count() As Long
    '///
    '/// Returns the Count of tasks in the Active View
    '///
    Count = prjApp.ActiveSelection.Tasks.Count
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Setters and Getters
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get TimePhaseDataResult() As Object
        Set TimePhaseDataResult = tpDic_
End Property

Public Property Get GetData() As Object
        'Returns last updated dictionary
        Set GetData = dataDic_
End Property

Public Property Get NewEnum() As IUnknown
    Dim objCol As Object
    'internal enum for objects with set ATTRIBUTES. Use Notepad to view ATTRIBUTES
    Set NewEnum = objCol.[_NewEnum]
End Property

Public Property Let SetApp(ByVal app As Object)
    'set Project Application. For use in MSProject Only. Use OpenFile() Otherwise
    Set prjApp = app
End Property

Public Property Let SetFile(ByVal file As Object)
    'set Project File. For use in MSProject Only. Use OpenFile() Otherwise
    Set prjFile = file
End Property

Public Property Let CloseOnExit(ByVal bool As Boolean)
    'Set wether to close the MSPRoject File and Applcation on Exit. False is to close Applcation/File True to keep Application/File open
    isProject = bool
End Property

Public Property Let Save(ByVal s As pj)
    'Set wether to Save the MSPRoject File and Applcation on Exit.
    pjSave = s
End Property

Public Sub ShowAllSubTasks(Optional ByVal outlineLevel As prjTaskOutline = prjOutlineShowLevelAll)

    '///
    '///    Change the shown level of tasks within MS Project.
    '///    OPtional Field:
    '///
    '///    outlineLevel as a prjTaskOutline. Sets the outline level 1-9 and all
    '///
        
    prjApp.OutlineShowTasks OutlineNumber:=outlineLevel, ExpandInsertedProject:=True
End Sub

Public Sub SelectSubData(ParentID As String, clearDic As Boolean, ParamArray Fields() As Variant)
                
    '///
    '///    For each task within current filtered data set, Find Task with UniqueID equal to ParentID, and then the selected Field(s) for the Children Tasks are added to a collection recursivley, which is then added to the dictionary dataDic_
    '///
    '///    Required Fields:
    '///
    '///    Fields - ParamArray to allow for infinite number of fields to be selected
    '///
    '///    Data to be returned through function GetData
    '///
    
    On Error Resume Next
    Dim prjTask As Object
    Dim field As Variant
    Dim colTask As Collection
    Dim strTemp As String
    Dim parentTask As Object
    Dim Sibling As Object
    
    'clear dictionary
    If clearDic Then
        dataDic_.RemoveAll
    End If
    
    'find task associated with ParentID
    prjApp.Find field:="UniqueID", Test:="equals", Value:=ParentID
    Set parentTask = prjApp.ActiveCell.task
    
    'Loop through Task Siblings
    For Each Sibling In parentTask.outlineChildren
        If Sibling.UniqueID <> CLng(ParentID) Then
            Set colTask = New Collection
            For Each field In Fields
                    strTemp = CallByName(Sibling, field, VbGet) 'Get Sibling Field Value
                    colTask.add strTemp 'add field value to collection
            Next field
            dataDic_.add Sibling.UniqueID, colTask 'add collection to dictionary using UniqueID as key
            Set colTask = Nothing
        End If
    Next Sibling

End Sub
