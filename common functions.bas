Option Explicit


    '///------------------------------------------------------------------
    '///   Namespace:      Monstah Developers
    '///   Class:          common functions.bas
    '///   Description:    Common Functions
    '///   Date: 11/04/16   Initial Release
    '///------------------------------------------------------------------
    '///    Worksheet Functions
    '///
    '///    -   Optimize
    '///    -   WorkSheetExists
    '///    -   LastRow
    '///    -   NextBlankCell
    '///    -   LastColumnNumber
    '///    -   LastColumnLetter
    '///    
    '///
    '///    File and Directory Functions
    '///
    '///    -   CheckPath
    '///    -   FileExists
    '///    -   openFilePicker
    '///    -   SelectFolder
    '///    -   CollectionToArray
    '///------------------------------------------------------------------


Dim OrigCalculationMethod As Integer

Private Sub optimize(opt As Boolean = True)

    '///
    '///    Speeds up the execution of VBA Code
    '///
    '///    Uses Global variable OrigCalculationMethod
    '///
    '///    Optional Field:
    '///        opt - Flag to turn on /off optimization. 
    '///        Use optimize(True) at beginning of and code and 
    '///        optimize(False) at the end of code
    '///

    Application.EnableEvents = Not opt
    Application.ScreenUpdating = Not opt
    Application.DisplayStatusBar = Not opt
    If opt = True Then
        OrigCalculationMethod = Application.Calculation
       Application.Calculation = xlCalculationManual
    Else
        Application.Calculation = OrigCalculationMethod
    End If
End Sub


Public Function WorkSheetExists(ByVal strSheet as string) As Boolean
    
    '///
    '///    Used to check if Worksheet Name exists within current active WorkBook
    '///
    
    dim xlSheet as WorkSheet

    err.clear
    
    WorkSheetExists = False

    on error resume next
    set xlSheet = application.ActiveWorkbook.Sheets(strSheet)

    on error goto err_function

    WorkSheetExists = Not xlSheet is nothing
    set xlSheet = nothing 
    Exit Function

err_function:
    set xlSheet = nothing
    err.raise "Not Found", , "Work Sheet not found in current WorkBook"
    WorkSheetExists = False
End Function

Public Function LastRow(optional strCol as string = "A", Optional ByVal strSheet As String) As Long

    '///
    '///    Finds the last row number using the Ctrl+Shift+End Method
    '///
    '///    utilises the WorkSheetExists Function
    '///
    '///    Optional Fields
    '///        strCol - Column letter to check for last Row. Default is column A
    '///        strSheet - SheetName as a string
    '///

    dim xlSheet as WorkSheet

    LastRow = 0

    if strSheet = "" then 
        Set xlSheet = Application.ActiveWorkbook.ActiveSheet
    elseif WorkSheetExists(strSheet) then
        Set xlSheet = application.ActiveWorkbook.Sheets(strSheet)
    else 
        goto err_function
    End if

    on error goto err_column
    LastRow = xlSheet.Cells(tmpSheet.Rows.Count, strCol).End(xlUp).Rows
    set xlSheet = nothing 
    Exit Function

err_function:
    set xlSheet = nothing
    err.raise "Not Found", , "No valid Worksheet found"
    Exit Function

err_column:
    set xlSheet = nothing
    err.Raise "Error", , "Invalid Column provided"
End Function

Public Function NextBlankCell(ByVal rngCell As Range, Optional searchDown as boolean = True) As Long
    
    '///
    '///    Finds the next empty cell row number, in a given Range
    '///
    '///    Optional Fields
    '///        searchDown - direction of the search. Default is True (Down)
    '///

    err.Clear
    NextBlankCellDown = 0

    if not rngCell = nothing then 
        if searchDown = True then
            NextBlankCellDown = rngCell.Offset(1,0).end(xlDown).row
        else    
            NextBlankCellDown = rngCell.Offset(-1,0).end(xlUp).row
        End if
    else 
        goto err_function
    end if
    Exit Function

err_function:
    err.raise "Not Found", , "No valid Range found"
    Exit Function
End Function

Public Function LastColumnNumber(Optional ByVal intRow As Integer = 1, Optional ByVal strSheet As String) As Integer

    '///
    '///    Finds the last used column for a given row and returns the column number
    '///
    '///    Optional Fields
    '///        intRow - the row to search
    '///        strSheet - the sheet name as a string to searh
    '///
    
    Dim xlSheet As Object

    Err.Clear
    LastColumnNumber = 0


    if strSheet = "" then 
        set xlSheet = Application.ActiveWorkbook.ActiveSheet
    elseif WorkSheetExists(strSheet) then
        set xlSheet = application.ActiveWorkbook.Sheets(strSheet)
        
    else 
        goto err_function
    End if
    
    LastColumnNumber = tmpSheet.Cells(intRow, tmpSheet.Columns.Count).End(xlToLeft).Column
    set xlSheet = nothing 
    Exit Function
err_function:

    set xlSheet = Nothing
    err.raise "Not Found", , "No valid column found"
    
End Function

Public Function LastColumnLetter(Optional ByVal intRow As Integer = 1, Optional ByVal strSheet As String) As String

    '///
    '///    Finds the last used column for a given row and returns the column Letter
    '///
    '///    Utilises LastColumnNumber Function
    '///
    '///    Optional Fields
    '///        intRow - the row to search
    '///        strSheet - the sheet name as a string to searh
    '///

    Dim lastColNumber As Integer
    Dim arr As Variant
    dim xlSheet as WorkSheet

    lastColNumber = LastColumnNumber(intRow, strSheet)
    
    
    if strSheet = "" then 
        set xlSheet = Application.ActiveWorkbook.ActiveSheet
    elseif WorkSheetExists(strSheet) then
        set xlSheet = application.ActiveWorkbook.Sheets(strSheet)
        
    else 
        goto err_function
    End if

    If LastError_ = 0 Then
        arr = Split(xlSheet.Cells(1, lastColNumber).Address(True, False), "$")
        LastColumnLetter = arr(0)
    Else
        LastColumnLetter = "A"
    End If

    set xlSheet = nothing 
    Exit Function

 err_function:
    set xlSheet = Nothing
    err.raise "Not Found", , "No valid WorkSheet found"   
End Function

 '/// File and Folder Functions

Function CheckPath(ByVal strPath As String) As Boolean
        
        '///
        '///    Checks to see if passed path exists.
        '///
        '///    Required Field:
        '///        strPath - Full Path excluding filename and extention
        '///
            
        
        Dim fso As Object
        CheckPath = False
        
        err.clear
        
        on error goto err_function

        If trim(strPath) = "" Then GoTo err_function
        
        Set fso = CreateObject("scripting.filesystemobject")
        
        if fso = nothing then GoTo err_function

        If right(strPath, 1) <> "\" Then
                strPath = strPath & "\"
        End If
        
        CheckPath = fso.FolderExists(strPath)
        
        Exit Function
err_function:
        CheckPath = False
End Function

Public Function FileExists(ByVal strPath As String, Optional ByVal isOpen As Boolean = False) As Boolean
    
    '///
    '///    Checks to see if pased file name, including full path exists.
    '///
    '///    Required Field:
    '///        strPath - Full Path to file incuding filename and extention
    '///
    '///    Optional Field:
    '///        isOpen - Optional Flag to also test if the file is currently locked by another process. Default is False
    '///
    
    Dim fso As Object
    Dim lngFileNum As Long
    
    err.Clear
    Set fso = CreateObject("scripting.filesystemobject")
    
    FileExists = False
    
    '///    Check file exists
    If fso.FileExists(strPath) = False Then
        FileExists = False
        err.Raise "Error", , "File not found!"
        Exit Function
    End If
    
    FileExists = True
    
    '///    Check to see if file is locked
    If isOpen Then
        On Error Resume Next
        
        lngFileNum = FreeFile()
        Open strPath For Input Lock Read As #lngFileNum
        Close lngFileNum
        On Error GoTo 0
        If err.Number <> 0 Then
            FileExists = False
            err.Raise "Error", , "File already in use"
        End If
    End If
    
    
End Function

Public Function openFilePicker(ByVal strTitle As String, optional multiFile as Boolean = False) As String

    '///
    '///    Displays standard windows file picker
    '///
    '///    Optional Field:
    '///        strTitle - Sets the titlte of the file picker
    '///        multiFile - Single or multiple file selection. Default is False (Single file only)
    '///
    '///    Returns selected full file path as a string
    '///
    
    Dim fd As FileDialog
    Dim fileName As String
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.AllowMultiSelect = multiFile
    'set allowed file types
    If Not strTitle = Null Or Not strTitle = "" Then
        fd.Title = strTitle
    End If
    If fd.Show = True Then
        If fd.SelectedItems(1) <> vbNullString Then
            fileName = fd.SelectedItems(1)
        Else
            'set any nulls to empty string
            fileName = ""
        End If
    Else
        'return empty string if cancelled
        fileName = ""
    End If
    Set fd = Nothing
    openFilePicker = fileName
End Function

Public Function SelectFolder(Optional ByVal strTitle As String, optional multiFolder as Boolean = False) As String
        
    '///
    '///    Displays standard windows folder picker
    '///
    '///    Optional Field:
    '///        strTitle - Sets the title of the folder picker
    '///        multiFolder - Single or multiple folder selection. Default is False (Single file only)
    '///
    '///    Returns selected folder path as a string
    '///
    
    Dim fd As FileDialog
    Dim strFolder As String
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.AllowMultiSelect = multiFile
    If Not strTitle = Null Or Not strTitle = "" Then
        fd.Title = strTitle
    End If
    
    If fd.Show = True Then
            strFolder = fd.SelectedItems(1)
    End If
    
    Set fd = Nothing
    SelectFolder = strFolder
End Function

 '/// Arrays

Public Function CollectionToArray(ByVal C As Collection) As Variant()
    
    '///
    '///    Converts a Collection to an array
    '///
    '///    Returns an array
    '///
    
    Dim tmpArr() As Variant
    ReDim tmpArr(0 To C.Count - 1)
    Dim i As Integer
    For i = 1 To C.Count
        tmpArr(i - 1) = C.Item(i)
    Next
    CollectionToArray = tmpArr
    
End Function
