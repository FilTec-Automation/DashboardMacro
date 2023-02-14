Public Function CollectUniques(rng As Range) As Collection
    
    Dim varArray As Variant, var As Variant
    Dim col As Collection
    
    'Guard clause - if Range is nothing, return a Nothing collection
    'Guard clause - if Range is empty, return a Nothing collection
    If rng Is Nothing Or WorksheetFunction.CountA(rng) = 0 Then
        Set CollectUniques = col
        Exit Function
    End If
        
    If rng.Count = 1 Then '<~ check for a single cell range
        Set col = New Collection
        col.Add Item:=CStr(rng.Value), Key:=CStr(rng.Value)
    Else '<~ otherwise the range contains multiple cells
        
        'Convert the passed-in range to a Variant array for SPEED and bind the Collection
        varArray = rng.Value
        Set col = New Collection
        
        'Ignore errors temporarily, as each attempt to add a repeat
        'entry to the collection will cause an error
        On Error Resume Next
        
            'Loop through everything in the variant array, adding
            'to the collection if it's not an empty string
            For Each var In varArray
                If CStr(var) <> vbNullString Then
                    col.Add Item:=CStr(var), Key:=CStr(var)
                End If
            Next var
    
        On Error GoTo 0
    End If
    
    'Return the contains-uniques-only collection
    Set CollectUniques = col
    
End Function
Public Function UpdateNote(taskNotes() As String) As String
    If UBound(taskNotes) = 0 Then
        UpdateNote = ""
        Debug.Print "Task Notes are empty"
        Exit Function
    End If

    ReDim Preserve taskNotes(UBound(taskNotes) - 1)
    UpdateNote = Join(taskNotes, " | ")
    Debug.Print "Notes string has been updated. Last Item removed"
End Function

Public Function ValidateTaskNote(taskNote As String) As Boolean
    Dim taskNoteElements() As String
    taskNoteElements = Split(taskNote, ",")
    
    If UBound(taskNoteElements) = 2 Then
        Debug.Print "Task notes have 3 sections"
        
        ' check if the note has 3 parts to it
        If UBound(taskNoteElements) <> 2 Then
            ValidateTaskNote = False
            Debug.Print "Task Notes don't have their 3 sections: Name, Status, Date"
            Exit Function
        End If
    ElseIf UBound(taskNoteElements) = 3 Then
        Debug.Print "Task notes have 4 sections"
    Else
        ValidateTaskNote = False
        Debug.Print "Task Notes don't have their 3 sections: Name, Status, Date"
        Exit Function
    End If
       

    ' check if the note has 3 parts to it
    If UBound(taskNoteElements) <> 2 Then
        ValidateTaskNote = False
        Debug.Print "Task Notes don't have their 3 sections: Name, Status, Date"
        Exit Function
    End If

    ' split the date element into it's own components
    Dim dateNoteElements() As String
    dateNoteElements = Split(taskNoteElements(2), ":")
    
    ' check if the date element has 2 components
    If UBound(dateNoteElements) <> 1 Then
        ValidateTaskNote = False
        Debug.Print "Task note is missing the date isn't complete"
        Exit Function
    End If

    ' remove the whitespace around the date component of the date element
    Dim noteDate As String
    noteDate = Trim(dateNoteElements(1))

    ' check if the date components is complete
    If Len(noteDate) <> 8 Then
        ValidateTaskNote = False
        Debug.Print "The date isn't complete"
        Exit Function
    End If

    ValidateTaskNote = True
    Debug.Print "Task note is good"
End Function

Public Function ValidateNotes(noteTxt As String) As String
    
    Dim taskNotes() As String
    taskNotes = Split(noteTxt, "|")
    Dim lastTaskNote As String
    lastTaskNote = taskNotes(UBound(taskNotes))

    If Not ValidateTaskNote(lastTaskNote) Then
        ValidateNotes = UpdateNote(taskNotes)
        Debug.Print "Last task note didn't pass validation. Had to be updated"
        Exit Function
    End If

    ValidateNotes = noteTxt
    Debug.Print "Original Note has passed validation"
End Function


Sub UpdateDataFromEE()

    Dim Cn As ADODB.Connection
    Dim Server_Name As String
    Dim Database_Name As String
    Dim User_ID As String
    Dim Password As String
    Dim SQLStr As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim src As Range
    Dim NumRows As Long
    Dim X As Long
    Dim myTxt As String
    Dim countTxt As String

    Application.ScreenUpdating = False

    Sheets("EE Data").Activate

    Server_Name = "EASYENG"
    Database_Name = "Filtec"
    User_ID = "EasyEngReports"
    Password = "EasyEngReports123!@#"
    SQLStr = "ReportsJobStockDetail"

    Set Cn = New ADODB.Connection
    Cn.Open "Driver={SQL Server};Server=" & Server_Name & ";Database=" & Database_Name & _
    ";Uid=" & User_ID & ";Pwd=" & Password & ";"

    rs.Open SQLStr, Cn, adOpenStatic

    Worksheets("EE Data").Cells.ClearContents
    Worksheets("EE Data").Cells(1, 1).CopyFromRecordset rs

    Worksheets("EE Data").Rows(1).Insert

    Worksheets("EE Data").Cells(1, 1).Value = "ID"
    Worksheets("EE Data").Cells(1, 2).Value = "Job Number"
    Worksheets("EE Data").Cells(1, 3).Value = "Job Created Date"
    Worksheets("EE Data").Cells(1, 4).Value = "Job Status"
    Worksheets("EE Data").Cells(1, 5).Value = "Job Description"
    Worksheets("EE Data").Cells(1, 6).Value = "Customer"
    Worksheets("EE Data").Cells(1, 7).Value = "Stock ID"
    Worksheets("EE Data").Cells(1, 8).Value = "Stock Item"
    Worksheets("EE Data").Cells(1, 9).Value = "Stock Item Category"
    Worksheets("EE Data").Cells(1, 10).Value = "Item Code"
    Worksheets("EE Data").Cells(1, 11).Value = "Job Stock Description"
    Worksheets("EE Data").Cells(1, 12).Value = "Notes"
    Worksheets("EE Data").Cells(1, 13).Value = "Barcode"
    Worksheets("EE Data").Cells(1, 14).Value = "Category"
    Worksheets("EE Data").Cells(1, 15).Value = "Part Number"
    Worksheets("EE Data").Cells(1, 16).Value = "Material"
    Worksheets("EE Data").Cells(1, 17).Value = "Created Date"
    Worksheets("EE Data").Cells(1, 18).Value = "Due Date"
    Worksheets("EE Data").Cells(1, 19).Value = "Order Number"
    Worksheets("EE Data").Cells(1, 20).Value = "Supplier"
    Worksheets("EE Data").Cells(1, 21).Value = "Group"
    Worksheets("EE Data").Cells(1, 22).Value = "Machining Process"
    Worksheets("EE Data").Cells(1, 23).Value = "Comments"
    Worksheets("EE Data").Cells(1, 24).Value = "Branch Identifier"
    Worksheets("EE Data").Cells(1, 25).Value = "Drawing Required"
    Worksheets("EE Data").Cells(1, 26).Value = "PDF Hyperlink"
    Worksheets("EE Data").Cells(1, 27).Value = "Status"
    Worksheets("EE Data").Cells(1, 28).Value = "Quantity"
    Worksheets("EE Data").Cells(1, 29).Value = "Budget Rate"
    Worksheets("EE Data").Cells(1, 30).Value = "Actual Rate"
    Worksheets("EE Data").Cells(1, 31).Value = "Actual Units"
    Worksheets("EE Data").Cells(1, 32).Value = "High Priority"
    Worksheets("EE Data").Cells(1, 33).Value = "Complete"
    Worksheets("EE Data").Cells(1, 34).Value = "Date Added"
    Worksheets("EE Data").Cells(1, 35).Value = "Designer"
    Worksheets("EE Data").Cells(1, 36).Value = "Assembly"
    Worksheets("EE Data").Cells(1, 37).Value = "Budget Time"


    Set src = Worksheets("EE Data").Range("A1").CurrentRegion
    Worksheets("EE Data").ListObjects.Add(SourceType:=xlSrcRange, Source:=src, xlListObjectHasHeaders:=xlYes, tablestyleName:="TableStyleLight8").Name = "EasyEng_Table"

    Worksheets("EE Data").Cells(1, 38).Value = "Functions"

    X = 1
    NumRows = ActiveWorkbook.Worksheets("EE Data").Range("A1", Range("A1").End(xlDown)).Rows.Count
    countTxt = "|"

    For X = 2 To NumRows
    Dim debugMsg As String
    debugMsg = "*************************************************************** Row " & X & " ***************************************************************"
    Debug.Print debugMsg
    myTxt = Worksheets("EE Data").Cells(X, 12).Text
    
    If myTxt <> "" Then
        Dim note As String
        note = ValidateNotes(myTxt)
        
        Debug.Print "********************* myTxt *********************"
        Debug.Print myTxt
        
        Debug.Print "********************* new note *********************"
        Debug.Print note
    End If

    Worksheets("EE Data").Cells(X, 38).Value = (Len(myTxt) - Len(Replace(myTxt, countTxt, ""))) + 1

    Next

    Sheets("Dashboard").Activate
    Sheets("Dashboard").Cells(2, 6) = Now

    rs.Close
    Set rs = Nothing
    Cn.Close
    Set Cn = Nothing

    Application.ScreenUpdating = True
    MsgBox ("DataFrame Successfully Updated")

End Sub

Sub UpdateResourceLog()

    Dim NumStatuses As Integer
    Dim MaxStatuses As Integer
    Dim NumRows As Long
    Dim X As Long
    Dim Y As Integer
    Dim Z As Single
    Dim WorkArray() As String
    Dim TTime As Single
    Dim src As Range
    Dim mystring As String
    Dim P As Integer
    Dim FunctionString As String
    Dim MyDate As Date

    'Call ErrorHandler
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Sheets("Resource Log").Activate

    Worksheets("Resource Log").Cells.ClearContents

    Worksheets("EE Data").Range("EasyEng_Table[#All]").Copy Worksheets("Resource Log").Range("A1")
    Worksheets("Resource Log").ListObjects(1).Name = "Resource_Table"
    
    Worksheets("Resource Log").ListObjects("Resource_Table").Range.AutoFilter Field:=4, Criteria1:="At Design"
    Worksheets("Resource Log").ListObjects("Resource_Table").DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
    Worksheets("Resource Log").ListObjects("Resource_Table").AutoFilter.ShowAllData
            
    
    
    X = 1
    MaxStatuses = Application.WorksheetFunction.Max(Range("Resource_Table[Functions]"))
    
    For X = 1 To MaxStatuses
    
        Worksheets("Resource Log").Cells(1, 38 + X).Value = "Function " & X
    
    Next
    
    X = 1
    NumRows = ActiveWorkbook.Worksheets("Resource Log").Range("A1", Range("A1").End(xlDown)).Rows.Count
    
    For X = 2 To NumRows
    
        If Worksheets("Resource Log").Cells(X, 12) <> "" Then
            
            Y = 1
            
            NumStatuses = Worksheets("Resource Log").Cells(X, 38).Value
            WorkArray = Split(Worksheets("Resource Log").Cells(X, 12).Value, "|")
            
            For Y = 1 To NumStatuses
            
                Worksheets("Resource Log").Cells(X, 38 + Y).Value = WorkArray(Y - 1)
            
            Next
            
        End If
    
    Next
    
    'Split Substrings Further
    X = 1

    For X = 1 To MaxStatuses
    
        Worksheets("Resource Log").Cells(1, 38 + MaxStatuses + ((5 * X) - 4)).Value = "Name " & X
        Worksheets("Resource Log").Cells(1, 38 + MaxStatuses + ((5 * X) - 3)).Value = "Status " & X
        Worksheets("Resource Log").Cells(1, 38 + MaxStatuses + ((5 * X) - 2)).Value = "Category " & X
        Worksheets("Resource Log").Cells(1, 38 + MaxStatuses + ((5 * X) - 1)).Value = "Date " & X
        Worksheets("Resource Log").Cells(1, 38 + MaxStatuses + (5 * X)).Value = "Time " & X
    
    Next
    
    X = 1
    
    For X = 2 To NumRows
      
        If Worksheets("Resource Log").Cells(X, 12) <> "" Then
            
            Y = 1
            
            NumStatuses = Worksheets("Resource Log").Cells(X, 38).Value
            
            For Y = 1 To NumStatuses
            
                FunctionString = Worksheets("Resource Log").Cells(X, 38 + Y).Value
                
                If InStr(FunctionString, "Time") > 0 Then
                    
                    WorkArray = Split(FunctionString, ",")
                    Worksheets("Resource Log").Cells(X, 38 + MaxStatuses + ((5 * Y) - 4)).Value = WorkArray(0)
                    Worksheets("Resource Log").Cells(X, 38 + MaxStatuses + ((5 * Y) - 2)).Value = Right(WorkArray(1), Len(WorkArray(1)) - 6)
                    MyDate = Format(Right(WorkArray(2), Len(WorkArray(2)) - 6), "dd/mm/yy")
                    Worksheets("Resource Log").Cells(X, 38 + MaxStatuses + ((5 * Y) - 1)).Value = MyDate
                    Worksheets("Resource Log").Cells(X, 38 + MaxStatuses + (5 * Y)).Value = Right(WorkArray(3), Len(WorkArray(3)) - 7)
                
                Else
                
                    WorkArray = Split(FunctionString, ",")
                    Worksheets("Resource Log").Cells(X, 38 + MaxStatuses + ((5 * Y) - 4)).Value = WorkArray(0)
                    Worksheets("Resource Log").Cells(X, 38 + MaxStatuses + ((5 * Y) - 3)).Value = Right(WorkArray(1), Len(WorkArray(1)) - 8)
                    MyDate = Format(Right(WorkArray(2), Len(WorkArray(2)) - 6), "dd/mm/yy")
                    Worksheets("Resource Log").Cells(X, 38 + MaxStatuses + ((5 * Y) - 1)).Value = MyDate
                    
                End If
            
            Next
            
        End If
    
    Next
    
    Worksheets("Resource Log").Cells(1, 38 + MaxStatuses + (5 * MaxStatuses) + 1) = "Total Time"

    X = 1
    
    For X = 2 To NumRows
    
        Y = 1
        NumStatuses = Worksheets("Resource Log").Cells(X, 38).Value
        TTime = 0
        
        For Y = 1 To NumStatuses
        
            TTime = TTime + Worksheets("Resource Log").Cells(X, 38 + MaxStatuses + ((Y * 5)))
        
        Next
        
        Worksheets("Resource Log").Cells(X, 38 + MaxStatuses + (5 * MaxStatuses) + 1) = TTime
    
    Next

    Worksheets("Resource Log").Cells(1, 38 + MaxStatuses + (5 * MaxStatuses) + 2) = "Total Time as a Percentage of Budget Time"

    X = 1
    
    For X = 2 To NumRows
    
        If Worksheets("Resource Log").Cells(X, 37) <> 0 Then
            Worksheets("Resource Log").Cells(X, 38 + MaxStatuses + (5 * MaxStatuses) + 2) = (Worksheets("Resource Log").Cells(X, 38 + MaxStatuses + (5 * MaxStatuses) + 1)) / (Worksheets("Resource Log").Cells(X, 37))
        Else:
            Worksheets("Resource Log").Cells(X, 38 + MaxStatuses + (5 * MaxStatuses) + 2) = 0
        End If
    
    Next

    Worksheets("Resource Log").Cells(1, 38 + MaxStatuses + (5 * MaxStatuses) + 4) = "Name"
    Worksheets("Resource Log").Cells(1, 38 + MaxStatuses + (5 * MaxStatuses) + 5) = "Job Code"
    Worksheets("Resource Log").Cells(1, 38 + MaxStatuses + (5 * MaxStatuses) + 6) = "Barcode"
    Worksheets("Resource Log").Cells(1, 38 + MaxStatuses + (5 * MaxStatuses) + 7) = "Part Number"
    Worksheets("Resource Log").Cells(1, 38 + MaxStatuses + (5 * MaxStatuses) + 8) = "Drawing Link"
    Worksheets("Resource Log").Cells(1, 38 + MaxStatuses + (5 * MaxStatuses) + 9) = "Status"
    Worksheets("Resource Log").Cells(1, 38 + MaxStatuses + (5 * MaxStatuses) + 10) = "Category"
    Worksheets("Resource Log").Cells(1, 38 + MaxStatuses + (5 * MaxStatuses) + 11) = "Date"
    Worksheets("Resource Log").Cells(1, 38 + MaxStatuses + (5 * MaxStatuses) + 12) = "Time"

    X = 1
    Z = 2

    For X = 2 To NumRows
    
    
        If Worksheets("Resource Log").Cells(X, 12).Value <> "" Then
        
            Y = 1
            NumStatuses = Worksheets("Resource Log").Cells(X, 38).Value
            
            For Y = 1 To NumStatuses
            
                Worksheets("Resource Log").Cells(Z, 38 + MaxStatuses + (5 * MaxStatuses) + 4) = Worksheets("Resource Log").Cells(X, 38 + MaxStatuses + ((5 * Y) - 4)).Value
                Worksheets("Resource Log").Cells(Z, 38 + MaxStatuses + (5 * MaxStatuses) + 5) = Worksheets("Resource Log").Cells(X, 2).Value
                Worksheets("Resource Log").Cells(Z, 38 + MaxStatuses + (5 * MaxStatuses) + 6) = Worksheets("Resource Log").Cells(X, 13).Value
                Worksheets("Resource Log").Cells(Z, 38 + MaxStatuses + (5 * MaxStatuses) + 7) = Worksheets("Resource Log").Cells(X, 15).Value
                Worksheets("Resource Log").Cells(Z, 38 + MaxStatuses + (5 * MaxStatuses) + 8) = Worksheets("Resource Log").Cells(X, 26).Value
                Worksheets("Resource Log").Cells(Z, 38 + MaxStatuses + (5 * MaxStatuses) + 9) = Worksheets("Resource Log").Cells(X, 38 + MaxStatuses + ((5 * Y) - 3)).Value
                Worksheets("Resource Log").Cells(Z, 38 + MaxStatuses + (5 * MaxStatuses) + 10) = Worksheets("Resource Log").Cells(X, 38 + MaxStatuses + ((5 * Y) - 2)).Value
                Worksheets("Resource Log").Cells(Z, 38 + MaxStatuses + (5 * MaxStatuses) + 11) = Worksheets("Resource Log").Cells(X, 38 + MaxStatuses + ((5 * Y) - 1)).Value
                Worksheets("Resource Log").Cells(Z, 38 + MaxStatuses + (5 * MaxStatuses) + 12) = Worksheets("Resource Log").Cells(X, 38 + MaxStatuses + ((5 * Y))).Value
            
                Z = Z + 1
            
            Next
           
         End If
    
    Next

    X = 2
    P = 38 + MaxStatuses + (5 * MaxStatuses) + 4
    NumRows = ActiveWorkbook.Worksheets("Resource Log").Range(Cells(1, P), Cells(1, P).End(xlDown)).Rows.Count
    
    For X = 2 To NumRows
    
        mystring = Left(Cells(X, 38 + MaxStatuses + (5 * MaxStatuses) + 4).Text, 1)
    
        If mystring = " " Then
        
            Cells(X, 38 + MaxStatuses + (5 * MaxStatuses) + 4).Value = Right(Cells(X, 38 + MaxStatuses + (5 * MaxStatuses) + 4), Len(Cells(X, 38 + MaxStatuses + (5 * MaxStatuses) + 4)) - 1)
        
        End If

    Next

    Set src = Worksheets("Resource Log").Cells(1, 38 + MaxStatuses + (5 * MaxStatuses) + 4).CurrentRegion
    Worksheets("Resource Log").ListObjects.Add(SourceType:=xlSrcRange, Source:=src, xlListObjectHasHeaders:=xlYes, tablestyleName:="TableStyleMedium28").Name = "Timesheet_Table"

    Sheets("Dashboard").Activate
    Sheets("Dashboard").Cells(3, 6) = Now
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox ("Resource Log Successfully Updated")
    Exit Sub

ErrorHandler:
    Sheets("Dashboard").Activate
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox ("There Is An Error On Row: " & X & " of the Datebase. Please Delete This Row on the EE Data Sheet and run the Macro Again") ',vbOKOnly,"DATABASE ERROR",,)
    Exit Sub

End Sub

Sub CollectUniquesNow()

    Dim rngUniques As Range, rngTarget As Range
    Dim varUniques As Variant
    Dim lngIdx As Long
    Dim NumRows As Long
    Dim colUniques As Collection
    Set colUniques = New Collection
    Dim rngTarget2 As String
    Dim MaxStatuses As Integer
    Dim P As Integer
    Dim src As Range

    Application.ScreenUpdating = False
    Sheets("Resource Log").Activate
    
    
    MaxStatuses = Application.WorksheetFunction.Max(Range("Resource_Table[Functions]"))
    P = 38 + MaxStatuses + (5 * MaxStatuses) + 4
    NumRows = ActiveWorkbook.Worksheets("Resource Log").Range(Cells(2, P), Cells(2, P).End(xlDown)).Rows.Count

    Set rngTarget = Worksheets("Resource Log").Range(Cells(2, 38 + MaxStatuses + (5 * MaxStatuses) + 4), Worksheets("Resource Log").Cells(NumRows, 38 + MaxStatuses + (5 * MaxStatuses) + 4))

    'Collect the uniques using the function we just wrote
    Set colUniques = CollectUniques(rngTarget)
    
    'Load a Variant array with the uniques
    '(in preparation for writing them to a new sheet)
    ReDim varUniques(colUniques.Count, 1)
    For lngIdx = 1 To colUniques.Count
        varUniques(lngIdx - 1, 0) = CStr(colUniques(lngIdx))
    Next lngIdx
    
    Sheets("Uniques").Activate
    Sheets("Uniques").Cells.ClearContents
    
    Worksheets("Uniques").Cells(1, 1).Value = "Employee Names"
    
    Set rngUniques = Worksheets("Uniques").Range("A2:A" & colUniques.Count + 1)
    rngUniques = varUniques

    Set src = Worksheets("Uniques").Range("A1").CurrentRegion
    Worksheets("Uniques").ListObjects.Add(SourceType:=xlSrcRange, Source:=src, xlListObjectHasHeaders:=xlYes, tablestyleName:="TableStyleLight8").Name = "Employee_Table"

    With Worksheets("Uniques").ListObjects("Employee_Table").Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("Employee_Table[Employee Names]"), SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
    
    
    Sheets("Resource Log").Activate

    Set rngTarget = Worksheets("Resource Log").Range(Cells(2, 38 + MaxStatuses + (5 * MaxStatuses) + 5), Worksheets("Resource Log").Cells(NumRows, 38 + MaxStatuses + (5 * MaxStatuses) + 5))

    'Collect the uniques using the function
    Set colUniques = CollectUniques(rngTarget)
    
    'Load a Variant array with the uniques
    '(in preparation for writing them to a new sheet)
    ReDim varUniques(colUniques.Count, 1)
    For lngIdx = 1 To colUniques.Count
        varUniques(lngIdx - 1, 0) = CStr(colUniques(lngIdx))
    Next lngIdx

    Sheets("Uniques").Activate

    Worksheets("Uniques").Cells(1, 3).Value = "Job Codes"
    
    Set rngUniques = Worksheets("Uniques").Range("C2:C" & colUniques.Count + 1)
    rngUniques = varUniques

    Set src = Worksheets("Uniques").Range("C1").CurrentRegion
    Worksheets("Uniques").ListObjects.Add(SourceType:=xlSrcRange, Source:=src, xlListObjectHasHeaders:=xlYes, tablestyleName:="TableStyleLight8").Name = "JobCode_Table"
    
    With Worksheets("Uniques").ListObjects("JobCode_Table").Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("JobCode_Table[Job Codes]"), SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
    
    
    Sheets("Resource Log").Activate

    Set rngTarget = Worksheets("Resource Log").Range(Cells(2, 14), Worksheets("Resource Log").Cells(NumRows, 14))

    'Collect the uniques using the function
    Set colUniques = CollectUniques(rngTarget)
    
    'Load a Variant array with the uniques
    '(in preparation for writing them to a new sheet)
    ReDim varUniques(colUniques.Count, 1)
    For lngIdx = 1 To colUniques.Count
        varUniques(lngIdx - 1, 0) = CStr(colUniques(lngIdx))
    Next lngIdx

    Sheets("Uniques").Activate

    Worksheets("Uniques").Cells(1, 5).Value = "Categories"
    
    Set rngUniques = Worksheets("Uniques").Range("E2:E" & colUniques.Count + 1)
    rngUniques = varUniques

    Set src = Worksheets("Uniques").Range("E1").CurrentRegion
    Worksheets("Uniques").ListObjects.Add(SourceType:=xlSrcRange, Source:=src, xlListObjectHasHeaders:=xlYes, tablestyleName:="TableStyleLight8").Name = "Category_Table"
    
    With Worksheets("Uniques").ListObjects("Category_Table").Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("Category_Table[Categories]"), SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With


    Sheets("Dashboard").Activate
    Sheets("Dashboard").Cells(4, 6) = Now
    Application.ScreenUpdating = True
    MsgBox ("All Unique Values SuccessFully Identified")

End Sub

Sub GenerateBudgetTimeReport()

With BudgetTimeReport
  .StartUpPosition = 0
  .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
  .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
  .Show

End With

End Sub

Sub GenerateEmployeeReport()

With EmployeeReport
  .StartUpPosition = 0
  .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
  .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
  .Show

End With

End Sub

Sub GenerateComponentReport()

With ComponentReport
  .StartUpPosition = 0
  .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
  .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
  .Show

End With

End Sub
Sub GenerateAssemblyReport()

With AssemblyChecker
  .StartUpPosition = 0
  .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
  .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
  .Show

End With

End Sub
Sub GenerateProjectReport()

With ProjectReport
  .StartUpPosition = 0
  .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
  .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
  .Show

End With

End Sub
Sub UpdateBudgeTimes()

Application.ScreenUpdating = False

With BudgetTimeUpdate
  .StartUpPosition = 0
  .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
  .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
  .Show

End With

End Sub
Sub CreateJobSysLabourCSV()

Application.ScreenUpdating = False

With JobSysLabour
  .StartUpPosition = 0
  .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
  .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
  .Show
End With

End Sub

Sub CreateQCReport()

Application.ScreenUpdating = False

With QCPassedReport
  .StartUpPosition = 0
  .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
  .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
  .Show
End With

End Sub



Sub WorkInProgress()

MsgBox ("This Function Is Not Yet Available")

End Sub
