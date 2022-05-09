VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Transform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' This is a class (a collection of subs and functions) which is responsible for transforming the worksheet from one format into another
' -- TODO: save data on create_report click and restore that data on error
' -- TODO: (double reset bug) Data(1,2) was causing an error when there wasn't a row 2.
'

' This variable stores the report data in a "Report" object we have created. See Report.cls to see what it does.
Private sheet_report As Report

'
' We use this to access the sheet_report object.
'
Property Get Report() As Report
    ' Return whatever is stored in the report variable.
    Set Report = sheet_report
End Property

'
' We use this to set the sheet_report object
'
Property Let Report(value As Report)
    ' Set the sheet_report variable to whatever we passed in
    Set sheet_report = value
End Property

' This sub is run every time a new Transform object is created.
' -- Its job is to set up our object.
' -- TODO: check the current sheet format and keep track of it
'
Private Sub class_initialize()
    ' Disable updating the sheet until we're done to speed things up.
    Application.ScreenUpdating = False
  
    ' Create a new report object and store it in our Report property (which represents the variable sheet_report)
    Set sheet_report = New Report
    ' Initialize the report with a refernce to this object
    Report.Init Me
End Sub

' Run initialization code with parameters. For some reason, VBA doesn't allow you to pass parameters to class_initialize... so we have to do it this way
Public Sub Init()
  ' Empty for now. We just create this as a template if we need to use it later.
End Sub

'
' This sub is run every time the Transform object is destroyed.
' -- Its job is to clean up once we're done running our code.
' -- This is one of the main reasons we're using objects; to ensure this code runs every time no matter what.
'
Private Sub Class_Terminate()
    ' Print a debug message for us (user doesn't see this)
    ' -- Ctrl+G in the code viewer will show these messages
    Debug.Print "Terminating the Transform object and performing cleanup"
    
    ' Clear our clipboard. This should not be strictly necessary
    ' -- but we're doing it anyway just to be safe
    Application.CutCopyMode = False
    
    ' Turn updating back on
    Application.ScreenUpdating = True
End Sub

'
' This function gets a cell BELOW the header. We outsource this work to the Report class, but we define it here just to make references
' -- to it shorter. For example: Data(1,1) instead of Report.Data(1,1)
'
' -- For example, if our header begins at B2, then:
' -- Data(1,1) would return Range(B3)
' -- Data(3,3) would return Range(D4)
'
Public Function Data(column As Long, row As Long) As Range
        ' Let our report object handle this logic for us
        Set Data = Report.Data(column, row)
End Function

'
' Duplicates all active rows below the header
' -- returns the number of rows we inserted
'
Public Function Duplicate_Rows() As Long
    ' Start below the header
    Dim startRow As Long: startRow = Report.header.row + 1
    ' i is our current row number. It changes as we loop through each row.
    Dim i As Long
    ' count is the number of rows we inserted. It increases each new operation
    Dim count As Long: count = 0
    ' Count the number of rows in our sheet which we need to duplicate
    Dim lastRow As Long: lastRow = Report.last_row()

    ' Print a debug message for us (user doesn't see this)
    ' -- Ctrl+G in the code viewer will show these messages
    Debug.Print "Duplicating rows from " & startRow & " to " & lastRow

    ' Loop through each row from bottom to top, stopping at "startRow"
    For i = lastRow To startRow Step -1
        ' Set the transaction type here.
        ' -- TODO: This shouldn't be in this function, but this is the easiest to do right now
        Cells(i, Report.header.column).value = "3"
        
        ' Copy the row contents
        Rows(i).Copy
        ' Paste the row contents into a newly inserted row
        Rows(i).Insert
        
        ' Modify the new row's transaction type
        ' -- TODO: This shouldn't be in this function, but this is the easiest to do right now
        Cells(i + 1, Report.header.column).value = "4"
        
        ' Keep track of the total of rows we inserted
        count = count + 1
    Next

    ' Print a debug message for us (user doesn't see this)
    ' -- Ctrl+G in the code viewer will show these messages
    Debug.Print "Duplicated " & count & " rows."

    ' Turn off copy mode, since we're done copying
    Application.CutCopyMode = False
    
    ' Return the number of rows we inserted, in case we need that one day
    ' -- right now we do not use this information
    DuplicateRows = count
End Function

'
' Switch the headers to the "report" view. This happens after clicking the "create report" button
' -- TODO: we have this in a sub because eventually we will copy these records from a template page, so it is easier for a non-programmer to change. But we aren't doing that now.
'
Private Sub Report_Headers()
    ' Print a debug message for us (user doesn't see this)
    ' -- Ctrl+G in the code viewer will show these messages
    Debug.Print "Switching to Report Headers"

    ' The starting column of our header. Stored here just to make things below look nicer.
    Dim startColumn As Long: startColumn = Report.header.column

    ' Clear all the comments in our header
    Report.Full_Range.ClearComments
    Report.Header_Range.ClearComments
    
    ' Insert a new column at the start of the data (this keeps data in the correct associated column)
    Columns(startColumn).Insert xlShiftToRight, xlFormatFromRightOrBelow
    ' Insert 2 columns at column 3:4 of the data
    Columns(startColumn + 3).Insert xlShiftToRight, xlFormatFromRightOrBelow
    Columns(startColumn + 3).Insert xlShiftToRight, xlFormatFromRightOrBelow
    
    ' Set the headers for each row
    ' -- This Range is equivalent to Range("B4:I4")
    Range(Data(0, 1), Data(7, 1)).value = Array("Transaction", "NYSLRS ID", "Empl Rcd", "Last 4 of social", "First 3 of first name", "Social Security Number", "First Name", "Last Name")

    ' Reset the header data in our Report class
    Report.find_header

    ' Define a special error handler here to skip adding comments if we encounter an error while doing so
    On Error GoTo error_handler

    ' Add a few comments to explain our data
    With Data(1, 1).AddComment
        .Visible = False
        .Text "Transaction type. Each employee needs at least one type 3 and one type 4 line."
    End With
    With Data(3, 1).AddComment
        .Visible = False
        .Text "Employee Record. This is 0 unless the employee has had another job in the retirement system."
    End With
    
    ' Exit before our error handler code. If we don't do this, our error handler code will run even on non-errors
    Exit Sub

    ' Our error handler
error_handler:
        ' Let us know there was a problem
        Debug.Print "Error adding comment. This usually happens because a previous comment wasn't cleared: #" & Err.Number & " " & Err.Description

End Sub

'
' Switch the headers to our initial view. This happens when the sheet is first loaded, or it is "reset"
' -- TODO: we have this in a sub because eventually we will copy these records from a template page, so it is easier for a non-programmer to change. But we aren't doing that now.
'
Private Sub Initial_Headers()
    ' Print a debug message for us (user doesn't see this)
    ' -- Ctrl+G in the code viewer will show these messages
    Debug.Print "Switching back to our initial headers."
    
    ' -- This Range is equivalent to Range("B4:I4")
    Range(Data(1, 1), Data(8, 1)).value = Array("NYSLRS ID", "Empl Rcd", "Social Security Number", "First Name", "Last Name", "", "", "")

    ' Reset the header data in our Report class
    Report.find_header

    ' Define a special error handler here to skip adding comments if we encounter an error while doing so
    On Error GoTo error_handler

    ' Add a few comments to explain our data
    With Data(2, 1).AddComment
        .Visible = False
        .Text "Employee Record. This is 0 unless the employee has had another job in the retirement system."
    End With

    ' Exit before our error handler code. If we don't do this, our error handler code will run even on non-errors
    Exit Sub

    ' Our error handler code
error_handler:
        ' Let us know there was a problem
        Debug.Print "Error adding comment. This usually happens because a previous comment wasn't cleared: #" & Err.Number & " " & Err.Description

End Sub

'
' This function loops through each row with a social and copies the last 4 to the "Last 4 of Social" column
'
Public Function redact_socials()
    ' Start looking for socials below the header row
    Dim startRow As Long: startRow = Report.header.row + 1
    ' i is our current row number. It changes as we loop through each row.
    Dim i As Long
    ' count is the number of rows we inserted. It increases each new operation
    Dim count As Long: count = 0
    ' Count the number of rows in our sheet which we need to duplicate
        Dim lastRow As Long: lastRow = Report.last_row()
    ' RegExp is an object that allows us to interact with "Regular Expressions", which is a language that searches text.
    ' -- we are using it to grab the last 4 digits from the social
    ' -- to use this, make sure the Regular Expressions 5.5 references are turned on in "Tools > References"
    Dim reg As RegExp
    ' Another regular expression to validate the social we're working with
    Dim validate As RegExp
    ' This variable will hold the matches we find when we run the regex
    Dim Matches As Object
    ' If the social is valid. This will change at each row.
    Dim is_valid As Boolean
    ' A variable to hold the full social of the row we're working with. This will change as we loop over each row
    Dim social As String

    ' Print a debug message for us (user doesn't see this)
    ' -- Ctrl+G in the code viewer will show these messages
    Debug.Print "Redacting socials from " & startRow & " to " & lastRow

    ' Set up our Regex options
    Set reg = New RegExp
    ' This pattern searches for digits (\d) in a set of 4 ({4}) followed immediately by the end of the cell ($).
        ' -- The \s* represents optional spaces which we ignore
    reg.Pattern = "\d{4}\s*$"

    ' Loop through each row from bottom to top, stopping at "startRow"
    For i = lastRow To startRow Step -1
        ' A debug line we don't use right now
        'Debug.Print "Checking social on " & i&; ":" & Report.Header.column + 5

        ' Get the full social from the current row (i) and from column 6 ("social security number")
        social = Cells(i, Report.header.column + 5).value
        
        ' Find matches within the cell value
        Set Matches = reg.Execute(social)
        
        ' If we didn't find any matches, the data is wrong
        If Matches.count < 1 Then
            ' Print a debug message for us (user doesn't see this)
            ' -- Ctrl+G in the code viewer will show these messages
            Debug.Print "Social passed validation but no matches found: " & social
            
            ' Let the user know something went wrong
            MsgBox "An error occurred processing the social on line " & i
            
            ' Move on to the next row. This one can't be processed.
            GoTo NextIteration
        End If
        
        ' Excel returns an array (a list) of results. We Can grab index 0 or loop through the array... so let's loop
        ' -- Advantages of looping occur when we have < 1 match (we avoid an error referencing index 0) or > 1 match (we process each one, so can print debug messages and maybe catch the issue)
        For Each m In Matches
            ' Set the "Last 4 of social" cell to what we found
            ' -- The cell is the current row (i) and column 4 ("Last 4 of social")
            ' -- TODO: Validate integrity (what does VBA return on failure?)
                Cells(i, Report.header.column + 3).value = m.value
        Next m
        
        ' Keep track of the total of rows we modified
        count = count + 1
        
        ' Label this section "NextIteration" so we can reference it with GoTo statements earlier.
        ' -- this is how VB handles control statements like "continue" and "break"
NextIteration:
    Next i

    ' Print a debug message for us (user doesn't see this)
    ' -- Ctrl+G in the code viewer will show these messages
    Debug.Print "Redacted " & count & " socials"
    
    ' Return the number of rows we modified, in case we need that one day
    ' -- right now we do not use this information
    redact_socials = count
    
End Function

'
' Validates all data before conversion
'
Public Function validate_data() As Boolean
    
    ' Our Validation class handles all the validation work for us
    Dim validator As Validation
    Set validator = New Validation
    validator.Report = Report

    ' If it told us that it failed, then communicate that to the function that called this one
    If validator.validate_data = False Then
        validate_data = False
        ' Exit here before we handle "success" logic below
        Exit Function
    End If

        
    ' If we got here, then we passed all validation
    ' -- Let the program that called this one know that we succeeded validating all entries
    validate_data = True
        
End Function

'
' This function copies the first 3 letters of the first name to the appropriate column
'
Public Function shorten_names()
    ' Start below the header
    Dim startRow As Long: startRow = Report.header.row + 1
    ' i is our current row number. It changes as we loop through each row.
    Dim i As Long
    ' count is the number of rows we inserted. It increases each new operation
    Dim count As Long: count = 0
    ' Count the number of rows in our sheet which we need to duplicate
        Dim lastRow As Long: lastRow = Report.last_row()
    ' RegExp is an object that allows us to interact with "Regular Expressions", which is a language that searches text.
    ' -- we are using it to grab the last 4 digits from the social
    ' -- to use this, make sure the Regular Expressions 5.5 references are turned on in "Tools > References"
    Dim reg As RegExp
    ' This variable will hold the matches we find when we run the regex
    Dim Matches As Object
    ' A variable to hold the first name of the row we're working with. This will change as we loop over each row
    Dim first_name As String

    ' Print a debug message for us (user doesn't see this)
    ' -- Ctrl+G in the code viewer will show these messages
    Debug.Print "Redacting socials from " & startRow & " to " & lastRow

    ' Set up our Regex options
    Set reg = New RegExp
    Set validate = New RegExp
    ' This pattern searches for the first 1 to 3 non-numeric characters
    ' -- It is important that we be as lenient as possible here, to account for non-latin characters in a user's name
    ' -- as well as names that are fewer than 3 characters (like "Al")
    reg.Pattern = "^\s*\D\D?\D?"

    ' Loop through each row from bottom to top, stopping at "startRow"
    For i = lastRow To startRow Step -1
        ' Get the full first name from the current row (i) and from column 7 ("first name")
        first_name = Cells(i, Report.header.column + 6).value
        
        ' Find matches within the cell value
        Set Matches = reg.Execute(first_name)
        
        ' If we didn't find any matches, the data is wrong
        If Matches.count < 1 Then
            ' Print a debug message for us (user doesn't see this)
            ' -- Ctrl+G in the code viewer will show these messages
            Debug.Print "Unable to get the first 3 characters of the first name: " & first_name
            
            ' Let the user know something went wrong
            MsgBox "An error occurred processing the first name on line " & i
            
            ' Move on to the next row. This one can't be processed.
            GoTo NextIteration
        End If
        
        ' Excel returns an array (a list) of results. We Can grab index 0 or loop through the array... so let's loop
        ' -- Advantages of looping occur when we have < 1 match (we avoid an error referencing index 0) or > 1 match (we process each one, so can print debug messages and maybe catch the issue)
        For Each m In Matches
            ' Set the "Last 4 of social" cell to what we found
            ' -- The cell is the current row (i) and column 5 ("First 3 of first name")
            ' -- TODO: Validate integrity (what does VBA return on failure?)
            Cells(i, Report.header.column + 4).value = m.value
        Next m
        
        ' Keep track of the total of rows we modified
        count = count + 1
        
        ' Label this section "NextIteration" so we can reference it with GoTo statements earlier.
        ' -- this is how VB handles control statements like "continue" and "break"
NextIteration:
    Next i

    ' Print a debug message for us (user doesn't see this)
    ' -- Ctrl+G in the code viewer will show these messages
    Debug.Print "Shortened " & count & " first names"
    
    ' Return the number of rows we modified, in case we need that one day
    ' -- right now we do not use this information
    shorten_names = count
End Function

'/*
'
' This sub sorts a report based on the lastname, first name from A-Z, then employment record
'
'*/
Public Sub Sort_Report()
    ' Turn off autofilters
    ActiveSheet.AutoFilterMode = False

    ' Sort the Active Sheet only
    With ActiveSheet.Sort
        ' Clear any previous sorting
        .SortFields.Clear
        ' Sort by "Last Name" A-Z
        .SortFields.Add key:=Columns(Report.header.column + 7), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ' Sort by "First Name" A-Z
        .SortFields.Add key:=Columns(Report.header.column + 6), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ' Sort by "Employee ID" 0-9
        .SortFields.Add key:=Columns(Report.header.column + 2), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ' Set the range to the header row to end, the header column start to finish
        .SetRange Range(Report.header, Data(8, 1).End(xlDown))
        ' The first row (A4:H4) is a header
        .header = xlYes
        ' Set orientation (probably unnecessary)
        .Orientation = xlTopToBottom
        ' Apply (run) the sort
        .Apply
    End With
End Sub

'/*
'
' Reset the page to its initial state
'
'*/
Public Sub Resetpage()
    ' Print a message for us. The user doesn't see this. Click Ctrl+G in the macro editor to see these messages.
    Debug.Print "Resetting the page."

    ' If we don't have a header at all, then the page is broken and needs to be reset the page completely piece by piece
    If Report.Has_Header = False Then
        ' Print a message for us. The user doesn't see this. Click Ctrl+G in the macro editor to see these messages.
        Debug.Print "Unable to find a header during reset."
        ' Another function handles all the logic for recreating the page
        Recreate_Page
        ' All done. Page should be "reset"
        Exit Sub
    End If
    'Range(Range("B6"), Range("F6")) = Array("NYSLRS ID", "Empl Rcd", "Social Security Number", "First Name", "Last Name")

    ' Save the report state for later. We're going to be making changes to the file which will impact what mode our Report class thinks we're in
    Dim isReport As Boolean: isReport = Report.isReport

    ' Clear all the comments in our header
    Report.Header_Range.ClearComments

    ' Select everything in the report and remove any text we find there
    Report.Full_Range().value = ""

    ' If we have a report, then we need to delete a few rows we inserted when we created it
    If isReport = True Then
        ' Print a message for us. The user doesn't see this. Click Ctrl+G in the macro editor to see these messages.
        Dim col As Long: col = Report.header.column

        ' Delete column 1, 4 and 5. Col+3 = 4, then once it is deleted, Col+3 will be filled by column 5.
        Columns(col + 3).EntireColumn.Delete
        Columns(col + 3).EntireColumn.Delete
        ' TODO: This should be Columns(col) but that is throwing an "object required" error. I'm not sure why that is
        ' -- So we're going to delete the wrong column and call it a day
        Columns(col + 1).EntireColumn.Delete
    End If

    ' Set our new headers
    Initial_Headers
    
    ' Put in sample data
    ' TODO Temporary
    Range(Data(1, 2), Data(5, 2)).value = Array("R10450174", "1", "687485841", "Reanne", "Dulanski")
    Range(Data(1, 3), Data(5, 3)).value = Array("R10450174", "2", "687485841", "Reanne", "Dulanski")
    Range(Data(1, 4), Data(5, 4)).value = Array("R10450174", "3", "687485841", "Reanne", "Dulanski")
    Range(Data(1, 5), Data(5, 5)).value = Array("R10450173", "1", "687-48-5852", "John", "Smith")
    Range(Data(1, 6), Data(5, 6)).value = Array("R 10450172", "2", "687 48 5863 ", "Jo-Anne", "Jacobs ")
    Range(Data(1, 7), Data(5, 7)).value = Array("r-10450173", "55", " 687 - 48 - 5852", "Ali' Baba", "Bush Jr 2nd")
End Sub

'/*
'
' Recreate the page piece by piece in case it gets messed up.
'
' This is going to be a hacky solution for now. This shouldn't be called hardly ever and if it fails the sheet can be redownloaded.
' -- it should only be used if the USER changes the sheet in a way that breaks it.
'
' TODO: Set up borders/backgrounds/etc
'
'*/
Public Function Recreate_Page() As Boolean
    ' Print a message for us. The user doesn't see this.
    Debug.Print "Recreating the page."

    ' Clear all data in every cell
    ' -- todo: use End() like a civilized person
    Range("B6:ZZ999999").value = ""

    ' Set new headers where they are supposed to be as of this writing
    Range("B6:F6").value = Array("NYSLRS ID", "Empl Rcd", "Social Security Number", "First Name", "Last Name")

End Function

'
' The new Create Report Button
' -- This will take an existing set of employee records and replace the sheet with that data
' -- formatted in the way our other excel tool needs.
'
Public Sub Create_Report()
    ' If we encounter an error, run our error handler to notify the user and fail (or continue) gracefully
    On Error GoTo error_handler

    ' Check if the sheet is in the "old" (reset) format, or the new (create report) format
    If Report.isInitial() = False Then
        If Report.isReport() Then
            ' The sheet appears to be in the new format
            ' Tell our user what happened
            MsgBox "The report was already generated. To create a new report, click 'Reset Page' and enter new data."
        Else
            ' The sheet appears to be in an unknown format
            ' Tell the user what happened
            MsgBox "Something is wrong with the sheet. Please click 'Reset Page' and insert your data again."
        End If

        ' Stop creating the report.
        Exit Sub
    End If

    ' If we have no data other than the headers, then we have nothing to do
    ' -- TODO: UsedRange is quick, but it counts rows with only spaces.
    If ActiveSheet.UsedRange.Rows.count < Report.header.row + 1 Then
        ' Tell our user what happened
        MsgBox "There is no data to report on. Please fill in some data and click the button again."
        ' Stop creating the report.
        Exit Sub
    End If
    
    ' Validate each row of data before proceeding with changes
    ' -- TODO: Doing it in one loop means formatting errors will prevent finding duplicates
    If validate_data() = False Then
        ' The data isn't valid, so don't continue changing anything
        Exit Sub
    End If

    ' Switch our headers to the report view, using a sub we defined above
    Report_Headers

    ' Copy the last 4 of the social
    redact_socials
    
    ' Shorten the first names
    shorten_names

    ' Sort the table by name/employee id
    Sort_Report

    ' Duplicate each row after our header
    Duplicate_Rows
    
    ' End the sub code here, because we don't want to execute our error handler code below unless it is called.
    Exit Sub

'
' This code is referenced when we encounter an error we didn't expect
' -- TODO: Check the error code and give more specific messages
'
error_handler:
    ' If an error is encountered in THIS section, then terminate the program.
    On Error GoTo 0

    ' Let our user know what happened
    MsgBox "An error was encountered. Please contact the transition team if you continue to encounter this message. " & Err.Description
    ' Log information about the error for us (users don't see this)
    ' -- Ctrl+G in the code viewer will show these messages
    Debug.Print "Error: (#" & Err.Number & ") " & Err.Description

    ' Try to continue
    Resume Next

End Sub
