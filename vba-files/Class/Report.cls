VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' This is a class (a collection of subs and functions) which is responsible for telling us things about the Report sheet, such as whether it
' -- has been generated, what row the report starts on, and so on.
'

' This variable stores the beginning of the header for the sheet. We determine this procedurally in case
' -- the sheet gets changed in a way we don't expect in the future.
' -- You can access this property using the Get/Let methods below, e.g. transform.Header = Range("A1")
Private Header_Start_Cell As Range

' A reference to the Transform object we're using. See Transform.cls to see what this object does.
Private transform_instance As Transform

'
' We use this to access the header.
'
Property Get header() As Range
    ' Return whatever is stored in the header_start_cell variable.
    ' -- this should always be set, because it is determined in Class_Initialize
    Set header = Header_Start_Cell
End Property

'
' We use this to set the header
'
Property Let header(value As Range)
    ' Set the header variable to whatever we passed in
    Set Header_Start_Cell = value
End Property

'
' We use this to access the transform object.
'
Property Get Transform() As Transform
    ' Return whatever is stored in the header_start_cell variable.
    ' -- this should always be set, because it is determined in Class_Initialize
    Set Transform = transform_instance
End Property

'
' We use this to set the transform object
'
Property Let Transform(value As Transform)
    ' Set the variable to whatever we passed in
    Set transform_instance = value
End Property
    
' This sub is run every time a new Report object is created.
' -- Its job is to set up our object.
' -- TODO: check the current sheet format and keep track of it
'
Private Sub class_initialize()
    ' Empty for now. We just create this as a template if we need to use it later.
End Sub

' Run initialization code with parameters. For some reason, VBA doesn't allow you to pass parameters to class_initialize... so we have to do it this way
Public Sub Init(Transform As Transform)
    ' Set the transform object to what we passed in
    Me.Transform = Transform

    ' Find the header and store it for later
    find_header
End Sub

'
' This function determines what header columns we have and what they should be called.
' -- if we change the sheet, we need to update the columns here.
'
Public Function column_data() As Scripting.Dictionary
    ' This stores a reference to each of our columns and the data it contains
    ' -- TODO: This should be handled when we initialize our Report object
    Dim Data As Scripting.Dictionary
    Set Data = New Scripting.Dictionary

    ' This is sloppy, but it works for right now. We hardcode these references to the column names and return those
    If isInitial = True Then
        Data.Add 1, "NYSLRS ID"
        Data.Add 2, "Employee Record"
        Data.Add 3, "SSN"
        Data.Add 4, "First Name"
        Data.Add 5, "Last Name"
    Else
        Data.Add 1, "Transaction"
        Data.Add 2, "NYSLRS ID"
        Data.Add 3, "Employee Record"
        Data.Add 4, "Last 4 of SSN"
        Data.Add 5, "First 3 of Name"
        Data.Add 6, "SSN"
        Data.Add 7, "First Name"
        Data.Add 8, "Last Name"
    End If

    ' Let the function that called this one know what we determined
    Set column_data = Data
End Function

'
' We use this to determine the first row of the report (which is currently the line below the header)
'
Property Get report_start_row() As Long
    ' Tell the function that called this one what we determined the row to be.
    report_start_row = header.row + 1
End Property


'
' This function gets a cell BELOW the header.
' -- For example, if our header begins at B2, then:
' -- Data(1,1) would return Range(B3)
' -- Data(3,3) would return Range(D4)
'
Public Function Data(column As Long, row As Long)
    ' Grab the cell at Header + Parameter and set our return value to it
    Set Data = Cells(row + header.row - 1, column + header.column - 1)
End Function

'
' Finds the column number for the provided column name
' -- TODO: Make this less dependent on using the exact string in column_data
'
Public Function column(name As String) As Integer
    ' Get our column data from a function we defined above
    Dim dict As Scripting.Dictionary
    Set dict = column_data()
    ' A variable to hold the key for the entry we're working with. This will change as the program runs.
    Dim key As Variant
    ' Loop through each entry in the column_data dictionary
    For Each key In dict.keys
        ' If the key matches the name we're looking for, success!
        If dict.Item(key) = name Then
            ' Let the function that called this one know that we found a result, and move that result over based on the header position
            column = CInt(key) + header.column - 1
            ' Exit here, because we found what we're looking for
            Exit Function
        End If
    Next
End Function


'/*
'
' This function searches the top of the sheet until we find the first column of our header
'
'*/
Public Function find_header(Optional recreated As Boolean = False) As Range
    ' A variable to store the cell we're checking. This changes over time
    Dim cell As Range

    ' Search for NYSLRS ID in the first 50 rows of the first A-M columns
    Set cell = Range("A1:M50").Find("Transaction")
    ' If we didn't find it, try something else
    If (cell Is Nothing) Then
        ' Try searching for "Transaction" instead
        Set cell = Range("A1:M50").Find("NYSLRS ID")
        ' We didn't find that either! There's a problem.
        If (cell Is Nothing) Then
            ' Unable to find a header!

            If recreated = False Then
                ' -- try to reset the page
                Transform.Recreate_Page

                ' Try running it again
                ' -- True is important here. It tells us that we ran the recreate function already.
                ' -- -- if we exclude it, we'll end up in an infinite loop.
                Set find_header = Me.find_header(True)
                ' Return whatever that function call returned and exit. It did everything we want to do already.
                Exit Function
            End If

            ' STILL no header... something is really wrong. Bail here.

            ' Print a message to us (the user doesn't see this)
            Debug.Print "Unable to find header"
            ' Send a message to the user
            MsgBox "An error was encountered reading the sheet."
            ' Throw an error, which tells our program to stop running (or handle it)
            Err.Raise Number:=vbObjectError + 513, Description:="Unable to find header"
        End If
    End If
    
    ' Print a message for us (the user doesn't see this). Press Ctrl+G to see these messages in the Macro editor
    Debug.Print "Found our header at " & cell.Address

    ' Reset our local header variable
    Set Header_Start_Cell = cell

    ' Let the function that called this one know what we found
    Set find_header = cell

End Function

'/*
'
' Determines if there is a header present on this page. Does not notify the user or do anything with that information.
'
' This is only used to determine if the page is broken so we can run code to fix it.
'
'*/
Public Function Has_Header() As Boolean
    ' A variable to store the cell we're checking. This changes over time
    Dim cell As Range

    ' Search for NYSLRS ID in the first 50 rows of the first A-M columns
    Set cell = Range("A1:M50").Find("Transaction")
    ' If we didn't find it, try something else
    If (cell Is Nothing) Then
        ' Try searching for "Transaction" instead
        Set cell = Range("A1:M50").Find("NYSLRS ID")
        ' We didn't find that either! There's a problem.
        If (cell Is Nothing) Then
            Has_Header = False
            Exit Function
        End If
    End If
    
    Has_Header = True
End Function

'
' Returns the number of the last row in this sheet. Turning this into a function because I'm not sure which is the most efficient/reliable way to do this, and I may want to swap
' -- the method out in the future.
'
Public Function last_row() As Long
    ' This is one approach. However, it counts blanks and if the macro is called multiple times, the number of rows it duplicates doubles each time, eventually locking the sheet.
    ' last_row = ActiveSheet.UsedRange.Row - 1 + ActiveSheet.UsedRange.Rows.count

    ' The next approach depends on a column we know will always have data in it. We will use "first name"
    Dim column As Integer
    If isInitial() Then
        ' In the initial state, the "First name" column is 4 (1+3 = 4)
        column = header.column + 3
    Else
        ' In the "report" state, the "first name" column is 7 (1+6 = 7)
        column = header.column + 6
    End If

    ' Use this approach for now. It finds the last cell in column 7 or 4 (first name) with data, and counts to that.
    last_row = Cells(Rows.count, column).End(xlUp).row
End Function

'/*
'
' Returns the number of the last column in this sheet. In the intial state, this will be something like column F (6). In the report state, it will be something like column I (9)
'
'*/
Public Function last_column() As Long
    ' Search the Header row from end to start until we find a cell with data. Then return the column there.
    last_column = Cells(header.row, Columns.count).End(xlToLeft).column
End Function
'/*
'
' Returns the range of cells containing the header
'
'*/
Public Function Header_Range() As Range
    ' A range from the start of the header to the last cell in the header
    Set Header_Range = Range(header.Address, Cells(header.row, last_column))
End Function

'/*
'
' Returns a range of cells representing this entire report
'
'*/
Public Function Full_Range() As Range
    ' Get the range from the header (1,1) to the last column and last row (8,last_row())
    Set Full_Range = Range(Data(1, 1), Data(last_column(), last_row()))
End Function

'
' Determines if the sheet is in its initial state
' -- this happens when it is first loaded, and when the "reset" button is clicked
'
Public Function isInitial() As Boolean
    
    ' If the first cell of the header is our initial value, then assume the rest of the sheet is in its initial state
    If Data(1, 1).value = "NYSLRS ID" Then
        ' Return True when this function ends
        isInitial = True
    Else
        ' Return False when this function ends
        isInitial = False
    End If
End Function

'
' Determines if the sheet is in the "report" format.
' -- this happens when the "create report" button is clicked
'
Public Function isReport() As Boolean
    ' If the first cell of the header is in our "report" state, then assume the rest of the sheet is in the same state too
    If Data(1, 1).value = "Transaction" Then
        ' Return True when this function ends
        isReport = True
    Else
        ' Return False when this function ends
        isReport = False
    End If
End Function
