VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Validation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/*
'
' This class is responsible for validating (and formatting) our data
' -- TODO: Dup social/employee #
' -- TODO: Dup social/dif names
' -- TODO: 12345 social
'
'*/

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
Public Property Let Report(value As Report)
    ' Set the sheet_report variable to whatever we passed in
    Set sheet_report = value
End Property

' This sub is run every time a new Validation object is created.
' -- Its job is to set up our object.
'
Private Sub class_initialize()
End Sub

' Run initialization code with parameters. For some reason, VBA doesn't allow you to pass parameters to class_initialize... so we have to do it this way
Public Sub Init()
  ' Empty for now. We just create this as a template if we need to use it later.
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

'/*
'
' This function validates the NYSLRS ID.
'   If the ID is invalid, it returns False
'   If the ID is valid, it formats it the way we expect it to be and returns a string
'
'*/
Public Function validate_id(value As String) As Variant

  ' The validation regex
  Dim reg As RegExp
  Set reg = New RegExp
  ' The validation regex pattern to use.
  '   ^       The start of the value
  '   \s*     Any number of spaces (including 0)
  '   [Rr]    The character "R" or "r", which start every ID
  '   [\s-]*  Any number of spaces or dashes (including 0)
  '   [012]   One single number which is either "0", "1", or "2". This starts every ID with plenty of room for the future.
  '   \d{7}   A series of 7 numbers
  '   \s*     Any number of spaces (including 0)
  '   $       The end of the value
  '   Parenthesis simply define a group (these will be called "SubMatches" below). They don't impact what gets matched.
  ' So this will match any 8 digit number starting with R
  reg.Pattern = "^\s*[rR][\s-]*([012]\d{7})\s*$"
  
  Dim Matches As Object: Set Matches = reg.Execute(value)
  
  ' Run the regex to test our value
  If Matches.count < 1 Then
      ' Let the program that called this one know that we failed validation
      validate_id = False
      
      ' Print a message for us (the user doesn't see that)
      Debug.Print "NYSLRS ID failed validation: " & value
      
      ' And we're all done.
      Exit Function
  End If

  ' Perform a validation check before we continue. This should never happen now, but might happen in the future
  ' -- if our code changes.
  If Matches(0).SubMatches.count < 1 Then
    ' Print a message for us. The user doesn't see this.
    Debug.Print "No submatches found in NYSLRS ID. This should never happen."
    
    ' Let the program that called this one know that we passed validation
    validate_id = True
    
    ' We can't continue formatting the social, so end here
    Exit Function
  End If

  ' This will store the new (formatted) social
  Dim format As String
  ' Set it to each digit we found without the stuff in between
  format = "R" & Matches(0).SubMatches(0)

  ' We should now have a 9 digit "number"
  If Len(format) <> 9 Then
      Debug.Print "Formatted NYSLRS ID is not the right number of digits: " & value
      GoTo Format_Error
  End If

  ' Let the program that called this one know that we passed validation
  ' -- and give it the formatted social so it can update the cell
  validate_id = format
  Exit Function

' Define a special error handler just for this function
Format_Error:
  ' Let us (not the user) know there was an issue
  Debug.Print "An error was encountered formatting the NYSLRS ID: " & value
  ' Validation has passed, but the id can't be formatted
  validate_id = True
End Function

'/*
'
' This function validates the Employee Record.
'   If the ID is invalid, it returns False
'   If the ID is valid, it formats it the way we expect it to be and returns a string
'
'*/
Public Function validate_employee_record(value As String) As Variant
  ' The validation regex
  Dim reg As RegExp
  Set reg = New RegExp
  ' The validation regex pattern to use.
  '   ^       The start of the value
  '   \s*     Any number of spaces (including 0)
  '   \d      A single number (0-9)
  '   \d?\d?  Two more optional numbers
  '   \s*     Any number of spaces (including 0)
  '   $       The end of the value
  ' So this will match any number between 0-999
  reg.Pattern = "^\s*(\d\d?\d?)\s*$"
  
  ' Run the regex to test our value
  If reg.test(value) = False Then
      ' Let the program that called this one know that we failed validation
      validate_employee_record = False
  Else
      ' Let the program that called this one know that we passed validation
      validate_employee_record = True
  End If
End Function

'/*
'
' This function validates the Social Security Number.
'   If the ID is invalid, it returns False
'   If the ID is valid, it formats it the way we expect it to be and returns a string
'
'*/
Public Function validate_social(value As String) As Variant
  ' The validation regex
  Dim reg As RegExp
  Set reg = New RegExp
  ' This pattern matches for:
  '   ^     - Start of the text
  '   \s*   - Any number of spaces
  '   0*    - Any number of preceeding 0s
  '   \d{3} - 3 digits
  '   -*    - An optional dash which may or may not be there
  '   \d{2} - 2 digits
  '   -*    - An optional dash which may or may not be there
  '   \d{4} - 4 digits
  '   $     - The end of the text
  '   Parenthesis simply define a group (these will be called "SubMatches" below). They don't impact what gets matched.
  ' So this will match "123-45-6789" as well as "123456789" and " 123 - 45 - 6789 " and various combinations in between
  reg.Pattern = "^\s*0*\s*(\d{3})\s*-*\s*(\d{2})\s*-*\s*(\d{4})\s*$"

  Dim Matches As Object: Set Matches = reg.Execute(value)
  
  ' Run the regex to test our value
  If Matches.count < 1 Then
      ' Let the program that called this one know that we failed validation
      validate_social = False
      
      ' Print a message for us (the user doesn't see that)
      Debug.Print "Social failed validation: " & value
      
      ' And we're all done.
      Exit Function
  End If

  '/*
  ' Format the SSN the way we expect it to be (no dashes/spaces/etc)
  '*/

  ' Handle errors differently here
  On Error GoTo Format_Error

  ' Perform a validation check before we continue. This should never happen now, but might happen in the future
  ' -- if our code changes.
  If Matches(0).SubMatches.count < 3 Then
      ' Print a message for us. The user doesn't see this.
      Debug.Print "No submatches found in social. This should never happen."
      
      ' Let the program that called this one know that we passed validation
      validate_social = True
      
      ' We can't continue formatting the social, so end here
      Exit Function
  End If

  ' This will store the new (formatted) social
  Dim format As String
  ' Set it to each digit we found without the stuff in between
  format = Matches(0).SubMatches(0) & Matches(0).SubMatches(1) & Matches(0).SubMatches(2)

  ' We should now have a 9 digit number
  If Len(format) <> 9 Then
      Debug.Print "Formatted social is not the right number of digits for a SSN: " & value
      GoTo Format_Error
  End If

  ' Let the program that called this one know that we passed validation
  ' -- and give it the formatted social so it can update the cell
  validate_social = format

  Exit Function

' Define a special error handler just for this function
Format_Error:
  ' Let us (not the user) know there was an issue
  Debug.Print "An error was encountered formatting the social: " & value
  ' Validation has passed, but the social can't be formatted
  validate_social = True
End Function

Public Function validate_first_name(value As String) As Boolean
  ' The validation regex
  Dim reg As RegExp
  Set reg = New RegExp
  ' The validation regex pattern to use.
  '   ^       The start of the value
  '   \s*     Any number of spaces (including 0)
  '   [^\s\d] A character which is NOT a space or a number (a last name can't begin with a number)
  '   .*      Any number of characters (including numbers and spaces) We want this to be as lenient as possible,
  '           --- because names are weird ("Jo Anne", "Jo-Anne", "Al", accents, single quotes, etc)
  '   \s*     Any number of spaces (including 0)
  '   Parenthesis simply define a group (these will be called "SubMatches" below). They don't impact what gets matched.
  ' So this will match "Joanne", "Jo-Anne", "Bob Sr 2nd", "Ali'Baba", and non-latin characters in names (such as those with accents)
  reg.Pattern = "^\s*([^\s\d].*)\s*"
  
  ' Run the regex to test our value
  If reg.test(value) = False Then
      ' Let the program that called this one know that we failed validation
      validate_first_name = False
  Else
      ' Let the program that called this one know that we passed validation
      validate_first_name = True
  End If
End Function

  '/*
  '
  ' This function validates a last name
  '
  '*/
Public Function validate_last_name(value As String) As Boolean
  ' The validation regex
  Dim reg As RegExp
  Set reg = New RegExp
  ' The validation regex pattern to use.
  '   ^       The start of the value
  '   \s*     Any number of spaces (including 0)
  '   [^\s\d] A character which is NOT a space or a number (a last name can't begin with a number)
  '   .*      Any number of characters (including numbers and spaces) We want this to be as lenient as possible,
  '           --- because names are weird ("Jo Anne", "Jo-Anne", "Al", accents, single quotes, etc)
  '   \s*     Any number of spaces (including 0)
  '   Parenthesis simply define a group (these will be called "SubMatches" below). They don't impact what gets matched.
  ' So this will match "Joanne", "Jo-Anne", "Bob Sr 2nd", "Ali'Baba", and non-latin characters in names (such as those with accents)
  reg.Pattern = "^\s*([^\s\d].*)\s*"
  
  ' Run the regex to test our value
  If reg.test(value) = False Then
      ' Let the program that called this one know that we failed validation
      validate_last_name = False
  Else
      ' Let the program that called this one know that we passed validation
      validate_last_name = True
  End If
End Function

'/*
'
' This function takes a row/column, grabs the value from that cell, and validates it based on the "type" supplied
'
' If the cell is invalid, we return False
' If the cell is valid, we format it the way we expect it to be presented and return True
'
'*/
Public Function validate_cell(row As Long, column As Integer, row_type As String) As Boolean
  ' Get the value of the cell
  Dim value As String: value = Cells(row, column).value
  ' The result from our validation function (used later). This is either false or a string
  Dim result As Variant
  
  ' Debug.Print "Validating cell " & row & ":" & column & ", " & row_type & ": " & value

  ' Depending on the type, run the appropriate validation function
  Select Case row_type
    ' Social Security Number
    Case "social", "ssn", "SSN"
      result = validate_social(value)
    ' NYSLRS ID
    Case "id", "nyslrsid", "nyslrs_id", "NYSLRS ID"
      result = validate_id(value)
    ' Employee Record
    Case "emprec", "employee record", "employee_record", "Employee Record"
      result = validate_employee_record(value)
    ' First Name
    Case "first", "first name", "first_name", "First Name"
      result = validate_first_name(value)
    ' Last Name
    Case "last", "last name", "last_name", "Last Name"
      result = validate_last_name(value)
    ' All other types:
    Case Else
      ' Print a message for us (the user doesn't see this)
      Debug.Print "validate_cell a bad type was supplied: " & row_type
      ' We weren't able to validate anything, so raise an error (which will be handled by the error handler in the calling function)
      Err.Raise Number:=vbObjectError + 513, Description:="Bad cell type: " & row_type
      ' In case our error handler calls a "Resume Next", we want to communicate that validation failed and exit
      ' -- we won't arrive at this point in the code otherwise.
      validate_cell = False
      Exit Function
  End Select
  
  ' If Validation failed (or succeeded without a new formatted value) then we're all done
  If result = False Or (result = True And VarType(result) <> VarType("String")) Then
    ' Tell the calling function that we failed (or succeeded without changes)
    validate_cell = result
    ' End our job here and don't do anything else
    Exit Function
  End If

  ' Otherwise... our validation function passed a string back for us to update our value
  Cells(row, column).value = result

  ' Validation must have passed to get here. Let the function that called this one know.
  validate_cell = True

End Function

'/*
'
' This function is responsible for finding (and notifying the user about) duplicate SSN's without a different employment #
' -- The cell parameter is a single cell with a SSN in it which we want to check all other rows for
'
'*/
Public Function find_duplicate_socials(cell As Range) As Boolean
  ' The last row with data
  Dim last_row As Long: last_row = Report.last_row
  ' The first row of data (1 below the header)
  Dim start_row As Integer: start_row = Report.header.row + 1
  ' The SSN column
  Dim ssn_col As Integer: ssn_col = Report.column("SSN")
  ' The Employment # column
  Dim empl_num_col As Integer: empl_num_col = Report.column("Employee Record")
  ' A reference to the row we're currently looking at. This will change as the program runs.
  Dim row As Long
  ' The values we retrieve. This will change as the program runs
  Dim ssn As String
  Dim empnum1 As String
  Dim empnum2 As String

  ' Iterate over each row in the sheet
  For row = start_row To last_row
    ' If the rows are the same, we don't need to compare them
    If cell.row <> row Then
      ' If the SSN for both rows match (and are not empty):
      ssn = Cells(row, ssn_col).value
      ' A debug line we don't use right now
      'Debug.Print "SSN: " & cell.value & "/" & ssn

      If cell.value <> "" And cell.value = ssn Then

        ' Find both employee numbers
        empnum1 = Cells(cell.row, empl_num_col).value
        empnum2 = Cells(row, empl_num_col).value

        ' A debug line we don't use right now
        'Debug.Print "Empnum: " & empnum1 & "/" & empnum2
        
        ' If the employment numbers for both rows ALSO match (and are not empty)
        If empnum1 <> "" And empnum1 = empnum2 Then
          ' Let the user know that they have duplicate data
          MsgBox "Row " & row & " and " & cell.row & " have the same SSN and employment #"
          ' Let the function that called this one know that we found duplicates
          find_duplicate_socials = True
          ' We don't need to run any more. Finding one is enough
          Exit Function
        End If
      End If
    End If
  Next row

  ' Let the function that called this one know that we didn't find any duplicates
  find_duplicate_socials = False
End Function

  '
' Validates all data before conversion
'
Public Function validate_data() As Boolean
    ' Start below the header
    Dim startRow As Long: startRow = Report.header.row + 1
    ' i is our current row number. It changes as we loop through each row.
    ' TODO: the row number we present to the user needs to be + header row
    Dim i As Long
    ' We will use this later to store the column number we are working with
    Dim col As Variant
    Dim column_name As String
    Dim column As String
    ' Count the number of rows in our sheet which we need to work with
    Dim lastRow As Long: lastRow = Report.last_row()
    ' A cell we're working with. This will change as the program runs
    Dim cell As Range
    
    ' This stores a reference to each of our columns and the data it contains
    ' -- TODO: This should be handled when we initialize our Transform object
    Dim column_data As Scripting.Dictionary
    Set column_data = Report.column_data

    ' Loop through each row from bottom to top, stopping at "startRow"
    For i = lastRow To startRow Step -1

        ' Loop through each column and run our validation code on it
        For Each col In column_data.keys
            column_name = column_data(col)
            ' Set a reference to the cell we're working with. This may be helpful later. CInt(col) is the Integer representation of the column. i is the row (which is already an integer)
            Set cell = Cells(i, CInt(col))

            ' Call the validation function with row number (i), column number as an integer (CInt(col)), and the column name (column_data(col))
            If validate_cell(i, CInt(col) + Report.header.column - 1, column_name) = False Then
                ' Let the user know what happened
                MsgBox column_name & " for row " & i & " is incorrect. Please check it and try again."
                ' Let the function that called this one know that we found an error
                validate_data = False
                ' We don't need to run anymore. One error is enough
                Exit Function
            End If

        Next col

        
        ' Find duplicate data (comparing to this row: i)
        If find_duplicate_data(i) = True Then
            ' No need to let the user know. Our find_duplicate_data function did that
            ' -- Let the function that called this one know that we failed validation
            validate_data = False
            ' No need to run anymore. One error is enough.
            Exit Function
        End If

        ' Find bad SSNs (like 123456789), comparing to this row: i
        If find_bad_ssn(i) = True Then
            ' No need to let the user know. Our find_duplicate_socials function did that
            ' -- Let the function that called this one know that we failed validation
            validate_data = False
            ' No need to run anymore. One error is enough.
            Exit Function
        End If
        
    ' Move on to the next row
    Next i
        
    ' If we got here, then we passed all validation
    ' -- Let the program that called this one know that we succeeded validating all entries
    validate_data = True
        
End Function


' Finds any duplicate data where we don't allow duplicates
'
Public Function find_duplicate_data(row As Long) As Boolean
    ' Get the cell in this row and the "SSN" column
    Dim cell As Range: Set cell = Cells(row, Report.column("SSN"))

    ' Find duplicate data
    If find_duplicate_socials(cell) = True Then
        ' No need to let the user know. Our find_duplicate_socials function did that
        ' -- Let the function that called this one know that we failed validation
        find_duplicate_data = True
        ' No need to run anymore. One error is enough.
        Exit Function
    End If
        
    ' If we got here, then we passed all validation
    ' -- Let the program that called this one know that we succeeded validating all entries
    find_duplicate_data = False
        
End Function

'/*
'
' This function finds any bad SSNs so we don't run into an error on submission to the website.
' -- for instance, "123456789"
'
'*/
Public Function find_bad_ssn(row As Long)
  ' Get the cell in this row and the "SSN" column
  Dim cell As Range: Set cell = Cells(row, Report.column("SSN"))

  ' If it's 123456789, the website won't take it.
  If cell.value = "123456789" Then
    Dim result As Integer
    ' Let the user know
    result = MsgBox("It looks like you used dummy data for row " & cell.row & ". Unfortunately, the website will not accept this number. Do you want me to create a sample SSN for you?", vbYesNo + vbQuestion)

    ' If they asked for a sample social, create one for them
    If result = vbYes Then
      ' The sample we'll create
      Dim sample As String
      ' This is a hacky solution, but it works for now.
      ' -- Prefix the social by a "9" so we avoid 1234... nonsense. Real socials cannot begin with a 9, but the website accepts them
      ' -- then use the row number to avoid collisions with other sample socials.
      ' -- -- Right("789" & cell.row, 4) prefixes the row number by 789 to make sure it is at least 4 characters long, then takes the right-most 4 digits
      ' -- Then add the row number to the end (prefixed by zeroes) to ensure no collisions
      sample = "9" & Right("789" & cell.row, 4) & Right("000" & cell.row, 4)

      ' Set the cell to this value we created
      cell.value = sample

      ' Now the SSN is valid. So we succeeded
      find_bad_ssn = False
      Exit Function

    ' If they said no, then bail out here because we have a bad SSN and can't continue
    End If

    ' Let the function calling this one know that we failed
    find_bad_ssn = True
    ' Exit here before our success logic below
    Exit Function
  End If

  ' If we got here, then we didn't find an error. Let the function that called this one know.
  find_bad_ssn = False
End Function
