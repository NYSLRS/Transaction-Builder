Attribute VB_Name = "Buttons"
'/*
' The Sub we call when the "create report" button is pressed.
' -- outsources this work to the Transform class
'*/
Sub Create_Report()
    ' Create a new Transform object
    Dim trans As Transform
    Set trans = New Transform
    ' Call the "Create_Report" sub in the transform class
    ' -- This handles the job for us
    trans.Create_Report
End Sub

'/*
'
' This macro is called when the "reset" button is clicked on the Transaction Builder tool
'
'*/
Sub Resetpage()
    ' Create a new Transform object
    Dim trans As Transform
    Set trans = New Transform
    ' Call the "Resetpage" sub in the transform class
    ' -- This handles the job for us
    trans.Resetpage
End Sub

'/*
'
' This macro is called when the "import txt" button is clicked
'
'*/
Sub Import_Txt()
    ' Create a new Transcode object
    Dim trans As Transcode
    Set trans = New Transcode
    ' Call the "Import" sub in the transcode class
    ' -- This handles the job for us
    trans.Import
End Sub

'/*
' Checks if a provided cell has Data Validation rules
' -- this is used to find the columns which use values from our Options sheet and make them look pretty
'*/
Function HasValidation(cell As Range) As Boolean
    ' The type of data validation to be used later.
    ' -- Initialize this to null so if it is not overwritten later, we know we didn't find anything.
    Dim t: t = Null

    ' If an error is encountered, ignore it
    On Error Resume Next
    ' Get the type of the validation on the cell
    t = cell.Validation.Type
    'MsgBox "Found validation type: " & t
    ' Reset the error handler to "exit"
    On Error GoTo 0

    ' Tell the function that called this one what we found
    ' -- if the "t" variable is not set to null, then we found something in the datavalidation property
    HasValidation = Not IsNull(t)
End Function
