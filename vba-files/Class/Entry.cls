VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'/*
'
' The entry class represents a single entry (cell) in the sheet.
' -- It expands the default Cell class provided by VBA with extra functionality
'
'*/

' The cell this references. We access this by calling Entry.cell (which is defined as a property below)
Private cell_ref As Range
Private errors() As comment

'
' We use this to access the cell_ref property
'
Property Get cell() As Range
    ' Return whatever is stored in the cell property.
    ' -- this should always be set
    Set cell = cell_ref
End Property

'
' We use this to set the cell this object is referencing
'
Property Let cell(value As Range)
    ' Set the variable to whatever we passed in
    Set cell_ref = value
End Property

'/*
'
' Show an error message next to the cell which can be dismissed
'
'*/
Public Function error(message As String) As Boolean
    ' The comment we create
    Dim comment As comment
    ' Create a new comment
    Set comment = cell.AddComment message
    ' Show it by default
    comment.Visible = True

    ' Add it to our error comment array
    errors.Add comment

    ' Print a message for us. The user doesn't see this.
    Debug.Print cell.Address & " data error: " & message

    error = True
End Function

'/*
'
' Similar to Range.ClearComments, but only deletes the error comments we added and not other comments explaining the field
'
'*/
Public Function Clear_Errors() As Integer
    ' A running count of how many comments we removed
    Dim count As Integer: count = 0
    ' The current comment we're working with
    Dim comment As comment
    ' Loop through each comment and remove it
    For Each comment In errors
        ' Keep track of how many we removed
        count = count + 1
        ' Remove it
        comment.Delete
    ' Process the next comment
    Next comment

    ' Report back on how many we removed
    Clear_Errors = count
End Function

'/*
'
' Creates a new entry object, referencing the provided cell, and returns it.
'
' In any modern language, the constructor would do this. But VBA doesn't allow that.
'
'*/
Public Static Function Factory(cell As Range) As entry
    ' The entry object we're creating
    Dim entry As entry: Set entry = New entry
    ' Set the cell
    entry.cell = cell

    ' Return it
    Factory = entry
End Function
