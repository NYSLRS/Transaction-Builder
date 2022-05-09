VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Allow sorting of rows when you double click on the header
' -- TODO: Allow asc/desc sorting
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' Set up some variables for later
    Dim KeyRange As Range
    Dim ColumnCount As Integer
    
    ' Determine the number of columns
    ColumnCount = Range("A4:H4").Columns.count
    ' By default, Cancel should be false
    Cancel = False
    ' If the cell we clicked on is within the defined header
    If Target.row = 4 And Target.column <= ColumnCount Then
        ' Cancel the default functionality
        Cancel = True
        ' Print a message for us (the user doesn't see this)
        Debug.Print "Sorting column " & Target.Address
        ' Sort the entire sheet (A:H, row 4 to end)
        Range("A4", Range("H4").End(xlDown)).Sort Key1:=Range(Target.Address), header:=xlYes
    End If
End Sub
