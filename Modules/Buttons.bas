'
' The Sub we call when the "create report" button is pressed.
' -- outsources this work to the Transform class
'
Sub Create_Report()
    ' Create a new Transform object
    Dim trans As Transform
    Set trans = New Transform
    ' Call the "Create_Report" sub in the transform class
    ' -- This handles the job for us
    trans.Create_Report
End Sub

Sub Resetpage()
    ' Create a new Transform object
    Dim trans As Transform
    Set trans = New Transform
    ' Call the "Resetpage" sub in the transform class
    ' -- This handles the job for us
    trans.Resetpage
End Sub
