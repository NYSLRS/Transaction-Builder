Attribute VB_Name = "Ribbon"

'Static training_mode_test As Boolean
Public training_mode As Boolean

'/*
'
' This sub turns on instructor mode when the "Training Mode" button in the Ribbon is clicked
'
'*/
Sub Instructor_Mode(control As IRibbonControl)
    if training_mode <> True Then
        ' Tell the training class to turn on training mode
        training_mode = True
        ' Print a message for us. The user doesn't see this.
        Debug.Print "Instruction Mode Enabled"
        ' Print a message for the user.
        MsgBox "Instruction Mode Enabled"
    End If
End Sub
