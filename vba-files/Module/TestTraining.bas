Attribute VB_Name = "TestTraining"
'@TestModule("Training")
'@Folder("Tests.Training")
'@ModuleDescription "Tests the Training.cls class"

Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object
Private Training As Training
Private Transform As Transform

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
    
    Set Training = New Training
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Training: SSN")