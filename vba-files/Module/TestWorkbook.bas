Attribute VB_Name = "TestWorkbook"
'@TestModule("ThisWorkbook")
'@Folder("Tests.Workbook")
'@ModuleDescription "Tests the ThisWorkbook.cls class"

Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object
Private Workbook As Workbook
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
    
    Set Workbook = New Workbook
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Workbook: SSN")