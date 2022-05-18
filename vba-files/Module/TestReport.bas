Attribute VB_Name = "TestReport"
'@TestModule("Report")
'@Folder("Tests.Report")
'@ModuleDescription "Tests the Report.cls class"

Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object
Private report As Report
Private sheet As Worksheet

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
    Set sheet = ThisWorkbook.Worksheets.Add
    Set report = New Report
    report.switchSheet(sheet)
    SetupHeaders()
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    sheet.Delete
    Set report = Nothing
End Sub

Private Sub SetupHeaders() 
    sheet.Range("A1").value = "NYSLRS ID"
    sheet.Range("A2").value = "Employee Record"
    sheet.Range("A3").value = "SSN"
    sheet.Range("A4").value = "First Name"
    sheet.Range("A5").value = "Last Name"
End Sub

'@TestMethod("Report: SSN")
Public Sub Test_Has_Header()
    Assert.IsTrue report.HasHeader
End Sub