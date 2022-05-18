Attribute VB_Name = "TestReport"
'@TestModule("Report")
'@Folder("Tests.Report")
'@ModuleDescription "Tests the Report.cls class"

Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object
Private report As Report
Private s As Worksheet
Private trans As Transform

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    
    Delete_Test_Sheets

    Create_Sheet
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing

    'Delete_Sheet s
End Sub

'@TestInitialize
Private Sub TestInitialize()
    ' Sometimes this gets reset in the middle of tests. It shouldn't... but VBA.
    If s Is Nothing Then 
        Create_Sheet
    End If

    ClearSheet

    'This method runs before every test in the module..
    SetupHeaders

    Set trans = New Transform
    ' Initialize the transform object
    trans.Init s

    Set report = trans.report
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    'ClearSheet
    Set report = Nothing
End Sub

Private Sub Create_Sheet()
    Set s = ThisWorkbook.Worksheets.Add
    s.Name = "Test WS " & Format(Now, "mmddhhmmss")
End Sub

Private Sub SetupHeaders( Optional rng As Range = Nothing )
    If rng Is Nothing Then
        Set rng = s.Range("A1:E1")
    End If 
    rng.value = Array("NYSLRS ID", "Employee Record", "SSN", "First Name", "Last Name")

    ' Reset the report header (if report is available)
    if Not report Is Nothing Then
        report.find_header
    End If
End Sub

Private Function CustomHeaders( Optional start_at As String = "C5" )
    Dim start_cell As Range: Set start_cell = s.Range(start_at)
    Dim end_cell As Range: Set end_cell = start_cell.Offset(, 4)
    Dim rng As Range: Set rng = s.Range(start_cell, end_cell)

    ClearSheet
    SetupHeaders rng

    CustomHeaders = rng
End Function

Private Sub ClearSheet()
    s.Range("A1:ZZ99999").value = ""
End Sub

Sub Delete_Sheet( ws As Worksheet )
    Application.DisplayAlerts = False 
    ws.Delete
    Application.DisplayAlerts = True 
End Sub

Sub Delete_Test_Sheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.name, 8) = "Test WS " Then
            Debug.Print "Deleting test worksheet " & ws.name
            Delete_Sheet ws 
        End If
    Next ws
End Sub

'@TestMethod("Report: Headers")
Public Sub Test_Has_Header()
    Assert.IsTrue report.Has_Header, "Unable to find initial header"
End Sub

'@TestMethod("Report: Headers")
Public Sub Test_Has_Header_Custom()  
    ClearSheet
    SetupHeaders s.Range("C5:G5")

    Assert.IsTrue report.Has_Header, "Unable to find custom header"
End Sub

'@TestMethod("Report: Headers")
Public Sub Test_Find_Header()
    Dim cell As Range: Set cell = report.find_header
    Assert.AreEqual cell.Address, "$A$1", "Cannot find initial header at A1"
End Sub

'@TestMethod("Report: Headers")
Public Sub Test_Find_Header_Custom()
    ClearSheet
    SetupHeaders s.Range("C5:G5")
    Dim cell As Range: Set cell = report.find_header
    Assert.AreEqual cell.Address, "$C$5", "Cannot find custom header at C5"
End Sub

'@TestMethod("Report: Data")
Public Sub Test_Last_Row()
    Dim last_row As Long: last_row = report.last_row
    Assert.IsTrue last_row = 1

    s.Range("A2:E2").value = "Test"
    last_row = report.last_row
    Assert.IsTrue last_row = 2, "Last row wrong with row 2 set"
    Debug.Print "last row is " & last_row

    s.Range("A3:E3").value = "Test"
    last_row = report.last_row
    Assert.IsTrue last_row = 3, "Last row wrong after 2nd row addition"

    ' Try individual columns
    s.Range("A4").value = "Test"
    last_row = report.last_row
    Assert.IsTrue last_row = 4, "Last row wrong when setting only column A"
    s.Range("B5").value = "Test"
    last_row = report.last_row
    Assert.IsTrue last_row = 5, "Last row wrong when setting only column B"
    s.Range("C6").value = "Test"
    last_row = report.last_row
    Assert.IsTrue last_row = 6, "Last row wrong when setting only column C"
    s.Range("D7").value = "Test"
    last_row = report.last_row
    Assert.IsTrue last_row = 7, "Last row wrong when setting only column D"
    s.Range("E8").value = "Test"
    last_row = report.last_row
    Assert.IsTrue last_row = 8, "Last row wrong when setting only column E"
End Sub

'@TestMethod("Report: Data")
Public Sub Test_Last_Row_Custom_Header()
    ClearSheet
    SetupHeaders s.Range("C5:G5")

    Dim last_row As Long: last_row = report.last_row
    Assert.IsTrue last_row = 5, "Last row wrong with custom header @ 5"

    s.Range("C6:G6").value = "Test"
    last_row = report.last_row
    Assert.IsTrue last_row = 6, "Last row wrong with custom header @ 6"
    s.Range("C7:G7").value = "Test"
    last_row = report.last_row
    Assert.IsTrue last_row = 7, "Last row wrong with custom header @ 7"

    ' Skip a line
    s.Range("C9:G9").value = "Test"
    last_row = report.last_row
    Assert.IsTrue last_row = 9, "Last row wrong with custom header after skipping line"

    ' Try another column
    s.Range("C10").value = "Test"
    last_row = report.last_row
    Assert.IsTrue last_row = 10, "Last row wrong when setting only column 1"
    s.Range("D11").value = "Test"
    last_row = report.last_row
    Assert.IsTrue last_row = 11, "Last row wrong when setting only column 2"
    s.Range("E12").value = "Test"
    last_row = report.last_row
    Assert.IsTrue last_row = 12, "Last row wrong when setting only column 3"
    s.Range("F13").value = "Test"
    last_row = report.last_row
    Assert.IsTrue last_row = 13, "Last row wrong when setting only column 4"
    s.Range("G14").value = "Test"
    last_row = report.last_row
    Assert.IsTrue last_row = 14, "Last row wrong when setting only column 5"
End Sub

'@TestMethod("Report: Headers")
Public Sub Test_Last_Column()
    Dim last_column As Long: last_column = report.last_column
    Assert.IsTrue last_column = 5  

    s.Range("F1").value = "Test"
    last_column = report.last_column
    Assert.IsTrue last_column = 6
    s.Range("G1").value = "Test"
    last_column = report.last_column
    Assert.IsTrue last_column = 7
End Sub

'@TestMethod("Report: Headers")
Public Sub Test_IsInitial() 
    ' Default sheet
    Assert.IsTrue report.IsInitial, "Report says it is not initial"

    ' Blank sheet
    ClearSheet
    Assert.IsFalse report.IsInitial, "Blank Report says it IS initial"

    ' Custom header
    SetupHeaders s.Range("C5:G5")
    Assert.IsTrue report.IsInitial, "Report says it is not initial on custom header"
End Sub

'@TestMethod("Report: Data")
Public Sub Test_Data() 
    Dim cell As Range

    Set cell = report.Data(1,1)
    Debug.Print "A2 ?= " & cell.Address
    Assert.IsTrue cell.Address = "$A$2", "Report says data(1,1) does not start on B2"
    Set cell = report.Data(2,2)
    Debug.Print "C3 ?= " & cell.Address
    Assert.IsTrue cell.Address = "$B$3", "Report says data(2,2) is not at C3"
End Sub

'@TestMethod("Report: Data")
Public Sub Test_Data_Custom_Header() 
    ClearSheet
    SetupHeaders s.Range("C5:G5")

    Dim cell As Range

    Set cell = report.Data(1,1)
    Debug.Print "D6 ?= " & cell.Address
    Assert.IsTrue cell.Address = "$C$6", "Report says data(1,1) for custom header does not start on D6"
    Set cell = report.Data(2,2)
    Debug.Print "D7 ?= " & cell.Address
    Assert.IsTrue cell.Address = "$D$7", "Report says data(2,2) for custom header is not at D7"
End Sub

'@TestMethod("Report: Data")
Public Sub Test_Data_Row()
    Dim row As Range
    Dim first_cell As Range

    Set row = Report.Data_Row(2)
    Assert.IsTrue row.count = 5, "Data_Row is not the right width: " & row.count
    Set first_cell = row.Item(1)
    Assert.IsTrue first_cell.Address = "$B$1", "Data_Row first cell is not correct: " & first_cell.Address

    CustomHeaders("C5")

    Set row = report.Data_Row(5)
    Assert.IsTrue row.count = 5, "Data_Row with custom header is not the right width: " & row.count
    Set first_cell = row.Item(1)
    Assert.IsTrue first_cell.Address = "$G$1", "Data_Row with custom header, first cell is not correct: " & first_cell.Address
End Sub

'@TestMethod("Report: Data")
Public Sub Test_Data_Column() 
    Dim col As Range
    Dim first_cell As Range
    Set col = report.Data_Column(2)
    Assert.IsTrue col.count = 1, "Data_Column is not the right height: " & col.count
    Set first_cell = col.Item(1)
    Assert.IsTrue first_cell.Address = "$B$1", "Data_Column first cell is not correct: " & first_cell.Address

    CustomHeaders("C5")

    Set col = report.Data_Column(5)
    Assert.IsTrue col.count = 1, "Data_Column with custom header is not the right width: " & col.count
    Set first_cell = col.Item(1)
    Assert.IsTrue first_cell.Address = "$G$1", "Data_Column with custom header, first cell is not correct: " & first_cell.Address
End Sub

'@TestMethod("Report: Data")
Public Sub Test_Data_Range() 
    Dim first_cell As Range
    Dim rng As Range: Set rng = report.Data_Range(2, 3, 5)
    Assert.IsTrue rng.count = 5, "Data_Range is not the right width: " & rng.count
    Set first_cell = rng.Item(1)
    Assert.IsTrue first_cell.Address = "$B$4", "Data_Range first cell is not correct: " & first_cell.Address

    CustomHeaders("C5")

    Set rng = report.Data_Range(5, 7, 3)
    Assert.IsTrue rng.count = 3, "Data_Range for custom header is not the right width: " & rng.count
    Set first_cell = rng.Item(1)
    Assert.IsTrue first_cell.Address = "$G$12", "Data_Range for custom header, first cell is not correct: " & first_cell.Address
End Sub

    
'@TestMethod("Report: Properties")
Public Sub Test_Switch_Sheet()
    Dim test_sheet As Worksheet
    Set test_sheet = ThisWorkbook.Worksheets.Add
    Dim name as String: name = "Test WS Switching " & Format(Now, "mmddhhmmss")
    test_sheet.name = name

    report.SwitchSheet test_sheet
    Assert.AreEqual report.sheet.name, name, "Switching sheet, names are not equal"
End Sub 

Public Function Check_Cell_Value( column As Long, row As Long, expected_value As String ) As Boolean
    Dim value As String: value = s.Cells(row, column).value
    Check_Cell_Value = (value = expected_value)
End Function

Public Function Check_Header_Value( name As String ) As Boolean
    Dim col As Long: col = report.column(name)
    Check_Header_Value = Check_Cell_Value(col, report.header.row, name)
End Function


'@TestMethod("Report: Header")
Public Sub Test_Get_Column_Data()
    Dim data As Scripting.Dictionary: Set data = report.column_data()
    Assert.IsTrue data.count = 5, "column_data returned <> 5 columns for no params: " & data.count
    Set data = report.column_data("Current")
    Assert.IsTrue data.count = 5, "column_data returned <> 5 columns for Current: " & data.count
    Set data = report.column_data("Report")
    Assert.IsTrue data.count = 8, "column_data returned <> 8 columns for a report: " & data.count
    Set data = report.column_data("Initial")
    Assert.IsTrue data.count = 5, "column_data returned <> 5 columns for initial: " & data.count
End Sub
        
'@TestMethod("Report: Properties")
Public Sub Test_Get_Column()
    Dim columns() As Variant: columns = Array("NYSLRS ID", "Employee Record", "SSN", "First Name", "Last Name")
    Dim col

    For Each col In columns
        Assert.IsTrue Check_Header_Value(CStr(col)), col & " Column could not be found"
    Next col
End Sub

'@TestMethod("Report: Header")
Public Sub Test_Report_Start_Row()
    Assert.IsTrue report.report_start_row = 2, "Report_start_row not 2 after initial state"

    CustomHeaders("C5")

    Assert.IsTrue report.report_start_row = 6, "Report_start_row not 6 after custom headers"
End Sub

'@TestMethod("Report: Header")
Public Sub Test_Header_Range()
    Dim rng As Range: Set rng = report.Header_Range
    Assert.IsTrue rng.count = 5, "Initial header_range is not equal to 5: " & rng.count
    Assert.IsTrue rng.Address = "$A$1", "Initial header_range doesn't start at A1: " & rng.Address

    ' Add to it
    s.Range("F1:G1").Value = "Test"
    Set rng = report.Header_Range
    Assert.IsTrue rng = 7, "Custom header_range is not equal to 7: " & rng.count
    Assert.IsTrue rng.Address = "$F$1", "Custom header_range doesn't start at F1: " & rng.Address
End Sub

'@TestMethod("Report: Data")
Public Sub Test_Full_Range()
    Dim rng As Range: Set rng = report.Full_Range

    Assert.IsTrue rng.count = 5, "Full range is not 5: " & rng.count
    Assert.IsTrue rng.Address = "$A$1:$E$1", "Full range address wrong: " & rng.Address

    ' Add data
    s.Range("A2:E2").value = "Test"
    Set rng = report.Full_Range
    Assert.IsTrue rng.count = 10, "Additional Full range is not 10: " & rng.count
    Assert.IsTrue rng.Address = "$A$1:$E$2", "Full range address wrong: " & rng.Address

    ' Add data with blanks
    s.Range("A3").value = "Test"
    Set rng = report.Full_Range
    Assert.IsTrue rng.count = 15, "Setting Column A results in bad Full_Range count " & rng.count
    Assert.IsTrue rng.Address = "$A$1:$E$3", "Setting Column A results in bad Full_Range address " & rng.Address
    s.Range("B4").value = "Test"
    Set rng = report.Full_Range
    Assert.IsTrue rng.count = 20, "Setting Column B results in bad Full_Range count " & rng.count
    Assert.IsTrue rng.Address = "$A$1:$E$4", "Setting Column A results in bad Full_Range address " & rng.Address
    s.Range("C5").value = "Test"
    Set rng = report.Full_Range
    Assert.IsTrue rng.count = 25, "Setting Column C results in bad Full_Range count " & rng.count
    Assert.IsTrue rng.Address = "$A$1:$E$5", "Setting Column A results in bad Full_Range address " & rng.Address
    s.Range("D6").value = "Test"
    Set rng = report.Full_Range
    Assert.IsTrue rng.count = 30, "Setting Column D results in bad Full_Range count " & rng.count
    Assert.IsTrue rng.Address = "$A$1:$E$6", "Setting Column A results in bad Full_Range address " & rng.Address
    s.Range("E7").value = "Test"
    Set rng = report.Full_Range
    Assert.IsTrue rng.count = 35, "Setting Column E results in bad Full_Range count " & rng.count
    Assert.IsTrue rng.Address = "$A$1:$E$7", "Setting Column A results in bad Full_Range address " & rng.Address

    'TODO
End Sub