'@TestModule("Validation")
'@Folder("Tests.Validation")
'@ModuleDescription "Tests the Validation.cls class"
Attribute VB_Name = "TestValidation"

Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object
Private Validation As Validation
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
    
    Set Transform = New Transform
    Set Validation = New Validation
    Validation.Report = Transform.Report
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("SSN")
Public Sub Validate_Passing_SSN()
    ' Success: normal SSN
    Assert.IsTrue Validation.validate_social("155505549")
End Sub

'@TestMethod("SSN")
Public Sub Validate_Passing_SSN_With_Dashes()
    ' Success: normal SSN with dashes
    Assert.IsTrue Validation.validate_social("155-50-5549")
End Sub

'@TestMethod("SSN")
Public Sub Validate_Passing_SSN_With_Spaces()
    ' Success: normal SSN with dashes and spaces
    Assert.IsTrue Validation.validate_social(" 155 - 50 - 5549 ")
    Assert.IsTrue Validation.validate_social("155 50 5549")
End Sub

'@TestMethod("SSN")
Public Sub Validate_Passing_Dummy_SSN()
    ' Success: normal SSN the website doesn't accept
    Assert.IsTrue Validation.validate_social("123-45-6789")
    Assert.IsTrue Validation.validate_social("123456789")
End Sub

'@TestMethod("SSN")
Public Sub Validate_Passing_SSN_Zero_Prefix()
    ' Success: normal SSN prefixed by zeroes
    Assert.IsTrue Validation.validate_social("0123456789")
    Assert.IsTrue Validation.validate_social("00123456789")
    Assert.IsTrue Validation.validate_social("000123456789")
    Assert.IsTrue Validation.validate_social("0000123456789")
End Sub

'@TestMethod("SSN")
Public Sub Validate_Failing_SSN_With_Nine()
    ' Failure: Invalid ssn begins with 9
    Assert.IsFalse Validation.validate_social("900-50-5549")
    Assert.IsFalse Validation.validate_social("923456789")
End Sub

'@TestMethod("SSN")
Public Sub Validate_Failing_SSN_Too_Many()
    ' Failure: too many characters
    Assert.IsFalse Validation.validate_social("1555055499")
    Assert.IsFalse Validation.validate_social("155-50-554910")
End Sub

'@TestMethod("SSN")
Public Sub Validate_Failing_SSN_Too_Few()
    ' Failure: too few characters
    Assert.IsFalse Validation.validate_social("155-50-554")
    Assert.IsFalse Validation.validate_social("155-50-55")
    Assert.IsFalse Validation.validate_social("89123456")
    Assert.IsFalse Validation.validate_social("7891234")
    Assert.IsFalse Validation.validate_social("678912")
    Assert.IsFalse Validation.validate_social("56789")
    Assert.IsFalse Validation.validate_social("4567")
    Assert.IsFalse Validation.validate_social("345")
    Assert.IsFalse Validation.validate_social("23")
    Assert.IsFalse Validation.validate_social("1")
    Assert.IsFalse Validation.validate_social("")
End Sub

'@TestMethod("SSN")
Public Sub Validate_Failing_SSN_Too_Few_With_Dashes()
    ' Failure: Exactly 9 characters but not 9 digits
    Assert.IsFalse Validation.validate_social("155-50-55")
    Assert.IsFalse Validation.validate_social("155 50 55")
    Assert.IsFalse Validation.validate_social(" 1555055 ")
End Sub

'@TestMethod("SSN")
Public Sub Validate_Failing_SSN_NonDigits()
    ' Failure: Non-digits present
    Assert.IsFalse Validation.validate_social("abc-de-fghi")
    Assert.IsFalse Validation.validate_social("abcdefghi")
    Assert.IsFalse Validation.validate_social("055-a1-2342")
    Assert.IsFalse Validation.validate_social("055-$1-2345")
    ' Failure: prefix/suffix
    Assert.IsFalse Validation.validate_social("@155-21-2345")
    Assert.IsFalse Validation.validate_social("155-21-2345/")
    ' Failure: all normal special chars
    Assert.IsFalse Validation.validate_social("123-45678!")
    Assert.IsFalse Validation.validate_social("123-45678@")
    Assert.IsFalse Validation.validate_social("123-45678#")
    Assert.IsFalse Validation.validate_social("123-45678$")
    Assert.IsFalse Validation.validate_social("123-45678%")
    Assert.IsFalse Validation.validate_social("123-45678^")
    Assert.IsFalse Validation.validate_social("123-45678&")
    Assert.IsFalse Validation.validate_social("123-45678*")
    Assert.IsFalse Validation.validate_social("123-45678(")
    Assert.IsFalse Validation.validate_social("123-45678)")
    Assert.IsFalse Validation.validate_social("123-45678{")
    Assert.IsFalse Validation.validate_social("123-45678}")
End Sub

