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

'@TestMethod("Validation: SSN")
Public Sub Validate_Passing_SSN()
    ' Success: normal SSN
    Assert.IsTrue Validation.validate_social("155505549"), "Basic SSN failed: 155505549"
End Sub

'@TestMethod("Validation: SSN")
Public Sub Validate_Passing_SSN_With_Dashes()
    ' Success: normal SSN with dashes
    Assert.IsTrue Validation.validate_social("155-50-5549"), "SSN with dashes failed: 155-50-5549"
End Sub

'@TestMethod("Validation: SSN")
Public Sub Validate_Passing_SSN_With_Spaces()
    ' Success: normal SSN with dashes and spaces
    Assert.IsTrue Validation.validate_social(" 155 - 50 - 5549 "), "SSN With dashes and spaces failed:  155 - 50 - 5549 "
    Assert.IsTrue Validation.validate_social("155 50 5549"), "SSN with spaces failed: 155 50 5549"
End Sub

'@TestMethod("Validation: SSN")
Public Sub Validate_Passing_Dummy_SSN()
    ' Success: normal SSN the website doesn't accept
    Assert.IsTrue Validation.validate_social("123-45-6789"), "Dummy SSN failed: 123-45-6789"
    Assert.IsTrue Validation.validate_social("123456789"), "Dummy SSN failed: 123456789"
End Sub

'@TestMethod("Validation: SSN")
Public Sub Validate_Passing_SSN_Zero_Prefix()
    ' Success: normal SSN prefixed by zeroes
    Assert.IsTrue Validation.validate_social("0123456789"), "0-prefix failed: 0123456789"
    Assert.IsTrue Validation.validate_social("00123456789"), "00-prefix failed: 00123456789"
    Assert.IsTrue Validation.validate_social("000123456789"), "000-prefix failed: 000123456789"
    Assert.IsTrue Validation.validate_social("0000123456789"), "0000-prefix failed: 0000123456789"
End Sub

'@TestMethod("Validation: SSN")
Public Sub Validate_Failing_SSN_With_Nine()
    ' Failure: Invalid ssn begins with 9
    Assert.IsFalse Validation.validate_social("900-50-5549"), "9-prefix SSN passed: 900-50-5549"
    Assert.IsFalse Validation.validate_social("923456789"), "9-prefix SSN passed: 923456789"
End Sub

'@TestMethod("Validation: SSN")
Public Sub Validate_Failing_SSN_Too_Many()
    ' Failure: too many characters
    Assert.IsFalse Validation.validate_social("1555055499"), "SSN too big passed: 1555055499"
    Assert.IsFalse Validation.validate_social("155-50-554910"), "SSN too big (w/ dashes) passed: 155-50-554910"
End Sub

'@TestMethod("Validation: SSN")
Public Sub Validate_Failing_SSN_Too_Few()
    ' Failure: too few characters
    Assert.IsFalse Validation.validate_social("155-50-554"), "SSN too small passed: 155-50-554"
    Assert.IsFalse Validation.validate_social("155-50-55"), "SSN too small passed: 155-50-55"
    Assert.IsFalse Validation.validate_social("89123456"), "SSN too small passed: 89123456"
    Assert.IsFalse Validation.validate_social("7891234")
    Assert.IsFalse Validation.validate_social("678912")
    Assert.IsFalse Validation.validate_social("56789")
    Assert.IsFalse Validation.validate_social("4567")
    Assert.IsFalse Validation.validate_social("345")
    Assert.IsFalse Validation.validate_social("23")
    Assert.IsFalse Validation.validate_social("1")
    Assert.IsFalse Validation.validate_social(""), "Empty SSN passed"
End Sub

'@TestMethod("Validation: SSN")
Public Sub Validate_Failing_SSN_Too_Few_With_Dashes()
    ' Failure: Exactly 9 characters but not 9 digits
    Assert.IsFalse Validation.validate_social("155-50-55"), "SSN too small (with dashes) passed: 155-50-55"
    Assert.IsFalse Validation.validate_social("155 50 55"), "SSN too small (with spaces) passed: 155 50 55"
    Assert.IsFalse Validation.validate_social(" 1555055 "), "SSN too small passed:  1555055 "
End Sub

'@TestMethod("Validation: SSN")
Public Sub Validate_Failing_SSN_NonDigits()
    ' Failure: Non-digits present
    Assert.IsFalse Validation.validate_social("abc-de-fghi"), "SSN with letters passed: abc-de-fghi"
    Assert.IsFalse Validation.validate_social("abcdefghi"), "SSN with letters passed: abcdefghi"
    Assert.IsFalse Validation.validate_social("055-a1-2342"), "SSN with one letter passed: 055-a1-2342"
    Assert.IsFalse Validation.validate_social("055-$1-2345"), "SSN with symbol passed: 055-$1-2345"
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

'@TestMethod("Validation: First Name")
Public Sub Validate_Good_First_Name()
    Assert.AreEqual Validation.validate_first_name("John"), "John", "Normal first name failed: John"
End Sub

'@TestMethod("Validation: First Name")
Public Sub Validate_Good_First_Name_Dashes()
    Assert.AreEqual Validation.validate_first_name("Jo-Anne"), "Jo-Anne", "Name with dashes failed: Jo-Anne"
End Sub

'@TestMethod("Validation: First Name")
Public Sub Validate_Good_First_Name_Spaces()
    Assert.AreEqual Validation.validate_first_name("Jo Anne"), "Jo Anne", "First Name with dashes failed: Jo Anne"
    Assert.AreEqual Validation.validate_first_name("Jo - Anne"), "Jo - Anne", "First Name with dashes failed: Jo - Anne"
    Assert.AreEqual Validation.validate_first_name("   Jo - Anne "), "Jo - Anne", "First Name with dashes/spaces failed:   Jo - Anne "
End Sub

'@TestMethod("Validation: First Name")
Public Sub Validate_First_Name_Special()
    Assert.AreEqual Validation.validate_first_name("Jo'hn."), "Jo'hn.", "First name special chars failed: Jo'hn."
    Assert.AreEqual Validation.validate_first_name(" !@#$%^&*(){}[]:;'<>,./?"), "!@#$%^&*(){}[]:;'<>,./?", "First name special chars failed:  !@#$%^&*(){}[]:;'<>,./?"
End Sub

'@TestMethod("Validation: First Name")
Public Sub Validate_First_Name_NonEnglish()
    Assert.AreEqual Validation.validate_first_name("陳大文"), "陳大文", "Chinese characters failed: 陳大文"
    Assert.AreEqual Validation.validate_first_name("Aarón"), "Aarón", "Spanish accent failed: Aarón"
    Assert.AreEqual Validation.validate_first_name("Aarão DiĎբonísio"), "Aarão DiĎբonísio", "Portugese accent failed: Aarão DiĎբonísio"
    Assert.AreEqual Validation.validate_first_name("Варфоломей"), "Варфоломей", "Russian characters failed: Варфоломей"
    Assert.AreEqual Validation.validate_first_name("Κωνσταντίνος"), "Κωνσταντίνος", "Greek characters failed: Κωνσταντίνος"
    Assert.AreEqual Validation.validate_first_name("Кир, Сайрус, Сайрес"), "Кир, Сайрус, Сайрес", "Ukranian characters failed: Кир, Сайрус, Сайрес"
    Assert.AreEqual Validation.validate_first_name("عباس"), "عباس", "Arabic characters failed: عباس"
    Assert.AreEqual Validation.validate_first_name("פֵא סוֹפִית,פֵה סוֹפִית"), "פֵא סוֹפִית,פֵה סוֹפִית", "Hebrew characters failed: פֵא סוֹפִית,פֵה סוֹפִית"
    
End Sub

'@TestMethod("Validation: First Name")
Public Sub Validate_First_Name_Lowercase()
    Assert.AreEqual Validation.validate_first_name("john"), "john", "Lowercase characters failed: john"
End Sub

'@TestMethod("Validation: First Name")
Public Sub Validate_First_Name_Short()
    Assert.AreEqual Validation.validate_first_name("Al"), "Al", "Short first name failed: Al"
    ' Real people have single character first names. "J" (for Jay) is one example.
    Assert.AreEqual Validation.validate_first_name("J"), "J", "Single character first name failed: J"
End Sub

'@TestMethod("Validation: First Name")
Public Sub Validate_First_Name_Long()
    ' This is a real first name of a real person
    Dim name As String: name = "Adolph Blaine Charles David Earl Frederick Gerald Hubert Irvin John Kenneth Lloyd Martin Nero Oliver Paul Quincy Randolph Sherman Thomas Uncas Victor William Xerxes Yancy Zeus Wolfeschlegelsteinhausenbergerdorffwelchevoralternwarengewissenhaftschaferswessenschafewarenwohlgepflegeundsorgfaltigkeitbeschutzenvonangreifendurchihrraubgierigfeindewelchevoralternzwolftausendjahresvorandieerscheinenvanderersteerdemenschderraumschiffgebrauchlichtalsseinursprungvonkraftgestartseinlangefahrthinzwischensternartigraumaufdersuchenachdiesternwelchegehabtbewohnbarplanetenkreisedrehensichundwohinderneurassevonverstandigmenschlichkeitkonntefortpflanzenundsicherfreuenanlebenslanglichfreudeundruhemitnichteinfurchtvorangreifenvonandererintelligentgeschopfsvonhinzwischensternartigraum"
    Assert.AreEqual Validation.validate_first_name(name), name, "Long First name failed"
End Sub

'@TestMethod("Validation: First Name")
Public Sub Validate_First_Name_Numbers()
    Assert.AreEqual Validation.validate_first_name("John 2nd"), "John 2nd", "Numbers failed in first name: John 2nd"
End Sub

'@TestMethod("Validation: First Name")
Public Sub Validate_First_Name_Empty()
    Assert.IsFalse Validation.validate_first_name(""), "Empty First Name Passed"
    Assert.IsFalse Validation.validate_first_name("     "), "Empty (spaces) First Name Passed"
End Sub

'@TestMethod("Validation: First Name")
Public Sub Validate_First_Name_StartsWithNumber()
    Assert.IsFalse Validation.validate_first_name("2"), "First name of '2' passed"
    Assert.IsFalse Validation.validate_first_name("2nd Josh"), "First name beginning with digit passed: 2nd Josh"
End Sub

'@TestMethod("Validation: Last Name")
Public Sub Validate_Good_Last_Name()
    Assert.AreEqual Validation.validate_last_name("Smith"), "Smith", "Normal Last name failed: Smith"
End Sub

'@TestMethod("Validation: Last Name")
Public Sub Validate_Good_Last_Name_Dashes()
    Assert.AreEqual Validation.validate_last_name("Jo-Anne"), "Jo-Anne", "Name with dashes failed: Jo-Anne"
End Sub

'@TestMethod("Validation: Last Name")
Public Sub Validate_Good_Last_Name_Spaces()
    Assert.AreEqual Validation.validate_last_name("Jo Anne"), "Jo Anne", "Name with dashes failed: Jo Anne"
    Assert.AreEqual Validation.validate_last_name("Jo - Anne"), "Jo - Anne", "Name with dashes failed: Jo - Anne"
    Assert.AreEqual Validation.validate_last_name("   Jo - Anne "), "Jo - Anne", "Name with dashes/spaces failed:   Jo - Anne "
End Sub

'@TestMethod("Validation: Last Name")
Public Sub validate_last_name_Special()
    Assert.AreEqual Validation.validate_last_name("Jo'hn."), "Jo'hn.", "Last name special chars failed: Jo'hn."
    Assert.AreEqual Validation.validate_last_name(" !@#$%^&*(){}[]:;'<>,./?"), "!@#$%^&*(){}[]:;'<>,./?", "Last name special chars failed:  !@#$%^&*(){}[]:;'<>,./?"
End Sub

'@TestMethod("Validation: Last Name")
Public Sub validate_last_name_NonEnglish()
    Assert.AreEqual Validation.validate_last_name("陳大文"), "陳大文", "Chinese characters failed: 陳大文"
    Assert.AreEqual Validation.validate_last_name("Aarón"), "Aarón", "Spanish accent failed: Aarón"
    Assert.AreEqual Validation.validate_last_name("Aarão DiĎբonísio"), "Aarão DiĎբonísio", "Portugese accent failed: Aarão DiĎբonísio"
    Assert.AreEqual Validation.validate_last_name("Варфоломей"), "Варфоломей", "Russian characters failed: Варфоломей"
    Assert.AreEqual Validation.validate_last_name("Κωνσταντίνος"), "Κωνσταντίνος", "Greek characters failed: Κωνσταντίνος"
    Assert.AreEqual Validation.validate_last_name("Кир, Сайрус, Сайрес"), "Кир, Сайрус, Сайрес", "Ukranian characters failed: Кир, Сайрус, Сайрес"
    Assert.AreEqual Validation.validate_last_name("عباس"), "عباس", "Arabic characters failed: عباس"
    Assert.AreEqual Validation.validate_last_name("פֵא סוֹפִית,פֵה סוֹפִית"), "פֵא סוֹפִית,פֵה סוֹפִית", "Hebrew characters failed: פֵא סוֹפִית,פֵה סוֹפִית"
End Sub

'@TestMethod("Validation: Last Name")
Public Sub validate_last_name_Short()
    Assert.AreEqual Validation.validate_last_name("Al"), "Al", "Short Last name failed: Al"
    Assert.AreEqual Validation.validate_last_name("B"), "B", "Short Last name failed: B"
End Sub

    
'@TestMethod("Validation: Last Name")
Public Sub validate_last_name_lowercase()
    Assert.AreEqual Validation.validate_last_name("leBrawn"), "leBrawn", "Lowercase last name failed"
End Sub

'@TestMethod("Validation: Last Name")
Public Sub validate_last_name_Long()
    Assert.AreEqual Validation.validate_last_name("abcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyz"), "abcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyz", "Long Last name failed"
End Sub

'@TestMethod("Validation: Last Name")
Public Sub validate_last_name_Numbers()
    Assert.AreEqual Validation.validate_last_name("John 2nd"), "John 2nd", "Numbers failed in Last name: John 2nd"
End Sub

'@TestMethod("Validation: Last Name")
Public Sub validate_last_name_Empty()
    Assert.IsFalse Validation.validate_last_name(""), "Empty Last Name Passed"
    Assert.IsFalse Validation.validate_last_name("     "), "Empty (spaces) Last Name Passed"
End Sub

'@TestMethod("Validation: Last Name")
Public Sub Validate_Last_Name_StartsWithNumber()
    Assert.IsFalse Validation.validate_last_name("2"), "Last name of '2' passed"
    Assert.IsFalse Validation.validate_last_name("2nd Smith"), "Last name beginning with digit passed: 2nd Josh"
End Sub

'@TestMethod("Validation: NYSLRS ID")
Public Sub Validate_Good_NYSLRS_ID()
    Assert.AreEqual Validation.validate_id("R02345678"), "R02345678", "Normal NYSLRS ID failed: R02345678"
    Assert.AreEqual Validation.validate_id("R12345678"), "R12345678", "Normal NYSLRS ID failed: R12345678"
    Assert.AreEqual Validation.validate_id("R22345678"), "R22345678", "Normal NYSLRS ID failed: R22345678"
End Sub

'@TestMethod("Validation: NYSLRS ID")
Public Sub Validate_Good_NYSLRS_ID_formatting()
    Assert.AreEqual Validation.validate_id("r10450174"), "R10450174", "Lowercase NYSLRS ID failed"
    Assert.AreEqual Validation.validate_id("r-10450174"), "R10450174", "Lowercase NYSLRS ID (with dash) failed"
    Assert.AreEqual Validation.validate_id("R-10450174"), "R10450174", "Dashed NYSLRS ID failed"
    Assert.AreEqual Validation.validate_id("R--10450174"), "R10450174", "Dashed NYSLRS ID failed: R--10450174"
    Assert.AreEqual Validation.validate_id("R- - -10450174"), "R10450174", "Dashed NYSLRS ID failed: R- - -10450174"
    Assert.AreEqual Validation.validate_id(" r - 10450174 "), "R10450174", "Spaces in NYSLRS ID failed"
End Sub

'@TestMethod("Validation: NYSLRS ID")
Public Sub Validate_Failed_NYSLRS_ID()
    Assert.IsFalse Validation.validate_id("A10450174"), "NYSLRS ID beginning with A passed: A10450174"
    Assert.IsFalse Validation.validate_id("A-10450174"), "NYSLRS ID beginning with A passed: A-10450174"
    Assert.IsFalse Validation.validate_id(" A - 10450174"), "NYSLRS ID beginning with A passed:  A - 10450174"
    Assert.IsFalse Validation.validate_id("$10450174"), "NYSLRS ID beginning with symbol passed: $10450174"
    Assert.IsFalse Validation.validate_id("10450174"), "NYSLRS ID without R failed: 10450174"
    Assert.IsFalse Validation.validate_id("104501749"), "NYSLRS ID without R failed: 104501749"
    Assert.IsFalse Validation.validate_id("-10450174"), "NYSLRS ID without R failed: -10450174"
    Assert.IsFalse Validation.validate_id("1-10450174"), "NYSLRS ID without R failed: 1-10450174"
End Sub

    
'@TestMethod("Validation: NYSLRS ID")
Public Sub Validate_Failed_NYSLRS_ID_TooHigh()
    Assert.IsFalse Validation.validate_id("R30450174"), "NYSLRS ID beginning with R3 passed (should be R1)"
    Assert.IsFalse Validation.validate_id("R40450174"), "NYSLRS ID beginning with R4 passed (should be R1)"
    Assert.IsFalse Validation.validate_id("R50450174"), "NYSLRS ID beginning with R5 passed (should be R1)"
    Assert.IsFalse Validation.validate_id("R60450174"), "NYSLRS ID beginning with R6 passed (should be R1)"
    Assert.IsFalse Validation.validate_id("R70450174"), "NYSLRS ID beginning with R7 passed (should be R1)"
    Assert.IsFalse Validation.validate_id("R80450174"), "NYSLRS ID beginning with R8 passed (should be R1)"
    Assert.IsFalse Validation.validate_id("R90450174"), "NYSLRS ID beginning with R9 passed (should be R1)"
End Sub

'@TestMethod("Validation: NYSLRS ID")
Public Sub validate_id_TooLong()
    Assert.IsFalse Validation.validate_id("R123456789"), "NYSLRS ID too long passed: R123456789"
    Assert.IsFalse Validation.validate_id("r - 1234567890"), "NYSLRS ID too long passed: r - 1234567890"
End Sub

'@TestMethod("Validation: NYSLRS ID")
Public Sub validate_id_TooShort()
    Assert.IsFalse Validation.validate_id("R1234567"), "NYSLRS ID too short passed: R1234567"
    Assert.IsFalse Validation.validate_id("r - 123456"), "NYSLRS ID too short passed: r - 123456"
    Assert.IsFalse Validation.validate_id("r - 12345"), "NYSLRS ID too short passed: r - 12345"
    Assert.IsFalse Validation.validate_id("r - 1234"), "NYSLRS ID too short passed: r - 1234"
    Assert.IsFalse Validation.validate_id("r - 123"), "NYSLRS ID too short passed: r - 123"
    Assert.IsFalse Validation.validate_id("r - 12"), "NYSLRS ID too short passed: r - 12"
    Assert.IsFalse Validation.validate_id("r - 1"), "NYSLRS ID too short passed: r - 1"
    Assert.IsFalse Validation.validate_id("r - "), "NYSLRS ID too short passed: r - "
End Sub

'@TestMethod("Validation: NYSLRS ID")
Public Sub validate_id_BadChars()
    Assert.IsFalse Validation.validate_id("R-12345678a"), "NYSLRS ID with bad characters passed: R-12345678a"
    Assert.IsFalse Validation.validate_id("R-12345678#"), "NYSLRS ID with bad characters passed: R-12345678#"
    Assert.IsFalse Validation.validate_id("R-12345678-"), "NYSLRS ID with bad characters passed: R-12345678-"
End Sub

    
'@TestMethod("Validation: Employee Record")
Public Sub validate_good_employee_record()
    Assert.AreEqual Validation.validate_employee_record("1"), "1", "Employee record failed: 1"
    Assert.AreEqual Validation.validate_employee_record("2"), "2", "Employee record failed: 2"
    Assert.AreEqual Validation.validate_employee_record("10"), "10", "Employee record failed: 10"
    Assert.AreEqual Validation.validate_employee_record("100"), "100", "Employee record failed: 100"
    Assert.AreEqual Validation.validate_employee_record("999"), "999", "Employee record failed: 999"
End Sub

'@TestMethod("Validation: Employee Record")
Public Sub validate_employee_record_formatting()
    Assert.AreEqual Validation.validate_employee_record(" 1 "), "1", "Employee record with spaces failed: 1"
    Assert.AreEqual Validation.validate_employee_record("  901   "), "901", "Employee record with spaces failed: 901"
End Sub

'@TestMethod("Validation: Employee Record")
Public Sub validate_employee_record_prefix_zero()
    Assert.AreEqual Validation.validate_employee_record("01"), "1", "Employee record with starting zeroes failed: 01"
    Assert.AreEqual Validation.validate_employee_record("00901"), "901", "Employee record with starting zeroes failed: 00901"
    Assert.AreEqual Validation.validate_employee_record("  0055 "), "55", "Employee record with starting zeroes failed:   0055  "
End Sub

'@TestMethod("Validation: Employee Record")
Public Sub validate_failed_employee_record_toolong()
    Assert.IsFalse Validation.validate_employee_record("1000"), "Employee record too long but passed: 1000"
    Assert.IsFalse Validation.validate_employee_record("99999"), "Employee record too long but passed: 99999"
End Sub

'@TestMethod("Validation: Employee Record")
Public Sub validate_failed_employee_record_empty()
    Assert.IsFalse Validation.validate_employee_record(""), "Employee record empty but passed"
    Assert.IsFalse Validation.validate_employee_record("   "), "Employee record only spaces but passed"
End Sub

'@TestMethod("Validation: Employee Record")
Public Sub validate_employee_record_onlyzero()
    Assert.AreEqual Validation.validate_employee_record("0"), "0", "Employee record of 0 failed"
    Assert.AreEqual Validation.validate_employee_record("00"), "0", "Employee record of 00 passed"
    Assert.AreEqual Validation.validate_employee_record(" 000 "), "0", "Employee record of 000 (with spaces) passed"
End Sub

'@TestMethod("Validation: Employee Record")
Public Sub validate_employee_record_nondigits()
    Assert.IsFalse Validation.validate_employee_record("a"), "Employee record non-digits passed: a"
    Assert.IsFalse Validation.validate_employee_record("$"), "Employee record non-digits passed: $"
    Assert.IsFalse Validation.validate_employee_record(" cab "), "Employee record non-digits passed:  cab "
    Assert.IsFalse Validation.validate_employee_record("-1"), "Employee record non-digits passed: -1 "
End Sub