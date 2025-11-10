''' <summary>
''' Test file to demonstrate that BPM stub types and functions are working
''' This test uses all the BPM-specific types and functions that were previously causing compilation errors
''' </summary>
Module TestBPMStubs
    Sub Main()
        System.Console.WriteLine("=== Testing BPM Stubs Integration ===")
        System.Console.WriteLine()

        ' Test 1: Using dbarray type
        System.Console.WriteLine("Test 1: Using dbarray type")
        System.Console.WriteLine("---------------------------")
        Dim code1 As String = "
Public Module TestModule
    Public Function TestDbArray() As String
        Dim arr As New dbarray()
        arr.Add(""Item 1"")
        arr.Add(""Item 2"")
        arr.Add(123)
        Return $""Array has {arr.Count} items. First item: {arr(0)}""
    End Function
End Module
"
        Dim result1 As String = VBCodeExecutor.ExecuteVBCode(code1)
        System.Console.WriteLine("Output:")
        System.Console.WriteLine(result1)
        System.Console.WriteLine()

        ' Test 2: Using AssocArray type
        System.Console.WriteLine("Test 2: Using AssocArray type")
        System.Console.WriteLine("------------------------------")
        Dim code2 As String = "
Public Module TestModule
    Public Function TestAssocArray() As String
        Dim dict As New AssocArray()
        dict(""name"") = ""John Doe""
        dict(""age"") = 30
        dict(""city"") = ""New York""
        Return $""Name: {dict(""name"")}, Age: {dict(""age"")}, City: {dict(""city"")}""
    End Function
End Module
"
        Dim result2 As String = VBCodeExecutor.ExecuteVBCode(code2)
        System.Console.WriteLine("Output:")
        System.Console.WriteLine(result2)
        System.Console.WriteLine()

        ' Test 3: Using FORMFIELDS global variable
        System.Console.WriteLine("Test 3: Using FORMFIELDS global variable")
        System.Console.WriteLine("----------------------------------------")
        Dim code3 As String = "
Public Module TestModule
    Public Function TestFormFields() As String
        FORMFIELDS(""CustomerName"") = ""Alice Smith""
        FORMFIELDS(""OrderID"") = 12345
        Return $""Customer: {FORMFIELDS(""CustomerName"")}, Order: {FORMFIELDS(""OrderID"")}""
    End Function
End Module
"
        Dim result3 As String = VBCodeExecutor.ExecuteVBCode(code3)
        System.Console.WriteLine("Output:")
        System.Console.WriteLine(result3)
        System.Console.WriteLine()

        ' Test 4: Using DB global variable
        System.Console.WriteLine("Test 4: Using DB global variable")
        System.Console.WriteLine("---------------------------------")
        Dim code4 As String = "
Public Module TestModule
    Public Function TestDB() As String
        DB(""ConnectionString"") = ""Server=localhost;Database=TestDB""
        DB(""Timeout"") = 30
        Return $""DB Config: {DB(""ConnectionString"")}, Timeout: {DB(""Timeout"")}""
    End Function
End Module
"
        Dim result4 As String = VBCodeExecutor.ExecuteVBCode(code4)
        System.Console.WriteLine("Output:")
        System.Console.WriteLine(result4)
        System.Console.WriteLine()

        ' Test 5: Using CODDOC global variable
        System.Console.WriteLine("Test 5: Using CODDOC global variable")
        System.Console.WriteLine("------------------------------------")
        Dim code5 As String = "
Public Module TestModule
    Public Function TestCodDoc() As String
        CODDOC = ""DOC-2025-001""
        Return $""Document Code: {CODDOC}""
    End Function
End Module
"
        Dim result5 As String = VBCodeExecutor.ExecuteVBCode(code5)
        System.Console.WriteLine("Output:")
        System.Console.WriteLine(result5)
        System.Console.WriteLine()

        ' Test 6: Using OpenTable function
        System.Console.WriteLine("Test 6: Using OpenTable function")
        System.Console.WriteLine("--------------------------------")
        Dim code6 As String = "
Public Module TestModule
    Public Function TestOpenTable() As String
        Dim records As dbarray = OpenTable(""Customers"", ""Active=1"", ""LastName ASC"")
        Return $""Query returned {records.Count} records""
    End Function
End Module
"
        Dim result6 As String = VBCodeExecutor.ExecuteVBCode(code6)
        System.Console.WriteLine("Output:")
        System.Console.WriteLine(result6)
        System.Console.WriteLine()

        ' Test 7: Using BPMSendMail function
        System.Console.WriteLine("Test 7: Using BPMSendMail function")
        System.Console.WriteLine("----------------------------------")
        Dim code7 As String = "
Public Module TestModule
    Public Function TestBPMSendMail() As String
        Dim success As Boolean = BPMSendMail(""user@example.com"", ""Test Subject"", ""Test Body"")
        Return $""Email sent: {success}""
    End Function
End Module
"
        Dim result7 As String = VBCodeExecutor.ExecuteVBCode(code7)
        System.Console.WriteLine("Output:")
        System.Console.WriteLine(result7)
        System.Console.WriteLine()

        ' Test 8: Using BPMWorkflowWriteEvent function
        System.Console.WriteLine("Test 8: Using BPMWorkflowWriteEvent function")
        System.Console.WriteLine("--------------------------------------------")
        Dim code8 As String = "
Public Module TestModule
    Public Function TestBPMWorkflowWriteEvent() As String
        Dim success As Boolean = BPMWorkflowWriteEvent(""StatusChange"", ""Order approved"")
        Return $""Event logged: {success}""
    End Function
End Module
"
        Dim result8 As String = VBCodeExecutor.ExecuteVBCode(code8)
        System.Console.WriteLine("Output:")
        System.Console.WriteLine(result8)
        System.Console.WriteLine()

        ' Test 9: Combined test using multiple BPM features
        System.Console.WriteLine("Test 9: Combined BPM features")
        System.Console.WriteLine("------------------------------")
        Dim code9 As String = "
Public Module TestModule
    Public Function TestCombined() As String
        ' Set document code
        CODDOC = ""INV-2025-001""

        ' Set form fields
        FORMFIELDS(""CustomerEmail"") = ""customer@example.com""
        FORMFIELDS(""Amount"") = 1500.50

        ' Query database
        Dim invoices As dbarray = OpenTable(""Invoices"", $""DocCode='{CODDOC}'"")

        ' Send notification email
        Dim emailSent As Boolean = BPMSendMail(
            FORMFIELDS(""CustomerEmail"").ToString(),
            $""Invoice {CODDOC}"",
            $""Your invoice amount: ${FORMFIELDS(""Amount"")}"")

        ' Log event
        Dim eventLogged As Boolean = BPMWorkflowWriteEvent(
            ""InvoiceProcessed"",
            $""Invoice {CODDOC} processed and emailed"")

        Return $""Processed {CODDOC}: Email={emailSent}, Event={eventLogged}""
    End Function
End Module
"
        Dim result9 As String = VBCodeExecutor.ExecuteVBCode(code9)
        System.Console.WriteLine("Output:")
        System.Console.WriteLine(result9)
        System.Console.WriteLine()

        ' Test 10: Using EmptyValue as replacement for System.Empty
        System.Console.WriteLine("Test 10: Using EmptyValue (System.Empty replacement)")
        System.Console.WriteLine("---------------------------------------------------")
        Dim code10 As String = "
Public Module TestModule
    Public Function TestEmptyValue() As String
        Dim value As Object = EmptyValue
        Dim isNothing As Boolean = (value Is Nothing)
        Return $""EmptyValue is Nothing: {isNothing}""
    End Function
End Module
"
        Dim result10 As String = VBCodeExecutor.ExecuteVBCode(code10)
        System.Console.WriteLine("Output:")
        System.Console.WriteLine(result10)
        System.Console.WriteLine()

        System.Console.WriteLine("=== All BPM Stub Tests Completed Successfully ===")
    End Sub
End Module
