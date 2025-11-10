Module TestComplexVariables
    Sub Main(args As String())
        System.Console.WriteLine("=== Testing Complex Variable Injection (DB, CODDOC, ACTIVITY) ===")
        System.Console.WriteLine()

        ' Simulate the user's scenario with DB, CODDOC, ACTIVITY variables
        System.Console.WriteLine("Test: Simulating DB, CODDOC, ACTIVITY variables")
        System.Console.WriteLine("------------------------------------------------")

        ' Create a mock database object (Dictionary to simulate)
        Dim mockDB As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"ConnectionString", "Server=localhost;Database=TestDB"},
            {"IsConnected", True}
        }

        ' Create FORMFIELDS (another Dictionary)
        Dim mockFormFields As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"Field1", "Value1"},
            {"Field2", "Value2"}
        }

        Dim code As String = "
Public Module TestModule
    Public Function TestVariables() As String
        Dim result As String = """"

        ' Try to access DB
        If DB IsNot Nothing Then
            result &= ""DB exists. ""
        End If

        ' Try to access CODDOC
        result &= ""CODDOC = "" & CODDOC.ToString() & "". ""

        ' Try to access ACTIVITY
        result &= ""ACTIVITY = "" & ACTIVITY.ToString() & "". ""

        ' Try to access FORMFIELDS
        If FORMFIELDS IsNot Nothing Then
            result &= ""FORMFIELDS exists.""
        End If

        Return result
    End Function
End Module
"

        Dim vars As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"FORMFIELDS", mockFormFields},
            {"DB", mockDB},
            {"ACTIVITY", 12345},
            {"CODDOC", "DOC123"}
        }

        System.Console.WriteLine("Variables being passed:")
        System.Console.WriteLine("  FORMFIELDS: Dictionary with 2 entries")
        System.Console.WriteLine("  DB: Dictionary simulating database")
        System.Console.WriteLine("  ACTIVITY: 12345 (Integer)")
        System.Console.WriteLine("  CODDOC: ""DOC123"" (String)")
        System.Console.WriteLine()

        System.Console.WriteLine("Executing code...")
        System.Console.WriteLine()

        Dim result As String = VBCodeExecutor.ExecuteVBCodeWithVariables(code, vars)
        System.Console.WriteLine("Result:")
        System.Console.WriteLine(result)
        System.Console.WriteLine()

        ' Test 2: Accessing properties of complex objects
        System.Console.WriteLine("Test 2: Accessing properties of complex objects")
        System.Console.WriteLine("-----------------------------------------------")

        Dim code2 As String = "
Public Module TestModule
    Public Function TestObjectProperties() As String
        Dim result As String = """"

        ' Try to access DB as Dictionary and get a value
        Dim dbDict As System.Collections.Generic.Dictionary(Of String, Object) = CType(DB, System.Collections.Generic.Dictionary(Of String, Object))
        If dbDict.ContainsKey(""ConnectionString"") Then
            result &= ""DB ConnectionString: "" & dbDict(""ConnectionString"").ToString() & "". ""
        End If

        ' Access FORMFIELDS
        Dim formDict As System.Collections.Generic.Dictionary(Of String, Object) = CType(FORMFIELDS, System.Collections.Generic.Dictionary(Of String, Object))
        If formDict.ContainsKey(""Field1"") Then
            result &= ""FORMFIELDS.Field1: "" & formDict(""Field1"").ToString()
        End If

        Return result
    End Function
End Module
"

        System.Console.WriteLine("Executing code with object property access...")
        System.Console.WriteLine()

        Dim result2 As String = VBCodeExecutor.ExecuteVBCodeWithVariables(code2, vars)
        System.Console.WriteLine("Result:")
        System.Console.WriteLine(result2)
        System.Console.WriteLine()

        System.Console.WriteLine("=== Tests completed ===")
    End Sub
End Module
