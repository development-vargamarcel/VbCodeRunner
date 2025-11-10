Module TestVariableInjection
    Sub Main(args As String())
        System.Console.WriteLine("=== Testing Variable Injection ===")
        System.Console.WriteLine()

        ' Test 1: Basic variable injection
        System.Console.WriteLine("Test 1: Basic variable injection")
        System.Console.WriteLine("---------------------------------")
        Dim code1 As String = "
Public Module TestModule
    Public Function Calculate() As Integer
        Return x * y + z
    End Function
End Module
"
        Dim vars1 As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"x", 10},
            {"y", 5},
            {"z", 3}
        }

        System.Console.WriteLine("Code before injection:")
        System.Console.WriteLine(code1)
        System.Console.WriteLine()

        Dim result1 As String = VBCodeExecutor.ExecuteVBCodeWithVariables(code1, vars1)
        System.Console.WriteLine("Result:")
        System.Console.WriteLine(result1)
        System.Console.WriteLine()

        ' Test 2: Without "Public" keyword
        System.Console.WriteLine("Test 2: Module without 'Public' keyword")
        System.Console.WriteLine("----------------------------------------")
        Dim code2 As String = "
Module TestModule
    Function Calculate() As Integer
        Return a + b
    End Function
End Module
"
        Dim vars2 As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"a", 100},
            {"b", 200}
        }

        Dim result2 As String = VBCodeExecutor.ExecuteVBCodeWithVariables(code2, vars2)
        System.Console.WriteLine("Result:")
        System.Console.WriteLine(result2)
        System.Console.WriteLine()

        ' Test 3: Different line endings (Windows CRLF)
        System.Console.WriteLine("Test 3: Different line endings")
        System.Console.WriteLine("-------------------------------")
        Dim code3 As String = "Public Module TestModule" & vbCrLf &
                              "    Public Function Test() As Integer" & vbCrLf &
                              "        Return num * 2" & vbCrLf &
                              "    End Function" & vbCrLf &
                              "End Module"
        Dim vars3 As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"num", 42}
        }

        System.Console.WriteLine("Code:")
        System.Console.WriteLine(code3)
        System.Console.WriteLine()

        Dim result3 As String = VBCodeExecutor.ExecuteVBCodeWithVariables(code3, vars3)
        System.Console.WriteLine("Result:")
        System.Console.WriteLine(result3)
        System.Console.WriteLine()

        System.Console.WriteLine("=== Tests completed ===")
    End Sub
End Module
