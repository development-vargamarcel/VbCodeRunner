Module Program
    Sub Main(args As String())
        System.Console.WriteLine("=== VB Code Executor using Roslyn ===")
        System.Console.WriteLine()

        System.Console.WriteLine("Example 1: Simple calculation")
        System.Console.WriteLine("------------------------------")
        Dim code1 As String = "
Public Module TestModule
    Public Function Calculate() As Integer
        Return 5 + 10 * 2
    End Function
End Module
"
        System.Console.WriteLine("Code to execute:")
        System.Console.WriteLine(code1)
        Dim result1 As String = VBCodeExecutor.ExecuteVBCode(code1)
        System.Console.WriteLine("Output:")
        System.Console.WriteLine(result1)
        System.Console.WriteLine()

        System.Console.WriteLine("Example 2: String manipulation with Console output")
        System.Console.WriteLine("---------------------------------------------------")
        Dim code2 As String = "
Public Module TestModule
    Public Function ProcessString() As String
        Dim text As String = ""Hello from dynamic VB code!""
        Console.WriteLine(""Processing: "" & text)
        Return text.ToUpper()
    End Function
End Module
"
        System.Console.WriteLine("Code to execute:")
        System.Console.WriteLine(code2)
        Dim result2 As String = VBCodeExecutor.ExecuteVBCode(code2)
        System.Console.WriteLine("Output:")
        System.Console.WriteLine(result2)
        System.Console.WriteLine()

        System.Console.WriteLine("Example 3: Loop with output")
        System.Console.WriteLine("---------------------------")
        Dim code3 As String = "
Public Module TestModule
    Public Sub PrintNumbers()
        For i As Integer = 1 To 5
            Console.WriteLine(""Number: "" & i)
        Next
    End Sub
End Module
"
        System.Console.WriteLine("Code to execute:")
        System.Console.WriteLine(code3)
        Dim result3 As String = VBCodeExecutor.ExecuteVBCode(code3)
        System.Console.WriteLine("Output:")
        System.Console.WriteLine(result3)
        System.Console.WriteLine()

        System.Console.WriteLine("Example 4: Compilation error (intentional)")
        System.Console.WriteLine("------------------------------------------")
        Dim code4 As String = "
Public Module TestModule
    Public Function BadCode() As Integer
        Return undefinedVariable
    End Function
End Module
"
        System.Console.WriteLine("Code to execute:")
        System.Console.WriteLine(code4)
        Dim result4 As String = VBCodeExecutor.ExecuteVBCode(code4)
        System.Console.WriteLine("Output:")
        System.Console.WriteLine(result4)
        System.Console.WriteLine()

        System.Console.WriteLine("Example 5: Runtime error (intentional)")
        System.Console.WriteLine("---------------------------------------")
        Dim code5 As String = "
Public Module TestModule
    Public Function DivideByZero() As Integer
        Dim x As Integer = 10
        Dim y As Integer = 0
        Return x \ y
    End Function
End Module
"
        System.Console.WriteLine("Code to execute:")
        System.Console.WriteLine(code5)
        Dim result5 As String = VBCodeExecutor.ExecuteVBCode(code5)
        System.Console.WriteLine("Output:")
        System.Console.WriteLine(result5)
        System.Console.WriteLine()

        System.Console.WriteLine("Example 6: Passing a single parameter")
        System.Console.WriteLine("--------------------------------------")
        Dim code6 As String = "
Public Module TestModule
    Public Function Square(x As Integer) As Integer
        Return x * x
    End Function
End Module
"
        System.Console.WriteLine("Code to execute:")
        System.Console.WriteLine(code6)
        System.Console.WriteLine("Parameters: 7")
        Dim result6 As String = VBCodeExecutor.ExecuteVBCode(code6, 7)
        System.Console.WriteLine("Output:")
        System.Console.WriteLine(result6)
        System.Console.WriteLine()

        System.Console.WriteLine("Example 7: Passing multiple parameters")
        System.Console.WriteLine("---------------------------------------")
        Dim code7 As String = "
Public Module TestModule
    Public Function Greet(name As String, age As Integer) As String
        Return $""Hello {name}, you are {age} years old!""
    End Function
End Module
"
        System.Console.WriteLine("Code to execute:")
        System.Console.WriteLine(code7)
        System.Console.WriteLine("Parameters: ""Alice"", 25")
        Dim result7 As String = VBCodeExecutor.ExecuteVBCode(code7, "Alice", 25)
        System.Console.WriteLine("Output:")
        System.Console.WriteLine(result7)
        System.Console.WriteLine()

        System.Console.WriteLine("Example 8: Parameters with calculations")
        System.Console.WriteLine("----------------------------------------")
        Dim code8 As String = "
Public Module TestModule
    Public Function CalculateSum(a As Double, b As Double, c As Double) As String
        Dim sum As Double = a + b + c
        Console.WriteLine($""Adding {a} + {b} + {c}"")
        Return $""Sum = {sum}""
    End Function
End Module
"
        System.Console.WriteLine("Code to execute:")
        System.Console.WriteLine(code8)
        System.Console.WriteLine("Parameters: 10.5, 20.3, 15.7")
        Dim result8 As String = VBCodeExecutor.ExecuteVBCode(code8, 10.5, 20.3, 15.7)
        System.Console.WriteLine("Output:")
        System.Console.WriteLine(result8)
        System.Console.WriteLine()

        System.Console.WriteLine("Example 9: Using Dictionary - Simple variables")
        System.Console.WriteLine("-----------------------------------------------")
        Dim code9 As String = "
Public Module TestModule
    Public Function Calculate() As Integer
        Return x * y + z
    End Function
End Module
"
        Dim vars9 As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"x", 10},
            {"y", 5},
            {"z", 3}
        }
        System.Console.WriteLine("Code to execute:")
        System.Console.WriteLine(code9)
        System.Console.WriteLine("Variables: x=10, y=5, z=3")
        Dim result9 As String = VBCodeExecutor.ExecuteVBCodeWithVariables(code9, vars9)
        System.Console.WriteLine("Output:")
        System.Console.WriteLine(result9)
        System.Console.WriteLine()

        System.Console.WriteLine("Example 10: Using Dictionary - Mixed types")
        System.Console.WriteLine("-------------------------------------------")
        Dim code10 As String = "
Public Module TestModule
    Public Function ProcessData() As String
        Console.WriteLine($""Name: {name}, Age: {age}, Active: {isActive}"")
        Return $""Hello {name}! You are {age} years old. Active status: {isActive}""
    End Function
End Module
"
        Dim vars10 As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"name", "Bob"},
            {"age", 30},
            {"isActive", True}
        }
        System.Console.WriteLine("Code to execute:")
        System.Console.WriteLine(code10)
        System.Console.WriteLine("Variables: name=""Bob"", age=30, isActive=True")
        Dim result10 As String = VBCodeExecutor.ExecuteVBCodeWithVariables(code10, vars10)
        System.Console.WriteLine("Output:")
        System.Console.WriteLine(result10)
        System.Console.WriteLine()

        System.Console.WriteLine("Example 11: Using Dictionary - String manipulation")
        System.Console.WriteLine("---------------------------------------------------")
        Dim code11 As String = "
Public Module TestModule
    Public Function CombineStrings() As String
        Dim result As String = firstName & "" "" & lastName
        Console.WriteLine($""Full name: {result}"")
        Return result.ToUpper()
    End Function
End Module
"
        Dim vars11 As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"firstName", "John"},
            {"lastName", "Smith"}
        }
        System.Console.WriteLine("Code to execute:")
        System.Console.WriteLine(code11)
        System.Console.WriteLine("Variables: firstName=""John"", lastName=""Smith""")
        Dim result11 As String = VBCodeExecutor.ExecuteVBCodeWithVariables(code11, vars11)
        System.Console.WriteLine("Output:")
        System.Console.WriteLine(result11)
        System.Console.WriteLine()

        System.Console.WriteLine("=== All examples completed ===")
    End Sub
End Module
