Imports System

Module Program
    Sub Main(args As String())
        Console.WriteLine("=== VB Code Executor using Roslyn ===")
        Console.WriteLine()
        
        Console.WriteLine("Example 1: Simple calculation")
        Console.WriteLine("------------------------------")
        Dim code1 As String = "
Public Module TestModule
    Public Function Calculate() As Integer
        Return 5 + 10 * 2
    End Function
End Module
"
        Console.WriteLine("Code to execute:")
        Console.WriteLine(code1)
        Dim result1 As String = VBCodeExecutor.ExecuteVBCode(code1)
        Console.WriteLine("Output:")
        Console.WriteLine(result1)
        Console.WriteLine()
        
        Console.WriteLine("Example 2: String manipulation with Console output")
        Console.WriteLine("---------------------------------------------------")
        Dim code2 As String = "
Public Module TestModule
    Public Function ProcessString() As String
        Dim text As String = ""Hello from dynamic VB code!""
        Console.WriteLine(""Processing: "" & text)
        Return text.ToUpper()
    End Function
End Module
"
        Console.WriteLine("Code to execute:")
        Console.WriteLine(code2)
        Dim result2 As String = VBCodeExecutor.ExecuteVBCode(code2)
        Console.WriteLine("Output:")
        Console.WriteLine(result2)
        Console.WriteLine()
        
        Console.WriteLine("Example 3: Loop with output")
        Console.WriteLine("---------------------------")
        Dim code3 As String = "
Public Module TestModule
    Public Sub PrintNumbers()
        For i As Integer = 1 To 5
            Console.WriteLine(""Number: "" & i)
        Next
    End Sub
End Module
"
        Console.WriteLine("Code to execute:")
        Console.WriteLine(code3)
        Dim result3 As String = VBCodeExecutor.ExecuteVBCode(code3)
        Console.WriteLine("Output:")
        Console.WriteLine(result3)
        Console.WriteLine()
        
        Console.WriteLine("Example 4: Compilation error (intentional)")
        Console.WriteLine("------------------------------------------")
        Dim code4 As String = "
Public Module TestModule
    Public Function BadCode() As Integer
        Return undefinedVariable
    End Function
End Module
"
        Console.WriteLine("Code to execute:")
        Console.WriteLine(code4)
        Dim result4 As String = VBCodeExecutor.ExecuteVBCode(code4)
        Console.WriteLine("Output:")
        Console.WriteLine(result4)
        Console.WriteLine()
        
        Console.WriteLine("Example 5: Runtime error (intentional)")
        Console.WriteLine("---------------------------------------")
        Dim code5 As String = "
Public Module TestModule
    Public Function DivideByZero() As Integer
        Dim x As Integer = 10
        Dim y As Integer = 0
        Return x \ y
    End Function
End Module
"
        Console.WriteLine("Code to execute:")
        Console.WriteLine(code5)
        Dim result5 As String = VBCodeExecutor.ExecuteVBCode(code5)
        Console.WriteLine("Output:")
        Console.WriteLine(result5)
        Console.WriteLine()
        
        Console.WriteLine("Example 6: Passing a single parameter")
        Console.WriteLine("--------------------------------------")
        Dim code6 As String = "
Public Module TestModule
    Public Function Square(x As Integer) As Integer
        Return x * x
    End Function
End Module
"
        Console.WriteLine("Code to execute:")
        Console.WriteLine(code6)
        Console.WriteLine("Parameters: 7")
        Dim result6 As String = VBCodeExecutor.ExecuteVBCode(code6, 7)
        Console.WriteLine("Output:")
        Console.WriteLine(result6)
        Console.WriteLine()
        
        Console.WriteLine("Example 7: Passing multiple parameters")
        Console.WriteLine("---------------------------------------")
        Dim code7 As String = "
Public Module TestModule
    Public Function Greet(name As String, age As Integer) As String
        Return $""Hello {name}, you are {age} years old!""
    End Function
End Module
"
        Console.WriteLine("Code to execute:")
        Console.WriteLine(code7)
        Console.WriteLine("Parameters: ""Alice"", 25")
        Dim result7 As String = VBCodeExecutor.ExecuteVBCode(code7, "Alice", 25)
        Console.WriteLine("Output:")
        Console.WriteLine(result7)
        Console.WriteLine()
        
        Console.WriteLine("Example 8: Parameters with calculations")
        Console.WriteLine("----------------------------------------")
        Dim code8 As String = "
Public Module TestModule
    Public Function CalculateSum(a As Double, b As Double, c As Double) As String
        Dim sum As Double = a + b + c
        Console.WriteLine($""Adding {a} + {b} + {c}"")
        Return $""Sum = {sum}""
    End Function
End Module
"
        Console.WriteLine("Code to execute:")
        Console.WriteLine(code8)
        Console.WriteLine("Parameters: 10.5, 20.3, 15.7")
        Dim result8 As String = VBCodeExecutor.ExecuteVBCode(code8, 10.5, 20.3, 15.7)
        Console.WriteLine("Output:")
        Console.WriteLine(result8)
        Console.WriteLine()
        
        Console.WriteLine("=== All examples completed ===")
    End Sub
End Module
