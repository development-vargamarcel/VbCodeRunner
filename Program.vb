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
        
        Console.WriteLine("=== All examples completed ===")
    End Sub
End Module
