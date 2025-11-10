Imports System

Module TestFunctionEntryPoints
    Sub Main()
        Console.WriteLine("=== Test: Function Entry Points ===")
        Console.WriteLine()

        ' Test Case 1: Multiple functions - calling specific function by name
        TestMultipleFunctionsWithEntryPoint()

        ' Test Case 2: Functions with parameters
        TestFunctionWithParametersAndEntryPoint()

        ' Test Case 3: ExecuteVBCodeWithVariables with entry point
        TestVariablesWithEntryPoint()

        ' Test Case 4: Error handling - wrong function name
        TestInvalidFunctionName()

        ' Test Case 5: Error handling - wrong parameter count
        TestWrongParameterCount()

        Console.WriteLine()
        Console.WriteLine("=== All Function Entry Point Tests Completed ===")
        Console.ReadLine()
    End Sub

    Sub TestMultipleFunctionsWithEntryPoint()
        Console.WriteLine("Test 1: Multiple Functions - Calling Specific Function by Name")
        Console.WriteLine("----------------------------------------------------------------")

        Dim code As String = "
Module TestModule
    Public Function FirstFunction() As String
        Return ""This is the first function""
    End Function

    Public Function SecondFunction() As String
        Return ""This is the second function""
    End Function

    Public Function ThirdFunction() As Integer
        Return 42
    End Function
End Module
"

        ' Call FirstFunction specifically
        Console.WriteLine("Calling FirstFunction:")
        Dim result1 = VBCodeExecutor.ExecuteVBCode(code, "FirstFunction")
        Console.WriteLine(result1)
        Console.WriteLine()

        ' Call SecondFunction specifically
        Console.WriteLine("Calling SecondFunction:")
        Dim result2 = VBCodeExecutor.ExecuteVBCode(code, "SecondFunction")
        Console.WriteLine(result2)
        Console.WriteLine()

        ' Call ThirdFunction specifically
        Console.WriteLine("Calling ThirdFunction:")
        Dim result3 = VBCodeExecutor.ExecuteVBCode(code, "ThirdFunction")
        Console.WriteLine(result3)
        Console.WriteLine()
    End Sub

    Sub TestFunctionWithParametersAndEntryPoint()
        Console.WriteLine("Test 2: Functions with Parameters - Specifying Entry Point")
        Console.WriteLine("------------------------------------------------------------")

        Dim code As String = "
Module TestModule
    Public Function Add(a As Integer, b As Integer) As Integer
        Return a + b
    End Function

    Public Function Multiply(a As Integer, b As Integer) As Integer
        Return a * b
    End Function

    Public Function Greet(name As String, age As Integer) As String
        Return $""Hello {name}, you are {age} years old!""
    End Function
End Module
"

        ' Call Add function
        Console.WriteLine("Calling Add(5, 3):")
        Dim result1 = VBCodeExecutor.ExecuteVBCode(code, "Add", 5, 3)
        Console.WriteLine(result1)
        Console.WriteLine()

        ' Call Multiply function
        Console.WriteLine("Calling Multiply(5, 3):")
        Dim result2 = VBCodeExecutor.ExecuteVBCode(code, "Multiply", 5, 3)
        Console.WriteLine(result2)
        Console.WriteLine()

        ' Call Greet function
        Console.WriteLine("Calling Greet(""Alice"", 30):")
        Dim result3 = VBCodeExecutor.ExecuteVBCode(code, "Greet", "Alice", 30)
        Console.WriteLine(result3)
        Console.WriteLine()
    End Sub

    Sub TestVariablesWithEntryPoint()
        Console.WriteLine("Test 3: ExecuteVBCodeWithVariables with Entry Point")
        Console.WriteLine("-----------------------------------------------------")

        Dim code As String = "
Module TestModule
    Public Function ProcessNumbers() As String
        Dim sum As Integer = x + y
        Return $""The sum of {x} and {y} is {sum}""
    End Function

    Public Function MultiplyNumbers() As String
        Dim product As Integer = x * y
        Return $""The product of {x} and {y} is {product}""
    End Function
End Module
"

        Dim variables As New Dictionary(Of String, Object) From {
            {"x", 10},
            {"y", 5}
        }

        ' Call ProcessNumbers function
        Console.WriteLine("Calling ProcessNumbers with x=10, y=5:")
        Dim result1 = VBCodeExecutor.ExecuteVBCodeWithVariables(code, variables, "ProcessNumbers")
        Console.WriteLine(result1)
        Console.WriteLine()

        ' Call MultiplyNumbers function
        Console.WriteLine("Calling MultiplyNumbers with x=10, y=5:")
        Dim result2 = VBCodeExecutor.ExecuteVBCodeWithVariables(code, variables, "MultiplyNumbers")
        Console.WriteLine(result2)
        Console.WriteLine()
    End Sub

    Sub TestInvalidFunctionName()
        Console.WriteLine("Test 4: Error Handling - Invalid Function Name")
        Console.WriteLine("------------------------------------------------")

        Dim code As String = "
Module TestModule
    Public Function ValidFunction() As String
        Return ""This is a valid function""
    End Function
End Module
"

        ' Try to call a function that doesn't exist
        Console.WriteLine("Calling NonExistentFunction (should error):")
        Dim result = VBCodeExecutor.ExecuteVBCode(code, "NonExistentFunction")
        Console.WriteLine(result)
        Console.WriteLine()
    End Sub

    Sub TestWrongParameterCount()
        Console.WriteLine("Test 5: Error Handling - Wrong Parameter Count")
        Console.WriteLine("------------------------------------------------")

        Dim code As String = "
Module TestModule
    Public Function RequiresTwoParams(a As Integer, b As Integer) As Integer
        Return a + b
    End Function
End Module
"

        ' Try to call with wrong number of parameters
        Console.WriteLine("Calling RequiresTwoParams with 1 parameter (should error):")
        Dim result = VBCodeExecutor.ExecuteVBCode(code, "RequiresTwoParams", 5)
        Console.WriteLine(result)
        Console.WriteLine()

        ' Try to call with no parameters
        Console.WriteLine("Calling RequiresTwoParams with 0 parameters (should error):")
        Dim result2 = VBCodeExecutor.ExecuteVBCode(code, "RequiresTwoParams")
        Console.WriteLine(result2)
        Console.WriteLine()
    End Sub
End Module
