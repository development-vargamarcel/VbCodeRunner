Public Module TestMainEntry
    ' Example 1: Main function with no parameters
    Public Function Main() As String
        Console.WriteLine("Hello from Main function!")
        Console.WriteLine("This is a test VB file executed from a file path")
        Return "Main function completed successfully"
    End Function

    ' Example 2: Main function with parameters
    Public Function MainWithParams(name As String, count As Integer) As String
        Console.WriteLine($"Processing: {name}")
        For i As Integer = 1 To count
            Console.WriteLine($"  Iteration {i}")
        Next
        Return $"Processed {name} with {count} iterations"
    End Function

    ' Example 3: Main function with calculations
    Public Function Calculate(a As Double, b As Double, operation As String) As Double
        Console.WriteLine($"Calculating: {a} {operation} {b}")

        Select Case operation
            Case "+"
                Return a + b
            Case "-"
                Return a - b
            Case "*"
                Return a * b
            Case "/"
                If b <> 0 Then
                    Return a / b
                Else
                    Console.WriteLine("Error: Division by zero")
                    Return 0
                End If
            Case Else
                Console.WriteLine("Unknown operation")
                Return 0
        End Select
    End Function
End Module
