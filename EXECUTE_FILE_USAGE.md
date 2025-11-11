# Execute VB File with Main Function

This document describes how to use the new file execution functionality in VBCodeExecutor.

## Overview

You can now execute VB code files directly by providing a file path and calling the Main function from that file.

## New Methods

### 1. ExecuteVBFile(filePath As String, ParamArray parameters As Object())

Executes a VB file and calls the Main function without any custom logger.

**Example:**
```vb
Dim result As String = VBCodeExecutor.ExecuteVBFile("MyScript.vb")
Console.WriteLine(result)
```

### 2. ExecuteVBFile(filePath As String, customLogger As Action(Of String), ParamArray parameters As Object())

Executes a VB file and calls the Main function with a custom logger.

**Example:**
```vb
Dim logFilePath As String = "output.log"
Dim logger As Action(Of String) = Sub(msg) File.AppendAllText(logFilePath, msg)
Dim result As String = VBCodeExecutor.ExecuteVBFile("MyScript.vb", logger)
Console.WriteLine(result)
```

### 3. ExecuteVBFileWithVariables(filePath As String, variables As Dictionary(Of String, Object), Optional customLogger As Action(Of String) = Nothing)

Executes a VB file with injected variables and calls the Main function.

**Example:**
```vb
Dim vars As New Dictionary(Of String, Object) From {
    {"userName", "Alice"},
    {"userAge", 30}
}
Dim result As String = VBCodeExecutor.ExecuteVBFileWithVariables("MyScript.vb", vars)
Console.WriteLine(result)
```

## Requirements for VB Files

The VB file should contain:

1. A public module
2. A public function named "Main" (case-insensitive)
3. The Main function can:
   - Accept no parameters
   - Accept parameters (must be passed via the parameters argument)
   - Return any type (will be converted to string in output)

## Example VB File

**MyScript.vb:**
```vb
Public Module MyScript
    Public Function Main() As String
        Console.WriteLine("Hello from Main function!")
        Console.WriteLine("Processing data...")
        Return "Task completed successfully"
    End Function
End Module
```

## Example VB File with Parameters

**MyCalculator.vb:**
```vb
Public Module MyCalculator
    Public Function Main(a As Double, b As Double, operation As String) As String
        Select Case operation
            Case "+"
                Return $"Result: {a + b}"
            Case "-"
                Return $"Result: {a - b}"
            Case "*"
                Return $"Result: {a * b}"
            Case "/"
                If b <> 0 Then
                    Return $"Result: {a / b}"
                Else
                    Return "Error: Division by zero"
                End If
            Case Else
                Return "Unknown operation"
        End Select
    End Function
End Module
```

**Usage:**
```vb
' Read the file content first if you need to pass parameters
Dim fileCode As String = File.ReadAllText("MyCalculator.vb")
Dim result As String = VBCodeExecutor.ExecuteVBCode(fileCode, "Main", 10.5, 5.0, "+")
Console.WriteLine(result)  ' Output: Result: 15.5
```

## Example VB File with Variables

**MyVariableScript.vb:**
```vb
Public Module MyVariableScript
    Public Function Main() As String
        Console.WriteLine($"User: {userName}")
        Console.WriteLine($"Age: {userAge}")
        Return $"Hello {userName}, age {userAge}"
    End Function
End Module
```

**Usage:**
```vb
Dim vars As New Dictionary(Of String, Object) From {
    {"userName", "Bob"},
    {"userAge", 25}
}
Dim result As String = VBCodeExecutor.ExecuteVBFileWithVariables("MyVariableScript.vb", vars)
Console.WriteLine(result)
```

## Features

- ✅ Execute VB files from disk
- ✅ Automatic Main function detection
- ✅ Support for custom loggers
- ✅ Variable injection support
- ✅ Parameter passing support
- ✅ Error handling and detailed error messages
- ✅ File existence validation

## Error Handling

The methods provide detailed error messages for common issues:
- File not found
- Compilation errors
- Runtime errors
- Missing Main function
- Parameter count mismatches
