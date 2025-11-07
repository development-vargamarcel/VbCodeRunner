# VB Code Executor using Roslyn

## Overview
A VB.NET application that provides a function to dynamically compile and execute VB code from a string using the Roslyn compiler. The function returns either the execution result or any compilation/runtime errors as a string.

## Project Structure
- **VBCodeExecutor.vb** - Main module containing the `ExecuteVBCode` and `ExecuteVBCodeWithVariables` functions
- **Program.vb** - Demo application with 11 example scenarios
- **VBCodeExecutor.vbproj** - Project file with Roslyn package references

## Key Features
- Dynamic VB code compilation using Roslyn
- **Pass parameters to dynamic code** - Support for any number of parameters of any type
- **Pass variables via Dictionary** - Each dictionary entry becomes a variable accessible in the code
- Automatic global imports (System, System.Collections.Generic, System.Linq)
- Comprehensive error handling for both compilation and runtime errors
- Captures Console output from executed code
- Returns formatted results or error messages

## Dependencies
- .NET 8.0
- Microsoft.CodeAnalysis.VisualBasic 4.14.0

## Usage Examples

### Basic Usage (No Parameters)
```vb
Dim vbCode As String = "
Public Module TestModule
    Public Function Calculate() As Integer
        Return 5 + 10 * 2
    End Function
End Module
"

Dim result As String = VBCodeExecutor.ExecuteVBCode(vbCode)
Console.WriteLine(result)  ' Output: "Result: 25"
```

### With Parameters
```vb
Dim vbCode As String = "
Public Module TestModule
    Public Function Greet(name As String, age As Integer) As String
        Return $""Hello {name}, you are {age} years old!""
    End Function
End Module
"

Dim result As String = VBCodeExecutor.ExecuteVBCode(vbCode, "Alice", 25)
Console.WriteLine(result)  ' Output: "Result: Hello Alice, you are 25 years old!"
```

### With Dictionary Variables
```vb
Dim vbCode As String = "
Public Module TestModule
    Public Function Calculate() As Integer
        Return x * y + z
    End Function
End Module
"

Dim vars As New Dictionary(Of String, Object) From {
    {"x", 10},
    {"y", 5},
    {"z", 3}
}

Dim result As String = VBCodeExecutor.ExecuteVBCodeWithVariables(vbCode, vars)
Console.WriteLine(result)  ' Output: "Result: 53"
```

## Function Signatures

### ExecuteVBCode - Method Parameters
```vb
Public Function ExecuteVBCode(vbCodeString As String, ParamArray parameters As Object()) As String
```

**Parameters:**
- `vbCodeString` - A string containing valid VB.NET code with a Module containing at least one public shared function or sub
- `parameters` (Optional) - Variable number of parameters to pass to the dynamic code's function

**Returns:**
- Success: Returns the execution result or Console output
- Compilation Error: Returns formatted compilation errors
- Runtime Error: Returns runtime exception details with stack trace
- Parameter Mismatch: Returns error if method with matching parameter count is not found

### ExecuteVBCodeWithVariables - Dictionary Variables
```vb
Public Function ExecuteVBCodeWithVariables(vbCodeString As String, variables As Dictionary(Of String, Object)) As String
```

**Parameters:**
- `vbCodeString` - A string containing valid VB.NET code with a Module
- `variables` - Dictionary where each key becomes a variable name accessible in the code, and the value becomes its value

**Returns:**
- Success: Returns the execution result or Console output
- Compilation Error: Returns formatted compilation errors
- Runtime Error: Returns runtime exception details with stack trace

## Recent Changes
- 2025-11-07: Added dictionary-based variable injection - each dictionary entry becomes a variable
- 2025-11-07: Added parameter passing support - can now pass variables to dynamic code
- 2025-11-07: Initial implementation with Roslyn compiler support
- Added global imports for System, System.Collections.Generic, and System.Linq
- Configured Microsoft.VisualBasic.Core assembly reference for VB runtime support
- Created comprehensive demo with 11 example scenarios (including parameter and dictionary examples)

## Architecture
The function works by:
1. Parsing the VB code string into a syntax tree
2. Creating a compilation with necessary assembly references and global imports
3. Emitting the compiled assembly to memory
4. Dynamically loading and invoking the first public static method found
5. Capturing Console output and return values
6. Handling compilation and runtime errors gracefully

## Notes
- Dynamic code must contain a Module with at least one public shared function or sub
- **Method Parameters**: When using `ExecuteVBCode` with parameters, the function matches methods with the same parameter count
- **Dictionary Variables**: When using `ExecuteVBCodeWithVariables`, dictionary keys become variable names accessible in the code
- Supported dictionary value types: Integer, String, Double, Boolean, Decimal, and other primitives
- Global imports are automatically added (System, System.Collections.Generic, System.Linq)
- Console output is captured and included in the result
- All errors are returned as formatted strings for easy debugging
