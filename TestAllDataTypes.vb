Module TestAllDataTypes
    ' Define a custom class for testing
    Public Class Person
        Public Property Name As String
        Public Property Age As Integer
        Public Property Email As String

        Public Sub New(name As String, age As Integer, email As String)
            Me.Name = name
            Me.Age = age
            Me.Email = email
        End Sub

        Public Overrides Function ToString() As String
            Return $"Person: {Name}, Age: {Age}, Email: {Email}"
        End Function
    End Class

    Public Structure Point
        Public X As Double
        Public Y As Double

        Public Sub New(x As Double, y As Double)
            Me.X = x
            Me.Y = y
        End Sub

        Public Overrides Function ToString() As String
            Return $"Point({X}, {Y})"
        End Function
    End Structure

    Sub Main(args As String())
        System.Console.WriteLine("=== Testing All Data Types with Variable Dictionary ===")
        System.Console.WriteLine()

        ' Test 1: Primitive types (Integer, String, Boolean, Double)
        System.Console.WriteLine("Test 1: Primitive Types")
        System.Console.WriteLine("------------------------")
        Dim code1 As String = "
Public Module TestModule
    Public Function TestPrimitives() As String
        Return $""Int: {intVal}, String: {strVal}, Bool: {boolVal}, Double: {dblVal}""
    End Function
End Module
"
        Dim vars1 As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"intVal", 42},
            {"strVal", "Hello World"},
            {"boolVal", True},
            {"dblVal", 3.14159}
        }
        Dim result1 As String = VBCodeExecutor.ExecuteVBCodeWithVariables(code1, vars1)
        System.Console.WriteLine("Result: " & result1)
        System.Console.WriteLine()

        ' Test 2: DateTime, TimeSpan, Guid
        System.Console.WriteLine("Test 2: DateTime, TimeSpan, and Guid")
        System.Console.WriteLine("-------------------------------------")
        Dim code2 As String = "
Public Module TestModule
    Public Function TestDateTimeTypes() As String
        Dim result As New System.Text.StringBuilder()
        result.AppendLine($""Date: {currentDate.ToString(""yyyy-MM-dd HH:mm:ss"")}"")
        result.AppendLine($""TimeSpan: {duration.TotalHours} hours"")
        result.AppendLine($""Guid: {uniqueId.ToString()}"")
        Return result.ToString()
    End Function
End Module
"
        Dim vars2 As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"currentDate", New System.DateTime(2025, 11, 10, 14, 30, 0)},
            {"duration", New System.TimeSpan(2, 30, 45)},
            {"uniqueId", System.Guid.NewGuid()}
        }
        Dim result2 As String = VBCodeExecutor.ExecuteVBCodeWithVariables(code2, vars2)
        System.Console.WriteLine("Result:")
        System.Console.WriteLine(result2)
        System.Console.WriteLine()

        ' Test 3: Arrays
        System.Console.WriteLine("Test 3: Arrays")
        System.Console.WriteLine("--------------")
        Dim code3 As String = "
Public Module TestModule
    Public Function TestArrays() As String
        Dim numbers As Integer() = DirectCast(numberArray, Integer())
        Dim names As String() = DirectCast(nameArray, String())

        Dim result As New System.Text.StringBuilder()
        result.AppendLine($""Numbers count: {numbers.Length}"")
        result.AppendLine($""Sum: {numbers.Sum()}"")
        result.AppendLine($""Names: {String.Join("", "", names)}"")
        Return result.ToString()
    End Function
End Module
"
        Dim vars3 As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"numberArray", New Integer() {1, 2, 3, 4, 5, 10, 20, 30}},
            {"nameArray", New String() {"Alice", "Bob", "Charlie"}}
        }
        Dim result3 As String = VBCodeExecutor.ExecuteVBCodeWithVariables(code3, vars3)
        System.Console.WriteLine("Result:")
        System.Console.WriteLine(result3)
        System.Console.WriteLine()

        ' Test 4: List collections
        System.Console.WriteLine("Test 4: List Collections")
        System.Console.WriteLine("------------------------")
        Dim code4 As String = "
Public Module TestModule
    Public Function TestLists() As String
        Dim intList As List(Of Integer) = DirectCast(myIntList, List(Of Integer))
        Dim strList As List(Of String) = DirectCast(myStrList, List(Of String))

        Dim result As New System.Text.StringBuilder()
        result.AppendLine($""Int list count: {intList.Count}"")
        result.AppendLine($""Max value: {intList.Max()}"")
        result.AppendLine($""String list: {String.Join("" | "", strList)}"")
        Return result.ToString()
    End Function
End Module
"
        Dim intList As New System.Collections.Generic.List(Of Integer) From {10, 20, 30, 40, 50}
        Dim strList As New System.Collections.Generic.List(Of String) From {"Apple", "Banana", "Cherry"}

        Dim vars4 As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"myIntList", intList},
            {"myStrList", strList}
        }
        Dim result4 As String = VBCodeExecutor.ExecuteVBCodeWithVariables(code4, vars4)
        System.Console.WriteLine("Result:")
        System.Console.WriteLine(result4)
        System.Console.WriteLine()

        ' Test 5: Dictionary collections
        System.Console.WriteLine("Test 5: Dictionary Collections")
        System.Console.WriteLine("------------------------------")
        Dim code5 As String = "
Public Module TestModule
    Public Function TestDictionaries() As String
        Dim scores As Dictionary(Of String, Integer) = DirectCast(scoreDict, Dictionary(Of String, Integer))
        Dim settings As Dictionary(Of String, String) = DirectCast(settingsDict, Dictionary(Of String, String))

        Dim result As New System.Text.StringBuilder()
        result.AppendLine($""Scores count: {scores.Count}"")
        result.AppendLine($""Alice's score: {scores(""Alice"")}"")
        result.AppendLine($""Theme: {settings(""theme"")}"")
        Return result.ToString()
    End Function
End Module
"
        Dim scoreDict As New System.Collections.Generic.Dictionary(Of String, Integer) From {
            {"Alice", 95},
            {"Bob", 87},
            {"Charlie", 92}
        }
        Dim settingsDict As New System.Collections.Generic.Dictionary(Of String, String) From {
            {"theme", "dark"},
            {"language", "en-US"}
        }

        Dim vars5 As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"scoreDict", scoreDict},
            {"settingsDict", settingsDict}
        }
        Dim result5 As String = VBCodeExecutor.ExecuteVBCodeWithVariables(code5, vars5)
        System.Console.WriteLine("Result:")
        System.Console.WriteLine(result5)
        System.Console.WriteLine()

        ' Test 6: Custom class objects
        System.Console.WriteLine("Test 6: Custom Class Objects")
        System.Console.WriteLine("----------------------------")
        Dim code6 As String = "
Public Module TestModule
    Public Function TestCustomObjects() As String
        Dim p As TestAllDataTypes.Person = DirectCast(person, TestAllDataTypes.Person)
        Dim pt As TestAllDataTypes.Point = DirectCast(point, TestAllDataTypes.Point)

        Dim result As New System.Text.StringBuilder()
        result.AppendLine($""Person: {p.Name}, {p.Age} years old"")
        result.AppendLine($""Email: {p.Email}"")
        result.AppendLine($""Point: ({pt.X}, {pt.Y})"")
        Return result.ToString()
    End Function
End Module
"
        Dim person As New Person("John Doe", 35, "john@example.com")
        Dim point As New Point(10.5, 20.75)

        Dim vars6 As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"person", person},
            {"point", point}
        }
        Dim result6 As String = VBCodeExecutor.ExecuteVBCodeWithVariables(code6, vars6)
        System.Console.WriteLine("Result:")
        System.Console.WriteLine(result6)
        System.Console.WriteLine()

        ' Test 7: Complex nested structures
        System.Console.WriteLine("Test 7: Complex Nested Structures")
        System.Console.WriteLine("----------------------------------")
        Dim code7 As String = "
Public Module TestModule
    Public Function TestComplexStructures() As String
        Dim peopleList As List(Of TestAllDataTypes.Person) = DirectCast(people, List(Of TestAllDataTypes.Person))
        Dim dataDict As Dictionary(Of String, Object) = DirectCast(data, Dictionary(Of String, Object))

        Dim result As New System.Text.StringBuilder()
        result.AppendLine($""People count: {peopleList.Count}"")
        For Each p In peopleList
            result.AppendLine($""  - {p.Name} ({p.Age})"")
        Next

        result.AppendLine($""Data keys: {String.Join("", "", dataDict.Keys)}"")
        result.AppendLine($""Config value: {dataDict(""config"")}"")
        Return result.ToString()
    End Function
End Module
"
        Dim peopleList As New System.Collections.Generic.List(Of Person) From {
            New Person("Alice", 28, "alice@example.com"),
            New Person("Bob", 34, "bob@example.com"),
            New Person("Carol", 31, "carol@example.com")
        }

        Dim dataDict As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"config", "production"},
            {"version", 1.5},
            {"enabled", True}
        }

        Dim vars7 As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"people", peopleList},
            {"data", dataDict}
        }
        Dim result7 As String = VBCodeExecutor.ExecuteVBCodeWithVariables(code7, vars7)
        System.Console.WriteLine("Result:")
        System.Console.WriteLine(result7)
        System.Console.WriteLine()

        ' Test 8: Null/Nothing values
        System.Console.WriteLine("Test 8: Null/Nothing Values")
        System.Console.WriteLine("---------------------------")
        Dim code8 As String = "
Public Module TestModule
    Public Function TestNullValues() As String
        Dim result As New System.Text.StringBuilder()
        result.AppendLine($""nullValue is Nothing: {nullValue Is Nothing}"")
        result.AppendLine($""validValue: {validValue}"")
        Return result.ToString()
    End Function
End Module
"
        Dim vars8 As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"nullValue", Nothing},
            {"validValue", "I'm not null!"}
        }
        Dim result8 As String = VBCodeExecutor.ExecuteVBCodeWithVariables(code8, vars8)
        System.Console.WriteLine("Result:")
        System.Console.WriteLine(result8)
        System.Console.WriteLine()

        System.Console.WriteLine("=== All tests completed successfully! ===")
    End Sub
End Module
