Module TestComprehensiveDataTypes
    ' Enum for testing
    Public Enum DayOfWeek
        Monday = 1
        Tuesday = 2
        Wednesday = 3
        Thursday = 4
        Friday = 5
        Saturday = 6
        Sunday = 7
    End Enum

    Public Enum Status As Byte
        Pending = 0
        Active = 1
        Completed = 2
    End Enum

    Sub Main(args As String())
        System.Console.WriteLine("=== Comprehensive Data Type Tests ===")
        System.Console.WriteLine()

        ' Test 1: All integer types (signed and unsigned)
        System.Console.WriteLine("Test 1: All Integer Types")
        System.Console.WriteLine("--------------------------")
        Dim code1 As String = "
Public Module TestModule
    Public Function TestIntegers() As String
        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine($""SByte: {sbyteVal} (Type: {sbyteVal.GetType().Name})"")
        sb.AppendLine($""Byte: {byteVal} (Type: {byteVal.GetType().Name})"")
        sb.AppendLine($""Short: {shortVal} (Type: {shortVal.GetType().Name})"")
        sb.AppendLine($""UShort: {ushortVal} (Type: {ushortVal.GetType().Name})"")
        sb.AppendLine($""Integer: {intVal} (Type: {intVal.GetType().Name})"")
        sb.AppendLine($""UInteger: {uintVal} (Type: {uintVal.GetType().Name})"")
        sb.AppendLine($""Long: {longVal} (Type: {longVal.GetType().Name})"")
        sb.AppendLine($""ULong: {ulongVal} (Type: {ulongVal.GetType().Name})"")
        Return sb.ToString()
    End Function
End Module
"
        Dim vars1 As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"sbyteVal", CSByte(-100)},
            {"byteVal", CByte(200)},
            {"shortVal", CShort(-30000)},
            {"ushortVal", CUShort(60000)},
            {"intVal", -2000000},
            {"uintVal", CUInt(4000000)},
            {"longVal", -9000000000000L},
            {"ulongVal", CULng(18000000000000UL)}
        }
        Dim result1 As String = VBCodeExecutor.ExecuteVBCodeWithVariables(code1, vars1)
        System.Console.WriteLine(result1)
        System.Console.WriteLine()

        ' Test 2: Char type
        System.Console.WriteLine("Test 2: Char Type")
        System.Console.WriteLine("------------------")
        Dim code2 As String = "
Public Module TestModule
    Public Function TestChar() As String
        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine($""Char1: '{char1}' (ASCII: {Microsoft.VisualBasic.AscW(char1)})"")
        sb.AppendLine($""Char2: '{char2}' (Unicode)"")
        sb.AppendLine($""Char3: '{char3}'"")
        Return sb.ToString()
    End Function
End Module
"
        Dim vars2 As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"char1", "A"c},
            {"char2", ChrW(&H2665)}, ' Heart symbol
            {"char3", " "c}
        }
        Dim result2 As String = VBCodeExecutor.ExecuteVBCodeWithVariables(code2, vars2)
        System.Console.WriteLine(result2)
        System.Console.WriteLine()

        ' Test 3: Special floating point values
        System.Console.WriteLine("Test 3: Special Floating Point Values")
        System.Console.WriteLine("--------------------------------------")
        Dim code3 As String = "
Public Module TestModule
    Public Function TestSpecialFloats() As String
        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine($""Regular Double: {regularDouble}"")
        sb.AppendLine($""NaN Double: {nanDouble} (IsNaN: {Double.IsNaN(nanDouble)})"")
        sb.AppendLine($""Positive Infinity: {posInf} (IsInfinity: {Double.IsPositiveInfinity(posInf)})"")
        sb.AppendLine($""Negative Infinity: {negInf} (IsInfinity: {Double.IsNegativeInfinity(negInf)})"")
        sb.AppendLine($""Regular Single: {regularSingle}"")
        sb.AppendLine($""NaN Single: {nanSingle} (IsNaN: {Single.IsNaN(nanSingle)})"")
        Return sb.ToString()
    End Function
End Module
"
        Dim vars3 As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"regularDouble", 3.14159},
            {"nanDouble", Double.NaN},
            {"posInf", Double.PositiveInfinity},
            {"negInf", Double.NegativeInfinity},
            {"regularSingle", 2.71828F},
            {"nanSingle", Single.NaN}
        }
        Dim result3 As String = VBCodeExecutor.ExecuteVBCodeWithVariables(code3, vars3)
        System.Console.WriteLine(result3)
        System.Console.WriteLine()

        ' Test 4: DateTimeOffset
        System.Console.WriteLine("Test 4: DateTimeOffset")
        System.Console.WriteLine("-----------------------")
        Dim code4 As String = "
Public Module TestModule
    Public Function TestDateTimeOffset() As String
        Dim sb As New System.Text.StringBuilder()
        Dim dto As DateTimeOffset = DirectCast(dtoVal, DateTimeOffset)
        sb.AppendLine($""DateTimeOffset: {dto.ToString(""yyyy-MM-dd HH:mm:ss zzz"")}"")
        sb.AppendLine($""UTC: {dto.UtcDateTime.ToString(""yyyy-MM-dd HH:mm:ss"")}"")
        sb.AppendLine($""Offset: {dto.Offset.TotalHours} hours"")
        Return sb.ToString()
    End Function
End Module
"
        Dim vars4 As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"dtoVal", New DateTimeOffset(2025, 11, 10, 14, 30, 0, New TimeSpan(-5, 0, 0))}
        }
        Dim result4 As String = VBCodeExecutor.ExecuteVBCodeWithVariables(code4, vars4)
        System.Console.WriteLine(result4)
        System.Console.WriteLine()

        ' Test 5: DBNull
        System.Console.WriteLine("Test 5: DBNull")
        System.Console.WriteLine("--------------")
        Dim code5 As String = "
Public Module TestModule
    Public Function TestDBNull() As String
        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine($""dbValue is DBNull: {TypeOf dbValue Is DBNull}"")
        sb.AppendLine($""dbValue equals DBNull.Value: {Object.Equals(dbValue, DBNull.Value)}"")
        Return sb.ToString()
    End Function
End Module
"
        Dim vars5 As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"dbValue", DBNull.Value}
        }
        Dim result5 As String = VBCodeExecutor.ExecuteVBCodeWithVariables(code5, vars5)
        System.Console.WriteLine(result5)
        System.Console.WriteLine()

        ' Test 6: Enum types
        System.Console.WriteLine("Test 6: Enum Types")
        System.Console.WriteLine("------------------")
        Dim code6 As String = "
Public Module TestModule
    Public Function TestEnums() As String
        Dim sb As New System.Text.StringBuilder()
        Dim day As TestComprehensiveDataTypes.DayOfWeek = DirectCast(dayVal, TestComprehensiveDataTypes.DayOfWeek)
        Dim status As TestComprehensiveDataTypes.Status = DirectCast(statusVal, TestComprehensiveDataTypes.Status)
        sb.AppendLine($""Day: {day} (Value: {CInt(day)})"")
        sb.AppendLine($""Status: {status} (Value: {CByte(status)})"")
        Return sb.ToString()
    End Function
End Module
"
        Dim vars6 As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"dayVal", DayOfWeek.Wednesday},
            {"statusVal", Status.Active}
        }
        Dim result6 As String = VBCodeExecutor.ExecuteVBCodeWithVariables(code6, vars6)
        System.Console.WriteLine(result6)
        System.Console.WriteLine()

        ' Test 7: DateTime with different DateTimeKind values
        System.Console.WriteLine("Test 7: DateTime with DateTimeKind")
        System.Console.WriteLine("-----------------------------------")
        Dim code7 As String = "
Public Module TestModule
    Public Function TestDateTimeKind() As String
        Dim sb As New System.Text.StringBuilder()
        Dim utcDate As DateTime = DirectCast(utcDateTime, DateTime)
        Dim localDate As DateTime = DirectCast(localDateTime, DateTime)
        Dim unspecDate As DateTime = DirectCast(unspecDateTime, DateTime)
        sb.AppendLine($""UTC: {utcDate.ToString(""yyyy-MM-dd HH:mm:ss"")} (Kind: {utcDate.Kind})"")
        sb.AppendLine($""Local: {localDate.ToString(""yyyy-MM-dd HH:mm:ss"")} (Kind: {localDate.Kind})"")
        sb.AppendLine($""Unspecified: {unspecDate.ToString(""yyyy-MM-dd HH:mm:ss"")} (Kind: {unspecDate.Kind})"")
        Return sb.ToString()
    End Function
End Module
"
        Dim vars7 As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"utcDateTime", DateTime.SpecifyKind(New DateTime(2025, 11, 10, 12, 0, 0), DateTimeKind.Utc)},
            {"localDateTime", DateTime.SpecifyKind(New DateTime(2025, 11, 10, 12, 0, 0), DateTimeKind.Local)},
            {"unspecDateTime", New DateTime(2025, 11, 10, 12, 0, 0)}
        }
        Dim result7 As String = VBCodeExecutor.ExecuteVBCodeWithVariables(code7, vars7)
        System.Console.WriteLine(result7)
        System.Console.WriteLine()

        ' Test 8: Multi-dimensional arrays
        System.Console.WriteLine("Test 8: Multi-dimensional Arrays")
        System.Console.WriteLine("--------------------------------")
        Dim code8 As String = "
Public Module TestModule
    Public Function TestMultiDimArrays() As String
        Dim sb As New System.Text.StringBuilder()
        Dim matrix As Integer(,) = DirectCast(matrix2d, Integer(,))
        Dim cube As Integer(,,) = DirectCast(matrix3d, Integer(,,))

        sb.AppendLine($""2D Array dimensions: {matrix.GetLength(0)} x {matrix.GetLength(1)}"")
        sb.AppendLine($""2D Array [1,1]: {matrix(1, 1)}"")
        sb.AppendLine($""3D Array dimensions: {cube.GetLength(0)} x {cube.GetLength(1)} x {cube.GetLength(2)}"")
        sb.AppendLine($""3D Array [0,1,1]: {cube(0, 1, 1)}"")

        Return sb.ToString()
    End Function
End Module
"
        Dim matrix2d(,) As Integer = {{1, 2, 3}, {4, 5, 6}}
        Dim matrix3d(,,) As Integer = {{{1, 2}, {3, 4}}, {{5, 6}, {7, 8}}}

        Dim vars8 As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"matrix2d", matrix2d},
            {"matrix3d", matrix3d}
        }
        Dim result8 As String = VBCodeExecutor.ExecuteVBCodeWithVariables(code8, vars8)
        System.Console.WriteLine(result8)
        System.Console.WriteLine()

        ' Test 9: Edge cases - Min/Max values
        System.Console.WriteLine("Test 9: Min/Max Edge Cases")
        System.Console.WriteLine("---------------------------")
        Dim code9 As String = "
Public Module TestModule
    Public Function TestMinMax() As String
        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine($""SByte Min: {sbyteMin}, Max: {sbyteMax}"")
        sb.AppendLine($""Byte Min: {byteMin}, Max: {byteMax}"")
        sb.AppendLine($""Int Min: {intMin}, Max: {intMax}"")
        sb.AppendLine($""ULong Max: {ulongMax}"")
        sb.AppendLine($""Decimal Min: {decMin}"")
        sb.AppendLine($""Decimal Max: {decMax}"")
        Return sb.ToString()
    End Function
End Module
"
        Dim vars9 As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"sbyteMin", SByte.MinValue},
            {"sbyteMax", SByte.MaxValue},
            {"byteMin", Byte.MinValue},
            {"byteMax", Byte.MaxValue},
            {"intMin", Integer.MinValue},
            {"intMax", Integer.MaxValue},
            {"ulongMax", ULong.MaxValue},
            {"decMin", Decimal.MinValue},
            {"decMax", Decimal.MaxValue}
        }
        Dim result9 As String = VBCodeExecutor.ExecuteVBCodeWithVariables(code9, vars9)
        System.Console.WriteLine(result9)
        System.Console.WriteLine()

        ' Test 10: Mixed type scenarios
        System.Console.WriteLine("Test 10: Mixed Type Scenarios")
        System.Console.WriteLine("-----------------------------")
        Dim code10 As String = "
Public Module TestModule
    Public Function TestMixedTypes() As String
        Dim sb As New System.Text.StringBuilder()

        ' Arithmetic with different integer types
        Dim sum As Long = CLng(byteVal) + CLng(shortVal) + intVal + longVal
        sb.AppendLine($""Sum of mixed integers: {sum}"")

        ' String concatenation with char
        Dim fullStr As String = charVal & stringVal
        sb.AppendLine($""Char + String: {fullStr}"")

        ' Date arithmetic
        Dim newDate As DateTime = DirectCast(dateVal, DateTime).AddDays(10)
        sb.AppendLine($""Date + 10 days: {newDate.ToString(""yyyy-MM-dd"")}"")

        ' Decimal precision
        Dim preciseCalc As Decimal = decVal * 1.5D
        sb.AppendLine($""Decimal calculation: {preciseCalc}"")

        Return sb.ToString()
    End Function
End Module
"
        Dim vars10 As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"byteVal", CByte(100)},
            {"shortVal", CShort(1000)},
            {"intVal", 10000},
            {"longVal", 100000L},
            {"charVal", "X"c},
            {"stringVal", "YZ"},
            {"dateVal", New DateTime(2025, 1, 1)},
            {"decVal", 123.456D}
        }
        Dim result10 As String = VBCodeExecutor.ExecuteVBCodeWithVariables(code10, vars10)
        System.Console.WriteLine(result10)
        System.Console.WriteLine()

        System.Console.WriteLine("=== All comprehensive tests completed successfully! ===")
    End Sub
End Module
