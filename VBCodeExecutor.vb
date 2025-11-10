Public Module VBCodeExecutor

    ''' <summary>
    ''' Thread-safe storage for complex objects that cannot be represented as VB literals
    ''' Key format: "{executionId}_{variableName}"
    ''' </summary>
    Private ReadOnly _objectStore As New System.Collections.Concurrent.ConcurrentDictionary(Of String, Object)()

    ''' <summary>
    ''' Public accessor for dynamically compiled code to retrieve stored objects
    ''' </summary>
    Public Function GetStoredObject(key As String) As Object
        Dim value As Object = Nothing
        If _objectStore.TryGetValue(key, value) Then
            Return value
        End If
        Return Nothing
    End Function

    ''' <summary>
    ''' Custom TextWriter that redirects output to an Action callback
    ''' </summary>
    Private Class ActionTextWriter
        Inherits System.IO.TextWriter

        Private ReadOnly _logAction As System.Action(Of String)
        Private ReadOnly _buffer As New System.Text.StringBuilder()

        Public Sub New(logAction As System.Action(Of String))
            _logAction = logAction
        End Sub

        Public Overrides Sub Write(value As Char)
            _buffer.Append(value)
            If value = Microsoft.VisualBasic.ControlChars.Lf Then
                Flush()
            End If
        End Sub

        Public Overrides Sub WriteLine(value As String)
            _buffer.Append(value)
            _buffer.Append(System.Environment.NewLine)
            Flush()
        End Sub

        Public Overrides Sub Flush()
            If _buffer.Length > 0 Then
                _logAction?.Invoke(_buffer.ToString())
                _buffer.Clear()
            End If
        End Sub

        Public Overrides ReadOnly Property Encoding As System.Text.Encoding
            Get
                Return System.Text.Encoding.UTF8
            End Get
        End Property
    End Class

    Public Function ExecuteVBCodeWithVariables(vbCodeString As String, variables As System.Collections.Generic.Dictionary(Of String, Object), Optional entryPointName As String = Nothing, Optional customLogger As System.Action(Of String) = Nothing) As String
        Dim executionId As String = System.Guid.NewGuid().ToString("N")
        Try
            Dim modifiedCode As String = InjectVariables(vbCodeString, variables, executionId)
            Dim result As String = ExecuteVBCodeInternal(modifiedCode, Nothing, entryPointName, customLogger)
            Return result
        Catch ex As System.Exception
            Return $"Fatal Error: {ex.Message}" & Microsoft.VisualBasic.ControlChars.CrLf & ex.StackTrace
        Finally
            ' Clean up stored objects for this execution
            CleanupExecution(executionId)
        End Try
    End Function

    Public Function ExecuteVBCode(vbCodeString As String, customLogger As System.Action(Of String), ParamArray parameters As Object()) As String
        Return ExecuteVBCodeInternal(vbCodeString, parameters, Nothing, customLogger)
    End Function

    Public Function ExecuteVBCode(vbCodeString As String, ParamArray parameters As Object()) As String
        Return ExecuteVBCodeInternal(vbCodeString, parameters, Nothing, Nothing)
    End Function

    Public Function ExecuteVBCode(vbCodeString As String, entryPointName As String, customLogger As System.Action(Of String), ParamArray parameters As Object()) As String
        Return ExecuteVBCodeInternal(vbCodeString, parameters, entryPointName, customLogger)
    End Function

    Public Function ExecuteVBCode(vbCodeString As String, entryPointName As String, ParamArray parameters As Object()) As String
        Return ExecuteVBCodeInternal(vbCodeString, parameters, entryPointName, Nothing)
    End Function

    Private Function InjectVariables(vbCodeString As String, variables As System.Collections.Generic.Dictionary(Of String, Object), executionId As String) As String
        If variables Is Nothing OrElse variables.Count = 0 Then
            Return vbCodeString
        End If

        Dim variableDeclarations As New System.Text.StringBuilder()

        For Each kvp As System.Collections.Generic.KeyValuePair(Of String, Object) In variables
            Dim varName As String = kvp.Key
            Dim varValue As Object = kvp.Value

            Dim valueStr As String = Nothing
            Dim useDirectAssignment As Boolean = True

            ' Check if this is a primitive type that can be represented as a literal
            If varValue Is Nothing Then
                valueStr = "Nothing"
            ElseIf TypeOf varValue Is String Then
                valueStr = $"""{varValue.ToString().Replace("""", """""")}"""
            ElseIf TypeOf varValue Is Char Then
                Dim charValue As Char = CChar(varValue)
                ' Use ChrW for proper Unicode character representation
                valueStr = $"ChrW({Microsoft.VisualBasic.AscW(charValue)})"
            ElseIf TypeOf varValue Is Boolean Then
                valueStr = If(CBool(varValue), "True", "False")
            ElseIf TypeOf varValue Is Integer Then
                valueStr = CInt(varValue).ToString(System.Globalization.CultureInfo.InvariantCulture)
            ElseIf TypeOf varValue Is Long Then
                valueStr = CLng(varValue).ToString(System.Globalization.CultureInfo.InvariantCulture) & "L"
            ElseIf TypeOf varValue Is Short Then
                valueStr = CShort(varValue).ToString(System.Globalization.CultureInfo.InvariantCulture) & "S"
            ElseIf TypeOf varValue Is Byte Then
                valueStr = CByte(varValue).ToString(System.Globalization.CultureInfo.InvariantCulture)
            ElseIf TypeOf varValue Is SByte Then
                valueStr = $"CSByte({CSByte(varValue).ToString(System.Globalization.CultureInfo.InvariantCulture)})"
            ElseIf TypeOf varValue Is UShort Then
                valueStr = $"CUShort({CUShort(varValue).ToString(System.Globalization.CultureInfo.InvariantCulture)})"
            ElseIf TypeOf varValue Is UInteger Then
                valueStr = CUInt(varValue).ToString(System.Globalization.CultureInfo.InvariantCulture) & "UI"
            ElseIf TypeOf varValue Is ULong Then
                valueStr = CULng(varValue).ToString(System.Globalization.CultureInfo.InvariantCulture) & "UL"
            ElseIf TypeOf varValue Is Double Then
                Dim dblValue As Double = CDbl(varValue)
                ' Handle special double values
                If Double.IsNaN(dblValue) Then
                    valueStr = "Double.NaN"
                ElseIf Double.IsPositiveInfinity(dblValue) Then
                    valueStr = "Double.PositiveInfinity"
                ElseIf Double.IsNegativeInfinity(dblValue) Then
                    valueStr = "Double.NegativeInfinity"
                Else
                    valueStr = dblValue.ToString("R", System.Globalization.CultureInfo.InvariantCulture)
                End If
            ElseIf TypeOf varValue Is Single Then
                Dim sngValue As Single = CSng(varValue)
                ' Handle special single values
                If Single.IsNaN(sngValue) Then
                    valueStr = "Single.NaN"
                ElseIf Single.IsPositiveInfinity(sngValue) Then
                    valueStr = "Single.PositiveInfinity"
                ElseIf Single.IsNegativeInfinity(sngValue) Then
                    valueStr = "Single.NegativeInfinity"
                Else
                    valueStr = sngValue.ToString("R", System.Globalization.CultureInfo.InvariantCulture) & "F"
                End If
            ElseIf TypeOf varValue Is Decimal Then
                valueStr = CDec(varValue).ToString(System.Globalization.CultureInfo.InvariantCulture) & "D"
            ElseIf TypeOf varValue Is System.DateTime Then
                ' DateTime: Store as ticks and DateTimeKind for precise representation
                Dim dt As System.DateTime = CType(varValue, System.DateTime)
                valueStr = $"New DateTime({dt.Ticks}L, DateTimeKind.{dt.Kind})"
            ElseIf TypeOf varValue Is System.DateTimeOffset Then
                ' DateTimeOffset: Store as ticks and offset
                Dim dto As System.DateTimeOffset = CType(varValue, System.DateTimeOffset)
                valueStr = $"New DateTimeOffset({dto.Ticks}L, New TimeSpan({dto.Offset.Ticks}L))"
            ElseIf TypeOf varValue Is System.TimeSpan Then
                ' TimeSpan: Store as ticks
                Dim ts As System.TimeSpan = CType(varValue, System.TimeSpan)
                valueStr = $"New TimeSpan({ts.Ticks}L)"
            ElseIf TypeOf varValue Is System.Guid Then
                ' Guid: Store as string and parse
                Dim g As System.Guid = CType(varValue, System.Guid)
                valueStr = $"New Guid(""{g.ToString()}"")"
            ElseIf TypeOf varValue Is System.DBNull Then
                ' DBNull
                valueStr = "System.DBNull.Value"
            ElseIf varValue.GetType().IsEnum Then
                ' Enum types: Convert to underlying type and cast back
                Dim enumType As System.Type = varValue.GetType()
                Dim underlyingValue As Object = System.Convert.ChangeType(varValue, System.Enum.GetUnderlyingType(enumType))
                Dim enumTypeName As String = enumType.FullName
                valueStr = $"CType({underlyingValue}, {enumTypeName})"
            Else
                ' Check if it's a Nullable type
                Dim actualType As System.Type = varValue.GetType()
                If actualType.IsGenericType AndAlso actualType.GetGenericTypeDefinition() Is GetType(System.Nullable(Of )) Then
                    ' This shouldn't normally happen since Nullable(Of T) with a value would be unwrapped
                    ' But handle it just in case by storing in object store
                    Dim storageKey As String = $"{executionId}_{varName}"
                    _objectStore.TryAdd(storageKey, varValue)
                    valueStr = $"VBCodeExecutor.GetStoredObject(""{storageKey}"")"
                    useDirectAssignment = False
                Else
                    ' Complex type: Store in object store and retrieve via reference
                    Dim storageKey As String = $"{executionId}_{varName}"
                    _objectStore.TryAdd(storageKey, varValue)
                    valueStr = $"VBCodeExecutor.GetStoredObject(""{storageKey}"")"
                    useDirectAssignment = False
                End If
            End If

            ' Generate the variable declaration
            If useDirectAssignment Then
                variableDeclarations.AppendLine($"    Private {varName} As Object = {valueStr}")
            Else
                ' For complex types, we need to cast appropriately
                Dim actualType As System.Type = varValue.GetType()
                Dim typeName As String = GetVBTypeName(actualType)
                variableDeclarations.AppendLine($"    Private {varName} As Object = {valueStr}")
            End If
        Next

        ' Try to find the module declaration with various patterns
        Dim modulePattern As String = "Module\s+\w+"
        Dim match As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(vbCodeString, modulePattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

        If match.Success Then
            ' Find the end of the line where the module is declared
            ' Handle different line ending types: CRLF (\r\n), LF (\n), or CR (\r)
            Dim moduleEndIndex As Integer = match.Index + match.Length
            Dim searchStartIndex As Integer = moduleEndIndex

            ' Find the next line ending
            Dim crlfIndex As Integer = vbCodeString.IndexOf(System.Environment.NewLine, searchStartIndex)
            Dim lfIndex As Integer = vbCodeString.IndexOf(Microsoft.VisualBasic.ControlChars.Lf, searchStartIndex)
            Dim crIndex As Integer = vbCodeString.IndexOf(Microsoft.VisualBasic.ControlChars.Cr, searchStartIndex)

            ' Choose the earliest line ending found
            Dim lineEndIndex As Integer = -1

            If crlfIndex >= 0 Then lineEndIndex = crlfIndex + System.Environment.NewLine.Length
            If lfIndex >= 0 AndAlso (lineEndIndex = -1 OrElse lfIndex < lineEndIndex) Then
                lineEndIndex = lfIndex + 1
            ElseIf crIndex >= 0 AndAlso (lineEndIndex = -1 OrElse crIndex < lineEndIndex) Then
                lineEndIndex = crIndex + 1
            End If

            If lineEndIndex > 0 Then
                ' Inject the variable declarations after the module line
                Dim result As String = vbCodeString.Substring(0, lineEndIndex) &
                                      variableDeclarations.ToString() &
                                      vbCodeString.Substring(lineEndIndex)
                Return result
            End If
        End If

        ' Fallback: If we couldn't find a module declaration, try to inject at the beginning
        ' This shouldn't normally happen, but provides a safety net
        Return variableDeclarations.ToString() & System.Environment.NewLine & vbCodeString
    End Function

    ''' <summary>
    ''' Helper function to get VB-friendly type names for better type hints
    ''' </summary>
    Private Function GetVBTypeName(type As System.Type) As String
        If type Is Nothing Then Return "Object"

        ' Handle signed integer types
        If type Is GetType(System.Int32) Then Return "Integer"
        If type Is GetType(System.Int64) Then Return "Long"
        If type Is GetType(System.Int16) Then Return "Short"
        If type Is GetType(System.SByte) Then Return "SByte"

        ' Handle unsigned integer types
        If type Is GetType(System.Byte) Then Return "Byte"
        If type Is GetType(System.UInt16) Then Return "UShort"
        If type Is GetType(System.UInt32) Then Return "UInteger"
        If type Is GetType(System.UInt64) Then Return "ULong"

        ' Handle floating point types
        If type Is GetType(System.Double) Then Return "Double"
        If type Is GetType(System.Single) Then Return "Single"
        If type Is GetType(System.Decimal) Then Return "Decimal"

        ' Handle other primitive types
        If type Is GetType(System.Boolean) Then Return "Boolean"
        If type Is GetType(System.Char) Then Return "Char"
        If type Is GetType(System.String) Then Return "String"

        ' Handle date/time types
        If type Is GetType(System.DateTime) Then Return "DateTime"
        If type Is GetType(System.DateTimeOffset) Then Return "DateTimeOffset"
        If type Is GetType(System.TimeSpan) Then Return "TimeSpan"

        ' Handle other common types
        If type Is GetType(System.Guid) Then Return "Guid"
        If type Is GetType(System.DBNull) Then Return "DBNull"

        ' Handle enums
        If type.IsEnum Then
            Return type.FullName
        End If

        ' For arrays
        If type.IsArray Then
            Dim elementType As System.Type = type.GetElementType()
            Dim rank As Integer = type.GetArrayRank()
            Dim commas As String = New String(","c, rank - 1)
            Return GetVBTypeName(elementType) & "(" & commas & ")"
        End If

        ' For generic types
        If type.IsGenericType Then
            Dim genericTypeDef As System.Type = type.GetGenericTypeDefinition()

            ' Handle Nullable types
            If genericTypeDef Is GetType(System.Nullable(Of )) Then
                Return GetVBTypeName(type.GetGenericArguments()(0)) & "?"
            End If

            If genericTypeDef Is GetType(System.Collections.Generic.List(Of )) Then
                Return $"List(Of {GetVBTypeName(type.GetGenericArguments()(0))})"
            ElseIf genericTypeDef Is GetType(System.Collections.Generic.Dictionary(Of ,)) Then
                Dim args As System.Type() = type.GetGenericArguments()
                Return $"Dictionary(Of {GetVBTypeName(args(0))}, {GetVBTypeName(args(1))})"
            ElseIf genericTypeDef Is GetType(System.Collections.Generic.IEnumerable(Of )) Then
                Return $"IEnumerable(Of {GetVBTypeName(type.GetGenericArguments()(0))})"
            ElseIf genericTypeDef Is GetType(System.Collections.Generic.IList(Of )) Then
                Return $"IList(Of {GetVBTypeName(type.GetGenericArguments()(0))})"
            ElseIf genericTypeDef Is GetType(System.Collections.Generic.IDictionary(Of ,)) Then
                Dim args As System.Type() = type.GetGenericArguments()
                Return $"IDictionary(Of {GetVBTypeName(args(0))}, {GetVBTypeName(args(1))})"
            End If
        End If

        ' Default: use the type's full name for better clarity
        Return If(String.IsNullOrEmpty(type.FullName), type.Name, type.FullName)
    End Function

    ''' <summary>
    ''' Clean up stored objects for a specific execution
    ''' </summary>
    Private Sub CleanupExecution(executionId As String)
        Dim keysToRemove As New System.Collections.Generic.List(Of String)()

        ' Find all keys for this execution
        For Each key As String In _objectStore.Keys
            If key.StartsWith(executionId & "_") Then
                keysToRemove.Add(key)
            End If
        Next

        ' Remove them
        For Each key As String In keysToRemove
            Dim removedValue As Object = Nothing
            _objectStore.TryRemove(key, removedValue)
        Next
    End Sub

    Private Function ExecuteVBCodeInternal(vbCodeString As String, parameters As Object(), entryPointName As String, customLogger As System.Action(Of String)) As String
        Try
            Dim syntaxTree As Microsoft.CodeAnalysis.SyntaxTree = Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxTree.ParseText(vbCodeString)

            Dim assemblyName As String = "DynamicAssembly_" & System.Guid.NewGuid().ToString("N")

            Dim references As Microsoft.CodeAnalysis.MetadataReference() = {
                Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(GetType(Object).Assembly.Location),
                Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(GetType(System.Linq.Enumerable).Assembly.Location),
                Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(GetType(Microsoft.VisualBasic.CompilerServices.StandardModuleAttribute).Assembly.Location),
                Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(System.Reflection.Assembly.Load("System").Location),
                Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(System.Reflection.Assembly.GetExecutingAssembly().Location)
            }

            Dim globalImports = {
                Microsoft.CodeAnalysis.VisualBasic.GlobalImport.Parse("System"),
                Microsoft.CodeAnalysis.VisualBasic.GlobalImport.Parse("System.Collections.Generic"),
                Microsoft.CodeAnalysis.VisualBasic.GlobalImport.Parse("System.Linq")
            }

            Dim compilationOptions = New Microsoft.CodeAnalysis.VisualBasic.VisualBasicCompilationOptions(Microsoft.CodeAnalysis.OutputKind.DynamicallyLinkedLibrary).
                WithGlobalImports(globalImports)

            Dim compilation As Microsoft.CodeAnalysis.VisualBasic.VisualBasicCompilation = Microsoft.CodeAnalysis.VisualBasic.VisualBasicCompilation.Create(
                assemblyName,
                syntaxTrees:={syntaxTree},
                references:=references,
                options:=compilationOptions
            )

            Using ms As New System.IO.MemoryStream()
                Dim result As Microsoft.CodeAnalysis.Emit.EmitResult = compilation.Emit(ms)

                If Not result.Success Then
                    Dim failures As System.Collections.Generic.IEnumerable(Of Microsoft.CodeAnalysis.Diagnostic) = System.Linq.Enumerable.Where(Of Microsoft.CodeAnalysis.Diagnostic)(result.Diagnostics, Function(diagnostic As Microsoft.CodeAnalysis.Diagnostic) diagnostic.IsWarningAsError OrElse diagnostic.Severity = Microsoft.CodeAnalysis.DiagnosticSeverity.Error)

                    Dim errorBuilder As New System.Text.StringBuilder()
                    errorBuilder.AppendLine("Compilation Errors:")
                    For Each diagnostic As Microsoft.CodeAnalysis.Diagnostic In failures
                        errorBuilder.AppendLine($"  {diagnostic.Id}: {diagnostic.GetMessage()}")
                    Next

                    Return errorBuilder.ToString()
                Else
                    ms.Seek(0, System.IO.SeekOrigin.Begin)
                    Dim assembly As System.Reflection.Assembly = System.Reflection.Assembly.Load(ms.ToArray())

                    Dim type As System.Type = System.Linq.Enumerable.FirstOrDefault(Of System.Type)(assembly.GetTypes())
                    If type Is Nothing Then
                        Return "Error: No types found in compiled assembly. Make sure your code includes a Module or Class."
                    End If

                    Dim methods As System.Reflection.MethodInfo() = type.GetMethods(System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Static)
                    Dim method As System.Reflection.MethodInfo = Nothing

                    ' If entry point name is specified, filter by name first
                    If Not String.IsNullOrEmpty(entryPointName) Then
                        If parameters IsNot Nothing AndAlso parameters.Length > 0 Then
                            method = System.Linq.Enumerable.FirstOrDefault(Of System.Reflection.MethodInfo)(methods, Function(m As System.Reflection.MethodInfo) m.Name = entryPointName AndAlso m.GetParameters().Length = parameters.Length)
                            If method Is Nothing Then
                                Return $"Error: No public shared method named '{entryPointName}' found with {parameters.Length} parameter(s). Make sure your code includes a method named '{entryPointName}' that accepts {parameters.Length} parameter(s)."
                            End If
                        Else
                            method = System.Linq.Enumerable.FirstOrDefault(Of System.Reflection.MethodInfo)(methods, Function(m As System.Reflection.MethodInfo) m.Name = entryPointName AndAlso m.GetParameters().Length = 0)
                            If method Is Nothing Then
                                ' Try to find the method with any parameter count
                                method = System.Linq.Enumerable.FirstOrDefault(Of System.Reflection.MethodInfo)(methods, Function(m As System.Reflection.MethodInfo) m.Name = entryPointName)
                                If method Is Nothing Then
                                    Return $"Error: No public shared method named '{entryPointName}' found. Make sure your code includes a method named '{entryPointName}'."
                                Else
                                    Return $"Error: Method '{entryPointName}' was found but it requires {method.GetParameters().Length} parameter(s). Please provide the required parameters."
                                End If
                            End If
                        End If
                    Else
                        ' No entry point specified, use original logic
                        If parameters IsNot Nothing AndAlso parameters.Length > 0 Then
                            method = System.Linq.Enumerable.FirstOrDefault(Of System.Reflection.MethodInfo)(methods, Function(m As System.Reflection.MethodInfo) m.GetParameters().Length = parameters.Length)
                            If method Is Nothing Then
                                Return $"Error: No public shared method found with {parameters.Length} parameter(s). Make sure your code includes a method that accepts {parameters.Length} parameter(s)."
                            End If
                        Else
                            method = System.Linq.Enumerable.FirstOrDefault(Of System.Reflection.MethodInfo)(methods, Function(m As System.Reflection.MethodInfo) m.GetParameters().Length = 0)
                            If method Is Nothing Then
                                method = System.Linq.Enumerable.FirstOrDefault(Of System.Reflection.MethodInfo)(methods)
                                If method Is Nothing Then
                                    Return "Error: No public shared methods found. Make sure your code includes a public shared function or sub."
                                End If
                            End If
                        End If
                    End If

                    Dim originalOut As System.IO.TextWriter = System.Console.Out
                    Dim originalError As System.IO.TextWriter = System.Console.Error

                    Try
                        Dim outputBuilder As New System.Text.StringBuilder()
                        Dim outputWriter As System.IO.TextWriter

                        ' Use custom logger if provided, otherwise use default StringWriter
                        If customLogger IsNot Nothing Then
                            outputWriter = New ActionTextWriter(customLogger)
                        Else
                            outputWriter = New System.IO.StringWriter(outputBuilder)
                        End If

                        Using outputWriter
                            System.Console.SetOut(outputWriter)
                            System.Console.SetError(outputWriter)

                            Dim methodResult As Object = method.Invoke(Nothing, parameters)

                            System.Console.SetOut(originalOut)
                            System.Console.SetError(originalError)

                            ' Flush any remaining buffered output
                            outputWriter.Flush()

                            Dim output As String = outputBuilder.ToString()

                            If methodResult IsNot Nothing Then
                                If System.String.IsNullOrEmpty(output) Then
                                    Return $"Result: {methodResult}"
                                Else
                                    Return $"{output}Result: {methodResult}"
                                End If
                            Else
                                If System.String.IsNullOrEmpty(output) Then
                                    Return "Execution completed successfully (no return value)."
                                Else
                                    Return output
                                End If
                            End If
                        End Using
                    Catch ex As System.Exception
                        System.Console.SetOut(originalOut)
                        System.Console.SetError(originalError)
                        Dim errorMessage As String = If(ex.InnerException IsNot Nothing, ex.InnerException.Message, ex.Message)
                        Return $"Runtime Error: {errorMessage}" & Microsoft.VisualBasic.ControlChars.CrLf & ex.StackTrace
                    End Try
                End If
            End Using

        Catch ex As System.Exception
            Return $"Fatal Error: {ex.Message}" & Microsoft.VisualBasic.ControlChars.CrLf & ex.StackTrace
        End Try
    End Function

End Module
