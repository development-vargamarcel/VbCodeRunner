Public Module VBCodeExecutor

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

    Public Function ExecuteVBCodeWithVariables(vbCodeString As String, variables As System.Collections.Generic.Dictionary(Of String, Object), Optional customLogger As System.Action(Of String) = Nothing) As String
        Try
            Dim modifiedCode As String = InjectVariables(vbCodeString, variables)
            Return ExecuteVBCodeInternal(modifiedCode, Nothing, customLogger)
        Catch ex As System.Exception
            Return $"Fatal Error: {ex.Message}" & Microsoft.VisualBasic.ControlChars.CrLf & ex.StackTrace
        End Try
    End Function

    Public Function ExecuteVBCode(vbCodeString As String, customLogger As System.Action(Of String), ParamArray parameters As Object()) As String
        Return ExecuteVBCodeInternal(vbCodeString, parameters, customLogger)
    End Function

    Public Function ExecuteVBCode(vbCodeString As String, ParamArray parameters As Object()) As String
        Return ExecuteVBCodeInternal(vbCodeString, parameters, Nothing)
    End Function

    Private Function InjectVariables(vbCodeString As String, variables As System.Collections.Generic.Dictionary(Of String, Object)) As String
        If variables Is Nothing OrElse variables.Count = 0 Then
            Return vbCodeString
        End If

        Dim variableDeclarations As New System.Text.StringBuilder()

        For Each kvp As System.Collections.Generic.KeyValuePair(Of String, Object) In variables
            Dim varName As String = kvp.Key
            Dim varValue As Object = kvp.Value

            Dim valueStr As String
            If varValue Is Nothing Then
                valueStr = "Nothing"
            ElseIf TypeOf varValue Is String Then
                valueStr = $"""{varValue.ToString().Replace("""", """""")}"""
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
            ElseIf TypeOf varValue Is Double Then
                valueStr = CDbl(varValue).ToString("R", System.Globalization.CultureInfo.InvariantCulture)
            ElseIf TypeOf varValue Is Single Then
                valueStr = CSng(varValue).ToString("R", System.Globalization.CultureInfo.InvariantCulture) & "F"
            ElseIf TypeOf varValue Is Decimal Then
                valueStr = CDec(varValue).ToString(System.Globalization.CultureInfo.InvariantCulture) & "D"
            Else
                valueStr = $"DirectCast(Nothing, Object)"
            End If

            ' Use Private instead of Dim for module-level variables
            variableDeclarations.AppendLine($"    Private {varName} As Object = {valueStr}")
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

    Private Function ExecuteVBCodeInternal(vbCodeString As String, parameters As Object(), customLogger As System.Action(Of String)) As String
        Try
            Dim syntaxTree As Microsoft.CodeAnalysis.SyntaxTree = Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxTree.ParseText(vbCodeString)

            Dim assemblyName As String = "DynamicAssembly_" & System.Guid.NewGuid().ToString("N")

            Dim references As Microsoft.CodeAnalysis.MetadataReference() = {
                Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(GetType(Object).Assembly.Location),
                Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(GetType(System.Linq.Enumerable).Assembly.Location),
                Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(GetType(Microsoft.VisualBasic.CompilerServices.StandardModuleAttribute).Assembly.Location),
                Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(System.Reflection.Assembly.Load("System").Location)
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
