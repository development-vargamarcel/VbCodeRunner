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
        Return ExecuteVBCodeWithVariables(vbCodeString, variables, Nothing, customLogger)
    End Function

    Public Function ExecuteVBCodeWithVariables(vbCodeString As String, variables As System.Collections.Generic.Dictionary(Of String, Object), analysisOptions As CodeAnalysisOptions, Optional customLogger As System.Action(Of String) = Nothing) As String
        Try
            Dim modifiedCode As String = InjectVariables(vbCodeString, variables)
            Return ExecuteVBCodeInternal(modifiedCode, Nothing, customLogger, analysisOptions)
        Catch ex As System.Exception
            Return $"Fatal Error: {ex.Message}" & Microsoft.VisualBasic.ControlChars.CrLf & ex.StackTrace
        End Try
    End Function

    Public Function ExecuteVBCode(vbCodeString As String, customLogger As System.Action(Of String), ParamArray parameters As Object()) As String
        Return ExecuteVBCodeInternal(vbCodeString, parameters, customLogger, Nothing)
    End Function

    Public Function ExecuteVBCode(vbCodeString As String, ParamArray parameters As Object()) As String
        Return ExecuteVBCodeInternal(vbCodeString, parameters, Nothing, Nothing)
    End Function

    ''' <summary>
    ''' Execute VB code with optional code analysis before execution
    ''' </summary>
    Public Function ExecuteVBCode(vbCodeString As String, analysisOptions As CodeAnalysisOptions, customLogger As System.Action(Of String), ParamArray parameters As Object()) As String
        Return ExecuteVBCodeInternal(vbCodeString, parameters, customLogger, analysisOptions)
    End Function

    ''' <summary>
    ''' Execute VB code with optional code analysis before execution
    ''' </summary>
    Public Function ExecuteVBCode(vbCodeString As String, analysisOptions As CodeAnalysisOptions, ParamArray parameters As Object()) As String
        Return ExecuteVBCodeInternal(vbCodeString, parameters, Nothing, analysisOptions)
    End Function

    Private Function InjectVariables(vbCodeString As String, variables As System.Collections.Generic.Dictionary(Of String, Object)) As String
        If variables Is Nothing OrElse variables.Count = 0 Then
            Return vbCodeString
        End If

        Dim variableDeclarations As New System.Text.StringBuilder()
        variableDeclarations.AppendLine()

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
                valueStr = CLng(varValue).ToString(System.Globalization.CultureInfo.InvariantCulture)
            ElseIf TypeOf varValue Is Short Then
                valueStr = CShort(varValue).ToString(System.Globalization.CultureInfo.InvariantCulture)
            ElseIf TypeOf varValue Is Byte Then
                valueStr = CByte(varValue).ToString(System.Globalization.CultureInfo.InvariantCulture)
            ElseIf TypeOf varValue Is Double Then
                valueStr = CDbl(varValue).ToString("R", System.Globalization.CultureInfo.InvariantCulture)
            ElseIf TypeOf varValue Is Single Then
                valueStr = CSng(varValue).ToString("R", System.Globalization.CultureInfo.InvariantCulture)
            ElseIf TypeOf varValue Is Decimal Then
                valueStr = CDec(varValue).ToString(System.Globalization.CultureInfo.InvariantCulture) & "D"
            Else
                valueStr = $"DirectCast(Nothing, Object)"
            End If

            variableDeclarations.AppendLine($"        Dim {varName} As Object = {valueStr}")
        Next

        Dim moduleStartIndex As Integer = vbCodeString.IndexOf("Public Module", System.StringComparison.OrdinalIgnoreCase)
        If moduleStartIndex = -1 Then
            moduleStartIndex = vbCodeString.IndexOf("Module", System.StringComparison.OrdinalIgnoreCase)
        End If

        If moduleStartIndex >= 0 Then
            Dim moduleLineEnd As Integer = vbCodeString.IndexOf(Microsoft.VisualBasic.ControlChars.Lf, moduleStartIndex)
            If moduleLineEnd = -1 Then moduleLineEnd = vbCodeString.IndexOf(Microsoft.VisualBasic.ControlChars.Cr, moduleStartIndex)

            If moduleLineEnd >= 0 Then
                Return vbCodeString.Substring(0, moduleLineEnd + 1) &
                       variableDeclarations.ToString() &
                       vbCodeString.Substring(moduleLineEnd + 1)
            End If
        End If

        Return vbCodeString
    End Function

    Private Function ExecuteVBCodeInternal(vbCodeString As String, parameters As Object(), customLogger As System.Action(Of String), analysisOptions As CodeAnalysisOptions) As String
        Try
            ' Perform code analysis if enabled
            If analysisOptions IsNot Nothing AndAlso analysisOptions.Enabled Then
                Dim analysisResult As CodeAnalyzer.AnalysisResult = CodeAnalyzer.AnalyzeCode(vbCodeString, analysisOptions)

                ' If analysis found issues and they prevent execution, return the report
                If Not analysisResult.Success Then
                    Return analysisResult.GetFormattedReport() & Microsoft.VisualBasic.ControlChars.CrLf & "Execution aborted due to analysis errors."
                End If

                ' If analysis succeeded but found warnings/info, prepend them to the output
                If analysisResult.HasIssues() Then
                    Dim report As String = analysisResult.GetFormattedReport()
                    ' Store the report to prepend to final output
                    Dim analysisReport As String = report & Microsoft.VisualBasic.ControlChars.CrLf & "Proceeding with execution..." & Microsoft.VisualBasic.ControlChars.CrLf & Microsoft.VisualBasic.ControlChars.CrLf

                    ' Execute the code and prepend the analysis report
                    Dim executionResult As String = ExecuteVBCodeInternalCore(vbCodeString, parameters, customLogger)
                    Return analysisReport & executionResult
                End If
            End If

            ' No analysis or no issues found, proceed with normal execution
            Return ExecuteVBCodeInternalCore(vbCodeString, parameters, customLogger)
        Catch ex As System.Exception
            Return $"Fatal Error: {ex.Message}" & Microsoft.VisualBasic.ControlChars.CrLf & ex.StackTrace
        End Try
    End Function

    Private Function ExecuteVBCodeInternalCore(vbCodeString As String, parameters As Object(), customLogger As System.Action(Of String)) As String
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
