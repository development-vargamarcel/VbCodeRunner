Public Module VBCodeExecutor

    Public Function ExecuteVBCodeWithVariables(vbCodeString As String, variables As System.Collections.Generic.Dictionary(Of String, Object)) As String
        Try
            Dim modifiedCode As String = InjectVariables(vbCodeString, variables)
            Return ExecuteVBCodeInternal(modifiedCode, Nothing)
        Catch ex As System.Exception
            Return $"Fatal Error: {ex.Message}" & vbCrLf & ex.StackTrace
        End Try
    End Function

    Public Function ExecuteVBCode(vbCodeString As String, ParamArray parameters As Object()) As String
        Return ExecuteVBCodeInternal(vbCodeString, parameters)
    End Function

    Private Function InjectVariables(vbCodeString As String, variables As System.Collections.Generic.Dictionary(Of String, Object)) As String
        If variables Is Nothing OrElse variables.Count = 0 Then
            Return vbCodeString
        End If

        Dim variableDeclarations As New System.Text.StringBuilder()
        variableDeclarations.AppendLine()

        For Each kvp In variables
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
            Dim moduleLineEnd As Integer = vbCodeString.IndexOf(vbLf, moduleStartIndex)
            If moduleLineEnd = -1 Then moduleLineEnd = vbCodeString.IndexOf(vbCr, moduleStartIndex)

            If moduleLineEnd >= 0 Then
                Return vbCodeString.Substring(0, moduleLineEnd + 1) &
                       variableDeclarations.ToString() &
                       vbCodeString.Substring(moduleLineEnd + 1)
            End If
        End If

        Return vbCodeString
    End Function

    Private Function ExecuteVBCodeInternal(vbCodeString As String, parameters As Object()) As String
        Try
            Dim syntaxTree As Microsoft.CodeAnalysis.SyntaxTree = Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxTree.ParseText(vbCodeString)

            Dim assemblyName As String = "DynamicAssembly_" & System.Guid.NewGuid().ToString("N")

            Dim references As Microsoft.CodeAnalysis.MetadataReference() = {
                Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(GetType(Object).Assembly.Location),
                Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(GetType(System.Console).Assembly.Location),
                Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(System.Reflection.Assembly.Load("System.Runtime").Location),
                Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(System.Reflection.Assembly.Load("System.Console").Location),
                Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(System.Reflection.Assembly.Load("netstandard").Location),
                Microsoft.CodeAnalysis.MetadataReference.CreateFromFile(GetType(Microsoft.VisualBasic.CompilerServices.StandardModuleAttribute).Assembly.Location)
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
                    Dim failures = result.Diagnostics.Where(Function(diagnostic) diagnostic.IsWarningAsError OrElse diagnostic.Severity = Microsoft.CodeAnalysis.DiagnosticSeverity.Error)

                    Dim errorBuilder As New System.Text.StringBuilder()
                    errorBuilder.AppendLine("Compilation Errors:")
                    For Each diagnostic In failures
                        errorBuilder.AppendLine($"  {diagnostic.Id}: {diagnostic.GetMessage()}")
                    Next

                    Return errorBuilder.ToString()
                Else
                    ms.Seek(0, System.IO.SeekOrigin.Begin)
                    Dim assembly As System.Reflection.Assembly = System.Reflection.Assembly.Load(ms.ToArray())

                    Dim type = assembly.GetTypes().FirstOrDefault()
                    If type Is Nothing Then
                        Return "Error: No types found in compiled assembly. Make sure your code includes a Module or Class."
                    End If

                    Dim methods = type.GetMethods(System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Static)
                    Dim method As System.Reflection.MethodInfo = Nothing

                    If parameters IsNot Nothing AndAlso parameters.Length > 0 Then
                        method = methods.FirstOrDefault(Function(m) m.GetParameters().Length = parameters.Length)
                        If method Is Nothing Then
                            Return $"Error: No public shared method found with {parameters.Length} parameter(s). Make sure your code includes a method that accepts {parameters.Length} parameter(s)."
                        End If
                    Else
                        method = methods.FirstOrDefault(Function(m) m.GetParameters().Length = 0)
                        If method Is Nothing Then
                            method = methods.FirstOrDefault()
                            If method Is Nothing Then
                                Return "Error: No public shared methods found. Make sure your code includes a public shared function or sub."
                            End If
                        End If
                    End If

                    Dim originalOut As System.IO.TextWriter = System.Console.Out
                    Dim originalError As System.IO.TextWriter = System.Console.Error

                    Try
                        Dim outputBuilder As New System.Text.StringBuilder()
                        Using sw As New System.IO.StringWriter(outputBuilder)
                            System.Console.SetOut(sw)
                            System.Console.SetError(sw)

                            Dim methodResult As Object = method.Invoke(Nothing, parameters)

                            System.Console.SetOut(originalOut)
                            System.Console.SetError(originalError)

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
                        Return $"Runtime Error: {errorMessage}" & vbCrLf & ex.StackTrace
                    End Try
                End If
            End Using

        Catch ex As System.Exception
            Return $"Fatal Error: {ex.Message}" & vbCrLf & ex.StackTrace
        End Try
    End Function

End Module
