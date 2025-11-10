''' <summary>
''' Provides code analysis capabilities for VB.NET code before execution
''' </summary>
Public Module CodeAnalyzer

    ''' <summary>
    ''' Result of code analysis
    ''' </summary>
    Public Class AnalysisResult
        Public Property Success As Boolean
        Public Property Errors As New System.Collections.Generic.List(Of String)
        Public Property Warnings As New System.Collections.Generic.List(Of String)
        Public Property Info As New System.Collections.Generic.List(Of String)
        Public Property ComplexityMetrics As String = ""

        Public Function HasIssues() As Boolean
            Return Errors.Count > 0 OrElse Warnings.Count > 0
        End Function

        Public Function GetFormattedReport() As String
            Dim report As New System.Text.StringBuilder()
            report.AppendLine("=== Code Analysis Report ===")
            report.AppendLine()

            If Errors.Count > 0 Then
                report.AppendLine($"ERRORS ({Errors.Count}):")
                For Each errorMsg As String In Errors
                    report.AppendLine($"  ❌ {errorMsg}")
                Next
                report.AppendLine()
            End If

            If Warnings.Count > 0 Then
                report.AppendLine($"WARNINGS ({Warnings.Count}):")
                For Each warnMsg As String In Warnings
                    report.AppendLine($"  ⚠️  {warnMsg}")
                Next
                report.AppendLine()
            End If

            If Info.Count > 0 Then
                report.AppendLine($"INFO ({Info.Count}):")
                For Each infoMsg As String In Info
                    report.AppendLine($"  ℹ️  {infoMsg}")
                Next
                report.AppendLine()
            End If

            If Not System.String.IsNullOrEmpty(ComplexityMetrics) Then
                report.AppendLine("COMPLEXITY METRICS:")
                report.AppendLine(ComplexityMetrics)
                report.AppendLine()
            End If

            If Not HasIssues() Then
                report.AppendLine("✅ No issues found!")
            End If

            report.AppendLine("===========================")
            Return report.ToString()
        End Function
    End Class

    ''' <summary>
    ''' Analyze VB.NET code based on the provided options
    ''' </summary>
    Public Function AnalyzeCode(vbCodeString As String, options As CodeAnalysisOptions) As AnalysisResult
        Dim result As New AnalysisResult()
        result.Success = True

        If Not options.Enabled Then
            result.Info.Add("Code analysis is disabled")
            Return result
        End If

        Try
            ' Parse the code
            Dim syntaxTree As Microsoft.CodeAnalysis.SyntaxTree = Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxTree.ParseText(vbCodeString)

            ' Check syntax
            If options.CheckSyntax Then
                CheckSyntaxErrors(syntaxTree, result)
            End If

            ' Check semantics
            If options.CheckSemantics Then
                CheckSemanticErrors(syntaxTree, result)
            End If

            ' Check security
            If options.CheckSecurity Then
                CheckSecurityIssues(vbCodeString, result)
            End If

            ' Report complexity
            If options.ReportComplexity Then
                ReportComplexityMetrics(syntaxTree, result)
            End If

            ' Determine overall success
            If result.Errors.Count > 0 Then
                result.Success = False
            ElseIf options.WarningsAsErrors AndAlso result.Warnings.Count > 0 Then
                result.Success = False
                result.Info.Add("Treating warnings as errors (WarningsAsErrors = True)")
            End If

        Catch ex As System.Exception
            result.Success = False
            result.Errors.Add($"Analysis failed: {ex.Message}")
        End Try

        Return result
    End Function

    Private Sub CheckSyntaxErrors(syntaxTree As Microsoft.CodeAnalysis.SyntaxTree, result As AnalysisResult)
        Dim diagnostics As System.Collections.Generic.IEnumerable(Of Microsoft.CodeAnalysis.Diagnostic) = syntaxTree.GetDiagnostics()

        For Each diagnostic As Microsoft.CodeAnalysis.Diagnostic In diagnostics
            Dim location As String = ""
            If diagnostic.Location.IsInSource Then
                Dim lineSpan = diagnostic.Location.GetLineSpan()
                location = $"Line {lineSpan.StartLinePosition.Line + 1}: "
            End If

            Select Case diagnostic.Severity
                Case Microsoft.CodeAnalysis.DiagnosticSeverity.Error
                    result.Errors.Add($"{location}{diagnostic.GetMessage()}")
                Case Microsoft.CodeAnalysis.DiagnosticSeverity.Warning
                    result.Warnings.Add($"{location}{diagnostic.GetMessage()}")
                Case Else
                    result.Info.Add($"{location}{diagnostic.GetMessage()}")
            End Select
        Next
    End Sub

    Private Sub CheckSemanticErrors(syntaxTree As Microsoft.CodeAnalysis.SyntaxTree, result As AnalysisResult)
        Try
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
                "AnalysisAssembly",
                syntaxTrees:={syntaxTree},
                references:=references,
                options:=compilationOptions
            )

            Dim semanticDiagnostics As System.Collections.Generic.IEnumerable(Of Microsoft.CodeAnalysis.Diagnostic) = compilation.GetDiagnostics()

            For Each diagnostic As Microsoft.CodeAnalysis.Diagnostic In semanticDiagnostics
                If diagnostic.Severity = Microsoft.CodeAnalysis.DiagnosticSeverity.Hidden Then
                    Continue For
                End If

                Dim location As String = ""
                If diagnostic.Location.IsInSource Then
                    Dim lineSpan = diagnostic.Location.GetLineSpan()
                    location = $"Line {lineSpan.StartLinePosition.Line + 1}: "
                End If

                Select Case diagnostic.Severity
                    Case Microsoft.CodeAnalysis.DiagnosticSeverity.Error
                        result.Errors.Add($"{location}{diagnostic.GetMessage()}")
                    Case Microsoft.CodeAnalysis.DiagnosticSeverity.Warning
                        result.Warnings.Add($"{location}{diagnostic.GetMessage()}")
                End Select
            Next

        Catch ex As System.Exception
            result.Warnings.Add($"Semantic analysis failed: {ex.Message}")
        End Try
    End Sub

    Private Sub CheckSecurityIssues(vbCodeString As String, result As AnalysisResult)
        ' Check for potentially dangerous namespaces/classes
        Dim securityPatterns As New System.Collections.Generic.Dictionary(Of String, String) From {
            {"System.IO.File", "File system operations detected"},
            {"System.IO.Directory", "Directory operations detected"},
            {"System.Diagnostics.Process", "Process execution detected"},
            {"System.Reflection", "Reflection usage detected"},
            {"System.Runtime.InteropServices", "Native interop detected"},
            {"Unsafe", "Unsafe code detected"},
            {"System.Net.WebClient", "Network operations detected"},
            {"System.Net.Http", "HTTP operations detected"},
            {"System.Data.SqlClient", "SQL database access detected"},
            {"System.Security.Cryptography", "Cryptography usage detected"}
        }

        For Each pattern As System.Collections.Generic.KeyValuePair(Of String, String) In securityPatterns
            If vbCodeString.Contains(pattern.Key) Then
                result.Warnings.Add($"Security: {pattern.Value} ({pattern.Key})")
            End If
        Next

        ' Check for potentially dangerous keywords
        If vbCodeString.Contains("DllImport") Then
            result.Warnings.Add("Security: DLL imports detected - potential native code execution")
        End If

        If vbCodeString.Contains("Invoke(") Then
            result.Info.Add("Reflection invoke detected - ensure input validation")
        End If
    End Sub

    Private Sub ReportComplexityMetrics(syntaxTree As Microsoft.CodeAnalysis.SyntaxTree, result As AnalysisResult)
        Dim root = syntaxTree.GetRoot()
        Dim metrics As New System.Text.StringBuilder()

        Dim codeLines As Integer = root.GetText().Lines.Count
        Dim nonEmptyLines As Integer = 0

        For Each line In root.GetText().Lines
            Dim lineText As String = line.ToString().Trim()
            If Not System.String.IsNullOrWhiteSpace(lineText) AndAlso Not lineText.StartsWith("'") Then
                nonEmptyLines += 1
            End If
        Next

        metrics.AppendLine($"  Total lines: {codeLines}")
        metrics.AppendLine($"  Non-empty lines: {nonEmptyLines}")

        ' Count various constructs
        Dim methodCount As Integer = System.Linq.Enumerable.Count(Of Microsoft.CodeAnalysis.VisualBasic.Syntax.MethodStatementSyntax)(System.Linq.Enumerable.OfType(Of Microsoft.CodeAnalysis.VisualBasic.Syntax.MethodStatementSyntax)(root.DescendantNodes()))
        Dim ifCount As Integer = System.Linq.Enumerable.Count(Of Microsoft.CodeAnalysis.VisualBasic.Syntax.IfStatementSyntax)(System.Linq.Enumerable.OfType(Of Microsoft.CodeAnalysis.VisualBasic.Syntax.IfStatementSyntax)(root.DescendantNodes()))
        Dim forCount As Integer = System.Linq.Enumerable.Count(Of Microsoft.CodeAnalysis.VisualBasic.Syntax.ForStatementSyntax)(System.Linq.Enumerable.OfType(Of Microsoft.CodeAnalysis.VisualBasic.Syntax.ForStatementSyntax)(root.DescendantNodes()))
        Dim forEachCount As Integer = System.Linq.Enumerable.Count(Of Microsoft.CodeAnalysis.VisualBasic.Syntax.ForEachStatementSyntax)(System.Linq.Enumerable.OfType(Of Microsoft.CodeAnalysis.VisualBasic.Syntax.ForEachStatementSyntax)(root.DescendantNodes()))
        Dim whileCount As Integer = System.Linq.Enumerable.Count(Of Microsoft.CodeAnalysis.VisualBasic.Syntax.WhileStatementSyntax)(System.Linq.Enumerable.OfType(Of Microsoft.CodeAnalysis.VisualBasic.Syntax.WhileStatementSyntax)(root.DescendantNodes()))
        Dim loopCount As Integer = forCount + forEachCount + whileCount

        metrics.AppendLine($"  Methods/Functions: {methodCount}")
        metrics.AppendLine($"  If statements: {ifCount}")
        metrics.AppendLine($"  Loops: {loopCount}")

        ' Simple cyclomatic complexity estimate
        Dim complexity As Integer = 1 + ifCount + loopCount
        metrics.AppendLine($"  Estimated cyclomatic complexity: {complexity}")

        If complexity > 10 Then
            result.Warnings.Add($"High complexity detected (complexity: {complexity})")
        End If

        result.ComplexityMetrics = metrics.ToString()
    End Sub

End Module
