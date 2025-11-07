Imports System
Imports System.IO
Imports System.Reflection
Imports System.Text
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.Emit

Public Module VBCodeExecutor

    Public Function ExecuteVBCode(vbCodeString As String) As String
        Try
            Dim syntaxTree As SyntaxTree = VisualBasicSyntaxTree.ParseText(vbCodeString)
            
            Dim assemblyName As String = "DynamicAssembly_" & Guid.NewGuid().ToString("N")
            
            Dim references As MetadataReference() = {
                MetadataReference.CreateFromFile(GetType(Object).Assembly.Location),
                MetadataReference.CreateFromFile(GetType(Console).Assembly.Location),
                MetadataReference.CreateFromFile(Assembly.Load("System.Runtime").Location),
                MetadataReference.CreateFromFile(Assembly.Load("System.Console").Location),
                MetadataReference.CreateFromFile(Assembly.Load("netstandard").Location)
            }
            
            Dim compilation As VisualBasicCompilation = VisualBasicCompilation.Create(
                assemblyName,
                syntaxTrees:={syntaxTree},
                references:=references,
                options:=New VisualBasicCompilationOptions(OutputKind.DynamicallyLinkedLibrary)
            )
            
            Using ms As New MemoryStream()
                Dim result As EmitResult = compilation.Emit(ms)
                
                If Not result.Success Then
                    Dim failures = result.Diagnostics.Where(Function(diagnostic) diagnostic.IsWarningAsError OrElse diagnostic.Severity = DiagnosticSeverity.Error)
                    
                    Dim errorBuilder As New StringBuilder()
                    errorBuilder.AppendLine("Compilation Errors:")
                    For Each diagnostic In failures
                        errorBuilder.AppendLine($"  {diagnostic.Id}: {diagnostic.GetMessage()}")
                    Next
                    
                    Return errorBuilder.ToString()
                Else
                    ms.Seek(0, SeekOrigin.Begin)
                    Dim assembly As Assembly = Assembly.Load(ms.ToArray())
                    
                    Dim type = assembly.GetTypes().FirstOrDefault()
                    If type Is Nothing Then
                        Return "Error: No types found in compiled assembly. Make sure your code includes a Module or Class."
                    End If
                    
                    Dim method = type.GetMethods(BindingFlags.Public Or BindingFlags.Static).FirstOrDefault()
                    If method Is Nothing Then
                        Return "Error: No public shared methods found. Make sure your code includes a public shared function or sub."
                    End If
                    
                    Dim originalOut As TextWriter = Console.Out
                    Dim originalError As TextWriter = Console.Error
                    
                    Try
                        Dim outputBuilder As New StringBuilder()
                        Using sw As New StringWriter(outputBuilder)
                            Console.SetOut(sw)
                            Console.SetError(sw)
                            
                            Dim methodResult As Object = method.Invoke(Nothing, Nothing)
                            
                            Console.SetOut(originalOut)
                            Console.SetError(originalError)
                            
                            Dim output As String = outputBuilder.ToString()
                            
                            If methodResult IsNot Nothing Then
                                If String.IsNullOrEmpty(output) Then
                                    Return $"Result: {methodResult}"
                                Else
                                    Return $"{output}Result: {methodResult}"
                                End If
                            Else
                                If String.IsNullOrEmpty(output) Then
                                    Return "Execution completed successfully (no return value)."
                                Else
                                    Return output
                                End If
                            End If
                        End Using
                    Catch ex As Exception
                        Console.SetOut(originalOut)
                        Console.SetError(originalError)
                        Return $"Runtime Error: {ex.InnerException?.Message ?? ex.Message}" & vbCrLf & ex.StackTrace
                    End Try
                End If
            End Using
            
        Catch ex As Exception
            Return $"Fatal Error: {ex.Message}" & vbCrLf & ex.StackTrace
        End Try
    End Function

End Module
