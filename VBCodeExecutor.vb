Imports System
Imports System.Globalization
Imports System.IO
Imports System.Reflection
Imports System.Text
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.Emit

Public Module VBCodeExecutor

    Public Function ExecuteVBCodeWithVariables(vbCodeString As String, variables As Dictionary(Of String, Object)) As String
        Try
            Dim modifiedCode As String = InjectVariables(vbCodeString, variables)
            Return ExecuteVBCodeInternal(modifiedCode, Nothing)
        Catch ex As Exception
            Return $"Fatal Error: {ex.Message}" & vbCrLf & ex.StackTrace
        End Try
    End Function

    Public Function ExecuteVBCode(vbCodeString As String, ParamArray parameters As Object()) As String
        Return ExecuteVBCodeInternal(vbCodeString, parameters)
    End Function

    Private Function InjectVariables(vbCodeString As String, variables As Dictionary(Of String, Object)) As String
        If variables Is Nothing OrElse variables.Count = 0 Then
            Return vbCodeString
        End If

        Dim variableDeclarations As New StringBuilder()
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
                valueStr = CInt(varValue).ToString(CultureInfo.InvariantCulture)
            ElseIf TypeOf varValue Is Long Then
                valueStr = CLng(varValue).ToString(CultureInfo.InvariantCulture)
            ElseIf TypeOf varValue Is Short Then
                valueStr = CShort(varValue).ToString(CultureInfo.InvariantCulture)
            ElseIf TypeOf varValue Is Byte Then
                valueStr = CByte(varValue).ToString(CultureInfo.InvariantCulture)
            ElseIf TypeOf varValue Is Double Then
                valueStr = CDbl(varValue).ToString("R", CultureInfo.InvariantCulture)
            ElseIf TypeOf varValue Is Single Then
                valueStr = CSng(varValue).ToString("R", CultureInfo.InvariantCulture)
            ElseIf TypeOf varValue Is Decimal Then
                valueStr = CDec(varValue).ToString(CultureInfo.InvariantCulture) & "D"
            Else
                valueStr = $"DirectCast(Nothing, Object)"
            End If
            
            variableDeclarations.AppendLine($"        Dim {varName} As Object = {valueStr}")
        Next
        
        Dim moduleStartIndex As Integer = vbCodeString.IndexOf("Public Module", StringComparison.OrdinalIgnoreCase)
        If moduleStartIndex = -1 Then
            moduleStartIndex = vbCodeString.IndexOf("Module", StringComparison.OrdinalIgnoreCase)
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
            Dim syntaxTree As SyntaxTree = VisualBasicSyntaxTree.ParseText(vbCodeString)
            
            Dim assemblyName As String = "DynamicAssembly_" & Guid.NewGuid().ToString("N")
            
            Dim references As MetadataReference() = {
                MetadataReference.CreateFromFile(GetType(Object).Assembly.Location),
                MetadataReference.CreateFromFile(GetType(Console).Assembly.Location),
                MetadataReference.CreateFromFile(Assembly.Load("System.Runtime").Location),
                MetadataReference.CreateFromFile(Assembly.Load("System.Console").Location),
                MetadataReference.CreateFromFile(Assembly.Load("netstandard").Location),
                MetadataReference.CreateFromFile(GetType(Microsoft.VisualBasic.CompilerServices.StandardModuleAttribute).Assembly.Location)
            }
            
            Dim globalImports = {
                GlobalImport.Parse("System"),
                GlobalImport.Parse("System.Collections.Generic"),
                GlobalImport.Parse("System.Linq")
            }
            
            Dim compilationOptions = New VisualBasicCompilationOptions(OutputKind.DynamicallyLinkedLibrary).
                WithGlobalImports(globalImports)
            
            Dim compilation As VisualBasicCompilation = VisualBasicCompilation.Create(
                assemblyName,
                syntaxTrees:={syntaxTree},
                references:=references,
                options:=compilationOptions
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
                    
                    Dim methods = type.GetMethods(BindingFlags.Public Or BindingFlags.Static)
                    Dim method As MethodInfo = Nothing
                    
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
                    
                    Dim originalOut As TextWriter = Console.Out
                    Dim originalError As TextWriter = Console.Error
                    
                    Try
                        Dim outputBuilder As New StringBuilder()
                        Using sw As New StringWriter(outputBuilder)
                            Console.SetOut(sw)
                            Console.SetError(sw)
                            
                            Dim methodResult As Object = method.Invoke(Nothing, parameters)
                            
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
                        Dim errorMessage As String = If(ex.InnerException IsNot Nothing, ex.InnerException.Message, ex.Message)
                        Return $"Runtime Error: {errorMessage}" & vbCrLf & ex.StackTrace
                    End Try
                End If
            End Using
            
        Catch ex As Exception
            Return $"Fatal Error: {ex.Message}" & vbCrLf & ex.StackTrace
        End Try
    End Function

End Module
