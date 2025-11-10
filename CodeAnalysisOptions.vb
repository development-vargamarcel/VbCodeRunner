''' <summary>
''' Configuration options for code analysis before execution
''' </summary>
Public Class CodeAnalysisOptions
    ''' <summary>
    ''' Enable or disable code analysis entirely
    ''' </summary>
    Public Property Enabled As Boolean = False

    ''' <summary>
    ''' Check for syntax errors
    ''' </summary>
    Public Property CheckSyntax As Boolean = True

    ''' <summary>
    ''' Check for semantic errors (undeclared variables, type mismatches, etc.)
    ''' </summary>
    Public Property CheckSemantics As Boolean = True

    ''' <summary>
    ''' Report potential security issues (e.g., System.IO usage, reflection)
    ''' </summary>
    Public Property CheckSecurity As Boolean = False

    ''' <summary>
    ''' Report code complexity metrics
    ''' </summary>
    Public Property ReportComplexity As Boolean = False

    ''' <summary>
    ''' Treat warnings as errors (prevent execution if warnings found)
    ''' </summary>
    Public Property WarningsAsErrors As Boolean = False

    ''' <summary>
    ''' Create a default CodeAnalysisOptions with analysis disabled
    ''' </summary>
    Public Shared Function CreateDisabled() As CodeAnalysisOptions
        Return New CodeAnalysisOptions With {.Enabled = False}
    End Function

    ''' <summary>
    ''' Create a CodeAnalysisOptions with basic analysis enabled (syntax and semantics only)
    ''' </summary>
    Public Shared Function CreateBasic() As CodeAnalysisOptions
        Return New CodeAnalysisOptions With {
            .Enabled = True,
            .CheckSyntax = True,
            .CheckSemantics = True,
            .CheckSecurity = False,
            .ReportComplexity = False
        }
    End Function

    ''' <summary>
    ''' Create a CodeAnalysisOptions with comprehensive analysis enabled
    ''' </summary>
    Public Shared Function CreateComprehensive() As CodeAnalysisOptions
        Return New CodeAnalysisOptions With {
            .Enabled = True,
            .CheckSyntax = True,
            .CheckSemantics = True,
            .CheckSecurity = True,
            .ReportComplexity = True
        }
    End Function
End Class
