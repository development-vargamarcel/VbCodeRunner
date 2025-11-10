''' <summary>
''' This module provides stub definitions for BPM-specific types and functions
''' that are commonly used in dynamic VB code but not available in the standard runtime.
''' </summary>
Public Module BPMStubs

    ''' <summary>
    ''' Stub class for dbarray - a custom array type used in BPM systems
    ''' </summary>
    Public Class dbarray
        Private _items As New System.Collections.Generic.List(Of Object)()

        Public Sub New()
        End Sub

        Public Sub New(ParamArray items As Object())
            If items IsNot Nothing Then
                _items.AddRange(items)
            End If
        End Sub

        Public Sub Add(item As Object)
            _items.Add(item)
        End Sub

        Public Function Item(index As Integer) As Object
            If index >= 0 AndAlso index < _items.Count Then
                Return _items(index)
            End If
            Return Nothing
        End Function

        Default Public Property Items(index As Integer) As Object
            Get
                Return Item(index)
            End Get
            Set(value As Object)
                If index >= 0 AndAlso index < _items.Count Then
                    _items(index) = value
                ElseIf index = _items.Count Then
                    _items.Add(value)
                End If
            End Set
        End Property

        Public ReadOnly Property Count() As Integer
            Get
                Return _items.Count
            End Get
        End Property

        Public Function ToArray() As Object()
            Return _items.ToArray()
        End Function
    End Class

    ''' <summary>
    ''' Stub class for AssocArray - an associative array (dictionary-like) type
    ''' </summary>
    Public Class AssocArray
        Private _dictionary As New System.Collections.Generic.Dictionary(Of String, Object)()

        Public Sub New()
        End Sub

        Public Sub Add(key As String, value As Object)
            _dictionary(key) = value
        End Sub

        Public Function Item(key As String) As Object
            Dim value As Object = Nothing
            If _dictionary.TryGetValue(key, value) Then
                Return value
            End If
            Return Nothing
        End Function

        Default Public Property Items(key As String) As Object
            Get
                Return Item(key)
            End Get
            Set(value As Object)
                _dictionary(key) = value
            End Set
        End Property

        Public ReadOnly Property Count() As Integer
            Get
                Return _dictionary.Count
            End Get
        End Property

        Public ReadOnly Property Keys() As System.Collections.Generic.ICollection(Of String)
            Get
                Return _dictionary.Keys
            End Get
        End Property

        Public Function ContainsKey(key As String) As Boolean
            Return _dictionary.ContainsKey(key)
        End Function
    End Class

    ''' <summary>
    ''' Stub for FORMFIELDS - typically a collection of form field values in BPM
    ''' </summary>
    Public ReadOnly FORMFIELDS As New AssocArray()

    ''' <summary>
    ''' Stub for DB - typically a database connection or query object in BPM
    ''' </summary>
    Public ReadOnly DB As New AssocArray()

    ''' <summary>
    ''' Stub for CODDOC - typically represents a document code in BPM systems
    ''' </summary>
    Public CODDOC As String = ""

    ''' <summary>
    ''' Stub function for OpenTable - opens a database table and returns a recordset/array
    ''' </summary>
    Public Function OpenTable(tableName As String, Optional filter As String = "", Optional orderBy As String = "") As dbarray
        System.Console.WriteLine($"[STUB] OpenTable called: Table={tableName}, Filter={filter}, OrderBy={orderBy}")
        Return New dbarray()
    End Function

    ''' <summary>
    ''' Stub function for BPMSendMail - sends an email in BPM systems
    ''' </summary>
    Public Function BPMSendMail(
        toAddress As String,
        Optional subject As String = "",
        Optional body As String = "",
        Optional fromAddress As String = "",
        Optional ccAddress As String = "",
        Optional bccAddress As String = "",
        Optional attachments As Object = Nothing) As Boolean

        System.Console.WriteLine($"[STUB] BPMSendMail called: To={toAddress}, Subject={subject}")
        Return True
    End Function

    ''' <summary>
    ''' Stub function for BPMWorkflowWriteEvent - writes a workflow event in BPM systems
    ''' </summary>
    Public Function BPMWorkflowWriteEvent(
        eventType As String,
        Optional description As String = "",
        Optional details As Object = Nothing) As Boolean

        System.Console.WriteLine($"[STUB] BPMWorkflowWriteEvent called: EventType={eventType}, Description={description}")
        Return True
    End Function

    ''' <summary>
    ''' Replacement for System.Empty which is not accessible (Friend-level)
    ''' This provides a public alternative
    ''' </summary>
    Public ReadOnly Property EmptyValue() As Object
        Get
            Return Nothing
        End Get
    End Property

End Module
