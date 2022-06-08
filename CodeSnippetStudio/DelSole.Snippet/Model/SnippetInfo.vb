Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports <xmlns="http://schemas.microsoft.com/VisualStudio/2005/CodeSnippet">

''' <summary>
''' Structured representation of a code snippet file with full-fidelity
''' </summary>
Public Class SnippetInfo
    Implements IDataErrorInfo
    Implements INotifyPropertyChanged

    Private _snippetFileName As String
    Private _snippetLanguage As String
    Private _snippetPath As String
    Private _snippetDescription As String

    ''' <summary>
    ''' A code snippet's file name
    ''' </summary>
    ''' <returns>String</returns>
    Public Property SnippetFileName As String
        Get
            Return _snippetFileName
        End Get
        Set(value As String)
            _snippetFileName = value
            OnPropertyChanged(NameOf(SnippetFileName))
            CheckValue(NameOf(SnippetFileName))
        End Set
    End Property

    ''' <summary>
    ''' The programming language the snippet targets
    ''' </summary>
    ''' <returns>String</returns>
    Public Property SnippetLanguage As String
        Get
            Return _snippetLanguage
        End Get
        Set(value As String)
            _snippetLanguage = value
            OnPropertyChanged(NameOf(SnippetLanguage))
            CheckValue(NameOf(SnippetLanguage))
        End Set
    End Property

    ''' <summary>
    ''' The snippet's path on disk
    ''' </summary>
    ''' <returns>String</returns>
    Public Property SnippetPath As String
        Get
            Return _snippetPath
        End Get
        Set(value As String)
            _snippetPath = value
            OnPropertyChanged(NameOf(SnippetPath))
            CheckValue(NameOf(SnippetPath))
        End Set
    End Property

    ''' <summary>
    ''' The snippet's description
    ''' </summary>
    ''' <returns>String</returns>
    Public Property SnippetDescription As String
        Get
            Return _snippetDescription
        End Get
        Set(value As String)
            _snippetDescription = value
            OnPropertyChanged(NameOf(SnippetDescription))
            CheckValue(NameOf(SnippetDescription))
        End Set
    End Property

    ''' <summary>
    ''' Return the full pathname for a code snippet, combining the values of <seealso cref="SnippetPath"/> and <seealso cref="SnippetFileName"/> properties
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property SnippetPathName As String
        Get
            Return IO.Path.Combine(SnippetPath, SnippetFileName)
        End Get
    End Property

    ''' <summary>
    ''' Holds a collection of validation errors
    ''' </summary>
    Private validationErrors As New Dictionary(Of String, String)

    ''' <summary>
    ''' Add a validation error to the collection of validation errors
    ''' passing the property name and the error message
    ''' </summary>
    ''' <param name="columnName"></param>
    ''' <param name="msg"></param>
    Protected Sub AddError(ByVal columnName As String, ByVal msg As String)
        If Not validationErrors.ContainsKey(columnName) Then
            validationErrors.Add(columnName, msg)
        End If
    End Sub

    ''' <summary>
    ''' Remove a validation error from the collection
    ''' </summary>
    ''' <param name="columnName"></param>
    Protected Sub RemoveError(ByVal columnName As String)
        If validationErrors.ContainsKey(columnName) Then
            validationErrors.Remove(columnName)
        End If
    End Sub

    ''' <summary>
    ''' Return True if the current instance has validation errors
    ''' </summary>
    ''' <returns>Boolean</returns>
    Public Overridable ReadOnly Property HasErrors() As Boolean
        Get
            Return (validationErrors.Any)
        End Get
    End Property

    ''' <summary>
    ''' Return an error message
    ''' </summary>
    ''' <returns>String</returns>
    Public ReadOnly Property [Error]() As String _
        Implements System.ComponentModel.IDataErrorInfo.Error
        Get
            If validationErrors.Any Then
                Return $"{TypeName(Me)} data is invalid."
            Else
                Return Nothing
            End If
        End Get
    End Property

    ''' <summary>
    ''' Default property for the specified property name
    ''' </summary>
    ''' <param name="columnName"></param>
    ''' <returns>String</returns>
    Default Public ReadOnly Property Item(ByVal columnName As String) As String _
        Implements System.ComponentModel.IDataErrorInfo.Item
        Get
            If validationErrors.ContainsKey(columnName) Then
                Return validationErrors(columnName).ToString
            Else
                Return Nothing
            End If
        End Get
    End Property

    ''' <summary>
    ''' Check for required values and add/removes
    ''' validation errors
    ''' </summary>
    ''' <param name="value"></param>
    Private Sub CheckValue(ByVal value As String)
        Select Case value
            Case Is = "SnippetFileName"
                If Me.SnippetFileName = "" Or String.IsNullOrEmpty(Me.SnippetFileName) Then
                    Me.AddError("SnippetFileName", "Value cannot be null")
                Else
                    Me.RemoveError("SnippetFileName")
                End If
            Case Is = "SnippetLanguage"
                If Me.SnippetLanguage = "" Or String.IsNullOrEmpty(Me.SnippetLanguage) Then
                    Me.AddError("SnippetLanguage", "Value cannot be null")
                Else
                    Me.RemoveError("SnippetLanguage")
                End If
            Case Is = "SnippetPath"
                If Me.SnippetPath = "" Or String.IsNullOrEmpty(Me.SnippetPath) Then
                    Me.AddError("SnippetPath", "Value cannot be null")
                Else
                    Me.RemoveError("SnippetPath")
                End If
        End Select
    End Sub

    ''' <summary>
    ''' Detect the target programming language
    ''' for the specified snippet file
    ''' </summary>
    ''' <param name="snippetFile"></param>
    ''' <returns></returns>
    Public Shared Function GetSnippetLanguage(snippetFile As String) As String
        Try
            Dim doc As XDocument = XDocument.Load(snippetFile)

            Dim query = From line In doc...<Code>
                        Select line.@Language

            Return query.FirstOrDefault.ToUpper

        Catch ex As Exception
            Dim sn As String = IO.File.ReadAllText(snippetFile)

            If sn.ToUpper.Contains("CODE LANGUAGE=""VB""") Then
                Return "VB"
            ElseIf sn.ToUpper.Contains("CODE LANGUAGE=""CSHARP""") = True Then
                Return "CSharp"
            ElseIf sn.ToUpper.Contains("CODE LANGUAGE=""SQL""") Then
                Return "SQL"
            ElseIf sn.ToUpper.Contains("CODE LANGUAGE=""JAVASCRIPT""") = True Then
                Return "JAVASCRIPT"
            ElseIf sn.ToUpper.Contains("CODE LANGUAGE=""XML""") = True Then
                Return "XML"
            Else
                Throw New NotSupportedException(snippetFile & " is not a supported snippet")
            End If
        End Try

    End Function

    ''' <summary>
    ''' Detect the description for the specified snippet file. If the operation fails, return the "Description unavailable" message
    ''' </summary>
    ''' <remarks>Note: invoking GetSnippetDescription might fail if the code snippet schema is based on v 1.0</remarks>
    ''' <param name="snippetFile"></param>
    ''' <returns></returns>
    Public Shared Function GetSnippetDescription(snippetFile As String) As String
        Try
            Dim doc As XDocument = XDocument.Load(snippetFile)

            Dim query = From line In doc...<Description>
                        Select line.Value

            Return query.FirstOrDefault
        Catch ex As Exception
            Return "Description unavailable"
        End Try
    End Function

    ''' <summary>
    ''' Notify callers for changes
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Public Event PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements INotifyPropertyChanged.PropertyChanged

    ''' <summary>
    ''' Raise the PropertyChanged event when instance data changes
    ''' </summary>
    ''' <param name="name"></param>
    Protected Sub OnPropertyChanged(ByVal name As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(name))
    End Sub
End Class

''' <summary>
''' A collection of <seealso cref="SnippetInfo"></seealso> objects/>
''' </summary>
Public Class SnippetInfoCollection
    Inherits ObservableCollection(Of SnippetInfo)
End Class

