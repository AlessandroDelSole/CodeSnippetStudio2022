Imports System.Collections.ObjectModel
Imports System.ComponentModel

Public Class SnippetFolder
    Implements INotifyPropertyChanged

    Private _folderName As String
    Private _subFolders As ObservableCollection(Of SnippetFolder)
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    ''' <summary>
    ''' The full directory name
    ''' </summary>
    ''' <returns></returns>
    Public Property FolderName As String
        Get
            Return _folderName
        End Get
        Set(value As String)
            _folderName = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(FolderName)))
        End Set
    End Property

    Public Sub New(folderName As String, criteria As String)
        Me.FolderName = folderName
        If criteria Is Nothing Or String.IsNullOrEmpty(criteria) Then
            SnippetFiles = PopulateSnippetFileList()
        Else
            SnippetFiles = PopulateSnippetFileList(criteria)
        End If
    End Sub

    Private _snippetFiles As IEnumerable(Of CodeSnippet)
    ''' <summary>
    ''' Return the list of snippet file names in the current folder
    ''' </summary>
    ''' <returns></returns>
    Public Property SnippetFiles As IEnumerable(Of CodeSnippet)
        Get
            Return _snippetFiles
        End Get
        Set(value As IEnumerable(Of CodeSnippet))
            _snippetFiles = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(SnippetFiles)))
        End Set
    End Property

    Public Function PopulateSnippetFileList() As IEnumerable(Of CodeSnippet)
        Dim files = From ff In IO.Directory.EnumerateFiles(FolderName)
                    Where IO.Path.GetExtension(ff).ToLower.Contains("snippet")
                    Select ff

        If files.Any Then
            Dim snips As New List(Of CodeSnippet)
            For Each snip In files
                Try
                    Dim newSnip = CodeSnippet.LoadSnippet(snip)
                    snips.Add(newSnip)
                Catch ex As Exception
                    'Error, ignore snippet
                End Try
            Next
            Return snips.AsEnumerable
        Else
            Return Nothing
        End If
    End Function

    Public Function PopulateSnippetFileList(criteria As String) As IEnumerable(Of CodeSnippet)
        Dim files = From ff In IO.Directory.EnumerateFiles(FolderName)
                    Where IO.Path.GetFileName(ff).ToLower.Contains(criteria) And IO.Path.GetExtension(ff).ToLower.Contains("snippet")
                    Select ff

        If files.Any Then
            Dim snips As New List(Of CodeSnippet)
            For Each snip In files
                Try
                    Dim newSnip = CodeSnippet.LoadSnippet(snip)
                    snips.Add(newSnip)
                Catch ex As Exception
                    'Error, ignore snippet
                End Try
            Next
            Return snips.AsEnumerable
        Else
            Return Nothing
        End If
    End Function
End Class

