Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports Ionic.Zip

Public Class SnippetLibrary
    Implements INotifyPropertyChanged

    Private _folders As ObservableCollection(Of SnippetFolder)

    Public Sub New()
        Me.Folders = New ObservableCollection(Of SnippetFolder)
    End Sub

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    ''' <summary>
    ''' Return the collection of folders in the library
    ''' </summary>
    ''' <returns></returns>
    Public Property Folders As ObservableCollection(Of SnippetFolder)
        Get
            Return _folders
        End Get
        Set(value As ObservableCollection(Of SnippetFolder))
            _folders = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(Folders)))
        End Set
    End Property

    ''' <summary>
    ''' Return the full list of code snippets inside the library. 
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property SnippetFiles As IEnumerable(Of String)
        Get

            If Not Me.Folders.Any Then
                Throw New InvalidOperationException("The folder collection is empty.")
            End If

            Dim fileNames As New List(Of String)
            For Each folder In Me.Folders
                Dim files =
                        IO.Directory.EnumerateFiles(folder.FolderName, "*.*", IO.SearchOption.AllDirectories).
                        Where(Function(f) IO.Path.GetExtension(f).ToLowerInvariant.Contains("snippet"))

                If files.Any Then
                    For Each fileName In files
                        fileNames.Add(fileName)
                    Next
                End If
            Next
            Return fileNames.AsEnumerable
        End Get
    End Property

    ''' <summary>
    ''' Save a library of folders containing code snippets
    ''' </summary>
    ''' <param name="pathName"></param>
    Public Sub SaveLibrary(pathName As String)
        Dim doc = <?xml version="1.0" encoding="utf-8"?>
                  <Folders>
                      <%= From fold In Me.Folders
                          Select <Folder FolderName=<%= fold.FolderName %>/> %>
                  </Folders>

        doc.Save(pathName)
    End Sub

    ''' <summary>
    ''' Load a snippet library list from disk. Make sure you load a list that was previously saved with the <seealso cref="SaveLibrary(String)"/> method.
    ''' </summary>
    ''' <param name="pathName"></param>
    ''' <remarks>If a folder in the library is not found on disk, then it's not added to the snippet library</remarks>
    Public Sub LoadLibrary(pathName As String)
        Dim doc = XDocument.Load(pathName)
        Dim query = From fold In doc...<Folder>
                    Select fold.@FolderName

        For Each element In query
            If IO.Directory.Exists(element) Then
                Dim folder As New SnippetFolder(element, Nothing)
                Me.Folders.Add(folder)
            End If
        Next
    End Sub

    ''' <summary>
    ''' Create a zip backup archive containing all the code snippet files in the library, maintaining the directory tree
    ''' </summary>
    ''' <param name="zipName"></param>
    Public Sub BackupLibraryToZip(zipName As String)
        Dim zip As New ZipFile(zipName)
        Try
            For Each fold In Folders
                Dim files = IO.Directory.EnumerateFiles(fold.FolderName).
                                Where(Function(f) IO.Path.GetExtension(f).ToLower = ".snippet" _
                                Or IO.Path.GetExtension(f).ToLower = ".json")

                zip.AddFiles(files)
            Next

            zip.Save()
        Catch ex As ArgumentException
            If IO.File.Exists(zipName) Then IO.File.Delete(zipName)
            BackupLibraryToZip(zipName)
        Catch ex As Exception
            Throw
        End Try
    End Sub
End Class

