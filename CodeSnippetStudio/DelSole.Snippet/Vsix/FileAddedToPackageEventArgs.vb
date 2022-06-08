Public Class FileAddedToPackageEventArgs
    Inherits EventArgs

    Public ReadOnly Property FileName As String

    Public Sub New(fileName As String)
        Me.FileName = fileName
    End Sub
End Class