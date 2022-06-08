Imports System.Collections.ObjectModel


''' <summary>
''' Represent a reference to a .NET assembly (VB Only).
''' The XML schema for code snippets does not support C# references
''' </summary>
Public Class Reference
    Property Assembly As String
    Property Url As String
End Class

''' <summary>
''' A collection of <seealso cref="Reference"/> objects
''' </summary>
Public Class References
    Inherits ObservableCollection(Of Reference)
End Class

