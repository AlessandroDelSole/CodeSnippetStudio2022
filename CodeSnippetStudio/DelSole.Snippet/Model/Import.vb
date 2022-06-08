Imports System.Collections.ObjectModel


''' <summary>
''' Represent an Imports directive (VB only). 
''' The XML schema for code snippets does not support C# using
''' </summary>
Public Class Import
    ''' <summary>
    ''' The namespace without the Imports keyword (e.g. System.Diagnostics)
    ''' </summary>
    ''' <returns>String</returns>
    Property ImportDirective As String
End Class

''' <summary>
''' A collection of <seealso cref="Import"/> directives/>
''' </summary>
Public Class [Imports]
    Inherits ObservableCollection(Of Import)
End Class
