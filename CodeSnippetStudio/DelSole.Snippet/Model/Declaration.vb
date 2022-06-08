Imports System.Collections.ObjectModel
Imports System.ComponentModel

''' <summary>
''' Represent a code replacement. Replacements are highlighted in the code editor and the user can
''' follow the suggestions offered by IntelliSense
''' </summary>
Public Class Declaration
    Implements INotifyPropertyChanged

    Private _editable As Boolean
    Public Property Editable As Boolean
        Get
            Return _editable
        End Get
        Set(value As Boolean)
            _editable = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(Editable)))
        End Set
    End Property

    Private _id As String

    ''' <summary>
    ''' The replacement ID
    ''' </summary>
    Public Property ID As String
        Get
            Return _id
        End Get
        Set(value As String)
            _id = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(ID)))
        End Set
    End Property

    Private _type As String
    ''' <summary>
    ''' The .NET type of the object 
    ''' </summary>
    ''' <returns></returns>
    Public Property [Type] As String
        Get
            Return _type
        End Get
        Set(value As String)
            _type = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(Type)))
        End Set
    End Property

    Private _toolTip As String
    ''' <summary>
    ''' A description explaining how the user can replace the code
    ''' </summary>
    ''' <returns></returns>
    Public Property ToolTip As String
        Get
            Return _toolTip
        End Get
        Set(value As String)
            _toolTip = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(ToolTip)))
        End Set
    End Property

    Private _default As String
    ''' <summary>
    ''' The default value for the replacement
    ''' </summary>
    ''' <returns></returns>
    Public Property [Default] As String
        Get
            Return _default
        End Get
        Set(value As String)
            _default = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf([Default])))
        End Set
    End Property

    Private _function As String
    ''' <summary>
    ''' A method that will be invoked when the snippet is inserted
    ''' </summary>
    ''' <returns></returns>
    Public Property [Function] As String
        Get
            Return _function
        End Get
        Set(value As String)
            _function = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf([Function])))
        End Set
    End Property

    Private _replacementType As String
    Public Property ReplacementType As String
        Get
            Return _replacementType
        End Get
        Set(value As String)
            _replacementType = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(ReplacementType)))
        End Set
    End Property

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Public Sub New()
        Me.Editable = True
        Me.ReplacementType = "Literal"
    End Sub
End Class

''' <summary>
''' A collection of <seealso cref="Declaration"/> objects
''' </summary>
Public Class Declarations
    Inherits ObservableCollection(Of Declaration)
End Class
