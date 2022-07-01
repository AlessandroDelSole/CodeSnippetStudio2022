Option Strict On
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CSharp
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports <xmlns="http://schemas.microsoft.com/VisualStudio/2005/CodeSnippet">


''' <summary>
''' Represent a code snippet and expose APIs to create and save a code snippet file. 
''' It differs from <seealso cref="SnippetInfo"/> which instead serves as a building block for generating .Vsix packages
''' </summary>
Public Class CodeSnippet
    Implements INotifyPropertyChanged
    Implements IDataErrorInfo

#Region "Roslyn Code Analysis"

    Private _diagnostics As ObservableCollection(Of Diagnostic)
    ''' <summary>
    ''' Return a collection of Roslyn diagnostics for the current code snippet
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Diagnostics As ObservableCollection(Of Diagnostic)
        Get
            Return _diagnostics
        End Get
    End Property

    ''' <summary>
    ''' Perform Roslyn live analysis over the snippet's code (VB and C# only).
    ''' The analysis result is assigned to the <seealso cref="Diagnostics"/> property
    ''' </summary>
    Public Sub AnalyzeCode()
        Select Case Language.ToUpper
            Case "VB"
                Dim tree = VisualBasicSyntaxTree.ParseText(Code).GetRoot
                If tree.ContainsDiagnostics Then
                    _diagnostics = New ObservableCollection(Of Diagnostic)(tree.GetDiagnostics)
                Else
                    _diagnostics = Nothing
                End If
            Case "CSHARP"
                Dim tree = CSharpSyntaxTree.ParseText(Code).GetRoot
                If tree.ContainsDiagnostics Then
                    _diagnostics = New ObservableCollection(Of Diagnostic)(tree.GetDiagnostics)
                Else
                    _diagnostics = Nothing
                End If
            Case Else
                _diagnostics = Nothing
        End Select
    End Sub
#End Region

#Region "Backing fields"
    Private _author As String
    Private _title As String
    Private _description As String
    Private _helpUrl As String
    Private _shortcut As String
    Private _keywords As String
    Private _kind As CodeSnippetKinds
    Private _namespaces As [Imports]
    Private _references As References
    Private _declarations As Declarations
    Private _code As String
    Private _language As String
    Private _snippetType As SnippetType
#End Region

#Region "Properties"
    Public Property IsDirty As Boolean

    ''' <summary>
    ''' The code snippet author
    ''' </summary>
    ''' <returns>String</returns>
    <Category("Properties")>
    <DisplayName("Author")>
    <Description("Author of code snippet. This is a required value.")>
    Public Property Author As String
        Get
            Return _author
        End Get
        Set(value As String)
            _author = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(Author)))
            IsDirty = True
            CheckValue(NameOf(Author))
        End Set
    End Property

    ''' <summary>
    ''' The code snippet title
    ''' </summary>
    ''' <returns>String</returns>
    <Category("Properties")>
    <DisplayName("Title")>
    <Description("Title of code snippet. This is a required value.")>
    Public Property Title As String
        Get
            Return _title
        End Get
        Set(value As String)
            _title = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(Title)))
            IsDirty = True
            CheckValue(NameOf(Title))
        End Set
    End Property

    ''' <summary>
    ''' The code snippet description
    ''' </summary>
    ''' <returns>String</returns>
    <Category("Properties")>
    <DisplayName("Description")>
    <Description("Description of code snippet. This is a required value.")>
    Public Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(Description)))
            IsDirty = True
        End Set
    End Property

    ''' <summary>
    ''' An optional URL that users can search for more information
    ''' </summary>
    ''' <returns>String</returns>
    <Category("Properties")>
    <DisplayName("HelpUrl")>
    <Description("URL where users can find help. This is an optional value.")>
    Public Property HelpUrl As String
        Get
            Return _helpUrl
        End Get
        Set(value As String)
            _helpUrl = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(HelpUrl)))
            IsDirty = True
        End Set
    End Property

    ''' <summary>
    ''' A keyboard shortcut that users can use to insert the code snippet quickly
    ''' </summary>
    ''' <returns></returns>
    <Category("Properties")>
    <DisplayName("Shortcut")>
    <Description("Keyboard shortcut for IntelliSense. This is an optional value.")>
    Public Property Shortcut As String
        Get
            Return _shortcut
        End Get
        Set(value As String)
            _shortcut = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(Shortcut)))
            IsDirty = True
        End Set
    End Property

    ''' <summary>
    ''' A comma-separated list of words that help categorize a code snippet
    ''' </summary>
    ''' <returns></returns>
    <Category("Properties")>
    <DisplayName("Keywords")>
    <Description("Comma separated list of keywords that help identifying a code snippet.")>
    Public Property Keywords As String
        Get
            Return _keywords
        End Get
        Set(value As String)
            _keywords = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(Keywords)))
            IsDirty = True
        End Set
    End Property

    ''' <summary>
    ''' Represent the kind of code snippet
    ''' </summary>
    ''' <returns><seealso cref="CodeSnippetKinds"/></returns>
    <Category("Properties")>
    <DisplayName("Kind")>
    <Description("Determine the kind of code snippet.")>
    Public Property Kind As CodeSnippetKinds
        Get
            Return _kind
        End Get
        Set(value As CodeSnippetKinds)
            _kind = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(Kind)))
            IsDirty = True
        End Set
    End Property

    <Category("Properties")>
    <DisplayName("Type")>
    <Description("Expansion or SurroundsWith snippet. Expansion is default")>
    Public Property Type As SnippetType
        Get
            Return _snippetType
        End Get
        Set(value As SnippetType)
            _snippetType = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(Type)))
            IsDirty = True
        End Set
    End Property

    ''' <summary>
    ''' VB Only: a list of namespaces that are required for the code snippet to work, without the Imports keyword
    ''' </summary>
    ''' <returns><seealso cref="[Imports]"/></returns>
    <Category("Properties")>
    <DisplayName("Namespaces")>
    <Description("The list of VB namespaces that the code snippet requires to work.")>
    Public Property Namespaces As [Imports]
        Get
            Return _namespaces
        End Get
        Set(value As [Imports])
            _namespaces = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(Namespaces)))
            IsDirty = True
        End Set
    End Property

    ''' <summary>
    ''' VB only: a list of assembly references that are required for the code snippet to work
    ''' </summary>
    ''' <returns><seealso cref="References"/></returns>
    <Category("Properties")>
    <DisplayName("References")>
    <Description("The list of assembly references that the code snippet requires to work (VB only).")>
    Public Property References As References
        Get
            Return _references
        End Get
        Set(value As References)
            _references = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(References)))
            IsDirty = True
        End Set
    End Property

    ''' <summary>
    ''' The list of identifiers/words and related suggestions the IntelliSense shows 
    ''' </summary>
    ''' <returns><seealso cref="Declarations"/></returns>
    <Category("Properties")>
    <DisplayName("Declarations")>
    <Description("The list of words/suggestions that the code editor will display.")>
    Public Property Declarations As Declarations
        Get
            Return _declarations
        End Get
        Set(value As Declarations)
            _declarations = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(Declarations)))
            IsDirty = True
        End Set
    End Property

    ''' <summary>
    ''' The actual source code of the snippet
    ''' </summary>
    ''' <returns>String</returns>
    <Category("Properties")>
    <DisplayName("Code")>
    <Description("The actual code for the code snippet.")>
    Public Property Code As String
        Get
            Return _code
        End Get
        Set(value As String)
            _code = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(Code)))
            CheckValue(NameOf(Code))
            IsDirty = True
        End Set
    End Property

    ''' <summary>
    ''' The programming language the code snippet is written with.
    ''' Supported values are: VB, CSHARP, SQL, XML, CPP, HTML, JAVASCRIPT
    ''' </summary>
    ''' <returns></returns>
    <Category("Properties")>
    <DisplayName("Language")>
    <Description("The programming language the snippet is written with.")>
    Public Property Language As String
        Get
            Return _language
        End Get
        Set(value As String)
            _language = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(Language)))
            CheckValue(NameOf(Language))
            IsDirty = True
        End Set
    End Property

    Private _fileName As String
    Public ReadOnly Property FileName As String
        Get
            Return _fileName
        End Get
    End Property

#End Region

#Region "Data validation support"
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
            Return validationErrors.Any
        End Get
    End Property

    ''' <summary>
    ''' Return an error message
    ''' </summary>
    ''' <returns>String</returns>
    Public ReadOnly Property [Error]() As String _
        Implements IDataErrorInfo.Error
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
        Implements IDataErrorInfo.Item
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
            Case Is = "Title"
                If Title = "" Or String.IsNullOrEmpty(Title) Then
                    AddError("Title", "Value cannot be null")
                Else
                    RemoveError("Title")
                End If
            Case Is = "Code"
                If Code = "" Or String.IsNullOrEmpty(Code) Then
                    AddError("Code", "Value cannot be null")
                Else
                    RemoveError("Code")
                End If
            Case Is = "Author"
                If Author = "" Or String.IsNullOrEmpty(Author) Then
                    AddError("Author", "Value cannot be null")
                Else
                    RemoveError("Author")
                End If
            Case Is = "Language"
                If Language = "" Or String.IsNullOrEmpty(Language) Then
                    AddError("Language", "Value cannot be null")
                ElseIf supportedLanguages.Contains(Language.ToUpper) = False Then
                    AddError("Language", $"{Language} language is not supported")
                Else
                    RemoveError("Language")
                End If
        End Select
    End Sub
#End Region

    ''' <summary>
    ''' An array representing the supported programming languages
    ''' </summary>
    Private supportedLanguages() As String = {"VB", "CSHARP", "SQL", "XML", "XAML", "CPP", "HTML", "JAVASCRIPT"}

    ''' <summary>
    ''' Raised when the value of a property changes.
    ''' </summary>
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    ''' <summary>
    ''' Type initialization
    ''' </summary>
    Public Sub New()
        My.User.InitializeWithWindowsUser()
        Author = My.User.Name
        Keywords = ""
        HelpUrl = ""
        Kind = CodeSnippetKinds.MethodBody
        Namespaces = New [Imports]
        References = New References
        Declarations = New Declarations
        Type = SnippetType.Expansion
        IsDirty = False
    End Sub

#Region "Load and save snippets"
    ''' <summary>
    ''' Return a schema-compliant string representing the code snippet kind
    ''' </summary>
    ''' <param name="kind"></param>
    ''' <returns>String</returns>
    Private Shared Function ReturnSnippetKind(kind As CodeSnippetKinds) As String
        Dim snippetKind As String
        Select Case kind
            Case CodeSnippetKinds.MethodBody
                snippetKind = "method body"
            Case CodeSnippetKinds.MethodDeclaration
                snippetKind = "method decl"
            Case CodeSnippetKinds.File
                snippetKind = "file"
            Case CodeSnippetKinds.TypeDeclaration
                snippetKind = "type decl"
            Case Else
                snippetKind = "any"
        End Select

        Return snippetKind
    End Function

    ''' <summary>
    ''' Create and save a code snippet file to disk
    ''' </summary>
    ''' <param name="fileName">The target snippet file name</param>
    Private Sub SaveCodeSnippet(fileName As String)
        Dim snippetKind As String = ReturnSnippetKind(Kind)

        Dim editedCode = Code

        If Declarations.Any Then
            For Each decl In Declarations
                editedCode = editedCode.Replace(decl.Default, "$" & decl.ID & "$")
            Next
        End If

        Dim keywords = Me.Keywords.Split(","c).AsEnumerable

        Dim snippetType As String
        If Language = "VB" Then
            snippetType = "Expansion"
        Else
            If Me.Type = Snippet.SnippetType.SurroundsWith Then
                snippetType = "SurroundsWith"
            Else
                snippetType = "Expansion"
            End If
        End If

        Dim cdata As New XCData(editedCode)
        Dim doc = <?xml version="1.0" encoding="utf-8"?>
                  <CodeSnippets xmlns="http://schemas.microsoft.com/VisualStudio/2005/CodeSnippet">
                      <CodeSnippet Format="1.0.0">
                          <Header>
                              <Title><%= Title %></Title>
                              <Author><%= Author %></Author>
                              <Description><%= Description %></Description>
                              <HelpUrl><%= HelpUrl %></HelpUrl>
                              <SnippetTypes>
                                  <SnippetType><%= snippetType %></SnippetType>
                              </SnippetTypes>
                              <Keywords>
                                  <%= From key In keywords
                                      Select <Keyword><%= key %></Keyword> %>
                              </Keywords>
                              <Shortcut><%= Shortcut %></Shortcut>
                          </Header>
                          <Snippet>
                              <References>
                                  <%= From ref In References
                                      Select <Reference>
                                                 <Assembly><%= ref.Assembly %></Assembly>
                                                 <Url><%= ref.Url %></Url>
                                             </Reference> %>
                              </References>
                              <Imports>
                                  <%= From imp In Namespaces
                                      Select <Import>
                                                 <Namespace><%= imp.ImportDirective %></Namespace>
                                             </Import> %>
                              </Imports>
                              <Declarations>
                                  <%= From decl In Declarations
                                      Where decl.ReplacementType.ToLower = "object"
                                      Select <Object Editable="true">
                                                 <ID><%= decl.ID %></ID>
                                                 <Type><%= decl.Type %></Type>
                                                 <ToolTip><%= decl.ToolTip %></ToolTip>
                                                 <Default><%= decl.Default %></Default>
                                                 <Function><%= decl.Function %></Function>
                                             </Object> %>
                                  <%= From decl In Declarations
                                      Where decl.ReplacementType.ToLower = "literal"
                                      Select <Literal Editable="true">
                                                 <ID><%= decl.ID %></ID>
                                                 <ToolTip><%= decl.ToolTip %></ToolTip>
                                                 <Default><%= decl.Default %></Default>
                                                 <Function><%= decl.Function %></Function>
                                             </Literal> %>
                              </Declarations>
                              <Code Language=<%= Language %> Kind=<%= snippetKind %>
                                  Delimiter="$"><%= cdata %></Code>
                          </Snippet>
                      </CodeSnippet>
                  </CodeSnippets>
        doc.Save(fileName)
        _fileName = Path.GetFileName(fileName)
    End Sub

    ''' <summary>
    ''' Create a new <seealso cref="CodeSnippet"/> from a code file
    ''' </summary>
    ''' <param name="fileName">Must be a .vb, .cs, .sql, .xml, .xaml, .cpp, or .js file</param>
    ''' <returns></returns>
    Public Shared Function ImportCodeFile(fileName As String) As CodeSnippet
        If Not File.Exists(fileName) Then
            Throw New FileNotFoundException("File not found", fileName)
        End If

        Dim snippet As New CodeSnippet

        Dim fileExt = Path.GetExtension(fileName).ToLower

        Select Case fileExt
            Case = ".vb"
                snippet.Language = "VB"
            Case = ".cs"
                snippet.Language = "CSharp"
            Case = ".js"
                snippet.Language = "JavaScript"
            Case = ".cpp"
                snippet.Language = "CPP"
            Case = ".sql"
                snippet.Language = "SQL"
            Case = ".xml"
                snippet.Language = "XML"
            Case = ".xaml"
                snippet.Language = "XAML"
            Case Else
                snippet = Nothing
                Throw New NotSupportedException($"Files with {fileExt} extension are not supported")
        End Select

        snippet.Code = File.ReadAllText(fileName)

        Return snippet
    End Function

    ''' <summary>
    ''' Load a code snippet from disk and return an instance of <seealso cref="CodeSnippet"/>
    ''' </summary>
    ''' <param name="fileName"></param>
    ''' <returns><seealso cref="CodeSnippet"/></returns>
    Public Shared Function LoadSnippet(fileName As String) As CodeSnippet
        If Not File.Exists(fileName) Then
            Throw New FileNotFoundException("File not found", fileName)
        End If

        Dim snippet As CodeSnippet

        'snippet for VS Code
        If Path.GetExtension(fileName).ToLower.EndsWith("json") Then
            snippet = New CodeSnippet

            Dim json = File.ReadAllText(fileName)
            Dim pos = json.IndexOf(":"c)
            json = json.Substring(pos + 1)

            Dim snippetBody As New StringBuilder
            Dim actualJson = JObject.Parse(json)
            Dim codeJsonArray = actualJson.SelectToken("body")
            For Each lineOfCode In codeJsonArray
                snippetBody.AppendLine(lineOfCode.ToString)
            Next
            snippet.Author = ""
            snippet.Code = snippetBody.ToString
            Return snippet
        Else
            'Snippet for VS
            Dim doc = XDocument.Load(fileName)
            snippet = New CodeSnippet

            snippet.Author = doc...<Author>.Value
            snippet.Code = doc...<Code>.Value
            snippet.Description = doc...<Description>.Value
            snippet.HelpUrl = doc...<HelpUrl>.Value
            snippet.Language = doc...<Code>.@Language
            snippet.Shortcut = doc...<Shortcut>.Value
            snippet.Title = doc...<Title>.Value

            Dim kwQuery = From kw In doc...<Keyword>
                          Select kw.Value

            If kwQuery.Any Then
                For Each kw In kwQuery
                    snippet.Keywords = snippet.Keywords & "," & kw
                Next
            End If

            If snippet.Language.ToUpper = "VB" Or snippet.Language.ToUpper = "CSHARP" Then
                Dim refQuery = From ref In doc...<Reference>
                               Select New Reference With {.Assembly = ref.<Assembly>.Value, .Url = ref.<Url>.Value}

                If refQuery.Any Then
                    For Each ref In refQuery
                        snippet.References.Add(ref)
                    Next
                End If

                Dim impQuery = From imp In doc...<Import>
                               Select New Import With {.ImportDirective = imp.<Namespace>.Value}

                If impQuery.Any Then
                    For Each imp In impQuery
                        snippet.Namespaces.Add(imp)
                    Next
                End If
            End If

            Dim litQuery = From decl In doc...<Literal>
                           Select New Declaration With {.ID = decl.<ID>.Value,
.Default = decl.<Default>.Value,
.Editable = True,
.Function = decl.<Function>.Value,
.ToolTip = decl.<ToolTip>.Value,
.Type = decl.<Type>.Value,
.ReplacementType = "Literal"}

            Dim objQuery = From decl In doc...<Object>
                           Select New Declaration With {.ID = decl.<ID>.Value,
.Default = decl.<Default>.Value,
.Editable = True,
.Function = decl.<Function>.Value,
.ToolTip = decl.<ToolTip>.Value,
.Type = decl.<Type>.Value,
.ReplacementType = "Object"}

            If litQuery.Any Then
                For Each lit In litQuery
                    snippet.Declarations.Add(lit)
                Next
            End If

            If objQuery.Any Then
                For Each obj In objQuery
                    snippet.Declarations.Add(obj)
                Next
            End If

            Try
                Select Case doc...<Code>.@Kind.ToLower
                    Case = "file"
                        snippet.Kind = CodeSnippetKinds.File
                    Case = "method body"
                        snippet.Kind = CodeSnippetKinds.MethodBody
                    Case = "method decl"
                        snippet.Kind = CodeSnippetKinds.MethodDeclaration
                    Case = "type decl"
                        snippet.Kind = CodeSnippetKinds.TypeDeclaration
                    Case Else
                        snippet.Kind = CodeSnippetKinds.Any
                End Select
            Catch ex As Exception
                snippet.Kind = CodeSnippetKinds.Any
            End Try

            For Each decl In snippet.Declarations
                snippet.Code = snippet.Code.Replace($"{doc...<Code>.@Delimiter}{decl.ID}{doc...<Code>.@Delimiter}", decl.Default)
            Next
            snippet._fileName = Path.GetFileName(fileName)
            Return snippet

        End If
    End Function

    ''' <summary>
    ''' Save a code snippet for Visual Studio Code
    ''' </summary>
    ''' <param name="fileName"></param>
    Private Sub SaveVSCodeSnippet(fileName As String)
        Dim editedCode = Code

        If Declarations.Any Then
            For Each decl In Declarations
                editedCode = editedCode.Replace(decl.Default, "${" & decl.ID & "}")
            Next
        End If

        Dim TextLines() As String = editedCode.Split(Environment.NewLine.ToCharArray)

        Using str As New StreamWriter(fileName)
            Using jw As New JsonTextWriter(str)
                jw.Formatting = Formatting.Indented
                jw.WriteStartObject()
                jw.WritePropertyName(Title)
                jw.WriteStartObject()
                jw.WritePropertyName("prefix")
                jw.WriteValue(Shortcut)
                jw.WritePropertyName("body")
                jw.WriteStartArray()

                For Each line As String In TextLines
                    If Not line.Length = 0 Then jw.WriteValue(line)
                Next
                jw.WriteEndArray()
                jw.WritePropertyName("description")
                jw.WriteValue(Description)
                jw.WriteEndObject()
                jw.WriteEndObject()
            End Using
        End Using
        _fileName = Path.GetFileName(fileName)
    End Sub

    ''' <summary>
    ''' Create and save a code snippet file to disk
    ''' </summary>
    ''' <param name="fileName">The target snippet file name</param>
    Public Sub SaveSnippet(fileName As String)
        If HasErrors Then
            Throw New InvalidOperationException("The supplied instance of CodeSnippet has errors that must be fixed first.")
        End If

        SaveCodeSnippet(fileName)
        IsDirty = False
    End Sub

    Private Shared Function ParseSublimeScopeSourceIntoSnippetLanguage(scope As String) As String
        scope = scope.ToLowerInvariant()
        Select Case scope
            Case = "source.sql"
                Return "SQL"
            Case = "source.cs"
                Return "CSharp"
            Case = "source.js"
                Return "JavaScript"
            Case = "text.xml"
                Return "XML"
            Case = "text.html"
                Return "HTML"
            Case = "source.c++"
                Return "CPP"
            Case Else   ' Must be fixed with a different default
                Return "VB"
        End Select
    End Function

    ''' <summary>
    ''' Convert a Sublime code snippet into an IntelliSense code snippet for Visual Studio
    ''' </summary>
    ''' <param name="fileName"></param>
    ''' <remarks>AS OF THIS VERSION, REPLACEMENTS ARE NOT PARSED. WORK IN PROGRESS</remarks>
    ''' <returns><seealso cref="CodeSnippet"/></returns>
    Public Shared Function ImportSublimeSnippet(fileName As String) As CodeSnippet
        If Not File.Exists(fileName) Then
            Throw New FileNotFoundException("File not found", fileName)
        End If

        Dim snippet As CodeSnippet

        Dim doc = XDocument.Load(fileName)
        snippet = New CodeSnippet

        My.User.InitializeWithWindowsUser()
        snippet.Author = My.User.Name
        snippet.Code = doc.Element("snippet").Element("content")?.Value
        snippet.Shortcut = doc.Element("snippet").Element("tabTrigger")?.Value
        snippet.Description = doc.Element("snippet").Element("description")?.Value
        snippet.Language = ParseSublimeScopeSourceIntoSnippetLanguage(doc.Element("snippet").Element("scope")?.Value)
        snippet.Title = IO.Path.GetFileNameWithoutExtension(fileName)
        Return snippet
    End Function
End Class

''' <summary>
''' Enumerate the code snippet kinds. 
''' Kind means the code snippet can be a method body, 
''' a method declaration, a type declaration, an entire source file,
''' or any other kind
''' </summary>
Public Enum CodeSnippetKinds
    MethodBody
    MethodDeclaration
    TypeDeclaration
    File
    Any
End Enum

#End Region



