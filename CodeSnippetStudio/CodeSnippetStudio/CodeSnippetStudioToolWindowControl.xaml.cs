using CodeSnippetStudio.Properties;
using DelSole.Snippet;
using Microsoft.CodeAnalysis;
using Microsoft.Win32;
using Newtonsoft.Json;
using Syncfusion.UI.Xaml.Grid;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Navigation;

namespace CodeSnippetStudio
{
    /// <summary>
    /// Interaction logic for CodeSnippetStudioToolWindowControl.
    /// </summary>
    public partial class CodeSnippetStudioToolWindowControl : UserControl
    {
        private VsixPackage VsixData { get; set; }
        private CodeSnippet SnippetData { get; set; }
        private ObservableCollection<Uri> IntelliSenseReferences { get; set; }
        private SnippetLibrary snippetLib { get; set; }

        private string LibraryName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "CodeSnippetStudioLibrary.xml");
        private Brush selectionBackground;

        /// <summary>
        /// Initializes a new instance of the <see cref="CodeSnippetStudioToolWindowControl"/> class.
        /// </summary>
        public CodeSnippetStudioToolWindowControl()
        {
            this.InitializeComponent();
        }


        private void ResetPkg()
        {
            this.VsixData = new VsixPackage();

            this.VsixGrid.DataContext = this.VsixData;
            this.PackageTab.Focus();
        }

        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }

        private void HidePropertiesFromPropertyGrid()
        {
            // Properties that must be hidden from the PropertyGrid
            this.SnippetPropertyGrid.HidePropertiesCollection.Add("Namespaces");
            this.SnippetPropertyGrid.HidePropertiesCollection.Add("Declarations");
            this.SnippetPropertyGrid.HidePropertiesCollection.Add("References");
            this.SnippetPropertyGrid.HidePropertiesCollection.Add("Language");
            this.SnippetPropertyGrid.HidePropertiesCollection.Add("Code");
            this.SnippetPropertyGrid.HidePropertiesCollection.Add("Error");
            this.SnippetPropertyGrid.HidePropertiesCollection.Add("HasErrors");
            this.SnippetPropertyGrid.HidePropertiesCollection.Add("IsDirty");
            this.SnippetPropertyGrid.HidePropertiesCollection.Add("FileName");
            this.SnippetPropertyGrid.HidePropertiesCollection.Add("Diagnostics");
            SnippetPropertyGrid.RefreshPropertygrid();
        }

        private void LoadSnippetLibrary()
        {
            this.snippetLib = new SnippetLibrary();
            this.LibraryTreeview.ItemsSource = snippetLib.Folders;

            Settings.Default.LibraryName = LibraryName;

            this.EditorRoot.DataContext = this.SnippetData;
            try
            {
                snippetLib.LoadLibrary(Settings.Default.LibraryName);
            }
            catch
            {
                // error loading library, ignore
            }
        }

        private Syncfusion.Windows.Edit.Languages LoadPreferredLanguage()
        {
            switch (Settings.Default.PreferredLanguage)
            {
                case "VB":
                    {
                        this.LanguageCombo.SelectedIndex = 0;
                        this.PrefLanguageCombo.SelectedIndex = 0;
                        return Syncfusion.Windows.Edit.Languages.VisualBasic;
                    }
                case "CSharp":
                    {
                        this.LanguageCombo.SelectedIndex = 1;
                        this.PrefLanguageCombo.SelectedIndex = 1;
                        return Syncfusion.Windows.Edit.Languages.CSharp;
                    }
                case "SQL":
                    {
                        this.LanguageCombo.SelectedIndex = 2;
                        this.PrefLanguageCombo.SelectedIndex = 2;
                        return Syncfusion.Windows.Edit.Languages.SQL;
                    }
                case "XML":
                    {
                        this.LanguageCombo.SelectedIndex = 3;
                        this.PrefLanguageCombo.SelectedIndex = 3;
                        return Syncfusion.Windows.Edit.Languages.XML;
                    }
                case "XAML":
                    {
                        this.LanguageCombo.SelectedIndex = 4;
                        this.PrefLanguageCombo.SelectedIndex = 4;
                        return Syncfusion.Windows.Edit.Languages.XAML;
                    }
                case "CPP":
                    {
                        this.LanguageCombo.SelectedIndex = 5;
                        this.PrefLanguageCombo.SelectedIndex = 5;
                        return Syncfusion.Windows.Edit.Languages.CSharp;
                    }
                case "HTML":
                    {
                        this.LanguageCombo.SelectedIndex = 6;
                        this.PrefLanguageCombo.SelectedIndex = 6;
                        return Syncfusion.Windows.Edit.Languages.XML;
                    }
                case "JavaScript":
                    {
                        this.LanguageCombo.SelectedIndex = 7;
                        this.PrefLanguageCombo.SelectedIndex = 7;
                        return Syncfusion.Windows.Edit.Languages.XML;
                    }

                default:
                    {
                        return Syncfusion.Windows.Edit.Languages.Text;
                    }
            }
        }


        private void EditorSetup()
        {
            this.RootTabControl.SelectedIndex = 0;
            this.editControl1.DocumentLanguage = LoadPreferredLanguage();

            if (Settings.Default.EditorForeColor != null)
            {
                editControl1.Foreground = Settings.Default.EditorForeColor;
                EditorForeColorPicker.Brush = Settings.Default.EditorForeColor;
            }
            else
            {
                EditorForeColorPicker.Brush = editControl1.Foreground;
            }

            if (Settings.Default.EditorSelectionColor != null)
            {
                editControl1.SelectionForeground = Settings.Default.EditorSelectionColor;
                EditorSelectionColorPicker.Brush = Settings.Default.EditorSelectionColor;
            }
            else
            {
                EditorSelectionColorPicker.Brush = editControl1.SelectionForeground;
            }

            this.selectionBackground = editControl1.SelectionBackground;

            this.IntelliSenseReferences = new ObservableCollection<Uri>();
            this.editControl1.AssemblyReferences = IntelliSenseReferences;
        }
        private void GenerateNewSnippet()
        {
            this.SnippetData = new CodeSnippet();
            this.SnippetData.Language = Settings.Default.PreferredLanguage;
            this.EditorRoot.DataContext = SnippetData;
            editControl1.SetValue(Syncfusion.Windows.Tools.Controls.DockingManager.HeaderProperty, "Untitled");
            this.SnippetData.IsDirty = false;
        }

        private void MyToolWindow_Loaded(object sender, RoutedEventArgs e)
        {
            if (this.SnippetData is null)
            {
                GenerateNewSnippet();

                LoadSnippetLibrary();
                HidePropertiesFromPropertyGrid();

                ResetPkg();

                EditorSetup();

                this.ImportsDataGrid.ItemsSource = SnippetData.Namespaces;
                this.RefDataGrid.ItemsSource = SnippetData.References;
                this.DeclarationsDataGrid.ItemsSource = SnippetData.Declarations;
            }
        }

        private void AddSnippetsButton_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Multiselect = true;
            dlg.Filter = "Snippet files (*.snippet)|*.snippet|All files|*.*";

            dlg.Title = "Select code snippets";
            if (dlg.ShowDialog() == true)
            {

                foreach (var item in dlg.FileNames)
                {
                    var sninfo = new SnippetInfo();
                    sninfo.SnippetFileName = Path.GetFileName(item);
                    sninfo.SnippetPath = Path.GetDirectoryName(item);
                    sninfo.SnippetLanguage = SnippetInfo.GetSnippetLanguage(item);
                    sninfo.SnippetDescription = SnippetInfo.GetSnippetDescription(item);
                    this.VsixData.CodeSnippets.Add(sninfo);
                }
                this.VSVsixTabControl.SelectedIndex = 1;
            }
            else
            {
                return;
            }
        }

        private void DisableDataGrids()
        {
            this.ImportsDataGrid.IsEnabled = false;
            this.RefDataGrid.IsEnabled = false;
        }

        private void EnableDataGrids()
        {
            this.ImportsDataGrid.IsEnabled = true;
            this.RefDataGrid.IsEnabled = true;
        }

        private void SetCurrentLanguage(string snippetLanguage)
        {
            switch (snippetLanguage.ToUpper() ?? "")
            {
                case "VB":
                    {
                        this.LanguageCombo.SelectedIndex = 0;
                        break;
                    }
                case "CSHARP":
                    {
                        this.LanguageCombo.SelectedIndex = 1;
                        break;
                    }
                case "SQL":
                    {
                        this.LanguageCombo.SelectedIndex = 2;
                        break;
                    }
                case "XML":
                    {
                        this.LanguageCombo.SelectedIndex = 3;
                        break;
                    }
                case "XAML":
                    {
                        this.LanguageCombo.SelectedIndex = 4;
                        break;
                    }
                case "CPP":
                    {
                        this.LanguageCombo.SelectedIndex = 5;
                        break;
                    }
                case "HTML":
                    {
                        this.LanguageCombo.SelectedIndex = 6;
                        break;
                    }
                case "JAVASCRIPT":
                    {
                        this.LanguageCombo.SelectedIndex = 7;
                        break;
                    }

                default:
                    {
                        this.LanguageCombo.SelectedIndex = 7;
                        break;
                    }
            }
        }


        private void ImportSumblimeButton_Click(object sender, RoutedEventArgs e)
        {
            if (SnippetData.IsDirty)
            {
                var result = MessageBox.Show("The current snippet has unsaved changes. Are you sure?", "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.No)
                {
                    return;
                }
            }

            var dlg = new OpenFileDialog();

            dlg.Title = "Select code  file";
            dlg.Filter = "All files|*.*";
            if (!(dlg.ShowDialog() == true))
            {
                return;
            }

            try
            {
                var tempData = CodeSnippet.ImportSublimeSnippet(dlg.FileName);
                if (tempData != null)
                {
                    this.SnippetData = null;
                    this.SnippetData = tempData;
                    this.EditorRoot.DataContext = this.SnippetData;
                    this.SnippetData.IsDirty = false;
                    editControl1.SetValue(Syncfusion.Windows.Tools.Controls.DockingManager.HeaderProperty, "Untitled");
                    SetCurrentLanguage(SnippetData.Language);
                }
            }
            catch (UriFormatException ex)
            {
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        private void AddDecButton_Click(object sender, RoutedEventArgs e)
        {
            if (editControl1.SelectedText == "")
                return;

            if (editControl1.SelectedText.ToLower() == "end" | editControl1.SelectedText.ToLower() == "select")
            {
                MessageBox.Show("Declarations are not supported for Select and End words.", "Code Snippet Studio", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            DockingManager.ActivateWindow("DeclarationsDataGrid");
            var newDecl = new Declaration();
            newDecl.Default = editControl1.SelectedText;

            newDecl.ID = editControl1.SelectedText;
            newDecl.ToolTip = "Replace with yours....";

            var query = from decl in SnippetData.Declarations
                        where decl.Default == newDecl.Default
                        select decl;

            if (query.Any())
            {
                MessageBox.Show("A declaration already exists for the specified word", "Code Snippet Studio", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            SnippetData.Declarations.Add(newDecl);
        }

        private void DeleteDecButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.DeclarationsDataGrid.SelectedItem is null)
            {
                return;
            }
            this.DeclarationsDataGrid.View.Remove(this.DeclarationsDataGrid.SelectedItem);
        }

        private void OpenVsixButton_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog();

            dlg.Title = "Select .vsix file";
            dlg.Filter = "Visual Studio Extension (*.vsix)|*.vsix|All files|*.*";
            if (!(dlg.ShowDialog() == true))
            {
                return;
            }
            VsixData = VsixPackage.OpenVsix(dlg.FileName);
            this.VsixGrid.DataContext = VsixData;
        }

        private void LoadCodeFileButton_Click(object sender, RoutedEventArgs e)
        {
            if (SnippetData.IsDirty)
            {
                var result = MessageBox.Show("The current snippet has unsaved changes. Are you sure?", "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.No)
                {
                    return;
                }
            }

            var dlg = new OpenFileDialog();

            dlg.Title = "Select code snippet file";
            dlg.Filter = "Snippet files (*.snippet)|*.snippet;*.vbsnippet;*.vssnippet|Json snippets for VS Code|*.json|All files|*.*";
            if (!(dlg.ShowDialog() == true))
            {
                return;
            }
            try
            {
                var tempData = CodeSnippet.LoadSnippet(dlg.FileName);
                if (tempData != null)
                {
                    this.SnippetData = null;
                    this.SnippetData = tempData;
                    this.EditorRoot.DataContext = this.SnippetData;
                    this.SnippetData.IsDirty = false;
                    editControl1.SetValue(Syncfusion.Windows.Tools.Controls.DockingManager.HeaderProperty, dlg.FileName);
                    if (!(Path.GetExtension(dlg.FileName).ToLower() == "json"))
                        SetCurrentLanguage(SnippetData.Language);
                }
            }
            catch (JsonReaderException)
            {
                MessageBox.Show("The .json snippet file is invalid", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        private void ImportVsiButton_Click(object sender, RoutedEventArgs e)
        {
            var inputFileDialog = new OpenFileDialog();

            inputFileDialog.Title = "Select .vsi file";
            inputFileDialog.Filter = "Visual Studio Content Installer (*.vsi)|*.vsi|All files|*.*";
            if (!(inputFileDialog.ShowDialog() == true))
            {
                return;
            }
            VsixData = VsiService.Vsi2Vsix(inputFileDialog.FileName);
            this.VsixGrid.DataContext = VsixData;
        }

        private void HelpButton_Click(object sender, RoutedEventArgs e)
        {
            Process.Start("https://github.com/AlessandroDelSole/CodeSnippetStudio2022/blob/master/CodeSnippetStudio/CodeSnippetStudio/Code_Snippet_Studio_User_Guide.pdf");
        }

        private void SaveSettings()
        {
            Settings.Default.PreferredLanguage = PrefLanguageCombo.SelectedItem.ToString();
        }

        private void PrefLanguageCombo_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            switch (cb.SelectedIndex)
            {
                case 0:
                    {
                        Settings.Default.PreferredLanguage = "VB";
                        Settings.Default.Save();
                        break;
                    }
                case 1:
                    {
                        Settings.Default.PreferredLanguage = "CSharp";
                        Settings.Default.Save();
                        break;
                    }
                case 2:
                    {
                        Settings.Default.PreferredLanguage = "SQL";
                        Settings.Default.Save();
                        break;
                    }
                case 3:
                    {
                        Settings.Default.PreferredLanguage = "XML";
                        Settings.Default.Save();
                        break;
                    }
                case 4:
                    {
                        Settings.Default.PreferredLanguage = "XAML";
                        Settings.Default.Save();
                        break;
                    }
                case 5:
                    {
                        Settings.Default.PreferredLanguage = "CPP";
                        Settings.Default.Save();
                        break;
                    }
                case 6:
                    {
                        Settings.Default.PreferredLanguage = "HTML";
                        Settings.Default.Save();
                        break;
                    }
                case 7:
                    {
                        Settings.Default.PreferredLanguage = "JavaScript";
                        Settings.Default.Save();
                        break;
                    }
            }
        }
        private void NewSnippetButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.SnippetData.IsDirty)
            {
                var result = MessageBox.Show("There are unsaved changes. Are you sure?", "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (result == MessageBoxResult.No)
                    return;
            }

            this.SnippetData = null;
            GenerateNewSnippet();
        }

        private void FontSizeTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            var result = default(double);
            if (double.TryParse(this.FontSizeTextBox.Text, out result) == true)
            {
                Settings.Default.EditorFontSize = result;
                Settings.Default.Save();
                this.editControl1.FontSize = result;
            }
            else
            {
                MessageBox.Show("Invalid value", "", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
        }

        private void AddRefButton_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog();
            dlg.Title = "Select .NET assembly";
            dlg.Filter = ".dll files (*.dll)|*.dll|All files|*.*";
            dlg.Multiselect = true;
            if (dlg.ShowDialog() == true)
            {
                foreach (var fname in dlg.FileNames)
                {
                    this.IntelliSenseReferences.Add(new Uri(fname));
                    var reference = new Reference();
                    reference.Assembly = Path.GetFileName(fname);
                    SnippetData.References.Add(reference);
                }
            }
        }

        private void DeleteRefButton_Click(object sender, RoutedEventArgs e)
        {
            Reference reference = RefDataGrid.SelectedItem as Reference;
            if (reference != null)
            {
                SnippetData.References.Remove(reference);
            }
        }

        private void LibraryTreeview_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            CodeSnippet item = LibraryTreeview.SelectedItem as CodeSnippet;
            if (item != null)
            {
                this.SnippetData = null;
                this.SnippetData = item;
                this.EditorRoot.DataContext = SnippetData;
                editControl1.SetValue(Syncfusion.Windows.Tools.Controls.DockingManager.HeaderProperty, SnippetData.FileName);
                if (!(Path.GetExtension(SnippetData.FileName).ToLower() == "json"))
                    SetCurrentLanguage(SnippetData.Language);
            }
        }

        private void AddLibFolderButton_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new System.Windows.Forms.FolderBrowserDialog();
            dlg.Description = "New library folder";
            dlg.ShowNewFolderButton = true;

            if (!(dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK))
            {
                return;
            }

            var query = from fold in snippetLib.Folders
                        where fold.FolderName.ToLower() == dlg.SelectedPath.ToLower()
                        select fold;

            if (query.Any())
            {
                // already exist
                MessageBox.Show("Folder already exist in the library", "Not allowed", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var newFolder = new SnippetFolder(dlg.SelectedPath, null);
            snippetLib.Folders.Add(newFolder);
            snippetLib.SaveLibrary(Settings.Default.LibraryName);
        }

        private void DeleteLibFolderButton_Click(object sender, RoutedEventArgs e)
        {
            SnippetFolder item = LibraryTreeview.SelectedItem as SnippetFolder;
            if (item != null)
            {
                snippetLib.Folders.Remove(item);
                snippetLib.SaveLibrary(Settings.Default.LibraryName);
            }
        }

        private void FilterLibraryTextBox_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                FilterSnippetList(this.FilterLibraryTextBox.Text);
            }
        }

        private void FilterSnippetList(string criteria)
        {
            if (string.IsNullOrEmpty(criteria))
            {
                this.LibraryTreeview.ItemsSource = null;
                this.LibraryTreeview.ItemsSource = this.snippetLib.Folders;
                return;
            }

            try
            {
                //var query = this.snippetLib.Folders.Where(f => f.SnippetFiles.Any(s => s.FileName != null && s.FileName.ToLowerInvariant().Contains(criteria.ToLowerInvariant())));
                var tempFolderList = new ObservableCollection<SnippetFolder>();
                foreach(var folder in snippetLib.Folders)
                {
                    var newTempFolder = new SnippetFolder(folder.FolderName, criteria);
                    tempFolderList.Add(newTempFolder);
                }
                this.LibraryTreeview.ItemsSource = tempFolderList;
            }
            catch (Exception)
            {
                this.LibraryTreeview.ItemsSource = this.snippetLib.Folders;
            }

        }

        private void FilterButton_Click(object sender, RoutedEventArgs e)
        {
            FilterSnippetList(this.FilterLibraryTextBox.Text);
        }

        private void BackupLibButton_Click(object sender, RoutedEventArgs e)
        {
            if (snippetLib.Folders.Any())
            {
                var dlg = new SaveFileDialog();
                dlg.Title = "Specify a zip archive name";
                dlg.Filter = "Zip archives|*.zip|All files|*.*";
                dlg.OverwritePrompt = true;

                if (dlg.ShowDialog() == true)
                {
                    try
                    {
                        snippetLib.BackupLibraryToZip(dlg.FileName);
                        MessageBox.Show($"{dlg.FileName} created successfully.");
                        return;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
        }
        private void AddFromLibButton_Click(object sender, RoutedEventArgs e)
        {
            VsixData?.PopulateFromSnippetLibrary(snippetLib);
        }

        private void AnalyzeCode()
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(SnippetData.Code))
                    SnippetData.AnalyzeCode();
                this.ErrorList.ItemsSource = this.SnippetData.Diagnostics;
            }
            catch
            {

            }
        }

        private void ErrorList_CurrentCellRequestNavigate(object sender, CurrentCellRequestNavigateEventArgs args)
        {
            Diagnostic diag = (Diagnostic)args.RowData;
            if (diag.Descriptor.HelpLinkUri != "")
            {
                Process.Start(diag.Descriptor.HelpLinkUri);
            }
        }

        private void ImportCodeFileButton_Click(object sender, RoutedEventArgs e)
        {
            if (SnippetData.IsDirty)
            {
                var result = MessageBox.Show("The current snippet has unsaved changes. Are you sure?", "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.No)
                {
                    return;
                }
            }

            var dlg = new OpenFileDialog();

            dlg.Title = "Select code  file";
            dlg.Filter = "Supported code files (.vb,.cs,.cpp,.js,.sql,.xml,.xaml)|*.cs;*.vb;*.cpp;*.sql;*.js;*.xml;*.xaml|All files|*.*";
            if (!(dlg.ShowDialog() == true))
            {
                return;
            }

            try
            {
                var tempData = CodeSnippet.ImportCodeFile(dlg.FileName);
                if (tempData != null)
                {
                    this.SnippetData = null;
                    this.SnippetData = tempData;
                    this.EditorRoot.DataContext = this.SnippetData;
                    this.SnippetData.IsDirty = false;
                    editControl1.SetValue(Syncfusion.Windows.Tools.Controls.DockingManager.HeaderProperty, "Untitled");
                    SetCurrentLanguage(SnippetData.Language);
                }
            }
            catch (UriFormatException ex)
            {
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        private void ErrorList_SelectionChanged(object sender, GridSelectionChangedEventArgs e)
        {
            Diagnostic diag = ErrorList.SelectedItem as Diagnostic;

            if (diag is null)
                return;

            var span = diag.Location.GetLineSpan();

            if (diag.DefaultSeverity == DiagnosticSeverity.Error)
            {
                editControl1.SelectionBackground = new SolidColorBrush(Colors.Red);
            }
            else if (diag.DefaultSeverity == DiagnosticSeverity.Warning)
            {
                editControl1.SelectionBackground = new SolidColorBrush(Colors.Yellow);
            }
            else
            {
                editControl1.SelectionBackground = this.selectionBackground;
            }

            editControl1.SelectLines(span.Span.Start.Line, span.Span.End.Line, span.Span.Start.Character, span.Span.End.Character);
        }


        private void editControl1_GotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            this.editControl1.SelectionBackground = this.selectionBackground;
        }

        private void EditorForeColorPicker_SelectedBrushChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            this.editControl1.Foreground = EditorForeColorPicker.Brush;
            Settings.Default.EditorForeColor = (SolidColorBrush)EditorForeColorPicker.Brush;
            Settings.Default.Save();
        }

        private void EditorSelectionColorPicker_SelectedBrushChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            this.editControl1.SelectionForeground = EditorSelectionColorPicker.Brush;
            Settings.Default.EditorSelectionColor = (SolidColorBrush)EditorSelectionColorPicker.Brush;
            Settings.Default.Save();
        }

        private void BuildVsixButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (this.VsixData.HasErrors)
            {
                System.Windows.MessageBox.Show("The metadata information is incomplete. Fix errors before compiling.", "Code Snippet Studio", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                return;
            }

            if (VsixData.CodeSnippets.Any() == false)
            {
                MessageBox.Show("The code snippet list is empty. Please add at least one before proceding.", "Code Snippet Studio", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                return;
            }

            var testLang = this.VsixData.TestLanguageGroup();

            if (testLang == false)
            {
                System.Windows.MessageBox.Show("You have added code snippets of different programming languages. " + Environment.NewLine + "VSIX packages offer the best customer experience possible with snippets of only one language." + "For this reason, leave snippets of only one language and remove others before building the package.", "Code Snippet Studio", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }

            try
            {
                var dlg = new SaveFileDialog();
                dlg.Title = "Specify the .vsix name and location";
                dlg.OverwritePrompt = true;
                dlg.Filter = "VSIX packages|*.vsix";
                if (dlg.ShowDialog() == true)
                {
                    this.VsixData.Build(dlg.FileName);
                    var result = MessageBox.Show("Package " + Path.GetFileName(dlg.FileName) + " created. Would you like to install the package for testing now?", "Code Snippet Studio", System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Question);
                    if (result == System.Windows.MessageBoxResult.No)
                    {
                        return;
                    }
                    else
                    {
                        Process.Start(dlg.FileName);
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message);
            }
        }

        private void GuideButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog();
            dlg.Title = "Select documentation file";
            dlg.Filter = "All supported files (*.doc, *.docx, *.rtf, *.txt, *.htm, *.html)|*.doc;*.docx;*.rtf;*.htm;*.html;*.txt|All files|*.*";
            if (dlg.ShowDialog() == true)
            {
                this.VsixData.GettingStartedGuide = dlg.FileName;
            }
        }

        private void IconButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog();
            dlg.Title = "Select icon file";
            dlg.Filter = "All supported files (*.jpg, *.png, *.ico, *.bmp, *.tif, *.tiff, *.gif)|*.jpg;*.png;*.ico;*.bmp;*.tiff;*.tif;*.gif|All files|*.*";
            if (dlg.ShowDialog() == true)
            {
                this.VsixData.IconPath = dlg.FileName;
            }
        }

        private void ImageButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog();
            dlg.Title = "Select preview image";
            dlg.Filter = "All supported files (*.jpg, *.png, *.ico, *.bmp, *.tif, *.tiff, *.gif)|*.jpg;*.png;*.ico;*.bmp;*.tiff;*.tif;*.gif|All files|*.*";
            if (dlg.ShowDialog() == true)
            {
                this.VsixData.PreviewImagePath = dlg.FileName;
            }
        }

        private void LicenseButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog();
            dlg.Title = "Select documentation file";
            dlg.Filter = "All supported files (*.rtf, *.txt)|*.rtf;*.txt|All files|*.*";
            if (dlg.ShowDialog() == true)
            {
                this.VsixData.License = dlg.FileName;
            }
        }

        private void RelNotesButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog();
            dlg.Title = "Select documentation file";
            dlg.Filter = "All supported files (*.rtf, *.txt, *.htm, *.html)|*.rtf;*.htm;*.html;*.txt|All files|*.*";
            if (dlg.ShowDialog() == true)
            {
                this.VsixData.ReleaseNotes = dlg.FileName;
            }
        }

        private void RemoveSnippetButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            try
            {
                var selectedList = this.CodeSnippetsDataGrid.SelectedItems.Cast<SnippetInfo>().ToList();
                if (!selectedList.Any())
                    return;

                this.VSVsixTabControl.SelectedIndex = 1;

                var result = MessageBox.Show("Are you sure you want to remove the selected snippet(s)?", "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.No)
                    return;

                foreach (var item in selectedList)
                    this.VsixData.CodeSnippets.Remove(item as SnippetInfo);
            }
            catch
            {

            }
        }

        private void ResetPkgButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            var result = MessageBox.Show("Are you sure?", "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result == MessageBoxResult.Yes)
                ResetPkg();
        }

        private void AboutButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Process.Start("https://github.com/AlessandroDelSole/CodeSnippetStudio2022");
        }

        private void VsiButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (VsixData.HasErrors)
            {
                MessageBox.Show("The package metadata contain errors that must be fixed before performing a conversion." + Environment.NewLine + "Please go to the Package and Share tab and check values in the Metadata nested tab.");
                return;
            }

            var inputFileDialog = new OpenFileDialog();
            string inputFile;
            string outputFile;

            inputFileDialog.Title = "Select .vsi file";
            inputFileDialog.Filter = "Visual Studio Community Installer (*.vsi)|*.vsi|All files|*.*";
            if (!(inputFileDialog.ShowDialog() == true))
            {
                return;
            }
            inputFile = inputFileDialog.FileName;

            var outputFileDialog = new SaveFileDialog();
            outputFileDialog.OverwritePrompt = true;
            outputFileDialog.Title = "Output .vsix file";
            outputFileDialog.Filter = ".vsix files (*.vsix)|*.vsix|All files|*.*";
            if (!(outputFileDialog.ShowDialog() == true))
            {
                return;
            }
            outputFile = outputFileDialog.FileName;

            VsiService.Vsi2Vsix(inputFile, outputFile, VsixData.SnippetFolderName, VsixData.PackageAuthor, VsixData.ProductName, VsixData.PackageDescription, VsixData.IconPath, VsixData.PreviewImagePath, VsixData.MoreInfoURL);
            MessageBox.Show($"Successfully converted {inputFile} into {outputFile}");
        }

        private void SignButton_Click(object sender, RoutedEventArgs e)
        {
            if (PfxPassword.Password.Length == 0)
            {
                MessageBox.Show("Please enter the password for the certificate file.");
                return;
            }

            if (PfxTextBox.Text == "" | string.IsNullOrEmpty(PfxTextBox.Text))
            {
                MessageBox.Show("Please specify a valid X.509 certificate file.");
                return;
            }

            if (!File.Exists(PfxTextBox.Text))
            {
                MessageBox.Show("The specified certificate file does not exist.");
                return;
            }

            var inputFileDialog = new OpenFileDialog();
            inputFileDialog.Title = "Select the .vsix you want to sign";
            inputFileDialog.Filter = "VSIX packages (*.vsix)|*.vsix|All files|*.*";
            if (inputFileDialog.ShowDialog() == true)
            {
                VsixPackage.SignVsix(inputFileDialog.FileName, PfxTextBox.Text, PfxPassword.Password);
                MessageBox.Show($"{inputFileDialog.FileName} signed successfully.");
            }
        }

        private void PfxButton_Click(object sender, RoutedEventArgs e)
        {
            var inputFileDialog = new OpenFileDialog();
            inputFileDialog.Title = "Select certificate file";
            inputFileDialog.Filter = "Certificate files (*.pfx)|*.pfx|All files|*.*";
            if (inputFileDialog.ShowDialog() == true)
            {
                this.PfxTextBox.Text = inputFileDialog.FileName;
            }
        }

        private void LanguageCombo_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            switch (cb.SelectedIndex)
            {
                case 0:
                    {
                        this.editControl1.DocumentLanguage = Syncfusion.Windows.Edit.Languages.VisualBasic;
                        this.SnippetData.Language = "VB";
                        EnableDataGrids();
                        AnalyzeCode();
                        break;
                    }
                case 1:
                    {
                        this.editControl1.DocumentLanguage = Syncfusion.Windows.Edit.Languages.CSharp;
                        this.SnippetData.Language = "CSharp";
                        //DisableDataGrids();
                        AnalyzeCode();
                        break;
                    }
                case 2:
                    {
                        this.editControl1.DocumentLanguage = Syncfusion.Windows.Edit.Languages.SQL;
                        this.SnippetData.Language = "SQL";
                        DisableDataGrids();
                        break;
                    }
                case 3:
                    {
                        this.editControl1.DocumentLanguage = Syncfusion.Windows.Edit.Languages.XML;
                        this.SnippetData.Language = "XML";
                        DisableDataGrids();
                        break;
                    }
                case 4:
                    {
                        this.editControl1.DocumentLanguage = Syncfusion.Windows.Edit.Languages.XAML;
                        this.SnippetData.Language = "XAML";
                        DisableDataGrids();
                        break;
                    }
                case 5:
                    {
                        this.editControl1.DocumentLanguage = Syncfusion.Windows.Edit.Languages.Text;
                        this.SnippetData.Language = "CPP";
                        DisableDataGrids();
                        break;
                    }
                case 6:
                    {
                        this.editControl1.DocumentLanguage = Syncfusion.Windows.Edit.Languages.XML;
                        this.SnippetData.Language = "HTML";
                        DisableDataGrids();
                        break;
                    }
                case 7:
                    {
                        this.editControl1.DocumentLanguage = Syncfusion.Windows.Edit.Languages.XML;
                        this.SnippetData.Language = "JavaScript";
                        DisableDataGrids();
                        break;
                    }
            }
        }

        private void SaveSnippetButton_Click(object sender, RoutedEventArgs e)
        {
            if (SnippetData.HasErrors)
            {
                MessageBox.Show("The current code snippet has errors that must be fixed before saving." + Environment.NewLine + "Ensure that Author, Title, Description, and snippet language have been supplied properly.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var dlg2 = new SaveFileDialog();
            dlg2.OverwritePrompt = true;
            dlg2.Title = "Output .snippet file";
            dlg2.Filter = ".snippet files (*.snippet)|*.snippet|All files|*.*";
            if (!(dlg2.ShowDialog() == true))
            {
                return;
            }

            if (SnippetData.Language == "" | string.IsNullOrEmpty(SnippetData.Language))
            {
                SnippetData.Language = Settings.Default.PreferredLanguage;
            }

            SnippetData.SaveSnippet(dlg2.FileName);
            editControl1.SetValue(Syncfusion.Windows.Tools.Controls.DockingManager.HeaderProperty, dlg2.FileName);
            MessageBox.Show($"{dlg2.FileName} saved correctly.");
        }

        private void ExtractButton_Click(object sender, RoutedEventArgs e)
        {
            var inputFileDialog = new OpenFileDialog();
            string inputFile;
            string outputFolder;

            inputFileDialog.Title = "Select .vsix file";
            inputFileDialog.Filter = "Visual Studio Extension (*.vsix)|*.vsix|All files|*.*";
            if (!(inputFileDialog.ShowDialog() == true))
            {
                return;
            }
            inputFile = inputFileDialog.FileName;

            var outputFolderDialog = new System.Windows.Forms.FolderBrowserDialog();
            outputFolderDialog.Description = "Select destination folder";
            outputFolderDialog.ShowNewFolderButton = true;

            if (!(outputFolderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK))
            {
                return;
            }

            outputFolder = outputFolderDialog.SelectedPath;

            VsixPackage.ExtractVsix(inputFile, outputFolder, OnlySnippetsCheckBox.IsChecked.Value);
            MessageBox.Show($"Successfully extracted {inputFile} into {outputFolder}");
        }

        private void SaveCodeFileButton_Click(object sender, RoutedEventArgs e)
        {
            if (editControl1.Text == "")
            {
                MessageBox.Show("Write some code first!");
                return;
            }

            string filter = "All files|*.*";
            switch (LanguageCombo.SelectedIndex)
            {
                case 0:
                    {
                        filter = "Visual Basic code file (.vb)|*.vb|All files|*.*";
                        break;
                    }
                case 1:
                    {
                        filter = "C# code file (.cs)|*.cs|All files|*.*";
                        break;
                    }
                case 2:
                    {
                        filter = "SQL code file (.sql)|*.sql|All files|*.*";
                        break;
                    }
                case 3:
                    {
                        filter = "XML file (.xml)|*.xml|All files|*.*";
                        break;
                    }
                case 4:
                    {
                        filter = "XAML file (.xaml)|*.xaml|All files|*.*";
                        break;
                    }
                case 5:
                    {
                        filter = "C++ code file (.cpp)|*.cpp|All files|*.*";
                        break;
                    }
                case 6:
                    {
                        filter = "HTML file (.htm)|*.htm|All files|*.*";
                        break;
                    }
                case 7:
                    {
                        filter = "JavaScript code file (.js)|*.js|All files|*.*";
                        break;
                    }
            }

            var dlg = new SaveFileDialog();
            dlg.Title = "Specify code file name";
            dlg.Filter = filter;
            dlg.OverwritePrompt = true;

            if (dlg.ShowDialog() == true)
            {
                File.WriteAllText(dlg.FileName, editControl1.Text);
            }
        }

        private void editControl1_TextChanged_1(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            SnippetData.Code = editControl1.Text;
            AnalyzeCode();
        }
    }

}