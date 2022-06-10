# Code Snippet Studio 2022


![](https://img.shields.io/badge/release-stable-brightgreen.svg)


Code Snippet Studio is an [extension for Visual Studio 2022](https://marketplace.visualstudio.com/items?itemName=AlessandroDelSole.CodeSnippetStudio2022) that makes it easy to create, edit, package, and share IntelliSense code snippets for **Visual Studio 2022**. For C# and Visual Basic snippets, it also provides live [Roslyn](https://github.com/dotnet/roslyn) code analysis as you type to immediately detect code issues.

With Code Snippet Studio, you have a fully-functional code editor with syntax highlighting, basic IntelliSense, and live code issue detection as you type, where you can write or paste your snippets and supply the information that is required by the respective schema references. Studio will generate the proper .snippet files for you. You will then be able to package a number of code snippets into a Visual Studio extension by building .vsix packages that you can share with other developers and that you can even publish to the Visual Studio Gallery! Additional tools are available to work with both .vsix and .vsi installers.


![Code Snippet Studio](https://github.com/AlessandroDelSole/CodeSnippetStudio2022/blob/master/adsCSS2022.gif)

**To show up Code Snippet Studio, in Visual Studio 2022 select View, Other Windows, Code Snippet Studio**

[Getting Started Guide](https://github.com/AlessandroDelSole/CodeSnippetStudio2022/blob/master/CodeSnippetStudio/CodeSnippetStudio/Code_Snippet_Studio_User_Guide.pdf)

## Breaking changes with version 2015:
- Extension themes are no longer available. This allows for a lightweight package that can adapt to the Visual Studio theme
- Support for Visual Studio Code is no longer available. The effort to maintain this feature is much higher than the benefits and real target users
- New iconography
- Filtering the library of snippets has now been fixed
- While the support libraries are still written in Visual Basic, the extension itself is now written in C#. This is to hopefully improve collaboration. 


Code Snippet Studio has been built using controls from [Syncfusion's Essential Studio for WPF (Community license)](https://www.syncfusion.com/products/communitylicense)
