' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Globalization

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Formatting
Imports Microsoft.CodeAnalysis.Simplification
Imports Microsoft.CodeAnalysis.Text
Imports Xunit

Namespace Roslyn.UnitTestFramework
    Public MustInherit Class CodeActionProviderTestFixture
        Protected Function CreateDocument(code As String) As Document
            Dim fileExtension As String = If(LanguageName = LanguageNames.CSharp, ".cs", ".vb")

            Dim projectId As ProjectId = ProjectId.CreateNewId(debugName:="TestProject")
            Dim documentId As DocumentId = DocumentId.CreateNewId(projectId, debugName:="Test" & fileExtension)

            Using NewAdhocWorkspace As AdhocWorkspace = New AdhocWorkspace()
                Return NewAdhocWorkspace.CurrentSolution.AddProject(projectId, "TestProject", "TestProject", LanguageName).AddMetadataReferences(projectId, SharedReferences.VisualBasicReferences("")).AddDocument(documentId, "Test" & fileExtension, SourceText.From(code)).GetDocument(documentId)
            End Using
        End Function

        Protected Sub VerifyDocument(expected As String, compareTokens As Boolean, document As Document)
            If document Is Nothing Then
                Throw New ArgumentNullException(NameOf(document))
            End If

            If compareTokens Then
                Me.VerifyTokens(expected, Format(document).ToString())
            Else
                VerifyText(expected, document)
            End If
        End Sub

        Private Shared Function Format(document As Document) As SyntaxNode
            Dim updatedDocument As Document = document.WithSyntaxRoot(document.GetSyntaxRootAsync().Result)
            Return Formatter.FormatAsync(Simplifier.ReduceAsync(updatedDocument, Simplifier.Annotation).Result, Formatter.Annotation).Result.GetSyntaxRootAsync().Result
        End Function

        Private Function ParseTokens(text As String) As IList(Of SyntaxToken)
#Disable Warning IDE0004 ' Remove Unnecessary Cast
            Return Me.ParseTokens(text).Select(Function(t As SyntaxToken) CType(t, SyntaxToken)).ToList()
#Enable Warning IDE0004 ' Remove Unnecessary Cast
        End Function

        Private Function VerifyTokens(expected As String, actual As String) As Boolean
            Dim expectedNewTokens As IList(Of SyntaxToken) = Me.ParseTokens(expected)
            Dim actualNewTokens As IList(Of SyntaxToken) = Me.ParseTokens(actual)

            For i As Integer = 0 To Math.Min(expectedNewTokens.Count, actualNewTokens.Count) - 1
                Assert.Equal(expectedNewTokens(i).ToString(), actualNewTokens(i).ToString())
            Next i

            If expectedNewTokens.Count <> actualNewTokens.Count Then
                Dim expectedDisplay As String = String.Join(" ", expectedNewTokens.Select(Function(t As SyntaxToken) t.ToString()))
                Dim actualDisplay As String = String.Join(" ", actualNewTokens.Select(Function(t As SyntaxToken) t.ToString()))
                Assert.True(False, String.Format(CultureInfo.InvariantCulture,
                                                 "Wrong token count. Expected '{0}', Actual '{1}', Expected Text: '{2}', Actual Text: '{3}'", expectedNewTokens.Count, actualNewTokens.Count, expectedDisplay, actualDisplay))
            End If

            Return True
        End Function

        Private Shared Function VerifyText(expected As String, document As Document) As Boolean
            Dim actual As String = Format(document).ToString()
            Assert.Equal(expected, actual, ignoreWhiteSpaceDifferences:=True)
            Return True
        End Function

        Protected MustOverride ReadOnly Property LanguageName() As String
    End Class
End Namespace
