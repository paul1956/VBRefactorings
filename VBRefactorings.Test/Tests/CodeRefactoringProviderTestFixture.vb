' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Threading

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeActions
Imports Microsoft.CodeAnalysis.CodeRefactorings
Imports Microsoft.CodeAnalysis.Text
Imports Xunit

Namespace Roslyn.UnitTestFramework
    Public MustInherit Class CodeRefactoringProviderTestFixture
        Inherits CodeActionProviderTestFixture

        Private Function GetRefactoring(document As Document, span As TextSpan) As IEnumerable(Of CodeAction)
            Dim provider As CodeRefactoringProvider = Me.CreateCodeRefactoringProvider()
            Dim actions As New List(Of CodeAction)()
            Dim context As New CodeRefactoringContext(document, span, Sub(a) actions.Add(a), CancellationToken.None)
            provider.ComputeRefactoringsAsync(context).Wait()
            Return actions
        End Function

        Protected Sub TestNoActions(markup As String)
            If markup Is Nothing Then
                Throw New ArgumentNullException(NameOf(markup))
            End If

            If Not markup.Contains(vbCr) Then
                markup = markup.Replace(vbLf, vbCrLf)
            End If

            Dim code As String = Nothing
            Dim span As TextSpan = Nothing
            MarkupTestFile.GetSpan(markup, code, span)

            Dim document As Document = Me.CreateDocument(code)
            Dim actions As IEnumerable(Of CodeAction) = Me.GetRefactoring(document, span)

            Assert.True(actions Is Nothing OrElse Not actions.Any)
        End Sub

        Protected Sub Test(markup As String, expected As String, Optional actionIndex As Integer = 0, Optional compareTokens As Boolean = False)

            If markup IsNot Nothing AndAlso Not markup.Contains(vbCr) Then
                markup = markup.Replace(vbLf, vbCrLf)
            End If

            If expected Is Nothing Then
                Throw New ArgumentNullException(NameOf(expected))
            End If

            If markup Is Nothing Then
                Throw New ArgumentNullException(NameOf(markup))
            End If

            If Not expected.Contains(vbCr) Then
                expected = expected.Replace(vbLf, vbCrLf)
            End If

            Dim code As String = Nothing
            Dim span As TextSpan = Nothing
            MarkupTestFile.GetSpan(markup, code, span)

            Dim document As Document = Me.CreateDocument(code)
            Dim actions As IEnumerable(Of CodeAction) = Me.GetRefactoring(document, span)

            Assert.NotNull(actions)

            Dim action As CodeAction = actions.ElementAt(actionIndex)
            Assert.NotNull(action)

            Dim edit As ApplyChangesOperation = action.GetOperationsAsync(CancellationToken.None).Result.OfType(Of ApplyChangesOperation)().First()
            Me.VerifyDocument(expected, compareTokens, edit.ChangedSolution.GetDocument(document.Id))
        End Sub

        Protected MustOverride Function CreateCodeRefactoringProvider() As CodeRefactoringProvider
    End Class
End Namespace
