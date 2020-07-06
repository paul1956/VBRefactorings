' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Globalization
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Testing
Imports VBRefactorings
Imports Xunit
Imports Verify = Microsoft.CodeAnalysis.VisualBasic.Testing.MSTest.CodeFixVerifier(Of VBRefactorings.Style.ByValAnalyzerFixAnalyzer, VBRefactorings.Style.ByValAnalyzerFixCodeFixProvider)

Namespace ByValAnalyzerFix.Test

    <TestClass>
    Public Class ByValAnalyzerFixUnitTest

        'Diagnostic And CodeFix both triggered And checked for
        <Fact>
        Public Async Sub TestByValNoExtraTrivia()

            Dim OriginalSource As String = "
Module Module1

    Sub Main(X As String, ByVal Y As Integer)

    End Sub

End Module"
            Dim expected As DiagnosticResult = Verify.Diagnostic(RemoveByValDiagnosticId) _
                                .WithLocation("/0/Test0.vb", 4, 27).
                                WithMessage(String.Format(CultureInfo.InvariantCulture, "Remove ByVal from parameter {0}", "1")).
                                WithSeverity(DiagnosticSeverity.Info)

            Dim FixedSource As String = "
Module Module1

    Sub Main(X As String, Y As Integer)

    End Sub

End Module"
            Await Verify.VerifyCodeFixAsync(OriginalSource, expected, FixedSource).ConfigureAwait(True)
        End Sub

        'Diagnostic And CodeFix both triggered And checked for
        <Fact>
        Public Async Sub TestByValWithLeadingTrivia()

            Dim OriginalSource As String = "
Module Module1

    Sub Main(X As String, ' Comment
             ByVal Y As Integer)

    End Sub

End Module"

            Dim expected As DiagnosticResult = Verify.Diagnostic(RemoveByValDiagnosticId) _
                                .WithLocation("/0/Test0.vb", 5, 14).
                                WithMessage(String.Format(CultureInfo.InvariantCulture, "Remove ByVal from parameter {0}", "1")).
                                WithSeverity(DiagnosticSeverity.Info)

            Dim FixedSource As String = "
Module Module1

    Sub Main(X As String, ' Comment
             Y As Integer)

    End Sub

End Module"
            Await Verify.VerifyCodeFixAsync(OriginalSource, expected, FixedSource).ConfigureAwait(True)
        End Sub

        <Fact>
        Public Async Sub TestByValWithTrailingMultiLineTrivia()

            Dim OriginalSource As String = "
Module Module1

    Sub Main(X As String, ByVal _ ' Comment
 _ ' Another Comment
             Y As Integer)

    End Sub

End Module"

            Dim expected As DiagnosticResult = Verify.Diagnostic(RemoveByValDiagnosticId) _
                                .WithLocation("/0/Test0.vb", 4, 27).
                                WithMessage(String.Format(CultureInfo.InvariantCulture, "Remove ByVal from parameter {0}", "1")).
                                WithSeverity(DiagnosticSeverity.Info)

            Dim FixedSource As String = "
Module Module1

    Sub Main(X As String, _ ' Comment
 _ ' Another Comment
Y As Integer)

    End Sub

End Module"
            Await Verify.VerifyCodeFixAsync(OriginalSource, expected, FixedSource).ConfigureAwait(True)
        End Sub

        <Fact>
        Public Async Sub TestByValWithTrailingTrivia()

            Dim OriginalSource As String = "
Module Module1

    Sub Main(X As String, ByVal _ ' Comment
             Y As Integer)

    End Sub

End Module"

            Dim expected As DiagnosticResult = Verify.Diagnostic(RemoveByValDiagnosticId) _
                                .WithLocation("/0/Test0.vb", 4, 27).
                                WithMessage(String.Format(CultureInfo.InvariantCulture, "Remove ByVal from parameter {0}", "1")).
                                WithSeverity(DiagnosticSeverity.Info)

            Dim FixedSource As String = "
Module Module1

    Sub Main(X As String, _ ' Comment
Y As Integer)

    End Sub

End Module"
            Await Verify.VerifyCodeFixAsync(OriginalSource, expected, FixedSource).ConfigureAwait(True)
        End Sub

        'No diagnostics Expected to show up
        <Fact>
        Public Async Sub TestNoByVal()
            Dim OriginalSource As String = ""
            Await Verify.VerifyCodeFixAsync(OriginalSource, OriginalSource).ConfigureAwait(True)
        End Sub

    End Class

End Namespace
