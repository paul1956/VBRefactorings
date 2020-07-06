' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Testing
Imports VBRefactorings
Imports Xunit

Imports Verify = Microsoft.CodeAnalysis.VisualBasic.Testing.MSTest.CodeFixVerifier(Of VBRefactorings.Style.AddAsClauseForLambdasAnalyzer, VBRefactorings.Style.AddAsClauseCodeFixProvider)

Namespace AddAsClauseForLambda.UnitTest

    <TestClass()> Public Class AddAsClauseForLambdaUnitTests

        <Fact>
        Public Async Sub AddAsClauseToMultiLineLambdaFunctionSplit()
            Const OriginalSource As String =
"Class Class1
    Sub Main()
        Dim getSortColumn As System.Func(Of Integer, String)
        getSortColumn = Function(index)
                            Select Case index
                                Case 0
                                    Return ""FirstName""
                                Case Else
                                    Return ""LastName""
                            End Select
                        End Function
    End Sub
End Class"

            Dim expected As DiagnosticResult = Verify.Diagnostic(AddAsClauseForLambdaDiagnosticId) _
                                .WithLocation("/0/Test0.vb", 4, 34).
                                WithMessage("Option Strict On requires all Lambda declarations to have an 'As' clause.").
                                WithSeverity(DiagnosticSeverity.Warning)
            Const FixedSource As String =
"Class Class1
    Sub Main()
        Dim getSortColumn As System.Func(Of Integer, String)
        getSortColumn = Function(index As Integer)
                            Select Case index
                                Case 0
                                    Return ""FirstName""
                                Case Else
                                    Return ""LastName""
                            End Select
                        End Function
    End Sub
End Class"
            Await Verify.VerifyCodeFixAsync(OriginalSource, expected, FixedSource).ConfigureAwait(True)
        End Sub

        <Fact>
        Public Async Sub AddAsClauseToMultiLineLambdaSubroutineSplit()
            Const OriginalSource As String =
"Class Class1
    Sub Main()
        Dim writeToLog As System.Action(Of String)
        writeToLog = Sub(msg)
                        Dim l As String = msg
                     End Sub
    End Sub
End Class"

            Dim expected As DiagnosticResult = Verify.Diagnostic(AddAsClauseForLambdaDiagnosticId) _
                                .WithLocation("/0/Test0.vb", 4, 26).
                                WithMessage("Option Strict On requires all Lambda declarations to have an 'As' clause.").
                                WithSeverity(DiagnosticSeverity.Warning)
            Const FixedSource As String =
"Class Class1
    Sub Main()
        Dim writeToLog As System.Action(Of String)
        writeToLog = Sub(msg As String)
                         Dim l As String = msg
                     End Sub
    End Sub
End Class"
            Await Verify.VerifyCodeFixAsync(OriginalSource, expected, FixedSource).ConfigureAwait(True)
        End Sub

        <Fact>
        Public Async Sub AddAsClauseToSingleLineLambdaFunctionSplit()
            Const OriginalSource As String =
"Class Class1
    Sub Main()
        Dim add1 As System.Func(Of UInteger, Integer, Integer)
        add1 = Function(Num, Num1) Num + Num1
    End Sub
End Class"
            Dim expected1 As DiagnosticResult = Verify.Diagnostic(AddAsClauseForLambdaDiagnosticId) _
                                .WithLocation("/0/Test0.vb", 4, 25).
                                WithMessage("Option Strict On requires all Lambda declarations to have an 'As' clause.").
                                WithSeverity(DiagnosticSeverity.Warning)

            Dim expected2 As DiagnosticResult = Verify.Diagnostic(AddAsClauseForLambdaDiagnosticId) _
                                .WithLocation("/0/Test0.vb", 4, 30).
                                WithMessage("Option Strict On requires all Lambda declarations to have an 'As' clause.").
                                WithSeverity(DiagnosticSeverity.Warning)

            Const FixedSource As String =
"Class Class1
    Sub Main()
        Dim add1 As System.Func(Of UInteger, Integer, Integer)
        add1 = Function(Num As UInteger, Num1 As Integer) Num + Num1
    End Sub
End Class"
            Await Verify.VerifyCodeFixAsync(OriginalSource, {expected1, expected2}, FixedSource).ConfigureAwait(True)
        End Sub

        <Fact>
        Public Async Sub AddAsClauseToSingleLineLambdaSubroutineSplit()
            Const OriginalSource As String =
"Imports System
Class Class1
    Sub Main()
        Dim writeMessage As System.Action(Of String)
        writeMessage = Sub(msg) Line(msg)
    End Sub
    Private Sub Line(msg As String)
        Throw New NotImplementedException()
    End Sub
End Class"

            Dim expected As DiagnosticResult = Verify.Diagnostic(AddAsClauseForLambdaDiagnosticId) _
                                .WithLocation("/0/Test0.vb", 5, 28).
                                WithMessage("Option Strict On requires all Lambda declarations to have an 'As' clause.").
                                WithSeverity(DiagnosticSeverity.Warning)

            Const FixedSource As String =
"Imports System
Class Class1
    Sub Main()
        Dim writeMessage As System.Action(Of String)
        writeMessage = Sub(msg As String) Line(msg)
    End Sub
    Private Sub Line(msg As String)
        Throw New NotImplementedException()
    End Sub
End Class"
            Await Verify.VerifyCodeFixAsync(OriginalSource, expected, FixedSource).ConfigureAwait(True)
        End Sub

    End Class

End Namespace
