' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Testing
Imports VBRefactorings
Imports Xunit
Imports Verify = Microsoft.CodeAnalysis.VisualBasic.Testing.MSTest.CodeFixVerifier(Of VBRefactorings.Style.AddAsClauseAnalyzer, VBRefactorings.Style.AddAsClauseCodeFixProvider)

Namespace AddAsClause.UnitTest

    <TestClass()> Public Class AddAsClauseFailingUnitTests

        <Fact>
        Public Async Sub VerifyFixAddAsClauseInDiagnosticArray()
            Const OriginalSource As String = "Imports System
Imports System.Collections.Generic
Partial Public MustInherit Class DiagnosticVerifier
    Protected Shared Function GetSortedDiagnosticsFromDocuments() As String()
        Dim diagnostics As New List(Of String)
        Dim results = SortDiagnostics(diagnostics)
        Return results
    End Function
    Private Shared Function SortDiagnostics(diagnostics As IEnumerable(Of String)) As String()
        Return Nothing
    End Function
End Class"
            Dim expected As DiagnosticResult = Verify.Diagnostic(AddAsClauseDiagnosticId) _
                                .WithLocation("/0/Test0.vb", 6, 13).
                                WithMessage("Option Strict On requires all variable declarations to have an 'As' clause.").
                                WithSeverity(DiagnosticSeverity.Error)

            Const FixedSource As String = "Imports System
Imports System.Collections.Generic
Partial Public MustInherit Class DiagnosticVerifier
    Protected Shared Function GetSortedDiagnosticsFromDocuments() As String()
        Dim diagnostics As New List(Of String)
        Dim results As String() = SortDiagnostics(diagnostics)
        Return results
    End Function
    Private Shared Function SortDiagnostics(diagnostics As IEnumerable(Of String)) As String()
        Return Nothing
    End Function
End Class"
            Await Verify.VerifyCodeFixAsync(OriginalSource, expected, FixedSource).ConfigureAwait(True)
        End Sub

        <Fact>
        Public Async Sub VerifyFixAddAsClauseInDimStatement()
            Const OriginalSource As String =
"Imports System
Imports AliasedType = NS.C(Of Integer)
Namespace NS
    Public Class C(Of T)
        Public Structure S(Of U)
        End Structure
    End Class
End Namespace
Module Program
    Public Sub Main()
        Dim s = New AliasedType.S(Of Long)
        Console.WriteLine(s.ToString)
    End Sub
End Module"
            Dim expected As DiagnosticResult = Verify.Diagnostic(AddAsClauseDiagnosticId).
                                WithLocation("/0/Test0.vb", 11, 13).
                                WithMessage("Option Strict On requires all variable declarations to have an 'As' clause.").
                                WithSeverity(DiagnosticSeverity.Error)

            Const FixedSource As String =
"Imports System
Imports AliasedType = NS.C(Of Integer)
Namespace NS
    Public Class C(Of T)
        Public Structure S(Of U)
        End Structure
    End Class
End Namespace
Module Program
    Public Sub Main()
        Dim s As New AliasedType.S(Of Long)
        Console.WriteLine(s.ToString)
    End Sub
End Module"
            Await Verify.VerifyCodeFixAsync(OriginalSource, expected, FixedSource).ConfigureAwait(True)
        End Sub

        <Fact>
        Public Async Sub VerifyFixAddAsClauseInDimStatementCollectionInitializer()
            Const OriginalSource As String = "
Module Program
    Public Sub Main()
        Dim s = {""X""c}
    End Sub
End Module"
            Dim expected As DiagnosticResult = Verify.Diagnostic(AddAsClauseDiagnosticId).
                                WithLocation("/0/Test0.vb", 4, 13).
                                WithMessage("Option Strict On requires all variable declarations to have an 'As' clause.").
                                WithSeverity(DiagnosticSeverity.Error)

            Const FixedSource As String = "
Module Program
    Public Sub Main()
        Dim s As Char() = {""X""c}
    End Sub
End Module"
            Await Verify.VerifyCodeFixAsync(OriginalSource, expected, FixedSource).ConfigureAwait(True)
        End Sub

        <Fact>
        Public Async Sub VerifyFixAddAsClauseInFor()
            Dim OriginalSource As String =
<text>Option Explicit Off
Option Infer Off
Option Strict On
Imports System
Class Class1
    Sub Main()
        For I = 0 to 1
            Console.Write(i)
        Next
    End Sub
End Class</text>.Value
            Dim expected As DiagnosticResult = Verify.Diagnostic(AddAsClauseDiagnosticId).
                                WithLocation("/0/Test0.vb", 7, 13).
                                WithMessage("Option Strict On requires all variable declarations to have an 'As' clause.").
                                WithSeverity(DiagnosticSeverity.Error)

            Dim FixedSource As String =
<text>Option Explicit Off
Option Infer Off
Option Strict On
Imports System
Class Class1
    Sub Main()
        For I As Integer = 0 to 1
            Console.Write(i)
        Next
    End Sub
End Class</text>.Value
            Await Verify.VerifyCodeFixAsync(OriginalSource, expected, FixedSource).ConfigureAwait(True)
        End Sub

        <Fact>
        Public Async Sub VerifyFixAddAsClauseInForEachArrayElement()
            Const OriginalSource As String =
"Imports System

Class Class1
    Sub Main()
        Dim Array1(10) As String
        For Each element In Array1
            Console.Write(element)
        Next
    End Sub
End Class"
            Dim expected As DiagnosticResult = Verify.Diagnostic(AddAsClauseDiagnosticId).
                                WithLocation("/0/Test0.vb", 6, 18).
                                WithMessage("Option Strict On requires all variable declarations to have an 'As' clause.").
                                WithSeverity(DiagnosticSeverity.Error)
            Const FixedSource As String =
"Imports System

Class Class1
    Sub Main()
        Dim Array1(10) As String
        For Each element As String In Array1
            Console.Write(element)
        Next
    End Sub
End Class"
            Await Verify.VerifyCodeFixAsync(OriginalSource, expected, FixedSource).ConfigureAwait(True)
        End Sub

        <Fact>
        Public Async Sub VerifyFixAddAsClauseInForEachArrayElement1()
            Const OriginalSource As String =
"Module Module1
    Sub M(s As String)
        For Each c In s
        Next
    End Sub
End Module"
            Dim expected As DiagnosticResult = Verify.Diagnostic(AddAsClauseDiagnosticId).
                                WithLocation("/0/Test0.vb", 3, 18).
                                WithMessage("Option Strict On requires all variable declarations to have an 'As' clause.").
                                WithSeverity(DiagnosticSeverity.Error)
            Const FixedSource As String =
"Module Module1
    Sub M(s As String)
        For Each c As Char In s
        Next
    End Sub
End Module"
            Await Verify.VerifyCodeFixAsync(OriginalSource, expected, FixedSource).ConfigureAwait(True)
        End Sub

        <Fact>
        Public Async Sub VerifyFixAddAsClauseInForEachByte()
            Const OriginalSource As String =
"Imports System
Class Class1
    Sub Main()
        For value = Byte.MinValue To Byte.MaxValue
           Console.Write($""{value}"")
        Next
    End Sub
End Class"
            Dim expected As DiagnosticResult = Verify.Diagnostic(AddAsClauseDiagnosticId).
                                WithLocation("/0/Test0.vb", 4, 13).
                                WithMessage("Option Strict On requires all variable declarations to have an 'As' clause.").
                                WithSeverity(DiagnosticSeverity.Error)
            Const FixedSource As String =
"Imports System
Class Class1
    Sub Main()
        For value As Byte = Byte.MinValue To Byte.MaxValue
           Console.Write($""{value}"")
        Next
    End Sub
End Class"
            Await Verify.VerifyCodeFixAsync(OriginalSource, expected, FixedSource).ConfigureAwait(True)
        End Sub

        <Fact>
        Public Async Sub VerifyFixAddAsClauseInForEachIEnumerable()
            Const OriginalSource As String =
"Imports System.Collections.Generic

Class Class1
    Sub Main()
       Dim annotatedNodesOrTokens As IEnumerable(Of String) = Nothing
       For Each annotatedNodeOrToken In annotatedNodesOrTokens
            Stop
       Next
    End Sub
End Class"
            Dim expected As DiagnosticResult = Verify.Diagnostic(AddAsClauseDiagnosticId).
                                WithLocation("/0/Test0.vb", 6, 17).
                                WithMessage("Option Strict On requires all variable declarations to have an 'As' clause.").
                                WithSeverity(DiagnosticSeverity.Error)

            Const FixedSource As String =
"Imports System.Collections.Generic

Class Class1
    Sub Main()
       Dim annotatedNodesOrTokens As IEnumerable(Of String) = Nothing
       For Each annotatedNodeOrToken As String In annotatedNodesOrTokens
            Stop
       Next
    End Sub
End Class"
            Await Verify.VerifyCodeFixAsync(OriginalSource, expected, FixedSource).ConfigureAwait(True)
        End Sub

        <Fact>
        Public Async Sub VerifyFixAddAsClauseInForEachListElement()
            Const OriginalSource As String =
"Imports System
Imports System.Collections.Generic

Class Class1
    Sub Main()
            Dim list1 As New List(Of Integer)({20, 30, 500})
            For Each element In list1
                Console.Write(element)
            Next
    End Sub
End Class"
            Dim expected As DiagnosticResult = Verify.Diagnostic(AddAsClauseDiagnosticId).
                                WithLocation("/0/Test0.vb", 7, 22).
                                WithMessage("Option Strict On requires all variable declarations to have an 'As' clause.").
                                WithSeverity(DiagnosticSeverity.Error)
            Const FixedSource As String =
"Imports System
Imports System.Collections.Generic

Class Class1
    Sub Main()
            Dim list1 As New List(Of Integer)({20, 30, 500})
            For Each element As Integer In list1
                Console.Write(element)
            Next
    End Sub
End Class"
            Await Verify.VerifyCodeFixAsync(OriginalSource, expected, FixedSource).ConfigureAwait(True)
        End Sub

        <Fact>
        Public Async Sub VerifyFixAddAsClauseToDim()
            Const OriginalSource As String =
"Class Class1
    Sub Main()
            Dim X = 1
    End Sub
End Class"
            Dim expected As DiagnosticResult = Verify.Diagnostic(AddAsClauseDiagnosticId).
                                WithLocation("/0/Test0.vb", 3, 17).
                                WithMessage("Option Strict On requires all variable declarations to have an 'As' clause.").
                                WithSeverity(DiagnosticSeverity.Error)
            Const FixedSource As String =
"Class Class1
    Sub Main()
            Dim X As Integer = 1
    End Sub
End Class"
            Await Verify.VerifyCodeFixAsync(OriginalSource, expected, FixedSource).ConfigureAwait(True)
        End Sub

    End Class

End Namespace
