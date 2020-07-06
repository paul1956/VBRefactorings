' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Testing
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Imports VBRefactorings
Imports Xunit

Imports Verify = Microsoft.CodeAnalysis.VisualBasic.Testing.MSTest.CodeFixVerifier(Of VBRefactorings.Style.AddAsClauseAnalyzer, VBRefactorings.Style.AddAsClauseCodeFixProvider)

Namespace AddAsClause.UnitTest

    <TestClass()> Public Class AddAsClauseUnitTests

        <Fact>
        Public Sub AssertTestAddAsClauseInDiagnosticArray1()
            Dim comp As VisualBasicCompilation = CreateCompilationWithMscorlibAndVBRuntime("
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Diagnostics

Partial Public MustInherit Class DiagnosticVerifier
    Protected Shared Function GetSortedDiagnosticsFromDocuments(analyzer As DiagnosticAnalyzer, documents As Document()) As Diagnostic()
        Dim diagnostics As List(Of Diagnostic) = New List(Of Diagnostic)
        Dim results = SortDiagnostics(diagnostics)
        Return results
    End Function
    Private Shared Function SortDiagnostics(diagnostics As IEnumerable(Of Diagnostic)) As Diagnostic()
        Return Nothing
    End Function
End Class")
            Dim tree As SyntaxTree = comp.SyntaxTrees(0)

            Dim model As SemanticModel = comp.GetSemanticModel(tree)
            Dim nodes As IEnumerable(Of SyntaxNode) = tree.GetCompilationUnitRoot().DescendantNodes()
            Dim node As InvocationExpressionSyntax = nodes.OfType(Of InvocationExpressionSyntax)().Single()

            Dim TypeInfo As TypeInfo = model.GetTypeInfo(node)
            Xunit.Assert.Equal("Microsoft.CodeAnalysis.Diagnostic()", TypeInfo.ConvertedType.ToTestDisplayString())
        End Sub

        <Fact>
        Public Sub AssertTestForeachRepro()
            Dim comp As VisualBasicCompilation = CreateCompilationWithMscorlibAndVBRuntime(
"Options Strict On
Imports  System
Module Module1
    Sub M(s As String)
        For Each c In s
        Next
    End Sub
End Module", "a.vb")

            Dim tree As SyntaxTree = comp.SyntaxTrees(0)

            Dim model As SemanticModel = comp.GetSemanticModel(tree)
            Dim nodes As IEnumerable(Of SyntaxNode) = tree.GetCompilationUnitRoot().DescendantNodes()
            Dim node As ForEachStatementSyntax = nodes.OfType(Of ForEachStatementSyntax)().Single()

            Dim foreachInfo As ForEachStatementInfo = model.GetForEachStatementInfo(node)
            Xunit.Assert.Equal("System.Char", foreachInfo.ElementType.ToTestDisplayString())
        End Sub

        <Fact>
        Public Async Sub VerifyDiagnosticAddAsClauseInForDeclared()
            Const OriginalSource As String =
"Option Strict On
Imports  System
Class Class1
    Sub Main()
            Dim I as Integer
            For I = 0 to 1
                Console.Write(i)
            Next
    End Sub
End Class"
            Await Verify.VerifyAnalyzerAsync(OriginalSource).ConfigureAwait(True)
        End Sub

        <Fact>
        Public Sub VerifyDiagnosticAddAsClauseInForEachWithParameter()
            Const OriginalSource As String =
"Class Class1
        Public Function C(sym As Integer) As Integer
            Dim x(10) As Integer
            For Each sym In x
                Return sym
            Next
            Return 0
        End Function
End Class"
            Verify.VerifyAnalyzerAsync(OriginalSource)

        End Sub

        <Fact>
        Public Sub VerifyDiagnosticAddAsClauseInForParameter()
            Const OriginalSource As String =
"Class Class1
    Sub Main(I as Integer)
            For I = 0 to 1
                Console.Write(i)
            Next
    End Sub
End Class"
            Verify.VerifyAnalyzerAsync(OriginalSource)
        End Sub

        <Fact>
        Public Async Sub VerifyFixAddAsClauseInDimNewStatement()
            Const OriginalSource As String =
"Imports  System.Collections.Generic
Module Program
    Public Sub Main()
        Dim s = New List(Of Long)
    End Sub
End Module"
            Dim expected As DiagnosticResult = Verify.Diagnostic(AddAsClauseDiagnosticId) _
                                .WithLocation("/0/Test0.vb", 4, 13).
                                WithMessage("Option Strict On requires all variable declarations to have an 'As' clause.").
                                WithSeverity(DiagnosticSeverity.Error)

            Const FixedSource As String =
"Imports  System.Collections.Generic
Module Program
    Public Sub Main()
        Dim s As New List(Of Long)
    End Sub
End Module"
            Await Verify.VerifyCodeFixAsync(OriginalSource, expected, FixedSource).ConfigureAwait(True)
        End Sub

    End Class

End Namespace
