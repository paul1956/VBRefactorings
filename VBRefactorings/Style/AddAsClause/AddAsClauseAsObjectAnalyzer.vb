' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Style

    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class AddAsClauseAsObjectAnalyzer
        Inherits DiagnosticAnalyzer

        Private Const Category As String = SupportedCategories.Style
        Private Const Description As String = "You should replace 'As Object' with more specific type."
        Private Const MessageFormat As String = "As Object should not be used."
        Private Const Title As String = "As Object should not be used when a more specific type is available"

        Friend Shared Rule As New DiagnosticDescriptor(
                    ChangeAsObjectToMoreSpecificDiagnosticId,
                    Title,
                    MessageFormat,
                    Category,
                    DiagnosticSeverity.Error,
                    isEnabledByDefault:=True,
                    Description,
                    helpLinkUri:=Nothing,
                    Array.Empty(Of String)
                    )

        Public Overrides ReadOnly Property SupportedDiagnostics As ImmutableArray(Of DiagnosticDescriptor)
            Get
                Return ImmutableArray.Create(Rule)
            End Get
        End Property

        Private Shared Sub AnalyzeClassForEachStatement(context As SyntaxNodeAnalysisContext)
            Try
                Dim model As SemanticModel = context.SemanticModel
                If model.OptionStrict = OptionStrict.On AndAlso model.OptionInfer = True Then Exit Sub
                Dim ForEachStatementDeclaration As ForEachStatementSyntax = DirectCast(context.Node, ForEachStatementSyntax)
                If ForEachStatementDeclaration Is Nothing Then Exit Sub
                Dim ControlVariable As VisualBasicSyntaxNode = ForEachStatementDeclaration.ControlVariable
                If ControlVariable Is Nothing Then Exit Sub
                Select Case ControlVariable.Kind
                    Case SyntaxKind.IdentifierName
                        If CType(ControlVariable, IdentifierNameSyntax).IsExplicitlyDeclared(model) Then
                            Exit Sub
                        End If
                    Case SyntaxKind.VariableDeclarator
                        Exit Sub
                    Case Else
                        Stop
                End Select

                Dim diag As Diagnostic = Diagnostic.Create(Rule, ControlVariable.GetLocation())
                context.ReportDiagnostic(diag)
            Catch ex As Exception When ex.HResult <> (New OperationCanceledException).HResult
                Stop
                Throw
            End Try
        End Sub

        Private Shared Sub AnalyzeClassForStatement(context As SyntaxNodeAnalysisContext)
            Try
                Dim model As SemanticModel = context.SemanticModel
                If model.OptionStrict = OptionStrict.On AndAlso model.OptionInfer = True Then Exit Sub
                Dim ForStatementDeclaration As ForStatementSyntax = DirectCast(context.Node, ForStatementSyntax)
                If ForStatementDeclaration Is Nothing Then Exit Sub
                Dim ControlVariable As VisualBasicSyntaxNode = ForStatementDeclaration.ControlVariable
                If ControlVariable Is Nothing Then Exit Sub
                Select Case ControlVariable.Kind
                    Case SyntaxKind.IdentifierName
                        Dim ID As SimpleNameSyntax = CType(ControlVariable, SimpleNameSyntax)
                        If IsExplicitlyDeclared(CType(ID, IdentifierNameSyntax), model) Then
                            Exit Sub
                        End If
                    Case SyntaxKind.VariableDeclarator
                        If CType(ControlVariable, VariableDeclaratorSyntax).AsClause IsNot Nothing Then
                            Exit Sub
                        End If
                    Case Else
                        Stop
                        Exit Sub
                End Select

                Dim diag As Diagnostic = Diagnostic.Create(Rule, ControlVariable.GetLocation())
                context.ReportDiagnostic(diag)
            Catch ex As Exception When ex.HResult <> (New OperationCanceledException).HResult
                Stop
                Throw
            End Try
        End Sub

        Private Shared Sub AnalyzeClassVariable(context As SyntaxNodeAnalysisContext)
            Try
                Dim model As SemanticModel = context.SemanticModel
                If model.OptionStrict = OptionStrict.On AndAlso model.OptionInfer = True Then
                    Exit Sub
                End If

                Dim diag As Diagnostic = Nothing
                Select Case context.Node.Kind
                    Case SyntaxKind.FunctionLambdaHeader, SyntaxKind.SubLambdaHeader
                        Dim _LambdaHeaderSyntax As LambdaHeaderSyntax = DirectCast(context.Node, LambdaHeaderSyntax)
                        If _LambdaHeaderSyntax.ParameterList Is Nothing Then
                            Exit Sub
                        End If
                        For Each param As ParameterSyntax In _LambdaHeaderSyntax.ParameterList.Parameters
                            If param.AsClause Is Nothing OrElse param.AsClause.ToString <> "As Object" Then
                                Continue For
                            End If
                        Next
                        Exit Sub
                    Case SyntaxKind.VariableDeclarator
                        Dim VariableDeclaration As VariableDeclaratorSyntax = DirectCast(context.Node, VariableDeclaratorSyntax)
                        If VariableDeclaration.AsClause Is Nothing Then
                            Exit Sub
                        End If
                        If VariableDeclaration.Initializer Is Nothing Then
                            Exit Sub
                        End If

                        If VariableDeclaration.AsClause Is Nothing OrElse VariableDeclaration.AsClause.ToString <> "As Object" Then
                            Exit Sub
                        End If
                        Dim Expression As ExpressionSyntax = VariableDeclaration.Initializer.Value
                        Dim variableITypeSymbol As (_ITypeSymbol As ITypeSymbol, _Error As Boolean) = Expression.DetermineType(context.SemanticModel, context.CancellationToken)
                        If variableITypeSymbol._Error Then
                            Exit Sub
                        End If
                        If variableITypeSymbol._ITypeSymbol.ToString = "Object" Then
                            Exit Sub
                        End If
                        diag = Diagnostic.Create(Rule, VariableDeclaration.GetLocation())
                    Case Else
                        Stop
                End Select
                context.ReportDiagnostic(diag)
            Catch ex As Exception When ex.HResult <> (New OperationCanceledException).HResult
                Stop
                Throw
            End Try
        End Sub

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf AnalyzeClassVariable, SyntaxKind.VariableDeclarator)
            context.RegisterSyntaxNodeAction(AddressOf AnalyzeClassForStatement, SyntaxKind.ForStatement)
            context.RegisterSyntaxNodeAction(AddressOf AnalyzeClassForEachStatement, SyntaxKind.ForEachStatement)
        End Sub

    End Class

End Namespace
