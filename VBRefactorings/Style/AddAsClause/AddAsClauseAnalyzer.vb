' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports VBRefactorings

Namespace Style

    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class AddAsClauseAnalyzer
        Inherits DiagnosticAnalyzer

        Private Const Description As String = "Option Strict On requires all variable declarations to have an 'As' clause."
        Friend Const Category As String = SupportedCategories.Style
        Friend Const MessageFormat As String = "Option Strict On requires all variable declarations to have an 'As' clause."
        Friend Const Title As String = "Option Strict On requires all variable declarations to have an 'As' clause."

        Friend Shared Rule As New DiagnosticDescriptor(
                        AddAsClauseDiagnosticId,
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

                Dim diag As Diagnostic = Diagnostic.Create(Rule, ControlVariable.GetLocation(), MessageFormat)
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
                    Case SyntaxKind.VariableDeclarator
                        If CType(ControlVariable, VariableDeclaratorSyntax).AsClause IsNot Nothing Then
                            Exit Sub
                        End If
                    Case Else
                        Stop
                        Exit Sub
                End Select
                Dim oldControlVariable As IdentifierNameSyntax = CType(ControlVariable, IdentifierNameSyntax)

                If oldControlVariable.IsExplicitlyDeclared(model) Then
                    Exit Sub
                End If

                Dim diag As Diagnostic = Diagnostic.Create(Rule, ControlVariable.GetLocation(), MessageFormat)
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
                            If param.AsClause Is Nothing Then
                                diag = Diagnostic.Create(Rule, param.GetLocation(), MessageFormat)
                                context.ReportDiagnostic(diag)
                            End If
                        Next
                        Exit Sub
                    Case SyntaxKind.VariableDeclarator
                        Dim VariableDeclaration As VariableDeclaratorSyntax = DirectCast(context.Node, VariableDeclaratorSyntax)
                        If VariableDeclaration.AsClause IsNot Nothing Then
                            Exit Sub
                        End If
                        If VariableDeclaration.Initializer Is Nothing Then
                            Exit Sub
                        End If

                        Dim node As ExpressionSyntax = VariableDeclaration.Initializer.Value
                        If node.Kind = SyntaxKind.NothingLiteralExpression Then
                            Exit Sub
                        End If

                        If Not node.MatchesKind(
                            SyntaxKind.AddExpression,
                            SyntaxKind.AndAlsoExpression,
                            SyntaxKind.AndExpression,
                            SyntaxKind.ArrayCreationExpression,
                            SyntaxKind.ArrayType,
                            SyntaxKind.AwaitExpression,
                            SyntaxKind.BinaryConditionalExpression,
                            SyntaxKind.CollectionInitializer,
                            SyntaxKind.ConcatenateExpression,
                            SyntaxKind.ConditionalAccessExpression,
                            SyntaxKind.CTypeExpression,
                            SyntaxKind.DictionaryAccessExpression,
                            SyntaxKind.DirectCastExpression,
                            SyntaxKind.EqualsExpression,
                            SyntaxKind.FalseLiteralExpression,
                            SyntaxKind.GenericName,
                            SyntaxKind.GreaterThanExpression,
                            SyntaxKind.IdentifierName,
                            SyntaxKind.InterpolatedStringExpression,
                            SyntaxKind.InvocationExpression,
                            SyntaxKind.IsNotExpression,
                            SyntaxKind.LessThanExpression,
                            SyntaxKind.MeExpression,
                            SyntaxKind.MultiLineFunctionLambdaExpression,
                            SyntaxKind.MultiLineSubLambdaExpression,
                            SyntaxKind.NotEqualsExpression,
                            SyntaxKind.NotExpression,
                            SyntaxKind.NumericLiteralExpression,
                            SyntaxKind.ObjectCreationExpression,
                            SyntaxKind.OrElseExpression,
                            SyntaxKind.OrExpression,
                            SyntaxKind.ParenthesizedExpression,
                            SyntaxKind.PredefinedCastExpression,
                            SyntaxKind.PredefinedType,
                            SyntaxKind.QueryExpression,
                            SyntaxKind.SimpleMemberAccessExpression,
                            SyntaxKind.SingleLineFunctionLambdaExpression,
                            SyntaxKind.SingleLineSubLambdaExpression,
                            SyntaxKind.StringLiteralExpression,
                            SyntaxKind.SubtractExpression,
                            SyntaxKind.TernaryConditionalExpression,
                            SyntaxKind.TrueLiteralExpression,
                            SyntaxKind.TryCastExpression,
                            SyntaxKind.TupleExpression,
                            SyntaxKind.TypeOfIsExpression,
                            SyntaxKind.UnaryMinusExpression
                            ) Then
                            Debug.Print($"Node.IsKind list is missing {node.Kind}")
                            Stop    ' Don't think this list is complete
                        End If
                        Dim expressionConvertedType As ITypeSymbol = model.GetTypeInfo(VariableDeclaration.Initializer.Value).ConvertedType
                        If expressionConvertedType Is Nothing OrElse expressionConvertedType.ToString().ToUpperInvariant.Contains("ANONYMOUS TYPE:") Then
                            Exit Sub
                        End If

                        diag = Diagnostic.Create(Rule, VariableDeclaration.GetLocation(), MessageFormat)
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
