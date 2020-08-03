' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Runtime.CompilerServices
Imports System.Threading
Imports Microsoft
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports VBRefactorings.Utilities

Public Module ExpressionSyntaxExtensions

    Private Function IsNameOrMemberAccessButNoExpression(node As SyntaxNode) As Boolean
        If node.IsKind(SyntaxKind.SimpleMemberAccessExpression) Then
            Dim memberAccess As MemberAccessExpressionSyntax = DirectCast(node, MemberAccessExpressionSyntax)

            Return memberAccess.Expression.IsKind(SyntaxKind.IdentifierName) OrElse IsNameOrMemberAccessButNoExpression(memberAccess.Expression)
        End If

        Return node.IsKind(SyntaxKind.IdentifierName)
    End Function

    Private Function ContainsOpenName(name As NameSyntax) As Boolean
        If TypeOf name Is QualifiedNameSyntax Then
            Dim qualifiedName As QualifiedNameSyntax = DirectCast(name, QualifiedNameSyntax)
            Return ContainsOpenName(qualifiedName.Left) OrElse ContainsOpenName(qualifiedName.Right)
        ElseIf TypeOf name Is CodeAnalysis.VisualBasic.Syntax.GenericNameSyntax Then
            Return DirectCast(name, GenericNameSyntax).IsUnboundGenericName
        Else
            Return False
        End If
    End Function

    <Extension>
    Public Function IsUnboundGenericName(Name As GenericNameSyntax) As Boolean
        Return Name.TypeArgumentList.Arguments.Any(SyntaxKind.OmittedArgument)
    End Function

    <Extension>
    Public Function DetermineType(expression As ExpressionSyntax,
                                      Model As SemanticModel,
                                      CancelToken As CancellationToken) As (_ITypeSymbol As ITypeSymbol, _Error As Boolean)
        ' If a parameter appears to have a void return type, then just use 'object' instead.
        Try
            If expression IsNot Nothing Then
                Dim typeInfo As TypeInfo = Model.GetTypeInfo(expression, CancelToken)
                Dim symbolInfo As SymbolInfo = Model.GetSymbolInfo(expression, CancelToken)
                If typeInfo.Type IsNot Nothing Then
                    If typeInfo.Type IsNot Nothing AndAlso typeInfo.Type.SpecialType = SpecialType.System_Void Then
                        Return (Model.Compilation.ObjectType, False)
                    End If

                    If typeInfo.Type.IsErrorType Then
                        Return (Model.Compilation.ObjectType, True)
                    ElseIf SymbolEqualityComparer.Default.Equals(typeInfo.Type, Model.Compilation.ObjectType) Then
                        Return (Model.Compilation.ObjectType, False)
                    End If
                End If
                Dim symbol As ISymbol = If(typeInfo.Type, symbolInfo.GetAnySymbol())
                If symbol IsNot Nothing Then
                    Return (symbol.ConvertToType(Model.Compilation), False)
                End If

                If TypeOf expression Is CollectionInitializerSyntax Then
                    Dim collectionInitializer As CollectionInitializerSyntax = DirectCast(expression, CollectionInitializerSyntax)
                    Return DetermineType(collectionInitializer, Model, CancelToken)
                End If
            End If
        Catch ex As Exception
            Stop
        End Try
        Return (Model.Compilation.ObjectType, True)
    End Function

    <Extension()>
    Private Function DetermineType(collectionInitializer As CollectionInitializerSyntax,
                                      semanticModel As SemanticModel,
                                      CancelToken As CancellationToken) As (ITypeSymbol, Boolean)
        Dim rank As Integer = 1
        While collectionInitializer.Initializers.Count > 0 AndAlso
                  collectionInitializer.Initializers(0).Kind = SyntaxKind.CollectionInitializer
            rank += 1
            collectionInitializer = DirectCast(collectionInitializer.Initializers(0), CollectionInitializerSyntax)
        End While

        Dim type As (ITypeSymbol, Boolean) = collectionInitializer.Initializers.FirstOrDefault().DetermineType(semanticModel, CancelToken)
        Return (semanticModel.Compilation.CreateArrayTypeSymbol(type.Item1, rank), type.Item2)
    End Function

End Module
