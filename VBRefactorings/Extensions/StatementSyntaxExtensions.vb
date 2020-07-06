' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Runtime.CompilerServices
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Public Module StatementSyntaxExtensions

    <Extension()>
    Public Function GetAttributes(member As StatementSyntax) As SyntaxList(Of AttributeListSyntax)
        If member IsNot Nothing Then
            Select Case member.Kind
                Case SyntaxKind.ClassBlock,
                        SyntaxKind.InterfaceBlock,
                        SyntaxKind.ModuleBlock,
                        SyntaxKind.StructureBlock
                    Return DirectCast(member, TypeBlockSyntax).BlockStatement.AttributeLists
                Case SyntaxKind.EnumBlock
                    Return DirectCast(member, EnumBlockSyntax).EnumStatement.AttributeLists
                Case SyntaxKind.ClassStatement,
                        SyntaxKind.InterfaceStatement,
                        SyntaxKind.ModuleStatement,
                        SyntaxKind.StructureStatement
                    Return DirectCast(member, TypeStatementSyntax).AttributeLists
                Case SyntaxKind.EnumStatement
                    Return DirectCast(member, EnumStatementSyntax).AttributeLists
                Case SyntaxKind.EnumMemberDeclaration
                    Return DirectCast(member, EnumMemberDeclarationSyntax).AttributeLists
                Case SyntaxKind.FieldDeclaration
                    Return DirectCast(member, FieldDeclarationSyntax).AttributeLists
                Case SyntaxKind.EventBlock
                    Return DirectCast(member, EventBlockSyntax).EventStatement.AttributeLists
                Case SyntaxKind.EventStatement
                    Return DirectCast(member, EventStatementSyntax).AttributeLists
                Case SyntaxKind.PropertyBlock
                    Return DirectCast(member, PropertyBlockSyntax).PropertyStatement.AttributeLists
                Case SyntaxKind.PropertyStatement
                    Return DirectCast(member, PropertyStatementSyntax).AttributeLists
                Case SyntaxKind.FunctionBlock,
                        SyntaxKind.SubBlock,
                        SyntaxKind.ConstructorBlock,
                        SyntaxKind.OperatorBlock,
                        SyntaxKind.GetAccessorBlock,
                        SyntaxKind.SetAccessorBlock,
                        SyntaxKind.AddHandlerAccessorBlock,
                        SyntaxKind.RemoveHandlerAccessorBlock,
                        SyntaxKind.RaiseEventAccessorBlock
                    Return DirectCast(member, MethodBlockBaseSyntax).BlockStatement.AttributeLists
                Case SyntaxKind.SubStatement,
                        SyntaxKind.FunctionStatement,
                        SyntaxKind.SubNewStatement,
                        SyntaxKind.OperatorStatement,
                        SyntaxKind.GetAccessorStatement,
                        SyntaxKind.SetAccessorStatement,
                        SyntaxKind.AddHandlerAccessorStatement,
                        SyntaxKind.RemoveHandlerAccessorStatement,
                        SyntaxKind.RaiseEventAccessorStatement,
                        SyntaxKind.DeclareSubStatement,
                        SyntaxKind.DeclareFunctionStatement,
                        SyntaxKind.DelegateSubStatement,
                        SyntaxKind.DelegateFunctionStatement
                    Return DirectCast(member, MethodBaseSyntax).AttributeLists
            End Select
        End If

        Return Nothing
    End Function

    <Extension>
    Public Function GetModifiers(node As DeclarationStatementSyntax) As SyntaxTokenList
        If TypeOf node Is MethodBlockBaseSyntax Then
            Return DirectCast(node, MethodBlockBaseSyntax).BlockStatement.Modifiers
        End If
        If TypeOf node Is FieldDeclarationSyntax Then
            Return DirectCast(node, FieldDeclarationSyntax).Modifiers
        End If
        If TypeOf node Is DelegateStatementSyntax Then
            Return DirectCast(node, DelegateStatementSyntax).Modifiers
        End If
        If TypeOf node Is PropertyStatementSyntax Then
            Return DirectCast(node, PropertyStatementSyntax).Modifiers
        End If
        If TypeOf node Is PropertyBlockSyntax Then
            Return DirectCast(node, PropertyBlockSyntax).PropertyStatement.Modifiers
        End If
        Return New SyntaxTokenList()
    End Function

    <Extension()>
    Public Function IsConstructorInitializer(statement As StatementSyntax) As Boolean
        If statement.IsParentKind(SyntaxKind.ConstructorBlock) AndAlso
               DirectCast(statement.Parent, ConstructorBlockSyntax).Statements.FirstOrDefault() Is statement Then

            Dim invocation As ExpressionSyntax
            If statement.IsKind(SyntaxKind.CallStatement) Then
                invocation = DirectCast(statement, CallStatementSyntax).Invocation
            ElseIf statement.IsKind(SyntaxKind.ExpressionStatement) Then
                invocation = DirectCast(statement, ExpressionStatementSyntax).Expression
            Else
                Return False
            End If

            Dim expression As ExpressionSyntax = Nothing
            If invocation.IsKind(SyntaxKind.InvocationExpression) Then
                expression = DirectCast(invocation, InvocationExpressionSyntax).Expression
            ElseIf invocation.IsKind(SyntaxKind.SimpleMemberAccessExpression) Then
                expression = invocation
            End If

            If expression IsNot Nothing Then
                If expression.IsKind(SyntaxKind.SimpleMemberAccessExpression) Then
                    Dim memberAccess As MemberAccessExpressionSyntax = DirectCast(expression, MemberAccessExpressionSyntax)

                    Return memberAccess.IsInConstructorInitializer()
                End If
            End If
        End If

        Return False
    End Function

    <Extension>
    Public Function IsInConstructorInitializer(expression As MemberAccessExpressionSyntax) As Boolean
        Dim constructorInitializer As StatementSyntax = expression.GetAncestorsOrThis(Of StatementSyntax)().FirstOrDefault(Function(n As StatementSyntax) n.IsConstructorInitializer())

        If constructorInitializer Is Nothing Then
            Return False
        End If

        ' have to make sure we're not inside a lambda inside the constructor initializer.
        If expression.GetAncestorOrThis(Of LambdaExpressionSyntax)() IsNot Nothing Then
            Return False
        End If

        Return True
    End Function

    <Extension()>
    Public Function WithAttributeLists(member As StatementSyntax, attributeLists As SyntaxList(Of AttributeListSyntax)) As StatementSyntax
        If member IsNot Nothing Then
            Select Case member.Kind
                Case SyntaxKind.ClassBlock,
                        SyntaxKind.InterfaceBlock,
                        SyntaxKind.ModuleBlock,
                        SyntaxKind.StructureBlock
                    Dim typeBlock As TypeBlockSyntax = DirectCast(member, TypeBlockSyntax)
                    Dim newBegin As TypeStatementSyntax = DirectCast(typeBlock.BlockStatement.WithAttributeLists(attributeLists), TypeStatementSyntax)
                    Return typeBlock.WithBlockStatement(newBegin)
                Case SyntaxKind.EnumBlock
                    Dim enumBlock As EnumBlockSyntax = DirectCast(member, EnumBlockSyntax)
                    Dim newEnumStatement As EnumStatementSyntax = enumBlock.EnumStatement.WithAttributeLists(attributeLists)
                    Return enumBlock.WithEnumStatement(newEnumStatement)
                Case SyntaxKind.ClassStatement
                    Return DirectCast(member, ClassStatementSyntax).WithAttributeLists(attributeLists)
                Case SyntaxKind.InterfaceStatement
                    Return DirectCast(member, InterfaceStatementSyntax).WithAttributeLists(attributeLists)
                Case SyntaxKind.ModuleStatement
                    Return DirectCast(member, ModuleStatementSyntax).WithAttributeLists(attributeLists)
                Case SyntaxKind.StructureStatement
                    Return DirectCast(member, StructureStatementSyntax).WithAttributeLists(attributeLists)
                Case SyntaxKind.EnumStatement
                    Return DirectCast(member, EnumStatementSyntax).WithAttributeLists(attributeLists)
                Case SyntaxKind.EnumMemberDeclaration
                    Return DirectCast(member, EnumMemberDeclarationSyntax).WithAttributeLists(attributeLists)
                Case SyntaxKind.FieldDeclaration
                    Return DirectCast(member, FieldDeclarationSyntax).WithAttributeLists(attributeLists)
                Case SyntaxKind.EventBlock
                    Dim eventBlock As EventBlockSyntax = DirectCast(member, EventBlockSyntax)
                    Dim newEventStatement As EventStatementSyntax = DirectCast(eventBlock.EventStatement.WithAttributeLists(attributeLists), EventStatementSyntax)
                    Return eventBlock.WithEventStatement(newEventStatement)
                Case SyntaxKind.EventStatement
                    Return DirectCast(member, EventStatementSyntax).WithAttributeLists(attributeLists)
                Case SyntaxKind.PropertyBlock
                    Dim propertyBlock As PropertyBlockSyntax = DirectCast(member, PropertyBlockSyntax)
                    Dim newPropertyStatement As PropertyStatementSyntax = propertyBlock.PropertyStatement.WithAttributeLists(attributeLists)
                    Return propertyBlock.WithPropertyStatement(newPropertyStatement)
                Case SyntaxKind.PropertyStatement
                    Return DirectCast(member, PropertyStatementSyntax).WithAttributeLists(attributeLists)
                Case SyntaxKind.FunctionBlock,
                        SyntaxKind.SubBlock,
                        SyntaxKind.ConstructorBlock,
                        SyntaxKind.OperatorBlock,
                        SyntaxKind.GetAccessorBlock,
                        SyntaxKind.SetAccessorBlock,
                        SyntaxKind.AddHandlerAccessorBlock,
                        SyntaxKind.RemoveHandlerAccessorBlock,
                        SyntaxKind.RaiseEventAccessorBlock
                    Dim methodBlock As MethodBlockBaseSyntax = DirectCast(member, MethodBlockBaseSyntax)
                    Dim newBegin As StatementSyntax = methodBlock.BlockStatement.WithAttributeLists(attributeLists)
                    Return methodBlock.ReplaceNode(methodBlock.BlockStatement, newBegin)
                Case SyntaxKind.SubStatement,
                        SyntaxKind.FunctionStatement
                    Return DirectCast(member, MethodStatementSyntax).WithAttributeLists(attributeLists)
                Case SyntaxKind.SubNewStatement
                    Return DirectCast(member, SubNewStatementSyntax).WithAttributeLists(attributeLists)
                Case SyntaxKind.OperatorStatement
                    Return DirectCast(member, OperatorStatementSyntax).WithAttributeLists(attributeLists)
                Case SyntaxKind.GetAccessorStatement,
                         SyntaxKind.SetAccessorStatement,
                         SyntaxKind.AddHandlerAccessorStatement,
                         SyntaxKind.RemoveHandlerAccessorStatement,
                         SyntaxKind.RaiseEventAccessorStatement
                    Return DirectCast(member, AccessorStatementSyntax).WithAttributeLists(attributeLists)
                Case SyntaxKind.DeclareSubStatement,
                         SyntaxKind.DeclareFunctionStatement
                    Return DirectCast(member, DeclareStatementSyntax).WithAttributeLists(attributeLists)
                Case SyntaxKind.DelegateSubStatement,
                        SyntaxKind.DelegateFunctionStatement
                    Return DirectCast(member, DelegateStatementSyntax).WithAttributeLists(attributeLists)
                Case SyntaxKind.IncompleteMember
                    Return DirectCast(member, IncompleteMemberSyntax).WithAttributeLists(attributeLists)
            End Select
        End If

        Return Nothing
    End Function

End Module
