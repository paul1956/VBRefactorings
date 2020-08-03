' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Runtime.CompilerServices
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Public Module IdentifierNameSyntaxExtensions

    ''' <summary>
    ''' Returns True if a ControlVariable Is Explicitly declared somewhere other than the current statement
    ''' </summary>
    ''' <param name="Identifier"></param>
    ''' <param name="Model">SemanticModel Model</param>
    ''' <returns></returns>
    <Extension>
    Public Function IsExplicitlyDeclared(Identifier As IdentifierNameSyntax, Model As SemanticModel) As Boolean
        Dim IdentifierSymbolInfo As SymbolInfo = Model.GetSymbolInfo(Identifier)
        Dim IdentifierSymbol As ISymbol = IdentifierSymbolInfo.Symbol
        If IdentifierSymbol Is Nothing Then
            Return False
        End If
        Dim _symbol As ISymbol = IdentifierSymbolInfo.Symbol
        If _symbol IsNot Nothing Then
            If TryCast(_symbol, IParameterSymbol) IsNot Nothing Then
                Return True
            End If
        End If

        Dim DeclaredCount As Integer = IdentifierSymbolInfo.Symbol.DeclaringSyntaxReferences.Count
        For Each DeclaringSyntaxReferences As SyntaxReference In IdentifierSymbolInfo.Symbol.DeclaringSyntaxReferences
            If Identifier.Span.Overlap(DeclaringSyntaxReferences.Span) IsNot Nothing Then
                DeclaredCount -= 1
            End If
        Next
        Return DeclaredCount > 0
    End Function

End Module
