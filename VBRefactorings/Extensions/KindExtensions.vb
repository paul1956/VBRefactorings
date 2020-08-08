' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Runtime.CompilerServices
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Public Module KindExtensions

    <Extension()>
    Public Function MatchesKind(Expression As ExpressionSyntax, ParamArray KindList() As SyntaxKind) As Boolean
        For Each kind As SyntaxKind In KindList
            If Expression.Kind = kind Then
                Return True
            End If
        Next
        Return False
    End Function

End Module
