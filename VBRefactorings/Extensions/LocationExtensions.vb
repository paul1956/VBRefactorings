' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Runtime.CompilerServices
Imports System.Threading
Imports Microsoft.CodeAnalysis

Public Module LocationExtensions

    <Extension>
    Public Function FindNode(location As Location, getInnermostNodeForTie As Boolean, CancelToken As CancellationToken) As SyntaxNode
        Return location.SourceTree.GetRoot(CancelToken).FindNode(location.SourceSpan, getInnermostNodeForTie:=getInnermostNodeForTie)
    End Function

End Module
