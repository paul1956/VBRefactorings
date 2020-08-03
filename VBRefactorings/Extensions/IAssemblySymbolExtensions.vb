' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Runtime.CompilerServices
Imports Microsoft.CodeAnalysis

Public Module IAssemblySymbolExtensions

    <Extension>
    Friend Function IsSameAssemblyOrHasFriendAccessTo(assembly As IAssemblySymbol, toAssembly As IAssemblySymbol) As Boolean
        Return SymbolEqualityComparer.Default.Equals(assembly, toAssembly) OrElse (assembly.IsInteractive AndAlso toAssembly.IsInteractive) OrElse toAssembly.GivesAccessTo(assembly)
    End Function

End Module
