' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

'Namespace Microsoft.CodeAnalysis.Shared.Extensions
Imports Microsoft.CodeAnalysis

Public Module IParameterSymbolExtensions

    <Runtime.CompilerServices.Extension>
    Public Function IsRefOrOut(symbol As IParameterSymbol) As Boolean
        Return symbol.RefKind <> RefKind.None
    End Function

End Module
'End Namespace
