' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Runtime.CompilerServices
' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.


Namespace Roslyn.UnitTestFramework
    Friend Module EnumerableExtensions
        <Extension>
        Public Function OrderBy(Of T)(source As IEnumerable(Of T), comparer As IComparer(Of T)) As IEnumerable(Of T)
            Return source.OrderBy(Function(t_) t_, comparer)
        End Function

        <Extension>
        Public Function OrderBy(Of T)(source As IEnumerable(Of T), compare As Comparison(Of T)) As IEnumerable(Of T)
            Return source.OrderBy(New ComparisonComparer(Of T)(compare))
        End Function
        Private Class ComparisonComparer(Of T)
            Inherits Comparer(Of T)

            Private ReadOnly _compare As Comparison(Of T)

            Public Sub New(compare As Comparison(Of T))
                _compare = compare
            End Sub

            Public Overrides Function Compare(x As T, y As T) As Integer
                Return _compare(x, y)
            End Function
        End Class
    End Module

End Namespace
