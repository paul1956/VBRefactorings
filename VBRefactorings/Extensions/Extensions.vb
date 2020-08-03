' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Runtime.CompilerServices

Public Module Extensions
    <Extension>
    Public Function ToDiagnosticId(ByVal diagnosticId As DiagnosticId) As String
        Return $"CC{CInt(Math.Truncate(diagnosticId)):D4}"
    End Function

    <Extension>
    Public Function AddRange(Of K, V)(ByVal dictionary As IDictionary(Of K, V), ByVal newValues As IDictionary(Of K, V)) As IDictionary(Of K, V)
        If dictionary Is Nothing OrElse newValues Is Nothing Then
            Return dictionary
        End If
        For Each kv As KeyValuePair(Of K, V) In newValues
            dictionary.Add(kv)
        Next kv
        Return dictionary
    End Function

    <Extension>
    Public Function EndsWithAny(ByVal text As String, ParamArray ByVal values() As String) As Boolean
        Return text.EndsWithAny(StringComparison.CurrentCulture, values)
    End Function

    <Extension>
    Public Function EndsWithAny(ByVal text As String, ByVal comparisonType As StringComparison, ParamArray ByVal values() As String) As Boolean
        For Each value As String In values
            If text.EndsWith(value, comparisonType) Then
                Return True
            End If
        Next value
        Return False
    End Function

    <Extension>
    Public Function ToLowerCaseFirstLetter(ByVal text As String) As String
        If String.IsNullOrWhiteSpace(text) Then
            Return text
        End If
        If text.Length = 1 Then
            Return text.ToLower()
        End If
        Return Char.ToLowerInvariant(text.Chars(0)) & text.Substring(1)
    End Function
End Module
