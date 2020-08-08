' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Runtime.CompilerServices

Namespace Roslyn.UnitTestFramework
    Friend Module DictionaryExtensions
        ' Copied from ConcurrentDictionary since IDictionary doesn't have this useful method
        <Extension>
        Public Function GetOrAdd(Of TKey, TValue)(dictionary As IDictionary(Of TKey, TValue), key As TKey, [function] As Func(Of TKey, TValue)) As TValue
            Dim value As TValue = Nothing
            If Not dictionary.TryGetValue(key, value) Then
                value = [function](key)
                dictionary.Add(key, value)
            End If

            Return value
        End Function

        <Extension>
        Public Function GetOrAdd(Of TKey, TValue)(dictionary As IDictionary(Of TKey, TValue), key As TKey, [function] As Func(Of TValue)) As TValue
            Return dictionary.GetOrAdd(key, Function(underscore As TKey) [function]())
        End Function
    End Module
End Namespace
