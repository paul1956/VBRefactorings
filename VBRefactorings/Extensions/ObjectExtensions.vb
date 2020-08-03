' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Runtime.CompilerServices

Public Module ObjectExtensions

#Region "TypeSwitch on Func<T>"
    <Extension>
    Public Function TypeSwitch(Of TBaseType, TDerivedType1 As TBaseType, TDerivedType2 As TBaseType, TDerivedType3 As TBaseType, TResult)(obj As TBaseType, matchFunc1 As Func(Of TDerivedType1, TResult), matchFunc2 As Func(Of TDerivedType2, TResult), matchFunc3 As Func(Of TDerivedType3, TResult), Optional defaultFunc As Func(Of TBaseType, TResult) = Nothing) As TResult
        If TypeOf obj Is TDerivedType1 Then
            Return matchFunc1(DirectCast(obj, TDerivedType1))
        ElseIf TypeOf obj Is TDerivedType2 Then
            Return matchFunc2(DirectCast(obj, TDerivedType2))
        ElseIf TypeOf obj Is TDerivedType3 Then
            Return matchFunc3(DirectCast(obj, TDerivedType3))
        ElseIf defaultFunc IsNot Nothing Then
            Return defaultFunc(obj)
        Else
            Return Nothing
        End If
    End Function
#End Region

End Module
