﻿' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Public Module ExceptionUtilities

    Public ReadOnly Property UnreachableException() As Exception
        Get
            Return New InvalidOperationException("This program location is thought to be unreachable.")
        End Get
    End Property

    Public Function UnexpectedValue(o As Object) As Exception
        Dim output As String = String.Format("Unexpected value '{0}' of type '{1}'", o, If(o IsNot Nothing, o.GetType().FullName, "<unknown>"))
        Debug.Assert(False, output)

        ' We do not throw from here because we don't want all Watson reports to be bucketed to this call.
        Return New InvalidOperationException(output)
    End Function
End Module
