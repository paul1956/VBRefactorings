﻿' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Runtime.CompilerServices

Public Module DiagnosticIdExtensions
    <Extension>
    Public Function ToDiagnosticId(diagnosticId As DiagnosticId) As String
        Return $"CC{CInt(Math.Truncate(diagnosticId)):D4}"
    End Function

End Module
