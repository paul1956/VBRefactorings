' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Public Module CheckSyntax

    Public Function CheckSyntax(source As String) As (Boolean, String)
        Dim tree As CompilationUnitSyntax = SyntaxFactory.ParseCompilationUnit(source)
        Dim strDetail As String = ""
        If tree.GetDiagnostics().Count = 0 Then
            strDetail &= "Diagnostic test passed!!!"
            Return (True, strDetail)
        End If

        Dim Diagnostics As IEnumerable(Of Diagnostic) = tree.GetDiagnostics
        For i As Integer = 1 To Diagnostics.Count
            Dim DiagnosticItem As Diagnostic = Diagnostics(i)
            strDetail &= i.ToString() & ". Info: " & DiagnosticItem.ToString()
            strDetail &= " Warning Level: " & DiagnosticItem.WarningLevel.ToString()
            strDetail &= " Severity Level: " & DiagnosticItem.Severity.ToString() & vbCrLf
            strDetail &= " Location: " & DiagnosticItem.Location.Kind.ToString()
            strDetail &= " Character at: " & DiagnosticItem.Location.GetLineSpan.Span.Start.Character.ToString()
            strDetail &= " On Line: " & DiagnosticItem.Location.GetLineSpan.Span.Start.Line.ToString()
        Next
        Return (False, strDetail)
    End Function

End Module
