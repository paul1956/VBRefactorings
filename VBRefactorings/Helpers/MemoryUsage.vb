' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Runtime.CompilerServices

Public Module MemoryUsage

    Public Function IsEnoughMemoryReport(OperationsCount As Integer, DocumentIdCount As Integer, CodeActionCount As Integer, codeActionTotal As Integer, Optional Limit As Long = 2516582400L) As Boolean
        Dim c As Process = Process.GetCurrentProcess
        Dim WorkingSet As Long = c.WorkingSet64
        Dim PrivateBytes As Long = c.PagedMemorySize64
        Dim GCTotalMemory As Long = GC.GetTotalMemory(True)
        Debug.Print($"OperationsCount = {OperationsCount}, DocumentIdCount = {DocumentIdCount}, CodeActionCount = {CodeActionCount}, codeActionTotal = {codeActionTotal}, Memory Usage (Working Set): {WorkingSet / 1024:N0}K VM, Size (Private Bytes): {PrivateBytes / 1024:N0}, K GC Total Memory: {GCTotalMemory:N0} bytes")
        If WorkingSet > Limit Then
            Return False
        End If
        Return True
    End Function

    Public Function IsEnoughMemoryReport(Limit As Long, <CallerMemberName> Optional memberName As String = Nothing, <CallerLineNumber()> Optional sourceLineNumber As Integer = 0) As Boolean
        Dim c As Process = Process.GetCurrentProcess
        Dim WorkingSet As Long = c.WorkingSet64
        Dim PrivateBytes As Long = c.PagedMemorySize64
        Dim GCTotalMemory As Long = GC.GetTotalMemory(True)
        Debug.Print($"Called By {memberName}, Member Line {sourceLineNumber}, Memory Usage (Working Set): {WorkingSet / 1024:N0}K VM, Size (Private Bytes): {PrivateBytes / 1024:N0}, K GC Total Memory: {GCTotalMemory:N0} bytes")
        If WorkingSet > Limit Then
            Return False
        End If
        Return True
    End Function

End Module
