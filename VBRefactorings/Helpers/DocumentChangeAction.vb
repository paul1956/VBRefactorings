' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Threading
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeActions
Imports Microsoft.CodeAnalysis.Text

Public NotInheritable Class DocumentChangeAction
    Inherits CodeAction

    Private ReadOnly _createChangedDocument As Func(Of CancellationToken, Task(Of Document))
    Private ReadOnly _title As String
    Private _severity As DiagnosticSeverity
    Private _textSpan As TextSpan

    Public Sub New(ByVal _TextSpan As TextSpan, ByVal _Severity As DiagnosticSeverity, ByVal title As String, ByVal createChangedDocument As Func(Of CancellationToken, Task(Of Document)))
        Me._severity = _Severity
        Me._textSpan = _TextSpan
        _title = title
        _createChangedDocument = createChangedDocument
    End Sub

    Public Property Severity() As DiagnosticSeverity
        Get
            Return _severity
        End Get
        Private Set(ByVal value As DiagnosticSeverity)
            _severity = value
        End Set
    End Property

    Public Property TextSpan() As TextSpan
        Get
            Return _textSpan
        End Get
        Private Set(ByVal value As TextSpan)
            _textSpan = value
        End Set
    End Property

    Public Overrides ReadOnly Property Title() As String
        Get
            Return _title
        End Get
    End Property

    Protected Overrides Function GetChangedDocumentAsync(ByVal CancelToken As CancellationToken) As Task(Of Document)
        If _createChangedDocument Is Nothing Then
            Return MyBase.GetChangedDocumentAsync(CancelToken)
        End If

        Dim task As Task(Of Document) = _createChangedDocument.Invoke(CancelToken)
        Return task
    End Function

End Class
