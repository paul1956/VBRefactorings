' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Threading
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeActions

Namespace Style

    Partial Public Class ByValAnalyzerFixCodeFixProvider

        Friend Class ByValAnalyzerFixCodeFixProviderCodeAction
            Inherits CodeAction

            Private ReadOnly _createChangedDocument As Func(Of Object, Task(Of Document))
            Private ReadOnly _title As String

            Public Sub New(title As String, createChangedDocument As Func(Of Object, Task(Of Document)))
                _title = title
                _createChangedDocument = createChangedDocument
            End Sub

            Public Overrides ReadOnly Property Title As String
                Get
                    Return _title
                End Get
            End Property

            Protected Overrides Function GetChangedDocumentAsync(CancelToken As CancellationToken) As Task(Of Document)
                Return _createChangedDocument(CancelToken)
            End Function

        End Class

    End Class

End Namespace
