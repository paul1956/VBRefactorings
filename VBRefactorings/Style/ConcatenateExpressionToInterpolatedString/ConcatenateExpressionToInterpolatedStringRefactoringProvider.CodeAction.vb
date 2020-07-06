' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Threading
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeActions

Namespace Style

    Partial Friend Class ConcatenateExpressionToInterpolatedStringRefactoringProvider

        Private Class ConcatenateExpressionToInterpolatedStringCodeAction
            Inherits CodeAction

            Private ReadOnly _title As String
            Private ReadOnly _generateDocument As Func(Of CancellationToken, Task(Of Document))

            Public Sub New(title As String, generateDocument As Func(Of CancellationToken, Task(Of Document)))
                _title = title
                _generateDocument = generateDocument
            End Sub

            Public Overrides ReadOnly Property Title As String
                Get
                    Return _title
                End Get
            End Property

            Protected Overrides Function GetChangedDocumentAsync(CancelToken As CancellationToken) As Task(Of Document)
                Return _generateDocument(CancelToken)
            End Function

        End Class

    End Class

End Namespace
