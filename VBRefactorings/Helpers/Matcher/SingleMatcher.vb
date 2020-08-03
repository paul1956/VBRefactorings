' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Partial Friend MustInherit Class Matcher(Of T)
    Private Class SingleMatcher
        Inherits Matcher(Of T)

        Private ReadOnly _description As String
        Private ReadOnly _predicate As Func(Of T, Boolean)

        Public Sub New(ByVal predicate As Func(Of T, Boolean), ByVal description As String)
            _predicate = predicate
            _description = description
        End Sub

        Public Overrides Function ToString() As String
            Return _description
        End Function

        Public Overrides Function TryMatch(ByVal sequence As IList(Of T), ByRef index As Integer) As Boolean
            If index < sequence.Count AndAlso _predicate(sequence(index)) Then
                index += 1
                Return True
            End If

            Return False
        End Function

    End Class
End Class
