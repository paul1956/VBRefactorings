' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Partial Friend MustInherit Class Matcher(Of T)
    Private Class RepeatMatcher
        Inherits Matcher(Of T)

        Private ReadOnly _matcher As Matcher(Of T)

        Public Sub New(M As Matcher(Of T))
            _matcher = M
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format("({0}*)", _matcher)
        End Function

        Public Overrides Function TryMatch(sequence As IList(Of T), ByRef index As Integer) As Boolean
            Do While _matcher.TryMatch(sequence, index)
            Loop

            Return True
        End Function

    End Class

End Class
