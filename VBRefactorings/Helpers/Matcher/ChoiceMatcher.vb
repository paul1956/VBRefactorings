' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Partial Friend MustInherit Class Matcher(Of T)
    Private Class ChoiceMatcher
        Inherits Matcher(Of T)

        Private ReadOnly _matcher1 As Matcher(Of T)
        Private ReadOnly _matcher2 As Matcher(Of T)

        Public Sub New(ByVal matcher1 As Matcher(Of T), ByVal matcher2 As Matcher(Of T))
            _matcher1 = matcher1
            _matcher2 = matcher2
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format("({0}|{1})", _matcher1, _matcher2)
        End Function

        Public Overrides Function TryMatch(ByVal sequence As IList(Of T), ByRef index As Integer) As Boolean
            Return _matcher1.TryMatch(sequence, index) OrElse _matcher2.TryMatch(sequence, index)
        End Function

    End Class

End Class
