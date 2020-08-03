' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Partial Friend MustInherit Class Matcher(Of T)
    Private Class SequenceMatcher
        Inherits Matcher(Of T)

        Private ReadOnly _matchers() As Matcher(Of T)

        Public Sub New(ParamArray ByVal matchers() As Matcher(Of T))
            _matchers = matchers
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format("({0})", String.Join(",", CType(_matchers, Object())))
        End Function

        Public Overrides Function TryMatch(ByVal sequence As IList(Of T), ByRef index As Integer) As Boolean
            Dim currentIndex As Integer = index
            For Each M As Matcher(Of T) In _matchers
                If Not M.TryMatch(sequence, currentIndex) Then
                    Return False
                End If
            Next M

            index = currentIndex
            Return True
        End Function

    End Class
End Class
