' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Partial Friend MustInherit Class Matcher(Of T)

    ' Tries to match this matcher against the provided sequence at the given index.  If the
    ' match succeeds, 'true' is returned, and 'index' points to the location after the match
    ' ends.  If the match fails, then false it returned and index remains the same.  Note: the
    ' matcher does not need to consume to the end of the sequence to succeed.
    Public MustOverride Function TryMatch(ByVal sequence As IList(Of T), ByRef index As Integer) As Boolean

    Friend Shared Function [Single](ByVal predicate As Func(Of T, Boolean), ByVal description As String) As Matcher(Of T)
        Return New SingleMatcher(predicate, description)
    End Function

    Friend Shared Function Choice(ByVal matcher1 As Matcher(Of T), ByVal matcher2 As Matcher(Of T)) As Matcher(Of T)
        Return New ChoiceMatcher(matcher1, matcher2)
    End Function

    Friend Shared Function OneOrMore(ByVal M As Matcher(Of T)) As Matcher(Of T)
        ' m+ is the same as (m m*)
        Return Sequence(M, Repeat(M))
    End Function

    Friend Shared Function Repeat(ByVal M As Matcher(Of T)) As Matcher(Of T)
        Return New RepeatMatcher(M)
    End Function

    Friend Shared Function Sequence(ParamArray ByVal matchers() As Matcher(Of T)) As Matcher(Of T)
        Return New SequenceMatcher(matchers)
    End Function

End Class
