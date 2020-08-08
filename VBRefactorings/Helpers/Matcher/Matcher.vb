' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Friend MustInherit Class Matcher

    ''' <summary>
    ''' Matcher that matches an element if the provide predicate returns true.
    ''' </summary>
    Public Shared Function [Single](Of T)(predicate As Func(Of T, Boolean), description As String) As Matcher(Of T)
        Return Matcher(Of T).Single(predicate, description)
    End Function

    ''' <summary>
    ''' Matcher equivalent to (m_1|m_2)
    ''' </summary>
    Public Shared Function Choice(Of T)(matcher1 As Matcher(Of T), matcher2 As Matcher(Of T)) As Matcher(Of T)
        Return Matcher(Of T).Choice(matcher1, matcher2)
    End Function

    ''' <summary>
    ''' Matcher equivalent to (m+)
    ''' </summary>
    Public Shared Function OneOrMore(Of T)(M As Matcher(Of T)) As Matcher(Of T)
        Return Matcher(Of T).OneOrMore(M)
    End Function

    ''' <summary>
    ''' Matcher equivalent to (m*)
    ''' </summary>
    'INSTANT VB NOTE: The parameter matcher was renamed since it may cause conflicts with calls to static members of the user-defined type with this name:
    Public Shared Function Repeat(Of T)(matcher_Renamed As Matcher(Of T)) As Matcher(Of T)
        Return Matcher(Of T).Repeat(matcher_Renamed)
    End Function

    ''' <summary>
    ''' Matcher equivalent to (m_1 ... m_n)
    ''' </summary>
    Public Shared Function Sequence(Of T)(ParamArray matchers() As Matcher(Of T)) As Matcher(Of T)
        Return Matcher(Of T).Sequence(matchers)
    End Function

End Class
