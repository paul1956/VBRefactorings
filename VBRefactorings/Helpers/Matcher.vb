' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On

Friend MustInherit Class Matcher(Of T)

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

    Private Class RepeatMatcher
        Inherits Matcher(Of T)

        Private ReadOnly _matcher As Matcher(Of T)

        Public Sub New(ByVal M As Matcher(Of T))
            _matcher = M
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format("({0}*)", _matcher)
        End Function

        Public Overrides Function TryMatch(ByVal sequence As IList(Of T), ByRef index As Integer) As Boolean
            Do While _matcher.TryMatch(sequence, index)
            Loop

            Return True
        End Function

    End Class

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

Friend MustInherit Class Matcher

    ''' <summary>
    ''' Matcher that matches an element if the provide predicate returns true.
    ''' </summary>
    Public Shared Function [Single](Of T)(ByVal predicate As Func(Of T, Boolean), ByVal description As String) As Matcher(Of T)
        Return Matcher(Of T).Single(predicate, description)
    End Function

    ''' <summary>
    ''' Matcher equivalent to (m_1|m_2)
    ''' </summary>
    Public Shared Function Choice(Of T)(ByVal matcher1 As Matcher(Of T), ByVal matcher2 As Matcher(Of T)) As Matcher(Of T)
        Return Matcher(Of T).Choice(matcher1, matcher2)
    End Function

    ''' <summary>
    ''' Matcher equivalent to (m+)
    ''' </summary>
    Public Shared Function OneOrMore(Of T)(ByVal M As Matcher(Of T)) As Matcher(Of T)
        Return Matcher(Of T).OneOrMore(M)
    End Function

    ''' <summary>
    ''' Matcher equivalent to (m*)
    ''' </summary>
    'INSTANT VB NOTE: The parameter matcher was renamed since it may cause conflicts with calls to static members of the user-defined type with this name:
    Public Shared Function Repeat(Of T)(ByVal matcher_Renamed As Matcher(Of T)) As Matcher(Of T)
        Return Matcher(Of T).Repeat(matcher_Renamed)
    End Function

    ''' <summary>
    ''' Matcher equivalent to (m_1 ... m_n)
    ''' </summary>
    Public Shared Function Sequence(Of T)(ParamArray ByVal matchers() As Matcher(Of T)) As Matcher(Of T)
        Return Matcher(Of T).Sequence(matchers)
    End Function

End Class