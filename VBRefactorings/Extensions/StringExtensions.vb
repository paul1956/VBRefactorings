' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Runtime.CompilerServices
Imports System.Text

Public Module StringExtensions

    Private Enum TitleCaseState
        First_Character
        Upper_Case_Word
        Title_Case_Word
    End Enum

    Private Function Replace(original As String, pattern As String, replacement As String, comparisonType As StringComparison, stringBuilderInitialSize As Integer) As String
        If original Is Nothing Then
            Return Nothing
        End If

        If String.IsNullOrEmpty(pattern) Then
            Return original
        End If

        Dim posCurrent As Integer = 0
        Dim lenPattern As Integer = pattern.Length
        Dim idxNext As Integer = original.IndexOf(pattern, comparisonType)
        Dim result As New StringBuilder(If(stringBuilderInitialSize < 0, Math.Min(4096, original.Length), stringBuilderInitialSize))

        While idxNext >= 0
            result.Append(original, posCurrent, idxNext - posCurrent)
            result.Append(replacement)

            posCurrent = idxNext + lenPattern

            idxNext = original.IndexOf(pattern, posCurrent, comparisonType)
        End While

        result.Append(original, posCurrent, original.Length - posCurrent)

        Return result.ToString()
    End Function

    <Extension()>
    Public Function Contains(ByVal s As String, ByVal StringList() As String) As Boolean
        If StringList Is Nothing OrElse StringList.Length = 0 Then
            Return False
        End If
        For Each strTemp As String In StringList
            If s.IndexOf(strTemp, StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return True
            End If
        Next
        Return False
    End Function

    <Extension>
    Public Function ConvertTabToSpace(textSnippet As String, tabSize As Integer, initialColumn As Integer, endPosition As Integer) As Integer
        Contracts.Contract.Requires(tabSize > 0)
        Contracts.Contract.Requires(endPosition >= 0 AndAlso endPosition <= textSnippet.Length)

        Dim column As Integer = initialColumn

        ' now this will calculate indentation regardless of actual content on the buffer except TAB
        For i As Integer = 0 To endPosition - 1
            column += If(textSnippet(i) = vbTab, tabSize - column Mod tabSize, 1)
        Next

        Return column - initialColumn
    End Function

    <Extension()>
    Public Function Count(ByVal value As String, ByVal ch As Char) As Integer
        Return value.Count(Function(c As Char) c = ch)
    End Function

    <Extension()>
    Public Sub DropLastElement(Of T)(ByRef a() As T)
        a.RemoveAt(a.GetUpperBound(0))
    End Sub

    <Extension>
    Public Function Join(source As IEnumerable(Of String), separator As String) As String
        If source Is Nothing Then
            Throw New ArgumentNullException(NameOf(source))
        End If

        If separator Is Nothing Then
            Throw New ArgumentNullException(NameOf(separator))
        End If

        Return String.Join(separator, source)
    End Function

    <Extension()>
    Public Function Left(ByVal str As String, ByVal Length As Integer) As String
        Return str.Substring(0, Math.Min(Length, str.Length))
    End Function

    <Extension()>
    Public Function Mid(ByVal s As String, ByVal index As Integer, ByVal Length As Integer) As String
        Return s.Substring(index, Length)
    End Function

    ' Remove element at index "index". Result is one element shorter.
    ' Similar to List.RemoveAt, but for arrays.
    <Extension()>
    Public Sub RemoveAt(Of T)(ByRef a() As T, ByVal index As Integer)
        ' Move elements after "index" down 1 position.
        Array.Copy(a, index + 1, a, index, a.GetUpperBound(0) - index)
        ' Shorten by 1 element.
        ReDim Preserve a(a.GetUpperBound(0) - 1)
    End Sub

    <Extension()>
    Public Function Replace(original As String, pattern As String, replacement As String, comparisonType As StringComparison) As String
        Return Replace(original, pattern, replacement, comparisonType, -1)
    End Function

    <Extension()>
    Public Function Right(ByVal str As String, ByVal Length As Integer) As String
        Return str.Substring(Math.Max(str.Length, Length) - Length)
    End Function

    <Extension()>
    Public Function TitleCaseSplit(str As String, Optional delim As String = " ", Optional upper_case_indicator As String = "") As String
        Dim chr As String, out As String
        Dim state As TitleCaseState
        Dim is_upper As Boolean
        chr = Mid(str, 1, 1)
        out = chr
        state = TitleCaseState.First_Character
        For i As Integer = 2 To str.Length
            chr = Mid(str, i, 1)
            is_upper = Char.IsUpper(CChar(chr))
            Select Case state
                Case TitleCaseState.First_Character
                    state = If(is_upper, TitleCaseState.Upper_Case_Word, TitleCaseState.Title_Case_Word)
                    out = $"{out}{chr}"
                Case TitleCaseState.Title_Case_Word
                    If is_upper Then
                        state = TitleCaseState.First_Character
                        out = $"{out}{delim}{chr}"
                    Else
                        out = $"{out}{chr}"
                    End If
                Case TitleCaseState.Upper_Case_Word
                    If is_upper Then
                        out = $"{Left(out, out.Length - 1)}{upper_case_indicator}{Right(out, 1)}{chr}"
                    Else
                        state = TitleCaseState.Title_Case_Word
                        out = $"{Left(out, out.Length - 1)}{delim}{Right(out, 1)}{chr}"
                    End If
            End Select
        Next
        TitleCaseSplit = out
    End Function

    ''' <summary>
    ''' Produces a new string truncated to the length specified in MaxStringLenght
    ''' </summary>
    ''' <param name="StrIn">Original String</param>
    ''' <param name="MaxStringLenght">Desired length</param>
    ''' <returns>new string truncated to the length specified in MaxStringLenght</returns>
    ''' <remarks>MaxStringLenght must be >=0 </remarks>
    <Extension()>
    Public Function Truncate(ByRef StrIn As String, ByVal MaxStringLenght As Integer) As String
        Contracts.Contract.Requires(MaxStringLenght >= 0, "MaxStringLenght can't be less than 0")
        If String.IsNullOrWhiteSpace(StrIn) Then
            Return String.Empty
        End If
        If StrIn.Trim.Length <= MaxStringLenght Then
            Return StrIn
        End If
        Return StrIn.Trim.Substring(0, MaxStringLenght).Trim
    End Function

End Module
