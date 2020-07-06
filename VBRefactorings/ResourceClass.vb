' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Option Explicit On
Option Infer Off
Option Strict On

Imports System.ComponentModel.Design
Imports System.IO
Imports VBRefactorings.EditorExtensions
Imports VSLangProj

Public Class ResourceXClass

    ''' <summary>
    ''' Class to speed up managing Resource files with option to add new resource
    ''' </summary>
    ''' <param name="CurrentFormFileDirectory"></param>
    Sub New(CurrentProjectFolderWithProjectFile As String)
        _ResourceFilenameWithPath = GetProjectGlobalizationFile(CurrentProjectFolderWithProjectFile)
        _ResourceFileNameWithoutExtension = Path.GetFileNameWithoutExtension(_ResourceFilenameWithPath)
        If String.IsNullOrWhiteSpace(_ResourceFilenameWithPath) Then
            Exit Sub
        End If
        _Initialized = InitializedValues.True
    End Sub

    Sub New()
        _Initialized = InitializedValues.Testing
        _ResourceFilenameWithPath = ""
        _ResourceFileNameWithoutExtension = ""
    End Sub

    <Flags>
    Public Enum InitializedValues
        [True]
        NoResourceDirectory
        Testing
    End Enum

    Private ReadOnly Property ResourceFilenameWithPath As String
    Public ReadOnly Property Initialized As InitializedValues = InitializedValues.NoResourceDirectory
    Public ReadOnly Property ResourceFileNameWithoutExtension As String
    Private Shared Function FindResourceFile(CultureFolderName As String, ProjectFileWithFullPath As String) As String
        Dim ProjectFilePath As String = Path.GetDirectoryName(ProjectFileWithFullPath)
        Dim FileList() As String = Directory.GetFiles(ProjectFilePath, $"Resource*.{CultureFolderName}.resx", SearchOption.TopDirectoryOnly)
        Select Case FileList.Length
            Case 0
                FileList = Directory.GetFiles(ProjectFilePath, $"Resource*.resx", SearchOption.TopDirectoryOnly)
                Select Case FileList.Count
                    Case 0
                        Return ""
                    Case 1
                        Return FileList(0)
                    Case Else
                        Stop
                End Select
            Case 1
                Return FileList(0)
            Case Else
                Stop
        End Select
        Return ""
    End Function

    Private Shared Function GetProjectGlobalizationFile(ProjectFileFullPath As String) As String
        Return FindResourceFile(Thread.CurrentThread.CurrentCulture.Name, ProjectFileFullPath)
    End Function

    Public Shared Function GetValidIdentifierName(BaseName As String, Unique As Boolean) As String
        If Unique Then
            Dim Hash As String = $"_{BaseName.GetHashCode.ToString(Globalization.CultureInfo.InvariantCulture).Replace("-", "", StringComparison.OrdinalIgnoreCase)}"
            Return BaseName.Truncate(1023 - Hash.Length) & Hash
        End If
        Dim ValidIdentifierName As String = "_"
        Dim CurrentCharacter As Integer = -1
        For i As Integer = 0 To BaseName.Count - 1
            Dim FirstChar As String = BaseName.Substring(i, 1)
            If SyntaxFacts.IsIdentifierStartCharacter(CChar(FirstChar)) AndAlso FirstChar <> "_" Then
                CurrentCharacter = i
                ValidIdentifierName = FirstChar.ToUpperInvariant
                Exit For
            End If
        Next
        Dim builder As New Text.StringBuilder()
        builder.Append(ValidIdentifierName)
        For i As Integer = CurrentCharacter + 1 To BaseName.Count - 1
            Dim CurrentChar As String = BaseName.Substring(i, 1)
            If SyntaxFacts.IsIdentifierPartCharacter(CChar(CurrentChar)) AndAlso CurrentChar <> "_" Then
                builder.Append(CurrentChar)
            End If
        Next
        Return builder.ToString().Truncate(1023)
    End Function

    ''' <summary>
    ''' Add a Resource to end of file with the option of adding Comments
    ''' </summary>
    ''' <param name="key"></param>
    ''' <param name="value"></param>
    ''' <param name="comment"></param>
    ''' <returns></returns>
    Public Function AddToResourceFile(ByVal key As String, ByVal value As String, Optional ByVal comment As String = "") As Boolean
        If Me.Initialized = InitializedValues.Testing Then
            Return True
        End If
        Using resourceWriter As New ResXResourceWriter(ResourceFilenameWithPath)
            'Get existing resources
            Using reader As New System.Resources.ResXResourceReader(ResourceFilenameWithPath) With {.UseResXDataNodes = True}
                For Each resEntry As DictionaryEntry In reader
                    Dim node As ResXDataNode = TryCast(resEntry.Value, ResXDataNode)
                    If node Is Nothing Then
                        Continue For
                    End If

                    If String.CompareOrdinal(key, node.Name) = 0 Then
                        If String.CompareOrdinal(value, node.GetValue(CType(Nothing, ITypeResolutionService)).ToString) = 0 Then
                            Return True
                        End If
                        ' Keep resources untouched. Alternatively modify this resource.
                        Return False
                    End If

                    resourceWriter.AddResource(node)
                Next resEntry
            End Using

            'Add new data (at the end of the file):
            resourceWriter.AddResource(New ResXDataNode(key, value) With {.comment = comment})
            'Write to file
            resourceWriter.Generate()
            resourceWriter.Close()
            Dim ResourceItem As VSProjectItem = CType(ProjectHelpers.DTE.Solution.FindProjectItem(ResourceFilenameWithPath).Object, VSProjectItem)
            ResourceItem.RunCustomTool()
        End Using
        Return True
    End Function

    Public Function GetResourceName(OriginalString As String) As String
        Dim ValidName As String = GetValidIdentifierName(OriginalString, False)
        If Me.Initialized = InitializedValues.Testing Then
            Return ValidName
        End If
        Using reader As New ResXResourceReader(ResourceFilenameWithPath) With {.UseResXDataNodes = True}
            For Each resEntry As DictionaryEntry In reader
                Dim node As ResXDataNode = TryCast(resEntry.Value, ResXDataNode)
                If node Is Nothing Then
                    Continue For
                End If

                If String.Equals(OriginalString, node.GetValue(CType(Nothing, ITypeResolutionService)).ToString, StringComparison.InvariantCulture) Then
                    Return node.Name
                End If
                If String.CompareOrdinal(ValidName, node.Name) = 0 Then
                    ' Found Name but string doesn't match so get a new name
                    Return GetValidIdentifierName(OriginalString, True)
                End If
            Next resEntry
        End Using
        Return ValidName
    End Function

    Public Function GetResourceNameFromValue(OriginalString As String) As String
        If Me.Initialized = InitializedValues.Testing Then
            Return OriginalString
        End If
        Dim ValidName As String = GetValidIdentifierName(OriginalString, False)
        Using reader As New ResXResourceReader(ResourceFilenameWithPath) With {.UseResXDataNodes = True}
            For Each resEntry As DictionaryEntry In reader
                Dim node As ResXDataNode = TryCast(resEntry.Value, ResXDataNode)
                If node Is Nothing Then
                    Continue For
                End If

                If OriginalString = node.GetValue(CType(Nothing, ITypeResolutionService)).ToString Then
                    Return node.Name
                End If
            Next resEntry
        End Using
        Return ""
    End Function

End Class
