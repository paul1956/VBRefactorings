' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Resources
Imports System.Threading
Imports Microsoft.CodeAnalysis.VisualBasic

Public Class ResourceXClass

    Public Initialized As InitializedValues = InitializedValues.NoResourceDirectory

    Private Const GlobalizationFile As String = "Resources.resw"

    ''' <summary>
    ''' Class to speed up managing Resource files with option to add new resource
    ''' </summary>
    ''' <param name="CurrentProjectFolderWithProjectFile"></param>
    Sub New(CurrentProjectFolderWithProjectFile As String)
        ResourceFilenameWithPath = GetGlobalizationFileDirectory(CurrentProjectFolderWithProjectFile)
        If String.IsNullOrWhiteSpace(ResourceFilenameWithPath) Then
            Exit Sub
        End If
        Initialized = InitializedValues.True
    End Sub

    Sub New()
        Initialized = InitializedValues.Testing
    End Sub

    <Flags>
    Public Enum InitializedValues
        [True]
        NoResourceDirectory
        Testing
    End Enum

    Private Property ResourceFilenameWithPath As String = ""

    Public Shared Function GetValidIdentifierName(BaseName As String, Unique As Boolean) As String
        If Unique Then
            Dim Hash As String = $"_{BaseName.GetHashCode.ToString.Replace("-", "")}"
            Return BaseName.Truncate(1023 - Hash.Length) & Hash
        End If
        Dim ValidIdentifierName As String = "_"
        Dim CurrentCharacter As Integer = -1
        For i As Integer = 0 To BaseName.Count - 1
            If SyntaxFacts.IsIdentifierStartCharacter(CChar(BaseName.Substring(i, 1))) Then
                CurrentCharacter = i
                ValidIdentifierName = BaseName.Substring(i, 1)
                Exit For
            End If
        Next
        Dim builder As New Text.StringBuilder()
        builder.Append(ValidIdentifierName)
        For i As Integer = CurrentCharacter + 1 To BaseName.Count - 1
            If SyntaxFacts.IsIdentifierPartCharacter(CChar(BaseName.Substring(i, 1))) Then
                builder.Append(BaseName.Substring(i, 1))
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
            Using reader As New ResXResourceReader(ResourceFilenameWithPath) With {.UseResXDataNodes = True}
                For Each resEntry As DictionaryEntry In reader
                    Dim node As ResXDataNode = TryCast(resEntry.Value, ResXDataNode)
                    If node Is Nothing Then
                        Continue For
                    End If

                    If String.CompareOrdinal(key, node.Name) = 0 Then
                        ' Keep resources untouched. Alternatively modify this resource.
                        Return False
                    End If

                    resourceWriter.AddResource(node)
                Next resEntry
            End Using

            'Add new data (at the end of the file):
            resourceWriter.AddResource(New ResXDataNode(key, value) With {.Comment = comment})

            'Write to file
            resourceWriter.Generate()
        End Using
        Return True
    End Function

    Public Function GetUniqueResourceName(BaseName As String) As String
        If Me.Initialized = InitializedValues.Testing Then
            Return BaseName
        End If
        Dim PossibleName As String = GetValidIdentifierName(BaseName, False)
        Using reader As New ResXResourceReader(ResourceFilenameWithPath) With {.UseResXDataNodes = True}
            For Each resEntry As DictionaryEntry In reader
                Dim node As ResXDataNode = TryCast(resEntry.Value, ResXDataNode)
                If node Is Nothing Then
                    Continue For
                End If

                If String.CompareOrdinal(PossibleName, node.Name) = 0 Then
                    ' Found
                    Return GetValidIdentifierName(PossibleName, True)
                End If
            Next resEntry
        End Using
        Return BaseName
    End Function

    Private Shared Function FindLanguageResourceFile(CultureFolderName As String, RootDirectory As String) As String
        If Not RootDirectory.Contains("\") Then
            Return String.Empty
        End If
        Dim Directorylist() As String = IO.Directory.GetDirectories(RootDirectory, CultureFolderName, IO.SearchOption.AllDirectories)
        If Directorylist.Length = 0 Then
            Return ""
        End If
        For Each DirectoryName As String In Directorylist
            Dim FileList() As String = IO.Directory.GetFiles(DirectoryName, GlobalizationFile, IO.SearchOption.TopDirectoryOnly)
            Select Case FileList.Length
                Case 0
                Case 1
                    Return FileList(0)
                Case Else
                    Stop
            End Select
        Next
        Return ""
    End Function

    Private Shared Function GetGlobalizationFileDirectory(CurrentFileFolder As String) As String
        Dim CurrentRootDirectory As String = GetParentDIrectory(CurrentFileFolder)
        Return FindLanguageResourceFile(Thread.CurrentThread.CurrentCulture.Name, CurrentRootDirectory)
    End Function

    Private Shared Function GetParentDIrectory(Path As String) As String
        Dim FolderList() As String = Path.Split("\"c)
        FolderList.DropLastElement
        Dim RootDirector As String = FolderList.Join("\"c)
        Return RootDirector
    End Function

End Class
