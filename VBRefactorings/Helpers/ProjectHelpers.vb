' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On

Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Imports EnvDTE

Imports EnvDTE80
Imports Microsoft.Build.Evaluation
Imports Microsoft.VisualStudio.Shell
Imports Microsoft.VisualStudio.Text

Namespace EditorExtensions
    Public Module ProjectHelpers
        Private _DTE As DTE2
        Public Const SolutionItemsFolder As String = "Solution Items"

        Public Property DTE() As DTE2
            Get
                ThreadHelper.ThrowIfNotOnUIThread()

                If _DTE Is Nothing Then
                    _DTE = TryCast(ServiceProvider.GlobalProvider.GetService(GetType(DTE)), DTE2)
                End If

                Return _DTE
            End Get
            Set(value As DTE2)
                _DTE = value
            End Set
        End Property

        Private Function FixAbsolutePath(absolutePath As String) As String
            If String.IsNullOrWhiteSpace(absolutePath) Then
                Return absolutePath
            End If

            Dim uniformlySeparated As String = absolutePath.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar)
            Dim doubleSlash As New String(Path.DirectorySeparatorChar, 2)
            Dim prependSeparator As Boolean = uniformlySeparated.StartsWith(doubleSlash, StringComparison.Ordinal)
            uniformlySeparated = uniformlySeparated.Replace(doubleSlash, New String(Path.DirectorySeparatorChar, 1))

            If prependSeparator Then
                uniformlySeparated = Path.DirectorySeparatorChar + uniformlySeparated
            End If

            Return uniformlySeparated
        End Function

        Private Function GetChildProjects(parent As Project) As IEnumerable(Of Project)
            Try
                If parent.Kind <> ProjectKinds.vsProjectKindSolutionFolder AndAlso parent.Collection Is Nothing Then
                    ' Unloaded
                    Return Enumerable.Empty(Of Project)()
                End If

                If Not String.IsNullOrEmpty(parent.FullName) Then
                    Return New Project() {parent}
                End If
            Catch generatedExceptionName As COMException
                Return Enumerable.Empty(Of Project)()
            End Try

            Return parent.ProjectItems.Cast(Of ProjectItem)().
                Where(Function(p As ProjectItem) p.SubProject IsNot Nothing).
                SelectMany(Function(p As ProjectItem) GetChildProjects(p.SubProject))
        End Function

        '''<summary>Gets the directory containing the project for the specified file.</summary>
        Private Function GetProjectFolder(item As ProjectItem) As String
            If item Is Nothing OrElse item.ContainingProject Is Nothing OrElse item.ContainingProject.Collection Is Nothing OrElse String.IsNullOrEmpty(item.ContainingProject.FullName) Then
                ' Solution items
                Return Nothing
            End If

            Return GetRootFolder(item.ContainingProject)
        End Function

        Friend Function GetProjectItem(fileName As String) As ProjectItem
            Try
                Return DTE.Solution.FindProjectItem(fileName)
            Catch exception As Exception
                Stop
                Return Nothing
            End Try
        End Function

        Public Sub AddFileToActiveProject(fileName As String, Optional itemType As String = Nothing)
            Dim project As Project = GetActiveProject()

            If project Is Nothing Then
                Return
            End If
            Try
                Dim projectFilePath As String = project.Properties.Item("FullPath").Value.ToString()
                Dim projectDirPath As String = Path.GetDirectoryName(projectFilePath)

                If Not fileName.StartsWith(projectDirPath, StringComparison.OrdinalIgnoreCase) Then
                    Return
                End If

                Dim item As ProjectItem = project.ProjectItems.AddFromFile(fileName)

                If itemType Is Nothing OrElse item Is Nothing OrElse project.FullName.Contains("://") Then
                    Return
                End If

                item.Properties.Item(NameOf(itemType)).Value = itemType
            Finally
            End Try
        End Sub

        <Extension>
        Public Sub AddFileToProject(project As Project, fileName As String, Optional itemType As String = Nothing)
            If project Is Nothing Then
                Return
            End If
            Try
                Dim projectFilePath As String = project.Properties.Item("FullPath").Value.ToString()
                Dim projectDirPath As String = Path.GetDirectoryName(projectFilePath)

                If Not fileName.StartsWith(projectDirPath, StringComparison.OrdinalIgnoreCase) Then
                    Return
                End If

                Dim item As ProjectItem = project.ProjectItems.AddFromFile(fileName)

                If itemType Is Nothing OrElse item Is Nothing OrElse project.FullName.Contains("://") Then
                    Return
                End If

                item.Properties.Item(NameOf(itemType)).Value = itemType
            Finally
            End Try
        End Sub

        Public Function AddFileToProject(parentFileName As String, fileName As String) As ProjectItem
            If Path.GetFullPath(parentFileName) = Path.GetFullPath(fileName) OrElse Not File.Exists(fileName) Then
                Return Nothing
            End If

            fileName = Path.GetFullPath(fileName)
            ' WAP projects don't like paths with forward slashes
            Dim item As ProjectItem = GetProjectItem(parentFileName)

            If item Is Nothing OrElse item.ContainingProject Is Nothing OrElse String.IsNullOrEmpty(item.ContainingProject.FullName) Then
                Return Nothing
            End If

            Dim dependentItem As ProjectItem = GetProjectItem(fileName)

            If dependentItem IsNot Nothing AndAlso item.ContainingProject.[GetType]().Name = "OAProject" AndAlso item.ProjectItems IsNot Nothing Then
                ' WinJS
                Dim addedItem As ProjectItem

                Try
                    addedItem = dependentItem.ProjectItems.AddFromFile(fileName)

                    ' create nesting
                    If Path.GetDirectoryName(parentFileName) = Path.GetDirectoryName(fileName) Then
                        addedItem.Properties.Item("DependentUpon").Value = Path.GetFileName(parentFileName)
                    End If
                Catch generatedExceptionName As COMException
                    Return dependentItem
                End Try

                Return addedItem
            End If

            If dependentItem IsNot Nothing Then
                ' File already exists in the project
                Return Nothing
            ElseIf item.ContainingProject.[GetType]().Name <> "OAProject" AndAlso item.ProjectItems IsNot Nothing AndAlso Path.GetDirectoryName(parentFileName) = Path.GetDirectoryName(fileName) Then
                ' WAP
                Try
                    Return item.ProjectItems.AddFromFile(fileName)
                Finally
                End Try
            ElseIf Path.GetFullPath(fileName).StartsWith(GetRootFolder(item.ContainingProject), StringComparison.OrdinalIgnoreCase) Then
                ' Website
                Try
                    Return item.ContainingProject.ProjectItems.AddFromFile(fileName)
                Finally
                End Try
            End If

            Return Nothing
        End Function

        '''<summary>Attempts to ensure that a file is writable.</summary>
        ''' <returns>True if the file is not under source control or was checked out; false if the checkout failed or an error occurred.</returns>
        Public Function CheckOutFileFromSourceControl(fileName As String) As Boolean
            Try
                If DTE Is Nothing OrElse Not File.Exists(fileName) OrElse DTE.Solution.FindProjectItem(fileName) Is Nothing Then
                    Return True
                End If

                If DTE.SourceControl.IsItemUnderSCC(fileName) AndAlso Not DTE.SourceControl.IsItemCheckedOut(fileName) Then
                    Return DTE.SourceControl.CheckOutItem(fileName)
                End If

                Return True
            Catch ex As Exception When ex.HResult <> (New OperationCanceledException).HResult
                System.Diagnostics.Debugger.Launch()
                Stop
                Return False
            End Try
        End Function

        Public Function CreateDirectoryInProject(path__1 As String) As Boolean
            If Not path__1.StartsWith(GetRootFolder(), StringComparison.OrdinalIgnoreCase) Then
                Return False
            End If

            ' Assuming all the files would have extension
            ' and all the paths without extensions are directory.
            Dim directory__2 As String = If(String.IsNullOrEmpty(Path.GetExtension(path__1)), path__1, Path.GetDirectoryName(path__1))

            Directory.CreateDirectory(directory__2)

            Return True
        End Function

        Public Function GetAbsolutePathFromSettings(settingsPath As String, filePath As String) As String
            If String.IsNullOrEmpty(settingsPath) Then
                Return filePath
            End If

            Dim targetFileName As String = Path.GetFileName(filePath)
            Dim sourceDir As String = Path.GetDirectoryName(filePath)

            ' If the output path is not project-relative, combine it directly.
            If Not settingsPath.StartsWith("~/", StringComparison.OrdinalIgnoreCase) AndAlso Not settingsPath.StartsWith("/", StringComparison.OrdinalIgnoreCase) Then
                Return Path.GetFullPath(Path.Combine(sourceDir, settingsPath, targetFileName))
            End If

            Dim rootDir As String = ProjectHelpers.GetRootFolder()

            If String.IsNullOrEmpty(rootDir) Then
                ' If no project is loaded, assume relative to file anyway
                rootDir = sourceDir
            End If

            Return Path.GetFullPath(Path.Combine(rootDir, settingsPath.TrimStart("~"c, "/"c), targetFileName))
        End Function

        Public Function GetActiveFile() As ProjectItem
            Dim doc As Document = DTE.ActiveDocument

            If doc Is Nothing Then
                Return Nothing
            End If

            If GetProjectFolder(doc.ProjectItem) IsNot Nothing Then
                Return doc.ProjectItem
            End If

            Return Nothing
        End Function

        '''<summary>Gets the currently active project (as reported by the Solution Explorer), if any.</summary>
        Public Function GetActiveProject() As Project
            Try
                Dim activeSolutionProjects As Array = TryCast(DTE.ActiveSolutionProjects, Array)

                If activeSolutionProjects IsNot Nothing AndAlso activeSolutionProjects.Length > 0 Then
                    Return TryCast(activeSolutionProjects.GetValue(0), Project)
                End If
            Catch ex As Exception
                Stop
            End Try

            Return Nothing
        End Function

        Public Function GetAllProjects() As IEnumerable(Of Project)
            Return DTE.Solution.Projects.Cast(Of Project)().
                SelectMany(AddressOf GetChildProjects).
                Union(DTE.Solution.Projects.Cast(Of Project)()).
                Where(Function(p As Project)
                          Try
                              Return Not String.IsNullOrEmpty(p.FullName)
                          Catch ex As Exception
                              Return False
                          End Try
                      End Function)
        End Function

        Public Iterator Function GetBundleConstituentFiles(files As IEnumerable(Of String), root As String, folder As String, bundleFileName As String) As IEnumerable(Of String)
            For Each file As String In files
                Yield If(Path.IsPathRooted(file), ToAbsoluteFilePath(file, root, folder), ToAbsoluteFilePath(file, root, Path.GetDirectoryName(bundleFileName)))
            Next
        End Function

        '''<summary>Gets the Project containing the specified file.</summary>
        Public Function GetProject(item As String) As Project
            Dim projectItem As ProjectItem = GetProjectItem(item)

            If projectItem Is Nothing Then
                Return Nothing
            End If

            Return projectItem.ContainingProject
        End Function

        '''<summary>Gets the directory containing the project for the specified file.</summary>
        Public Function GetProjectFolder(fileNameOrFolder As String) As String
            If String.IsNullOrEmpty(fileNameOrFolder) Then
                Return GetRootFolder()
            End If

            Dim item As ProjectItem = GetProjectItem(fileNameOrFolder)
            Dim projectFolder As String = Nothing

            If item IsNot Nothing Then
                projectFolder = GetProjectFolder(item)
            End If

            Return projectFolder
        End Function

        '''<summary>Gets the base directory of a specific Project, or of the active project if no parameter is passed.</summary>
        Public Function GetRootFolder(Optional project As Project = Nothing) As String
            Try
                project = If(project, GetActiveProject())

                If project Is Nothing OrElse project.Collection Is Nothing Then
                    Dim doc As Document = DTE.ActiveDocument
                    If doc IsNot Nothing AndAlso Not String.IsNullOrEmpty(doc.FullName) Then
                        Return GetProjectFolder(doc.FullName)
                    End If
                    Return String.Empty
                End If
                If String.IsNullOrEmpty(project.FullName) Then
                    Return Nothing
                End If
                Dim fullPath As String
                Try
                    fullPath = TryCast(project.Properties.Item(NameOf(fullPath)).Value, String)
                Catch generatedExceptionName As ArgumentException
                    Try
                        ' MFC projects don't have FullPath, and there seems to be no way to query existence
                        fullPath = TryCast(project.Properties.Item("ProjectDirectory").Value, String)
                    Catch generatedExceptionName1 As ArgumentException
                        ' Installer projects have a ProjectPath.
                        fullPath = TryCast(project.Properties.Item("ProjectPath").Value, String)
                    End Try
                End Try

                If String.IsNullOrEmpty(fullPath) Then
                    Return If(File.Exists(project.FullName), Path.GetDirectoryName(project.FullName), "")
                End If

                If Directory.Exists(fullPath) Then
                    Return fullPath
                End If
                If File.Exists(fullPath) Then
                    Return Path.GetDirectoryName(fullPath)
                End If

                Return String.Empty
            Catch ex As Exception
                Return String.Empty
            End Try
        End Function

        '''<summary>Gets the paths to all files included in the selection, including files within selected folders.</summary>
        Public Function GetSelectedFilePaths() As IEnumerable(Of String)
            Return GetSelectedItemPaths().SelectMany(Function(p As String) If(Directory.Exists(p), Directory.EnumerateFiles(p, "*", SearchOption.AllDirectories), New String() {p}))
        End Function

        '''<summary>Gets the full paths to the currently selected item(s) in the Solution Explorer.</summary>
        Public Iterator Function GetSelectedItemPaths(Optional mDTE As DTE2 = Nothing) As IEnumerable(Of String)
            Dim items As Array = DirectCast(If(mDTE, DTE).ToolWindows.SolutionExplorer.SelectedItems, Array)
            For Each selItem As UIHierarchyItem In items
                Dim item As ProjectItem = TryCast(selItem.[Object], ProjectItem)

                If item IsNot Nothing AndAlso item.Properties IsNot Nothing Then
                    Yield item.Properties.Item("FullPath").Value.ToString()
                End If
            Next
        End Function

        '''<summary>Gets the currently selected file(s) in the Solution Explorer.</summary>
        Public Iterator Function GetSelectedItems() As IEnumerable(Of ProjectItem)
            Dim items As Array = DirectCast(DTE.ToolWindows.SolutionExplorer.SelectedItems, Array)
            For Each selItem As UIHierarchyItem In items
                Dim item As ProjectItem = TryCast(selItem.[Object], ProjectItem)

                If item IsNot Nothing Then
                    Yield item
                End If
            Next
        End Function

        '''<summary>Gets the currently selected project(s) in the Solution Explorer.</summary>
        Public Iterator Function GetSelectedProjects() As IEnumerable(Of Project)
            Dim items As Array = DirectCast(DTE.ToolWindows.SolutionExplorer.SelectedItems, Array)
            For Each selItem As UIHierarchyItem In items
                Dim item As Project = TryCast(selItem.[Object], Project)

                If item IsNot Nothing Then
                    Yield item
                End If
            Next
        End Function

        '''<summary>Gets the directory containing the active solution file.</summary>
        Public Function GetSolutionFolderPath() As String
            Dim solution As EnvDTE.Solution = DTE.Solution

            If solution Is Nothing Then
                Return Nothing
            End If

            If String.IsNullOrEmpty(solution.FullName) Then
                Return GetRootFolder()
            End If

            Return Path.GetDirectoryName(solution.FullName)
        End Function

        '''<summary>Gets the Solution Items solution folder in the current solution, creating it if it doesn't exist.</summary>
        Public Function GetSolutionItemsProject() As Project
            Dim solution As Solution2 = TryCast(DTE.Solution, Solution2)
            Return If(solution.Projects.OfType(Of Project)().FirstOrDefault(Function(p As Project) p.Name.Equals(SolutionItemsFolder, StringComparison.OrdinalIgnoreCase)), solution.AddSolutionFolder(SolutionItemsFolder))
        End Function

        '''<summary>Indicates whether a Project is a Web Application, Web Site, or WinJS project.</summary>
        <System.Runtime.CompilerServices.Extension>
        Public Function IsWebProject(project As Project) As Boolean
            ' Web site project
            If project.Kind.Equals("{E24C65DC-7377-472B-9ABA-BC803B73C61A}", StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
            Try
                Return project.Properties.Item("WebApplication.UseIISExpress") IsNot Nothing
            Finally
            End Try

            Return False
        End Function

#Region "ToAbsoluteFilePath()"

        '''<summary>Converts a relative URL to an absolute path on disk, as resolved from the specified file.</summary>
        Private Function ToAbsoluteFilePath(relativeUrl As String, file As ProjectItem) As String
            If file Is Nothing Then
                Return Nothing
            End If

            Dim baseFolder As String = If(file.Properties Is Nothing, Nothing, ProjectHelpers.GetProjectFolder(file))
            Return ToAbsoluteFilePath(relativeUrl, GetProjectFolder(file), baseFolder)
        End Function

        '''<summary>Converts a relative URL to an absolute path on disk, as resolved from the specified relative or base directory.</summary>
        '''<param name="relativeUrl">The URL to resolve.</param>
        '''<param name="projectRoot">The root directory to resolve absolute URLs from.</param>
        '''<param name="baseFolder">The source directory to resolve relative URLs from.</param>
        Private Function ToAbsoluteFilePath(relativeUrl As String, projectRoot As String, baseFolder As String) As String
            Dim imageUrl As String = relativeUrl.Trim({"'"c, """"c})
            Dim relUri As New Uri(imageUrl, UriKind.RelativeOrAbsolute)

            If relUri.IsAbsoluteUri Then
                Return relUri.LocalPath
            End If

            If relUri.OriginalString.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar).StartsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.OrdinalIgnoreCase) Then
                baseFolder = Nothing
                relUri = New Uri(relUri.OriginalString.Substring(1), UriKind.Relative)
            End If

            If projectRoot Is Nothing AndAlso baseFolder Is Nothing Then
                Return ""
            End If

            Dim root As String = (If(baseFolder, projectRoot)).Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar)

            If File.Exists(root) Then
                root = Path.GetDirectoryName(root)
            End If

            If Not root.EndsWith(New String(Path.DirectorySeparatorChar, 1), StringComparison.OrdinalIgnoreCase) Then
                root += Path.DirectorySeparatorChar
            End If

            Try
                Dim rootUri As New Uri(root, UriKind.Absolute)

                Return FixAbsolutePath(New Uri(rootUri, relUri).LocalPath)
            Catch generatedExceptionName As UriFormatException
                Return String.Empty
            End Try
        End Function

        '''<summary>Converts a relative URL to an absolute path on disk, as resolved from the specified file.</summary>
        Public Function ToAbsoluteFilePath(relativeUrl As String, relativeToFile As String) As String
            Dim file As ProjectItem = GetProjectItem(relativeToFile)
            If file Is Nothing OrElse file.Properties Is Nothing Then
                Return ToAbsoluteFilePath(relativeUrl, GetRootFolder(), Path.GetDirectoryName(relativeToFile))
            End If
            Return ToAbsoluteFilePath(relativeUrl, GetProjectFolder(file), Path.GetDirectoryName(relativeToFile))
        End Function

        '''<summary>Converts a relative URL to an absolute path on disk, as resolved from the active file.</summary>
        Public Function ToAbsoluteFilePathFromActiveFile(relativeUrl As String) As String
            Return ToAbsoluteFilePath(relativeUrl, GetActiveFile())
        End Function

#End Region

    End Module
End Namespace
