' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic

Module CompilationTestUtils
    ' Use the mscorlib from our current process
    Public Function CreateCompilationWithMscorlibAndVBRuntime(Text As String, Optional FileName As String = "Test") As VisualBasicCompilation
        Dim SyntaxTree As SyntaxTree = SyntaxFactory.ParseSyntaxTree(Text)
        Return VisualBasicCompilation.Create(
                      FileName,
                      syntaxTrees:={SyntaxTree},
                      SharedReferences.VisualBasicReferences(""))
    End Function

End Module
