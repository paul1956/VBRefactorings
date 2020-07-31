' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeRefactorings
Imports VBRefactorings.Style
Imports VBRefactoringsTests.Roslyn.UnitTestFramework
Imports Xunit

<TestClass()> Public Class ConvertStringToXMLLiteralUnitTests
    Inherits CodeRefactoringProviderTestFixture

    Protected Overrides ReadOnly Property LanguageName As String
        Get
            Return LanguageNames.VisualBasic
        End Get
    End Property

    Protected Overrides Function CreateCodeRefactoringProvider() As CodeRefactoringProvider
        Return New ConstStringToXMLLiteralRefactoring
    End Function

    <Fact()>
    Public Sub TestNoActionNotAStringLiteral()
        Const code As String = "Class Class1
    Sub Main()
        Const FixedlSource As String = ""[||]Class C
    End Class""
    End Sub
End Class"

        Const expected As String = "Class Class1
  Sub Main()
       Dim FixedlSource As String = <text>Class C
    End Class</text>.Value
    End Sub
End Class"
        Me.Test(code, expected)
    End Sub

End Class
