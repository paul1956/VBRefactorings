' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeRefactorings
Imports UnitTestProject1.Roslyn.UnitTestFramework
Imports VBRefactorings.InternationalizationResources

Imports Xunit

Namespace UnitTest

    Public Class MoveStringToResourceUnitTests
        Inherits CodeRefactoringProviderTestFixture

        Protected Overrides ReadOnly Property LanguageName As String
            Get
                Return LanguageNames.VisualBasic
            End Get
        End Property

        Protected Overrides Function CreateCodeRefactoringProvider() As CodeRefactoringProvider
            Return New MoveStringToResourceFileRefactoringProvider
        End Function

        <Fact()>
        Public Sub TestNoActionNotAStringLiteral()
            Dim code As String = <text>
Class C
    Property [|P|] As String
        Get
    End Property
End Class</text>.Value
            TestNoActions(code)
        End Sub

        <Fact()>
        Public Sub TestRefactorinExpressionStatement()
            Dim code As String = <text>Class C
                    Sub Main(args As String())
                        x([|"Test"|])
                    End Sub
                    Sub x(ByRef P As String)
                       Dim x = P
                    End Sub
        End Class</text>.Value

            Dim expected As String = <text>Class C
                    Sub Main(args As String())
                        x(My.Resources..Test)
                    End Sub
                    Sub x(ByRef P As String)
                       Dim x = P
                    End Sub
        End Class</text>.Value
            Test(code, expected)
        End Sub

        <Fact()>
        Public Sub TestRefactorinFunctionOrSubWithOptionalParameter()
            Dim code As String = <text>Class C
            Function X(Optional P As String = [|"Test"|]) AS String
            Return P
            End Function
        End Class</text>.Value
            TestNoActions(code)
        End Sub

        <Fact()>
        Public Sub TestRefactoringInPropertyStatement()
            Dim code As String = <text>Structure S
            Property P As String = [|"Test"|]
        End Structure</text>.Value

            Dim expected As String = <text>Structure S
            Property P As String = My.Resources..Test
        End Structure</text>.Value
            Test(code, expected)
        End Sub

        <Fact()>
        Public Sub TestRefactoringOnAutoProperty()
            Dim code As String = <text>Class C
            Property P As String = [|"Test"|]
        End Class</text>.Value

            Dim expected As String = <text>Class C
            Property P As String = My.Resources..Test
        End Class</text>.Value
            Test(code, expected)
        End Sub

        <Fact()>
        Public Sub TestRefactoringOnField()
            Dim code As String = <text>Class C
            Private P As String = [|"Test"|]
        End Class</text>.Value

            Dim expected As String = <text>Class C
            Private P As String = My.Resources..Test
        End Class</text>.Value
            Test(code, expected)
        End Sub

        <Fact()>
        Public Sub TestRefactorinLocalDeclarationStatement()
            Dim code As String = <text>Class C
                    Sub Main(args As String())
                        Dim P as String = [|"Test"|]
                    End Sub
        End Class</text>.Value
            Dim expected As String = <text>Class C
                    Sub Main(args As String())
                        Dim P as String = My.Resources..Test
                    End Sub
        End Class</text>.Value
            Test(code, expected)
        End Sub

        <Fact()>
        Public Sub TestRefactorinMethodStatementWithOptionalParams()
            Dim code As String = <text>Class C
                Sub X(Optional P As String = [|"Test"|])
                    stop
                End Sub
            End Class</text>.Value
            TestNoActions(code)
        End Sub

        <Fact()>
        Public Sub TestRefactorinSimpleArgument()
            Dim code As String = <text>Class C
        Sub Main(args As String())
            x(P:= [|"Test"|])
        End Sub
        Sub x(A As String, Optional P As String = "")
            Dim x = P
        End Sub
    End Class
</text>.Value

            Dim expected As String = <text>Class C
        Sub Main(args As String())
            x(P:= My.Resources..Test)
        End Sub
        Sub x(A As String, Optional P As String = "")
            Dim x = P
        End Sub
    End Class
</text>.Value
            Test(code, expected)
        End Sub

    End Class

End Namespace
