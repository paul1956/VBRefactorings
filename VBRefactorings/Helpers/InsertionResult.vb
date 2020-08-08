' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeRefactorings

Public NotInheritable Class InsertionResult

    ''' <summary>
    ''' Gets the context the insertion is invoked at.
    ''' </summary>
    Private _privateContext As CodeRefactoringContext

    ''' <summary>
    ''' Gets the location of the type part the node should be inserted to.
    ''' </summary>
    Private _privateLocation As Location

    ''' <summary>
    ''' Gets the node that should be inserted.
    ''' </summary>
    Private _privateNode As SyntaxNode

    ''' <summary>
    ''' Gets the type the node should be inserted to.
    ''' </summary>
    Private _privateType As INamedTypeSymbol

    Public Sub New(context As CodeRefactoringContext, node As SyntaxNode, type As INamedTypeSymbol, location As Location)
        Me.Context = context
        Me.Node = node
        Me.Type = type
        Me.Location = location
    End Sub

    Public Property Context() As CodeRefactoringContext
        Get
            Return _privateContext
        End Get
        Private Set(value As CodeRefactoringContext)
            _privateContext = value
        End Set
    End Property

    Public Property Location() As Location
        Get
            Return _privateLocation
        End Get
        Private Set(value As Location)
            _privateLocation = value
        End Set
    End Property

    Public Property Node() As SyntaxNode
        Get
            Return _privateNode
        End Get
        Private Set(value As SyntaxNode)
            _privateNode = value
        End Set
    End Property

    Public Property Type() As INamedTypeSymbol
        Get
            Return _privateType
        End Get
        Private Set(value As INamedTypeSymbol)
            _privateType = value
        End Set
    End Property

    Public Shared Function GuessCorrectLocation(context As CodeRefactoringContext, locations As Immutable.ImmutableArray(Of Location)) As Location
        For Each Loc As Location In locations
            If context.Document.FilePath = Loc.SourceTree.FilePath Then
                Return Loc
            End If
        Next
        Return locations(0)
    End Function

End Class
