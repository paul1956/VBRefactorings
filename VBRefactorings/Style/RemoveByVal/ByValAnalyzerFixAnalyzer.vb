' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Style

    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class ByValAnalyzerFixAnalyzer
        Inherits DiagnosticAnalyzer

        Public Const DiagnosticId As String = RemoveByValDiagnosticId

        ' You can change these strings in the Resources.resx file. If you do not want your analyzer to be localize-able, you can use regular strings for Title and MessageFormat.
        ' See https://github.com/dotnet/roslyn/blob/master/docs/analyzers/Localizing%20Analyzers.md for more on localization
        Private Shared ReadOnly Title As LocalizableString = New LocalizableResourceString(NameOf(My.Resources.Resources.RemoveByValTitle), My.Resources.Resources.ResourceManager, GetType(My.Resources.Resources))

        Private Shared ReadOnly MessageFormat As LocalizableString = New LocalizableResourceString(NameOf(My.Resources.Resources.RemoveByValMessageFormat), My.Resources.Resources.ResourceManager, GetType(My.Resources.Resources))
        Private Shared ReadOnly Description As LocalizableString = New LocalizableResourceString(NameOf(My.Resources.Resources.RemoveByValDescription), My.Resources.Resources.ResourceManager, GetType(My.Resources.Resources))
        Private Const Category As String = "Style"

        Private Shared ReadOnly Rule As New DiagnosticDescriptor(
                                                    DiagnosticId,
                                                    Title,
                                                    MessageFormat,
                                                    Category,
                                                    DiagnosticSeverity.Info,
                                                    isEnabledByDefault:=True,
                                                    Description,
                                                    helpLinkUri:="",
                                                    Array.Empty(Of String)
                                                    )

        Public Overrides ReadOnly Property SupportedDiagnostics As ImmutableArray(Of DiagnosticDescriptor)
            Get
                Return ImmutableArray.Create(Rule)
            End Get
        End Property

        Public Overrides Sub Initialize(context As AnalysisContext)
            ' TODO: Consider registering other actions that act on syntax instead of or in addition to symbols
            ' See https://github.com/dotnet/roslyn/blob/master/docs/analyzers/Analyzer%20Actions%20Semantics.md for more information
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf AnalyzeMethod, SyntaxKind.FunctionStatement, SyntaxKind.SubStatement)
        End Sub

        Private Sub AnalyzeMethod(ByVal context As SyntaxNodeAnalysisContext)
            Dim SubFunctionStatement As MethodStatementSyntax = CType(context.Node, MethodStatementSyntax)
            For i As Integer = 0 To SubFunctionStatement.ParameterList.Parameters.Count - 1
                Dim Parameter As ParameterSyntax = SubFunctionStatement.ParameterList.Parameters(i)
                For Each Modifier As SyntaxToken In Parameter.Modifiers
                    If Modifier.IsKind(SyntaxKind.ByValKeyword) Then
                        Dim diag As Diagnostic = Diagnostic.Create(Rule, Modifier.GetLocation, i)
                        context.ReportDiagnostic(diag)
                    End If
                Next
            Next
        End Sub

    End Class

End Namespace
