' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Runtime.CompilerServices

Imports Microsoft.CodeAnalysis

Module Extensions
    <Flags>
    Friend Enum SymbolDisplayCompilerInternalOptions
        ''' <summary>
        ''' None
        ''' </summary>
        None = 0

        ''' <summary>
        ''' ".ctor" instead of "Goo"
        ''' </summary>
        UseMetadataMethodNames = 1 << 0

        ''' <summary>
        ''' "List`1" instead of "List&lt;T&gt;" ("List(of T)" in VB). Overrides GenericsOptions on
        ''' types.
        ''' </summary>
        UseArityForGenericTypes = 1 << 1

        ''' <summary>
        ''' Append "[Missing]" to missing Metadata types (for testing).
        ''' </summary>
        FlagMissingMetadataTypes = 1 << 2

        ''' <summary>
        ''' Include the Script type when qualifying type names.
        ''' </summary>
        IncludeScriptType = 1 << 3

        ''' <summary>
        ''' Include custom modifiers (e.g. modopt([mscorlib]System.Runtime.CompilerServices.IsConst)) on
        ''' the member (return) type and parameters.
        ''' </summary>
        ''' <remarks>
        ''' CONSIDER: custom modifiers are part of the public API, so we might want to move this to SymbolDisplayMemberOptions.
        ''' </remarks>
        IncludeCustomModifiers = 1 << 4

        ''' <summary>
        ''' For a type written as "int[][,]" in C#, then
        '''   a) setting this option will produce "int[,][]", and
        '''   b) not setting this option will produce "int[][,]".
        ''' </summary>
        ReverseArrayRankSpecifiers = 1 << 5
    End Enum

    ''' <summary>
    ''' A verbose format for displaying symbols (useful for testing).
    ''' </summary>
    ReadOnly TestFormat As New SymbolDisplayFormat(globalNamespaceStyle:=SymbolDisplayGlobalNamespaceStyle.OmittedAsContaining,
                                                   typeQualificationStyle:=SymbolDisplayTypeQualificationStyle.NameAndContainingTypesAndNamespaces,
                                                   propertyStyle:=SymbolDisplayPropertyStyle.ShowReadWriteDescriptor,
                                                   localOptions:=SymbolDisplayLocalOptions.IncludeType,
                                                   genericsOptions:=SymbolDisplayGenericsOptions.IncludeTypeParameters Or SymbolDisplayGenericsOptions.IncludeVariance,
                                                   memberOptions:=SymbolDisplayMemberOptions.IncludeParameters Or SymbolDisplayMemberOptions.IncludeContainingType Or SymbolDisplayMemberOptions.IncludeType Or SymbolDisplayMemberOptions.IncludeRef Or SymbolDisplayMemberOptions.IncludeExplicitInterface,
                                                   kindOptions:=SymbolDisplayKindOptions.IncludeMemberKeyword,
                                                   parameterOptions:=SymbolDisplayParameterOptions.IncludeOptionalBrackets Or SymbolDisplayParameterOptions.IncludeDefaultValue Or SymbolDisplayParameterOptions.IncludeParamsRefOut Or SymbolDisplayParameterOptions.IncludeExtensionThis Or SymbolDisplayParameterOptions.IncludeType Or SymbolDisplayParameterOptions.IncludeName,
                                                   miscellaneousOptions:=SymbolDisplayMiscellaneousOptions.EscapeKeywordIdentifiers Or SymbolDisplayMiscellaneousOptions.UseErrorTypeSymbolName)

    <Extension()>
    Function GetCode(xml As XElement) As String
        Dim code As String = xml.Value

        If code.First() = vbLf Then
            code = code.Remove(0, 1)
        End If

        If code.Last() = vbLf Then
            code = code.Remove(code.Length - 1)
        End If

        Return code.Replace(vbLf, vbCrLf)

    End Function

    <Extension>
    Public Function ToTestDisplayString(symbol As ISymbol) As String
        Return symbol.ToDisplayString(TestFormat)
    End Function

End Module
