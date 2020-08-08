' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Text

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports Xunit

Public Class SymbolsAndSemantics

    Private Shared Function GetMembers(parent As ISymbol) As IEnumerable(Of ISymbol)
        Dim container As INamespaceOrTypeSymbol = TryCast(parent, INamespaceOrTypeSymbol)
        If container IsNot Nothing Then
            Return container.GetMembers().AsEnumerable()
        End If

        Return Enumerable.Empty(Of ISymbol)()
    End Function

    Private Sub EnumSymbols(symbol As ISymbol, builder As StringBuilder)
        builder.AppendLine(symbol.ToString())

        For Each childSymbol As ISymbol In GetMembers(symbol)
            Me.EnumSymbols(childSymbol, builder)
        Next
    End Sub

    <Fact>
    Public Sub AnalyzeRegionControlFlow()
        Dim code As String =
<text>
Class C
    Public Sub F()
        Goto L1 ' 1

'start
        L1: Stop

        If False Then
            Return
        End If
'end
        Goto L1 ' 2
    End Sub
End Class
</text>.GetCode()

        Dim testCode As TestCodeContainer = New TestCodeContainer(code)

        Dim firstStatement As StatementSyntax = Nothing
        Dim lastStatement As StatementSyntax = Nothing
        testCode.GetStatementsBetweenMarkers(firstStatement, lastStatement)

        Dim controlFlowAnalysis1 As ControlFlowAnalysis = testCode.SemanticModel.AnalyzeControlFlow(firstStatement, lastStatement)
        Assert.Equal(1, controlFlowAnalysis1.EntryPoints.Length())
        Assert.Equal(1, controlFlowAnalysis1.ExitPoints.Length())
        Assert.True(controlFlowAnalysis1.EndPointIsReachable)

        Dim methodBlock As MethodBlockSyntax = testCode.SyntaxTree.GetRoot().DescendantNodes().OfType(Of MethodBlockSyntax)().First()
        Dim controlFlowAnalysis2 As ControlFlowAnalysis = testCode.SemanticModel.AnalyzeControlFlow(methodBlock.Statements.First, methodBlock.Statements.Last)
        Assert.False(controlFlowAnalysis2.EndPointIsReachable)
    End Sub

    <Fact>
    Public Sub AnalyzeRegionDataFlow()
        Dim code As String =
<text>
Class C
    Public Sub F(x As Integer)
        Dim a As Integer
'start
        Dim b As Integer
        Dim x As Integer
        Dim y As Integer = 1

        If True Then
            Dim z As String = "a"
        End If
'end
        Dim c As Integer
    End Sub
End Class
</text>.GetCode()

        Dim testCode As TestCodeContainer = New TestCodeContainer(code)

        Dim firstStatement As StatementSyntax = Nothing
        Dim lastStatement As StatementSyntax = Nothing
        testCode.GetStatementsBetweenMarkers(firstStatement, lastStatement)

        Dim dataFlowAnalysis As DataFlowAnalysis = testCode.SemanticModel.AnalyzeDataFlow(firstStatement, lastStatement)
        Assert.Equal("b,x,y,z", String.Join(",", dataFlowAnalysis.VariablesDeclared.Select(Function(s As ISymbol) s.Name)))
    End Sub

    <Fact>
    Public Sub BindNameToSymbol()
        Dim code As TestCodeContainer = New TestCodeContainer("Imports System")
        Dim compilationUnit As CompilationUnitSyntax = CType(code.SyntaxTree.GetRoot(), CompilationUnitSyntax)

        Dim name As NameSyntax = CType(compilationUnit.Imports(0).ImportsClauses.First(), SimpleImportsClauseSyntax).Name
        Assert.Equal("System", name.ToString())

        Dim nameInfo As SymbolInfo = code.SemanticModel.GetSymbolInfo(name)
        Dim nameSymbol As INamespaceSymbol = CType(nameInfo.Symbol, INamespaceSymbol)
        Assert.True(nameSymbol.GetNamespaceMembers().Any(Function(s As INamespaceSymbol) s.Name = "Collections"))
    End Sub

    <Fact>
    Public Sub EnumerateSymbolsInCompilation()
        Dim file1 As String =
<text>
Public Class Animal
    Public Overridable Sub MakeSound()
    End Sub
End Class
</text>.GetCode()

        Dim file2 As String =
<text>
Class Cat
    Inherits Animal

    Public Overrides Sub MakeSound()
    End Sub
End Class
</text>.GetCode()

        Dim comp As VisualBasicCompilation = VisualBasicCompilation.Create(
            "test",
            syntaxTrees:={SyntaxFactory.ParseSyntaxTree(file1), SyntaxFactory.ParseSyntaxTree(file2)},
            SharedReferences.VisualBasicReferences(""))

        Dim globalNamespace As INamespaceSymbol = comp.SourceModule.GlobalNamespace

        Dim builder As StringBuilder = New StringBuilder()
        Me.EnumSymbols(globalNamespace, builder)

        Dim expected As String = "Global" & vbCrLf &
                                 "My" & vbCrLf &
                                 "My.InternalXmlHelper" & vbCrLf &
                                 "Private Sub New()" & vbCrLf &
                                 "Public Shared Property Value(source As System.Collections.Generic.IEnumerable(Of System.Xml.Linq.XElement)) As String" & vbCrLf &
                                 "Public Shared Property Get Value(source As System.Collections.Generic.IEnumerable(Of System.Xml.Linq.XElement)) As String" & vbCrLf &
                                 "Public Shared Property Set Value(source As System.Collections.Generic.IEnumerable(Of System.Xml.Linq.XElement), value As String)" & vbCrLf &
                                 "Public Shared Property AttributeValue(source As System.Collections.Generic.IEnumerable(Of System.Xml.Linq.XElement), name As System.Xml.Linq.XName) As String" & vbCrLf &
                                 "Public Shared Property Get AttributeValue(source As System.Collections.Generic.IEnumerable(Of System.Xml.Linq.XElement), name As System.Xml.Linq.XName) As String" & vbCrLf &
                                 "Public Shared Property Set AttributeValue(source As System.Collections.Generic.IEnumerable(Of System.Xml.Linq.XElement), name As System.Xml.Linq.XName, value As String)" & vbCrLf &
                                 "Public Shared Property AttributeValue(source As System.Xml.Linq.XElement, name As System.Xml.Linq.XName) As String" & vbCrLf &
                                 "Public Shared Property Get AttributeValue(source As System.Xml.Linq.XElement, name As System.Xml.Linq.XName) As String" & vbCrLf &
                                 "Public Shared Property Set AttributeValue(source As System.Xml.Linq.XElement, name As System.Xml.Linq.XName, value As String)" & vbCrLf &
                                 "Public Shared Function CreateAttribute(name As System.Xml.Linq.XName, value As Object) As System.Xml.Linq.XAttribute" & vbCrLf &
                                 "Public Shared Function CreateNamespaceAttribute(name As System.Xml.Linq.XName, ns As System.Xml.Linq.XNamespace) As System.Xml.Linq.XAttribute" & vbCrLf &
                                 "Public Shared Function RemoveNamespaceAttributes(inScopePrefixes As String(), inScopeNs As System.Xml.Linq.XNamespace(), attributes As System.Collections.Generic.List(Of System.Xml.Linq.XAttribute), obj As Object) As Object" & vbCrLf &
                                 "Public Shared Function RemoveNamespaceAttributes(inScopePrefixes As String(), inScopeNs As System.Xml.Linq.XNamespace(), attributes As System.Collections.Generic.List(Of System.Xml.Linq.XAttribute), obj As System.Collections.IEnumerable) As System.Collections.IEnumerable" & vbCrLf &
                                 "My.InternalXmlHelper.RemoveNamespaceAttributesClosure" & vbCrLf &
                                 "Private ReadOnly m_inScopePrefixes As String()" & vbCrLf &
                                 "Private ReadOnly m_inScopeNs As System.Xml.Linq.XNamespace()" & vbCrLf &
                                 "Private ReadOnly m_attributes As System.Collections.Generic.List(Of System.Xml.Linq.XAttribute)" & vbCrLf &
                                 "Friend Sub New(inScopePrefixes As String(), inScopeNs As System.Xml.Linq.XNamespace(), attributes As System.Collections.Generic.List(Of System.Xml.Linq.XAttribute))" & vbCrLf &
                                 "Friend Function ProcessXElement(elem As System.Xml.Linq.XElement) As System.Xml.Linq.XElement" & vbCrLf &
                                 "Friend Function ProcessObject(obj As Object) As Object" & vbCrLf &
                                 "Public Shared Function RemoveNamespaceAttributes(inScopePrefixes As String(), inScopeNs As System.Xml.Linq.XNamespace(), attributes As System.Collections.Generic.List(Of System.Xml.Linq.XAttribute), e As System.Xml.Linq.XElement) As System.Xml.Linq.XElement" & vbCrLf &
                                 "Microsoft" & vbCrLf &
                                 "Microsoft.VisualBasic" & vbCrLf &
                                 "Microsoft.VisualBasic.Embedded" & vbCrLf &
                                 "Public Sub New()" & vbCrLf &
                                 "Animal" & vbCrLf &
                                 "Public Sub New()" & vbCrLf &
                                 "Public Overridable Sub MakeSound()" & vbCrLf &
                                 "Cat" & vbCrLf &
                                 "Public Sub New()" & vbCrLf &
                                 "Public Overrides Sub MakeSound()" & vbCrLf

        Assert.Equal(expected, builder.ToString())
    End Sub

    <Fact>
    Public Sub FailedOverloadResolution()
        Dim code As String =
<text>
Option Strict On
Module Module1
    Sub Main()
        X$.F("hello")
    End Sub
End Module
Module X
    Sub F()
    End Sub
    Sub F(i As Integer)
    End Sub
End Module
</text>.GetCode()

        Dim testCode As TestCodeContainer = New TestCodeContainer(code)

        Dim expression As ExpressionSyntax = CType(testCode.SyntaxNode, ExpressionSyntax)
        Dim typeInfo As TypeInfo = testCode.SemanticModel.GetTypeInfo(expression)
        Dim semanticInfo As SymbolInfo = testCode.SemanticModel.GetSymbolInfo(expression)

        Assert.Null(typeInfo.Type)
        Assert.Null(typeInfo.ConvertedType)
        Assert.Null(semanticInfo.Symbol)
        Assert.Equal(CandidateReason.OverloadResolutionFailure, semanticInfo.CandidateReason)
        Assert.Equal(1, semanticInfo.CandidateSymbols.Length)

        Assert.Equal("Public Sub F(i As Integer)", semanticInfo.CandidateSymbols(0).ToDisplayString())
        Assert.Equal(SymbolKind.Method, semanticInfo.CandidateSymbols(0).Kind)

        Dim memberGroup As Immutable.ImmutableArray(Of ISymbol) = testCode.SemanticModel.GetMemberGroup(expression)

        Assert.Equal(2, memberGroup.Length)

        Dim sortedMethodGroup As ISymbol() = Aggregate s In memberGroup.AsEnumerable()
                                Order By s.ToDisplayString()
                                Into ToArray()

        Assert.Equal("Public Sub F()", sortedMethodGroup(0).ToDisplayString())
        Assert.Equal("Public Sub F(i As Integer)", sortedMethodGroup(1).ToDisplayString())
    End Sub

    <Fact>
    Public Sub GetDeclaredSymbol()
        Dim code As String =
<text>
Namespace Acme
    Friend Class C$lass1
    End Class
End Namespace
</text>.GetCode()

        Dim testCode As TestCodeContainer = New TestCodeContainer(code)
        Dim symbol As INamedTypeSymbol = testCode.SemanticModel.GetDeclaredSymbol(CType(testCode.SyntaxNode, TypeStatementSyntax))

        Assert.Equal(True, symbol.CanBeReferencedByName)
        Assert.Equal("Acme", symbol.ContainingNamespace.Name)
        Assert.Equal(Microsoft.CodeAnalysis.Accessibility.Friend, symbol.DeclaredAccessibility)
        Assert.Equal(SymbolKind.NamedType, symbol.Kind)
        Assert.Equal("Class1", symbol.Name)
        Assert.Equal("Acme.Class1", symbol.ToDisplayString())
        Assert.Equal("Acme.Class1", symbol.ToString())
    End Sub

    <Fact>
    Public Sub GetExpressionType()
        Dim code As String =
<text>
Class C
    Public Shared Sub Method()
        Dim local As String = New C().ToString() &amp; String.Empty
    End Sub
End Class
</text>.GetCode()

        Dim testCode As TestCodeContainer = New TestCodeContainer(code)

        Dim localDeclaration As LocalDeclarationStatementSyntax = testCode.SyntaxTree.GetRoot().DescendantNodes().OfType(Of LocalDeclarationStatementSyntax).First()
        Dim initializer As ExpressionSyntax = localDeclaration.Declarators.First().Initializer.Value
        Dim semanticInfo As TypeInfo = testCode.SemanticModel.GetTypeInfo(initializer)
        Assert.Equal("String", semanticInfo.Type.Name)
    End Sub

    <Fact>
    Public Sub GetSymbolXmlDocComments()
        Dim code As String =
<text>
''' &lt;summary&gt;
''' This is a test class!
''' &lt;/summary&gt;
Class C$lass1
End Class
</text>.GetCode()

        Dim testCode As TestCodeContainer = New TestCodeContainer(code)

        Dim symbol As INamedTypeSymbol = testCode.SemanticModel.GetDeclaredSymbol(CType(testCode.SyntaxNode, TypeStatementSyntax))
        Dim actualXml As String = symbol.GetDocumentationCommentXml()
        Dim expectedXml As String = "<member name=""T:Class1""> <summary> This is a test class! </summary></member>"
        Assert.Equal(expectedXml, actualXml.Replace(vbCr, "").Replace(vbLf, ""))
    End Sub

    <Fact>
    Public Sub SemanticFactsTests()
        Dim code As String =
<text>
Class C1
    Sub M(i As Integer)
    End Sub
End Class
Class C2
    Sub M(i As Integer)
    End Sub
End Class
</text>.GetCode()

        Dim testCode As TestCodeContainer = New TestCodeContainer(code)

        Dim classNode1 As ClassStatementSyntax = CType(testCode.SyntaxTree.GetRoot().FindToken(testCode.Text.IndexOf("C1", StringComparison.OrdinalIgnoreCase)).Parent, ClassStatementSyntax)
        Dim classNode2 As ClassStatementSyntax = CType(testCode.SyntaxTree.GetRoot().FindToken(testCode.Text.IndexOf("C2", StringComparison.OrdinalIgnoreCase)).Parent, ClassStatementSyntax)

        Dim class1 As INamedTypeSymbol = testCode.SemanticModel.GetDeclaredSymbol(classNode1)
        Dim class2 As INamedTypeSymbol = testCode.SemanticModel.GetDeclaredSymbol(classNode2)

        Dim method1 As IMethodSymbol = CType(class1.GetMembers().First(), IMethodSymbol)
        Dim method2 As IMethodSymbol = CType(class2.GetMembers().First(), IMethodSymbol)

        ' TODO: this API has been made internal. What is the replacement? Do we even need it here?
        ' Assert.IsTrue(Symbol.HaveSameSignature(method1, method2))
    End Sub

    <Fact>
    Public Sub SymbolDisplayFormatTest()
        Dim code As String =
<text>
Class C1(Of T)
End Class
Class C2
    Public Shared Function M(Of TSource)(source as C1(Of TSource), index as Integer) As TSource
    End Function
End Class
</text>.GetCode()

        Dim testCode As TestCodeContainer = New TestCodeContainer(code)

        Dim displayFormat As SymbolDisplayFormat = New SymbolDisplayFormat(
            genericsOptions:=
                SymbolDisplayGenericsOptions.IncludeTypeParameters Or
                SymbolDisplayGenericsOptions.IncludeVariance,
            memberOptions:=
                SymbolDisplayMemberOptions.IncludeParameters Or
                SymbolDisplayMemberOptions.IncludeModifiers Or
                SymbolDisplayMemberOptions.IncludeAccessibility Or
                SymbolDisplayMemberOptions.IncludeType Or
                SymbolDisplayMemberOptions.IncludeContainingType,
            kindOptions:=
                SymbolDisplayKindOptions.IncludeMemberKeyword,
            parameterOptions:=
                SymbolDisplayParameterOptions.IncludeExtensionThis Or
                SymbolDisplayParameterOptions.IncludeType Or
                SymbolDisplayParameterOptions.IncludeName Or
                SymbolDisplayParameterOptions.IncludeDefaultValue,
            miscellaneousOptions:=
                SymbolDisplayMiscellaneousOptions.UseSpecialTypes)

        Dim symbol As ISymbol = testCode.Compilation.SourceModule.GlobalNamespace.GetTypeMembers("C2").First().GetMembers("M").First()
        Assert.Equal(
            "Public Shared Function C2.M(Of TSource)(source As C1(Of TSource), index As Integer) As TSource",
            symbol.ToDisplayString(displayFormat))
    End Sub

End Class
