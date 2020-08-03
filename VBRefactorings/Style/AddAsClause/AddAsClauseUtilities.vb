' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Text
Imports System.Threading
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Editing
Imports Microsoft.CodeAnalysis.Formatting
Imports Microsoft.CodeAnalysis.Simplification
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports Microsoft.CodeAnalysis.VisualBasic.SyntaxFactory
Imports VBRefactorings.Utilities

Namespace Style
    Public Module AddAsClauseUtilities
        Private Const DictionaryOf As String = "Dictionary(Of "
        Private Const KeyCollection As String = ").KeyCollection"
        Private Const MemoryLimit As Long = CType(Long.MaxValue * 0.75, Long)
        Private Const ValueCollection As String = ").ValueCollection"

        Private Function FixExpressionType(expressionType As ITypeSymbol) As TypeSyntax
            Dim ExpressionTypeString As String = expressionType.ToDisplayString
            If expressionType.IsTupleType Then
                Return ParseTypeName($" {ExpressionTypeString}")
            End If
            If String.IsNullOrEmpty(ExpressionTypeString) Then
                ExpressionTypeString = expressionType.ToString
            End If
            Dim Index As Integer = ExpressionTypeString.IndexOf(" As ", StringComparison.InvariantCulture)
            If Index > 0 Then
                ExpressionTypeString = ExpressionTypeString.Substring(0, Index).Trim & ")"
            End If
            If ExpressionTypeString = "Int32" Then
                Return ParseTypeName("Integer")
            End If
            Return ParseTypeName($" {ExpressionTypeString}")
        End Function

        Private Function GetAsClause(_forEachStatement As ForEachStatementSyntax, Model As SemanticModel, CancelToken As CancellationToken) As AsClauseSyntax
            Dim _ForEachStatementInfo As ForEachStatementInfo = Model.GetForEachStatementInfo(_forEachStatement)
            Dim ElementITypeSymbol As ITypeSymbol = _ForEachStatementInfo.ElementType
            If ElementITypeSymbol IsNot Nothing Then
                Return SimpleAsClause(FixExpressionType(ElementITypeSymbol).
                                        WithAdditionalAnnotations(Simplifier.Annotation)
                                        ).NormalizeWhitespace()
            End If
            Dim ForEachExpressionSyntax As ExpressionSyntax = _forEachStatement.Expression
            Dim ElementTypeInfo As TypeInfo = Model.GetTypeInfo(ForEachExpressionSyntax)

            If ElementTypeInfo.Type IsNot Nothing AndAlso ElementTypeInfo.Type.IsArrayType Then
                Dim ExpressionArrayTypeSymbol As IArrayTypeSymbol = TryCast(ElementTypeInfo.Type, IArrayTypeSymbol)
                If ExpressionArrayTypeSymbol Is Nothing Then
                    Stop
                    Return Nothing
                End If
                Return SimpleAsClause(
                                    FixExpressionType(ExpressionArrayTypeSymbol.ElementType).
                                    WithAdditionalAnnotations(Simplifier.Annotation)
                                    ).NormalizeWhitespace()
                Select Case ElementTypeInfo.Type.GetTypeArguments.Count
                    Case 0
                    Case 1
                        ElementITypeSymbol = ElementTypeInfo.Type.GetTypeArguments(0)

                        If ElementITypeSymbol IsNot Nothing Then
                            Return SimpleAsClause(
                                             FixExpressionType(ElementITypeSymbol).
                                             WithAdditionalAnnotations(Simplifier.Annotation)
                                             ).NormalizeWhitespace()
                        End If
                    Case Else
                End Select
            End If
            Dim ExpressionConvertedType As ITypeSymbol = ElementTypeInfo.ConvertedType
            If ExpressionConvertedType Is Nothing Then
                Return Nothing
            End If
            Select Case ExpressionConvertedType.TypeKind
                Case TypeKind.Array
                    Stop
                Case TypeKind.Class
                    Dim ExpressionConvertedTypetoString As String = ExpressionConvertedType.ToString

                    Dim Index As Integer = ExpressionConvertedTypetoString.IndexOf("(Of", StringComparison.InvariantCulture)
                    Dim PossibleDictionaryOfIndex As Integer = -1
                    If Index + 4 < ExpressionConvertedTypetoString.Length Then
                        PossibleDictionaryOfIndex = ExpressionConvertedTypetoString.Left(Index + 4).IndexOf(DictionaryOf, StringComparison.InvariantCultureIgnoreCase)
                    End If
                    If PossibleDictionaryOfIndex < 0 Then
                        Dim typeSyntax1 As TypeSyntax = GetTypeSyntaxFromInterface(ExpressionConvertedType)
                        If typeSyntax1 Is Nothing Then
                            Return Nothing
                        End If
                        Return SimpleAsClause(
                                                typeSyntax1.
                                                WithAdditionalAnnotations(Simplifier.Annotation)
                                                ).NormalizeWhitespace()
                    End If
                    ExpressionConvertedTypetoString = ExpressionConvertedTypetoString.Substring(PossibleDictionaryOfIndex + DictionaryOf.Length)
                    ' If we are hear we have a some kind of Dictionary
                    If ExpressionConvertedTypetoString.EndsWith(ValueCollection, StringComparison.InvariantCultureIgnoreCase) Then
                        ' if we are here we want the Values Collection
                        ExpressionConvertedTypetoString = ExpressionConvertedTypetoString.Replace(ValueCollection.Substring(1), String.Empty)
                        Dim DictionaryTypeCollection As List(Of String) = GetTypesExtracted(ExpressionConvertedTypetoString)
                        Return SimpleAsClause(
                                                ParseTypeName($" {DictionaryTypeCollection(1)}").
                                                WithAdditionalAnnotations(Simplifier.Annotation)
                                                ).NormalizeWhitespace()
                    ElseIf ExpressionConvertedTypetoString.EndsWith(KeyCollection, StringComparison.InvariantCultureIgnoreCase) Then
                        ' if we are here we want the Key Collection
                        ExpressionConvertedTypetoString = ExpressionConvertedTypetoString.Replace(KeyCollection.Substring(1), String.Empty)
                        Dim DictionaryTypeCollection As List(Of String) = GetTypesExtracted(ExpressionConvertedTypetoString)
                        Return SimpleAsClause(
                                                ParseTypeName($" {DictionaryTypeCollection(0)}").
                                                WithAdditionalAnnotations(Simplifier.Annotation)
                                                ).NormalizeWhitespace()
                    Else
                        If ExpressionConvertedTypetoString.Right(1) = ")" Then
                            ExpressionConvertedTypetoString = ExpressionConvertedTypetoString.Substring(0, ExpressionConvertedTypetoString.Length - 1)
                        End If
                        Return SimpleAsClause(
                                            ParseTypeName($" KeyValuePair(Of {ExpressionConvertedTypetoString})").
                                            WithAdditionalAnnotations(Simplifier.Annotation)
                                            ).NormalizeWhitespace()
                    End If
                Case TypeKind.Struct
                    Dim typeSyntax2 As TypeSyntax = GetTypeSyntaxFromInterface(ExpressionConvertedType)
                    If typeSyntax2 Is Nothing Then
                        Return Nothing
                    End If
                    Return SimpleAsClause(
                                            typeSyntax2.
                                            WithAdditionalAnnotations(Simplifier.Annotation)
                                            ).NormalizeWhitespace()
                Case TypeKind.Interface
                    Dim typeSyntax3 As TypeSyntax = GetTypeSyntaxFromInterface(ExpressionConvertedType)
                    If typeSyntax3 Is Nothing Then
                        Dim expressionTypeTuple As (pITypeSymbol As ITypeSymbol, _Error As Boolean) = ForEachExpressionSyntax.DetermineType(Model, CancelToken)
                        If expressionTypeTuple._Error Then
                            Return Nothing
                        End If
                        If expressionTypeTuple.pITypeSymbol IsNot Nothing AndAlso expressionTypeTuple.pITypeSymbol.IsArrayType Then
                            Dim ExpressionArrayTypeSymbol As IArrayTypeSymbol = TryCast(expressionTypeTuple.pITypeSymbol, IArrayTypeSymbol)
                            If ExpressionArrayTypeSymbol Is Nothing Then
                                Stop
                                Return Nothing
                            End If
                            Return SimpleAsClause(
                                    FixExpressionType(ExpressionArrayTypeSymbol.ElementType).
                                    WithAdditionalAnnotations(Simplifier.Annotation)
                                    ).NormalizeWhitespace()

                        End If
                        Return Nothing
                    End If
                    Return SimpleAsClause(
                                            typeSyntax3.
                                            WithAdditionalAnnotations(Simplifier.Annotation)
                                            ).NormalizeWhitespace()
                Case TypeKind.Error
                    Return Nothing
                Case Else
                    Stop
            End Select
            Return Nothing
        End Function

        Private Function GetAsClause(EnumSpecialHandling As Boolean, ExpressionValue As ExpressionSyntax, model As SemanticModel, variableITypeSymbol As ITypeSymbol) As AsClauseSyntax
            Dim FinalTypeSyntax As TypeSyntax = Nothing
            Try
                Select Case ExpressionValue.Kind
                    Case SyntaxKind.AddExpression,
                     SyntaxKind.AndExpression,
                     SyntaxKind.ArrayCreationExpression,
                     SyntaxKind.DictionaryAccessExpression,
                     SyntaxKind.ExclusiveOrExpression,
                     SyntaxKind.IdentifierName,
                     SyntaxKind.LeftShiftExpression,
                     SyntaxKind.ModuloExpression,
                     SyntaxKind.NotExpression,
                     SyntaxKind.NumericLiteralExpression,
                     SyntaxKind.OrExpression,
                     SyntaxKind.ParenthesizedExpression,
                     SyntaxKind.PredefinedCastExpression,
                     SyntaxKind.RightShiftExpression,
                     SyntaxKind.SubtractExpression,
                     SyntaxKind.UnaryMinusExpression,
                     SyntaxKind.XmlAttributeAccessExpression,
                     SyntaxKind.XmlCDataSection,
                     SyntaxKind.XmlDocument,
                     SyntaxKind.XmlElement,
                     SyntaxKind.XmlElementAccessExpression
                        FinalTypeSyntax = FixExpressionType(variableITypeSymbol)
                    Case SyntaxKind.CharacterLiteralExpression
                        FinalTypeSyntax = FixExpressionType(variableITypeSymbol) ' Char
                    Case SyntaxKind.CTypeExpression
                        FinalTypeSyntax = DirectCast(ExpressionValue, CTypeExpressionSyntax).Type
                    Case SyntaxKind.DirectCastExpression
                        FinalTypeSyntax = DirectCast(ExpressionValue, DirectCastExpressionSyntax).Type
                    Case SyntaxKind.TryCastExpression
                        FinalTypeSyntax = DirectCast(ExpressionValue, TryCastExpressionSyntax).Type
                    Case SyntaxKind.ObjectCreationExpression
                        Dim ExpressionString As String = ExpressionValue.ToString

                        Dim typ As TypeSyntax = TryCast(ExpressionValue, ObjectCreationExpressionSyntax).Type

                        Dim ClassName As String = ExpressionString.Trim.Replace("New ", "")
                        If variableITypeSymbol.ToString.Trim.EndsWith(ClassName, StringComparison.InvariantCultureIgnoreCase) OrElse
                            typ.ToString.EndsWith(ClassName, StringComparison.InvariantCultureIgnoreCase) Then
                            Dim NewExpression As NewExpressionSyntax = CType(ParseExpression(ExpressionString), NewExpressionSyntax)
                            Return AsNewClause(Token(SyntaxKind.AsKeyword), NewExpression).WithTriviaFrom(ExpressionValue).WithAdditionalAnnotations(Simplifier.Annotation).NormalizeWhitespace()
                        End If
                        FinalTypeSyntax = FixExpressionType(variableITypeSymbol).WithAdditionalAnnotations(Simplifier.Annotation)
                    Case SyntaxKind.ConditionalAccessExpression,
                     SyntaxKind.SimpleMemberAccessExpression,
                     SyntaxKind.TernaryConditionalExpression
                        If EnumSpecialHandling AndAlso variableITypeSymbol.IsEnumType Then
                            FinalTypeSyntax = ParseTypeName(GetNonErrorEnumUnderlyingType(variableITypeSymbol).ToString)
                        Else
                            FinalTypeSyntax = FixExpressionType(variableITypeSymbol)
                        End If
                    Case SyntaxKind.BinaryConditionalExpression
                        FinalTypeSyntax = FixExpressionType(variableITypeSymbol)
                    Case SyntaxKind.NothingLiteralExpression
                        FinalTypeSyntax = ParseTypeName($" [Object]")
                    Case SyntaxKind.QueryExpression
                        FinalTypeSyntax = FixExpressionType(variableITypeSymbol)
                    Case SyntaxKind.AndAlsoExpression,
                     SyntaxKind.EqualsExpression,
                     SyntaxKind.FalseLiteralExpression,
                     SyntaxKind.GreaterThanExpression,
                     SyntaxKind.GreaterThanOrEqualExpression,
                     SyntaxKind.IsNotExpression,
                     SyntaxKind.IsExpression,
                     SyntaxKind.LessThanExpression,
                     SyntaxKind.LessThanOrEqualExpression,
                     SyntaxKind.NotEqualsExpression,
                     SyntaxKind.OrElseExpression,
                     SyntaxKind.TrueLiteralExpression,
                     SyntaxKind.TypeOfIsExpression
                        Dim ConvertedType As String = variableITypeSymbol.ToDisplayString
                        If Not ConvertedType.TrimEnd("?"c).EndsWith("Boolean", StringComparison.InvariantCulture) Then
                            FinalTypeSyntax = ParseTypeName($" Boolean", StringComparison.InvariantCulture)
                        Else
                            FinalTypeSyntax = ParseTypeName($" {ConvertedType}")
                        End If
                    Case SyntaxKind.ConcatenateExpression,
                     SyntaxKind.InterpolatedStringExpression,
                     SyntaxKind.StringLiteralExpression
                        If Not variableITypeSymbol.ToDisplayString.EndsWith("String", StringComparison.InvariantCulture) Then
                            Stop
                        End If
                        FinalTypeSyntax = ParseTypeName($" String")
                    Case SyntaxKind.InvocationExpression
                        Dim ExpressionInfo As TypeInfo = model.GetTypeInfo(ExpressionValue)
                        If ExpressionInfo.Type Is Nothing Then
                            Return Nothing
                        End If
                        Dim ConvertedTypeString As String = variableITypeSymbol.ToMinimalDisplayString(model, 0, SymbolDisplayFormat.MinimallyQualifiedFormat)
                        If ConvertedTypeString.Contains({"anonymous type:", "<generated method>"}) Then
                            Return Nothing
                        End If
                        If ConvertedTypeString = "?" Then
                            FinalTypeSyntax = ParseTypeName($" Object")
                        Else

                            If ConvertedTypeString.Contains("(Of ") Then
                                Dim Keyword As String = "Optional"
                                For Each Keyword In SyntaxFacts.GetKeywordKinds()
                                    If ConvertedTypeString.Contains($" {Keyword}(Of") Then
                                        ConvertedTypeString = ConvertedTypeString.Replace($" {Keyword}(Of", " [{Keyword}](Of")
                                        Exit For
                                    End If
                                    If ConvertedTypeString.Contains(".{Keyword}(Of") Then
                                        ConvertedTypeString = ConvertedTypeString.Replace($".{Keyword}(Of", ".[{Keyword}](Of")
                                        Exit For
                                    End If
                                Next
                                FinalTypeSyntax = ParseTypeName($" {ConvertedTypeString} ")
                            Else
                                If String.IsNullOrWhiteSpace(variableITypeSymbol.ToDisplayString) Then
                                    FinalTypeSyntax = FixExpressionType(variableITypeSymbol)
                                Else
                                    FinalTypeSyntax = FixExpressionType(variableITypeSymbol)
                                End If
                            End If
                        End If
                    Case SyntaxKind.AwaitExpression
                        Dim _AwaitExpressionInfo As AwaitExpressionInfo = model.GetAwaitExpressionInfo(CType(ExpressionValue, AwaitExpressionSyntax), CancellationToken.None)
                        FinalTypeSyntax = ParseTypeName($" { _AwaitExpressionInfo.GetResultMethod.ReturnType}")
                    Case SyntaxKind.CollectionInitializer
                        FinalTypeSyntax = FixExpressionType(variableITypeSymbol)
                    Case SyntaxKind.MeExpression
                        FinalTypeSyntax = FixExpressionType(variableITypeSymbol)
                    Case SyntaxKind.FunctionLambdaHeader,
                    SyntaxKind.MultiLineFunctionLambdaExpression,
                    SyntaxKind.MultiLineSubLambdaExpression,
                    SyntaxKind.SingleLineFunctionLambdaExpression,
                    SyntaxKind.SingleLineSubLambdaExpression,
                    SyntaxKind.SubLambdaHeader
                        FinalTypeSyntax = FixExpressionType(variableITypeSymbol)
                    Case SyntaxKind.EqualsValue
                        FinalTypeSyntax = FixExpressionType(variableITypeSymbol)
                    Case SyntaxKind.TupleExpression
                        FinalTypeSyntax = ParseTypeName(variableITypeSymbol.ToString)
                    Case Else
                        Stop
                        FinalTypeSyntax = FixExpressionType(variableITypeSymbol)
                End Select
            Catch ex As Exception When ex.HResult <> (New OperationCanceledException).HResult
                If ex.HResult = (New OutOfMemoryException).HResult Then
                    Return Nothing
                End If
                Stop
            End Try
            If FinalTypeSyntax Is Nothing Then
                Return Nothing
            Else
                Debug.Print($"Leaving GetAsClause for {ExpressionValue} = {FinalTypeSyntax}")
                Return SimpleAsClause(FinalTypeSyntax.WithLeadingTrivia(ExpressionValue.GetLeadingTrivia()).WithAdditionalAnnotations(Simplifier.Annotation)
                                          ).NormalizeWhitespace()

            End If
        End Function

        Private Function GetNewInvocation(lVariableDeclarator As VariableDeclaratorSyntax, model As SemanticModel, ByRef variableITypeSymbol As ITypeSymbol) As VariableDeclaratorSyntax
            Try
                Dim lGetAsClause As AsClauseSyntax = GetAsClause(False, lVariableDeclarator.Initializer.Value, model, variableITypeSymbol)
                If lGetAsClause Is Nothing Then
                    Return lVariableDeclarator
                End If
                If lGetAsClause.ToString.Contains("As New ") Then
                    Return lVariableDeclarator.ReplaceNode(lVariableDeclarator, VariableDeclarator(lVariableDeclarator.Names, lGetAsClause, Nothing).WithAdditionalAnnotations(Simplifier.Annotation)).WithAdditionalAnnotations(Simplifier.Annotation)
                End If
                Return lVariableDeclarator.WithAsClause(lGetAsClause.WithTrailingTrivia(Whitespace(" "))).WithInitializer(lVariableDeclarator.Initializer)
            Catch ex As Exception When ex.HResult <> (New OperationCanceledException).HResult
                Stop
                Throw
            End Try
            Return lVariableDeclarator
        End Function

        Private Function GetNonErrorEnumUnderlyingType(type As ITypeSymbol) As INamedTypeSymbol
            Try
                If type.TypeKind = TypeKind.Enum Then
                    Dim underlying As INamedTypeSymbol = DirectCast(type, INamedTypeSymbol).EnumUnderlyingType

                    If underlying.Kind <> SymbolKind.ErrorType Then
                        Return underlying
                    End If
                End If
                Return DirectCast(type, INamedTypeSymbol)
            Catch ex As Exception
                Stop
            End Try
            Return Nothing
        End Function

        Private Function GetTypes(sometype As ITypeSymbol) As List(Of TypeSyntax)
            Dim ReturnTypeList As New List(Of TypeSyntax)
            Dim ExpectedHeader() As String = {NameOf(System), "."}
            Dim ExpectedMid() As String = {"(", "Of", " "}
            Dim Parts() As SymbolDisplayPart = sometype.ToDisplayParts.ToArray
            Dim Types As New List(Of String)
            Dim i As Integer
            For i = 0 To 1
                If ExpectedHeader(i) <> Parts(i).ToString Then
                    Return ReturnTypeList
                End If
            Next
            Dim FunctionOffset As Integer
            Select Case Parts(2).ToString
                Case "Func"
                    FunctionOffset = 1
                Case NameOf(Action)
                    FunctionOffset = 0
                Case Else
                    Return ReturnTypeList
            End Select

            For i = 3 To 5
                If ExpectedMid(i - 3) <> Parts(i).ToString Then
                    Return ReturnTypeList
                End If
            Next
            Dim TypeList As New StringBuilder
            For i = 6 To sometype.ToDisplayParts.Count - 2
                TypeList.Append(Parts(i))
            Next
            Dim TypeListString As String = TypeList.ToString.Trim
            Dim TypeListSplit As List(Of String)
            If TypeListString.IndexOf("Of", StringComparison.InvariantCulture) < 0 Then
                TypeListSplit = TypeListString.Split(","c).ToList
            Else
                TypeListSplit = GetTypesExtracted(TypeListString)
            End If
            For k As Integer = 0 To TypeListSplit.Count - 1 - FunctionOffset
                If TypeListSplit(k).Contains("(") And Not TypeListSplit(k).Contains(")") Then
                    Throw New FormatException($"GetTypes Type contains '(' but not ')' - {TypeListSplit(k)}")
                End If
                Select Case TypeListSplit(k).Count("("c) - TypeListSplit(k).Count(")"c)
                    Case 0
                    Case < 0
                        Throw New FormatException($"GetTypes Type contains to many ')' - {TypeListSplit(k)}")
                    Case > 0
                        Throw New FormatException($"GetTypes Type contains to many ')' - {TypeListSplit(k)}")
                End Select
                ReturnTypeList.Add(ParseTypeName(TypeListSplit(k).Trim))
            Next
            If ReturnTypeList.Count = 0 Then
                Stop
            End If
            Return ReturnTypeList
        End Function

        Private Function GetTypesExtracted(TypeListString As String) As List(Of String)
            Dim TypeListSplit As New List(Of String)
            Dim ModifiedTypeListStringBuilder As New StringBuilder
            For j As Integer = 0 To TypeListString.Count - 1
                Select Case TypeListString.Substring(j, 1)
                    Case ","
                        If ModifiedTypeListStringBuilder.Length > 0 Then
                            TypeListSplit.Add(ModifiedTypeListStringBuilder.ToString)
                            ModifiedTypeListStringBuilder.Clear()
                        End If
                    Case "("
                        Dim IndexOfCloseParen As Integer = TypeListString.Substring(j).IndexOf(")", StringComparison.InvariantCulture)
                        Dim IndexOfNextOpenParen As Integer = TypeListString.Substring(j + 1).IndexOf("(", StringComparison.InvariantCulture)
                        If IndexOfCloseParen > -1 Then
                            ModifiedTypeListStringBuilder.Append(TypeListString.Substring(j, IndexOfCloseParen + 1))
                            If IndexOfNextOpenParen < IndexOfCloseParen Then
                                j += IndexOfCloseParen
                                Continue For
                            End If
                        End If
                        j += IndexOfCloseParen
                        If j < TypeListString.Count - 1 Then
                            If TypeListString.Substring(j, 1) = "." Then
                                ModifiedTypeListStringBuilder.Append(".")
                            Else
                                TypeListSplit.Add(ModifiedTypeListStringBuilder.ToString)
                                ModifiedTypeListStringBuilder.Clear()
                            End If
                        End If
                    Case Else
                        ModifiedTypeListStringBuilder.Append(TypeListString.Substring(j, 1))
                End Select
            Next
            If ModifiedTypeListStringBuilder.Length > 0 Then
                TypeListSplit.Add(ModifiedTypeListStringBuilder.ToString.Trim)
                ModifiedTypeListStringBuilder.Clear()
            End If
            Return TypeListSplit
        End Function

        Private Function GetTypeSyntaxFromInterface(expressionConvertedType As ITypeSymbol) As TypeSyntax
            Const IEnumerableOf As String = "IEnumerable(Of "

            If Not expressionConvertedType.AllInterfaces.Any Then
                If expressionConvertedType.ToString.EndsWith("IArityEnumerable", StringComparison.InvariantCulture) Then
                    Return ParseName("Integer")
                End If
                Return Nothing
            End If
            For Each NamedType As INamedTypeSymbol In expressionConvertedType.AllInterfaces
                Dim index As Integer = NamedType.ToString.IndexOf(IEnumerableOf, StringComparison.InvariantCultureIgnoreCase)
                If index > 0 Then
                    Dim NewType As String = NamedType.ToString.Substring(index + IEnumerableOf.Length)
                    Return ParseName(NewType)
                End If
            Next

            Dim index1 As Integer = expressionConvertedType.ToString.IndexOf(IEnumerableOf, StringComparison.InvariantCultureIgnoreCase)
            If index1 > 0 Then
                Dim NewType As String = expressionConvertedType.ToString.Substring(index1 + IEnumerableOf.Length)
                Return ParseName(NewType)
            End If
            Return Nothing
        End Function

        Public Function AddAsClauseAsync(root As SyntaxNode, Model As SemanticModel, CurrentDocument As Document, OldVariableDeclarator As VariableDeclaratorSyntax) As SyntaxNode
            Contracts.Contract.Assume(CurrentDocument IsNot Nothing)
            If Not IsEnoughMemoryReport(MemoryLimit) Then
                Return OldVariableDeclarator
            End If
            Dim NewVariableDeclarator As VariableDeclaratorSyntax
            Try
                Dim oldNode As SyntaxNode = root.FindNode(OldVariableDeclarator.GetLocation.SourceSpan)
                NewVariableDeclarator = NewVariableDeclaratorSyntaxWithAsClauseAsync(Model, CurrentDocument, OldVariableDeclarator)
                If NewVariableDeclarator Is Nothing OrElse NewVariableDeclarator.Equals(OldVariableDeclarator) Then
                    Return OldVariableDeclarator
                End If
                Return NewVariableDeclarator
            Catch ex As Exception When ex.HResult <> (New OperationCanceledException).HResult
                Return OldVariableDeclarator
            Catch ex As Exception
                Throw
            End Try
        End Function

        Public Function AddAsClauseAsync(Model As SemanticModel, CurrentDocument As Document, LambdaHeader As LambdaHeaderSyntax) As SyntaxNode
            Contracts.Contract.Assume(CurrentDocument IsNot Nothing)
            Contracts.Contract.Assume(LambdaHeader IsNot Nothing)
            If Not IsEnoughMemoryReport(MemoryLimit) Then
                Return LambdaHeader
            End If
            Try
                Dim lLabdaExpression As LambdaExpressionSyntax = LambdaHeader.FirstAncestorOfType(Of LambdaExpressionSyntax)
                If lLabdaExpression Is Nothing Then
                    Return LambdaHeader
                End If
                Dim NewParameters As SeparatedSyntaxList(Of ParameterSyntax) = NewLambdaParameterListWithAsClause(Model, CurrentDocument, lLabdaExpression)
                If NewParameters.Count <> LambdaHeader.ParameterList.Parameters.Count Then
                    Return LambdaHeader
                End If
                Return LambdaHeader.WithParameterList(ParameterList(NewParameters).WithTriviaFrom(LambdaHeader.ParameterList))
            Catch ex As Exception When ex.HResult <> (New OperationCanceledException).HResult
                Return LambdaHeader
            Catch ex As Exception
                Throw
            End Try
        End Function

        Public Async Function AddAsClauseAsync(CurrentDocument As Document, ForStatement As ForStatementSyntax, CancelToken As CancellationToken) As Task(Of SyntaxNode)
            Contracts.Contract.Assume(CurrentDocument IsNot Nothing)
            Contracts.Contract.Assume(ForStatement IsNot Nothing)
            If Not IsEnoughMemoryReport(MemoryLimit) Then
                Return ForStatement
            End If
            Try
                Dim root As SyntaxNode = Await CurrentDocument.GetSyntaxRootAsync(CancelToken).ConfigureAwait(False)
                Dim newControlVariable As VariableDeclaratorSyntax = Await NewForControlVariableWithAsClause(CurrentDocument, ForStatement, CancelToken).ConfigureAwait(False)
                If newControlVariable Is Nothing Then
                    Return ForStatement
                End If
                Dim oldNode As SyntaxNode = root.FindNode(ForStatement.ControlVariable.GetLocation.SourceSpan)
                Return ForStatement.WithControlVariable(newControlVariable)
            Catch ex As Exception When ex.HResult <> (New OperationCanceledException).HResult
                Return ForStatement
            Catch ex As Exception
                Throw
            End Try
        End Function

        Public Async Function AddAsClauseAsync(CurrentDocument As Document, ForEachStatement As ForEachStatementSyntax, CancelToken As CancellationToken) As Task(Of SyntaxNode)
            Contracts.Contract.Assume(CurrentDocument IsNot Nothing)
            Contracts.Contract.Assume(ForEachStatement IsNot Nothing)
            If Not IsEnoughMemoryReport(MemoryLimit) Then
                Return ForEachStatement
            End If
            Try
                Dim root As SyntaxNode = Await CurrentDocument.GetSyntaxRootAsync(CancelToken).ConfigureAwait(False)
                Dim oldControlVariable As VisualBasicSyntaxNode = ForEachStatement.ControlVariable
                If oldControlVariable Is Nothing Then
                    Return ForEachStatement
                End If
                Dim newControlVariable As VariableDeclaratorSyntax = Await NewForEachControlVariableWithAsClause(CurrentDocument, ForEachStatement, CancelToken).ConfigureAwait(False)
                If newControlVariable Is Nothing Then
                    Return ForEachStatement
                End If
                Return ForEachStatement.WithControlVariable(newControlVariable)
            Catch ex As Exception When ex.HResult <> (New OperationCanceledException).HResult
                Return ForEachStatement
            Catch ex As Exception
                Throw
            End Try
        End Function

        Public Async Function AddAsClauseDocumentAsync(root As SyntaxNode, Model As SemanticModel, CurrentDocument As Document, OldVariableDeclarator As VariableDeclaratorSyntax, CancelToken As CancellationToken) As Task(Of Document)
            Dim editor As DocumentEditor = Await DocumentEditor.CreateAsync(CurrentDocument, CancelToken)
            editor.ReplaceNode(OldVariableDeclarator, AddAsClauseAsync(root, Model, CurrentDocument, OldVariableDeclarator))
            Return editor.GetChangedDocument
        End Function

        Public Async Function AddAsClauseDocumentAsync(CurrentDocument As Document, ForStatement As ForStatementSyntax, CancelToken As CancellationToken) As Task(Of Document)
            Dim editor As DocumentEditor = Await DocumentEditor.CreateAsync(CurrentDocument, CancelToken)
            editor.ReplaceNode(ForStatement, Await AddAsClauseAsync(CurrentDocument, ForStatement, CancelToken))
            Return editor.GetChangedDocument
        End Function

        Public Async Function AddAsClauseDocumentAsync(CurrentDocument As Document, ForEachStatement As ForEachStatementSyntax, CancelToken As CancellationToken) As Task(Of Document)
            Dim editor As DocumentEditor = Await DocumentEditor.CreateAsync(CurrentDocument, CancelToken)
            editor.ReplaceNode(ForEachStatement, Await AddAsClauseAsync(CurrentDocument, ForEachStatement, CancelToken))
            Return editor.GetChangedDocument
        End Function

        Public Async Function NewForControlVariableWithAsClause(CurrentDocument As Document, ForStatement As ForStatementSyntax, CancelToken As CancellationToken) As Task(Of VariableDeclaratorSyntax)
            Contracts.Contract.Assume(CurrentDocument IsNot Nothing)
            Contracts.Contract.Assume(ForStatement IsNot Nothing)
            Dim newControlVariable As VariableDeclaratorSyntax = Nothing
            Try

                Dim fromValue As ExpressionSyntax = ForStatement.FromValue
                Dim Model As SemanticModel = Await CurrentDocument.GetSemanticModelAsync(CancelToken).ConfigureAwait(False)
                Dim variableITypeSymbol As (pITypeSymbol As ITypeSymbol, _Error As Boolean) = fromValue.DetermineType(Model, CancelToken)
                If variableITypeSymbol._Error Then
                    Return Nothing
                End If

                Dim AsClause As AsClauseSyntax = GetAsClause(True, fromValue, Model, variableITypeSymbol.pITypeSymbol)
                If AsClause Is Nothing Then
                    Return Nothing
                End If
                Dim oldControlVariable As VisualBasicSyntaxNode = ForStatement.ControlVariable
                If oldControlVariable Is Nothing Then
                    Return Nothing
                End If
                If TypeOf oldControlVariable Is IdentifierNameSyntax Then
                    Return VariableDeclarator(ModifiedIdentifier(DirectCast(oldControlVariable, IdentifierNameSyntax).Identifier)).WithAsClause(AsClause).WithTriviaFrom(oldControlVariable)
                ElseIf TypeOf oldControlVariable Is VariableDeclaratorSyntax Then
                    Dim oldControlVariable1 As VariableDeclaratorSyntax = DirectCast(oldControlVariable, VariableDeclaratorSyntax)
                    If oldControlVariable1.Names.Count <> 1 Then
                        Return Nothing
                    End If
                    newControlVariable = VariableDeclarator(SingletonSeparatedList(
                                                                oldControlVariable1.Names(0)),
                                                                AsClause,
                                                                initializer:=Nothing
                                                                ).
                                                                WithTriviaFrom(oldControlVariable).WithAdditionalAnnotations(Simplifier.Annotation, Formatter.Annotation)
                Else
                    Stop
                End If
            Catch ex As Exception
                Stop
            End Try
            Return newControlVariable
        End Function

        Public Async Function NewForEachControlVariableWithAsClause(CurrentDocument As Document, ForEachStatement As ForEachStatementSyntax, CancelToken As CancellationToken) As Task(Of VariableDeclaratorSyntax)
            Contracts.Contract.Assume(CurrentDocument IsNot Nothing)
            Contracts.Contract.Assume(ForEachStatement IsNot Nothing)
            Dim NewControlVariable As VariableDeclaratorSyntax = Nothing
            Try
                Dim Model As SemanticModel = Await CurrentDocument.GetSemanticModelAsync(CancelToken).ConfigureAwait(False)
                Dim AsClause As AsClauseSyntax = GetAsClause(ForEachStatement, Model, CancelToken)
                If AsClause Is Nothing Then
                    Return Nothing
                End If
                Dim oldControlVariable As VisualBasicSyntaxNode = ForEachStatement.ControlVariable
                If oldControlVariable Is Nothing Then
                    Return NewControlVariable
                End If
                If TypeOf oldControlVariable Is IdentifierNameSyntax Then
                    Return VariableDeclarator(ModifiedIdentifier(DirectCast(oldControlVariable, IdentifierNameSyntax).Identifier)).WithAsClause(AsClause).
                        WithTriviaFrom(oldControlVariable)
                ElseIf TypeOf oldControlVariable Is VariableDeclaratorSyntax Then
                    Dim oldControlVariable1 As VariableDeclaratorSyntax = DirectCast(oldControlVariable, VariableDeclaratorSyntax)
                    If oldControlVariable1.Names.Count <> 1 Then
                        Return Nothing
                    End If
                    NewControlVariable = VariableDeclarator(SingletonSeparatedList(
                                                                oldControlVariable1.Names(0)),
                                                                AsClause,
                                                                initializer:=Nothing
                                                                ).
                                                                WithTriviaFrom(oldControlVariable).WithAdditionalAnnotations(Simplifier.Annotation, Formatter.Annotation)
                Else
                    Stop
                End If
            Catch ex As Exception
                Stop
            End Try
            Return NewControlVariable
        End Function

        Public Function NewLambdaParameterListWithAsClause(model As SemanticModel, CurrentDocument As Document, LambdaExpression As LambdaExpressionSyntax) As SeparatedSyntaxList(Of ParameterSyntax)
            Contracts.Contract.Assume(CurrentDocument IsNot Nothing)
            Contracts.Contract.Assume(LambdaExpression IsNot Nothing)
            Dim NewParameters As SeparatedSyntaxList(Of ParameterSyntax)
            Try
                Dim sometype As ITypeSymbol = LambdaExpression.DetermineType(model, CancellationToken.None)._ITypeSymbol
                Dim Types As List(Of TypeSyntax) = GetTypes(sometype)
                If Types.Count = 0 Then
                    Return NewParameters
                End If

                Dim Parameters As SeparatedSyntaxList(Of ParameterSyntax) = LambdaExpression.SubOrFunctionHeader.ParameterList.Parameters

                If Parameters.Count = 0 Then
                    Exit Try
                End If
                Dim i As Integer = 0
                For Each param As ParameterSyntax In Parameters
                    Dim NewSimpleAsClause As SimpleAsClauseSyntax = SimpleAsClause(
                                        Token(SyntaxKind.AsKeyword),
                                                                                   New SyntaxList(Of AttributeListSyntax),
                                                                                   Types(i).WithLeadingTrivia(Space).WithAdditionalAnnotations(Simplifier.Annotation))
                    NewParameters = NewParameters.Add(param.WithTrailingTrivia(Space).WithAsClause(NewSimpleAsClause))
                    i += 1
                Next
            Catch ex As Exception
                Stop
            End Try
            Return NewParameters
        End Function

        Public Function NewVariableDeclaratorSyntaxWithAsClauseAsync(model As SemanticModel, currentDocument As Document, variableDeclarator As VariableDeclaratorSyntax) As VariableDeclaratorSyntax
            Contracts.Contract.Assume(currentDocument IsNot Nothing)
            If variableDeclarator Is Nothing Then
                Throw New ArgumentException($"Argument {NameOf(variableDeclarator)} can't be nothing")
            End If
            Dim variableDeclaratorSyntax1 As VariableDeclaratorSyntax = Nothing
            Try
                Dim Expression As ExpressionSyntax = variableDeclarator.Initializer.Value
                Dim variableITypeSymbol As (pITypeSymbol As ITypeSymbol, _Error As Boolean) = Expression.DetermineType(model, CancellationToken.None)
                If variableITypeSymbol._Error Then
                    Return Nothing
                End If
                If variableITypeSymbol.ToString.Contains(" ?)") Then
                    Return Nothing
                End If
                variableDeclaratorSyntax1 = GetNewInvocation(variableDeclarator, model, variableITypeSymbol.pITypeSymbol).WithTriviaFrom(variableDeclarator)
            Catch ex As Exception
                Stop
            End Try
            Return variableDeclaratorSyntax1
        End Function

    End Module
End Namespace
