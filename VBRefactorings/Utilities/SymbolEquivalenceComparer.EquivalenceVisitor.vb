' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.
'

Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Shared.Extensions

Namespace Utilities

    Partial Friend Class SymbolEquivalenceComparer

        Private Class EquivalenceVisitor
            Private ReadOnly _compareMethodTypeParametersByIndex As Boolean
            Private ReadOnly _objectAndDynamicCompareEqually As Boolean
            Private ReadOnly _symbolEquivalenceComparer As SymbolEquivalenceComparer

            Public Sub New(ByVal symbolEquivalenceComparer As SymbolEquivalenceComparer, ByVal compareMethodTypeParametersByIndex As Boolean, ByVal objectAndDynamicCompareEqually As Boolean)
                _symbolEquivalenceComparer = symbolEquivalenceComparer
                _compareMethodTypeParametersByIndex = compareMethodTypeParametersByIndex
                _objectAndDynamicCompareEqually = objectAndDynamicCompareEqually
            End Sub

            Private Shared Function AreCompatibleMethodKinds(ByVal kind1 As MethodKind, ByVal kind2 As MethodKind) As Boolean
                If kind1 = kind2 Then
                    Return True
                End If

                If (kind1 = MethodKind.Ordinary AndAlso kind2.IsPropertyAccessor()) OrElse (kind1.IsPropertyAccessor() AndAlso kind2 = MethodKind.Ordinary) Then
                    Return True
                End If

                Return False
            End Function

            Private Shared Function DynamicTypesAreEquivalent(ByVal _1 As IDynamicTypeSymbol, ByVal _2 As IDynamicTypeSymbol) As Boolean
                Return True
            End Function

            Private Shared Function HaveSameLocation(ByVal x As ISymbol, ByVal y As ISymbol) As Boolean
                Return x.Locations.Length = 1 AndAlso y.Locations.Length = 1 AndAlso x.Locations.First().Equals(y.Locations.First())
            End Function

            Private Shared Function LabelsAreEquivalent(ByVal x As ILabelSymbol, ByVal y As ILabelSymbol) As Boolean
                Return x.Name = y.Name AndAlso HaveSameLocation(x, y)
            End Function

            Private Shared Function LocalsAreEquivalent(ByVal x As ILocalSymbol, ByVal y As ILocalSymbol) As Boolean
                Return HaveSameLocation(x, y)
            End Function

            Private Shared Function PreprocessingSymbolsAreEquivalent(ByVal x As IPreprocessingSymbol, ByVal y As IPreprocessingSymbol) As Boolean
                Return x.Name = y.Name
            End Function

            Private Shared Function RangeVariablesAreEquivalent(ByVal x As IRangeVariableSymbol, ByVal y As IRangeVariableSymbol) As Boolean
                Return HaveSameLocation(x, y)
            End Function

            Private Function AreEquivalentWorker(ByVal x As ISymbol, ByVal y As ISymbol, ByVal k As SymbolKind, ByVal equivalentTypesWithDifferingAssemblies As Dictionary(Of INamedTypeSymbol, INamedTypeSymbol)) As Boolean
                Contracts.Contract.Requires(x.Kind = y.Kind AndAlso x.Kind = k)
                Select Case k
                    Case SymbolKind.ArrayType
                        Return Me.ArrayTypesAreEquivalent(DirectCast(x, IArrayTypeSymbol), DirectCast(y, IArrayTypeSymbol), equivalentTypesWithDifferingAssemblies)
                    Case SymbolKind.Assembly
                        Return Me.AssembliesAreEquivalent(DirectCast(x, IAssemblySymbol), DirectCast(y, IAssemblySymbol))
                    Case SymbolKind.DynamicType
                        Return DynamicTypesAreEquivalent(DirectCast(x, IDynamicTypeSymbol), DirectCast(y, IDynamicTypeSymbol))
                    Case SymbolKind.Event
                        Return Me.EventsAreEquivalent(DirectCast(x, IEventSymbol), DirectCast(y, IEventSymbol), equivalentTypesWithDifferingAssemblies)
                    Case SymbolKind.Field
                        Return Me.FieldsAreEquivalent(DirectCast(x, IFieldSymbol), DirectCast(y, IFieldSymbol), equivalentTypesWithDifferingAssemblies)
                    Case SymbolKind.Label
                        Return LabelsAreEquivalent(DirectCast(x, ILabelSymbol), DirectCast(y, ILabelSymbol))
                    Case SymbolKind.Local
                        Return LocalsAreEquivalent(DirectCast(x, ILocalSymbol), DirectCast(y, ILocalSymbol))
                    Case SymbolKind.Method
                        Return Me.MethodsAreEquivalent(DirectCast(x, IMethodSymbol), DirectCast(y, IMethodSymbol), equivalentTypesWithDifferingAssemblies)
                    Case SymbolKind.NetModule
                        Return Me.ModulesAreEquivalent(DirectCast(x, IModuleSymbol), DirectCast(y, IModuleSymbol))
                    Case SymbolKind.NamedType, SymbolKind.ErrorType
                        Return Me.NamedTypesAreEquivalent(DirectCast(x, INamedTypeSymbol), DirectCast(y, INamedTypeSymbol), equivalentTypesWithDifferingAssemblies)
                    Case SymbolKind.Namespace
                        Return Me.NamespacesAreEquivalent(DirectCast(x, INamespaceSymbol), DirectCast(y, INamespaceSymbol), equivalentTypesWithDifferingAssemblies)
                    Case SymbolKind.Parameter
                        Return Me.ParametersAreEquivalent(DirectCast(x, IParameterSymbol), DirectCast(y, IParameterSymbol), equivalentTypesWithDifferingAssemblies)
                    Case SymbolKind.PointerType
                        Return Me.PointerTypesAreEquivalent(DirectCast(x, IPointerTypeSymbol), DirectCast(y, IPointerTypeSymbol), equivalentTypesWithDifferingAssemblies)
                    Case SymbolKind.Property
                        Return Me.PropertiesAreEquivalent(DirectCast(x, IPropertySymbol), DirectCast(y, IPropertySymbol), equivalentTypesWithDifferingAssemblies)
                    Case SymbolKind.RangeVariable
                        Return RangeVariablesAreEquivalent(DirectCast(x, IRangeVariableSymbol), DirectCast(y, IRangeVariableSymbol))
                    Case SymbolKind.TypeParameter
                        Return Me.TypeParametersAreEquivalent(DirectCast(x, ITypeParameterSymbol), DirectCast(y, ITypeParameterSymbol), equivalentTypesWithDifferingAssemblies)
                    Case SymbolKind.Preprocessing
                        Return PreprocessingSymbolsAreEquivalent(DirectCast(x, IPreprocessingSymbol), DirectCast(y, IPreprocessingSymbol))
                    Case Else
                        Return False
                End Select
            End Function

            Private Function ArrayTypesAreEquivalent(ByVal x As IArrayTypeSymbol, ByVal y As IArrayTypeSymbol, ByVal equivalentTypesWithDifferingAssemblies As Dictionary(Of INamedTypeSymbol, INamedTypeSymbol)) As Boolean
                Return x.Rank = y.Rank AndAlso Me.AreEquivalent(x.CustomModifiers, y.CustomModifiers, equivalentTypesWithDifferingAssemblies) AndAlso Me.AreEquivalent(x.ElementType, y.ElementType, equivalentTypesWithDifferingAssemblies)
            End Function

            Private Function AssembliesAreEquivalent(ByVal x As IAssemblySymbol, ByVal y As IAssemblySymbol) As Boolean
                Return If(_symbolEquivalenceComparer._assemblyComparerOpt?.Equals(x, y), True)
            End Function

            Private Function EventsAreEquivalent(ByVal x As IEventSymbol, ByVal y As IEventSymbol, ByVal equivalentTypesWithDifferingAssemblies As Dictionary(Of INamedTypeSymbol, INamedTypeSymbol)) As Boolean
                Return x.Name = y.Name AndAlso Me.AreEquivalent(x.ContainingSymbol, y.ContainingSymbol, equivalentTypesWithDifferingAssemblies)
            End Function

            Private Function FieldsAreEquivalent(ByVal x As IFieldSymbol, ByVal y As IFieldSymbol, ByVal equivalentTypesWithDifferingAssemblies As Dictionary(Of INamedTypeSymbol, INamedTypeSymbol)) As Boolean
                Return x.Name = y.Name AndAlso Me.AreEquivalent(x.CustomModifiers, y.CustomModifiers, equivalentTypesWithDifferingAssemblies) AndAlso Me.AreEquivalent(x.ContainingSymbol, y.ContainingSymbol, equivalentTypesWithDifferingAssemblies)
            End Function

            Private Function HandleAnonymousTypes(ByVal x As INamedTypeSymbol, ByVal y As INamedTypeSymbol, ByVal equivalentTypesWithDifferingAssemblies As Dictionary(Of INamedTypeSymbol, INamedTypeSymbol)) As Boolean
                If x.TypeKind = TypeKind.Delegate Then
                    Return Me.AreEquivalent(x.DelegateInvokeMethod, y.DelegateInvokeMethod, equivalentTypesWithDifferingAssemblies)
                Else
                    Dim xMembers As IEnumerable(Of IPropertySymbol) = x.GetValidAnonymousTypeProperties()
                    Dim yMembers As IEnumerable(Of IPropertySymbol) = y.GetValidAnonymousTypeProperties()

                    Dim xMembersEnumerator As IEnumerator(Of IPropertySymbol) = xMembers.GetEnumerator()
                    Dim yMembersEnumerator As IEnumerator(Of IPropertySymbol) = yMembers.GetEnumerator()

                    Do While xMembersEnumerator.MoveNext()
                        If Not yMembersEnumerator.MoveNext() Then
                            Return False
                        End If

                        Dim p1 As IPropertySymbol = xMembersEnumerator.Current
                        Dim p2 As IPropertySymbol = yMembersEnumerator.Current

                        If p1.Name <> p2.Name OrElse p1.IsReadOnly <> p2.IsReadOnly OrElse Not Me.AreEquivalent(p1.Type, p2.Type, equivalentTypesWithDifferingAssemblies) Then
                            Return False
                        End If
                    Loop

                    Return Not yMembersEnumerator.MoveNext()
                End If
            End Function

            ''' <summary>
            ''' Worker for comparing two named types for equivalence. Note: The two
            ''' types must have the same TypeKind.
            ''' </summary>
            ''' <param name="x">The first type to compare</param>
            ''' <param name="y">The second type to compare</param>
            ''' <param name="equivalentTypesWithDifferingAssemblies">
            ''' Map of equivalent non-nested types to be populated, such that each key-value pair of named types are equivalent but reside in different assemblies.
            ''' This map is populated only if we are ignoring assemblies for symbol equivalence comparison, i.e. <see cref="_assemblyComparerOpt"/> is true.
            ''' </param>
            ''' <returns>True if the two types are equivalent.</returns>
            Private Function HandleNamedTypesWorker(ByVal x As INamedTypeSymbol, ByVal y As INamedTypeSymbol, ByVal equivalentTypesWithDifferingAssemblies As Dictionary(Of INamedTypeSymbol, INamedTypeSymbol)) As Boolean
                Debug.Assert(GetTypeKind(x) = GetTypeKind(y))

                If x.IsTupleType OrElse y.IsTupleType Then
                    If x.IsTupleType <> y.IsTupleType Then
                        Return False
                    End If

                    Dim xElements As ImmutableArray(Of IFieldSymbol) = x.TupleElements
                    Dim yElements As ImmutableArray(Of IFieldSymbol) = y.TupleElements

                    If xElements.Length <> yElements.Length Then
                        Return False
                    End If

                    For i As Integer = 0 To xElements.Length - 1
                        If Not Me.AreEquivalent(xElements(i).Type, yElements(i).Type, equivalentTypesWithDifferingAssemblies) Then
                            Return False
                        End If
                    Next i

                    Return True
                End If

                If x.IsDefinition <> y.IsDefinition OrElse IsConstructedFromSelf(x) <> IsConstructedFromSelf(y) OrElse x.Arity <> y.Arity OrElse x.Name <> y.Name OrElse x.IsAnonymousType <> y.IsAnonymousType OrElse x.IsUnboundGenericType <> y.IsUnboundGenericType Then
                    Return False
                End If

                If Not Me.AreEquivalent(x.ContainingSymbol, y.ContainingSymbol, equivalentTypesWithDifferingAssemblies) Then
                    Return False
                End If

                ' Above check makes sure that the containing assemblies are considered the same by the assembly comparer being used.
                ' If they are in fact not the same (have different name) and the caller requested to know about such types add {x, y}
                ' to equivalentTypesWithDifferingAssemblies map.
                If equivalentTypesWithDifferingAssemblies IsNot Nothing AndAlso x.ContainingType Is Nothing AndAlso x.ContainingAssembly IsNot Nothing AndAlso Not AssemblyIdentityComparer.SimpleNameComparer.Equals(x.ContainingAssembly.Name, y.ContainingAssembly.Name) AndAlso Not equivalentTypesWithDifferingAssemblies.ContainsKey(x) Then
                    equivalentTypesWithDifferingAssemblies.Add(x, y)
                End If

                If x.IsAnonymousType Then
                    Return Me.HandleAnonymousTypes(x, y, equivalentTypesWithDifferingAssemblies)
                End If

                ' They look very similar at this point.  In the case of non constructed types, we're
                ' done.  However, if they are constructed, then their type arguments have to match
                ' as well.
                Return IsConstructedFromSelf(x) OrElse x.IsUnboundGenericType OrElse Me.TypeArgumentsAreEquivalent(x.TypeArguments, y.TypeArguments, equivalentTypesWithDifferingAssemblies)
            End Function

            Private Function MethodsAreEquivalent(ByVal x As IMethodSymbol, ByVal y As IMethodSymbol, ByVal equivalentTypesWithDifferingAssemblies As Dictionary(Of INamedTypeSymbol, INamedTypeSymbol)) As Boolean
                If Not AreCompatibleMethodKinds(x.MethodKind, y.MethodKind) Then
                    Return False
                End If

                If x.MethodKind = MethodKind.ReducedExtension Then
                    Dim rx As IMethodSymbol = x.ReducedFrom
                    Dim ry As IMethodSymbol = y.ReducedFrom

                    ' reduced from symbols are equivalent
                    If Not Me.AreEquivalent(rx, ry, equivalentTypesWithDifferingAssemblies) Then
                        Return False
                    End If

                    ' receiver types are equivalent
                    If Not Me.AreEquivalent(x.ReceiverType, y.ReceiverType, equivalentTypesWithDifferingAssemblies) Then
                        Return False
                    End If
                Else
                    If x.MethodKind = MethodKind.AnonymousFunction OrElse x.MethodKind = MethodKind.LocalFunction Then
                        ' Treat local and anonymous functions just like we do ILocalSymbols.
                        ' They're only equivalent if they have the same location.
                        Return HaveSameLocation(x, y)
                    End If

                    If IsPartialMethodDefinitionPart(x) <> IsPartialMethodDefinitionPart(y) OrElse IsPartialMethodImplementationPart(x) <> IsPartialMethodImplementationPart(y) OrElse x.IsDefinition <> y.IsDefinition OrElse IsConstructedFromSelf(x) <> IsConstructedFromSelf(y) OrElse x.Arity <> y.Arity OrElse x.Parameters.Length <> y.Parameters.Length OrElse x.Name <> y.Name Then
                        Return False
                    End If

                    Dim checkContainingType_ As Boolean = CheckContainingType(x)
                    If checkContainingType_ Then
                        If Not Me.AreEquivalent(x.ContainingSymbol, y.ContainingSymbol, equivalentTypesWithDifferingAssemblies) Then
                            Return False
                        End If
                    End If

                    If Not Me.ParametersAreEquivalent(x.Parameters, y.Parameters, equivalentTypesWithDifferingAssemblies) Then
                        Return False
                    End If

                    If Not Me.ReturnTypesAreEquivalent(x, y, equivalentTypesWithDifferingAssemblies) Then
                        Return False
                    End If
                End If

                ' If it's an unconstructed method, then we don't need to check the type arguments.
                If IsConstructedFromSelf(x) Then
                    Return True
                End If

                Return Me.TypeArgumentsAreEquivalent(x.TypeArguments, y.TypeArguments, equivalentTypesWithDifferingAssemblies)
            End Function

            Private Function ModulesAreEquivalent(ByVal x As IModuleSymbol, ByVal y As IModuleSymbol) As Boolean
                Return Me.AssembliesAreEquivalent(x.ContainingAssembly, y.ContainingAssembly) AndAlso x.Name = y.Name
            End Function

            Private Function NamedTypesAreEquivalent(ByVal x As INamedTypeSymbol, ByVal y As INamedTypeSymbol, ByVal equivalentTypesWithDifferingAssemblies As Dictionary(Of INamedTypeSymbol, INamedTypeSymbol)) As Boolean
                ' PERF: Avoid multiple virtual calls to fetch the TypeKind property
                Dim xTypeKind As TypeKind = GetTypeKind(x)
                Dim yTypeKind As TypeKind = GetTypeKind(y)

                If xTypeKind = TypeKind.Error OrElse yTypeKind = TypeKind.Error Then
                    ' Slow path: x or y is an error type. We need to compare
                    ' all the candidates in both.
                    Return Me.NamedTypesAreEquivalentError(x, y, equivalentTypesWithDifferingAssemblies)
                End If

                ' Fast path: we can compare the symbols directly,
                ' avoiding any allocations associated with the Unwrap()
                ' enumerator.
                Return xTypeKind = yTypeKind AndAlso Me.HandleNamedTypesWorker(x, y, equivalentTypesWithDifferingAssemblies)
            End Function

            Private Function NamedTypesAreEquivalentError(ByVal x As INamedTypeSymbol, ByVal y As INamedTypeSymbol, ByVal equivalentTypesWithDifferingAssemblies As Dictionary(Of INamedTypeSymbol, INamedTypeSymbol)) As Boolean
                For Each type1 As INamedTypeSymbol In Unwrap(x)
                    Dim typeKind1 As TypeKind = GetTypeKind(type1)
                    For Each type2 As INamedTypeSymbol In Unwrap(y)
                        Dim typeKind2 As TypeKind = GetTypeKind(type2)
                        If typeKind1 = typeKind2 AndAlso Me.HandleNamedTypesWorker(type1, type2, equivalentTypesWithDifferingAssemblies) Then
                            Return True
                        End If
                    Next type2
                Next type1

                Return False
            End Function

            Private Function NamespacesAreEquivalent(ByVal x As INamespaceSymbol, ByVal y As INamespaceSymbol, ByVal equivalentTypesWithDifferingAssemblies As Dictionary(Of INamedTypeSymbol, INamedTypeSymbol)) As Boolean
                If x.IsGlobalNamespace <> y.IsGlobalNamespace OrElse x.Name <> y.Name Then
                    Return False
                End If

                If x.IsGlobalNamespace AndAlso _symbolEquivalenceComparer._assemblyComparerOpt Is Nothing Then
                    ' No need to compare the containers of global namespace when assembly identities are ignored.
                    Return True
                End If

                Return Me.AreEquivalent(x.ContainingSymbol, y.ContainingSymbol, equivalentTypesWithDifferingAssemblies)
            End Function

            Private Function ParametersAreEquivalent(ByVal xParameters As ImmutableArray(Of IParameterSymbol), ByVal yParameters As ImmutableArray(Of IParameterSymbol), ByVal equivalentTypesWithDifferingAssemblies As Dictionary(Of INamedTypeSymbol, INamedTypeSymbol), Optional ByVal compareParameterName As Boolean = False, Optional ByVal isParameterNameCaseSensitive As Boolean = False) As Boolean
                ' Note the special parameter comparer we pass in.  We do this so we don't end up
                ' infinitely looping between parameters -> type parameters -> methods -> parameters
                Dim count As Integer = xParameters.Length
                If yParameters.Length <> count Then
                    Return False
                End If

                For i As Integer = 0 To count - 1
                    If Not _symbolEquivalenceComparer.ParameterEquivalenceComparer.Equals(xParameters(i), yParameters(i), equivalentTypesWithDifferingAssemblies, compareParameterName, isParameterNameCaseSensitive) Then
                        Return False
                    End If
                Next i

                Return True
            End Function

            Private Function ParametersAreEquivalent(ByVal x As IParameterSymbol, ByVal y As IParameterSymbol, ByVal equivalentTypesWithDifferingAssemblies As Dictionary(Of INamedTypeSymbol, INamedTypeSymbol)) As Boolean
                Return x.IsRefOrOut() = y.IsRefOrOut() AndAlso x.Name = y.Name AndAlso Me.AreEquivalent(x.CustomModifiers, y.CustomModifiers, equivalentTypesWithDifferingAssemblies) AndAlso Me.AreEquivalent(x.Type, y.Type, equivalentTypesWithDifferingAssemblies) AndAlso Me.AreEquivalent(x.ContainingSymbol, y.ContainingSymbol, equivalentTypesWithDifferingAssemblies)
            End Function

            Private Function PointerTypesAreEquivalent(ByVal x As IPointerTypeSymbol, ByVal y As IPointerTypeSymbol, ByVal equivalentTypesWithDifferingAssemblies As Dictionary(Of INamedTypeSymbol, INamedTypeSymbol)) As Boolean
                Return Me.AreEquivalent(x.CustomModifiers, y.CustomModifiers, equivalentTypesWithDifferingAssemblies) AndAlso Me.AreEquivalent(x.PointedAtType, y.PointedAtType, equivalentTypesWithDifferingAssemblies)
            End Function

            Private Function PropertiesAreEquivalent(ByVal x As IPropertySymbol, ByVal y As IPropertySymbol, ByVal equivalentTypesWithDifferingAssemblies As Dictionary(Of INamedTypeSymbol, INamedTypeSymbol)) As Boolean
                If x.ContainingType.IsAnonymousType AndAlso y.ContainingType.IsAnonymousType Then
                    ' We can short circuit here and just use the symbols themselves to determine
                    ' equality.  This will properly handle things like the VB case where two
                    ' anonymous types will be considered the same if they have properties that
                    ' differ in casing.
                    If SymbolEqualityComparer.Default.Equals(x, y) Then
                        Return True
                    End If
                End If

                Return x.IsIndexer = y.IsIndexer AndAlso x.MetadataName = y.MetadataName AndAlso x.Parameters.Length = y.Parameters.Length AndAlso Me.ParametersAreEquivalent(x.Parameters, y.Parameters, equivalentTypesWithDifferingAssemblies) AndAlso Me.AreEquivalent(x.ContainingSymbol, y.ContainingSymbol, equivalentTypesWithDifferingAssemblies)
            End Function

            Private Function TypeArgumentsAreEquivalent(ByVal xTypeArguments As ImmutableArray(Of ITypeSymbol), ByVal yTypeArguments As ImmutableArray(Of ITypeSymbol), ByVal equivalentTypesWithDifferingAssemblies As Dictionary(Of INamedTypeSymbol, INamedTypeSymbol)) As Boolean
                Dim count As Integer = xTypeArguments.Length
                If yTypeArguments.Length <> count Then
                    Return False
                End If

                For i As Integer = 0 To count - 1
                    If Not Me.AreEquivalent(xTypeArguments(i), yTypeArguments(i), equivalentTypesWithDifferingAssemblies) Then
                        Return False
                    End If
                Next i

                Return True
            End Function

            Private Function TypeParametersAreEquivalent(ByVal x As ITypeParameterSymbol, ByVal y As ITypeParameterSymbol, ByVal equivalentTypesWithDifferingAssemblies As Dictionary(Of INamedTypeSymbol, INamedTypeSymbol)) As Boolean
                Contracts.Contract.Requires((x.TypeParameterKind = TypeParameterKind.Method AndAlso IsConstructedFromSelf(x.DeclaringMethod)) OrElse (x.TypeParameterKind = TypeParameterKind.Type AndAlso IsConstructedFromSelf(x.ContainingType)) OrElse x.TypeParameterKind = TypeParameterKind.Cref)
                Contracts.Contract.Requires((y.TypeParameterKind = TypeParameterKind.Method AndAlso IsConstructedFromSelf(y.DeclaringMethod)) OrElse (y.TypeParameterKind = TypeParameterKind.Type AndAlso IsConstructedFromSelf(y.ContainingType)) OrElse y.TypeParameterKind = TypeParameterKind.Cref)

                If x.Ordinal <> y.Ordinal OrElse x.TypeParameterKind <> y.TypeParameterKind Then
                    Return False
                End If

                ' If this is a method type parameter, and we are in 'non-recurse' mode (because
                ' we're comparing method parameters), then we're done at this point.  The types are
                ' equal.
                If x.TypeParameterKind = TypeParameterKind.Method AndAlso _compareMethodTypeParametersByIndex Then
                    Return True
                End If

                If x.TypeParameterKind = TypeParameterKind.Type AndAlso x.ContainingType.IsAnonymousType Then
                    ' Anonymous type type parameters compare by index as well to prevent
                    ' recursion.
                    Return True
                End If

                If x.TypeParameterKind = TypeParameterKind.Cref Then
                    Return True
                End If

                Return Me.AreEquivalent(x.ContainingSymbol, y.ContainingSymbol, equivalentTypesWithDifferingAssemblies)
            End Function

            Friend Function AreEquivalent(ByVal x As CustomModifier, ByVal y As CustomModifier, ByVal equivalentTypesWithDifferingAssemblies As Dictionary(Of INamedTypeSymbol, INamedTypeSymbol)) As Boolean
                Return x.IsOptional = y.IsOptional AndAlso Me.AreEquivalent(x.Modifier, y.Modifier, equivalentTypesWithDifferingAssemblies)
            End Function

            Friend Function AreEquivalent(ByVal x As ImmutableArray(Of CustomModifier), ByVal y As ImmutableArray(Of CustomModifier), ByVal equivalentTypesWithDifferingAssemblies As Dictionary(Of INamedTypeSymbol, INamedTypeSymbol)) As Boolean
                Debug.Assert(Not x.IsDefault AndAlso Not y.IsDefault)
                If x.Length <> y.Length Then
                    Return False
                End If

                For i As Integer = 0 To x.Length - 1
                    If Not Me.AreEquivalent(x(i), y(i), equivalentTypesWithDifferingAssemblies) Then
                        Return False
                    End If
                Next i

                Return True
            End Function

            Friend Function ReturnTypesAreEquivalent(ByVal x As IMethodSymbol, ByVal y As IMethodSymbol, Optional ByVal equivalentTypesWithDifferingAssemblies As Dictionary(Of INamedTypeSymbol, INamedTypeSymbol) = Nothing) As Boolean
                Return _symbolEquivalenceComparer.SignatureTypeEquivalenceComparer.Equals(x.ReturnType, y.ReturnType, equivalentTypesWithDifferingAssemblies) AndAlso Me.AreEquivalent(x.ReturnTypeCustomModifiers, y.ReturnTypeCustomModifiers, equivalentTypesWithDifferingAssemblies)
            End Function

            Public Function AreEquivalent(ByVal x As ISymbol, ByVal y As ISymbol, ByVal equivalentTypesWithDifferingAssemblies As Dictionary(Of INamedTypeSymbol, INamedTypeSymbol)) As Boolean
                If ReferenceEquals(x, y) Then
                    Return True
                End If

                If x Is Nothing OrElse y Is Nothing Then
                    Return False
                End If

                Dim xKind As SymbolKind = GetKindAndUnwrapAlias(x)
                Dim yKind As SymbolKind = GetKindAndUnwrapAlias(y)

                ' Normally, if they're different types, then they're not the same.
                If xKind <> yKind Then
                    ' Special case.  If we're comparing signatures then we want to compare 'object'
                    ' and 'dynamic' as the same.  However, since they're different types, we don't
                    ' want to bail out using the above check.
                    Return _objectAndDynamicCompareEqually AndAlso ((yKind = SymbolKind.DynamicType AndAlso xKind = SymbolKind.NamedType AndAlso DirectCast(x, ITypeSymbol).SpecialType = SpecialType.System_Object) OrElse (xKind = SymbolKind.DynamicType AndAlso yKind = SymbolKind.NamedType AndAlso DirectCast(y, ITypeSymbol).SpecialType = SpecialType.System_Object))
                End If

                Return Me.AreEquivalentWorker(x, y, xKind, equivalentTypesWithDifferingAssemblies)

            End Function

        End Class

    End Class

End Namespace
