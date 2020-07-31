' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.
'

Imports System.Collections.Immutable
Imports HashLibrary
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Shared.Extensions

Namespace Utilities

    Partial Friend Class SymbolEquivalenceComparer

        Private Class GetHashCodeVisitor
            Private ReadOnly _compareMethodTypeParametersByIndex As Boolean
            Private ReadOnly _objectAndDynamicCompareEqually As Boolean
            Private ReadOnly _parameterAggregator As Func(Of Integer, IParameterSymbol, Integer)
            Private ReadOnly _symbolAggregator As Func(Of Integer, ISymbol, Integer)
            Private ReadOnly _symbolEquivalenceComparer As SymbolEquivalenceComparer

            Public Sub New(ByVal symbolEquivalenceComparer As SymbolEquivalenceComparer, ByVal compareMethodTypeParametersByIndex As Boolean, ByVal objectAndDynamicCompareEqually As Boolean)
                _symbolEquivalenceComparer = symbolEquivalenceComparer
                _compareMethodTypeParametersByIndex = compareMethodTypeParametersByIndex
                _objectAndDynamicCompareEqually = objectAndDynamicCompareEqually
                _parameterAggregator = Function(acc, sym) CodeRefactoringHash.Combine(symbolEquivalenceComparer.ParameterEquivalenceComparer.GetHashCode(sym), acc)
                _symbolAggregator = Function(acc, sym) Me.GetHashCode(sym, acc)
            End Sub

            Private Shared Function CombineHashCodes(Of T)(ByVal array As ImmutableArray(Of T), ByVal currentHash As Integer, ByVal func As Func(Of Integer, T, Integer)) As Integer
                Return array.Aggregate(currentHash, func)
            End Function

            Private Shared Function CombineHashCodes(ByVal x As ILabelSymbol, ByVal currentHash As Integer) As Integer
                Return CodeRefactoringHash.Combine(x.Name, CodeRefactoringHash.Combine(x.Locations.FirstOrDefault(), currentHash))
            End Function

            Private Shared Function CombineHashCodes(ByVal x As ILocalSymbol, ByVal currentHash As Integer) As Integer
                Return CodeRefactoringHash.Combine(x.Locations.FirstOrDefault(), currentHash)
            End Function

            Private Shared Function CombineHashCodes(ByVal x As IRangeVariableSymbol, ByVal currentHash As Integer) As Integer
                Return CodeRefactoringHash.Combine(x.Locations.FirstOrDefault(), currentHash)
            End Function

            Private Shared Function CombineHashCodes(ByVal x As IPreprocessingSymbol, ByVal currentHash As Integer) As Integer
                Return CodeRefactoringHash.Combine(x.GetHashCode(), currentHash)
            End Function

            Private Function CombineAnonymousTypeHashCode(ByVal x As INamedTypeSymbol, ByVal currentHash As Integer) As Integer
                If x.TypeKind = TypeKind.Delegate Then
                    Return Me.GetHashCode(x.DelegateInvokeMethod, currentHash)
                Else
                    Dim xMembers As IEnumerable(Of IPropertySymbol) = x.GetValidAnonymousTypeProperties()

                    Return xMembers.Aggregate(currentHash, Function(a, p)
                                                               Return CodeRefactoringHash.Combine(p.Name, CodeRefactoringHash.Combine(p.IsReadOnly, Me.GetHashCode(p.Type, a)))
                                                           End Function)
                End If
            End Function

            Private Function CombineHashCodes(ByVal x As IArrayTypeSymbol, ByVal currentHash As Integer) As Integer
                Return CodeRefactoringHash.Combine(x.Rank, Me.GetHashCode(x.ElementType, currentHash))
            End Function

            Private Function CombineHashCodes(ByVal x As IAssemblySymbol, ByVal currentHash As Integer) As Integer
                Return CodeRefactoringHash.Combine(If(_symbolEquivalenceComparer._assemblyComparerOpt?.GetHashCode(x), 0), currentHash)
            End Function

            Private Function CombineHashCodes(ByVal x As IFieldSymbol, ByVal currentHash As Integer) As Integer
                Return CodeRefactoringHash.Combine(x.Name, Me.GetHashCode(x.ContainingSymbol, currentHash))
            End Function

            Private Function CombineHashCodes(ByVal x As IMethodSymbol, ByVal currentHash As Integer) As Integer
                currentHash = CodeRefactoringHash.Combine(x.MetadataName, currentHash)
                If x.MethodKind = MethodKind.AnonymousFunction Then
                    Return CodeRefactoringHash.Combine(x.Locations.FirstOrDefault(), currentHash)
                End If

                currentHash = CodeRefactoringHash.Combine(IsPartialMethodImplementationPart(x), CodeRefactoringHash.Combine(IsPartialMethodDefinitionPart(x), CodeRefactoringHash.Combine(x.IsDefinition, CodeRefactoringHash.Combine(IsConstructedFromSelf(x), CodeRefactoringHash.Combine(x.Arity, CodeRefactoringHash.Combine(x.Parameters.Length, CodeRefactoringHash.Combine(x.Name, currentHash)))))))

                Dim checkContainingType_ As Boolean = CheckContainingType(x)
                If checkContainingType_ Then
                    currentHash = Me.GetHashCode(x.ContainingSymbol, currentHash)
                End If

                currentHash = CombineHashCodes(x.Parameters, currentHash, _parameterAggregator)

                Return If(IsConstructedFromSelf(x), currentHash, CombineHashCodes(x.TypeArguments, currentHash, _symbolAggregator))
            End Function

            Private Function CombineHashCodes(ByVal x As IModuleSymbol, ByVal currentHash As Integer) As Integer
                Return Me.CombineHashCodes(x.ContainingAssembly, CodeRefactoringHash.Combine(x.Name, currentHash))
            End Function

            Private Function CombineHashCodes(ByVal x As INamedTypeSymbol, ByVal currentHash As Integer) As Integer
                currentHash = Me.CombineNamedTypeHashCode(x, currentHash)

                Dim errorType As IErrorTypeSymbol = TryCast(x, IErrorTypeSymbol)
                If errorType IsNot Nothing Then
                    For Each candidate As ISymbol In errorType.CandidateSymbols
                        Dim candidateNamedType As INamedTypeSymbol = TryCast(candidate, INamedTypeSymbol)
                        If candidateNamedType IsNot Nothing Then
                            currentHash = Me.CombineNamedTypeHashCode(candidateNamedType, currentHash)
                        End If
                    Next candidate
                End If

                Return currentHash
            End Function

            Private Function CombineHashCodes(ByVal x As INamespaceSymbol, ByVal currentHash As Integer) As Integer
                If x.IsGlobalNamespace AndAlso _symbolEquivalenceComparer._assemblyComparerOpt Is Nothing Then
                    ' Exclude global namespace's container's hash when assemblies can differ.
                    Return CodeRefactoringHash.Combine(x.Name, currentHash)
                End If

                Return CodeRefactoringHash.Combine(x.IsGlobalNamespace, CodeRefactoringHash.Combine(x.Name, Me.GetHashCode(x.ContainingSymbol, currentHash)))
            End Function

            Private Function CombineHashCodes(ByVal x As IParameterSymbol, ByVal currentHash As Integer) As Integer
                Return CodeRefactoringHash.Combine(x.IsRefOrOut(), CodeRefactoringHash.Combine(x.Name, Me.GetHashCode(x.Type, Me.GetHashCode(x.ContainingSymbol, currentHash))))
            End Function

            Private Function CombineHashCodes(ByVal x As IPointerTypeSymbol, ByVal currentHash As Integer) As Integer
                Return CodeRefactoringHash.Combine(GetType(IPointerTypeSymbol).GetHashCode(), Me.GetHashCode(x.PointedAtType, currentHash))
            End Function

            Private Function CombineHashCodes(ByVal x As IPropertySymbol, ByVal currentHash As Integer) As Integer
                currentHash = CodeRefactoringHash.Combine(x.IsIndexer, CodeRefactoringHash.Combine(x.Name, CodeRefactoringHash.Combine(x.Parameters.Length, Me.GetHashCode(x.ContainingSymbol, currentHash))))

                Return CombineHashCodes(x.Parameters, currentHash, _parameterAggregator)
            End Function

            Private Function CombineHashCodes(ByVal x As IEventSymbol, ByVal currentHash As Integer) As Integer
                Return CodeRefactoringHash.Combine(x.Name, Me.GetHashCode(x.ContainingSymbol, currentHash))
            End Function

            Private Function CombineNamedTypeHashCode(ByVal x As INamedTypeSymbol, ByVal currentHash As Integer) As Integer
                If x.IsTupleType Then
                    Return CodeRefactoringHash.Combine(currentHash, CodeRefactoringHash.CombineValues(x.TupleElements))
                End If

                ' If we want object and dynamic to be the same, and this is 'object', then return
                ' the same hash we do for 'dynamic'.
                currentHash = CodeRefactoringHash.Combine(x.IsDefinition, CodeRefactoringHash.Combine(IsConstructedFromSelf(x), CodeRefactoringHash.Combine(x.Arity, CodeRefactoringHash.Combine(CInt(Math.Truncate(GetTypeKind(x))), CodeRefactoringHash.Combine(x.Name, CodeRefactoringHash.Combine(x.IsAnonymousType, CodeRefactoringHash.Combine(x.IsUnboundGenericType, Me.GetHashCode(x.ContainingSymbol, currentHash))))))))

                If x.IsAnonymousType Then
                    Return Me.CombineAnonymousTypeHashCode(x, currentHash)
                End If

                Return If(IsConstructedFromSelf(x) OrElse x.IsUnboundGenericType, currentHash, CombineHashCodes(x.TypeArguments, currentHash, _symbolAggregator))
            End Function

            Private Function GetHashCodeWorker(ByVal x As ISymbol, ByVal currentHash As Integer) As Integer
                Select Case x.Kind
                    Case SymbolKind.ArrayType
                        Return Me.CombineHashCodes(DirectCast(x, IArrayTypeSymbol), currentHash)
                    Case SymbolKind.Assembly
                        Return Me.CombineHashCodes(DirectCast(x, IAssemblySymbol), currentHash)
                    Case SymbolKind.Event
                        Return Me.CombineHashCodes(DirectCast(x, IEventSymbol), currentHash)
                    Case SymbolKind.Field
                        Return Me.CombineHashCodes(DirectCast(x, IFieldSymbol), currentHash)
                    Case SymbolKind.Label
                        Return CombineHashCodes(DirectCast(x, ILabelSymbol), currentHash)
                    Case SymbolKind.Local
                        Return CombineHashCodes(DirectCast(x, ILocalSymbol), currentHash)
                    Case SymbolKind.Method
                        Return Me.CombineHashCodes(DirectCast(x, IMethodSymbol), currentHash)
                    Case SymbolKind.NetModule
                        Return Me.CombineHashCodes(DirectCast(x, IModuleSymbol), currentHash)
                    Case SymbolKind.NamedType
                        Return Me.CombineHashCodes(DirectCast(x, INamedTypeSymbol), currentHash)
                    Case SymbolKind.Namespace
                        Return Me.CombineHashCodes(DirectCast(x, INamespaceSymbol), currentHash)
                    Case SymbolKind.Parameter
                        Return Me.CombineHashCodes(DirectCast(x, IParameterSymbol), currentHash)
                    Case SymbolKind.PointerType
                        Return Me.CombineHashCodes(DirectCast(x, IPointerTypeSymbol), currentHash)
                    Case SymbolKind.Property
                        Return Me.CombineHashCodes(DirectCast(x, IPropertySymbol), currentHash)
                    Case SymbolKind.RangeVariable
                        Return CombineHashCodes(DirectCast(x, IRangeVariableSymbol), currentHash)
                    Case SymbolKind.TypeParameter
                        Return Me.CombineHashCodes(DirectCast(x, ITypeParameterSymbol), currentHash)
                    Case SymbolKind.Preprocessing
                        Return CombineHashCodes(DirectCast(x, IPreprocessingSymbol), currentHash)
                    Case Else
                        Return -1
                End Select
            End Function

            Public Function CombineHashCodes(ByVal x As ITypeParameterSymbol, ByVal currentHash As Integer) As Integer
                Contracts.Contract.Requires((x.TypeParameterKind = TypeParameterKind.Method AndAlso IsConstructedFromSelf(x.DeclaringMethod)) OrElse (x.TypeParameterKind = TypeParameterKind.Type AndAlso IsConstructedFromSelf(x.ContainingType)) OrElse x.TypeParameterKind = TypeParameterKind.Cref)

                currentHash = CodeRefactoringHash.Combine(x.Ordinal, CodeRefactoringHash.Combine(CInt(Math.Truncate(x.TypeParameterKind)), currentHash))

                If x.TypeParameterKind = TypeParameterKind.Method AndAlso _compareMethodTypeParametersByIndex Then
                    Return currentHash
                End If

                If x.TypeParameterKind = TypeParameterKind.Type AndAlso x.ContainingType.IsAnonymousType Then
                    ' Anonymous type, type parameters compare by index as well to prevent
                    ' recursion.
                    Return currentHash
                End If

                If x.TypeParameterKind = TypeParameterKind.Cref Then
                    Return currentHash
                End If

                Return Me.GetHashCode(x.ContainingSymbol, currentHash)
            End Function

            Public Shadows Function GetHashCode(ByVal x As ISymbol, ByVal currentHash As Integer) As Integer
                If x Is Nothing Then
                    Return 0
                End If

                x = UnwrapAlias(x)

                ' Special case.  If we're comparing signatures then we want to compare 'object'
                ' and 'dynamic' as the same.  However, since they're different types, we don't
                ' want to bail out using the above check.

                If x.Kind = SymbolKind.DynamicType OrElse (_objectAndDynamicCompareEqually AndAlso IsObjectType(x)) Then
                    Return CodeRefactoringHash.Combine(GetType(IDynamicTypeSymbol), currentHash)
                End If

                Return Me.GetHashCodeWorker(x, currentHash)
            End Function

        End Class

    End Class

End Namespace
