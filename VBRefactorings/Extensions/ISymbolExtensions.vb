' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports System.Runtime.CompilerServices
Imports Microsoft.CodeAnalysis

' All type argument must be accessible.

Public Module ISymbolExtensions

    Public Enum SymbolVisibility
        [Public]
        Internal
        [Private]
    End Enum

    ''' <summary>
    ''' Checks if 'symbol' is accessible from within 'within', which must be a INamedTypeSymbol
    ''' or an IAssemblySymbol.  If 'symbol' is accessed off of an expression then
    ''' 'throughTypeOpt' is the type of that expression. This is needed to properly do protected
    ''' access checks. Sets "failedThroughTypeCheck" to true if this protected check failed.
    ''' </summary>
    '// NOTE(cyrusn): I expect this function to be called a lot.  As such, I do not do any memory
    '// allocations in the function itself (including not making any iterators).  This does mean
    '// that certain helper functions that we'd like to call are inlined in this method to
    '// prevent the overhead of returning collections or enumerators.
    Private Function IsSymbolAccessibleCore(symbol As ISymbol, within As ISymbol, throughTypeOpt As ITypeSymbol, ByRef failedThroughTypeCheck As Boolean) As Boolean ' must be assembly or named type symbol
        '			Contract.ThrowIfNull(symbol);
        '			Contract.ThrowIfNull(within);
        '			Contract.Requires(within is INamedTypeSymbol || within is IAssemblySymbol);

        failedThroughTypeCheck = False
        ' var withinAssembly = (within as IAssemblySymbol) ?? ((INamedTypeSymbol)within).ContainingAssembly;

        Select Case symbol.Kind
            Case SymbolKind.Alias
                Return IsSymbolAccessibleCore(DirectCast(symbol, IAliasSymbol).Target, within, throughTypeOpt, failedThroughTypeCheck)

            Case SymbolKind.ArrayType
                Return IsSymbolAccessibleCore(DirectCast(symbol, IArrayTypeSymbol).ElementType, within, Nothing, failedThroughTypeCheck)

            Case SymbolKind.PointerType
                Return IsSymbolAccessibleCore(DirectCast(symbol, IPointerTypeSymbol).PointedAtType, within, Nothing, failedThroughTypeCheck)

            Case SymbolKind.NamedType
                Return IsNamedTypeAccessible(DirectCast(symbol, INamedTypeSymbol), within)

            Case SymbolKind.ErrorType
                Return True

            Case SymbolKind.TypeParameter, SymbolKind.Parameter, SymbolKind.Local, SymbolKind.Label, SymbolKind.Namespace, SymbolKind.DynamicType, SymbolKind.Assembly, SymbolKind.NetModule, SymbolKind.RangeVariable
                ' These types of symbols are always accessible (if visible).
                Return True

            Case SymbolKind.Method, SymbolKind.Property, SymbolKind.Field, SymbolKind.Event
                If symbol.IsStatic Then
                    ' static members aren't accessed "through" an "instance" of any type.  So we
                    ' null out the "through" instance here.  This ensures that we'll understand
                    ' accessing protected statics properly.
                    throughTypeOpt = Nothing
                End If

                ' If this is a synthesized operator of dynamic, it's always accessible.
                If symbol.IsKind(SymbolKind.Method) AndAlso DirectCast(symbol, IMethodSymbol).MethodKind = MethodKind.BuiltinOperator AndAlso symbol.ContainingSymbol.IsKind(SymbolKind.DynamicType) Then
                    Return True
                End If

                ' If it's a synthesized operator on a pointer, use the pointer's PointedAtType.
                If symbol.IsKind(SymbolKind.Method) AndAlso DirectCast(symbol, IMethodSymbol).MethodKind = MethodKind.BuiltinOperator AndAlso symbol.ContainingSymbol.IsKind(SymbolKind.PointerType) Then
                    Return IsSymbolAccessibleCore(DirectCast(symbol.ContainingSymbol, IPointerTypeSymbol).PointedAtType, within, Nothing, failedThroughTypeCheck)
                End If

                Return IsMemberAccessible(symbol.ContainingType, symbol.DeclaredAccessibility, within, throughTypeOpt, failedThroughTypeCheck)

            Case Else
                Throw UnexpectedValue(symbol.Kind)
        End Select
    End Function

    <Extension>
    Public Function ActionType(compilation As Compilation) As INamedTypeSymbol
        Return compilation.GetTypeByMetadataName(GetType(Action).FullName)
    End Function

    <Extension>
    Public Function ConvertToType(symbol As ISymbol, compilation As Compilation, Optional extensionUsedAsInstance As Boolean = False) As ITypeSymbol
        Dim _type As ITypeSymbol = TryCast(symbol, ITypeSymbol)
        If _type IsNot Nothing Then
            Return _type
        End If

        Dim method As IMethodSymbol = TryCast(symbol, IMethodSymbol)
        If method IsNot Nothing AndAlso Not method.Parameters.Any(Function(p As IParameterSymbol) p.RefKind <> RefKind.None) Then
            ' Convert the symbol to Func<...> or Action<...>
            If method.ReturnsVoid Then
                Dim count As Integer = If(extensionUsedAsInstance, method.Parameters.Length - 1, method.Parameters.Length)
                Dim skip As Integer = If(extensionUsedAsInstance, 1, 0)
                count = Math.Max(0, count)
                If count = 0 Then
                    ' Action
                    Return compilation.ActionType()
                Else
                    ' Action<TArg1, ..., TArgN>
                    Dim actionName As String = "System.Action`" & count
                    'INSTANT VB NOTE: The variable actionType was renamed since Visual Basic does not handle local variables named the same as class members well:
                    Dim actionType_Renamed As INamedTypeSymbol = compilation.GetTypeByMetadataName(actionName)

                    If actionType_Renamed IsNot Nothing Then
                        Dim types() As ITypeSymbol = method.Parameters.
                            Skip(skip).
                            Select(Function(p As IParameterSymbol) If(p.Type, compilation.GetSpecialType(SpecialType.System_Object))).
                            ToArray()
                        Return actionType_Renamed.Construct(types)
                    End If
                End If
            Else
                ' Func<TArg1,...,TArgN,TReturn>
                '
                ' +1 for the return type.
                Dim count As Integer = If(extensionUsedAsInstance, method.Parameters.Length - 1, method.Parameters.Length)
                Dim skip As Integer = If(extensionUsedAsInstance, 1, 0)
                Dim functionName As String = "System.Func`" & (count + 1)
                Dim functionType As INamedTypeSymbol = compilation.GetTypeByMetadataName(functionName)

                If functionType IsNot Nothing Then
                    Try
                        Dim types() As ITypeSymbol = method.Parameters.
                            Skip(skip).Select(Function(p As IParameterSymbol) p.Type).
                            Concat(method.ReturnType).
                            Select(Function(t As ITypeSymbol) If(t, compilation.GetSpecialType(SpecialType.System_Object))).
                            ToArray()
                        Return functionType.Construct(types)
                    Catch ex As Exception
                        Stop
                    End Try
                End If
            End If
        End If

        ' Otherwise, just default to object.
        Return compilation.ObjectType
    End Function

    <Extension>
    Public Function GetResultantVisibility(symbol As ISymbol) As SymbolVisibility
        ' Start by assuming it's visible.
        Dim visibility As SymbolVisibility = SymbolVisibility.Public

        Select Case symbol.Kind
            Case SymbolKind.Alias
                ' Aliases are uber private.  They're only visible in the same file that they
                ' were declared in.
                Return SymbolVisibility.Private

            Case SymbolKind.Parameter
                ' Parameters are only as visible as their containing symbol
                Return GetResultantVisibility(symbol.ContainingSymbol)

            Case SymbolKind.TypeParameter
                ' Type Parameters are private.
                Return SymbolVisibility.Private
        End Select

        Do While symbol IsNot Nothing AndAlso symbol.Kind <> SymbolKind.Namespace
            Select Case symbol.DeclaredAccessibility
                    ' If we see anything private, then the symbol is private.
                Case Accessibility.NotApplicable, Accessibility.Private
                    Return SymbolVisibility.Private

                    ' If we see anything internal, then knock it down from public to
                    ' internal.
                Case Accessibility.Internal, Accessibility.ProtectedAndInternal
                    visibility = SymbolVisibility.Internal

                    ' For anything else (Public, Protected, ProtectedOrInternal), the
                    ' symbol stays at the level we've gotten so far.
            End Select

            symbol = symbol.ContainingSymbol
        Loop

        Return visibility
    End Function

    <Extension>
    Public Function GetTypeArguments(symbol As ISymbol) As ImmutableArray(Of ITypeSymbol)
        'ORIGINAL LINE: return symbol.TypeSwitch( (IMethodSymbol m) => m.TypeArguments, (INamedTypeSymbol nt) => nt.TypeArguments, _ => ImmutableArray.Create<ITypeSymbol>());
        Return symbol.TypeSwitch(
                Function(m As IMethodSymbol) m.TypeArguments,
                Function(nt As INamedTypeSymbol) nt.TypeArguments,
                Function(underscore As ISymbol) ImmutableArray.Create(Of ITypeSymbol)())
    End Function

    <Extension>
    Public Function GetValidAnonymousTypeProperties(symbol As ISymbol) As IEnumerable(Of IPropertySymbol)
        ' Contract.ThrowIfFalse(symbol.IsNormalAnonymousType());
        Return DirectCast(symbol, INamedTypeSymbol).GetMembers().OfType(Of IPropertySymbol)().Where(Function(p As IPropertySymbol) p.CanBeReferencedByName)
    End Function

    <Extension>
    Public Function IsArrayType(symbol As ISymbol) As Boolean
        Return CBool(symbol?.Kind = SymbolKind.ArrayType)
    End Function

    <Extension>
    Public Function IsAttribute(symbol As ISymbol) As Boolean
        Return CBool(TryCast(symbol, ITypeSymbol)?.IsAttribute() = True)
    End Function

    <Extension>
    Public Function IsInterfaceType(symbol As ISymbol) As Boolean
        Return CBool(TryCast(symbol, ITypeSymbol)?.IsInterfaceType() = True)
    End Function

    <Extension>
    Public Function IsKind(symbol As ISymbol, kind As SymbolKind) As Boolean
        Return symbol.MatchesKind(kind)
    End Function

    <Extension>
    Public Function IsModuleType(symbol As ISymbol) As Boolean
        Return CBool(TryCast(symbol, ITypeSymbol)?.IsModuleType() = True)
    End Function

    <Extension>
    Public Function MatchesKind(symbol As ISymbol, kind As SymbolKind) As Boolean
        Return CBool(symbol?.Kind = kind)
    End Function

End Module
