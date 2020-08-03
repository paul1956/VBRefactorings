' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Runtime.InteropServices
Imports Microsoft.CodeAnalysis
Imports VBRefactorings.Utilities

Partial Public Module ISymbolExtensions

    ' Is a member with declared accessibility "declaredAccessibility" accessible from within
    ' "within", which must be a named type or an assembly.
    Private Function IsMemberAccessible(
            containingType As INamedTypeSymbol,
            declaredAccessibility As Accessibility,
            within As ISymbol,
            throughTypeOpt As ITypeSymbol,
            <Out> ByRef failedThroughTypeCheck As Boolean) As Boolean
        Debug.Assert(TypeOf within Is INamedTypeSymbol OrElse TypeOf within Is IAssemblySymbol)
        Contracts.Contract.Requires(containingType IsNot Nothing)

        failedThroughTypeCheck = False

        Dim originalContainingType As INamedTypeSymbol = containingType.OriginalDefinition
        Dim withinNamedType As INamedTypeSymbol = TryCast(within, INamedTypeSymbol)
        Dim withinAssembly As IAssemblySymbol = If(TryCast(within, IAssemblySymbol), CType(within, INamedTypeSymbol).ContainingAssembly)

        ' A nested symbol is only accessible to us if its container is accessible as well.
        If Not IsNamedTypeAccessible(containingType, within) Then
            Return False
        End If

        Select Case declaredAccessibility
            Case Accessibility.NotApplicable
                ' TODO(cyrusn): Is this the right thing to do here?  Should the caller ever be
                ' asking about the accessibility of a symbol that has "NotApplicable" as its
                ' value?  For now, I'm preserving the behavior of the existing code.  But perhaps
                ' we should fail here and require the caller to not do this?
                Return True

            Case Accessibility.[Public]
                ' Public symbols are always accessible from any context
                Return True

            Case Accessibility.[Private]
                ' All expressions in the current submission (top-level or nested in a method or
                ' type) can access previous submission's private top-level members. Previous
                ' submissions are treated like outer classes for the current submission - the
                ' inner class can access private members of the outer class.
                If withinAssembly.IsInteractive AndAlso containingType.IsScriptClass Then
                    Return True
                End If

                ' private members never accessible from outside a type.
                Return withinNamedType IsNot Nothing AndAlso IsPrivateSymbolAccessible(withinNamedType, originalContainingType)

            Case Accessibility.Internal
                ' An internal type is accessible if we're in the same assembly or we have
                ' friend access to the assembly it was defined in.
                Return withinAssembly.IsSameAssemblyOrHasFriendAccessTo(containingType.ContainingAssembly)

            Case Accessibility.ProtectedAndInternal
                If Not withinAssembly.IsSameAssemblyOrHasFriendAccessTo(containingType.ContainingAssembly) Then
                    ' We require internal access.  If we don't have it, then this symbol is
                    ' definitely not accessible to us.
                    Return False
                End If

                ' We had internal access.  Also have to make sure we have protected access.
                Return IsProtectedSymbolAccessible(withinNamedType, withinAssembly, throughTypeOpt, originalContainingType, failedThroughTypeCheck)

            Case Accessibility.ProtectedOrInternal
                If withinAssembly.IsSameAssemblyOrHasFriendAccessTo(containingType.ContainingAssembly) Then
                    ' If we have internal access to this symbol, then that's sufficient.  no
                    ' need to do the complicated protected case.
                    Return True
                End If

                ' We don't have internal access.  But if we have protected access then that's
                ' sufficient.
                Return IsProtectedSymbolAccessible(withinNamedType, withinAssembly, throughTypeOpt, originalContainingType, failedThroughTypeCheck)

            Case Accessibility.[Protected]
                Return IsProtectedSymbolAccessible(withinNamedType, withinAssembly, throughTypeOpt, originalContainingType, failedThroughTypeCheck)
            Case Else
                Throw UnexpectedValue(declaredAccessibility)
        End Select
    End Function

    ' Is the named type "type" accessible from within "within", which must be a named type or
    ' an assembly.
    Friend Function IsNamedTypeAccessible(type As INamedTypeSymbol, within As ISymbol) As Boolean
        Debug.Assert(TypeOf within Is INamedTypeSymbol OrElse TypeOf within Is IAssemblySymbol)
        Contracts.Contract.Requires(type IsNot Nothing)

        If type.IsErrorType() Then
            ' Always assume that error types are accessible.
            Return True
        End If

        Dim unused As Boolean
        If Not type.IsDefinition Then
            For Each typeArg As ITypeSymbol In type.TypeArguments
                ' type parameters are always accessible, so don't check those (so common it's
                ' worth optimizing this).
                If typeArg.Kind <> SymbolKind.TypeParameter AndAlso
                    typeArg.TypeKind <> Global.Microsoft.CodeAnalysis.TypeKind.[Error] AndAlso
                    Not IsSymbolAccessibleCore(typeArg, within, Nothing, unused) Then
                    Return False
                End If
            Next
        End If

        Dim containingType As INamedTypeSymbol = type.ContainingType
        Return If(containingType Is Nothing,
                IsNonNestedTypeAccessible(type.ContainingAssembly, type.DeclaredAccessibility, within), IsMemberAccessible(type.ContainingType, type.DeclaredAccessibility, within, Nothing, unused))
    End Function

    ' Is the type "withinType" nested within the original type "originalContainingType".
    Private Function IsNestedWithinOriginalContainingType(
            withinType As INamedTypeSymbol,
            originalContainingType As INamedTypeSymbol) As Boolean
        Contracts.Contract.Requires(withinType IsNot Nothing)
        Contracts.Contract.Requires(originalContainingType IsNot Nothing)

        ' Walk up my parent chain and see if I eventually hit the owner.  If so then I'm a
        ' nested type of that owner and I'm allowed access to everything inside of it.
        Dim current As INamedTypeSymbol = withinType.OriginalDefinition
        While current IsNot Nothing
            Debug.Assert(current.IsDefinition)
            If SymbolEqualityComparer.Default.Equals(current, originalContainingType) Then
                Return True
            End If

            ' NOTE(cyrusn): The container of an 'original' type is always original.
            current = current.ContainingType
        End While

        Return False
    End Function

    ' Is a top-level type with accessibility "declaredAccessibility" inside assembly "assembly"
    ' accessible from "within", which must be a named type of an assembly.
    Private Function IsNonNestedTypeAccessible(
            assembly As IAssemblySymbol,
            declaredAccessibility As Accessibility,
            within As ISymbol) As Boolean
        Debug.Assert(TypeOf within Is INamedTypeSymbol OrElse TypeOf within Is IAssemblySymbol)
        Contracts.Contract.Requires(assembly IsNot Nothing)
        Dim withinAssembly As IAssemblySymbol = If(TryCast(within, IAssemblySymbol), CType(within, INamedTypeSymbol).ContainingAssembly)

        Select Case declaredAccessibility
            Case Accessibility.NotApplicable, Accessibility.[Public]
                ' Public symbols are always accessible from any context
                Return True

            Case Accessibility.[Private], Accessibility.[Protected], Accessibility.ProtectedAndInternal
                ' Shouldn't happen except in error cases.
                Return False

            Case Accessibility.Internal, Accessibility.ProtectedOrInternal
                ' An internal type is accessible if we're in the same assembly or we have
                ' friend access to the assembly it was defined in.
                Return withinAssembly.IsSameAssemblyOrHasFriendAccessTo(assembly)
            Case Else
                Throw UnexpectedValue(declaredAccessibility)
        End Select
    End Function

    ' Is a private symbol access
    Private Function IsPrivateSymbolAccessible(
            within As ISymbol,
            originalContainingType As INamedTypeSymbol) As Boolean
        Debug.Assert(TypeOf within Is INamedTypeSymbol OrElse TypeOf within Is IAssemblySymbol)

        Dim withinType As INamedTypeSymbol = TryCast(within, INamedTypeSymbol)
        If withinType Is Nothing Then
            ' If we're not within a type, we can't access a private symbol
            Return False
        End If

        ' A private symbol is accessible if we're (optionally nested) inside the type that it
        ' was defined in.
        Return IsNestedWithinOriginalContainingType(withinType, originalContainingType)
    End Function

    ' Is a protected symbol inside "originalContainingType" accessible from within "within",
    ' which much be a named type or an assembly.
    Private Function IsProtectedSymbolAccessible(
            withinType As INamedTypeSymbol,
            withinAssembly As IAssemblySymbol,
            throughTypeOpt As ITypeSymbol,
            originalContainingType As INamedTypeSymbol,
            <Out> ByRef failedThroughTypeCheck As Boolean) As Boolean
        failedThroughTypeCheck = False

        ' It is not an error to define protected member in a sealed Script class,
        ' it's just a warning. The member behaves like a private one - it is visible
        ' in all subsequent submissions.
        If withinAssembly.IsInteractive AndAlso originalContainingType.IsScriptClass Then
            Return True
        End If

        If withinType Is Nothing Then
            ' If we're not within a type, we can't access a protected symbol
            Return False
        End If

        ' A protected symbol is accessible if we're (optionally nested) inside the type that it
        ' was defined in.
        ' NOTE: It is helpful to consider 'protected' as *increasing* the
        ' accessibility domain of a private member, rather than *decreasing* that of a public
        ' member. Members are naturally private; the protected, internal and public access
        ' modifiers all increase the accessibility domain. Since private members are accessible
        ' to nested types, so are protected members.
        ' NOTE: We do this check up front as it is very fast and easy to do.
        If IsNestedWithinOriginalContainingType(withinType, originalContainingType) Then
            Return True
        End If

        ' Protected is really confusing.  Check out 3.5.3 of the language spec "protected access
        ' for instance members" to see how it works.  I actually got the code for this from
        ' LangCompiler::CheckAccessCore

        Dim current As INamedTypeSymbol = withinType.OriginalDefinition
        Dim originalThroughTypeOpt As ITypeSymbol = throughTypeOpt?.OriginalDefinition
        While current IsNot Nothing
            Debug.Assert(current.IsDefinition)

            If current.InheritsFromOrEqualsIgnoringConstruction(originalContainingType) Then
                ' NOTE(cyrusn): We're continually walking up the 'throughType's inheritance
                ' chain.  We could compute it up front and cache it in a set.  However, i
                ' don't want to allocate memory in this function.  Also, in practice
                ' inheritance chains should be very short.  As such, it might actually be
                ' slower to create and check inside the set versus just walking the
                ' inheritance chain.
                If originalThroughTypeOpt Is Nothing OrElse
                    originalThroughTypeOpt.InheritsFromOrEqualsIgnoringConstruction(current) Then
                    Return True
                Else
                    failedThroughTypeCheck = True
                End If
            End If

            ' NOTE(cyrusn): The container of an original type is always original.
            current = current.ContainingType
        End While

        Return False
    End Function

End Module
