using Microsoft.CodeAnalysis;
using SpreadCheetah.SourceGenerator.Models;
using System.Collections.Immutable;

namespace SpreadCheetah.SourceGenerator;

internal static class Diagnostics
{
    private const string Category = "SpreadCheetah.SourceGenerator";

    public static ImmutableArray<DiagnosticDescriptor> AllDescriptors =>
    [
        NoPropertiesFoundDescriptor,
        UnsupportedTypeForCellValueDescriptor,
        DuplicateColumnOrderDescriptor,
        InvalidColumnHeaderPropertyReferenceDescriptor,
        UnsupportedTypeForAttributeDescriptor,
        InvalidAttributeArgumentDescriptor,
        AttributeTypeArgumentMustInheritDescriptor,
        AttributeCombinationNotSupportedDescriptor,
        AttributeTypeArgumentMustHaveDefaultConstructorDescriptor
    ];

    // TODO: Class-level diagnostic
    public static Diagnostic NoPropertiesFound(Location? location, string rowTypeName)
        => Diagnostic.Create(NoPropertiesFoundDescriptor, location, rowTypeName);

    private static readonly DiagnosticDescriptor NoPropertiesFoundDescriptor = new(
        id: "SPCH1001",
        title: "Missing properties with public getters",
        messageFormat: "The type '{0}' has no properties with public getters. This will cause an empty row to be added.",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Warning,
        isEnabledByDefault: true);

    // TODO: Class-level diagnostic
    public static Diagnostic UnsupportedTypeForCellValue(Location? location, string rowTypeName, string unsupportedPropertyTypeName)
        => Diagnostic.Create(UnsupportedTypeForCellValueDescriptor, location, rowTypeName, unsupportedPropertyTypeName);

    private static readonly DiagnosticDescriptor UnsupportedTypeForCellValueDescriptor = new(
        id: "SPCH1002",
        title: "Unsupported type for cell value",
        messageFormat: "The type '{0}' has a property of type '{1}' which is not supported as a cell value. The property will be ignored when creating the row.",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Warning,
        isEnabledByDefault: true);

    // TODO: Class-level diagnostic
    public static DiagnosticInfo DuplicateColumnOrder(LocationInfo? location, string className)
        => new(DuplicateColumnOrderDescriptor, location, new([className]));

    private static readonly DiagnosticDescriptor DuplicateColumnOrderDescriptor = new(
        id: "SPCH1003",
        title: "Duplicate column ordering",
        messageFormat: "The type '{0}' has two or more properties with the same column order",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static Diagnostic InvalidColumnHeaderPropertyReference(Location? location, string propertyName, string typeFullName)
        => Diagnostic.Create(InvalidColumnHeaderPropertyReferenceDescriptor, location, [propertyName, typeFullName]);

    private static readonly DiagnosticDescriptor InvalidColumnHeaderPropertyReferenceDescriptor = new(
        id: "SPCH1004",
        title: "Invalid ColumnHeader property reference",
        messageFormat: "'{0}' on type '{1}' is not a valid property reference. It must be a static property, have a public getter, and the return type must be a string (or string?).",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static Diagnostic UnsupportedTypeForAttribute(Location? location, string attributeName, string typeFullName)
        => Diagnostic.Create(UnsupportedTypeForAttributeDescriptor, location, [attributeName, typeFullName]);

    private static readonly DiagnosticDescriptor UnsupportedTypeForAttributeDescriptor = new(
        id: "SPCH1005",
        title: "Unsupported type for attribute",
        messageFormat: "{0} is not supported on properties of type '{1}'",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static Diagnostic InvalidAttributeArgument(Location? location, string attributeName)
        => Diagnostic.Create(InvalidAttributeArgumentDescriptor, location, [attributeName]);

    private static readonly DiagnosticDescriptor InvalidAttributeArgumentDescriptor = new(
        id: "SPCH1006",
        title: "Invalid attribute argument",
        messageFormat: "Invalid argument for {1}",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static Diagnostic AttributeTypeArgumentMustInherit(Location? location, string typeName, string attributeName, string baseClassName)
        => Diagnostic.Create(AttributeTypeArgumentMustInheritDescriptor, location, [typeName, attributeName, baseClassName]);

    private static readonly DiagnosticDescriptor AttributeTypeArgumentMustInheritDescriptor = new(
        id: "SPCH1007",
        title: "Invalid attribute type argument",
        messageFormat: "Type '{0}' is an invalid argument for {1} because it does not inherit {2}",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static Diagnostic AttributeCombinationNotSupported(Location? location, string attribute1, string attribute2)
        => Diagnostic.Create(AttributeCombinationNotSupportedDescriptor, location, [attribute1, attribute2]);

    private static readonly DiagnosticDescriptor AttributeCombinationNotSupportedDescriptor = new(
        id: "SPCH1008",
        title: "Attribute combination not supported",
        messageFormat: "Having both {0} and {1} on a property is not supported",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static Diagnostic AttributeTypeArgumentMustHaveDefaultConstructor(Location? location, string typeName)
        => Diagnostic.Create(AttributeTypeArgumentMustHaveDefaultConstructorDescriptor, location, [typeName]);

    public static readonly DiagnosticDescriptor AttributeTypeArgumentMustHaveDefaultConstructorDescriptor = new(
        id: "SPCH1009",
        title: "Type must have a public parameterless constructor",
        messageFormat: "Type '{0}' must have a public parameterless constructor",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);
}
