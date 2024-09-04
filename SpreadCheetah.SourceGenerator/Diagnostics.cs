using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator;

internal static class Diagnostics
{
    private const string Category = "SpreadCheetah.SourceGenerator";

    public static readonly DiagnosticDescriptor NoPropertiesFound = new(
        id: "SPCH1001",
        title: "Missing properties with public getters",
        messageFormat: "The type '{0}' has no properties with public getters. This will cause an empty row to be added.",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Warning,
        isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor UnsupportedTypeForCellValue = new(
        id: "SPCH1002",
        title: "Unsupported type for cell value",
        messageFormat: "The type '{0}' has a property of type '{1}' which is not supported as a cell value. The property will be ignored when creating the row.",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Warning,
        isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor DuplicateColumnOrder = new(
        id: "SPCH1003",
        title: "Duplicate column ordering",
        messageFormat: "The type '{0}' has two or more properties with the same column order",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor InvalidColumnHeaderPropertyReference = new(
        id: "SPCH1004",
        title: "Invalid ColumnHeader property reference",
        messageFormat: "'{0}' on type '{1}' is not a valid property reference. It must be a static property, have a public getter, and the return type must be a string (or string?).",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor UnsupportedTypeForCellValueLengthLimit = new(
        id: "SPCH1005",
        title: "Unsupported type for CellValueLengthLimit attribute",
        messageFormat: "The CellValueLengthLimit attribute is not supported on properties of type '{0}'",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor InvalidAttributeArgument = new(
        id: "SPCH1006",
        title: "Invalid attribute argument",
        messageFormat: "'{0}' is an invalid argument for attribute '{1}'",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);
    
    public static readonly DiagnosticDescriptor CellValueConverterTypeNotInheritCellValueConverter = new(
        id: "SPCH1007",
        title: "The type provided for CellValueConverterAttribute must inherit СellValueConverter<> class",
        messageFormat: "'{0}' is not inherit PropertyCellValueConverter<>",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);
    
    public static readonly DiagnosticDescriptor 
        UnsupportedCellValueConverterAttributeWithCellValueTruncateAttributeTogether = new(
        id: "SPCH1008",
        title: "The type has CellValueConverterAttribute and TruncateValueAttribute which is not supported",
        messageFormat: "'{0}' has CellValueConverterAttribute and TruncateValueAttribute, only one of this attribute on property allowed at once",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor CellValueConverterWithoutPublicParameterlessConstructor = new(
        id: "SPCH1009",
        title: "The type must have a public parameterless constructor",
        messageFormat: "'{0}' inherit CellValueConverter but doesn't have a public parameterless constructor",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);
    
    public static readonly DiagnosticDescriptor CellValueConverterArgumentTypeNotSameAsPropertyType = new(
        id: "SPCH1010",
        title: "CellValueConverter generic different from the property type",
        messageFormat: "'{0}' has different type that property type",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);
}
