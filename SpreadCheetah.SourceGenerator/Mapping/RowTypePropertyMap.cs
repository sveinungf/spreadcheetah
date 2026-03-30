using Microsoft.CodeAnalysis;
using SpreadCheetah.SourceGenerator.Extensions;
using SpreadCheetah.SourceGenerator.Models;

namespace SpreadCheetah.SourceGenerator.Mapping;

internal static class RowTypePropertyMap
{
    public static RowTypeProperty ToRowTypeProperty(
        this PropertyAttributeData data,
        IPropertySymbol propertySymbol)
    {
        return new RowTypeProperty
        {
            CellFormat = data.CellFormat,
            CellStyle = data.CellStyle,
            CellValueConverter = data.CellValueConverter,
            CellValueTruncate = data.CellValueTruncate,
            ColumnHeader = data.ColumnHeader?.ToColumnHeaderInfo(),
            ColumnWidth = data.ColumnWidth,
            Formula = data.CellValueConverter is null ? propertySymbol.Type.ToPropertyFormula() : null,
            Name = propertySymbol.Name
        };
    }

    private static ColumnHeaderInfo ToColumnHeaderInfo(this ColumnHeader columnHeader)
    {
        var fullPropertyReference = columnHeader.PropertyReference is { } reference
            ? $"{reference.TypeFullName}.{reference.PropertyName}"
            : null;

        return new ColumnHeaderInfo(
            columnHeader.RawString,
            fullPropertyReference);
    }
}
