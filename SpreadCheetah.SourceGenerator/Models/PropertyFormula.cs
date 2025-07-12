namespace SpreadCheetah.SourceGenerator.Models;

internal readonly record struct PropertyFormula(FormulaType Type);

internal enum FormulaType
{
    GeneralNullable,
    GeneralNonNullable,
    HyperlinkFromUri
}