using ClosedXML.Excel;
using SpreadCheetah.TestHelpers.Interfaces;
using SpreadCheetah.TestHelpers.Models;
using System.Globalization;

namespace SpreadCheetah.TestHelpers.Implementations;

internal sealed class ClosedXmlConditionalFormatRule(
    IXLConditionalFormat conditionalFormat)
    : IConditionalFormatRule
{
    public bool IsCellIsRule => conditionalFormat.ConditionalFormatType is XLConditionalFormatType.CellIs;
    public bool IsDuplicateValuesRule => conditionalFormat.ConditionalFormatType is XLConditionalFormatType.IsDuplicate;
    public bool IsMatchesFormulaRule => conditionalFormat.ConditionalFormatType is XLConditionalFormatType.Expression;
    public bool IsUniqueValuesRule => conditionalFormat.ConditionalFormatType is XLConditionalFormatType.IsUnique;
    public IStyle Style => ClosedXmlStyle.Create(conditionalFormat.Style);

    public string CellRangeReference
    {
        get
        {
            var rangeAddress = conditionalFormat.Range.RangeAddress;
            return rangeAddress.NumberOfCells == 1
                ? rangeAddress.FirstAddress.ToStringRelative()
                : rangeAddress.ToStringRelative();
        }
    }

    public string Formula
    {
        get
        {
            if (!IsMatchesFormulaRule)
                throw new InvalidOperationException("The conditional format rule is not a formula rule.");

            var value = conditionalFormat.Values.Values.SingleOrDefault();
            if (value is not { IsFormula: true, Value.Length: > 0 })
                throw new InvalidOperationException("The conditional format rule does not have a formula value.");

            return value.Value;
        }
    }

    public double NumericValue => NumericValues.Single();
    public IReadOnlyList<double> NumericValues => field ??= CreateNumericValues();

    private List<double> CreateNumericValues()
    {
        if (!IsCellIsRule)
            throw new InvalidOperationException("The conditional format rule is not a cellIs rule.");

        var values = conditionalFormat.Values.Values;
        if (values.Any(x => x.Value.Length == 0))
            throw new InvalidOperationException("The conditional format rule does not have a numeric value.");

        return [.. values.Select(x => double.Parse(x.Value, CultureInfo.InvariantCulture))];
    }

    public ConditionalFormatOperator RuleOperator => !IsCellIsRule
        ? throw new InvalidOperationException("The conditional format rule is not a cellIs rule.")
        : Map(conditionalFormat.Operator);

    private static ConditionalFormatOperator Map(XLCFOperator op)
    {
        return op switch
        {
            XLCFOperator.Equal => ConditionalFormatOperator.EqualTo,
            XLCFOperator.NotEqual => ConditionalFormatOperator.NotEqualTo,
            XLCFOperator.GreaterThan => ConditionalFormatOperator.GreaterThan,
            XLCFOperator.LessThan => ConditionalFormatOperator.LessThan,
            XLCFOperator.EqualOrGreaterThan => ConditionalFormatOperator.GreaterThanOrEqualTo,
            XLCFOperator.EqualOrLessThan => ConditionalFormatOperator.LessThanOrEqualTo,
            XLCFOperator.Between => ConditionalFormatOperator.Between,
            XLCFOperator.NotBetween => ConditionalFormatOperator.NotBetween,
            XLCFOperator.Contains => ConditionalFormatOperator.Contains,
            XLCFOperator.NotContains => ConditionalFormatOperator.NotContains,
            XLCFOperator.StartsWith => ConditionalFormatOperator.StartsWith,
            XLCFOperator.EndsWith => ConditionalFormatOperator.EndsWith,
            _ => throw new ArgumentOutOfRangeException(nameof(op), op, null)
        };
    }
}
