namespace SpreadCheetah.SourceGeneration;

/// <summary>
/// Instructs the SpreadCheetah source generator to specify the order between columns.
/// The source generator creates columns from properties, and using the attribute on a property will set the order for the resulting column.
/// <list type="bullet">
///   <item><description>The order will be in ascending order based on the <c>order</c> parameter. E.g. the column from property with <c>ColumnOrder(1)</c> will be before
///   the column from property with <c>ColumnOrder(2)</c>.</description></item>
///   <item><description>The value for the <c>order</c> parameter can be any integer, including negative integers.</description></item>
///   <item><description>Two different properties on a type can not have the same value for the <c>order</c> parameter.</description></item>
///   <item><description>The values used for a type's properties don't have to be consecutive.</description></item>
///   <item><description>Properties that don't have the attribute will implicitly get the first available value for <c>order</c> starting at 1.</description></item>
/// </list>
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class ColumnOrderAttribute(int order) : Attribute
{
    /// <summary>
    /// Returns the constructor's <c>order</c> argument value.
    /// </summary>
    public int? Order { get; } = order;
}
