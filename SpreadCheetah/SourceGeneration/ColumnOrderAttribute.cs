namespace SpreadCheetah.SourceGeneration;

/// <summary>
/// Instructs the SpreadCheetah source generator to specify column ordering.
/// The order can be customized by using this attribute on a type's properties.
/// <list type="bullet">
///   <item><description>The order will be in ascending order based on the <c>order</c> parameter. E.g. a property with <c>ColumnOrder(1)</c> will be in the column before the column for property with <c>ColumnOrder(2)</c>.</description></item>
///   <item><description>The value for the <c>order</c> parameter can be any integer, including negative integers.</description></item>
///   <item><description>Two properties can not have the same value for the <c>order</c> parameter.</description></item>
///   <item><description>Properties that don't have the attribute will implicitly get the first non-duplicate value for <c>order</c> starting at 1.</description></item>
/// </list>
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class ColumnOrderAttribute(int order) : Attribute;
