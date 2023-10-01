namespace SpreadCheetah.SourceGenerator.Test.Models;

public class ClassWithAllSupportedTypes
{
    public string StringValue { get; set; } = "";
    public string? NullableStringValue { get; set; }
    public int IntValue { get; set; }
    public int? NullableIntValue { get; set; }
    public long LongValue { get; set; }
    public long? NullableLongValue { get; set; }
    public float FloatValue { get; set; }
    public float? NullableFloatValue { get; set; }
    public double DoubleValue { get; set; }
    public double? NullableDoubleValue { get; set; }
    public decimal DecimalValue { get; set; }
    public decimal? NullableDecimalValue { get; set; }
    public DateTime DateTimeValue { get; set; }
    public DateTime? NullableDateTimeValue { get; set; }
    public bool BoolValue { get; set; }
    public bool? NullableBoolValue { get; set; }
}
