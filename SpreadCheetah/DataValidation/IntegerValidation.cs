using System.Text;

namespace SpreadCheetah.DataValidation;

public sealed class IntegerValidation : BaseValidation
{
    internal int Value1 { get; }
    internal int Value2 { get; }
    internal override ValidationType Type => ValidationType.Integer;

    private IntegerValidation(int value1, ValidationOperator validationOperator)
    {
        Value1 = value1;
        HasValue1 = true;
        Operator = validationOperator;
    }

    private IntegerValidation(int value1, int value2, ValidationOperator validationOperator)
        : this(value1, validationOperator)
    {
        Value2 = value2;
        HasValue2 = true;
    }

    internal override void AppendValue1(StringBuilder sb) => sb.Append(Value1);
    internal override void AppendValue2(StringBuilder sb) => sb.Append(Value2);

    public static IntegerValidation Between(int minimum, int maximum)
    {
        // TODO: Validate arguments
        return new IntegerValidation(minimum, maximum, ValidationOperator.Between);
    }

    public static IntegerValidation NotBetween(int minimum, int maximum)
    {
        // TODO: Validate arguments
        return new IntegerValidation(minimum, maximum, ValidationOperator.NotBetween);
    }
}
