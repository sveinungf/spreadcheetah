using SpreadCheetah.Helpers;
using System.Text;

namespace SpreadCheetah.Validations;

public abstract class BaseValidation
{
    internal bool HasValue1 { get; set; }
    internal bool HasValue2 { get; set; }
    internal ValidationOperator Operator { get; set; }
    internal abstract ValidationType Type { get; }
    internal abstract void AppendValue1(StringBuilder sb);
    internal abstract void AppendValue2(StringBuilder sb);

    public bool IgnoreBlank { get; set; } = true;
    public bool ShowErrorAlert { get; set; } = true;
    public bool ShowInputMessage { get; set; } = true;
    public string? ErrorTitle { get; set; } // TODO: Validate length
    public string? ErrorMessage { get; set; } // TODO: Validate length
    public string? InputTitle { get; set; } // TODO: Validate length
    public string? InputMessage { get; set; } // TODO: Validate length

    public ValidationErrorType ErrorType
    {
        get => _errorType;
        set => _errorType = EnumHelper.IsDefined(value)
            ? value
            : throw new ArgumentOutOfRangeException(nameof(value), value, null);
    }

    private ValidationErrorType _errorType;
}
