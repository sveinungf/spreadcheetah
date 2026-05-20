using SpreadCheetah.Styling;

namespace SpreadCheetah.TestHelpers.Interfaces;

public interface IStyleNumberFormat
{
    string? CustomFormat { get; }

    StandardNumberFormat? StandardFormat { get; }
}
