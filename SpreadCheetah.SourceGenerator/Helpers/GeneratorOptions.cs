namespace SpreadCheetah.SourceGenerator.Helpers;

internal sealed class GeneratorOptions
{
    public bool SuppressWarnings { get; }

    public GeneratorOptions(bool suppressWarnings)
    {
        SuppressWarnings = suppressWarnings;
    }
}
