namespace SpreadCheetah.SourceGenerator.Helpers;

internal class GeneratorOptions
{
    public bool SuppressWarnings { get; }

    public GeneratorOptions(bool suppressWarnings)
    {
        SuppressWarnings = suppressWarnings;
    }
}
