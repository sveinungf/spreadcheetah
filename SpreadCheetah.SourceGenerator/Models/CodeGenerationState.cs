namespace SpreadCheetah.SourceGenerator.Models;

internal sealed class CodeGenerationState(Dictionary<string, string> cellValueConverters)
{
    public bool DoGenerateConstructNullableFormulaCell { get; private set; }
    public bool DoGenerateConstructTruncatedDataCell { get; private set; }

    public void RequireConstructNullableFormulaCell() => DoGenerateConstructNullableFormulaCell = true;
    public void RequireConstructTruncatedDataCell() => DoGenerateConstructTruncatedDataCell = true;
    public string GetValueConverter(string typeName) => cellValueConverters[typeName];
}
