namespace SpreadCheetah
{
    public readonly struct Formula
    {
        public string? FormulaText { get; }

        public Formula(string? formulaText) => FormulaText = formulaText;
    }
}
