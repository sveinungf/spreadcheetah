using SpreadCheetah.CellReferences;
using SpreadCheetah.Helpers;
using SpreadCheetah.MetadataXml.Attributes;

namespace SpreadCheetah.ConditionalFormatting.Internal;

internal sealed record ImmutableMatchesFormulaFormatRule : ImmutableConditionalFormatRule
{
    public Formula Formula { get; init; }

    private BufferWriteProgress _progress;

    public override bool TryWrite(SpreadsheetBuffer buffer, int priority, SimpleSingleCellReference topLeftCell)
    {
        var dxfIdAttribute = new IntAttribute("dxfId"u8, StyleDxfId);
        var priorityAttribute = new IntAttribute("priority"u8, priority);
        var formulaText = Formula.GetFormulaText((int)topLeftCell.Row, topLeftCell.Column);

        var success = buffer.TryWrite(
            _progress, out _progress,
            $"{"<cfRule type=\"expression\""u8}" +
            $"{dxfIdAttribute}" +
            $"{priorityAttribute}" +
            $"{"><formula>"u8}" +
            $"{formulaText}" +
            $"{"</formula></cfRule>"u8}");

        if (success)
            _progress = default;

        return success;
    }
}
