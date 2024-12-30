namespace SpreadCheetah.Tables;

public sealed class Table
{
    internal string Name { get; }
    internal TableStyle Style { get; }

    public bool AutoFilter { get; set; } = true; // TODO: Maybe bool is not enough, if one would want to do pre-defined filtering.
    public bool BandedColumns { get; set; }
    public bool BandedRows { get; set; } = true;
    public int? NumberOfColumns { get; set; } // TODO: Validate in setter. Should only be used if number of columns is different from header row columns.

    public Table(string name, TableStyle style)
    {
        Name = name; // TODO: Validate
        Style = style; // TODO: Validate
    }

    public TableColumnOptions Column(int columnNumber)
    {
        // TODO: Implement
        // TODO: Consider having AutoFilter stuff in TableColumnOptions
        throw new NotImplementedException();
    }
}
