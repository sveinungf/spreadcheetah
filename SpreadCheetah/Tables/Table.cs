namespace SpreadCheetah.Tables;

public sealed class Table
{
    internal string Name { get; }
    internal TableStyle Style { get; }

    public bool AutoFilter { get; set; } = true;
    public bool BandedColumns { get; set; }
    public bool BandedRows { get; set; } = true;
    public int? NumberOfColumns { get; set; } // TODO: Validate in setter

    public Table(string name, TableStyle style)
    {
        Name = name; // TODO: Validate
        Style = style; // TODO: Validate
    }

    public TableColumnOptions Column(int columnNumber)
    {
        // TODO
        throw new NotImplementedException();
    }
}
