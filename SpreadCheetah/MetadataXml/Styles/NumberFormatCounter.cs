namespace SpreadCheetah.MetadataXml.Styles;

internal sealed class NumberFormatCounter
{
    // Custom number formats start at 165, after the built-in number formats.
    private int _currentId = 165;

    public int GetNextId() => _currentId++;
}
