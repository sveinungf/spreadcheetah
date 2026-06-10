namespace SpreadCheetah.MetadataXml.Worksheets;

internal sealed class ConditionalFormatPriorityCounter
{
    private int _currentPriority = 1;
    public int IncrementPriority() => _currentPriority++;
    public int DecrementPriority() => _currentPriority--;
}
