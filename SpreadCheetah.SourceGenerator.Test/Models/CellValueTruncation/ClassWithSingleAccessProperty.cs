using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellValueTruncation;

public class ClassWithSingleAccessProperty(string theValue)
{
    private int _accessCounter;

    [CellValueTruncate(1)]
    public string Value
    {
        get
        {
            if (_accessCounter > 0)
                throw new InvalidOperationException("The property was accessed more than once");

            ++_accessCounter;
            return theValue;
        }
    }
}
