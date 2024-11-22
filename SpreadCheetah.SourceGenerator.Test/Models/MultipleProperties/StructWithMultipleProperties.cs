namespace SpreadCheetah.SourceGenerator.Test.Models.MultipleProperties;

#pragma warning disable IDE0250 // Make struct 'readonly'
public struct StructWithMultipleProperties
#pragma warning restore IDE0250 // Make struct 'readonly'
{
    public string FirstName { get; }
    public string LastName { get; }
    public int Age { get; }

    public StructWithMultipleProperties(string firstName, string lastName, int age)
    {
        FirstName = firstName;
        LastName = lastName;
        Age = age;
    }
}
