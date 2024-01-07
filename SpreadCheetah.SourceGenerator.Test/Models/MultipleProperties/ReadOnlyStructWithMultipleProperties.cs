namespace SpreadCheetah.SourceGenerator.Test.Models.MultipleProperties;

public readonly struct ReadOnlyStructWithMultipleProperties
{
    public string FirstName { get; }
    public string LastName { get; }
    public int Age { get; }

    public ReadOnlyStructWithMultipleProperties(string firstName, string lastName, int age)
    {
        FirstName = firstName;
        LastName = lastName;
        Age = age;
    }
}
