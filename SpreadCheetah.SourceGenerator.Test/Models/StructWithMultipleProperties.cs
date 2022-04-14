namespace SpreadCheetah.SourceGenerator.Test.Models;

public struct StructWithMultipleProperties
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
