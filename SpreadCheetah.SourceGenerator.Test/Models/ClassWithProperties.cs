namespace SpreadCheetah.SourceGenerator.Test.Models;

public class ClassWithProperties
{
    public string FirstName { get; }
    public string LastName { get; }
    public int Age { get; }

    public ClassWithProperties(string firstName, string lastName, int age)
    {
        FirstName = firstName;
        LastName = lastName;
        Age = age;
    }
}
