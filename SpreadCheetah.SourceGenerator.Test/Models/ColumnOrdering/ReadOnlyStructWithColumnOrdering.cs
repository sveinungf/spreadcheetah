using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.ColumnOrdering;

public readonly struct ReadOnlyStructWithColumnOrdering(string firstName, string lastName, decimal gpa, int age)
{
    [ColumnOrder(2)]
    public string FirstName { get; } = firstName;
    [ColumnOrder(1)]
    public string LastName { get; } = lastName;
    public decimal Gpa { get; } = gpa;
    [ColumnOrder(3)]
    public int Age { get; } = age;
}