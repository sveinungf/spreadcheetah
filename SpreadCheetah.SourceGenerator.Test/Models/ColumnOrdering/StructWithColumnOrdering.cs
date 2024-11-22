using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.ColumnOrdering;

#pragma warning disable IDE0250 // Make struct 'readonly'
public struct StructWithColumnOrdering(string firstName, string lastName, decimal gpa, int age)
#pragma warning restore IDE0250 // Make struct 'readonly'
{
    [ColumnOrder(2)]
    public string FirstName { get; } = firstName;
    [ColumnOrder(1)]
    public string LastName { get; } = lastName;
    public decimal Gpa { get; } = gpa;
    [ColumnOrder(3)]
    public int Age { get; } = age;
}