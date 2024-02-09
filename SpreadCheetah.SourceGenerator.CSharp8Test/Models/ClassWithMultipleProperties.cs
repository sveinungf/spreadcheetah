using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.CSharp8Test.Models
{
    public class ClassWithMultipleProperties
    {
        [ColumnHeader("Last name")]
        public string LastName { get; }
        [ColumnOrder(1)]
        public string FirstName { get; }
        public int Age { get; }

        public ClassWithMultipleProperties(string firstName, string lastName, int age)
        {
            FirstName = firstName;
            LastName = lastName;
            Age = age;
        }
    }
}
