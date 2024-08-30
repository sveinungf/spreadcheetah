using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.CSharp8Test.Models
{
    [InheritColumns]
    public class ClassWithMultipleProperties : BaseClass
    {
        [ColumnHeader("Last name")]
        [CellStyle("Last name style")]
        [CellValueTruncate(20)]
        public string LastName { get; }

        [ColumnOrder(1)]
        public string FirstName { get; }

        [ColumnWidth(5)]
        public int Age { get; }

        public ClassWithMultipleProperties(string id, string firstName, string lastName, int age)
            : base(id)
        {
            FirstName = firstName;
            LastName = lastName;
            Age = age;
        }
    }
}
