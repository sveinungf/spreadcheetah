namespace SpreadCheetah.SourceGenerator.Test.Models
{
    public struct StructWithProperties
    {
        public string FirstName { get; }
        public string LastName { get; }
        public int Age { get; }

        public StructWithProperties(string firstName, string lastName, int age)
        {
            FirstName = firstName;
            LastName = lastName;
            Age = age;
        }
    }
}
