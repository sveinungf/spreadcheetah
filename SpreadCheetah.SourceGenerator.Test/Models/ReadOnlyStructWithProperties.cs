namespace SpreadCheetah.SourceGenerator.Test.Models
{
    public readonly struct ReadOnlyStructWithProperties
    {
        public string FirstName { get; }
        public string LastName { get; }
        public int Age { get; }

        public ReadOnlyStructWithProperties(string firstName, string lastName, int age)
        {
            FirstName = firstName;
            LastName = lastName;
            Age = age;
        }
    }
}
