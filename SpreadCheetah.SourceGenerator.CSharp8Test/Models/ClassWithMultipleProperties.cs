namespace SpreadCheetah.SourceGenerator.CSharp8Test.Models
{
    public class ClassWithMultipleProperties
    {
        public string FirstName { get; }
        public string LastName { get; }
        public int Age { get; }

        public ClassWithMultipleProperties(string firstName, string lastName, int age)
        {
            FirstName = firstName;
            LastName = lastName;
            Age = age;
        }
    }
}
