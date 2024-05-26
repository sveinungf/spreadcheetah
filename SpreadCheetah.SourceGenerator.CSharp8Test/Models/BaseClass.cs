namespace SpreadCheetah.SourceGenerator.CSharp8Test.Models
{
    public class BaseClass
    {
        public string Id { get; }

        protected BaseClass(string id)
        {
            Id = id;
        }
    }
}
