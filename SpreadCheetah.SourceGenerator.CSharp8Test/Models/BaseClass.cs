using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.CSharp8Test.Models
{
    public class BaseClass
    {
        [CellStyle("Id style")]
        [CellValueConverter(typeof(RemoveDashesValueConverter))]
        public string Id { get; }

        protected BaseClass(string id)
        {
            Id = id;
        }
    }
}
