using Microsoft.CodeAnalysis;
using System.Collections.Generic;
using System.Linq;

namespace SpreadCheetah.SourceGenerator
{
    internal class ClassPropertiesInfo
    {
        public ITypeSymbol ClassType { get; }
        public List<string> PropertyNames { get; }

        public ClassPropertiesInfo(ITypeSymbol classType)
        {
            ClassType = classType;
            PropertyNames = GetPropertyNames(classType).ToList();
        }

        private static IEnumerable<string> GetPropertyNames(ITypeSymbol classType)
        {
            // TODO: Only take allowed types
            foreach (var member in classType.GetMembers())
            {
                if (member is IPropertySymbol
                    {
                        DeclaredAccessibility: Accessibility.Public,
                        IsStatic: false,
                        IsWriteOnly: false
                    })
                {
                    yield return member.Name;
                }
            }
        }
    }
}
