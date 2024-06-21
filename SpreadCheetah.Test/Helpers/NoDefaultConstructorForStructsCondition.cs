using ArchUnitNET.Domain;
using ArchUnitNET.Domain.Extensions;
using ArchUnitNET.Fluent.Conditions;

namespace SpreadCheetah.Test.Helpers;

internal sealed class NoDefaultConstructorForStructsCondition : ICondition<IType>
{
    public string Description => "not have a default constructor";

    public IEnumerable<ConditionResult> Check(IEnumerable<IType> objects, Architecture architecture)
    {
        foreach (var obj in objects)
        {
            yield return Check(obj);
        }
    }

    private static ConditionResult Check(IType type)
    {
        var pass = true;

        if (type is Struct structType)
        {
            foreach (var constructor in structType.GetConstructors())
            {
                if (constructor.IsStatic is true)
                    continue;

                if (!constructor.Parameters.Any())
                {
                    pass = false;
                    break;
                }
            }
        }

        return new ConditionResult(type, pass: pass, failDescription: pass ? null : "has a default constructor");
    }

    public bool CheckEmpty() => true;
}
