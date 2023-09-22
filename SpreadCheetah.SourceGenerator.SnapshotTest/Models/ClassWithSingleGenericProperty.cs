namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models;

public class ClassWithSingleGenericProperty<T>
{
    public T Value { get; set; }

    public ClassWithSingleGenericProperty(T value) => Value = value;
}
