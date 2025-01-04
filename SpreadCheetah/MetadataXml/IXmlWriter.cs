namespace SpreadCheetah.MetadataXml;

internal interface IXmlWriter<out T> where T : struct
{
    bool Current { get; }
    bool MoveNext();
    T GetEnumerator();
}
