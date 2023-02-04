using System.Net;
using System.Text;

namespace SpreadCheetah.MetadataXml;

internal static class AutoFilterXml
{
    public static async ValueTask WriteAsync(
        Stream stream,
        SpreadsheetBuffer buffer,
        string filterRange,
        CancellationToken token)
    {
        var sb = new StringBuilder("<autoFilter ");

        sb.AppendTextAttribute("ref", filterRange);

        sb.Append("/>");
        await buffer.WriteStringAsync(sb, stream, token).ConfigureAwait(false);
    }

    private static void AppendTextAttribute(this StringBuilder sb, string attribute, string value)
    {
        sb.Append(attribute);
        sb.Append("=\"");
        sb.Append(WebUtility.HtmlEncode(value));
        sb.Append("\" ");
    }
}