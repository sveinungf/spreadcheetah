using System.Text;

namespace SpreadCheetah.MetadataXml;

internal static class CellMergeXml
{
    public static async ValueTask WriteAsync(
        Stream stream,
        SpreadsheetBuffer buffer,
        HashSet<CellReference> cellMerges,
        CancellationToken token)
    {
        var sb = new StringBuilder("<mergeCells count=\"");
        sb.Append(cellMerges.Count);
        sb.Append("\">");

        foreach (var cellMerge in cellMerges)
        {
            sb.Append("<mergeCell ref=\"");
            sb.Append(cellMerge.Reference);
            sb.Append("\"/>");
        }

        sb.Append("</mergeCells>");
        await buffer.WriteStringAsync(sb, stream, token).ConfigureAwait(false);
    }
}
