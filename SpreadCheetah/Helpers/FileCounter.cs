using SpreadCheetah.Images;
using SpreadCheetah.Images.Internal;
using SpreadCheetah.Tables.Internal;
using System.Diagnostics;

namespace SpreadCheetah.Helpers;

internal sealed class FileCounter
{
    public int? CurrentWorksheetDrawingsFileIndex { get; private set; }
    public int? CurrentWorksheetNotesFileIndex { get; private set; }
    public int? CurrentWorksheetTableFileStartIndex { get; private set; }
    public int CurrentWorksheetTableCount { get; private set; }

    public int WorksheetsWithImages { get; private set; }
    public int WorksheetsWithNotes { get; private set; }
    public int TotalTables => TableNames?.Count ?? 0;
    public int TotalEmbeddedImages { get; private set; }
    public int TotalAddedImages { get; set; }
    public EmbeddedImageTypes EmbeddedImageTypes { get; private set; }

    private HashSet<string>? TableNames { get; set; }

    public void AddEmbeddedImage(ImageType type)
    {
        Debug.Assert(type > ImageType.None && EnumPolyfill.IsDefined(type));

        if (type == ImageType.Png)
            EmbeddedImageTypes |= EmbeddedImageTypes.Png;
        else if (type == ImageType.Jpeg)
            EmbeddedImageTypes |= EmbeddedImageTypes.Jpeg;

        TotalEmbeddedImages++;
    }

    public void ImageForCurrentWorksheet()
    {
        if (CurrentWorksheetDrawingsFileIndex is null)
        {
            ++WorksheetsWithImages;
            CurrentWorksheetDrawingsFileIndex = WorksheetsWithImages;
        }
    }

    public void NoteForCurrentWorksheet()
    {
        if (CurrentWorksheetNotesFileIndex is null)
        {
            ++WorksheetsWithNotes;
            CurrentWorksheetNotesFileIndex = WorksheetsWithNotes;
        }
    }

    public string AssignTableNameForCurrentWorksheet()
    {
        var tableNames = TableNames ??= new(StringComparer.OrdinalIgnoreCase);
        var tableName = TableNameGenerator.GenerateUniqueTableName(tableNames);
        var result = TryAssignTableNameForCurrentWorksheet(tableName);
        Debug.Assert(result);
        return tableName;
    }

    public bool TryAssignTableNameForCurrentWorksheet(string tableName)
    {
        var tableNames = TableNames ??= new(StringComparer.OrdinalIgnoreCase);

        if (!tableNames.Add(tableName))
            return false;

        CurrentWorksheetTableCount++;
        CurrentWorksheetTableFileStartIndex ??= tableNames.Count;
        return true;
    }

    public bool CurrentWorksheetHasRelationships =>
        CurrentWorksheetDrawingsFileIndex is not null ||
        CurrentWorksheetNotesFileIndex is not null ||
        CurrentWorksheetTableCount > 0;

    public void ResetCurrentWorksheet()
    {
        CurrentWorksheetDrawingsFileIndex = null;
        CurrentWorksheetNotesFileIndex = null;
        CurrentWorksheetTableCount = 0;
        CurrentWorksheetTableFileStartIndex = null;
    }
}
