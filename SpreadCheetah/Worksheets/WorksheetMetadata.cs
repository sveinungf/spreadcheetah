namespace SpreadCheetah.Worksheets;

internal readonly record struct WorksheetMetadata(
    string Name,
    string Path,
    WorksheetVisibility Visibility,
    int? NotesFileIndex);
