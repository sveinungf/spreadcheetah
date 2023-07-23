namespace SpreadCheetah.Worksheets;

internal readonly record struct WorksheetMetadata(
    string Name,
    string Path, // TODO: Just store an integer?
    WorksheetVisibility Visibility,
    int? NotesFileIndex);
