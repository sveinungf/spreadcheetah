namespace SpreadCheetah.Styling.Internal;

internal readonly record struct AddedStyle(
    int? AlignmentIndex,
    int? BorderIndex,
    int? FillIndex,
    int? FontIndex,
    int? CustomFormatIndex,
    StandardNumberFormat? StandardFormat,
    string? Name,
    StyleNameVisibility? Visibility);