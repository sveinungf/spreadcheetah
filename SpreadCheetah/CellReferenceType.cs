namespace SpreadCheetah;

[Flags]
internal enum CellReferenceType
{
    None = 0,
    Relative = 1,
    Absolute = 2,
    RelativeOrAbsolute = Relative | Absolute
}
