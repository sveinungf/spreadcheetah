using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.ColumnHeader;

public class ClassWithSpecialCharacterColumnHeaders
{
    [ColumnHeader("First name")]
    public string? FirstName { get; set; }

    [ColumnHeader("")]
    public string? LastName { get; set; }

    [ColumnHeader("Nationality (escaped characters \", \', \\)")]
    public string? Nationality { get; set; }

    [ColumnHeader("Address line 1 (escaped characters \n, \r\n, \t)")]
    public string? AddressLine1 { get; set; }

    [ColumnHeader(@"Address line 2 (verbatim
string: "", \)")]
    public string? AddressLine2 { get; set; }

    [ColumnHeader("""
        Age (
            raw
            string
            literal
        )
    """)]
    public int Age { get; set; }

    [ColumnHeader("Note (unicode escape sequence 🌉, \ud83d\udc4d, \xE7)")]
    public string? Note { get; set; }

    private const string Constant = "This is a constant";

    [ColumnHeader($"Note 2 (constant interpolated string: {Constant})")]
    public string? Note2 { get; set; }
}
