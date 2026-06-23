using SpreadCheetah.Formulas;

namespace SpreadCheetah.Test.Tests;

public static class R1C1FormulaConverterTests
{
    [Theory]
    // Absolute cell references (independent of anchor)
    [InlineData("R1C1", 1, 1, "$A$1")]
    [InlineData("R5C3", 1, 1, "$C$5")]
    [InlineData("R5C3", 9, 9, "$C$5")]
    // Relative cell references (depend on anchor)
    [InlineData("RC", 3, 2, "B3")]
    [InlineData("RC[-1]", 3, 2, "A3")]
    [InlineData("R[-1]C", 3, 2, "B2")]
    [InlineData("R[1]C[1]", 3, 2, "C4")]
    [InlineData("R[10]C[5]", 1, 1, "F11")]
    // Explicit plus sign in relative offset
    [InlineData("R[+1]C[+1]", 3, 2, "C4")]
    // Lowercase r/c are treated the same as uppercase
    [InlineData("rc[-1]", 3, 2, "A3")]
    [InlineData("r1c1", 1, 1, "$A$1")]
    // Mixed references
    [InlineData("R1C[1]", 3, 2, "C$1")]
    [InlineData("R[1]C1", 3, 2, "$A4")]
    // Column names beyond 'Z'
    [InlineData("R1C27", 1, 1, "$AA$1")]
    // Whole row references
    [InlineData("R5", 3, 2, "$5:$5")]
    [InlineData("R[2]", 3, 2, "5:5")]
    [InlineData("R", 3, 2, "3:3")]
    // Whole column references
    [InlineData("C3", 3, 2, "$C:$C")]
    [InlineData("C[-1]", 3, 2, "A:A")]
    [InlineData("C", 3, 2, "B:B")]
    // Ranges
    [InlineData("R1C1:R2C2", 1, 1, "$A$1:$B$2")]
    [InlineData("RC:R[1]C[1]", 3, 2, "B3:C4")]
    [InlineData("R1:R3", 1, 1, "$1:$3")]
    [InlineData("C1:C3", 1, 1, "$A:$C")]
    public static void R1C1FormulaConverter_ToA1_Reference(string formula, int row, int column, string expected)
    {
        // Act
        var actual = R1C1FormulaConverter.ToA1(formula, row, column);

        // Assert
        Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData("SUM(RC[-2]:RC[-1])", 1, 3, "SUM(A1:B1)")]
    [InlineData("RC[-1]*RC[-2]", 5, 3, "B5*A5")]
    [InlineData("IF(RC[-1]>0,\"Yes\",\"No\")", 2, 2, "IF(A2>0,\"Yes\",\"No\")")]
    [InlineData("Sheet1!R1C1", 5, 5, "Sheet1!$A$1")]
    [InlineData("'My Sheet'!RC[-1]", 3, 2, "'My Sheet'!A3")]
    // Escaped single quote inside a quoted sheet name
    [InlineData("'O''Brien'!R1C1", 5, 5, "'O''Brien'!$A$1")]
    // An unquoted sheet name that itself looks like a reference must be left intact
    [InlineData("C5!RC[-1]", 3, 2, "C5!A3")]
    // A reference followed by ':' but not by another reference is not turned into a range
    [InlineData("R1C1:", 1, 1, "$A$1:")]
    [InlineData("R1C1:foo", 1, 1, "$A$1:foo")]
    public static void R1C1FormulaConverter_ToA1_WithinExpression(string formula, int row, int column, string expected)
    {
        // Act
        var actual = R1C1FormulaConverter.ToA1(formula, row, column);

        // Assert
        Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData("", 1, 1, "")]
    // Quoted text that looks like references must not be converted
    [InlineData("\"RC[-1]\"", 1, 1, "\"RC[-1]\"")]
    [InlineData("\"R1C1\"", 1, 1, "\"R1C1\"")]
    // Escaped double quotes inside a string literal
    [InlineData("\"Total \"\"RC\"\" here\"", 1, 1, "\"Total \"\"RC\"\" here\"")]
    // Unterminated string literal is left as-is
    [InlineData("\"unterminated", 1, 1, "\"unterminated")]
    // Function names and identifiers containing R/C must not be converted
    [InlineData("ROUND(RC[-1],2)", 1, 3, "ROUND(B1,2)")]
    [InlineData("COUNT(RC)", 2, 2, "COUNT(B2)")]
    // An R/C that is part of a longer identifier must not be converted
    [InlineData("FooR1C1", 1, 1, "FooR1C1")]
    // Malformed references are passed through unchanged
    [InlineData("R[1", 3, 2, "R[1")]
    [InlineData("RC[", 3, 2, "RC[")]
    public static void R1C1FormulaConverter_ToA1_DoesNotConvertNonReferences(string formula, int row, int column, string expected)
    {
        // Act
        var actual = R1C1FormulaConverter.ToA1(formula, row, column);

        // Assert
        Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData("RC[-1]", 1, 1)] // Column 0 does not exist
    [InlineData("R[-1]C", 1, 1)] // Row 0 does not exist
    [InlineData("RC[1]", 1, 16384)] // Beyond last column
    [InlineData("R[1]C", 1048576, 1)] // Beyond last row
    [InlineData("R1C16385", 1, 1)] // Absolute column beyond the maximum
    [InlineData("R10000000C1", 1, 1)] // Absolute row beyond the maximum (number overflow cap)
    public static void R1C1FormulaConverter_ToA1_OutOfBoundsThrows(string formula, int row, int column)
    {
        // Act & Assert
        Assert.Throws<ArgumentException>(() => R1C1FormulaConverter.ToA1(formula, row, column));
    }
}
