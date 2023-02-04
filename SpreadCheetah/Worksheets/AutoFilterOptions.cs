using System;
using System.Text.RegularExpressions;
using SpreadCheetah.Helpers;

namespace SpreadCheetah.Worksheets;

/// <summary>
/// Provides auto filter options to be used with <see cref="WorksheetOptions"/>.
/// </summary>
public class AutoFilterOptions
{
    private static Regex Regex { get; } = new Regex(@"^([A-Z]+)(\d+):(\1\d+|[A-Z]+\2)$", RegexOptions.None, TimeSpan.FromSeconds(1));
    
    /// <summary>
    /// The range of columns and rows to filter
    /// </summary>
    public string Range
    {
        get => _range;
        set => _range = Regex.Matches(value).Count == 0
            ? throw new ArgumentException("Range must match Excel format of A1:Z26")
            : value;
    }

    public void SetRangeWithHelper(int? rowStart, int columnStart, int? columnEnd, int? rowEnd) =>
        _range = RangeHelper.GetRange(rowStart, columnStart, columnEnd, rowEnd);

    private string _range;

    ///<summary>
    /// Enable auto filtering
    /// </summary>
    public bool Enabled
    {
        get => _enabled;
        set => _enabled = value;
    }

    private bool _enabled;
}