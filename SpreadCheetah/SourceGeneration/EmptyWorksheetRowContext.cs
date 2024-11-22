using SpreadCheetah.Styling;

#if !NET6_0_OR_GREATER
using ArgumentNullException = SpreadCheetah.Helpers.Backporting.ArgumentNullExceptionBackport;
#endif

namespace SpreadCheetah.SourceGeneration;

/// <summary>
/// Used by the SpreadCheetah source generator for types without valid publicly visible getters.
/// Should not be used directly.
/// </summary>
public static class EmptyWorksheetRowContext
{
    /// <summary>
    /// Creates metadata for a type without valid publicly visible getters.
    /// Should not be used directly.
    /// </summary>
    public static WorksheetRowTypeInfo<T> CreateTypeInfo<T>() => WorksheetRowMetadataServices.CreateObjectInfo<T>(AddHeaderRowAsync, AddAsRowAsync, AddRangeAsRowsAsync);

    private static ValueTask AddHeaderRowAsync(Spreadsheet spreadsheet, StyleId? _, CancellationToken token)
    {
        return spreadsheet.AddRowAsync(ReadOnlyMemory<DataCell>.Empty, token);
    }

    private static ValueTask AddAsRowAsync<T>(Spreadsheet spreadsheet, T _, CancellationToken token)
    {
        ArgumentNullException.ThrowIfNull(spreadsheet);
        return spreadsheet.AddRowAsync(ReadOnlyMemory<DataCell>.Empty, token);
    }

    private static ValueTask AddRangeAsRowsAsync<T>(Spreadsheet spreadsheet, IEnumerable<T> objs, CancellationToken token)
    {
        ArgumentNullException.ThrowIfNull(spreadsheet);
        ArgumentNullException.ThrowIfNull(objs);
        return AddRangeAsEmptyRowsAsync(spreadsheet, objs, token);
    }

    private static async ValueTask AddRangeAsEmptyRowsAsync<T>(Spreadsheet spreadsheet, IEnumerable<T> objs, CancellationToken token)
    {
        foreach (var _ in objs)
        {
            await spreadsheet.AddRowAsync(ReadOnlyMemory<DataCell>.Empty, token).ConfigureAwait(false);
        }
    }
}
