using Bogus;
using SpreadCheetah.Validations;

namespace SpreadCheetah.Test.Helpers;

internal static class DataValidationGenerator
{
    public static IList<DataValidation> Generate(int count)
    {
        var testValidations = new Faker<DataValidation>()
            .CustomInstantiator(f => f.PickRandom(Factories).Invoke(f))
            .RuleFor(x => x.ErrorMessage, f => f.Commerce.ProductName().OrNull(f, .1f))
            .RuleFor(x => x.ErrorTitle, f => f.Lorem.Word().OrNull(f, .1f))
            .RuleFor(x => x.ErrorType, f => f.Random.Enum<ValidationErrorType>())
            .RuleFor(x => x.IgnoreBlank, f => f.Random.Bool())
            .RuleFor(x => x.InputMessage, f => f.Commerce.ProductName().OrNull(f, .1f))
            .RuleFor(x => x.InputTitle, f => f.Lorem.Word().OrNull(f, .1f))
            .RuleFor(x => x.ShowErrorAlert, f => f.Random.Bool())
            .RuleFor(x => x.ShowInputMessage, f => f.Random.Bool());

        return testValidations.Generate(count);
    }

    private static readonly Func<Faker, DataValidation>[] Factories =
    [
        f => DataValidation.DateTimeBetween(f.Random.SimpleDateTimePair(out var max), max),
        f => DataValidation.DateTimeEqualTo(f.Random.SimpleDateTime(f.Date)),
        f => DataValidation.DateTimeGreaterThan(f.Random.SimpleDateTime(f.Date)),
        f => DataValidation.DateTimeGreaterThanOrEqualTo(f.Random.SimpleDateTime(f.Date)),
        f => DataValidation.DateTimeLessThan(f.Random.SimpleDateTime(f.Date)),
        f => DataValidation.DateTimeLessThanOrEqualTo(f.Random.SimpleDateTime(f.Date)),
        f => DataValidation.DateTimeNotBetween(f.Random.SimpleDateTimePair(out var max), max),
        f => DataValidation.DateTimeNotEqualTo(f.Random.SimpleDateTime(f.Date)),


        f => DataValidation.DecimalBetween(f.Random.SimpleDoublePair(out var max), max),
        f => DataValidation.DecimalEqualTo(f.Random.SimpleDouble()),
        f => DataValidation.DecimalGreaterThan(f.Random.SimpleDouble()),
        f => DataValidation.DecimalGreaterThanOrEqualTo(f.Random.SimpleDouble()),
        f => DataValidation.DecimalLessThan(f.Random.SimpleDouble()),
        f => DataValidation.DecimalLessThanOrEqualTo(f.Random.SimpleDouble()),
        f => DataValidation.DecimalNotBetween(f.Random.SimpleDoublePair(out var max), max),
        f => DataValidation.DecimalNotEqualTo(f.Random.SimpleDouble()),
        f => DataValidation.IntegerBetween(f.Random.IntPair(out var max), max),
        f => DataValidation.IntegerEqualTo(f.Random.Int()),
        f => DataValidation.IntegerGreaterThan(f.Random.Int()),
        f => DataValidation.IntegerGreaterThanOrEqualTo(f.Random.Int()),
        f => DataValidation.IntegerLessThan(f.Random.Int()),
        f => DataValidation.IntegerLessThanOrEqualTo(f.Random.Int()),
        f => DataValidation.IntegerNotBetween(f.Random.IntPair(out var max), max),
        f => DataValidation.IntegerNotEqualTo(f.Random.Int()),
        f => DataValidation.ListValues(f.Lorem.Words(), f.Random.Bool()),
        f => DataValidation.ListValuesFromCells(f.Random.CellRange(), f.Random.Bool()),
        f => DataValidation.ListValuesFromCells(f.Lorem.Word(), f.Random.CellRange(), f.Random.Bool()),
        f => DataValidation.TextLengthBetween(f.Random.IntPair(out var max), max),
        f => DataValidation.TextLengthEqualTo(f.Random.Int()),
        f => DataValidation.TextLengthGreaterThan(f.Random.Int()),
        f => DataValidation.TextLengthGreaterThanOrEqualTo(f.Random.Int()),
        f => DataValidation.TextLengthLessThan(f.Random.Int()),
        f => DataValidation.TextLengthLessThanOrEqualTo(f.Random.Int()),
        f => DataValidation.TextLengthNotBetween(f.Random.IntPair(out var max), max),
        f => DataValidation.TextLengthNotEqualTo(f.Random.Int()),
    ];
}
