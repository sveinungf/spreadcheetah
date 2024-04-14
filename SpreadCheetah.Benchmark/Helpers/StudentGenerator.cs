namespace SpreadCheetah.Benchmark.Helpers;

internal static class StudentGenerator
{
    private static readonly string[] Firstnames =
    [
        "Ada", "Aksel", "Ella", "Emil", "Emma", "Filip", "Henrik", "Ingrid", "Jakob",
        "Lucas", "Maja", "Noah", "Nora", "Olivia", "Oskar", "Sara", "Sofie", "William"
    ];

    private static readonly string[] Lastnames =
    [
        "Berg", "Dahl", "Eide", "Haugen", "Hansen", "Holm", "Johannessen", "Nguyen", "Nordmann", "Solberg", "Strand"
    ];

    public static List<Student> Generate(int count)
    {
        var result = new List<Student>(count);
        var r = new Random();

        for (var i = 0; i <= count; ++i)
        {
            result.Add(new Student
            {
                BirthYear = r.Next(1970, 2000),
                Enrolled = r.NextBoolean(),
                Firstname = Firstnames[r.Next(0, Firstnames.Length)],
                Gpa = r.NextDecimal(1, 4),
                LastName = Lastnames[r.Next(0, Lastnames.Length)]
            });
        }

        return result;
    }
}
