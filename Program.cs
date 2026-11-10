using System.Data.Odbc;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using EntityFrameworkCore.Jet.Infrastructure;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Diagnostics;

// Shows that EntityFrameworkCore.Jet can't run a query with no table of its own, such as a top-level Any(), against
// an Access database it didn't create: the SQL reads from a one-row `#Dual` table, which only a database the provider
// creates has, so Access answers "cannot find the input table or query '#Dual'".
//
// Usage: EFCore.Jet.DualTableRepro [path to an existing, empty .mdb or .accdb]
// Without a path, it creates a new Access 2000 format .mdb in the temp folder.

if (!OperatingSystem.IsWindows())
{
    Console.Error.WriteLine("The Access ODBC driver only runs on Windows.");
    return 1;
}

return Repro.Run(args);

[SupportedOSPlatform("windows")]
internal static class Repro
{
    private const string Driver = "Microsoft Access Driver (*.mdb, *.accdb)";

    public static int Run(string[] args)
    {
        var path = args.Length > 0 ? Path.GetFullPath(args[0]) : NewDatabasePath();
        if (args.Length == 0)
        {
            AccessDatabase.Create(Driver, path);
        }

        var connectionString = $"Driver={{{Driver}}};DBQ={path};";
        AccessDatabase.CreateNotesTable(connectionString);
        var id = AddNote(connectionString);

        Console.WriteLine($"Process architecture: {RuntimeInformation.ProcessArchitecture}");
        Console.WriteLine($"Database: {path}");
        Console.WriteLine($"JetDualTable.Name: {JetDualTable.Name}");
        Console.WriteLine();

        var results = new List<Result>
        {
            Attempt("Notes.Any(x => x.Id == id)", "no #Dual table",
                () => Query(connectionString, x => x.Notes.Any(n => n.Id == id))),

            Attempt("Notes.Count(x => x.Id == id) > 0", "no #Dual table",
                () => Query(connectionString, x => x.Notes.Count(n => n.Id == id) > 0)),

            // A query of another shape, because EF caches the SQL of the first one, with `#Dual` in it
            Attempt("Notes.Any(x => x.Body == body)", "JetDualTable.CustomName = \"MSysRelationships\"", () =>
            {
                JetDualTable.CustomName = "MSysRelationships";
                try
                {
                    return Query(connectionString, x => x.Notes.Any(n => n.Body == "a note"));
                }
                finally
                {
                    JetDualTable.CustomName = "";
                }
            }),

            Attempt("Notes.Any(x => x.Id == id)", "#Dual table created as the provider creates it", () =>
            {
                AccessDatabase.CreateDualTable(connectionString);
                return Query(connectionString, x => x.Notes.Any(n => n.Id == id));
            })
        };

        foreach (var result in results)
        {
            Console.WriteLine($"Query:   {result.Query}");
            Console.WriteLine($"Setup:   {result.Setup}");
            Console.WriteLine($"Outcome: {result.Outcome}");
            Console.WriteLine();
        }

        return 0;
    }

    private static string NewDatabasePath() =>
        Path.Combine(Path.GetTempPath(), $"DualTableRepro-{DateTime.Now:yyyyMMdd-HHmmss}.mdb");

    private static int AddNote(string connectionString)
    {
        using var context = NotesContext.Create(connectionString);
        var note = new Note { Body = "a note" };
        context.Notes.Add(note);
        context.SaveChanges();

        return note.Id;
    }

    private static string Query(string connectionString, Func<NotesContext, bool> query)
    {
        using var context = NotesContext.Create(connectionString);
        return $"returned {query(context)}";
    }

    private static Result Attempt(string query, string setup, Func<string> attempt)
    {
        try
        {
            return new(query, setup, attempt());
        }
        catch (OdbcException e)
        {
            return new(query, setup, $"FAILED: {e.Message}");
        }
    }

    private sealed record Result(string Query, string Setup, string Outcome);
}

public sealed class Note
{
    public int Id { get; set; }
    public string? Body { get; set; }
}

[SupportedOSPlatform("windows")]
public sealed class NotesContext(DbContextOptions<NotesContext> options) : DbContext(options)
{
    public DbSet<Note> Notes => Set<Note>();

    public static NotesContext Create(string connectionString) =>
        new(new DbContextOptionsBuilder<NotesContext>()
            .UseJetOdbc(connectionString)
            .LogTo(Console.WriteLine, [RelationalEventId.CommandError])
            .Options);

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        var note = modelBuilder.Entity<Note>();
        note.ToTable("Notes");
        note.Property(x => x.Body).HasMaxLength(255);
    }
}

/// <summary>
/// Creates the database with the ODBC driver itself, as Access or another program would, rather than EnsureCreated,
/// which would also create the `#Dual` table
/// </summary>
[SupportedOSPlatform("windows")]
internal static partial class AccessDatabase
{
    private const ushort OdbcAddDsn = 1;

    public static void Create(string driver, string path)
    {
        // CREATE_DBV4 creates the Access 2000 (Jet 4) format
        if (!SQLConfigDataSourceW(IntPtr.Zero, OdbcAddDsn, driver, $"CREATE_DBV4=\"{path}\" General\0"))
        {
            throw new InvalidOperationException(
                $"Couldn't create {path} with the \"{driver}\" ODBC driver: {InstallerError()}");
        }
    }

    public static void CreateNotesTable(string connectionString) =>
        Execute(connectionString,
            "CREATE TABLE Notes (Id COUNTER CONSTRAINT PK_Notes PRIMARY KEY, Body VARCHAR(255))");

    /// <summary>
    /// The table and row that the provider's JetConnection.CreateDatabase adds, with EnsureDualTable
    /// </summary>
    public static void CreateDualTable(string connectionString)
    {
        Execute(connectionString,
            "CREATE TABLE `#Dual` (`ID` INTEGER NOT NULL CONSTRAINT `PrimaryKey` PRIMARY KEY)");
        Execute(connectionString, "INSERT INTO `#Dual` (`ID`) VALUES (1)");
    }

    private static void Execute(string connectionString, string sql)
    {
        using var connection = new OdbcConnection(connectionString);
        connection.Open();
        using var command = connection.CreateCommand();
        command.CommandText = sql;
        command.ExecuteNonQuery();
    }

    private static unsafe string InstallerError()
    {
        const ushort capacity = 512;
        var message = stackalloc char[capacity];
        return SQLInstallerErrorW(1, out var code, message, capacity, out var length) is 0 or 1
            ? $"{new string(message, 0, Math.Min((int)length, capacity))} (error {code})"
            : "no error was given";
    }

    [LibraryImport("odbccp32.dll", StringMarshalling = StringMarshalling.Utf16)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool SQLConfigDataSourceW(IntPtr hwndParent, ushort fRequest, string lpszDriver,
        string lpszAttributes);

    [LibraryImport("odbccp32.dll")]
    private static unsafe partial short SQLInstallerErrorW(ushort iError, out uint pfErrorCode, char* lpszErrorMsg,
        ushort cbErrorMsgMax, out ushort pcbErrorMsg);
}
