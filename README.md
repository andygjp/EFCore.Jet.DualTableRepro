# EFCore.Jet dual table repro

Shows that [EntityFrameworkCore.Jet](https://github.com/CirrusRedOrg/EntityFrameworkCore.Jet) can't run a query that
has no table of its own, such as a top-level `Any()`, against an Access database the provider didn't create.

Jet requires a `FROM`, so the provider adds one that reads from a one-row table, `JetDualTable.Name`:

```sql
SELECT EXISTS (
    SELECT 1
    FROM `Notes` AS `n`
    WHERE `n`.`Id` = @id)
FROM (SELECT COUNT(*) FROM `#Dual`)
```

`JetDualTable.Name` is `#Dual` unless `JetDualTable.CustomName` is set. Only a database the provider creates
(`JetConnection.CreateDatabase`, which calls `EnsureDualTable`) has that table. Only the scaffolding factory looks for
an alternative (`MSysAccessStorage`, `MSysRelationships`), so an application using a database made by Access gets:

```
ERROR [42S02] [Microsoft][ODBC Microsoft Access Driver] The Microsoft Jet database engine cannot find the input table or query '#Dual'.  Make sure it exists and that its name is spelled correctly.
```

## What it does

1. Creates a new, empty Access 2000 format `.mdb` in the temp folder, with the ODBC driver's own `CREATE_DBV4`
   (`SQLConfigDataSourceW`), as a database made by Access would be. `EnsureCreated` would add the `#Dual` table.
2. Creates a table `Notes`, and adds a note with EF.
3. Runs these queries, each in a new context, and prints what each returned or the error:
   1. `Notes.Any(x => x.Id == id)`, which fails.
   2. `Notes.Count(x => x.Id == id) > 0`, which reads `FROM Notes` and so runs. This is the workaround we use.
   3. `Notes.Any(...)` with `JetDualTable.CustomName = "MSysRelationships"`, the alternative the provider's comments
      suggest. Whether it runs depends on whether the user may read that system table.
   4. The first query again, after creating `#Dual` with one row, as `EnsureDualTable` does.

EF logs the SQL of each command that fails, so the output also shows the `FROM (SELECT COUNT(*) FROM #Dual)`.

## Requirements

- Windows, with the Microsoft Access Database Engine ODBC driver, "Microsoft Access Driver (*.mdb, *.accdb)". Office
  installs it, as does the Access Database Engine redistributable.
- The .NET 11 SDK.

The process must have the same architecture as the driver, which has the architecture of the Office that installed
it. The project builds for `win-x64`, self-contained, for 64-bit Office. On Windows on ARM, it runs under x64
emulation.

## Run

```shell
dotnet run
```

For 32-bit Office:

```shell
dotnet run -r win-x86
```

To use an existing empty database instead of creating one, pass its path:

```shell
dotnet run -- C:\path\to\Empty.mdb
```

If you get `ERROR [IM003] Specified driver could not be loaded due to system error 193`, the process and the driver
have different architectures: use the other `-r`.

## Expected output

After the logged SQL of the failing commands:

```
Query:   Notes.Any(x => x.Id == id)
Setup:   no #Dual table
Outcome: FAILED: ERROR [42S02] [Microsoft][ODBC Microsoft Access Driver] The Microsoft Jet database engine cannot find the input table or query '#Dual'.  Make sure it exists and that its name is spelled correctly.

Query:   Notes.Count(x => x.Id == id) > 0
Setup:   no #Dual table
Outcome: returned True

Query:   Notes.Any(x => x.Body == body)
Setup:   JetDualTable.CustomName = "MSysRelationships"
Outcome: returned True, or FAILED if the user can't read MSysRelationships

Query:   Notes.Any(x => x.Id == id)
Setup:   #Dual table created as the provider creates it
Outcome: returned True
```

## Versions

- EntityFrameworkCore.Jet.Odbc 11.0.0-alpha.5, with the EF Core and System.Data.Odbc versions it depends on
- .NET 11
