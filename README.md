# EFCore.Jet dual table repro

Showed that [EntityFrameworkCore.Jet](https://github.com/CirrusRedOrg/EntityFrameworkCore.Jet) couldn't run a query
that has no table of its own, such as a top-level `Any()`, against an Access database the provider didn't create:
[issue #307](https://github.com/CirrusRedOrg/EntityFrameworkCore.Jet/issues/307). 11.0.0-beta.1 fixes it
([#308](https://github.com/CirrusRedOrg/EntityFrameworkCore.Jet/pull/308)), and the project now uses that version.

## The issue, up to 11.0.0-alpha.5

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
ERROR [42S02] [Microsoft][ODBC Microsoft Access Driver] The Microsoft Access database engine cannot find the input table or query '#Dual'. Make sure it exists and that its name is spelled correctly.
```

## The fix, in 11.0.0-beta.1

The stand-in table is now `JetDualTable.CustomName` when set, then `#Dual` when the open database has it, then the
first table the query itself reads. So on a database without `#Dual`, the `Any()` above reads `FROM Notes`:

```sql
SELECT EXISTS (
    SELECT 1
    FROM `Notes` AS `n`
    WHERE `n`.`Id` = @id)
FROM (SELECT COUNT(*) FROM `Notes`)
```

`JetDualTable.Name` is now blank unless `CustomName` is set or the scaffolding factory has detected a table.
`CustomName = "MSysRelationships"` still fails, as `CustomName` comes first and ACE won't let the default user read
that table.

## What it does

1. Creates a new, empty Access 2000 format `.mdb` in the temp folder, with the ODBC driver's own `CREATE_DBV4`
   (`SQLConfigDataSourceW`), as a database made by Access would be. `EnsureCreated` would add the `#Dual` table.
2. Creates a table `Notes`, and adds a note with EF.
3. Runs these queries, each in a new context, and prints what each returned or the error:
   1. `Notes.Any(x => x.Id == id)`, which failed up to alpha.5 and runs on beta.1.
   2. `Notes.Count(x => x.Id == id) > 0`, which reads `FROM Notes` and so runs. This was the workaround we used.
   3. `Notes.Any(...)` with `JetDualTable.CustomName = "MSysRelationships"`, the alternative the provider's comments
      suggested, which still fails: the default user may not read that system table.
   4. The first query again, after creating `#Dual` with one row, as `EnsureDualTable` does.

EF logs the SQL of each command that fails, so the output also shows the `FROM (SELECT COUNT(*) FROM ...)` of each
failure.

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

## Output

From a run on Windows 11 on ARM, as x64, with the driver from Microsoft 365 (Office 16, x64), on 11.0.0-beta.1:

```
Process architecture: X64
Database: C:\Users\andrewpattison\AppData\Local\Temp\DualTableRepro-20261116-072956.mdb
JetDualTable.Name: 

fail: 16/11/2026 07:30:08.770 RelationalEventId.CommandError[20102] (Microsoft.EntityFrameworkCore.Database.Command) 
      Failed executing DbCommand (6ms) [Parameters=[], CommandType='Text', CommandTimeout='30']
      SELECT EXISTS (
          SELECT 1
          FROM `Notes` AS `n`
          WHERE `n`.`Body` = 'a note')
      FROM (SELECT COUNT(*) FROM `MSysRelationships`)
Query:   Notes.Any(x => x.Id == id)
Setup:   no #Dual table
Outcome: returned True

Query:   Notes.Count(x => x.Id == id) > 0
Setup:   no #Dual table
Outcome: returned True

Query:   Notes.Any(x => x.Body == body)
Setup:   JetDualTable.CustomName = "MSysRelationships"
Outcome: FAILED: ERROR [42000] [Microsoft][ODBC Microsoft Access Driver] Record(s) cannot be read; no read permission on 'MSysRelationships'.

Query:   Notes.Any(x => x.Id == id)
Setup:   #Dual table created as the provider creates it
Outcome: returned True
```

The same run on 11.0.0-alpha.5:

```
Process architecture: X64
Database: C:\Users\andrewpattison\AppData\Local\Temp\DualTableRepro-20261110-170910.mdb
JetDualTable.Name: #Dual

fail: 10/11/2026 17:09:17.947 RelationalEventId.CommandError[20102] (Microsoft.EntityFrameworkCore.Database.Command) 
      Failed executing DbCommand (29ms) [Parameters=[@id='?' (DbType = Int32)], CommandType='Text', CommandTimeout='30']
      SELECT EXISTS (
          SELECT 1
          FROM `Notes` AS `n`
          WHERE `n`.`Id` = @id)
      FROM (SELECT COUNT(*) FROM `#Dual`)
fail: 10/11/2026 17:09:18.646 RelationalEventId.CommandError[20102] (Microsoft.EntityFrameworkCore.Database.Command) 
      Failed executing DbCommand (7ms) [Parameters=[], CommandType='Text', CommandTimeout='30']
      SELECT EXISTS (
          SELECT 1
          FROM `Notes` AS `n`
          WHERE `n`.`Body` = 'a note')
      FROM (SELECT COUNT(*) FROM `MSysRelationships`)
Query:   Notes.Any(x => x.Id == id)
Setup:   no #Dual table
Outcome: FAILED: ERROR [42S02] [Microsoft][ODBC Microsoft Access Driver] The Microsoft Access database engine cannot find the input table or query '#Dual'. Make sure it exists and that its name is spelled correctly.

Query:   Notes.Count(x => x.Id == id) > 0
Setup:   no #Dual table
Outcome: returned True

Query:   Notes.Any(x => x.Body == body)
Setup:   JetDualTable.CustomName = "MSysRelationships"
Outcome: FAILED: ERROR [42000] [Microsoft][ODBC Microsoft Access Driver] Record(s) cannot be read; no read permission on 'MSysRelationships'.

Query:   Notes.Any(x => x.Id == id)
Setup:   #Dual table created as the provider creates it
Outcome: returned True
```

## Versions

- EntityFrameworkCore.Jet.Odbc 11.0.0-beta.1, with the EF Core and System.Data.Odbc versions it depends on
- .NET 11
