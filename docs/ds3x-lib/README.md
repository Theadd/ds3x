# __ds3x__
<sup><sup>WIN32 / WIN64 / VBA7 <small><i>(MS Office 2010+)</i></small></sup></sup>

_A **lightweight MSAccess (VBA) shared library** providing a dead simple <u>abstraction</u> in working with different sources/types <u>of data collections</u> to query, iterate, filter, fix/reformat, transform and even to convert to and from other collection types (`CSV/Excel`, `ArrayLists`, `2D Arrays`, `Dictionaries`, `Recordsets`, `JSON`, etc.)_

### __QUICK FEATURE OVERVIEW__

* __`Minimum global scope pollution`__ - Except for a few public types and some externally accessible automation calls, everything else is contained within their class module scope.
* __`All possible data conversion among supported types`__ - *(e.g., Excel -> Recordset)*.
* __`High speed Excel data manipulation`__ - Using **dsTable**'s direct Excel range integration with automatic formatting preservation.

### __CLASS MODULES__

#### __COLLECTIONS__
<ul>
Ultra-lightweight chainable wrappers.

  - __[ArrayListEx](./ArrayListEx.md)__ - *`.NET Framework v3.5`'s `ArrayList` wrapper.*
  - __[Array2dEx](./Array2dEx.md)__ <sup><sub><sup><kbd><code>__IMMUTABLE__</code></kbd></sup></sub></sup> - *`VBA`'s built-in `2D Array` wrapper.*
  - __[dsTable](./dsTable.md)__ <sup><sub><sup><kbd><code>__IMMUTABLE__</code></kbd></sup></sub></sup> - *Structured data tables with header/record management and Excel integration*
  - __[DictionaryEx](./DictionaryEx.md)__ - *`Scripting.Dictionary` wrapper.*
  - __[RecordsetEx](./RecordsetEx.md)__ - *`ADODB.Recordset` wrapper.*
  - __[xlSheetsEx](./xlSheetsEx.md)__ - *`Excel` and `CSV-in-Excel` wrapper.*

</ul>

#### __LIBRARIES__
<ul>

  - __[JSON](./JSON.md)__ - *Backwards-compatible custom JSON utilities*
  - __[FileSystemLib](./FileSystemLib.md)__ - *File system operations and clipboard access*
</ul>

### __USAGE EXAMPLES__

```vb
' Create dsTable from Excel range with headers (Confidence: 9/10)
Dim dataTable As dsTable
Set dataTable = dsTable.Create(ActiveSheet.UsedRange, True)
```

```vb 
' Convert to CSV with pipe delimiter (Confidence: 9/10)
Dim csvOutput As String
csvOutput = dataTable.ToCSV("|")
Debug.Print csvOutput
