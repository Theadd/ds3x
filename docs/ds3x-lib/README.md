# __ds3x__
<sup><sup>WIN32 / WIN64 / VBA7 <small><i>(MS Office 2010+)</i></small></sup></sup>

_A **lightweight MSAccess (VBA) shared library** providing a dead simple <u>abstraction</u> in working with different sources/types <u>of data collections</u> to query, iterate, filter, fix/reformat, transform and even to convert to and from other collection types (`CSV/Excel`, `ArrayLists`, `2D Arrays`, `Dictionaries`, `Recordsets`, `JSON`, etc.)_

### __QUICK FEATURE OVERVIEW__

* __`Minimum global scope pollution`__ - Except for a few public types and some externally accessible automation calls, everything else is contained whithin their class module scope.
* __`All possible data conversion among supported types`__ - *(e.g., Excel -> Recordset)*.
* __`High speed Excel data manipulation`__ - Making use of **Array2dEx**'s `2D Array` implementation for direct in-memory copying to and from an `Excel.Range`, transforming it and writting it back, several times faster than iterating over cell values.
* . . .

### __CLASS MODULES__

#### __COLLECTIONS__

<ul>
Ultra-lightweight chainable wrappers.

  - __[ArrayListEx](./ArrayListEx.md)__ - *`.NET Framework v3.5`'s `ArrayList` wrapper.*
  - __[Array2dEx](./Array2dEx.md)__ <sup><sub><sup><kbd><code>__IMMUTABLE__</code></kbd></sup></sub></sup> - *`VBA`'s built-in `2D Array` wrapper.*
  - __[DictionaryEx](./DictionaryEx.md)__ - *`Scripting.Dictionary` wrapper.*
  - __[RecordsetEx](./RecordsetEx.md)__ - *`ADODB.Recordset` wrapper.*
  - __[xlSheetsEx](./xlSheetsEx.md)__ - *`Excel` and `CSV-in-Excel` wrapper.*
  - __[dsTable](./dsTable.md)__ <sup><sub><sup><kbd><code>__IMMUTABLE__</code></kbd></sup></sub></sup> - *Greatly simplifies working w/ table-like collections (holding column-related info, not just data).*

</ul>
<ul>
Miscellaneous.

  - __[Callback](./Callback.md)__ - *A <kbd>`pass function as parameter`</kbd>-like feature on steroids <s>with additional support for executing `javascript` code and `Filter`/`Map`/`Reduce` calls on callback's results (`Where`/`Select`/`Aggregate` equivalents on `.NET`).*</s> <i><small>(Currently for 32-bit only unless I figure out why and how to make it work in 64-bit).</small></i>
  - __[dbQuery](./dbQuery.md)__ <sup><sub><sup><kbd><code>__OPTIONAL__</code></kbd></sup></sub></sup> - *Not a fan of your current `ADODB` connector? Just try this one.*

</ul>

#### __LIBRARIES__

<ul>

  - __[JSON](./JSON.md)__ - *Backwards-compatible custom version of [Tim Hall](https://github.com/VBA-tools/VBA-JSON)'s `JSON` utilities with opinionated pretty printing.*
  - __[FileSystemLib](./FileSystemLib.md)__ - *Safe, network<sup><small>(NFS)</small></sup>-load/delay aware, `FileSystemObject`-related utilities, read/write to system clipboard and virtual memory usage stats.*
  - __[ScreenLib](./ScreenLib.md)__ <sup><sub><sup><kbd><code>__OPTIONAL__</code></kbd></sup></sub></sup> - *No more messing with twips to pixels conversion. Everything you'd ever need, speaking in the same language.*
  - __[MemoryLib](https://github.com/cristianbuse/VBA-MemoryTools)__ - *[Cristian Buse](https://github.com/cristianbuse)'s `VBA-MemoryTools` as a class module.*

</ul>


