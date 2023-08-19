# __ds3x__

_A **lightweight MSAccess (VBA) shared library** providing a dead simple <u>abstraction</u> in working with different sources/types <u>of data collections</u> to query, iterate, filter, fix/reformat, transform and even to convert to and from other collection types (`CSV/Excel`, `ArrayLists`, `2D Arrays`, `Dictionaries`, `Recordsets`, `JSON`, etc.)_


### __QUICK FEATURE OVERVIEW__

<ul>
<strong><u>RAW</u> INPUT/OUTPUT DATA SUPPORT</strong>

`CSV`, `Excel`, `ADODB.Recordset`, `2D Array`, `Jagged Array`, `JSON`.

</ul>
<ul>
<strong>SUPPORTED OPERATIONS</strong>

  - __All possible data conversion among supported types__ *(e.g., `Excel` -> `Recordset`)*.
  - __dsLiveEd__ - _(Setup and manage lists of transformation or other kind of tasks to be applied in a `PowerQuery`-like way)._
    - __Modes of Use__
      - __Live Editor UI__ - *It can be embedded in your application as a project reference or opened as an external (standalone) application.*
      - __Headless mode__ - *Those list of tasks to be applied can be exported as a presset in JSON, which can be imported in your application and just generate the resulting output without having to involve any visible UI.*
      - __Automation__ - *Formerly known as `OLE Automation`, allows to programatically use `dsLiveEd` from another application without even having to include `ds3x` as a project reference.*
    - __Immutability support__ - *Allowing to go back and forward within the resulting state of each and every single transformation task applied.*

</ul>


#### __CLASS MODULES__

<ul>
Ultra-lightweight chainable wrappers.

  - __[ArrayListEx](./docs/ArrayListEx.md)__ - *`.NET Framework v3.5`'s `ArrayList` wrapper.*
  - __[Array2dEx](./docs/Array2dEx.md)__ - *`VBA`'s built-in `2D Array` wrapper.*
  - __[DictionaryEx](./docs/DictionaryEx.md)__ - *`Scripting.Dictionary` wrapper.*
  - __[RecordsetEx](./docs/RecordsetEx.md)__ - *`ADODB.Recordset` wrapper.*
  - <sup><sub><sup><kbd><code>__OPTIONAL__</code></kbd></sup></sub></sup> __[xlSheetsEx](./docs/xlSheetsEx.md)__ - *`CSV`/`Excel` wrapper.*
  - <sup><sub><sup><kbd><code>__OPTIONAL__</code></kbd></sup></sub></sup> __[dsTable](./docs/dsTable.md)__ - *Greatly simplifies working w/ table-like collections (holding column-related info, not just data).*
  - <sup><sub><sup><kbd><code>__OPTIONAL__</code></kbd></sup></sub></sup> __[Callback](./docs/Callback.md)__ - *A <kbd>`pass function as parameter`</kbd>-like feature on steroids.*

</ul>
<ul>
Miscellaneous.

  - __[JSON](./docs/JSON.md)__ - *Backwards-compatible custom version of [Tim Hall](https://github.com/VBA-tools/VBA-JSON)'s `JSON` utilities.*
  - <sup><sub><sup><kbd><code>__OPTIONAL__</code></kbd></sup></sub></sup> __[dbQuery](./docs/dbQuery.md)__ - *Not a fan of your current `ADODB` connector? Just try this one.*

</ul>

#### __STANDARD MODULES__

<ul>

  - <sup><sub><sup><kbd><code>__OPTIONAL__</code></kbd></sup></sub></sup> __[FileSystemLib](./docs/FileSystemLib.md)__ - *Safe, network<sup><small>(NFS)</small></sup>-load/delay aware, `FileSystemObject`-related, clipboard and memory usage utilities.*
  - <sup><sub><sup><kbd><code>__OPTIONAL__</code></kbd></sup></sub></sup> __[ScreenLib](./docs/ScreenLib.md)__ - *No more messing with twips to pixels conversion. Everything you'd ever need, speaking in the same language.*
  
</ul>

