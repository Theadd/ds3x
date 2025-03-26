## **`FileSystemLib` Class** <sup><sub><sup>

_Advanced file system operations with network-safe retry patterns and multi-encoding support._

---

- Implements **TryWait pattern** for robust network file system operations
- Automatic charset detection for 6 encoding formats
- Path normalization with project-relative resolution
- `UTF-8/16` text streaming with `BOM` handling
- All methods returning `FileSystemLib` are chainable

---

### **Usage Examples**

* Read text file with automatic encoding detection:

```vb
Dim content As String
If FileSystemLib.TryReadAllTextInFile(FileSystemLib.Resolve("../data.csv"), content) Then
    Debug.Print "File content: " & content
End If
' Direct unsafe version:
content = FileSystemLib.ReadAllTextInFile(FileSystemLib.Resolve("../data.csv"))
```

---

### **API Overview**

```vb
' Properties
Property Get PathSeparator() As String
Property Get Resolve(Path As String, [RelativeTo]) As String
Property Get StreamReader(TargetFile As String) As ADODB.Stream
Property Get StreamWriter() As ADODB.Stream

' TryWait Pattern Functions
Function TryWaitFileWriteAccess(TargetFile, [retryAttempts=100]) As Boolean
Function TryWaitFileExists(TargetFile, [retryAttempts=100]) As Boolean
Function TryWaitFolderExists(TargetPath, [retryAttempts=100]) As Boolean

' Encoding-Safe IO
Function TryReadAllTextInFile(TargetFile, Content, [Charset]) As Boolean
Function TryWriteTextToFile(TargetFile, Content, [overwrite], [AsUnicode]) As Boolean

' Path Handling
Function PathCombine(directoryPath, filename) As String
Function TryCreateFolder(TargetPath) As Boolean
Function TryGetFileInAncestors(TargetFile, [BackwardMovesLimit]) As Boolean

' Clipboard
Function SystemClipboard([StoreText]) As String
```

<table width="100%"><caption>

### **`FileSystemLib` API**  
</caption>
<thead><tr><th colspan="2">PATH HANDLING</th></tr></thead>
<tbody>

<tr><td align="left" valign="top">

```vb
Resolve(Path, [RelativeTo]) As String
```
</td><td align="left" valign="top">
Normalizes paths and converts relative to absolute using CodeProject.Path as base. Handles network paths and mixed slashes.
<details><summary>Example</summary>
<code>Resolve("data/files", "G:\projects")</code> → <code>"G:\projects\data\files"</code>
</details>
</td></tr>

<tr><td align="left" valign="top">

```vb
PathCombine(directoryPath, filename) As String
```
</td><td align="left" valign="top">
Safe path concatenation with proper separator handling. Network path aware.
</td></tr>

</tbody>

<thead><tr><th colspan="2">TRYWAIT PATTERN</th></tr></thead>
<tbody>

<tr><td align="left" valign="top">

```vb
TryWaitFileWriteAccess(TargetFile, [retryAttempts=100]) As Boolean
```
</td><td align="left" valign="top">
Ensures file write access through atomic rename checks. Essential for network files.
<details><summary>Network Safety</summary>
Uses temporary file renaming to verify write locks. Retries every 100ms up to specified attempts.
</details>
</td></tr>

<tr><td align="left" valign="top">

```vb
TryWaitFileExists(TargetFile, [retryAttempts=100]) As Boolean
```
</td><td align="left" valign="top">
Waits for file existence with retry logic. Handles network latency.
</td></tr>

<tr><td align="left" valign="top">

```vb
TryWaitFolderExists(TargetPath, [retryAttempts=100]) As Boolean
```
</td><td align="left" valign="top">
Verifies folder existence with network-safe retries.
</td></tr>

</tbody>

<thead><tr><th colspan="2">ENCODING HANDLING</th></tr></thead>
<tbody>

<tr><td align="left" valign="top">

```vb
TryReadAllTextInFile(TargetFile, Content, [Charset]) As Boolean
```
</td><td align="left" valign="top">

Reads text files with automatic detection of:
`ANSI`, `UTF-8`, `UTF-16(LE/BE)`, `UTF-8-BOM`, `UTF-8/ANSI`
Uses optimized streaming for large files.
</td></tr>

<tr><td align="left" valign="top">

```vb
TryWriteTextToFile(TargetFile, Content, [overwrite], [AsUnicode]) As Boolean
```
</td><td align="left" valign="top">
Writes text with specified encoding (default Unicode). Implements atomic write-through caching.
</td></tr>

</tbody>

<thead><tr><th colspan="2">STREAM HANDLING</th></tr></thead>
<tbody>

<tr><td align="left" valign="top">

```vb
Property Get StreamReader(TargetFile) As ADODB.Stream
```
</td><td align="left" valign="top">
Creates UTF-aware text reader with automatic BOM detection. 
<details><summary>Features</summary>
- Handles files up to 4GB<br/>
- Buffered reading for performance<br/>
- Automatic charset detection
</details>
</td></tr>

<tr><td align="left" valign="top">

```vb
Property Get StreamWriter() As ADODB.Stream
```
</td><td align="left" valign="top">
Creates configurable text writer with encoding options.
</td></tr>

</tbody>

<thead><tr><th colspan="2">CLIPBOARD</th></tr></thead>
<tbody>

<tr><td align="left" valign="top">

```vb
SystemClipboard([StoreText]) As String
```
</td><td align="left" valign="top">
Reads/writes system clipboard with UTF-8 support. Handles large text (up to 1MB).
</td></tr>

</tbody>

<thead><tr><th colspan="2">FILE OPERATIONS</th></tr></thead>
<tbody>

<tr><td align="left" valign="top">

```vb
TryCreateFolder(TargetPath) As Boolean
```
</td><td align="left" valign="top">
Creates directories with parent hierarchy. Returns False if path is file.
</td></tr>

<tr><td align="left" valign="top">

```vb
TryGetFileInAncestors(TargetFile, [BackwardMovesLimit]) As Boolean
```
</td><td align="left" valign="top">
Searches parent directories for file. Useful for config file discovery.
</td></tr>

<tr><td align="left" valign="top">

```vb
TryShellExecute(TargetFile, [ShowStyle]) As Boolean
```
</td><td align="left" valign="top">
Opens files with default handler. Network path safe.
</td></tr>

</tbody>
</table>
