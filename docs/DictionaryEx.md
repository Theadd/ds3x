## **`DictionaryEx` Class** <sup><sub><sup> &nbsp; <kbd><code>__CHAINABLE__</code></kbd></sup></sub></sup>

_A **lightweight** wrapper around `Scripting.Dictionary` objects._

Implements __[ICollectionEx](./ICollectionEx.md)__.

---

- `DictionaryEx` instances can hold keys and values of any type, like objects or arrays, not only value-types.
- All methods returning an `DictionaryEx` are chainable.
- Publicly exposes the inner `Scripting.Dictionary` instance for extensibility purposes with no code alterning needed.
- Can directly access elements in nested levels, supporting any combination of dictionary-like and array-like elements.
- It's enumerable (for each support) and with `.Item` as the default class member.
- `DictionaryEx.Create()` supports creating new instances by directly converting from:
  - **Scripting.Dictionary** and **DictionaryEx**.
  - Plain **2D Arrays** and **Jagged Arrays** (array of arrays).
  - **ArrayList** and **ArrayListEx**.
  - **Array2dEx**.
  - **JSON** string.

---

### **Usage Examples**

* Write a value of a nested path in a JSON file.

```vb
FileSystemLib.TryWriteTextToFile _
    "../package.json", _
    DictionaryEx.Create(FileSystemLib.ReadAllTextInFile("../package.json", False)) _
        .SetValue("ds3x.dev-tests[0].values", Array(1, 3, 5, 7)) _
        .ToJSON(), _
    asUnicode:=False
```

---

### **API Overview**

```vb
' Fields
Public Instance As Scripting.Dictionary
' Properties
Property Get Count() As Long
Property Get ColumnCount() As Long
Property Get Item(Key As Variant) As Variant
Property Get Key(Key As Variant, NewKey as Variant)
Property Get GetValue(Key As Variant, Optional DefaultValue As Variant) As Variant
Property Get Row(Index As Long) As Variant
' Functions
Function SetValue(Key As Variant, Value As Variant) As DictionaryEx
Function Create(Optional DictionaryLike As Variant) As DictionaryEx
Function Bind(DictionaryLike As Variant) As DictionaryEx
Function Unbind() As DictionaryEx
Function Add(Key As Variant, Value As Variant) As DictionaryEx
Function AddRange(Target As Varriant) As DictionaryEx
Function GetRange(Optional Index As Long, Optional GetCount As Long) As DictionaryEx
Function GetRows(RowIndexes As Variant, Optional ModIndex As Long) As DictionaryEx
Function CopyToRange(RangeObject As Excel.Range, Optional ApplyUserLocale As Boolean, Optional WriteHeaders As Boolean) As DictionaryEx
Function Remove(Key As Variant) As DictionaryEx
Function Clear() As DictionaryEx
Function Exists(Key as Variant) As Boolean
Function Items() As Variant()
Function Keys() As Variant()
Function Entries() As Variant()
Function Clone() As DictionaryEx
Function Duplicate() As DictionaryEx
Function ToCSV(Optional Delimiter As String) As String
Function ToJSON() As String
Function ToExcel() As String
```


<table width="100%"><caption>

### **`DictionaryEx` API**  
</caption>
<thead><tr><th colspan="2">FIELDS</th></tr></thead>
<tbody>


<tr><td align="left" valign="top">

```vb
Instance As Scripting.Dictionary
```
</td><td align="left" valign="top">
The <code>Scripting.Dictionary</code> object wrapped by this <code>DictionaryEx</code> instance.
</td></tr>

</tbody>


<thead><tr><th colspan="2">PROPERTIES</th></tr></thead>
<tbody>


<tr><td align="left" valign="top">

```vb
Get Count() As Long
```
</td><td align="left" valign="top">
Gets the number of elements in the <code>Scripting.Dictionary</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
Get ColumnCount() As Long
```
</td><td align="left" valign="top">
This is always 2, (keys and values).
</td></tr>


<tr><td align="left" valign="top">

```vb
Get Item(Key As Variant) As Variant
Let Item(Key As Variant, Value As Variant)
Set Item(Key As Variant, Value As Variant)
```
</td><td align="left" valign="top">
Sets or gets the <code>Value</code> at the specified <code>Key</code>. It's the <u>default class member</u>.
<br/>When the <code>Key</code> is a <code>String</code>, supports directly access to other dictionary-like or array-like within nested levels.
</td></tr>


<tr><td align="left" valign="top">

```vb
Let Key(Key As Variant, NewKey as Variant)
```
</td><td align="left" valign="top">
Updates a key, replacing the old key with the new key.
</td></tr>


<tr><td align="left" valign="top">

```vb
Get GetValue(Key As Variant, Optional DefaultValue As Variant) As Variant
```
</td><td align="left" valign="top">
Gets the <code>Value</code> at the specified <code>Key</code> if exists, or <code>DefaultValue</code> otherwise.
</td></tr>


<tr><td align="left" valign="top">

```vb
Get Row(Index As Long) As Variant
```
</td><td align="left" valign="top">
Gets the key-value <code>Array</code> at the specified row <code>Index</code>. 
</td></tr>


</tbody>



<thead><tr><th colspan="2">FUNCTIONS</th></tr></thead>
<tbody>


<tr><td align="left" valign="top">

```vb
SetValue(Key As Variant, Value As Variant) As DictionaryEx
```
</td><td align="left" valign="top">
Chainable wrapper around <code>Item</code> setter.
</td></tr>


<tr><td align="left" valign="top">

```vb
Create(Optional DictionaryLike As Variant) As DictionaryEx
```
</td><td align="left" valign="top">
When no parameter is provided, returns a new <code>DictionaryEx</code> instance with an empty <code>Scripting.Dictionary</code>.
<br/>
Otherwise, returns a new instance containing the values obtained by converting the provided <code>DictionaryLike</code> to a <code>Scripting.Dictionary</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
Bind(DictionaryLike As Variant) As DictionaryEx
```
</td><td align="left" valign="top">
Instead of returning a new instance that references the provided <code>DictionaryLike</code>, points this <code>DictionaryEx</code> instance it.
</td></tr>


<tr><td align="left" valign="top">

```vb
Unbind() As DictionaryEx
```
</td><td align="left" valign="top">
Removes the reference to the <code>Scripting.Dictionary</code> instance.
</td></tr>


<tr><td align="left" valign="top">

```vb
Add(Key As Variant, Value As Variant) As DictionaryEx
```
</td><td align="left" valign="top">
Adds or replaces a key-value pair to the <code>Scripting.Dictionary</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
AddRange(Target As Varriant) As DictionaryEx
```
</td><td align="left" valign="top">
Adds or replaces all key-value pairs from <code>Target</code> dictionary-like collection.
</td></tr>


<tr><td align="left" valign="top">

```vb
GetRange(Optional Index As Long, Optional GetCount As Long) As DictionaryEx
```
</td><td align="left" valign="top">
Returns a new <code>DictionaryEx</code> with the specified range of elements.
</td></tr>


<tr><td align="left" valign="top">

```vb
GetRows(RowIndexes As Variant, Optional ModIndex As Long) As DictionaryEx
```
</td><td align="left" valign="top">
Returns a new <code>DictionaryEx</code> instance containing only those rows specified in the <code>RowIndexes</code> array.
<details><summary><code>PARAMETERS</code></summary><ul>
<li><kbd>RowIndexes</kbd> → An <code>Array</code> of row indices.</li>
<li><kbd>ModIndex</kbd> → A signed integer to shift values in the <code>RowIndexes</code> array.
<br/>Default value: <code>0</code></li>
</ul></details>
</td></tr>


<tr><td align="left" valign="top">

```vb
CopyToRange(RangeObject As Excel.Range, Optional ApplyUserLocale As Boolean, Optional WriteHeaders As Boolean) As DictionaryEx
```
</td><td align="left" valign="top">
Internally converts it to a <code>dsTable</code> and calls it's <code>CopyToRange</code> method instead of implementing it's own.
<br/><em>@see: <code>dsTable.CopyToRange</code>.</em>
</td></tr>


<tr><td align="left" valign="top">

```vb
Remove(Key As Variant) As DictionaryEx
```
</td><td align="left" valign="top">
Removes a key-value pair from the <code>Scripting.Dictionary</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
Clear() As DictionaryEx
```
</td><td align="left" valign="top">
Removes all elements from the <code>Scripting.Dictionary</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
Exists(Key as Variant) As Boolean
```
</td><td align="left" valign="top">
Returns whether the specified <code>Key</code> exists in the <code>Scripting.Dictionary</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
Items() As Variant()
```
</td><td align="left" valign="top">
Returns an <code>Array</code> containing all values in the <code>Scripting.Dictionary</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
Keys() As Variant()
```
</td><td align="left" valign="top">
Returns an <code>Array</code> containing all existing keys in the <code>Scripting.Dictionary</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
Entries() As Variant()
```
</td><td align="left" valign="top">
Returns an <code>Array</code> containing all the key-value pairs in the <code>Scripting.Dictionary</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
Clone() As DictionaryEx
```
</td><td align="left" valign="top">
Returns a new <code>DictionaryEx</code> containing a shallow copy of all the key-value pairs in this <code>Scripting.Dictionary</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
Duplicate() As DictionaryEx
```
</td><td align="left" valign="top">
Returns a new <code>DictionaryEx</code> containing a deep copy of all the key-value pairs in this <code>Scripting.Dictionary</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
ToCSV(Optional Delimiter As String) As String
```
</td><td align="left" valign="top">
Returns a <code>String</code> representing this instance in <code>CSV</code>-style format.
<details><summary><code>PARAMETERS</code></summary><ul>
<li><kbd>Delimiter</kbd> → Character <code>String</code> used as delimiter between values within each row.<br/>Default value: <code>","</code></li>
</ul></details>
</td></tr>


<tr><td align="left" valign="top">

```vb
ToJSON() As String
```
</td><td align="left" valign="top">
Returns a <code>String</code> representing this instance in <code>JSON</code> format.
</td></tr>


<tr><td align="left" valign="top">

```vb
ToExcel() As String
```
</td><td align="left" valign="top">
Provides a simple way of direct copy-paste to an <code>Excel</code> document. <em>@see: <code>FileSystemLib.SystemClipboard</code>.</em>
<br/>Same as <code>.ToCSV(vbTab)</code>.
</td></tr>


</tbody>


<thead><tr><th colspan="2">PROCEDURES</th></tr></thead>
<tbody>


<tr><td align="left" valign="top">

```vb
Dispose()
```
</td><td align="left" valign="top">
Tells <code>DictionaryEx</code> that this instance won't be needed anymore and it can be safely disposed.
<br/>
<em>This shouldn't be necessary in most cases since all objects are automatically destroyed when there's no reference to them.</em>
</td></tr>


</tbody>

</table>




