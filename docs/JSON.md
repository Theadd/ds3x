## **`JSON` Class**

_Backwards-compatible custom version of [Tim Hall](https://github.com/VBA-tools/VBA-JSON)'s `JSON` utilities._

---

### **API Overview**

```vb
' Static
JSON.Stringify(Value As Variant, Optional Whitespace As Variant, Optional CurrentIndentation As Long) As String
JSON.Parse(JSON As String, Optional CollectionsAsArrays As Boolean, Optional UnquotedKeysAllowed As Boolean) As Variant
```


<table width="100%"><caption>

### **`JSON` API**  
</caption>


<thead><tr><th colspan="2">STATIC</th></tr></thead>
<tbody>


<tr><td align="left" valign="top">

```vb
Stringify(Value As Variant, Optional Whitespace As Variant, Optional CurrentIndentation As Long) As String
```
</td><td align="left" valign="top">
Returns a <code>String</code> representing the specified <code>Value</code> in <b>JSON</b> format.
</td></tr>


<tr><td align="left" valign="top">

```vb
Parse(JSON As String, Optional CollectionsAsArrays As Boolean, Optional UnquotedKeysAllowed As Boolean) As Variant
```
</td><td align="left" valign="top">
Returns the parsed <b>JSON</b> <code>String</code> as a <code>Variant</code> value-type or object-type.
</td></tr>


</tbody>

</table>

