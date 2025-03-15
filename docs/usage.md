# DS3X Usage Guide

## 1. Data Ingestion

### Excel/CSV Import
Load data from an Excel worksheet into a `dsTable` for high-performance processing:
```vba
Dim ds As dsTable
Set ds = dsTable.CreateFromExcelRange(Worksheets("Data").UsedRange)
```

### Access Table Import
Convert an ADODB.Recordset from an Access query into a `dsTable`:
```vba
Dim rs As ADODB.Recordset
Set rs = CurrentDb.OpenRecordset("SELECT * FROM Customers")
Set ds = dsTable.CreateFromRecordset(rs)
```

---

## 2. Data Transformation

### Filtering Rows
Keep only rows where "Sales > $1000":
```vba
Dim filteredDs As dsTable
Set filteredDs = ds.GetRange(0, ds.Count).Where("Sales > 1000")
```

### Joining Tables
Combine customer data with order details using a common key:
```vba
Dim joinedDs As dsTable
Set joinedDs = customersDs.Join(ordersDs, "CustomerID")
```

---

## 3. Live Editor Workflows

### Task Creation & Presets
Save a data cleaning workflow as a reusable preset file:
```vba
dsLiveEd.SavePreset "C:\Presets\clean_data.ds3x"
```

### Export Results
Export transformed data to Excel with formatting:
```vba
ds.CopyToExcel "Sheet1", "C:\Output\report.xlsx"
```

---

## 4. Performance Optimization

### Batch Processing
Process large datasets efficiently in chunks:
```vba
Dim chunkSize As Long: chunkSize = 10000
For i = 0 To ds.Count Step chunkSize
    ProcessChunk ds.GetRange(i, chunkSize)
Next
```

### Memory Management
Use `Array2dEx` for Excel interactions to avoid slow cell-by-cell updates:
```vba
Dim arr2d As Array2dEx
Set arr2d = Array2dEx.CreateFromExcelRange(ws.Range("A1:D100000"))
```

---

## Key Patterns from .clinerules
- **Always use**: `dsTable.CreateFromExcelRange` for Excel imports
- **Avoid**: Cell-by-cell Excel updates - use `Array2dEx` instead
- **For large datasets**: Use `ArrayListLib.CreateBlank` to initialize collections
