VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dsTaskDefinitions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder "ds3x.LiveEditor"
Option Compare Database
Option Explicit
Option Base 0

Private pTaskController As dsTaskController
Private pTasks As New DictionaryEx

Public Property Get TaskController() As dsTaskController: Set TaskController = pTaskController: End Property
Public Property Set TaskController(ByRef Value As dsTaskController): Set pTaskController = Value: End Property
Public Property Get Tasks() As DictionaryEx: Set Tasks = pTasks: End Property
Public Property Get AppExcelVisibilityAuto() As Boolean: AppExcelVisibilityAuto = False: End Property


' --- DEBUG / DEV ---

Public Function AsRecordsetTable(ByRef TargetTable As Variant) As dsTable
    Dim Target As dsTable, rX As RecordsetEx
    Set Target = TargetTable
    
    Set rX = RecordsetEx.Create(Target.NamedRecordset)
    Set AsRecordsetTable = pTaskController.AddResource(dsTable.Create(rX, True))
End Function

Public Function CreateMassiveTable() As dsTable
    Dim dsT As dsTable, t() As Variant, a2X As Array2dEx, i As Long, sX As ArraySlice, aX As ArrayListEx, v As Variant, sgX As ArraySliceGroup
    Dim NumRows As Long, NumRowSlices As Long, NumColumns As Long, r() As Byte
    NumRows = 42000: NumRowSlices = 50000: NumColumns = 1000
    ReDim r(0 To NumColumns - 1)
    ReDim t(0 To NumRows - 1, 0 To 0)
    
    For i = 0 To NumRows - 1
        t(i, 0) = i
    Next i
    Set a2X = New Array2dEx
    a2X.SetInstanceByRef t
    Set sX = ArraySlice.Create(a2X)
    Set aX = New ArrayListEx
    v = Array(0, NumRows)
    For i = 0 To NumRowSlices - 1
        aX.Add v
    Next i
    Set sX = sX.SliceAllRanges(aX)
    Set sgX = New ArraySliceGroup
    sgX.SliceGroups = Array(sX)
    Set sgX = sgX.GetRange(, , r)
    Set dsT = dsTable.Create(sgX)
    Set CreateMassiveTable = pTaskController.AddResource(dsT)
End Function

' ---

Public Function EnumerateRows(ByRef TargetTable As Variant, Optional ByVal FromNumber As Variant, Optional ByVal ToNumber As Variant) As dsTable
    Dim t() As Variant, i As Long, n As Long, Target As dsTable, FromNum As Long, ToNum As Long
    Set Target = TargetTable
    FromNum = CLng(GetParam(FromNumber, CLng(1)))
    ToNum = CLng(GetParam(ToNumber, CLng(-1)))
    ReDim t(0 To Target.Count - 1, 0 To 0)
    
    n = FromNum
    For i = 0 To Target.Count - 1
        t(i, 0) = n
        If n = ToNum Then
            n = FromNum
        Else
            n = n + 1
        End If
    Next i
    
    Set EnumerateRows = dsTable _
        .Create(Array2dEx.Create(t)) _
        .SetHeaders(ArrayListEx.Create(Array(Array(dsTable.CreateColumnHeader("#"))))) _
        .Join(Target)
End Function

Public Function SetColumnName(ByRef TargetTable As Variant, Optional ByVal ColumnIndexes As Variant, Optional ByVal ColumnName As Variant) As dsTable
    Dim Target As dsTable, e As Long, colIndexes As ArrayListEx
    Dim sgX As ArraySliceGroup, sX As ArraySlice, v As Variant, cH As Variant, vItems As Variant
    ColumnIndexes = GetParam(ColumnIndexes, Array())
    ColumnName = GetParam(ColumnName, "%1")
    
    Set Target = TargetTable
    Set colIndexes = AsListOfColumnIndexes(Target, ColumnIndexes)
    Set sgX = ArraySliceGroup.Create(Target.Headers)
    
    For Each cH In colIndexes
        v = sgX.SliceAt(CLng(cH)).Row(0)
        vItems = Array(dsApp.ApplyCustomVarsOn(Printf(ColumnName, cH, e, cH + 1, e + 1)), CollectionsLib.ArrayItem(v, 1), CollectionsLib.ArrayItem(v, 2), CollectionsLib.ArrayItem(v, 3))
        Set sgX = sgX.Replace(ArraySlice.Create(Array2dEx.Create(Array(Array(vItems)))), CLng(cH))
        e = e + 1
    Next cH
    
    Set SetColumnName = dsTable.Create(Target.Records).SetHeaders(sgX.Instance)
End Function

Public Function SetNumberFormat(ByRef TargetTable As Variant, Optional ByVal ColumnIndexes As Variant, Optional ByVal NumberFormat As Variant) As dsTable
    Dim Target As dsTable, e As Long, colIndexes As ArrayListEx
    Dim sgX As ArraySliceGroup, v As Variant, cH As Variant, vItems As Variant
    ColumnIndexes = GetParam(ColumnIndexes, Array())
    NumberFormat = GetParam(NumberFormat, "General")
    
    Set Target = TargetTable
    Set colIndexes = AsListOfColumnIndexes(Target, ColumnIndexes)
    Set sgX = ArraySliceGroup.Create(Target.Headers)
    
    For Each cH In colIndexes
        v = sgX.SliceAt(CLng(cH)).Row(0)
        vItems = Array(CollectionsLib.ArrayItem(v, 0), dsApp.ApplyCustomVarsOn(Printf(NumberFormat, cH, e, cH + 1, e + 1)), CollectionsLib.ArrayItem(v, 2), CollectionsLib.ArrayItem(v, 3))
        Set sgX = sgX.Replace(ArraySlice.Create(Array2dEx.Create(Array(Array(vItems)))), CLng(cH))
        e = e + 1
    Next cH
    
    Set SetNumberFormat = dsTable.Create(Target.Records).SetHeaders(sgX.Instance)
End Function

'RecordsetEx.Filter("REPORT_DATE < '20/04/2023' AND REPORT_STATUS = 'RESUELTA'")
Public Function FilterQuery(ByRef TargetTable As Variant, Optional ByVal QueryFilter As Variant) As dsTable
    Dim rX As RecordsetEx, Target As dsTable
    Set Target = TargetTable
    QueryFilter = GetParam(QueryFilter, vbNullString)

    Set rX = RecordsetEx.Create(Target.NamedRecordset).Filter(QueryFilter)
    Set FilterQuery = dsTable.Create(rX).SetHeaders(Target.Headers)
End Function

Public Function FilterValue(ByRef TargetTable As Variant, ByVal ColumnIndex As Variant, Optional ByVal Value As Variant, Optional ByVal FilterOut As Variant) As dsTable
    Dim Target As dsTable, t() As Variant, aX As ArrayListEx, sgX As ArraySliceGroup, i As Long, sValue As String
    Set Target = TargetTable
    FilterOut = GetParam(FilterOut, False)
    Value = GetParam(Value, vbNullString)
    sValue = CStr(Value)

    Set sgX = CollectionsLib.AsArraySliceGroup(Target.Records)
    Set aX = ArrayListEx.Create()
    sgX.SliceAt(CLng(ColumnIndex)).SlicedArray t
    If FilterOut Then
        For i = 0 To UBound(t, 1)
            If CStr(Nz(t(i, 0), vbNullString)) <> sValue Then aX.Add i
        Next i
    Else
        For i = 0 To UBound(t, 1)
            If CStr(Nz(t(i, 0), vbNullString)) = sValue Then aX.Add i
        Next i
    End If
    
    Set FilterValue = dsTable.Create(sgX.GetRows(aX.ToArray()), False).SetHeaders(Target.Headers)
End Function

Public Function GetRange(ByRef TargetTable As Variant, Optional ByVal Index As Variant, Optional ByVal GetCount As Variant, Optional ByVal ColumnIndexes As Variant) As dsTable
    Dim Target As dsTable
    Set Target = TargetTable
    Index = GetParam(Index, 0)
    GetCount = GetParam(GetCount, Target.Count - Index)
    ColumnIndexes = GetParam(ColumnIndexes, Array())
    
    If Not IsArray(ColumnIndexes) Then ColumnIndexes = Array()
    
    If ArrayListEx.CountElementsIn(ColumnIndexes) > 0 Then
        Set GetRange = pTaskController.AddResource(Target.GetRange(Index, GetCount, ColumnIndexes))
    Else
        Set GetRange = pTaskController.AddResource(Target.GetRange(Index, GetCount))
    End If
End Function

Public Function Join(ByRef TargetTable As Variant, ByRef OtherTable As Variant) As dsTable
    Dim Target As dsTable, Other As dsTable
    Set Target = TargetTable
    Set Other = OtherTable

    Set Join = pTaskController.AddResource(Target.Join(Other))
End Function

Public Function Append(ByRef SourceTable As Variant, ByRef TargetTable As Variant) As dsTable
    Dim Target As dsTable, Source As dsTable
    Set Target = TargetTable
    Set Source = SourceTable

    Set Append = pTaskController.AddResource(Source.AddRange(Target))
End Function

' CallbackFunction can also be a global function, for instance: Callback.Create(, "MillisecondsToTime")
Public Function ReplaceValues(ByRef TargetTable As Variant, ByRef Replacer As Variant, Optional ByVal ColumnIndexes As Variant) As dsTable
    Dim t() As Variant, i As Long, iMax As Long, c As Long, cMax As Long, Target As dsTable, cb As Callback, v As Variant
    Dim dRepl As DictionaryEx, colIndexes As New ArrayListEx, sgX As ArraySliceGroup, vSliced() As Variant, r As Variant
    Set Target = TargetTable
    ColumnIndexes = GetParam(ColumnIndexes, Array())
    iMax = Target.Count - 1
    cMax = Target.ColumnCount - 1
    ReDim t(0 To iMax, 0 To cMax)
    If ArrayListEx.CountElementsIn(ColumnIndexes) = 0 Then
        ColumnIndexes = CollectionsLib.ArrayRange(0, cMax + 1)
    End If
    For Each v In ColumnIndexes
        colIndexes.Add CLng(v)
    Next v
    Set sgX = CollectionsLib.AsArraySliceGroup(Target.Records)
    
    If TypeName(Replacer) Like "*Dictionary*" Then
        Set dRepl = DictionaryEx.Create(Replacer)
        For Each v In colIndexes
            c = CLng(v)
            sgX.SliceAt(c).SlicedArray vSliced
            t = vSliced
            For i = 0 To iMax
                r = t(i, 0)
                If Not IsNull(r) Then
                    t(i, 0) = dRepl.GetValue(r, r)
                End If
            Next i
            Set sgX = sgX.Replace(ArraySlice.Create(t), c)
        Next v
    Else
        Set cb = Replacer
        For Each v In colIndexes
            c = CLng(v)
            sgX.SliceAt(c).SlicedArray vSliced
            t = vSliced
            For i = 0 To iMax
                r = t(i, 0)
                If Not IsNull(r) Then
                    t(i, 0) = cb.Execute(r).Result
                End If
            Next i
            Set sgX = sgX.Replace(ArraySlice.Create(t), c)
        Next v
    End If
    
    Set ReplaceValues = pTaskController.AddResource(dsTable.Create(sgX, False).SetHeaders(Target.Headers.GetRange(0, Target.ColumnCount)))
End Function

Public Function StringValuesToHex(ByRef TargetTable As Variant, Optional ByVal ColumnIndexes As Variant) As dsTable
    Dim t() As Variant, i As Long, iMax As Long, c As Long, k As Long, kMax As Long, Target As dsTable, sgX As ArraySliceGroup, vSliced() As Variant
    Set Target = TargetTable
    ColumnIndexes = GetParam(ColumnIndexes, Array())
    iMax = Target.Count - 1
    If ArrayListEx.CountElementsIn(ColumnIndexes) = 0 Then
        ColumnIndexes = CollectionsLib.ArrayRange(0, Target.ColumnCount)
    End If
    kMax = ArrayListEx.CountElementsIn(ColumnIndexes) - 1
    Set sgX = CollectionsLib.AsArraySliceGroup(Target.Records)
    
    For k = 0 To kMax
        c = CLng(ColumnIndexes(k))
        sgX.SliceAt(c).SlicedArray vSliced
        t = vSliced
        For i = 0 To iMax
            If VarType(t(i, 0)) = vbString Then
                t(i, 0) = AsciiToHex(t(i, 0))
            End If
        Next i
        Set sgX = sgX.Replace(ArraySlice.Create(t), c)
    Next k
    
    Set StringValuesToHex = pTaskController.AddResource(dsTable.Create(sgX, False).SetHeaders(Target.Headers.GetRange(0, Target.ColumnCount)))
End Function

Public Function HexToStringValues(ByRef TargetTable As Variant, Optional ByVal ColumnIndexes As Variant) As dsTable
    Dim t() As Variant, i As Long, iMax As Long, c As Long, k As Long, kMax As Long, Target As dsTable, sgX As ArraySliceGroup, vSliced() As Variant
    Set Target = TargetTable
    ColumnIndexes = GetParam(ColumnIndexes, Array())
    iMax = Target.Count - 1
    If ArrayListEx.CountElementsIn(ColumnIndexes) = 0 Then
        ColumnIndexes = CollectionsLib.ArrayRange(0, Target.ColumnCount)
    End If
    kMax = ArrayListEx.CountElementsIn(ColumnIndexes) - 1
    Set sgX = CollectionsLib.AsArraySliceGroup(Target.Records)
    
    For k = 0 To kMax
        c = CLng(ColumnIndexes(k))
        sgX.SliceAt(c).SlicedArray vSliced
        t = vSliced
        For i = 0 To iMax
            If VarType(t(i, 0)) = vbString Then
                t(i, 0) = HexToAscii(t(i, 0))
            End If
        Next i
        Set sgX = sgX.Replace(ArraySlice.Create(t), c)
    Next k
    
    Set HexToStringValues = pTaskController.AddResource(dsTable.Create(sgX, False).SetHeaders(Target.Headers.GetRange(0, Target.ColumnCount)))
End Function

' TODO: Convert to the new dsTable.Headers implementation
'Public Function ReplaceColumnNames(ByRef TargetTable As Variant, ByRef Replacer As Variant) As dsTable
'    Dim t() As Variant, i As Long, cH As Variant, Target As dsTable, d As Scripting.Dictionary, dX As New DictionaryEx
'    Dim cb As Callback, dRepl As DictionaryEx
'    Set Target = TargetTable
'
'    cH = Target.Headers.Row(0)
'    ReDim t(0 To 0, 0 To Target.ColumnCount - 1)
'
'    If TypeName(Replacer) Like "*Dictionary*" Then
'        Set dRepl = DictionaryEx.Create(Replacer)
'        For i = 0 To Target.ColumnCount - 1
'            Set d = cH(i)
'            Set dX = DictionaryEx.Bind(d).Clone()
'
'            dX.Add "ColumnName", dRepl.GetValue(dX.Item("ColumnName"), dX.Item("ColumnName"))
'            Set t(0, i) = dX.Instance
'        Next i
'    Else
'        Set cb = Replacer
'        For i = 0 To Target.ColumnCount - 1
'            Set d = cH(i)
'            Set dX = DictionaryEx.Bind(d).Clone()
'
'            dX.Add "ColumnName", cb.Execute(dX.Item("ColumnName")).Result
'            Set t(0, i) = dX.Instance
'        Next i
'    End If
'
'    Set ReplaceColumnNames = pTaskController.AddResource(dsTable.Create(Target.Records).SetHeaders(Array2dEx.Create(t)))
'End Function

Public Function RemoveColumns(ByRef TargetTable As Variant, ByVal ColumnIndexes As Variant) As dsTable
    Dim t() As Variant, i As Long, Target As dsTable, aX As ArrayListEx, c As Variant
    Set Target = TargetTable
    
    Set aX = ArrayListEx.Create(CollectionsLib.ArrayRange(0, ColumnCount(Target)))
    For Each c In ColumnIndexes
        aX.Remove CLng(c)
    Next c
    t = aX.ToArray
    
    Set RemoveColumns = Target.GetRange(, , t)
End Function

Public Function MoveColumnsToTheEnd(ByRef TargetTable As Variant, ByVal ColumnIndexes As Variant) As dsTable
    Dim t() As Variant, i As Long, Target As dsTable, aX As ArrayListEx, c As Variant, cX As ArrayListEx
    Dim TableA As dsTable, TableB As dsTable
    
    Set Target = TargetTable
    Set TableA = RemoveColumns(Target, ColumnIndexes)
    Set TableB = Target.GetRange(, , ColumnIndexes)

    Set MoveColumnsToTheEnd = TableA.Join(TableB)
End Function

Public Function MoveColumnsToTheStart(ByRef TargetTable As Variant, ByVal ColumnIndexes As Variant) As dsTable
    Dim TableA As dsTable, TableB As dsTable, Target As dsTable
    Set Target = TargetTable
    
    Set TableA = RemoveColumns(Target, ColumnIndexes)
    Set TableB = Target.GetRange(, , ColumnIndexes)

    Set MoveColumnsToTheStart = TableB.Join(TableA)
End Function

Public Function MoveColumn(ByRef TargetTable As Variant, ByVal ColumnIndex As Variant, Optional ByVal MoveCount As Variant) As dsTable
    Dim Target As dsTable, aX As ArrayListEx
    Set Target = TargetTable
    MoveCount = CLng(Val(GetParam(MoveCount, 1)))

    Set aX = ArrayListEx.Create(CollectionsLib.ArrayRange(0, ColumnCount(Target)))
    aX.RemoveAt CLng(ColumnIndex)
    aX.Insert ColumnIndex + MoveCount, CLng(ColumnIndex)
    
    Set MoveColumn = Target.GetRange(, , aX.ToArray)
End Function



' --- DATA SOURCE TASKS ---

Public Function DataSource() As dsTable
    With TaskController()
        Set DataSource = .AddResource(.DataSource(.RebuildSequence(.SequenceIndex())("Id")))
    End With
End Function

Public Function WorksheetSource() As xlSheetsEx
    With TaskController()
        Set WorksheetSource = .AddResource(.DataSource(.RebuildSequence(.SequenceIndex())("Id")))
    End With
End Function

Public Function LoadFromExcelFile(ByVal TargetFile As Variant, Optional ByVal WorksheetNameOrIndex As Variant, Optional ByVal FirstRowHasHeaders As Variant, Optional ByVal Options As Variant, Optional ByVal DisableCache As Variant) As dsTable
    Dim dsT As dsTable, xlX As xlSheetsEx, xlRange As Excel.Range
    DisableCache = GetParam(DisableCache, False)
    
    If Not DisableCache Then
        On Error GoTo LoadTableFromFileFallback
        Set dsT = TaskController().DataSource(ActiveTaskId())
        If IsObject(dsT) And Not (dsT Is Nothing) Then
            Set LoadFromExcelFile = dsT
            Exit Function
        End If
    End If
    
LoadTableFromFileFallback:
    On Error GoTo 0
    FirstRowHasHeaders = GetParam(FirstRowHasHeaders, True)
    WorksheetNameOrIndex = GetParam(WorksheetNameOrIndex, 0)
    
    Set xlX = OpenWorksheetFromFile(TargetFile, Options)
    On Error GoTo Finally
    If WorksheetNameOrIndex <> 0 Then
        Set xlX = GetWorksheet(xlX, WorksheetNameOrIndex)
    End If
    Set dsT = LoadFromRange(xlX, xlX.UsedRange, FirstRowHasHeaders)
    If Not DisableCache Then
        Set TaskController().DataSource(ActiveTaskId()) = dsT
    End If
Finally:
    On Error Resume Next
    xlX.DiscardChangesAndExit
    Set LoadFromExcelFile = dsT
    On Error GoTo 0
End Function

Public Function OpenWorksheetFromFile(ByVal TargetFile As Variant, Optional ByVal Options As Variant) As xlSheetsEx
    Dim xlSheet As xlSheetsEx, dX As DictionaryEx
    
    Set Options = GetParamAsDictionaryEx(Options)
    ' {"UpdateLinks": false, "ReadOnly": true, "Local": true, "UTF8": false, "NoTextQualifier": true}
    Set dX = DictionaryEx.Create() _
                         .Add("UpdateLinks", Options.GetValue("UpdateLinks", False)) _
                         .Add("ReadOnly", Options.GetValue("ReadOnly", True)) _
                         .Add("Local", Options.GetValue("Local", True)) _
                         .Add("UTF8", Options.GetValue("UTF8", False)) _
                         .Add("NoTextQualifier", Options.GetValue("NoTextQualifier", True))
                         
    Set xlSheet = xlSheetsEx.Create(, TargetFile, dX)
    xlSheet.WindowVisibility = AppExcelVisibilityAuto

    Set OpenWorksheetFromFile = pTaskController.AddResource(xlSheet)
End Function

Public Function LoadFromRange(ByRef SourceSheet As Variant, Optional ByVal RangeAddress As Variant, Optional ByVal FirstRowHasHeaders As Variant) As dsTable
    Dim xlSheet As xlSheetsEx, xlRange As Excel.Range
    Set xlSheet = SourceSheet
    FirstRowHasHeaders = GetParam(FirstRowHasHeaders, True)
    
    If IsObject(RangeAddress) Then
        Set xlRange = RangeAddress
    Else
        Set xlRange = GetRangeInWorksheet(xlSheet, RangeAddress)
    End If

    With xlRange
        If FirstRowHasHeaders Then
            Set LoadFromRange = pTaskController.AddResource(dsTable.Create(xlRange, True))
        Else
            Set LoadFromRange = pTaskController.AddResource(dsTable _
                .Create(xlRange, False) _
                .SetHeaders(CollectionsLib.ArrayRange(1, .Columns.Count)))
        End If
    End With
End Function

Public Function LoadFromFileAsJSON(ByVal TargetFile As Variant, Optional ByVal DisableCache As Variant) As dsTable
    Dim dsT As dsTable, sContent As String, sTarget As String
    DisableCache = GetParam(DisableCache, False)
    
    If Not DisableCache Then
        On Error GoTo LoadFromFileAsJSONFallback
        Set dsT = TaskController().DataSource(ActiveTaskId())
        If IsObject(dsT) And Not (dsT Is Nothing) Then
            Set LoadFromFileAsJSON = dsT
            Exit Function
        End If
    End If
    
LoadFromFileAsJSONFallback:
    On Error GoTo 0
    TargetFile = GetParam(TargetFile, "")
    sTarget = FileSystemLib.GetFileMatchingPattern(FileSystemLib.Resolve(TargetFile))
    
    If Not FileSystemLib.TryReadAllTextInFile(sTarget, sContent) Then
        Err.Raise 17, , "Unable to read text from provided file"
    End If
    Set dsT = dsTable.Create(sContent)
    If Not DisableCache Then
        Set TaskController().DataSource(ActiveTaskId()) = dsT
    End If

    Set LoadFromFileAsJSON = pTaskController.AddResource(dsT)
End Function

Public Function LoadFromFileAsCSV(ByVal TargetFile As Variant, Optional ByVal Options As Variant, Optional ByVal DisableCache As Variant) As dsTable
    Dim dsT As dsTable, sContent As String, sTarget As String, dX As DictionaryEx
    DisableCache = GetParam(DisableCache, False)
    Set Options = GetParamAsDictionaryEx(Options)

    ' {"Delimiter": "Auto", "InLocalFormat": true, "AutoHeaders": true, "TextQualifier": """"}
    Set dX = DictionaryEx.Create() _
        .Add("Delimiter", Options.GetValue("Delimiter", "Auto")) _
        .Add("InLocalFormat", Options.GetValue("InLocalFormat", False)) _
        .Add("AutoHeaders", Options.GetValue("AutoHeaders", True)) _
        .Add("TextQualifier", Options.GetValue("TextQualifier", """"))
    
    If Not DisableCache Then
        On Error GoTo LoadFromFileAsCSVFallback
        Set dsT = TaskController().DataSource(ActiveTaskId())
        If IsObject(dsT) And Not (dsT Is Nothing) Then
            Set LoadFromFileAsCSV = dsT
            Exit Function
        End If
    End If
    
LoadFromFileAsCSVFallback:
    On Error GoTo 0
    TargetFile = GetParam(TargetFile, "")
    sTarget = FileSystemLib.GetFileMatchingPattern(FileSystemLib.Resolve(TargetFile))
    
    Set dsT = LoadFromCSV(sTarget, dX("Delimiter"), dX("InLocalFormat"), dX("AutoHeaders"), dX("TextQualifier"))
    If Not DisableCache Then
        Set TaskController().DataSource(ActiveTaskId()) = dsT
    End If

    Set LoadFromFileAsCSV = pTaskController.AddResource(dsT)
End Function

Public Function LoadFromSQL(ByVal QueryString As Variant, Optional ByVal ConnectionString As Variant, Optional ByVal DisableCache As Variant) As dsTable
    Dim dsT As dsTable, dbQ As dsQuery
    DisableCache = GetParam(DisableCache, False)
    
    If Not DisableCache Then
        On Error GoTo LoadFromSQLFallback
        Set dsT = TaskController().DataSource(ActiveTaskId())
        If IsObject(dsT) And Not (dsT Is Nothing) Then
            Set LoadFromSQL = dsT
            Exit Function
        End If
    End If
    
LoadFromSQLFallback:
    On Error GoTo 0
    ConnectionString = GetParam(ConnectionString, "")
    Set dbQ = New dsQuery
    If Trim(CStr(ConnectionString)) <> "" Then dbQ.ConnectionString = CStr(ConnectionString)
    dbQ.CloseOnTerminate = False
    pTaskController.AddResource dbQ
    Set dsT = dsTable.Create(dbQ(CStr(QueryString)), True)
    If Not DisableCache Then
        Set TaskController().DataSource(ActiveTaskId()) = dsT
    End If

    Set LoadFromSQL = pTaskController.AddResource(dsT)
End Function


' --- GENERATION TASKS ---

Public Function SaveToFileAsJSON(ByRef SourceTable As Variant, ByVal FilePath As Variant) As dsTable
    Dim Target As dsTable
    FilePath = GetParam(FilePath, "")
    Set Target = SourceTable

    If Not FileSystemLib.TryWriteTextToFile(FileSystemLib.Resolve(FilePath), JSON.Stringify(Target, 2), True, False) Then
        Err.Raise 17, , "Failed to save dsTable as JSON to """ & CStr(FilePath) & """"
    End If
    Set SaveToFileAsJSON = Target
End Function

Public Function SaveToFileAsCSV(ByRef SourceTable As Variant, ByVal FilePath As Variant, Optional ByVal Delimiter As Variant, Optional ByVal InLocalFormat As Variant, Optional ByVal OverWriteIfExists As Variant) As dsTable
    Dim Target As dsTable
    FilePath = GetParam(FilePath, "")
    Delimiter = GetParam(Delimiter, ",")
    InLocalFormat = GetParam(InLocalFormat, False)
    OverWriteIfExists = GetParam(OverWriteIfExists, True)
    Set Target = SourceTable
    
    With FileSystemLib.StreamWriter()
        .WriteText Target.ToCSV(Delimiter, InLocalFormat)
        If OverWriteIfExists Then
            .SaveToFile FileSystemLib.Resolve(FilePath), adSaveCreateOverWrite
        Else
            .SaveToFile FileSystemLib.Resolve(FilePath), adSaveCreateNotExist
        End If
        .Close
    End With

    Set SaveToFileAsCSV = Target
End Function


' --- XL GENERATION TASKS ---

Public Function CreateNewExcelApp(Optional ByVal WorksheetName As Variant) As xlSheetsEx
    Dim xlSheet As xlSheetsEx
    WorksheetName = GetParam(WorksheetName, "Sheet%1")
    
    Set xlSheet = xlSheetsEx.Create(WorksheetName)
    xlSheet.WindowVisibility = AppExcelVisibilityAuto
    
    Set CreateNewExcelApp = pTaskController.AddResource(xlSheet)
End Function

Public Function CopyToRangeInNewExcelApp(ByRef SourceTable As Variant, Optional ByVal RangeAddress As Variant, Optional ByVal WorksheetName As Variant) As xlSheetsEx
    Dim Target As dsTable, xlSheet As xlSheetsEx, xlRange As Excel.Range
    Assign RangeAddress, GetParam(RangeAddress, "A1")
    WorksheetName = GetParam(WorksheetName, "Sheet%1")
    
    Set Target = SourceTable
    Set xlSheet = pTaskController.AddResource(xlSheetsEx.Create(WorksheetName))
    xlSheet.WindowVisibility = AppExcelVisibilityAuto
    Set xlRange = xlSheet.Range(RangeAddress)
    Target.CopyToRange xlRange, True
    
    Set CopyToRangeInNewExcelApp = xlSheet
End Function

Public Function CopyToRange(ByRef SourceTable As Variant, ByRef TargetSheet As Variant, Optional ByVal RangeAddress As Variant, Optional ByVal WriteHeaders As Variant, Optional ByVal ApplyUserLocale As Variant) As xlSheetsEx
    Dim Target As dsTable, xlSheet As xlSheetsEx, xlRange As Excel.Range
    Assign RangeAddress, GetParam(RangeAddress, "A1")
    WriteHeaders = GetParam(WriteHeaders, True)
    ApplyUserLocale = GetParam(ApplyUserLocale, True)
    
    Set Target = SourceTable
    Set xlSheet = TargetSheet
    If IsObject(RangeAddress) Then
        Set xlRange = RangeAddress
    Else
        Set xlRange = xlSheet.Range(RangeAddress)
    End If
    Target.CopyToRange xlRange, CBool(ApplyUserLocale), CBool(WriteHeaders)
    
    Set CopyToRange = xlSheet
End Function

Public Function CopyRange(ByRef SourceSheet As Variant, ByRef TargetSheet As Variant, Optional ByVal SourceRangeAddress As Variant, Optional ByVal TargetRangeAddress As Variant) As xlSheetsEx
    Dim Source As xlSheetsEx, sourceRange As Excel.Range, xlSheet As xlSheetsEx, xlRange As Excel.Range
    Assign TargetRangeAddress, GetParam(TargetRangeAddress, "A1")
    
    Set Source = SourceSheet
    Source.Instance.Select
    If IsObject(SourceRangeAddress) Then
        Set sourceRange = SourceRangeAddress
    Else
        Set sourceRange = GetRangeInWorksheet(Source, SourceRangeAddress)
    End If
    sourceRange.Copy
    
    Set xlSheet = TargetSheet
    xlSheet.Instance.Select
    If IsObject(TargetRangeAddress) Then
        Set xlRange = TargetRangeAddress
    Else
        Set xlRange = xlSheet.Range(TargetRangeAddress)
    End If
    xlRange.Select
    xlSheet.Instance.Paste
    
    Set CopyRange = xlSheet
End Function

Public Function AddWorksheet(ByRef SourceSheet As Variant, Optional ByVal WorksheetName As Variant) As xlSheetsEx
    Dim xlSheet As xlSheetsEx
    WorksheetName = GetParam(WorksheetName, "Sheet%1")
    
    Set xlSheet = SourceSheet
    Set AddWorksheet = pTaskController.AddResource(xlSheet.AddSheet(WorksheetName))
End Function

Public Function GetWorksheet(ByRef SourceSheet As Variant, ByVal WorksheetNameOrIndex As Variant) As xlSheetsEx
    Dim xlSheet As xlSheetsEx
    
    Set xlSheet = SourceSheet
    Set GetWorksheet = pTaskController.AddResource(xlSheet.GetSheet(WorksheetNameOrIndex))
End Function

Public Function SaveWorkbook(ByRef SourceSheet As Variant, ByVal FilePath As Variant) As xlSheetsEx
    Dim xlSheet As xlSheetsEx
    
    Set xlSheet = SourceSheet
    Set SaveWorkbook = pTaskController.AddResource(xlSheet.SaveWorkbook(CStr(FileSystemLib.Resolve(FilePath))))
End Function



' --- XL FORMATTING TASKS ---

Public Function AutoSizeCells(ByRef SourceSheet As Variant, Optional ByVal RangeAddress As Variant) As xlSheetsEx
    Dim xlSheet As xlSheetsEx, xlRange As Excel.Range
    Set xlSheet = SourceSheet
    
    If IsObject(RangeAddress) Then
        Set xlRange = RangeAddress
    Else
        Set xlRange = GetRangeInWorksheet(xlSheet, RangeAddress)
    End If
    
    Set AutoSizeCells = xlSheet.AutoSizeCells(xlRange)
End Function

Public Function FormatAsTable(ByRef SourceSheet As Variant, Optional ByVal TableName As Variant, Optional ByVal TableStyle As Variant) As xlSheetsEx
    Dim xlSheet As xlSheetsEx
    Set xlSheet = SourceSheet
    TableName = GetParam(TableName, "xlSheetTable%1")
    TableStyle = GetParam(TableStyle, "TableStyleMedium6")

    Set FormatAsTable = xlSheet.FormatAsTable(TableName, TableStyle)
End Function

Public Function FreezeHeaders(ByRef SourceSheet As Variant, Optional ByVal HeaderRows As Variant, Optional ByVal HeaderColumns As Variant) As xlSheetsEx
    Dim xlSheet As xlSheetsEx
    Set xlSheet = SourceSheet
    HeaderRows = GetParam(HeaderRows, 1)
    HeaderColumns = GetParam(HeaderColumns, 0)

    Set FreezeHeaders = xlSheet.FreezeHeaders(HeaderRows, HeaderColumns)
End Function

Public Function AutoFormatCells(ByRef SourceSheet As Variant, Optional ByVal RangeAddress As Variant) As xlSheetsEx
    Dim xlSheet As xlSheetsEx, xlRange As Excel.Range
    Set xlSheet = SourceSheet
    
    If IsObject(RangeAddress) Then
        Set xlRange = RangeAddress
    Else
        Set xlRange = GetRangeInWorksheet(xlSheet, RangeAddress)
    End If
    
    Set AutoFormatCells = xlSheet.AutoFormatCells(xlRange)
End Function

Public Function WindowVisibility(ByRef SourceSheet As Variant, Optional ByVal Visible As Variant) As xlSheetsEx
    Dim xlSheet As xlSheetsEx
    Set xlSheet = SourceSheet

    xlSheet.WindowVisibility = GetParam(Visible, True)
    
    Set WindowVisibility = xlSheet
End Function




' --- PRIVATE ---

Private Sub Class_Initialize()
    Set pTasks.Instance = New Scripting.Dictionary
    AddAllTasks
End Sub

Private Function StringLines(ParamArray Lines() As Variant) As String
    StringLines = VBA.Join(Lines, " " & vbNewLine)
    StringLines = Replace(StringLines, "@param ", "<font face=""Consolas"" size=1 color=""#33CC33"">@param</font> ")
End Function

Private Sub AddTask(ByVal TaskName As String, _
                    ByVal TaskParameters As String, _
                    Optional ByVal Description As String = "", _
                    Optional ByVal Usage As String = "", _
                    Optional ByVal TaskType As String = "T", _
                    Optional ByVal Requires As Variant)
    If IsMissing(Requires) Then
        Select Case Left(TaskType, 1)
            Case "T": Requires = "[TD]*"
            Case "D": Requires = "[TD]*"
            Case "X": Requires = "X*"
            Case Else: Stop
        End Select
    End If
    pTasks.Add TaskName, Array(TaskName, TokenizeParameters(TaskParameters), Description, Usage, TaskType, Requires)
End Sub

Private Function TokenizeParameters(ByVal TaskParameters As String) As Variant()
    Dim t() As String, Tokens() As Variant, i As Long
    
    t = VBA.Split(TaskParameters, ", ")
    If Not UBound(t) >= 0 Then
        TokenizeParameters = Array("", 0, Array())
        Exit Function
    End If
    ReDim Tokens(UBound(t))
    
    For i = LBound(t) To UBound(t)
        Tokens(i) = VBA.Split(t(i), ": ", 2)
    Next i
    
    TokenizeParameters = Array(TaskParameters, 1 + UBound(Tokens) - LBound(Tokens), Tokens)
End Function

Private Sub AddAllTasks()

    AddTask "EnumerateRows", "[Source]: Table, [FromNumber = 1]: Long, [ToNumber = -1]: Long", _
            Description:=StringLines( _
                "Inserts a new column at the beginning of the table by enumerating its values from <b>FromNumber</b> to <b>ToNumber</b> repeatedly until the end. Ignores the value of <b>ToNumber</b> if it is less than <b>FromNumber</b>.")
    
    AddTask "SetColumnName", "[Source]: Table, [ColumnIndexes]: Long|Array(), [ColumnName]: String", _
            Description:=StringLines( _
                "@param [<b>ColumnIndexes</b>] - A column index position or an array of indexes. When no value is provided, all columns are included.", _
                "@param [<b>ColumnName</b>] - The string to use for renaming that/those column/s. Supports <b>CustomVar</b>s and from <b>%1</b> to <b>%4</b>, being:", _
                "    <b>%1</b>: The column index.", _
                "    <b>%2</b>: the column index within the <b>ColumnIndexes</b> array.", _
                "    <b>%3</b> and <b>%4</b>: Position instead of index for <b>%1</b> and <b>%2</b> respectively.")
    
    AddTask "SetNumberFormat", "[Source]: Table, [ColumnIndexes]: Long|Array(), [NumberFormat]: String", _
            Description:=StringLines( _
                "Select a single Column before picking this task to see it's current <b>NumberFormat</b>.", _
                "", _
                "@param [<b>ColumnIndexes</b>] - A column index position or an array of indexes. When no value is provided, all columns are included.", _
                "@param [<b>NumberFormat</b>] - Excel's <b>NumberFormat</b> string. Supports <b>CustomVar</b>s and from <b>%1</b> to <b>%4</b>, being:", _
                "    <b>%1</b>: The column index.", _
                "    <b>%2</b>: the column index within the <b>ColumnIndexes</b> array.", _
                "    <b>%3</b> and <b>%4</b>: Position instead of index for <b>%1</b> and <b>%2</b> respectively.", _
                "", _
                "NOTE: This task will&nbsp;<u>only</u> apply <b>NumberFormat</b> to corresponding columns in an <b>Excel.Worksheet</b> <u>just before</u> copying the data to the specified <b>Excel.Range</b>.")
                    
    AddTask "FilterQuery", "[Source]: Table, QueryFilter: String", _
            Usage:="{...TaskName: 'FilterQuery', Values: [, ""REPORT_DATE < '20/04/2023' AND REPORT_STATUS = 'RESUELTA'""]}"
    
    AddTask "FilterValue", "[Source]: Table, ColumnIndex: Long, [Value]: Variant, [FilterOut = False]: Boolean"
    
    AddTask "GetRange", "[Source]: Table, [Index]: Long, [GetCount]: Long, [ColumnIndexes]: Array()", _
            Description:=StringLines( _
                "Returns a subset of rows and/or columns from the Source Table. The ColumnIndexes Array can contain sparse elements.", _
                "", _
                "@param [Index] - The zero-based row index at which the range starts.", _
                "@param [GetCount] - The number of rows in the range.", _
                "@param [ColumnIndexes] - If specified, an array of column indexes to include in the range, otherwise, no columns are excluded.")

    AddTask "Join", "[Source]: Table, Target: Table", _
            Description:="Concatenate all elements of <b>Target</b> Table as additional columns into <b>Source</b> Table."
            
    AddTask "Append", "[Source]: Table, Target: Table", _
            Description:="Appends all rows in <b>Target</b> Table the end of <b>Source</b> Table."

    AddTask "ReplaceValues", "[Source]: Table, Replacer: Callable|Dictionary, [ColumnIndexes]: Array()", _
            Description:=StringLines( _
                "Returns a copy of the provided Source Table with the values within the provided ColumnIndexes being replaced by the provided Replacer.", _
                "", _
                "@param Replacer - A name to a global function or a Scripting.Dictionary in JSON format.", _
                "@param [ColumnIndexes] - If specified, an array of column indexes to include in the range, otherwise, all columns are included."), _
            Usage:="{...TaskName: 'ReplaceValues', Values: [, 'SecondsToHMS']}"
            
    AddTask "StringValuesToHex", "[Source]: Table, [ColumnIndexes]: Array()", _
            Description:=StringLines( _
                "Returns a copy of the provided Source Table with the <b>string</b> values within the provided ColumnIndexes being replaced by it's corresponding hex string.", _
                "", _
                "@param [ColumnIndexes] - If specified, an array of column indexes to include in the range, otherwise, all columns are included.")
    
    AddTask "HexToStringValues", "[Source]: Table, [ColumnIndexes]: Array()", _
            Description:=StringLines( _
                "Returns a copy of the provided Source Table with the <b>hex</b> values within the provided ColumnIndexes being replaced by it's corresponding <b>ascii</b> string.", _
                "", _
                "@param [ColumnIndexes] - If specified, an array of column indexes to include in the range, otherwise, all columns are included.")
            
    ' TODO: AddTask "ReplaceColumnNames", "[Source]: Table, Replacer: Callable|Dictionary"
    
    AddTask "RemoveColumns", "[Source]: Table, ColumnIndexes: Array()"
    
    AddTask "MoveColumnsToTheEnd", "[Source]: Table, ColumnIndexes: Array()"
    
    AddTask "MoveColumnsToTheStart", "[Source]: Table, ColumnIndexes: Array()"
    
    AddTask "MoveColumn", "[Source]: Table, ColumnIndex: Long, [MoveCount = 1]: Long"

    ' --- DATA SOURCE TASKS ---
    
    AddTask "DataSource", "", _
            Description:="Provides an empty data source for programmatic use in code.", _
            TaskType:="DS", Requires:="*"
    
    AddTask "WorksheetSource", "", _
            Description:="Provides an empty worksheet for programmatic use in code.", _
            TaskType:="XS", Requires:="*"
    
    AddTask "LoadFromRange", "[Source]: Worksheet, [RangeAddress = .UsedRange()]: String|Excel.Range, [FirstRowHasHeaders = True]: Boolean", TaskType:="DS", Requires:="X*"
    
    AddTask "LoadFromExcelFile", "TargetFile: String, [WorksheetNameOrIndex]: String|Long, [FirstRowHasHeaders = True]: Boolean, " & _
                                 "[Options]: String|Dictionary, [DisableCache = False]: Boolean", TaskType:="DS", Requires:="*"
                                 
    AddTask "LoadFromFileAsJSON", "TargetFile: String, [DisableCache = False]: Boolean", TaskType:="DS", Requires:="*"

    AddTask "LoadFromFileAsCSV", "TargetFile: String, [Options]: String|Dictionary, [DisableCache = False]: Boolean", TaskType:="DS", Requires:="*"
    
    AddTask "LoadFromSQL", "QueryString: String, [ConnectionString]: String, [DisableCache = False]: Boolean", TaskType:="DS", Requires:="*"
    
    ' --- GENERATION TASKS ---
    
    AddTask "SaveToFileAsJSON", "[Source]: Table, TargetFile: String", TaskType:="DSG"
    
    AddTask "SaveToFileAsCSV", "[Source]: Table, TargetFile: String, [Delimiter = ',']: String, [InLocalFormat = False]: Boolean, [OverWriteIfExists = True]: Boolean", TaskType:="DSG"
    
    ' --- XL GENERATION TASKS ---
    
    AddTask "OpenWorksheetFromFile", "TargetFile: String, [Options]: String|Dictionary", TaskType:="XLG", Requires:="*"
    
    AddTask "CreateNewExcelApp", "[WorksheetName = 'Sheet%1']: String", TaskType:="XLG", Requires:="*"
    
    AddTask "CopyToRangeInNewExcelApp", "[Source]: Table, [RangeAddress = 'A1']: String, [WorksheetName = 'Sheet%1']: String", TaskType:="XLG", Requires:="[TD]*"

    AddTask "CopyToRange", "[Source]: Table, Target: Worksheet, [RangeAddress = 'A1']: String, [WriteHeaders = True]: Boolean, [ApplyUserLocale = True]: Boolean", TaskType:="XLG"
    
    AddTask "CopyRange", "[Source]: Worksheet, Target: Worksheet, [SourceRangeAddress = .UsedRange()]: String|Excel.Range, [TargetRangeAddress = 'A1']: String|Excel.Range", TaskType:="XLG"
    
    AddTask "AddWorksheet", "[Source]: Worksheet, [WorksheetName = 'Sheet%1']: String", TaskType:="XLG"
    
    AddTask "GetWorksheet", "[Source]: Worksheet, WorksheetNameOrIndex: String|Long", TaskType:="XLG"
    
    AddTask "SaveWorkbook", "[Source]: Worksheet, TargetFile: String", TaskType:="XLG"
    
    ' --- XL FORMATTING TASKS ---
    
    AddTask "AutoSizeCells", "[Source]: Worksheet, [RangeAddress = .UsedRange()]: String|Excel.Range", TaskType:="XL"
    
    AddTask "FormatAsTable", "[Source]: Worksheet, [TableName = 'xlSheetTable%1']: String, TableStyle: String", TaskType:="XL"
    
    AddTask "FreezeHeaders", "[Source]: Worksheet, [HeaderRows = 1]: Long, [HeaderColumns = 0]: Long", TaskType:="XL"
    
    AddTask "AutoFormatCells", "[Source]: Worksheet, [RangeAddress = .UsedRange()]: String|Excel.Range", TaskType:="XL"

    AddTask "WindowVisibility", "[Source]: Worksheet, [Visible = True]: Boolean", TaskType:="XL"
    
    ' --- DEV / DEBUG ---
    
    AddTask "AsRecordsetTable", "[Source]: Table", _
            Description:="<b><font color=#AA0000>ONLY FOR DEVELOPMENT / DEBUGGING PURPOSES, WILL BE REMOVED SHORTLY.</font></b>", _
            TaskType:="DBG", Requires:="*"
    
    AddTask "CreateMassiveTable", "", _
            Description:="Quickly provides a new Table generated with <b>2.100.000.000</b> rows * <b>1.000</b> columns of sample data.", _
            TaskType:="DBG", Requires:="*"
    
End Sub





' --- UTILITY FUNCTIONS ---

Public Function ActiveTaskId() As String: ActiveTaskId = TaskController().RebuildSequence(TaskController().SequenceIndex())("Id"): End Function

Public Function ColumnCount(ByRef TargetTable As Variant) As Long
    Dim Target As dsTable
    Set Target = TargetTable

    ColumnCount = Target.ColumnCount
End Function

Public Function GetRangeInWorksheet(ByRef SourceSheet As Variant, Optional ByVal RangeAddress As Variant) As Excel.Range
    Dim xlSheet As xlSheetsEx, xlRange As Excel.Range
    Set xlSheet = SourceSheet
    
    If IsMissing(RangeAddress) Then
        Set xlRange = xlSheet.UsedRange
    Else
        Assign RangeAddress, GetParam(RangeAddress, xlSheet.UsedRange)
        If VarType(RangeAddress) = vbString Then
            Set xlRange = xlSheet.Range(CStr(RangeAddress))
        ElseIf IsObject(RangeAddress) Then
            If TypeName(RangeAddress) Like "*Range*" Then
                Set xlRange = RangeAddress
            End If
        End If
    End If
    
    If xlRange Is Nothing Then Set xlRange = xlSheet.UsedRange
    
    Set GetRangeInWorksheet = xlRange
End Function

Private Function AsListOfColumnIndexes(ByRef TargetTable As Variant, Optional ByVal ColumnIndexes As Variant) As ArrayListEx
    Dim Target As dsTable, colIndexes As ArrayListEx, v As Variant
    ColumnIndexes = GetParam(ColumnIndexes, Array())
    Set Target = TargetTable
    Set colIndexes = ArrayListEx.Create()
    
    If Not IsArray(ColumnIndexes) Then
        If CStr(ColumnIndexes) = "" Then
            ColumnIndexes = Array()
        Else
            ColumnIndexes = Array(CLng(ColumnIndexes))
        End If
    End If
    If ArrayListEx.CountElementsIn(ColumnIndexes) = 0 Then
        ColumnIndexes = CollectionsLib.ArrayRange(0, Target.ColumnCount)
    End If
    For Each v In ColumnIndexes
        colIndexes.Add CLng(v)
    Next v
    
    Set AsListOfColumnIndexes = colIndexes
End Function

Private Function GetParam(Optional ByVal Value As Variant, Optional ByVal DefaultValue As Variant) As Variant
    If Not IsMissing(Value) Then
        If Not (IsError(Value) Or IsNull(Value) Or IsEmpty(Value)) Then
            If IsObject(Value) Then
                Set GetParam = Value
            Else
                Select Case VarType(DefaultValue)
                    Case vbBoolean: GetParam = CBool(Value)
                    Case vbLong: GetParam = CLng(Value)
                    Case vbInteger: GetParam = CLng(Value)
                    Case Else: GetParam = Value
                End Select
            End If
            Exit Function
        End If
    End If
    If IsObject(DefaultValue) Then
        Set GetParam = DefaultValue
    Else
        GetParam = DefaultValue
    End If
End Function

Private Function GetParamAsDictionaryEx(Optional ByVal Value As Variant) As DictionaryEx
    If Not IsMissing(Value) Then
        If IsObject(Value) Then
            If TypeName(Value) Like "*Dictionary*" Then
                Set GetParamAsDictionaryEx = DictionaryEx.Create(JSON.Parse(JSON.Stringify(Value), True, True))
            Else
                Err.Raise 13
            End If
        Else
            If IsError(Value) Or IsEmpty(Value) Or IsNull(Value) Then GoTo Fallback
            If Trim(CStr(Value)) = "" Then GoTo Fallback
            Set GetParamAsDictionaryEx = DictionaryEx.Create(JSON.Parse(Value, True, True))
        End If
    Else
        GoTo Fallback
    End If
    
    Exit Function
Fallback:
    Set GetParamAsDictionaryEx = DictionaryEx.Create()
End Function

Private Sub Assign(ByRef Target, Optional ByRef Value As Variant)
    If Not IsMissing(Value) Then
        If IsObject(Value) Then
            Set Target = Value
        Else
            Target = Value
        End If
    End If
End Sub

Private Function AsciiToHex(ByVal Value As String) As String
    Dim i As Long, iMax As Long, t() As String
    iMax = Len(Value)
    ReDim t(1 To iMax)
    
    For i = 1 To iMax
        t(i) = Right("0" & VBA.Hex$(VBA.AscW(VBA.Mid$(Value, i, 1))), 2)
    Next i
    
    AsciiToHex = VBA.Join(t, "")
End Function

Private Function HexToAscii(ByVal Value As String) As String
    Dim i As Long, iMax As Long, t() As String
    iMax = CLng(Len(Value) / 2) - 1
    ReDim t(0 To iMax)
    
    For i = 0 To iMax
        t(i) = VBA.ChrW(CLng("&H" & VBA.Mid$(Value, 1 + (2 * i), 2)))
    Next i
    
    HexToAscii = VBA.Join(t, "")
End Function

Private Function LoadFromCSV(ByRef TargetFile As String, Optional ByVal Delimiter As String = "Auto", _
                                          Optional ByVal InLocalFormat As Boolean = True, Optional ByVal AutoHeaders As Boolean = True, _
                                          Optional ByVal cQuote As String = """") As dsTable
    Dim nCols As Long, c As Long, ContentString As String, leftOvers As Variant
    Dim colHeaders As ArrayListEx, a2X As Array2dEx, dsT As dsTable, sLines As ArrayListEx, cDblQuote As String
    cDblQuote = cQuote & cQuote

    FileSystemLib.TryReadAllTextInFile TargetFile, ContentString
    Set sLines = ArrayListEx.Create()
    If Not CollectionsLib.StringLineSlicesV2(ContentString, sLines, leftOvers, , , cQuote) Then
        sLines.Add leftOvers
    End If
    If Delimiter = "Auto" Then
        Delimiter = AutoDetectDelimiterInCSV(VBA.Mid$(ContentString, sLines(0)(0), sLines(0)(1)))
    End If
    Set colHeaders = ArrayListEx.Create(Split(CollectionsLib.TokenizeArgs(VBA.Mid$(ContentString, sLines(0)(0), sLines(0)(1)), Delimiter), vbNullChar))
    If AutoHeaders Then sLines.RemoveAt 0
    nCols = colHeaders.Count
    
    Set a2X = Array2dEx.Create()
    If sLines.Count > 0 Then
        If sLines(0)(0) > 8 Then
            DelimitedLineSlicesTo2dArrayV2 ContentString, sLines, a2X, nCols, Delimiter, InLocalFormat, cQuote, cDblQuote
        Else
            DelimitedLineSlicesTo2dArray ContentString, sLines, a2X, nCols, Delimiter, InLocalFormat, cQuote, cDblQuote
        End If
    End If
    Set dsT = dsTable.Create(a2X, False)
    If AutoHeaders Then
        For c = 0 To nCols - 1
            If colHeaders(c) = cDblQuote Then
                colHeaders(c) = vbNullString
            Else
                colHeaders(c) = Replace(Replace(Replace(colHeaders(c), cDblQuote, vbNullChar), cQuote, vbNullString), vbNullChar, cQuote)
            End If
        Next c
        dsT.SetHeaders colHeaders.ToArray()
    End If
    Set LoadFromCSV = dsT
End Function

Private Sub DelimitedLineSlicesTo2dArray(ByRef SearchString As String, ByRef LineSlices As ArrayListEx, ByRef a2X As Array2dEx, _
                                         ByVal nCols As Long, Optional ByVal Delimiter As String = ",", Optional ByVal InLocalFormat As Boolean = False, _
                                         Optional ByVal cQuote As String = """", Optional ByVal cDblQuote As String = """""")
    Dim r As Long, i As Long, iDel As Long, iQt As Long, uStr As String, nQt As Long, rMax As Long
    Dim t() As Variant, line As String, c As Long, sLines() As Variant
    rMax = LineSlices.Count - 1
    ReDim t(0 To rMax, 0 To nCols - 1)
    sLines = LineSlices.ToArray()
    
    For r = 0 To rMax
        line = VBA.Mid$(SearchString, sLines(r)(0), sLines(r)(1))
        i = 1
        c = 0
        iDel = InStr(i, line, Delimiter, vbBinaryCompare)
        Do While (iDel > 0)
            uStr = VBA.Mid$(line, i, iDel - i)
            iQt = InStr(1, uStr, cQuote, vbBinaryCompare)
            nQt = Sgn(iQt)
            Do
                Do While (iQt > 0)
                    iQt = InStr(1 + iQt, uStr, cQuote, vbBinaryCompare)
                    nQt = nQt + Sgn(iQt)
                Loop
                If (nQt Mod 2) = 1 Then
                    iQt = iDel - i
                    iDel = InStr(1 + i + iQt, line, Delimiter, vbBinaryCompare)
                    uStr = VBA.Mid$(line, i, iDel - i)
                End If
            Loop Until ((nQt Mod 2) = 0)
            If VBA.Mid$(uStr, 1, 1) = cQuote Then
                t(r, c) = VBA.Replace(VBA.Mid$(uStr, 2, Len(uStr) - 2), cDblQuote, cQuote, , , vbBinaryCompare)
            Else
                If IsNumeric(uStr) Then
                    If InLocalFormat Then
                        t(r, c) = CDbl(CCur(uStr))
                    Else
                        t(r, c) = Val(uStr)
                    End If
                ElseIf uStr = vbNullString Then
                    t(r, c) = Empty
                ElseIf IsDate(uStr) Then
                    t(r, c) = CDate(uStr)
                Else
                    t(r, c) = uStr
                End If
            End If
            i = iDel + 1
            c = c + 1
            iDel = InStr(i, line, Delimiter, vbBinaryCompare)
            If iDel = 0 Then
                If i <= Len(line) Then iDel = Len(line) + 1
            End If
        Loop
    Next r
    a2X.SetInstanceByRef t
End Sub

Private Sub DelimitedLineSlicesTo2dArrayV2(ByRef SearchString As String, ByRef LineSlices As ArrayListEx, ByRef a2X As Array2dEx, _
                                         ByVal nCols As Long, Optional ByVal Delimiter As String = ",", Optional ByVal InLocalFormat As Boolean = False, _
                                         Optional ByVal cQuote As String = """", Optional ByVal cDblQuote As String = """""")
    Dim r As Long, i As Long, iDel As Long, iQt As Long, uStr As String, nQt As Long
    Dim t() As Variant, line As String, c As Long, sPtr As LongPtr, rMax As Long
    Dim e As Long, wStr As String, sLines() As Variant, linePtr As LongPtr
    
    rMax = LineSlices.Count - 1
    ReDim t(0 To rMax, 0 To nCols - 1)
    sLines = LineSlices.ToArray()
    sPtr = StrPtr(SearchString)
    
    For r = 0 To rMax
        linePtr = sPtr + ((sLines(r)(0) - 1) * INT_SIZE) - 4
        MemoryLib.MemLong(linePtr - 4) = (sLines(r)(1) * 2) + 4
        MemoryLib.MemLongPtr(VarPtr(line)) = linePtr
        i = 3
        c = 0
        iDel = InStr(i, line, Delimiter, vbBinaryCompare)
        Do While (iDel > 0)
            MemoryLib.MemLong(linePtr + ((i - 1) * INT_SIZE) - 4) = (iDel - i) * 2
            MemoryLib.MemLongPtr(VarPtr(uStr)) = linePtr + ((i - 1) * INT_SIZE)
            iQt = InStr(1, uStr, cQuote, vbBinaryCompare)
            nQt = Sgn(iQt)
            Do
                Do While (iQt > 0)
                    iQt = InStr(1 + iQt, uStr, cQuote, vbBinaryCompare)
                    nQt = nQt + Sgn(iQt)
                Loop
                If (nQt Mod 2) = 1 Then
                    iQt = iDel - i
                    iDel = InStr(1 + i + iQt, line, Delimiter, vbBinaryCompare)
                    MemoryLib.MemLong(linePtr + ((i - 1) * INT_SIZE) - 4) = (iDel - i) * 2
                End If
            Loop Until ((nQt Mod 2) = 0)
            If VBA.Mid$(uStr, 1, 1) = cQuote Then
                t(r, c) = VBA.Replace(VBA.Mid$(uStr, 2, Len(uStr) - 2), cDblQuote, cQuote, , , vbBinaryCompare)
            Else
                wStr = CStr(uStr)
                If IsNumeric(wStr) Then
                    If InLocalFormat Then
                        t(r, c) = CDbl(CCur(wStr))
                    Else
                        t(r, c) = Val(wStr)
                    End If
                ElseIf wStr = vbNullString Then
                    t(r, c) = Empty
                ElseIf IsDate(wStr) Then
                    t(r, c) = CDate(wStr)
                Else
                    t(r, c) = wStr
                End If
            End If
            i = iDel + 1
            c = c + 1
            iDel = InStr(i, line, Delimiter, vbBinaryCompare)
            If iDel = 0 Then
                If i <= Len(line) Then iDel = Len(line) + 1
            End If
        Loop
    Next r
    MemoryLib.MemLongPtr(VarPtr(uStr)) = CLngPtr(0)
    MemoryLib.MemLongPtr(VarPtr(line)) = CLngPtr(0)
    a2X.SetInstanceByRef t
End Sub

Private Function MakeJagged2dArrayContainer(ByVal Rows As Long, ByVal Columns As Long) As Variant
    Dim t() As Variant, r() As Variant, i As Long
    ReDim r(0 To Rows - 1, 0 To 0)
    ReDim t(0 To Columns - 1)
    For i = 0 To Columns - 1
        t(i) = r
    Next i
    MakeJagged2dArrayContainer = t
End Function

Private Function AutoDetectDelimiterInCSV(ByRef ContentString As String) As String
    AutoDetectDelimiterInCSV = IIf(UBound(VBA.Split(ContentString, ",")) > UBound(VBA.Split(ContentString, ";")), ",", ";")
End Function

Private Function C2ArraySlice(ByRef Value As Variant) As ArraySlice: Set C2ArraySlice = Value: End Function


