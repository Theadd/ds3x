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



Public Function AddRowNumbers(ByRef TargetTable As Variant, Optional ByVal UseAsFirstNumber As Long = 1) As dsTable
    Dim t() As Variant, i As Long, Target As dsTable
    Set Target = TargetTable
    ReDim t(0 To Target.Count - 1, 0 To 0)
    
    For i = 0 To Target.Count - 1
        t(i, 0) = i + UseAsFirstNumber
    Next i
    
    Set AddRowNumbers = dsTable _
                            .Create(Array2dEx.Create(t)) _
                            .SetHeaders(ArrayListEx.Create(Array(Array(dsTable.CreateColumnHeader("#"))))) _
                            .Join(Target)
End Function

Public Function SetColumnName(ByRef TargetTable As Variant, Optional ByVal ColumnIndexes As Variant, Optional ByVal ColumnName As Variant) As dsTable
    Dim t() As Variant, i As Long, cH As Variant, Target As dsTable, d As Scripting.Dictionary, e As Long, colIndexes As ArrayListEx
    ColumnIndexes = GetParam(ColumnIndexes, Array())
    ColumnName = GetParam(ColumnName, "%1")
    
    Set Target = TargetTable
    cH = Target.HeaderList(0)
    Set colIndexes = AsListOfColumnIndexes(Target, ColumnIndexes)
    ReDim t(0 To 0, 0 To Target.ColumnCount - 1)
    
    For i = 0 To Target.ColumnCount - 1
        If colIndexes.Contains(i) Then
            Set d = cH(i)
            Set t(0, i) = DictionaryEx.Bind(d).Clone().Add("ColumnName", dsApp.ApplyCustomVarsOn(Printf(ColumnName, i, e, i + 1, e + 1))).Instance
            e = e + 1
        Else
            Set d = cH(i)
            Set t(0, i) = d
        End If
    Next i
    
    Set SetColumnName = dsTable.Create(Target.Records).SetHeaders(Array2dEx.Create(t))
End Function

Public Function SetNumberFormat(ByRef TargetTable As Variant, Optional ByVal ColumnIndexes As Variant, Optional ByVal NumberFormat As Variant) As dsTable
    Dim t() As Variant, i As Long, cH As Variant, Target As dsTable, d As Scripting.Dictionary, colIndexes As ArrayListEx, e As Long
    ColumnIndexes = GetParam(ColumnIndexes, Array())
    NumberFormat = GetParam(NumberFormat, "General")
    
    Set Target = TargetTable
    cH = Target.HeaderList(0)
    Set colIndexes = AsListOfColumnIndexes(Target, ColumnIndexes)
    ReDim t(0 To 0, 0 To Target.ColumnCount - 1)
    
    For i = 0 To Target.ColumnCount - 1
        If colIndexes.Contains(i) Then
            Set d = cH(i)
            Set t(0, i) = DictionaryEx.Bind(d).Clone().Add("NumberFormat", dsApp.ApplyCustomVarsOn(Printf(NumberFormat, i, e, i + 1, e + 1))).Instance
            e = e + 1
        Else
            Set d = cH(i)
            Set t(0, i) = d
        End If
    Next i
    
    Set SetNumberFormat = dsTable.Create(Target.Records).SetHeaders(Array2dEx.Create(t))
End Function

'RecordsetEx.Filter("REPORT_DATE < '20/04/2023' AND REPORT_STATUS = 'RESUELTA'")
Public Function Filter(ByRef TargetTable As Variant, Optional ByVal QueryFilter As Variant) As dsTable
    Dim rX As RecordsetEx, Target As dsTable
    Set Target = TargetTable
    QueryFilter = GetParam(QueryFilter, "")

    Set rX = RecordsetEx.Create(Target.NamedRecordset).Filter(QueryFilter)
    Set Filter = dsTable.Create(rX).SetHeaders(Target.Headers)
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

' CallbackFunction can also be a global function, for instance: Callback.Create(, "MillisecondsToTime")
Public Function ReplaceValues(ByRef TargetTable As Variant, ByRef Replacer As Variant, Optional ByVal ColumnIndexes As Variant) As dsTable
    Dim t() As Variant, i As Long, iMax As Long, c As Long, cMax As Long, Target As dsTable, cb As Callback, v As Variant
    Dim dRepl As DictionaryEx, colIndexes As New ArrayListEx
    Set Target = TargetTable
    ColumnIndexes = GetParam(ColumnIndexes, Array())
    iMax = Target.Count - 1
    cMax = Target.ColumnCount - 1
    ReDim t(0 To iMax, 0 To cMax)
    If ArrayListEx.CountElementsIn(ColumnIndexes) = 0 Then
        ColumnIndexes = CollectionsLib.ArrayRange(0, cMax)
    End If
    For Each v In ColumnIndexes
        colIndexes.Add CLng(v)
    Next v
    
    If TypeName(Replacer) Like "*Dictionary*" Then
        Set dRepl = DictionaryEx.Create(Replacer)
        With Target.RecordArray
            For c = 0 To cMax
                If colIndexes.Contains(c) Then
                    For i = 0 To iMax
                        v = .Instance(i, c)
                        If Not IsNull(v) Then
                            t(i, c) = dRepl.GetValue(v, v)
                        End If
                    Next i
                Else
                    For i = 0 To iMax
                        t(i, c) = .Instance(i, c)
                    Next i
                End If
            Next c
        End With
    Else
        Set cb = Replacer
        With Target.RecordArray
            For c = 0 To cMax
                If colIndexes.Contains(c) Then
                    For i = 0 To iMax
                        v = .Instance(i, c)
                        If Not IsNull(v) Then
                            t(i, c) = cb.Execute(v).result
                        End If
                    Next i
                Else
                    For i = 0 To iMax
                        t(i, c) = .Instance(i, c)
                    Next i
                End If
            Next c
        End With
    End If
    
    Set ReplaceValues = pTaskController.AddResource(dsTable.Create(Array2dEx.Create(t)).SetHeaders(Target.Headers))
End Function

Public Function ReplaceColumnNames(ByRef TargetTable As Variant, ByRef Replacer As Variant) As dsTable
    Dim t() As Variant, i As Long, cH As Variant, Target As dsTable, d As Scripting.Dictionary, dX As New DictionaryEx
    Dim cb As Callback, dRepl As DictionaryEx
    Set Target = TargetTable
    
    cH = Target.HeaderList(0)
    ReDim t(0 To 0, 0 To Target.ColumnCount - 1)
    
    If TypeName(Replacer) Like "*Dictionary*" Then
        Set dRepl = DictionaryEx.Create(Replacer)
        For i = 0 To Target.ColumnCount - 1
            Set d = cH(i)
            Set dX = DictionaryEx.Bind(d).Clone()
    
            dX.Add "ColumnName", dRepl.GetValue(dX.Item("ColumnName"), dX.Item("ColumnName"))
            Set t(0, i) = dX.Instance
        Next i
    Else
        Set cb = Replacer
        For i = 0 To Target.ColumnCount - 1
            Set d = cH(i)
            Set dX = DictionaryEx.Bind(d).Clone()
    
            dX.Add "ColumnName", cb.Execute(dX.Item("ColumnName")).result
            Set t(0, i) = dX.Instance
        Next i
    End If
    
    Set ReplaceColumnNames = pTaskController.AddResource(dsTable.Create(Target.Records).SetHeaders(Array2dEx.Create(t)))
End Function

Public Function RemoveColumns(ByRef TargetTable As Variant, ByVal ColumnIndexes As Variant) As dsTable
    Dim t() As Variant, i As Long, Target As dsTable, aX As ArrayListEx, c As Variant
    Set Target = TargetTable
    
    Set aX = ArrayListEx.Create(CollectionsLib.ArrayRange(0, ColumnCount(Target) - 1))
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

    Set aX = ArrayListEx.Create(CollectionsLib.ArrayRange(0, ColumnCount(Target) - 1))
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
                .Create(ArrayListEx.Create(xlRange)) _
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
    
    If Not FileSystemLib.TryReadAllTextInFile(sTarget, sContent, False) Then
        Err.Raise 17, , "Unable to read text from provided file"
    End If
    Set dsT = dsTable.Create(sContent)
    If Not DisableCache Then
        Set TaskController().DataSource(ActiveTaskId()) = dsT
    End If

    Set LoadFromFileAsJSON = pTaskController.AddResource(dsT)
End Function

Public Function LoadFromSQL(ByVal QueryString As Variant, Optional ByVal ConnectionString As Variant, Optional ByVal DisableCache As Variant) As dsTable
    Dim dsT As dsTable, dbQ As dbQuery
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
    Set dbQ = New dbQuery
    If Trim(CStr(ConnectionString)) <> "" Then dbQ.ConnectionString = CStr(ConnectionString)
    dbQ.CloseOnTerminate = False
    
    Set dsT = dsTable.Create(dbQ.Create(CStr(QueryString)), True)
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

Public Function CopyToRange(ByRef SourceTable As Variant, ByRef TargetSheet As Variant, Optional ByVal RangeAddress As Variant, Optional ByVal WriteHeaders As Variant) As xlSheetsEx
    Dim Target As dsTable, xlSheet As xlSheetsEx, xlRange As Excel.Range
    Assign RangeAddress, GetParam(RangeAddress, "A1")
    WriteHeaders = GetParam(WriteHeaders, True)
    
    Set Target = SourceTable
    Set xlSheet = TargetSheet
    If IsObject(RangeAddress) Then
        Set xlRange = RangeAddress
    Else
        Set xlRange = xlSheet.Range(RangeAddress)
    End If
    Target.CopyToRange xlRange, True, CBool(WriteHeaders)
    
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

Private Function StringLines(ParamArray Lines() As Variant) As String: StringLines = VBA.Join(Lines, " " & vbNewLine): End Function

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

    AddTask "AddRowNumbers", "[Source]: Table, [FirstNumber = 1]: Long"
    
    AddTask "SetColumnName", "[Source]: Table, [ColumnIndexes]: Long|Array(), [ColumnName]: String", _
            Description:=StringLines( _
                "@param [ColumnIndexes] - A column index position or an array of indexes. When no value is provided, all columns are included.", _
                "@param [ColumnName] - The string to use for renaming that/those column/s. Supports custom vars and %1 to %4, being:", _
                "    %1: The column index.", _
                "    %2: the column index within the ColumnIndexes array.", _
                "    %3 and %4: Position instead of index for %1 and %2 respectively.")
    
    AddTask "SetNumberFormat", "[Source]: Table, [ColumnIndexes]: Long|Array(), [NumberFormat]: String", _
            Description:=StringLines( _
                "Select a single Column before picking this task to see it's current .NumberFormat value.", _
                "", _
                "@param [ColumnIndexes] - A column index position or an array of indexes. When no value is provided, all columns are included.", _
                "@param [NumberFormat] - Excel's .NumberFormat string. Supports custom vars and %1 to %4, being:", _
                "    %1: The column index.", _
                "    %2: the column index within the ColumnIndexes array.", _
                "    %3 and %4: Position instead of index for %1 and %2 respectively.", _
                "", _
                "NOTE: This task will <u>only</u> apply <b>.NumberFormat</b> to corresponding columns in an <b>Excel.Worksheet</b> <u>just before</u> copying the data to the specified <b>Excel.Range</b>.")
                    
    AddTask "Filter", "[Source]: Table, QueryFilter: String", _
            Usage:="{...TaskName: 'Filter', Values: [, ""REPORT_DATE < '20/04/2023' AND REPORT_STATUS = 'RESUELTA'""]}"
    
    AddTask "GetRange", "[Source]: Table, [Index]: Long, [GetCount]: Long, [ColumnIndexes]: Array()", _
            Description:=StringLines( _
                "Returns a subset of rows and/or columns from the Source Table. The ColumnIndexes Array can contain sparse elements.", _
                "", _
                "@param [Index] - The zero-based row index at which the range starts.", _
                "@param [GetCount] - The number of rows in the range.", _
                "@param [ColumnIndexes] - If specified, an array of column indexes to include in the range, otherwise, no columns are excluded.")

    AddTask "Join", "[Source]: Table, Other: Table", _
            Description:="Concatenate all elements of Other Table as additional columns into Source Table."

    AddTask "ReplaceValues", "[Source]: Table, Replacer: Callable|Dictionary, [ColumnIndexes]: Array()", _
            Description:=StringLines( _
                "Returns a copy of the provided Source Table with the values within the provided ColumnIndexes being replaced by the provided Replacer.", _
                "", _
                "@param Replacer - A name to a global function or a Scripting.Dictionary in JSON format.", _
                "@param [ColumnIndexes] - If specified, an array of column indexes to include in the range, otherwise, all columns are included."), _
            Usage:="{...TaskName: 'ReplaceValues', Values: [, 'SecondsToHMS']}"
            
    AddTask "ReplaceColumnNames", "[Source]: Table, Replacer: Callable|Dictionary"
    
    AddTask "RemoveColumns", "[Source]: Table, ColumnIndexes: Array()"
    
    AddTask "MoveColumnsToTheEnd", "[Source]: Table, ColumnIndexes: Array()"
    
    AddTask "MoveColumnsToTheStart", "[Source]: Table, ColumnIndexes: Array()"
    
    AddTask "MoveColumn", "[Source]: Table, ColumnIndex: Long, [MoveCount = 1]: Long"

    ' --- DATA SOURCE TASKS ---
    
    AddTask "DataSource", "", TaskType:="DS", Requires:="*"
    
    AddTask "WorksheetSource", "", TaskType:="XS", Requires:="*"
    
    AddTask "LoadFromRange", "[Source]: Worksheet, [RangeAddress = .UsedRange()]: String|Excel.Range, [FirstRowHasHeaders = True]: Boolean", TaskType:="DS", Requires:="X*"
    
    AddTask "LoadFromExcelFile", "TargetFile: String, [WorksheetNameOrIndex]: String|Long, [FirstRowHasHeaders = True]: Boolean, " & _
                                 "[Options]: String|Dictionary, [DisableCache = False]: Boolean", TaskType:="DS", Requires:="*"
                                 
    AddTask "LoadFromFileAsJSON", "TargetFile: String, [DisableCache = False]: Boolean", TaskType:="DS", Requires:="*"
    
    AddTask "LoadFromSQL", "QueryString: String, [ConnectionString]: String, [DisableCache = False]: Boolean", TaskType:="DS", Requires:="*"
    
    ' --- GENERATION TASKS ---
    
    AddTask "SaveToFileAsJSON", "[Source]: Table, TargetFile: String", TaskType:="DSG"
    
    ' --- XL GENERATION TASKS ---
    
    AddTask "OpenWorksheetFromFile", "TargetFile: String, [Options]: String|Dictionary", TaskType:="XLG", Requires:="*"
    
    AddTask "CreateNewExcelApp", "[WorksheetName = 'Sheet%1']: String", TaskType:="XLG", Requires:="*"
    
    AddTask "CopyToRangeInNewExcelApp", "[Source]: Table, [RangeAddress = 'A1']: String, [WorksheetName = 'Sheet%1']: String", TaskType:="XLG", Requires:="[TD]*"

    AddTask "CopyToRange", "[Source]: Table, Target: Worksheet, [RangeAddress = 'A1']: String, [WriteHeaders = True]: Boolean", TaskType:="XLG"
    
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
        ColumnIndexes = CollectionsLib.ArrayRange(0, Target.ColumnCount - 1)
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

Private Sub Assign(ByRef Target, Optional ByRef Value As Variant)
    If Not IsMissing(Value) Then
        If IsObject(Value) Then
            Set Target = Value
        Else
            Target = Value
        End If
    End If
End Sub


