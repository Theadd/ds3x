VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xlSheetsEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Public Instance As Excel.Worksheet

Private pAllSheets As ArrayListEx



' --- PROPERTIES ---

Friend Property Get AllSheets() As ArrayListEx
    If pAllSheets Is Nothing Then Set pAllSheets = ArrayListEx.Create()
    Set AllSheets = pAllSheets
End Property

Public Property Get WindowVisibility() As Boolean
    ' If Instance Is Nothing Then Exit Property
    WindowVisibility = Instance.Application.Visible
End Property

Public Property Let WindowVisibility(ByVal ShouldBeVisible As Boolean)
    ' If Instance Is Nothing Then Exit Property
    SetExcelApplicationVisibility Instance.Application, ShouldBeVisible
    ' If ShouldBeVisible Then Debug.Print "TODO: @SEE Excel.Workbook.HighlightChangesOnScreen"
End Property

Public Property Get Workbook() As Excel.Workbook: Set Workbook = Instance.Application.Workbooks(1): End Property

Public Property Get SheetIndex() As Long
    SheetIndex = -1
    If Instance Is Nothing Then Exit Property
    SheetIndex = Instance.Index
End Property

Public Property Get SheetName() As String
    If Instance Is Nothing Then Exit Property
    SheetName = Instance.Name
End Property

Public Property Let SheetName(ByVal Value As String)
    Instance.Name = Value
End Property

Public Property Get Cells() As Excel.Range
    Set Cells = Instance.Cells
End Property

Public Property Get Columns() As Excel.Range
    Set Columns = Instance.Columns
End Property

Public Property Get Range(ByVal Cell1 As Variant, Optional ByVal Cell2 As Variant) As Excel.Range
    If IsMissing(Cell2) Then
        Set Range = Instance.Range(Cell1)
    Else
        Set Range = Instance.Range(Cell1, Cell2)
    End If
End Property

Public Property Get UsedRange() As Excel.Range
    Set UsedRange = Instance.UsedRange
End Property

Public Property Get Protected() As Boolean
    Protected = Instance.ProtectContents
End Property

Public Property Let Protected(ByVal Value As Boolean)
    Me.Workbook.Protect Structure:=Value, Windows:=Value
    If Not Value Then
        Me.Workbook.UnProtect
        Instance.UnProtect
    Else
        Instance.Protect
    End If
End Property

Public Property Let DefaultSaveFormat(ByVal Value As Excel.XlFileFormat)
    Instance.Application.DefaultSaveFormat = Value
End Property



' --- METHODS ---

Public Function GetSheet(ByVal SheetNameOrIndex As Variant) As xlSheetsEx
    Dim sTarget As Excel.Worksheet
    Set sTarget = Me.Workbook.Sheets(SheetNameOrIndex)
    Set GetSheet = New xlSheetsEx
    AllSheets.Add GetSheet
    GetSheet.Bind Me
    Set GetSheet.Instance = sTarget
End Function

Public Function CreateFrom(ByRef Target As Object) As xlSheetsEx
    Set CreateFrom = New xlSheetsEx
    If TypeOf Target Is Excel.Worksheet Then
        Set CreateFrom.Instance = Target
    ElseIf TypeOf Target Is Excel.Range Then
        Set CreateFrom.Instance = Target.Worksheet
    Else
        Err.Raise 13
    End If

    CreateFrom.AllSheets.Add CreateFrom
End Function

Public Function Create(Optional ByVal WorksheetName As String = "Sheet%1", Optional ByVal TargetFile As Variant, Optional ByVal Options As DictionaryEx = Nothing) As xlSheetsEx
    Dim xlApp As Excel.Application, xWb As Excel.Workbook, xWs As Excel.Worksheet
    Set Create = New xlSheetsEx
    If Options Is Nothing Then Set Options = DictionaryEx.Create()
    
    On Error GoTo Fallback
    Set xlApp = New Excel.Application
    xlApp.DefaultSaveFormat = xlOpenXMLWorkbookMacroEnabled
    SetExcelApplicationVisibility xlApp, False
    
    If Not IsMissing(TargetFile) Then
        If TryWaitFileExists(TargetFile, 0) Then
            Set xWb = LoadWorkbookAt(xlApp, TargetFile, Options)
            Set xWs = xWb.Worksheets(1)
        Else
            ' TODO: Save empty worksheet to file
        End If
    Else
        Set xWb = xlApp.Workbooks.Add(1)
        xWb.PrecisionAsDisplayed = True
        Set xWs = xWb.Sheets(1)
        xWs.Name = Printf(WorksheetName, 1)
    End If
    
    Set Create.Instance = xWs
    Create.AllSheets.Add Create
    
    Exit Function
Fallback:
    On Error Resume Next
    xlApp.ScreenUpdating = True
    xlApp.Workbooks(1).Close SaveChanges:=False
    xlApp.Quit
    Set Instance = Nothing
    Set pAllSheets = Nothing
    Set xlApp = Nothing
    Set Create = Nothing
End Function

Private Function LoadWorkbookAt(ByRef xlApp As Excel.Application, ByVal TargetFile, ByVal Options As DictionaryEx) As Excel.Workbook
    Dim TTextQualifier As Long
    
    If Options.GetValue("NoTextQualifier", True) Then
        TTextQualifier = xlTextQualifierNone
    Else
        TTextQualifier = xlTextQualifierDoubleQuote
    End If
    If Options.GetValue("UTF8", False) Then
        xlApp.Application.Workbooks.OpenText TargetFile, _
            Origin:=65001, _
            TextQualifier:=TTextQualifier, _
            DataType:=xlDelimited, _
            Semicolon:=True, _
            Local:=Options.GetValue("ReadOnly", True)
'        xlApp.Application.Workbooks.OpenText TargetFile, Origin:=65001, DataType:=xlDelimited, Other:=True, OtherChar:=";"
        Set LoadWorkbookAt = xlApp.Application.Workbooks(1)
    Else
        Set LoadWorkbookAt = xlApp.Application.Workbooks.Open(TargetFile, _
            UpdateLinks:=IIf(Options.GetValue("UpdateLinks", False), 3, 0), _
            ReadOnly:=Options.GetValue("ReadOnly", False), _
            Local:=Options.GetValue("Local", True))
    End If
End Function

Public Function AddSheet(Optional ByVal WorksheetName As String = "Sheet%1") As xlSheetsEx
    With Me.Workbook
        .Sheets.Add Type:=xlWorksheet, After:=.Worksheets(.Worksheets.Count)
        Set AddSheet = GetSheet(.Sheets.Count)
        AddSheet.SheetName = Printf(WorksheetName, .Sheets.Count)
    End With
End Function

Public Function SaveWorkbook(ByVal TargetPath As String, Optional ByVal EnableMacros As Boolean = False) As xlSheetsEx
    If EnableMacros Then
        Me.Workbook.SaveAs TargetPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled, Local:=True
    Else
        Me.Workbook.SaveAs TargetPath, FileFormat:=xlOpenXMLWorkbook, Local:=True
    End If
    Set SaveWorkbook = Me
End Function

Friend Function Bind(ByVal Other As xlSheetsEx) As xlSheetsEx
    Unbind
    Set pAllSheets = Other.AllSheets
    Set Instance = Other.Instance
    Set Bind = Me
End Function

Friend Sub Unbind()
    On Error Resume Next
    If pAllSheets Is Nothing Then Exit Sub
    
    AllSheets.Remove Me
    Set Instance = Nothing
    Set pAllSheets = Nothing
End Sub

Public Sub Dispose()
    On Error Resume Next
    Dim shouldContinue As Boolean
    
    If Not Instance Is Nothing Then
        If Not pAllSheets Is Nothing Then
            If pAllSheets.Count > 1 Then
                Unbind
            Else
                shouldContinue = True
            End If
        Else
            shouldContinue = True
        End If
        
        If shouldContinue Then
            If WindowVisibility Then
                ReleaseWorkbook
            Else
                DiscardChangesAndExit
            End If
        End If
    End If
End Sub

Public Sub ReleaseWorkbook()
    Me.WindowVisibility = True
    Me.Protected = False
    DoEvents
    
    UnbindAllSheets
End Sub

Public Sub DiscardChangesAndExit()
    Dim xlApp As Excel.Application
    On Error Resume Next
    
    If Not Instance Is Nothing Then
        Set xlApp = Instance.Application
        Me.Workbook.Close SaveChanges:=False
        xlApp.ScreenUpdating = True
        UnbindAllSheets
        xlApp.Quit
        Set xlApp = Nothing
        Set Instance = Nothing
    End If
    Set pAllSheets = Nothing
End Sub




' --- CELL FORMATTING AND FIXING UTILITIES ---

Public Function AutoSizeCells(ByVal TargetRange As Excel.Range) As xlSheetsEx
    Dim c As Long
    
    With TargetRange
        .Cells.EntireColumn.AutoFit
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        
        For c = 1 To .Columns.Count
            Select Case .Columns(c).ColumnWidth
                Case Is >= 80: .Columns(c).ColumnWidth = 75
                Case Is < 15: .Columns(c).ColumnWidth = 15
                Case Is >= 60: .Columns(c).ColumnWidth = 50
            End Select
        Next c
    End With
    Set AutoSizeCells = Me
End Function

Public Function FormatAsTable(Optional ByVal TableName As String = "xlSheetTable%1", Optional ByVal TableStyle As String = "TableStyleMedium6") As xlSheetsEx
    TableName = Printf(TableName, Me.SheetIndex, Me.SheetName)
    Instance.ListObjects.Add(1, Instance.UsedRange, , 1).Name = TableName
    Instance.ListObjects(TableName).TableStyle = TableStyle
    Set FormatAsTable = Me
End Function

Public Function FreezeHeaders(Optional ByVal HeaderRows As Long = 1, Optional ByVal HeaderColumns As Long = 0) As xlSheetsEx
    Instance.Application.ScreenUpdating = True

    With Instance.Application.ActiveWindow
        .SplitColumn = HeaderColumns
        .SplitRow = HeaderRows
        .FreezePanes = True
    End With
    
    Set FreezeHeaders = Me
End Function

Public Function AutoFormatCells(Optional ByVal TargetRange As Excel.Range = Nothing) As xlSheetsEx
    Dim c As Long, res As String
    If TargetRange Is Nothing Then Set TargetRange = Me.UsedRange
    
    For c = 1 To TargetRange.Columns.Count
        res = GetCommonFormatForRange(TargetRange.Columns(c))
        If res <> "General" Then
            TargetRange.Columns(c).NumberFormatLocal = res
        End If
    Next c
    
    Set AutoFormatCells = Me
End Function



' --- PRIVATE ---

Private Function GetCommonFormatForRange(ByVal Rng As Excel.Range) As String
    Dim Item As Variant, i As Long, k As Long, iMax As Long, isDefaultSet As Boolean
    iMax = 10
    GetCommonFormatForRange = "General"
    
    For Each Item In Rng.Cells
        If Item.Text <> "" Then
            
            If Not isDefaultSet Then
                If Item.NumberFormat = "General" Then
                    GetCommonFormatForRange = "General"
                Else
                    GetCommonFormatForRange = Item.NumberFormatLocal
                End If
                isDefaultSet = True
            End If
            
            If Item.Text <> Item.FormulaR1C1Local Then
                If Item.NumberFormat = "General" Then
                    If Item.Text Like "*,*E+*" Then
                        GetCommonFormatForRange = "0"
                        Exit For
                    End If
                Else
                    If Item.NumberFormatLocal Like "[#][,.][#]?0[,.]*" Then
                        GetCommonFormatForRange = "#.##0,00;-#.##0,00"
                    Else
                        GetCommonFormatForRange = Item.NumberFormatLocal
                    End If
                    Exit For
                End If
            End If
            
            i = i + 1
            If i >= iMax Then Exit For
        End If
        
        k = k + 1
        If k >= 300 Then Exit For
    Next Item
    
End Function

Private Sub SetExcelApplicationVisibility(ByVal xlApp As Excel.Application, ByVal ShouldBeVisible As Boolean)
    With xlApp
        .ScreenUpdating = ShouldBeVisible
        .DisplayAlerts = ShouldBeVisible
        .Visible = ShouldBeVisible
    End With
End Sub

Private Sub UnbindAllSheets()
    Dim sItem As xlSheetsEx, i As Long, iStart As Long
    
    On Error GoTo Finally
    iStart = AllSheets.Count - 1
    For i = iStart To 0 Step -1
        On Error Resume Next
        Set sItem = AllSheets(i)
        sItem.Unbind
        Set sItem = Nothing
        On Error GoTo 0
    Next i
Finally:
End Sub

Private Sub Class_Terminate()
    Debug.Print "IN xlSheetsEx.Terminate()"
    If Not Instance Is Nothing Then
        On Error Resume Next
        With Instance.Application
            .ScreenUpdating = True
            .Workbooks(1).Close SaveChanges:=False
            .Quit
        End With
        Set Instance = Nothing
        Set pAllSheets = Nothing
    End If
    'Unbind
End Sub
