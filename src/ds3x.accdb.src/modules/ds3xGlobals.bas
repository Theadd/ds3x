Attribute VB_Name = "ds3xGlobals"
Option Compare Database
Option Explicit
Option Base 0


' --- ACCESS WINDOW HIDE / SHOW ---

Global Const SW_HIDE = 0
Global Const SW_SHOWNORMAL = 1
Global Const SW_SHOWMINIMIZED = 2
Global Const SW_SHOWMAXIMIZED = 3
Global Const SW_SHOW = 5
'Forces a top-level window onto the taskbar when the window is visible.
Public Const WS_EX_APPWINDOW As Long = &H40000

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type BOUNDS
    X As Long
    Y As Long
    W As Long
    h As Long
End Type

' EDIT: CopyMemory
Public Type REMOTE_MEMORY
    memValue As Variant
    remoteVT As Variant 'Will be linked to the first 2 bytes of 'memValue' - see 'InitRemoteMemory'
    isInitialized As Boolean 'In case state is lost
End Type

Public Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Public Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    #If Win64 Then
        dummyPadding As Long
        pvData As LongLong
    #Else
        pvData As Long
    #End If
    rgsabound0 As SAFEARRAYBOUND
End Type


' --- Automation ---

Public Function RunApplicationCommandArgs()
    #If AutomationSupport = 1 Then
        dsApp.ExecuteAutomationCommandArgs VBA.Command$()
    #End If
End Function

#If AutomationSupport = 1 Then

    Public Function IsTaskRunning(Optional ByVal TaskNamePattern As String = "*") As Boolean
        If IsEmpty(dsApp.ActiveTask) Then Exit Function
        IsTaskRunning = ((dsApp.ActiveTask(0) Like "*" & TaskNamePattern & "*") Or (dsApp.ActiveTask(1) Like "*" & TaskNamePattern & "*"))
    End Function
    
    Public Function HasFailedToRunAllTasks() As Boolean
        HasFailedToRunAllTasks = Not (IsEmpty(dsApp.FailedTask))
    End Function
    
    Public Sub SetCustomVar(ByVal VarName As String, ByVal VarValue As Variant)
        dsApp.CustomVar(VarName) = VarValue
    End Sub
    
    ' Adds a preset file to the runnable tasks queue
    Public Sub AddRunnableTask(ByVal TargetPath As String, Optional ByVal RunnableTaskName As String = "", Optional ByVal OnErrorResumeNext As Boolean = False)
        dsApp.RunnableTasks.Add Array(TargetPath, RunnableTaskName, OnErrorResumeNext)
    End Sub
    
    Public Sub ClearAllRunnableTasks()
        dsApp.RunnableTasks.Clear
    End Sub
    
    ' Sequentially executes all runnable tasks in queue
    Public Sub RunAllAsync()
        DoCmd.OpenForm "DS_ASYNC_RUNNER", WindowMode:=acHidden
        Forms("DS_ASYNC_RUNNER").RunAsync
    End Sub
    
    Public Function NumTasksInQueue() As Long
        NumTasksInQueue = dsApp.RunnableTasks.Count
    End Function
    
#End If


' --- Utility Functions ---

' USAGE: Printf("Name: %1, Age: %2", "John", 32) -> "Name: John, Age: 32"
Public Function Printf(ByVal mask As String, ParamArray Tokens() As Variant) As String
    Dim parts() As String: parts = Split(mask, "%")
    Dim i As Long, j As Long, isFound As Boolean, s As String
    
    For i = LBound(parts) + 1 To UBound(parts)
        If LenB(parts(i)) = 0 Then
            parts(i) = "%"
        Else
            isFound = False
            For j = UBound(Tokens) To LBound(Tokens) Step -1
                s = CStr(j + 1)
                If Left$(parts(i), Len(s)) = s Then
                    parts(i) = Tokens(j) & Right$(parts(i), Len(parts(i)) - Len(s))
                    isFound = True
                    Exit For
                End If
            Next j
            If Not isFound Then
                parts(i) = "%" & parts(i)
            End If
        End If
    Next i
    
    Printf = Join(parts, vbNullString)
End Function

'Public Function SecondsToHMS(ByVal Value As Long) As String
'    Dim hrs As Long, mins, secs, m As Integer, t As String
'    On Error GoTo Finally
'
'    hrs = Fix(Value / 3600)
'    mins = Fix(Value / 60) Mod 60
'    secs = Fix((Value Mod 60) / 1)
'
'    If hrs >= 2 Or (hrs = 1 And mins > 39) Then
'        t = t & " " & hrs & "h"
'        If mins >= 5 Then t = t & " " & mins & "m"
'    Else
'        If (hrs = 1 And mins <= 39) Then mins = mins + 60
'        If mins >= 1 Then
'            t = t & " " & mins & "m"
'            If mins < 15 Then t = t & " " & secs & "s"
'        Else
'            t = t & " " & secs & "s"
'        End If
'    End If
'
'    SecondsToHMS = VBA.Mid(t, 2)
'Finally:
'End Function

'' USAGES: XDaysAgo(15), XDaysAgo("7 days..."), XDaysAgo("-7 in a week")
'Public Function XDaysAgo(ByVal Value As Variant, Optional ByVal DateFormat As String = "dd MMM") As String
'    If VBA.InStr(1, Value, CStr(Val(Value))) > 0 Then
'        Value = UCase(VBA.Format$(DateAdd("d", 0 - Int(Val(Value)), Date), DateFormat, vbMonday))
'    End If
'    XDaysAgo = CStr(Value)
'End Function

''*******************************************************************************
''Alternative for CopyMemory - not affected by API speed issues on Windows
''--------------------------
''Mac - wrapper around CopyMemory/memmove
''Win - bytesCount 1 to 2147483647 - no API calls. Uses a combination of
''      REMOTE_MEMORY/SAFEARRAY_1D structs as well as native Strings and Arrays
''      to manipulate memory. Works within size limitation of Strings in VBA
''      For some smaller sizes (<=5) optimizes via MemLong, MemInt, MemByte etc.
''    - bytesCount < 0 or > 2147483647 - wrapper around CopyMemory/RtlMoveMemory
''*******************************************************************************
'Public Sub MemCopy(ByVal destinationPtr As LongPtr _
'                 , ByVal sourcePtr As LongPtr _
'                 , ByVal bytesCount As LongPtr)

'Public Sub TestA2dCopy()
'#If Win64 Then
'    Const PTR_SIZE As Long = 8
'    Const VARIANT_SIZE As Long = 24
'#Else
'    Const PTR_SIZE As Long = 4
'    Const VARIANT_SIZE As Long = 16
'#End If
'    Dim aX As Array2dEx, bX As Array2dEx, srcX As ArrayListEx, t() As Variant, nX As ArrayListEx, bbX As ArrayListEx
'    Dim oCol As New Collection, oDic As New Scripting.Dictionary
'
'    oCol.Add 15
'    oCol.Add 72
'    oCol.Add 33
'
'    oDic("First") = 15
'    oDic("Second") = 72
'    oDic("Third") = 33
'
''    Set nX = ArrayListEx.Create(Array(10, 20, 30))
'''            Array(-1, 310, nX, 131), _
'''            Array(2, 320, ArrayListEx.Create(Array("ArrayList", nX, "SomeString", "You there!", "SomeNumber", -44)), 132), _
'''            Array(-1, 310, "Hello World!ONE", 131), _
'''            Array(2, 320, "Hello World!TWO", 132), _
'''''            Array(-1, 310, oCol, 131), _
'''''            Array(2, 320, oDic, 132), _
'''
''
''    Set bbX = ArrayListEx.Create(Array("ArrayList", nX, "SomeString", "You there!", "SomeNumber", -44))
'''            Array(0, 300, "Hello World!ZERO", 130), _
'''            Array(-1, 310, nX, 131), _
'''            Array(2, 320, bbX, 132), _
'''
'    Set srcX = ArrayListEx.Create(Array( _
'            Array(-1, 310, "Hello World!ONE", 131), _
'            Array(2, 320, "Hello World!TWO", 132), _
'            Array(-3, 330, "Hello World!THREE", 133), _
'            Array(4, 340, "Hello World!FOUR", 134), _
'            Array(-5, 350, "Hello World!FIVE", 135) _
'        ))
'    Set aX = Array2dEx.Create(srcX)
'
''    ReDim t(0 To 5, 0 To 3)
''    Stop
'    'MemoryLib.MemCopy VBA.VarPtr(t(0, 0)), VBA.VarPtr(aX.Instance(0, 0)), VARIANT_SIZE * (6 * 4)
''    CopyMemory ByVal VBA.VarPtr(t(0, 0)), ByVal VBA.VarPtr(aX.Instance(0, 0)), VARIANT_SIZE * (6 * 4)
''    CopyMemory ByVal VBA.VarPtr(t(0, 0)), ByVal VBA.VarPtr(aX.Instance(1, 0)), VARIANT_SIZE * 4
''    CopyMemory ByVal VBA.VarPtr(t(0, 1)), ByVal VBA.VarPtr(aX.Instance(1, 1)), VARIANT_SIZE * 4
''    CopyMemory ByVal VBA.VarPtr(t(0, 2)), ByVal VBA.VarPtr(aX.Instance(1, 2)), VARIANT_SIZE * 4
''    Stop
'    Set bX = aX.GetRange(1, 3, Array(0, 2, 3))
'    Debug.Print JSON.Stringify(bX)
'    'Debug.Print "Columns: " & CStr(bX.ColumnCount)
''    Stop
'End Sub

''*******************************************************************************
''Copy a param array to another array of Variants while preserving ByRef elements
''*******************************************************************************
'Public Sub CloneParamArray(ByRef firstElem As Variant _
'                         , ByVal elemCount As Long _
'                         , ByRef outArray() As Variant)
'    ReDim outArray(0 To elemCount - 1)
'    MemCopy VarPtr(outArray(0)), VarPtr(firstElem), VARIANT_SIZE * elemCount
'    '
'    Static sArr As SAFEARRAY_1D 'Fake array of VarTypes (Integers)
'    Static rmArr As REMOTE_MEMORY
'    '
'    If Not rmArr.isInitialized Then
'        With sArr
'            .cDims = 1
'            .fFeatures = FADF_HAVEVARTYPE
'            .cbElements = INT_SIZE
'        End With
'        InitRemoteMemory rmArr
'        rmArr.memValue = VarPtr(sArr)
'    End If
'    sArr.rgsabound0.cElements = elemCount * VT_SPACING
'    sArr.pvData = VarPtr(outArray(0))
'    '
'    FixByValElements outArray, rmArr, rmArr.remoteVT
'End Sub
'
''*******************************************************************************
''Utility for 'CloneParamArray' - avoid deallocation on elements passed ByVal
''e.g. if original ParamArray has a pointer to a BSTR then safely clear the copy
''*******************************************************************************
'Private Sub FixByValElements(ByRef arr() As Variant _
'                           , ByRef rmArr As REMOTE_MEMORY _
'                           , ByRef vtArr As Variant)
'    Dim i As Long
'    Dim v As Variant
'    Dim vtIndex As Long: vtIndex = 0
'    Dim vt As VbVarType
'    '
'    vtArr = vbArray + vbInteger
'    For i = 0 To UBound(arr)
'        vt = rmArr.memValue(vtIndex)
'        If (vt And VT_BYREF) = 0 Then
'            If (vt And vbArray) = vbArray Or vt = vbObject Or vt = vbString _
'            Or vt = vbDataObject Or vt = vbUserDefinedType Then
'                If vt = vbObject Then Set v = arr(i) Else v = arr(i)
'                rmArr.memValue(vtIndex) = vbEmpty 'Avoid deallocation
'                If vt = vbObject Then Set arr(i) = v Else arr(i) = v
'            End If
'        End If
'        vtIndex = vtIndex + VT_SPACING
'    Next i
'    vtArr = vbEmpty
'End Sub

'' USAGE: CollectionsLib.Tokenize("--task=..\Tasks\Entry.json --name=""Initial task"" --exec")   => ['--task=..\Tasks\Entry.json', '--name="Initial task"', '--exec']
'Public Function G3Tokenize(ByVal SearchString As String, Optional ByVal Tokenizer As String = " ") As Variant
'    G3Tokenize = Split(G3TokenizeArgs(SearchString, Tokenizer), VBA.Chr$(0))
'End Function
'
'Public Function G3ParseToken(ByVal Target As String, Optional ByVal Splitter As String = "=") As Variant
'    G3ParseToken = G3ParseTokenizedArg(Target, Splitter)
'End Function
'
'
'' USAGE: sArgv() = Split(G3TokenizeArgs("one ""two twoB twoC"" three ""four fourB"" five"), Chr$(0))
'Public Function G3TokenizeArgs(ByVal SearchString As String, Optional ByVal Tokenizer As String = " ") As String
'   Dim sArgs As String, sChar As String, nCount As Long, bQuotes As Boolean
'
'   For nCount = 1 To Len(SearchString)
'      sChar = Mid$(SearchString, nCount, 1)
'      If sChar = Chr$(34) Then
'         bQuotes = Not bQuotes
'      End If
'      If sChar = Tokenizer Then
'         If bQuotes Then
'            sArgs = sArgs & sChar
'         Else
'            sArgs = sArgs & Chr$(0)
'         End If
'      Else
'         sArgs = sArgs & sChar
'      End If
'   Next
'   G3TokenizeArgs = sArgs
'End Function
'
'Public Function G3ParseTokenizedArg(ByVal Target As String, Optional ByVal Splitter As String = "=") As Variant
'    Dim t(0 To 1) As Variant, r As Variant
'
'    r = VBA.Split(Target, Splitter, 2)
'    t(0) = r(0)
'    If UBound(r) = 1 Then
'        If (Left(r(1), 1) = """" Or Left(r(1), 1) = "'") And (Left(r(1), 1) = Right(r(1), 1)) Then
'            t(1) = VBA.Mid$(r(1), 2, Len(r(1)) - 2)
'        Else
'            If r(1) = "true" Or r(1) = "false" Then
'                t(1) = CBool(r(1))
'            Else
'                t(1) = r(1)
'            End If
'        End If
'    Else
'        t(1) = True
'    End If
'
'    G3ParseTokenizedArg = t
'End Function
'
'Public Sub TestLoadFromFileAsCSV()
'    Dim t0 As Variant, t2 As Variant, sContent As String, r As Long, sCell As String, t() As Variant, nRows As Long, nCols As Long, c As Long, rStart As Long
'    Dim cNull As String, rTSplitter As String, vDelimiter As String, InLocalFormat As Boolean, AutoHeaders As Boolean, colHeaders As ArrayListEx
'    cNull = VBA.Chr$(0): rTSplitter = vbCr & cNull: vDelimiter = ";": InLocalFormat = True: AutoHeaders = True
'
'    sContent = FileSystemLib.ReadAllTextInFile("C:\dev\samples\TEST_LoadFromFileAsCSV_v2.csv", False)
'    t0 = Split(G3TokenizeArgs(sContent, vbLf), rTSplitter)
'    Set colHeaders = ArrayListEx.Create(Split(G3TokenizeArgs(t0(0), vDelimiter), cNull))
'    nRows = IIf(t0(UBound(t0)) = vbNullString, UBound(t0) - 1, UBound(t0)) + 1  ' TODO: AutoHeaders
'    nCols = colHeaders.Count
'    rStart = IIf(AutoHeaders, 1, 0)
'    ReDim t(0 To nRows - (1 + rStart), 0 To nCols - 1)
'
'    For r = 0 To nRows - (1 + rStart)
'        t2 = Split(G3TokenizeArgs(t0(r + rStart), vDelimiter), cNull)
'        For c = 0 To nCols - 1
'            sCell = Replace(Replace(Replace(t2(c), """""", cNull), """", ""), cNull, """")
'            If Len(sCell) <> Len(t2(c)) Then
'                t(r, c) = sCell
'            Else
'                Select Case True
'                    Case sCell = vbNullString: t(r, c) = Empty
'                    Case IsNumeric(sCell): t(r, c) = IIf(InLocalFormat, CDbl(CCur(sCell)), CDbl(Val(sCell)))
'                    Case IsDate(sCell): t(r, c) = CDate(sCell)
'                    Case Else: t(r, c) = sCell
'                End Select
'            End If
'        Next c
'    Next r
'
'    Dim a2X As Array2dEx
'    Dim dsT As dsTable
'
'    Set a2X = Array2dEx.Create()
'    a2X.Instance = CollectionsLib.GetArrayByRef(t)
'
'    Set dsT = dsTable.Create(a2X, False)
'
'    If AutoHeaders Then
'        For c = 0 To nCols - 1
'            colHeaders(c) = Replace(Replace(Replace(colHeaders(c), """""", cNull), """", ""), cNull, """")
'        Next c
'        dsT.SetHeaders colHeaders.ToArray()
'    End If
'
'    Debug.Print dsT.ToJSON()
'End Sub
'

'Public Sub TestStringLineSlicer()
'    Dim sContent As String, aX As ArrayListEx, i As Long, v As Variant, bX As ArrayListEx
'
'    sContent = VBA.Join(Array("First line", "Second one", "Third" & vbLf & "Another one", "... And last one!"), vbNewLine)
'    Set aX = StringLineSlices(sContent)
'
'    Debug.Print JSON.Stringify(aX)
'    Set bX = ArrayListEx.Create()
'
'    For Each v In aX
'        bX.Add VBA.Mid$(sContent, v(0), v(1))
'    Next v
'
'    Debug.Print JSON.Stringify(bX, 2)
'End Sub
'
'Public Sub SecondTestStringLineSlicer()
'    Dim sContent As String, aX As ArrayListEx
'
'    Stop
'    sContent = FileSystemLib.ReadAllTextInFile("C:\dev\samples\PDM_INSYSAN_v3.csv", False)
'    Stop
'    Set aX = StringLineSlices(sContent)
'    Stop
'    Debug.Print "LINES: " & CStr(aX.Count)
'End Sub

Public Sub TestArraySlicingStuff()
    Dim t() As Variant, sgX As ArraySliceGroup, a2dX As Array2dEx, sg2 As ArraySliceGroup, sl3x As ArraySlice, vReverse As Variant, b2dX As Array2dEx, c2dX As Array2dEx
    ReDim t(0 To 5, 0 To 3)
    
    t(0, 0) = "not zero":    t(0, 1) = 300:  t(0, 2) = "ZERO":   t(0, 3) = 130
    t(1, 0) = "-1":   t(1, 1) = 310:  t(1, 2) = "ONE":    t(1, 3) = 131
    t(2, 0) = "2ac":    t(2, 1) = 320:  t(2, 2) = "TWO":    t(2, 3) = 132
    t(3, 0) = "f -3a":   t(3, 1) = 330:  t(3, 2) = "THREE":  t(3, 3) = 133
    t(4, 0) = 4:    t(4, 1) = 340:  t(4, 2) = "FOUR":   t(4, 3) = 134
    t(5, 0) = "It's -5":   t(5, 1) = 350:  t(5, 2) = "FIVE":   t(5, 3) = 135

    Set a2dX = Array2dEx.Create()
    'a2dX.Instance = CollectionsLib.GetArrayByRef(t)
    a2dX.Instance = t
    
    Set sgX = ArraySliceGroup.Create(a2dX)
    Set sg2 = sgX.GetRange(1, 3, Array(0, 1, 2))
    
    Set sl3x = sg2.SliceAt(0)
    
    'Stop
    
    Set c2dX = New Array2dEx
    
    'c2dX.Instance = CollectionsLib.GetArrayByRef(sl3x.Instance)
    c2dX.Instance = sl3x.Instance
    'vReverse = sl3x.Instance
    'vReverse = CollectionsLib.GetArrayByRef(sl3x.Instance)
'    Stop
    'Debug.Print JSON.Stringify(a2dX, 2)
    
'    Stop
    
    Debug.Print JSON.Stringify(c2dX)
    
'    Stop
'    Set b2dX = Array2dEx.Create(vReverse)
'    Stop
'    Debug.Print JSON.Stringify(b2dX, 2)
'    Stop
End Sub




