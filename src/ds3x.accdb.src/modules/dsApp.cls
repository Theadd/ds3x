VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dsApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("ds3x")
Option Compare Database
'@PredeclaredId
Option Explicit
Option Base 0


#If AutomationSupport = 1 Then
    Private pActiveTask As Variant
    Private pFailedTask As Variant
    Private pRunnableTasks As ArrayListEx
    Private pLiveEd As dsLiveEd
    Private pQuitOnRunAll As Boolean
#End If

Private pCustomVars As DictionaryEx


Public Property Get LoadConfig(Optional ByVal ConfigFile As String = "package.json") As DictionaryEx: Set LoadConfig = GetConfigFromFile(ConfigFile): End Property
Public Property Let Log(ByVal LogLevel As String, ByVal LogMessage As String): AppendToLogFile LogLevel, LogMessage: End Property

Public Property Get CustomVars() As DictionaryEx
    If pCustomVars Is Nothing Then InitializeCustomVars
    Set CustomVars = pCustomVars
End Property

Public Property Get CustomVar(ByVal VarName As String) As Variant
    If pCustomVars Is Nothing Then InitializeCustomVars
    If Not (Left(VarName, 2) = "${") Then VarName = "${" & VarName & "}"
    CustomVar = pCustomVars.GetValue(VarName, vbNullString)
End Property

Public Property Let CustomVar(ByVal VarName As String, ByVal VarValue As Variant)
    If pCustomVars Is Nothing Then InitializeCustomVars
    If Not (Left(VarName, 2) = "${") Then VarName = "${" & VarName & "}"
    pCustomVars(VarName) = VarValue
End Property

Public Property Get ApplyCustomVarsOn(ByVal Target As Variant) As Variant
    If pCustomVars Is Nothing Then InitializeCustomVars
    If pCustomVars.Count = 0 Then GoTo NoCustomVars
    
    If IsObject(Target) Then
        Set ApplyCustomVarsOn = ReplaceCustomVars(Target)
    Else
        ApplyCustomVarsOn = ReplaceCustomVars(Target)
    End If
    
    Exit Property
NoCustomVars:
    If IsObject(Target) Then
        Set ApplyCustomVarsOn = Target
    Else
        ApplyCustomVarsOn = Target
    End If
End Property

#If AutomationSupport = 1 Then
    
    Public Property Get ActiveTask() As Variant: ActiveTask = pActiveTask: End Property
    Public Property Get FailedTask() As Variant: FailedTask = pFailedTask: End Property

    Public Property Get RunnableTasks() As ArrayListEx
        If pRunnableTasks Is Nothing Then Set pRunnableTasks = ArrayListEx.Create()
        Set RunnableTasks = pRunnableTasks
    End Property
    
    Public Property Get RunAll() As Boolean
        Dim t As Variant, shouldContinue As Boolean, g As Variant, sTarget As String
        On Error GoTo ErrorHandler
        pFailedTask = Empty
        
        For Each t In RunnableTasks.ToArray()
            pActiveTask = t
            Set pLiveEd = Nothing
            
            Set pLiveEd = New dsLiveEd
            sTarget = FileSystemLib.Resolve(ApplyCustomVarsOn(CStr(t(0))))
            If CStr(t(1)) = "" Then
                CustomVar("TaskName") = FileSystemLib.FSO.GetBaseName(sTarget)
            Else
                CustomVar("TaskName") = ApplyCustomVarsOn(CStr(t(1)))
            End If

            shouldContinue = pLiveEd.ImportPreset(sTarget)
            If shouldContinue Then
                shouldContinue = pLiveEd.TryApply(g)
                If Not shouldContinue Then
                    If CBool(t(2)) Then
                        shouldContinue = True   ' Where `t(2)` is the `OnErrorResumeNext` (optional) parameter provided to the `AddRunnableTask()` call.
                        dsApp.Log("warning") = "Resuming execution after an error occurred while running the preset at " & sTarget
                    End If
                End If
            Else
                dsApp.Log("error") = "Failed to import preset at " & sTarget
            End If
            CustomVar("TaskName") = ""
            
            If shouldContinue Then
                pActiveTask = Empty
                RunnableTasks.RemoveAt 0
            Else
                pFailedTask = pActiveTask
                pActiveTask = Empty
                dsApp.Log("error") = "Error while running the preset at " & sTarget
                Exit For
            End If
        Next t
        
Finally:
        Set pLiveEd = Nothing
        RunAll = shouldContinue
        If pQuitOnRunAll Then Application.Quit acQuitPrompt
        Exit Property
ErrorHandler:
        dsApp.Log("critical") = "Unhandled error @dsAppGlobals.RunAll() - " & Err.Description
        GoTo Finally
    End Property
    
    
    Public Sub ExecuteAutomationCommandArgs(ByVal sArgs As String)
        Dim vArgs As Variant, sArg As Variant, vArg As Variant, isAsync As Boolean, isExec As Boolean, isContinue As Boolean
        Dim aX As ArrayListEx, dX As DictionaryEx
        
        sArgs = Trim(sArgs)
        If sArgs = "" Then Exit Sub
        
        Set aX = ArrayListEx.Create()
        Set dX = DictionaryEx.Create()
        vArgs = CollectionsLib.Tokenize(sArgs)
        For Each sArg In vArgs
            vArg = CollectionsLib.ParseToken(CStr(sArg))
            Select Case vArg(0)
                Case "--task"
                    aX.Add Array(vArg(1), "")
                Case "--exec"
                    isExec = CBool(vArg(1))
                Case "--async"
                    isAsync = CBool(vArg(1))
                Case "--continue"
                    isContinue = CBool(vArg(1))
                Case Else
                    If Left(vArg(0), 6) = "--var-" Then
                        dX.Add VBA.Mid$(vArg(0), 7, Len(vArg(0))), vArg(1)
                    ElseIf Left(vArg(0), 7) = "--task-" Then
                        aX.Add Array(vArg(1), VBA.Mid$(vArg(0), 8, Len(vArg(0))))
                    Else
                        MsgBox "Unknown parameter: " & CStr(vArg(0))
                        GoTo HandleUnknownParamter
                    End If
            End Select
        Next sArg
        
        For Each vArg In dX
            CustomVar(vArg(0)) = vArg(1)
        Next vArg
        
        For Each vArg In aX
            AddRunnableTask CStr(vArg(0)), CStr(vArg(1)), isContinue
        Next vArg
        
        If isExec Then
            pQuitOnRunAll = True
            If isAsync Then
                RunAllAsync
            Else
                Debug.Print "[INFO] RunAll = " & CStr(dsApp.RunAll())
            End If
        End If
        
        Exit Sub
HandleUnknownParamter:
        ' Abort
        Application.Quit
    End Sub

#Else

    Public Property Get RunAll() As Boolean
        Err.Raise 425
    End Property

#End If


' --- CustomVars ---

Private Sub InitializeCustomVars()
    Dim s As String, dX As DictionaryEx, Item As Variant
    Set pCustomVars = DictionaryEx.Create()
    
    pCustomVars.Add "${Timestamp}", CStr(DateDiff("s", DateValue("1970-01-01"), Now()))
    pCustomVars.Add "${Date}", VBA.Format$(Date, "yyyymmdd")
    pCustomVars.Add "${DateTime}", VBA.Format$(Now(), "yyyymmddhhMMss")
    pCustomVars.Add "${Year}", VBA.Format$(Date, "yyyy")
    pCustomVars.Add "${Month}", VBA.Format$(Date, "mm")
    pCustomVars.Add "${Day}", VBA.Format$(Date, "dd")
    pCustomVars.Add "${Username}", Nz(CreateObject("WScript.Network").UserName, "")
    pCustomVars.Add "${UserProfile}", VBA.Environ$("USERPROFILE")
    pCustomVars.Add "${Temp}", VBA.Environ$("TEMP")
    pCustomVars.Add "${ApplicationPath}", Application.CurrentProject.Path
    
    s = FileSystemLib.PathCombine(Application.CurrentProject.Path, "package.json")
    If FileSystemLib.TryGetFileInAncestors(s, 3) Then
        If FileSystemLib.TryReadAllTextInFile(s, s, False) Then
            On Error GoTo Finally
            Set dX = DictionaryEx.Create(DictionaryEx.Create(s)("ds3x.CustomVars"))
            For Each Item In dX
                dsApp.CustomVar(CStr(Item(0))) = CStr(Item(1))
            Next Item
        End If
    End If
Finally:
End Sub

Private Function ReplaceCustomVars(ByVal Target As Variant) As Variant
    Select Case CLng(VarType(Target))
        Case Is = CLng(vbString)
            ReplaceCustomVars = ReplaceCustomVarsOnString(Target)
        Case Is >= CLng(vbArray)
            ReplaceCustomVars = ReplaceCustomVarsOnIterable(Target)
        Case Else
            If IsObject(Target) Then
                Set ReplaceCustomVars = Target
            Else
                ReplaceCustomVars = Target
            End If
    End Select
End Function

Private Function ReplaceCustomVarsOnIterable(ByVal Target As Variant) As Variant
    Dim aX As ArrayListEx, Item As Variant
    Set aX = ArrayListEx.Create()
    
    For Each Item In Target
        aX.Add ReplaceCustomVars(Item)
    Next Item
    
    ReplaceCustomVarsOnIterable = aX.ToArray()
End Function

Private Function ReplaceCustomVarsOnString(ByVal Target As String) As String
    Dim cKey As Variant
    
    For Each cKey In pCustomVars.Keys()
        If InStr(1, Target, CStr(cKey), vbTextCompare) > 0 Then
            Target = Replace(Target, CStr(cKey), CStr(pCustomVars(cKey)), Compare:=vbTextCompare)
        End If
    Next cKey
    
    ReplaceCustomVarsOnString = Target
End Function


' --- Logging ---

Private Sub AppendToLogFile(ByVal LogLevel As String, ByVal LogMessage As String)
    Static sLogFile As String, isFileMissing As Boolean
    Dim sItem As String, sFile As Variant
    
    sItem = Printf("[%1] %2%3 - %4", UCase(Left(LogLevel, 12)), VBA.Space$(13 - Len(Left(LogLevel, 12))), Time(), LogMessage)
    Debug.Print sItem
    If isFileMissing Then Exit Sub
    
    On Error GoTo HandleFileMissing
    If sLogFile = vbNullString Then
        sFile = dsApp.LoadConfig("package.json")("ds3x.LogFile")
        If IsEmpty(sFile) Or Trim(CStr(sFile)) = "" Then GoTo HandleFileMissing
        sLogFile = FileSystemLib.Resolve(ApplyCustomVarsOn(sFile))
    End If
    
    FileSystemLib.TryAppendTextInFile sLogFile, sItem & vbNewLine, False
    
    Exit Sub
HandleFileMissing:
    isFileMissing = True
End Sub

Private Function GetConfigFromFile(Optional ByVal TargetFile As String = "package.json") As DictionaryEx
    Dim s As String
    On Error GoTo Finally
    Set GetConfigFromFile = DictionaryEx.Create()
    
    s = FileSystemLib.Resolve(TargetFile)
    If FileSystemLib.TryGetFileInAncestors(s, 3) Then
        If FileSystemLib.TryReadAllTextInFile(s, s, False) Then
            Set GetConfigFromFile = DictionaryEx.Create(s)
        End If
    End If
Finally:
End Function




