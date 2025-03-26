Attribute VB_Name = "ds3xGlobals"
Option Compare Database
Option Explicit
Option Base 0


' --- ACCESS WINDOW HIDE / SHOW ---

Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOW = 5


' --- ScreenLib Types ---

Public Type POINTAPI
    x As Long
    Y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type BOUNDS
    x As Long
    Y As Long
    W As Long
    h As Long
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
    Public Sub AddRunnableTask(ByVal targetPath As String, Optional ByVal RunnableTaskName As String = "", Optional ByVal OnErrorResumeNext As Boolean = False)
        dsApp.RunnableTasks.Add Array(targetPath, RunnableTaskName, OnErrorResumeNext)
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

Public Function HttpGetRequest(url As String, Optional headers As Variant) As String
  Dim objHTTP As Object
  Dim strResponse As String
  
  On Error GoTo ErrorHandler
  
  Set objHTTP = CreateObject("MSXML2.XMLHTTP")
  objHTTP.Open "GET", url, False
  
  If Not IsMissing(headers) Then
    Dim key As Variant
    For Each key In headers
      objHTTP.setRequestHeader key, headers(key)
    Next key
  End If
  
  objHTTP.send
  
  If objHTTP.Status = 200 Then
    strResponse = objHTTP.responseText
  Else
    strResponse = "Error: " & objHTTP.Status & " - " & objHTTP.statusText
  End If
  
  HttpGetRequest = strResponse
  
  Set objHTTP = Nothing
  Exit Function
  
ErrorHandler:
  HttpGetRequest = "Error: " & Err.Number & " - " & Err.Description
  Set objHTTP = Nothing
End Function
