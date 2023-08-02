VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Callback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Private pInstance As Object
Private pName As String
Private pCallType As VBA.VbCallType


' --- PUBLIC ---

Public result As Variant
Public Event ExecuteComplete(TargetCallable As Callback)


' --- PROPERTY SETTERS ---

'Public Property Set Instance(ByRef Value As Object): Set pInstance = Value: End Property
'Public Property Let Name(ByVal Value As String): pName = Value: End Property
'Public Property Let CallType(ByVal Value As VBA.VbCallType): pCallType = Value: End Property


' --- PUBLIC FUNCTIONS ---

Public Function Create(Optional ByRef ObjectInstance As Object = Nothing, Optional ByVal CallableName As String = "", Optional ByVal CallableType As VBA.VbCallType = VBA.VbCallType.VbMethod) As Callback
    Set Create = New Callback
    Create.Bind ObjectInstance, CallableName, CallableType
End Function

Public Function Bind(Optional ByRef ObjectInstance As Object = Nothing, Optional ByVal CallableName As String = "", Optional ByVal CallableType As VBA.VbCallType = VBA.VbCallType.VbMethod) As Callback
    Set pInstance = ObjectInstance
    pName = CallableName
    pCallType = CallableType
    Set Bind = Me
End Function

Public Function Unbind() As Callback
    Set pInstance = Nothing
    pName = ""
    pCallType = VBA.VbCallType.VbMethod
    Set Unbind = Me
End Function

Public Function Self() As Callback: Set Self = Me: End Function

Public Function Execute(ParamArray vArgs() As Variant) As Callback
Attribute Execute.VB_UserMemId = 0
    If Not pInstance Is Nothing Then
        ' CallByName
        Select Case UBound(vArgs)
            Case -1: Assign result, CallByName(pInstance, pName, VbMethod)
            Case 0: Assign result, CallByName(pInstance, pName, VbMethod, vArgs(0))
            Case 1: Assign result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1))
            Case 2: Assign result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1), vArgs(2))
            Case 3: Assign result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1), vArgs(2), vArgs(3))
            Case 4: Assign result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4))
        End Select
    ElseIf pInstance Is Nothing And pName <> "" Then
        ' Application.Run
        Select Case UBound(vArgs)
            Case -1: Assign result, Application.Run(pName)
            Case 0: Assign result, Application.Run(pName, vArgs(0))
            Case 1: Assign result, Application.Run(pName, vArgs(0), vArgs(1))
            Case 2: Assign result, Application.Run(pName, vArgs(0), vArgs(1), vArgs(2))
            Case 3: Assign result, Application.Run(pName, vArgs(0), vArgs(1), vArgs(2), vArgs(3))
            Case 4: Assign result, Application.Run(pName, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4))
        End Select
    Else
        Select Case UBound(vArgs)
            Case -1: Assign result, Empty
            Case 0: Assign result, vArgs(0)
            Case Else: Assign result, CArray(vArgs)
        End Select
    End If
    
    Set Execute = Me
    RaiseEvent ExecuteComplete(Execute)
'    Set Execute = Me
'    ExecuteOnArray vArgs
End Function

'Public Function ExecuteOnArray(ByVal vArgs As Variant) As Callback
Public Function ExecuteOnArray(ByRef vArgs() As Variant) As Callback
    'If Not IsArray(vArgs) Then vArgs = Array(vArgs)
    If Not pInstance Is Nothing Then
        ' CallByName
        Select Case UBound(vArgs)
            Case -1: Assign result, CallByName(pInstance, pName, VbMethod)
            Case 0: Assign result, CallByName(pInstance, pName, VbMethod, vArgs(0))
            Case 1: Assign result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1))
            Case 2: Assign result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1), vArgs(2))
            Case 3: Assign result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1), vArgs(2), vArgs(3))
            Case 4: Assign result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4))
            Case 5: Assign result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5))
            Case 6: Assign result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5), vArgs(6))
            Case 7: Assign result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5), vArgs(6), vArgs(7))
            Case 8: Assign result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5), vArgs(6), vArgs(7), vArgs(8))
            Case 9: Assign result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5), vArgs(6), vArgs(7), vArgs(8), vArgs(9))
        End Select
    ElseIf pInstance Is Nothing And pName <> "" Then
        ' Application.Run
        Select Case UBound(vArgs)
            Case -1: Assign result, Application.Run(pName)
            Case 0: Assign result, Application.Run(pName, vArgs(0))
            Case 1: Assign result, Application.Run(pName, vArgs(0), vArgs(1))
            Case 2: Assign result, Application.Run(pName, vArgs(0), vArgs(1), vArgs(2))
            Case 3: Assign result, Application.Run(pName, vArgs(0), vArgs(1), vArgs(2), vArgs(3))
            Case 4: Assign result, Application.Run(pName, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4))
            Case 5: Assign result, Application.Run(pName, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5))
            Case 6: Assign result, Application.Run(pName, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5), vArgs(6))
            Case 7: Assign result, Application.Run(pName, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5), vArgs(6), vArgs(7))
            Case 8: Assign result, Application.Run(pName, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5), vArgs(6), vArgs(7), vArgs(8))
            Case 9: Assign result, Application.Run(pName, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5), vArgs(6), vArgs(7), vArgs(8), vArgs(9))
        End Select
    Else
        Select Case UBound(vArgs)
            Case -1: Assign result, Empty
            Case 0: Assign result, vArgs(0)
            Case Else: Assign result, CArray(vArgs)
        End Select
    End If
    
    Set ExecuteOnArray = Me
    RaiseEvent ExecuteComplete(ExecuteOnArray)
End Function

' Just like Array.Where() in .NET or Array.Filter() in JavaScript.
'
' @REQUIRES:
'     1. A reference to "Microsoft Script Control 1.0"
'     2. ".NET Framework v3.5" installed and a reference to "mscorlib.dll" (by opening "mscorlib.tlb" from references window, usually found at "C:\Windows\Microsoft.NET\Framework\v4.0.30319")
'
' @EXAMPLE:
'     Callback.Create()(Array(1, 3, 5, 2, 4, 6)).Filter("x => x < 5")   '.Result contains: [1, 3, 2, 4]
'
Public Function Filter(ByVal PredicateFunction As String) As Callback
    Dim aList As Object, i As Long, eScript As Object, jsFunc As Variant
    
    If TryParseAsJScriptFunction(PredicateFunction, "fn", jsFunc) Then
        Set aList = CreateObject("System.Collections.ArrayList")
        Set eScript = CreateObject("MSScriptControl.ScriptControl")
        eScript.Language = "JScript"
        eScript.ExecuteStatement "var Result = eval(" & JSON.Stringify(result) & ");"
        eScript.AddCode jsFunc(0)
        
        For i = 0 To UBound(result)
            If CBool(eScript.Eval("fn(Result[" & i & "], " & i & ", Result);")) Then
                aList.Add result(i)
            End If
        Next i
    
        result = aList.ToArray
    End If
    
    Set Filter = Me
End Function

Public Function Map(ByVal SelectorFunction As String) As Callback
    Dim aList As Object, i As Long, eScript As Object, jsFunc As Variant
    
    If TryParseAsJScriptFunction(SelectorFunction, "fn", jsFunc) Then
        Set aList = CreateObject("System.Collections.ArrayList")
        Set eScript = CreateObject("MSScriptControl.ScriptControl")
        eScript.Language = "JScript"
        eScript.AddCode CollectionsLib.JScriptCode("JSON.serialize")
        eScript.ExecuteStatement "var Result = eval(" & JSON.Stringify(result) & ");"
        eScript.AddCode jsFunc(0)
        
        For i = 0 To UBound(result)
            aList.Add JSON.Parse(eScript.Eval("JSON.serialize(fn(Result[" & i & "], " & i & ", Result));"))
        Next i
        
        result = aList.ToArray
    End If

    Set Map = Me
End Function

Public Function Reduce(ByVal AccumulatorFunction As String, Optional ByVal InitialValue As Variant) As Callback
    Dim aList As Object, i As Long, eScript As Object, jsFunc As Variant, iStart As Long
    
    If TryParseAsJScriptFunction(AccumulatorFunction, "fn", jsFunc) Then
        Set aList = CreateObject("System.Collections.ArrayList")
        Set eScript = CreateObject("MSScriptControl.ScriptControl")
        eScript.Language = "JScript"
        eScript.AddCode CollectionsLib.JScriptCode("JSON.serialize")
        eScript.ExecuteStatement "var Result = eval(" & JSON.Stringify(result) & ");"
        iStart = LBound(result)
        If IsMissing(InitialValue) Then
            InitialValue = result(LBound(result))
            iStart = LBound(result) + 1
        End If
        eScript.ExecuteStatement "var __Acc = eval(" & JSON.Stringify(InitialValue) & ");"
        eScript.AddCode jsFunc(0)
        
        For i = iStart To UBound(result)
            If i = UBound(result) Then
                Assign result, JSON.Parse(eScript.Eval("JSON.serialize(fn(__Acc, Result[" & i & "], " & i & ", Result))"))
            Else
                eScript.ExecuteStatement "__Acc = fn(__Acc, Result[" & i & "], " & i & ", Result);"
            End If
        Next i
    End If

    Set Reduce = Me
End Function


Public Function PrintAll() As Callback
    Debug.Print JSON.Stringify(result, 2)
    Set PrintAll = Me
End Function


' --- PRIVATE METHODS ---

Private Sub Assign(ByRef Target, Optional ByRef Value As Variant)
    If Not IsMissing(Value) Then
        If IsObject(Value) Then
            Set Target = Value
        Else
            Target = Value
        End If
    End If
End Sub

' Convert Variant To Varaint()
Private Function CArray(ByVal xs) As Variant()
    On Error GoTo ErrorHandler
    CArray = xs
ExitFunc:
    Exit Function
ErrorHandler:
    Debug.Print "[ERROR] " & CStr(Err.Number) & " - " & CStr(Err.Description)
    Resume ExitFunc
End Function



