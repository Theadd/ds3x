VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Callback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "ds3x.Collections"
Option Compare Database
Option Explicit

Private pInstance As Object
Private pName As String
Private pCallType As VBA.VbCallType


' --- PUBLIC ---

Public Result As Variant
Public Event ExecuteComplete(TargetCallable As Callback)


' --- PROPERTY SETTERS ---

'Public Property Set Instance(ByRef Value As Object): Set pInstance = Value: End Property
'Public Property Let Name(ByVal Value As String): pName = Value: End Property
'Public Property Let CallType(ByVal Value As VBA.VbCallType): pCallType = Value: End Property


' --- PUBLIC FUNCTIONS ---

Public Function Create(Optional ByRef ObjectInstance As Object = Nothing, Optional ByVal CallableName As String = vbNullString, Optional ByVal CallableType As VBA.VbCallType = VBA.VbCallType.VbMethod) As Callback
    Set Create = New Callback
    Create.Bind ObjectInstance, CallableName, CallableType
End Function

Public Function Bind(Optional ByRef ObjectInstance As Object = Nothing, Optional ByVal CallableName As String = vbNullString, Optional ByVal CallableType As VBA.VbCallType = VBA.VbCallType.VbMethod) As Callback
    Set pInstance = ObjectInstance
    pName = CallableName
    pCallType = CallableType
    Set Bind = Me
End Function

Public Function Unbind() As Callback
    Set pInstance = Nothing
    pName = vbNullString
    pCallType = VBA.VbCallType.VbMethod
    Set Unbind = Me
End Function

Public Function Self() As Callback: Set Self = Me: End Function

'@DefaultMember
Public Function Execute(ParamArray vArgs() As Variant) As Callback
Attribute Execute.VB_UserMemId = 0
    If Not pInstance Is Nothing Then
        ' CallByName
        Select Case UBound(vArgs)
            Case -1: Assign Result, CallByName(pInstance, pName, VbMethod)
            Case 0: Assign Result, CallByName(pInstance, pName, VbMethod, vArgs(0))
            Case 1: Assign Result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1))
            Case 2: Assign Result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1), vArgs(2))
            Case 3: Assign Result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1), vArgs(2), vArgs(3))
            Case 4: Assign Result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4))
            Case 5: Assign Result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5))
            Case 6: Assign Result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5), vArgs(6))
            Case 7: Assign Result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5), vArgs(6), vArgs(7))
            Case 8: Assign Result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5), vArgs(6), vArgs(7), vArgs(8))
            Case 9: Assign Result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5), vArgs(6), vArgs(7), vArgs(8), vArgs(9))
        End Select
    ElseIf pInstance Is Nothing And pName <> "" Then
        ' Application.Run
        Select Case UBound(vArgs)
            Case -1: Assign Result, Application.Run(pName)
            Case 0: Assign Result, Application.Run(pName, vArgs(0))
            Case 1: Assign Result, Application.Run(pName, vArgs(0), vArgs(1))
            Case 2: Assign Result, Application.Run(pName, vArgs(0), vArgs(1), vArgs(2))
            Case 3: Assign Result, Application.Run(pName, vArgs(0), vArgs(1), vArgs(2), vArgs(3))
            Case 4: Assign Result, Application.Run(pName, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4))
            Case 5: Assign Result, Application.Run(pName, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5))
            Case 6: Assign Result, Application.Run(pName, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5), vArgs(6))
            Case 7: Assign Result, Application.Run(pName, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5), vArgs(6), vArgs(7))
            Case 8: Assign Result, Application.Run(pName, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5), vArgs(6), vArgs(7), vArgs(8))
            Case 9: Assign Result, Application.Run(pName, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5), vArgs(6), vArgs(7), vArgs(8), vArgs(9))
        End Select
    Else
        Select Case UBound(vArgs)
            Case -1: Assign Result, Empty
            Case 0: Assign Result, vArgs(0)
            Case Else: Assign Result, CArray(vArgs)
        End Select
    End If

    Set Execute = Me
    RaiseEvent ExecuteComplete(Execute)
End Function


Public Function ExecuteOnArray(ByRef vArgs() As Variant) As Callback
    If Not pInstance Is Nothing Then
        ' CallByName
        Select Case UBound(vArgs)
            Case -1: Assign Result, CallByName(pInstance, pName, VbMethod)
            Case 0: Assign Result, CallByName(pInstance, pName, VbMethod, vArgs(0))
            Case 1: Assign Result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1))
            Case 2: Assign Result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1), vArgs(2))
            Case 3: Assign Result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1), vArgs(2), vArgs(3))
            Case 4: Assign Result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4))
            Case 5: Assign Result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5))
            Case 6: Assign Result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5), vArgs(6))
            Case 7: Assign Result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5), vArgs(6), vArgs(7))
            Case 8: Assign Result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5), vArgs(6), vArgs(7), vArgs(8))
            Case 9: Assign Result, CallByName(pInstance, pName, VbMethod, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5), vArgs(6), vArgs(7), vArgs(8), vArgs(9))
        End Select
    ElseIf pInstance Is Nothing And pName <> "" Then
        ' Application.Run
        Select Case UBound(vArgs)
            Case -1: Assign Result, Application.Run(pName)
            Case 0: Assign Result, Application.Run(pName, vArgs(0))
            Case 1: Assign Result, Application.Run(pName, vArgs(0), vArgs(1))
            Case 2: Assign Result, Application.Run(pName, vArgs(0), vArgs(1), vArgs(2))
            Case 3: Assign Result, Application.Run(pName, vArgs(0), vArgs(1), vArgs(2), vArgs(3))
            Case 4: Assign Result, Application.Run(pName, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4))
            Case 5: Assign Result, Application.Run(pName, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5))
            Case 6: Assign Result, Application.Run(pName, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5), vArgs(6))
            Case 7: Assign Result, Application.Run(pName, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5), vArgs(6), vArgs(7))
            Case 8: Assign Result, Application.Run(pName, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5), vArgs(6), vArgs(7), vArgs(8))
            Case 9: Assign Result, Application.Run(pName, vArgs(0), vArgs(1), vArgs(2), vArgs(3), vArgs(4), vArgs(5), vArgs(6), vArgs(7), vArgs(8), vArgs(9))
        End Select
    Else
        Select Case UBound(vArgs)
            Case -1: Assign Result, Empty
            Case 0: Assign Result, vArgs(0)
            Case Else: Assign Result, CArray(vArgs)
        End Select
    End If

    Set ExecuteOnArray = Me
    RaiseEvent ExecuteComplete(ExecuteOnArray)
End Function

Public Function PrintAll() As Callback
    Debug.Print JSON.Stringify(Result, 2)
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
