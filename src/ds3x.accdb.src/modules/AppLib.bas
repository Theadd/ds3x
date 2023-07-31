Attribute VB_Name = "AppLib"
Option Compare Database
Option Explicit

' MODULO: APP LIBRARY




Public Function IsFormOpenAsStandaloneWindow(ByVal TargetFormName As String, ByRef TargetForm As Access.Form) As Boolean
    On Error GoTo Finally
    IsFormOpenAsStandaloneWindow = False
    
    If CurrentProject.AllForms(TargetFormName).IsLoaded Then
        Set TargetForm = Forms(TargetFormName)
        IsFormOpenAsStandaloneWindow = True
    End If
Finally:
End Function

Public Function GetParentFormOfControl(ByRef bControl As Access.Control) As Access.Form
    On Error GoTo HandleError
    Dim bForm As Access.Form
    
    Set bForm = bControl.Parent
    Set GetParentFormOfControl = bForm
    Exit Function
    
HandleError:
    Set bForm = GetParentFormOfControl(bControl.Parent)
    Set GetParentFormOfControl = bForm
End Function

Public Function GetTopParentFormWindow(ByRef TargetForm As Access.Form, Optional ByRef UpToParentHwnd As Long = -1) As Access.Form
    On Error GoTo Finally
    
    If TargetForm.hWnd = UpToParentHwnd Then GoTo Finally
    
    If TypeOf TargetForm.Parent Is Form Then
        Set GetTopParentFormWindow = GetTopParentFormWindow(TargetForm.Parent, UpToParentHwnd)
    Else
        Set GetTopParentFormWindow = GetTopParentFormWindow(TargetForm.Parent.Parent, UpToParentHwnd)
    End If
    Exit Function
    
Finally:
    Set GetTopParentFormWindow = TargetForm
End Function

Public Function IsAncestorOfForm(ByRef ParentForm As Access.Form, ByRef ChildForm As Access.Form) As Boolean
    IsAncestorOfForm = (GetTopParentFormWindow(ChildForm, ParentForm.hWnd).hWnd = ParentForm.hWnd)
End Function

Public Function TryGetActiveFormWindow(ByRef TargetForm As Access.Form) As Boolean
    On Error GoTo ErrorHandler
    
    Set TargetForm = Screen.ActiveForm
    
    TryGetActiveFormWindow = (Len(TargetForm.Name) > 0)
    Exit Function
    
ErrorHandler:
End Function



