Attribute VB_Name = "StyleLib"
' StyleLib Module
Option Compare Database
Option Explicit



Public Enum StyleType
    Locked
    Disabled
    Grayed
    Invalid
End Enum

Private Const SC_GRAYED_COLOR As Long = 15461355
Private Const SC_GRAYED_SHADE As Long = 92
Private Const SC_GRAYED_THEME As Long = 1
Private Const SC_GRAYED_TINT As Long = 100

Private Const SC_WHITE_COLOR As Long = 16777215
Private Const SC_WHITE_SHADE As Long = 100
Private Const SC_WHITE_THEME As Long = 1
Private Const SC_WHITE_TINT As Long = 100

Private Const SC_INVALID_COLOR As Long = 2366701




'''''''''''''''''''''''''''''''' ALL IN ONE ''''''''''''''''''''''''''''''

Public Function AddStyle(ByRef TargetControl As Access.Control, ParamArray Styles() As Variant) As Access.Control
    Dim Style As Variant
    
    For Each Style In Styles
        Select Case Style
            Case StyleType.Locked
                
            Case StyleType.Disabled
            
            Case StyleType.Grayed
                SetGrayedStyle TargetControl, True
            Case StyleType.Invalid
                SetGridlineAsValid TargetControl, False
            Case Else
                If DEBUG_MODE_ENABLED Then Stop
        End Select
    Next Style
    
    Set AddStyle = TargetControl
End Function

Public Function RemoveStyle(ByRef TargetControl As Access.Control, ParamArray Styles() As Variant) As Access.Control
    Dim Style As Variant
    
    For Each Style In Styles
        Select Case Style
            Case StyleType.Locked
                
            Case StyleType.Disabled
            
            Case StyleType.Grayed
                SetGrayedStyle TargetControl, False
            Case StyleType.Invalid
                SetGridlineAsValid TargetControl, True
            Case Else
                If DEBUG_MODE_ENABLED Then Stop
        End Select
    Next Style
    
    Set RemoveStyle = TargetControl
End Function

Public Function HasStyle(ByRef TargetControl As Access.Control, ByVal Style As StyleType) As Boolean
    Select Case Style
        Case StyleType.Locked
            HasStyle = (TargetControl.Locked)
        Case StyleType.Disabled
            HasStyle = (Not TargetControl.Enabled)
        Case StyleType.Grayed
            HasStyle = (TargetControl.BackColor = SC_GRAYED_COLOR)
        Case StyleType.Invalid
            HasStyle = (TargetControl.GridlineColor = SC_INVALID_COLOR)
        Case Else
            ' TODO
            If DEBUG_MODE_ENABLED Then Stop
    End Select
End Function

Public Function ResetAllControlStylesInForm(ByRef TargetForm As Access.Form)
    Dim i As Long, isLocked As Boolean
    isLocked = (Not TargetForm.AllowEdits)
    
    For i = 0 To TargetForm.Controls.Count - 1
        Select Case TargetForm.Controls(i).ControlType
            Case acToggleButton
                SetGridlineAsValid TargetForm.Controls(i), True
            Case acTextBox, acComboBox
                SetGridlineAsValid TargetForm.Controls(i), True
                SetGrayedStyle TargetForm.Controls(i), isLocked
            Case acCommandButton
                SetGridlineAsValid TargetForm.Controls(i), True
                SetGrayedStyle TargetForm.Controls(i), isLocked
                If Len(TargetForm.Controls(i).Caption) = 1 Then TargetForm.Controls(i).Enabled = (Not isLocked)
            Case acListBox
                'SetGridlineAsValid TargetForm.Controls(i), True
                SetGrayedStyle TargetForm.Controls(i), isLocked
        End Select
    Next i
End Function




''''''''''''''''''''''''''' GRAYED / WHITE '''''''''''''''''''''''''''

Private Function SetGrayedStyle(ByRef TargetControl As Access.Control, ByVal asGrayed As Boolean)
    Dim TargetLabel As Access.Label, styleLabel As Boolean
    
    'If TargetControl.BackColor = IIf(asGrayed, SC_GRAYED_COLOR, SC_WHITE_COLOR) Then Exit Function
    If TargetControl.BackColor <> IIf(asGrayed, SC_WHITE_COLOR, SC_GRAYED_COLOR) Then Exit Function
    
    If TryGetAttachedLabel(TargetControl, TargetLabel) Then
        styleLabel = (TargetLabel.BackColor = TargetControl.BackColor)
    End If
    
    With TargetControl
        .BackColor = IIf(asGrayed, SC_GRAYED_COLOR, SC_WHITE_COLOR)
        .BackShade = IIf(asGrayed, SC_GRAYED_SHADE, SC_WHITE_SHADE)
        .BackThemeColorIndex = IIf(asGrayed, SC_GRAYED_THEME, SC_WHITE_THEME)
        .BackTint = IIf(asGrayed, SC_GRAYED_TINT, SC_WHITE_TINT)
        .BorderColor = IIf(asGrayed, SC_GRAYED_COLOR, SC_WHITE_COLOR)
        .BorderShade = IIf(asGrayed, SC_GRAYED_SHADE, SC_WHITE_SHADE)
        .BorderThemeColorIndex = IIf(asGrayed, SC_GRAYED_THEME, SC_WHITE_THEME)
        .BorderTint = IIf(asGrayed, SC_GRAYED_TINT, SC_WHITE_TINT)
    End With
    
    If styleLabel Then
        With TargetLabel
            .BackColor = IIf(asGrayed, SC_GRAYED_COLOR, SC_WHITE_COLOR)
            .BackShade = IIf(asGrayed, SC_GRAYED_SHADE, SC_WHITE_SHADE)
            .BackThemeColorIndex = IIf(asGrayed, SC_GRAYED_THEME, SC_WHITE_THEME)
            .BackTint = IIf(asGrayed, SC_GRAYED_TINT, SC_WHITE_TINT)
            .BorderColor = IIf(asGrayed, SC_GRAYED_COLOR, SC_WHITE_COLOR)
            .BorderShade = IIf(asGrayed, SC_GRAYED_SHADE, SC_WHITE_SHADE)
            .BorderThemeColorIndex = IIf(asGrayed, SC_GRAYED_THEME, SC_WHITE_THEME)
            .BorderTint = IIf(asGrayed, SC_GRAYED_TINT, SC_WHITE_TINT)
        End With
    End If
End Function



''''''''''''''''''''''''' GRIDLINE VALIDATION '''''''''''''''''''''''''''''

Public Function SetGridlineAsValid(ByRef TargetControl As Access.Control, Optional ByVal esValido As Boolean = True) As Boolean
    Dim TargetLabel As Access.Label
    On Error GoTo Finally
    
    If Not esValido Then
        If TargetControl.GridlineColor <> SC_INVALID_COLOR Then
            TargetControl.Tag = CStr(TargetControl.GridlineColor)
            TargetControl.GridlineColor = SC_INVALID_COLOR
            If TryGetAttachedLabel(TargetControl, TargetLabel) Then
                TargetLabel.GridlineColor = SC_INVALID_COLOR
            End If
        End If
    Else
        If TargetControl.GridlineColor = SC_INVALID_COLOR Then
            TargetControl.GridlineColor = CLng(TargetControl.Tag)
            If TryGetAttachedLabel(TargetControl, TargetLabel) Then
                TargetLabel.GridlineColor = CLng(TargetControl.Tag)
            End If
            TargetControl.Tag = ""
        End If
    End If
    
Finally:
    SetGridlineAsValid = esValido
End Function

Public Function GridlineValidation(ByRef TargetControl As Access.Control, ParamArray ValidValues() As Variant) As Boolean
    Dim esValido As Boolean, i As Long, strValor As String, TargetLabel As Access.Label
    esValido = (UBound(ValidValues) = -1)
    On Error GoTo Finally
    
    If Not esValido Then
        strValor = CStr(Nz(TargetControl, ""))
        For i = 0 To UBound(ValidValues)
            If CStr(ValidValues(i)) = strValor Then
                esValido = True
                Exit For
            End If
        Next i
    End If
    
    If Not esValido Then
        If TargetControl.GridlineColor <> SC_INVALID_COLOR Then
            TargetControl.Tag = CStr(TargetControl.GridlineColor)
            TargetControl.GridlineColor = SC_INVALID_COLOR
            If TryGetAttachedLabel(TargetControl, TargetLabel) Then
                TargetLabel.GridlineColor = SC_INVALID_COLOR
            End If
        End If
    Else
        If TargetControl.GridlineColor = SC_INVALID_COLOR Then
            TargetControl.GridlineColor = CLng(TargetControl.Tag)
            If TryGetAttachedLabel(TargetControl, TargetLabel) Then
                TargetLabel.GridlineColor = CLng(TargetControl.Tag)
            End If
            TargetControl.Tag = ""
        End If
    End If
    
Finally:
    GridlineValidation = esValido
End Function





Public Function SetControlAsEnabled(ByRef TargetControl As Access.Control, Optional ByVal Enable As Boolean = True) As Boolean
    SetControlAsEnabled = Enable
    SetGrayedStyle TargetControl, (Not Enable)
    On Error GoTo Fallback
    TargetControl.Locked = (Not Enable)
    Exit Function
Fallback:
    TargetControl.Enabled = Enable
End Function




'''''''''''''''''''''''''''''''''''''' HELPERS ''''''''''''''''''''''''''''''''''

Public Function TryGetAttachedLabel(ByRef TargetControl As Access.Control, ByRef AttachedLabel As Access.Label) As Boolean
    On Error GoTo MissingLabel
    
    If TypeName(TargetControl.Controls(0)) = "Label" Then
        Set AttachedLabel = TargetControl.Controls(0)
        TryGetAttachedLabel = True
        Exit Function
    End If

MissingLabel:
    TryGetAttachedLabel = False
End Function


