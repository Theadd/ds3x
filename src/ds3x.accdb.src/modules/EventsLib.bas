Attribute VB_Name = "EventsLib"
' EventsLib Module
Option Compare Database
Option Explicit



Private Function GetScreenRectOfActiveForm() As RECT
    On Error GoTo Fallback
    
    GetScreenRectOfActiveForm = GetScreenRectOfPoint(PointInRect(GetWindowRect(Screen.ActiveForm), DirectionType.Center))
    Exit Function
Fallback:
    GetScreenRectOfActiveForm = GetScreenRectOfPoint(GetPoint())
End Function


Public Sub Focus(ByRef frm As Access.Form)
    On Error Resume Next
    frm.Controls("HiddenControl").SetFocus
    DoEvents
End Sub




'''''''''''''''''''''''' KEYBOARD SHORTCUTS ''''''''''''''''''''''''

Public Function TryHandleKeyboardShortcut(ByRef mTargetForm As Access.Form, ByRef KeyCode As Integer, ByRef Shift As Integer) As Boolean
    On Error GoTo ErrorHandler
    Dim isHandled As Boolean

    'Debug.Print Printf("IN TryHandleKeyboardShortcut(%1, %2, %3)", mTargetForm.Name, KeyCode, Shift)
    
    Select Case Shift
        Case 2
            Select Case KeyCode
                Case vbKeyA, vbKeyE
                    ' CTRL + A, CTRL + E
                      Select Case GetActiveControlAcType
                          Case acTextBox, acComboBox
                              isHandled = TrySelectAllTextInTextBox(Screen.ActiveControl)
                      End Select
            End Select
        Case 0
            ' No active key modifiers
        Case 4
            Select Case KeyCode
                Case vbKeyF11
                    isHandled = True
            End Select
    End Select
    
    ' TEXTBOX SPECIALS
    On Error GoTo NoSpecialsToHandle
    If Not isHandled Then
        Select Case GetActiveControlAcType
            Case acTextBox, acComboBox
                If Len(Nz(Screen.ActiveControl.StatusBarText)) > 0 Then
                    isHandled = TryHandleTextBoxSpecials(Screen.ActiveControl, KeyCode, Shift)
                End If
        End Select
    End If
    
NoSpecialsToHandle:
    On Error GoTo ErrorHandler
'    ' LIMIT TEXTBOX MAX LENGTH
'    If Not isHandled And (GetActiveControlType = "TextBox") Then
'        If Len(Nz(Screen.ActiveControl.StatusBarText)) > 0 Then
'            isHandled = TryHandleTextBoxSpecials(Screen.ActiveControl, KeyCode, Shift)
'        End If
'    End If
    
    If isHandled Then
        KeyCode = 0
    End If
    TryHandleKeyboardShortcut = isHandled
    
    Exit Function
ErrorHandler:
    TryHandleKeyboardShortcut = False
    If DEBUG_MODE_ENABLED Then Stop
End Function

Private Function TryTriggerOnEscapeEvent(ByRef mTargetForm As Access.Form) As Boolean
    On Error GoTo NoEscapeActionFound
    
    mTargetForm.OnEscape
    TryTriggerOnEscapeEvent = True
    
NoEscapeActionFound:
End Function

Public Function TrySelectAllTextInTextBox(ByRef TargetTextBox As Access.Control) As Boolean
    On Error GoTo Fallback
    Dim AuxLong As Long
    
    TargetTextBox.SetFocus
    AuxLong = Len(TargetTextBox.Text)
    TargetTextBox.selStart = 0
    TargetTextBox.SelLength = AuxLong
    TrySelectAllTextInTextBox = True
Fallback:
End Function

Private Function GetActiveControlType() As String
    On Error GoTo Fallback
    
    GetActiveControlType = TypeName(Screen.ActiveControl)
    Exit Function
Fallback:
    GetActiveControlType = ""
End Function

Private Function GetActiveControlAcType() As Long
    On Error GoTo Finally
    
    GetActiveControlAcType = Screen.ActiveControl.ControlType
    
Finally:
End Function

Private Function TryHandleTextBoxMaxLength(ByVal maxLen As Long, ByRef TargetTextBox As Access.Control, ByRef KeyCode As Integer, ByRef Shift As Integer) As Boolean
    Dim curLen As Long, Aux As String, startSlice As String, sValue As String, endSlice As String, midSlice As String
    Dim availCapacity As Long, i As Long, cSize As Long, kAscii As Integer
    
    With TargetTextBox
        curLen = GetTextBoxLength(TargetTextBox)
        
        If curLen > maxLen Then
            TargetTextBox.Text = Left(GetTextUnsafe(TargetTextBox), maxLen)
            curLen = maxLen
        End If
        
        If KeyCode = vbKeyV And Shift = 2 Then
            ' CTRL + V
            sValue = GetTextUnsafe(TargetTextBox)
            startSlice = VBA.Mid(sValue, 1, .selStart)
            endSlice = VBA.Mid(sValue, 1 + .selStart + .SelLength)
            Aux = GetTextFromClipboard
            
            cSize = Len(startSlice) + Len(endSlice)
            availCapacity = maxLen - cSize
            midSlice = ""
            
            For i = 1 To Len(Aux)
                If Len(midSlice) >= availCapacity Then Exit For
                kAscii = VBA.AscW(VBA.Mid$(Aux, i, 1))
                If Not TryHandleTextBoxLateSpecials(TargetTextBox, kAscii) Then
                    midSlice = midSlice & VBA.ChrW(kAscii)
                End If
            Next i
            
            TargetTextBox.Text = startSlice & midSlice & endSlice
            TryHandleTextBoxMaxLength = True
            .selStart = Len(startSlice) + Len(midSlice)
            .SelLength = 0

        ElseIf curLen = maxLen Then
            Select Case KeyCode
                Case vbKeyDown, vbKeyUp, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyDelete, vbKeyHome, vbKeyEnd, vbKeyEscape
                    ' Allowed
                Case Else
                    TryHandleTextBoxMaxLength = (.SelLength < 1)
            End Select
        End If
    End With
    
End Function

Private Function TryHandleTextBoxSpecials(ByRef TargetTextBox As Access.Control, ByRef KeyCode As Integer, ByRef Shift As Integer) As Boolean
    ' TextBox SPECIAL PROPERTIES HANDLING ON KeyDown EVENT
    Dim Item As Variant
    
    With TargetTextBox
        For Each Item In Split(Nz(.StatusBarText, ""), ";")
            Select Case Split(Item, "=")(0)
                Case "MAXLEN"
                    TryHandleTextBoxSpecials = TryHandleTextBoxMaxLength(CLng(Val(Split(Item, "=")(1))), TargetTextBox, KeyCode, Shift)
            End Select
        Next Item
    End With
End Function

Private Function TryHandleTextBoxLateSpecials(ByRef TargetTextBox As Access.Control, ByRef KeyAscii As Integer) As Boolean
    ' TextBox SPECIAL PROPERTIES (LATE) HANDLING ON KeyPress EVENT
    Dim Item As Variant, Aux As String, isHandled As Boolean
    
    With TargetTextBox
        For Each Item In Split(Nz(.StatusBarText, ""), ";")
            Select Case Split(Item, "=")(0)
                Case "ALLOWEDCHARS"
                    Aux = Split(Item, "=", 2)(1)
                    isHandled = isHandled _
                                Or (Not ( _
                                    InStr(Aux, VBA.Chr$(KeyAscii)) > 0 _
                                    Or (KeyAscii = 8))) ' BACKSPACE KEY
                                    
                Case "APPLY"
                    Select Case Split(Item, "=")(1)
                        Case "NO_UNSAFE_CHARS"
                            Select Case KeyAscii
                                Case 34, 39, 59, 92 ' DISCARD: " ' ; \
                                    isHandled = True
                            End Select
                        Case "UCASE"
                            If KeyAscii > 32 Then
                                KeyAscii = VBA.AscW(UCase(VBA.Chr$(KeyAscii)))
                            End If
                        Case "LCASE"
                            If KeyAscii > 32 Then
                                KeyAscii = VBA.AscW(LCase(VBA.Chr$(KeyAscii)))
                            End If
                        Case "SINGLE_LINE"
                            Select Case KeyAscii
                                Case 10, 13
                                    isHandled = True
                            End Select
                    End Select
                    
                Case "REGEX"
                    isHandled = Not (TryValidateRegexPattern(VBA.Chr$(KeyAscii), Split(Item, "=", 2)(1)) _
                                    Or (KeyAscii = 8)) ' BACKSPACE KEY
                
            End Select
        Next Item
    End With
    
    TryHandleTextBoxLateSpecials = isHandled
End Function

Private Function TryValidateRegexPattern(ByVal mValue As String, ByVal mPattern As String) As Boolean
    Dim rExp As Object: Set rExp = CreateObject("VBSCript.RegExp")
    On Error GoTo Finally
    
    rExp.IgnoreCase = False
    rExp.Global = True
    rExp.Pattern = mPattern ' "^([a-zA-Z0-9_\-\.]+)@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,3})$"
    TryValidateRegexPattern = rExp.test(mValue)
    
    Exit Function
Finally:
    If DEBUG_MODE_ENABLED Then Stop
End Function

Public Function TryHandleKeyPress(ByRef KeyAscii As Integer) As Boolean
    'Debug.Print Printf("IN TryHandleKeyPress FOR %1", KeyAscii)
    On Error GoTo Finally
    
    Select Case Screen.ActiveControl.ControlType
        Case acTextBox, acComboBox
            If TryHandleTextBoxLateSpecials(Screen.ActiveControl, KeyAscii) Then
                ' isHandled
                TryHandleKeyPress = True
                KeyAscii = 0
            End If
    End Select
Finally:
End Function


Private Function GetTextUnsafe(ByRef mTextBox As Access.Control) As String
    On Error Resume Next
    GetTextUnsafe = mTextBox.Value
    GetTextUnsafe = mTextBox.Text
End Function

Private Function GetTextBoxLength(ByRef mTextBox As Access.Control) As Long
    On Error Resume Next
    GetTextBoxLength = Len(mTextBox.Value)
    GetTextBoxLength = Len(mTextBox.Text)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''


Public Function AutoExpandDropdownX(ByRef mCombo As Access.ComboBox, ByVal X As Single)
    On Error GoTo Finally
    Dim mControl As Access.Control
    Set mControl = mCombo
    
    If IsControlLocked(mControl) Then Exit Function
    
    If X <= (mCombo.Width - 265) Then
        If Len(Nz(mCombo.Text, "")) > 0 And mCombo.SelLength = 0 Then Exit Function
        mCombo.Dropdown
    End If
    
    Exit Function
Finally:
    Debug.Print "ERROR COUGHT!"
End Function

Public Function AutoExpandDropdown()
    On Error GoTo Finally

    If IsControlLocked(Screen.ActiveControl) Then Exit Function
    
    Select Case GetActiveControlType
        Case "ComboBox"
            Screen.ActiveControl.Dropdown
        Case "TextBox"
            If Screen.ActiveControl.ShowDatePicker = 1 And UCase(Right(Screen.ActiveControl.Format, 4)) = "DATE" Then
                SendKeys "%{DOWN}", False
            End If
    End Select

Finally:
End Function

Private Function IsControlLocked(ByRef TargetControl As Access.Control) As Boolean
    On Error Resume Next
    
    Select Case True
        Case TargetControl.Locked
            IsControlLocked = True
            Exit Function
        Case (Not TargetControl.Enabled)
            IsControlLocked = True
            Exit Function
        Case (Not GetParentFormOfControl(TargetControl).AllowEdits)
            IsControlLocked = True
            Exit Function
    End Select
End Function





