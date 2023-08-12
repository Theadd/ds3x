Attribute VB_Name = "dsCommonLib"
Option Compare Database
Option Explicit

Private Const SC_GRAYED_COLOR As Long = 15461355
Private Const SC_GRAYED_SHADE As Long = 92
Private Const SC_GRAYED_THEME As Long = 1
Private Const SC_GRAYED_TINT As Long = 100

Private Const SC_WHITE_COLOR As Long = 16777215
Private Const SC_WHITE_SHADE As Long = 100
Private Const SC_WHITE_THEME As Long = 1
Private Const SC_WHITE_TINT As Long = 100

Private Const SC_INVALID_COLOR As Long = 2366701


' --- Memory Status ---

Private Declare PtrSafe Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

#If Win64 Then
    Declare PtrSafe Sub GlobalMemoryStatusEx Lib "kernel32" (lpBuffer As MEMORYSTATUS)

    Private Type LARGE_INTEGER
        LowPart As Long
        HighPart As Long
    End Type

    Public Type MEMORYSTATUS
       dwLength As Long
       dwMemoryLoad As Long
       dwTotalPhys As LARGE_INTEGER
       dwAvailPhys As LARGE_INTEGER
       dwTotalPageFile As LARGE_INTEGER
       dwAvailPageFile As LARGE_INTEGER
       dwTotalVirtual As LARGE_INTEGER
       dwAvailVirtual As LARGE_INTEGER
       dwAvailExtendedVirtual As LARGE_INTEGER
    End Type
    
    Public Function GetAvailableVirtualMemory() As Long
        On Error GoTo Finally
        Dim Mem As MEMORYSTATUS
    
        Mem.dwLength = LenB(Mem)
        GlobalMemoryStatusEx Mem
    
        GetAvailableVirtualMemory = BytesToMB(Mem.dwAvailVirtual)
Finally:
    End Function
    
    Private Function LargeIntToCurrency(liInput As LARGE_INTEGER) As Currency
        CopyMemory LargeIntToCurrency, liInput, LenB(liInput)
        LargeIntToCurrency = LargeIntToCurrency * 10000
    End Function

    Private Function BytesToMB(RawValue As LARGE_INTEGER) As Long
        Dim Value As Currency
        Value = LargeIntToCurrency(RawValue)
        Select Case Value
            Case Is > (2 ^ 20)
                BytesToMB = CLng(CStr(Round(Value / (2 ^ 20), 2)))
            Case Else
                BytesToMB = 0
        End Select
    End Function
#Else
    Declare PtrSafe Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
    
    Public Type MEMORYSTATUS
       dwLength As Long
       dwMemoryLoad As Long
       dwTotalPhys As Long
       dwAvailPhys As Long
       dwTotalPageFile As Long
       dwAvailPageFile As Long
       dwTotalVirtual As Long
       dwAvailVirtual As Long
    End Type
    
    Public Function GetAvailableVirtualMemory() As Long
        On Error GoTo Finally
        Dim Mem As MEMORYSTATUS
    
        Mem.dwLength = LenB(Mem)
        GlobalMemoryStatus Mem
    
        GetAvailableVirtualMemory = BytesToMB(Mem.dwAvailVirtual)
Finally:
    End Function
    
    Private Function BytesToMB(Value As Long) As Long
        Select Case Value
            Case Is > (2 ^ 20)
                BytesToMB = CLng(CStr(Round(Value / (2 ^ 20), 2)))
            Case Else
                BytesToMB = 0
        End Select
    End Function
#End If

' ---
' TODO: REMOVE
Public Sub ShowMemStats()
    Dim Mem As MEMORYSTATUS
    Mem.dwLength = LenB(Mem)
#If Win64 Then
    GlobalMemoryStatusEx Mem
#Else
    GlobalMemoryStatus Mem
#End If

    Debug.Print "Memory load:", , Mem.dwMemoryLoad; "%"
    Debug.Print
    Debug.Print "Total physical memory:", BytesToMB(Mem.dwTotalPhys)
    Debug.Print "Physical memory free: ", BytesToMB(Mem.dwAvailPhys)
    Debug.Print
    Debug.Print "Total paging file:", BytesToMB(Mem.dwTotalPageFile)
    Debug.Print "Paging file  free: ", BytesToMB(Mem.dwAvailPageFile)
    Debug.Print
    Debug.Print "Total virtual memory:", BytesToMB(Mem.dwTotalVirtual)
    Debug.Print "Virtual memory free: ", BytesToMB(Mem.dwAvailVirtual)
#If Win64 Then
    Debug.Print "Virtual memory free: ", BytesToMB(Mem.dwAvailExtendedVirtual)
#End If

End Sub
' ---

' --- Utility Functions ---

Public Function SecondsToHMS(ByVal Value As Long) As String
    Dim hrs As Long, mins, secs, m As Integer, t As String
    On Error GoTo Finally
    
    hrs = Fix(Value / 3600)
    mins = Fix(Value / 60) Mod 60
    secs = Fix((Value Mod 60) / 1)
    
    If hrs >= 2 Or (hrs = 1 And mins > 39) Then
        t = t & " " & hrs & "h"
        If mins >= 5 Then t = t & " " & mins & "m"
    Else
        If (hrs = 1 And mins <= 39) Then mins = mins + 60
        If mins >= 1 Then
            t = t & " " & mins & "m"
            If mins < 15 Then t = t & " " & secs & "s"
        Else
            t = t & " " & secs & "s"
        End If
    End If
    
    SecondsToHMS = VBA.Mid(t, 2)
Finally:
End Function

' USAGES: XDaysAgo(15), XDaysAgo("7 days..."), XDaysAgo("-7 in a week")
Public Function XDaysAgo(ByVal Value As Variant, Optional ByVal DateFormat As String = "dd MMM") As String
    If VBA.InStr(1, Value, CStr(Val(Value))) > 0 Then
        Value = UCase(VBA.Format$(DateAdd("d", 0 - Int(Val(Value)), Date), DateFormat, vbMonday))
    End If
    XDaysAgo = CStr(Value)
End Function

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


' --- System Clipboard ---

' Read/Write to Clipboard. Source: ExcelHero.com (Daniel Ferry)
Public Function SystemClipboard(Optional StoreText As String) As String
    Dim X As Variant: X = StoreText ' 64-bit support

    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            Select Case True
                Case Len(StoreText): .SetData "text", X
                Case Else: SystemClipboard = .GetData("text")
            End Select
        End With
    End With
End Function


' --- StyleLib ---

Private Function SetGrayedStyle(ByRef TargetControl As Access.Control, ByVal asGrayed As Boolean)
    Dim TargetLabel As Access.Label, styleLabel As Boolean

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

Public Function SetControlAsEnabled(ByRef TargetControl As Access.Control, Optional ByVal Enable As Boolean = True) As Boolean
    SetControlAsEnabled = Enable
    SetGrayedStyle TargetControl, (Not Enable)
    On Error GoTo Fallback
    TargetControl.Locked = (Not Enable)
    Exit Function
Fallback:
    TargetControl.Enabled = Enable
End Function

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





