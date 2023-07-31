Attribute VB_Name = "SystemLib"
Option Compare Database
Option Explicit


''''''''''''''''''' CLIPBOARD '''''''''''''''''''

Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32" Alias "GetClipboardDataA" (ByVal wFormat As Long) As LongPtr

Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long

Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr

' NO EXACT REPLACE FOR: Private Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare PtrSafe Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As LongPtr

Private Const GHND = &H42
Private Const CF_TEXT = 1
Private Const MaxSize = 4096





' --- Enabling MouseWheel Scrolling in Multiline TextBox ---

Private Const WM_VSCROLL = &H115
Private Const SB_LINEUP = 0
Private Const SB_LINEDOWN = 1

Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
'Private Declare Function SendMessage Lib "user32" _
'   Alias "SendMessageA" _
'   (ByVal hWnd As Long, _
'   ByVal wMsg As Long, _
'   ByVal wParam As Long, _
'   LParam As Any) _
'   As Long

Private Declare PtrSafe Function apiGetFocus Lib "user32" Alias "GetFocus" () As LongPtr

' ---



'''''''''''''''''' UNICODE FRIENDLY MSGBOX REPLACEMENT '''''''''''''''''''''

#If VBA7 Then
    Private Declare PtrSafe Function MessageBoxW Lib "user32" _
                                    (ByVal hWnd As LongPtr, _
                                     ByVal lpText As LongPtr, _
                                     ByVal lpCaption As LongPtr, _
                                     ByVal wType As Long) As Long
#Else
    Private Declare Function MessageBoxW Lib "user32" _
                            (ByVal hWnd As Long, _
                             ByVal lpText As Long, _
                             ByVal lpCaption As Long, _
                             ByVal wType As Long) As Long
#End If


''''''''''''''''''' MEMORY STATS '''''''''''''''''

'Inspired by: https://stackoverflow.com/a/48626253/154439  (h/t Charles Williams)
#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
#Else
    Private Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
#End If


#If Win64 Then
    Declare PtrSafe Sub GlobalMemoryStatusEx Lib "kernel32" (lpBuffer As MEMORYSTATUS)

    Private Type LARGE_INTEGER
        LowPart As Long
        HighPart As Long
    End Type

    'GlobalMemoryStatusEx outputs memory sizes in 64-bit *un*-signed integers;
    '   LongLong won't give us correct values because it is a signed type;
    '   the workaround is to use a custom data type and convert the result to Currency,
    '   as Currency is a fixed-point numeric data type supporting large values
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

#Else
    Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
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
#End If












#If Win64 Then
'Convert raw 64-bit unsigned integers to Currency data type
Private Function LargeIntToCurrency(liInput As LARGE_INTEGER) As Currency
    'copy 8 bytes from the large integer to an empty currency
    CopyMemory LargeIntToCurrency, liInput, LenB(liInput)
    'adjust it
    LargeIntToCurrency = LargeIntToCurrency * 10000
End Function
#End If

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
    Debug.Print "Total physical memory:", BytesToXB(Mem.dwTotalPhys)
    Debug.Print "Physical memory free: ", BytesToXB(Mem.dwAvailPhys)
    Debug.Print
    Debug.Print "Total paging file:", BytesToXB(Mem.dwTotalPageFile)
    Debug.Print "Paging file  free: ", BytesToXB(Mem.dwAvailPageFile)
    Debug.Print
    Debug.Print "Total virtual memory:", BytesToXB(Mem.dwTotalVirtual)
    Debug.Print "Virtual memory free: ", BytesToXB(Mem.dwAvailVirtual)
#If Win64 Then
    Debug.Print "Virtual memory free: ", BytesToXB(Mem.dwAvailExtendedVirtual)
#End If

End Sub

Public Function TryGetMemoryUsage(ByRef mLoad As Long, ByRef mFreeVirtual As Long, ByRef mTotalVirtual As Long, ByRef mFreePhysical As Long, ByRef mTotalPhysical As Long) As Boolean
    On Error GoTo Finally
    Dim Mem As MEMORYSTATUS

    Mem.dwLength = LenB(Mem)
#If Win64 Then
    GlobalMemoryStatusEx Mem
#Else
    GlobalMemoryStatus Mem
#End If

    mLoad = Mem.dwMemoryLoad
    mFreeVirtual = BytesToMB(Mem.dwAvailVirtual)
    mTotalVirtual = BytesToMB(Mem.dwTotalVirtual)
    mFreePhysical = BytesToMB(Mem.dwAvailPhys)
    mTotalPhysical = BytesToMB(Mem.dwTotalPhys)

    TryGetMemoryUsage = True
    Exit Function

Finally:
    TryGetMemoryUsage = False
End Function

Public Sub ShowAvailableVirtualMemory()
    On Error GoTo Finally
    Dim Mem As MEMORYSTATUS

    Mem.dwLength = LenB(Mem)
#If Win64 Then
    GlobalMemoryStatusEx Mem
#Else
    GlobalMemoryStatus Mem
#End If

    Debug.Print CStr(BytesToMB(Mem.dwAvailVirtual)) & " MB / " & CStr(BytesToMB(Mem.dwTotalVirtual)) & " MB"
Finally:
End Sub

Public Function GetAvailableVirtualMemory() As Long
    On Error GoTo Finally
    Dim Mem As MEMORYSTATUS

    Mem.dwLength = LenB(Mem)
#If Win64 Then
    GlobalMemoryStatusEx Mem
#Else
    GlobalMemoryStatus Mem
#End If

    GetAvailableVirtualMemory = BytesToMB(Mem.dwAvailVirtual)
Finally:
End Function

'Convert raw byte count to a more human readable format
#If Win64 Then
Private Function BytesToXB(RawValue As LARGE_INTEGER) As String
    Dim Value As Currency
    Value = LargeIntToCurrency(RawValue)

#Else
Private Function BytesToXB(Value As Long) As String

#End If
    Select Case Value
    Case Is > (2 ^ 30)
        BytesToXB = Round(Value / (2 ^ 30), 2) & " GB"
    Case Is > (2 ^ 20)
        BytesToXB = Round(Value / (2 ^ 20), 2) & " MB"
    Case Is > (2 ^ 10)
        BytesToXB = Round(Value / (2 ^ 10), 2) & " KB"
    Case Else
        BytesToXB = Value & " B"
    End Select
End Function


#If Win64 Then
Private Function BytesToMB(RawValue As LARGE_INTEGER) As Long
    Dim Value As Currency
    Value = LargeIntToCurrency(RawValue)

#Else
Private Function BytesToMB(Value As Long) As Long

#End If
    Select Case Value
        Case Is > (2 ^ 20)
            BytesToMB = CLng(CStr(Round(Value / (2 ^ 20), 2)))
        Case Else
            BytesToMB = 0
    End Select
End Function

Public Function MveNF(ByVal Val As String, Optional ByVal n As Long = 13) As String
    Dim i As Long, res As String, lMax As Long, cCode As Long, cMod As Long
    lMax = Len(Val)
    res = ""
    
    For i = 1 To lMax
        cCode = VBA.AscW(VBA.Mid(Val, i, 1))
        cCode = cCode + n
        If cCode < 32 Then
            cCode = 127 - (32 - cCode)
        ElseIf cCode > 126 Then
            cCode = 31 + (cCode - 126)
        End If
        res = res & VBA.ChrW(cCode)
    Next i
    
    MveNF = res
End Function

Public Function MveNB(ByVal Val As String, Optional ByVal n As Long = 13) As String
    MveNB = MveNF(Val, n * (-1))
End Function


''''''''''''''''''''''' CLIPBOARD '''''''''''''''''''''

Public Function GetTextFromClipboard() As String
    GetTextFromClipboard = ClipBoard_GetData
End Function

Private Function ClipBoard_GetData() As String
   Dim hClipMemory As LongPtr
   Dim lpClipMemory As LongPtr
   Dim MyString As String
   Dim RetVal As LongPtr
 
   If OpenClipboard(0&) = 0 Then
      Debug.Print "Cannot open Clipboard. Another app. may have it open"
      Exit Function
   End If
          
   ' Obtain the handle to the global memory
   ' block that is referencing the text.
   hClipMemory = GetClipboardData(CF_TEXT)
   If IsNull(hClipMemory) Then
      Debug.Print "Could not allocate memory"
      GoTo OutOfHere
   End If
 
   ' Lock Clipboard memory so we can reference
   ' the actual data string.
   lpClipMemory = GlobalLock(hClipMemory)
 
   If Not IsNull(lpClipMemory) Then
      MyString = Space$(MaxSize)
      RetVal = lstrcpy(MyString, lpClipMemory)
      RetVal = GlobalUnlock(hClipMemory)
       
      ' Peel off the null terminating character.
      MyString = Mid(MyString, 1, InStr(1, MyString, Chr$(0), 0) - 1)
   Else
      Debug.Print "Could not lock memory to copy string from."
   End If
 
OutOfHere:
 
   RetVal = CloseClipboard()
   ClipBoard_GetData = MyString
 
End Function

Public Function AsLockedToken(fpath As String) As Boolean
    On Error GoTo Finally
    Dim fs As New Scripting.FileSystemObject, cbd As Long, Aux As String
    AsLockedToken = True
    
    With fs.OpenTextFile(fpath, ForReading, False, TristateFalse)
        cbd = DateDiff("d", Date, DateValue(MveNF(.ReadAll)))
        .Close
    End With
    AsLockedToken = Not (cbd >= 0 And cbd <= 5)
Finally:
End Function






'''''''''''''''''' UNICODE FRIENDLY MSGBOX REPLACEMENT '''''''''''''''''''''

' ----------------------------------------------------------------
' Procedure : MsgBox
' Author    : Mike Wolfe
' Date      : 2020-12-24
' Source    : https://nolongerset.com/unicode-msgbox-v2/
' Purpose   : Unicode-safe drop-in replacement for the VBA MsgBox function.
' ----------------------------------------------------------------
Function MsgBox(Prompt As Variant, _
                Optional Buttons As VbMsgBoxStyle = vbOKOnly, _
                Optional Title As Variant) As VbMsgBoxResult
    
    'Set the default MsgBox title to Microsoft Access/Excel/Word/etc.
    Dim Caption As String
    If IsMissing(Title) Then
        Caption = Application.Name
    Else
        Caption = Title
    End If
    
#If VBA7 Then
    Dim hWnd As LongPtr
#Else
    Dim hWnd As Long
#End If
    
    ' Avoiding compile errors by making the Application object late-bound
    Dim mApp As Object
    Set mApp = Application
    
    hWnd = mApp.hWndAccessApp

    MsgBox = MessageBoxW(hWnd, StrPtr(Prompt), StrPtr(Caption), Buttons)
End Function






' --- Enabling MouseWheel Scrolling in Multiline TextBox ---

' Scroll multi-line textboxes with the mouse wheel. The textbox must have the focus.
'
' Call this sub in the MouseWheel event of the form(s) containing multi-line textboxes, like this:
'
' Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)
'    Call MouseWheelScroll(Count)
' End Sub
'
' Sources
' http://www.access-programmers.co.uk/forums/showthread.php?t=195679
' http://www.extramiledata.com/scroll-microsoft-access-text-box-using-mouse-wheel/

Public Sub MouseWheelScroll(ByVal Count As Long)

    Dim LinesToScroll As Integer
    Dim hwndActiveControl As LongPtr
    
    If Screen.ActiveControl.Properties.Item("ControlType") = acTextBox Then
        hwndActiveControl = fhWnd(Screen.ActiveControl)
        For LinesToScroll = 1 To Abs(Count)
            SendMessage hwndActiveControl, WM_VSCROLL, IIf(Count < 0, SB_LINEUP, SB_LINEDOWN), 0&
        Next
    End If

End Sub

' Source: http://access.mvps.org/access/api/api0027.htm
' Code Courtesy of Dev Ashish

Private Function fhWnd(ctl As Control) As LongPtr
    
    On Error Resume Next
    ' We only use this function for Screen.ActiveControl, so this is not necessary.
    ' I can't remember if I found it harmful in some situations.
    ' ctl.SetFocus
    
    fhWnd = apiGetFocus
    On Error GoTo 0
    
End Function

' ---


' --- VBE DEVTOOLS ---

Public Sub SwitchToPreviousActiveCodeWindow()
    If DEBUG_MODE_ENABLED Then
        Debug.Print "VBE Window visible = " & CStr(Application.VBE.MainWindow.Visible)
        Application.VBE.MainWindow.Collection(2).SetFocus
        
    End If
End Sub


