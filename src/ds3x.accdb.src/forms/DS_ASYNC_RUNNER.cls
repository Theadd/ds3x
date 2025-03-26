Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "ds3x.UI.Misc"
Option Compare Database
Option Explicit


Public Sub RunAsync()
    Debug.Print "[INFO] DS_ASYNC_RUNNER.RunAsync"
    Me.TimerInterval = 100
End Sub

Private Sub Form_Timer()
    On Error Resume Next
    Me.TimerInterval = 0
    Run
    DoCmd.Close acForm, "DS_ASYNC_RUNNER", acSaveNo
    On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Me.TimerInterval = 0
    On Error GoTo 0
End Sub

Private Sub Run()
    Debug.Print "[INFO] DS_ASYNC_RUNNER.Run = " & CStr(dsApp.RunAll())
End Sub
