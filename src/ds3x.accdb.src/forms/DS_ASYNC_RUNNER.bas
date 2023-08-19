Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =2777
    DatasheetFontHeight =11
    Left =3225
    Top =3030
    Right =28545
    Bottom =15225
    OnUnload ="[Event Procedure]"
    RecSrcDt = Begin
        0x09f591aa690ae640
    End
    DatasheetFontName ="Calibri"
    OnTimer ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Section
            Height =963
            Name ="Detalle"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Me.TimerInterval = 0
End Sub

Private Sub Run()
    Debug.Print "[INFO] DS_ASYNC_RUNNER.Run = " & CStr(dsApp.RunAll())
End Sub
