VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSystemLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "ds3x.Libraries"
Option Compare Database
Option Explicit


'
' @REQUIRES:
'   1. A reference to "Microsoft Office 16.0 Object Library"
'

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As LongPtr)
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare PtrSafe Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

#If Win64 Then
    Private Declare PtrSafe Sub GlobalMemoryStatusEx Lib "kernel32" (lpBuffer As MEMORYSTATUS)

    Private Type LARGE_INTEGER
        LowPart As Long
        HighPart As Long
    End Type

    Private Type MEMORYSTATUS
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
    Private Declare PtrSafe Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
    
    Private Type MEMORYSTATUS
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
    
Private pFSO As Scripting.FileSystemObject


' --- PUBLIC PROPERTIES ---

Public Property Get FSO() As Scripting.FileSystemObject
    If pFSO Is Nothing Then Set pFSO = New Scripting.FileSystemObject
    Set FSO = pFSO
End Property

Public Property Get PathSeparator() As String: PathSeparator = "\": End Property
Public Property Get Resolve(ByVal Path As String, Optional ByVal RelativeTo As Variant) As String: Resolve = PathResolve(Path, RelativeTo): End Property
Public Property Get ReadAllTextInFile(ByVal TargetFile As String, Optional ByVal asUnicode As Boolean = True) As String: Dim s As String: FileSystemLib.TryReadAllTextInFile TargetFile, s, asUnicode: ReadAllTextInFile = s: End Property
Public Property Get FileInAncestors(ByVal TargetFile As String, Optional ByVal BackwardMovesLimit As Long = -1) As String: FileInAncestors = IIf(TryGetFileInAncestors(TargetFile, BackwardMovesLimit), TargetFile, ""): End Property


' --- FILE SYSTEM UTILITIES ---

' Function TryWriteAccessOfSaveAsDialog(path, [saveAsFileExtension="*.xlsx"], [windowTitle="Save As"], [retryAttempts=100]) As Boolean
'
' Abre un cuadro de diálogo de "Guardar como..." y se asegura de poder acceder en modo de
' escritura al archivo proporcionado por el usuario.
'
'   @param path - a string passed by reference used to provide optional path to a directory by
'                 default that will be replaced by the full path of the file selected by the user.
'
'   @param saveAsFileExtension - [Optional] allowed file extension (using wilcards).
'                                           Defaults to "*.xlsx".
'
'   @param windowTitle - [Optional] Defaults to "Save As".
'
'   @param retryAttempts - [Optional] Máximo número de intentos (100ms) que esperará hasta tener
'                                     acceso al archivo.
'
Public Function TryWriteAccessOfSaveAsDialog(ByRef Path As String, Optional ByVal saveAsFileExtension As String = "*.xlsx", Optional ByVal windowTitle As String = "Save As", Optional ByVal retryAttempts As Integer = 100) As Boolean
    Dim success As Boolean: success = TrySaveAsDialog(Path, saveAsFileExtension, windowTitle)
    
    If success Then
        success = TryWaitFileWriteAccess(Path, retryAttempts)
    End If
    
    TryWriteAccessOfSaveAsDialog = success
End Function


Public Function TrySaveAsDialog(ByRef Path As String, Optional ByVal saveAsFileExtension As String = "*.xlsx", Optional ByVal windowTitle As String = "Save As") As Boolean
    Dim selectedItem As Variant

    With Application.FileDialog(msoFileDialogSaveAs)
        'Setup prefered view style
        .InitialView = msoFileDialogViewDetails
        .AllowMultiSelect = False
        'Setup default filename, it can contain a initial path too
        .InitialFileName = PathCombine(Path, saveAsFileExtension)
        .Title = windowTitle

        If .Show Then
            'Step through eachString in the FileDialogSelectedItems collection.
            For Each selectedItem In .selectedItems
                'selectedItem is aString that contains the path of each selected item.
                'Use any file I/O functions that you want to work with this path.
                'This example displays the path in a message box.
                Path = "" & selectedItem
            Next selectedItem
            TrySaveAsDialog = (Len(Path) > 0)
        End If
    End With
End Function

Public Function TryFileOpenDialog(ByRef Path As String, Optional ByVal initialPath As String = "", Optional ByVal commaSeparatedStringFiltersPairs As String = "All files,*.*,Excel files,*.xlsx;*.xls;*.xlsm", Optional ByVal windowTitle As String = "Open") As Boolean
    Dim selectedItem As Variant, i As Integer, lastIndex As Integer, filtersArray() As String

    With Application.FileDialog(msoFileDialogOpen)
        'Setup prefered view style
        .InitialView = msoFileDialogViewDetails
        .AllowMultiSelect = False
        'Setup default filename, it can contain a initial path too
        .InitialFileName = initialPath
        .Title = windowTitle
        ' Rebuild filters
        .Filters.Clear
        filtersArray = Split(commaSeparatedStringFiltersPairs, ",")
        lastIndex = UBound(filtersArray)
        
        For i = 0 To lastIndex Step 2
            .Filters.Add filtersArray(i), filtersArray(i + 1)
        Next i
        
        If .Show Then
            'Step through eachString in the FileDialogSelectedItems collection.
            For Each selectedItem In .selectedItems
                'selectedItem is aString that contains the path of each selected item.
                'Use any file I/O functions that you want to work with this path.
                'This example displays the path in a message box.
                Path = "" & selectedItem
            Next selectedItem
            TryFileOpenDialog = (Len(Path) > 0)
        End If
    End With
End Function

Public Function TryMultiSelectFileOpenDialog(ByRef paths As Collection, Optional ByVal initialPath As String = "", Optional ByVal commaSeparatedStringFiltersPairs As String = "All files,*.*,Excel files,*.xlsx;*.xls;*.xlsm,CSV files,*.csv", Optional ByVal windowTitle As String = "Open") As Boolean
    Dim selectedItem As Variant, i As Integer, lastIndex As Integer, filtersArray() As String

    With Application.FileDialog(msoFileDialogOpen)
        'Setup prefered view style
        .InitialView = msoFileDialogViewDetails
        .AllowMultiSelect = True
        'Setup default filename, it can contain a initial path too
        .InitialFileName = initialPath
        .Title = windowTitle
        ' Rebuild filters
        .Filters.Clear
        filtersArray = Split(commaSeparatedStringFiltersPairs, ",")
        lastIndex = UBound(filtersArray)
        
        For i = 0 To lastIndex Step 2
            .Filters.Add filtersArray(i), filtersArray(i + 1)
        Next i
        
        If .Show Then
            'Step through eachString in the FileDialogSelectedItems collection.
            For Each selectedItem In .selectedItems
                paths.Add "" & selectedItem
                If Not TryMultiSelectFileOpenDialog Then
                    TryMultiSelectFileOpenDialog = (Len("" & selectedItem) > 0)
                End If
            Next selectedItem
        End If
    End With
End Function

' Devuelve el resultado normalizado de combinar la ruta de un directorio con un archivo en ese directorio.
Public Function PathCombine(ByVal directoryPath As String, ByVal filename As String) As String
    PathCombine = FSO.BuildPath(directoryPath, filename)
End Function

' Si no existe el directorio en `targetPath` lo crea y devuelve true, si `targetPath` corresponde a un archivo
' o se produce un error, devuelve false.
Public Function TryCreateFolder(ByVal TargetPath As String) As Boolean
    On Error GoTo CreateFolderError
    'If the path exists as a file, the function fails.
    If FSO.FileExists(TargetPath) Then
        TryCreateFolder = False
        Exit Function
    End If
    
    'If the path already exists as a folder, don't do anything and return success.
    If FSO.FolderExists(TargetPath) Then
        TryCreateFolder = True
        Exit Function
    End If
    
    TryCreateFolder = (Not (FSO.CreateFolder(TargetPath) Is Nothing))
CreateFolderError:
End Function

' Devuelve si existe o no el directorio especificado, esperando 0.1 segundos a cada intento fallido,
' @SEE: FileSystemLib.TryWaitFileExists
Public Function TryWaitFolderExists(ByVal TargetPath As String, Optional ByVal retryAttempts As Integer = 100) As Boolean
    On Error GoTo Finally
    Dim Exists As Boolean
    
    'If the path exists as a file, it can't be a directory
    If FSO.FileExists(TargetPath) Then Exit Function
    
    Exists = FSO.FolderExists(TargetPath)
    While Not Exists And retryAttempts > 0
        retryAttempts = retryAttempts - 1
        Sleep 100
        Exists = FSO.FolderExists(TargetPath)
    Wend
    
    TryWaitFolderExists = Exists
Finally:
End Function

' Devuelve si existe o no el archivo especificado, esperando 0.1 segundos a cada intento fallido,
' si no estás creando dicho archivo en este momento, no hay razón para repetir ningún intento,
' pásale un 0 al segundo parámetro (retryAttempts).
Public Function TryWaitFileExists(ByVal TargetPath As String, Optional ByVal retryAttempts As Integer = 100) As Boolean
    Dim Exists As Boolean
    
    Exists = FSO.FileExists(TargetPath)
    
    While Not Exists And retryAttempts > 0
        retryAttempts = retryAttempts - 1
        Sleep 100
        Exists = FSO.FileExists(TargetPath)
    Wend
    
    TryWaitFileExists = Exists
End Function

Public Function TryWaitFileMatchingPatternExists(ByVal directoryPath As String, Optional ByVal filenamePattern As String = "*.csv", Optional ByVal retryAttempts As Integer = 100) As Boolean
    Dim Exists As Boolean, TargetPath As String
    
    TargetPath = FSO.BuildPath(directoryPath, filenamePattern)
    Exists = (Dir(TargetPath) <> "")
    
    While Not Exists And retryAttempts > 0
        retryAttempts = retryAttempts - 1
        Sleep 100
        Exists = (Dir(TargetPath) <> "")
    Wend
    
    TryWaitFileMatchingPatternExists = Exists
End Function

' Al escribir en una unidad de red y una vez el archivo en questión exista, se esperará, un número máximo
' de intentos, hasta que el archivo sea liberado (escrito por completo y cerrado).
' Si el archivo especificado no existe, se asegurará de que podamos crear ese fichero, de lo contrario devolverá false.
Public Function TryWaitFileWriteAccess(ByVal TargetPath As String, Optional ByVal retryAttempts As Integer = 100) As Boolean
    Dim stream As TextStream
    TryWaitFileWriteAccess = False

    If TryWaitFileExists(TargetPath, 0) Then
        If TryWaitFileExists(TargetPath & "~", 0) Then
            
            Debug.Print "Error no contemplado, ya existe un archivo con el mismo nombre y acabado con '~' en el mismo directorio. " & TargetPath
            GoTo ExitFunction
        Else
            While True
            
                If TryMoveFile(TargetPath, TargetPath & "~") Then
                    GoTo MoveFileBack
                End If

                If retryAttempts = 0 Then GoTo ExitFunction
                
                Sleep 100
                retryAttempts = retryAttempts - 1
                
            Wend

MoveFileBack:
            On Error GoTo HandleUnexpectedError
            FSO.MoveFile TargetPath & "~", TargetPath
            
            If TryWaitFileExists(TargetPath, retryAttempts) Then
                TryWaitFileWriteAccess = True
            End If
        End If
    Else
        ' There's no existing file at targetPath, ensure write access
        Set stream = FSO.CreateTextFile(TargetPath, overwrite:=False)
        stream.Close
        
        If TryWaitFileExists(TargetPath, retryAttempts) Then
            
            If TryWaitFileWriteAccess(TargetPath, retryAttempts) Then
                FSO.DeleteFile TargetPath, True
                TryWaitFileWriteAccess = True
            End If
            
        End If
    End If

ExitFunction:
    Exit Function
HandleUnexpectedError:
    TryWaitFileWriteAccess = False
    Debug.Print "Error inesperado al devolver un archivo a su nombre original. " & TargetPath
End Function

Public Function GetFileMatchingPattern(ByVal Pattern As String) As String
    Dim sPath As String, sName As String
    On Error GoTo Finally
    
    sPath = FSO.GetParentFolderName(Pattern)
    sName = VBA.Dir(Pattern)
    If sName <> "" Then
        GetFileMatchingPattern = FSO.BuildPath(sPath, sName)
    End If
Finally:
End Function

Public Function TryMoveFile(ByVal sourceFilePath As String, ByVal destinationFilePath As String) As Boolean
    On Error GoTo MoveFail
    TryMoveFile = True
    
    FSO.MoveFile sourceFilePath, destinationFilePath
    
    Exit Function
MoveFail:
    TryMoveFile = False
End Function

Public Function TryCopyFile(ByVal sourceFilePath As String, ByVal destinationFilePath As String, Optional ByVal overwriteExisting As Boolean = True) As Boolean
    On Error GoTo CopyFail
    TryCopyFile = True
    
    FSO.CopyFile sourceFilePath, destinationFilePath, overwriteExisting
    
    Exit Function
CopyFail:
    TryCopyFile = False
End Function

Public Function TryFollowHyperlink(ByVal TargetFile As String) As Boolean
    On Error GoTo Finally
    
    Application.FollowHyperlink TargetFile
    TryFollowHyperlink = True
    Exit Function
Finally:
End Function

Public Function TryKill(ByVal TargetFile As String) As Boolean
    On Error GoTo Finally
    
    Kill TargetFile
    TryKill = True
    Exit Function
Finally:
End Function

Public Function TryWaitFileKill(ByRef TargetFile As String, Optional ByVal retryAttempts As Integer = 100) As Boolean
    If TryWaitFileExists(TargetFile, 0) Then
        If TryWaitFileWriteAccess(TargetFile, retryAttempts) Then
            TryWaitFileKill = (TryKill(TargetFile))
        End If
    Else
        TryWaitFileKill = True  ' TargetFile does not even exist, no need to kill.
    End If
End Function


' Devuelve si existe o no un archivo en el directorio especificado sin importar su extensión y asigna
' la ruta completa de ese archivo al tercer parametro, targetFilePath, que es pasado por referencia
Public Function TryGetFileWithoutExtension(ByVal targetDirectory As String, ByVal filenameWithoutExtension As String, ByRef targetFilePath As String) As Boolean
    Dim StrFile As String, found As Boolean
    targetFilePath = PathCombine(targetDirectory, filenameWithoutExtension)
    TryGetFileWithoutExtension = False
    found = False
    
    StrFile = Dir(targetFilePath & ".*")
    
    Do While Len(StrFile) > 0
        found = True
        Exit Do
    Loop
    
    If found Then
        targetFilePath = targetDirectory & StrFile
        TryGetFileWithoutExtension = True
    End If
    
End Function

Public Function GetFileExtension(ByVal FilePath As String) As String
    GetFileExtension = FSO.GetExtensionName(FilePath)
End Function

Public Function GetFileName(ByVal FilePath As String) As String
    GetFileName = FSO.GetFileName(FilePath)
End Function

Public Function TryGetFileInAncestors(ByRef TargetFile As String, Optional ByVal BackwardMovesLimit As Long = -1) As Boolean
    Dim sName As String, sPath As String
    sName = GetFileName(TargetFile)
    sPath = TargetFile
    
    Do
        sPath = FSO.GetParentFolderName(sPath)
        If sPath = vbNullString Then Exit Do
        
        If TryWaitFileExists(PathCombine(sPath, sName), 0) Then
            TargetFile = PathCombine(sPath, sName)
            TryGetFileInAncestors = True
            Exit Do
        End If
        
        BackwardMovesLimit = BackwardMovesLimit - 1
    Loop Until (BackwardMovesLimit = -1)
End Function


Public Function TryWriteTextToFile(ByVal TargetFile As String, ByRef Content As String, Optional ByVal overwriteIfAlreadyExists As Boolean = True, Optional ByVal asUnicode As Boolean = True) As Boolean
    On Error GoTo ErrorHandler
    Dim stream As TextStream
    
    Set stream = FSO.CreateTextFile(FileSystemLib.Resolve(TargetFile), overwriteIfAlreadyExists, asUnicode)
    stream.Write (Content)
    stream.Close
    
    TryWriteTextToFile = True
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    TryWriteTextToFile = False
    If Not (stream Is Nothing) Then stream.Close
End Function

Public Function TryReadAllTextInFile(ByVal TargetFile As String, ByRef Content As String, Optional ByVal asUnicode As Boolean = True) As Boolean
    On Error GoTo ErrorHandler
    Dim stream As TextStream
    
    Set stream = FSO.OpenTextFile(FileSystemLib.Resolve(TargetFile), ForReading, False, IIf(asUnicode, TristateTrue, TristateFalse))
    Content = stream.ReadAll
    stream.Close
    
    TryReadAllTextInFile = True
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    TryReadAllTextInFile = False
    If Not (stream Is Nothing) Then stream.Close
End Function

Public Function TryAppendTextInFile(ByVal TargetFile As String, ByVal Content As String, Optional ByVal asUnicode As Boolean = True) As Boolean
    On Error GoTo ErrorHandler
    Dim stream As TextStream
    
    Set stream = FSO.OpenTextFile(FileSystemLib.Resolve(TargetFile), ForAppending, True, IIf(asUnicode, TristateTrue, TristateFalse))
    stream.Write Content
    stream.Close
    
    TryAppendTextInFile = True
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    TryAppendTextInFile = False
    If Not (stream Is Nothing) Then stream.Close
End Function

Public Function TryShellExecute(ByRef TargetFile As String, Optional ByVal ShowStyle As VbAppWinStyle = vbNormalFocus) As Boolean
    On Error GoTo Finally

    TryShellExecute = (ShellExecute(0&, vbNullString, TargetFile, vbNullString, vbNullString, ShowStyle) >= 32)
Finally:
End Function

Public Function GetAllFilesInPath(ByVal TargetPath As String, Optional ByVal FilePattern As String = "*") As Collection
    Dim allFiles As New Collection, aMatch As String
    
    aMatch = Dir(PathCombine(TargetPath, FilePattern))
    
    While aMatch <> ""
        allFiles.Add aMatch
        aMatch = Dir()
    Wend
    
    Set GetAllFilesInPath = allFiles
End Function

Private Function PathResolve(ByVal Path As String, Optional ByVal RelativeTo As Variant) As String
    PathResolve = FSO.GetAbsolutePathName(Path)
    If VBA.Left$(PathResolve, 2) <> VBA.UCase$(VBA.Left$(Path, 2)) Then
        If IsMissing(RelativeTo) Then RelativeTo = Application.CodeProject.Path
        PathResolve = FSO.GetAbsolutePathName(FSO.BuildPath(RelativeTo, Path))
    End If
End Function


' --- Memory Usage ---

#If Win64 Then

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


