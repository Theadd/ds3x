﻿Attribute VB_Name = "FileSystemLib"
Option Compare Database
Option Explicit


'
' @REQUIRES:
'   1. A reference to "Microsoft Office 16.0 Object Library"
'

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As LongPtr)
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private pFSO As Scripting.FileSystemObject


' --- PUBLIC PROPERTIES ---

Public Property Get FSO() As Scripting.FileSystemObject
    If pFSO Is Nothing Then Set pFSO = New Scripting.FileSystemObject
    Set FSO = pFSO
End Property

Public Property Get PathSeparator() As String: PathSeparator = "\": End Property


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
Public Function TryWriteAccessOfSaveAsDialog(ByRef path As String, Optional ByVal saveAsFileExtension As String = "*.xlsx", Optional ByVal windowTitle As String = "Save As", Optional ByVal retryAttempts As Integer = 100) As Boolean
    Dim success As Boolean: success = TrySaveAsDialog(path, saveAsFileExtension, windowTitle)
    
    If success Then
        success = TryWaitFileWriteAccess(path, retryAttempts)
    End If
    
    TryWriteAccessOfSaveAsDialog = success
End Function


Public Function TrySaveAsDialog(ByRef path As String, Optional ByVal saveAsFileExtension As String = "*.xlsx", Optional ByVal windowTitle As String = "Save As") As Boolean
    Dim selectedItem As Variant

    With Application.FileDialog(msoFileDialogSaveAs)
        'Setup prefered view style
        .InitialView = msoFileDialogViewDetails
        .AllowMultiSelect = False
        'Setup default filename, it can contain a initial path too
        .InitialFileName = PathCombine(path, saveAsFileExtension)
        .Title = windowTitle

        If .Show Then
            'Step through eachString in the FileDialogSelectedItems collection.
            For Each selectedItem In .selectedItems
                'selectedItem is aString that contains the path of each selected item.
                'Use any file I/O functions that you want to work with this path.
                'This example displays the path in a message box.
                path = "" & selectedItem
            Next selectedItem
            TrySaveAsDialog = (Len(path) > 0)
        End If
    End With
End Function

Public Function TryFileOpenDialog(ByRef path As String, Optional ByVal initialPath As String = "", Optional ByVal commaSeparatedStringFiltersPairs As String = "All files,*.*,Excel files,*.xlsx;*.xls;*.xlsm", Optional ByVal windowTitle As String = "Open") As Boolean
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
                path = "" & selectedItem
            Next selectedItem
            TryFileOpenDialog = (Len(path) > 0)
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
                FSO.deletefile TargetPath, True
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
    
    Set stream = FSO.CreateTextFile(TargetFile, overwriteIfAlreadyExists, asUnicode)
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
    
    Set stream = FSO.OpenTextFile(TargetFile, ForReading, False, IIf(asUnicode, TristateTrue, TristateFalse))
    Content = stream.ReadAll
    stream.Close
    
    TryReadAllTextInFile = True
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    TryReadAllTextInFile = False
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





