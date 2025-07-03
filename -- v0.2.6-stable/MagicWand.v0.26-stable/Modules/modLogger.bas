Attribute VB_Name = "modLogger"
' #############################################
' ## modLogger – CSV logging per machine/user
' #############################################

Option Explicit

Public Sub LogAction(actionType As String, folderPath As String, includeSubfolders As Boolean, _
                     exportPDF As Boolean, exportPDFType As String, altPDFPath As String, _
                     keepOriginal As Boolean, files As Long, replacements As Long, _
                     pdfs As Long, duration As Long, notes As String)

    Dim logFolder As String, logPath As String
    Dim fso As Object
    Dim fnum As Integer
    Dim machineID As String, userID As String
    Dim header As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    logFolder = ThisDocument.path & "\logs"
    If Not fso.FolderExists(logFolder) Then
        fso.CreateFolder logFolder
        SetAttr logFolder, vbHidden
    End If

    machineID = Environ("COMPUTERNAME")
    If machineID = "" Then
        machineID = Environ("USERNAME")
        If machineID = "" Then machineID = "UnknownMachine"
    End If
    logPath = logFolder & "\" & machineID & ".csv"

    userID = Environ("USERNAME")
    If userID = "" Then userID = "UnknownUser"

    fnum = FreeFile
    If Not fso.FileExists(logPath) Then
        Open logPath For Output As #fnum
        header = "Timestamp;User;Version;ActionType;FolderPath;Subfolders;PDFExport;PDFType;" & _
                 "AltPDFPathUsed;PreserveOriginals;FilesProcessed;ReplacementsMade;PDFsGenerated;" & _
                 "DurationSeconds;Notes"
        Print #fnum, header
    Else
        Open logPath For Append As #fnum
    End If

    Print #fnum, Format(Now, "yyyy-mm-dd HH:nn:ss") & ";" & _
                  userID & ";" & _
                  APP_VERSION & ";" & _
                  actionType & ";" & _
                  folderPath & ";" & _
                  CStr(includeSubfolders) & ";" & _
                  CStr(exportPDF) & ";" & _
                  exportPDFType & ";" & _
                  CStr(altPDFPath <> "") & ";" & _
                  CStr(keepOriginal) & ";" & _
                  files & ";" & _
                  replacements & ";" & _
                  pdfs & ";" & _
                  duration & ";" & _
                  notes
    Close #fnum
End Sub


' === Returns the Windows username as user ID (not email) ===
Public Function GetUserID() As String
    On Error Resume Next
    GetUserID = Environ("USERNAME")
    If GetUserID = "" Then GetUserID = "UnknownUser"
End Function




