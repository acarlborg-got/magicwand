Attribute VB_Name = "modStats"
' #############################################
' ## modStats – Log summary & productivity
' #############################################

Option Explicit

Public Sub ShowMyEfficiency()
    Dim logPath As String
    logPath = ThisDocument.path & "\logs\" & Environ("COMPUTERNAME") & ".csv"

    If Dir(logPath) = "" Then
        MsgBox "No log file found for this user.", vbExclamation
        Exit Sub
    End If

    Dim fnum As Integer, line As String, totalSeconds As Double
    Dim parts() As String
    Dim files As Long, replaces As Long, pdfs As Long, estSeconds As Double

    fnum = FreeFile
    Open logPath For Input As #fnum
    Line Input #fnum, line ' skip header

    Do While Not EOF(fnum)
        Line Input #fnum, line
        parts = Split(line, ";")
        If UBound(parts) >= 13 Then
            estSeconds = EstimateTimeSaved(parts(3), CLng(parts(10)), CLng(parts(11)), CLng(parts(12)))
            totalSeconds = totalSeconds + estSeconds
            files = files + CLng(parts(10))
            replaces = replaces + CLng(parts(11))
            pdfs = pdfs + CLng(parts(12))
        End If
    Loop
    Close #fnum

    MsgBox "Your estimated time saved:" & vbCrLf & _
           FormatTime(totalSeconds) & vbCrLf & _
           "Files: " & files & ", Replacements: " & replaces & ", PDFs: " & pdfs, vbInformation
End Sub


Public Sub ShowTeamEfficiency()
    Dim logFolder As String
    logFolder = ThisDocument.path & "\logs\"

    Dim f As Object, fso As Object, file As Object
    Dim totalSeconds As Double, totalFiles As Long, totalReps As Long, totalPDFs As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(logFolder) Then
        MsgBox "No log folder found.", vbExclamation
        Exit Sub
    End If

    For Each file In fso.GetFolder(logFolder).files
        If LCase(Right(file.Name, 4)) = ".csv" Then
            Dim fnum As Integer, line As String, parts() As String, estSeconds As Double
            fnum = FreeFile
            Open file.path For Input As #fnum
            Line Input #fnum, line ' header

            Do While Not EOF(fnum)
                Line Input #fnum, line
                parts = Split(line, ";")
                If UBound(parts) >= 13 Then
                    estSeconds = EstimateTimeSaved(parts(3), CLng(parts(10)), CLng(parts(11)), CLng(parts(12)))
                    totalSeconds = totalSeconds + estSeconds
                    totalFiles = totalFiles + CLng(parts(10))
                    totalReps = totalReps + CLng(parts(11))
                    totalPDFs = totalPDFs + CLng(parts(12))
                End If
            Loop
            Close #fnum
        End If
    Next

    MsgBox "Total team time saved:" & vbCrLf & _
           FormatTime(totalSeconds) & vbCrLf & _
           "Files: " & totalFiles & ", Replacements: " & totalReps & ", PDFs: " & totalPDFs, vbInformation
End Sub
Public Sub ShowGlobalEfficiency()
    Dim folderPath As String
    folderPath = ThisDocument.path & "\logs"

    Dim fso As Object, file As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        frmReplaceTool.lblGlobalStats.Caption = "No logs found."
        Exit Sub
    End If

    Dim totalFiles As Long, totalReplacements As Long, totalPDFs As Long, totalSeconds As Double
    Dim userDict As Object: Set userDict = CreateObject("Scripting.Dictionary")

    For Each file In fso.GetFolder(folderPath).files
        If LCase(fso.GetExtensionName(file.Name)) = "csv" Then
            Dim fnum As Integer: fnum = FreeFile
            Dim line As String, parts() As String
            Dim isHeader As Boolean: isHeader = True
            Open file.path For Input As #fnum
            Do Until EOF(fnum)
                Line Input #fnum, line
                If isHeader Then
                    isHeader = False
                Else
                    parts = Split(line, ";")
                    If UBound(parts) >= 13 Then
                        Dim uID As String: uID = LCase(Trim(parts(1)))
                        If Not userDict.Exists(uID) Then userDict.Add uID, True

                        totalFiles = totalFiles + CLng(Val(parts(10)))
                        totalReplacements = totalReplacements + CLng(Val(parts(11)))
                        totalPDFs = totalPDFs + CLng(Val(parts(12)))
                        totalSeconds = totalSeconds + EstimateTimeSaved(parts(3), CLng(Val(parts(10))), CLng(Val(parts(11))), CLng(Val(parts(12))))
                    End If
                End If
            Loop
            Close #fnum
        End If
    Next

    Dim savedTimeStr As String
    savedTimeStr = FormatTime(totalSeconds)

    frmReplaceTool.lblGlobalStats.Caption = "Global Statistics" & vbCrLf & _
        "Total users: " & userDict.count & vbCrLf & _
        "Total Files processed: " & totalFiles & vbCrLf & _
        "Total Replacements made: " & totalReplacements & vbCrLf & _
        "Total PDFs exported: " & totalPDFs & vbCrLf & _
        "Total Estimated time saved: " & savedTimeStr
End Sub

Public Sub ShowLocalEfficiency()
    Dim folderPath As String
    folderPath = ThisDocument.path & "\logs"

    Dim fso As Object, file As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        frmReplaceTool.lblLocalStats.Caption = "No personal logs found."
        Exit Sub
    End If

    Dim files As Long, reps As Long, pdfs As Long
    Dim secondsActual As Double, secondsEstimated As Double
    Dim userID As String: userID = LCase(GetUserID())

    For Each file In fso.GetFolder(folderPath).files
        If LCase(fso.GetExtensionName(file.Name)) = "csv" Then
            Dim fnum As Integer: fnum = FreeFile
            Dim line As String, parts() As String
            Dim isHeader As Boolean: isHeader = True

            Open file.path For Input As #fnum
            Do Until EOF(fnum)
                Line Input #fnum, line
                If isHeader Then
                    isHeader = False
                Else
                    parts = Split(line, ";")
                    If UBound(parts) >= 13 Then
                        Dim csvUser As String
                        csvUser = LCase(Trim(parts(1)))
                        If csvUser = userID Then
                            files = files + CLng(Val(parts(10)))
                            reps = reps + CLng(Val(parts(11)))
                            pdfs = pdfs + CLng(Val(parts(12)))
                            secondsActual = secondsActual + CDbl(Replace(parts(13), ",", "."))
                            secondsEstimated = secondsEstimated + EstimateTimeSaved(parts(3), _
                                CLng(Val(parts(10))), CLng(Val(parts(11))), CLng(Val(parts(12))))
                        End If
                    End If
                End If
            Loop
            Close #fnum
        End If
    Next

    If files + reps + pdfs = 0 Then
        frmReplaceTool.lblLocalStats.Caption = "No personal logs found."
    Else
        frmReplaceTool.lblLocalStats.Caption = "My Statistics" & vbCrLf & _
            "User ID: " & userID & vbCrLf & _
            "Files processed: " & files & vbCrLf & _
            "Replacements made: " & reps & vbCrLf & _
            "PDFs exported: " & pdfs & vbCrLf & _
            "Estimated time saved: " & FormatTime(secondsEstimated)
    End If
End Sub



Private Sub ParseLogFile(filePath As String, ByRef files As Long, ByRef reps As Long, ByRef pdfs As Long, ByRef seconds As Long)
    Dim fnum As Integer: fnum = FreeFile
    Dim line As String, parts() As String
    Dim isHeader As Boolean: isHeader = True
    
    On Error Resume Next
    Open filePath For Input As #fnum
    Do Until EOF(fnum)
        Line Input #fnum, line
        If isHeader Then
            isHeader = False
        Else
            parts = Split(line, ";")
            If UBound(parts) >= 7 Then
                files = files + CLng(Val(parts(4)))
                reps = reps + CLng(Val(parts(5)))
                pdfs = pdfs + CLng(Val(parts(6)))
                seconds = seconds + CDbl(Replace(parts(7), ",", "."))
            End If
        End If
    Loop
    Close #fnum
End Sub

