Attribute VB_Name = "modEfficiency"
' #############################################
' ## modEfficiency � Manual time estimation
' #############################################

Option Explicit

' === H�rdkodad "prislista" i sekunder per �tg�rd ===
Public Function EstimateTimeSaved(actionType As String, _
                                  files As Long, _
                                  replacements As Long, _
                                  pdfs As Long) As Double
    Dim secPerFile As Double, secPerReplace As Double, secPerPDF As Double

    Select Case actionType
        Case "Replace+PDF"
            secPerFile = 90 ' s�ka, ers�tta, spara
            secPerReplace = 30 ' extra f�r varje ers�ttning
            secPerPDF = 30 ' export och d�pning

        Case "FindDates"
            secPerFile = 20

        Case "Spellcheck"
            secPerFile = 45

        Case Else
            secPerFile = 30
    End Select

    EstimateTimeSaved = files * secPerFile + replacements * secPerReplace + pdfs * secPerPDF
End Function

' === Formatera tid i arbetsdagar, timmar och minuter ===
Public Function FormatTime(seconds As Double) As String
    Dim days As Long, hours As Long, minutes As Long

    days = seconds \ (60 * 60 * 7.5) ' arbetsdagar � 7.5 h
    hours = (seconds Mod (60 * 60 * 7.5)) \ 3600
    minutes = (seconds Mod 3600) \ 60

    FormatTime = days & "d " & hours & "h " & minutes & "min"
End Function


