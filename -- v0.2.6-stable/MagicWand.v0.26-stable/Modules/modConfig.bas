Attribute VB_Name = "modConfig"
' #############################################
' ## modConfig – Versionshantering
' #############################################

Option Explicit

Public Const APP_VERSION As String = "v0.26 [Stable]"

Public Function GetAppVersion() As String
    GetAppVersion = "MagicWand " & APP_VERSION
End Function


