Attribute VB_Name = "modConfig"
' === modConfig – Versionshantering & formulärtitlar ===

Option Explicit

Public Const APP_VERSION As String = "v0.3 [Dev]"

Public Function GetAppVersion() As String
    GetAppVersion = "MagicWand " & APP_VERSION
End Function

Public Function GetTitle_frmIndexBrowser() As String
    GetTitle_frmIndexBrowser = ":: MagicWand | Index folders"
End Function

Public Function GetTitle_frmReplaceTool() As String
    GetTitle_frmReplaceTool = ":: MagicWand | Search & Replace"
End Function

