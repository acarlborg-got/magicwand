Attribute VB_Name = "modValidate"


Public Function ValidateProcessingSettings() As Boolean
    With frmReplaceTool
        ' Default to valid
        ValidateProcessingSettings = True

        ' ?? Preserve original + no way to separate modified files
        If .chkKeepOriginal.Value = True Then
            If Trim(.txtPreserveSubFolder.text) = "" And _
               Trim(.txtPrefix.text) = "" And Trim(.txtSuffix.text) = "" Then
                MsgBox "Preserve original is enabled, but no prefix, suffix, or subfolder is defined." & vbCrLf & _
                       "To avoid overwriting the original files, you must define at least one.", vbExclamation, "Invalid Settings"
                ValidateProcessingSettings = False
                Exit Function
            End If
        Else
            ' ? Overwrite warning
            If MsgBox("You are about to overwrite original Word files." & vbCrLf & _
                      "Are you sure you want to continue?", vbYesNo + vbCritical, "Confirm Overwrite") = vbNo Then
                ValidateProcessingSettings = False
                Exit Function
            End If
        End If
    End With
End Function




