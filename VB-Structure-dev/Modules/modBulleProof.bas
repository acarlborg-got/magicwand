Attribute VB_Name = "modBulleProof"
Sub AutoExec()
    ' D�ljer linjalen vid uppstart
    On Error Resume Next
    Application.ActiveWindow.DisplayRulers = False
    On Error GoTo 0
End Sub

Sub ViewRuler()
    ' Blockerar Ctrl+Shift+R och f�rs�k att visa linjalen
    MsgBox "�tkomst till linjalen �r permanent inaktiverad! Ring inte IT, anv�nd formatstilar...", vbExclamation
    On Error Resume Next
    Application.ActiveWindow.DisplayRulers = False
    On Error GoTo 0
End Sub

Sub AutoOpen()
    Call SkapaFormatstilar
    Call DisableFormatting
End Sub

Sub SkapaFormatstilar()
    Dim doc As Document
    Set doc = ActiveDocument

    Dim style As style

    ' Gulmarkerad
    Set style = doc.Styles.Add(Name:="Gulmarkerad", Type:=wdStyleTypeCharacter)
    With style
        .Font.Name = "Arial"
        .Font.Size = 11
        .Font.Color = wdColorBlack
        .Font.HighlightColorIndex = wdYellow
        .Priority = 1
        .QuickStyle = True
    End With

    ' Bl� understruken
    Set style = doc.Styles.Add(Name:="Bl� Understruken", Type:=wdStyleTypeCharacter)
    With style
        .Font.Name = "Arial"
        .Font.Size = 11
        .Font.Color = wdColorBlue
        .Font.Underline = wdUnderlineSingle
        .Priority = 1
        .QuickStyle = True
    End With

' R�d �verstruken
Set style = doc.Styles.Add(Name:="Utg�r", Type:=wdStyleTypeCharacter)
With style
��� .Font.Name = "Arial"
��� .Font.Size = 11
��� .Font.Color = wdColorRed
��� .Font.StrikeThrough = True
��� .Priority = 1
��� .QuickStyle = True
End With

    ' Punktlista AMA
    Set style = doc.Styles.Add(Name:="Punktlista AMA", Type:=wdStyleTypeParagraph)
    With style
        .ParagraphFormat.LeftIndent = CentimetersToPoints(0.5)
        .ParagraphFormat.SpaceAfter = 6
        .ParagraphFormat.SpaceBefore = 6
        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
        .Priority = 1
        .QuickStyle = True
    End With

    ' Numrerad AMA
    Set style = doc.Styles.Add(Name:="Numrerad AMA", Type:=wdStyleTypeParagraph)
    With style
        .ParagraphFormat.LeftIndent = CentimetersToPoints(0.5)
        .ParagraphFormat.SpaceAfter = 6
        .ParagraphFormat.SpaceBefore = 6
        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
        .Priority = 1
        .QuickStyle = True
    End With
End Sub

Sub DisableFormatting()
    Dim ctrl As CommandBarControl
    Dim ids As Variant

    ' ID:n f�r vanliga formateringskommandon
    ids = Array(21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, _
                1605, 1606, 1607, 1608, 1609, 12295, 12296)

    For Each ID In ids
        On Error Resume Next
        Set ctrl = Application.CommandBars.FindControl(ID:=ID)
        If Not ctrl Is Nothing Then ctrl.Enabled = False
        On Error GoTo 0
    Next
End Sub


