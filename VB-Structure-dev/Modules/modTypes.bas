Attribute VB_Name = "modTypes"


' ========================================
' modTypes – Globala typer för indexering
' ========================================

Option Explicit

' === Komplett metadata för varje fil i index ===
Public Type IndexedFile
    ID As Long
    filePath As String
    fileName As String
    parentFolderID As Long
    extension As String
    lastModified As Date
    selected As Boolean
End Type

' === Mappstruktur med djup ===
Public Type IndexedFolder
    ID As Long
    folderPath As String
    selected As Boolean
    depth As Long
End Type

' === Endast sparat urval: filnamn + mappreferens ===
Public Type tFileSelection
    fileName As String
    parentID As Long
    filePath As String
End Type

' === Globala arrayer ===
Public IndexedFolders() As IndexedFolder
Public IndexedFiles() As IndexedFile
Public selectedFiles() As tFileSelection
Public selectedFolders() As IndexedFolder

' === Insamlad metadata per Wordfil ===
Public Type FileMetadata
    Title As String
    Subject As String
    Author As String
    Keywords As String
    Datum As String
    Handlaggare As String
    Konstruktor As String
    DocumentDate As String
End Type

Public MetadataIndex() As FileMetadata

' === Regel för villkorstyrd metadata ===
Public Type MetaRule
    Field As String
    Condition As String
    Action As String
    valueSource As String
    valueText As String
End Type

' === modTypes – Globala typer för indexering och statistik ===

Public Type tMatchStat
    foundText As String
    count As Long
    sourceFile As String
End Type

Public MatchStats() As tMatchStat


