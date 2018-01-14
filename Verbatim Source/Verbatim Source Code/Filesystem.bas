Attribute VB_Name = "Filesystem"
Option Explicit

Function GetSubfoldersInFolder(FolderPath) As String

    Dim Script As String
    
    Script = "tell application ""Finder""" & Chr(13)
    Script = Script & "set r to """"" & Chr(13)
    Script = Script & "set myFolders to folders of folder""" & FolderPath & """" & Chr(13)
    Script = Script & "repeat with f in myFolders" & Chr(13)
    Script = Script & "set r to (r & f as string) & ""\n""" & Chr(13)
    Script = Script & "end repeat" & Chr(13)
    Script = Script & "return r" & Chr(13)
    Script = Script & "end tell"
    
    #If MAC_OFFICE_VERSION >= 15 Then
        GetSubfoldersInFolder = AppleScriptTask("Verbatim.scpt", "GetSubfoldersInFolder", FolderPath)
    #Else
        GetSubfoldersInFolder = MacScript(Script)
    #End If
    
    'Trim trailing newline
    If Right(GetSubfoldersInFolder, 1) = Chr(10) Or Right(GetSubfoldersInFolder, 1) = Chr(13) Then GetSubfoldersInFolder = Left(GetSubfoldersInFolder, Len(GetSubfoldersInFolder) - 1)
    
End Function
Function GetFilesInFolder(FolderPath) As String

    Dim POSIXPath As String
    Dim Script As String
    
    POSIXPath = MacScript("tell text 1 thru -2 of " & Chr(34) & FolderPath & Chr(34) & " to return quoted form of it's POSIX Path")
    
    Script = "set streamEditorCommand to " & Chr(34) & " |  tr  [/:] [:/] " & Chr(34) & Chr(13)
    Script = Script & "set streamEditorCommand to streamEditorCommand & " & Chr(34)
    Script = Script & " | sed -e " & Chr(34) & "  & quoted form of (" & Chr(34) & " s.:." & Chr(34)
    Script = Script & "  & (POSIX file " & Chr(34) & "/" & Chr(34) & "  as string) & " & Chr(34) & "." & Chr(34) & " )" & Chr(13)
    Script = Script & "do shell script """ & "find -E " & POSIXPath
    Script = Script & " -iregex " & "'.*/[^~][^/]*\\." & "(docx|doc|docm|dot|dotm)" & "$' " & "-maxdepth 1"
    Script = Script & """ & streamEditorCommand without altering line endings"

    #If MAC_OFFICE_VERSION >= 15 Then
        GetFilesInFolder = AppleScriptTask("Verbatim.scpt", "GetFilesInFolder", POSIXPath)
    #Else
        GetFilesInFolder = MacScript(Script)
    #End If
    
    'Trim trailing newline
    If Right(GetFilesInFolder, 1) = Chr(10) Or Right(GetFilesInFolder, 1) = Chr(13) Then GetFilesInFolder = Left(GetFilesInFolder, Len(GetFilesInFolder) - 1)
        
End Function

Sub KillFileOnMac(FileName As String)
'Built-in Kill doesn't work with filenames over 28 characters

    Dim Script As String
    
    On Error Resume Next
    
    #If MAC_OFFICE_VERSION >= 15 Then
        AppleScriptTask "Verbatim.scpt", "KillFileOnMac", FileName
    #Else
        Script = "tell application " & Chr(34) & "Finder" & Chr(34) & Chr(13)
        Script = Script & "do shell script ""rm "" & quoted form of posix path of " & Chr(34) & FileName & Chr(34) & Chr(13)
        Script = Script & "end tell"
        MacScript (Script)
    #End If
    
End Sub
