'  ExtendableLoader.bas
'
' || Dynamically loads macros on the startup
' || of Microsoft Word. Uses a plaintext .bas
' || file from the configuration directory.
' ||
' || Configuration Directory Locatio
' ||     MacOS:
' ||     Windows:

Public isLoaded As Boolean
isLoaded = False

' Runs
Sub AutoExec()
    ' Load document

End Sub

' Just in case something screws up with the
' AutoExec function. Runs on initialization
' of a new document.
Sub AutoOpen()
    If Not isLoaded Then
        ' Load document
End Sub
