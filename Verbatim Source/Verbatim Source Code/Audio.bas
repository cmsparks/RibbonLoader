Attribute VB_Name = "Audio"
Option Explicit

Sub StartRecord()

    Dim Script As String
    Dim PressedControl As CommandBarControl
    
    'If Word 2011, get pressed ribbon control
    If Application.Version < "15" Then
        Set PressedControl = CommandBars.ActionControl
        If PressedControl Is Nothing Then Exit Sub
    End If
    
    On Error GoTo Handler
    
    'Create script
    Script = "tell application ""QuickTime Player""" & vbCr
    Script = Script & "new audio recording" & vbCr
    Script = Script & "document ""Audio Recording"" start" & vbCr
    Script = Script & "end tell"
    
    'Start recording
    #If MAC_OFFICE_VERSION >= 15 Then
        AppleScriptTask "Verbatim.scpt", "StartRecord", ""
    #Else
        MacScript (Script)
    #End If

    'Notify
    MsgBox "Audio Recording Started. Select the menu item again to stop."

    Exit Sub
    
    Set PressedControl = Nothing
    
Handler:
    RecordAudioToggle = False
    If Application.Version < "15" Then PressedControl.Caption = "Start Audio Recording"
    Set PressedControl = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Sub SaveRecord()

    Dim AudioDir As String
    Dim FileName As String
    Dim Script As String
    
    Dim PressedControl As CommandBarControl
    
    'If Word 2011, get pressed ribbon control
    If Application.Version < "15" Then
        Set PressedControl = CommandBars.ActionControl
        If PressedControl Is Nothing Then Exit Sub
    End If
    
    On Error GoTo Handler
    
    'Get Audio recording directory from settings
    AudioDir = GetSetting("Verbatim", "Paperless", "AudioDir", "?")
    
    'If blank, save to desktop
    If AudioDir = "?" Then
        AudioDir = MacScript("return the path to the desktop folder as text")
    End If
    
    'Add trailing :
    If Right(AudioDir, 1) <> ":" Then AudioDir = AudioDir & ":"
    
GetFileName:
    'Get name for recording
    FileName = InputBox("Please enter a name for your saved audio file. It will be saved to the following directory:" & vbCrLf & "(Configurable in Settings)" & vbCrLf & AudioDir, "Save Audio Recording", "Recording " & Format(Now, "m d yyyy hmmAMPM"))
    
    'Exit if no file name or user pressed Cancel, recording is still active
    If FileName = "" Then
        RecordAudioToggle = True
        If Application.Version < "15" Then PressedControl.Caption = "Stop Audio Recording"
        Exit Sub
    End If
    
    'Clean up filename and ensure correct extension
    FileName = Strings.OnlyAlphaNumericChars(FileName)
    If Right(FileName, 4) <> ".m4a" Then FileName = FileName & ".m4a"
    FileName = AudioDir & FileName
    
    'Test if file exists
    #If MAC_OFFICE_VERSION >= 15 Then
        If AppleScriptTask("Verbatim.scpt", "FileExists", FileName) = "true" Then
            If MsgBox("File exists. Overwrite?", vbYesNo) = vbNo Then GoTo GetFileName
        End If
    #Else
        If MacScript("tell application ""Finder""" & Chr(13) & "exists file """ & FileName & """" & Chr(13) & "end tell") = "true" Then
            If MsgBox("File exists. Overwrite?", vbYesNo) = vbNo Then GoTo GetFileName
        End If
    #End If
    
    'Create script
    Script = "tell application ""Finder""" & vbCr
    Script = Script & "set exportFile to """ & FileName & """" & vbCr
    Script = Script & "tell application ""QuickTime Player""" & vbCr
    Script = Script & "stop document ""Audio Recording""" & vbCr
    Script = Script & "tell last item of documents" & vbCr
    Script = Script & "export in file exportFile using settings preset ""Audio Only""" & vbCr
    Script = Script & "close without saving" & vbCr
    Script = Script & "end tell" & vbCr
    'Script = Script & "ignoring application responses" & vbCr
    'Script = Script & "quit" & vbCr
    'Script = Script & "end ignoring" & vbCr
    Script = Script & "end tell" & vbCr
    Script = Script & "end tell"
    
    'Stop and save recording
    #If MAC_OFFICE_VERSION >= 15 Then
        AppleScriptTask "Verbatim.scpt", "SaveRecord", FileName
    #Else
        MacScript (Script)
    #End If
    
    'Notify
    MsgBox "Recording being saved as:" & vbCrLf & FileName, vbOKOnly

    Set PressedControl = Nothing
    
    Exit Sub
    
Handler:
    RecordAudioToggle = True
    If Application.Version < "15" Then PressedControl.Caption = "Stop Audio Recording"
    Set PressedControl = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub
