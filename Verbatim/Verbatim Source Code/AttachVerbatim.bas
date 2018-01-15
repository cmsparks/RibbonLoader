Attribute VB_Name = "AttachVerbatim"
'This module will be copied to Normal.dotm to enable "Always On" mode and the Verbatimize button
Option Explicit

Sub AutoExec()
    
    If GetSetting("Verbatim", "Admin", "AlwaysOn", False) = True Then
            Call AttachVerbatim
    End If
    
End Sub

Sub AttachVerbatim()
    
    On Error GoTo Handler
    
    'Check template exists
    #If MAC_OFFICE_VERSION >= 15 Then
        If AppleScriptTask("Verbatim.scpt", "FileExists", "Macintosh HD" & Replace(Replace(Application.NormalTemplate.Path & "/Debate.dotm", ".localized", ""), "/", ":")) = "false" Then
            Application.StatusBar = "Debate.dotm not found in your Templates folder - it must be installed correctly to attach it."
            Exit Sub
        End If
    #Else
        If MacScript("tell application ""Finder""" & Chr(13) & "exists file """ & Application.NormalTemplate.Path & ":My Templates:Debate.dotm" & """" & Chr(13) & "end tell") = "false" Then
            Application.StatusBar = "Debate.dotm not found in your My Templates folder - it must be installed correctly to attach it."
            Exit Sub
        End If
    #End If

    'If starting Word from scratch, add a new doc based on the template - will suppress Word's built-in doc
    If Application.Documents.Count = 0 Then
        #If MAC_OFFICE_VERSION >= 15 Then
            Application.Documents.Add Template:=Application.NormalTemplate.Path & "/Debate.dotm"
        #Else
            Application.Documents.Add Template:=Application.NormalTemplate.Path & ":My Templates:Debate.dotm"
        #End If
    Else
        'Attach Verbatim to the current doc
        #If MAC_OFFICE_VERSION >= 15 Then
            ActiveDocument.AttachedTemplate = Application.NormalTemplate.Path & "/Debate.dotm"
            Application.AddIns(Application.NormalTemplate.Path & "/Debate.dotm").Installed = True
        #Else
            ActiveDocument.AttachedTemplate = Application.NormalTemplate.Path & ":My Templates:Debate.dotm"
        #End If
        ActiveDocument.UpdateStyles
    End If
    
    'Make debate toolbar visible
    If Application.Version < "15" Then CommandBars("Verbatim").Visible = True

    Exit Sub
    
Handler:
    Application.StatusBar = "Error Attaching Verbatim. Error " & Err.Number & ": " & Err.Description

End Sub
