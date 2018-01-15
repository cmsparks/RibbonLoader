Attribute VB_Name = "Troubleshooting"
Option Explicit

'*************************************************************************************
'* INSTALL CHECK FUNCTIONS                                                                               *
'*************************************************************************************
Function InstallCheckNormal(Optional Notify As Boolean) As Boolean
'Checks if Verbatim is installed as Normal.dotm and optionally notifies the user.
    
    Dim msg As String
    
    On Error Resume Next
    
    If ActiveDocument.AttachedTemplate.Name = "Normal.dotm" Then
        InstallCheckNormal = True
        
        If Notify = True Then
            msg = "WARNING - Verbatim appears to be incorrectly installed as Normal.dotm. " _
            & "Verbatim is not designed to be your normal template, and many features will " _
            & "not work correctly. Instead, you should change the filename back to Debate.dotm " _
            & "and use the ""Always On"" mode in the Verbatim settings. This message can be suppressed in the Verbatim settings."
            MsgBox (msg)
        End If
        
    Else
        InstallCheckNormal = False
    End If

End Function

Function InstallCheckTemplateName(Optional Notify As Boolean) As Boolean
'Checks if Verbatim is installed as the wrong filename and optionally notifies the user

    Dim msg As String
    
    On Error Resume Next
    
    If ActiveDocument.AttachedTemplate.Name <> "Debate.dotm" Then
        InstallCheckTemplateName = True
        
        If Notify = True Then
            msg = "WARNING - Verbatim appears to be installed incorrectly. " _
            & "Verbatim should always be named ""Debate.dotm"" or many features will not work correctly. " _
            & "If you have changed the filename, it will break compatibility with others. " _
            & "It is strongly recommended you change the file name back. " _
            & "This warning can be suppressed in the Verbatim Options."
            MsgBox (msg)
        End If

    Else
        InstallCheckTemplateName = False
    End If
    
End Function

Function InstallCheckTemplateLocation(Optional Notify As Boolean) As Boolean
'Checks if Verbatim is installed in the wrong location and optionally notifes the user

    Dim msg As String
    
    On Error Resume Next
    
    #If MAC_OFFICE_VERSION >= 15 Then
    If ActiveDocument.AttachedTemplate.Path <> Application.NormalTemplate.Path Then
    #Else
    If ActiveDocument.AttachedTemplate.Path <> Application.NormalTemplate.Path & ":My Templates" Then
    #End If
        InstallCheckTemplateLocation = True
        
        If Notify = True Then
            msg = "WARNING - Verbatim appears to be installed in the wrong location. " _
            & "The Verbatim template file (Debate.dotm) should be located in your My Templates folder, usually located at: " _
            & vbCrLf & "Word 2011:" _
            & vbCrLf & "~/Library/Application Support/Microsoft/Office/User Templates/My Templates" _
            & vbCrLf & "or Word 2016:" _
            & vbCrLf & "~/Library/Group Containers/UBF8T346G9.Office/User Content/Templates" _
            & "Using it from a different location will break many features. " _
            & "You can open your templates folder or suppress this warning in the Verbatim settings."
            MsgBox (msg)
        End If

    Else
        InstallCheckTemplateLocation = False
    End If
    
End Function

Function CheckSaveFormat(Optional Notify As Boolean) As Boolean
'Check if default save format is .docx and optionally notifies the user

    Dim msg As String
    
    On Error Resume Next
    
    If Application.DefaultSaveFormat = "Doc" Or Application.DefaultSaveFormat = "Doc97" Then
        CheckSaveFormat = True
        
        If Notify = True Then
            msg = "Your default save format appears to be set to .doc instead of .docx"
            msg = msg & " - It is highly recommended that you use the .docx format instead. "
            msg = msg & "Change automatically?" & vbCrLf & "(This warning can be supressed in the Verbatim options)"
            If MsgBox(msg, vbYesNo) = vbYes Then Application.DefaultSaveFormat = "WordDocument"
        End If
    
    Else
        CheckSaveFormat = False
    End If
    
End Function

Function CheckDocx(Optional Notify As Boolean) As Boolean
'Check if current document is a .doc

    Dim msg As String
    
    On Error Resume Next
    
    If Right(ActiveDocument.Name, 3) = "doc" Then
        CheckDocx = True
        
        If Notify = True Then
            msg = "This file is saved as .doc instead of .docx"
            msg = msg & " - It is highly recommended that you use the .docx format instead. "
            msg = msg & "Save as .docx automatically? This will overwrite any current file in the same directory with the same name." & vbCrLf & "(This warning can be supressed in the Verbatim options)"
            If MsgBox(msg, vbYesNo) = vbYes Then
                ActiveDocument.SaveAs FileName:=Left(ActiveDocument.FullName, InStrRev(ActiveDocument.FullName, ".") - 1), FileFormat:=wdFormatXMLDocument
            End If
        End If
    
    Else
        CheckDocx = False
    End If

End Function

'*************************************************************************************
'* FIX FUNCTIONS                                                                                                   *
'*************************************************************************************
Sub DeleteDuplicateTemplates()
    
    Dim FilePath As String
    
    On Error Resume Next
    
    'Check for "Debate.dotm" in the Desktop and Downloads folders, prompt to delete if found
    FilePath = MacScript("return the path to the desktop folder as string") & "Debate.dotm"
    
    #If MAC_OFFICE_VERSION >= 15 Then
        If AppleScriptTask("Verbatim.scpt", "FileExists", FilePath) = "true" Then
    #Else
         If MacScript("tell application ""Finder""" & Chr(13) & "exists file """ & FilePath & """" & Chr(13) & "end tell") = "true" Then
    #End If
            If MsgBox("A duplicate copy of Debate.dotm was found on your Desktop - this can cause interoperability issues. Attempt to delete automatically?", vbYesNo) = vbYes Then
                Call Filesystem.KillFileOnMac(FilePath)
            End If
        End If
    
    FilePath = MacScript("return the path to the downloads folder as string") & "Debate.dotm"
    
    #If MAC_OFFICE_VERSION >= 15 Then
        If AppleScriptTask("Verbatim.scpt", "FileExists", FilePath) = "true" Then
    #Else
        If MacScript("tell application ""Finder""" & Chr(13) & "exists file """ & FilePath & """" & Chr(13) & "end tell") = "true" Then
    #End If
            If MsgBox("A duplicate copy of Debate.dotm was found in your Downloads folder - this can cause interoperability issues. Attempt to delete automatically?", vbYesNo) = vbYes Then
                Call Filesystem.KillFileOnMac(FilePath)
            End If
        End If
    
End Sub

Sub SetDefaultSave()
    Application.DefaultSaveFormat = "WordDocument"
End Sub
