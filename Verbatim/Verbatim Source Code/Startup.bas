Attribute VB_Name = "Startup"
' Verbatim Mac
' Copyright © 2015 Aaron Hardy
' http://paperlessdebate.com
' ashtarcommunications@gmail.com
'
' Verbatim is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License 3.0 as published by
' the Free Software Foundation.
'
' Verbatim is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License 3 for more details.
'
' For a copy of the GNU General Public License 3 see:
' http://www.gnu.org/licenses/gpl-3.0.txt

Option Explicit

Sub AutoOpen()
    
    Call Startup.Start

End Sub

Sub AutoNew()
    
    On Error Resume Next
    
    'Add doc variables with name and version number
    ThisDocument.Variables.Add Name:="Creator", Value:=GetSetting("Verbatim", "Main", "Name", "")
    ThisDocument.Variables.Add Name:="Team", Value:=GetSetting("Verbatim", "Main", "TeamName", "")
    ThisDocument.Variables.Add Name:="VerbatimVersion", Value:=Settings.GetVersion
    ThisDocument.Variables.Add Name:="VerbatimMac", Value:="True"
    ThisDocument.Saved = True
    
    Call Startup.Start
       
End Sub

Sub AutoClose()
    
    On Error Resume Next
    
    'If current doc was active speech doc, clear it
    If Paperless.ActiveSpeechDoc = ActiveDocument.Name Then Paperless.ActiveSpeechDoc = ""
    
    'Check if current file is a .doc file instead of a .docx and default save settings
    If GetSetting("Verbatim", "Admin", "SuppressDocCheck", False) = False Then
        Call Troubleshooting.CheckDocx(Notify:=True)
        Call Troubleshooting.CheckSaveFormat(Notify:=True)
    End If
    
    'If last doc, check if audio recording is still on
    If Application.Documents.Count = 1 And Toolbar.RecordAudioToggle = True Then
        If MsgBox("Audio recording appears to be active. Stop and save recording now? If you answer ""No"", recording will be lost.", vbYesNo) = vbYes Then Audio.SaveRecord
    End If
        
End Sub

Sub Start()
    
    Dim c As CommandBar
    Dim FoundToolbar As Boolean
    
    On Error Resume Next
          
    #If MAC_OFFICE_VERSION >= 15 Then
        'If Verbatim 2016 toolbar exists, make it visible
        For Each c In Application.CommandBars
            If c.Name = "Verbatim2016" Then
                FoundToolbar = True
                c.Visible = True
            End If
        Next c
    #Else
        'If Verbatim toolbar is already built, just make it visible
        For Each c In Application.CommandBars
            If c.Name = "Verbatim" Then
                FoundToolbar = True
                c.Visible = True
            End If
        Next c
    #End If
        
    #If MAC_OFFICE_VERSION >= 15 Then
        'Do nothing - can't build toolbars in VBA in Word 2016
    #Else
        'If no toolbar, build it
        If FoundToolbar = False Then
            Call Toolbar.BuildVerbatimToolbar
            CommandBars("Verbatim").Visible = True
        End If
        
        'Reposition window to avoid toolbar overlap
        If GetSetting("Verbatim", "View", "ToolbarPosition", "Top") = "Top" Then
            If FoundToolbar = False And ActiveWindow.Top < CommandBars("Verbatim").Height Then
                ActiveWindow.Top = 34 'Mac Word subtracts 34
                ActiveWindow.Height = Application.UsableHeight - 34
            End If
            ActiveWindow.Left = 0
        Else
            If ActiveWindow.Left < CommandBars("Verbatim").Width Then ActiveWindow.Left = 100
            ActiveWindow.Top = 0
        End If
    
    #End If
    
    'Set default view, zoom, and navigation pane
    Call View.DefaultView
    Call View.SetZoom
    ActiveWindow.DocumentMap = True
    
    'Refresh document styles from template if setting checked and not editing template itself
    If GetSetting("Verbatim", "Format", "AutoUpdateStyles", True) = True And ActiveDocument.FullName <> ActiveDocument.AttachedTemplate.FullName Then ActiveDocument.UpdateStyles
    ActiveDocument.Saved = True
       
    'Check if it's the first run
    If GetSetting("Verbatim", "Admin", "FirstRun", True) = True Then
        Call Startup.FirstRun
    Else
        
        'If first document opened and warnings not suppressed, check if template is incorrectly installed.
        If GetSetting("Verbatim", "Admin", "SuppressInstallChecks", False) = False And Application.Documents.Count = 1 Then
            If Troubleshooting.InstallCheckNormal = True Or Troubleshooting.InstallCheckTemplateName = True Or Troubleshooting.InstallCheckTemplateLocation = True Then
                If MsgBox("Verbatim appears to be installed incorrectly. Would you like to open the Troubleshooter? This message can be suppressed in the Verbatim settings.", vbYesNo) = vbYes Then
                    Call Settings.ShowTroubleshooter
                    Exit Sub
                End If
            End If
        End If
        
        'Check for updates weekly on Wednesdays
        If GetSetting("Verbatim", "Admin", "AutoUpdateCheck", True) = True Then
            If DateDiff("d", GetSetting("Verbatim", "Main", "LastUpdateCheck"), Now) > 6 Then
                If DatePart("w", Now) = 4 Then
                    Call Settings.UpdateCheck
                    Exit Sub
                End If
            End If
        End If
        
    End If

    'Check for custom code to import
    If GetSetting("Verbatim", "Main", "ImportCustomCode", False) = True Then
        Call Settings.ImportCustomCode(Notify:=True)
    End If

End Sub

Sub FirstRun()
    
    'Set FirstRun to False for future
    SaveSetting "Verbatim", "Admin", "FirstRun", False
    
    'Unverbatimize Normal to clear out old installs
    Call Settings.UnverbatimizeNormal
    
    'Setup keyboard shortcuts
    Call Settings.ResetKeyboardShortcuts

    #If MAC_OFFICE_VERSION >= 15 Then
        'Do nothing, can't build toolbars in VBA in Word 2016
    #Else
        'Rebuild Toolbar
        Call Toolbar.BuildVerbatimToolbar
    #End If
    
    'Run setup wizard
    Call Settings.ShowSetupWizard

End Sub
