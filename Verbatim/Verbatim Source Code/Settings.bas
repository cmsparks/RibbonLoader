Attribute VB_Name = "Settings"
Option Explicit

'*************************************************************************************
'* SHOW FORMS FUNCTIONS                                                                                  *
'*************************************************************************************

Sub ShowSettingsForm()

    Dim SettingsForm As frmSettings
    Set SettingsForm = New frmSettings
    SettingsForm.Show

End Sub

Sub ShowVerbatimHelp()

    Dim HelpForm As frmHelp
    Set HelpForm = New frmHelp
    HelpForm.Show

End Sub

Sub ShowCheatSheet()

    Dim CheatSheet As frmCheatSheet
    Set CheatSheet = New frmCheatSheet
    CheatSheet.Show

End Sub

Sub ShowSetupWizard()

    Dim SetupWizard As frmSetupWizard
    Set SetupWizard = New frmSetupWizard
    SetupWizard.Show
    
End Sub

Sub ShowTroubleshooter()

    Dim Troubleshooter As frmTroubleshooter
    Set Troubleshooter = New frmTroubleshooter
    Troubleshooter.Show
    
End Sub

'*************************************************************************************
'* VERBATIMIZE FUNCTIONS                                                                                   *
'*************************************************************************************

Sub VerbatimizeNormal(Optional Notify As Boolean)
'Copies the "AttachVerbatim" module to the normal template and adds a button to the Standard Toolbar
'Must copy a whole module, not an individual macro

    On Error GoTo Handler

    'Unverbatimize first to start with a clean slate
    Call Settings.UnverbatimizeNormal(Notify:=False)
   
    'Copy the AttachVerbatim module
    Application.OrganizerCopy Source:=ActiveDocument.AttachedTemplate.FullName, Destination:=Application.NormalTemplate.FullName, Name:="AttachVerbatim", Object:=3

    'Make a global template, or add Verbatimize button to Normal template
    #If MAC_OFFICE_VERSION >= 15 Then
        Application.AddIns.Add (Application.NormalTemplate.Path & "/Debate.dotm")
    #Else
        Call Settings.CreateVerbatimizeButton
    #End If
    
    'Notify
    If Notify = True Then MsgBox "Normal template successfully verbatimized!"

    Exit Sub
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Sub UnverbatimizeNormal(Optional Notify As Boolean)

    Dim StandardToolbar As CommandBar
    Dim c
    Dim Toolbar As CommandBar
    
    'Delete module (and old one) from normal template - turn off error checking in case they don't exist
    On Error Resume Next
    Application.OrganizerDelete Source:=Application.NormalTemplate.FullName, Name:="AttachVerbatim", Object:=3
    Application.OrganizerDelete Source:=Application.NormalTemplate.FullName, Name:="Verbatim_AttachTemplate", Object:=3

    On Error GoTo Handler
    
    'If 2011, delete the Verbatimize button from the Standard toolbar
    If Application.Version < "15" Then
        CustomizationContext = NormalTemplate
        Set StandardToolbar = Application.CommandBars("Standard")
        For Each c In StandardToolbar.Controls
            If c.Tag = "Verbatimize" Then c.Delete
        Next c
    End If
    
    'If 2011, delete the old toolbar if it exists
    If Application.Version < "15" Then
        For Each Toolbar In Application.CommandBars
            If Toolbar.Name = "VerbatimNormal" Then Toolbar.Delete
        Next Toolbar
    End If
    
    CustomizationContext = ActiveDocument
    
    'Notify
    If Notify = True Then MsgBox "Normal template successfully un-verbatimized!"
    
    Set StandardToolbar = Nothing
    
    Exit Sub
    
Handler:
    Set StandardToolbar = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Sub CreateVerbatimizeButton()

    Dim StandardToolbar As CommandBar
    Dim VerbatimizeButton As CommandBarControl

    If Application.Version < "15" Then

        CustomizationContext = NormalTemplate
        
        Set StandardToolbar = Application.CommandBars("Standard")
        StandardToolbar.Visible = True
        
        'Add a button for the Attach template macro
        Set VerbatimizeButton = StandardToolbar.Controls.Add(Before:=1)
        VerbatimizeButton.Caption = "Verbatimize"
        VerbatimizeButton.DescriptionText = "Verbatimize the current document"
        VerbatimizeButton.TooltipText = "Verbatimize"
        VerbatimizeButton.BeginGroup = True
        VerbatimizeButton.Tag = "Verbatimize"
        VerbatimizeButton.FaceId = 938    'Pin/note icon
        VerbatimizeButton.Style = msoButtonIcon
        VerbatimizeButton.OnAction = "AttachVerbatim.AttachVerbatim"
        
        CustomizationContext = ActiveDocument
        
        Set StandardToolbar = Nothing
        Set VerbatimizeButton = Nothing
    Else
        Exit Sub
    End If
    
End Sub

'*************************************************************************************
'* IMPORT/EXPORT FUNCTIONS                                                                              *
'*************************************************************************************

Sub ImportCustomCode(Optional Notify As Boolean)

    Dim p As VBIDE.VBProject

    'Turn on Error Handling
    On Error GoTo Handler

    'Set registry setting to avoid repeatedly trying to import code
    SaveSetting "Verbatim", "Main", "ImportCustomCode", False

    'Make sure custom code file exists
    #If MAC_OFFICE_VERSION >= 15 Then
    If AppleScriptTask("Verbatim.scpt", "FileExists", "Macintosh HD" & Replace(Application.NormalTemplate.Path & "/VerbatimCustomCode.bas", "/", ":")) = "false" Then
    #Else
    If MacScript("tell application ""Finder""" & Chr(13) & "exists file """ & Application.NormalTemplate.Path & ":My Templates:VerbatimCustomCode.bas" & """" & Chr(13) & "end tell") = "false" Then
    #End If
        If Notify = True Then MsgBox "No custom code module found in your Templates folder. It must be named ""VerbatimCustomCode.bas"" to import."
        Exit Sub
    End If
    
    'Warn user
    If MsgBox("Attemping to import custom code - this will overwrite your current custom code module. Proceed?", vbOKCancel) = vbCancel Then Exit Sub
    
    'Delete current Custom code module - turn off error checking temporarily in case it doesn't exist
    On Error Resume Next
    Application.OrganizerDelete Source:=ActiveDocument.AttachedTemplate.FullName, Name:="Custom", Object:=3
    On Error GoTo Handler
    
    'Import the module and delete the file
    Set p = FindVBProject(ActiveDocument.AttachedTemplate.Path & "/" & ActiveDocument.AttachedTemplate)
    If p Is Nothing Then
        MsgBox "Failed to import custom code."
        Exit Sub
    End If
    
    #If MAC_OFFICE_VERSION >= 15 Then
        p.VBComponents.Import (Application.NormalTemplate.Path & "/VerbatimCustomCode.bas")
        Call Filesystem.KillFileOnMac(Application.NormalTemplate.Path & "/VerbatimCustomCode.bas")
    #Else
        p.VBComponents.Import (Application.NormalTemplate.Path & ":My Templates:VerbatimCustomCode.bas")
        Call Filesystem.KillFileOnMac(Application.NormalTemplate.Path & ":My Templates:VerbatimCustomCode.bas")
    #End If
    
    If Notify = True Then MsgBox "Custom code successfully imported!"

    Set p = Nothing

    Exit Sub

Handler:
    Set p = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Sub ExportCustomCode(Optional Notify As Boolean)

    Dim p As VBIDE.VBProject
    Dim Module As VBIDE.VBComponent
    
    'Turn on Error Handling
    On Error GoTo Handler

    'Find the Custom code module - accessing it directly is removed on the Mac
    Set p = FindVBProject(ActiveDocument.AttachedTemplate.Path & "/" & ActiveDocument.AttachedTemplate)
    If p Is Nothing Then
        MsgBox "Failed to find the Custom code module."
        Exit Sub
    End If
    
    Set Module = p.VBComponents("Custom")
    If Module.CodeModule.CountOfLines <= 1 Then
        If Notify = True Then MsgBox "No custom code found."
        Exit Sub
    End If
    
    'Export the module
    #If MAC_OFFICE_VERSION >= 15 Then
        Module.Export Application.NormalTemplate.Path & "/VerbatimCustomCode.bas"
    #Else
        Module.Export Application.NormalTemplate.Path & ":My Templates:VerbatimCustomCode.bas"
    #End If
    
    'Set registry for automatic import on startup
    SaveSetting "Verbatim", "Main", "ImportCustomCode", True
    
    If Notify = True Then MsgBox "Custom code exported as VerbatimCustomCode.bas to your Templates folder."
   
    Set p = Nothing
    Set Module = Nothing
    
    Exit Sub

Handler:
    Set p = Nothing
    Set Module = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Private Function FindVBProject(d As String) As VBIDE.VBProject
    
    Dim p As VBIDE.VBProject
    
    On Error Resume Next
    
    For Each p In Application.VBE.VBProjects
        If (p.FileName = d) Then
            Set FindVBProject = p
            Exit Function
        End If
    Next
    
End Function

'*************************************************************************************
'* UPDATE FUNCTIONS                                                                                           *
'*************************************************************************************

Sub UpdateCheck(Optional Notify As Boolean)
'Check for Verbatim updates
    
    Dim XML As String
    Dim Script As String
    
    Dim DownloadURL As String
    Dim DownloadFile As String
    
    'Turn on error checking
    On Error GoTo Handler

    Application.StatusBar = "Checking for Verbatim updates..."

    'Get update XML
    XML = MacScript("do shell script ""curl 'http://update.paperlessdebate.com/verbatim.xml'""")
    
    'Exit if the request fails
    If Len(XML) < 1 Then
        Application.StatusBar = "Update Check Failed"
        SaveSetting "Verbatim", "Main", "LastUpdateCheck", Now
        If Notify = True Then MsgBox "Update Check Failed."
        Exit Sub
    End If
    
    'Set LastUpdateCheck
    SaveSetting "Verbatim", "Main", "LastUpdateCheck", Now
    
    'If newer version is found
    If Mid(XML, InStr(XML, "<macversion>") + 12, InStr(XML, "</macversion>") - InStr(XML, "<macversion>") - 12) > Settings.GetVersion Then
        
        'Confirm update
        If MsgBox("There is a newer version of Verbatim available for download. Would you like to close Word and update automatically? You will be given the option of saving any open files, and any custom code will be exported automatically.", vbYesNo) = vbNo Then Exit Sub
            
        Application.StatusBar = "Downloading updates..."
        
        'Get the URL for latest PC version
        DownloadURL = Mid(XML, InStr(XML, "<macurl>") + 8, InStr(XML, "</macurl>") - InStr(XML, "<macurl>") - 8)
        
        'Save file to disk
        DownloadFile = MacScript("return POSIX path of (path to temporary items from user domain) as string")
        DownloadFile = DownloadFile & Mid(XML, InStr(XML, "<macfilename>") + 13, InStr(XML, "</macfilename>") - InStr(XML, "<macfilename>") - 13)
        MacScript ("do shell script ""curl -o '" & DownloadFile & "' '" & DownloadURL & "'""")
        
        'Try exporting settings
        Call Settings.ExportCustomCode(Notify:=False)
        
        'Launch installer
        Application.StatusBar = "Launching installer..."
        MacScript ("do shell script ""open '" & DownloadFile & "'""")
        Application.Quit wdPromptToSaveChanges
    
    Else
        Application.StatusBar = "No Verbatim updates found."
        If Notify = True Then MsgBox "No Verbatim updates found."
    End If
         
    Exit Sub

Handler:
    Application.StatusBar = "Update Check Failed. Error " & Err.Number & ": " & Err.Description
    If Notify = True Then MsgBox "Update Check Failed. Error " & Err.Number & ": " & Err.Description

End Sub

'*************************************************************************************
'* KEYBOARD FUNCTIONS                                                                                       *
'*************************************************************************************

Sub ChangeKeyboardShortcut(KeyName As WdKey, MacroName As String)
'Change keyboard shortcuts in template
    
    Application.CustomizationContext = ActiveDocument.AttachedTemplate
    
    Select Case MacroName
        Case Is = "Paste"
            KeyBindings.Add wdKeyCategoryMacro, "Formatting.PasteText", BuildKeyCode(KeyName)
        Case Is = "Condense"
            KeyBindings.Add wdKeyCategoryMacro, "Formatting.Condense", BuildKeyCode(KeyName)
        Case Is = "Pocket"
            KeyBindings.Add wdKeyCategoryStyle, "Pocket", BuildKeyCode(KeyName)
        Case Is = "Hat"
            KeyBindings.Add wdKeyCategoryStyle, "Hat", BuildKeyCode(KeyName)
        Case Is = "Block"
            KeyBindings.Add wdKeyCategoryStyle, "Block", BuildKeyCode(KeyName)
        Case Is = "Tag"
            KeyBindings.Add wdKeyCategoryStyle, "Tag", BuildKeyCode(KeyName)
        Case Is = "Cite"
            KeyBindings.Add wdKeyCategoryStyle, "Cite", BuildKeyCode(KeyName)
        Case Is = "Underline"
            KeyBindings.Add wdKeyCategoryMacro, "Formatting.ToggleUnderline", BuildKeyCode(KeyName)
        Case Is = "Emphasis"
            KeyBindings.Add wdKeyCategoryStyle, "Emphasis", BuildKeyCode(KeyName)
        Case Is = "Highlight"
            KeyBindings.Add wdKeyCategoryMacro, "Formatting.Highlight", BuildKeyCode(KeyName)
        Case Is = "Clear"
            KeyBindings.Add wdKeyCategoryMacro, "Formatting.ClearToNormal", BuildKeyCode(KeyName)
        Case Is = "Shrink Text"
            KeyBindings.Add wdKeyCategoryMacro, "Formatting.ShrinkText", BuildKeyCode(KeyName)
        Case Is = "Select Similar"
            KeyBindings.Add wdKeyCategoryMacro, "Formatting.SelectSimilar", BuildKeyCode(KeyName)
        Case Else
            'Nothing
        
    End Select
    
    Application.CustomizationContext = ThisDocument

End Sub

Sub ResetKeyboardShortcuts()
      
    On Error Resume Next
    
    'Clear old keybindings
    Call Settings.RemoveKeyBindings
    
    'Save defaults
    SaveSetting "Verbatim", "Keyboard", "F2Shortcut", "Paste"
    SaveSetting "Verbatim", "Keyboard", "F3Shortcut", "Condense"
    SaveSetting "Verbatim", "Keyboard", "F4Shortcut", "Pocket"
    SaveSetting "Verbatim", "Keyboard", "F5Shortcut", "Hat"
    SaveSetting "Verbatim", "Keyboard", "F6Shortcut", "Block"
    SaveSetting "Verbatim", "Keyboard", "F7Shortcut", "Tag"
    SaveSetting "Verbatim", "Keyboard", "F8Shortcut", "Cite"
    SaveSetting "Verbatim", "Keyboard", "F9Shortcut", "Underline"
    SaveSetting "Verbatim", "Keyboard", "F10Shortcut", "Emphasis"
    SaveSetting "Verbatim", "Keyboard", "F11Shortcut", "Highlight"
    SaveSetting "Verbatim", "Keyboard", "F12Shortcut", "Clear"

    'Save shortcuts in the template
    Application.CustomizationContext = ActiveDocument.AttachedTemplate

    'Set keyboard shortcuts
    KeyBindings.Add wdKeyCategoryMacro, "Settings.ShowVerbatimHelp", BuildKeyCode(wdKeyF1)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.PasteText", BuildKeyCode(wdKeyF2)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.Condense", BuildKeyCode(wdKeyF3)
    KeyBindings.Add wdKeyCategoryStyle, "Pocket", BuildKeyCode(wdKeyF4)
    KeyBindings.Add wdKeyCategoryStyle, "Hat", BuildKeyCode(wdKeyF5)
    KeyBindings.Add wdKeyCategoryStyle, "Block", BuildKeyCode(wdKeyF6)
    KeyBindings.Add wdKeyCategoryStyle, "Tag", BuildKeyCode(wdKeyF7)
    KeyBindings.Add wdKeyCategoryStyle, "Cite", BuildKeyCode(wdKeyF8)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.ToggleUnderline", BuildKeyCode(wdKeyF9)
    KeyBindings.Add wdKeyCategoryStyle, "Emphasis", BuildKeyCode(wdKeyF10)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.Highlight", BuildKeyCode(wdKeyF11)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.ClearToNormal", BuildKeyCode(wdKeyF12)
    
    KeyBindings.Add wdKeyCategoryMacro, "View.SwitchWindows", BuildKeyCode(wdKeyControl, wdKeyTab)
    KeyBindings.Add wdKeyCategoryMacro, "Settings.ShowSettingsForm", BuildKeyCode(wdKeyAlt, wdKeyF1)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.GetFromCiteMaker", BuildKeyCode(wdKeyAlt, wdKeyF2)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.SelectSimilar", BuildKeyCode(wdKeyCommand, wdKeyF2)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.CondenseNoPilcrows", BuildKeyCode(wdKeyCommand, wdKeyF3)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.ShrinkText", BuildKeyCode(wdKeyAlt, wdKeyF3)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.ShrinkText", BuildKeyCode(wdKeyCommand, wdKey8)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.AutoFormatCite", BuildKeyCode(wdKeyCommand, wdKeyF8)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.AutoEmphasizeFirst", BuildKeyCode(wdKeyCommand, wdKeyF10)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.CopyPreviousCite", BuildKeyCode(wdKeyAlt, wdKeyF8)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.AutoUnderline", BuildKeyCode(wdKeyAlt, wdKeyF9)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.RemoveEmphasis", BuildKeyCode(wdKeyCommand, wdKeyShift, wdKeyF10)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.UpdateStyles", BuildKeyCode(wdKeyCommand, wdKeyF12)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.AutoNumberTags", BuildKeyCode(wdKeyCommand, wdKeyAlt, wdKey3)
    
    KeyBindings.Add wdKeyCategoryMacro, "Paperless.MoveUp", BuildKeyCode(wdKeyAlt, vbKeyUp)
    KeyBindings.Add wdKeyCategoryMacro, "Paperless.MoveDown", BuildKeyCode(wdKeyAlt, vbKeyDown)
    KeyBindings.Add wdKeyCategoryMacro, "Paperless.DeleteHeading", BuildKeyCode(wdKeyAlt, vbKeyLeft)
    KeyBindings.Add wdKeyCategoryMacro, "Paperless.SendToSpeech", BuildKeyCode(wdKeyAlt, vbKeyRight)
    KeyBindings.Add wdKeyCategoryMacro, "Paperless.SendToSpeech", BuildKeyCode(wdKeyBackSingleQuote)
    
    KeyBindings.Add wdKeyCategoryMacro, "Email.ShowEmailForm", BuildKeyCode(wdKeyCommand, wdKeyShift, wdKeyE)
    KeyBindings.Add wdKeyCategoryMacro, "Paperless.CopyToUSB", BuildKeyCode(wdKeyCommand, wdKeyShift, wdKeyS)
    KeyBindings.Add wdKeyCategoryMacro, "PaDS.PaDSPublic", BuildKeyCode(wdKeyCommand, wdKeyShift, wdKeyW)
    KeyBindings.Add wdKeyCategoryMacro, "PaDS.UploadToPaDSDummy", BuildKeyCode(wdKeyCommand, wdKeyAlt, wdKeyS)
    KeyBindings.Add wdKeyCategoryMacro, "PaDS.OpenFromPaDSDummy", BuildKeyCode(wdKeyCommand, wdKeyAlt, wdKeyO)
    KeyBindings.Add wdKeyCategoryMacro, "View.ArrangeWindows", BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyTab)
    KeyBindings.Add wdKeyCategoryMacro, "Paperless.StartTimer", BuildKeyCode(wdKeyCommand, wdKeyShift, wdKeyT)
    KeyBindings.Add wdKeyCategoryMacro, "Caselist.CiteRequest", BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyQ)
    KeyBindings.Add wdKeyCategoryMacro, "Stats.ShowStatsForm", BuildKeyCode(wdKeyCommand, wdKeyShift, wdKeyI)
    KeyBindings.Add wdKeyCategoryMacro, "View.InvisibilityOff", BuildKeyCode(wdKeyCommand, wdKeyShift, wdKeyV)
    KeyBindings.Add wdKeyCategoryMacro, "Paperless.NewSpeech", BuildKeyCode(wdKeyCommand, wdKeyShift, wdKeyN)
    
    'Save template
    ActiveDocument.AttachedTemplate.Save

    'Reset customization context
    Application.CustomizationContext = ThisDocument

End Sub

Sub RemoveKeyBindings()

    Dim k As KeyBinding
    
    For Each k In Application.KeyBindings
        k.Clear
    Next k

End Sub

'*************************************************************************************
'* MISC FUNCTIONS                                                                                                *
'*************************************************************************************

Sub LaunchWebsite(URL As String)

    Dim Script

    On Error GoTo Handler
    
    Script = "tell application ""Safari""" & vbCrLf
    Script = Script & "open location """ & URL & """" & vbCrLf
    Script = Script & "activate" & vbCrLf
    Script = Script & "end tell"
    
    #If MAC_OFFICE_VERSION >= 15 Then
        AppleScriptTask "Verbatim.scpt", "LaunchWebsite", URL
    #Else
        MacScript (Script)
    #End If
    
    Exit Sub
    
Handler:
        MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Sub OpenWordHelp()
    Help wdHelp
End Sub

Sub OpenTemplatesFolder()
    
    Dim FolderPath As String
    
    FolderPath = MacScript("return POSIX path of (path to library folder from user domain) as string")
    FolderPath = FolderPath & "Application Support/Microsoft/Office/User Templates/My Templates"
    
    #If MAC_OFFICE_VERSION >= 15 Then
        AppleScriptTask "Verbatim.scpt", "OpenFolder", Application.NormalTemplate.Path
    #Else
        MacScript ("do shell script ""open '" & FolderPath & "'""")
    #End If

End Sub

Sub QuitWord()
    Application.Quit wdPromptToSaveChanges
End Sub

Function GetVersion() As String
    GetVersion = ActiveDocument.AttachedTemplate.BuiltInDocumentProperties(wdPropertyKeywords)
End Function

