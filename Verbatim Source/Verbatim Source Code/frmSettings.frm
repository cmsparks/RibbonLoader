Attribute VB_Name = "frmSettings"
Attribute VB_Base = "0{6BC2A2A4-C5A7-4D70-9189-763E28E2D684}{94D6FFAF-AAED-4501-B3BF-840512439874}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub UserForm_Initialize()

    Dim FontSize As Integer
    Dim f
    Dim MacroArray
    
    'Turn on Error handling
    On Error GoTo Handler
    
    'Get Settings from the plist to populate the settings boxes
    
    'Main Tab
    Me.txtSchoolName.Value = GetSetting("Verbatim", "Main", "SchoolName", "?")
    Me.txtName.Value = GetSetting("Verbatim", "Main", "Name", "?")
    
    If GetSetting("Verbatim", "Main", "CollegeHS", "College") = "College" Then
        Me.optCollege.Value = True
    Else
        Me.optHS.Value = True
    End If
    
    Me.txtWPM.Value = GetSetting("Verbatim", "Main", "WPM", 350)
    
    Me.txtTabroomUsername.Value = GetSetting("Verbatim", "Main", "TabroomUsername", "?")
    
    Me.chkAutomaticUpdates.Value = GetSetting("Verbatim", "Main", "AutomaticUpdates", True)
    Me.lblLastUpdateCheck.Caption = "Last Update Check:" & vbCrLf & _
        Format(GetSetting("Verbatim", "Main", "LastUpdateCheck", "?"), "mm-dd-yy hh:mm")
    
    'Admin Tab
    Me.chkAlwaysOn.Value = GetSetting("Verbatim", "Admin", "AlwaysOn", True)
    Me.chkAutoUpdateStyles.Value = GetSetting("Verbatim", "Admin", "AutoUpdateStyles", True)
    Me.chkSuppressInstallChecks.Value = GetSetting("Verbatim", "Admin", "SuppressInstallChecks", False)
    Me.chkSuppressDocCheck.Value = GetSetting("Verbatim", "Admin", "SuppressDocCheck", False)
    Me.chkFirstRun = GetSetting("Verbatim", "Admin", "FirstRun", False)
    
    'View Tab
    If GetSetting("Verbatim", "View", "DefaultView", "Web") = "Web" Then
        Me.optWebView.Value = True
    Else
        Me.optDraftView.Value = True
    End If
    
    If GetSetting("Verbatim", "View", "ToolbarPosition", "Top") = "Top" Then
        Me.optToolbarTop.Value = True
    Else
        Me.optToolbarLeft.Value = True
    End If
    
    Me.spnDocs.Value = GetSetting("Verbatim", "View", "DocsPct", 50)
    Me.spnSpeech.Value = GetSetting("Verbatim", "View", "SpeechPct", 50)
    
    Me.spnZoomPct.Value = GetSetting("Verbatim", "View", "ZoomPct", 100)
    
    'Paperless Tab
    Me.chkAutoSaveSpeech.Value = GetSetting("Verbatim", "Paperless", "AutoSaveSpeech", False)
    Me.cboAutoSaveDir.Value = GetSetting("Verbatim", "Paperless", "AutoSaveDir", "?")
    Me.chkStripSpeech.Value = GetSetting("Verbatim", "Paperless", "StripSpeech", True)
    Me.cboSearchDir.Value = GetSetting("Verbatim", "Paperless", "SearchDir", "?")
    Me.cboAutoOpenDir.Value = GetSetting("Verbatim", "Paperless", "AutoOpenDir", "?")
    Me.cboAudioDir.Value = GetSetting("Verbatim", "Paperless", "AudioDir", "?")
    Me.cboTimerApp.Value = GetSetting("Verbatim", "Paperless", "TimerApp", "?")
      
    'Populate Format Tab Comboboxes - Allow 8pt-32pt
    FontSize = 8
    Do While FontSize < 33
        Me.cboNormalSize.AddItem FontSize
        Me.cboPocketSize.AddItem FontSize
        Me.cboHatSize.AddItem FontSize
        Me.cboBlockSize.AddItem FontSize
        Me.cboTagSize.AddItem FontSize
        Me.cboCiteSize.AddItem FontSize
        Me.cboUnderlineSize.AddItem FontSize
        Me.cboEmphasisSize.AddItem FontSize
        FontSize = FontSize + 1
    Loop
    
    'Populate Format Tab Normal Font Combobox
    For Each f In Application.FontNames
        Me.cboNormalFont.AddItem f
    Next f
    
    'Populate Format Tab Emphasis box size combobox
    Me.cboEmphasisBoxSize.AddItem "1pt"
    Me.cboEmphasisBoxSize.AddItem "1.5pt"
    Me.cboEmphasisBoxSize.AddItem "2.25pt"
    Me.cboEmphasisBoxSize.AddItem "3pt"
    
    'Format Tab
    Me.cboNormalSize.Value = GetSetting("Verbatim", "Format", "NormalSize", 11)
    Me.cboNormalFont.Value = GetSetting("Verbatim", "Format", "NormalFont", "Calibri")
    
    If GetSetting("Verbatim", "Format", "Spacing", "Wide") = "Wide" Then
        Me.optSpacingWide.Value = True
    Else
        Me.optSpacingNarrow.Value = True
    End If
    
    Me.cboPocketSize.Value = GetSetting("Verbatim", "Format", "PocketSize", 26)
    Me.cboHatSize.Value = GetSetting("Verbatim", "Format", "HatSize", 22)
    Me.cboBlockSize.Value = GetSetting("Verbatim", "Format", "BlockSize", 16)
    Me.cboTagSize.Value = GetSetting("Verbatim", "Format", "TagSize", 13)
    
    Me.cboCiteSize.Value = GetSetting("Verbatim", "Format", "CiteSize", 13)
    Me.chkUnderlineCite.Value = GetSetting("Verbatim", "Format", "UnderlineCite", False)
    
    Me.cboUnderlineSize.Value = GetSetting("Verbatim", "Format", "UnderlineSize", 11)
    Me.chkBoldUnderline.Value = GetSetting("Verbatim", "Format", "BoldUnderline", False)
    
    Me.cboEmphasisSize.Value = GetSetting("Verbatim", "Format", "EmphasisSize", 11)
    Me.chkEmphasisBold.Value = GetSetting("Verbatim", "Format", "EmphasisBold", True)
    Me.chkEmphasisItalic.Value = GetSetting("Verbatim", "Format", "EmphasisItalic", False)
    Me.chkEmphasisBox.Value = GetSetting("Verbatim", "Format", "EmphasisBox", False)
    Me.cboEmphasisBoxSize.Value = GetSetting("Verbatim", "Format", "EmphasisBoxSize", "1pt")
    
    Me.chkParagraphIntegrity.Value = GetSetting("Verbatim", "Format", "ParagraphIntegrity", False)
    Me.chkUsePilcrows.Value = GetSetting("Verbatim", "Format", "UsePilcrows", False)
    
    If GetSetting("Verbatim", "Format", "ShrinkMode", "Paragraph") = "Paragraph" Then
        Me.optParagraph.Value = True
    Else
        Me.optSelected.Value = True
    End If
    
    Me.chkAutoUnderlineEmphasis.Value = GetSetting("Verbatim", "Format", "AutoUnderlineEmphasis", False)
            
    'Populate Keyboard Tab Comboboxes
    MacroArray = Array("Paste", "Condense", "Pocket", "Hat", "Block", "Tag", "Cite", "Underline", "Emphasis", "Highlight", "Clear", "Shrink Text", "Select Similar")
    
    Me.cboF2Shortcut.List = MacroArray
    Me.cboF3Shortcut.List = MacroArray
    Me.cboF4Shortcut.List = MacroArray
    Me.cboF5Shortcut.List = MacroArray
    Me.cboF6Shortcut.List = MacroArray
    Me.cboF7Shortcut.List = MacroArray
    Me.cboF8Shortcut.List = MacroArray
    Me.cboF9Shortcut.List = MacroArray
    Me.cboF10Shortcut.List = MacroArray
    Me.cboF11Shortcut.List = MacroArray
    Me.cboF12Shortcut.List = MacroArray
    
    'Keyboard Tab
    Me.cboF2Shortcut.Value = GetSetting("Verbatim", "Keyboard", "F2Shortcut", "Paste")
    Me.cboF3Shortcut.Value = GetSetting("Verbatim", "Keyboard", "F3Shortcut", "Condense")
    Me.cboF4Shortcut.Value = GetSetting("Verbatim", "Keyboard", "F4Shortcut", "Pocket")
    Me.cboF5Shortcut.Value = GetSetting("Verbatim", "Keyboard", "F5Shortcut", "Hat")
    Me.cboF6Shortcut.Value = GetSetting("Verbatim", "Keyboard", "F6Shortcut", "Block")
    Me.cboF7Shortcut.Value = GetSetting("Verbatim", "Keyboard", "F7Shortcut", "Tag")
    Me.cboF8Shortcut.Value = GetSetting("Verbatim", "Keyboard", "F8Shortcut", "Cite")
    Me.cboF9Shortcut.Value = GetSetting("Verbatim", "Keyboard", "F9Shortcut", "Underline")
    Me.cboF10Shortcut.Value = GetSetting("Verbatim", "Keyboard", "F10Shortcut", "Emphasis")
    Me.cboF11Shortcut.Value = GetSetting("Verbatim", "Keyboard", "F11Shortcut", "Highlight")
    Me.cboF12Shortcut.Value = GetSetting("Verbatim", "Keyboard", "F12Shortcut", "Clear")
    
    'VTub Tab
    Me.cboVTubPath.Value = GetSetting("Verbatim", "VTub", "VTubPath", "?")
    Me.chkVTubRefreshPrompt.Value = GetSetting("Verbatim", "VTub", "VTubRefreshPrompt", True)
    
    'PaDS Tab
    Me.txtPaDSSiteName.Value = GetSetting("Verbatim", "PaDS", "PaDSSiteName", "?")
    Me.chkManualPaDSFolders.Value = GetSetting("Verbatim", "PaDS", "ManualPaDSFolders", False)
    Me.txtCoauthoringFolder.Value = GetSetting("Verbatim", "PaDS", "CoauthoringFolder", "?")
    Me.txtPublicFolder.Value = GetSetting("Verbatim", "PaDS", "PublicFolder", "?")
    
    'Caselist Tab
    Select Case GetSetting("Verbatim", "Caselist", "DefaultWiki", "openCaselist")
        Case Is = "openCaselist"
            Me.optOpenCaselist.Value = True
        Case Is = "NDCAPolicy"
            Me.optNDCAPolicy.Value = True
        Case Is = "NDCALD"
            Me.optNDCALD.Value = True
        Case Else
            Me.optOpenCaselist.Value = True
    End Select
    
    Me.cboCaselistSchoolName.Value = GetSetting("Verbatim", "Caselist", "CaselistSchoolName", "?")
    Me.cboCaselistTeamName.Value = GetSetting("Verbatim", "Caselist", "CaselistTeamName", "?")
    Me.txtCustomPrefixes.Value = GetSetting("Verbatim", "Caselist", "CustomPrefixes", "?")
       
    'About Tab
    Me.lblAbout2.Caption = "Verbatim Mac v. " & Settings.GetVersion
    
    'Disable incompatible settings in 2016
    If Application.Version < "15" Then
    Else
        Me.mpgSettings(6).Enabled = False
        Me.mpgSettings(7).Enabled = False
        Me.btnTutorial.Enabled = False
        Me.optToolbarLeft.Enabled = False
        Me.optToolbarTop.Enabled = False
        Me.chkAutoSaveSpeech.Enabled = False
        Me.cboAutoSaveDir.Enabled = False
        
        Me.cboAutoSaveDir.DropButtonStyle = fmDropButtonStyleArrow
        Me.cboSearchDir.DropButtonStyle = fmDropButtonStyleArrow
        Me.cboAutoOpenDir.DropButtonStyle = fmDropButtonStyleArrow
        Me.cboAudioDir.DropButtonStyle = fmDropButtonStyleArrow
        Me.cboTimerApp.DropButtonStyle = fmDropButtonStyleArrow
        Me.cboVTubPath.DropButtonStyle = fmDropButtonStyleArrow
        
    End If
    
    Exit Sub
        
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Private Sub btnResetAllSettings_Click()
'Resets all settings to the default

    On Error GoTo Handler
    
    'Prompt for confirmation
    If MsgBox("This will reset all settings to their default values - changes will not be committed until you click Save. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
    
    'Main Tab
    Me.txtSchoolName.Value = ""
    Me.txtName.Value = ""
    Me.optCollege.Value = True
    Me.txtWPM.Value = 350
    Me.txtTabroomUsername.Value = ""
    Me.txtTabroomPassword.Value = ""
    Me.chkAutomaticUpdates.Value = True
    
    'Admin Tab
    Me.chkAlwaysOn.Value = True
    Me.chkAutoUpdateStyles.Value = True
    Me.chkSuppressInstallChecks.Value = False
    Me.chkSuppressDocCheck.Value = False
    Me.chkFirstRun.Value = False
    
    'View Tab
    Me.optWebView.Value = True
    Me.optToolbarTop.Value = True
    Me.spnDocs.Value = 50
    Me.spnSpeech.Value = 50
    Me.spnZoomPct.Value = 100
    
    'Paperless Tab
    Me.chkAutoSaveSpeech.Value = False
    Me.cboAutoSaveDir.Value = ""
    Me.chkStripSpeech.Value = True
    Me.cboSearchDir.Value = ""
    Me.cboAutoOpenDir.Value = ""
    Me.cboAudioDir.Value = ""
    Me.cboTimerApp.Value = ""
    
    'Format Tab
    Me.cboNormalSize.Value = 11
    Me.cboNormalFont.Value = "Calibri"
    Me.optSpacingWide.Value = True
    Me.cboPocketSize.Value = 26
    Me.cboHatSize.Value = 22
    Me.cboBlockSize.Value = 16
    Me.cboTagSize.Value = 13
    Me.cboCiteSize.Value = 13
    Me.chkUnderlineCite.Value = False
    Me.cboUnderlineSize.Value = 11
    Me.chkBoldUnderline.Value = False
    Me.cboEmphasisSize.Value = 11
    Me.chkEmphasisBold.Value = True
    Me.chkEmphasisItalic.Value = False
    Me.chkEmphasisBox.Value = False
    Me.cboEmphasisBoxSize.Value = "1pt"
    
    Me.chkParagraphIntegrity.Value = False
    Me.chkUsePilcrows.Value = False
    Me.optParagraph.Value = True
    Me.chkAutoUnderlineEmphasis.Value = False
    
    'Keyboard Tab
    Me.cboF2Shortcut.Value = "Paste"
    Me.cboF3Shortcut.Value = "Condense"
    Me.cboF4Shortcut.Value = "Pocket"
    Me.cboF5Shortcut.Value = "Hat"
    Me.cboF6Shortcut.Value = "Block"
    Me.cboF7Shortcut.Value = "Tag"
    Me.cboF8Shortcut.Value = "Cite"
    Me.cboF9Shortcut.Value = "Underline"
    Me.cboF10Shortcut.Value = "Emphasis"
    Me.cboF11Shortcut.Value = "Highlight"
    Me.cboF12Shortcut.Value = "Clear"
    
    'VTub Tab
    Me.cboVTubPath.Value = ""
    Me.chkVTubRefreshPrompt.Value = True
    
    'PaDS Tab
    Me.txtPaDSSiteName.Value = ""
    Me.chkManualPaDSFolders.Value = False
    Me.txtCoauthoringFolder.Value = ""
    Me.txtPublicFolder.Value = ""
    
    'Caselist Tab
    Me.optOpenCaselist.Value = True
    Me.cboCaselistSchoolName.Value = ""
    Me.cboCaselistTeamName.Value = ""
    Me.txtCustomPrefixes.Value = ""
    
    'About Tab
    
    Exit Sub
        
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Private Sub btnSave_Click()
'Save Settings to Registry

    Dim Menu As CommandBarControl
    Dim c

    Dim DebateTemplate As Document
    Dim CloseDebateTemplate As Boolean
    
    'Turn on Error handling
    On Error GoTo Handler
    
    'Main Tab
    SaveSetting "Verbatim", "Main", "SchoolName", Me.txtSchoolName.Value
    SaveSetting "Verbatim", "Main", "Name", Me.txtName.Value
    
    If Me.optCollege.Value = True Then
        SaveSetting "Verbatim", "Main", "CollegeHS", "College"
    Else
        SaveSetting "Verbatim", "Main", "CollegeHS", "HS"
    End If
        
    SaveSetting "Verbatim", "Main", "WPM", Me.txtWPM.Value
    SaveSetting "Verbatim", "Main", "TabroomUsername", Trim(Me.txtTabroomUsername.Value)
    If Me.txtTabroomPassword.Value <> "" Then
        SaveSetting "Verbatim", "Main", "TabroomPassword", XOREncryption(Me.txtTabroomPassword.Value)
    End If
    SaveSetting "Verbatim", "Main", "AutomaticUpdates", Me.chkAutomaticUpdates.Value
    
    'Admin Tab
    SaveSetting "Verbatim", "Admin", "AlwaysOn", Me.chkAlwaysOn.Value
    SaveSetting "Verbatim", "Admin", "AutoUpdateStyles", Me.chkAutoUpdateStyles.Value
    SaveSetting "Verbatim", "Admin", "SuppressInstallChecks", Me.chkSuppressInstallChecks.Value
    SaveSetting "Verbatim", "Admin", "SuppressDocCheck", Me.chkSuppressDocCheck.Value
    SaveSetting "Verbatim", "Admin", "FirstRun", Me.chkFirstRun.Value
    
    'View Tab
    If Me.optWebView.Value = True Then
        SaveSetting "Verbatim", "View", "DefaultView", "Web"
    Else
        SaveSetting "Verbatim", "View", "DefaultView", "Draft"
    End If

    If Me.optToolbarTop.Value = True Then
        SaveSetting "Verbatim", "View", "ToolbarPosition", "Top"
    Else
        SaveSetting "Verbatim", "View", "ToolbarPosition", "Left"
    End If
    
    SaveSetting "Verbatim", "View", "DocsPct", Me.spnDocs.Value
    SaveSetting "Verbatim", "View", "SpeechPct", Me.spnSpeech.Value
    
    SaveSetting "Verbatim", "View", "ZoomPct", Me.spnZoomPct.Value
    
    'Paperless Tab
    SaveSetting "Verbatim", "Paperless", "AutoSaveSpeech", Me.chkAutoSaveSpeech.Value
    SaveSetting "Verbatim", "Paperless", "AutoSaveDir", Me.cboAutoSaveDir.Value
    SaveSetting "Verbatim", "Paperless", "StripSpeech", Me.chkStripSpeech.Value
    SaveSetting "Verbatim", "Paperless", "SearchDir", Me.cboSearchDir.Value
    SaveSetting "Verbatim", "Paperless", "AutoOpenDir", Me.cboAutoOpenDir.Value
    SaveSetting "Verbatim", "Paperless", "AudioDir", Me.cboAudioDir.Value
    SaveSetting "Verbatim", "Paperless", "TimerApp", Me.cboTimerApp.Value
    
    'Format Tab
    SaveSetting "Verbatim", "Format", "NormalSize", Me.cboNormalSize.Value
    SaveSetting "Verbatim", "Format", "NormalFont", Me.cboNormalFont.Value
    
    If Me.optSpacingWide.Value = True Then
        SaveSetting "Verbatim", "Format", "Spacing", "Wide"
    Else
        SaveSetting "Verbatim", "Format", "Spacing", "Narrow"
    End If
    
    SaveSetting "Verbatim", "Format", "PocketSize", Me.cboPocketSize.Value
    SaveSetting "Verbatim", "Format", "HatSize", Me.cboHatSize.Value
    SaveSetting "Verbatim", "Format", "BlockSize", Me.cboBlockSize.Value
    SaveSetting "Verbatim", "Format", "TagSize", Me.cboTagSize.Value
    SaveSetting "Verbatim", "Format", "CiteSize", Me.cboCiteSize.Value
    SaveSetting "Verbatim", "Format", "UnderlineCite", Me.chkUnderlineCite.Value
    SaveSetting "Verbatim", "Format", "UnderlineSize", Me.cboUnderlineSize.Value
    SaveSetting "Verbatim", "Format", "BoldUnderline", Me.chkBoldUnderline.Value
    SaveSetting "Verbatim", "Format", "EmphasisSize", Me.cboEmphasisSize.Value
    SaveSetting "Verbatim", "Format", "EmphasisBold", Me.chkEmphasisBold.Value
    SaveSetting "Verbatim", "Format", "EmphasisItalic", Me.chkEmphasisItalic.Value
    SaveSetting "Verbatim", "Format", "EmphasisBox", Me.chkEmphasisBox.Value
    SaveSetting "Verbatim", "Format", "EmphasisBoxSize", Me.cboEmphasisBoxSize.Value
    SaveSetting "Verbatim", "Format", "ParagraphIntegrity", Me.chkParagraphIntegrity.Value
    SaveSetting "Verbatim", "Format", "UsePilcrows", Me.chkUsePilcrows.Value
    
    If Me.optParagraph.Value = True Then
        SaveSetting "Verbatim", "Format", "ShrinkMode", "Paragraph"
    Else
        SaveSetting "Verbatim", "Format", "ShrinkMode", "Selected"
    End If
    
    SaveSetting "Verbatim", "Format", "AutoUnderlineEmphasis", Me.chkAutoUnderlineEmphasis.Value
    
    'Check if Template itself is open, or open it as a Document
    If ActiveDocument.FullName = ActiveDocument.AttachedTemplate.FullName Then
        Set DebateTemplate = ActiveDocument
        CloseDebateTemplate = False
    Else
        Set DebateTemplate = ActiveDocument.AttachedTemplate.OpenAsDocument
        CloseDebateTemplate = True
    End If
    
    'Update template styles based on Format settings
    DebateTemplate.Styles("Normal").Font.Size = Me.cboNormalSize.Value
    DebateTemplate.Styles("Normal").Font.Name = Me.cboNormalFont.Value
    
    If Me.optSpacingWide.Value = True Then
        DebateTemplate.Styles("Normal").ParagraphFormat.SpaceBefore = 0
        DebateTemplate.Styles("Normal").ParagraphFormat.SpaceAfter = 8
        DebateTemplate.Styles("Normal").ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
        DebateTemplate.Styles("Normal").ParagraphFormat.LineSpacing = LinesToPoints(1.08)
        DebateTemplate.Styles("Pocket").ParagraphFormat.SpaceBefore = 12
        DebateTemplate.Styles("Pocket").ParagraphFormat.SpaceAfter = 0
        DebateTemplate.Styles("Hat").ParagraphFormat.SpaceBefore = 2
        DebateTemplate.Styles("Hat").ParagraphFormat.SpaceAfter = 0
        DebateTemplate.Styles("Block").ParagraphFormat.SpaceBefore = 2
        DebateTemplate.Styles("Block").ParagraphFormat.SpaceAfter = 0
        DebateTemplate.Styles("Tag").ParagraphFormat.SpaceBefore = 2
        DebateTemplate.Styles("Tag").ParagraphFormat.SpaceAfter = 0
    Else
        DebateTemplate.Styles("Normal").ParagraphFormat.SpaceBefore = 0
        DebateTemplate.Styles("Normal").ParagraphFormat.SpaceAfter = 0
        DebateTemplate.Styles("Normal").ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
        DebateTemplate.Styles("Pocket").ParagraphFormat.SpaceBefore = 24
        DebateTemplate.Styles("Pocket").ParagraphFormat.SpaceAfter = 0
        DebateTemplate.Styles("Hat").ParagraphFormat.SpaceBefore = 24
        DebateTemplate.Styles("Hat").ParagraphFormat.SpaceAfter = 0
        DebateTemplate.Styles("Block").ParagraphFormat.SpaceBefore = 10
        DebateTemplate.Styles("Block").ParagraphFormat.SpaceAfter = 0
        DebateTemplate.Styles("Tag").ParagraphFormat.SpaceBefore = 10
        DebateTemplate.Styles("Tag").ParagraphFormat.SpaceAfter = 0
    End If
    
    DebateTemplate.Styles("Pocket").Font.Size = Me.cboPocketSize.Value
    DebateTemplate.Styles("Hat").Font.Size = Me.cboHatSize.Value
    DebateTemplate.Styles("Block").Font.Size = Me.cboBlockSize.Value
    DebateTemplate.Styles("Tag").Font.Size = Me.cboTagSize.Value
    DebateTemplate.Styles("Cite").Font.Size = Me.cboCiteSize.Value
    If Me.chkUnderlineCite.Value = True Then
        DebateTemplate.Styles("Cite").Font.Underline = wdUnderlineSingle
    Else
        DebateTemplate.Styles("Cite").Font.Underline = wdUnderlineNone
    End If
    DebateTemplate.Styles("Underline").Font.Size = Me.cboUnderlineSize.Value
    If Me.chkBoldUnderline.Value = True Then
        DebateTemplate.Styles("Underline").Font.Bold = True
    Else
        DebateTemplate.Styles("Underline").Font.Bold = False
    End If
    DebateTemplate.Styles("Emphasis").Font.Size = Me.cboEmphasisSize.Value
    DebateTemplate.Styles("Emphasis").Font.Name = Me.cboNormalFont.Value
    DebateTemplate.Styles("Emphasis").Font.Bold = Me.chkEmphasisBold.Value
    DebateTemplate.Styles("Emphasis").Font.Italic = Me.chkEmphasisItalic.Value
    
    If Me.chkEmphasisBox.Value = True Then
        DebateTemplate.Styles("Emphasis").Font.Borders(1).LineStyle = wdLineStyleSingle
    
        Select Case Me.cboEmphasisBoxSize.Value
            Case Is = "1pt"
                DebateTemplate.Styles("Emphasis").Font.Borders(1).LineWidth = wdLineWidth100pt
            Case Is = "1.5pt"
                DebateTemplate.Styles("Emphasis").Font.Borders(1).LineWidth = wdLineWidth150pt
            Case Is = "2.25pt"
                DebateTemplate.Styles("Emphasis").Font.Borders(1).LineWidth = wdLineWidth225pt
            Case Is = "3pt"
                DebateTemplate.Styles("Emphasis").Font.Borders(1).LineWidth = wdLineWidth300pt
            Case Else
                DebateTemplate.Styles("Emphasis").Font.Borders(1).LineWidth = wdLineWidth100pt
        End Select
    Else
        DebateTemplate.Styles("Emphasis").Font.Borders(1).LineStyle = wdLineStyleNone
    End If
    
    'Keyboard Tab
    SaveSetting "Verbatim", "Keyboard", "F2Shortcut", Me.cboF2Shortcut.Value
    SaveSetting "Verbatim", "Keyboard", "F3Shortcut", Me.cboF3Shortcut.Value
    SaveSetting "Verbatim", "Keyboard", "F4Shortcut", Me.cboF4Shortcut.Value
    SaveSetting "Verbatim", "Keyboard", "F5Shortcut", Me.cboF5Shortcut.Value
    SaveSetting "Verbatim", "Keyboard", "F6Shortcut", Me.cboF6Shortcut.Value
    SaveSetting "Verbatim", "Keyboard", "F7Shortcut", Me.cboF7Shortcut.Value
    SaveSetting "Verbatim", "Keyboard", "F8Shortcut", Me.cboF8Shortcut.Value
    SaveSetting "Verbatim", "Keyboard", "F9Shortcut", Me.cboF9Shortcut.Value
    SaveSetting "Verbatim", "Keyboard", "F10Shortcut", Me.cboF10Shortcut.Value
    SaveSetting "Verbatim", "Keyboard", "F11Shortcut", Me.cboF11Shortcut.Value
    SaveSetting "Verbatim", "Keyboard", "F12Shortcut", Me.cboF12Shortcut.Value
    
    'Update template keyboard shortcuts based on keyboard settings
    Call Settings.ChangeKeyboardShortcut(wdKeyF2, Me.cboF2Shortcut.Value)
    Call Settings.ChangeKeyboardShortcut(wdKeyF3, Me.cboF3Shortcut.Value)
    Call Settings.ChangeKeyboardShortcut(wdKeyF4, Me.cboF4Shortcut.Value)
    Call Settings.ChangeKeyboardShortcut(wdKeyF5, Me.cboF5Shortcut.Value)
    Call Settings.ChangeKeyboardShortcut(wdKeyF6, Me.cboF6Shortcut.Value)
    Call Settings.ChangeKeyboardShortcut(wdKeyF7, Me.cboF7Shortcut.Value)
    Call Settings.ChangeKeyboardShortcut(wdKeyF8, Me.cboF8Shortcut.Value)
    Call Settings.ChangeKeyboardShortcut(wdKeyF9, Me.cboF9Shortcut.Value)
    Call Settings.ChangeKeyboardShortcut(wdKeyF10, Me.cboF10Shortcut.Value)
    Call Settings.ChangeKeyboardShortcut(wdKeyF11, Me.cboF11Shortcut.Value)
    Call Settings.ChangeKeyboardShortcut(wdKeyF12, Me.cboF12Shortcut.Value)
    
    'Close template if opened separately
    If CloseDebateTemplate = True Then
        DebateTemplate.Close SaveChanges:=wdSaveChanges
    End If
    
    ActiveDocument.UpdateStyles
    
    'VTub Tab
    SaveSetting "Verbatim", "VTub", "VTubPath", Me.cboVTubPath.Value
    SaveSetting "Verbatim", "VTub", "VTubRefreshPrompt", chkVTubRefreshPrompt.Value
    
    'PaDS Tab
    SaveSetting "Verbatim", "PaDS", "PaDSSiteName", Me.txtPaDSSiteName.Value
    SaveSetting "Verbatim", "PaDS", "ManualPaDSFolders", Me.chkManualPaDSFolders.Value
    SaveSetting "Verbatim", "PaDS", "CoauthoringFolder", Me.txtCoauthoringFolder.Value
    SaveSetting "Verbatim", "PaDS", "PublicFolder", Me.txtPublicFolder.Value
    Call PaDS.ClearPaDSCookie 'Delete cookie in case credentials changed
    
    'Caselist Tab
    If Me.optOpenCaselist.Value = True Then SaveSetting "Verbatim", "Caselist", "DefaultWiki", "openCaselist"
    If Me.optNDCAPolicy.Value = True Then SaveSetting "Verbatim", "Caselist", "DefaultWiki", "NDCAPolicy"
    If Me.optNDCALD.Value = True Then SaveSetting "Verbatim", "Caselist", "DefaultWiki", "NDCALD"
    
    SaveSetting "Verbatim", "Caselist", "CaselistSchoolName", Me.cboCaselistSchoolName.Value
    If Me.cboCaselistTeamName.Value <> "No teams found." Then SaveSetting "Verbatim", "Caselist", "CaselistTeamName", Me.cboCaselistTeamName.Value
    SaveSetting "Verbatim", "Caselist", "CustomPrefixes", Me.txtCustomPrefixes.Value
    
    'About Tab
    SaveSetting "Verbatim", "Main", "Version", Settings.GetVersion
    
    'If Word 2011, update toolbar
    If Application.Version < "15" Then
    
        'Reset F Key button captions
    
        CommandBars.FindControl(Tag:="F2Button").Caption = Me.cboF2Shortcut.Value
        CommandBars.FindControl(Tag:="F3Button").Caption = Me.cboF3Shortcut.Value
        CommandBars.FindControl(Tag:="F4Button").Caption = Me.cboF4Shortcut.Value
        CommandBars.FindControl(Tag:="F5Button").Caption = Me.cboF5Shortcut.Value
        CommandBars.FindControl(Tag:="F6Button").Caption = Me.cboF6Shortcut.Value
        CommandBars.FindControl(Tag:="F7Button").Caption = Me.cboF7Shortcut.Value
        CommandBars.FindControl(Tag:="F8Button").Caption = Me.cboF8Shortcut.Value
        CommandBars.FindControl(Tag:="F9Button").Caption = Me.cboF9Shortcut.Value
        CommandBars.FindControl(Tag:="F10Button").Caption = Me.cboF10Shortcut.Value
        CommandBars.FindControl(Tag:="F11Button").Caption = Me.cboF11Shortcut.Value
        CommandBars.FindControl(Tag:="F12Button").Caption = Me.cboF12Shortcut.Value
    
    
        'Clear dynamic menus
        Set Menu = CommandBars.FindControl(Tag:="NewSpeechMenu")
        For Each c In Menu.Controls
            c.Delete
        Next c
        
        Set Menu = CommandBars.FindControl(Tag:="VirtualTub")
        For Each c In Menu.Controls
            c.Delete
        Next c
        
        Set Menu = CommandBars.FindControl(Tag:="CoauthoringMenu")
        For Each c In Menu.Controls
            c.Delete
        Next c
    End If
    
    'Tell user to restart
    MsgBox "Settings saved. You may need to completely close and restart Word for them to take effect."
    
    'Clean up
    Set Menu = Nothing
    
    'Unload the form
    Unload Me
    Exit Sub
        
Handler:
    Set Menu = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Private Sub btnCancel_Click()
  Unload Me
End Sub

'*************************************************************************************
'* MAIN TAB                                                                                                            *
'*************************************************************************************

Private Sub lblWPMLink_Click()
    Settings.LaunchWebsite ("http://www.readingsoft.com/")
End Sub

Private Sub lblTabroomRegister_Click()
    Settings.LaunchWebsite ("https://www.tabroom.com/user/login/new_user.mhtml")
End Sub

Private Sub btnUpdateCheck_Click()
    Settings.UpdateCheck (True)
End Sub

'*************************************************************************************
'* ADMIN TAB                                                                                                          *
'*************************************************************************************

Private Sub btnVerbatimizeNormal_Click()
    Call Settings.VerbatimizeNormal(Notify:=True)
End Sub

Private Sub btnUnverbatimizeNormal_Click()
    Call Settings.UnverbatimizeNormal(Notify:=False)
End Sub

Private Sub btnTemplatesFolder_Click()
    Settings.OpenTemplatesFolder
End Sub

Private Sub btnTutorial_Click()
    Unload Me
    Call Tutorial.LaunchTutorial
End Sub

Private Sub btnSetupWizard_Click()
    Unload Me
    Settings.ShowSetupWizard
End Sub

Private Sub btnTroubleshooter_Click()
    Unload Me
    Call Settings.ShowTroubleshooter
End Sub

Private Sub btnImportSettings_Click()

    Dim PlistFile As String
    Dim PlistLockFile As String
    Dim SettingsFileName As String
    Dim FilePath As String

    'Turn on Error Handling
    On Error GoTo Handler

    'Get path to Verbatim plist file
    PlistFile = MacScript("return (path to preferences folder) as string") & "Verbatim.plist"
    PlistLockFile = PlistFile & ".lockfile"

    'Create MacScript for picking file
    FilePath = MacScript("set FilePath to (choose file with prompt ""Choose Verbatim settings (.plist) file:"") as string")
    
    On Error Resume Next
    
    'Delete settings
    DeleteSetting "Verbatim", "Format"
    
    'Delete old files
    Kill PlistFile
    Kill PlistLockFile
    On Error GoTo Handler
    
    'Copy new file to the Preferences folder
    FileCopy Source:=FilePath, Destination:=PlistFile

    'Report Success
    MsgBox "Settings successfully imported. To apply them, you must:" & vbCrLf & "a) Completely close and restart Word" & vbCrLf & "b) Open the Verbatim Settings and click ""Save""" & vbCrLf & "c) Completely close and restart Word again"
        
    Unload Me
    
    Exit Sub
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Private Sub btnExportSettings_Click()

    Dim PlistFile As String
    Dim SettingsFileName As String
    Dim RootFolder As String
    Dim FolderPath As String
    Dim ExportFolder As String
    
    'Turn on error checking
    On Error GoTo Handler
    
    'Get path to Verbatim plist file
    PlistFile = MacScript("return (path to preferences folder) as string") & "Verbatim.plist"
    
    'Create SettingsFile name
    SettingsFileName = "VerbatimSettings"
    If txtSchoolName.Value <> "" Then
        SettingsFileName = SettingsFileName & " - " & txtSchoolName.Value
    End If
    If txtName.Value <> "" Then
        SettingsFileName = SettingsFileName & " - " & txtName.Value
    End If
    SettingsFileName = SettingsFileName & ".plist"
    
    'Select folder for export
    RootFolder = MacScript("return (path to desktop folder) as String")
    FolderPath = MacScript("(choose folder with prompt ""Select the folder to export to...""" & "default location alias """ & RootFolder & """) as string")
    On Error GoTo 0

    If FolderPath <> "" Then
        SettingsFileName = FolderPath & SettingsFileName
    Else
        SettingsFileName = RootFolder & SettingsFileName
    End If

    'Copy Verbatim plist file
    FileCopy Source:=PlistFile, Destination:=SettingsFileName
    
    'Report success
    MsgBox "Settings successfully exported as:" & vbCrLf & SettingsFileName
    
    Exit Sub

Handler:
    If Err.Number = 5 Then Exit Sub
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Private Sub btnImportCustomCode_Click()
    Call Settings.ImportCustomCode(Notify:=True)
End Sub

Private Sub btnExportCustomCode_Click()
    Call Settings.ExportCustomCode(Notify:=True)
End Sub

'*************************************************************************************
'* VIEW TAB                                                                                                             *
'*************************************************************************************

Private Sub spnDocs_Change()
    Me.txtDocPct.Value = Me.spnDocs.Value
    Me.lblDocs.Width = 200 * Me.spnDocs.Value / 100
    Me.lblSpeech.Width = (200 * Me.spnSpeech.Value / 100)
    Me.lblSpeech.Left = 200 - Me.lblSpeech.Width
End Sub

Private Sub spnSpeech_Change()
    Me.txtSpeechPct.Value = Me.spnSpeech.Value
    Me.lblDocs.Width = 200 * Me.spnDocs.Value / 100
    Me.lblSpeech.Width = (200 * Me.spnSpeech.Value / 100)
    Me.lblSpeech.Left = 200 - Me.lblSpeech.Width
End Sub

Private Sub txtDocPct_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If IsNumeric(Me.txtDocPct.Value) Then Me.spnDocs.Value = Me.txtDocPct.Value
End Sub

Private Sub txtSpeechPct_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If IsNumeric(Me.txtSpeechPct.Value) Then Me.spnSpeech.Value = Me.txtSpeechPct.Value
End Sub

Private Sub spnZoomPct_Change()
    Me.txtZoomPct.Value = Me.spnZoomPct.Value
End Sub

Private Sub txtZoomPct_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Not IsNumeric(Me.txtZoomPct.Value) Then
        MsgBox "You must input a number between 10 and 500."
        Exit Sub
    Else
        If Me.txtZoomPct.Value < 10 Then
            Me.spnZoomPct.Value = 10
        ElseIf Me.txtZoomPct.Value > 500 Then
            Me.spnZoomPct.Value = 500
        Else
            Me.spnZoomPct.Value = Me.txtZoomPct.Value
        End If
    End If
End Sub

Private Sub btnResetView_Click()

    If MsgBox("This will reset view settings to their default values. Changes will not be committed until you click Save. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
    
    Me.optWebView.Value = True
    Me.optToolbarTop.Value = True
    Me.spnDocs.Value = 50
    Me.spnSpeech.Value = 50
    Me.spnZoomPct.Value = 100
End Sub

'*************************************************************************************
'* PAPERLESS TAB                                                                                                    *
'*************************************************************************************

Private Sub chkAutoSaveSpeech_Change()
    If Me.chkAutoSaveSpeech.Value = True Then
        Me.cboAutoSaveDir.Enabled = True
        Me.lblAutoSaveDir.Enabled = True
    Else
        Me.cboAutoSaveDir.Enabled = False
        Me.lblAutoSaveDir.Enabled = False
    End If
End Sub

Private Sub cboAutoSaveDir_DropButtonClick()

    Dim RootFolder As String
    Dim FolderPath As String

    On Error Resume Next
    
    RootFolder = MacScript("return (path to desktop folder) as String")
    FolderPath = MacScript("(choose folder with prompt ""Select the folder""" & "default location alias """ & RootFolder & """) as string")
    On Error GoTo 0
    
    If FolderPath <> "" Then
        Me.cboAutoSaveDir.Value = FolderPath
    End If

    'Reset focus to avoid getting stuck
    Me.btnCancel.SetFocus

End Sub
Private Sub cboSearchDir_DropButtonClick()

    Dim RootFolder As String
    Dim FolderPath As String

    On Error Resume Next
    
    RootFolder = MacScript("return (path to desktop folder) as String")
    FolderPath = MacScript("(choose folder with prompt ""Select the folder""" & "default location alias """ & RootFolder & """) as string")
    On Error GoTo 0
    
    If FolderPath <> "" Then
        Me.cboSearchDir.Value = FolderPath
    End If

    'Reset focus to avoid getting stuck
    Me.btnCancel.SetFocus
    
End Sub
Private Sub cboAutoOpenDir_DropButtonClick()

    Dim RootFolder As String
    Dim FolderPath As String

    On Error Resume Next
    
    RootFolder = MacScript("return (path to desktop folder) as String")
    FolderPath = MacScript("(choose folder with prompt ""Select the folder""" & "default location alias """ & RootFolder & """) as string")
    On Error GoTo 0
    
    If FolderPath <> "" Then
        Me.cboAutoOpenDir.Value = FolderPath
    End If
    
    'Reset focus to avoid getting stuck
    Me.btnCancel.SetFocus
    
End Sub

Private Sub cboAudioDir_DropButtonClick()

    Dim RootFolder As String
    Dim FolderPath As String

    On Error Resume Next
    
    RootFolder = MacScript("return (path to desktop folder) as String")
    FolderPath = MacScript("(choose folder with prompt ""Select the folder""" & "default location alias """ & RootFolder & """) as string")
    On Error GoTo 0
    
    If FolderPath <> "" Then
        Me.cboAudioDir.Value = FolderPath
    End If

    'Reset focus to avoid getting stuck
    Me.btnCancel.SetFocus

End Sub

Private Sub cboTimerApp_DropButtonClick()

    Dim RootFolder As String
    Dim AppPath As String

    On Error Resume Next
    
    RootFolder = MacScript("return (path to applications folder) as String")
    AppPath = MacScript("(choose file with prompt ""Select the timer application""" & "default location alias """ & RootFolder & """) as string")
    On Error GoTo 0
    
    If AppPath <> "" Then
        Me.cboTimerApp.Value = AppPath
    End If

    'Reset focus to avoid getting stuck
    Me.btnCancel.SetFocus

End Sub

'*************************************************************************************
'* FORMAT TAB                                                                                                       *
'*************************************************************************************

Private Sub cboNormalFont_Change()
    'Changes the font sample
    Me.lblFontSample2.Font.Name = Me.cboNormalFont.Value
End Sub

Private Sub chkEmphasisBox_Change()
    If Me.chkEmphasisBox.Value = True Then
        Me.cboEmphasisBoxSize.Enabled = True
    Else
        Me.cboEmphasisBoxSize.Enabled = False
    End If

End Sub

Private Sub chkParagraphIntegrity_Change()
    'Disable Pilcrows button if unchecked
    If Me.chkParagraphIntegrity.Value = False Then
        Me.chkUsePilcrows.Enabled = False
    Else
        Me.chkUsePilcrows.Enabled = True
    End If

End Sub

Private Sub btnResetFormatting_Click()
'Resets formatting settings to the default
    
    On Error GoTo Handler
    
    'Prompt for confirmation
    If MsgBox("This will reset formatting settings to their default values. Changes will not be committed until you click Save. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
    
    'Format Tab
    Me.cboNormalSize.Value = 11
    Me.cboNormalFont.Value = "Calibri"
    Me.optSpacingWide.Value = True
    Me.cboPocketSize.Value = 26
    Me.cboHatSize.Value = 22
    Me.cboBlockSize.Value = 16
    Me.cboTagSize.Value = 13
    Me.cboCiteSize.Value = 13
    Me.chkUnderlineCite.Value = False
    Me.cboUnderlineSize.Value = 11
    Me.chkBoldUnderline.Value = False
    Me.cboEmphasisSize.Value = 11
    Me.chkEmphasisBold.Value = True
    Me.chkEmphasisItalic.Value = False
    Me.chkEmphasisBox.Value = False
    Me.cboEmphasisBoxSize.Value = "1pt"
    
    Me.chkParagraphIntegrity.Value = False
    Me.chkUsePilcrows.Value = False
    Me.optParagraph.Value = True
    
    Exit Sub
        
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

'*************************************************************************************
'* KEYBOARD TAB                                                                                                    *
'*************************************************************************************

Private Sub btnOtherKeyboardShortcuts_Click()
    'Shows the Customize Keyboard dialogue
    Dialogs(wdDialogToolsCustomizeKeyboard).Show
End Sub

Private Sub btnResetKeyboard_Click()
'Resets keyboard settings to the default
    
    On Error GoTo Handler
    
    'Prompt for confirmation
    If MsgBox("This will reset keyboard settings to their default values. Changes will not be committed until you click Save. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
    
    'Keyboard Tab
    Me.cboF2Shortcut.Value = "Paste"
    Me.cboF3Shortcut.Value = "Condense"
    Me.cboF4Shortcut.Value = "Pocket"
    Me.cboF5Shortcut.Value = "Hat"
    Me.cboF6Shortcut.Value = "Block"
    Me.cboF7Shortcut.Value = "Tag"
    Me.cboF8Shortcut.Value = "Cite"
    Me.cboF9Shortcut.Value = "Underline"
    Me.cboF10Shortcut.Value = "Emphasis"
    Me.cboF11Shortcut.Value = "Highlight"
    Me.cboF12Shortcut.Value = "Clear"
    
    Exit Sub
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

'*************************************************************************************
'* VTUB TAB                                                                                                            *
'*************************************************************************************

Private Sub cboVTubPath_DropButtonClick()
    
    Dim RootFolder As String
    Dim FolderPath As String

    On Error Resume Next
    
    RootFolder = MacScript("return (path to desktop folder) as String")
    FolderPath = MacScript("(choose folder with prompt ""Select the folder""" & "default location alias """ & RootFolder & """) as string")
    On Error GoTo 0
    
    If FolderPath <> "" Then
        Me.cboVTubPath.Value = FolderPath
    End If
    
    'Save immediately so the Create button can find a path
    SaveSetting "Verbatim", "VTub", "VTubPath", Me.cboVTubPath.Value
    
    'Reset focus to avoid getting stuck
    Me.btnCancel.SetFocus
    
End Sub

Private Sub btnCreateVTub_Click()
    If Me.cboVTubPath.Value = "" Then
        MsgBox "You must select a path for the VTub first."
        Exit Sub
    Else
        Me.Hide
        Call VirtualTub.VTubCreate
        Me.Show
    End If
End Sub

'*************************************************************************************
'* PADS TAB                                                                                                            *
'*************************************************************************************

Private Sub txtPaDSSiteName_Change()
    If Me.chkManualPaDSFolders.Value = False Then
        Me.txtCoauthoringFolder.Value = "http://" & Me.txtPaDSSiteName.Value & ".paperlessdebate.com/Team Tubs/"
        Me.txtPublicFolder.Value = "http://" & Me.txtPaDSSiteName.Value & ".paperlessdebate.com/Public/"
    End If
End Sub

Private Sub chkManualPaDSFolders_Click()
    If Me.chkManualPaDSFolders.Value = True Then
        Me.lblCoauthoringFolder.Enabled = True
        Me.txtCoauthoringFolder.Enabled = True
        Me.lblPublicFolder.Enabled = True
        Me.txtPublicFolder.Enabled = True
        Me.txtPaDSSiteName.Enabled = False
    Else
        Me.lblCoauthoringFolder.Enabled = False
        Me.txtCoauthoringFolder.Enabled = False
        Me.lblPublicFolder.Enabled = False
        Me.txtPublicFolder.Enabled = False
        Me.txtPaDSSiteName.Enabled = True
    End If
End Sub

Private Sub lblPaDSLink_Click()
    Settings.LaunchWebsite ("http://paperlessdebate.com/pads/")
End Sub

Private Sub btnClearPaDSCookie_Click()
    Call PaDS.ClearPaDSCookie
End Sub

'*************************************************************************************
'* CASELIST TAB                                                                                                      *
'*************************************************************************************

Private Sub optOpenCaselist_Change()
    Me.cboCaselistSchoolName.Value = ""
    Me.cboCaselistSchoolName.Clear
    Me.cboCaselistTeamName.Value = ""
    Me.cboCaselistTeamName.Clear
End Sub

Private Sub optNDCAPolicy_Change()
    Me.cboCaselistSchoolName.Value = ""
    Me.cboCaselistSchoolName.Clear
    Me.cboCaselistTeamName.Value = ""
    Me.cboCaselistTeamName.Clear
End Sub

Private Sub optNDCALD_Change()
    Me.cboCaselistSchoolName.Value = ""
    Me.cboCaselistSchoolName.Clear
    Me.cboCaselistTeamName.Value = ""
    Me.cboCaselistTeamName.Clear
End Sub

Private Sub cboCaselistSchoolName_Change()
    Me.cboCaselistTeamName.Value = ""
    Me.cboCaselistTeamName.Clear
End Sub

Private Sub cboCaselistSchoolName_DropButtonClick()
'Populates the SchoolName combo box with schools from the caselist
        
    'If the list is already populated, exit
    If Me.cboCaselistSchoolName.ListCount > 0 Then Exit Sub
        
    'Clear ComboBoxes - clear TeamName too, so there's not a mismatch when changing
    Me.cboCaselistSchoolName.Value = ""
    Me.cboCaselistTeamName.Value = ""
    Me.cboCaselistSchoolName.Clear
    Me.cboCaselistTeamName.Clear
        
    'Populate box
    If Me.optOpenCaselist.Value = True Then Call GetCaselistSchoolNames("openCaselist", Me.cboCaselistSchoolName)
    If Me.optNDCAPolicy.Value = True Then Call GetCaselistSchoolNames("NDCAPolicy", Me.cboCaselistSchoolName)
    If Me.optNDCALD.Value = True Then Call GetCaselistSchoolNames("NDCALD", Me.cboCaselistSchoolName)
    
    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Private Sub cboCaselistTeamName_DropButtonClick()
'Populates the TeamName combo box with pages from the school's space

    'If the list is already populated, exit
    If Me.cboCaselistTeamName.ListCount > 0 Then Exit Sub
    
    'Check CaselistSchoolName has a value
    If Me.cboCaselistSchoolName.Value = "" Then
        Me.cboCaselistTeamName.Value = "Please choose a school first"
        Me.cboCaselistTeamName.Clear
        Exit Sub
    End If
    
    'Clear ComboBox
    Me.cboCaselistTeamName.Value = ""
    Me.cboCaselistTeamName.Clear
    
    'Turn on error checking
    On Error GoTo Handler
  
    If Me.optOpenCaselist.Value = True Then Call GetCaselistTeamNames("openCaselist", Me.cboCaselistSchoolName.Value, Me.cboCaselistTeamName)
    If Me.optNDCAPolicy.Value = True Then Call GetCaselistTeamNames("NDCAPolicy", Me.cboCaselistSchoolName.Value, Me.cboCaselistTeamName)
    If Me.optNDCALD.Value = True Then Call GetCaselistTeamNames("NDCALD", Me.cboCaselistSchoolName.Value, Me.cboCaselistTeamName)
    
    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

'*************************************************************************************
'* ABOUT TAB                                                                                                          *
'*************************************************************************************

Private Sub lblAbout5_Click()
    Settings.LaunchWebsite ("http://paperlessdebate.com/")
End Sub

Private Sub btnVerbatimHelp_Click()
    Settings.ShowVerbatimHelp
End Sub
