Attribute VB_Name = "frmSetupWizard"
Attribute VB_Base = "0{1040DA01-C5E2-436C-ADD0-88BFB7382A56}{E65E6809-FBE0-4BD0-822D-E72C4B3B1AC2}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub UserForm_Initialize()

    On Error GoTo Handler
    
    'Run install checks
    If Troubleshooting.InstallCheckNormal = True Then
        Me.lblInstallWarning.ForeColor = vbRed
        Me.lblInstallWarning.Caption = "WARNING - Verbatim appears to be incorrectly installed as Normal.dotm. Verbatim is not designed to be your normal template, and many features will not work correctly. Instead, you should change the filename back to Debate.dotm and use the ""Always On"" mode in the Verbatim settings." & vbCrLf & vbCrLf & "Please close Verbatim and install correctly before proceeding."
    ElseIf Troubleshooting.InstallCheckTemplateName = True Then
        Me.lblInstallWarning.ForeColor = vbRed
        Me.lblInstallWarning.Caption = "WARNING - Verbatim appears to be installed incorrectly. Verbatim should always be named ""Debate.dotm"" or many features will not work correctly. It is strongly recommended you change the file name back." & vbCrLf & vbCrLf & "Please close Verbatim and install correctly before proceeding."
    ElseIf Troubleshooting.InstallCheckTemplateLocation = True Then
        Me.lblInstallWarning.ForeColor = vbRed
        Me.lblInstallWarning.Caption = "WARNING - Verbatim appears to be installed in the wrong location. The Verbatim template file (Debate.dotm) should be located in your My Templates folder, usually located at: ~/Library/Application Support/Microsoft/Office/User Templates/My Templates. Using it from a different location will break many features." & vbCrLf & vbCrLf & "Please close Verbatim and install correctly before proceeding."
    ElseIf MacScript("do shell script ""defaults read com.microsoft.Word '" & Left(Application.Version, 2) & "\\Options\\Options:EnableMacroVirusProtection'""") = "1" Then
        Me.lblInstallWarning.ForeColor = vbRed
        Me.lblInstallWarning.Caption = "WARNING - You appear to have Macro Security enabled. This will cause Verbatim to run poorly and frequently prompt you to ""Enable Macros.""" & vbCrLf & vbCrLf & "You can disable Macro Security in Word Preferences - Security."
    Else
        Me.lblInstallWarning.Caption = "Verbatim appears to be installed correctly."
    End If
    
    'Set defaults
    Me.chkVerbatimizeNormal.Value = True
    Me.chkAlwaysOn.Value = True
    
    If GetSetting("Verbatim", "Main", "CollegeHS", "College") = "College" Then
        Me.optCollege.Value = True
    Else
        Me.optHS.Value = True
    End If
    
    Me.txtTabroomUsername.Value = GetSetting("Verbatim", "Main", "TabroomUsername", "?")
    
    If GetSetting("Verbatim", "PaDS", "PaDSSiteName", "?") <> "" And GetSetting("Verbatim", "PaDS", "PaDSSiteName", "?") <> "?" Then
        Me.optPaDSYes.Value = True
        Me.txtPaDSSiteName.Value = GetSetting("Verbatim", "PaDS", "PaDSSiteName", "?")
    Else
        Me.optPaDSNo.Value = True
    End If
    
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
    
    'Disable tutorial button if 2016
    If Application.Version >= "15" Then
        Me.chkTutorial.Value = False
        Me.chkTutorial.Enabled = False
    Else
        Me.chkTutorial.Value = True
    End If
    
    Exit Sub
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub mpgSetupWizard_Change()

    Select Case Me.mpgSetupWizard.Value
        Case Is = 0
            Me.btnBack.Visible = False
        Case Is = 1
            Me.btnBack.Visible = True
        Case Is = 5
            Me.btnNext.Caption = "Next-->"
        Case Is = 6
            Me.btnNext.Caption = "Finish"
        Case Else
            'Do nothing
    End Select
    
End Sub
Private Sub btnNext_Click()

    On Error Resume Next
    
    If Me.mpgSetupWizard.Value < 6 Then
        Me.mpgSetupWizard.Value = Me.mpgSetupWizard.Value + 1
    Else 'Finish button
        
        'Install tab
        
        'Configure tab
        If Me.chkVerbatimizeNormal.Value = True Then Call Settings.VerbatimizeNormal(Notify:=True)
        SaveSetting "Verbatim", "Admin", "AlwaysOn", Me.chkAlwaysOn.Value
        
        If Me.optCollege.Value = True Then
            SaveSetting "Verbatim", "Main", "CollegeHS", "College"
        Else
            SaveSetting "Verbatim", "Main", "CollegeHS", "HS"
        End If
        
        'Accounts tab
        SaveSetting "Verbatim", "Main", "TabroomUsername", Me.txtTabroomUsername.Value
        If Me.txtTabroomPassword.Value <> "" Then
            SaveSetting "Verbatim", "Main", "TabroomPassword", XOREncryption(Me.txtTabroomPassword.Value)
        End If
        
        'PaDS tab
        SaveSetting "Verbatim", "PaDS", "PaDSSiteName", Me.txtPaDSSiteName.Value
        SaveSetting "Verbatim", "PaDS", "PublicFolder", "http://" & Me.txtPaDSSiteName.Value & ".paperlessdebate.com/Public/"
        SaveSetting "Verbatim", "PaDS", "CoauthoringFolder", "http://" & Me.txtPaDSSiteName.Value & ".paperlessdebate.com/Team Tubs/"
        
        'Caselist tab
        If Me.optOpenCaselist.Value = True Then SaveSetting "Verbatim", "Caselist", "DefaultWiki", "openCaselist"
        If Me.optNDCAPolicy.Value = True Then SaveSetting "Verbatim", "Caselist", "DefaultWiki", "NDCAPolicy"
        If Me.optNDCALD.Value = True Then SaveSetting "Verbatim", "Caselist", "DefaultWiki", "NDCALD"
        
        SaveSetting "Verbatim", "Caselist", "CaselistSchoolName", Me.cboCaselistSchoolName.Value
        If Me.cboCaselistTeamName.Value <> "No teams found." Then SaveSetting "Verbatim", "Caselist", "CaselistTeamName", Me.cboCaselistTeamName.Value
        
        Unload Me
        
        If Me.chkTutorial.Value = True Then Call Tutorial.LaunchTutorial
        
    End If
    
    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    
End Sub

Private Sub btnBack_Click()
    Me.mpgSetupWizard.Value = Me.mpgSetupWizard.Value - 1
End Sub

Private Sub btnCancel_Click()
    If MsgBox("Are you sure you want to exit without completing the Setup Wizard?", vbYesNo) = vbYes Then Unload Me
End Sub

Private Sub lblTabroomRegister_Click()
    Settings.LaunchWebsite ("https://www.tabroom.com/user/login/new_user.mhtml")
End Sub

Private Sub optPaDSYes_Click()
    Me.lblPaDSSiteName.Visible = True
    Me.txtPaDSSiteName.Visible = True
End Sub
Private Sub optPaDSNo_Click()
    Me.lblPaDSSiteName.Visible = False
    Me.txtPaDSSiteName.Visible = False
End Sub

Private Sub lblPaDSLink_Click()
    Settings.LaunchWebsite ("http://paperlessdebate.com/pads/")
End Sub

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
