Attribute VB_Name = "frmTroubleshooter"
Attribute VB_Base = "0{A9FDE364-A8C6-45E9-9B2A-A205149F9C0E}{DF20329B-0B62-4846-9D82-0A03F8C2F700}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub UserForm_Activate()

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
        #If MAC_OFFICE_VERSION >= 15 Then
            Me.lblInstallWarning.Caption = "WARNING - Verbatim appears to be installed in the wrong location. The Verbatim template file (Debate.dotm) should be located in your Word Templates folder, usually located at:" & vbCrLf & "~/Library/Group Containers/UBF8T346G9.Office/User Content/Templates" & vbCrLf & "Using it from a different location will break many features." & vbCrLf & vbCrLf & "Please close Verbatim and install correctly before proceeding."
        #Else
            Me.lblInstallWarning.Caption = "WARNING - Verbatim appears to be installed in the wrong location. The Verbatim template file (Debate.dotm) should be located in your Word Templates folder, usually located at:" & vbCrLf & "~/Library/Application Support/Microsoft/Office/User Templates/My Templates" & vbCrLf & "Using it from a different location will break many features." & vbCrLf & vbCrLf & "Please close Verbatim and install correctly before proceeding."
        #End If
    Else
        Me.lblInstallWarning.ForeColor = vbBlack
        Me.lblInstallWarning.Caption = "Verbatim appears to be installed correctly."
    End If
    
    'Run rest of checks
    Call CheckMacroSecurity
    Call CheckDuplicateTemplates
    
    Exit Sub
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub imgInfoMacroSecurity_Click()
    MsgBox "Verbatim works best when your Macro Security Settings are turned off. This will prevent Word from prompting you to enable macros each time you open Word. Macro Security settings can be changed in Word Preferences - Security.", vbInformation
End Sub
Private Sub imgInfoDuplicateTemplates_Click()
    MsgBox "Many people inadvertently leave copies of the Verbatim template in other locations on their computer, such as the Desktop or Downloads folder. This can cause difficulties with file interoperability. It's best if you delete all copies of Debate.dotm except the one in your Templates folder.", vbInformation
End Sub

Private Sub CheckMacroSecurity()
       
    On Error Resume Next
    
    #If MAC_OFFICE_VERSION >= 15 Then
        If AppleScriptTask("Verbatim.scpt", "CheckMacroSecurity", "") = "0" Then
    #Else
        If MacScript("do shell script ""defaults read com.microsoft.Word '" & Left(Application.Version, 2) & "\\Options\\Options:EnableMacroVirusProtection'""") = "0" Then
    #End If
        Me.lblMacroSecurity.ForeColor = vbBlack
        Me.imgYesMacroSecurity.Visible = True
        Me.imgNoMacroSecurity.Visible = False
    Else
        Me.lblMacroSecurity.ForeColor = vbRed
        Me.imgYesMacroSecurity.Visible = False
        Me.imgNoMacroSecurity.Visible = True
    End If
    
End Sub

Private Sub CheckDuplicateTemplates()

    Dim DesktopPath As String
    Dim DownloadsPath As String
    
    On Error Resume Next
    
    DesktopPath = MacScript("return the path to the desktop folder as string") & "Debate.dotm"
    DownloadsPath = MacScript("return the path to the downloads folder as string") & "Debate.dotm"
    
    #If MAC_OFFICE_VERSION >= 15 Then
    If AppleScriptTask("Verbatim.scpt", "FileExists", DesktopPath) = "false" And AppleScriptTask("Verbatim.scpt", "FileExists", DownloadsPath) = "false" Then
    #Else
    If MacScript("tell application ""Finder""" & Chr(13) & "exists file """ & DesktopPath & """" & Chr(13) & "end tell") = "false" And _
    MacScript("tell application ""Finder""" & Chr(13) & "exists file """ & DownloadsPath & """" & Chr(13) & "end tell") = "false" Then
    #End If
        Me.lblDuplicateTemplates.ForeColor = vbBlack
        Me.imgYesDuplicateTemplates.Visible = True
        Me.imgNoDuplicateTemplates.Visible = False
        Me.btnFixDuplicateTemplates.Visible = False
    Else
        Me.lblDuplicateTemplates.ForeColor = vbRed
        Me.imgYesDuplicateTemplates.Visible = False
        Me.imgNoDuplicateTemplates.Visible = True
        Me.btnFixDuplicateTemplates.Visible = True
    End If
    
End Sub

Private Sub btnFixDuplicateTemplates_Click()
    Call Troubleshooting.DeleteDuplicateTemplates
    Me.Hide
    Me.Show
End Sub
