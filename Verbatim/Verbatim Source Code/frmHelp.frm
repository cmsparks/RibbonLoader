Attribute VB_Name = "frmHelp"
Attribute VB_Base = "0{A77E080C-267E-4726-AE35-12465610F6AC}{BE1F8590-F53C-4308-B3B3-258620495711}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnManual_Click()
    Unload Me
    Settings.LaunchWebsite ("http://paperlessdebate.com/verbatim")
End Sub

Private Sub btnTutorial_Click()
    Unload Me
    Call Tutorial.LaunchTutorial
End Sub

Private Sub btnTroubleshooter_Click()
    Unload Me
    Call Settings.ShowTroubleshooter
End Sub

Private Sub btnSettings_Click()
    Unload Me
    Call Settings.ShowSettingsForm
End Sub

Private Sub btnOfficeHelp_Click()
    Unload Me
    Call Settings.OpenWordHelp
End Sub

Private Sub btnCheatSheet_Click()
    Unload Me
    Call Settings.ShowCheatSheet
End Sub

Private Sub UserForm_Initialize()

    'Disable incompatible features in 2016
    If Application.Version < "15" Then
        Me.btnTutorial.Enabled = False
    End If
End Sub
