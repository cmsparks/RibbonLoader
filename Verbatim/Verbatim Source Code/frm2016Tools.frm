Attribute VB_Name = "frm2016Tools"
Attribute VB_Base = "0{FB611679-75A8-413E-A8FB-3598F2A46CD4}{AF54F46C-529A-453D-B80B-2BB24716C36A}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub UserForm_Initialize()
    
    On Error GoTo Handler
    
    If Toolbar.RecordAudioToggle = True Then
        Me.btnRecordAudio.Caption = "STOP Audio Recording"
        Me.btnRecordAudio.BackColor = wdRed
    Else
        Me.btnRecordAudio.Caption = "START Audio Recording"
        Me.btnRecordAudio.BackColor = wdGreen
    End If
    
    Exit Sub
    
Handler:
        MsgBox "Error " & Err.Number & ": " & Err.Description
    
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnStartTimer_Click()
    Unload Me
    Paperless.StartTimer
End Sub

Private Sub btnDocumentStats_Click()
    Unload Me
        Stats.ShowStatsForm
End Sub

Private Sub btnConvertBackfile_Click()
    Unload Me
    Convert.ShowConvertForm
End Sub

Private Sub btnRecordAudio_Click()

    Unload Me
    
    If Toolbar.RecordAudioToggle = False Then
        Toolbar.RecordAudioToggle = True
        Call Audio.StartRecord
    Else
        Toolbar.RecordAudioToggle = False
        Call Audio.SaveRecord
    End If

End Sub
Private Sub btnAutoOpenFolder_Click()
    Unload Me
    Paperless.AutoOpenFolder
End Sub

Private Sub btnAddWarrant_Click()
    Unload Me
    Paperless.NewWarrant
End Sub

Private Sub btnDeleteAllWarrants_Click()
    Unload Me
    Paperless.DeleteAllWarrants
End Sub
