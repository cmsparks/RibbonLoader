Attribute VB_Name = "frm2016Share"
Attribute VB_Base = "0{04343DF6-FE2A-4AB5-A1B5-92A78EC45C24}{0330DF67-AC69-4DC1-ADCA-B7C3A6B16713}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub UserForm_Initialize()
    
    On Error GoTo Handler
    
    Exit Sub
    
Handler:
        MsgBox "Error " & Err.Number & ": " & Err.Description
    
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnCopyToUSB_Click()
    Unload Me
    Paperless.CopyToUSB
End Sub

Private Sub btnEmail_Click()
    Unload Me
    Email.ShowEmailForm
End Sub
