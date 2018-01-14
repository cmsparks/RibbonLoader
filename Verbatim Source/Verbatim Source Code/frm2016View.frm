Attribute VB_Name = "frm2016View"
Attribute VB_Base = "0{C21F7DDD-8E55-4754-9E8E-08585195AD08}{1FBC7431-B0DB-4D39-AC22-474E926664E5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub UserForm_Initialize()
    
    On Error GoTo Handler
    
    If Toolbar.InvisibilityToggle = True Then
        Me.btnInvisibilityMode.Caption = "Turn OFF Invisibility Mode"
        Me.btnInvisibilityMode.BackColor = wdRed
    Else
        Me.btnInvisibilityMode.Caption = "Turn ON Invisibility Mode"
        Me.btnInvisibilityMode.BackColor = wdGreen
    End If
    
    Exit Sub
    
Handler:
        MsgBox "Error " & Err.Number & ": " & Err.Description
    
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnDefaultView_Click()
    Unload Me
    View.DefaultView
End Sub

Private Sub btnInvisibilityMode_Click()

    Unload Me
    
    If Toolbar.InvisibilityToggle = False Then
        Toolbar.InvisibilityToggle = True
        Call View.InvisibilityOn
    Else
        Toolbar.InvisibilityToggle = False
        Call View.InvisibilityOff
    End If

End Sub
