Attribute VB_Name = "frmChooseSpeechDoc"
Attribute VB_Base = "0{546D38E6-F04B-49C3-B50C-5884F4C5F8E0}{8A31D0A2-4F54-4242-AD1B-5ED19B76C1E2}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim w As Window
    Dim i As Integer
    
    On Error GoTo Handler
    
    'Loop through open Windows - use Windows because Application.Documents collection gets corrupted
    For Each w In Application.Windows
        If InStr(LCase(w.Document.Name), "speech") Then
            Me.lboxDocuments.AddItem w.Document.Name, 0
        Else
            Me.lboxDocuments.AddItem w.Document.Name
        End If
    Next w
    
    'Select the active speech doc
    For i = 0 To Me.lboxDocuments.ListCount - 1
        If Me.lboxDocuments.List(i) = ActiveSpeechDoc Then
            Me.lboxDocuments.Selected(i) = True
        'ElseIf InStr(LCase(Me.lboxDocuments.List(i)), "speech") Then
        '    Me.lboxDocuments.Selected(i) = True
        End If
    Next i

    Exit Sub
    
Handler:
    'Periodic inexplicable runtime error
    If Err.Number = 5097 Then
        MsgBox "Your Word interface has been corrupted. Try restarting Word to fix it."
    Else
        MsgBox "Error " & Err.Number & ": " & Err.Description
    End If
    
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnChooseSpeechDoc_Click()

    Dim i As Integer
    Dim DocSelected As Boolean

    'Make sure a document is selected
    For i = 0 To Me.lboxDocuments.ListCount - 1
        If Me.lboxDocuments.Selected(i) = True Then DocSelected = True
    Next i

    If DocSelected = False Then
        MsgBox "You must select a document first."
    Else
        ActiveSpeechDoc = Me.lboxDocuments.Value
        Unload Me
    End If
End Sub
