Attribute VB_Name = "frm2016Caselist"
Attribute VB_Base = "0{FA817245-189B-4439-9375-FF0AC0FD34BD}{96C79F6B-2455-4FAA-9E0C-1C3910CD5895}"
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

Private Sub btnCaselistWizard_Click()
    Unload Me
    Caselist.ShowCaselistWizard
End Sub

Private Sub btnConvertToWiki_Click()
    Unload Me
    Caselist.Word2XWikiCites
End Sub

Private Sub btnCiteRequestDoc_Click()
    Unload Me
    Caselist.CiteRequestDoc
End Sub

Private Sub btnCiteRequest_Click()
    Unload Me
    Caselist.CiteRequest
End Sub

Private Sub btnCombineDocs_Click()
    Unload Me
    Caselist.ShowCombineDocs
End Sub

