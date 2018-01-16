' SpeechHandler.bas
' || VBA file containing all speech
' || document related macros.
' || Linked to the .officeUI XML
' || elements.
Public SpeechDoc As Document

Sub SetSpeechDoc()
    
End Sub
Sub SendToSpeech()
'Sends content to the Speech doc.  Sends currently selected text,
'or if nothing is selected, the current tag, block, hat, or pocket
'Speech marker doesn't work on Mac Word reading view - code left here in case of new version

    Dim CurrentDoc As String
    Dim d As Document
    Dim FoundDoc As Long

    'Save active document name
    CurrentDoc = ActiveDocument.Name

    'Turn off screen updating for the heavy-lifting
    Application.ScreenUpdating = False
    
    'If text is selected, copy and send it.  Add a return if not in the selection.
    If Selection.End > Selection.Start Then
        Selection.Copy
        
        'Trap for sending to middle of text
        If SpeechDoc.ActiveWindow.Selection.Start <> SpeechDoc.ActiveWindow.Selection.Paragraphs(1).Range.Start Then
            If MsgBox("Sending to the middle of text. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
        End If
        
        SpeechDoc.ActiveWindow.Selection.Paste
        If Selection.Characters.Last.Text <> Chr(13) Then
            SpeechDoc.ActiveWindow.Selection.TypeParagraph
        End If
        Exit Sub
    End If
    
    'If nothing is selected, select the current card, block, hat or pocket
    Call Paperless.SelectHeadingAndContent
        
    'If still nothing selected, exit
    If Selection.Start = Selection.End Then Exit Sub
        
    'Copy the unit
    Selection.Copy
    
    'Trap for sending to middle of text or sending a card into a block/hat
    If SpeechDoc.ActiveWindow.Selection.Start <> SpeechDoc.ActiveWindow.Selection.Paragraphs(1).Range.Start Then
       If MsgBox("Sending to the middle of text. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
    End If
    If Selection.Paragraphs(1).OutlineLevel = 4 Then
        If SpeechDoc.ActiveWindow.Selection.Paragraphs.OutlineLevel < wdOutlineLevel4 Then
            If MsgBox("Sending a card into a block, hat, or pocket.  Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
        End If
    End If
    
    'Paste it
    SpeechDoc.ActiveWindow.Selection.Paste
        
    'Reset Selection
    Selection.Collapse
    
    Set SpeechDoc = Nothing
    
    Application.ScreenUpdating = True
    
    Exit Sub

End Sub
Sub NewSpeech()
    
End Sub
Sub GetDocuments()
    Dim Docs()
    ReDim Docs(0)
    For Each w In Application.Windows
        ReDim Preserve Docs(UBound(Docs) + 1)
        Docs(UBound(Docs) - 1) = w.Document.Name
    Next w
    ReDim Preserve Docs(UBound(Docs) - 1)
    Debug.Print (Docs(0))
End Sub
