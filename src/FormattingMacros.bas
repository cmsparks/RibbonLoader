' FormattingMacros.bas
' || VBA file containing all formatting related
' || macros. Linked to hotkeys and .officeUI
' || XML elements.
Sub Pocket()
    Selection.Style = ActiveDocument.Styles("Pocket")
End Sub
Sub Hat()
    Selection.Style = ActiveDocument.Styles("Hat")
End Sub
Sub Block()
    Selection.Style = ActiveDocument.Styles("Block")
End Sub
Sub Tag()
    Selection.Style = ActiveDocument.Styles("Tag")
End Sub
Sub Cite()
    Selection.Style = ActiveDocument.Styles("Cite")
End Sub
Sub Emphasis()
    Selection.Style = ActiveDocument.Styles("Emphasis")
End Sub
Sub RemoveAllHighlighting()
    'Modified Verbatim RemoveEmphasis Macro
    Selection.HomeKey Unit:=wdStory
    Selection.Collapse

    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Highlight = True
        Do While (.Execute(Forward:=True) = True) = True
            If Selection.Range.HighlightColorIndex <> wdAuto Then
                Selection.Range.HighlightColorIndex = wdAuto
                Selection.Collapse direction:=wdCollapseEnd
            End If
        Loop
        .ClearFormatting
        .Replacement.ClearFormatting
    End With
End Sub
Sub RemoveAllUnderline()
    'Modified Verbatim RemoveEmphasis Macro
    Selection.HomeKey Unit:=wdStory
    Selection.Collapse

    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = "Underline"
        .Replacement.Style = "Normal/Card"
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .Replacement.ClearFormatting
        End With
End Sub
Sub UniHighlight()
    ' Modified Verbatim UniHighlight Macro
    With Selection.Find
        .ClearFormatting
        .Highlight = True
        .Replacement.ClearFormatting
        .Replacement.Highlight = True
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        Execute Replace:=wdReplaceAll
    End With
End Sub
Sub Condense()
'Modifed Verbatim Condense Macro
'Removes white-space from selection
    'Turn off Screen Updating
    Application.ScreenUpdating = False
    
    'If selection is too short, exit
    If Len(Selection) < 2 Then Exit Sub
        
    'If end of selection is a line break, shorten it
    If Selection.Characters.Last = vbCr Then Selection.MoveEnd , -1
    'Condense everything except hard returns
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Wrap = wdFindStop
        .Text = "^m"                    'page breaks
        .Replacement.Text = " "
        
        .Text = "^t"                    'tabs
        .Replacement.Text = " "
        
        .Text = "^s"                    'non-breaking space
        .Replacement.Text = " "
        
        .Text = "^b"                    'section break
        .Replacement.Text = " "
        
        .Text = "^l"                    'new line
        .Replacement.Text = " "
        
        .Text = "^n"                    'column break
        .Replacement.Text = " "
        
        .Text = "^p"
        .Replacement.Text = " "
                   
        .Text = "  "
        .Replacement.Text = " "
    
        .Execute Replace:=wdReplaceAll
            
        If Selection.Characters(1) = " " And _
        Selection.Paragraphs(1).Range.Start = Selection.Start Then _
        Selection.Characters(1).Delete
    End With
    
    'Clear find dialogue
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    
    'Turn on Screen Updating
    Application.ScreenUpdating = True

End Sub
Sub UnderlineMode()
' Verbatim Underline Mode Macro
' TODO: Rewrite and improve efficiency
    Dim PressedControl As CommandBarControl
    
    On Error GoTo Handler
    
    'If Word 2011, get pressed control
    If Application.Version < "15" Then
        Set PressedControl = CommandBars.ActionControl
        If PressedControl Is Nothing Then Exit Sub
    End If
    
    'If mode is off, turn it on:
    If UnderlineModeToggle = False Then
        UnderlineModeToggle = True
        If Application.Version < "15" Then PressedControl.Caption = "Turn Off Underline Mode"
        MsgBox "Underline Mode is turned ON. Click the menu item again to turn off."
        Application.StatusBar = "Underline Mode ON - press button again to cancel."
      
      Do
        DoEvents 'Give control back to application
        If Selection.Type > 1 Then
            If Selection.Paragraphs.OutlineLevel = wdOutlineLevelBodyText Then 'Only affect cards
                If Selection.Font.Underline = wdUnderlineNone Then 'Testing for style here instead doesn't work
                    Selection.Style = "Underline"
                Else
                    Selection.ClearFormatting
                End If
                Selection.Collapse 0 '0 Direction allows keyboard to underline to the right
            End If
        End If
      Loop Until UnderlineModeToggle = False 'Loop until button is pressed again
      
    'Turn it off
    Else
        UnderlineModeToggle = False
        If Application.Version < "15" Then PressedControl.Caption = "Turn On Underline Mode"
        MsgBox "Underline Mode is turned OFF."
        Application.StatusBar = "Underline Mode OFF"
    End If
    
    Set PressedControl = Nothing
    
    Exit Sub
    
Handler:
    UnderlineModeToggle = False
    If Application.Version < "15" Then PressedControl.Caption = "Turn On Underline Mode"
    Application.StatusBar = "Underline Mode OFF"
    Set PressedControl = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub
Public Sub SelectHeadingAndContent()
' Unmodified Verbatim Macro
' TODO: Improve efficiency
    Dim OLevel As Integer
    
    'Move to start of current paragraph and collapse the selection
    Selection.StartOf Unit:=wdParagraph
    Selection.Collapse
        
    'Move backwards through each paragraph to find the first tag, block title, hat, pocket or the top of the document
    Do While True
        If Selection.Paragraphs.OutlineLevel < wdOutlineLevel5 Then Exit Do 'Headings 1-4
        If Selection.Start <= ActiveDocument.Range.Start Then 'Top of document
            Application.StatusBar = "Nothing found to select"
            Exit Sub
        End If
        Selection.Move Unit:=wdParagraph, Count:=-1
    Loop
        
    'Get current outline level
    OLevel = Selection.Paragraphs.OutlineLevel
    
    'Extend selection until hitting the bottom or a bigger outline level
    Selection.MoveEnd Unit:=wdParagraph, Count:=1
    Do While True And Selection.End <> ActiveDocument.Range.End
        Selection.MoveEnd Unit:=wdParagraph, Count:=1
        If Selection.Paragraphs.Last.OutlineLevel <= OLevel Then
            Selection.MoveEnd Unit:=wdParagraph, Count:=-1
            Exit Do 'Bigger Outline Level
        End If
    Loop

End Sub
