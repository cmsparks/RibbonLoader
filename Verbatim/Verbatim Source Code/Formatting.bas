Attribute VB_Name = "Formatting"
Option Explicit

Sub UnderlineMode()

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
            If Selection.Paragraphs.outlineLevel = wdOutlineLevelBodyText Then 'Only affect cards
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

Sub ToggleUnderline()
'Toggles any style underlined text to Normal and back to Underline style
    
    'Check for all underlinining, not a specific style, to be more universal
    If Selection.Font.Underline = 1 Then
        Selection.ClearFormatting
    Else
        Selection.Style = "Underline"
    End If

End Sub

Sub PasteText()
' Pastes unformatted text
' Normal Clipboard DataObject is unreliable in Mac VBA - pastes extra characters, so use the built-in method instead

    Selection.PasteSpecial DataType:=wdPasteText

End Sub

Sub Highlight()
    WordBasic.Highlight
End Sub

Sub ShrinkText()
'Cycles non-underlined text in the current paragraph down a size at a time from 11-4pt
'Differences in un-underlined font size will be normalized automatically

    Dim SelectionStart As Long
    Dim SelectionEnd As Long
    Dim FoundFontSize As Long
    
    'Turn off screen updating
    Application.ScreenUpdating = False
    
    'If in "Paragraph" mode, select current paragraph
    If GetSetting("Verbatim", "Format", "ShrinkMode", "Paragraph") = "Paragraph" Then
        'Move selection to start and end of paragraph
        If Selection.Start <> Selection.Paragraphs(1).Range.Start Then Selection.Paragraphs(1).Range.Select
        If Selection.End <> Selection.Paragraphs(1).Range.End Then Selection.Paragraphs(1).Range.Select
    End If
    
    'If not text, exit
    If Selection.Type < 2 Then Exit Sub
    
    'Save selection
    SelectionStart = Selection.Start
    SelectionEnd = Selection.End
    
    'Make sure at least some text is underlined - solves the "shrink rest of document" bug
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Wrap = wdFindStop
        .Format = True
        .Font.Underline = 1
        .Execute
    End With
    
    If Selection.Find.Found = False Then
        Application.StatusBar = "At least some text must be underlined to shrink."
        Exit Sub
    End If
        
    'Reset Selection
    Selection.Start = SelectionStart
    Selection.End = SelectionEnd
        
    'Find first un-underlined part of card and test font size
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Wrap = wdFindStop
        .Format = True
        .Font.Underline = 0
        .Execute
    End With
    
    FoundFontSize = Selection.Font.Size
    
    'Reset selection
    Selection.Start = SelectionStart
    Selection.End = SelectionEnd
    
    'Depending on font size, shrink or reset to normal text size
    Select Case FoundFontSize
        Case Is = wdUndefined   'Multiple font sizes, shrink to 8
            Selection.Find.Replacement.Font.Size = 8
        Case Is > 8
            Selection.Find.Replacement.Font.Size = 8
        Case Is = 8
            Selection.Find.Replacement.Font.Size = 7
        Case Is = 7
            Selection.Find.Replacement.Font.Size = 6
        Case Is = 6
            Selection.Find.Replacement.Font.Size = 5
        Case Is = 5
            Selection.Find.Replacement.Font.Size = 4
        Case Is = 4
            Selection.Find.Replacement.Font.Size = ActiveDocument.Styles("Normal").Font.Size
        Case Else   'Anything weird, go back to normal text size
            Selection.Find.Replacement.Font.Size = ActiveDocument.Styles("Normal").Font.Size
    End Select
    
    'Replace the text and reset Find dialogue
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
        
    'Shrink pilcrows too, just in case they've been underlined
    Call Formatting.ShrinkPilcrows
    
    'Turn on Screen Updating
    Application.ScreenUpdating = True
    
End Sub

Sub ShrinkAll()

    Dim ShrinkMode As String
    Dim p
    
    'Temporarily override ShrinkMode to Paragraph mode
    ShrinkMode = GetSetting("Verbatim", "Format", "ShrinkMode", "Paragraph")
    SaveSetting "Verbatim", "Format", "ShrinkMode", "Paragraph"
    
    'Loop all paragraphs, shrink if body text
    For Each p In ActiveDocument.Paragraphs
        If p.outlineLevel = wdOutlineLevelBodyText Then
            Selection.Start = p.Range.Start
            Call Formatting.ShrinkText
        End If
    Next p
    
    'Restore setting
    SaveSetting "Verbatim", "Format", "ShrinkMode", ShrinkMode

End Sub

Sub Condense()
'Removes white-space from selection and optionally retains paragraph integrity

    Dim CondenseRange As Range
    
    'Turn off Screen Updating
    Application.ScreenUpdating = False
    
    'If selection is too short, exit
    If Len(Selection) < 2 Then Exit Sub
        
    'If end of selection is a line break, shorten it
    If Selection.Characters.Last = vbCr Then Selection.MoveEnd , -1
    
    'Save selection
    Set CondenseRange = Selection.Range
    
    'Condense everything except hard returns
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Wrap = wdFindStop
    
        .Text = "^m"                    'page breaks
        .Replacement.Text = " "
        .Execute Replace:=wdReplaceAll
        
        .Text = "^t"                    'tabs
        .Replacement.Text = " "
        .Execute Replace:=wdReplaceAll
        
        .Text = "^s"                    'non-breaking space
        .Replacement.Text = " "
        .Execute Replace:=wdReplaceAll
        
        .Text = "^b"                    'section break
        .Replacement.Text = " "
        .Execute Replace:=wdReplaceAll
        
        .Text = "^l"                    'new line
        .Replacement.Text = " "
        .Execute Replace:=wdReplaceAll
        
        .Text = "^n"                    'column break
        .Replacement.Text = " "
        .Execute Replace:=wdReplaceAll
    End With
    
    'If paragraph integrity is off, just condense
    If GetSetting("Verbatim", "Format", "ParagraphIntegrity", False) = False Then
        With Selection.Find
            .Text = "^p"
            .Replacement.Text = " "
            .Execute Replace:=wdReplaceAll
        
            .Text = "  "
            .Replacement.Text = " "
            
            While InStr(Selection, "  ")
                .Execute Replace:=wdReplaceAll
            Wend
            
            If Selection.Characters(1) = " " And _
            Selection.Paragraphs(1).Range.Start = Selection.Start Then _
            Selection.Characters(1).Delete
        End With
    
    Else
        'If paragraph integrity and Pilcrows are on, replace paragraph breaks with Pilcrow sign
    
        If GetSetting("Verbatim", "Format", "UsePilcrows", False) = True Then
            With Selection.Find
                .Text = "^p"
                .Replacement.Text = Chr(182) & " " 'Pilcrow sign
                .Replacement.Font.Size = 6
                .Execute Replace:=wdReplaceAll
                
                .Text = Chr(182) & " " & Chr(182)
                .Replacement.Text = Chr(182)
                
                While InStr(Selection, Chr(182) & " " & Chr(182))
                    .Execute Replace:=wdReplaceAll
                Wend
                
                .Text = "  "
                .Replacement.ClearFormatting
                .Replacement.Text = " "
                
                While InStr(Selection, "  ")
                    .Execute Replace:=wdReplaceAll
                Wend
                
                If Selection.Characters(1) = " " And _
                Selection.Paragraphs(1).Range.Start = Selection.Start Then _
                Selection.Characters(1).Delete
                
                'Remove trailing pilcrows
                If Selection.Characters.Last.Previous = Chr(182) Then Selection.Characters.Last.Previous.Delete
            End With
    
        Else 'Else, paragraph integrity is off and Pilcrows are off, leave single paragraph marks
        
            With Selection.Find
                .Text = "^p^w"
                .Execute
                .Replacement.Text = "^p"
                Do Until .Found = False
                    CondenseRange.Select
                    .Execute Replace:=wdReplaceAll
                    CondenseRange.Select
                    .Execute
                Loop
                
                .Text = "^p^p"
                .Execute
                .Replacement.Text = "^p"
                Do Until .Found = False
                    CondenseRange.Select
                    .Execute Replace:=wdReplaceAll
                    CondenseRange.Select
                    .Execute
                Loop
                
                .Text = "  "
                .Replacement.Text = " "
                .Execute Replace:=wdReplaceAll
                
                If Selection.Characters(1) = " " And _
                Selection.Paragraphs(1).Range.Start = Selection.Start Then _
                Selection.Characters(1).Delete
            End With
    
        End If
    End If
    
    'Clear find dialogue
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    
    'Turn on Screen Updating
    Application.ScreenUpdating = True

End Sub

Sub ShrinkPilcrows()
'Shrinks, un-underlines and unbolds all pilcrows in current paragraph to 8pt
'If run with the insertion point at the very beginning of the document, shrinks all pilcrows

    'Turn off screen updating
    Application.ScreenUpdating = False
    
    'If at beginning of document, shrink all pilcrows and exit
    If Selection.Start <= ActiveDocument.Range.Start Then
        Selection.Collapse
        With Selection.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = Chr(182)
            .Replacement.Text = Chr(182)
            .Replacement.Font.Size = 6
            .Replacement.Font.Underline = wdUnderlineNone
            .Replacement.Font.Bold = 0
            .Format = True
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
            
            .ClearFormatting
            .Replacement.ClearFormatting
        End With
        
        Exit Sub
    End If
    
    'If in "Paragraph" mode, select current paragraph
    If GetSetting("Verbatim", "Format", "ShrinkMode", "Paragraph") = "Paragraph" Then
        'Move selection to start and end of paragraph
        If Selection.Start <> Selection.Paragraphs(1).Range.Start Then Selection.Paragraphs(1).Range.Select
        If Selection.End <> Selection.Paragraphs(1).Range.End Then Selection.Paragraphs(1).Range.Select
    End If
    
    'If not text or no selection, exit
    If Selection.Type < 2 Then Exit Sub
    If Selection.Start = Selection.End Then Exit Sub
    
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = Chr(182)
        .Replacement.Text = Chr(182)
        .Replacement.Font.Size = 6
        .Replacement.Font.Underline = wdUnderlineNone
        .Replacement.Font.Bold = 0
        .Format = True
        .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
        
    'Turn on Screen Updating
    Application.ScreenUpdating = True
    
End Sub

Sub ClearToNormal()
'Clears all formatting if text is selected, otherwise sets the paragraph style to Normal

    If Selection.End > Selection.Start Then
        Selection.ClearFormatting
    Else
        Selection.Paragraphs.Style = ActiveDocument.Styles("Normal")
    End If

End Sub

Sub CopyPreviousCite()
'Duplicates previous cite - only works with one-line cites
    
    Dim StartLocation As Long
    
    'Save Current Location
    StartLocation = Selection.Start
    
    'Find previous cite
    With Selection.Find
      .ClearFormatting
      .Text = ""
      .Wrap = wdFindStop
      .Format = True
      .Style = ActiveDocument.Styles("Cite")
      .Forward = False
      .Execute
    End With
    
    'If nothing found, exit
    If Selection.Find.Found = False Then
        Application.StatusBar = "No Cite Found"
        Exit Sub
    End If
    
    'If found, select the whole paragraph
    Selection.Collapse
    Selection.StartOf Unit:=wdParagraph
    Selection.MoveEnd Unit:=wdParagraph, Count:=1
    Selection.Copy
    
    'Return to original location and paste
    Selection.Start = StartLocation
    Selection.Collapse
    Selection.Paste

End Sub

Sub UniHighlight()
'Replaces all highlighting in the document with the selected color

    Selection.Find.ClearFormatting
    Selection.Find.Highlight = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Highlight = True
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

End Sub

Sub RemoveBlanks()
'Removes blank lines from appearing in the Navigation Pane by setting them to Normal text

    Dim p

    'Prompt user to confirm
    If MsgBox("Removing blank lines from the Nav Pane is irreversible. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub

    For Each p In ActiveDocument.Paragraphs
        If p.outlineLevel < wdOutlineLevel5 And Len(p) = 1 Then
            p.Style = "Normal"
        End If
    Next p

End Sub

Sub ShowComments()
'Toggles showing comments

    With ActiveWindow.View
        If .ShowRevisionsAndComments Then
            .ShowRevisionsAndComments = False
        Else
            .ShowRevisionsAndComments = True
        End If
    End With
    
End Sub

Sub InsertHeader()
'Inserts a custom header based on team/user information in Verbatim settings

    ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range.Text = GetSetting("Verbatim", "Main", "SchoolName", "?") & vbCrLf & "File Title" & vbTab & vbTab & GetSetting("Verbatim", "Main", "Name", "?")
    ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).PageNumbers.Add (wdAlignPageNumberRight)

End Sub

Sub UpdateStyles()
'Updates document styles from template

    ActiveDocument.UpdateStyles
    
End Sub

Sub SelectSimilar()
    
    'Turn off error checking
    On Error Resume Next
    
    Application.ScreenUpdating = False
    
    If Selection.Font.Underline = wdUnderlineNone And Selection.Font.Size = ActiveDocument.Styles("Normal").Font.Size Then
        ActiveDocument.Content.Font.Shrink
        WordBasic.SelectSimilarFormatting
        ActiveDocument.Content.Font.Grow
    Else
        WordBasic.SelectSimilarFormatting
    End If
    
    Application.ScreenUpdating = True
    
End Sub

Sub RemoveHyperlinks()
'Remove all hyperlinks from document

    Dim i
    Dim Count As Long
    
    For i = ActiveDocument.Hyperlinks.Count To 1 Step -1
        ActiveDocument.Hyperlinks(i).Delete
        Count = Count + 1
    Next i

    Application.StatusBar = Count & " hyperlinks removed."

End Sub

Sub RemovePilcrows()

    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = Chr(182)
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .Replacement.ClearFormatting
    End With

End Sub

Sub AutoFormatCite()
    
    Dim r As Range
       
    'Set range to current paragraph
    Set r = Selection.Paragraphs(1).Range
    
    'Find first comma
    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Text = ","
        .Wrap = wdFindStop
        .Execute
    End With
    
    'Select word before comma
    r.MoveStart Unit:=wdWord, Count:=-1
    r.MoveEnd Unit:=wdCharacter, Count:=-1
    
    'If it's numeric, it's the year, so expand backwards one word to catch last name, make a cite, and exit
    If IsNumeric(r.Text) = True Then
        r.MoveStart Unit:=wdWord, Count:=-1
        r.Style = "Cite"
        Exit Sub
    Else 'If non-numeric, it's likely the last name, so make a cite
        r.Style = "Cite"
        r.Collapse 0
    End If
    
    'Move right in paragraph until finding a digit - should be the start of the date, then extend to get whole date
    r.MoveStartUntil Cset:="0123456789", Count:=Len(Selection.Paragraphs(1).Range.Text)
    r.MoveEndWhile Cset:="-/0123456789", Count:=Len(Selection.Paragraphs(1).Range.Text)

    'If end of date doesn't match current year, make year portion a cite
    If Right(r.Text, 2) <> Right(Year(Date), 2) Then
        r.Collapse 0
        r.MoveStartWhile Cset:="0123456789", Count:=-4
        r.Style = "Cite"
    
    'Otherwise, skip year potion and extend backwards to make rest of date a cite
    Else
        r.Collapse 0
        r.MoveStartWhile Cset:="0123456789", Count:=-4
        r.Collapse
        r.MoveStartWhile Cset:="-/", Count:=-1
        r.Collapse
        r.MoveStartWhile Cset:="-/0123456789", Count:=-5
        r.Style = "Cite"
    End If

    Set r = Nothing

End Sub

Sub ReformatCiteDates()

    'Go to top of document
    Selection.HomeKey Unit:=wdStory
    Selection.Collapse
    
    'Find each occurrence of the Cite style
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Style = "Cite"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute
        
        'Find all matches
        Do Until .Found = False
            'Select existing cite and clear formatting, then re-format
            Selection.Collapse
            Selection.StartOf Unit:=wdParagraph
            Selection.MoveEnd Unit:=wdParagraph, Count:=1
            Selection.ClearFormatting
            Formatting.AutoFormatCite
            
            'Move down to avoid getting stuck
            Selection.MoveDown Unit:=wdParagraph, Count:=1
            .Execute
        Loop
        
        .ClearFormatting
        .Replacement.ClearFormatting
    End With

End Sub

Sub AutoUnderline()

    Dim Tag As Range
    Dim TagWord As Range
    
    Dim SI As SynonymInfo
    Dim Synonyms() As String
    Dim TagSynonyms As Collection
    Set TagSynonyms = New Collection
    
    Dim i
    Dim j
    Dim k
    
    Dim w As Range
    Dim CardEnd
    
    Dim IntersectionCount
    
    Dim RangeArray() As Range
    Dim RangeScoreArray()
    
    On Error GoTo Handler
    
    'If cursor isn't on a tag, exit
    Selection.Collapse
    If Selection.Paragraphs.outlineLevel <> wdOutlineLevel4 Or Len(Selection.Paragraphs(1).Range.Text) < 2 Then
        MsgBox "Cursor must be in a tag to automatically underline a card."
        Exit Sub
    End If
    
    'Select tag and loop words, add each word and all synonyms if adjective, noun, adverb, or verb
    Selection.Paragraphs(1).Range.Select
    For Each TagWord In Selection.Words
        TagSynonyms.Add TagWord.Text
        Set SI = SynonymInfo(TagWord.Text)
        If SI.MeaningCount > 0 Then
            For i = 1 To SI.MeaningCount
                If SI.PartOfSpeechList(i) < 4 Then '0=Adjective, 1=Noun, 2=Adverb, 3=Verb
                    Synonyms = SI.SynonymList(i)
                    For j = 1 To UBound(Synonyms)
                        TagSynonyms.Add Synonyms(j)
                    Next j
                End If
            Next i
        End If
    Next TagWord
        
    'Select card, then deselect tag - exit if no more content
    Call Paperless.SelectHeadingAndContent
    If Selection.Paragraphs.Count < 2 Then Exit Sub
    Selection.MoveStart Unit:=wdParagraph, Count:=1
    
    'If cite detected in 1st or 2nd paragraph, skip to next to avoid underlining cite
    If Selection.Paragraphs.Count > 2 Then
        If InStr(Selection.Paragraphs(2).Range.Text, "http://") > 0 Or Selection.Paragraphs(2).Range.Characters(1) Like "[(<]" Or Selection.Paragraphs(2).Range.Characters(1) Like "[[]" Then
            Selection.MoveStart Unit:=wdParagraph, Count:=2
        ElseIf InStr(Selection.Paragraphs(1).Range.Text, "http://") > 0 Or Selection.Paragraphs(1).Range.Characters(1) Like "[(<]" Or Selection.Paragraphs(1).Range.Characters(1) Like "[[]" Then
            Selection.MoveStart Unit:=wdParagraph, Count:=1
        End If
    ElseIf InStr(Selection.Paragraphs(1).Range.Text, "http://") > 0 Or Selection.Paragraphs(1).Range.Characters(1) Like "[(<]" Or Selection.Paragraphs(1).Range.Characters(1) Like "[[]" Then
        Selection.MoveStart Unit:=wdParagraph, Count:=1
    End If
    
    'Set end point, then collapse and start processing card
    CardEnd = Selection.End
    Selection.Collapse
    Do While Selection.End <> CardEnd
        
        Selection.MoveEnd wdWord, 1
        Select Case Trim(Selection.Words(Selection.Words.Count).Text)
            'Extend the chunk until we find punctuation, then roll back 1 word
            Case Is = ".", ",", """", "(", ")", "),", ":", ";", " - ", Chr(151) '151 = em dash
                Selection.MoveEnd wdWord, -1
                
                'Loop all words in chunk and compare to all words in TagSynonyms - count it if it matches
                For Each w In Selection.Words
                        For i = 1 To TagSynonyms.Count
                            If LCase(Trim(w.Text)) = LCase(Trim(TagSynonyms(i))) Then IntersectionCount = IntersectionCount + 1
                        Next i
                Next w
                
                'Increase array size
                If Not Not RangeArray Then
                    ReDim Preserve RangeArray(0 To UBound(RangeArray) + 1) As Range
                    ReDim Preserve RangeScoreArray(0 To UBound(RangeScoreArray) + 1)
                Else
                    ReDim RangeArray(0 To 0) As Range
                    ReDim RangeScoreArray(0 To 0)
                End If
                
                'Add the range for the chunk and the normalized chunk score to the arrays, then reset counter
                Set RangeArray(UBound(RangeArray)) = Selection.Range
                RangeScoreArray(UBound(RangeScoreArray)) = IntersectionCount / Selection.Words.Count
                IntersectionCount = 0
                
                'Move one word right to skip punctuation and start new chunk
                Selection.MoveEnd wdWord, 1
                Selection.Collapse Direction:=0
                
        End Select
    Loop
    Selection.Collapse
    
    'Loop through arrays - set style if chunk score is high enough
    '0.1 threshold is pretty good for underline relevance, 0.25 is pretty good for emphasis
    For i = 0 To UBound(RangeArray)
        If RangeScoreArray(i) >= 0.1 Then RangeArray(i).Style = "Underline"
        If GetSetting("Verbatim", "Format", "AutoUnderlineEmphasis", False) = True Then
            If RangeScoreArray(i) >= 0.25 Then RangeArray(i).Style = "Emphasis"
        End If
    Next
    
    'Clean up
    Set TagSynonyms = Nothing
    Set SI = Nothing

    Exit Sub
    
Handler:
    Set TagSynonyms = Nothing
    Set SI = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Sub AutoEmphasizeFirst()

    Dim w As Range
    For Each w In Selection.Words
        w.Characters(1).Style = "Emphasis"
    Next w

End Sub

Sub CondenseNoPilcrows()

    Dim ParagraphIntegrity As Boolean
    
    ParagraphIntegrity = GetSetting("Verbatim", "Format", "ParagraphIntegrity", False)
    SaveSetting "Verbatim", "Format", "ParagraphIntegrity", False
    Call Formatting.Condense
    SaveSetting "Verbatim", "Format", "ParagraphIntegrity", ParagraphIntegrity

End Sub

Sub RemoveEmphasis()

    'Go to top of document
    Selection.HomeKey Unit:=wdStory
    Selection.Collapse

    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Style = "Emphasis"
        .Replacement.Style = "Underline"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .Replacement.ClearFormatting
    End With

End Sub

Sub GetFromCiteMaker()
    
    Dim Script As String
     
    On Error GoTo Handler
    
    Script = "tell application ""Google Chrome""" & vbCr
    Script = Script & "activate" & vbCr
    Script = Script & "delay 1" & vbCr
    Script = Script & "tell application ""System Events"" to keystroke ""c"" using {control down, option down}" & vbCr
    Script = Script & "delay 1" & vbCr
    Script = Script & "end tell" & vbCr
    Script = Script & "tell application ""Microsoft Word""" & vbCr
    Script = Script & "activate" & vbCr
    Script = Script & "end tell"
    
    #If MAC_OFFICE_VERSION >= 15 Then
        AppleScriptTask "Verbatim.scpt", "GetFromCiteMaker", ""
    #Else
        MacScript (Script)
    #End If
    
    Call Formatting.PasteText
    
    Exit Sub
    
Handler:
    MsgBox "Getting from CiteMaker failed - ensure Google Chrome and the CiteMaker extension are installed and open." & vbCrLf & vbCrLf & "Error " & Err.Number & ": " & Err.Description
End Sub

Sub AutoNumberTags()
    Dim p As Paragraph
    Dim i As Long
    
    'Remove pre-existing numbers
    Call Formatting.DeNumberTags
    
    'Loop through each tag and insert a number - restart numbering on any larger heading
    For Each p In ActiveDocument.Paragraphs
        Select Case p.outlineLevel
            Case Is = 1, 2, 3
                i = 0
            Case Is = 4
                If Len(p.Range.Text) > 1 Then
                    i = i + 1
                    p.Range.InsertBefore i & ". "
                End If
            Case Is > 4
                'Do Nothing
        End Select
    Next p

End Sub

Sub DeNumberTags()
    Dim p As Paragraph
    Dim r As Range
    
    'Loop through each tag
    For Each p In ActiveDocument.Paragraphs
        If p.outlineLevel = 4 Then
            
            'Delete numbers from beginning of line if there's a delimiter, then trim
            Set r = p.Range
            r.Collapse
            r.MoveEndWhile Cset:="0123456789.():-", Count:=5
            If Left(r.Text, 5) Like "*[.():-]*" And r.Text <> "" Then
                r.Delete
                r.Collapse
                r.MoveEndWhile Cset:=" "
                If r.Text <> "" Then r.Delete
            End If
        End If
    Next p
    
    Set r = Nothing
End Sub

Sub FixFakeTags()

    Dim p As Paragraph
    
    For Each p In ActiveDocument.Paragraphs
        If p.outlineLevel = wdOutlineLevelBodyText And p.Range.Bold = True And p.Range.Font.Size > ActiveDocument.Styles("Underline").Font.Size Then
            p.Range.Select
            Selection.ClearFormatting
            p.Style = "Tag"
        End If
    Next p
    
End Sub
