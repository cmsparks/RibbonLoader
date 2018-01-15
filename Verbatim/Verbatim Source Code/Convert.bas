Attribute VB_Name = "Convert"
Option Explicit

Sub ShowConvertForm()
    Dim ConvertForm As frmConvert
    Set ConvertForm = New frmConvert
    ConvertForm.Show
End Sub

Sub ConvertBackfile(File As String, ConvertFrom As Integer)

    Dim NewFileName As String
        
    'Turn off error checking - too many ways for it to go wrong
    On Error Resume Next
    
    'Check template exists
    #If MAC_OFFICE_VERSION >= 15 Then
        If AppleScriptTask("Verbatim.scpt", "FileExists", "Macintosh HD" & Replace(Replace(Application.NormalTemplate.Path & "/Debate.dotm", ".localized", ""), "/", ":")) = "false" Then
            Application.StatusBar = "Debate.dotm not found in your Templates folder - it must be installed to convert files."
            Exit Sub
        End If
    #Else
        If MacScript("tell application ""Finder""" & Chr(13) & "exists file """ & Application.NormalTemplate.Path & ":My Templates:Debate.dotm" & """" & Chr(13) & "end tell") = "false" Then
            Application.StatusBar = "Debate.dotm not found in your My Templates folder - it must be installed to convert files."
            Exit Sub
        End If
    #End If
    
    'If converting a file other than the active document, open it before copying
    If File <> ThisDocument.FullName Then
        'Open the file in the background and activate it
        Documents.Open FileName:=File
        Documents(File).Activate
    
        'Copy everything except header/footer
        ThisDocument.Content.Select
        Selection.Copy

        'Close original file
        Documents(File).Close SaveChanges:=wdDoNotSaveChanges
    Else
        'Copy everything except header/footer
        ThisDocument.Content.Select
        Selection.Copy
    End If
    
    'Add new document based on debate template
    #If MAC_OFFICE_VERSION >= 15 Then
        Application.Documents.Add Template:=Application.NormalTemplate.Path & "/Debate.dotm"
    #Else
        Application.Documents.Add Template:=Application.NormalTemplate.Path & ":My Templates:Debate.dotm"
    #End If

    'Paste into new document. If converting a non-Verbatim file, match destination formatting
    If ConvertFrom = 3 Then 'non-Verbatim
        Selection.PasteAndFormat (wdFormatSurroundingFormattingWithEmphasis)
    Else
        Selection.Paste
    End If
    
    'Go to top of document and collapse selection
    Selection.HomeKey Unit:=wdStory
    Selection.Collapse
        
    'Call conversion macro based on userform
    Select Case ConvertFrom
        Case Is = 1 'Verbatim 3
            ConvertFromV3
        Case Is = 2 'Verbatim 2
            ConvertFromV2
        Case Is = 3 'Non-Verbatim
            ConvertNonVerbatim
        Case Is = 4 'Synergy
            ConvertSynergy
    End Select
    
    'Create new file name
    NewFileName = Left(File, InStrRev(File, ".") - 1) & " - Converted.docx"
    
    'Check if NewFileName already exists
    #If MAC_OFFICE_VERSION >= 15 Then
    If AppleScriptTask("Verbatim.scpt", "FileExists", NewFileName) = "true" Then
    #Else
    If MacScript("tell application ""Finder""" & Chr(13) & "exists file """ & NewFileName & """" & Chr(13) & "end tell") = "true" Then
    #End If
        With Application.Dialogs(wdDialogFileSaveAs)
            .Name = NewFileName
            .Show
        End With
    Else
        ActiveDocument.SaveAs FileName:=NewFileName, FileFormat:=wdFormatXMLDocument
    End If
    
End Sub

Private Sub ConvertFromV3()
'Converts files produced in Verbatim version 3.x to Verbatim 4 format
'Only requires decreasing each heading level by one to make room for the "Pocket" style

    Dim p

    For Each p In ActiveDocument.Paragraphs
        If p.outlineLevel = wdOutlineLevel3 Then p.Style = "Tag"
        If p.outlineLevel = wdOutlineLevel2 Then p.Style = "Block"
        If p.outlineLevel = wdOutlineLevel1 Then p.Style = "Hat"
    Next p

End Sub

Private Sub ConvertFromV2()
' Converts files produced in Verbatim version 2.x to Verbatim 4 format
    
    Dim p
    Dim s

    'Turn off error checking -- too many ways for it to go wrong
    On Error Resume Next
    
    'Replace styles and blank lines
    'Cite commented out because it's a character style for most people
    'Cites will be replaced manually later
    For Each p In ActiveDocument.Paragraphs
        If InStr(p.Style, "hat") <> 0 Then
            p.Style = "Heading 2,Hat"
        ElseIf InStr(p.Style, "Block Title") <> 0 Then
            p.Style = "Heading 3,Block"
        ElseIf InStr(p.Style, "tag") <> 0 Then
            p.Style = "Heading 4,Tag"
        'ElseIf InStr(p.Style, "cite") <> 0 Then
         '   p.Style = "Style Style Bold + 12 pt,Cite"
        ElseIf InStr(p.Style, "Normal") <> 0 Then
            p.Style = "Normal,Normal/Card"
        ElseIf InStr(p.Style, "card") <> 0 Then
            p.Style = "Normal,Normal/Card"
        ElseIf InStr(p.Style, "underline") <> 0 Then
            p.Style = "Style Bold Underline,Underline"
        ElseIf InStr(p.Style, "Emphasis2") <> 0 Then
            p.Style = "Emphasis"
        Else
            p.Style = "Normal,Normal/Card"
        End If
                    
        'Remove blank lines
        If p.outlineLevel < wdOutlineLevel5 And Len(p) = 1 Then
            p.Style = "Normal,Normal/Card"
        End If
    Next p
    
    'Replace underlined text
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Wrap = wdFindStop
        .Replacement.Text = ""
        .Format = True
        .Style = "underline"
        .Replacement.Style = "Style Bold Underline,Underline"
        .Execute Replace:=wdReplaceAll
    End With
    
    'Replace cites
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Wrap = wdFindStop
        .Replacement.Text = ""
        .Format = True
        .Style = "cite"
        .Replacement.Style = "Style Style Bold + 12 pt,Cite"
        .Execute Replace:=wdReplaceAll
    End With
    
    'Replace improperly formatted cites as "tags"
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Wrap = wdFindStop
        .Replacement.Text = ""
        .Format = True
        .Style = "tag"
        .Replacement.Style = "Style Style Bold + 12 pt,Cite"
        .Execute Replace:=wdReplaceAll
    End With
    
    'Replace Emphasis
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Wrap = wdFindStop
        .Replacement.Text = ""
        .Format = True
        .Style = "Emphasis2"
        .Replacement.Style = "Emphasis"
        .Execute Replace:=wdReplaceAll
    End With
    
    'Delete Old styles - deactivate to try and save underlining
    With ActiveDocument.Styles
        .Item("hat").Delete
        .Item("Block Title").Delete
        .Item("Block Title2").Delete
        .Item("Block Heading").Delete
        .Item("Block Headings").Delete
        .Item("Block Name").Delete
        .Item("tag").Delete
        .Item("Tags").Delete
        .Item("cite").Delete
        .Item("Cites").Delete
        .Item("Author-Date").Delete
        .Item("Emphasis2").Delete
        .Item("underline").Delete
        .Item("card").Delete
        .Item("Cards").Delete
    End With
        
    'Delete Char styles
    For Each s In ActiveDocument.Styles
        If InStr(s, "char") <> 0 Then s.Delete
        If InStr(s, "Char") <> 0 Then s.Delete
    Next s
                    
    'Go to top of document and collapse selection
    Selection.HomeKey Unit:=wdStory
    Selection.Collapse
        
    'Remove Page Breaks
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^m"
        .Wrap = wdFindStop
        .Replacement.Text = ""
        .Execute Replace:=wdReplaceAll
    End With
        
End Sub

Private Sub ConvertNonVerbatim()
'Attempts to convert files from non-Verbatim templates to Verbatim 4 Format
'Works by pasting the file in to "Match Destination Formatting" and then trying to guess headings

    'Turn off error checking -- too many ways for it to go wrong
    On Error Resume Next
        
    'Delete TOC if exists - otherwise slows macro too much
    If ActiveDocument.TablesOfContents(1) Then ActiveDocument.TablesOfContents(1).Delete
        
    'Converts in this order to ensure headings changes don't ruin underlining
    ConvertNonVerbatim_Underline
    ConvertNonVerbatim_Headings
    ConvertNonVerbatim_Cites
    ConvertNonVerbatim_PageBreaks
    ConvertNonVerbatim_RemoveBlanks
    ConvertNonVerbatim_HatFixer

End Sub

Private Sub ConvertNonVerbatim_Underline()

    'Go to top of document and collapse selection
    Selection.HomeKey Unit:=wdStory
    Selection.Collapse
    
    'Change underlining from font to Style
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Wrap = wdFindStop
        .Replacement.Text = ""
        .Format = True
        .Font.Underline = wdUnderlineSingle
        .Replacement.Style = "Underline"
        .Execute Replace:=wdReplaceAll
    End With

End Sub
Private Sub ConvertNonVerbatim_Headings()
'Tries to identify headings by leading and trailing blank lines
'Doesn't work perfectly, but usually catches the bulk of headings

    Dim p
    Dim BlankAfter
    Dim IsBold
    Dim ParaLength
    
    'Turn off error checking -- too many ways for it to go wrong
    On Error Resume Next
        
    'Go to top of document and collapse selection
    Selection.HomeKey Unit:=wdStory
    Selection.Collapse
    
    For Each p In ActiveDocument.Paragraphs
        
        'Clear Variables
        BlankAfter = False
        IsBold = False
        
        'Set first line as hat
        If p.Range.Start = ActiveDocument.Range.Start Then
            p.Style = "Block"
            GoTo NextP
        End If
        
        'Check if bold text by checking last letter of line
        If p.Range.End - p.Range.Start > 3 Then
            Selection.Start = p.Range.End - 3
        Else
            Selection.Start = p.Range.Start
        End If
        Selection.End = p.Range.End
        With Selection.Find
            .ClearFormatting
            .Text = ""
            .Wrap = wdFindStop
            .Replacement.Text = ""
            .Format = True
            .Font.Bold = True
            .Execute
        End With
        If Selection.Find.Found = True Then
            IsBold = True
        Else
            GoTo NextP
        End If
            
        'Check if cite by searching for non-bold text too
        Selection.Start = p.Range.Start
        Selection.End = p.Range.End
        With Selection.Find
            .ClearFormatting
            .Text = ""
            .Wrap = wdFindStop
            .Replacement.Text = ""
            .Format = True
            .Font.Bold = False
            .Execute
        End With
        If Selection.Find.Found = True Then
            IsBold = False
        End If
                    
        'Check if blank line after
        Selection.Start = p.Range.Start
        Selection.Collapse
        Selection.Move Unit:=wdParagraph, Count:=1
        ParaLength = Selection.Paragraphs(1).Range.End - Selection.Paragraphs(1).Range.Start
        If ParaLength < 2 Then BlankAfter = True
        If ParaLength > 2 And Len(Trim(Selection.Paragraphs(1))) < 2 Then BlankAfter = True
        
        'Assign Styles
        If IsBold = True Then
            If BlankAfter = True Then
                p.Style = "Block"
            Else
                p.Style = "Tag"
            End If
        End If
           
NextP:
    Next p

End Sub

Private Sub ConvertNonVerbatim_Cites()

    'Go to top of document and collapse selection
    Selection.HomeKey Unit:=wdStory
    Selection.Collapse
    
    'Replace Cites
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Size = 11
        .Bold = True
        .Underline = wdUnderlineNone
    End With
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Style Style Bold + 12 pt,Cite")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

End Sub

Private Sub ConvertNonVerbatim_PageBreaks()

    'Go to top of document and collapse selection
    Selection.HomeKey Unit:=wdStory
    Selection.Collapse
    
    'Remove Page Breaks
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^m"
        .Wrap = wdFindStop
        .Replacement.Text = ""
        .Execute Replace:=wdReplaceAll
    End With
    
End Sub

Private Sub ConvertNonVerbatim_RemoveBlanks()
'Removes blank lines from NavPane
    Dim p
    
    For Each p In ActiveDocument.Paragraphs
        If p.outlineLevel < wdOutlineLevel5 And Len(p) = 1 Then
            p.Style = "Normal"
        End If
    Next p

End Sub

Private Sub ConvertNonVerbatim_HatFixer()
'Identifies Hats with leading or trailing ***'s

    Dim p
    
    For Each p In ActiveDocument.Paragraphs
     If InStr(p, "***") = 1 Then p.Style = "Hat"
    Next p

End Sub

Sub ConvertSynergy()
' Converts files produced in Synergy and derivatives (such as the Georgetown Template) to Verbatim 4 format
' Does not work perfectly, because Synergy doesn't have standard heading/style names across different implementations
    
    Dim p
    Dim s

    'Turn off error checking -- too many ways for it to go wrong
    On Error Resume Next
    
    'Delete TOC if exists - otherwise slows macro too much
    If ActiveDocument.TablesOfContents(1) Then ActiveDocument.TablesOfContents(1).Delete
    
    'Convert hats - Synergy "Hat" styles already work, so this just catches Hat's with ***'s
    ConvertNonVerbatim_HatFixer
    
    'Convert headings to block titles - Synergy only uses Heading 1 by default
    For Each p In ActiveDocument.Paragraphs
        If p.outlineLevel = wdOutlineLevel1 Then p.Style = "Heading 3,Block"
    Next p
    
    'Try manually replacing styles and blank lines
    For Each p In ActiveDocument.Paragraphs
        If InStr(p.Style, "hat") <> 0 Then
            p.Style = "Heading 2,Hat"
        ElseIf InStr(p.Style, "HAT") <> 0 Then
            p.Style = "Heading 2,Hat"
        ElseIf InStr(p.Style, "Block Title") <> 0 Then
            p.Style = "Heading 3,Block"
        ElseIf InStr(p.Style, "Block Heading") <> 0 Then
            p.Style = "Heading 3,Block"
        ElseIf InStr(p.Style, "Block Name") <> 0 Then
            p.Style = "Heading 3,Block"
        ElseIf InStr(p.Style, "tag") <> 0 Then
            p.Style = "Heading 4,Tag"
        ElseIf InStr(p.Style, "Tag") <> 0 Then
            p.Style = "Heading 4,Tag"
        ElseIf InStr(p.Style, "Normal") <> 0 Then
            p.Style = "Normal,Normal/Card"
        ElseIf InStr(p.Style, "card") <> 0 Then
            p.Style = "Normal,Normal/Card"
         ElseIf InStr(p.Style, "Cards") <> 0 Then
            p.Style = "Normal,Normal/Card"
        ElseIf InStr(p.Style, "underline") <> 0 Then
            p.Style = "Style Bold Underline,Underline"
        ElseIf InStr(p.Style, "Debate Underline") <> 0 Then
            p.Style = "Style Bold Underline,Underline"
        ElseIf InStr(p.Style, "Underline + Bold") <> 0 Then
            p.Style = "Style Bold Underline,Underline"
         ElseIf InStr(p.Style, "text bold") <> 0 Then
            p.Style = "Style Bold Underline,Underline"
        ElseIf InStr(p.Style, "Emphasis2") <> 0 Then
            p.Style = "Emphasis"
        ElseIf p.outlineLevel = wdOutlineLevelBodyText Then
            p.Style = "Normal,Normal/Card"
        End If
    
    Next p
  
    'Replace cites - done separately from above because it's a character style
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Wrap = wdFindStop
        .Replacement.Text = ""
        .Format = True
        .Style = "cite"
        .Replacement.Style = "Style Style Bold + 12 pt,Cite"
        .Execute Replace:=wdReplaceAll
    End With
    
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Wrap = wdFindStop
        .Replacement.Text = ""
        .Format = True
        .Style = "Cites"
        .Replacement.Style = "Style Style Bold + 12 pt,Cite"
        .Execute Replace:=wdReplaceAll
    End With
    
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Wrap = wdFindStop
        .Replacement.Text = ""
        .Format = True
        .Style = "Author-Date"
        .Replacement.Style = "Style Style Bold + 12 pt,Cite"
        .Execute Replace:=wdReplaceAll
    End With
    
    'Replace improperly formatted cites as "tags"
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Wrap = wdFindStop
        .Replacement.Text = ""
        .Format = True
        .Style = "tag"
        .Replacement.Style = "Style Style Bold + 12 pt,Cite"
        .Execute Replace:=wdReplaceAll
    End With
    
    'Replace Emphasis
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Wrap = wdFindStop
        .Replacement.Text = ""
        .Format = True
        .Style = "Emphasis2"
        .Replacement.Style = "Emphasis"
        .Execute Replace:=wdReplaceAll
    End With
          
    'Delete Old styles - deactivate to try and save underlining
    With ActiveDocument.Styles
        .Item("hat").Delete
        .Item("Block Title").Delete
        .Item("Block Title2").Delete
        .Item("Block Heading").Delete
        .Item("Block Headings").Delete
        .Item("Block Name").Delete
        .Item("tag").Delete
        .Item("Tags").Delete
        .Item("cite").Delete
        .Item("Cites").Delete
        .Item("Author-Date").Delete
        .Item("Emphasis2").Delete
        .Item("underline").Delete
        .Item("card").Delete
        .Item("Cards").Delete
    End With
        
    'Delete Char styles
    For Each s In ActiveDocument.Styles
        If InStr(s, "char") <> 0 Then s.Delete
        If InStr(s, "Char") <> 0 Then s.Delete
    Next s
          
    'Go to top of document and collapse selection
    Selection.HomeKey Unit:=wdStory
    Selection.Collapse
        
    'Remove Page Breaks
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^m"
        .Wrap = wdFindStop
        .Replacement.Text = ""
        .Execute Replace:=wdReplaceAll
    End With
    
End Sub
