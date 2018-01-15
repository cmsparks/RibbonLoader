Attribute VB_Name = "Caselist"
Option Explicit

Sub ShowCaselistWizard()
    Dim CaselistWizard As frmCaselist
    Set CaselistWizard = New frmCaselist
    CaselistWizard.Show
End Sub

Sub ShowCombineDocs()
    Dim CombineDocs As frmCombineDocs
    Set CombineDocs = New frmCombineDocs
    CombineDocs.Show
End Sub

'*************************************************************************************
'* CITEIFY FUNCTIONS                                                                                            *
'*************************************************************************************

Sub CiteRequest()

    Selection.Collapse
    
    'Make sure cursor is in a card
    If Selection.Paragraphs.outlineLevel <> wdOutlineLevelBodyText Then
        MsgBox "Cursor must be in card text - it appears to be in a heading."
        Exit Sub
    End If
    
    'If card is longer than 50 words, remove all but the first and last few
    With Selection
        .StartOf Unit:=wdParagraph
        .MoveEnd Unit:=wdParagraph, Count:=1
        If .Range.ComputeStatistics(wdStatisticWords) > 50 Then
            .Range.HighlightColorIndex = wdNoHighlight 'Remove highlighting
            .MoveStart Unit:=wdWord, Count:=15
            .MoveEnd Unit:=wdWord, Count:=-15
            .TypeText vbCrLf & "AND" & vbCrLf
        Else
            MsgBox "Cut longer cards!"
        End If
    
    End With

End Sub

Public Sub CiteRequestAll()

    Dim p
    Dim r As Range
    
    Dim pCount As Long
    Dim i
    
    'Delete blank paragraphs to make processing easier
    For Each p In ActiveDocument.Paragraphs
        If Len(p) = 1 Then p.Range.Delete
    Next p
    
    'Go to top of document
    Selection.HomeKey Unit:=wdStory
    Selection.Collapse

    'Find tags
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .ParagraphFormat.outlineLevel = wdOutlineLevel4
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Forward = True
        .Wrap = wdFindContinue
        
        'Loop all tags
        Do While .Execute And Selection.End <> ActiveDocument.Range.End
            
            'Select card
            Call Paperless.SelectHeadingAndContent
            
            'If less than 3 paragraphs (tag, cite, card), something's weird so don't do anything
            If Selection.Paragraphs.Count < 3 Then
                'Do Nothing
            
            'If 3 paragraphs, cite request 3rd paragraph, which will almost always be the card text
            ElseIf Selection.Paragraphs.Count = 3 Then
                Set r = Selection.Paragraphs(3).Range
                
            'If 4 or more paragraphs, non-obvious cite
            Else
                
                'If 2nd, 3rd or 4th paragraph has a URL, start range with next paragraph
                If InStr(Selection.Paragraphs(2).Range.Text, "http://") > 0 Then
                    Set r = ActiveDocument.Range(Start:=Selection.Paragraphs(3).Range.Start, End:=Selection.Range.End)
                ElseIf InStr(Selection.Paragraphs(3).Range.Text, "http://") > 0 Then
                    Set r = ActiveDocument.Range(Start:=Selection.Paragraphs(4).Range.Start, End:=Selection.Range.End)
                ElseIf InStr(Selection.Paragraphs(4).Range.Text, "http://") > 0 Then
                    Set r = ActiveDocument.Range(Start:=Selection.Paragraphs(5).Range.Start, End:=Selection.Range.End)
                
                'No URL found, try brackets
                Else
                    
                    'If starting character of 2nd, 3rd or 4th paragraph is one of (<[, it's likely a cite
                    If Selection.Paragraphs(2).Range.Characters(1) Like "[(<]" Or Selection.Paragraphs(2).Range.Characters(1) Like "[[]" Then
                        Set r = ActiveDocument.Range(Start:=Selection.Paragraphs(3).Range.Start, End:=Selection.Range.End)
                    ElseIf Selection.Paragraphs(3).Range.Characters(1) Like "[(<]" Or Selection.Paragraphs(3).Range.Characters(1) Like "[[]" Then
                        Set r = ActiveDocument.Range(Start:=Selection.Paragraphs(4).Range.Start, End:=Selection.Range.End)
                    ElseIf Selection.Paragraphs(4).Range.Characters(1) Like "[(<]" Or Selection.Paragraphs(4).Range.Characters(1) Like "[[]" Then
                        Set r = ActiveDocument.Range(Start:=Selection.Paragraphs(5).Range.Start, End:=Selection.Range.End)
                    
                    'No Bracket found, try line-length
                    Else
                        'If 2nd paragraph is a short line, it's likely to be a 2-line cite, so cite request paragraphs 4+
                        If Selection.Paragraphs(2).Range.Characters.Count < 100 Then
                            Set r = ActiveDocument.Range(Start:=Selection.Paragraphs(4).Range.Start, End:=Selection.Range.End)
                        'Else it's likely a single line cite, so cite request paragraphs 3+
                        Else
                            Set r = ActiveDocument.Range(Start:=Selection.Paragraphs(3).Range.Start, End:=Selection.Range.End)
                        End If
                    End If
                End If
            End If
            
            'Cite request the range
            If Not r Is Nothing Then
                If r.Words.Count > 50 Then
                    r.MoveStart Unit:=wdWord, Count:=15
                    r.MoveEnd Unit:=wdWord, Count:=-15
                    r.Text = vbCrLf & "AND" & vbCrLf
                End If
            End If
            
            'Reset range for next loop
            Set r = Nothing
                
            'Collapse right so find moves on
            Selection.Collapse wdCollapseEnd
                        
        Loop
    End With
       
    'Add a newline before each heading to keep plaintext output clean
    'Have to go backwards on Mac to avoid infinite loop from adding element to Paragraphs collection
    pCount = ActiveDocument.Paragraphs.Count
    For i = pCount To 1 Step -1
        If ActiveDocument.Paragraphs(i).outlineLevel < 5 Then
            ActiveDocument.Paragraphs(i).Range.InsertBefore vbCrLf
            ActiveDocument.Paragraphs(i).OutlineDemoteToBody
        End If
    Next
    
End Sub

Sub CiteRequestDoc()
    
    On Error GoTo Handler

    'Check template exists
    #If MAC_OFFICE_VERSION >= 15 Then
        If AppleScriptTask("Verbatim.scpt", "FileExists", "Macintosh HD" & Replace(Replace(Application.NormalTemplate.Path & "/Debate.dotm", ".localized", ""), "/", ":")) = "false" Then
            Application.StatusBar = "Debate.dotm not found in your Templates folder - it must be installed to create a cite request doc."
            Exit Sub
        End If
    #Else
        If MacScript("tell application ""Finder""" & Chr(13) & "exists file """ & Application.NormalTemplate.Path & ":My Templates:Debate.dotm" & """" & Chr(13) & "end tell") = "false" Then
            Application.StatusBar = "Debate.dotm not found in your My Templates folder - it must be installed to create a cite request doc."
            Exit Sub
        End If
    #End If
    
    'Copy everything except header/footer
    ActiveDocument.Content.Select
    Selection.Copy

    'Add new document based on debate template
    #If MAC_OFFICE_VERSION >= 15 Then
        Application.Documents.Add Template:=Application.NormalTemplate.Path & "/Debate.dotm"
    #Else
        Application.Documents.Add Template:=Application.NormalTemplate.Path & ":My Templates:Debate.dotm"
    #End If
    
    'Paste into new document
    Selection.Paste
    
    'Go to top of document and collapse selection
    Selection.HomeKey Unit:=wdStory
    Selection.Collapse

    'Convert all cites
    Call Caselist.CiteRequestAll
    
    'Remove highlighting
    ActiveDocument.Content.Select
    Selection.Range.HighlightColorIndex = wdNoHighlight 'Remove highlighting
    Selection.Collapse
    
    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

'*************************************************************************************
'* WIKIFY FUNCTIONS                                                                                             *
'*************************************************************************************

Sub Word2XWikiCites()

    'Cite request and wikify doc
    Call Caselist.CiteRequestDoc
    Call Caselist.Word2XWikiMain
    
    'Clear all formatting
    ActiveDocument.Content.Select
    Selection.ClearFormatting

    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Public Sub Word2XWikiMain()
'Based on Word2MediaWiki, modified for XWiki Syntax:
'http://www.mediawiki.org/wiki/Word_macros
'Bold/Italic/Underline text is just set to normal to keep the output clean

    Application.ScreenUpdating = False
    
    On Error Resume Next
       
    Call XWikiReplaceQuotes
    Call XWikiReplaceDashes
    Call Formatting.RemovePilcrows
    Call XWikiEscapeChars
    Call XWikiReplaceGroups
    Call XWikiConvertHyperlinks
    Call XWikiConvertH1
    Call XWikiConvertH2
    Call XWikiConvertH3
    Call XWikiConvertH4
    Call XWikiConvertH5
    Call XWikiConvertCites
    Call XWikiConvertItalic
    Call XWikiConvertBold
    Call XWikiConvertUnderline
    Call XWikiConvertSuperscript
    Call XWikiConvertSubscript
    Call XWikiRemoveHighlighting
    Call XWikiRemoveComments
    
    ' Copy to clipboard
    ActiveDocument.Content.Copy
    Application.ScreenUpdating = True
    
End Sub

Private Sub XWikiReplaceQuotes()
'Replace all smart quotes with their dumb equivalents
    Dim Quotes As Boolean
    Quotes = Options.AutoFormatAsYouTypeReplaceQuotes
    Options.AutoFormatAsYouTypeReplaceQuotes = False
    ReplaceString ChrW(8220), """"
    ReplaceString ChrW(8221), """"
    ReplaceString "_", "'"
    ReplaceString "", "'"
    Options.AutoFormatAsYouTypeReplaceQuotes = Quotes
End Sub

Private Sub XWikiReplaceDashes()
    ReplaceString "--", ChrW(8212)
End Sub

Private Sub XWikiReplaceGroups()
    ReplaceString "(((", "~(~(~("
    ReplaceString ")))", "~)~)~)"
End Sub

Private Sub XWikiEscapeChars()
    'EscapeCharacter "*"
    EscapeCharacter "#"
    'EscapeCharacter "_"
    'EscapeCharacter "-"
    'EscapeCharacter "+"
    EscapeCharacter "{"
    EscapeCharacter "}"
    EscapeCharacter "["
    EscapeCharacter "]"
    EscapeCharacter "~"
    EscapeCharacter "^^"
    EscapeCharacter "|"
    'EscapeCharacter "'"
End Sub

Private Function EscapeCharacter(Char As String)
    ReplaceString Char, "~" & Char
End Function

Private Function ReplaceString(findStr As String, replacementStr As String)
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = findStr
        .Replacement.Text = replacementStr
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Function

Private Sub XWikiConvertHyperlinks()
    Call Formatting.RemoveHyperlinks
End Sub

Private Sub XWikiConvertH1()
    ReplaceHeading wdOutlineLevel1, "="
End Sub

Private Sub XWikiConvertH2()
    ReplaceHeading wdOutlineLevel2, "=="
End Sub

Private Sub XWikiConvertH3()
    ReplaceHeading wdOutlineLevel3, "==="
End Sub

Private Sub XWikiConvertH4()
    ReplaceHeading wdOutlineLevel4, "===="
End Sub

Private Sub XWikiConvertH5()
    ReplaceHeading wdOutlineLevel5, "====="
End Sub

Private Function ReplaceHeading(outlineLevel As String, headerPrefix As String)
    ActiveDocument.Select
    With Selection.Find
        .ClearFormatting
        .ParagraphFormat.outlineLevel = outlineLevel
        .Text = ""
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Forward = True
        .Wrap = wdFindContinue
        Do While .Execute
            With Selection
                If InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If
              
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    .InsertBefore headerPrefix
                    .InsertBefore vbCr
                    .InsertAfter headerPrefix
                End If
                .Style = ActiveDocument.Styles(wdStyleNormal)
            End With
        Loop
    End With
End Function

Private Sub XWikiConvertCites()
    
    On Error Resume Next
    
    ActiveDocument.Select
    With Selection.Find
        .ClearFormatting
        .Style = "Style Style Bold"
        .Text = ""
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Forward = True
        .Wrap = wdFindContinue
        Do While .Execute
            With Selection
                If Len(.Text) > 1 And InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    .InsertBefore "**"
                    .InsertAfter "**"
                End If

                .Style = ActiveDocument.Styles("Default Paragraph Font")
                .Font.Bold = False
            End With
        Loop
    End With
End Sub

Private Sub XWikiConvertItalic()
    ActiveDocument.Select
    With Selection.Find
        .ClearFormatting
        .Font.Italic = True
        .Text = ""
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Forward = True
        .Wrap = wdFindContinue
        Do While .Execute
            With Selection
                If Len(.Text) > 1 And InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    '.InsertBefore "//"
                    '.InsertAfter "//"
                End If

                .Style = ActiveDocument.Styles("Default Paragraph Font")
                .Font.Italic = False
            End With
        Loop
    End With
End Sub

Private Sub XWikiConvertBold()
    ActiveDocument.Select
    With Selection.Find
        .ClearFormatting
        .Font.Bold = True
        .Text = ""
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Forward = True
        .Wrap = wdFindContinue
        Do While .Execute
            With Selection
                If Len(.Text) > 1 And InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                '    .InsertBefore "**"
                '    .InsertAfter "**"
                End If

                .Style = ActiveDocument.Styles("Default Paragraph Font")
                .Font.Bold = False
            End With
        Loop
    End With
End Sub

Private Sub XWikiConvertUnderline()

    ActiveDocument.Select
    With Selection.Find
        .ClearFormatting
        .Font.Underline = True
        .Text = ""
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Forward = True
        .Wrap = wdFindContinue
        Do While .Execute
            With Selection
                If Len(.Text) > 1 And InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    '.InsertBefore "__"
                    '.InsertAfter "__"
                End If
                .Style = ActiveDocument.Styles("Default Paragraph Font")
                .Font.Underline = False
            End With
        Loop
    End With
End Sub

Private Sub XWikiConvertSuperscript()
    ActiveDocument.Select
    With Selection.Find
        .ClearFormatting
        .Font.Superscript = True
        .Text = ""
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Forward = True
        .Wrap = wdFindContinue
        Do While .Execute
            With Selection
                .Text = Trim(.Text)
                If Len(.Text) > 1 And InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If
             
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    .InsertBefore ("^^")
                    .InsertAfter ("^^")
                End If

                .Style = ActiveDocument.Styles("Default Paragraph Font")
                .Font.Superscript = False
            End With
        Loop
    End With
End Sub

Private Sub XWikiConvertSubscript()
    ActiveDocument.Select
    With Selection.Find
        .ClearFormatting
        .Font.Subscript = True
        .Text = ""
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Forward = True
        .Wrap = wdFindContinue
        Do While .Execute
            With Selection
                .Text = Trim(.Text)
                If Len(.Text) > 1 And InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If

                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    .InsertBefore (",,")
                    .InsertAfter (",,")
                End If
                .Style = ActiveDocument.Styles("Default Paragraph Font")
                .Font.Subscript = False
            End With
        Loop
    End With
End Sub

Private Sub XWikiRemoveHighlighting()
    Selection.WholeStory
    Selection.Range.HighlightColorIndex = wdNoHighlight
End Sub

Private Sub XWikiRemoveComments()
    Dim i
    For i = ActiveDocument.Comments.Count To 1 Step -1
        ActiveDocument.Comments(i).Delete
    Next i
End Sub

'*************************************************************************************
'* CASELIST INFO FUNCTIONS                                                                                 *
'*************************************************************************************

Sub GetCaselistSchoolNames(Caselist As String, c As control)

    Dim CaselistURL As String
    Dim Script As String
    Dim XML As String
    
    Dim ExcludedSpaces As String
    Dim SpaceName As String
        
    'Turn on error checking
    On Error GoTo Handler
    
    'Get URL for appropriate caselist
    Select Case Caselist
        Case Is = "openCaselist"
            CaselistURL = GetCaselistURL("openCaselist")
        Case Is = "NDCAPolicy"
            CaselistURL = GetCaselistURL("NDCAPolicy")
        Case Is = "NDCALD"
            CaselistURL = GetCaselistURL("NDCALD")
        Case Else
            CaselistURL = GetCaselistURL("openCaselist")
    End Select
    
    'Exit if error
    If CaselistURL = "HTTP Error" Then
        c.AddItem "Internet error."
        Exit Sub
    End If
         
    'Spaces to exclude from list
    ExcludedSpaces = "Admin;AnnotationCode;Blog;Caselist;ColorThemes;Panels;Sandbox;Scheduler;Stats;XWiki;Dashboard;AppWithinMinutes;Main"
    
    'Set Mouse Pointer
    System.Cursor = wdCursorWait
    
    'Construct curl request and send
    Script = "do shell script ""curl '" & CaselistURL & "'"""
    XML = MacScript(Script)
    
    'Parse XML and add Space name to control if not an excluded space
    Do While InStr(XML, "<name>") > 0
        SpaceName = Mid(XML, InStr(XML, "<name>") + 6, InStr(XML, "</name") - InStr(XML, "<name>") - 6)
        If InStr(ExcludedSpaces, SpaceName) < 1 Then c.AddItem SpaceName
        XML = Mid(XML, InStr(XML, "</name>") + 7)
    Loop
    
    'Reset mouse cursor
    System.Cursor = wdCursorNormal
    Exit Sub

Handler:
    System.Cursor = wdCursorNormal
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Public Sub GetCaselistTeamNames(Caselist As String, School As String, c As control)

    Dim CaselistURL As String
    Dim Script As String
    Dim XML As String
    Dim TeamName As String

    'Turn on error checking
    On Error GoTo Handler
    
    'Get URL for appropriate caselist
    Select Case Caselist
        Case Is = "openCaselist"
            CaselistURL = GetCaselistURL("openCaselist")
        Case Is = "NDCAPolicy"
            CaselistURL = GetCaselistURL("NDCAPolicy")
        Case Is = "NDCALD"
            CaselistURL = GetCaselistURL("NDCALD")
        Case Else
            CaselistURL = GetCaselistURL("openCaselist")
    End Select
    
    'Exit if error
    If CaselistURL = "HTTP Error" Then
        c.AddItem "Internet error."
        Exit Sub
    End If
    
    'Caselist.TeamClass/ removed from URL because pages with deleted objects break
    CaselistURL = CaselistURL & School & "/pages/WebHome/objects/"
    
    'Set Mouse Pointer
    System.Cursor = wdCursorWait
    
    'Construct curl request and send
    Script = "do shell script ""curl '" & CaselistURL & "'"""
    XML = MacScript(Script)
    
    'Parse XML and add Team name to control
    Do While InStr(XML, "<headline>") > 0
        TeamName = Mid(XML, InStr(XML, "<headline>") + 10, InStr(XML, "</headline") - InStr(XML, "<headline>") - 10)
        c.AddItem Split(Left(TeamName, Len(TeamName) - 4), ".")(1)
        XML = Mid(XML, InStr(XML, "</headline>") + 11)
    Loop

    If c.ListCount = 0 Then c.AddItem "No teams found."
        
    'Rest mouse cursor
    System.Cursor = wdCursorNormal
    Exit Sub

Handler:
    System.Cursor = wdCursorNormal
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Public Function GetCaselistURL(Caselist As String) As String

    Dim XML As String
    
    'Turn off error checking
    On Error Resume Next

    XML = MacScript("do shell script ""curl 'http://update.paperlessdebate.com/caselist.xml'""")

    'Exit if the request fails
    If Len(XML) < 1 Then
        GetCaselistURL = "HTTP Error"
        Exit Function
    End If
    
    Select Case Caselist
        Case Is = "openCaselist"
            GetCaselistURL = Mid(XML, InStr(XML, "<opencaselist>") + 14, InStr(XML, "</opencaselist>") - InStr(XML, "<opencaselist>") - 14)
        Case Is = "NDCAPolicy"
            GetCaselistURL = Trim(Mid(XML, InStr(XML, "<hspolicy>") + 10, InStr(XML, "</hspolicy>") - InStr(XML, "<hspolicy>") - 10))
        Case Is = "NDCALD"
            GetCaselistURL = Trim(Mid(XML, InStr(XML, "<hsld>") + 6, InStr(XML, "</hsld>") - InStr(XML, "<hsld>") - 6))
        Case Else
            GetCaselistURL = Trim(Mid(XML, InStr(XML, "<opencaselist>") + 14, InStr(XML, "</opencaselist>") - InStr(XML, "<opencaselist>") - 14))
    End Select
            
    Exit Function

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Function
