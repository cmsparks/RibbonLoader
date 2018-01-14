Attribute VB_Name = "View"
Option Explicit

Sub ArrangeWindows()
   
    Dim CurrentWindow As Window
    Dim w As Window
    Dim l As Long
    Dim t As Long
    Dim LeftSplitPct As Single
    Dim RightSplitPct As Single
    
    On Error GoTo Handler
        
    'Save current window
    Set CurrentWindow = ActiveWindow
            
    'Maximize current window to find usable area edges
    ActiveWindow.WindowState = wdWindowStateMaximize
    l = ActiveWindow.Left
    t = ActiveWindow.Top
    
    'Set to zero if maximized window returns a negative number
    If l < 0 Then l = 0
    If t < 0 Then t = 0
       
    'Get split percentages
    LeftSplitPct = GetSetting("Verbatim", "View", "DocsPct", 50) / 100
    RightSplitPct = GetSetting("Verbatim", "View", "SpeechPct", 50) / 100
        
    'Loop through open windows and organize
    For Each w In Application.Windows
    
        'Windows must not be minimized/maximized to assign properties
        w.WindowState = wdWindowStateNormal
        
        'If it's the ActiveSpeechDoc, put on right
        If ActiveSpeechDoc <> "" And w.Document.Name = ActiveSpeechDoc Then
            w.Width = Application.UsableWidth * RightSplitPct
            w.Left = Application.UsableWidth - w.Width
            
            'If toolbar is on the left, manually adjust to avoid covering window
            If GetSetting("Verbatim", "View", "ToolbarPosition", "Top") = "Left" Then w.Width = w.Width - 50
            If GetSetting("Verbatim", "View", "ToolbarPosition", "Top") = "Left" Then w.Left = w.Left + 50
            
            w.Height = Application.UsableHeight
            w.Top = t
        
        'If no ActiveSpeechDoc, put any doc with "Speech" in the name on the right
        ElseIf InStr(LCase(w.Document.Name), "speech") > 0 Then
            w.Width = Application.UsableWidth * RightSplitPct
            w.Left = Application.UsableWidth - w.Width
            
            'If toolbar is on the left, manually adjust to avoid covering window
            If GetSetting("Verbatim", "View", "ToolbarPosition", "Top") = "Left" Then w.Width = w.Width - 50
            If GetSetting("Verbatim", "View", "ToolbarPosition", "Top") = "Left" Then w.Left = w.Left + 50
            
            w.Height = Application.UsableHeight
            w.Top = t
        
        'Else put doc on the left
        Else
            w.Width = Application.UsableWidth * LeftSplitPct
            w.Left = l
            
            'If toolbar is on the left, manually adjust to avoid covering window
            If GetSetting("Verbatim", "View", "ToolbarPosition", "Top") = "Left" Then w.Width = w.Width - 50
            If GetSetting("Verbatim", "View", "ToolbarPosition", "Top") = "Left" Then w.Left = w.Left + 100
            
            w.Height = Application.UsableHeight
            w.Top = t
        End If
        
    Next w
    
    'Activate original window and clean up
    CurrentWindow.Activate
    Set CurrentWindow = Nothing

    Exit Sub
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Sub SwitchWindows()
'Cycle through all open windows
    
    Dim i As Long
    For i = 1 To Documents.Count
        If Documents(i).Name = ActiveDocument.Name Then Exit For
    Next i
    
    If i = 1 Then i = Documents.Count + 1
    Documents(i - 1).Activate
    
End Sub

Sub ToggleReadingView()

    If ActiveWindow.View.FullScreen = False Then
        ActiveWindow.View.FullScreen = True
    Else
        Call View.DefaultView
    End If
   
End Sub

Sub DefaultView()
    
    If GetSetting("Verbatim", "View", "DefaultView", "Web") = "Web" Then
        ActiveWindow.ActivePane.View.Type = wdOnlineView
    Else
        ActiveWindow.ActivePane.View.Type = wdNormalView
    End If

End Sub

Sub SetZoom()
    ActiveWindow.ActivePane.View.Zoom.Percentage = GetSetting("Verbatim", "View", "ZoomPct", "100")
End Sub

Sub InvisibilityOn()

    Dim p
    Dim pCount As Long
 
    pCount = 0
    
    'Make sure status bar is visible for progress indicator
    Application.StatusBar = True
 
    'Loop each paragraph
    For Each p In ActiveDocument.Paragraphs
        pCount = pCount + 1
        Application.StatusBar = "Processing paragraph " & pCount & " of " & ActiveDocument.Paragraphs.Count
        
        'Select each non-blank body text paragraph
        If p.outlineLevel = wdOutlineLevelBodyText And Len(p) > 1 Then
            p.Range.Select
            
            'Highlight the cites so they don't disappear
            With Selection.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = ""
                .Wrap = wdFindStop
                .Replacement.Text = ""
                .Format = True
                .Style = "Cite"
                .Execute
                
                'Skip the paragraph if cite is found
                If .Found = True Then GoTo Skip
            End With
            
            'Select the paragraph, shorten to keep line breaks
            p.Range.Select
            Selection.MoveEndWhile Cset:=vbCrLf, Count:=-1
            Selection.MoveEndWhile Cset:=" ", Count:=-1
            Selection.MoveStartWhile Cset:=vbCrLf, Count:=1
            Selection.MoveStartWhile Cset:=" ", Count:=1
            
            'Hide all non-highlighted text
            With Selection.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = "[! ]"
                .Wrap = wdFindStop
                .MatchWildcards = True
                .Format = True
                .Highlight = False
                .ParagraphFormat.outlineLevel = wdOutlineLevelBodyText
                .Replacement.Font.Hidden = True
                .Execute Replace:=wdReplaceAll
            End With
            
        End If
Skip:
    Next p

    'Clean up and supress errors
    Selection.Find.ClearFormatting
    Selection.Find.MatchWildcards = False
    Selection.Find.Replacement.ClearFormatting
                    
    ActiveDocument.ShowGrammaticalErrors = False
    ActiveDocument.ShowSpellingErrors = False

End Sub

Sub InvisibilityOff()
    'Set the whole doc visible
    ActiveDocument.Range.Font.Hidden = False
    
    'Turn error checking back on but set it to checked
    ActiveDocument.ShowGrammaticalErrors = False
    ActiveDocument.ShowSpellingErrors = False
    ActiveDocument.GrammarChecked = True
    ActiveDocument.SpellingChecked = True
    ActiveDocument.ShowGrammaticalErrors = True
    ActiveDocument.ShowSpellingErrors = True

End Sub
