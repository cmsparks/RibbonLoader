Attribute VB_Name = "Tutorial"
Option Explicit

Public TutorialStep As Long
Public TutorialDoc As String

Sub LaunchTutorial()

    Dim d As Document
    
    If Application.Version >= "15" Then
        MsgBox "The Verbatim Tutorial only works on Mac Office 2011. For more information, see the online manual on paperlessdebate.com"
        Exit Sub
    End If
    
    'If more than one non-empty doc is open, prompt to close
    If Documents.Count > 1 Or ActiveDocument.Words.Count > 1 Then
        If MsgBox("Tutorial can only be run while a single blank document is open. Open a new blank doc and close everything else?", vbYesNo) = vbYes Then
            TutorialDoc = Documents.Add(ActiveDocument.AttachedTemplate.FullName)
            
            For Each d In Documents
                If d <> TutorialDoc Then d.Close wdPromptToSaveChanges
            Next d
        Else
            Exit Sub
        End If
    End If
    
    'Start tutorial
    TutorialStep = 1
    Call Tutorial.RunTutorial
    
End Sub

Private Sub RunTutorial()
    
    Dim d
    
    On Error GoTo Handler
    
    'Clear the tutorial doc
    Call Tutorial.ClearTutorialDoc
    
    Select Case TutorialStep
    
        'Introduction
        Case Is = 1
            Selection.Style = "Block"
            Selection.TypeText "Step 1/18 - Start" & vbCrLf
            Selection.Style = "Tag"
            Selection.TypeText "Welcome to the Verbatim tutorial!" & vbCrLf
            Selection.TypeText "Just follow the prompts to step through the tutorial." & vbCrLf
            Selection.TypeText "The Status Bar at the bottom of the screen will show you the time remaining on each step." & vbCrLf
            Selection.ClearFormatting
            If MsgBox("This is the interactive Verbatim tutorial. For each step, you will be given a few seconds to read the instructions and try things out, then you will be asked if you're ready to move on." & vbCrLf & vbCrLf & "The complete tutorial will take less than 5 minutes - let's get started!", vbOKCancel) = vbCancel Then Exit Sub
            TutorialStep = TutorialStep + 1
            Call Tutorial.RunTutorial
        
        'Toolbar
        Case Is = 2
            Selection.Style = "Block"
            Selection.TypeText "Step 2/18 - Verbatim Toolbar" & vbCrLf
            Selection.Style = "Tag"
            Selection.TypeText "You should see the Verbatim toolbar at the top or left of the screen - it contains buttons for almost every feature. Many features also have keyboard shortcuts." & vbCrLf
            Selection.TypeText "Features in each step will usually be displayed as the only active button on the toolbar. Some features will be disabled during the tutorial." & vbCrLf
            Selection.ClearFormatting
            If Tutorial.TutorialTimer("10") = True Then
                Call Toolbar.BuildVerbatimToolbar
                Exit Sub
            End If
            Call Tutorial.RunTutorial
                            
        'F keys
        Case Is = 3
            Call Tutorial.TutorialToolbarControls("F2Button,F3Button,F4Button,F5Button,F6Button,F7Button,F8Button,F9Button,F10Button,F11Button,F12Button,ShrinkText")
            
            Selection.Style = "Block"
            Selection.TypeText "Step 3/18 - Formatting" & vbCrLf
            Selection.Style = "Tag"
            Selection.TypeText "The first block of controls on the Verbatim toolbar show basic formatting functions for things like Blocks and Tags, and their corresponding F-key shortcuts. You can configure these shortcuts in the Verbatim settings." & vbCrLf
            Selection.TypeText "Try using some of the F-key shortcuts to paste or format text:" & vbCrLf
            Selection.ClearFormatting
            Selection.TypeText vbCrLf & "For example, if you" & vbCrLf & vbCrLf & "select these four paragraphs" & vbCrLf & vbCrLf & "and press F3, they will be condensed" & vbCrLf & vbCrLf & "to a single paragraph." & vbCrLf & vbCrLf
            Selection.Style = "Tag"
            Selection.TypeText "Remember to always use the Paste function instead of Ctrl-V when pasting from the internet to remove weird formatting!" & vbCrLf
            Selection.ClearFormatting
            If Tutorial.TutorialTimer("20") = True Then
                Call Toolbar.BuildVerbatimToolbar
                Exit Sub
            End If
            Call Tutorial.RunTutorial
            
        'Heading levels
        Case Is = 4
            Call Tutorial.TutorialToolbarControls("F2Button,F3Button,F4Button,F5Button,F6Button,F7Button,F8Button,F9Button,F10Button,F11Button,F12Button,ShrinkText")
            
            Selection.Style = "Block"
            Selection.TypeText "Step 4/18 - File Organization" & vbCrLf
            Selection.Style = "Tag"
            Selection.TypeText "Think of each Word document like an expando - with Pockets, Hats, Blocks and Tags, you have 4 levels available for organizing your files. Note how these levels show up in a hierarchy in the Document Map on the left." & vbCrLf
            Selection.TypeText "Tip: You can move a unit (Pocket/Block/Hat/Tag) up or down in the document hierarchy by putting your cursor in the heading and using the keyboard shortcuts Alt+Up and Alt+Down. Try moving some tags or blocks around below." & vbCrLf
            Selection.ClearFormatting
            Selection.TypeParagraph
            Selection.Style = "Pocket"
            Selection.TypeText "Pocket" & vbCrLf
            Selection.Style = "Hat"
            Selection.TypeText "Hat" & vbCrLf
            Selection.Style = "Block"
            Selection.TypeText "Block" & vbCrLf
            Selection.Style = "Tag"
            Selection.TypeText "Tag 1" & vbCrLf
            Selection.TypeText "Tag 2" & vbCrLf
            Selection.Style = "Block"
            Selection.TypeText "Block 2" & vbCrLf
            Selection.Style = "Tag"
            Selection.TypeText "Tag 3" & vbCrLf
            Selection.TypeText "Tag 4" & vbCrLf
            Selection.ClearFormatting
            If Tutorial.TutorialTimer("20") = True Then
                Call Toolbar.BuildVerbatimToolbar
                Exit Sub
            End If
            Call Tutorial.RunTutorial

        'Send To Speech button
        Case Is = 5
            Call Tutorial.TutorialToolbarControls("SendToSpeech")
            
            Selection.Style = "Block"
            Selection.TypeText "Step 5/18 - Send To Speech" & vbCrLf
            Selection.Style = "Tag"
            Selection.TypeText "The arrow on the toolbar sends the current Pocket, Hat, Block, or Card (or the selected text) to the active Speech document. You can also press the `\~ key instead (next to the number 1 key)." & vbCrLf
            Selection.TypeText "Try it out! Click in the sample text below and try sending to the newly opened speech doc. When you're done with this step, the temporary speech doc will be closed automatically." & vbCrLf
            
            Selection.Style = "Block"
            Selection.TypeText "Block Title" & vbCrLf
            Selection.ClearFormatting
            ActiveDocument.AttachedTemplate.AutoTextEntries("VSCVerbatimSampleCard").Insert Where:=Selection.Range, RichText:=True

            Dim TempSpeechDoc As String
            TempSpeechDoc = Documents.Add(Template:=ActiveDocument.AttachedTemplate.FullName)
            ActiveSpeechDoc = TempSpeechDoc
            Call View.ArrangeWindows
            
            'Re-activate the main document - should be 2nd in Documents collection
            Documents(2).ActiveWindow.Activate

            If Tutorial.TutorialTimer("20") = True Then
                Call Toolbar.BuildVerbatimToolbar
                Exit Sub
            End If
            Call Tutorial.RunTutorial

        'Speech doc chooser
        Case Is = 6
            For Each d In Documents
                If d = ActiveSpeechDoc Then Documents(ActiveSpeechDoc).Close wdDoNotSaveChanges
            Next d
            
            'Re-clear document after speech doc closed
            Call Tutorial.ClearTutorialDoc

            Call Tutorial.TutorialToolbarControls("ChooseSpeechDoc")
            
            Selection.Style = "Block"
            Selection.TypeText "Step 6/18 - Speech Documents" & vbCrLf
            Selection.Style = "Tag"
            Selection.TypeText "By default, any document with the name ""Speech"" in the name will be your speech document for sending things to. You can create a speech document with the New Speech button." & vbCrLf
            Selection.TypeText "If you use the drop-down menu next to it instead, it will let you select from a list of pre-selected speech names, including ones auto-detected from the tournament you're at." & vbCrLf
            Selection.TypeText "Tip: For the auto-naming functions to work, you must enter your tabroom.com username and password in the Verbatim settings, and the tournament you are attending must be run on Tabroom." & vbCrLf & vbCrLf
            Selection.TypeText "The Choose Speech Doc button lets you choose ANY document you want as the current speech document." & vbCrLf
            Selection.TypeText "Try clicking the button to see the Choose Speech Doc form - close the form when you're ready to proceed." & vbCrLf
            If Tutorial.TutorialTimer("15") = True Then
                Call Toolbar.BuildVerbatimToolbar
                Exit Sub
            End If
            Call Tutorial.RunTutorial
            
        'Windows arranger
        Case Is = 7
            Call Tutorial.TutorialToolbarControls("WindowArranger")
            
            Selection.Style = "Block"
            Selection.TypeText "Step 7/18 - Window Arranger" & vbCrLf
            Selection.Style = "Tag"
            Selection.TypeText "The Window Arranger button helps arrange your docs for greater efficiency, split-screen with your Speech on the right (like in the Send to Speech step). Try dragging the window out of place and then using the button to organize it again." & vbCrLf
            Selection.TypeText "Tip: You can configure the layout of the automatic window arranger in the Verbatim Settings." & vbCrLf
            Selection.TypeText "The keyboard shortcut for automatically arranging your open windows is Ctrl+Shift+Tab." & vbCrLf

            If Tutorial.TutorialTimer("10") = True Then
                Call Toolbar.BuildVerbatimToolbar
                Exit Sub
            End If
            Call Tutorial.RunTutorial

        'VTub
        Case Is = 8
            
            Call Tutorial.TutorialToolbarControls("VTubMenu")

            Selection.Style = "Block"
            Selection.TypeText "Step 8/18 - Virtual Tub" & vbCrLf
            Selection.Style = "Tag"
            Selection.TypeText "The VTub menu opens your ""Virtual Tub,"" which lets you insert sections of documents without needing to actually open them. It must be configured in the Verbatim Settings before use." & vbCrLf
            Selection.TypeText "Tip: The Virtual Tub is designed to be used with a relatively small number of files that are very well organized - it's not meant for your entire Tub. The VTub can be tricky to set up - make sure to read the manual if you're having trouble." & vbCrLf

            If Tutorial.TutorialTimer("10") = True Then
                Call Toolbar.BuildVerbatimToolbar
                Exit Sub
            End If
            Call Tutorial.RunTutorial
            
        'Search
        Case Is = 9
            Call Tutorial.TutorialToolbarControls("Search")
        
            Selection.Style = "Block"
            Selection.TypeText "Step 9/18 - Search" & vbCrLf
            Selection.Style = "Tag"
            Selection.TypeText "You can type a search term into the search box and press Enter - the dropdown menu will then contain a list of documents on your computer which contain that phrase, which you can open just by clicking. The results are disabled during this tutorial." & vbCrLf
            Selection.TypeText "Tip: By default, the Search box will search everything under your Documents folder. You can set a more specific search location in the Verbatim settings." & vbCrLf
            
            If Tutorial.TutorialTimer("15") = True Then
                Call Toolbar.BuildVerbatimToolbar
                Exit Sub
            End If
            Call Tutorial.RunTutorial
            
        'Tools Menu
        Case Is = 10
                        
            Call Tutorial.TutorialToolbarControls("Tools")
        
            CommandBars.FindControl(Tag:="Tools").Controls(3).Enabled = False 'Doc combiner
            CommandBars.FindControl(Tag:="Tools").Controls(7).Enabled = False 'Convert Backfile
            CommandBars.FindControl(Tag:="Tools").Controls(8).Enabled = False 'Auto Open Folder
            
            Selection.Style = "Block"
            Selection.TypeText "Step 10/18 - Tools Menu" & vbCrLf
            Selection.Style = "Tag"
            Selection.TypeText "The tools menu contains useful features like quick access to any Timer program you have installed, an audio recorder for quickly capturing speeches, and a ""Stats"" window that will estimate how long your speech doc would take to read." & vbCrLf
            Selection.TypeText "For suggestions on a Mac timer program, see the Verbatim manual." & vbCrLf
            Selection.TypeText "The Speech Combiner starts a wizard that lets you quickly combine documents, for example to combine speech docs into one post-round document for the judge." & vbCrLf
            Selection.TypeText "You can configure the ""Auto Open"" folder in the Verbatim settings. It will watch any folder you choose (e.g. a PaDS or Dropbox folder) and automatically open any new document which appears there." & vbCrLf
            Selection.TypeText "The default directory for saving recorded audio can be configured in the Verbatim settings. You can also configure your words-per-minute count for a more accurate time estimate in the Stats form." & vbCrLf
            Selection.ClearFormatting
            
            If Tutorial.TutorialTimer("20") = True Then
                Call Toolbar.BuildVerbatimToolbar
                Exit Sub
            End If
            Call Tutorial.RunTutorial
        
        'Share menu
        Case Is = 11
            Call Tutorial.TutorialToolbarControls("Share")
            
            Selection.Style = "Block"
            Selection.TypeText "Step 11/18 - Share Menu" & vbCrLf
            Selection.Style = "Tag"
            Selection.TypeText "The Share menu lets you quickly share a speech document via USB, Email, or a public PaDS folder." & vbCrLf
            Selection.TypeText "The USB option (Cmd+Shift+S) will copy your current document to the root of any plugged in USB drive - even more than one at a time!" & vbCrLf
            Selection.TypeText "To use the Email functions, you must first set up an account in Apple Mail - it works easily with Gmail, and you can find more instructions in the Verbatim manual. It will also try to automatically look up your opponents email addresses from tabroom.com - make sure you've entered your Tabroom username and password." & vbCrLf
            
            If Tutorial.TutorialTimer("15") = True Then
                Call Toolbar.BuildVerbatimToolbar
                Exit Sub
            End If
            Call Tutorial.RunTutorial

        'View menu
        Case Is = 12
            Call Tutorial.TutorialToolbarControls("View")
            
            CommandBars.FindControl(Tag:="View").Controls(2).Enabled = False 'Reading Mode
            
            Selection.Style = "Block"
            Selection.TypeText "Step 12/18 - View Menu" & vbCrLf
            Selection.Style = "Tag"
            Selection.TypeText "The View menu lets you quickly change your view, like toggling the Navigation Pane or switching between Web and Read view." & vbCrLf
            Selection.TypeText "Tip: If you prefer ""Draft"" view to Web View, you can configure your default view in the Verbatim settings." & vbCrLf
            Selection.TypeText "While in Reading View, you can advance pages using the arrow keys or mousewheel." & vbCrLf
            Selection.TypeText """Invisibility Mode,"" temporarily hides all non-highlighted card text for easier reading or judging. Go to the next step to see it in action on the cards below." & vbCrLf
            Selection.ClearFormatting
            ActiveDocument.AttachedTemplate.AutoTextEntries("VSCVerbatimSampleCard").Insert Where:=Selection.Range, RichText:=True
            ActiveDocument.AttachedTemplate.AutoTextEntries("VSCVerbatimSampleCard").Insert Where:=Selection.Range, RichText:=True
            
            If Tutorial.TutorialTimer("10") = True Then
                Call Toolbar.BuildVerbatimToolbar
                Exit Sub
            End If
            Call Tutorial.RunTutorial

        'Invisibility
        Case Is = 13
        
            Call Tutorial.TutorialToolbarControls("InvisibilityMode")
            
            Selection.Style = "Block"
            Selection.TypeText "Step 13/18 - Invisibility Mode" & vbCrLf
            Selection.Style = "Tag"
            Selection.TypeText "The cards below have invisibility mode turned on - go to the next step to turn it back off." & vbCrLf
            Selection.ClearFormatting
            ActiveDocument.AttachedTemplate.AutoTextEntries("VSCVerbatimSampleCard").Insert Where:=Selection.Range, RichText:=True
            ActiveDocument.AttachedTemplate.AutoTextEntries("VSCVerbatimSampleCard").Insert Where:=Selection.Range, RichText:=True
            
            Call View.InvisibilityOn
            
            If Tutorial.TutorialTimer("10") = True Then
                Call Toolbar.BuildVerbatimToolbar
                Exit Sub
            End If
            Call Tutorial.RunTutorial

        'Coauthor Menu
        Case Is = 14
            Call View.InvisibilityOff
            Call Tutorial.TutorialToolbarControls("CoauthoringMenu")
            
            Selection.Style = "Block"
            Selection.TypeText "Step 14/18 - Coauthoring" & vbCrLf
            Selection.Style = "Tag"
            Selection.TypeText "If you have a PaDS account, the Coauthor menu will let you quickly upload or open documents for coauthoring with other people." & vbCrLf
            Selection.TypeText """Coauthoring"" is when multiple people edit the same Word document simultaneously, for example partners prepping the same speech document." & vbCrLf
            Selection.TypeText "You can configure the default locations for using PaDS in the Verbatim settings." & vbCrLf
            Selection.TypeText "For more info on PaDS, check out: " & vbCrLf & "http://paperlessdebate.com/pads" & vbCrLf

            If Tutorial.TutorialTimer("10") = True Then
                Call Toolbar.BuildVerbatimToolbar
                Exit Sub
            End If
            Call Tutorial.RunTutorial
            
        'Format Menu
        Case Is = 15
            Call Tutorial.TutorialToolbarControls("Format")
            
            CommandBars.FindControl(Tag:="Format").Controls(2).Enabled = False 'Underline Mode
            CommandBars.FindControl(Tag:="Format").Controls(19).Enabled = False 'Get From CiteMaker
            
            Selection.Style = "Block"
            Selection.TypeText "Step 15/18 - Format Menu" & vbCrLf
            Selection.Style = "Tag"
            Selection.TypeText "The Format Menu contains a number of useful features." & vbCrLf
            Selection.TypeText "The Automatic Underliner attempts to automatically underline a card based on the tag - make sure to check the results! To use the automatic underliner, your cursor must be on the tag. The better and more specific your tag, the better it will work." & vbCrLf
            Selection.TypeText "You can ""Shrink"" all cards in the document, remove unwanted pilcrows, blanks, hyperlinks, or emphasis, standardize highlighting, automatically format cites, or add numbers to all your tags." & vbCrLf
            Selection.TypeText "If you have the CiteMaker extension for Chrome installed, you can also automatically paste in the cite from the top tab in Chrome. For more information on the CiteMaker extension, see paperlessdebate.com."
            
            If Tutorial.TutorialTimer("20") = True Then
                Call Toolbar.BuildVerbatimToolbar
                Exit Sub
            End If
            Call Tutorial.RunTutorial
            
        'Caselist Menu
        Case Is = 16
            Call Tutorial.TutorialToolbarControls("Caselist")
            
            CommandBars.FindControl(Tag:="Caselist").Controls(1).Enabled = False 'Caselist Wizard
            CommandBars.FindControl(Tag:="Caselist").Controls(2).Enabled = False 'Wikify
            CommandBars.FindControl(Tag:="Caselist").Controls(3).Enabled = False 'Cite Request Doc
            
            Selection.Style = "Block"
            Selection.TypeText "Step 16/18 - Caselist Menu" & vbCrLf
            Selection.Style = "Tag"
            Selection.TypeText "The Caselist Menu lets you automatically upload cites or open source documents to the caselist of your choice, or convert your docs to cites and/or wiki syntax for manual posting." & vbCrLf
            Selection.TypeText "Caselist functions can be configured in the Verbatim settings, and require you to set up a tabroom.com account." & vbCrLf
            
            If Tutorial.TutorialTimer("10") = True Then
                Call Toolbar.BuildVerbatimToolbar
                Exit Sub
            End If
            Call Tutorial.RunTutorial

        'Link and Help
        Case Is = 17
            Call Tutorial.TutorialToolbarControls("SettingsMenu")
            
            Selection.Style = "Block"
            Selection.TypeText "Step 17/18 - Settings Menu" & vbCrLf
            Selection.Style = "Tag"
            Selection.TypeText "The Settings menu lets you open the main Verbatim settings, where you can configure all the features you've just seen, or lets you get more help from paperlessdebate.com or the built-in Verbatim help." & vbCrLf
            Selection.TypeText "Tip: You can also open the Verbatim help at any time by pressing F1." & vbCrLf
            Selection.TypeText "The Cheat Sheet opens a handy info sheet of all the Verbatim keyboard shortcuts." & vbCrLf

            If Tutorial.TutorialTimer("10") = True Then
                Call Toolbar.BuildVerbatimToolbar
                Exit Sub
            End If
            Call Tutorial.RunTutorial

        'Finish
        Case Is = 18
            Call Toolbar.BuildVerbatimToolbar
            
            Selection.Style = "Block"
            Selection.TypeText "Step 18/18 - Finish" & vbCrLf
            Selection.Style = "Tag"
            Selection.TypeText "That's it! For more information read the built-in help or the manual on paperlessdebate.com" & vbCrLf
            
        Case Else
            Exit Sub
            
    End Select
    
    Exit Sub
    
Handler:
    Call Toolbar.BuildVerbatimToolbar
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Private Function TutorialTimer(Timer As String) As Boolean

    Dim NextStep As Boolean
    Dim StartTime

    'Loop until the user says to move on
    Do While NextStep = False
        
        'Empty loop for the timer length, return control to application
        StartTime = Now()
        Do Until Now() > StartTime + TimeValue("00:00:" & Timer)
            DoEvents
            'Update status bar with step timer
            Application.StatusBar = "You have " & CLng(Timer) - DateDiff("s", StartTime, Now) + 1 & " seconds remaining on this tutorial step."
        Loop
        
        'Once timer runs out, prompt whether to move on
        Select Case MsgBox("Are you ready to move on to the next step?" & vbCrLf & vbCrLf & "Select ""Yes"" to move on, ""No"" for additional time on this step, or ""Cancel"" to exit the tutorial.", vbYesNoCancel)
            Case Is = vbYes
                'Increment tutorial step and exit loop to go back to main sub
                TutorialStep = TutorialStep + 1
                NextStep = True
            Case Is = vbNo
                'Nothing, stay in loop
            Case Is = vbCancel
                'Return true to tell main sub to quit
                TutorialTimer = True
                Exit Function
        End Select
    Loop
    
End Function

Private Sub ClearTutorialDoc()
    Selection.WholeStory
    Selection.Delete
    Selection.ClearFormatting
End Sub

Private Sub TutorialToolbarControls(TutorialControls As String)
    
    Dim TutorialControlArray
    Dim tc
    Dim VerbatimToolbar As CommandBar
    Dim c

    On Error GoTo Handler

    'Split the control array
    TutorialControlArray = Split(TutorialControls, ",")

    'Get the Verbatim Toolbar
    Set VerbatimToolbar = Application.CommandBars("Verbatim")

    'Disable all controls except tutorial control
    If TutorialControls <> "" Then
    
        For Each c In VerbatimToolbar.Controls
            c.Enabled = False
        Next c
    
        For Each tc In TutorialControlArray
            For Each c In VerbatimToolbar.Controls
                If c.Tag = tc Then c.Enabled = True
            Next c
        Next tc
    End If

    'Clean up
    Set VerbatimToolbar = Nothing
    
    Exit Sub
    
Handler:
    Set VerbatimToolbar = Nothing
    Call Toolbar.BuildVerbatimToolbar
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub
