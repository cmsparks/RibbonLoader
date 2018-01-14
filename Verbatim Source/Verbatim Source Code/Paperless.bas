Attribute VB_Name = "Paperless"
Option Explicit

Public ActiveSpeechDoc As String

'*************************************************************************************
'* TOOLBAR FUNCTIONS                                                                                         *
'*************************************************************************************

Sub ShowChooseSpeechDoc()

    Dim ChooseSpeechDocForm As frmChooseSpeechDoc
    Set ChooseSpeechDocForm = New frmChooseSpeechDoc
    ChooseSpeechDocForm.Show

End Sub

Sub AutoOpenFolder()
'Runs in the background to automatically open all documents in the speech folder.
    
    Dim PressedControl As CommandBarControl
    
    Dim AutoOpenDir As String
    
    Dim FileList
    Dim f
    
    Dim d As Document
    Dim IsOpen As Boolean
       
    On Error GoTo Handler
    
    'If Word 2011, get pressed control
    If Application.Version < "15" Then
        Set PressedControl = CommandBars.ActionControl
        If PressedControl Is Nothing Then Exit Sub
    End If
    
    If AutoOpenFolderToggle = False Then
       
        'Check for auto open folder
        AutoOpenDir = GetSetting("Verbatim", "Paperless", "AutoOpenDir", "?")
        If AutoOpenDir = "?" Then
            If MsgBox("You have not set an Auto Open folder. Open settings now?", vbYesNo) = vbYes Then
                Call Settings.ShowSettingsForm
                AutoOpenFolderToggle = False
                Exit Sub
            Else
                AutoOpenFolderToggle = False
                Exit Sub
            End If
        End If
    
        'Ensure a trailing :
        If Right(AutoOpenDir, 1) <> ":" Then AutoOpenDir = AutoOpenDir & ":"
    
        'Notify it's turned on
        If MsgBox("This will start a listener that automatically opens all documents in the root of your Auto Open folder:" & vbCrLf & AutoOpenDir, vbOKCancel) = vbCancel Then
            AutoOpenFolderToggle = False
            Exit Sub
        End If
        
        AutoOpenFolderToggle = True
        If Application.Version < "15" Then PressedControl.Caption = "Turn Off Auto Open Folder"
        
        'Loop until unpressed
        Do
            DoEvents
            FileList = Split(Filesystem.GetFilesInFolder(AutoOpenDir), Chr(10))
            
            'Loop all files - if not open, open it
            For Each f In FileList
                IsOpen = False
                For Each d In Application.Documents
                    If d.FullName = f Then IsOpen = True
                Next d
                
                If IsOpen = False And (Right(f, 3) = "doc" Or Right(f, 4) = "docx" Or Right(f, 3) = "rtf") Then Documents.Open f
            Next f
        Loop Until AutoOpenFolderToggle = False
    
    'Turn it off
    Else
        AutoOpenFolderToggle = False
        If Application.Version < "15" Then PressedControl.Caption = "Turn On Auto Open Folder"
        MsgBox "Stopped listening to the Auto Open folder.", vbInformation
    End If
    
    Set PressedControl = Nothing
    
    Exit Sub
    
Handler:
    AutoOpenFolderToggle = False
    If Application.Version < "15" Then PressedControl.Caption = "Turn On Auto Open Folder"
    Set PressedControl = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Sub GetSpeeches(Optional FromScratch As Boolean)

    Dim Menu As CommandBarControl
    Dim MenuItem As CommandBarButton
    Dim c
    
    Dim RoundArray As Variant
    Dim i
    Dim Tournament As String
    Dim RoundNum As String
    Dim Side As String
    Dim Opponent As String
    
    On Error GoTo Handler
    
    'Find Speeches Menu
    Set Menu = CommandBars.FindControl(Tag:="NewSpeechMenu")
    
    'If rebuilding from scratch, remove existing controls
    If FromScratch = True Then
        For Each c In Menu.Controls
            c.Delete
        Next c
    End If
    
    'Exit if the menu is already built
    If Menu.Controls.Count > 0 Then Exit Sub
       
    'Set Mouse Pointer and update progress bar
    System.Cursor = wdCursorWait
    ProgressBar = "Getting round info from Tabroom.com "
    Application.StatusBar = ProgressBar
       
    'Get rounds from tabroom - use email setting to limit rounds and make faster
    RoundArray = Tabroom.GetTabroomRounds(Email:=True)
    
    If IsArray(RoundArray) Then
        If UBound(RoundArray, 1) > 0 Then
            For i = 0 To 1
                'Update Progress Bar
                ProgressBar = ProgressBar & ChrW(9609)
                Application.StatusBar = ProgressBar
    
                Tournament = Trim(RoundArray(i, 0))
                RoundNum = Trim(RoundArray(i, 1))
                Side = Trim(RoundArray(i, 2))
                Opponent = Trim(RoundArray(i, 3))
                
                Select Case RoundNum
                    Case "1", "2", "3", "4", "5", "6", "7", "8"
                        RoundNum = "Round " & RoundNum
                    Case Else
                        'Do nothing
                End Select
                
                If Side = "Aff" Then
                    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
                    MenuItem.Caption = "2AC" & " " & Tournament & " " & RoundNum & " vs " & Opponent
                    MenuItem.Tag = "2AC" & " " & Tournament & " " & RoundNum & " vs " & Opponent
                    MenuItem.FaceId = 1717 'Page with arrow. 3813 = Page with plus
                    MenuItem.OnAction = "Paperless.NewSpeechFromMenu"
                    
                    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
                    MenuItem.Caption = "1AR" & " " & Tournament & " " & RoundNum & " vs " & Opponent
                    MenuItem.Tag = "1AR" & " " & Tournament & " " & RoundNum & " vs " & Opponent
                    MenuItem.FaceId = 1717 'Page with arrow. 3813 = Page with plus
                    MenuItem.OnAction = "Paperless.NewSpeechFromMenu"
                    
                    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
                    MenuItem.Caption = "2AR" & " " & Tournament & " " & RoundNum & " vs " & Opponent
                    MenuItem.Tag = "2AR" & " " & Tournament & " " & RoundNum & " vs " & Opponent
                    MenuItem.FaceId = 1717 'Page with arrow. 3813 = Page with plus
                    MenuItem.OnAction = "Paperless.NewSpeechFromMenu"
                    
                    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
                    MenuItem.Caption = "1AC" & " " & Tournament & " " & RoundNum & " vs " & Opponent
                    MenuItem.Tag = "1AC" & " " & Tournament & " " & RoundNum & " vs " & Opponent
                    MenuItem.FaceId = 1717 'Page with arrow. 3813 = Page with plus
                    MenuItem.OnAction = "Paperless.NewSpeechFromMenu"
                    
                    'Separator
                    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
                    MenuItem.Caption = ""
                    MenuItem.Tag = "SpeechSeparator" & i
                    MenuItem.Enabled = False
                
                Else
                    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
                    MenuItem.Caption = "1NC" & " " & Tournament & " " & RoundNum & " vs " & Opponent
                    MenuItem.Tag = "1NC" & " " & Tournament & " " & RoundNum & " vs " & Opponent
                    MenuItem.FaceId = 1717 'Page with arrow. 3813 = Page with plus
                    MenuItem.OnAction = "Paperless.NewSpeechFromMenu"
                    
                    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
                    MenuItem.Caption = "2NC" & " " & Tournament & " " & RoundNum & " vs " & Opponent
                    MenuItem.Tag = "2NC" & " " & Tournament & " " & RoundNum & " vs " & Opponent
                    MenuItem.FaceId = 1717 'Page with arrow. 3813 = Page with plus
                    MenuItem.OnAction = "Paperless.NewSpeechFromMenu"
                    
                    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
                    MenuItem.Caption = "1NR" & " " & Tournament & " " & RoundNum & " vs " & Opponent
                    MenuItem.Tag = "1NR" & " " & Tournament & " " & RoundNum & " vs " & Opponent
                    MenuItem.FaceId = 1717 'Page with arrow. 3813 = Page with plus
                    MenuItem.OnAction = "Paperless.NewSpeechFromMenu"
                    
                    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
                    MenuItem.Caption = "2NR" & " " & Tournament & " " & RoundNum & " vs " & Opponent
                    MenuItem.Tag = "2NR" & " " & Tournament & " " & RoundNum & " vs " & Opponent
                    MenuItem.FaceId = 1717 'Page with arrow. 3813 = Page with plus
                    MenuItem.OnAction = "Paperless.NewSpeechFromMenu"
                    
                    'Separator
                    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
                    MenuItem.Caption = ""
                    MenuItem.Tag = "SpeechSeparator" & i
                    MenuItem.Enabled = False
    
                End If
                
            Next
        End If
    End If
    
    'Update progress bar
    ProgressBar = ProgressBar & ChrW(9609)
    Application.StatusBar = ProgressBar
    
    'Add default speech options
    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "2AC"
    MenuItem.Tag = "2AC"
    MenuItem.FaceId = 1717 'Page with arrow. 3813 = Page with plus
    MenuItem.OnAction = "Paperless.NewSpeechFromMenu"
    
    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "1AR"
    MenuItem.Tag = "1AR"
    MenuItem.FaceId = 1717 'Page with arrow. 3813 = Page with plus
    MenuItem.OnAction = "Paperless.NewSpeechFromMenu"
    
    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "2AR"
    MenuItem.Tag = "2AR"
    MenuItem.FaceId = 1717 'Page with arrow. 3813 = Page with plus
    MenuItem.OnAction = "Paperless.NewSpeechFromMenu"
    
    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "1AC"
    MenuItem.Tag = "1AC"
    MenuItem.FaceId = 1717 'Page with arrow. 3813 = Page with plus
    MenuItem.OnAction = "Paperless.NewSpeechFromMenu"
    
    'Separator
    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = ""
    MenuItem.Tag = "SpeechSeparator3"
    MenuItem.Enabled = False
    
    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "1NC"
    MenuItem.Tag = "1NC"
    MenuItem.FaceId = 1717 'Page with arrow. 3813 = Page with plus
    MenuItem.OnAction = "Paperless.NewSpeechFromMenu"
    
    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "2NC"
    MenuItem.Tag = "2NC"
    MenuItem.FaceId = 1717 'Page with arrow. 3813 = Page with plus
    MenuItem.OnAction = "Paperless.NewSpeechFromMenu"
    
    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "1NR"
    MenuItem.Tag = "1NR"
    MenuItem.FaceId = 1717 'Page with arrow. 3813 = Page with plus
    MenuItem.OnAction = "Paperless.NewSpeechFromMenu"
    
    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "2NR"
    MenuItem.Tag = "2NR"
    MenuItem.FaceId = 1717 'Page with arrow. 3813 = Page with plus
    MenuItem.OnAction = "Paperless.NewSpeechFromMenu"
    
    'Separator
    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = ""
    MenuItem.Tag = "SpeechSeparator4"
    MenuItem.Enabled = False
    
    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "New Speech"
    MenuItem.DescriptionText = "New Speech (Ctrl+Shift+N)"
    MenuItem.FaceId = 1717 'Page with arrow. 3813 = Page with plus
    MenuItem.Tag = "NewSpeech1"
    MenuItem.OnAction = "Toolbar.AssignButtonActions"
    
    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "New Document"
    MenuItem.DescriptionText = "New Verbatim Document"
    MenuItem.FaceId = 1544 'Blank page
    MenuItem.Tag = "NewDocument"
    MenuItem.OnAction = "Toolbar.AssignButtonActions"

    Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Refresh Speeches"
    MenuItem.DescriptionText = "Refresh New Speech Menu"
    MenuItem.FaceId = 8085 'Blue refresh
    MenuItem.Tag = "RefreshSpeeches"
    MenuItem.OnAction = "Toolbar.AssignButtonActions"

    'Set template as saved to avoid prompts
    ActiveDocument.AttachedTemplate.Saved = True
    
    'Clean Up
    System.Cursor = wdCursorNormal
    Application.StatusBar = "Successfully fetched speeches from Tabroom.com"
    Set Menu = Nothing
    Set MenuItem = Nothing

    Exit Sub
    
Handler:
    System.Cursor = wdCursorNormal
    Set Menu = Nothing
    Set MenuItem = Nothing
    Application.StatusBar = "Failed to fetch speeches from Tabroom.com"
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Sub NewSpeechFromMenu()

    Dim AutoSaveDirectory As String
    Dim FileName As String
    Dim h
    Dim PressedControl As CommandBarButton
    
    'Get filename from tag of most recently pressed control
    Set PressedControl = CommandBars.ActionControl

    FileName = PressedControl.Tag
    If FileName = "" Then
        Set PressedControl = Nothing
        Exit Sub
    End If
    
    'Add a new document based on the template
    Call Paperless.NewDocument
   
    'If Tag is just the speech name, add a date
    If Len(FileName) = 3 Then
        If Hour(Now) > 12 Then h = Hour(Now) - 12 & "PM"
        If Hour(Now) <= 12 Then h = Hour(Now) & "AM"
        FileName = FileName & " " & Month(Now) & "-" & Day(Now) & " " & h
    End If
    
    'Add speech to the name
    FileName = "Speech " & FileName
    
    'Autosave or open save dialog
    If GetSetting("Verbatim", "Paperless", "AutoSaveSpeech", False) = True Then
        AutoSaveDirectory = Trim(GetSetting("Verbatim", "Paperless", "AutoSaveDir", CurDir()))
        If Right(AutoSaveDirectory, 1) <> ":" Then AutoSaveDirectory = AutoSaveDirectory & ":"
        FileName = AutoSaveDirectory & FileName
        ActiveDocument.SaveAs FileName:=FileName, FileFormat:=wdFormatXMLDocument
    Else
        With Application.Dialogs(wdDialogFileSaveAs)
            .Name = FileName
            If .Show = 0 Then Exit Sub
        End With
    End If

    Set PressedControl = Nothing
    
    Exit Sub

Handler:
    Set PressedControl = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

'*************************************************************************************
'* MOVE AND SELECT FUNCTIONS                                                                          *
'*************************************************************************************

Public Sub SelectHeadingAndContent()

    Dim OLevel As Integer
    
    'Move to start of current paragraph and collapse the selection
    Selection.StartOf Unit:=wdParagraph
    Selection.Collapse
        
    'Move backwards through each paragraph to find the first tag, block title, hat, pocket or the top of the document
    Do While True
        If Selection.Paragraphs.outlineLevel < wdOutlineLevel5 Then Exit Do 'Headings 1-4
        If Selection.Start <= ActiveDocument.Range.Start Then 'Top of document
            Application.StatusBar = "Nothing found to select"
            Exit Sub
        End If
        Selection.Move Unit:=wdParagraph, Count:=-1
    Loop
        
    'Get current outline level
    OLevel = Selection.Paragraphs.outlineLevel
    
    'Extend selection until hitting the bottom or a bigger outline level
    Selection.MoveEnd Unit:=wdParagraph, Count:=1
    Do While True And Selection.End <> ActiveDocument.Range.End
        Selection.MoveEnd Unit:=wdParagraph, Count:=1
        If Selection.Paragraphs.Last.outlineLevel <= OLevel Then
            Selection.MoveEnd Unit:=wdParagraph, Count:=-1
            Exit Do 'Bigger Outline Level
        End If
    Loop

End Sub

Sub MoveUp()
'Moves the current pocket, hat, block, or tag, up one level in the document outline

    Dim OLevel As Long
    Dim CurrentView As Long
    Dim StartLocation As Long
    
    On Error GoTo Handler
    
    Application.ScreenUpdating = False
    
    'Save current view
    CurrentView = ActiveWindow.ActivePane.View.Type
    
    'Move to start of current paragraph and collapse the selection
    Selection.StartOf Unit:=wdParagraph
    Selection.Collapse
    
    'Move backwards through each paragraph to find the first tag, block title, hat, pocket, or the top of the document
    Do While True
        If Selection.Start <= ActiveDocument.Range.Start Then Exit Sub 'Top of doc
        If Selection.Paragraphs.outlineLevel < wdOutlineLevel5 Then Exit Do 'Headings 1-4
        Selection.Move Unit:=wdParagraph, Count:=-1
    Loop
        
    'Get current outline level
    OLevel = Selection.Paragraphs.outlineLevel
    
    'Check to make sure you're not moving a card above a block
    If OLevel = 4 Then
        StartLocation = Selection.Start 'Save current location
        Do While True
            Selection.Move Unit:=wdParagraph, Count:=-1
            If Selection.Start <= ActiveDocument.Range.Start Then
                Selection.Start = StartLocation
                Exit Sub
            End If
            If Selection.Paragraphs.outlineLevel = wdOutlineLevel4 Then
                Selection.Start = StartLocation
                Exit Do
            End If
            If Selection.Paragraphs.outlineLevel < wdOutlineLevel4 Then
                Application.StatusBar = "Already the first card on this block"
                Selection.Start = StartLocation
                Exit Sub
            End If
        Loop
    End If
    
    'Switch to outline view and collapse to current level
    ActiveWindow.ActivePane.View.Type = wdOutlineView
    ActiveWindow.View.ShowHeading OLevel
    
    'Move up
    'Selection.Range.Relocate wdRelocateUp - CRASHES WORD 2013
    Application.Run "OutlineMoveUp"
    Selection.Collapse

    'Switch back to previous view
    ActiveWindow.ActivePane.View.Type = CurrentView
    
    Application.ScreenUpdating = True

    Exit Sub
    
Handler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Sub MoveDown()
'Moves the current pocket, hat, block, or tag down one level in the document outline

    Dim OLevel As Long
    Dim CurrentView As Long
    Dim StartLocation As Long
    
    On Error GoTo Handler
    Application.ScreenUpdating = False
    
    'Save current view
    CurrentView = ActiveWindow.ActivePane.View.Type
    
    'Move to start of current paragraph and collapse the selection
    Selection.StartOf Unit:=wdParagraph
    Selection.Collapse
    
    'Move backwards through each paragraph to find the first tag, block title, hat, pocket, or the top of the document
    Do While True
        If Selection.Paragraphs.outlineLevel < wdOutlineLevel5 Then
            Exit Do 'Headings 1-4
        Else
            Application.StatusBar = "Nothing found to move"
            Exit Sub
        End If
        Selection.Move Unit:=wdParagraph, Count:=-1
    Loop
        
    'Get current outline level
    OLevel = Selection.Paragraphs.outlineLevel
    
    'Check to make sure you're not already at the bottom
    StartLocation = Selection.Start 'Save current location
    Do While True
        Selection.Move Unit:=wdParagraph, Count:=1
        If Selection.End + 1 >= ActiveDocument.Range.End Then
                Selection.Start = StartLocation
                Selection.Collapse
                Application.StatusBar = "Already at the bottom"
                Exit Sub
        End If
        If Selection.Paragraphs.outlineLevel <= OLevel Then
            Selection.Start = StartLocation
            Selection.Collapse
            Exit Do
        End If
    Loop
    
    'Check to make sure you're not moving a card off a block or the bottom
    If OLevel = 4 Then
        StartLocation = Selection.Start 'Save current location
        Do While True
            Selection.Move Unit:=wdParagraph, Count:=1
            If Selection.End + 1 >= ActiveDocument.Range.End Then
                Selection.Start = StartLocation
                Selection.Collapse
                Exit Sub
            End If
            If Selection.Paragraphs.outlineLevel = wdOutlineLevel4 Then
                Selection.Start = StartLocation
                Selection.Collapse
                Exit Do
            End If
            If Selection.Paragraphs.outlineLevel < wdOutlineLevel4 Then
                Application.StatusBar = "Already the last card on this block"
                Selection.Start = StartLocation
                Selection.Collapse
                Exit Sub
            End If
        Loop
    End If
    
    'Switch to outline view and collapse to current level
    ActiveWindow.ActivePane.View.Type = wdOutlineView
    ActiveWindow.View.ShowHeading OLevel

    'Move down
    'Selection.Range.Relocate wdRelocateDown - CRASHES WORD 2013
    Application.Run "OutlineMoveDown"
    Selection.Collapse

    'Switch back to previous view
    ActiveWindow.ActivePane.View.Type = CurrentView
    
    Application.ScreenUpdating = True
    
    Exit Sub
    
Handler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Sub DeleteHeading()
'Deletes the current card, block, hat, or pocket

    Call Paperless.SelectHeadingAndContent
    Selection.Delete

End Sub

'*************************************************************************************
'* SEND FUNCTIONS                                                                                               *
'*************************************************************************************

Sub SendToSpeech()

'Sends content to the Speech doc.  Sends currently selected text,
'or if nothing is selected, the current tag, block, hat, or pocket
'Speech marker doesn't work on Mac Word reading view - code left here in case of new version

    Dim CurrentDoc As String
    Dim SpeechDoc As Document
    Dim d As Document
    Dim FoundDoc As Long

    On Error GoTo Handler
    
    'If in reading mode, enter a stopped reading marker
    'If ActiveWindow.View.FullScreen Then
    '        Selection.Collapse
    '        If Selection.Words(1).End <> Selection.End Then Selection.MoveRight wdWord
    '        Selection.Font.Color = wdColorRed
    '        Selection.Font.Size = 18
    '        Selection.TypeText Chr(167) & " Marked " & FormatDateTime(Time, 4) & " " & Chr(167) & " "
    '        Exit Sub
    '    End If
   ' End If

    'Save active document name
    CurrentDoc = ActiveDocument.Name

SpeechDocCheck:

    'If there's an active speech doc, use it
    If ActiveSpeechDoc <> "" Then
        For Each d In Application.Documents
            If d.Name = ActiveSpeechDoc Then
                Set SpeechDoc = Application.Documents(ActiveSpeechDoc)
            End If
        Next d
    Else
        'Look for a document with "speech" in the title
        For Each d In Application.Documents
            If InStr(LCase(d.Name), "speech") Then
                FoundDoc = FoundDoc + 1
                If FoundDoc = 1 Then Set SpeechDoc = d
            End If
        Next d
        
        'If no Speech doc is found, prompt whether to create one.
        'If yes, create a new document based on the current template to save, then retry
        If FoundDoc = 0 Then
            If MsgBox("Speech document is not open - create one?", vbYesNo, "Create Speech?") = vbNo Then
                Exit Sub
            Else
                'Create New Speech Doc
                Call Paperless.NewSpeech
            
                'Switch focus back after save
                Documents(CurrentDoc).Activate
                GoTo SpeechDocCheck:
            End If
        End If
    
        'If multiple Speech docs are open, warn the user.
        If FoundDoc > 1 Then
            Call Paperless.ShowChooseSpeechDoc
            Exit Sub
        End If
    End If
    
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
    If Selection.Paragraphs(1).outlineLevel = 4 Then
        If SpeechDoc.ActiveWindow.Selection.Paragraphs.outlineLevel < wdOutlineLevel4 Then
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

Handler:
    Application.ScreenUpdating = True
    Set SpeechDoc = Nothing
    
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

'*************************************************************************************
'* DOCUMENT FUNCTIONS                                                                                     *
'*************************************************************************************

Sub NewDocument()
'Adds a new document based on the debate template
    
    #If MAC_OFFICE_VERSION >= 15 Then
        Application.Documents.Add Template:=Application.NormalTemplate.Path & "/Debate.dotm"
    #Else
        Application.Documents.Add Template:=Application.NormalTemplate.Path & ":My Templates:Debate.dotm"
    #End If
    
End Sub

Sub NewSpeech()
'Creates a new Speech document
    
    Dim SpeechName As String
    Dim FileName As String
    Dim h
    Dim AutoSaveDirectory As String
 
    On Error GoTo Handler
    
SpeechName:
    'Get input for which Speech to name it
    SpeechName = InputBox("Which Speech (1NC, 2AC, etc...)? You can also add extra info about the round.", "New Speech", "e.g. 2AC Round 3 vs Hogwarts")
    If SpeechName = "" Then Exit Sub
    If SpeechName = "e.g. 2AC Round 3 vs Hogwarts" Then GoTo SpeechName
    SpeechName = Trim(ScrubString(SpeechName))
    SpeechName = Replace(SpeechName, "/", "")
    SpeechName = Replace(SpeechName, "\", "")
    SpeechName = Replace(SpeechName, ":", "")
    
    'Create filename
    If Hour(Now) > 12 Then h = Hour(Now) - 12 & "PM"
    If Hour(Now) <= 12 Then h = Hour(Now) & "AM"
    FileName = "Speech " & SpeechName & " " & Month(Now) & "-" & Day(Now) & " " & h

    'Add new document based on template
    Call Paperless.NewDocument
 
    'If AutoSave is set, save the doc - otherwise bring up Save As dialogue with default name set
    If Application.Version < "15" And GetSetting("Verbatim", "Paperless", "AutoSaveSpeech", False) = True Then
        AutoSaveDirectory = Trim(GetSetting("Verbatim", "Paperless", "AutoSaveDir", CurDir()))
        If Right(AutoSaveDirectory, 1) <> ":" Then AutoSaveDirectory = AutoSaveDirectory & ":"
        FileName = AutoSaveDirectory & FileName
        ActiveDocument.SaveAs FileName:=FileName, FileFormat:=wdFormatXMLDocument
    Else
        With Application.Dialogs(wdDialogFileSaveAs)
            .Name = FileName
            If .Show = 0 Then Exit Sub
        End With
    End If
    
    Exit Sub
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

'*************************************************************************************
'* TOOL FUNCTIONS                                                                                               *
'*************************************************************************************

Sub CopyToUSB()
'Copies the current file to the root folder of the first found USB drive
    
    Dim POSIXActive
    Dim FileName As String
    
    Dim MountPoints As String
    Dim MountPointArray
    Dim m

    On Error GoTo Handler
    
    'Get full POSIX path of current file
    POSIXActive = MacScript("return POSIX path of """ & ActiveDocument.FullName & """")
    
    'Get list of mounted USB drives - throws an error if none plugged in, so turn off error checking temporarily
    On Error Resume Next
    #If MAC_OFFICE_VERSION >= 15 Then
        MountPoints = AppleScriptTask("Verbatim.scpt", "RunShellScript", "system_profiler SPUSBDataType | grep 'Mount Point'")
    #Else
        MountPoints = MacScript("do shell script ""system_profiler SPUSBDataType | grep 'Mount Point'""")
    #End If
    On Error GoTo Handler
    
    'Exit if no USB drives found
    If MountPoints = "" Then
        MsgBox "No USB drives found!"
        Exit Sub
    End If
    
    'Split into array and loop each drive
    MountPointArray = Split(MountPoints, Chr(13))
    For Each m In MountPointArray
        m = Trim(Replace(m, "Mount Point: ", "")) & "/" 'Get just the mount path and add a trailing /
        
        'Strip "Speech" if option set
        If GetSetting("Verbatim", "Paperless", "StripSpeech", True) = True And Len(ActiveDocument.Name) > 11 Then
            FileName = Trim(Replace(ActiveDocument.Name, "speech", "", 1, -1, vbTextCompare))
        Else
            FileName = ActiveDocument.Name
        End If
        
        'Check if file already exists on USB
        #If MAC_OFFICE_VERSION >= 15 Then
        If AppleScriptTask("Verbatim.scpt", "RunShellScript", "test -e '" & m & FileName & "'; echo $?") = "0" Then
        #Else
        If MacScript("do shell script ""test -e '" & m & FileName & "'; echo $?""") = "0" Then
        #End If
            If MsgBox("File Exists.  Overwrite?", vbOKCancel) = vbCancel Then Exit Sub
        End If
        
        'Save File locally
        ActiveDocument.Save
        
        'Copy To USB
        #If MAC_OFFICE_VERSION >= 15 Then
            AppleScriptTask "Verbatim.scpt", "RunShellScript", "cp '" & POSIXActive & "' '" & m & FileName & "'"
        #Else
            MacScript ("do shell script ""cp '" & POSIXActive & "' '" & m & FileName & "'""")
        #End If
        MsgBox "Sucessfully copied to USB!"
        
    Next m
    
    Exit Sub
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Sub StartTimer()
'Starts a user supplied timer.
    
    Dim TimerApp As String
    Dim TimerAppPOSIX As String
    
    On Error GoTo Handler
    
    'Get path to timer app
    TimerApp = GetSetting("Verbatim", "Paperless", "TimerApp", "?")
    
    'If not set, try default
    If TimerApp = "?" Then TimerApp = MacScript("return path to applications folder as string") & "Debate Timer for Mac.app"

    'Make sure timer app exists
    #If MAC_OFFICE_VERSION >= 15 Then
    If AppleScriptTask("Verbatim.scpt", "FileExists", TimerApp) = "false" Then
    #Else
    If MacScript("tell application ""Finder""" & Chr(13) & "exists file """ & TimerApp & """" & Chr(13) & "end tell") = "false" Then
    #End If
        MsgBox "Timer application not found. Ensure you have one installed and enter the correct path to the application in the Verbatim Settings." & vbCrLf & vbCrLf & "See the Verbatim manual on paperlessdebate.com for suggestions of Mac timer programs."
        Exit Sub
    Else
        'If java app selected, run it from the shell
        If Right(TimerApp, 5) = ".jar:" Or Right(TimerApp, 4) = ".jar" Then
            TimerAppPOSIX = MacScript("return POSIX path of """ & TimerApp & """")
            #If MAC_OFFICE_VERSION >= 15 Then
                AppleScriptTask "Verbatim.scpt", "RunShellScript", "open '" & TimerAppPOSIX & "'"
            #Else
                MacScript ("do shell script ""open '" & TimerAppPOSIX & "'""")
            #End If
        Else
            #If MAC_OFFICE_VERSION >= 15 Then
                AppleScriptTask "Verbatim.scpt", "ActivateTimer", TimerApp
            #Else
                MacScript ("tell application """ & TimerApp & """ to activate")
            #End If
        End If
    End If

    Exit Sub
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

'*************************************************************************************
'* WARRANT FUNCTIONS                                                                                        *
'*************************************************************************************

Sub NewWarrant()
    Selection.Comments.Add Range:=Selection.Range
End Sub

Sub DeleteAllWarrants()
    Dim c As Comment
    For Each c In ActiveDocument.Comments
        c.Delete
    Next c
End Sub
