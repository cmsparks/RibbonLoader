Attribute VB_Name = "frmCaselist"
Attribute VB_Base = "0{2DA71CCE-7635-4AAA-9C39-A175AFFC1DA4}{A5D89ECF-8FD6-46BB-9E76-DD9004030B27}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub UserForm_Initialize()
    
    'Add Round and side options
    Me.cboRound.AddItem "Round 1"
    Me.cboRound.AddItem "Round 2"
    Me.cboRound.AddItem "Round 3"
    Me.cboRound.AddItem "Round 4"
    Me.cboRound.AddItem "Round 5"
    Me.cboRound.AddItem "Round 6"
    Me.cboRound.AddItem "Round 7"
    Me.cboRound.AddItem "Round 8"
    Me.cboRound.AddItem "Round 9"
    Me.cboRound.AddItem "Quads"
    Me.cboRound.AddItem "Triples"
    Me.cboRound.AddItem "Doubles"
    Me.cboRound.AddItem "Octas"
    Me.cboRound.AddItem "Quarters"
    Me.cboRound.AddItem "Semis"
    Me.cboRound.AddItem "Finals"
    
    Me.cboSide.AddItem "Aff"
    Me.cboSide.AddItem "Neg"
    
    'Initialize entry counter
    Me.txtEntryCount.Value = 0
    
End Sub

Private Sub UserForm_Activate()
       
    Dim RoundArray As Variant
    Dim i
    
    'Turn on error checking
    On Error GoTo Handler
    
    'Make sure tabroom username/password are filled out in the settings
    If GetSetting("Verbatim", "Main", "TabroomUsername", "?") = "?" Or GetSetting("Verbatim", "Main", "TabroomPassword", "?") = "?" Then
        If MsgBox("You must enter a tabroom.com account in the Verbatim Settings to upload to the caselist. Open Settings Now?", vbYesNo) = vbYes Then
            Me.Hide
            Settings.ShowSettingsForm
            Me.Show
        Else
            Unload Me
        End If
    End If
    
    'Make sure default school/team are filled out in the settings
    If GetSetting("Verbatim", "Caselist", "CaselistSchoolName", "?") = "?" Or GetSetting("Verbatim", "Caselist", "CaselistTeamName", "?") = "?" Then
        If MsgBox("You must set a default school/team to upload to the caselist. Open Settings Now?", vbYesNo) = vbYes Then
            Me.Hide
            Settings.ShowSettingsForm
            Me.Show
        Else
            Unload Me
        End If
    End If
    
    'Reset Select Round box and add a blank item
    Me.cboSelectRound.Clear
    Me.cboSelectRound.AddItem
    Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 0) = ""
    
    'Get rounds from Tabroom
    RoundArray = Tabroom.GetTabroomRounds()
    
    'Loop Rounds and save Round info for later retrieval
    If IsArray(RoundArray) Then
        For i = 0 To UBound(RoundArray, 1)
            
            Me.cboSelectRound.AddItem
            Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 0) = RoundArray(i, 0) & " " & RoundArray(i, 1) & " " & RoundArray(i, 2) & " vs " & RoundArray(i, 3)
            Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 1) = RoundArray(i, 0) 'Tournament
            Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 2) = RoundArray(i, 1) 'RoundName
            Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 3) = RoundArray(i, 2) 'Side
            Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 4) = RoundArray(i, 3) 'Opponent
            Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 5) = RoundArray(i, 4) 'Judge
        Next
    End If
    
    'Show caselist settings
    Select Case GetSetting("Verbatim", "Caselist", "DefaultWiki", "openCaselist")
        Case Is = "openCaselist"
            Me.lblSettingsWiki.Caption = "Wiki: openCaselist"
        Case Is = "NDCAPolicy"
            Me.lblSettingsWiki.Caption = "Wiki: NDCA Policy"
        Case Is = "NDCALD"
            Me.lblSettingsWiki.Caption = "Wiki: NDCA LD"
        Case Else
            Me.lblSettingsWiki.Caption = "Wiki: openCaselist"
    End Select
    
    Me.lblSettingsSchool.Caption = "School: " & GetSetting("Verbatim", "Caselist", "CaselistSchoolName", "?")
    Me.lblSettingsTeam.Caption = "Team: " & GetSetting("Verbatim", "Caselist", "CaselistTeamName", "?")
        
    'Process doc to get Pockets for automatic cite entry maker
    Call RefreshPockets
  
    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub mpgCaselist_Change()

    Select Case Me.mpgCaselist.Value
        Case Is = 0 'Round Info tab
            Me.btnBack.Visible = False
        Case Is = 1 'Round Report tab
            Me.btnBack.Visible = True
        Case Is = 2 'Cites tab
            'Do nothing
        Case Is = 3 'Open Source tab
            Me.btnNext.Caption = "Next-->"
        Case Is = 4 'Submit tab
            Me.btnNext.Caption = "Upload"
            
            'Recompute submission info
            Me.lblSubmitTournament.Caption = "Tournament: " & Me.txtTournament.Value
            Me.lblSubmitRound.Caption = "Round: " & Me.cboRound.Value
            Me.lblSubmitSide.Caption = "Side: " & Me.cboSide.Value
            Me.lblSubmitOpponent.Caption = "Opponent: " & Me.txtOpponent.Value
            Me.lblSubmitJudge.Caption = "Judge(s): " & Me.txtJudge.Value
            
            'Display images for selected options
            If Me.txtRoundReport.Value <> "" Then
                Me.imgYesRoundReport.Visible = True
                Me.imgNoRoundReport.Visible = False
            Else
                Me.imgYesRoundReport.Visible = False
                Me.imgNoRoundReport.Visible = True
            End If
            
            If Me.txtEntryCount.Value > 0 Then
                Me.imgYesCites.Visible = True
                Me.imgNoCites.Visible = False
            Else
                Me.imgYesCites.Visible = False
                Me.imgNoCites.Visible = True
            End If
            
            If Me.chkOpenSource.Value = True Then
                Me.imgYesOpenSource.Visible = True
                Me.imgNoOpenSource.Visible = False
            Else
                Me.imgYesOpenSource.Visible = False
                Me.imgNoOpenSource.Visible = True
            End If
            
        Case Else
            'Do nothing
    End Select
    
End Sub

Private Sub btnNext_Click()

    Dim Caselist As String
    Dim School As String
    Dim Team As String

    Dim CaselistStem As String
    Dim CaselistURL As String
    Dim RoundObjectNumber As String
    Dim CiteObjectNumbers As String
    Dim CaselistFileName As String
    Dim OSURL As String
    
    Dim i
    Dim c As control

    'Turn on error checking
    On Error GoTo Handler
    
    'Validate Round Info
    If Me.mpgCaselist.Value = 0 Then
        
        Me.txtTournament.Value = Trim(ScrubString(Me.txtTournament.Value))
        If Me.txtTournament.Value = "" Then
            MsgBox "You must enter a value for Tournament."
            Exit Sub
        End If
        
        Me.cboRound.Value = Trim(ScrubString(Me.cboRound.Value))
        If Me.cboRound.Value = "" Then
            MsgBox "You must enter a value for Round."
            Exit Sub
        End If
        
        If Me.cboSide.Value = "" Then
            MsgBox "You must enter a value for Side."
            Exit Sub
        End If
        
        If Me.cboSide.Value <> "Aff" And Me.cboSide.Value <> "Neg" Then
            MsgBox "Side must be ""Aff"" or ""Neg."""
            Exit Sub
        End If

        Me.txtOpponent.Value = Trim(ScrubString(Me.txtOpponent.Value))
        If Me.txtOpponent.Value = "" Then
            MsgBox "You must enter a value for Opponent."
            Exit Sub
        End If

        Me.txtJudge.Value = Trim(ScrubString(Me.txtJudge.Value))
    End If
    
    'Validate Round Report
    If Me.mpgCaselist.Value = 1 Then
        Me.txtRoundReport.Value = Trim(ScrubString(Me.txtRoundReport.Value))
    End If
    
    'Validate Caselist
    If Me.mpgCaselist.Value = 2 Then
        If Me.txtEntryCount.Value = 0 Then 'Nag if no entries
            If MsgBox("Are you sure you want to skip creating cite entries? It's easy, and open source is NOT a replacement for good disclosure practices.", vbYesNo) = vbNo Then Exit Sub
        Else
            
            For i = 1 To Me.txtEntryCount.Value
                
                'Validate entry titles
                Set c = Me.fEntries.Controls("cboPrefix" & i)
                c.Value = Trim(ScrubString(c.Value))
                
                Set c = Me.fEntries.Controls("txtEntryTitle" & i)
                c.Value = Trim(ScrubString(c.Value))
                If c.Value = "" Then
                    MsgBox "You must include a title for Entry " & i
                    Exit Sub
                End If
                
                'Make Entry text XML safe - don't need to escape ' and " because they're not in attribute text
                Set c = Me.fEntries.Controls("txtEntryText" & i)
                c.Value = Replace(c.Value, "&", "&amp;")
                c.Value = Replace(c.Value, "<", "&lt;")
                c.Value = Replace(c.Value, ">", "&gt;")
                
            Next i
        End If
    End If
    
    'Move tabs
    If Me.mpgCaselist.Value < 4 Then
        Me.mpgCaselist(Me.mpgCaselist.Value + 1).Enabled = True
        Me.mpgCaselist.Value = Me.mpgCaselist.Value + 1
        Me.mpgCaselist(Me.mpgCaselist.Value - 1).Enabled = False
    
    Else 'Upload button
        
        'Validate either cites or open source selected
        If Me.txtEntryCount.Value = 0 And Me.chkOpenSource.Value = False Then
            MsgBox "Nothing to upload - you must include either cite entries or select open source."
            Exit Sub
        End If
        
        'Use default caselist settings unless temporary box checked
        If Me.chkTemporaryPage.Value = False Then
            Caselist = GetSetting("Verbatim", "Caselist", "DefaultWiki", "openCaselist")
            School = GetSetting("Verbatim", "Caselist", "CaselistSchoolName", "?")
            Team = GetSetting("Verbatim", "Caselist", "CaselistTeamName", "?")
        Else
            'Validate school and team selected
            If Me.cboCaselistSchoolName.Value = "" Or Me.cboCaselistSchoolName.Value = "?" Then
                MsgBox "You must select a school name."
                Exit Sub
            End If
            
            If Me.cboCaselistTeamName.Value = "" Or Me.cboCaselistTeamName.Value = "?" Then
                MsgBox "You must select a team name."
                Exit Sub
            End If
            
            If Me.optOpenCaselist.Value = True Then Caselist = "openCaselist"
            If Me.optNDCAPolicy.Value = True Then Caselist = "NDCAPolicy"
            If Me.optNDCALD.Value = True Then Caselist = "NDCALD"
            School = Me.cboCaselistSchoolName.Value
            Team = Me.cboCaselistTeamName.Value
            
        End If
        
        'Initialize progress label - cover everything else
        Me.lblProgress.Visible = True
        Me.lblProgress.ZOrder (0)
        Me.lblProgress.Top = 0
        Me.lblProgress.Left = 6
        Me.lblProgress.Height = 370
        Me.lblProgress.Width = 285
        Me.lblProgress.ForeColor = vbBlack
        Me.lblProgress.Caption = vbCrLf & "Computing upload location....."
        Me.Repaint 'Update form
        
        'Get URL for appropriate caselist
        Select Case Caselist
            Case Is = "openCaselist"
                CaselistStem = GetCaselistURL("openCaselist")
            Case Is = "NDCAPolicy"
                CaselistStem = GetCaselistURL("NDCAPolicy")
            Case Is = "NDCALD"
                CaselistStem = GetCaselistURL("NDCALD")
            Case Else
                CaselistStem = GetCaselistURL("openCaselist")
        End Select
    
        'Exit if error
        If CaselistStem = "HTTP Error" Then
            MsgBox "Internet error. Please try again later."
            Me.lblProgress.Visible = False 'Hide progress bar
            Exit Sub
        End If
        
        'Construct Caselist URL for uploading objects
        CaselistURL = CaselistStem & School & "/pages/" & Team & " " & Me.cboSide.Value & "/objects"
               
        Me.lblProgress.Caption = Me.lblProgress.Caption & "done." & vbCrLf
        Me.lblProgress.Caption = Me.lblProgress.Caption & "Creating Round....."
        Me.Repaint

        'Create Round, get round number
        RoundObjectNumber = CaselistCreateRound(CaselistURL)
        
        If RoundObjectNumber = "" Then
            MsgBox "Failed to create round - check your round info."
            Me.lblProgress.Visible = False 'Hide progress bar
            Exit Sub
        End If
        
        'Upload cite entries linked to the round, get cite object numbers - pass in school and team in case using temporary page
        If Me.txtEntryCount.Value > 0 Then
            Me.lblProgress.Caption = Me.lblProgress.Caption & "Creating Cite Entries....."
            Me.Repaint
            CiteObjectNumbers = CaselistCreateCiteEntries(CaselistURL, RoundObjectNumber, School, Team)
            
            If CiteObjectNumbers = "" Then
                MsgBox "Failed to create cite entries - check your cite info."
                Me.lblProgress.Visible = False 'Hide progress bar
                Exit Sub
            End If
        End If
        
        'Upload OS if checked
        If Me.chkOpenSource.Value = True Then
            Me.lblProgress.Caption = Me.lblProgress.Caption & "Uploading Open Source....."
            Me.Repaint
            
            'Construct file name and URL
            CaselistFileName = School & "-" & Replace(Team, " ", "-") & "-" & Me.cboSide.Value & "-" & Me.txtTournament.Value & "-" & Me.cboRound.Value & Right(ActiveDocument.FullName, Len(ActiveDocument.FullName) - InStrRev(ActiveDocument.FullName, ".") + 1)
            CaselistURL = CaselistStem & School & "/pages/" & Team & " " & Me.cboSide.Value & "/attachments/" & CaselistFileName
            
            'Upload file and get attachment URL
            OSURL = UploadOpenSource(CaselistURL)
            
            If OSURL = "" Then
                MsgBox "Failed to upload open source."
                Me.lblProgress.Visible = False 'Hide progress bar
                Exit Sub
            End If
        End If
        
        'Update Round object with entry object numbers and attachment URL
        Me.lblProgress.Caption = Me.lblProgress.Caption & "Updating Round....."
        Me.Repaint
        CaselistURL = CaselistStem & School & "/pages/" & Team & " " & Me.cboSide.Value & "/objects/Caselist.RoundClass/" & RoundObjectNumber & "/properties"
        Call CaselistUpdateRound(CaselistURL, RoundObjectNumber, CiteObjectNumbers, OSURL)
        
        'Close form
        Unload Me
    
    End If
    
    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    
End Sub

Private Sub btnBack_Click()
    Me.mpgCaselist(Me.mpgCaselist.Value - 1).Enabled = True
    Me.mpgCaselist.Value = Me.mpgCaselist.Value - 1
    Me.mpgCaselist(Me.mpgCaselist.Value + 1).Enabled = False
End Sub

'*************************************************************************************
'* ROUND INFO TAB                                                                                                *
'*************************************************************************************

Private Sub cboSelectRound_Change()
   
    On Error GoTo Handler
    
    'If list is empty, exit
    If Me.cboSelectRound.ListCount = 0 Then Exit Sub
    
    'If selected item is the first blank line, clear boxes
    If Me.cboSelectRound.ListIndex = 0 Then
        Me.txtTournament.Value = ""
        Me.cboRound.Value = ""
        Me.cboSide.Value = ""
        Me.txtOpponent.Value = ""
        Me.txtJudge.Value = ""
        
    'Tabroom round is selected - fill out boxes
    Else
        Me.txtTournament.Value = Trim(Me.cboSelectRound.List(Me.cboSelectRound.ListIndex, 1))
        Me.cboRound.Value = Trim(Me.cboSelectRound.List(Me.cboSelectRound.ListIndex, 2))
        Me.cboSide.Value = Trim(Me.cboSelectRound.List(Me.cboSelectRound.ListIndex, 3))
        Me.txtOpponent.Value = Trim(Me.cboSelectRound.List(Me.cboSelectRound.ListIndex, 4))
        Me.txtJudge.Value = Trim(Me.cboSelectRound.List(Me.cboSelectRound.ListIndex, 5))
        
    End If

    Exit Sub
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    
End Sub

'*************************************************************************************
'* CITES TAB                                                                                                            *
'*************************************************************************************

Private Sub btnRefreshPockets_Click()
    Call RefreshPockets
End Sub

Private Sub RefreshPockets()

    Dim txt
    Dim p
    
    'Clear list box
    Me.lboxPockets.Clear
    
    'Parse doc, add to list box
    For Each p In ActiveDocument.Paragraphs
        If p.outlineLevel = 1 Then
            txt = Trim(p.Range.Text)
            txt = Replace(txt, vbCrLf, "")
            txt = Replace(txt, Chr(10), "")
            txt = Replace(txt, Chr(13), "")
            txt = Replace(txt, Chr(12), "")
            Me.lboxPockets.AddItem
            Me.lboxPockets.List(Me.lboxPockets.ListCount - 1, 0) = txt
            Me.lboxPockets.List(Me.lboxPockets.ListCount - 1, 1) = p.Range.Start
        End If
    Next p

End Sub

Private Sub btnSelectAllPockets_Click()

    Dim i
    
    For i = 0 To Me.lboxPockets.ListCount - 1
        Me.lboxPockets.Selected(i) = True
    Next
        
End Sub

Private Sub btnSelectNoPockets_Click()

    Dim i
    
    For i = 0 To Me.lboxPockets.ListCount - 1
        Me.lboxPockets.Selected(i) = False
    Next
        
End Sub

Private Sub btnPockets2Entries_Click()
    
    Dim i
    
    'Loop all selected items in box and create a cite entry, then remove from list
    For i = Me.lboxPockets.ListCount - 1 To 0 Step -1
        If Me.lboxPockets.Selected(i) = True Then
            Call AddCiteEntry
            Selection.Start = lboxPockets.List(i, 1)
            Call Paperless.SelectHeadingAndContent
            Call WikifySelection
            Me.lboxPockets.RemoveItem (i)
        End If
    Next i
    
End Sub

Private Sub txtEntryCount_Change()

    'Hide instructions when entry created, or make visible when all entries deleted
    If Me.txtEntryCount.Value > 0 Then
        Me.lblEntriesInstructions.Visible = False
    Else
        Me.lblEntriesInstructions.Visible = True
    End If

End Sub

Private Sub AddCiteEntry()
        
    Dim EntryLabel As control
    Dim PrefixLabel As control
    Dim PrefixBox As control
    Dim TitleLabel As control
    Dim TitleBox As control
    Dim EntryText As control
    
    Dim Prefixes
    Dim p
    
    'Increment the entry counter - also makes instructions invisible
    Me.txtEntryCount.Value = Me.txtEntryCount.Value + 1
    
    'Create Entry Label - all other positioning keyed off this
    Set EntryLabel = Me.fEntries.Controls.Add("Forms.Label.1", "lblEntry" & Me.txtEntryCount.Value)
    EntryLabel.Caption = "Entry " & Me.txtEntryCount.Value & ":"
    EntryLabel.Height = 12
    EntryLabel.Width = 48
    EntryLabel.Left = 5
    EntryLabel.Top = Me.fEntries.ScrollHeight + 10
    
    'Create Prefix Label
    Set PrefixLabel = Me.fEntries.Controls.Add("Forms.Label.1", "lblPrefix" & Me.txtEntryCount.Value)
    PrefixLabel.Caption = "Prefix (optional)"
    PrefixLabel.Height = 12
    PrefixLabel.Width = 65
    PrefixLabel.Left = 5
    PrefixLabel.Top = EntryLabel.Top + EntryLabel.Height
    
    'Create Prefix Box
    Set PrefixBox = Me.fEntries.Controls.Add("Forms.ComboBox.1", "cboPrefix" & Me.txtEntryCount.Value)
    PrefixBox.Height = 18
    PrefixBox.Width = 65
    PrefixBox.Left = 5
    PrefixBox.Top = PrefixLabel.Top + PrefixLabel.Height
    
    'Add custom prefixes
    If GetSetting("Verbatim", "Caselist", "CustomPrefixes", "?") <> "?" Then
        Prefixes = Split(GetSetting("Verbatim", "Caselist", "CustomPrefixes", "?"), ",")
        For Each p In Prefixes
            PrefixBox.AddItem Trim(p)
        Next p
    End If
    
    'Add built-in prefixes
    PrefixBox.AddItem "1AC -"
    PrefixBox.AddItem "2AC -"
    PrefixBox.AddItem "Adv -"
    PrefixBox.AddItem "DA -"
    PrefixBox.AddItem "CP -"
    PrefixBox.AddItem "K -"
    PrefixBox.AddItem "T -"
        
    'Create Title Label
    Set TitleLabel = Me.fEntries.Controls.Add("Forms.Label.1", "lblEntryTitle" & Me.txtEntryCount.Value)
    TitleLabel.Caption = "Entry Title"
    TitleLabel.Height = 12
    TitleLabel.Width = 65
    TitleLabel.Left = PrefixLabel.Left + PrefixLabel.Width + 10
    TitleLabel.Top = PrefixLabel.Top
    
    'Create Title Box
    Set TitleBox = Me.fEntries.Controls.Add("Forms.TextBox.1", "txtEntryTitle" & Me.txtEntryCount.Value)
    TitleBox.Height = 18
    TitleBox.Width = 170
    TitleBox.Left = PrefixBox.Left + PrefixBox.Width + 10
    TitleBox.Top = PrefixBox.Top
    
    'Create Entry Box
    Set EntryText = Me.fEntries.Controls.Add("Forms.TextBox.1", "txtEntryText" & Me.txtEntryCount.Value)
    EntryText.Height = 100
    EntryText.Width = PrefixBox.Width + 10 + TitleBox.Width
    EntryText.Left = PrefixBox.Left
    EntryText.Top = PrefixBox.Top + PrefixBox.Height + 5
    EntryText.MultiLine = True
    EntryText.EnterKeyBehavior = True
    EntryText.ScrollBars = 2
    EntryText.Font.Size = 8
    
    'Add ScrollHeight and scroll to bottom
    Me.fEntries.ScrollHeight = Me.fEntries.ScrollHeight + 160
    Me.fEntries.ScrollTop = Me.fEntries.ScrollHeight
End Sub

Private Sub btnDeleteCiteEntry_Click()

    If Me.txtEntryCount.Value > 0 Then
    
        'Delete last entry
        Me.fEntries.Controls.Remove ("lblEntry" & Me.txtEntryCount.Value)
        Me.fEntries.Controls.Remove ("lblPrefix" & Me.txtEntryCount.Value)
        Me.fEntries.Controls.Remove ("cboPrefix" & Me.txtEntryCount.Value)
        Me.fEntries.Controls.Remove ("lblEntryTitle" & Me.txtEntryCount.Value)
        Me.fEntries.Controls.Remove ("txtEntryTitle" & Me.txtEntryCount.Value)
        Me.fEntries.Controls.Remove ("txtEntryText" & Me.txtEntryCount.Value)
        
        'Decrement counter
        Me.txtEntryCount.Value = Me.txtEntryCount.Value - 1
    
        'Remove excess ScrollHeight
        Me.fEntries.ScrollHeight = Me.fEntries.ScrollHeight - 160
    End If
    
End Sub

Private Sub WikifySelection()

    Dim p
    
    On Error GoTo Handler
    
    'Turn off screen updating
    Application.ScreenUpdating = False
    
    'Set entry title to text of first header in selection
    Me.Controls("txtEntryTitle" + Me.txtEntryCount.Value).Value = Left(Selection.Paragraphs(1).Range.Text, Len(Selection.Paragraphs(1).Range.Text) - 1)
    
    'Copy selection
    Selection.Copy
    
    'Add new document based on debate template
    Application.Documents.Add Template:=ActiveDocument.AttachedTemplate.FullName
    
    'Paste into new document
    Selection.Paste
    
    'Go to top of document and collapse selection
    Selection.HomeKey Unit:=wdStory
    Selection.Collapse

    'Convert to cites
    Call Caselist.CiteRequestAll

    'Wikify and clear formatting
    Call Caselist.Word2XWikiMain
    ActiveDocument.Content.Select
    Selection.ClearFormatting
    
    'Set EntryText
    Me.Controls("txtEntryText" + Me.txtEntryCount.Value).Value = Selection.Text
    
    'Close temporary doc without saving
    ActiveDocument.Close wdDoNotSaveChanges
    
    'Turn on screen updating
    Application.ScreenUpdating = True

    Exit Sub
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    
End Sub

'*************************************************************************************
'* SUBMIT TAB                                                                                                         *
'*************************************************************************************

Private Sub btnSettings_Click()
    Me.Hide
    Settings.ShowSettingsForm
    Me.Show
End Sub

Private Sub chkTemporaryPage_Click()

    If Me.chkTemporaryPage.Value = True Then
        Me.fChooseWiki.Visible = True
        Me.lblCaselistSchoolName.Visible = True
        Me.cboCaselistSchoolName.Visible = True
        Me.lblCaselistTeamName.Visible = True
        Me.cboCaselistTeamName.Visible = True
        
        Me.optOpenCaselist.Value = True
    Else
        Me.fChooseWiki.Visible = False
        Me.lblCaselistSchoolName.Visible = False
        Me.cboCaselistSchoolName.Visible = False
        Me.lblCaselistTeamName.Visible = False
        Me.cboCaselistTeamName.Visible = False
    End If

End Sub

Private Sub optOpenCaselist_Change()
    Me.cboCaselistSchoolName.Value = ""
    Me.cboCaselistSchoolName.Clear
    Me.cboCaselistTeamName.Value = ""
    Me.cboCaselistTeamName.Clear
End Sub

Private Sub optNDCAPolicy_Change()
    Me.cboCaselistSchoolName.Value = ""
    Me.cboCaselistSchoolName.Clear
    Me.cboCaselistTeamName.Value = ""
    Me.cboCaselistTeamName.Clear
End Sub

Private Sub optNDCALD_Change()
    Me.cboCaselistSchoolName.Value = ""
    Me.cboCaselistSchoolName.Clear
    Me.cboCaselistTeamName.Value = ""
    Me.cboCaselistTeamName.Clear
End Sub

Private Sub cboCaselistSchoolName_Change()
    Me.cboCaselistTeamName.Value = ""
    Me.cboCaselistTeamName.Clear
End Sub

Private Sub cboCaselistSchoolName_DropButtonClick()
'Populates the SchoolName combo box with schools from the caselist
        
    'If the list is already populated, exit
    If Me.cboCaselistSchoolName.ListCount > 0 Then Exit Sub
        
    'Clear ComboBoxes - clear TeamName too, so there's not a mismatch when changing
    Me.cboCaselistSchoolName.Value = ""
    Me.cboCaselistTeamName.Value = ""
    Me.cboCaselistSchoolName.Clear
    Me.cboCaselistTeamName.Clear
        
    'Populate box
    If Me.optOpenCaselist.Value = True Then Call Caselist.GetCaselistSchoolNames("openCaselist", Me.cboCaselistSchoolName)
    If Me.optNDCAPolicy.Value = True Then Call Caselist.GetCaselistSchoolNames("NDCAPolicy", Me.cboCaselistSchoolName)
    If Me.optNDCALD.Value = True Then Call Caselist.GetCaselistSchoolNames("NDCALD", Me.cboCaselistSchoolName)
    
    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Private Sub cboCaselistTeamName_DropButtonClick()
'Populates the TeamName combo box with teams from the school's space

    'If the list is already populated, exit
    If Me.cboCaselistTeamName.ListCount > 0 Then Exit Sub
    
    'Check CaselistSchoolName has a value
    If Me.cboCaselistSchoolName.Value = "" Then
        Me.cboCaselistTeamName.Value = "Please choose a school first"
        Me.cboCaselistTeamName.Clear
        Exit Sub
    End If
    
    'Clear ComboBox
    Me.cboCaselistTeamName.Value = ""
    Me.cboCaselistTeamName.Clear
  
    If Me.optOpenCaselist.Value = True Then Call Caselist.GetCaselistTeamNames("openCaselist", Me.cboCaselistSchoolName.Value, Me.cboCaselistTeamName)
    If Me.optNDCAPolicy.Value = True Then Call Caselist.GetCaselistTeamNames("NDCAPolicy", Me.cboCaselistSchoolName.Value, Me.cboCaselistTeamName)
    If Me.optNDCALD.Value = True Then Call Caselist.GetCaselistTeamNames("NDCALD", Me.cboCaselistSchoolName.Value, Me.cboCaselistTeamName)
    
    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

'*************************************************************************************
'* UPLOAD FUNCTIONS                                                                                           *
'*************************************************************************************

Private Function CaselistCreateRound(CaselistURL As String) As String
    
    Dim Script As String
    Dim XML As String
    
    Dim TempFile
    Dim TempFilePath As String
    Dim TempFilePOSIX As String
    
    'Turn on error checking
    On Error GoTo Handler
    
    'Create XML
    XML = XML & "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
    XML = XML & "<object xmlns=""http://www.xwiki.org"">"
    XML = XML & "<className>Caselist.RoundClass</className>"
    XML = XML & "<property name=""Tournament""><value>" & Me.txtTournament.Value & "</value></property>"
    XML = XML & "<property name=""Round""><value>" & Me.cboRound.Value & "</value></property>"
    XML = XML & "<property name=""Opponent""><value>" & Me.txtOpponent.Value & "</value></property>"
    XML = XML & "<property name=""Judge""><value>" & Me.txtJudge.Value & "</value></property>"
    If Me.txtRoundReport.Value <> "" Then XML = XML & "<property name=""RoundReport""><value>" & Me.txtRoundReport.Value & "</value></property>"
    XML = XML & "</object>"
    
    'Write XML to temporary file - include POSIX path for use with curl
    TempFilePath = MacScript("return (path to temporary items from user domain) as string") & "caselist.xml"
    TempFilePOSIX = MacScript("return POSIX path of (path to temporary items from user domain) as string") & "caselist.xml"
    
    #If MAC_OFFICE_VERSION >= 15 Then
        If AppleScriptTask("Verbatim.scpt", "FileExists", TempFilePath) = "true" Then Kill TempFilePath
    #Else
        If MacScript("tell application ""Finder""" & Chr(13) & "exists file """ & TempFilePath & """" & Chr(13) & "end tell") = "true" Then Kill TempFilePath 'Kill temp file if it exists
    #End If
    
    TempFile = FreeFile
    Open TempFilePath For Output As #TempFile
    Print #TempFile, XML
    Close #TempFile

    #If MAC_OFFICE_VERSION >= 15 Then
        Script = "curl -i --data-binary '@" & TempFilePOSIX & "' "
        Script = Script & "-H 'Content-Type:application/xml' -H 'Accept:application/xml' "
        Script = Script & "-u " & GetSetting("Verbatim", "Main", "TabroomUsername", "?") & ":" & XORDecryption(GetSetting("Verbatim", "Main", "TabroomPassword", "?")) & " "
        Script = Script & "'" & Replace(CaselistURL, " ", "%20") & "'"
        XML = AppleScriptTask("Verbatim.scpt", "RunShellScript", Script)
    #Else
    
        'Construct curl request and send
        Script = "do shell script ""curl -i --data-binary '@" & TempFilePOSIX & "' "
        Script = Script & "-H 'Content-Type:application/xml' -H 'Accept:application/xml' "
        Script = Script & "-u " & GetSetting("Verbatim", "Main", "TabroomUsername", "?") & ":" & XORDecryption(GetSetting("Verbatim", "Main", "TabroomPassword", "?")) & " "
        Script = Script & "'" & Replace(CaselistURL, " ", "%20") & "'"""
        XML = MacScript(Script)

    #End If

    'Delete temp file
    Kill TempFilePath
    
    'Get the status code and update progress bar
    Select Case Mid(XML, 10, 3)
        Case Is = "201" 'Created
            'Process XML and return new round object number
            CaselistCreateRound = Mid(XML, InStr(XML, "<number>") + 8, InStr(XML, "</number>") - InStr(XML, "<number>") - 8)
            Me.lblProgress.Caption = Me.lblProgress.Caption & "done." & vbCrLf
            Me.Repaint
        Case Is = "400" 'Badly Formed
            CaselistCreateRound = ""
            Me.lblProgress.Caption = "Failed to create round due to badly formed syntax - check your round info." & vbCrLf
            Me.lblProgress.ForeColor = vbRed
            Me.Repaint
        Case Is = "401" 'Unauthorized
            CaselistCreateRound = ""
            Me.lblProgress.Caption = "Bad Username/Password - please check your tabroom account info in the Verbatim settings." & vbCrLf
            Me.lblProgress.ForeColor = vbRed
            Me.Repaint
        Case Is = "404" 'Not Found
            CaselistCreateRound = ""
            Me.lblProgress.Caption = "School/Team page not found - please check your caselist info in the Verbatim settings." & vbCrLf
            Me.lblProgress.ForeColor = vbRed
            Me.Repaint
        Case Else
            CaselistCreateRound = ""
            Me.lblProgress.Caption = "Unknown Error. HTTP Status Code: " & Mid(XML, 10, 3) & vbCrLf
            Me.lblProgress.ForeColor = vbRed
            Me.Repaint
    End Select
    
    Exit Function

Handler:
    Me.lblProgress.Caption = Me.lblProgress.Caption & "failed." & vbCrLf
    Me.Repaint
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Function

Public Function CaselistCreateCiteEntries(CaselistURL As String, ParentRound As String, CaselistSchool As String, CaselistTeam As String) As String
    
    Dim i
    
    Dim cPrefix As control
    Dim cTitle As control
    Dim cCites As control
    Dim EntryTitle As String
    Dim EntryText As String
    
    Dim Script As String
    Dim XML As String
    
    Dim TempFile
    Dim TempFilePath As String
    Dim TempFilePOSIX As String
    
    'Turn on error checking
    On Error GoTo Handler
    
    'Loop all cite entries and upload each one
    For i = 1 To Me.txtEntryCount.Value
            
        'Get entry info
        Set cPrefix = Me.fEntries.Controls("cboPrefix" & i)
        Set cTitle = Me.fEntries.Controls("txtEntryTitle" & i)
        Set cCites = Me.fEntries.Controls("txtEntryText" & i)
        
        'Create title and entry
        EntryTitle = cPrefix.Value & " " & cTitle.Value
        EntryText = cCites.Value
              
        'Create XML
        XML = "" 'Clear at beginning
        XML = XML & "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
        XML = XML & "<object xmlns=""http://www.xwiki.org"">"
        XML = XML & "<className>Caselist.CitesClass</className>"
        XML = XML & "<property name=""Team""><value>" & CaselistSchool & " " & CaselistTeam & " " & Me.cboSide.Value & "</value></property>"
        XML = XML & "<property name=""Tournament""><value>" & Me.txtTournament.Value & "</value></property>"
        XML = XML & "<property name=""Round""><value>" & Me.cboRound.Value & "</value></property>"
        XML = XML & "<property name=""Opponent""><value>" & Me.txtOpponent.Value & "</value></property>"
        XML = XML & "<property name=""Judge""><value>" & Me.txtJudge.Value & "</value></property>"
        XML = XML & "<property name=""ParentRound""><value>" & ParentRound & "</value></property>"
        XML = XML & "<property name=""Title""><value>" & EntryTitle & "</value></property>"
        XML = XML & "<property name=""Cites""><value>" & EntryText & "</value></property>"
        XML = XML & "</object>"

        'Write XML to temporary file - include POSIX path for use with curl
        TempFilePath = MacScript("return (path to temporary items from user domain) as string") & "caselist.xml"
        TempFilePOSIX = MacScript("return POSIX path of (path to temporary items from user domain) as string") & "caselist.xml"
        
        #If MAC_OFFICE_VERSION >= 15 Then
            If AppleScriptTask("Verbatim.scpt", "FileExists", TempFilePath) = "true" Then Kill TempFilePath
        #Else
            If MacScript("tell application ""Finder""" & Chr(13) & "exists file """ & TempFilePath & """" & Chr(13) & "end tell") = "true" Then Kill TempFilePath 'Kill temp file if it exists
        #End If
    
        TempFile = FreeFile
        Open TempFilePath For Output As #TempFile
        Print #TempFile, XML
        Close #TempFile
    
        #If MAC_OFFICE_VERSION >= 15 Then
            Script = "curl -i --data-binary '@" & TempFilePOSIX & "' "
            Script = Script & "-H 'Content-Type:application/xml' -H 'Accept:application/xml' "
            Script = Script & "-u " & GetSetting("Verbatim", "Main", "TabroomUsername", "?") & ":" & XORDecryption(GetSetting("Verbatim", "Main", "TabroomPassword", "?")) & " "
            Script = Script & "'" & Replace(CaselistURL, " ", "%20") & "'"
            XML = AppleScriptTask("Verbatim.scpt", "RunShellScript", Script)
        #Else
        
            'Construct curl request and send
            Script = "do shell script ""curl -i --data-binary '@" & TempFilePOSIX & "' "
            Script = Script & "-H 'Content-Type:application/xml' -H 'Accept:application/xml' "
            Script = Script & "-u " & GetSetting("Verbatim", "Main", "TabroomUsername", "?") & ":" & XORDecryption(GetSetting("Verbatim", "Main", "TabroomPassword", "?")) & " "
            Script = Script & "'" & Replace(CaselistURL, " ", "%20") & "'"""
            XML = MacScript(Script)
    
        #End If
    
        'Delete temp file
        Kill TempFilePath
               
        'Get the status code and update progress bar
        Select Case Mid(XML, 10, 3)
            Case Is = "201" 'Created
                'Process XML and return comma-delimited list of new cite object numbers
                CaselistCreateCiteEntries = CaselistCreateCiteEntries & Mid(XML, InStr(XML, "<number>") + 8, InStr(XML, "</number>") - InStr(XML, "<number>") - 8) & ","
            Case Is = "400" 'Badly Formed
                Me.lblProgress.Caption = "Failed to create a cite entry due to badly formed syntax - check your cite info. If you have more than one cite entry, the rest will still be attempted." & vbCrLf
                Me.lblProgress.ForeColor = vbRed
                Me.Repaint
            Case Is = "401" 'Unauthorized
                CaselistCreateCiteEntries = ""
                Me.lblProgress.Caption = "Bad Username/Password - please check your tabroom account info in the Verbatim settings." & vbCrLf
                Me.lblProgress.ForeColor = vbRed
                Me.Repaint
            Case Is = "404" 'Not Found
                CaselistCreateCiteEntries = ""
                Me.lblProgress.Caption = "School/Team page not found - please check your caselist info in the Verbatim settings." & vbCrLf
                Me.lblProgress.ForeColor = vbRed
                Me.Repaint
            Case Else
                CaselistCreateCiteEntries = ""
                Me.lblProgress.Caption = "Unknown Error. HTTP Status Code: " & Mid(XML, 10, 3) & vbCrLf
                Me.lblProgress.ForeColor = vbRed
                Me.Repaint
        End Select
        
    Next i

    'Update progress bar if successful
    If CaselistCreateCiteEntries <> "" Then
        Me.lblProgress.Caption = Me.lblProgress.Caption & "done." & vbCrLf
        Me.Repaint
    End If
    
    Exit Function

Handler:
    Me.lblProgress.Caption = Me.lblProgress.Caption & "failed." & vbCrLf
    Me.Repaint
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Function

Public Function UploadOpenSource(CaselistURL As String) As String

    Dim TempFilePOSIX As String
    Dim Script As String
    Dim XML As String
    
    'Turn on error checking
    On Error GoTo Handler
    
    'Make sure doc is saved
    If ActiveDocument.Saved = False Then ActiveDocument.Save
    
    'Create a temporary copy of the current file to upload
    TempFilePOSIX = MacScript("return POSIX path of (path to temporary items from user domain) as string") & ActiveDocument.Name
    #If MAC_OFFICE_VERSION >= 15 Then
        Script = "cp " & ActiveDocument.FullName & " " & TempFilePOSIX
        AppleScriptTask "Verbatim.scpt", "RunShellScript", Script
    #Else
        MacScript ("set tempfolder to path to temporary items from user domain" & Chr(13) & "tell application ""Finder""" & Chr(13) & "duplicate file """ & ActiveDocument.FullName & """ to folder tempfolder with replacing" & Chr(13) & "end tell")
    #End If
    
    #If MAC_OFFICE_VERSION >= 15 Then
        Script = "curl -i -X PUT --data-binary '@" & TempFilePOSIX & "' "
        Script = Script & "-H 'Content-Type:application/xml' -H 'Accept:application/xml' "
        Script = Script & "-u " & GetSetting("Verbatim", "Main", "TabroomUsername", "?") & ":" & XORDecryption(GetSetting("Verbatim", "Main", "TabroomPassword", "?")) & " "
        Script = Script & "'" & Replace(CaselistURL, " ", "%20") & "'"
        XML = AppleScriptTask("Verbatim.scpt", "RunShellScript", Script)
    #Else
    
        'Construct curl request and send
        Script = "do shell script ""curl -i -X PUT --data-binary '@" & TempFilePOSIX & "' "
        Script = Script & "-H 'Content-Type:application/xml' -H 'Accept:application/xml' "
        Script = Script & "-u " & GetSetting("Verbatim", "Main", "TabroomUsername", "?") & ":" & XORDecryption(GetSetting("Verbatim", "Main", "TabroomPassword", "?")) & " "
        Script = Script & "'" & Replace(CaselistURL, " ", "%20") & "'"""
        XML = MacScript(Script)
    
    #End If
    
    'Delete temp file
    Filesystem.KillFileOnMac MacScript("return path to temporary items from user domain as string") & ActiveDocument.Name
    
    'Get the status code and update progress bar
    Select Case Mid(XML, 33, 3)
        Case Is = "201" 'Created
            'Process XML and return URL of uploaded attachment
            UploadOpenSource = Mid(XML, InStr(XML, "<xwikiAbsoluteUrl>") + 18, InStr(XML, "</xwikiAbsoluteUrl>") - InStr(XML, "<xwikiAbsoluteUrl>") - 18)
        Case Is = "202" 'Created as Update
            Debug.Print "File with same name already exists on the page - Open Source posted successfully as a new version of the attachment."
            'Process XML and return URL of uploaded attachment
            UploadOpenSource = Mid(XML, InStr(XML, "<xwikiAbsoluteUrl>") + 18, InStr(XML, "</xwikiAbsoluteUrl>") - InStr(XML, "<xwikiAbsoluteUrl>") - 18)
        Case Is = "401" 'Unauthorized
            UploadOpenSource = ""
            Me.lblProgress.Caption = "Open Source upload failed due to Bad Username/Password. Check your tabroom account info in the Verbatim settings." & vbCrLf
            Me.lblProgress.ForeColor = vbRed
            Me.Repaint
        Case Is = "404" 'Not Found
            UploadOpenSource = ""
            Me.lblProgress.Caption = "Open Source upload failed because the School/Team page was not found. Check your caselist info in the Verbatim settings." & vbCrLf
            Me.lblProgress.ForeColor = vbRed
            Me.Repaint
        Case Else
            UploadOpenSource = ""
            Me.lblProgress.Caption = "Open Source upload failed due to an unknown Error. HTTP Status Code: " & Mid(XML, 33, 3) & vbCrLf
            Me.lblProgress.ForeColor = vbRed
            Me.Repaint
    End Select
    
    'Update progress bar if successful
    If UploadOpenSource <> "" Then
        Me.lblProgress.Caption = Me.lblProgress.Caption & "done." & vbCrLf
        Me.Repaint
    End If
        
    Exit Function
    
Handler:
    'Update progress bar and clean up
    Me.lblProgress.Caption = Me.lblProgress.Caption & "failed." & vbCrLf
    Me.Repaint
    Filesystem.KillFileOnMac MacScript("return path to temporary items from user domain as string") & ActiveDocument.Name
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Function

Public Sub CaselistUpdateRound(CaselistURL As String, RoundObjectNumber As String, CiteObjectNumbers As String, OSURL As String)
    
    Dim Script As String
    Dim XML As String
    
    Dim TempFile
    Dim TempFilePath As String
    Dim TempFilePOSIX As String
    
    'Turn on error checking
    On Error GoTo Handler
    
    'Update Cites Property
    XML = ""
    XML = XML & "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
    XML = XML & "<property xmlns=""http://www.xwiki.org"">"
    XML = XML & "<value>" & CiteObjectNumbers & "</value>"
    XML = XML & "</property>"
    
    'Write XML to temporary file - include POSIX path for use with curl
    TempFilePath = MacScript("return (path to temporary items from user domain) as string") & "caselist.xml"
    TempFilePOSIX = MacScript("return POSIX path of (path to temporary items from user domain) as string") & "caselist.xml"
    
    #If MAC_OFFICE_VERSION >= 15 Then
        If AppleScriptTask("Verbatim.scpt", "FileExists", TempFilePath) = "true" Then Kill TempFilePath
    #Else
        If MacScript("tell application ""Finder""" & Chr(13) & "exists file """ & TempFilePath & """" & Chr(13) & "end tell") = "true" Then Kill TempFilePath 'Kill temp file if it exists
    #End If
        
    TempFile = FreeFile
    Open TempFilePath For Output As #TempFile
    Print #TempFile, XML
    Close #TempFile

    #If MAC_OFFICE_VERSION >= 15 Then
        Script = "curl -i -X PUT --data-binary '@" & TempFilePOSIX & "' "
        Script = Script & "-H 'Content-Type:application/xml' -H 'Accept:application/xml' "
        Script = Script & "-u " & GetSetting("Verbatim", "Main", "TabroomUsername", "?") & ":" & XORDecryption(GetSetting("Verbatim", "Main", "TabroomPassword", "?")) & " "
        Script = Script & "'" & Replace(CaselistURL, " ", "%20") & "/Cites/" & "'"
        XML = AppleScriptTask("Verbatim.scpt", "RunShellScript", Script)
    #Else
    
        'Construct curl request and send
        Script = "do shell script ""curl -i -X PUT --data-binary '@" & TempFilePOSIX & "' "
        Script = Script & "-H 'Content-Type:application/xml' -H 'Accept:application/xml' "
        Script = Script & "-u " & GetSetting("Verbatim", "Main", "TabroomUsername", "?") & ":" & XORDecryption(GetSetting("Verbatim", "Main", "TabroomPassword", "?")) & " "
        Script = Script & "'" & Replace(CaselistURL, " ", "%20") & "/Cites/" & "'"""
        XML = MacScript(Script)

    #End If

    'Delete temp file
    Kill TempFilePath
    
    'Update Open Source Property
    XML = ""
    XML = XML & "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
    XML = XML & "<property xmlns=""http://www.xwiki.org"">"
    XML = XML & "<value>" & OSURL & "</value>"
    XML = XML & "</property>"
    
    'Write XML to temporary file - include POSIX path for use with curl
    TempFilePath = MacScript("return (path to temporary items from user domain) as string") & "caselist.xml"
    TempFilePOSIX = MacScript("return POSIX path of (path to temporary items from user domain) as string") & "caselist.xml"
    
    #If MAC_OFFICE_VERSION >= 15 Then
        If AppleScriptTask("Verbatim.scpt", "FileExists", TempFilePath) = "true" Then Kill TempFilePath
    #Else
        If MacScript("tell application ""Finder""" & Chr(13) & "exists file """ & TempFilePath & """" & Chr(13) & "end tell") = "true" Then Kill TempFilePath 'Kill temp file if it exists
    #End If
    
    TempFile = FreeFile
    Open TempFilePath For Output As #TempFile
    Print #TempFile, XML
    Close #TempFile

    #If MAC_OFFICE_VERSION >= 15 Then
        Script = "curl -i -X PUT --data-binary '@" & TempFilePOSIX & "' "
        Script = Script & "-H 'Content-Type:application/xml' -H 'Accept:application/xml' "
        Script = Script & "-u " & GetSetting("Verbatim", "Main", "TabroomUsername", "?") & ":" & XORDecryption(GetSetting("Verbatim", "Main", "TabroomPassword", "?")) & " "
        Script = Script & "'" & Replace(CaselistURL, " ", "%20") & "/OpenSource/" & "'"
        XML = AppleScriptTask("Verbatim.scpt", "RunShellScript", Script)
    #Else
    
        'Construct curl request and send
        Script = "do shell script ""curl -i -X PUT --data-binary '@" & TempFilePOSIX & "' "
        Script = Script & "-H 'Content-Type:application/xml' -H 'Accept:application/xml' "
        Script = Script & "-u " & GetSetting("Verbatim", "Main", "TabroomUsername", "?") & ":" & XORDecryption(GetSetting("Verbatim", "Main", "TabroomPassword", "?")) & " "
        Script = Script & "'" & Replace(CaselistURL, " ", "%20") & "/OpenSource/" & "'"""
        XML = MacScript(Script)

    #End If

    'Delete temp file
    Kill TempFilePath
      
    'Get the status code and update progress bar
    Select Case Mid(XML, 10, 3)
        Case Is = "202" 'Updated
            Me.lblProgress.Caption = Me.lblProgress.Caption & "done." & vbCrLf
            Me.lblProgress.Caption = Me.lblProgress.Caption & "Caselist Upload Successful!" & vbCrLf
            Me.Repaint
            MsgBox "Caselist Upload Successful!"
        Case Is = "400" 'Badly Formed
            Me.lblProgress.Caption = "Failed to update round due to badly formed syntax - check your round info." & vbCrLf
            Me.lblProgress.ForeColor = vbRed
            Me.Repaint
        Case Is = "401" 'Unauthorized
            Me.lblProgress.Caption = "Bad Username/Password - please check your tabroom account info in the Verbatim settings." & vbCrLf
            Me.lblProgress.ForeColor = vbRed
            Me.Repaint
        Case Is = "404" 'Not Found
            Me.lblProgress.Caption = "School/Team page not found - please check your caselist info in the Verbatim settings." & vbCrLf
            Me.lblProgress.ForeColor = vbRed
            Me.Repaint
        Case Else
            Me.lblProgress.Caption = "Unknown Error. HTTP Status Code: " & Mid(XML, 10, 3) & vbCrLf
            Me.lblProgress.ForeColor = vbRed
            Me.Repaint
    End Select

    Exit Sub

Handler:
    Me.lblProgress.Caption = Me.lblProgress.Caption & "failed." & vbCrLf
    Me.Repaint
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub
