Attribute VB_Name = "frmEmail"
Attribute VB_Base = "0{59687B23-24C0-4B73-91F3-767A815AC327}{B1BF5748-C9AB-4A28-8389-443B98D580B4}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub UserForm_Activate()

    Dim RoundArray As Variant
    Dim TabroomURL As String
    
    Dim FavoriteEmails
    Dim i
    Dim Entry
    
    'Turn on error checking
    On Error GoTo Handler
    
    'Reset Select Round box and add a blank item
    Me.cboSelectRound.Clear
    Me.cboSelectRound.AddItem
    Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 0) = ""
        
    'Get rounds from Tabroom, including emails
    RoundArray = Tabroom.GetTabroomRounds(True)
    
    'Loop Rounds and save Round info for later retrieval
    If IsArray(RoundArray) Then
        For i = 0 To UBound(RoundArray, 1)
            Me.cboSelectRound.AddItem
            Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 0) = RoundArray(i, 0) & " " & RoundArray(i, 1) & " " & RoundArray(i, 2) & " vs " & RoundArray(i, 3)
            Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 1) = RoundArray(i, 0) 'Tournament
            Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 2) = RoundArray(i, 1) 'Round Name
            Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 3) = RoundArray(i, 5) 'Student Names
            Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 4) = RoundArray(i, 6) 'Student Emails
            Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 5) = RoundArray(i, 7) 'Judge Names
            Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 6) = RoundArray(i, 8) 'Judge Emails
        Next
    End If
    
    'Populate FavoriteEmails box
    Me.lboxFavoriteEmails.Clear
    FavoriteEmails = Split(GetSetting("Verbatim", "Email", "FavoriteEmails", "?"), ";") 'Names/Emails are saved in string as Name,Email;
    For i = 0 To UBound(FavoriteEmails) - 1
        Entry = Split(FavoriteEmails(i), ",")
        Me.lboxFavoriteEmails.AddItem
        Me.lboxFavoriteEmails.List(Me.lboxFavoriteEmails.ListCount - 1, 0) = Entry(0)
        Me.lboxFavoriteEmails.List(Me.lboxFavoriteEmails.ListCount - 1, 1) = Entry(1)
    Next i

    'Show current file name
    Me.lblFileName.Caption = ActiveDocument.Name

    'Show email settings or warn user to enter
    If Email.CheckAppleMailConfigured = False Then
        Me.lblEmailInfo.ForeColor = vbRed
        Me.lblEmailInfo.Caption = "It doesn't look like you've set up an account in Apple Mail. Click the button at right to open."
    Else
        Me.lblEmailInfo.ForeColor = vbBlack
        Me.lblEmailInfo.Caption = "Apple Mail looks like it has an account configured!"
    End If

    'Readjust ColumnWidths
    Call SetScroll
 
    'If data returned from Tabroom, select first round - done after SetScroll to avoid unselecting
    If Me.cboSelectRound.ListCount > 1 Then Me.cboSelectRound.ListIndex = 1
    
    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnSend_Click()

    Dim i
    Dim SendTo As String
    Dim TempFile As String
    Dim POSIXCurrent As String
    Dim POSIXTemp As String
    
    On Error GoTo Handler
    
    'Warn if email not configured
    If Email.CheckAppleMailConfigured = False Then
        If MsgBox("Apple Mail doesn't appear to have any accounts configured - sending email is likely to fail. Please configure an email account in Apple Mail first." & vbCrLf & vbCrLf & "Try sending anyway?", vbYesNo) = vbNo Then Exit Sub
    End If
       
    'Loop all boxes and get selected emails
    For i = 0 To Me.lboxDebaters.ListCount - 1
        If Me.lboxDebaters.Selected(i) = True Then SendTo = SendTo & Me.lboxDebaters.List(i, 1) & ","
    Next i
    
    For i = 0 To Me.lboxJudges.ListCount - 1
        If Me.lboxJudges.Selected(i) = True Then SendTo = SendTo & Me.lboxJudges.List(i, 1) & ","
    Next i
    
    For i = 0 To Me.lboxFavoriteEmails.ListCount - 1
        If Me.lboxFavoriteEmails.Selected(i) = True Then SendTo = SendTo & Me.lboxFavoriteEmails.List(i, 1) & ","
    Next i
    
    'Chop off trailing comma
    If Right(SendTo, 1) = "," Then SendTo = Left(SendTo, Len(SendTo) - 1)
    
    'If nothing selected, exit
    If SendTo = "" Then
        MsgBox "No emails selected!"
        Exit Sub
    End If
    
    'Save document before sending
    ActiveDocument.Save
    
    'Check size of document, exit if larger than 5MB
    If FileLen(ActiveDocument.FullName) > 5000000 Then
        MsgBox "File is too large to send (" & Round(FileLen(ActiveDocument.FullName) / 1024, 0) & " KB)." & vbCrLf & vbCrLf & "Verbatim can only send files up to 5MB in size."
        Exit Sub
    End If
    
    'Get POSIX path of current file - shell cp works better than finder
    POSIXCurrent = MacScript("return POSIX path of """ & ActiveDocument.FullName & """")

    'Strip "Speech" if option set and the name isn't just "Speech"
    If GetSetting("Verbatim", "Paperless", "StripSpeech", True) = True And Len(ActiveDocument.Name) > 11 Then
        TempFile = MacScript("return path to temporary items from user domain as string") & Trim(Replace(ActiveDocument.Name, "speech", "", 1, -1, vbTextCompare))
        POSIXTemp = MacScript("return POSIX path of (path to temporary items from user domain as string)") & Trim(Replace(ActiveDocument.Name, "speech", "", 1, -1, vbTextCompare))
    Else
        TempFile = MacScript("return path to temporary items from user domain as string") & ActiveDocument.Name
        POSIXTemp = MacScript("return POSIX path of (path to temporary items from user domain as string)") & ActiveDocument.Name
    End If
    
    'Create a temp copy of the file
    #If MAC_OFFICE_VERSION >= 15 Then
        AppleScriptTask "Verbatim.scpt", "RunShellScript", "cp '" & POSIXCurrent & "' '" & POSIXTemp & "'"
    #Else
        MacScript ("do shell script ""cp '" & POSIXCurrent & "' '" & POSIXTemp & "'""")
    #End If

    'Send email
    Call Email.SendMail(SendTo, Me.txtSubject.Value, Me.txtMessage.Value, TempFile)
    
    'Delete the temp file
    Filesystem.KillFileOnMac TempFile
    
    'Close form
    Unload Me

    Exit Sub

Handler:
    Filesystem.KillFileOnMac TempFile
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Private Sub cboSelectRound_Change()

    Dim StudentNameArray As Variant
    Dim StudentEmailArray As Variant
    Dim JudgeNameArray As Variant
    Dim JudgeEmailArray As Variant
    Dim i
    
    On Error GoTo Handler
    
    'If list is empty, exit
    If Me.cboSelectRound.ListCount = 0 Then Exit Sub
    
    'If selected item is the first blank line, disable controls and clear boxes
    If Me.cboSelectRound.ListIndex = 0 Then
        Me.lblDebaters.Enabled = False
        Me.lboxDebaters.Clear
        Me.lboxDebaters.Enabled = False
        Me.btnNoDebaters.Enabled = False
        Me.btnAllDebaters.Enabled = False
        Me.lblJudges.Enabled = False
        Me.lboxJudges.Clear
        Me.lboxJudges.Enabled = False
        Me.btnNoJudges.Enabled = False
        Me.btnAllJudges.Enabled = False
        Me.txtSubject.Value = ""
        
    'Tabroom round is selected - enable controls
    Else
        Me.lblDebaters.Enabled = True
        Me.lboxDebaters.Enabled = True
        Me.lboxDebaters.Clear
        Me.btnNoDebaters.Enabled = True
        Me.btnAllDebaters.Enabled = True
        Me.lblJudges.Enabled = True
        Me.lboxJudges.Enabled = True
        Me.lboxJudges.Clear
        Me.btnNoJudges.Enabled = True
        Me.btnAllJudges.Enabled = True
        
        'Set the subject line
        Me.txtSubject.Value = Trim(Me.cboSelectRound.List(Me.cboSelectRound.ListIndex, 1)) & " " & Trim(Me.cboSelectRound.List(Me.cboSelectRound.ListIndex, 2))
    
        'Populate Debaters box
        StudentNameArray = Split(Me.cboSelectRound.List(Me.cboSelectRound.ListIndex, 3), ";")
        StudentEmailArray = Split(Me.cboSelectRound.List(Me.cboSelectRound.ListIndex, 4), ";")
        For i = 0 To UBound(StudentNameArray)
                Me.lboxDebaters.AddItem
                Me.lboxDebaters.List(Me.lboxDebaters.ListCount - 1, 0) = StudentNameArray(i)
                Me.lboxDebaters.List(Me.lboxDebaters.ListCount - 1, 1) = StudentEmailArray(i)
        Next
    
        'Select all debaters by default
        For i = 0 To Me.lboxDebaters.ListCount - 1
            Me.lboxDebaters.Selected(i) = True
        Next i
        
        'Populate Judges box
        JudgeNameArray = Split(Me.cboSelectRound.List(Me.cboSelectRound.ListIndex, 5), ";")
        JudgeEmailArray = Split(Me.cboSelectRound.List(Me.cboSelectRound.ListIndex, 6), ";")
        For i = 0 To UBound(JudgeNameArray)
                Me.lboxJudges.AddItem
                Me.lboxJudges.List(Me.lboxJudges.ListCount - 1, 0) = JudgeNameArray(i)
                Me.lboxJudges.List(Me.lboxJudges.ListCount - 1, 1) = JudgeEmailArray(i)
        Next

    End If

    Exit Sub
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    
End Sub

Private Sub btnNoDebaters_Click()
    Dim i
    For i = 0 To Me.lboxDebaters.ListCount - 1
        Me.lboxDebaters.Selected(i) = False
    Next i
End Sub

Private Sub btnAllDebaters_Click()
    Dim i
    For i = 0 To Me.lboxDebaters.ListCount - 1
        Me.lboxDebaters.Selected(i) = True
    Next i
End Sub

Private Sub btnNoJudges_Click()
    Dim i
    For i = 0 To Me.lboxJudges.ListCount - 1
        Me.lboxJudges.Selected(i) = False
    Next i
End Sub

Private Sub btnAllJudges_Click()
    Dim i
    For i = 0 To Me.lboxJudges.ListCount - 1
        Me.lboxJudges.Selected(i) = True
    Next i
End Sub

Private Sub btnAddEmail_Click()
    Dim i
    Dim FavoriteEmails As String
    
    'If no email is entered, exit
    If Me.txtManualEmail.Value = "" Then Exit Sub
    
    'Validate email address - @ and . required, "," and ";" prohibited because those are delimiters
    If InStr(Me.txtManualEmail.Value, "@") < 1 Or InStr(Me.txtManualEmail.Value, ".") < 1 Or InStr(Me.txtManualEmail.Value, ",") > 0 Or InStr(Me.txtManualEmail.Value, ";") > 0 Then
        MsgBox "Please enter a valid email address."
        Exit Sub
    End If

    'Checks passed, add a selected item at the top of the box, use name if filled out
    Me.lboxFavoriteEmails.AddItem , 0
    Me.lboxFavoriteEmails.List(0, 1) = Me.txtManualEmail.Value
    If Me.txtManualName.Value <> "" Then Me.lboxFavoriteEmails.List(0, 0) = Me.txtManualName.Value
    Me.lboxFavoriteEmails.Selected(0) = True
    
    'Save the updated Favorites list to the registry
    For i = 0 To Me.lboxFavoriteEmails.ListCount - 1
        FavoriteEmails = FavoriteEmails & Me.lboxFavoriteEmails.List(i, 0) & "," & Me.lboxFavoriteEmails.List(i, 1) & ";"
    Next i
    SaveSetting "Verbatim", "Email", "FavoriteEmails", FavoriteEmails

    'Reset the boxes
    Me.txtManualName.Value = ""
    Me.txtManualEmail.Value = ""

End Sub

Private Sub btnDeleteSelected_Click()
    Dim i
    Dim FavoriteEmails As String
    
    'Confirm deletion, then step backwards through list and remove selected items to prevent re-indexing list
    If MsgBox("Are you sure you want to remove the selected emails from your favorites?", vbYesNo) = vbYes Then
        For i = Me.lboxFavoriteEmails.ListCount - 1 To 0 Step -1
            If Me.lboxFavoriteEmails.Selected(i) = True Then Me.lboxFavoriteEmails.RemoveItem (i)
        Next i
    End If
    
    'Save the updated Favorites list to the registry
    For i = 0 To Me.lboxFavoriteEmails.ListCount - 1
        FavoriteEmails = FavoriteEmails & Me.lboxFavoriteEmails.List(i, 0) & "," & Me.lboxFavoriteEmails.List(i, 1) & ";"
    Next i
    SaveSetting "Verbatim", "Email", "FavoriteEmails", FavoriteEmails
    
End Sub

Private Sub btnDeleteAll_Click()
    'Confirm deletion, then clear the box and the registry
    If MsgBox("Are you sure you want to remove all emails from your favorites?", vbYesNo) = vbYes Then
        Me.lboxFavoriteEmails.Clear
    End If
    
    SaveSetting "Verbatim", "Email", "FavoriteEmails", ""

End Sub

Private Sub btnSettings_Click()
    #If MAC_OFFICE_VERSION >= 15 Then
        AppleScriptTask "Verbatim.scpt", "MailSettings", ""
    #Else
        MacScript ("tell application ""Mail"" to activate")
    #End If
End Sub

Private Sub SetScroll()
    
    Dim i
    Dim NameWidth As Integer
    Dim EmailWidth As Integer
    
    'Set default column widths
    NameWidth = 110
    EmailWidth = 156
    
    'Loop through the Debaters listbox, set column widths if larger than the default
    If Me.lboxDebaters.ListCount <> 0 Then
        For i = 0 To Me.lboxDebaters.ListCount - 1
            Me.txtResizeName = Me.lboxDebaters.List(i, 0)
            Me.txtResizeEmail = Me.lboxDebaters.List(i, 1)
            If Me.txtResizeName.Width > NameWidth Then NameWidth = Me.txtResizeName.Width
            If Me.txtResizeEmail.Width > EmailWidth Then EmailWidth = Me.txtResizeEmail.Width
        Next i

        Me.lboxDebaters.ColumnWidths = NameWidth & ";" & EmailWidth
    End If
    
    'Reset the defaults
    NameWidth = 110
    EmailWidth = 156
    
    'Loop through the Judges listbox, set column widths if larger than the default
    If Me.lboxJudges.ListCount <> 0 Then
        For i = 0 To Me.lboxJudges.ListCount - 1
            Me.txtResizeName = Me.lboxJudges.List(i, 0)
            Me.txtResizeEmail = Me.lboxJudges.List(i, 1)
            If Me.txtResizeName.Width > NameWidth Then NameWidth = Me.txtResizeName.Width
            If Me.txtResizeEmail.Width > EmailWidth Then EmailWidth = Me.txtResizeEmail.Width
        Next i

        Me.lboxJudges.ColumnWidths = NameWidth & ";" & EmailWidth
    End If
    
    'Reset the defaults
    NameWidth = 110
    EmailWidth = 156
    
    'Loop through the Favorites listbox, set column widths if larger than the default
    If Me.lboxFavoriteEmails.ListCount <> 0 Then
        For i = 0 To Me.lboxFavoriteEmails.ListCount - 1
            Me.txtResizeName = Me.lboxFavoriteEmails.List(i, 0)
            Me.txtResizeEmail = Me.lboxFavoriteEmails.List(i, 1)
            If Me.txtResizeName.Width > NameWidth Then NameWidth = Me.txtResizeName.Width
            If Me.txtResizeEmail.Width > EmailWidth Then EmailWidth = Me.txtResizeEmail.Width
        Next i

        Me.lboxFavoriteEmails.ColumnWidths = NameWidth & ";" & EmailWidth
    End If
    
End Sub
