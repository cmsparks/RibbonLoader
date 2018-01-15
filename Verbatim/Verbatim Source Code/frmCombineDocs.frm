Attribute VB_Name = "frmCombineDocs"
Attribute VB_Base = "0{C0B8850C-A2D2-4C98-8FFF-CD67906FE1B8}{207FC747-E585-4F89-A50B-C621D40F6A79}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub UserForm_Initialize()

    Dim rFile As RecentFile
    
    Dim RoundArray As Variant
    Dim i
    
    'Turn off error checking for recent files - Mac Word has issues with the collection
    On Error Resume Next
    
    'Add all recent files to the box
    For Each rFile In Application.RecentFiles
        Me.lboxRecentDocs.AddItem
        Me.lboxRecentDocs.List(Me.lboxRecentDocs.ListCount - 1, 0) = rFile.Name
        
        '2016 only returns the path, not the full path
        If Application.Version >= "15" Then
            Me.lboxRecentDocs.List(Me.lboxRecentDocs.ListCount - 1, 1) = rFile.Path & "/" & rFile.Name
        Else
            Me.lboxRecentDocs.List(Me.lboxRecentDocs.ListCount - 1, 1) = rFile.Path
        End If
    Next rFile
       
    'Turn on error checking
    On Error GoTo Handler
       
    'Exit if info not entered in settings
    If GetSetting("Verbatim", "Main", "TabroomUsername", "?") = "?" Or GetSetting("Verbatim", "Main", "TabroomPassword", "?") = "?" Or GetSetting("Verbatim", "Caselist", "CaselistSchoolName", "?") = "?" Or GetSetting("Verbatim", "Caselist", "CaselistTeamName", "?") = "?" Then Exit Sub
    
    'Reset AutoName box and add a blank item
    Me.cboAutoName.Clear
    Me.cboAutoName.AddItem
    Me.cboAutoName.List(Me.cboAutoName.ListCount - 1) = ""
    
    'Get rounds from Tabroom
    RoundArray = Tabroom.GetTabroomRounds()
    
    'Loop Rounds and compute file name
    If IsArray(RoundArray) Then
        For i = 0 To UBound(RoundArray, 1)
            Me.cboAutoName.AddItem
            Me.cboAutoName.List(Me.cboAutoName.ListCount - 1) = GetSetting("Verbatim", "Caselist", "CaselistSchoolName") & "-" & Replace(GetSetting("Verbatim", "Caselist", "CaselistTeamName"), " ", "-") & "-" & RoundArray(i, 2) & "-" & RoundArray(i, 0) & "-" & Replace(Trim(RoundArray(i, 1)), " ", "-") & ".docx"
        Next
    End If
    
    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnManualAdd_Click()

    Dim FilePath As String
    
    On Error Resume Next
            
    'Show the file picker
    FilePath = MacScript("set FilePath to (choose file with prompt ""Choose file to combine:"") as string")
    
    'Add selected file to the box
    If FilePath <> "" Then
        Me.lboxManualDocs.AddItem , 0
        Me.lboxManualDocs.List(0, 0) = Right(FilePath, Len(FilePath) - InStrRev(FilePath, ":"))
        Me.lboxManualDocs.List(0, 1) = FilePath
        Me.lboxManualDocs.Selected(0) = True
    End If
        
End Sub

Private Sub btnCombine_Click()

    Dim i As Integer
    Dim FileCount As Integer
    
    On Error GoTo Handler
    
    'Make sure only docx, doc and rtf files are selected
    For i = 0 To Me.lboxRecentDocs.ListCount - 1
        If Me.lboxRecentDocs.Selected(i) = True Then
            If Right(Me.lboxRecentDocs.List(i, 1), Len(Me.lboxRecentDocs.List(i, 1)) - InStrRev(Me.lboxRecentDocs.List(i, 1), ".")) <> "docx" And _
            Right(Me.lboxRecentDocs.List(i, 1), Len(Me.lboxRecentDocs.List(i, 1)) - InStrRev(Me.lboxRecentDocs.List(i, 1), ".")) <> "doc" And _
            Right(Me.lboxRecentDocs.List(i, 1), Len(Me.lboxRecentDocs.List(i, 1)) - InStrRev(Me.lboxRecentDocs.List(i, 1), ".")) <> "rtf" Then
                MsgBox "You can only combine .docx, .doc, and .rtf files - please deselect other file formats before proceeding."
                Exit Sub
            End If
            FileCount = FileCount + 1
        End If
    Next i
    For i = 0 To Me.lboxManualDocs.ListCount - 1
         If Me.lboxManualDocs.Selected(i) = True Then
            If Right(Me.lboxManualDocs.List(i, 1), Len(Me.lboxManualDocs.List(i, 1)) - InStrRev(Me.lboxManualDocs.List(i, 1), ".")) <> "docx" And _
            Right(Me.lboxManualDocs.List(i, 1), Len(Me.lboxManualDocs.List(i, 1)) - InStrRev(Me.lboxManualDocs.List(i, 1), ".")) <> "doc" And _
            Right(Me.lboxManualDocs.List(i, 1), Len(Me.lboxManualDocs.List(i, 1)) - InStrRev(Me.lboxManualDocs.List(i, 1), ".")) <> "rtf" Then
                MsgBox "You can only combine .docx, .doc, and .rtf files - please deselect other file formats before proceeding."
                Exit Sub
            End If
            FileCount = FileCount + 1
        End If
    Next i
    
    'Make sure at least 2 files are selected
    If FileCount < 2 Then
        MsgBox "You must select at least 2 files to combine."
        Exit Sub
    End If
        
    'Add a new blank document
    Call Paperless.NewDocument
  
    'Insert selected files in new pockets
    For i = 0 To Me.lboxRecentDocs.ListCount - 1
        If Me.lboxRecentDocs.Selected(i) = True Then
            Selection.TypeText Left(Me.lboxRecentDocs.List(i, 0), InStrRev(Me.lboxRecentDocs.List(i, 0), ".") - 1)
            Selection.Style = "Pocket"
            Selection.TypeParagraph
            Selection.InsertFile Me.lboxRecentDocs.List(i, 1)
        End If
    Next i

    For i = 0 To Me.lboxManualDocs.ListCount - 1
        If Me.lboxManualDocs.Selected(i) = True Then
            Selection.TypeText Left(Me.lboxManualDocs.List(i, 0), InStrRev(Me.lboxManualDocs.List(i, 0), ".") - 1)
            Selection.Style = "Pocket"
            Selection.TypeParagraph
            Selection.InsertFile Me.lboxManualDocs.List(i, 1)
        End If
    Next i
    
    'Save file
    If GetSetting("Verbatim", "Paperless", "AutoSaveDir", "?") <> "?" And Me.cboAutoName.Value <> "" Then
        If Right(Trim(GetSetting("Verbatim", "Paperless", "AutoSaveDir", "?")), 1) = ":" Then
            ActiveDocument.SaveAs FileName:=GetSetting("Verbatim", "Paperless", "AutoSaveDir", "?") & Me.cboAutoName.Value, FileFormat:=wdFormatXMLDocument
        Else
            ActiveDocument.SaveAs FileName:=GetSetting("Verbatim", "Paperless", "AutoSaveDir", "?") & ":" & Me.cboAutoName.Value, FileFormat:=wdFormatXMLDocument
        End If
    Else
        With Application.Dialogs(wdDialogFileSaveAs)
            If Me.cboAutoName.Value <> "" Then
                .Name = Me.cboAutoName.Value
            Else
                .Name = "Combined Doc"
            End If
            .Show
        End With
    End If
    
    Unload Me
    
    Exit Sub
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub
