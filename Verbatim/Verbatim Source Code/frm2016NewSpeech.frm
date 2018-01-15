Attribute VB_Name = "frm2016NewSpeech"
Attribute VB_Base = "0{EFCEFE10-1537-4FE3-AB7C-451DD93C6832}{CC098899-5D29-49C2-902E-679AA7A24551}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub UserForm_Initialize()
    
    Dim RoundArray As Variant
    Dim i
    Dim Tournament As String
    Dim RoundNum As String
    Dim Side As String
    Dim Opponent As String
    
    On Error GoTo Handler
    
    'Add default speech names
    Me.lboxSpeeches.AddItem "2AC"
    Me.lboxSpeeches.AddItem "1AR"
    Me.lboxSpeeches.AddItem "2AR"
    Me.lboxSpeeches.AddItem "1AC"
    
    Me.lboxSpeeches.AddItem "1NC"
    Me.lboxSpeeches.AddItem "2NC"
    Me.lboxSpeeches.AddItem "1NR"
    Me.lboxSpeeches.AddItem "2NR"
    
    Me.lboxSpeeches.AddItem "New Document"
    
    'Exit if info not entered in settings
    If GetSetting("Verbatim", "Main", "TabroomUsername", "?") = "?" Or GetSetting("Verbatim", "Main", "TabroomPassword", "?") = "?" Then Exit Sub
    
    'Get rounds from Tabroom
    RoundArray = Tabroom.GetTabroomRounds(Email:=True)
    
    'Loop Rounds and compute speech name
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
                    Me.lboxSpeeches.AddItem "1AC" & " " & Tournament & " " & RoundNum & " vs " & Opponent, 0
                    Me.lboxSpeeches.AddItem "2AR" & " " & Tournament & " " & RoundNum & " vs " & Opponent, 0
                    Me.lboxSpeeches.AddItem "1AR" & " " & Tournament & " " & RoundNum & " vs " & Opponent, 0
                    Me.lboxSpeeches.AddItem "2AC" & " " & Tournament & " " & RoundNum & " vs " & Opponent, 0
                    
                Else
                    Me.lboxSpeeches.AddItem "2NR" & " " & Tournament & " " & RoundNum & " vs " & Opponent, 0
                    Me.lboxSpeeches.AddItem "1NR" & " " & Tournament & " " & RoundNum & " vs " & Opponent, 0
                    Me.lboxSpeeches.AddItem "2NC" & " " & Tournament & " " & RoundNum & " vs " & Opponent, 0
                    Me.lboxSpeeches.AddItem "1NC" & " " & Tournament & " " & RoundNum & " vs " & Opponent, 0
                End If
                
            Next
        End If
    End If
    
    Exit Sub
    
Handler:
        MsgBox "Error " & Err.Number & ": " & Err.Description
    
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnCreateSpeech_Click()

    Dim i As Integer
    Dim SpeechSelected As Boolean
    
    Dim FileName As String
    Dim h As String
    
    Dim AutoSaveDirectory As String

    'Make sure a speech is selected
    For i = 0 To Me.lboxSpeeches.ListCount - 1
        If Me.lboxSpeeches.Selected(i) = True Then SpeechSelected = True
    Next i

    If SpeechSelected = False Then
        MsgBox "You must select a speech first."
    Else
        
        'Set FileName from box
        FileName = Me.lboxSpeeches.Value
        
        'Unload the menu
        Unload Me
        
        'If New Document selected, just add a new document and quit
        If FileName = "New Document" Then
            Call Paperless.NewDocument
            Exit Sub
        Else
        
            'Add a new document based on the template
            Call Paperless.NewDocument
            
            'If filename  is just the speech name, add a date
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
        End If
    End If

End Sub

