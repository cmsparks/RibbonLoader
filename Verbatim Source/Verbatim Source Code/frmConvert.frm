Attribute VB_Name = "frmConvert"
Attribute VB_Base = "0{F44306B3-65A4-4637-AFFD-1E04AC6E1C5B}{26368222-AC0C-4044-B70C-960F79E5816C}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub UserForm_Initialize()

    'Clear the text boxes
    Me.txtConvertFile.Value = ""
    Me.optConvert1 = True

End Sub

Private Sub btnCancelConvert_Click()
    Unload Me
End Sub

Private Sub btnConvert_Click()

    Select Case Me.mpgConvert.Value
    
        Case Is = 0 'ConvertThis is selected
            If Me.optConvert1 Then
                Convert.ConvertBackfile File:=ActiveDocument.FullName, ConvertFrom:=1 'Verbatim 3
            ElseIf Me.optConvert2 Then
                Convert.ConvertBackfile File:=ActiveDocument.FullName, ConvertFrom:=2 'Verbatim 2
            ElseIf Me.optConvert3 Then
                Convert.ConvertBackfile File:=ActiveDocument.FullName, ConvertFrom:=3 'Non-Verbatim
            ElseIf Me.optConvert4 Then
                Convert.ConvertBackfile File:=ActiveDocument.FullName, ConvertFrom:=4 'Synergy
            End If
    
        Case Is = 1 'ConvertFile is selected
            If Me.optConvert1 Then
                Convert.ConvertBackfile File:=txtConvertFile.Value, ConvertFrom:=1 'Verbatim 3
            ElseIf Me.optConvert2 Then
                Convert.ConvertBackfile File:=txtConvertFile.Value, ConvertFrom:=2 'Verbatim 2
            ElseIf Me.optConvert3 Then
                Convert.ConvertBackfile File:=txtConvertFile.Value, ConvertFrom:=3 'Non-Verbatim
            ElseIf Me.optConvert4 Then
                Convert.ConvertBackfile File:=txtConvertFile.Value, ConvertFrom:=4 'Synergy
            End If
    End Select
        
    'Close form after conversion
    Unload Me

End Sub

Private Sub btnConvertFileBrowse_Click()

    On Error Resume Next
    
    Dim FileOpen As Word.Dialog
    Dim Result
    
    'Show the built-in file picker, only allow picking 1 file at a time
    Set FileOpen = Word.Dialogs(wdDialogFileOpen)
    Result = FileOpen.Display
    Select Case Result
    Case -2, -1 'close, OK
        'Populate the Folder with the current directory, set by the file dialog
        Me.txtConvertFile.Value = FileOpen.Name
    Case 0 'cancel
        Exit Sub
    Case Else
        Exit Sub
    End Select
    
    Set FileOpen = Nothing

End Sub
