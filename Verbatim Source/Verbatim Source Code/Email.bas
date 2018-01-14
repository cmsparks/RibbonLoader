Attribute VB_Name = "Email"
Option Explicit

Sub ShowEmailForm()
    Dim EmailForm As frmEmail
    Set EmailForm = New frmEmail
    EmailForm.Show
End Sub

Sub SendMail(SendTo As String, Subject As String, Message As String, Optional Attachment As String)

    Dim Script As String
    
    On Error GoTo Handler
    
    'Create script
    Script = Script & "tell application " & Chr(34) & "Mail" & Chr(34) & Chr(13)
    Script = Script & "set NewMail to make new outgoing message with properties "
    Script = Script & "{content:""" & Message & """, subject:""" & Subject & """ , visible:true}" & Chr(13)
    Script = Script & "tell NewMail" & Chr(13)
    
    If SendTo <> "" Then Script = Script & "make new to recipient at end of to recipients with properties " & "{address:""" & SendTo & """}" & Chr(13)

    If Attachment <> "" Then
        Script = Script & "tell content" & Chr(13)
        Script = Script & "make new attachment with properties " & "{file name:""" & Attachment & """ as alias} " & "at after the last paragraph" & Chr(13)
        Script = Script & "end tell" & Chr(13)
    End If

    Script = Script & "send" & Chr(13)
    Script = Script & "end tell" & Chr(13)
    Script = Script & "end tell"

    'Run script
    #If MAC_OFFICE_VERSION >= 15 Then
        AppleScriptTask "Verbatim.scpt", "SendMail", Message & ";" & Subject & ";" & SendTo & ";" & Attachment
    #Else
        MacScript (Script)
    #End If
    
    Exit Sub
    
Handler:
    If Err.Number = 5 Then
        MsgBox "Sending mail failed - Apple Mail appears to be configured incorrectly."
    Else
        MsgBox "Sending mail failed. Error " & Err.Number & ": " & Err.Description
    End If
End Sub

Public Function CheckAppleMailConfigured() As Boolean
    
    Dim AccountsPlist As String
    
    'Get the path to the Apple Mail Accounts.plist file
    AccountsPlist = MacScript("return POSIX path of (path to library folder from user domain) as string") & "Mail/V2/MailData/Accounts.plist"
    
    'If Accounts.plist contains the string IMAPAccount or POPAccount, an account is configured, otherwise it hasn't been set up
    If MacScript("do shell script ""grep -e IMAPAccount -e POPAccount '" & AccountsPlist & "' | wc -l""") > 0 Then
        CheckAppleMailConfigured = True
    Else
        CheckAppleMailConfigured = False
    End If

End Function
