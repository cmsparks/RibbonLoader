Attribute VB_Name = "PaDS"
Option Explicit

'Globals to ensure menu ID's, button ID's etc. increment correctly
Public CoauthoringRootNode As CommandBarControl
Public CoauthoringMenuIDNumber As Long
Public CoauthoringButtonIDNumber As Long
Public CoauthoringDepth As Long

'*************************************************************************************
'* TOOLBAR FUNCTIONS                                                                                         *
'*************************************************************************************

Sub AutoCoauthoring()
    
    Dim PressedControl As CommandBarControl

    On Error GoTo Handler

    'Get pressed control
    Set PressedControl = CommandBars.ActionControl
    If PressedControl Is Nothing Then Exit Sub

    'If mode is off, turn it on
    If AutoCoauthoringToggle = False Then
            
        'Warn if current document can't be coauthored
        If ActiveDocument.CoAuthoring.CanShare = False Then
            MsgBox "Error - current document cannot be coauthored. Check that you are using Word 2011+, and the file is a .docx and saved to a Sharepoint server like PaDS."
            AutoCoauthoringToggle = False
            Exit Sub
        End If
    
        AutoCoauthoringToggle = True
        PressedControl.Caption = "Turn Off Auto Coauthoring Updates"
        MsgBox "Automatic coauthoring updates are turned ON. Note that this will save your document every time someone else edits it - use only if you trust others edits and can risk losing the ability to undo your own changes."
      
        Do
            DoEvents 'Give control back to application
      
            'If coauthoring updates exist, save the file to get them
            If ActiveDocument.CoAuthoring.PendingUpdates Then
                ActiveDocument.Save
            End If
        Loop Until AutoCoauthoringToggle = False 'Loop until button is pressed again
      
    Else
        AutoCoauthoringToggle = False
        PressedControl.Caption = "Turn On Auto Coauthoring Updates"
        MsgBox "Automatic coauthoring updates are turned OFF."
        
    End If
    
    Set PressedControl = Nothing
    
    Exit Sub

Handler:
    AutoCoauthoringToggle = False
    PressedControl.Caption = "Turn On Auto Coauthoring Updates"
    Set PressedControl = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Sub GetCoauthoringContent(Optional FromScratch As Boolean)
'Get content for dynamic menu from PaDS
        
    Dim c
    Dim Button As CommandBarControl
    
    On Error GoTo Handler
    
    'Check URL isn't blank, or prompt for settings
    If GetSetting("Verbatim", "PaDS", "CoauthoringFolder", "?") = "?" Then
        If MsgBox("You have not set a PaDS URL. Open Settings?", vbOKCancel) = vbOK Then
          Call Settings.ShowSettingsForm
          Exit Sub
        Else
          Exit Sub
        End If
    End If
        
    'Check for tabroom username/password
    If GetSetting("Verbatim", "Main", "TabroomUsername", "?") = "?" Or GetSetting("Verbatim", "Main", "TabroomPassword", "?") = "?" Then
        If MsgBox("You have not configured your tabroom username/password. Open Settings?", vbOKCancel) = vbOK Then
          Call Settings.ShowSettingsForm
          Exit Sub
        Else
          Exit Sub
        End If
    End If
        
    'Get root coauthoring menu
    Set CoauthoringRootNode = CommandBars.FindControl(Tag:="CoauthoringMenu")
    
    'If building from scratch, delete existing controls
    If FromScratch = True Then
        For Each c In CoauthoringRootNode.Controls
            c.Delete
        Next c
    End If
    
    'Exit if already built
    If CoauthoringRootNode.Controls.Count > 0 Then Exit Sub
    
    'Set Mouse Pointer and update status bar
    System.Cursor = wdCursorWait
    ProgressBar = "Refreshing coauthoring menu "
    Application.StatusBar = ProgressBar
        
    'Add Upload Button
    Set Button = CoauthoringRootNode.Controls.Add(Type:=msoControlButton)
    Button.Caption = "Upload To PaDS"
    Button.DescriptionText = "Uploads the document to PaDS to coauthor (Cmd+Alt+S). Coauthoring location can be configured in the Verbatim settings. (Ctrl+Alt+S)"
    Button.FaceId = 1756 'Heads
    Button.Style = msoButtonIconAndCaption
    Button.Tag = "UploadToPaDS"
    Button.OnAction = "PaDS.UploadToPaDS"
    
    'Separator
    Set Button = CoauthoringRootNode.Controls.Add(Type:=msoControlButton)
    Button.Caption = ""
    Button.Tag = "CoauthoringSeparator1"
    Button.Enabled = False
    
    'Reset counters
    CoauthoringDepth = 0
    CoauthoringMenuIDNumber = 0
    CoauthoringButtonIDNumber = 0
    
    'Seed the recursion with the top level of the coauthoring folder
    Call PaDS.GetCoauthoringRecursion(PaDS.SharepointGetFolderContents(URL:=GetSetting("Verbatim", "PaDS", "CoauthoringFolder"), SubFolders:=True), CoauthoringRootNode)
    
    'Separator
    Set Button = CoauthoringRootNode.Controls.Add(Type:=msoControlButton)
    Button.Caption = ""
    Button.Tag = "CoauthoringSeparator2"
    Button.Enabled = False
    
    'Add Open button
    Set Button = CoauthoringRootNode.Controls.Add(Type:=msoControlButton)
    Button.Caption = "Open PaDS Folder"
    Button.DescriptionText = "Opens a document from your PaDS Coauthoring location (Cmd+Alt+O). Coauthoring location can be configured in the Verbatim settings. (Ctrl+Alt+O)"
    Button.FaceId = 23 'Open Folder
    Button.Style = msoButtonIconAndCaption
    Button.Tag = "OpenPaDSFolder"
    Button.OnAction = "PaDS.OpenFromPaDS"
    
    'Add Refresh button
    Set Button = CoauthoringRootNode.Controls.Add(Type:=msoControlButton)
    Button.Caption = "Refresh Coauthor Menu"
    Button.DescriptionText = "Refreshes your coauthoring folder."
    Button.FaceId = 8085 'Blue refresh
    Button.Style = msoButtonIconAndCaption
    Button.Tag = "RefreshCoauthoring"
    Button.OnAction = "Toolbar.AssignButtonActions"
    
    'Auto Coauthoring button
    Set Button = CoauthoringRootNode.Controls.Add(Type:=msoControlButton)
    If AutoCoauthoringToggle = False Then
        Button.Caption = "Turn On Auto Coauthoring Updates"
    Else
        Button.Caption = "Turn Off Auto Coauthoring Updates"
    End If
    Button.DescriptionText = "Automatically saves and updates your document any time coauthoring updates are available."
    Button.FaceId = 1020 'Refresh page
    Button.Style = msoButtonIconAndCaption
    Button.Tag = "RefreshCoauthoring"
    Button.OnAction = "PaDS.AutoCoauthoring"
    
    'Add Settings button
    Set Button = CoauthoringRootNode.Controls.Add(Type:=msoControlButton)
    Button.Caption = "PaDS Settings"
    Button.DescriptionText = "Opens the Verbatim settings to configure coauthoring."
    Button.FaceId = 2144 'Gears
    Button.Style = msoButtonIconAndCaption
    Button.Tag = "CoauthoringSettings"
    Button.OnAction = "Settings.ShowSettingsForm"
    
    'Set template as saved to avoid prompts
    ActiveDocument.AttachedTemplate.Saved = True
    
    'Clean up
    System.Cursor = wdCursorNormal
    Application.StatusBar = "Coauthoring menu updated!"
    Set CoauthoringRootNode = Nothing
    Set Button = Nothing
    
    Exit Sub

Handler:
    System.Cursor = wdCursorNormal
    Set CoauthoringRootNode = Nothing
    Set Button = Nothing
    Application.StatusBar = "Coauthoring menu update failed!"
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Private Sub GetCoauthoringRecursion(NodeList, Parent As CommandBarControl)

    Dim Elem
    Dim Menu As CommandBarControl
    Dim Button As CommandBarControl
    Dim FileNodeList

    'Increment progress bar
    ProgressBar = ProgressBar & ChrW(9609)
    Application.StatusBar = ProgressBar
    
    'Increment depth counter
    CoauthoringDepth = CoauthoringDepth + 1
    
    'Iterate through each subfolder in first 4 depth levels - XML menu is limited to 5 levels and we need 1 for the file
    If CoauthoringDepth < 5 Then
    
        'Get each subfolder from the passed in list
        For Each Elem In NodeList
            'Skip if built-in "Forms" folder
            If Right(Elem, 5) <> "Forms" Then
                'Create a menu for the subfolder and append it
                CoauthoringMenuIDNumber = CoauthoringMenuIDNumber + 1
                Set Menu = Parent.Controls.Add(Type:=msoControlPopup)
                Menu.Caption = Right(Elem, Len(Elem) - InStrRev(Elem, "/"))
                Menu.Tag = "CoauthoringMenu" & CoauthoringMenuIDNumber
                Menu.Parameter = Elem
                
                'Reseed the recursion macro with the new subfolder node - this ensures a loop to the bottom level
                Call PaDS.GetCoauthoringRecursion(PaDS.SharepointGetFolderContents(URL:="http://" & GetSetting("Verbatim", "PaDS", "PaDSSiteName") & ".paperlessdebate.com" & Elem, SubFolders:=True), Menu)
                
                'Decrement depth level when coming back out of subfolder
                CoauthoringDepth = CoauthoringDepth - 1
                
            End If
        Next Elem
    End If
        
    'Get files list for current folder
    If CoauthoringDepth = 1 Then
        FileNodeList = PaDS.SharepointGetFolderContents(URL:=GetSetting("Verbatim", "PaDS", "CoauthoringFolder"), SubFolders:=False)
    Else
        FileNodeList = PaDS.SharepointGetFolderContents(URL:="http://" & GetSetting("Verbatim", "PaDS", "PaDSSiteName") & ".paperlessdebate.com/" & Parent.Parameter, SubFolders:=False)
    End If
    
    'Increment progress bar
    ProgressBar = ProgressBar & ChrW(9609)
    Application.StatusBar = ProgressBar
    
    'Iterate through each file in the folder
    For Each Elem In FileNodeList
                        
        'Create a button for the file
        CoauthoringButtonIDNumber = CoauthoringButtonIDNumber + 1 'Increment Menu number to ensure a unique ID
        Set Button = Parent.Controls.Add(Type:=msoControlButton)
        Button.Caption = Right(Elem, Len(Elem) - InStrRev(Elem, "/"))
        Button.FaceId = 1544 'Blank page
        Button.Style = msoButtonIconAndCaption
        Button.Tag = "CoauthoringButton" & CoauthoringButtonIDNumber
        Button.Parameter = Elem
        Button.OnAction = "PaDS.OpenFromPaDS"
      
    Next Elem
    
    'Clean up
    Set Menu = Nothing
    Set Button = Nothing
    
    Exit Sub
    
Handler:
    Set Menu = Nothing
    Set Button = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

'*************************************************************************************
'* UPLOAD/OPEN FUNCTIONS                                                                                 *
'*************************************************************************************

Sub UploadToPaDS(Optional UploadToPublic As Boolean)

    Dim PaDSURL As String
    Dim PaDSFileName As String
    
    'Turn on error checking
    On Error GoTo Handler
     
    'Get PaDS URL from Registry
    If UploadToPublic = True Then
        PaDSURL = GetSetting("Verbatim", "PaDS", "PublicFolder", "?")
    Else
        PaDSURL = GetSetting("Verbatim", "PaDS", "CoauthoringFolder", "?")
    End If
    
    'Check URL isn't blank, or prompt for settings
    If PaDSURL = "?" Then
        If MsgBox("You have not set a PaDS URL. Open Settings?", vbOKCancel) = vbOK Then
          Call Settings.ShowSettingsForm
          Exit Sub
        Else
          Exit Sub
        End If
    End If
        
    'Check for tabroom username/password
    If GetSetting("Verbatim", "Main", "TabroomUsername", "?") = "?" Or GetSetting("Verbatim", "Main", "TabroomPassword", "?") = "?" Then
        If MsgBox("You have not configured your tabroom username/password. Open Settings?", vbOKCancel) = vbOK Then
          Call Settings.ShowSettingsForm
          Exit Sub
        Else
          Exit Sub
        End If
    End If
        
    'Save File locally
    ActiveDocument.Save

    'Create file name
    If Right(PaDSURL, 1) = "/" Then
        PaDSFileName = PaDSURL & ActiveDocument.Name
    Else
        PaDSFileName = PaDSURL & "/" & ActiveDocument.Name
    End If
        
    'Check if file already exists
    If PaDS.SharepointURLExists(PaDSFileName) = True Then
        If MsgBox("File Exists.  Overwrite?", vbOKCancel) = vbCancel Then Exit Sub
    End If
   
    'Upload the file to PaDS using SaveAs - ensures the file is open on PaDS during speech for marking
    ActiveDocument.SaveAs FileName:=PaDSFileName

    'Report success
    MsgBox "Upload Successful. You are now working off PaDS:" & vbCrLf & PaDSFileName

    Exit Sub
    
Handler:
    If Err.Number = 4198 Then
        MsgBox "Upload failed. Try opening a file from the coauthoring menu first, then retry the upload."
    Else
        MsgBox "Error " & Err.Number & ": " & Err.Description
    End If
End Sub

Sub PaDSPublic()
    Call PaDS.UploadToPaDS(UploadToPublic:=True)
End Sub

Sub OpenFromPaDS()
    
    Dim PressedControl As CommandBarButton
    
    Dim URL As String
    Dim Script As String
    
    'Check Coauthoring URL isn't blank, or prompt for settings
    If GetSetting("Verbatim", "PaDS", "CoauthoringFolder", "?") = "?" Then
        If MsgBox("You have not set a PaDS URL. Open Settings?", vbOKCancel) = vbOK Then
          Call Settings.ShowSettingsForm
          Exit Sub
        Else
          Exit Sub
        End If
    End If
        
    'Check for tabroom username/password
    If GetSetting("Verbatim", "Main", "TabroomUsername", "?") = "?" Or GetSetting("Verbatim", "Main", "TabroomPassword", "?") = "?" Then
        If MsgBox("You have not configured your tabroom username/password. Open Settings?", vbOKCancel) = vbOK Then
          Call Settings.ShowSettingsForm
          Exit Sub
        Else
          Exit Sub
        End If
    End If
    
    'Check that access for assistive devices is enabled
    If Left(System.Version, 4) = "10.9" Or Left(System.Version, 5) = "10.10" Then 'Mavericks and up - assistive devices permission has to be granted one application at a time
            If MacScript("tell application ""System Events"" to return UI elements enabled") = "false" Then
                If MsgBox("To open files from the Coauthor menu, you must add Word to the list of authorized programs in System Preferences - Security & Privacy - Accessibility. Open Now?", vbYesNo) = vbYes Then
                    Script = "tell application ""System Preferences""" & vbCrLf
                    Script = Script & "set securityPane to pane id ""com.apple.preference.security""" & vbCrLf
                    Script = Script & "tell securityPane to reveal anchor ""Privacy_Accessibility""" & vbCrLf
                    Script = Script & "activate" & vbCrLf
                    Script = Script & "end tell"
                    MacScript (Script)
                End If
                Exit Sub
            End If
    ElseIf Left(System.Version, 4) = "10.8" Or Left(System.Version, 4) = "10.7" Or Left(System.Version, 4) = "10.6" Then 'Mountain Lion and down
        If MacScript("tell application ""System Events"" to return UI elements enabled") = "false" Then
            If MsgBox("To open files from the Coauthor menu, you must first check ""Enable access for assistive devices"" in System Preferences - Accessibility. Open Now?", vbYesNo) = vbYes Then
                Script = "tell application ""System Preferences""" & vbCrLf
                Script = Script & "set the current pane to pane id ""com.apple.preference.universalaccess""" & vbCrLf
                Script = Script & "activate" & vbCrLf
                Script = Script & "end tell"
                MacScript (Script)
            End If
            Exit Sub
        End If
    End If
                
    'Get which control was pressed and open file or folder
    Set PressedControl = CommandBars.ActionControl
    If PressedControl.Tag = "OpenPaDSFolder" Then
        Settings.LaunchWebsite (GetSetting("Verbatim", "PaDS", "CoauthoringFolder", "?"))
    Else
        'Create URL and encode spaces
        URL = "http://" & GetSetting("Verbatim", "PaDS", "PaDSSiteName") & ".paperlessdebate.com" & PressedControl.Parameter
        URL = Replace(URL, " ", "%20")
        
        'Construct script - have to shell out to osascript because VBA MacScript crashes while GUI scripting
        Script = "do shell script ""osascript -e 'tell application \""System Events\""' " & _
        "-e 'tell process \""Microsoft Word\""' " & _
        "-e 'tell menu bar 1' " & _
        "-e 'tell menu bar item \""File\""' " & _
        "-e 'tell menu \""File\""' " & _
        "-e 'click menu item \""Open URL...\""' " & _
        "-e 'end tell' " & _
        "-e 'end tell' " & _
        "-e 'end tell' " & _
        "-e 'tell window \""Open Url\""' " & _
        "-e 'set value of combo box 1 to \""" & URL & "\""' " & _
        "-e 'end tell' " & _
        "-e 'click button \""Open\"" of window \""Open URL\""' " & _
        "-e 'end tell' " & _
        "-e 'end tell' " & _
        "> /dev/null 2>&1 &"""
        
        MacScript (Script)
        
    End If
            
    'Clean up
    Set PressedControl = Nothing
    
End Sub

Sub OpenFromPaDSDummy()
'Dummy macro to let keybindings work
    Call PaDS.OpenFromPaDS
End Sub

Sub UploadToPaDSDummy()
'Dummy macro to let keybindings work
    Call PaDS.UploadToPaDS
End Sub

'*************************************************************************************
'* AUTHORIZATION FUNCTIONS                                                                             *
'*************************************************************************************

Sub GetPaDSCookie(Site As String, Username As String, Password As String)

    Dim PaDSURL As String
    Dim Script As String
    Dim XML As String
    Dim XMLArray
    Dim ViewStateKey As String
    Dim EventValidationKey As String
    Dim CookiePOSIX As String
    
    'Turn on error checking
    On Error GoTo Handler
    
    'Retrieve login page with curl and filter with grep
    PaDSURL = "http://" & Site & ".paperlessdebate.com/_forms/default.aspx?ReturnUrl=%2f_layouts%2f15%2fAuthenticate.aspx%3fSource%3d%252F&Source=%2F"
    Script = "do shell script ""curl -sS '" & PaDSURL & "' | grep -e VIEWSTATE -e EVENTVALIDATION | grep -v VIEWSTATEGENERATOR"""
    
    'Get VIEWSTATE and EVENTVALIDATION keys
    XML = MacScript(Script)
    XMLArray = Split(XML, Chr(13))
    ViewStateKey = Mid(XMLArray(0), InStr(XMLArray(0), "value=""") + 7, InStr(XMLArray(0), """ />") - InStr(XMLArray(0), "value=""") - 7)
    EventValidationKey = Mid(XMLArray(1), InStr(XMLArray(1), "value=""") + 7, InStr(XMLArray(1), """ />") - InStr(XMLArray(1), "value=""") - 7)
    
    'URL Encode keys
    ViewStateKey = Replace(ViewStateKey, "/", "%2F")
    ViewStateKey = Replace(ViewStateKey, "+", "%2B")
    EventValidationKey = Replace(EventValidationKey, "/", "%2F")
    EventValidationKey = Replace(EventValidationKey, "+", "%2B")
    
    'Set cookie location and delete if it exists
    Call PaDS.ClearPaDSCookie
    CookiePOSIX = MacScript("return POSIX path of (path to temporary items from user domain) as string") & "PaDSCookie.txt"
    
    'Login and get cookie with curl
    Script = "do shell script ""curl -d '"
    Script = Script & "__VIEWSTATE=" & ViewStateKey
    Script = Script & "&__EVENTVALIDATION=" & EventValidationKey
    Script = Script & "&ctl00%24PlaceHolderMain%24signInControl%24UserName=" & Username
    Script = Script & "&ctl00%24PlaceHolderMain%24signInControl%24password=" & Password
    Script = Script & "&ctl00%24PlaceHolderMain%24signInControl%24login=Sign+In"
    Script = Script & "&ctl00%24PlaceHolderMain%24signInControl%24RememberMe=on"
    Script = Script & "' -c '" & CookiePOSIX & "' '" & PaDSURL & "'"""
    XML = MacScript(Script)

    Exit Sub
    
Handler:
        MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Sub ClearPaDSCookie()

    Dim CookiePath As String
    
    On Error Resume Next
    
    'Set cookie location and delete if it exists
    CookiePath = MacScript("return (path to temporary items from user domain) as string") & "PaDSCookie.txt"
    If MacScript("tell application ""Finder""" & Chr(13) & "exists file """ & CookiePath & """" & Chr(13) & "end tell") = "true" Then Kill CookiePath

End Sub

'*************************************************************************************
'* API FUNCTIONS                                                                                                   *
'*************************************************************************************

Function SharepointURLExists(URL As String, Optional Folder As Boolean) As Boolean

    Dim APIURL As String
    Dim CookiePath As String
    Dim CookiePOSIX As String
    
    Dim Script As String
    Dim XML As String
       
    'Construct API file URL
    APIURL = Left(URL, InStr(8, URL, "/"))
    If Folder = True Then
        APIURL = APIURL & "_api/web/GetFolderByServerRelativeUrl('"
    Else
        APIURL = APIURL & "_api/web/GetFileByServerRelativeUrl('"
    End If
    APIURL = APIURL & Mid(URL, InStr(8, URL, "/"))
    APIURL = APIURL & "')"
    APIURL = Replace(APIURL, " ", "%20")
    
    'Set cookie location
    CookiePath = MacScript("return (path to temporary items from user domain) as string") & "PaDSCookie.txt"
    CookiePOSIX = MacScript("return POSIX path of (path to temporary items from user domain) as string") & "PaDSCookie.txt"

    'If cookie doesn't exist, or is older than 5 minutes, get it
    If MacScript("tell application ""Finder""" & Chr(13) & "exists file """ & CookiePath & """" & Chr(13) & "end tell") = "false" Then
        Call PaDS.GetPaDSCookie(GetSetting("Verbatim", "PaDS", "PaDSSiteName", "?"), GetSetting("Verbatim", "Main", "TabroomUsername", "?"), XORDecryption(GetSetting("Verbatim", "Main", "TabroomPassword", "?")))
    ElseIf DateDiff("n", MacScript("do shell script ""stat -f %Sm -t '%m/%d/%Y %H:%M:%S' '" & CookiePOSIX & "'"""), Now) > 5 Then
        Call PaDS.GetPaDSCookie(GetSetting("Verbatim", "PaDS", "PaDSSiteName", "?"), GetSetting("Verbatim", "Main", "TabroomUsername", "?"), XORDecryption(GetSetting("Verbatim", "Main", "TabroomPassword", "?")))
    End If

    'Construct curl request and send
    Script = "do shell script ""curl --cookie '" & CookiePOSIX & "' \""" & APIURL & "\"""""
    XML = MacScript(Script)
    
    'If response includes an "m:error" element, it doesn't exist
    If InStr(XML, "m:error") > 0 Then
        SharepointURLExists = False
    Else
        SharepointURLExists = True
    End If

End Function

Function SharepointGetFolderContents(URL As String, SubFolders As Boolean) As Variant

    Dim FolderURL As String
    Dim CookiePath As String
    Dim CookiePOSIX As String
    
    Dim Script As String
    Dim XML As String
    Dim URLList As String
    Dim URLArray
    
    'Construct API folder URL
    FolderURL = Left(URL, InStr(8, URL, "/"))
    FolderURL = FolderURL & "_api/web/GetFolderByServerRelativeUrl('"
    FolderURL = FolderURL & Mid(URL, InStr(8, URL, "/"))
    If Right(FolderURL, 1) = "/" Then FolderURL = Left(FolderURL, Len(FolderURL) - 1)
    If SubFolders = True Then
        FolderURL = FolderURL & "')/Folders/"
    Else
        FolderURL = FolderURL & "')/Files/"
    End If
    FolderURL = Replace(FolderURL, " ", "%20")
    
    'Set cookie location
    CookiePath = MacScript("return (path to temporary items from user domain) as string") & "PaDSCookie.txt"
    CookiePOSIX = MacScript("return POSIX path of (path to temporary items from user domain) as string") & "PaDSCookie.txt"

    'If cookie doesn't exist, or is older than 5 minutes, get it
    If MacScript("tell application ""Finder""" & Chr(13) & "exists file """ & CookiePath & """" & Chr(13) & "end tell") = "false" Then
        Call PaDS.GetPaDSCookie(GetSetting("Verbatim", "PaDS", "PaDSSiteName", "?"), GetSetting("Verbatim", "Main", "TabroomUsername", "?"), XORDecryption(GetSetting("Verbatim", "Main", "TabroomPassword", "?")))
    ElseIf DateDiff("n", MacScript("do shell script ""stat -f %Sm -t '%m/%d/%Y %H:%M:%S' '" & CookiePOSIX & "'"""), Now) > 5 Then
        Call PaDS.GetPaDSCookie(GetSetting("Verbatim", "PaDS", "PaDSSiteName", "?"), GetSetting("Verbatim", "Main", "TabroomUsername", "?"), XORDecryption(GetSetting("Verbatim", "Main", "TabroomPassword", "?")))
    End If

    'Construct curl request and send
    Script = "do shell script ""curl --cookie '" & CookiePOSIX & "' \""" & FolderURL & "\"""""
    XML = MacScript(Script)
    
    'Parse out URL values
    Do While InStr(XML, "<d:ServerRelativeUrl>") > 0
        URLList = URLList & Mid(XML, InStr(XML, "<d:ServerRelativeUrl>") + 21, InStr(XML, "</d:ServerRelativeUrl") - InStr(XML, "<d:ServerRelativeUrl>") - 21) & ";"
        XML = Mid(XML, InStr(XML, "</d:ServerRelativeUrl>") + 22)
    Loop
    
    'Strip off trailing ; to prevent empty element in array
    If Right(URLList, 1) = ";" Then URLList = Left(URLList, Len(URLList) - 1)
    
    'Return as an array
    URLArray = Split(URLList, ";")
    SharepointGetFolderContents = URLArray
    
End Function
