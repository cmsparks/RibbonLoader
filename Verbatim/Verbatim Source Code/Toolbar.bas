Attribute VB_Name = "Toolbar"
Option Explicit

'Togglebutton state variables
Public AutoOpenFolderToggle As Boolean
Public AutoCoauthoringToggle As Boolean
Public RecordAudioToggle As Boolean
Public InvisibilityToggle As Boolean
Public UnderlineModeToggle As Boolean

Public ProgressBar As String

#If MAC_OFFICE_VERSION >= 15 Then
    'Do nothing - can't build toolbars on Word 2016
#Else

Sub BuildVerbatimToolbar()

    Dim Toolbar As CommandBar
    Dim VerbatimToolbar As CommandBar
    Dim ButtonControl As CommandBarButton
    Dim MenuControl As CommandBarControl
    Dim PopupControl As CommandBarPopup
    
    On Error GoTo Handler
    
    CustomizationContext = ActiveDocument.AttachedTemplate
    
    'Delete any preexisting toolbars to start from scratch
    For Each Toolbar In Application.CommandBars
        If Toolbar.Name = "Verbatim" Then Toolbar.Delete
    Next Toolbar
    For Each Toolbar In Application.CommandBars
        If Toolbar.Name = "Verbatim2016" Then Toolbar.Delete
    Next Toolbar
    
    'Create Toolbar and make sure it's visible and first
    If GetSetting("Verbatim", "View", "ToolbarPosition", "Top") = "Top" Then
        Set VerbatimToolbar = CommandBars.Add(Name:="Verbatim", Position:=msoBarTop)
    Else
        Set VerbatimToolbar = CommandBars.Add(Name:="Verbatim", Position:=msoBarLeft)
    End If
    VerbatimToolbar.Visible = True
    VerbatimToolbar.RowIndex = 1
        
    'F2
    Set ButtonControl = VerbatimToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = GetSetting("Verbatim", "Keyboard", "F2Shortcut", "Paste")
    ButtonControl.FaceId = 3778 'F2
    ButtonControl.Style = msoButtonIconAndCaption
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "F2Button"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'F3
    Set ButtonControl = VerbatimToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = GetSetting("Verbatim", "Keyboard", "F3Shortcut", "Condense")
    ButtonControl.FaceId = 3779 'F3
    ButtonControl.Style = msoButtonIconAndCaption
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "F3Button"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'F4
    Set ButtonControl = VerbatimToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = GetSetting("Verbatim", "Keyboard", "F4Shortcut", "Pocket")
    ButtonControl.FaceId = 3780 'F4
    ButtonControl.Style = msoButtonIconAndCaption
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "F4Button"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'F5
    Set ButtonControl = VerbatimToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = GetSetting("Verbatim", "Keyboard", "F5Shortcut", "Hat")
    ButtonControl.FaceId = 3781 'F5
    ButtonControl.Style = msoButtonIconAndCaption
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "F5Button"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'F6
    Set ButtonControl = VerbatimToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = GetSetting("Verbatim", "Keyboard", "F6Shortcut", "Block")
    ButtonControl.FaceId = 3782 'F6
    ButtonControl.Style = msoButtonIconAndCaption
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "F6Button"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'F7
    Set ButtonControl = VerbatimToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = GetSetting("Verbatim", "Keyboard", "F7Shortcut", "Tag")
    ButtonControl.FaceId = 3783 'F7
    ButtonControl.Style = msoButtonIconAndCaption
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "F7Button"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'F8
    Set ButtonControl = VerbatimToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = GetSetting("Verbatim", "Keyboard", "F8Shortcut", "Cite")
    ButtonControl.FaceId = 3784 'F8
    ButtonControl.Style = msoButtonIconAndCaption
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "F8Button"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'F9
    Set ButtonControl = VerbatimToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = GetSetting("Verbatim", "Keyboard", "F9Shortcut", "Underline")
    ButtonControl.FaceId = 3785 'F9
    ButtonControl.Style = msoButtonIconAndCaption
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "F9Button"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'F10
    Set ButtonControl = VerbatimToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = GetSetting("Verbatim", "Keyboard", "F10Shortcut", "Emphasis")
    ButtonControl.FaceId = 3786 'F10
    ButtonControl.Style = msoButtonIconAndCaption
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "F10Button"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'F11
    Set ButtonControl = VerbatimToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = GetSetting("Verbatim", "Keyboard", "F11Shortcut", "Highlight")
    ButtonControl.FaceId = 3787 'F11
    ButtonControl.Style = msoButtonIconAndCaption
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "F11Button"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'F12
    Set ButtonControl = VerbatimToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = GetSetting("Verbatim", "Keyboard", "F12Shortcut", "Clear")
    ButtonControl.FaceId = 3788 'F12
    ButtonControl.Style = msoButtonIconAndCaption
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "F12Button"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'Shrink Text
    Set ButtonControl = VerbatimToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = "Shrink"
    ButtonControl.DescriptionText = "Shrink Text (Cmd+8 or Alt+F3)"
    ButtonControl.Style = msoButtonIconAndCaption
    ButtonControl.FaceId = 1845 '8 Ball
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "ShrinkText"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
                
    'Send To Speech
    Set ButtonControl = VerbatimToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = "Send To Speech"
    ButtonControl.DescriptionText = "Send To Speech (` key or Alt+Right)"
    ButtonControl.FaceId = 39 'Right Arrow
    ButtonControl.Style = msoButtonIcon
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "SendToSpeech"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"

    'Choose Doc
    Set ButtonControl = VerbatimToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = "Choose Doc"
    ButtonControl.DescriptionText = "Choose Speech Doc"
    ButtonControl.FaceId = 837 'Checkbox
    ButtonControl.Style = msoButtonIcon
    ButtonControl.Tag = "ChooseSpeechDoc"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"

    'Arrange Windows
    Set ButtonControl = VerbatimToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = "Arrange Windows"
    ButtonControl.DescriptionText = "Arrange Windows (Ctrl+Shift+Tab)"
    ButtonControl.FaceId = 53 'Cascading windows
    ButtonControl.Style = msoButtonIcon
    ButtonControl.Tag = "WindowArranger"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'New Speech Menu
    Set ButtonControl = VerbatimToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = "New Speech"
    ButtonControl.DescriptionText = "New Speech (Cmd+Shift+N)"
    ButtonControl.FaceId = 1717 'Page with arrow. 3813 = Page with plus
    ButtonControl.Style = msoButtonIcon
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "NewSpeech"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    Set MenuControl = VerbatimToolbar.Controls.Add(Type:=msoControlPopup)
    MenuControl.Caption = ""
    MenuControl.DescriptionText = "New Speech"
    MenuControl.Tag = "NewSpeechMenu"
    MenuControl.OnAction = "Toolbar.AssignButtonActions"
    
    'VTub
    Set ButtonControl = VerbatimToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = ""
    ButtonControl.DescriptionText = "Virtual Tub"
    ButtonControl.FaceId = 1399 'Empty Box
    ButtonControl.Style = msoButtonIcon
    ButtonControl.BeginGroup = True
    ButtonControl.Enabled = False
    ButtonControl.Tag = "VirtualTubFake"
    
    Set MenuControl = VerbatimToolbar.Controls.Add(Type:=msoControlPopup)
    MenuControl.Caption = "VTub"
    MenuControl.DescriptionText = "Virtual Tub"
    MenuControl.Tag = "VirtualTub"
    MenuControl.OnAction = "Toolbar.AssignButtonActions"
    
    'Search
    Set MenuControl = VerbatimToolbar.Controls.Add(Type:=msoControlEdit)
    MenuControl.Caption = ""
    MenuControl.DescriptionText = "Search"
    MenuControl.BeginGroup = True
    MenuControl.Tag = "Search"
    MenuControl.OnAction = "Toolbar.AssignButtonActions"
    
    Set MenuControl = VerbatimToolbar.Controls.Add(Type:=msoControlPopup)
    MenuControl.Caption = ""
    MenuControl.DescriptionText = "Search Results"
    MenuControl.Tag = "SearchResults"
    
    'Tools Menu
    Set ButtonControl = VerbatimToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = ""
    ButtonControl.DescriptionText = "Tools"
    ButtonControl.FaceId = 2933 'Tools
    ButtonControl.Style = msoButtonIcon
    ButtonControl.BeginGroup = True
    ButtonControl.Enabled = False
    ButtonControl.Tag = "ToolsFake"
    
    Set MenuControl = VerbatimToolbar.Controls.Add(Type:=msoControlPopup)
    MenuControl.Caption = "Tools"
    MenuControl.DescriptionText = "Tools"
    MenuControl.Tag = "Tools"
    
        'Timer
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Start Timer"
        ButtonControl.DescriptionText = "Start Timer (Cmd+Shift+T)"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 608 'Alarm clock
        ButtonControl.Tag = "StartTimer"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
        'Stats
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Document Stats"
        ButtonControl.DescriptionText = "Document Stats (Cmd+Shift+I)"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 50 'Calculator
        ButtonControl.Tag = "DocumentStats"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
        'Combine Docs
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Combine Docs"
        ButtonControl.DescriptionText = "Combine Docs"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 486 '2 boxes combining
        ButtonControl.Tag = "CombineDocs"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
        'Record Audio
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        If RecordAudioToggle = False Then
            ButtonControl.Caption = "Start Recording Audio"
        Else
            ButtonControl.Caption = "Stop Recording Audio"
        End If
        ButtonControl.DescriptionText = "Record Audio"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 27 'Microphone
        ButtonControl.Tag = "RecordAudio"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
        'Add Warrant
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton, ID:=1589)
        ButtonControl.Caption = "New Warrant"
        ButtonControl.DescriptionText = "New Warrant"
        
        'Delete Warrants
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Delete All Warrants"
        ButtonControl.DescriptionText = "Delete All Warrants"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 1592 'Delete Comment
        ButtonControl.Tag = "DeleteAllWarrants"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
        'Convert
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Convert Backfile"
        ButtonControl.DescriptionText = "Convert Backfile"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 3271 'Wand
        ButtonControl.Tag = "ConvertBackfile"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
            
        'Auto Open Folder
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        If AutoOpenFolderToggle = False Then
            ButtonControl.Caption = "Start Auto Open Folder"
        Else
            ButtonControl.Caption = "Stop Auto Open Folder"
        End If
        ButtonControl.DescriptionText = "Auto Open Folder"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 23 'Open folder
        ButtonControl.Tag = "AutoOpenFolder"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'Share Menu
    Set ButtonControl = VerbatimToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = ""
    ButtonControl.DescriptionText = "Share"
    ButtonControl.FaceId = 3975 'Share
    ButtonControl.Style = msoButtonIcon
    ButtonControl.BeginGroup = True
    ButtonControl.Enabled = False
    ButtonControl.Tag = "ShareFake"
    
    Set MenuControl = VerbatimToolbar.Controls.Add(Type:=msoControlPopup)
    MenuControl.Caption = "Share"
    MenuControl.DescriptionText = "Share"
    MenuControl.Tag = "Share"
    
        'Copy to USB
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Copy To USB"
        ButtonControl.DescriptionText = "Copy To USB (Cmd+Shift+S)"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 777 'Paperclip
        ButtonControl.Tag = "CopyToUSB"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
        'Email
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Send Email"
        ButtonControl.DescriptionText = "Send Email (Cmd+Shift+E)"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 1675 'Envelope
        ButtonControl.Tag = "SendEmail"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
        'PaDS Public
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "PaDS Public Folder"
        ButtonControl.DescriptionText = "Upload to PaDS Public Folder (Cmd+Shift+W)"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 3823 'Save to web
        ButtonControl.Tag = "PaDSPublic"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
    'View Menu
    Set ButtonControl = VerbatimToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = ""
    ButtonControl.DescriptionText = "View"
    ButtonControl.FaceId = 48 'Eyeglasses
    ButtonControl.Style = msoButtonIcon
    ButtonControl.BeginGroup = True
    ButtonControl.Enabled = False
    ButtonControl.Tag = "ViewFake"
    
    Set MenuControl = VerbatimToolbar.Controls.Add(Type:=msoControlPopup)
    MenuControl.Caption = "View"
    MenuControl.DescriptionText = "View"
    MenuControl.Tag = "View"
    
        'Default View
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Default View"
        ButtonControl.DescriptionText = "Default View"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 176 'Book side by side
        ButtonControl.Tag = "DefaultView"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
          
        'Reading View
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton, ID:=9377)
        ButtonControl.Caption = "Reading View"
        ButtonControl.DescriptionText = "Reading View"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 247
        
        'Doc Map
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton, ID:=1714)
      
        'Invisibility Mode
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        If InvisibilityToggle = False Then
            ButtonControl.Caption = "Turn On Invisibility Mode"
        Else
            ButtonControl.Caption = "Turn Off Invisibility Mode"
        End If
        ButtonControl.DescriptionText = "Toggle Invisibility Mode"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 2174 'Blue eye
        ButtonControl.Tag = "InvisibilityMode"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
      
        'Show Paragraphs
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton, ID:=119)
        ButtonControl.Caption = "Show Paragraph Formatting"
        ButtonControl.DescriptionText = "Show Paragraph Formatting"
        ButtonControl.Style = msoButtonIconAndCaption
        
        'Window List
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton, ID:=959)
        ButtonControl.FaceId = 303

    'Coauthor Menu
    Set ButtonControl = VerbatimToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = ""
    ButtonControl.DescriptionText = "Coauthor"
    ButtonControl.FaceId = 1756 'Heads
    ButtonControl.Style = msoButtonIcon
    ButtonControl.BeginGroup = True
    ButtonControl.Enabled = False
    ButtonControl.Tag = "CoauthorFake"
    
    Set MenuControl = VerbatimToolbar.Controls.Add(Type:=msoControlPopup)
    MenuControl.Caption = "Coauthor"
    MenuControl.DescriptionText = "Coauthor"
    MenuControl.Tag = "CoauthoringMenu"
    MenuControl.OnAction = "Toolbar.AssignButtonActions"
    
    'Format Menu
    Set ButtonControl = VerbatimToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = ""
    ButtonControl.DescriptionText = "Format"
    ButtonControl.FaceId = 254 'FormatStyle
    ButtonControl.Style = msoButtonIcon
    ButtonControl.BeginGroup = True
    ButtonControl.Enabled = False
    ButtonControl.Tag = "FormatFake"
    
    Set MenuControl = VerbatimToolbar.Controls.Add(Type:=msoControlPopup)
    MenuControl.Caption = "Format"
    MenuControl.DescriptionText = "Format"
    MenuControl.Tag = "Format"
        
        'Auto Underline
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Auto Underline"
        ButtonControl.DescriptionText = "Auto Underline (Alt+F9)"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 1709 'Zap red text
        ButtonControl.Tag = "AutoUnderline"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
        'Underline Mode
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        If UnderlineModeToggle = False Then
            ButtonControl.Caption = "Turn On Underline Mode"
        Else
            ButtonControl.Caption = "Turn Off Underline Mode"
        End If
        ButtonControl.DescriptionText = "Underline Mode"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 375 'Underscore
        ButtonControl.Tag = "UnderlineMode"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
        'Update Styles
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Update Styles"
        ButtonControl.DescriptionText = "Update Styles (Cmd+F12)"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 62 'Update Styles
        ButtonControl.Tag = "UpdateStyles"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
        'Select Similar
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Select Similar"
        ButtonControl.DescriptionText = "Select Similar (Cmd+F2)"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 8035 'AB to AB
        ButtonControl.Tag = "SelectSimilar"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
        'Shrink All
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Shrink All"
        ButtonControl.DescriptionText = "Shrink All"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 1845 '8 Ball
        ButtonControl.Tag = "ShrinkAll"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
        'Shrink Pilcrows
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Shrink Pilcrows"
        ButtonControl.DescriptionText = "Shrink Pilcrows"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 948 'Small Pilcrow
        ButtonControl.Tag = "ShrinkPilcrows"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
        'Remove Pilcrows
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Remove Pilcrows"
        ButtonControl.DescriptionText = "Remove Pilcrows"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 119 'Pilcrow
        ButtonControl.Tag = "RemovePilcrows"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
        'Remove Blanks
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Remove Blanks"
        ButtonControl.DescriptionText = "Remove Blanks"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 143 'Blank line between pages
        ButtonControl.Tag = "RemoveBlanks"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
        'Remove Hyperlinks
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Remove Hyperlinks"
        ButtonControl.DescriptionText = "Remove Hyperlinks"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 2309 'Break Link
        ButtonControl.Tag = "RemoveHyperlinks"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
        'Remove Bookmarks
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Remove Bookmarks"
        ButtonControl.DescriptionText = "Remove Bookmarks"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 2528 'Delete Flag
        ButtonControl.Tag = "RemoveBookmarks"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
        'Remove Emphasis
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Remove Emphasis"
        ButtonControl.DescriptionText = "Remove Emphasis (Cmd+Shift+F10)"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 6243 'Pencil no on page
        ButtonControl.Tag = "RemoveEmphasis"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
        'Auto Emphasize First
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Auto Emphasize First"
        ButtonControl.DescriptionText = "Auto Emphasize First Letters (Cmd+F10)"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 223 'Small a in box
        ButtonControl.Tag = "AutoEmphasizeFirst"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
        'Fix Fake Tags
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Fix Fake Tags"
        ButtonControl.DescriptionText = "Fix Fake Tags"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 16195 'Small a to A
        ButtonControl.Tag = "FixFakeTags"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
        'Unihighlight
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Standardize Highlighting"
        ButtonControl.DescriptionText = "Standardize Highlighting"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 661 'Dripping paint bucket
        ButtonControl.Tag = "UniHighlight"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
        'Insert Header
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Insert Header"
        ButtonControl.DescriptionText = "Insert Header"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 9118 'Header Strip
        ButtonControl.Tag = "InsertHeader"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"

        'Duplicate Cite
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Duplicate Cite"
        ButtonControl.DescriptionText = "Duplicate Cite (Alt+F8)"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 321 'Duplicate
        ButtonControl.Tag = "DuplicateCite"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
        'Auto Format Cite
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Auto Format Cite"
        ButtonControl.DescriptionText = "Auto Format Cite (Cmd+F8)"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 1979 'Check mark v card
        ButtonControl.Tag = "AutoFormatCite"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"

        'Reformat Cite Dates
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Reformat Cite Dates"
        ButtonControl.DescriptionText = "Reformat Cite Dates"
        ButtonControl.FaceId = 125 'Calendar
        ButtonControl.Tag = "ReformatCiteDates"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"

        'Auto Number Tags
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Auto Number Tags"
        ButtonControl.DescriptionText = "Auto Number Tags (Cmd+Alt+3)"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 127 'Number sign page
        ButtonControl.Tag = "AutoNumberTags"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"

        'DeNumber Tags
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "DeNumber Tags"
        ButtonControl.DescriptionText = "DeNumber Tags"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 3793 'Bullet list
        ButtonControl.Tag = "DeNumberTags"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
        'Get From CiteMaker
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Get From CiteMaker"
        ButtonControl.DescriptionText = "Get From CiteMaker (Alt+F2)"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 1015 'Download web page
        ButtonControl.Tag = "GetFromCiteMaker"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"

    'Caselist Menu
    Set ButtonControl = VerbatimToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = ""
    ButtonControl.DescriptionText = "Caselist"
    ButtonControl.FaceId = 3903 'Upload to web page
    ButtonControl.Style = msoButtonIcon
    ButtonControl.BeginGroup = True
    ButtonControl.Enabled = False
    ButtonControl.Tag = "CaselistFake"
    
    Set MenuControl = VerbatimToolbar.Controls.Add(Type:=msoControlPopup)
    MenuControl.Caption = "Caselist"
    MenuControl.DescriptionText = "Caselist"
    MenuControl.Tag = "Caselist"
    
        'Caselist Wizard
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Caselist Wizard"
        ButtonControl.DescriptionText = "Caselist Wizard"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 625 'Wand
        ButtonControl.Tag = "CaselistWizard"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
        'Wikify
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Wikify"
        ButtonControl.DescriptionText = "Wikify"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 567 'Word icon
        ButtonControl.Tag = "ConvertToWiki"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
        'Citeify
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Citeify"
        ButtonControl.DescriptionText = "Citeify"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 527 'Wand page
        ButtonControl.Tag = "CiteRequestDoc"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
        'Cite Request
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Cite Request Card"
        ButtonControl.DescriptionText = "Cite Request Card (Ctrl+Shift+Q)"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 325 'Condense
        ButtonControl.Tag = "CiteRequest"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
    'Settings Menu
    Set ButtonControl = VerbatimToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = ""
    ButtonControl.DescriptionText = "Settings"
    ButtonControl.FaceId = 2144 'Gears. 611 = Hammer
    ButtonControl.Style = msoButtonIcon
    ButtonControl.BeginGroup = True
    ButtonControl.Enabled = False
    ButtonControl.Tag = "SettingsFake"
    
    Set MenuControl = VerbatimToolbar.Controls.Add(Type:=msoControlPopup)
    MenuControl.Caption = "Settings"
    MenuControl.DescriptionText = "Settings"
    MenuControl.Tag = "SettingsMenu"
    
        'Verbatim Settings
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Verbatim Settings"
        ButtonControl.DescriptionText = "Verbatim Settings (Alt+F1)"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 2144 'Gears
        ButtonControl.Tag = "VerbatimSettings"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
        'Verbatim Help
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Verbatim Help"
        ButtonControl.DescriptionText = "Verbatim Help (F1)"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 1743 'Book
        ButtonControl.Tag = "VerbatimHelp"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
        'Cheat Sheet
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "Shortcut Cheat Sheet"
        ButtonControl.DescriptionText = "Shortcut Cheat Sheet"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 487 'Info
        ButtonControl.Tag = "CheatSheet"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
        'Launch Website
        Set ButtonControl = MenuControl.Controls.Add(Type:=msoControlButton)
        ButtonControl.Caption = "paperlessdebate.com"
        ButtonControl.DescriptionText = "Launch paperlessdebate.com"
        ButtonControl.Style = msoButtonIconAndCaption
        ButtonControl.FaceId = 610 'Globe
        ButtonControl.Tag = "LaunchWebsite"
        ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
    'Set position and width
    If GetSetting("Verbatim", "View", "ToolbarPosition", "Top") = "Top" Then
        VerbatimToolbar.Position = msoBarTop
        VerbatimToolbar.Width = 950 '950 comes out 901
    Else
        VerbatimToolbar.Position = msoBarLeft
        VerbatimToolbar.Width = 100 '100 comes out 93
    End If
    
    'Save template if editing it
    ActiveDocument.AttachedTemplate.Saved = True
    
    'Clean Up
    Set VerbatimToolbar = Nothing
    Set ButtonControl = Nothing
    Set MenuControl = Nothing
    Set PopupControl = Nothing
    
    Exit Sub
        
Handler:
    Set VerbatimToolbar = Nothing
    Set ButtonControl = Nothing
    Set MenuControl = Nothing
    Set PopupControl = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Sub BuildVerbatim2016Toolbar()
'Builds a pared down toolbar for Word 2016 with no menus
'Must be run from Word 2011, since Word 2016 can't build toolbars

    Dim Toolbar As CommandBar
    Dim Verbatim2016Toolbar As CommandBar
    Dim ButtonControl As CommandBarButton
    Dim MenuControl As CommandBarControl
    Dim PopupControl As CommandBarPopup
    
    On Error GoTo Handler
    
    CustomizationContext = ActiveDocument.AttachedTemplate
    
    'Delete any preexisting toolbars to start from scratch
    For Each Toolbar In Application.CommandBars
        If Toolbar.Name = "Verbatim" Then Toolbar.Delete
    Next Toolbar
    For Each Toolbar In Application.CommandBars
        If Toolbar.Name = "Verbatim2016" Then Toolbar.Delete
    Next Toolbar
    
    'Create Toolbar and make sure it's visible and first
    Set Verbatim2016Toolbar = CommandBars.Add(Name:="Verbatim2016", Position:=msoBarTop)
    Verbatim2016Toolbar.Visible = True
    Verbatim2016Toolbar.RowIndex = 1
    
    'Verbatimize
    Set ButtonControl = Verbatim2016Toolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = "Verbatimize"
    ButtonControl.DescriptionText = "Verbatimize Current Document"
    ButtonControl.FaceId = 938 'Pin/note icon
    ButtonControl.Style = msoButtonIcon
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "Verbatimize"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'Send To Speech
    Set ButtonControl = Verbatim2016Toolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = "Send To Speech"
    ButtonControl.DescriptionText = "Send To Speech (` key or Alt+Right)"
    ButtonControl.FaceId = 39 'Right Arrow
    ButtonControl.Style = msoButtonIcon
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "SendToSpeech"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'Choose Doc
    Set ButtonControl = Verbatim2016Toolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = "Choose Doc"
    ButtonControl.DescriptionText = "Choose Speech Doc"
    ButtonControl.FaceId = 837 'Checkbox
    ButtonControl.Style = msoButtonIcon
    ButtonControl.Tag = "ChooseSpeechDoc"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"

    'Arrange Windows
    Set ButtonControl = Verbatim2016Toolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = "Arrange Windows"
    ButtonControl.DescriptionText = "Arrange Windows (Ctrl+Shift+Tab)"
    ButtonControl.FaceId = 53 'Cascading windows
    ButtonControl.Style = msoButtonIcon
    ButtonControl.Tag = "WindowArranger"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'New Speech
    Set ButtonControl = Verbatim2016Toolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = "New Speech"
    ButtonControl.DescriptionText = "New Speech (Cmd+Shift+N)"
    ButtonControl.FaceId = 1717 'Page with arrow. 3813 = Page with plus
    ButtonControl.Style = msoButtonIcon
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "NewSpeech"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'New Speech Menu
    Set ButtonControl = Verbatim2016Toolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = "v"
    ButtonControl.DescriptionText = "New Speech Menu"
    ButtonControl.Style = msoButtonCaption
    ButtonControl.Tag = "2016NewSpeechMenu"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'New Doc
    Set ButtonControl = Verbatim2016Toolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = "New Document"
    ButtonControl.DescriptionText = "New Verbatim Document"
    ButtonControl.FaceId = 1544 'Blank Page
    ButtonControl.Style = msoButtonIcon
    ButtonControl.Tag = "NewDocument"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'F2
    Set ButtonControl = Verbatim2016Toolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = GetSetting("Verbatim", "Keyboard", "F2Shortcut", "Paste")
    ButtonControl.FaceId = 3778 'F2
    ButtonControl.Style = msoButtonIconAndCaption
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "F2Button"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'F3
    Set ButtonControl = Verbatim2016Toolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = GetSetting("Verbatim", "Keyboard", "F3Shortcut", "Condense")
    ButtonControl.FaceId = 3779 'F3
    ButtonControl.Style = msoButtonIconAndCaption
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "F3Button"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'F4
    Set ButtonControl = Verbatim2016Toolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = GetSetting("Verbatim", "Keyboard", "F4Shortcut", "Pocket")
    ButtonControl.FaceId = 3780 'F4
    ButtonControl.Style = msoButtonIconAndCaption
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "F4Button"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'F5
    Set ButtonControl = Verbatim2016Toolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = GetSetting("Verbatim", "Keyboard", "F5Shortcut", "Hat")
    ButtonControl.FaceId = 3781 'F5
    ButtonControl.Style = msoButtonIconAndCaption
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "F5Button"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'F6
    Set ButtonControl = Verbatim2016Toolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = GetSetting("Verbatim", "Keyboard", "F6Shortcut", "Block")
    ButtonControl.FaceId = 3782 'F6
    ButtonControl.Style = msoButtonIconAndCaption
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "F6Button"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'F7
    Set ButtonControl = Verbatim2016Toolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = GetSetting("Verbatim", "Keyboard", "F7Shortcut", "Tag")
    ButtonControl.FaceId = 3783 'F7
    ButtonControl.Style = msoButtonIconAndCaption
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "F7Button"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'F8
    Set ButtonControl = Verbatim2016Toolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = GetSetting("Verbatim", "Keyboard", "F8Shortcut", "Cite")
    ButtonControl.FaceId = 3784 'F8
    ButtonControl.Style = msoButtonIconAndCaption
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "F8Button"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'F9
    Set ButtonControl = Verbatim2016Toolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = GetSetting("Verbatim", "Keyboard", "F9Shortcut", "Underline")
    ButtonControl.FaceId = 3785 'F9
    ButtonControl.Style = msoButtonIconAndCaption
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "F9Button"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'F10
    Set ButtonControl = Verbatim2016Toolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = GetSetting("Verbatim", "Keyboard", "F10Shortcut", "Emphasis")
    ButtonControl.FaceId = 3786 'F10
    ButtonControl.Style = msoButtonIconAndCaption
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "F10Button"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'F11
    Set ButtonControl = Verbatim2016Toolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = GetSetting("Verbatim", "Keyboard", "F11Shortcut", "Highlight")
    ButtonControl.FaceId = 3787 'F11
    ButtonControl.Style = msoButtonIconAndCaption
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "F11Button"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'F12
    Set ButtonControl = Verbatim2016Toolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = GetSetting("Verbatim", "Keyboard", "F12Shortcut", "Clear")
    ButtonControl.FaceId = 3788 'F12
    ButtonControl.Style = msoButtonIconAndCaption
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "F12Button"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'Shrink Text
    Set ButtonControl = Verbatim2016Toolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = "Shrink"
    ButtonControl.DescriptionText = "Shrink Text (Cmd+8 or Alt+F3)"
    ButtonControl.Style = msoButtonIconAndCaption
    ButtonControl.FaceId = 1845 '8 Ball
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "ShrinkText"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
                
    'Tools Menu
    Set ButtonControl = Verbatim2016Toolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = "Tools"
    ButtonControl.DescriptionText = "Tools"
    ButtonControl.TooltipText = "Tools Menu"
    ButtonControl.FaceId = 2933 'Tools
    ButtonControl.Style = msoButtonIcon
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "2016ToolsMenu"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'Share Menu
    Set ButtonControl = Verbatim2016Toolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = "Share"
    ButtonControl.DescriptionText = "Share"
    ButtonControl.TooltipText = "Share Menu"
    ButtonControl.FaceId = 3975 'Share
    ButtonControl.Style = msoButtonIcon
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "2016ShareMenu"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'View Menu
    Set ButtonControl = Verbatim2016Toolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = "View"
    ButtonControl.DescriptionText = "View"
    ButtonControl.TooltipText = "View Menu"
    ButtonControl.FaceId = 48 'Eyeglasses
    ButtonControl.Style = msoButtonIcon
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "2016ViewMenu"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
    'Format Menu
    Set ButtonControl = Verbatim2016Toolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = "Format"
    ButtonControl.DescriptionText = "Format"
    ButtonControl.TooltipText = "Format Menu"
    ButtonControl.FaceId = 254 'FormatStyle
    ButtonControl.Style = msoButtonIcon
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "2016FormatMenu"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
       
    'Caselist Menu
    Set ButtonControl = Verbatim2016Toolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = "Caselist"
    ButtonControl.DescriptionText = "Caselist"
    ButtonControl.TooltipText = "Caselist Menu"
    ButtonControl.FaceId = 3903 'Upload to web page
    ButtonControl.Style = msoButtonIcon
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "2016CaselistMenu"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
    
    'Settings Menu
    Set ButtonControl = Verbatim2016Toolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = "Settings"
    ButtonControl.DescriptionText = "Settings"
    ButtonControl.TooltipText = "Settings Menu"
    ButtonControl.FaceId = 2144 'Gears. 611 = Hammer
    ButtonControl.Style = msoButtonIcon
    ButtonControl.BeginGroup = True
    ButtonControl.Tag = "2016SettingsMenu"
    ButtonControl.OnAction = "Toolbar.AssignButtonActions"
        
    'Save template if editing it
    ActiveDocument.AttachedTemplate.Saved = True
    
    'Clean Up
    Set Verbatim2016Toolbar = Nothing
    Set ButtonControl = Nothing
    
    Exit Sub
        
Handler:
    Set Verbatim2016Toolbar = Nothing
    Set ButtonControl = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

#End If 'End of 2016 Conditional

Sub AssignButtonActions()
   
    Dim PressedControl As CommandBarControl
    
    'Get pressed control
    Set PressedControl = CommandBars.ActionControl
    If PressedControl Is Nothing Then Exit Sub
    
    CustomizationContext = ActiveDocument.AttachedTemplate
    
    'Choose the action based on the control's tag
    Select Case PressedControl.Tag
    
        'F Keys
        Case Is = "F2Button"
            FindKey(wdKeyF2).Execute
        Case Is = "F3Button"
            FindKey(wdKeyF3).Execute
        Case Is = "F4Button"
            FindKey(wdKeyF4).Execute
        Case Is = "F5Button"
            FindKey(wdKeyF5).Execute
        Case Is = "F6Button"
            FindKey(wdKeyF6).Execute
        Case Is = "F7Button"
            FindKey(wdKeyF7).Execute
        Case Is = "F8Button"
            FindKey(wdKeyF8).Execute
        Case Is = "F9Button"
            FindKey(wdKeyF9).Execute
        Case Is = "F10Button"
            FindKey(wdKeyF10).Execute
        Case Is = "F11Button"
            FindKey(wdKeyF11).Execute
        Case Is = "F12Button"
            FindKey(wdKeyF12).Execute
    
        'Paperless Group
        Case Is = "SendToSpeech"
            Paperless.SendToSpeech
        Case Is = "ChooseSpeechDoc"
            Paperless.ShowChooseSpeechDoc
        Case Is = "WindowArranger"
            View.ArrangeWindows
        
        Case Is = "NewSpeechMenu"
            Paperless.GetSpeeches
        Case Is = "NewSpeech", "NewSpeech1"
            Paperless.NewSpeech
        Case Is = "NewDocument"
            Paperless.NewDocument
        Case Is = "RefreshSpeeches"
            Call Paperless.GetSpeeches(FromScratch:=True)
             
        'Virtual Tub
        Case Is = "VirtualTub"
            VirtualTub.GetVTubContent
        Case Is = "RefreshVTub"
            VirtualTub.VTubRefresh
        
        'Search
        Case Is = "Search"
            Search.GetSearchResultsContent
        
        'Share Group
        Case Is = "CopyToUSB"
            Paperless.CopyToUSB
        Case Is = "SendEmail"
            Email.ShowEmailForm
        Case Is = "PaDSPublic"
            Call PaDS.UploadToPaDS(UploadToPublic:=True)
        
        'Tools Group
        Case Is = "StartTimer"
            Paperless.StartTimer
        Case Is = "DocumentStats"
            Stats.ShowStatsForm
        Case Is = "RecordAudio"
            If RecordAudioToggle = False Then
                RecordAudioToggle = True
                PressedControl.Caption = "Stop Audio Recording"
                Call Audio.StartRecord
            Else
                RecordAudioToggle = False
                PressedControl.Caption = "Start Audio Recording"
                Call Audio.SaveRecord
            End If
        Case Is = "CombineDocs"
            Caselist.ShowCombineDocs
        Case Is = "NewWarrant"
            Paperless.NewWarrant
        Case Is = "DeleteAllWarrants"
            Paperless.DeleteAllWarrants
        Case Is = "AutoOpenFolder"
            Call Paperless.AutoOpenFolder
        
        'View Group
        Case Is = "DefaultView"
            View.DefaultView
                    
        Case Is = "InvisibilityMode"
            If InvisibilityToggle = False Then
                InvisibilityToggle = True
                Call View.InvisibilityOn
                PressedControl.Caption = "Turn Off Invisibility Mode"
                MsgBox "Invisibility Mode Turned ON. Press the menu option again to turn off."
            Else
                InvisibilityToggle = False
                Call View.InvisibilityOff
                PressedControl.Caption = "Turn On Invisibility Mode"
            End If
            
        'PaDS Group
        Case Is = "CoauthoringMenu"
            PaDS.GetCoauthoringContent
        Case Is = "RefreshCoauthoring"
            Call PaDS.GetCoauthoringContent(FromScratch:=True)
                    
        'Format Group
        Case Is = "ShrinkText"
            Formatting.ShrinkText
        Case Is = "ConvertBackfile"
            Convert.ShowConvertForm
        Case Is = "AutoUnderline"
            Formatting.AutoUnderline
        Case Is = "UnderlineMode"
                Call Formatting.UnderlineMode
        Case Is = "UpdateStyles"
            Formatting.UpdateStyles
        Case Is = "SelectSimilar"
            Formatting.SelectSimilar
        Case Is = "ShrinkAll"
            Formatting.ShrinkAll
        Case Is = "ShrinkPilcrows"
            Formatting.ShrinkPilcrows
        Case Is = "RemovePilcrows"
            Formatting.RemovePilcrows
        Case Is = "RemoveBlanks"
            Formatting.RemoveBlanks
        Case Is = "RemoveHyperlinks"
            Formatting.RemoveHyperlinks
        Case Is = "RemoveBookmarks"
            VirtualTub.RemoveBookmarks
        Case Is = "RemoveEmphasis"
            Formatting.RemoveEmphasis
        Case Is = "AutoEmphasizeFirst"
            Formatting.AutoEmphasizeFirst
        Case Is = "FixFakeTags"
            Formatting.FixFakeTags
        Case Is = "UniHighlight"
            Formatting.UniHighlight
        Case Is = "InsertHeader"
            Formatting.InsertHeader
    
        Case Is = "DuplicateCite"
            Formatting.CopyPreviousCite
        Case Is = "AutoFormatCite"
            Formatting.AutoFormatCite
        Case Is = "ReformatCiteDates"
            Formatting.ReformatCiteDates
        Case Is = "AutoNumberTags"
            Formatting.AutoNumberTags
        Case Is = "DeNumberTags"
            Formatting.DeNumberTags
        Case Is = "GetFromCiteMaker"
            Formatting.GetFromCiteMaker
            
        'Caselist Group
        Case Is = "CaselistWizard"
            Caselist.ShowCaselistWizard
        Case Is = "ConvertToWiki"
            Caselist.Word2XWikiCites
        Case Is = "CiteRequestDoc"
            Caselist.CiteRequestDoc
        Case Is = "CiteRequest"
            Caselist.CiteRequest
        
        'Settings Group
        Case Is = "LaunchWebsite"
            Settings.LaunchWebsite ("http://paperlessdebate.com")
        Case Is = "VerbatimHelp"
            Settings.ShowVerbatimHelp
        Case Is = "CheatSheet"
            Settings.ShowCheatSheet
        Case Is = "VerbatimSettings"
            Settings.ShowSettingsForm
        
        '2016 Menu Items
        Case Is = "Verbatimize"
            AttachVerbatim.AttachVerbatim
        Case Is = "2016NewSpeechMenu"
            Toolbar.Show2016NewSpeechMenu
        Case Is = "2016ToolsMenu"
            Toolbar.Show2016ToolsMenu
        Case Is = "2016ShareMenu"
            Toolbar.Show2016ShareMenu
        Case Is = "2016ViewMenu"
            Toolbar.Show2016ViewMenu
        Case Is = "2016FormatMenu"
            Toolbar.Show2016FormatMenu
        Case Is = "2016CaselistMenu"
            Toolbar.Show2016CaselistMenu
        Case Is = "2016SettingsMenu"
            Settings.ShowVerbatimHelp
            
        Case Else
            'Do Nothing

    End Select

    'Clean Up
    Set PressedControl = Nothing
    
    Exit Sub

Handler:
    Set PressedControl = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

'2016 Menu Forms
Sub Show2016NewSpeechMenu()
    Dim NewSpeechForm As frm2016NewSpeech
    Set NewSpeechForm = New frm2016NewSpeech
    NewSpeechForm.Show
End Sub
Sub Show2016ToolsMenu()
    Dim ToolsForm As frm2016Tools
    Set ToolsForm = New frm2016Tools
    ToolsForm.Show
End Sub
Sub Show2016ShareMenu()
    Dim ShareForm As frm2016Share
    Set ShareForm = New frm2016Share
    ShareForm.Show
End Sub
Sub Show2016ViewMenu()
    Dim ViewForm As frm2016View
    Set ViewForm = New frm2016View
    ViewForm.Show
End Sub
Sub Show2016FormatMenu()
    Dim FormatForm As frm2016Format
    Set FormatForm = New frm2016Format
    FormatForm.Show
End Sub
Sub Show2016CaselistMenu()
    Dim CaselistForm As frm2016Caselist
    Set CaselistForm = New frm2016Caselist
    CaselistForm.Show
End Sub

