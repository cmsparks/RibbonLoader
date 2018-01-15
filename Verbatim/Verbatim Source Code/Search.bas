Attribute VB_Name = "Search"
Option Explicit

Sub GetSearchResultsContent()

    Dim SearchText As String
    Dim SearchDir As String
    Dim SearchResults As String
    Dim SearchResultsArray
    Dim s
    
    Dim Menu As CommandBarControl
    Dim MenuItem As CommandBarButton
    Dim m As CommandBarControl

    'Get search text from toolbar
    SearchText = CommandBars.FindControl(Tag:="Search").Text

    'Clear search results
    Set Menu = CommandBars.FindControl(Tag:="SearchResults")
    For Each m In Menu.Controls
        m.Delete
    Next m

    'Set search location
    If GetSetting("Verbatim", "Paperless", "SearchDir", "?") <> "?" Then
        SearchDir = MacScript("return quoted form of POSIX path of """ & GetSetting("Verbatim", "Paperless", "SearchDir", "?") & """")
    Else
        SearchDir = MacScript("return quoted form of POSIX path of (path to documents folder)")
    End If

    'Get search results
    SearchResults = MacScript("do shell script ""mdfind -onlyin " & SearchDir & " \""" & "kMDItemFSName == '*.doc*' && kMDItemTextContent == '" & SearchText & "'c \"" | head -n 25""")
    
    'If nothing found, exit
    If SearchResults = "" Then
        Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
        MenuItem.Caption = "No results found"
        
        Set Menu = Nothing
        Set MenuItem = Nothing
        Exit Sub
    End If
    
    'Split search results into an array, and add a button for each result
    SearchResultsArray = Split(SearchResults, Chr(13))
    For Each s In SearchResultsArray
        Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
        MenuItem.Caption = Mid(s, InStrRev(s, "/") + 1)
        MenuItem.TooltipText = s
        MenuItem.DescriptionText = s
        MenuItem.Tag = s
        MenuItem.FaceId = 1544 'Blank page
        MenuItem.OnAction = "Search.OpenSearchResult"
    Next s
    
    'Set template as saved to avoid prompts
    ActiveDocument.AttachedTemplate.Saved = True
    
    'Clean up
    Set Menu = Nothing
    Set MenuItem = Nothing

    Exit Sub
    
Handler:
    Set Menu = Nothing
    Set MenuItem = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Sub OpenSearchResult()
    
    Dim PressedControl As CommandBarButton
    
    'Get filename from tag of most recently pressed control
    Set PressedControl = CommandBars.ActionControl

    'If blank control, exit
    If PressedControl.Tag = "" Then
        Set PressedControl = Nothing
        Exit Sub
    End If
    
    'Open the document
    Documents.Open MacScript("return POSIX file """ & PressedControl.Tag & """ as string")
    
    'Clean up
    Set PressedControl = Nothing
    
    Exit Sub

Handler:
    Set PressedControl = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub
