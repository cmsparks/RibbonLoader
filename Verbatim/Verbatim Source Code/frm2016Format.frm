Attribute VB_Name = "frm2016Format"
Attribute VB_Base = "0{E23178BE-7137-4C7A-95FF-F23B35F239E5}{F41FD499-63F7-444C-94F6-776709799227}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub UserForm_Initialize()
    
    On Error GoTo Handler
    
    Me.lboxFunctions.AddItem "Shrink Text"
    Me.lboxFunctions.AddItem "Convert Backfile"
    Me.lboxFunctions.AddItem "Auto Underline"
        
    If UnderlineModeToggle = False Then
        Me.lboxFunctions.AddItem "Turn ON Underline Mode"
    Else
        Me.lboxFunctions.AddItem "Turn OFF Underline Mode"
    End If
    
    Me.lboxFunctions.AddItem "Update Styles"
    Me.lboxFunctions.AddItem "Select Similar"
    Me.lboxFunctions.AddItem "Shrink All"
    Me.lboxFunctions.AddItem "Shrink Pilcrows"
    Me.lboxFunctions.AddItem "Remove Pilcrows"
    Me.lboxFunctions.AddItem "Remove Blanks"
    Me.lboxFunctions.AddItem "Remove Hyperlinks"
    Me.lboxFunctions.AddItem "Remove Bookmarks"
    Me.lboxFunctions.AddItem "Remove Emphasis"
    Me.lboxFunctions.AddItem "Auto Emphasize First"
    Me.lboxFunctions.AddItem "Fix Fake Tags"
    Me.lboxFunctions.AddItem "UniHighlight"
    Me.lboxFunctions.AddItem "Insert Header"
    Me.lboxFunctions.AddItem "Duplicate Cite"
    Me.lboxFunctions.AddItem "Auto Format Cite"
    Me.lboxFunctions.AddItem "Reformat Cite Dates"
    Me.lboxFunctions.AddItem "Auto Number Tags"
    Me.lboxFunctions.AddItem "De-Number Tags"
    Me.lboxFunctions.AddItem "Get From CiteMaker"
    
    Exit Sub
    
Handler:
        MsgBox "Error " & Err.Number & ": " & Err.Description
    
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnRun_Click()

    Dim i As Integer
    Dim FunctionSelected As Boolean

    'Make sure a function is selected
    For i = 0 To Me.lboxFunctions.ListCount - 1
        If Me.lboxFunctions.Selected(i) = True Then FunctionSelected = True
    Next i

    If FunctionSelected = False Then
        MsgBox "You must select a function first."
    Else
        
        'Unload the menu and run the function
        Select Case Me.lboxFunctions.Value
        
            Case Is = "Shrink Text"
                Unload Me
                Formatting.ShrinkText
            Case Is = "Auto Underline"
                Unload Me
                Formatting.AutoUnderline
            Case Is = "Turn ON Underline Mode"
                Unload Me
                Formatting.UnderlineMode
            Case Is = "Turn OFF Underline Mode"
                Unload Me
                Formatting.UnderlineMode
            Case Is = "Update Styles"
                Unload Me
                Formatting.UpdateStyles
            Case Is = "Select Similar"
                Unload Me
                Formatting.SelectSimilar
            Case Is = "Shrink All"
                Unload Me
                Formatting.ShrinkAll
            Case Is = "Shrink Pilcrows"
                Unload Me
                Formatting.ShrinkPilcrows
            Case Is = "Remove Pilcrows"
                Unload Me
                Formatting.RemovePilcrows
            Case Is = "Remove Blanks"
                Unload Me
                Formatting.RemoveBlanks
            Case Is = "Remove Hyperlinks"
                Unload Me
                Formatting.RemoveHyperlinks
            Case Is = "Remove Bookmarks"
                Unload Me
                VirtualTub.RemoveBookmarks
            Case Is = "Remove Emphasis"
                Unload Me
                Formatting.RemoveEmphasis
            Case Is = "Auto Emphasize First"
                Unload Me
                Formatting.AutoEmphasizeFirst
            Case Is = "Fix Fake Tags"
                Unload Me
                Formatting.FixFakeTags
            Case Is = "UniHighlight"
                Unload Me
                Formatting.UniHighlight
            Case Is = "Insert Header"
                Unload Me
                Formatting.InsertHeader
            Case Is = "Duplicate Cite"
                Unload Me
                Formatting.CopyPreviousCite
            Case Is = "Auto Format Cite"
                Unload Me
                Formatting.AutoFormatCite
            Case Is = "Reformat Cite Dates"
                Unload Me
                Formatting.ReformatCiteDates
            Case Is = "Auto Number Tags"
                Unload Me
                Formatting.AutoNumberTags
            Case Is = "De-Number Tags"
                Unload Me
                Formatting.DeNumberTags
            Case Is = "Get From CiteMaker"
                Unload Me
                Formatting.GetFromCiteMaker
                
            Case Else
                Unload Me
                'Do Nothing
            
        End Select
    
    End If

End Sub
