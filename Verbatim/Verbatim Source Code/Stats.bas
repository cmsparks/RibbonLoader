Attribute VB_Name = "Stats"
Option Explicit

Sub ShowStatsForm()

    Dim StatsForm As frmStats
    
    If Toolbar.InvisibilityToggle = True Then
        MsgBox "Stats form cannot be opened while in Invisibility Mode. Please turn off Invisibility Mode and try again."
        Exit Sub
    End If
    
    Set StatsForm = New frmStats
    StatsForm.Show
    
End Sub

