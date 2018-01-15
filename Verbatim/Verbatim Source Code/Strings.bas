Attribute VB_Name = "Strings"
Option Explicit

Public Function OnlyAlphaNumericChars(ByVal OrigString As String) As String
    Dim lLen As Long
    Dim sAns As String
    Dim lCtr As Long
    Dim sChar As String
    
    OrigString = Trim(OrigString)
    lLen = Len(OrigString)
    For lCtr = 1 To lLen
        sChar = Mid(OrigString, lCtr, 1)
        If IsAlphaNumeric(Mid(OrigString, lCtr, 1)) Then
            sAns = sAns & sChar
        End If
    Next
        
    OnlyAlphaNumericChars = sAns

End Function

Private Function IsAlphaNumeric(sChr As String) As Boolean
        IsAlphaNumeric = sChr Like "[0-9A-Za-z ]"
        'IsSafeChar = sChr Like "[,.!@$%^():;'""_+=0-9A-Za-z -]"
End Function

Public Function OnlySafeChars(ByVal OrigString As String) As String
    Dim lLen As Long
    Dim sAns As String
    Dim lCtr As Long
    Dim sChar As String
    
    OrigString = Trim(OrigString)
    lLen = Len(OrigString)
    For lCtr = 1 To lLen
        sChar = Mid(OrigString, lCtr, 1)
        If IsSafeChar(Mid(OrigString, lCtr, 1)) Then
            sAns = sAns & sChar
        End If
    Next
        
    OnlySafeChars = sAns

End Function

Private Function IsSafeChar(sChr As String) As Boolean
        IsSafeChar = sChr Like "[*0-9A-Za-z -]"
End Function

Public Function ScrubString(s As String) As String

    s = Replace(s, "&", "")
    s = Replace(s, "?", "")
    s = Replace(s, "%", "")
    s = Replace(s, "[", "")
    s = Replace(s, "]", "")
    s = Replace(s, "{", "")
    s = Replace(s, "}", "")
    s = Replace(s, "<", "")
    s = Replace(s, ">", "")
    s = Replace(s, "#", "")
    s = Replace(s, "(((", "~(~(~(")
    s = Replace(s, ")))", "~)~)~)")
    ScrubString = s

End Function

Public Function URLEncode(s As String, Optional SpaceAsPlus As Boolean = False) As String

  Dim StringLen As Long
  Dim Result
  StringLen = Len(s)

  If StringLen > 0 Then
    ReDim Result(StringLen) As String
    Dim i As Long, CharCode As Integer
    Dim Char As String, Space As String

    If SpaceAsPlus Then Space = "+" Else Space = "%20"

    For i = 1 To StringLen
      Char = Mid$(s, i, 1)
      CharCode = Asc(Char)
      Select Case CharCode
        Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
          Result(i) = Char
        Case 32
          Result(i) = Space
        Case 0 To 15
          Result(i) = "%0" & Hex(CharCode)
        Case Else
          Result(i) = "%" & Hex(CharCode)
      End Select
    Next i
    URLEncode = Join(Result, "")
  End If
End Function
