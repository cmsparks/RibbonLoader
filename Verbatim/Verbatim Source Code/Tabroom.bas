Attribute VB_Name = "Tabroom"
Option Explicit

Public Function GetTabroomRounds(Optional Email As Boolean) As Variant

    Dim Script As String
    Dim XML As String
    Dim XMLArray
    Dim RoundArray() As String
    Dim i
    
    Dim ParseText As String
    
    'Turn off error checking in case no internet
    On Error Resume Next
    
    'Construct curl request and send
    Script = "do shell script ""curl -d 'username="
    Script = Script & GetSetting("Verbatim", "Main", "TabroomUsername", "?") & "&password="
    Script = Script & XORDecryption(GetSetting("Verbatim", "Main", "TabroomPassword", "?"))
    If Email = True Then Script = Script & "&email=1"
    Script = Script & "' 'https://www.tabroom.com/api/verbatim.mhtml'"""
    XML = MacScript(Script)
    
    'Chunk the XML by ROUND tag into an array
    XMLArray = Split(XML, "<ROUND>")

    If Email = True Then
    
        'Resize the 2d round array to one less than the XML array
        ReDim RoundArray(0 To UBound(XMLArray) - 1, 0 To 8) As String
    
        'Parse each ROUND in XML array, ignoring zero-indexed element, which is XML boilerplate
        For i = 1 To UBound(XMLArray)
            RoundArray(i - 1, 0) = Mid(XMLArray(i), InStr(XMLArray(i), "<TOURNAMENT>") + 12, InStr(XMLArray(i), "</TOURNAMENT>") - InStr(XMLArray(i), "<TOURNAMENT>") - 12)
            RoundArray(i - 1, 1) = Mid(XMLArray(i), InStr(XMLArray(i), "<ROUND_NAME>") + 12, InStr(XMLArray(i), "</ROUND_NAME>") - InStr(XMLArray(i), "<ROUND_NAME>") - 12)
            RoundArray(i - 1, 2) = Mid(XMLArray(i), InStr(XMLArray(i), "<SIDE>") + 6, InStr(XMLArray(i), "</SIDE>") - InStr(XMLArray(i), "<SIDE>") - 6)
            RoundArray(i - 1, 3) = Mid(XMLArray(i), InStr(XMLArray(i), "<OPPONENT>") + 10, InStr(XMLArray(i), "</OPPONENT>") - InStr(XMLArray(i), "<OPPONENT>") - 10)
            RoundArray(i - 1, 4) = Mid(XMLArray(i), InStr(XMLArray(i), "<JUDGE>") + 7, InStr(XMLArray(i), "</JUDGE>") - InStr(XMLArray(i), "<JUDGE>") - 7)
            
            'Copy current round text for parsing names and emails
            ParseText = XMLArray(i)
            Do While InStr(ParseText, "<STUDENT_NAME>") > 0
                If Len(RoundArray(i - 1, 5)) > 0 Then RoundArray(i - 1, 5) = RoundArray(i - 1, 5) & ";"
                RoundArray(i - 1, 5) = RoundArray(i - 1, 5) & Mid(ParseText, InStr(ParseText, "<STUDENT_NAME>") + 14, InStr(ParseText, "</STUDENT_NAME") - InStr(ParseText, "<STUDENT_NAME>") - 14)
                ParseText = Mid(ParseText, InStr(ParseText, "</STUDENT_NAME>") + 15)
            Loop
            If Right(RoundArray(i - 1, 5), 1) = ";" Then RoundArray(i - 1, 5) = Left(RoundArray(i - 1, 5), Len(RoundArray(i - 1, 5)) - 1)
            
            ParseText = XMLArray(i)
            Do While InStr(ParseText, "<STUDENT_EMAIL>") > 0
                If Len(RoundArray(i - 1, 6)) > 0 Then RoundArray(i - 1, 6) = RoundArray(i - 1, 6) & ";"
                RoundArray(i - 1, 6) = RoundArray(i - 1, 6) & Mid(ParseText, InStr(ParseText, "<STUDENT_EMAIL>") + 15, InStr(ParseText, "</STUDENT_EMAIL") - InStr(ParseText, "<STUDENT_EMAIL>") - 15)
                ParseText = Mid(ParseText, InStr(ParseText, "</STUDENT_EMAIL>") + 16)
            Loop
            If Right(RoundArray(i - 1, 6), 1) = ";" Then RoundArray(i - 1, 6) = Left(RoundArray(i - 1, 6), Len(RoundArray(i - 1, 6)) - 1)
            
            ParseText = XMLArray(i)
            Do While InStr(ParseText, "<JUDGE_NAME>") > 0
                If Len(RoundArray(i - 1, 7)) > 0 Then RoundArray(i - 1, 7) = RoundArray(i - 1, 7) & ";"
                RoundArray(i - 1, 7) = RoundArray(i - 1, 7) & Mid(ParseText, InStr(ParseText, "<JUDGE_NAME>") + 12, InStr(ParseText, "</JUDGE_NAME") - InStr(ParseText, "<JUDGE_NAME>") - 12)
                ParseText = Mid(ParseText, InStr(ParseText, "</JUDGE_NAME>") + 13)
            Loop
            If Right(RoundArray(i - 1, 7), 1) = ";" Then RoundArray(i - 1, 7) = Left(RoundArray(i - 1, 7), Len(RoundArray(i - 1, 7)) - 1)
            
            ParseText = XMLArray(i)
            Do While InStr(ParseText, "<JUDGE_EMAIL>") > 0
                If Len(RoundArray(i - 1, 8)) > 0 Then RoundArray(i - 1, 8) = RoundArray(i - 1, 8) & ";"
                RoundArray(i - 1, 8) = RoundArray(i - 1, 8) & Mid(ParseText, InStr(ParseText, "<JUDGE_EMAIL>") + 13, InStr(ParseText, "</JUDGE_EMAIL") - InStr(ParseText, "<JUDGE_EMAIL>") - 13)
                ParseText = Mid(ParseText, InStr(ParseText, "</JUDGE_EMAIL>") + 14)
            Loop
            If Right(RoundArray(i - 1, 8), 1) = ";" Then RoundArray(i - 1, 8) = Left(RoundArray(i - 1, 8), Len(RoundArray(i - 1, 8)) - 1)
            
        Next i
    
    Else
        'Resize the 2d round array to one less than the XML array
        ReDim RoundArray(0 To UBound(XMLArray) - 1, 0 To 4) As String
    
        'Parse each ROUND in XML array, ignoring zero-indexed element, which is XML boilerplate
        For i = 1 To UBound(XMLArray)
            RoundArray(i - 1, 0) = Mid(XMLArray(i), InStr(XMLArray(i), "<TOURNAMENT>") + 12, InStr(XMLArray(i), "</TOURNAMENT>") - InStr(XMLArray(i), "<TOURNAMENT>") - 12)
            RoundArray(i - 1, 1) = Mid(XMLArray(i), InStr(XMLArray(i), "<ROUND_NAME>") + 12, InStr(XMLArray(i), "</ROUND_NAME>") - InStr(XMLArray(i), "<ROUND_NAME>") - 12)
            RoundArray(i - 1, 2) = Mid(XMLArray(i), InStr(XMLArray(i), "<SIDE>") + 6, InStr(XMLArray(i), "</SIDE>") - InStr(XMLArray(i), "<SIDE>") - 6)
            RoundArray(i - 1, 3) = Mid(XMLArray(i), InStr(XMLArray(i), "<OPPONENT>") + 10, InStr(XMLArray(i), "</OPPONENT>") - InStr(XMLArray(i), "<OPPONENT>") - 10)
            RoundArray(i - 1, 4) = Mid(XMLArray(i), InStr(XMLArray(i), "<JUDGE>") + 7, InStr(XMLArray(i), "</JUDGE>") - InStr(XMLArray(i), "<JUDGE>") - 7)
        Next i

    End If
    
    'Return the round array or an empty string - double Not checks if array is empty
    If Not Not RoundArray Then
        GetTabroomRounds = RoundArray
    Else
        GetTabroomRounds = ""
    End If
    
End Function
