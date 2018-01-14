Attribute VB_Name = "Encryption"
Option Explicit

Public Function XORDecryption(DataIn As String) As String
    
    Dim CodeKey As String
    
    Dim lonDataPtr As Long
    Dim strDataOut As String
    Dim intXOrValue1 As Integer
    Dim intXOrValue2 As Integer
    
    'Generate a unique CodeKey
    CodeKey = GetSerial()
    
    For lonDataPtr = 1 To (Len(DataIn) / 2)
        'The first value to be XOr-ed comes from the data to be encrypted
        intXOrValue1 = Val("&H" & (Mid$(DataIn, (2 * lonDataPtr) - 1, 2)))
        'The second value comes from the code key
        intXOrValue2 = Asc(Mid$(CodeKey, ((lonDataPtr Mod Len(CodeKey)) + 1), 1))
        
        strDataOut = strDataOut + Chr(intXOrValue1 Xor intXOrValue2)
    Next lonDataPtr
    XORDecryption = strDataOut
End Function

Public Function XOREncryption(DataIn As String) As String
    
    Dim CodeKey As String
    
    Dim lonDataPtr As Long
    Dim strDataOut As String
    Dim temp As Integer
    Dim tempstring As String
    Dim intXOrValue1 As Integer
    Dim intXOrValue2 As Integer

    'Generate a unique CodeKey
    CodeKey = GetSerial()
    
    For lonDataPtr = 1 To Len(DataIn)
        'The first value to be XOr-ed comes from the data to be encrypted
        intXOrValue1 = Asc(Mid$(DataIn, lonDataPtr, 1))
        'The second value comes from the code key
        intXOrValue2 = Asc(Mid$(CodeKey, ((lonDataPtr Mod Len(CodeKey)) + 1), 1))
        
        temp = (intXOrValue1 Xor intXOrValue2)
        tempstring = Hex(temp)
        If Len(tempstring) = 1 Then tempstring = "0" & tempstring
        
        strDataOut = strDataOut + tempstring
    Next lonDataPtr
    XOREncryption = strDataOut
End Function

Private Function GetSerial() As String
'Generates a unique computer ID

    Dim Serial As String
    Dim Script As String
    
    'Turn off error checking - if anything goes wrong it will use a default value
    On Error Resume Next
    
    'Get the computer's serial number from the system_profiler
    Script = "do shell script ""system_profiler SPHardwareDataType | awk '/Serial/ {print $4}'"""
    Serial = MacScript(Script)
   
    'Convert to hex
    Serial = Hex(Serial)
    
    'If something went wrong above or a real number wasn't returned, set a default
    If Len(Serial) < 3 Then
        Serial = "dj2ijg84nvnwj38gnm90dopqm9256dmn"
    End If
    
    'Set return value
    GetSerial = Serial

End Function

