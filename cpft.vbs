'''
' Plaintext File Transfer
' 
' By Jan Seidl
'
' usage: cpft.vbs [decode/encode] <infile> <outfile>
'
'''


Set args = WScript.Arguments

If args.Count() = "3" Then

    Dim inFile, outFile
    action = args.Item(0)
    inFile = args.Item(1)
    outFile = args.Item(2)

    If action = "" Or inFile = "" Or outFile = "" Then
        Err.Raise 1, "Main", "Please supply: cpft.vbs [decode/encode] <infile> <outfile>"
    Else
        DeployEncodedFile action, inFile, outFile
    End If
Else
    Err.Raise 1, "Main", "Please supply: cpft.vbs [decode/encode] <infile> <outfile>"
End If


'' /** BEGIN>>> BASE 64 FUNCTIONS **/

' ATTENTION!
' This code uses base64 functions from
'
' Antonin Foller, http://www.motobit.com
'
'

Function Base64Encode(inData)
  'rfc1521
  '2001 Antonin Foller, Motobit Software, http://Motobit.cz
  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim cOut, sOut, I
  
  'For each group of 3 bytes
  For I = 1 To Len(inData) Step 3
    Dim nGroup, pOut, sGroup
    
    'Create one long from this 3 bytes.
    nGroup = &H10000 * Asc(Mid(inData, I, 1)) + _
      &H100 * MyASC(Mid(inData, I + 1, 1)) + MyASC(Mid(inData, I + 2, 1))
    
    'Oct splits the long To 8 groups with 3 bits
    nGroup = Oct(nGroup)
    
    'Add leading zeros
    nGroup = String(8 - Len(nGroup), "0") & nGroup
    
    'Convert To base64
    pOut = Mid(Base64, CLng("&o" & Mid(nGroup, 1, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 3, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 5, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 7, 2)) + 1, 1)
    
    'Add the part To OutPut string
    sOut = sOut + pOut
    
    'Add a new line For Each 76 chars In dest (76*3/4 = 57)
    'If (I + 2) Mod 57 = 0 Then sOut = sOut + vbCrLf
  Next
  Select Case Len(inData) Mod 3
    Case 1: '8 bit final
      sOut = Left(sOut, Len(sOut) - 2) + "=="
    Case 2: '16 bit final
      sOut = Left(sOut, Len(sOut) - 1) + "="
  End Select
  Base64Encode = sOut
End Function

' Decodes a base-64 encoded string (BSTR type).
' 1999 - 2004 Antonin Foller, http://www.motobit.com
' 1.01 - solves problem with Access And 'Compare Database' (InStr)
Function Base64Decode(ByVal base64String)
  'rfc1521
  '1999 Antonin Foller, Motobit Software, http://Motobit.cz
  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim dataLength, sOut, groupBegin
  
  'remove white spaces, If any
  base64String = Replace(base64String, vbCrLf, "")
  base64String = Replace(base64String, vbTab, "")
  base64String = Replace(base64String, " ", "")
  
  'The source must consists from groups with Len of 4 chars
  dataLength = Len(base64String)
  If dataLength Mod 4 <> 0 Then
    Err.Raise 1, "Base64Decode", "Bad Base64 string."
    Exit Function
  End If

  
  ' Now decode each group:
  For groupBegin = 1 To dataLength Step 4
    Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut
    ' Each data group encodes up To 3 actual bytes.
    numDataBytes = 3
    nGroup = 0

    For CharCounter = 0 To 3
      ' Convert each character into 6 bits of data, And add it To
      ' an integer For temporary storage.  If a character is a '=', there
      ' is one fewer data byte.  (There can only be a maximum of 2 '=' In
      ' the whole string.)

      thisChar = Mid(base64String, groupBegin + CharCounter, 1)

      If thisChar = "=" Then
        numDataBytes = numDataBytes - 1
        thisData = 0
      Else
        thisData = InStr(1, Base64, thisChar, vbBinaryCompare) - 1
      End If
      If thisData = -1 Then
        Err.Raise 2, "Base64Decode", "Bad character In Base64 string."
        Exit Function
      End If

      nGroup = 64 * nGroup + thisData
    Next
    
    'Hex splits the long To 6 groups with 4 bits
    nGroup = Hex(nGroup)
    
    'Add leading zeros
    nGroup = String(6 - Len(nGroup), "0") & nGroup
    
    'Convert the 3 byte hex integer (6 chars) To 3 characters
    pOut = Chr(CByte("&H" & Mid(nGroup, 1, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 3, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 5, 2)))
    
    'add numDataBytes characters To out string
    sOut = sOut & Left(pOut, numDataBytes)
  Next

  Base64Decode = sOut
End Function

Function MyASC(OneChar)
  If OneChar = "" Then MyASC = 0 Else MyASC = Asc(OneChar)
End Function


'' /** <<<END BASE 64 FUNCTIONS **/

'' /** BEGIN>>> FILE OPERATIONS **/

Sub WriteOutput(content, outFile)

        wscript.Echo "Writing output to "+outFile
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.CreateTextFile(outFile)
    objFile.Write(content)
    objFile.Close
        wscript.Echo "Wrote "+outFile+".Process completed."

End Sub

'' /** <<<END FILE OPERATIONS **/


'' /** BEGIN>>> MAIN **/

' Open File, Decode, Write to File
Sub DeployEncodedFile(action, inFile, outFile)

    Dim objFSO, objFile, objReadFile, strContents
    Dim deltaSize, strSize

    Const ForReading = 1

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.GetFile(inFile)
    If objFile.Size > 0 Then
        Set objReadFile = objFSO.OpenTextFile(inFile, ForReading)
        strContents = Trim(objReadFile.ReadAll)

        If action = "decode" Then

            ' Fix weird glitch
            ' Yes, I hate VBS and I'm too lazy to debug
            ' Workaround time!
            strSize = Len(strContents)
            deltaSize = strSize-3
            strContents = Mid(strContents, 4, deltaSize)

            wscript.Echo "Decoding "+inFile+". This uses a pure-vbs implementation and its SLOW. Grab a coffee."
            strContents = Base64Decode(strContents)
            wscript.Echo "Decoding complete."
        Else

            wscript.Echo "Encoding "+inFile+". This uses a pure-vbs implementation and its SLOW. Grab a coffee."
            strContents = Base64Encode(strContents)
            wscript.Echo "Encoding complete."
        End if

        WriteOutput strContents, outFile
        objReadFile.Close
    Else
        Wscript.Echo "The file is empty or cannot be read."
    End If
End Sub

'' /** <<<END MAIN **/
