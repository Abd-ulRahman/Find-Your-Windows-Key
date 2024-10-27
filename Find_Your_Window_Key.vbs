Const HKLM = &H80000002

wscript.echo "View Product Keys | Microsoft Products" & vbCrLf

'Install Date 
Computer = "."
Set objWMIService = GetObject("winmgmts:\\" & Computer & "\root\cimv2")
Set Obj = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")

Dim objshell,path,ProductData,ProductID
Set objshell = CreateObject("WScript.Shell")

dim InsDate

For Each item in Obj
  InsDate = item.InstallDate
  ' Gather Operating System Information
  Caption = Item.Caption
  OSArchitecture = Item.OSArchitecture
  CSDVersion = Item.CSDVersion
  Version = Item.Version
  Next

dim NewDate

NewDate = mid(InsDate,9,2) & ":" & mid(InsDate,11,2) & ":" & mid(InsDate,13,2)
NewDate = NewDate & " " & mid(InsDate,7,2) & "/" & mid(InsDate,5,2) & "/" & mid(InsDate,1,4)

QueryWindowsProductKeys() 

'wscript.echo vbCrLf & "Office Keys" & vbCrLf

Function DecodeProductKey(arrKey)
    Const KeyOffset = 52
  If Not IsArray(arrKey) Then Exit Function
    intIsWin8 = BitShiftRight(arrKey(66) \ 6, 3) And 1    
    arrKey(66) = arrKey(66) And 247 Or BitShiftLeft(intIsWin8 And 2,2)
    i = 24
    strChars = "BCDFGHJKMPQRTVWXY2346789"
    strKeyOutput = ""
    While i > -1
        intCur = 0
        intX = 14
        While intX > -1
            intCur = BitShiftLeft(intCur,8)
            intCur = arrKey(intX + KeyOffset) + intCur
            arrKey(intX + KeyOffset) = Int(intCur / 24) 
            intCur = intCur Mod 24
            intX = intX - 1
        Wend
        i = i - 1
        strKeyOutput = Mid(strChars,intCur + 1,1) & strKeyOutput
        intLast = intCur
    Wend
    If intIsWin8 = 1 Then
        strKeyOutput = Mid(strKeyOutput,2,intLast) & "N" & Right(strKeyOutput,Len(strKeyOutput) - (intLast + 1))    
    End If
    strKeyGUIDOutput = Mid(strKeyOutput,1,5) & "-" & Mid(strKeyOutput,6,5) & "-" & Mid(strKeyOutput,11,5) & "-" & Mid(strKeyOutput,16,5) & "-" & Mid(strKeyOutput,21,5)
    DecodeProductKey = strKeyGUIDOutput
End Function

Function RegReadBinary(strRegPath,strRegValue)
    Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    objReg.GetBinaryValue HKLM,strRegPath,strRegValue,arrRegBinaryData
    RegReadBinary = arrRegBinaryData
    Set objReg = Nothing
End Function

Function BitShiftLeft(intValue,intShift)
    BitShiftLeft = intValue * 2 ^ intShift
End Function

Function BitShiftRight(intValue,intShift)
    BitShiftRight = Int(intValue / (2 ^ intShift))
End Function

'Windows Product Key
Sub QueryWindowsProductKeys()
    strWinKey = CheckWindowsKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion","DigitalProductId",52)
    Path = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\"
    If strWinKey <> "" Then
       'Registry key value
        ProductID = objshell.RegRead(Path & "ProductID")
        wscript.echo "Product: " & Caption & Version & " (" & OSArchitecture & ")"
        wscript.echo "ProductID : "  & ProductID 
        wscript.echo "Installation Date: " & NewDate 
	    ProductData =  "Product: " & Caption  & Version & " (" & OSArchitecture & ")"  & vbNewLine & "ProductID : "  & ProductID & vbNewLine & "Installation Date: " &  NewDate  & vbNewLine & "Installed Key: " &  strWinKey
        WriteData ProductData
        Exit Sub
    End If
    strWinKey = CheckWindowsKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion","DigitalProductId4",808)
    If strWinKey <> "" Then
       'Registry key value
        ProductID = objshell.RegRead(Path & "ProductID")
        wscript.echo "Product: " & Caption & Version & " (" & OSArchitecture & ")"
        wscript.echo "ProductID : "  & ProductID 
        wscript.echo "Installation Date: " & NewDate 
	    ProductData =  "Product: " & Caption  & Version & " (" & OSArchitecture & ")"  & vbNewLine & "ProductID : "  & ProductID & vbNewLine & "Installation Date: " &  NewDate  & vbNewLine & "Installed Key: " &  strWinKey
        WriteData ProductData
        Exit Sub
    End If
    strWinKey = CheckWindowsKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion\DefaultProductKey","DigitalProductId",52)
    If strWinKey <> "" Then
       'Registry key value
        ProductID = objshell.RegRead(Path & "ProductID")
        wscript.echo "Product: " & Caption & Version & " (" & OSArchitecture & ")"
        wscript.echo "ProductID : "  & ProductID 
        wscript.echo "Installation Date: " & NewDate 
	    ProductData =  "Product: " & Caption  & Version & " (" & OSArchitecture & ")"  & vbNewLine & "ProductID : "  & ProductID & vbNewLine & "Installation Date: " &  NewDate  & vbNewLine & "Installed Key: " &  strWinKey
        WriteData ProductData
        Exit Sub
    End If
    strWinKey = CheckWindowsKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion\DefaultProductKey","DigitalProductId4",808)
    If strWinKey <> "" Then
       'Registry key value
        ProductID = objshell.RegRead(Path & "ProductID")
        wscript.echo "Product: " & Caption & Version & " (" & OSArchitecture & ")"
        wscript.echo "ProductID : "  & ProductID 
        wscript.echo "Installation Date: " & NewDate 
	    ProductData =  "Product: " & Caption  & Version & " (" & OSArchitecture & ")"  & vbNewLine & "ProductID : "  & ProductID & vbNewLine & "Installation Date: " &  NewDate  & vbNewLine & "Installed Key: " &  strWinKey
        WriteData ProductData
        Exit Sub
    End If
End Sub

Function CheckWindowsKey(strRegPath,strRegValue,intKeyOffset)
    strWinKey = DecodeProductKey(RegReadBinary(strRegPath,strRegValue))
    If strWinKey <> "BBBBB-BBBBB-BBBBB-BBBBB-BBBBB" And strWinKey <> "" Then
        CheckWindowsKey = strWinKey
    Else
        CheckWindowsKey = ""
    End If
End Function

Function RegReadBinary(strRegPath,strRegValue)
    Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    objReg.GetBinaryValue HKLM,strRegPath,strRegValue,arrRegBinaryData
    RegReadBinary = arrRegBinaryData
    Set objReg = Nothing
End Function

Function OsArch()
    Set objShell = WScript.CreateObject("WScript.Shell")
    If objShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%") = "%ProgramFiles(x86)%" Then
        OsArch = "x86" 
    Else
        OsArch = "x64"
    End If
    Set objShell = Nothing
End Function

Sub WriteData(strValue)

    Dim fso, fName, txt,objshell,UserName

    Set objShell = CreateObject("WScript.Shell")
    UserName = objshell.ExpandEnvironmentStrings("%UserName%") 
    'Create a text file on desktop 
    fName = "C:\Users\" & UserName & "\Desktop\WindowsKeyInfo.txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txt = fso.CreateTextFile(fName)
    txt.Writeline strValue
    txt.Close

    WScript.Echo  strValue

End Sub