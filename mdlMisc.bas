Attribute VB_Name = "mdlMisc"
Option Explicit

#If Win16 Then
    Declare Function WritePrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal NewString As String, ByVal FileName As String) As Integer
    Declare Function GetPrivateProfileString Lib "Kernel" Alias "GetPrivateProfilestring" (ByVal AppName As String, ByVal KeyName As Any, ByVal default As String, ByVal ReturnedString As String, ByVal MAXSIZE As Integer, ByVal FileName As String) As Integer
#Else
    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If

Function File_INIRead(Section, KeyName, FileName As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    File_INIRead = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName, "", sRet, Len(sRet), FileName))
End Function

Function File_INIWrite(sSection As String, sKeyName As String, sNewString As String, sFileName) As String
    Dim r
    r = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
    File_INIWrite = sNewString
End Function


Public Function RandomNumber(ByVal Low As Long, ByVal High As Long) As Long
    RandomNumber = (High - Low + 1) * Rnd + Low
End Function

Public Function Encrypt(ByVal What As String) As String
    On Error GoTo EncryptError
    Dim x As Integer, Number As Long
    For x = 1 To Len(What)
        DoEvents
        Number = RandomNumber(10, 99)                 'A random number with two digits
        Encrypt = Encrypt & Number & Int(Asc(Mid(What, x, 1))) * Number & "-"
    Next x
    For x = 1 To Len(Encrypt)
        DoEvents
        Select Case Mid(Encrypt, x, 1)
            Case "1": Mid(Encrypt, x, 1) = "'"      ''
            Case "2": Mid(Encrypt, x, 1) = Chr(34)  '"
            Case "3": Mid(Encrypt, x, 1) = "."      '.
            Case "4": Mid(Encrypt, x, 1) = ","      ',
            Case "5": Mid(Encrypt, x, 1) = "~"      '~
            Case "6": Mid(Encrypt, x, 1) = "]"      '
            Case "7": Mid(Encrypt, x, 1) = "*"      '*
            Case "8": Mid(Encrypt, x, 1) = "|"      '|
            Case "9": Mid(Encrypt, x, 1) = "-"      '-
            Case "0": Mid(Encrypt, x, 1) = "`"      '_
            Case "-": Mid(Encrypt, x, 1) = "_"      '`
        End Select
    Next x
    Exit Function
EncryptError:
    Encrypt = ""
End Function

Public Function Decrypt(ByVal What As String) As String
    On Error GoTo DecryptError
    Dim x As Integer
    For x = 1 To Len(What)
        DoEvents
        Select Case Mid(What, x, 1)
            Case "'": Mid(What, x, 1) = "1"
            Case Chr(34): Mid(What, x, 1) = "2"     '"
            Case ".": Mid(What, x, 1) = "3"
            Case ",": Mid(What, x, 1) = "4"
            Case "~": Mid(What, x, 1) = "5"
            Case "]": Mid(What, x, 1) = "6"
            Case "*": Mid(What, x, 1) = "7"
            Case "|": Mid(What, x, 1) = "8"
            Case "-": Mid(What, x, 1) = "9"
            Case "`": Mid(What, x, 1) = "0"
            Case "_": Mid(What, x, 1) = "-"
        End Select
    Next x
    For x = 1 To Len(What)
        DoEvents
        Decrypt = Decrypt & Chr(Mid(Mid(What, x, InStr(x, What, "-") - x), 3) / _
        Mid(Mid(What, x, InStr(x, What, "-") - x), 1, 2))
        x = InStr(x, What, "-")
    Next x
    Exit Function
DecryptError:
    Decrypt = ""
End Function
