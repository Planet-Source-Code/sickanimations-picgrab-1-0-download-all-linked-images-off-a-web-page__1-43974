Attribute VB_Name = "modNameCode"
Function CodeToName(Code As String, FileName As String, Digits As Single, Number As Single) As String
Dim FileTitle As String, FileExtention As String, FileNumber As String

FileTitle = GetFileTitle(FileName)
FileExtention = GetFileExtention(FileName)
FileNumber = GetNumber(Digits, Number)

Code = Replace(Code, "/title/", FileTitle)
Code = Replace(Code, "/extention/", FileExtention)
Code = Replace(Code, "/number/", FileNumber)

CodeToName = Code

End Function


Function GetFileTitle(FileName As String) As String
Dim TempString As String, LastPosition As Single, OldLastPosition As Single

TempString = FileName

If InStr(1, TempString, ".") = 0 Then GetFileTitle = FileName: Exit Function

CheckNext:

OldLastPosition = LastPosition
LastPosition = LastPosition + InStr(1, TempString, ".")
TempString = Mid(TempString, LastPosition + 1)

If LastPosition <> OldLastPosition Then GoTo CheckNext

GetFileTitle = Mid(FileName, 1, LastPosition - 1)

End Function


Function GetFileExtention(FileName As String) As String
Dim temp() As String

temp() = Split(FileName, ".")

If UBound(temp) < 1 Then GetFileExtention = "": Exit Function

GetFileExtention = temp(UBound(temp))

End Function


Function GetNumber(Digits As Single, Number As Single) As String
Dim NumberStr As String

NumberStr = CStr(Val(Number))

FillPlaceHolder:

If Len(NumberStr) < Digits Then
    NumberStr = "0" & NumberStr
    GoTo FillPlaceHolder
End If

GetNumber = NumberStr

End Function

Function GetDirectory(Path As String) As String
Dim temp() As String
temp() = Split(Path, "\")
If UBound(temp) < 0 Then GetDirectory = Path: Exit Function

For i = 0 To UBound(temp) - 1
GetDirectory = GetDirectory & temp(i) & "\"
Next i

End Function

Function GetFile(Path As String) As String
Dim temp() As String
temp() = Split(Path, "\")

If UBound(temp) < 0 Then
    GetFile = Path
    Exit Function
End If

GetFile = temp(UBound(temp))

End Function

