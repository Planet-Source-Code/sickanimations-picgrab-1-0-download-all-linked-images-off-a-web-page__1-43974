Attribute VB_Name = "modMain"
Public Sub Status(text As String)
frmMain.lblStatus.Caption = text
End Sub

Sub ParseLinks(document As String, Extention As String, SourceDir As String, LstBox As ListBox)
Dim NoSpaces As String
Dim StartLinkPosition As Single, EndLinkPosition As Single
Dim LinkURL As String
Dim TempDir As String

Status "Data Recieved, Parsing code."

NoSpaces = Replace(document, " ", "")

If Right(SourceDir, 1) <> "/" Then SourceDir = SourceDir & "/"


NextLink:

StartLinkPosition = InStr(1, NoSpaces, "href=" & Chr(34)) + Len("href=" & Chr(34))
If StartLinkPosition > Len("href=" & Chr(34)) Then
    NoSpaces = Mid(NoSpaces, StartLinkPosition)
    EndLinkPosition = InStr(1, NoSpaces, Chr(34)) - Len(Chr(34))
    
    LinkURL = Left(NoSpaces, EndLinkPosition)
    
    Select Case Extention
    Case "*"
        If Left(LCase(LinkURL), Len("http://")) <> "http://" Then
        
        LinkURL = CorrectLink(SourceDir, LinkURL)
        
        End If
    Case Else
        'Geting links to specified file types
        If Right(LinkURL, Len(Extention)) <> Extention Then GoTo NextLink 'Not type specified
        
        If Left(LCase(LinkURL), Len("http://")) <> "http://" Then
        
        LinkURL = CorrectLink(SourceDir, LinkURL)
        
        End If
    End Select
    
    LstBox.AddItem LinkURL
    
    GoTo NextLink
    
Else

End If

End Sub

Function GetParent(URL As String)
Dim temp() As String
temp() = Split(URL, "/")

If UBound(temp) < 1 Then GetParent = URL: Exit Function

For i = 1 To UBound(temp)
    GetParent = GetParent & temp(i) & "/"
Next i

End Function

Function GetFileName(URL As String)
Dim temp() As String

temp() = Split(URL, "/")

GetFileName = temp(UBound(temp))

End Function
Function CorrectLink(URL As String, Link As String) As String
Dim HostName As String

HostName = GetHostName(URL)

CheckLink:
If Left(Link, Len("../")) = "../" Then
    'This refers to a parent directory, we have to get it
    Link = Mid(Link, Len("../"))
    URL = GetParent(URL)
    GoTo CheckLink

ElseIf Left(LCase(Link), Len("/")) = "/" Then
    CorrectLink = HostName & Link
ElseIf Left(Link, Len("javascript:")) = "javascript:" Then
    'This is a javascript link
    CorrectLink = Link

ElseIf Left(Link, Len("mailto:")) = "mailto:" Then
    'This is a email link
    CorrectLink = Link

ElseIf Left(LCase(Link), Len("http://")) = "http://" Then
    'This is an absolute link
    CorrectLink = URL & Link

Else
    'It is a relative link
    CorrectLink = URL & Link
    
End If

End Function

Function GetHostName(URL As String)
Dim temp() As String
If Left(URL, Len("http://")) <> "http://" Then URL = URL & "http://"

temp() = Split(Mid(URL, Len("http://*")), "/")

If UBound(temp) < 0 Then
    GetHostName = URL
Else
    GetHostName = "http://" & temp(0)
End If

End Function

Function SaveFile(URL As String, FilePath As String) As Single
Dim FileData() As Byte
On Error Resume Next

TransferError = False

FileData = frmSaving.iTransfer.OpenURL(URL, icByteArray)

If TransferError = True Then
    On Error Resume Next
    SaveFile = 0
Exit Function
End If

SaveStatus "File recieved, saving."

Open FilePath For Binary Access Write As #1
Put #1, , FileData()
Close #1

SaveStatus "File saved."

SaveFile = 1

End Function

Public Sub SaveStatus(StatusText As String)
frmSaving.lblInformation.Caption = StatusText
End Sub
