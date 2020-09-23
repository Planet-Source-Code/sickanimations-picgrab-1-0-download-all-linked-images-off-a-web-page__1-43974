VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmSaving 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Saving..."
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   1750
      Width           =   1095
   End
   Begin VB.CommandButton cmdSkip 
      Caption         =   "Skip"
      Height          =   375
      Left            =   3840
      TabIndex        =   9
      Top             =   1750
      Width           =   1095
   End
   Begin VB.PictureBox picProgressImage 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   0
      Picture         =   "frmSaving.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   6000
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.PictureBox pbProgress 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   6000
      TabIndex        =   5
      Top             =   2160
      Width           =   6000
   End
   Begin VB.ListBox lstFiles 
      Height          =   1230
      Left            =   6240
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   0
      Picture         =   "frmSaving.frx":4694
      ScaleHeight     =   1125
      ScaleWidth      =   6000
      TabIndex        =   2
      Top             =   0
      Width           =   6000
      Begin InetCtlsObjects.Inet iTransfer 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
   End
   Begin VB.PictureBox picImage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   0
      Picture         =   "frmSaving.frx":1A668
      ScaleHeight     =   150
      ScaleWidth      =   6030
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   6030
   End
   Begin VB.Timer tmrAnimate 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picAnimation 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   0
      ScaleHeight     =   2.778
      ScaleMode       =   0  'User
      ScaleWidth      =   384.652
      TabIndex        =   0
      Top             =   1125
      Width           =   6015
   End
   Begin VB.Line lProgress2 
      X1              =   6000
      X2              =   0
      Y1              =   2410
      Y2              =   2410
   End
   Begin VB.Line lProgress 
      X1              =   6000
      X2              =   0
      Y1              =   2150
      Y2              =   2150
   End
   Begin VB.Label lblTotalProgress 
      Caption         =   "Image 0 of 0"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1920
      Width           =   6015
   End
   Begin VB.Label lblInformation 
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1320
      Width           =   6015
   End
End
Attribute VB_Name = "frmSaving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Value As Single
Dim Image1Left As Single, Image2Left As Single



Private Sub cmdCancel_Click()
Select Case LCase(cmdCancel.Caption)
Case "cancel"
    
    Dim Result As VbMsgBoxResult
    
    Result = MsgBox("Are you sure you want to end all file downloads?", vbYesNo + vbQuestion, "Skip File?")
    If Result = vbNo Then Exit Sub
    
    DownloadFile = False
    Unload Me
    
Case "exit"
    
    DownloadFile = False
    Status "Downloaded Files."
    Unload Me

End Select
End Sub

Private Sub cmdSkip_Click()
Dim Result As VbMsgBoxResult

Result = MsgBox("Are you sure you want to skip this file?", vbYesNo + vbQuestion, "Skip File?")
End Sub

Private Sub Form_Load()

Image1Left = 200
Image2Left = -200

cmdCancel.Caption = "Cancel"
cmdSkip.Enabled = True

SaveStatus "Please wait..."

End Sub

Private Sub iTransfer_StateChanged(ByVal State As Integer)
Select Case State
Case 0
    SaveStatus "Done."
Case 1
    SaveStatus "Resolving host..."
Case 2
    SaveStatus "Host Resolved."
Case 3
    SaveStatus "Connecting to host..."
Case 4
    SaveStatus "Connected to URL."
Case 5
    SaveStatus "Preparing data request..."
Case 6
    SaveStatus "Data request sent."
Case 7
    SaveStatus "Recieving data..."
Case 8
    SaveStatus "Data recieved."
Case 9
    SaveStatus "Disconecting..."
Case 10
    SaveStatus "Disconected."
Case 11
    SaveStatus "Transfer Error!"
    TransferError = True
Case 12
    SaveStatus "Data transfer complete."
End Select

End Sub

Private Sub tmrAnimate_Timer()

If Image1Left > 400 Then Image1Left = -400
If Image2Left > 400 Then Image2Left = -400

picAnimation.Cls

BitBlt picAnimation.hDC, Image1Left, 0, 405, 5, picImage.hDC, 0, 0, vbSrcCopy
BitBlt picAnimation.hDC, Image2Left, 0, 405, 5, picImage.hDC, 0, 0, vbSrcCopy

picAnimation.Refresh

Image1Left = Image1Left + 2
Image2Left = Image2Left + 2

End Sub

Function Progress(Maximum As Single, Value As Single, MaxWidth As Single) As Single
Dim Percentage As Single, Width As Single

If Maximum <= 0 Or MaxWidth <= 0 Then
    Progress = 0
    Exit Function
End If

If Value >= Maximum Then
    Progress = MaxWidth
    Exit Function
End If


Percentage = (Value * 100) / Maximum

SetProgress:

Progress = (Percentage / 100) * MaxWidth

End Function

Sub DrawProgress(Max As Single, Value As Single)
Dim ImageWidth As Single

ImageWidth = Progress(Max, Value, pbProgress.Width)

pbProgress.Cls
BitBlt pbProgress.hDC, 0, 0, ImageWidth / Screen.TwipsPerPixelX, 30, picProgressImage.hDC, 0, 0, vbSrcCopy
pbProgress.Refresh

End Sub

Public Sub RecieveFiles()
Dim URL As String, OriginalName As String, FileName As String

Me.Show , frmMain

For i = 0 To lstFiles.ListCount - 1

    If DownloadFile = False Then Exit Sub
    
    lstFiles.ListIndex = i
    
    URL = lstFiles.text
    
    OriginalName = modMain.GetFileName(URL)
    
    Select Case sNameCode
    Case ""
        FileName = ""
        sFilePath = sFileName
    Case Else
        sNameCode = GetSetting("PicGrab", "Settings", "NameCode", "/title/ #/number/./extention/")
        FileName = modNameCode.CodeToName(sNameCode, OriginalName, sMinDigits, sCurrentNumber)
    End Select
    
    Me.lblTotalProgress.Caption = "Image " & i + 1 & " of " & lstFiles.ListCount
    DrawProgress lstFiles.ListCount, i + 1
    
    If SaveFile(URL, sFilePath & "\" & FileName) = 0 Then
        SaveStatus "Transfer error, operation aborted."
        tmrAnimate.Enabled = False
        BitBlt picAnimation.hDC, 0, 0, 405, 5, picImage.hDC, 0, 0, vbSrcCopy
        DrawProgress lstFiles.ListCount, lstFiles.ListCount
        cmdSkip.Enabled = False
        cmdCancel.Caption = "Exit"
        Exit Sub
    End If

sCurrentNumber = sCurrentNumber + 1
SaveSetting "PicGrab", "Settings", "StartFrom", sCurrentNumber

Next i

DownloadFile = False

Image1Left = 0
Image2Left = 0
tmrAnimate_Timer
tmrAnimate.Enabled = False

SaveStatus "All transfers complete."
cmdSkip.Enabled = False
cmdCancel.Caption = "Exit"
End Sub
