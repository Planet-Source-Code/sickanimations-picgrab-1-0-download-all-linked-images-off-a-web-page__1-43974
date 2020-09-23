VERSION 5.00
Begin VB.Form frmSaveAll 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Save"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox drive 
      Height          =   315
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   4575
   End
   Begin VB.Frame fmeFiles 
      Caption         =   "Files"
      Height          =   3975
      Left            =   4680
      TabIndex        =   12
      Top             =   0
      Width           =   4575
      Begin VB.TextBox txtOutputName 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   3120
         Width           =   3255
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   2400
         TabIndex        =   16
         Top             =   3480
         Width           =   2055
      End
      Begin VB.CommandButton cmdSaveAll 
         Caption         =   "Save files in list"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   3480
         Width           =   2175
      End
      Begin VB.ListBox lstFiles 
         Height          =   2790
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label lblSavedName 
         Caption         =   "Saved Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3120
         Width           =   975
      End
   End
   Begin VB.Frame fmeName 
      Caption         =   "Name"
      Height          =   1215
      Left            =   0
      TabIndex        =   6
      Top             =   2160
      Width           =   4575
      Begin VB.CommandButton cmdDefault 
         Caption         =   "Default"
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "Help"
         Height          =   255
         Left            =   3480
         TabIndex        =   9
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Text            =   "/title/ #/number/./extention/"
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label lblHelp 
         Caption         =   "If confused, click default."
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblTypeName 
         Caption         =   "Name Code:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fmeNumber 
      Caption         =   "Number:"
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   3360
      Width           =   4575
      Begin VB.TextBox txtMinimumDigits 
         Height          =   285
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   5
         Text            =   "4"
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtStartFrom 
         Height          =   285
         Left            =   960
         MaxLength       =   4
         TabIndex        =   3
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Minimum Digits:"
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblStartOn 
         Caption         =   "Start From:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
End
Attribute VB_Name = "frmSaveAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentFileNumber As Single

Private Sub cmdDefault_Click()
txtName.text = "/title/ #/number/./extention/"
End Sub

Private Sub cmdHelp_Click()
frmNameCodeHelp.Show vbModal, Me
End Sub

Private Sub cmdRemove_Click()
On Error Resume Next
lstFiles.RemoveItem lstFiles.ListIndex
End Sub

Private Sub cmdSaveAll_Click()


sNameCode = txtName.text
sMinDigits = Val(txtMinimumDigits.text)
sCurrentNumber = Val(txtStartFrom.text)
sFilePath = Dir1.Path

Me.Hide

frmSaving.lstFiles.Clear

For i = 0 To lstFiles.ListCount - 1
    lstFiles.ListIndex = i
    frmSaving.lstFiles.AddItem lstFiles.text
Next i

DownloadFile = True
frmSaving.RecieveFiles

End Sub

Private Sub Dir1_Change()
SaveSetting "PicGrab", "Settings", "Path", Dir1.Path
End Sub

Private Sub drive_Change()
Dir1.Path = drive.drive
End Sub

Private Sub Form_Load()
On Error Resume Next
txtName.text = GetSetting("PicGrab", "Settings", "NameCode", "/title/ #/number/./extention/")
txtMinimumDigits.text = GetSetting("PicGrab", "Settings", "MinDigits", "4")
txtStartFrom.text = GetSetting("PicGrab", "Settings", "StartFrom", "0")
Dir1.Path = GetSetting("PicGrab", "Settings", "Path", App.Path)
End Sub

Private Sub lstFiles_Click()
If lstFiles.ListIndex < 0 Then Exit Sub
txtOutputName.text = modNameCode.CodeToName(txtName.text, modMain.GetFileName(lstFiles.text), Val(txtMinimumDigits.text), Val(txtStartFrom.text) + lstFiles.ListIndex)
End Sub


Private Sub txtMinimumDigits_LostFocus()
txtMinimumDigits.text = Val(txtMinimumDigits.text)
If Val(txtMinimumDigits.text) > 5 Then txtMinimumDigits.text = "5"

SaveSetting "PicGrab", "Settings", "MinDigits", Val(txtMinimumDigits.text)
End Sub

Private Sub txtName_LostFocus()

If lstFiles.ListCount > 1 And InStr(1, txtName.text, "/name/") = 0 And InStr(1, txtName.text, "/number/") = 0 Then
    MsgBox "You MUST include '/number/' or '/name/' in the namecode." & vbCrLf & "Reason: All files will have the same name.", vbExclamation + vbOKOnly
    cmdDefault_Click
End If

SaveSetting "PicGrab", "Settings", "NameCode", txtName.text
End Sub

Private Sub txtStartFrom_LostFocus()
txtStartFrom.text = Val(txtStartFrom.text)
SaveSetting "PicGrab", "Settings", "StartFrom", Val(txtStartFrom.text)
End Sub
