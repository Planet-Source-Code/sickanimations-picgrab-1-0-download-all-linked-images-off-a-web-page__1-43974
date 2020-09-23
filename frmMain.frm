VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "[s]Animations PicGrab"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkClearImages 
      Appearance      =   0  'Flat
      Caption         =   "Clear Images on navigate."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   3000
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.CheckBox chkClearLinks 
      Appearance      =   0  'Flat
      Caption         =   "Clear links on navigate."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   4560
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdNavigate 
      Caption         =   "Navigate to URL"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton cmdRemoveAll 
      Caption         =   "Remove All"
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdSaveAll 
      Caption         =   "Save All"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.ListBox lstLinks 
      Appearance      =   0  'Flat
      Height          =   810
      Left            =   0
      TabIndex        =   5
      Top             =   3720
      Width           =   5175
   End
   Begin InetCtlsObjects.Inet iTransfer 
      Left            =   4680
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin VB.ListBox lstImages 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   5175
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   5040
      Width           =   5175
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   5040
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   5040
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label lblLinks 
      BackStyle       =   0  'Transparent
      Caption         =   "All Links:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   3480
      Width           =   735
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   5040
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label lblImages 
      BackStyle       =   0  'Transparent
      Caption         =   "Images Found:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5040
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblAddress 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNavigate_Click()
If lstLinks.text = "" Then Exit Sub
txtAddress.text = lstLinks.text
txtAddress_KeyUp 13, 0
End Sub

Private Sub cmdRemove_Click()
On Error Resume Next
lstImages.RemoveItem lstImages.ListIndex
End Sub

Private Sub cmdRemoveAll_Click()
lstImages.Clear
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrorHandler

If lstImages.text = "" Then MsgBox "You must select an image from the list.", vbExclamation + vbOKOnly: Exit Sub


ChoosePath:

cd.FileName = modMain.GetFileName(lstImages.text)
cd.ShowSave

Unload frmSaveAll

frmSaveAll.Dir1.Path = GetDirectory(cd.FileName)
frmSaveAll.txtName.text = modNameCode.GetFile(cd.FileName)
frmSaveAll.lstFiles.Clear
frmSaveAll.lstFiles.AddItem lstImages.text

frmSaveAll.Show vbModal, Me

Exit Sub
ErrorHandler:
Status Err.Description
End Sub

Private Sub cmdSaveAll_Click()

If lstImages.ListCount <= 0 Then
    MsgBox "There are no images to save.", vbExclamation + vbOKOnly
    Exit Sub
End If

Unload frmSaveAll

For i = 0 To lstImages.ListCount - 1
    lstImages.ListIndex = i
    frmSaveAll.lstFiles.AddItem lstImages.text
Next i

frmSaveAll.Show vbModal, Me

End Sub

Private Sub Form_Click()
'frmSaving.Show vbModal, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub iTransfer_StateChanged(ByVal State As Integer)
Select Case State
Case 0

Case 1
Status "Resolving host..."
Case 2
Status "Host Resolved."
Case 3
Status "Connecting to host..."
Case 4
Status "Connected to URL."
Case 5
Status "Preparing data request..."
Case 6
Status "Data request sent."
Case 7
Status "Recieving data..."
Case 8
Status "Data recieved."
Case 9
Status "Disconecting..."
Case 10
Status "Disconected."
Case 11
Status "Transfer Error!"
Case 12
Status "Data transfer complete."
End Select
End Sub

Private Sub lstLinks_Click()
Debug.Print lstLinks.text
End Sub

Private Sub txtAddress_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> 13 Then Exit Sub

Dim DocumentData As String, URL As String, NetDir As String, temp() As String

iTransfer.Cancel

URL = txtAddress.text
URL = Replace(URL, " ", "%20")


temp() = Split(URL, "/")

For i = 0 To UBound(temp) - 1
    NetDir = NetDir & temp(i) & "/"
Next i


If Left(URL, Len("http://")) <> "http://" Then URL = "http://" & URL

Status "Opening URL..."

DocumentData = iTransfer.OpenURL(URL, icString)

If chkClearImages.Value = 1 Then lstImages.Clear
If chkClearLinks.Value = 1 Then lstLinks.Clear

modMain.ParseLinks DocumentData, ".jpg", NetDir, frmMain.lstImages
modMain.ParseLinks DocumentData, "*", NetDir, frmMain.lstLinks

End Sub
