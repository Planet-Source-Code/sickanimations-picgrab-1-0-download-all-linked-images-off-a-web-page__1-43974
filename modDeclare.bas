Attribute VB_Name = "modDeclare"
Public StartTransfer As Boolean

Public sNameCode As String, sFileName As String, sMinDigits As Single, sCurrentNumber As Single, sFilePath As String
Public TransferError As Boolean
Public DownloadFile As Boolean

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
