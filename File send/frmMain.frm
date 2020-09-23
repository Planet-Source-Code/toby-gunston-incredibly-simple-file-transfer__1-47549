VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Transfer"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock wsSend 
      Left            =   1440
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   360
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsWait 
      Left            =   960
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   500
   End
   Begin VB.Frame Frame2 
      Caption         =   "Send a file"
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   6615
      Begin VB.TextBox txtIP 
         Height          =   375
         Left            =   2280
         TabIndex        =   16
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtSendToPort 
         Height          =   375
         Left            =   4200
         TabIndex        =   15
         Text            =   "500"
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton cmdBrowseFTS 
         Caption         =   "B&rowse"
         Height          =   375
         Left            =   5160
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtFileToSend 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   4815
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "&Send"
         Height          =   375
         Left            =   5160
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Port:"
         Height          =   195
         Left            =   3720
         TabIndex        =   14
         Top             =   1080
         Width           =   330
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "IP Address:"
         Height          =   195
         Left            =   1320
         TabIndex        =   13
         Top             =   1080
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Connect to:"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "File to send:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Wait for a file"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtWaitOnPort 
         Height          =   375
         Left            =   4200
         TabIndex        =   11
         Text            =   "500"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtSaveTo 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   4815
      End
      Begin VB.CommandButton cmdBrowseST 
         Caption         =   "&Browse"
         Height          =   375
         Left            =   5160
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdWait 
         Caption         =   "&Wait"
         Height          =   375
         Left            =   5160
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Wait on Port:"
         Height          =   195
         Left            =   3120
         TabIndex        =   10
         Top             =   1080
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Save to:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FileOpen As Boolean

Private Sub cmdBrowseFTS_Click()
dlgOpen.ShowOpen
If dlgOpen.Filename <> "" Then
    txtFileToSend = dlgOpen.Filename
End If
End Sub

Private Sub cmdBrowseST_Click()
frmBrowse.Show 1
If SaveTo <> "" Then txtSaveTo = SaveTo
End Sub

Private Sub cmdSend_Click()
Dim Data As String              'Stores the data we're going to send
Dim BytesRemaining As Long      'Stores the number of bytes still to be sent
Dim NumOfBytes As Long          'Max number of bytes to be sent each time

NumOfBytes = 1024
DisableControls False
'Reset the control (IE it might already be connected)
wsSend.Close
'Connect to the specified IP Address and Port
wsSend.Connect txtIP, txtSendToPort
'Wait until we are connected
Do While wsSend.State <> sckConnected
    DoEvents
Loop
'Send the filename of our file first
wsSend.SendData GetFileTitle(txtFileToSend)
DoEvents
Open txtFileToSend For Binary Access Read As #1
'Get the amount of bytes we're going to send
BytesRemaining = FileLen(txtFileToSend)
'While we're not at the end of our file...
Do While EOF(1) = False
    'Get a chunk of data from the file
    Data = Input(NumOfBytes, 1)
    'Decrease the value in BytesRemaining by how much data we took
    BytesRemaining = BytesRemaining - NumOfBytes
    'Send the data
    wsSend.SendData Data
    'Display how much is still to be sent
    Me.Caption = "Bytes left: " & BytesRemaining
    'Wait until weve finished the above before continuing
    DoEvents
Loop
Close #1
'Let the other end know we've finished sending data
wsSend.SendData "xx"
DoEvents
DisableControls True
Me.Caption = "File sent"
End Sub

Private Sub cmdWait_Click()
DisableControls False
'Reset the control (IE it might already be connected)
wsWait.Close
'Set the port number the user has to connect to
wsWait.LocalPort = txtWaitOnPort
'Listen for a connection
wsWait.Listen
'Display the staus to the user
Me.Caption = "Waiting..."
End Sub

Private Sub Form_Load()
'Default to the current machine (for testing purposes)
txtIP = wsSend.LocalIP
End Sub

Private Sub wsWait_ConnectionRequest(ByVal requestID As Long)
'Reset the control (IE it might already be connected)
wsWait.Close
'Accept the connection
wsWait.Accept requestID
End Sub

Private Sub wsWait_DataArrival(ByVal bytesTotal As Long)
'Stores the data we receive
Dim Data As String

'Let the user know we're downloading data
Me.Caption = "Downloading file, please wait..."
'If we havnt already got the filename then get it
If FileOpen = False Then
    'Get the filename
    wsWait.GetData Data
    Open txtSaveTo & "\" & Data For Binary Access Write As #1
    FileOpen = True
    Exit Sub
End If
'Get the data sent to us
wsWait.GetData Data

'If the user hasn't finished sending then save the data
If Data <> "xx" Then
    Put #1, , Data
Else
    'Else they have so close the file etc
    Close #1
    FileOpen = False
    Me.Caption = "Complete"
    DisableControls True
End If

End Sub

Public Sub DisableControls(Enable As Boolean)
'Disable or enable the controls on the main form
txtSaveTo.Enabled = Enable
cmdBrowseST.Enabled = Enable
txtWaitOnPort.Enabled = Enable

txtFileToSend.Enabled = Enable
cmdBrowseFTS.Enabled = Enable
cmdWait.Enabled = Enable
cmdSend.Enabled = Enable
txtIP.Enabled = Enable
txtSendToPort.Enabled = Enable
End Sub

Public Function GetFileTitle(Filename As String) As String
'Used to get the filetitle from a filename
Dim Count As Integer

Do While Left(Right(Filename, Count), 1) <> "\"
    GetFileTitle = Right(Filename, Count)
    Count = Count + 1
Loop
End Function
