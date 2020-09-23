VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2640
      Top             =   3720
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.TextBox txtFrom 
         Height          =   405
         Left            =   2640
         TabIndex        =   7
         Text            =   "ENTER YOUR EMAIL ADDRESS HERE"
         Top             =   600
         Width           =   4335
      End
      Begin VB.ComboBox cboSubject 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox txtMessage 
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Text            =   "frmMain.frx":0000
         Top             =   1200
         Width           =   6855
      End
      Begin VB.Label Label2 
         Caption         =   "Enter Your Email Address Here"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Select Subject"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3240
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "status"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   3720
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************
'PROGRAM BY GERRY MC DONNELL 2003
'www.gerrymcdonnell.com
'TOOL TO REPORT PROBLEMS COMMNEST IN UR APP
'BASED ON EMAIL CODE BY Saurabh

'NOTES
'MAKE SURE U LIMIT THE NUMBER OF TIMES AN EMAILC AN SENT OTHER WISE ULL GET LOADS
'DONT FORGET TO FILL IN UR EMAIL ADDRESS


'IMPROVMENTS:
'***********
'IF UCAN IMPROVE THIS PROGRAM PLEASE SEND ME A COPY
'*************************************************


Dim myappname As String, Subject As String, myemailaddress As String

'EMAIL RELATED VARS
Private DataAvailable As Boolean
Dim inData As String
Private timer As Long
Private change As Boolean
Private Const TIME_OUT = 30

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdSend_Click()
Subject = myappname & cboSubject.List(cboSubject.ListIndex)
'FILL UP OUR EMAIL STRUCTURE
With Myemail
    .From = txtFrom.Text
    .Subject = Subject
    .To = myemailaddress
    .Msg = txtMessage.Text
End With

'NOW WE SEND IT
    cmdSend.Enabled = False
    lblStatus.Caption = "Connecting..."
    Winsock1.Connect Myemail.SMTP, "25"     'Connect to server
    
End Sub

Private Sub Form_Load()

myappname = "Test App Name - "

'change this
'PUT UR EMAIL ADDRESS HERE SO AS THE REPORT CAN BE SNET TO U
myemailaddress = "youremailadress@domain.com"

With cboSubject
    .AddItem " Comment on program"
    .AddItem " Report Bug"
    .AddItem " Program Suggestion"
    .AddItem " Ask a Question"
    .AddItem " Technical Support"
End With

'change this
'EMAIL FORMAT
With Myemail
    .Format = "plain;"
    'YOU WILL HAVE TO FIND A PUBLIC SMTP SERVER
    'I CANT FIND A RELIABLE ONE YET
    .SMTP = "mindspring.com"
End With

End Sub



'WINSOCK SUBS
'************************************************************************
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If Not Number = sckSuccess Then
        MsgBox Description          'Display error
        Timer1.Enabled = False
        CloseConn True
    End If
End Sub



Private Sub Winsock1_DataArrival _
(ByVal bytesTotal As Long)
    Dim data As String
    Winsock1.GetData data, vbString
    'Add data arrived data to the already arrived data
    inData = inData + data
    'Wait till a line is recieved (with CR LF in the end)
    If StrComp(Right$(inData, 2), vbCrLf) = 0 Then DataAvailable = True
End Sub
Private Sub Winsock1_Connect()
    lblStatus.Caption = "Connected"
    timer = 0
    Timer1.Enabled = True
    While Not DataAvailable         'Wait for reply
        If Winsock1.State = sckClosed Then Exit Sub
        DoEvents
    Wend
    Timer1.Enabled = False
    
    Dim reply As String
    Dim tmp() As String
    reply = inData
    inData = ""
    DataAvailable = False
    tmp = Split(reply, " ")
    If Not Val(tmp(0)) = 220 Then           'Error occured
        MsgBox "Server returned the following error:" + vbCrLf + reply
        CloseConn False
        Exit Sub
    End If
    lblStatus.Caption = "Receiving Welcome Message"
    'Start the process
    Winsock1.SendData "HELO " + Winsock1.LocalHostName + vbCrLf
    DoEvents
    timer = 0
    Timer1.Enabled = True
    While Not DataAvailable         'Wait for reply
        If Winsock1.State = sckClosed Then Exit Sub
        DoEvents
    Wend
    Timer1.Enabled = False
    reply = inData
    inData = ""
    DataAvailable = False
    tmp = Split(reply, " ")
    If Not Val(tmp(0)) = 250 Then
        MsgBox "Server returned the following error:" + vbCrLf + reply
        CloseConn False
        Exit Sub
    End If
    'Send MAIL FROM
    Winsock1.SendData "MAIL FROM:<" + Myemail.From + ">" + vbCrLf
    DoEvents
    timer = 0
    Timer1.Enabled = True
    While Not DataAvailable         'Wait for reply
        If Winsock1.State = sckClosed Then Exit Sub
        DoEvents
    Wend
    Timer1.Enabled = False
    reply = inData
    inData = ""
    DataAvailable = False
    tmp = Split(reply, " ")
    If Not Val(tmp(0)) = 250 Then
        MsgBox "Server returned the following error:" + vbCrLf + reply
        CloseConn True
        Exit Sub
    End If
    'Send RCPT TO
    Winsock1.SendData "RCPT TO:<" + Myemail.To + ">" + vbCrLf
    DoEvents
    timer = 0
    Timer1.Enabled = True
    While Not DataAvailable         'Wait for reply
        If Winsock1.State = sckClosed Then Exit Sub
        DoEvents
    Wend
    Timer1.Enabled = False
    reply = inData
    inData = ""
    DataAvailable = False
    tmp = Split(reply, " ")
    If Not Val(tmp(0)) = 250 Then
        MsgBox "Server returned the following error:" + vbCrLf + reply
        CloseConn True
        Exit Sub
    End If
    'Send DATA
    DoEvents
    Winsock1.SendData "DATA" + vbCrLf
    DoEvents
    timer = 0
    Timer1.Enabled = True
    While Not DataAvailable         'Wait for reply
        If Winsock1.State = sckClosed Then Exit Sub
        DoEvents
    Wend
    Timer1.Enabled = False
    reply = inData
    inData = ""
    DataAvailable = False
    tmp = Split(reply, " ")
    If Not Val(tmp(0)) = 354 Then
        MsgBox "Server returned the following error:" + vbCrLf + reply
        CloseConn False
        Exit Sub
    End If
    lblStatus.Caption = "Sending Mail . . ."
    
    
    'Send the E-Mail
    Winsock1.SendData "From: <" + Myemail.From + ">" + vbCrLf + _
                      "To: " + Myemail.To + vbCrLf + _
                      "Subject: " + Myemail.Subject + vbCrLf + _
                      "X-Mailer: Gmcd'sBugReport V1" + vbCrLf + _
                      "Mime-Version: 1.0" + vbCrLf + _
                      "Content-Type: text/" + Myemail.Format + vbTab + "charset=us-ascii" + vbCrLf + vbCrLf + _
                      txtMessage.Text
    Winsock1.SendData vbCrLf + "." + vbCrLf
    
    
    DoEvents
    timer = 0
    Timer1.Enabled = True
    While Not DataAvailable             'Wait for reply
        If Winsock1.State = sckClosed Then Exit Sub
        DoEvents
    Wend
    Timer1.Enabled = False
    reply = inData
    inData = ""
    DataAvailable = False
    tmp = Split(reply, " ")
    If Not Val(tmp(0)) = 250 Then               'Error occured
        MsgBox "Server returned the following error:" + vbCrLf + reply
        CloseConn False
        Exit Sub
    End If
    Winsock1.SendData "QUIT"
    MsgBox "Report Sent Successfully", vbInformation, "Done"
    CloseConn False
End Sub

Private Sub Timer1_Timer()
    timer = timer + 1
    If timer = TIME_OUT Then
        CloseConn True              'Disconnect if timed out
        MsgBox "Could not connect to host " + Myemail.SMTP + vbCrLf + "Operation timed out"
        Timer1.Enabled = False
    End If
End Sub
Private Sub CloseConn(Err As Boolean)
'Close Connection & enable contrls
    Winsock1.Close
    lblStatus.Caption = "Send"
    cmdSend.Enabled = True
End Sub
