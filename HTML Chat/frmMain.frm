VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HTML Chat"
   ClientHeight    =   3480
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8160
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Client Options"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5160
      TabIndex        =   5
      Top             =   0
      Width           =   3015
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Text            =   "127.0.0.1"
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Disconnect"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Connect"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Connect to IP:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Chat Area"
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton Command10 
         Caption         =   "http://"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         MaskColor       =   &H8000000F&
         TabIndex        =   29
         Top             =   2160
         Width           =   855
      End
      Begin RichTextLib.RichTextBox rtbRichText 
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   2640
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   1296
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmMain.frx":0000
      End
      Begin VB.CommandButton Command8 
         Caption         =   "U"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   23
         Top             =   2160
         Width           =   255
      End
      Begin VB.CommandButton Command7 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   2160
         Width           =   255
      End
      Begin VB.CommandButton Command6 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2160
         Width           =   255
      End
      Begin VB.CommandButton Send 
         Caption         =   "Send"
         Height          =   735
         Left            =   3960
         TabIndex        =   4
         Top             =   2640
         Width           =   975
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   1695
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4815
         ExtentX         =   8493
         ExtentY         =   2990
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin RichTextLib.RichTextBox receivebox 
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   2640
         Visible         =   0   'False
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   1296
         _Version        =   393217
         TextRTF         =   $"frmMain.frx":0082
      End
      Begin RichTextLib.RichTextBox txtHTML 
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   2640
         Visible         =   0   'False
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   1296
         _Version        =   393217
         TextRTF         =   $"frmMain.frx":0104
      End
      Begin RichTextLib.RichTextBox txtBlank 
         Height          =   735
         Left            =   120
         TabIndex        =   17
         Top             =   2640
         Visible         =   0   'False
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   1296
         _Version        =   393217
         TextRTF         =   $"frmMain.frx":0186
      End
      Begin RichTextLib.RichTextBox sendbox 
         Height          =   735
         Left            =   120
         TabIndex        =   28
         Top             =   2640
         Visible         =   0   'False
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   1296
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmMain.frx":0208
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   4920
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   4920
         Y1              =   2040
         Y2              =   2040
      End
   End
   Begin MSWinsockLib.Winsock WS 
      Left            =   240
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Caption         =   "Chat Options:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5160
      TabIndex        =   18
      Top             =   1080
      Width           =   3015
      Begin VB.CommandButton Command9 
         Caption         =   "Exit"
         Height          =   615
         Left            =   1800
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtNick 
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Text            =   "Nickname"
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "YourChat Nick:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Server Options"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5160
      TabIndex        =   6
      Top             =   2040
      Width           =   3015
      Begin VB.ListBox List1 
         Height          =   840
         ItemData        =   "frmMain.frx":028A
         Left            =   120
         List            =   "frmMain.frx":028C
         TabIndex        =   14
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Restart"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Stop Server"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Start Server"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Connected IPs:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSWinsockLib.Winsock WS2 
      Left            =   240
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lastcount 
      Caption         =   "0"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "Client Is Connected"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   26
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Server Is Running"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   27
      Top             =   600
      Width           =   1815
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function RichToHTML(rtbRichTextBox As RichTextLib.RichTextBox, Optional lngStartPosition As Long, Optional lngEndPosition As Long) As String
'The Conversion from RichText to HTML
Dim blnBold As Boolean, blnUnderline As Boolean, blnStrikeThru As Boolean
Dim blnItalic As Boolean, strLastFont As String, lngLastFontColor As Long
Dim strHTML As String, lngColor As Long, lngRed As Long, lngGreen As Long
Dim lngBlue As Long, lngCurText As Long, strHex As String, intLastAlignment As Integer
Const AlignLeft = 0, AlignRight = 1, AlignCenter = 2
If IsMissing(lngStartPosition&) Then lngStartPosition& = 0
If IsMissing(lngEndPosition&) Then lngEndPosition& = Len(rtbRichTextBox.Text)
lngLastFontColor& = -1
   For lngCurText& = lngStartPosition& To lngEndPosition&
       rtbRichTextBox.SelStart = lngCurText&
       rtbRichTextBox.SelLength = 1
   
          If intLastAlignment% <> rtbRichTextBox.SelAlignment Then
             intLastAlignment% = rtbRichTextBox.SelAlignment

                Select Case rtbRichTextBox.SelAlignment
                   Case AlignLeft: strHTML$ = strHTML$ & "<p align=left>"
                   Case AlignRight: strHTML$ = strHTML$ & "<p align=right>"
                   Case AlignCenter: strHTML$ = strHTML$ & "<p align=center>"
                End Select
                
          End If

          If blnBold <> rtbRichTextBox.SelBold Then
               If rtbRichTextBox.SelBold = True Then
                 strHTML$ = strHTML$ & "<b>"
               Else
                 strHTML$ = strHTML$ & "</b>"
               End If
             blnBold = rtbRichTextBox.SelBold
          End If

          If blnUnderline <> rtbRichTextBox.SelUnderline Then
               If rtbRichTextBox.SelUnderline = True Then
                 strHTML$ = strHTML$ & "<u>"
               Else
                 strHTML$ = strHTML$ & "</u>"
               End If
             blnUnderline = rtbRichTextBox.SelUnderline
          End If
   

          If blnItalic <> rtbRichTextBox.SelItalic Then
               If rtbRichTextBox.SelItalic = True Then
                 strHTML$ = strHTML$ & "<i>"
               Else
                 strHTML$ = strHTML$ & "</i>"
               End If
             blnItalic = rtbRichTextBox.SelItalic
          End If


          If blnStrikeThru <> rtbRichTextBox.SelStrikeThru Then
               If rtbRichTextBox.SelStrikeThru = True Then
                 strHTML$ = strHTML$ & "<s>"
               Else
                 strHTML$ = strHTML$ & "</s>"
               End If
             blnStrikeThru = rtbRichTextBox.SelStrikeThru
          End If

         If strLastFont$ <> rtbRichTextBox.SelFontName Then
            strLastFont$ = rtbRichTextBox.SelFontName
            strHTML$ = strHTML$ + "<font face=""" & strLastFont$ & """>"
         End If

         If lngLastFontColor& <> rtbRichTextBox.SelColor Then
            lngLastFontColor& = rtbRichTextBox.SelColor
            strHex$ = Hex(rtbRichTextBox.SelColor)
            strHex$ = String$(6 - Len(strHex$), "0") & strHex$
            strHex$ = Right$(strHex$, 2) & Mid$(strHex$, 3, 2) & Left$(strHex$, 2)
            
            strHTML$ = strHTML$ + "<font color=#" & strHex$ & ">"
        End If
 
     strHTML$ = strHTML$ + rtbRichTextBox.SelText

   Next lngCurText&

RichToHTML = strHTML$

End Function

Private Sub Command1_Click()
'Change your nickname to client, mostly for me testing and not wanting to change
'the nick all the time when I restarted the program
txtNick.Text = "Client"
'Get rid of the server controls because we do not need them if we are a client
Frame3.Visible = False
'Make sure we can not click connect because we are already connected
Command1.Enabled = False
'Make it so we can click disconnect if we want to disconnect from the server
Command2.Enabled = True
'Set the port of the server that we are going to connect to
WS2.RemotePort = 789
'Tell the client where the server is located that we want to connect to
WS2.RemoteHost = Text1
'Time to connect to the server, let's do it
WS2.Connect
End Sub

Private Sub Command10_Click()
'Show the Insert Link form so we can put a link in our message
frmLink.Show
End Sub

Private Sub Command2_Click()
'Okay, we are disconnecting so we want to show the server controls again
Frame3.Visible = True
'If we are disconnected we will need the connect button if we want to connect again
Command1.Enabled = True
'We do not need to disconnect because we are already disconnected
Command2.Enabled = False
'Close the connection to the server
WS2.Close
End Sub

Private Sub Command3_Click()
'Change your nickname to server, mostly for me testing and not wanting to change
'the nick all the time when I restarted the program
txtNick.Text = "Server"
'Get rid of the client controls because we will not be needing them as the server
Frame2.Visible = False
'The server is started, we do not need to start it again, so why not disable the button
Command3.Enabled = False
'We started the server, if we want to stop it we need to enable the button to do so
Command4.Enabled = True
'If we want to restart the server, let's make sure the button is enabled so we can
Command5.Enabled = True
'Set the port that we will be listening to for the client to connect to
WS.LocalPort = 789
'Okay, start listening for the client to connect
WS.Listen
End Sub

Private Sub Command4_Click()
'We are not a server anymore, we may want to be a client so we make the controls visible
Frame2.Visible = True
'We stopped the server so we need to make it possible to start again by enabling the button
Command3.Enabled = True
'We do not need the button to stop the server because it is already stopped
'let's get rid of it for now
Command4.Enabled = False
'If we are not the server it will not be possible to restart so let's get rid of the button
Command5.Enabled = False
'Close the connection so we will stop listening for the client
WS.Close
End Sub

Private Sub Command5_Click()
'Restarting the server by closing the winsock, re-opening it and setting the port
WS.Close
WS.LocalPort = 789
WS.Listen
End Sub

Private Sub Command6_Click()
'If the text is not bold, make it bold, if it is bold already make it not bold
If rtbRichText.SelBold = False Then
   rtbRichText.SelBold = True
Else
   rtbRichText.SelBold = False
End If
End Sub

Private Sub Command7_Click()
'If the text is not italic, make it italic, if it is italic already make it not italic
If rtbRichText.SelItalic = False Then
   rtbRichText.SelItalic = True
Else
   rtbRichText.SelItalic = False
End If
End Sub

Private Sub Command8_Click()
'If the text is not underline, make it underline, if it is underline already make it not underline
If rtbRichText.SelUnderline = False Then
   rtbRichText.SelUnderline = True
Else
   rtbRichText.SelUnderline = False
End If
End Sub

Private Sub Command9_Click()
'Tell the program we do not need it anymore
Unload Me
End Sub

Private Sub Form_Load()
'Just in case, to prevent a halt on an error
On Error Resume Next
DoEvents
'We want to create a new HTML file so that we can use it for our messages
Open App.Path & "\temp.html" For Output As #1: Print #1, txtBlank.Text: Close #1
DoEvents
'We want to load the blank HTML file into the messenger so that there is no
'page not found error
WebBrowser1.Navigate App.Path & "\temp.html"
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub Send_Click()
Dim strData As String
'We have an incriment in the HTML file so that you load the HTML file at the last message instead of the top
lastcount.Caption = lastcount.Caption + 1
'Convert the rich text to HTML and put it in your box, make sure the name is red
txtHTML.Text = "<a name='last" & lastcount.Caption & "'></a><font color='red'><b>" & txtNick.Text & ": " & "</b></font>" & RichToHTML(rtbRichText, 0&, Len(rtbRichText.Text)) & "</font></font></b></u></i><br>"
'Convert the rich text to HTML that you are sending, make sure the name is blue
sendbox.Text = "<a name='last" & lastcount.Caption & "'></a><font color='blue'><b>" & txtNick.Text & ": " & "</b></font>" & RichToHTML(rtbRichText, 0&, Len(rtbRichText.Text)) & "</font></font></b></u></i><br>"
'Set the name of your message so the server knows what to send
strData = sendbox.Text
'I have to put this here to prevent a crash
On Error Resume Next
'Send the message to the Client
WS.SendData strData
'Send the message to the Server
WS2.SendData strData
'Set the send text box to blank so you don't have to
rtbRichText.Text = txtBlank.Text
End Sub

Private Sub txtHTML_Change()
'Just in case, to prevent a halt on an error
On Error Resume Next
DoEvents
'When we send our message it is converted from rich text to HTML, now we want to add it
'to our HTML file so that we can see it
Open App.Path & "\temp.html" For Append As #1: Print #1, txtHTML.Text: Close #1
DoEvents
'Reload our HTML file into the messenger so that we can read our message we just added
WebBrowser1.Navigate App.Path & "\temp.html#last" & lastcount.Caption
End Sub

Private Sub rtbRichText_KeyPress(KeyAscii As Integer)
'If we are typing and we press the enter key we want it to assume that we are sending
'our message so we make it click the send button for us when we press enter
If KeyAscii = (13) Then
Send_Click
End If
End Sub

Private Sub WS2_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
'Watch for a new message from the server
WS2.GetData strData
'We first want to put the message in a text box so we can add it to the HTML file
receivebox = strData
'Incriment our counter so that the HTML file will load to this message
lastcount.Caption = lastcount.Caption + 1
'Just in case, to prevent a halt on an error
On Error Resume Next
DoEvents
'We want to save the new message to the HTML file so we can load it into the messenger
'we are appending so it adds the message instead of overwriting what we already have
Open App.Path & "\temp.html" For Append As #1: Print #1, receivebox.Text: Close #1
DoEvents
'We are reloading the chat area so that we can see the new message
WebBrowser1.Navigate App.Path & "\temp.html#last" & lastcount.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Get rid of the HTML file off of the system when we close the program
FileSystem.Kill App.Path & "\temp.html"
End Sub

Private Sub WS_ConnectionRequest(ByVal requestID As Long)
'If we are the server, we can see the IP of the connected client
List1.AddItem WS.RemoteHostIP
'Make sure that we are not using our winsock already, or else the client will crash
WS.Close
'Let the client connect
WS.Accept requestID
End Sub

Private Sub WS_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
'We are the server, and the client is sending us a message
WS.GetData strData
'We first want to put the message in a text box so we can add it to the HTML file
receivebox = strData
'Incriment our counter so that the HTML file will load to this message
lastcount.Caption = lastcount.Caption + 1
'Just in case, to prevent a halt on an error
On Error Resume Next
DoEvents
'We want to save the new message to the HTML file so we can load it into the messenger
'we are appending so it adds the message instead of overwriting what we already have
Open App.Path & "\temp.html" For Append As #1: Print #1, receivebox.Text: Close #1
DoEvents
'We are reloading the chat area so that we can see the new message
WebBrowser1.Navigate App.Path & "\temp.html#last" & lastcount.Caption
End Sub

Private Sub WS_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'If there is an error with the connection this tells us what it is (Server Side)
Call MsgBox(Description, vbExclamation, "Error Num." & Number)
End Sub

Private Sub WS2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'If there is an error with the connection this tells us what it is (Client Side)
Call MsgBox(Description, vbExclamation, "Error Num." & Number)
End Sub
