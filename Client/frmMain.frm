VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Sample-Application(Client)"
   ClientHeight    =   2985
   ClientLeft      =   2115
   ClientTop       =   1545
   ClientWidth     =   4740
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   4740
   Begin VB.ListBox lst 
      Height          =   2400
      Left            =   2700
      TabIndex        =   3
      Top             =   60
      Width           =   1995
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   435
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   1035
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock sck 
      Left            =   240
      Top             =   180
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   2700
      Width           =   450
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********'
' ThFabba's Winsock Sample Application V1.15 (for SP5 of the control) - Client
'  Copyright © 2001-2002 Thomas Faber
'
' last edit: 17th Mar 2002
'
' this program is a simple winsock-client program
' the information is written into a listbox
'
' okay... i made this sample some time ago and just made the code _
  a bit clearer in this version.
' Now I don't use the Winsock Control myself in my programs...
' I use Winsock API. If you want to learn using the Winsock API _
  i suggest the very good tutorial by Oleg Gdalevich at http://www.vbip.com
'
' now something about my coding style...
' Option Explicit:
    ' I always use this cause it saves memory, is faster and _
      helps to avoid errors due to mis-spelling
' Comments
    ' At the end of nearly every line in this program it put comments _
      to make this code easier to understand
    ' If there is a comment on a line without code, it _
      describes the following lines or the follwing function
' Indenting
    ' to make my code more readable i put tab characters _
      before lines in blocks(like if/do/open and in this _
      sample also function/sub)
' Variables
    ' I always try to give varibles names that are easy to _
      understand and show what they are for. I also use the _
      hungarian notation mostly to indicate the type of a variable
    ' prefixes i use for vb data types (including controls) in this code:
        ' int   -   Integer
        ' str   -   String
        ' bol   -   Boolean
        ' txt   -   TextBox
        ' lst   -   ListBox
        ' sck   -   Winsock
        ' cmd   -   CommandButton
        ' lbl   -   Label
' ByVal/ByRef
    ' I alwyas write which of them to use although ByRef is standard.
    ' If i don't need to change the value of an argument i use ByVal _
      (except for strings, there i always use byref)
'
' you may do everything(change, copy, ...) with this code but: _
- you are not allowed to make anyone pay for this code _
- if u copy the code you have to credit me as the original creator _
- you are not allowed to remove/change this information
'
' I hope you can improve your winsock-programming skills _
  with the help of this code :)
'
' please give me comments + suggestions!
'
' special thanks to:
    ' hmm! looks like i did this all on my own :)
    ' send me improvements or new features u added _
      so thatll change :)
'
' sorry for all grammatical and orthographical mistakes _
  and for my terrible non-apostrophe and non-capital-letter english :)
'
'
'                                                        Thomas Faber
' Email me: Th-Faber@gmx.net
'
'**********'

'*** Begin of code ***

Option Explicit
Option Base 0
Option Compare Text

'close the connection to the server
Public Sub CloseConnection()
    sck.Close 'close the winsock
    lblStatus.Caption = "Disconnected" 'show that we are no more connected
    cmdConnect.Caption = "&Connect" 'change the button's caption
End Sub
'connect button was clicked
Private Sub cmdConnect_Click()
    If cmdConnect.Caption = "&Connect" Then 'connect to the server
        sck.RemotePort = "1234" 'Set the port
        sck.RemoteHost = InputBox( _
            "Please type the HostName/IP-Address here." & vbNewLine & _
            "You can also run server and client on the same computer." & vbNewLine & _
            "To do that just let the server listen and don't change the text of this Input-Box.", _
            "Enter HostName/IP-Address", sck.LocalIP) 'ask for the remotehost
        sck.Connect 'start the connection
        lblStatus.Caption = "Connecting" 'show that we are now trying to connect
        cmdConnect.Caption = "&Disconnect" 'change the button's caption
    ElseIf cmdConnect.Caption = "&Disconnect" Then 'close the connection to the server
        CloseConnection
    End If
End Sub
'send button was clicked. send the data
Private Sub cmdSend_Click()
    On Local Error GoTo TheError
    sck.SendData txtData.Text 'send the data to the server
    txtData.Text = "" 'clear the textbox
    Exit Sub 'thats it. dont execute the code of the error handler since there was no error
TheError:
    MsgBox Err.Description, vbMsgBoxHelpButton Or _
        vbExclamation, "Error " & Err.Number, Err.HelpFile, Err.HelpContext 'show the error
End Sub
'connection is closed by the server
Private Sub sck_Close()
    CloseConnection
End Sub
'connection is established
Private Sub sck_Connect()
    cmdSend.Enabled = True 'since we are connected now, we can send data
    lblStatus.Caption = "Connected" 'show that we are connected
End Sub
'data from the server is arriving!
Private Sub sck_DataArrival(ByVal bytesTotal As Long)
    Dim strGotData As String 'variable to receive the data
    
    sck.GetData strGotData 'put the data into the string
    lst.AddItem strGotData 'add line of data to the listbox
    '
    ' This is not a good way doing this! but this is not a sample on how to _
      parse data. Maybe itll be in later versions ;)
    '
    lst.ListIndex = lst.ListCount - 1 'select new line in the listbox
End Sub
Private Sub sck_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description, vbMsgBoxHelpButton Or _
        vbExclamation, "Error " & Number, HelpFile, HelpContext 'show the error
End Sub

'*** End of code ***
