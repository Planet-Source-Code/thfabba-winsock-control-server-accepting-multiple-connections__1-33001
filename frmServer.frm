VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   Caption         =   "Sample-Application(Server)"
   ClientHeight    =   3195
   ClientLeft      =   2115
   ClientTop       =   1545
   ClientWidth     =   5520
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close(Kick Client)"
      Enabled         =   0   'False
      Height          =   435
      Left            =   120
      TabIndex        =   10
      Top             =   2220
      Width           =   1215
   End
   Begin VB.ListBox lst 
      Height          =   2790
      Left            =   3120
      TabIndex        =   4
      Top             =   60
      Width           =   2295
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send Data"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   435
      Left            =   1680
      TabIndex        =   3
      Top             =   780
      Width           =   1215
   End
   Begin VB.TextBox txtIndex 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "Index(not 0)"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "Data"
      Top             =   60
      Width           =   1515
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "Start &Listening"
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   780
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock sck 
      Index           =   0
      Left            =   2700
      Top             =   300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Index-TextBox"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   360
      Width           =   1035
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Data-TextBox"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   60
      Width           =   1035
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "With this Button the text in the Data-TextBox is sent to the winsock with the index in the Index-TextBox"
      Height          =   1455
      Index           =   3
      Left            =   1680
      TabIndex        =   7
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Press this button to make this program let clients connect"
      Height          =   855
      Index           =   2
      Left            =   180
      TabIndex        =   6
      Top             =   1320
      Width           =   1155
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty DataFormat 
         Type            =   2
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1031
         SubFormatType   =   9
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   2940
      Width           =   450
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********'
' ThFabba's Winsock Sample Application V1.15 (for SP5 of the control) - Server
'  Copyright © 2001-2002 Thomas Faber
'
' last edit: 17th Mar 2002
'
' this program shows how to make a server that allows _
  multiple(i think up to 32766) connections
' the information is written into a listbox
' start/stop listening, sending/receiving data(text) _
  and kicking clients are supported
'
' okay... i made this sample some time ago and just made the code _
  a bit clearer in this version.
' Now I don't use the Winsock Control myself in my programs...
' I use Winsock api and if you want to make big server applications _
  this will be really necessary cause i heard Winsock control is _
  not very stable. If you want to learn using the winsock api _
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

Private bolWinsockIsUsed() As Boolean 'is that winsock used?
'this function returns a free index for a new client and prepares a winsock
Public Function GetFreeWinSock() As Integer
    Dim intIndex As Integer 'counter variable. will store the index of the new winsock control
    
    For intIndex = 1 To UBound(bolWinsockIsUsed) 'loop thru the array
        If bolWinsockIsUsed(intIndex) = False Then 'is that index used?
            'this index is not used so we can use it now
            GoTo Prepare 'prepare the winsock
        End If
    Next intIndex
    intIndex = UBound(bolWinsockIsUsed) + 1 'if all current indexes are used - add a new one
    ReDim Preserve bolWinsockIsUsed(intIndex) 'redim the array for the new index
Prepare:            'prepare the winsock
    Load sck(intIndex) 'add a new winsock control to the form
    sck(intIndex).Close 'close the winsock to prepare it for a Connection
    sck(intIndex).LocalPort = 0 'give that winsock a random free local port
    bolWinsockIsUsed(intIndex) = True 'that index is used now
    GetFreeWinSock = intIndex 'return the index
End Function
'This function closes the winsock with index I
Public Sub CloseWinsock(ByVal intIndex As Integer)
    sck(intIndex).Close 'close the socket
    bolWinsockIsUsed(intIndex) = False 'the winsock is not used now
    Unload sck(intIndex) 'we wont need that winsock for now
    lst.AddItem "Client(" & CStr(intIndex) & ") disconnected." 'add to the listbox
End Sub
'the close button was clicked. "kick" the selected client
Private Sub cmdClose_Click()
    On Local Error GoTo TheError 'the app shouldn't stop if u typed something wrong into one of the text-boxes
    CloseWinsock CInt(txtIndex.Text) 'close selected winsock
    Exit Sub 'thats it. dont execute the code of the error handler since there was no error
TheError: 'error handler
    MsgBox Err.Description, vbMsgBoxHelpButton Or _
        vbExclamation, "Error " & Err.Number, Err.HelpFile, Err.HelpContext 'show the error
End Sub
'listen button was clicked
Private Sub cmdListen_Click()
    Dim intIndex As Integer 'counter variable
    
    If cmdListen.Caption = "Start &Listening" Then 'we have to start listening
        sck(0).Close 'close the first winsock for a Connection
        sck(0).LocalPort = "1234" 'any number but multiple servers cannot use the same port
        'SAME PORT FOR SERVER AND CLIENTS!
        
        sck(0).Listen 'the first winsock (0) listens and gives all clients to the other winsocks
        lblStatus.Caption = "Listening" 'show that we are listening now
        cmdListen.Caption = "Stop &Listening" 'change the caption of the button
    ElseIf cmdListen.Caption = "Stop &Listening" Then 'we have to stop listening
        For intIndex = 0 To UBound(bolWinsockIsUsed) 'loop thru all winsocks and close them
            If bolWinsockIsUsed(intIndex) = True Then
                sck(intIndex).Close 'close the winsock(stop the connection)
                bolWinsockIsUsed(intIndex) = False 'the winsock isn't used now
            End If
        Next intIndex
        cmdSend.Enabled = False 'disable send button since there cant be a client to send to
        cmdClose.Enabled = False 'disable close button since there cant be a client we could "kick"
        lblStatus.Caption = "Server closed" 'show that the server is closed now
        cmdListen.Caption = "Start &Listening" 'change the caption of the button
    End If
End Sub
'send button was pressed
Private Sub cmdSend_Click()
    On Local Error GoTo TheError 'the app shouldn't stop if u typed something wrong into one of the text-boxes
    sck(txtIndex.Text).SendData txtData.Text 'send the data to the winsock with the index in the textbox
    txtData.Text = vbNullString 'clear textbox now
    Exit Sub 'thats it. dont execute the code of the error handler since there was no error
TheError: 'error handler
    MsgBox Err.Description, vbMsgBoxHelpButton Or _
        vbExclamation, "Error " & Err.Number, Err.HelpFile, Err.HelpContext 'display error
End Sub
'form loads. initialize all the things we need
Private Sub Form_Load()
    ReDim bolWinsockIsUsed(0) 'initialize the array
    bolWinsockIsUsed(0) = False 'first winsock (index 0) listens but you cant send data thru it
End Sub
'an error in the winsock occured
Private Sub sck_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description, vbMsgBoxHelpButton Or _
        vbExclamation, "Error " & Number, HelpFile, HelpContext 'display error
End Sub
'the text of the index textbox changed. lets check if its valid
Private Sub txtIndex_Change()
    Dim intIndex As Integer 'index that was entered
    'checks if the current text of this textbox is a valid index
    On Local Error GoTo AnErrorOcured 'input is invalid
    intIndex = CInt(txtIndex.Text) 'save index written to the textbox in the variable
    If bolWinsockIsUsed(intIndex) = True Then 'input is valid. so enable the buttons
        cmdSend.Enabled = True 'enable send button
        cmdClose.Enabled = True 'enable close button
        Exit Sub 'thats it. dont execute the code of the error handler since there was no error
    End If
AnErrorOcured: 'error handler
    'input is invalid. so disble the buttons
    cmdSend.Enabled = False 'disable send button
    cmdClose.Enabled = False 'disable close button
End Sub
'a key was pressed in the port textbox
Private Sub txtIndex_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub 'backspace - we will allow to use this :)
    If IsNumeric(Chr$(KeyAscii)) = False Then _
        KeyAscii = 0 'don't accept anything else than numbers
End Sub
'u clicked out of the textbox so tha we will remove leading zeros now...
Private Sub txtIndex_LostFocus()
    On Local Error GoTo 0 'dont worry about errors
    txtIndex.Text = CInt(txtIndex.Text) 'convert index to integer value. ie remove leading 0s
End Sub
'a user disconnects
Private Sub sck_Close(Index As Integer)
    CloseWinsock Index 'close the winsock correctly and set all variables right
End Sub
'some1 wants to connect
Private Sub sck_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim intNewIndex As Integer 'index for new client
    
    lblStatus = "Connection requested(" & CStr(requestID) & ")"  'show that some1 wants to connect
    intNewIndex = GetFreeWinSock 'get index for new client
    sck(intNewIndex).Accept requestID 'let the new winsock accept the connection
    '
    ' Here you could check if this user may connect. Maybe ill add this in a _
      future version. Or you could do that and send me ;)
    '
    lblStatus = "Connection from " & sck(intNewIndex).RemoteHostIP & "(" & CStr(requestID) & ") accepted" 'show that the connection was accepted
    lst.AddItem "Client(" & CStr(intNewIndex) & ") connected." 'add to the listbox
End Sub
'we got data
Private Sub sck_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strGotData As String 'variable to receive the data
    
    sck(Index).GetData strGotData 'put the data into the string
    lst.AddItem "Client(" & CStr(Index) & "):" & strGotData 'add line of data to the listbox
    '
    ' This is not a good way doing this! but this is not a sample on how to _
      parse data. Maybe itll be in later versions ;)
    '
    lst.ListIndex = lst.ListCount - 1 'select new line in the listbox
End Sub

'*** End of code ***
