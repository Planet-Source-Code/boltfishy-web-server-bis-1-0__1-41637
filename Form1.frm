VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Winsock1(0).LocalIP - Stopped"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtfData 
      Height          =   1215
      Left            =   4560
      TabIndex        =   20
      Top             =   5160
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   2143
      _Version        =   393217
      TextRTF         =   $"Form1.frx":0442
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   6480
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6165
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Config"
      TabPicture(0)   =   "Form1.frx":04B9
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdRun"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdStop"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Stats"
      TabPicture(1)   =   "Form1.frx":04D5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(1)=   "Label4"
      Tab(1).Control(2)=   "Image2"
      Tab(1).Control(3)=   "Frame5"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Errors"
      TabPicture(2)   =   "Form1.frx":04F1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label5"
      Tab(2).Control(1)=   "Label6"
      Tab(2).Control(2)=   "Image3"
      Tab(2).Control(3)=   "Frame4"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Options"
      TabPicture(3)   =   "Form1.frx":050D
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label9"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label10"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Image4"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Frame6"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "About"
      TabPicture(4)   =   "Form1.frx":0529
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Image5"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label11"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label12"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Frame7"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      Begin VB.Frame Frame4 
         Caption         =   "Error Path"
         Height          =   735
         Left            =   -74880
         TabIndex        =   31
         Top             =   1080
         Width           =   4215
         Begin VB.TextBox txtError 
            Height          =   375
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Options"
         Height          =   2175
         Left            =   -74880
         TabIndex        =   26
         Top             =   1080
         Width           =   4215
         Begin VB.CheckBox chkMin 
            Caption         =   "Start Minimized"
            Height          =   330
            Left            =   240
            TabIndex        =   30
            ToolTipText     =   "Should BIS go straight to the System Tray when launched?"
            Top             =   240
            Width           =   2895
         End
         Begin VB.CheckBox chkStart 
            Caption         =   "Run Server Automatically"
            Height          =   330
            Left            =   240
            TabIndex        =   29
            ToolTipText     =   "Automatically starts the server upon the launch of BIS"
            Top             =   600
            Width           =   2895
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save Settings"
            Height          =   375
            Left            =   240
            TabIndex        =   28
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CheckBox chkSaveLog 
            Caption         =   "Save Log File"
            Height          =   285
            Left            =   240
            TabIndex        =   27
            ToolTipText     =   "Saves a log file of all website activity"
            Top             =   960
            Width           =   2295
         End
         Begin VB.Timer tmrSave 
            Interval        =   19999
            Left            =   3720
            Top             =   240
         End
         Begin VB.Label Label13 
            Caption         =   "[settings are automatically saved on exit]"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   1320
            Width           =   3015
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "About"
         Height          =   975
         Left            =   -74880
         TabIndex        =   21
         Top             =   1080
         Width           =   4215
         Begin VB.Label Label7 
            Caption         =   "by Mischa Balen"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   360
            Width           =   3855
         End
         Begin VB.Label Label8 
            Caption         =   "http://www26.brinkster.com/boltfish/"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   600
            Width           =   3855
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Recently Logged Activity"
         Height          =   2175
         Left            =   -74880
         TabIndex        =   16
         Top             =   1080
         Width           =   4215
         Begin VB.TextBox txtLog 
            Height          =   1815
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Port"
         Height          =   735
         Left            =   2280
         TabIndex        =   9
         Top             =   2040
         Width           =   2055
         Begin VB.TextBox txtPort 
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Text            =   "80"
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   2880
         Width           =   855
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "Run"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   855
      End
      Begin VB.Frame Frame2 
         Caption         =   "Index File"
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Width           =   2055
         Begin VB.CommandButton cmdBrowse2 
            Caption         =   "..."
            Height          =   375
            Left            =   1440
            TabIndex        =   11
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtIndex 
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Website Path"
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   4215
         Begin VB.TextBox txtPath 
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.Label Label12 
         Caption         =   "Boltfish Internet Services 2002"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74040
         TabIndex        =   25
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Label11 
         Caption         =   "Roxfort Studios"
         Height          =   255
         Left            =   -74160
         TabIndex        =   24
         Top             =   480
         Width           =   3495
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   -74760
         Picture         =   "Form1.frx":0545
         Top             =   480
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   -74760
         Picture         =   "Form1.frx":0987
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label10 
         Caption         =   "Roxfort Studios"
         Height          =   255
         Left            =   -74160
         TabIndex        =   19
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label9 
         Caption         =   "Boltfish Internet Services 2002"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74040
         TabIndex        =   18
         Top             =   720
         Width           =   3375
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   -74760
         Picture         =   "Form1.frx":0DC9
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label6 
         Caption         =   "Roxfort Studios"
         Height          =   255
         Left            =   -74160
         TabIndex        =   15
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label5 
         Caption         =   "Boltfish Internet Services 2002"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74040
         TabIndex        =   14
         Top             =   720
         Width           =   3375
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   -74760
         Picture         =   "Form1.frx":120B
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label4 
         Caption         =   "Roxfort Studios"
         Height          =   255
         Left            =   -74160
         TabIndex        =   13
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   "Boltfish Internet Services 2002"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74040
         TabIndex        =   12
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "Boltfish Internet Services 2002"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "Roxfort Studios"
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   480
         Width           =   3495
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "Form1.frx":164D
         Top             =   480
         Width           =   480
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   5880
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuRun 
         Caption         =   "Run Server"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Stop Server"
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBrowse 
         Caption         =   "Choose Location"
      End
      Begin VB.Menu mnuVisit 
         Caption         =   "Visit Website"
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'by Mischa Balen
'some source 'borrowed'


Dim Connections As Long
Dim m As Long 'users


Private Sub chkSaveLog_Click()
If chkSaveLog.Value = 1 Then
SaveLog = 1
End If
End Sub

Private Sub cmdBrowse2_Click()

Dim File1 As String
CD1.ShowOpen
File1 = FreeFile
    
    If CD1.FileName <> "" Then 'if file name is true
        File1 = CD1.FileName 'return file path
            ElseIf CD1.FileName = "" Then
            File1 = ""
        Exit Sub
    End If
    
    If FileExists(File1) = True Then 'if file exists
        txtIndex.Text = GetFileName(File1) 'add it to text box
        txtPath.Text = GetPath(File1)
        
            ElseIf FileExists(File1) = False Then
                File1 = ""
            CD1.FileName = ""
        Exit Sub
    End If

End Sub

Private Sub cmdRun_Click() 'run server

cmdRun.Enabled = False: cmdStop.Enabled = True: cmdBrowse2.Enabled = False: txtPort.Enabled = False: txtIndex.Enabled = False

Me.Caption = Winsock1(0).LocalIP & " - Running"

Winsock1(0).LocalPort = txtPort.Text
Winsock1(0).Listen

End Sub

Private Sub cmdSave_Click() 'save config settings
Dim sFile As String
sFile = Space(256)

sFile = App.Path
sFile = sFile + "\bis.ini"

WritePrivateProfileString "BIS Config", "Root", txtPath.Text, sFile
WritePrivateProfileString "BIS Config", "Index", txtIndex.Text, sFile
WritePrivateProfileString "BIS Config", "Port", txtPort.Text, sFile
WritePrivateProfileString "BIS Config", "Minimized", CStr(chkMin.Value), sFile
WritePrivateProfileString "BIS Config", "AutoStart", CStr(chkStart.Value), sFile
WritePrivateProfileString "BIS Config", "SaveLog", CStr(chkSaveLog.Value), sFile
WritePrivateProfileString "BIS Config", "ErrorPath", txtError.Text, sFile

End Sub

Private Sub cmdStop_Click() 'stop server

cmdRun.Enabled = True: cmdStop.Enabled = False: cmdBrowse2.Enabled = True: txtPort.Enabled = True: txtIndex.Enabled = True

Me.Caption = Winsock1(0).LocalIP & " - Stopped"

Winsock1(0).Close

End Sub


Private Sub Form_Load()

Dim sFile As String
sFile = Space(256)

sFile = App.Path
sFile = sFile + "\bis.ini"

        
        With nid 'with system tray
            .cbSize = Len(nid)
            .hwnd = Me.hwnd
            .uId = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            .uCallBackMessage = WM_MOUSEMOVE
            .hIcon = Me.Icon 'use form's icon in tray
            .szTip = "BIS " & App.Major & "." & App.Minor & "." & App.Revision & vbNullChar
        End With
        
    Shell_NotifyIcon NIM_ADD, nid 'add to tray
    
Me.Caption = Winsock1(0).LocalIP & " - Stopped"

If FileExists(App.Path & "\bis.ini") = False Then
MsgBox ("Deafult configuration being loaded!"), vbInformation + vbOKOnly, "BIS"
WritePrivateProfileString "BIS Config", "Root", App.Path & "\", sFile
WritePrivateProfileString "BIS Config", "Index", "index.htm", sFile
WritePrivateProfileString "BIS Config", "Port", "80", sFile
WritePrivateProfileString "BIS Config", "Minimized", "0", sFile
WritePrivateProfileString "BIS Config", "AutoStart", "0", sFile
WritePrivateProfileString "BIS Config", "SaveLog", "1", sFile
WritePrivateProfileString "BIS Config", "ErrorPath", App.Path & "\errors", sFile
End If

'let's read the ini file and load our settings

Root = Space(256)
Root = Left(Root, GetPrivateProfileString("BIS Config", "Root", "NULL", Root, Len(Root), sFile))
Index = Space(64)
Index = Left(Index, GetPrivateProfileString("BIS Config", "Index", "NULL", Index, Len(Index), sFile))
Port = GetPrivateProfileInt("BIS Config", "Port", 0, sFile)
AutoStart = GetPrivateProfileInt("BIS Config", "AutoStart", 0, sFile)
Minimized = GetPrivateProfileInt("BIS Config", "Minimized", 0, sFile)
SaveLog = GetPrivateProfileInt("BIS Config", "SaveLog", 0, sFile)
ErrorPath = Space(256)
ErrorPath = Left(ErrorPath, GetPrivateProfileString("BIS Config", "ErrorPath", "NULL", ErrorPath, Len(ErrorPath), sFile))


txtPath.Text = Root
txtIndex.Text = Index
txtPort.Text = Port
txtError.Text = ErrorPath

If AutoStart = 1 Then 'if we should run server
    chkStart.Value = 1 'set check box
    Call cmdRun_Click 'start server
End If

If Minimized = 1 Then 'if we should minimize the prog
    chkMin.Value = 1 'set check box
    Me.Hide 'and hide prog
End If

If SaveLog = 1 Then 'if we should save a log file
    chkSaveLog.Value = 1 'set check box
End If


End Sub

Private Sub Form_Resize()

If WindowState = 1 And Visible = True Then
    Me.Hide
End If

End Sub

Private Sub Form_Unload(Cancel As Integer) 'on form unload
    Call cmdSave_Click 'save settings
    Winsock1(m).Close
    Shell_NotifyIcon NIM_DELETE, nid 'remove from tray
    Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Result, Action As Long
    
    'there are two display modes and we need to find out
    'which one the application is using
    
    If Me.ScaleMode = vbPixels Then
        Action = X
    Else
        Action = X / Screen.TwipsPerPixelX
    End If
    
Select Case Action

    Case WM_LBUTTONDBLCLK 'Left Button Double Click
        Me.WindowState = vbNormal 'put into taskbar
            Result = SetForegroundWindow(Me.hwnd)
        Me.Show 'show form
    
    Case WM_RBUTTONUP 'Right Button Up
        Result = SetForegroundWindow(Me.hwnd)
        PopupMenu mnuFile 'popup menu, cool eh?
    
    End Select
    
End Sub

Private Sub mnuBrowse_Click()
Call cmdBrowse2_Click
End Sub

Private Sub mnuExit_Click()
    Unload Me
    Winsock1(m).Close
    Shell_NotifyIcon NIM_DELETE, nid 'remove from tray
    End
End Sub

Private Sub mnuRun_Click()
Call cmdRun_Click
End Sub

Private Sub mnuStop_Click()
Call cmdStop_Click
End Sub


Private Sub tmrSave_Timer()

If SaveLog = 1 Then

Open App.Path + "\" + "log.txt" For Append As #1 'open the file
    Print #1, txtLog.Text
        txtLog.Text = ""
            'set the contents of it to the text box
                Close #1 'close the file

End If

End Sub

'WINSOCK STUFF

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)

If Winsock1(m).State <> sckckosed Then Winsock1(m).Close

Winsock1(m).Accept requestID

Connections = Connections + 1

End Sub


Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)

On Error GoTo ErrSub

Winsock1(m).GetData Data, vbString, bytesTotal

Request = Mid(Data, 5, InStr(5, Data, " ") - 5)
If Request <> "/" Then
'/.. Protection
If InStr(1, Request, "/..", vbTextCompare) Then rtfData.LoadFile ErrorPath + "413.htm": Winsock1(m).SendData rtfData.Text: Exit Sub

End If

txtLog.Text = txtLog.Text & Data
myData = Split(Data, " ", -1, vbBinaryCompare)

If myData(1) = "/" Then
    rtfData.LoadFile txtPath & txtIndex
        DoEvents
            Winsock1(m).SendData rtfData.Text
Else
    rtfData.LoadFile txtPath & Right(myData(1), Len(myData(1)) - 1)
        DoEvents
            Winsock1(m).SendData rtfData.Text
End If

ErrSub:

On Error Resume Next
If Err.Number <> 0 Then
If Err.Number = 75 Then rtfData.LoadFile ErrorPath + "404.htm": Winsock1(m).SendData rtfData.Text

End If


End Sub

Private Sub Winsock1_SendComplete(Index As Integer)
Winsock1(m).Close
Winsock1(m).Listen
End Sub

