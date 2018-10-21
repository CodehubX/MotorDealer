VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form login 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   3525
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4035
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3525
   ScaleWidth      =   4035
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   120
      Top             =   3000
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   3000
      Top             =   3000
   End
   Begin VB.Timer Timer1 
      Interval        =   221
      Left            =   3480
      Top             =   3000
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   2160
      Picture         =   "Form1.frx":1C59
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Log In"
      Height          =   615
      Left            =   120
      Picture         =   "Form1.frx":71FC
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1080
      TabIndex        =   4
      Text            =   "Pilih"
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1800
      Width           =   2775
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   2880
      Width           =   615
      URL             =   "D:\IBAS\JANGAN DI BUKA!!!\VISUAL BASIC\VB UKOM IBAS FIX\Avenged Sevenfold - Gunslinger.mp3"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1085
      _cy             =   873
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Selamat datang di Dealer Motor yamaha."
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Combo1 = "Admin" And Text1 = "Myadmin" Then
sstab.Show
Me.Hide
ElseIf Combo1 = "User" And Text1 = "Myuser" Then
MDIForm1.Show
Me.Hide
Text1.Enabled = False
Else
MsgBox "Password yang anda masukan salah!", vbYesNo, "info"
Text1 = ""
Text1.SetFocus
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
With Combo1
.AddItem "Admin"
.AddItem "User"
End With
End Sub

Private Sub Text1_keypress(keyascii As Integer)
If keyascii = 13 Then
Command1.SetFocus
End If
End Sub

Private Sub Timer1_Timer()
Label4.Caption = Format(Time, "hh.mm.ss.am/pm")
Label5.Caption = Format(Date, "dd-mm-yyyy")
End Sub

Private Sub Timer2_Timer()
If ProgressBar1.Value = 100 Then
ProgressBar1.Visible = False
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Text1.Visible = True
Combo1.Visible = True
Command1.Visible = True
Command2.Visible = True
Else
ProgressBar1.Value = Val(ProgressBar1.Value) + Val(1)
Label2.Visible = False
Label3.Visible = False
Text1.Visible = False
Combo1.Visible = False
Command1.Visible = False
Command2.Visible = False
End If
End Sub

Private Sub Timer3_Timer()
Label1.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End Sub
