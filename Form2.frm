VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form pembeli 
   BackColor       =   &H0080FF80&
   Caption         =   "Form2"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8790
   LinkTopic       =   "Form2"
   ScaleHeight     =   5055
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   221
      Left            =   1440
      Top             =   4440
   End
   Begin VB.CommandButton Command8 
      Caption         =   "REFRESH"
      Height          =   255
      Left            =   6840
      TabIndex        =   22
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1095
      Left            =   4560
      TabIndex        =   19
      Top             =   1320
      Width           =   3855
      Begin VB.CommandButton Command7 
         Caption         =   "SEARCH"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Button"
      Height          =   1095
      Left            =   4560
      TabIndex        =   12
      Top             =   120
      Width           =   3855
      Begin VB.CommandButton Command9 
         Caption         =   "EDIT"
         Height          =   735
         Left            =   2640
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "SAVE"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ADD"
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "DELETE"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "EXIT"
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "NEXT"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "PRINT"
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form2.frx":0000
      Height          =   1695
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   2990
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3240
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ukom.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ukom.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select*from pembeli"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   1440
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "No Telp"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "No Hp"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Lengkap"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "No KTP"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "data pembeli"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "pembeli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Then
MsgBox "Masih ada data yang belum terisi!", , "error"
Else
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("no_ktp") = Text1
Adodc1.Recordset.Fields("nama_lengkap") = Text2
Adodc1.Recordset.Fields("alamat") = Text3
Adodc1.Recordset.Fields("no_hp") = Text4
Adodc1.Recordset.Fields("no_telp") = Text5
Adodc1.Recordset.Update
MsgBox "Data telah berhasil tersimpan", vbOKOnly, "Sucses!"
Command1.Enabled = False
End If
End Sub

Private Sub Command2_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text1.SetFocus
End Sub

Private Sub Command3_Click()
With Adodc1.Recordset
.Delete
.MoveFirst
End With
End Sub

Private Sub Command4_Click()
psn = MsgBox("apakah anda ingin keluar?", vbYesNo, "info")
If psn = vbYes Then
Unload Me
End If
End Sub

Private Sub Command5_Click()
motor.Show
Me.Hide
End Sub

Private Sub Command6_Click()
ReportPembeli.Show
End Sub

Private Sub Command7_Click()
Adodc1.RecordSource = "select*from pembeli where no_ktp like '" & Text6.Text & "'"
Adodc1.Refresh
End Sub

Private Sub Command8_Click()
Adodc1.RecordSource = "select *from pembeli"
Adodc1.Refresh
End Sub

Private Sub Command9_Click()
On Error Resume Next
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Then
MsgBox "Masih ada data yang belum terisi!", , "error"
Else
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("no_ktp") = Text1
Adodc1.Recordset.Fields("nama_lengkap") = Text2
Adodc1.Recordset.Fields("alamat") = Text3
Adodc1.Recordset.Fields("no_hp") = Text4
Adodc1.Recordset.Fields("no_telp") = Text5
Adodc1.Recordset.Update
MsgBox "Data telah terganti", vbOKOnly, "Sucses!"
End If
End Sub

Private Sub DataGrid1_Click()
Text1 = Adodc1.Recordset!no_ktp
Text2 = Adodc1.Recordset!nama_lengkap
Text3 = Adodc1.Recordset!alamat
Text4 = Adodc1.Recordset!no_hp
Text5 = Adodc1.Recordset!no_telp
End Sub

Private Sub Text1_keypress(keyascii As Integer)
If keyascii = 13 Then
Text2.SetFocus
End If
End Sub

Private Sub Text2_keypress(keyascii As Integer)
If keyascii = 13 Then
Text3.SetFocus
End If
End Sub

Private Sub Text3_keypress(keyascii As Integer)
If keyascii = 13 Then
Text4.SetFocus
End If
End Sub

Private Sub Text4_keypress(keyascii As Integer)
If keyascii = 13 Then
Text5.SetFocus
End If
End Sub

Private Sub Text5_keypress(keyascii As Integer)
If keyascii = 13 Then
Command1.SetFocus
End If
End Sub

Private Sub Timer1_Timer()
Label8.Caption = Format(Time, "hh.mm.ss.am/pm")
Label7.Caption = Format(Date, "dd-mm-yyyy")
End Sub
