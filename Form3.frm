VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form motor 
   BackColor       =   &H0080FF80&
   Caption         =   "Form3"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13500
   LinkTopic       =   "Form3"
   ScaleHeight     =   5400
   ScaleWidth      =   13500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   221
      Left            =   11520
      Top             =   120
   End
   Begin VB.CommandButton Command11 
      Caption         =   "REFRESH"
      Height          =   255
      Left            =   5640
      TabIndex        =   24
      Top             =   2400
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "SEARCH"
      Height          =   975
      Left            =   4680
      TabIndex        =   21
      Top             =   1800
      Width           =   2055
      Begin VB.CommandButton Command10 
         Caption         =   "SEARCH"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "PAKET  1-3 TAHUN"
      Height          =   255
      Left            =   6960
      TabIndex        =   20
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Button"
      Height          =   1215
      Left            =   4680
      TabIndex        =   13
      Top             =   600
      Width           =   4095
      Begin VB.CommandButton Command12 
         Caption         =   "EDIT"
         Height          =   735
         Left            =   2520
         TabIndex        =   25
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "SAVE"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "DELETE"
         Height          =   255
         Left            =   1320
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "EXIT"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "ADD"
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "NEXT"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "PRINT"
         Height          =   255
         Left            =   1320
         TabIndex        =   14
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   4575
      Left            =   8880
      ScaleHeight     =   4515
      ScaleWidth      =   4395
      TabIndex        =   12
      Top             =   600
      Width           =   4455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   11
      Text            =   "Motor"
      Top             =   1200
      Width           =   3015
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form3.frx":0000
      Height          =   2175
      Left            =   240
      TabIndex        =   10
      Top             =   2880
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   3836
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
   Begin VB.CommandButton Command7 
      Caption         =   "CHASH"
      Height          =   255
      Left            =   6960
      TabIndex        =   9
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Paket"
      Height          =   975
      Left            =   6840
      TabIndex        =   8
      Top             =   1800
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3120
      Top             =   720
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
      RecordSource    =   "select*from motor"
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
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   12120
      TabIndex        =   27
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   12120
      TabIndex        =   26
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Merk Motor"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Motor"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "DATA MOTOR"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "motor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_click()
If Combo1 = "YMH-BS-01" Then
Picture1.Picture = LoadPicture("gambar\a.jpg")
Text1 = "Yamaha"
Text2 = "Cowo"
Text3 = "35000000"
ElseIf Combo1 = "YMH-BS-02" Then
Picture1.Picture = LoadPicture("gambar\b.jpg")
Text1 = "Yamaha"
Text2 = "Cowo"
Text3 = "35000000"
ElseIf Combo1 = "YMH-BS-03" Then
Picture1.Picture = LoadPicture("gambar\c.jpg")
Text1 = "Yamaha"
Text2 = "Cowo"
Text3 = "35000000"
ElseIf Combo1 = "YMH-BS-04" Then
Picture1.Picture = LoadPicture("gambar\d.jpg")
Text1 = "Yamaha"
Text2 = "Cowo"
Text3 = "35000000"
ElseIf Combo1 = "YMH-SCP-01" Then
Picture1.Picture = LoadPicture("gambar\e.jpg")
Text1 = "Yamaha"
Text2 = "Cowo"
Text3 = "45000000"
ElseIf Combo1 = "YMH-SCP-02" Then
Picture1.Picture = LoadPicture("gambar\f.jpg")
Text1 = "Yamaha"
Text2 = "Cowo"
Text3 = "45000000"
ElseIf Combo1 = "YMH-SCP-03" Then
Picture1.Picture = LoadPicture("gambar\g.jpg")
Text1 = "Yamaha"
Text2 = "Cowo"
Text3 = "45000000"
ElseIf Combo1 = "YMH-SCP-04" Then
Picture1.Picture = LoadPicture("gambar\h.jpg")
Text1 = "Yamaha"
Text2 = "Cowo"
Text3 = "45000000"
ElseIf Combo1 = "YMH-VR-01" Then
Picture1.Picture = LoadPicture("gambar\i.jpg")
Text1 = "Yamaha"
Text2 = "Cowo-Cewe"
Text3 = "15000000"
ElseIf Combo1 = "YMH-VR-02" Then
Picture1.Picture = LoadPicture("gambar\j.jpg")
Text1 = "Yamaha"
Text2 = "Cowo-Cewe"
Text3 = "15000000"
ElseIf Combo1 = "YMH-VR-03" Then
Picture1.Picture = LoadPicture("gambar\k.jpg")
Text1 = "Yamaha"
Text2 = "Cowo-Cewe"
Text3 = "15000000"
ElseIf Combo1 = "YMH-VR-04" Then
Picture1.Picture = LoadPicture("gambar\l.jpg")
Text1 = "Yamaha"
Text2 = "Cowo-Cewe"
Text3 = "15000000"
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
If Combo1 = "" Or Text1 = "" Or Text2 = "" Or Text3 = "" Then
MsgBox " Masih ada data yang belum terisi!!!", , "error"
Else
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("kode_motor") = Combo1
Adodc1.Recordset.Fields("merk_motor") = Text1
Adodc1.Recordset.Fields("type") = Text2
Adodc1.Recordset.Fields("harga") = Text3
Adodc1.Recordset.Update
MsgBox "Data telah berhasil tersimpan!", vbOKOnly, "Sucses!"
Command1.Enabled = False
End If
End Sub

Private Sub Command10_Click()
Adodc1.RecordSource = "select *from motor where kode_motor like '" & Text5.Text & "'"
Adodc1.Refresh
End Sub

Private Sub Command11_Click()
Adodc1.RecordSource = "select*from motor"
Adodc1.Refresh
End Sub

Private Sub Command12_Click()
On Error Resume Next
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Combo1 = "" Then
MsgBox " Masih ada data yang belum terisi!!!", , "error"
Else
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("kode_motor") = Text1
Adodc1.Recordset.Fields("merk_motor") = Text2
Adodc1.Recordset.Fields("type") = Text3
Adodc1.Recordset.Fields("harga") = Text4
Adodc1.Recordset.Fields("gambar") = Combo1
Adodc1.Recordset.Update
MsgBox "Data telah terganti!", vbOKOnly, "Sucses!"
End If
End Sub

Private Sub Command2_Click()
With Adodc1.Recordset
.Delete
.MoveFirst
End With
End Sub

Private Sub Command3_Click()
psn = MsgBox("apakah anda ingin keluar?", vbYesNo, "info")
If psn = vbYes Then
Unload Me
End If
End Sub

Private Sub Command4_Click()
Combo1 = ""
Text1 = ""
Text2 = ""
Text3 = ""
Text1.SetFocus
End Sub

Private Sub Command6_Click()
ReportMotor.Show
End Sub

Private Sub Command7_Click()
chas.Show
Me.Hide
End Sub

Private Sub Command8_Click()
cicilan1.Show
Me.Hide
End Sub

Private Sub Command9_Click()
credit.Show
Me.Hide
End Sub

Private Sub DataGrid1_Click()
Text1 = Adodc1.Recordset!kode_motor
Text2 = Adodc1.Recordset!merk_motor
Text3 = Adodc1.Recordset!Type
Text4 = Adodc1.Recordset!harga
'Picture1 = Adodc1.Recordset!gambar
End Sub

Private Sub Form_Load()
With Combo1
.AddItem "YMH-BS-01"
.AddItem "YMH-BS-02"
.AddItem "YMH-BS-03"
.AddItem "YMH-BS-04"
.AddItem "YMH-SCP-01"
.AddItem "YMH-SCP-02"
.AddItem "YMH-SCP-03"
.AddItem "YMH-SCP-04"
.AddItem "YMH-VR-01"
.AddItem "YMH-VR-02"
.AddItem "YMH-VR-03"
.AddItem "YMH-VR-04"
End With
End Sub

Private Sub Text4_keypress(keyascii As Integer)
If keyascii = 13 Then
Command1.SetFocus
End If
End Sub

Private Sub Timer1_Timer()
Label6.Caption = Format(Time, "hh.mm.ss.am/pm")
Label7.Caption = Format(Date, "dd-mm-yyyy")
End Sub
