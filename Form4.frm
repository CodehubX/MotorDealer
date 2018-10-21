VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form chas 
   BackColor       =   &H0080FF80&
   Caption         =   "Form4"
   ClientHeight    =   5295
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9840
   LinkTopic       =   "Form4"
   ScaleHeight     =   5295
   ScaleWidth      =   9840
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   221
      Left            =   7920
      Top             =   120
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   29
      Text            =   "Terima Kasih Telah Membeli Produk Kami, Semoga Anda Puas Dengan Pelayanan Kami."
      Top             =   3000
      Width           =   8295
   End
   Begin VB.CommandButton Command8 
      Caption         =   "REFRESH"
      Height          =   375
      Left            =   6960
      TabIndex        =   26
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      Height          =   975
      Left            =   4560
      TabIndex        =   23
      Top             =   1920
      Width           =   3615
      Begin VB.CommandButton Command10 
         Caption         =   "LOGOUT"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton Command7 
         Caption         =   "SEARCH"
         Height          =   375
         Left            =   2400
         TabIndex        =   25
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Button"
      Height          =   1095
      Left            =   4560
      TabIndex        =   16
      Top             =   840
      Width           =   3615
      Begin VB.CommandButton Command9 
         Caption         =   "EDIT"
         Height          =   735
         Left            =   2520
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "SAVE"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "DELETE"
         Height          =   255
         Left            =   1320
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "ADD"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "EXIT"
         Height          =   255
         Left            =   1320
         TabIndex        =   19
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "NEXT"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "PRINT"
         Height          =   255
         Left            =   1320
         TabIndex        =   17
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Adodc"
      Height          =   1095
      Left            =   8160
      TabIndex        =   15
      Top             =   840
      Width           =   1575
      Begin MSAdodcLib.Adodc Adodc5 
         Height          =   330
         Left            =   240
         Top             =   600
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
         CommandType     =   8
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
         Caption         =   "Adodc5"
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
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   330
         Left            =   240
         Top             =   240
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
         CommandType     =   8
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
         RecordSource    =   ""
         Caption         =   "Adodc4"
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "Adodc"
      Height          =   975
      Left            =   8160
      TabIndex        =   14
      Top             =   1920
      Width           =   1575
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   330
         Left            =   240
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
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
         RecordSource    =   "select *from pembeli"
         Caption         =   "Adodc3"
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
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   240
         Top             =   240
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
         CommandType     =   8
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
         RecordSource    =   ""
         Caption         =   "Adodc2"
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
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form4.frx":0000
      Height          =   1815
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   3201
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
      Left            =   2760
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      RecordSource    =   "select*from chas"
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "Form4.frx":0015
      DataField       =   "kode_motor"
      DataSource      =   "Adodc5"
      Height          =   315
      Left            =   1440
      TabIndex        =   12
      Top             =   1560
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "kode_motor"
      Text            =   "Pilih..."
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Form4.frx":002A
      DataField       =   "no_ktp"
      DataSource      =   "Adodc3"
      Height          =   315
      Left            =   1440
      TabIndex        =   11
      Top             =   840
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "no_ktp"
      Text            =   "Pilih..."
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Top             =   2640
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   8520
      TabIndex        =   32
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   8520
      TabIndex        =   31
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Kembali"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Bayar"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Motor"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Lengkap"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "No KTP"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "BELI CHAS"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "chas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If DataCombo1 = "" Or Text1 = "" Or DataCombo2 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
MsgBox "Masih ada data yang belum di isi", , "error"
Else
Adodc1.Recordset.AddNew
Adodc1.Recordset("no_ktp") = DataCombo1
Adodc1.Recordset("nama_lengkap") = Text1
Adodc1.Recordset("kode_motor") = DataCombo2
Adodc1.Recordset("harga") = Text2
Adodc1.Recordset("bayar") = Text3
Adodc1.Recordset("kembali") = Text4
Adodc1.Recordset.Update
MsgBox "Data sudah tersimpan", vbOKOnly, "Sucses!"
Command1.Enabled = False
End If
End Sub

Private Sub Command10_Click()
login.Show
Me.Hide
End Sub

Private Sub Command2_Click()
With Adodc1.Recordset
.Delete
.MoveFirst
End With
End Sub

Private Sub Command3_Click()
DataCombo1 = ""
Text1 = ""
DataCombo2 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text1.SetFocus
End Sub

Private Sub Command4_Click()
psn = MsgBox("are you sure exit ?", vbYesNo, "info")
If psn = vbYes Then
Unload Me
End If
End Sub

Private Sub Command5_Click()
Form5.Show
Me.Hide
End Sub

Private Sub Command6_Click()
ReportChas.Show
End Sub

Private Sub Command7_Click()
Adodc1.RecordSource = "select*from chas where no_ktp like '" & Text5.Text & "'"
Adodc1.Refresh
End Sub

Private Sub Command8_Click()
Adodc1.RecordSource = "select*from chas"
Adodc1.Refresh
End Sub

Private Sub Command9_Click()
On Error Resume Next
If DataCombo1 = "" Or Text1 = "" Or DataCombo2 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
MsgBox "Masih ada data yang belum di isi", , "error"
Else
Adodc1.Recordset("no_ktp") = DataCombo1
Adodc1.Recordset("nama_lengkap") = Text1
Adodc1.Recordset("kode_motor") = DataCombo2
Adodc1.Recordset("harga") = Text2
Adodc1.Recordset("bayar") = Text3
Adodc1.Recordset("kembali") = Text4
Adodc1.Recordset.Update
MsgBox "Data sudah terganti", vbOKOnly, "Sucses!"
End If
End Sub

Private Sub DataCombo1_Click(Area As Integer)
muncul
End Sub

Public Sub muncul()
Adodc2.RecordSource = "select * from pembeli where no_ktp ='" & DataCombo1.Text & "'"
Adodc2.Refresh
Text1.Text = Adodc2.Recordset!nama_lengkap
End Sub

Private Sub DataCombo2_Click(Area As Integer)
tampil
End Sub

Public Sub tampil()
Adodc4.RecordSource = "select * from motor where kode_motor ='" & DataCombo2.Text & "'"
Adodc4.Refresh
Text2.Text = Adodc4.Recordset!harga
End Sub

Private Sub DataGrid1_Click()
DataCombo1 = Adodc1.Recordset!no_ktp
Text1 = Adodc1.Recordset!nama_lengkap
DataCombo2 = Adodc1.Recordset!kode_motor
Text2 = Adodc1.Recordset!harga
Text3 = Adodc1.Recordset!bayar
Text4 = Adodc1.Recordset!kembali
End Sub

Private Sub Text3_keypress(keyascii As Integer)
If keyascii = 13 Then
Text4.SetFocus
Text4.Text = Val(Text3.Text) - Val(Text2.Text)
End If
End Sub

Private Sub Text4_keypress(keyascii As Integer)
If keyascii = 13 Then
Command1.SetFocus
End If
End Sub



Private Sub Timer1_Timer()
Label9.Caption = Format(Time, "hh.mm.ss.am/pm")
Label10.Caption = Format(Date, "dd-mm-yyyy")
End Sub
