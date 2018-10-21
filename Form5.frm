VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form credit 
   BackColor       =   &H0080FF80&
   Caption         =   "Form5"
   ClientHeight    =   7065
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8700
   LinkTopic       =   "Form5"
   ScaleHeight     =   7065
   ScaleWidth      =   8700
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   221
      Left            =   6840
      Top             =   120
   End
   Begin VB.CommandButton Command10 
      Caption         =   "REFRESH"
      Height          =   255
      Left            =   6360
      TabIndex        =   43
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      Height          =   975
      Left            =   4800
      TabIndex        =   40
      Top             =   3960
      Width           =   2895
      Begin VB.CommandButton Command9 
         Caption         =   "SEARCH"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "LOGOUT"
      Height          =   375
      Left            =   7200
      TabIndex        =   39
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   36
      Text            =   "Anda Harus Membayar Sesuai Cicilan "
      Top             =   4680
      Width           =   3255
   End
   Begin VB.Frame Frame3 
      Caption         =   "Button"
      Height          =   1095
      Left            =   4680
      TabIndex        =   27
      Top             =   720
      Width           =   3855
      Begin VB.CommandButton Command7 
         Caption         =   "EDIT"
         Height          =   375
         Left            =   2520
         TabIndex        =   38
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "SAVE"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "DELETE"
         Height          =   255
         Left            =   1320
         TabIndex        =   32
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "AXIT"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "ADD"
         Height          =   255
         Left            =   1320
         TabIndex        =   30
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "NEXT"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "PRINT"
         Height          =   255
         Left            =   1320
         TabIndex        =   28
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   26
      Top             =   4320
      Width           =   3255
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   23
      Top             =   3960
      Width           =   3255
   End
   Begin VB.Frame Frame2 
      Caption         =   "Adodc"
      Height          =   1095
      Left            =   6240
      TabIndex        =   21
      Top             =   2760
      Width           =   1455
      Begin MSAdodcLib.Adodc Adodc5 
         Height          =   330
         Left            =   120
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
         Left            =   120
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
      Height          =   1095
      Left            =   4800
      TabIndex        =   20
      Top             =   2760
      Width           =   1455
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   330
         Left            =   120
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
         RecordSource    =   "select*from pembeli"
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
         Left            =   120
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
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   19
      Top             =   3600
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1440
      TabIndex        =   17
      Text            =   "Pilih..."
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   16
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   15
      Top             =   2520
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   14
      Top             =   2160
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   13
      Top             =   1800
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   12
      Top             =   1080
      Width           =   3255
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form5.frx":0000
      Height          =   1695
      Left            =   120
      TabIndex        =   11
      Top             =   5280
      Width           =   8415
      _ExtentX        =   14843
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
      Left            =   3360
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
      RecordSource    =   "select*from credit"
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
      Bindings        =   "Form5.frx":0015
      DataField       =   "kode_motor"
      DataSource      =   "Adodc5"
      Height          =   315
      Left            =   1440
      TabIndex        =   9
      Top             =   1440
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "kode_motor"
      Text            =   "Pilih..."
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Form5.frx":002A
      DataField       =   "no_ktp"
      DataSource      =   "Adodc3"
      Height          =   315
      Left            =   1440
      TabIndex        =   8
      Top             =   720
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "no_ktp"
      Text            =   "Pilih..."
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   7560
      TabIndex        =   45
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   7560
      TabIndex        =   44
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Bulan"
      Height          =   255
      Left            =   2400
      TabIndex        =   37
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Uang Muka Minimal 1000000"
      Height          =   255
      Left            =   4800
      TabIndex        =   34
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Cicilan Per Bulan"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Harga"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   255
      Left            =   2400
      TabIndex        =   22
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Bunga"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Bunga"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Lama Cicilan"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Sisa"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Uang Muka"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Motor"
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
      Caption         =   "BELI CREDIT"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   26.25
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
      Width           =   3135
   End
End
Attribute VB_Name = "credit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_click()
If Combo1 = "12" Then
Text5 = "10"
Text6.Text = Val(Text4.Text) * Val(Text5.Text / 100)
ElseIf Combo1 = "24" Then
Text5 = "20"
Text6.Text = Val(Text4.Text) * Val(Text5.Text / 100)
ElseIf Combo1 = "36" Then
Text5 = "30"
Text6.Text = Val(Text4.Text) * Val(Text5.Text / 100)
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
If DataCombo1 = "" Or Text1 = "" Or DataCombo2 = "" Or Text2 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Combo1 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Or Text8 = "" Or Text9 = "" Then
MsgBox "Masih ada data yang belum terisi!!!, ,error"
Else
Adodc1.Recordset.AddNew
Adodc1.Recordset("no_ktp") = DataCombo1
Adodc1.Recordset("nama_lengkap") = Text1
Adodc1.Recordset("kode_motor") = DataCombo2
Adodc1.Recordset("harga") = Text2
Adodc1.Recordset("uang_muka") = Text3
Adodc1.Recordset("sisa") = Text4
Adodc1.Recordset("lama_cicilan") = Combo1
Adodc1.Recordset("bunga") = Text5
Adodc1.Recordset("total_bunga") = Text6
Adodc1.Recordset("total_harga") = Text7
Adodc1.Recordset("ppb") = Text8
Adodc1.Recordset("keterangan") = Text9
Adodc1.Recordset.Update
MsgBox "Data telah tersimpan!", vbOKOnly, "Sucses!"
Command1.Enabled = False
End If
End Sub

Private Sub Command10_Click()
Adodc1.RecordSource = "select*from credit"
Adodc1.Refresh
End Sub

Private Sub Command2_Click()
With Adodc1.Recordset
.Delete
.MoveFirst
End With
End Sub

Private Sub Command3_Click()
psn = MsgBox("Are You Sure?", vbYesNo, "info")
If psn = vbYes Then
Unload Me
End If
End Sub

Private Sub Command4_Click()
DataCombo1 = ""
Text1 = ""
DataCombo2 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
End Sub

Private Sub Command5_Click()
Form6.Show
Me.Hide
End Sub

Private Sub Command6_Click()
ReportCredit.Show
End Sub

Private Sub Command7_Click()
On Error Resume Next
If DataCombo1 = "" Or Text1 = "" Or DataCombo2 = "" Or Text2 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Combo1 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Or Text8 = "" Or Text9 = "" Then
MsgBox "Masih ada data yang belum terisi!!!, ,error"
Else
Adodc1.Recordset("no_ktp") = DataCombo1
Adodc1.Recordset("nama_lengkap") = Text1
Adodc1.Recordset("kode_motor") = DataCombo2
Adodc1.Recordset("harga") = Text2
Adodc1.Recordset("uang_muka") = Text3
Adodc1.Recordset("sisa") = Text4
Adodc1.Recordset("lama_cicilan") = Combo1
Adodc1.Recordset("bunga") = Text5
Adodc1.Recordset("total_bunga") = Text6
Adodc1.Recordset("total_harga") = Text7
Adodc1.Recordset("ppb") = Text8
Adodc1.Recordset("keterangan") = Text9
Adodc1.Recordset.Update
MsgBox "Data telah terganti!", vbOKOnly, "Sucses!"
End If
End Sub

Private Sub Command8_Click()
login.Show
Me.Hide
End Sub

Private Sub Command9_Click()
Adodc1.RecordSource = "select*from credit where no_ktp like '" & Text10.Text & "'"
Adodc1.Refresh
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
Adodc4.RecordSource = "select *from motor where kode_motor ='" & DataCombo2.Text & "'"
Adodc4.Refresh
Text2.Text = Adodc4.Recordset!harga
End Sub

Private Sub DataGrid1_Click()
DataCombo1 = Adodc1.Recordset!no_ktp
Text1 = Adodc1.Recordset!nama_lengkap
DataCombo2 = Adodc1.Recordset!kode_motor
Text2 = Adodc1.Recordset!harga
Text3 = Adodc1.Recordset!uang_muka
Text4 = Adodc1.Recordset!sisa
Combo1 = Adodc1.Recordset!lama_cicilan
Text5 = Adodc1.Recordset!bunga
Text6 = Adodc1.Recordset!total_bunga
Text7 = Adodc1.Recordset!total_harga
Text8 = Adodc1.Recordset!ppb
Text9 = Adodc1.Recordset!keterangan
End Sub

Private Sub Form_Load()
With Combo1
.AddItem "12"
.AddItem "24"
.AddItem "36"
End With
End Sub

Private Sub Text3_keypress(keyascii As Integer)
If keyascii = 13 Then
Text4.SetFocus
Text4.Text = Val(Text2.Text) - Val(Text3.Text)
End If
End Sub

Private Sub Text5_keypress(keyascii As Integer)
If keyascii = 13 Then
Text6.SetFocus
Text6.Text = (Text5.Text / 12)
End If
End Sub

Private Sub Text6_(keyascii As Integer)
If keyascii = 13 Then
Command1.SetFocus
End If
End Sub

Private Sub Text6_keypress(keyascii As Integer)
If keyascii = 13 Then
Text7.SetFocus
Text7.Text = Val(Text4.Text) + Val(Text6.Text)
End If
End Sub

Private Sub Text7_keypress(keyascii As Integer)
If keyascii = 13 Then
Text8.SetFocus
Text8.Text = Val(Text7.Text) / Val(Combo1.Text)
End If
End Sub


Private Sub Timer1_Timer()
Label17.Caption = Format(Time, "hh.mm.ss.am/pm")
Label18.Caption = Format(Date, "dd-mm-yyyy")
End Sub
