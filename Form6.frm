VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form cicilan1 
   BackColor       =   &H0080FF80&
   Caption         =   "Form6"
   ClientHeight    =   8595
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15045
   LinkTopic       =   "Form6"
   ScaleHeight     =   8595
   ScaleWidth      =   15045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "ADD"
      Height          =   255
      Left            =   11280
      TabIndex        =   89
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "DELETE"
      Height          =   255
      Left            =   10200
      TabIndex        =   88
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "EXIT"
      Height          =   255
      Left            =   11280
      TabIndex        =   87
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NEXT"
      Height          =   255
      Left            =   10200
      TabIndex        =   86
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PRINT"
      Height          =   255
      Left            =   11280
      TabIndex        =   85
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SAVE"
      Height          =   255
      Left            =   10200
      TabIndex        =   84
      Top             =   5160
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form6.frx":0000
      Height          =   1575
      Left            =   120
      TabIndex        =   83
      Top             =   6600
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   2778
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
      Left            =   5040
      Top             =   6120
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
      CommandType     =   2
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
      RecordSource    =   "cicilan"
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
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   975
      Left            =   11640
      TabIndex        =   82
      Top             =   4080
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
      Caption         =   "Frame1"
      Height          =   975
      Left            =   10200
      TabIndex        =   81
      Top             =   4080
      Width           =   1455
      Begin MSAdodcLib.Adodc Adodc3 
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
   Begin VB.TextBox Text26 
      Height          =   285
      Left            =   11520
      TabIndex        =   80
      Top             =   3600
      Width           =   3375
   End
   Begin VB.TextBox Text25 
      Height          =   285
      Left            =   11520
      TabIndex        =   79
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox Text24 
      Height          =   285
      Left            =   11520
      TabIndex        =   78
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox Text23 
      Height          =   285
      Left            =   11520
      TabIndex        =   77
      Top             =   2160
      Width           =   3375
   End
   Begin VB.TextBox Text22 
      Height          =   285
      Left            =   11520
      TabIndex        =   76
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox Text21 
      Height          =   285
      Left            =   11520
      TabIndex        =   75
      Top             =   1080
      Width           =   3375
   End
   Begin VB.ComboBox Combo12 
      Height          =   315
      Left            =   11520
      TabIndex        =   74
      Text            =   "Combo1"
      Top             =   2880
      Width           =   3375
   End
   Begin VB.ComboBox Combo11 
      Height          =   315
      Left            =   11520
      TabIndex        =   73
      Text            =   "Combo1"
      Top             =   1800
      Width           =   3375
   End
   Begin VB.ComboBox Combo10 
      Height          =   315
      Left            =   11520
      TabIndex        =   72
      Text            =   "Combo1"
      Top             =   720
      Width           =   3375
   End
   Begin VB.ComboBox Combo9 
      Height          =   315
      Left            =   6480
      TabIndex        =   63
      Text            =   "Combo1"
      Top             =   5040
      Width           =   3375
   End
   Begin VB.ComboBox Combo8 
      Height          =   315
      Left            =   6480
      TabIndex        =   62
      Text            =   "Combo1"
      Top             =   3960
      Width           =   3375
   End
   Begin VB.ComboBox Combo7 
      Height          =   315
      Left            =   6480
      TabIndex        =   61
      Text            =   "Combo1"
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox Text20 
      Height          =   285
      Left            =   6480
      TabIndex        =   60
      Top             =   5760
      Width           =   3375
   End
   Begin VB.TextBox Text19 
      Height          =   285
      Left            =   6480
      TabIndex        =   59
      Top             =   5400
      Width           =   3375
   End
   Begin VB.TextBox Text18 
      Height          =   285
      Left            =   6480
      TabIndex        =   58
      Top             =   4680
      Width           =   3375
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   6480
      TabIndex        =   57
      Top             =   4320
      Width           =   3375
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   6480
      TabIndex        =   56
      Top             =   3600
      Width           =   3375
   End
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   6480
      TabIndex        =   55
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   6480
      TabIndex        =   54
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   6480
      TabIndex        =   53
      Top             =   2160
      Width           =   3375
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   6480
      TabIndex        =   52
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   6480
      TabIndex        =   51
      Top             =   1080
      Width           =   3375
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      Left            =   6480
      TabIndex        =   50
      Text            =   "Combo1"
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   1440
      TabIndex        =   45
      Top             =   6120
      Width           =   3375
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   1440
      TabIndex        =   44
      Top             =   5760
      Width           =   3375
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   1440
      TabIndex        =   43
      Top             =   5040
      Width           =   3375
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1440
      TabIndex        =   28
      Top             =   4680
      Width           =   3375
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1440
      TabIndex        =   27
      Top             =   3960
      Width           =   3375
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1440
      TabIndex        =   26
      Top             =   3600
      Width           =   3375
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   6480
      TabIndex        =   25
      Text            =   "Combo1"
      Top             =   720
      Width           =   3375
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   1440
      TabIndex        =   24
      Text            =   "Combo1"
      Top             =   5400
      Width           =   3375
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1440
      TabIndex        =   23
      Text            =   "Combo1"
      Top             =   4320
      Width           =   3375
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1440
      TabIndex        =   22
      Text            =   "Combo1"
      Top             =   3240
      Width           =   3375
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1440
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   2160
      Width           =   3375
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1440
      TabIndex        =   12
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1440
      TabIndex        =   11
      Top             =   2520
      Width           =   3375
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "Form6.frx":0015
      DataField       =   "kode_motor"
      DataSource      =   "Adodc5"
      Height          =   315
      Left            =   1440
      TabIndex        =   10
      Top             =   1440
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "kode_motor"
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Form6.frx":002A
      DataField       =   "no_ktp"
      DataSource      =   "Adodc3"
      Height          =   315
      Left            =   1440
      TabIndex        =   9
      Top             =   720
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "no_ktp"
      Text            =   "DataCombo1"
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label Label41 
      Caption         =   "Sisa Cicilan"
      Height          =   255
      Left            =   10200
      TabIndex        =   71
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label40 
      Caption         =   "Bayar Cicilan"
      Height          =   255
      Left            =   10200
      TabIndex        =   70
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label39 
      Caption         =   "Cicilan Ke 12"
      Height          =   255
      Left            =   10200
      TabIndex        =   69
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label38 
      Caption         =   "Sisa Cicilan"
      Height          =   255
      Left            =   10200
      TabIndex        =   68
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label37 
      Caption         =   "Bayar Cicilan"
      Height          =   255
      Left            =   10200
      TabIndex        =   67
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label36 
      Caption         =   "Cicilan Ke 11"
      Height          =   255
      Left            =   10200
      TabIndex        =   66
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label35 
      Caption         =   "Sisa Cicilan"
      Height          =   255
      Left            =   10200
      TabIndex        =   65
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label34 
      Caption         =   "Bayar Cicilan"
      Height          =   255
      Left            =   10200
      TabIndex        =   64
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label33 
      Caption         =   "Cicilan Ke 10"
      Height          =   255
      Left            =   10200
      TabIndex        =   49
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label32 
      Caption         =   "Sisa Cicilan"
      Height          =   255
      Left            =   5160
      TabIndex        =   48
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label31 
      Caption         =   "Bayar Cicilan"
      Height          =   255
      Left            =   5160
      TabIndex        =   47
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label30 
      Caption         =   "Cicilan Ke 9"
      Height          =   255
      Left            =   5160
      TabIndex        =   46
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label29 
      Caption         =   "Sisa Cicilan"
      Height          =   255
      Left            =   5160
      TabIndex        =   42
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label28 
      Caption         =   "Bayar Cicilan"
      Height          =   255
      Left            =   5160
      TabIndex        =   41
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label27 
      Caption         =   "Cicilan Ke 8"
      Height          =   255
      Left            =   5160
      TabIndex        =   40
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label26 
      Caption         =   "Sisa Cicilan"
      Height          =   255
      Left            =   5160
      TabIndex        =   39
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label25 
      Caption         =   "Bayar Cicilan"
      Height          =   255
      Left            =   5160
      TabIndex        =   38
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label24 
      Caption         =   "Cicilan Ke 7"
      Height          =   255
      Left            =   5160
      TabIndex        =   37
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label23 
      Caption         =   "Sisa Cicilan"
      Height          =   255
      Left            =   5160
      TabIndex        =   36
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label22 
      Caption         =   "Bayar Cicilan"
      Height          =   255
      Left            =   5160
      TabIndex        =   35
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label21 
      Caption         =   "Cicilan Ke 6"
      Height          =   255
      Left            =   5160
      TabIndex        =   34
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label20 
      Caption         =   "Sisa Cicilan"
      Height          =   255
      Left            =   5160
      TabIndex        =   33
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label19 
      Caption         =   "Bayar Cicilan"
      Height          =   255
      Left            =   5160
      TabIndex        =   32
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label18 
      Caption         =   "Cicilan Ke 5"
      Height          =   255
      Left            =   5160
      TabIndex        =   31
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label17 
      Caption         =   "Sisa Cicilan"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "Bayar Cicilan"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "Cicilan ke 4"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "Sisa Cicilan"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Bayar Cicilan"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "Cicilan Ke 3"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Sisa Cicilan"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Bayar Cicilan"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Cicilan ke2"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Sisa Cicilan"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Bayar Cicilan"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Cicilan Ke 1"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Harga"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Kode Motor"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Nama Lengkap"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "No KTP"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Bayar Cicilam Motor"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
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
      Width           =   3495
   End
End
Attribute VB_Name = "cicilan1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_click()
If Combo1 = "1" Then
Text3 = ""
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

