VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_empresa 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empresa"
   ClientHeight    =   6930
   ClientLeft      =   2760
   ClientTop       =   1680
   ClientWidth     =   8910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmd_confirmar2 
      BackColor       =   &H00808080&
      Caption         =   "Confirmar Modif"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmd_cancelar 
      BackColor       =   &H00808080&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmd_confirmaralta 
      BackColor       =   &H00808080&
      Caption         =   "Confirmar Alta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmd_modif 
      BackColor       =   &H00808080&
      Caption         =   "Modificar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmd_baja 
      BackColor       =   &H00808080&
      Caption         =   "Baja"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmd_alta 
      BackColor       =   &H00808080&
      Caption         =   "Alta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3480
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   7200
      Top             =   6480
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
      Connect         =   "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=proyecto1;Data Source=."
      OLEDBString     =   "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=proyecto1;Data Source=."
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from empresa"
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frm_empresa.frx":0000
      Height          =   2175
      Left            =   0
      TabIndex        =   29
      Top             =   4800
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "tipo"
         Caption         =   "Tipo de transporte"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "capacidad"
         Caption         =   "Capacidad en Kilos"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "marca"
         Caption         =   "Marca del transporte"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "dis_max"
         Caption         =   "Distancia máxima (Km)"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1995,024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2355,024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1995,024
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmd_ocultemp 
      BackColor       =   &H00808080&
      Caption         =   "O&cultar transportes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CommandButton cmd_mostmedios 
      BackColor       =   &H00808080&
      Caption         =   "Mo&strar transportes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CommandButton cmd_volver 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      Picture         =   "frm_empresa.frx":0015
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4080
      Width           =   735
   End
   Begin VB.Frame frm_vol_neg 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Volumen de negocios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5280
      TabIndex        =   18
      Top             =   2160
      Width           =   3255
      Begin VB.OptionButton opt_volumen3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Más de 60.000 Euros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   1320
         Width           =   2295
      End
      Begin VB.OptionButton opt_volumen2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Entre 30.000 y 60.000 Euros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   840
         Width           =   2775
      End
      Begin VB.OptionButton opt_volumen1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Menos de 30.000 Euros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame frm_ambito 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ambito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   5280
      TabIndex        =   17
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton opt_ambito3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Internacional"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   1440
         Width           =   1695
      End
      Begin VB.OptionButton opt_ambito2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nacional"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton opt_ambito1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Local"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmd_ocultar 
      BackColor       =   &H00808080&
      Caption         =   "&Ocultar empleados "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmd_lista_emp 
      BackColor       =   &H00808080&
      Caption         =   "&Mostrar empleados "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4200
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   720
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Connect         =   "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=proyecto1;Data Source=."
      OLEDBString     =   "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=proyecto1;Data Source=."
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from empresa"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm_empresa.frx":03F0
      Height          =   2055
      Left            =   0
      TabIndex        =   9
      Top             =   4800
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777215
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "dni"
         Caption         =   "Dni"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "nombre"
         Caption         =   "Nombre"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "apellido"
         Caption         =   "Apellido"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "car_rod"
         Caption         =   "C.Rodado"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "Sí"
            FalseValue      =   "No"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "car_fer"
         Caption         =   "C.Ferroviario"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "Sí"
            FalseValue      =   "No"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "car_aer"
         Caption         =   "C.Aéreo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "Sí"
            FalseValue      =   "No"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "car_mar"
         Caption         =   "C.Marítimo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "Sí"
            FalseValue      =   "No"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   7
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1305,071
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1844,787
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1244,976
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   794,835
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1049,953
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   6
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   2160
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   2160
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   2160
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton cmd_mover 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   2880
      Picture         =   "frm_empresa.frx":0405
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton cmd_mover 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   2280
      Picture         =   "frm_empresa.frx":1047
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton cmd_mover 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1680
      Picture         =   "frm_empresa.frx":1C89
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton cmd_mover 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1080
      Picture         =   "frm_empresa.frx":28CB
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   615
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3720
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Connect         =   "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=proyecto1;Data Source=."
      OLEDBString     =   "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=proyecto1;Data Source=."
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.Line Line8 
      X1              =   2280
      X2              =   2160
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line7 
      X1              =   960
      X2              =   1080
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line6 
      X1              =   2400
      X2              =   2280
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label lblempresa 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   26
      Top             =   0
      Width           =   1095
   End
   Begin VB.Line Line5 
      X1              =   8760
      X2              =   8760
      Y1              =   120
      Y2              =   3960
   End
   Begin VB.Line Line4 
      X1              =   8760
      X2              =   120
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line3 
      X1              =   960
      X2              =   120
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line2 
      X1              =   8760
      X2              =   2400
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   120
      Y1              =   3960
      Y2              =   120
   End
   Begin VB.Label lbl_pres_mant 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fondos mantenimiento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label lblnum_emp 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nº Empleados:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lbl_dir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Director:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lbl_dom 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Domicilio:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblnombre 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frm_empresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim rs As adodb.Recordset
Public variablecodigo As Integer

Private Sub cmd_alta_Click()
    cmd_confirmaralta.Visible = True
    cmd_baja.Visible = False
    cmd_modif.Visible = False
    cmd_cancelar.Visible = True
    
    frm_empresa.Height = 5325
    cmd_mostmedios.Visible = False
    cmd_ocultemp.Visible = False
    cmd_ocultar.Visible = False
    cmd_lista_emp.Visible = False
    
    Text1(6).Visible = False
    lblnum_emp.Visible = False
    
    Text1(0).SetFocus
    Text1(0).Locked = False
    Text1(1).Locked = False
    Text1(3).Locked = False
    Text1(4).Locked = False
    
    For i = 0 To 3
        cmd_mover(i).Enabled = False
    Next
    
    Text1(0).Text = ""
    Text1(1).Text = ""
    Text1(3).Text = ""
    
    Text1(4).Text = ""
    
    opt_volumen3.Enabled = True
    opt_volumen2.Enabled = True
    opt_volumen1.Enabled = True
    
    opt_ambito1.Enabled = True
    opt_ambito2.Enabled = True
    opt_ambito3.Enabled = True
    
End Sub

Private Sub cmd_baja_Click()
Dim respuesta As String
Dim frasesql, frasesql2, frasesql3 As String

    If (rs.EOF Or rs.BOF) Then
        MsgBox "No hay registros activos para eliminar", vbOKOnly, "Gestión"
        Exit Sub
    Else
        respuesta = MsgBox("¿Estas seguro de borrar el registro?", vbDefaultButton2 + vbYesNo, "Gestión")
        
        If respuesta = 6 Then
        
            frasesql = "DELETE FROM empresa WHERE nom_emp = '" & Trim(Text1(0).Text) & "' AND codigo NOT IN (SELECT cod_emp FROM m_trans) AND codigo NOT IN (SELECT cod_emp from empleado) AND codigo NOT IN (SELECT cod_emp from c_mante)"
            cn.Execute (frasesql)
            
            If rs.EOF And rs.BOF Then
                Exit Sub
            Else
            
                rs.MoveFirst
            
                rs.Find ("nom_emp  = '" & Trim(Text1(0).Text) & "'")
            
                If rs.EOF Then

                    MsgBox "Baja realizada con éxito", vbOKOnly, "Gestión"
                        
                Else
                    MsgBox "No se puede realizar la baja ya que esa empresa esta en relacion con otros datos", vbOKOnly, "Gestion"
                
                End If
                
            End If
            
            cmd_mover_Click 0
        End If
    End If
End Sub

Private Sub cmd_cancelar_Click()
    Call restaurar
End Sub

Private Sub cmd_confirmar2_Click()
Dim respuesta As String
Dim frasesql As String
Dim ambito As String
Dim vol_negs As String

If Text1(0).Text = "" Or IsNumeric(Text1(0).Text) Then
    MsgBox "El nombre de la empresa debe estar relleno y no debe de ser numérico", vbOKOnly, "Gestión"
Else
    If (rs.EOF Or rs.BOF) Then
        MsgBox "No hay registros activos para modificar", vbOKOnly, "Gestión"
    Else
        respuesta = MsgBox("¿Estas seguro de modificar el registro?", vbDefaultButton2 + vbYesNo, "Gestión")
        
        If respuesta = 6 Then
        
            If opt_ambito1.Value = True Then
                ambito = "Local"
            ElseIf opt_ambito2.Value = True Then
                ambito = "Nacional"
            ElseIf opt_ambito3.Value = True Then
                ambito = "Internacional"
            End If
            
            If opt_volumen1.Value = True Then
                vol_negs = "<30000"
            ElseIf opt_volumen2.Value = True Then
                vol_negs = "30000-60000"
            ElseIf opt_volumen3.Value = True Then
                vol_negs = ">60000"
            End If
            
            rs.MoveFirst
            
            rs.Find ("nom_emp = '" & Text1(0).Text & "'")
            
            If rs.EOF Then
            
                frasesql = "UPDATE empresa SET nom_emp = '" & Trim(Text1(0).Text) & "',domicilio = '" & Trim(Text1(1).Text) & "'," & _
                "tip_empres = '" & ambito & "', nom_dir = '" & Trim(Text1(3).Text) & "',pres_mant = " & Val(Text1(4).Text) & ",vol_neg = '" & vol_negs & "' WHERE codigo = '" & Trim(Text1(2).Text) & "'"
            
                cn.Execute (frasesql)
            
                MsgBox "Modificación realizada con éxito", vbOKOnly, "Gestión"
                frm_distribu.Adodc6.Refresh
            
            Else
            
                MsgBox "No se puede realizar la modificacion ya que el nombre de empresa asignado corresponde ya al de una empresa existente", vbOKOnly, "Gestion"
            
            End If
            
            cmd_confirmar2.Visible = False
            cmd_cancelar.Visible = False
            cmd_alta.Visible = True
            cmd_baja.Visible = True
            cmd_modif.Visible = True
    
            Text1(0).Locked = True
            Text1(1).Locked = True
            Text1(3).Locked = True
            Text1(4).Locked = True
            
            For i = 0 To 3
                cmd_mover(i).Enabled = True
            Next
            
            Call restaurar
            
            cmd_mover_Click 0
        End If
    End If
End If
End Sub

Private Sub cmd_confirmaralta_Click()
    Dim frasesql As String
    Dim ambito, vol_neg As String
    Dim numemple As Integer
    
    numemple = 0
    
    If Text1(0).Text = "" Or Text1(1).Text = "" Or Text1(3).Text = "" Or Text1(4).Text = "" Then
        MsgBox "Por favor rellene todas las casillas con los datos correspondientes", vbOKOnly, "Gestión"
    Else
        If Not IsNumeric(Text1(4).Text) Then
            MsgBox "El numero de empleados y los fondos para el mantenimiento deben ser numéricos", vbOKOnly, "Gestión"
            Text1(4).Text = ""
            
        Else
            If opt_ambito1.Value = True Then
                ambito = "Local"
            ElseIf opt_ambito2.Value = True Then
                ambito = "Nacional"
            Else
                ambito = "Internacional"
            End If
            
            If opt_volumen1.Value = True Then
                vol_neg = "<30000"
            ElseIf opt_volumen2.Value = True Then
                vol_neg = "30000-60000"
            Else
                vol_neg = ">60000"
            End If
            
            rs.MoveFirst
            
            rs.Find ("nom_emp = '" & Text1(0).Text & "'")
            
            If rs.EOF Then
            
                frasesql = "Insert into empresa values('" & Text1(0).Text & "','" & Text1(1).Text & "','" & ambito & "','" & Text1(3).Text & "','" & Text1(4).Text & "','" & vol_neg & "','" & numemple & "')"
                cn.Execute frasesql
                MsgBox "Alta realizada con éxito", vbOKOnly, "Gestión"
            Else
            
                MsgBox "No se puede realizar alta de empresa ya que el nombre asignado pertenece al de otra empresa", vbOKOnly, "Gestión"
                Text1(0).Text = ""
                Exit Sub
            End If
            
            Call restaurar
        End If
    End If
End Sub

Private Sub cmd_lista_emp_Click()
    frm_empresa.Height = 7590
    cmd_lista_emp.Visible = False
    cmd_ocultar.Visible = True
    DataGrid1.Visible = True
    cmd_mostmedios.Visible = False
End Sub

Private Sub cmd_lista_emp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_lista_emp.BackColor = &HEEEEEE
End Sub

Private Sub cmd_lista_emp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_lista_emp.BackColor = &H808080
End Sub

Private Sub cmd_modif_Click()

    frm_empresa.Height = 5325
    cmd_mostmedios.Visible = False
    cmd_ocultemp.Visible = False
    cmd_ocultar.Visible = False
    cmd_lista_emp.Visible = False
    
    cmd_confirmar2.Visible = True
    cmd_baja.Visible = False
    cmd_modif.Visible = False
    cmd_cancelar.Visible = True
    cmd_confirmaralta.Visible = False
    
    Text1(0).SetFocus
    Text1(0).Locked = False
    Text1(1).Locked = False
    Text1(3).Locked = False
    Text1(4).Locked = False
    Text1(6).Locked = False
    
    For i = 0 To 3
        cmd_mover(i).Enabled = False
    Next
    
    opt_volumen3.Enabled = True
    opt_volumen2.Enabled = True
    opt_volumen1.Enabled = True
    
    opt_ambito1.Enabled = True
    opt_ambito2.Enabled = True
    opt_ambito3.Enabled = True
    
End Sub

Private Sub cmd_mostmedios_Click()
    frm_empresa.Height = 7590
    cmd_mostmedios.Visible = False
    cmd_ocultemp.Visible = True
    DataGrid2.Visible = True
    cmd_lista_emp.Visible = False
End Sub

Private Sub cmd_mostmedios_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_mostmedios.BackColor = &HFFFFFF
End Sub

Private Sub cmd_mostmedios_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_mostmedios.BackColor = &H808080
End Sub

Private Sub cmd_mover_Click(Index As Integer)
    On Error Resume Next
    
    If rs.BOF = True And rs.EOF = True Then
        Exit Sub
    End If
    
    Select Case Index
        Case 0
            rs.MoveFirst
        Case 1
            rs.MovePrevious
        Case 2
            rs.MoveNext
        Case 3
            rs.MoveLast
    End Select
    
    If rs.BOF Then
        MsgBox "Ya está en el primer registro", vbOKOnly, "Advertencia"
        rs.MoveFirst
    ElseIf rs.EOF Then
        MsgBox "Ya está en el último registro", vbOKOnly, "Advertencia"
        rs.MoveLast
    End If
 
 Call mostrarempresa
 
End Sub

Private Sub cmd_mover_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Select Case Index
    Case 0:
    cmd_mover(0).BackColor = &HFFFFFF
    cmd_mover(0).Picture = LoadPicture(".\iconos\primeropulsado.bmp")
    Case 1:
    cmd_mover(1).BackColor = &HFFFFFF
    cmd_mover(1).Picture = LoadPicture(".\iconos\anteriorpulsado.bmp")
    Case 2:
    cmd_mover(2).BackColor = &HFFFFFF
    cmd_mover(2).Picture = LoadPicture(".\iconos\siguientepulsado.bmp")
    Case 3:
    cmd_mover(3).BackColor = &HFFFFFF
    cmd_mover(3).Picture = LoadPicture(".\iconos\ultimopulsado.bmp")
    End Select
    
End Sub

Private Sub cmd_mover_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
    Case 0:
    cmd_mover(0).BackColor = &H808080
    cmd_mover(0).Picture = LoadPicture(".\iconos\primerosinpulsar.bmp")
    Case 1:
    cmd_mover(1).BackColor = &H808080
    cmd_mover(1).Picture = LoadPicture(".\iconos\anteriorsinpulsar.bmp")
    Case 2:
    cmd_mover(2).BackColor = &H808080
    cmd_mover(2).Picture = LoadPicture(".\iconos\siguientesinpulsar.bmp")
    Case 3:
    cmd_mover(3).BackColor = &H808080
    cmd_mover(3).Picture = LoadPicture(".\iconos\ultimosinpulsar.bmp")
    End Select
End Sub

Private Sub cmd_ocultar_Click()
    cmd_lista_emp.Visible = True
    cmd_ocultar.Visible = False
    frm_empresa.Height = 5325
    DataGrid1.Visible = False
    cmd_mostmedios.Visible = True
End Sub

Private Sub cmd_ocultar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_ocultar.BackColor = &HEEEEEE
End Sub

Private Sub cmd_ocultar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_ocultar.BackColor = &H808080
End Sub

Private Sub cmd_ocultemp_Click()
    cmd_mostmedios.Visible = True
    cmd_ocultemp.Visible = False
    frm_empresa.Height = 5325
    DataGrid2.Visible = False
    cmd_lista_emp.Visible = True
End Sub

Private Sub cmd_ocultemp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_ocultemp.BackColor = &HFFFFFF
End Sub

Private Sub cmd_ocultemp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_ocultemp.BackColor = &H808080
End Sub

Private Sub cmd_volver_Click()
    frm_empresa.Hide
    cmd_confirmaralta.Visible = False
    cmd_cancelar.Visible = False
    cmd_baja.Visible = True
    cmd_modif.Visible = True
    cmd_confirmar2.Visible = False
    
    cmd_mostmedios.Visible = True
    cmd_ocultemp.Visible = False
    cmd_ocultar.Visible = False
    cmd_lista_emp.Visible = True
    
    Text1(0).Locked = True
    Text1(1).Locked = True
    Text1(3).Locked = True
    Text1(4).Locked = True
    
    For i = 0 To 3
        cmd_mover(i).Enabled = True
    Next
    
    frm_ambito.Enabled = True
    frm_vol_neg.Enabled = True
    
    cmd_mover_Click 0
    
    Call volver
End Sub

Private Sub cmd_volver_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_volver.BackColor = &HEEEEEE
End Sub

Private Sub cmd_volver_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_volver.BackColor = &H808080
End Sub

Private Sub Form_Activate()
    cmd_mover_Click 0
End Sub

Private Sub Form_Load()

    Set cn = New adodb.Connection
    Set rs = New adodb.Recordset

    Call conexion

    rs.Open "Select * from empresa", cn, adOpenDynamic, adLockOptimistic

    DataGrid1.Visible = False
    DataGrid2.Visible = False

    cmd_confirmaralta.Visible = False
    cmd_cancelar.Visible = False
    cmd_confirmar2.Visible = False

    cmd_ocultar_Click
    cmd_ocultemp_Click

    cmd_mover_Click 0

    For i = 0 To 6
        DataGrid1.Columns(i).AllowSizing = False
    Next

    For i = 0 To 3
        DataGrid2.Columns(i).AllowSizing = False
    Next

End Sub

Private Sub mostrarempresa()

Dim empleadostotales As Integer
Dim frasesql As String

    Text1(0).Text = " "
    Text1(1).Text = " "
    Text1(3).Text = " "
    Text1(4).Text = " "
    Text1(6).Text = " "
    
        With rs
        
            Text1(0).Text = .Fields("nom_emp")
            Text1(1).Text = .Fields("domicilio")
            Text1(3).Text = .Fields("nom_dir")
            Text1(4).Text = .Fields("pres_mant")
            Text1(6).Text = .Fields("num_emple")
            Text1(2).Text = .Fields("codigo")
            
            If .Fields("tip_empres") = "Local" Then
            
                opt_ambito1.Value = True
                opt_ambito1.Enabled = True
                opt_ambito2.Enabled = False
                opt_ambito3.Enabled = False
                
            ElseIf .Fields("tip_empres") = "Nacional" Then
            
                opt_ambito2.Value = True
                opt_ambito2.Enabled = True
                opt_ambito1.Enabled = False
                opt_ambito3.Enabled = False
                
            Else
            
                opt_ambito3.Value = True
                opt_ambito3.Enabled = True
                opt_ambito2.Enabled = False
                opt_ambito1.Enabled = False
            
            End If
            
            
            If .Fields("vol_neg") = "<30000" Then
                
                opt_volumen1.Value = True
                opt_volumen1.Enabled = True
                opt_volumen2.Enabled = False
                opt_volumen3.Enabled = False
                
            ElseIf .Fields("vol_neg") = "30000-60000" Then
                
                opt_volumen2.Value = True
                opt_volumen2.Enabled = True
                opt_volumen1.Enabled = False
                opt_volumen3.Enabled = False
                
                
            Else
            
                opt_volumen3.Value = True
                opt_volumen3.Enabled = True
                opt_volumen2.Enabled = False
                opt_volumen1.Enabled = False
                
            End If
        End With
End Sub

Private Sub Text1_Change(Index As Integer)
    Adodc2.RecordSource = "Select * from empleado where empleado.cod_emp =  " & Val(Text1(2).Text)
    
    DataGrid1.Refresh
    Adodc2.Refresh
    
    Adodc3.RecordSource = "Select * from m_trans where m_trans.cod_emp =  " & Val(Text1(2).Text)
    
    DataGrid2.Refresh
    Adodc3.Refresh
End Sub

Private Sub restaurar()
    cmd_confirmaralta.Visible = False
    cmd_confirmar2.Visible = False
    cmd_cancelar.Visible = False
    cmd_alta.Visible = True
    cmd_baja.Visible = True
    cmd_modif.Visible = True
    
    cmd_mostmedios.Visible = True
    cmd_ocultemp.Visible = False
    cmd_ocultar.Visible = False
    cmd_lista_emp.Visible = True
    
    Text1(0).Locked = True
    Text1(1).Locked = True
    Text1(3).Locked = True
    Text1(4).Locked = True
    Text1(6).Locked = True
    
    For i = 0 To 3
        cmd_mover(i).Enabled = True
    Next
    
    frm_ambito.Enabled = True
    frm_vol_neg.Enabled = True
    Text1(6).Visible = True
    lblnum_emp.Visible = True
    
    cmd_mover_Click 0
End Sub
