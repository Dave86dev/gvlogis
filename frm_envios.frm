VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_envios 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Envios"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10020
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   10020
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   8280
      Top             =   960
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
      RecordSource    =   "select  * from empleado"
      Caption         =   "Adodc6"
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
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "frm_envios.frx":0000
      Height          =   315
      Left            =   5400
      TabIndex        =   12
      Top             =   960
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "nom_emp"
      BoundColumn     =   "codigo"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   375
      Left            =   4800
      Top             =   1800
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      RecordSource    =   "select dni,(nombre+'  '+apellido) as ncompleto,cod_emp from empleado"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   1680
      Top             =   1440
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      RecordSource    =   "select cod_mtrans,cod_carga from reparto where asignado = 0 order by cod_carga asc"
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "frm_envios.frx":0015
      Height          =   315
      Left            =   1680
      TabIndex        =   11
      Top             =   960
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "cod_carga"
      BoundColumn     =   "cod_mtrans"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   240
      Top             =   4440
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   $"frm_envios.frx":002A
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frm_envios.frx":0177
      Height          =   1935
      Left            =   240
      TabIndex        =   10
      Top             =   2400
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   3413
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
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "cod_carga"
         Caption         =   "Carga"
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
         DataField       =   "cod_itine"
         Caption         =   "cod_itine"
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
         DataField       =   "cod_mtrans"
         Caption         =   "cod_mtrans"
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
         DataField       =   "fecha"
         Caption         =   "Fecha"
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
      BeginProperty Column04 
         DataField       =   "asignado"
         Caption         =   "asignado"
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
      BeginProperty Column05 
         DataField       =   "proced"
         Caption         =   "Procedencia"
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
      BeginProperty Column06 
         DataField       =   "destino"
         Caption         =   "Destino"
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
      BeginProperty Column07 
         DataField       =   "descripcion"
         Caption         =   "Descripcion"
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
      BeginProperty Column08 
         DataField       =   "marca"
         Caption         =   "M.transporte"
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
      BeginProperty Column09 
         DataField       =   "nom_emp"
         Caption         =   "Empresa"
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
            Object.Visible         =   -1  'True
            ColumnWidth     =   599,811
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column03 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column05 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1395,213
         EndProperty
         BeginProperty Column06 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1395,213
         EndProperty
         BeginProperty Column07 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   2099,906
         EndProperty
         BeginProperty Column08 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1349,858
         EndProperty
         BeginProperty Column09 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1094,74
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo DataCombo5 
      Bindings        =   "frm_envios.frx":018C
      Height          =   315
      Left            =   1680
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "ncompleto"
      BoundColumn     =   "dni"
      Text            =   ""
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton cmd_confirmar 
      BackColor       =   &H00808080&
      Caption         =   "Confirmar"
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
      TabIndex        =   6
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton cmd_asignar 
      BackColor       =   &H00808080&
      Caption         =   "Envios por asignar"
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
      TabIndex        =   5
      Top             =   960
      Width           =   4095
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   6840
      Top             =   5160
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      RecordSource    =   "select * from pilota"
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
      Bindings        =   "frm_envios.frx":01A1
      Height          =   1935
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   3413
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
            LCID            =   3082
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
            LCID            =   3082
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
      Left            =   6840
      Top             =   4800
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "select dni,(nombre+'   '+apellido) As Ncompleto from empleado order by nombre asc"
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frm_envios.frx":01B6
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   1800
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "Ncompleto"
      BoundColumn     =   "dni"
      Text            =   ""
   End
   Begin VB.CommandButton cmd_volver 
      BackColor       =   &H00808080&
      Height          =   615
      Left            =   9240
      Picture         =   "frm_envios.frx":01CB
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label lblenvios 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Envios:"
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
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblempleado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Empleado:"
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
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lbldaenvios 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "RELACION DE ENVIOS POR EMPLEADO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frm_envios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs_reparto As adodb.Recordset

Private Sub cmd_asignar_Click()
    Call restaurar
    cmd_asignar.Visible = False
    lblenvios.Visible = True
    DataCombo2.Visible = True
    cmd_cancelar.Visible = True
    cmd_confirmar.Visible = True
    DataCombo1.Visible = False
    DataGrid2.Visible = True
End Sub

Private Sub cmd_cancelar_Click()
    Call restaurar
End Sub

Private Sub cmd_confirmar_Click()
    Dim frasesql, frasesql2 As String
    Dim fecha As Date
    
    If DataCombo2.Text = "" Or DataCombo3.Text = "" Or DataCombo5.Text = "" Then
        MsgBox "Por favor selecciona todos los datos correspondientes al envio", vbOKOnly, "Gestion"
    Else
        rs_reparto.MoveFirst
        
        Do While Not rs_reparto.EOF
            If rs_reparto.Fields("cod_carga") = DataCombo2.Text And rs_reparto.Fields("asignado") = 0 Then
              fecha = rs_reparto.Fields("fecha")
              Exit Do
            Else
              rs_reparto.MoveNext
            End If
        Loop
        
        frasesql = "UPDATE reparto SET asignado = 1 WHERE cod_carga = '" & DataCombo2.Text & "'"
        cn.Execute (frasesql)
        
        
        frasesql2 = "INSERT INTO pilota VALUES ('" & DataCombo2.BoundText & "','" & DataCombo5.BoundText & "','" & fecha & "')"
        cn.Execute (frasesql2)
        
        MsgBox "Asignacion de envio realizada correctamente", vbOKOnly, "Gestion"
        
        Adodc4.RecordSource = "select reparto.*,carga.proced,carga.destino,carga.descripcion,m_trans.marca,empresa.nom_emp FROM reparto INNER JOIN (carga INNER JOIN (m_trans INNER JOIN empresa ON empresa.codigo = m_trans.cod_emp)ON m_trans.codigo = carga.cod_mtrans) ON reparto.cod_carga = carga.codigo WHERE reparto.asignado = 0 ORDER BY reparto.cod_carga ASC"
        Adodc4.Refresh
        Adodc5.Refresh
        
        DataCombo2.Text = ""
        DataCombo5.Text = ""
        DataCombo3.Visible = False
        DataCombo5.Visible = False
        Adodc3.Refresh
        frm_distribu.Adodc6.Refresh
        
    End If
    
End Sub

Private Sub cmd_volver_Click()
    Call restaurar
    frm_envios.Hide
    Call volver
End Sub

Private Sub cmd_volver_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_volver.BackColor = &HFFFFFF
End Sub

Private Sub cmd_volver_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_volver.BackColor = &H808080
End Sub

Private Sub DataCombo1_Change()
    Adodc2.RecordSource = "select m_trans.marca, carga.descripcion,pilota.fecha FROM reparto INNER JOIN (carga INNER JOIN (m_trans INNER JOIN pilota ON pilota.cod_mtrans = m_trans.codigo) ON m_trans.codigo = carga.cod_mtrans) ON carga.codigo = reparto.cod_carga WHERE pilota.dni_emple = '" & DataCombo1.BoundText & "' AND reparto.fecha = pilota.fecha AND reparto.asignado = 1"
    Adodc2.Refresh
    DataGrid1.Refresh
    DataGrid1.Visible = True
    
    DataGrid1.Columns(0).Caption = "Medio de transporte"
    DataGrid1.Columns(1).Caption = "Carga"
    DataGrid1.Columns(2).Caption = "Fecha del envio"
  
    DataGrid1.Columns(0).Width = 2300
    DataGrid1.Columns(1).Width = 2100
    DataGrid1.Columns(2).Width = 1500
   
    For i = 0 To 2
         DataGrid1.Columns(i).AllowSizing = False
    Next
    
End Sub

Private Sub DataCombo2_Change()
    
    DataCombo3.Visible = True
    
    Adodc6.RecordSource = "select reparto.*,carga.proced,carga.destino,carga.descripcion,m_trans.marca,empresa.nom_emp,empresa.codigo FROM reparto INNER JOIN (carga INNER JOIN (m_trans INNER JOIN empresa ON empresa.codigo = m_trans.cod_emp)ON m_trans.codigo = carga.cod_mtrans) ON reparto.cod_carga = carga.codigo WHERE asignado = 0 and cod_carga = '" & DataCombo2.Text & "' ORDER BY reparto.cod_carga ASC"
    DataCombo3.DataField = "nom_emp"
    DataCombo3.BoundColumn = "codigo"
    DataCombo3.Text = ""
    Adodc6.Refresh
        
End Sub

Private Sub DataCombo3_Change()
     DataCombo5.Visible = True
     Adodc5.RecordSource = "select dni,(nombre+'  '+apellido) as ncompleto,cod_emp from empleado where cod_emp = '" & DataCombo3.BoundText & "'"
     Adodc5.Refresh
End Sub

Private Sub Form_Load()
    
    Set cn = New adodb.Connection
    Set rs_reparto = New adodb.Recordset
    
    Call conexion
    
    rs_reparto.Open "Select * from reparto", cn, adOpenDynamic, adLockOptimistic
    
    Call restaurar
    
End Sub

Private Sub restaurar()
    DataCombo1.Text = ""
    DataGrid1.Visible = False
    lblenvios.Visible = False
    DataCombo2.Visible = False
    DataCombo2.Text = ""
    cmd_asignar.Visible = True
    cmd_cancelar.Visible = False
    cmd_confirmar.Visible = False
    DataCombo1.Visible = True
    DataCombo5.Visible = False
    DataCombo5.Text = ""
    DataGrid2.Visible = False
    DataCombo3.Visible = False
End Sub

