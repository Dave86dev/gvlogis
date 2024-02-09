VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_m_trans 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medios de transporte"
   ClientHeight    =   6555
   ClientLeft      =   210
   ClientTop       =   975
   ClientWidth     =   9570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   9570
   StartUpPosition =   2  'CenterScreen
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "frm_m_trans.frx":0000
      Height          =   315
      Left            =   2160
      TabIndex        =   35
      Top             =   2760
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "nom_emp"
      BoundColumn     =   "codigo"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   4680
      Top             =   3960
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
   Begin VB.CommandButton cmd_confirmar2 
      BackColor       =   &H00808080&
      Caption         =   "Confirmar Modificación"
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3960
      Width           =   2895
   End
   Begin VB.CommandButton cmd_confirmar 
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
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   4560
      Width           =   2895
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
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5760
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   6720
      Top             =   0
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
      RecordSource    =   "Select  marca,codigo FROM m_trans ORDER BY marca ASC"
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
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   2160
      MaxLength       =   7
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   2160
      MaxLength       =   7
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   1680
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   2160
      MaxLength       =   7
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   2160
      MaxLength       =   20
      TabIndex        =   20
      Top             =   720
      Width           =   3015
   End
   Begin VB.Frame frm_tipo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipo de transporte"
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
      Left            =   6240
      TabIndex        =   19
      Top             =   360
      Width           =   3135
      Begin VB.OptionButton opt_tipo4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ferroviario"
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
         Left            =   240
         TabIndex        =   27
         Top             =   1440
         Width           =   2655
      End
      Begin VB.OptionButton opt_tipo3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Marítimo"
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
         Left            =   240
         TabIndex        =   26
         Top             =   1080
         Width           =   2655
      End
      Begin VB.OptionButton opt_tipo2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aéreo"
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
         Left            =   240
         TabIndex        =   25
         Top             =   720
         Width           =   2655
      End
      Begin VB.OptionButton opt_tipo1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Rodado"
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
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame frm_capa 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Capacidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6240
      TabIndex        =   15
      Top             =   2280
      Width           =   3135
      Begin VB.OptionButton opt_cap3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Más de 100.000 kilos"
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
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   2535
      End
      Begin VB.OptionButton opt_cap2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "De 51.000 a 100.000 kilos"
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
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   2655
      End
      Begin VB.OptionButton opt_cap1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Menos de 50.000 kilos"
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
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   2655
      End
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
      Left            =   8520
      Picture         =   "frm_m_trans.frx":0015
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5760
      Width           =   735
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   3960
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
   Begin VB.CommandButton cmd_eliminar 
      BackColor       =   &H00808080&
      Caption         =   "Baja "
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5160
      Width           =   2895
   End
   Begin VB.CommandButton cmd_modif 
      BackColor       =   &H00808080&
      Caption         =   "&Modificar "
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4560
      Width           =   2895
   End
   Begin VB.CommandButton cmd_nuevo_mtrans 
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
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Añadir medio transporte"
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Frame frm_busc 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Búsqueda de medios de transporte"
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
      Left            =   240
      TabIndex        =   10
      Top             =   4560
      Width           =   5775
      Begin VB.CommandButton cmd_buscar 
         BackColor       =   &H00808080&
         Caption         =   "Buscar"
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
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1080
         Width           =   2295
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frm_m_trans.frx":03F0
         Height          =   315
         Left            =   1920
         TabIndex        =   29
         Top             =   600
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "marca"
         BoundColumn     =   "codigo"
         Text            =   ""
      End
      Begin VB.Label lblmarca2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Marca / Modelo:"
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
         TabIndex        =   28
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   5760
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   480
      Visible         =   0   'False
      Width           =   150
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
      Left            =   3480
      Picture         =   "frm_m_trans.frx":0405
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3600
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
      Left            =   2880
      Picture         =   "frm_m_trans.frx":1047
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3600
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
      Left            =   2280
      Picture         =   "frm_m_trans.frx":1C89
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3600
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
      Left            =   1680
      Picture         =   "frm_m_trans.frx":28CB
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblempresa 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Empresa:"
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
      Left            =   480
      TabIndex        =   34
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Line Line5 
      X1              =   3600
      X2              =   6000
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label lblmtrans2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Medio de transporte"
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
      Left            =   1200
      TabIndex        =   9
      Top             =   240
      Width           =   2295
   End
   Begin VB.Line Line4 
      X1              =   1080
      X2              =   240
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   240
      Y1              =   360
      Y2              =   4320
   End
   Begin VB.Line Line2 
      X1              =   6000
      X2              =   6000
      Y1              =   360
      Y2              =   4320
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   6000
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label lbl_dis_max 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Distancia máxima:"
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
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblpotencia 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Potencia:"
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
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblpeso 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Peso:"
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
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblmarca 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marca:"
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
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "frm_m_trans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs_trans As adodb.Recordset

Private Sub cmd_buscar_Click()

Dim stringbusc

    If DataCombo1.Text = "" Then
        MsgBox "Por favor selecciona un medio de transporte", vbOKOnly, "Gestión"
    Else
        stringbusc = DataCombo1.Text
        
        rs_trans.MoveFirst
        
        rs_trans.Find ("marca = '" & stringbusc & "'")
        
        DataCombo1.Text = ""
        Call mostrartransporte
       
    End If
End Sub



Private Sub cmd_cancelar_Click()
    
    frm_busc.Enabled = True
    
    cmd_modif.Visible = True
    cmd_eliminar.Visible = True
    cmd_confirmar.Visible = False
    cmd_confirmar2.Visible = False
    
    opt_cap1.Enabled = False
    opt_cap2.Enabled = False
    opt_cap3.Enabled = False
    
    opt_tipo1.Enabled = False
    opt_tipo2.Enabled = False
    opt_tipo3.Enabled = False
    opt_tipo4.Enabled = False
    
    For i = 0 To 3
        cmd_mover(i).Enabled = True
    Next
    
    For i = 1 To 4
        Text1(i).Locked = True
    Next
    DataCombo2.Visible = False
    
    cmd_nuevo_mtrans.Visible = True
    lblempresa.Visible = False
    cmd_cancelar.Visible = False
    cmd_mover_Click 0
    Call mostrartransporte
    
End Sub



Private Sub cmd_confirmar_Click()

Dim frasesql As String
Dim tipotrans, capacidad As String

If Text1(1).Text = "" Or Text1(2).Text = "" Or Text1(3).Text = "" Or Text1(4).Text = "" Or DataCombo2.Text = "" Then
    MsgBox "Por favor rellene todos los datos correspondientes del medio de transporte", vbOKOnly, "Gestión"
Else
    If Not IsNumeric(Text1(2).Text) Or Not IsNumeric(Text1(3).Text) Or Not IsNumeric(Text1(4).Text) Then
        MsgBox "Recuerde que los datos de Peso, Potencia y Distancia máxima han de ser numéricos", vbOKOnly, "Gestión"
        Text1(2).Text = ""
        Text1(3).Text = ""
        Text1(4).Text = ""
    Else
        If opt_tipo1.Value = True Then
            tipotrans = "Rodado"
        ElseIf opt_tipo2.Value = True Then
            tipotrans = "Aéreo"
        ElseIf opt_tipo3.Value = True Then
            tipotrans = "Maritimo"
        Else
            tipotrans = "Ferroviario"
        End If
        
        If opt_cap1.Value = True Then
            capacidad = "<50000"
        ElseIf opt_cap2.Value = True Then
            capacidad = "51000-100000"
        Else
            capacidad = ">100000"
        End If
        
        frasesql = "INSERT INTO m_trans VALUES ('" & tipotrans & "','" & capacidad & "','" & Text1(1).Text & "','" & Text1(2).Text & "','" & Text1(3).Text & "','" & Text1(4).Text & "','" & DataCombo2.BoundText & "')"
        
        cn.Execute (frasesql)
        
        MsgBox "Alta realizada correctamente", vbOKOnly, "Gestión"
        Call vuelta
        
    End If
End If
End Sub

Private Sub cmd_confirmar2_Click()
Dim respuesta As String
Dim frasesql As String

If Text1(1).Text = "" Then
    MsgBox "No puede dejar en blanco el nombre del medio de transporte", vbOKOnly, "Gestión"
    Text1(1).SetFocus
Else
    If (rs_trans.EOF Or rs_trans.BOF) Then
        MsgBox "No hay registros activos para modificar", vbOKOnly, "Gestión"
    Else
        respuesta = MsgBox("¿Estas seguro de modificar el registro?", vbDefaultButton2 + vbYesNo, "Gestión")
        
        If respuesta = 6 Then
        
            frasesql = "UPDATE m_trans SET marca = '" & Trim(Text1(1).Text) & "',peso = '" & Trim(Text1(2).Text) & "'," & _
            "potencia = '" & Trim(Text1(3).Text) & "', dis_max = '" & Trim(Text1(4).Text) & "' WHERE codigo = '" & Trim(Text1(0).Text) & "'"
            
            cn.Execute (frasesql)
            
            MsgBox "Modificación realizada con éxito", vbOKOnly, "Gestión"
            
            frm_distribu.Adodc6.Refresh
            frm_reparación.Adodc4.Refresh
            
            cmd_confirmar.Visible = False
            cmd_confirmar2.Visible = False
            cmd_cancelar.Visible = False
            
            opt_cap1.Enabled = True
    
            DataCombo1.Text = ""
            DataCombo1.Refresh
            Adodc2.Refresh
            
            cmd_eliminar.Visible = True
            cmd_modif.Visible = True
            cmd_nuevo_mtrans.Visible = True
    
            Text1(1).Locked = True
            Text1(2).Locked = True
            Text1(3).Locked = True
            Text1(4).Locked = True
            
            frm_busc.Enabled = True
            
            For i = 0 To 3
                cmd_mover(i).Enabled = True
            Next
    
            For i = 1 To 4
                Text1(i).Locked = True
            Next
            
            cmd_mover_Click 0
        End If
    End If
End If
End Sub

Private Sub cmd_eliminar_Click()
Dim respuesta As String
Dim frasesql As String

    If (rs_trans.EOF Or rs_trans.BOF) Then
        MsgBox "No hay registros activos para eliminar", vbOKOnly, "Gestión"
    Else
        respuesta = MsgBox("¿Estas seguro de borrar el registro?", vbDefaultButton2 + vbYesNo, "Gestión")
        
        If respuesta = 6 Then
            frasesql = "DELETE FROM m_trans WHERE codigo = '" & Trim(Text1(0).Text) & "' AND codigo NOT IN (SELECT cod_mtrans FROM repara) AND codigo NOT IN (SELECT cod_mtrans from reparto) AND codigo NOT IN (SELECT cod_mtrans from carga) AND codigo NOT IN (Select cod_mtrans from pilota)"
            cn.Execute (frasesql)
            
            If rs_trans.EOF And rs_trans.BOF Then
            
                Exit Sub
                
            Else
            
                rs_trans.MoveFirst
                
                rs_trans.Find ("codigo = '" & Text1(0).Text & "'")
                
                If rs_trans.EOF Then
                    MsgBox "Baja realizada con éxito", vbOKOnly, "Gestión"
                Else
                    MsgBox "La baja no se puede realizar ya que hay datos relacionados con ese transporte", vbOKOnly, "Gestion"
                End If
                
                Call vuelta
                
                cmd_mover_Click 0
            End If
        End If
    End If
End Sub

Private Sub cmd_eliminar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_eliminar.BackColor = &HFFFFFF
End Sub

Private Sub cmd_eliminar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_eliminar.BackColor = &H808080
End Sub

Private Sub cmd_modif_Click()
    
    cmd_confirmar2.Visible = True
    cmd_cancelar.Visible = True
    cmd_eliminar.Visible = False
    cmd_modif.Visible = False
    
    opt_cap1.Enabled = True
    opt_cap2.Enabled = True
    opt_cap3.Enabled = True
    
    opt_tipo1.Enabled = True
    opt_tipo2.Enabled = True
    opt_tipo3.Enabled = True
    opt_tipo4.Enabled = True
    
    For i = 0 To 3
        cmd_mover(i).Enabled = False
    Next
    
    frm_busc.Enabled = False
    
    Text1(1).SetFocus
    Text1(1).Locked = False
    Text1(2).Locked = False
    Text1(3).Locked = False
    Text1(4).Locked = False
    
End Sub

Private Sub cmd_modif_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_modif.BackColor = &HFFFFFF
End Sub

Private Sub cmd_modif_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_modif.BackColor = &H808080
End Sub

Private Sub cmd_mover_Click(Index As Integer)
    
    On Error Resume Next

    If rs_trans.BOF = True And rs_trans.EOF = True Then
        Exit Sub
    End If
    
    Select Case Index
    
        Case 0
            rs_trans.MoveFirst
        Case 1
            rs_trans.MovePrevious
        Case 2
            rs_trans.MoveNext
        Case 3
            rs_trans.MoveLast
    End Select
    
    If rs_trans.BOF Then
        MsgBox "Ya está en el primer registro", vbOKOnly, "Advertencia"
        rs_trans.MoveFirst
    ElseIf rs_trans.EOF Then
        MsgBox "Ya está en el último registro", vbOKOnly, "Advertencia"
        rs_trans.MoveLast
    End If

    Call mostrartransporte
    
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

Private Sub cmd_nuevo_mtrans_Click()

    frm_busc.Enabled = False
    
    cmd_modif.Visible = False
    cmd_eliminar.Visible = False
    cmd_nuevo_mtrans.Visible = False
    
    lblempresa.Visible = True
    DataCombo2.Visible = True
    
    cmd_cancelar.Visible = True
    cmd_confirmar.Visible = True
    
    opt_tipo1.Enabled = True
    opt_tipo2.Enabled = True
    opt_tipo3.Enabled = True
    opt_tipo4.Enabled = True
    
    opt_cap1.Enabled = True
    opt_cap2.Enabled = True
    opt_cap3.Enabled = True
    
    For i = 1 To 4
        Text1(i).Text = ""
        Text1(i).Locked = False
    Next
    
    For i = 0 To 3
        cmd_mover(i).Enabled = False
    Next
    
End Sub

Private Sub cmd_nuevo_mtrans_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_nuevo_mtrans.BackColor = &HFFFFFF
End Sub

Private Sub cmd_nuevo_mtrans_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_nuevo_mtrans.BackColor = &H808080
End Sub

Private Sub cmd_volver_Click()
    frm_m_trans.Hide
    cmd_confirmar2.Visible = False
    DataCombo1.Text = ""
    cmd_nuevo_mtrans.Visible = True
    Call vuelta
    Call volver
End Sub

Private Sub cmd_volver_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_volver.BackColor = &HFFFFFF
End Sub

Private Sub cmd_volver_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_volver.BackColor = &H808080
End Sub

Private Sub Form_Load()
    
    Set cn = New adodb.Connection
    Set rs_trans = New adodb.Recordset

    Call conexion

    rs_trans.Open "Select * from m_trans", cn, adOpenDynamic, adLockOptimistic
    
    For i = 0 To 4
        Text1(i).Locked = True
    Next
    
    cmd_cancelar.Visible = False
    cmd_confirmar.Visible = False
    cmd_confirmar2.Visible = False
    lblempresa.Visible = False
    DataCombo2.Visible = False
    
    cmd_mover_Click 0
    
End Sub

Private Sub mostrartransporte()

    For i = 0 To 4
        Text1(i).Text = " "
    Next
            With rs_trans
                   Text1(0).Text = .Fields("codigo")
                   Text1(1).Text = .Fields("marca")
                   Text1(2).Text = .Fields("peso")
                   Text1(3).Text = .Fields("potencia")
                   Text1(4).Text = .Fields("dis_max")
                   
                   If .Fields("tipo") = "Rodado" Then
                        opt_tipo1.Value = True
                        opt_tipo1.Enabled = True
                        opt_tipo2.Enabled = False
                        opt_tipo3.Enabled = False
                        opt_tipo4.Enabled = False
                   ElseIf .Fields("tipo") = "Aéreo" Then
                        opt_tipo2.Value = True
                        opt_tipo2.Enabled = True
                        opt_tipo1.Enabled = False
                        opt_tipo3.Enabled = False
                        opt_tipo4.Enabled = False
                   ElseIf .Fields("tipo") = "Maritimo" Then
                        opt_tipo3.Value = True
                        opt_tipo3.Enabled = True
                        opt_tipo1.Enabled = False
                        opt_tipo2.Enabled = False
                        opt_tipo4.Enabled = False
                   Else
                        opt_tipo4.Value = True
                        opt_tipo4.Enabled = True
                        opt_tipo1.Enabled = False
                        opt_tipo2.Enabled = False
                        opt_tipo3.Enabled = False
                   End If
                   
                   
                   If .Fields("capacidad") = "<50000" Then
                        opt_cap1.Value = True
                        opt_cap1.Enabled = True
                        opt_cap2.Enabled = False
                        opt_cap3.Enabled = False
                   ElseIf .Fields("capacidad") = "51000-100000" Then
                        opt_cap2.Value = True
                        opt_cap1.Enabled = False
                        opt_cap2.Enabled = True
                        opt_cap3.Enabled = False
                   Else
                        opt_cap3.Value = True
                        opt_cap1.Enabled = False
                        opt_cap2.Enabled = False
                        opt_cap3.Enabled = True
                   End If
            End With
End Sub

Private Sub vuelta()

    frm_busc.Enabled = True
    
    opt_cap1.Enabled = False
    opt_cap2.Enabled = False
    opt_cap3.Enabled = False
    
    opt_tipo1.Enabled = False
    opt_tipo2.Enabled = False
    opt_tipo3.Enabled = False
    opt_tipo4.Enabled = False
    
    cmd_modif.Visible = True
    cmd_eliminar.Visible = True
    
    For i = 0 To 3
        cmd_mover(i).Enabled = True
    Next
    
    For i = 1 To 4
        Text1(i).Locked = True
    Next
    
    cmd_cancelar.Visible = False
    cmd_confirmar.Visible = False
    
    lblempresa.Visible = False
    DataCombo2.Visible = False
    DataCombo2.Text = ""
    
    Adodc2.Refresh
    
    cmd_mover_Click 0
    
    Call mostrartransporte

End Sub

