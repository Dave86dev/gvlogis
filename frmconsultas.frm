VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmconsultas 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3675
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10590
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra_leyenda 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Leyenda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Label lblleyenda 
         BackColor       =   &H00E0E0E0&
         Caption         =   "-1 = SI              0 = NO"
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
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
   End
   Begin MSDataListLib.DataCombo DComboempresa 
      Bindings        =   "frmconsultas.frx":0000
      DataField       =   "nom_emp"
      Height          =   315
      Left            =   4440
      TabIndex        =   4
      Top             =   600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "nom_emp"
      BoundColumn     =   "codigo"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1680
      Top             =   3000
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      RecordSource    =   "Select empresa.codigo,empresa.nom_emp from empresa ORDER BY nom_emp"
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
      Bindings        =   "frmconsultas.frx":0015
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   10335
      _ExtentX        =   18230
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
      Height          =   375
      Left            =   120
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      RecordSource    =   "Select * from empresa where vol_neg = '>60000' ORDER BY nom_emp"
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
   Begin VB.CommandButton cmdcerrar 
      BackColor       =   &H00808080&
      Caption         =   "&Cerrar"
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label lblselec 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Seleccione el nombre de la empresa"
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
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label lblenunciado 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10575
   End
End
Attribute VB_Name = "frmconsultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcerrar_Click()
    frmlistconsul.Show
    frmconsultas.Hide
    frmconsultas.Caption = ""
    frmconsultas.lblenunciado.Caption = ""
    frmconsultas.lblselec.Visible = False
    frmconsultas.DComboempresa.Visible = False
    frmconsultas.DComboempresa.Text = ""
    frmconsultas.DataGrid1.Visible = False
    fra_leyenda.Visible = False
End Sub

Private Sub cmdcerrar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdcerrar.BackColor = &HEEEEEE
End Sub

Private Sub cmdcerrar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdcerrar.BackColor = &H8000000F
End Sub

Private Sub DComboempresa_Change()

 If lblenunciado.Caption = "Listado de empleados de una empresa." Then
   
   frmconsultas.DataGrid1.Visible = True
   frmconsultas.Adodc1.RecordSource = "Select dni,nombre,apellido,car_rod,car_fer,car_aer,car_mar from  Empleado inner join empresa ON empleado.cod_emp = empresa.codigo Where empleado.cod_emp = '" & frmconsultas.DComboempresa.BoundText & "'"
   frmconsultas.Adodc1.Refresh
   
   frmconsultas.DataGrid1.Columns(0).Caption = "Dni"
   frmconsultas.DataGrid1.Columns(1).Caption = "Nombre"
   frmconsultas.DataGrid1.Columns(2).Caption = "Apellidos"
   frmconsultas.DataGrid1.Columns(3).Caption = "C.Rodado"
   frmconsultas.DataGrid1.Columns(4).Caption = "C.Ferroviario"
   frmconsultas.DataGrid1.Columns(5).Caption = "C.Aéreo"
   frmconsultas.DataGrid1.Columns(6).Caption = "C.Marítimo"
            
   frmconsultas.DataGrid1.Columns(0).Width = 1100
   frmconsultas.DataGrid1.Columns(1).Width = 1400
   frmconsultas.DataGrid1.Columns(2).Width = 2200
   frmconsultas.DataGrid1.Columns(3).Width = 1000
   frmconsultas.DataGrid1.Columns(4).Width = 1400
   frmconsultas.DataGrid1.Columns(5).Width = 1000
   frmconsultas.DataGrid1.Columns(6).Width = 1000
   
   For i = 0 To 6
        frmconsultas.DataGrid1.Columns(i).AllowSizing = False
   Next
End If
 
 If lblenunciado.Caption = "Listado de medios de transporte de una empresa." Then
    
   frmconsultas.DataGrid1.Visible = True
   frmconsultas.Adodc1.RecordSource = "Select tipo,marca,potencia,dis_max  from m_trans inner join empresa ON m_trans.cod_emp = empresa.codigo Where m_trans.cod_emp = '" & frmconsultas.DComboempresa.BoundText & "'"
   frmconsultas.Adodc1.Refresh
   
   frmconsultas.DataGrid1.Columns(0).Caption = "Tipo"
   frmconsultas.DataGrid1.Columns(1).Caption = "Marca"
   frmconsultas.DataGrid1.Columns(2).Caption = "Potencia"
   frmconsultas.DataGrid1.Columns(3).Caption = "Distancia máxima"
   
   frmconsultas.DataGrid1.Columns(0).Width = 1100
   frmconsultas.DataGrid1.Columns(1).Width = 1700
   frmconsultas.DataGrid1.Columns(2).Width = 1600
   frmconsultas.DataGrid1.Columns(3).Width = 2000
   
   For i = 0 To 3
        frmconsultas.DataGrid1.Columns(i).AllowSizing = False
   Next
   
 End If
 
 If lblenunciado.Caption = "Listado de talleres concertados por una empresa." Then
   
   frmconsultas.DataGrid1.Visible = True
   frmconsultas.Adodc1.RecordSource = "Select c_mante.nombre,c_mante.especialidad,c_mante.direc from c_mante inner join empresa ON c_mante.cod_emp = empresa.codigo Where c_mante.cod_emp = '" & frmconsultas.DComboempresa.BoundText & "'"
   frmconsultas.Adodc1.Refresh
   
   frmconsultas.DataGrid1.Columns(0).Caption = "Centro de mantenimiento"
   frmconsultas.DataGrid1.Columns(1).Caption = "Especialidad"
   frmconsultas.DataGrid1.Columns(2).Caption = "Dirección"
   
   frmconsultas.DataGrid1.Columns(0).Width = 3200
   frmconsultas.DataGrid1.Columns(1).Width = 2000
   frmconsultas.DataGrid1.Columns(2).Width = 4000
   
   For i = 0 To 2
        frmconsultas.DataGrid1.Columns(i).AllowSizing = False
   Next
   
 End If

End Sub

Private Sub Form_Load()
    Set cn = New ADODB.Connection
    Call conexion
End Sub
