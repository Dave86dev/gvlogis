VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frm_listadosej 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Muestra"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   10710
   StartUpPosition =   2  'CenterScreen
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   3015
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   6255
      _Version        =   524288
      _ExtentX        =   11033
      _ExtentY        =   5318
      _StockProps     =   1
      BackColor       =   14737632
      Year            =   2008
      Month           =   2
      Day             =   28
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2400
      Top             =   3840
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
      RecordSource    =   "select * from reparto"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   360
      Top             =   3840
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
      RecordSource    =   "select * from carga"
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
      Bindings        =   "frm_listadosej.frx":0000
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm_listadosej.frx":0015
      Height          =   2175
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   3836
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
   Begin VB.CommandButton cmd_cerrar 
      BackColor       =   &H00808080&
      Caption         =   "Cerrar"
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
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label lblenunciado 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   9375
   End
End
Attribute VB_Name = "frm_listadosej"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cerrar_Click()
    frm_listadosej.Hide
    frm_list.Show
    lblenunciado.Caption = ""
    DataCombo1.Text = ""
    DataGrid1.Visible = False
    Calendar1.Visible = False
    cmd_confirmar.Visible = False
End Sub

Private Sub cmd_confirmar_Click()
    Calendar1.Visible = False
    cmd_confirmar.Visible = False
    
    DataGrid1.Visible = True
    Adodc1.RecordSource = "select m_trans.marca,reparto.cod_carga,carga.descripcion,carga.proced,carga.destino FROM carga INNER join (reparto INNER JOIN m_trans ON m_trans.codigo = reparto.cod_mtrans) ON reparto.cod_carga = carga.codigo WHERE reparto.fecha = '" & Calendar1.Value & "' "
    Adodc1.Refresh
    DataGrid1.Refresh
    
    DataGrid1.Columns(0).Caption = "M.transporte"
    DataGrid1.Columns(1).Caption = "Codigo de carga"
    DataGrid1.Columns(2).Caption = "Descripcion"
    DataGrid1.Columns(3).Caption = "Procedencia"
    DataGrid1.Columns(4).Caption = "Destino"
    
    DataGrid1.Columns(1).Width = 1900
    
End Sub

Private Sub DataCombo1_Change()

    If lblenunciado.Caption = "Dado un itinerario, que medios de transporte lo han realizado y en que fechas." Then

        DataGrid1.Visible = True
  
        Adodc1.RecordSource = "select m_trans.marca,m_trans.peso,m_trans.potencia,reparto.fecha FROM m_trans INNER JOIN reparto ON reparto.cod_mtrans = m_trans.codigo WHERE cod_itine = '" & DataCombo1.BoundText & "'"
        Adodc1.Refresh
        DataGrid1.Refresh
        
        DataGrid1.Columns(0).Caption = "M.transporte"
        DataGrid1.Columns(1).Caption = "Peso"
        DataGrid1.Columns(2).Caption = "Potencia"
        DataGrid1.Columns(3).Caption = "Fecha"
       
    ElseIf lblenunciado.Caption = "Dado un medio de transporte, que itinerarios y/o cargas ha realizado." Then

        DataGrid1.Visible = True

        Adodc1.RecordSource = "select reparto.cod_carga,carga.proced,carga.destino,carga.descripcion,reparto.fecha from m_trans inner join (reparto inner join carga on carga.codigo = reparto.cod_carga) on reparto.cod_mtrans = m_trans.codigo WHERE m_trans.codigo = '" & DataCombo1.BoundText & "'"
        Adodc1.Refresh
        DataGrid1.Refresh
        
        DataGrid1.Columns(0).Caption = "Codigo carga"
        DataGrid1.Columns(1).Caption = "Procedencia"
        DataGrid1.Columns(2).Caption = "Destino"
        DataGrid1.Columns(3).Caption = "Descripcion"
        DataGrid1.Columns(4).Caption = "Fecha"

        DataGrid1.Columns(0).Width = 2000
    End If
    
End Sub

Private Sub Form_Load()

End Sub
