VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_carga 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   8325
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   4680
      Top             =   4560
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
      RecordSource    =   "select  distinct tipo from m_trans"
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
   Begin MSDataListLib.DataCombo DataCombo4 
      Bindings        =   "frm_carga.frx":0000
      Height          =   315
      Left            =   2280
      TabIndex        =   17
      Top             =   3960
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "tipo"
      BoundColumn     =   "tipo"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   6600
      Top             =   4440
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      RecordSource    =   "select * from m_trans"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   5640
      Top             =   3240
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      RecordSource    =   "select distinct punt_dest,codigo from itinerario order by punt_dest ASC"
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
      Left            =   5640
      Top             =   2640
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      RecordSource    =   "select distinct punt_part from itinerario order by punt_part ASC"
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
   Begin VB.TextBox txteur 
      Height          =   315
      Left            =   2280
      MaxLength       =   7
      TabIndex        =   16
      Top             =   4680
      Width           =   2055
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "frm_carga.frx":0015
      Height          =   315
      Left            =   5160
      TabIndex        =   15
      Top             =   3960
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "marca"
      BoundColumn     =   "codigo"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "frm_carga.frx":002A
      Height          =   315
      Left            =   2280
      TabIndex        =   14
      Top             =   3240
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "punt_dest"
      BoundColumn     =   "codigo"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frm_carga.frx":003F
      Height          =   315
      Left            =   2280
      TabIndex        =   13
      Top             =   2640
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "punt_part"
      BoundColumn     =   "codigo"
      Text            =   ""
   End
   Begin VB.TextBox txtdesc 
      Height          =   315
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   12
      Top             =   2040
      Width           =   4815
   End
   Begin VB.Frame fra_carga 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Administracion de cargas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8055
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
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton cmd_confirmar 
         BackColor       =   &H00808080&
         Caption         =   "Confirmar "
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
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton cmd_alta 
         BackColor       =   &H00808080&
         Caption         =   "Nueva Carga"
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
         TabIndex        =   4
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmd_volver 
      Height          =   615
      Left            =   7440
      Picture         =   "frm_carga.frx":0054
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm_carga.frx":042F
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   6376
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "codigo"
         Caption         =   "codigo"
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
         DataField       =   "valor"
         Caption         =   "Valor en Euros"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#.##0,00 ""€"""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column05 
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
      BeginProperty Column06 
         DataField       =   "marca"
         Caption         =   "M.Transporte"
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
            Object.Visible         =   0   'False
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1995,024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1305,071
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1305,071
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1395,213
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1395,213
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   5640
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
      RecordSource    =   "select carga.*,m_trans.marca from carga inner join m_trans on carga.cod_mtrans = m_trans.codigo"
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
   Begin VB.Label lblvalor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Valor en Euros:"
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
      Left            =   120
      TabIndex        =   11
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label lblmtrans 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Medios de transporte:"
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
      TabIndex        =   10
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label lbldestino 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Destino:"
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
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label lblproced 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Procedencia:"
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
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label lbldesc 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Descripcion:"
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
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lbldatoscargas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos de cargas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
End
Attribute VB_Name = "frm_carga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_carga As adodb.Recordset

Private Sub cmd_alta_Click()
    cmd_confirmar.Visible = True
    cmd_cancelar.Visible = True
    cmd_alta.Visible = False
    DataGrid1.Visible = False
    DataCombo1.Visible = True
    DataCombo3.Visible = False
    DataCombo2.Visible = False
    DataCombo4.Visible = True
    txtdesc.Visible = True
    txteur.Visible = True
End Sub

Private Sub cmd_cancelar_Click()
    Call restaurar
End Sub

Private Sub cmd_confirmar_Click()
Dim frasesql As String

    If txteur.Text = "" Or txtdesc.Text = "" Or DataCombo1.Text = "" Or DataCombo2.Text = "" Or DataCombo3.Text = "" Then
        MsgBox "Debe de rellenar todos los campos para dar de alta una carga", vbOKOnly, "Gestion"
    Else
        If Not IsNumeric(txteur.Text) Then
            MsgBox "El valor introducido debe de ser una cifra numerica", vbOKOnly, "Gestion"
        Else
        
            rs_carga.MoveFirst
            
            Do While Not rs_carga.EOF
                If rs_carga.Fields("descripcion") = txtdesc.Text And rs_carga.Fields("proced") = DataCombo1.Text And rs_carga.Fields("destino") = DataCombo2.Text And rs_carga.Fields("valor") = txteur.Text And rs_carga.Fields("cod_mtrans") = DataCombo3.BoundText Then
                    MsgBox "Ya existe una carga con el itinerario, descripcion, precio y medio de transporte asignado identica", vbOKOnly, "Gestion"
                    txtdesc.Text = ""
                    DataCombo1.Text = ""
                    DataCombo2.Text = ""
                    txteur.Text = ""
                    DataCombo3.BoundText = ""
                    Exit Sub
                Else
                    rs_carga.MoveNext
                End If
            Loop
            
            frasesql = "INSERT INTO carga VALUES ('" & txtdesc.Text & "','" & DataCombo1.Text & "','" & DataCombo2.Text & "'," & txteur.Text & ",'" & DataCombo3.BoundText & "')"
            
            cn.Execute (frasesql)
            
            MsgBox "Alta realizada correctamente", vbOKOnly, "Gestion"
            
            Adodc1.RecordSource = "select carga.*,m_trans.marca from carga inner join m_trans on carga.cod_mtrans = m_trans.codigo"
            Adodc1.Refresh
            DataGrid1.Refresh
            frm_distribu.Adodc4.Refresh
            
            Call restaurar
            
        End If
    End If
End Sub

Private Sub cmd_volver_Click()
    Call restaurar
    Call volver
End Sub

Private Sub DataCombo1_Change()
    DataCombo2.Visible = True
    Adodc3.RecordSource = "select punt_dest from itinerario where punt_part = '" & DataCombo1.Text & "'"
    Adodc3.Refresh
    DataCombo2.Text = ""
End Sub

Private Sub DataCombo4_Change()
    DataCombo3.Visible = True
    Adodc4.RecordSource = "select marca,codigo from m_trans where tipo = '" & DataCombo4.Text & "'"
    DataCombo3.BoundColumn = "codigo"
    Adodc4.Refresh
    DataCombo3.Text = ""
End Sub


Private Sub Form_Load()

    Set cn = New adodb.Connection
    
    Call conexion
    
    Set rs_carga = New adodb.Recordset
    
    rs_carga.Open "select * from carga", cn, adOpenDynamic, adLockOptimistic
    
    Call restaurar
End Sub

Private Sub restaurar()
    cmd_confirmar.Visible = False
    cmd_cancelar.Visible = False
    cmd_alta.Visible = True
    DataGrid1.Visible = True

    txtdesc.Visible = False
    txteur.Visible = False
    
    DataCombo1.Text = ""
    DataCombo2.Text = ""
    DataCombo3.Text = ""
    DataCombo4.Text = ""
    DataCombo1.Visible = False
    DataCombo3.Visible = False
    DataCombo4.Visible = False
    DataCombo2.Visible = False
    txtdesc.Text = ""
    txteur.Text = ""
End Sub
