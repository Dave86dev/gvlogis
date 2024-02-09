VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frm_distribu 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Distribución"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm_distri.frx":0000
      Height          =   3615
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   9135
      _ExtentX        =   16113
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
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "cod_carga"
         Caption         =   "cod_carga"
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
         Caption         =   "Asig."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "Si"
            FalseValue      =   "No"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   7
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
            Object.Visible         =   0   'False
            ColumnWidth     =   975,118
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column03 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column04 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   599,811
         EndProperty
         BeginProperty Column05 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1305,071
         EndProperty
         BeginProperty Column06 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1305,071
         EndProperty
         BeginProperty Column07 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1904,882
         EndProperty
         BeginProperty Column08 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1305,071
         EndProperty
         BeginProperty Column09 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1094,74
         EndProperty
      EndProperty
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   3375
      Left            =   120
      TabIndex        =   11
      Top             =   3960
      Width           =   7335
      _Version        =   524288
      _ExtentX        =   12938
      _ExtentY        =   5953
      _StockProps     =   1
      BackColor       =   14737632
      Year            =   2008
      Month           =   2
      Day             =   25
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   8160
      Top             =   0
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
      RecordSource    =   "select * from reparto"
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   375
      Left            =   6840
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      RecordSource    =   "select * from carga"
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
      Picture         =   "frm_distri.frx":0015
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6720
      Width           =   735
   End
   Begin VB.Frame frm_carga 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Carga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   5655
      Begin MSDataListLib.DataCombo DataCombo4 
         Bindings        =   "frm_distri.frx":03F0
         Height          =   315
         Left            =   2400
         TabIndex        =   8
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   ""
      End
      Begin VB.Label lblselcar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Seleccione una carga:"
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
         TabIndex        =   7
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame frm_m_trans 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Medio de Transporte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5880
      TabIndex        =   4
      Top             =   1320
      Width           =   3375
      Begin VB.TextBox txttrans 
         Height          =   405
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame frm_itinerarios 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Itinerario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   9135
      Begin VB.TextBox txtdestino 
         Height          =   375
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtprocedencia 
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   600
         Width           =   2055
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   315
         Left            =   360
         TabIndex        =   3
         Top             =   2160
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
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
         Height          =   375
         Left            =   5640
         TabIndex        =   16
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblprocedencia 
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
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmd_añadirdis 
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
      Left            =   720
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.Frame frm_distri 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos de distribución"
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
      TabIndex        =   0
      Top             =   120
      Width           =   9135
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
         Height          =   375
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   480
         Width           =   1695
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
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   480
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frm_distribu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs_datos As adodb.Recordset
Dim rs_datos2 As adodb.Recordset
Dim rs_itinerario As adodb.Recordset
Dim rs_reparto As adodb.Recordset

Dim codigoitine As Integer

Private Sub cmd_añadirdis_Click()
    frm_itinerarios.Visible = True
    frm_m_trans.Visible = True
    frm_carga.Visible = True
    DataGrid1.Visible = False
    cmd_cancelar.Visible = True
    cmd_confirmar.Visible = True
    Calendar1.Visible = True
    cmd_añadirdis.Visible = False
End Sub

Private Sub cmd_añadirdis_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_añadirdis.BackColor = &HFFFFFF
End Sub

Private Sub cmd_añadirdis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_añadirdis.BackColor = &H808080
End Sub

Private Sub cmd_cancelar_Click()
    frm_itinerarios.Visible = False
    frm_m_trans.Visible = False
    frm_carga.Visible = False
    DataGrid1.Visible = True
    cmd_cancelar.Visible = False
    Calendar1.Visible = False
    cmd_añadirdis.Visible = True
    Call restaurar
End Sub

Private Sub cmd_cancelar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_cancelar.BackColor = &HFFFFFF
End Sub

Private Sub cmd_cancelar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_cancelar.BackColor = &H808080
End Sub

Private Sub cmd_confirmar_Click()
    Dim frasesql, frasesql2 As String
    Dim fecha As Date
    Dim itine, codtrans, indice As Integer
    
    indice = 0
    
    If DataCombo4.Text = "" Or Calendar1.Value = "" Then
        MsgBox "Por favor selecciona una carga y una fecha para el reparto", vbOKOnly, "Gestion"
    Else
        
        rs_reparto.MoveFirst
        
        Do While Not rs_reparto.EOF
            If rs_reparto.Fields("cod_carga") = DataCombo4.BoundText And rs_reparto.Fields("fecha") = Calendar1.Value Then
                MsgBox "No se realizo el alta de distribucion ya que ya existe una distribucion igual para esa fecha", vbOKOnly, "Gestion"
                Exit Sub
            Else
                rs_reparto.MoveNext
            End If
        Loop
        
        rs_itinerario.MoveFirst
        
        Do While Not rs_itinerario.EOF
            If rs_itinerario.Fields("punt_part") = txtprocedencia.Text And rs_itinerario.Fields("punt_dest") = txtdestino.Text Then
                itine = rs_itinerario.Fields("codigo")
                Exit Do
            Else
                rs_itinerario.MoveNext
            End If
        Loop
        
        rs_datos2.MoveFirst
        
        Do While Not rs_datos2.EOF
            If rs_datos2.Fields("marca") = txttrans.Text Then
                codtrans = rs_datos2.Fields("codigo")
                Exit Do
            Else
                rs_datos2.MoveNext
            End If
        Loop
        
        If Calendar1.Value < Date Then
        
            MsgBox "La fecha no puede ser inferior a la del dia presente", vbOKOnly, "Gestion"
            Calendar1.Value = Date
            Exit Sub
        Else
            fecha = Calendar1.Value
        End If
        
        rs_reparto.MoveFirst
        
        Do While Not rs_reparto.EOF
            If rs_reparto.Fields("fecha") = fecha And rs_reparto.Fields("cod_mtrans") = codtrans Then
                MsgBox "La distribucion seleccionada se encuentra en curso actualmente en la fecha seleccionada", vbOKOnly, "Gestion"
                Exit Sub
            Else
                rs_reparto.MoveNext
            End If
        Loop
        
        frasesql = "INSERT INTO reparto VALUES ('" & DataCombo4.BoundText & "'," & itine & "," & codtrans & ",'" & fecha & "'," & indice & ")"
        
        cn.Execute (frasesql)
        
        MsgBox "Alta de distribucion realizada correctamente", vbOKOnly, "Gestion"
        
        txttrans.Text = ""
        txtprocedencia.Text = ""
        txtdestino.Text = ""
        Calendar1.Value = Date
        
        Adodc6.RecordSource = "select reparto.*,carga.proced,carga.destino,carga.descripcion,m_trans.marca,empresa.nom_emp FROM reparto INNER JOIN (carga INNER JOIN (m_trans INNER JOIN empresa ON empresa.codigo = m_trans.cod_emp)ON m_trans.codigo = carga.cod_mtrans) ON reparto.cod_carga = carga.codigo ORDER BY reparto.fecha DESC"
        Adodc6.Refresh
        
        cmd_cancelar_Click
        
    End If
End Sub

Private Sub cmd_volver_Click()
    frm_itinerarios.Visible = False
    frm_m_trans.Visible = False
    frm_carga.Visible = False
    DataGrid1.Visible = True
    cmd_cancelar.Visible = False
    frm_distribu.Hide
    Calendar1.Visible = False
    cmd_añadirdis.Visible = True
    Call restaurar
    Call volver
End Sub

Private Sub cmd_volver_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_volver.BackColor = &HFFFFFF
End Sub

Private Sub cmd_volver_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_volver.BackColor = &H808080
End Sub

Private Sub DataCombo4_Change()

    Dim codigotrans As Integer

    rs_datos.MoveFirst

    rs_datos.Find ("codigo = " & Val(DataCombo4.BoundText) & "")

    If rs_datos.EOF Or rs_datos.BOF Then
        Exit Sub
    End If

    txtprocedencia.Text = rs_datos.Fields("proced")
    txtdestino.Text = rs_datos.Fields("destino")
    codigotrans = Val(rs_datos.Fields("cod_mtrans"))

    rs_datos2.MoveFirst

    rs_datos2.Find ("codigo = " & codigotrans & "")

    txttrans.Text = rs_datos2.Fields("marca")

End Sub

Private Sub Form_Load()
    
    Set cn = New adodb.Connection
    Set rs_datos = New adodb.Recordset
    Set rs_datos2 = New adodb.Recordset
    Set rs_itinerario = New adodb.Recordset
    Set rs_reparto = New adodb.Recordset

    Call conexion
    
    rs_datos.Open "Select * from carga", cn, adOpenDynamic, adLockOptimistic
    rs_datos2.Open "select * from m_trans", cn, adOpenDynamic, adLockOptimistic
    rs_itinerario.Open "select * from itinerario", cn, adOpenDynamic, adLockOptimistic
    rs_reparto.Open "select * from reparto", cn, adOpenDynamic, adLockOptimistic
    
    Adodc6.RecordSource = "select reparto.*,carga.proced,carga.destino,carga.descripcion,m_trans.marca,empresa.nom_emp FROM reparto INNER JOIN (carga INNER JOIN (m_trans INNER JOIN empresa ON empresa.codigo = m_trans.cod_emp)ON m_trans.codigo = carga.cod_mtrans) ON reparto.cod_carga = carga.codigo ORDER BY reparto.fecha DESC"
    Adodc6.Refresh
    
    cmd_cancelar.Visible = False
    
    frm_itinerarios.Visible = False
    frm_m_trans.Visible = False
    frm_carga.Visible = False
    cmd_confirmar.Visible = False
    Calendar1.Visible = False
    
    codigoitine = 0
    
End Sub

Private Sub restaurar()
    DataCombo4.Text = ""
    txtprocedencia.Text = ""
    txtdestino.Text = ""
    txttrans.Text = ""
    cmd_confirmar.Visible = False
End Sub

