VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_reparación 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reparaciones"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   9690
   StartUpPosition =   2  'CenterScreen
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3000
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   5400
      Top             =   0
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
      Connect         =   "Provider=SQLOLEDB.1;Persist Security Info=True;User ID=sa;Initial Catalog=proyecto1;Data Source=."
      OLEDBString     =   "Provider=SQLOLEDB.1;Persist Security Info=True;User ID=sa;Initial Catalog=proyecto1;Data Source=."
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"frm_reparación.frx":0000
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm_reparación.frx":01C5
      Height          =   2175
      Left            =   480
      TabIndex        =   19
      Top             =   600
      Width           =   8655
      _ExtentX        =   15266
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "nombre"
         Caption         =   "Taller"
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
         DataField       =   "precio"
         Caption         =   "Precio"
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
      BeginProperty Column02 
         DataField       =   "tiempo"
         Caption         =   "Horas"
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
      BeginProperty Column05 
         DataField       =   "descripcion"
         Caption         =   "Reparación"
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
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1305,071
         EndProperty
         BeginProperty Column02 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   599,811
         EndProperty
         BeginProperty Column03 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1305,071
         EndProperty
         BeginProperty Column04 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1305,071
         EndProperty
         BeginProperty Column05 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   6480
      MaxLength       =   7
      TabIndex        =   18
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2400
      MaxLength       =   7
      TabIndex        =   16
      Top             =   2280
      Width           =   1935
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "frm_reparación.frx":01DA
      Height          =   315
      Left            =   2400
      TabIndex        =   14
      Top             =   1200
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "marca"
      BoundColumn     =   "tipo"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   3120
      Top             =   0
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
      RecordSource    =   "select * from  m_trans"
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
      Left            =   1440
      Top             =   0
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
      RecordSource    =   "Select * from m_trans"
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
      Left            =   0
      Top             =   0
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
      RecordSource    =   "select * from reparaciones"
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
   Begin VB.CommandButton cmd_verificar 
      BackColor       =   &H00808080&
      Caption         =   "Verificar"
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmd_cancelar2 
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3000
      Width           =   1455
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "frm_reparación.frx":01EF
      Height          =   315
      Left            =   2400
      TabIndex        =   11
      Top             =   1800
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "nombre"
      BoundColumn     =   "codigo"
      Text            =   ""
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   9
      Top             =   1200
      Width           =   5775
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frm_reparación.frx":0204
      Height          =   315
      Left            =   2400
      TabIndex        =   8
      Top             =   720
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "descripcion"
      BoundColumn     =   "codigo"
      Text            =   ""
   End
   Begin VB.CommandButton cmd_nuevarep 
      BackColor       =   &H00808080&
      Caption         =   "Alta reparación"
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      Width           =   3135
   End
   Begin VB.CommandButton cmd_volver 
      BackColor       =   &H00808080&
      Height          =   615
      Left            =   8280
      Picture         =   "frm_reparación.frx":0219
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmd_cancelar 
      BackColor       =   &H00808080&
      Caption         =   "C&ancelar"
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmd_confirmar 
      BackColor       =   &H00808080&
      Caption         =   "&Confirmar"
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
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lbltiempo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tiempo en horas:"
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
      Left            =   4560
      TabIndex        =   17
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblprecio 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Precio en Euros:"
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
      TabIndex        =   15
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lbldesc 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Descripción:"
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
      Left            =   720
      TabIndex        =   10
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "C.Mantenimiento:"
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
      Left            =   720
      TabIndex        =   6
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Reparación"
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
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.Line Line5 
      X1              =   2400
      X2              =   9240
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line4 
      X1              =   360
      X2              =   960
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line3 
      X1              =   360
      X2              =   9240
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line2 
      X1              =   9240
      X2              =   9240
      Y1              =   360
      Y2              =   3720
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   360
      Y1              =   360
      Y2              =   3720
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "M.Transporte:"
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
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Reparaciones:"
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
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frm_reparación"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim rs_repara As adodb.Recordset
Dim rs_mtrans As adodb.Recordset
Dim rs_reparacion As adodb.Recordset

Private Sub cmd_alta_Click()
    DataGrid1.Visible = False
    cmd_alta.Visible = False
    cmd_nuevarep.Visible = False
End Sub

Private Sub cmd_cancelar_Click()
    Call controlesres
End Sub

Private Sub cmd_cancelar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_cancelar.BackColor = &HFFFFFF
End Sub

Private Sub cmd_cancelar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_cancelar.BackColor = &H808080
End Sub

Private Sub cmd_cancelar2_Click()
    Call controlesres
End Sub

Private Sub cmd_cancelar2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_cancelar2.BackColor = &HFFFFFF
End Sub

Private Sub cmd_cancelar2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_cancelar2.BackColor = &H808080
End Sub

Private Sub cmd_confirmar_Click()
    Dim frasesql As String
    Dim codigotrans As Integer
    Dim fecha As Date
    
    codigotrans = 0
    fecha = Date
    
    If DataCombo1.Text = "" Or DataCombo2.Text = "" Or DataCombo3.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
        MsgBox "No olvides rellenar todos los datos correspondientes a la reparación", vbOKOnly, "Gestión"
    Else
        If Not IsNumeric(Text2.Text) Or Not IsNumeric(Text3.Text) Then
            MsgBox "Los datos del precio en euros y tiempo en horas son numéricos", vbOKOnly, "Gestión"
            Text2.Text = ""
            Text3.Text = ""
        Else
        
        rs_mtrans.MoveFirst
        
        rs_mtrans.Find ("marca = '" & DataCombo3.Text & "'")
        
        If Not rs_mtrans.EOF Then
            codigotrans = rs_mtrans.Fields("codigo")
        End If
            
        rs_repara.MoveFirst
            
            Do While Not rs_repara.EOF
                If rs_repara.Fields("fecha") = Date And rs_repara.Fields("cod_mtrans") = codigotrans And rs_repara.Fields("cod_cmante") = DataCombo2.BoundText Then
                    MsgBox "No se puede realizar esa reparación ya que está actualmente en curso", vbOKOnly, "Gestión"
                    DataCombo3.Text = ""
                    DataCombo2.Text = ""
                    Text2.Text = ""
                    Text3.Text = ""
                    Exit Sub
                Else
                    rs_repara.MoveNext
                End If
            Loop
            
            
        frasesql = "INSERT INTO repara VALUES (" & Text2.Text & ",'" & Text3.Text & "','" & fecha & "','" & DataCombo1.BoundText & "'," & codigotrans & ",'" & DataCombo2.BoundText & "')"
        cn.Execute (frasesql)
        
        MsgBox "Alta realizada con éxito", vbOKOnly, "Gestión"
        Adodc4.Refresh
        
        Call controlesres
        
        End If
    End If
End Sub

Private Sub cmd_confirmar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_confirmar.BackColor = &HFFFFFF
End Sub

Private Sub cmd_confirmar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_confirmar.BackColor = &H808080
End Sub

Private Sub cmd_nuevarep_Click()
    Text1.Visible = True
    lbldesc.Visible = True
    cmd_cancelar2.Visible = True
    cmd_verificar.Visible = True
    
    DataCombo1.Visible = False
    DataCombo2.Visible = False
    DataCombo3.Visible = False
    
    lbltiempo.Visible = False
    lblprecio.Visible = False
    Text2.Visible = False
    Text3.Visible = False
    
    Label1.Visible = False
    Label2.Visible = False
    Label4.Visible = False
    
    DataGrid1.Visible = False
    cmd_alta.Visible = False
    
    cmd_volver.Visible = False
    cmd_confirmar.Visible = False
    cmd_cancelar.Visible = False
    cmd_nuevarep.Visible = False
    
End Sub

Private Sub cmd_nuevarep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_nuevarep.BackColor = &HFFFFFF
End Sub

Private Sub cmd_nuevarep_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_nuevarep.BackColor = &H808080
End Sub

Private Sub cmd_verificar_Click()
    
    Dim frasesql As String
    
    If Text1.Text = "" Then
        MsgBox "Por favor introduce una nuevo tipo de reparación a dar de alta", vbOKOnly, "Gestión"
    Else
        rs_reparacion.MoveFirst
        
        rs_reparacion.Find ("descripcion = '" & Text1.Text & "'")
        
        If rs_reparacion.EOF Then
            frasesql = "INSERT INTO reparaciones VALUES ('" & Text1.Text & "')"
            cn.Execute (frasesql)
            MsgBox "Alta realizada correctamente", vbOKOnly, "Gestión"
            Text1.Text = ""
            Adodc4.Refresh
            Adodc1.Refresh
        Else
            MsgBox "No se realizó el alta ya que la reparación incluida ya existe en la base de datos", vbOKOnly, "Gestión"
            Text1.Text = ""
            Exit Sub
        End If
    End If
    
Call controlesres

End Sub

Private Sub cmd_verificar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_verificar.BackColor = &HFFFFFF
End Sub

Private Sub cmd_verificar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_verificar.BackColor = &H808080
End Sub

Private Sub cmd_volver_Click()
    frm_reparación.Hide
    Call controlesres
    Call volver
End Sub

Private Sub cmd_volver_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_volver.BackColor = &HFFFFFF
End Sub

Private Sub cmd_volver_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_volver.BackColor = &H808080
End Sub

Private Sub DataCombo3_Change()
    DataCombo2.Visible = True
    Label4.Visible = True
    DataCombo2.Text = ""
    
    Adodc3.RecordSource = "Select * from c_mante WHERE c_mante.especialidad = '" & DataCombo3.BoundText & "'"
    
    Adodc3.Refresh
    DataCombo2.Refresh
    
End Sub

Private Sub Form_Load()
    Set cn = New adodb.Connection
    Set rs_repara = New adodb.Recordset
    Set rs_mtrans = New adodb.Recordset
    Set rs_reparacion = New adodb.Recordset

    Call conexion

    rs_repara.Open "select * from repara", cn, adOpenDynamic, adLockOptimistic
    rs_mtrans.Open "select * from m_trans", cn, adOpenDynamic, adLockOptimistic
    rs_reparacion.Open "select * from reparaciones", cn, adOpenDynamic, adLockOptimistic

    Call controlesres
End Sub

Private Sub controlesres()
    Text1.Visible = False
    lbldesc.Visible = False
    cmd_cancelar2.Visible = False
    
    DataCombo1.Visible = True
    DataCombo2.Visible = True
    DataCombo3.Visible = True
    
    Label1.Visible = True
    Label2.Visible = True
    Label4.Visible = True
    
    cmd_volver.Visible = True
    cmd_confirmar.Visible = True
    cmd_cancelar.Visible = True
    cmd_nuevarep.Visible = True
    cmd_verificar.Visible = False
    DataCombo2.Visible = False
    Label4.Visible = False
    
    lbltiempo.Visible = True
    lblprecio.Visible = True
    Text2.Visible = True
    Text3.Visible = True
    
    DataCombo3.Text = ""
    Label4.Visible = False
    DataCombo2.Visible = False
    
    DataGrid1.Visible = True
    cmd_alta.Visible = True
    cmd_nuevarep.Visible = True
    
End Sub
