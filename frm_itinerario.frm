VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_itinerario 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Itinerarios"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5370
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
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
      Left            =   4440
      Picture         =   "frm_itinerario.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4920
      Width           =   735
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   1920
      Top             =   4920
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
      RecordSource    =   "Select DISTINCT punt_dest FROM itinerario ORDER BY punt_dest ASC"
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
      Left            =   3240
      Top             =   4920
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
      RecordSource    =   "Select DISTINCT punt_part FROM itinerario ORDER BY punt_part ASC"
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
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   3000
      Width           =   975
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "frm_itinerario.frx":03DB
      Height          =   315
      Left            =   1320
      TabIndex        =   10
      Top             =   2280
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "punt_dest"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frm_itinerario.frx":03F0
      DataField       =   "punt_part"
      Height          =   315
      Left            =   1320
      TabIndex        =   8
      Top             =   1560
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "punt_part"
      BoundColumn     =   "punt_part"
      Text            =   ""
   End
   Begin VB.Frame frm_gestiti 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Gestión de itinerarios"
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
      TabIndex        =   2
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton cmd_cancelar 
         BackColor       =   &H00808080&
         Caption         =   "&Cancelar"
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
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmd_guardar 
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmd_modificariti 
         BackColor       =   &H00808080&
         Caption         =   "&Modificar"
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
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmd_anadiriti 
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame frm_itinedisp 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Itinerarios disponibles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   5055
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frm_itinerario.frx":0405
         Height          =   2775
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   4895
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "punt_part"
            Caption         =   "Punto de partida"
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
            DataField       =   "punt_dest"
            Caption         =   "Punto de destino"
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
            DataField       =   "indice"
            Caption         =   "Índice"
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
               ColumnWidth     =   1604,976
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1604,976
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   900,284
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   4920
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      RecordSource    =   "Select * from itinerario"
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
   Begin VB.Label lblcalificacion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Calificación:"
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
      Top             =   3000
      Width           =   1095
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
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label lblorigen 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Origen:"
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
      Top             =   1560
      Width           =   735
   End
End
Attribute VB_Name = "frm_itinerario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs_itine As adodb.Recordset

Private Sub cmd_anadiriti_Click()
        
    cmd_anadiriti.Visible = False
    cmd_modificariti.Visible = False
    
    cmd_volver.Visible = False
    
    Call botones_anadmos
    
End Sub

Private Sub cmd_anadiriti_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_anadiriti.BackColor = &HFFFFFF
End Sub

Private Sub cmd_anadiriti_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_anadiriti.BackColor = &H808080
End Sub

Private Sub cmd_borrariti_Click()

Dim respuesta As String

    respuesta = MsgBox("¿Esta seguro de proceder a borrar?", vbDefaultButton2 + vbYesNo, "Gestión")
    
    If respuesta = 6 Then
        MsgBox "Borre el registro seleccionado manualmente pulsando suprimir", vbOKOnly, "Gestión"
        DataGrid1.AllowDelete = True
    End If
End Sub

Private Sub cmd_cancelar_Click()
    Call botones_anadoc
    cmd_anadiriti.Visible = True
    cmd_modificariti.Visible = True
    cmd_volver.Visible = True
    DataCombo1.Text = ""
    DataCombo2.Text = ""
    DataGrid1.AllowUpdate = False
    cmd_anadiriti.Visible = True
End Sub

Private Sub cmd_cancelar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_cancelar.BackColor = &HFFFFFF
End Sub

Private Sub cmd_cancelar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_cancelar.BackColor = &H808080
End Sub

Private Sub cmd_guardar_Click()

Dim frasesql As String

    If DataCombo1.Text = "" Or DataCombo2.Text = "" Or Combo1.Text = "" Then
        MsgBox "Por favor acaba de rellenar los datos del itinerario", vbOKOnly, "Gestión"
    Else
    
        rs_itine.MoveFirst
        
        Do While Not rs_itine.EOF
            If rs_itine.Fields("punt_part") = DataCombo1.Text And rs_itine.Fields("punt_dest") = DataCombo2.Text Then
                MsgBox "No se puede realizar alta de itinerario ya que la procedencia y destino asignados pertenece a otro itinerario", vbOKOnly, "Gestión"
                DataCombo1.Text = ""
                DataCombo2.Text = ""
                
                Exit Sub
            Else
                rs_itine.MoveNext
            End If
        Loop
        
        frasesql = "Insert into itinerario values ('" & DataCombo1.Text & "', '" & DataCombo2.Text & "','" & Combo1.Text & "')"
        
        cn.Execute (frasesql)
        MsgBox "Alta de itinerario realizada con éxito", vbOKOnly, "Gestión"
        DataCombo1.Text = ""
        DataCombo2.Text = ""
        
        Adodc1.Refresh
        DataGrid1.Refresh
        
    End If
End Sub

Private Sub cmd_guardar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_guardar.BackColor = &HFFFFFF
End Sub

Private Sub cmd_guardar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_guardar.BackColor = &H808080
End Sub

Private Sub cmd_modificariti_Click()
    cmd_cancelar.Visible = True
    cmd_modificariti.Visible = False
    DataGrid1.AllowUpdate = True
    cmd_anadiriti.Visible = False
    MsgBox "Ahora ya puede realizar modificaciones en los itinerarios", vbOKOnly, "Gestión"
End Sub

Private Sub cmd_modificariti_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_modificariti.BackColor = &HFFFFFF
End Sub

Private Sub cmd_modificariti_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_modificariti.BackColor = &H808080
End Sub

Private Sub cmd_volver_Click()
    frm_itinerario.Hide
    cmd_anadiriti.Visible = True
    cmd_modificariti.Visible = True
    cmd_volver.Visible = True
    
    DataGrid1.AllowUpdate = False
    cmd_anadiriti.Visible = True
    Call volver
End Sub

Private Sub cmd_volver_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_volver.BackColor = &HFFFFFF
End Sub

Private Sub cmd_volver_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_volver.BackColor = &H808080
End Sub

Private Sub DataGrid1_AfterDelete()
    DataGrid1.Refresh
    DataGrid1.AllowDelete = False
End Sub

Private Sub DataGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
rs_itine.MoveFirst
            
    Do While Not rs_itine.EOF
            If rs_itine.Fields("punt_part") = DataGrid1.Columns(0).Text And rs_itine.Fields("punt_dest") = DataGrid1.Columns(1).Text Then
                MsgBox "No se puede modificar el itinerario ya que la procedencia y destino asignados pertenece a otro itinerario", vbOKOnly, "Gestión"
                DataGrid1.Columns(0).Text = ""
                DataGrid1.Columns(1).Text = ""
                DataGrid1.AllowUpdate = False
                DataGrid1.Refresh
                Exit Sub
                
            Else
                rs_itine.MoveNext
            End If
    Loop
End Sub

Private Sub Form_Load()

    Set cn = New adodb.Connection
    
    Call conexion
    
    Set rs_itine = New adodb.Recordset
    
    rs_itine.Open "select * from itinerario", cn, adOpenDynamic, adLockOptimistic
    
    For i = 1 To 10
        Combo1.AddItem i
    Next
    
    Call botones_anadoc
    
End Sub

Private Sub botones_anadoc()
    cmd_cancelar.Visible = False
    cmd_guardar.Visible = False
    lblorigen.Visible = False
    DataCombo1.Visible = False
    lbldestino.Visible = False
    DataCombo2.Visible = False
    lblcalificacion.Visible = False
    Combo1.Visible = False
    frm_itinedisp.Visible = True
End Sub

Private Sub botones_anadmos()
    cmd_cancelar.Visible = True
    cmd_guardar.Visible = True
    lblorigen.Visible = True
    DataCombo1.Visible = True
    lbldestino.Visible = True
    DataCombo2.Visible = True
    lblcalificacion.Visible = True
    Combo1.Visible = True
    frm_itinedisp.Visible = False
End Sub

