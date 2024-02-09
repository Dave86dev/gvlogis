VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_Cmante 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Centros de mantenimiento"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4080
      TabIndex        =   30
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5760
      Top             =   3120
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
      RecordSource    =   "select * from empresa"
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
      Bindings        =   "frm_Cmante.frx":0000
      Height          =   315
      Left            =   1320
      TabIndex        =   29
      Top             =   1680
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "nom_emp"
      BoundColumn     =   "codigo"
      Text            =   ""
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3120
      Width           =   1335
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmd_confirmar 
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Frame frm_buscmante 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Búsqueda de centros de mantenimiento"
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
      TabIndex        =   18
      Top             =   3720
      Width           =   7695
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
         Height          =   375
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtbusc 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2520
         TabIndex        =   19
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label lblnombrecentmant 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nombre del centro:"
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
         TabIndex        =   21
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   5160
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   13
      Top             =   480
      Width           =   2535
   End
   Begin VB.CommandButton cmdvolver 
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
      Left            =   6840
      Picture         =   "frm_Cmante.frx":0015
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdmodificar 
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdeliminar 
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdagregar 
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3120
      Width           =   1335
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
      Left            =   3120
      Picture         =   "frm_Cmante.frx":03F0
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
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
      Left            =   2520
      Picture         =   "frm_Cmante.frx":1032
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
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
      Left            =   1920
      Picture         =   "frm_Cmante.frx":1C74
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
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
      Left            =   1320
      Picture         =   "frm_Cmante.frx":28B6
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   615
   End
   Begin VB.Frame Fraesp 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   4920
      TabIndex        =   3
      Top             =   360
      Width           =   2535
      Begin VB.OptionButton opt_fer 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ferroviarios"
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
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton opt_aer 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aéreos"
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
         Top             =   1800
         Width           =   1935
      End
      Begin VB.OptionButton opt_mar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Marítimos"
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
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton opt_rod 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Rodados"
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
         Width           =   1815
      End
      Begin VB.Label lblespe 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Especialidad"
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
         Left            =   360
         TabIndex        =   17
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "C.Mantenimiento"
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
      Left            =   960
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.Line Line5 
      X1              =   840
      X2              =   120
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line4 
      X1              =   7800
      X2              =   2760
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line3 
      X1              =   7800
      X2              =   7800
      Y1              =   240
      Y2              =   2880
   End
   Begin VB.Line Line2 
      X1              =   7800
      X2              =   120
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   120
      Y1              =   240
      Y2              =   2880
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
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lbldirección 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dirección:"
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
      TabIndex        =   1
      Top             =   1080
      Width           =   855
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
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frm_Cmante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs_cmante As adodb.Recordset

Private Sub cmd_buscar_Click()

    If txtbusc.Text = "" Then
        MsgBox "Introduce un nombre de centro a buscar", vbOKOnly, "Gestión"
    Else
        If IsNumeric(Trim(txtbusc.Text)) Then
            MsgBox "El nombre del centro no puede ser numérico", vbOKOnly, "Gestión"
            txtbusc.Text = ""
        Else
            stringbusca = Trim(txtbusc.Text)
            
            rs_cmante.MoveFirst
            
            rs_cmante.Find ("nombre = '" & stringbusca & "'")
            
            If rs_cmante.EOF Then
                MsgBox "El nombre del centro no se encuentra en el registro", vbOKOnly, "Gestión"
                rs_cmante.MoveFirst
                txtbusc.Text = ""
            Else
                Call mostrarmante
                txtbusc.Text = ""
            End If
            
        End If
    End If
End Sub

Private Sub cmd_cancelar_Click()
    Call restaurar
    cmd_mover_Click 0
End Sub

Private Sub cmd_confirmar_Click()
Dim respuesta, frasesql, nombreorig As String

nombreorig = rs_cmante.Fields("nombre")

If Text1(0).Text = "" Then
    MsgBox "El nombre del centro de mantenimiento debe de estar relleno y no ser numérico", vbOKOnly, "Gestión"
    Text1(0).SetFocus
Else

    If (rs_cmante.EOF Or rs_cmante.BOF) Then
        MsgBox "No hay registros activos para modificar", vbOKOnly, "Gestión"
    Else
        respuesta = MsgBox("¿Estas seguro de modificar el registro?", vbDefaultButton2 + vbYesNo, "Gestión")
        
        If respuesta = 6 Then
        
            If rs_cmante.Fields("nombre") <> Text1(0).Text Then
        
                rs_cmante.MoveFirst
            
                rs_cmante.Find ("nombre = '" & Text1(0).Text & "'")
            
                If rs_cmante.EOF Then
            
                    frasesql = "UPDATE c_mante SET nombre = '" & Trim(Text1(0).Text) & "',direc = '" & Trim(Text1(1).Text) & "' WHERE codigo = '" & Trim(Text1(3).Text) & "'"
             
                    cn.Execute (frasesql)
            
                    MsgBox "Modificación realizada con éxito", vbOKOnly, "Gestión"
                
                Else
            
                    MsgBox "No se puede modificar el nombre asignado ya que ya pertenece al de otro taller", vbOKOnly, "Gestion"
                    Text1(0).Text = ""
                Exit Sub
                End If
            Else
                frasesql = "UPDATE c_mante SET nombre = '" & Trim(Text1(0).Text) & "',direc = '" & Trim(Text1(1).Text) & "' WHERE codigo = '" & Trim(Text1(3).Text) & "'"
             
                cn.Execute (frasesql)
            
                MsgBox "Modificación realizada con éxito", vbOKOnly, "Gestión"
            End If
            
            Call restaurar
            
            cmd_mover_Click 0
        End If
    End If
End If
End Sub

Private Sub cmd_confirmaralta_Click()
    Dim frasesql, especialidad As String
    
    If Text1(0).Text = "" Or Text1(1).Text = "" Or DataCombo1.Text = "" Then
        MsgBox "No olvides rellenar los campos correspondientes para el nuevo centro", vbOKOnly, "Gestión"
    Else
        If opt_aer.Value = True Then
            especialidad = "Aéreo"
        ElseIf opt_mar.Value = True Then
            especialidad = "Maritimo"
        ElseIf opt_rod.Value = True Then
            especialidad = "Rodado"
        Else
            especialidad = "Ferroviario"
        End If
        
        rs_cmante.MoveFirst
        
        rs_cmante.Find ("nombre = '" & Text1(0).Text & "'")
        
        If rs_cmante.EOF Then
        
            frasesql = "INSERT INTO c_mante VALUES ('" & Text1(0).Text & "','" & Text1(1).Text & "','" & especialidad & "','" & DataCombo1.BoundText & "')"
        
            cn.Execute frasesql
            
            MsgBox "Alta realizada correctamente", vbOKOnly, "Gestión"
            
        Else
        
            MsgBox "El alta no se puede realizar ya que existe otro centro de mantenimiento con el mismo nombre asignado", vbOKOnly, "Gestion"
            Text1(0).Text = ""
            Exit Sub
        End If
        Call restaurar
        cmd_mover_Click 0
        
    End If
End Sub

Private Sub cmd_mover_Click(Index As Integer)
    On Error Resume Next
    
    If rs_cmante.BOF = True And rs_cmante.EOF = True Then
        Exit Sub
    End If
    
    Select Case Index
        Case 0
            rs_cmante.MoveFirst
        Case 1
            rs_cmante.MovePrevious
        Case 2
            rs_cmante.MoveNext
        Case 3
            rs_cmante.MoveLast
    End Select
    
    If rs_cmante.BOF Then
        MsgBox "Ya está en el primer registro", vbOKOnly, "Advertencia"
        rs_cmante.MoveFirst
    ElseIf rs_cmante.EOF Then
        MsgBox "Ya está en el último registro", vbOKOnly, "Advertencia"
        rs_cmante.MoveLast
    End If
    
    Call mostrarmante
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

Private Sub cmdagregar_Click()
    For i = 0 To 3
        cmd_mover(i).Enabled = False
    Next
    
    For i = 0 To 2
        Text1(i).Locked = False
        Text1(i).Text = ""
    Next
    
    Text1(2).Visible = False
    
    Fraesp.Enabled = True
    frm_buscmante.Enabled = False
    cmdagregar.Visible = False
    cmd_confirmaralta.Visible = True
    cmd_cancelar.Visible = True
    DataCombo1.Visible = True
            
End Sub

Private Sub cmdagregar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdagregar.BackColor = &HFFFFFF
End Sub

Private Sub cmdagregar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdagregar.BackColor = &H808080
End Sub

Private Sub cmdeliminar_Click()
Dim respuesta As String
Dim frasesql, frasesql2 As String

    If (rs_cmante.EOF Or rs_cmante.BOF) Then
        MsgBox "No hay registros activos para eliminar", vbOKOnly, "Gestión"
    Else
        respuesta = MsgBox("¿Estas seguro de borrar el registro?", vbDefaultButton2 + vbYesNo, "Gestión")
        
        If respuesta = 6 Then
            
            frasesql = "DELETE FROM c_mante WHERE codigo = '" & Trim(Text2.Text) & "' AND codigo NOT IN (SELECT cod_cmante FROM repara)"
           
            cn.Execute (frasesql)
            
            If rs_cmante.EOF And rs_cmante.BOF Then
                Exit Sub
            Else
            
                rs_cmante.MoveFirst
            
                rs_cmante.Find ("codigo  = '" & Trim(Text2.Text) & "'")
            
                If rs_cmante.EOF Then

                    MsgBox "Baja realizada con éxito", vbOKOnly, "Gestión"
                
                Else
                    MsgBox "No se puede realizar la baja ya que ese centro de mantenimiento esta en relacion con otros datos", vbOKOnly, "Gestion"
                
                End If
                
            End If
        
            Call restaurar
                
            cmd_mover_Click 0
        End If
    End If
End Sub

Private Sub cmdeliminar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdeliminar.BackColor = &HFFFFFF
End Sub

Private Sub cmdeliminar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdeliminar.BackColor = &H808080
End Sub

Private Sub cmdmodificar_Click()
    frm_buscmante.Enabled = False
    cmdagregar.Visible = False
    cmdeliminar.Visible = False
    cmd_confirmar.Visible = True
    cmd_cancelar.Visible = True
    Fraesp.Enabled = True
    cmdmodificar.Visible = False
    
    For i = 0 To 3
        cmd_mover(i).Enabled = False
    Next
    
    For i = 0 To 1
        Text1(i).Locked = False
    Next
    
    Text1(2).Enabled = False
        
    Text1(0).SetFocus
    
End Sub

Private Sub cmdmodificar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdmodificar.BackColor = &HFFFFFF
End Sub

Private Sub cmdmodificar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdmodificar.BackColor = &H808080
End Sub

Private Sub cmdvolver_Click()
    frm_Cmante.Hide
    Call restaurar
    cmd_mover_Click 0
    Fraesp.Enabled = False
    Call volver
End Sub

Private Sub cmdvolver_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdvolver.BackColor = &HFFFFFF
End Sub

Private Sub cmdvolver_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdvolver.BackColor = &H808080
End Sub

Private Sub Form_Load()
    Set cn = New adodb.Connection
    Set rs_cmante = New adodb.Recordset

    Call conexion
    
    rs_cmante.Open "Select c_mante.codigo,c_mante.nombre,c_mante.direc,especialidad,c_mante.cod_emp,  empresa.nom_emp FROM c_mante INNER JOIN empresa ON empresa.codigo = c_mante.cod_emp", cn, adOpenDynamic, adLockOptimistic
    
    cmd_confirmar.Visible = False
    cmd_cancelar.Visible = False
    cmd_confirmaralta.Visible = False
    
    cmd_mover_Click 0
    
End Sub


Private Sub mostrarmante()
    Text1(0).Text = " "
    Text1(1).Text = " "
    Text1(2).Text = " "
    Text1(3).Text = " "
    
    Fraesp.Enabled = False
    
            With rs_cmante
                Text1(0).Text = .Fields("nombre")
                Text1(1).Text = .Fields("direc")
                Text1(2).Text = .Fields("nom_emp")
                Text1(3).Text = .Fields("cod_emp")
                Text2.Text = .Fields("codigo")
                
                If .Fields("especialidad") = "Maritimo" Then
                    opt_mar.Value = True
                    opt_aer.Value = False
                    opt_rod.Value = False
                    opt_fer.Value = False
                ElseIf .Fields("especialidad") = "Rodado" Then
                    opt_mar.Value = False
                    opt_aer.Value = False
                    opt_rod.Value = True
                    opt_fer.Value = False
                ElseIf .Fields("especialidad") = "Aéreo" Then
                    opt_mar.Value = False
                    opt_aer.Value = True
                    opt_rod.Value = False
                    opt_fer.Value = False
                Else
                    opt_mar.Value = False
                    opt_aer.Value = False
                    opt_rod.Value = False
                    opt_fer.Value = True
                End If
                
            End With
End Sub

Private Sub restaurar()
    frm_buscmante.Enabled = True
    cmdagregar.Visible = True
    cmdeliminar.Visible = True
    cmd_confirmar.Visible = False
    cmd_cancelar.Visible = False
    Fraesp.Enabled = False
    cmd_confirmaralta.Visible = False
    Fraesp.Enabled = False
    DataCombo1.Visible = False
    Text1(2).Visible = True
    cmdmodificar.Visible = True
    
    For i = 0 To 3
        cmd_mover(i).Enabled = True
    Next
    
    For i = 0 To 1
        Text1(i).Locked = True
    Next
    
     Text1(2).Enabled = True
End Sub


            
            
            
            

