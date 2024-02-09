VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_empleado 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empleado"
   ClientHeight    =   6090
   ClientLeft      =   4635
   ClientTop       =   2595
   ClientWidth     =   9735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   9735
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   7320
      Top             =   3360
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
      RecordSource    =   "Select * from empresa"
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frm_empleado.frx":0000
      Height          =   315
      Left            =   1800
      TabIndex        =   33
      Top             =   1800
      Width           =   2055
      _ExtentX        =   3625
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtdni 
      Height          =   285
      Left            =   1800
      TabIndex        =   31
      Text            =   "Text2"
      Top             =   120
      Visible         =   0   'False
      Width           =   255
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3360
      Width           =   2175
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
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Frame frabuscemp 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Búsqueda de empleados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      TabIndex        =   23
      Top             =   3960
      Width           =   9255
      Begin VB.CommandButton cmd_buscarsig 
         BackColor       =   &H00808080&
         Caption         =   "Buscar siguiente"
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
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1200
         Visible         =   0   'False
         Width           =   1935
      End
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
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton opt_buscdni 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar por dni"
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
         Width           =   1935
      End
      Begin VB.OptionButton opt_nombre 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar por nombre"
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
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtbusc 
         Height          =   375
         Left            =   2880
         TabIndex        =   24
         Top             =   480
         Width           =   3735
      End
   End
   Begin VB.TextBox txt_carnet 
      Height          =   285
      Left            =   120
      TabIndex        =   21
      Text            =   "Text2"
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmd_baja 
      BackColor       =   &H00808080&
      Caption         =   "&Baja"
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
      TabIndex        =   19
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmd_alta 
      BackColor       =   &H00808080&
      Caption         =   "&Alta"
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
      TabIndex        =   18
      Top             =   3360
      Width           =   2175
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
      Left            =   8760
      Picture         =   "frm_empleado.frx":0015
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3240
      Width           =   735
   End
   Begin VB.Frame frm_carnet 
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
      Height          =   2535
      Left            =   6480
      TabIndex        =   12
      Top             =   600
      Width           =   3015
      Begin VB.CheckBox chk_fer 
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
         TabIndex        =   16
         Top             =   1920
         Width           =   2295
      End
      Begin VB.CheckBox chk_mar 
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
         TabIndex        =   15
         Top             =   1440
         Width           =   2295
      End
      Begin VB.CheckBox chk_aer 
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
         TabIndex        =   14
         Top             =   960
         Width           =   2295
      End
      Begin VB.CheckBox chk_rod 
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
         TabIndex        =   13
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lbltipcar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tipos de carnet"
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
         TabIndex        =   22
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   1800
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   1800
      MaxLength       =   8
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   3480
      MaxLength       =   35
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   840
      Width           =   1695
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
      Left            =   3720
      Picture         =   "frm_empleado.frx":03F0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
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
      Left            =   3120
      Picture         =   "frm_empleado.frx":1032
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
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
      Left            =   2520
      Picture         =   "frm_empleado.frx":1C74
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
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
      Left            =   1920
      Picture         =   "frm_empleado.frx":28B6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   615
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4800
      Top             =   2640
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
   Begin VB.Line Line6 
      X1              =   2160
      X2              =   6360
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line5 
      X1              =   720
      X2              =   240
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblemple 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Empleados"
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
      Left            =   840
      TabIndex        =   11
      Top             =   480
      Width           =   1215
   End
   Begin VB.Line Line4 
      X1              =   6360
      X2              =   6360
      Y1              =   600
      Y2              =   720
   End
   Begin VB.Line Line3 
      X1              =   6360
      X2              =   6360
      Y1              =   720
      Y2              =   3120
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   240
      Y1              =   600
      Y2              =   3120
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   6360
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label lblemp 
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
      TabIndex        =   10
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lbldni 
      BackColor       =   &H00E0E0E0&
      Caption         =   "DNI:"
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
      TabIndex        =   9
      Top             =   1440
      Width           =   735
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
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "frm_empleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs_emple As adodb.Recordset

Private Sub cmd_alta_Click()

    Text1(3).Visible = False
    DataCombo1.Visible = True
    
    For i = 0 To 3
        cmd_mover(i).Enabled = False
    Next
    
    For i = 0 To 3
        Text1(i).Locked = False
        Text1(i).Text = ""
    Next
    
    cmd_confirmaralta.Visible = True
    cmd_modif.Visible = False
    cmd_alta.Visible = False
    cmd_cancelar.Visible = True
    frabuscemp.Enabled = False
    frm_carnet.Enabled = True
    chk_rod.Value = 0
    chk_mar.Value = 0
    chk_aer.Value = 0
    chk_fer.Value = 0
    
End Sub

Private Sub cmd_alta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_alta.BackColor = &HFFFFFF
End Sub

Private Sub cmd_alta_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_alta.BackColor = &H808080
End Sub

Private Sub cmd_baja_Click()
Dim respuesta As String
Dim frasesql, frasesql2, frasesql3 As String

    If (rs_emple.EOF Or rs_emple.BOF) Then
        MsgBox "No hay registros activos para eliminar", vbOKOnly, "Gestión"
    Else
        respuesta = MsgBox("¿Estas seguro de borrar el registro?", vbDefaultButton2 + vbYesNo, "Gestión")
        
        If respuesta = 6 Then
            frasesql2 = "Delete from pilota where dni_emple = '" & Text1(2).Text & "'"
            frasesql = "DELETE FROM empleado WHERE dni = '" & Trim(txtdni.Text) & "'"
            cn.Execute (frasesql2)
            cn.Execute (frasesql)
            MsgBox "Baja realizada con éxito", vbOKOnly, "Gestión"
            frasesql3 = "UPDATE empresa SET num_emple = num_emple-1 WHERE nom_emp = '" & Text1(3).Text & "'"
            cn.Execute (frasesql3)
            cmd_mover_Click 0
        End If
    End If
End Sub

Private Sub cmd_baja_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_baja.BackColor = &HFFFFFF
End Sub

Private Sub cmd_baja_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_baja.BackColor = &H808080
End Sub

Private Sub cmd_buscar_Click()

    If txtbusc.Text = "" Then
        MsgBox "Rellena en el cuadro de texto el empleado a buscar", vbOKOnly, "Gestión"
    Else
        If opt_nombre.Value = False And opt_buscdni.Value = False Then
            MsgBox "Elige un método de búsqueda", vbOKOnly, "Gestión"
        Else
            If opt_nombre.Value = True Then
                If IsNumeric(txtbusc.Text) Then
                    MsgBox "Recuerda que el nombre no contiene cifras numéricas", vbOKOnly, "Gestión"
                    txtbusc.Text = ""
                Else
                    
                    stringbusca = Trim(txtbusc.Text)
                    rs_emple.MoveFirst
                    
                    rs_emple.Find ("nombre = '" & stringbusca & "'")
                    
                    If rs_emple.EOF Then
                        MsgBox "El nombre no se encuentra en el registro", vbOKOnly, "Gestión"
                        rs_emple.MoveFirst
                    Else
                        
                        Call mostrarempleado
                        txtbusc.Text = ""
                        cmd_buscarsig.Visible = True
                    End If
                  End If
                End If
            End If
         
            If opt_buscdni.Value = True Then
                If Not IsNumeric(txtbusc.Text) Then
                    MsgBox "Recuerda que el DNI debe de ser una cifra numérica", vbOKOnly, "Gestión"
                    txtbusc.Text = ""
                Else
                    
                    stringbusca = Trim(txtbusc.Text)
                    rs_emple.MoveFirst
                    
                    rs_emple.Find ("dni = '" & stringbusca & "'")
                    
                    If rs_emple.EOF Then
                        MsgBox "El Dni no se encuentra en el registro", vbOKOnly, "Gestión"
                        rs_emple.MoveFirst
                    Else
                        
                        Call mostrarempleado
                        txtbusc.Text = ""
                    End If
                End If
            End If
        End If
End Sub

Private Sub cmd_buscarsig_Click()
    If opt_nombre.Value = False And opt_buscdni.Value = False Then
            MsgBox "Elige un método de búsqueda", vbOKOnly, "Gestión"
    Else
            If opt_nombre.Value = True Then
                If IsNumeric(txtbusc.Text) Then
                    MsgBox "Recuerda que el nombre no contiene cifras numéricas", vbOKOnly, "Gestión"
                    txtbusc.Text = ""
                Else
                    
                    rs_emple.Find ("nombre = '" & stringbusca & "'"), 1, adSearchForward
                    
                    If rs_emple.EOF Then
                        MsgBox "El nombre no se encuentra en el registro", vbOKOnly, "Gestión"
                        rs_emple.MoveFirst
                    Else
                        Call mostrarempleado
                        txtbusc.Text = ""
                    End If
                End If
            End If
         
            If opt_buscdni.Value = True Then
                MsgBox "Recuerda que los DNI son unicos, por lo tanto no hay mas registros iguales", vbOKOnly, "Gestion"
                txtbusc.Text = ""
            End If
    End If
End Sub

Private Sub cmd_cancelar_Click()
    Call restablecer
End Sub

Private Sub cmd_confirmar2_Click()
Dim respuesta As String
Dim frasesql As String

If Text1(0).Text = "" Or IsNumeric(Text1(0).Text) Then
    MsgBox "El nombre del empleado debe estar relleno y no debe de ser numérico", vbOKOnly, "Gestión"
    Text1(0).SetFocus
Else
    If (rs_emple.EOF Or rs_emple.BOF) Then
        MsgBox "No hay registros activos para modificar", vbOKOnly, "Gestión"
    Else
        respuesta = MsgBox("¿Estas seguro de modificar el registro?", vbDefaultButton2 + vbYesNo, "Gestión")
        
        If respuesta = 6 Then
            frasesql = "UPDATE empleado SET nombre = '" & Trim(Text1(0).Text) & "',apellido = '" & Trim(Text1(1).Text) & "' WHERE dni = '" & Trim(txtdni.Text) & "'"
            
            cn.Execute (frasesql)
            
            MsgBox "Modificación realizada con éxito", vbOKOnly, "Gestión"
            
            cmd_confirmar2.Visible = False
            cmd_cancelar.Visible = False
            
            cmd_baja.Visible = True
            cmd_alta.Visible = True
            
            Text1(0).Locked = True
            Text1(1).Locked = True
            Text1(3).Locked = True
            
            For i = 0 To 3
                cmd_mover(i).Enabled = True
            Next
            
            cmd_mover_Click 0
        End If
    End If
End If
End Sub

Private Sub cmd_confirmaralta_Click()
    Dim frasesql, frasesql2, especialidad As String
    Dim aer, fer, rod, mar As Integer
    
    aer = 0
    fer = 0
    rod = 0
    mar = 0
    
    cmd_confirmaralta.Visible = False
    
    If Text1(0).Text = "" Or Text1(1).Text = "" Or Text1(2).Text = "" Or DataCombo1.Text = "" Then
        MsgBox "No olvide rellenar los datos de nuevo empleado", vbOKOnly, "Gestión"
        cmd_confirmaralta.Visible = True
        Exit Sub
    Else
        If Not IsNumeric(Text1(2).Text) Then
            MsgBox "El Dni debe de ser numérico obligatoriamente", vbOKOnly, "Gestión"
            cmd_confirmaralta.Visible = True
            Exit Sub
        Else
            If chk_rod.Value = 0 And chk_fer.Value = 0 And chk_aer.Value = 0 And chk_mar.Value = 0 Then
                MsgBox "Por favor selecciona una especialidad de pilotar para el nuevo empleado", vbOKOnly, "Gestión"
                cmd_confirmaralta.Visible = True
                Exit Sub
            Else
                If chk_rod.Value = 1 Then
                    rod = 1
                End If
                
                If chk_aer.Value = 1 Then
                    aer = 1
                End If
                
                If chk_mar.Value = 1 Then
                    mar = 1
                End If
                
                If chk_fer.Value = 1 Then
                    fer = 1
                End If
                
                rs_emple.MoveFirst
                
                rs_emple.Find ("dni = '" & Text1(2).Text & "'")
                
                If rs_emple.EOF Then
                
                    frasesql = "INSERT INTO empleado VALUES ('" & Text1(2).Text & "','" & Text1(0).Text & "','" & Text1(1).Text & "'," & Val(rod) & "," & Val(fer) & "," & Val(aer) & "," & Val(mar) & ",'" & DataCombo1.BoundText & "')"
                
                    cn.Execute (frasesql)
                
                    MsgBox "Alta realizada correctamente", vbOKOnly, "Gestión"
                    frasesql2 = "UPDATE empresa SET num_emple = num_emple+1 WHERE codigo = '" & DataCombo1.BoundText & "'"
                    cn.Execute (frasesql2)
                    
                Else
                
                    MsgBox "No se va a realizar el alta ya que el DNI introducido ya corresponde al de un empleado", vbOKOnly, "Gestión"
                    Text1(2).Text = ""
                    cmd_confirmaralta.Visible = True
                    Exit Sub
                End If
                
                Call restablecer
                
            End If
        End If
    End If
End Sub

Private Sub cmd_modif_Click()

    For i = 0 To 3
        cmd_mover(i).Enabled = False
    Next
    
    For i = 0 To 3
        Text1(i).Locked = False
    Next
    
    Text1(2).Locked = True
    Text1(2).Enabled = False
    Text1(0).SetFocus
    
    frabuscemp.Enabled = False
    frm_carnet.Enabled = True
    cmd_alta.Visible = False
    cmd_baja.Visible = False
    cmd_cancelar.Visible = True
    cmd_confirmar2.Visible = True
    cmd_modif.Visible = False
End Sub

Private Sub cmd_modif_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_modif.BackColor = &HFFFFFF
End Sub

Private Sub cmd_modif_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_modif.BackColor = &H808080
End Sub

Private Sub cmd_mover_Click(Index As Integer)
    On Error Resume Next
    
    If rs_emple.BOF = True And rs_emple.EOF = True Then
        Exit Sub
    End If
    
    Select Case Index
    
        Case 0
            rs_emple.MoveFirst
        Case 1
            rs_emple.MovePrevious
        Case 2
            rs_emple.MoveNext
        Case 3
            rs_emple.MoveLast
    
    End Select
    
    If rs_emple.BOF Then
        MsgBox "Ya está en el primer registro", vbOKOnly, "Advertencia"
        rs_emple.MoveFirst
    ElseIf rs_emple.EOF Then
        MsgBox "Ya está en el último registro", vbOKOnly, "Advertencia"
        rs_emple.MoveLast
    End If
    
    Call mostrarempleado
    
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

Private Sub cmd_volver_Click()
    frm_empleado.Hide
    frm_carnet.Enabled = False
    cmd_confirmar2.Visible = False
    cmd_cancelar.Visible = False
    frabuscemp.Enabled = True
    Text1(2).Enabled = True
    
    Call restablecer
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
    Set rs_emple = New adodb.Recordset
    
    Call conexion
    
    rs_emple.Open "Select empleado.dni, empleado.nombre, empleado.apellido, empleado.car_rod, empleado.car_fer, empleado.car_aer, empleado.car_mar, empleado.cod_emp, empresa.nom_emp, empresa.codigo FROM empleado INNER JOIN empresa ON (empleado.cod_emp = empresa.codigo)", cn, adOpenDynamic, adLockOptimistic
    
    For i = 0 To 3
        Text1(i).Locked = True
    Next
    
    cmd_confirmar2.Visible = False
    cmd_cancelar.Visible = False
    cmd_confirmaralta.Visible = False
    DataCombo1.Visible = False
    
    
    cmd_mover_Click 0
    
End Sub

Private Sub mostrarempleado()

    For i = 0 To 3
        Text1(i).Text = " "
    Next
    
    chk_fer.Value = 0
    chk_rod.Value = 0
    chk_aer.Value = 0
    chk_mar.Value = 0
    
    frm_carnet.Enabled = False
    
        With rs_emple
            Text1(0).Text = .Fields("nombre")
            Text1(1).Text = .Fields("apellido")
            Text1(2).Text = .Fields("dni")
            txtdni.Text = .Fields("dni")
            Text1(3).Text = .Fields("nom_emp")
            
            If .Fields("car_rod") = True Then
                chk_rod.Value = 1
            End If
            
            If .Fields("car_fer") = True Then
                chk_fer.Value = 1
            End If
            
            If .Fields("car_aer") = True Then
                chk_aer.Value = 1
            End If
            
            If .Fields("car_mar") = True Then
                chk_mar.Value = 1
            End If
            
        End With
End Sub

Private Sub restablecer()
    For i = 0 To 3
        cmd_mover(i).Enabled = True
    Next
    
    For i = 0 To 3
        Text1(i).Locked = True
    Next
    
    frabuscemp.Enabled = True
    frm_carnet.Enabled = False
    Text1(2).Enabled = True
    cmd_alta.Visible = True
    cmd_baja.Visible = True
    DataCombo1.Visible = False
    Text1(3).Visible = True
    cmd_cancelar.Visible = False
    cmd_confirmar2.Visible = False
    cmd_buscarsig.Visible = False
    cmd_mover_Click 0
    cmd_confirmaralta.Visible = False
    cmd_modif.Visible = True
End Sub


