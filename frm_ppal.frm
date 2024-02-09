VERSION 5.00
Begin VB.Form frm_ppal 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GV Logis 0.5"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11355
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   11355
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Listados informativos"
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
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6360
      Width           =   3855
   End
   Begin VB.Frame fra_3 
      BackColor       =   &H00E0E0E0&
      Height          =   3615
      Left            =   7680
      TabIndex        =   12
      Top             =   2280
      Width           =   3375
      Begin VB.CommandButton cmd_carga 
         BackColor       =   &H00808080&
         Caption         =   "Carga"
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2760
         Width           =   2415
      End
      Begin VB.CommandButton cmd_itinerarios 
         BackColor       =   &H00808080&
         Caption         =   "Itinerarios"
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2040
         Width           =   2415
      End
      Begin VB.CommandButton cmd_pilotar 
         BackColor       =   &H00808080&
         Caption         =   "Envios"
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CommandButton cmd_distri 
         BackColor       =   &H00808080&
         Caption         =   "Distribución"
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   600
         Width           =   2415
      End
   End
   Begin VB.Frame fra_2 
      BackColor       =   &H00E0E0E0&
      Height          =   3615
      Left            =   3960
      TabIndex        =   10
      Top             =   2280
      Width           =   3375
      Begin VB.CommandButton cmd_reparaciones 
         BackColor       =   &H00808080&
         Caption         =   "Reparaciones"
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1560
         Width           =   2415
      End
      Begin VB.CommandButton cmd_c_mante 
         BackColor       =   &H00808080&
         Caption         =   "Centros de mantenimiento"
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   600
         Width           =   2415
      End
   End
   Begin VB.Frame fra_1 
      BackColor       =   &H00E0E0E0&
      Height          =   3615
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   3375
      Begin VB.CommandButton cmd_m_trans 
         BackColor       =   &H00808080&
         Caption         =   "Medios de transporte"
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2520
         Width           =   2415
      End
      Begin VB.CommandButton cmd_empleados 
         BackColor       =   &H00808080&
         Caption         =   "Empleados"
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1560
         Width           =   2415
      End
      Begin VB.CommandButton cmd_empresa 
         BackColor       =   &H00808080&
         Caption         =   "Empresas"
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   600
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmd_consultasprov 
      BackColor       =   &H00808080&
      Caption         =   "Consultas Informativas"
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
      TabIndex        =   1
      Top             =   6360
      Width           =   3855
   End
   Begin VB.CommandButton cmd_salir 
      BackColor       =   &H00808080&
      Caption         =   "Salir"
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
      TabIndex        =   0
      Top             =   6360
      Width           =   2655
   End
   Begin VB.Image Image4 
      Height          =   735
      Left            =   240
      Picture         =   "frm_ppal.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   3615
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   3615
      Left            =   240
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Label lblgestser 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Gestión de servicios"
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
      Left            =   8160
      TabIndex        =   5
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label lblgestmant 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Gestión de mantenimiento"
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
      Left            =   4080
      TabIndex        =   4
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label lblgestemp 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Gestión de empresas"
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
      Left            =   720
      TabIndex        =   3
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   3615
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Line Line4 
      X1              =   11040
      X2              =   11040
      Y1              =   1920
      Y2              =   1200
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   240
      Y1              =   1920
      Y2              =   1200
   End
   Begin VB.Label lblgestion 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "GESTOR DE INFRAESTRUCTURAS EMPRESARIALES PARA EL TRANSPORTE"
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
      Left            =   2040
      TabIndex        =   2
      Top             =   360
      Width           =   8775
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   11040
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   11040
      Y1              =   1920
      Y2              =   1920
   End
End
Attribute VB_Name = "frm_ppal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_c_mante_Click()
    frm_Cmante.Show
    frm_ppal.Enabled = False
End Sub

Private Sub cmd_c_mante_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_c_mante.BackColor = &HFFFFFF
End Sub

Private Sub cmd_c_mante_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_c_mante.BackColor = &H808080
End Sub

Private Sub cmd_carga_Click()
    frm_carga.Show
    frm_ppal.Enabled = False
End Sub

Private Sub cmd_consultasprov_Click()
    frmlistconsul.Show
    frm_ppal.Enabled = False
End Sub

Private Sub cmd_consultasprov_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_consultasprov.BackColor = &HFFFFFF
End Sub

Private Sub cmd_consultasprov_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_consultasprov.BackColor = &H808080
End Sub

Private Sub cmd_distri_Click()
    frm_distribu.Show
    frm_ppal.Enabled = False
End Sub

Private Sub cmd_distri_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_distri.BackColor = &HFFFFFF
End Sub

Private Sub cmd_distri_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_distri.BackColor = &H808080
End Sub

Private Sub cmd_empleados_Click()
    frm_empleado.Show
    frm_ppal.Enabled = False
End Sub

Private Sub cmd_empleados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_empleados.BackColor = &HFFFFFF
End Sub

Private Sub cmd_empleados_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_empleados.BackColor = &H808080
End Sub

Private Sub cmd_empresa_Click()
    frm_empresa.Show
    frm_ppal.Enabled = False
End Sub

Private Sub cmd_empresa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_empresa.BackColor = &HFFFFFF
End Sub

Private Sub cmd_empresa_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_empresa.BackColor = &H808080
End Sub

Private Sub cmd_itinerarios_Click()
    frm_itinerario.Show
    frm_ppal.Enabled = False
End Sub

Private Sub cmd_itinerarios_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_itinerarios.BackColor = &HFFFFFF
End Sub

Private Sub cmd_itinerarios_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_itinerarios.BackColor = &H808080
End Sub

Private Sub cmd_m_trans_Click()
    frm_m_trans.Show
    frm_ppal.Enabled = False
End Sub

Private Sub cmd_m_trans_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_m_trans.BackColor = &HFFFFFF
End Sub

Private Sub cmd_m_trans_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_m_trans.BackColor = &H808080
End Sub

Private Sub cmd_pilotar_Click()
    frm_envios.Adodc4.RecordSource = "select reparto.*,carga.proced,carga.destino,carga.descripcion,m_trans.marca,empresa.nom_emp FROM reparto INNER JOIN (carga INNER JOIN (m_trans INNER JOIN empresa ON empresa.codigo = m_trans.cod_emp)ON m_trans.codigo = carga.cod_mtrans) ON reparto.cod_carga = carga.codigo WHERE reparto.asignado = 0 ORDER BY reparto.cod_carga ASC"
    frm_envios.Adodc4.Refresh
    frm_envios.Adodc3.RecordSource = "select cod_carga,cod_mtrans FROM reparto where asignado = 0 order by cod_carga asc"
    frm_envios.Adodc3.Refresh
    frm_envios.Show
    frm_ppal.Enabled = False
End Sub

Private Sub cmd_pilotar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_pilotar.BackColor = &HFFFFFF
End Sub

Private Sub cmd_pilotar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_pilotar.BackColor = &H808080
End Sub

Private Sub cmd_reparaciones_Click()
    frm_reparación.Show
    frm_ppal.Enabled = False
End Sub

Private Sub cmd_reparaciones_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_reparaciones.BackColor = &HFFFFFF
End Sub

Private Sub cmd_reparaciones_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_reparaciones.BackColor = &H808080
End Sub

Private Sub cmd_salir_Click()
    End
    cn.Close
End Sub

Private Sub cmd_salir_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_salir.BackColor = &HFFFFFF
End Sub

Private Sub cmd_salir_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_salir.BackColor = &H808080
End Sub

Private Sub Command1_Click()
    frm_list.Show
    frm_ppal.Enabled = False
End Sub

Private Sub Form_Load()
    Call controlesoriginal
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Visible = False
    Image2.Visible = True
    Image3.Visible = True
    
    fra_1.Visible = True
    fra_2.Visible = False
    fra_3.Visible = False
    
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.Visible = False
    Image1.Visible = True
    Image3.Visible = True
    
    fra_1.Visible = False
    fra_2.Visible = True
    fra_3.Visible = False
    
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image3.Visible = False
    Image1.Visible = True
    Image2.Visible = True
    
    fra_1.Visible = False
    fra_2.Visible = False
    fra_3.Visible = True
    
End Sub
