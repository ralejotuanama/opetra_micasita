VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_Caj_CiePag_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   9465
   ClientLeft      =   45
   ClientTop       =   1740
   ClientWidth     =   12900
   Icon            =   "OpeTra_frm_805.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   12900
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9465
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   13065
      _Version        =   65536
      _ExtentX        =   23045
      _ExtentY        =   16695
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   11
         Top             =   780
         Width           =   12810
         _Version        =   65536
         _ExtentX        =   22595
         _ExtentY        =   1138
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_805.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Limpiar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_805.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Buscar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   2400
            Picture         =   "OpeTra_frm_805.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_GasCie 
            Height          =   585
            Left            =   1215
            Picture         =   "OpeTra_frm_805.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Pago Masivo"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_GasAdm 
            Height          =   585
            Left            =   1800
            Picture         =   "OpeTra_frm_805.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Pago por Cliente"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   12210
            Picture         =   "OpeTra_frm_805.frx":14FE
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   12
         Top             =   60
         Width           =   12810
         _Version        =   65536
         _ExtentX        =   22595
         _ExtentY        =   1191
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin Threed.SSPanel SSPanel7 
            Height          =   255
            Left            =   630
            TabIndex        =   13
            Top             =   60
            Width           =   7875
            _Version        =   65536
            _ExtentX        =   13891
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Operaciones Financieras"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   255
            Left            =   630
            TabIndex        =   14
            Top             =   330
            Width           =   7875
            _Version        =   65536
            _ExtentX        =   13891
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Créditos Hipotecarios - Gastos de Cierre - Pago Proveedores"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_805.frx":1940
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   6675
         Left            =   30
         TabIndex        =   15
         Top             =   2100
         Width           =   12810
         _Version        =   65536
         _ExtentX        =   22595
         _ExtentY        =   11774
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin Threed.SSPanel pnl_Tit_Estado 
            Height          =   285
            Left            =   10650
            TabIndex        =   24
            Top             =   60
            Width           =   1810
            _Version        =   65536
            _ExtentX        =   3193
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Estado"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_Tit_NumSol 
            Height          =   285
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Solicitud"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_Tit_DocIde 
            Height          =   285
            Left            =   2610
            TabIndex        =   17
            Top             =   60
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "ID Cliente"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_Tit_NomCli 
            Height          =   285
            Left            =   3780
            TabIndex        =   18
            Top             =   60
            Width           =   4830
            _Version        =   65536
            _ExtentX        =   8520
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Apellidos y Nombres"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_Tit_Import 
            Height          =   285
            Left            =   9360
            TabIndex        =   19
            Top             =   60
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Saldo"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   6285
            Left            =   30
            TabIndex        =   9
            Top             =   360
            Width           =   12765
            _ExtentX        =   22516
            _ExtentY        =   11086
            _Version        =   393216
            Rows            =   30
            Cols            =   14
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_Moneda 
            Height          =   285
            Left            =   8580
            TabIndex        =   20
            Top             =   60
            Width           =   795
            _Version        =   65536
            _ExtentX        =   1411
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Moneda"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_Tit_FecSol 
            Height          =   285
            Left            =   1560
            TabIndex        =   27
            Top             =   60
            Width           =   1080
            _Version        =   65536
            _ExtentX        =   1905
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Solicitud"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   585
         Left            =   30
         TabIndex        =   21
         Top             =   8820
         Width           =   12810
         _Version        =   65536
         _ExtentX        =   22595
         _ExtentY        =   1032
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin VB.ComboBox cmb_Buscar 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   120
            Width           =   2595
         End
         Begin VB.TextBox txt_Buscar 
            Height          =   315
            Left            =   5730
            MaxLength       =   100
            TabIndex        =   8
            Top             =   150
            Width           =   4935
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Buscar Por:"
            Height          =   195
            Left            =   4740
            TabIndex        =   23
            Top             =   210
            Width           =   825
         End
         Begin VB.Label lbl_NomEti 
            AutoSize        =   -1  'True
            Caption         =   "Columna a Buscar:"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   22
            Top             =   210
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   585
         Left            =   30
         TabIndex        =   25
         Top             =   1470
         Width           =   12810
         _Version        =   65536
         _ExtentX        =   22595
         _ExtentY        =   1032
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin VB.ComboBox cmb_Situacion 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   150
            Width           =   3855
         End
         Begin VB.Label Label1 
            Caption         =   "Mostrar :"
            Height          =   285
            Left            =   150
            TabIndex        =   26
            Top             =   210
            Width           =   825
         End
      End
   End
End
Attribute VB_Name = "frm_Caj_CiePag_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_dbl_PorITF     As Double

Private Type arr_RegSol
   r_str_NumSol  As String
   r_int_DUPLIC  As Integer
   r_lng_FECCRE  As Long
   r_str_CodPrd  As String
   r_str_CodSub  As String
   r_int_TITTDO  As Integer
   r_lng_TITNDO  As Long
   r_lng_PAGFEC  As Long
   r_int_Situac  As Integer
   r_str_NomCli  As String
   r_dbl_PAGCLI  As Double
   r_dbl_PAGPRV  As Double
   r_dbl_SALDO   As Double
   r_str_Moneda  As String
   r_int_TipMon  As Integer
   r_str_Situac  As String
   r_lng_FecSol  As Long
End Type
   
Dim l_arr_RegDup()      As arr_RegSol

Private Sub cmd_ExpExc_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

'Private Type r_Arr_NumSol
'   str_NumSol         As String
'   dbl_Saldo          As Double
'End Type
'Dim arr_NumSol()       As r_Arr_NumSol

Private Sub cmd_GasCie_Click()
    frm_Caj_CiePag_03.Show 1
End Sub

Private Sub cmd_GasAdm_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   grd_Listad.Col = 2
   moddat_g_int_TipDoc = Left(grd_Listad.Text, 1)
   moddat_g_str_NumDoc = Mid(grd_Listad.Text, 3)
   
   grd_Listad.Col = 3
   moddat_g_str_NomCli = grd_Listad.Text
   
   grd_Listad.Col = 5
   moddat_g_dbl_IngDec = grd_Listad.Text
   
   grd_Listad.Col = 6
   moddat_g_str_FecIng = grd_Listad.Text
   
   grd_Listad.Col = 7
   moddat_g_str_NumSol = grd_Listad.Text
   
   grd_Listad.Col = 8
   moddat_g_int_TipMon = CInt(grd_Listad.Text)
   
   grd_Listad.Col = 9
   moddat_g_str_CodPrd = grd_Listad.Text
   
   grd_Listad.Col = 10
   moddat_g_str_CodSub = grd_Listad.Text
   
   grd_Listad.Col = 11
   moddat_g_int_Situac = grd_Listad.Text
   
   
   Call gs_RefrescaGrid(grd_Listad)
   moddat_g_int_FlgAct = 1
   Screen.MousePointer = 0
   
   frm_Caj_CiePag_02.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call cmd_Buscar_Click
   Call fs_Habilitado(True)
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   
   cmb_Situacion.Clear
   cmb_Situacion.AddItem "VIGENTES"
   cmb_Situacion.AddItem "CANCELADOS"
   cmb_Situacion.AddItem "<< TODOS >>"
   cmb_Situacion.ListIndex = 0
   
   cmb_Buscar.Clear
   cmb_Buscar.AddItem "NINGUNA"
   cmb_Buscar.AddItem "NRO SOLICITUD"
   cmb_Buscar.AddItem "FECHA SOLICITUD"
   cmb_Buscar.AddItem "ID CLIENTE"
   cmb_Buscar.AddItem "APELLIDOS Y NOMBRES"
   cmb_Buscar.AddItem "ESTADO"
   cmb_Buscar.ListIndex = 0

   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 1500 'GASADM_NUMSOL
   grd_Listad.ColWidth(1) = 1050 'SOLMAE_FECSOL
   grd_Listad.ColWidth(2) = 1190 'SOLMAE_TITNDO
   grd_Listad.ColWidth(3) = 4800 'NOM_CLIENTE
   grd_Listad.ColWidth(4) = 760 'MONEDA
   grd_Listad.ColWidth(5) = 1290 'SALDO
   grd_Listad.ColWidth(6) = 0    'GASADM_PAGFEC
   grd_Listad.ColWidth(7) = 0    'GASADM_NUMSOL
   grd_Listad.ColWidth(8) = 0    'SOLMAE_TIPMON
   grd_Listad.ColWidth(9) = 0    'SOLMAE_CODPRD
   grd_Listad.ColWidth(10) = 0   'SOLMAE_CODSUB
   grd_Listad.ColWidth(11) = 0   'SOLMAE_SITUAC
   grd_Listad.ColWidth(12) = 1800 'ESTADO
   grd_Listad.ColWidth(13) = 0 'SOLMAE_FECSOL
      
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_Listad.ColAlignment(12) = flexAlignCenterCenter
   grd_Listad.Rows = 0
End Sub

Public Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_Listad)
   cmb_Buscar.ListIndex = 0
   cmb_Situacion.ListIndex = 0
   cmb_Situacion.Enabled = True
End Sub

Private Sub cmd_Buscar_Click()
   Screen.MousePointer = 11
   Call fs_Buscar
   Screen.MousePointer = 0
   cmb_Situacion.Enabled = False
End Sub

Public Sub fs_Habilitado(ByVal p_Habilitado As Boolean)
   cmd_GasAdm.Enabled = p_Habilitado
   grd_Listad.Enabled = p_Habilitado
End Sub

Public Function fs_Query(p_Motrar As Integer) As String
   fs_Query = ""

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT GASADM_NUMSOL, SOLMAE_CODPRD, SOLMAE_CODSUB, SOLMAE_TITTDO, SOLMAE_TITNDO  ," 'GASADM_PAGFEC,  "
   g_str_Parame = g_str_Parame & "        (TRIM(E.DATGEN_APEPAT) || ' ' || TRIM(E.DATGEN_APEMAT) || ' ' ||  TRIM(E.DATGEN_NOMBRE)) AS NOM_CLIENTE,  "
   g_str_Parame = g_str_Parame & "        PAGO_CLIENTE, PAGO_PROVEEDOR, SALDO, TRIM(C.PARDES_DESCRI) AS MONEDA, SOLMAE_TIPMON,  "
   If p_Motrar = 1 Then 'VIGENTE
      g_str_Parame = g_str_Parame & "        DUP.DUPLICADO ,  "
   End If
   
   g_str_Parame = g_str_Parame & "        B.SEGFECCRE, SOLMAE_SITUAC, TRIM(F.PARDES_DESCRI) AS SITUACION, SOLMAE_FECSOL, "
   g_str_Parame = g_str_Parame & "        (SELECT MIN(GASADM_PAGFEC) "
   g_str_Parame = g_str_Parame & "           FROM TRA_GASADM E"
   g_str_Parame = g_str_Parame & "          WHERE E.GASADM_NUMSOL = A.GASADM_NUMSOL "
   g_str_Parame = g_str_Parame & "            AND GASADM_CODGAS <> 13 GROUP BY E.GASADM_NUMSOL) FECHA_PAGO "
   
   g_str_Parame = g_str_Parame & "   FROM (SELECT GASADM_NUMSOL, NVL(SUM(GASADM_PAGIMP),0) AS PAGO_CLIENTE, NVL(SUM(GASADM_MTOPAGPRV),0) PAGO_PROVEEDOR,  " 'GASADM_PAGFEC,
   g_str_Parame = g_str_Parame & "                NVL(NVL(Sum(GASADM_PAGIMP), 0) - NVL(Sum(GASADM_MTOPAGPRV), 0), 0) AS SALDO  "
   g_str_Parame = g_str_Parame & "           From TRA_GASADM  "
   'g_str_Parame = g_str_Parame & "          WHERE SUBSTR(GASADM_NUMSOL,1,3) NOT IN ('001','003')"
   g_str_Parame = g_str_Parame & "          GROUP BY GASADM_NUMSOL) A  " ', GASADM_PAGFEC
   g_str_Parame = g_str_Parame & "          INNER JOIN CRE_SOLMAE B ON B.SOLMAE_NUMERO = A.GASADM_NUMSOL AND SOLMAE_SITUAC IN (1,2,3) "
   g_str_Parame = g_str_Parame & "          INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 229 AND C.PARDES_CODITE = B.SOLMAE_TIPMON  "
   g_str_Parame = g_str_Parame & "          INNER JOIN CLI_DATGEN E ON E.DATGEN_TIPDOC = B.SOLMAE_TITTDO AND E.DATGEN_NUMDOC = B.SOLMAE_TITNDO  "
   g_str_Parame = g_str_Parame & "          INNER JOIN MNT_PARDES F ON F.PARDES_CODGRP = 020 AND F.PARDES_CODITE = SOLMAE_SITUAC  "
   If p_Motrar = 1 Then 'VIGENTE
      g_str_Parame = g_str_Parame & "          INNER JOIN (SELECT DATGEN_NUMDOC, COUNT(EE.DATGEN_NUMDOC) DUPLICADO  "
      g_str_Parame = g_str_Parame & "                        FROM (SELECT GASADM_NUMSOL, NVL(NVL(SUM(GASADM_PAGIMP), 0) - NVL(Sum(GASADM_MTOPAGPRV), 0), 0) AS SALDO  "
      g_str_Parame = g_str_Parame & "                                From TRA_GASADM  "
      'g_str_Parame = g_str_Parame & "                               WHERE SUBSTR(GASADM_NUMSOL,1,3) NOT IN ('001','003')  "
      g_str_Parame = g_str_Parame & "                               GROUP BY GASADM_NUMSOL) AA  " ', GASADM_PAGFEC
      g_str_Parame = g_str_Parame & "                       INNER JOIN CRE_SOLMAE BB ON BB.SOLMAE_NUMERO = AA.GASADM_NUMSOL AND SOLMAE_SITUAC IN (1,2,3)  "
      g_str_Parame = g_str_Parame & "                       INNER JOIN CLI_DATGEN EE ON EE.DATGEN_TIPDOC = BB.SOLMAE_TITTDO AND EE.DATGEN_NUMDOC = BB.SOLMAE_TITNDO  "
      g_str_Parame = g_str_Parame & "                       Where SALDO > 0  "
      g_str_Parame = g_str_Parame & "                       GROUP BY EE.DATGEN_NUMDOC) DUP ON  DUP.DATGEN_NUMDOC = E.DATGEN_NUMDOC  "
   End If
   
   If p_Motrar = 1 Then 'VIGENTE
      g_str_Parame = g_str_Parame & "  WHERE SALDO > 0  "
   Else 'CANCELADO
      g_str_Parame = g_str_Parame & "  WHERE SALDO = 0  "
   End If
   
   If cmb_Buscar.ListIndex > 0 Then
      If cmb_Buscar.ListIndex = 1 Then 'NRO SOLICITUD
         g_str_Parame = g_str_Parame & "   AND SUBSTR(GASADM_NUMSOL,1,3)||'-'||SUBSTR(GASADM_NUMSOL,4,3)||'-'||SUBSTR(GASADM_NUMSOL,7,2)||'-'||SUBSTR(GASADM_NUMSOL,9,4) = '" & Trim(txt_Buscar.Text) & "' "
      ElseIf cmb_Buscar.ListIndex = 2 Then 'FECHA SOLICITUD
         g_str_Parame = g_str_Parame & "   AND SUBSTR(SOLMAE_FECSOL,7,2)||'/'||SUBSTR(SOLMAE_FECSOL,5,2)||'/'||SUBSTR(SOLMAE_FECSOL,1,4) LIKE '%" & UCase(Trim(txt_Buscar.Text)) & "%' "
      ElseIf cmb_Buscar.ListIndex = 3 Then 'ID-CLIENTE
         g_str_Parame = g_str_Parame & "   AND TRIM(SOLMAE_TITNDO) = '" & Trim(txt_Buscar.Text) & "' "
      ElseIf cmb_Buscar.ListIndex = 4 Then 'APELLIDOS Y  NOMBRES
         g_str_Parame = g_str_Parame & "   AND TRIM(E.DATGEN_APEPAT)|| ' ' ||TRIM(E.DATGEN_APEMAT)|| ' ' ||TRIM(E.DATGEN_NOMBRE) like '%" & UCase(Trim(txt_Buscar.Text)) & "%' "
      ElseIf cmb_Buscar.ListIndex = 5 Then 'ESTADO
         g_str_Parame = g_str_Parame & "   AND TRIM(F.PARDES_DESCRI) LIKE '%" & UCase(Trim(txt_Buscar.Text)) & "%'  "
      End If
   End If
   If p_Motrar = 1 Then 'VIGENTE
      g_str_Parame = g_str_Parame & "  ORDER BY DUPLICADO, GASADM_NUMSOL, SOLMAE_TITNDO  "
   Else
      g_str_Parame = g_str_Parame & "  ORDER BY GASADM_NUMSOL, SOLMAE_TITNDO  "
   End If
   fs_Query = g_str_Parame
End Function

Public Sub fs_Buscar()
Dim r_dbl_ITFGas     As Double
Dim r_rst_NumPag     As ADODB.Recordset
Dim r_int_DifDia     As Integer

   Call gs_LimpiaGrid(grd_Listad)
   ReDim l_arr_RegDup(0)
   
   '=====================VIGENTES=============================
   If cmb_Situacion.ListIndex = 0 Or cmb_Situacion.ListIndex = 2 Then
      g_str_Parame = fs_Query(1)
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
   
      'No existen registros
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         Exit Sub
      End If
      
      grd_Listad.Redraw = False
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!DUPLICADO = 1 Then
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            
            grd_Listad.Col = 0
            grd_Listad.Text = gf_Formato_NumSol(g_rst_Princi!GASADM_NUMSOL)
            
            grd_Listad.Col = 1
            grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!SOLMAE_FECSOL)
            
            grd_Listad.Col = 2
            grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
            
            grd_Listad.Col = 3
            grd_Listad.Text = Trim(g_rst_Princi!NOM_CLIENTE)
            
            grd_Listad.Col = 4
            grd_Listad.Text = Trim(g_rst_Princi!MONEDA)
                    
            grd_Listad.Col = 5
            grd_Listad.Text = Format(g_rst_Princi!SALDO, "###,###,##0.00")
            
            grd_Listad.Col = 6
            grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!FECHA_PAGO) 'GASADM_PAGFEC
            
            grd_Listad.Col = 7
            grd_Listad.Text = Trim(g_rst_Princi!GASADM_NUMSOL & "")
            
            grd_Listad.Col = 8
            grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_TIPMON)
            
            grd_Listad.Col = 9
            grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CODPRD & "")
               
            grd_Listad.Col = 10
            grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CODSUB & "")
            
            grd_Listad.Col = 11
            grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_SITUAC & "")
            
            grd_Listad.Col = 12
            grd_Listad.Text = Trim(g_rst_Princi!SITUACION & "")
            
            grd_Listad.Col = 13
            grd_Listad.Text = g_rst_Princi!SOLMAE_FECSOL
            
           'siguiente registro
            g_rst_Princi.MoveNext
         Else
            '***AGREGAR AL ARREGLO
            ReDim Preserve l_arr_RegDup(UBound(l_arr_RegDup) + 1)
            l_arr_RegDup(UBound(l_arr_RegDup)).r_str_NumSol = Trim(g_rst_Princi!GASADM_NUMSOL & "")
            l_arr_RegDup(UBound(l_arr_RegDup)).r_int_DUPLIC = g_rst_Princi!DUPLICADO
            l_arr_RegDup(UBound(l_arr_RegDup)).r_lng_FECCRE = g_rst_Princi!SEGFECCRE
            l_arr_RegDup(UBound(l_arr_RegDup)).r_str_CodPrd = Trim(g_rst_Princi!SOLMAE_CODPRD & "")
            l_arr_RegDup(UBound(l_arr_RegDup)).r_str_CodSub = Trim(g_rst_Princi!SOLMAE_CODSUB & "")
            l_arr_RegDup(UBound(l_arr_RegDup)).r_int_TITTDO = g_rst_Princi!SOLMAE_TITTDO
            l_arr_RegDup(UBound(l_arr_RegDup)).r_lng_TITNDO = g_rst_Princi!SOLMAE_TITNDO
            l_arr_RegDup(UBound(l_arr_RegDup)).r_lng_PAGFEC = g_rst_Princi!FECHA_PAGO 'GASADM_PAGFEC
            l_arr_RegDup(UBound(l_arr_RegDup)).r_int_Situac = g_rst_Princi!SOLMAE_SITUAC
            l_arr_RegDup(UBound(l_arr_RegDup)).r_str_NomCli = Trim(g_rst_Princi!NOM_CLIENTE & "")
            l_arr_RegDup(UBound(l_arr_RegDup)).r_dbl_PAGCLI = g_rst_Princi!PAGO_CLIENTE
            l_arr_RegDup(UBound(l_arr_RegDup)).r_dbl_PAGPRV = g_rst_Princi!PAGO_PROVEEDOR
            l_arr_RegDup(UBound(l_arr_RegDup)).r_dbl_SALDO = g_rst_Princi!SALDO
            l_arr_RegDup(UBound(l_arr_RegDup)).r_str_Moneda = Trim(g_rst_Princi!MONEDA & "")
            l_arr_RegDup(UBound(l_arr_RegDup)).r_int_TipMon = g_rst_Princi!SOLMAE_TIPMON
            l_arr_RegDup(UBound(l_arr_RegDup)).r_str_Situac = g_rst_Princi!SITUACION
            l_arr_RegDup(UBound(l_arr_RegDup)).r_lng_FecSol = g_rst_Princi!SOLMAE_FECSOL
            'siguiente registro
            g_rst_Princi.MoveNext
         End If
      Loop
      '!!!COLUMNA ADICIONA GRILLA - SE DEBE ADICIONAR A LA DUPLICIDAD!!!
      Call fs_GrdDuplicado
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   '=====================CANCELADOS=============================
   If cmb_Situacion.ListIndex = 1 Or cmb_Situacion.ListIndex = 2 Then
      g_str_Parame = fs_Query(2)
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      'No existen registros
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         Exit Sub
      End If
      
      grd_Listad.Redraw = False
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            
            grd_Listad.Col = 0
            grd_Listad.Text = gf_Formato_NumSol(g_rst_Princi!GASADM_NUMSOL)
            
            grd_Listad.Col = 1
            grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!SOLMAE_FECSOL)
            
            grd_Listad.Col = 2
            grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
            
            grd_Listad.Col = 3
            grd_Listad.Text = Trim(g_rst_Princi!NOM_CLIENTE)
            
            grd_Listad.Col = 4
            grd_Listad.Text = Trim(g_rst_Princi!MONEDA)
                    
            grd_Listad.Col = 5
            grd_Listad.Text = Format(g_rst_Princi!SALDO, "###,###,##0.00")
            
            grd_Listad.Col = 6
            grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!FECHA_PAGO) 'g_rst_Princi!GASADM_PAGFEC
            
            grd_Listad.Col = 7
            grd_Listad.Text = Trim(g_rst_Princi!GASADM_NUMSOL & "")
            
            grd_Listad.Col = 8
            grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_TIPMON)
            
            grd_Listad.Col = 9
            grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CODPRD & "")
               
            grd_Listad.Col = 10
            grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CODSUB & "")
            
            grd_Listad.Col = 11
            grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_SITUAC & "")
            
            grd_Listad.Col = 12
            grd_Listad.Text = Trim(g_rst_Princi!SITUACION & "")
            
            grd_Listad.Col = 13
            grd_Listad.Text = g_rst_Princi!SOLMAE_FECSOL
            
            'siguiente registro
            g_rst_Princi.MoveNext
      Loop
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   If grd_Listad.Rows = 0 Then
      MsgBox "No se encontraron solicitudes para la asignación del pago a proveedor.", vbInformation, modgen_g_str_NomPlt
   End If
   
   grd_Listad.Redraw = True
   'Ordenando por Nombres de Clientes
   pnl_Tit_NomCli.Tag = "A"
   Call gs_SorteaGrid(grd_Listad, 2, "C")
   Call gs_UbiIniGrid(grd_Listad)
End Sub



Private Sub fs_GrdDuplicado()
   Dim r_int_ConfiA     As Integer
   Dim r_int_ConfiB     As Integer
   Dim r_int_ConfiC     As Integer
   
   Dim r_str_NumSol     As String:  Dim r_int_DUPLIC As Integer
   Dim r_lng_FECCRE     As Long:    Dim r_str_CodPrd As String
   Dim r_str_CodSub     As String:  Dim r_int_TITTDO As Integer
   Dim r_lng_TITNDO     As Long:    Dim r_lng_PAGFEC As Long
   Dim r_int_Situac     As Integer: Dim r_str_NomCli As String
   Dim r_dbl_PAGCLI     As Double:  Dim r_dbl_PAGPRV As Double
   Dim r_dbl_SALDO      As Double:  Dim r_str_Moneda As String
   Dim r_int_TipMon     As Integer: Dim r_str_Situac As String
   Dim r_lng_FecSol     As Long
   
   For r_int_ConfiA = 1 To UBound(l_arr_RegDup)
       If l_arr_RegDup(r_int_ConfiA).r_int_DUPLIC <> 1 Then
          r_str_NumSol = l_arr_RegDup(r_int_ConfiA).r_str_NumSol
          r_int_DUPLIC = l_arr_RegDup(r_int_ConfiA).r_int_DUPLIC
          r_lng_FECCRE = l_arr_RegDup(r_int_ConfiA).r_lng_FECCRE
          r_str_CodPrd = l_arr_RegDup(r_int_ConfiA).r_str_CodPrd
          r_str_CodSub = l_arr_RegDup(r_int_ConfiA).r_str_CodSub
          r_int_TITTDO = l_arr_RegDup(r_int_ConfiA).r_int_TITTDO
          r_lng_TITNDO = l_arr_RegDup(r_int_ConfiA).r_lng_TITNDO
          r_lng_PAGFEC = l_arr_RegDup(r_int_ConfiA).r_lng_PAGFEC
          r_int_Situac = l_arr_RegDup(r_int_ConfiA).r_int_Situac
          r_str_NomCli = l_arr_RegDup(r_int_ConfiA).r_str_NomCli
          r_dbl_PAGCLI = l_arr_RegDup(r_int_ConfiA).r_dbl_PAGCLI
          r_dbl_PAGPRV = l_arr_RegDup(r_int_ConfiA).r_dbl_PAGPRV
          r_dbl_SALDO = l_arr_RegDup(r_int_ConfiA).r_dbl_SALDO
          r_str_Moneda = l_arr_RegDup(r_int_ConfiA).r_str_Moneda
          r_int_TipMon = l_arr_RegDup(r_int_ConfiA).r_int_TipMon
          r_str_Situac = l_arr_RegDup(r_int_ConfiA).r_str_Situac
          r_lng_FecSol = l_arr_RegDup(r_int_ConfiA).r_lng_FecSol
         
          For r_int_ConfiB = 1 To UBound(l_arr_RegDup)
              If l_arr_RegDup(r_int_ConfiA).r_lng_TITNDO = l_arr_RegDup(r_int_ConfiB).r_lng_TITNDO Then
                 If l_arr_RegDup(r_int_ConfiA).r_lng_FECCRE < l_arr_RegDup(r_int_ConfiB).r_lng_FECCRE Then
                    r_str_NumSol = l_arr_RegDup(r_int_ConfiB).r_str_NumSol
                    r_int_DUPLIC = l_arr_RegDup(r_int_ConfiB).r_int_DUPLIC
                    r_lng_FECCRE = l_arr_RegDup(r_int_ConfiB).r_lng_FECCRE
                    r_str_CodPrd = l_arr_RegDup(r_int_ConfiB).r_str_CodPrd
                    r_str_CodSub = l_arr_RegDup(r_int_ConfiB).r_str_CodSub
                    r_int_TITTDO = l_arr_RegDup(r_int_ConfiB).r_int_TITTDO
                    r_lng_TITNDO = l_arr_RegDup(r_int_ConfiB).r_lng_TITNDO
                    r_lng_PAGFEC = l_arr_RegDup(r_int_ConfiB).r_lng_PAGFEC
                    r_int_Situac = l_arr_RegDup(r_int_ConfiB).r_int_Situac
                    r_str_NomCli = l_arr_RegDup(r_int_ConfiB).r_str_NomCli
                    r_dbl_PAGCLI = l_arr_RegDup(r_int_ConfiB).r_dbl_PAGCLI
                    r_dbl_PAGPRV = l_arr_RegDup(r_int_ConfiB).r_dbl_PAGPRV
                    r_dbl_SALDO = l_arr_RegDup(r_int_ConfiB).r_dbl_SALDO
                    r_str_Moneda = l_arr_RegDup(r_int_ConfiB).r_str_Moneda
                    r_int_TipMon = l_arr_RegDup(r_int_ConfiB).r_int_TipMon
                    r_str_Situac = l_arr_RegDup(r_int_ConfiB).r_str_Situac
                    r_lng_FecSol = l_arr_RegDup(r_int_ConfiB).r_lng_FecSol
                 End If
              End If
          Next
          'update no duplicado
          For r_int_ConfiC = 1 To UBound(l_arr_RegDup)
              If l_arr_RegDup(r_int_ConfiC).r_lng_TITNDO = r_lng_TITNDO Then
                 l_arr_RegDup(r_int_ConfiC).r_int_DUPLIC = 1
              End If
          Next
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0
         grd_Listad.Text = gf_Formato_NumSol(r_str_NumSol)
         grd_Listad.Col = 1
         grd_Listad.Text = gf_FormatoFecha(r_lng_FecSol)
         grd_Listad.Col = 2
         grd_Listad.Text = CStr(r_int_TITTDO) & "-" & Trim(r_lng_TITNDO)
         grd_Listad.Col = 3
         grd_Listad.Text = Trim(r_str_NomCli)
         grd_Listad.Col = 4
         grd_Listad.Text = Trim(r_str_Moneda)
         grd_Listad.Col = 5
         grd_Listad.Text = Format(r_dbl_SALDO, "###,###,##0.00")
         grd_Listad.Col = 6
         grd_Listad.Text = gf_FormatoFecha(r_lng_PAGFEC)
         grd_Listad.Col = 7
         grd_Listad.Text = Trim(r_str_NumSol & "")
         grd_Listad.Col = 8
         grd_Listad.Text = CStr(r_int_TipMon)
         grd_Listad.Col = 9
         grd_Listad.Text = Trim(r_str_CodPrd & "")
         grd_Listad.Col = 10
         grd_Listad.Text = Trim(r_str_CodSub & "")
         grd_Listad.Col = 11
         grd_Listad.Text = Trim(r_int_Situac & "")
         grd_Listad.Col = 12
         grd_Listad.Text = Trim(r_str_Situac & "")
       End If
   Next
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_int_NumFil        As Integer
Dim r_int_Contar        As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "REPORTE DE GASTOS DE CIERRE - PAGO PROVEEDORES (" & UCase(Trim(cmb_Situacion.Text)) & ") "
      .Range(.Cells(2, 2), .Cells(2, 8)).Merge
      .Range(.Cells(2, 2), .Cells(2, 8)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 8)).HorizontalAlignment = xlHAlignCenter

      .Cells(4, 2) = "NRO SOLICITUD"
      .Cells(4, 3) = "FECHA SOLICITUD"
      .Cells(4, 4) = "ID CLIENTE"
      .Cells(4, 5) = "APELLIDOS Y NOMBRES"
      .Cells(4, 6) = "MONEDA"
      .Cells(4, 7) = "SALDO"
      .Cells(4, 8) = "ESTADO"
      .Cells(1, 8) = "'" & Format(CDate(date), "dd/mm/yyyy")
       
      .Range(.Cells(4, 2), .Cells(4, 8)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(4, 2), .Cells(4, 8)).Font.Bold = True
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 17 'nro solicitud
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 16 'fecha solictud
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 15 'id-cliente
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 55 'apellidos nombre
      .Columns("E").HorizontalAlignment = xlHAlignLeft
      .Columns("F").ColumnWidth = 9 'moneda
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 18 'saldo
      .Columns("G").NumberFormat = "###,###,##0.00"
      .Columns("G").HorizontalAlignment = xlHAlignRight
      .Columns("H").ColumnWidth = 22 'estado
      .Columns("H").HorizontalAlignment = xlHAlignCenter
            
      .Range(.Cells(1, 1), .Cells(10, 8)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(10, 8)).Font.Size = 11
            
      r_int_NumFil = 5
      For r_int_Contar = 0 To grd_Listad.Rows - 1
         .Cells(r_int_NumFil, 2) = "'" & grd_Listad.TextMatrix(r_int_Contar, 0)
         .Cells(r_int_NumFil, 3) = "'" & grd_Listad.TextMatrix(r_int_Contar, 1)
         .Cells(r_int_NumFil, 4) = grd_Listad.TextMatrix(r_int_Contar, 2)
         .Cells(r_int_NumFil, 5) = grd_Listad.TextMatrix(r_int_Contar, 3)
         .Cells(r_int_NumFil, 6) = grd_Listad.TextMatrix(r_int_Contar, 4)
         .Cells(r_int_NumFil, 7) = grd_Listad.TextMatrix(r_int_Contar, 5)
         .Cells(r_int_NumFil, 8) = grd_Listad.TextMatrix(r_int_Contar, 12)
         
         r_int_NumFil = r_int_NumFil + 1
      Next
      .Range(.Cells(4, 4), .Cells(4, 8)).HorizontalAlignment = xlHAlignCenter
      .Cells(1, 8).HorizontalAlignment = xlHAlignRight
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_GasAdm_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub pnl_Tit_NumSol_Click()
   If Trim(pnl_Tit_NumSol.Tag) = "" Then
      pnl_Tit_NumSol.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Tit_NumSol.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub

Private Sub pnl_Tit_DocIde_Click()
   If Trim(pnl_Tit_DocIde.Tag) = "" Then
      pnl_Tit_DocIde.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_Tit_DocIde.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub

Private Sub pnl_Tit_NomCli_Click()
   If Trim(pnl_Tit_NomCli.Tag) = "" Then
      pnl_Tit_NomCli.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 3, "C")
   Else
      pnl_Tit_NomCli.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 3, "C-")
   End If
End Sub

Private Sub pnl_Tit_Moneda_Click()
   If Trim(pnl_Tit_Moneda.Tag) = "" Then
      pnl_Tit_Moneda.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 4, "C")
   Else
      pnl_Tit_Moneda.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 4, "C-")
   End If
End Sub

Private Sub pnl_Tit_Import_Click()
   If Trim(pnl_Tit_Import.Tag) = "" Then
      pnl_Tit_Import.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 5, "N")
   Else
      pnl_Tit_Import.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 5, "N-")
   End If
End Sub

Private Sub pnl_Tit_Estado_Click()
   If Trim(pnl_Tit_Estado.Tag) = "" Then
      pnl_Tit_Estado.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 12, "C")
   Else
      pnl_Tit_Estado.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 12, "C-")
   End If
End Sub

Private Sub pnl_Tit_FecSol_Click()
   If Trim(pnl_Tit_FecSol.Tag) = "" Then
      pnl_Tit_FecSol.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 13, "N")
   Else
      pnl_Tit_FecSol.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 13, "N-")
   End If
End Sub

Private Sub cmb_Buscar_Click()
    If (cmb_Buscar.ListIndex = 0 Or cmb_Buscar.ListIndex = -1) Then
        txt_Buscar.Enabled = False
        Call gs_SetFocus(cmd_Buscar)
    Else
        txt_Buscar.Enabled = True
    End If
    txt_Buscar.Text = ""
End Sub

Private Sub cmb_Buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If (txt_Buscar.Enabled = False) Then
          Call gs_SetFocus(cmd_Buscar)
      Else
          Call gs_SetFocus(txt_Buscar)
      End If
   End If
End Sub

Private Sub txt_Buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   Else
      If cmb_Buscar.ListIndex = 3 Then
         KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
      Else
         KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
      End If
   End If
End Sub

'Private Function fs_Busca_Saldo(ByVal p_NumSol As String) As String
'Dim r_int_Contad     As Integer
'
'   fs_Busca_Saldo = ""
'   For r_int_Contad = 1 To UBound(arr_NumSol)
'      If Trim(arr_NumSol(r_int_Contad).str_NumSol) = Trim(p_NumSol) Then
'         fs_Busca_Saldo = Trim(arr_NumSol(r_int_Contad).dbl_Saldo)
'         Exit For
'      End If
'   Next r_int_Contad
'End Function
