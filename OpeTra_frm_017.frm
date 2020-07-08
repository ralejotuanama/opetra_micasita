VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_GasAdm_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   8790
   ClientLeft      =   1080
   ClientTop       =   1920
   ClientWidth     =   17475
   Icon            =   "OpeTra_frm_017.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   17475
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8775
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   17490
      _Version        =   65536
      _ExtentX        =   30850
      _ExtentY        =   15487
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
         Left            =   60
         TabIndex        =   4
         Top             =   750
         Width           =   17385
         _Version        =   65536
         _ExtentX        =   30665
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
         Begin VB.CommandButton cmd_Export 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_017.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Exportar datos a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_GasAdm 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_017.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Evaluar Solicitud"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   16770
            Picture         =   "OpeTra_frm_017.frx":0BE0
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salida"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   5
         Top             =   30
         Width           =   17385
         _Version        =   65536
         _ExtentX        =   30665
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
            Height          =   585
            Left            =   660
            TabIndex        =   12
            Top             =   30
            Width           =   3975
            _Version        =   65536
            _ExtentX        =   7011
            _ExtentY        =   1032
            _StockProps     =   15
            Caption         =   "Asignación de Gastos de Cierre"
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
            Picture         =   "OpeTra_frm_017.frx":1022
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   6525
         Left            =   60
         TabIndex        =   6
         Top             =   1440
         Width           =   17385
         _Version        =   65536
         _ExtentX        =   30665
         _ExtentY        =   11509
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
         Begin Threed.SSPanel pnl_Tit_Produc 
            Height          =   285
            Left            =   60
            TabIndex        =   7
            Top             =   60
            Width           =   3135
            _Version        =   65536
            _ExtentX        =   5530
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Producto"
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
            Left            =   8220
            TabIndex        =   8
            Top             =   60
            Width           =   1500
            _Version        =   65536
            _ExtentX        =   2646
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
            Left            =   9600
            TabIndex        =   9
            Top             =   60
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
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
            Left            =   11040
            TabIndex        =   10
            Top             =   60
            Width           =   3495
            _Version        =   65536
            _ExtentX        =   6165
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
         Begin Threed.SSPanel pnl_Tit_FecSol 
            Height          =   285
            Left            =   14520
            TabIndex        =   11
            Top             =   60
            Width           =   1260
            _Version        =   65536
            _ExtentX        =   2222
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Solicitud"
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
            Height          =   6120
            Left            =   30
            TabIndex        =   0
            Top             =   360
            Width           =   17295
            _ExtentX        =   30506
            _ExtentY        =   10795
            _Version        =   393216
            Rows            =   30
            Cols            =   14
            FixedRows       =   0
            FixedCols       =   0
            ForeColor       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_FecEva 
            Height          =   285
            Left            =   15750
            TabIndex        =   13
            Top             =   60
            Width           =   1230
            _Version        =   65536
            _ExtentX        =   2170
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F.Ing. Instancia"
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
         Begin Threed.SSPanel pnl_Tit_Proyec 
            Height          =   285
            Left            =   3120
            TabIndex        =   24
            Top             =   60
            Width           =   4485
            _Version        =   65536
            _ExtentX        =   7911
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Proyecto"
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
         Begin Threed.SSPanel pnl_Tit_Tipo 
            Height          =   285
            Left            =   7605
            TabIndex        =   25
            Top             =   60
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1085
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo"
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
      Begin Threed.SSPanel SSPanel4 
         Height          =   690
         Left            =   60
         TabIndex        =   14
         Top             =   8010
         Width           =   17355
         _Version        =   65536
         _ExtentX        =   30612
         _ExtentY        =   1217
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
         Begin Threed.SSPanel pnl_OpeObs 
            Height          =   315
            Left            =   11790
            TabIndex        =   17
            Top             =   210
            Width           =   945
            _Version        =   65536
            _ExtentX        =   1667
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_OpeNue 
            Height          =   315
            Left            =   6720
            TabIndex        =   18
            Top             =   210
            Width           =   945
            _Version        =   65536
            _ExtentX        =   1667
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_OpeAsi 
            Height          =   315
            Left            =   2400
            TabIndex        =   19
            Top             =   210
            Width           =   945
            _Version        =   65536
            _ExtentX        =   1667
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_TotReg 
            Height          =   315
            Left            =   16140
            TabIndex        =   21
            Top             =   210
            Width           =   945
            _Version        =   65536
            _ExtentX        =   1676
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel SSPanel5 
            Height          =   315
            Left            =   4710
            TabIndex        =   15
            Top             =   210
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "OPERACIONES NUEVAS"
            ForeColor       =   255
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
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel SSPanel9 
            Height          =   315
            Left            =   14220
            TabIndex        =   22
            Top             =   210
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "TOTAL DE REGISTROS"
            ForeColor       =   0
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
            FloodColor      =   0
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel SSPanel11 
            Height          =   315
            Left            =   150
            TabIndex        =   20
            Top             =   210
            Width           =   2235
            _Version        =   65536
            _ExtentX        =   3942
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "OPERACIONES ASIGNADAS"
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
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   315
            Left            =   9330
            TabIndex        =   16
            Top             =   210
            Width           =   2445
            _Version        =   65536
            _ExtentX        =   4313
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "OPERACIONES OBSERVADAS"
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   1
         End
      End
   End
End
Attribute VB_Name = "frm_GasAdm_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Export_Click()
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_GasAdm_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 0
   moddat_g_str_NomPrd = grd_Listad.Text

   grd_Listad.Col = 3
   moddat_g_str_NumSol = Mid(grd_Listad.Text, 1, 3) & Mid(grd_Listad.Text, 5, 3) & Mid(grd_Listad.Text, 9, 2) & Mid(grd_Listad.Text, 12, 4)
   
   grd_Listad.Col = 4
   moddat_g_int_TipDoc = Left(grd_Listad.Text, 1)
   moddat_g_str_NumDoc = Mid(grd_Listad.Text, 3)
         
   grd_Listad.Col = 5
   moddat_g_str_NomCli = grd_Listad.Text
   
   grd_Listad.Col = 6
   moddat_g_str_FecIng = grd_Listad.Text
   
   grd_Listad.Col = 7
   moddat_g_str_CodPrd = grd_Listad.Text
   
   grd_Listad.Col = 8
   moddat_g_str_CodSub = grd_Listad.Text
   
   grd_Listad.Col = 10
   moddat_g_int_TipMon = CInt(grd_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Listad)
   
   If moddat_g_int_TipMon <> 1 Then
      If moddat_gf_Obtiene_TipCam(1, 2) = 0 Then
         MsgBox "Debe registrar el Tipo de Cambio para la moneda " & moddat_gf_Consulta_ParDes("204", CStr(2)) & ".", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   End If
   
   moddat_g_int_FlgAct = 1
   frm_GasAdm_11.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Buscar
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 3130                 'PRODUCTO
   grd_Listad.ColWidth(1) = 4410                 'PROYECTO
   grd_Listad.ColWidth(2) = 630                  'TIPO
   grd_Listad.ColWidth(3) = 1480                 'SOLICITUD
   grd_Listad.ColWidth(4) = 1320                 'CLIENTE
   grd_Listad.ColWidth(5) = 3510                 'APELLIDOS Y NOMBRES
   grd_Listad.ColWidth(6) = 1230                 'FECHA DE SOLICITUD
   grd_Listad.ColWidth(7) = 0
   grd_Listad.ColWidth(8) = 0
   grd_Listad.ColWidth(9) = 0
   grd_Listad.ColWidth(10) = 0
   grd_Listad.ColWidth(11) = 1200                'FECHA DE INGRESO A INSTANCIA
   grd_Listad.ColWidth(12) = 0
   grd_Listad.ColWidth(13) = 0
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter
   grd_Listad.ColAlignment(6) = flexAlignCenterCenter
   grd_Listad.ColAlignment(11) = flexAlignCenterCenter
End Sub

Private Sub fs_Buscar()
Dim r_rst_Genera     As ADODB.Recordset
Dim r_var_ColCel
Dim r_int_FlgAsg     As Integer
Dim r_int_OpeAsi     As Integer
Dim r_int_OpeNue     As Integer
Dim r_int_OpeObs     As Integer
Dim r_int_TotReg     As Integer
   
   r_int_OpeAsi = 0
   r_int_OpeNue = 0
   r_int_OpeObs = 0
   r_int_TotReg = 0
   pnl_OpeAsi.Caption = ""
   pnl_OpeNue.Caption = ""
   pnl_OpeObs.Caption = ""
   pnl_TotReg.Caption = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT SOLMAE_CODPRD, SOLMAE_NUMERO, SOLMAE_TITTDO, SOLMAE_TITNDO, F.DATGEN_TITULO, SOLMAE_FECSOL, SOLMAE_CODSUB, SOLMAE_TIPMON, "
   g_str_Parame = g_str_Parame & "        TRIM(D.DATGEN_APEPAT)||' '||TRIM(D.DATGEN_APEMAT)||' '||TRIM(D.DATGEN_NOMBRE) AS NOM_CLIENTE, "
   g_str_Parame = g_str_Parame & "        (SELECT COUNT(*) FROM TRA_GASADM C WHERE C.GASADM_NUMSOL = A.SOLMAE_NUMERO) AS TOTGSTCIE, "
   g_str_Parame = g_str_Parame & "        (SELECT TRIM(E.PRODUC_DESCRI) FROM CRE_PRODUC E WHERE E.PRODUC_CODIGO = LPAD(SOLMAE_CODPRD,3,'0')) NOM_PRODUCTO, "
   g_str_Parame = g_str_Parame & "        (SELECT E.SEGUIM_FECINI FROM TRA_SEGUIM E WHERE E.SEGUIM_NUMSOL = A.SOLMAE_NUMERO AND E.SEGUIM_CODINS = 31) AS FEC_INGRESO, "
   g_str_Parame = g_str_Parame & "        (SELECT COUNT(*) FROM TRA_SEGDET F WHERE F.SEGDET_NUMSOL = A.SOLMAE_NUMERO AND F.SEGDET_CODINS = 32 AND TRIM(F.SEGDET_OBSERV) IS NOT NULL AND TRIM(F.SEGDET_OBSDES) IS NULL) AS NO_RESUELTO "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & "  INNER JOIN CLI_DATGEN D ON D.DATGEN_TIPDOC = A.SOLMAE_TITTDO AND D.DATGEN_NUMDOC = A.SOLMAE_TITNDO "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CRE_SOLINM E ON E.SOLINM_NUMSOL = A.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & "   LEFT JOIN PRY_DATGEN F ON F.DATGEN_CODIGO = E.SOLINM_PRYCOD "
   g_str_Parame = g_str_Parame & "  WHERE A.SOLMAE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "    AND A.SOLMAE_CODINS >= 31 "
   g_str_Parame = g_str_Parame & "    AND A.SOLMAE_CODINS <= 41 "
   g_str_Parame = g_str_Parame & "    AND ((SELECT COUNT(*) FROM TRA_GASADM B WHERE B.GASADM_NUMSOL = A.SOLMAE_NUMERO AND GASADM_SITUAC = 1) = 0 "
   g_str_Parame = g_str_Parame & "     OR  (SELECT COUNT(*) FROM TRA_GASADM B WHERE B.GASADM_NUMSOL = A.SOLMAE_NUMERO AND GASADM_SITUAC = 2) > 0 )"
   g_str_Parame = g_str_Parame & "  ORDER BY A.SOLMAE_NUMERO ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   grd_Listad.Redraw = False
   Call gs_LimpiaGrid(grd_Listad)
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!TOTGSTCIE > 0 Then
            r_var_ColCel = modgen_g_con_ColNeg
         Else
            r_var_ColCel = modgen_g_con_ColRoj
         End If
         If g_rst_Princi!NO_RESUELTO > 0 Then
            r_var_ColCel = modgen_g_con_ColVer
         End If
         
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.CellForeColor = r_var_ColCel
         grd_Listad.Text = g_rst_Princi!NOM_PRODUCTO
         
         grd_Listad.Col = 1
         grd_Listad.CellForeColor = r_var_ColCel
         If IsNull(g_rst_Princi!DATGEN_TITULO) Then
            grd_Listad.Text = ""
         Else
            grd_Listad.Text = Trim(g_rst_Princi!DATGEN_TITULO)
         End If
         
         grd_Listad.Col = 2
         grd_Listad.CellForeColor = r_var_ColCel
         grd_Listad.Text = fs_ObtnerTipo(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
         
         grd_Listad.Col = 3
         grd_Listad.CellForeColor = r_var_ColCel
         grd_Listad.Text = Left(g_rst_Princi!SOLMAE_NUMERO, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 4, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 7, 2) & "-" & Right(g_rst_Princi!SOLMAE_NUMERO, 4)
         
         grd_Listad.Col = 4
         grd_Listad.CellForeColor = r_var_ColCel
         grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
         
         grd_Listad.Col = 5
         grd_Listad.CellForeColor = r_var_ColCel
         grd_Listad.Text = g_rst_Princi!NOM_CLIENTE 'moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
         
         grd_Listad.Col = 6
         grd_Listad.CellForeColor = r_var_ColCel
         grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
         
         grd_Listad.Col = 7
         grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CODPRD & "")
         
         grd_Listad.Col = 8
         grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CODSUB & "")
         
         grd_Listad.Col = 9
         grd_Listad.Text = g_rst_Princi!SOLMAE_FECSOL
         
         grd_Listad.Col = 10
         grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_TIPMON)
         
         grd_Listad.Col = 11
         grd_Listad.CellForeColor = r_var_ColCel
         grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!FEC_INGRESO))
         
         grd_Listad.Col = 12
         grd_Listad.Text = g_rst_Princi!FEC_INGRESO
         
         grd_Listad.Col = 13
         Select Case r_var_ColCel
               Case modgen_g_con_ColNeg:
                    r_int_OpeAsi = r_int_OpeAsi + 1
                    grd_Listad.Text = "ASIGNADA"
               Case modgen_g_con_ColRoj:
                    r_int_OpeNue = r_int_OpeNue + 1
                    grd_Listad.Text = "NUEVA"
               Case modgen_g_con_ColVer:
                    r_int_OpeObs = r_int_OpeObs + 1
                    grd_Listad.Text = "OBSERVADA"
         End Select
         
         r_int_TotReg = r_int_TotReg + 1
         g_rst_Princi.MoveNext
      Loop
   End If
   
   grd_Listad.Redraw = True
   
   If grd_Listad.Rows = 0 Then
      cmd_GasAdm.Enabled = False
      MsgBox "No se encontraron Solicitudes Pendientes de Asignación de Gastos Administrativos.", vbInformation, modgen_g_str_NomPlt
   Else
      'Ordenando por Nombre de Cliente
      pnl_Tit_NomCli.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 5, "C")
      Call gs_UbiIniGrid(grd_Listad)
   End If
   
   pnl_OpeAsi.Caption = r_int_OpeAsi & " "
   pnl_OpeNue.Caption = r_int_OpeNue & " "
   pnl_OpeObs.Caption = r_int_OpeObs & " "
   pnl_TotReg.Caption = r_int_TotReg & " "
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_NroFil     As Integer
Dim r_int_Conta      As Integer
Dim r_int_Fila       As Integer

   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
      
   With r_obj_Excel.ActiveSheet
      .Columns("A").HorizontalAlignment = xlHAlignLeft
      .Columns("B").HorizontalAlignment = xlHAlignLeft
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").HorizontalAlignment = xlHAlignLeft
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      
      r_int_NroFil = 2
      
      .Cells(1, 5) = "FECHA IMPRESION: " & Format(date, "dd/mm/yyyy")
      .Range(.Cells(1, 5), .Cells(1, 9)).Merge
      .Cells(1, 5).HorizontalAlignment = xlHAlignCenter
      
      .Cells(r_int_NroFil, 1) = "ASIGNACION DE GASTOS DE CIERRE"
      .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 1).Font.Bold = True
      .Range("A" & r_int_NroFil & ":I" & r_int_NroFil).Merge
      
      r_int_NroFil = 1
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 9)).Font.Bold = True
      r_int_NroFil = 4
      .Cells(r_int_NroFil, 1) = "PRODUCTO"
      .Cells(r_int_NroFil, 2) = "PROYECTO"
      .Cells(r_int_NroFil, 3) = "TIPO"
      .Cells(r_int_NroFil, 4) = "NRO. SOLICITUD"
      .Cells(r_int_NroFil, 5) = "ID CLIENTE"
      .Cells(r_int_NroFil, 6) = "APELLIDOS Y NOMBRES"
      .Cells(r_int_NroFil, 7) = "FECHA SOLICITUD"
      .Cells(r_int_NroFil, 8) = "F. ING. INSTANCIA"
      .Cells(r_int_NroFil, 9) = "ESTADO OPERACION"
      
      .Columns("A").ColumnWidth = 56
      .Columns("B").ColumnWidth = 56
      .Columns("C").ColumnWidth = 15
      .Columns("D").ColumnWidth = 18
      .Columns("E").ColumnWidth = 15
      .Columns("F").ColumnWidth = 46
      .Columns("G").ColumnWidth = 16
      .Columns("H").ColumnWidth = 17
      .Columns("I").ColumnWidth = 20
      
      r_int_Fila = 0
      For r_int_Conta = 1 To grd_Listad.Rows
         .Cells(r_int_Conta + 4, 1) = UCase(grd_Listad.TextMatrix(r_int_Fila, 0))
         .Cells(r_int_Conta + 4, 2) = UCase(grd_Listad.TextMatrix(r_int_Fila, 1))
         .Cells(r_int_Conta + 4, 3) = UCase(grd_Listad.TextMatrix(r_int_Fila, 2))
         .Cells(r_int_Conta + 4, 4) = UCase(grd_Listad.TextMatrix(r_int_Fila, 3))
         .Cells(r_int_Conta + 4, 5) = UCase(grd_Listad.TextMatrix(r_int_Fila, 4))
         .Cells(r_int_Conta + 4, 6) = UCase(grd_Listad.TextMatrix(r_int_Fila, 5))
         .Cells(r_int_Conta + 4, 7) = "'" & UCase(grd_Listad.TextMatrix(r_int_Fila, 6))
         .Cells(r_int_Conta + 4, 8) = "'" & UCase(grd_Listad.TextMatrix(r_int_Fila, 11))
         .Cells(r_int_Conta + 4, 9) = UCase(grd_Listad.TextMatrix(r_int_Fila, 13))
         r_int_Fila = r_int_Fila + 1
      Next
      'TITUTLO
      .Range(.Cells(2, 1), .Cells(2, 9)).HorizontalAlignment = xlHAlignCenter
      'NOMBRE COLUMNA
      .Range(.Cells(4, 1), .Cells(4, 9)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 1), .Cells(4, 9)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(4, 1), .Cells(4, 9)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(4, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(4, 9)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(4, 9)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(4, 9)).Borders(xlInsideVertical).LineStyle = xlContinuous
   End With
   r_obj_Excel.Sheets(1).Name = "Asig. Gastos Cierre"
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

Private Sub pnl_Tit_DocIde_Click()
   If Len(Trim(pnl_Tit_DocIde.Tag)) = 0 Or pnl_Tit_DocIde.Tag = "D" Then
      pnl_Tit_DocIde.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 4, "C")
   Else
      pnl_Tit_DocIde.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 4, "C-")
   End If
End Sub

Private Sub pnl_Tit_FecEva_Click()
   If Len(Trim(pnl_Tit_FecEva.Tag)) = 0 Or pnl_Tit_FecEva.Tag = "D" Then
      pnl_Tit_FecEva.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 12, "N")
   Else
      pnl_Tit_FecEva.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 12, "N-")
   End If
End Sub

Private Sub pnl_Tit_FecSol_Click()
   If Len(Trim(pnl_Tit_FecSol.Tag)) = 0 Or pnl_Tit_FecSol.Tag = "D" Then
      pnl_Tit_FecSol.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 9, "N")
   Else
      pnl_Tit_FecSol.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 9, "N-")
   End If
End Sub

Private Sub pnl_Tit_NomCli_Click()
   If Len(Trim(pnl_Tit_NomCli.Tag)) = 0 Or pnl_Tit_NomCli.Tag = "D" Then
      pnl_Tit_NomCli.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 5, "C")
   Else
      pnl_Tit_NomCli.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 5, "C-")
   End If
End Sub

Private Sub pnl_Tit_NumSol_Click()
   If Len(Trim(pnl_Tit_NumSol.Tag)) = 0 Or pnl_Tit_NumSol.Tag = "D" Then
      pnl_Tit_NumSol.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 3, "C")
   Else
      pnl_Tit_NumSol.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 3, "C-")
   End If
End Sub

Private Sub pnl_Tit_Produc_Click()
   If Len(Trim(pnl_Tit_Produc.Tag)) = 0 Or pnl_Tit_Produc.Tag = "D" Then
      pnl_Tit_Produc.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Tit_Produc.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub

Private Sub pnl_Tit_Proyec_Click()
   If Len(Trim(pnl_Tit_Proyec.Tag)) = 0 Or pnl_Tit_Proyec.Tag = "D" Then
      pnl_Tit_Proyec.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Tit_Proyec.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub

Private Sub pnl_Tit_Tipo_Click()
   If Len(Trim(pnl_Tit_Tipo.Tag)) = 0 Or pnl_Tit_Tipo.Tag = "D" Then
      pnl_Tit_Tipo.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_Tit_Tipo.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub

Private Function fs_ObtnerTipo(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String) As String
Dim r_str_Parame  As String

   fs_ObtnerTipo = ""
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT COUNT(SOLMAE_TITNDO) AS CANTIDAD  "
   r_str_Parame = r_str_Parame & "   FROM CRE_SOLMAE  "
   r_str_Parame = r_str_Parame & "  WHERE SOLMAE_TITTDO = " & p_TipDoc & ""
   r_str_Parame = r_str_Parame & "    AND SOLMAE_TITNDO = '" & p_NumDoc & "'"
    
   If Not gf_EjecutaSQL(r_str_Parame, g_rst_GenAux, 3) Then
      Exit Function
   End If
     
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      g_rst_GenAux.MoveFirst
      If g_rst_GenAux!CANTIDAD = 1 Then
         fs_ObtnerTipo = "N"
      Else
         fs_ObtnerTipo = "R"
      End If
   End If
End Function

