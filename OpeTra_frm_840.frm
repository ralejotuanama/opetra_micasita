VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_Ges_TecPro_16 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11400
   Icon            =   "OpeTra_frm_840.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10350
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   10340
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      _Version        =   65536
      _ExtentX        =   20135
      _ExtentY        =   18239
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   585
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   11295
         _Version        =   65536
         _ExtentX        =   19923
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   210
            Left            =   660
            TabIndex        =   2
            Top             =   240
            Width           =   5280
            _Version        =   65536
            _ExtentX        =   9313
            _ExtentY        =   370
            _StockProps     =   15
            Caption         =   "Posición Consolidada de Entidades Técnicas"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.24
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
            Picture         =   "OpeTra_frm_840.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel12 
         Height          =   645
         Left            =   60
         TabIndex        =   3
         Top             =   690
         Width           =   11295
         _Version        =   65536
         _ExtentX        =   19923
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10680
            Picture         =   "OpeTra_frm_840.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_840.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   60
            TabIndex        =   6
            Top             =   1740
            Width           =   1065
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   1155
         Left            =   60
         TabIndex        =   7
         Top             =   1380
         Width           =   11295
         _Version        =   65536
         _ExtentX        =   19923
         _ExtentY        =   2037
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
         Begin Threed.SSPanel pnl_RazSoc 
            Height          =   315
            Left            =   1620
            TabIndex        =   8
            Top             =   450
            Width           =   9435
            _Version        =   65536
            _ExtentX        =   16642
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_TipDoc 
            Height          =   315
            Left            =   1620
            TabIndex        =   9
            Top             =   120
            Width           =   4575
            _Version        =   65536
            _ExtentX        =   8070
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_NroDoc 
            Height          =   315
            Left            =   8040
            TabIndex        =   10
            Top             =   120
            Width           =   3015
            _Version        =   65536
            _ExtentX        =   5318
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_TipEmp 
            Height          =   315
            Left            =   1620
            TabIndex        =   11
            Top             =   780
            Width           =   4575
            _Version        =   65536
            _ExtentX        =   8070
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Label lbl_TipDoc 
            Caption         =   "Tipo Documento:"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   150
            Width           =   1335
         End
         Begin VB.Label lbl_NumDoc 
            Caption         =   "Nro. Documento:"
            Height          =   225
            Left            =   6450
            TabIndex        =   14
            Top             =   150
            Width           =   1335
         End
         Begin VB.Label lbl_RazSoc 
            Caption         =   "Razón Social:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lbl_TipEmp 
            Caption         =   "Tipo Empresa:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   810
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   7665
         Left            =   60
         TabIndex        =   16
         Top             =   2595
         Width           =   11295
         _Version        =   65536
         _ExtentX        =   19923
         _ExtentY        =   13520
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
         Begin TabDlg.SSTab SSTab1 
            Height          =   7620
            Left            =   30
            TabIndex        =   17
            Top             =   30
            Width           =   11235
            _ExtentX        =   19817
            _ExtentY        =   13441
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Resumen"
            TabPicture(0)   =   "OpeTra_frm_840.frx":0A62
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SSPanel41"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Garantías"
            TabPicture(1)   =   "OpeTra_frm_840.frx":0A7E
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SSPanel13"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).ControlCount=   1
            Begin Threed.SSPanel SSPanel41 
               Height          =   7185
               Left            =   45
               TabIndex        =   18
               Top             =   360
               Width           =   11145
               _Version        =   65536
               _ExtentX        =   19659
               _ExtentY        =   12674
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
               Begin MSFlexGridLib.MSFlexGrid grd_Listad 
                  Height          =   7170
                  Index           =   0
                  Left            =   45
                  TabIndex        =   19
                  Top             =   30
                  Width           =   11100
                  _ExtentX        =   19579
                  _ExtentY        =   12647
                  _Version        =   393216
                  Rows            =   21
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
            End
            Begin Threed.SSPanel SSPanel13 
               Height          =   7185
               Left            =   -74955
               TabIndex        =   20
               Top             =   360
               Width           =   11145
               _Version        =   65536
               _ExtentX        =   19659
               _ExtentY        =   12674
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
               Begin MSFlexGridLib.MSFlexGrid grd_Listad 
                  Height          =   7170
                  Index           =   1
                  Left            =   45
                  TabIndex        =   21
                  Top             =   30
                  Width           =   11100
                  _ExtentX        =   19579
                  _ExtentY        =   12647
                  _Version        =   393216
                  Rows            =   21
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
            End
         End
      End
   End
End
Attribute VB_Name = "frm_Ges_TecPro_16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim l_str_PerMes As String
'Dim l_str_PerAno As String
Dim l_int_TipGar As Integer
Dim l_int_MonGar As Integer
Dim l_dbl_MtoGar As Double
Dim l_str_CodSbs As String

Private Sub Form_Load()
   Dim r_arr_Mtz()      As moddat_g_tpo_DatCom
   
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_RazSoc.Caption = ""
   
   'Limpia grids
   Call fs_IniciaGrid

   pnl_TipDoc.Caption = moddat_gf_Consulta_ParDes("118", CStr(moddat_g_int_TipDoc))
   pnl_NroDoc.Caption = moddat_g_str_NumDoc
   pnl_RazSoc.Caption = moddat_g_str_NomCli
   pnl_TipEmp.Caption = moddat_g_str_Descri
   
   'Buscar Información de Resumen ET
   Call fs_DatResEte(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(0))             'Buscar Información de Resumen Empresa
   Call fs_DatGarEte(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(1))             'Buscar Información de Garantía
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_DatResEte(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_GrdLis As MSFlexGrid)
Dim r_str_FecAct  As String

   r_str_FecAct = Format(Now, "yyyymmdd")
   
   'Buscando Información del Crédito
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT MAEETE_TIPDOC    , MAEETE_NUMDOC                  , MAEETE_TIPEMP                  , NVL(MAEETE_LINASI_IND,0) AS LINASI_IND, NVL(MAEETE_LINASI_DIR,0) AS LINASI_DIR, MAEPRV_RAZSOC, MAEETE_PORRET, MAEETE_FECVCT, MAEETE_CODCIU, MAEETE_NOMCOR, "
   g_str_Parame = g_str_Parame & "        MAEETE_APEPAT    , MAEETE_APEMAT                  , MAEETE_APECAS                  , MAEETE_PRINOM                         , MAEETE_SEGNOM                         , MAEETE_NOMREP, MAEETE_TDOREP, MAEETE_NDOREP, MAEETE_DIRREP, "
   g_str_Parame = g_str_Parame & "        MAEETE_TELREP    , NVL(MAEETE_ADMFLJ,0) AS ADMFLJ , NVL(MAEETE_IMPHIP,0) AS IMPHIP , NVL(MAEETE_IMPLIQ,0) AS IMPLIQ        , MAEETE_FECAPR                         , MAEETE_UBIGEO, MAEETE_CODSBS, MAEETE_SITUAC, A.SEGUSUCRE, "
   g_str_Parame = g_str_Parame & "        MAEETE_LINREV_IND, MAEETE_LINNRE_IND              , MAEETE_LINREV_DIR              , MAEETE_LINNRE_DIR                     , "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT COUNT(MAECFI_NUMREF)  "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "               WHERE MAECFI_TIPDOC = MAEETE_TIPDOC AND MAECFI_NUMDOC = MAEETE_NUMDOC AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                 AND MAECFI_CODPRD IN ('026','027') AND MAECFI_CODMOD <> '005' AND MAECFI_CODMOD <> '008' "
   g_str_Parame = g_str_Parame & "               GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC ),0) AS NRO_CARTA_CF, "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT COUNT(MAECFI_NUMREF)  "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "               WHERE MAECFI_TIPDOC = MAEETE_TIPDOC AND MAECFI_NUMDOC = MAEETE_NUMDOC AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                 AND MAECFI_CODPRD IN ('026') AND MAECFI_CODMOD = '005'"
   g_str_Parame = g_str_Parame & "               GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC ),0) AS NRO_CARTA_AD, "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT COUNT(MAECFI_NUMREF)  "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "               WHERE MAECFI_TIPDOC = MAEETE_TIPDOC AND MAECFI_NUMDOC = MAEETE_NUMDOC AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                 AND MAECFI_CODPRD IN ('026','027') AND MAECFI_CODMOD = '008'"
   g_str_Parame = g_str_Parame & "               GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC ),0) AS NRO_CARTA_CSO, "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT NVL(SUM(MAECFI_GARFIA),0) "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "               WHERE MAECFI_TIPDOC = MAEETE_TIPDOC AND MAECFI_NUMDOC = MAEETE_NUMDOC AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                 AND MAECFI_CODPRD IN ('026','027') AND MAECFI_CODMOD <> '005' AND MAECFI_CODMOD <> '008' "
   g_str_Parame = g_str_Parame & "               GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC),0) AS LINEA_UTILIZADA_CF, "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT NVL(SUM(MAECFI_GARFIA),0) "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "               WHERE MAECFI_TIPDOC = MAEETE_TIPDOC AND MAECFI_NUMDOC = MAEETE_NUMDOC AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                 AND MAECFI_CODPRD IN ('026') AND MAECFI_CODMOD = '005'"
   g_str_Parame = g_str_Parame & "               GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC),0) AS LINEA_UTILIZADA_AD, "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT NVL(SUM(MAECFI_GARFIA),0) "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "               WHERE MAECFI_TIPDOC = MAEETE_TIPDOC AND MAECFI_NUMDOC = MAEETE_NUMDOC AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                 AND MAECFI_CODPRD IN ('026','027') AND MAECFI_CODMOD = '008'"
   g_str_Parame = g_str_Parame & "               GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC),0) AS LINEA_UTILIZADA_CSO, "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT NVL(SUM(MAECFI_GARFIA),0) "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "               WHERE MAECFI_TIPDOC = MAEETE_TIPDOC AND MAECFI_NUMDOC = MAEETE_NUMDOC AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                 AND MAECFI_CODPRD IN ('026','027') "
   g_str_Parame = g_str_Parame & "               GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC),0) AS LINEA_UTILIZADA_IND, "
      
   g_str_Parame = g_str_Parame & "        NVL(( SELECT NVL(SUM(NVL(MAEGAR_MTOGAR_INM,0) + NVL(MAEGAR_MTOGAR_ES1,0) + NVL(MAEGAR_MTOGAR_ES2,0) + NVL(MAEGAR_MTOGAR_DE1,0) + NVL(MAEGAR_MTOGAR_DE2,0)),0)"
   g_str_Parame = g_str_Parame & "                FROM TPR_MAEGAR INNER JOIN TPR_MAECFI ON MAECFI_TIPDOC = MAEGAR_TIPDOC AND MAECFI_NUMDOC = MAEGAR_NUMDOC AND TRIM(MAECFI_NUMREF) = TRIM (MAEGAR_NUMREF) AND MAECFI_CODPRD IN ('026','027')"
   g_str_Parame = g_str_Parame & "               WHERE MAEGAR_TIPDOC = MAEETE_TIPDOC And MAEGAR_NUMDOC = MAEETE_NUMDOC And MAEGAR_SITUAC = 1 And MAEGAR_TIPGAR = 1"
   g_str_Parame = g_str_Parame & "               GROUP BY MAEGAR_TIPDOC, MAEGAR_NUMDOC),0) AS GARANTIA_LIQUIDA_IND, "
                     
   g_str_Parame = g_str_Parame & "        NVL(( SELECT NVL(SUM(NVL(MAEGAR_MTOGAR_INM,0) + NVL(MAEGAR_MTOGAR_ES1,0) + NVL(MAEGAR_MTOGAR_ES2,0) + NVL(MAEGAR_MTOGAR_DE1,0) + NVL(MAEGAR_MTOGAR_DE2,0)),0)"
   g_str_Parame = g_str_Parame & "                FROM TPR_MAEGAR"
   g_str_Parame = g_str_Parame & "               WHERE MAEGAR_TIPDOC = MAEETE_TIPDOC And MAEGAR_NUMDOC = MAEETE_NUMDOC And MAEGAR_SITUAC = 1 And MAEGAR_TIPGAR = 2"
   g_str_Parame = g_str_Parame & "               GROUP BY MAEGAR_TIPDOC, MAEGAR_NUMDOC),0) AS GARANTIA_HIPOTECARIO_IND, "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT COUNT(MAECFI_NUMREF)  "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "               WHERE MAECFI_TIPDOC = MAEETE_TIPDOC AND MAECFI_NUMDOC = MAEETE_NUMDOC AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                 AND MAECFI_CODPRD = '008' AND MAECFI_CODSUB = '008' AND MAECFI_CODMOD = '001' "
   g_str_Parame = g_str_Parame & "               GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC ),0) AS NRO_DIR_LIN_CREDITO, "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT COUNT(MAECFI_NUMREF)  "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "               WHERE MAECFI_TIPDOC = MAEETE_TIPDOC AND MAECFI_NUMDOC = MAEETE_NUMDOC AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                 AND MAECFI_CODPRD = '008' AND MAECFI_CODSUB = '008' AND MAECFI_CODMOD = '002' "
   g_str_Parame = g_str_Parame & "               GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC ),0) AS NRO_DIR_CRED_PUNTUAL, "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT NVL(SUM(MAECFI_IMPFIA),0) "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "               WHERE MAECFI_TIPDOC = MAEETE_TIPDOC AND MAECFI_NUMDOC = MAEETE_NUMDOC AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                 AND MAECFI_CODPRD = '008' AND MAECFI_CODSUB = '008' AND MAECFI_CODMOD = '001' "
   g_str_Parame = g_str_Parame & "               GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC),0) AS LINEA_UTIL_DIR_LIN_CREDITO, "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT NVL(SUM(MAECFI_IMPFIA),0) "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "               WHERE MAECFI_TIPDOC = MAEETE_TIPDOC AND MAECFI_NUMDOC = MAEETE_NUMDOC AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                 AND MAECFI_CODPRD = '008' AND MAECFI_CODSUB = '008' AND MAECFI_CODMOD = '002' "
   g_str_Parame = g_str_Parame & "               GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC),0) AS LINEA_UTIL_DIR_CRE_PUNTUAL, "
     
   g_str_Parame = g_str_Parame & "        NVL(( SELECT NVL(SUM(MAECFI_IMPFIA),0) "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "               WHERE MAECFI_TIPDOC = MAEETE_TIPDOC AND MAECFI_NUMDOC = MAEETE_NUMDOC AND MAECFI_CODPRD = '008' AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "               GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC),0) AS LINEA_UTILIZADA_DIR, "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT NVL(SUM(NVL(MAEGAR_MTOGAR_INM,0) + NVL(MAEGAR_MTOGAR_ES1,0) + NVL(MAEGAR_MTOGAR_ES2,0) + NVL(MAEGAR_MTOGAR_DE1,0) + NVL(MAEGAR_MTOGAR_DE2,0)),0)"
   g_str_Parame = g_str_Parame & "                FROM TPR_MAEGAR INNER JOIN TPR_MAECFI ON MAECFI_TIPDOC = MAEGAR_TIPDOC AND MAECFI_NUMDOC = MAEGAR_NUMDOC AND TRIM(MAECFI_NUMREF) = TRIM (MAEGAR_NUMREF) AND MAECFI_CODPRD IN ('008')"
   g_str_Parame = g_str_Parame & "               WHERE MAEGAR_TIPDOC = MAEETE_TIPDOC And MAEGAR_NUMDOC = MAEETE_NUMDOC And MAEGAR_SITUAC = 1 And MAEGAR_TIPGAR = 1"
   g_str_Parame = g_str_Parame & "               GROUP BY MAEGAR_TIPDOC, MAEGAR_NUMDOC),0) AS GARANTIA_LIQUIDA_DIR, "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT ROUND(SUM(CASE WHEN MAECFI_CODPRD = '008' AND MAECFI_CODSUB = '008' AND MAECFI_CODMOD = '001' THEN "
   g_str_Parame = g_str_Parame & "                                    MAECFI_IMPFIA * (power((1 + C.MAECFI_PORTEA/100),(trunc(to_date(TO_CHAR(" & "'" & r_str_FecAct & "' ),'YYYYMMDD')) - trunc(to_date(TO_CHAR(C.MAECFI_EMIFIA),'YYYYMMDD')))/360))- C.MAECFI_IMPFIA"
   g_str_Parame = g_str_Parame & "                               ELSE 0 END) ,2) "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI C"
   g_str_Parame = g_str_Parame & "               WHERE C.MAECFI_TIPDOC = MAEETE_TIPDOC"
   g_str_Parame = g_str_Parame & "                 AND C.MAECFI_NUMDOC = MAEETE_NUMDOC"
   g_str_Parame = g_str_Parame & "                 AND C.MAECFI_CODPRD = '008'"
   g_str_Parame = g_str_Parame & "                 AND C.MAECFI_CODSUB = '008'"
   g_str_Parame = g_str_Parame & "                 AND C.MAECFI_CODMOD = '001'"
   g_str_Parame = g_str_Parame & "                 AND C.MAECFI_SITUAC = 1),0) AS INTERES_ACUM_ACTUAL_LC, "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT ROUND(SUM(CASE WHEN MAECFI_CODPRD = '008' AND MAECFI_CODSUB = '008' AND MAECFI_CODMOD = '002' THEN "
   g_str_Parame = g_str_Parame & "                                    MAECFI_IMPFIA * (power((1 + C.MAECFI_PORTEA/100),(trunc(to_date(TO_CHAR(" & "'" & r_str_FecAct & "' ),'YYYYMMDD')) - trunc(to_date(TO_CHAR(C.MAECFI_EMIFIA),'YYYYMMDD')))/360))- C.MAECFI_IMPFIA"
   g_str_Parame = g_str_Parame & "                               ELSE 0 END) ,2) "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI C"
   g_str_Parame = g_str_Parame & "               WHERE C.MAECFI_TIPDOC = MAEETE_TIPDOC "
   g_str_Parame = g_str_Parame & "                 AND C.MAECFI_NUMDOC = MAEETE_NUMDOC "
   g_str_Parame = g_str_Parame & "                 AND C.MAECFI_CODPRD = '008' "
   g_str_Parame = g_str_Parame & "                 AND C.MAECFI_CODSUB = '008' "
   g_str_Parame = g_str_Parame & "                 AND C.MAECFI_CODMOD = '002' "
   g_str_Parame = g_str_Parame & "                 AND C.MAECFI_SITUAC = 1),0) AS INTERES_ACUM_ACTUAL_CP, "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT ROUND(SUM(CASE WHEN MAECFI_CODPRD = '008' AND MAECFI_CODSUB = '008' AND MAECFI_CODMOD = '001' THEN "
   g_str_Parame = g_str_Parame & "                                    MAECFI_IMPFIA * (power((1 + C.MAECFI_PORTEA/100),(trunc(to_date(TO_CHAR(C.MAECFI_VTOFIA),'YYYYMMDD')) - trunc(to_date(TO_CHAR(C.MAECFI_EMIFIA),'YYYYMMDD')))/360))- C.MAECFI_IMPFIA"
   g_str_Parame = g_str_Parame & "                               ELSE 0 END) ,2) "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI C "
   g_str_Parame = g_str_Parame & "               WHERE C.MAECFI_TIPDOC = MAEETE_TIPDOC "
   g_str_Parame = g_str_Parame & "                 AND C.MAECFI_NUMDOC = MAEETE_NUMDOC "
   g_str_Parame = g_str_Parame & "                 AND C.MAECFI_CODPRD = '008' "
   g_str_Parame = g_str_Parame & "                 AND C.MAECFI_CODSUB = '008' "
   g_str_Parame = g_str_Parame & "                 AND C.MAECFI_CODMOD = '001' "
   g_str_Parame = g_str_Parame & "                 AND C.MAECFI_SITUAC = 1),0) AS INTERES_ACUM_VCTO_LC, "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT ROUND(SUM(CASE WHEN MAECFI_CODPRD = '008' AND MAECFI_CODSUB = '008' AND MAECFI_CODMOD = '002' THEN "
   g_str_Parame = g_str_Parame & "                                    MAECFI_IMPFIA * (power((1 + C.MAECFI_PORTEA/100),(trunc(to_date(TO_CHAR(C.MAECFI_VTOFIA),'YYYYMMDD')) - trunc(to_date(TO_CHAR(C.MAECFI_EMIFIA),'YYYYMMDD')))/360))- C.MAECFI_IMPFIA"
   g_str_Parame = g_str_Parame & "                               ELSE 0 END) ,2) "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI C "
   g_str_Parame = g_str_Parame & "               WHERE C.MAECFI_TIPDOC = MAEETE_TIPDOC "
   g_str_Parame = g_str_Parame & "                 AND C.MAECFI_NUMDOC = MAEETE_NUMDOC "
   g_str_Parame = g_str_Parame & "                 AND C.MAECFI_CODPRD = '008' "
   g_str_Parame = g_str_Parame & "                 AND C.MAECFI_CODSUB = '008' "
   g_str_Parame = g_str_Parame & "                 AND C.MAECFI_CODMOD = '002' "
   g_str_Parame = g_str_Parame & "                 AND C.MAECFI_SITUAC = 1),0) AS INTERES_ACUM_VCTO_CP, "
   
   g_str_Parame = g_str_Parame & "        NVL((SELECT NVL(SUM(NVL(D.MAEGAR_MTOGAR_INM,0) + NVL(D.MAEGAR_MTOGAR_ES1,0) + NVL(D.MAEGAR_MTOGAR_ES2,0) + NVL(D.MAEGAR_MTOGAR_DE1,0) + NVL(D.MAEGAR_MTOGAR_DE2,0)),0)"
   g_str_Parame = g_str_Parame & "               FROM TPR_MAEGAR D "
   g_str_Parame = g_str_Parame & "              WHERE D.MAEGAR_TIPDOC = MAEETE_TIPDOC "
   g_str_Parame = g_str_Parame & "                AND D.MAEGAR_NUMDOC = MAEETE_NUMDOC "
   g_str_Parame = g_str_Parame & "                AND D.MAEGAR_TIPGAR = 1 "
   g_str_Parame = g_str_Parame & "                AND D.MAEGAR_SITUAC = 1),0) AS GARANTIA_LIQUIDA, "
         
   g_str_Parame = g_str_Parame & "        NVL((SELECT NVL(SUM(NVL(D.MAEGAR_MTOGAR_INM,0) + NVL(D.MAEGAR_MTOGAR_ES1,0) + NVL(D.MAEGAR_MTOGAR_ES2,0) + NVL(D.MAEGAR_MTOGAR_DE1,0) + NVL(D.MAEGAR_MTOGAR_DE2,0)),0)"
   g_str_Parame = g_str_Parame & "               FROM TPR_MAEGAR D "
   g_str_Parame = g_str_Parame & "              WHERE D.MAEGAR_TIPDOC = MAEETE_TIPDOC "
   g_str_Parame = g_str_Parame & "                AND D.MAEGAR_NUMDOC = MAEETE_NUMDOC "
   g_str_Parame = g_str_Parame & "                AND D.MAEGAR_TIPGAR = 2 "
   g_str_Parame = g_str_Parame & "                AND D.MAEGAR_SITUAC = 1),0) AS GARANTIA_HIPOTECARIA, "
         
   g_str_Parame = g_str_Parame & "        NVL((SELECT NVL(SUM(NVL(D.MAEGAR_MTOGAR_INM,0) + NVL(D.MAEGAR_MTOGAR_ES1,0) + NVL(D.MAEGAR_MTOGAR_ES2,0) + NVL(D.MAEGAR_MTOGAR_DE1,0) + NVL(D.MAEGAR_MTOGAR_DE2,0)),0)"
   g_str_Parame = g_str_Parame & "               FROM TPR_MAEGAR D "
   g_str_Parame = g_str_Parame & "              WHERE D.MAEGAR_TIPDOC = MAEETE_TIPDOC "
   g_str_Parame = g_str_Parame & "                AND D.MAEGAR_NUMDOC = MAEETE_NUMDOC "
   g_str_Parame = g_str_Parame & "                AND D.MAEGAR_TIPGAR = 5 "
   g_str_Parame = g_str_Parame & "                AND D.MAEGAR_SITUAC = 1),0) AS GARANTIA_PAGARE "
        
   g_str_Parame = g_str_Parame & "   FROM TPR_MAEETE A INNER JOIN CNTBL_MAEPRV ON MAEETE_TIPDOC = MAEPRV_TIPDOC AND MAEETE_NUMDOC = MAEPRV_NUMDOC "
   g_str_Parame = g_str_Parame & "  WHERE MAEETE_TIPDOC = " & CStr(p_TipDoc) & " AND MAEETE_NUMDOC = '" & CStr(p_NumDoc) & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
       
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   moddat_g_int_TipMon = 1
   
   'Situación de Crédito
   moddat_g_int_Situac = g_rst_Princi!MAEETE_SITUAC
   moddat_g_str_Situac = moddat_gf_Consulta_ParDes("529", CStr(g_rst_Princi!MAEETE_SITUAC))
  
   Call gs_LimpiaGrid(p_GrdLis)
   p_GrdLis.Redraw = False
   
   'Cargando en Grid
   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   
   p_GrdLis.Col = 0
   p_GrdLis.CellFontBold = True
   p_GrdLis.Text = "Documento de Identidad"
   
   p_GrdLis.Col = 1
   p_GrdLis.CellFontBold = True
   p_GrdLis.Text = CStr(g_rst_Princi!MAEETE_TIPDOC) & " - " & CStr(g_rst_Princi!MAEETE_NUMDOC) 'moddat_gf_Consulta_ParDes("118", g_rst_Princi!MAEETE_TIPDOC)
      
   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.CellFontBold = True
   p_GrdLis.Text = "Cliente"
   
   p_GrdLis.Col = 1
   p_GrdLis.CellFontBold = True
   p_GrdLis.Text = CStr(g_rst_Princi!MAEPRV_RAZSOC)

   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.Text = "Código SBS"
   
   p_GrdLis.Col = 1
   p_GrdLis.Text = Trim(g_rst_Princi!MAEETE_CODSBS)

   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.Text = "Tipo de Empresa"
   
   p_GrdLis.Col = 1
   p_GrdLis.Text = moddat_g_str_Descri
   
   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.Text = "CIUU"
   
   p_GrdLis.Col = 1
   p_GrdLis.Text = moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!MAEETE_CODCIU))
   
   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.CellFontBold = True
   p_GrdLis.Text = "Situación"
   
   p_GrdLis.Col = 1
   p_GrdLis.CellFontBold = True
   p_GrdLis.Text = moddat_g_str_Situac
'
'   p_GrdLis.Rows = p_GrdLis.Rows + 2
'   p_GrdLis.Row = p_GrdLis.Rows - 1
'   p_GrdLis.Col = 0
'   p_GrdLis.Text = "Documento de Identidad"
'   p_GrdLis.Col = 1
'   p_GrdLis.Text = CStr(g_rst_Princi!MAEETE_TDOREP) & "-" & CStr(g_rst_Princi!MAEETE_NDOREP) 'moddat_gf_Consulta_ParDes("118", g_rst_Princi!MAEETE_TDOREP)
   
   p_GrdLis.Rows = p_GrdLis.Rows + 2
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.Text = "Representante Legal"
   p_GrdLis.Col = 1
   
   If Not IsNull(g_rst_Princi!MAEETE_TDOREP) Then
      If Not IsNull(g_rst_Princi!MAEETE_NDOREP) Then
         p_GrdLis.Text = CStr(g_rst_Princi!MAEETE_TDOREP) & " - " & CStr(g_rst_Princi!MAEETE_NDOREP) & " / " & Trim(g_rst_Princi!MAEETE_NOMREP)
      End If
   End If
   
   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.Text = "Dirección"
   
   p_GrdLis.Col = 1
   If Not IsNull(g_rst_Princi!MAEETE_DIRREP) Then
      p_GrdLis.Text = UCase(CStr(g_rst_Princi!MAEETE_DIRREP))
   End If
   
   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.Text = "Departamento - Provincia - Distrito"
   
   p_GrdLis.Col = 1
   If Not IsNull(g_rst_Princi!MAEETE_UBIGEO) Then
      p_GrdLis.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!MAEETE_UBIGEO, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!MAEETE_UBIGEO, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!MAEETE_UBIGEO))
   End If
   
   p_GrdLis.Rows = p_GrdLis.Rows + 2
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.CellFontBold = True
   p_GrdLis.Text = "Línea Asignada"
   
   p_GrdLis.Col = 1
   p_GrdLis.CellFontBold = True
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!LINASI_IND + g_rst_Princi!LINASI_DIR, 12, 2)
   
   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   
   p_GrdLis.Col = 0
   p_GrdLis.CellFontBold = True
   p_GrdLis.Text = "       Créditos Indirectos"
   
   p_GrdLis.Col = 1
   p_GrdLis.CellFontBold = True
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = "       " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!LINASI_IND, 12, 2)
   
   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.Text = "                Revolvente"

   p_GrdLis.Col = 1
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = "              " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!MAEETE_LINREV_IND, 12, 2)

   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.Text = "                No Revolvente"

   p_GrdLis.Col = 1
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = "              " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!MAEETE_LINNRE_IND, 12, 2)

'   p_GrdLis.Rows = p_GrdLis.Rows + 1
'   p_GrdLis.Row = p_GrdLis.Rows - 1
'   p_GrdLis.Col = 0
'   p_GrdLis.Text = "                Administración de Flujos"
'
'   p_GrdLis.Col = 1
'   p_GrdLis.CellFontName = "Lucida Console"
'   p_GrdLis.CellFontSize = 8
'   p_GrdLis.Text = "              " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!ADMFLJ, 12, 2)
'
'   p_GrdLis.Rows = p_GrdLis.Rows + 1
'   p_GrdLis.Row = p_GrdLis.Rows - 1
'   p_GrdLis.Col = 0
'   p_GrdLis.Text = "                Hipoteca"
'
'   p_GrdLis.Col = 1
'   p_GrdLis.CellFontName = "Lucida Console"
'   p_GrdLis.CellFontSize = 8
'   p_GrdLis.Text = "              " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!IMPHIP, 12, 2)
'
'   p_GrdLis.Rows = p_GrdLis.Rows + 1
'   p_GrdLis.Row = p_GrdLis.Rows - 1
'   p_GrdLis.Col = 0
'   p_GrdLis.Text = "                Líquida"
'
'   p_GrdLis.Col = 1
'   p_GrdLis.CellFontName = "Lucida Console"
'   p_GrdLis.CellFontSize = 8
'   p_GrdLis.Text = "              " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!IMPLIQ, 12, 2)
   
   
   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.CellFontBold = True
   p_GrdLis.Text = "       Créditos Directos"
   
   p_GrdLis.Col = 1
   p_GrdLis.CellFontBold = True
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = "       " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!LINASI_DIR, 12, 2)
   
   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.Text = "                Revolvente"

   p_GrdLis.Col = 1
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = "              " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!MAEETE_LINREV_DIR, 12, 2)

   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.Text = "                No Revolvente"

   p_GrdLis.Col = 1
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = "              " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!MAEETE_LINNRE_DIR, 12, 2)


   'Línea Utilizada
   p_GrdLis.Rows = p_GrdLis.Rows + 2
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.CellFontBold = True
   p_GrdLis.Text = "Línea Utilizada"
   
   p_GrdLis.Col = 1
   p_GrdLis.CellFontBold = True
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = " Cantidad: " & g_rst_Princi!NRO_CARTA_CF + g_rst_Princi!NRO_CARTA_AD + g_rst_Princi!NRO_CARTA_CSO + g_rst_Princi!NRO_DIR_LIN_CREDITO + g_rst_Princi!NRO_DIR_CRED_PUNTUAL & "    -    " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!LINEA_UTILIZADA_IND + g_rst_Princi!LINEA_UTILIZADA_DIR, 12, 2)   'gf_FormatoNumero(g_rst_Princi!LINEA_UTILIZADA_CF + g_rst_Princi!LINEA_UTILIZADA_AD + g_rst_Princi!LINEA_UTILIZADA_CSO, 12, 2)

   
   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.CellFontBold = True
   p_GrdLis.Text = "       Créditos Indirectos"
   
   p_GrdLis.Col = 1
   p_GrdLis.CellFontBold = True
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = "       " & "Cantidad: " & g_rst_Princi!NRO_CARTA_CF + g_rst_Princi!NRO_CARTA_AD + g_rst_Princi!NRO_CARTA_CSO & "    -    " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!LINEA_UTILIZADA_IND, 12, 2)
   
   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.Text = "                Cartas Fianzas"
   
   p_GrdLis.Col = 1
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = "              " & "Cantidad: " & g_rst_Princi!NRO_CARTA_CF & "    -    " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!LINEA_UTILIZADA_CF, 12, 2)
   
   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.Text = "                Adendas"
   
   p_GrdLis.Col = 1
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = "              " & "Cantidad: " & g_rst_Princi!NRO_CARTA_AD & "    -    " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!LINEA_UTILIZADA_AD, 12, 2)
   'p_GrdLis.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!LINEA_UTILIZADA_AD, 12, 2)
   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.Text = "                Cartas Seriedad Oferta"
   
   p_GrdLis.Col = 1
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = "              " & "Cantidad: " & g_rst_Princi!NRO_CARTA_CSO & "    -    " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!LINEA_UTILIZADA_CSO, 12, 2)
   
   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.CellFontBold = True
   p_GrdLis.Text = "       Créditos Directos"
   
   p_GrdLis.Col = 1
   p_GrdLis.CellFontBold = True
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = "       " & "Cantidad: " & g_rst_Princi!NRO_DIR_LIN_CREDITO + g_rst_Princi!NRO_DIR_CRED_PUNTUAL & "    -    " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!LINEA_UTILIZADA_DIR, 12, 2)
   
   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.Text = "                Línea de Crédito"
   
   p_GrdLis.Col = 1
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = "              " & "Cantidad: " & g_rst_Princi!NRO_DIR_LIN_CREDITO & "    -    " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!LINEA_UTIL_DIR_LIN_CREDITO, 12, 2)
   
   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.Text = "                Crédito Puntual"
   
   p_GrdLis.Col = 1
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = "              " & "Cantidad: " & g_rst_Princi!NRO_DIR_CRED_PUNTUAL & "    -    " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!LINEA_UTIL_DIR_CRE_PUNTUAL, 12, 2)

   'Línea Disponible
   p_GrdLis.Rows = p_GrdLis.Rows + 2
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.CellFontBold = True
   p_GrdLis.Text = "Línea Disponible"
   
   p_GrdLis.Col = 1
   p_GrdLis.CellFontBold = True
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero((g_rst_Princi!LINASI_IND + g_rst_Princi!LINASI_DIR) - (g_rst_Princi!LINEA_UTILIZADA_IND + g_rst_Princi!LINEA_UTILIZADA_DIR), 12, 2)
   
   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   
   p_GrdLis.Col = 0
   p_GrdLis.CellFontBold = True
   p_GrdLis.Text = "       Créditos Indirectos"
   
   p_GrdLis.Col = 1
'   p_GrdLis.CellFontBold = True
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = "       " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!LINASI_IND - g_rst_Princi!LINEA_UTILIZADA_IND, 12, 2)
   
   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.CellFontBold = True
   p_GrdLis.Text = "       Créditos Directos"
   
   p_GrdLis.Col = 1
'   p_GrdLis.CellFontBold = True
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = "       " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!LINASI_DIR - g_rst_Princi!LINEA_UTILIZADA_DIR, 12, 2)
   
   'Intereses
   p_GrdLis.Rows = p_GrdLis.Rows + 2
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.CellFontBold = True
   p_GrdLis.Text = "Interés Devengado"
   
'   p_GrdLis.Col = 1
'   p_GrdLis.CellFontBold = True
'   p_GrdLis.CellFontName = "Lucida Console"
'   p_GrdLis.CellFontSize = 8
'   p_GrdLis.Text = " Interés Total: " & g_rst_Princi!NRO_CARTA_CF + g_rst_Princi!NRO_CARTA_AD + g_rst_Princi!NRO_CARTA_CSO + g_rst_Princi!NRO_DIR_LIN_CREDITO + g_rst_Princi!NRO_DIR_CRED_PUNTUAL & "    -    " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!LINEA_UTILIZADA_IND + g_rst_Princi!LINEA_UTILIZADA_DIR, 12, 2)   'gf_FormatoNumero(g_rst_Princi!LINEA_UTILIZADA_CF + g_rst_Princi!LINEA_UTILIZADA_AD + g_rst_Princi!LINEA_UTILIZADA_CSO, 12, 2)


   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.CellFontBold = True
   p_GrdLis.Text = "       Créditos Directos"
   
   p_GrdLis.Col = 1
   p_GrdLis.CellFontBold = True
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = "       " & "Fec. Actual: " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!INTERES_ACUM_ACTUAL_LC + g_rst_Princi!INTERES_ACUM_ACTUAL_CP, 12, 2)
   '& "    -    " & "Fec. Vcto.: " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!LINEA_UTILIZADA_DIR, 12, 2)
   
   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 1
   p_GrdLis.CellFontBold = True
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = "       " & "Fec. Vcto. : " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!INTERES_ACUM_VCTO_LC + g_rst_Princi!INTERES_ACUM_VCTO_CP, 12, 2)


   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.Text = "                Línea de Crédito"
   
   p_GrdLis.Col = 1
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = "              " & "Fec. Actual: " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!INTERES_ACUM_ACTUAL_LC, 12, 2)
   
   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 1
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = "              " & "Fec. Vcto. : " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!INTERES_ACUM_VCTO_LC, 12, 2)

   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.Text = "                Crédito Puntual"
   
   p_GrdLis.Col = 1
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = "              " & "Fec. Actual: " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!INTERES_ACUM_ACTUAL_CP, 12, 2)
   
   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 1
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = "              " & "Fec. Vcto. : " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!INTERES_ACUM_VCTO_CP, 12, 2)

   
   'Garantías
   p_GrdLis.Rows = p_GrdLis.Rows + 2
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.CellFontBold = True
   p_GrdLis.Text = "Garantías"
   
   p_GrdLis.Col = 1
   p_GrdLis.CellFontBold = True
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!GARANTIA_HIPOTECARIA + g_rst_Princi!GARANTIA_LIQUIDA + g_rst_Princi!GARANTIA_PAGARE, 12, 2)
   
   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.Text = "                Hipoteca"
   
   p_GrdLis.Col = 1
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = "              " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!GARANTIA_HIPOTECARIA, 12, 2)
   
   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.Text = "                Líquida"
   
   p_GrdLis.Col = 1
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = "              " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!GARANTIA_LIQUIDA, 12, 2)
   
   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.Text = "                Pagaré"
   
   p_GrdLis.Col = 1
   p_GrdLis.CellFontName = "Lucida Console"
   p_GrdLis.CellFontSize = 8
   p_GrdLis.Text = "              " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!GARANTIA_PAGARE, 12, 2)
   
   '***************************************
   
   p_GrdLis.Rows = p_GrdLis.Rows + 2
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.Text = "Fecha Aprobación"
   
   p_GrdLis.Col = 1
   If Not IsNull(g_rst_Princi!MAEETE_FECAPR) Then
      p_GrdLis.Text = gf_FormatoFecha(CStr(g_rst_Princi!MAEETE_FECAPR))
   Else
      p_GrdLis.Text = ""
   End If
   
   p_GrdLis.Rows = p_GrdLis.Rows + 1
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.Text = "Fecha Vencimiento"
   
   p_GrdLis.Col = 1
   If Not IsNull(g_rst_Princi!MAEETE_FECVCT) Then
      p_GrdLis.Text = gf_FormatoFecha(CStr(g_rst_Princi!MAEETE_FECVCT))
   Else
      p_GrdLis.Text = ""
   End If

   p_GrdLis.Rows = p_GrdLis.Rows + 2
   p_GrdLis.Row = p_GrdLis.Rows - 1
   p_GrdLis.Col = 0
   p_GrdLis.Text = "Usuario"
   
   p_GrdLis.Col = 1
   p_GrdLis.Text = Trim(g_rst_Princi!SEGUSUCRE & "")
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   p_GrdLis.Redraw = True
   Call gs_UbiIniGrid(p_GrdLis)
End Sub
Private Sub fs_IniciaGrid()

Dim r_int_Contad     As Integer

   'Grid de Resumen
   grd_Listad(0).ColWidth(0) = 3000
   grd_Listad(0).ColWidth(1) = 7940
   grd_Listad(0).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(0).ColAlignment(1) = flexAlignLeftCenter
   grd_Listad(0).Rows = 0
   
   'Grid Posición Integral
   grd_Listad(1).ColWidth(0) = 3000
   grd_Listad(1).ColWidth(1) = 7940
   grd_Listad(1).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(1).ColAlignment(1) = flexAlignLeftCenter
   grd_Listad(1).Rows = 0
   
'   SSTab1.TabVisible(1) = False
End Sub

Private Sub fs_DatGarEte(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_GrdLis As MSFlexGrid)
   
   'Buscando Información de Garantía
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT MAEGAR_TIPGAR, MAEGAR_TIPMON_INM, " 'MAEGAR_NUMREF
   g_str_Parame = g_str_Parame & "         SUM(NVL(NVL(MAEGAR_MTOGAR_INM,0) + NVL(MAEGAR_MTOGAR_ES1,0) + NVL(MAEGAR_MTOGAR_ES2,0) + NVL(MAEGAR_MTOGAR_DE1,0) + NVL(MAEGAR_MTOGAR_DE2,0),0)) AS IMPORTE_GARANTIA "
   g_str_Parame = g_str_Parame & "    FROM TPR_MAEGAR "
   g_str_Parame = g_str_Parame & "   WHERE MAEGAR_TIPDOC = '" & p_TipDoc & "' "
   g_str_Parame = g_str_Parame & "     AND MAEGAR_NUMDOC = '" & p_NumDoc & "' "
   g_str_Parame = g_str_Parame & "     AND MAEGAR_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "   GROUP BY MAEGAR_TIPGAR , MAEGAR_TIPMON_INM"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      SSTab1.TabVisible(1) = False
      Exit Sub
   End If

   g_rst_Princi.MoveFirst
   
   Call gs_LimpiaGrid(p_GrdLis)
   p_GrdLis.Redraw = False
   
   'Cargando en Grid
   p_GrdLis.Rows = 0
   
   Do While Not g_rst_Princi.EOF

      p_GrdLis.Rows = p_GrdLis.Rows + 1
      p_GrdLis.Row = p_GrdLis.Rows - 1
      p_GrdLis.Col = 0
      p_GrdLis.Text = "Tipo Garantía"
   
      p_GrdLis.Col = 1
      p_GrdLis.Text = moddat_gf_Consulta_ParDes("527", CStr(g_rst_Princi!MAEGAR_TIPGAR))
   
      p_GrdLis.Rows = p_GrdLis.Rows + 1
      p_GrdLis.Row = p_GrdLis.Rows - 1
      p_GrdLis.Col = 0
      p_GrdLis.Text = "Moneda Garantía"
   
      p_GrdLis.Col = 1
      p_GrdLis.Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!MAEGAR_TIPMON_INM))
   
      p_GrdLis.Rows = p_GrdLis.Rows + 1
      p_GrdLis.Row = p_GrdLis.Rows - 1
      p_GrdLis.Col = 0
      p_GrdLis.Text = "Monto Garantía"

      p_GrdLis.Col = 1
      p_GrdLis.CellFontName = "Lucida Console"
      p_GrdLis.CellFontSize = 8
      p_GrdLis.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!MAEGAR_TIPMON_INM)) & " " & gf_FormatoNumero(g_rst_Princi!IMPORTE_GARANTIA, 12, 2)
   
   
      p_GrdLis.Rows = p_GrdLis.Rows + 1
      p_GrdLis.Row = p_GrdLis.Rows - 1
      p_GrdLis.Col = 0
      p_GrdLis.Text = "Créditos de Referencia"
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "   SELECT NVL(WM_CONCAT(TRIM(MAEGAR_NUMREF)),'-') AS REFERENCIA "
      g_str_Parame = g_str_Parame & "     FROM TPR_MAEGAR"
      g_str_Parame = g_str_Parame & "    WHERE MAEGAR_TIPDOC = '" & p_TipDoc & "' "
      g_str_Parame = g_str_Parame & "      AND MAEGAR_NUMDOC = '" & p_NumDoc & "' "
      g_str_Parame = g_str_Parame & "      AND MAEGAR_SITUAC = 1 "
      g_str_Parame = g_str_Parame & "      AND MAEGAR_TIPGAR = " & Trim(g_rst_Princi!MAEGAR_TIPGAR) & ""
      g_str_Parame = g_str_Parame & "    GROUP BY MAEGAR_TIPGAR "
             
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
         Exit Sub
      End If
   
      If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
         g_rst_GenAux.Close
         Set g_rst_GenAux = Nothing
      End If
      
      g_rst_GenAux.MoveFirst
      
      p_GrdLis.Col = 1
      p_GrdLis.Text = CStr(Trim(Replace(g_rst_GenAux!REFERENCIA, ",", " - ")))
      
      p_GrdLis.Rows = p_GrdLis.Rows + 1
      
      g_rst_Princi.MoveNext
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   p_GrdLis.Redraw = True
   Call gs_UbiIniGrid(p_GrdLis)
End Sub

Public Function gf_Buscar_NomEmp(ByVal p_CodEmp As Integer) As String
   gf_Buscar_NomEmp = ""
   
   g_str_Parame = "SELECT * FROM CTB_EMPSUP WHERE EMPSUP_CODIGO = " & p_CodEmp & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      gf_Buscar_NomEmp = Trim(g_rst_Listas!EMPSUP_NOMBRE)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function gf_Buscar_TipCre(ByVal p_CodCre As Integer) As String
   gf_Buscar_TipCre = ""
   g_str_Parame = "SELECT * FROM CTB_TIPCRE WHERE TIPCRE_CODIGO = " & p_CodCre
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      gf_Buscar_TipCre = Trim(g_rst_Listas!TIPCRE_DESCRI)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_ExpExc_Click()
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
      
   Call fs_GenExc
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_str_ParAux        As String
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
Dim r_int_NroFil        As Integer
Dim r_int_NoFlLi        As Integer
       
   Screen.MousePointer = 11
    
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 2
   r_obj_Excel.Workbooks.Add
   
   r_obj_Excel.Sheets(1).Name = "Resumen"
   
   With r_obj_Excel.Sheets(1)
      .Cells(2, 2) = "REPORTE DE ENTIDAD TECNICA - POSICION CONSOLIDADA"
      .Range(.Cells(2, 1), .Cells(2, 4)).Merge
      .Range(.Cells(2, 1), .Cells(2, 4)).Font.Bold = True
      .Range(.Cells(2, 1), .Cells(2, 4)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(2, 1), .Cells(2, 4)).Font.Size = 14
      
'      .Columns("A").ColumnWidth = 1
'      .Columns("B").ColumnWidth = 20
'      .Columns("B").NumberFormat = "###,###,###,##0.00"
'      .Columns("B").HorizontalAlignment = xlHAlignRight
'      .Columns("C").ColumnWidth = 20
'      .Columns("C").NumberFormat = "###,###,###,##0.00"
      .Columns("C").HorizontalAlignment = xlHAlignLeft
'      .Columns("D").ColumnWidth = 17
'      .Columns("D").HorizontalAlignment = xlHAlignCenter
'      .Columns("D").NumberFormat = "0.00"
'      .Columns("E").ColumnWidth = 17
'      .Columns("E").HorizontalAlignment = xlHAlignCenter
'      .Columns("E").NumberFormat = "0.00"
'      .Columns("F").ColumnWidth = 17
'      .Columns("F").HorizontalAlignment = xlHAlignCenter

      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(3, 1), .Cells(99, 99)).Font.Size = 11
      
      r_int_NroFil = 4
      
      For r_int_NoFlLi = 0 To grd_Listad(0).Rows - 1
      
         .Cells(r_int_NroFil, 2) = grd_Listad(0).TextMatrix(r_int_NoFlLi, 0)
         .Cells(r_int_NroFil, 3) = grd_Listad(0).TextMatrix(r_int_NoFlLi, 1)
         
         r_int_NroFil = r_int_NroFil + 1
      Next r_int_NoFlLi
      
      .Range(.Cells(4, 2), .Cells(5, 3)).Font.Bold = True
      .Range(.Cells(9, 2), .Cells(9, 3)).Font.Bold = True
      .Range(.Cells(15, 2), .Cells(16, 3)).Font.Bold = True
      .Range(.Cells(20, 2), .Cells(20, 3)).Font.Bold = True
      .Range(.Cells(22, 2), .Cells(23, 3)).Font.Bold = True
      .Range(.Cells(27, 2), .Cells(27, 3)).Font.Bold = True
      
      With .Range(.Cells(4, 2), .Cells(r_int_NroFil - 1, 3))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous
          .Borders(xlEdgeTop).LineStyle = xlContinuous
          .Borders(xlEdgeBottom).LineStyle = xlContinuous
          .Borders(xlEdgeRight).LineStyle = xlContinuous
          .Borders(xlInsideVertical).LineStyle = xlContinuous
          .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
    .Columns("B:B").EntireColumn.AutoFit
    .Columns("C:C").EntireColumn.AutoFit
         
   End With
   
   r_obj_Excel.Sheets(2).Name = "Garantía"
   
   With r_obj_Excel.Sheets(2)
      .Cells(2, 2) = "REPORTE DE ENTIDAD TECNICA - GARANTIA"
      .Range(.Cells(2, 1), .Cells(2, 4)).Merge
      .Range(.Cells(2, 1), .Cells(2, 4)).Font.Bold = True
      .Range(.Cells(2, 1), .Cells(2, 4)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(2, 1), .Cells(2, 4)).Font.Size = 14
      .Columns("C").HorizontalAlignment = xlHAlignLeft
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(3, 1), .Cells(99, 99)).Font.Size = 11
      
      r_int_NroFil = 4
      
      For r_int_NoFlLi = 0 To grd_Listad(1).Rows - 1
      
         .Cells(r_int_NroFil, 2) = grd_Listad(1).TextMatrix(r_int_NoFlLi, 0)
         .Cells(r_int_NroFil, 3) = grd_Listad(1).TextMatrix(r_int_NoFlLi, 1)
         
         r_int_NroFil = r_int_NroFil + 1
      Next r_int_NoFlLi
      
      .Range(.Cells(4, 2), .Cells(5, 3)).Font.Bold = True
      .Range(.Cells(9, 2), .Cells(9, 3)).Font.Bold = True
      .Range(.Cells(15, 2), .Cells(16, 3)).Font.Bold = True
      .Range(.Cells(20, 2), .Cells(20, 3)).Font.Bold = True
      .Range(.Cells(22, 2), .Cells(23, 3)).Font.Bold = True
      .Range(.Cells(27, 2), .Cells(27, 3)).Font.Bold = True
      
      With .Range(.Cells(4, 2), .Cells(r_int_NroFil - 1, 3))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous
          .Borders(xlEdgeTop).LineStyle = xlContinuous
          .Borders(xlEdgeBottom).LineStyle = xlContinuous
          .Borders(xlEdgeRight).LineStyle = xlContinuous
          .Borders(xlInsideVertical).LineStyle = xlContinuous
          .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
    .Columns("B:B").EntireColumn.AutoFit
    .Columns("C:C").EntireColumn.AutoFit
         
   End With
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
   
End Sub
