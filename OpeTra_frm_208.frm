VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_Con_SolHip_52 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9840
   ClientLeft      =   5370
   ClientTop       =   1935
   ClientWidth     =   11595
   Icon            =   "OpeTra_frm_208.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9840
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9825
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11595
      _Version        =   65536
      _ExtentX        =   20452
      _ExtentY        =   17330
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
         TabIndex        =   1
         Top             =   750
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_208.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "OpeTra_frm_208.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir de la Ventana"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
            Height          =   495
            Left            =   630
            TabIndex        =   4
            Top             =   60
            Width           =   7725
            _Version        =   65536
            _ExtentX        =   13626
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Consulta de Solicitud de Crédito Hipotecario"
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
            Picture         =   "OpeTra_frm_208.frx":0758
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel14 
         Height          =   4125
         Left            =   30
         TabIndex        =   5
         Top             =   1440
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   7276
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
         Begin MSFlexGridLib.MSFlexGrid grd_DatSol 
            Height          =   3735
            Left            =   60
            TabIndex        =   6
            Top             =   330
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   6588
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.Label Label2 
            Caption         =   "Datos Generales de Solicitud de Crédito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   7
            Top             =   60
            Width           =   3945
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   4155
         Left            =   30
         TabIndex        =   8
         Top             =   5610
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   7329
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisIns 
            Height          =   3495
            Left            =   60
            TabIndex        =   9
            Top             =   630
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   6165
            _Version        =   393216
            Rows            =   21
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel13 
            Height          =   285
            Left            =   5100
            TabIndex        =   10
            Top             =   330
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Fin Eval."
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   285
            Left            =   7590
            TabIndex        =   11
            Top             =   330
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3360
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Situación"
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
         Begin Threed.SSPanel SSPanel9 
            Height          =   285
            Left            =   3720
            TabIndex        =   12
            Top             =   330
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Inicio Eval."
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
         Begin Threed.SSPanel SSPanel12 
            Height          =   285
            Left            =   90
            TabIndex        =   13
            Top             =   330
            Width           =   3645
            _Version        =   65536
            _ExtentX        =   6429
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Instancia"
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   6480
            TabIndex        =   14
            Top             =   330
            Width           =   1125
            _Version        =   65536
            _ExtentX        =   1984
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Días Transc."
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
         Begin Threed.SSPanel SSPanel3 
            Height          =   285
            Left            =   9480
            TabIndex        =   17
            Top             =   330
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Usuario"
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
         Begin VB.Label Label1 
            Caption         =   "Seguimiento por Instancias"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   15
            Top             =   60
            Width           =   3165
         End
      End
   End
End
Attribute VB_Name = "frm_Con_SolHip_52"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_ExpExc_Click()
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
      
   Call fs_GenExc
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call modmip_gs_DatSolCre(moddat_g_str_NumSol, grd_DatSol) 'Call fs_Buscar_DatGen
   Call fs_Buscar_Seguim
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Grid de Datos de la Solicitud
   grd_DatSol.ColWidth(0) = 2600
   grd_DatSol.ColWidth(1) = 8470
   grd_DatSol.ColAlignment(0) = flexAlignLeftCenter
   grd_DatSol.ColAlignment(1) = flexAlignLeftCenter
   
   'Inicializando Grid de Instancias
   grd_LisIns.Cols = 8
   grd_LisIns.ColWidth(0) = 3600
   grd_LisIns.ColWidth(1) = 1385
   grd_LisIns.ColWidth(2) = 1385
   grd_LisIns.ColWidth(3) = 1115
   grd_LisIns.ColWidth(4) = 1900
   grd_LisIns.ColWidth(5) = 0
   grd_LisIns.ColWidth(6) = 0
   grd_LisIns.ColWidth(7) = 1650
   grd_LisIns.ColAlignment(0) = flexAlignLeftCenter
   grd_LisIns.ColAlignment(1) = flexAlignCenterCenter
   grd_LisIns.ColAlignment(2) = flexAlignCenterCenter
   grd_LisIns.ColAlignment(3) = flexAlignRightCenter
   grd_LisIns.ColAlignment(4) = flexAlignCenterCenter
   grd_LisIns.ColAlignment(7) = flexAlignCenterCenter
End Sub

Private Sub fs_Buscar_DatGen_Ant()
   Dim r_str_CodPry     As String
   Dim r_str_NomPry     As String
   Dim r_str_CodBco     As String
   
   Call gs_LimpiaGrid(grd_DatSol)
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If

   g_rst_Princi.MoveFirst
   
   'Cliente
   moddat_g_int_TipDoc = g_rst_Princi!SOLMAE_TITTDO
   moddat_g_str_NumDoc = Trim(g_rst_Princi!SOLMAE_TITNDO & "")
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   'Cónyuge
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
   moddat_g_str_CygNom = ""
   
   If g_rst_Princi!SOLMAE_CYGTDO > 0 Then
      moddat_g_int_CygTDo = g_rst_Princi!SOLMAE_CYGTDO
      moddat_g_str_CygNDo = Trim(g_rst_Princi!SOLMAE_CYGNDO & "")
      moddat_g_str_CygNom = moddat_gf_Buscar_NomCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo)
   End If
   
   'Producto
   moddat_g_str_CodPrd = g_rst_Princi!SOLMAE_CODPRD
   moddat_g_str_NomPrd = moddat_gf_Consulta_Produc(Trim(g_rst_Princi!SOLMAE_CODPRD))
   moddat_g_str_CodSub = g_rst_Princi!SOLMAE_CODSUB
   
   'Moneda
   moddat_g_int_TipMon = g_rst_Princi!SOLMAE_TIPMON
   moddat_g_str_Moneda = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLMAE_TIPMON))
   
   'Modalidad
   moddat_g_str_CodMod = ""
   moddat_g_str_DesMod = ""
   
   If Len(Trim(g_rst_Princi!SOLMAE_CODMOD & "")) > 0 Then
      moddat_g_str_CodMod = Trim(g_rst_Princi!SOLMAE_CODMOD & "")
      moddat_g_str_DesMod = moddat_gf_Buscar_NomMod(Trim(g_rst_Princi!SOLMAE_CODPRD), moddat_g_str_CodMod)
   End If
   
   'Ejecutivo de Seguimiento
   moddat_g_str_CodEjeSeg = Trim(g_rst_Princi!SOLMAE_EJESEG)
   moddat_g_str_NomEjeSeg = moddat_gf_Buscar_NomEje(Trim(g_rst_Princi!SOLMAE_EJESEG))
   
   'Consejero Hipotecario
   moddat_g_str_CodConHip = Trim(g_rst_Princi!SOLMAE_CONHIP)
   moddat_g_str_NomConHip = moddat_gf_Buscar_NomEje(Trim(g_rst_Princi!SOLMAE_CONHIP))
   
   'Fecha de Ingreso
   moddat_g_str_FecIng = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
   
   'Situación
   moddat_g_int_Situac = g_rst_Princi!SOLMAE_SITUAC
   moddat_g_str_Situac = moddat_gf_Consulta_ParDes("020", CStr(g_rst_Princi!SOLMAE_SITUAC))
   
   'Inmueble Identificado
   moddat_g_int_InmIde = g_rst_Princi!SOLMAE_INMIDE
   
   'Instancia Actual
   moddat_g_int_InsAct = g_rst_Princi!SOLMAE_CODINS
   
   'Obteniendo Información del Inmueble
   Call moddat_gs_Consulta_DatInm(moddat_g_str_NumSol, moddat_g_str_Direcc, moddat_g_str_Distri, r_str_CodPry, r_str_NomPry, r_str_CodBco)
   
   'Cargando en Grid
   grd_DatSol.Rows = grd_DatSol.Rows + 1:       grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                          grd_DatSol.Text = "Número de Solicitud"
   grd_DatSol.Col = 1:                          grd_DatSol.Text = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1:       grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                          grd_DatSol.Text = "Cliente"
   grd_DatSol.Col = 1:                          grd_DatSol.Text = CStr(g_rst_Princi!SOLMAE_TITTDO) & " - " & Trim(g_rst_Princi!SOLMAE_TITNDO) & " / " & moddat_g_str_NomCli
   
   If g_rst_Princi!SOLMAE_CYGTDO > 0 Then
      grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0:                       grd_DatSol.Text = "Cónyuge"
      grd_DatSol.Col = 1:                       grd_DatSol.Text = CStr(moddat_g_int_CygTDo) & " - " & Trim(moddat_g_str_CygNDo) & " / " & moddat_g_str_CygNom
   End If
   
   'Apoderado
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE DATGEN_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " AND DATGEN_NUMDOC = '" & Trim(moddat_g_str_NumDoc) & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera!DATGEN_APOTDO > 0 Then
      grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0:                       grd_DatSol.Text = "Apoderado"
      grd_DatSol.Col = 1:                       grd_DatSol.Text = CStr(g_rst_Genera!DATGEN_APOTDO) & " - " & Trim(g_rst_Genera!DATGEN_APONDO) & " / " & Trim(g_rst_Genera!DATGEN_APOAPP) & " " & Trim(g_rst_Genera!DATGEN_APOAPM) & " " & Trim(g_rst_Genera!DATGEN_APONOM)
   End If
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1:       grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                          grd_DatSol.Text = "Producto"
   grd_DatSol.Col = 1:                          grd_DatSol.Text = moddat_g_str_NomPrd & " / " & moddat_gf_Consulta_SubPrd(g_rst_Princi!SOLMAE_CODPRD, g_rst_Princi!SOLMAE_CODSUB)
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1:       grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                          grd_DatSol.Text = "Primera Vivienda"
   grd_DatSol.Col = 1:                          grd_DatSol.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!SOLMAE_PRIVIV))
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1:       grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                          grd_DatSol.Text = "Moneda Préstamo"
   grd_DatSol.Col = 1:                          grd_DatSol.Text = moddat_g_str_Moneda
   
   If Len(Trim(moddat_g_str_Direcc)) > 0 Then
      grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0:                       grd_DatSol.Text = "Modalidad"
      grd_DatSol.Col = 1:                       grd_DatSol.Text = moddat_g_str_DesMod
   
      grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0:                       grd_DatSol.Text = "Dirección Inmueble"
      grd_DatSol.Col = 1:                       grd_DatSol.Text = moddat_g_str_Direcc
      
      grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0:                       grd_DatSol.Text = "Distrito"
      grd_DatSol.Col = 1:                       grd_DatSol.Text = moddat_g_str_Distri
      
      If Len(Trim(r_str_CodPry)) > 0 Then
         grd_DatSol.Rows = grd_DatSol.Rows + 1: grd_DatSol.Row = grd_DatSol.Rows - 1
         grd_DatSol.Col = 0:                    grd_DatSol.Text = "Proyecto Inmobiliario"
         grd_DatSol.Col = 1:                    grd_DatSol.Text = moddat_gf_Consulta_NomPry(r_str_CodPry)
      ElseIf Len(Trim(r_str_NomPry)) > 0 Then
         grd_DatSol.Rows = grd_DatSol.Rows + 1: grd_DatSol.Row = grd_DatSol.Rows - 1
         grd_DatSol.Col = 0:                    grd_DatSol.Text = "Proyecto Inmobiliario"
         grd_DatSol.Col = 1:                    grd_DatSol.Text = r_str_NomPry & " (" & moddat_gf_Consulta_ParDes("513", r_str_CodBco) & ")"
      End If
   End If
   
   If g_rst_Princi!SOLMAE_COMVTA_SOL > 0 Then
      grd_DatSol.Rows = grd_DatSol.Rows + 2:    grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0:                       grd_DatSol.Text = "Valor Compra Venta"
      grd_DatSol.Col = 1:                       grd_DatSol.CellFontName = "Lucida Console"
      grd_DatSol.CellFontSize = 8:              grd_DatSol.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & IIf(moddat_g_int_TipMon = 1, gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_SOL, 12, 2), gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_DOL, 12, 2))
      
      grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0:                       grd_DatSol.Text = "Aporte Propio"
      grd_DatSol.Col = 1:                       grd_DatSol.CellFontName = "Lucida Console"
      
      'If moddat_g_str_CodPrd = "021" Or moddat_g_str_CodPrd = "022" Or moddat_g_str_CodPrd = "023" Then
      If InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 And moddat_g_str_CodPrd <> "019" Then
         grd_DatSol.CellFontSize = 8:              grd_DatSol.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & IIf(moddat_g_int_TipMon = 1, gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_SOL, 12, 2), gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_DOL, 12, 2)) & " (INCLUYE BBP " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & Format(g_rst_Princi!SOLMAE_FMVBBP, "###,###,##0.00") & ") "
      Else
         grd_DatSol.CellFontSize = 8:              grd_DatSol.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & IIf(moddat_g_int_TipMon = 1, gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_SOL, 12, 2), gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_DOL, 12, 2))
      End If
      
      grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0:                       grd_DatSol.Text = "Monto Préstamo"
      grd_DatSol.Col = 1:                       grd_DatSol.CellFontName = "Lucida Console"
      grd_DatSol.CellFontSize = 8:              grd_DatSol.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_MPR, 12, 2)
   End If
      
   grd_DatSol.Rows = grd_DatSol.Rows + 2:    grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                       grd_DatSol.Text = "Tasa de Interés"
   grd_DatSol.Col = 1:                       grd_DatSol.Text = Format(g_rst_Princi!SOLMAE_TASINT, "##0.00") & "%"
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                       grd_DatSol.Text = "Plazo"
   grd_DatSol.Col = 1:                       grd_DatSol.Text = CStr(g_rst_Princi!SOLMAE_PLAANO) & " Años"
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                       grd_DatSol.Text = "Número de Cuotas"
   grd_DatSol.Col = 1:                       grd_DatSol.Text = CStr(g_rst_Princi!SOLMAE_PLAANO * 12)
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                       grd_DatSol.Text = "Período de Gracia"
   grd_DatSol.Col = 1:                       grd_DatSol.Text = CStr(g_rst_Princi!SOLMAE_PERGRA) & " Meses"

   grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                       grd_DatSol.Text = "Cuotas Extraordinarias"
   grd_DatSol.Col = 1:                       grd_DatSol.Text = moddat_gf_Consulta_ParDes("277", CStr(g_rst_Princi!SOLMAE_CUOEXT))
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                       grd_DatSol.Text = "Compañía de Seguros"
   grd_DatSol.Col = 1:                       grd_DatSol.Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!SOLMAE_ESGDES & "")
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                       grd_DatSol.Text = "Tipo de Seguro Desgravamen"
   grd_DatSol.Col = 1:                       grd_DatSol.Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!SOLMAE_ESGDES, g_rst_Princi!SOLMAE_TIPSEG)
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                       grd_DatSol.Text = "Día de Pago"
   grd_DatSol.Col = 1:                       grd_DatSol.Text = Format(g_rst_Princi!SOLMAE_DIAPAG, "00")
   
   grd_DatSol.Rows = grd_DatSol.Rows + 2:    grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                       grd_DatSol.Text = "Situación"
   grd_DatSol.Col = 1:                       grd_DatSol.Text = moddat_g_str_Situac
   grd_DatSol.CellFontBold = True:           grd_DatSol.CellForeColor = modgen_g_con_ColNar
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                       grd_DatSol.Text = "Fecha de Ingreso"
   grd_DatSol.Col = 1:                       grd_DatSol.Text = moddat_g_str_FecIng
   
   grd_DatSol.Rows = grd_DatSol.Rows + 2:    grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                       grd_DatSol.Text = "Consejero Hipotecario"
   grd_DatSol.Col = 1:                       grd_DatSol.Text = moddat_g_str_NomConHip
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0:                       grd_DatSol.Text = "Ejecutivo Seguimiento"
   grd_DatSol.Col = 1:                       grd_DatSol.Text = moddat_g_str_NomEjeSeg
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1
   
   If g_rst_Genera!DATGEN_TDOVIN > 0 Then
      grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0:                       grd_DatSol.Text = "Vinculado"
      grd_DatSol.Col = 1
      
      If g_rst_Genera!DATGEN_TIPVIN = 1 Then
         grd_DatSol.Text = "TRABAJADOR"
      ElseIf g_rst_Genera!DATGEN_TIPVIN = 2 Or g_rst_Genera!DATGEN_TIPVIN = 3 Then
         grd_DatSol.Text = "VINCULADO A TRABAJADOR (" & modmip_gf_Consulta_NomTra(g_rst_Genera!DATGEN_TDOVIN, Trim(g_rst_Genera!DATGEN_NDOVIN)) & ")"
      ElseIf g_rst_Genera!DATGEN_TIPVIN = 4 Then
         grd_DatSol.Text = "FUNCIONARIO"
      ElseIf g_rst_Genera!DATGEN_TIPVIN = 5 Then
         grd_DatSol.Text = "VINCULADO A FUNCIONARIO (" & modmip_gf_Consulta_NomOtrFun(g_rst_Genera!DATGEN_TDOVIN, Trim(g_rst_Genera!DATGEN_NDOVIN)) & ")"
      Else
         grd_DatSol.Text = ""
      End If
   End If
   
   If g_rst_Genera!DATGEN_TDOACC > 0 Then
      grd_DatSol.Rows = grd_DatSol.Rows + 1:    grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0:                       grd_DatSol.Text = "Accionista"
      grd_DatSol.Col = 1
      
      If g_rst_Genera!DATGEN_ACCVIN = 1 Then
         grd_DatSol.Text = "ACCIONISTA"
      ElseIf g_rst_Genera!DATGEN_ACCVIN = 2 Then
         grd_DatSol.Text = "VINCULADO A ACCIONISTA (" & modmip_gf_Consulta_NomAcc(g_rst_Genera!DATGEN_TDOACC, Trim(g_rst_Genera!DATGEN_NDOACC)) & ")"
      End If
   End If
   
   modmip_g_int_PaiRes = g_rst_Genera!DATGEN_PAIRES
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_DatSol)
End Sub

Private Sub grd_DatSol_SelChange()
   If grd_DatSol.Rows > 2 Then
      grd_DatSol.RowSel = grd_DatSol.Row
   End If
End Sub

Private Sub grd_LisIns_DblClick()
   Dim r_int_Situac     As Integer

   If grd_LisIns.Rows = 0 Then
      Exit Sub
   End If
   
   grd_LisIns.Col = 5
   moddat_g_int_InsAct = CInt(grd_LisIns.Text)
   grd_LisIns.Col = 6
   r_int_Situac = CInt(grd_LisIns.Text)

   Call gs_RefrescaGrid(grd_LisIns)
   moddat_g_int_FlgAct = 1
   
   Select Case moddat_g_int_InsAct
      Case 11: frm_Con_SolHip_53.Show 1
      Case 21: frm_Con_SolHip_54.Show 1
      Case 31: frm_Con_SolHip_55.Show 1
      Case 32: frm_Con_SolHip_56.Show 1
      Case 41: frm_Con_SolHip_57.Show 1
      Case 42: frm_Con_SolHip_58.Show 1
      Case 51: frm_Con_SolHip_59.Show 1
      Case 61: frm_Con_SolHip_60.Show 1
      Case 62
         If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then  '"003" "004" "007" "009" "010" "012" "013" "014" "015" "016" "017" "018" Or moddat_g_str_CodPrd = "019" Or moddat_g_str_CodPrd = "020" Or moddat_g_str_CodPrd = "021" Or moddat_g_str_CodPrd = "022" Or moddat_g_str_CodPrd = "023" Then
            frm_Con_SolHip_61.Show 1
         Else
            frm_Con_SolHip_62.Show 1
         End If
      Case 72: frm_Con_SolHip_63.Show 1
      Case 81: frm_Con_SolHip_64.Show 1
      Case 91: frm_Con_SolHip_65.Show 1
   End Select
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call modmip_gs_DatSolCre(moddat_g_str_NumSol, grd_DatSol) 'Call fs_Buscar_DatGen
      Call fs_Buscar_Seguim
      Screen.MousePointer = 0
   End If
End Sub

Private Sub grd_LisIns_SelChange()
   If grd_LisIns.Rows > 2 Then
      grd_LisIns.RowSel = grd_LisIns.Row
   End If
End Sub

Private Sub fs_Buscar_Seguim()
   Dim r_int_DiaTra     As Integer
   Dim r_int_DiaTas     As Integer
   Dim r_int_DiaSeg     As Integer
   Dim r_int_DiaPol     As Integer
   Dim r_int_DiaMVi     As Integer
   
   Call gs_LimpiaGrid(grd_LisIns)
   r_int_DiaTra = 0
   r_int_DiaTas = 0
   r_int_DiaSeg = 0
   r_int_DiaPol = 0
   r_int_DiaMVi = 0
      
   g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE SEGUIM_NUMSOL = '" & moddat_g_str_NumSol & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   g_rst_Princi.MoveFirst
   grd_LisIns.Redraw = False
   
   Do While Not g_rst_Princi.EOF
      grd_LisIns.Rows = grd_LisIns.Rows + 1
      grd_LisIns.Row = grd_LisIns.Rows - 1
      
      'Instancia
      grd_LisIns.Col = 0
      grd_LisIns.Text = moddat_gf_Consulta_ParDes("002", Format(g_rst_Princi!SEGUIM_CODINS, "000000"))
      
      grd_LisIns.Col = 5
      grd_LisIns.Text = g_rst_Princi!SEGUIM_CODINS
      
      'Fecha de Inicio
      grd_LisIns.Col = 1
      grd_LisIns.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECINI))
      
      'Fecha de Fin
      grd_LisIns.Col = 2
      If g_rst_Princi!SEGUIM_FECFIN > 0 Then
         grd_LisIns.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECFIN))
         
         'Días Transcurridos
         grd_LisIns.Col = 3
         grd_LisIns.Text = CStr(g_rst_Princi!SEGUIM_DIATRA)
         
         If g_rst_Princi!SEGUIM_CODINS = 41 Or g_rst_Princi!SEGUIM_CODINS = 42 Then
            If g_rst_Princi!SEGUIM_CODINS = 41 Then
               r_int_DiaTas = g_rst_Princi!SEGUIM_DIATRA
            Else
               r_int_DiaSeg = g_rst_Princi!SEGUIM_DIATRA
            End If
            
            If g_rst_Princi!SEGUIM_CODINS = 42 Then
               If r_int_DiaTas > r_int_DiaSeg Then
                  r_int_DiaTra = r_int_DiaTra + r_int_DiaTas
               Else
                  r_int_DiaTra = r_int_DiaTra + r_int_DiaSeg
               End If
            End If
         ElseIf g_rst_Princi!SEGUIM_CODINS = 61 Or g_rst_Princi!SEGUIM_CODINS = 62 Then
            If g_rst_Princi!SEGUIM_CODINS = 61 Then
               r_int_DiaPol = g_rst_Princi!SEGUIM_DIATRA
            Else
               r_int_DiaMVi = g_rst_Princi!SEGUIM_DIATRA
            End If
            
            If g_rst_Princi!SEGUIM_CODINS = 62 Or (g_rst_Princi!SEGUIM_CODINS = 61 And moddat_g_str_CodPrd = "002") Then
               If r_int_DiaPol > r_int_DiaMVi Then
                  r_int_DiaTra = r_int_DiaTra + r_int_DiaPol
               Else
                  r_int_DiaTra = r_int_DiaTra + r_int_DiaMVi
               End If
            End If
         Else
            r_int_DiaTra = r_int_DiaTra + g_rst_Princi!SEGUIM_DIATRA
         End If
      Else
         If moddat_g_int_Situac = 1 Then
            r_int_DiaTra = r_int_DiaTra + CInt(date - CDate(gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECINI))))
         Else
            r_int_DiaTra = r_int_DiaTra + CInt(date - CDate(gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECINI))))
         End If
      End If
      
      'Situación
      grd_LisIns.Col = 4
      grd_LisIns.Text = moddat_gf_Consulta_ParDes("023", CStr(g_rst_Princi!SEGUIM_SITUAC))
      
      grd_LisIns.Col = 6
      grd_LisIns.Text = CStr(g_rst_Princi!SEGUIM_SITUAC)
      
      grd_LisIns.Col = 7
      If IsNull(g_rst_Princi!SEGUSUACT) Then
         grd_LisIns.Text = ""
      Else
         grd_LisIns.Text = Trim(g_rst_Princi!SEGUSUACT)
      End If

      g_rst_Princi.MoveNext
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   grd_LisIns.Redraw = True
   Call gs_UbiIniGrid(grd_LisIns)
End Sub

Private Function ff_ObsRec(ByVal p_NumSol As String, ByVal p_TipRec As Integer) As String
   ff_ObsRec = " "
   
   If p_TipRec = 1 Then
      g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
      g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & p_NumSol & "' AND "
      g_str_Parame = g_str_Parame & "SEGDET_CODOCU = 13 "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Function
      End If
   
      DoEvents
      g_rst_Genera.MoveFirst
      ff_ObsRec = Trim(g_rst_Genera!SEGDET_OBSERV & "")
   
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   ElseIf p_TipRec = 3 Then
      g_str_Parame = "SELECT * FROM TRA_RECADM WHERE "
      g_str_Parame = g_str_Parame & "RECADM_NUMSOL = '" & p_NumSol & "' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Function
      End If
   
      DoEvents
      g_rst_Genera.MoveFirst
      ff_ObsRec = Trim(g_rst_Genera!RECADM_OBSERV & "")
   
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
End Function

Private Function ff_ObsExp(ByVal p_NumSol As String, ByVal p_CodIns As Integer) As String
   ff_ObsExp = " "
      
   g_str_Parame = "SELECT SEGDET_OBSERV "
   g_str_Parame = g_str_Parame & "FROM (SELECT RANK() OVER (ORDER BY SEGFECCRE DESC, SEGHORCRE DESC) R,SEGDET_OBSERV FROM TRA_SEGDET"
   g_str_Parame = g_str_Parame & " WHERE  SEGDET_NUMSOL='" & p_NumSol & "' AND SEGDET_CODINS='" & p_CodIns & "' )"
   g_str_Parame = g_str_Parame & " WHERE R < 2"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If

   DoEvents
   g_rst_Genera.MoveFirst
   ff_ObsExp = Trim(g_rst_Genera!SEGDET_OBSERV & "")

   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Private Sub fs_GenExc()
   Dim r_str_nrohij As String
   Dim r_str_clinom As String
   Dim r_str_edahij As String
   Dim r_str_comvta As String
   Dim r_str_apopro As String
   Dim r_str_mtopre As String
   Dim r_str_nompfs As String
   Dim r_str_nomcar As String
   Dim r_str_fecing As String
   Dim r_str_cuoapr As String
   Dim r_str_cuomen As String
   Dim r_str_ingliq As String
   Dim r_str_obs As String
   
   'r_str_CodPry
   Dim r_obj_Excel      As Excel.Application
      
   r_str_nrohij = ""
   r_str_clinom = ""
   r_str_edahij = ""
   r_str_comvta = ""
   r_str_apopro = ""
   r_str_mtopre = ""
   r_str_nompfs = ""
   r_str_nomcar = ""
   r_str_fecing = ""
   r_str_obs = ""
   r_str_cuoapr = ""
   r_str_ingliq = ""
   r_str_cuomen = ""
   
   g_str_Parame = "SELECT TRIM(EJECMC_NOMBRE) AS CNJNOM ,TRIM(EJECMC_APEPAT) AS CNJPAT,TRIM(EJECMC_APEMAT) AS CNJMAT  "
   g_str_Parame = g_str_Parame & ",TRIM(DATGEN_APEPAT) AS APEPAT, TRIM(DATGEN_APEMAT) AS APEMAT "
   g_str_Parame = g_str_Parame & ",TRIM(DATGEN_NOMBRE) AS CLINOM, TRIM(DATGEN_APECAS) AS APECAS "
   g_str_Parame = g_str_Parame & ",trunc((SYSDATE - to_date(DATGEN_NACFEC, 'YYYY/MM/DD'))/365,0) AS CLIEDA"
   g_str_Parame = g_str_Parame & ",TRIM(EC.PARDES_DESCRI) AS ESTCIV ,TRIM(NE.PARDES_DESCRI) AS NIVELESTUDIO "
   g_str_Parame = g_str_Parame & ",DATGEN_PROFES, DATGEN_DEPECO,TRIM(OP.PARDES_DESCRI) AS NOMOCP "
   g_str_Parame = g_str_Parame & ",DATGEN_EDAD01,DATGEN_EDAD02,DATGEN_EDAD03,DATGEN_EDAD04,DATGEN_EDAD05 "
   g_str_Parame = g_str_Parame & ",TRIM(ACTECO_DEP_NOMCOM) AS NOMCOM, ACTECO_DEP_FECING,ACTECO_DEP_CODCAR, ACTECO_DEP_NOMCAR "
   g_str_Parame = g_str_Parame & ",SOLINM_PRYMCS,PRY.DATGEN_TITULO AS SOLINM_PRYNOM,SOLINM_RAZSOC_PRO,TRIM(EVA.PARDES_DESCRI) AS TIPEVA "
   g_str_Parame = g_str_Parame & ",SOLMAE_TIPMON, SOLMAE_COMVTA_SOL,SOLMAE_COMVTA_DOL,SOLMAE_APOPRO_SOL,SOLMAE_APOPRO_DOL "
   g_str_Parame = g_str_Parame & ",SOLMAE_MTOPRE_SOL, SOLMAE_MTOPRE_DOL,SOLMAE_PLAANO,SOLMAE_PLAANO,SOLMAE_CODINS  "
   g_str_Parame = g_str_Parame & ",SOLMAE_NUMERO, EVACRE_CUOMPR, EVACRE_INGNET,OP.PARDES_CODITE AS CODITE, EVACRE_CUOSOL "
   g_str_Parame = g_str_Parame & "FROM CLI_DATGEN "
   g_str_Parame = g_str_Parame & "LEFT JOIN  MNT_PARDES EC  ON (DATGEN_ESTCIV = EC.PARDES_CODITE AND  EC.PARDES_CODGRP = 205 ) "
   g_str_Parame = g_str_Parame & "LEFT JOIN  MNT_PARDES NE  ON (DATGEN_NIVEST = NE.PARDES_CODITE AND  NE.PARDES_CODGRP = 209 ) "
   g_str_Parame = g_str_Parame & "LEFT JOIN  MNT_PARDES OP  ON (DATGEN_OCUPAC = OP.PARDES_CODITE AND  OP.PARDES_CODGRP = 008 ) "
   g_str_Parame = g_str_Parame & "JOIN       CLI_ACTECO     ON (DATGEN_NUMDOC = ACTECO_CLINDO) "
   g_str_Parame = g_str_Parame & "JOIN       CRE_SOLMAE     ON (SOLMAE_TITNDO = DATGEN_NUMDOC AND SOLMAE_TITTDO = DATGEN_TIPDOC ) "
   g_str_Parame = g_str_Parame & "LEFT JOIN  MNT_PARDES EVA ON (SOLMAE_TIPEVA = EVA.PARDES_CODITE AND EVA.PARDES_CODGRP = 038 ) "
   g_str_Parame = g_str_Parame & "JOIN       CRE_EJECMC     ON (SOLMAE_CONHIP =EJECMC_CODEJE) "
   g_str_Parame = g_str_Parame & "LEFT JOIN  CRE_SOLINM     ON (SOLINM_NUMSOL= SOLMAE_NUMERO) "
   g_str_Parame = g_str_Parame & "LEFT JOIN  TRA_EVACRE     ON (EVACRE_NUMSOL= SOLMAE_NUMERO ) "
   g_str_Parame = g_str_Parame & "LEFT JOIN  PRY_DATGEN PRY ON (SOLINM_PRYCOD = PRY.DATGEN_CODIGO) "
   g_str_Parame = g_str_Parame & "WHERE SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "'"
          
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron datos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
            
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
     
   With r_obj_Excel.ActiveSheet
      'Unir celdas
      .Range("B7") = "FICHA PARA REVISION - EXPEDIENTE COMITÉ DE CREDITOS"
      .Range("B7:D6").Font.Underline = True
      .Range("B7:F7").Merge
      
      .Range("B42:F42").Merge
      .Range("B49:F49").Merge
      .Range("B56:F56").Merge
      .Range("B66:F66").Merge
      .Range("B76:F76").Merge
      .Range("B58:F64").Merge
      
      'Bordes de las celdas
      .Range("B42:F42").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("B42:F42").Borders(xlEdgeLeft).Weight = xlThin
      .Range("B42:F42").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("B42:F42").Borders(xlEdgeTop).Weight = xlThin
      .Range("B42:F42").Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range("B42:F42").Borders(xlEdgeRight).Weight = xlThin
      .Range("B42:F42").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range("B42:F42").Borders(xlEdgeBottom).Weight = xlThin
      
      .Range("B49:F49").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("B49:F49").Borders(xlEdgeLeft).Weight = xlThin
      .Range("B49:F49").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("B49:F49").Borders(xlEdgeTop).Weight = xlThin
      .Range("B49:F49").Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range("B49:F49").Borders(xlEdgeRight).Weight = xlThin
      .Range("B49:F49").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range("B49:F49").Borders(xlEdgeBottom).Weight = xlThin
      
      .Range("B56:F56").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("B56:F56").Borders(xlEdgeLeft).Weight = xlThin
      .Range("B56:F56").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("B56:F56").Borders(xlEdgeTop).Weight = xlThin
      .Range("B56:F56").Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range("B56:F56").Borders(xlEdgeRight).Weight = xlThin
      .Range("B56:F56").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range("B56:F56").Borders(xlEdgeBottom).Weight = xlThin
       
      .Range("B66:F66").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("B66:F66").Borders(xlEdgeLeft).Weight = xlThin
      .Range("B66:F66").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("B66:F66").Borders(xlEdgeTop).Weight = xlThin
      .Range("B66:F66").Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range("B66:F66").Borders(xlEdgeRight).Weight = xlThin
      .Range("B66:F66").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range("B66:F66").Borders(xlEdgeBottom).Weight = xlThin
      
      .Range("B76:F76").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("B76:F76").Borders(xlEdgeLeft).Weight = xlThin
      .Range("B76:F76").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("B76:F76").Borders(xlEdgeTop).Weight = xlThin
      .Range("B76:F76").Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range("B76:F76").Borders(xlEdgeRight).Weight = xlThin
      .Range("B76:F76").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range("B76:F76").Borders(xlEdgeBottom).Weight = xlThin
      
      .Range("B58:F64").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("B58:F64").Borders(xlEdgeLeft).Weight = xlThin
      .Range("B58:F64").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("B58:F64").Borders(xlEdgeTop).Weight = xlThin
      .Range("B58:F64").Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range("B58:F64").Borders(xlEdgeRight).Weight = xlThin
      .Range("B58:F64").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range("B58:F64").Borders(xlEdgeBottom).Weight = xlThin
      
      .Range("B68:F74").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("B68:F74").Borders(xlEdgeLeft).Weight = xlThin
      .Range("B68:F74").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("B68:F74").Borders(xlEdgeTop).Weight = xlThin
      .Range("B68:F74").Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range("B68:F74").Borders(xlEdgeRight).Weight = xlThin
      .Range("B68:F74").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range("B68:F74").Borders(xlEdgeBottom).Weight = xlThin
      
      .Range("B78:F84").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("B78:F84").Borders(xlEdgeLeft).Weight = xlThin
      .Range("B78:F84").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("B78:F84").Borders(xlEdgeTop).Weight = xlThin
      .Range("B78:F84").Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range("B78:F84").Borders(xlEdgeRight).Weight = xlThin
      .Range("B78:F84").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range("B78:F84").Borders(xlEdgeBottom).Weight = xlThin
      
      .Cells(2, 6) = "Dpto. Tecnología e Informática"
      .Cells(3, 6) = "Desarrollo de Sistemas"
      .Cells(11, 2) = "CONSEJERO:"
      .Cells(13, 2) = "CLIENTE:"
      .Cells(13, 4) = "NRO DE SOLICITUD:"
      .Cells(15, 2) = "EDAD:"
      .Cells(17, 2) = "ESTADO CIVIL:"
      .Cells(17, 4) = "NIVEL DE ESTUDIOS:"
      .Cells(19, 2) = "NRO DE DEPENDIENTES:"
      .Cells(19, 4) = "EDADES DE LOS HIJOS:"
      .Cells(21, 2) = "PROFESION:"
      .Cells(21, 4) = "OCUPACION PRINCIPAL:"
      .Cells(23, 2) = "EMPLEADOR:"
      .Cells(25, 2) = "FECHA DE INGRESO:"
      .Cells(25, 4) = "CARGO:"
      .Cells(27, 2) = "PROYECTO MICASITA:"
      .Cells(27, 4) = "NOMBRE DEL PROYECTO:"
      .Cells(29, 2) = "RAZON SOCIAL PROMOTOR:"
      .Cells(31, 2) = "TIPO DE EVALUACION:"
      .Cells(33, 2) = "VALOR DE COMPRAVENTA:"
      .Cells(35, 2) = "APORTE PROPIO:"
      .Cells(37, 2) = "MONTO PRESTAMO:"
      .Cells(39, 2) = "PLAZO:"
      
      .Cells(42, 2) = "DATOS DE LA SIMULACION"
      .Cells(44, 2) = "CUOTA MENSUAL:"
      .Cells(44, 4) = "CUOTA CON PBP:"
      .Cells(46, 2) = "INGRESO REQUERIDO:"
      
      .Cells(49, 2) = "DATOS DE LA EVALUACION CREDITICIA"
      .Cells(51, 2) = "CUOT. MEN. APROB.: (ORG)"
      .Cells(51, 4) = "CUOT. MEN. APROB.: (SOL)"
      .Cells(53, 2) = "TOT. INGRESO LIQ. NETO:"
      
      .Cells(56, 2) = "EVALUACION CREDITICIA"
      .Cells(66, 2) = "SUSTENTO DE REEVALUACION"
      .Cells(76, 2) = "NUEVO REQUERIMIENTO"
           
      .Columns("A").ColumnWidth = 1.43
      .Columns("B").ColumnWidth = 21.3
      .Columns("C").ColumnWidth = 33.86
      .Columns("D").ColumnWidth = 20.29
      .Columns("E").ColumnWidth = 19.29
      
      .Rows("1:65536").RowHeight = 11.25
      
      .Cells(42, 2).Font.Bold = True
      .Cells(49, 2).Font.Bold = True
      .Cells(56, 2).Font.Bold = True
      .Cells(66, 2).Font.Bold = True
      .Cells(76, 2).Font.Bold = True
      .Columns("B").Font.Bold = True
      .Columns("D").Font.Bold = True
      .Range("B7").Font.Bold = True
            
      .Columns("B").HorizontalAlignment = xlHAlignLeft
      .Columns("C").HorizontalAlignment = xlHAlignLeft
      .Columns("D").HorizontalAlignment = xlHAlignLeft
      .Columns("E").HorizontalAlignment = xlHAlignLeft
      .Range("B7").HorizontalAlignment = xlHAlignCenter
      .Range("B58:F64").HorizontalAlignment = xlHAlignLeft
      .Cells(2, 6).HorizontalAlignment = xlHAlignRight
      .Cells(3, 6).HorizontalAlignment = xlHAlignRight
                                    
      .Range("B58:F64").VerticalAlignment = xlVAlignTop
      .Range("B58:F64").WrapText = True
                                    
      .Range("A1:F85").Font.Name = "Arial"
      .Range("A1:F85").Font.Size = 8
   End With
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      'NOMBRE DEL CLIENTE
      'r_str_clinom = IIf(IsNull(g_rst_Princi!APECAS) = True, g_rst_Princi!APEPAT & " " & g_rst_Princi!APEMAT & " " & g_rst_Princi!CLINOM, g_rst_Princi!APEPAT & " " & g_rst_Princi!APECAS & " " & g_rst_Princi!APEMAT & " " & g_rst_Princi!CLINOM)
      If IsNull(g_rst_Princi!APECAS) Then
         r_str_clinom = Trim(g_rst_Princi!APEPAT) & " " & Trim(g_rst_Princi!APEMAT) & " " & Trim(g_rst_Princi!CLINOM)
      Else
         If Len(Trim(g_rst_Princi!APECAS)) = 0 Then
            r_str_clinom = Trim(g_rst_Princi!APEPAT) & " " & Trim(g_rst_Princi!APEMAT) & " " & Trim(g_rst_Princi!CLINOM)
         Else
            r_str_clinom = Trim(g_rst_Princi!APEPAT) & " " & Trim(g_rst_Princi!APEMAT) & " DE " & Trim(g_rst_Princi!APECAS) & " " & Trim(g_rst_Princi!CLINOM)
         End If
      End If
      
      'OBSERVACIONES
      r_str_obs = ff_ObsExp(g_rst_Princi!SOLMAE_NUMERO, g_rst_Princi!SOLMAE_CODINS)
    
      'NRO DE DEPENDIENTES
      If g_rst_Princi!DatGen_DepEco <> 0 Then
        r_str_nrohij = CStr(g_rst_Princi!DatGen_DepEco)
      End If
      
      'EDADES DE LOS HIJOS
      Select Case g_rst_Princi!DatGen_DepEco
         Case 1
            r_str_edahij = CStr(g_rst_Princi!DatGen_EDAD01)
            r_str_edahij = IIf(g_rst_Princi!DatGen_EDAD01 = 1, r_str_edahij & " AÑO", _
            r_str_edahij & " AÑOS")
            r_obj_Excel.ActiveSheet.Cells(19, 4) = "EDAD DEL HIJO:"
         Case 2
            r_str_edahij = CStr(g_rst_Princi!DatGen_EDAD01) & " Y " & CStr(g_rst_Princi!DatGen_EDAD02) & " AÑOS"
         Case 3
            r_str_edahij = CStr(g_rst_Princi!DatGen_EDAD01) & ", " & CStr(g_rst_Princi!DatGen_EDAD02) & " Y " & _
            CStr(g_rst_Princi!DatGen_EDAD03) & " AÑOS"
         Case 4
            r_str_edahij = CStr(g_rst_Princi!DatGen_EDAD01) & ", " & CStr(g_rst_Princi!DatGen_EDAD02) & ", " & _
            CStr(g_rst_Princi!DatGen_EDAD03) & " Y " & CStr(g_rst_Princi!DatGen_EDAD04) & " AÑOS"
         Case 5
            r_str_edahij = CStr(g_rst_Princi!DatGen_EDAD01) & ", " & CStr(g_rst_Princi!DatGen_EDAD02) & ", " & _
            CStr(g_rst_Princi!DatGen_EDAD03) & ", " & CStr(g_rst_Princi!DatGen_EDAD04) & " Y " & _
            CStr(g_rst_Princi!DatGen_EDAD05) & " AÑOS"
      End Select
      
      Select Case g_rst_Princi!SOLMAE_TIPMON
         Case 1 'SOLES
            'VALOR DE COMPRAVENTA
            r_str_comvta = IIf(IsNull(g_rst_Princi!SOLMAE_COMVTA_SOL) = True, "", "S/.     " & CStr(Format(g_rst_Princi!SOLMAE_COMVTA_SOL, "###,###,##0.00")))
            'APORTE PROPIO
            r_str_apopro = IIf(IsNull(g_rst_Princi!SOLMAE_APOPRO_SOL) = True, "", "S/.     " & CStr(Format(g_rst_Princi!SOLMAE_APOPRO_SOL, "###,###,##0.00")))
            'MONTRO PRESTAMO
            r_str_mtopre = IIf(IsNull(g_rst_Princi!SOLMAE_MTOPRE_SOL) = True, "", "S/.     " & CStr(Format(g_rst_Princi!SOLMAE_MTOPRE_SOL, "###,###,##0.00")))
            'CUOTA MENSUAL APROBADA
            r_str_cuoapr = IIf(IsNull(g_rst_Princi!EVACRE_CUOMPR) = True, "", "S/.     " & CStr(Format(g_rst_Princi!EVACRE_CUOMPR, "###,###,##0.00")))
         Case 2 'DOLARES
            r_str_comvta = IIf(IsNull(g_rst_Princi!SOLMAE_COMVTA_DOL) = True, "", "US$     " & CStr(Format(g_rst_Princi!SOLMAE_COMVTA_DOL, "###,###,##0.00")))
            r_str_apopro = IIf(IsNull(g_rst_Princi!SOLMAE_APOPRO_DOL) = True, "", "US$     " & CStr(Format(g_rst_Princi!SOLMAE_APOPRO_DOL, "###,###,##0.00")))
            r_str_mtopre = IIf(IsNull(g_rst_Princi!SOLMAE_MTOPRE_DOL) = True, "", "US$     " & CStr(Format(g_rst_Princi!SOLMAE_MTOPRE_DOL, "###,###,##0.00")))
            r_str_cuoapr = IIf(IsNull(g_rst_Princi!EVACRE_CUOMPR) = True, "", "US$     " & CStr(Format(g_rst_Princi!EVACRE_CUOMPR, "###,###,##0.00")))
      End Select
      
      'INGRESO NETO LIQUIDO
      r_str_cuomen = IIf(IsNull(g_rst_Princi!EVACRE_CUOSOL) = True, "", "S/.     " & CStr(Format(g_rst_Princi!EVACRE_CUOSOL, "###,###,##0.00")))
      r_str_ingliq = IIf(IsNull(g_rst_Princi!EVACRE_INGNET) = True, "", "S/.     " & CStr(Format(g_rst_Princi!EVACRE_INGNET, "###,###,##0.00")))
      
      'PROFESION
      r_str_nompfs = moddat_gf_Consulta_ParDes("501", CStr(g_rst_Princi!DatGen_Profes))
      If g_rst_Princi!DatGen_Profes = "000099" Then
         r_str_nompfs = " "
      End If
      
      'CARGO
      If IsNull(g_rst_Princi!ActEco_Dep_CodCar) = False Then
         r_str_nomcar = IIf(g_rst_Princi!ActEco_Dep_CodCar = "999999", Trim(g_rst_Princi!ActEco_Dep_NomCar & ""), moddat_gf_Consulta_ParDes("503", g_rst_Princi!ActEco_Dep_CodCar))
      End If
      
      'FECHA DE INGRESO
      If g_rst_Princi!ActEco_Dep_FecIng <> 0 Then
         r_str_fecing = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecIng))
      End If
      
      r_obj_Excel.ActiveSheet.Cells(11, 3) = g_rst_Princi!CNJPAT & " " & g_rst_Princi!CNJMAT & " " & g_rst_Princi!CNJNOM
      r_obj_Excel.ActiveSheet.Cells(13, 3) = r_str_clinom
      r_obj_Excel.ActiveSheet.Cells(15, 3) = g_rst_Princi!CLIEDA & " AÑOS"
      r_obj_Excel.ActiveSheet.Cells(17, 3) = g_rst_Princi!ESTCIV
      r_obj_Excel.ActiveSheet.Cells(13, 5) = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO) ' NRO DE SOLICITUD
      r_obj_Excel.ActiveSheet.Cells(17, 5) = g_rst_Princi!NIVELESTUDIO
      r_obj_Excel.ActiveSheet.Cells(19, 3) = r_str_nrohij
      r_obj_Excel.ActiveSheet.Cells(19, 5) = r_str_edahij
      r_obj_Excel.ActiveSheet.Cells(21, 3) = r_str_nompfs
      r_obj_Excel.ActiveSheet.Cells(21, 5) = g_rst_Princi!NOMOCP
      r_obj_Excel.ActiveSheet.Cells(23, 3) = Trim(g_rst_Princi!NOMCOM)
      r_obj_Excel.ActiveSheet.Cells(25, 5) = Trim(r_str_nomcar)
      r_obj_Excel.ActiveSheet.Cells(25, 3) = r_str_fecing
      r_obj_Excel.ActiveSheet.Cells(27, 3) = IIf(g_rst_Princi!SOLINM_PRYMCS = 1, "SI", "NO")
      r_obj_Excel.ActiveSheet.Cells(27, 5) = Trim(g_rst_Princi!SOLINM_PRYNOM)
      r_obj_Excel.ActiveSheet.Cells(29, 3) = Trim(g_rst_Princi!SOLINM_RAZSOC_PRO)
      r_obj_Excel.ActiveSheet.Cells(31, 3) = g_rst_Princi!TIPEVA
      r_obj_Excel.ActiveSheet.Cells(33, 3) = CStr(r_str_comvta)
      r_obj_Excel.ActiveSheet.Cells(35, 3) = CStr(r_str_apopro)
      r_obj_Excel.ActiveSheet.Cells(37, 3) = CStr(r_str_mtopre)
      r_obj_Excel.ActiveSheet.Cells(39, 3) = g_rst_Princi!SOLMAE_PLAANO & " AÑOS"
      r_obj_Excel.ActiveSheet.Cells(51, 3) = r_str_cuoapr
      r_obj_Excel.ActiveSheet.Cells(51, 5) = r_str_cuomen
      r_obj_Excel.ActiveSheet.Cells(53, 3) = r_str_ingliq
      r_obj_Excel.ActiveSheet.Cells(58, 2) = Trim(r_str_obs)
      g_rst_Princi.MoveNext
      DoEvents
   Loop
  
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
     
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

