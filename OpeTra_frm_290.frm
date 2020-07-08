VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Pro_EvaPBP_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   8940
   ClientLeft      =   3435
   ClientTop       =   2160
   ClientWidth     =   14790
   Icon            =   "OpeTra_frm_290.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   14790
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8955
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14805
      _Version        =   65536
      _ExtentX        =   26114
      _ExtentY        =   15796
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
         Height          =   6975
         Left            =   30
         TabIndex        =   1
         Top             =   1920
         Width           =   14715
         _Version        =   65536
         _ExtentX        =   25956
         _ExtentY        =   12303
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
         Begin Threed.SSPanel pnl_Tit_NumOpe 
            Height          =   555
            Left            =   60
            TabIndex        =   2
            Top             =   60
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "Nro. Operación"
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
            Height          =   555
            Left            =   1680
            TabIndex        =   3
            Top             =   60
            Width           =   3975
            _Version        =   65536
            _ExtentX        =   7011
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "Cliente"
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
         Begin Threed.SSPanel pnl_Tit_FlgPBP 
            Height          =   555
            Left            =   8310
            TabIndex        =   4
            Top             =   60
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "Flag PBP"
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
         Begin Threed.SSPanel pnl_Tit_NumCuo 
            Height          =   555
            Left            =   9480
            TabIndex        =   5
            Top             =   60
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "Nro. Cuota TC"
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
            Height          =   6315
            Left            =   30
            TabIndex        =   6
            Top             =   630
            Width           =   14625
            _ExtentX        =   25797
            _ExtentY        =   11139
            _Version        =   393216
            Rows            =   30
            Cols            =   10
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_PenIni 
            Height          =   285
            Left            =   12390
            TabIndex        =   12
            Top             =   330
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Inicio"
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
         Begin Threed.SSPanel pnl_Tit_PenFin 
            Height          =   285
            Left            =   13350
            TabIndex        =   13
            Top             =   330
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fin"
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
         Begin Threed.SSPanel pnl_Tit_Produc 
            Height          =   555
            Left            =   5640
            TabIndex        =   18
            Top             =   60
            Width           =   2685
            _Version        =   65536
            _ExtentX        =   4736
            _ExtentY        =   979
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   285
            Left            =   12390
            TabIndex        =   19
            Top             =   60
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cuotas a Penalizar"
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
         Begin Threed.SSPanel pnl_Tit_EvaIni 
            Height          =   285
            Left            =   10470
            TabIndex        =   23
            Top             =   330
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Inicio"
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
         Begin Threed.SSPanel pnl_Tit_EvaFin 
            Height          =   285
            Left            =   11430
            TabIndex        =   24
            Top             =   330
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fin"
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
         Begin Threed.SSPanel SSPanel14 
            Height          =   285
            Left            =   10470
            TabIndex        =   25
            Top             =   60
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cuotas a Evaluar"
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   14715
         _Version        =   65536
         _ExtentX        =   25956
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
            Height          =   555
            Left            =   570
            TabIndex        =   8
            Top             =   30
            Width           =   4605
            _Version        =   65536
            _ExtentX        =   8123
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "Evaluación y Asignación de Premio Buen Pagador"
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
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   12030
            Top             =   180
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Presentación Preliminar"
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_290.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   9
         Top             =   750
         Width           =   14715
         _Version        =   65536
         _ExtentX        =   25956
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
            Left            =   1830
            Picture         =   "OpeTra_frm_290.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Exportar a Excel Evaluación PBP"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_290.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Imprimir Evaluación PBP"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Acepta 
            Height          =   585
            Left            =   2430
            Picture         =   "OpeTra_frm_290.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Asignar PBP"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_290.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Modificar Evaluación PBP"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   14100
            Picture         =   "OpeTra_frm_290.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Regene 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_290.frx":14B8
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Regenerar Propuesta de Asignación PBP"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   435
         Left            =   30
         TabIndex        =   20
         Top             =   1440
         Width           =   14715
         _Version        =   65536
         _ExtentX        =   25956
         _ExtentY        =   767
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
         Begin Threed.SSPanel pnl_Period 
            Height          =   315
            Left            =   1440
            TabIndex        =   21
            Top             =   60
            Width           =   13215
            _Version        =   65536
            _ExtentX        =   23310
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
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
            Alignment       =   1
         End
         Begin VB.Label Label1 
            Caption         =   "Período:"
            Height          =   315
            Left            =   60
            TabIndex        =   22
            Top             =   60
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_Pro_EvaPBP_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function MesEspanol(mes As Integer) As String
   Select Case mes
      Case 1:      MesEspanol = "ENERO"
      Case 2:      MesEspanol = "FEBRERO"
      Case 3:      MesEspanol = "MARZO"
      Case 4:      MesEspanol = "ABRIL"
      Case 5:      MesEspanol = "MAYO"
      Case 6:      MesEspanol = "JUNIO"
      Case 7:      MesEspanol = "JULIO"
      Case 8:      MesEspanol = "AGOSTO"
      Case 9:      MesEspanol = "SETIEMBRE"
      Case 10:     MesEspanol = "OCTUBRE"
      Case 11:     MesEspanol = "NOVIEMBRE"
      Case 12:     MesEspanol = "DICIEMBRE"
   End Select
End Function

Private Sub cmd_Acepta_Click()
   'Validar que no existan Bonos Pendientes de Asignación
   If modmip_gf_TotalPBP(CInt(moddat_g_str_Codigo), CInt(moddat_g_str_CodIte), 3) > 0 Then
      MsgBox "No se puede dar por cerrado la Asignación de PBP porque existen clientes Pendientes de Evaluación.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de Asignar el PBP?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_CierrePeriodo
   Screen.MousePointer = 0
   
   moddat_g_int_FlgAct = 2
   Unload Me
End Sub

Private Sub cmd_Editar_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 9
   moddat_g_str_NumOpe = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)
   moddat_g_int_FlgAct_1 = 1
   
   frm_Pro_EvaPBP_04.Show 1
   
   If moddat_g_int_FlgAct_1 = 2 Then
      moddat_g_int_FlgAct = 2
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_ExpExc_Click()
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_ExpExcel
   Call fs_ExpExcel_Calificacion
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Imprim_Click()
   If MsgBox("¿Está seguro de imprimir el reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Se conecta al crystal report
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   'Se envia las tablas correspondientes en el orden que fueron utilizadas
   crp_Imprim.DataFiles(0) = "CRE_DETPBP"
   crp_Imprim.DataFiles(1) = "CRE_HIPMAE"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "CRE_PRODUC"
   
   'Se pone la llamada del nombre del reporte y se escoge donde se destinara el reporte
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_EVAPBP_11.RPT"
   crp_Imprim.SelectionFormula = "{CRE_DETPBP.DETPBP_PERMES} = " & moddat_g_str_Codigo & " AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CRE_DETPBP.DETPBP_PERANO} = " & moddat_g_str_CodIte & " "
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Action = 1
End Sub

Private Sub cmd_Regene_Click()
   If MsgBox("¿Está seguro de volver a ejecutar el Proceso?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Eliminando Cabecera
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM CRE_CABPBP "
   g_str_Parame = g_str_Parame & " WHERE CABPBP_PERMES = " & moddat_g_str_Codigo & " "
   g_str_Parame = g_str_Parame & "   AND CABPBP_PERANO = " & moddat_g_str_CodIte & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "No se pudo leer la tabla CRE_CABPBP.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If

   'Eliminando Detalle
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM CRE_DETPBP "
   g_str_Parame = g_str_Parame & " WHERE DETPBP_PERMES = " & moddat_g_str_Codigo & " "
   g_str_Parame = g_str_Parame & "   AND DETPBP_PERANO = " & moddat_g_str_CodIte & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "No se pudo leer la tabla CRE_DETPBP.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GeneraPBP(CInt(moddat_g_str_Codigo), CInt(moddat_g_str_CodIte))
   Screen.MousePointer = 0
   
   MsgBox "El Proceso ha terminado correctamente.", vbInformation, modgen_g_str_NomPlt
   moddat_g_int_FlgAct = 2
   
   Screen.MousePointer = 11
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   pnl_Period.Caption = moddat_gf_Consulta_ParDes("033", moddat_g_str_Codigo) & " - " & moddat_g_str_CodIte
   
   Call gs_CentraForm(Me)
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Buscar
   
   Call gs_SetFocus(cmd_Regene)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 1625
   grd_Listad.ColWidth(1) = 3965
   grd_Listad.ColWidth(2) = 2675
   grd_Listad.ColWidth(3) = 1175
   grd_Listad.ColWidth(4) = 995
   grd_Listad.ColWidth(5) = 965
   grd_Listad.ColWidth(6) = 965
   grd_Listad.ColWidth(7) = 965
   grd_Listad.ColWidth(8) = 965
   grd_Listad.ColWidth(9) = 0
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignCenterCenter
   grd_Listad.ColAlignment(7) = flexAlignCenterCenter
   grd_Listad.ColAlignment(8) = flexAlignCenterCenter
End Sub

Private Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_Buscar()
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT A.DETPBP_NUMOPE, B.HIPMAE_TDOCLI, B.HIPMAE_NDOCLI, TRIM(D.PRODUC_DESCRI) AS PRODUCTO, "
   g_str_Parame = g_str_Parame & "       A.DETPBP_FLGPBP, A.DETPBP_CUOCON, A.DETPBP_CEVAIN, A.DETPBP_CEVAFN, "
   g_str_Parame = g_str_Parame & "       A.DETPBP_CIPNCL, A.DETPBP_CFPNCL, TRIM(E.PARDES_DESCRI) AS FLAG_PBP, "
   g_str_Parame = g_str_Parame & "       TRIM(C.DATGEN_APEPAT)||'/'||TRIM(C.DATGEN_APEMAT)||'/'||TRIM(C.DATGEN_NOMBRE) AS NOMBRE_CLIENTE "
   g_str_Parame = g_str_Parame & "  FROM CRE_DETPBP A "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_HIPMAE B ON B.HIPMAE_NUMOPE = A.DETPBP_NUMOPE "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN C ON C.DATGEN_TIPDOC = B.HIPMAE_TDOCLI AND TRIM(C.DATGEN_NUMDOC) = TRIM(B.HIPMAE_NDOCLI) "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_PRODUC D ON D.PRODUC_CODIGO = B.HIPMAE_CODPRD "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = '275' AND E.PARDES_CODITE = A.DETPBP_FLGPBP "
   g_str_Parame = g_str_Parame & " WHERE DETPBP_PERMES = " & moddat_g_str_Codigo & " "
   g_str_Parame = g_str_Parame & "   AND DETPBP_PERANO = " & moddat_g_str_CodIte & " "
   g_str_Parame = g_str_Parame & " ORDER BY HIPMAE_CODPRD ASC, DATGEN_APEPAT ASC, DATGEN_APEMAT ASC, DATGEN_NOMBRE ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      MsgBox "No se pudo leer la tabla CRE_DETPBP.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
     Exit Sub
   End If
   
   grd_Listad.Redraw = False
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = gf_Formato_NumOpe(g_rst_Princi!DETPBP_NUMOPE)
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!NOMBRE_CLIENTE)
      
      grd_Listad.Col = 2
      grd_Listad.Text = Trim(g_rst_Princi!PRODUCTO)
      
      grd_Listad.Col = 3
      grd_Listad.Text = Trim(g_rst_Princi!FLAG_PBP)
      
      grd_Listad.Col = 4
      grd_Listad.Text = CStr(g_rst_Princi!DETPBP_CUOCON)
      
      grd_Listad.Col = 5
      grd_Listad.Text = CStr(g_rst_Princi!DETPBP_CEVAIN)
      
      grd_Listad.Col = 6
      grd_Listad.Text = CStr(g_rst_Princi!DETPBP_CEVAFN)
      
      grd_Listad.Col = 7
      grd_Listad.Text = CStr(g_rst_Princi!DETPBP_CIPNCL)
      
      grd_Listad.Col = 8
      grd_Listad.Text = CStr(g_rst_Princi!DETPBP_CFPNCL)
      
      grd_Listad.Col = 9
      grd_Listad.Text = Trim(g_rst_Princi!DETPBP_NUMOPE)
      
      g_rst_Princi.MoveNext
   Loop
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Call gs_UbiIniGrid(grd_Listad)
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_Editar_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub fs_GeneraPBP(ByVal p_PerMes As Integer, ByVal p_PerAno As Integer)
Dim r_arr_ParPrd()            As moddat_tpo_Genera
Dim r_arr_DiaAtr(6)           As Integer
Dim r_str_NumOpe              As String
Dim r_int_NumCuo              As Integer
Dim r_int_CuoIni              As Integer
Dim r_int_CuoFin              As Integer
Dim r_int_AtrMax              As Integer
Dim r_int_DiaAtr              As Integer
Dim r_str_FecVct              As String
Dim r_str_FecPag              As String
Dim r_rst_Cuotas              As ADODB.Recordset
Dim r_rst_CuoCon              As ADODB.Recordset
Dim r_int_Contad              As Integer
Dim r_int_FlgApl              As Integer
Dim r_int_FlgPBP              As Integer

Dim r_dbl_CapCof              As Double
Dim r_dbl_IntCof              As Double
Dim r_dbl_ComCof              As Double

Dim r_dbl_CapPen_Cli          As Double
Dim r_dbl_IntPen_Cli          As Double
Dim r_dbl_CapPen_Ult_Cli      As Double
Dim r_dbl_IntPen_Ult_Cli      As Double

Dim r_dbl_CapPen_Cli_1        As Double
Dim r_dbl_CapPen_Cli_2        As Double
Dim r_dbl_CapPen_Cli_3        As Double
Dim r_dbl_CapPen_Cli_4        As Double
Dim r_dbl_CapPen_Cli_5        As Double
Dim r_dbl_CapPen_Cli_6        As Double

Dim r_dbl_IntPen_Cli_1        As Double
Dim r_dbl_IntPen_Cli_2        As Double
Dim r_dbl_IntPen_Cli_3        As Double
Dim r_dbl_IntPen_Cli_4        As Double
Dim r_dbl_IntPen_Cli_5        As Double
Dim r_dbl_IntPen_Cli_6        As Double

Dim r_dbl_CapPen_Cof          As Double
Dim r_dbl_IntPen_Cof          As Double
Dim r_dbl_ComPen_Cof          As Double
Dim r_dbl_CapPen_Ult_Cof      As Double
Dim r_dbl_IntPen_Ult_Cof      As Double
Dim r_dbl_ComPen_Ult_Cof      As Double

Dim r_int_CuoDis              As Integer
Dim r_int_DiaMax              As Integer
Dim r_str_FecIni              As String
Dim r_str_FecFin              As String
Dim r_dbl_CapCli              As Double
Dim r_dbl_IntCli              As Double

Dim r_int_NumEva              As Integer
Dim r_int_EvaAsg              As Integer
Dim r_int_EvaPer              As Integer

Dim r_int_CuoIni_Eval         As Integer
Dim r_int_CuoFin_Eval         As Integer
Dim r_int_CuoIni_Cast         As Integer
Dim r_int_CuoFin_Cast         As Integer
   
   r_int_NumEva = 0
   r_int_EvaAsg = 0
   r_int_EvaPer = 0
   r_int_CuoIni_Eval = 0
   r_int_CuoFin_Eval = 0
   r_int_CuoIni_Cast = 0
   r_int_CuoFin_Cast = 0
   
   r_str_FecIni = Format(CDate("01/" & Format(p_PerMes, "00") & "/" & Format(p_PerAno, "0000")), "yyyymmdd")
   r_str_FecFin = Format(CDate(Format(ff_Ultimo_Dia_Mes(p_PerMes, p_PerAno), "00") & "/" & Format(p_PerMes, "00") & "/" & Format(p_PerAno, "0000")), "yyyymmdd")
   
   'Creando Cursor Principal
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT HIPMAE_CODPRD, HIPMAE_CODSUB, HIPCUO_NUMOPE, HIPMAE_SALCON, HIPCUO_NUMCUO, "
   g_str_Parame = g_str_Parame & "       HIPCUO_CAPITA, HIPCUO_INTERE, HIPCUO_COMCOF, HIPCUO_SALCAP "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO A, CRE_HIPMAE B "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = HIPMAE_NUMOPE "
   g_str_Parame = g_str_Parame & "   AND HIPMAE_SITUAC = 2 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_FECVCT >= " & r_str_FecIni & " "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_FECVCT <= " & r_str_FecFin & " "
   g_str_Parame = g_str_Parame & "   AND ((HIPMAE_CODPRD = '001' AND HIPCUO_TIPCRO = 4) OR "
   g_str_Parame = g_str_Parame & "        (HIPMAE_CODPRD = '003' AND HIPCUO_TIPCRO = 4) OR "
   g_str_Parame = g_str_Parame & "        (HIPMAE_CODPRD = '004' AND HIPCUO_TIPCRO = 4) OR "
   g_str_Parame = g_str_Parame & "        (HIPMAE_CODPRD = '006' AND HIPCUO_TIPCRO = 2) OR "
   g_str_Parame = g_str_Parame & "        (HIPMAE_CODPRD = '007' AND HIPCUO_TIPCRO = 4) OR "
   g_str_Parame = g_str_Parame & "        (HIPMAE_CODPRD = '009' AND HIPCUO_TIPCRO = 4) OR "
   g_str_Parame = g_str_Parame & "        (HIPMAE_CODPRD = '010' AND HIPCUO_TIPCRO = 4) OR "
   g_str_Parame = g_str_Parame & "        (HIPMAE_CODPRD = '013' AND HIPCUO_TIPCRO = 4) OR "
   g_str_Parame = g_str_Parame & "        (HIPMAE_CODPRD = '014' AND HIPCUO_TIPCRO = 4) OR "
   g_str_Parame = g_str_Parame & "        (HIPMAE_CODPRD = '015' AND HIPCUO_TIPCRO = 4) OR "
   g_str_Parame = g_str_Parame & "        (HIPMAE_CODPRD = '016' AND HIPCUO_TIPCRO = 4) OR "
   g_str_Parame = g_str_Parame & "        (HIPMAE_CODPRD = '017' AND HIPCUO_TIPCRO = 4) OR "
   g_str_Parame = g_str_Parame & "        (HIPMAE_CODPRD = '018' AND HIPCUO_TIPCRO = 4)) "
   
   g_str_Parame = g_str_Parame & " UNION "
   
   g_str_Parame = g_str_Parame & "SELECT HIPMAE_CODPRD, HIPMAE_CODSUB, HIPCUO_NUMOPE, HIPMAE_SALCON, HIPCUO_NUMCUO, HIPCUO_CAPITA, HIPCUO_INTERE, HIPCUO_COMCOF, HIPCUO_SALCAP "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO A "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_HIPMAE B ON B.HIPMAE_NUMOPE = A.HIPCUO_NUMOPE "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE IN (SELECT HIPMAE_NUMOPE FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & "                         Where HIPMAE_SITUAC = 9 "
   g_str_Parame = g_str_Parame & "                           AND HIPMAE_FECCAN >= " & r_str_FecIni & " AND HIPMAE_FECCAN <= " & r_str_FecFin & " AND HIPMAE_CUOPEN = 0 AND HIPMAE_CODPRD IN ('001','003','004','006','007','009','010','013','014','015','016','017','018')) "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 4 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_FECVCT >= " & r_str_FecIni
   g_str_Parame = g_str_Parame & "   AND HIPCUO_FECVCT <= " & r_str_FecFin
   g_str_Parame = g_str_Parame & " ORDER BY HIPCUO_NUMOPE ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      'Para leer Parámetro de Días de Atraso Máximo para Aplicación de PBP
      r_int_AtrMax = 0
      
      If moddat_gf_Consulta_ParSubPrd(r_arr_ParPrd, g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "051", "041") Then
         r_int_AtrMax = r_arr_ParPrd(1).Genera_Cantid
      End If
      
      r_str_NumOpe = g_rst_Princi!HIPCUO_NUMOPE
      r_int_NumCuo = CInt(g_rst_Princi!HIPCUO_NUMCUO)
      
      'Cuotas de Evaluacion
      r_int_CuoIni_Eval = (r_int_NumCuo * 6) - 6
      r_int_CuoFin_Eval = (r_int_NumCuo * 6) - 1
      If r_int_NumCuo = 1 Then
         r_int_CuoIni_Eval = 1
      End If
      
      'Para leer Cuotas de TNC de cada Operación y determinar Días de Atraso de cada cuota
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT HIPCUO_NUMCUO, HIPCUO_FECVCT, HIPCUO_FECPAG, HIPCUO_SITUAC "
      g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
      g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & g_rst_Princi!HIPCUO_NUMOPE & "' "
      g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 1 "
      g_str_Parame = g_str_Parame & "   AND HIPCUO_NUMCUO >= " & CStr(r_int_CuoIni_Eval) & " "
      g_str_Parame = g_str_Parame & "   AND HIPCUO_NUMCUO <= " & CStr(r_int_CuoFin_Eval) & " "
      g_str_Parame = g_str_Parame & " ORDER BY HIPCUO_NUMCUO ASC "
   
      If Not gf_EjecutaSQL(g_str_Parame, r_rst_Cuotas, 3) Then
         Exit Sub
      End If
   
      If Not (r_rst_Cuotas.BOF And r_rst_Cuotas.EOF) Then
         r_int_Contad = 1
         r_int_FlgApl = 0
         r_int_DiaMax = 0
         r_rst_Cuotas.MoveFirst
         
         Do While Not r_rst_Cuotas.EOF
            If r_rst_Cuotas!HIPCUO_SITUAC = 2 Then
               r_str_FecVct = gf_FormatoFecha(CStr(r_rst_Cuotas!HIPCUO_FECVCT))
               r_int_DiaAtr = 0
               If CInt((date - CDate(1)) - CDate(r_str_FecVct)) > r_int_AtrMax Then
                  r_int_DiaAtr = CInt((date - CDate(1)) - CDate(r_str_FecVct))
               Else
                  r_int_FlgApl = 1
               End If
            Else
               r_str_FecVct = gf_FormatoFecha(CStr(r_rst_Cuotas!HIPCUO_FECVCT))
               r_str_FecPag = gf_FormatoFecha(CStr(r_rst_Cuotas!HIPCUO_FECPAG))
               
               If CDate(r_str_FecPag) > CDate(r_str_FecVct) Then
                  r_int_DiaAtr = CInt(CDate(r_str_FecPag) - CDate(r_str_FecVct))
               Else
                  r_int_DiaAtr = 0
               End If
            End If
            
            r_arr_DiaAtr(r_int_Contad) = r_int_DiaAtr
            If r_int_DiaAtr > r_int_DiaMax Then
               r_int_DiaMax = r_int_DiaAtr
            End If
            
            r_int_Contad = r_int_Contad + 1
            r_rst_Cuotas.MoveNext
            DoEvents
         Loop
      End If
   
      r_rst_Cuotas.Close
      Set r_rst_Cuotas = Nothing
         
      'Para determinar si se asigna PBP
      r_int_FlgPBP = 3
      
      If (r_int_FlgApl = 0) Or (r_int_FlgApl = 1 And r_int_DiaMax > r_int_DiaAtr) Then
         r_int_FlgPBP = 1
         If r_int_CuoIni_Eval = 1 Then
            For r_int_Contad = 1 To 5
               If r_arr_DiaAtr(r_int_Contad) > r_int_AtrMax Then
                  r_int_FlgPBP = 2
               End If
            Next r_int_Contad
         Else
            For r_int_Contad = 1 To 6
               If r_arr_DiaAtr(r_int_Contad) > r_int_AtrMax Then
                  r_int_FlgPBP = 2
               End If
            Next r_int_Contad
         End If
      End If
      
      'Leyendo Capital, Interes y Comisión de Cronograma Cliente y Cofide/Mivivienda
      r_dbl_CapCli = 0
      r_dbl_IntCli = 0
      r_dbl_CapCof = 0
      r_dbl_IntCof = 0
      r_dbl_ComCof = 0
      
      If g_rst_Princi!HIPMAE_CODPRD = "006" Then
         r_dbl_CapCli = g_rst_Princi!HIPCUO_CAPITA
         r_dbl_IntCli = g_rst_Princi!HIPCUO_INTERE
      Else
         r_dbl_CapCof = g_rst_Princi!HIPCUO_CAPITA
         r_dbl_IntCof = g_rst_Princi!HIPCUO_INTERE
         r_dbl_ComCof = g_rst_Princi!HIPCUO_COMCOF
         r_dbl_CapCli = r_dbl_CapCof
         
         'Leer Cuota de TC (Cliente)
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "SELECT HIPCUO_INTERE "
         g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
         g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & g_rst_Princi!HIPCUO_NUMOPE & "' "
         g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 2 "
         g_str_Parame = g_str_Parame & "   AND HIPCUO_NUMCUO = " & CStr(r_int_NumCuo) & " "
         
         If Not gf_EjecutaSQL(g_str_Parame, r_rst_CuoCon, 3) Then
            Exit Sub
         End If
         
         r_rst_CuoCon.MoveFirst
         r_dbl_IntCli = r_rst_CuoCon!HIPCUO_INTERE
         
         r_rst_CuoCon.Close
         Set r_rst_CuoCon = Nothing
      End If
      
      'Obteniendo Nro. de Cuotas a Distribuir Penalidad
      r_int_CuoDis = 6
      r_int_CuoIni_Cast = (r_int_NumCuo * 6) + 1
      r_int_CuoFin_Cast = (r_int_NumCuo * 6) + 6
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT COUNT(*) AS TOTCUO "
      g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
      g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & g_rst_Princi!HIPCUO_NUMOPE & "' "
      g_str_Parame = g_str_Parame & "   AND HIPCUO_NUMCUO >= " & CStr(r_int_CuoIni_Cast) & " "
      g_str_Parame = g_str_Parame & "   AND HIPCUO_NUMCUO <= " & CStr(r_int_CuoFin_Cast) & " "
      g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 1 "
      g_str_Parame = g_str_Parame & "   AND HIPCUO_SITUAC = 2 "
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      r_int_CuoDis = g_rst_Genera!TOTCUO
      If r_int_CuoDis = 0 Then r_int_FlgPBP = 1
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
         
      'Aplicar Penalidad PBP
      r_dbl_CapPen_Cli = 0:      r_dbl_IntPen_Cli = 0
      r_dbl_CapPen_Ult_Cli = 0:  r_dbl_IntPen_Ult_Cli = 0
      
      r_dbl_CapPen_Cli_1 = 0: r_dbl_CapPen_Cli_2 = 0: r_dbl_CapPen_Cli_3 = 0: r_dbl_CapPen_Cli_4 = 0: r_dbl_CapPen_Cli_5 = 0: r_dbl_CapPen_Cli_6 = 0
      r_dbl_IntPen_Cli_1 = 0: r_dbl_IntPen_Cli_2 = 0: r_dbl_IntPen_Cli_3 = 0: r_dbl_IntPen_Cli_4 = 0: r_dbl_IntPen_Cli_5 = 0: r_dbl_IntPen_Cli_6 = 0
      
      r_dbl_CapPen_Cof = 0:      r_dbl_IntPen_Cof = 0:      r_dbl_ComPen_Cof = 0
      r_dbl_CapPen_Ult_Cof = 0:  r_dbl_IntPen_Ult_Cof = 0:  r_dbl_ComPen_Ult_Cof = 0
         
      If r_int_FlgPBP = 2 Then
         r_int_EvaPer = r_int_EvaPer + 1
      
         r_dbl_CapPen_Cli = CDbl(Format(r_dbl_CapCli / r_int_CuoDis, "######0.00"))
         r_dbl_IntPen_Cli = CDbl(Format(r_dbl_IntCli / r_int_CuoDis, "#####0.00"))
         r_dbl_CapPen_Cof = CDbl(Format(r_dbl_CapCof / 6, "######0.00"))
         r_dbl_IntPen_Cof = CDbl(Format(r_dbl_IntCof / 6, "#####0.00"))
         r_dbl_ComPen_Cof = CDbl(Format(r_dbl_ComCof / 6, "#####0.00"))
         
         'Ajustando Capital Ultima Cuota
         If r_dbl_CapPen_Cli * r_int_CuoDis < r_dbl_CapCli Then
            r_dbl_CapPen_Ult_Cli = r_dbl_CapPen_Cli + (r_dbl_CapCli - (r_dbl_CapPen_Cli * r_int_CuoDis))
         Else
            r_dbl_CapPen_Ult_Cli = r_dbl_CapPen_Cli - ((r_dbl_CapPen_Cli * r_int_CuoDis) - r_dbl_CapCli)
         End If
         
         'Ajustando Interes Ultima Cuota
         If r_dbl_IntPen_Cli * r_int_CuoDis < r_dbl_IntCli Then
            r_dbl_IntPen_Ult_Cli = r_dbl_IntPen_Cli + (r_dbl_IntCli - (r_dbl_IntPen_Cli * r_int_CuoDis))
         Else
            r_dbl_IntPen_Ult_Cli = r_dbl_IntPen_Cli - ((r_dbl_IntPen_Cli * r_int_CuoDis) - r_dbl_IntCli)
         End If
         
         'Ajustando Capital COFIDE Ultima Cuota
         If r_dbl_CapPen_Cof * r_int_CuoDis < r_dbl_CapCof Then
            r_dbl_CapPen_Ult_Cof = r_dbl_CapPen_Cof + (r_dbl_CapCof - (r_dbl_CapPen_Cof * 6))
         Else
            r_dbl_CapPen_Ult_Cof = r_dbl_CapPen_Cof - ((r_dbl_CapPen_Cof * 6) - r_dbl_CapCof)
         End If
         
         'Ajustando Interes COFIDE Ultima Cuota
         If r_dbl_IntPen_Cof * r_int_CuoDis < r_dbl_IntCof Then
            r_dbl_IntPen_Ult_Cof = r_dbl_IntPen_Cof + (r_dbl_IntCof - (r_dbl_IntPen_Cof * 6))
         Else
            r_dbl_IntPen_Ult_Cof = r_dbl_IntPen_Cof - ((r_dbl_IntPen_Cof * 6) - r_dbl_IntCof)
         End If
         
         'Ajustando Interes COFIDE Ultima Cuota
         If r_dbl_ComPen_Cof * r_int_CuoDis < r_dbl_ComCof Then
            r_dbl_ComPen_Ult_Cof = r_dbl_ComPen_Cof + (r_dbl_ComCof - (r_dbl_ComPen_Cof * 6))
         Else
            r_dbl_ComPen_Ult_Cof = r_dbl_ComPen_Cof - ((r_dbl_ComPen_Cof * 6) - r_dbl_ComCof)
         End If
         
         If r_int_CuoDis = 6 Then
            r_dbl_CapPen_Cli_1 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_2 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_3 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_4 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_5 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_6 = r_dbl_CapPen_Ult_Cli
            r_dbl_IntPen_Cli_1 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_2 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_3 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_4 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_5 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_6 = r_dbl_IntPen_Ult_Cli
         ElseIf r_int_CuoDis = 5 Then
            r_dbl_CapPen_Cli_2 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_3 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_4 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_5 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_6 = r_dbl_CapPen_Ult_Cli
            r_dbl_IntPen_Cli_2 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_3 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_4 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_5 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_6 = r_dbl_IntPen_Ult_Cli
         ElseIf r_int_CuoDis = 4 Then
            r_dbl_CapPen_Cli_3 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_4 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_5 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_6 = r_dbl_CapPen_Ult_Cli
            r_dbl_IntPen_Cli_3 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_4 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_5 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_6 = r_dbl_IntPen_Ult_Cli
         ElseIf r_int_CuoDis = 3 Then
            r_dbl_CapPen_Cli_4 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_5 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_6 = r_dbl_CapPen_Ult_Cli
            r_dbl_IntPen_Cli_4 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_5 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_6 = r_dbl_IntPen_Ult_Cli
         ElseIf r_int_CuoDis = 2 Then
            r_dbl_CapPen_Cli_5 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_6 = r_dbl_CapPen_Ult_Cli
            r_dbl_IntPen_Cli_5 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_6 = r_dbl_IntPen_Ult_Cli
         ElseIf r_int_CuoDis = 1 Then
            r_dbl_CapPen_Cli_6 = r_dbl_CapPen_Ult_Cli
            r_dbl_IntPen_Cli_6 = r_dbl_IntPen_Ult_Cli
         End If
      ElseIf r_int_FlgPBP = 1 Then
         r_int_EvaAsg = r_int_EvaAsg + 1
      End If
      
      r_int_NumEva = r_int_NumEva + 1
      
      'Grabando en CRE_DETPBP
      g_str_Parame = "USP_CRE_DETPBP ("
      g_str_Parame = g_str_Parame & CStr(p_PerMes) & ", "
      g_str_Parame = g_str_Parame & CStr(p_PerAno) & ", "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!HIPCUO_NUMOPE & "', "
      g_str_Parame = g_str_Parame & CStr(r_int_FlgPBP) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_NumCuo) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapCli) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntCli) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapCof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntCof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_ComCof) & ", "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!HIPCUO_SALCAP + g_rst_Princi!HIPCUO_CAPITA) & ", "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!HIPCUO_SALCAP) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_CuoIni_Eval) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_CuoFin_Eval) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_CuoIni_Cast) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_CuoFin_Cast) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapPen_Cli_1) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntPen_Cli_1) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapPen_Cli_2) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntPen_Cli_2) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapPen_Cli_3) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntPen_Cli_3) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapPen_Cli_4) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntPen_Cli_4) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapPen_Cli_5) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntPen_Cli_5) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapPen_Cli_6) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntPen_Cli_6) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_CuoIni_Cast) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_CuoFin_Cast) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_ComPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_ComPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_ComPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_ComPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_ComPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapPen_Ult_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntPen_Ult_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_ComPen_Ult_Cof) & ", "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & "1)"
            
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         MsgBox "No se pudo ejecutar el procedimiento USP_CRE_DETPBP.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      g_rst_Princi.MoveNext
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Grabando Cabecera
   g_str_Parame = "USP_CRE_CABPBP ("
   g_str_Parame = g_str_Parame & CStr(p_PerMes) & ", "
   g_str_Parame = g_str_Parame & CStr(p_PerAno) & ", "
   g_str_Parame = g_str_Parame & CStr(r_int_NumEva) & ", "
   g_str_Parame = g_str_Parame & CStr(r_int_EvaAsg) & ", "
   g_str_Parame = g_str_Parame & CStr(r_int_EvaPer) & ", "
   g_str_Parame = g_str_Parame & "1, "
   
   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
   g_str_Parame = g_str_Parame & "1)"
         
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      MsgBox "No se pudo ejecutar el procedimiento USP_CRE_CABPBP.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
End Sub

Private Sub pnl_Tit_NumOpe_Click()
   If Len(Trim(pnl_Tit_NumOpe.Tag)) = 0 Or pnl_Tit_NumOpe.Tag = "D" Then
      pnl_Tit_NumOpe.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Tit_NumOpe.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub

Private Sub pnl_Tit_NomCli_Click()
   If Len(Trim(pnl_Tit_NomCli.Tag)) = 0 Or pnl_Tit_NomCli.Tag = "D" Then
      pnl_Tit_NomCli.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Tit_NomCli.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub

Private Sub pnl_Tit_Produc_Click()
   If Len(Trim(pnl_Tit_Produc.Tag)) = 0 Or pnl_Tit_Produc.Tag = "D" Then
      pnl_Tit_Produc.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_Tit_Produc.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub

Private Sub pnl_Tit_FlgPBP_Click()
   If Len(Trim(pnl_Tit_FlgPBP.Tag)) = 0 Or pnl_Tit_FlgPBP.Tag = "D" Then
      pnl_Tit_FlgPBP.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 3, "C")
   Else
      pnl_Tit_FlgPBP.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 3, "C-")
   End If
End Sub

Private Sub pnl_Tit_NumCuo_Click()
   If Len(Trim(pnl_Tit_NumCuo.Tag)) = 0 Or pnl_Tit_NumCuo.Tag = "D" Then
      pnl_Tit_NumCuo.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 4, "N")
   Else
      pnl_Tit_NumCuo.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 4, "N-")
   End If
End Sub

Private Sub pnl_Tit_EvaIni_Click()
   If Len(Trim(pnl_Tit_EvaIni.Tag)) = 0 Or pnl_Tit_EvaIni.Tag = "D" Then
      pnl_Tit_EvaIni.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 5, "N")
   Else
      pnl_Tit_EvaIni.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 5, "N-")
   End If
End Sub

Private Sub pnl_Tit_EvaFin_Click()
   If Len(Trim(pnl_Tit_EvaFin.Tag)) = 0 Or pnl_Tit_EvaFin.Tag = "D" Then
      pnl_Tit_EvaFin.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 6, "N")
   Else
      pnl_Tit_EvaFin.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 6, "N-")
   End If
End Sub

Private Sub pnl_Tit_PenIni_Click()
   If Len(Trim(pnl_Tit_PenIni.Tag)) = 0 Or pnl_Tit_PenIni.Tag = "D" Then
      pnl_Tit_PenIni.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 7, "N")
   Else
      pnl_Tit_PenIni.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 7, "N-")
   End If
End Sub

Private Sub pnl_Tit_PenFin_Click()
   If Len(Trim(pnl_Tit_PenFin.Tag)) = 0 Or pnl_Tit_PenFin.Tag = "D" Then
      pnl_Tit_PenFin.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 8, "N")
   Else
      pnl_Tit_PenFin.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 8, "N-")
   End If
End Sub

Private Sub fs_ExpExcel()
Dim r_obj_Excel         As Excel.Application
Dim r_obj_NomArc        As New Excel.Workbook
Dim r_obj_NomHoj        As New Excel.Worksheet
Dim r_int_ConVer        As Integer
Dim r_rst_Princi        As ADODB.Recordset
Dim r_int_Cont          As Integer
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_DETPBP A, CRE_HIPMAE B "
   g_str_Parame = g_str_Parame & " WHERE DETPBP_PERMES = " & moddat_g_str_Codigo & " "
   g_str_Parame = g_str_Parame & "   AND DETPBP_PERANO = " & moddat_g_str_CodIte & " "
   g_str_Parame = g_str_Parame & "   AND DETPBP_NUMOPE = HIPMAE_NUMOPE "
   g_str_Parame = g_str_Parame & " ORDER BY DETPBP_NUMOPE ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If r_rst_Princi.BOF And r_rst_Princi.EOF Then
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      MsgBox "No se encontraron datos a reportar.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   Set r_obj_NomArc = r_obj_Excel.Workbooks.Add
   Set r_obj_NomHoj = r_obj_NomArc.Worksheets("Hoja1")
   
   With r_obj_NomArc.ActiveSheet
      .Cells(1, 1) = "ITEM":                             .Columns("A").ColumnWidth = 8
      .Cells(1, 2) = "PERIODO":                          .Columns("B").ColumnWidth = 15:        .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 3) = "PRODUCTO":                         .Columns("C").ColumnWidth = 30:        .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 4) = "SITUACION PBP":                    .Columns("D").ColumnWidth = 15:        .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 5) = "NRO. OPERACION":                   .Columns("E").ColumnWidth = 20:        .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 6) = "CLIENTE":                          .Columns("F").ColumnWidth = 50
      .Cells(1, 7) = "NRO. OPERACION MIVIVIENDA":        .Columns("G").ColumnWidth = 30:        .Columns("G").HorizontalAlignment = xlHAlignCenter:      .Columns("G").NumberFormat = "@"
      .Cells(1, 8) = "NRO. OPERACION COFIDE":            .Columns("H").ColumnWidth = 30:        .Columns("H").HorizontalAlignment = xlHAlignCenter:      .Columns("H").NumberFormat = "@"
      .Cells(1, 9) = "CUOTA TC":                         .Columns("I").ColumnWidth = 15:        .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 10) = "CAPITAL TRAMO CLIENTE":           .Columns("J").ColumnWidth = 30:        .Columns("J").NumberFormat = "###,##0.00"
      .Cells(1, 11) = "INTERES TRAMO CLIENTE":           .Columns("K").ColumnWidth = 30:        .Columns("K").NumberFormat = "###,##0.00"
      .Cells(1, 12) = "TOTAL CUOTA TRAMO CLIENTE":       .Columns("L").ColumnWidth = 30:        .Columns("L").NumberFormat = "###,##0.00"
      .Cells(1, 13) = "CAPITAL TRAMO COFIDE/MVI":        .Columns("M").ColumnWidth = 30:        .Columns("M").NumberFormat = "###,##0.00"
      .Cells(1, 14) = "INTERES TRAMO COFIDE/MVI":        .Columns("N").ColumnWidth = 30:        .Columns("N").NumberFormat = "###,##0.00"
      .Cells(1, 15) = "COMISION TRAMO COFIDE/MVI":       .Columns("O").ColumnWidth = 30:        .Columns("O").NumberFormat = "###,##0.00"
      .Cells(1, 16) = "TOTAL CUOTA TRAMO COFIDE/MVI":    .Columns("P").ColumnWidth = 30:        .Columns("P").NumberFormat = "###,##0.00"
      .Cells(1, 17) = "CUOTA INICIO EVALUACION":         .Columns("Q").ColumnWidth = 25:        .Columns("Q").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 18) = "CUOTA FIN EVALUACION":            .Columns("R").ColumnWidth = 25:        .Columns("R").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 19) = "CUOTA INICIO PENALIDAD":          .Columns("S").ColumnWidth = 25:        .Columns("S").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 20) = "CUOTA FIN PENALIDAD":             .Columns("T").ColumnWidth = 25:        .Columns("T").HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(1, 1), .Cells(1, 20)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 20)).HorizontalAlignment = xlHAlignCenter
   End With
   
   r_rst_Princi.MoveFirst
   r_int_ConVer = 2
   r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer - 1, 1), r_obj_Excel.Cells(r_int_ConVer - 1, 20)).Borders(xlEdgeTop).LineStyle = xlContinuous
   
   Do While Not r_rst_Princi.EOF
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Format(r_rst_Princi!DETPBP_PERMES, "00") & " - " & Format(r_rst_Princi!DETPBP_PERANO, "0000")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = moddat_gf_Consulta_Produc(r_rst_Princi!HIPMAE_CODPRD)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = moddat_gf_Consulta_ParDes("275", CStr(r_rst_Princi!DETPBP_FLGPBP))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = gf_Formato_NumOpe(r_rst_Princi!HIPMAE_NUMOPE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = moddat_gf_Buscar_NomCli(r_rst_Princi!HIPMAE_TDOCLI, r_rst_Princi!HIPMAE_NDOCLI)
      
      If r_rst_Princi!HIPMAE_CODPRD = "001" Or r_rst_Princi!HIPMAE_CODPRD = "003" Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = Trim(r_rst_Princi!HIPMAE_OPEMVI & "")
      End If
      
      If r_rst_Princi!HIPMAE_CODPRD = "003" Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = Trim(r_rst_Princi!HIPMAE_OPEMV1 & "")
      ElseIf r_rst_Princi!HIPMAE_CODPRD <> "001" And r_rst_Princi!HIPMAE_CODPRD <> "006" Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = Trim(r_rst_Princi!HIPMAE_OPEMVI & "")
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = CStr(r_rst_Princi!DETPBP_CUOCON)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = r_rst_Princi!DETPBP_CAPCLI
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = r_rst_Princi!DETPBP_INTCLI
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = r_rst_Princi!DETPBP_CAPCLI + r_rst_Princi!DETPBP_INTCLI
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = r_rst_Princi!DETPBP_CAPADE
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = r_rst_Princi!DETPBP_INTADE
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = r_rst_Princi!DETPBP_COMADE
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = r_rst_Princi!DETPBP_CAPADE + r_rst_Princi!DETPBP_INTADE + r_rst_Princi!DETPBP_COMADE
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = r_rst_Princi!DETPBP_CEVAIN
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = r_rst_Princi!DETPBP_CEVAFN
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = r_rst_Princi!DETPBP_CIPNCL
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = r_rst_Princi!DETPBP_CFPNCL
      
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer, 1), r_obj_Excel.Cells(r_int_ConVer, 20)).Borders(xlEdgeTop).LineStyle = xlContinuous
      r_int_ConVer = r_int_ConVer + 1
      r_rst_Princi.MoveNext
   Loop

   r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer, 1), r_obj_Excel.Cells(r_int_ConVer, 20)).Borders(xlEdgeTop).LineStyle = xlContinuous
   r_obj_Excel.Range(r_obj_Excel.Cells(1, 1), r_obj_Excel.Cells(1, 20)).Interior.Color = RGB(146, 208, 80)
   
   For r_int_Cont = 1 To r_int_ConVer - 1
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 1), r_obj_Excel.Cells(r_int_Cont, 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 2), r_obj_Excel.Cells(r_int_Cont, 2)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 3), r_obj_Excel.Cells(r_int_Cont, 3)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 4), r_obj_Excel.Cells(r_int_Cont, 4)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 5), r_obj_Excel.Cells(r_int_Cont, 5)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 6), r_obj_Excel.Cells(r_int_Cont, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 7), r_obj_Excel.Cells(r_int_Cont, 7)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 8), r_obj_Excel.Cells(r_int_Cont, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 9), r_obj_Excel.Cells(r_int_Cont, 9)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 10), r_obj_Excel.Cells(r_int_Cont, 10)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 11), r_obj_Excel.Cells(r_int_Cont, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 12), r_obj_Excel.Cells(r_int_Cont, 12)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 13), r_obj_Excel.Cells(r_int_Cont, 13)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 14), r_obj_Excel.Cells(r_int_Cont, 14)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 15), r_obj_Excel.Cells(r_int_Cont, 15)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 16), r_obj_Excel.Cells(r_int_Cont, 16)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 17), r_obj_Excel.Cells(r_int_Cont, 17)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 18), r_obj_Excel.Cells(r_int_Cont, 18)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 19), r_obj_Excel.Cells(r_int_Cont, 19)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 20), r_obj_Excel.Cells(r_int_Cont, 20)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 21), r_obj_Excel.Cells(r_int_Cont, 21)).Borders(xlEdgeLeft).LineStyle = xlContinuous
   Next

   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   'Ordenando por Producto y Cliente
   r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(1, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer - 1, 20)).Font.Size = 9
   r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(1, 2), r_obj_Excel.Cells(r_int_ConVer - 1, 20)).Sort r_obj_Excel.Range("C1"), xlAscending, r_obj_Excel.Range("D1"), , xlAscending, r_obj_Excel.Range("F1"), , xlAscending, , , xlYes
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_ExpExcel_Calificacion()
Dim r_rst_Princi        As ADODB.Recordset
Dim r_obj_Excel         As Excel.Application
Dim r_int_ConVer        As Integer
Dim r_int_Cont          As Integer
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT " & moddat_g_str_CodIte & "    AS EJERCICIO, "
   g_str_Parame = g_str_Parame & "       " & moddat_g_str_Codigo & "    AS PERIODO, "
   g_str_Parame = g_str_Parame & "       '10005339'                     AS CODIGO_IFI, "
   g_str_Parame = g_str_Parame & "       'EDPYME MICASITA'              AS NOMBRE_IFI, "
   g_str_Parame = g_str_Parame & "       CASE WHEN SUBSTR(DETPBP_NUMOPE,1,3) = '004' THEN '803' ELSE '808' END AS PRODUCTO, "
   g_str_Parame = g_str_Parame & "       TRIM(B.HIPMAE_OPEMVI)          AS NRO_PRESTAMO, "
   g_str_Parame = g_str_Parame & "       TRIM(C.DATGEN_NOMBRE)          AS NOMBRES, "
   g_str_Parame = g_str_Parame & "       TRIM(C.DATGEN_APEPAT)||' '||TRIM(C.DATGEN_APEMAT) AS APELLIDOS, "
   g_str_Parame = g_str_Parame & "       CASE WHEN C.DATGEN_TIPDOC = 1 THEN 'PE2'"
   g_str_Parame = g_str_Parame & "            WHEN C.DATGEN_TIPDOC = 2 THEN 'PE0'"
   g_str_Parame = g_str_Parame & "            WHEN C.DATGEN_TIPDOC = 5 THEN 'PE4' END      AS TIPO_DOCUMENTO, "
   g_str_Parame = g_str_Parame & "       C.DATGEN_NUMDOC                AS NUMERO_DOCUMENTO,"
   g_str_Parame = g_str_Parame & "       CASE WHEN DETPBP_FLGPBP = 1 THEN ' ' ELSE 'X' END AS CLASE_PRODUCTO "
   g_str_Parame = g_str_Parame & "  FROM CRE_DETPBP A "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_HIPMAE B ON B.HIPMAE_NUMOPE = A.DETPBP_NUMOPE "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN C ON C.DATGEN_TIPDOC = B.HIPMAE_TDOCLI AND TRIM(C.DATGEN_NUMDOC) = TRIM(B.HIPMAE_NDOCLI) "
   g_str_Parame = g_str_Parame & " WHERE A.DETPBP_PERMES = " & moddat_g_str_Codigo & " "
   g_str_Parame = g_str_Parame & "   AND A.DETPBP_PERANO = " & moddat_g_str_CodIte & " "
      
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If r_rst_Princi.BOF And r_rst_Princi.EOF Then
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      MsgBox "No se encontraron datos a reportar.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   r_obj_Excel.Sheets(1).Name = "Hoja 1"
   
   With r_obj_Excel.ActiveSheet
      .Range(.Cells(2, 1), .Cells(2, 11)).Merge
      .Range(.Cells(3, 2), .Cells(3, 3)).Merge
      
      .Cells(2, 1) = "ARCHIVO DE CALIFICACION - PBP"
      .Cells(3, 1) = "IFI:"
      .Cells(4, 1) = "MES:"
      .Cells(5, 1) = "AÑO:"
      .Cells(3, 2) = "EDPYME MICASITA"
      .Cells(4, 2) = MesEspanol(Val(moddat_g_str_Codigo))
      .Cells(5, 2) = moddat_g_str_CodIte
      
      .Cells(7, 1) = "EJERCICIO":               .Columns("A").ColumnWidth = 10
      .Cells(7, 2) = "PERIODO":                 .Columns("B").ColumnWidth = 10
      .Cells(7, 3) = "IFI":                     .Columns("C").ColumnWidth = 15
      .Cells(7, 4) = "NOMBRE IFI":              .Columns("D").ColumnWidth = 25
      .Cells(7, 5) = "CL.PROD.":                .Columns("E").ColumnWidth = 10
      .Cells(7, 6) = "N° PRESTAMO":             .Columns("F").ColumnWidth = 20
      .Cells(7, 7) = "NOMBRES":                 .Columns("G").ColumnWidth = 35
      .Cells(7, 8) = "APELLIDOS":               .Columns("H").ColumnWidth = 35
      .Cells(7, 9) = "TIPO DOC.":               .Columns("I").ColumnWidth = 10
      .Cells(7, 10) = "NUMERO DOC.":            .Columns("J").ColumnWidth = 15
      .Cells(7, 11) = "PREMIO":                 .Columns("K").ColumnWidth = 10
      
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").HorizontalAlignment = xlHAlignLeft
      .Columns("H").HorizontalAlignment = xlHAlignLeft
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(2, 1), .Cells(2, 20)).Font.Bold = True
      .Range(.Cells(3, 1), .Cells(3, 2)).Font.Bold = True
      .Range(.Cells(2, 1), .Cells(2, 20)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(7, 7), .Cells(7, 7)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(7, 8), .Cells(7, 8)).HorizontalAlignment = xlHAlignCenter
   End With
   
   r_rst_Princi.MoveFirst
   r_int_ConVer = 8
   r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer - 1, 1), r_obj_Excel.Cells(r_int_ConVer - 1, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
   
   Do While Not r_rst_Princi.EOF
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = moddat_g_str_CodIte
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = moddat_g_str_Codigo
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = r_rst_Princi!CODIGO_IFI
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = r_rst_Princi!NOMBRE_IFI
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = r_rst_Princi!PRODUCTO
      If IsNull(r_rst_Princi!NRO_PRESTAMO) Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = ""
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = "'" & fs_Obtiene_NroOperacion(r_rst_Princi!NRO_PRESTAMO)
      End If
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = r_rst_Princi!NOMBRES
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = r_rst_Princi!APELLIDOS
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = r_rst_Princi!TIPO_DOCUMENTO
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = r_rst_Princi!NUMERO_DOCUMENTO
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = r_rst_Princi!CLASE_PRODUCTO

      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer, 1), r_obj_Excel.Cells(r_int_ConVer, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
      r_int_ConVer = r_int_ConVer + 1
      r_rst_Princi.MoveNext
   Loop
   
   r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer, 1), r_obj_Excel.Cells(r_int_ConVer, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
   r_obj_Excel.Range(r_obj_Excel.Cells(7, 1), r_obj_Excel.Cells(7, 11)).Interior.Color = RGB(146, 208, 80)
   r_obj_Excel.Range(r_obj_Excel.Cells(7, 1), r_obj_Excel.Cells(7, 11)).Font.Bold = True
   
   For r_int_Cont = 7 To r_int_ConVer - 1
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 1), r_obj_Excel.Cells(r_int_Cont, 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 2), r_obj_Excel.Cells(r_int_Cont, 2)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 3), r_obj_Excel.Cells(r_int_Cont, 3)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 4), r_obj_Excel.Cells(r_int_Cont, 4)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 5), r_obj_Excel.Cells(r_int_Cont, 5)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 6), r_obj_Excel.Cells(r_int_Cont, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 7), r_obj_Excel.Cells(r_int_Cont, 7)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 8), r_obj_Excel.Cells(r_int_Cont, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 9), r_obj_Excel.Cells(r_int_Cont, 9)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 10), r_obj_Excel.Cells(r_int_Cont, 10)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 11), r_obj_Excel.Cells(r_int_Cont, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_Cont, 12), r_obj_Excel.Cells(r_int_Cont, 12)).Borders(xlEdgeLeft).LineStyle = xlContinuous
   Next
      
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_CierrePeriodo()
Dim r_obj_Excel         As Excel.Application
Dim r_obj_NomArc        As New Excel.Workbook
Dim r_obj_NomHoj        As New Excel.Worksheet
Dim r_int_ConVer        As Integer
Dim r_rst_Princi        As ADODB.Recordset
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_DETPBP "
   g_str_Parame = g_str_Parame & " WHERE DETPBP_PERMES = " & moddat_g_str_Codigo & " "
   g_str_Parame = g_str_Parame & "   AND DETPBP_PERANO = " & moddat_g_str_CodIte & " "
   g_str_Parame = g_str_Parame & " ORDER BY DETPBP_NUMOPE ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If r_rst_Princi.BOF And r_rst_Princi.EOF Then
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      Exit Sub
   End If
   
   r_rst_Princi.MoveFirst
   Do While Not r_rst_Princi.EOF
      'Actualizando Cuota del Tramo Concesional - Cronograma Cliente
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "UPDATE CRE_HIPCUO "
      g_str_Parame = g_str_Parame & "   SET HIPCUO_IMPPAG = HIPCUO_CAPITA + HIPCUO_INTERE, "
      g_str_Parame = g_str_Parame & "       HIPCUO_CAPPAG = HIPCUO_CAPITA, "
      g_str_Parame = g_str_Parame & "       HIPCUO_INTPAG = HIPCUO_INTERE, "
      g_str_Parame = g_str_Parame & "       HIPCUO_SITUAC = 1 , "
      g_str_Parame = g_str_Parame & "       HIPCUO_FECPAG = " & Format(date, "yyyymmdd") & " "
      g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & r_rst_Princi!DETPBP_NUMOPE & "' "
      g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 2 "
      g_str_Parame = g_str_Parame & "   AND HIPCUO_NUMCUO = " & CStr(r_rst_Princi!DETPBP_CUOCON)
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
          Exit Sub
      End If
      
      'Actualizando Cuota del Tramo Concesional - Cronograma COFIDE / Mivivienda
      If Left(r_rst_Princi!DETPBP_NUMOPE, 3) <> "006" Then
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "UPDATE CRE_HIPCUO "
         g_str_Parame = g_str_Parame & "   SET HIPCUO_IMPPAG = HIPCUO_CAPITA + HIPCUO_INTERE + HIPCUO_COMCOF, "
         g_str_Parame = g_str_Parame & "       HIPCUO_CAPPAG = HIPCUO_CAPITA, "
         g_str_Parame = g_str_Parame & "       HIPCUO_INTPAG = HIPCUO_INTERE, "
         g_str_Parame = g_str_Parame & "       HIPCUO_CCFPAG_CON = HIPCUO_COMCOF, "
         g_str_Parame = g_str_Parame & "       HIPCUO_SITUAC = 1, "
         g_str_Parame = g_str_Parame & "       HIPCUO_FECPAG = " & Format(date, "yyyymmdd") & " "
         g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & r_rst_Princi!DETPBP_NUMOPE & "' "
         g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 4 "
         g_str_Parame = g_str_Parame & "   AND HIPCUO_NUMCUO = " & CStr(r_rst_Princi!DETPBP_CUOCON) & " "
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            Exit Sub
         End If
      End If
      
      'Actualizando Saldo Tramo Concesional en CRE_HIPMAE
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "UPDATE CRE_HIPMAE "
      g_str_Parame = g_str_Parame & "   SET HIPMAE_SALCON = " & CStr(r_rst_Princi!DETPBP_SALNUE) & " "
      g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & r_rst_Princi!DETPBP_NUMOPE & "' "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         Exit Sub
      End If
      
      If r_rst_Princi!DETPBP_FLGPBP = 2 Then
         'Aplicando Penalidades en Tramo del Cliente
         Call fs_AplicaPenalidad(r_rst_Princi!DETPBP_NUMOPE, r_rst_Princi!DETPBP_CIPNCL, 1, r_rst_Princi!DETPBP_CUOCON, r_rst_Princi!DETPBP_CAPCL1, r_rst_Princi!DETPBP_INPCL1, 0)
         Call fs_AplicaPenalidad(r_rst_Princi!DETPBP_NUMOPE, r_rst_Princi!DETPBP_CIPNCL + 1, 1, r_rst_Princi!DETPBP_CUOCON, r_rst_Princi!DETPBP_CAPCL2, r_rst_Princi!DETPBP_INPCL2, 0)
         Call fs_AplicaPenalidad(r_rst_Princi!DETPBP_NUMOPE, r_rst_Princi!DETPBP_CIPNCL + 2, 1, r_rst_Princi!DETPBP_CUOCON, r_rst_Princi!DETPBP_CAPCL3, r_rst_Princi!DETPBP_INPCL3, 0)
         Call fs_AplicaPenalidad(r_rst_Princi!DETPBP_NUMOPE, r_rst_Princi!DETPBP_CIPNCL + 3, 1, r_rst_Princi!DETPBP_CUOCON, r_rst_Princi!DETPBP_CAPCL4, r_rst_Princi!DETPBP_INPCL4, 0)
         Call fs_AplicaPenalidad(r_rst_Princi!DETPBP_NUMOPE, r_rst_Princi!DETPBP_CIPNCL + 4, 1, r_rst_Princi!DETPBP_CUOCON, r_rst_Princi!DETPBP_CAPCL5, r_rst_Princi!DETPBP_INPCL5, 0)
         Call fs_AplicaPenalidad(r_rst_Princi!DETPBP_NUMOPE, r_rst_Princi!DETPBP_CIPNCL + 5, 1, r_rst_Princi!DETPBP_CUOCON, r_rst_Princi!DETPBP_CAPCL6, r_rst_Princi!DETPBP_INPCL6, 0)
         
         If Left(r_rst_Princi!DETPBP_NUMOPE, 3) <> "001" And Left(r_rst_Princi!DETPBP_NUMOPE, 3) <> "003" And Left(r_rst_Princi!DETPBP_NUMOPE, 3) <> "006" Then
            'Aplicando Penalidades en Tramo de COFIDE
            Call fs_AplicaPenalidad(r_rst_Princi!DETPBP_NUMOPE, r_rst_Princi!DETPBP_CIPNCO, 3, r_rst_Princi!DETPBP_CUOCON, r_rst_Princi!DETPBP_CAPCO1, r_rst_Princi!DETPBP_INPCO1, r_rst_Princi!DETPBP_COPCO1)
            Call fs_AplicaPenalidad(r_rst_Princi!DETPBP_NUMOPE, r_rst_Princi!DETPBP_CIPNCO + 1, 3, r_rst_Princi!DETPBP_CUOCON, r_rst_Princi!DETPBP_CAPCO2, r_rst_Princi!DETPBP_INPCO2, r_rst_Princi!DETPBP_COPCO2)
            Call fs_AplicaPenalidad(r_rst_Princi!DETPBP_NUMOPE, r_rst_Princi!DETPBP_CIPNCO + 2, 3, r_rst_Princi!DETPBP_CUOCON, r_rst_Princi!DETPBP_CAPCO3, r_rst_Princi!DETPBP_INPCO3, r_rst_Princi!DETPBP_COPCO3)
            Call fs_AplicaPenalidad(r_rst_Princi!DETPBP_NUMOPE, r_rst_Princi!DETPBP_CIPNCO + 3, 3, r_rst_Princi!DETPBP_CUOCON, r_rst_Princi!DETPBP_CAPCO4, r_rst_Princi!DETPBP_INPCO4, r_rst_Princi!DETPBP_COPCO4)
            Call fs_AplicaPenalidad(r_rst_Princi!DETPBP_NUMOPE, r_rst_Princi!DETPBP_CIPNCO + 4, 3, r_rst_Princi!DETPBP_CUOCON, r_rst_Princi!DETPBP_CAPCO5, r_rst_Princi!DETPBP_INPCO5, r_rst_Princi!DETPBP_COPCO5)
            Call fs_AplicaPenalidad(r_rst_Princi!DETPBP_NUMOPE, r_rst_Princi!DETPBP_CIPNCO + 5, 3, r_rst_Princi!DETPBP_CUOCON, r_rst_Princi!DETPBP_CAPCO6, r_rst_Princi!DETPBP_INPCO6, r_rst_Princi!DETPBP_COPCO6)
         End If
      End If
      
      r_rst_Princi.MoveNext
   Loop
      
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   'Actualizando Situación de Evaluación en Cabecera
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "UPDATE CRE_CABPBP "
   g_str_Parame = g_str_Parame & "   SET CABPBP_SITUAC = 2 "
   g_str_Parame = g_str_Parame & " WHERE CABPBP_PERMES = " & moddat_g_str_Codigo & " "
   g_str_Parame = g_str_Parame & "   AND CABPBP_PERANO = " & moddat_g_str_CodIte & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
       Exit Sub
   End If
End Sub

Private Sub fs_AplicaPenalidad(ByVal p_NumOpe As String, ByVal p_NumCuo As Integer, ByVal p_TipCro As Integer, ByVal p_CuoCon As Integer, ByVal p_CapPen As Double, ByVal p_IntPen As Double, ByVal p_ComPen As Double)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "UPDATE CRE_HIPCUO "
   g_str_Parame = g_str_Parame & "   SET HIPCUO_CAPBBP = " & p_CapPen & ", "
   g_str_Parame = g_str_Parame & "       HIPCUO_INTBBP = " & p_IntPen & ", "
   If p_ComPen > 0 Then
      g_str_Parame = g_str_Parame & "       HIPCUO_COMPBP = " & p_ComPen & ", "
   End If
   g_str_Parame = g_str_Parame & "       HIPCUO_CUOBBP = " & CStr(p_CuoCon) & " "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & p_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = " & CStr(p_TipCro) & " "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_NUMCUO = " & CStr(p_NumCuo) & " "
   
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      Exit Sub
   End If
End Sub

Private Function fs_Obtiene_NroOperacion(ByVal p_NumOpeMiv As String) As String
Dim r_str_Tempo1     As String
Dim r_str_Tempo2     As String
   
   If IsNull(p_NumOpeMiv) Then
      fs_Obtiene_NroOperacion = ""
      Exit Function
   End If
   
   r_str_Tempo1 = Left(Trim(p_NumOpeMiv), 8)
   r_str_Tempo2 = Right(Trim(p_NumOpeMiv), 5)
   fs_Obtiene_NroOperacion = r_str_Tempo1 & r_str_Tempo2
End Function
