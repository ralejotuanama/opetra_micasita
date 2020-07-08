VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.MDIForm frm_MnuPri_01 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   8235
   ClientLeft      =   870
   ClientTop       =   2220
   ClientWidth     =   9825
   Icon            =   "OpeTra_frm_008.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9825
      _Version        =   65536
      _ExtentX        =   17330
      _ExtentY        =   1138
      _StockProps     =   15
      BackColor       =   -2147483633
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
      Begin VB.CommandButton cmd_CamCon 
         Height          =   585
         Left            =   30
         Picture         =   "OpeTra_frm_008.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_Salida 
         Height          =   585
         Left            =   630
         Picture         =   "OpeTra_frm_008.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir"
         Top             =   30
         Width           =   585
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   3
      Top             =   7845
      Width           =   9825
      _Version        =   65536
      _ExtentX        =   17330
      _ExtentY        =   688
      _StockProps     =   15
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelOuter      =   0
      BevelInner      =   1
      Begin Threed.SSPanel pnl_EntDat 
         Height          =   315
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   3900
         _Version        =   65536
         _ExtentX        =   6879
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "lm_db_db1 - prod1"
         ForeColor       =   32768
         BackColor       =   -2147483633
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
      End
      Begin Threed.SSPanel pnl_NumVer 
         Height          =   315
         Left            =   3960
         TabIndex        =   5
         Top             =   30
         Width           =   2100
         _Version        =   65536
         _ExtentX        =   3704
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "rev. 008-1028.1"
         ForeColor       =   32768
         BackColor       =   -2147483633
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
      End
   End
   Begin VB.Menu mnuSol 
      Caption         =   "Trámites Solicitud Crédito Hipotecario"
      Begin VB.Menu mnuSol_Opcion 
         Caption         =   "Asignación de Gastos de Cierre"
         Index           =   1
      End
      Begin VB.Menu mnuSol_Opcion 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuSol_Opcion 
         Caption         =   "Tasación del Inmueble"
         Index           =   3
      End
      Begin VB.Menu mnuSol_Opcion 
         Caption         =   "Evaluación de Seguros"
         Index           =   4
      End
      Begin VB.Menu mnuSol_Opcion 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuSol_Opcion 
         Caption         =   "Pólizas de Seguro"
         Index           =   6
      End
      Begin VB.Menu mnuSol_Opcion 
         Caption         =   "Trámites COFIDE"
         Index           =   7
      End
      Begin VB.Menu mnuSol_Opcion 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuSol_Opcion 
         Caption         =   "Autorización de Desembolso"
         Index           =   9
      End
      Begin VB.Menu mnuSol_Opcion 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuSol_Opcion 
         Caption         =   "Rechazo Administrativo"
         Index           =   11
      End
      Begin VB.Menu mnuSol_Opcion 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuSol_Opcion 
         Caption         =   "Levantamiento Aprobación Condicionada (Tasación del Inmueble)"
         Index           =   13
      End
      Begin VB.Menu mnuSol_Opcion 
         Caption         =   "Levantamiento Aprobación Condicionada (Evaluación de Seguros)"
         Index           =   14
      End
      Begin VB.Menu mnuSol_Opcion 
         Caption         =   "Levantamiento Aprobación Condicionada (Pólizas de Seguros)"
         Index           =   15
      End
      Begin VB.Menu mnuSol_Opcion 
         Caption         =   "Levantamiento Aprobación Condicionada (Trámites Cofide)"
         Index           =   16
      End
      Begin VB.Menu mnuSol_Opcion 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu mnuSol_Opcion 
         Caption         =   "Modificación de Solicitud de Crédito"
         Index           =   18
      End
   End
   Begin VB.Menu mnuHip 
      Caption         =   "Créditos Hipotecarios"
      Begin VB.Menu mnuHip_Opcion 
         Caption         =   "Activación de Crédito Hipotecario"
         Index           =   1
      End
      Begin VB.Menu mnuHip_Opcion 
         Caption         =   "Desembolso de Crédito Hipotecario"
         Index           =   2
      End
      Begin VB.Menu mnuHip_Opcion 
         Caption         =   "Activar Seguro Inmueble de Crédito"
         Index           =   3
      End
      Begin VB.Menu mnuHip_Opcion 
         Caption         =   "Endosar Seguro de Desgravamen"
         Index           =   4
      End
      Begin VB.Menu mnuHip_Opcion 
         Caption         =   "Cambio de Fecha de Pago"
         Index           =   5
      End
      Begin VB.Menu mnuHip_Opcion 
         Caption         =   "Generación Carta Constitución Hipotecas"
         Index           =   6
      End
      Begin VB.Menu mnuHip_Opcion 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuHip_Opcion 
         Caption         =   "Simulación de Crédito Hipotecario"
         Index           =   8
      End
      Begin VB.Menu mnuHip_Opcion 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuHip_Opcion 
         Caption         =   "Gestión Operativa de Crédito Hipotecario"
         Index           =   10
      End
      Begin VB.Menu mnuHip_Opcion 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuHip_Opcion 
         Caption         =   "Posición Consolidada del Cliente"
         Index           =   12
      End
   End
   Begin VB.Menu mnuOpe 
      Caption         =   "Operaciones Financieras"
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "Carga de Archivo de Recaudación"
         Index           =   1
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "Generación de Archivo de Recaudo"
         Index           =   2
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "Carga de Cronogramas Tipo 1, 2, 3, 4, 5"
         Index           =   4
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "Carga de Archivos"
         Index           =   5
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "Carga de Archivo Conciliación de Pagos Mensuales Cofide "
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "Operaciones - Maestro"
         Index           =   7
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "Operaciones - Cuadre Cierre"
         Index           =   8
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "Operaciones - Adjudicados"
         Index           =   9
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "Créditos Hipotecarios - Cobro de Gastos de Cierre "
         Index           =   11
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "Créditos Hipotecarios - Transferencia de Gastos de Cierre"
         Index           =   12
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "Créditos Hipotecarios - Devolución de Gastos de Cierre"
         Index           =   13
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "Créditos Hipotecarios - Pago Proveedores de Gastos de Cierre"
         Index           =   14
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "Créditos Hipotecarios - Cobro de Cuotas"
         Index           =   15
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "-"
         Index           =   16
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "Créditos Hipotecarios - Prepago Parcial"
         Index           =   17
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "Créditos Hipotecarios - Prepago Total"
         Index           =   18
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "Créditos Hipotecarios - Seguimiento de Prepagos"
         Index           =   19
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "Créditos Hipotecarios - Consulta de Prepagos"
         Index           =   20
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "-"
         Index           =   21
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "Créditos Hipotecarios - Actualización de Tasación"
         Index           =   22
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "Créditos Hipotecarios - Proceso Actualización de Garantías"
         Index           =   23
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "Créditos Hipotecarios - Consulta de Garantías Actualizadas"
         Index           =   24
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "-"
         Index           =   25
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "Desembolso al Promotor"
         Index           =   26
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "Cuentas por Pagar"
         Index           =   27
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "Techo Propio"
         Index           =   28
      End
   End
   Begin VB.Menu mnuMnt 
      Caption         =   "Mantenimiento"
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Mantenimiento de Empresas de Peritaje"
         Index           =   1
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Mantenimiento de Empresas de Seguros"
         Index           =   2
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Mantenimiento de Notarias"
         Index           =   3
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Mantenimiento de Comisiones Mivivienda - Cofide"
         Index           =   5
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Mantenimiento de Días Feriados"
         Index           =   6
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Países"
         Index           =   8
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Ciudades Extranjeras"
         Index           =   9
      End
   End
   Begin VB.Menu mnuCon 
      Caption         =   "Consultas"
      Begin VB.Menu mnuCon_Opcion 
         Caption         =   "Operaciones Financieras"
         Index           =   1
      End
      Begin VB.Menu mnuCon_Opcion 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuCon_Opcion 
         Caption         =   "Tipo de Cambio Comercial"
         Index           =   3
      End
   End
   Begin VB.Menu mnuRpt 
      Caption         =   "Reportes"
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Operaciones Financieras "
         Index           =   1
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Procesos Operativos"
         Index           =   2
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Asignación de PBP"
         Index           =   3
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte General de Cartas Fianzas"
         Index           =   5
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Cartas Fianzas Pendiente de Regularización"
         Index           =   6
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Saldos Garantías (Cartas Fianzas)"
         Index           =   7
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Pagos a Mivivienda"
         Index           =   9
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Pagos de Seguros"
         Index           =   10
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Pagos Mensuales (COFIDE)"
         Index           =   12
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Saldos por Pagar (COFIDE)"
         Index           =   13
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Comparativo de Saldos miCasita-Cofide"
         Index           =   14
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Comparativo de Pagos Mensuales miCasita-Cofide"
         Index           =   15
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "-"
         Index           =   16
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Créditos Desembolsados (Mensual)"
         Index           =   17
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Saldos de Creditos Hipotecarios (Mensual)"
         Index           =   18
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Gastos de Cierre"
         Index           =   19
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Clasificacion de Cartera"
         Index           =   20
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reportes Varios"
         Index           =   21
      End
   End
   Begin VB.Menu mnuTrm 
      Caption         =   "Créditos Hipotecarios"
      Visible         =   0   'False
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "Simulación de Créditos Hipotecarios"
         Index           =   1
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "Asignación de Gastos de Cierre"
         Index           =   3
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "Reintegro de Gastos de Cierre"
         Index           =   4
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "Devolución de Gastos de Cierre"
         Index           =   5
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "Tasación del Inmueble"
         Index           =   7
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "Evaluación de Seguros"
         Index           =   8
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "Pólizas de Seguro"
         Index           =   10
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "Trámites COFIDE"
         Index           =   11
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "Autorización de Desembolso"
         Index           =   13
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "Activación de Crédito Hipotecario"
         Index           =   15
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "Desembolso de Crédito Hipotecario"
         Index           =   16
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "Rechazo Administrativo"
         Index           =   18
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "-"
         Index           =   19
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "Levantamiento de Aprobación Condicionada en Tasación del Inmueble"
         Index           =   20
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "Levantamiento de Aprobación Condicionada en Evaluación de Seguros"
         Index           =   21
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "Levantamiento de Aprobación Condicionada en Pólizas de Seguros"
         Index           =   22
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "Levantamiento de Aprobación Condicionada en Trámites Mivivienda-Cofide"
         Index           =   23
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "-"
         Index           =   24
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "Modificación de Solicitud"
         Index           =   25
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "-"
         Index           =   26
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "Gestión Operativa de Crédito"
         Index           =   27
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "-"
         Index           =   28
      End
      Begin VB.Menu mnuTrm_Opcion 
         Caption         =   "Simulación Operativa de Créditos Hipotecarios"
         Index           =   29
      End
   End
   Begin VB.Menu mnuCns 
      Caption         =   "Consultas"
      Visible         =   0   'False
      Begin VB.Menu mnuCns_Opcion 
         Caption         =   "Consulta de Solicitud de Crédito Hipotecario"
         Index           =   1
      End
      Begin VB.Menu mnuCns_Opcion 
         Caption         =   "Consulta de Crédito Hipotecario"
         Index           =   2
      End
      Begin VB.Menu mnuCns_Opcion 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuCns_Opcion 
         Caption         =   "Consulta de Tipo de Cambio"
         Index           =   4
      End
      Begin VB.Menu mnuCns_Opcion 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuCns_Opcion 
         Caption         =   "Consulta de Operaciones Financieras"
         Index           =   6
      End
   End
   Begin VB.Menu mnuPro 
      Caption         =   "Procesos"
      Begin VB.Menu mnuPro_Opcion 
         Caption         =   "Evaluación y Asignación de Premio Buen Pagador"
         Index           =   1
      End
      Begin VB.Menu mnuPro_Opcion 
         Caption         =   "Informe de Pagos a Mivivienda"
         Index           =   2
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frm_MnuPri_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_CamCon_Click()
   If modgen_g_str_CodUsu <> "DESARROLLO" Then
      frm_IdeUsu_02.Show 1
   End If
End Sub

Private Sub cmd_Salida_Click()
   If MsgBox("¿Está seguro de salir de la Plataforma?", vbExclamation + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
      Call gs_Desconecta_Servidor
      End
   End If
End Sub

Private Sub MDIForm_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   'Call fs_HabSeg
   Call moddat_gf_Cargar_AgrPrd
   Screen.MousePointer = 0
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   If MsgBox("¿Está seguro de salir de la Plataforma?", vbExclamation + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
      Call gs_Desconecta_Servidor
      End
   Else
      Cancel = True
   End If
End Sub

Private Sub mnuSol_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1
         'Asignar Gastos de Cierre
         frm_GasAdm_01.Show 1
         
      Case 3
         'Tasacion de inmueble
         frm_Tra_EvaTas_01.Show 1
         
      Case 4
         'Seguro de inmueble
         frm_Tra_EvaSeg_01.Show 1
         
      Case 6
         'Polizas de seguro
         frm_Tra_PolSeg_01.Show 1
         
      Case 7
         'Tramites Cofide
         frm_Tra_TraCof_01.Show 1
         
      Case 9
         'Autoreizacion de desembolso
         frm_Tra_AutDes_01.Show 1
         
      Case 11
         'Rechazo Administrativo
         frm_Hip_RecAdm_01.Show 1
         
      Case 13
         'Levantamiento Aprobacion Condicionada (Tasacion de inmueble)
         frm_Lev_TasInm_01.Show 1
         
      Case 14
         'Levantamiento Aprobacion Condicionada (Seguro de inmueble)
         frm_Lev_EvaSeg_01.Show 1
         
      Case 15
         'Levantamiento Aprobacion Condicionada (Polizas de seguro)
         frm_Lev_PolSeg_01.Show 1
         
      Case 16
         'Levantamiento Aprobacion Condicionada (Tramites Cofide)
         frm_Lev_TraCof_01.Show 1
         
      Case 18
         'Modificacion de Solicitud de Credito
         frm_ModSol_01.Show 1
   End Select
End Sub

Private Sub mnuHip_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1
         'Activacion del credito hipotecario
         frm_ActOpe_01.Show 1
         
      Case 2
         'Desembolso del credito hipotecario
         frm_Desemb_03.Show 1
         
      Case 3
         'Activar Seguro del inmueble
         frm_Pro_AsgSegInm_01.Show 1
         
      Case 4
         'Endosar Seguro Desgravamen
         moddat_g_int_FlgPre = 5
         frm_Con_PrePgo_01.Show 1
         
      Case 5
         'Cambio de Fecha de Pago
         moddat_g_int_FlgPre = 6
         frm_Con_PrePgo_01.Show 1
         
      Case 6
         'Constitucion de HIpotecas
         frm_Rpt_Cofide_05.Show 1
         
      Case 8
         'Simulacion del credito hipotecario
         frm_SimCre_11.Show 1
         
      Case 10
         'Gestion operatica del credito hipotecario
         moddat_g_int_FlgCre = 1
         frm_Ges_CreHip_01.Show 1
         
      Case 12
         'Posicion consolidada del cliente
         moddat_g_int_FlgCre = 2
         frm_Ges_CreHip_01.Show 1
   End Select
End Sub

Private Sub mnuOpe_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1
         'Carga de Archivo de Recaudo
         frm_Caj_CarArc_01.Show 1
         
      Case 2
         'Generación de Archivo de Recaudo
         frm_Caj_GenArc_01.Show 1
         
      Case 4
         'Carga de Cronogramas Tipo 1, 2, 3, 4
         moddat_g_int_FlgPre = 4
         frm_Con_PrePgo_01.Show 1
         
      Case 5
         'Carga archivo de saldos trimestrales cofide
         frm_Pro_SalCof_01.Show 1
         
      Case 6
         'Carga archivo de pagos mensuales cofide
         'frm_Pro_CbzCof.Show 1
         
      Case 7
         'Operaciones - Maestro
         frm_Con_Cuadre_02.Show 1
         
      Case 8
         'Operaciones - Cuadre Cierre
         frm_Con_Cuadre_01.Show 1
         
      Case 9
         'Operaciones - Adjudicados
         frm_Con_Cuadre_04.Show 1
         
      Case 11
         'Cobro de Gastos de Cierre (Credito Hipotecario)
         frm_Caj_GasCie_01.Show 1
         
      Case 12
         'Traslado de Gastos de Cierre (Crédito Hipotecario)
         frm_Caj_GasCie_03.Show 1
         
      Case 13
         'Devolución de Gastos de Cierre (Crédito Hipotecario)
         
      Case 14
         'Pago de Gastos de Cierre (Credito Hipotecario)
         frm_Caj_CiePag_01.Show 1
         
      Case 15
         'Cobro de Cuotas (Crédito Hipotecario)
         moddat_g_int_TipPan = 1
         frm_Caj_CreHip_01.Show 1
         
      Case 17
         'Prepago Parcial
         moddat_g_int_FlgPre = 1
         frm_Con_PrePgo_01.Show 1
         
      Case 18
         'Prepago Total
         moddat_g_int_FlgPre = 2
         frm_Con_PrePgo_01.Show 1
         
      Case 19
         'Seguimiento de prepagos
         frm_Con_PreSeg_01.Show 1
         
      Case 20
         'Consulta de prepagos
         frm_Con_PreCon_01.Show 1
      
      Case 22
         'Actualizacion de Tasacion
         moddat_g_int_FlgPre = 3
         frm_Con_PrePgo_01.Show 1
         
      Case 23
         'Proceso de Actualizacion de Garantias
         moddat_g_int_FlgPre = 3
         frm_Tas_ActReg_04.Show 1
         
      Case 24
         'Consulta de Garantias Actualizadas
         frm_Tas_ActReg_05.Show 1
         
      Case 26
         'Desembolso al Promotor
         frm_RegDes_01.Show 1
   
      Case 27
         'Cuentas por Pagar
         frm_Con_CtaPag_01.Show 1
   
      Case 28
         'Cartas Fianzas - Techpo Propio AVN
         frm_Ges_TecPro_01.Show 1
   
   End Select
End Sub

Private Sub mnuMnt_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1
         'Empresas de Peritaje
         moddat_g_int_TipRec = 1
         frm_EmpPer_01.Show 1
         
      Case 2
         'Empresa de Seguros
         moddat_g_int_TipRec = 2
         frm_EmpPer_01.Show 1
         
      Case 3
         'Empresa de Seguros
         moddat_g_int_TipRec = 3
         frm_EmpPer_01.Show 1
         
      Case 4
         'Comisiones Mivivienda
         frm_ComMvi_01.Show 1
         
      Case 5
         'Días Feriados
         frm_DiaFer_01.Show 1
   End Select
End Sub

Private Sub mnuCon_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1
         'Consulta de Operaciones Financieras
         frm_Con_OpeFin_01.Show 1
         
      Case 3
         'Consulta de Tipo de Cambio Comercial
         frm_ConTCa_01.Show 1
   End Select
End Sub

Private Sub mnuRpt_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1
         'Operaciones financieras por sucursal
         frm_Rpt_OpeFin_01.Show 1
         
      Case 2
         'Operaciones financieras por tipo de movimiento
         frm_RptSol_01.Show 1
         
      Case 3
         'Reporte de Asignación de PBP
         frm_Rpt_MviCof_01.Show 1
         
      Case 5
         'Reporte general de cartas fianzas
         frm_RptFia_01.Show 1
         
      Case 6
         'Reporte de cartas fianzas pendientes de regularizacion
         frm_RptFia_02.Show 1
         
      Case 7
         'Reporte de saldos de garantias (cartas fianzas)
         frm_RptFia_03.Show 1
         
      Case 9
         'Reporte de pagos a mivivienda
         frm_Pro_MViPag_01.Show 1
         
      Case 10
         'Reporte de pagos a seguros
         frm_Pro_MViPag_02.Show 1
         
      Case 12
         'Pagos mensuales cofide
         frm_Rpt_Cofide_01.Show 1
         
      Case 13
         'Saldos a pagar cofide
         frm_Rpt_Cofide_02.Show 1
         
      Case 14
         'Reporte comparativo de saldos trimestrales cofide-micasita
         frm_Rpt_Cofide_03.Show 1
         
      Case 15
         'Reporte comparativo de pagos mensuales cofide-micasita
         frm_Rpt_Cofide_04.Show 1
         
      Case 17
         'Reporte de creditos desembolsados
         frm_Rpt_CreDes_01.Show 1
         
      Case 18
         'Reporte de saldos mensuales
         frm_Rpt_CreSal_01.Show 1
         
      Case 19
         'Reporte de gastos de cierre (pago a proveedores)
         frm_Rpt_PagPrv_01.Show 1
         
      Case 20
         'Reporte de Clasificacion de cartera
         frm_Rpt_ClaCar_01.Show 1
   
      Case 21
         'Reporte Varios
         frm_Rpt_RepVar_01.Show 1
   
   End Select
End Sub

Private Sub mnuPro_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1
         'Proceso de evaluacion y asignacion PBP
         frm_Pro_EvaPBP_01.Show 1
         
      Case 2
         'Informe de pagos a mivivienda
         frm_Pro_PagMvi_01.Show 1
   End Select
End Sub

'**********************************************
Private Sub fs_HabSeg()
Dim r_int_Posici     As Integer
Dim r_str_CodMen     As String
   
   'pnl_Seg_NomUsu.Caption = modgen_g_str_CodUsu
   pnl_NumVer.Caption = modgen_g_str_NumRev
   pnl_EntDat.Caption = moddat_g_str_NomEsq & " - " & UCase(moddat_g_str_EntDat)
   
   'Desactivando todas las opciones
   For r_int_Posici = 1 To mnuSol_Opcion.Count
      If mnuSol_Opcion(r_int_Posici).Caption <> "-" Then
         mnuSol_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici

   For r_int_Posici = 1 To mnuHip_Opcion.Count
      If mnuHip_Opcion(r_int_Posici).Caption <> "-" Then
         mnuHip_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   For r_int_Posici = 1 To mnuOpe_Opcion.Count
      If mnuOpe_Opcion(r_int_Posici).Caption <> "-" Then
         mnuOpe_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   For r_int_Posici = 1 To mnuMnt_Opcion.Count
      If mnuMnt_Opcion(r_int_Posici).Caption <> "-" Then
         mnuMnt_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   For r_int_Posici = 1 To mnuCon_Opcion.Count
      If mnuCon_Opcion(r_int_Posici).Caption <> "-" Then
         mnuCon_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   For r_int_Posici = 1 To mnuRpt_Opcion.Count
      If mnuRpt_Opcion(r_int_Posici).Caption <> "-" Then
         mnuRpt_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   For r_int_Posici = 1 To mnuTrm_Opcion.Count
      If mnuTrm_Opcion(r_int_Posici).Caption <> "-" Then
         mnuTrm_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   For r_int_Posici = 1 To mnuCns_Opcion.Count
      If mnuCns_Opcion(r_int_Posici).Caption <> "-" Then
         mnuCns_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   For r_int_Posici = 1 To mnuPro_Opcion.Count
      If mnuPro_Opcion(r_int_Posici).Caption <> "-" Then
         mnuPro_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   'Verificando si todas las Opciones están habilitadas
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM SEG_PLTOPC "
   g_str_Parame = g_str_Parame & " WHERE PLTOPC_CODPLT = '" & UCase(App.EXEName) & "' "
   g_str_Parame = g_str_Parame & "   AND PLTOPC_FLGMEN = 2 "
   g_str_Parame = g_str_Parame & " ORDER BY PLTOPC_CODMEN ASC, PLTOPC_CODSUB ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_CodMen = Trim(g_rst_Princi!PLTOPC_CODMEN)
         Do While Not g_rst_Princi.EOF And r_str_CodMen = Trim(g_rst_Princi!PLTOPC_CODMEN)
            Select Case r_str_CodMen
               Case "MNUSOL": mnuSol_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
               Case "MNUHIP": mnuHip_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
               Case "MNUOPE": mnuOpe_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
               Case "MNUMNT": mnuMnt_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
               Case "MNUCON": mnuCon_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
               Case "MNURPT": mnuRpt_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
               Case "MNUPRO": mnuPro_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
            End Select
            
            g_rst_Princi.MoveNext
            If g_rst_Princi.EOF Then
               Exit Do
            End If
         Loop
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Verificando por Plantilla de Acceso
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM SEG_PLTPLA "
   g_str_Parame = g_str_Parame & " WHERE PLTPLA_CODPLT = '" & UCase(App.EXEName) & "' "
   g_str_Parame = g_str_Parame & "   AND PLTPLA_TIPUSU = '" & CStr(modgen_g_int_TipUsu) & "' "
   g_str_Parame = g_str_Parame & " ORDER BY PLTPLA_CODMEN ASC, PLTPLA_CODSUB ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_CodMen = Trim(g_rst_Princi!PLTPLA_CODMEN)
         Do While Not g_rst_Princi.EOF And r_str_CodMen = Trim(g_rst_Princi!PLTPLA_CODMEN)
            Select Case r_str_CodMen
               Case "MNUSOL": mnuSol_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
               Case "MNUHIP": mnuHip_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
               Case "MNUOPE": mnuOpe_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
               Case "MNUMNT": mnuMnt_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
               Case "MNUCON": mnuCon_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
               Case "MNURPT": mnuRpt_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
               Case "MNUPRO": mnuPro_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
            End Select
            
            g_rst_Princi.MoveNext
            If g_rst_Princi.EOF Then
               Exit Do
            End If
         Loop
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   'Verificando por Personalización de Opciones
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM SEG_PLTUSU "
   g_str_Parame = g_str_Parame & " WHERE PLTUSU_CODUSU = '" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "   AND PLTUSU_CODPLT = '" & UCase(App.EXEName) & "' "
   g_str_Parame = g_str_Parame & " ORDER BY PLTUSU_CODMEN ASC, PLTUSU_CODSUB ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_CodMen = Trim(g_rst_Princi!PLTUSU_CODMEN)
         Do While Not g_rst_Princi.EOF And r_str_CodMen = Trim(g_rst_Princi!PLTUSU_CODMEN)
            Select Case r_str_CodMen
               Case "MNUSOL": mnuSol_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
               Case "MNUHIP": mnuHip_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
               Case "MNUOPE": mnuOpe_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
               Case "MNUMNT": mnuMnt_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
               Case "MNUCON": mnuCon_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
               Case "MNURPT": mnuRpt_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
               Case "MNUPRO": mnuPro_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
            End Select
            
            g_rst_Princi.MoveNext
            If g_rst_Princi.EOF Then
               Exit Do
            End If
         Loop
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub mnuTrm_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1   'Simulación de Créditos Hipotecarios
         frm_SimCre_11.Show 1
      
      Case 7   'Trámite Tasación del Inmueble
         frm_Tra_EvaTas_01.Show 1
      
      Case 8   'Trámite Evaluación de Seguros
         frm_Tra_EvaSeg_01.Show 1
      
      Case 10  'Trámite Pólizas de Seguros
         frm_Tra_PolSeg_01.Show 1
      
      Case 11  'Trámite Trámites COFIDE
         frm_Tra_TraCof_01.Show 1
         
      Case 13  'Trámite Autorización de Desembolso
         frm_Tra_AutDes_01.Show 1
         
      Case 15  'Trámite Activación de Desembolso
         frm_Tra_ActOpe_01.Show 1
         
      Case 16  'Trámite de Desembolso
         frm_Tra_Desemb_01.Show 1
                  
      Case 18  'Rechazo Administrativo
         frm_Hip_RecAdm_01.Show 1
         
      Case 20  'Levantamiento Aprobación Condicionada Tasación del Inmueble
         frm_Lev_TasInm_01.Show 1
         
      Case 21  'Levantamiento Aprobación Condicionada Evaluación de Seguros
         frm_Lev_EvaSeg_01.Show 1
         
      Case 22  'Levantamiento Aprobación Condicionada Pólizas de Seguros
         frm_Lev_PolSeg_01.Show 1
         
      Case 23  'Levantamiento Aprobación Condicionada Trámites Mivivienda-Cofide
         frm_Lev_TraCof_01.Show 1
         
      Case 29  'Simulación Operativa de Créditos Hipotecarios
         frm_Sim_OpeHip_01.Show 1
   End Select
End Sub
