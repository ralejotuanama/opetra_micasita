VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frm_Seg_SolHip_53 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9810
   ClientLeft      =   3015
   ClientTop       =   540
   ClientWidth     =   11580
   Icon            =   "OpeTra_frm_182.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9810
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9825
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
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
         Begin VB.CommandButton cmd_ConAFP 
            Height          =   585
            Left            =   2430
            Picture         =   "OpeTra_frm_182.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Pre-Conformidad AFP"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_NueObs 
            Height          =   585
            Left            =   1830
            Picture         =   "OpeTra_frm_182.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Enviar a Créditos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "OpeTra_frm_182.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir de la Ventana"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_VerGas 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_182.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Consulta de Gastos de Cierre"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DatInm 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_182.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Modificación de Datos del Inmueble"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DatCli 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_182.frx":1636
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Modificación de Datos del Cliente"
            Top             =   30
            Width           =   585
         End
         Begin Threed.SSPanel pnl_ObsAFP 
            Height          =   555
            Left            =   6900
            TabIndex        =   20
            Top             =   60
            Width           =   3795
            _Version        =   65536
            _ExtentX        =   6694
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "OBSERVACION PRE CONFORMIDAD AFP PENDIENTE DE DESCARGO"
            ForeColor       =   16777215
            BackColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Outline         =   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   7
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
            Left            =   690
            TabIndex        =   8
            Top             =   60
            Width           =   7725
            _Version        =   65536
            _ExtentX        =   13626
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Seguimiento de Solicitud de Crédito Hipotecario"
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
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   10560
            Top             =   30
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            AddressEditFieldCount=   1
            AddressModifiable=   0   'False
            AddressResolveUI=   0   'False
            FetchSorted     =   0   'False
            FetchUnreadOnly =   0   'False
         End
         Begin MSMAPI.MAPISession mps_Sesion 
            Left            =   9960
            Top             =   30
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_182.frx":1940
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel14 
         Height          =   4125
         Left            =   30
         TabIndex        =   9
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
            TabIndex        =   10
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
            TabIndex        =   11
            Top             =   60
            Width           =   3945
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   4155
         Left            =   30
         TabIndex        =   12
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
            TabIndex        =   13
            Top             =   630
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   6165
            _Version        =   393216
            Rows            =   21
            Cols            =   7
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel13 
            Height          =   285
            Left            =   6300
            TabIndex        =   14
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
            Left            =   8790
            TabIndex        =   15
            Top             =   330
            Width           =   2385
            _Version        =   65536
            _ExtentX        =   4207
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
            Left            =   4920
            TabIndex        =   16
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
            TabIndex        =   17
            Top             =   330
            Width           =   4845
            _Version        =   65536
            _ExtentX        =   8546
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
            Left            =   7680
            TabIndex        =   18
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
            TabIndex        =   19
            Top             =   60
            Width           =   3165
         End
      End
   End
End
Attribute VB_Name = "frm_Seg_SolHip_53"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_int_ObsPen        As Integer

Private Sub cmd_ConAFP_Click()
   If moddat_g_int_MtoAfp = 0 Then
      MsgBox "La solicitud no registra monto de AFP.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
      
   moddat_g_int_TipRep = 2
   frm_Tra_CarAFP_02.Show 1
End Sub

Private Sub cmd_DatCli_Click()
   If moddat_g_int_InsAct <> 11 Then
      MsgBox "La solicitud ha sido enviada a Evaluación Crediticia. No se pueden modificar los datos del Cliente. Coordine con el Dpto. de Créditos.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_int_FlgAct = 1
   moddat_g_int_FlgGrb = 2
   frm_MntCli_52.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call modmip_gs_DatSolCre(moddat_g_str_NumSol, grd_DatSol)
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_DatInm_Click()
   If moddat_g_int_InsAct >= 41 Then
      MsgBox "La información del Inmueble sólo puede ser modificada antes del envío a Tasación.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If moddat_g_int_InmIde = 1 Then
      moddat_g_int_FlgGrb = 2
   Else
      moddat_g_int_FlgGrb = 1
   End If
   
   moddat_g_int_FlgAct = 1
   frm_Seg_SolHip_54.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call modmip_gs_DatSolCre(moddat_g_str_NumSol, grd_DatSol)
      Screen.MousePointer = 0
      moddat_g_int_InmIde = 1
   End If
End Sub

Private Sub cmd_NueObs_Click()
Dim r_int_NumObs     As Integer
Dim r_str_Parame     As String
Dim r_rst_Genera     As ADODB.Recordset
Dim r_int_Resul      As Integer
Dim r_str_CodPry     As String
Dim r_str_CodMod     As String
Dim r_str_CodPrd     As String
Dim r_str_DesMod     As String

   '*********** Insertamos una auto-observación para que supervisor realice el descargo ************
   If moddat_g_int_FlgActEnv = 1 Then
      
      r_str_CodPry = ""
      r_str_CodMod = ""
      r_str_CodPrd = ""
      r_str_DesMod = ""
      
      r_str_Parame = ""
      r_str_Parame = r_str_Parame & "  SELECT SOLMAE_CODPRD, SOLMAE_CODMOD, SOLINM_PRYCOD "
      r_str_Parame = r_str_Parame & "    FROM CRE_SOLMAE "
      r_str_Parame = r_str_Parame & "         INNER JOIN CRE_SOLINM ON SOLINM_NUMSOL = SOLMAE_NUMERO "
      r_str_Parame = r_str_Parame & "   WHERE SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "'"
      
      If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
         r_rst_Genera.MoveFirst
         If Not IsNull(r_rst_Genera!SOLINM_PRYCOD) Then
            r_str_CodPry = Trim(r_rst_Genera!SOLINM_PRYCOD)
         End If
         r_str_CodMod = Trim(r_rst_Genera!SOLMAE_CODMOD)
         r_str_CodPrd = Trim(r_rst_Genera!SOLMAE_CODPRD)
         r_str_DesMod = moddat_gf_Buscar_NomMod(Trim(r_str_CodPrd), r_str_CodMod)
      End If
      
      If InStr(r_str_DesMod, "TERMINADO") = 0 Then  'Bien Terminado
      
         r_str_Parame = ""
         r_str_Parame = r_str_Parame & "SELECT NVL((SELECT COUNT(*) "
         r_str_Parame = r_str_Parame & "              FROM CRE_SOLINM "
         r_str_Parame = r_str_Parame & "             WHERE SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "'),0) AS CONTEO, "
         r_str_Parame = r_str_Parame & "       NVL((SELECT X.DATGEN_PRYAPR "
         r_str_Parame = r_str_Parame & "              FROM PRY_DATGEN X "
         r_str_Parame = r_str_Parame & "             WHERE DATGEN_CODIGO = (SELECT SOLINM_PRYCOD FROM CRE_SOLINM A "
         r_str_Parame = r_str_Parame & "                                     WHERE A.SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "')),0) AS PRYAPR "
         r_str_Parame = r_str_Parame & "  FROM DUAL "
      
         If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
            Exit Sub
         End If
         
         If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
            r_rst_Genera.MoveFirst
            If r_rst_Genera!CONTEO = 0 Then
               MsgBox "Esta pendiente por registrar el inmueble.", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
            If r_rst_Genera!PRYAPR <> 1 Then
               MsgBox "El proyecto no está aprobado coordinar con las áreas correspondientes.", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
         End If
         
         r_rst_Genera.Close
         Set r_rst_Genera = Nothing
         
      End If
      
      '=======================================================
      If InStr(r_str_DesMod, "TERMINADO") = 0 Then  'Bien Terminado
         'Valida los Gastos de Cierre
         r_int_Resul = gf_Valida_GastoCierre(r_str_CodPrd, r_str_CodPry)
         
         If r_int_Resul = 1 Then
            MsgBox "El proyecto asociado no tiene empresa de peritaje asignado, favor actualizar información en la plataforma de Operaciones.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         ElseIf r_int_Resul = 2 Then
            MsgBox "El proyecto asociado no tiene notaría asignada, favor actualizar información en la plataforma de Operaciones.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         ElseIf r_int_Resul = 3 Then
            'MsgBox "La notaria asociada al proyecto no tiene registrado los parámetros necesarios para el cálculo de los gastos de cierre, favor actualizar información en la plataforma de Operaciones.", vbExclamation, modgen_g_str_NomPlt
            'Exit Sub
            'MsgBox "Los gastos de cierre no se calcularán porque no se han registrado los parámetros de notaría, favor actualizar información en la plataforma de Operaciones.", vbExclamation, modgen_g_str_NomPlt
            If MsgBox("La notaria asociada al proyecto no tiene registrado los parámetros necesarios para el cálculo de los gastos de cierre, favor coordinar con el área legal la actualización de la información en caso contrario no se generaran los gastos de cierre." & vbCrLf & "Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               Exit Sub
            End If
         ElseIf r_int_Resul = 4 Then
            MsgBox "La empresa de peritaje asociada al proyecto no tiene registrado los parámetros necesarios para el cálculo de los gastos de cierre, favor actualizar información en la plataforma de Operaciones.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
      End If
      
      If MsgBox("¿Está seguro de Enviar?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
   
      r_int_NumObs = 0
      r_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
      r_str_Parame = r_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
      r_str_Parame = r_str_Parame & "SEGDET_CODINS = 11 AND "
      r_str_Parame = r_str_Parame & "SEGDET_CODOCU = 21 "
      r_str_Parame = r_str_Parame & "ORDER BY SEGDET_NUMOBS DESC"
   
      If Not gf_EjecutaSQL(r_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
         
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         Do While Not g_rst_Princi.EOF
            r_int_NumObs = r_int_NumObs + 1
            g_rst_Princi.MoveNext
         Loop
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      r_int_NumObs = r_int_NumObs + 1
      moddat_g_str_Observ = "PENDIENTE DE ENVÍO A RECEPCIÓN DE SOLICITUDES"
      
      'Grabando en Detalle de Seguimiento
      If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 11, 92, CStr(r_int_NumObs), moddat_g_str_Observ, 1, 0) Then
         Exit Sub
      End If
      
      'Actualizando en Instancia si es una Observación
      If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, 11, 0, 3, 2) Then
         Exit Sub
      End If
      
      moddat_g_int_NumObs = r_int_NumObs

      '*************************************************************************************************
      moddat_g_str_DesObs = "ENVIADO A RECEPCIÓN DE SOLICITUDES"
      moddat_g_int_FlgAct_1 = 2
   
      If moddat_g_int_FlgAct_1 = 2 Then
         If Not moddat_gf_Modifica_SegDet_Observ(moddat_g_str_NumSol, 11, 92, CStr(moddat_g_int_NumObs), moddat_g_str_DesObs, 2) Then
            Exit Sub
         End If
         
         'Actualizando en Instancia
         If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, 11, 0, 9, 2) Then
            Exit Sub
         End If
         
         'Enviando Correo Electrónico
         modgen_g_str_Mail_Asunto = moddat_gf_Consulta_ParDes("002", CStr(11)) & " - ENVÍO DE SOLICITUD A CRÉDITOS " & "(Cliente: " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " - " & moddat_g_str_NomCli & ")"
         
         modgen_g_str_Mail_Mensaj = ""
         modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & gf_Formato_NumSol(moddat_g_str_NumSol) & Chr(13)
         modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
         modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
         modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "FECHA               : " & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & Chr(13)
         modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "HORA                : " & Format(Time, "hh:mm:ss") & Chr(13)
         modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
         modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & moddat_g_str_DesObs
      
         Call fs_Envia_CorreoEle(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, moddat_g_str_NumSol, 11, True, False, False)
      
         Screen.MousePointer = 11
         Call fs_Buscar_Seguim
         Me.cmd_NueObs.Visible = False
         Screen.MousePointer = 0
         moddat_g_int_FlgAct = 2
      End If
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_VerGas_Click()
   frm_Seg_SolHip_56.Show 1
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call modmip_gs_DatSolCre(moddat_g_str_NumSol, grd_DatSol)
   Call fs_Buscar_Seguim
   Call fs_Buscar_ObsPenAFP
   
   'Si no hay Aprobaciones Condicionadas Pendiente
   If l_int_ObsPen = 0 Then
      pnl_ObsAFP.Visible = False
   End If
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
   grd_LisIns.ColWidth(0) = 4835
   grd_LisIns.ColWidth(1) = 1385
   grd_LisIns.ColWidth(2) = 1385
   grd_LisIns.ColWidth(3) = 1115
   grd_LisIns.ColWidth(4) = 2375
   grd_LisIns.ColWidth(5) = 0
   grd_LisIns.ColWidth(6) = 0
   grd_LisIns.ColAlignment(0) = flexAlignLeftCenter
   grd_LisIns.ColAlignment(1) = flexAlignCenterCenter
   grd_LisIns.ColAlignment(2) = flexAlignCenterCenter
   grd_LisIns.ColAlignment(3) = flexAlignRightCenter
   grd_LisIns.ColAlignment(4) = flexAlignLeftCenter
End Sub

Private Sub grd_DatSol_SelChange()
   If grd_DatSol.Rows > 2 Then
      grd_DatSol.RowSel = grd_DatSol.Row
   End If
End Sub

Private Sub grd_LisIns_DblClick()
Dim r_int_Situac     As Integer
   
   If moddat_g_int_FlgActEnv = 0 Then
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
         Case 11
            If r_int_Situac <> 1 And r_int_Situac <> 2 Then
               frm_Seg_SolHip_61.Show 1
            Else
               frm_Seg_SolHip_62.Show 1
            End If
            
         Case 21
            If r_int_Situac <> 1 And r_int_Situac <> 2 Then
               frm_Seg_SolHip_63.Show 1
            Else
               frm_Seg_SolHip_64.Show 1
            End If
            
         Case 31
            If r_int_Situac <> 1 And r_int_Situac <> 2 Then
               If moddat_g_int_TipMon <> 1 Then
                  If moddat_gf_Obtiene_TipCam(1, moddat_g_int_TipMon) = 0 Then
                     MsgBox "Debe solicitar el ingreso del Tipo de Cambio de " & moddat_gf_Consulta_ParDes("204", CStr(moddat_g_int_TipMon)) & ".", vbExclamation, modgen_g_str_NomPlt
                     Exit Sub
                  End If
               End If
               
               frm_Seg_SolHip_65.Show 1
            Else
               frm_Seg_SolHip_66.Show 1
            End If
         
         Case 32  'Trámites del Cliente
            If r_int_Situac <> 1 And r_int_Situac <> 2 Then
               frm_Seg_SolHip_67.Show 1
            Else
               frm_Seg_SolHip_68.Show 1
            End If
            
         Case 41
            If r_int_Situac <> 1 And r_int_Situac <> 2 Then
               frm_Seg_SolHip_69.Show 1
            Else
               frm_Seg_SolHip_70.Show 1
            End If
         
         Case 42
            If r_int_Situac <> 1 And r_int_Situac <> 2 Then
               frm_Seg_SolHip_71.Show 1
            Else
               frm_Seg_SolHip_72.Show 1
            End If
         
         Case 51
            If r_int_Situac <> 1 And r_int_Situac <> 2 Then
               frm_Seg_SolHip_73.Show 1
            Else
               frm_Seg_SolHip_74.Show 1
            End If
            
         Case 61
            If r_int_Situac <> 1 And r_int_Situac <> 2 Then
               frm_Seg_SolHip_75.Show 1
            Else
               frm_Seg_SolHip_76.Show 1
            End If
         
         Case 62
            If r_int_Situac <> 1 And r_int_Situac <> 2 Then
               If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then   '"003" "004" "007" "009" "010" "012" "013" "014" "015" "016" "017" "018" "019" "021" "022" "023"
                  frm_Seg_SolHip_77.Show 1
               End If
            Else
               If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then   '"003" "004" "007" "009" "010" "012" "013" "014" "015" "016" "017" "018" "019" "021" "022" "023"
                  frm_Seg_SolHip_78.Show 1
               Else
                  frm_Seg_SolHip_79.Show 1
               End If
            End If
            
         Case 72
            If r_int_Situac <> 1 And r_int_Situac <> 2 Then
               frm_Seg_SolHip_80.Show 1
            Else
               frm_Seg_SolHip_81.Show 1
            End If
            
         Case 81
            frm_Seg_SolHip_82.Show 1
      End Select
      
      If moddat_g_int_FlgAct = 2 Then
         Screen.MousePointer = 11
         Call modmip_gs_DatSolCre(moddat_g_str_NumSol, grd_DatSol)
         Call fs_Buscar_Seguim
         Screen.MousePointer = 0
      End If
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
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * "
      g_str_Parame = g_str_Parame & "  FROM TRA_SEGDET "
      g_str_Parame = g_str_Parame & " WHERE SEGDET_NUMSOL = '" & p_NumSol & "' "
      g_str_Parame = g_str_Parame & "   AND SEGDET_CODOCU = 13 "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Function
      End If
   
      DoEvents
      g_rst_Genera.MoveFirst
      ff_ObsRec = Trim(g_rst_Genera!SEGDET_OBSERV & "")
   
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      
   ElseIf p_TipRec = 3 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * "
      g_str_Parame = g_str_Parame & "  FROM TRA_RECADM "
      g_str_Parame = g_str_Parame & " WHERE RECADM_NUMSOL = '" & p_NumSol & "' "
   
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

Private Sub fs_Buscar_ObsPenAFP()
   l_int_ObsPen = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.SEGFECCRE, A.SEGHORCRE, A.SEGDET_CODOCU, A.SEGFECACT, A.SEGHORACT, A.SEGDET_OBSERV, A.SEGDET_OBSDES, A.SEGDET_NUMOBS "
   g_str_Parame = g_str_Parame & "   FROM TRA_SEGDET A "
   g_str_Parame = g_str_Parame & "  WHERE A.SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & "    AND A.SEGDET_CODINS = 33 " '" & moddat_g_int_CodIns & " "
   g_str_Parame = g_str_Parame & "    AND A.SEGDET_CODOCU = 99 "
   g_str_Parame = g_str_Parame & "  ORDER BY A.SEGFECCRE DESC, A.SEGHORCRE DESC "
   
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
      If g_rst_Princi!SEGFECACT = 0 Then
         l_int_ObsPen = 1
      End If
  
      g_rst_Princi.MoveNext
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub
