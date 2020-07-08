VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_Tra_TraCof_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   8475
   ClientLeft      =   1200
   ClientTop       =   2505
   ClientWidth     =   14130
   Icon            =   "OpeTra_frm_287.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   14130
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8460
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14145
      _Version        =   65536
      _ExtentX        =   24950
      _ExtentY        =   14922
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
         TabIndex        =   1
         Top             =   780
         Width           =   14025
         _Version        =   65536
         _ExtentX        =   24739
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
            Left            =   630
            Picture         =   "OpeTra_frm_287.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_EvaSol 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_287.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   13410
            Picture         =   "OpeTra_frm_287.frx":0BE0
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   6015
         Left            =   60
         TabIndex        =   4
         Top             =   1470
         Width           =   14025
         _Version        =   65536
         _ExtentX        =   24739
         _ExtentY        =   10610
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
            Height          =   5985
            Left            =   30
            TabIndex        =   5
            Top             =   30
            Width           =   13935
            _ExtentX        =   24580
            _ExtentY        =   10557
            _Version        =   393216
            Rows            =   30
            Cols            =   13
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            SelectionMode   =   1
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   6
         Top             =   60
         Width           =   14025
         _Version        =   65536
         _ExtentX        =   24739
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
         Begin Threed.SSPanel pnl_TitPri 
            Height          =   315
            Left            =   690
            TabIndex        =   7
            Top             =   30
            Width           =   8565
            _Version        =   65536
            _ExtentX        =   15108
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Solicitud de Crédito Hipotecario"
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
         Begin Threed.SSPanel pnl_TitSec 
            Height          =   315
            Left            =   690
            TabIndex        =   8
            Top             =   330
            Width           =   8565
            _Version        =   65536
            _ExtentX        =   15108
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Trámites COFIDE"
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
            Picture         =   "OpeTra_frm_287.frx":1022
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   900
         Left            =   60
         TabIndex        =   10
         Top             =   7500
         Width           =   14025
         _Version        =   65536
         _ExtentX        =   24739
         _ExtentY        =   1579
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   315
            Left            =   240
            TabIndex        =   11
            Top             =   120
            Width           =   1845
            _Version        =   65536
            _ExtentX        =   3254
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "ENVIADO A COFIDE"
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Left            =   240
            TabIndex        =   12
            Top             =   480
            Width           =   2085
            _Version        =   65536
            _ExtentX        =   3678
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "NO ENVIADO A COFIDE"
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
         Begin Threed.SSPanel pnl_SenCof 
            Height          =   315
            Left            =   2880
            TabIndex        =   13
            Top             =   480
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_EnvCof 
            Height          =   315
            Left            =   2880
            TabIndex        =   14
            Top             =   120
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
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
            Alignment       =   4
         End
      End
   End
End
Attribute VB_Name = "frm_Tra_TraCof_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_EvaSol_Click()
   If grd_Listad.Rows = 1 Then
      Exit Sub
   End If

   grd_Listad.Col = 0
   moddat_g_str_NomPrd = grd_Listad.Text

   grd_Listad.Col = 1
   moddat_g_str_NumSol = Mid(grd_Listad.Text, 1, 3) & Mid(grd_Listad.Text, 5, 3) & Mid(grd_Listad.Text, 9, 2) & Mid(grd_Listad.Text, 12, 4)
   
   grd_Listad.Col = 2
   moddat_g_int_TipDoc = Left(grd_Listad.Text, 1)
   moddat_g_str_NumDoc = Mid(grd_Listad.Text, 3)
         
   grd_Listad.Col = 3
   moddat_g_str_NomCli = grd_Listad.Text
   
   grd_Listad.Col = 4
   moddat_g_str_FecIng = grd_Listad.Text
   
   grd_Listad.Col = 11
   moddat_g_str_CodPrd = grd_Listad.Text
   
   grd_Listad.Col = 12
   moddat_g_str_CodSub = grd_Listad.Text
   
   grd_Listad.Col = 15
   moddat_g_int_TipMon = CInt(grd_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Listad)
   
   If moddat_g_int_TipMon <> 1 Then
      If moddat_gf_Obtiene_TipCam(1, moddat_g_int_TipMon) = 0 Then
         MsgBox "No se encontró Tipo de Cambio registrado para " & moddat_gf_Consulta_ParDes("204", CStr(moddat_g_int_TipMon)) & ".", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   End If
   
   moddat_g_int_FlgAct = 1
   frm_Tra_TraCof_02.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_ExpExc_Click()
   If grd_Listad.Rows = 0 Then
      MsgBox "No existe datos.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
       
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
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
   grd_Listad.Cols = 21 '19
   grd_Listad.ColWidth(0) = 1895
   grd_Listad.ColWidth(1) = 1375
   grd_Listad.ColWidth(2) = 1235
   grd_Listad.ColWidth(3) = 3455
   grd_Listad.ColWidth(4) = 1195
   grd_Listad.ColWidth(5) = 1195
   grd_Listad.ColWidth(6) = 1670
   grd_Listad.ColWidth(7) = 1670
   grd_Listad.ColWidth(8) = 1650 '1580
   grd_Listad.ColWidth(9) = 0
   grd_Listad.ColWidth(10) = 1580
   grd_Listad.ColWidth(11) = 0
   grd_Listad.ColWidth(12) = 0
   grd_Listad.ColWidth(13) = 0
   grd_Listad.ColWidth(14) = 0
   grd_Listad.ColWidth(15) = 0
   grd_Listad.ColWidth(16) = 0
   grd_Listad.ColWidth(17) = 0
   grd_Listad.ColWidth(18) = 0
   grd_Listad.ColWidth(19) = 0
   grd_Listad.ColWidth(20) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignCenterCenter
   grd_Listad.ColAlignment(7) = flexAlignCenterCenter
   grd_Listad.ColAlignment(8) = flexAlignCenterCenter
   grd_Listad.ColAlignment(10) = flexAlignCenterCenter
End Sub

Private Sub fs_Buscar()
Dim r_int_EnvCof     As Integer
Dim r_int_NEnCof     As Integer
   
   r_int_EnvCof = 0
   r_int_NEnCof = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT SOLMAE_NUMERO, PRODUC_DESCRI, SOLMAE_TITTDO, SOLMAE_TITNDO, SOLMAE_FECSOL, SOLMAE_CONHIP, SOLMAE_CODPRD, "
   g_str_Parame = g_str_Parame & "       SOLMAE_CODSUB, SOLMAE_FECSOL, SOLMAE_TIPMON, D.SEGUIM_FECINI, D.SEGUIM_SITUAC, D.SEGUIM_FECINI, "
   g_str_Parame = g_str_Parame & "       (SELECT X.EVACOF_FECENV FROM TRA_EVACOF X WHERE X.EVACOF_NUMSOL = A.SOLMAE_NUMERO) AS FECHA_ENVIO_COF, "
   g_str_Parame = g_str_Parame & "       DATGEN_NOMBRE, DATGEN_APEPAT, DATGEN_APEMAT, DATGEN_APECAS, TRIM(E.PARDES_DESCRI) AS SITUACION, "
   g_str_Parame = g_str_Parame & "       DECODE((SELECT LENGTH(TRIM(SEGDET_CODOCU)) FROM TRA_SEGDET "
   g_str_Parame = g_str_Parame & "                WHERE SEGDET_NUMSOL = A.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & "                  AND SEGDET_CODINS = 62 AND SEGDET_CODOCU = 36),NULL,'NO','SI') AS ENVIO_COFIDE "
   g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_PRODUC B ON PRODUC_CODIGO = SOLMAE_CODPRD "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN C ON DATGEN_TIPDOC = SOLMAE_TITTDO AND DATGEN_NUMDOC = SOLMAE_TITNDO "
   g_str_Parame = g_str_Parame & " INNER JOIN TRA_SEGUIM D ON D.SEGUIM_NUMSOL = SOLMAE_NUMERO AND SEGUIM_CODINS = 62 AND (SEGUIM_SITUAC = 9 OR SEGUIM_SITUAC = 3) "
   'g_str_Parame = g_str_Parame & " INNER JOIN TRA_SEGUIM D ON D.SEGUIM_NUMSOL = SOLMAE_NUMERO AND SEGUIM_CODINS = 62 "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 23 AND E.PARDES_CODITE = D.SEGUIM_SITUAC "
   g_str_Parame = g_str_Parame & " WHERE SOLMAE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "   AND SOLMAE_CODINS = 61 "
   g_str_Parame = g_str_Parame & "   AND SOLMAE_CODPRD IN (" & moddat_g_str_AgrTFMV & ") "
   'g_str_Parame = g_str_Parame & " WHERE SOLMAE_NUMERO IN ('022001191655','022001191149','022001191431','022001191470','022001191506','022001191370','022001191354','022001191243','022001191469','022001191462','022001191267','022001191565','022001191649','022001191698','022001191653','022001191294') "
   g_str_Parame = g_str_Parame & " ORDER BY SOLMAE_NUMERO ASC "
 
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   Call gs_LimpiaGrid(grd_Listad)
   
   'CABECERA
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.FixedRows = 1

   grd_Listad.Row = 0
   grd_Listad.Col = 0:   grd_Listad.Text = "Producto":               grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 1:   grd_Listad.Text = "Nro. Solicitud":         grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 2:   grd_Listad.Text = "DOI Cliente":            grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 3:   grd_Listad.Text = "Apellidos y Nombres":    grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 4:   grd_Listad.Text = "F. Solicitud":           grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 5:   grd_Listad.Text = "F. Ing. Inst.":          grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 6:   grd_Listad.Text = "Situación Instancia":    grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 7:   grd_Listad.Text = "Enviado a COFIDE":       grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 8:   grd_Listad.Text = "Fecha Envio COFIDE":     grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 10:   grd_Listad.Text = "Consej. Hipotecario":    grd_Listad.CellAlignment = flexAlignCenterCenter
   
   grd_Listad.Rows = grd_Listad.Rows - 1
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad.Redraw = False
      
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = Trim(g_rst_Princi!PRODUC_DESCRI)
         
         grd_Listad.Col = 1
         grd_Listad.Text = Left(g_rst_Princi!SOLMAE_NUMERO, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 4, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 7, 2) & "-" & Right(g_rst_Princi!SOLMAE_NUMERO, 4)
         
         grd_Listad.Col = 2
         grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
         
         grd_Listad.Col = 3
         grd_Listad.Text = Trim(g_rst_Princi!DatGen_ApePat) & " " & Trim(g_rst_Princi!DatGen_ApeMat) & IIf(Len(Trim(g_rst_Princi!DATGEN_APECAS)) > 0, " DE " & Trim(g_rst_Princi!DATGEN_APECAS), "") & " " & Trim(g_rst_Princi!DatGen_Nombre)
         
         grd_Listad.Col = 4
         grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
         
         grd_Listad.Col = 5
         grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECINI))
         
         grd_Listad.Col = 6
         grd_Listad.Text = g_rst_Princi!SITUACION
         
         grd_Listad.Col = 7
         grd_Listad.Text = g_rst_Princi!ENVIO_COFIDE
         
         If grd_Listad.Text = "SI" Then
            r_int_EnvCof = r_int_EnvCof + 1
         ElseIf grd_Listad.Text = "NO" Then
            r_int_NEnCof = r_int_NEnCof + 1
         End If
         
         If Len(Trim(g_rst_Princi!FECHA_ENVIO_COF & "")) > 3 Then
            grd_Listad.Col = 8
            grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!FECHA_ENVIO_COF))
            
            grd_Listad.Col = 9
            grd_Listad.Text = g_rst_Princi!FECHA_ENVIO_COF
         End If
         
         grd_Listad.Col = 10
         grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CONHIP & "")
         
         grd_Listad.Col = 11
         grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CODPRD & "")
         
         grd_Listad.Col = 12
         grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CODSUB & "")
         
         grd_Listad.Col = 13
         grd_Listad.Text = g_rst_Princi!SOLMAE_FECSOL
         
         grd_Listad.Col = 14
         grd_Listad.Text = g_rst_Princi!SEGUIM_FECINI
         
         grd_Listad.Col = 15
         grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_TIPMON)
         
         grd_Listad.Col = 16
         grd_Listad.Text = CStr(g_rst_Princi!DatGen_ApePat)
         
         grd_Listad.Col = 17
         grd_Listad.Text = CStr(g_rst_Princi!DatGen_ApeMat)
         
         grd_Listad.Col = 18
         grd_Listad.Text = CStr(g_rst_Princi!DatGen_Nombre)
         
         grd_Listad.Col = 19
         grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_NUMERO)
         
         grd_Listad.Col = 20
         grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_TITTDO) & Trim(g_rst_Princi!SOLMAE_TITNDO)
         
         g_rst_Princi.MoveNext
      Loop
      grd_Listad.Redraw = True
      
      If grd_Listad.Rows > 0 Then
         'Ordenando por Nombre de Cliente
         Call gs_SorteaGrid(grd_Listad, 3, "C")
      End If
      If grd_Listad.Rows > 1 Then
         Call gs_UbicaGrid(grd_Listad, 1)
      End If
   Else
      cmd_EvaSol.Enabled = False
      MsgBox "No se encontraron Solicitudes registradas.", vbInformation, modgen_g_str_NomPlt
   End If
   
   pnl_EnvCof.Caption = r_int_EnvCof & " "
   pnl_SenCof.Caption = r_int_NEnCof & " "
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_NroFil     As Integer
Dim r_int_nroaux     As Integer
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   r_int_NroFil = 1
   
   With r_obj_Excel.ActiveSheet
      .Cells(r_int_NroFil, 1) = "FECHA":                    .Columns("A").ColumnWidth = 12
      .Cells(r_int_NroFil, 2) = "IFI":                      .Columns("B").ColumnWidth = 22
      .Cells(r_int_NroFil, 3) = "APELLIDO PATERNO":         .Columns("C").ColumnWidth = 20
      .Cells(r_int_NroFil, 4) = "APELLIDO MATERNO":         .Columns("D").ColumnWidth = 25
      .Cells(r_int_NroFil, 5) = "NOMBRES":                  .Columns("E").ColumnWidth = 25
      .Cells(r_int_NroFil, 6) = "NOMBRE DEL PRODUCTO":      .Columns("F").ColumnWidth = 50
      .Cells(r_int_NroFil, 7) = "ENVIADO A COFIDE":         .Columns("G").ColumnWidth = 20
      .Cells(r_int_NroFil, 8) = "FECHA ENVIADO COFIDE":     .Columns("H").ColumnWidth = 22
      .Cells(r_int_NroFil, 9) = "CONSEJ. HIPOTECARIO":      .Columns("I").ColumnWidth = 22
      .Cells(r_int_NroFil, 10) = "NUMERO SOLICITUD":        .Columns("J").ColumnWidth = 20
      
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 10)).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 10)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").HorizontalAlignment = xlHAlignLeft
      .Columns("D").HorizontalAlignment = xlHAlignLeft
      .Columns("E").HorizontalAlignment = xlHAlignLeft
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      
      r_int_NroFil = r_int_NroFil + 1
      For r_int_nroaux = 1 To grd_Listad.Rows - 1
         .Cells(r_int_NroFil, 1) = "'" & Format(moddat_g_str_FecSis, "DD/MM/YYYY")
         .Cells(r_int_NroFil, 2) = "EDPYME MICASITA SA"
         .Cells(r_int_NroFil, 3) = grd_Listad.TextMatrix(r_int_nroaux, 16)
         .Cells(r_int_NroFil, 4) = grd_Listad.TextMatrix(r_int_nroaux, 17)
         .Cells(r_int_NroFil, 5) = grd_Listad.TextMatrix(r_int_nroaux, 18)
         .Cells(r_int_NroFil, 6) = grd_Listad.TextMatrix(r_int_nroaux, 0)
         .Cells(r_int_NroFil, 7) = grd_Listad.TextMatrix(r_int_nroaux, 7)
         .Cells(r_int_NroFil, 8) = "'" & grd_Listad.TextMatrix(r_int_nroaux, 8)
         .Cells(r_int_NroFil, 9) = "'" & grd_Listad.TextMatrix(r_int_nroaux, 10)
         .Cells(r_int_NroFil, 10) = grd_Listad.TextMatrix(r_int_nroaux, 1)
         
         r_int_NroFil = r_int_NroFil + 1
      Next
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub grd_Listad_Click()
   Static Modo  As Boolean
   If (grd_Listad.MouseRow = 0) Then
       If grd_Listad.MouseCol = 4 Then
          grd_Listad.Col = 13 '11
       ElseIf grd_Listad.MouseCol = 5 Then
          grd_Listad.Col = 14 '12
       ElseIf grd_Listad.MouseCol = 8 Then
          grd_Listad.Col = 9
       ElseIf grd_Listad.MouseCol = 2 Then
          grd_Listad.Col = 20
       Else
          grd_Listad.Col = grd_Listad.MouseCol
       End If
       If Modo Then
       ' Ordena en forma ascendente
           grd_Listad.Sort = 2
           Modo = False
       ' Ordena en forma descendente
       Else
           grd_Listad.Sort = 1
           Modo = True
       End If
       If grd_Listad.Rows > 1 Then
          Call gs_UbicaGrid(grd_Listad, 1)
       Else
          Call gs_UbicaGrid(grd_Listad, 0)
       End If
   End If
End Sub

Private Sub grd_Listad_DblClick()
   If grd_Listad.MouseRow > 0 Then
      Call cmd_EvaSol_Click
   End If
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub
