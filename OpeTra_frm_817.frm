VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Con_PreSeg_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14865
   Icon            =   "OpeTra_frm_817.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   14865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   5025
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14865
      _Version        =   65536
      _ExtentX        =   26220
      _ExtentY        =   8864
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
      Begin Threed.SSPanel SSPanel5 
         Height          =   645
         Left            =   60
         TabIndex        =   1
         Top             =   4320
         Width           =   14775
         _Version        =   65536
         _ExtentX        =   26061
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
         Begin EditLib.fpDateTime ipp_FecEnv 
            Height          =   315
            Left            =   1710
            TabIndex        =   2
            Top             =   180
            Width           =   1335
            _Version        =   196608
            _ExtentX        =   2355
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   3
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "28/09/2004"
            DateCalcMethod  =   0
            DateTimeFormat  =   0
            UserDefinedFormat=   ""
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "F. Envío a COFIDE:"
            Height          =   195
            Left            =   180
            TabIndex        =   3
            Top             =   240
            Width           =   1425
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   60
         TabIndex        =   4
         Top             =   780
         Width           =   14775
         _Version        =   65536
         _ExtentX        =   26061
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_817.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   14160
            Picture         =   "OpeTra_frm_817.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   7
         Top             =   60
         Width           =   14775
         _Version        =   65536
         _ExtentX        =   26061
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
            Height          =   315
            Left            =   690
            TabIndex        =   8
            Top             =   30
            Width           =   8685
            _Version        =   65536
            _ExtentX        =   15319
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Créditos Hipotecarios"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   315
            Left            =   690
            TabIndex        =   9
            Top             =   330
            Width           =   8685
            _Version        =   65536
            _ExtentX        =   15319
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Seguimiento de Prepagos"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
            Picture         =   "OpeTra_frm_817.frx":0890
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   2805
         Left            =   60
         TabIndex        =   10
         Top             =   1470
         Width           =   14775
         _Version        =   65536
         _ExtentX        =   26061
         _ExtentY        =   4948
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisPrePag 
            Height          =   2295
            Left            =   30
            TabIndex        =   11
            Top             =   480
            Width           =   14775
            _ExtentX        =   26061
            _ExtentY        =   4048
            _Version        =   393216
            Rows            =   26
            Cols            =   7
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_NumOpe 
            Height          =   285
            Left            =   60
            TabIndex        =   12
            Top             =   120
            Width           =   2000
            _Version        =   65536
            _ExtentX        =   3528
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cód. COFIDE"
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
         Begin Threed.SSPanel pnl_Tit_TipPpg 
            Height          =   285
            Left            =   5685
            TabIndex        =   13
            Top             =   120
            Width           =   2610
            _Version        =   65536
            _ExtentX        =   4604
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo de Prepago"
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
         Begin Threed.SSPanel pnl_Tit_FecPro 
            Height          =   285
            Left            =   9675
            TabIndex        =   14
            Top             =   120
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Mto. Aplicar Capital"
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
         Begin Threed.SSPanel pnl_Tit_DoiCli 
            Height          =   285
            Left            =   11310
            TabIndex        =   15
            Top             =   120
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
         Begin Threed.SSPanel pnl_Tit_NomCli 
            Height          =   285
            Left            =   1965
            TabIndex        =   16
            Top             =   120
            Width           =   3840
            _Version        =   65536
            _ExtentX        =   6773
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nombre Cliente"
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
         Begin Threed.SSPanel pnl_Tit_FecPpg 
            Height          =   285
            Left            =   8250
            TabIndex        =   17
            Top             =   120
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Prepago"
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
   End
End
Attribute VB_Name = "frm_Con_PreSeg_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Grabar_Click()
Dim r_int_Contad As Integer
Dim r_int_NroFil As Integer

   If CDate(ipp_FecEnv.Text) < date Then
      MsgBox "La Fecha de Envío no puede ser menor a la Fecha Actual.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecEnv)
      Exit Sub
   End If

   If MsgBox("¿Está seguro de grabar e imprimir los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
     
   'Elimina datos de la tabla temporal
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " DELETE FROM RPT_TABLA_TEMP "
   g_str_Parame = g_str_Parame & "  WHERE RPT_USUCRE = '" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "    AND RPT_TERCRE = '" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "    AND RPT_NOMBRE = 'REPORTE PREPAGOS A COFIDE' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
       Exit Sub
   End If
   
   'Grabacion en Tabla de Temporal
   For r_int_NroFil = 0 To grd_LisPrePag.Rows - 1
   
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "INSERT INTO RPT_TABLA_TEMP("
      g_str_Parame = g_str_Parame & "  RPT_PERMES, "
      g_str_Parame = g_str_Parame & "  RPT_PERANO, "
      g_str_Parame = g_str_Parame & "  RPT_TERCRE, "
      g_str_Parame = g_str_Parame & "  RPT_USUCRE, "
      g_str_Parame = g_str_Parame & "  RPT_NOMBRE, "
      g_str_Parame = g_str_Parame & "  RPT_MONEDA, "
      g_str_Parame = g_str_Parame & "  RPT_FECCRE, "
      g_str_Parame = g_str_Parame & "  RPT_HORCRE, "
      g_str_Parame = g_str_Parame & "  RPT_CODIGO, "
      g_str_Parame = g_str_Parame & "  RPT_DESCRI, "
      g_str_Parame = g_str_Parame & "  RPT_VALCAD01, "
      g_str_Parame = g_str_Parame & "  RPT_VALCAD02, "
      g_str_Parame = g_str_Parame & "  RPT_VALCAD03, "
      g_str_Parame = g_str_Parame & "  RPT_VALCAD04, "
      g_str_Parame = g_str_Parame & "  RPT_VALCAD05, "
      g_str_Parame = g_str_Parame & "  RPT_VALCAD06) "
      
      g_str_Parame = g_str_Parame & "VALUES ("
      g_str_Parame = g_str_Parame & "" & Month(date) & ", "
      g_str_Parame = g_str_Parame & "" & Year(date) & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'REPORTE PREPAGOS A COFIDE', "
      g_str_Parame = g_str_Parame & "'1', "
      g_str_Parame = g_str_Parame & "'" & Format(date, "DDMMYYYY") & "', "
      g_str_Parame = g_str_Parame & "'" & Format(Time, "HHMMSS") & "', "
      g_str_Parame = g_str_Parame & "" & r_int_NroFil & ", "
      g_str_Parame = g_str_Parame & "'" & grd_LisPrePag.TextMatrix(r_int_NroFil, 1) & "', "
      g_str_Parame = g_str_Parame & "'" & grd_LisPrePag.TextMatrix(r_int_NroFil, 0) & "', "
      g_str_Parame = g_str_Parame & "'" & grd_LisPrePag.TextMatrix(r_int_NroFil, 2) & "', "
      g_str_Parame = g_str_Parame & "'" & grd_LisPrePag.TextMatrix(r_int_NroFil, 3) & "', "
      g_str_Parame = g_str_Parame & "'" & grd_LisPrePag.TextMatrix(r_int_NroFil, 4) & "', "
      g_str_Parame = g_str_Parame & "'" & grd_LisPrePag.TextMatrix(r_int_NroFil, 5) & "', "
      g_str_Parame = g_str_Parame & "'" & grd_LisPrePag.TextMatrix(r_int_NroFil, 6) & "')"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         Exit Sub
      End If
      
      'Actualiza el estado del prepago
      g_str_Parame = "USP_ACTUALIZA_CRE_PPGCAB ("
      g_str_Parame = g_str_Parame & "'" & grd_LisPrePag.TextMatrix(r_int_NroFil, 6) & "', "
      g_str_Parame = g_str_Parame & "" & Format(CDate(grd_LisPrePag.TextMatrix(r_int_NroFil, 3)), "yyyymmdd") & " , 2, "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecEnv.Text), "yyyymmdd") & ", 0, 0, 0 ) "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         MsgBox "No se pudo completar la actualización del estado de los datos.", vbInformation, modgen_g_con_PltPar
         Exit Sub
      End If
   Next r_int_NroFil

   Set g_rst_Princi = Nothing
   Set g_rst_Genera = Nothing
   
   moddat_g_int_FlgAct = 2
   Screen.MousePointer = 0
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
  
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Buscar_DatPrepag
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(ipp_FecEnv)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   ipp_FecEnv.Text = Format(date, "dd/mm/yyyy")
   grd_LisPrePag.ColWidth(0) = 2000
   grd_LisPrePag.ColWidth(1) = 3650
   grd_LisPrePag.ColWidth(2) = 2600
   grd_LisPrePag.ColWidth(3) = 1380
   grd_LisPrePag.ColWidth(4) = 1650
   grd_LisPrePag.ColWidth(5) = 3150
   grd_LisPrePag.ColWidth(6) = 0
   grd_LisPrePag.ColAlignment(0) = flexAlignCenterCenter
   grd_LisPrePag.ColAlignment(1) = flexAlignLeftCenter
   grd_LisPrePag.ColAlignment(2) = flexAlignCenterCenter
   grd_LisPrePag.ColAlignment(3) = flexAlignCenterCenter
   grd_LisPrePag.ColAlignment(4) = flexAlignRightCenter
   grd_LisPrePag.ColAlignment(5) = flexAlignLeftCenter
End Sub

Private Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_LisPrePag)
End Sub

Private Sub fs_Buscar_DatPrepag()
Dim r_int_Contad As Integer
  
   For r_int_Contad = 1 To UBound(modatecli_g_arr_TitOpe)
   
      g_str_Parame = " "
      g_str_Parame = g_str_Parame & " SELECT CH.HIPMAE_OPEMVI, TRIM(CL.DATGEN_APEPAT)||' '||TRIM(CL.DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS CLIENTE, "
      g_str_Parame = g_str_Parame & "        PP.PPGCAB_TIPPPG , PP.PPGCAB_FECPPG, (PP.PPGCAB_MTOAPL + PP.PPGCAB_PBPPER) AS MONTO_APLICA, CH.HIPMAE_CODPRD, PP.PPGCAB_TIPPPGPAR, CH.HIPMAE_MONEDA, "
      g_str_Parame = g_str_Parame & "        PP.PPGCAB_REDANO, PP.PPGCAB_NUMOPE"
      g_str_Parame = g_str_Parame & "   FROM CRE_PPGCAB PP"
      g_str_Parame = g_str_Parame & "        INNER JOIN CRE_HIPMAE CH ON CH.HIPMAE_NUMOPE = PP.PPGCAB_NUMOPE "
      g_str_Parame = g_str_Parame & "        INNER JOIN CLI_DATGEN CL ON CL.DATGEN_TIPDOC = CH.HIPMAE_TDOCLI AND CL.DATGEN_NUMDOC = CH.HIPMAE_NDOCLI "
      g_str_Parame = g_str_Parame & "  WHERE CH.HIPMAE_NUMOPE = '" & modatecli_g_arr_TitOpe(r_int_Contad).CreHip_NumOpe & "' "
      g_str_Parame = g_str_Parame & "    AND PP.PPGCAB_FECPPG = '" & modatecli_g_arr_TitOpe(r_int_Contad).CreHip_FecAct & "' "
      g_str_Parame = g_str_Parame & "    AND PP.PPGCAB_FLGEST = 1 "
      g_str_Parame = g_str_Parame & "  ORDER BY PP.PPGCAB_NUMOPE ASC, PP.PPGCAB_FECPPG ASC "
   
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
         grd_LisPrePag.Rows = grd_LisPrePag.Rows + 1
         grd_LisPrePag.Row = grd_LisPrePag.Rows - 1
         
         'CÓDIGO COFIDE
         grd_LisPrePag.Col = 0
         grd_LisPrePag.Text = IIf(IsNull(g_rst_Princi!HIPMAE_OPEMVI), "", Trim(g_rst_Princi!HIPMAE_OPEMVI))
         
         'NOMBRE DEL CLIENTE
         grd_LisPrePag.Col = 1
         grd_LisPrePag.Text = Trim(g_rst_Princi!CLIENTE)
         
         'TIPO PREPAGO
         grd_LisPrePag.Col = 2
         If g_rst_Princi!PPGCAB_TIPPPG = 1 Then
            If g_rst_Princi!PPGCAB_TIPPPGPAR = 1 Then
               grd_LisPrePag.Text = "PARCIAL - RED MONTO"
            Else
               grd_LisPrePag.Text = "PARCIAL - RED PLAZO " & Trim(g_rst_Princi!PPGCAB_REDANO) & " AÑOS"
            End If
         Else
           grd_LisPrePag.Text = "TOTAL"
         End If
         
         'FECHA DE PREPAGO
         grd_LisPrePag.Col = 3
         grd_LisPrePag.Text = gf_FormatoFecha(CStr(g_rst_Princi!PPGCAB_FECPPG))
         
         'MONTO APLICAR CAPITAL
         grd_LisPrePag.Col = 4
         If g_rst_Princi!PPGCAB_TIPPPG = 1 Then
            If g_rst_Princi!HIPMAE_MONEDA = 1 Then
               grd_LisPrePag.Text = "S/.   " & Format(g_rst_Princi!MONTO_APLICA, "###,###,###,##0.00")
            Else
               grd_LisPrePag.Text = "US$   " & Format(g_rst_Princi!MONTO_APLICA, "###,###,###,##0.00")
            End If
         Else
            grd_LisPrePag.Text = "-"
         End If
      
         'PRODUCTO
         grd_LisPrePag.Col = 5
         grd_LisPrePag.Text = moddat_gf_Consulta_Produc(Trim(g_rst_Princi!HIPMAE_CODPRD))
         
         'OPERACIÒN
         grd_LisPrePag.Col = 6
         grd_LisPrePag.Text = Trim(g_rst_Princi!PPGCAB_NUMOPE)
         
         g_rst_Princi.MoveNext
      Loop
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   Next r_int_Contad
   
   grd_LisPrePag.Redraw = True
   Call gs_UbiIniGrid(grd_LisPrePag)
End Sub

Private Sub ipp_FecEnv_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub
