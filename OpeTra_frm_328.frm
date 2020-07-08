VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Rpt_CreRef_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   8550
   ClientLeft      =   3345
   ClientTop       =   1950
   ClientWidth     =   12390
   Icon            =   "OpeTra_frm_328.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   12390
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8565
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12465
      _Version        =   65536
      _ExtentX        =   21987
      _ExtentY        =   15108
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
         Height          =   675
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   12315
         _Version        =   65536
         _ExtentX        =   21722
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
            Height          =   285
            Left            =   660
            TabIndex        =   2
            Top             =   180
            Width           =   6855
            _Version        =   65536
            _ExtentX        =   12091
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Reporte de Crédito Refinanciado"
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
            Left            =   90
            Picture         =   "OpeTra_frm_328.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   3
         Top             =   750
         Width           =   12315
         _Version        =   65536
         _ExtentX        =   21722
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
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_328.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Buscar Operaciones"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   11700
            Picture         =   "OpeTra_frm_328.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_328.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   6
         Top             =   1440
         Width           =   12315
         _Version        =   65536
         _ExtentX        =   21722
         _ExtentY        =   1349
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
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1440
            TabIndex        =   7
            Top             =   60
            Width           =   1545
            _Version        =   196608
            _ExtentX        =   2725
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
            Text            =   "01/01/2008"
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
         Begin EditLib.fpDateTime ipp_FecFin 
            Height          =   315
            Left            =   1440
            TabIndex        =   8
            Top             =   390
            Width           =   1545
            _Version        =   196608
            _ExtentX        =   2725
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
            Text            =   "01/01/2008"
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
         Begin VB.Label Label2 
            Caption         =   "Fecha de Inicio:"
            Height          =   225
            Left            =   60
            TabIndex        =   10
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label20 
            Caption         =   "Fecha de Fin:"
            Height          =   195
            Left            =   60
            TabIndex        =   9
            Top             =   450
            Width           =   1125
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   6195
         Left            =   60
         TabIndex        =   11
         Top             =   2280
         Width           =   12315
         _Version        =   65536
         _ExtentX        =   21722
         _ExtentY        =   10927
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
            TabIndex        =   12
            Top             =   60
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "DNI"
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
            Left            =   1200
            TabIndex        =   13
            Top             =   60
            Width           =   3810
            _Version        =   65536
            _ExtentX        =   6720
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
         Begin Threed.SSPanel pnl_Tit_DocIde 
            Height          =   285
            Left            =   5010
            TabIndex        =   14
            Top             =   60
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro Operación"
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
            Left            =   6330
            TabIndex        =   15
            Top             =   60
            Width           =   1230
            _Version        =   65536
            _ExtentX        =   2170
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Desemb."
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
            Height          =   5805
            Left            =   30
            TabIndex        =   16
            Top             =   360
            Width           =   12275
            _ExtentX        =   21643
            _ExtentY        =   10239
            _Version        =   393216
            Rows            =   30
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   285
            Left            =   7560
            TabIndex        =   17
            Top             =   60
            Width           =   2025
            _Version        =   65536
            _ExtentX        =   3572
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Moneda Desembolso"
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   9570
            TabIndex        =   18
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Monto Desemb."
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   285
            Left            =   10770
            TabIndex        =   20
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
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
      End
   End
End
Attribute VB_Name = "frm_Rpt_CreRef_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Buscar_Click()
   If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
      MsgBox "La fecha de fin no puede ser menor a la fecha de inicio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecIni)
      Exit Sub
   End If
   
   Call fs_BusCli
End Sub

Private Sub cmd_ExpExc_Click()
   If grd_Listad.Rows = 0 Then
      MsgBox "No existe información a exportar.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Exportar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
  
   Call fs_Inicia
   Call fs_Limpia
   Call gs_CentraForm(Me)

   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 0
   grd_Listad.ColWidth(1) = 1140
   grd_Listad.ColWidth(2) = 3810
   grd_Listad.ColWidth(3) = 1320
   grd_Listad.ColWidth(4) = 1220
   grd_Listad.ColWidth(5) = 2020
   grd_Listad.ColWidth(6) = 1200
   grd_Listad.ColWidth(7) = 1200
    
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_Listad.ColAlignment(7) = flexAlignCenterCenter
End Sub

Private Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_Listad)
   
   ipp_FecIni.Text = Format(date - CDate(30), "dd/mm/yyyy")
   ipp_FecFin.Text = Format(date, "dd/mm/yyyy")
End Sub

Private Sub fs_BusCli()
Dim r_dbl_TotPag  As Double

   Call fs_Inicia
   
   'Buscando Información del Crédito
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT "
   g_str_Parame = g_str_Parame & "A.HIPMAE_TDOCLI ||' - ' || TRIM(A.HIPMAE_NDOCLI) AS DNI, "
   g_str_Parame = g_str_Parame & "TRIM(B.DATGEN_APEPAT) ||' ' || TRIM(B.DATGEN_APEMAT) ||' '|| TRIM(B.DATGEN_NOMBRE) AS NOMBRE, "
   g_str_Parame = g_str_Parame & "SUBSTR(A.HIPMAE_NUMOPE,1,3) ||'-' || SUBSTR(A.HIPMAE_NUMOPE,4,2) ||'-' ||SUBSTR(A.HIPMAE_NUMOPE,6,5) AS NUMOPE, "
   g_str_Parame = g_str_Parame & "TO_DATE(A.HIPMAE_FECDES,'YYYY/MM/DD') AS FECHA, "
   g_str_Parame = g_str_Parame & "TRIM(C.PARDES_DESCRI) AS MONEDA, "
   g_str_Parame = g_str_Parame & "ROUND(A.HIPMAE_IMPDES, 2) AS MONTO, "
   g_str_Parame = g_str_Parame & "TRIM(D.PARDES_DESCRI) AS ESTADO "
   g_str_Parame = g_str_Parame & "FROM CRE_HIPMAE A "
   g_str_Parame = g_str_Parame & "INNER JOIN CLI_DATGEN B ON (B.DATGEN_NUMDOC = A.HIPMAE_NDOCLI AND B.DATGEN_TIPDOC = A.HIPMAE_TDOCLI) "
   g_str_Parame = g_str_Parame & "INNER JOIN MNT_PARDES C ON (C.PARDES_CODGRP = '204' AND C.PARDES_CODITE = A.HIPMAE_MONEDA) "
   g_str_Parame = g_str_Parame & "INNER JOIN MNT_PARDES D ON (D.PARDES_CODGRP = '027' AND D.PARDES_CODITE = A.HIPMAE_SITUAC) "
   g_str_Parame = g_str_Parame & "WHERE HIPMAE_REFINA = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY NUMOPE "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "El cliente no cuenta con una clasificación.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_dbl_TotPag = 0
      
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
                         
         grd_Listad.Col = 1
         grd_Listad.Text = CStr(g_rst_Princi!DNI)
           
         grd_Listad.Col = 2
         grd_Listad.Text = CStr(g_rst_Princi!NOMBRE)
         
         grd_Listad.Col = 3
         grd_Listad.Text = CStr(g_rst_Princi!NUMOPE)
         
         grd_Listad.Col = 4
         grd_Listad.Text = CStr(g_rst_Princi!Fecha)
         
         grd_Listad.Col = 5
         grd_Listad.Text = CStr(g_rst_Princi!MONEDA)
         
         grd_Listad.Col = 6
         grd_Listad.Text = Format(g_rst_Princi!MONTO, "###,###,##0.00")
         
         grd_Listad.Col = 7
         grd_Listad.Text = CStr(g_rst_Princi!ESTADO)
         g_rst_Princi.MoveNext
      Loop
      
      Call gs_UbiIniGrid(grd_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Exportar()
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_nrofil     As Integer
   Dim r_int_nroaux     As Integer
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   r_int_nrofil = 1
   
   With r_obj_Excel.ActiveSheet
      .Cells(r_int_nrofil, 1) = "DNI":                   .Columns("A").ColumnWidth = 12
      .Cells(r_int_nrofil, 2) = "APELLIDOS Y NOMBRES":   .Columns("B").ColumnWidth = 45
      .Cells(r_int_nrofil, 3) = "NRO DE OPERACIÓN":      .Columns("C").ColumnWidth = 17
      .Cells(r_int_nrofil, 4) = "FECHA DESEMBOLSO":      .Columns("D").ColumnWidth = 17
      .Cells(r_int_nrofil, 5) = "MONEDA DESEMBOLSO":     .Columns("E").ColumnWidth = 20
      .Cells(r_int_nrofil, 6) = "MONTO DESEMBOLSO":      .Columns("F").ColumnWidth = 18
      .Cells(r_int_nrofil, 7) = "SITUACION":             .Columns("G").ColumnWidth = 14
      
      .Cells(r_int_nrofil, 1).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_nrofil, 2).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_nrofil, 3).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_nrofil, 4).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_nrofil, 5).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_nrofil, 6).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_nrofil, 7).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      
      .Cells(r_int_nrofil, 1).Font.Bold = True
      .Cells(r_int_nrofil, 2).Font.Bold = True
      .Cells(r_int_nrofil, 3).Font.Bold = True
      .Cells(r_int_nrofil, 4).Font.Bold = True
      .Cells(r_int_nrofil, 5).Font.Bold = True
      .Cells(r_int_nrofil, 6).Font.Bold = True
      .Cells(r_int_nrofil, 7).Font.Bold = True
 
      r_int_nrofil = r_int_nrofil + 1
      
      For r_int_nroaux = 0 To grd_Listad.Rows - 1
         .Cells(r_int_nrofil, 1) = grd_Listad.TextMatrix(r_int_nroaux, 1)
         .Cells(r_int_nrofil, 2) = grd_Listad.TextMatrix(r_int_nroaux, 2)
         .Cells(r_int_nrofil, 3) = grd_Listad.TextMatrix(r_int_nroaux, 3)
         .Cells(r_int_nrofil, 4) = grd_Listad.TextMatrix(r_int_nroaux, 4)
         .Cells(r_int_nrofil, 5) = grd_Listad.TextMatrix(r_int_nroaux, 5)
         .Cells(r_int_nrofil, 6) = grd_Listad.TextMatrix(r_int_nroaux, 6)
         .Cells(r_int_nrofil, 7) = grd_Listad.TextMatrix(r_int_nroaux, 7)
         r_int_nrofil = r_int_nrofil + 1
      Next
      
      .Columns("F").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      r_int_nrofil = r_int_nrofil + 2
       
      .Range(.Cells(1, 8), .Cells(r_int_nrofil, 8)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(r_int_nrofil, 8)).Font.Size = 9
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

