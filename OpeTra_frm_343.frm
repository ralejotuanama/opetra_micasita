VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Rpt_Transf_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12315
   Icon            =   "OpeTra_frm_343.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   12315
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7665
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12315
      _Version        =   65536
      _ExtentX        =   21722
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   675
         Left            =   60
         TabIndex        =   7
         Top             =   30
         Width           =   12195
         _Version        =   65536
         _ExtentX        =   21511
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
            TabIndex        =   8
            Top             =   180
            Width           =   3135
            _Version        =   65536
            _ExtentX        =   5530
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Reporte de Créditos Transferidos"
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
            Picture         =   "OpeTra_frm_343.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   675
         Left            =   60
         TabIndex        =   9
         Top             =   720
         Width           =   12195
         _Version        =   65536
         _ExtentX        =   21511
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   615
            Left            =   11535
            Picture         =   "OpeTra_frm_343.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   15
            Width           =   615
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   615
            Left            =   15
            Picture         =   "OpeTra_frm_343.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Buscar datos"
            Top             =   15
            Width           =   615
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   615
            Left            =   630
            Picture         =   "OpeTra_frm_343.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Exportar a Excel"
            Top             =   15
            Width           =   615
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   825
         Left            =   60
         TabIndex        =   10
         Top             =   1410
         Width           =   12195
         _Version        =   65536
         _ExtentX        =   21511
         _ExtentY        =   1455
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
            TabIndex        =   0
            Top             =   90
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
            TabIndex        =   1
            Top             =   420
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
         Begin VB.Label Label20 
            Caption         =   "Fecha de Fin:"
            Height          =   195
            Left            =   180
            TabIndex        =   12
            Top             =   480
            Width           =   1125
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha de Inicio:"
            Height          =   225
            Left            =   180
            TabIndex        =   11
            Top             =   150
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   5355
         Left            =   60
         TabIndex        =   13
         Top             =   2250
         Width           =   12195
         _Version        =   65536
         _ExtentX        =   21511
         _ExtentY        =   9446
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
         Begin Threed.SSPanel SSPanel6 
            Height          =   285
            Left            =   45
            TabIndex        =   19
            Top             =   60
            Width           =   525
            _Version        =   65536
            _ExtentX        =   926
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Item"
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
         Begin Threed.SSPanel pnl_Tit_DNI 
            Height          =   285
            Left            =   1740
            TabIndex        =   14
            Top             =   60
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Doc. Ident."
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
         Begin Threed.SSPanel pnl_Tit_ApeNom 
            Height          =   285
            Left            =   2910
            TabIndex        =   15
            Top             =   60
            Width           =   4260
            _Version        =   65536
            _ExtentX        =   7514
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
         Begin Threed.SSPanel pnl_Tit_NroOper 
            Height          =   285
            Left            =   555
            TabIndex        =   16
            Top             =   60
            Width           =   1230
            _Version        =   65536
            _ExtentX        =   2170
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   4965
            Left            =   0
            TabIndex        =   5
            Top             =   360
            Width           =   12150
            _ExtentX        =   21431
            _ExtentY        =   8758
            _Version        =   393216
            Rows            =   30
            Cols            =   6
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_Producto 
            Height          =   285
            Left            =   7080
            TabIndex        =   17
            Top             =   60
            Width           =   3585
            _Version        =   65536
            _ExtentX        =   6324
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
         Begin Threed.SSPanel pnl_Tit_FecTra 
            Height          =   285
            Left            =   10650
            TabIndex        =   18
            Top             =   60
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fec.Transfer."
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
Attribute VB_Name = "frm_Rpt_Transf_01"
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
   
   Call gs_SetFocus(ipp_FecIni)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 500
   grd_Listad.ColWidth(1) = 1200
   grd_Listad.ColWidth(2) = 1200
   grd_Listad.ColWidth(3) = 4200
   grd_Listad.ColWidth(4) = 3500
   grd_Listad.ColWidth(5) = 1200
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
End Sub

Private Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_Listad)
   ipp_FecIni.Text = Format(date - CDate(30), "dd/mm/yyyy")
   ipp_FecFin.Text = Format(date, "dd/mm/yyyy")
End Sub

Private Sub fs_BusCli()
   Call fs_Inicia
   
   'Buscando Información del Crédito
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT ROWNUM,"
   g_str_Parame = g_str_Parame & "       HIPMAE_NUMOPE AS OPERACION, HIPMAE_TDOCLI||'-'||TRIM(HIPMAE_NDOCLI) AS DOC_IDENTIDAD,"
   g_str_Parame = g_str_Parame & "       TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS NOMBRE_CLIENTE,"
   g_str_Parame = g_str_Parame & "       TRIM(C.PRODUC_DESCRI) AS PRODUCTO, "
   g_str_Parame = g_str_Parame & "       SUBSTR(HIPMAE_FECCAN,7,2)||'/'||SUBSTR(HIPMAE_FECCAN,5,2)||'/'||SUBSTR(HIPMAE_FECCAN,1,4) AS FECHA_TRANSFERENCIA "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN B ON B.DATGEN_TIPDOC = HIPMAE_TDOCLI AND DATGEN_NUMDOC = HIPMAE_NDOCLI "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_PRODUC C ON C.PRODUC_CODIGO = HIPMAE_CODPRD "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES D ON (D.PARDES_CODGRP = '027' AND D.PARDES_CODITE = HIPMAE_SITUAC) "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_SITUAC = 6 AND HIPMAE_FECCAN >= '" & Format(ipp_FecIni.Text, "YYYYMMDD") & "' "
   g_str_Parame = g_str_Parame & "   AND HIPMAE_FECCAN <= '" & Format(ipp_FecFin.Text, "YYYYMMDD") & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraros operaciones.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecIni)
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = CStr(g_rst_Princi!ROWNUM)
         
         grd_Listad.Col = 1
         grd_Listad.Text = CStr(g_rst_Princi!OPERACION)
           
         grd_Listad.Col = 2
         grd_Listad.Text = CStr(g_rst_Princi!DOC_IDENTIDAD)
         
         grd_Listad.Col = 3
         grd_Listad.Text = CStr(g_rst_Princi!NOMBRE_CLIENTE)
         
         grd_Listad.Col = 4
         grd_Listad.Text = CStr(g_rst_Princi!PRODUCTO)
         
         grd_Listad.Col = 5
         grd_Listad.Text = CStr(g_rst_Princi!FECHA_TRANSFERENCIA)
         
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
      .Cells(r_int_nrofil, 1) = "ITEM":                        .Columns("A").ColumnWidth = 10
      .Cells(r_int_nrofil, 2) = "OPERACION":                   .Columns("B").ColumnWidth = 12
      .Cells(r_int_nrofil, 3) = "DNI":                         .Columns("C").ColumnWidth = 12
      .Cells(r_int_nrofil, 4) = "APELLIDOS Y NOMBRES":         .Columns("D").ColumnWidth = 45
      .Cells(r_int_nrofil, 5) = "PRODUCTO":                    .Columns("E").ColumnWidth = 40
      .Cells(r_int_nrofil, 6) = "FEC.TRANSF.":                 .Columns("F").ColumnWidth = 12
      
      .Cells(r_int_nrofil, 1).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_nrofil, 2).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_nrofil, 3).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_nrofil, 4).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_nrofil, 5).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_nrofil, 6).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      
      .Cells(r_int_nrofil, 1).Font.Bold = True
      .Cells(r_int_nrofil, 2).Font.Bold = True
      .Cells(r_int_nrofil, 3).Font.Bold = True
      .Cells(r_int_nrofil, 4).Font.Bold = True
      .Cells(r_int_nrofil, 5).Font.Bold = True
      .Cells(r_int_nrofil, 6).Font.Bold = True
 
      r_int_nrofil = r_int_nrofil + 1
      
      For r_int_nroaux = 0 To grd_Listad.Rows - 1
         .Cells(r_int_nrofil, 1) = grd_Listad.TextMatrix(r_int_nroaux, 0)
         .Cells(r_int_nrofil, 2) = "'" & (grd_Listad.TextMatrix(r_int_nroaux, 1))
         .Cells(r_int_nrofil, 3) = grd_Listad.TextMatrix(r_int_nroaux, 2)
         .Cells(r_int_nrofil, 4) = grd_Listad.TextMatrix(r_int_nroaux, 3)
         .Cells(r_int_nrofil, 5) = grd_Listad.TextMatrix(r_int_nroaux, 4)
         .Cells(r_int_nrofil, 6) = "'" & grd_Listad.TextMatrix(r_int_nroaux, 5)
         r_int_nrofil = r_int_nrofil + 1
      Next

      .Range(.Cells(1, 8), .Cells(r_int_nrofil, 8)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(r_int_nrofil, 8)).Font.Size = 9
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFin)
   End If
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub pnl_Tit_NroOper_Click()
   If Len(Trim(pnl_Tit_NroOper.Tag)) = 0 Or pnl_Tit_NroOper.Tag = "D" Then
      pnl_Tit_NroOper.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Tit_NroOper.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub

Private Sub pnl_Tit_FecTra_Click()
   If Len(Trim(pnl_Tit_FecTra.Tag)) = 0 Or pnl_Tit_FecTra.Tag = "D" Then
      pnl_Tit_FecTra.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 5, "C")
   Else
      pnl_Tit_FecTra.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 5, "C-")
   End If
End Sub

Private Sub pnl_Tit_ApeNom_Click()
   If Len(Trim(pnl_Tit_ApeNom.Tag)) = 0 Or pnl_Tit_ApeNom.Tag = "D" Then
      pnl_Tit_ApeNom.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 3, "C")
   Else
      pnl_Tit_ApeNom.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 3, "C-")
   End If
End Sub

Private Sub pnl_Tit_DNI_Click()
   If Len(Trim(pnl_Tit_DNI.Tag)) = 0 Or pnl_Tit_DNI.Tag = "D" Then
      pnl_Tit_DNI.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_Tit_DNI.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub

Private Sub pnl_Tit_Producto_Click()
   If Len(Trim(pnl_Tit_Producto.Tag)) = 0 Or pnl_Tit_Producto.Tag = "D" Then
      pnl_Tit_Producto.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 4, "C")
   Else
      pnl_Tit_Producto.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 4, "C-")
   End If
End Sub

