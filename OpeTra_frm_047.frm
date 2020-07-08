VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Caj_GasCie_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form3"
   ClientHeight    =   7860
   ClientLeft      =   5760
   ClientTop       =   2400
   ClientWidth     =   8340
   Icon            =   "OpeTra_frm_047.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel2 
      Height          =   7875
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8355
      _Version        =   65536
      _ExtentX        =   14737
      _ExtentY        =   13891
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
         Height          =   1425
         Left            =   30
         TabIndex        =   7
         Top             =   6390
         Width           =   8265
         _Version        =   65536
         _ExtentX        =   14579
         _ExtentY        =   2514
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
         Begin VB.ComboBox cmb_CodBan 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   6585
         End
         Begin VB.TextBox txt_NumCom 
            Height          =   315
            Left            =   1620
            MaxLength       =   25
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   1050
            Width           =   2775
         End
         Begin VB.ComboBox cmb_CtaBan 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   390
            Width           =   3795
         End
         Begin EditLib.fpDateTime ipp_FecPag 
            Height          =   315
            Left            =   1620
            TabIndex        =   2
            Top             =   720
            Width           =   1305
            _Version        =   196608
            _ExtentX        =   2302
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
         Begin VB.Label Label6 
            Caption         =   "Banco:"
            Height          =   315
            Left            =   60
            TabIndex        =   11
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Nro. Cuenta:"
            Height          =   285
            Left            =   60
            TabIndex        =   10
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Fecha de Pago:"
            Height          =   315
            Left            =   60
            TabIndex        =   9
            Top             =   720
            Width           =   1245
         End
         Begin VB.Label Label11 
            Caption         =   "Nro. Comprobante:"
            Height          =   285
            Left            =   60
            TabIndex        =   8
            Top             =   1050
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   12
         Top             =   30
         Width           =   8265
         _Version        =   65536
         _ExtentX        =   14579
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
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   7650
            Top             =   60
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
            Left            =   6930
            Top             =   60
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
         Begin MSComDlg.CommonDialog dlg_Guarda 
            Left            =   6300
            Top             =   60
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   5880
            Top             =   90
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   255
            Left            =   600
            TabIndex        =   36
            Top             =   60
            Width           =   5000
            _Version        =   65536
            _ExtentX        =   8819
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Operaciones Financieras"
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
         Begin Threed.SSPanel SSPanel4 
            Height          =   255
            Left            =   600
            TabIndex        =   37
            Top             =   330
            Width           =   5000
            _Version        =   65536
            _ExtentX        =   8819
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Créditos Hipotecarios - Cobro de Gastos de Cierre"
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
            Picture         =   "OpeTra_frm_047.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   1425
         Left            =   30
         TabIndex        =   13
         Top             =   1440
         Width           =   8265
         _Version        =   65536
         _ExtentX        =   14579
         _ExtentY        =   2514
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1620
            TabIndex        =   14
            Top             =   60
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3360
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
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   1620
            TabIndex        =   15
            Top             =   720
            Width           =   6585
            _Version        =   65536
            _ExtentX        =   11615
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
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   315
            Left            =   1620
            TabIndex        =   29
            Top             =   1050
            Width           =   2925
            _Version        =   65536
            _ExtentX        =   5159
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
         Begin Threed.SSPanel pnl_DocIde 
            Height          =   315
            Left            =   1620
            TabIndex        =   31
            Top             =   390
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3360
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
         Begin VB.Label Label12 
            Caption         =   "DOI Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   32
            Top             =   390
            Width           =   1395
         End
         Begin VB.Label Label1 
            Caption         =   "Moneda:"
            Height          =   315
            Left            =   60
            TabIndex        =   30
            Top             =   1050
            Width           =   1125
         End
         Begin VB.Label Label2 
            Caption         =   "Nro. de Solicitud:"
            Height          =   315
            Left            =   60
            TabIndex        =   17
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label Label3 
            Caption         =   "Nombre Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   16
            Top             =   720
            Width           =   1125
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   3435
         Left            =   30
         TabIndex        =   18
         Top             =   2910
         Width           =   8265
         _Version        =   65536
         _ExtentX        =   14579
         _ExtentY        =   6059
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
         Begin Threed.SSPanel pnl_Import 
            Height          =   315
            Left            =   6270
            TabIndex        =   19
            Top             =   2370
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
            RoundedCorners  =   0   'False
            Font3D          =   2
            Alignment       =   4
         End
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   2025
            Left            =   30
            TabIndex        =   20
            Top             =   330
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   3572
            _Version        =   393216
            Rows            =   21
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   60
            TabIndex        =   21
            Top             =   60
            Width           =   6225
            _Version        =   65536
            _ExtentX        =   10980
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Concepto"
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
            Left            =   6270
            TabIndex        =   22
            Top             =   60
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Importe"
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
         Begin Threed.SSPanel pnl_ITFImp 
            Height          =   315
            Left            =   6270
            TabIndex        =   23
            Top             =   2700
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
            RoundedCorners  =   0   'False
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_TotImp 
            Height          =   315
            Left            =   6270
            TabIndex        =   24
            Top             =   3030
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   556
            _StockProps     =   15
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
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Alignment       =   4
         End
         Begin VB.Label lbl_Moneda 
            Alignment       =   1  'Right Justify
            Caption         =   "US$"
            Height          =   315
            Index           =   2
            Left            =   5730
            TabIndex        =   35
            Top             =   3060
            Width           =   495
         End
         Begin VB.Label lbl_Moneda 
            Alignment       =   1  'Right Justify
            Caption         =   "US$"
            Height          =   315
            Index           =   1
            Left            =   5730
            TabIndex        =   34
            Top             =   2730
            Width           =   495
         End
         Begin VB.Label lbl_Moneda 
            Alignment       =   1  'Right Justify
            Caption         =   "US$"
            Height          =   315
            Index           =   0
            Left            =   5730
            TabIndex        =   33
            Top             =   2370
            Width           =   495
         End
         Begin VB.Label Label9 
            Caption         =   "Total:"
            Height          =   285
            Left            =   4860
            TabIndex        =   27
            Top             =   3030
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "ITF:"
            Height          =   285
            Left            =   4860
            TabIndex        =   26
            Top             =   2700
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Sub-Total:"
            Height          =   285
            Left            =   4860
            TabIndex        =   25
            Top             =   2370
            Width           =   855
         End
      End
      Begin Threed.SSPanel SSPanel12 
         Height          =   645
         Left            =   30
         TabIndex        =   28
         Top             =   750
         Width           =   8265
         _Version        =   65536
         _ExtentX        =   14579
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
            Left            =   7650
            Picture         =   "OpeTra_frm_047.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_047.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_Caj_GasCie_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_dbl_PorITF     As Double
Dim l_arr_CodBan()   As moddat_tpo_Genera
Dim l_arr_CtaBan()   As moddat_tpo_Genera

Private Sub cmb_CodBan_Click()
   If cmb_CodBan.ListIndex > -1 Then
      Screen.MousePointer = 11
      Call moddat_gs_Carga_CtaBan(l_arr_CodBan(cmb_CodBan.ListIndex + 1).Genera_Codigo, cmb_CtaBan, l_arr_CtaBan)
      Screen.MousePointer = 0
         
      Call gs_SetFocus(cmb_CtaBan)
   Else
      cmb_CtaBan.Clear
   End If
End Sub

Private Sub cmb_CodBan_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CodBan_Click
   End If
End Sub

Private Sub cmb_CtaBan_Click()
   Call gs_SetFocus(ipp_FecPag)
End Sub

Private Sub cmb_CtaBan_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CtaBan_Click
   End If
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_str_Operac     As String
   Dim r_lng_NumMov     As Long
   Dim r_int_Contad     As Integer
   Dim r_int_CodGas     As Integer
   Dim r_dbl_Import     As Double
   
   If cmb_CodBan.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Banco.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodBan)
      Exit Sub
   End If
   
   If cmb_CtaBan.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Número de Cuenta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CtaBan)
      Exit Sub
   End If
   
   If CDate(ipp_FecPag.Text) > date Then
      MsgBox "Debe ingresar una fecha correcta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecPag)
      Exit Sub
   End If
   
   If Len(Trim(txt_NumCom.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Comprobante.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumCom)
      Exit Sub
   End If
   
   If l_arr_CtaBan(cmb_CtaBan.ListIndex + 1).Genera_TipPar <> moddat_g_int_TipMon Then
      MsgBox "La Moneda de la Cuenta no coincide con la Moneda de Asignación de los Gastos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CtaBan)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de registrar el pago?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   r_str_Operac = moddat_gf_Consulta_Operac(moddat_g_str_CodPrd, "210")
   r_str_Operac = CStr(moddat_g_int_TipMon) & Right(r_str_Operac, 5)
   
   'Obteniendo Número de Movimiento
   r_lng_NumMov = opecaj_gf_Genera_NumMov()
   
   'Registrando Movimiento
   If Not opecaj_gf_Inserta_CajMov(modgen_g_str_CodUsu, "1101", moddat_g_str_NumSol, "", moddat_g_int_TipDoc, moddat_g_str_NumDoc, l_arr_CodBan(cmb_CodBan.ListIndex + 1).Genera_Codigo, Format(CDate(ipp_FecPag.Text), "yyyymmdd"), l_arr_CtaBan(cmb_CtaBan.ListIndex + 1).Genera_Codigo, txt_NumCom.Text, moddat_g_int_TipMon, CDbl(pnl_Import.Caption), 0, modgen_g_str_CodSuc, 0, 0, 0, l_dbl_PorITF, CDbl(pnl_ITFImp.Caption), CDbl(pnl_TotImp.Caption), 0, "0", r_str_Operac, r_lng_NumMov, 1, "0", "", "", "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0) Then
      Exit Sub
   End If
   
   'Actualizando Saldo de Caja
   'If Not opecaj_gf_ActualizaSaldo(l_arr_CodBan(cmb_CodBan.ListIndex + 1).Genera_Codigo, moddat_g_int_TipMon, CDbl(pnl_TotImp.Caption)) Then
   '   Exit Sub
   'End If
   
   'Actualizando Pago en Tabla de Gasto Administrativo
   grd_Listad.Redraw = False
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      grd_Listad.Row = r_int_Contad

      grd_Listad.Col = 2
      r_int_CodGas = CInt(grd_Listad.Text)
      
      grd_Listad.Col = 1
      r_dbl_Import = CDbl(grd_Listad.Text)
      
      If Not opecaj_gf_Pago_GasAdm(moddat_g_str_NumSol, r_int_CodGas, moddat_g_int_TipMon, r_dbl_Import, l_dbl_PorITF, Format(CDate(ipp_FecPag.Text), "yyyymmdd"), r_str_Operac) Then
         Exit Sub
      End If
   Next r_int_Contad
   grd_Listad.Redraw = True
   
   'Actualizando en Seguimiento de Tasacion Pago de Gastos Administrativos
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, moddat_g_int_CodIns, 25, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   'Enviando Correo Electrónico
   modgen_g_str_Mail_Asunto = "PAGO DE GASTOS DE CIERRE (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
   
   modgen_g_str_Mail_Mensaj = ""
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
   
   Call fs_Envia_CorreoEle(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, "", 0, False, False, False)
   
   moddat_g_int_FlgAct = 2
   
   'Borrar Spool de PC (Cabecera)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_COMPGC WHERE "
   g_str_Parame = g_str_Parame & "COMPGC_CODTER = '" & modgen_g_str_NombPC & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
       Exit Sub
   End If
   
   'Borrar Spool de PC (Detalle)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_COMPGD WHERE "
   g_str_Parame = g_str_Parame & "COMPGD_CODTER = '" & modgen_g_str_NombPC & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
       Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   Call opecaj_gs_ComPago(modgen_g_str_CodSuc, CStr(r_lng_NumMov), Format(CDate(moddat_g_str_FecSis), "yyyymmdd"), 1, 1)
   
   Screen.MousePointer = 0
   
   'Se conecta al crystal report
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat

   'Se envia las tablas correspondientes en el orden que fueron utilizadas
   crp_Imprim.DataFiles(0) = "RPT_COMPGC"
   crp_Imprim.DataFiles(1) = "RPT_COMPGD"
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_COMPAG_01.RPT"
   crp_Imprim.SelectionFormula = "{RPT_COMPGC.COMPGC_CODTER} = '" & modgen_g_str_NombPC & "'"
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
   
   cmd_Grabar.Enabled = False
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
   Call moddat_gs_FecSis
   
   grd_Listad.ColWidth(0) = 6210
   grd_Listad.ColWidth(1) = 1630
   grd_Listad.ColWidth(2) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignRightCenter

   Call moddat_gs_Carga_LisIte(cmb_CodBan, l_arr_CodBan, 1, "516")
   
   'Obteniendo ITF
   l_dbl_PorITF = opecaj_gf_Consulta_ITF(Format(CDate(moddat_g_str_FecSis), "yyyymmdd"), 1)
End Sub

Private Sub fs_Buscar()
   Dim r_dbl_Import     As Double

   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   pnl_DocIde.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc
   pnl_NomCli.Caption = moddat_g_str_NomCli
   pnl_Moneda.Caption = moddat_gf_Consulta_ParDes("204", CStr(moddat_g_int_TipMon))
   
   lbl_Moneda(0).Caption = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon))
   lbl_Moneda(1).Caption = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon))
   lbl_Moneda(2).Caption = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon))
   
   pnl_Import.Caption = "0.00 "
   pnl_ITFImp.Caption = "0.00 "
   pnl_TotImp.Caption = "0.00 "
   
   cmb_CodBan.ListIndex = -1
   cmb_CtaBan.Clear
   ipp_FecPag.Text = Format(date, "dd/mm/yyyy")
   txt_NumCom.Text = ""
   
   r_dbl_Import = 0
   
   Call gs_LimpiaGrid(grd_Listad)
   
   'Lista de Gastos Administrativos
   g_str_Parame = "SELECT * FROM TRA_GASADM WHERE "
   g_str_Parame = g_str_Parame & "GASADM_NUMSOL = '" & moddat_g_str_NumSol & "' AND GASADM_SITUAC = 2 "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad.Redraw = False
      
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         
         grd_Listad.Row = grd_Listad.Rows - 1
         
         'Buscando Descripción de Gastos Administrativos
         grd_Listad.Col = 0
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "007", Format(g_rst_Princi!GASADM_CODGAS, "00") & Format(g_rst_Princi!GASADM_TIPMON, "0")) Then
            grd_Listad.Text = Trim(moddat_g_arr_Genera(1).Genera_Nombre)
         End If
         
         grd_Listad.Col = 2
         grd_Listad.Text = g_rst_Princi!GASADM_CODGAS
         
         'Importe
         grd_Listad.Col = 1
         grd_Listad.Text = Format(g_rst_Princi!GASADM_IMPORT, "###,###,##0.00")
         
         r_dbl_Import = r_dbl_Import + g_rst_Princi!GASADM_IMPORT
         
         g_rst_Princi.MoveNext
      Loop
         
      grd_Listad.Redraw = True
      
      Call gs_UbiIniGrid(grd_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   pnl_Import.Caption = Format(r_dbl_Import, "###,###,##0.00") & " "
   pnl_ITFImp.Caption = gf_NueImp_Numero(gf_Truncar_Numero(CDbl(pnl_Import.Caption) * (l_dbl_PorITF / 100), 2)) & " "
   pnl_TotImp.Caption = Format(CDbl(pnl_Import.Caption) + CDbl(Trim(pnl_ITFImp.Caption)), "###,###,##0.00") & " "
End Sub

Private Sub ipp_FecPag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumCom)
   End If
End Sub

Private Sub txt_NumCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
   End If
End Sub

Private Sub txt_NumCom_GotFocus()
   Call gs_SelecTodo(txt_NumCom)
End Sub



