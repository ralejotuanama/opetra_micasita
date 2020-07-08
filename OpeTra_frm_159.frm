VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_SolCre_54 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   9555
   ClientLeft      =   3480
   ClientTop       =   495
   ClientWidth     =   11685
   Icon            =   "OpeTra_frm_159.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9555
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9555
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   11685
      _Version        =   65536
      _ExtentX        =   20611
      _ExtentY        =   16854
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
      Begin Threed.SSPanel SSPanel9 
         Height          =   495
         Left            =   30
         TabIndex        =   33
         Top             =   6420
         Width           =   11625
         _Version        =   65536
         _ExtentX        =   20505
         _ExtentY        =   873
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
         Begin VB.ComboBox cmb_InsFin 
            Height          =   315
            Left            =   8940
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   90
            Width           =   2550
         End
         Begin VB.ComboBox cmb_MonAho 
            Height          =   315
            Left            =   3240
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   90
            Width           =   2010
         End
         Begin EditLib.fpDoubleSingle ipp_MtoAho 
            Height          =   315
            Left            =   1290
            TabIndex        =   21
            Top             =   90
            Width           =   1020
            _Version        =   196608
            _ExtentX        =   1799
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
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
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
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "50000"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpLongInteger ipp_MesAho 
            Height          =   315
            Left            =   6540
            TabIndex        =   23
            Top             =   90
            Width           =   690
            _Version        =   196608
            _ExtentX        =   1217
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
            ButtonStyle     =   1
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
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
            Text            =   "0"
            MaxValue        =   "12"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label19 
            Caption         =   "Institución Financ.:"
            Height          =   315
            Left            =   7380
            TabIndex        =   37
            Top             =   120
            Width           =   1605
         End
         Begin VB.Label Label1 
            Caption         =   "Moneda:"
            Height          =   315
            Left            =   2490
            TabIndex        =   36
            Top             =   120
            Width           =   645
         End
         Begin VB.Label Label18 
            Caption         =   "Monto Ahorro:"
            Height          =   285
            Left            =   120
            TabIndex        =   35
            Top             =   120
            Width           =   1035
         End
         Begin VB.Label Label22 
            Caption         =   "Meses Ahorro:"
            Height          =   285
            Left            =   5400
            TabIndex        =   34
            Top             =   120
            Width           =   1155
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1695
         Left            =   30
         TabIndex        =   38
         Top             =   7830
         Width           =   11625
         _Version        =   65536
         _ExtentX        =   20505
         _ExtentY        =   2990
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
            Height          =   1305
            Left            =   60
            TabIndex        =   28
            Top             =   330
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   2302
            _Version        =   393216
            Rows            =   12
            Cols            =   7
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   9780
            TabIndex        =   39
            Top             =   60
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Recepcionado"
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
         Begin Threed.SSPanel SSPanel6 
            Height          =   285
            Left            =   90
            TabIndex        =   40
            Top             =   60
            Width           =   9705
            _Version        =   65536
            _ExtentX        =   17119
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Documento"
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
      Begin Threed.SSPanel SSPanel10 
         Height          =   885
         Left            =   30
         TabIndex        =   41
         Top             =   6930
         Width           =   11625
         _Version        =   65536
         _ExtentX        =   20505
         _ExtentY        =   1561
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
         Begin VB.ComboBox cmb_EjeSeg 
            Height          =   315
            Left            =   8010
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   450
            Width           =   3540
         End
         Begin VB.ComboBox cmb_ConHip 
            Height          =   315
            Left            =   8010
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   90
            Width           =   3540
         End
         Begin VB.TextBox txt_Observ 
            Height          =   675
            Left            =   1290
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   25
            Top             =   90
            Width           =   4545
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Ejecutivo de Seguimiento:"
            Height          =   165
            Left            =   6060
            TabIndex        =   44
            Top             =   510
            Width           =   1845
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Consejero Hipotecario:"
            Height          =   195
            Left            =   6060
            TabIndex        =   43
            Top             =   150
            Width           =   1845
         End
         Begin VB.Label Label5 
            Caption         =   "Observaciones:"
            Height          =   315
            Left            =   90
            TabIndex        =   42
            Top             =   120
            Width           =   1905
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   45
         Top             =   660
         Width           =   11625
         _Version        =   65536
         _ExtentX        =   20505
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
            Left            =   10980
            Picture         =   "OpeTra_frm_159.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_159.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Aceptar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_SimCre 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_159.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Simulación de Créditos Hipotecarios"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   4275
         Left            =   30
         TabIndex        =   46
         Top             =   2130
         Width           =   11625
         _Version        =   65536
         _ExtentX        =   20505
         _ExtentY        =   7541
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
         Begin VB.ComboBox cmb_TasEsp 
            Height          =   315
            Left            =   8940
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   3120
            Width           =   2550
         End
         Begin VB.ComboBox cmb_EmpSeg 
            Height          =   315
            Left            =   8940
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   2460
            Width           =   2550
         End
         Begin VB.ComboBox cmb_SegDes 
            Height          =   315
            Left            =   8940
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   2790
            Width           =   2550
         End
         Begin VB.ComboBox cmb_PriViv 
            Height          =   315
            Left            =   8940
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   450
            Width           =   2595
         End
         Begin VB.ComboBox cmb_TipEva 
            Height          =   315
            Left            =   8940
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   120
            Width           =   2595
         End
         Begin VB.ComboBox cmb_DiaPag 
            Height          =   315
            Left            =   8940
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   2130
            Width           =   2550
         End
         Begin VB.ComboBox cmb_CuoDbl 
            Height          =   315
            Left            =   8940
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1800
            Width           =   2550
         End
         Begin EditLib.fpDoubleSingle ipp_ComVta 
            Height          =   315
            Left            =   2010
            TabIndex        =   0
            Top             =   450
            Width           =   1215
            _Version        =   196608
            _ExtentX        =   2143
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
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
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
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDoubleSingle ipp_ApoPro 
            Height          =   315
            Left            =   2010
            TabIndex        =   3
            Top             =   1470
            Width           =   1215
            _Version        =   196608
            _ExtentX        =   2143
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
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
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
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDoubleSingle ipp_MtoAFP 
            Height          =   315
            Left            =   2010
            TabIndex        =   6
            Top             =   2460
            Width           =   1215
            _Version        =   196608
            _ExtentX        =   2143
            _ExtentY        =   556
            Enabled         =   0   'False
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
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
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
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin Threed.SSPanel pnl_PorIni 
            Height          =   225
            Left            =   3210
            TabIndex        =   60
            Top             =   1170
            Width           =   885
            _Version        =   65536
            _ExtentX        =   1561
            _ExtentY        =   397
            _StockProps     =   15
            Caption         =   "(24.0000%) "
            ForeColor       =   32768
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_FmvBbp 
            Height          =   315
            Left            =   2010
            TabIndex        =   4
            Top             =   1800
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_MtoBMS 
            Height          =   315
            Left            =   2010
            TabIndex        =   7
            Top             =   2790
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_CuoIni_Sol 
            Height          =   315
            Left            =   5550
            TabIndex        =   66
            Top             =   1140
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_MtoAFP_Sol 
            Height          =   315
            Left            =   5550
            TabIndex        =   67
            Top             =   2460
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_ApoPro_Sol 
            Height          =   315
            Left            =   5550
            TabIndex        =   68
            Top             =   1470
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_FmvBbp_Sol 
            Height          =   315
            Left            =   5550
            TabIndex        =   69
            Top             =   1800
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_MtoPre_Sol 
            Height          =   315
            Left            =   5550
            TabIndex        =   70
            Top             =   3180
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_MtoPre 
            Height          =   315
            Left            =   1710
            TabIndex        =   8
            Top             =   3180
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_ValGas_Sol 
            Height          =   315
            Left            =   5550
            TabIndex        =   71
            Top             =   3510
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_TotPre 
            Height          =   315
            Left            =   1710
            TabIndex        =   11
            Top             =   3840
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_MtoBMS_Sol 
            Height          =   315
            Left            =   5550
            TabIndex        =   73
            Top             =   2790
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_MefPbp 
            Height          =   315
            Left            =   2010
            TabIndex        =   5
            Top             =   2130
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_MefPbp_Sol 
            Height          =   315
            Left            =   5550
            TabIndex        =   76
            Top             =   2130
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_ValTot 
            Height          =   315
            Left            =   1710
            TabIndex        =   79
            Top             =   120
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin EditLib.fpDoubleSingle ipp_ValEst 
            Height          =   315
            Left            =   2010
            TabIndex        =   1
            Top             =   780
            Width           =   1215
            _Version        =   196608
            _ExtentX        =   2143
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
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
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
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin Threed.SSPanel pnl_CuoIni 
            Height          =   315
            Left            =   1710
            TabIndex        =   2
            Top             =   1140
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_ValGas 
            Height          =   315
            Left            =   1710
            TabIndex        =   10
            Top             =   3510
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin EditLib.fpLongInteger ipp_PlaAno 
            Height          =   315
            Left            =   8940
            TabIndex        =   14
            Top             =   1140
            Width           =   2550
            _Version        =   196608
            _ExtentX        =   4498
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
            ButtonStyle     =   1
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
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
            Text            =   "0"
            MaxValue        =   "30"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpLongInteger ipp_PerGra 
            Height          =   315
            Left            =   8940
            TabIndex        =   15
            Top             =   1470
            Width           =   2550
            _Version        =   196608
            _ExtentX        =   4498
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
            ButtonStyle     =   1
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
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
            Text            =   "0"
            MaxValue        =   "6"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin Threed.SSPanel pnl_ValTot_Sol 
            Height          =   315
            Left            =   5550
            TabIndex        =   91
            Top             =   120
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_ComVta_Sol 
            Height          =   315
            Left            =   5550
            TabIndex        =   93
            Top             =   450
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_ValEst_Sol 
            Height          =   315
            Left            =   5550
            TabIndex        =   95
            Top             =   780
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_TotPre_Sol 
            Height          =   315
            Left            =   5550
            TabIndex        =   97
            Top             =   3840
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSCheck chk_Gastos 
            Height          =   225
            Left            =   120
            TabIndex        =   9
            Top             =   3570
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Incluye Gastos"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Tasa Especial:"
            Height          =   195
            Left            =   7380
            TabIndex        =   100
            Top             =   3180
            Width           =   1050
         End
         Begin VB.Label Label33 
            Caption         =   "Valores en S/.:"
            Height          =   315
            Left            =   4350
            TabIndex        =   99
            Top             =   3570
            Width           =   1185
         End
         Begin VB.Label Label28 
            Caption         =   "Valores en S/.:"
            Height          =   315
            Left            =   4350
            TabIndex        =   98
            Top             =   3900
            Width           =   1185
         End
         Begin VB.Label Label16 
            Caption         =   "Valores en S/.:"
            Height          =   315
            Left            =   4350
            TabIndex        =   96
            Top             =   840
            Width           =   1185
         End
         Begin VB.Label Label12 
            Caption         =   "Valores en S/.:"
            Height          =   315
            Left            =   4350
            TabIndex        =   94
            Top             =   510
            Width           =   1185
         End
         Begin VB.Label Label10 
            Caption         =   "Valores en S/.:"
            Height          =   315
            Left            =   4350
            TabIndex        =   92
            Top             =   180
            Width           =   1185
         End
         Begin VB.Label Label14 
            Caption         =   "Compañía Seguros:"
            Height          =   315
            Left            =   7380
            TabIndex        =   90
            Top             =   2520
            Width           =   1605
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo Seg. Desgrav.:"
            Height          =   315
            Left            =   7380
            TabIndex        =   89
            Top             =   2850
            Width           =   1605
         End
         Begin VB.Label Label23 
            Caption         =   "Primera Vivienda:"
            Height          =   315
            Left            =   7380
            TabIndex        =   88
            Top             =   510
            Width           =   1605
         End
         Begin VB.Label Label9 
            Caption         =   "Tipo de Evaluación:"
            Height          =   315
            Left            =   7380
            TabIndex        =   87
            Top             =   180
            Width           =   1605
         End
         Begin VB.Label Label6 
            Caption         =   "Día de Pago:"
            Height          =   315
            Left            =   7380
            TabIndex        =   86
            Top             =   2190
            Width           =   1605
         End
         Begin VB.Label Label32 
            Caption         =   "Cuotas Dobles:"
            Height          =   315
            Left            =   7380
            TabIndex        =   85
            Top             =   1860
            Width           =   1605
         End
         Begin VB.Label Label4 
            Caption         =   "Período Gracia:"
            Height          =   285
            Left            =   7380
            TabIndex        =   84
            Top             =   1530
            Width           =   1605
         End
         Begin VB.Label Label29 
            Caption         =   "Plazo en Años:"
            Height          =   285
            Left            =   7380
            TabIndex        =   83
            Top             =   1200
            Width           =   1605
         End
         Begin VB.Label Label44 
            Caption         =   "Cuota Inicial"
            Height          =   285
            Left            =   120
            TabIndex        =   82
            Top             =   1200
            Width           =   1605
         End
         Begin VB.Label Label43 
            Caption         =   "Valor Estacio.:"
            Height          =   285
            Left            =   510
            TabIndex        =   81
            Top             =   840
            Width           =   1155
         End
         Begin VB.Label Label42 
            Caption         =   "Valor Total Vivienda"
            Height          =   285
            Left            =   120
            TabIndex        =   80
            Top             =   180
            Width           =   1605
         End
         Begin VB.Label Label41 
            Caption         =   "Bono PBP:"
            Height          =   285
            Left            =   510
            TabIndex        =   78
            Top             =   2190
            Width           =   1245
         End
         Begin VB.Label Label40 
            Caption         =   "Valores en S/.:"
            Height          =   315
            Left            =   4350
            TabIndex        =   77
            Top             =   2190
            Width           =   1185
         End
         Begin VB.Label Label38 
            Caption         =   "Total Préstamo"
            Height          =   285
            Left            =   120
            TabIndex        =   75
            Top             =   3900
            Width           =   1575
         End
         Begin VB.Label Label37 
            Caption         =   "Valores en S/.:"
            Height          =   315
            Left            =   4350
            TabIndex        =   74
            Top             =   2850
            Width           =   1185
         End
         Begin VB.Label Label34 
            Caption         =   "Monto Préstamo"
            Height          =   285
            Left            =   120
            TabIndex        =   72
            Top             =   3240
            Width           =   1575
         End
         Begin VB.Label Label27 
            Caption         =   "Valores en S/.:"
            Height          =   315
            Left            =   4350
            TabIndex        =   65
            Top             =   3240
            Width           =   1185
         End
         Begin VB.Label Label26 
            Caption         =   "Valores en S/.:"
            Height          =   315
            Left            =   4350
            TabIndex        =   64
            Top             =   1860
            Width           =   1185
         End
         Begin VB.Label Label25 
            Caption         =   "Monto BMS"
            Height          =   285
            Left            =   510
            TabIndex        =   63
            Top             =   2850
            Width           =   1245
         End
         Begin VB.Label Label24 
            Caption         =   "Bono BBP:"
            Height          =   285
            Left            =   510
            TabIndex        =   62
            Top             =   1860
            Width           =   1245
         End
         Begin VB.Label Label13 
            Caption         =   "Valores en S/.:"
            Height          =   285
            Left            =   4350
            TabIndex        =   61
            Top             =   2520
            Width           =   1185
         End
         Begin VB.Label Label15 
            Caption         =   "Monto AFP (25%)"
            Height          =   285
            Left            =   510
            TabIndex        =   51
            Top             =   2520
            Width           =   1245
         End
         Begin VB.Label Label35 
            Caption         =   "Valor Inmueble:"
            Height          =   285
            Left            =   510
            TabIndex        =   50
            Top             =   510
            Width           =   1185
         End
         Begin VB.Label Label2 
            Caption         =   "Aporte Propio:"
            Height          =   285
            Left            =   510
            TabIndex        =   49
            Top             =   1560
            Width           =   1245
         End
         Begin VB.Label Label17 
            Caption         =   "Valores en S/.:"
            Height          =   315
            Left            =   4350
            TabIndex        =   48
            Top             =   1530
            Width           =   1185
         End
         Begin VB.Label Label11 
            Caption         =   "Valores en S/.:"
            Height          =   315
            Left            =   4350
            TabIndex        =   47
            Top             =   1200
            Width           =   1185
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   615
         Left            =   30
         TabIndex        =   52
         Top             =   30
         Width           =   11625
         _Version        =   65536
         _ExtentX        =   20505
         _ExtentY        =   1085
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   255
            Left            =   660
            TabIndex        =   58
            Top             =   30
            Width           =   6165
            _Version        =   65536
            _ExtentX        =   10874
            _ExtentY        =   450
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   255
            Left            =   660
            TabIndex        =   59
            Top             =   300
            Width           =   6165
            _Version        =   65536
            _ExtentX        =   10874
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Datos del Crédito"
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
            Picture         =   "OpeTra_frm_159.frx":0A62
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   795
         Left            =   30
         TabIndex        =   53
         Top             =   1320
         Width           =   11625
         _Version        =   65536
         _ExtentX        =   20505
         _ExtentY        =   1402
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
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   1860
            TabIndex        =   54
            Top             =   90
            Width           =   9645
            _Version        =   65536
            _ExtentX        =   17013
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "CREDITO HIPOTECARIO MICASITA"
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1860
            TabIndex        =   55
            Top             =   420
            Width           =   9645
            _Version        =   65536
            _ExtentX        =   17013
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
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
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   120
            TabIndex        =   57
            Top             =   450
            Width           =   1755
         End
         Begin VB.Label Label21 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   120
            TabIndex        =   56
            Top             =   120
            Width           =   1755
         End
      End
   End
End
Attribute VB_Name = "frm_SolCre_54"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_EmpSeg()      As moddat_tpo_Genera
Dim l_arr_Modali()      As moddat_tpo_Genera
Dim l_arr_CuoExt()      As moddat_tpo_Genera
Dim l_arr_DiaPag()      As moddat_tpo_Genera
Dim l_arr_ConHip()      As moddat_tpo_Genera
Dim l_arr_EjeSeg()      As moddat_tpo_Genera
Dim l_arr_ParPrd()      As moddat_tpo_Genera
Dim l_arr_InsFin()      As moddat_tpo_Genera
Dim l_arr_TipEva()      As moddat_tpo_Genera
Dim l_dbl_TipCam        As Double
Dim l_int_GraMax        As Integer
Dim l_int_FlgAfeBV      As Integer
Dim l_int_FlgTipAfe     As Integer
Dim l_dbl_ValAfeBV      As Double
Dim l_dbl_MPSMS         As Double

Private Sub chk_Gastos_Click(Value As Integer)
   If chk_Gastos.Value = False Then
      pnl_ValGas.Caption = "0.00 "
   End If
   
   If chk_Gastos.Value = True Then
      If modatecli_g_arr_DatInm(1).DatInm_InmIde = 2 Then
         MsgBox "Debe ingresar Datos del Inmueble.", vbExclamation, modgen_g_str_NomPlt
         chk_Gastos.Value = 0
         Call gs_SetFocus(chk_Gastos)
         Exit Sub
      Else
         Call fs_Calcul
         Call fs_Calcular_Prestamo
         Call fs_Calcular_GCierre
      End If
   Else
      Call fs_Calcul
      Call fs_Calcular_Prestamo
      Call fs_Calcular_GCierre
   End If
End Sub

Private Sub chk_Gastos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipEva)
   End If
End Sub

Private Sub chk_Gastos_LostFocus()
   Call fs_Calcul
   Call fs_Calcular_Prestamo
   Call fs_Calcular_GCierre
End Sub

'**************************************************************************************************
'* BOTONES
'**************************************************************************************************
Private Sub cmd_Grabar_Click()
Dim r_int_Contad           As Integer
Dim r_int_FlgDoc           As Integer
Dim r_dbl_ValMin_ComVta    As Double
Dim r_dbl_ValMax_ComVta    As Double
Dim r_dbl_PorMin_ApoPro    As Double
Dim r_dbl_PorMax_ApoPro    As Double
Dim r_dbl_PorMax_MtoPre    As Double
Dim r_dbl_ValMin_MtoPre    As Double
Dim r_dbl_ValMax_MtoPre    As Double
Dim r_int_EdaMin           As Integer
Dim r_int_EdaMax           As Integer
Dim r_int_EdaAct           As Integer
Dim r_dbl_Aho_ApoMin       As Double
Dim r_dbl_Aho_ApoTp1       As Double
Dim r_dbl_Aho_ApoTp2       As Double
Dim r_dbl_Aho_ApoRgI       As Double
Dim r_dbl_Aho_ApoRgF       As Double
Dim r_dbl_ApoMin           As Double
Dim r_dbl_CuoAho           As Double
Dim r_dbl_Aho_CuoMin       As Double
Dim r_dbl_Ini_ApoMin       As Double
Dim r_dbl_Ini_PlaMin       As Double
Dim r_dbl_Ini_PlaMax       As Double
Dim r_dbl_PrcMin           As Double
Dim r_dbl_PrcMax           As Double
Dim r_dbl_ValAfe           As Double

   Call ipp_ApoPro_Change
   Call fs_Calcul
   Call fs_Calcular_Prestamo
   Call fs_Calcular_GCierre

   If cmb_TipEva.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Evaluación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipEva)
      Exit Sub
   End If
   If cmb_PriViv.ListIndex = -1 Then
      MsgBox "Debe seleccionar si es Primera Vivienda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PriViv)
      Exit Sub
   End If
   
   If modatecli_g_arr_DatInm(1).DatInm_FlgEst <> 0 Then
      If modatecli_g_arr_DatInm(1).DatInm_FlgEst = 1 Then
         If CDbl(ipp_ValEst.Text) = 0 Then
            MsgBox "Debe ingresar el valor del estacionamiento.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ValEst)
            Exit Sub
         End If
      Else
         If CDbl(ipp_ValEst.Text) <> 0 Then
            MsgBox "El valor del estacionamiento debe de ser cero.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ValEst)
            Exit Sub
         End If
      End If
   End If
   
   'If moddat_g_str_CodPrd = "001" Then
   '   If cmb_TipEva.ItemData(cmb_TipEva.ListIndex) = 4 Then
   '      MsgBox "El Producto no acepta este Tipo de Evaluación.", vbExclamation, modgen_g_str_NomPlt
   '      Call gs_SetFocus(cmb_TipEva)
   '      Exit Sub
   '   End If
   'ElseIf moddat_g_str_CodPrd = "003" Or moddat_g_str_CodPrd = "002" Then
   '   If cmb_TipEva.ItemData(cmb_TipEva.ListIndex) = 3 Or cmb_TipEva.ItemData(cmb_TipEva.ListIndex) = 4 Then
   '      MsgBox "El Producto no acepta este Tipo de Evaluación.", vbExclamation, modgen_g_str_NomPlt
   '      Call gs_SetFocus(cmb_TipEva)
   '      Exit Sub
   '   End If
   'End If
   
   If CInt(l_arr_TipEva(cmb_TipEva.ListIndex + 1).Genera_Codigo) = 3 Or CInt(l_arr_TipEva(cmb_TipEva.ListIndex + 1).Genera_Codigo) = 4 Then
      If modatecli_g_arr_DatInm(1).DatInm_InmIde = 2 Then
         MsgBox "Este Tipo de Evaluación exige que el Inmueble sea identificado. Registre la información del Inmueble.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipEva)
         Exit Sub
      End If
   End If
   
   'Buscar Parámetros en Productos
   r_dbl_ValMin_ComVta = 0
   r_dbl_ValMax_ComVta = 0
   r_dbl_PorMin_ApoPro = 0
   r_dbl_PorMax_ApoPro = 0
   r_dbl_PorMax_MtoPre = 0
   r_dbl_ValMin_MtoPre = 0
   r_dbl_ValMax_MtoPre = 0
   r_dbl_PrcMin = 0
   r_dbl_PrcMax = 0
   
   'Edad Minima del cliente por producto y subproducto
   If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "011") Then
      r_int_EdaMin = l_arr_ParPrd(1).Genera_ValMin
   End If
   
   'Edad Máxima del Cliente (incluido plazo prestamo)
   If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "012") Then
      r_int_EdaMax = l_arr_ParPrd(1).Genera_Cantid
   End If
   
   'Para obtener Valor Máximo del Inmueble
   Select Case moddat_g_str_CodPrd > 0
      Case InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd)   '"001"                    'En UIT (Mínimo y Máximo)
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "022") Then
            r_dbl_ValMin_ComVta = l_arr_ParPrd(1).Genera_ValMin * moddat_gf_Consulta_ParVal("001", "002")
            r_dbl_ValMax_ComVta = l_arr_ParPrd(1).Genera_ValMax * moddat_gf_Consulta_ParVal("001", "002")
         End If
      
      Case InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd)  '"002", "011"             'En Montos
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "021") Then
            r_dbl_ValMax_ComVta = l_arr_ParPrd(1).Genera_Cantid
         End If
      
      Case InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd)   '"003"                    'En UIT
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "022") Then
            r_dbl_ValMin_ComVta = l_arr_ParPrd(1).Genera_ValMin * moddat_gf_Consulta_ParVal("001", "002")
            r_dbl_ValMax_ComVta = l_arr_ParPrd(1).Genera_ValMax * moddat_gf_Consulta_ParVal("001", "002")
         End If
      
      Case InStr(moddat_g_str_AgrMIHG, moddat_g_str_CodPrd)  '"004"                    'En UIT (Mínimo y Máximo)
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "022") Then
            r_dbl_ValMin_ComVta = l_arr_ParPrd(1).Genera_ValMin * moddat_gf_Consulta_ParVal("001", "002")
            r_dbl_ValMax_ComVta = l_arr_ParPrd(1).Genera_ValMax * moddat_gf_Consulta_ParVal("001", "002")
         End If
   
      Case InStr(moddat_g_str_Agr2MIC, moddat_g_str_CodPrd) Or InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) Or InStr(moddat_g_str_Agr2FMV, moddat_g_str_CodPrd) '"007", "009", "010", "013", "014", "015", "016", "017", "018", "019", "021", "022", "023"    'En UIT (Mínimo y Máximo)
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "022") Then
            r_dbl_ValMin_ComVta = l_arr_ParPrd(1).Genera_ValMin * moddat_gf_Consulta_ParVal("001", "002")
            r_dbl_ValMax_ComVta = l_arr_ParPrd(1).Genera_ValMax * moddat_gf_Consulta_ParVal("001", "002")
            r_dbl_ValMin_ComVta = Format(r_dbl_ValMin_ComVta, "###,###,##0.00")
            r_dbl_ValMax_ComVta = Format(r_dbl_ValMax_ComVta, "###,###,##0.00")
         End If
   
   End Select
   
   'Para obtener % Mínimo de Aporte Propio
   If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "022") Then
      r_dbl_PorMin_ApoPro = l_arr_ParPrd(1).Genera_Cantid
   End If
   
   'Para obtener % Máximo de Monto de Préstamo
   If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "023") Then
      r_dbl_PorMax_MtoPre = l_arr_ParPrd(1).Genera_Cantid
   End If
   
   'Para obtener Monto Mínimo y Máximo de Préstamo
   If InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then    '"002" "011"
      'En Montos (Maximo)
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "024") Then
         r_dbl_ValMax_MtoPre = l_arr_ParPrd(1).Genera_Cantid
      End If
      
      'En Montos (Minimo)
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "026") Then
         r_dbl_ValMin_MtoPre = l_arr_ParPrd(1).Genera_Cantid
      End If
      
   ElseIf InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then   '"001" "003"
      'En UIT
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "023") Then
         r_dbl_ValMax_MtoPre = l_arr_ParPrd(1).Genera_Cantid * moddat_gf_Consulta_ParVal("001", "002")
      End If
      
   ElseIf InStr(moddat_g_str_AgrMIHG, moddat_g_str_CodPrd) > 0 Then  '"004"
      'Porcentaje para Valor Minimo
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "024") Then
         r_dbl_PrcMin = l_arr_ParPrd(1).Genera_Cantid
      End If
   
      'Porcentaje para Valor Máximo
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "025") Then
         r_dbl_PrcMax = l_arr_ParPrd(1).Genera_Cantid
      End If
   
      'En UIT
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "023") Then
         r_dbl_ValMin_MtoPre = l_arr_ParPrd(1).Genera_ValMin * moddat_gf_Consulta_ParVal("001", "002") * r_dbl_PrcMin / 100
         r_dbl_ValMax_MtoPre = l_arr_ParPrd(1).Genera_ValMax * moddat_gf_Consulta_ParVal("001", "002") * r_dbl_PrcMax / 100
      End If
      
   ElseIf InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr2FMV, moddat_g_str_CodPrd) > 0 Then '"007" "009" "010" "012" "013" "014" "015" "016" "017" "018" "019" "021" "022" "023"
      'Porcentaje para Valor Minimo
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "024") Then
         r_dbl_PrcMin = l_arr_ParPrd(1).Genera_Cantid
      End If
   
      'Porcentaje para Valor Máximo
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "025") Then
         r_dbl_PrcMax = l_arr_ParPrd(1).Genera_Cantid
      End If
   
      'En UIT
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "023") Then
         r_dbl_ValMin_MtoPre = l_arr_ParPrd(1).Genera_ValMin * moddat_gf_Consulta_ParVal("001", "002") * r_dbl_PrcMin / 100
         r_dbl_ValMax_MtoPre = l_arr_ParPrd(1).Genera_ValMax * moddat_gf_Consulta_ParVal("001", "002") * r_dbl_PrcMax / 100
         r_dbl_ValMin_MtoPre = Format(r_dbl_ValMin_MtoPre, "###,###,##0.00")
         r_dbl_ValMax_MtoPre = Format(r_dbl_ValMax_MtoPre, "###,###,##0.00")
      End If
   End If
   
   If CDbl(ipp_ComVta.Text) = 0 Then
      MsgBox "Debe ingresar el Valor de Compra-Venta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ComVta)
      Exit Sub
   End If
   
   'Validando Valor de Compra Venta
   If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then  '"001" "004" "003" "007" "009" "010" "013" "014" "015" "016" "017" "018" "019" "021" "022" "023"
      If CDbl(pnl_ValTot.Caption) < r_dbl_ValMin_ComVta Then  '  pnl_ComVta_Sol.Caption
         MsgBox "El Valor de Compra-Venta no cubre el mínimo requerido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ComVta)
         Exit Sub
      End If
      
      If CDbl(pnl_ValTot.Caption) > r_dbl_ValMax_ComVta Then   'pnl_ComVta_Sol
         MsgBox "El Valor de Compra-Venta excede el permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ComVta)
         Exit Sub
      End If
      
   ElseIf InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then   '"002" "011"
      If CDbl(pnl_ValTot.Caption) > r_dbl_ValMax_ComVta Then                                 ' pnl_ComVta_Dol.Caption
         MsgBox "El Valor de Compra-Venta excede el permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ComVta)
         Exit Sub
      End If
   End If

   If CDbl(ipp_ApoPro.Text) + CDbl(ipp_MtoAFP.Value) = 0 Then
      MsgBox "Debe ingresar el Aporte Propio y/o Monto AFP.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ApoPro)
      Exit Sub
   End If
   
   If CDbl(ipp_ApoPro.Text) > CDbl(pnl_ValTot.Caption) Then    'ipp_ComVta.Text
      MsgBox "El Aporte Propio no puede ser mayor al Valor de Compra Venta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ApoPro)
      Exit Sub
   End If
   
   If moddat_g_str_CodPrd = "023" Then
      If CDbl(Format((CDbl(ipp_ApoPro.Value) + CDbl(ipp_MtoAFP.Value) + CDbl(pnl_FmvBbp.Caption) + CDbl(pnl_MefPbp.Caption) + CDbl(pnl_MtoBMS.Caption)) / CDbl(pnl_ValTot.Caption) * 100, "###0.0000")) < r_dbl_PorMin_ApoPro Then     'ipp_ComVta.Value
         MsgBox "El Aporte Propio no cubre el mínimo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
   Else
      If CDbl(Format((CDbl(ipp_ApoPro.Value) + CDbl(ipp_MtoAFP.Value)) / CDbl(pnl_ValTot.Caption) * 100, "###0.0000")) < r_dbl_PorMin_ApoPro Then 'ipp_ComVta.Value
         MsgBox "El Aporte Propio no cubre el mínimo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
   End If
   
   If CDbl(pnl_MtoPre.Caption) / CDbl(pnl_ValTot.Caption) * 100 > r_dbl_PorMax_MtoPre Then  'ipp_ComVta.Text
      MsgBox "El Monto del Aporte Propio no cubre el mínimo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ApoPro)
      Exit Sub
   End If

   If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then '"003" "004" "007" "009" "010" "012" "013" "014" "015" "016" "017" "018" "019" "021" "022" "023"
      'Para obtener % Maximo de Aporte Propio
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "027") Then
         r_dbl_PorMax_ApoPro = l_arr_ParPrd(1).Genera_Cantid
      End If
      
      If CDbl(Format(CDbl(ipp_ApoPro.Value) / CDbl(pnl_ValTot.Caption) * 100, "###0.000000")) > r_dbl_PorMax_ApoPro Then      'ipp_ComVta.Value
         MsgBox "El Aporte Propio sobrepasa el máximo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
   End If
   
   'Validando Monto de Préstamo
   If InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then  '"002" "011"
      If moddat_g_int_TipMon = 1 Then
         If CDbl(pnl_TotPre.Caption) < r_dbl_ValMin_MtoPre Then
            MsgBox "El Monto del Préstamo no cubre el mínimo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
         
         If CDbl(pnl_TotPre.Caption) > r_dbl_ValMax_MtoPre Then
            MsgBox "El Monto del Préstamo excede el máximo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
      Else
'         If CDbl(pnl_MtoPre_Dol.Caption) < r_dbl_ValMin_MtoPre Then
'            MsgBox "El Monto del Préstamo no cubre el mínimo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
'            Call gs_SetFocus(ipp_ApoPro)
'            Exit Sub
'         End If
'         If CDbl(pnl_MtoPre_Dol.Caption) > r_dbl_ValMax_MtoPre Then
'            MsgBox "El Monto del Préstamo excede el máximo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
'            Call gs_SetFocus(ipp_ApoPro)
'            Exit Sub
'         End If
      End If
      
   ElseIf InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then  '"001" "003"
      If CDbl(pnl_TotPre.Caption) > r_dbl_ValMax_MtoPre Then
         MsgBox "El Monto del Préstamo excede el máximo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
      
   ElseIf InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then '"004" "007" "009" "010" "012" "013" "014" "015" "016" "017" "018" "019" "021" "022" "023" "024" "025"
      If CDbl(pnl_TotPre.Caption) < r_dbl_ValMin_MtoPre Then
         MsgBox "El Monto del Préstamo no cubre el mínimo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
      If CDbl(pnl_TotPre.Caption) > r_dbl_ValMax_MtoPre Then
         MsgBox "El Monto del Préstamo excede el máximo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
   End If

   If CDbl(ipp_PlaAno.Text) = 0 Then
      MsgBox "Debe ingresar el Plazo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaAno)
      Exit Sub
   End If
   
   If Not (CInt(ipp_PlaAno.Text) >= ipp_PlaAno.MinValue And CInt(ipp_PlaAno.Text) <= ipp_PlaAno.MaxValue) Then
      MsgBox "El Plazo indicado no se ajusta a los Parámetros permitidos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaAno)
      Exit Sub
   End If

   r_int_EdaAct = CInt(Left(gs_CalcularEdad(CDate(moddat_g_str_FecNac_Tit), date), 2))
   
   If r_int_EdaMin > r_int_EdaAct Then
      MsgBox "La Edad del Titular es menor que la minima permitida según parámetros.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaAno)
      Exit Sub
   End If
   
   If r_int_EdaAct + CInt(ipp_PlaAno.Text) > r_int_EdaMax Then
      MsgBox "La Edad del Titular más el Plazo del Préstamo excede el parámetro permitido. El Plazo máximo podría ser de " & CStr(r_int_EdaMax - r_int_EdaAct) & " años.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaAno)
      Exit Sub
   End If

   If Not (CInt(ipp_PerGra.Text) >= ipp_PerGra.MinValue And CInt(ipp_PerGra.Text) <= ipp_PerGra.MaxValue) Then
      MsgBox "El Período de Gracia no se ajusta a los Parámetros permitidos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerGra)
      Exit Sub
   End If

   If cmb_CuoDbl.ListIndex = -1 Then
      MsgBox "Debe seleccionar si desea cuotas extraordinarias (dobles).", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CuoDbl)
      Exit Sub
   End If
   
   If cmb_EmpSeg.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Empresa de Seguros.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_EmpSeg)
      Exit Sub
   End If

   If cmb_SegDes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Seguro de Desgravamen.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_SegDes)
      Exit Sub
   End If
   
   If cmb_TasEsp.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Tasa Especial.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TasEsp)
      Exit Sub
   End If
   
   If moddat_g_int_EstCiv <> 2 And moddat_g_int_EstCiv <> 5 Then
      If cmb_SegDes.ItemData(cmb_SegDes.ListIndex) = 12 Then
         MsgBox "El Cliente no requiere tomar Seguro de Desgravamen Mancomunado.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_SegDes)
         Exit Sub
      End If
   End If
   
   'Si cliente complementa Renta
   If moddat_g_int_ComRta = 1 Then
      If cmb_SegDes.ItemData(cmb_SegDes.ListIndex) <> 12 Then
         MsgBox "El Cliente presenta Complemento de Renta debe seleccionar el Seguro de Desgravamen Mancomunado.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_SegDes)
         Exit Sub
      End If
   End If
   
   If cmb_SegDes.ItemData(cmb_SegDes.ListIndex) = 12 Then
      r_int_EdaAct = CInt(Left(gs_CalcularEdad(CDate(moddat_g_str_FecNac_Cyg), date), 2))
      
      If r_int_EdaAct + CInt(ipp_PlaAno.Text) > r_int_EdaMax Then
         MsgBox "La Edad del Cónyuge más el Plazo del Préstamo excede el parámetro permitido. El Plazo máximo podría ser de " & CStr(r_int_EdaMax - r_int_EdaAct) & " años.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_PlaAno)
         Exit Sub
      End If
   
      If r_int_EdaMin > r_int_EdaAct Then
         MsgBox "La Edad del Cónyuge es menor que la minima permitida según parámetros.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_PlaAno)
         Exit Sub
      End If
   End If
   
   If cmb_DiaPag.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Día de Pago.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_DiaPag)
      Exit Sub
   End If
   
   'Evaluación Normal
   If CInt(l_arr_TipEva(cmb_TipEva.ListIndex + 1).Genera_Codigo) = 1 Then
      'Validando que Clientes de Provincias cumplan con Aporte Inicial mínimo
      r_dbl_ApoMin = (CDbl(ipp_ApoPro.Text) + CDbl(Me.ipp_MtoAFP.Value) + CDbl(pnl_FmvBbp.Caption) + CDbl(pnl_MefPbp.Caption) + CDbl(pnl_MtoBMS.Caption)) / CDbl(pnl_ValTot.Caption) * 100       'ipp_ComVta.Value
      
      If moddat_g_str_UbiGeo <> "1501" And moddat_g_str_UbiGeo <> "0701" Then
         r_dbl_Ini_ApoMin = 0
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "025") Then
            r_dbl_Ini_ApoMin = l_arr_ParPrd(1).Genera_Cantid
         End If
         
         If r_dbl_ApoMin < r_dbl_Ini_ApoMin Then
            MsgBox "Cliente de Provincias. El Aporte Inicial es menor al Aporte Inicial mínimo requerido. (" & CStr(r_dbl_Ini_ApoMin) & "%).", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
      End If
   End If
   
   'Ahorro Programado
   If CInt(l_arr_TipEva(cmb_TipEva.ListIndex + 1).Genera_Codigo) = 2 Then
      'Clientes de Provincia no tienen acceso a este Tipo de Evaluación
      'If moddat_g_str_UbiGeo <> "1501" And moddat_g_str_UbiGeo <> "0701" Then
      '   MsgBox "Este tipo de Evaluación sólo está permitida para clientes que residen en Lima Metropolitana o Callao.", vbExclamation, modgen_g_str_NomPlt
      '   Call gs_SetFocus(cmb_InsFin)
      '   Exit Sub
      'End If
   
      If cmb_InsFin.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Institución Financiera donde tiene sus ahorros.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_InsFin)
         Exit Sub
      End If
      If cmb_MonAho.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Moneda de su ahorro.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_MonAho)
         Exit Sub
      End If
      If ipp_MtoAho.Value = 0 Then
         MsgBox "Debe ingresar el Monto Mínimo Mensual de su Ahorro.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_MtoAho)
         Exit Sub
      End If
      If ipp_MesAho.Value = 0 Then
         MsgBox "Debe ingresar los Meses Ahorrados.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_MesAho)
         Exit Sub
      End If
      
      r_dbl_CuoAho = 0
      If cmb_MonAho.ItemData(cmb_MonAho.ListIndex) = 1 Then
         r_dbl_CuoAho = CDbl(ipp_MtoAho.Text) / l_dbl_TipCam
      ElseIf cmb_MonAho.ItemData(cmb_MonAho.ListIndex) = 2 Then
         r_dbl_CuoAho = CDbl(ipp_MtoAho.Text)
      End If
      r_dbl_CuoAho = CDbl(Format(r_dbl_CuoAho, "###,##0.00"))
      
      r_dbl_Aho_CuoMin = 0
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "001") Then
         r_dbl_Aho_CuoMin = l_arr_ParPrd(1).Genera_Cantid
      End If
      
      If r_dbl_CuoAho < r_dbl_Aho_CuoMin Then
         MsgBox "El Importe de la Cuota Mensual Ahorrada no cumple con el mínimo requerido.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_MtoAho)
         Exit Sub
      End If
      
      r_dbl_ApoMin = (CDbl(ipp_ApoPro.Text) + CDbl(ipp_MtoAFP.Value) + CDbl(pnl_FmvBbp.Caption) + CDbl(pnl_MefPbp.Caption) + CDbl(pnl_MtoBMS.Caption)) / CDbl(pnl_ValTot.Caption) * 100   'ipp_ComVta.Text
      r_dbl_Aho_ApoTp1 = 0
      
      If InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then     '"002" "011" Then
         If r_dbl_ApoMin < 20 Then
            MsgBox "El Aporte Propio no cubre el mínimo permitido para el Tipo de Evaluación. (20%).", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
         
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "012") Then
            r_dbl_Aho_ApoTp1 = l_arr_ParPrd(1).Genera_Cantid
         End If
      
         If CInt(ipp_MesAho.Text) < r_dbl_Aho_ApoTp1 Then
            MsgBox "El Cliente no cumple con el Tiempo Mínimo de Ahorro requerido. (" & CStr(r_dbl_Aho_ApoTp1) & " meses).", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_MesAho)
            Exit Sub
         End If
      ElseIf InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then  '"003"
         If r_dbl_ApoMin >= 20 And r_dbl_ApoMin < 30 Then
            If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "013") Then
               r_dbl_Aho_ApoTp1 = l_arr_ParPrd(1).Genera_Cantid
            End If
            
            If CInt(ipp_MesAho.Text) < r_dbl_Aho_ApoTp1 Then
               MsgBox "El Cliente no cumple con el Tiempo Mínimo de Ahorro requerido. (" & CStr(r_dbl_Aho_ApoTp1) & " meses).", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_MesAho)
               Exit Sub
            End If
         
         ElseIf r_dbl_ApoMin >= 30 Then
            If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "014") Then
               r_dbl_Aho_ApoTp1 = l_arr_ParPrd(1).Genera_Cantid
            End If
            
            If CInt(ipp_MesAho.Text) < r_dbl_Aho_ApoTp1 Then
               MsgBox "El Cliente no cumple con el Tiempo Mínimo de Ahorro requerido. (" & CStr(r_dbl_Aho_ApoTp1) & " meses).", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_MesAho)
               Exit Sub
            End If
         
         Else
            MsgBox "El Aporte Propio no cubre el mínimo permitido para el Tipo de Evaluación. (20%).", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
      ElseIf InStr(moddat_g_str_AgrMIHG, moddat_g_str_CodPrd) > 0 Then '"004"
         If r_dbl_ApoMin >= 10 And r_dbl_ApoMin < 20 Then
            If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "011") Then
               r_dbl_Aho_ApoTp1 = l_arr_ParPrd(1).Genera_Cantid
            End If
            
            If CInt(ipp_MesAho.Text) < r_dbl_Aho_ApoTp1 Then
               MsgBox "El Cliente no cumple con el Tiempo Mínimo de Ahorro requerido. (" & CStr(r_dbl_Aho_ApoTp1) & " meses).", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_MesAho)
               Exit Sub
            End If
         
         ElseIf r_dbl_ApoMin >= 20 Then
            If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "012") Then
               r_dbl_Aho_ApoTp1 = l_arr_ParPrd(1).Genera_Cantid
            End If
            
            If CInt(ipp_MesAho.Text) < r_dbl_Aho_ApoTp1 Then
               MsgBox "El Cliente no cumple con el Tiempo Mínimo de Ahorro requerido. (" & CStr(r_dbl_Aho_ApoTp1) & " meses).", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_MesAho)
               Exit Sub
            End If
         
         Else
            MsgBox "El Aporte Propio no cubre el mínimo permitido para el Tipo de Evaluación. (10%).", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
         
      ElseIf InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr2FMV, moddat_g_str_CodPrd) > 0 Then  '"007" "009" "010" "012" "013" "014" "015" "016" "017" "018" "019" "021" "022" "023"
         If r_dbl_ApoMin >= 10 And r_dbl_ApoMin < 20 Then
            If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "011") Then
               r_dbl_Aho_ApoTp1 = l_arr_ParPrd(1).Genera_Cantid
            End If
            
            If CInt(ipp_MesAho.Text) < r_dbl_Aho_ApoTp1 Then
               MsgBox "El Cliente no cumple con el Tiempo Mínimo de Ahorro requerido. (" & CStr(r_dbl_Aho_ApoTp1) & " meses).", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_MesAho)
               Exit Sub
            End If
         
         ElseIf r_dbl_ApoMin >= 20 Then
            If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "012") Then
               r_dbl_Aho_ApoTp1 = l_arr_ParPrd(1).Genera_Cantid
            End If
            
            If CInt(ipp_MesAho.Text) < r_dbl_Aho_ApoTp1 Then
               MsgBox "El Cliente no cumple con el Tiempo Mínimo de Ahorro requerido. (" & CStr(r_dbl_Aho_ApoTp1) & " meses).", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_MesAho)
               Exit Sub
            End If
         
         Else
            MsgBox "El Aporte Propio no cubre el mínimo permitido para el Tipo de Evaluación. (10%).", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
      End If
   End If
   
   '30%-35%
   If CInt(l_arr_TipEva(cmb_TipEva.ListIndex + 1).Genera_Codigo) = 3 Then
      r_dbl_ApoMin = CDbl(Format((CDbl(ipp_ApoPro.Text) + CDbl(ipp_MtoAFP.Value) + CDbl(pnl_FmvBbp.Caption) + CDbl(pnl_MefPbp.Caption) + CDbl(pnl_MtoBMS.Caption)) / CDbl(pnl_ValTot.Caption) * 100, "###0.00"))    'ipp_ComVta.Value
      
      If modatecli_g_arr_DatInm(1).DatInm_PryMCs = 1 Then
         r_dbl_Ini_ApoMin = 0
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "053", "001") Then
            r_dbl_Ini_ApoMin = l_arr_ParPrd(1).Genera_Cantid
         End If
         
         If r_dbl_ApoMin < r_dbl_Ini_ApoMin Then
            MsgBox "El Aporte Inicial es menor al Aporte Inicial mínimo requerido.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
      Else
         r_dbl_Ini_ApoMin = 0
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "053", "002") Then
            r_dbl_Ini_ApoMin = l_arr_ParPrd(1).Genera_Cantid
         End If
         
         If r_dbl_ApoMin < r_dbl_Ini_ApoMin Then
            MsgBox "El Aporte Inicial es menor al Aporte Inicial mínimo requerido.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
      End If
      
      'If moddat_g_str_UbiGeo <> "1501" And moddat_g_str_UbiGeo <> "0701" Then
      '   r_dbl_Ini_ApoMin = 0
      '   If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "053", "004") Then
      '      r_dbl_Ini_ApoMin = l_arr_ParPrd(1).Genera_Cantid
      '   End If
      '
      '   If r_dbl_ApoMin < r_dbl_Ini_ApoMin Then
      '      MsgBox "Cliente de Provincias. El Aporte Inicial es menor al Aporte Inicial mínimo requerido. (" & CStr(r_dbl_Ini_ApoMin) & "%).", vbExclamation, modgen_g_str_NomPlt
      '      Call gs_SetFocus(ipp_ApoPro)
      '      Exit Sub
      '   End If
      'End If
   End If
   
   '50% Inicial Sin Evaluación
   If CInt(l_arr_TipEva(cmb_TipEva.ListIndex + 1).Genera_Codigo) = 4 Then
      r_dbl_ApoMin = CDbl(ipp_ApoPro.Text) / CDbl(ipp_ComVta.Text) * 100
      
      r_dbl_Ini_ApoMin = 0
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "054", "001") Then
         r_dbl_Ini_ApoMin = l_arr_ParPrd(1).Genera_Cantid
      End If
      
      If r_dbl_ApoMin < r_dbl_Ini_ApoMin Then
         MsgBox "El Aporte Inicial es menor al Aporte Inicial mínimo requerido.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
   End If
   
   If cmb_ConHip.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Consejero Hipotecario.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_ConHip)
      Exit Sub
   End If
   
   If cmb_EjeSeg.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Ejecutivo de Seguimiento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_EjeSeg)
      Exit Sub
   End If
   
   'Validando Documentos a Recibir
   grd_Listad.Redraw = False
   r_int_FlgDoc = 1
   
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      grd_Listad.Row = r_int_Contad
      
      grd_Listad.Col = 1
      If Trim(grd_Listad.Text) = "X" Then
         r_int_FlgDoc = 2
         Exit For
      End If
   Next r_int_Contad
   
   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)
   
   If r_int_FlgDoc = 1 Then
      MsgBox "Debe seleccionar los Documentos Crediticios que han sido recibidos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(grd_Listad)
      Exit Sub
   End If
   
   'Valida los datos del bono verde
   If InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then
      If l_int_FlgAfeBV = 1 And l_int_FlgTipAfe <> 0 And l_dbl_ValAfeBV <> 0 And CDbl(pnl_MtoBMS.Caption) = 0 Then
         MsgBox "Verificar montos ingresados, está afecto a bono verde", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ComVta)
         Exit Sub
      End If
   End If
   
   If ipp_ComVta.Text <= modatecli_g_dbl_MtoFin Then
      If l_int_FlgAfeBV = 1 And l_int_FlgTipAfe <> 0 Or l_dbl_ValAfeBV <> 0 Then r_dbl_ValAfe = modatecli_g_dbl_BMSTas
   Else
      If l_int_FlgAfeBV = 0 Or l_int_FlgTipAfe = 0 Or l_dbl_ValAfeBV = 0 Then l_dbl_ValAfeBV = 0
      r_dbl_ValAfe = l_dbl_ValAfeBV
   End If
   
   If InStr(moddat_g_str_AgrTMIC, moddat_g_str_CodPrd) = 0 Then
      If Format((CDbl(l_dbl_MPSMS) * r_dbl_ValAfe) / (1 + r_dbl_ValAfe), "###,###,##0.00") & " " <> Me.pnl_MtoBMS.Caption Then  'pnl_MPSBMS.Caption
         MsgBox "El Bono Mivivienda Sostenible no es correcto", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ComVta)
         Exit Sub
      End If
   End If
   
   Dim r_int_Resul As Integer
   
   'Valida si tiene Ingresado los datos del inmueble para hallar gastos de cierre
   If chk_Gastos.Value = True Then
      If modatecli_g_arr_DatInm(1).DatInm_InmIde = 2 Then
         MsgBox "Debe ingresar Datos del Inmueble.", vbExclamation, modgen_g_str_NomPlt
         chk_Gastos.Value = False
         Exit Sub
      Else
         'Valida si el proyecto ha sido aprobado
         r_int_Resul = fs_Valida_PryAsg
         
         If r_int_Resul <> 1 Then
            MsgBox "El Proyecto seleccionado no se encuentra aprobado. Los Gastos de Cierre se deberán ingresar manualmente.", vbExclamation, modgen_g_str_NomPlt
            chk_Gastos.Value = False
            Exit Sub
         End If
         
         'Valida si hay Notaria y Empresa Tasadora asignada al Proyecto
         r_int_Resul = gf_Valida_GastoCierre(moddat_g_str_CodPrd, modatecli_g_arr_DatInm(1).DatInm_CodPry)
         
         If r_int_Resul = 1 Then
            MsgBox "El proyecto asociado no tiene empresa de peritaje asignado, favor actualizar información en la plataforma de Operaciones.", vbExclamation, modgen_g_str_NomPlt
            chk_Gastos.Value = False
            Exit Sub
         ElseIf r_int_Resul = 2 Then
            MsgBox "El proyecto asociado no tiene notaría asignada, favor actualizar información en la plataforma de Operaciones.", vbExclamation, modgen_g_str_NomPlt
            chk_Gastos.Value = False
            Exit Sub
         ElseIf r_int_Resul = 3 Then
            'MsgBox "La notaria asociada al proyecto no tiene registrado los parámetros necesarios para el cálculo de los gastos de cierre, favor actualizar información en la plataforma de Operaciones.", vbExclamation, modgen_g_str_NomPlt
            'chk_Gastos.Value = False
            'Exit Sub
            If MsgBox("La notaria asociada al proyecto no tiene registrado los parámetros necesarios para el cálculo de los gastos de cierre, favor coordinar con el área legal la actualización de la información en caso contrario no se generaran los gastos de cierre." & vbCrLf & "Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               chk_Gastos.Value = False
               Exit Sub
            End If
         ElseIf r_int_Resul = 4 Then
            MsgBox "La empresa de peritaje asociada al proyecto no tiene registrado los parámetros necesarios para el cálculo de los gastos de cierre, favor actualizar información en la plataforma de Operaciones.", vbExclamation, modgen_g_str_NomPlt
            chk_Gastos.Value = False
            Exit Sub
         End If
      End If
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
     
   Call modatecli_gs_Limpia_DatCre
   
   modatecli_g_arr_DatCre(1).DatCre_TipEva = CInt(l_arr_TipEva(cmb_TipEva.ListIndex + 1).Genera_Codigo)
   modatecli_g_arr_DatCre(1).DatCre_TipMon = moddat_g_int_TipMon
   modatecli_g_arr_DatCre(1).DatCre_ComVta = CDbl(pnl_ValTot.Caption)
   modatecli_g_arr_DatCre(1).DatCre_CuoIni = CDbl(ipp_ApoPro.Text) + CDbl(pnl_FmvBbp.Caption) + CDbl(pnl_MefPbp.Caption) + CDbl(ipp_MtoAFP.Text) + CDbl(pnl_MtoBMS.Caption)
   modatecli_g_arr_DatCre(1).DatCre_ApoPro = CDbl(ipp_ApoPro.Text)
   modatecli_g_arr_DatCre(1).DatCre_FmvBbp = CDbl(pnl_FmvBbp.Caption)
   modatecli_g_arr_DatCre(1).DatCre_MefPbp = CDbl(pnl_MefPbp.Caption)
   modatecli_g_arr_DatCre(1).DatCre_MtoAFP = CDbl(ipp_MtoAFP.Text)
   modatecli_g_arr_DatCre(1).DatCre_MPSBMS = CDbl(l_dbl_MPSMS)
   modatecli_g_arr_DatCre(1).DatCre_MtoBMS = CDbl(pnl_MtoBMS.Caption)
   modatecli_g_arr_DatCre(1).DatCre_MtoPre = CDbl(pnl_TotPre.Caption)
   If InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then
      If CDbl(pnl_MtoBMS.Caption) = 0 Then
         modatecli_g_arr_DatCre(1).DatCre_BMSTas = 0
      Else
         If CDbl(l_dbl_MPSMS) <= modatecli_g_dbl_MtoFin Then
            modatecli_g_arr_DatCre(1).DatCre_BMSTas = CDbl(modatecli_g_dbl_BMSTas)
         Else
            modatecli_g_arr_DatCre(1).DatCre_BMSTas = CDbl(l_dbl_ValAfeBV)
         End If
      End If
   Else
      modatecli_g_arr_DatCre(1).DatCre_BMSTas = 0
   End If
   modatecli_g_arr_DatCre(1).DatCre_MtoInm = CDbl(ipp_ComVta.Text)
   modatecli_g_arr_DatCre(1).DatCre_MtoEst = CDbl(ipp_ValEst.Text)
   modatecli_g_arr_DatCre(1).DatCre_MtoGCi = CDbl(pnl_ValGas.Caption)
   modatecli_g_arr_DatCre(1).DatCre_PreMto = CDbl(pnl_MtoPre.Caption)
   modatecli_g_arr_DatCre(1).DatCre_TipCam = l_dbl_TipCam
   modatecli_g_arr_DatCre(1).DatCre_ComVta_Sol = CDbl(pnl_ValTot_Sol.Caption)
   modatecli_g_arr_DatCre(1).DatCre_CuoIni_Sol = CDbl(pnl_ApoPro_Sol.Caption) + CDbl(pnl_FmvBbp_Sol.Caption) + CDbl(pnl_MefPbp_Sol.Caption) + CDbl(pnl_MtoAFP_Sol.Caption) + CDbl(pnl_MtoBMS_Sol.Caption)
   modatecli_g_arr_DatCre(1).DatCre_ApoPro_Sol = CDbl(pnl_ApoPro_Sol.Caption)
   modatecli_g_arr_DatCre(1).DatCre_FmvBbp_Sol = CDbl(pnl_FmvBbp_Sol.Caption)
   modatecli_g_arr_DatCre(1).DatCre_MefPbp_Sol = CDbl(pnl_MefPbp_Sol.Caption)
   modatecli_g_arr_DatCre(1).DatCre_MtoAFP_Sol = CDbl(pnl_MtoAFP_Sol.Caption)
   modatecli_g_arr_DatCre(1).DatCre_MPSBMS_Sol = l_dbl_MPSMS
   modatecli_g_arr_DatCre(1).DatCre_MtoBMS_Sol = CDbl(pnl_MtoBMS_Sol.Caption)
   modatecli_g_arr_DatCre(1).DatCre_MtoPre_Sol = CDbl(pnl_TotPre_Sol.Caption)
   modatecli_g_arr_DatCre(1).DatCre_MtoInm_Sol = CDbl(pnl_ComVta_Sol.Caption)
   modatecli_g_arr_DatCre(1).DatCre_MtoEst_Sol = CDbl(pnl_ValEst_Sol.Caption)
   modatecli_g_arr_DatCre(1).DatCre_MtoGCi_Sol = CDbl(pnl_ValGas_Sol.Caption)
   modatecli_g_arr_DatCre(1).DatCre_PreMto_Sol = CDbl(pnl_MtoPre_Sol.Caption)
   
   modatecli_g_arr_DatCre(1).DatCre_ComVta_Dol = CDbl(pnl_ValTot_Sol.Caption) * l_dbl_TipCam
   modatecli_g_arr_DatCre(1).DatCre_CuoIni_Dol = CDbl(pnl_ApoPro_Sol.Caption) * l_dbl_TipCam + CDbl(pnl_FmvBbp_Sol.Caption) * l_dbl_TipCam + CDbl(pnl_MefPbp_Sol.Caption) * l_dbl_TipCam + CDbl(pnl_MtoAFP_Sol.Caption) * l_dbl_TipCam + CDbl(pnl_MtoBMS_Sol.Caption) * l_dbl_TipCam
   modatecli_g_arr_DatCre(1).DatCre_ApoPro_Dol = CDbl(pnl_ApoPro_Sol.Caption) * l_dbl_TipCam
   modatecli_g_arr_DatCre(1).DatCre_FmvBbp_Dol = CDbl(pnl_FmvBbp_Sol.Caption) * l_dbl_TipCam
   modatecli_g_arr_DatCre(1).DatCre_MefPbp_Dol = CDbl(pnl_MefPbp_Sol.Caption) * l_dbl_TipCam
   modatecli_g_arr_DatCre(1).DatCre_MtoAFP_Dol = CDbl(pnl_MtoAFP_Sol.Caption) * l_dbl_TipCam
   modatecli_g_arr_DatCre(1).DatCre_MPSBMS_Dol = l_dbl_MPSMS * l_dbl_TipCam
   modatecli_g_arr_DatCre(1).DatCre_MtoBMS_Dol = CDbl(pnl_MtoBMS_Sol.Caption) * l_dbl_TipCam
   modatecli_g_arr_DatCre(1).DatCre_MtoPre_Dol = CDbl(pnl_TotPre_Sol.Caption) * l_dbl_TipCam
   modatecli_g_arr_DatCre(1).DatCre_MtoInm_Dol = CDbl(pnl_ComVta_Sol.Caption) * l_dbl_TipCam
   modatecli_g_arr_DatCre(1).DatCre_MtoEst_Dol = CDbl(pnl_ValEst_Sol.Caption) * l_dbl_TipCam
   modatecli_g_arr_DatCre(1).DatCre_MtoGCi_Dol = CDbl(pnl_ValGas_Sol.Caption) * l_dbl_TipCam
   modatecli_g_arr_DatCre(1).DatCre_PreMto_Dol = CDbl(pnl_MtoPre_Sol.Caption) * l_dbl_TipCam
   
   modatecli_g_arr_DatCre(1).DatCre_PlaAno = ipp_PlaAno.Value
   modatecli_g_arr_DatCre(1).DatCre_PerGra = ipp_PerGra.Value
   modatecli_g_arr_DatCre(1).DatCre_CuoExt = cmb_CuoDbl.ItemData(cmb_CuoDbl.ListIndex)
   modatecli_g_arr_DatCre(1).DatCre_ESgDes = l_arr_EmpSeg(cmb_EmpSeg.ListIndex + 1).Genera_Codigo
   modatecli_g_arr_DatCre(1).DatCre_TipSeg = cmb_SegDes.ItemData(cmb_SegDes.ListIndex)
   modatecli_g_arr_DatCre(1).DatCre_TasEsp = cmb_TasEsp.ItemData(cmb_TasEsp.ListIndex)
   modatecli_g_arr_DatCre(1).DatCre_ESgViv = l_arr_EmpSeg(cmb_EmpSeg.ListIndex + 1).Genera_Codigo
   modatecli_g_arr_DatCre(1).DatCre_DiaPag = CInt(l_arr_DiaPag(cmb_DiaPag.ListIndex + 1).Genera_Codigo)
   
   If CInt(l_arr_TipEva(cmb_TipEva.ListIndex + 1).Genera_Codigo) = 2 Then
      modatecli_g_arr_DatCre(1).DatCre_InsFin = l_arr_InsFin(cmb_InsFin.ListIndex + 1).Genera_Codigo
      modatecli_g_arr_DatCre(1).DatCre_MonAho = cmb_MonAho.ItemData(cmb_MonAho.ListIndex)
      modatecli_g_arr_DatCre(1).DatCre_MtoAho = CDbl(ipp_MtoAho.Text)
      modatecli_g_arr_DatCre(1).DatCre_MesAho = CDbl(ipp_MesAho.Text)
   Else
      modatecli_g_arr_DatCre(1).DatCre_InsFin = ""
      modatecli_g_arr_DatCre(1).DatCre_MonAho = 0
      modatecli_g_arr_DatCre(1).DatCre_MtoAho = 0
      modatecli_g_arr_DatCre(1).DatCre_MesAho = 0
   End If
   
   modatecli_g_arr_DatCre(1).DatCre_PriViv = cmb_PriViv.ItemData(cmb_PriViv.ListIndex)
   modatecli_g_arr_DatCre(1).DatCre_ConHip = l_arr_ConHip(cmb_ConHip.ListIndex + 1).Genera_Codigo
   modatecli_g_arr_DatCre(1).DatCre_EjeSeg = l_arr_EjeSeg(cmb_EjeSeg.ListIndex + 1).Genera_Codigo
   modatecli_g_arr_DatCre(1).DatCre_Observ = txt_Observ.Text
   
   'Cargando Documentos
   ReDim modatecli_g_arr_DocCre(0)
   
   grd_Listad.Redraw = False
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      grd_Listad.Row = r_int_Contad
      
      grd_Listad.Col = 1
      If grd_Listad.Text = "X" Then
         ReDim Preserve modatecli_g_arr_DocCre(UBound(modatecli_g_arr_DocCre) + 1)
         
         grd_Listad.Col = 2
         modatecli_g_arr_DocCre(UBound(modatecli_g_arr_DocCre)).DocCre_TipDoc = CInt(grd_Listad.Text)
         grd_Listad.Col = 3
         modatecli_g_arr_DocCre(UBound(modatecli_g_arr_DocCre)).DocCre_CodGrp = grd_Listad.Text
         grd_Listad.Col = 4
         modatecli_g_arr_DocCre(UBound(modatecli_g_arr_DocCre)).DocCre_CodAct = CInt(grd_Listad.Text)
         grd_Listad.Col = 5
         modatecli_g_arr_DocCre(UBound(modatecli_g_arr_DocCre)).DocCre_CodIte = grd_Listad.Text
      End If
   Next r_int_Contad
   
   grd_Listad.Redraw = True
   modatecli_g_int_DatCreTit = 2
   Unload Me
End Sub

Private Function fs_Valida_PryAsg() As Integer
Dim r_str_Parame  As String
Dim r_rst_Genera  As ADODB.Recordset
   
   fs_Valida_PryAsg = 0
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "      SELECT NVL(X.DATGEN_PRYAPR,0) AS PRYAPR "
   r_str_Parame = r_str_Parame & "        FROM PRY_DATGEN X "
   r_str_Parame = r_str_Parame & "       WHERE X.DATGEN_CODIGO =  '" & modatecli_g_arr_DatInm(1).DatInm_CodPry & "' "

   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
      Exit Function
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst

      If r_rst_Genera!PRYAPR = 1 Then
         fs_Valida_PryAsg = 1
      End If
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Function

'Private Function fs_Valida_NotPer() As Integer
'Dim r_str_Princi  As String
'Dim r_rst_Genera  As ADODB.Recordset
'
'   fs_Valida_NotPer = 0
'
'   'GASPAR_TIPTAB: 1=PERITO, 2=NOTARIA
'
'   r_str_Princi = ""
'   r_str_Princi = r_str_Princi & " SELECT "
'   r_str_Princi = r_str_Princi & "       ( SELECT COUNT(*) "
'   r_str_Princi = r_str_Princi & "           FROM TRA_GASPAR A "
'   r_str_Princi = r_str_Princi & "                INNER JOIN PRY_DATGEN ON TRIM(TO_CHAR(TRIM(DATGEN_CODPRT), '000000')) = TRIM(GASPAR_CODEMP) AND GASPAR_CODPRY = DATGEN_CODIGO "
'   r_str_Princi = r_str_Princi & "          WHERE GASPAR_TIPTAB = 1 AND GASPAR_CODPRD = '" & moddat_g_str_CodPrd & "' AND GASPAR_CODPRY = '" & modatecli_g_arr_DatInm(1).DatInm_CodPry & "' AND GASPAR_SITUAC = 1 ) AS PERITO, "
'
'   r_str_Princi = r_str_Princi & "       ( SELECT COUNT(*) "
'   r_str_Princi = r_str_Princi & "           FROM TRA_GASPAR A "
'   r_str_Princi = r_str_Princi & "                INNER JOIN PRY_DATGEN ON TRIM(TO_CHAR(TRIM(DATGEN_CODPRT), '000000')) = TRIM(GASPAR_CODEMP) AND GASPAR_CODPRY = DATGEN_CODIGO "
'   r_str_Princi = r_str_Princi & "          WHERE GASPAR_TIPTAB = 1 AND GASPAR_CODPRD = '" & moddat_g_str_CodPrd & "' AND GASPAR_CODPRY = '" & modatecli_g_arr_DatInm(1).DatInm_CodPry & "' AND GASPAR_SITUAC = 1 "
'   r_str_Princi = r_str_Princi & "            AND GASPAR_GASTAS_MTO > 0 ) AS MONTO_PERITO, "
'
'   r_str_Princi = r_str_Princi & "       ( SELECT COUNT(*) "
'   r_str_Princi = r_str_Princi & "           FROM TRA_GASPAR A "
'   r_str_Princi = r_str_Princi & "                INNER JOIN PRY_DATGEN ON TRIM(TO_CHAR(TRIM(DATGEN_CODNOT), '000000')) = TRIM(GASPAR_CODEMP) AND GASPAR_CODPRY = DATGEN_CODIGO "
'   r_str_Princi = r_str_Princi & "          WHERE GASPAR_TIPTAB = 2 AND GASPAR_CODPRD = '" & moddat_g_str_CodPrd & "' AND GASPAR_CODPRY = '" & modatecli_g_arr_DatInm(1).DatInm_CodPry & "' AND GASPAR_SITUAC = 1 ) AS NOTARIA, "
'
'   r_str_Princi = r_str_Princi & "       ( SELECT COUNT(*) "
'   r_str_Princi = r_str_Princi & "           FROM TRA_GASPAR A "
'   r_str_Princi = r_str_Princi & "                INNER JOIN PRY_DATGEN ON TRIM(TO_CHAR(TRIM(DATGEN_CODNOT), '000000')) = TRIM(GASPAR_CODEMP) AND GASPAR_CODPRY = DATGEN_CODIGO "
'   r_str_Princi = r_str_Princi & "          WHERE GASPAR_TIPTAB = 2 AND GASPAR_CODPRD = '" & moddat_g_str_CodPrd & "' AND GASPAR_CODPRY = '" & modatecli_g_arr_DatInm(1).DatInm_CodPry & "' AND GASPAR_SITUAC = 1 "
'   r_str_Princi = r_str_Princi & "            AND GASPAR_GASNOT_MTO > 0 "
'   r_str_Princi = r_str_Princi & "            AND GASPAR_REGMIN_TASINM > 0 "
'   r_str_Princi = r_str_Princi & "            AND GASPAR_REGMIN_FACINM > 0 "
'   r_str_Princi = r_str_Princi & "            AND GASPAR_REGMIN_FICINM > 0 "
'   r_str_Princi = r_str_Princi & "            AND GASPAR_REGMIN_TASEST > 0 "
'   r_str_Princi = r_str_Princi & "            AND GASPAR_REGMIN_FACEST > 0 "
'   r_str_Princi = r_str_Princi & "            AND GASPAR_REGMIN_FICEST > 0 "
'   r_str_Princi = r_str_Princi & "            AND GASPAR_REGGAR_TAS000 > 0 "
'   r_str_Princi = r_str_Princi & "            AND GASPAR_REGGAR_TAS001 > 0 "
'   r_str_Princi = r_str_Princi & "            AND GASPAR_REGGAR_TAS002 > 0 ) AS TASAS_NOTARIA "
'
'   r_str_Princi = r_str_Princi & "   FROM DUAL "
'
'
'   If Not gf_EjecutaSQL(r_str_Princi, r_rst_Genera, 3) Then
'      Exit Function
'   End If
'
'   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
'      r_rst_Genera.MoveFirst
'
'      If CInt(r_rst_Genera!PERITO) = 0 Then
'         fs_Valida_NotPer = 1
'      ElseIf CInt(r_rst_Genera!PERITO) > 0 And r_rst_Genera!MONTO_PERITO = 0 Then
'         fs_Valida_NotPer = 2
'      ElseIf CInt(r_rst_Genera!NOTARIA) = 0 Then
'         fs_Valida_NotPer = 3
'      ElseIf CInt(r_rst_Genera!NOTARIA) > 0 And r_rst_Genera!TASAS_NOTARIA = 0 Then
'         fs_Valida_NotPer = 4
'      End If
'   End If
'
'   r_rst_Genera.Close
'   Set r_rst_Genera = Nothing
'End Function

Private Sub cmd_SimCre_Click()
   If moddat_gf_Obtiene_TipCam(1, 2) = 0 Then
      MsgBox "No se encuentra registrado el Tipo de Cambio. No puede ingresar a esta opción.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   frm_SimCre_11.Show 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

'**************************************************************************************************
'* FORM
'**************************************************************************************************
Private Sub Form_Load()
Dim r_int_Contad     As Integer
Dim r_int_ConAux     As Integer
Dim r_int_TipDoc     As Integer
Dim r_str_CodGrp     As String
Dim r_int_CodAct     As Integer
Dim r_str_CodIte     As String

   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   pnl_Produc.Caption = moddat_gf_Consulta_Produc(moddat_g_str_CodPrd)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   l_dbl_TipCam = moddat_gf_Obtiene_TipCam(1, 2)
   
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Carga_Docume
   
   ipp_ComVta.Enabled = True
   ipp_ApoPro.Enabled = True
   ipp_MtoAFP.Enabled = True
   pnl_MtoPre.Enabled = False
   ipp_PlaAno.Enabled = True
   ipp_PerGra.Enabled = True
   cmb_EmpSeg.Enabled = True
   cmb_SegDes.Enabled = True
   cmb_DiaPag.Enabled = True
   
   If modatecli_g_arr_DatCre(1).DatCre_ComVta > 0 Then
      cmb_TipEva.ListIndex = gf_Busca_Arregl(l_arr_TipEva, Format(modatecli_g_arr_DatCre(1).DatCre_TipEva, "000")) - 1
      pnl_ValTot.Caption = Format(modatecli_g_arr_DatCre(1).DatCre_ComVta, "##,###,##0.00") & " "
      ipp_ComVta.Value = modatecli_g_arr_DatCre(1).DatCre_MtoInm
      ipp_ValEst.Value = modatecli_g_arr_DatCre(1).DatCre_MtoEst
      ipp_ApoPro.Value = modatecli_g_arr_DatCre(1).DatCre_ApoPro
      pnl_FmvBbp.Caption = Format(modatecli_g_arr_DatCre(1).DatCre_FmvBbp, "##,###,##0.00") & " "
      pnl_MefPbp.Caption = Format(modatecli_g_arr_DatCre(1).DatCre_MefPbp, "##,###,##0.00") & " "
      ipp_MtoAFP.Text = modatecli_g_arr_DatCre(1).DatCre_MtoAFP
      l_dbl_MPSMS = Format(modatecli_g_arr_DatCre(1).DatCre_MPSBMS, "##,###,##0.00") & " "
      pnl_MtoBMS.Caption = Format(modatecli_g_arr_DatCre(1).DatCre_MtoBMS, "##,###,##0.00") & " "
      pnl_MtoPre.Caption = Format(modatecli_g_arr_DatCre(1).DatCre_PreMto, "##,###,##0.00") & " "
      pnl_ValGas.Caption = Format(modatecli_g_arr_DatCre(1).DatCre_MtoGCi, "##,###,##0.00") & " "
      If CDbl(pnl_ValGas.Caption) > 0 Then
         chk_Gastos.Value = True
      Else
         chk_Gastos.Value = False
      End If
      pnl_TotPre.Caption = Format(modatecli_g_arr_DatCre(1).DatCre_MtoPre, "##,###,##0.00") & " "
      Call fs_Calcul
      
      If modatecli_g_arr_DatCre(1).DatCre_PlaAno = 0 Then
         'Plazo de Crédito
         If moddat_gf_Consulta_SubPrd_Arregl(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub) Then
            ipp_PlaAno.MinValue = moddat_g_arr_Genera(1).Genera_PlzMin
            ipp_PlaAno.MaxValue = moddat_g_arr_Genera(1).Genera_PlzMax
         End If
         
         'Periodo de Gracia
         l_int_GraMax = 0
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "008", "002") Then
            ipp_PerGra.MinValue = moddat_g_arr_Genera(1).Genera_ValMin
            ipp_PerGra.MaxValue = moddat_g_arr_Genera(1).Genera_ValMax
            l_int_GraMax = moddat_g_arr_Genera(1).Genera_ValMax
         End If
      Else
         ipp_PlaAno.Value = modatecli_g_arr_DatCre(1).DatCre_PlaAno
         ipp_PerGra.Value = modatecli_g_arr_DatCre(1).DatCre_PerGra
      End If
         
      Call gs_BuscarCombo_Item(cmb_CuoDbl, modatecli_g_arr_DatCre(1).DatCre_CuoExt)
      cmb_EmpSeg.ListIndex = gf_Busca_Arregl(l_arr_EmpSeg, modatecli_g_arr_DatCre(1).DatCre_ESgDes) - 1
      Call moddat_gs_Carga_TipSeg(cmb_SegDes, l_arr_EmpSeg(cmb_EmpSeg.ListIndex + 1).Genera_Codigo)
      Call gs_BuscarCombo_Item(cmb_SegDes, modatecli_g_arr_DatCre(1).DatCre_TipSeg)
      Call gs_BuscarCombo_Item(cmb_TasEsp, modatecli_g_arr_DatCre(1).DatCre_TasEsp)
      cmb_DiaPag.ListIndex = gf_Busca_Arregl(l_arr_DiaPag, Format(modatecli_g_arr_DatCre(1).DatCre_DiaPag, "000")) - 1
      
      If cmb_TipEva.ListIndex <> -1 Then
         If CInt(l_arr_TipEva(cmb_TipEva.ListIndex + 1).Genera_Codigo) = 2 Then
            cmb_InsFin.ListIndex = gf_Busca_Arregl(l_arr_InsFin, modatecli_g_arr_DatCre(1).DatCre_InsFin) - 1
            Call gs_BuscarCombo_Item(cmb_MonAho, modatecli_g_arr_DatCre(1).DatCre_MonAho)
            ipp_MtoAho.Text = modatecli_g_arr_DatCre(1).DatCre_MtoAho
            ipp_MesAho.Text = modatecli_g_arr_DatCre(1).DatCre_MesAho
            
            cmb_InsFin.Enabled = True
            cmb_MonAho.Enabled = True
            ipp_MtoAho.Enabled = True
            ipp_MesAho.Enabled = True
         End If
      End If
      Call gs_BuscarCombo_Item(cmb_PriViv, modatecli_g_arr_DatCre(1).DatCre_PriViv)
      cmb_ConHip.ListIndex = gf_Busca_Arregl(l_arr_ConHip, modatecli_g_arr_DatCre(1).DatCre_ConHip) - 1
      cmb_EjeSeg.ListIndex = gf_Busca_Arregl(l_arr_EjeSeg, modatecli_g_arr_DatCre(1).DatCre_EjeSeg) - 1
      txt_Observ.Text = modatecli_g_arr_DatCre(1).DatCre_Observ
   
      grd_Listad.Redraw = False
      For r_int_Contad = 0 To grd_Listad.Rows - 1
         grd_Listad.Row = r_int_Contad
         grd_Listad.Col = 2: r_int_TipDoc = CInt(grd_Listad.Text)
         grd_Listad.Col = 3: r_str_CodGrp = grd_Listad.Text
         grd_Listad.Col = 4: r_int_CodAct = CInt(grd_Listad.Text)
         grd_Listad.Col = 5: r_str_CodIte = grd_Listad.Text
         
         For r_int_ConAux = 0 To UBound(modatecli_g_arr_DocCre)
            If modatecli_g_arr_DocCre(r_int_ConAux).DocCre_TipDoc = r_int_TipDoc And modatecli_g_arr_DocCre(r_int_ConAux).DocCre_CodGrp = r_str_CodGrp And modatecli_g_arr_DocCre(r_int_ConAux).DocCre_CodAct = r_int_CodAct And modatecli_g_arr_DocCre(r_int_ConAux).DocCre_CodIte = r_str_CodIte Then
               grd_Listad.Col = 1
               grd_Listad.Text = "X"
               Exit For
            End If
         Next r_int_ConAux
      Next r_int_Contad
      
      grd_Listad.Redraw = True
      Call gs_UbiIniGrid(grd_Listad)
   End If
   
   'ipp_ValEst.Enabled = False
   'If modatecli_g_arr_DatInm(1).DatInm_FlgEst = 1 Then
   '   ipp_ValEst.Enabled = True
   'End If
   
   'Obteniendo Información del Bomo MiVivienda Sostenible
   Call moddat_gs_Consulta_DatBMS(modatecli_g_arr_DatInm(1).DatInm_CodPry, l_int_FlgAfeBV, l_int_FlgTipAfe, l_dbl_ValAfeBV)
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

'**************************************************************************************************
'* PROCEDIMIENTOS
'**************************************************************************************************
Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_PriViv, 1, "214")
   Call moddat_gs_Carga_ParSubPrd(cmb_TipEva, l_arr_TipEva(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "014")
   
   'Plazo de Crédito
   If moddat_gf_Consulta_SubPrd_Arregl(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub) Then
      ipp_PlaAno.MinValue = moddat_g_arr_Genera(1).Genera_PlzMin
      ipp_PlaAno.MaxValue = moddat_g_arr_Genera(1).Genera_PlzMax
   End If
   
   'Periodo de Gracia
   l_int_GraMax = 0
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "008", "002") Then
      ipp_PerGra.MinValue = moddat_g_arr_Genera(1).Genera_ValMin
      ipp_PerGra.MaxValue = moddat_g_arr_Genera(1).Genera_ValMax
      l_int_GraMax = moddat_g_arr_Genera(1).Genera_ValMax
   End If
   
   Call moddat_gs_Carga_ParSubPrd(cmb_DiaPag, l_arr_DiaPag(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "009")
   Call moddat_gs_Carga_EmpSeg(cmb_EmpSeg, l_arr_EmpSeg, 1)
   Call moddat_gs_Carga_EjecMC(cmb_ConHip, l_arr_ConHip, 121)
   Call moddat_gs_Carga_EjecMC(cmb_EjeSeg, l_arr_EjeSeg, 131)
   Call moddat_gs_Carga_LisIte_Combo(cmb_CuoDbl, 1, "277")
   
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 9690
   grd_Listad.ColWidth(1) = 1140
   grd_Listad.ColWidth(2) = 0
   grd_Listad.ColWidth(3) = 0
   grd_Listad.ColWidth(4) = 0
   grd_Listad.ColWidth(5) = 0
   grd_Listad.ColWidth(6) = 0
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   
   'Ahorro Programado
   Call moddat_gs_Carga_LisIte(cmb_InsFin, l_arr_InsFin, 1, "505")
   Call moddat_gs_Carga_LisIte_Combo(cmb_MonAho, 1, "204")
   
   'Bono - Producto FMV MAS BBP y BBP COMPLEMENTO INICIAL
   pnl_FmvBbp.Caption = "0.00 "
   pnl_MefPbp.Caption = "0.00 "
   
   If InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 And moddat_g_str_CodPrd <> "019" Then
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "012") Then
         pnl_FmvBbp.Caption = Format(l_arr_ParPrd(1).Genera_Cantid, "###,###,##0.00") & " "
      End If
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "013") Then
         pnl_MefPbp.Caption = Format(l_arr_ParPrd(1).Genera_Cantid, "###,###,##0.00") & " "
      End If
   End If
   
   'Tasa Especial
   Call moddat_gs_Carga_LisIte_Combo(cmb_TasEsp, 1, "522")
End Sub

Private Sub fs_Limpia()
   ipp_ComVta.Value = 0
   ipp_ValEst.Value = 0
   ipp_ApoPro.Value = 0
   ipp_MtoAFP.Value = 0
   ipp_MtoAho.Value = 0
   
   pnl_ValTot.Caption = "0.00 "
   pnl_CuoIni.Caption = "0.00 "
   pnl_MtoBMS.Caption = "0.00 "
   pnl_MtoPre.Caption = "0.00 "
   pnl_ValGas.Caption = "0.00 "
   pnl_TotPre.Caption = "0.00 "
   
   pnl_ValTot_Sol.Caption = "0.00 "
   pnl_ComVta_Sol.Caption = "0.00 "
   pnl_CuoIni_Sol.Caption = "0.00 "
   pnl_ApoPro_Sol.Caption = "0.00 "
   pnl_FmvBbp_Sol.Caption = "0.00 "
   pnl_MefPbp_Sol.Caption = "0.00 "
   pnl_MtoAFP_Sol.Caption = "0.00 "
   pnl_MtoBMS_Sol.Caption = "0.00 "
   pnl_MtoPre_Sol.Caption = "0.00 "
   pnl_ValGas_Sol.Caption = "0.00 "
   pnl_TotPre_Sol.Caption = "0.00 "
   
   ipp_PlaAno.Value = ipp_PlaAno.MinValue
   ipp_PerGra.Value = 0
   cmb_CuoDbl.ListIndex = -1
   cmb_EmpSeg.ListIndex = -1
   cmb_SegDes.Clear
   cmb_DiaPag.ListIndex = -1
   txt_Observ.Text = ""
   cmb_ConHip.ListIndex = -1
   cmb_EjeSeg.ListIndex = -1
   
   ipp_ComVta.Enabled = False
   ipp_ApoPro.Enabled = False
   pnl_MtoPre.Enabled = False
   ipp_PlaAno.Enabled = False
   ipp_PerGra.Enabled = False
   cmb_EmpSeg.Enabled = False
   cmb_SegDes.Enabled = False
   cmb_DiaPag.Enabled = False
   Call gs_LimpiaGrid(grd_Listad)
   
   cmb_InsFin.ListIndex = -1
   cmb_MonAho.ListIndex = -1
   ipp_MtoAho.Value = 0
   ipp_MesAho.Value = 0
   cmb_InsFin.Enabled = False
   cmb_MonAho.Enabled = False
   ipp_MtoAho.Enabled = False
   ipp_MesAho.Enabled = False
End Sub

Private Sub fs_Carga_Docume()
Dim r_int_ActPri_Cli    As Integer
Dim r_int_ActSec_Cli    As Integer
Dim r_int_ActPri_Cyg    As Integer
Dim r_int_ActSec_Cyg    As Integer
   
   '0 - Descripción
   '1 - Selección
   '2 - Tipo de Origen de Documento
   '3 - Código de Grupo
   '4 - Código de Actividad Económica
   '5 - Código de Item
   Call gs_LimpiaGrid(grd_Listad)
   
   'Documentos Crediticios
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARPRD "
   g_str_Parame = g_str_Parame & " WHERE PARPRD_CODPRD = '" & moddat_g_str_CodPrd & "' "
   g_str_Parame = g_str_Parame & "   AND PARPRD_CODSUB = '" & moddat_g_str_CodSub & "' "
   g_str_Parame = g_str_Parame & "   AND PARPRD_CODGRP = '011' "
   g_str_Parame = g_str_Parame & "   AND PARPRD_CODITE <> '000' "
   g_str_Parame = g_str_Parame & "   AND PARPRD_SITUAC = 1 "
   g_str_Parame = g_str_Parame & " ORDER BY PARPRD_CODITE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      grd_Listad.Redraw = False
      g_rst_Genera.MoveFirst
      Do While Not g_rst_Genera.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0: grd_Listad.Text = Trim(g_rst_Genera!PARPRD_DESCRI)
         grd_Listad.Col = 1: grd_Listad.Text = ""
         grd_Listad.Col = 2: grd_Listad.Text = "1"
         grd_Listad.Col = 3: grd_Listad.Text = "011"
         grd_Listad.Col = 4: grd_Listad.Text = "0"
         grd_Listad.Col = 5: grd_Listad.Text = g_rst_Genera!PARPRD_CODITE
         g_rst_Genera.MoveNext
      Loop
      grd_Listad.Redraw = True
   End If
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   r_int_ActPri_Cli = moddat_gf_Consulta_ActEco(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 1)
   r_int_ActSec_Cli = moddat_gf_Consulta_ActEco(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 2)
   r_int_ActPri_Cyg = moddat_gf_Consulta_ActEco(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 1)
   r_int_ActSec_Cyg = moddat_gf_Consulta_ActEco(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 2)
   
   'Documentos por Actividad Económica Titular - Actividad Principal
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARACT "
   g_str_Parame = g_str_Parame & " WHERE PARACT_CODPRD = '" & moddat_g_str_CodPrd & "' "
   g_str_Parame = g_str_Parame & "   AND PARACT_CODSUB = '" & moddat_g_str_CodSub & "' "
   g_str_Parame = g_str_Parame & "   AND PARACT_CODACT = " & CStr(r_int_ActPri_Cli) & " "
   g_str_Parame = g_str_Parame & "   AND PARACT_CODGRP = '002' "
   g_str_Parame = g_str_Parame & "   AND PARACT_CODITE <> '000' "
   g_str_Parame = g_str_Parame & "   AND PARACT_SITUAC = 1 "
   g_str_Parame = g_str_Parame & " ORDER BY PARACT_CODITE ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      grd_Listad.Redraw = False
      g_rst_Genera.MoveFirst
      Do While Not g_rst_Genera.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0:  grd_Listad.Text = Trim(g_rst_Genera!PARACT_DESCRI)
         grd_Listad.Col = 1:  grd_Listad.Text = ""
         grd_Listad.Col = 2:  grd_Listad.Text = "2"
         grd_Listad.Col = 3:  grd_Listad.Text = "002"
         grd_Listad.Col = 4:  grd_Listad.Text = r_int_ActPri_Cli
         grd_Listad.Col = 5:  grd_Listad.Text = g_rst_Genera!PARACT_CODITE
         g_rst_Genera.MoveNext
      Loop
      grd_Listad.Redraw = True
   End If
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   'Documentos por Actividad Económica Titular - Actividad Secundaria
   If r_int_ActPri_Cli <> r_int_ActSec_Cli And r_int_ActSec_Cli > 0 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARACT "
      g_str_Parame = g_str_Parame & " WHERE PARACT_CODPRD = '" & moddat_g_str_CodPrd & "' "
      g_str_Parame = g_str_Parame & "   AND PARACT_CODSUB = '" & moddat_g_str_CodSub & "' "
      g_str_Parame = g_str_Parame & "   AND PARACT_CODACT = " & CStr(r_int_ActSec_Cli) & " "
      g_str_Parame = g_str_Parame & "   AND PARACT_CODGRP = '002' "
      g_str_Parame = g_str_Parame & "   AND PARACT_CODITE <> '000' "
      g_str_Parame = g_str_Parame & "   AND PARACT_SITUAC = 1 "
      g_str_Parame = g_str_Parame & " ORDER BY PARACT_CODITE ASC "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         grd_Listad.Redraw = False
         g_rst_Genera.MoveFirst
         Do While Not g_rst_Genera.EOF
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            
            grd_Listad.Col = 0:  grd_Listad.Text = Trim(g_rst_Genera!PARACT_DESCRI)
            grd_Listad.Col = 1:  grd_Listad.Text = ""
            grd_Listad.Col = 2:  grd_Listad.Text = "2"
            grd_Listad.Col = 3:  grd_Listad.Text = "002"
            grd_Listad.Col = 4:  grd_Listad.Text = r_int_ActSec_Cli
            grd_Listad.Col = 5:  grd_Listad.Text = g_rst_Genera!PARACT_CODITE
            g_rst_Genera.MoveNext
         Loop
         grd_Listad.Redraw = True
      End If
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
   
   'Documentos por Actividad Económica Cónyuge - Actividad Principal
   If r_int_ActPri_Cyg > 0 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARACT "
      g_str_Parame = g_str_Parame & " WHERE PARACT_CODPRD = '" & moddat_g_str_CodPrd & "' "
      g_str_Parame = g_str_Parame & "   AND PARACT_CODSUB = '" & moddat_g_str_CodSub & "' "
      g_str_Parame = g_str_Parame & "   AND PARACT_CODACT = " & CStr(r_int_ActPri_Cyg) & " "
      g_str_Parame = g_str_Parame & "   AND PARACT_CODGRP = '003' "
      g_str_Parame = g_str_Parame & "   AND PARACT_CODITE <> '000' "
      g_str_Parame = g_str_Parame & "   AND PARACT_SITUAC = 1 "
      g_str_Parame = g_str_Parame & " ORDER BY PARACT_CODITE ASC "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         grd_Listad.Redraw = False
         g_rst_Genera.MoveFirst
         Do While Not g_rst_Genera.EOF
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            
            grd_Listad.Col = 0:  grd_Listad.Text = Trim(g_rst_Genera!PARACT_DESCRI)
            grd_Listad.Col = 1:  grd_Listad.Text = ""
            grd_Listad.Col = 2:  grd_Listad.Text = "2"
            grd_Listad.Col = 3:  grd_Listad.Text = "003"
            grd_Listad.Col = 4:  grd_Listad.Text = r_int_ActPri_Cyg
            grd_Listad.Col = 5:  grd_Listad.Text = g_rst_Genera!PARACT_CODITE
            g_rst_Genera.MoveNext
         Loop
         grd_Listad.Redraw = True
      End If
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
   
   'Documentos por Actividad Económica Cónyuge - Actividad Secundaria
   If r_int_ActPri_Cyg <> r_int_ActSec_Cyg And r_int_ActPri_Cyg > 0 And r_int_ActSec_Cyg > 0 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARACT "
      g_str_Parame = g_str_Parame & " WHERE PARACT_CODPRD = '" & moddat_g_str_CodPrd & "' "
      g_str_Parame = g_str_Parame & "   AND PARACT_CODSUB = '" & moddat_g_str_CodSub & "' "
      g_str_Parame = g_str_Parame & "   AND PARACT_CODACT = " & CStr(r_int_ActSec_Cyg) & " "
      g_str_Parame = g_str_Parame & "   AND PARACT_CODGRP = '003' "
      g_str_Parame = g_str_Parame & "   AND PARACT_CODITE <> '000' "
      g_str_Parame = g_str_Parame & "   AND PARACT_SITUAC = 1 "
      g_str_Parame = g_str_Parame & "ORDER BY PARACT_CODITE ASC "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         grd_Listad.Redraw = False
         g_rst_Genera.MoveFirst
         Do While Not g_rst_Genera.EOF
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            
            grd_Listad.Col = 0:  grd_Listad.Text = Trim(g_rst_Genera!PARACT_DESCRI)
            grd_Listad.Col = 1:  grd_Listad.Text = ""
            grd_Listad.Col = 2:  grd_Listad.Text = "2"
            grd_Listad.Col = 3:  grd_Listad.Text = "003"
            grd_Listad.Col = 4:  grd_Listad.Text = r_int_ActSec_Cyg
            grd_Listad.Col = 5:  grd_Listad.Text = g_rst_Genera!PARACT_CODITE
            g_rst_Genera.MoveNext
         Loop
         grd_Listad.Redraw = True
      End If
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
End Sub

Private Sub fs_Bono_Verde(ByVal p_ValAfe As Double)
   If InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then
      If l_int_FlgAfeBV = 1 Then
         pnl_MtoBMS.Caption = Format((CDbl(l_dbl_MPSMS) * p_ValAfe) / (1 + p_ValAfe), "###,###,##0.00") & " "  'pnl_MPSBMS.Caption
         If pnl_MtoBMS.Caption < 0 Then pnl_MtoBMS.Caption = "0.00 "
      End If
      pnl_MtoPre.Caption = Format(CDbl(l_dbl_MPSMS) - CDbl(pnl_MtoBMS.Caption), "###,###,##0.00") & " "        'pnl_MPSBMS.Caption
      If pnl_MtoPre.Caption < 0 Then pnl_MtoPre.Caption = "0.00 "
   Else
      pnl_MtoBMS.Caption = "0.00 "
   End If
End Sub

Private Sub fs_Calcular_MtoSBMS()
   l_dbl_MPSMS = Format(CDbl(pnl_ValTot.Caption) - CDbl(ipp_ApoPro.Text) - (CDbl(pnl_FmvBbp.Caption) + CDbl(pnl_MefPbp.Caption)) - CDbl(ipp_MtoAFP.Text), "###,###,##0.00") & " "
   If l_dbl_MPSMS < 0 Then l_dbl_MPSMS = 0
End Sub

Private Sub fs_Calcul()
   If CDbl(pnl_ValTot.Caption) > 0 Then
      If moddat_g_str_CodPrd = "023" Then
         pnl_PorIni.Caption = "(" & Format((CDbl(ipp_ApoPro.Text) + CDbl(ipp_MtoAFP) + CDbl(pnl_FmvBbp.Caption) + CDbl(pnl_MefPbp.Caption) + CDbl(pnl_MtoBMS.Caption)) / CDbl(pnl_ValTot.Caption) * 100, "##0.0000") & "%) "
      Else
         pnl_PorIni.Caption = "(" & Format((CDbl(ipp_ApoPro.Text) + CDbl(ipp_MtoAFP)) / CDbl(pnl_ValTot.Caption) * 100, "##0.0000") & "%) "
      End If
   Else
      pnl_PorIni.Caption = "(0.00%) "
   End If
   
   If moddat_g_int_TipMon = 1 Then
      pnl_ValTot_Sol.Caption = Format(CDbl(pnl_ValTot.Caption), "###,###,##0.00") & " "
      pnl_ComVta_Sol.Caption = Format(CDbl(ipp_ComVta.Text), "###,###,##0.00") & " "
      pnl_ValEst_Sol.Caption = Format(CDbl(ipp_ValEst.Text), "###,###,##0.00") & " "
      pnl_CuoIni_Sol.Caption = Format(CDbl(pnl_CuoIni.Caption), "###,###,##0.00") & " "
      pnl_ApoPro_Sol.Caption = Format(CDbl(ipp_ApoPro.Text), "###,###,##0.00") & " "
      pnl_FmvBbp_Sol.Caption = Format(CDbl(pnl_FmvBbp.Caption), "###,###,##0.00") & " "
      pnl_MefPbp_Sol.Caption = Format(CDbl(pnl_MefPbp.Caption), "###,###,##0.00") & " "
      pnl_MtoAFP_Sol.Caption = Format(CDbl(ipp_MtoAFP.Text), "###,###,##0.00") & " "
      pnl_MtoBMS_Sol.Caption = Format(CDbl(pnl_MtoBMS.Caption), "###,###,##0.00") & " "
      pnl_MtoPre_Sol.Caption = Format(CDbl(pnl_MtoPre.Caption), "###,###,##0.00") & " "
      pnl_ValGas_Sol.Caption = Format(CDbl(pnl_ValGas.Caption), "###,###,##0.00") & " "
      pnl_TotPre_Sol.Caption = Format(CDbl(pnl_TotPre.Caption), "###,###,##0.00") & " "
   Else
      pnl_ValTot_Sol.Caption = Format(CDbl(pnl_ValTot.Caption) * l_dbl_TipCam, "###,###,##0.00") & " "
      pnl_ComVta_Sol.Caption = Format(CDbl(ipp_ComVta.Text) * l_dbl_TipCam, "###,###,##0.00") & " "
      pnl_ValEst_Sol.Caption = Format(CDbl(ipp_ValEst.Text) * l_dbl_TipCam, "###,###,##0.00") & " "
      pnl_ApoPro_Sol.Caption = Format((CDbl(ipp_ApoPro.Text)) * l_dbl_TipCam, "###,###,##0.00") & " "
      pnl_FmvBbp_Sol.Caption = Format((CDbl(pnl_FmvBbp.Caption)) * l_dbl_TipCam, "###,###,##0.00") & " "
      pnl_MefPbp_Sol.Caption = Format((CDbl(pnl_MefPbp.Caption)) * l_dbl_TipCam, "###,###,##0.00") & " "
      pnl_MtoAFP_Sol.Caption = Format((CDbl(ipp_MtoAFP.Text)) * l_dbl_TipCam, "###,###,##0.00") & " "
      pnl_MtoBMS_Sol.Caption = Format((CDbl(pnl_MtoBMS.Caption)) * l_dbl_TipCam, "###,###,##0.00") & " "
      pnl_MtoPre_Sol.Caption = Format(CDbl(pnl_MtoPre.Caption) * l_dbl_TipCam, "###,###,##0.00") & " "
      pnl_ValGas_Sol.Caption = Format(CDbl(pnl_ValGas.Caption) * l_dbl_TipCam, "###,###,##0.00") & " "
      pnl_TotPre_Sol.Caption = Format(CDbl(pnl_TotPre.Caption) * l_dbl_TipCam, "###,###,##0.00") & " "
   End If
End Sub

'**************************************************************************************************
'* CONTROLES
'**************************************************************************************************
Private Sub cmb_TipEva_Click()
   If cmb_TipEva.ListIndex > -1 Then
      If CInt(l_arr_TipEva(cmb_TipEva.ListIndex + 1).Genera_Codigo) = 2 Then
         cmb_InsFin.Enabled = True
         cmb_MonAho.Enabled = True
         ipp_MtoAho.Enabled = True
         ipp_MesAho.Enabled = True
      Else
         cmb_InsFin.Enabled = False
         cmb_MonAho.Enabled = False
         ipp_MtoAho.Enabled = False
         ipp_MesAho.Enabled = False
         cmb_InsFin.ListIndex = -1
         cmb_MonAho.ListIndex = -1
         ipp_MtoAho.Value = 0
         ipp_MesAho.Value = 0
      End If
   End If
   Call gs_SetFocus(cmb_PriViv)
End Sub

Private Sub cmb_TipEva_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipEva_Click
   End If
End Sub

Private Sub cmb_PriViv_Click()
   If cmb_PriViv.ListIndex > -1 Then
      Call gs_SetFocus(ipp_PlaAno)
   End If
End Sub

Private Sub cmb_PriViv_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_PriViv_Click
   End If
End Sub

Private Sub ipp_ApoPro_LostFocus()
   Call fs_Calcul
   Call fs_Calcular_Prestamo
End Sub

Private Sub ipp_ComVta_Change()
   Call ipp_ApoPro_Change
   Call fs_Calcul
   Call fs_Calcular_Prestamo
End Sub

Private Sub ipp_ComVta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValEst)
      ipp_ComVta_Change
   End If
End Sub

Private Sub ipp_ApoPro_Change()
   Call fs_Calcular_MtoSBMS
   
   If CDbl(pnl_ValTot.Caption) > 0 Then
      If moddat_g_str_CodPrd = "023" Then
         pnl_PorIni.Caption = "(" & Format((CDbl(ipp_ApoPro.Text) + CDbl(ipp_MtoAFP) + CDbl(pnl_FmvBbp.Caption) + CDbl(pnl_MefPbp.Caption) + CDbl(pnl_MtoBMS.Caption)) / CDbl(pnl_ValTot.Caption) * 100, "##0.0000") & "%) " 'ipp_ComVta.Text
      Else
         pnl_PorIni.Caption = "(" & Format((CDbl(ipp_ApoPro.Text) + CDbl(ipp_MtoAFP)) / CDbl(pnl_ValTot.Caption) * 100, "##0.0000") & "%) " 'ipp_ComVta.Text
      End If
   Else
      pnl_PorIni.Caption = "(0.00%) "
   End If
   
   pnl_MtoPre.Caption = Format(CDbl(pnl_ValTot.Caption) - CDbl(ipp_ApoPro.Text) - (CDbl(pnl_FmvBbp.Caption) + CDbl(pnl_MefPbp.Caption)) - CDbl(ipp_MtoAFP.Text) - CDbl(pnl_MtoBMS.Caption), "###,###,##0.00") & " " 'ipp_ComVta.Text
   If pnl_MtoPre.Caption < 0 Then pnl_MtoPre.Caption = "0.00 "

   If CDbl(pnl_ValTot.Caption) <= modatecli_g_dbl_MtoFin Then 'ipp_ComVta.Text
      If l_int_FlgAfeBV = 1 And l_int_FlgTipAfe <> 0 Or l_dbl_ValAfeBV <> 0 Then Call fs_Bono_Verde(modatecli_g_dbl_BMSTas)
   Else
      If l_int_FlgAfeBV = 0 Or l_int_FlgTipAfe = 0 Or l_dbl_ValAfeBV = 0 Then l_dbl_ValAfeBV = 0
      Call fs_Bono_Verde(l_dbl_ValAfeBV)
   End If
   
   Call fs_Calcul
   Call fs_Calcular_Prestamo
End Sub

Private Sub ipp_ApoPro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoAFP)
      ipp_ApoPro_Change
   End If
End Sub

Private Sub ipp_ComVta_LostFocus()
   Call fs_Calcul
   Call fs_Calcular_Prestamo
End Sub

Private Sub ipp_MtoAFP_Change()
   Call fs_Calcular_MtoSBMS
   
   pnl_MtoPre.Caption = Format(CDbl(pnl_ValTot.Caption) - CDbl(ipp_ApoPro.Text) - (CDbl(pnl_FmvBbp.Caption) + CDbl(pnl_MefPbp.Caption)) - CDbl(ipp_MtoAFP.Text) - CDbl(pnl_MtoBMS.Caption), "###,###,##0.00") & " " 'ipp_ComVta.Text
   If pnl_MtoPre.Caption < 0 Then pnl_MtoPre.Caption = "0.00 "

   If pnl_ValTot.Caption <= modatecli_g_dbl_MtoFin Then 'ipp_ComVta.Text
      If l_int_FlgAfeBV = 1 And l_int_FlgTipAfe <> 0 Or l_dbl_ValAfeBV <> 0 Then Call fs_Bono_Verde(modatecli_g_dbl_BMSTas)
   Else
      If l_int_FlgAfeBV = 0 Or l_int_FlgTipAfe = 0 Or l_dbl_ValAfeBV = 0 Then l_dbl_ValAfeBV = 0
      Call fs_Bono_Verde(l_dbl_ValAfeBV)
   End If
   
   Call fs_Calcul
   Call fs_Calcular_Prestamo
End Sub

Private Sub ipp_MtoAFP_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(chk_Gastos)
      ipp_MtoAFP_Change
   End If
End Sub

Private Sub ipp_MtoAFP_LostFocus()
   Call fs_Calcul
   Call fs_Calcular_Prestamo
End Sub

Private Sub ipp_PlaAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PerGra)
   End If
End Sub

Private Sub ipp_PerGra_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CuoDbl)
   End If
End Sub

Private Sub cmb_CuoDbl_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_DiaPag)
   End If
End Sub

Private Sub cmb_DiaPag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_EmpSeg)
   End If
End Sub

Private Sub cmb_EmpSeg_Click()
   If cmb_EmpSeg.ListIndex > -1 Then
      Screen.MousePointer = 11
      Call moddat_gs_Carga_TipSeg(cmb_SegDes, l_arr_EmpSeg(cmb_EmpSeg.ListIndex + 1).Genera_Codigo)
      Screen.MousePointer = 0
      Call gs_SetFocus(cmb_SegDes)
   Else
      cmb_SegDes.Clear
   End If
End Sub

Private Sub cmb_EmpSeg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_EmpSeg_Click
   End If
End Sub

Private Sub cmb_SegDes_Click()
   If cmb_SegDes.ListIndex > -1 Then
      Call gs_SetFocus(cmb_TasEsp)
   End If
End Sub

Private Sub cmb_TasEsp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_MtoAho.Enabled Then
         Call gs_SetFocus(ipp_MtoAho)
      Else
         Call gs_SetFocus(txt_Observ)
      End If
   End If
End Sub

Private Sub cmb_SegDes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_SegDes_Click
   End If
End Sub

Private Sub cmb_InsFin_Click()
   Call gs_SetFocus(txt_Observ)
End Sub

Private Sub cmb_InsFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_InsFin_Click
   End If
End Sub

Private Sub ipp_MesAho_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_InsFin)
   End If
End Sub

Private Sub cmb_MonAho_Click()
   Call gs_SetFocus(ipp_MesAho)
End Sub

Private Sub cmb_MonAho_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_MonAho_Click
   End If
End Sub

Private Sub ipp_MtoAho_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_MonAho)
   End If
End Sub

Private Sub ipp_ValEst_Change()
   Call fs_Calcul
   Call fs_Calcular_Prestamo
   Call fs_Calcular_GCierre
End Sub

Private Sub ipp_ValEst_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ApoPro)
   End If
End Sub

Private Sub ipp_ValEst_LostFocus()
   Call fs_Calcul
   Call fs_Calcular_Prestamo
   Call fs_Calcular_GCierre
End Sub

Private Sub txt_Observ_GotFocus()
   Call gs_SelecTodo(txt_Observ)
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_ConHip)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub cmb_ConHip_Click()
   Call gs_SetFocus(cmb_EjeSeg)
End Sub

Private Sub cmb_ConHip_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_ConHip_Click
   End If
End Sub

Private Sub cmb_EjeSeg_Click()
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmb_EjeSeg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_EjeSeg_Click
   End If
End Sub

Private Sub grd_Listad_DblClick()
   If grd_Listad.Rows > 0 Then
      grd_Listad.Col = 1
      If grd_Listad.Text = "X" Then
         grd_Listad.Text = ""
      Else
         grd_Listad.Text = "X"
      End If
      Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub

Private Sub grd_Listad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 32 Then
      Call grd_Listad_DblClick
   End If
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub fs_Calcular_GCierre()
Dim r_int_Resul As String

   If chk_Gastos.Value = True Then
   
      'Valida si tiene Datos del inmueble
      If modatecli_g_arr_DatInm(1).DatInm_InmIde = 2 Then
         MsgBox "El Proyecto no está aprobado coordinar con las áreas correspondientes.", vbExclamation, modgen_g_str_NomPlt
         chk_Gastos.Value = False
         Exit Sub
      End If
      
      'Valida que el Proyecto se encuentre aprobado
      If fs_Valida_PryAsg = 0 Then
         MsgBox "El Proyecto seleccionado no se encuentra aprobado. Los Gastos de Cierre se deberán ingresar manualmente.", vbExclamation, modgen_g_str_NomPlt
         chk_Gastos.Value = False
         Exit Sub
      End If
      
      'Valida que se haya ingresado Notaria y Empresa Tasadora al Proyecto
      r_int_Resul = gf_Valida_GastoCierre(moddat_g_str_CodPrd, modatecli_g_arr_DatInm(1).DatInm_CodPry)
   
      If r_int_Resul = 1 Then
         MsgBox "El proyecto no tiene empresa de peritaje asignada, favor actualizar información en la plataforma de Operaciones y/o Proyectos.", vbExclamation, modgen_g_str_NomPlt
         chk_Gastos.Value = False
         Exit Sub
      ElseIf r_int_Resul = 2 Then
         MsgBox "La empresa de peritaje asociada al proyecto no tiene registrado los parámetros necesarios para el cálculo de los gastos de cierre, favor actualizar información en la plataforma de Operaciones.", vbExclamation, modgen_g_str_NomPlt
         chk_Gastos.Value = False
         Exit Sub
      ElseIf r_int_Resul = 3 Then
'         MsgBox "El proyecto asociado no tiene notaría asignada, favor actualizar información en la plataforma de Operaciones y/o Proyectos.", vbExclamation, modgen_g_str_NomPlt
'         chk_Gastos.Value = False
'         Exit Sub
         If MsgBox("La notaria asociada al proyecto no tiene registrado los parámetros necesarios para el cálculo de los gastos de cierre, favor coordinar con el área legal la actualización de la información en caso contrario no se generaran los gastos de cierre." & vbCrLf & "Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            chk_Gastos.Value = False
            Exit Sub
         End If
      ElseIf r_int_Resul = 4 Then
         MsgBox "La notaria asociada al proyecto no tiene registrado los parámetros necesarios para el cálculo de los gastos de cierre, favor actualizar información en la plataforma de Operaciones.", vbExclamation, modgen_g_str_NomPlt
         chk_Gastos.Value = False
         Exit Sub
      End If
   
      pnl_ValGas.Caption = Format(CDbl(gf_Genera_Gastos_Cierre(moddat_g_str_CodPrd, modatecli_g_arr_DatInm(1).DatInm_CodPry, Format(modatecli_g_arr_DatInm(1).DatInm_Modali, "000"), CDbl(ipp_ComVta.Value), CDbl(ipp_ValEst.Value), CDbl(pnl_MtoPre.Caption), 0, 0, 0, 0, 0, 0)), "##,###,##0.00") & " "
      If CDbl(pnl_ValGas.Caption) < 0 Then pnl_ValGas.Caption = "0.00 "
      If CDbl(pnl_ValGas.Caption) = 0 And CDbl(ipp_ValEst.Value = 0) Then pnl_ValGas.Caption = "0.00 "
   Else
      pnl_ValGas.Caption = "0.00 "
   End If
   pnl_TotPre.Caption = Format(CDbl(pnl_MtoPre.Caption) + CDbl(pnl_ValGas.Caption), "##,###,##0.00") & " "
End Sub


Private Sub fs_Calcular_Prestamo()
   'Valor inmueble
   pnl_ValTot.Caption = Format(CDbl(ipp_ComVta.Value) + CDbl(ipp_ValEst.Value), "##,###,##0.00") & " "
   If CDbl(pnl_ValTot.Caption) < 0 Then pnl_ValTot.Caption = "0.00 "
   
   'Inicial
   Call fs_Calcular_MtoSBMS
   If CDbl(l_dbl_MPSMS) <= modatecli_g_dbl_MtoFin Then
      Call fs_Bono_Verde(modatecli_g_dbl_BMSTas)
   End If
   
   pnl_CuoIni.Caption = Format(CDbl(ipp_ApoPro.Value) + CDbl(pnl_FmvBbp.Caption) + CDbl(pnl_MefPbp.Caption) + CDbl(ipp_MtoAFP.Text) + CDbl(pnl_MtoBMS.Caption), "##,###,##0.00") & " "
   If CDbl(pnl_CuoIni.Caption) < 0 Then pnl_CuoIni.Caption = "0.00 "
   
   'Prestamo
   pnl_MtoPre.Caption = Format(CDbl(pnl_ValTot.Caption) - CDbl(pnl_CuoIni.Caption), "##,###,##0.00") & " "
   If CDbl(pnl_MtoPre.Caption) < 0 Then pnl_MtoPre.Caption = "0.00 "
   pnl_TotPre.Caption = Format(CDbl(pnl_MtoPre.Caption) + CDbl(pnl_ValGas.Caption), "##,###,##0.00") & " "
   DoEvents
End Sub
