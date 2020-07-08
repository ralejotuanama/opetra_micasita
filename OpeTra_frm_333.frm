VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_Con_Cuadre_01 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9915
   ClientLeft      =   5085
   ClientTop       =   2115
   ClientWidth     =   11775
   Icon            =   "OpeTra_frm_333.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9915
   ScaleMode       =   0  'User
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel111 
      Height          =   9915
      Left            =   0
      TabIndex        =   53
      Top             =   0
      Width           =   11775
      _Version        =   65536
      _ExtentX        =   20770
      _ExtentY        =   17489
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
      Begin TabDlg.SSTab SSTab1 
         Height          =   6075
         Left            =   60
         TabIndex        =   62
         Top             =   3750
         Width           =   11625
         _ExtentX        =   20505
         _ExtentY        =   10716
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Datos de Saldos"
         TabPicture(0)   =   "OpeTra_frm_333.frx":000C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SSPanel2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "SSPanel5"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Provisiones y Otros Datos"
         TabPicture(1)   =   "OpeTra_frm_333.frx":0028
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SSPanel6"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "SSPanel8"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "SSPanel9"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).ControlCount=   3
         Begin Threed.SSPanel SSPanel9 
            Height          =   1545
            Left            =   -69150
            TabIndex        =   144
            Top             =   2430
            Width           =   5685
            _Version        =   65536
            _ExtentX        =   10028
            _ExtentY        =   2725
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
            Begin VB.ComboBox cmb_MonGar 
               Height          =   315
               Left            =   1545
               Style           =   2  'Dropdown List
               TabIndex        =   51
               Top             =   780
               Width           =   2760
            End
            Begin VB.ComboBox cmb_TipGar 
               Height          =   315
               Left            =   1545
               Style           =   2  'Dropdown List
               TabIndex        =   50
               Top             =   450
               Width           =   2760
            End
            Begin EditLib.fpDoubleSingle fpd_MtoGar 
               Height          =   315
               Left            =   1545
               TabIndex        =   52
               Top             =   1110
               Width           =   1170
               _Version        =   196608
               _ExtentX        =   2064
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
               MaxValue        =   "9000000000"
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
            Begin VB.Label Label76 
               Caption         =   "Monto Garantía"
               Height          =   315
               Left            =   120
               TabIndex        =   148
               Top             =   1170
               Width           =   1365
            End
            Begin VB.Label Label75 
               Caption         =   "Moneda Garantía"
               Height          =   195
               Left            =   120
               TabIndex        =   147
               Top             =   810
               Width           =   1365
            End
            Begin VB.Label Label67 
               Caption         =   "Tipo Garantía"
               Height          =   195
               Left            =   120
               TabIndex        =   146
               Top             =   480
               Width           =   1365
            End
            Begin VB.Label Label74 
               Caption         =   "Otros Datos"
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
               Left            =   120
               TabIndex        =   145
               Top             =   120
               UseMnemonic     =   0   'False
               Width           =   3015
            End
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   1545
            Left            =   -74910
            TabIndex        =   136
            Top             =   2430
            Width           =   5715
            _Version        =   65536
            _ExtentX        =   10081
            _ExtentY        =   2725
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
            Begin VB.ComboBox cmb_Castig 
               Height          =   315
               Left            =   4410
               Style           =   2  'Dropdown List
               TabIndex        =   49
               Top             =   1080
               Width           =   1140
            End
            Begin VB.ComboBox cmb_Judici 
               Height          =   315
               Left            =   4410
               Style           =   2  'Dropdown List
               TabIndex        =   48
               Top             =   750
               Width           =   1140
            End
            Begin VB.ComboBox cmb_Refina 
               Height          =   315
               Left            =   4410
               Style           =   2  'Dropdown List
               TabIndex        =   47
               Top             =   420
               Width           =   1140
            End
            Begin EditLib.fpDoubleSingle fpd_CuoPag 
               Height          =   315
               Left            =   1530
               TabIndex        =   45
               Top             =   750
               Width           =   1140
               _Version        =   196608
               _ExtentX        =   2011
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
               Text            =   "0"
               DecimalPlaces   =   0
               DecimalPoint    =   "."
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "360"
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
            Begin EditLib.fpDoubleSingle fpd_NumCuo 
               Height          =   315
               Left            =   1530
               TabIndex        =   44
               Top             =   420
               Width           =   1140
               _Version        =   196608
               _ExtentX        =   2011
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
               Text            =   "0"
               DecimalPlaces   =   0
               DecimalPoint    =   "."
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "360"
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
            Begin EditLib.fpDoubleSingle fpd_CuoPen 
               Height          =   315
               Left            =   1530
               TabIndex        =   46
               Top             =   1080
               Width           =   1140
               _Version        =   196608
               _ExtentX        =   2011
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
               Text            =   "0"
               DecimalPlaces   =   0
               DecimalPoint    =   "."
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "360"
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
            Begin VB.Label Label73 
               Caption         =   "Flag Castigado:"
               Height          =   195
               Left            =   2895
               TabIndex        =   143
               Top             =   1110
               Width           =   1515
            End
            Begin VB.Label Label72 
               Caption         =   "Flag Judicial:"
               Height          =   195
               Left            =   2895
               TabIndex        =   142
               Top             =   780
               Width           =   1515
            End
            Begin VB.Label Label71 
               Caption         =   "Flag Refinanciado:"
               Height          =   195
               Left            =   2895
               TabIndex        =   141
               Top             =   450
               Width           =   1515
            End
            Begin VB.Label Label70 
               Caption         =   "Cuotas Pagadas:"
               Height          =   270
               Left            =   120
               TabIndex        =   140
               Top             =   810
               Width           =   1425
            End
            Begin VB.Label Label69 
               Caption         =   "N° de Cuotas:"
               Height          =   315
               Left            =   120
               TabIndex        =   139
               Top             =   480
               Width           =   1275
            End
            Begin VB.Label Label68 
               Caption         =   "Cuotas Pendientes:"
               Height          =   315
               Left            =   120
               TabIndex        =   138
               Top             =   1140
               Width           =   1425
            End
            Begin VB.Label Label66 
               Caption         =   "Otros Datos"
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
               Left            =   120
               TabIndex        =   137
               Top             =   90
               UseMnemonic     =   0   'False
               Width           =   3015
            End
         End
         Begin Threed.SSPanel SSPanel5 
            Height          =   2295
            Left            =   90
            TabIndex        =   63
            Top             =   450
            Width           =   11445
            _Version        =   65536
            _ExtentX        =   20188
            _ExtentY        =   4048
            _StockProps     =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Begin EditLib.fpDoubleSingle fpd_CapAmort 
               Height          =   315
               Left            =   2040
               TabIndex        =   11
               Top             =   1170
               Width           =   1320
               _Version        =   196608
               _ExtentX        =   2328
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
               MaxValue        =   "9000000000"
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
            Begin EditLib.fpDoubleSingle fpd_IntSuspC 
               Height          =   315
               Left            =   9990
               TabIndex        =   16
               Top             =   840
               Width           =   1230
               _Version        =   196608
               _ExtentX        =   2170
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
               MaxValue        =   "9000000000"
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
            Begin EditLib.fpDoubleSingle fpd_SalConcC 
               Height          =   315
               Left            =   6060
               TabIndex        =   13
               Top             =   840
               Width           =   1320
               _Version        =   196608
               _ExtentX        =   2328
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
               MaxValue        =   "9000000000"
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
            Begin Threed.SSPanel pnl_CapDesmb 
               Height          =   315
               Left            =   2040
               TabIndex        =   64
               Top             =   510
               Width           =   1320
               _Version        =   65536
               _ExtentX        =   2328
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "135,000.00 "
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
               RoundedCorners  =   0   'False
               Font3D          =   2
               Alignment       =   4
            End
            Begin EditLib.fpDoubleSingle fpd_SalNConC 
               Height          =   315
               Left            =   6060
               TabIndex        =   12
               Top             =   510
               Width           =   1320
               _Version        =   196608
               _ExtentX        =   2328
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
               MaxValue        =   "9000000000"
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
            Begin EditLib.fpDoubleSingle fpd_IntDevnC 
               Height          =   315
               Left            =   9990
               TabIndex        =   15
               Top             =   510
               Width           =   1230
               _Version        =   196608
               _ExtentX        =   2170
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
               MaxValue        =   "9000000000"
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
            Begin Threed.SSPanel pnl_SldDeud1 
               Height          =   315
               Left            =   2040
               TabIndex        =   65
               Top             =   1830
               Width           =   1320
               _Version        =   65536
               _ExtentX        =   2328
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "15,000.00 "
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
               RoundedCorners  =   0   'False
               Font3D          =   2
               Alignment       =   4
            End
            Begin EditLib.fpDoubleSingle fpd_CapVencd 
               Height          =   315
               Left            =   9990
               TabIndex        =   19
               Top             =   1830
               Width           =   1230
               _Version        =   196608
               _ExtentX        =   2170
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
               MaxValue        =   "9000000000"
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
            Begin EditLib.fpLongInteger fpd_DiaMoroC 
               Height          =   315
               Left            =   9990
               TabIndex        =   18
               Top             =   1500
               Width           =   1230
               _Version        =   196608
               _ExtentX        =   2170
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
               Text            =   "0"
               MaxValue        =   "2147483647"
               MinValue        =   "0"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ""
               UseSeparator    =   0   'False
               IncInt          =   1
               BorderGrayAreaColor=   -2147483637
               ThreeDOnFocusInvert=   0   'False
               ThreeDFrameColor=   -2147483637
               Appearance      =   2
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483637
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin Threed.SSPanel pnl_SldDeud2 
               Height          =   315
               Left            =   6060
               TabIndex        =   66
               Top             =   1830
               Width           =   1320
               _Version        =   65536
               _ExtentX        =   2328
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "15,000.00 "
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
               RoundedCorners  =   0   'False
               Font3D          =   2
               Alignment       =   4
            End
            Begin EditLib.fpDoubleSingle fpd_IntMoraC 
               Height          =   315
               Left            =   9990
               TabIndex        =   17
               Top             =   1170
               Width           =   1230
               _Version        =   196608
               _ExtentX        =   2170
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
               MaxValue        =   "9000000000"
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
            Begin EditLib.fpDoubleSingle fpd_IntCapit 
               Height          =   315
               Left            =   2040
               TabIndex        =   10
               Top             =   840
               Width           =   1320
               _Version        =   196608
               _ExtentX        =   2328
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
               MaxValue        =   "9000000000"
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
            Begin EditLib.fpDoubleSingle fpd_PBPPerdi 
               Height          =   315
               Left            =   6060
               TabIndex        =   14
               Top             =   1170
               Width           =   1320
               _Version        =   196608
               _ExtentX        =   2328
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
               MaxValue        =   "9000000000"
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
            Begin VB.Label Label17 
               Caption         =   "Capital Amortizado:"
               Height          =   315
               Left            =   120
               TabIndex        =   88
               Top             =   1200
               Width           =   1725
            End
            Begin VB.Label Label19 
               Caption         =   "Saldo Concesional:"
               Height          =   315
               Left            =   4095
               TabIndex        =   87
               Top             =   870
               Width           =   1770
            End
            Begin VB.Label Label20 
               Caption         =   "Interés Capitalizado:"
               Height          =   315
               Left            =   120
               TabIndex        =   86
               Top             =   870
               Width           =   1725
            End
            Begin VB.Label Label21 
               Caption         =   "Capital Desembolsado:"
               Height          =   315
               Left            =   120
               TabIndex        =   85
               Top             =   540
               Width           =   1725
            End
            Begin VB.Label Label26 
               Caption         =   "Saldo No Concesional:"
               Height          =   315
               Left            =   4095
               TabIndex        =   84
               Top             =   540
               Width           =   1770
            End
            Begin VB.Label Label29 
               Caption         =   "Información del Padrón de Deudores (Contabilidad)"
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
               Left            =   120
               TabIndex        =   83
               Top             =   120
               UseMnemonic     =   0   'False
               Width           =   4965
            End
            Begin VB.Label Label7 
               Caption         =   "Interés en Suspenso:"
               Height          =   315
               Left            =   8220
               TabIndex        =   82
               Top             =   870
               Width           =   1755
            End
            Begin VB.Label Label6 
               Caption         =   "Interés Devengado:"
               Height          =   315
               Left            =   8220
               TabIndex        =   81
               Top             =   540
               Width           =   1755
            End
            Begin VB.Label Label1 
               Caption         =   "Saldo de la Deuda (1):"
               Height          =   315
               Left            =   120
               TabIndex        =   80
               Top             =   1860
               Width           =   1725
            End
            Begin VB.Line Line1 
               X1              =   120
               X2              =   3480
               Y1              =   1650
               Y2              =   1650
            End
            Begin VB.Label Label3 
               Caption         =   "+"
               Height          =   315
               Left            =   3465
               TabIndex        =   79
               Top             =   540
               Width           =   165
            End
            Begin VB.Label Label4 
               Caption         =   "+"
               Height          =   315
               Left            =   3465
               TabIndex        =   78
               Top             =   900
               Width           =   165
            End
            Begin VB.Label Label5 
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   3495
               TabIndex        =   77
               Top             =   1230
               Width           =   165
            End
            Begin VB.Label Label8 
               Caption         =   "PBP Perdido (TC):"
               Height          =   315
               Left            =   4095
               TabIndex        =   76
               Top             =   1200
               Width           =   1770
            End
            Begin VB.Line Line2 
               X1              =   4050
               X2              =   7380
               Y1              =   1650
               Y2              =   1650
            End
            Begin VB.Label Label9 
               Caption         =   "Capital Vencido:"
               Height          =   315
               Left            =   8220
               TabIndex        =   75
               Top             =   1860
               Width           =   1755
            End
            Begin VB.Label Label11 
               Caption         =   "Días de Morosidad:"
               Height          =   285
               Left            =   8220
               TabIndex        =   74
               Top             =   1530
               Width           =   1755
            End
            Begin VB.Label Label22 
               Caption         =   "Saldo de la Deuda (2):"
               Height          =   315
               Left            =   4095
               TabIndex        =   73
               Top             =   1860
               Width           =   1770
            End
            Begin VB.Label Label23 
               Caption         =   "+"
               Height          =   315
               Left            =   7470
               TabIndex        =   72
               Top             =   540
               Width           =   165
            End
            Begin VB.Label Label24 
               Caption         =   "+"
               Height          =   315
               Left            =   7470
               TabIndex        =   71
               Top             =   900
               Width           =   165
            End
            Begin VB.Label Label25 
               Caption         =   "+"
               Height          =   315
               Left            =   7470
               TabIndex        =   70
               Top             =   1230
               Width           =   165
            End
            Begin VB.Label Label30 
               Caption         =   "="
               Height          =   315
               Left            =   3465
               TabIndex        =   69
               Top             =   1860
               Width           =   165
            End
            Begin VB.Label Label31 
               Caption         =   "="
               Height          =   315
               Left            =   7485
               TabIndex        =   68
               Top             =   1860
               Width           =   165
            End
            Begin VB.Label Label12 
               Caption         =   "Interés Moratorio:"
               Height          =   315
               Left            =   8220
               TabIndex        =   67
               Top             =   1200
               Width           =   1755
            End
            Begin VB.Line Line6 
               X1              =   3840
               X2              =   3840
               Y1              =   510
               Y2              =   2130
            End
            Begin VB.Line Line7 
               X1              =   7890
               X2              =   7890
               Y1              =   510
               Y2              =   2130
            End
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   3165
            Left            =   90
            TabIndex        =   89
            Top             =   2790
            Width           =   11445
            _Version        =   65536
            _ExtentX        =   20188
            _ExtentY        =   5583
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
            Begin EditLib.fpDoubleSingle fpd_SalNConO 
               Height          =   315
               Left            =   6060
               TabIndex        =   22
               Top             =   510
               Width           =   1320
               _Version        =   196608
               _ExtentX        =   2328
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
               MaxValue        =   "9000000000"
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
            Begin EditLib.fpDoubleSingle fpd_SalConcO 
               Height          =   315
               Left            =   6060
               TabIndex        =   23
               Top             =   840
               Width           =   1320
               _Version        =   196608
               _ExtentX        =   2328
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
               MaxValue        =   "9000000000"
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
            Begin Threed.SSPanel pnl_SldDeud4 
               Height          =   315
               Left            =   6060
               TabIndex        =   90
               Top             =   1500
               Width           =   1320
               _Version        =   65536
               _ExtentX        =   2328
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "15,000.00 "
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
               RoundedCorners  =   0   'False
               Font3D          =   2
               Alignment       =   4
            End
            Begin EditLib.fpDoubleSingle fpd_TotCapVi 
               Height          =   315
               Left            =   2040
               TabIndex        =   20
               Top             =   510
               Width           =   1320
               _Version        =   196608
               _ExtentX        =   2328
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
               MaxValue        =   "9000000000"
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
            Begin EditLib.fpDoubleSingle fpd_TotCapVe 
               Height          =   315
               Left            =   2040
               TabIndex        =   21
               Top             =   840
               Width           =   1320
               _Version        =   196608
               _ExtentX        =   2328
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
               MaxValue        =   "9000000000"
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
            Begin Threed.SSPanel pnl_SldDeud3 
               Height          =   315
               Left            =   2040
               TabIndex        =   91
               Top             =   1500
               Width           =   1320
               _Version        =   65536
               _ExtentX        =   2328
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "15,000.00 "
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
               RoundedCorners  =   0   'False
               Font3D          =   2
               Alignment       =   4
            End
            Begin EditLib.fpDoubleSingle fpd_IntSuspO 
               Height          =   315
               Left            =   9990
               TabIndex        =   25
               Top             =   840
               Width           =   1230
               _Version        =   196608
               _ExtentX        =   2170
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
               MaxValue        =   "9000000000"
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
            Begin EditLib.fpDoubleSingle fpd_IntDevnO 
               Height          =   315
               Left            =   9990
               TabIndex        =   24
               Top             =   510
               Width           =   1230
               _Version        =   196608
               _ExtentX        =   2170
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
               MaxValue        =   "9000000000"
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
            Begin EditLib.fpLongInteger fpd_DiaMoroO 
               Height          =   315
               Left            =   9990
               TabIndex        =   27
               Top             =   1500
               Width           =   1230
               _Version        =   196608
               _ExtentX        =   2170
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
               Text            =   "0"
               MaxValue        =   "2147483647"
               MinValue        =   "0"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ""
               UseSeparator    =   0   'False
               IncInt          =   1
               BorderGrayAreaColor=   -2147483637
               ThreeDOnFocusInvert=   0   'False
               ThreeDFrameColor=   -2147483637
               Appearance      =   2
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483637
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDoubleSingle fpd_IntMoraO 
               Height          =   315
               Left            =   9990
               TabIndex        =   26
               Top             =   1170
               Width           =   1230
               _Version        =   196608
               _ExtentX        =   2170
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
               MaxValue        =   "9000000000"
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
            Begin Threed.SSPanel pnl_PrstaCon 
               Height          =   315
               Left            =   2040
               TabIndex        =   92
               Top             =   2670
               Width           =   1320
               _Version        =   65536
               _ExtentX        =   2328
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "15,000.00 "
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
               RoundedCorners  =   0   'False
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_SaldNCon 
               Height          =   315
               Left            =   3930
               TabIndex        =   93
               Top             =   2340
               Width           =   1320
               _Version        =   65536
               _ExtentX        =   2328
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "15,000.00 "
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
               RoundedCorners  =   0   'False
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_SaldoCon 
               Height          =   315
               Left            =   3930
               TabIndex        =   94
               Top             =   2670
               Width           =   1320
               _Version        =   65536
               _ExtentX        =   2328
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "15,000.00 "
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
               RoundedCorners  =   0   'False
               Font3D          =   2
               Alignment       =   4
            End
            Begin EditLib.fpDoubleSingle fpd_AmortCon 
               Height          =   315
               Left            =   5790
               TabIndex        =   30
               Top             =   2670
               Width           =   1230
               _Version        =   196608
               _ExtentX        =   2170
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
               MaxValue        =   "9000000000"
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
            Begin EditLib.fpDoubleSingle fpd_AmorNCon 
               Height          =   315
               Left            =   5790
               TabIndex        =   29
               Top             =   2340
               Width           =   1230
               _Version        =   196608
               _ExtentX        =   2170
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
               MaxValue        =   "9000000000"
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
            Begin EditLib.fpDoubleSingle fpd_PrstNCon 
               Height          =   315
               Left            =   2040
               TabIndex        =   28
               Top             =   2340
               Width           =   1320
               _Version        =   196608
               _ExtentX        =   2328
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
               MaxValue        =   "9000000000"
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
            Begin VB.Label Label13 
               Caption         =   "Información del Maestro de Saldos (Operaciones)"
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
               Left            =   120
               TabIndex        =   120
               Top             =   120
               UseMnemonic     =   0   'False
               Width           =   4605
            End
            Begin VB.Label Label33 
               Caption         =   "Saldo Concesional (Incluye PBP Perdido):"
               Height          =   450
               Left            =   4095
               TabIndex        =   119
               Top             =   810
               Width           =   1785
            End
            Begin VB.Label Label39 
               Caption         =   "Saldo No Concesional:"
               Height          =   270
               Left            =   4095
               TabIndex        =   118
               Top             =   540
               Width           =   1785
            End
            Begin VB.Line Line3 
               X1              =   120
               X2              =   3480
               Y1              =   1320
               Y2              =   1320
            End
            Begin VB.Label Label15 
               Caption         =   "Saldo de la Deuda (4):"
               Height          =   315
               Left            =   4095
               TabIndex        =   117
               Top             =   1530
               Width           =   1725
            End
            Begin VB.Label Label16 
               Caption         =   "+"
               Height          =   315
               Left            =   7470
               TabIndex        =   116
               Top             =   870
               Width           =   165
            End
            Begin VB.Label Label28 
               Caption         =   "+"
               Height          =   315
               Left            =   7470
               TabIndex        =   115
               Top             =   570
               Width           =   165
            End
            Begin VB.Label Label14 
               Caption         =   "Total Capital Vigente:"
               Height          =   315
               Left            =   120
               TabIndex        =   114
               Top             =   540
               Width           =   1785
            End
            Begin VB.Label Label18 
               Caption         =   "Total Capital Vencido:"
               Height          =   315
               Left            =   120
               TabIndex        =   113
               Top             =   870
               Width           =   1785
            End
            Begin VB.Line Line4 
               X1              =   4050
               X2              =   7380
               Y1              =   1320
               Y2              =   1320
            End
            Begin VB.Label Label27 
               Caption         =   "Saldo de la Deuda (3):"
               Height          =   315
               Left            =   120
               TabIndex        =   112
               Top             =   1530
               Width           =   1725
            End
            Begin VB.Label Label32 
               Caption         =   "+"
               Height          =   315
               Left            =   3450
               TabIndex        =   111
               Top             =   900
               Width           =   165
            End
            Begin VB.Label Label34 
               Caption         =   "+"
               Height          =   315
               Left            =   3450
               TabIndex        =   110
               Top             =   570
               Width           =   165
            End
            Begin VB.Label Label10 
               Caption         =   "Días de Morosidad:"
               Height          =   285
               Left            =   8220
               TabIndex        =   109
               Top             =   1530
               Width           =   1755
            End
            Begin VB.Label Label35 
               Caption         =   "Interés Devengado:"
               Height          =   315
               Left            =   8220
               TabIndex        =   108
               Top             =   540
               Width           =   1755
            End
            Begin VB.Label Label36 
               Caption         =   "Interés en Suspenso:"
               Height          =   315
               Left            =   8220
               TabIndex        =   107
               Top             =   870
               Width           =   1755
            End
            Begin VB.Label Label37 
               Caption         =   "Interés Moratorio:"
               Height          =   315
               Left            =   8220
               TabIndex        =   106
               Top             =   1200
               Width           =   1755
            End
            Begin VB.Line Line5 
               X1              =   120
               X2              =   11370
               Y1              =   1980
               Y2              =   1980
            End
            Begin VB.Label Label38 
               Caption         =   "Tramo No Concesional:"
               Height          =   315
               Left            =   90
               TabIndex        =   105
               Top             =   2370
               Width           =   1725
            End
            Begin VB.Label Label40 
               Caption         =   "Tramo Concesional:"
               Height          =   315
               Left            =   90
               TabIndex        =   104
               Top             =   2700
               Width           =   1725
            End
            Begin VB.Label Label41 
               Caption         =   "="
               Height          =   315
               Left            =   3450
               TabIndex        =   103
               Top             =   1530
               Width           =   165
            End
            Begin VB.Label Label42 
               Caption         =   "="
               Height          =   315
               Left            =   7470
               TabIndex        =   102
               Top             =   1530
               Width           =   165
            End
            Begin VB.Label Label43 
               Caption         =   "="
               Height          =   315
               Left            =   3450
               TabIndex        =   101
               Top             =   2370
               Width           =   165
            End
            Begin VB.Label Label44 
               Caption         =   "="
               Height          =   315
               Left            =   3450
               TabIndex        =   100
               Top             =   2700
               Width           =   165
            End
            Begin VB.Label Label47 
               Caption         =   "+"
               Height          =   315
               Left            =   5460
               TabIndex        =   99
               Top             =   2370
               Width           =   165
            End
            Begin VB.Label Label48 
               Caption         =   "+"
               Height          =   315
               Left            =   5460
               TabIndex        =   98
               Top             =   2700
               Width           =   165
            End
            Begin VB.Line Line8 
               X1              =   3840
               X2              =   3840
               Y1              =   510
               Y2              =   1800
            End
            Begin VB.Line Line9 
               X1              =   7890
               X2              =   7890
               Y1              =   390
               Y2              =   1680
            End
            Begin VB.Label Label45 
               Caption         =   "Monto Préstamo"
               Height          =   315
               Left            =   2100
               TabIndex        =   97
               Top             =   2130
               Width           =   1305
            End
            Begin VB.Label Label46 
               Caption         =   "Monto Saldo"
               Height          =   315
               Left            =   4140
               TabIndex        =   96
               Top             =   2130
               Width           =   1035
            End
            Begin VB.Label Label49 
               Caption         =   "Monto Amortizado"
               Height          =   315
               Left            =   5760
               TabIndex        =   95
               Top             =   2130
               Width           =   1335
            End
         End
         Begin Threed.SSPanel SSPanel6 
            Height          =   1905
            Left            =   -74910
            TabIndex        =   121
            Top             =   450
            Width           =   11445
            _Version        =   65536
            _ExtentX        =   20188
            _ExtentY        =   3360
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
            Begin EditLib.fpDoubleSingle fpd_CbrFmvRC 
               Height          =   315
               Left            =   4410
               TabIndex        =   43
               Top             =   1410
               Width           =   1170
               _Version        =   196608
               _ExtentX        =   2064
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
               MaxValue        =   "9000000000"
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
            Begin EditLib.fpDoubleSingle fpd_ClaAli 
               Height          =   315
               Left            =   4410
               TabIndex        =   32
               Top             =   420
               Width           =   1170
               _Version        =   196608
               _ExtentX        =   2064
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
               Text            =   "0"
               DecimalPlaces   =   0
               DecimalPoint    =   "."
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "9"
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
            Begin EditLib.fpDoubleSingle fpd_ClaCli 
               Height          =   315
               Left            =   1530
               TabIndex        =   31
               Top             =   420
               Width           =   1140
               _Version        =   196608
               _ExtentX        =   2011
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
               Text            =   "0"
               DecimalPlaces   =   0
               DecimalPoint    =   "."
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "9"
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
            Begin EditLib.fpDoubleSingle fpd_ClaPrv 
               Height          =   315
               Left            =   7290
               TabIndex        =   33
               Top             =   420
               Width           =   1170
               _Version        =   196608
               _ExtentX        =   2064
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
               Text            =   "0"
               DecimalPlaces   =   0
               DecimalPoint    =   "."
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "9"
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
            Begin EditLib.fpDoubleSingle fpd_PrvCic 
               Height          =   315
               Left            =   1530
               TabIndex        =   39
               Top             =   1080
               Width           =   1140
               _Version        =   196608
               _ExtentX        =   2011
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
               MaxValue        =   "9000000000"
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
            Begin EditLib.fpDoubleSingle fpd_PrvGen 
               Height          =   315
               Left            =   1530
               TabIndex        =   35
               Top             =   750
               Width           =   1140
               _Version        =   196608
               _ExtentX        =   2011
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
               MaxValue        =   "9000000000"
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
            Begin EditLib.fpDoubleSingle fpd_PrvEsp 
               Height          =   315
               Left            =   7290
               TabIndex        =   37
               Top             =   750
               Width           =   1170
               _Version        =   196608
               _ExtentX        =   2064
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
               MaxValue        =   "9000000000"
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
            Begin EditLib.fpDoubleSingle fpd_PrvVol 
               Height          =   315
               Left            =   10110
               TabIndex        =   38
               Top             =   750
               Width           =   1170
               _Version        =   196608
               _ExtentX        =   2064
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
               MaxValue        =   "9000000000"
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
            Begin EditLib.fpDoubleSingle fpd_Aplica 
               Height          =   315
               Left            =   10110
               TabIndex        =   34
               Top             =   420
               Width           =   1170
               _Version        =   196608
               _ExtentX        =   2064
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
               MaxValue        =   "9000000000"
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
            Begin EditLib.fpDoubleSingle fpd_PrvGenRC 
               Height          =   315
               Left            =   4410
               TabIndex        =   36
               Top             =   750
               Width           =   1170
               _Version        =   196608
               _ExtentX        =   2064
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
               MaxValue        =   "9000000000"
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
            Begin EditLib.fpDoubleSingle fpd_PrvCicRC 
               Height          =   315
               Left            =   4410
               TabIndex        =   40
               Top             =   1080
               Width           =   1170
               _Version        =   196608
               _ExtentX        =   2064
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
               MaxValue        =   "9000000000"
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
            Begin EditLib.fpDoubleSingle fpd_PrvRip 
               Height          =   315
               Left            =   7290
               TabIndex        =   41
               Top             =   1080
               Width           =   1170
               _Version        =   196608
               _ExtentX        =   2064
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
               MaxValue        =   "9000000000"
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
            Begin EditLib.fpDoubleSingle fpd_CbrFmv 
               Height          =   315
               Left            =   1530
               TabIndex        =   42
               Top             =   1410
               Width           =   1140
               _Version        =   196608
               _ExtentX        =   2011
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
               MaxValue        =   "9000000000"
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
            Begin VB.Label Label65 
               Caption         =   "Cobertura FMV RC:"
               Height          =   270
               Left            =   2895
               TabIndex        =   135
               Top             =   1470
               Width           =   1425
            End
            Begin VB.Label Label64 
               Caption         =   "Cobertura FMV:"
               Height          =   270
               Left            =   120
               TabIndex        =   134
               Top             =   1470
               Width           =   1275
            End
            Begin VB.Label Label63 
               Caption         =   "Prov. Riesgo País:"
               Height          =   315
               Left            =   5850
               TabIndex        =   133
               Top             =   1140
               Width           =   1425
            End
            Begin VB.Label Label62 
               Caption         =   "Prov. Ciclica RC:"
               Height          =   270
               Left            =   2895
               TabIndex        =   132
               Top             =   1140
               Width           =   1425
            End
            Begin VB.Label Label61 
               Caption         =   "Prov. Generica RC:"
               Height          =   315
               Left            =   2895
               TabIndex        =   131
               Top             =   810
               Width           =   1425
            End
            Begin VB.Label Label60 
               AutoSize        =   -1  'True
               Caption         =   "Aplicación:"
               Height          =   195
               Left            =   8655
               TabIndex        =   130
               Top             =   480
               Width           =   1275
            End
            Begin VB.Label Label59 
               Caption         =   "Prov. Voluntaria:"
               Height          =   315
               Left            =   8640
               TabIndex        =   129
               Top             =   810
               Width           =   1275
            End
            Begin VB.Label Label52 
               Caption         =   "Provisiones"
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
               Left            =   120
               TabIndex        =   128
               Top             =   120
               UseMnemonic     =   0   'False
               Width           =   3015
            End
            Begin VB.Label Label53 
               Caption         =   "Clasifica. Provisión:"
               Height          =   315
               Left            =   5850
               TabIndex        =   127
               Top             =   480
               Width           =   1425
            End
            Begin VB.Label Label54 
               Caption         =   "Clasifica. Interna:"
               Height          =   315
               Left            =   120
               TabIndex        =   126
               Top             =   480
               Width           =   1275
            End
            Begin VB.Label Label55 
               Caption         =   "Clasifica. Alineada:"
               Height          =   270
               Left            =   2895
               TabIndex        =   125
               Top             =   480
               Width           =   1425
            End
            Begin VB.Label Label56 
               Caption         =   "Prov. Especifica:"
               Height          =   315
               Left            =   5850
               TabIndex        =   124
               Top             =   810
               Width           =   1425
            End
            Begin VB.Label Label57 
               Caption         =   "Prov. Generica:"
               Height          =   315
               Left            =   120
               TabIndex        =   123
               Top             =   810
               Width           =   1275
            End
            Begin VB.Label Label58 
               Caption         =   "Prov. Ciclica:"
               Height          =   270
               Left            =   120
               TabIndex        =   122
               Top             =   1140
               Width           =   1275
            End
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   2235
         Left            =   60
         TabIndex        =   56
         Top             =   1470
         Width           =   11655
         _Version        =   65536
         _ExtentX        =   20558
         _ExtentY        =   3942
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
            Height          =   1815
            Left            =   60
            TabIndex        =   9
            Top             =   330
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   3201
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
            Caption         =   "Datos del Crédito"
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
            TabIndex        =   58
            Top             =   90
            Width           =   1875
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   675
         Left            =   60
         TabIndex        =   54
         Top             =   60
         Width           =   11655
         _Version        =   65536
         _ExtentX        =   20558
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
            Left            =   690
            TabIndex        =   57
            Top             =   30
            Width           =   5565
            _Version        =   65536
            _ExtentX        =   9816
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "Cuadre de Operaciones: Padrón Contable vs Saldos Operativos"
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
            Picture         =   "OpeTra_frm_333.frx":0044
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   645
         Left            =   60
         TabIndex        =   55
         Top             =   780
         Width           =   11655
         _Version        =   65536
         _ExtentX        =   20558
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
            Left            =   8040
            Picture         =   "OpeTra_frm_333.frx":034E
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Buscar Operación"
            Top             =   30
            Width           =   585
         End
         Begin VB.ComboBox cmb_PerMes 
            Height          =   315
            Left            =   3390
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   180
            Width           =   1725
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   10440
            Picture         =   "OpeTra_frm_333.frx":0658
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   11040
            Picture         =   "OpeTra_frm_333.frx":0962
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Enabled         =   0   'False
            Height          =   585
            Left            =   9840
            Picture         =   "OpeTra_frm_333.frx":0DA4
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_VerPag 
            Enabled         =   0   'False
            Height          =   585
            Left            =   8640
            Picture         =   "OpeTra_frm_333.frx":11E6
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Consulta Pagos del Cliente"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ImpCro 
            Enabled         =   0   'False
            Height          =   585
            Left            =   9240
            Picture         =   "OpeTra_frm_333.frx":14F0
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Consulta Cronograma de Pagos"
            Top             =   30
            Width           =   585
         End
         Begin MSMask.MaskEdBox msk_NumOpe 
            Height          =   315
            Left            =   1410
            TabIndex        =   0
            Top             =   180
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            Mask            =   "###-##-#####"
            PromptChar      =   " "
         End
         Begin EditLib.fpDoubleSingle ipp_PerAno 
            Height          =   315
            Left            =   5910
            TabIndex        =   2
            Top             =   180
            Width           =   825
            _Version        =   196608
            _ExtentX        =   1455
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
            DecimalPlaces   =   0
            DecimalPoint    =   "."
            FixedPoint      =   0   'False
            LeadZero        =   0
            MaxValue        =   "9999"
            MinValue        =   "1900"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   0   'False
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
         Begin VB.Label Label51 
            Caption         =   "Año:"
            Height          =   255
            Left            =   5460
            TabIndex        =   61
            Top             =   240
            Width           =   405
         End
         Begin VB.Label Label50 
            Caption         =   "Mes:"
            Height          =   315
            Left            =   2910
            TabIndex        =   60
            Top             =   210
            Width           =   555
         End
         Begin VB.Label lbl_Numero 
            Caption         =   "Nro. Operación:"
            Height          =   285
            Left            =   120
            TabIndex        =   59
            Top             =   210
            Width           =   1275
         End
      End
   End
End
Attribute VB_Name = "frm_Con_Cuadre_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_dbl_MtoPre           As Double
Dim l_dbl_PreNco           As Double
Dim l_dbl_PreCon           As Double
Dim l_dbl_SalNco           As Double
Dim l_dbl_SalCon           As Double

'Lista de variables para ser usadas en el proceso de Auditoria
Dim l_dbl_IntCapCtb        As Double
Dim l_dbl_CapAmoCtb        As Double
Dim l_dbl_SalNoConCtb      As Double
Dim l_dbl_SalConCtb        As Double
Dim l_dbl_SalPBPCtb        As Double
Dim l_dbl_IntDevCtb        As Double
Dim l_dbl_IntSusCtb        As Double
Dim l_dbl_IntMorCtb        As Double
Dim l_int_DiaMorCtb        As Integer
Dim l_dbl_CapVenCtb        As Double
Dim l_dbl_CapVigOpe        As Double
Dim l_dbl_CapVenOpe        As Double
Dim l_dbl_SalNoConOpe      As Double
Dim l_dbl_SalConOpe        As Double
Dim l_dbl_IntDevOpe        As Double
Dim l_dbl_IntSusOpe        As Double
Dim l_dbl_IntMorOpe        As Double
Dim l_int_DiaMorOpe        As Integer
Dim l_dbl_TraNoConOpe      As Double
Dim l_dbl_AmNoCoOpe        As Double
Dim l_dbl_AmoConOpe        As Double

Dim l_int_ClaInt           As Integer
Dim l_int_ClaAli           As Integer
Dim l_int_ClaPro           As Integer
Dim l_dbl_ProvGen          As Double
Dim l_dbl_ProvCic          As Double
Dim l_dbl_ProvEsp          As Double
Dim l_dbl_ProvVol          As Double
Dim l_dbl_Aplica           As Double
Dim l_dbl_PrvGenRc         As Double
Dim l_dbl_PrvCicRc         As Double
Dim l_dbl_PrvRip           As Double
Dim l_dbl_CBRFMV           As Double
Dim l_dbl_CBRFMV_RC        As Double
Dim l_int_NumCuo           As Integer
Dim l_int_CuoPag           As Integer
Dim l_int_CuoPen           As Integer
Dim l_str_FlgRef           As String
Dim l_str_FlgJud           As String
Dim l_str_FlgCas           As String
Dim l_str_TipGar           As String
Dim l_str_MonGar           As String
Dim l_dbl_MtoGar           As Double

Private Sub Grabar_Auditoria()
Dim r_str_Proceso    As String
Dim r_str_Tabla      As String
Dim r_str_Descri     As String
Dim r_str_Descri1    As String
Dim r_str_Descri2    As String
Dim r_str_Descri3    As String
Dim r_str_Usuario    As String
Dim r_str_Plataforma As String
Dim r_str_Terminal   As String
Dim r_str_Sucursal   As String

   r_str_Proceso = "OPERACIONES CIERRE"
   r_str_Tabla = "CRE_HIPCIE"
   r_str_Usuario = modgen_g_str_CodUsu
   r_str_Terminal = modgen_g_str_NombPC
   r_str_Plataforma = UCase(App.EXEName)
   r_str_Sucursal = modgen_g_str_CodSuc
   r_str_Descri1 = ""
   r_str_Descri2 = ""
   r_str_Descri3 = ""

   'Verificacion de datos modificados para ser guardados como Auditoria
   If l_dbl_IntCapCtb <> fpd_IntCapit.Text Then
      r_str_Descri = r_str_Descri + "Interes Capitalizado / Contabilidad (Antes: " & Format(l_dbl_IntCapCtb, "#,###,##0.00") & ")  (Nuevo: " & fpd_IntCapit.Text & ")" + Chr(13)
   End If
   If l_dbl_CapAmoCtb <> fpd_CapAmort.Text Then
      r_str_Descri = r_str_Descri + "Capital Amortizado / Contabilidad (Antes: " & Format(l_dbl_CapAmoCtb, "#,###,##0.00") & ")  (Nuevo: " & fpd_CapAmort.Text & ")" + Chr(13)
   End If
   If l_dbl_SalNoConCtb <> fpd_SalNConC.Text Then
      r_str_Descri = r_str_Descri + "Saldo No Concesional / Contabilidad (Antes: " & Format(l_dbl_SalNoConCtb, "#,###,##0.00") & ")  (Nuevo: " & fpd_SalNConC.Text & ")" + Chr(13)
   End If
   If l_dbl_SalConCtb <> fpd_SalConcC.Text Then
      r_str_Descri = r_str_Descri + "Saldo Concesional / Contabilidad (Antes: " & Format(l_dbl_SalConCtb, "#,###,##0.00") & ")  (Nuevo: " & fpd_SalConcC.Text & ")" + Chr(13)
   End If
   If l_dbl_SalPBPCtb <> fpd_PBPPerdi.Text Then
      r_str_Descri = r_str_Descri + "PBP Perdido(TC) / Contabilidad (Antes: " & Format(l_dbl_SalPBPCtb, "#,###,##0.00") & ")  (Nuevo: " & fpd_PBPPerdi.Text & ")" + Chr(13)
   End If
   If l_dbl_IntDevCtb <> fpd_IntDevnC.Text Then
      r_str_Descri = r_str_Descri + "Interes Devengado / Contabilidad (Antes: " & Format(l_dbl_IntDevCtb, "#,###,##0.00") & ")  (Nuevo: " & fpd_IntDevnC.Text & ")" + Chr(13)
   End If
   If l_dbl_IntSusCtb <> fpd_IntSuspC.Text Then
      r_str_Descri = r_str_Descri + "Interes en Suspenso / Contabilidad (Antes: " & Format(l_dbl_IntSusCtb, "#,###,##0.00") & ")  (Nuevo: " & fpd_IntSuspC.Text & ")" + Chr(13)
   End If
   If l_dbl_IntMorCtb <> fpd_IntMoraC.Text Then
      r_str_Descri = r_str_Descri + "Interes Moratorio / Contabilidad (Antes: " & Format(l_dbl_IntMorCtb, "#,###,##0.00") & ")  (Nuevo: " & fpd_IntMoraC.Text & ")" + Chr(13)
   End If
   If l_int_DiaMorCtb <> fpd_DiaMoroC.Text Then
      r_str_Descri = r_str_Descri + "Dias de Morosidad / Contabilidad (Antes: " & Format(l_int_DiaMorCtb, "#,###,##0") & ")  (Nuevo: " & fpd_DiaMoroC.Text & ")" + Chr(13)
   End If
   If l_dbl_CapVenCtb <> fpd_CapVencd.Text Then
      r_str_Descri = r_str_Descri + "Capital Vencido / Contabilidad (Antes: " & Format(l_dbl_CapVenCtb, "#,###,##0.00") & ")  (Nuevo: " & fpd_CapVencd.Text & ")" + Chr(13)
   End If

   'Operaciones
   If l_dbl_CapVigOpe <> fpd_TotCapVi.Text Then
      r_str_Descri = r_str_Descri + "Total Capital Vigente / Operaciones (Antes: " & Format(l_dbl_CapVigOpe, "#,###,##0.00") & ")  (Nuevo: " & fpd_TotCapVi.Text & ")" + Chr(13)
   End If
   If l_dbl_CapVenOpe <> fpd_TotCapVe.Text Then
      r_str_Descri = r_str_Descri + "Total Capital Vencido / Operaciones (Antes: " & Format(l_dbl_CapVenOpe, "#,###,##0.00") & ")  (Nuevo: " & fpd_TotCapVe.Text & ")" + Chr(13)
   End If
   If l_dbl_SalNoConOpe <> fpd_SalNConO.Text Then
      r_str_Descri = r_str_Descri + "Saldo No Concesional / Operaciones (Antes: " & Format(l_dbl_SalNoConOpe, "#,###,##0.00") & ")  (Nuevo: " & fpd_SalNConO.Text & ")" + Chr(13)
   End If
   If l_dbl_SalConOpe <> fpd_SalConcO.Text Then
      r_str_Descri = r_str_Descri + "Saldo Concesional / Operaciones (Antes: " & Format(l_dbl_SalConOpe, "#,###,##0.00") & ")  (Nuevo: " & fpd_SalConcO.Text & ")" + Chr(13)
   End If
   If l_dbl_IntDevOpe <> fpd_IntDevnO.Text Then
      r_str_Descri = r_str_Descri + "Interes Devengado / Operaciones (Antes: " & Format(l_dbl_IntDevOpe, "#,###,##0.00") & ")  (Nuevo: " & fpd_IntDevnO.Text & ")" + Chr(13)
   End If
   If l_dbl_IntSusOpe <> fpd_IntSuspO.Text Then
      r_str_Descri = r_str_Descri + "Interes Suspenso / Operaciones (Antes: " & Format(l_dbl_IntSusOpe, "#,###,##0.00") & ")  (Nuevo: " & fpd_IntSuspO.Text & ")" + Chr(13)
   End If
   If l_dbl_IntMorOpe <> fpd_IntMoraO.Text Then
      r_str_Descri = r_str_Descri + "Interes Moratoria / Operaciones (Antes: " & Format(l_dbl_IntMorOpe, "#,###,##0.00") & ")  (Nuevo: " & fpd_IntMoraO.Text & ")" + Chr(13)
   End If
   If l_int_DiaMorOpe <> fpd_DiaMoroO.Text Then
      r_str_Descri = r_str_Descri + "Dias de Morosidad / Operaciones (Antes: " & Format(l_int_DiaMorOpe, "#,###,##0") & ")  (Nuevo: " & fpd_DiaMoroO.Text & ")" + Chr(13)
   End If
   If l_dbl_TraNoConOpe <> fpd_PrstNCon.Text Then
      r_str_Descri = r_str_Descri + "Tramo No Concesional - Monto Prestamo / Operaciones (Antes: " & Format(l_dbl_TraNoConOpe, "#,###,##0.00") & ")  (Nuevo: " & fpd_PrstNCon.Text & ")" + Chr(13)
   End If
   If l_dbl_AmNoCoOpe <> fpd_AmorNCon.Text Then
      r_str_Descri = r_str_Descri + "Tramo No Concesional - Monto Amortizado / Operaciones (Antes: " & Format(l_dbl_AmNoCoOpe, "#,###,##0.00") & ")  (Nuevo: " & fpd_AmorNCon.Text & ")" + Chr(13)
   End If
   If l_dbl_AmoConOpe <> fpd_AmortCon.Text Then
      r_str_Descri = r_str_Descri + "Tramo Concesional - Monto Amortizado / Operaciones (Antes: " & Format(l_dbl_AmoConOpe, "#,###,##0.00") & ")  (Nuevo: " & fpd_AmortCon.Text & ")" + Chr(13)
   End If
   If l_int_ClaInt <> fpd_ClaCli.Text Then
      r_str_Descri = r_str_Descri + "Clasificacion Interna / Provisiones (Antes: " & Format(l_int_ClaInt, "#,###,##0") & ")  (Nuevo: " & fpd_ClaCli.Text & ")" + Chr(13)
   End If
   If l_int_ClaAli <> fpd_ClaAli.Text Then
      r_str_Descri = r_str_Descri + "Clasificacion Alineada / Provisiones (Antes: " & Format(l_int_ClaAli, "#,###,##0") & ")  (Nuevo: " & fpd_ClaAli.Text & ")" + Chr(13)
   End If
   If l_int_ClaPro <> fpd_ClaPrv.Text Then
      r_str_Descri = r_str_Descri + "Clasificacion Provision / Provisiones (Antes: " & Format(l_int_ClaPro, "#,###,##0") & ")  (Nuevo: " & fpd_ClaPrv.Text & ")" + Chr(13)
   End If
   If l_dbl_ProvGen <> fpd_PrvGen.Text Then
      r_str_Descri = r_str_Descri + "Provision Generica (Antes: " & Format(l_dbl_ProvGen, "#,###,##0.00") & ")  (Nuevo: " & fpd_PrvGen.Text & ")" + Chr(13)
   End If
   If l_dbl_ProvCic <> fpd_PrvCic.Text Then
      r_str_Descri = r_str_Descri + "Provision Ciclica (Antes: " & Format(l_dbl_ProvCic, "#,###,##0.00") & ")  (Nuevo: " & fpd_PrvCic.Text & ")" + Chr(13)
   End If
   If l_dbl_ProvEsp <> fpd_PrvEsp.Text Then
      r_str_Descri = r_str_Descri + "Provision Especifica (Antes: " & Format(l_dbl_ProvEsp, "#,###,##0.00") & ")  (Nuevo: " & fpd_PrvEsp.Text & ")" + Chr(13)
   End If
   If l_dbl_ProvVol <> fpd_PrvVol.Text Then
      r_str_Descri = r_str_Descri + "Provision Voluntaria (Antes: " & Format(l_dbl_ProvVol, "#,###,##0.00") & ")  (Nuevo: " & fpd_PrvVol.Text & ")" + Chr(13)
   End If
   If l_dbl_Aplica <> fpd_Aplica.Text Then
      r_str_Descri = r_str_Descri + "Aplicacion Prociclica (Antes: " & Format(l_dbl_Aplica, "#,###,##0.00") & ")  (Nuevo: " & fpd_Aplica.Text & ")" + Chr(13)
   End If
   If l_dbl_PrvGenRc <> fpd_PrvGenRC.Text Then
      r_str_Descri = r_str_Descri + "Provisión Genérica RC (Antes: " & Format(l_dbl_PrvGenRc, "#,###,##0.00") & ")  (Nuevo: " & fpd_PrvGenRC.Text & ")" + Chr(13)
   End If
   If l_dbl_PrvCicRc <> fpd_PrvCicRC.Text Then
      r_str_Descri = r_str_Descri + "Provisión Prociclica RC (Antes: " & Format(l_dbl_PrvCicRc, "#,###,##0.00") & ")  (Nuevo: " & fpd_PrvCicRC.Text & ")" + Chr(13)
   End If
   If l_dbl_PrvRip <> fpd_PrvRip.Text Then
      r_str_Descri = r_str_Descri + "Provisión Riesgo País (Antes: " & Format(l_dbl_PrvRip, "#,###,##0.00") & ")  (Nuevo: " & fpd_PrvRip.Text & ")" + Chr(13)
   End If
   If l_dbl_CBRFMV <> fpd_CbrFmv.Text Then
      r_str_Descri = r_str_Descri + "COBERTURA FMV (Antes: " & Format(l_dbl_CBRFMV, "#,###,##0.00") & ")  (Nuevo: " & fpd_CbrFmv.Text & ")" + Chr(13)
   End If
   If l_dbl_CBRFMV_RC <> fpd_CbrFmvRC.Text Then
      r_str_Descri = r_str_Descri + "COBERTURA FMV RC (Antes: " & Format(l_dbl_CBRFMV_RC, "#,###,##0.00") & ")  (Nuevo: " & fpd_CbrFmvRC.Text & ")" + Chr(13)
   End If
   If l_int_NumCuo <> fpd_NumCuo.Text Then
      r_str_Descri = r_str_Descri + "Número de Cuota (Antes: " & Format(l_int_NumCuo, "#,###,##0.00") & ")  (Nuevo: " & fpd_NumCuo.Text & ")" + Chr(13)
   End If
   If l_int_CuoPag <> fpd_CuoPag.Text Then
      r_str_Descri = r_str_Descri + "Cuotas Pagadas (Antes: " & Format(l_int_CuoPag, "#,###,##0.00") & ")  (Nuevo: " & fpd_CuoPag.Text & ")" + Chr(13)
   End If
   If l_int_CuoPen <> fpd_CuoPen.Text Then
      r_str_Descri = r_str_Descri + "Cuotas Pendientes (Antes: " & Format(l_int_CuoPen, "#,###,##0.00") & ")  (Nuevo: " & fpd_CuoPen.Text & ")" + Chr(13)
   End If
   If l_str_FlgRef <> cmb_Refina.Text Then
      r_str_Descri = r_str_Descri + "Refinanciamiento (Antes: " & Format(l_str_FlgRef, "#,###,##0.00") & ")  (Nuevo: " & cmb_Refina.Text & ")" + Chr(13)
   End If
   If l_str_FlgJud <> cmb_Judici.Text Then
      r_str_Descri = r_str_Descri + "Judicial (Antes: " & Format(l_str_FlgJud, "#,###,##0.00") & ")  (Nuevo: " & cmb_Judici.Text & ")" + Chr(13)
   End If
   If l_str_FlgCas <> cmb_Castig.Text Then
      r_str_Descri = r_str_Descri + "Castigado (Antes: " & Format(l_str_FlgCas, "#,###,##0.00") & ")  (Nuevo: " & cmb_Castig.Text & ")" + Chr(13)
   End If
   If l_str_TipGar <> Cmb_TipGar.Text Then
      r_str_Descri = r_str_Descri + "Tipo Garantía (Antes: " & Format(l_str_TipGar, "#,###,##0.00") & ")  (Nuevo: " & Cmb_TipGar.Text & ")" + Chr(13)
   End If
   If l_str_MonGar <> cmb_MonGar.Text Then
      r_str_Descri = r_str_Descri + "Moneda Garantía (Antes: " & Format(l_str_MonGar, "#,###,##0.00") & ")  (Nuevo: " & cmb_MonGar.Text & ")" + Chr(13)
   End If
   If l_dbl_MtoGar <> fpd_MtoGar.Text Then
      r_str_Descri = r_str_Descri + "Monto Garantía (Antes: " & Format(l_dbl_MtoGar, "#,###,##0.00") & ")  (Nuevo: " & fpd_MtoGar.Text & ")" + Chr(13)
   End If
   
   'Tomar posiciones de descripcion (cada 2000 caracteres) para guardar en Auditoria
   r_str_Descri1 = Mid(r_str_Descri, 1, 2000)
   r_str_Descri2 = Mid(r_str_Descri, 2001, 4000)

   'Grabacion en Tabla de Auditoria
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "INSERT INTO CRE_AUDIT("
   g_str_Parame = g_str_Parame & "  AUDIT_PROCES, "
   g_str_Parame = g_str_Parame & "  AUDIT_TBLAFE, "
   g_str_Parame = g_str_Parame & "  AUDIT_NUMOPE, "
   g_str_Parame = g_str_Parame & "  AUDIT_PERIOD, "
   g_str_Parame = g_str_Parame & "  AUDIT_FECHA, "
   g_str_Parame = g_str_Parame & "  AUDIT_HORA, "
   g_str_Parame = g_str_Parame & "  AUDIT_DESCR1, "
   g_str_Parame = g_str_Parame & "  AUDIT_DESCR2, "
   g_str_Parame = g_str_Parame & "  AUDIT_DESCR3, "
   g_str_Parame = g_str_Parame & "  SEGUSUCRE, "
   g_str_Parame = g_str_Parame & "  SEGFECCRE, "
   g_str_Parame = g_str_Parame & "  SEGHORCRE, "
   g_str_Parame = g_str_Parame & "  SEGPLTCRE, "
   g_str_Parame = g_str_Parame & "  SEGTERCRE, "
   g_str_Parame = g_str_Parame & "  SEGSUCCRE) "
   g_str_Parame = g_str_Parame & "VALUES ("
   g_str_Parame = g_str_Parame & "'" & r_str_Proceso & "', "
   g_str_Parame = g_str_Parame & "'" & r_str_Tabla & "', "
   g_str_Parame = g_str_Parame & "'" & msk_NumOpe.Text & "', "
   g_str_Parame = g_str_Parame & "" & ipp_PerAno.Text & Format(cmb_PerMes.ListIndex + 1, "00") & ", "
   g_str_Parame = g_str_Parame & "'" & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & "', "
   g_str_Parame = g_str_Parame & "'" & Format(Time, "HHMMSS") & "', "
   g_str_Parame = g_str_Parame & "'" & r_str_Descri1 & "', "
   g_str_Parame = g_str_Parame & "'" & r_str_Descri2 & "', "
   g_str_Parame = g_str_Parame & "'" & r_str_Descri3 & "', "
   g_str_Parame = g_str_Parame & "'" & r_str_Usuario & "', "
   g_str_Parame = g_str_Parame & "'" & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & "', "
   g_str_Parame = g_str_Parame & "'" & Format(Time, "HHMMSS") & "', "
   g_str_Parame = g_str_Parame & "'" & r_str_Plataforma & "', "
   g_str_Parame = g_str_Parame & "'" & r_str_Terminal & "', "
   g_str_Parame = g_str_Parame & "'" & r_str_Sucursal & "' )"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      Exit Sub
   End If
End Sub

Private Sub cmd_Buscar_Click()
Dim r_int_Moneda     As Integer
    
    If Len(Trim(msk_NumOpe.Text)) < 10 Then
        MsgBox "Debe ingresar el Número de Operación.", vbExclamation, modgen_g_con_OpeTra
        msk_NumOpe.Text = ""
        msk_NumOpe.Mask = "###-##-#####"
        Call gs_SetFocus(msk_NumOpe)
        Exit Sub
    End If
    If cmb_PerMes.ListIndex = -1 Then
        MsgBox "Debe seleccionar el Mes.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(cmb_PerMes)
        Exit Sub
    End If
    If ipp_PerAno.Text < 2010 Then
        MsgBox "Debe ingresar el año correcto.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(ipp_PerAno)
        Exit Sub
    End If
    
    moddat_g_str_NumOpe = msk_NumOpe.Text
    Screen.MousePointer = 11
    Me.Enabled = False
    
    Call fs_Buscar_Credito
    
    cmd_Grabar.Enabled = True
    If moddat_g_int_Situac <> 2 Then
      cmd_Grabar.Enabled = False
    End If
    
    'Asignacion de variables usadas en el Proceso de Auditoria.
    l_dbl_IntCapCtb = fpd_IntCapit.Text
    l_dbl_CapAmoCtb = fpd_CapAmort.Text
    l_dbl_SalNoConCtb = fpd_SalNConC.Text
    l_dbl_SalConCtb = fpd_SalConcC.Text
    l_dbl_SalPBPCtb = fpd_PBPPerdi.Text
    l_dbl_IntDevCtb = fpd_IntDevnC.Text
    l_dbl_IntSusCtb = fpd_IntSuspC.Text
    l_dbl_IntMorCtb = fpd_IntMoraC.Text
    l_int_DiaMorCtb = fpd_DiaMoroC.Text
    l_dbl_CapVenCtb = fpd_CapVencd.Text
    l_dbl_CapVigOpe = fpd_TotCapVi.Text
    l_dbl_CapVenOpe = fpd_TotCapVe.Text
    l_dbl_SalNoConOpe = fpd_SalNConO.Text
    l_dbl_SalConOpe = fpd_SalConcO.Text
    l_dbl_IntDevOpe = fpd_IntDevnO.Text
    l_dbl_IntSusOpe = fpd_IntSuspO.Text
    l_dbl_IntMorOpe = fpd_IntMoraO.Text
    l_int_DiaMorOpe = fpd_DiaMoroO.Text
    l_dbl_TraNoConOpe = fpd_PrstNCon.Text
    l_dbl_AmNoCoOpe = fpd_AmorNCon.Text
    l_dbl_AmoConOpe = fpd_AmortCon.Text
    l_int_ClaInt = fpd_ClaCli.Text
    l_int_ClaAli = fpd_ClaAli.Text
    l_int_ClaPro = fpd_ClaPrv.Text
    l_dbl_ProvGen = fpd_PrvGen.Text
    l_dbl_ProvCic = fpd_PrvCic.Text
    l_dbl_ProvEsp = fpd_PrvEsp.Text
    l_dbl_ProvVol = fpd_PrvVol.Text
    l_dbl_Aplica = fpd_Aplica.Text
    
    l_dbl_PrvGenRc = fpd_PrvGenRC.Text
    l_dbl_PrvCicRc = fpd_PrvCicRC.Text
    l_dbl_PrvRip = fpd_PrvRip.Text
    l_dbl_CBRFMV = fpd_CbrFmv.Text
    l_dbl_CBRFMV_RC = fpd_CbrFmvRC.Text
    l_int_NumCuo = fpd_NumCuo.Text
    l_int_CuoPag = fpd_CuoPag.Text
    l_int_CuoPen = fpd_CuoPen.Text
    l_str_FlgRef = Me.cmb_Refina.Text
    l_str_FlgJud = Me.cmb_Judici.Text
    l_str_FlgCas = Me.cmb_Castig.Text
    l_str_TipGar = Me.Cmb_TipGar.Text
    l_str_MonGar = Me.cmb_MonGar.Text
    l_dbl_MtoGar = fpd_MtoGar.Text
   
    Me.Enabled = True
    Screen.MousePointer = 0
    
    If moddat_g_int_CntErr = 2 Then
        msk_NumOpe.Text = ""
        msk_NumOpe.Mask = "###-##-#####"
        Call gs_SetFocus(msk_NumOpe)
    End If
End Sub

Private Sub cmd_Limpia_Click()
    Call gs_LimpiaGrid(grd_Listad)
    Call fs_Validar_Botones(False)
    Call fs_Limpiar
    Call gs_SetFocus(msk_NumOpe)
End Sub

Private Sub cmd_VerPag_Click()
   frm_Ges_CreHip_05.Show 1
End Sub

Private Sub cmd_ImpCro_Click()
   modmip_g_int_OrdAct = 2
   frm_Ges_CreHip_07.l_str_PerAno = CStr(ipp_PerAno.Text)
   frm_Ges_CreHip_07.l_str_PerMes = CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   frm_Ges_CreHip_07.Show 1
End Sub

Private Sub cmd_Grabar_Click()
Dim r_dbl_Auxil1  As Double
Dim r_dbl_Auxil2  As Double
Dim r_dbl_Auxil3  As Double
Dim r_str_Mnsaje  As String
Dim r_bol_MsjeOk  As Boolean
    
   r_bol_MsjeOk = True
   r_str_Mnsaje = ""
   
   r_dbl_Auxil1 = CDbl(Trim(pnl_SldDeud1.Caption))
   r_dbl_Auxil2 = CDbl(Trim(pnl_SldDeud2.Caption))
   If Format(r_dbl_Auxil1, "#######0.00") <> Format(r_dbl_Auxil2, "########0.00") Then
       r_bol_MsjeOk = False
       r_str_Mnsaje = r_str_Mnsaje & "El Saldo de la Deuda 1 es diferente al Saldo de la Deuda 2. "
   End If
   
   r_dbl_Auxil1 = CDbl(Trim(pnl_SldDeud3.Caption))
   r_dbl_Auxil2 = CDbl(Trim(pnl_SldDeud4.Caption))
   If Format(r_dbl_Auxil1, "#######0.00") <> Format(r_dbl_Auxil2, "#######0.00") Then
       r_bol_MsjeOk = False
       r_str_Mnsaje = r_str_Mnsaje & Chr(13) & "El Saldo de la Deuda 3 es diferente al Saldo de la Deuda 4. "
   End If
   
   r_dbl_Auxil1 = CDbl(Trim(pnl_SldDeud1.Caption))
   r_dbl_Auxil2 = CDbl(Trim(pnl_SldDeud3.Caption))
   If Format(r_dbl_Auxil1, "#######0.00") <> Format(r_dbl_Auxil2, "#######0.00") Then
       r_bol_MsjeOk = False
       r_str_Mnsaje = r_str_Mnsaje & Chr(13) & "El Saldo de la Deuda 1 es diferente al Saldo de la Deuda 3."
   End If
   
   r_dbl_Auxil1 = CDbl(fpd_IntDevnC.Text)
   r_dbl_Auxil2 = CDbl(fpd_IntDevnO.Text)
   If Format(r_dbl_Auxil1, "#######0.00") <> Format(r_dbl_Auxil2, "#######0.00") Then
       r_bol_MsjeOk = False
       r_str_Mnsaje = r_str_Mnsaje & Chr(13) & "El Interés Devengado del Padrón de Deudores es diferente al Maestro de Saldos"
   End If
   
   r_dbl_Auxil1 = CDbl(fpd_IntSuspC.Text)
   r_dbl_Auxil2 = CDbl(fpd_IntSuspO.Text)
   If Format(r_dbl_Auxil1, "#######0.00") <> Format(r_dbl_Auxil2, "#######0.00") Then
       r_bol_MsjeOk = False
       r_str_Mnsaje = r_str_Mnsaje & Chr(13) & "El Interés en Suspenso del Padrón de Deudores es diferente al Maestro de Saldos"
   End If
   
   r_dbl_Auxil1 = CDbl(fpd_IntMoraC.Text)
   r_dbl_Auxil2 = CDbl(fpd_IntMoraO.Text)
   If Format(r_dbl_Auxil1, "#######0.00") <> Format(r_dbl_Auxil2, "#######0.00") Then
       r_bol_MsjeOk = False
       r_str_Mnsaje = r_str_Mnsaje & Chr(13) & "El Interés Moratorio del Padrón de Deudores es diferente al Maestro de Saldos"
   End If
   
   r_dbl_Auxil1 = CDbl(fpd_DiaMoroC.Text)
   r_dbl_Auxil2 = CDbl(fpd_DiaMoroO.Text)
   If Format(r_dbl_Auxil1, "#######0.00") <> Format(r_dbl_Auxil2, "#######0.00") Then
       r_bol_MsjeOk = False
       r_str_Mnsaje = r_str_Mnsaje & Chr(13) & "Los Dias de Morosidad del Padrón de Deudores son diferentes al Maestro de Saldos"
   End If
   
   r_dbl_Auxil1 = CDbl(Trim(pnl_SaldNCon.Caption))
   r_dbl_Auxil2 = CDbl(fpd_AmorNCon.Text)
   r_dbl_Auxil3 = CDbl(fpd_PrstNCon.Text)
   If Format(r_dbl_Auxil1 + r_dbl_Auxil2, "#######0.00") <> Format(r_dbl_Auxil3, "#######0.00") Then
       r_bol_MsjeOk = False
       r_str_Mnsaje = r_str_Mnsaje & Chr(13) & "La suma del Saldo del Tramo No Concesional y del Monto Amortizado es diferente al Monto del Préstamo."
   End If
   
   r_dbl_Auxil1 = CDbl(Trim(pnl_SaldoCon.Caption))
   r_dbl_Auxil2 = CDbl(fpd_AmortCon.Text)
   r_dbl_Auxil3 = CDbl(Trim(pnl_PrstaCon.Caption))
   If Format(r_dbl_Auxil1 + r_dbl_Auxil2, "#######0.00") <> Format(r_dbl_Auxil3, "#######0.00") Then
       r_bol_MsjeOk = False
       r_str_Mnsaje = r_str_Mnsaje & Chr(13) & "La suma del Saldo del Tramo Concesional y del Monto Amortizado es diferente al Monto del Préstamo."
   End If
   
   If cmb_Refina.ListIndex = -1 Then
       r_bol_MsjeOk = False
       r_str_Mnsaje = r_str_Mnsaje & Chr(13) & "No ha seleccionado si es o no Refinanciado."
   End If
   
   If cmb_Judici.ListIndex = -1 Then
       r_bol_MsjeOk = False
       r_str_Mnsaje = r_str_Mnsaje & Chr(13) & "No ha seleccionado si es o no Judicial."
   End If
   
   If cmb_Castig.ListIndex = -1 Then
       r_bol_MsjeOk = False
       r_str_Mnsaje = r_str_Mnsaje & Chr(13) & "No ha seleccionado si es o no Castigado."
   End If
   
   If Cmb_TipGar.ListIndex = -1 Then
       r_bol_MsjeOk = False
       r_str_Mnsaje = r_str_Mnsaje & Chr(13) & "No ha seleccionado Tipo de Garantía."
   End If
   
   If cmb_MonGar.ListIndex = -1 Then
       r_bol_MsjeOk = False
       r_str_Mnsaje = r_str_Mnsaje & Chr(13) & "No ha seleccionado Moneda de Garantía."
   End If
   
   If r_bol_MsjeOk Then
       r_str_Mnsaje = "¿Está seguro de grabar los datos?"
   Else
       r_str_Mnsaje = r_str_Mnsaje & Chr(13) & Chr(13) & "Se encontraron incidencias en el proceso ¿Está seguro de proseguir con la grabación de los datos?"
   End If
   
   If MsgBox("" & r_str_Mnsaje, vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Me.Enabled = False
   
   Call fs_Grabar
   Call cmd_Limpia_Click
   Me.Enabled = True
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Me.Caption = modgen_g_str_NomPlt
    
    Call fs_Inicia
    Call fs_Limpiar
    Call fs_Validar_Botones(False)
    Call gs_LimpiaGrid(grd_Listad)
    Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
    Call moddat_gs_Carga_LisIte_Combo(cmb_Refina, 1, "214")
    Call moddat_gs_Carga_LisIte_Combo(cmb_Judici, 1, "214")
    Call moddat_gs_Carga_LisIte_Combo(cmb_Castig, 1, "214")
    Call moddat_gs_Carga_LisIte_Combo(Cmb_TipGar, 1, "241")
    Call moddat_gs_Carga_LisIte_Combo(cmb_MonGar, 1, "204")
    
    Call gs_CentraForm(Me)
    Call gs_SetFocus(msk_NumOpe)
    Screen.MousePointer = 0
End Sub
 
Private Sub fs_Inicia()
    'Inicializando Grid de Datos del Crédito
    grd_Listad.ColWidth(0) = 2900
    grd_Listad.ColWidth(1) = 8150
    grd_Listad.ColAlignment(0) = flexAlignLeftCenter
    grd_Listad.ColAlignment(1) = flexAlignLeftCenter
End Sub

Private Sub fs_Buscar()
   Dim r_str_CodPry     As String
   Dim r_str_NomPry     As String
   Dim r_str_CodBco     As String
   
   'Buscando Información del Crédito
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "(HIPMAE_SITUAC = 2 OR HIPMAE_SITUAC = 6 OR HIPMAE_SITUAC = 9)"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
     
   g_rst_Princi.MoveFirst
   
   'Almacenando en Variables Globales
   moddat_g_int_TipDoc = g_rst_Princi!HIPMAE_TDOCLI
   moddat_g_str_NumDoc = Trim(g_rst_Princi!HIPMAE_NDOCLI)
   moddat_g_str_NumSol = Trim(g_rst_Princi!HIPMAE_NUMSOL)
   moddat_g_str_NumOpe = Trim(g_rst_Princi!HIPMAE_NUMOPE)
   
   'Obteniendo Nombre de Cliente
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   'Obteniendo Nombre y DOI de Cónyuge
   moddat_g_int_CygTDo = g_rst_Princi!HIPMAE_TDOCYG
   moddat_g_str_CygNDo = ""
   moddat_g_str_CygNom = ""
   
   If moddat_g_int_CygTDo > 0 Then
      moddat_g_str_CygNDo = Trim(g_rst_Princi!HIPMAE_NDOCYG & "")
      moddat_g_str_CygNom = moddat_gf_Buscar_NomCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo)
   End If
   
   'Obteniendo Descripción de Producto
   moddat_g_str_CodPrd = Trim(g_rst_Princi!HIPMAE_CODPRD)
   moddat_g_str_NomPrd = moddat_gf_Consulta_Produc(Trim(g_rst_Princi!HIPMAE_CODPRD))

   moddat_g_str_CodSub = Trim(g_rst_Princi!HIPMAE_CODSUB)

   'Obeniendo Modalidad de Producto
   moddat_g_str_CodMod = Trim(g_rst_Princi!HIPMAE_CODMOD)
   moddat_g_str_DesMod = moddat_gf_Buscar_NomMod(Trim(g_rst_Princi!HIPMAE_CODPRD), moddat_g_str_CodMod)
   
   'Ejecutivo de Seguimiento
   moddat_g_str_CodEjeSeg = Trim(g_rst_Princi!HIPMAE_EJESEG & "")
   moddat_g_str_NomEjeSeg = moddat_gf_Buscar_NomEje(moddat_g_str_CodEjeSeg)

   'Consejero Hipotecario
   moddat_g_str_CodConHip = Trim(g_rst_Princi!HIPMAE_CONHIP & "")
   moddat_g_str_NomConHip = moddat_gf_Buscar_NomEje(moddat_g_str_CodConHip)

   'Moneda
   moddat_g_str_Moneda = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPMAE_MONEDA))
   moddat_g_int_TipMon = g_rst_Princi!HIPMAE_MONEDA                           'Moneda Préstamo
   moddat_g_dbl_MtoPre = g_rst_Princi!HIPMAE_MTOPRE                           'Monto Préstamo
   moddat_g_int_CuoPen = g_rst_Princi!HIPMAE_CUOPEN                           'Cuotas Pendientes
   moddat_g_int_TotCuo = g_rst_Princi!HIPMAE_NUMCUO                           'Total de Cuotas
   moddat_g_dbl_SalCap = g_rst_Princi!HIPMAE_SALCAP                           'Saldo Capital
   moddat_g_str_FecApr = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))    'Fecha Desembolso
   
   'Situación de Crédito
   moddat_g_int_Situac = g_rst_Princi!HIPMAE_SITUAC
   moddat_g_str_Situac = moddat_gf_Consulta_ParDes("027", CStr(g_rst_Princi!HIPMAE_SITUAC))
   
   'Obteniendo Información del Inmueble
   Call moddat_gs_Consulta_DatInm(moddat_g_str_NumSol, moddat_g_str_Direcc, moddat_g_str_Distri, r_str_CodPry, r_str_NomPry, r_str_CodBco)
   
   'Cargando en Grid
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.CellFontBold = True
   grd_Listad.Text = "Número de Operación"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontBold = True
   grd_Listad.Text = gf_Formato_NumOpe(g_rst_Princi!HIPMAE_NUMOPE)
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.CellFontBold = True
   grd_Listad.Text = "Situación"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontBold = True
   grd_Listad.Text = moddat_g_str_Situac
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.CellFontBold = True
   grd_Listad.Text = "Cliente"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontBold = True
   grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_TDOCLI) & " - " & Trim(g_rst_Princi!HIPMAE_NDOCLI) & " / " & moddat_g_str_NomCli
   
   If g_rst_Princi!HIPMAE_TDOCYG > 0 Then
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Cónyuge"
      
      grd_Listad.Col = 1
      grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_TDOCYG) & " - " & Trim(g_rst_Princi!HIPMAE_NDOCYG) & " / " & moddat_g_str_CygNom
   End If
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Producto"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_NomPrd & " / " & moddat_gf_Consulta_SubPrd(g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB)
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Moneda Préstamo"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_Moneda
   
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Modalidad"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_DesMod
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Dirección Inmueble"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_Direcc
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Distrito"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_Distri
   
   If g_rst_Princi!HIPMAE_PRYMCS = 1 Or (g_rst_Princi!HIPMAE_PRYMCS = 2 And CInt(g_rst_Princi!HIPMAE_CODMOD) = 2 Or CInt(g_rst_Princi!HIPMAE_CODMOD) = 3) Then
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Proyecto Inmobiliario"
      
      grd_Listad.Col = 1
      grd_Listad.Text = moddat_gf_Consulta_NomPry(g_rst_Princi!HIPMAE_PRYINM & "")
      
      If g_rst_Princi!HIPMAE_PRYMCS = 2 Then
         grd_Listad.Text = grd_Listad.Text & " (" & moddat_gf_Consulta_ParDes("513", r_str_CodBco) & ")"
      End If
   End If
   
   If moddat_g_int_TipMon = 1 Then
      grd_Listad.Rows = grd_Listad.Rows + 2
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Valor Compra Venta"
      
      grd_Listad.Col = 1
      grd_Listad.CellFontName = "Lucida Console"
      grd_Listad.CellFontSize = 8
      grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_CVTSOL, 12, 2)
      
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Aporte Propio"
      
      grd_Listad.Col = 1
      grd_Listad.CellFontName = "Lucida Console"
      grd_Listad.CellFontSize = 8
      grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APOSOL, 12, 2)
   Else
      grd_Listad.Rows = grd_Listad.Rows + 2
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Valor Compra Venta"
      
      grd_Listad.Col = 1
      grd_Listad.CellFontName = "Lucida Console"
      grd_Listad.CellFontSize = 8
      grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_CVTDOL, 12, 2)
   
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Aporte Propio"
      
      grd_Listad.Col = 1
      grd_Listad.CellFontName = "Lucida Console"
      grd_Listad.CellFontSize = 8
      grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APODOL, 12, 2)
   End If
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Monto Desembolsado"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPDES, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Monto Préstamo"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_MTOPRE, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Interés Capitalizado"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_INTCAP, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Total Préstamo"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_TOTPRE, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Fecha Activación"
   
   grd_Listad.Col = 1
   grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECACT))
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Fecha Desembolso"
   
   grd_Listad.Col = 1
   grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))
   
   If g_rst_Princi!HIPMAE_FECESC > 0 Then
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Fecha Firma EE.PP"
      
      grd_Listad.Col = 1
      grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECESC))
   End If
   
   If moddat_g_str_CodPrd <> "002" Then
      grd_Listad.Rows = grd_Listad.Rows + 2
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      
      Select Case moddat_g_str_CodPrd > 0
         Case InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd): grd_Listad.Text = "Nro. Operación Mivivienda"   '"001"
         Case InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd): grd_Listad.Text = "Nro. Operación COFIDE"       '"003"
         Case InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd): grd_Listad.Text = "Nro. Operación COFIDE"      '"004", "007", "009", "010", "013", "014", "015", "016", "017", "018", "019", "020", "021", "022", "023"
      End Select
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!HIPMAE_OPEMVI & "")
      
      If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "003" Then
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0
         grd_Listad.Text = "Nro. Operación Mivivienda"
         
         grd_Listad.Col = 1
         grd_Listad.Text = Trim(g_rst_Princi!HIPMAE_OPEMV1 & "")
      End If
      
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Monto Préstamo (Tramo No Conces.)"
      
      grd_Listad.Col = 1
      grd_Listad.CellFontName = "Lucida Console"
      grd_Listad.CellFontSize = 8
      grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPNCO, 12, 2)
   
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Monto Préstamo (Tramo Conces.)"
      
      grd_Listad.Col = 1
      grd_Listad.CellFontName = "Lucida Console"
      grd_Listad.CellFontSize = 8
      grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPCON, 12, 2)
      
      If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "001" Or moddat_g_str_CodPrd = "003" Then
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0
         grd_Listad.Text = "Tasa de Interés Mivivienda"
      
         grd_Listad.Col = 1
         grd_Listad.Text = Format(g_rst_Princi!HIPMAE_TASMVI, "##0.00") & " %"
      End If
      
      If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "003" Or moddat_g_str_CodPrd = "007" Or moddat_g_str_CodPrd = "009" Or moddat_g_str_CodPrd = "010" Or moddat_g_str_CodPrd = "013" Or moddat_g_str_CodPrd = "014" Or moddat_g_str_CodPrd = "015" Or moddat_g_str_CodPrd = "016" Or moddat_g_str_CodPrd = "017" Or moddat_g_str_CodPrd = "018" Or moddat_g_str_CodPrd = "019" Or moddat_g_str_CodPrd = "020" Or moddat_g_str_CodPrd = "021" Or moddat_g_str_CodPrd = "022" Or moddat_g_str_CodPrd = "023" Then
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0
         grd_Listad.Text = "Tasa de Interés COFIDE"
      
         grd_Listad.Col = 1
         grd_Listad.Text = Format(g_rst_Princi!HIPMAE_TASCOF, "##0.00") & " %"
         
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0
         grd_Listad.Text = "Tasa de Comisión COFIDE"
         
         grd_Listad.Col = 1
         grd_Listad.Text = Format(g_rst_Princi!HIPMAE_COMCOF, "##0.00") & " %"
      End If
   End If
   
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Plazo"
   
   grd_Listad.Col = 1
   grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_PLAANO) & " Años"
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Tasa de Interés"
   
   grd_Listad.Col = 1
   grd_Listad.Text = Format(g_rst_Princi!HIPMAE_TASINT, "##0.00") & " %"
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Nro. de Cuotas"
   
   grd_Listad.Col = 1
   grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_NUMCUO)
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Período de Gracia"
   
   grd_Listad.Col = 1
   grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_PERGRA) & " Meses"
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Compañía de Seguros"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!HIPMAE_SEGPRE & "")
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Tipo de Seguro Desg."
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!HIPMAE_SEGPRE, g_rst_Princi!HIPMAE_TIPSEG)
      
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Tipo Garantía"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_gf_Consulta_ParDes("241", CStr(g_rst_Princi!HIPMAE_TIPGAR))
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Monto Garantía"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!HIPMAE_MONGAR)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_MTOGAR, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Saldo Capital"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCAP + g_rst_Princi!HIPMAE_SALCON, 12, 2)
      
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Cuotas Pendientes de Pago"
   
   grd_Listad.Col = 1
   grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_CUOPEN)
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Días de Atraso"
   
   grd_Listad.Col = 1
   grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_DIAMOR) & " Días"
   
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Saldo Capital (Tramo No Conces.)"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCAP, 12, 2)
      
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Saldo Capital (Tramo Conces.)"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCON, 12, 2)
 
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Consejero Hipotecario"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_NomConHip
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Ejecutivo de Seguimiento"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_NomEjeSeg
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_Listad)
End Sub

Private Function fs_Obtiene_FechaPago(ByVal p_NumOpe As String, ByVal p_TipCro As Integer, ByVal p_FecDes As String) As String
   Dim r_rst_Temp    As Recordset
   fs_Obtiene_FechaPago = p_FecDes
   
   g_str_Parame = "SELECT HIPCUO_FECVCT FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & p_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = " & p_TipCro & " "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_FECVCT DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Temp, 3) Then
       Exit Function
   End If
   
   If Not (r_rst_Temp.BOF And r_rst_Temp.EOF) Then
      r_rst_Temp.MoveFirst
      fs_Obtiene_FechaPago = r_rst_Temp!HIPCUO_FECVCT
   End If
   
   r_rst_Temp.Close
   Set r_rst_Temp = Nothing
End Function

Private Sub fs_Grabar()
    g_str_Parame = ""
    g_str_Parame = g_str_Parame & "USP_CUADRE_SALPAD("
    g_str_Parame = g_str_Parame & "'" & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & "', "
    g_str_Parame = g_str_Parame & "'" & ipp_PerAno & "', "
    g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
    g_str_Parame = g_str_Parame & CDbl(fpd_IntCapit.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_CapAmort.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_SalNConC.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_SalConcC.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_IntDevnC.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_IntSuspC.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_IntMoraC.Text) & ", "
    g_str_Parame = g_str_Parame & CInt(fpd_DiaMoroC.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_CapVencd.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_TotCapVi.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_TotCapVe.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_SalNConO.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_SalConcO.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_IntDevnO.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_IntSuspO.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_IntMoraO.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_DiaMoroO.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_AmorNCon.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_AmortCon.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_PrstNCon.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_PBPPerdi.Text) & ", "
    g_str_Parame = g_str_Parame & CInt(fpd_ClaCli.Text) & ", "
    g_str_Parame = g_str_Parame & CInt(fpd_ClaAli.Text) & ", "
    g_str_Parame = g_str_Parame & CInt(fpd_ClaPrv.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_PrvGen.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_PrvCic.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_PrvEsp.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_PrvVol.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_Aplica.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_PrvGenRC.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_PrvCicRC.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_PrvRip.Text) & ", "
    
    g_str_Parame = g_str_Parame & CDbl(fpd_CbrFmv.Text) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_CbrFmvRC.Text) & ", "
    g_str_Parame = g_str_Parame & CInt(fpd_NumCuo.Text) & ", "
    g_str_Parame = g_str_Parame & CInt(fpd_CuoPag.Text) & ", "
    g_str_Parame = g_str_Parame & CInt(fpd_CuoPen.Text) & ", "
    g_str_Parame = g_str_Parame & IIf(cmb_Refina.ItemData(cmb_Refina.ListIndex) = 2, 0, cmb_Refina.ItemData(cmb_Refina.ListIndex)) & ", "
    g_str_Parame = g_str_Parame & IIf(cmb_Judici.ItemData(cmb_Judici.ListIndex) = 2, 0, cmb_Judici.ItemData(cmb_Judici.ListIndex)) & ", "
    g_str_Parame = g_str_Parame & IIf(cmb_Castig.ItemData(cmb_Castig.ListIndex) = 2, 0, cmb_Castig.ItemData(cmb_Castig.ListIndex)) & ", "
    g_str_Parame = g_str_Parame & Cmb_TipGar.ItemData(Cmb_TipGar.ListIndex) & ", "
    g_str_Parame = g_str_Parame & cmb_MonGar.ItemData(cmb_MonGar.ListIndex) & ", "
    g_str_Parame = g_str_Parame & CDbl(fpd_MtoGar.Text) & ", "
    
    'Datos de Auditoria
    g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                  'Código Usuario
    g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                  'Nombre Terminal
    g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                   'Nombre Ejecutable
    g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                  'Código Sucursal
     
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
        Exit Sub
    Else
        Call Grabar_Auditoria
        MsgBox "El proceso se grabó exitosamente.", vbInformation, modgen_g_str_NomPlt
    End If
End Sub

Private Sub fs_Limpiar()
   msk_NumOpe.Text = ""
   cmb_PerMes.ListIndex = -1
   ipp_PerAno.Text = Year(date)
   'fpd_SalConcC.Enabled = True
   'fpd_SalConcO.Enabled = True
   'fpd_AmortCon.Enabled = True
   Call fs_Limpiar_DatCre
   
   'Limpieza de variables usadas para el proceso de Auditoria.
   l_dbl_IntCapCtb = 0
   l_dbl_CapAmoCtb = 0
   l_dbl_SalNoConCtb = 0
   l_dbl_SalConCtb = 0
   l_dbl_SalPBPCtb = 0
   l_dbl_IntDevCtb = 0
   l_dbl_IntSusCtb = 0
   l_dbl_IntMorCtb = 0
   l_int_DiaMorCtb = 0
   l_dbl_CapVenCtb = 0
   l_dbl_CapVigOpe = 0
   l_dbl_CapVenOpe = 0
   l_dbl_SalNoConOpe = 0
   l_dbl_SalConOpe = 0
   l_dbl_IntDevOpe = 0
   l_dbl_IntSusOpe = 0
   l_dbl_IntMorOpe = 0
   l_int_DiaMorOpe = 0
   l_dbl_TraNoConOpe = 0
   l_dbl_AmNoCoOpe = 0
   l_dbl_AmoConOpe = 0
   l_int_ClaInt = 0
   l_int_ClaAli = 0
   l_int_ClaPro = 0
   l_dbl_ProvGen = 0
   l_dbl_ProvCic = 0
   l_dbl_ProvEsp = 0
   l_dbl_ProvVol = 0
   l_dbl_Aplica = 0
End Sub

Private Sub fs_Limpiar_DatCre()
    pnl_CapDesmb.Caption = "0.00 "
    fpd_IntCapit.Text = 0
    fpd_CapAmort.Text = 0
    pnl_SldDeud1.Caption = "0.00 "
    fpd_SalNConC.Text = 0
    fpd_SalConcC.Text = 0
    fpd_PBPPerdi.Text = 0
    pnl_SldDeud2.Caption = "0.00 "
    fpd_IntDevnC.Text = 0
    fpd_IntSuspC.Text = 0
    fpd_IntMoraC.Text = 0
    fpd_DiaMoroC.Text = 0
    fpd_CapVencd.Text = 0
    fpd_TotCapVi.Text = 0
    fpd_TotCapVe.Text = 0
    pnl_SldDeud3.Caption = "0.00 "
    fpd_SalNConO.Text = 0
    fpd_SalConcO.Text = 0
    pnl_SldDeud4.Caption = "0.00 "
    fpd_IntDevnO.Text = 0
    fpd_IntSuspO.Text = 0
    fpd_IntMoraO.Text = 0
    fpd_DiaMoroO.Text = 0
    
    fpd_PrstNCon.Text = 0
    pnl_PrstaCon.Caption = "0.00 "
    pnl_SaldNCon.Caption = "0.00 "
    pnl_SaldoCon.Caption = "0.00 "
    fpd_AmorNCon.Text = 0
    fpd_AmortCon.Text = 0
    fpd_ClaCli.Text = 0
    fpd_ClaAli.Text = 0
    fpd_ClaPrv.Text = 0
    fpd_PrvGen.Text = 0
    fpd_PrvCic.Text = 0
    fpd_PrvEsp.Text = 0
    fpd_PrvVol.Text = 0
    fpd_Aplica.Text = 0
    
    fpd_PrvGenRC.Text = "0.00 "
    fpd_PrvCicRC.Text = "0.00 "
    fpd_PrvRip.Text = "0.00 "
    fpd_CbrFmv.Text = "0.00 "
    fpd_CbrFmvRC.Text = "0.00 "
    fpd_NumCuo.Text = 0
    fpd_CuoPag.Text = 0
    fpd_CuoPen.Text = 0
    cmb_Refina.ListIndex = -1
    cmb_Judici.ListIndex = -1
    cmb_Castig.ListIndex = -1
    Cmb_TipGar.ListIndex = -1
    cmb_MonGar.ListIndex = -1
    fpd_MtoGar.Text = "0.00 "
End Sub

Private Sub fs_Buscar_Credito_ant()
Dim r_str_CodPry     As String
Dim r_str_NomPry     As String
Dim r_str_CodBco     As String
    
   moddat_g_int_CntErr = 1
   Call gs_LimpiaGrid(grd_Listad)
   Call fs_Limpiar_DatCre
   fpd_SalConcC.Enabled = True
   fpd_SalConcO.Enabled = True
   fpd_AmortCon.Enabled = True
   
   'Buscando Información del Crédito
   g_str_Parame = " "
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPMAE_SITUAC in (2,6,9)"
    
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      moddat_g_int_CntErr = 2
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No existe ninguna Operación registrada con ese Número. ", vbExclamation, modgen_g_con_OpeTra
      Call gs_LimpiaGrid(grd_Listad)
      Call fs_Limpiar
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      moddat_g_int_CntErr = 2
      Exit Sub
   End If
   
   If g_rst_Princi!HIPMAE_SITUAC = 6 Then
      MsgBox "Operación se encuentra transferida.", vbExclamation, modgen_g_con_OpeTra
   End If
   
   If g_rst_Princi!HIPMAE_SITUAC = 9 Then
      MsgBox "Operación se encuentra cancelada.", vbExclamation, modgen_g_con_OpeTra
   End If
   
   Call fs_Validar_Botones(True)
   g_rst_Princi.MoveFirst
   
   'Almacenando en Variables Globales
   moddat_g_int_TipDoc = g_rst_Princi!HIPMAE_TDOCLI
   moddat_g_str_NumDoc = Trim(g_rst_Princi!HIPMAE_NDOCLI)
   moddat_g_str_NumSol = Trim(g_rst_Princi!HIPMAE_NUMSOL)
   moddat_g_str_NumOpe = Trim(g_rst_Princi!HIPMAE_NUMOPE)
   
   'Obteniendo Nombre de Cliente
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   'Obteniendo Nombre y DOI de Cónyuge
   moddat_g_int_CygTDo = g_rst_Princi!HIPMAE_TDOCYG
   moddat_g_str_CygNDo = ""
   moddat_g_str_CygNom = ""
   
   If moddat_g_int_CygTDo > 0 Then
      moddat_g_str_CygNDo = Trim(g_rst_Princi!HIPMAE_NDOCYG & "")
      moddat_g_str_CygNom = moddat_gf_Buscar_NomCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo)
   End If
   
   'Obteniendo Descripción de Producto
   moddat_g_str_CodPrd = Trim(g_rst_Princi!HIPMAE_CODPRD)
   moddat_g_str_NomPrd = moddat_gf_Consulta_Produc(Trim(g_rst_Princi!HIPMAE_CODPRD))
   moddat_g_str_CodSub = Trim(g_rst_Princi!HIPMAE_CODSUB)
   
   'Obeniendo Modalidad de Producto
   moddat_g_str_CodMod = Trim(g_rst_Princi!HIPMAE_CODMOD)
   moddat_g_str_DesMod = moddat_gf_Buscar_NomMod(Trim(g_rst_Princi!HIPMAE_CODPRD), moddat_g_str_CodMod)
   
   'Ejecutivo de Seguimiento
   moddat_g_str_CodEjeSeg = Trim(g_rst_Princi!HIPMAE_EJESEG & "")
   moddat_g_str_NomEjeSeg = moddat_gf_Buscar_NomEje(moddat_g_str_CodEjeSeg)
   
   'Consejero Hipotecario
   moddat_g_str_CodConHip = Trim(g_rst_Princi!HIPMAE_CONHIP & "")
   moddat_g_str_NomConHip = moddat_gf_Buscar_NomEje(moddat_g_str_CodConHip)
   
   'Moneda
   moddat_g_str_Moneda = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPMAE_MONEDA))
   moddat_g_int_TipMon = g_rst_Princi!HIPMAE_MONEDA
   moddat_g_dbl_MtoPre = g_rst_Princi!HIPMAE_MTOPRE                  'Monto Préstamo
   moddat_g_int_CuoPen = g_rst_Princi!HIPMAE_CUOPEN                  'Cuotas Pendientes
   moddat_g_int_TotCuo = g_rst_Princi!HIPMAE_NUMCUO                  'Total de Cuotas
   moddat_g_dbl_SalCap = g_rst_Princi!HIPMAE_SALCAP                  'Saldo Capital
   moddat_g_str_FecApr = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))
   
   'Situación de Crédito
   moddat_g_int_Situac = g_rst_Princi!HIPMAE_SITUAC
   moddat_g_str_Situac = moddat_gf_Consulta_ParDes("027", CStr(g_rst_Princi!HIPMAE_SITUAC))
   
   'Obteniendo Información del Inmueble
   Call moddat_gs_Consulta_DatInm(moddat_g_str_NumSol, moddat_g_str_Direcc, moddat_g_str_Distri, r_str_CodPry, r_str_NomPry, r_str_CodBco)
   
   'Cargando en Grid
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.CellFontBold = True
   grd_Listad.Text = "Número de Operación"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontBold = True
   grd_Listad.Text = gf_Formato_NumOpe(g_rst_Princi!HIPMAE_NUMOPE)
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.CellFontBold = True
   grd_Listad.Text = "Situación"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontBold = True
   grd_Listad.Text = moddat_g_str_Situac
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.CellFontBold = True
   grd_Listad.Text = "Cliente"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontBold = True
   grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_TDOCLI) & " - " & Trim(g_rst_Princi!HIPMAE_NDOCLI) & " / " & moddat_g_str_NomCli
   
   If g_rst_Princi!HIPMAE_TDOCYG > 0 Then
       grd_Listad.Rows = grd_Listad.Rows + 1
       grd_Listad.Row = grd_Listad.Rows - 1
       grd_Listad.Col = 0
       grd_Listad.Text = "Cónyuge"
       
       grd_Listad.Col = 1
       grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_TDOCYG) & " - " & Trim(g_rst_Princi!HIPMAE_NDOCYG) & " / " & moddat_g_str_CygNom
   End If
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Producto"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_NomPrd & " / " & moddat_gf_Consulta_SubPrd(g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB)
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Moneda Préstamo"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_Moneda
   
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Modalidad"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_DesMod
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Dirección Inmueble"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_Direcc
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Distrito"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_Distri
   
   If g_rst_Princi!HIPMAE_PRYMCS = 1 Or (g_rst_Princi!HIPMAE_PRYMCS = 2 And CInt(g_rst_Princi!HIPMAE_CODMOD) = 2 Or CInt(g_rst_Princi!HIPMAE_CODMOD) = 3) Then
       grd_Listad.Rows = grd_Listad.Rows + 1
       grd_Listad.Row = grd_Listad.Rows - 1
       grd_Listad.Col = 0
       grd_Listad.Text = "Proyecto Inmobiliario"
       
       grd_Listad.Col = 1
       grd_Listad.Text = moddat_gf_Consulta_NomPry(g_rst_Princi!HIPMAE_PRYINM & "")
       
       If g_rst_Princi!HIPMAE_PRYMCS = 2 Then
           grd_Listad.Text = grd_Listad.Text & " (" & moddat_gf_Consulta_ParDes("513", r_str_CodBco) & ")"
       End If
   End If
   
   If moddat_g_int_TipMon = 1 Then
       grd_Listad.Rows = grd_Listad.Rows + 2
       grd_Listad.Row = grd_Listad.Rows - 1
       grd_Listad.Col = 0
       grd_Listad.Text = "Valor Compra Venta"
       
       grd_Listad.Col = 1
       grd_Listad.CellFontName = "Lucida Console"
       grd_Listad.CellFontSize = 8
       grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_CVTSOL, 12, 2)
       
       grd_Listad.Rows = grd_Listad.Rows + 1
       grd_Listad.Row = grd_Listad.Rows - 1
       grd_Listad.Col = 0
       grd_Listad.Text = "Aporte Propio"
       
       grd_Listad.Col = 1
       grd_Listad.CellFontName = "Lucida Console"
       grd_Listad.CellFontSize = 8
       grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APOSOL, 12, 2)
   Else
       grd_Listad.Rows = grd_Listad.Rows + 2
       grd_Listad.Row = grd_Listad.Rows - 1
       grd_Listad.Col = 0
       grd_Listad.Text = "Valor Compra Venta"
       
      grd_Listad.Col = 1
      grd_Listad.CellFontName = "Lucida Console"
      grd_Listad.CellFontSize = 8
      grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_CVTDOL, 12, 2)
   
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Aporte Propio"
      
      grd_Listad.Col = 1
      grd_Listad.CellFontName = "Lucida Console"
      grd_Listad.CellFontSize = 8
      grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APODOL, 12, 2)
   End If
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Monto Desembolsado"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPDES, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Monto Préstamo"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_MTOPRE, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Interés Capitalizado"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_INTCAP, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Total Préstamo"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_TOTPRE, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Fecha Activación"
   
   grd_Listad.Col = 1
   grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECACT))
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Fecha Desembolso"
   
   grd_Listad.Col = 1
   grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))
   
   If g_rst_Princi!HIPMAE_FECESC > 0 Then
       grd_Listad.Rows = grd_Listad.Rows + 1
       grd_Listad.Row = grd_Listad.Rows - 1
       grd_Listad.Col = 0
       grd_Listad.Text = "Fecha Firma EE.PP"
       
       grd_Listad.Col = 1
       grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECESC))
   End If
   
   If moddat_g_str_CodPrd <> "002" Then
       grd_Listad.Rows = grd_Listad.Rows + 2
       grd_Listad.Row = grd_Listad.Rows - 1
       grd_Listad.Col = 0
       
       Select Case moddat_g_str_CodPrd > 0
           Case InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd): grd_Listad.Text = "Nro. Operación Mivivienda" '"001"
           Case InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd): grd_Listad.Text = "Nro. Operación COFIDE"     '"003"
           Case InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd): grd_Listad.Text = "Nro. Operación COFIDE"    '"004", "007", "009", "010", "013", "014", "015", "016", "017", "018", "019", "020", "021", "022", "023"
       End Select
     
       grd_Listad.Col = 1
       grd_Listad.Text = Trim(g_rst_Princi!HIPMAE_OPEMVI & "")
       
       If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "003" Then
           grd_Listad.Rows = grd_Listad.Rows + 1
           grd_Listad.Row = grd_Listad.Rows - 1
           grd_Listad.Col = 0
           grd_Listad.Text = "Nro. Operación Mivivienda"
           
           grd_Listad.Col = 1
           grd_Listad.Text = Trim(g_rst_Princi!HIPMAE_OPEMV1 & "")
       End If
     
       grd_Listad.Rows = grd_Listad.Rows + 1
       grd_Listad.Row = grd_Listad.Rows - 1
       grd_Listad.Col = 0
       grd_Listad.Text = "Monto Préstamo (Tramo No Conces.)"
       
       grd_Listad.Col = 1
       grd_Listad.CellFontName = "Lucida Console"
       grd_Listad.CellFontSize = 8
       grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPNCO, 12, 2)
       
       grd_Listad.Rows = grd_Listad.Rows + 1
       grd_Listad.Row = grd_Listad.Rows - 1
       grd_Listad.Col = 0
       grd_Listad.Text = "Monto Préstamo (Tramo Conces.)"
       
       grd_Listad.Col = 1
       grd_Listad.CellFontName = "Lucida Console"
       grd_Listad.CellFontSize = 8
       grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPCON, 12, 2)
     
       If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "001" Or moddat_g_str_CodPrd = "003" Then
          grd_Listad.Rows = grd_Listad.Rows + 1
          grd_Listad.Row = grd_Listad.Rows - 1
          grd_Listad.Col = 0
          grd_Listad.Text = "Tasa de Interés Mivivienda"
       
          grd_Listad.Col = 1
          grd_Listad.Text = Format(g_rst_Princi!HIPMAE_TASMVI, "##0.00") & " %"
       End If
     
       If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "003" Or moddat_g_str_CodPrd = "007" Or moddat_g_str_CodPrd = "009" Or moddat_g_str_CodPrd = "010" Or moddat_g_str_CodPrd = "013" Or moddat_g_str_CodPrd = "014" Or moddat_g_str_CodPrd = "015" Or moddat_g_str_CodPrd = "016" Or moddat_g_str_CodPrd = "017" Or moddat_g_str_CodPrd = "018" Or moddat_g_str_CodPrd = "019" Or moddat_g_str_CodPrd = "020" Or moddat_g_str_CodPrd = "021" Or moddat_g_str_CodPrd = "022" Or moddat_g_str_CodPrd = "023" Then
           grd_Listad.Rows = grd_Listad.Rows + 1
           grd_Listad.Row = grd_Listad.Rows - 1
           grd_Listad.Col = 0
           grd_Listad.Text = "Tasa de Interés COFIDE"
           
           grd_Listad.Col = 1
           grd_Listad.Text = Format(g_rst_Princi!HIPMAE_TASCOF, "##0.00") & " %"
           
           grd_Listad.Rows = grd_Listad.Rows + 1
           grd_Listad.Row = grd_Listad.Rows - 1
           grd_Listad.Col = 0
           grd_Listad.Text = "Tasa de Comisión COFIDE"
           
           grd_Listad.Col = 1
           grd_Listad.Text = Format(g_rst_Princi!HIPMAE_COMCOF, "##0.00") & " %"
       End If
   End If
   
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Plazo"
   
   grd_Listad.Col = 1
   grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_PLAANO) & " Años"
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Tasa de Interés"
   
   grd_Listad.Col = 1
   grd_Listad.Text = Format(g_rst_Princi!HIPMAE_TASINT, "##0.00") & " %"
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Nro. de Cuotas"
   
   grd_Listad.Col = 1
   grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_NUMCUO)
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Período de Gracia"
   
   grd_Listad.Col = 1
   grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_PERGRA) & " Meses"
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Compañía de Seguros"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!HIPMAE_SEGPRE & "")
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Tipo de Seguro Desg."
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!HIPMAE_SEGPRE, g_rst_Princi!HIPMAE_TIPSEG)
   
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Tipo Garantía"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_gf_Consulta_ParDes("241", CStr(g_rst_Princi!HIPMAE_TIPGAR))
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Monto Garantía"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!HIPMAE_MONGAR)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_MTOGAR, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Saldo Capital"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCAP + g_rst_Princi!HIPMAE_SALCON, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Cuotas Pendientes de Pago"
   
   grd_Listad.Col = 1
   grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_CUOPEN)
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Días de Atraso"
   
   grd_Listad.Col = 1
   grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_DIAMOR) & " Días"
   
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Saldo Capital (Tramo No Conces.)"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCAP, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Saldo Capital (Tramo Conces.)"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCON, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Consejero Hipotecario"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_NomConHip
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Ejecutivo de Seguimiento"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_NomEjeSeg
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_Listad)
   
   'Datos del Padron
   Call fs_Buscar_DatCred
End Sub

Private Sub fs_Buscar_Credito()
Dim r_str_CodPry     As String
Dim r_str_NomPry     As String
Dim r_str_CodBco     As String
    
   moddat_g_int_CntErr = 1
   
   Call gs_LimpiaGrid(grd_Listad)
   Call fs_Limpiar_DatCre
   fpd_SalConcC.Enabled = True
   fpd_SalConcO.Enabled = True
   fpd_AmortCon.Enabled = True
   
   'Buscando Información del Crédito
   g_str_Parame = " "
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_HIPMAE  "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND HIPMAE_SITUAC in (2,6,9)"
    
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      moddat_g_int_CntErr = 2
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No existe ninguna Operación registrada con ese Número. ", vbExclamation, modgen_g_con_OpeTra
      Call gs_LimpiaGrid(grd_Listad)
      Call fs_Limpiar
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      moddat_g_int_CntErr = 2
      Exit Sub
   End If
   
   If g_rst_Princi!HIPMAE_SITUAC = 6 Then
      MsgBox "Operación se encuentra transferida.", vbExclamation, modgen_g_con_OpeTra
   End If
   
   If g_rst_Princi!HIPMAE_SITUAC = 9 Then
      MsgBox "Operación se encuentra cancelada.", vbExclamation, modgen_g_con_OpeTra
   End If
   
   Call fs_Validar_Botones(True)
      
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Información del Crédito
   Call modmip_gs_DatNumOpe(moddat_g_str_NumOpe, grd_Listad)
   
   'Datos del Padron
   Call fs_Buscar_DatCred
End Sub

Private Sub fs_Buscar_DatCred()
   'DATOS DE CONTABILIDAD
   g_str_Parame = " "
   g_str_Parame = g_str_Parame & " SELECT CAPITAL_DESEMBOLSADO, CAPITAL_INTERES, CAPITAL_AMORTIZADO, SALDO_NO_CONCESIONAL,"
   g_str_Parame = g_str_Parame & "        SALDO_CONCESIONAL, INTERES, INTERES_COMP, INTERES_MOR, DIAS_MOROSIDAD, CAPITAL_VENCIDO "
   g_str_Parame = g_str_Parame & "  FROM  CREDITO_CIERRE_FINMES"
   g_str_Parame = g_str_Parame & "  WHERE MES = '" & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & "'"
   g_str_Parame = g_str_Parame & "    AND ANO = '" & ipp_PerAno.Text & "'"
   g_str_Parame = g_str_Parame & "    AND CREDITO = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      MsgBox "No se encontró datos del crédito del mes de " & cmb_PerMes.Text & ".", vbExclamation, modgen_g_con_OpeTra
      Call fs_Limpiar_DatCre
      Call fs_Validar_Botones(False)
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      Exit Sub
   End If
    
   g_rst_Listas.MoveFirst
   pnl_CapDesmb.Caption = Format(CDbl(g_rst_Listas!CAPITAL_DESEMBOLSADO), "######,#00.00") & " "
   fpd_IntCapit.Text = Format(CDbl(g_rst_Listas!CAPITAL_INTERES), "######,#00.00") & " "
   fpd_CapAmort.Text = CDbl(g_rst_Listas!CAPITAL_AMORTIZADO)
   Call fpd_CapAmort_Change
   
   fpd_SalNConC.Text = CDbl(g_rst_Listas!SALDO_NO_CONCESIONAL)
   fpd_SalConcC.Text = CDbl(g_rst_Listas!SALDO_CONCESIONAL)
   fpd_IntDevnC.Text = CDbl(g_rst_Listas!INTERES)
   fpd_IntSuspC.Text = CDbl(g_rst_Listas!INTERES_COMP)
   fpd_IntMoraC.Text = CDbl(g_rst_Listas!INTERES_MOR)
   fpd_DiaMoroC.Text = CDbl(g_rst_Listas!DIAS_MOROSIDAD)
   fpd_CapVencd.Text = CDbl(g_rst_Listas!CAPITAL_VENCIDO)
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
   
   'DATOS DE OPERACION
   g_str_Parame = " "
   g_str_Parame = g_str_Parame & " SELECT HIPCIE_CAPVIG, HIPCIE_CAPVEN, HIPCIE_SALCAP  , HIPCIE_SALCON      , HIPCIE_ACUDVG, HIPCIE_ACUDVC, "
   g_str_Parame = g_str_Parame & "        HIPCIE_INTMOR, HIPCIE_DIAMOR, HIPCIE_PRENCO  , HIPCIE_CPTNAM      , HIPCIE_PRECON, HIPCIE_CPTCAM, "
   g_str_Parame = g_str_Parame & "        HIPCIE_PERPBP, HIPCIE_TOTPRE, SUBSTR(HIPCIE_NUMOPE,1,3) AS CODPROD, HIPCIE_INTCAP, "
   g_str_Parame = g_str_Parame & "        HIPCIE_CLACLI, HIPCIE_CLAALI, HIPCIE_CLAPRV  , HIPCIE_PRVGEN      , HIPCIE_PRVESP, HIPCIE_PRVCIC, "
   g_str_Parame = g_str_Parame & "        HIPCIE_PRVVOL, HIPCIE_APLCIC, HIPCIE_CBRFMV  , HIPCIE_CBRFMV_RC   , HIPCIE_PRVGEN_RC, HIPCIE_PRVCIC_RC, "
   g_str_Parame = g_str_Parame & "        HIPCIE_TIPGAR, HIPCIE_PRVRIP, HIPCIE_NUMCUO, HIPCIE_CUOPAG, HIPCIE_CUOPEN, "
   g_str_Parame = g_str_Parame & "        CASE WHEN HIPCIE_FLGREF = 0 THEN 2 ELSE HIPCIE_FLGREF END AS  FLGREF,  "
   g_str_Parame = g_str_Parame & "        CASE WHEN HIPCIE_FLGJUD = 0 THEN 2 ELSE HIPCIE_FLGJUD END AS  FLGJUD,  "
   g_str_Parame = g_str_Parame & "        CASE WHEN HIPCIE_FLGCAS = 0 THEN 2 ELSE HIPCIE_FLGCAS END AS  FLGCAS,  "
   g_str_Parame = g_str_Parame & "        HIPCIE_MONGAR, HIPCIE_MTOGAR "
   g_str_Parame = g_str_Parame & "   FROM CRE_HIPCIE"
   g_str_Parame = g_str_Parame & "  WHERE HIPCIE_PERMES = '" & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & "'"
   g_str_Parame = g_str_Parame & "    AND HIPCIE_PERANO = '" & ipp_PerAno.Text & "'"
   g_str_Parame = g_str_Parame & "    AND HIPCIE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Sub
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Sub
   End If
   
   g_rst_GenAux.MoveFirst
   fpd_TotCapVi.Text = CDbl(g_rst_GenAux!HIPCIE_CAPVIG)
   fpd_TotCapVe.Text = CDbl(g_rst_GenAux!HIPCIE_CAPVEN)
   Call fpd_TotCapVe_Change
   
   fpd_SalNConO.Text = CDbl(g_rst_GenAux!HIPCIE_SALCAP)
   fpd_SalConcO.Text = CDbl(g_rst_GenAux!HIPCIE_SALCON)
   Call fpd_SalConcO_Change
   
   fpd_IntDevnO.Text = CDbl(g_rst_GenAux!HIPCIE_ACUDVG)
   fpd_IntSuspO.Text = CDbl(g_rst_GenAux!HIPCIE_ACUDVC)
   fpd_IntMoraO.Text = CDbl(g_rst_GenAux!HIPCIE_INTMOR)
   fpd_DiaMoroO.Text = CDbl(g_rst_GenAux!HIPCIE_DIAMOR)
   
   If (CInt(g_rst_GenAux!CODPROD) = 2 Or CInt(g_rst_GenAux!CODPROD) = 11) Then
       fpd_PrstNCon.Text = Format(CDbl(g_rst_GenAux!HIPCIE_TOTPRE), "###,###,#00.00") '& " "
       fpd_SalConcC.Enabled = False
       fpd_SalConcO.Enabled = False
       fpd_AmortCon.Enabled = False
       fpd_PBPPerdi.Enabled = False
   Else
       fpd_PrstNCon.Text = Format(CDbl(g_rst_GenAux!HIPCIE_PRENCO), "###,###,#00.00") & " "
   End If
   
   pnl_PrstaCon.Caption = Format(CDbl(g_rst_GenAux!HIPCIE_PRECON), "###,###,#00.00") & " "
   fpd_PBPPerdi.Text = Format(CDbl(g_rst_GenAux!HIPCIE_PERPBP), "###,###,#00.00") & " "
   Call fpd_SalConcC_Change
   
   fpd_AmorNCon.Text = CDbl(g_rst_GenAux!HIPCIE_CPTNAM)
   fpd_AmortCon.Text = CDbl(g_rst_GenAux!HIPCIE_CPTCAM)
   If IsNull(g_rst_GenAux!HIPCIE_CLACLI) Then
      fpd_ClaCli.Text = Format(0, "###,###,#00.00") & " "
   Else
      fpd_ClaCli.Text = Format(CDbl(g_rst_GenAux!HIPCIE_CLACLI), "###,###,#00.00") & " "
   End If
   If IsNull(g_rst_GenAux!HIPCIE_CLAALI) Then
      fpd_ClaAli.Text = Format(0, "###,###,#00.00") & " "
   Else
      fpd_ClaAli.Text = Format(CDbl(g_rst_GenAux!HIPCIE_CLAALI), "###,###,#00.00") & " "
   End If
   If IsNull(g_rst_GenAux!HIPCIE_CLAPRV) Then
      fpd_ClaPrv.Text = Format(0, "###,###,#00.00") & " "
   Else
      fpd_ClaPrv.Text = Format(CDbl(g_rst_GenAux!HIPCIE_CLAPRV), "###,###,#00.00") & " "
   End If
   If IsNull(g_rst_GenAux!HIPCIE_PRVGEN) Then
      fpd_PrvGen.Text = Format(0, "###,###,#00.00") & " "
   Else
      fpd_PrvGen.Text = Format(CDbl(g_rst_GenAux!HIPCIE_PRVGEN), "###,###,#00.00") & " "
   End If
   If IsNull(g_rst_GenAux!HIPCIE_PRVCIC) Then
      fpd_PrvCic.Text = Format(0, "###,###,#00.00") & " "
   Else
      fpd_PrvCic.Text = Format(CDbl(g_rst_GenAux!HIPCIE_PRVCIC), "###,###,#00.00") & " "
   End If
   If IsNull(g_rst_GenAux!HIPCIE_PRVESP) Then
      fpd_PrvEsp.Text = Format(0, "###,###,#00.00") & " "
   Else
      fpd_PrvEsp.Text = Format(CDbl(g_rst_GenAux!HIPCIE_PRVESP), "###,###,#00.00") & " "
   End If
   If IsNull(g_rst_GenAux!HIPCIE_PRVVOL) Then
      fpd_PrvVol.Text = Format(0, "###,###,#00.00") & " "
   Else
      fpd_PrvVol.Text = Format(CDbl(g_rst_GenAux!HIPCIE_PRVVOL), "###,###,#00.00") & " "
   End If
   If IsNull(g_rst_GenAux!HIPCIE_APLCIC) Then
      fpd_Aplica.Text = Format(0, "###,###,#00.00") & " "
   Else
      fpd_Aplica.Text = Format(CDbl(g_rst_GenAux!HIPCIE_APLCIC), "###,###,#00.00") & " "
   End If
   '--
   If IsNull(g_rst_GenAux!HIPCIE_PRVGEN_RC) Then
      fpd_PrvGenRC.Text = Format(0, "###,###,#00.00") & " "
   Else
      fpd_PrvGenRC.Text = Format(CDbl(g_rst_GenAux!HIPCIE_PRVGEN_RC), "###,###,#00.00") & " "
   End If
   
   If IsNull(g_rst_GenAux!HIPCIE_PRVCIC_RC) Then
      fpd_PrvCicRC.Text = Format(0, "###,###,#00.00") & " "
   Else
      fpd_PrvCicRC.Text = Format(CDbl(g_rst_GenAux!HIPCIE_PRVCIC_RC), "###,###,#00.00") & " "
   End If
   
   If IsNull(g_rst_GenAux!HIPCIE_PRVRIP) Then
      fpd_PrvRip.Text = Format(0, "###,###,#00.00") & " "
   Else
      fpd_PrvRip.Text = Format(CDbl(g_rst_GenAux!HIPCIE_PRVRIP), "###,###,#00.00") & " "
   End If
   
   If IsNull(g_rst_GenAux!HIPCIE_CBRFMV) Then
      fpd_CbrFmv.Text = Format(0, "###,###,#00.00") & " "
   Else
      fpd_CbrFmv.Text = Format(CDbl(g_rst_GenAux!HIPCIE_CBRFMV), "###,###,#00.00") & " "
   End If
   
   If IsNull(g_rst_GenAux!HIPCIE_CBRFMV_RC) Then
      fpd_CbrFmvRC.Text = Format(0, "###,###,#00.00") & " "
   Else
      fpd_CbrFmvRC.Text = Format(CDbl(g_rst_GenAux!HIPCIE_CBRFMV_RC), "###,###,#00.00") & " "
   End If
   
   fpd_NumCuo.Text = g_rst_GenAux!HIPCIE_NUMCUO
   fpd_CuoPag.Text = g_rst_GenAux!HIPCIE_CUOPAG
   fpd_CuoPen.Text = g_rst_GenAux!HIPCIE_CUOPEN
   
   Call gs_BuscarCombo_Item(cmb_Refina, g_rst_GenAux!FLGREF)
   Call gs_BuscarCombo_Item(cmb_Judici, g_rst_GenAux!FLGJUD)
   Call gs_BuscarCombo_Item(cmb_Castig, g_rst_GenAux!FLGCAS)
   Call gs_BuscarCombo_Item(Cmb_TipGar, g_rst_GenAux!HIPCIE_TIPGAR)
   Call gs_BuscarCombo_Item(cmb_MonGar, g_rst_GenAux!HIPCIE_MONGAR)
   
   If IsNull(g_rst_GenAux!HIPCIE_MTOGAR) Then
      fpd_MtoGar.Text = Format(0, "###,###,#00.00") & " "
   Else
      fpd_MtoGar.Text = Format(CDbl(g_rst_GenAux!HIPCIE_MTOGAR), "###,###,#00.00") & " "
   End If
      
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
End Sub

Private Sub fs_Validar_Botones(ByVal r_bol_FlagEn As Boolean)
    msk_NumOpe.Enabled = Not r_bol_FlagEn
    cmb_PerMes.Enabled = Not r_bol_FlagEn
    ipp_PerAno.Enabled = Not r_bol_FlagEn
    cmd_Buscar.Enabled = Not r_bol_FlagEn
    cmd_VerPag.Enabled = r_bol_FlagEn
    cmd_ImpCro.Enabled = r_bol_FlagEn
    cmd_Grabar.Enabled = r_bol_FlagEn
    
    fpd_IntCapit.Enabled = r_bol_FlagEn
    fpd_CapAmort.Enabled = r_bol_FlagEn
    fpd_SalNConC.Enabled = r_bol_FlagEn
    fpd_SalConcC.Enabled = r_bol_FlagEn
    fpd_PBPPerdi.Enabled = r_bol_FlagEn
    fpd_IntDevnC.Enabled = r_bol_FlagEn
    fpd_IntSuspC.Enabled = r_bol_FlagEn
    fpd_IntMoraC.Enabled = r_bol_FlagEn
    fpd_DiaMoroC.Enabled = r_bol_FlagEn
    fpd_CapVencd.Enabled = r_bol_FlagEn
    fpd_TotCapVi.Enabled = r_bol_FlagEn
    fpd_TotCapVe.Enabled = r_bol_FlagEn
    fpd_SalNConO.Enabled = r_bol_FlagEn
    fpd_SalConcO.Enabled = r_bol_FlagEn
    fpd_IntDevnO.Enabled = r_bol_FlagEn
    fpd_IntSuspO.Enabled = r_bol_FlagEn
    fpd_IntMoraO.Enabled = r_bol_FlagEn
    fpd_DiaMoroO.Enabled = r_bol_FlagEn
    fpd_PrstNCon.Enabled = r_bol_FlagEn
    fpd_AmorNCon.Enabled = r_bol_FlagEn
    fpd_AmortCon.Enabled = r_bol_FlagEn
    fpd_ClaCli.Enabled = r_bol_FlagEn
    fpd_ClaAli.Enabled = r_bol_FlagEn
    fpd_ClaPrv.Enabled = r_bol_FlagEn
    fpd_PrvGen.Enabled = r_bol_FlagEn
    fpd_PrvCic.Enabled = r_bol_FlagEn
    fpd_PrvEsp.Enabled = r_bol_FlagEn
    fpd_PrvVol.Enabled = r_bol_FlagEn
    fpd_Aplica.Enabled = r_bol_FlagEn
End Sub

Private Sub msk_NumOpe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(cmb_PerMes)
    End If
End Sub

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(ipp_PerAno)
    End If
End Sub
 
Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(cmd_Buscar)
    End If
End Sub

Private Sub fpd_IntCapit_Change()
   pnl_SldDeud1.Caption = Format(CDbl(pnl_CapDesmb.Caption) + CDbl(fpd_IntCapit.Text) - CDbl(fpd_CapAmort.Text), "###,###,#00.00") & " "
End Sub

Private Sub fpd_IntCapit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_CapAmort)
    End If
End Sub

Private Sub fpd_CapAmort_Change()
    pnl_SldDeud1.Caption = Format(CDbl(pnl_CapDesmb.Caption) + CDbl(fpd_IntCapit.Text) - CDbl(fpd_CapAmort.Text), "###,###,#00.00") & " "
End Sub

Private Sub fpd_CapAmort_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_SalNConC)
    End If
End Sub

Private Sub fpd_SalNConC_Change()
    pnl_SldDeud2.Caption = Format(CDbl(fpd_SalNConC.Text) + CDbl(fpd_SalConcC.Text) + CDbl(fpd_PBPPerdi.Text), "###,###,#00.00") & " "
End Sub

Private Sub fpd_SalNConC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_SalConcC)
    End If
End Sub

Private Sub fpd_SalConcC_Change()
    pnl_SldDeud2.Caption = Format(CDbl(fpd_SalNConC.Text) + CDbl(fpd_SalConcC.Text) + CDbl(fpd_PBPPerdi.Text), "###,###,#00.00") & " "
End Sub

Private Sub fpd_SalConcC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_PBPPerdi)
    End If
End Sub

Private Sub fpd_PBPPerdi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_IntDevnC)
    End If
End Sub

Private Sub fpd_PBPPerdi_Change()
    pnl_SldDeud2.Caption = Format(CDbl(fpd_SalNConC.Text) + CDbl(fpd_SalConcC.Text) + CDbl(fpd_PBPPerdi.Text), "###,###,#00.00") & " "
End Sub

Private Sub fpd_IntDevnC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_IntSuspC)
    End If
End Sub

Private Sub fpd_IntSuspC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_IntMoraC)
    End If
End Sub

Private Sub fpd_IntMoraC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_DiaMoroC)
    End If
End Sub

Private Sub fpd_DiaMoroC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_CapVencd)
    End If
End Sub

Private Sub fpd_CapVencd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_TotCapVi)
    End If
End Sub

Private Sub fpd_TotCapVi_Change()
    pnl_SldDeud3.Caption = Format(CDbl(fpd_TotCapVi.Text) + CDbl(fpd_TotCapVe.Text), "###,###,#00.00") & " "
End Sub

Private Sub fpd_TotCapVi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_TotCapVe)
    End If
End Sub

Private Sub fpd_TotCapVe_Change()
    pnl_SldDeud3.Caption = Format(CDbl(fpd_TotCapVi.Text) + CDbl(fpd_TotCapVe.Text), "###,###,#00.00") & " "
End Sub

Private Sub fpd_TotCapVe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_SalNConO)
    End If
End Sub

Private Sub fpd_SalNConO_Change()
    pnl_SldDeud4.Caption = Format(CDbl(fpd_SalNConO.Text) + CDbl(fpd_SalConcO.Text), "###,###,#00.00") & " "
    pnl_SaldNCon.Caption = Format(CDbl(fpd_SalNConO.Text), "######,#00.00") & " "
End Sub

Private Sub fpd_SalNConO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_SalConcO)
    End If
End Sub

Private Sub fpd_SalConcO_Change()
    pnl_SldDeud4.Caption = Format(CDbl(fpd_SalNConO.Text) + CDbl(fpd_SalConcO.Text), "###,###,#00.00") & " "
    pnl_SaldoCon.Caption = Format(CDbl(fpd_SalConcO.Text), "######,#00.00") & " "
End Sub

Private Sub fpd_SalConcO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_IntDevnO)
    End If
End Sub

Private Sub fpd_IntDevnO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_IntSuspO)
    End If
End Sub

Private Sub fpd_IntSuspO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_IntMoraO)
    End If
End Sub

Private Sub fpd_IntMoraO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_DiaMoroO)
    End If
End Sub

Private Sub fpd_DiaMoroO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_PrstNCon)
    End If
End Sub

Private Sub fpd_PrstNCon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_AmorNCon)
    End If
End Sub

Private Sub fpd_AmorNCon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_AmortCon)
    End If
End Sub
  
Private Sub fpd_AmortCon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         Me.SSTab1.Tab = 1
        Call gs_SetFocus(fpd_ClaCli)
    End If
End Sub

Private Sub fpd_ClaCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_ClaAli)
    End If
End Sub

Private Sub fpd_ClaAli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_ClaPrv)
    End If
End Sub

Private Sub fpd_ClaPrv_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_Aplica)
    End If
End Sub

Private Sub fpd_Aplica_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_PrvGen)
    End If
End Sub

Private Sub fpd_PrvGen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_PrvGenRC)
    End If
End Sub

Private Sub fpd_PrvGenRC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_PrvEsp)
    End If
End Sub

Private Sub fpd_PrvEsp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_PrvVol)
    End If
End Sub

Private Sub fpd_PrvVol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_PrvCic)
    End If
End Sub

Private Sub fpd_PrvCic_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_PrvCicRC)
    End If
End Sub

Private Sub fpd_PrvCicRC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_PrvRip)
    End If
End Sub

Private Sub fpd_PrvRip_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_CbrFmv)
    End If
End Sub

Private Sub fpd_CbrFmv_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_CbrFmvRC)
    End If
End Sub

Private Sub fpd_CbrFmvRC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_NumCuo)
    End If
End Sub

Private Sub fpd_NumCuo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_CuoPag)
    End If
End Sub

Private Sub fpd_CuoPag_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_CuoPen)
    End If
End Sub

Private Sub fpd_CuoPen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(cmb_Refina)
    End If
End Sub

Private Sub cmb_Refina_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(cmb_Judici)
    End If
End Sub

Private Sub cmb_Judici_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(cmb_Castig)
    End If
End Sub

Private Sub cmb_Castig_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(Cmb_TipGar)
    End If
End Sub

Private Sub cmb_TipGar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(cmb_MonGar)
    End If
End Sub

Private Sub cmb_MonGar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_MtoGar)
    End If
End Sub

Private Sub fpd_MtoGar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(cmd_Grabar)
    End If
End Sub
