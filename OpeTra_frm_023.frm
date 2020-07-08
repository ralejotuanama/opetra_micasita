VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frm_EvaTas_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   9960
   ClientLeft      =   2880
   ClientTop       =   555
   ClientWidth     =   11250
   Icon            =   "OpeTra_frm_023.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9960
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9960
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   11250
      _Version        =   65536
      _ExtentX        =   19844
      _ExtentY        =   17568
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
      Begin Threed.SSPanel SSPanel8 
         Height          =   1875
         Left            =   30
         TabIndex        =   78
         Top             =   7230
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
         _ExtentY        =   3307
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
         Begin Threed.SSPanel pnl_AreTer 
            Height          =   315
            Left            =   1860
            TabIndex        =   79
            Top             =   390
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
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
         Begin Threed.SSPanel pnl_AreCon 
            Height          =   315
            Left            =   1860
            TabIndex        =   115
            Top             =   720
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   60
            Left            =   30
            TabIndex        =   117
            Top             =   1080
            Width           =   11115
            _Version        =   65536
            _ExtentX        =   19606
            _ExtentY        =   106
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
         End
         Begin Threed.SSPanel pnl_SumAse 
            Height          =   315
            Left            =   1860
            TabIndex        =   118
            Top             =   1170
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
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
         Begin Threed.SSPanel pnl_ValCom 
            Height          =   315
            Left            =   6030
            TabIndex        =   120
            Top             =   1170
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
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
         Begin Threed.SSPanel pnl_ValRea 
            Height          =   315
            Left            =   9780
            TabIndex        =   122
            Top             =   1170
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
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
         Begin Threed.SSPanel pnl_ValTer 
            Height          =   315
            Left            =   1860
            TabIndex        =   124
            Top             =   1500
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
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
         Begin Threed.SSPanel pnl_ValEdi 
            Height          =   315
            Left            =   6030
            TabIndex        =   126
            Top             =   1500
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
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
         Begin Threed.SSPanel pnl_ValACo 
            Height          =   315
            Left            =   9780
            TabIndex        =   128
            Top             =   1500
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
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
         Begin VB.Label Label56 
            Caption         =   "Valor Areas Comunes:"
            Height          =   315
            Left            =   7920
            TabIndex        =   129
            Top             =   1500
            Width           =   1815
         End
         Begin VB.Label Label55 
            Caption         =   "Valor Edificación:"
            Height          =   315
            Left            =   4170
            TabIndex        =   127
            Top             =   1500
            Width           =   1485
         End
         Begin VB.Label Label54 
            Caption         =   "Valor Terreno:"
            Height          =   315
            Left            =   90
            TabIndex        =   125
            Top             =   1500
            Width           =   1485
         End
         Begin VB.Label Label53 
            Caption         =   "Valor Realización:"
            Height          =   315
            Left            =   7920
            TabIndex        =   123
            Top             =   1170
            Width           =   1485
         End
         Begin VB.Label Label52 
            Caption         =   "Valor Comercial:"
            Height          =   315
            Left            =   4170
            TabIndex        =   121
            Top             =   1170
            Width           =   1485
         End
         Begin VB.Label Label51 
            Caption         =   "Suma Asegurada:"
            Height          =   315
            Left            =   90
            TabIndex        =   119
            Top             =   1170
            Width           =   1485
         End
         Begin VB.Label Label50 
            Caption         =   "Area Construcción:"
            Height          =   315
            Left            =   90
            TabIndex        =   116
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label49 
            Caption         =   "Totales"
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
            TabIndex        =   114
            Top             =   60
            Width           =   1065
         End
         Begin VB.Label Label11 
            Caption         =   "Area Terreno:"
            Height          =   315
            Left            =   90
            TabIndex        =   80
            Top             =   390
            Width           =   1185
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   2505
         Left            =   30
         TabIndex        =   66
         Top             =   4680
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
         _ExtentY        =   4419
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
         Begin TabDlg.SSTab tab_Genera 
            Height          =   2385
            Left            =   60
            TabIndex        =   7
            Top             =   60
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   4207
            _Version        =   393216
            Style           =   1
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "Inmueble"
            TabPicture(0)   =   "OpeTra_frm_023.frx":000C
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label22"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label23"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label15"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Label16"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "Label17"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "Label18"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "Label19"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "Label20"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "ipp_ValRea_Inm"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "ipp_ValCom_Inm"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "ipp_ValACo_Inm"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "ipp_ValEdi_Inm"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "ipp_ValTer_Inm"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "ipp_SumAse_Inm"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "ipp_AreCon_Inm"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "ipp_AreTer_Inm"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "SSPanel5"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).ControlCount=   17
            TabCaption(1)   =   "Estacionamiento 1"
            TabPicture(1)   =   "OpeTra_frm_023.frx":0028
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Label9"
            Tab(1).Control(1)=   "Label10"
            Tab(1).Control(2)=   "Label24"
            Tab(1).Control(3)=   "Label25"
            Tab(1).Control(4)=   "Label26"
            Tab(1).Control(5)=   "Label27"
            Tab(1).Control(6)=   "Label28"
            Tab(1).Control(7)=   "Label29"
            Tab(1).Control(8)=   "Label30"
            Tab(1).Control(9)=   "SSPanel10"
            Tab(1).Control(10)=   "ipp_ValRea_Es1"
            Tab(1).Control(11)=   "ipp_ValCom_Es1"
            Tab(1).Control(12)=   "ipp_ValACo_Es1"
            Tab(1).Control(13)=   "ipp_ValEdi_Es1"
            Tab(1).Control(14)=   "ipp_ValTer_Es1"
            Tab(1).Control(15)=   "ipp_SumAse_Es1"
            Tab(1).Control(16)=   "ipp_AreCon_Es1"
            Tab(1).Control(17)=   "ipp_AreTer_Es1"
            Tab(1).Control(18)=   "SSPanel9"
            Tab(1).Control(19)=   "cmb_FlgEst_Es1"
            Tab(1).ControlCount=   20
            TabCaption(2)   =   "Estacionamiento 2"
            TabPicture(2)   =   "OpeTra_frm_023.frx":0044
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "cmb_FlgEst_Es2"
            Tab(2).Control(1)=   "SSPanel11"
            Tab(2).Control(2)=   "ipp_AreTer_Es2"
            Tab(2).Control(3)=   "ipp_AreCon_Es2"
            Tab(2).Control(4)=   "ipp_SumAse_Es2"
            Tab(2).Control(5)=   "ipp_ValTer_Es2"
            Tab(2).Control(6)=   "ipp_ValEdi_Es2"
            Tab(2).Control(7)=   "ipp_ValACo_Es2"
            Tab(2).Control(8)=   "ipp_ValCom_Es2"
            Tab(2).Control(9)=   "ipp_ValRea_Es2"
            Tab(2).Control(10)=   "SSPanel12"
            Tab(2).Control(11)=   "Label39"
            Tab(2).Control(12)=   "Label38"
            Tab(2).Control(13)=   "Label37"
            Tab(2).Control(14)=   "Label36"
            Tab(2).Control(15)=   "Label35"
            Tab(2).Control(16)=   "Label34"
            Tab(2).Control(17)=   "Label33"
            Tab(2).Control(18)=   "Label32"
            Tab(2).Control(19)=   "Label31"
            Tab(2).ControlCount=   20
            TabCaption(3)   =   "Depósito"
            TabPicture(3)   =   "OpeTra_frm_023.frx":0060
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "Label40"
            Tab(3).Control(1)=   "Label41"
            Tab(3).Control(2)=   "Label42"
            Tab(3).Control(3)=   "Label43"
            Tab(3).Control(4)=   "Label44"
            Tab(3).Control(5)=   "Label45"
            Tab(3).Control(6)=   "Label46"
            Tab(3).Control(7)=   "Label47"
            Tab(3).Control(8)=   "Label48"
            Tab(3).Control(9)=   "SSPanel14"
            Tab(3).Control(10)=   "ipp_ValRea_Dep"
            Tab(3).Control(11)=   "ipp_ValCom_Dep"
            Tab(3).Control(12)=   "ipp_ValACo_Dep"
            Tab(3).Control(13)=   "ipp_ValEdi_Dep"
            Tab(3).Control(14)=   "ipp_ValTer_Dep"
            Tab(3).Control(15)=   "ipp_SumAse_Dep"
            Tab(3).Control(16)=   "ipp_AreCon_Dep"
            Tab(3).Control(17)=   "ipp_AreTer_Dep"
            Tab(3).Control(18)=   "SSPanel13"
            Tab(3).Control(19)=   "cmb_FlgEst_Dep"
            Tab(3).ControlCount=   20
            Begin VB.ComboBox cmb_FlgEst_Dep 
               Height          =   315
               Left            =   -73140
               Style           =   2  'Dropdown List
               TabIndex        =   34
               Top             =   390
               Width           =   975
            End
            Begin VB.ComboBox cmb_FlgEst_Es2 
               Height          =   315
               Left            =   -73140
               Style           =   2  'Dropdown List
               TabIndex        =   25
               Top             =   390
               Width           =   975
            End
            Begin VB.ComboBox cmb_FlgEst_Es1 
               Height          =   315
               Left            =   -73140
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   390
               Width           =   975
            End
            Begin Threed.SSPanel SSPanel5 
               Height          =   60
               Left            =   30
               TabIndex        =   75
               Top             =   1080
               Width           =   10965
               _Version        =   65536
               _ExtentX        =   19341
               _ExtentY        =   106
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
            End
            Begin EditLib.fpDoubleSingle ipp_AreTer_Inm 
               Height          =   315
               Left            =   1860
               TabIndex        =   8
               Top             =   390
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_AreCon_Inm 
               Height          =   315
               Left            =   1860
               TabIndex        =   9
               Top             =   720
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_SumAse_Inm 
               Height          =   315
               Left            =   1860
               TabIndex        =   10
               Top             =   1200
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValTer_Inm 
               Height          =   315
               Left            =   1860
               TabIndex        =   13
               Top             =   1530
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValEdi_Inm 
               Height          =   315
               Left            =   5520
               TabIndex        =   14
               Top             =   1530
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValACo_Inm 
               Height          =   315
               Left            =   9450
               TabIndex        =   15
               Top             =   1530
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValCom_Inm 
               Height          =   315
               Left            =   5520
               TabIndex        =   11
               Top             =   1200
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValRea_Inm 
               Height          =   315
               Left            =   9450
               TabIndex        =   12
               Top             =   1200
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
               MinValue        =   "-9000000000"
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
            Begin Threed.SSPanel SSPanel9 
               Height          =   60
               Left            =   -74970
               TabIndex        =   82
               Top             =   750
               Width           =   10965
               _Version        =   65536
               _ExtentX        =   19341
               _ExtentY        =   106
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
            End
            Begin EditLib.fpDoubleSingle ipp_AreTer_Es1 
               Height          =   315
               Left            =   -73140
               TabIndex        =   17
               Top             =   870
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_AreCon_Es1 
               Height          =   315
               Left            =   -73140
               TabIndex        =   18
               Top             =   1200
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_SumAse_Es1 
               Height          =   315
               Left            =   -73140
               TabIndex        =   19
               Top             =   1680
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValTer_Es1 
               Height          =   315
               Left            =   -73140
               TabIndex        =   22
               Top             =   2010
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValEdi_Es1 
               Height          =   315
               Left            =   -69480
               TabIndex        =   23
               Top             =   2010
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValACo_Es1 
               Height          =   315
               Left            =   -65550
               TabIndex        =   24
               Top             =   2010
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValCom_Es1 
               Height          =   315
               Left            =   -69480
               TabIndex        =   20
               Top             =   1680
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValRea_Es1 
               Height          =   315
               Left            =   -65550
               TabIndex        =   21
               Top             =   1680
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
               MinValue        =   "-9000000000"
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
            Begin Threed.SSPanel SSPanel10 
               Height          =   60
               Left            =   -74970
               TabIndex        =   91
               Top             =   1560
               Width           =   10965
               _Version        =   65536
               _ExtentX        =   19341
               _ExtentY        =   106
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
            End
            Begin Threed.SSPanel SSPanel11 
               Height          =   60
               Left            =   -74970
               TabIndex        =   93
               Top             =   750
               Width           =   10965
               _Version        =   65536
               _ExtentX        =   19341
               _ExtentY        =   106
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
            End
            Begin EditLib.fpDoubleSingle ipp_AreTer_Es2 
               Height          =   315
               Left            =   -73140
               TabIndex        =   26
               Top             =   870
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_AreCon_Es2 
               Height          =   315
               Left            =   -73140
               TabIndex        =   27
               Top             =   1200
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_SumAse_Es2 
               Height          =   315
               Left            =   -73140
               TabIndex        =   28
               Top             =   1680
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValTer_Es2 
               Height          =   315
               Left            =   -73140
               TabIndex        =   31
               Top             =   2010
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValEdi_Es2 
               Height          =   315
               Left            =   -69480
               TabIndex        =   32
               Top             =   2010
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValACo_Es2 
               Height          =   315
               Left            =   -65550
               TabIndex        =   33
               Top             =   2010
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValCom_Es2 
               Height          =   315
               Left            =   -69480
               TabIndex        =   29
               Top             =   1680
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValRea_Es2 
               Height          =   315
               Left            =   -65550
               TabIndex        =   30
               Top             =   1680
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
               MinValue        =   "-9000000000"
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
            Begin Threed.SSPanel SSPanel12 
               Height          =   60
               Left            =   -74970
               TabIndex        =   102
               Top             =   1560
               Width           =   10965
               _Version        =   65536
               _ExtentX        =   19341
               _ExtentY        =   106
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
            End
            Begin Threed.SSPanel SSPanel13 
               Height          =   60
               Left            =   -74970
               TabIndex        =   104
               Top             =   750
               Width           =   10965
               _Version        =   65536
               _ExtentX        =   19341
               _ExtentY        =   106
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
            End
            Begin EditLib.fpDoubleSingle ipp_AreTer_Dep 
               Height          =   315
               Left            =   -73140
               TabIndex        =   35
               Top             =   870
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_AreCon_Dep 
               Height          =   315
               Left            =   -73140
               TabIndex        =   36
               Top             =   1200
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_SumAse_Dep 
               Height          =   315
               Left            =   -73140
               TabIndex        =   37
               Top             =   1680
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValTer_Dep 
               Height          =   315
               Left            =   -73140
               TabIndex        =   40
               Top             =   2010
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValEdi_Dep 
               Height          =   315
               Left            =   -69480
               TabIndex        =   41
               Top             =   2010
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValACo_Dep 
               Height          =   315
               Left            =   -65550
               TabIndex        =   42
               Top             =   2010
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValCom_Dep 
               Height          =   315
               Left            =   -69480
               TabIndex        =   38
               Top             =   1680
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValRea_Dep 
               Height          =   315
               Left            =   -65550
               TabIndex        =   39
               Top             =   1680
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
               MinValue        =   "-9000000000"
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
            Begin Threed.SSPanel SSPanel14 
               Height          =   60
               Left            =   -74970
               TabIndex        =   113
               Top             =   1560
               Width           =   10965
               _Version        =   65536
               _ExtentX        =   19341
               _ExtentY        =   106
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
            End
            Begin VB.Label Label48 
               Caption         =   "Valor Realización:"
               Height          =   285
               Left            =   -67320
               TabIndex        =   112
               Top             =   1680
               Width           =   1485
            End
            Begin VB.Label Label47 
               Caption         =   "Valor Comercial:"
               Height          =   285
               Left            =   -70980
               TabIndex        =   111
               Top             =   1680
               Width           =   1485
            End
            Begin VB.Label Label46 
               Caption         =   "Valor Areas Comunes:"
               Height          =   285
               Left            =   -67320
               TabIndex        =   110
               Top             =   2010
               Width           =   1725
            End
            Begin VB.Label Label45 
               Caption         =   "Valor Edificación:"
               Height          =   285
               Left            =   -70980
               TabIndex        =   109
               Top             =   2010
               Width           =   1485
            End
            Begin VB.Label Label44 
               Caption         =   "Valor Terreno:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   108
               Top             =   2010
               Width           =   1485
            End
            Begin VB.Label Label43 
               Caption         =   "Suma Asegurada:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   107
               Top             =   1680
               Width           =   1485
            End
            Begin VB.Label Label42 
               Caption         =   "Area Construcción:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   106
               Top             =   1200
               Width           =   1485
            End
            Begin VB.Label Label41 
               Caption         =   "Area Terreno:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   105
               Top             =   870
               Width           =   1485
            End
            Begin VB.Label Label40 
               Caption         =   "Depósito:"
               Height          =   315
               Left            =   -74910
               TabIndex        =   103
               Top             =   390
               Width           =   1365
            End
            Begin VB.Label Label39 
               Caption         =   "Valor Realización:"
               Height          =   285
               Left            =   -67320
               TabIndex        =   101
               Top             =   1680
               Width           =   1485
            End
            Begin VB.Label Label38 
               Caption         =   "Valor Comercial:"
               Height          =   285
               Left            =   -70980
               TabIndex        =   100
               Top             =   1680
               Width           =   1485
            End
            Begin VB.Label Label37 
               Caption         =   "Valor Areas Comunes:"
               Height          =   285
               Left            =   -67320
               TabIndex        =   99
               Top             =   2010
               Width           =   1725
            End
            Begin VB.Label Label36 
               Caption         =   "Valor Edificación:"
               Height          =   285
               Left            =   -70980
               TabIndex        =   98
               Top             =   2010
               Width           =   1485
            End
            Begin VB.Label Label35 
               Caption         =   "Valor Terreno:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   97
               Top             =   2010
               Width           =   1485
            End
            Begin VB.Label Label34 
               Caption         =   "Suma Asegurada:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   96
               Top             =   1680
               Width           =   1485
            End
            Begin VB.Label Label33 
               Caption         =   "Area Construcción:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   95
               Top             =   1200
               Width           =   1485
            End
            Begin VB.Label Label32 
               Caption         =   "Area Terreno:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   94
               Top             =   870
               Width           =   1485
            End
            Begin VB.Label Label31 
               Caption         =   "Estacionamiento 2:"
               Height          =   315
               Left            =   -74910
               TabIndex        =   92
               Top             =   390
               Width           =   1365
            End
            Begin VB.Label Label30 
               Caption         =   "Valor Realización:"
               Height          =   285
               Left            =   -67320
               TabIndex        =   90
               Top             =   1680
               Width           =   1485
            End
            Begin VB.Label Label29 
               Caption         =   "Valor Comercial:"
               Height          =   285
               Left            =   -70980
               TabIndex        =   89
               Top             =   1680
               Width           =   1485
            End
            Begin VB.Label Label28 
               Caption         =   "Valor Areas Comunes:"
               Height          =   285
               Left            =   -67320
               TabIndex        =   88
               Top             =   2010
               Width           =   1725
            End
            Begin VB.Label Label27 
               Caption         =   "Valor Edificación:"
               Height          =   285
               Left            =   -70980
               TabIndex        =   87
               Top             =   2010
               Width           =   1485
            End
            Begin VB.Label Label26 
               Caption         =   "Valor Terreno:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   86
               Top             =   2010
               Width           =   1485
            End
            Begin VB.Label Label25 
               Caption         =   "Suma Asegurada:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   85
               Top             =   1680
               Width           =   1485
            End
            Begin VB.Label Label24 
               Caption         =   "Area Construcción:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   84
               Top             =   1200
               Width           =   1485
            End
            Begin VB.Label Label10 
               Caption         =   "Area Terreno:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   83
               Top             =   870
               Width           =   1485
            End
            Begin VB.Label Label9 
               Caption         =   "Estacionamiento 1:"
               Height          =   315
               Left            =   -74910
               TabIndex        =   81
               Top             =   390
               Width           =   1365
            End
            Begin VB.Label Label20 
               Caption         =   "Valor Realización:"
               Height          =   285
               Left            =   7680
               TabIndex        =   74
               Top             =   1200
               Width           =   1485
            End
            Begin VB.Label Label19 
               Caption         =   "Valor Comercial:"
               Height          =   285
               Left            =   4020
               TabIndex        =   73
               Top             =   1200
               Width           =   1485
            End
            Begin VB.Label Label18 
               Caption         =   "Valor Areas Comunes:"
               Height          =   285
               Left            =   7680
               TabIndex        =   72
               Top             =   1530
               Width           =   1725
            End
            Begin VB.Label Label17 
               Caption         =   "Valor Edificación:"
               Height          =   285
               Left            =   4020
               TabIndex        =   71
               Top             =   1530
               Width           =   1485
            End
            Begin VB.Label Label16 
               Caption         =   "Valor Terreno:"
               Height          =   285
               Left            =   90
               TabIndex        =   70
               Top             =   1530
               Width           =   1485
            End
            Begin VB.Label Label15 
               Caption         =   "Suma Asegurada:"
               Height          =   285
               Left            =   90
               TabIndex        =   69
               Top             =   1200
               Width           =   1485
            End
            Begin VB.Label Label23 
               Caption         =   "Area Construcción:"
               Height          =   285
               Left            =   90
               TabIndex        =   68
               Top             =   720
               Width           =   1485
            End
            Begin VB.Label Label22 
               Caption         =   "Area Terreno:"
               Height          =   285
               Left            =   90
               TabIndex        =   67
               Top             =   390
               Width           =   1485
            End
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   2745
         Left            =   30
         TabIndex        =   46
         Top             =   1890
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
         _ExtentY        =   4842
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
         Begin VB.ComboBox cmb_MatCon 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   140
            Top             =   2040
            Width           =   9255
         End
         Begin VB.ComboBox cmb_UsoInm 
            Height          =   315
            Left            =   8340
            Style           =   2  'Dropdown List
            TabIndex        =   138
            Top             =   1710
            Width           =   2775
         End
         Begin VB.ComboBox cmb_TipInm 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   136
            Top             =   1710
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipMon 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2370
            Width           =   3315
         End
         Begin VB.TextBox txt_NumInf 
            Height          =   315
            Left            =   1860
            MaxLength       =   25
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   720
            Width           =   1635
         End
         Begin VB.TextBox txt_CodPer 
            Height          =   315
            Left            =   8340
            MaxLength       =   25
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   390
            Width           =   1635
         End
         Begin VB.ComboBox cmb_EmpPer 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   3345
         End
         Begin VB.TextBox txt_NomPer 
            Height          =   315
            Left            =   1860
            MaxLength       =   60
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   390
            Width           =   3345
         End
         Begin EditLib.fpDateTime ipp_FecEva 
            Height          =   315
            Left            =   8340
            TabIndex        =   4
            Top             =   720
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
         Begin EditLib.fpDoubleSingle ipp_TipCam 
            Height          =   315
            Left            =   8340
            TabIndex        =   6
            Top             =   2370
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
            Text            =   "0.0000"
            DecimalPlaces   =   4
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
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
         Begin EditLib.fpDoubleSingle ipp_AnoCon 
            Height          =   315
            Left            =   1860
            TabIndex        =   130
            Top             =   1050
            Width           =   1065
            _Version        =   196608
            _ExtentX        =   1879
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
         Begin EditLib.fpDoubleSingle ipp_NumPis 
            Height          =   315
            Left            =   1860
            TabIndex        =   132
            Top             =   1380
            Width           =   765
            _Version        =   196608
            _ExtentX        =   1349
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
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "99"
            MinValue        =   "1"
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
         Begin EditLib.fpDoubleSingle ipp_NumSot 
            Height          =   315
            Left            =   8340
            TabIndex        =   134
            Top             =   1380
            Width           =   765
            _Version        =   196608
            _ExtentX        =   1349
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
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "99"
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
         Begin VB.Label Label62 
            Caption         =   "Material Construcción:"
            Height          =   315
            Left            =   60
            TabIndex        =   141
            Top             =   2040
            Width           =   1635
         End
         Begin VB.Label Label61 
            Caption         =   "Uso Inmueble:"
            Height          =   315
            Left            =   6780
            TabIndex        =   139
            Top             =   1710
            Width           =   1065
         End
         Begin VB.Label Label60 
            Caption         =   "Tipo Inmueble:"
            Height          =   315
            Left            =   60
            TabIndex        =   137
            Top             =   1710
            Width           =   1065
         End
         Begin VB.Label Label59 
            Caption         =   "Nro. Sótanos:"
            Height          =   285
            Left            =   6780
            TabIndex        =   135
            Top             =   1380
            Width           =   1485
         End
         Begin VB.Label Label58 
            Caption         =   "Nro. Pisos:"
            Height          =   285
            Left            =   60
            TabIndex        =   133
            Top             =   1380
            Width           =   1485
         End
         Begin VB.Label Label57 
            Caption         =   "Año Construcción:"
            Height          =   285
            Left            =   60
            TabIndex        =   131
            Top             =   1050
            Width           =   1485
         End
         Begin VB.Label Label14 
            Caption         =   "Tipo de Cambio:"
            Height          =   315
            Left            =   6780
            TabIndex        =   53
            Top             =   2370
            Width           =   1365
         End
         Begin VB.Label Label13 
            Caption         =   "Moneda:"
            Height          =   315
            Left            =   60
            TabIndex        =   52
            Top             =   2370
            Width           =   1065
         End
         Begin VB.Label Label12 
            Caption         =   "F. Evaluación:"
            Height          =   285
            Left            =   6780
            TabIndex        =   51
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label8 
            Caption         =   "Número Informe:"
            Height          =   285
            Left            =   60
            TabIndex        =   50
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label7 
            Caption         =   "Código REPEV SBS:"
            Height          =   285
            Left            =   6780
            TabIndex        =   49
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label5 
            Caption         =   "Empresa Peritaje:"
            Height          =   285
            Left            =   60
            TabIndex        =   48
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label Label4 
            Caption         =   "Nombre Perito:"
            Height          =   285
            Left            =   60
            TabIndex        =   47
            Top             =   390
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   54
         Top             =   30
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
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
            Left            =   630
            TabIndex        =   55
            Top             =   60
            Width           =   5415
            _Version        =   65536
            _ExtentX        =   9551
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Tasación del Inmueble"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
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
            Picture         =   "OpeTra_frm_023.frx":007C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   1095
         Left            =   30
         TabIndex        =   56
         Top             =   750
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
         _ExtentY        =   1931
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
            TabIndex        =   57
            Top             =   60
            Width           =   4485
            _Version        =   65536
            _ExtentX        =   7911
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1860
            TabIndex        =   58
            Top             =   390
            Width           =   2535
            _Version        =   65536
            _ExtentX        =   4471
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
         Begin Threed.SSPanel pnl_FecIng 
            Height          =   315
            Left            =   9690
            TabIndex        =   59
            Top             =   60
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
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
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_IngIns 
            Height          =   315
            Left            =   9690
            TabIndex        =   60
            Top             =   390
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
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
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1860
            TabIndex        =   76
            Top             =   720
            Width           =   9255
            _Version        =   65536
            _ExtentX        =   16325
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
         Begin VB.Label Label2 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   77
            Top             =   720
            Width           =   1125
         End
         Begin VB.Label Label3 
            Caption         =   "F. Ingreso Solicitud:"
            Height          =   315
            Left            =   8040
            TabIndex        =   64
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label21 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   63
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   62
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "F. Ingreso Instancia:"
            Height          =   315
            Left            =   8040
            TabIndex        =   61
            Top             =   390
            Width           =   1455
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   765
         Left            =   30
         TabIndex        =   65
         Top             =   9150
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10470
            Picture         =   "OpeTra_frm_023.frx":0386
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   9780
            Picture         =   "OpeTra_frm_023.frx":07C8
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frm_EvaTas_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_EmpPer()   As moddat_tpo_Genera
Dim l_arr_ParPrd()   As moddat_tpo_Genera

Private Sub cmb_EmpPer_Click()
   Call gs_SetFocus(txt_NomPer)
End Sub

Private Sub cmb_EmpPer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_EmpPer)
   End If
End Sub

Private Sub cmb_FlgEst_Dep_Click()
   If cmb_FlgEst_Dep.ListIndex > -1 Then
      If cmb_FlgEst_Dep.ItemData(cmb_FlgEst_Dep.ListIndex) = 1 Then
         ipp_AreTer_Dep.Enabled = True
         ipp_AreCon_Dep.Enabled = True
         ipp_SumAse_Dep.Enabled = True
         ipp_ValCom_Dep.Enabled = True
         ipp_ValRea_Dep.Enabled = True
         ipp_ValTer_Dep.Enabled = True
         ipp_ValEdi_Dep.Enabled = True
         ipp_ValACo_Dep.Enabled = True
         
         Call gs_SetFocus(ipp_AreTer_Dep)
      Else
         ipp_AreTer_Dep.Enabled = False
         ipp_AreCon_Dep.Enabled = False
         ipp_SumAse_Dep.Enabled = False
         ipp_ValCom_Dep.Enabled = False
         ipp_ValRea_Dep.Enabled = False
         ipp_ValTer_Dep.Enabled = False
         ipp_ValEdi_Dep.Enabled = False
         ipp_ValACo_Dep.Enabled = False
      
         ipp_AreTer_Dep.Value = 0
         ipp_AreCon_Dep.Value = 0
         ipp_SumAse_Dep.Value = 0
         ipp_ValCom_Dep.Value = 0
         ipp_ValRea_Dep.Value = 0
         ipp_ValTer_Dep.Value = 0
         ipp_ValEdi_Dep.Value = 0
         ipp_ValACo_Dep.Value = 0
         
         Call gs_SetFocus(cmd_Grabar)
      End If
   End If
End Sub

Private Sub cmb_FlgEst_Dep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_FlgEst_Dep_Click
   End If
End Sub

Private Sub cmb_FlgEst_Es1_Click()
   If cmb_FlgEst_Es1.ListIndex > -1 Then
      If cmb_FlgEst_Es1.ItemData(cmb_FlgEst_Es1.ListIndex) = 1 Then
         ipp_AreTer_Es1.Enabled = True
         ipp_AreCon_Es1.Enabled = True
         ipp_SumAse_Es1.Enabled = True
         ipp_ValCom_Es1.Enabled = True
         ipp_ValRea_Es1.Enabled = True
         ipp_ValTer_Es1.Enabled = True
         ipp_ValEdi_Es1.Enabled = True
         ipp_ValACo_Es1.Enabled = True
         
         Call gs_SetFocus(ipp_AreTer_Es1)
      Else
         ipp_AreTer_Es1.Enabled = False
         ipp_AreCon_Es1.Enabled = False
         ipp_SumAse_Es1.Enabled = False
         ipp_ValCom_Es1.Enabled = False
         ipp_ValRea_Es1.Enabled = False
         ipp_ValTer_Es1.Enabled = False
         ipp_ValEdi_Es1.Enabled = False
         ipp_ValACo_Es1.Enabled = False
      
         ipp_AreTer_Es1.Value = 0
         ipp_AreCon_Es1.Value = 0
         ipp_SumAse_Es1.Value = 0
         ipp_ValCom_Es1.Value = 0
         ipp_ValRea_Es1.Value = 0
         ipp_ValTer_Es1.Value = 0
         ipp_ValEdi_Es1.Value = 0
         ipp_ValACo_Es1.Value = 0
         
         tab_Genera.Tab = 2
         Call gs_SetFocus(cmb_FlgEst_Es2)
      End If
   End If
End Sub

Private Sub cmb_FlgEst_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_FlgEst_Es1_Click
   End If
End Sub

Private Sub cmb_FlgEst_Es2_Click()
   If cmb_FlgEst_Es2.ListIndex > -1 Then
      If cmb_FlgEst_Es2.ItemData(cmb_FlgEst_Es2.ListIndex) = 1 Then
         ipp_AreTer_Es2.Enabled = True
         ipp_AreCon_Es2.Enabled = True
         ipp_SumAse_Es2.Enabled = True
         ipp_ValCom_Es2.Enabled = True
         ipp_ValRea_Es2.Enabled = True
         ipp_ValTer_Es2.Enabled = True
         ipp_ValEdi_Es2.Enabled = True
         ipp_ValACo_Es2.Enabled = True
         
         Call gs_SetFocus(ipp_AreTer_Es2)
      Else
         ipp_AreTer_Es2.Enabled = False
         ipp_AreCon_Es2.Enabled = False
         ipp_SumAse_Es2.Enabled = False
         ipp_ValCom_Es2.Enabled = False
         ipp_ValRea_Es2.Enabled = False
         ipp_ValTer_Es2.Enabled = False
         ipp_ValEdi_Es2.Enabled = False
         ipp_ValACo_Es2.Enabled = False
      
         ipp_AreTer_Es2.Value = 0
         ipp_AreCon_Es2.Value = 0
         ipp_SumAse_Es2.Value = 0
         ipp_ValCom_Es2.Value = 0
         ipp_ValRea_Es2.Value = 0
         ipp_ValTer_Es2.Value = 0
         ipp_ValEdi_Es2.Value = 0
         ipp_ValACo_Es2.Value = 0
         
         tab_Genera.Tab = 3
         Call gs_SetFocus(cmb_FlgEst_Dep)
      End If
   End If
End Sub

Private Sub cmb_FlgEst_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_FlgEst_Es2_Click
   End If
End Sub

Private Sub cmb_MatCon_Click()
   Call gs_SetFocus(cmb_TipMon)
End Sub

Private Sub cmb_MatCon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_MatCon_Click
   End If
End Sub

Private Sub cmb_TipInm_Click()
   Call gs_SetFocus(cmb_UsoInm)
End Sub

Private Sub cmb_TipInm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipInm_Click
   End If
End Sub

Private Sub cmb_TipMon_Click()
   Call gs_SetFocus(ipp_TipCam)
End Sub

Private Sub cmb_TipMon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipMon_Click
   End If
End Sub

Private Sub cmb_UsoInm_Click()
   Call gs_SetFocus(cmb_MatCon)
End Sub

Private Sub cmb_UsoInm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_UsoInm_Click
   End If
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_dbl_Valmax_ValEdi    As Double
   Dim r_dbl_MtoPre           As Double
   Dim r_dbl_IntCap           As Double
   Dim r_dbl_ComVta           As Double
   Dim r_dbl_TipCam           As Double
   Dim r_dbl_TCaMPr           As Double
   Dim r_dbl_ValRea           As Double
   Dim r_dbl_PorMax_ValGar    As Double
   Dim r_dbl_PorMin_ValGrv    As Double
   Dim r_dbl_ValGar           As Double

   If cmb_EmpPer.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Empresa de Peritaje.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_EmpPer)
      Exit Sub
   End If
   
   If Len(Trim(txt_NomPer.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre del Perito.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomPer)
      Exit Sub
   End If

   If Len(Trim(txt_CodPer.Text)) = 0 Then
      MsgBox "Debe ingresar el Código REPEV SBS del Perito.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_CodPer)
      Exit Sub
   End If

   If Len(Trim(txt_NumInf.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Informe del Perito.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumInf)
      Exit Sub
   End If
   
   If CDate(ipp_FecEva.Text) > Date Then
      MsgBox "La Fecha de Evaluación no puede ser mayor a la Fecha Actual.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecEva)
      Exit Sub
   End If
      
   If ipp_AnoCon.Value = 0 Then
      MsgBox "Debe ingresar el Año de Construcción.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_AnoCon)
      Exit Sub
   End If
      
   If ipp_NumPis.Value = 0 Then
      MsgBox "Debe ingresar el Nro. de Pisos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_NumPis)
      Exit Sub
   End If
      
   If cmb_TipInm.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Inmueble.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipInm)
      Exit Sub
   End If
      
   If cmb_UsoInm.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Uso del Inmueble.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_UsoInm)
      Exit Sub
   End If
      
   If cmb_MatCon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Material de Construcción.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_MatCon)
      Exit Sub
   End If
      
   If cmb_TipMon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Moneda de la Valuación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipMon)
      Exit Sub
   End If
   
   If cmb_TipMon.ItemData(cmb_TipMon.ListIndex) <> moddat_g_int_TipMon Then
      MsgBox "La Moneda de la Tasación debe ser igual a la Moneda del Préstamo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipMon)
      Exit Sub
   End If
   
   If ipp_TipCam.Value = 0 Then
      MsgBox "Debe ingresar el Tipo de Cambio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_TipCam)
      Exit Sub
   End If
   
   If ipp_AreTer_Inm.Value = 0 Then
      MsgBox "Debe ingresar el Area del Terreno.", vbExclamation, modgen_g_str_NomPlt
      
      tab_Genera.Tab = 0
      Call gs_SetFocus(ipp_AreTer_Inm)
      Exit Sub
   End If

   If ipp_AreCon_Inm.Value = 0 Then
      MsgBox "Debe ingresar el Area Construida.", vbExclamation, modgen_g_str_NomPlt
      
      tab_Genera.Tab = 0
      Call gs_SetFocus(ipp_AreCon_Inm)
      Exit Sub
   End If

   If ipp_SumAse_Inm.Value = 0 Then
      MsgBox "Debe ingresar la Suma Asegurada.", vbExclamation, modgen_g_str_NomPlt
      
      tab_Genera.Tab = 0
      Call gs_SetFocus(ipp_SumAse_Inm)
      Exit Sub
   End If

   If ipp_ValCom_Inm.Value = 0 Then
      MsgBox "Debe ingresar el Valor Comercial.", vbExclamation, modgen_g_str_NomPlt
      
      tab_Genera.Tab = 0
      Call gs_SetFocus(ipp_ValCom_Inm)
      Exit Sub
   End If

   If ipp_ValRea_Inm.Value = 0 Then
      MsgBox "Debe ingresar el Valor de Realización.", vbExclamation, modgen_g_str_NomPlt
      
      tab_Genera.Tab = 0
      Call gs_SetFocus(ipp_ValRea_Inm)
      Exit Sub
   End If
   
   If ipp_ValTer_Inm.Value = 0 Then
      MsgBox "Debe ingresar el Valor del Terreno.", vbExclamation, modgen_g_str_NomPlt
      
      tab_Genera.Tab = 0
      Call gs_SetFocus(ipp_ValTer_Inm)
      Exit Sub
   End If
   
   If ipp_ValEdi_Inm.Value = 0 Then
      MsgBox "Debe ingresar el Valor de Edificación.", vbExclamation, modgen_g_str_NomPlt
      
      tab_Genera.Tab = 0
      Call gs_SetFocus(ipp_ValEdi_Inm)
      Exit Sub
   End If
   
'   If ipp_ValACo_Inm.Value = 0 Then
'      MsgBox "Debe ingresar el Valor de Areas Comunes.", vbExclamation, modgen_g_str_NomPlt
'
'      tab_Genera.Tab = 0
'      Call gs_SetFocus(ipp_ValACo_Inm)
'      Exit Sub
'   End If
   
   'Estacionamiento 1
   If cmb_FlgEst_Es1.ListIndex = -1 Then
      MsgBox "Debe seleccionar si hay valorización por Estacionamiento.", vbExclamation, modgen_g_str_NomPlt
      
      tab_Genera.Tab = 1
      Call gs_SetFocus(cmb_FlgEst_Es1)
      Exit Sub
   End If
   
   If cmb_FlgEst_Es1.ItemData(cmb_FlgEst_Es1.ListIndex) = 1 Then
      If ipp_AreTer_Es1.Value = 0 Then
         MsgBox "Debe ingresar el Area del Terreno.", vbExclamation, modgen_g_str_NomPlt
         
         tab_Genera.Tab = 1
         Call gs_SetFocus(ipp_AreTer_Es1)
         Exit Sub
      End If
   
      'If ipp_AreCon_Es1.Value = 0 Then
      '   MsgBox "Debe ingresar el Area Construida.", vbExclamation, modgen_g_str_NomPlt
      '
      '   tab_Genera.Tab = 1
      '   Call gs_SetFocus(ipp_AreCon_Es1)
      '   Exit Sub
      'End If
   
      'If ipp_SumAse_Es1.Value = 0 Then
      '   MsgBox "Debe ingresar la Suma Asegurada.", vbExclamation, modgen_g_str_NomPlt
      '
      '   tab_Genera.Tab = 1
      '   Call gs_SetFocus(ipp_SumAse_Es1)
      '   Exit Sub
      'End If
   
      If ipp_ValCom_Es1.Value = 0 Then
         MsgBox "Debe ingresar el Valor Comercial.", vbExclamation, modgen_g_str_NomPlt
         
         tab_Genera.Tab = 1
         Call gs_SetFocus(ipp_ValCom_Es1)
         Exit Sub
      End If
   
      If ipp_ValRea_Es1.Value = 0 Then
         MsgBox "Debe ingresar el Valor de Realización.", vbExclamation, modgen_g_str_NomPlt
         
         tab_Genera.Tab = 1
         Call gs_SetFocus(ipp_ValRea_Es1)
         Exit Sub
      End If
      
      If ipp_ValTer_Es1.Value = 0 Then
         MsgBox "Debe ingresar el Valor del Terreno.", vbExclamation, modgen_g_str_NomPlt
         
         tab_Genera.Tab = 1
         Call gs_SetFocus(ipp_ValTer_Es1)
         Exit Sub
      End If
      
      If ipp_ValEdi_Es1.Value = 0 Then
         MsgBox "Debe ingresar el Valor de Edificación.", vbExclamation, modgen_g_str_NomPlt
         
         tab_Genera.Tab = 1
         Call gs_SetFocus(ipp_ValEdi_Es1)
         Exit Sub
      End If
      
'      If ipp_ValACo_Es1.Value = 0 Then
'         MsgBox "Debe ingresar el Valor de Areas Comunes.", vbExclamation, modgen_g_str_NomPlt
'
'         tab_Genera.Tab = 1
'         Call gs_SetFocus(ipp_ValACo_Es1)
'         Exit Sub
'      End If
   End If
   
   'Estacionamiento 2
   If cmb_FlgEst_Es2.ListIndex = -1 Then
      MsgBox "Debe seleccionar si hay valorización por Estacionamiento.", vbExclamation, modgen_g_str_NomPlt
      
      tab_Genera.Tab = 2
      Call gs_SetFocus(cmb_FlgEst_Es2)
      Exit Sub
   End If
   
   If cmb_FlgEst_Es2.ItemData(cmb_FlgEst_Es2.ListIndex) = 1 Then
      If ipp_AreTer_Es2.Value = 0 Then
         MsgBox "Debe ingresar el Area del Terreno.", vbExclamation, modgen_g_str_NomPlt
         
         tab_Genera.Tab = 2
         Call gs_SetFocus(ipp_AreTer_Es2)
         Exit Sub
      End If
   
      'If ipp_AreCon_Es2.Value = 0 Then
      '   MsgBox "Debe ingresar el Area Construida.", vbExclamation, modgen_g_str_NomPlt
      '
      '   tab_Genera.Tab = 2
      '   Call gs_SetFocus(ipp_AreCon_Es2)
      '   Exit Sub
      'End If
   
      'If ipp_SumAse_Es2.Value = 0 Then
      '   MsgBox "Debe ingresar la Suma Asegurada.", vbExclamation, modgen_g_str_NomPlt
      '
      '   tab_Genera.Tab = 2
      '   Call gs_SetFocus(ipp_SumAse_Es2)
      '   Exit Sub
      'End If
   
      If ipp_ValCom_Es2.Value = 0 Then
         MsgBox "Debe ingresar el Valor Comercial.", vbExclamation, modgen_g_str_NomPlt
         
         tab_Genera.Tab = 2
         Call gs_SetFocus(ipp_ValCom_Es2)
         Exit Sub
      End If
   
      If ipp_ValRea_Es2.Value = 0 Then
         MsgBox "Debe ingresar el Valor de Realización.", vbExclamation, modgen_g_str_NomPlt
         
         tab_Genera.Tab = 2
         Call gs_SetFocus(ipp_ValRea_Es2)
         Exit Sub
      End If
      
      If ipp_ValTer_Es2.Value = 0 Then
         MsgBox "Debe ingresar el Valor del Terreno.", vbExclamation, modgen_g_str_NomPlt
         
         tab_Genera.Tab = 2
         Call gs_SetFocus(ipp_ValTer_Es2)
         Exit Sub
      End If
      
      If ipp_ValEdi_Es2.Value = 0 Then
         MsgBox "Debe ingresar el Valor de Edificación.", vbExclamation, modgen_g_str_NomPlt
         
         tab_Genera.Tab = 2
         Call gs_SetFocus(ipp_ValEdi_Es2)
         Exit Sub
      End If
      
'      If ipp_ValACo_Es2.Value = 0 Then
'         MsgBox "Debe ingresar el Valor de Areas Comunes.", vbExclamation, modgen_g_str_NomPlt
'
'         tab_Genera.Tab = 2
'         Call gs_SetFocus(ipp_ValACo_Es2)
'         Exit Sub
'      End If
   End If
   
   'Depósito
   If cmb_FlgEst_Dep.ListIndex = -1 Then
      MsgBox "Debe seleccionar si hay valorización por el Depósito.", vbExclamation, modgen_g_str_NomPlt
      
      tab_Genera.Tab = 3
      Call gs_SetFocus(cmb_FlgEst_Dep)
      Exit Sub
   End If
   
   If cmb_FlgEst_Dep.ItemData(cmb_FlgEst_Dep.ListIndex) = 1 Then
      If ipp_AreTer_Dep.Value = 0 Then
         MsgBox "Debe ingresar el Area del Terreno.", vbExclamation, modgen_g_str_NomPlt
         
         tab_Genera.Tab = 3
         Call gs_SetFocus(ipp_AreTer_Dep)
         Exit Sub
      End If
   
      'If ipp_AreCon_Dep.Value = 0 Then
      '   MsgBox "Debe ingresar el Area Construida.", vbExclamation, modgen_g_str_NomPlt
      '
      '   tab_Genera.Tab = 3
      '   Call gs_SetFocus(ipp_AreCon_Dep)
      '   Exit Sub
      'End If
   
      'If ipp_SumAse_Dep.Value = 0 Then
      '   MsgBox "Debe ingresar la Suma Asegurada.", vbExclamation, modgen_g_str_NomPlt
      '
      '   tab_Genera.Tab = 3
      '   Call gs_SetFocus(ipp_SumAse_Dep)
      '   Exit Sub
      'End If
   
      If ipp_ValCom_Dep.Value = 0 Then
         MsgBox "Debe ingresar el Valor Comercial.", vbExclamation, modgen_g_str_NomPlt
         
         tab_Genera.Tab = 3
         Call gs_SetFocus(ipp_ValCom_Dep)
         Exit Sub
      End If
   
      If ipp_ValRea_Dep.Value = 0 Then
         MsgBox "Debe ingresar el Valor de Realización.", vbExclamation, modgen_g_str_NomPlt
         
         tab_Genera.Tab = 3
         Call gs_SetFocus(ipp_ValRea_Dep)
         Exit Sub
      End If
      
      If ipp_ValTer_Dep.Value = 0 Then
         MsgBox "Debe ingresar el Valor del Terreno.", vbExclamation, modgen_g_str_NomPlt
         
         tab_Genera.Tab = 3
         Call gs_SetFocus(ipp_ValTer_Dep)
         Exit Sub
      End If
      
      If ipp_ValEdi_Dep.Value = 0 Then
         MsgBox "Debe ingresar el Valor de Edificación.", vbExclamation, modgen_g_str_NomPlt
         
         tab_Genera.Tab = 3
         Call gs_SetFocus(ipp_ValEdi_Dep)
         Exit Sub
      End If
      
'      If ipp_ValACo_Dep.Value = 0 Then
'         MsgBox "Debe ingresar el Valor de Areas Comunes.", vbExclamation, modgen_g_str_NomPlt
'
'         tab_Genera.Tab = 3
'         Call gs_SetFocus(ipp_ValACo_Dep)
'         Exit Sub
'      End If
   End If
   
   'Obteniendo Monto de Préstamo
   r_dbl_MtoPre = 0
   r_dbl_IntCap = 0
   r_dbl_ComVta = 0
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_dbl_MtoPre = g_rst_Princi!SOLMAE_MTOPRE_MPR
      r_dbl_IntCap = g_rst_Princi!SOLMAE_INTGRA
      
      If moddat_g_int_TipMon = 1 Then
         r_dbl_ComVta = g_rst_Princi!SOLMAE_COMVTA_SOL
      Else
         r_dbl_ComVta = g_rst_Princi!SOLMAE_COMVTA_DOL
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_dbl_TCaMPr = 0
   r_dbl_TipCam = 0
   
   If cmb_TipMon.ItemData(cmb_TipMon.ListIndex) <> 1 Then
      r_dbl_TipCam = CDbl(ipp_TipCam.Text)
   End If
   
   'Validando contra Parámetros de Productos
   If moddat_g_str_CodPrd = "001" Or moddat_g_str_CodPrd = "003" Then
      r_dbl_Valmax_ValEdi = 0
   
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "021") Then
         r_dbl_Valmax_ValEdi = l_arr_ParPrd(1).Genera_Cantid * moddat_gf_Consulta_ParVal("001", "002")
      End If
      
      'Validando Valor de Edificación
      If cmb_TipMon.ItemData(cmb_TipMon.ListIndex) <> 1 Then
         If CDbl(pnl_ValEdi.Caption) * r_dbl_TipCam > r_dbl_Valmax_ValEdi Then
            MsgBox "El Valor de Edificación excede el permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
      Else
         If CDbl(pnl_ValEdi.Caption) > r_dbl_Valmax_ValEdi Then
            MsgBox "El Valor de Edificación excede el permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
      End If
   End If
   
   If moddat_g_str_CodPrd = "001" Or moddat_g_str_CodPrd = "003" Or moddat_g_str_CodPrd = "004" Then
      'Validando Relación Valor de Gravamen Hipotecario (Valor Comercial) / Monto del Préstamo
      r_dbl_PorMin_ValGrv = 0
   
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "031") Then
         r_dbl_PorMin_ValGrv = l_arr_ParPrd(1).Genera_Cantid
      End If
   
      If Not (CDbl(pnl_ValCom.Caption) / r_dbl_MtoPre * 100 >= r_dbl_PorMin_ValGrv) Then
         MsgBox "La relación Valor Comercial / Monto de Préstamo no cumple con el Parámetro permitido para el Producto." & Format(CDbl(pnl_ValCom.Caption) / r_dbl_MtoPre * 100, "##0.00") & "%", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   
      'Determinando Valor de la Garantía
      If CDbl(pnl_ValCom.Caption) < r_dbl_ComVta Then
         r_dbl_ValGar = CDbl(pnl_ValCom.Caption)
      Else
         r_dbl_ValGar = r_dbl_ComVta
      End If
      
      'Validando Relación Monto del Préstamo / Valor de la Garantía
      r_dbl_PorMax_ValGar = 0
   
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "032") Then
         r_dbl_PorMax_ValGar = l_arr_ParPrd(1).Genera_Cantid
      End If
   
      If Not (r_dbl_MtoPre / r_dbl_ValGar * 100 <= r_dbl_PorMax_ValGar) Then
         MsgBox "La relación Monto del Préstamo / Valor de la Garantía no cumple con el Parámetro permitido para el Producto." & Format(r_dbl_MtoPre / r_dbl_ValGar * 100, "##0.00") & "%", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   
      'Validando Relación Monto del Préstamo + Interés Capitalizado / Valor de la Garantía
      r_dbl_PorMax_ValGar = 0
   
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "033") Then
         r_dbl_PorMax_ValGar = l_arr_ParPrd(1).Genera_Cantid
      End If
   
      If Not ((r_dbl_MtoPre + r_dbl_IntCap) / r_dbl_ValGar * 100 <= r_dbl_PorMax_ValGar) Then
         MsgBox "La relación Monto del Préstamo + Intereses Capitalizados / Valor de la Garantía no cumple con el Parámetro permitido para el Producto." & Format((r_dbl_MtoPre + r_dbl_IntCap) / r_dbl_ValGar * 100, "##0.00") & "%", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   End If
   
   'miCasita - Validando Valor de Realización contra Monto de Préstamo
   If moddat_g_str_CodPrd = "002" Then
      If r_dbl_MtoPre > CDbl(pnl_ValRea.Caption) Then
         MsgBox "El Valor de Realización es menor al Monto del Préstamo, no se puede otorgar este crédito.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      
      Call moddat_gs_FecSis
      
      g_str_Parame = "USP_TRA_EVATAS ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_EmpPer(cmb_EmpPer.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NomPer.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_CodPer.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NumInf.Text & "', "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecEva.Text), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AnoCon.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_NumPis.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_NumSot.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_TipInm.ItemData(cmb_TipInm.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_UsoInm.ItemData(cmb_UsoInm.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_MatCon.ItemData(cmb_MatCon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_TipMon.ItemData(cmb_TipMon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_TipCam.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_TCaMPr) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AreTer_Inm.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AreCon_Inm.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_SumAse_Inm.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValCom_Inm.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValRea_Inm.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValTer_Inm.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValEdi_Inm.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValACo_Inm.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_FlgEst_Es1.ItemData(cmb_FlgEst_Es1.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AreTer_Es1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AreCon_Es1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_SumAse_Es1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValCom_Es1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValRea_Es1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValTer_Es1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValEdi_Es1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValACo_Es1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_FlgEst_Es2.ItemData(cmb_FlgEst_Es2.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AreTer_Es2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AreCon_Es2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_SumAse_Es2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValCom_Es2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValRea_Es2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValTer_Es2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValEdi_Es2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValACo_Es2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_FlgEst_Dep.ItemData(cmb_FlgEst_Dep.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AreTer_Dep.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AreCon_Dep.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_SumAse_Dep.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValCom_Dep.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValRea_Dep.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValTer_Dep.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValEdi_Dep.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValACo_Dep.Value) & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ") "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      Screen.MousePointer = 0
   Loop

   moddat_g_int_FlgAct = 2
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_Produc.Caption = moddat_g_str_NomPrd
   pnl_FecIng.Caption = moddat_g_str_FecIng
   pnl_IngIns.Caption = moddat_gf_FecIng_Ins(moddat_g_str_NumSol, 41)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicia
   Call fs_Buscar_DatEva
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte(cmb_EmpPer, l_arr_EmpPer, 1, "507")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipInm, 1, "221")
   Call moddat_gs_Carga_LisIte_Combo(cmb_UsoInm, 1, "222")
   Call moddat_gs_Carga_LisIte_Combo(cmb_MatCon, 1, "223")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipMon, 1, "204")
   Call moddat_gs_Carga_LisIte_Combo(cmb_FlgEst_Es1, 1, "214")
   Call moddat_gs_Carga_LisIte_Combo(cmb_FlgEst_Es2, 1, "214")
   Call moddat_gs_Carga_LisIte_Combo(cmb_FlgEst_Dep, 1, "214")

   cmb_EmpPer.ListIndex = -1
   txt_NomPer.Text = ""
   txt_CodPer.Text = ""
   txt_NumInf.Text = ""
   ipp_FecEva.Text = Format(Date, "dd/mm/yyyy")
   cmb_TipMon.ListIndex = -1
   ipp_TipCam.Value = 0
   
   ipp_AreTer_Inm.Value = 0
   ipp_AreCon_Inm.Value = 0
   ipp_SumAse_Inm.Value = 0
   ipp_ValCom_Inm.Value = 0
   ipp_ValRea_Inm.Value = 0
   ipp_ValTer_Inm.Value = 0
   ipp_ValEdi_Inm.Value = 0
   ipp_ValACo_Inm.Value = 0

   cmb_FlgEst_Es1.ListIndex = -1
   ipp_AreTer_Es1.Value = 0
   ipp_AreCon_Es1.Value = 0
   ipp_SumAse_Es1.Value = 0
   ipp_ValCom_Es1.Value = 0
   ipp_ValRea_Es1.Value = 0
   ipp_ValTer_Es1.Value = 0
   ipp_ValEdi_Es1.Value = 0
   ipp_ValACo_Es1.Value = 0

   ipp_AreTer_Es1.Enabled = False
   ipp_AreCon_Es1.Enabled = False
   ipp_SumAse_Es1.Enabled = False
   ipp_ValCom_Es1.Enabled = False
   ipp_ValRea_Es1.Enabled = False
   ipp_ValTer_Es1.Enabled = False
   ipp_ValEdi_Es1.Enabled = False
   ipp_ValACo_Es1.Enabled = False

   cmb_FlgEst_Es2.ListIndex = -1
   ipp_AreTer_Es2.Value = 0
   ipp_AreCon_Es2.Value = 0
   ipp_SumAse_Es2.Value = 0
   ipp_ValCom_Es2.Value = 0
   ipp_ValRea_Es2.Value = 0
   ipp_ValTer_Es2.Value = 0
   ipp_ValEdi_Es2.Value = 0
   ipp_ValACo_Es2.Value = 0

   ipp_AreTer_Es2.Enabled = False
   ipp_AreCon_Es2.Enabled = False
   ipp_SumAse_Es2.Enabled = False
   ipp_ValCom_Es2.Enabled = False
   ipp_ValRea_Es2.Enabled = False
   ipp_ValTer_Es2.Enabled = False
   ipp_ValEdi_Es2.Enabled = False
   ipp_ValACo_Es2.Enabled = False

   cmb_FlgEst_Dep.ListIndex = -1
   ipp_AreTer_Dep.Value = 0
   ipp_AreCon_Dep.Value = 0
   ipp_SumAse_Dep.Value = 0
   ipp_ValCom_Dep.Value = 0
   ipp_ValRea_Dep.Value = 0
   ipp_ValTer_Dep.Value = 0
   ipp_ValEdi_Dep.Value = 0
   ipp_ValACo_Dep.Value = 0

   ipp_AreTer_Dep.Enabled = False
   ipp_AreCon_Dep.Enabled = False
   ipp_SumAse_Dep.Enabled = False
   ipp_ValCom_Dep.Enabled = False
   ipp_ValRea_Dep.Enabled = False
   ipp_ValTer_Dep.Enabled = False
   ipp_ValEdi_Dep.Enabled = False
   ipp_ValACo_Dep.Enabled = False
End Sub

Private Sub ipp_AnoCon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_NumPis)
   End If
End Sub

Private Sub ipp_AreCon_Inm_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_AreCon_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_SumAse_Inm)
   End If
End Sub

Private Sub ipp_AreTer_Inm_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_AreTer_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AreCon_Inm)
   End If
End Sub

Private Sub ipp_NumPis_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_NumSot)
   End If
End Sub

Private Sub ipp_NumSot_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipInm)
   End If
End Sub

Private Sub ipp_SumAse_Inm_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_SumAse_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValCom_Inm)
   End If
End Sub

Private Sub ipp_ValCom_Inm_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValCom_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValRea_Inm)
   End If
End Sub

Private Sub ipp_ValRea_Inm_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValRea_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValTer_Inm)
   End If
End Sub

Private Sub ipp_ValTer_Inm_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValTer_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValEdi_Inm)
   End If
End Sub

Private Sub ipp_ValEdi_Inm_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValEdi_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValACo_Inm)
   End If
End Sub

Private Sub ipp_ValACo_Inm_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValACo_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      tab_Genera.Tab = 1
      Call gs_SetFocus(cmb_FlgEst_Es1)
   End If
End Sub

Private Sub ipp_AreCon_Es1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_AreCon_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_SumAse_Es1)
   End If
End Sub

Private Sub ipp_AreTer_Es1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_AreTer_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AreCon_Es1)
   End If
End Sub

Private Sub ipp_SumAse_Es1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_SumAse_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValCom_Es1)
   End If
End Sub

Private Sub ipp_ValCom_Es1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValCom_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValRea_Es1)
   End If
End Sub

Private Sub ipp_ValRea_Es1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValRea_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValTer_Es1)
   End If
End Sub

Private Sub ipp_ValTer_Es1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValTer_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValEdi_Es1)
   End If
End Sub

Private Sub ipp_ValEdi_Es1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValEdi_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValACo_Es1)
   End If
End Sub

Private Sub ipp_ValACo_Es1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValACo_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      tab_Genera.Tab = 2
      Call gs_SetFocus(cmb_FlgEst_Es2)
   End If
End Sub

Private Sub ipp_AreCon_Es2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_AreCon_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_SumAse_Es2)
   End If
End Sub

Private Sub ipp_AreTer_Es2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_AreTer_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AreCon_Es2)
   End If
End Sub

Private Sub ipp_SumAse_Es2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_SumAse_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValCom_Es2)
   End If
End Sub

Private Sub ipp_ValCom_Es2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValCom_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValRea_Es2)
   End If
End Sub

Private Sub ipp_ValRea_Es2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValRea_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValTer_Es2)
   End If
End Sub

Private Sub ipp_ValTer_Es2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValTer_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValEdi_Es2)
   End If
End Sub

Private Sub ipp_ValEdi_Es2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValEdi_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValACo_Es2)
   End If
End Sub

Private Sub ipp_ValACo_Es2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValACo_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      tab_Genera.Tab = 1
      Call gs_SetFocus(cmb_FlgEst_Dep)
   End If
End Sub

Private Sub ipp_AreCon_Dep_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_AreCon_Dep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_SumAse_Dep)
   End If
End Sub

Private Sub ipp_AreTer_Dep_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_AreTer_Dep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AreCon_Dep)
   End If
End Sub

Private Sub ipp_SumAse_Dep_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_SumAse_Dep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValCom_Dep)
   End If
End Sub

Private Sub ipp_ValCom_Dep_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValCom_Dep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValRea_Dep)
   End If
End Sub

Private Sub ipp_ValRea_Dep_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValRea_Dep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValTer_Dep)
   End If
End Sub

Private Sub ipp_ValTer_Dep_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValTer_Dep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValEdi_Dep)
   End If
End Sub

Private Sub ipp_ValEdi_Dep_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValEdi_Dep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValACo_Dep)
   End If
End Sub

Private Sub ipp_ValACo_Dep_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValACo_Dep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      tab_Genera.Tab = 1
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub ipp_FecEva_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AnoCon)
   End If
End Sub

Private Sub ipp_TipCam_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      tab_Genera.Tab = 0
      Call gs_SetFocus(ipp_AreTer_Inm)
   End If
End Sub

Private Sub txt_CodPer_GotFocus()
   Call gs_SelecTodo(txt_CodPer)
End Sub

Private Sub txt_CodPer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumInf)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ .")
   End If
End Sub

Private Sub txt_NomPer_GotFocus()
   Call gs_SelecTodo(txt_NomPer)
End Sub

Private Sub txt_NomPer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_CodPer)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "-_ .")
   End If
End Sub

Private Sub txt_NumInf_GotFocus()
   Call gs_SelecTodo(txt_NumInf)
End Sub

Private Sub txt_NumInf_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecEva)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-/():;:;")
   End If
End Sub

Private Sub fs_Calcul()
   pnl_AreTer.Caption = Format(CDbl(ipp_AreTer_Inm.Text) + CDbl(ipp_AreTer_Es1.Text) + CDbl(ipp_AreTer_Es2.Text) + CDbl(ipp_AreTer_Dep.Text), "###,###,##0.00") & " "
   pnl_AreCon.Caption = Format(CDbl(ipp_AreCon_Inm.Text) + CDbl(ipp_AreCon_Es1.Text) + CDbl(ipp_AreCon_Es2.Text) + CDbl(ipp_AreCon_Dep.Text), "###,###,##0.00") & " "
   
   pnl_SumAse.Caption = Format(CDbl(ipp_SumAse_Inm.Text) + CDbl(ipp_SumAse_Es1.Text) + CDbl(ipp_SumAse_Es2.Text) + CDbl(ipp_SumAse_Dep.Text), "###,###,##0.00") & " "
   pnl_ValCom.Caption = Format(CDbl(ipp_ValCom_Inm.Text) + CDbl(ipp_ValCom_Es1.Text) + CDbl(ipp_ValCom_Es2.Text) + CDbl(ipp_ValCom_Dep.Text), "###,###,##0.00") & " "
   pnl_ValRea.Caption = Format(CDbl(ipp_ValRea_Inm.Text) + CDbl(ipp_ValRea_Es1.Text) + CDbl(ipp_ValRea_Es2.Text) + CDbl(ipp_ValRea_Dep.Text), "###,###,##0.00") & " "
   pnl_ValTer.Caption = Format(CDbl(ipp_ValTer_Inm.Text) + CDbl(ipp_ValTer_Es1.Text) + CDbl(ipp_ValTer_Es2.Text) + CDbl(ipp_ValTer_Dep.Text), "###,###,##0.00") & " "
   pnl_ValEdi.Caption = Format(CDbl(ipp_ValEdi_Inm.Text) + CDbl(ipp_ValEdi_Es1.Text) + CDbl(ipp_ValEdi_Es2.Text) + CDbl(ipp_ValEdi_Dep.Text), "###,###,##0.00") & " "
   pnl_ValACo.Caption = Format(CDbl(ipp_ValACo_Inm.Text) + CDbl(ipp_ValACo_Es1.Text) + CDbl(ipp_ValACo_Es2.Text) + CDbl(ipp_ValACo_Dep.Text), "###,###,##0.00") & " "
End Sub

Private Sub fs_Buscar_DatEva()
   moddat_g_int_FlgGrb = 1
   
   g_str_Parame = "SELECT * FROM TRA_EVATAS WHERE "
   g_str_Parame = g_str_Parame & "EVATAS_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      moddat_g_int_FlgGrb = 2
         
      cmb_EmpPer.ListIndex = gf_Busca_Arregl(l_arr_EmpPer, g_rst_Princi!EVATAS_CODEMP) - 1
      txt_NomPer.Text = Trim(g_rst_Princi!EVATAS_NOMPER & "")
      txt_CodPer.Text = Trim(g_rst_Princi!EVATAS_CODPER & "")
      txt_NumInf.Text = Trim(g_rst_Princi!EVATAS_NUMINF & "")
      ipp_FecEva.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVATAS_FECEVA))
      
      ipp_AnoCon.Value = g_rst_Princi!EVATAS_ANOCON
      ipp_NumPis.Value = g_rst_Princi!EVATAS_NUMPIS
      ipp_NumSot.Value = g_rst_Princi!EVATAS_NUMSOT
      
      Call gs_BuscarCombo_Item(cmb_TipInm, g_rst_Princi!EVATAS_TIPINM)
      Call gs_BuscarCombo_Item(cmb_UsoInm, g_rst_Princi!EVATAS_USOINM)
      Call gs_BuscarCombo_Item(cmb_MatCon, g_rst_Princi!EVATAS_MATCON)
      
      Call gs_BuscarCombo_Item(cmb_TipMon, g_rst_Princi!EVATAS_TIPMON)
      
      ipp_TipCam.Value = g_rst_Princi!EVATAS_TIPCAM
      
      ipp_AreTer_Inm.Value = g_rst_Princi!EVATAS_ARETER_INM
      ipp_AreCon_Inm.Value = g_rst_Princi!EVATAS_ARECON_INM
      ipp_SumAse_Inm.Value = g_rst_Princi!EVATAS_SUMASE_INM
      ipp_ValCom_Inm.Value = g_rst_Princi!EVATAS_VALCOM_INM
      ipp_ValRea_Inm.Value = g_rst_Princi!EVATAS_VALREA_INM
      ipp_ValTer_Inm.Value = g_rst_Princi!EVATAS_VALTER_INM
      ipp_ValEdi_Inm.Value = g_rst_Princi!EVATAS_VALEDI_INM
      ipp_ValACo_Inm.Value = g_rst_Princi!EVATAS_VALACO_INM
      
      Call gs_BuscarCombo_Item(cmb_FlgEst_Es1, g_rst_Princi!EVATAS_FLGEST_ES1)
      
      If cmb_FlgEst_Es1.ItemData(cmb_FlgEst_Es1.ListIndex) = 1 Then
         ipp_AreTer_Es1.Value = g_rst_Princi!EVATAS_ARETER_ES1
         ipp_AreCon_Es1.Value = g_rst_Princi!EVATAS_ARECON_ES1
         ipp_SumAse_Es1.Value = g_rst_Princi!EVATAS_SUMASE_ES1
         ipp_ValCom_Es1.Value = g_rst_Princi!EVATAS_VALCOM_ES1
         ipp_ValRea_Es1.Value = g_rst_Princi!EVATAS_VALREA_ES1
         ipp_ValTer_Es1.Value = g_rst_Princi!EVATAS_VALTER_ES1
         ipp_ValEdi_Es1.Value = g_rst_Princi!EVATAS_VALEDI_ES1
         ipp_ValACo_Es1.Value = g_rst_Princi!EVATAS_VALACO_ES1
      
         ipp_AreTer_Es1.Enabled = True
         ipp_AreCon_Es1.Enabled = True
         ipp_SumAse_Es1.Enabled = True
         ipp_ValCom_Es1.Enabled = True
         ipp_ValRea_Es1.Enabled = True
         ipp_ValTer_Es1.Enabled = True
         ipp_ValEdi_Es1.Enabled = True
         ipp_ValACo_Es1.Enabled = True
      End If
   
      Call gs_BuscarCombo_Item(cmb_FlgEst_Es2, g_rst_Princi!EVATAS_FLGEST_ES2)
      
      If cmb_FlgEst_Es2.ItemData(cmb_FlgEst_Es2.ListIndex) = 1 Then
         ipp_AreTer_Es2.Value = g_rst_Princi!EVATAS_ARETER_ES2
         ipp_AreCon_Es2.Value = g_rst_Princi!EVATAS_ARECON_ES2
         ipp_SumAse_Es2.Value = g_rst_Princi!EVATAS_SUMASE_ES2
         ipp_ValCom_Es2.Value = g_rst_Princi!EVATAS_VALCOM_ES2
         ipp_ValRea_Es2.Value = g_rst_Princi!EVATAS_VALREA_ES2
         ipp_ValTer_Es2.Value = g_rst_Princi!EVATAS_VALTER_ES2
         ipp_ValEdi_Es2.Value = g_rst_Princi!EVATAS_VALEDI_ES2
         ipp_ValACo_Es2.Value = g_rst_Princi!EVATAS_VALACO_ES2
      
         ipp_AreTer_Es2.Enabled = True
         ipp_AreCon_Es2.Enabled = True
         ipp_SumAse_Es2.Enabled = True
         ipp_ValCom_Es2.Enabled = True
         ipp_ValRea_Es2.Enabled = True
         ipp_ValTer_Es2.Enabled = True
         ipp_ValEdi_Es2.Enabled = True
         ipp_ValACo_Es2.Enabled = True
      End If
      
      Call gs_BuscarCombo_Item(cmb_FlgEst_Dep, g_rst_Princi!EVATAS_FLGEST_DEP)
      
      If cmb_FlgEst_Dep.ItemData(cmb_FlgEst_Dep.ListIndex) = 1 Then
         ipp_AreTer_Dep.Value = g_rst_Princi!EVATAS_ARETER_DEP
         ipp_AreCon_Dep.Value = g_rst_Princi!EVATAS_ARECON_DEP
         ipp_SumAse_Dep.Value = g_rst_Princi!EVATAS_SUMASE_DEP
         ipp_ValCom_Dep.Value = g_rst_Princi!EVATAS_VALCOM_DEP
         ipp_ValRea_Dep.Value = g_rst_Princi!EVATAS_VALREA_DEP
         ipp_ValTer_Dep.Value = g_rst_Princi!EVATAS_VALTER_DEP
         ipp_ValEdi_Dep.Value = g_rst_Princi!EVATAS_VALEDI_DEP
         ipp_ValACo_Dep.Value = g_rst_Princi!EVATAS_VALACO_DEP
      
         ipp_AreTer_Dep.Enabled = True
         ipp_AreCon_Dep.Enabled = True
         ipp_SumAse_Dep.Enabled = True
         ipp_ValCom_Dep.Enabled = True
         ipp_ValRea_Dep.Enabled = True
         ipp_ValTer_Dep.Enabled = True
         ipp_ValEdi_Dep.Enabled = True
         ipp_ValACo_Dep.Enabled = True
      End If
      
      Call fs_Calcul
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   tab_Genera.Tab = 0
End Sub
