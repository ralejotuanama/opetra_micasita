VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Ges_TecPro_06 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13770
   Icon            =   "OpeTra_frm_822.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10485
   ScaleWidth      =   13770
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10575
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   13815
      _Version        =   65536
      _ExtentX        =   24368
      _ExtentY        =   18653
      _StockProps     =   15
      BackColor       =   14215660
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
         Height          =   10515
         Left            =   -30
         TabIndex        =   17
         Top             =   30
         Width           =   13785
         _Version        =   65536
         _ExtentX        =   24315
         _ExtentY        =   18547
         _StockProps     =   15
         BackColor       =   14215660
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
            Height          =   3075
            Left            =   30
            TabIndex        =   51
            Top             =   7350
            Width           =   13695
            _Version        =   65536
            _ExtentX        =   24156
            _ExtentY        =   5424
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
            Begin VB.TextBox txt_NumMov 
               Height          =   315
               Left            =   2280
               MaxLength       =   25
               TabIndex        =   64
               Top             =   1890
               Width           =   2175
            End
            Begin VB.TextBox txt_Refer 
               Height          =   555
               Left            =   2280
               MaxLength       =   100
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   63
               Top             =   2250
               Width           =   11190
            End
            Begin VB.ComboBox cmb_CtaBan 
               Enabled         =   0   'False
               Height          =   315
               Left            =   2280
               Style           =   2  'Dropdown List
               TabIndex        =   60
               Top             =   1530
               Width           =   4635
            End
            Begin VB.ComboBox cmb_CodBan 
               Enabled         =   0   'False
               Height          =   315
               Left            =   2280
               Style           =   2  'Dropdown List
               TabIndex        =   59
               Top             =   1170
               Width           =   4635
            End
            Begin VB.TextBox txt_NumDoc 
               Height          =   315
               Left            =   2280
               MaxLength       =   25
               TabIndex        =   7
               Top             =   450
               Width           =   4635
            End
            Begin VB.ComboBox cmb_TipDoc 
               Height          =   315
               Left            =   2280
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   120
               Width           =   4635
            End
            Begin Threed.SSPanel pnl_CCIBan 
               Height          =   315
               Left            =   9480
               TabIndex        =   52
               Top             =   1530
               Width           =   3975
               _Version        =   65536
               _ExtentX        =   7011
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
            Begin Threed.SSPanel pnl_RScPrv 
               Height          =   315
               Left            =   2280
               TabIndex        =   53
               Top             =   810
               Width           =   11205
               _Version        =   65536
               _ExtentX        =   19764
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
            Begin VB.Label Label14 
               Caption         =   "Num. Movimiento:"
               Height          =   225
               Left            =   150
               TabIndex        =   66
               Top             =   1935
               Width           =   1365
            End
            Begin VB.Label Label15 
               Caption         =   "Referencia:"
               Height          =   195
               Left            =   150
               TabIndex        =   65
               Top             =   2430
               Width           =   1245
            End
            Begin VB.Label Label3 
               Caption         =   "Cuenta Corriente:"
               Height          =   255
               Left            =   120
               TabIndex        =   62
               Top             =   1560
               Width           =   1245
            End
            Begin VB.Label Label1 
               Caption         =   "Entidad Financiera:"
               Height          =   225
               Left            =   120
               TabIndex        =   61
               Top             =   1215
               Width           =   1485
            End
            Begin VB.Label Label6 
               Caption         =   "CCI:"
               Height          =   225
               Left            =   7920
               TabIndex        =   57
               Top             =   1560
               Width           =   1155
            End
            Begin VB.Label Label11 
               Caption         =   "Nro. Doc. Proveedor:"
               Height          =   225
               Left            =   120
               TabIndex        =   56
               Top             =   495
               Width           =   1845
            End
            Begin VB.Label Label12 
               Caption         =   "Razón Social:"
               Height          =   255
               Left            =   120
               TabIndex        =   55
               Top             =   840
               Width           =   1155
            End
            Begin VB.Label Label13 
               Caption         =   "Tipo Doc. Proveedor:"
               Height          =   255
               Left            =   120
               TabIndex        =   54
               Top             =   150
               Width           =   1605
            End
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   675
            Left            =   30
            TabIndex        =   18
            Top             =   720
            Width           =   13695
            _Version        =   65536
            _ExtentX        =   24156
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
            Begin VB.CommandButton cmd_ExpLiq 
               Height          =   585
               Left            =   1830
               Picture         =   "OpeTra_frm_822.frx":000C
               Style           =   1  'Graphical
               TabIndex        =   12
               ToolTipText     =   "Imprimir Orden de Trabajo"
               Top             =   30
               Width           =   585
            End
            Begin VB.CommandButton cmd_Salida 
               Height          =   585
               Left            =   13080
               Picture         =   "OpeTra_frm_822.frx":044E
               Style           =   1  'Graphical
               TabIndex        =   14
               ToolTipText     =   "Salir"
               Top             =   30
               Width           =   585
            End
            Begin VB.CommandButton cmd_ExpExc 
               Height          =   585
               Left            =   1230
               Picture         =   "OpeTra_frm_822.frx":0890
               Style           =   1  'Graphical
               TabIndex        =   11
               ToolTipText     =   "Exportar a Excel"
               Top             =   30
               Width           =   585
            End
            Begin VB.CommandButton cmd_Cancel 
               Height          =   585
               Left            =   12480
               Picture         =   "OpeTra_frm_822.frx":0B9A
               Style           =   1  'Graphical
               TabIndex        =   13
               ToolTipText     =   "Cancelar"
               Top             =   30
               Width           =   585
            End
            Begin VB.CommandButton cmd_Grabar 
               Height          =   585
               Left            =   11880
               Picture         =   "OpeTra_frm_822.frx":0EA4
               Style           =   1  'Graphical
               TabIndex        =   8
               ToolTipText     =   "Grabar Datos"
               Top             =   30
               Width           =   585
            End
            Begin VB.CommandButton cmd_Borrar 
               Height          =   585
               Left            =   630
               Picture         =   "OpeTra_frm_822.frx":12E6
               Style           =   1  'Graphical
               TabIndex        =   10
               ToolTipText     =   "Borrar Registro"
               Top             =   30
               Width           =   585
            End
            Begin VB.CommandButton cmd_Editar 
               Height          =   585
               Left            =   4650
               Picture         =   "OpeTra_frm_822.frx":15F0
               Style           =   1  'Graphical
               TabIndex        =   15
               ToolTipText     =   "Modificar Registro"
               Top             =   30
               Visible         =   0   'False
               Width           =   585
            End
            Begin VB.CommandButton cmd_Agrega 
               Height          =   585
               Left            =   30
               Picture         =   "OpeTra_frm_822.frx":18FA
               Style           =   1  'Graphical
               TabIndex        =   9
               ToolTipText     =   "Nuevo Registro"
               Top             =   30
               Width           =   585
            End
         End
         Begin Threed.SSPanel SSPanel11 
            Height          =   1965
            Left            =   30
            TabIndex        =   19
            Top             =   5340
            Width           =   13695
            _Version        =   65536
            _ExtentX        =   24156
            _ExtentY        =   3466
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
            Begin VB.ComboBox cmb_ForPag 
               Enabled         =   0   'False
               Height          =   315
               Left            =   2310
               Style           =   2  'Dropdown List
               TabIndex        =   2
               Top             =   480
               Width           =   4635
            End
            Begin VB.ComboBox cmb_Moneda 
               Height          =   315
               ItemData        =   "OpeTra_frm_822.frx":1C04
               Left            =   9480
               List            =   "OpeTra_frm_822.frx":1C06
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   1560
               Width           =   4035
            End
            Begin VB.ComboBox cmb_TipOper 
               Height          =   315
               Left            =   2910
               Style           =   2  'Dropdown List
               TabIndex        =   0
               Top             =   120
               Width           =   4035
            End
            Begin EditLib.fpDoubleSingle ipp_ImpDes 
               Height          =   315
               Left            =   2310
               TabIndex        =   4
               Top             =   1560
               Width           =   1635
               _Version        =   196608
               _ExtentX        =   2893
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
            Begin Threed.SSPanel pnl_ImpSal 
               Height          =   315
               Left            =   2310
               TabIndex        =   20
               Top             =   1200
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
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
            Begin Threed.SSPanel pnl_CodOpe 
               Height          =   315
               Left            =   2310
               TabIndex        =   21
               Top             =   120
               Width           =   555
               _Version        =   65536
               _ExtentX        =   979
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
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
            Begin Threed.SSPanel pnl_ImpTot 
               Height          =   315
               Left            =   2310
               TabIndex        =   45
               Top             =   840
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "0.00  "
               ForeColor       =   32768
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
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
            Begin EditLib.fpDoubleSingle ipp_PorDes 
               Height          =   315
               Left            =   5820
               TabIndex        =   3
               Top             =   840
               Width           =   1125
               _Version        =   196608
               _ExtentX        =   1984
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
            Begin EditLib.fpDateTime ipp_FecOpe 
               Height          =   315
               Left            =   9480
               TabIndex        =   1
               Top             =   120
               Width           =   1635
               _Version        =   196608
               _ExtentX        =   2884
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
            Begin VB.Label Label17 
               Caption         =   "Forma de Pago:"
               Height          =   225
               Left            =   120
               TabIndex        =   58
               Top             =   525
               Width           =   1245
            End
            Begin VB.Label Label10 
               Caption         =   "% de Desembolso:"
               Height          =   255
               Left            =   4350
               TabIndex        =   46
               Top             =   870
               Width           =   1395
            End
            Begin VB.Label Label9 
               Caption         =   "Monto Desembolsado:"
               Height          =   285
               Left            =   120
               TabIndex        =   44
               Top             =   1575
               Width           =   1755
            End
            Begin VB.Label Label8 
               Caption         =   "Saldo:"
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   1230
               Width           =   1395
            End
            Begin VB.Label Label5 
               Caption         =   "Valor:"
               Height          =   255
               Left            =   120
               TabIndex        =   42
               Top             =   870
               Width           =   1365
            End
            Begin VB.Label Label4 
               Caption         =   "Moneda:"
               Height          =   225
               Left            =   7920
               TabIndex        =   24
               Top             =   1605
               Width           =   1545
            End
            Begin VB.Label Label2 
               Caption         =   "Tipo de Operación:"
               Height          =   285
               Left            =   120
               TabIndex        =   23
               Top             =   135
               Width           =   1395
            End
            Begin VB.Label Label16 
               Caption         =   "Fecha Operación:"
               Height          =   285
               Left            =   7920
               TabIndex        =   22
               Top             =   135
               Width           =   1305
            End
         End
         Begin Threed.SSPanel SSPanel6 
            Height          =   675
            Left            =   30
            TabIndex        =   25
            Top             =   30
            Width           =   13695
            _Version        =   65536
            _ExtentX        =   24156
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
               Left            =   630
               TabIndex        =   26
               Top             =   30
               Width           =   4215
               _Version        =   65536
               _ExtentX        =   7435
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Gestión de Crédito Hipotecario"
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
            Begin Threed.SSPanel SSPanel15 
               Height          =   315
               Left            =   630
               TabIndex        =   27
               Top             =   330
               Width           =   4215
               _Version        =   65536
               _ExtentX        =   7435
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Techo Propio - Gestión de Operaciones"
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
               Picture         =   "OpeTra_frm_822.frx":1C08
               Top             =   60
               Width           =   480
            End
         End
         Begin Threed.SSPanel SSPanel9 
            Height          =   1155
            Left            =   30
            TabIndex        =   28
            Top             =   1440
            Width           =   13695
            _Version        =   65536
            _ExtentX        =   24156
            _ExtentY        =   2037
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
            Begin Threed.SSPanel pnl_RazSoc 
               Height          =   315
               Left            =   1620
               TabIndex        =   29
               Top             =   450
               Width           =   5685
               _Version        =   65536
               _ExtentX        =   10028
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
            Begin Threed.SSPanel pnl_TipDoc 
               Height          =   315
               Left            =   1620
               TabIndex        =   30
               Top             =   120
               Width           =   5685
               _Version        =   65536
               _ExtentX        =   10028
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
            Begin Threed.SSPanel pnl_NroDoc 
               Height          =   315
               Left            =   9690
               TabIndex        =   31
               Top             =   120
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
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
            Begin Threed.SSPanel pnl_TipEmp 
               Height          =   315
               Left            =   1620
               TabIndex        =   32
               Top             =   780
               Width           =   2955
               _Version        =   65536
               _ExtentX        =   5212
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
            Begin Threed.SSPanel pnl_NumRef 
               Height          =   315
               Left            =   9690
               TabIndex        =   33
               Top             =   450
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
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
            Begin VB.Label lbl_TipEmp 
               Caption         =   "Tipo Empresa:"
               Height          =   255
               Left            =   120
               TabIndex        =   38
               Top             =   810
               Width           =   1335
            End
            Begin VB.Label lbl_RazSoc 
               Caption         =   "Razón Social:"
               Height          =   255
               Left            =   120
               TabIndex        =   37
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label lbl_NumDoc 
               Caption         =   "Nro. Documento:"
               Height          =   225
               Left            =   8100
               TabIndex        =   36
               Top             =   150
               Width           =   1335
            End
            Begin VB.Label lbl_TipDoc 
               Caption         =   "Tipo Documento:"
               Height          =   255
               Left            =   120
               TabIndex        =   35
               Top             =   150
               Width           =   1335
            End
            Begin VB.Label Label7 
               Caption         =   "Nro. Referencia:"
               Height          =   255
               Left            =   8100
               TabIndex        =   34
               Top             =   480
               Width           =   1335
            End
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   2655
            Left            =   30
            TabIndex        =   39
            Top             =   2640
            Width           =   13695
            _Version        =   65536
            _ExtentX        =   24156
            _ExtentY        =   4683
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
               Height          =   2085
               Left            =   90
               TabIndex        =   40
               Top             =   450
               Width           =   13290
               _ExtentX        =   23442
               _ExtentY        =   3678
               _Version        =   393216
               Rows            =   21
               Cols            =   12
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
               Appearance      =   0
            End
            Begin Threed.SSPanel pnl_Tit_TipOpe 
               Height          =   285
               Left            =   90
               TabIndex        =   41
               Top             =   150
               Width           =   3255
               _Version        =   65536
               _ExtentX        =   5741
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Tipo de Operación"
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
            Begin Threed.SSPanel pnl_Tit_FecOpe 
               Height          =   285
               Left            =   3330
               TabIndex        =   47
               Top             =   150
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Fecha Operación"
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
            Begin Threed.SSPanel pnl_Tit_TipMon 
               Height          =   285
               Left            =   4890
               TabIndex        =   48
               Top             =   150
               Width           =   2895
               _Version        =   65536
               _ExtentX        =   5106
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Moneda"
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
            Begin Threed.SSPanel pnl_Tit_MtoDes 
               Height          =   285
               Left            =   7770
               TabIndex        =   49
               Top             =   150
               Width           =   2115
               _Version        =   65536
               _ExtentX        =   3731
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Pagado"
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
            Begin Threed.SSPanel SSPanel3 
               Height          =   285
               Left            =   9870
               TabIndex        =   50
               Top             =   150
               Width           =   2115
               _Version        =   65536
               _ExtentX        =   3731
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo"
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
End
Attribute VB_Name = "frm_Ges_TecPro_06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_int_NumIte     As Integer
Dim l_arr_CodBan()   As moddat_tpo_Genera
Dim l_arr_CtaBan()   As moddat_tpo_Genera
Dim l_dbl_PorRet     As Double

Private Sub cmb_CodBan_Click()
   cmb_CtaBan.Clear
   cmb_CtaBan.Clear
   pnl_CCIBan.Caption = ""
   If cmb_Moneda.ListIndex = -1 Then cmb_CodBan.ListIndex = -1: Call gs_SetFocus(cmb_Moneda): Exit Sub
   
   If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 4 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 7 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 14 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 21 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 22 Then
      If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 21 Then
         If cmb_TipDoc.ListIndex <> -1 And cmb_CodBan.ListIndex <> -1 Then
            Call fs_Buscar_CtaCte(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex), txt_NumDoc.Text)
         Else
            Call fs_Buscar_CtaCte(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
         End If
      Else
         Call fs_Buscar_CtaCte(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
      End If
   ElseIf cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 5 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 18 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 21 Then
      If cmb_TipDoc.ListIndex <> -1 And cmb_CodBan.ListIndex <> -1 Then
         Call fs_Buscar_CtaCte(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex), txt_NumDoc.Text)
      'Else
      
      End If
   End If
End Sub

Private Sub fs_Buscar_CtaCte(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String)
   If cmb_CodBan.ListIndex = -1 Then
      Exit Sub
   End If
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT CASE WHEN MAEPRV_CODBNC_MN1 = " & "" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "" & " THEN MAEPRV_CTACRR_MN1 END AS CTACRR_MN1, " 'MAEPRV_CTACRR_MN1, "
   g_str_Parame = g_str_Parame & "        CASE WHEN MAEPRV_CODBNC_MN2 = " & "" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "" & " THEN MAEPRV_CTACRR_MN2 END AS CTACRR_MN2, " 'MAEPRV_CTACRR_MN2 , "
   g_str_Parame = g_str_Parame & "        CASE WHEN MAEPRV_CODBNC_MN3 = " & "" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "" & " THEN MAEPRV_CTACRR_MN3 END AS CTACRR_MN3, " 'MAEPRV_CTACRR_MN3, "
   g_str_Parame = g_str_Parame & "        CASE WHEN MAEPRV_CODBNC_MN1 = " & "" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "" & " THEN MAEPRV_NROCCI_MN1 END AS NROCCI_MN1, " 'MAEPRV_NROCCI_MN1, "
   g_str_Parame = g_str_Parame & "        CASE WHEN MAEPRV_CODBNC_MN2 = " & "" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "" & " THEN MAEPRV_NROCCI_MN2 END AS NROCCI_MN2, " 'MAEPRV_NROCCI_MN2 , "
   g_str_Parame = g_str_Parame & "        CASE WHEN MAEPRV_CODBNC_MN3 = " & "" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "" & " THEN MAEPRV_NROCCI_MN3 END AS NROCCI_MN3, " 'MAEPRV_NROCCI_MN3, "
   g_str_Parame = g_str_Parame & "        CASE WHEN MAEPRV_CODBNC_DL1 = " & "" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "" & " THEN MAEPRV_CTACRR_DL1 END AS CTACRR_DL1, " 'MAEPRV_CTACRR_DL1, "
   g_str_Parame = g_str_Parame & "        CASE WHEN MAEPRV_CODBNC_DL2 = " & "" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "" & " THEN MAEPRV_CTACRR_DL2 END AS CTACRR_DL2, " 'MAEPRV_CTACRR_DL2, "
   g_str_Parame = g_str_Parame & "        CASE WHEN MAEPRV_CODBNC_DL3 = " & "" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "" & " THEN MAEPRV_CTACRR_DL3 END AS CTACRR_DL3, " 'MAEPRV_CTACRR_DL3, "
   g_str_Parame = g_str_Parame & "        CASE WHEN MAEPRV_CODBNC_DL1 = " & "" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "" & " THEN MAEPRV_NROCCI_DL1 END AS NROCCI_DL1, " 'MAEPRV_NROCCI_DL1, "
   g_str_Parame = g_str_Parame & "        CASE WHEN MAEPRV_CODBNC_DL2 = " & "" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "" & " THEN MAEPRV_NROCCI_DL2 END AS NROCCI_DL2, " 'MAEPRV_NROCCI_DL2, "
   g_str_Parame = g_str_Parame & "        CASE WHEN MAEPRV_CODBNC_DL3 = " & "" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "" & " THEN MAEPRV_NROCCI_DL3 END AS NROCCI_DL3  " 'MAEPRV_NROCCI_DL3  "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_MAEPRV "
   g_str_Parame = g_str_Parame & "  WHERE MAEPRV_TIPDOC = " & p_TipDoc & ""
   g_str_Parame = g_str_Parame & "    AND MAEPRV_NUMDOC =  '" & p_NumDoc & "'"
   
   If cmb_Moneda.ListIndex <> -1 Then
      If cmb_CodBan.ListIndex <> -1 Then
         If cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = 1 Then
            g_str_Parame = g_str_Parame & "    AND (MAEPRV_CODBNC_MN1 = '" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "'"
            g_str_Parame = g_str_Parame & "     OR  MAEPRV_CODBNC_MN2 = '" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "'"
            g_str_Parame = g_str_Parame & "     OR  MAEPRV_CODBNC_MN3 = '" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "' )"
         Else
            g_str_Parame = g_str_Parame & "    AND (MAEPRV_CODBNC_DL1 = '" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "'"
            g_str_Parame = g_str_Parame & "     OR  MAEPRV_CODBNC_DL2 = '" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "'"
            g_str_Parame = g_str_Parame & "     OR  MAEPRV_CODBNC_DL2 = '" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "' )"
         End If
      End If
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         
         If cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = 1 Then
            'CTA - SOLES
            If Not IsNull(g_rst_Princi!CTACRR_MN1) Then 'MAEPRV_
               cmb_CtaBan.AddItem (g_rst_Princi!CTACRR_MN1) & "- " & "(S/.)" '
            End If
            If Not IsNull(g_rst_Princi!CTACRR_MN2) Then 'MAEPRV_
               cmb_CtaBan.AddItem (g_rst_Princi!CTACRR_MN2) & "- " & "(S/.)"
            End If
            If Not IsNull(g_rst_Princi!CTACRR_MN3) Then 'MAEPRV_
               cmb_CtaBan.AddItem (g_rst_Princi!CTACRR_MN3) & "- " & "(S/.)"
            End If
         Else
            'CTA - DOLARES
            If Not IsNull(g_rst_Princi!CTACRR_DL1) Then 'MAEPRV_
               cmb_CtaBan.AddItem (g_rst_Princi!CTACRR_DL1) & "- " & "(US$)"
            End If
            If Not IsNull(g_rst_Princi!CTACRR_DL2) Then 'MAEPRV_
               cmb_CtaBan.AddItem (g_rst_Princi!CTACRR_DL2) & "- " & "(US$)"
            End If
            If Not IsNull(g_rst_Princi!CTACRR_DL3) Then 'MAEPRV_
               cmb_CtaBan.AddItem (g_rst_Princi!CTACRR_DL3) & "- " & "(US$)"
            End If
          
         End If
         g_rst_Princi.MoveNext
      Loop
   Else
      cmb_CtaBan.Clear
   End If
End Sub

Private Sub cmb_CodBan_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CtaBan)
   End If
End Sub

Private Sub cmb_CtaBan_Click()
   pnl_CCIBan.Caption = ""
   
   If cmb_CtaBan.ListIndex = -1 Then Exit Sub
   If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 4 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 7 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 14 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 21 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 22 Then
      If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 21 Then
         If cmb_TipDoc.ListIndex <> -1 Then
            Call fs_Buscar_CCI(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex), txt_NumDoc.Text)
         Else
            Call fs_Buscar_CCI(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
         End If
      Else
         Call fs_Buscar_CCI(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
      End If
   ElseIf cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 5 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 18 Then
      If cmb_TipDoc.ListIndex <> -1 Then
         Call fs_Buscar_CCI(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex), txt_NumDoc.Text)
      End If
   End If
   Call gs_SetFocus(txt_NumMov)
End Sub

Private Sub fs_Buscar_CCI(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT CASE WHEN MAEPRV_CODBNC_MN1 = " & "" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "" & " AND MAEPRV_CTACRR_MN1 = " & "'" & Trim(Mid(cmb_CtaBan.Text, 1, InStr(cmb_CtaBan.Text, "-") - 1)) & "'" & " THEN MAEPRV_NROCCI_MN1 END AS NROCCI_MN1, "
   g_str_Parame = g_str_Parame & "        CASE WHEN MAEPRV_CODBNC_MN2 = " & "" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "" & " AND MAEPRV_CTACRR_MN2 = " & "'" & Trim(Mid(cmb_CtaBan.Text, 1, InStr(cmb_CtaBan.Text, "-") - 1)) & "'" & " THEN MAEPRV_NROCCI_MN2 END AS NROCCI_MN2, "
   g_str_Parame = g_str_Parame & "        CASE WHEN MAEPRV_CODBNC_MN3 = " & "" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "" & " AND MAEPRV_CTACRR_MN3 = " & "'" & Trim(Mid(cmb_CtaBan.Text, 1, InStr(cmb_CtaBan.Text, "-") - 1)) & "'" & " THEN MAEPRV_NROCCI_MN3 END AS NROCCI_MN3, "
   g_str_Parame = g_str_Parame & "        CASE WHEN MAEPRV_CODBNC_DL1 = " & "" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "" & " AND MAEPRV_CTACRR_DL1 = " & "'" & Trim(Mid(cmb_CtaBan.Text, 1, InStr(cmb_CtaBan.Text, "-") - 1)) & "'" & " THEN MAEPRV_NROCCI_DL1 END AS NROCCI_DL1, "
   g_str_Parame = g_str_Parame & "        CASE WHEN MAEPRV_CODBNC_DL2 = " & "" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "" & " AND MAEPRV_CTACRR_DL2 = " & "'" & Trim(Mid(cmb_CtaBan.Text, 1, InStr(cmb_CtaBan.Text, "-") - 1)) & "'" & " THEN MAEPRV_NROCCI_DL2 END AS NROCCI_DL2, "
   g_str_Parame = g_str_Parame & "        CASE WHEN MAEPRV_CODBNC_DL3 = " & "" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "" & " AND MAEPRV_CTACRR_DL3 = " & "'" & Trim(Mid(cmb_CtaBan.Text, 1, InStr(cmb_CtaBan.Text, "-") - 1)) & "'" & " THEN MAEPRV_NROCCI_DL3 END AS NROCCI_DL3 "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_MAEPRV "
   g_str_Parame = g_str_Parame & "  WHERE MAEPRV_TIPDOC = " & p_TipDoc & ""
   g_str_Parame = g_str_Parame & "    AND MAEPRV_NUMDOC =  '" & p_NumDoc & "'"
   
   If cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = 1 Then
      g_str_Parame = g_str_Parame & "    AND (MAEPRV_CODBNC_MN1 = '" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "'"
      g_str_Parame = g_str_Parame & "     OR  MAEPRV_CODBNC_MN2 = '" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "'"
      g_str_Parame = g_str_Parame & "     OR  MAEPRV_CODBNC_MN3 = '" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "') "
      g_str_Parame = g_str_Parame & "    AND (MAEPRV_CTACRR_MN1 = '" & Trim(Mid(cmb_CtaBan.Text, 1, InStr(cmb_CtaBan.Text, "-") - 1)) & "'"
      g_str_Parame = g_str_Parame & "     OR  MAEPRV_CTACRR_MN2 = '" & Trim(Mid(cmb_CtaBan.Text, 1, InStr(cmb_CtaBan.Text, "-") - 1)) & "'"
      g_str_Parame = g_str_Parame & "     OR  MAEPRV_CTACRR_MN3 = '" & Trim(Mid(cmb_CtaBan.Text, 1, InStr(cmb_CtaBan.Text, "-") - 1)) & "' )"
   Else
      g_str_Parame = g_str_Parame & "    AND (MAEPRV_CODBNC_DL1 = '" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "'"
      g_str_Parame = g_str_Parame & "     OR  MAEPRV_CODBNC_DL2 = '" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "'"
      g_str_Parame = g_str_Parame & "     OR  MAEPRV_CODBNC_DL3 = '" & cmb_CodBan.ItemData(cmb_CodBan.ListIndex) & "') "
      g_str_Parame = g_str_Parame & "    AND (MAEPRV_CTACRR_DL1 = '" & Trim(Mid(cmb_CtaBan.Text, 1, InStr(cmb_CtaBan.Text, "-") - 1)) & "'"
      g_str_Parame = g_str_Parame & "     OR  MAEPRV_CTACRR_DL2 = '" & Trim(Mid(cmb_CtaBan.Text, 1, InStr(cmb_CtaBan.Text, "-") - 1)) & "'"
      g_str_Parame = g_str_Parame & "     OR  MAEPRV_CTACRR_DL3 = '" & Trim(Mid(cmb_CtaBan.Text, 1, InStr(cmb_CtaBan.Text, "-") - 1)) & "' )"
   End If

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst

      Do While Not g_rst_Princi.EOF
         If cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = 1 Then
            'CCI - SOLES
            If Not IsNull(g_rst_Princi!NROCCI_MN1) Then 'MAEPRV_
               pnl_CCIBan.Caption = (g_rst_Princi!NROCCI_MN1)
            End If
            If Not IsNull(g_rst_Princi!NROCCI_MN2) Then
               pnl_CCIBan.Caption = (g_rst_Princi!NROCCI_MN2)
            End If
            If Not IsNull(g_rst_Princi!NROCCI_MN3) Then
               pnl_CCIBan.Caption = (g_rst_Princi!NROCCI_MN3)
            End If
         Else
            'CCI - DÓLARES
            If Not IsNull(g_rst_Princi!NROCCI_DL1) Then
               pnl_CCIBan.Caption = (g_rst_Princi!NROCCI_DL1)
            End If
            If Not IsNull(g_rst_Princi!NROCCI_DL2) Then
               pnl_CCIBan.Caption = (g_rst_Princi!NROCCI_DL2)
            End If
            If Not IsNull(g_rst_Princi!NROCCI_DL3) Then
               pnl_CCIBan.Caption = (g_rst_Princi!NROCCI_DL3)
            End If
         End If
         g_rst_Princi.MoveNext
      Loop
   Else
      pnl_CCIBan.Caption = ""
   End If
End Sub

Private Sub cmb_CtaBan_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub cmb_ForPag_Click()
   If cmb_ForPag.ListIndex <> -1 Then
      If cmb_ForPag.ItemData(cmb_ForPag.ListIndex) = 1 Then
         cmb_CodBan.Enabled = False
         cmb_CodBan.ListIndex = -1
         cmb_CtaBan.Enabled = False
         cmb_CtaBan.ListIndex = -1
         pnl_CCIBan.Enabled = False
         pnl_CCIBan.Caption = Empty
         cmb_TipDoc.Enabled = True
         txt_NumDoc.Enabled = True
         pnl_RScPrv.Enabled = True
         Call fs_Limpia_DatBan
         If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 14 Then
            Call gs_SetFocus(cmb_TipDoc)
         Else
            Call gs_SetFocus(ipp_ImpDes)
         End If
      Else
         If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 4 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 7 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 14 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 21 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 22 Then
            cmb_CodBan.Enabled = True
            cmb_CtaBan.Enabled = True
            pnl_CCIBan.Enabled = True
            
            If moddat_g_int_FlgAct = 2 Then
               cmb_TipDoc.Enabled = False
               txt_NumDoc.Enabled = False
               cmb_CodBan.Enabled = False
               cmb_CtaBan.Enabled = False
               pnl_CCIBan.Enabled = False
            Else
               If moddat_g_str_CodMod = "008" Then 'moddat_g_str_CodPrd = "026" AND moddat_g_str_CodSub = "002"
                  cmb_TipDoc.Enabled = True
                  txt_NumDoc.Enabled = True
               Else
                  cmb_TipDoc.Enabled = False
                  txt_NumDoc.Enabled = False
               End If
            End If
            
            If moddat_g_int_FlgAct <> 2 Then 'cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 21 and
               cmb_TipDoc.ListIndex = -1
               txt_NumDoc.Text = Empty
               pnl_RScPrv.Caption = Empty
               Call fs_Buscar_Banco(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
            End If
            
         Else
            cmb_CodBan.Enabled = True
            cmb_CtaBan.Enabled = True
            pnl_CCIBan.Enabled = True
            cmb_TipDoc.Enabled = True
            txt_NumDoc.Enabled = True
         End If
         Call gs_SetFocus(ipp_ImpDes)
      End If
   End If
End Sub

Private Sub cmb_ForPag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ImpDes)
   End If
End Sub

Private Sub cmb_Moneda_Click()
   If cmb_TipOper.ListIndex <> -1 Then
      If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 4 And cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 5 And cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 7 And cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 14 And cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 18 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 21 Then
         Call gs_SetFocus(txt_Refer)
      Else
         If cmb_ForPag.ListIndex <> -1 Then
            If cmb_ForPag.ItemData(cmb_ForPag.ListIndex) = 1 Then
               Call gs_SetFocus(cmb_TipDoc)
            Else
               If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 5 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 18 Then
                  Call gs_SetFocus(cmb_TipDoc)
               Else
                  cmb_CodBan.ListIndex = -1
                  Call gs_SetFocus(cmb_CodBan)
               End If
            End If
         Else
            Call gs_SetFocus(cmb_CodBan)
         End If
      End If
   End If
End Sub


Private Sub cmb_Moneda_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_TipOper.ListIndex <> -1 Then
         If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 4 And cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 5 And cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 7 And cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 14 And cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 18 And cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 21 And cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 22 Then
            Call gs_SetFocus(txt_Refer)
         Else
            If cmb_ForPag.ListIndex <> -1 Then
               If cmb_ForPag.ItemData(cmb_ForPag.ListIndex) = 1 Then
                  Call gs_SetFocus(cmb_TipDoc)
               Else
                  If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 4 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 21 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 22 Then
                     Call gs_SetFocus(cmb_CodBan)
                  ElseIf cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 5 Then
                     Call gs_SetFocus(cmb_TipDoc)
                  End If
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub cmb_TipDoc_Click()
     Call gs_SetFocus(txt_NumDoc)
     If moddat_g_int_FlgAct <> 2 Then
      cmb_CodBan.Clear
      cmb_CtaBan.Clear
     End If
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumDoc)
   End If
End Sub

Private Sub cmb_TipOper_Click()
Dim r_dbl_ImpTot     As Double
Dim r_dbl_ImpSal     As Double
Dim r_dbl_ImpPag     As Double
   
   'Call fs_Limpia
   If cmb_TipOper.ListIndex <> -1 Then
      If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 4 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 7 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 21 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 22 Then
         If cmb_ForPag.ListIndex <> -1 Then
            If cmb_ForPag.ItemData(cmb_ForPag.ListIndex) = 1 Then
               cmb_TipDoc.Enabled = True
               txt_NumDoc.Enabled = True
               pnl_RScPrv.Enabled = True
               cmb_CodBan.Enabled = False
               cmb_CtaBan.Enabled = False
               pnl_CCIBan.Enabled = False
               cmb_ForPag.Enabled = True
               cmb_CodBan.ListIndex = -1
               cmb_CtaBan.ListIndex = -1
               pnl_CCIBan.Caption = Empty
            Else
               cmb_TipDoc.Enabled = False
               txt_NumDoc.Enabled = False
               pnl_RScPrv.Enabled = False
               cmb_CodBan.Enabled = True
               cmb_CtaBan.Enabled = True
               pnl_CCIBan.Enabled = True
               cmb_ForPag.Enabled = True
               
               cmb_TipDoc.ListIndex = -1
               txt_NumDoc.Text = Empty
               pnl_RScPrv.Caption = Empty
               cmb_CodBan.ListIndex = -1
               cmb_CtaBan.ListIndex = -1
               pnl_CCIBan.Caption = Empty
            End If
         Else
            cmb_ForPag.Enabled = True
            cmb_TipDoc.Enabled = False
            txt_NumDoc.Enabled = False
            cmb_CodBan.Enabled = False
            cmb_CtaBan.Enabled = False
         End If
      ElseIf cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 5 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 14 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 18 Then
         If cmb_ForPag.ListIndex <> -1 Then
            If cmb_ForPag.ItemData(cmb_ForPag.ListIndex) = 1 Then
               cmb_TipDoc.Enabled = True 'False
               txt_NumDoc.Enabled = True 'False
               pnl_RScPrv.Enabled = True 'False
               cmb_CodBan.Enabled = False 'True
               cmb_CtaBan.Enabled = False 'True
               pnl_CCIBan.Enabled = False 'True
               cmb_ForPag.Enabled = True
            Else
               cmb_TipDoc.Enabled = True
               txt_NumDoc.Enabled = True
               pnl_RScPrv.Enabled = True
               cmb_CodBan.Enabled = True
               cmb_CtaBan.Enabled = True
               pnl_CCIBan.Enabled = True
               cmb_ForPag.Enabled = True
               cmb_TipDoc.ListIndex = -1
               txt_NumDoc.Text = Empty
               pnl_RScPrv.Caption = Empty
               cmb_CodBan.ListIndex = -1
               cmb_CtaBan.ListIndex = -1
               pnl_CCIBan.Caption = Empty
            End If
         Else
            cmb_TipDoc.Enabled = False
            txt_NumDoc.Enabled = False
            pnl_RScPrv.Enabled = True
            cmb_CodBan.Enabled = False
            cmb_CtaBan.Enabled = False
            pnl_CCIBan.Enabled = True
            cmb_ForPag.Enabled = True
            cmb_CodBan.Clear
         End If
      Else
         cmb_TipDoc.Enabled = False
         txt_NumDoc.Enabled = False
         pnl_RScPrv.Enabled = False
         cmb_CodBan.Enabled = False
         cmb_CtaBan.Enabled = False
         pnl_CCIBan.Enabled = False
         cmb_ForPag.Enabled = False
         Call fs_Limpia_DatBan
         cmb_ForPag.ListIndex = -1
      End If

      cmb_CodBan.ListIndex = -1
      cmb_CtaBan.Clear
      pnl_CCIBan.Caption = ""
      pnl_CodOpe.Caption = Format(cmb_TipOper.ItemData(cmb_TipOper.ListIndex), "000") & " "
      
      If moddat_g_int_FlgAct <> 2 Then Call fs_Buscar
      
      pnl_ImpTot.Caption = Format(0, "###,###,###,##0.00") & "  "
      pnl_ImpSal.Caption = Format(0, "###,###,###,##0.00") & "  "
      ipp_ImpDes.Value = Format(0, "###,###,###,##0.00")
        
      If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 1 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 2 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 13 Then
         Call fs_Buscar_Saldo_Comision(r_dbl_ImpTot, r_dbl_ImpSal, r_dbl_ImpPag)
         pnl_ImpTot.Caption = Format(CDbl(r_dbl_ImpTot), "###,###,###,##0.00") & "  "
         pnl_ImpSal.Caption = Format(CDbl(r_dbl_ImpSal), "###,###,###,##0.00") & "  "
'         If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 15 Then
'            Call fs_ActOper(False)
'            cmb_Moneda.ListIndex = 0
'            If r_dbl_ImpSal < 0 Then pnl_ImpSal.Caption = Format(CDbl(0), "###,###,###,##0.00") & "  "
'         Else
            Call fs_ActOper(True)
'         End If
      ElseIf cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 3 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 11 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 12 Or _
             cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 19 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 22 Then
         Call fs_Buscar_Saldo_Fondos(r_dbl_ImpTot, r_dbl_ImpSal)
         pnl_ImpTot.Caption = Format(CDbl(r_dbl_ImpTot), "###,###,###,##0.00") & "  "
         pnl_ImpSal.Caption = Format(CDbl(r_dbl_ImpSal), "###,###,###,##0.00") & "  "
         Call fs_ActOper(True)
      ElseIf cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 4 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 5 Then
         Call fs_Buscar_Saldo_Desembolso
         Call fs_ActOper(True)
         Call fs_Buscar_Banco(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
      ElseIf cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 6 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 20 Then
         Call fs_Buscar_Saldo_Garantia(cmb_TipOper.ItemData(cmb_TipOper.ListIndex))
         Call fs_ActOper(True)
      ElseIf cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 7 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 21 Then
         Call fs_Buscar_Saldo_Garantia_Devolucion(cmb_TipOper.ItemData(cmb_TipOper.ListIndex))
         Call fs_ActOper(True)
         Call fs_Buscar_Banco(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
      ElseIf cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 8 Then
         Call fs_ActOper(False)
      ElseIf cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 9 Then
         Call fs_ActOper(False)
      ElseIf cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 14 Then
         Call fs_Buscar_Saldo_Retencion
         Call fs_ActOper(False)
         Call fs_Buscar_Banco(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
         ipp_ImpDes.Value = Format(CDbl(pnl_ImpSal.Caption), "###,###,###,##0.00")
      ElseIf cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 15 Then
         Call fs_Buscar_Extorno_Comision(r_dbl_ImpTot, r_dbl_ImpSal, r_dbl_ImpPag)
         Call fs_ActOper(False)
         cmb_Moneda.ListIndex = 0
         pnl_ImpTot.Caption = Format(CDbl(r_dbl_ImpTot), "###,###,###,##0.00") & "  "
         pnl_ImpSal.Caption = Format(CDbl(r_dbl_ImpSal), "###,###,###,##0.00") & "  "
      ElseIf cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 18 Then
         Call fs_Buscar_Saldo_Desembolso
         Call fs_ActOper(True)
         Call fs_Buscar_Banco(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
      Else
         pnl_ImpTot.Caption = Format(CDbl(0), "###,###,###,##0.00") & "  "
         pnl_ImpSal.Caption = Format(CDbl(0), "###,###,###,##0.00") & "  "
      End If
   End If
   
   Call gs_SetFocus(ipp_FecOpe)
End Sub

Private Sub cmb_TipOper_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecOpe)
   End If
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgAct = 1
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   
   Call fs_Activa(True)
   If Me.grd_Listad.Row <= 0 Then
      cmd_Editar.Enabled = False
      cmd_Borrar.Enabled = False
      cmd_ExpExc.Enabled = False
      cmd_ExpLiq.Enabled = False
   Else
      'cmd_Editar.Enabled = True
      'cmd_Borrar.Enabled = True
   End If
   Call gs_SetFocus(cmb_TipOper)
End Sub

Private Sub cmd_Borrar_Click()
Dim r_str_EstMod     As String
Dim r_str_NroAsi     As String
Dim r_str_NroLib     As String

   If CInt(grd_Listad.TextMatrix(grd_Listad.Row, 9)) = 2 Or CInt(grd_Listad.TextMatrix(grd_Listad.Row, 9)) = 10 Or CInt(grd_Listad.TextMatrix(grd_Listad.Row, 9)) = 3 Or CInt(grd_Listad.TextMatrix(grd_Listad.Row, 9)) = 6 Or CInt(grd_Listad.TextMatrix(grd_Listad.Row, 9)) = 19 Or CInt(grd_Listad.TextMatrix(grd_Listad.Row, 9)) = 20 Then
      If fs_Validar_MovReg = True Then
         If CInt(grd_Listad.TextMatrix(grd_Listad.Row, 9)) = 2 Or CInt(grd_Listad.TextMatrix(grd_Listad.Row, 9)) = 10 Then
            MsgBox "No se puede eliminar, el registro es parte de los Fondos Recibidos.", vbExclamation, modgen_g_str_NomPlt
         ElseIf CInt(grd_Listad.TextMatrix(grd_Listad.Row, 9)) = 3 Or CInt(grd_Listad.TextMatrix(grd_Listad.Row, 9)) = 19 Then
            MsgBox "No se puede eliminar, Verifique que el registro no tenga Desembolsos o Pagos a Cuenta.", vbExclamation, modgen_g_str_NomPlt
         ElseIf CInt(grd_Listad.TextMatrix(grd_Listad.Row, 9)) = 6 Or CInt(grd_Listad.TextMatrix(grd_Listad.Row, 9)) = 20 Then
            MsgBox "No se puede eliminar, Verifique que el registro no tenga Devolución de Garantía.", vbExclamation, modgen_g_str_NomPlt
         End If
         Call gs_SetFocus(grd_Listad)
         Exit Sub
      End If
   End If
   
   'Validar que ya se encuentra autorizado
   Call fs_Validar_ComAut(grd_Listad.TextMatrix(grd_Listad.Row, 9), grd_Listad.TextMatrix(grd_Listad.Row, 10), Trim(moddat_g_str_DesIte), r_str_EstMod, r_str_NroAsi) 'pnl_NumRef.Caption
   If r_str_EstMod = "PENDIENTE" Then
      MsgBox "No se puede eliminar, el registro está por ser Autorizado. " & vbCrLf & "Favor verificar el módulo de Autorización", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   ElseIf r_str_EstMod = "APROBADO" Then
      MsgBox "No se puede eliminar, el registro ya se encuentra Autorizado. " & vbCrLf & "Favor verificar el módulo de Compensación", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   ElseIf r_str_EstMod = "APLICADO" Then
      MsgBox "No se puede eliminar, el registro está por Pagarse. " & vbCrLf & "Favor verificar el módulo de Compensación", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   ElseIf r_str_EstMod = "PAGADO" Then
      MsgBox "No se puede eliminar, el registro ya se encuentra Pagado. " & vbCrLf & "Favor verificar el módulo de Compensación", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   ElseIf InStr(r_str_NroAsi, "LM/") > 0 And r_str_EstMod = "" Then
      MsgBox "No se puede eliminar, el registro está asociado a un asiento contable. " & vbCrLf & "Favor verificar en la Plataforma de Contabilidad", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   ElseIf InStr(r_str_EstMod, "RECHAZADO") > 0 Then
      r_str_NroLib = Mid(Mid(r_str_NroAsi, 1, InStrRev(r_str_NroAsi, "/") - 1), InStrRev(Mid(r_str_NroAsi, 1, InStrRev(r_str_NroAsi, "/") - 1), "/") + 1)
      r_str_NroAsi = Mid(r_str_NroAsi, InStrRev(r_str_NroAsi, "/") + 1)
      
      If MsgBox("Recuerde eliminar el Asiento Contable Nro. " & r_str_NroAsi & " del Libro " & r_str_NroLib & ", ¿Está seguro de eliminar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
      
      moddat_g_int_FlgGOK = False
      moddat_g_int_CntErr = 0
      Do While moddat_g_int_FlgGOK = False
         Screen.MousePointer = 11
      
         'Grabando Información de Carta Fianza
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " USP_TPR_MAERDE_ELIMINA ("
         g_str_Parame = g_str_Parame & CStr(grd_Listad.TextMatrix(grd_Listad.Row, 10)) & ", "
         g_str_Parame = g_str_Parame & "'" & CStr(Replace(moddat_g_str_DesIte, "-", "")) & "') " 'pnl_NumRef.Caption
                     
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
         End If
         
         If moddat_g_int_CntErr = 6 Then
            If MsgBox("No se pudo completar la eliminación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
               Exit Sub
            Else
               moddat_g_int_CntErr = 0
            End If
         End If
         
         Screen.MousePointer = 0
      Loop
   Else
      MsgBox "No se puede eliminar, pertenece a una carga inicial. ", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
     
   'Actualiza la Grilla
   Call fs_Buscar
   Call fs_Activa(False)
   frm_Ges_TecPro_03.fs_Buscar_Creditos_Indirectos
   frm_Ges_TecPro_03.fs_Buscar_Creditos_Directos
End Sub

Private Function fs_Validar_MovReg() As Boolean
   fs_Validar_MovReg = False
   
   If CInt(grd_Listad.TextMatrix(grd_Listad.Row, 9)) = 2 Or CInt(grd_Listad.TextMatrix(grd_Listad.Row, 9)) = 10 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "  SELECT COUNT(*) AS CONTADOR "
      g_str_Parame = g_str_Parame & "    FROM TPR_MAERDE "
      g_str_Parame = g_str_Parame & "   WHERE MAERDE_NUMREF = '" & CStr(Replace(moddat_g_str_DesIte, "-", "")) & "' " 'pnl_NumRef.Caption
      g_str_Parame = g_str_Parame & "     AND (MAERDE_CODIGO IN (3,19)) "
   
   ElseIf CInt(grd_Listad.TextMatrix(grd_Listad.Row, 9)) = 3 Or CInt(grd_Listad.TextMatrix(grd_Listad.Row, 9)) = 19 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "  SELECT COUNT(*) AS CONTADOR "
      g_str_Parame = g_str_Parame & "    FROM TPR_MAERDE "
      g_str_Parame = g_str_Parame & "   WHERE MAERDE_NUMREF = '" & CStr(Replace(moddat_g_str_DesIte, "-", "")) & "' " 'pnl_NumRef.Caption
      g_str_Parame = g_str_Parame & "     AND (MAERDE_CODIGO = 4 OR MAERDE_CODIGO = 5) "
      
   ElseIf CInt(grd_Listad.TextMatrix(grd_Listad.Row, 9)) = 6 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "  SELECT COUNT(*) AS CONTADOR "
      g_str_Parame = g_str_Parame & "    FROM TPR_MAERDE "
      g_str_Parame = g_str_Parame & "   WHERE MAERDE_NUMREF = '" & CStr(Replace(moddat_g_str_DesIte, "-", "")) & "' " 'pnl_NumRef.Caption
      g_str_Parame = g_str_Parame & "     AND MAERDE_CODIGO = 7 "
   
   ElseIf CInt(grd_Listad.TextMatrix(grd_Listad.Row, 9)) = 20 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "  SELECT COUNT(*) AS CONTADOR "
      g_str_Parame = g_str_Parame & "    FROM TPR_MAERDE "
      g_str_Parame = g_str_Parame & "   WHERE MAERDE_NUMREF = '" & CStr(Replace(moddat_g_str_DesIte, "-", "")) & "' "
      g_str_Parame = g_str_Parame & "     AND MAERDE_CODIGO = 21 "
   
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
     Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Function
   End If
     
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      If (g_rst_GenAux!CONTADOR) > 0 Then
         fs_Validar_MovReg = True
      End If
   End If
End Function

Private Sub cmd_Cancel_Click()
   Call fs_Limpia
   Call fs_Buscar
   Call fs_Activa(False)
End Sub

Private Sub cmd_Editar_Click()
Dim r_str_Parame     As String
   
   moddat_g_int_FlgAct = 2
   Call fs_Limpia
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "  SELECT A.MAERDE_CODIGO, TRIM(B.PARDES_DESCRI) OPERACION , A.MAERDE_FECASG, TRIM(C.PARDES_DESCRI) MONEDA, A.MAERDE_IMPORT, MAERDE_OPEREF, "
   r_str_Parame = r_str_Parame & "         D.MAEPRV_RAZSOC, A.MAERDE_CODBAN, MAERDE_CTACTE, A.MAERDE_NUMCCI  , A.MAERDE_TDOPRV, A.MAERDE_NDOPRV, A.MAERDE_NUMMOV, A.MAERDE_FORPAG "
   r_str_Parame = r_str_Parame & "    FROM TPR_MAERDE A "
   r_str_Parame = r_str_Parame & "           INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = 528 AND B.PARDES_CODITE = A.MAERDE_CODIGO "
   r_str_Parame = r_str_Parame & "           INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 204 AND C.PARDES_CODITE = A.MAERDE_TIPMON "
   r_str_Parame = r_str_Parame & "            LEFT JOIN CNTBL_MAEPRV D ON D.MAEPRV_TIPDOC = A.MAERDE_TDOPRV AND D.MAEPRV_NUMDOC = A.MAERDE_NDOPRV "
   r_str_Parame = r_str_Parame & "   WHERE A.MAERDE_NUMREF = '" & CStr(moddat_g_str_DesIte) & "'" 'Replace(pnl_NumRef.Caption, "-", "")
   r_str_Parame = r_str_Parame & "     AND A.MAERDE_NUMITE = " & grd_Listad.TextMatrix(grd_Listad.Row, 10) & ""
      
   If Not gf_EjecutaSQL(r_str_Parame, g_rst_GenAux, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      If moddat_g_str_DesObs = "" Then
         Call fs_Activa(True)
      Else
         SSPanel11.Enabled = False
         SSPanel5.Enabled = False
      End If
      Call gs_BuscarCombo(cmb_TipOper, g_rst_GenAux!OPERACION)

      If cmb_TipOper.ListIndex = -1 Then
         Call fs_Limpia
         Call fs_Activa(False)
         Exit Sub
      End If
      
      ipp_FecOpe.Text = Format(gf_FormatoFecha(CStr(g_rst_GenAux!MAERDE_FECASG)), "dd/mm/yyyy")
      cmb_Moneda.Text = g_rst_GenAux!Moneda
      ipp_ImpDes.Value = Format(g_rst_GenAux!MAERDE_IMPORT, "###,###,###,##0.00")
      pnl_ImpSal.Caption = Format(CDbl(ipp_ImpDes.Value) + CDbl(grd_Listad.TextMatrix(grd_Listad.Row, 4)), "###,###,###,##0.00") & "  "
      txt_Refer.Text = IIf(IsNull(g_rst_GenAux!MAERDE_OPEREF), "", Trim(g_rst_GenAux!MAERDE_OPEREF))
      
      If Not IsNull(g_rst_GenAux!MAERDE_TDOPRV) And g_rst_GenAux!MAERDE_TDOPRV > 0 Then
         cmb_TipDoc.Text = moddat_gf_Consulta_ParDes("118", g_rst_GenAux!MAERDE_TDOPRV)
         txt_NumDoc.Text = g_rst_GenAux!MAERDE_NDOPRV
         Call fs_Buscar_Banco(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex), txt_NumDoc.Text)
         pnl_RScPrv.Caption = g_rst_GenAux!MAEPRV_RAZSOC
      End If
      
      If (g_rst_GenAux!MAERDE_CODBAN) > 0 Then
         Call gs_BuscarCombo(cmb_CodBan, moddat_gf_Consulta_ParDes("122", g_rst_GenAux!MAERDE_CODBAN))
    
         If cmb_CodBan <> "" Then
            If cmb_Moneda.ItemData(cmb_Moneda.ListIndex) = 1 Then
               Call gs_BuscarCombo_Text(cmb_CtaBan, g_rst_GenAux!MAERDE_CTACTE, 0)
               If cmb_CtaBan <> "" Then
                  cmb_CtaBan.Text = Trim(g_rst_GenAux!MAERDE_CTACTE) & "- " & "(S/.)"
               End If
            Else
               cmb_CtaBan.Text = Trim(g_rst_GenAux!MAERDE_CTACTE) & "- " & "(US$)"
            End If
         End If
         pnl_CCIBan.Caption = IIf(IsNull(g_rst_GenAux!MAERDE_NUMCCI), "", Trim(g_rst_GenAux!MAERDE_NUMCCI))
      End If

      txt_NumMov.Text = IIf(IsNull(g_rst_GenAux!MAERDE_NUMMOV), "", g_rst_GenAux!MAERDE_NUMMOV)
      l_int_NumIte = grd_Listad.TextMatrix(grd_Listad.Row, 10)
      If Not IsNull(g_rst_GenAux!MAERDE_FORPAG) Then
         If g_rst_GenAux!MAERDE_FORPAG > 0 Then
            cmb_ForPag.Text = moddat_gf_Consulta_ParDes("531", g_rst_GenAux!MAERDE_FORPAG)
         End If
      End If
   End If

   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
End Sub

Private Function fs_Validar_ComAut(ByVal p_Codigo As Integer, ByVal p_NumIte As String, ByVal p_NumRef As String, ByRef p_Descri As String, ByRef p_NroAsi As String)
   p_Descri = ""
   p_NroAsi = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT TRIM (PARDES_DESCRI) AS CONDICION "
   g_str_Parame = g_str_Parame & "    FROM CNTBL_COMAUT "
   g_str_Parame = g_str_Parame & "         INNER JOIN MNT_PARDES ON PARDES_CODGRP = 137 AND PARDES_CODITE = COMAUT_CODEST "
   g_str_Parame = g_str_Parame & "   WHERE COMAUT_CODOPE = ( SELECT MAERDE_CODOPE "
   g_str_Parame = g_str_Parame & "                             FROM TPR_MAERDE "
   g_str_Parame = g_str_Parame & "                            WHERE MAERDE_CODIGO = " & p_Codigo & ""
   g_str_Parame = g_str_Parame & "                              AND MAERDE_NUMITE = '" & p_NumIte & "'"
   g_str_Parame = g_str_Parame & "                              AND MAERDE_NUMREF = '" & Replace(p_NumRef, "-", "") & "')"
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      p_Descri = g_rst_Genera!CONDICION
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
     
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "   SELECT MAERDE_NROCNT "
   g_str_Parame = g_str_Parame & "     FROM TPR_MAERDE "
   g_str_Parame = g_str_Parame & "    WHERE MAERDE_CODIGO = " & p_Codigo & ""
   g_str_Parame = g_str_Parame & "      AND MAERDE_NUMITE = '" & p_NumIte & "'"
   g_str_Parame = g_str_Parame & "      AND MAERDE_NUMREF = '" & Replace(p_NumRef, "-", "") & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      If Not IsNull(g_rst_Genera!MAERDE_NROCNT) Then
         p_NroAsi = g_rst_Genera!MAERDE_NROCNT
      End If
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Private Sub cmd_ExpExc_Click()
    'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_str_ParAux        As String
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
Dim r_int_NroFil        As Integer
Dim r_int_NoFlLi        As Integer
   
    r_int_NroFil = 9
    Set r_obj_Excel = New Excel.Application
    r_obj_Excel.SheetsInNewWorkbook = 1
    r_obj_Excel.Workbooks.Add

    With r_obj_Excel.ActiveSheet
        .Cells(2, 2) = "REPORTE DE GESTIÓN DE OPERACIONES"
        .Range(.Cells(2, 2), .Cells(2, 6)).Merge
        .Range(.Cells(2, 2), .Cells(2, 6)).Font.Bold = True
        .Range(.Cells(2, 2), .Cells(2, 6)).HorizontalAlignment = xlHAlignCenter
        .Range(.Cells(2, 2), .Cells(2, 6)).Font.Size = 14

        .Cells(4, 2) = "TIPO DE DOCUMENTO"
        .Cells(4, 3) = Trim(pnl_TipDoc.Caption)
        .Cells(5, 2) = "NRO. DOCUMENTO"
        .Cells(5, 3) = "'" & Trim(pnl_NroDoc.Caption)
        .Cells(6, 2) = "RAZÓN SOCIAL"
        .Cells(6, 3) = Trim(pnl_RazSoc.Caption)
        .Cells(7, 2) = "NÚMERO"
        .Cells(7, 3) = "'" & pnl_NumRef.Caption
        .Range(.Cells(3, 2), .Cells(7, 2)).Font.Bold = True
        
        .Cells(r_int_NroFil, 2) = "TIPO DE OPERACION"
        .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 2)).Merge
        .Cells(r_int_NroFil, 3) = "FECHA DE OPERACION"
        .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil + 1, 3)).Merge
        .Cells(r_int_NroFil, 4) = "MONEDA"
        .Range(.Cells(r_int_NroFil, 4), .Cells(r_int_NroFil + 1, 4)).Merge
        .Cells(r_int_NroFil, 5) = "PAGADO"
        .Range(.Cells(r_int_NroFil, 5), .Cells(r_int_NroFil + 1, 5)).Merge
        .Cells(r_int_NroFil, 6) = "SALDO"
        .Range(.Cells(r_int_NroFil, 6), .Cells(r_int_NroFil + 1, 6)).Merge
        
        .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 6)).Interior.Color = RGB(146, 208, 80)
        .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 6)).Font.Bold = True
        .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 6)).HorizontalAlignment = xlHAlignCenter
        
        .Columns("A").ColumnWidth = 1
        .Columns("B").ColumnWidth = 28 '25.5
        .Columns("C").ColumnWidth = 13
        .Columns("D").ColumnWidth = 22
        .Columns("D").HorizontalAlignment = xlHAlignCenter
        .Columns("E").ColumnWidth = 13.5
        .Columns("E").NumberFormat = "###,###,###,##0.00"
        .Columns("E").HorizontalAlignment = xlHAlignRight
        .Columns("F").ColumnWidth = 13.5
        .Columns("F").NumberFormat = "###,###,###,##0.00"
        .Columns("F").HorizontalAlignment = xlHAlignRight
                
        With .Range(.Cells(8, 2), .Cells(9, 6))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
        
        .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
        .Range(.Cells(3, 1), .Cells(99, 99)).Font.Size = 11
        r_int_NroFil = r_int_NroFil + 2
         
        For r_int_NoFlLi = 0 To grd_Listad.Rows - 1
            .Cells(r_int_NroFil, 2) = "'" & grd_Listad.TextMatrix(r_int_NoFlLi, 0)
            .Cells(r_int_NroFil, 3) = "'" & grd_Listad.TextMatrix(r_int_NoFlLi, 1)
            .Cells(r_int_NroFil, 4) = "'" & grd_Listad.TextMatrix(r_int_NoFlLi, 2)
            .Cells(r_int_NroFil, 5) = grd_Listad.TextMatrix(r_int_NoFlLi, 3)
            .Cells(r_int_NroFil, 6) = grd_Listad.TextMatrix(r_int_NoFlLi, 4)
            
            r_int_NroFil = r_int_NroFil + 1
        Next r_int_NoFlLi
        
        With .Range(.Cells(10, 3), .Cells(r_int_NroFil, 3))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
         
        With .Range(.Cells(9, 2), .Cells(r_int_NroFil - 1, 6))
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        End With
   End With
   
   r_obj_Excel.Visible = True
End Sub

Private Sub cmd_ExpLiq_Click()
   If grd_Listad.TextMatrix(grd_Listad.Row, 9) = 4 Or grd_Listad.TextMatrix(grd_Listad.Row, 9) = 5 Or grd_Listad.TextMatrix(grd_Listad.Row, 9) = 7 Or grd_Listad.TextMatrix(grd_Listad.Row, 9) = 14 Or grd_Listad.TextMatrix(grd_Listad.Row, 9) = 18 Or grd_Listad.TextMatrix(grd_Listad.Row, 9) = 21 Then
      If MsgBox("¿Está seguro de generar Orden de Liquidación?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
      Screen.MousePointer = 11
      cmd_ExpLiq.Enabled = False
      Call fs_GenExc_Liquidacion
      cmd_ExpLiq.Enabled = True
      Screen.MousePointer = 0
   Else
      MsgBox "Opción no habilitada", vbInformation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub fs_GenExc_Liquidacion()
Dim r_obj_Excel         As Excel.Application
Dim r_str_ParAux        As String
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
Dim r_int_NroFil        As Integer
Dim r_int_NoFlLi        As Integer
Dim r_int_NroIte        As Integer
Dim r_int_CodOpe         As Integer

   r_int_NroFil = 9
   r_int_CodOpe = grd_Listad.TextMatrix(grd_Listad.Row, 9)
   r_int_NroIte = grd_Listad.TextMatrix(grd_Listad.Row, 10)
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT MAERDE_NUMREF, MAERDE_IMPORT, MAERDE_OPEREF, MAERDE_TDOPRV, MAERDE_NDOPRV, MAERDE_CTACTE, "
   g_str_Parame = g_str_Parame & "         MAEETE_TDOREP, MAEETE_NDOREP, MAEETE_NOMREP, MAEETE_DIRREP, MAEETE_TELREP "
   g_str_Parame = g_str_Parame & "    FROM TPR_MAERDE "
   g_str_Parame = g_str_Parame & "         INNER JOIN TPR_MAECFI ON MAECFI_NUMREF = MAERDE_NUMREF "
   g_str_Parame = g_str_Parame & "         INNER JOIN TPR_MAEETE ON MAEETE_TIPDOC = MAECFI_TIPDOC AND MAEETE_NUMDOC = MAECFI_NUMDOC "
   g_str_Parame = g_str_Parame & "   WHERE MAERDE_NUMREF = '" & CStr(moddat_g_str_DesIte) & "' " 'moddat_g_str_NumFia
   g_str_Parame = g_str_Parame & "     AND MAERDE_NUMITE = " & r_int_NroIte & " "
        
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      Set r_obj_Excel = New Excel.Application
      r_obj_Excel.SheetsInNewWorkbook = 1
      r_obj_Excel.Workbooks.Add
      
      r_obj_Excel.ActiveSheet.PageSetup.LeftMargin = r_obj_Excel.ActiveSheet.Application.InchesToPoints(0.4)
      r_obj_Excel.ActiveSheet.PageSetup.RightMargin = r_obj_Excel.ActiveSheet.Application.InchesToPoints(0.3)
      With r_obj_Excel.ActiveSheet
         .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Arial Narrow"
         .Columns("A:AM").ColumnWidth = 3
         .Rows("1:71").RowHeight = 12
         
         .Range(.Cells(3, 1), .Cells(99, 99)).Font.Size = 10
         .Range(.Cells(3, 1), .Cells(99, 39)).ColumnWidth = 1.9
         .Columns("A").ColumnWidth = 0.7
         .Columns("AK").ColumnWidth = 0.7
         .Columns("AL").ColumnWidth = 0.7
                    
         .Range(.Cells(2, 2), .Cells(2, 39)).Merge
         .Range(.Cells(3, 3), .Cells(3, 39)).Merge
         .Cells(2, 2) = "EDPYME MICASITA SA"
         .Cells(3, 3) = "ORDEN DE LIQUIDACIÓN DE OPERACIONES PARA ET     "
         
         .Range(.Cells(2, 2), .Cells(3, 6)).Font.Bold = True
         .Range(.Cells(2, 2), .Cells(3, 6)).HorizontalAlignment = xlHAlignCenter
         .Range(.Cells(2, 2), .Cells(3, 6)).Font.Size = 11
         .Rows("2:2").EntireRow.AutoFit
         .Rows("3:3").EntireRow.AutoFit
         
         .Cells(5, 2) = "I. DATOS DEL CLIENTE"
         .Cells(5, 34) = "N°"
         .Range(.Cells(5, 2), .Cells(5, 35)).Font.Bold = True

         .Range(.Cells(5, 35), .Cells(5, 37)).Merge
         .Cells(7, 2) = "Cliente"
         .Cells(7, 7) = Trim(pnl_RazSoc.Caption)
         .Cells(7, 24) = "DOI"
         .Cells(7, 27) = "'" & Trim(pnl_NroDoc.Caption)
        ' .Range(.Cells(7, 27), .Cells(7, 37)).Merge
         
         .Cells(9, 2) = "Reptte. Legal"
         .Cells(9, 7) = Trim(g_rst_Princi!MAEETE_NOMREP)
         .Cells(9, 24) = "DNI"
         .Cells(9, 27) = "'" & g_rst_Princi!MAEETE_TDOREP & "-" & g_rst_Princi!MAEETE_NDOREP
         
         .Cells(11, 2) = "Dirección"
         .Cells(11, 7) = Trim(g_rst_Princi!MAEETE_DIRREP)
         .Cells(11, 24) = "Teléfono"
         .Cells(11, 27) = "'" & g_rst_Princi!MAEETE_TELREP
         
         .Cells(13, 2) = "Carta Fianza N°"
         .Cells(13, 7) = Trim(pnl_NumRef.Caption)
         .Cells(13, 24) = "Fecha"
         .Range(.Cells(13, 27), .Cells(13, 30)).Merge
         .Cells(13, 27) = "'" & Format(Now, "DD/MM/YYYY")
         
         .Cells(17, 2) = "II. INSTRUCCIONES PARA OPERACIONES"
         .Range(.Cells(17, 2), .Cells(17, 2)).Font.Bold = True
        
         .Cells(19, 3) = "'" & "1.-"
         .Cells(19, 4) = "Transferencia de Fondos"
        
         .Cells(21, 4) = "Desembolso al Cliente"
         .Range(.Cells(21, 4), .Cells(21, 17)).Merge
         .Cells(21, 22) = "Abono Cta. Empresa Supervisora"
         .Range(.Cells(21, 22), .Cells(21, 35)).Merge
         .Cells(23, 4) = "N° Cta. **"
         .Cells(23, 22) = "N° Cta."
         .Cells(25, 4) = "Importe S/."
         .Cells(25, 22) = "Importe S/."
         .Cells(27, 4) = "Concepto:"
         .Cells(27, 22) = "Empresa Sup."
         
         If r_int_CodOpe = 4 Or r_int_CodOpe = 21 Then
            .Cells(23, 8) = "'" & g_rst_Princi!MAERDE_CTACTE
            .Cells(25, 8) = "S/." & Format(g_rst_Princi!MAERDE_IMPORT, "###,###,##0.00")
            .Cells(25, 8).NumberFormat = "###,###,###,##0.00"
            .Range(.Cells(25, 8), .Cells(25, 13)).Merge
            .Cells(27, 8) = Format(g_rst_Princi!MAERDE_OPEREF, "###,###,##0.00")
          ElseIf r_int_CodOpe = 5 Then
            .Cells(23, 26) = "'" & g_rst_Princi!MAERDE_CTACTE
            .Cells(25, 26) = "S/." & Format(g_rst_Princi!MAERDE_IMPORT, "###,###,##0.00")
            .Cells(25, 26).NumberFormat = "###,###,###,##0.00"
            .Range(.Cells(25, 26), .Cells(25, 31)).Merge
            If IsNull(g_rst_Princi!MAERDE_NDOPRV) Then
               .Cells(27, 26) = ""
            Else
               .Cells(27, 26) = fs_BuscarProv(g_rst_Princi!MAERDE_TDOPRV, g_rst_Princi!MAERDE_NDOPRV)
            End If
          End If
         .Cells(30, 3) = "'" & "2.-"
         .Cells(30, 4) = "Cargo a la Cta Administradora por Otros conceptos "
         .Cells(32, 4) = "N° Cta."
         .Cells(32, 22) = "Descripción"
         .Cells(34, 4) = "Importe S/."
         .Cells(38, 3) = "Observaciones:"

         .Cells(56, 3) = "  Firma y Sello"
         .Range(.Cells(56, 3), .Cells(56, 9)).Merge
         .Range(.Cells(56, 3), .Cells(56, 9)).HorizontalAlignment = xlHAlignCenter
         
         .Cells(56, 12) = "  Firma y Sello"
         .Range(.Cells(56, 12), .Cells(56, 18)).Merge
         .Range(.Cells(56, 12), .Cells(56, 18)).HorizontalAlignment = xlHAlignCenter
         
         .Cells(56, 21) = "  Firma y Sello"
         .Range(.Cells(56, 21), .Cells(56, 27)).Merge
         .Range(.Cells(56, 21), .Cells(56, 27)).HorizontalAlignment = xlHAlignCenter
                 
         .Cells(56, 30) = "  Firma y Sello"
         .Range(.Cells(56, 30), .Cells(56, 36)).Merge
         .Range(.Cells(56, 30), .Cells(56, 36)).HorizontalAlignment = xlHAlignCenter
         
         With .Range(.Cells(1, 1), .Cells(62, 39))
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
         End With
         
         With .Range(.Cells(1, 1), .Cells(62, 39)).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .PatternTintAndShade = 0
         End With
         'PLOMO
         .Range(.Cells(5, 1), .Cells(5, 39)).Interior.Color = RGB(245, 245, 245)
         .Range(.Cells(17, 1), .Cells(17, 39)).Interior.Color = RGB(245, 245, 245)
         .Range(.Cells(21, 4), .Cells(21, 17)).Interior.Color = RGB(245, 245, 245)
         .Range(.Cells(21, 22), .Cells(21, 35)).Interior.Color = RGB(245, 245, 245)
         
         'NARANJA
         .Range(.Cells(7, 7), .Cells(7, 20)).Interior.Color = RGB(255, 229, 204)
         .Range(.Cells(7, 27), .Cells(7, 34)).Interior.Color = RGB(255, 229, 204)
         .Range(.Cells(9, 7), .Cells(9, 20)).Interior.Color = RGB(255, 229, 204)
         .Range(.Cells(9, 27), .Cells(9, 34)).Interior.Color = RGB(255, 229, 204)
         .Range(.Cells(11, 7), .Cells(11, 20)).Interior.Color = RGB(255, 229, 204)
         .Range(.Cells(11, 27), .Cells(11, 34)).Interior.Color = RGB(255, 229, 204)
         .Range(.Cells(13, 7), .Cells(13, 20)).Interior.Color = RGB(255, 229, 204)
         .Range(.Cells(13, 27), .Cells(13, 30)).Interior.Color = RGB(255, 229, 204)
         
         .Range(.Cells(23, 8), .Cells(23, 17)).Interior.Color = RGB(255, 229, 204)
         .Range(.Cells(23, 26), .Cells(23, 35)).Interior.Color = RGB(255, 229, 204)
         .Range(.Cells(25, 8), .Cells(25, 13)).Interior.Color = RGB(255, 229, 204)
         .Range(.Cells(25, 26), .Cells(25, 31)).Interior.Color = RGB(255, 229, 204)
         .Range(.Cells(27, 26), .Cells(27, 35)).Interior.Color = RGB(255, 229, 204)
         
         .Range(.Cells(32, 8), .Cells(32, 17)).Interior.Color = RGB(255, 229, 204)
         .Range(.Cells(34, 8), .Cells(34, 13)).Interior.Color = RGB(255, 229, 204)
         
         'VERDE
         .Range(.Cells(19, 2), .Cells(19, 38)).Interior.Color = RGB(220, 239, 220)
         .Range(.Cells(30, 2), .Cells(30, 38)).Interior.Color = RGB(220, 239, 220)
         
         With .Range(.Cells(5, 35), .Cells(5, 37))
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
         End With
         With .Range(.Cells(27, 8), .Cells(27, 20))
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
         End With
         With .Range(.Cells(32, 26), .Cells(32, 36))
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
         End With
         With .Range(.Cells(34, 26), .Cells(34, 36))
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
         End With
         With .Range(.Cells(38, 8), .Cells(38, 36))
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
         End With
         With .Range(.Cells(40, 8), .Cells(40, 36))
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
         End With
         With .Range(.Cells(42, 8), .Cells(42, 36))
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
         End With
         With .Range(.Cells(55, 3), .Cells(55, 10))
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
         End With
         With .Range(.Cells(55, 12), .Cells(55, 19))
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
         End With
         With .Range(.Cells(55, 21), .Cells(55, 28))
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
         End With
         With .Range(.Cells(55, 30), .Cells(55, 37))
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
         End With
      End With
      r_obj_Excel.Visible = True
   End If
End Sub

Private Sub cmd_Grabar_Click()
Dim r_dbl_ValImp  As Double
Dim r_str_DesGlo  As String
Dim r_int_NumIte  As Integer
Dim r_str_CtaDeb  As String
Dim r_str_CtaHab  As String
Dim r_dbl_ImpTot  As Double
Dim r_dbl_ImpSal  As Double
Dim r_dbl_ImpPag  As Double
Dim r_dbl_ImpCFi  As Double
Dim r_dbl_ImpGar  As Double
Dim r_int_CodBan  As Integer
Dim r_int_ImpDes  As Double
Dim r_int_Contad  As Integer
Dim r_bol_FlgCom  As Boolean
Dim r_dbl_ImpRet  As Double
Dim r_int_TdoPrv  As Integer
Dim r_str_NdoPrv  As String

   r_int_NumIte = 0
   
   'Validaciones
   If cmb_TipOper.ListIndex = -1 Then
      MsgBox "Debe seleccionar Tipo de Operación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipOper)
      Exit Sub
   End If
      
   If CInt(cmb_TipOper.ItemData(cmb_TipOper.ListIndex)) < 8 Or CInt(cmb_TipOper.ItemData(cmb_TipOper.ListIndex)) > 9 Then
'      If CDate(ipp_FecOpe.Text) > date Then
'         MsgBox "Debe ingresar una Fecha de Emisión válida.", vbExclamation, modgen_g_str_NomPlt
'         Call gs_SetFocus(ipp_FecOpe)
'         Exit Sub
'      End If
      If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 4 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 5 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 7 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 14 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 18 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 21 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 22 Then
         If cmb_ForPag.ListIndex = -1 Then
            MsgBox "Debe seleccionar Forma de Pago.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_ForPag)
            Exit Sub
         End If
      End If
      If CInt(cmb_TipOper.ItemData(cmb_TipOper.ListIndex)) <> 15 Then
         If ipp_ImpDes.Value = 0 Then
            MsgBox "Debe ingresar Monto desembolsado.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ImpDes)
            Exit Sub
         End If
      Else
         If pnl_ImpTot.Caption = 0 Then
            MsgBox "No se puede realizar el Extorno de Comisión porque Valor es Cero.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
      End If
      If cmb_Moneda.ListIndex = -1 Then
         MsgBox "Debe seleccionar Moneda.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Moneda)
         Exit Sub
      End If
      If ipp_ImpDes.Value > CDbl(pnl_ImpSal.Caption) Then
         MsgBox "El Monto desembolsado no puede ser mayor al Saldo adeudado.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ImpDes)
         Exit Sub
      End If
      If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 4 Then
         If ipp_ImpDes.Value > 300000 Then
            MsgBox "El importe a desembolsar supera los S/.300,000.00, por favor partir el pago", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ImpDes)
            Exit Sub
         End If
      End If
      If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 5 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 18 Then
         If cmb_TipDoc.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Tipo de Documento del Proveedor", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_TipDoc)
            Exit Sub
         End If
         If Len(Trim(txt_NumDoc.Text)) = 0 Then
            MsgBox "Debe ingresar el Número de Documento del Proveedor", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumDoc)
            Exit Sub
         End If
         If Len(Trim(Me.pnl_RScPrv.Caption)) = 0 Then
            MsgBox "Debe ingresar un documento válido", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumDoc)
            Exit Sub
         End If
      End If
      If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 4 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 5 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 7 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 14 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 18 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 21 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 22 Then
         If cmb_ForPag.ItemData(cmb_ForPag.ListIndex) = 2 Then
            If cmb_CodBan.ListIndex = -1 Then
               MsgBox "Debe seleccionar Banco.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(cmb_CodBan)
               Exit Sub
            End If
            If cmb_CtaBan.ListIndex = -1 Then
               MsgBox "Debe seleccionar Cuenta Corriente", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(cmb_CtaBan)
               Exit Sub
            End If
         Else
            If cmb_TipDoc.ListIndex = -1 Then
               MsgBox "Debe seleccionar el Tipo de Documento del Proveedor", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(cmb_TipDoc)
               Exit Sub
            End If
            If Len(Trim(txt_NumDoc.Text)) = 0 Then
               MsgBox "Debe ingresar el Número de Documento del Proveedor", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(txt_NumDoc)
               Exit Sub
            End If
            If Len(Trim(Me.pnl_RScPrv.Caption)) = 0 Then
               MsgBox "Debe ingresar un documento válido", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(txt_NumDoc)
               Exit Sub
            End If
         End If
      End If
   Else
'      If CDate(ipp_FecOpe.Text) > date Then
'         MsgBox "Debe ingresar una Fecha de Emisión válida.", vbExclamation, modgen_g_str_NomPlt
'         Call gs_SetFocus(ipp_FecOpe)
'         Exit Sub
'      End If
      If CInt(cmb_TipOper.ItemData(cmb_TipOper.ListIndex)) = 8 Then
         If fs_Validar_LiqCFi = False Then
            Exit Sub
         End If
      End If
      If CInt(cmb_TipOper.ItemData(cmb_TipOper.ListIndex)) = 9 Then
         If fs_Validar_CanCFi = False Then
            MsgBox "Verifique que el registro no tenga movimientos", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
      End If
   End If
   
   If (Format(ipp_FecOpe.Text, "yyyymmdd") < Format(moddat_g_str_FecIni, "yyyymmdd") Or Format(ipp_FecOpe.Text, "yyyymmdd") > Format(moddat_g_str_FecFin, "yyyymmdd")) Then
       MsgBox "No es posible registrar movimientos en un periodo errado.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(ipp_FecOpe)
       Exit Sub
   End If
   
   'Validaciones para Edición
   If moddat_g_int_FlgAct = 2 Then
      'Valida que los importes modificados de Comisiones no sean mayores al importe total de comisiones
      If CInt(pnl_CodOpe.Caption) = 1 Or CInt(pnl_CodOpe.Caption) = 2 Then
         If fs_Validar_Mto_ComPag = False Then
            MsgBox "El Importe Comisión es menor a las Comisiones pagadas.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ImpDes)
            Exit Sub
         End If
      End If
      
      'Valida que los importes modificados de Fondos Recibidos no sean mayores al importe total de Fondos
      If CInt(pnl_CodOpe.Caption) = 3 Or CInt(pnl_CodOpe.Caption) = 19 Then
         If fs_Validar_Mto_FRePag = False Then
            MsgBox "El Importe Fondos Recibidos es menor a lo Recibido y/o al Desemboldo pagado.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ImpDes)
            Exit Sub
         End If
      End If
       
       'Valida que los importes modificados de Desembolsos y/o pagos no sean mayores al importe de fondos recibidos
      If CInt(pnl_CodOpe.Caption) = 4 Or CInt(pnl_CodOpe.Caption) = 5 Then
         If fs_Validar_Mto_ValDes = False Then
            MsgBox "El Importe Desembolsado al Cliente es menor a los Fondos Recibidos.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ImpDes)
            Exit Sub
         End If
      End If
      
      'Valida que los importes modificados de garantía no sean mayores al importe de garantía
      If CInt(pnl_CodOpe.Caption) = 6 Or CInt(pnl_CodOpe.Caption) = 20 Then
         If fs_Validar_Mto_GarPag = False Then
            MsgBox "El Importe Garantía es menor a la Garantía pagada.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ImpDes)
            Exit Sub
         End If
      End If
   End If
   
   'Valida que el Monto ingresado para compensar sea menor e igual al Saldo con el que se cuenta
   If CInt(pnl_CodOpe.Caption) = 1 Then
      If fs_Validar_Compensacion = False Then
         MsgBox "El Importe ingresado es mayor al Saldo.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ImpDes)
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   If CInt(pnl_CodOpe.Caption) < 8 Or CInt(pnl_CodOpe.Caption) > 9 Then
      If moddat_g_int_FlgAct = 1 Then
         r_int_NumIte = fs_GeneraNumIte
      Else
         r_int_NumIte = l_int_NumIte
      End If
   End If
   
   If cmb_CodBan.ListIndex <> -1 Then
      r_int_CodBan = cmb_CodBan.ItemData(cmb_CodBan.ListIndex)
   End If
   
   If txt_NumDoc.Text <> "" And cmb_TipDoc.ListIndex <> -1 Then
      r_int_TdoPrv = cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
      r_str_NdoPrv = CStr(txt_NumDoc.Text)
   End If
   
   If CInt(cmb_TipOper.ItemData(cmb_TipOper.ListIndex)) <> 8 And CInt(cmb_TipOper.ItemData(cmb_TipOper.ListIndex)) <> 9 Then
      
      'Descuento al Fondo Recibido, la Comisión y el Porcentaje de Retención
      If CInt(cmb_TipOper.ItemData(cmb_TipOper.ListIndex)) = 3 Or CInt(cmb_TipOper.ItemData(cmb_TipOper.ListIndex)) = 11 Or CInt(cmb_TipOper.ItemData(cmb_TipOper.ListIndex)) = 12 Or CInt(cmb_TipOper.ItemData(cmb_TipOper.ListIndex)) = 19 Then
      
         'Saldo de la Comisión a pagar
         Call fs_Buscar_Saldo_Comision(r_dbl_ImpTot, r_dbl_ImpSal, r_dbl_ImpPag)
      
         If r_dbl_ImpSal > 0 And CDbl(ipp_ImpDes.Value) > CDbl(r_dbl_ImpSal) Then
            If (ipp_ImpDes.Value - r_dbl_ImpSal - (ipp_ImpDes.Value * l_dbl_PorRet)) > 0 Then
               'Ingreso a TPR_MAERDE - Importe de Comisión                          'pnl_NumRef.Caption
               Call fs_Ing_Maerde(r_int_NumIte, 2, Replace(moddat_g_str_DesIte, "-", ""), ipp_FecOpe.Value, cmb_Moneda.ItemData(cmb_Moneda.ListIndex), CDbl(r_dbl_ImpSal), Trim(txt_Refer.Text), r_int_CodBan, cmb_CtaBan.Text, pnl_CCIBan.Caption, moddat_g_int_FlgAct, Trim(txt_NumMov.Text))
            Else
               r_dbl_ImpSal = 0
            End If
         Else
            r_dbl_ImpSal = 0           'El Monto Recibido no cubre el saldo de la comisión
         End If
         
         'Ingreso a TPR_MAERDE - Importe de Retención (Fondos)
         r_int_NumIte = fs_GeneraNumIte
         r_dbl_ImpRet = Format(CDbl((ipp_ImpDes.Value) * l_dbl_PorRet), "###,###,###,##0.00")
      
         If r_dbl_ImpRet > 0 Then                                                'pnl_NumRef.Caption
            Call fs_Ing_Maerde(r_int_NumIte, 10, Replace(moddat_g_str_DesIte, "-", ""), ipp_FecOpe.Value, cmb_Moneda.ItemData(cmb_Moneda.ListIndex), CDbl(r_dbl_ImpRet), Trim(txt_Refer.Text), r_int_CodBan, cmb_CtaBan.Text, pnl_CCIBan.Caption, moddat_g_int_FlgAct, Trim(txt_NumMov.Text))
         End If
         
         'Ingreso a TPR_MAERDE - Importe de Fondos Recibidos (FMV)
          r_int_NumIte = fs_GeneraNumIte
         
         'El monto recibido debe ser mayor a la comisión para que se pueda descontar el monto de comisión
         If ipp_ImpDes.Value >= r_dbl_ImpSal Then
            r_int_ImpDes = Format(CDbl(ipp_ImpDes.Value) - CDbl(r_dbl_ImpSal) - CDbl((ipp_ImpDes.Value) * l_dbl_PorRet), "###,###,###,##0.00")
         Else
            r_int_ImpDes = Format(CDbl(ipp_ImpDes.Value) - CDbl((ipp_ImpDes.Value) * l_dbl_PorRet), "###,###,###,##0.00")
         End If
      
      ElseIf CInt(cmb_TipOper.ItemData(cmb_TipOper.ListIndex)) = 14 Then
        
        r_int_ImpDes = ipp_ImpDes.Value
        Call fs_Ing_Maerde(r_int_NumIte, cmb_TipOper.ItemData(cmb_TipOper.ListIndex), Replace(moddat_g_str_DesIte, "-", ""), ipp_FecOpe.Value, cmb_Moneda.ItemData(cmb_Moneda.ListIndex), CDbl(r_int_ImpDes * (-1)), Trim(txt_Refer.Text), r_int_CodBan, cmb_CtaBan.Text, pnl_CCIBan.Caption, moddat_g_int_FlgAct, Trim(txt_NumMov.Text), r_int_TdoPrv, IIf(r_str_NdoPrv = "", Empty, CStr(r_str_NdoPrv)))
        r_int_NumIte = fs_GeneraNumIte
        
        'Genera un nuevo registro de Fondo del Monto Retenido
        Call fs_Ing_Maerde(r_int_NumIte, 3, Replace(moddat_g_str_DesIte, "-", ""), ipp_FecOpe.Value, cmb_Moneda.ItemData(cmb_Moneda.ListIndex), CDbl(r_int_ImpDes), "DEVOLUCIÓN DE RETENCIÓN", r_int_CodBan, cmb_CtaBan.Text, pnl_CCIBan.Caption, moddat_g_int_FlgAct, Trim(txt_NumMov.Text), r_int_TdoPrv, IIf(r_str_NdoPrv = "", Empty, CStr(r_str_NdoPrv)))
      
      ElseIf CInt(cmb_TipOper.ItemData(cmb_TipOper.ListIndex)) = 15 Then
         
         r_int_ImpDes = pnl_ImpTot.Caption
         
         'Actualiza el Monto de Comisión en TPR_MAECFI, y el MAERDE_SITUAC = 3 (anulado) en el COdOpe = 2
         Call fs_Ing_MaeCfi(CStr(Replace(moddat_g_str_DesIte, "-", "")), moddat_g_int_TipDoc, CStr(moddat_g_str_NumDoc), CInt(pnl_CodOpe.Caption), "")
         
         'Ingresa el monto de la Comisión como FMV
         Call fs_Ing_Maerde(r_int_NumIte, 3, Replace(moddat_g_str_DesIte, "-", ""), ipp_FecOpe.Value, cmb_Moneda.ItemData(cmb_Moneda.ListIndex), CDbl(r_int_ImpDes), "MONTO DE COMISIÓN", r_int_CodBan, cmb_CtaBan.Text, pnl_CCIBan.Caption, moddat_g_int_FlgAct, Trim(txt_NumMov.Text), r_int_TdoPrv, IIf(r_str_NdoPrv = "", Empty, CStr(r_str_NdoPrv)))

         r_int_NumIte = fs_GeneraNumIte 'Para que ingrese el registro de Extorno de Comisión
         
'      ElseIf CInt(cmb_TipOper.ItemData(cmb_TipOper.ListIndex)) = 22 Then
        
'        r_int_ImpDes = ipp_ImpDes.Value
'        Call fs_Ing_Maerde(r_int_NumIte, cmb_TipOper.ItemData(cmb_TipOper.ListIndex), Replace(moddat_g_str_DesIte, "-", ""), ipp_FecOpe.Value, cmb_Moneda.ItemData(cmb_Moneda.ListIndex), CDbl(r_int_ImpDes * (-1)), Trim(txt_Refer.Text), r_int_CodBan, cmb_CtaBan.Text, pnl_CCIBan.Caption, moddat_g_int_FlgAct, Trim(txt_NumMov.Text), r_int_TdoPrv, IIf(r_str_NdoPrv = "", Empty, CStr(r_str_NdoPrv)))
'
      Else
         r_int_ImpDes = ipp_ImpDes.Value
      End If
      
      If r_int_ImpDes > 0 Then                                                   'pnl_NumRef.Caption
         If CInt(cmb_TipOper.ItemData(cmb_TipOper.ListIndex)) <> 14 Then         'And CInt(cmb_TipOper.ItemData(cmb_TipOper.ListIndex)) <> 22
            Call fs_Ing_Maerde(r_int_NumIte, cmb_TipOper.ItemData(cmb_TipOper.ListIndex), Replace(moddat_g_str_DesIte, "-", ""), ipp_FecOpe.Value, cmb_Moneda.ItemData(cmb_Moneda.ListIndex), CDbl(r_int_ImpDes), Trim(txt_Refer.Text), r_int_CodBan, cmb_CtaBan.Text, pnl_CCIBan.Caption, moddat_g_int_FlgAct, Trim(txt_NumMov.Text), r_int_TdoPrv, IIf(r_str_NdoPrv = "", Empty, CStr(r_str_NdoPrv)))
         End If
      End If
      
   Else                                               'CANCELACIÓN Y ANULACIÓN
      'Para Cancelación y Anulación                   'pnl_NumRef.Caption
      Call fs_Ing_MaeCfi(CStr(Replace(moddat_g_str_DesIte, "-", "")), moddat_g_int_TipDoc, CStr(moddat_g_str_NumDoc), IIf(CInt(pnl_CodOpe.Caption) = 8, 3, IIf(CInt(pnl_CodOpe.Caption) = 9, 4, 0)), ipp_FecOpe.Value)
      
      'Cuando se CANCELA, se libera el monto de la CF CANCELADA que se encuentre asociada en la Garantía
      If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = "008" Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = "018" Then 'pnl_NumRef.Caption
         Call fs_DesAsociar_NumRef(CStr(Replace(moddat_g_str_NumFia, "-", "")), moddat_g_int_TipDoc, CStr(moddat_g_str_NumDoc)) 'moddat_g_str_DesIte
      End If
   
   End If
                                                                                                                                                                    
   'Generar Asientos automáticos
   r_str_DesGlo = Trim(cmb_TipOper.Text)
   
   If moddat_g_str_CodPrd <> "008" Then      'CREDITOS INDIRECTOS
      If moddat_g_int_FlgGOK = True And moddat_g_int_FlgAct = 1 Then
         
         If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 1 Then
            r_str_CtaDeb = "251419010110"
            If Mid(moddat_g_str_DesIte, 1, 1) = 2 Then   'AD
               r_str_CtaHab = "151719010114"
            Else                                         'CF
               r_str_CtaHab = "151719010114"  '"151719010104"
            End If
            
         ElseIf cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 3 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 11 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 12 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 19 Then
            r_str_CtaDeb = "111301060102"
            r_str_CtaHab = "251419010110"                '"151719010112"
            r_dbl_ValImp = CDbl(ipp_ImpDes.Value)        'CDbl(r_int_ImpDes)
            
            If r_dbl_ValImp > 0 Then                     'And r_dbl_ImpSal > 0
               Call fs_GeneraAsiento(Trim(Replace(pnl_NumRef.Caption, "-", "")), r_str_DesGlo, CStr(moddat_g_str_NumDoc), CStr(moddat_g_str_NomCli), r_str_CtaDeb, r_str_CtaHab, CDbl(r_dbl_ValImp), r_int_NumIte, r_dbl_ImpSal)
            End If
   
         ElseIf cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 4 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 5 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 18 Then
            r_str_CtaDeb = "251419010110"
            r_str_CtaHab = "251419010109"                '"111301060102"
         ElseIf cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 6 Then
            r_str_CtaDeb = "111702090601"
            r_str_CtaHab = "251419010110"                '"151719010104"
         ElseIf cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 7 Then
            r_str_CtaDeb = "251419010110"                '"251419010109"
            r_str_CtaHab = "111702090601"
         ElseIf cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 13 Then
            r_str_CtaDeb = "111301060102"
            r_str_CtaHab = "151719010114"  '"151719010104"
         ElseIf cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 15 Then
            r_str_CtaDeb = "291102070101"
            r_str_CtaHab = "151719010114"  '"151719010104"
         ElseIf cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 20 Then
            r_str_CtaDeb = "111301060102"
            r_str_CtaHab = "251419010110"
         ElseIf cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 21 Then
            r_str_CtaDeb = "251419010110"
            r_str_CtaHab = "111301060102"
'         ElseIf cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 22 Then
'            r_str_CtaDeb = "251419010110"
'            r_str_CtaHab = "111301060102"
         Else
           GoTo Finalizar
         End If
         
         If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = "008" Then                   'CANCELACION
            
            'GARANTIZADO
            Call fs_Buscar_Importe_CarFia(r_dbl_ImpCFi, r_dbl_ImpGar)
            If r_dbl_ImpGar > 0 Then
               Call fs_GeneraAsiento(Trim(Replace(pnl_NumRef.Caption, "-", "")), r_str_DesGlo, CStr(moddat_g_str_NumDoc), CStr(moddat_g_str_NomCli), "721201010101", "711201010101", CDbl(r_dbl_ImpGar), r_int_NumIte)
            End If
            
            'SALDO COMISION
            Call fs_Buscar_Saldo_Comision(r_dbl_ImpTot, r_dbl_ImpSal, r_dbl_ImpPag)
            If r_dbl_ImpSal > 0 Then
               Call fs_GeneraAsiento(Trim(Replace(pnl_NumRef.Caption, "-", "")), r_str_DesGlo, CStr(moddat_g_str_NumDoc), CStr(moddat_g_str_NomCli), "291102070101", "151719010114", CDbl(r_dbl_ImpSal), r_int_NumIte)  '"151719010104"
            End If
             
            'IMPORTE COMISION
            If r_dbl_ImpPag > 0 Then
               Call fs_GeneraAsiento(Trim(Replace(pnl_NumRef.Caption, "-", "")), r_str_DesGlo, CStr(moddat_g_str_NumDoc), CStr(moddat_g_str_NomCli), "291102070101", "521102010101", CDbl(r_dbl_ImpPag), r_int_NumIte)
            End If
            
            'SALDO FONDOS RECIBIDOS
   '         Call fs_Buscar_Saldo_Fondos(r_dbl_ImpTot, r_dbl_ImpSal)
   '         If r_dbl_ImpSal > 0 Then
   '            Call fs_GeneraAsiento(Trim(Replace(pnl_NumRef.Caption, "-", "")), r_str_DesGlo, CStr(moddat_g_str_NumDoc), CStr(moddat_g_str_NomCli), "251419010110", "151719010112", CDbl(r_dbl_ImpSal), r_int_NumIte)
   '         End If
                   
         ElseIf cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = "009" Then               'ANULACION
            
            'GARANTIZADO
            Call fs_Buscar_Importe_CarFia(r_dbl_ImpCFi, r_dbl_ImpGar)
            If r_dbl_ImpGar > 0 Then
               Call fs_GeneraAsiento(Trim(Replace(pnl_NumRef.Caption, "-", "")), r_str_DesGlo, CStr(moddat_g_str_NumDoc), CStr(moddat_g_str_NomCli), "721201010101", "711201010101", CDbl(r_dbl_ImpGar), r_int_NumIte)
            End If
            
            'SALDO COMISION
            Call fs_Buscar_Saldo_Comision(r_dbl_ImpTot, r_dbl_ImpSal, r_dbl_ImpPag)
            If r_dbl_ImpSal > 0 Then
               Call fs_GeneraAsiento(Trim(Replace(pnl_NumRef.Caption, "-", "")), r_str_DesGlo, CStr(moddat_g_str_NumDoc), CStr(moddat_g_str_NomCli), "291102070101", "151719010114", CDbl(r_dbl_ImpSal), r_int_NumIte)  '"151719010104"
            End If
            
   '         'SALDO FONDOS RECIBIDOS
   '         Call fs_Buscar_Saldo_Fondos(r_dbl_ImpTot, r_dbl_ImpSal)
   '         If r_dbl_ImpSal > 0 Then
   '            Call fs_GeneraAsiento(Trim(Replace(pnl_NumRef.Caption, "-", "")), r_str_DesGlo, CStr(moddat_g_str_NumDoc), CStr(moddat_g_str_NomCli), "251419010110", "151719010112", CDbl(r_dbl_ImpSal), r_int_NumIte)
   '         End If
            
         ElseIf cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> "003" And cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> "019" Then              'OTROS MOVIMIENTOS
            r_dbl_ValImp = CDbl(ipp_ImpDes.Value)
            If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 15 Then
               r_dbl_ValImp = CDbl(pnl_ImpTot.Caption)
            End If
            If r_dbl_ValImp > 0 Then
               Call fs_GeneraAsiento(Trim(Replace(pnl_NumRef.Caption, "-", "")), r_str_DesGlo, CStr(moddat_g_str_NumDoc), CStr(moddat_g_str_NomCli), r_str_CtaDeb, r_str_CtaHab, CDbl(r_dbl_ValImp), r_int_NumIte, r_dbl_ImpSal, r_int_TdoPrv, IIf(r_str_NdoPrv = "", Empty, CStr(r_str_NdoPrv)))
            End If
         End If
      End If
   
   Else 'CREDITOS DIRECTOS
   
   End If
   
   If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = "007" Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = "021" Then
      Call fs_DesAsociar_NumRef(CStr(Replace(moddat_g_str_NumFia, "-", "")), moddat_g_int_TipDoc, CStr(moddat_g_str_NumDoc))
   End If
   
Finalizar:
   MsgBox "Actualización realizada satisfactoriamente.", vbInformation, modgen_g_con_PltPar
   Call fs_Limpia
   Call fs_Buscar
   Call fs_Activa(False)
   Call gs_SetFocus(cmd_Agrega)
   frm_Ges_TecPro_03.fs_Buscar_Creditos_Indirectos
   frm_Ges_TecPro_03.fs_Buscar_Creditos_Directos
   'Unload Me
End Sub

Private Function fs_Ing_Maerde(ByVal p_NumIte As Integer, ByVal p_TipOper As Integer, ByVal p_NumRef As String, ByVal p_FecOpe As String, ByVal p_TipMon As Integer, ByVal p_ImpDes As Double, _
                               ByVal p_Refer As String, ByVal p_CodBan As Integer, ByVal p_CtaBan As String, ByVal p_CCIBan As String, ByVal p_Flag As Integer, ByVal p_NumMov As String, Optional ByVal p_TipDoc As Integer, _
                               Optional ByVal p_NumDoc As String) As Integer
   fs_Ing_Maerde = False
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
      
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
'      Call moddat_gs_FecSis
      
      g_str_Parame = "USP_TPR_MAERDE ("
      g_str_Parame = g_str_Parame & CStr(p_NumIte) & ", "
      g_str_Parame = g_str_Parame & CStr(p_TipOper) & ", "
      g_str_Parame = g_str_Parame & "'" & CStr(Replace(p_NumRef, "-", "")) & "', "
      g_str_Parame = g_str_Parame & "'" & Format(p_FecOpe, "yyyymmdd") & "', "
      g_str_Parame = g_str_Parame & CStr(p_TipMon) & ", "
      g_str_Parame = g_str_Parame & CDbl(p_ImpDes) & ", "
      g_str_Parame = g_str_Parame & "'" & Trim(p_Refer) & "', "
      g_str_Parame = g_str_Parame & CStr(p_TipDoc) & ", "
      If p_NumDoc = Empty Then
         g_str_Parame = g_str_Parame & "'', "
      Else
         g_str_Parame = g_str_Parame & CStr(p_NumDoc) & ", "
      End If
      
      g_str_Parame = g_str_Parame & CStr(p_CodBan) & ", "
      If InStr(cmb_CtaBan.Text, "-") > 0 Then
         g_str_Parame = g_str_Parame & "'" & Trim(Mid(p_CtaBan, 1, InStr(p_CtaBan, "-") - 1)) & "', "
      Else
         g_str_Parame = g_str_Parame & "'" & Trim(p_CtaBan) & "', "
      End If
      g_str_Parame = g_str_Parame & "'" & Trim(p_CCIBan) & "', "
      If cmb_ForPag.ListIndex = -1 Then
         g_str_Parame = g_str_Parame & "0, "
      Else
         g_str_Parame = g_str_Parame & CStr(CStr(cmb_ForPag.ItemData(cmb_ForPag.ListIndex))) & ", "
      End If
      g_str_Parame = g_str_Parame & CStr(p_Flag) & ", "
      g_str_Parame = g_str_Parame & "'" & CStr(p_NumMov) & "', "

      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      Screen.MousePointer = 0
   Loop
   
   fs_Ing_Maerde = True
End Function

Private Function fs_DesAsociar_NumRef(ByVal p_NumRef As String, ByVal p_TipDoc As Integer, ByVal p_NumDoc As String) As Integer
   fs_DesAsociar_NumRef = False
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
      
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
'      Call moddat_gs_FecSis
               
      'Grabando Información Estado de Carta Fianza
      g_str_Parame = "USP_TPR_MAEGAR_NUMREF ("
      g_str_Parame = g_str_Parame & "'" & CStr(p_NumRef) & "', "
      g_str_Parame = g_str_Parame & CStr(p_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & CStr(p_NumDoc) & "', "

       'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function

         Else
            moddat_g_int_CntErr = 0
         End If
      End If

      Screen.MousePointer = 0
   Loop
   fs_DesAsociar_NumRef = True
End Function

Private Function fs_Ing_MaeCfi(ByVal p_NumRef As String, ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_Flag As Integer, ByVal p_FecCan As String) As Integer
   fs_Ing_MaeCfi = False
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
      
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
'      Call moddat_gs_FecSis
               
      'Grabando Información Estado de Carta Fianza
      g_str_Parame = "USP_TPR_MAECFI ("
      g_str_Parame = g_str_Parame & "'', "                                             'as_codprd
      g_str_Parame = g_str_Parame & "'', "                                             'as_subprd
      g_str_Parame = g_str_Parame & "'', "                                             'as_codmod
      g_str_Parame = g_str_Parame & "'" & CStr(p_NumRef) & "', "                       'as_numref
      g_str_Parame = g_str_Parame & "'', "                                             'as_emifia
      g_str_Parame = g_str_Parame & CStr(p_TipDoc) & ", "                              'as_tipdoc
      g_str_Parame = g_str_Parame & "'" & CStr(p_NumDoc) & "', "                       'as_numdoc
      g_str_Parame = g_str_Parame & "'', "                                             'as_plzfia
      g_str_Parame = g_str_Parame & "'', "                                             'as_vtofia
      g_str_Parame = g_str_Parame & "'', "                                             'as_monfia
      g_str_Parame = g_str_Parame & "'', "                                             'as_impfia
      g_str_Parame = g_str_Parame & "'', "                                             'as_garfia
      g_str_Parame = g_str_Parame & "'', "                                             'as_tasfia
      g_str_Parame = g_str_Parame & "'', "                                             'as_comfia
      g_str_Parame = g_str_Parame & "'', "                                             'as_minfia
      g_str_Parame = g_str_Parame & "'', "                                             'as_porret
      g_str_Parame = g_str_Parame & "'', "                                             'as_numant
      If p_FecCan = "" Then
         g_str_Parame = g_str_Parame & "'', "
      Else
         g_str_Parame = g_str_Parame & "'" & Format(p_FecCan, "yyyymmdd") & "', "      'as_feccan
      End If
      g_str_Parame = g_str_Parame & "'', "                                             'as_numade
      g_str_Parame = g_str_Parame & "'', "                                             'as_codpry
      g_str_Parame = g_str_Parame & "'', "                                             'as_parreg
      g_str_Parame = g_str_Parame & 0 & ", "                                           'as_codete
      g_str_Parame = g_str_Parame & 0 & ", "                                           'as_tiprec
      g_str_Parame = g_str_Parame & 0 & ", "                                           'as_nompry
      g_str_Parame = g_str_Parame & 0 & ", "                                           'as_portea
      g_str_Parame = g_str_Parame & 0 & ", "                                           'as_porgar
      g_str_Parame = g_str_Parame & 0 & ", "                                           'as_esclie
      g_str_Parame = g_str_Parame & 0 & ", "                                           'as_tiplin
      g_str_Parame = g_str_Parame & 0 & ", "                                           'as_lincre
      g_str_Parame = g_str_Parame & p_Flag & ", "                                      'as_insupd

       'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function

         Else
            moddat_g_int_CntErr = 0
         End If
      End If

      Screen.MousePointer = 0
   Loop
   fs_Ing_MaeCfi = True
End Function

Private Function fs_Validar_Mto_ComPag() As Boolean
Dim r_dbl_Mto_Pag As Double

   fs_Validar_Mto_ComPag = False
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT MAECFI_NUMREF, MAECFI_COMFIA COMISION, NVL(COMISION_PAGADO,0) COMISION_PAGADO "
   g_str_Parame = g_str_Parame & "    FROM TPR_MAECFI"
   g_str_Parame = g_str_Parame & "           LEFT JOIN (SELECT MAERDE_NUMREF, NVL(SUM(NVL(MAERDE_IMPORT,0)),0) COMISION_PAGADO "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE "
   g_str_Parame = g_str_Parame & "                       WHERE MAERDE_CODIGO = 1 OR MAERDE_CODIGO = 2 "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF) B ON MAERDE_NUMREF = MAECFI_NUMREF"
   g_str_Parame = g_str_Parame & "   WHERE MAECFI_NUMREF =  '" & CStr(moddat_g_str_DesIte) & "' " 'moddat_g_str_NumFia
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
     Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      fs_Validar_Mto_ComPag = True
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Function
   End If
     
   r_dbl_Mto_Pag = grd_Listad.TextMatrix(grd_Listad.Row, 3)
   
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      If CDbl(g_rst_GenAux!COMISION_PAGADO) - CDbl(r_dbl_Mto_Pag) + CDbl(ipp_ImpDes.Value) <= CDbl(g_rst_GenAux!COMISION) Then
         fs_Validar_Mto_ComPag = True
      End If
   End If
End Function

Private Function fs_Validar_Mto_FRePag() As Boolean
Dim r_dbl_Mto_Pag    As Double
Dim r_bol_Flag1      As Boolean
Dim r_bol_Flag2      As Boolean

   fs_Validar_Mto_FRePag = False
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT MAECFI_NUMREF, MAECFI_IMPFIA FONDOS, NVL(FONDOS_PAGADO,0) FONDOS_PAGADO "
   g_str_Parame = g_str_Parame & "    FROM TPR_MAECFI"
   g_str_Parame = g_str_Parame & "           LEFT JOIN (SELECT MAERDE_NUMREF, NVL(SUM(NVL(MAERDE_IMPORT,0)),0) FONDOS_PAGADO "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE "
   'g_str_Parame = g_str_Parame & "                       WHERE MAERDE_CODIGO = 3 "
   g_str_Parame = g_str_Parame & "                       WHERE MAERDE_CODIGO IN (3,19) "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF) B ON MAERDE_NUMREF = MAECFI_NUMREF"
   g_str_Parame = g_str_Parame & "   WHERE MAECFI_NUMREF =  '" & CStr(moddat_g_str_DesIte) & "' " 'moddat_g_str_NumFia
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
     Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      r_bol_Flag1 = True
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Function
   End If
     
   r_dbl_Mto_Pag = grd_Listad.TextMatrix(grd_Listad.Row, 3)
   
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      If CDbl(g_rst_GenAux!FONDOS_PAGADO) - CDbl(r_dbl_Mto_Pag) + CDbl(ipp_ImpDes.Value) <= CDbl(g_rst_GenAux!FONDOS) Then
         r_bol_Flag1 = True
      End If
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT SUM(A.MAERDE_IMPORT) DESEMBOLSO_RECIBIDO, NVL(DESEMBOLSO_PAGADO,0) DESEMBOLSO_PAGADO "
   g_str_Parame = g_str_Parame & "    FROM TPR_MAERDE A "
   g_str_Parame = g_str_Parame & "           LEFT JOIN (SELECT MAERDE_NUMREF, NVL(SUM(NVL(MAERDE_IMPORT,0)),0) DESEMBOLSO_PAGADO "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE "
   g_str_Parame = g_str_Parame & "                       WHERE (MAERDE_CODIGO = 4 OR MAERDE_CODIGO = 5) "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF) B ON B.MAERDE_NUMREF = A.MAERDE_NUMREF"
   g_str_Parame = g_str_Parame & "   WHERE A.MAERDE_NUMREF =  '" & CStr(moddat_g_str_DesIte) & "' " 'moddat_g_str_NumFia
   'g_str_Parame = g_str_Parame & "     AND MAERDE_CODIGO = 3 "
   g_str_Parame = g_str_Parame & "     AND MAERDE_CODIGO IN (3,19) "
   g_str_Parame = g_str_Parame & "   GROUP BY DESEMBOLSO_PAGADO "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
     Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      r_bol_Flag2 = True
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Function
   End If
       
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      If CDbl(g_rst_GenAux!DESEMBOLSO_PAGADO) <= CDbl(g_rst_GenAux!DESEMBOLSO_RECIBIDO) - CDbl(r_dbl_Mto_Pag) + CDbl(ipp_ImpDes.Value) Then
         r_bol_Flag2 = True
      End If
   End If
   If r_bol_Flag1 = True And r_bol_Flag2 = True Then
      fs_Validar_Mto_FRePag = True
   End If
End Function

Private Function fs_Validar_Mto_ValDes() As Boolean
Dim r_dbl_Mto_Pag As Double

   fs_Validar_Mto_ValDes = False
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT SUM(A.MAERDE_IMPORT) DESEMBOLSO_RECIBIDO, NVL(DESEMBOLSO_PAGADO,0) DESEMBOLSO_PAGADO "
   g_str_Parame = g_str_Parame & "    FROM TPR_MAERDE A "
   g_str_Parame = g_str_Parame & "           LEFT JOIN (SELECT MAERDE_NUMREF, NVL(SUM(NVL(MAERDE_IMPORT,0)),0) DESEMBOLSO_PAGADO "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE "
   g_str_Parame = g_str_Parame & "                       WHERE (MAERDE_CODIGO = 4 OR MAERDE_CODIGO = 5) "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF) B ON B.MAERDE_NUMREF = A.MAERDE_NUMREF"
   g_str_Parame = g_str_Parame & "   WHERE A.MAERDE_NUMREF =  '" & CStr(moddat_g_str_DesIte) & "' " 'moddat_g_str_NumFia
   'g_str_Parame = g_str_Parame & "     AND MAERDE_CODIGO = 3 "
   g_str_Parame = g_str_Parame & "     AND MAERDE_CODIGO IN (3,19) "
   g_str_Parame = g_str_Parame & "   GROUP BY DESEMBOLSO_PAGADO "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
     Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      fs_Validar_Mto_ValDes = True
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Function
   End If
     
   r_dbl_Mto_Pag = grd_Listad.TextMatrix(grd_Listad.Row, 3)
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      If CDbl(g_rst_GenAux!DESEMBOLSO_PAGADO) - CDbl(r_dbl_Mto_Pag) + CDbl(ipp_ImpDes.Value) <= CDbl(g_rst_GenAux!DESEMBOLSO_RECIBIDO) Then
         fs_Validar_Mto_ValDes = True
      End If
   End If
End Function

Private Function fs_Validar_Mto_GarPag() As Boolean
Dim r_dbl_Mto_Pag As Double

   fs_Validar_Mto_GarPag = False
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT MAEGAR_NUMREF, NVL(SUM(NVL(MAEGAR_MTOGAR_INM,0) + NVL(MAEGAR_MTOGAR_ES1,0) + NVL(MAEGAR_MTOGAR_ES2,0) + NVL(MAEGAR_MTOGAR_DE1,0) + NVL(MAEGAR_MTOGAR_DE2,0)),0) GARANTIA, "
   g_str_Parame = g_str_Parame & "         NVL((SELECT NVL(SUM(NVL(MAERDE_IMPORT,0)),0) "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAERDE "
   g_str_Parame = g_str_Parame & "               WHERE MAERDE_NUMREF = MAEGAR_NUMREF "
   g_str_Parame = g_str_Parame & "                 AND MAERDE_CODIGO = 6),0) GARANTIA_PAGADO "
   g_str_Parame = g_str_Parame & "    FROM TPR_MAEGAR "
   g_str_Parame = g_str_Parame & "   WHERE MAEGAR_NUMREF = '" & CStr(moddat_g_str_DesIte) & "' " 'moddat_g_str_NumFia
   g_str_Parame = g_str_Parame & "   GROUP BY MAEGAR_NUMREF "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      fs_Validar_Mto_GarPag = True
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Function
   End If
     
   r_dbl_Mto_Pag = grd_Listad.TextMatrix(grd_Listad.Row, 3)
   
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      If CDbl(g_rst_GenAux!GARANTIA_PAGADO) - CDbl(r_dbl_Mto_Pag) + CDbl(ipp_ImpDes.Value) <= CDbl(g_rst_GenAux!GARANTIA) Then
         fs_Validar_Mto_GarPag = True
      End If
   End If
End Function

Private Function fs_GeneraNumIte() As Integer
   fs_GeneraNumIte = 0
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT NVL(MAX(MAERDE_NUMITE),0) NUMITE FROM TPR_MAERDE WHERE MAERDE_NUMREF =  '" & CStr(Replace(moddat_g_str_DesIte, "-", "")) & "'" 'pnl_NumRef.Caption
    
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
       Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Function
   End If
     
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      fs_GeneraNumIte = g_rst_GenAux!NUMITE + 1
   End If
End Function

Public Sub fs_GeneraAsiento(ByVal p_NumRef As String, ByVal p_Glosa As String, ByVal p_NumDoc As String, ByVal p_RazSoc As String, _
                            ByVal p_CtaDeb As String, ByVal p_CtaHab As String, ByVal p_Importe As Double, ByVal p_NumIte As Integer, _
                            Optional ByVal p_SalCom As Double, Optional ByVal p_TdoPrv As Integer, Optional ByVal p_NdoPrv As String)
                            
Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
Dim r_int_NumIte        As Integer
Dim r_str_AsiGen        As String
Dim r_str_Origen        As String
Dim r_str_TipNot        As String
Dim r_int_NumLib        As Integer
Dim r_int_NumAsi        As Integer
Dim r_str_FecCon        As String
Dim r_str_FecReg        As String
Dim r_str_CtaCtb        As String
Dim r_str_DebHab        As String
Dim r_str_Glosa         As String
Dim r_dbl_MtoSol        As Double
Dim r_dbl_MtoDol        As Double
Dim r_dbl_importe       As Double
Dim r_dbl_TipCam        As Double
Dim r_int_ConAux        As Integer
Dim r_str_NroCnt        As String
Dim r_str_CodOpe        As String
Dim r_int_Contad        As Integer
Dim r_int_TipOpe        As Integer

   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1013"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   r_str_Origen = "LM"
   r_str_TipNot = "D"
   r_int_NumLib = 6
   r_str_AsiGen = ""
   r_int_NumAsi = 0
   r_int_NumIte = 0
   
   'Obteniendo Tipo de Cambio del día
   r_dbl_TipCam = moddat_gf_ObtieneTipCamDia(2, 2, Format(CDate(date), "yyyymmdd"), 2)
   
   'Obteniendo el Número de Asiento
   r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, CInt(moddat_g_str_CodAno), CInt(moddat_g_str_CodMes), r_str_Origen, r_int_NumLib)
   r_str_AsiGen = CStr(r_int_NumAsi)
   r_str_FecCon = CDate(ipp_FecOpe.Text)
   r_str_FecReg = moddat_g_str_FecSis
   r_str_Glosa = Trim(p_Glosa)
   
   'Insertar en CABECERA
   Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, CInt(moddat_g_str_CodAno), CInt(moddat_g_str_CodMes), r_int_NumLib, r_int_NumAsi, Format(1, "000"), r_dbl_TipCam, r_str_TipNot, Trim(r_str_Glosa), r_str_FecCon, "1")

   '*************************************************
   'GENERACION DE ASIENTOS CONTABLES DE OPERACIONES
   '*************************************************
   
   If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 3 And cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 19 Then
      r_int_TipOpe = cmb_TipOper.ItemData(cmb_TipOper.ListIndex)
      
      If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 6 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 7 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 20 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 21 Then
         r_int_Contad = 4
      Else
         r_int_Contad = 2
      End If
                  
      For r_int_ConAux = 1 To r_int_Contad '2
          r_dbl_importe = p_Importe
          
          If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 6 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 7 Then
            If r_int_ConAux = 1 Then r_str_DebHab = IIf(r_int_TipOpe = 6, "D", "H"): r_str_CtaCtb = IIf(r_int_TipOpe = 6, p_CtaDeb, p_CtaHab)
            If r_int_ConAux = 2 Then r_str_DebHab = IIf(r_int_TipOpe = 6, "H", "D"): r_str_CtaCtb = IIf(r_int_TipOpe = 6, p_CtaHab, p_CtaDeb)
            If r_int_ConAux = 3 Then r_str_DebHab = IIf(r_int_TipOpe = 6, "D", "H"): r_str_CtaCtb = IIf(r_int_TipOpe = 6, p_CtaHab, p_CtaDeb)
            If r_int_ConAux = 4 Then r_str_DebHab = IIf(r_int_TipOpe = 6, "H", "D"): r_str_CtaCtb = "111301060102"
          ElseIf cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 20 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 21 Then
            If r_int_ConAux = 1 Then r_str_DebHab = IIf(r_int_TipOpe = 20, "D", "H"): r_str_CtaCtb = IIf(r_int_TipOpe = 20, p_CtaDeb, p_CtaHab)
            If r_int_ConAux = 2 Then r_str_DebHab = IIf(r_int_TipOpe = 20, "H", "D"): r_str_CtaCtb = IIf(r_int_TipOpe = 20, p_CtaHab, p_CtaDeb)
            If r_int_ConAux = 3 Then r_str_DebHab = IIf(r_int_TipOpe = 20, "D", "H"): r_str_CtaCtb = "111702090601"
            If r_int_ConAux = 4 Then r_str_DebHab = IIf(r_int_TipOpe = 20, "H", "D"): r_str_CtaCtb = IIf(r_int_TipOpe = 20, p_CtaDeb, p_CtaHab)
          Else
            If r_int_ConAux = 1 Then r_str_DebHab = "D": r_str_CtaCtb = p_CtaDeb Else r_str_DebHab = "H": r_str_CtaCtb = p_CtaHab
          End If
          
          If Len(Trim(txt_NumMov.Text)) > 0 Then
            r_str_Glosa = IIf(Mid(p_NumRef, 1, 1) = 2, "AD", IIf(Mid(p_NumRef, 1, 1) = 3, "CSO", "CF")) & Trim(p_NumRef) & "/" & Trim(txt_NumMov.Text) & "/" & Trim(p_NumDoc) & "/" & Trim(p_RazSoc)
          Else
            r_str_Glosa = IIf(Mid(p_NumRef, 1, 1) = 2, "AD", IIf(Mid(p_NumRef, 1, 1) = 3, "CSO", "CF")) & Trim(p_NumRef) & "/" & Trim(p_NumDoc) & "/" & Trim(p_RazSoc)
          End If
          r_str_Glosa = Trim(Mid(r_str_Glosa, 1, 60))
         
          If (r_dbl_importe > 0) Then
              r_int_NumIte = r_int_NumIte + 1
              r_dbl_MtoSol = Format(r_dbl_importe, "###,###,##0.00")
              r_dbl_MtoDol = Format(0, "###,###,##0.00")
              
              Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, CInt(moddat_g_str_CodAno), CInt(moddat_g_str_CodMes), r_int_NumLib, r_int_NumAsi, r_int_NumIte, r_str_CtaCtb, CDate(r_str_FecCon), r_str_Glosa, r_str_DebHab, r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecCon))
              r_dbl_importe = 0
          End If
      Next r_int_ConAux

   Else
      If l_dbl_PorRet = 0 Then
         r_int_Contad = 3
      Else
         r_int_Contad = 4
      End If
   
      For r_int_ConAux = 1 To r_int_Contad
         If l_dbl_PorRet = 0 Then
            If r_int_ConAux = 1 Then r_dbl_importe = p_Importe + p_SalCom
            If r_int_ConAux = 2 Then r_dbl_importe = p_SalCom
            If r_int_ConAux = 3 Then r_dbl_importe = p_Importe
            
            If r_int_ConAux = 1 Then r_str_DebHab = "D":  r_str_CtaCtb = p_CtaDeb
            If r_int_ConAux = 2 Then r_str_DebHab = "H":  r_str_CtaCtb = "151719010114"  '"151719010104"
            If r_int_ConAux = 3 Then r_str_DebHab = "H":  r_str_CtaCtb = p_CtaHab
            
         Else
            If r_int_ConAux = 1 Or r_int_ConAux = 4 Then r_dbl_importe = p_Importe
            If r_int_ConAux = 2 Or r_int_ConAux = 3 Then r_dbl_importe = p_SalCom
            
            If r_int_ConAux = 1 Then r_str_DebHab = "D":  r_str_CtaCtb = p_CtaDeb
            If r_int_ConAux = 2 Then r_str_DebHab = "H":  r_str_CtaCtb = IIf(Mid(p_NumRef, 1, 1) = 2, "151719010114", "151719010114")  '"151719010104"
            If r_int_ConAux = 3 Then r_str_DebHab = "D":  r_str_CtaCtb = "251419010110"
            If r_int_ConAux = 4 Then r_str_DebHab = "H":  r_str_CtaCtb = p_CtaHab
         End If
         
         If Len(Trim(txt_NumMov.Text)) > 0 Then
            r_str_Glosa = IIf(Mid(p_NumRef, 1, 1) = 2, "AD", IIf(Mid(p_NumRef, 1, 1) = 3, "CSO", "CF")) & Trim(p_NumRef) & "/" & Trim(txt_NumMov.Text) & "/" & Trim(p_NumDoc) & "/" & Trim(p_RazSoc)
         Else
            r_str_Glosa = IIf(Mid(p_NumRef, 1, 1) = 2, "AD", IIf(Mid(p_NumRef, 1, 1) = 3, "CSO", "CF")) & Trim(p_NumRef) & "/" & Trim(p_NumDoc) & "/" & Trim(p_RazSoc)
         End If
         r_str_Glosa = Trim(Mid(r_str_Glosa, 1, 60))
         
         If (r_dbl_importe > 0) Then
             r_int_NumIte = r_int_NumIte + 1
             r_dbl_MtoSol = Format(r_dbl_importe, "###,###,##0.00")
             r_dbl_MtoDol = Format(0, "###,###,##0.00")
              
             Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, CInt(moddat_g_str_CodAno), CInt(moddat_g_str_CodMes), r_int_NumLib, r_int_NumAsi, r_int_NumIte, r_str_CtaCtb, CDate(r_str_FecCon), r_str_Glosa, r_str_DebHab, r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecCon))
             r_dbl_importe = 0
         End If
      Next r_int_ConAux
   End If
   r_str_NroCnt = r_str_Origen & "/" & moddat_g_str_CodAno & "/" & Format(moddat_g_str_CodMes, "00") & "/" & r_int_NumLib & "/" & r_int_NumAsi
   
   'Grabando en TPR_MAERDE, año/mes/nro_libro/nro_asiento
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE TPR_MAERDE "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET MAERDE_NROCNT = '" & CStr(r_str_NroCnt) & "'"
   modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE MAERDE_NUMREF = '" & p_NumRef & "'"
   
   If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 3 And cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 19 Then
      modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND MAERDE_NUMITE = " & p_NumIte & ""
   Else
      'modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND MAERDE_CODIGO IN (2,3,10) "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND MAERDE_CODIGO IN (2,3,10,19) "
   End If
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Grabar, 2) Then
      Exit Sub
   End If

   If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = "004" Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = "005" Or _
      cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = "018" Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = "021" Then
      r_str_CodOpe = modmip_gf_Genera_CodGen(3, 2)
      
      'Grabando en tra_gasadm, año/mes/nro_libro/nro_asiento
      modprc_g_str_CadEje = ""
      modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE TPR_MAERDE "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET MAERDE_CODOPE = " & r_str_CodOpe & ""
      modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE MAERDE_NUMREF = '" & p_NumRef & "'"
      modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND MAERDE_NUMITE = " & p_NumIte & ""
       
      If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Grabar, 2) Then
         Exit Sub
      End If
             
      'Agregando en la tabla CNTBL_COMAUT - PARA LAS APROBACIONES
      If p_TdoPrv = 0 Then
         Call fs_InsertaCompensacion(r_str_CodOpe, CStr(r_str_NroCnt), Trim(p_Glosa), moddat_g_int_TipDoc, moddat_g_str_NumDoc)
      Else
         Call fs_InsertaCompensacion(r_str_CodOpe, CStr(r_str_NroCnt), Trim(p_Glosa), p_TdoPrv, p_NdoPrv)
      End If
   End If
End Sub

Private Sub fs_InsertaCompensacion(ByVal p_CodOpe As String, ByVal p_DatCta As String, ByVal p_Descri As String, ByVal p_TdoPrv As Integer, ByVal p_NdoPrv As String)
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
           
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      
      g_str_Parame = "USP_CNTBL_COMAUT ("
      g_str_Parame = g_str_Parame & "'" & p_CodOpe & "', "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecOpe.Text), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & CStr(p_TdoPrv) & ", "                                                                  'moddat_g_int_TipDoc
      g_str_Parame = g_str_Parame & "'" & p_NdoPrv & "', "                                                                 'moddat_g_str_NumDoc
      g_str_Parame = g_str_Parame & 1 & ", "                                                                               'Tipo de Moneda
      g_str_Parame = g_str_Parame & CDbl(ipp_ImpDes.Text) & ", "
      
      If cmb_ForPag.ItemData(cmb_ForPag.ListIndex) = 1 Then
         g_str_Parame = g_str_Parame & "'', "
'      ElseIf cmb_ForPag.ItemData(cmb_ForPag.ListIndex) = 2 And moddat_g_str_CodMod = "008" Then
'         g_str_Parame = g_str_Parame & "'', "
      Else
         g_str_Parame = g_str_Parame & CStr(cmb_CodBan.ItemData(cmb_CodBan.ListIndex)) & ", "                              'Código del Banco - Proveedor
      End If
      If cmb_ForPag.ItemData(cmb_ForPag.ListIndex) = 1 Then
         g_str_Parame = g_str_Parame & "'', "
'      ElseIf cmb_ForPag.ItemData(cmb_ForPag.ListIndex) = 2 And moddat_g_str_CodMod = "008" Then
'         g_str_Parame = g_str_Parame & "'', "
      Else
         If CStr(cmb_CodBan.ItemData(cmb_CodBan.ListIndex)) = 11 Then
            g_str_Parame = g_str_Parame & "'" & Trim(Mid(cmb_CtaBan.Text, 1, InStr(cmb_CtaBan.Text, "-") - 1)) & "', "     'Cuenta Corriente - Proveedor
         Else
            g_str_Parame = g_str_Parame & "'" & Trim(pnl_CCIBan.Caption) & "', "                                           'CCI - Proveedor
         End If
      End If
      g_str_Parame = g_str_Parame & "'251419010109', "
      g_str_Parame = g_str_Parame & "'" & p_DatCta & "', "
      g_str_Parame = g_str_Parame & "'" & p_Descri & "', "
      g_str_Parame = g_str_Parame & 1 & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
           
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
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call gs_CentraForm(Me)
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Buscar
   Call fs_Activa(False)
   
   If moddat_g_str_DesObs <> "VIGENTE" Then
      cmd_Agrega.Enabled = False
      cmd_Borrar.Enabled = False
   Else
      moddat_g_str_DesObs = ""
   End If
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Moneda
   Call moddat_gs_Carga_LisIte_Combo(cmb_Moneda, 1, "204")
   
   'Tipo de Operación
   'Call moddat_gs_Carga_LisIte_Combo(cmb_TipOper, 1, "528")
   If moddat_g_str_TipCre = 1 And moddat_g_int_TipPan = 1 Then
      If CInt(Mid(moddat_g_str_NumFia, 1, 1)) <> 3 Then
         Call fs_Carga_TipOpe_Combo(cmb_TipOper, "528", 0)
      Else
         Call fs_Carga_TipOpe_Combo(cmb_TipOper, "528", 1)
      End If
   ElseIf moddat_g_str_TipCre = 2 And moddat_g_int_TipPan = 0 Then
      Call fs_Carga_TipOpe_Combo(cmb_TipOper, "528", 2)
   End If
   'Forma de Pago
    Call moddat_gs_Carga_LisIte_Combo(cmb_ForPag, 1, "531")
    
   'Bancos
   Call fs_Buscar_Banco(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   'Tipo de Documento
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "118")
     
'   'Año y mes
'   Call moddat_gf_ConsultaPerMesActivo("000001", 1, moddat_g_str_FecIni, moddat_g_str_FecFin, moddat_g_str_CodMes, moddat_g_str_CodAno)
   
   pnl_TipDoc.Caption = moddat_gf_Consulta_ParDes("118", moddat_g_int_TipDoc)
   pnl_NroDoc.Caption = moddat_g_str_NumDoc
   pnl_RazSoc.Caption = moddat_g_str_NomCli
   pnl_TipEmp.Caption = moddat_g_str_Descri
   
   If moddat_g_str_NumFia <> "" Then
      If moddat_g_str_CodPrd = "008" Then
         pnl_NumRef.Caption = gf_Formato_NumRef(moddat_g_str_NumFia, 1)
      Else
         pnl_NumRef.Caption = gf_Formato_NumRef(moddat_g_str_NumFia, Mid(moddat_g_str_NumFia, 1, 1))
      End If
   End If
   
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 3250
   grd_Listad.ColWidth(1) = 1560
   grd_Listad.ColWidth(2) = 2860
   grd_Listad.ColWidth(3) = 2080
   grd_Listad.ColWidth(4) = 2120
   grd_Listad.ColWidth(5) = 0
   grd_Listad.ColWidth(6) = 0
   grd_Listad.ColWidth(7) = 0
   grd_Listad.ColWidth(8) = 0
   grd_Listad.ColWidth(9) = 0
   grd_Listad.ColWidth(10) = 0
   grd_Listad.ColWidth(11) = 0
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_Listad.ColAlignment(4) = flexAlignRightCenter
End Sub

Private Sub fs_Carga_TipOpe_Combo(p_Combo As ComboBox, ByVal p_CodGrp As String, ByVal p_CodTip As Integer)
   p_Combo.Clear
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT PARDES_CODITE AS CODOPE, TRIM(PARDES_DESCRI) DESCRI "
   g_str_Parame = g_str_Parame & "   FROM MNT_PARDES "
   g_str_Parame = g_str_Parame & "  WHERE PARDES_CODGRP = '" & p_CodGrp & "' "
   
   If p_CodTip = 0 Then
      g_str_Parame = g_str_Parame & "    AND PARDES_CODITE NOT IN (0,2,10,16,17,20,21) " ',6,7
   ElseIf p_CodTip = 1 Then
       g_str_Parame = g_str_Parame & "    AND PARDES_CODITE IN (13,20,21,4,8,9)"
   ElseIf p_CodTip = 2 Then
      g_str_Parame = g_str_Parame & "   AND PARDES_CODITE IN (2,4,8,11)"
   End If
   
   g_str_Parame = g_str_Parame & "    AND PARDES_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "  ORDER BY PARDES_CODITE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     
     Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim$(g_rst_Genera!DESCRI)
      p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Genera!CODOPE)
      
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

'Private Function fs_Formato_NumRef(ByVal p_Numref As String) As String
'   p_Numref = Format(p_Numref, "0000000000")
'   'fs_Formato_NumRef = Left(p_Numref, 4) & "-" & Mid(p_Numref, 5, 2) & "-" & Right(p_Numref, 4)
'   fs_Formato_NumRef = Mid(p_Numref, 1, 1) & Mid(p_Numref, 2, 2) & "-" & Mid(p_Numref, 4, 2) & "-" & Right(p_Numref, 5)
'End Function

Function fs_Buscar_Importe_CarFia(ByRef p_ImpCFi, ByRef p_ImpGar)
   p_ImpCFi = 0
   p_ImpGar = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT MAECFI_IMPFIA, MAECFI_GARFIA "
   g_str_Parame = g_str_Parame & "    FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "   WHERE MAECFI_NUMREF = '" & CStr(moddat_g_str_DesIte) & "' " 'moddat_g_str_NumFia
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Function
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      If Trim(pnl_CodOpe.Caption) <> "" Then
         p_ImpCFi = Format(CDbl(g_rst_Princi!MAECFI_IMPFIA), "###,###,###,##0.00")
         p_ImpGar = Format(CDbl(g_rst_Princi!MAECFI_GARFIA), "###,###,###,##0.00")
      End If
   End If
End Function

Function fs_Buscar_Saldo_Comision(ByRef p_ImpTot As Double, ByRef p_ImpSal As Double, ByRef p_ImpPag As Double)  'As Double

   p_ImpTot = 0
   p_ImpSal = 0
 
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT A.MAECFI_COMFIA IMPORTE_COMISION, NVL(SUM(NVL(B.MAERDE_IMPORT,0)),0) IMPORTE_PAGADO "
   g_str_Parame = g_str_Parame & "    FROM TPR_MAECFI A "
   g_str_Parame = g_str_Parame & "         LEFT JOIN (SELECT B.MAERDE_NUMREF , B.MAERDE_NUMITE, B.MAERDE_IMPORT "
   g_str_Parame = g_str_Parame & "                      FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                     WHERE ( MAERDE_CODIGO = 1 OR MAERDE_CODIGO = 2 OR MAERDE_CODIGO = 13))B " '
   g_str_Parame = g_str_Parame & "                        ON B.MAERDE_NUMREF = A.MAECFI_NUMREF "
   g_str_Parame = g_str_Parame & "   WHERE A.MAECFI_NUMREF = '" & CStr(moddat_g_str_DesIte) & "' " ' moddat_g_str_NumFia
   g_str_Parame = g_str_Parame & "   GROUP BY MAECFI_COMFIA "
   
'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & "  SELECT "
'   g_str_Parame = g_str_Parame & "         CASE WHEN (NVL((SELECT NVL(SUM(MAECFI_COMFIA),0) "
'   g_str_Parame = g_str_Parame & "                           FROM TPR_MAECFI B "
'   g_str_Parame = g_str_Parame & "                          WHERE B.MAECFI_REFORI = A.MAECFI_REFORI "
'   g_str_Parame = g_str_Parame & "                          GROUP BY MAECFI_REFORI),0)) = 0 THEN MAECFI_COMFIA "
'   g_str_Parame = g_str_Parame & "         ELSE "
'   g_str_Parame = g_str_Parame & "                   NVL((SELECT NVL(SUM(MAECFI_COMFIA),0) "
'   g_str_Parame = g_str_Parame & "                          FROM TPR_MAECFI B "
'   g_str_Parame = g_str_Parame & "                         WHERE B.MAECFI_REFORI = A.MAECFI_REFORI "
'   g_str_Parame = g_str_Parame & "                         GROUP BY MAECFI_REFORI),0) "
'   g_str_Parame = g_str_Parame & "         END AS IMPORTE_COMISION, "
'   g_str_Parame = g_str_Parame & "         NVL(SUM(NVL(B.MAERDE_IMPORT,0)),0) IMPORTE_PAGADO "
'   g_str_Parame = g_str_Parame & "    FROM TPR_MAECFI A "
'   g_str_Parame = g_str_Parame & "         LEFT JOIN (SELECT B.MAERDE_NUMREF , B.MAERDE_NUMITE, B.MAERDE_IMPORT "
'   g_str_Parame = g_str_Parame & "                      FROM TPR_MAERDE B "
'   g_str_Parame = g_str_Parame & "                     WHERE (MAERDE_CODIGO = 1 OR MAERDE_CODIGO = 2))B "
'   g_str_Parame = g_str_Parame & "                        ON B.MAERDE_NUMREF = A.MAECFI_NUMREF "
'   g_str_Parame = g_str_Parame & "   WHERE A.MAECFI_NUMREF = '" & CStr(moddat_g_str_NumFia) & "' "
'   g_str_Parame = g_str_Parame & "   GROUP BY MAECFI_COMFIA , MAECFI_REFORI "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Function
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      If Trim(pnl_CodOpe.Caption) <> "" Then
         p_ImpTot = Format(CDbl(g_rst_Princi!IMPORTE_COMISION), "###,###,###,##0.00")
         p_ImpPag = Format(CDbl(g_rst_Princi!IMPORTE_PAGADO), "###,###,###,##0.00")
         p_ImpSal = Format(CDbl(g_rst_Princi!IMPORTE_COMISION) - CDbl(g_rst_Princi!IMPORTE_PAGADO), "###,###,###,##0.00")
      End If
   End If
End Function
Function fs_Buscar_Saldo_Comision_SinCompensacion(ByRef p_ImpTot As Double, ByRef p_ImpSal As Double, ByRef p_ImpPag As Double)  'As Double

   p_ImpTot = 0
   p_ImpSal = 0
 
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT A.MAECFI_COMFIA IMPORTE_COMISION, NVL(SUM(NVL(B.MAERDE_IMPORT,0)),0) IMPORTE_PAGADO "
   g_str_Parame = g_str_Parame & "    FROM TPR_MAECFI A "
   g_str_Parame = g_str_Parame & "         LEFT JOIN (SELECT B.MAERDE_NUMREF , B.MAERDE_NUMITE, B.MAERDE_IMPORT "
   g_str_Parame = g_str_Parame & "                      FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                     WHERE (MAERDE_CODIGO = 2 OR MAERDE_CODIGO = 13))B " '
   g_str_Parame = g_str_Parame & "                        ON B.MAERDE_NUMREF = A.MAECFI_NUMREF "
   g_str_Parame = g_str_Parame & "   WHERE A.MAECFI_NUMREF = '" & CStr(moddat_g_str_DesIte) & "' "
   g_str_Parame = g_str_Parame & "   GROUP BY MAECFI_COMFIA "
     
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Function
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      If Trim(pnl_CodOpe.Caption) <> "" Then
         p_ImpTot = Format(CDbl(g_rst_Princi!IMPORTE_COMISION), "###,###,###,##0.00")
         p_ImpPag = Format(CDbl(g_rst_Princi!IMPORTE_PAGADO), "###,###,###,##0.00")
         p_ImpSal = Format(CDbl(g_rst_Princi!IMPORTE_COMISION) - CDbl(g_rst_Princi!IMPORTE_PAGADO), "###,###,###,##0.00")
      End If
   End If
End Function
Function fs_Buscar_Extorno_Comision(ByRef p_ImpTot As Double, ByRef p_ImpSal As Double, ByRef p_ImpPag As Double)

   p_ImpTot = 0
   p_ImpSal = 0
 
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT A.MAECFI_COMFIA IMPORTE_COMISION, NVL(SUM(NVL(B.MAERDE_IMPORT,0)),0) IMPORTE_PAGADO "
   g_str_Parame = g_str_Parame & "    FROM TPR_MAECFI A "
   g_str_Parame = g_str_Parame & "         LEFT JOIN (SELECT B.MAERDE_NUMREF , B.MAERDE_NUMITE, B.MAERDE_IMPORT "
   g_str_Parame = g_str_Parame & "                      FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                     WHERE (MAERDE_CODIGO = 2) AND MAERDE_SITUAC = 1 )B "
   g_str_Parame = g_str_Parame & "                        ON B.MAERDE_NUMREF = A.MAECFI_NUMREF "
   g_str_Parame = g_str_Parame & "   WHERE A.MAECFI_NUMREF = '" & CStr(moddat_g_str_DesIte) & "' "
   g_str_Parame = g_str_Parame & "   GROUP BY MAECFI_COMFIA "
   
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Function
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      If Trim(pnl_CodOpe.Caption) <> "" Then
         p_ImpTot = Format(CDbl(g_rst_Princi!IMPORTE_COMISION), "###,###,###,##0.00")
         p_ImpPag = Format(CDbl(g_rst_Princi!IMPORTE_PAGADO), "###,###,###,##0.00")
         p_ImpSal = Format(CDbl(g_rst_Princi!IMPORTE_COMISION) - CDbl(g_rst_Princi!IMPORTE_PAGADO), "###,###,###,##0.00")
      End If
   End If
End Function
Function fs_Buscar_Saldo_Fondos(ByRef p_ImpTot As Double, ByRef p_ImpSal As Double)
Dim r_int_Contad     As Integer
Dim r_dbl_ImpTot     As Double
Dim r_dbl_ImpSal     As Double
Dim r_dbl_ImpPag     As Double
Dim r_dbl_PorRet     As Double
Dim r_dbl_ImpRet     As Double
Dim r_dbl_PAGCLI     As Double
   
   g_str_Parame = ""
   'SI HA SIDO RENOVADA Y ESTÁ VIGENTE

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT CASE WHEN A.MAECFI_NUMREN > 0 THEN "
   g_str_Parame = g_str_Parame & "                       NVL(( SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                               FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                              WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF "
   g_str_Parame = g_str_Parame & "                                AND MAERDE_CODIGO = 16 "
   g_str_Parame = g_str_Parame & "                                AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                              GROUP BY MAERDE_NUMREF),0) "
   
   g_str_Parame = g_str_Parame & "          ELSE"
   g_str_Parame = g_str_Parame & "                       A.MAECFI_IMPFIA "
   g_str_Parame = g_str_Parame & "          END                                                                                              IMPORTE_FONDOS          , "
   g_str_Parame = g_str_Parame & "          NVL(B.IMPORT,0)                                                                                  IMPORTE_PAGADO          , "
   g_str_Parame = g_str_Parame & "          MAECFI_PORRET                                                                                                            , "
   g_str_Parame = g_str_Parame & "          NVL(C.IMPORT,0)                                                                                  IMPORTE_RETENIDO        , "
   g_str_Parame = g_str_Parame & "          NVL(D.IMPORT,0)                                                                                  PAGO_COMISION_CLIENTE   , "
   g_str_Parame = g_str_Parame & "          NVL(E.IMPORT,0)                                                                                  DEVOLUCION_RETENCION       "
   g_str_Parame = g_str_Parame & "     FROM TPR_MAECFI A "
   g_str_Parame = g_str_Parame & "          LEFT JOIN (SELECT B.MAERDE_NUMREF , SUM(B.MAERDE_IMPORT) AS IMPORT "
   g_str_Parame = g_str_Parame & "                       FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                      WHERE (MAERDE_CODIGO = 3 OR MAERDE_CODIGO = 19 OR MAERDE_CODIGO = 11 OR MAERDE_CODIGO = 12)"
   g_str_Parame = g_str_Parame & "                   GROUP BY B.MAERDE_NUMREF )B "
   g_str_Parame = g_str_Parame & "                         ON B.MAERDE_NUMREF = A.MAECFI_NUMREF "
   g_str_Parame = g_str_Parame & "          LEFT JOIN (SELECT C.MAERDE_NUMREF , SUM(C.MAERDE_IMPORT) AS IMPORT "
   g_str_Parame = g_str_Parame & "                       FROM TPR_MAERDE C "
   g_str_Parame = g_str_Parame & "                      WHERE (C.MAERDE_CODIGO = 10)"
   g_str_Parame = g_str_Parame & "                   GROUP BY C.MAERDE_NUMREF )C "
   g_str_Parame = g_str_Parame & "                         ON C.MAERDE_NUMREF = A.MAECFI_NUMREF "
   g_str_Parame = g_str_Parame & "          LEFT JOIN (SELECT D.MAERDE_NUMREF , SUM(D.MAERDE_IMPORT) AS IMPORT "
   g_str_Parame = g_str_Parame & "                       FROM TPR_MAERDE D "
   g_str_Parame = g_str_Parame & "                      WHERE (D.MAERDE_CODIGO = 13)"
   g_str_Parame = g_str_Parame & "                      GROUP BY D.MAERDE_NUMREF )D "
   g_str_Parame = g_str_Parame & "                         ON D.MAERDE_NUMREF = A.MAECFI_NUMREF "
   g_str_Parame = g_str_Parame & "          LEFT JOIN (SELECT E.MAERDE_NUMREF , SUM(E.MAERDE_IMPORT ) AS IMPORT "
   g_str_Parame = g_str_Parame & "                       FROM TPR_MAERDE E "
   g_str_Parame = g_str_Parame & "                      WHERE (E.MAERDE_CODIGO = 14)"
   g_str_Parame = g_str_Parame & "                      GROUP BY E.MAERDE_NUMREF )E "
   g_str_Parame = g_str_Parame & "                         ON E.MAERDE_NUMREF = A.MAECFI_NUMREF "
   g_str_Parame = g_str_Parame & "    WHERE A.MAECFI_NUMREF = '" & CStr(moddat_g_str_DesIte) & "'"
   g_str_Parame = g_str_Parame & "    GROUP BY A.MAECFI_IMPFIA, A.MAECFI_PORRET, A.MAECFI_SITUAC, A.MAECFI_NUMREF, A.MAECFI_TIPDOC, A.MAECFI_NUMDOC, A.MAECFI_NUMREN, B.IMPORT, C.IMPORT, D.IMPORT, E.IMPORT "
   
    
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Function
   End If

   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Function
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      If Trim(pnl_CodOpe.Caption) <> "" Then
         l_dbl_PorRet = CDbl(g_rst_Genera!MAECFI_PORRET) / 100
         
         p_ImpTot = Format(CDbl(g_rst_Genera!IMPORTE_FONDOS), "###,###,###,##0.00")
         p_ImpSal = Format(CDbl(g_rst_Genera!IMPORTE_FONDOS) - CDbl(g_rst_Genera!IMPORTE_PAGADO), "###,###,###,##0.00")
         r_dbl_ImpRet = Format(CDbl(g_rst_Genera!IMPORTE_RETENIDO) + CDbl(g_rst_Genera!DEVOLUCION_RETENCION), "###,###,###,##0.00")
         r_dbl_PAGCLI = CDbl(g_rst_Genera!PAGO_COMISION_CLIENTE)
          
         Call fs_Buscar_Saldo_Comision_SinCompensacion(r_dbl_ImpTot, r_dbl_ImpSal, r_dbl_ImpPag)
         
         If CDbl(r_dbl_PAGCLI) = 0 Then   'debe descontarse cuando la comisión se haya descontado por fondos recibidos y no abonada por el cliente
               p_ImpSal = Format(CDbl(g_rst_Genera!IMPORTE_FONDOS) - (CDbl(g_rst_Genera!IMPORTE_PAGADO) + CDbl(r_dbl_ImpRet) + CDbl(r_dbl_ImpPag)), "###,###,###,##0.00")
         Else
               p_ImpSal = Format(CDbl(g_rst_Genera!IMPORTE_FONDOS) - CDbl(g_rst_Genera!IMPORTE_PAGADO) - CDbl(r_dbl_ImpRet), "###,###,###,##0.00")
         End If
         If p_ImpSal < 0 Then p_ImpSal = 0
        
      End If
   End If
End Function
Private Sub fs_Buscar()
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_RPT_SALTPR_CARFIA ("
   g_str_Parame = g_str_Parame & "'" & Format(Month(Now), "00") & "', "
   g_str_Parame = g_str_Parame & "'" & Format(Year(Now), "0000") & "', "
   g_str_Parame = g_str_Parame & "'REPORTE SALDOS TRP_CARFIA', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      
   If Trim(pnl_CodOpe.Caption) = "" Then
      g_str_Parame = g_str_Parame & "'" & CStr(moddat_g_str_DesIte) & "'," 'moddat_g_str_NumFia
      g_str_Parame = g_str_Parame & " 1 ,"
      g_str_Parame = g_str_Parame & "'') "
   Else
      g_str_Parame = g_str_Parame & "'" & CStr(moddat_g_str_DesIte) & "'," 'moddat_g_str_NumFia
      g_str_Parame = g_str_Parame & " 2,"
      g_str_Parame = g_str_Parame & CInt(CStr(pnl_CodOpe.Caption)) & ") "
   End If
      
   DoEvents: DoEvents: DoEvents
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   DoEvents: DoEvents: DoEvents
      
   grd_Listad.Redraw = False
   Call gs_LimpiaGrid(grd_Listad)
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     grd_Listad.Redraw = True
     Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = moddat_gf_Consulta_ParDes("528", g_rst_Princi!OPERACION)
            
         grd_Listad.Col = 1
         If Not IsNull(g_rst_Princi!FECASG) Then
            grd_Listad.Text = Format(gf_FormatoFecha(CStr(g_rst_Princi!FECASG)), "dd/mm/yyyy")
         Else
            grd_Listad.Text = ""
         End If
         
         grd_Listad.Col = 2
         grd_Listad.Text = moddat_gf_Consulta_ParDes("204", g_rst_Princi!Moneda)
         
         grd_Listad.Col = 3
         grd_Listad.Text = Format(CStr(g_rst_Princi!PAGADO), "###,###,###,##0.00")
         
         grd_Listad.Col = 4
         grd_Listad.Text = Format(CStr(g_rst_Princi!SALDO), "###,###,###,##0.00")
         
         grd_Listad.Col = 5
         grd_Listad.Text = IIf(IsNull(g_rst_Princi!REFERENCIA), "", g_rst_Princi!REFERENCIA)
         
         grd_Listad.Col = 6
         grd_Listad.Text = IIf(IsNull(g_rst_Princi!CODBAN), "", g_rst_Princi!CODBAN)
         
         grd_Listad.Col = 7
         grd_Listad.Text = IIf(IsNull(g_rst_Princi!CTACTE), "", g_rst_Princi!CTACTE)
         
         grd_Listad.Col = 8
         grd_Listad.Text = IIf(IsNull(g_rst_Princi!NUMCCI), "", g_rst_Princi!NUMCCI)
         
         grd_Listad.Col = 9
         grd_Listad.Text = CInt(g_rst_Princi!OPERACION)
         
         grd_Listad.Col = 10
         grd_Listad.Text = CInt(g_rst_Princi!NUMITE)
         
         grd_Listad.Col = 11
         grd_Listad.Text = IIf(IsNull(g_rst_Princi!NUMMOV), "", g_rst_Princi!NUMMOV)
         
         g_rst_Princi.MoveNext
      Loop
   End If
 
   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)
End Sub

Private Sub fs_Buscar_Saldo_Desembolso()
Dim r_dbl_ImpSal     As Double
Dim r_dbl_ImpOpe     As Double
   
   pnl_ImpTot.Caption = Format(0, "###,###,###,##0.00") & "  "
   pnl_ImpSal.Caption = Format(0, "###,###,###,##0.00") & "  "
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT CASE WHEN (SELECT B.MAECFI_NUMREN "
   g_str_Parame = g_str_Parame & "                      FROM TPR_MAECFI B "
   g_str_Parame = g_str_Parame & "                     WHERE B.MAECFI_NUMREF = A.MAERDE_NUMREF ) > 0 THEN"

   g_str_Parame = g_str_Parame & "                      NVL(( SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                              FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                             WHERE B.MAERDE_NUMREF = A.MAERDE_NUMREF "
   'g_str_Parame = g_str_Parame & "                               AND (B.MAERDE_CODIGO = 2 OR B.MAERDE_CODIGO = 3 OR B.MAERDE_CODIGO = 10 OR B.MAERDE_CODIGO = 11 OR B.MAERDE_CODIGO = 12 OR B.MAERDE_CODIGO = 14 OR B.MAERDE_CODIGO = 17 )"
   g_str_Parame = g_str_Parame & "                               AND (B.MAERDE_CODIGO = 2 OR B.MAERDE_CODIGO = 3 OR B.MAERDE_CODIGO = 19 OR B.MAERDE_CODIGO = 10 OR B.MAERDE_CODIGO = 11 OR B.MAERDE_CODIGO = 12 OR B.MAERDE_CODIGO = 14 OR B.MAERDE_CODIGO = 17 )"
   g_str_Parame = g_str_Parame & "                               AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                             GROUP BY MAERDE_NUMREF),0)                                                                         "
   g_str_Parame = g_str_Parame & "         ELSE                                                                                                                   "
  'g_str_Parame = g_str_Parame & "                      NVL(SUM(A.MAERDE_IMPORT), 0)                                                                              "
   
   g_str_Parame = g_str_Parame & "                      NVL(( SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                              FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                             WHERE B.MAERDE_NUMREF = A.MAERDE_NUMREF "
   g_str_Parame = g_str_Parame & "                               AND (B.MAERDE_CODIGO = 2 OR B.MAERDE_CODIGO = 3 OR B.MAERDE_CODIGO = 19 OR B.MAERDE_CODIGO = 10 OR B.MAERDE_CODIGO = 11 OR B.MAERDE_CODIGO = 12 OR B.MAERDE_CODIGO = 14) "
   g_str_Parame = g_str_Parame & "                               AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                             GROUP BY MAERDE_NUMREF),0)  "
   
   g_str_Parame = g_str_Parame & "         END                                                                                              IMPORTE_RECIBIDO,     "
   g_str_Parame = g_str_Parame & "         NVL((IMPORTE_PAGADO),0)                                                                          IMPORTE_PAGADO,       "
   g_str_Parame = g_str_Parame & "         NVL((DEVOLUCION_GARANTIA),0)                                                                     DEVOLUCION_GARANTIA , "
   g_str_Parame = g_str_Parame & "         NVL((DEVOLUCION_FONDO_CLIENTE),0)                                                                DEVOLUCION_FONDO_CLIENTE , "
   g_str_Parame = g_str_Parame & "         NVL((EXTORNO_COMISION),0)                                                                        EXTORNO_COMISION "
   
   g_str_Parame = g_str_Parame & "    FROM TPR_MAERDE A "
   g_str_Parame = g_str_Parame & "         LEFT JOIN (SELECT B.MAERDE_NUMREF, NVL(SUM(NVL(B.MAERDE_IMPORT,0)),0) IMPORTE_PAGADO "
   g_str_Parame = g_str_Parame & "                      FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                     WHERE (MAERDE_CODIGO = 1 OR MAERDE_CODIGO = 2 OR MAERDE_CODIGO = 4 OR MAERDE_CODIGO = 5 OR MAERDE_CODIGO = 10 OR MAERDE_CODIGO = 14 OR MAERDE_CODIGO = 6 OR MAERDE_CODIGO = 18) "
   g_str_Parame = g_str_Parame & "                     GROUP BY B.MAERDE_NUMREF ) B "
   g_str_Parame = g_str_Parame & "                        ON B.MAERDE_NUMREF = A.MAERDE_NUMREF "
   g_str_Parame = g_str_Parame & "         LEFT JOIN (SELECT C.MAERDE_NUMREF, NVL(SUM(NVL(C.MAERDE_IMPORT,0)),0) DEVOLUCION_GARANTIA "
   g_str_Parame = g_str_Parame & "                      FROM TPR_MAERDE C "
   g_str_Parame = g_str_Parame & "                     WHERE C.MAERDE_CODIGO = 7 "
   g_str_Parame = g_str_Parame & "                     GROUP BY C.MAERDE_NUMREF ) C "
   g_str_Parame = g_str_Parame & "                        ON C.MAERDE_NUMREF = A.MAERDE_NUMREF "
   g_str_Parame = g_str_Parame & "         LEFT JOIN (SELECT D.MAERDE_NUMREF, NVL(SUM(NVL(D.MAERDE_IMPORT,0)),0) EXTORNO_COMISION "
   g_str_Parame = g_str_Parame & "                      FROM TPR_MAERDE D "
   g_str_Parame = g_str_Parame & "                     WHERE D.MAERDE_CODIGO = 15 "
   g_str_Parame = g_str_Parame & "                     GROUP BY D.MAERDE_NUMREF ) D "
   g_str_Parame = g_str_Parame & "                        ON D.MAERDE_NUMREF = A.MAERDE_NUMREF "
   g_str_Parame = g_str_Parame & "         LEFT JOIN (SELECT E.MAERDE_NUMREF, NVL(SUM(NVL(E.MAERDE_IMPORT,0)),0) DEVOLUCION_FONDO_CLIENTE "
   g_str_Parame = g_str_Parame & "                      FROM TPR_MAERDE E "
   g_str_Parame = g_str_Parame & "                     WHERE E.MAERDE_CODIGO = 22 "
   g_str_Parame = g_str_Parame & "                     GROUP BY E.MAERDE_NUMREF ) E "
   g_str_Parame = g_str_Parame & "                        ON E.MAERDE_NUMREF = A.MAERDE_NUMREF "
   g_str_Parame = g_str_Parame & "   WHERE A.MAERDE_NUMREF = '" & CStr(moddat_g_str_DesIte) & "'"
   g_str_Parame = g_str_Parame & "   GROUP BY IMPORTE_PAGADO, DEVOLUCION_GARANTIA, EXTORNO_COMISION, A.MAERDE_NUMREF, DEVOLUCION_FONDO_CLIENTE "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      If Trim(pnl_CodOpe.Caption) <> "" Then
         pnl_ImpTot.Caption = Format(CDbl(g_rst_Princi!IMPORTE_RECIBIDO), "###,###,###,##0.00") & "  "  '- CDbl(g_rst_Princi!DEVOLUCION_GARANTIA)
         r_dbl_ImpSal = CDbl(g_rst_Princi!IMPORTE_RECIBIDO) - CDbl(g_rst_Princi!IMPORTE_PAGADO) + CDbl(g_rst_Princi!DEVOLUCION_GARANTIA) + CDbl(g_rst_Princi!DEVOLUCION_FONDO_CLIENTE) '+ CDbl(g_rst_Princi!EXTORNO_COMISION)
         pnl_ImpSal.Caption = Format(CDbl(r_dbl_ImpSal), "###,###,###,##0.00") & "  "
      End If
   End If
End Sub
Function fs_Validar_Compensacion() As Boolean
Dim r_dbl_ImpSal  As Double
   
   fs_Validar_Compensacion = False
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT CASE WHEN (SELECT B.MAECFI_NUMREN "
   g_str_Parame = g_str_Parame & "                      FROM TPR_MAECFI B "
   g_str_Parame = g_str_Parame & "                     WHERE B.MAECFI_NUMREF = A.MAERDE_NUMREF ) > 0 THEN"

   g_str_Parame = g_str_Parame & "                      NVL(( SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                              FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                             WHERE B.MAERDE_NUMREF = A.MAERDE_NUMREF "
   g_str_Parame = g_str_Parame & "                               AND (B.MAERDE_CODIGO = 2 OR B.MAERDE_CODIGO = 3 OR B.MAERDE_CODIGO = 19 OR B.MAERDE_CODIGO = 10 OR B.MAERDE_CODIGO = 11 OR B.MAERDE_CODIGO = 12 OR B.MAERDE_CODIGO = 14 OR B.MAERDE_CODIGO = 17 )"
   g_str_Parame = g_str_Parame & "                               AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                             GROUP BY MAERDE_NUMREF),0)                                                                         "
   g_str_Parame = g_str_Parame & "         ELSE                                                                                                                   "
   
   g_str_Parame = g_str_Parame & "                      NVL(( SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                              FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                             WHERE B.MAERDE_NUMREF = A.MAERDE_NUMREF "
   g_str_Parame = g_str_Parame & "                               AND (B.MAERDE_CODIGO = 2 OR B.MAERDE_CODIGO = 3 OR B.MAERDE_CODIGO = 19 OR B.MAERDE_CODIGO = 10 OR B.MAERDE_CODIGO = 11 OR B.MAERDE_CODIGO = 12 OR B.MAERDE_CODIGO = 14) "
   g_str_Parame = g_str_Parame & "                               AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                             GROUP BY MAERDE_NUMREF),0)  "
   
   g_str_Parame = g_str_Parame & "         END                                                                                              IMPORTE_RECIBIDO,     "
   g_str_Parame = g_str_Parame & "         NVL((IMPORTE_PAGADO),0)                                                                          IMPORTE_PAGADO,       "
   g_str_Parame = g_str_Parame & "         NVL((DEVOLUCION_GARANTIA),0)                                                                     DEVOLUCION_GARANTIA , "
   g_str_Parame = g_str_Parame & "         NVL((DEVOLUCION_FONDO_CLIENTE),0)                                                                DEVOLUCION_FONDO_CLIENTE , "
   g_str_Parame = g_str_Parame & "         NVL((EXTORNO_COMISION),0)                                                                        EXTORNO_COMISION "
   
   g_str_Parame = g_str_Parame & "    FROM TPR_MAERDE A "
   g_str_Parame = g_str_Parame & "         LEFT JOIN (SELECT B.MAERDE_NUMREF, NVL(SUM(NVL(B.MAERDE_IMPORT,0)),0) IMPORTE_PAGADO "
   g_str_Parame = g_str_Parame & "                      FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                     WHERE (MAERDE_CODIGO = 1 OR MAERDE_CODIGO = 2 OR MAERDE_CODIGO = 4 OR MAERDE_CODIGO = 5 OR MAERDE_CODIGO = 10 OR MAERDE_CODIGO = 14 OR MAERDE_CODIGO = 6 OR MAERDE_CODIGO = 18) "
   g_str_Parame = g_str_Parame & "                     GROUP BY B.MAERDE_NUMREF ) B "
   g_str_Parame = g_str_Parame & "                        ON B.MAERDE_NUMREF = A.MAERDE_NUMREF "
   g_str_Parame = g_str_Parame & "         LEFT JOIN (SELECT C.MAERDE_NUMREF, NVL(SUM(NVL(C.MAERDE_IMPORT,0)),0) DEVOLUCION_GARANTIA "
   g_str_Parame = g_str_Parame & "                      FROM TPR_MAERDE C "
   g_str_Parame = g_str_Parame & "                     WHERE C.MAERDE_CODIGO = 7 "
   g_str_Parame = g_str_Parame & "                     GROUP BY C.MAERDE_NUMREF ) C "
   g_str_Parame = g_str_Parame & "                        ON C.MAERDE_NUMREF = A.MAERDE_NUMREF "
   g_str_Parame = g_str_Parame & "         LEFT JOIN (SELECT D.MAERDE_NUMREF, NVL(SUM(NVL(D.MAERDE_IMPORT,0)),0) EXTORNO_COMISION "
   g_str_Parame = g_str_Parame & "                      FROM TPR_MAERDE D "
   g_str_Parame = g_str_Parame & "                     WHERE D.MAERDE_CODIGO = 15 "
   g_str_Parame = g_str_Parame & "                     GROUP BY D.MAERDE_NUMREF ) D "
   g_str_Parame = g_str_Parame & "                        ON D.MAERDE_NUMREF = A.MAERDE_NUMREF "
   g_str_Parame = g_str_Parame & "         LEFT JOIN (SELECT E.MAERDE_NUMREF, NVL(SUM(NVL(E.MAERDE_IMPORT,0)),0) DEVOLUCION_FONDO_CLIENTE "
   g_str_Parame = g_str_Parame & "                      FROM TPR_MAERDE E "
   g_str_Parame = g_str_Parame & "                     WHERE E.MAERDE_CODIGO = 22 "
   g_str_Parame = g_str_Parame & "                     GROUP BY E.MAERDE_NUMREF ) E "
   g_str_Parame = g_str_Parame & "                        ON E.MAERDE_NUMREF = A.MAERDE_NUMREF "
   g_str_Parame = g_str_Parame & "   WHERE A.MAERDE_NUMREF = '" & CStr(moddat_g_str_DesIte) & "'"
   g_str_Parame = g_str_Parame & "   GROUP BY IMPORTE_PAGADO, DEVOLUCION_GARANTIA, EXTORNO_COMISION, A.MAERDE_NUMREF, DEVOLUCION_FONDO_CLIENTE "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Function
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      If Trim(pnl_CodOpe.Caption) <> "" Then
         r_dbl_ImpSal = CDbl(g_rst_Princi!IMPORTE_RECIBIDO) - CDbl(g_rst_Princi!IMPORTE_PAGADO) + CDbl(g_rst_Princi!DEVOLUCION_GARANTIA) + CDbl(g_rst_Princi!DEVOLUCION_FONDO_CLIENTE) '+ CDbl(g_rst_Princi!EXTORNO_COMISION)
         If CDbl(ipp_ImpDes.Value) <= CDbl(r_dbl_ImpSal) Then
            fs_Validar_Compensacion = True
         End If
      End If
   End If
End Function

Private Sub fs_Buscar_Saldo_Retencion()
Dim r_dbl_ImpSal     As Double
Dim r_dbl_ImpOpe     As Double
   
   pnl_ImpTot.Caption = Format(0, "###,###,###,##0.00") & "  "
   pnl_ImpSal.Caption = Format(0, "###,###,###,##0.00") & "  "
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT NVL(SUM(A.MAERDE_IMPORT), 0) IMPORTE_RETENCION, NVL((IMPORTE_DEVOLUCION),0) IMPORTE_DEVOLUCION, A.MAERDE_TIPMON "
   g_str_Parame = g_str_Parame & "    FROM TPR_MAERDE A "
   g_str_Parame = g_str_Parame & "         LEFT JOIN (SELECT B.MAERDE_NUMREF, NVL(SUM(NVL(B.MAERDE_IMPORT,0)),0) IMPORTE_DEVOLUCION "
   g_str_Parame = g_str_Parame & "                      FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                     WHERE (MAERDE_CODIGO = 14) "
   g_str_Parame = g_str_Parame & "                     GROUP BY B.MAERDE_NUMREF ) B "
   g_str_Parame = g_str_Parame & "                        ON B.MAERDE_NUMREF = A.MAERDE_NUMREF "
   g_str_Parame = g_str_Parame & "   WHERE A.MAERDE_NUMREF = '" & CStr(moddat_g_str_DesIte) & "' " 'moddat_g_str_NumFia
   g_str_Parame = g_str_Parame & "     AND ( A.MAERDE_CODIGO = 10) "
   g_str_Parame = g_str_Parame & "   GROUP BY IMPORTE_DEVOLUCION, A.MAERDE_TIPMON "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      If Trim(pnl_CodOpe.Caption) <> "" Then
         pnl_ImpTot.Caption = Format(CDbl(g_rst_Princi!IMPORTE_RETENCION), "###,###,###,##0.00") & "  "
         r_dbl_ImpSal = CDbl(g_rst_Princi!IMPORTE_RETENCION) + CDbl(g_rst_Princi!IMPORTE_DEVOLUCION)
         pnl_ImpSal.Caption = Format(CDbl(r_dbl_ImpSal), "###,###,###,##0.00") & "  "
         cmb_Moneda.Text = moddat_gf_Consulta_ParDes("204", g_rst_Princi!MAERDE_TIPMON)
        
      End If
   End If
End Sub
Private Sub fs_Buscar_Saldo_Garantia(ByVal p_CodOpe As Integer)
'Garantía Líquida
Dim r_dbl_ImpSal     As Double

   pnl_ImpTot.Caption = Format(0, "###,###,###,##0.00") & "  "
   pnl_ImpSal.Caption = Format(0, "###,###,###,##0.00") & "  "

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT MAEGAR_NUMREF, NVL(SUM(NVL(MAEGAR_MTOGAR_INM,0) + NVL(MAEGAR_MTOGAR_ES1,0) + NVL(MAEGAR_MTOGAR_ES2,0) + NVL(MAEGAR_MTOGAR_DE1,0) + NVL(MAEGAR_MTOGAR_DE2,0)),0) IMPORTE_GARANTIA, "
   g_str_Parame = g_str_Parame & "        (SELECT NVL(SUM(NVL(MAERDE_IMPORT,0)),0)"
   g_str_Parame = g_str_Parame & "           FROM TPR_MAERDE "
   g_str_Parame = g_str_Parame & "          WHERE MAERDE_NUMREF = MAEGAR_NUMREF "
   g_str_Parame = g_str_Parame & "            AND MAERDE_CODIGO = " & p_CodOpe & ") AS IMPORTE_PAGADO "
   g_str_Parame = g_str_Parame & "   FROM TPR_MAEGAR A "
   g_str_Parame = g_str_Parame & "  WHERE MAEGAR_NUMREF = CASE WHEN (SELECT COUNT(B.MAECFI_NUMANT) "
   g_str_Parame = g_str_Parame & "                                     FROM TPR_MAECFI B "
   g_str_Parame = g_str_Parame & "                                    WHERE MAECFI_NUMREF = '" & CStr(moddat_g_str_DesIte) & "' ) > 0 THEN (SELECT B.MAECFI_NUMANT "
   g_str_Parame = g_str_Parame & "                                                                                                            FROM TPR_MAECFI B "
   g_str_Parame = g_str_Parame & "                                                                                                           WHERE MAECFI_NUMREF = '" & CStr(moddat_g_str_DesIte) & "')"
   g_str_Parame = g_str_Parame & "                        ELSE  '" & CStr(moddat_g_str_DesIte) & "' END "
   g_str_Parame = g_str_Parame & "    AND MAEGAR_TIPGAR = 1 "
   'g_str_Parame = g_str_Parame & "    AND MAEGAR_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "  GROUP BY MAEGAR_NUMREF "
  
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
      
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      r_dbl_ImpSal = Format(CDbl(g_rst_Princi!IMPORTE_GARANTIA) - CDbl(g_rst_Princi!IMPORTE_PAGADO), "###,###,###,##0.00")
      If Trim(pnl_CodOpe.Caption) <> "" Then
         pnl_ImpTot.Caption = Format(CDbl(g_rst_Princi!IMPORTE_GARANTIA), "###,###,###,##0.00") & "  "
         pnl_ImpSal.Caption = Format(CDbl(r_dbl_ImpSal), "###,###,###,##0.00") & "  "
      End If
   End If
End Sub

Private Sub fs_Buscar_Saldo_Garantia_Devolucion(ByVal p_CodOpe As Integer)
'Garantía Líquida

Dim r_dbl_ImpSal     As Double
Dim r_dbl_ImpOpe     As Double
Dim r_int_CodGar     As Integer

   pnl_ImpTot.Caption = Format(0, "###,###,###,##0.00") & "  "
   pnl_ImpSal.Caption = Format(0, "###,###,###,##0.00") & "  "
   
   If p_CodOpe = 7 Then          'Devolución Garantía Líquida
      r_int_CodGar = 6           'Retención Garantía Líquida
   ElseIf p_CodOpe = 21 Then     'Devolución Garantía Cliente
      r_int_CodGar = 20          'Pago Garantía Cliente
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "    SELECT NVL(SUM(A.MAERDE_IMPORT), 0) IMPORTE_RECIBIDO, NVL((IMPORTE_DEVUELTO),0) IMPORTE_DEVUELTO "
   g_str_Parame = g_str_Parame & "      FROM TPR_MAERDE A "
   g_str_Parame = g_str_Parame & "           LEFT JOIN (SELECT B.MAERDE_NUMREF, NVL(SUM(NVL(B.MAERDE_IMPORT,0)),0) IMPORTE_DEVUELTO "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                       WHERE (MAERDE_CODIGO = " & p_CodOpe & " ) "
   g_str_Parame = g_str_Parame & "                       GROUP BY B.MAERDE_NUMREF ) B "
   g_str_Parame = g_str_Parame & "                          ON B.MAERDE_NUMREF = A.MAERDE_NUMREF "
   g_str_Parame = g_str_Parame & "   WHERE A.MAERDE_NUMREF = '" & CStr(moddat_g_str_DesIte) & "' " 'moddat_g_str_NumFia
   g_str_Parame = g_str_Parame & "     AND A.MAERDE_CODIGO = " & r_int_CodGar & " "
   g_str_Parame = g_str_Parame & "   GROUP BY IMPORTE_DEVUELTO "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      If Trim(pnl_CodOpe.Caption) <> "" Then
         pnl_ImpTot.Caption = Format(CDbl(g_rst_Princi!IMPORTE_RECIBIDO), "###,###,###,##0.00") & "  "
         r_dbl_ImpSal = CDbl(g_rst_Princi!IMPORTE_RECIBIDO) - CDbl(g_rst_Princi!IMPORTE_DEVUELTO)
         pnl_ImpSal.Caption = Format(CDbl(r_dbl_ImpSal), "###,###,###,##0.00") & "  "
      End If
   End If
End Sub

Function fs_Validar_LiqCFi() As Boolean
fs_Validar_LiqCFi = False
 
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT NVL((SELECT NVL(SUM(NVL(B.MAERDE_IMPORT,0)),0)  "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAERDE B "
   'g_str_Parame = g_str_Parame & "               WHERE MAERDE_CODIGO = 3 "
   g_str_Parame = g_str_Parame & "               WHERE MAERDE_CODIGO IN (3,19) "
   g_str_Parame = g_str_Parame & "                 AND MAERDE_NUMREF = '" & CStr(moddat_g_str_DesIte) & "' " 'moddat_g_str_NumFia
   g_str_Parame = g_str_Parame & "               GROUP BY B.MAERDE_NUMREF), 0) FONDOS, "
   g_str_Parame = g_str_Parame & "         NVL((SELECT NVL(SUM(NVL(B.MAERDE_IMPORT,0)),0) FONDOS "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "               WHERE (MAERDE_CODIGO = 4 Or MAERDE_CODIGO = 5)"
   g_str_Parame = g_str_Parame & "                 AND MAERDE_NUMREF = '" & CStr(moddat_g_str_DesIte) & "' " 'moddat_g_str_NumFia
   g_str_Parame = g_str_Parame & "               GROUP BY B.MAERDE_NUMREF), 0) DESEMBOLSO , "
   g_str_Parame = g_str_Parame & "         NVL((SELECT NVL(MAECFI_COMFIA,0) - NVL(SUM(NVL(B.MAERDE_IMPORT,0)),0) "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI A "
   g_str_Parame = g_str_Parame & "                     LEFT JOIN (SELECT B.MAERDE_NUMREF , B.MAERDE_NUMITE, B.MAERDE_IMPORT "
   g_str_Parame = g_str_Parame & "                                  FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                                 WHERE (MAERDE_CODIGO = 1 OR MAERDE_CODIGO = 2))B "
   g_str_Parame = g_str_Parame & "                                      ON B.MAERDE_NUMREF = A.MAECFI_NUMREF "
   g_str_Parame = g_str_Parame & "               WHERE A.MAECFI_NUMREF = '" & CStr(moddat_g_str_DesIte) & "' " 'moddat_g_str_NumFia
   g_str_Parame = g_str_Parame & "               GROUP BY MAECFI_COMFIA),0) SALDO_COMISION "
   g_str_Parame = g_str_Parame & "    FROM DUAL "
        
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Function
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      If CDbl(g_rst_Princi!FONDOS) < 0 Then
         MsgBox "No se puede Liquidar, la Operación no ha culminado su flujo", vbExclamation, modgen_g_str_NomPlt
      Else
         fs_Validar_LiqCFi = True
      End If
      
'      If CDbl(g_rst_Princi!FONDOS) > 0 Then
'         If CDbl(g_rst_Princi!FONDOS) = CDbl(g_rst_Princi!DESEMBOLSO) And CDbl(g_rst_Princi!SALDO_COMISION) = 0 Then
'            fs_Validar_LiqCFi = True
'         Else
'            If CDbl(g_rst_Princi!SALDO_COMISION) > 0 Then
'               MsgBox "No se puede Liquidar, aún falta pagar Comisión", vbExclamation, modgen_g_str_NomPlt
'            Else
'               MsgBox "No se puede Liquidar, aún no ha sido desembolsado todo el Fondo Recibido", vbExclamation, modgen_g_str_NomPlt
'            End If
'         End If
'      Else
'         MsgBox "No se puede Liquidar, la Operación no ha culminado su flujo", vbExclamation, modgen_g_str_NomPlt
'      End If
   End If
End Function

Function fs_Validar_CanCFi() As Boolean
fs_Validar_CanCFi = False
 
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT COUNT(*) CONTADOR  "
   g_str_Parame = g_str_Parame & "    FROM TPR_MAERDE "
   g_str_Parame = g_str_Parame & "   WHERE MAERDE_NUMREF = '" & CStr(moddat_g_str_DesIte) & "' " 'moddat_g_str_NumFia
        
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Function
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      If CDbl(g_rst_Princi!CONTADOR) = 0 Then
         fs_Validar_CanCFi = True
      End If
   End If
End Function

Private Sub fs_Limpia()
   cmb_Moneda.ListIndex = -1
   cmb_TipOper.ListIndex = -1
   txt_Refer.Text = ""
   ipp_PorDes.Value = "0.00%"
   pnl_ImpTot.Caption = Format(0, "###,###,###,##0.00") & "  "
   pnl_ImpSal.Caption = Format(0, "###,###,###,##0.00") & "  "
   ipp_ImpDes.Value = Format(0, "###,###,###,##0.00")
'   ipp_FecOpe.Text = Format(date, "dd/mm/yyyy")
   ipp_FecOpe.DateMax = modctb_str_FecFin
   ipp_FecOpe.DateMin = modctb_str_FecIni
   If (Format(moddat_g_str_FecSis, "yyyymmdd") <= Format(modctb_str_FecFin, "yyyymmdd")) Then
       ipp_FecOpe.Text = moddat_g_str_FecSis
   Else
       ipp_FecOpe.Text = modctb_str_FecFin
   End If
   
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   pnl_RScPrv.Caption = ""
   cmb_CodBan.ListIndex = -1
   cmb_CtaBan.Clear
   pnl_CCIBan.Caption = ""
   txt_NumMov.Text = ""
   pnl_CodOpe.Caption = ""
   cmb_ForPag.ListIndex = -1
'   l_str_FilSel = ""
   l_int_NumIte = 0
End Sub
Private Sub fs_Limpia_DatBan()
   
   cmb_TipDoc.ListIndex = -1
   pnl_CCIBan.Caption = Empty
   txt_NumDoc.Text = Empty
   pnl_RScPrv.Caption = Empty
   cmb_CodBan.ListIndex = -1
   cmb_CtaBan.ListIndex = -1
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
Dim r_int_NumFil As Integer

   cmb_TipOper.Enabled = p_Activa
   txt_Refer.Enabled = p_Activa
   cmb_Moneda.Enabled = p_Activa
   ipp_FecOpe.Enabled = p_Activa
   pnl_ImpSal.Enabled = p_Activa
   pnl_ImpTot.Enabled = p_Activa
   ipp_ImpDes.Enabled = p_Activa
   ipp_PorDes.Enabled = p_Activa
   txt_NumMov.Enabled = p_Activa
   cmb_TipDoc.Enabled = p_Activa
   txt_NumDoc.Enabled = p_Activa
   pnl_RScPrv.Enabled = p_Activa
   cmb_CodBan.Enabled = p_Activa
   cmb_CtaBan.Enabled = p_Activa
   cmb_ForPag.Enabled = p_Activa
   pnl_CCIBan.Enabled = p_Activa
   
   cmd_Agrega.Enabled = Not p_Activa
   If Me.grd_Listad.Row < 0 Then
      cmd_Editar.Enabled = p_Activa
      cmd_Borrar.Enabled = p_Activa
      cmd_ExpExc.Enabled = p_Activa
      cmd_ExpLiq.Enabled = p_Activa
   Else
      cmd_Editar.Enabled = Not p_Activa
      cmd_Borrar.Enabled = Not p_Activa
      cmd_ExpExc.Enabled = Not p_Activa
      cmd_ExpLiq.Enabled = Not p_Activa
   End If
   cmd_Grabar.Enabled = p_Activa
   cmd_Cancel.Enabled = p_Activa
   
   'solo cuando es desembolso
   cmd_ExpLiq.Enabled = p_Activa
   For r_int_NumFil = 0 To grd_Listad.Rows - 1
      If CInt(grd_Listad.TextMatrix(r_int_NumFil, 9)) = 4 Or CInt(grd_Listad.TextMatrix(r_int_NumFil, 9)) = 17 Then
         cmd_ExpLiq.Enabled = Not p_Activa
         Exit For
      End If
   Next
End Sub

Private Sub fs_ActOper(ByVal p_Activa As Integer)
   If CInt(pnl_CodOpe.Caption) = 8 Or Int(pnl_CodOpe.Caption) = 9 Or Int(pnl_CodOpe.Caption) = 14 Then
      If Int(pnl_CodOpe.Caption) = 9 Then
         ipp_FecOpe.Enabled = p_Activa
      Else
         ipp_FecOpe.Enabled = Not p_Activa
      End If
      txt_Refer.Enabled = Not p_Activa
      
   Else
      ipp_FecOpe.Enabled = p_Activa
      txt_Refer.Enabled = p_Activa
   End If
   cmb_Moneda.Enabled = p_Activa
   ipp_PorDes.Enabled = p_Activa
   ipp_ImpDes.Enabled = p_Activa
End Sub

Private Sub grd_Listad_DblClick()
   'If moddat_g_str_DesObs = "VIGENTE" Then
      cmd_Editar_Click
   'End If
End Sub


Private Sub ipp_FecOpe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 8 Then
         Call gs_SetFocus(txt_Refer)
      Else
         If cmb_TipOper.ListIndex <> -1 Then
            If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 4 And cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 5 And cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 7 And cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 14 And cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 18 And cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 21 Then
               Call gs_SetFocus(ipp_ImpDes)
            Else
               Call gs_SetFocus(cmb_ForPag)
            End If
         End If
      End If
   Else
      KeyAscii = 0
   End If
End Sub

'Private Sub ipp_FecOpe_LostFocus()
'   If (Format(ipp_FecOpe.Text, "yyyymmdd") < Format(moddat_g_str_FecIni, "yyyymmdd") Or _
'       Format(ipp_FecOpe.Text, "yyyymmdd") > Format(moddat_g_str_FecFin, "yyyymmdd")) Then
'       MsgBox "El movimiento que intenta registrar está en un periodo errado.", vbExclamation, modgen_g_str_NomPlt
'       Call gs_SetFocus(ipp_FecOpe)
'   End If
'End Sub

Private Sub ipp_ImpDes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Moneda)
   End If
End Sub

Private Sub ipp_PorDes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ImpDes)
   End If
End Sub

Private Sub ipp_PorDes_LostFocus()
   Call fs_Calcular
End Sub



Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      pnl_RScPrv.Caption = ""

      If cmb_TipDoc.ListIndex <> -1 Then
          pnl_RScPrv.Caption = fs_BuscarProv(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex), Trim(txt_NumDoc.Text))
          Call fs_Buscar_Banco(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex), txt_NumDoc.Text)
      End If
      
      If cmb_ForPag.ListIndex <> -1 Then
         If cmb_ForPag.ItemData(cmb_ForPag.ListIndex) = 1 Then
            Call gs_SetFocus(txt_Refer)
         Else
             If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 4 And cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 14 And cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 18 Then
               Call gs_SetFocus(cmb_CodBan)
            Else
               If cmb_ForPag.ItemData(cmb_ForPag.ListIndex) = 2 Then
                  Call gs_SetFocus(cmb_CodBan)
               Else
                  Call gs_SetFocus(txt_NumMov)
               End If
            End If
         End If
      End If
    Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
    End If
End Sub

Private Function fs_BuscarProv(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String) As String

    'pnl_RScPrv.Caption = ""
    fs_BuscarProv = ""
    
    g_str_Parame = ""
    g_str_Parame = g_str_Parame & " SELECT MAEPRV_TIPDOC, MAEPRV_NUMDOC, MAEPRV_RAZSOC "
    g_str_Parame = g_str_Parame & "   FROM CNTBL_MAEPRV A  "
    g_str_Parame = g_str_Parame & "  WHERE MAEPRV_SITUAC = 1  "
    
    If p_TipDoc > 0 Then
         g_str_Parame = g_str_Parame & "   AND MAEPRV_TIPDOC = " & p_TipDoc & "  "
         g_str_Parame = g_str_Parame & "   AND MAEPRV_NUMDOC = '" & p_NumDoc & "' "
    End If
    'If Len(Trim(p_NumDoc)) > 0 Then
    '     g_str_Parame = g_str_Parame & "   AND MAEPRV_NUMDOC = '" & p_NumDoc & "' "
    'End If
    
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Function
    End If
    
    If g_rst_Princi.BOF And g_rst_Princi.EOF Then
       g_rst_Princi.Close
       Set g_rst_Princi = Nothing
       Screen.MousePointer = 0
       Exit Function
    End If
    
    g_rst_Princi.MoveFirst
    If Not g_rst_Princi.EOF Then
       'pnl_RScPrv.Caption = Trim(g_rst_Princi!MAEPRV_RAZSOC & "")
       fs_BuscarProv = Trim(g_rst_Princi!MAEPRV_RAZSOC & "")
    End If
    
    g_rst_Princi.Close
    Set g_rst_Princi = Nothing
    Screen.MousePointer = 0
End Function

Private Sub fs_Buscar_Banco(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String)
    
    cmb_CodBan.ListIndex = -1
    cmb_CodBan.Clear
    cmb_CtaBan.Clear
    pnl_CCIBan.Caption = ""
    
    g_str_Parame = ""
    g_str_Parame = g_str_Parame & "  SELECT NVL(MAEPRV_CODBNC_MN1,0) AS CODBNC_MN1, TRIM(B.PARDES_DESCRI) AS BANCO_MN1,"
    g_str_Parame = g_str_Parame & "         NVL(MAEPRV_CODBNC_MN2,0) AS CODBNC_MN2, TRIM(C.PARDES_DESCRI) AS BANCO_MN2,"
    g_str_Parame = g_str_Parame & "         NVL(MAEPRV_CODBNC_MN3,0) AS CODBNC_MN3, TRIM(D.PARDES_DESCRI) AS BANCO_MN3,"
    g_str_Parame = g_str_Parame & "         NVL(MAEPRV_CODBNC_DL1,0) AS CODBNC_DL1, TRIM(E.PARDES_DESCRI) AS BANCO_DL1,"
    g_str_Parame = g_str_Parame & "         NVL(MAEPRV_CODBNC_DL2,0) AS CODBNC_DL2, TRIM(F.PARDES_DESCRI) AS BANCO_DL2,"
    g_str_Parame = g_str_Parame & "         NVL(MAEPRV_CODBNC_DL3,0) AS CODBNC_DL3, TRIM(G.PARDES_DESCRI) AS BANCO_DL3"
    g_str_Parame = g_str_Parame & "    FROM CNTBL_MAEPRV A"
    g_str_Parame = g_str_Parame & "         LEFT JOIN MNT_PARDES B ON A.MAEPRV_CODBNC_MN1 = B.PARDES_CODITE AND B.PARDES_CODGRP = 122"
    g_str_Parame = g_str_Parame & "         LEFT JOIN MNT_PARDES C ON A.MAEPRV_CODBNC_MN2 = C.PARDES_CODITE AND C.PARDES_CODGRP = 122"
    g_str_Parame = g_str_Parame & "         LEFT JOIN MNT_PARDES D ON A.MAEPRV_CODBNC_MN3 = D.PARDES_CODITE AND D.PARDES_CODGRP = 122"
    g_str_Parame = g_str_Parame & "         LEFT JOIN MNT_PARDES E ON A.MAEPRV_CODBNC_DL1 = E.PARDES_CODITE AND E.PARDES_CODGRP = 122"
    g_str_Parame = g_str_Parame & "         LEFT JOIN MNT_PARDES F ON A.MAEPRV_CODBNC_DL2 = F.PARDES_CODITE AND F.PARDES_CODGRP = 122"
    g_str_Parame = g_str_Parame & "         LEFT JOIN MNT_PARDES G ON A.MAEPRV_CODBNC_DL3 = G.PARDES_CODITE AND G.PARDES_CODGRP = 122"
    g_str_Parame = g_str_Parame & "   WHERE MAEPRV_SITUAC = 1  "
    g_str_Parame = g_str_Parame & "     AND MAEPRV_TIPDOC = " & p_TipDoc & "  "
    g_str_Parame = g_str_Parame & "     AND MAEPRV_NUMDOC = '" & p_NumDoc & "' "
    g_str_Parame = g_str_Parame & "  ORDER BY MAEPRV_TIPDOC, MAEPRV_RAZSOC ASC  "
    
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
    End If
    
    If g_rst_Princi.BOF And g_rst_Princi.EOF Then
       g_rst_Princi.Close
       Set g_rst_Princi = Nothing
       Screen.MousePointer = 0
       Exit Sub
    End If
    
    g_rst_Princi.MoveFirst
    If Not g_rst_Princi.EOF Then
      If CLng(g_rst_Princi!CODBNC_MN1) <> 0 Then
         cmb_CodBan.AddItem Trim$(g_rst_Princi!BANCO_MN1)
         cmb_CodBan.ItemData(cmb_CodBan.NewIndex) = CLng(g_rst_Princi!CODBNC_MN1)
      End If
      If CLng(g_rst_Princi!CODBNC_MN2) <> 0 And CLng(g_rst_Princi!CODBNC_MN1) <> CLng(g_rst_Princi!CODBNC_MN2) Then
         cmb_CodBan.AddItem Trim$(g_rst_Princi!BANCO_MN2)
         cmb_CodBan.ItemData(cmb_CodBan.NewIndex) = CLng(g_rst_Princi!CODBNC_MN2)
      End If
      If CLng(g_rst_Princi!CODBNC_MN3) <> 0 And CLng(g_rst_Princi!CODBNC_MN1) <> CLng(g_rst_Princi!CODBNC_MN2) And CLng(g_rst_Princi!CODBNC_MN2) <> CLng(g_rst_Princi!CODBNC_MN3) Then
         cmb_CodBan.AddItem Trim$(g_rst_Princi!BANCO_MN3)
         cmb_CodBan.ItemData(cmb_CodBan.NewIndex) = CLng(g_rst_Princi!CODBNC_MN3)
      End If
      If CLng(g_rst_Princi!CODBNC_DL1) <> 0 And CLng(g_rst_Princi!CODBNC_DL1) <> CLng(g_rst_Princi!CODBNC_MN1) And CLng(g_rst_Princi!CODBNC_DL1) <> CLng(g_rst_Princi!CODBNC_MN2) And CLng(g_rst_Princi!CODBNC_DL1) <> CLng(g_rst_Princi!CODBNC_MN3) Then
         cmb_CodBan.AddItem Trim$(g_rst_Princi!BANCO_DL1)
         cmb_CodBan.ItemData(cmb_CodBan.NewIndex) = CLng(g_rst_Princi!CODBNC_DL1)
      End If
      If CLng(g_rst_Princi!CODBNC_DL2) <> 0 And CLng(g_rst_Princi!CODBNC_DL2) <> CLng(g_rst_Princi!CODBNC_DL1) Then
         cmb_CodBan.AddItem Trim$(g_rst_Princi!BANCO_DL2)
         cmb_CodBan.ItemData(cmb_CodBan.NewIndex) = CLng(g_rst_Princi!CODBNC_DL2)
      End If
      If CLng(g_rst_Princi!CODBNC_DL3) <> 0 And CLng(g_rst_Princi!CODBNC_DL3) <> CLng(g_rst_Princi!CODBNC_DL2) Then
         cmb_CodBan.AddItem Trim$(g_rst_Princi!BANCO_DL3)
         cmb_CodBan.ItemData(cmb_CodBan.NewIndex) = CLng(g_rst_Princi!CODBNC_DL3)
      End If
    End If
    
    g_rst_Princi.Close
    Set g_rst_Princi = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub txt_NumDoc_LostFocus()
   txt_NumDoc_KeyPress (13)
End Sub

Private Sub txt_NumMov_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Refer)
   End If
End Sub

Private Sub Txt_Refer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
'      If cmb_TipOper.ListIndex <> -1 Then
'         If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 4 Or cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 7 Then
'            Call gs_SetFocus(cmd_Grabar)
'         ElseIf cmb_TipOper.ItemData(cmb_TipOper.ListIndex) = 5 Then
'            Call gs_SetFocus(cmb_TipDoc)
'         Else
            Call gs_SetFocus(cmd_Grabar)
'         End If
'      End If
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-( )%$.:@_?¿º#= ")
   End If
End Sub

Private Sub fs_Calcular()
Dim r_dbl_PorDes  As Double
   
    If CDbl(Replace(ipp_PorDes.Value, "%", "")) < 0 Then
        ipp_ImpDes.Caption = "0.00 "
    Else
        r_dbl_PorDes = CDbl(Replace(ipp_PorDes.Value, "%", "")) / 100
        If cmb_TipOper.ListIndex <> -1 Then
            If cmb_TipOper.ItemData(cmb_TipOper.ListIndex) <> 14 Then
                ipp_ImpDes.Value = Format(CDbl(pnl_ImpTot.Caption) * r_dbl_PorDes, "##,###,##0.00") & " "
            End If
        Else
            ipp_ImpDes.Value = Format(CDbl(pnl_ImpTot.Caption) * r_dbl_PorDes, "##,###,##0.00") & " "
        End If
    End If
End Sub
Private Function fs_Validar_Renovacion(ByVal p_NumRef As String) As Boolean
   
   fs_Validar_Renovacion = False
  
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT COUNT(*) CONTADOR "
   g_str_Parame = g_str_Parame & "   FROM TPR_MAECFI A "
   g_str_Parame = g_str_Parame & "  WHERE MAECFI_REFORI = (SELECT MAECFI_REFORI"
   g_str_Parame = g_str_Parame & "                           FROM TPR_MAECFI"
   g_str_Parame = g_str_Parame & "                          WHERE MAECFI_NUMREF = '" & CStr(p_NumRef) & "' "
   g_str_Parame = g_str_Parame & "                            AND MAECFI_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " "
   g_str_Parame = g_str_Parame & "                            AND MAECFI_NUMDOC = '" & CStr(moddat_g_str_NumDoc) & "' ) "
   g_str_Parame = g_str_Parame & "    AND MAECFI_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " "
   g_str_Parame = g_str_Parame & "    AND MAECFI_NUMDOC = '" & CStr(moddat_g_str_NumDoc) & "' "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Function
   End If
       
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     Exit Function
   End If
   
   g_rst_Princi.MoveFirst
   
   If CInt(g_rst_Princi!CONTADOR) > 1 Then
      fs_Validar_Renovacion = True
   End If
End Function
