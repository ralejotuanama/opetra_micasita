VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_EmpPer_07 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13485
   Icon            =   "OpeTra_frm_405.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9780
   ScaleWidth      =   13485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   9825
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   13515
      _Version        =   65536
      _ExtentX        =   23839
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
      Begin Threed.SSPanel SSPanel5 
         Height          =   645
         Left            =   60
         TabIndex        =   20
         Top             =   770
         Width           =   13380
         _Version        =   65536
         _ExtentX        =   23601
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
            Left            =   12780
            Picture         =   "OpeTra_frm_405.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   0
            Picture         =   "OpeTra_frm_405.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Imprimir Listado"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   525
         Left            =   60
         TabIndex        =   21
         Top             =   1440
         Width           =   13380
         _Version        =   65536
         _ExtentX        =   23601
         _ExtentY        =   926
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
         Begin Threed.SSPanel pnl_EmpPer 
            Height          =   405
            Left            =   990
            TabIndex        =   22
            Top             =   60
            Width           =   12210
            _Version        =   65536
            _ExtentX        =   21537
            _ExtentY        =   714
            _StockProps     =   15
            Caption         =   "SSPanel3"
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
         Begin VB.Label lbl_NomEmp 
            Caption         =   "Notaria:"
            Height          =   195
            Left            =   210
            TabIndex        =   23
            Top             =   150
            Width           =   825
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   960
         Left            =   60
         TabIndex        =   24
         Top             =   2000
         Width           =   13380
         _Version        =   65536
         _ExtentX        =   23601
         _ExtentY        =   1693
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
         Begin VB.TextBox txt_DirEle2 
            Height          =   315
            Left            =   5340
            MaxLength       =   50
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   180
            Width           =   3500
         End
         Begin VB.TextBox txt_DirEle1 
            Height          =   315
            Left            =   990
            MaxLength       =   50
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   180
            Width           =   3500
         End
         Begin VB.TextBox txt_DirEle3 
            Height          =   315
            Left            =   9690
            MaxLength       =   50
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   180
            Width           =   3500
         End
         Begin VB.TextBox txt_DirEle4 
            Height          =   315
            Left            =   990
            MaxLength       =   50
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   510
            Width           =   3500
         End
         Begin VB.TextBox txt_DirEle5 
            Height          =   315
            Left            =   5340
            MaxLength       =   50
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   510
            Width           =   3500
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Correo 2:"
            Height          =   195
            Left            =   4620
            TabIndex        =   29
            Top             =   240
            Width           =   645
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Correo 1:"
            Height          =   195
            Left            =   210
            TabIndex        =   28
            Top             =   240
            Width           =   645
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Correo 3:"
            Height          =   195
            Left            =   8970
            TabIndex        =   27
            Top             =   240
            Width           =   645
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Correo 4:"
            Height          =   195
            Left            =   210
            TabIndex        =   26
            Top             =   570
            Width           =   645
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Correo 5:"
            Height          =   195
            Left            =   4620
            TabIndex        =   25
            Top             =   570
            Width           =   645
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   675
         Left            =   60
         TabIndex        =   30
         Top             =   60
         Width           =   13380
         _Version        =   65536
         _ExtentX        =   23601
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
            Height          =   270
            Left            =   630
            TabIndex        =   31
            Top             =   210
            Width           =   3105
            _Version        =   65536
            _ExtentX        =   5477
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Parametros de Gastos de Cierre"
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
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   12900
            Top             =   30
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
         Begin VB.Image Image1 
            Height          =   480
            Left            =   90
            Picture         =   "OpeTra_frm_405.frx":0890
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   6750
         Left            =   60
         TabIndex        =   32
         Top             =   2985
         Width           =   13380
         _Version        =   65536
         _ExtentX        =   23601
         _ExtentY        =   11906
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
            Height          =   3015
            Left            =   60
            TabIndex        =   33
            Top             =   60
            Width           =   13260
            _ExtentX        =   23389
            _ExtentY        =   5318
            _Version        =   393216
            Rows            =   1
            Cols            =   19
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel9 
            Height          =   2895
            Left            =   75
            TabIndex        =   34
            Top             =   3795
            Width           =   13245
            _Version        =   65536
            _ExtentX        =   23363
            _ExtentY        =   5106
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
            Begin VB.CommandButton cmd_Cancelar 
               Height          =   585
               Left            =   12450
               Picture         =   "OpeTra_frm_405.frx":0B9A
               Style           =   1  'Graphical
               TabIndex        =   16
               ToolTipText     =   "Cancelar"
               Top             =   2160
               Width           =   585
            End
            Begin VB.CommandButton cmd_Aceptar 
               Height          =   585
               Left            =   11850
               Picture         =   "OpeTra_frm_405.frx":0FDC
               Style           =   1  'Graphical
               TabIndex        =   15
               Tag             =   "0"
               ToolTipText     =   "Agregar Cuenta"
               Top             =   2160
               Width           =   585
            End
            Begin VB.Frame Frame3 
               Caption         =   "Datos"
               Height          =   855
               Left            =   6780
               TabIndex        =   44
               Top             =   1950
               Width           =   4155
               Begin EditLib.fpDoubleSingle txt_GasNot_Mto 
                  Height          =   315
                  Left            =   2760
                  TabIndex        =   14
                  Top             =   300
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
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Monto Gastos Notariales:"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   45
                  Top             =   360
                  Width           =   1785
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "Inscripción Garantía"
               Height          =   855
               Left            =   6780
               TabIndex        =   40
               Top             =   930
               Width           =   6285
               Begin EditLib.fpDoubleSingle txt_RegGar_Tas000 
                  Height          =   315
                  Left            =   825
                  TabIndex        =   11
                  Top             =   350
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
                  Text            =   "0.0000"
                  DecimalPlaces   =   4
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
               Begin EditLib.fpDoubleSingle txt_RegGar_Tas001 
                  Height          =   315
                  Left            =   2760
                  TabIndex        =   12
                  Top             =   350
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
                  Text            =   "0.0000"
                  DecimalPlaces   =   4
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
               Begin EditLib.fpDoubleSingle txt_RegGar_Tas002 
                  Height          =   315
                  Left            =   4920
                  TabIndex        =   13
                  Top             =   350
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
                  Text            =   "0.0000"
                  DecimalPlaces   =   4
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
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  Caption         =   "Tasa:"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   43
                  Top             =   410
                  Width           =   405
               End
               Begin VB.Label Label10 
                  AutoSize        =   -1  'True
                  Caption         =   "Factor:"
                  Height          =   195
                  Left            =   2160
                  TabIndex        =   42
                  Top             =   410
                  Width           =   495
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  Caption         =   "Por Ficha:"
                  Height          =   195
                  Left            =   4050
                  TabIndex        =   41
                  Top             =   410
                  Width           =   720
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "Minuta Compra Venta"
               Height          =   1875
               Left            =   150
               TabIndex        =   39
               Top             =   930
               Width           =   6555
               Begin VB.Frame Frame5 
                  Caption         =   "Estacionamiento"
                  Height          =   735
                  Left            =   120
                  TabIndex        =   51
                  Top             =   1035
                  Width           =   6285
                  Begin EditLib.fpDoubleSingle txt_RegMin_TasEst 
                     Height          =   315
                     Left            =   820
                     TabIndex        =   54
                     Top             =   290
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
                     Text            =   "0.0000"
                     DecimalPlaces   =   4
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
                  Begin EditLib.fpDoubleSingle txt_RegMin_FacEst 
                     Height          =   315
                     Left            =   2760
                     TabIndex        =   58
                     Top             =   240
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
                     Text            =   "0.0000"
                     DecimalPlaces   =   4
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
                  Begin EditLib.fpDoubleSingle txt_RegMin_FicEst 
                     Height          =   315
                     Left            =   4920
                     TabIndex        =   59
                     Top             =   285
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
                     Text            =   "0.0000"
                     DecimalPlaces   =   4
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
                  Begin VB.Label Label6 
                     AutoSize        =   -1  'True
                     Caption         =   "Tasa :"
                     Height          =   195
                     Left            =   240
                     TabIndex        =   55
                     Top             =   300
                     Width           =   450
                  End
                  Begin VB.Label Label12 
                     AutoSize        =   -1  'True
                     Caption         =   "Factor:"
                     Height          =   195
                     Left            =   2160
                     TabIndex        =   53
                     Top             =   300
                     Width           =   495
                  End
                  Begin VB.Label Label8 
                     AutoSize        =   -1  'True
                     Caption         =   "Por Ficha:"
                     Height          =   195
                     Left            =   4050
                     TabIndex        =   52
                     Top             =   345
                     Width           =   720
                  End
               End
               Begin VB.Frame Frame4 
                  Caption         =   "Inmueble"
                  Height          =   735
                  Left            =   120
                  TabIndex        =   46
                  Top             =   240
                  Width           =   6285
                  Begin EditLib.fpDoubleSingle txt_RegMin_TasInm 
                     Height          =   315
                     Left            =   820
                     TabIndex        =   49
                     Top             =   290
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
                     Text            =   "0.0000"
                     DecimalPlaces   =   4
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
                  Begin EditLib.fpDoubleSingle txt_RegMin_FacInm 
                     Height          =   315
                     Left            =   2760
                     TabIndex        =   56
                     Top             =   285
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
                     Text            =   "0.0000"
                     DecimalPlaces   =   4
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
                  Begin EditLib.fpDoubleSingle txt_RegMin_FicInm 
                     Height          =   315
                     Left            =   4920
                     TabIndex        =   57
                     Top             =   285
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
                     Text            =   "0.0000"
                     DecimalPlaces   =   4
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
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     Caption         =   "Tasa :"
                     Height          =   195
                     Left            =   240
                     TabIndex        =   50
                     Top             =   300
                     Width           =   450
                  End
                  Begin VB.Label Label7 
                     AutoSize        =   -1  'True
                     Caption         =   "Por Ficha:"
                     Height          =   195
                     Left            =   4050
                     TabIndex        =   48
                     Top             =   345
                     Width           =   720
                  End
                  Begin VB.Label Label4 
                     AutoSize        =   -1  'True
                     Caption         =   "Factor:"
                     Height          =   195
                     Left            =   2160
                     TabIndex        =   47
                     Top             =   345
                     Width           =   495
                  End
               End
            End
            Begin VB.ComboBox cmb_Moneda 
               Height          =   315
               Left            =   990
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   510
               Width           =   2400
            End
            Begin VB.ComboBox cmb_Producto 
               Height          =   315
               Left            =   990
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   180
               Width           =   5070
            End
            Begin VB.ComboBox cmb_Proyecto 
               Height          =   315
               Left            =   7110
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   180
               Width           =   6105
            End
            Begin VB.Label Label84 
               AutoSize        =   -1  'True
               Caption         =   "Moneda:"
               Height          =   195
               Left            =   180
               TabIndex        =   38
               Top             =   555
               Width           =   630
            End
            Begin VB.Label Label81 
               AutoSize        =   -1  'True
               Caption         =   "Producto:"
               Height          =   195
               Left            =   180
               TabIndex        =   37
               Top             =   240
               Width           =   690
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Proyecto:"
               Height          =   195
               Left            =   6300
               TabIndex        =   36
               Top             =   240
               Width           =   675
            End
         End
         Begin Threed.SSPanel pnl_Boton_Cta 
            Height          =   675
            Left            =   60
            TabIndex        =   35
            Top             =   3090
            Width           =   13275
            _Version        =   65536
            _ExtentX        =   23407
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
            BorderWidth     =   1
            BevelOuter      =   0
            BevelInner      =   1
            Begin VB.CommandButton cmd_Editar 
               Height          =   570
               Left            =   12645
               Picture         =   "OpeTra_frm_405.frx":12E6
               Style           =   1  'Graphical
               TabIndex        =   7
               ToolTipText     =   "Editar Cuenta"
               Top             =   60
               Width           =   585
            End
            Begin VB.CommandButton cmd_Borrar 
               Height          =   570
               Left            =   12045
               Picture         =   "OpeTra_frm_405.frx":15F0
               Style           =   1  'Graphical
               TabIndex        =   6
               ToolTipText     =   "Eliminar Cuenta"
               Top             =   60
               Width           =   585
            End
            Begin VB.CommandButton cmd_Nuevo 
               Height          =   570
               Left            =   11445
               Picture         =   "OpeTra_frm_405.frx":18FA
               Style           =   1  'Graphical
               TabIndex        =   5
               ToolTipText     =   "Nueva Cuenta"
               Top             =   60
               Width           =   585
            End
         End
      End
   End
End
Attribute VB_Name = "frm_EmpPer_07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_str_Codigo  As String

Private Sub cmb_Moneda_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_RegMin_TasInm)
   End If
End Sub

Private Sub cmb_Producto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Proyecto)
   End If
End Sub

Private Sub cmb_Proyecto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Moneda)
   End If
End Sub

Private Sub cmd_Aceptar_Click()
   If cmd_Aceptar.Tag = "0" Or cmd_Aceptar.Tag = "" Then
      Exit Sub
   End If
   
   If cmb_Producto.ListIndex = -1 Then
      MsgBox "Debe seleccionar un producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Producto)
      Exit Sub
   End If
   If cmb_Proyecto.ListIndex = -1 Then
      MsgBox "Debe seleccionar un proyecto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Proyecto)
      Exit Sub
   End If
   If cmb_Moneda.ListIndex = -1 Then
      MsgBox "Debe seleccionar un tipo de moneda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Moneda)
      Exit Sub
   End If
   
   If CDbl(txt_RegMin_TasInm.Text) = 0 Then
      MsgBox "Debe de ingresar la tasa del inmueble, en el grupo minuta compra venta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_RegMin_TasInm)
      Exit Sub
   End If
   If CDbl(txt_RegMin_FacInm.Text) = 0 Then
      MsgBox "Debe de ingresar factor del inmueble, en el grupo minuta compra venta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_RegMin_FacInm)
      Exit Sub
   End If
   If CDbl(txt_RegMin_FicInm.Text) = 0 Then
      MsgBox "Debe de ingresar importe por ficha del inmueble, en el grupo minuta compra venta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_RegMin_FicInm)
      Exit Sub
   End If
   If CDbl(txt_RegMin_TasEst.Text) = 0 Then
      MsgBox "Debe de ingresar la tasa de estacionamiento, en el grupo minuta compra venta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_RegMin_TasEst)
      Exit Sub
   End If
   If CDbl(txt_RegMin_FacEst.Text) = 0 Then
      MsgBox "Debe de ingresar factor del estacionamiento, en el grupo minuta compra venta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_RegMin_FacEst)
      Exit Sub
   End If
   If CDbl(txt_RegMin_FicEst.Text) = 0 Then
      MsgBox "Debe de ingresar importe por ficha del estacionamiento, en el grupo minuta compra venta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_RegMin_FicEst)
      Exit Sub
   End If
   
   If CDbl(txt_RegGar_Tas000.Text) = 0 Then
      MsgBox "Debe de ingresar una tasa, en el grupo inscripción de garantía.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_RegGar_Tas000)
      Exit Sub
   End If
   If CDbl(txt_RegGar_Tas001.Text) = 0 Then
      MsgBox "Debe de ingresar una tasa adicional 1, en el grupo inscripción de garantía.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_RegGar_Tas001)
      Exit Sub
   End If
   If CDbl(txt_RegGar_Tas002.Text) = 0 Then
      MsgBox "Debe de ingresar una tasa adicional 2, en el grupo inscripción de garantía.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_RegGar_Tas002)
      Exit Sub
   End If
   If CDbl(txt_GasNot_Mto.Text) = 0 Then
      MsgBox "Debe de ingresar el monto gastos notariales, en el grupo datos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_GasNot_Mto)
      Exit Sub
   End If
   '--------------------------
   
    If cmd_Aceptar.Tag = 1 Then
       'Insertar
       If MsgBox("¿Está seguro de insertar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
          Exit Sub
       End If
       
       If fs_Valida_Insercion(cmb_Producto.ItemData(cmb_Producto.ListIndex), cmb_Proyecto.ItemData(cmb_Proyecto.ListIndex)) Then
          grd_Listad.Rows = grd_Listad.Rows + 1
          grd_Listad.Row = grd_Listad.Rows - 1
          
          grd_Listad.Col = 0
          grd_Listad.Text = ""
          grd_Listad.Col = 1
          grd_Listad.Text = cmb_Producto.ItemData(cmb_Producto.ListIndex)
          grd_Listad.Col = 2
          grd_Listad.Text = Trim(cmb_Producto.Text)
          grd_Listad.Col = 3
          grd_Listad.Text = cmb_Proyecto.ItemData(cmb_Proyecto.ListIndex)
          grd_Listad.Col = 4
          grd_Listad.Text = Trim(cmb_Proyecto.Text)
          grd_Listad.Col = 5
          grd_Listad.Text = cmb_Moneda.ItemData(cmb_Moneda.ListIndex)
          grd_Listad.Col = 6
          grd_Listad.Text = Trim(cmb_Moneda.Text)
          grd_Listad.Col = 7
          grd_Listad.Text = txt_GasNot_Mto.Text    'GASPAR_GASNOT_MTO
          grd_Listad.Col = 8
          grd_Listad.Text = txt_RegMin_TasInm.Text 'GASPAR_REGMIN_TASINM
          grd_Listad.Col = 9
          grd_Listad.Text = txt_RegMin_FacInm.Text 'GASPAR_REGMIN_FACINM
          grd_Listad.Col = 10
          grd_Listad.Text = txt_RegMin_FicInm.Text 'GASPAR_REGMIN_FICINM
          grd_Listad.Col = 11
          grd_Listad.Text = txt_RegMin_TasEst.Text 'GASPAR_REGMIN_TASEST
          grd_Listad.Col = 12
          grd_Listad.Text = txt_RegMin_FacEst.Text 'GASPAR_REGMIN_FACEST
          grd_Listad.Col = 13
          grd_Listad.Text = txt_RegMin_FicEst.Text 'GASPAR_REGMIN_FICEST
          grd_Listad.Col = 14
          grd_Listad.Text = txt_RegGar_Tas000.Text 'GASPAR_REGGAR_TAS000
          grd_Listad.Col = 15
          grd_Listad.Text = txt_RegGar_Tas001.Text 'GASPAR_REGGAR_TAS001
          grd_Listad.Col = 16
          grd_Listad.Text = txt_RegGar_Tas002.Text 'GASPAR_REGGAR_TAS002
          grd_Listad.Col = 17
          grd_Listad.Text = 1 'estado
          grd_Listad.Col = 18
          grd_Listad.Text = 1 'input
       Else
          MsgBox "Los parámetros ya se encuentran ingresados.", vbExclamation, modgen_g_str_NomPlt
          Exit Sub
       End If
       
    ElseIf cmd_Aceptar.Tag = 2 Then
      'Actualizar
       If MsgBox("¿Está seguro de actualizar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
          Exit Sub
       End If
       grd_Listad.TextMatrix(grd_Listad.Row, 1) = cmb_Producto.ItemData(cmb_Producto.ListIndex) 'l_arr_Produc(cmb_Producto.ListIndex + 1).Genera_Codigo
       grd_Listad.TextMatrix(grd_Listad.Row, 2) = Trim(cmb_Producto.Text)
       grd_Listad.TextMatrix(grd_Listad.Row, 3) = cmb_Proyecto.ItemData(cmb_Proyecto.ListIndex) 'l_arr_Proyec(cmb_Proyecto.ListIndex + 1).Genera_Codigo
       grd_Listad.TextMatrix(grd_Listad.Row, 4) = Trim(cmb_Proyecto.Text)
       grd_Listad.TextMatrix(grd_Listad.Row, 5) = cmb_Moneda.ItemData(cmb_Moneda.ListIndex)
       grd_Listad.TextMatrix(grd_Listad.Row, 6) = Trim(cmb_Moneda.Text)
       grd_Listad.TextMatrix(grd_Listad.Row, 7) = txt_GasNot_Mto.Text      'GASPAR_GASNOT_MTO
       grd_Listad.TextMatrix(grd_Listad.Row, 8) = txt_RegMin_TasInm.Text   'GASPAR_REGMIN_TASINM
       grd_Listad.TextMatrix(grd_Listad.Row, 9) = txt_RegMin_FacInm.Text   'GASPAR_REGMIN_FACINM
       grd_Listad.TextMatrix(grd_Listad.Row, 10) = txt_RegMin_FicEst.Text  'GASPAR_REGMIN_FICINM
       grd_Listad.TextMatrix(grd_Listad.Row, 11) = txt_RegMin_TasEst.Text  'GASPAR_REGMIN_TASEST
       grd_Listad.TextMatrix(grd_Listad.Row, 12) = txt_RegMin_FacEst.Text  'GASPAR_REGMIN_FACEST
       grd_Listad.TextMatrix(grd_Listad.Row, 13) = txt_RegMin_FicEst.Text  'GASPAR_REGMIN_FICEST
       grd_Listad.TextMatrix(grd_Listad.Row, 14) = txt_RegGar_Tas000.Text  'GASPAR_REGGAR_TAS000
       grd_Listad.TextMatrix(grd_Listad.Row, 15) = txt_RegGar_Tas001.Text  'GASPAR_REGGAR_TAS001
       grd_Listad.TextMatrix(grd_Listad.Row, 16) = txt_RegGar_Tas002.Text  'GASPAR_REGGAR_TAS002
       
       If Trim(grd_Listad.TextMatrix(grd_Listad.Row, 0)) = "" Then
          grd_Listad.TextMatrix(grd_Listad.Row, 18) = 1 'INSERT
       Else
          grd_Listad.TextMatrix(grd_Listad.Row, 18) = 2 'UPDATE
       End If
    End If
    
    Call cmd_Cancelar_Click
End Sub

Private Sub cmd_Borrar_Click()
   If grd_Listad.Rows = 1 Then
      Exit Sub
   End If
   If grd_Listad.Row = 0 Then
      Exit Sub
   End If

   If MsgBox("¿Está seguro que desea eliminar?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   If Trim(grd_Listad.TextMatrix(grd_Listad.Row, 0)) = "" Then
       grd_Listad.RemoveItem (grd_Listad.Row)
   Else
       grd_Listad.TextMatrix(grd_Listad.Row, 6) = 0
       grd_Listad.RowHeight(grd_Listad.Row) = 0
       
       If Trim(grd_Listad.TextMatrix(grd_Listad.Row, 0)) = "" Then
          grd_Listad.TextMatrix(grd_Listad.Row, 17) = 1 'INSERT
       Else
          grd_Listad.TextMatrix(grd_Listad.Row, 17) = 2 'UPDATE
       End If
   End If
End Sub

Private Sub cmd_Cancelar_Click()
   cmb_Producto.ListIndex = -1
   cmb_Proyecto.ListIndex = -1
   cmb_Moneda.ListIndex = -1
   
   txt_GasNot_Mto.Text = "0.00"        'GASPAR_GASNOT_MTO
   txt_RegMin_TasInm.Text = "0.0000"   'GASPAR_REGMIN_TASINM
   txt_RegMin_FacInm.Text = "0.00"     'GASPAR_REGMIN_FACINM
   txt_RegMin_FicInm.Text = "0.00"     'GASPAR_REGMIN_FICINM
   txt_RegMin_TasEst.Text = "0.0000"   'GASPAR_REGMIN_TASEST
   txt_RegMin_FacEst.Text = "0.00"     'GASPAR_REGMIN_FACEST
   txt_RegMin_FicEst.Text = "0.00"     'GASPAR_REGMIN_FICEST
   txt_RegGar_Tas000.Text = "0.0000"   'GASPAR_REGGAR_TAS000
   txt_RegGar_Tas001.Text = "0.0000"   'GASPAR_REGGAR_TAS001
   txt_RegGar_Tas002.Text = "0.0000"   'GASPAR_REGGAR_TAS002
   
   cmb_Producto.Enabled = False
   cmb_Proyecto.Enabled = False
   cmb_Moneda.Enabled = False
   
   txt_GasNot_Mto.Enabled = False
   txt_RegMin_TasInm.Enabled = False
   txt_RegMin_FacInm.Enabled = False
   txt_RegMin_FicInm.Enabled = False
   txt_RegMin_TasEst.Enabled = False
   txt_RegMin_FacEst.Enabled = False
   txt_RegMin_FicEst.Enabled = False
   txt_RegGar_Tas000.Enabled = False
   txt_RegGar_Tas001.Enabled = False
   txt_RegGar_Tas002.Enabled = False
   
   cmd_Aceptar.Enabled = False
   cmd_Cancelar.Enabled = False
   cmd_Nuevo.Enabled = True
   cmd_Borrar.Enabled = True
   cmd_Editar.Enabled = True
         
   grd_Listad.Enabled = True
   cmd_Aceptar.Tag = 0
   Call gs_SetFocus(cmd_Nuevo)
End Sub

Private Sub cmd_Editar_Click()
   If grd_Listad.Rows = 1 Then
      Exit Sub
   End If
   If grd_Listad.Row = 0 Then
      Exit Sub
   End If
     
   Call gs_BuscarCombo_Item(cmb_Producto, grd_Listad.TextMatrix(grd_Listad.Row, 1))
   Call gs_BuscarCombo_Item(cmb_Proyecto, grd_Listad.TextMatrix(grd_Listad.Row, 3))
   Call gs_BuscarCombo_Item(cmb_Moneda, grd_Listad.TextMatrix(grd_Listad.Row, 5))
   
   txt_GasNot_Mto.Text = grd_Listad.TextMatrix(grd_Listad.Row, 7)       'GASPAR_GASNOT_MTO
   txt_RegMin_TasInm.Text = grd_Listad.TextMatrix(grd_Listad.Row, 8)    'GASPAR_REGMIN_TASINM
   txt_RegMin_FacInm.Text = grd_Listad.TextMatrix(grd_Listad.Row, 9)   'GASPAR_REGMIN_FACINM
   txt_RegMin_FicInm.Text = grd_Listad.TextMatrix(grd_Listad.Row, 10)   'GASPAR_REGMIN_FICINM
   txt_RegMin_TasEst.Text = grd_Listad.TextMatrix(grd_Listad.Row, 11)   'GASPAR_REGMIN_TASEST
   txt_RegMin_FacEst.Text = grd_Listad.TextMatrix(grd_Listad.Row, 12)   'GASPAR_REGMIN_FACEST
   txt_RegMin_FicEst.Text = grd_Listad.TextMatrix(grd_Listad.Row, 13)   'GASPAR_REGMIN_FICEST
   txt_RegGar_Tas002.Text = grd_Listad.TextMatrix(grd_Listad.Row, 16)   'GASPAR_REGGAR_TAS002
   
   cmb_Producto.Enabled = True
   cmb_Proyecto.Enabled = True
   cmb_Moneda.Enabled = True
   
   txt_RegMin_TasInm.Enabled = True
   txt_RegMin_FacInm.Enabled = True
   txt_RegMin_FicInm.Enabled = True
   txt_RegMin_TasEst.Enabled = True
   txt_RegMin_FacEst.Enabled = True
   txt_RegMin_FicEst.Enabled = True
   txt_RegGar_Tas000.Enabled = True
   txt_RegGar_Tas001.Enabled = True
   txt_RegGar_Tas002.Enabled = True
   txt_GasNot_Mto.Enabled = True
   
   cmd_Aceptar.Enabled = True
   cmd_Cancelar.Enabled = True
   cmd_Nuevo.Enabled = False
   cmd_Borrar.Enabled = False
   cmd_Editar.Enabled = False
   
   grd_Listad.Enabled = False
   Call gs_SetFocus(cmb_Producto)
   cmd_Aceptar.Tag = 2
End Sub

Private Sub cmd_Nuevo_Click()
   cmb_Producto.ListIndex = -1
   cmb_Proyecto.ListIndex = -1
   cmb_Moneda.ListIndex = -1

   txt_GasNot_Mto.Text = "0.00"        'GASPAR_GASNOT_MTO
   txt_RegMin_TasInm.Text = "0.0000"   'GASPAR_REGMIN_TASINM
   txt_RegMin_FacInm.Text = "0.00"     'GASPAR_REGMIN_FACINM
   txt_RegMin_FicInm.Text = "0.00"     'GASPAR_REGMIN_FICINM
   txt_RegMin_TasEst.Text = "0.0000"   'GASPAR_REGMIN_TASEST
   txt_RegMin_FacEst.Text = "0.00"     'GASPAR_REGMIN_FACEST
   txt_RegMin_FicEst.Text = "0.00"     'GASPAR_REGMIN_FICEST
   txt_RegGar_Tas000.Text = "0.0000"   'GASPAR_REGGAR_TAS000
   txt_RegGar_Tas001.Text = "0.0000"   'GASPAR_REGGAR_TAS001
   txt_RegGar_Tas002.Text = "0.0000"   'GASPAR_REGGAR_TAS002
   
   cmb_Producto.Enabled = True
   cmb_Proyecto.Enabled = True
   cmb_Moneda.Enabled = True
   
   txt_RegMin_TasInm.Enabled = True
   txt_RegMin_FacInm.Enabled = True
   txt_RegMin_FicInm.Enabled = True
   txt_RegMin_TasEst.Enabled = True
   txt_RegMin_FacEst.Enabled = True
   txt_RegMin_FicEst.Enabled = True
   txt_RegGar_Tas000.Enabled = True
   txt_RegGar_Tas001.Enabled = True
   txt_RegGar_Tas002.Enabled = True
   txt_GasNot_Mto.Enabled = True
   
   cmd_Aceptar.Enabled = True
   cmd_Cancelar.Enabled = True
   cmd_Nuevo.Enabled = False
   cmd_Borrar.Enabled = False
   cmd_Editar.Enabled = False
   
   grd_Listad.Enabled = False
   Call gs_SetFocus(cmb_Producto)
   cmd_Aceptar.Tag = 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Call gs_CentraForm(Me)
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
Dim r_str_Parame    As String
Dim r_rst_Genera    As ADODB.Recordset

   pnl_EmpPer.Caption = moddat_g_str_CodGrp & " - " & moddat_g_str_DesGrp
   txt_DirEle1.Text = ""
   txt_DirEle2.Text = ""
   txt_DirEle3.Text = ""
   txt_DirEle4.Text = ""
   txt_DirEle5.Text = ""
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_Moneda, 1, "204")
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT A.PRODUC_CODIGO, A.PRODUC_DESCRI "
   r_str_Parame = r_str_Parame & "   FROM CRE_PRODUC A "
   r_str_Parame = r_str_Parame & "  WHERE PRODUC_SITCOM = 1 "
   r_str_Parame = r_str_Parame & "    AND PRODUC_CODCLA = 4 "
   r_str_Parame = r_str_Parame & "  ORDER BY PRODUC_CODIGO ASC "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      Do While Not r_rst_Genera.EOF
         cmb_Producto.AddItem Trim$(r_rst_Genera!PRODUC_DESCRI)
         cmb_Producto.ItemData(cmb_Producto.NewIndex) = CLng(r_rst_Genera!Produc_Codigo)
         r_rst_Genera.MoveNext
      Loop
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT DATGEN_CODIGO, DATGEN_TITULO "
   r_str_Parame = r_str_Parame & "   FROM PRY_DATGEN A "
   r_str_Parame = r_str_Parame & "  WHERE DATGEN_SITUAC = 1 "
   r_str_Parame = r_str_Parame & "  ORDER BY DATGEN_TITULO "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      Do While Not r_rst_Genera.EOF
         cmb_Proyecto.AddItem Trim$(r_rst_Genera!DATGEN_TITULO)
         cmb_Proyecto.ItemData(cmb_Proyecto.NewIndex) = CLng(r_rst_Genera!DATGEN_CODIGO)
         r_rst_Genera.MoveNext
      Loop
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
   
   Call cmd_Cancelar_Click
   Call gs_LimpiaGrid(grd_Listad)
   
   grd_Listad.ColWidth(0) = 0       'ITEM
   grd_Listad.ColWidth(1) = 0       'CODIGO_PRODUCTO
   grd_Listad.ColWidth(2) = 3290    'NOMBRE PRODUCTO
   grd_Listad.ColWidth(3) = 0       'CODIGO_PROPYECTO
   grd_Listad.ColWidth(4) = 4470    'NOMBRE PROYECTO
   grd_Listad.ColWidth(5) = 0       'CODIGO_MONEDA
   grd_Listad.ColWidth(6) = 900     'MONEDA
   grd_Listad.ColWidth(7) = 1400    'GASPAR_GASNOT_MTO
   grd_Listad.ColWidth(8) = 1400    'GASPAR_REGMIN_TASINM
   grd_Listad.ColWidth(9) = 1400    'GASPAR_REGMIN_FACINM
   grd_Listad.ColWidth(10) = 1400   'GASPAR_REGMIN_FICEST
   grd_Listad.ColWidth(11) = 1400   'GASPAR_REGMIN_TASEST
   grd_Listad.ColWidth(12) = 1400   'GASPAR_REGMIN_FACEST
   grd_Listad.ColWidth(13) = 1400   'GASPAR_REGMIN_FICEST
   grd_Listad.ColWidth(14) = 1400   'GASPAR_REGGAR_TAS000
   grd_Listad.ColWidth(15) = 1400   'GASPAR_REGGAR_TAS001
   grd_Listad.ColWidth(16) = 1400   'GASPAR_REGGAR_TAS002
   grd_Listad.ColWidth(17) = 0      'ESTADO
   grd_Listad.ColWidth(18) = 0      'INPUT
   
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter
   grd_Listad.ColAlignment(6) = flexAlignLeftCenter
   grd_Listad.ColAlignment(7) = flexAlignRightCenter
      
   grd_Listad.Rows = 1
   grd_Listad.TextMatrix(0, 0) = "ITEM"
   grd_Listad.TextMatrix(0, 1) = ""
   grd_Listad.TextMatrix(0, 2) = "NOMBRE PRODUCTO"
   grd_Listad.TextMatrix(0, 3) = ""
   grd_Listad.TextMatrix(0, 4) = "NOMBRE PROYECTO"
   grd_Listad.TextMatrix(0, 5) = ""
   grd_Listad.TextMatrix(0, 6) = "MONEDA"
   grd_Listad.TextMatrix(0, 7) = "MTO GAST NOT"
   grd_Listad.TextMatrix(0, 8) = "MIN TAS INM"
   grd_Listad.TextMatrix(0, 9) = "MIN FAC INM"
   grd_Listad.TextMatrix(0, 10) = "MIN FIC INM"
   grd_Listad.TextMatrix(0, 11) = "MIN TAS EST"
   grd_Listad.TextMatrix(0, 12) = "MIN FAC EST"
   grd_Listad.TextMatrix(0, 13) = "MIN FIC EST"
   grd_Listad.TextMatrix(0, 14) = "GAR TASA"
   grd_Listad.TextMatrix(0, 15) = "GAR TAS ADI1"
   grd_Listad.TextMatrix(0, 16) = "GAR TAS ADI2"
   
   Dim r_int_Fila As Integer
   For r_int_Fila = 1 To 16
       grd_Listad.Row = 0
       grd_Listad.Col = r_int_Fila
       grd_Listad.CellAlignment = flexAlignCenterCenter
       grd_Listad.CellBackColor = &H4000&     '&HE0E0E0
       grd_Listad.CellForeColor = &HFFFFFF    '&HE0E0E0
   Next
End Sub

Private Sub fs_Buscar()
Dim r_str_Parame    As String
Dim r_rst_Princi    As ADODB.Recordset

   l_str_Codigo = 0
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT DATEMP_CODEMP, DATEMP_IMPORT, DATEMP_DIRELE1, "
   r_str_Parame = r_str_Parame & "        DATEMP_DIRELE2, DATEMP_DIRELE3, DATEMP_DIRELE4, DATEMP_DIRELE5 "
   r_str_Parame = r_str_Parame & "   FROM MNT_DATEMP "
   r_str_Parame = r_str_Parame & "  WHERE DATEMP_CODEMP = '" & moddat_g_str_CodGrp & "' "
   r_str_Parame = r_str_Parame & "    AND DATEMP_TIPTAB = 2 "

   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      r_rst_Princi.MoveFirst
      l_str_Codigo = r_rst_Princi!DATEMP_CODEMP
      txt_DirEle1.Text = Trim(r_rst_Princi!DATEMP_DIRELE1 & "")
      txt_DirEle2.Text = Trim(r_rst_Princi!DATEMP_DIRELE2 & "")
      txt_DirEle3.Text = Trim(r_rst_Princi!DATEMP_DIRELE3 & "")
      txt_DirEle4.Text = Trim(r_rst_Princi!DATEMP_DIRELE4 & "")
      txt_DirEle5.Text = Trim(r_rst_Princi!DATEMP_DIRELE5 & "")
   End If
      
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT GASPAR_CODGAS       , GASPAR_CODEMP           , GASPAR_TIPTAB        , GASPAR_CODPRD        , TRIM(C.PRODUC_DESCRI) AS NOMBRE_PRODUCTO     , "
   r_str_Parame = r_str_Parame & "        GASPAR_CODPRY       , TRIM(B.DATGEN_TITULO) AS NOMBRE_PROYECTO       , GASPAR_CODMON        , TRIM(D.PARDES_DESCRI) AS MONEDA              , "
   r_str_Parame = r_str_Parame & "        GASPAR_CODGAS       , GASPAR_CODEMP           , GASPAR_TIPTAB        , GASPAR_CODPRD        , GASPAR_CODPRY                                , "
   r_str_Parame = r_str_Parame & "        GASPAR_CODMON       , GASPAR_GASNOT_MTO       , GASPAR_REGMIN_TASINM , GASPAR_REGMIN_FACINM , GASPAR_REGMIN_FICINM , GASPAR_REGMIN_TASEST  , "
   r_str_Parame = r_str_Parame & "        GASPAR_REGMIN_FACEST, GASPAR_REGMIN_FICEST    , GASPAR_REGGAR_TAS000 , GASPAR_REGGAR_TAS001 , GASPAR_REGGAR_TAS002 , "
   r_str_Parame = r_str_Parame & "        GASPAR_SITUAC "
   r_str_Parame = r_str_Parame & "   FROM TRA_GASPAR A "
   r_str_Parame = r_str_Parame & "  INNER JOIN PRY_DATGEN B ON B.DATGEN_CODIGO = A.GASPAR_CODPRY "
   r_str_Parame = r_str_Parame & "  INNER JOIN CRE_PRODUC C ON C.PRODUC_CODIGO = A.GASPAR_CODPRD "
   r_str_Parame = r_str_Parame & "  INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 204 AND D.PARDES_CODITE = A.GASPAR_CODMON "
   r_str_Parame = r_str_Parame & "  WHERE GASPAR_CODEMP = '" & moddat_g_str_CodGrp & "' "
   r_str_Parame = r_str_Parame & "    AND GASPAR_TIPTAB = 2 "
   r_str_Parame = r_str_Parame & "    AND GASPAR_SITUAC = 1 "
   r_str_Parame = r_str_Parame & "  ORDER BY NOMBRE_PROYECTO, NOMBRE_PRODUCTO "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      r_rst_Princi.MoveFirst

      Do While Not r_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = r_rst_Princi!GASPAR_CODGAS
         
         grd_Listad.Col = 1
         grd_Listad.Text = r_rst_Princi!GASPAR_CODPRD
         
         grd_Listad.Col = 2
         grd_Listad.Text = Trim(r_rst_Princi!NOMBRE_PRODUCTO)
         
         grd_Listad.Col = 3
         grd_Listad.Text = r_rst_Princi!GASPAR_CODPRY
         
         grd_Listad.Col = 4
         grd_Listad.Text = Trim(r_rst_Princi!NOMBRE_PROYECTO)
         
         grd_Listad.Col = 5
         grd_Listad.Text = r_rst_Princi!GASPAR_CODMON
         
         grd_Listad.Col = 6
         grd_Listad.Text = r_rst_Princi!Moneda
         '------------------------------------
         grd_Listad.Col = 7
         grd_Listad.Text = Format(r_rst_Princi!GASPAR_GASNOT_MTO, "###,###,##0.00")
         
         grd_Listad.Col = 8
         grd_Listad.Text = Format(r_rst_Princi!GASPAR_REGMIN_TASINM, "###,###,##0.00")
         
         grd_Listad.Col = 9
         grd_Listad.Text = Format(r_rst_Princi!GASPAR_REGMIN_FACINM, "###,###,##0.0000")
         
         grd_Listad.Col = 10
         grd_Listad.Text = Format(r_rst_Princi!GASPAR_REGMIN_FICINM, "###,###,##0.0000")
         
         grd_Listad.Col = 11
         grd_Listad.Text = Format(r_rst_Princi!GASPAR_REGMIN_TASEST, "###,###,##0.00")
         
         grd_Listad.Col = 12
         grd_Listad.Text = Format(r_rst_Princi!GASPAR_REGMIN_FACEST, "###,###,##0.0000")
         
         grd_Listad.Col = 13
         grd_Listad.Text = Format(r_rst_Princi!GASPAR_REGMIN_FICEST, "###,###,##0.0000")
         
         grd_Listad.Col = 14
         grd_Listad.Text = Format(r_rst_Princi!GASPAR_REGGAR_TAS000, "###,###,##0.0000")
         
         grd_Listad.Col = 15
         grd_Listad.Text = Format(r_rst_Princi!GASPAR_REGGAR_TAS001, "###,###,##0.0000")
         
         grd_Listad.Col = 16
         grd_Listad.Text = Format(r_rst_Princi!GASPAR_REGGAR_TAS002, "###,###,##0.00")
         '------------------------------------
         grd_Listad.Col = 17
         grd_Listad.Text = r_rst_Princi!GASPAR_SITUAC
         
         grd_Listad.Col = 18
         grd_Listad.Text = 0
                  
         r_rst_Princi.MoveNext
      Loop
      
      grd_Listad.FixedRows = 1
      Call gs_UbiIniGrid(grd_Listad)
   End If
      
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
End Sub

Private Sub cmd_Grabar_Click()
Dim r_str_Parame     As String
Dim r_rst_Genera     As ADODB.Recordset
Dim r_int_Fila       As Integer

   If Len(Trim(moddat_g_str_CodGrp)) = 0 Then
      MsgBox "Debe de seleccionar una notaria", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_GasNot_Mto)
      Exit Sub
   End If
   If Len(Trim(pnl_EmpPer.Caption)) = 0 Then
      MsgBox "Debe de seleccionar una notaria", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_GasNot_Mto)
      Exit Sub
   End If
'   If Len(Trim(txt_DirEle1.Text)) = 0 Then
'      MsgBox "Debe se ingresar un correo.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(txt_DirEle1)
'      Exit Sub
'   Else
'      If gf_ValidarEmail(txt_DirEle1) = False Then
'         MsgBox "El correo ingresado es erroneo.", vbExclamation, modgen_g_str_NomPlt
'         Call gs_SetFocus(txt_DirEle1)
'         Exit Sub
'      End If
'   End If
   If Len(Trim(txt_DirEle2.Text)) > 0 Then
      If gf_ValidarEmail(txt_DirEle2) = False Then
         MsgBox "El correo ingresado es erroneo.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_DirEle2)
         Exit Sub
      End If
   End If
   If Len(Trim(txt_DirEle3.Text)) > 0 Then
      If gf_ValidarEmail(txt_DirEle3) = False Then
         MsgBox "El correo ingresado es erroneo.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_DirEle3)
         Exit Sub
      End If
   End If
   If Len(Trim(txt_DirEle4.Text)) > 0 Then
      If gf_ValidarEmail(txt_DirEle4) = False Then
         MsgBox "El correo ingresado es erroneo.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_DirEle4)
         Exit Sub
      End If
   End If
   If Len(Trim(txt_DirEle5.Text)) > 0 Then
      If gf_ValidarEmail(txt_DirEle5) = False Then
         MsgBox "El correo ingresado es erroneo.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_DirEle5)
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "USP_MNT_DATEMP ("
   r_str_Parame = r_str_Parame & "'" & moddat_g_str_CodGrp & "',"
   r_str_Parame = r_str_Parame & "2," 'Tipo Tabla
   r_str_Parame = r_str_Parame & "0,"
   r_str_Parame = r_str_Parame & "'" & Trim(txt_DirEle1.Text) & "',"
   r_str_Parame = r_str_Parame & "'" & Trim(txt_DirEle2.Text) & "',"
   r_str_Parame = r_str_Parame & "'" & Trim(txt_DirEle3.Text) & "',"
   r_str_Parame = r_str_Parame & "'" & Trim(txt_DirEle4.Text) & "',"
   r_str_Parame = r_str_Parame & "'" & Trim(txt_DirEle5.Text) & "',"
   r_str_Parame = r_str_Parame & "1,"
   'Datos de Auditoria
   r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   r_str_Parame = r_str_Parame & "'" & modgen_g_str_NombPC & "', "
   r_str_Parame = r_str_Parame & "'" & UCase(App.EXEName) & "', "
   r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodSuc & "', "
   r_str_Parame = r_str_Parame & "1) "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 2) Then
      Exit Sub
   End If
   
   For r_int_Fila = 1 To grd_Listad.Rows - 1
       If grd_Listad.TextMatrix(r_int_Fila, 18) = 1 Or grd_Listad.TextMatrix(r_int_Fila, 18) = 2 Then
          r_str_Parame = ""
          r_str_Parame = r_str_Parame & " USP_TRA_GASPAR ( "
          r_str_Parame = r_str_Parame & "'" & grd_Listad.TextMatrix(r_int_Fila, 0) & "',"                         'GASPAR_CODGAS
          r_str_Parame = r_str_Parame & "'" & moddat_g_str_CodGrp & "',"                                          'GASPAR_CODEMP
          r_str_Parame = r_str_Parame & "2,"                                                                      'GASPAR_TIPTAB
          r_str_Parame = r_str_Parame & "'" & Format(Trim(grd_Listad.TextMatrix(r_int_Fila, 1)), "000") & "',"    'GASPAR_CODPRDS
          r_str_Parame = r_str_Parame & "'" & Format(Trim(grd_Listad.TextMatrix(r_int_Fila, 3)), "000000") & "'," 'GASPAR_CODPRY
          r_str_Parame = r_str_Parame & "'" & grd_Listad.TextMatrix(r_int_Fila, 5) & "',"                         'GASPAR_CODMON
          r_str_Parame = r_str_Parame & "0,"                                                                      'GASPAR_GASTAS_MTO
          r_str_Parame = r_str_Parame & CDbl(grd_Listad.TextMatrix(r_int_Fila, 7)) & ","                          'GASPAR_GASNOT_MTO
          r_str_Parame = r_str_Parame & CDbl(grd_Listad.TextMatrix(r_int_Fila, 8)) & ","                          'GASPAR_REGMIN_TASINM
          r_str_Parame = r_str_Parame & CDbl(grd_Listad.TextMatrix(r_int_Fila, 9)) & ","                          'GASPAR_REGMIN_FACINM
          r_str_Parame = r_str_Parame & CDbl(grd_Listad.TextMatrix(r_int_Fila, 10)) & ","                         'GASPAR_REGMIN_FICINM
          r_str_Parame = r_str_Parame & CDbl(grd_Listad.TextMatrix(r_int_Fila, 11)) & ","                         'GASPAR_REGMIN_TASEST
          r_str_Parame = r_str_Parame & CDbl(grd_Listad.TextMatrix(r_int_Fila, 12)) & ","                         'GASPAR_REGMIN_FACEST
          r_str_Parame = r_str_Parame & CDbl(grd_Listad.TextMatrix(r_int_Fila, 13)) & ","                         'GASPAR_REGMIN_FICEST
          r_str_Parame = r_str_Parame & CDbl(grd_Listad.TextMatrix(r_int_Fila, 14)) & ","                         'GASPAR_REGGAR_TAS000
          r_str_Parame = r_str_Parame & CDbl(grd_Listad.TextMatrix(r_int_Fila, 15)) & ","                         'GASPAR_REGGAR_TAS001
          r_str_Parame = r_str_Parame & CDbl(grd_Listad.TextMatrix(r_int_Fila, 16)) & ","                         'GASPAR_REGGAR_TAS002
          r_str_Parame = r_str_Parame & CDbl(grd_Listad.TextMatrix(r_int_Fila, 17)) & ","                         'GASPAR_SITUAC
          'Datos de Auditoria
          r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodUsu & "', "
          r_str_Parame = r_str_Parame & "'" & modgen_g_str_NombPC & "', "
          r_str_Parame = r_str_Parame & "'" & UCase(App.EXEName) & "', "
          r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodSuc & "', "
          r_str_Parame = r_str_Parame & grd_Listad.TextMatrix(r_int_Fila, 18) & ") "
          
         If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 2) Then
            Exit Sub
         End If
       End If
   Next
  
   MsgBox "El proceso se grabó exitosamente.", vbInformation, modgen_g_str_NomPlt
   Unload Me
End Sub

Private Sub txt_DirEle1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_DirEle2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "@_.-")
   End If
End Sub

Private Sub txt_DirEle2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_DirEle3)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "@_.-")
   End If
End Sub

Private Sub txt_DirEle3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_DirEle4)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "@_.-")
   End If
End Sub

Private Sub txt_DirEle4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_DirEle5)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "@_.-")
   End If
End Sub

Private Sub txt_DirEle5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Nuevo)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "@_.-")
   End If
End Sub

Private Sub txt_GasNot_Mto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Aceptar)
   End If
End Sub

Private Sub txt_RegGar_Tas000_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_RegGar_Tas001)
   End If
End Sub

Private Sub txt_RegGar_Tas001_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_RegGar_Tas002)
   End If
End Sub

Private Sub txt_RegGar_Tas002_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_GasNot_Mto)
   End If
End Sub

Private Sub txt_RegMin_FacEst_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_RegMin_FicEst)
   End If
End Sub

Private Sub txt_RegMin_FacInm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_RegMin_FicInm)
   End If
End Sub

Private Sub txt_RegMin_FicEst_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_RegGar_Tas000)
   End If
End Sub

Private Sub txt_RegMin_FicInm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_RegMin_TasEst)
   End If
End Sub

Private Sub txt_RegMin_TasEst_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_RegMin_FacEst)
   End If
End Sub

Private Sub txt_RegMin_TasInm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_RegMin_FacInm)
   End If
End Sub

Private Function fs_Valida_Insercion(ByVal p_CodPrd As String, ByVal p_CodPry As String) As Integer
Dim r_str_Parame        As String

   fs_Valida_Insercion = 0
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "  SELECT COUNT(*) AS CONTADOR "
   r_str_Parame = r_str_Parame & "    FROM TRA_GASPAR "
   r_str_Parame = r_str_Parame & "         INNER JOIN PRY_DATGEN ON DATGEN_CODIGO = GASPAR_CODPRY AND TRIM(TO_CHAR(TRIM(DATGEN_CODNOT), '000000')) = GASPAR_CODEMP "
   r_str_Parame = r_str_Parame & "   WHERE GASPAR_CODPRD = '" & Format(p_CodPrd, "000") & "'"
   r_str_Parame = r_str_Parame & "     AND GASPAR_CODPRY = '" & Format(p_CodPry, "000000") & "'"
   r_str_Parame = r_str_Parame & "     AND GASPAR_TIPTAB = 2 "
   
   If Not gf_EjecutaSQL(r_str_Parame, g_rst_Genera, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      If g_rst_Genera!CONTADOR = 0 Then
         fs_Valida_Insercion = 1
      End If
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function
