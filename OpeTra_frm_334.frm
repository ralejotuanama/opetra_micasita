VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_Con_Cuadre_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   9690
   ClientLeft      =   5085
   ClientTop       =   2115
   ClientWidth     =   11805
   Icon            =   "OpeTra_frm_334.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9657.646
   ScaleMode       =   0  'User
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel111 
      Height          =   9675
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   11820
      _Version        =   65536
      _ExtentX        =   20849
      _ExtentY        =   17066
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
      Begin Threed.SSPanel SSPanel4 
         Height          =   2445
         Left            =   30
         TabIndex        =   40
         Top             =   1380
         Width           =   11775
         _Version        =   65536
         _ExtentX        =   20770
         _ExtentY        =   4313
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
            Height          =   2055
            Left            =   30
            TabIndex        =   7
            Top             =   360
            Width           =   11685
            _ExtentX        =   20611
            _ExtentY        =   3625
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
            TabIndex        =   42
            Top             =   90
            Width           =   1875
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   675
         Left            =   30
         TabIndex        =   38
         Top             =   30
         Width           =   11775
         _Version        =   65536
         _ExtentX        =   20770
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
            TabIndex        =   41
            Top             =   30
            Width           =   5565
            _Version        =   65536
            _ExtentX        =   9816
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "Cuadre de Operaciones (Maestro)"
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
            Picture         =   "OpeTra_frm_334.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   645
         Left            =   30
         TabIndex        =   39
         Top             =   720
         Width           =   11775
         _Version        =   65536
         _ExtentX        =   20770
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
            Left            =   8160
            Picture         =   "OpeTra_frm_334.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Buscar Crédito por Número de Operación"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   10560
            Picture         =   "OpeTra_frm_334.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   11160
            Picture         =   "OpeTra_frm_334.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Enabled         =   0   'False
            Height          =   585
            Left            =   9960
            Picture         =   "OpeTra_frm_334.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_VerPag 
            Enabled         =   0   'False
            Height          =   585
            Left            =   8760
            Picture         =   "OpeTra_frm_334.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Consulta Pagos del Cliente"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ImpCro 
            Enabled         =   0   'False
            Height          =   585
            Left            =   9360
            Picture         =   "OpeTra_frm_334.frx":14B8
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Consulta Cronograma de Pagos"
            Top             =   30
            Width           =   585
         End
         Begin MSMask.MaskEdBox msk_NumOpe 
            Height          =   315
            Left            =   1410
            TabIndex        =   0
            Top             =   180
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            Mask            =   "###-##-#####"
            PromptChar      =   " "
         End
         Begin VB.Label lbl_Numero 
            Caption         =   "Nro. Operación:"
            Height          =   285
            Left            =   120
            TabIndex        =   43
            Top             =   210
            Width           =   1275
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   5805
         Left            =   30
         TabIndex        =   44
         Top             =   3840
         Width           =   11775
         _Version        =   65536
         _ExtentX        =   20770
         _ExtentY        =   10239
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
         Begin TabDlg.SSTab SSTab1 
            Height          =   5625
            Left            =   60
            TabIndex        =   45
            Top             =   90
            Width           =   11625
            _ExtentX        =   20505
            _ExtentY        =   9922
            _Version        =   393216
            Tabs            =   1
            TabsPerRow      =   1
            TabHeight       =   520
            TabCaption(0)   =   "Datos Principales"
            TabPicture(0)   =   "OpeTra_frm_334.frx":17C2
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Frame1"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Frame2"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Frame3"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Frame4"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "Frame5"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "Frame6"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).ControlCount=   6
            Begin VB.Frame Frame6 
               Caption         =   "Numero de Operacion Externos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1065
               Left            =   7350
               TabIndex        =   77
               Top             =   450
               Width           =   4125
               Begin VB.TextBox txt_NMiVivi 
                  Height          =   315
                  Left            =   1620
                  MaxLength       =   25
                  TabIndex        =   13
                  Top             =   600
                  Width           =   2355
               End
               Begin VB.TextBox txt_NCofide 
                  Height          =   315
                  Left            =   1620
                  MaxLength       =   20
                  TabIndex        =   12
                  Top             =   270
                  Width           =   2355
               End
               Begin VB.Label Label17 
                  AutoSize        =   -1  'True
                  Caption         =   "Nº MiVivienda:"
                  Height          =   195
                  Left            =   150
                  TabIndex        =   79
                  Top             =   630
                  Width           =   1050
               End
               Begin VB.Label Label16 
                  AutoSize        =   -1  'True
                  Caption         =   "Código Prestatario:"
                  Height          =   195
                  Left            =   150
                  TabIndex        =   78
                  Top             =   300
                  Width           =   1335
               End
            End
            Begin VB.Frame Frame5 
               Caption         =   "Datos de la Garantia"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1425
               Left            =   7350
               TabIndex        =   73
               Top             =   1560
               Width           =   4125
               Begin VB.ComboBox cmb_TipGar 
                  Height          =   315
                  Left            =   1620
                  Style           =   2  'Dropdown List
                  TabIndex        =   17
                  Top             =   300
                  Width           =   2370
               End
               Begin VB.ComboBox cmb_TipMon 
                  Height          =   315
                  ItemData        =   "OpeTra_frm_334.frx":17DE
                  Left            =   1620
                  List            =   "OpeTra_frm_334.frx":17E0
                  Style           =   2  'Dropdown List
                  TabIndex        =   18
                  Top             =   630
                  Width           =   2370
               End
               Begin EditLib.fpDoubleSingle fpd_MonGar 
                  Height          =   315
                  Left            =   1620
                  TabIndex        =   19
                  Top             =   960
                  Width           =   1515
                  _Version        =   196608
                  _ExtentX        =   2672
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
               Begin VB.Label Label11 
                  Caption         =   "Monto Garantía:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   76
                  Top             =   990
                  Width           =   1155
               End
               Begin VB.Label Label10 
                  Caption         =   "Moneda Garantía:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   75
                  Top             =   660
                  Width           =   1425
               End
               Begin VB.Label Label9 
                  Caption         =   "Tipo Garantía:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   74
                  Top             =   330
                  Width           =   1275
               End
            End
            Begin VB.Frame Frame4 
               Caption         =   "Otros Datos de la Operacion"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2415
               Left            =   3690
               TabIndex        =   62
               Top             =   3030
               Width           =   7785
               Begin VB.ComboBox cmb_SitAdj 
                  Height          =   315
                  Left            =   5280
                  Style           =   2  'Dropdown List
                  TabIndex        =   82
                  Top             =   1950
                  Width           =   2370
               End
               Begin VB.TextBox txt_CodigoCustodia 
                  Height          =   315
                  Left            =   5280
                  MaxLength       =   20
                  TabIndex        =   36
                  Top             =   1620
                  Width           =   2355
               End
               Begin VB.ComboBox cmb_Refina 
                  Height          =   315
                  Left            =   5280
                  Style           =   2  'Dropdown List
                  TabIndex        =   32
                  Top             =   300
                  Width           =   2370
               End
               Begin VB.ComboBox cmb_Judici 
                  Height          =   315
                  Left            =   5280
                  Style           =   2  'Dropdown List
                  TabIndex        =   33
                  Top             =   630
                  Width           =   2370
               End
               Begin VB.ComboBox cmb_Castig 
                  Height          =   315
                  Left            =   5280
                  Style           =   2  'Dropdown List
                  TabIndex        =   34
                  Top             =   960
                  Width           =   2370
               End
               Begin VB.ComboBox cmb_EnvCuo 
                  Height          =   315
                  Left            =   5280
                  Style           =   2  'Dropdown List
                  TabIndex        =   35
                  Top             =   1290
                  Width           =   2370
               End
               Begin EditLib.fpDoubleSingle fpd_CapPag 
                  Height          =   315
                  Left            =   1770
                  TabIndex        =   27
                  Top             =   630
                  Width           =   1575
                  _Version        =   196608
                  _ExtentX        =   2778
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
               Begin EditLib.fpDoubleSingle fpd_CapVen 
                  Height          =   315
                  Left            =   1770
                  TabIndex        =   28
                  Top             =   960
                  Width           =   1575
                  _Version        =   196608
                  _ExtentX        =   2778
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
               Begin EditLib.fpDoubleSingle fpd_DiaPag 
                  Height          =   315
                  Left            =   1770
                  TabIndex        =   26
                  Top             =   300
                  Width           =   1575
                  _Version        =   196608
                  _ExtentX        =   2778
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
                  MaxValue        =   "9000000000"
                  MinValue        =   "0"
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
               Begin EditLib.fpDoubleSingle fpd_CuoPen 
                  Height          =   315
                  Left            =   1770
                  TabIndex        =   29
                  Top             =   1290
                  Width           =   1575
                  _Version        =   196608
                  _ExtentX        =   2778
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
                  MaxValue        =   "9000000000"
                  MinValue        =   "0"
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
               Begin EditLib.fpDoubleSingle fpd_CuoPag 
                  Height          =   315
                  Left            =   1770
                  TabIndex        =   30
                  Top             =   1620
                  Width           =   1575
                  _Version        =   196608
                  _ExtentX        =   2778
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
                  MaxValue        =   "9000000000"
                  MinValue        =   "0"
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
               Begin EditLib.fpDoubleSingle fpd_DiaAtr 
                  Height          =   315
                  Left            =   1770
                  TabIndex        =   31
                  Top             =   1950
                  Width           =   1575
                  _Version        =   196608
                  _ExtentX        =   2778
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
                  MaxValue        =   "9000000000"
                  MinValue        =   "0"
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
               Begin VB.Label Label31 
                  AutoSize        =   -1  'True
                  Caption         =   "Adjudicado:"
                  Height          =   195
                  Left            =   3750
                  TabIndex        =   81
                  Top             =   1980
                  Width           =   840
               End
               Begin VB.Label Label30 
                  AutoSize        =   -1  'True
                  Caption         =   "Código Custodia:"
                  Height          =   195
                  Left            =   3750
                  TabIndex        =   80
                  Top             =   1650
                  Width           =   1200
               End
               Begin VB.Label Label20 
                  AutoSize        =   -1  'True
                  Caption         =   "Capital Pagado:"
                  Height          =   195
                  Left            =   150
                  TabIndex        =   72
                  Top             =   660
                  Width           =   1125
               End
               Begin VB.Label Label23 
                  AutoSize        =   -1  'True
                  Caption         =   "Capital Vencido"
                  Height          =   195
                  Left            =   150
                  TabIndex        =   71
                  Top             =   990
                  Width           =   1110
               End
               Begin VB.Label Label19 
                  AutoSize        =   -1  'True
                  Caption         =   "Día de Pago:"
                  Height          =   195
                  Left            =   150
                  TabIndex        =   70
                  Top             =   330
                  Width           =   960
               End
               Begin VB.Label Label22 
                  AutoSize        =   -1  'True
                  Caption         =   "Cuotas Pendientes:"
                  Height          =   195
                  Left            =   150
                  TabIndex        =   69
                  Top             =   1320
                  Width           =   1380
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Cuotas Pagadas:"
                  Height          =   195
                  Left            =   150
                  TabIndex        =   68
                  Top             =   1650
                  Width           =   1215
               End
               Begin VB.Label Label24 
                  AutoSize        =   -1  'True
                  Caption         =   "Día de Atraso:"
                  Height          =   195
                  Left            =   150
                  TabIndex        =   67
                  Top             =   1980
                  Width           =   1035
               End
               Begin VB.Label Label25 
                  Caption         =   "Refinanciado:"
                  Height          =   195
                  Left            =   3750
                  TabIndex        =   66
                  Top             =   330
                  Width           =   1275
               End
               Begin VB.Label Label26 
                  Caption         =   "Judicial:"
                  Height          =   195
                  Left            =   3750
                  TabIndex        =   65
                  Top             =   660
                  Width           =   1275
               End
               Begin VB.Label Label27 
                  Caption         =   "Castigado:"
                  Height          =   195
                  Left            =   3750
                  TabIndex        =   64
                  Top             =   990
                  Width           =   1275
               End
               Begin VB.Label Label28 
                  Caption         =   "Envio de Cuotas:"
                  Height          =   195
                  Left            =   3750
                  TabIndex        =   63
                  Top             =   1320
                  Width           =   1275
               End
            End
            Begin VB.Frame Frame3 
               Caption         =   "Datos del Prestamo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2415
               Left            =   150
               TabIndex        =   55
               Top             =   3030
               Width           =   3435
               Begin EditLib.fpDoubleSingle fpd_MonPre 
                  Height          =   315
                  Left            =   1620
                  TabIndex        =   20
                  Top             =   300
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
               Begin EditLib.fpDoubleSingle fpd_IntCap 
                  Height          =   315
                  Left            =   1620
                  TabIndex        =   21
                  Top             =   630
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
               Begin EditLib.fpDoubleSingle fpd_TotPre 
                  Height          =   315
                  Left            =   1620
                  TabIndex        =   22
                  Top             =   960
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
               Begin EditLib.fpDoubleSingle fpd_CuoFij 
                  Height          =   315
                  Left            =   1620
                  TabIndex        =   23
                  Top             =   1290
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
               Begin EditLib.fpDoubleSingle fpd_SalCap 
                  Height          =   315
                  Left            =   1620
                  TabIndex        =   24
                  Top             =   1620
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
               Begin EditLib.fpDoubleSingle fpd_SalCapTC 
                  Height          =   315
                  Left            =   1620
                  TabIndex        =   25
                  Top             =   1950
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
               Begin VB.Label Label15 
                  AutoSize        =   -1  'True
                  Caption         =   "Monto Cuota Fija:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   61
                  Top             =   1320
                  Width           =   1215
               End
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  Caption         =   "Monto Prestamo:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   60
                  Top             =   330
                  Width           =   1200
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  Caption         =   "Interes Capitalizado:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   59
                  Top             =   660
                  Width           =   1425
               End
               Begin VB.Label Label14 
                  AutoSize        =   -1  'True
                  Caption         =   "Total Prestamo:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   58
                  Top             =   990
                  Width           =   1080
               End
               Begin VB.Label Label18 
                  AutoSize        =   -1  'True
                  Caption         =   "Saldo Capital:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   57
                  Top             =   1650
                  Width           =   945
               End
               Begin VB.Label Label21 
                  AutoSize        =   -1  'True
                  Caption         =   "Saldo Capital TC:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   56
                  Top             =   1980
                  Width           =   1200
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "Datos del Proyecto"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1425
               Left            =   150
               TabIndex        =   51
               Top             =   1560
               Width           =   7095
               Begin VB.ComboBox cmb_NomPry 
                  Height          =   315
                  Left            =   1620
                  Style           =   2  'Dropdown List
                  TabIndex        =   14
                  Top             =   300
                  Width           =   5355
               End
               Begin VB.ComboBox cmb_EstVin 
                  Height          =   315
                  Left            =   1620
                  Style           =   2  'Dropdown List
                  TabIndex        =   15
                  Top             =   630
                  Width           =   1800
               End
               Begin VB.ComboBox cmb_HipMtz 
                  Height          =   315
                  Left            =   1620
                  Style           =   2  'Dropdown List
                  TabIndex        =   16
                  Top             =   960
                  Width           =   1800
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  Caption         =   "Nombre Proyecto :"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   54
                  Top             =   330
                  Width           =   1320
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  Caption         =   "Vinculado:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   53
                  Top             =   660
                  Width           =   750
               End
               Begin VB.Label Label29 
                  AutoSize        =   -1  'True
                  Caption         =   "Hipoteca Matriz:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   52
                  Top             =   990
                  Width           =   1155
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "Datos del Titular y Conyuge"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1065
               Left            =   150
               TabIndex        =   46
               Top             =   450
               Width           =   7095
               Begin VB.ComboBox cmb_TipDoc 
                  Height          =   315
                  Left            =   1620
                  Style           =   2  'Dropdown List
                  TabIndex        =   8
                  Top             =   300
                  Width           =   2385
               End
               Begin VB.TextBox txt_NumDoc 
                  Height          =   315
                  Left            =   5640
                  MaxLength       =   12
                  TabIndex        =   9
                  Top             =   300
                  Width           =   1305
               End
               Begin VB.TextBox txt_DocCon 
                  Height          =   315
                  Left            =   5640
                  MaxLength       =   12
                  TabIndex        =   11
                  Top             =   630
                  Width           =   1305
               End
               Begin VB.ComboBox cmb_TDoCon 
                  Height          =   315
                  Left            =   1620
                  Style           =   2  'Dropdown List
                  TabIndex        =   10
                  Top             =   630
                  Width           =   2385
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "T.Docum. Titular:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   50
                  Top             =   330
                  Width           =   1230
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  Caption         =   "Nro. Doc. Titular:"
                  Height          =   195
                  Left            =   4170
                  TabIndex        =   49
                  Top             =   330
                  Width           =   1215
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  Caption         =   "Nro. Doc. Conyuge:"
                  Height          =   195
                  Left            =   4170
                  TabIndex        =   48
                  Top             =   660
                  Width           =   1410
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  Caption         =   "T.Docum. Conyuge:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   47
                  Top             =   660
                  Width           =   1425
               End
            End
         End
      End
   End
End
Attribute VB_Name = "frm_Con_Cuadre_02"
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
Dim l_rst_Proy             As ADODB.Recordset

'Lista de variables para ser usadas en el proceso de Auditoria
Dim l_str_TDocTit          As String
Dim l_str_NDocTit          As String
Dim l_str_TDocCon          As String
Dim l_str_NDocCon          As String
Dim l_str_NomPry           As String
Dim l_str_EstVin           As String
Dim l_str_HipMtz           As String
Dim l_dbl_MonPres          As Double
Dim l_dbl_IntCap           As Double
Dim l_dbl_TotPres          As Double
Dim l_dbl_MonCuoF          As Double
Dim l_dbl_SalCap           As Double
Dim l_dbl_SalCapTC         As Double
Dim l_int_DiaPago          As Integer
Dim l_dbl_CapPag           As Double
Dim l_dbl_CapVen           As Double
Dim l_int_CuoPend          As Integer
Dim l_int_CuoPag           As Integer
Dim l_int_DiaAtra          As Integer
Dim l_str_Refinan          As String
Dim l_str_Judici           As String
Dim l_str_Castig           As String
Dim l_str_EnvCuo           As String
Dim l_str_CodPres          As String
Dim l_str_NumViv           As String
Dim l_str_TipGar           As String
Dim l_str_MonGar           As String
Dim l_dbl_MtoGar           As Double
Dim l_str_CodCus           As String
Dim l_str_SitAdj           As String

Private Sub ManejoControles(estado As Boolean)
Dim Control As Object
   
   For Each Control In Me.Controls
      If TypeOf Control Is ComboBox Then
         Control.Enabled = estado
      End If
      
      If TypeOf Control Is TextBox Then
         Control.Enabled = estado
      End If
      
      If TypeOf Control Is fpDoubleSingle Then
         Control.Enabled = estado
      End If
   Next
End Sub

Private Sub cmb_EnvCuo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub
Private Sub cmb_SitAdj_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Call gs_SetFocus(cmd_Grabar)
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
   
   Screen.MousePointer = 11
   Me.Enabled = False
   moddat_g_str_NumOpe = msk_NumOpe.Text
   Call fs_Buscar_Credito
   Call gs_SetFocus(cmb_TipDoc)
   
   'Asignacion de variables usadas en el Proceso de Auditoria.
   l_str_TDocTit = cmb_TipDoc.Text
   l_str_NDocTit = txt_NumDoc.Text
   l_str_TDocCon = cmb_TDoCon.Text
   l_str_NDocCon = txt_DocCon.Text
   l_str_NomPry = cmb_NomPry.Text
   l_str_EstVin = cmb_EstVin.Text
   l_str_HipMtz = cmb_HipMtz.Text
   l_dbl_MonPres = fpd_MonPre.Text
   l_dbl_IntCap = fpd_IntCap.Text
   l_dbl_TotPres = fpd_TotPre.Text
   l_dbl_MonCuoF = fpd_CuoFij.Text
   l_dbl_SalCap = fpd_SalCap.Text
   l_dbl_SalCapTC = fpd_SalCapTC.Text
   l_int_DiaPago = fpd_DiaPag.Text
   l_dbl_CapPag = fpd_CapPag.Text
   l_dbl_CapVen = fpd_CapVen.Text
   l_int_CuoPend = fpd_CuoPen.Text
   l_int_CuoPag = fpd_CuoPag.Text
   l_int_DiaAtra = fpd_DiaAtr.Text
   l_str_Refinan = cmb_Refina.Text
   l_str_Judici = cmb_Judici.Text
   l_str_Castig = cmb_Castig.Text
   l_str_EnvCuo = cmb_EnvCuo.Text
   l_str_CodPres = txt_NCofide.Text
   l_str_NumViv = txt_NMiVivi.Text
   l_str_TipGar = Cmb_TipGar.Text
   l_str_MonGar = cmb_TipMon.Text
   l_dbl_MtoGar = fpd_MonGar.Text
   l_str_CodCus = txt_CodigoCustodia.Text
   l_str_SitAdj = cmb_SitAdj.Text
   
   Call ManejoControles(True)
   
   Me.Enabled = True
   Screen.MousePointer = 0
   
   If moddat_g_int_CntErr = 2 Then
      msk_NumOpe.Text = ""
      msk_NumOpe.Mask = "###-##-#####"
      Call gs_SetFocus(msk_NumOpe)
   End If
End Sub

Private Sub cmd_VerPag_Click()
   frm_Ges_CreHip_05.Show 1
End Sub

Private Sub cmd_ImpCro_Click()
   modmip_g_int_OrdAct = 1
   frm_Ges_CreHip_07.cmd_Cronog.Visible = True
   frm_Ges_CreHip_07.Show 1
End Sub

Private Sub cmd_Grabar_Click()
Dim r_str_Mnsaje  As String
   
   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar Tipo de Documento del Titular.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   If Len(Trim(txt_NumDoc.Text)) = 0 Then
      MsgBox "Debe ingresar Numero de Documento del Titular.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If
   If cmb_TDoCon.Text = "" And Len(Trim(txt_DocCon.Text)) > 0 Then
      MsgBox "Debe seleccionar Tipo de Documento de Conyuge.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TDoCon)
      Exit Sub
   End If
   If cmb_TDoCon.Text <> "" And Len(Trim(txt_DocCon.Text)) = 0 Then
      MsgBox "Debe ingresar Numero de Documento de Conyuge.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_DocCon)
      Exit Sub
   End If
   'If cmb_TipMon.Text = "" Then
   '   MsgBox "Debe seleccionar la Moneda de Garantia.", vbExclamation, modgen_g_str_NomPlt
   '   Call gs_SetFocus(cmb_TipMon)
   '   Exit Sub
   'End If
   'If fpd_MonGar.Text = 0 Then
   '   MsgBox "Debe ingresar Monto de Garantia.", vbExclamation, modgen_g_str_NomPlt
   '   Call gs_SetFocus(fpd_MonGar)
   '   Exit Sub
   'End If
   If fpd_MonPre.Text = 0 Then
      MsgBox "Debe ingresar Monto del Prestamo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(fpd_MonPre)
      Exit Sub
   End If
   If fpd_TotPre.Text = 0 Then
      MsgBox "Debe ingresar Total del Prestamo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(fpd_TotPre)
      Exit Sub
   End If
   If fpd_CuoFij.Text = 0 Then
      MsgBox "Debe ingresar Monto Cuota Fija.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(fpd_CuoFij)
      Exit Sub
   End If
   If fpd_SalCap.Text = 0 Then
      MsgBox "Debe ingresar Saldo Capital.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(fpd_SalCap)
      Exit Sub
   End If
   If fpd_DiaPag.Text = 0 Or fpd_DiaPag.Text = "" Then
      MsgBox "Debe ingresar el Dia de Pago.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(fpd_DiaPag)
      Exit Sub
   End If
   If fpd_CuoPen.Text = 0 Or fpd_CuoPen.Text = "" Then
      MsgBox "Debe ingresar Cuota(s) Pendiente(s).", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(fpd_CuoPen)
      Exit Sub
   End If
   If cmb_Refina.ListIndex = -1 Then
      MsgBox "Debe seleccionar si el cliente esta Refinanciado.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Refina)
      Exit Sub
   End If
   If cmb_Judici.ListIndex = -1 Then
      MsgBox "Debe seleccionar si el cliente esta en Judicial.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Judici)
      Exit Sub
   End If
   If cmb_Castig.ListIndex = -1 Then
      MsgBox "Debe seleccionar si el cliente esta en Castigo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Castig)
      Exit Sub
   End If
   If cmb_EnvCuo.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Envio de Cuotas.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_EnvCuo)
      Exit Sub
   End If
   If cmb_SitAdj.ListIndex = -1 Then
      MsgBox "Debe seleccionar si el cliente esta en Adjudicado.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_SitAdj)
      Exit Sub
   End If
   
   r_str_Mnsaje = "¿Está seguro de grabar los datos?"
   If MsgBox(r_str_Mnsaje, vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Grabando Información del Cliente
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   Screen.MousePointer = 11
      
   g_str_Parame = ""
   g_str_Parame = "USP_PRY_MAEOPE ("
   g_str_Parame = g_str_Parame & "'" & msk_NumOpe.Text & "', "
   g_str_Parame = g_str_Parame & "'" & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & "', "
   g_str_Parame = g_str_Parame & "'" & txt_NumDoc.Text & "', "
   
   If cmb_TDoCon.Text = "" Then
      g_str_Parame = g_str_Parame & "'0', "
   Else
      g_str_Parame = g_str_Parame & "'" & CStr(cmb_TDoCon.ItemData(cmb_TDoCon.ListIndex)) & "', "
   End If

   g_str_Parame = g_str_Parame & "'" & txt_DocCon.Text & "', "
   g_str_Parame = g_str_Parame & "'" & CStr(cmb_EstVin.ItemData(cmb_EstVin.ListIndex)) & "', "
   
   If cmb_NomPry.Text = "" Then
      g_str_Parame = g_str_Parame & "'', "
   Else
      g_str_Parame = g_str_Parame & "'" & Proyecto_Buscar(cmb_NomPry.Text) & "', "
   End If
   
   g_str_Parame = g_str_Parame & " " & CDbl(fpd_SalCap.Text) & ", "
   g_str_Parame = g_str_Parame & " " & CDbl(fpd_SalCapTC.Text) & ", "
   g_str_Parame = g_str_Parame & "'" & CStr(Cmb_TipGar.ItemData(Cmb_TipGar.ListIndex)) & "', "
   If cmb_TipMon.ListIndex = -1 Then
      g_str_Parame = g_str_Parame & "'0', "
   Else
      g_str_Parame = g_str_Parame & "'" & CStr(cmb_TipMon.ItemData(cmb_TipMon.ListIndex)) & "', "
   End If
   g_str_Parame = g_str_Parame & " " & CDbl(fpd_MonGar.Text) & ", "
   g_str_Parame = g_str_Parame & "'" & txt_NMiVivi.Text & "', "
   g_str_Parame = g_str_Parame & "'" & txt_NCofide.Text & "', "
   g_str_Parame = g_str_Parame & " " & CDbl(fpd_CuoFij.Text) & ", "
   g_str_Parame = g_str_Parame & " " & CDbl(fpd_MonPre.Text) & ", "
   g_str_Parame = g_str_Parame & " " & CDbl(fpd_IntCap.Text) & ", "
   g_str_Parame = g_str_Parame & " " & CDbl(fpd_TotPre.Text) & ", "
   g_str_Parame = g_str_Parame & " " & CDbl(fpd_CapPag.Text) & ", "
   g_str_Parame = g_str_Parame & " " & CDbl(fpd_CapVen.Text) & ", "
   g_str_Parame = g_str_Parame & "'" & fpd_DiaPag.Text & "', "
   g_str_Parame = g_str_Parame & "'" & fpd_CuoPen.Text & "',"
   g_str_Parame = g_str_Parame & "'" & fpd_CuoPag.Text & "',"
   g_str_Parame = g_str_Parame & "'" & fpd_DiaAtr.Text & "',"
   g_str_Parame = g_str_Parame & " " & CStr(cmb_Refina.ItemData(cmb_Refina.ListIndex)) & ", "
   g_str_Parame = g_str_Parame & " " & CStr(cmb_Judici.ItemData(cmb_Judici.ListIndex)) & ", "
   g_str_Parame = g_str_Parame & " " & CStr(cmb_Castig.ItemData(cmb_Castig.ListIndex)) & ", "
   g_str_Parame = g_str_Parame & " " & CStr(cmb_EnvCuo.ItemData(cmb_EnvCuo.ListIndex)) & ", "
   
   If cmb_HipMtz.ListIndex = -1 Then
      g_str_Parame = g_str_Parame & "'' ,"
   Else
      g_str_Parame = g_str_Parame & "'" & CStr(cmb_HipMtz.ItemData(cmb_HipMtz.ListIndex)) & "', "
   End If
   g_str_Parame = g_str_Parame & "'" & Trim(txt_CodigoCustodia.Text) & "', "
   g_str_Parame = g_str_Parame & " " & CStr(cmb_SitAdj.ItemData(cmb_SitAdj.ListIndex)) & " ) "
    
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      moddat_g_int_CntErr = moddat_g_int_CntErr + 1
   Else
      moddat_g_int_FlgGOK = True
   End If
   
   If moddat_g_int_CntErr = 6 Then
      If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      Else
         moddat_g_int_CntErr = 0
      End If
   End If
   
   Call Grabar_Auditoria
   
   MsgBox "Los datos se registraron correctamente.", vbInformation, modgen_g_str_NomPlt
   Call cmd_Limpia_Click
   Me.Enabled = True
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Limpia_Click()
   Call gs_LimpiaGrid(grd_Listad)
   Call fs_Validar_Botones(False)
   Call fs_Limpiar
   Call gs_SetFocus(msk_NumOpe)
   Call ManejoControles(False)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

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

   r_str_Proceso = "OPERACIONES MAESTRO"
   r_str_Tabla = "CRE_HIPMAE"
   r_str_Usuario = modgen_g_str_CodUsu
   r_str_Terminal = modgen_g_str_NombPC
   r_str_Plataforma = UCase(App.EXEName)
   r_str_Sucursal = modgen_g_str_CodSuc
   r_str_Descri1 = ""
   r_str_Descri2 = ""
   r_str_Descri3 = ""
   
   'Verificacion de datos modificados para ser guardados como Auditoria
   If l_str_TDocTit <> cmb_TipDoc.Text Then
      r_str_Descri = r_str_Descri + "Tipo Documento Titular (Antes: " & l_str_TDocTit & ")  (Nuevo: " & cmb_TipDoc.Text & ")" + Chr(13)
   End If
   If l_str_NDocTit <> txt_NumDoc.Text Then
      r_str_Descri = r_str_Descri + "Nro.Doc.Titular (Antes: " & l_str_NDocTit & ")  (Nuevo: " & txt_NumDoc.Text & ")" + Chr(13)
   End If
   If l_str_TDocCon <> cmb_TDoCon.Text Then
      r_str_Descri = r_str_Descri + "Tipo Documento Conyugue (Antes: " & l_str_TDocCon & ")  (Nuevo: " & cmb_TDoCon.Text & ")" + Chr(13)
   End If
   If l_str_NDocCon <> txt_DocCon.Text Then
      r_str_Descri = r_str_Descri + "Nro.Doc.Conyugue (Antes: " & l_str_NDocCon & ")  (Nuevo: " & txt_DocCon.Text & ")" + Chr(13)
   End If
   If l_str_NomPry <> cmb_NomPry.Text Then
      r_str_Descri = r_str_Descri + "Nombre Proyecto (Antes: " & l_str_NomPry & ")  (Nuevo: " & cmb_NomPry.Text & ")" + Chr(13)
   End If
   If l_str_EstVin <> cmb_EstVin.Text Then
      r_str_Descri = r_str_Descri + "Vinculado (Antes: " & l_str_EstVin & ")  (Nuevo: " & cmb_EstVin.Text & ")" + Chr(13)
   End If
   If l_str_HipMtz <> cmb_HipMtz.Text Then
      r_str_Descri = r_str_Descri + "Hipoteca Matriz (Antes: " & l_str_HipMtz & ")  (Nuevo: " & cmb_HipMtz.Text & ")" + Chr(13)
   End If
   If l_dbl_MonPres <> fpd_MonPre.Text Then
      r_str_Descri = r_str_Descri + "Monto Prestamo (Antes: " & Format(l_dbl_MonPres, "#,###,##0.00") & ")  (Nuevo: " & fpd_MonPre.Text & ")" + Chr(13)
   End If
   If l_dbl_IntCap <> fpd_IntCap.Text Then
      r_str_Descri = r_str_Descri + "Interes Capitalizado (Antes: " & Format(l_dbl_IntCap, "#,###,##0.00") & ")  (Nuevo: " & fpd_IntCap.Text & ")" + Chr(13)
   End If
   If l_dbl_TotPres <> fpd_TotPre.Text Then
      r_str_Descri = r_str_Descri + "Total Prestamo (Antes: " & Format(l_dbl_TotPres, "#,###,##0.00") & ")  (Nuevo: " & fpd_TotPre.Text & ")" + Chr(13)
   End If
   If l_dbl_MonCuoF <> fpd_CuoFij.Text Then
      r_str_Descri = r_str_Descri + "Monto Cuota Fija (Antes: " & Format(l_dbl_MonCuoF, "#,###,##0.00") & ")  (Nuevo: " & fpd_CuoFij.Text & ")" + Chr(13)
   End If
   If l_dbl_SalCap <> fpd_SalCap.Text Then
      r_str_Descri = r_str_Descri + "Saldo Capital (Antes: " & Format(l_dbl_SalCap, "#,###,##0.00") & ")  (Nuevo: " & fpd_SalCap.Text & ")" + Chr(13)
   End If
   If l_dbl_SalCapTC <> fpd_SalCapTC.Text Then
      r_str_Descri = r_str_Descri + "Saldo Capital TC (Antes: " & Format(l_dbl_SalCapTC, "#,###,##0.00") & ")  (Nuevo: " & fpd_SalCapTC.Text & ")" + Chr(13)
   End If
   If l_int_DiaPago <> fpd_DiaPag.Text Then
      r_str_Descri = r_str_Descri + "Dia de Pago (Antes: " & l_int_DiaPago & ")  (Nuevo: " & fpd_DiaPag.Text & ")" + Chr(13)
   End If
   If l_dbl_CapPag <> fpd_CapPag.Text Then
      r_str_Descri = r_str_Descri + "Capital Pagado (Antes: " & Format(l_dbl_CapPag, "#,###,##0.00") & ")  (Nuevo: " & fpd_CapPag.Text & ")" + Chr(13)
   End If
   If l_dbl_CapVen <> fpd_CapVen.Text Then
      r_str_Descri = r_str_Descri + "Capital Vencido (Antes: " & Format(l_dbl_CapVen, "#,###,##0.00") & ")  (Nuevo: " & fpd_CapVen.Text & ")" + Chr(13)
   End If
   If l_int_CuoPend <> fpd_CuoPen.Text Then
      r_str_Descri = r_str_Descri + "Cuotas Pendientes (Antes: " & l_int_CuoPend & ")  (Nuevo: " & fpd_CuoPen.Text & ")" + Chr(13)
   End If
   If l_int_CuoPag <> fpd_CuoPag.Text Then
      r_str_Descri = r_str_Descri + "Cuotas Pagadas (Antes: " & l_int_CuoPag & ")  (Nuevo: " & fpd_CuoPag.Text & ")" + Chr(13)
   End If
   If l_int_DiaAtra <> fpd_DiaAtr.Text Then
      r_str_Descri = r_str_Descri + "Dias Atraso (Antes: " & l_int_DiaAtra & ")  (Nuevo: " & fpd_DiaAtr.Text & ")" + Chr(13)
   End If
   If l_str_Refinan <> cmb_Refina.Text Then
      r_str_Descri = r_str_Descri + "Refinanciado (Antes: " & l_str_Refinan & ")  (Nuevo: " & cmb_Refina.Text & ")" + Chr(13)
   End If
   If l_str_Judici <> cmb_Judici.Text Then
      r_str_Descri = r_str_Descri + "Judicial (Antes: " & l_str_Judici & ")  (Nuevo: " & cmb_Judici.Text & ")" + Chr(13)
   End If
   If l_str_Castig <> cmb_Castig.Text Then
      r_str_Descri = r_str_Descri + "Castigado (Antes: " & l_str_Castig & ")  (Nuevo: " & cmb_Castig.Text & ")" + Chr(13)
   End If
   If l_str_EnvCuo <> cmb_EnvCuo.Text Then
      r_str_Descri = r_str_Descri + "Envio de Cuotas (Antes: " & l_str_EnvCuo & ")  (Nuevo: " & cmb_EnvCuo.Text & ")" + Chr(13)
   End If
   If l_str_CodPres <> txt_NCofide.Text Then
      r_str_Descri = r_str_Descri + "Codigo Prestatario (Antes: " & l_str_CodPres & ")  (Nuevo: " & txt_NCofide.Text & ")" + Chr(13)
   End If
   If l_str_NumViv <> txt_NMiVivi.Text Then
      r_str_Descri = r_str_Descri + "Nº MiVivienda (Antes: " & l_str_NumViv & ")  (Nuevo: " & txt_NMiVivi.Text & ")" + Chr(13)
   End If
   If l_str_TipGar <> Cmb_TipGar.Text Then
      r_str_Descri = r_str_Descri + "Tipo Garantia (Antes: " & l_str_TipGar & ")  (Nuevo: " & Cmb_TipGar.Text & ")" + Chr(13)
   End If
   If l_str_MonGar <> cmb_TipMon.Text Then
      r_str_Descri = r_str_Descri + "Moneda Garantia (Antes: " & l_str_MonGar & ")  (Nuevo: " & cmb_TipMon.Text & ")" + Chr(13)
   End If
   If l_dbl_MtoGar <> fpd_MonGar.Text Then
      r_str_Descri = r_str_Descri + "Monto Garantia (Antes: " & l_dbl_MtoGar & ")  (Nuevo: " & fpd_MonGar.Text & ")" + Chr(13)
   End If
   If l_str_CodCus <> txt_CodigoCustodia.Text Then
      r_str_Descri = r_str_Descri + "Código de Custodia (Antes: " & l_str_CodCus & ")  (Nuevo: " & txt_CodigoCustodia.Text & ")" + Chr(13)
   End If
   If l_str_SitAdj <> cmb_SitAdj.Text Then
      r_str_Descri = r_str_Descri + "Situación de Adjudicación (Antes: " & l_str_SitAdj & ")  (Nuevo: " & cmb_SitAdj.Text & ")" + Chr(13)
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
   g_str_Parame = g_str_Parame & "0, "
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

Private Sub Form_Load()
    Screen.MousePointer = 11
    Me.Caption = modgen_g_str_NomPlt
    
    Call fs_Inicia
    Call fs_Limpiar
    Call fs_Validar_Botones(False)
    Call gs_LimpiaGrid(grd_Listad)
    Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "230")
    Call moddat_gs_Carga_LisIte_Combo(cmb_TDoCon, 1, "230")
    Call moddat_gs_Carga_LisIte_Combo(cmb_EstVin, 1, "214")
    Call moddat_gs_Carga_LisIte_Combo(Cmb_TipGar, 1, "241")
    Call moddat_gs_Carga_LisIte_Combo(cmb_TipMon, 1, "204")
    Call moddat_gs_Carga_LisIte_Combo(cmb_HipMtz, 1, "378")
    Call Proyecto_Listar
        
    Call gs_CentraForm(Me)
    Call gs_SetFocus(msk_NumOpe)
    
    Call ManejoControles(False)
    
    Screen.MousePointer = 0
End Sub
 
Private Sub fs_Inicia()
    'Inicializando Grid de Datos del Crédito
    grd_Listad.ColWidth(0) = 2900
    grd_Listad.ColWidth(1) = 8150
    grd_Listad.ColAlignment(0) = flexAlignLeftCenter
    grd_Listad.ColAlignment(1) = flexAlignLeftCenter
End Sub

Private Sub fs_Limpiar()
   msk_NumOpe.Text = ""
   txt_NumDoc.Text = ""
   txt_DocCon.Text = ""
   txt_NCofide.Text = ""
   txt_NMiVivi.Text = ""
   txt_CodigoCustodia.Text = ""
   
   fpd_MonGar.Text = "0"
   fpd_MonPre.Text = "0"
   fpd_TotPre.Text = "0"
   fpd_IntCap.Text = "0"
   fpd_CuoFij.Text = "0"
   fpd_SalCap.Text = "0"
   fpd_DiaPag.Text = "0"
   fpd_CapPag.Text = "0"
   fpd_SalCapTC.Text = "0"
   fpd_CuoPen.Text = "0"
   fpd_CapVen.Text = "0"
   fpd_CuoPag.Text = "0"
   fpd_DiaAtr.Text = "0"
   
   cmb_TipDoc.ListIndex = -1
   cmb_TDoCon.ListIndex = -1
   cmb_NomPry.ListIndex = -1
   cmb_EstVin.ListIndex = -1
   cmb_HipMtz.ListIndex = -1
   Cmb_TipGar.ListIndex = -1
   cmb_TipMon.ListIndex = -1
   
   cmb_Refina.Clear
   cmb_Refina.AddItem "NO"
   cmb_Refina.ItemData(cmb_Refina.NewIndex) = 0
   cmb_Refina.AddItem "SI"
   cmb_Refina.ItemData(cmb_Refina.NewIndex) = 1
   cmb_Refina.ListIndex = -1
   
   cmb_Judici.Clear
   cmb_Judici.AddItem "NO"
   cmb_Judici.ItemData(cmb_Judici.NewIndex) = 0
   cmb_Judici.AddItem "SI"
   cmb_Judici.ItemData(cmb_Judici.NewIndex) = 1
   cmb_Judici.ListIndex = -1
   
   cmb_Castig.Clear
   cmb_Castig.AddItem "NO"
   cmb_Castig.ItemData(cmb_Castig.NewIndex) = 0
   cmb_Castig.AddItem "SI"
   cmb_Castig.ItemData(cmb_Castig.NewIndex) = 1
   cmb_Castig.ListIndex = -1
   
   cmb_EnvCuo.Clear
   cmb_EnvCuo.AddItem "NO"
   cmb_EnvCuo.ItemData(cmb_EnvCuo.NewIndex) = 0
   cmb_EnvCuo.AddItem "SI"
   cmb_EnvCuo.ItemData(cmb_EnvCuo.NewIndex) = 1
   cmb_EnvCuo.ListIndex = -1
   
   cmb_SitAdj.Clear
   cmb_SitAdj.AddItem "NO"
   cmb_SitAdj.ItemData(cmb_SitAdj.NewIndex) = 0
   cmb_SitAdj.AddItem "SI"
   cmb_SitAdj.ItemData(cmb_SitAdj.NewIndex) = 1
   cmb_SitAdj.ListIndex = -1
   
   'Limpieza de variables usadas para el proceso de Auditoria.
   l_str_TDocTit = ""
   l_str_NDocTit = ""
   l_str_TDocCon = ""
   l_str_NDocCon = ""
   l_str_NomPry = ""
   l_str_EstVin = ""
   l_dbl_MonPres = 0
   l_dbl_IntCap = 0
   l_dbl_TotPres = 0
   l_dbl_MonCuoF = 0
   l_dbl_SalCap = 0
   l_dbl_SalCapTC = 0
   l_int_DiaPago = 0
   l_dbl_CapPag = 0
   l_dbl_CapVen = 0
   l_int_CuoPend = 0
   l_int_CuoPag = 0
   l_int_DiaAtra = 0
   l_str_Refinan = ""
   l_str_Judici = ""
   l_str_Castig = ""
   l_str_CodPres = ""
   l_str_NumViv = ""
   l_str_TipGar = ""
   l_str_MonGar = ""
   l_dbl_MtoGar = 0
   l_str_CodCus = ""
   l_str_SitAdj = ""
   
End Sub
 
Private Sub fs_Validar_Botones(ByVal r_bol_FlagEn As Boolean)
    msk_NumOpe.Enabled = Not r_bol_FlagEn
    cmd_Buscar.Enabled = Not r_bol_FlagEn
    cmd_VerPag.Enabled = r_bol_FlagEn
    cmd_ImpCro.Enabled = r_bol_FlagEn
    cmd_Grabar.Enabled = r_bol_FlagEn
End Sub

Private Sub Proyecto_Listar()
   g_str_Parame = ""
   g_str_Parame = g_str_Parame + " SELECT DATGEN_CODIGO,  "
   g_str_Parame = g_str_Parame + "        CASE WHEN DATGEN_PRYMCS = 1 THEN  TRIM(DATGEN_TITULO) || ' - ' || 'VINCULADO' "
   g_str_Parame = g_str_Parame + "        ELSE TRIM(DATGEN_TITULO) || ' - ' || TRIM(MNT_PARDES.PARDES_DESCRI) END AS DATGEN_TITULO "
   g_str_Parame = g_str_Parame + "   FROM PRY_DATGEN LEFT "
   g_str_Parame = g_str_Parame + "   JOIN MNT_PARDES ON MNT_PARDES.PARDES_CODGRP = 513 AND PRY_DATGEN.DATGEN_CODBCO=MNT_PARDES.PARDES_CODITE "
   g_str_Parame = g_str_Parame + "  ORDER BY DATGEN_TITULO"

   If Not gf_EjecutaSQL(g_str_Parame, l_rst_Proy, 3) Then
      Exit Sub
   End If
   
   cmb_NomPry.Clear
   Do While Not l_rst_Proy.EOF
      cmb_NomPry.AddItem Trim(l_rst_Proy!DATGEN_TITULO)
      l_rst_Proy.MoveNext
   Loop
End Sub

Private Function Proyecto_Buscar(Desc_Titulo As String) As String
Dim r_rst_Proy    As ADODB.Recordset
Dim r_str_Para    As String
   
   r_str_Para = ""
   r_str_Para = r_str_Para & "SELECT DATGEN_CODIGO, "
   r_str_Para = r_str_Para & "       CASE WHEN DATGEN_PRYMCS = 1 THEN  TRIM(DATGEN_TITULO) || ' - ' || 'VINCULADO' "
   r_str_Para = r_str_Para & "       ELSE TRIM(DATGEN_TITULO) || ' - ' || TRIM(MNT_PARDES.PARDES_DESCRI) END AS DATGEN_TITULO "
   r_str_Para = r_str_Para & "  FROM PRY_DATGEN "
   r_str_Para = r_str_Para & "  LEFT JOIN MNT_PARDES ON MNT_PARDES.PARDES_CODGRP = 513 AND PRY_DATGEN.DATGEN_CODBCO = MNT_PARDES.PARDES_CODITE "
   r_str_Para = r_str_Para & " WHERE DATGEN_TITULO ='" & Mid(Desc_Titulo, 1, InStr(1, Desc_Titulo, "-") - 2) & "' "
   
   If Not gf_EjecutaSQL(r_str_Para, r_rst_Proy, 3) Then
      Exit Function
   End If
   
   If Not (r_rst_Proy.EOF And r_rst_Proy.BOF) Then
      Proyecto_Buscar = r_rst_Proy!DATGEN_CODIGO
   End If
End Function

Private Sub fs_Buscar_Credito()
   Dim r_str_CodPry     As String
   Dim r_str_NomPry     As String
   Dim r_str_CodBco     As String
    
   moddat_g_int_CntErr = 1
'   Call gs_LimpiaGrid(grd_Listad)
   
   'Buscando Información del Crédito
   g_str_Parame = " "
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND HIPMAE_SITUAC IN (2,6,9)"
    
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
      MsgBox "Operación se encuentra transferida", vbExclamation, modgen_g_con_OpeTra
      Call gs_LimpiaGrid(grd_Listad)
      Call fs_Limpiar
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      moddat_g_int_CntErr = 2
      Exit Sub
   End If
   
   If g_rst_Princi!HIPMAE_SITUAC = 9 Then
      MsgBox "Operación se encuentra cancelada", vbExclamation, modgen_g_con_OpeTra
      If g_rst_Princi!HIPMAE_SITADJ = 0 Then
         Call gs_LimpiaGrid(grd_Listad)
         Call fs_Limpiar
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         moddat_g_int_CntErr = 2
         Exit Sub
      End If
   End If

   Call fs_Validar_Botones(True)
      
   g_rst_Princi.MoveFirst
    
   'Mostrar datos en cada Frame respectivo, para el mantenimiento del formulario
   '----------------------------------------------------------------------------------------
   Call gs_BuscarCombo_Item(cmb_TipDoc, g_rst_Princi!HIPMAE_TDOCLI)
   Call gs_BuscarCombo_Item(cmb_TDoCon, IIf(IsNull(g_rst_Princi!HIPMAE_TDOCYG), 0, g_rst_Princi!HIPMAE_TDOCYG))
   txt_NumDoc.Text = Trim(g_rst_Princi!HIPMAE_NDOCLI) & ""
   txt_DocCon.Text = Trim(g_rst_Princi!HIPMAE_NDOCYG) & ""

   If (g_rst_Princi!HIPMAE_PRYINM & "") <> "" Then
      l_rst_Proy.MoveFirst
      l_rst_Proy.Find "DATGEN_CODIGO='" & g_rst_Princi!HIPMAE_PRYINM & "'"
      If Not (l_rst_Proy.EOF And l_rst_Proy.BOF) Then
         cmb_NomPry.Text = Trim(l_rst_Proy!DATGEN_TITULO)
      End If
   End If
   
   cmb_EstVin.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!HIPMAE_PRYMCS))
   If g_rst_Princi!HIPMAE_TIPGAR <> 0 Then
      Cmb_TipGar.Text = moddat_gf_Consulta_ParDes("241", CStr(g_rst_Princi!HIPMAE_TIPGAR))
   End If
   
   If Not IsNull(g_rst_Princi!HIPMAE_HIPMTZ) Then
      cmb_HipMtz.Text = moddat_gf_Consulta_ParDes("378", CStr(g_rst_Princi!HIPMAE_HIPMTZ))
   End If
   
   If CStr(g_rst_Princi!HIPMAE_MONGAR) > 0 Then
      cmb_TipMon.Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPMAE_MONGAR))
   End If
   
   Call gs_BuscarCombo_Item(cmb_Refina, g_rst_Princi!HIPMAE_REFINA)
   Call gs_BuscarCombo_Item(cmb_Judici, g_rst_Princi!HIPMAE_JUDICI)
   Call gs_BuscarCombo_Item(cmb_Castig, g_rst_Princi!HIPMAE_CASTIG)
   
   If IsNull(g_rst_Princi!HIPMAE_ENVCUO) Then
      Call gs_BuscarCombo_Item(cmb_EnvCuo, 1)
   Else
      If g_rst_Princi!HIPMAE_ENVCUO = 1 Then
         Call gs_BuscarCombo_Item(cmb_EnvCuo, 1)
      Else
         Call gs_BuscarCombo_Item(cmb_EnvCuo, 0)
      End If
   End If
   
   fpd_MonGar.Text = gf_FormatoNumero(g_rst_Princi!HIPMAE_MTOGAR, 12, 2)
   fpd_MonPre.Text = gf_FormatoNumero(g_rst_Princi!HIPMAE_MTOPRE, 12, 2)
   fpd_TotPre.Text = gf_FormatoNumero(g_rst_Princi!HIPMAE_TOTPRE, 12, 2)
   fpd_IntCap.Text = gf_FormatoNumero(g_rst_Princi!HIPMAE_INTCAP, 12, 2)
   fpd_CuoFij.Text = g_rst_Princi!HIPMAE_CUOFIJ & ""
   txt_NCofide.Text = Trim(g_rst_Princi!HIPMAE_CODCOF) & ""
   txt_NMiVivi.Text = Trim(g_rst_Princi!HIPMAE_OPEMVI) & ""
   
   fpd_SalCap.Text = gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCAP, 12, 2)
   fpd_SalCapTC.Text = gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCON, 12, 2)
   fpd_DiaPag.Text = g_rst_Princi!HIPMAE_DIAPAG & ""
   fpd_CapPag.Text = g_rst_Princi!HIPMAE_PAGCAP & ""
   fpd_CapVen.Text = g_rst_Princi!HIPMAE_CAPVEN & ""
   fpd_CuoPen.Text = g_rst_Princi!HIPMAE_CUOPEN & ""
   fpd_CuoPag.Text = g_rst_Princi!HIPMAE_CUOPAG & ""
   fpd_DiaAtr.Text = g_rst_Princi!HIPMAE_DIAMOR & ""
   
   txt_CodigoCustodia.Text = g_rst_Princi!HIPMAE_CODCUS & ""
   If Not IsNull(g_rst_Princi!HIPMAE_SITADJ) Then
      Call gs_BuscarCombo_Item(cmb_SitAdj, g_rst_Princi!HIPMAE_SITADJ)
   Else
      Call gs_BuscarCombo_Item(cmb_SitAdj, 0)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
      
   'Información del Crédito
   Call modmip_gs_DatNumOpe(moddat_g_str_NumOpe, grd_Listad)
End Sub

Private Sub fs_Buscar_Credito_ant()
   Dim r_str_CodPry     As String
   Dim r_str_NomPry     As String
   Dim r_str_CodBco     As String
    
   moddat_g_int_CntErr = 1
   Call gs_LimpiaGrid(grd_Listad)
   
   'Buscando Información del Crédito
   g_str_Parame = " "
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND HIPMAE_SITUAC IN (2,6,9)"
    
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
      MsgBox "Operación se encuentra transferida", vbExclamation, modgen_g_con_OpeTra
      Call gs_LimpiaGrid(grd_Listad)
      Call fs_Limpiar
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      moddat_g_int_CntErr = 2
      Exit Sub
   End If
   
   If g_rst_Princi!HIPMAE_SITUAC = 9 Then
      MsgBox "Operación se encuentra cancelada", vbExclamation, modgen_g_con_OpeTra
      Call gs_LimpiaGrid(grd_Listad)
      Call fs_Limpiar
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      moddat_g_int_CntErr = 2
      Exit Sub
   End If

   Call fs_Validar_Botones(True)
   g_rst_Princi.MoveFirst
    
   'Almacenando en Variables Globales
   moddat_g_int_TipDoc = g_rst_Princi!HIPMAE_TDOCLI
   moddat_g_str_NumDoc = Trim(g_rst_Princi!HIPMAE_NDOCLI) & ""
   moddat_g_str_NumSol = Trim(g_rst_Princi!HIPMAE_NUMSOL)
   moddat_g_str_NumOpe = Trim(g_rst_Princi!HIPMAE_NUMOPE)

   'Obteniendo Nombre de Cliente
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   'Obteniendo Nombre y DOI de Cónyuge
   moddat_g_int_CygTDo = IIf(IsNull(g_rst_Princi!HIPMAE_TDOCYG), 0, g_rst_Princi!HIPMAE_TDOCYG)
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
   
   'Obteniendo Modalidad de Producto
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
   moddat_g_dbl_MtoPre = g_rst_Princi!HIPMAE_MTOPRE           'Monto Préstamo
   moddat_g_int_CuoPen = g_rst_Princi!HIPMAE_CUOPEN           'Cuotas Pendientes
   moddat_g_int_TotCuo = g_rst_Princi!HIPMAE_NUMCUO           'Total de Cuotas
   moddat_g_dbl_SalCap = g_rst_Princi!HIPMAE_SALCAP           'Saldo Capital
    
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
         Case InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd):  grd_Listad.Text = "Nro. Operación Mivivienda"  '"001"
         Case InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd):  grd_Listad.Text = "Nro. Operación COFIDE"      '"003"
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
   
   'Mostrar datos en cada Frame respectivo, para el mantenimiento del formulario
   '----------------------------------------------------------------------------------------
   Call gs_BuscarCombo_Item(cmb_TipDoc, g_rst_Princi!HIPMAE_TDOCLI)
   Call gs_BuscarCombo_Item(cmb_TDoCon, IIf(IsNull(g_rst_Princi!HIPMAE_TDOCYG), 0, g_rst_Princi!HIPMAE_TDOCYG))
   txt_NumDoc.Text = Trim(g_rst_Princi!HIPMAE_NDOCLI) & ""
   txt_DocCon.Text = Trim(g_rst_Princi!HIPMAE_NDOCYG) & ""

   If (g_rst_Princi!HIPMAE_PRYINM & "") <> "" Then
      l_rst_Proy.MoveFirst
      l_rst_Proy.Find "DATGEN_CODIGO='" & g_rst_Princi!HIPMAE_PRYINM & "'"
      If Not (l_rst_Proy.EOF And l_rst_Proy.BOF) Then
         cmb_NomPry.Text = Trim(l_rst_Proy!DATGEN_TITULO)
      End If
   End If
   
   cmb_EstVin.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!HIPMAE_PRYMCS))
   If g_rst_Princi!HIPMAE_TIPGAR <> 0 Then
      Cmb_TipGar.Text = moddat_gf_Consulta_ParDes("241", CStr(g_rst_Princi!HIPMAE_TIPGAR))
   End If
   
   If Not IsNull(g_rst_Princi!HIPMAE_HIPMTZ) Then
      cmb_HipMtz.Text = moddat_gf_Consulta_ParDes("378", CStr(g_rst_Princi!HIPMAE_HIPMTZ))
   End If
   
   If CStr(g_rst_Princi!HIPMAE_MONGAR) > 0 Then
      cmb_TipMon.Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPMAE_MONGAR))
   End If
   
   Call gs_BuscarCombo_Item(cmb_Refina, g_rst_Princi!HIPMAE_REFINA)
   Call gs_BuscarCombo_Item(cmb_Judici, g_rst_Princi!HIPMAE_JUDICI)
   Call gs_BuscarCombo_Item(cmb_Castig, g_rst_Princi!HIPMAE_CASTIG)
   
   If IsNull(g_rst_Princi!HIPMAE_ENVCUO) Then
      Call gs_BuscarCombo_Item(cmb_EnvCuo, 1)
   Else
      If g_rst_Princi!HIPMAE_ENVCUO = 1 Then
         Call gs_BuscarCombo_Item(cmb_EnvCuo, 1)
      Else
         Call gs_BuscarCombo_Item(cmb_EnvCuo, 0)
      End If
   End If
   
   fpd_MonGar.Text = gf_FormatoNumero(g_rst_Princi!HIPMAE_MTOGAR, 12, 2)
   fpd_MonPre.Text = gf_FormatoNumero(g_rst_Princi!HIPMAE_MTOPRE, 12, 2)
   fpd_TotPre.Text = gf_FormatoNumero(g_rst_Princi!HIPMAE_TOTPRE, 12, 2)
   fpd_IntCap.Text = gf_FormatoNumero(g_rst_Princi!HIPMAE_INTCAP, 12, 2)
   fpd_CuoFij.Text = g_rst_Princi!HIPMAE_CUOFIJ & ""
   txt_NCofide.Text = Trim(g_rst_Princi!HIPMAE_CODCOF) & ""
   txt_NMiVivi.Text = Trim(g_rst_Princi!HIPMAE_OPEMVI) & ""
   
   fpd_SalCap.Text = gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCAP, 12, 2)
   fpd_SalCapTC.Text = gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCON, 12, 2)
   fpd_DiaPag.Text = g_rst_Princi!HIPMAE_DIAPAG & ""
   fpd_CapPag.Text = g_rst_Princi!HIPMAE_PAGCAP & ""
   fpd_CapVen.Text = g_rst_Princi!HIPMAE_CAPVEN & ""
   fpd_CuoPen.Text = g_rst_Princi!HIPMAE_CUOPEN & ""
   fpd_CuoPag.Text = g_rst_Princi!HIPMAE_CUOPAG & ""
   fpd_DiaAtr.Text = g_rst_Princi!HIPMAE_DIAMOR & ""
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_Listad)
End Sub

'**************************
Private Sub msk_NumOpe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(cmd_Buscar)
    End If
End Sub

Private Sub cmb_TipDoc_Click()
   If cmb_TipDoc.ListIndex > -1 Then
      Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
         Case 1:  txt_NumDoc.MaxLength = 8
         Case 2:  txt_NumDoc.MaxLength = 12
         Case 3:  txt_NumDoc.MaxLength = 12
      End Select
   End If
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumDoc)
   End If
   If KeyAscii = 8 Then
      cmb_TipDoc.ListIndex = -1
   End If
End Sub

Private Sub txt_CodigoCustodia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Call gs_SetFocus(cmb_SitAdj)
   End If
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Call gs_SetFocus(cmb_TDoCon)
   End If
End Sub

Private Sub cmb_TDoCon_Click()
   If cmb_TDoCon.ListIndex > -1 Then
      Select Case cmb_TDoCon.ItemData(cmb_TDoCon.ListIndex)
         Case 1:  txt_DocCon.MaxLength = 8
         Case 2:  txt_DocCon.MaxLength = 12
         Case 3:  txt_DocCon.MaxLength = 12
      End Select
   End If
End Sub

Private Sub cmb_TDoCon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_DocCon)
   End If
   If KeyAscii = 8 Then
      cmb_TDoCon.ListIndex = -1
   End If
End Sub

Private Sub txt_DocCon_GotFocus()
   Call gs_SelecTodo(txt_DocCon)
End Sub

Private Sub txt_DocCon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(txt_NCofide)
    End If
End Sub

Private Sub txt_NCofide_GotFocus()
   Call gs_SelecTodo(txt_NCofide)
End Sub

Private Sub txt_NCofide_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(txt_NMiVivi)
    End If
End Sub

Private Sub txt_NMiVivi_GotFocus()
   Call gs_SelecTodo(txt_NMiVivi)
End Sub

Private Sub txt_NMiVivi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(cmb_NomPry)
    End If
End Sub

Private Sub cmb_NomPry_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(cmb_EstVin)
    End If
End Sub

Private Sub cmb_EstVin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(cmb_HipMtz)
    End If
End Sub

Private Sub cmb_HipMtz_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(Cmb_TipGar)
    End If
End Sub

Private Sub cmb_TipGar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(cmb_TipMon)
    End If
End Sub

Private Sub cmb_TipMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_MonGar)
    End If
End Sub

Private Sub fpd_MonGar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_MonPre)
    End If
End Sub

Private Sub fpd_MonPre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_IntCap)
    End If
End Sub

Private Sub fpd_IntCap_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_TotPre)
    End If
End Sub

Private Sub fpd_TotPre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_CuoFij)
    End If
End Sub

Private Sub fpd_CuoFij_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_SalCap)
    End If
End Sub

Private Sub fpd_SalCap_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_SalCapTC)
    End If
End Sub

Private Sub fpd_SalCapTC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_DiaPag)
    End If
End Sub

Private Sub fpd_DiaPag_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_CapPag)
    End If
End Sub

Private Sub fpd_CapPag_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_CapVen)
    End If
End Sub

Private Sub fpd_CapVen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_CuoPen)
    End If
End Sub

Private Sub fpd_CuoPen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_CuoPag)
    End If
End Sub

Private Sub fpd_CuoPag_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(fpd_DiaAtr)
    End If
End Sub

Private Sub fpd_DiaAtr_KeyPress(KeyAscii As Integer)
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
        Call gs_SetFocus(cmb_EnvCuo)
    End If
End Sub

