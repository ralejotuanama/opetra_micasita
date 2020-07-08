VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frm_GasAdm_11 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10560
   ClientLeft      =   3120
   ClientTop       =   2115
   ClientWidth     =   11610
   Icon            =   "OpeTra_frm_085.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10560
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10755
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11610
      _Version        =   65536
      _ExtentX        =   20479
      _ExtentY        =   18971
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
         Height          =   2140
         Left            =   30
         TabIndex        =   11
         Top             =   2205
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   3775
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
            Height          =   2010
            Left            =   60
            TabIndex        =   12
            Top             =   60
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   3545
            _Version        =   393216
            Style           =   1
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "Datos del Cliente"
            TabPicture(0)   =   "OpeTra_frm_085.frx":000C
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Datos del Cónyuge"
            TabPicture(1)   =   "OpeTra_frm_085.frx":0028
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(3)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Datos del Inmueble"
            TabPicture(2)   =   "OpeTra_frm_085.frx":0044
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(1)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Datos del Crédito"
            TabPicture(3)   =   "OpeTra_frm_085.frx":0060
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "grd_Listad(2)"
            Tab(3).ControlCount=   1
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1560
               Index           =   0
               Left            =   60
               TabIndex        =   13
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   2752
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   3405
               Index           =   7
               Left            =   -74910
               TabIndex        =   14
               Top             =   3660
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   6006
               _Version        =   393216
               Rows            =   21
               Cols            =   1
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1560
               Index           =   2
               Left            =   -74940
               TabIndex        =   38
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   2752
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1560
               Index           =   1
               Left            =   -74940
               TabIndex        =   39
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   2752
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1560
               Index           =   3
               Left            =   -74940
               TabIndex        =   40
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   2752
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   15
         Top             =   30
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
            TabIndex        =   16
            Top             =   60
            Width           =   5415
            _Version        =   65536
            _ExtentX        =   9551
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Asignación de Gastos de Cierre"
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
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   10860
            Top             =   30
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
            Left            =   10290
            Top             =   30
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_085.frx":007C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   17
         Top             =   740
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   1650
            TabIndex        =   18
            Top             =   60
            Width           =   6465
            _Version        =   65536
            _ExtentX        =   11404
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
            Left            =   1650
            TabIndex        =   19
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
            Left            =   10050
            TabIndex        =   20
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
            Left            =   10050
            TabIndex        =   21
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
         Begin VB.Label Label3 
            Caption         =   "F. Ingreso Solicitud:"
            Height          =   315
            Left            =   8400
            TabIndex        =   25
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label21 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   24
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   23
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "F. Ingreso Instancia:"
            Height          =   315
            Left            =   8400
            TabIndex        =   22
            Top             =   390
            Width           =   1455
         End
      End
      Begin Threed.SSPanel SSPanel39 
         Height          =   645
         Left            =   30
         TabIndex        =   26
         Top             =   1530
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin VB.CommandButton cmd_Export 
            Height          =   585
            Left            =   3630
            Picture         =   "OpeTra_frm_085.frx":0386
            Style           =   1  'Graphical
            TabIndex        =   90
            ToolTipText     =   "Exportar datos a Excel"
            Top             =   30
            Width           =   585
         End
         Begin Threed.SSPanel pnl_Gastos 
            Height          =   345
            Left            =   4950
            TabIndex        =   89
            Top             =   180
            Width           =   5355
            _Version        =   65536
            _ExtentX        =   9446
            _ExtentY        =   609
            _StockProps     =   15
            Caption         =   "FINANCIA GASTOS DE CIERRE"
            ForeColor       =   255
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            FloodColor      =   255
         End
         Begin VB.CommandButton cmd_Comment 
            Height          =   585
            Left            =   3030
            Picture         =   "OpeTra_frm_085.frx":0690
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Enviar comentario soibre gastos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   1830
            Picture         =   "OpeTra_frm_085.frx":099A
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Cancel 
            Height          =   585
            Left            =   2430
            Picture         =   "OpeTra_frm_085.frx":0DDC
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_EdiIte 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_085.frx":10E6
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Edición de Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_NueIte 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_085.frx":13F0
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_BorIte 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_085.frx":16FA
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "OpeTra_frm_085.frx":1A04
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   765
         Left            =   30
         TabIndex        =   27
         Top             =   9780
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin VB.ComboBox cmb_GasAdm 
            Height          =   315
            Left            =   1650
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   60
            Width           =   3855
         End
         Begin EditLib.fpDoubleSingle ipp_Import 
            Height          =   315
            Left            =   1650
            TabIndex        =   9
            Top             =   390
            Width           =   975
            _Version        =   196608
            _ExtentX        =   1720
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
         Begin Threed.SSPanel pnl_MonGas 
            Height          =   315
            Left            =   2610
            TabIndex        =   28
            Top             =   390
            Width           =   2835
            _Version        =   65536
            _ExtentX        =   5001
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
         Begin VB.Label Label5 
            Caption         =   "Concepto de Gasto:"
            Height          =   285
            Left            =   60
            TabIndex        =   30
            Top             =   60
            Width           =   1485
         End
         Begin VB.Label Label8 
            Caption         =   "Importe:"
            Height          =   285
            Left            =   60
            TabIndex        =   29
            Top             =   390
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   2025
         Left            =   30
         TabIndex        =   31
         Top             =   7725
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   3572
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
         Begin MSFlexGridLib.MSFlexGrid grd_GasAdm 
            Height          =   1305
            Left            =   30
            TabIndex        =   7
            Top             =   330
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   2302
            _Version        =   393216
            Rows            =   5
            Cols            =   4
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
            TabIndex        =   32
            Top             =   60
            Width           =   6165
            _Version        =   65536
            _ExtentX        =   10874
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
         Begin Threed.SSPanel SSPanel13 
            Height          =   285
            Left            =   6210
            TabIndex        =   33
            Top             =   60
            Width           =   2685
            _Version        =   65536
            _ExtentX        =   4736
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo de Moneda"
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
            Left            =   8880
            TabIndex        =   34
            Top             =   60
            Width           =   2295
            _Version        =   65536
            _ExtentX        =   4048
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
         Begin Threed.SSPanel pnl_TotGas 
            Height          =   315
            Left            =   8880
            TabIndex        =   35
            Top             =   1650
            Width           =   2295
            _Version        =   65536
            _ExtentX        =   4048
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
            Alignment       =   4
         End
         Begin VB.Label lbl_Moneda 
            Alignment       =   1  'Right Justify
            Caption         =   "US$"
            Height          =   315
            Left            =   8310
            TabIndex        =   37
            Top             =   1650
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "Total de Gastos ==>"
            Height          =   315
            Left            =   6420
            TabIndex        =   36
            Top             =   1650
            Width           =   1635
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   3285
         Left            =   30
         TabIndex        =   41
         Top             =   4390
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   5794
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
         Begin TabDlg.SSTab tab_Seguim 
            Height          =   3190
            Left            =   60
            TabIndex        =   42
            Top             =   60
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   5636
            _Version        =   393216
            Style           =   1
            TabHeight       =   520
            TabCaption(0)   =   "Seguimiento en Instancia"
            TabPicture(0)   =   "OpeTra_frm_085.frx":1E46
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label7"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label2"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label11"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "pnl_DesOcu"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "SSPanel12"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "SSPanel14"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "SSPanel9"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "grd_LisOcu"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "SSPanel8"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "txt_Observ"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "txt_Descar"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).ControlCount=   11
            TabCaption(1)   =   "Excepciones Aplicadas"
            TabPicture(1)   =   "OpeTra_frm_085.frx":1E62
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "lbl_motivo"
            Tab(1).Control(1)=   "Label17"
            Tab(1).Control(2)=   "Label18"
            Tab(1).Control(3)=   "Label19"
            Tab(1).Control(4)=   "SSPanel29"
            Tab(1).Control(5)=   "SSPanel28"
            Tab(1).Control(6)=   "SSPanel27"
            Tab(1).Control(7)=   "SSPanel26"
            Tab(1).Control(8)=   "pnl_motivo"
            Tab(1).Control(9)=   "pnl_TipAut"
            Tab(1).Control(10)=   "pnl_DesExc"
            Tab(1).Control(11)=   "grd_LisExc"
            Tab(1).Control(12)=   "txt_ObsExc"
            Tab(1).ControlCount=   13
            TabCaption(2)   =   "Aprobación Condicionada"
            TabPicture(2)   =   "OpeTra_frm_085.frx":1E7E
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Label12"
            Tab(2).Control(1)=   "Label15"
            Tab(2).Control(2)=   "Label14"
            Tab(2).Control(3)=   "Label16"
            Tab(2).Control(4)=   "SSPanel30"
            Tab(2).Control(5)=   "SSPanel17"
            Tab(2).Control(6)=   "SSPanel18"
            Tab(2).Control(7)=   "pnl_InsCon"
            Tab(2).Control(8)=   "grd_LisCon"
            Tab(2).Control(9)=   "SSPanel20"
            Tab(2).Control(10)=   "SSPanel19"
            Tab(2).Control(11)=   "txt_ObsCon"
            Tab(2).Control(12)=   "txt_LevCon"
            Tab(2).ControlCount=   13
            Begin VB.TextBox txt_LevCon 
               Height          =   620
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   47
               Text            =   "OpeTra_frm_085.frx":1E9A
               Top             =   2520
               Width           =   10005
            End
            Begin VB.TextBox txt_ObsCon 
               Height          =   525
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   46
               Text            =   "OpeTra_frm_085.frx":1E9E
               Top             =   1980
               Width           =   10005
            End
            Begin VB.TextBox txt_ObsExc 
               Height          =   795
               Left            =   -73770
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   45
               Text            =   "OpeTra_frm_085.frx":1EA2
               Top             =   1980
               Width           =   10065
            End
            Begin VB.TextBox txt_Descar 
               Height          =   620
               Left            =   1320
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   44
               Text            =   "OpeTra_frm_085.frx":1EA6
               Top             =   2520
               Width           =   10005
            End
            Begin VB.TextBox txt_Observ 
               Height          =   525
               Left            =   1320
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   43
               Text            =   "OpeTra_frm_085.frx":1EAA
               Top             =   1980
               Width           =   10005
            End
            Begin Threed.SSPanel SSPanel8 
               Height          =   45
               Left            =   30
               TabIndex        =   48
               Top             =   1560
               Width           =   11325
               _Version        =   65536
               _ExtentX        =   19976
               _ExtentY        =   79
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
            Begin MSFlexGridLib.MSFlexGrid grd_LisOcu 
               Height          =   855
               Left            =   30
               TabIndex        =   49
               Top             =   660
               Width           =   11325
               _ExtentX        =   19976
               _ExtentY        =   1508
               _Version        =   393216
               Rows            =   21
               Cols            =   5
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel9 
               Height          =   285
               Left            =   60
               TabIndex        =   50
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Ocurrencia"
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
            Begin Threed.SSPanel SSPanel14 
               Height          =   285
               Left            =   2400
               TabIndex        =   51
               Top             =   360
               Width           =   8595
               _Version        =   65536
               _ExtentX        =   15161
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Descripción Ocurrencia"
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
            Begin Threed.SSPanel SSPanel12 
               Height          =   285
               Left            =   1230
               TabIndex        =   52
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "H. Ocurrencia"
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
            Begin Threed.SSPanel pnl_DesOcu 
               Height          =   315
               Left            =   1320
               TabIndex        =   53
               Top             =   1650
               Width           =   10005
               _Version        =   65536
               _ExtentX        =   17648
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Día: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
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
            Begin MSFlexGridLib.MSFlexGrid grd_LisExc 
               Height          =   855
               Left            =   -74970
               TabIndex        =   54
               Top             =   660
               Width           =   11325
               _ExtentX        =   19976
               _ExtentY        =   1508
               _Version        =   393216
               Rows            =   21
               Cols            =   6
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel15 
               Height          =   285
               Left            =   -74940
               TabIndex        =   55
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Excepción"
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
            Begin Threed.SSPanel SSPanel16 
               Height          =   285
               Left            =   -69330
               TabIndex        =   56
               Top             =   360
               Width           =   5325
               _Version        =   65536
               _ExtentX        =   9393
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Descripción Excepción"
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
            Begin Threed.SSPanel SSPanel19 
               Height          =   45
               Left            =   -74970
               TabIndex        =   57
               Top             =   1560
               Width           =   11325
               _Version        =   65536
               _ExtentX        =   19976
               _ExtentY        =   79
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
            Begin Threed.SSPanel pnl_DesExc 
               Height          =   315
               Left            =   -73770
               TabIndex        =   58
               Top             =   1650
               Width           =   10065
               _Version        =   65536
               _ExtentX        =   17754
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Día: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
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
            Begin Threed.SSPanel pnl_TipAut 
               Height          =   315
               Left            =   -73770
               TabIndex        =   59
               Top             =   2790
               Width           =   4005
               _Version        =   65536
               _ExtentX        =   7064
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "INGRESO A INSTANCIA"
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
            Begin Threed.SSPanel SSPanel20 
               Height          =   45
               Left            =   -74970
               TabIndex        =   60
               Top             =   1560
               Width           =   11325
               _Version        =   65536
               _ExtentX        =   19976
               _ExtentY        =   79
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
            Begin MSFlexGridLib.MSFlexGrid grd_LisCon 
               Height          =   855
               Left            =   -74970
               TabIndex        =   61
               Top             =   660
               Width           =   11325
               _ExtentX        =   19976
               _ExtentY        =   1508
               _Version        =   393216
               Rows            =   21
               Cols            =   4
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel21 
               Height          =   285
               Left            =   -74940
               TabIndex        =   62
               Top             =   360
               Width           =   2745
               _Version        =   65536
               _ExtentX        =   4842
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Instancia"
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
            Begin Threed.SSPanel SSPanel22 
               Height          =   285
               Left            =   -65610
               TabIndex        =   63
               Top             =   360
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
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
            Begin Threed.SSPanel SSPanel23 
               Height          =   285
               Left            =   -72210
               TabIndex        =   64
               Top             =   360
               Width           =   6615
               _Version        =   65536
               _ExtentX        =   11668
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Condiciones de Aprobación"
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
            Begin Threed.SSPanel pnl_InsCon 
               Height          =   315
               Left            =   -73680
               TabIndex        =   65
               Top             =   1650
               Width           =   10005
               _Version        =   65536
               _ExtentX        =   17648
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Día: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
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
            Begin Threed.SSPanel pnl_motivo 
               Height          =   315
               Left            =   -68970
               TabIndex        =   66
               Top             =   2790
               Width           =   5265
               _Version        =   65536
               _ExtentX        =   9287
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "MOTIVO DE EXCEPCION"
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
               Begin Threed.SSPanel SSPanel25 
                  Height          =   315
                  Left            =   6090
                  TabIndex        =   67
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   5205
                  _Version        =   65536
                  _ExtentX        =   9181
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "INGRESOS 4A CATEG. NO SE PUEDEN CONFIRMAR "
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
            End
            Begin Threed.SSPanel SSPanel26 
               Height          =   285
               Left            =   -74940
               TabIndex        =   78
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Excepción"
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
            Begin Threed.SSPanel SSPanel27 
               Height          =   285
               Left            =   -69330
               TabIndex        =   79
               Top             =   360
               Width           =   5325
               _Version        =   65536
               _ExtentX        =   9393
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Descripción Excepción"
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
            Begin Threed.SSPanel SSPanel28 
               Height          =   285
               Left            =   -73770
               TabIndex        =   80
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "H. Excepción"
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
            Begin Threed.SSPanel SSPanel29 
               Height          =   285
               Left            =   -72600
               TabIndex        =   81
               Top             =   360
               Width           =   3285
               _Version        =   65536
               _ExtentX        =   5794
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Instancia"
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
            Begin Threed.SSPanel SSPanel18 
               Height          =   285
               Left            =   -74940
               TabIndex        =   82
               Top             =   360
               Width           =   2745
               _Version        =   65536
               _ExtentX        =   4842
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Instancia"
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
            Begin Threed.SSPanel SSPanel17 
               Height          =   285
               Left            =   -65610
               TabIndex        =   83
               Top             =   360
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
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
            Begin Threed.SSPanel SSPanel30 
               Height          =   285
               Left            =   -72210
               TabIndex        =   84
               Top             =   360
               Width           =   6615
               _Version        =   65536
               _ExtentX        =   11668
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Condiciones de Aprobación"
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
            Begin VB.Label Label19 
               Caption         =   "Descripción:"
               Height          =   495
               Left            =   -74940
               TabIndex        =   88
               Top             =   1980
               Width           =   1155
            End
            Begin VB.Label Label18 
               Caption         =   "Excepción:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   87
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label17 
               Caption         =   "Autorizado por:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   86
               Top             =   2790
               Width           =   1095
            End
            Begin VB.Label Label16 
               Caption         =   "Levantamiento de Condiciones:"
               Height          =   495
               Left            =   -74940
               TabIndex        =   85
               Top             =   2580
               Width           =   1245
            End
            Begin VB.Label Label13 
               Caption         =   "Levantamiento de Condiciones:"
               Height          =   615
               Left            =   -74940
               TabIndex        =   77
               Top             =   2640
               Width           =   1215
            End
            Begin VB.Label Label14 
               Caption         =   "Instancia:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   76
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label15 
               Caption         =   "Condiciones de Aprobación:"
               Height          =   495
               Left            =   -74940
               TabIndex        =   75
               Top             =   1980
               Width           =   1155
            End
            Begin VB.Label Label12 
               Caption         =   "Descripción:"
               Height          =   495
               Left            =   -74940
               TabIndex        =   74
               Top             =   1980
               Width           =   1155
            End
            Begin VB.Label Label10 
               Caption         =   "Excepción:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   73
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label9 
               Caption         =   "Autorizado por:"
               Height          =   315
               Left            =   -74910
               TabIndex        =   72
               Top             =   2970
               Width           =   1095
            End
            Begin VB.Label Label11 
               Caption         =   "Descargo:"
               Height          =   315
               Left            =   60
               TabIndex        =   71
               Top             =   2640
               Width           =   1035
            End
            Begin VB.Label Label2 
               Caption         =   "Ocurrencia:"
               Height          =   315
               Left            =   60
               TabIndex        =   70
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label7 
               Caption         =   "Comentario u Observación:"
               Height          =   495
               Left            =   60
               TabIndex        =   69
               Top             =   1980
               Width           =   1155
            End
            Begin VB.Label lbl_motivo 
               Caption         =   "Motivo:"
               Height          =   255
               Left            =   -69630
               TabIndex        =   68
               Top             =   2850
               Width           =   645
            End
         End
      End
   End
End
Attribute VB_Name = "frm_GasAdm_11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_GasAdm()   As moddat_tpo_Genera
Dim l_int_AprCon     As Integer

Private Sub cmb_GasAdm_Click()
   Call gs_SetFocus(ipp_Import)
   If cmb_GasAdm.ListIndex > -1 Then
      ipp_Import.Value = l_arr_GasAdm(cmb_GasAdm.ListIndex + 1).Genera_Cantid
      pnl_MonGas.Caption = moddat_gf_Consulta_ParDes("204", Right(l_arr_GasAdm(cmb_GasAdm.ListIndex + 1).Genera_Codigo, 1))
   End If
End Sub

Private Sub cmb_GasAdm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_GasAdm_Click
   End If
End Sub

Private Sub cmd_BorIte_Click()
Dim r_str_CodGas  As String
   
   grd_GasAdm.Col = 3
   r_str_CodGas = grd_GasAdm.Text
   Call gs_RefrescaGrid(grd_GasAdm)
   
   If MsgBox("¿Está seguro de eliminar el Gasto Administrativo asignado?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
      Exit Sub
   End If
   
   g_str_Parame = "USP_BORRAR_TRA_GASADM ('" & moddat_g_str_NumSol & "', " & Left(r_str_CodGas, 2) & "), 1"
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If

   Screen.MousePointer = 11
   Call fs_Buscar_GasAdm
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_LimpiaItem
   Call fs_ActivaItem(False)
   Call gs_SetFocus(grd_GasAdm)
End Sub

Private Sub cmd_Comment_Click()
   frm_GasAdm_12.Show 1
End Sub

Private Sub cmd_EdiIte_Click()
Dim r_str_CodGas  As Integer
Dim r_int_Situac  As String
   
   grd_GasAdm.Col = 3
   r_str_CodGas = grd_GasAdm.Text
   Call gs_RefrescaGrid(grd_GasAdm)
   moddat_g_int_FlgGrb = 2
   
   g_str_Parame = "SELECT * FROM TRA_GASADM WHERE "
   g_str_Parame = g_str_Parame & "GASADM_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "GASADM_CODGAS = " & Left(r_str_CodGas, 2) & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
      
   g_rst_Princi.MoveFirst
   cmb_GasAdm.ListIndex = gf_Busca_Arregl(l_arr_GasAdm, Format(r_str_CodGas, "000")) - 1
   ipp_Import.Value = g_rst_Princi!GASADM_IMPORT
   r_int_Situac = g_rst_Princi!GASADM_SITUAC
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Valida si ya está pagado
   
   If r_int_Situac = 2 Then
      Call fs_ActivaItem(True)
   Else
       MsgBox "El Gasto Administrativo ya fue pagado.", vbExclamation, modgen_g_str_NomPlt
   End If

'   Call fs_ActivaItem(True)
   cmb_GasAdm.Enabled = False
   Call gs_SetFocus(ipp_Import)
End Sub

Private Sub cmd_Export_Click()
   If grd_GasAdm.Rows = 0 Then
      MsgBox "La solicitud no posee gastos administrativos.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Grabar_Click()
   If cmb_GasAdm.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Gasto Administrativo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_GasAdm)
      Exit Sub
   End If
   If ipp_Import.Value = 0 Then
      MsgBox "Debe ingresar el Importe del Gasto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Import)
      Exit Sub
   End If

   'Validar que el Gasto no este ingresado si es Agregar
   If moddat_g_int_FlgGrb = 1 Then
      g_str_Parame = "SELECT * FROM TRA_GASADM WHERE "
      g_str_Parame = g_str_Parame & "GASADM_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
      g_str_Parame = g_str_Parame & "GASADM_CODGAS = " & Left(l_arr_GasAdm(cmb_GasAdm.ListIndex + 1).Genera_Codigo, 2)
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         MsgBox "El Gasto Administrativo ya fue ingresado.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_GasAdm)
         Exit Sub
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      Call moddat_gs_FecSis
      
      g_str_Parame = "USP_TRA_GASADM_ASIGNA ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & Left(l_arr_GasAdm(cmb_GasAdm.ListIndex + 1).Genera_Codigo, 2) & ", "
      g_str_Parame = g_str_Parame & Right(l_arr_GasAdm(cmb_GasAdm.ListIndex + 1).Genera_Codigo, 1) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(ipp_Import.Text)) & ", "
      g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & "2, "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ", "
      g_str_Parame = g_str_Parame & "1)"
      
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
   Screen.MousePointer = 11
   Call fs_Buscar_GasAdm
   Screen.MousePointer = 0
   
   If moddat_g_int_FlgGrb = 1 Then
      Call fs_LimpiaItem
      Call fs_ActivaItem(True)
      Call gs_SetFocus(cmb_GasAdm)
   Else
      Call cmd_Cancel_Click
   End If
End Sub

Private Sub cmd_NueIte_Click()
   moddat_g_int_FlgGrb = 1
   Call fs_ActivaItem(True)
   Call gs_SetFocus(cmb_GasAdm)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
Dim r_arr_Mtz()      As moddat_g_tpo_DatCom

   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   pnl_Produc.Caption = moddat_g_str_NomPrd
   pnl_FecIng.Caption = moddat_g_str_FecIng
   pnl_IngIns.Caption = moddat_gf_FecIng_Ins(moddat_g_str_NumSol, 32)
   
   Call fs_Inicia
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
   Call modmip_gs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(0), 0)      'Buscar Información del Cliente
   Call modmip_gs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, grd_Listad(3), 1)      'Buscar Información del Cónyuge
   Call modmip_gs_DatInm(grd_Listad(1), True)
   
   'Datos del Crédito
   Call modmip_gs_DatCre(grd_Listad(2), r_arr_Mtz)
   moddat_g_str_CodEjeSeg = r_arr_Mtz(0).DatCom_EjeSeg
   moddat_g_str_CodConHip = r_arr_Mtz(0).DatCom_ConHip
   moddat_g_int_InsAct = r_arr_Mtz(0).DatCom_CodIns
   lbl_Moneda.Caption = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon))
   
   'Buscando Gastos Administrativos
   Call fs_ActivaItem(False)
   Call fs_Buscar_GasAdm
   Call fs_Buscar_LisOcu      'Buscando Ocurrencias de Instancia
   Call fs_Buscar_LisExc      'Buscando Excepciones
   Call fs_Buscar_LisCon      'Buscando Aprobaciones Condicionadas
   Call fs_Buscar_FinGci      'Buscando Financiacion de gastos de cierre
   
   'Si no hay Excepciones aplicadas
   If grd_LisExc.Rows = 0 Then
      tab_Seguim.TabVisible(1) = False
   End If

   'Si no hay Aprobaciones Condicionadas
   If grd_LisCon.Rows = 0 Then
      tab_Seguim.TabVisible(2) = False
   End If
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
Dim r_int_Contad     As Integer
   
   Call moddat_gs_Carga_ParSubPrd_Combo(cmb_GasAdm, l_arr_GasAdm(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "007", moddat_g_int_TipMon)
   
   grd_GasAdm.ColWidth(0) = 6150
   grd_GasAdm.ColWidth(1) = 2610
   grd_GasAdm.ColWidth(2) = 2280
   grd_GasAdm.ColWidth(3) = 0
   grd_GasAdm.ColAlignment(0) = flexAlignLeftCenter
   grd_GasAdm.ColAlignment(1) = flexAlignCenterCenter
   grd_GasAdm.ColAlignment(2) = flexAlignRightCenter
   
   'Inicializando Grid de Cliente y de Cónyuge
   For r_int_Contad = 0 To 3
      grd_Listad(r_int_Contad).ColWidth(0) = 3000
      grd_Listad(r_int_Contad).ColWidth(1) = 7940
      grd_Listad(r_int_Contad).ColAlignment(0) = flexAlignLeftCenter
      grd_Listad(r_int_Contad).ColAlignment(1) = flexAlignLeftCenter
      Call gs_LimpiaGrid(grd_Listad(r_int_Contad))
   Next r_int_Contad
   
   'Lista de Ocurrencias
   grd_LisOcu.ColWidth(0) = 1155
   grd_LisOcu.ColWidth(1) = 1185
   grd_LisOcu.ColWidth(2) = 8595
   grd_LisOcu.ColWidth(3) = 0
   grd_LisOcu.ColWidth(4) = 0
   grd_LisOcu.ColAlignment(0) = flexAlignCenterCenter
   grd_LisOcu.ColAlignment(1) = flexAlignCenterCenter
   grd_LisOcu.ColAlignment(2) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_LisOcu)
   
   pnl_Gastos.Visible = False
   pnl_DesOcu.Caption = ""
   txt_Observ.Text = ""
   txt_Descar.Text = ""

   'Lista de Excepciones
   grd_LisExc.ColWidth(0) = 1175
   grd_LisExc.ColWidth(1) = 1175
   grd_LisExc.ColWidth(2) = 3275
   grd_LisExc.ColWidth(3) = 5325
   grd_LisExc.ColWidth(4) = 0
   grd_LisExc.ColWidth(5) = 0
   grd_LisExc.ColAlignment(0) = flexAlignCenterCenter
   grd_LisExc.ColAlignment(1) = flexAlignCenterCenter
   grd_LisExc.ColAlignment(2) = flexAlignLeftCenter
   grd_LisExc.ColAlignment(3) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_LisExc)

   pnl_DesExc.Caption = ""
   txt_ObsExc.Text = ""
   pnl_TipAut.Caption = ""
   pnl_motivo.Caption = ""

   'Lista de Aprobaciones Condicionadas
   grd_LisCon.ColWidth(0) = 2735
   grd_LisCon.ColWidth(1) = 6605
   grd_LisCon.ColWidth(2) = 1625
   grd_LisCon.ColWidth(3) = 0
   grd_LisCon.ColAlignment(0) = flexAlignLeftCenter
   grd_LisCon.ColAlignment(1) = flexAlignLeftCenter
   grd_LisCon.ColAlignment(2) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_LisCon)

   pnl_InsCon.Caption = ""
   txt_ObsCon.Text = ""
   txt_LevCon.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If moddat_g_int_FlgAct = 2 Then
      'Registrar Asignación de Gastos de Cierre en Seguimiento
      If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, moddat_g_int_InsAct, 24, 0, "", 0, 0) Then
         Exit Sub
      End If
      
      'Enviar Correo de Asignación de Gastos de Cierre
      modgen_g_str_Mail_Asunto = "ASIGNACION DE GASTOS DE CIERRE (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
      modgen_g_str_Mail_Mensaj = ""
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
      
      Call fs_Envia_CorreoEle(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, "", 0, False, False, False)
   End If
End Sub

Private Sub grd_Listad_SelChange(Index As Integer)
   If grd_Listad(Index).Rows > 2 Then
      grd_Listad(Index).RowSel = grd_Listad(Index).Row
   End If
End Sub

Private Sub grd_GasAdm_DblClick()
   Call cmd_EdiIte_Click
End Sub

Private Sub grd_GasAdm_SelChange()
   If grd_GasAdm.Rows > 2 Then
      grd_GasAdm.RowSel = grd_GasAdm.Row
   End If
End Sub

Private Sub ipp_Import_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub fs_Buscar_GasAdm()
Dim r_dbl_TotGas     As Double

   r_dbl_TotGas = 0
   pnl_TotGas.Caption = "0.00 "
   cmd_NueIte.Enabled = True
   cmd_EdiIte.Enabled = False
   cmd_BorIte.Enabled = False
   grd_GasAdm.Enabled = False
   Call gs_LimpiaGrid(grd_GasAdm)
   
   g_str_Parame = "SELECT * FROM TRA_GASADM WHERE "
   g_str_Parame = g_str_Parame & "GASADM_NUMSOL = '" & moddat_g_str_NumSol & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   grd_GasAdm.Redraw = False
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_GasAdm.Rows = grd_GasAdm.Rows + 1
      grd_GasAdm.Row = grd_GasAdm.Rows - 1
      
      'Buscando Descripción de Gastos Administrativos
      grd_GasAdm.Col = 0
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "007", Format(g_rst_Princi!GASADM_CODGAS, "00") & CStr(g_rst_Princi!GASADM_TIPMON)) Then
         grd_GasAdm.Text = Trim(moddat_g_arr_Genera(1).Genera_Nombre)
      End If
      
      grd_GasAdm.Col = 1
      grd_GasAdm.Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!GASADM_TIPMON))
      
      grd_GasAdm.Col = 2
      grd_GasAdm.Text = Format(g_rst_Princi!GASADM_IMPORT, "###,###,##0.00")
      
      grd_GasAdm.Col = 3
      grd_GasAdm.Text = Format(g_rst_Princi!GASADM_CODGAS, "00") & Format(g_rst_Princi!GASADM_TIPMON, "0")
      
      r_dbl_TotGas = r_dbl_TotGas + g_rst_Princi!GASADM_IMPORT
      g_rst_Princi.MoveNext
   Loop
   
   pnl_TotGas.Caption = Format(r_dbl_TotGas, "###,###,##0.00") & " "
   grd_GasAdm.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If grd_GasAdm.Rows > 0 Then
      cmd_EdiIte.Enabled = True
      cmd_BorIte.Enabled = True
      grd_GasAdm.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_GasAdm)
   Call gs_SetFocus(grd_GasAdm)
End Sub

Private Sub fs_LimpiaItem()
   cmb_GasAdm.ListIndex = -1
   ipp_Import.Value = 0
   pnl_MonGas.Caption = ""
End Sub

Private Sub fs_ActivaItem(ByVal p_Habilita As Integer)
   cmb_GasAdm.Enabled = p_Habilita
   ipp_Import.Enabled = p_Habilita
   cmd_Grabar.Enabled = p_Habilita
   cmd_Cancel.Enabled = p_Habilita
   cmd_NueIte.Enabled = Not p_Habilita
   cmd_BorIte.Enabled = Not p_Habilita
   cmd_EdiIte.Enabled = Not p_Habilita
End Sub

Private Sub fs_Buscar_LisOcu()
   Call gs_LimpiaGrid(grd_LisOcu)
   moddat_g_int_NumObs = 0
   
   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = 32 " '62
   g_str_Parame = g_str_Parame & "ORDER BY SEGFECCRE DESC, SEGHORCRE DESC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   grd_LisOcu.Redraw = False
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_LisOcu.Rows = grd_LisOcu.Rows + 1
      grd_LisOcu.Row = grd_LisOcu.Rows - 1
      
      'Fecha de Ocurrencia
      grd_LisOcu.Col = 0
      grd_LisOcu.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECCRE))
      
      'Hora de Ocurrencia
      grd_LisOcu.Col = 1
      grd_LisOcu.Text = gf_FormatoHora(Format(g_rst_Princi!SEGHORCRE, "000000"))
      
      'Descripción Ocurrencia
      grd_LisOcu.Col = 2
      grd_LisOcu.Text = moddat_gf_Consulta_ParDes("004", Format(g_rst_Princi!SEGDET_CODOCU, "000000"))
      
      If g_rst_Princi!SEGDET_CODOCU = 21 Then
         If g_rst_Princi!SEGFECACT > 0 Then
            grd_LisOcu.Text = grd_LisOcu.Text & " (DESCARGO EFECTUADO - " & gf_FormatoFecha(CStr(g_rst_Princi!SEGFECACT))
            grd_LisOcu.Text = grd_LisOcu.Text & " / " & gf_FormatoHora(Format(g_rst_Princi!SEGHORACT, "000000")) & ")"
         Else
            moddat_g_int_NumObs = g_rst_Princi!SEGDET_NUMOBS
         End If
      End If
      
      grd_LisOcu.Col = 3
      grd_LisOcu.Text = Trim(g_rst_Princi!SEGDET_OBSERV & "")
      
      grd_LisOcu.Col = 4
      grd_LisOcu.Text = Trim(g_rst_Princi!SEGDET_OBSDES & "")
      
      g_rst_Princi.MoveNext
   Loop
   grd_LisOcu.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_LisOcu)
   Call grd_LisOcu_Click
End Sub

Private Sub fs_Buscar_LisExc()
Dim r_str_FecOcu  As String
   
   Call gs_LimpiaGrid(grd_LisExc)
   
   g_str_Parame = modgen_gf_Buscar_Excepc
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   grd_LisExc.Redraw = False
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_LisExc.Rows = grd_LisExc.Rows + 1
      grd_LisExc.Row = grd_LisExc.Rows - 1
      
      'Fecha de Excepción
      grd_LisExc.Col = 0
      grd_LisExc.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECCRE))
      
      'Hora de Excepción
      grd_LisExc.Col = 1
      grd_LisExc.Text = gf_FormatoHora(Format(g_rst_Princi!SEGHORCRE, "000000"))
      
      'Instancia
      grd_LisExc.Col = 2
      grd_LisExc.Text = moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SEGEXC_CODINS))
      
      'Descripción Excepción
      grd_LisExc.Col = 3
      grd_LisExc.Text = Trim(g_rst_Princi!SEGEXC_DESCRI & "")
      
      'Tipo Autorización
      grd_LisExc.Col = 4
      grd_LisExc.Text = moddat_gf_Consulta_ParDes("243", CStr(g_rst_Princi!SEGEXC_TIPAUT))
      
      'Motivo de Excepción
      grd_LisExc.Col = 5
      grd_LisExc.Text = Trim(g_rst_Princi!PARDES_DESCRI)
      
      g_rst_Princi.MoveNext
   Loop
   grd_LisExc.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_LisExc)
   Call grd_LisExc_Click
End Sub

Private Sub fs_Buscar_LisCon()
   l_int_AprCon = 0
   Call gs_LimpiaGrid(grd_LisCon)
   
   g_str_Parame = "SELECT * FROM TRA_SEGCON WHERE "
   g_str_Parame = g_str_Parame & "SEGCON_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & "ORDER BY SEGCON_SITUAC ASC, SEGCON_CODINS DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   grd_LisCon.Redraw = False
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_LisCon.Rows = grd_LisCon.Rows + 1
      grd_LisCon.Row = grd_LisCon.Rows - 1
      
      'Instancia
      grd_LisCon.Col = 0
      grd_LisCon.Text = moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SEGCON_CODINS))
      
      'Descripción Condiciones
      grd_LisCon.Col = 1
      grd_LisCon.Text = Trim(g_rst_Princi!SEGCON_OBSCON & "")
      
      'Situación
      grd_LisCon.Col = 2
      grd_LisCon.Text = moddat_gf_Consulta_ParDes("244", CStr(g_rst_Princi!SEGCON_SITUAC))
      
      If g_rst_Princi!SEGCON_SITUAC = 1 Then
         l_int_AprCon = 1
      End If
      
      'Descripción Levantamiento Condiciones
      grd_LisCon.Col = 3
      grd_LisCon.Text = Trim(g_rst_Princi!SEGCON_OBSLEV & "")
      
      g_rst_Princi.MoveNext
   Loop
   grd_LisCon.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_LisCon)
   Call grd_LisCon_Click
End Sub

Private Sub fs_Buscar_FinGci()
Dim r_str_Parame        As String
Dim r_rst_Princi        As ADODB.Recordset

   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "SELECT * FROM CRE_SOLMAE "
   r_str_Parame = r_str_Parame & " WHERE SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If r_rst_Princi.BOF And r_rst_Princi.EOF Then
      r_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   r_rst_Princi.MoveFirst
   If r_rst_Princi!SOLMAE_MTOGCI > 0 Then
      pnl_Gastos.Visible = True
   End If

   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
End Sub

Private Sub fs_GenExc()
Dim r_str_Parame     As String
Dim r_rst_Princi     As ADODB.Recordset
Dim r_obj_Excel      As Excel.Application
Dim r_int_NroFil     As Integer
Dim r_int_Conta      As Integer
Dim r_int_Fila       As Integer
Dim tipoInmueble     As String
Dim nomProyecto      As String
Dim valInmue         As Double
Dim valEstacio       As Double
Dim valDepo          As Double
Dim valTotInmue      As Double
Dim valPrest         As Double
Dim gasTasa          As Double
Dim gasNotari        As Double
Dim gasBloq          As Double
Dim gasRegMin        As Double
Dim gasRegHip        As Double
Dim itf              As Double
Dim Moneda           As String

   On Error GoTo ErrProc
   valTotInmue = 0
   gasTasa = 0
   gasNotari = 0
   gasBloq = 0
   gasRegMin = 0
   gasRegHip = 0
   itf = 0
   valInmue = 0
   valEstacio = 0
   valDepo = 0
   valPrest = 0
   tipoInmueble = ""
   nomProyecto = ""
   
   '******
   r_int_Fila = 0
   For r_int_Conta = 1 To grd_Listad(1).Rows
      grd_Listad(1).Row = r_int_Fila
      grd_Listad(1).Col = 0
      If (Trim(grd_Listad(1).Text) = "Tipo de Inmueble") Then
         grd_Listad(1).Col = 1
         tipoInmueble = grd_Listad(1).Text
      ElseIf (Trim(grd_Listad(1).Text) = "Nombre Proyecto") Or (Trim(grd_Listad(1).Text) = "Proyecto Vinculado") Or (Trim(grd_Listad(1).Text) = "Proyecto No Vinculado") Then
         grd_Listad(1).Col = 1
         nomProyecto = grd_Listad(1).Text
         Exit For
      End If
      r_int_Fila = r_int_Fila + 1
   Next
   
   '******
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "SELECT * FROM CRE_SOLMAE "
   r_str_Parame = r_str_Parame & " WHERE SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If r_rst_Princi.BOF And r_rst_Princi.EOF Then
      r_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   r_rst_Princi.MoveFirst
   valInmue = r_rst_Princi!SOLMAE_MTOINM
   valEstacio = r_rst_Princi!SOLMAE_MTOEST
   valDepo = 0
   valPrest = r_rst_Princi!SOLMAE_MTOPRE_SOL
   valTotInmue = valInmue + valEstacio + valDepo

   r_rst_Princi.Close
   Set r_rst_Princi = Nothing

   r_int_Fila = 0
   grd_GasAdm.Row = r_int_Fila
   grd_GasAdm.Col = 1
   If (Trim(grd_GasAdm.Text) = "SOLES") Then
      Moneda = "S/."
   Else
      Moneda = "$"
   End If
   For r_int_Conta = 1 To grd_GasAdm.Rows
      grd_GasAdm.Row = r_int_Fila
      grd_GasAdm.Col = 0
      If (Trim(grd_GasAdm.Text) = "GASTOS DE TASACION") Then
         grd_GasAdm.Col = 2
         gasTasa = CDbl(grd_GasAdm.Text)
      ElseIf (Trim(grd_GasAdm.Text) = "GASTOS NOTARIALES") Then
         grd_GasAdm.Col = 2
         gasNotari = CDbl(grd_GasAdm.Text)
      ElseIf (Trim(grd_GasAdm.Text) = "ITF") Then
         grd_GasAdm.Col = 2
         itf = CDbl(grd_GasAdm.Text)
      ElseIf (Trim(grd_GasAdm.Text) = "GASTOS REGISTRALES - BLOQUEO REGISTRAL") Then
         grd_GasAdm.Col = 2
         gasBloq = CDbl(grd_GasAdm.Text)
      ElseIf (Trim(grd_GasAdm.Text) = "GASTOS REGISTRALES - MINUTA COMPRA VENTA") Then
         grd_GasAdm.Col = 2
         gasRegMin = CDbl(grd_GasAdm.Text)
      ElseIf (Trim(grd_GasAdm.Text) = "GASTOS REGISTRALES - INSCRIPCION DE GARANTIA") Then
         grd_GasAdm.Col = 2
         gasRegHip = CDbl(grd_GasAdm.Text)
      End If
      r_int_Fila = r_int_Fila + 1
   Next
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Columns("A").HorizontalAlignment = xlHAlignLeft
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      
      r_int_NroFil = 1
      
      .Cells(r_int_NroFil, 1) = UCase(moddat_g_str_NomPrd) & " - " & UCase(nomProyecto)
      .Cells(r_int_NroFil, 1).Font.Bold = True
      .Cells(r_int_NroFil, 1).Font.Size = 12
      .Cells(r_int_NroFil, 1).Font.Color = RGB(255, 255, 255)
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 3)).Merge
      .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignCenter
      
      r_int_NroFil = r_int_NroFil + 2
      .Cells(r_int_NroFil, 1) = "Tipo de Inmueble"
      .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignLeft
      .Cells(r_int_NroFil, 1).Font.Bold = True
      .Cells(r_int_NroFil, 2) = tipoInmueble
      .Cells(r_int_NroFil, 2).Font.Bold = True
      
      r_int_NroFil = r_int_NroFil + 2
      .Cells(r_int_NroFil, 1) = "Valor del inmueble en " & Moneda
      .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignLeft
      .Cells(r_int_NroFil, 1).Font.Bold = True
      .Cells(r_int_NroFil, 2) = valInmue
      .Cells(r_int_NroFil, 2).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
      .Cells(r_int_NroFil, 2).Font.Bold = True
      
      r_int_NroFil = r_int_NroFil + 2
      .Cells(r_int_NroFil, 1) = "Valor del estacionamiento en " & Moneda
      .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignLeft
      .Cells(r_int_NroFil, 1).Font.Bold = True
      .Cells(r_int_NroFil, 2) = valEstacio
      .Cells(r_int_NroFil, 2).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
      .Cells(r_int_NroFil, 2).Font.Bold = True
      
      r_int_NroFil = r_int_NroFil + 2
      .Cells(r_int_NroFil, 1) = "Valor del deposito en " & Moneda
      .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignLeft
      .Cells(r_int_NroFil, 1).Font.Bold = True
      .Cells(r_int_NroFil, 2) = valDepo
      .Cells(r_int_NroFil, 2).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
      .Cells(r_int_NroFil, 2).Font.Bold = True
      
      r_int_NroFil = r_int_NroFil + 2
      .Cells(r_int_NroFil, 1) = "Valor total del inmueble en " & Moneda
      .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignLeft
      .Cells(r_int_NroFil, 1).Font.Bold = True
      .Cells(r_int_NroFil, 2) = valTotInmue
      .Cells(r_int_NroFil, 2).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
      .Cells(r_int_NroFil, 2).Font.Bold = True
      
      r_int_NroFil = r_int_NroFil + 2
      .Cells(r_int_NroFil, 1) = "Valor del prestamo en " & Moneda
      .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignLeft
      .Cells(r_int_NroFil, 1).Font.Bold = True
      .Cells(r_int_NroFil, 2) = valPrest
      .Cells(r_int_NroFil, 2).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
      .Cells(r_int_NroFil, 2).Font.Bold = True
      
      r_int_NroFil = r_int_NroFil + 2
      .Cells(r_int_NroFil, 1) = "GASTOS DE CIERRE EN " & Moneda
      .Cells(r_int_NroFil, 1).Font.Color = RGB(255, 255, 255)
      .Cells(r_int_NroFil, 1).Font.Size = 12
      .Cells(r_int_NroFil, 1).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 3)).Merge
      .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignCenter
      
      If (gasTasa > 0) Then
         r_int_NroFil = r_int_NroFil + 1
         .Cells(r_int_NroFil, 1) = "Gastos de Tasación"
         .Cells(r_int_NroFil, 1).Font.Bold = True
         .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignLeft
         .Cells(r_int_NroFil, 3) = gasTasa
         .Cells(r_int_NroFil, 3).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
      End If
      
      If (gasNotari > 0) Then
         r_int_NroFil = r_int_NroFil + 1
         .Cells(r_int_NroFil, 1) = "Gastos Notariales"
         .Cells(r_int_NroFil, 1).Font.Bold = True
         .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignLeft
         .Cells(r_int_NroFil, 3) = gasNotari
         .Cells(r_int_NroFil, 3).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
      End If
      
      If (gasBloq > 0) Then
         r_int_NroFil = r_int_NroFil + 1
         .Cells(r_int_NroFil, 1) = "Bloqueo registral"
         .Cells(r_int_NroFil, 1).Font.Bold = True
         .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignLeft
         .Cells(r_int_NroFil, 3) = gasBloq
         .Cells(r_int_NroFil, 3).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
      End If
      
      If (gasRegMin > 0) Then
         r_int_NroFil = r_int_NroFil + 1
         .Cells(r_int_NroFil, 1) = "Gastos por registro de minuta"
         .Cells(r_int_NroFil, 1).Font.Bold = True
         .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignLeft
         .Cells(r_int_NroFil, 3) = gasRegMin
         .Cells(r_int_NroFil, 3).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
      End If
      
      If (gasRegHip > 0) Then
         r_int_NroFil = r_int_NroFil + 1
         .Cells(r_int_NroFil, 1) = "Gastos por registro de hipoteca"
         .Cells(r_int_NroFil, 1).Font.Bold = True
         .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignLeft
         .Cells(r_int_NroFil, 3) = gasRegHip
         .Cells(r_int_NroFil, 3).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
      End If
      
      If (itf > 0) Then
         r_int_NroFil = r_int_NroFil + 1
         .Cells(r_int_NroFil, 1) = "ITF"
         .Cells(r_int_NroFil, 1).Font.Bold = True
         .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignLeft
         .Cells(r_int_NroFil, 3) = itf
         .Cells(r_int_NroFil, 3).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
      End If
      
      r_int_NroFil = r_int_NroFil + 1
      .Cells(r_int_NroFil, 1) = "TOTAL GASTOS DE CIERRE EN " & Moneda
      .Cells(r_int_NroFil, 1).Font.Color = RGB(255, 255, 255)
      .Cells(r_int_NroFil, 1).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 2)).Merge
      .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 2)).Interior.Color = RGB(51, 153, 102)
      .Cells(r_int_NroFil, 3) = gasTasa + gasNotari + gasRegMin + gasRegHip + itf + gasBloq
      .Cells(r_int_NroFil, 3).Font.Bold = True
      .Cells(r_int_NroFil, 3).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 3)).Borders(xlEdgeTop).LineStyle = xlContinuous
      
      .Columns("A").ColumnWidth = 50.71
      .Columns("B").ColumnWidth = 33.43
      .Columns("C").ColumnWidth = 17.57
      
      'TITUTLO
      .Range(.Cells(1, 1), .Cells(1, 3)).Interior.Color = RGB(51, 153, 102)
      .Range(.Cells(15, 1), .Cells(15, 3)).Interior.Color = RGB(51, 153, 102)
      
      .Range(.Cells(15, 1), .Cells(15, 3)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(16, 1), .Cells(16, 3)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Cells(15, 3).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(16, 3), .Cells(r_int_NroFil, 3)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(16, 3), .Cells(r_int_NroFil, 3)).Borders(xlEdgeRight).LineStyle = xlContinuous
      
      .Range(.Cells(r_int_NroFil + 1, 1), .Cells(r_int_NroFil + 1, 3)).Borders(xlEdgeTop).LineStyle = xlContinuous
   End With
   
   r_obj_Excel.Sheets(1).Name = "Hoja1"
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
   Exit Sub
   
ErrProc:
   MsgBox Err.Description
   Resume Next
End Sub

Private Sub grd_LisOcu_Click()
Dim r_str_FecOcu     As String
Dim r_str_HorOcu     As String
Dim r_str_DesOcu     As String

   If grd_LisOcu.Rows > 0 Then
      grd_LisOcu.Col = 0
      r_str_FecOcu = grd_LisOcu.Text
      
      grd_LisOcu.Col = 1
      r_str_HorOcu = grd_LisOcu.Text
      
      grd_LisOcu.Col = 2
      r_str_DesOcu = grd_LisOcu.Text
      
      pnl_DesOcu.Caption = "Día: " & r_str_FecOcu & " - " & r_str_HorOcu & " hrs. - " & r_str_DesOcu
   
      grd_LisOcu.Col = 3
      txt_Observ.Text = grd_LisOcu.Text
      
      grd_LisOcu.Col = 4
      txt_Descar.Text = grd_LisOcu.Text
      
      Call gs_RefrescaGrid(grd_LisOcu)
   End If
End Sub

Private Sub grd_LisOcu_SelChange()
   If grd_LisOcu.Rows > 2 Then
      grd_LisOcu.RowSel = grd_LisOcu.Row
   End If
   
'   Call grd_LisOcu_Click
End Sub

Private Sub grd_LisExc_Click()
Dim r_str_FecExc     As String
Dim r_str_HorExc     As String
Dim r_str_InsExc     As String

   If grd_LisExc.Rows > 0 Then
      grd_LisExc.Col = 0
      r_str_FecExc = grd_LisExc.Text
      
      grd_LisExc.Col = 1
      r_str_HorExc = grd_LisExc.Text
      
      grd_LisExc.Col = 2
      r_str_InsExc = grd_LisExc.Text
      
      pnl_DesExc.Caption = "Día: " & r_str_FecExc & " - " & r_str_HorExc & " hrs. - " & r_str_InsExc
   
      grd_LisExc.Col = 3
      txt_ObsExc.Text = grd_LisExc.Text
      
      grd_LisExc.Col = 4
      pnl_TipAut.Caption = grd_LisExc.Text
      
      If LCase(Trim(r_str_InsExc)) = LCase("EVALUACION CREDITICIA") Then
         grd_LisExc.Col = 5
         pnl_motivo.Caption = IIf(grd_LisExc.Text = "0", " ", grd_LisExc.Text)
         pnl_motivo.Visible = True
         lbl_motivo.Visible = True
      Else
         pnl_motivo.Visible = False
         lbl_motivo.Visible = False
         pnl_motivo.Caption = ""
      End If
      
      Call gs_SetFocus(grd_LisExc)
      Call gs_RefrescaGrid(grd_LisExc)
   Else
      pnl_DesExc.Caption = ""
      txt_ObsExc.Text = ""
      pnl_TipAut.Caption = ""
      pnl_motivo.Caption = ""
   End If
End Sub

Private Sub grd_LisExc_SelChange()
   If grd_LisExc.Rows > 2 Then
      grd_LisExc.RowSel = grd_LisExc.Row
   End If
   
'   Call grd_LisExc_Click
End Sub

Private Sub grd_LisCon_Click()
Dim r_str_FecOcu     As String
Dim r_str_HorOcu     As String
Dim r_str_DesOcu     As String

   If grd_LisCon.Rows > 0 Then
      grd_LisCon.Col = 0
      pnl_InsCon.Caption = grd_LisCon.Text
   
      grd_LisCon.Col = 1
      txt_ObsCon.Text = grd_LisCon.Text
      
      grd_LisCon.Col = 3
      txt_LevCon.Text = grd_LisCon.Text
      
      Call gs_RefrescaGrid(grd_LisCon)
   End If
End Sub

Private Sub grd_LisCon_SelChange()
   If grd_LisCon.Rows > 2 Then
      grd_LisCon.RowSel = grd_LisCon.Row
   End If
   
'   Call grd_LisCon_Click
End Sub

Private Sub txt_Descar_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_LevCon_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_ObsCon_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_ObsExc_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub
