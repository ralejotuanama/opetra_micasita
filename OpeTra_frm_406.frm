VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_EmpPer_08 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11355
   Icon            =   "OpeTra_frm_406.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   8925
      Left            =   -30
      TabIndex        =   17
      Top             =   0
      Width           =   11445
      _Version        =   65536
      _ExtentX        =   20188
      _ExtentY        =   15743
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
         Height          =   525
         Left            =   75
         TabIndex        =   18
         Top             =   1470
         Width           =   11265
         _Version        =   65536
         _ExtentX        =   19870
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
            Left            =   1260
            TabIndex        =   19
            Top             =   60
            Width           =   9915
            _Version        =   65536
            _ExtentX        =   17489
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
            AutoSize        =   -1  'True
            Caption         =   "Empresa:"
            Height          =   195
            Left            =   330
            TabIndex        =   20
            Top             =   150
            Width           =   660
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   645
         Left            =   75
         TabIndex        =   21
         Top             =   780
         Width           =   11265
         _Version        =   65536
         _ExtentX        =   19870
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_406.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Imprimir Listado"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10650
            Picture         =   "OpeTra_frm_406.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   675
         Left            =   75
         TabIndex        =   22
         Top             =   60
         Width           =   11265
         _Version        =   65536
         _ExtentX        =   19870
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
            TabIndex        =   23
            Top             =   210
            Width           =   3105
            _Version        =   65536
            _ExtentX        =   5477
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Datos de la Empresa Peritaje"
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
            Left            =   10800
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
            Picture         =   "OpeTra_frm_406.frx":0890
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   1605
         Left            =   75
         TabIndex        =   24
         Top             =   2040
         Width           =   11265
         _Version        =   65536
         _ExtentX        =   19870
         _ExtentY        =   2831
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
         Begin VB.TextBox txt_DirEle5 
            Height          =   315
            Left            =   1260
            MaxLength       =   50
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   1140
            Width           =   4140
         End
         Begin VB.TextBox txt_DirEle4 
            Height          =   315
            Left            =   6690
            MaxLength       =   50
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   810
            Width           =   4140
         End
         Begin VB.TextBox txt_DirEle3 
            Height          =   315
            Left            =   1260
            MaxLength       =   50
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   810
            Width           =   4140
         End
         Begin VB.TextBox txt_DirEle1 
            Height          =   315
            Left            =   1260
            MaxLength       =   50
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   480
            Width           =   4140
         End
         Begin VB.TextBox txt_DirEle2 
            Height          =   315
            Left            =   6690
            MaxLength       =   50
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   480
            Width           =   4140
         End
         Begin EditLib.fpDoubleSingle txt_Importe1 
            Height          =   315
            Left            =   1260
            TabIndex        =   0
            Top             =   150
            Width           =   1380
            _Version        =   196608
            _ExtentX        =   2434
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
            MaxValue        =   "90000"
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
            Caption         =   "Correo 5:"
            Height          =   195
            Left            =   330
            TabIndex        =   30
            Top             =   1200
            Width           =   645
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Correo 4:"
            Height          =   195
            Left            =   5820
            TabIndex        =   29
            Top             =   870
            Width           =   645
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Correo 3:"
            Height          =   195
            Left            =   330
            TabIndex        =   28
            Top             =   870
            Width           =   645
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Left            =   330
            TabIndex        =   27
            Top             =   210
            Width           =   570
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Correo 1:"
            Height          =   195
            Left            =   330
            TabIndex        =   26
            Top             =   540
            Width           =   645
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Correo 2:"
            Height          =   195
            Left            =   5820
            TabIndex        =   25
            Top             =   540
            Width           =   645
         End
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   5070
         Left            =   75
         TabIndex        =   31
         Top             =   3690
         Width           =   11265
         _Version        =   65536
         _ExtentX        =   19870
         _ExtentY        =   8943
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
         Begin Threed.SSPanel pnl_Boton_Cta 
            Height          =   675
            Left            =   30
            TabIndex        =   32
            Top             =   2925
            Width           =   11190
            _Version        =   65536
            _ExtentX        =   19738
            _ExtentY        =   1199
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
               Left            =   10545
               Picture         =   "OpeTra_frm_406.frx":0B9A
               Style           =   1  'Graphical
               TabIndex        =   8
               ToolTipText     =   "Editar Cuenta"
               Top             =   60
               Width           =   585
            End
            Begin VB.CommandButton cmd_Borrar 
               Height          =   570
               Left            =   9945
               Picture         =   "OpeTra_frm_406.frx":0EA4
               Style           =   1  'Graphical
               TabIndex        =   7
               ToolTipText     =   "Eliminar Cuenta"
               Top             =   60
               Width           =   585
            End
            Begin VB.CommandButton cmd_Nuevo 
               Height          =   570
               Left            =   9345
               Picture         =   "OpeTra_frm_406.frx":11AE
               Style           =   1  'Graphical
               TabIndex        =   6
               ToolTipText     =   "Nueva Cuenta"
               Top             =   60
               Width           =   585
            End
         End
         Begin Threed.SSPanel pnl_Dato_Cta 
            Height          =   1425
            Left            =   30
            TabIndex        =   33
            Top             =   3600
            Width           =   11190
            _Version        =   65536
            _ExtentX        =   19738
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
            BorderWidth     =   1
            BevelOuter      =   0
            BevelInner      =   1
            Begin VB.ComboBox cmb_Proyecto 
               Height          =   315
               Left            =   1260
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   570
               Width           =   7260
            End
            Begin VB.ComboBox cmb_Producto 
               Height          =   315
               Left            =   1260
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   240
               Width           =   7260
            End
            Begin VB.CommandButton cmd_Aceptar 
               Height          =   585
               Left            =   9960
               Picture         =   "OpeTra_frm_406.frx":14B8
               Style           =   1  'Graphical
               TabIndex        =   13
               Tag             =   "0"
               ToolTipText     =   "Agregar Cuenta"
               Top             =   780
               Width           =   585
            End
            Begin VB.CommandButton cmd_Cancelar 
               Height          =   585
               Left            =   10560
               Picture         =   "OpeTra_frm_406.frx":17C2
               Style           =   1  'Graphical
               TabIndex        =   14
               ToolTipText     =   "Cancelar"
               Top             =   780
               Width           =   585
            End
            Begin VB.ComboBox cmb_Moneda 
               Height          =   315
               Left            =   1260
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   900
               Width           =   2400
            End
            Begin EditLib.fpDoubleSingle txt_Importe2 
               Height          =   315
               Left            =   7140
               TabIndex        =   12
               Top             =   900
               Width           =   1380
               _Version        =   196608
               _ExtentX        =   2434
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
               MaxValue        =   "90000"
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
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Proyecto:"
               Height          =   195
               Left            =   330
               TabIndex        =   37
               Top             =   630
               Width           =   675
            End
            Begin VB.Label Label81 
               AutoSize        =   -1  'True
               Caption         =   "Producto:"
               Height          =   195
               Left            =   330
               TabIndex        =   36
               Top             =   300
               Width           =   690
            End
            Begin VB.Label Label84 
               AutoSize        =   -1  'True
               Caption         =   "Moneda:"
               Height          =   195
               Left            =   330
               TabIndex        =   35
               Top             =   945
               Width           =   630
            End
            Begin VB.Label Label86 
               AutoSize        =   -1  'True
               Caption         =   "Importe:"
               Height          =   195
               Left            =   6450
               TabIndex        =   34
               Top             =   945
               Width           =   570
            End
         End
         Begin Threed.SSPanel pnl_Producto 
            Height          =   285
            Left            =   90
            TabIndex        =   38
            Top             =   90
            Width           =   3315
            _Version        =   65536
            _ExtentX        =   5847
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
         Begin Threed.SSPanel pnl_Proyecto 
            Height          =   285
            Left            =   3390
            TabIndex        =   39
            Top             =   90
            Width           =   4485
            _Version        =   65536
            _ExtentX        =   7911
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Proyecto"
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
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   285
            Left            =   7860
            TabIndex        =   40
            Top             =   90
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
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
         Begin Threed.SSPanel pnl_Importe 
            Height          =   285
            Left            =   9510
            TabIndex        =   41
            Top             =   90
            Width           =   1370
            _Version        =   65536
            _ExtentX        =   2417
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   2505
            Left            =   75
            TabIndex        =   42
            Top             =   390
            Width           =   11130
            _ExtentX        =   19632
            _ExtentY        =   4419
            _Version        =   393216
            Rows            =   30
            Cols            =   10
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
      End
   End
End
Attribute VB_Name = "frm_EmpPer_08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_str_Codigo      As String
Dim l_arr_Produc()    As moddat_tpo_Genera
Dim l_arr_Proyec()    As moddat_tpo_Genera

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
    If CDbl(txt_Importe2.Text) = 0 Then
       MsgBox "Debe de ingresar un importe", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(txt_Importe2)
       Exit Sub
    End If
   
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
          grd_Listad.Text = cmb_Producto.ItemData(cmb_Producto.ListIndex) ' l_arr_Produc(cmb_Producto.ListIndex + 1).Genera_Codigo
      
          grd_Listad.Col = 2
          grd_Listad.Text = Trim(cmb_Producto.Text)
          
          grd_Listad.Col = 3
          grd_Listad.Text = cmb_Proyecto.ItemData(cmb_Proyecto.ListIndex) 'l_arr_Proyec(cmb_Proyecto.ListIndex + 1).Genera_Codigo
          
          grd_Listad.Col = 4
          grd_Listad.Text = Trim(cmb_Proyecto.Text)
          
          grd_Listad.Col = 5
          grd_Listad.Text = cmb_Moneda.ItemData(cmb_Moneda.ListIndex)
          
          grd_Listad.Col = 6
          grd_Listad.Text = Trim(cmb_Moneda.Text)
          
          grd_Listad.Col = 7
          grd_Listad.Text = txt_Importe2.Text & " "
          
          grd_Listad.Col = 8
          grd_Listad.Text = 1
          
          grd_Listad.Col = 9
          grd_Listad.Text = 1
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
       grd_Listad.TextMatrix(grd_Listad.Row, 7) = txt_Importe2.Text & " "
       
       If Trim(grd_Listad.TextMatrix(grd_Listad.Row, 0)) = "" Then
          grd_Listad.TextMatrix(grd_Listad.Row, 9) = 1 'INSERT
       Else
          grd_Listad.TextMatrix(grd_Listad.Row, 9) = 2 'UPDATE
       End If
    End If
    
    Call cmd_Cancelar_Click
End Sub

Private Sub cmd_Borrar_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   If MsgBox("¿Está seguro que desea eliminar?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   If Trim(grd_Listad.TextMatrix(grd_Listad.Row, 0)) = "" Then
       grd_Listad.RemoveItem (grd_Listad.Row)
   Else
       grd_Listad.TextMatrix(grd_Listad.Row, 8) = 0
       grd_Listad.RowHeight(grd_Listad.Row) = 0
       
       If Trim(grd_Listad.TextMatrix(grd_Listad.Row, 0)) = "" Then
          grd_Listad.TextMatrix(grd_Listad.Row, 9) = 1 'INSERT
       Else
          grd_Listad.TextMatrix(grd_Listad.Row, 9) = 2 'UPDATE
       End If
   End If
End Sub

Private Sub cmd_Cancelar_Click()
   cmb_Producto.ListIndex = -1
   cmb_Proyecto.ListIndex = -1
   cmb_Moneda.ListIndex = -1
   txt_Importe2.Text = "0.00"
   
   cmb_Producto.Enabled = False
   cmb_Proyecto.Enabled = False
   cmb_Moneda.Enabled = False
   txt_Importe2.Enabled = False
   
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
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   Call gs_BuscarCombo_Item(cmb_Producto, grd_Listad.TextMatrix(grd_Listad.Row, 1))
   Call gs_BuscarCombo_Item(cmb_Proyecto, grd_Listad.TextMatrix(grd_Listad.Row, 3))
   Call gs_BuscarCombo_Item(cmb_Moneda, grd_Listad.TextMatrix(grd_Listad.Row, 5))
   txt_Importe2.Text = grd_Listad.TextMatrix(grd_Listad.Row, 7)
   
   cmb_Producto.Enabled = True
   cmb_Proyecto.Enabled = True
   cmb_Moneda.Enabled = True
   txt_Importe2.Enabled = True
   
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
   txt_Importe2.Text = "0.00"
   
   cmb_Producto.Enabled = True
   cmb_Proyecto.Enabled = True
   cmb_Moneda.Enabled = True
   txt_Importe2.Enabled = True
   
   cmd_Aceptar.Enabled = True
   cmd_Cancelar.Enabled = True
   cmd_Nuevo.Enabled = False
   cmd_Borrar.Enabled = False
   cmd_Editar.Enabled = False
   
   grd_Listad.Enabled = False
   Call gs_SetFocus(cmb_Producto)
   cmd_Aceptar.Tag = 1
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Call gs_CentraForm(Me)
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Buscar
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
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
   
   'Call moddat_gs_Carga_Produc_Comerc(cmb_Producto, l_arr_Produc, 4)
   'Call moddat_gs_Carga_Proyec(cmb_Proyecto, l_arr_Proyec)
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
   '-----------------------------
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
   '-----------------------------
   Call cmd_Cancelar_Click
   Call gs_LimpiaGrid(grd_Listad)
   
   grd_Listad.ColWidth(0) = 0    'ITEM
   grd_Listad.ColWidth(1) = 0    'CODIGO_PRODUCTO
   grd_Listad.ColWidth(2) = 3290 'NOMBRE PRODUCTO
   grd_Listad.ColWidth(3) = 0    'CODIGO_PROPYECTO
   grd_Listad.ColWidth(4) = 4470 'NOMBRE PROYECTO
   grd_Listad.ColWidth(5) = 0    'CODIGO_MONEDA
   grd_Listad.ColWidth(6) = 1655 'MONEDA
   grd_Listad.ColWidth(7) = 1365 'IMPORTE
   grd_Listad.ColWidth(8) = 0    'ESTADO
   grd_Listad.ColWidth(9) = 0    'INPUT
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter
   grd_Listad.ColAlignment(6) = flexAlignLeftCenter
   grd_Listad.ColAlignment(7) = flexAlignRightCenter
End Sub

Private Sub fs_Buscar()
Dim r_str_Parame    As String
Dim r_rst_Princi    As ADODB.Recordset

   l_str_Codigo = ""
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT DATEMP_CODEMP,  DATEMP_IMPORT,  DATEMP_DIRELE1, "
   r_str_Parame = r_str_Parame & "        DATEMP_DIRELE2, DATEMP_DIRELE3, DATEMP_DIRELE4, DATEMP_DIRELE5 "
   r_str_Parame = r_str_Parame & "   FROM MNT_DATEMP "
   r_str_Parame = r_str_Parame & "  WHERE DATEMP_CODEMP = '" & moddat_g_str_CodGrp & "' "
   r_str_Parame = r_str_Parame & "    AND DATEMP_TIPTAB = 1 "

   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      r_rst_Princi.MoveFirst
      l_str_Codigo = r_rst_Princi!DATEMP_CODEMP
      txt_Importe1.Text = r_rst_Princi!DATEMP_IMPORT
      txt_DirEle1.Text = Trim(r_rst_Princi!DATEMP_DIRELE1 & "")
      txt_DirEle2.Text = Trim(r_rst_Princi!DATEMP_DIRELE2 & "")
      txt_DirEle3.Text = Trim(r_rst_Princi!DATEMP_DIRELE3 & "")
      txt_DirEle4.Text = Trim(r_rst_Princi!DATEMP_DIRELE4 & "")
      txt_DirEle5.Text = Trim(r_rst_Princi!DATEMP_DIRELE5 & "")
   End If
      
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   '------------------------------------------------------------
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT GASPAR_CODGAS, GASPAR_CODEMP, GASPAR_TIPTAB, GASPAR_CODPRD, TRIM(C.PRODUC_DESCRI) AS NOMBRE_PRODUCTO, "
   r_str_Parame = r_str_Parame & "        GASPAR_CODPRY, TRIM(B.DATGEN_TITULO) AS NOMBRE_PROYECTO   , GASPAR_CODMON, GASPAR_GASTAS_MTO, "
   r_str_Parame = r_str_Parame & "        TRIM(D.PARDES_DESCRI) As MONEDA, GASPAR_SITUAC "
   r_str_Parame = r_str_Parame & "   FROM TRA_GASPAR A "
   r_str_Parame = r_str_Parame & "  INNER JOIN PRY_DATGEN B ON B.DATGEN_CODIGO = A.GASPAR_CODPRY "
   r_str_Parame = r_str_Parame & "  INNER JOIN CRE_PRODUC C ON C.PRODUC_CODIGO = A.GASPAR_CODPRD "
   r_str_Parame = r_str_Parame & "  INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 204 AND D.PARDES_CODITE = A.GASPAR_CODMON "
   r_str_Parame = r_str_Parame & "  WHERE GASPAR_CODEMP = '" & moddat_g_str_CodGrp & "' "
   r_str_Parame = r_str_Parame & "    AND GASPAR_TIPTAB = 1 "
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
         
         grd_Listad.Col = 7
         grd_Listad.Text = Format(r_rst_Princi!GASPAR_GASTAS_MTO, "###,###,##0.00") & " "
         
         grd_Listad.Col = 8
         grd_Listad.Text = r_rst_Princi!GASPAR_SITUAC
         
         grd_Listad.Col = 9
         grd_Listad.Text = 0
                  
         r_rst_Princi.MoveNext
      Loop
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
      MsgBox "Debe de seleccionar una empresa.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Importe1)
      Exit Sub
   End If

   If Len(Trim(pnl_EmpPer.Caption)) = 0 Then
      MsgBox "Debe de seleccionar una empresa.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Importe1)
      Exit Sub
   End If
   
   If CDbl(txt_Importe1.Text) = 0 Then
      MsgBox "Debe se ingresar un importe.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Importe1)
      Exit Sub
   End If
   
   If Len(Trim(txt_DirEle1.Text)) = 0 Then
      MsgBox "Debe se ingresar un correo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_DirEle1)
      Exit Sub
   Else
      If gf_ValidarEmail(txt_DirEle1) = False Then
         MsgBox "El correo ingresado es erroneo.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_DirEle1)
         Exit Sub
      End If
   End If
   
   If Len(Trim(txt_DirEle2.Text)) > 0 Then
      If gf_ValidarEmail(txt_DirEle2) = False Then
         MsgBox "El correo ingresado es erroneo.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_DirEle2)
         Exit Sub
      End If
   End If
'----------------------------
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
   r_str_Parame = r_str_Parame & "1," 'Tipo Tabla
   r_str_Parame = r_str_Parame & CDbl(txt_Importe1.Text) & ","
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
   
   For r_int_Fila = 0 To grd_Listad.Rows - 1
       If grd_Listad.TextMatrix(r_int_Fila, 9) = 1 Or grd_Listad.TextMatrix(r_int_Fila, 9) = 2 Then
          r_str_Parame = ""
          r_str_Parame = r_str_Parame & " USP_TRA_GASPAR ( "
          r_str_Parame = r_str_Parame & "'" & grd_Listad.TextMatrix(r_int_Fila, 0) & "',"                         'GASPAR_CODGAS
          r_str_Parame = r_str_Parame & "'" & moddat_g_str_CodGrp & "',"                                          'GASPAR_CODEMP
          r_str_Parame = r_str_Parame & "1,"                                                                      'GASPAR_TIPTAB
          r_str_Parame = r_str_Parame & "'" & Format(Trim(grd_Listad.TextMatrix(r_int_Fila, 1)), "000") & "',"    'GASPAR_CODPRDS
          r_str_Parame = r_str_Parame & "'" & Format(Trim(grd_Listad.TextMatrix(r_int_Fila, 3)), "000000") & "'," 'GASPAR_CODPRY
          r_str_Parame = r_str_Parame & "'" & grd_Listad.TextMatrix(r_int_Fila, 5) & "',"                         'GASPAR_CODMON
          r_str_Parame = r_str_Parame & "'" & CDbl(grd_Listad.TextMatrix(r_int_Fila, 7)) & "',"                   'GASPAR_GASTAS_MTO
          r_str_Parame = r_str_Parame & "0,"                                                                      'GASPAR_GASNOT_MTO
          r_str_Parame = r_str_Parame & "0,"                                                                      'GASPAR_REGMIN_TASINM
          r_str_Parame = r_str_Parame & "0,"                                                                      'GASPAR_REGMIN_FACINM
          r_str_Parame = r_str_Parame & "0,"                                                                      'GASPAR_REGMIN_FICINM
          r_str_Parame = r_str_Parame & "0,"                                                                      'GASPAR_REGMIN_TASEST
          r_str_Parame = r_str_Parame & "0,"                                                                      'GASPAR_REGMIN_FACEST
          r_str_Parame = r_str_Parame & "0,"                                                                      'GASPAR_REGMIN_FICEST
          r_str_Parame = r_str_Parame & "0,"                                                                      'GASPAR_REGGAR_TAS000
          r_str_Parame = r_str_Parame & "0,"                                                                      'GASPAR_REGGAR_TAS001
          r_str_Parame = r_str_Parame & "0,"                                                                      'GASPAR_REGGAR_TAS002
          r_str_Parame = r_str_Parame & grd_Listad.TextMatrix(r_int_Fila, 8) & ","                                'GASPAR_SITUAC
          
          'Datos de Auditoria
          r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodUsu & "', "
          r_str_Parame = r_str_Parame & "'" & modgen_g_str_NombPC & "', "
          r_str_Parame = r_str_Parame & "'" & UCase(App.EXEName) & "', "
          r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodSuc & "', "
          r_str_Parame = r_str_Parame & grd_Listad.TextMatrix(r_int_Fila, 9) & ") "
          
         If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 2) Then
            Exit Sub
         End If
       End If
   Next
  
   MsgBox "El proceso se grabó exitosamente.", vbInformation, modgen_g_str_NomPlt
   Unload Me
End Sub

Private Sub pnl_Importe_Click()
   If Len(Trim(pnl_Importe.Tag)) = 0 Or pnl_Importe.Tag = "D" Then
      pnl_Importe.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 7, "N")
   Else
      pnl_Importe.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 7, "N-")
   End If
End Sub

Private Sub pnl_Moneda_Click()
   If Len(Trim(pnl_Moneda.Tag)) = 0 Or pnl_Moneda.Tag = "D" Then
      pnl_Moneda.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 5, "C")
   Else
      pnl_Moneda.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 5, "C-")
   End If
End Sub

Private Sub pnl_Producto_Click()
   If Len(Trim(pnl_Producto.Tag)) = 0 Or pnl_Producto.Tag = "D" Then
      pnl_Producto.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_Producto.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub

Private Sub pnl_Proyecto_Click()
   If Len(Trim(pnl_Proyecto.Tag)) = 0 Or pnl_Proyecto.Tag = "D" Then
      pnl_Proyecto.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 4, "C")
   Else
      pnl_Proyecto.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 4, "C-")
   End If
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

Private Sub txt_Importe1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_DirEle1)
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

Private Sub cmb_Moneda_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Importe2)
   End If
End Sub

Private Sub txt_Importe2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Aceptar)
   End If
End Sub
Private Function fs_Valida_Insercion(ByVal p_CodPrd As String, ByVal p_CodPry As String) As Integer
Dim r_str_Parame        As String

   fs_Valida_Insercion = 0
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "  SELECT COUNT(*) AS CONTADOR "
   r_str_Parame = r_str_Parame & "    FROM TRA_GASPAR "
   r_str_Parame = r_str_Parame & "         INNER JOIN PRY_DATGEN ON DATGEN_CODIGO = GASPAR_CODPRY AND TRIM(TO_CHAR(TRIM(DATGEN_CODPRT), '000000')) = GASPAR_CODEMP "
   r_str_Parame = r_str_Parame & "   WHERE GASPAR_CODPRD = '" & Format(p_CodPrd, "000") & "'"
   r_str_Parame = r_str_Parame & "     AND GASPAR_CODPRY = '" & Format(p_CodPry, "000000") & "'"
   r_str_Parame = r_str_Parame & "     AND GASPAR_TIPTAB = 1 "
   
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
