VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_Caj_PPgHip_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10560
   ClientLeft      =   3855
   ClientTop       =   705
   ClientWidth     =   11595
   Icon            =   "OpeTra_frm_063.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10560
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel2 
      Height          =   11475
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   20241
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
      Begin Threed.SSPanel SSPanel70 
         Height          =   705
         Left            =   30
         TabIndex        =   39
         Top             =   6150
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   1244
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
         Begin Threed.SSPanel pnl_NueSal 
            Height          =   315
            Left            =   1950
            TabIndex        =   40
            Top             =   30
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
         Begin EditLib.fpLongInteger ipp_NuePla 
            Height          =   315
            Left            =   1950
            TabIndex        =   42
            Top             =   360
            Width           =   735
            _Version        =   196608
            _ExtentX        =   1296
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
            MaxValue        =   "999"
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
         Begin Threed.SSPanel pnl_CuoPag 
            Height          =   315
            Left            =   8670
            TabIndex        =   48
            Top             =   360
            Width           =   945
            _Version        =   65536
            _ExtentX        =   1667
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "240 "
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
         Begin Threed.SSPanel pnl_CuoApr 
            Height          =   315
            Left            =   8670
            TabIndex        =   50
            Top             =   30
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
         Begin VB.Label Label23 
            Caption         =   "Cuota Fija Aprobada:"
            Height          =   285
            Left            =   6780
            LinkItem        =   "60"
            TabIndex        =   51
            Top             =   30
            Width           =   1635
         End
         Begin VB.Label Label22 
            Caption         =   "Nro. Cuotas Pagadas:"
            Height          =   285
            Left            =   6780
            LinkItem        =   "60"
            TabIndex        =   49
            Top             =   360
            Width           =   1635
         End
         Begin VB.Label Label29 
            Caption         =   "Nro. Cuotas Restantes:"
            Height          =   285
            Left            =   60
            LinkItem        =   "60"
            TabIndex        =   43
            Top             =   360
            Width           =   1665
         End
         Begin VB.Label Label20 
            Caption         =   "Nuevo Saldo Capital:"
            Height          =   285
            Left            =   60
            LinkItem        =   "60"
            TabIndex        =   41
            Top             =   30
            Width           =   1635
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1425
         Left            =   30
         TabIndex        =   29
         Top             =   4680
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin Threed.SSPanel pnl_Penali 
            Height          =   315
            Left            =   1950
            TabIndex        =   31
            Top             =   1050
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
         Begin Threed.SSPanel pnl_SegDes 
            Height          =   315
            Left            =   1950
            TabIndex        =   33
            Top             =   60
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
         Begin Threed.SSPanel pnl_Intere 
            Height          =   315
            Left            =   1950
            TabIndex        =   35
            Top             =   390
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
         Begin Threed.SSPanel pnl_Capita 
            Height          =   315
            Left            =   1950
            TabIndex        =   37
            Top             =   720
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
         Begin VB.Label Label19 
            Caption         =   "Capital:"
            Height          =   285
            Left            =   60
            TabIndex        =   38
            Top             =   720
            Width           =   1365
         End
         Begin VB.Label Label18 
            Caption         =   "Intereses:"
            Height          =   285
            Left            =   60
            TabIndex        =   36
            Top             =   390
            Width           =   1365
         End
         Begin VB.Label Label17 
            Caption         =   "Seguro Desgrav.:"
            Height          =   285
            Left            =   60
            TabIndex        =   34
            Top             =   60
            Width           =   1365
         End
         Begin VB.Label Label8 
            Caption         =   "Penalidad:"
            Height          =   285
            Left            =   60
            TabIndex        =   32
            Top             =   1050
            Width           =   975
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   2415
         Left            =   30
         TabIndex        =   1
         Top             =   2220
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   4260
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
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   60
            Width           =   9525
         End
         Begin VB.TextBox txt_NumCom 
            Height          =   315
            Left            =   1950
            MaxLength       =   25
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   720
            Width           =   1635
         End
         Begin VB.ComboBox cmb_CtaBan 
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   390
            Width           =   4125
         End
         Begin EditLib.fpDateTime ipp_FecPag 
            Height          =   315
            Left            =   1950
            TabIndex        =   5
            Top             =   1050
            Width           =   1605
            _Version        =   196608
            _ExtentX        =   2831
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
         Begin EditLib.fpDoubleSingle ipp_Import 
            Height          =   315
            Left            =   1950
            TabIndex        =   6
            Top             =   1380
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
         Begin Threed.SSPanel pnl_ITFImp 
            Height          =   315
            Left            =   1950
            TabIndex        =   7
            Top             =   1710
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
         Begin Threed.SSPanel pnl_TotImp 
            Height          =   315
            Left            =   1950
            TabIndex        =   8
            Top             =   2040
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   16777215
            BackColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin VB.Label Label6 
            Caption         =   "Banco:"
            Height          =   315
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Nro. Cuenta:"
            Height          =   285
            Left            =   60
            TabIndex        =   15
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Fecha de Pago:"
            Height          =   315
            Left            =   60
            TabIndex        =   14
            Top             =   1050
            Width           =   1245
         End
         Begin VB.Label Label11 
            Caption         =   "Nro. Comprobante:"
            Height          =   285
            Left            =   60
            TabIndex        =   13
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label41 
            Caption         =   "Importe Depositado:"
            Height          =   285
            Left            =   60
            TabIndex        =   12
            Top             =   1380
            Width           =   1515
         End
         Begin VB.Label Label4 
            Caption         =   "(*) Debe incluir ITF"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1950
            TabIndex        =   11
            Top             =   1440
            Width           =   1515
         End
         Begin VB.Label Label9 
            Caption         =   "Neto Pagado:"
            Height          =   285
            Left            =   60
            TabIndex        =   10
            Top             =   2040
            Width           =   1365
         End
         Begin VB.Label Label13 
            Caption         =   "ITF:"
            Height          =   285
            Left            =   60
            TabIndex        =   9
            Top             =   1710
            Width           =   975
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   1425
         Left            =   30
         TabIndex        =   17
         Top             =   750
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin Threed.SSPanel pnl_NumOpe 
            Height          =   315
            Left            =   1950
            TabIndex        =   18
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
            Left            =   1950
            TabIndex        =   19
            Top             =   720
            Width           =   9525
            _Version        =   65536
            _ExtentX        =   16801
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
            Left            =   1950
            TabIndex        =   20
            Top             =   1050
            Width           =   4125
            _Version        =   65536
            _ExtentX        =   7276
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
            Left            =   1950
            TabIndex        =   21
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
            TabIndex        =   25
            Top             =   390
            Width           =   1395
         End
         Begin VB.Label Label1 
            Caption         =   "Moneda:"
            Height          =   315
            Left            =   60
            TabIndex        =   24
            Top             =   1050
            Width           =   1125
         End
         Begin VB.Label Label2 
            Caption         =   "Nro. de Operación:"
            Height          =   315
            Left            =   60
            TabIndex        =   23
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label Label3 
            Caption         =   "Nombre Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   22
            Top             =   720
            Width           =   1125
         End
      End
      Begin Threed.SSPanel SSPanel12 
         Height          =   765
         Left            =   30
         TabIndex        =   26
         Top             =   10650
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10830
            Picture         =   "OpeTra_frm_063.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   10140
            Picture         =   "OpeTra_frm_063.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
         Begin MSComDlg.CommonDialog dlg_Guarda 
            Left            =   4050
            Top             =   120
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
      Begin Threed.SSPanel SSPanel22 
         Height          =   2895
         Left            =   30
         TabIndex        =   30
         Top             =   7710
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin TabDlg.SSTab tab_Cronog 
            Height          =   2775
            Left            =   60
            TabIndex        =   52
            Top             =   60
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   4895
            _Version        =   393216
            Style           =   1
            Tabs            =   5
            TabsPerRow      =   5
            TabHeight       =   520
            TabCaption(0)   =   "Cronograma - Cliente TNC"
            TabPicture(0)   =   "OpeTra_frm_063.frx":0890
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "pnl_CliNCo_TotCuo"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "pnl_CliNCo_OtrCar"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "SSPanel62"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "SSPanel61"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "SSPanel59"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "SSPanel36"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "SSPanel35"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "SSPanel34"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "SSPanel33"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "SSPanel4"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "grd_CliNCo_Listad"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "pnl_CliNCo_Intere"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "pnl_CliNCo_SegPre"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "pnl_CliNCo_SegViv"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "pnl_CliNCo_Capita"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "SSPanel30"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).ControlCount=   16
            TabCaption(1)   =   "Cliente - Tramo Concesional"
            TabPicture(1)   =   "OpeTra_frm_063.frx":08AC
            Tab(1).ControlEnabled=   0   'False
            Tab(1).ControlCount=   0
            TabCaption(2)   =   "Mivivienda - Tramo No Concesional"
            TabPicture(2)   =   "OpeTra_frm_063.frx":08C8
            Tab(2).ControlEnabled=   0   'False
            Tab(2).ControlCount=   0
            TabCaption(3)   =   "Mivivienda - Tramo Concesional"
            TabPicture(3)   =   "OpeTra_frm_063.frx":08E4
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "pnl_MViCon_TotCuo"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Cofide"
            TabPicture(4)   =   "OpeTra_frm_063.frx":0900
            Tab(4).ControlEnabled=   0   'False
            Tab(4).ControlCount=   0
            Begin Threed.SSPanel pnl_MViCon_TotCuo 
               Height          =   285
               Left            =   -67470
               TabIndex        =   53
               Top             =   2370
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel SSPanel30 
               Height          =   285
               Left            =   3450
               TabIndex        =   54
               Top             =   360
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel pnl_CliNCo_Capita 
               Height          =   285
               Left            =   2280
               TabIndex        =   55
               Top             =   2370
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_CliNCo_SegViv 
               Height          =   285
               Left            =   5790
               TabIndex        =   56
               Top             =   2370
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_CliNCo_SegPre 
               Height          =   285
               Left            =   4620
               TabIndex        =   57
               Top             =   2370
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_CliNCo_Intere 
               Height          =   285
               Left            =   3450
               TabIndex        =   58
               Top             =   2370
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin MSFlexGridLib.MSFlexGrid grd_CliNCo_Listad 
               Height          =   1695
               Left            =   30
               TabIndex        =   59
               Top             =   660
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   2990
               _Version        =   393216
               Rows            =   21
               Cols            =   9
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel23 
               Height          =   285
               Left            =   -67530
               TabIndex        =   60
               Top             =   360
               Width           =   2370
               _Version        =   65536
               _ExtentX        =   4180
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel25 
               Height          =   285
               Left            =   -65190
               TabIndex        =   61
               Top             =   360
               Width           =   2370
               _Version        =   65536
               _ExtentX        =   4180
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel26 
               Height          =   285
               Left            =   -74940
               TabIndex        =   62
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
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
               Left            =   -73770
               TabIndex        =   63
               Top             =   360
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
               ForeColor       =   16777215
               BackColor       =   32768
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
               Left            =   -71970
               TabIndex        =   64
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
               ForeColor       =   16777215
               BackColor       =   32768
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
               Left            =   -70140
               TabIndex        =   65
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel31 
               Height          =   285
               Left            =   -66480
               TabIndex        =   66
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel32 
               Height          =   285
               Left            =   -64650
               TabIndex        =   67
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel37 
               Height          =   285
               Left            =   -68310
               TabIndex        =   68
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Comisión"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel40 
               Height          =   285
               Left            =   -74940
               TabIndex        =   69
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel48 
               Height          =   285
               Left            =   -73770
               TabIndex        =   70
               Top             =   360
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel50 
               Height          =   285
               Left            =   -71970
               TabIndex        =   71
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel51 
               Height          =   285
               Left            =   -70140
               TabIndex        =   72
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel52 
               Height          =   285
               Left            =   -66480
               TabIndex        =   73
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel53 
               Height          =   285
               Left            =   -64650
               TabIndex        =   74
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel57 
               Height          =   285
               Left            =   -68310
               TabIndex        =   75
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Comisión"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel4 
               Height          =   285
               Left            =   60
               TabIndex        =   76
               Top             =   360
               Width           =   795
               _Version        =   65536
               _ExtentX        =   1402
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel33 
               Height          =   285
               Left            =   840
               TabIndex        =   77
               Top             =   360
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel34 
               Height          =   285
               Left            =   2280
               TabIndex        =   78
               Top             =   360
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel35 
               Height          =   285
               Left            =   8130
               TabIndex        =   79
               Top             =   360
               Width           =   1320
               _Version        =   65536
               _ExtentX        =   2328
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel36 
               Height          =   285
               Left            =   9420
               TabIndex        =   80
               Top             =   360
               Width           =   1560
               _Version        =   65536
               _ExtentX        =   2752
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel59 
               Height          =   285
               Left            =   4620
               TabIndex        =   81
               Top             =   360
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg. Prest."
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel61 
               Height          =   285
               Left            =   5790
               TabIndex        =   82
               Top             =   360
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg. Vivienda"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel62 
               Height          =   285
               Left            =   6960
               TabIndex        =   83
               Top             =   360
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Portes"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel pnl_CliNCo_OtrCar 
               Height          =   285
               Left            =   6960
               TabIndex        =   84
               Top             =   2370
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_CliNCo_TotCuo 
               Height          =   285
               Left            =   8130
               TabIndex        =   85
               Top             =   2370
               Width           =   1320
               _Version        =   65536
               _ExtentX        =   2328
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel SSPanel14 
               Height          =   285
               Left            =   -70950
               TabIndex        =   86
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin MSFlexGridLib.MSFlexGrid grd_MviCon_Listad 
               Height          =   1695
               Left            =   -74970
               TabIndex        =   87
               Top             =   660
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   2990
               _Version        =   393216
               Rows            =   21
               Cols            =   7
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel15 
               Height          =   285
               Left            =   -74940
               TabIndex        =   88
               Top             =   360
               Width           =   765
               _Version        =   65536
               _ExtentX        =   1349
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
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
               Left            =   -74190
               TabIndex        =   89
               Top             =   360
               Width           =   1515
               _Version        =   65536
               _ExtentX        =   2672
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
               ForeColor       =   16777215
               BackColor       =   32768
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
               Left            =   -72690
               TabIndex        =   90
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
               ForeColor       =   16777215
               BackColor       =   32768
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
               Left            =   -67470
               TabIndex        =   91
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
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
               Height          =   285
               Left            =   -65730
               TabIndex        =   92
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel20 
               Height          =   285
               Left            =   -69210
               TabIndex        =   93
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Comisión COFIDE"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel pnl_MViCon_Intere 
               Height          =   285
               Left            =   -70950
               TabIndex        =   94
               Top             =   2370
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_MViCon_Capita 
               Height          =   285
               Left            =   -72690
               TabIndex        =   95
               Top             =   2370
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_MViCon_Comisi 
               Height          =   285
               Left            =   -69210
               TabIndex        =   96
               Top             =   2370
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_CliCon_TotCuo 
               Height          =   285
               Left            =   -68370
               TabIndex        =   97
               Top             =   2370
               Width           =   2170
               _Version        =   65536
               _ExtentX        =   3828
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel SSPanel9 
               Height          =   285
               Left            =   -70530
               TabIndex        =   98
               Top             =   360
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3828
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interes"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin MSFlexGridLib.MSFlexGrid grd_CliCon_Listad 
               Height          =   1695
               Left            =   -74970
               TabIndex        =   99
               Top             =   660
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   2990
               _Version        =   393216
               Rows            =   21
               Cols            =   6
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel10 
               Height          =   285
               Left            =   -74940
               TabIndex        =   100
               Top             =   360
               Width           =   765
               _Version        =   65536
               _ExtentX        =   1349
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel11 
               Height          =   285
               Left            =   -74190
               TabIndex        =   101
               Top             =   360
               Width           =   1515
               _Version        =   65536
               _ExtentX        =   2672
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel8 
               Height          =   285
               Left            =   -72690
               TabIndex        =   102
               Top             =   360
               Width           =   2170
               _Version        =   65536
               _ExtentX        =   3828
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
               ForeColor       =   16777215
               BackColor       =   32768
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
               Left            =   -68370
               TabIndex        =   103
               Top             =   360
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3828
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel21 
               Height          =   285
               Left            =   -66210
               TabIndex        =   104
               Top             =   360
               Width           =   2235
               _Version        =   65536
               _ExtentX        =   3942
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel pnl_CliCon_Intere 
               Height          =   285
               Left            =   -70530
               TabIndex        =   105
               Top             =   2370
               Width           =   2170
               _Version        =   65536
               _ExtentX        =   3828
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_CliCon_Capita 
               Height          =   285
               Left            =   -72690
               TabIndex        =   106
               Top             =   2370
               Width           =   2170
               _Version        =   65536
               _ExtentX        =   3828
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin MSFlexGridLib.MSFlexGrid grd_MViNCo_Listad 
               Height          =   1695
               Left            =   -74970
               TabIndex        =   107
               Top             =   660
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   2990
               _Version        =   393216
               Rows            =   21
               Cols            =   10
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel24 
               Height          =   285
               Left            =   -71790
               TabIndex        =   108
               Top             =   360
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel38 
               Height          =   285
               Left            =   -74940
               TabIndex        =   109
               Top             =   360
               Width           =   705
               _Version        =   65536
               _ExtentX        =   1244
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel41 
               Height          =   285
               Left            =   -74250
               TabIndex        =   110
               Top             =   360
               Width           =   1425
               _Version        =   65536
               _ExtentX        =   2514
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel42 
               Height          =   285
               Left            =   -72840
               TabIndex        =   111
               Top             =   360
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel43 
               Height          =   285
               Left            =   -66390
               TabIndex        =   112
               Top             =   360
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel44 
               Height          =   285
               Left            =   -65310
               TabIndex        =   113
               Top             =   360
               Width           =   1290
               _Version        =   65536
               _ExtentX        =   2275
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel45 
               Height          =   285
               Left            =   -70710
               TabIndex        =   114
               Top             =   360
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg. Prest."
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel46 
               Height          =   285
               Left            =   -69630
               TabIndex        =   115
               Top             =   360
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg. Vivienda"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel47 
               Height          =   285
               Left            =   -68550
               TabIndex        =   116
               Top             =   360
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Portes"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel49 
               Height          =   285
               Left            =   -67470
               TabIndex        =   117
               Top             =   360
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "C. COFIDE"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel pnl_MViNCo_Capita 
               Height          =   285
               Left            =   -72840
               TabIndex        =   118
               Top             =   2370
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_MViNCo_SegViv 
               Height          =   285
               Left            =   -69630
               TabIndex        =   119
               Top             =   2370
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_MViNCo_SegPre 
               Height          =   285
               Left            =   -70710
               TabIndex        =   120
               Top             =   2370
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_MViNCo_Intere 
               Height          =   285
               Left            =   -71790
               TabIndex        =   121
               Top             =   2370
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_MViNCo_OtrCar 
               Height          =   285
               Left            =   -68550
               TabIndex        =   122
               Top             =   2370
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_MViNCo_TotCuo 
               Height          =   285
               Left            =   -66390
               TabIndex        =   123
               Top             =   2370
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_MViNCo_Comisi 
               Height          =   285
               Left            =   -67470
               TabIndex        =   124
               Top             =   2370
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_CofNCo_TotCuo 
               Height          =   285
               Left            =   -67500
               TabIndex        =   125
               Top             =   2370
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel SSPanel55 
               Height          =   285
               Left            =   -70980
               TabIndex        =   126
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin MSFlexGridLib.MSFlexGrid grd_CofNCo_Listad 
               Height          =   1695
               Left            =   -74970
               TabIndex        =   127
               Top             =   660
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   2990
               _Version        =   393216
               Rows            =   21
               Cols            =   7
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel56 
               Height          =   285
               Left            =   -74940
               TabIndex        =   128
               Top             =   360
               Width           =   765
               _Version        =   65536
               _ExtentX        =   1349
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel58 
               Height          =   285
               Left            =   -74220
               TabIndex        =   129
               Top             =   360
               Width           =   1515
               _Version        =   65536
               _ExtentX        =   2672
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel60 
               Height          =   285
               Left            =   -72720
               TabIndex        =   130
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel63 
               Height          =   285
               Left            =   -67500
               TabIndex        =   131
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel64 
               Height          =   285
               Left            =   -65760
               TabIndex        =   132
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel65 
               Height          =   285
               Left            =   -69240
               TabIndex        =   133
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Comisión COFIDE"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel pnl_CofNCo_Intere 
               Height          =   285
               Left            =   -70980
               TabIndex        =   134
               Top             =   2370
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_CofNCo_Capita 
               Height          =   285
               Left            =   -72720
               TabIndex        =   135
               Top             =   2370
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_CofNCo_Comisi 
               Height          =   285
               Left            =   -69240
               TabIndex        =   136
               Top             =   2370
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin VB.Label Label15 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   -73230
               TabIndex        =   139
               Top             =   1470
               Width           =   945
            End
            Begin VB.Label Label14 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   -72930
               TabIndex        =   138
               Top             =   1470
               Width           =   945
            End
            Begin VB.Label Label5 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   -72930
               TabIndex        =   137
               Top             =   1470
               Width           =   945
            End
         End
      End
      Begin Threed.SSPanel SSPanel39 
         Height          =   765
         Left            =   30
         TabIndex        =   44
         Top             =   6900
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
         Begin VB.CommandButton cmd_CalCuo 
            Height          =   675
            Left            =   10830
            Picture         =   "OpeTra_frm_063.frx":091C
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Calcular Cuota"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   46
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
            Height          =   585
            Left            =   660
            TabIndex        =   47
            Top             =   30
            Width           =   6435
            _Version        =   65536
            _ExtentX        =   11351
            _ExtentY        =   1032
            _StockProps     =   15
            Caption         =   "Cobro por Banco - Crédito Hipotecario - Prepagos"
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
            Picture         =   "OpeTra_frm_063.frx":0C2E
            Top             =   60
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_Caj_PPgHip_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_CliNCo()         As modcal_g_est_CuoCli
Dim l_arr_CliCon()         As modcal_g_est_CuoCli
Dim l_arr_MViCon()         As modcal_g_est_CuoCli
Dim l_arr_MViNCo()         As modcal_g_est_CuoCli
Dim l_arr_CofNCo()         As modcal_g_est_CuoCli
Dim l_dbl_PorITF        As Double
Dim l_arr_CodBan()      As moddat_tpo_Genera
Dim l_arr_CtaBan()      As moddat_tpo_Genera
Dim l_int_CuoPen        As Integer
Dim l_int_NumCuo        As Integer
Dim l_int_CuoExt        As Integer
Dim l_int_AplViv        As Integer
Dim l_int_DiaPag        As Integer
Dim l_dbl_TasInt        As Double
Dim l_dbl_SegDes        As Double
Dim l_dbl_FoIDes        As Double
Dim l_dbl_FoIViv        As Double
Dim l_dbl_Portes        As Double
Dim l_dbl_ValInm        As Double
Dim l_dbl_TipCam        As Double
Dim l_dbl_CuoApr        As Double
Dim l_dbl_TasCof        As Double
Dim l_dbl_ComCof        As Double
Dim l_dbl_TasMVi        As Double
Dim l_dbl_CuoFij        As Double
Dim l_dbl_SalCap        As Double
Dim l_dbl_Cuo_Portes    As Double
Dim l_dbl_Cuo_SegInm    As Double
Dim l_dbl_Cuo_SegDes    As Double
Dim l_dbl_PPG_PorPen    As Double
Dim l_int_PPG_NumCuo    As Integer
Dim l_dbl_IntDia        As Double
Dim l_str_FecVct        As String
Dim l_int_CuoPag        As Integer
Dim l_str_UltPag        As String
Dim l_str_PrxVct        As String
Dim l_dbl_PorCon        As Double

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
   Call gs_SetFocus(txt_NumCom)
End Sub

Private Sub cmb_CtaBan_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CtaBan_Click
   End If
End Sub

Private Sub cmd_CalCuo_Click()
   Dim r_dbl_MtoNCo     As Double
   Dim r_dbl_MtoCon     As Double

   If Not (CDate(ipp_FecPag.Text) >= CDate(l_str_UltPag) And CDate(ipp_FecPag.Text) < CDate(l_str_PrxVct)) Then
      MsgBox "La Fecha de Pago debe estar entre la Fecha del Ultimo Pago del Cliente y la Fecha de Próximo Vencimiento de su cuota.", vbExclamation, modgen_g_str_NomPlt
      
      Call gs_SetFocus(ipp_FecPag)
      Exit Sub
   End If

   If CDbl(pnl_TotImp.Caption) < CDbl(Format(l_int_PPG_NumCuo * l_dbl_CuoFij, "#####0.00")) Then
      MsgBox "El Monto mínimo de Prepago tiene que ser mayor al importe de " & CStr(l_int_PPG_NumCuo) & " cuotas fijas.", vbExclamation, modgen_g_str_NomPlt
      
      Call gs_SetFocus(ipp_Import)
      Exit Sub
   End If
   
   If (ipp_NuePla.Value + l_int_CuoPag) Mod 12 > 0 Then
      MsgBox "El Número de Cuotas del Crédito tiene que ser múltiplo de 12.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_NuePla)
      Exit Sub
   End If
   
   
   'Calcular la Cuota
   
   Select Case moddat_g_str_CodPrd
      Case "002"
         'Calculando Cronograma de Pagos
         Call gs_Cronog_MiCasita(l_arr_CliNCo(), l_dbl_ValInm, CDbl(pnl_NueSal.Caption), ipp_NuePla.Value, 2, l_dbl_TasInt, l_dbl_FoIDes, l_int_AplViv, l_dbl_FoIViv, l_dbl_Portes, CDate(ipp_FecPag.Text), l_int_DiaPag, 0, , 1)
         
      Case "001"
         Call gs_Cronog_CRCPBP_NC_PrePag(l_arr_CliNCo(), CDbl(pnl_NueSal.Caption), l_dbl_PorCon, l_dbl_ValInm, ipp_NuePla.Value, l_dbl_TasInt, l_dbl_FoIDes, l_int_AplViv, l_dbl_FoIViv, l_dbl_Portes, ipp_FecPag.Text, l_str_PrxVct, l_int_DiaPag, r_dbl_MtoNCo, r_dbl_MtoCon)
   End Select
   
   Call fs_Cronog
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Buscar_DatGen
   
   Call gs_CentraForm(Me)

   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_FecSis
   
   Call moddat_gs_Carga_LisIte(cmb_CodBan, l_arr_CodBan, 1, "516")
   
   cmb_CodBan.ListIndex = -1
   cmb_CtaBan.Clear
   ipp_FecPag.Text = Format(date, "dd/mm/yyyy")
   txt_NumCom.Text = ""
   ipp_Import.Value = 0
   
   pnl_ITFImp.Caption = "0.00 "
   pnl_TotImp.Caption = "0.00 "
   
   'Cliente No Concesional
   grd_CliNCo_Listad.ColWidth(0) = 795
   grd_CliNCo_Listad.ColWidth(1) = 1425
   grd_CliNCo_Listad.ColWidth(2) = 1180
   grd_CliNCo_Listad.ColWidth(3) = 1170
   grd_CliNCo_Listad.ColWidth(4) = 1160
   grd_CliNCo_Listad.ColWidth(5) = 1160
   grd_CliNCo_Listad.ColWidth(6) = 1160
   grd_CliNCo_Listad.ColWidth(7) = 1320
   grd_CliNCo_Listad.ColWidth(8) = 1560
   
   grd_CliNCo_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CliNCo_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CliNCo_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(8) = flexAlignRightCenter

   'Mivivienda No Concesional
   grd_MViNCo_Listad.ColWidth(0) = 695
   grd_MViNCo_Listad.ColWidth(1) = 1415
   grd_MViNCo_Listad.ColWidth(2) = 1070
   grd_MViNCo_Listad.ColWidth(3) = 1070
   grd_MViNCo_Listad.ColWidth(4) = 1080
   grd_MViNCo_Listad.ColWidth(5) = 1080
   grd_MViNCo_Listad.ColWidth(6) = 1080
   grd_MViNCo_Listad.ColWidth(7) = 1080
   grd_MViNCo_Listad.ColWidth(8) = 1080
   grd_MViNCo_Listad.ColWidth(9) = 1290
   
   grd_MViNCo_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_MViNCo_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_MViNCo_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(8) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(9) = flexAlignRightCenter

   'Mivivienda Concesional
   grd_MviCon_Listad.ColWidth(0) = 770
   grd_MviCon_Listad.ColWidth(1) = 1485
   grd_MviCon_Listad.ColWidth(2) = 1730
   grd_MviCon_Listad.ColWidth(3) = 1740
   grd_MviCon_Listad.ColWidth(4) = 1740
   grd_MviCon_Listad.ColWidth(5) = 1740
   grd_MviCon_Listad.ColWidth(6) = 1740
   
   grd_MviCon_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_MviCon_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_MviCon_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_MviCon_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_MviCon_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_MviCon_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_MviCon_Listad.ColAlignment(6) = flexAlignRightCenter

   'Cliente Concesional
   grd_CliCon_Listad.ColWidth(0) = 770
   grd_CliCon_Listad.ColWidth(1) = 1485
   grd_CliCon_Listad.ColWidth(2) = 2170
   grd_CliCon_Listad.ColWidth(3) = 2160
   grd_CliCon_Listad.ColWidth(4) = 2170
   grd_CliCon_Listad.ColWidth(5) = 2170
   
   grd_CliCon_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CliCon_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CliCon_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(5) = flexAlignRightCenter

   'Cofide No Concesional
   grd_CofNCo_Listad.ColWidth(0) = 770
   grd_CofNCo_Listad.ColWidth(1) = 1485
   grd_CofNCo_Listad.ColWidth(2) = 1730
   grd_CofNCo_Listad.ColWidth(3) = 1740
   grd_CofNCo_Listad.ColWidth(4) = 1740
   grd_CofNCo_Listad.ColWidth(5) = 1740
   grd_CofNCo_Listad.ColWidth(6) = 1740
   
   grd_CofNCo_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CofNCo_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CofNCo_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(6) = flexAlignRightCenter

   Call gs_LimpiaGrid(grd_CliNCo_Listad)
   Call gs_LimpiaGrid(grd_CliCon_Listad)
   Call gs_LimpiaGrid(grd_MViNCo_Listad)
   Call gs_LimpiaGrid(grd_MviCon_Listad)
   Call gs_LimpiaGrid(grd_CofNCo_Listad)
   
   pnl_CliNCo_Capita.Caption = "0.00 "
   pnl_CliNCo_Intere.Caption = "0.00 "
   pnl_CliNCo_SegPre.Caption = "0.00 "
   pnl_CliNCo_SegViv.Caption = "0.00 "
   pnl_CliNCo_OtrCar.Caption = "0.00 "
   pnl_CliNCo_TotCuo.Caption = "0.00 "
   
   pnl_MViNCo_Capita.Caption = "0.00 "
   pnl_MViNCo_Intere.Caption = "0.00 "
   pnl_MViNCo_SegPre.Caption = "0.00 "
   pnl_MViNCo_SegViv.Caption = "0.00 "
   pnl_MViNCo_OtrCar.Caption = "0.00 "
   pnl_MViNCo_TotCuo.Caption = "0.00 "
   
   pnl_CofNCo_Capita.Caption = "0.00 "
   pnl_CofNCo_Intere.Caption = "0.00 "
   pnl_CofNCo_Comisi.Caption = "0.00 "
   pnl_CofNCo_TotCuo.Caption = "0.00 "
   
   pnl_CliCon_Capita.Caption = "0.00 "
   pnl_CliCon_Intere.Caption = "0.00 "
   pnl_CliCon_TotCuo.Caption = "0.00 "
   
   pnl_MViCon_Capita.Caption = "0.00 "
   pnl_MViCon_Intere.Caption = "0.00 "
   pnl_MViCon_Comisi.Caption = "0.00 "
   pnl_MViCon_TotCuo.Caption = "0.00 "
End Sub

Private Sub fs_Buscar_DatGen()
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   pnl_NumOpe.Caption = gf_Formato_NumOpe(moddat_g_str_NumOpe)
   pnl_DocIde.Caption = CStr(g_rst_Princi!hipmae_tdocli) & "-" & Trim(g_rst_Princi!hipmae_ndocli)
   pnl_NomCli.Caption = moddat_gf_Buscar_NomCli(g_rst_Princi!hipmae_tdocli, g_rst_Princi!hipmae_ndocli)
   pnl_Moneda.Caption = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!hipmae_moneda))
   
   moddat_g_int_TipDoc = g_rst_Princi!hipmae_tdocli
   moddat_g_str_NumDoc = Trim(g_rst_Princi!hipmae_ndocli)
   moddat_g_str_NumSol = Trim(g_rst_Princi!hipmae_numsol)
   moddat_g_str_CodPrd = Trim(g_rst_Princi!HIPMAE_CODPRD)
   moddat_g_str_CodSub = Trim(g_rst_Princi!HIPMAE_CODSUB)
   moddat_g_str_CodMod = Trim(g_rst_Princi!HIPMAE_CODMOD)
   moddat_g_int_TipMon = g_rst_Princi!hipmae_moneda

   l_int_NumCuo = g_rst_Princi!HIPMAE_NUMCUO
   l_int_CuoPen = g_rst_Princi!HIPMAE_CUOPEN
   l_int_CuoPag = l_int_NumCuo - l_int_CuoPen
   
   pnl_CuoPag.Caption = CStr(l_int_CuoPag) & " "
   
   'Obteniendo ITF
   If g_rst_Princi!HIPMAE_INDITF = 2 Then
      l_dbl_PorITF = 0
   Else
      l_dbl_PorITF = opecaj_gf_Consulta_ITF(Format(CDate(moddat_g_str_FecSis), "yyyymmdd"), 1)
   End If

   'Obteniendo Información de Tasas
   l_int_CuoExt = 0
   l_int_AplViv = 0
   l_int_DiaPag = 0
   l_dbl_TasInt = 0
   l_dbl_FoIDes = 0
   l_dbl_FoIViv = 0
   l_dbl_Portes = 0
   l_dbl_ValInm = 0
   l_dbl_TipCam = 0
   l_dbl_CuoApr = 0
   l_dbl_SalCap = 0
      
   l_dbl_TasInt = g_rst_Princi!HIPMAE_TASINT
   l_int_CuoExt = g_rst_Princi!HIPMAE_CUOANO
   l_dbl_FoIDes = g_rst_Princi!HIPMAE_FOIPRE
   l_int_AplViv = g_rst_Princi!HIPMAE_APLVIV
   l_dbl_FoIViv = g_rst_Princi!HIPMAE_FOIVIV
   l_dbl_Portes = g_rst_Princi!HIPMAE_OTRIMP
   l_int_DiaPag = g_rst_Princi!HIPMAE_DIAPAG
   l_dbl_TasMVi = g_rst_Princi!HIPMAE_TASMVI
   l_dbl_TasCof = g_rst_Princi!HIPMAE_TASCOF
   l_dbl_ComCof = g_rst_Princi!HIPMAE_COMCOF
   l_dbl_CuoFij = g_rst_Princi!HIPMAE_CUOFIJ
   l_dbl_SalCap = g_rst_Princi!HIPMAE_SALCAP + g_rst_Princi!HIPMAE_SALCON
   
   l_dbl_PorCon = CDbl(Format(g_rst_Princi!HIPMAE_IMPCON / (g_rst_Princi!HIPMAE_IMPNCO + g_rst_Princi!HIPMAE_IMPCON) * 100, "##0.0000"))
   
   If g_rst_Princi!HIPMAE_VCTANT > 0 Then
      l_str_FecVct = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_VCTANT))
   Else
      l_str_FecVct = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))
   End If
   
   l_str_PrxVct = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_PRXVCT))
   
   If g_rst_Princi!HIPMAE_ULTPAG > 0 Then
      l_str_UltPag = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_ULTPAG))
   Else
      l_str_UltPag = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))
   End If
   
   l_dbl_IntDia = (1 + (l_dbl_TasInt / 100)) ^ (1 / 360) - 1      'Calculando Tasa Diaria de Interes
   l_dbl_SegDes = (1 + (l_dbl_FoIDes / 100)) ^ (1 / 30) - 1       'Calculando Tasa Diaria de Seguro de Desgravamen
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Obteniendo Parámetros Mivivienda
   l_dbl_TasCof = 0
   l_dbl_ComCof = 0
   l_dbl_TasMVi = 0

   'Obteniendo Valor de Inmueble
   g_str_Parame = "SELECT * FROM TRA_EVATAS WHERE "
   g_str_Parame = g_str_Parame & "EVATAS_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      l_dbl_ValInm = g_rst_Princi!EVATAS_SUMASE_INM + g_rst_Princi!EVATAS_SUMASE_ES1 + g_rst_Princi!EVATAS_SUMASE_ES2 + g_rst_Princi!EVATAS_SUMASE_DEP
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Obteniendo Tipo de Cambio
   If moddat_g_int_TipMon <> 1 Then
      l_dbl_TipCam = moddat_gf_Obtiene_TipCam(1, moddat_g_int_TipMon)
   End If
   
   'Para obtener Cuota Aprobada en Créditos
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
   
      'If moddat_g_int_TipMon <> 1 Then
      '   l_dbl_CuoApr = CDbl(Format(g_rst_Princi!SOLMAE_CUOAPR_SOL / l_dbl_TipCam, "#####0.00"))
      'Else
      '   l_dbl_CuoApr = CDbl(Format(g_rst_Princi!SOLMAE_CUOAPR_SOL, "#####0.00"))
      'End If
      
      l_dbl_CuoApr = CDbl(Format(g_rst_Princi!SOLMAE_CUOAPR_MPR, "#####0.00"))
   End If
   
   pnl_CuoApr.Caption = Format(l_dbl_CuoApr, "###,##0.00") & " "
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   'Datos de Cuota
   l_dbl_Cuo_Portes = 0
   l_dbl_Cuo_SegInm = 0
   l_dbl_Cuo_SegDes = 0
   
   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1 AND "
   g_str_Parame = g_str_Parame & "HIPCUO_SITUAC = 2 "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_NUMCUO ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      l_dbl_Cuo_Portes = CDbl(Format(g_rst_Princi!HIPCUO_OTRORG - g_rst_Princi!HIPCUO_OTRPAG, "###,###,##0.00"))
      l_dbl_Cuo_SegInm = CDbl(Format(g_rst_Princi!HIPCUO_VIVORG - g_rst_Princi!HIPCUO_VIVPAG, "###,###,##0.00"))
      l_dbl_Cuo_SegDes = CDbl(Format(g_rst_Princi!HIPCUO_DESORG - g_rst_Princi!HIPCUO_DESPAG, "###,###,##0.00"))
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
   pnl_Penali.Caption = "0.00 "
   pnl_SegDes.Caption = Format(l_dbl_Cuo_SegDes, "###,##0.00") & " "
   pnl_Intere.Caption = "0.00 "
   pnl_Capita.Caption = "0.00 "

   'Buscando Porcentaje de Penalidad
   l_dbl_PPG_PorPen = 0
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "002", "501") Then
      l_dbl_PPG_PorPen = moddat_g_arr_Genera(1).Genera_Cantid
   End If
   
   l_int_PPG_NumCuo = 0
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "002", "502") Then
      l_int_PPG_NumCuo = moddat_g_arr_Genera(1).Genera_Cantid
   End If

   ipp_NuePla.Value = l_int_CuoPen
   ipp_NuePla.MaxValue = l_int_CuoPen
   
   Select Case moddat_g_str_CodPrd
      Case "001"
         tab_Cronog.TabCaption(0) = "Cliente - No Concesional"
         tab_Cronog.TabCaption(1) = "Cliente - Concesional"
         tab_Cronog.TabCaption(2) = "Mivivienda - No Concesional"
         tab_Cronog.TabCaption(3) = "Mivivienda - Concesional"
         
         tab_Cronog.TabVisible(1) = True
         tab_Cronog.TabVisible(2) = True
         tab_Cronog.TabVisible(3) = True
         tab_Cronog.TabVisible(4) = False
         
      Case "002"
         tab_Cronog.TabCaption(0) = "Cliente"
   
         tab_Cronog.TabVisible(1) = False
         tab_Cronog.TabVisible(2) = False
         tab_Cronog.TabVisible(3) = False
         tab_Cronog.TabVisible(4) = False
   
      Case "003"
         tab_Cronog.TabCaption(0) = "Cliente - No Concesional"
         tab_Cronog.TabCaption(1) = "Cliente - Concesional"
         tab_Cronog.TabCaption(2) = "Mivivienda - No Concesional"
         tab_Cronog.TabCaption(3) = "Mivivienda - Concesional"
         tab_Cronog.TabCaption(4) = "Cofide"
         
         tab_Cronog.TabVisible(1) = True
         tab_Cronog.TabVisible(2) = True
         tab_Cronog.TabVisible(3) = True
         tab_Cronog.TabVisible(4) = True
         
      Case "004"
         tab_Cronog.TabCaption(0) = "Cliente - No Concesional"
         tab_Cronog.TabCaption(1) = "Cliente - Concesional"
         tab_Cronog.TabCaption(2) = "Cofide - No Concesional"
         tab_Cronog.TabCaption(3) = "Cofide - Concesional"
   
         tab_Cronog.TabVisible(1) = True
         tab_Cronog.TabVisible(2) = True
         tab_Cronog.TabVisible(3) = True
         tab_Cronog.TabVisible(4) = False
   End Select
End Sub

Private Sub ipp_FecPag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Import)
   End If
End Sub

Private Sub ipp_FecPag_LostFocus()
   If IsDate(ipp_FecPag.Text) Then
      If Not (CDate(ipp_FecPag.Text) >= CDate(l_str_UltPag) And CDate(ipp_FecPag.Text) < CDate(l_str_PrxVct)) Then
         MsgBox "La Fecha de Pago debe estar entre la Fecha del Ultimo Pago del Cliente y la Fecha de Próximo Vencimiento de su cuota.", vbExclamation, modgen_g_str_NomPlt
         
         Call gs_SetFocus(ipp_FecPag)
         Exit Sub
      End If
      
      If CDbl(ipp_Import.Text) > 0 Then
         Call ipp_Import_LostFocus
      End If
   End If
End Sub

Private Sub ipp_Import_LostFocus()
   
   pnl_ITFImp.Caption = gf_NueImp_Numero(gf_Truncar_Numero(CDbl(ipp_Import.Text) * (l_dbl_PorITF / 100), 2)) & " "
   pnl_TotImp.Caption = Format(CDbl(ipp_Import.Text) - CDbl(Trim(pnl_ITFImp.Caption)), "###,###,##0.00") & " "
   
   If CDbl(pnl_TotImp.Caption) < CDbl(Format(l_int_PPG_NumCuo * l_dbl_CuoFij, "#####0.00")) Then
      MsgBox "El Monto mínimo de Prepago tiene que ser mayor al importe de " & CStr(l_int_PPG_NumCuo) & " cuotas fijas.", vbExclamation, modgen_g_str_NomPlt
      
      Call gs_SetFocus(ipp_Import)
      Exit Sub
   End If
   
   pnl_SegDes.Caption = Format(l_dbl_SalCap * (1 + l_dbl_SegDes) ^ CInt(CDate(ipp_FecPag.Text) - CDate(l_str_FecVct)) - l_dbl_SalCap, "###,##0.00") & " "
   pnl_Intere.Caption = Format(l_dbl_SalCap * (1 + l_dbl_IntDia) ^ CInt(CDate(ipp_FecPag.Text) - CDate(l_str_FecVct)) - l_dbl_SalCap, "###,##0.00") & " "

   pnl_Penali.Caption = Format(l_dbl_PPG_PorPen / 100 * (CDbl(pnl_TotImp.Caption) - CDbl(pnl_SegDes.Caption) - CDbl(pnl_Intere.Caption)), "###,##0.00") & " "
   
   pnl_Capita.Caption = Format(CDbl(pnl_TotImp.Caption) - CDbl(pnl_Penali.Caption) - CDbl(pnl_SegDes.Caption) - CDbl(pnl_Intere.Caption), "###,##0.00") & " "
   
   pnl_NueSal.Caption = Format(l_dbl_SalCap - CDbl(pnl_Capita.Caption), "###,##0.00") & " "
   
   Call gs_LimpiaGrid(grd_CliNCo_Listad)
   Call gs_LimpiaGrid(grd_CliCon_Listad)
   Call gs_LimpiaGrid(grd_MViNCo_Listad)
   Call gs_LimpiaGrid(grd_MviCon_Listad)
   Call gs_LimpiaGrid(grd_CofNCo_Listad)
   
   pnl_CliNCo_Capita.Caption = "0.00 "
   pnl_CliNCo_Intere.Caption = "0.00 "
   pnl_CliNCo_SegPre.Caption = "0.00 "
   pnl_CliNCo_SegViv.Caption = "0.00 "
   pnl_CliNCo_OtrCar.Caption = "0.00 "
   pnl_CliNCo_TotCuo.Caption = "0.00 "
   
   pnl_MViNCo_Capita.Caption = "0.00 "
   pnl_MViNCo_Intere.Caption = "0.00 "
   pnl_MViNCo_SegPre.Caption = "0.00 "
   pnl_MViNCo_SegViv.Caption = "0.00 "
   pnl_MViNCo_OtrCar.Caption = "0.00 "
   pnl_MViNCo_TotCuo.Caption = "0.00 "
   
   pnl_CofNCo_Capita.Caption = "0.00 "
   pnl_CofNCo_Intere.Caption = "0.00 "
   pnl_CofNCo_Comisi.Caption = "0.00 "
   pnl_CofNCo_TotCuo.Caption = "0.00 "
   
   pnl_CliCon_Capita.Caption = "0.00 "
   pnl_CliCon_Intere.Caption = "0.00 "
   pnl_CliCon_TotCuo.Caption = "0.00 "
   
   pnl_MViCon_Capita.Caption = "0.00 "
   pnl_MViCon_Intere.Caption = "0.00 "
   pnl_MViCon_Comisi.Caption = "0.00 "
   pnl_MViCon_TotCuo.Caption = "0.00 "
End Sub

Private Sub ipp_Import_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub ipp_NuePla_Change()
   Call gs_LimpiaGrid(grd_CliNCo_Listad)
   Call gs_LimpiaGrid(grd_CliCon_Listad)
   Call gs_LimpiaGrid(grd_MViNCo_Listad)
   Call gs_LimpiaGrid(grd_MviCon_Listad)
   Call gs_LimpiaGrid(grd_CofNCo_Listad)
   
   pnl_CliNCo_Capita.Caption = "0.00 "
   pnl_CliNCo_Intere.Caption = "0.00 "
   pnl_CliNCo_SegPre.Caption = "0.00 "
   pnl_CliNCo_SegViv.Caption = "0.00 "
   pnl_CliNCo_OtrCar.Caption = "0.00 "
   pnl_CliNCo_TotCuo.Caption = "0.00 "
   
   pnl_MViNCo_Capita.Caption = "0.00 "
   pnl_MViNCo_Intere.Caption = "0.00 "
   pnl_MViNCo_SegPre.Caption = "0.00 "
   pnl_MViNCo_SegViv.Caption = "0.00 "
   pnl_MViNCo_OtrCar.Caption = "0.00 "
   pnl_MViNCo_TotCuo.Caption = "0.00 "
   
   pnl_CofNCo_Capita.Caption = "0.00 "
   pnl_CofNCo_Intere.Caption = "0.00 "
   pnl_CofNCo_Comisi.Caption = "0.00 "
   pnl_CofNCo_TotCuo.Caption = "0.00 "
   
   pnl_CliCon_Capita.Caption = "0.00 "
   pnl_CliCon_Intere.Caption = "0.00 "
   pnl_CliCon_TotCuo.Caption = "0.00 "
   
   pnl_MViCon_Capita.Caption = "0.00 "
   pnl_MViCon_Intere.Caption = "0.00 "
   pnl_MViCon_Comisi.Caption = "0.00 "
   pnl_MViCon_TotCuo.Caption = "0.00 "
End Sub

Private Sub txt_NumCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecPag)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
   End If
End Sub

Private Sub txt_NumCom_GotFocus()
   Call gs_SelecTodo(txt_NumCom)
End Sub

Private Sub fs_Cronog()
   Dim r_dbl_Capita  As Double
   Dim r_dbl_Intere  As Double
   Dim r_dbl_Comisi  As Double
   Dim r_dbl_SegPre  As Double
   Dim r_dbl_SegViv  As Double
   Dim r_dbl_OtrCar  As Double
   Dim r_dbl_TotCuo  As Double
   Dim r_int_Contad  As Integer
   
   'Cliente No Concesional
   r_dbl_Capita = 0
   r_dbl_Intere = 0
   r_dbl_SegPre = 0
   r_dbl_SegViv = 0
   r_dbl_OtrCar = 0
   r_dbl_TotCuo = 0
   
   grd_CliNCo_Listad.Redraw = False
   For r_int_Contad = 1 To UBound(l_arr_CliNCo)
      grd_CliNCo_Listad.Rows = grd_CliNCo_Listad.Rows + 1
      grd_CliNCo_Listad.Row = grd_CliNCo_Listad.Rows - 1
      
      'Número de Cuota
      grd_CliNCo_Listad.Col = 0
      grd_CliNCo_Listad.Text = Format(r_int_Contad, "000")
   
      'Fecha de Vencimiento
      grd_CliNCo_Listad.Col = 1
      grd_CliNCo_Listad.Text = l_arr_CliNCo(r_int_Contad).CuoCli_FecVct
   
      'Capital
      grd_CliNCo_Listad.Col = 2
      grd_CliNCo_Listad.Text = Format(l_arr_CliNCo(r_int_Contad).CuoCli_Capita, "###,###,##0.00")
      r_dbl_Capita = r_dbl_Capita + CDbl(grd_CliNCo_Listad)
      
      'Interes
      grd_CliNCo_Listad.Col = 3
      grd_CliNCo_Listad.Text = Format(l_arr_CliNCo(r_int_Contad).CuoCli_Intere, "###,###,##0.00")
      r_dbl_Intere = r_dbl_Intere + CDbl(grd_CliNCo_Listad)
   
      'Seguro Desgravamen
      grd_CliNCo_Listad.Col = 4
      grd_CliNCo_Listad.Text = Format(l_arr_CliNCo(r_int_Contad).CuoCli_SegPre, "###,###,##0.00")
      r_dbl_SegPre = r_dbl_SegPre + CDbl(grd_CliNCo_Listad)
   
      'Seguro Vivienda
      grd_CliNCo_Listad.Col = 5
      grd_CliNCo_Listad.Text = Format(l_arr_CliNCo(r_int_Contad).CuoCli_SegViv, "###,###,##0.00")
      r_dbl_SegViv = r_dbl_SegViv + CDbl(grd_CliNCo_Listad)
   
      'Otros Cargos
      grd_CliNCo_Listad.Col = 6
      grd_CliNCo_Listad.Text = Format(l_arr_CliNCo(r_int_Contad).CuoCli_Portes, "###,###,##0.00")
      r_dbl_OtrCar = r_dbl_OtrCar + CDbl(grd_CliNCo_Listad)
   
      'Valor Cuota
      grd_CliNCo_Listad.Col = 7
      grd_CliNCo_Listad.Text = Format(l_arr_CliNCo(r_int_Contad).CuoCli_ValCuo, "###,###,##0.00")
      r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_CliNCo_Listad)
   
      'Saldo Capital
      grd_CliNCo_Listad.Col = 8
      grd_CliNCo_Listad.Text = Format(l_arr_CliNCo(r_int_Contad).CuoCli_SalCap, "###,###,##0.00")
   Next r_int_Contad
   
   grd_CliNCo_Listad.Redraw = True
   
   pnl_CliNCo_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
   pnl_CliNCo_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
   pnl_CliNCo_SegPre.Caption = Format(r_dbl_SegPre, "###,###,##0.00") & " "
   pnl_CliNCo_SegViv.Caption = Format(r_dbl_SegViv, "###,###,##0.00") & " "
   pnl_CliNCo_OtrCar.Caption = Format(r_dbl_OtrCar, "###,###,##0.00") & " "
   pnl_CliNCo_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
   
   If grd_CliNCo_Listad.Rows > 0 Then
      Call gs_UbiIniGrid(grd_CliNCo_Listad)
   End If


   If moddat_g_str_CodPrd <> "002" Then
      'Mivivienda No Concesional
      r_dbl_Capita = 0
      r_dbl_Intere = 0
      r_dbl_SegPre = 0
      r_dbl_SegViv = 0
      r_dbl_OtrCar = 0
      r_dbl_Comisi = 0
      r_dbl_TotCuo = 0
      
      grd_MViNCo_Listad.Redraw = False
      For r_int_Contad = 1 To UBound(l_arr_MViNCo)
         grd_MViNCo_Listad.Rows = grd_MViNCo_Listad.Rows + 1
         grd_MViNCo_Listad.Row = grd_MViNCo_Listad.Rows - 1
         
         'Número de Cuota
         grd_MViNCo_Listad.Col = 0
         grd_MViNCo_Listad.Text = Format(r_int_Contad, "000")
      
         'Fecha de Vencimiento
         grd_MViNCo_Listad.Col = 1
         grd_MViNCo_Listad.Text = l_arr_MViNCo(r_int_Contad).CuoCli_FecVct
      
         'Capital
         grd_MViNCo_Listad.Col = 2
         grd_MViNCo_Listad.Text = Format(l_arr_MViNCo(r_int_Contad).CuoCli_Capita, "###,###,##0.00")
         r_dbl_Capita = r_dbl_Capita + CDbl(grd_MViNCo_Listad)
         
         'Interes
         grd_MViNCo_Listad.Col = 3
         grd_MViNCo_Listad.Text = Format(l_arr_MViNCo(r_int_Contad).CuoCli_Intere, "###,###,##0.00")
         r_dbl_Intere = r_dbl_Intere + CDbl(grd_MViNCo_Listad)
      
         'Seguro Desgravamen
         grd_MViNCo_Listad.Col = 4
         grd_MViNCo_Listad.Text = Format(l_arr_MViNCo(r_int_Contad).CuoCli_SegPre, "###,###,##0.00")
         r_dbl_SegPre = r_dbl_SegPre + CDbl(grd_MViNCo_Listad)
      
         'Seguro Vivienda
         grd_MViNCo_Listad.Col = 5
         grd_MViNCo_Listad.Text = Format(l_arr_MViNCo(r_int_Contad).CuoCli_SegViv, "###,###,##0.00")
         r_dbl_SegViv = r_dbl_SegViv + CDbl(grd_MViNCo_Listad)
      
         'Otros Cargos
         grd_MViNCo_Listad.Col = 6
         grd_MViNCo_Listad.Text = Format(l_arr_MViNCo(r_int_Contad).CuoCli_Portes, "###,###,##0.00")
         r_dbl_OtrCar = r_dbl_OtrCar + CDbl(grd_MViNCo_Listad)
      
         'Comisión COFIDE
         grd_MViNCo_Listad.Col = 7
         grd_MViNCo_Listad.Text = Format(l_arr_MViNCo(r_int_Contad).CuoCli_Comisi, "###,###,##0.00")
         r_dbl_Comisi = r_dbl_Comisi + CDbl(grd_MViNCo_Listad)
      
         'Valor Cuota
         grd_MViNCo_Listad.Col = 8
         grd_MViNCo_Listad.Text = Format(l_arr_MViNCo(r_int_Contad).CuoCli_ValCuo, "###,###,##0.00")
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_MViNCo_Listad)
      
         'Saldo Capital
         grd_MViNCo_Listad.Col = 9
         grd_MViNCo_Listad.Text = Format(l_arr_MViNCo(r_int_Contad).CuoCli_SalCap, "###,###,##0.00")
      Next r_int_Contad
      
      grd_MViNCo_Listad.Redraw = True
      
      pnl_MViNCo_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
      pnl_MViNCo_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
      pnl_MViNCo_SegPre.Caption = Format(r_dbl_SegPre, "###,###,##0.00") & " "
      pnl_MViNCo_SegViv.Caption = Format(r_dbl_SegViv, "###,###,##0.00") & " "
      pnl_MViNCo_OtrCar.Caption = Format(r_dbl_OtrCar, "###,###,##0.00") & " "
      pnl_MViNCo_Comisi.Caption = Format(r_dbl_Comisi, "###,###,##0.00") & " "
      pnl_MViNCo_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
      
      If grd_MViNCo_Listad.Rows > 0 Then
         Call gs_UbiIniGrid(grd_MViNCo_Listad)
      End If
      
   
      'Mivivienda Concesional
      r_dbl_Capita = 0
      r_dbl_Intere = 0
      r_dbl_Comisi = 0
      r_dbl_TotCuo = 0
      
      grd_MviCon_Listad.Redraw = False
      For r_int_Contad = 1 To UBound(l_arr_MViCon)
         grd_MviCon_Listad.Rows = grd_MviCon_Listad.Rows + 1
         grd_MviCon_Listad.Row = grd_MviCon_Listad.Rows - 1
         
         'Número de Cuota
         grd_MviCon_Listad.Col = 0
         grd_MviCon_Listad.Text = Format(r_int_Contad, "000")
      
         'Fecha de Vencimiento
         grd_MviCon_Listad.Col = 1
         grd_MviCon_Listad.Text = l_arr_MViCon(r_int_Contad).CuoCli_FecVct
      
         'Capital
         grd_MviCon_Listad.Col = 2
         grd_MviCon_Listad.Text = Format(l_arr_MViCon(r_int_Contad).CuoCli_Capita, "###,###,##0.00")
         r_dbl_Capita = r_dbl_Capita + CDbl(grd_MviCon_Listad)
         
         'Interes
         grd_MviCon_Listad.Col = 3
         grd_MviCon_Listad.Text = Format(l_arr_MViCon(r_int_Contad).CuoCli_Intere, "###,###,##0.00")
         r_dbl_Intere = r_dbl_Intere + CDbl(grd_MviCon_Listad)
      
         'Comisión
         grd_MviCon_Listad.Col = 4
         grd_MviCon_Listad.Text = Format(l_arr_MViCon(r_int_Contad).CuoCli_Comisi, "###,###,##0.00")
         r_dbl_Comisi = r_dbl_Comisi + CDbl(grd_MviCon_Listad)
      
         'Valor Cuota
         grd_MviCon_Listad.Col = 5
         grd_MviCon_Listad.Text = Format(l_arr_MViCon(r_int_Contad).CuoCli_ValCuo, "###,###,##0.00")
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_MviCon_Listad)
      
         'Saldo Capital
         grd_MviCon_Listad.Col = 6
         grd_MviCon_Listad.Text = Format(l_arr_MViCon(r_int_Contad).CuoCli_SalCap, "###,###,##0.00")
      Next r_int_Contad
      
      grd_MviCon_Listad.Redraw = True
      
      pnl_MViCon_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
      pnl_MViCon_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
      pnl_MViCon_Comisi.Caption = Format(r_dbl_Comisi, "###,###,##0.00") & " "
      pnl_MViCon_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
      
      If grd_MviCon_Listad.Rows > 0 Then
         Call gs_UbiIniGrid(grd_MviCon_Listad)
      End If
      
      'Cliente Concesional
      r_dbl_Capita = 0
      r_dbl_Intere = 0
      r_dbl_Comisi = 0
      r_dbl_TotCuo = 0
      
      grd_CliCon_Listad.Redraw = False
      For r_int_Contad = 1 To UBound(l_arr_CliCon)
         grd_CliCon_Listad.Rows = grd_CliCon_Listad.Rows + 1
         grd_CliCon_Listad.Row = grd_CliCon_Listad.Rows - 1
         
         'Número de Cuota
         grd_CliCon_Listad.Col = 0
         grd_CliCon_Listad.Text = Format(r_int_Contad, "000")
      
         'Fecha de Vencimiento
         grd_CliCon_Listad.Col = 1
         grd_CliCon_Listad.Text = l_arr_CliCon(r_int_Contad).CuoCli_FecVct
      
         'Capital
         grd_CliCon_Listad.Col = 2
         grd_CliCon_Listad.Text = Format(l_arr_CliCon(r_int_Contad).CuoCli_Capita, "###,###,##0.00")
         r_dbl_Capita = r_dbl_Capita + CDbl(grd_CliCon_Listad)
         
         'Interes
         grd_CliCon_Listad.Col = 3
         grd_CliCon_Listad.Text = Format(l_arr_CliCon(r_int_Contad).CuoCli_Intere, "###,###,##0.00")
         r_dbl_Intere = r_dbl_Intere + CDbl(grd_CliCon_Listad)
      
         'Valor Cuota
         grd_CliCon_Listad.Col = 4
         grd_CliCon_Listad.Text = Format(l_arr_CliCon(r_int_Contad).CuoCli_ValCuo, "###,###,##0.00")
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_CliCon_Listad)
      
         'Saldo Capital
         grd_CliCon_Listad.Col = 5
         grd_CliCon_Listad.Text = Format(l_arr_CliCon(r_int_Contad).CuoCli_SalCap, "###,###,##0.00")
      Next r_int_Contad
      
      grd_CliCon_Listad.Redraw = True
      
      pnl_CliCon_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
      pnl_CliCon_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
      pnl_CliCon_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
      
      If grd_CliCon_Listad.Rows > 0 Then
         Call gs_UbiIniGrid(grd_CliCon_Listad)
      End If
   
      If moddat_g_str_CodPrd = "003" Then
         'Cofide No Concesional
         r_dbl_Capita = 0
         r_dbl_Intere = 0
         r_dbl_Comisi = 0
         r_dbl_TotCuo = 0
         
         grd_CofNCo_Listad.Redraw = False
         For r_int_Contad = 1 To UBound(l_arr_CofNCo)
            grd_CofNCo_Listad.Rows = grd_CofNCo_Listad.Rows + 1
            grd_CofNCo_Listad.Row = grd_CofNCo_Listad.Rows - 1
            
            'Número de Cuota
            grd_CofNCo_Listad.Col = 0
            grd_CofNCo_Listad.Text = Format(r_int_Contad, "000")
         
            'Fecha de Vencimiento
            grd_CofNCo_Listad.Col = 1
            grd_CofNCo_Listad.Text = l_arr_CofNCo(r_int_Contad).CuoCli_FecVct
         
            'Capital
            grd_CofNCo_Listad.Col = 2
            grd_CofNCo_Listad.Text = Format(l_arr_CofNCo(r_int_Contad).CuoCli_Capita, "###,###,##0.00")
            r_dbl_Capita = r_dbl_Capita + CDbl(grd_CofNCo_Listad)
            
            'Interes
            grd_CofNCo_Listad.Col = 3
            grd_CofNCo_Listad.Text = Format(l_arr_CofNCo(r_int_Contad).CuoCli_Intere, "###,###,##0.00")
            r_dbl_Intere = r_dbl_Intere + CDbl(grd_CofNCo_Listad)
         
            'Comisión
            grd_CofNCo_Listad.Col = 4
            grd_CofNCo_Listad.Text = Format(l_arr_CofNCo(r_int_Contad).CuoCli_Comisi, "###,###,##0.00")
            r_dbl_Comisi = r_dbl_Comisi + CDbl(grd_CofNCo_Listad)
         
            'Valor Cuota
            grd_CofNCo_Listad.Col = 5
            grd_CofNCo_Listad.Text = Format(l_arr_CofNCo(r_int_Contad).CuoCli_ValCuo, "###,###,##0.00")
            r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_CofNCo_Listad)
         
            'Saldo Capital
            grd_CofNCo_Listad.Col = 6
            grd_CofNCo_Listad.Text = Format(l_arr_CofNCo(r_int_Contad).CuoCli_SalCap, "###,###,##0.00")
         Next r_int_Contad
         
         grd_CofNCo_Listad.Redraw = True
         
         pnl_CofNCo_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
         pnl_CofNCo_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
         pnl_CofNCo_Comisi.Caption = Format(r_dbl_Comisi, "###,###,##0.00") & " "
         pnl_CofNCo_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
         
         If grd_CofNCo_Listad.Rows > 0 Then
            Call gs_UbiIniGrid(grd_CofNCo_Listad)
         End If
      End If
   End If
End Sub


