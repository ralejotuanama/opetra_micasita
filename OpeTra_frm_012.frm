VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Gar_CreHip_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   9975
   ClientLeft      =   960
   ClientTop       =   615
   ClientWidth     =   12825
   Icon            =   "OpeTra_frm_012.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9975
   ScaleWidth      =   12825
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9975
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   12825
      _Version        =   65536
      _ExtentX        =   22622
      _ExtentY        =   17595
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   1335
         Left            =   30
         TabIndex        =   23
         Top             =   750
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   2355
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
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   1560
            TabIndex        =   24
            Top             =   390
            Width           =   11135
            _Version        =   65536
            _ExtentX        =   19641
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
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
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_NumOpe 
            Height          =   315
            Left            =   1560
            TabIndex        =   25
            Top             =   60
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "001-01-00005"
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
         End
         Begin Threed.SSPanel pnl_Situac 
            Height          =   315
            Left            =   3150
            TabIndex        =   26
            Top             =   60
            Width           =   4035
            _Version        =   65536
            _ExtentX        =   7117
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "CREDITO VIGENTE"
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
         End
         Begin Threed.SSPanel pnl_Direcc 
            Height          =   555
            Left            =   1560
            TabIndex        =   27
            Top             =   720
            Width           =   11130
            _Version        =   65536
            _ExtentX        =   19641
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
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
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   315
            Left            =   9450
            TabIndex        =   28
            Top             =   30
            Width           =   3225
            _Version        =   65536
            _ExtentX        =   5689
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
            Alignment       =   1
         End
         Begin VB.Label Label2 
            Caption         =   "Moneda:"
            Height          =   315
            Left            =   7950
            TabIndex        =   32
            Top             =   30
            Width           =   945
         End
         Begin VB.Label Label3 
            Caption         =   "Nro. de Operación:"
            Height          =   315
            Left            =   60
            TabIndex        =   31
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label Label4 
            Caption         =   "Cliente Titular:"
            Height          =   315
            Left            =   60
            TabIndex        =   30
            Top             =   390
            Width           =   1395
         End
         Begin VB.Label Label5 
            Caption         =   "Dirección Inmueble:"
            Height          =   405
            Left            =   60
            TabIndex        =   29
            Top             =   720
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   4035
         Left            =   30
         TabIndex        =   33
         Top             =   5070
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   7117
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
         Begin VB.TextBox txt_NumIns 
            Height          =   315
            Left            =   8850
            MaxLength       =   25
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   720
            Width           =   1425
         End
         Begin VB.ComboBox cmb_AtrGar 
            Height          =   315
            Left            =   8850
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   390
            Width           =   3825
         End
         Begin VB.ComboBox cmb_TipGar 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   390
            Width           =   3825
         End
         Begin VB.ComboBox cmb_SedReg 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1380
            Width           =   3825
         End
         Begin VB.ComboBox cmb_BieGar 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   60
            Width           =   3825
         End
         Begin VB.TextBox txt_NumTom 
            Height          =   315
            Left            =   1560
            MaxLength       =   25
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   2700
            Width           =   1425
         End
         Begin VB.TextBox txt_NumFoj 
            Height          =   315
            Left            =   8850
            MaxLength       =   25
            TabIndex        =   16
            Text            =   "Text1"
            Top             =   2730
            Width           =   1425
         End
         Begin VB.TextBox txt_NumLib 
            Height          =   315
            Left            =   10290
            MaxLength       =   25
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   2730
            Width           =   1425
         End
         Begin VB.TextBox txt_Observ 
            Height          =   615
            Left            =   1560
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Text            =   "OpeTra_frm_012.frx":000C
            Top             =   3360
            Width           =   11115
         End
         Begin VB.ComboBox cmb_TDoReg 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1710
            Width           =   3825
         End
         Begin VB.TextBox txt_NumPar 
            Height          =   315
            Left            =   1560
            MaxLength       =   25
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   2040
            Width           =   1425
         End
         Begin VB.TextBox txt_NumAs1 
            Height          =   315
            Left            =   8850
            MaxLength       =   25
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   2070
            Width           =   1425
         End
         Begin VB.TextBox txt_NumFic 
            Height          =   315
            Left            =   1560
            MaxLength       =   25
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   2370
            Width           =   1425
         End
         Begin VB.TextBox txt_NumAs2 
            Height          =   315
            Left            =   8850
            MaxLength       =   25
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   2400
            Width           =   1425
         End
         Begin EditLib.fpDateTime ipp_FecCon 
            Height          =   315
            Left            =   1560
            TabIndex        =   8
            Top             =   1050
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
            ThreeDInsideHighlightColor=   -2147483633
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
            ThreeDTextHighlightColor=   -2147483633
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
            ThreeDFrameColor=   -2147483633
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
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime ipp_FecIns 
            Height          =   315
            Left            =   1560
            TabIndex        =   6
            Top             =   720
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
            ThreeDInsideHighlightColor=   -2147483633
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
            ThreeDTextHighlightColor=   -2147483633
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
            ThreeDFrameColor=   -2147483633
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
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDoubleSingle ipp_MtoHip 
            Height          =   315
            Left            =   1560
            TabIndex        =   18
            Top             =   3030
            Width           =   1455
            _Version        =   196608
            _ExtentX        =   2566
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
         Begin VB.Label Label19 
            Caption         =   "Nro. Inscripción:"
            Height          =   285
            Left            =   7200
            TabIndex        =   54
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Monto Hipotecado:"
            Height          =   285
            Index           =   5
            Left            =   60
            TabIndex        =   48
            Top             =   3030
            Width           =   1395
         End
         Begin VB.Label Label18 
            Caption         =   "Atributo Garantía:"
            Height          =   315
            Left            =   7200
            TabIndex        =   47
            Top             =   390
            Width           =   1305
         End
         Begin VB.Label Label17 
            Caption         =   "Tipo de Garantía:"
            Height          =   315
            Left            =   60
            TabIndex        =   46
            Top             =   390
            Width           =   1305
         End
         Begin VB.Label Label10 
            Caption         =   "Sede Registral:"
            Height          =   315
            Left            =   60
            TabIndex        =   45
            Top             =   1380
            Width           =   1305
         End
         Begin VB.Label Label7 
            Caption         =   "Fecha Inscripción:"
            Height          =   315
            Left            =   60
            TabIndex        =   44
            Top             =   720
            Width           =   1425
         End
         Begin VB.Label Label9 
            Caption         =   "Bien en Garantía:"
            Height          =   315
            Left            =   60
            TabIndex        =   43
            Top             =   60
            Width           =   1305
         End
         Begin VB.Label Label14 
            Caption         =   "Nro. Tomo:"
            Height          =   285
            Left            =   60
            TabIndex        =   42
            Top             =   2700
            Width           =   1485
         End
         Begin VB.Label Label15 
            Caption         =   "Fecha Constitucion:"
            Height          =   315
            Left            =   60
            TabIndex        =   41
            Top             =   1050
            Width           =   1425
         End
         Begin VB.Label Label21 
            Caption         =   "Nro. Foja / Nro. Libro:"
            Height          =   285
            Left            =   7200
            TabIndex        =   40
            Top             =   2700
            Width           =   1545
         End
         Begin VB.Label Label25 
            Caption         =   "Observaciones:"
            Height          =   315
            Left            =   60
            TabIndex        =   39
            Top             =   3360
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Doc. Registral:"
            Height          =   315
            Left            =   60
            TabIndex        =   38
            Top             =   1710
            Width           =   1485
         End
         Begin VB.Label Label8 
            Caption         =   "Partida Electrónica:"
            Height          =   285
            Left            =   60
            TabIndex        =   37
            Top             =   2040
            Width           =   1485
         End
         Begin VB.Label Label12 
            Caption         =   "Asiento:"
            Height          =   285
            Left            =   7200
            TabIndex        =   36
            Top             =   2040
            Width           =   765
         End
         Begin VB.Label Label13 
            Caption         =   "Ficha Registral:"
            Height          =   285
            Left            =   60
            TabIndex        =   35
            Top             =   2370
            Width           =   1485
         End
         Begin VB.Label Label16 
            Caption         =   "Asiento:"
            Height          =   285
            Left            =   7200
            TabIndex        =   34
            Top             =   2370
            Width           =   765
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   765
         Left            =   30
         TabIndex        =   49
         Top             =   2130
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
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
            Left            =   12030
            Picture         =   "OpeTra_frm_012.frx":0010
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   11340
            Picture         =   "OpeTra_frm_012.frx":0452
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   52
         Top             =   30
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
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
            TabIndex        =   53
            Top             =   60
            Width           =   6585
            _Version        =   65536
            _ExtentX        =   11615
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Gestión de Garantías - Registro de Hipotecas"
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
            Picture         =   "OpeTra_frm_012.frx":0894
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1275
         Left            =   30
         TabIndex        =   55
         Top             =   2940
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   2249
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
            Height          =   885
            Left            =   30
            TabIndex        =   56
            Top             =   360
            Width           =   12645
            _ExtentX        =   22304
            _ExtentY        =   1561
            _Version        =   393216
            Rows            =   21
            Cols            =   17
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel5 
            Height          =   285
            Left            =   60
            TabIndex        =   57
            Top             =   60
            Width           =   2445
            _Version        =   65536
            _ExtentX        =   4313
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Bien en Garantía"
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
            Left            =   8100
            TabIndex        =   58
            Top             =   60
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Inscripción"
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   285
            Left            =   10740
            TabIndex        =   59
            Top             =   60
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Importe"
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
         Begin Threed.SSPanel SSPanel14 
            Height          =   285
            Left            =   9420
            TabIndex        =   60
            Top             =   60
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Constitución"
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
            Left            =   2490
            TabIndex        =   61
            Top             =   60
            Width           =   5625
            _Version        =   65536
            _ExtentX        =   9922
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Documento Registral"
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
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   765
         Left            =   30
         TabIndex        =   62
         Top             =   9150
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
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
         Begin VB.CommandButton cmd_Cancel 
            Height          =   675
            Left            =   12000
            Picture         =   "OpeTra_frm_012.frx":0B9E
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Acepta 
            Height          =   675
            Left            =   11310
            Picture         =   "OpeTra_frm_012.frx":0EA8
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Aceptar Datos"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel12 
         Height          =   765
         Left            =   30
         TabIndex        =   63
         Top             =   4260
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
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
         Begin VB.CommandButton cmd_Borrar 
            Height          =   675
            Left            =   12030
            Picture         =   "OpeTra_frm_012.frx":11B2
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   675
            Left            =   10650
            Picture         =   "OpeTra_frm_012.frx":14BC
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   675
            Left            =   11340
            Picture         =   "OpeTra_frm_012.frx":17C6
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Modificar Registro"
            Top             =   30
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frm_Gar_CreHip_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_TipGar()   As moddat_tpo_Genera
Dim l_arr_AtrGar()   As moddat_tpo_Genera

Private Sub cmb_AtrGar_Click()
   Call gs_SetFocus(ipp_FecIns)
End Sub

Private Sub cmb_AtrGar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_AtrGar_Click
   End If
End Sub

Private Sub cmb_BieGar_Click()
   Call gs_SetFocus(cmb_TipGar)
End Sub

Private Sub cmb_BieGar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_BieGar_Click
   End If
End Sub

Private Sub cmb_SedReg_Click()
   Call gs_SetFocus(cmb_TDoReg)
End Sub

Private Sub cmb_SedReg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_SedReg_Click
   End If
End Sub

Private Sub cmb_TDoReg_Click()
   If cmb_TDoReg.ListIndex > -1 Then
      Select Case cmb_TDoReg.ItemData(cmb_TDoReg.ListIndex)
         Case 1
            txt_NumPar.Enabled = True
            txt_NumAs1.Enabled = True
            
            txt_NumFic.Enabled = False
            txt_NumAs2.Enabled = False
            
            txt_NumTom.Enabled = False
            txt_NumFoj.Enabled = False
            txt_NumLib.Enabled = False
            
            txt_NumFic.Text = ""
            txt_NumAs2.Text = ""
            txt_NumTom.Text = ""
            txt_NumFoj.Text = ""
            txt_NumLib.Text = ""
            
            Call gs_SetFocus(txt_NumPar)
            
         Case 2
            txt_NumPar.Enabled = False
            txt_NumAs1.Enabled = False
            
            txt_NumFic.Enabled = True
            txt_NumAs2.Enabled = True
            
            txt_NumTom.Enabled = False
            txt_NumFoj.Enabled = False
            txt_NumLib.Enabled = False
            
            txt_NumPar.Text = ""
            txt_NumAs1.Text = ""
            txt_NumTom.Text = ""
            txt_NumFoj.Text = ""
            txt_NumLib.Text = ""
            
            Call gs_SetFocus(txt_NumFic)
         Case 3
            txt_NumPar.Enabled = False
            txt_NumAs1.Enabled = False
            
            txt_NumFic.Enabled = False
            txt_NumAs2.Enabled = False
            
            txt_NumTom.Enabled = True
            txt_NumFoj.Enabled = True
            txt_NumLib.Enabled = True
         
            txt_NumPar.Text = ""
            txt_NumAs1.Text = ""
            txt_NumFic.Text = ""
            txt_NumAs2.Text = ""
      
            Call gs_SetFocus(txt_NumTom)
      End Select
   End If
End Sub

Private Sub cmb_TDoReg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TDoReg_Click
   End If
End Sub

Private Sub cmb_TipGar_Click()
   Call gs_SetFocus(cmb_AtrGar)
   
   If cmb_TipGar.ListIndex > -1 Then
      Call moddat_gs_Carga_AtrGar(cmb_AtrGar, l_arr_AtrGar(), l_arr_TipGar(cmb_TipGar.ListIndex + 1).Genera_Codigo, 2)
   Else
      cmb_AtrGar.Clear
   End If
End Sub

Private Sub cmb_TipGar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipGar_Click
   End If
End Sub

Private Sub cmd_Acepta_Click()
   Dim r_int_Contad     As Integer
   Dim r_int_FlgEnc     As Integer
   Dim r_str_ParFic     As String
   Dim r_str_NumAsi     As String

   If cmb_BieGar.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Bien en Garantía.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_BieGar)
      Exit Sub
   End If
   
   If cmb_TipGar.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Garantía.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipGar)
      Exit Sub
   End If
   
   If cmb_AtrGar.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Atributo de la Garantía.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_AtrGar)
      Exit Sub
   End If
   
   If Len(Trim(txt_NumIns.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Inscripción.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumIns)
      Exit Sub
   End If
   
   If CDate(ipp_FecCon.Text) < CDate(ipp_FecIns.Text) Then
      MsgBox "La Fecha de Constitución no debe ser menor a la Fecha de Inscripción.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecCon)
      Exit Sub
   End If
   
   If cmb_SedReg.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Sede Registral.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_SedReg)
      Exit Sub
   End If
   
   If cmb_TDoReg.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento Registral.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TDoReg)
      Exit Sub
   End If
   
   r_str_ParFic = ""
   r_str_NumAsi = ""
      
   Select Case cmb_TDoReg.ItemData(cmb_TDoReg.ListIndex)
      Case 1
         If Len(Trim(txt_NumPar.Text)) = 0 Then
            MsgBox "Debe ingresar el Número de Partida Electrónica.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumPar)
            Exit Sub
         End If
         
         If Len(Trim(txt_NumAs1.Text)) = 0 Then
            MsgBox "Debe ingresar el Número de Asiento.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumAs1)
            Exit Sub
         End If
         
         r_str_ParFic = txt_NumPar.Text
         r_str_NumAsi = txt_NumAs1.Text
         
      Case 2
         If Len(Trim(txt_NumFic.Text)) = 0 Then
            MsgBox "Debe ingresar el Número de Ficha Registral.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumFic)
            Exit Sub
         End If
         
         If Len(Trim(txt_NumAs2.Text)) = 0 Then
            MsgBox "Debe ingresar el Número de Asiento.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumAs2)
            Exit Sub
         End If
         
         r_str_ParFic = txt_NumFic.Text
         r_str_NumAsi = txt_NumAs2.Text
      
      Case 3
         If Len(Trim(txt_NumTom.Text)) = 0 Then
            MsgBox "Debe ingresar el Número de Tomo.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumTom)
            Exit Sub
         End If
      
         If Len(Trim(txt_NumFoj.Text)) = 0 Then
            MsgBox "Debe ingresar el Número de Fojas.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumFoj)
            Exit Sub
         End If
      
         If Len(Trim(txt_NumLib.Text)) = 0 Then
            MsgBox "Debe ingresar el Número de Libro.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumLib)
            Exit Sub
         End If
   End Select

   If CDbl(ipp_MtoHip.Text) = 0 Then
      MsgBox "Debe ingresar el Monto Hipotecado.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_MtoHip)
      Exit Sub
   End If
   
   'Validando que no haya sido ingresado el mismo bien
   If moddat_g_int_FlgGrb = 1 Then
      grd_Listad.Col = 5
      r_int_FlgEnc = 0
      
      For r_int_Contad = 0 To grd_Listad.Rows - 1
         grd_Listad.Row = r_int_Contad
         If CInt(grd_Listad.Text) = cmb_BieGar.ItemData(cmb_BieGar.ListIndex) Then
            r_int_FlgEnc = 1
            Exit For
         End If
      Next r_int_Contad
      
      If grd_Listad.Rows > 0 Then
         Call gs_RefrescaGrid(grd_Listad)
      End If
      
      If r_int_FlgEnc = 1 Then
         MsgBox "El Bien en Garantía ya ha sido registrado.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_BieGar)
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   If moddat_g_int_FlgGrb = 1 Then
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
   End If

   grd_Listad.Col = 0
   grd_Listad.Text = cmb_BieGar.Text
   
   grd_Listad.Col = 1
   
   Select Case cmb_TDoReg.ItemData(cmb_TDoReg.ListIndex)
      Case 1:  grd_Listad.Text = Trim(cmb_TDoReg.Text) & " NRO. " & txt_NumPar.Text & " ASIENTO NRO. " & txt_NumAs1.Text
      Case 2:  grd_Listad.Text = Trim(cmb_TDoReg.Text) & " NRO. " & txt_NumFic.Text & " ASIENTO NRO. " & txt_NumAs2.Text
      Case 3:  grd_Listad.Text = "TOMO NRO. " & txt_NumTom.Text & " FOJA NRO. " & txt_NumFoj.Text & " LIBRO NRO. " & txt_NumLib.Text
   End Select
   
   grd_Listad.Col = 2
   grd_Listad.Text = ipp_FecIns.Text
   
   grd_Listad.Col = 3
   grd_Listad.Text = ipp_FecCon.Text
   
   grd_Listad.Col = 4
   grd_Listad.Text = ipp_MtoHip.Text
   
   grd_Listad.Col = 5
   grd_Listad.Text = cmb_BieGar.ItemData(cmb_BieGar.ListIndex)
   
   grd_Listad.Col = 6
   grd_Listad.Text = l_arr_TipGar(cmb_TipGar.ListIndex + 1).Genera_Codigo
   
   grd_Listad.Col = 7
   grd_Listad.Text = l_arr_AtrGar(cmb_AtrGar.ListIndex + 1).Genera_Codigo
   
   grd_Listad.Col = 8
   grd_Listad.Text = txt_NumIns.Text
   
   grd_Listad.Col = 9
   grd_Listad.Text = cmb_SedReg.ItemData(cmb_SedReg.ListIndex)
   
   grd_Listad.Col = 10
   grd_Listad.Text = cmb_TDoReg.ItemData(cmb_TDoReg.ListIndex)
   
   
   Select Case cmb_TDoReg.ItemData(cmb_TDoReg.ListIndex)
      Case 1
         grd_Listad.Col = 11
         grd_Listad.Text = txt_NumPar.Text
         
         grd_Listad.Col = 12
         grd_Listad.Text = txt_NumAs1.Text
         
      Case 2
         grd_Listad.Col = 11
         grd_Listad.Text = txt_NumFic.Text
         
         grd_Listad.Col = 12
         grd_Listad.Text = txt_NumAs2.Text
      
      Case 3
         grd_Listad.Col = 13
         grd_Listad.Text = txt_NumTom.Text
         
         grd_Listad.Col = 14
         grd_Listad.Text = txt_NumFoj.Text
   
         grd_Listad.Col = 15
         grd_Listad.Text = txt_NumLib.Text
   End Select
   
   grd_Listad.Col = 16
   grd_Listad.Text = txt_Observ.Text
   
   Call gs_RefrescaGrid(grd_Listad)
   Call cmd_Cancel_Click
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   
   Call fs_Activa(False)
   Call gs_SetFocus(cmb_BieGar)
End Sub

Private Sub cmd_Borrar_Click()
   If MsgBox("¿Está seguro de eliminar la actividad?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   If grd_Listad.Rows = 1 Then
      grd_Listad.Rows = 0
   Else
      grd_Listad.RemoveItem grd_Listad.Row
   End If
   
   If grd_Listad.Rows = 0 Then
      cmd_Borrar.Enabled = False
      cmd_Editar.Enabled = False
   End If
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_Limpia
   Call fs_Activa(True)
   
   If grd_Listad.Rows = 0 Then
      cmd_Editar.Enabled = False
      cmd_Borrar.Enabled = False
      
      Call gs_SetFocus(cmd_Agrega)
   Else
      Call gs_RefrescaGrid(grd_Listad)
      Call gs_SetFocus(grd_Listad)
   End If
End Sub

Private Sub cmd_Editar_Click()
   Call fs_Activa(False)

   grd_Listad.Col = 5
   Call gs_BuscarCombo_Item(cmb_BieGar, CInt(grd_Listad.Text))
   
   grd_Listad.Col = 6
   cmb_TipGar.ListIndex = gf_Busca_Arregl(l_arr_TipGar, grd_Listad.Text) - 1
   
   Call moddat_gs_Carga_AtrGar(cmb_AtrGar, l_arr_AtrGar(), l_arr_TipGar(cmb_TipGar.ListIndex + 1).Genera_Codigo, 2)
   
   grd_Listad.Col = 7
   cmb_AtrGar.ListIndex = gf_Busca_Arregl(l_arr_AtrGar, grd_Listad.Text) - 1
   
   grd_Listad.Col = 2
   ipp_FecIns.Text = grd_Listad.Text
   
   grd_Listad.Col = 8
   txt_NumIns.Text = grd_Listad.Text
   
   grd_Listad.Col = 3
   ipp_FecCon.Text = grd_Listad.Text
   
   grd_Listad.Col = 9
   Call gs_BuscarCombo_Item(cmb_SedReg, CInt(grd_Listad.Text))
   
   grd_Listad.Col = 10
   Call gs_BuscarCombo_Item(cmb_TDoReg, CInt(grd_Listad.Text))
   
   Call cmb_TDoReg_Click
   
   Select Case cmb_TDoReg.ItemData(cmb_TDoReg.ListIndex)
      Case 1
         grd_Listad.Col = 11
         txt_NumPar.Text = grd_Listad.Text
         
         grd_Listad.Col = 12
         txt_NumAs1.Text = grd_Listad.Text
         
      Case 2
         grd_Listad.Col = 11
         txt_NumFic.Text = grd_Listad.Text
         
         grd_Listad.Col = 12
         txt_NumAs2.Text = grd_Listad.Text
      
      Case 3
         grd_Listad.Col = 13
         txt_NumTom.Text = grd_Listad.Text
         
         grd_Listad.Col = 14
         txt_NumFoj.Text = grd_Listad.Text
         
         grd_Listad.Col = 15
         txt_NumLib.Text = grd_Listad.Text
   End Select
   
   grd_Listad.Col = 4
   ipp_MtoHip.Value = CDbl(grd_Listad.Text)
   
   grd_Listad.Col = 16
   txt_Observ.Text = grd_Listad.Text
   
   moddat_g_int_FlgGrb = 2
   
   cmb_BieGar.Enabled = False
   Call gs_SetFocus(cmb_TipGar)
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_int_Contad     As Integer
   Dim r_int_BieGar     As Integer
   Dim r_int_TDoReg     As Integer
   Dim r_str_TipGar     As String
   Dim r_str_AtrGar     As String
   Dim r_str_FecIns     As String
   Dim r_str_NumIns     As String
   Dim r_str_FecCon     As String
   Dim r_str_SedReg     As String
   Dim r_str_ParFic     As String
   Dim r_str_NumAsi     As String
   Dim r_str_NumTom     As String
   Dim r_str_NumFoj     As String
   Dim r_str_NumLib     As String
   Dim r_str_MtoHip     As String
   Dim r_str_Observ     As String
   Dim r_str_SitCre     As String
   Dim r_str_SitAnt     As String
   Dim r_str_Descri     As String
   Dim r_str_Operac     As String
   
   If grd_Listad.Rows = 0 Then
      MsgBox "Debe registrar la(s) Hipoteca(s).", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Agrega)
      Exit Sub
   End If

   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      grd_Listad.Row = r_int_Contad
      
      grd_Listad.Col = 1
      r_str_Descri = grd_Listad.Text
      
      grd_Listad.Col = 5
      r_int_BieGar = CInt(grd_Listad.Text)
   
      grd_Listad.Col = 6
      r_str_TipGar = Format(grd_Listad.Text, "000000")
   
      grd_Listad.Col = 7
      r_str_AtrGar = Format(grd_Listad.Text, "000000")
   
      grd_Listad.Col = 2
      r_str_FecIns = Format(CDate(grd_Listad.Text), "yyyymmdd")
   
      grd_Listad.Col = 8
      r_str_NumIns = grd_Listad.Text
   
      grd_Listad.Col = 3
      r_str_FecCon = Format(CDate(grd_Listad.Text), "yyyymmdd")
   
      grd_Listad.Col = 9
      r_str_SedReg = Format(grd_Listad.Text, "0000")
   
      grd_Listad.Col = 10
      r_int_TDoReg = CInt(grd_Listad.Text)
   
      Select Case r_int_TDoReg
         Case 1, 2
            grd_Listad.Col = 11
            r_str_ParFic = grd_Listad.Text
         
            grd_Listad.Col = 12
            r_str_NumAsi = grd_Listad.Text
         
         Case 3
            grd_Listad.Col = 13
            r_str_NumTom = grd_Listad.Text
         
            grd_Listad.Col = 14
            r_str_NumFoj = grd_Listad.Text
         
            grd_Listad.Col = 15
            r_str_NumLib = grd_Listad.Text
      End Select
   
      grd_Listad.Col = 4
      r_str_MtoHip = CStr(CDbl(grd_Listad.Text))
   
      grd_Listad.Col = 16
      r_str_Observ = grd_Listad.Text
      
      
      'Buscando Información en Maestro de Creditos
      g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
      g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         r_str_SitCre = CStr(g_rst_Genera!HIPMAE_SITCRE)
         r_str_SitAnt = CStr(g_rst_Genera!HIPMAE_SITANT)
      End If
   
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      
      
      r_str_Operac = moddat_gf_Consulta_Operac(moddat_g_str_CodPrd, "41")
      
      'Registrando Garantía
      moddat_g_int_FlgGOK = False
      moddat_g_int_CntErr = 0
      
      Do While moddat_g_int_FlgGOK = False
         g_str_Parame = "USP_CRE_HIPGAR_CREA ("
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
         g_str_Parame = g_str_Parame & CStr(r_int_BieGar) & ", "
         g_str_Parame = g_str_Parame & "'" & r_str_TipGar & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_AtrGar & "', "
         g_str_Parame = g_str_Parame & r_str_FecIns & ", "
         g_str_Parame = g_str_Parame & "'" & r_str_NumIns & "', "
         g_str_Parame = g_str_Parame & r_str_FecCon & ", "
         g_str_Parame = g_str_Parame & "'" & r_str_SedReg & "', "
         g_str_Parame = g_str_Parame & CStr(r_int_TDoReg) & ", "
         g_str_Parame = g_str_Parame & "'" & r_str_ParFic & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_NumAsi & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_NumTom & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_NumFoj & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_NumLib & "', "
         g_str_Parame = g_str_Parame & "2, "
         g_str_Parame = g_str_Parame & r_str_MtoHip & ", "
         g_str_Parame = g_str_Parame & "'" & r_str_Observ & "', "
         g_str_Parame = g_str_Parame & "'" & CStr(moddat_g_int_TipDoc) & Trim(moddat_g_str_NumDoc) & "', "
               
         'Datos de Auditoria
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                              'Código Sucursal
            
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
         End If
   
         If moddat_g_int_CntErr = 6 Then
            If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
               Exit Sub
            Else
               moddat_g_int_CntErr = 0
            End If
         End If
      Loop
   
   
   
      'Registrando Archivo para Contabilización
      moddat_g_int_FlgGOK = False
      moddat_g_int_CntErr = 0
      
      Do While moddat_g_int_FlgGOK = False
         g_str_Parame = "USP_CREDITO_MOV_CONTAB_INSERTA ("
         g_str_Parame = g_str_Parame & "'" & r_str_SitCre & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_SitAnt & "', "
         g_str_Parame = g_str_Parame & "1,"
         g_str_Parame = g_str_Parame & r_str_FecCon & ", "
         g_str_Parame = g_str_Parame & "'" & Left(r_str_Operac, 3) & "002', "
         g_str_Parame = g_str_Parame & "'002', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
         g_str_Parame = g_str_Parame & r_str_MtoHip & ", "
         
         g_str_Parame = g_str_Parame & "'" & Mid(r_str_Descri, 1, 50) & "')"
            
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
         End If
   
         If moddat_g_int_CntErr = 6 Then
            If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
               Exit Sub
            Else
               moddat_g_int_CntErr = 0
            End If
         End If
      Loop
   Next r_int_Contad
   
   MsgBox "Se registraron las Hipotecas correctamente.", vbInformation, modgen_g_str_NomPlt
   
   moddat_g_int_FlgAct = 2
   
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_NumOpe.Caption = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
   pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   pnl_Moneda.Caption = moddat_g_str_Moneda
   pnl_Situac.Caption = moddat_g_str_Situac
   
   pnl_Direcc.Caption = moddat_g_str_Direcc & Chr(10) & Chr(13) & moddat_g_str_Distri
   
   Call fs_Inicia
   
   Call fs_Activa(True)
   Call fs_Limpia
   
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub ipp_FecCon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_SedReg)
   End If
End Sub

Private Sub ipp_FecIns_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumIns)
   End If
End Sub

Private Sub ipp_MtoHip_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Observ)
   End If
End Sub

Private Sub txt_NumIns_GotFocus()
   Call gs_SelecTodo(txt_NumIns)
End Sub

Private Sub txt_NumIns_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecCon)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumPar_GotFocus()
   Call gs_SelecTodo(txt_NumPar)
End Sub

Private Sub txt_NumPar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumAs1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumAs1_GotFocus()
   Call gs_SelecTodo(txt_NumAs1)
End Sub

Private Sub txt_NumAs1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoHip)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumFic_GotFocus()
   Call gs_SelecTodo(txt_NumFic)
End Sub

Private Sub txt_NumFic_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumAs2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumAs2_GotFocus()
   Call gs_SelecTodo(txt_NumAs2)
End Sub

Private Sub txt_NumAs2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoHip)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumTom_GotFocus()
   Call gs_SelecTodo(txt_NumTom)
End Sub

Private Sub txt_NumTom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumFoj)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*")
   End If
End Sub

Private Sub txt_NumFoj_GotFocus()
   Call gs_SelecTodo(txt_NumFoj)
End Sub

Private Sub txt_NumFoj_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumLib)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*")
   End If
End Sub

Private Sub txt_NumLib_GotFocus()
   Call gs_SelecTodo(txt_NumLib)
End Sub

Private Sub txt_NumLib_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoHip)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*")
   End If
End Sub

Private Sub txt_Observ_GotFocus()
   Call gs_SelecTodo(txt_Observ)
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Acepta)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_TDoReg, 1, "026")
   Call moddat_gs_Carga_LisIte_Combo(cmb_BieGar, 1, "030")
   Call moddat_gs_Carga_LisIte_Combo(cmb_SedReg, 1, "511")
   
   Call moddat_gs_Carga_TipGar(cmb_TipGar, l_arr_TipGar())
   
   grd_Listad.ColWidth(0) = 2430
   grd_Listad.ColWidth(1) = 5595
   grd_Listad.ColWidth(2) = 1305
   grd_Listad.ColWidth(3) = 1310
   grd_Listad.ColWidth(4) = 1575
   grd_Listad.ColWidth(5) = 0
   grd_Listad.ColWidth(6) = 0
   grd_Listad.ColWidth(7) = 0
   grd_Listad.ColWidth(8) = 0
   grd_Listad.ColWidth(9) = 0
   grd_Listad.ColWidth(10) = 0
   grd_Listad.ColWidth(11) = 0
   grd_Listad.ColWidth(12) = 0
   grd_Listad.ColWidth(13) = 0
   grd_Listad.ColWidth(14) = 0
   grd_Listad.ColWidth(15) = 0
   grd_Listad.ColWidth(16) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignRightCenter

   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_Limpia()
   cmb_BieGar.ListIndex = -1
   cmb_TipGar.ListIndex = -1
   cmb_AtrGar.Clear
   ipp_FecIns.Text = Format(Date, "dd/mm/yyyy")
   txt_NumIns.Text = ""
   ipp_FecCon.Text = Format(Date, "dd/mm/yyyy")
   cmb_SedReg.ListIndex = -1
   cmb_TDoReg.ListIndex = -1
   txt_NumPar.Text = ""
   txt_NumAs1.Text = ""
   txt_NumFic.Text = ""
   txt_NumAs2.Text = ""
   txt_NumTom.Text = ""
   txt_NumFoj.Text = ""
   txt_NumLib.Text = ""
   ipp_MtoHip.Value = 0
   txt_Observ.Text = ""
      
   txt_NumPar.Enabled = False
   txt_NumAs1.Enabled = False
   txt_NumFic.Enabled = False
   txt_NumAs2.Enabled = False
   txt_NumTom.Enabled = False
   txt_NumFoj.Enabled = False
   txt_NumLib.Enabled = False
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   grd_Listad.Enabled = p_Activa
   cmd_Agrega.Enabled = p_Activa
   cmd_Editar.Enabled = p_Activa
   cmd_Borrar.Enabled = p_Activa
   cmd_Grabar.Enabled = p_Activa
   
   cmb_BieGar.Enabled = Not p_Activa
   cmb_TipGar.Enabled = Not p_Activa
   cmb_AtrGar.Enabled = Not p_Activa
   ipp_FecIns.Enabled = Not p_Activa
   txt_NumIns.Enabled = Not p_Activa
   ipp_FecCon.Enabled = Not p_Activa
   cmb_SedReg.Enabled = Not p_Activa
   cmb_TDoReg.Enabled = Not p_Activa
   
   txt_NumPar.Enabled = False
   txt_NumAs1.Enabled = False
   txt_NumFic.Enabled = False
   txt_NumAs2.Enabled = False
   txt_NumTom.Enabled = False
   txt_NumFoj.Enabled = False
   txt_NumLib.Enabled = False
   
   ipp_MtoHip.Enabled = Not p_Activa
   txt_Observ.Enabled = Not p_Activa
   
   cmd_Acepta.Enabled = Not p_Activa
   cmd_Cancel.Enabled = Not p_Activa
End Sub


