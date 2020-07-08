VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Con_PreTot_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   9930
   ClientLeft      =   435
   ClientTop       =   3090
   ClientWidth     =   11685
   Icon            =   "OpeTra_frm_325.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel111 
      Height          =   10065
      Left            =   0
      TabIndex        =   7
      Top             =   -120
      Width           =   11685
      _Version        =   65536
      _ExtentX        =   20611
      _ExtentY        =   17754
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   1860
         Left            =   30
         TabIndex        =   15
         Top             =   8160
         Width           =   11625
         _Version        =   65536
         _ExtentX        =   20505
         _ExtentY        =   3281
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
         Begin VB.ComboBox cmb_MotPpg 
            Height          =   315
            Left            =   1470
            TabIndex        =   18
            Text            =   "MOTIVO DEL PREPAGO"
            Top             =   360
            Width           =   10080
         End
         Begin VB.TextBox txt_ObsPpg 
            Height          =   1065
            Left            =   1470
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   20
            Top             =   690
            Width           =   10065
         End
         Begin VB.Label Label9 
            Caption         =   "Comentarios"
            Height          =   315
            Left            =   120
            TabIndex        =   40
            Top             =   720
            Width           =   1290
         End
         Begin VB.Label Label31 
            Caption         =   "Motivo Prepago"
            Height          =   315
            Left            =   120
            TabIndex        =   39
            Top             =   390
            Width           =   1290
         End
         Begin VB.Label Label8 
            Caption         =   "Observaciones"
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
            TabIndex        =   38
            Top             =   60
            Width           =   2085
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   3380
         Left            =   30
         TabIndex        =   17
         Top             =   4740
         Width           =   11625
         _Version        =   65536
         _ExtentX        =   20505
         _ExtentY        =   5962
         _StockProps     =   15
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
         Begin EditLib.fpDateTime ipp_FecPre 
            Height          =   315
            Left            =   1350
            TabIndex        =   8
            Top             =   390
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
         Begin Threed.SSPanel pnl_SaldoTNC1 
            Height          =   315
            Left            =   7290
            TabIndex        =   19
            Top             =   390
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
         Begin Threed.SSPanel pnl_SaldoTC1 
            Height          =   315
            Left            =   10080
            TabIndex        =   21
            Top             =   390
            Width           =   1440
            _Version        =   65536
            _ExtentX        =   2540
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
         Begin EditLib.fpDoubleSingle txt_MontoITF 
            Height          =   315
            Left            =   4470
            TabIndex        =   16
            Top             =   2910
            Width           =   1200
            _Version        =   196608
            _ExtentX        =   2117
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
            MaxValue        =   "9000"
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
         Begin Threed.SSPanel pnl_MtoApl 
            Height          =   315
            Left            =   1590
            TabIndex        =   22
            Top             =   2910
            Width           =   1440
            _Version        =   65536
            _ExtentX        =   2540
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
         Begin Threed.SSPanel pnl_Val_AsgInm 
            Height          =   315
            Left            =   4230
            TabIndex        =   23
            Top             =   390
            Width           =   1440
            _Version        =   65536
            _ExtentX        =   2540
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "250,000.00 "
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
         Begin Threed.SSPanel pnl_MtoPrepagar 
            Height          =   315
            Left            =   7590
            TabIndex        =   36
            Top             =   2910
            Width           =   1260
            _Version        =   65536
            _ExtentX        =   2222
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
         Begin TabDlg.SSTab Tab_Deuda 
            Height          =   1920
            Left            =   90
            TabIndex        =   41
            Top             =   840
            Width           =   11475
            _ExtentX        =   20241
            _ExtentY        =   3387
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Deuda Vigente"
            TabPicture(0)   =   "OpeTra_frm_325.frx":000C
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label22"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label23"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label19"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Label1"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "Label24"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "Label25"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "Label7"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "Label6"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "Label17"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "Label13"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "Label12"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "Label10"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "Label14"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "Label15"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "txt_IntLeg"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "pnl_DevBBP"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "pnl_IntPbpPerdido"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).Control(17)=   "pnl_CapPbpPerdido"
            Tab(0).Control(17).Enabled=   0   'False
            Tab(0).Control(18)=   "pnl_DeuPen"
            Tab(0).Control(18).Enabled=   0   'False
            Tab(0).Control(19)=   "txt_SegDes"
            Tab(0).Control(19).Enabled=   0   'False
            Tab(0).Control(20)=   "txt_InteresTNC"
            Tab(0).Control(20).Enabled=   0   'False
            Tab(0).Control(21)=   "txt_InteresTC"
            Tab(0).Control(21).Enabled=   0   'False
            Tab(0).Control(22)=   "txt_MontoPortes"
            Tab(0).Control(22).Enabled=   0   'False
            Tab(0).Control(23)=   "pnl_DiasTC"
            Tab(0).Control(23).Enabled=   0   'False
            Tab(0).Control(24)=   "pnl_DiasTNC"
            Tab(0).Control(24).Enabled=   0   'False
            Tab(0).Control(25)=   "txt_SegInm"
            Tab(0).Control(25).Enabled=   0   'False
            Tab(0).Control(26)=   "pnl_UltPagTC"
            Tab(0).Control(26).Enabled=   0   'False
            Tab(0).Control(27)=   "pnl_UltPagTNC"
            Tab(0).Control(27).Enabled=   0   'False
            Tab(0).ControlCount=   28
            TabCaption(1)   =   "Deuda Pendiente (*)"
            TabPicture(1)   =   "OpeTra_frm_325.frx":0028
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_DeuPen"
            Tab(1).ControlCount=   1
            Begin Threed.SSPanel pnl_UltPagTNC 
               Height          =   315
               Left            =   1500
               TabIndex        =   42
               Top             =   450
               Width           =   1440
               _Version        =   65536
               _ExtentX        =   2540
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "31/12/2011"
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
            End
            Begin Threed.SSPanel pnl_UltPagTC 
               Height          =   315
               Left            =   1500
               TabIndex        =   43
               Top             =   795
               Width           =   1440
               _Version        =   65536
               _ExtentX        =   2540
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "31/12/2011"
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
            End
            Begin EditLib.fpDoubleSingle txt_SegInm 
               Height          =   315
               Left            =   4380
               TabIndex        =   12
               Top             =   1125
               Width           =   1200
               _Version        =   196608
               _ExtentX        =   2117
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
            Begin Threed.SSPanel pnl_DiasTNC 
               Height          =   315
               Left            =   4380
               TabIndex        =   44
               Top             =   450
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "30 "
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
            Begin Threed.SSPanel pnl_DiasTC 
               Height          =   315
               Left            =   4380
               TabIndex        =   45
               Top             =   795
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "180 "
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
            Begin EditLib.fpDoubleSingle txt_MontoPortes 
               Height          =   315
               Left            =   7470
               TabIndex        =   13
               Top             =   1125
               Width           =   1260
               _Version        =   196608
               _ExtentX        =   2222
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
            Begin EditLib.fpDoubleSingle txt_InteresTC 
               Height          =   315
               Left            =   7470
               TabIndex        =   10
               Top             =   795
               Width           =   1260
               _Version        =   196608
               _ExtentX        =   2222
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
            Begin EditLib.fpDoubleSingle txt_InteresTNC 
               Height          =   315
               Left            =   7470
               TabIndex        =   9
               Top             =   450
               Width           =   1260
               _Version        =   196608
               _ExtentX        =   2222
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
            Begin EditLib.fpDoubleSingle txt_SegDes 
               Height          =   315
               Left            =   1500
               TabIndex        =   11
               Top             =   1125
               Width           =   1440
               _Version        =   196608
               _ExtentX        =   2540
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
            Begin MSFlexGridLib.MSFlexGrid grd_DeuPen 
               Height          =   1245
               Left            =   -74880
               TabIndex        =   55
               Top             =   420
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   2196
               _Version        =   393216
               Rows            =   4
               Cols            =   16
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               SelectionMode   =   1
            End
            Begin Threed.SSPanel pnl_DeuPen 
               Height          =   315
               Left            =   10155
               TabIndex        =   56
               Top             =   1110
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
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
            Begin Threed.SSPanel pnl_CapPbpPerdido 
               Height          =   315
               Left            =   10155
               TabIndex        =   58
               Top             =   450
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
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
            Begin Threed.SSPanel pnl_IntPbpPerdido 
               Height          =   315
               Left            =   10155
               TabIndex        =   59
               Top             =   780
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
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
            Begin Threed.SSPanel pnl_DevBBP 
               Height          =   315
               Left            =   1500
               TabIndex        =   62
               Top             =   1450
               Width           =   1440
               _Version        =   65536
               _ExtentX        =   2540
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
            Begin EditLib.fpDoubleSingle txt_IntLeg 
               Height          =   315
               Left            =   4380
               TabIndex        =   14
               Top             =   1450
               Width           =   1200
               _Version        =   196608
               _ExtentX        =   2117
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
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Interés Legal"
               Height          =   195
               Left            =   3105
               TabIndex        =   64
               Top             =   1530
               Width           =   915
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Devolución BBP"
               Height          =   195
               Left            =   120
               TabIndex        =   63
               Top             =   1500
               Width           =   1170
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Capital PBP"
               Height          =   195
               Left            =   8955
               TabIndex        =   61
               Top             =   510
               Width           =   840
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Interés PBP"
               Height          =   195
               Left            =   8955
               TabIndex        =   60
               Top             =   840
               Width           =   840
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Deuda Pend.(*)"
               Height          =   195
               Left            =   8955
               TabIndex        =   57
               Top             =   1170
               Width           =   1095
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Seguro Desgrav."
               Height          =   195
               Left            =   120
               TabIndex        =   54
               Top             =   1185
               Width           =   1200
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Interés TNC a la fecha"
               Height          =   195
               Left            =   5760
               TabIndex        =   53
               Top             =   510
               Width           =   1605
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Interés TC a la fecha"
               Height          =   195
               Left            =   5760
               TabIndex        =   52
               Top             =   855
               Width           =   1485
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               Caption         =   "Dias TNC"
               Height          =   195
               Left            =   3105
               TabIndex        =   51
               Top             =   510
               Width           =   690
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "Dias TC"
               Height          =   195
               Left            =   3105
               TabIndex        =   50
               Top             =   855
               Width           =   570
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Monto Portes"
               Height          =   195
               Left            =   5760
               TabIndex        =   49
               Top             =   1185
               Width           =   945
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Seguro Inmueble"
               Height          =   195
               Left            =   3105
               TabIndex        =   48
               Top             =   1185
               Width           =   1200
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               Caption         =   "Último Vcto. TNC"
               Height          =   195
               Left            =   120
               TabIndex        =   47
               Top             =   510
               Width           =   1230
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "Último Vcto. TC"
               Height          =   195
               Left            =   120
               TabIndex        =   46
               Top             =   855
               Width           =   1110
            End
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Total Prepagar"
            Height          =   195
            Left            =   5880
            TabIndex        =   37
            Top             =   2970
            Width           =   1050
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Total Antes ITF"
            Height          =   195
            Left            =   240
            TabIndex        =   30
            Top             =   2970
            Width           =   1095
         End
         Begin VB.Label Label11 
            Caption         =   "Fecha Prepago"
            Height          =   315
            Left            =   120
            TabIndex        =   29
            Top             =   450
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Valor Asegur. Inm."
            Height          =   315
            Left            =   2850
            TabIndex        =   28
            Top             =   450
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Saldo Actual TC"
            Height          =   315
            Left            =   8820
            TabIndex        =   27
            Top             =   450
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "Saldo Actual TNC"
            Height          =   315
            Left            =   5940
            TabIndex        =   26
            Top             =   450
            Width           =   1335
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Monto ITF"
            Height          =   195
            Left            =   3210
            TabIndex        =   25
            Top             =   2970
            Width           =   735
         End
         Begin VB.Label Label29 
            Caption         =   "Datos del Prepago"
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
            TabIndex        =   24
            Top             =   60
            UseMnemonic     =   0   'False
            Width           =   1875
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   3405
         Left            =   30
         TabIndex        =   31
         Top             =   1440
         Width           =   11625
         _Version        =   65536
         _ExtentX        =   20505
         _ExtentY        =   6006
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
            Height          =   2955
            Left            =   60
            TabIndex        =   6
            Top             =   330
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   5212
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
            Left            =   120
            TabIndex        =   32
            Top             =   60
            Width           =   1875
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   645
         Left            =   30
         TabIndex        =   33
         Top             =   780
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   2430
            Picture         =   "OpeTra_frm_325.frx":0044
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Grabar Prepago"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Simula 
            Height          =   585
            Left            =   1830
            Picture         =   "OpeTra_frm_325.frx":0486
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Exportar simulación de liquidación"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ImpCro 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_325.frx":0790
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Consulta Cronograma de Cuotas"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_VerPag 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_325.frx":0A9A
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Consulta Pagos del Cliente"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   11010
            Picture         =   "OpeTra_frm_325.frx":0DA4
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_PolSeg 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_325.frx":11E6
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Consulta Pólizas de Seguros"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   615
         Left            =   30
         TabIndex        =   34
         Top             =   150
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
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   10950
            Top             =   150
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   555
            Left            =   690
            TabIndex        =   35
            Top             =   30
            Width           =   4575
            _Version        =   65536
            _ExtentX        =   8070
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "Prepago Total de Crédito Hipotecario - Registro"
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
            Left            =   10350
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
            Left            =   9780
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
            Left            =   90
            Picture         =   "OpeTra_frm_325.frx":14F0
            Top             =   60
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_Con_PreTot_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_dbl_PorNco           As Double
Dim l_dbl_PorCon           As Double
Dim l_str_PrxVct           As String
Dim l_int_NumCuo           As Integer
Dim l_int_PagCuo           As Integer
Dim l_int_PerGra           As Integer
Dim l_int_CodMod           As Integer
Dim l_int_FlgOpe           As Integer
Dim l_dbl_SalNco           As Double
Dim l_dbl_SalCon           As Double
Dim l_int_CodPrd           As Integer
Dim l_dbl_TasInt           As Double
Dim l_dbl_SegDes           As Double
Dim l_dbl_SegInm           As Double
Dim l_dbl_PorITF           As Double
Dim l_int_TipSeg           As Integer
Dim l_str_NomAch           As String

Private Sub cmd_VerPag_Click()
   frm_Ges_CreHip_05.Show 1
End Sub

Private Sub cmd_ImpCro_Click()
   modmip_g_int_OrdAct = 1
   frm_Ges_CreHip_07.Show 1
End Sub

Private Sub cmd_PolSeg_Click()
   frm_Con_PolSeg_01.Show 1
End Sub

Private Sub cmd_Simula_Click()
   'Validaciones
   'If CDbl(txt_MontoPortes.Text) = 0 Then
   '   MsgBox "El Monto de Portes debe ser mayor a 0.", vbExclamation, modgen_g_str_NomPlt
   '   Call gs_SetFocus(txt_MontoPortes)
   '   Exit Sub
   'End If
   If CDbl(pnl_MtoApl.Caption) = 0 Then
      MsgBox "El Monto Total Antes ITF debe ser mayor a 0.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecPre)
      Exit Sub
   End If
   If CDbl(pnl_MtoPrepagar.Caption) = 0 Then
      MsgBox "El monto Total Prepagar debe ser mayor a 0.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecPre)
      Exit Sub
   End If
   If CInt(pnl_DiasTNC.Caption) > 30 Then
      MsgBox "El campo 'Dias TNC' no puede ser mayor a 30.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecPre)
      Exit Sub
   End If
   If cmb_MotPpg.ListIndex = -1 Then
      MsgBox "Debe seleccionar el motivo del prepago total.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_MotPpg)
      Exit Sub
   End If
   'If CDbl(pnl_DeuPen) > 0# Then
   '   MsgBox "No se puede procesar el Prepago Total, porque tiene Cuotas Pendientes.", vbExclamation, modgen_g_str_NomPlt
   '   Exit Sub
   'End If
   If CDbl(pnl_DevBBP.Caption) > 0 Then
      If CDbl(txt_IntLeg.Value) = 0 Then
         MsgBox "El monto del Interes Legal debe ser mayor a 0.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_IntLeg)
         Exit Sub
      End If
   End If
   If InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then
      If CDbl(txt_MontoITF.Text) = 0 Then
         MsgBox "El Monto de ITF debe ser mayor a 0.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_MontoITF)
         Exit Sub
      End If
   End If
   
   'confirma simulacion
   If MsgBox("¿Está seguro de imprimir simulación de liquidación de Prepago Total ?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   'imprime simulacion
   Screen.MousePointer = 11
   If InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Then
      Call fs_PpgTot_Micasita(False, l_str_NomAch)
   Else
      Call fs_PpgTot_Mivivienda(False, l_str_NomAch)
   End If
   
   'Graba datos en Solicitud de tabla prepagos
   Call fs_usp_cre_ppgsol
   If moddat_g_int_CntErr = 1 Then
      MsgBox "No se pudo completar el procedimiento 'usp_cre_ppgsol'.", vbCritical, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Simula)
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Grabar_Click()
   'Validaciones
   'If CDbl(txt_MontoPortes.Text) = 0 Then
   '   MsgBox "El Monto de Portes debe ser mayor a 0.", vbExclamation, modgen_g_str_NomPlt
   '   Call gs_SetFocus(txt_MontoPortes)
   '   Exit Sub
   'End If
   If CDbl(pnl_MtoApl.Caption) = 0 Then
      MsgBox "El Monto Total Antes ITF debe ser mayor a 0.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_IntLeg)
      Exit Sub
   End If
   If CDbl(pnl_MtoPrepagar.Caption) = 0 Then
      MsgBox "El monto Total Prepagar debe ser mayor a 0.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_IntLeg)
      Exit Sub
   End If
   If CInt(pnl_DiasTNC.Caption) > 30 Then
      MsgBox "El campo 'Dias TNC' no puede ser mayor a 30.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecPre)
      Exit Sub
   End If
   If cmb_MotPpg.ListIndex = -1 Then
      MsgBox "Debe seleccionar el motivo del prepago total.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_MotPpg)
      Exit Sub
   End If
   If CDbl(pnl_DeuPen) > 0# Then
      MsgBox "No se puede procesar el Prepago Total, porque tiene Cuotas Pendientes.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   If CDbl(pnl_DevBBP.Caption) > 0 Then
      If CDbl(txt_IntLeg.Value) = 0 Then
         MsgBox "El monto del Interes Legal debe ser mayor a 0.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_IntLeg)
         Exit Sub
      End If
   End If
   If InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then
      If CDbl(txt_MontoITF.Text) = 0 Then
         MsgBox "El Monto de ITF debe ser mayor a 0.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_MontoITF)
         Exit Sub
      End If
   End If
   
'   If fs_validar_ppgpnd = 1 Then
'      MsgBox "Existe un prepago pendiente de regularización de COFIDE.", vbExclamation, modgen_g_str_NomPlt
'         Call gs_SetFocus(cmd_Grabar)
'         Exit Sub
'   End If

   'confirma grabacion
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'graba prepago
   Screen.MousePointer = 11
   Call fs_Insert_ppgcab
   
   'Actualiza estado en tabla prepagos
   If moddat_g_int_CntErr = 0 Then
      If fs_usp_actualiza_cre_ppgcab = 1 Then
         MsgBox "No se pudo completar el procedimiento 'usp_actualiza_cre_ppgcab'.", vbCritical, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_Grabar)
         Screen.MousePointer = 0
         Exit Sub
      End If
   End If
   
   'imprime simulacion
   Screen.MousePointer = 11
   If InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Then
      Call fs_PpgTot_Micasita(True, l_str_NomAch)
   Else
      Call fs_PpgTot_Mivivienda(True, l_str_NomAch)
   End If
   
   'Enviando Correo usuarios
   Call fs_Envia_Correo
   
   'Enviando Correo plaft
   Call fs_Envia_Correo_Plaft
   
   'Imprime liquidacion
   If moddat_g_int_CntErr = 0 Then
      If InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Then
         Call fs_PpgTot_Micasita(False, l_str_NomAch)
      Else
         Call fs_PpgTot_Mivivienda(False, l_str_NomAch)
      End If
      
      Screen.MousePointer = 0
      Unload Me
   End If
   Screen.MousePointer = 0
End Sub

'* actualiza el estado en la tabla de cabecera de prepagos
Private Function fs_usp_actualiza_cre_ppgcab() As Integer
   fs_usp_actualiza_cre_ppgcab = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "USP_ACTUALIZA_CRE_PPGCAB ( "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
   g_str_Parame = g_str_Parame & "" & ipp_FecPre.Year & IIf(Len(Trim(ipp_FecPre.Month)) = 1, 0 & ipp_FecPre.Month, ipp_FecPre.Month) & IIf(Len(Trim(ipp_FecPre.Day)) = 1, 0 & ipp_FecPre.Day, ipp_FecPre.Day) & " , 1, 0, 0, 0, 0) "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      fs_usp_actualiza_cre_ppgcab = 1
   End If
End Function

Private Sub fs_Envia_Correo()
Dim r_str_Mensaj     As String
Dim r_str_Asunto     As String

   ReDim moddat_g_arr_Genera(0)

   r_str_Asunto = "PREPAGO TOTAL DE CREDITO HIPOTECARIO (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
   r_str_Mensaj = ""
   r_str_Mensaj = r_str_Mensaj & "NUMERO DE OPERACION : " & moddat_g_str_NumOpe & Chr(13)
   r_str_Mensaj = r_str_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
   r_str_Mensaj = r_str_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
   r_str_Mensaj = r_str_Mensaj & Chr(13)
   r_str_Mensaj = r_str_Mensaj & Chr(13)
   r_str_Mensaj = r_str_Mensaj & "ESTIMADOS: "
   r_str_Mensaj = r_str_Mensaj & Chr(13)
   r_str_Mensaj = r_str_Mensaj & "FAVOR DE PROCEDER CON LA EMISION DE LA MINUTA DE LEVANTAMIENTO DE HIPOTECA DEL CLIENTE."
   r_str_Mensaj = r_str_Mensaj & ""
   r_str_Mensaj = r_str_Mensaj & "SALUDOS CORDIALES"

   'Evaluador de Operaciones
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(221, moddat_g_arr_Genera)

   'Director de Producción
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(200, moddat_g_arr_Genera)

   'Jefe de Legal
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(230, moddat_g_arr_Genera)
   
   'Evaluador Legal 1
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(231, moddat_g_arr_Genera)

   'Evaluador Legal 2
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(232, moddat_g_arr_Genera)

   'Plataforma
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(122, moddat_g_arr_Genera)

   Call moddat_gs_EnvCor(mps_Sesion, mps_Mensaj, moddat_g_arr_Genera, r_str_Asunto, r_str_Mensaj, l_str_NomAch, g_str_RutLog & "\")
End Sub

Private Sub fs_Envia_Correo_Plaft()
Dim r_str_Mensaj     As String
Dim r_str_Asunto     As String

   If Mid(moddat_g_str_NumOpe, 1, 3) = "001" Or Mid(moddat_g_str_NumOpe, 1, 3) = "002" Then
      If CDbl(pnl_MtoPrepagar.Caption) < 10000 Then
         Exit Sub
      End If
   Else
      If CDbl(pnl_MtoPrepagar.Caption) < 30000 Then
         Exit Sub
      End If
   End If
   
   ReDim moddat_g_arr_Genera(0)

   r_str_Asunto = "ALERTA DE PREPAGO (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
   r_str_Mensaj = ""
   r_str_Mensaj = r_str_Mensaj & "NUMERO DE OPERACION : " & moddat_g_str_NumOpe & Chr(13)
   r_str_Mensaj = r_str_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
   r_str_Mensaj = r_str_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
   r_str_Mensaj = r_str_Mensaj & Chr(13)
   r_str_Mensaj = r_str_Mensaj & "El sistema alerto el prepago detallado en el adjunto"
  
   'Jefe de Legal
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(230, moddat_g_arr_Genera)
   
   Call moddat_gs_EnvCor(mps_Sesion, mps_Mensaj, moddat_g_arr_Genera, r_str_Asunto, r_str_Mensaj, l_str_NomAch, g_str_RutLog & "\")
End Sub

Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Buscar_Cuotas_Vencidas
   Call fs_Buscar
   Call fs_Buscar_Credito
   Call fs_Calcula_PbpPerdido(moddat_g_str_NumOpe)
   Call fs_ValidaTiempo(moddat_g_str_NumOpe, True)
   Call gs_CentraForm(Me)
   If UCase(App.EXEName) = "OPETRA" Then
      cmd_Grabar.Enabled = True
   Else
      cmd_Grabar.Enabled = False
   End If
   
   Call gs_SetFocus(ipp_FecPre)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Grid de Datos del Crédito
   grd_Listad.ColWidth(0) = 2900
   grd_Listad.ColWidth(1) = 8150
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_Listad)
   
   pnl_Val_AsgInm.Caption = "0.00 "
   pnl_SaldoTNC1.Caption = "0.00 "
   pnl_SaldoTC1.Caption = "0.00 "
   pnl_UltPagTNC.Caption = " "
   pnl_UltPagTC.Caption = " "
   pnl_DiasTNC.Caption = "0 "
   pnl_DiasTC.Caption = "0 "
   txt_InteresTNC.Text = 0
   txt_InteresTC.Text = 0
   txt_SegDes.Text = 0
   txt_SegInm.Text = 0
   txt_MontoITF.Text = 0
   pnl_MtoApl.Caption = "0.00 "
   txt_ObsPpg.Text = " "
   pnl_DevBBP.Caption = "0.00 "
   txt_IntLeg.Text = 0
   cmb_MotPpg.ListIndex = -1
   
   If UCase(App.EXEName) = "OPETRA" Then
      cmd_Grabar.Enabled = True
      txt_InteresTNC.Enabled = True
      txt_InteresTC.Enabled = True
      txt_SegDes.Enabled = True
      txt_SegInm.Enabled = True
      txt_MontoPortes.Enabled = True
      txt_MontoITF.Enabled = True
      txt_IntLeg.Enabled = True
   Else
      cmd_Grabar.Enabled = False
      txt_InteresTNC.Enabled = False
      txt_InteresTC.Enabled = False
      txt_SegDes.Enabled = False
      txt_SegInm.Enabled = False
      txt_MontoPortes.Enabled = False
      txt_MontoITF.Enabled = False
      txt_IntLeg.Enabled = False
   End If
   
   'Cuotas Pendientes
   grd_DeuPen.ColWidth(0) = 575
   grd_DeuPen.ColWidth(1) = 630
   grd_DeuPen.ColWidth(2) = 1020
   grd_DeuPen.ColWidth(3) = 0
   grd_DeuPen.ColWidth(4) = 965
   grd_DeuPen.ColWidth(5) = 965
   grd_DeuPen.ColWidth(6) = 965
   grd_DeuPen.ColWidth(7) = 965
   grd_DeuPen.ColWidth(8) = 840
   grd_DeuPen.ColWidth(9) = 965
   grd_DeuPen.ColWidth(10) = 965
   grd_DeuPen.ColWidth(11) = 0
   grd_DeuPen.ColWidth(12) = 0
   grd_DeuPen.ColWidth(13) = 965
   grd_DeuPen.ColWidth(14) = 0
   grd_DeuPen.ColWidth(15) = 965
   grd_DeuPen.ColAlignment(0) = flexAlignCenterCenter
   grd_DeuPen.ColAlignment(1) = flexAlignCenterCenter
   grd_DeuPen.ColAlignment(2) = flexAlignCenterCenter
   grd_DeuPen.ColAlignment(3) = flexAlignCenterCenter
   grd_DeuPen.ColAlignment(4) = flexAlignRightCenter
   grd_DeuPen.ColAlignment(5) = flexAlignRightCenter
   grd_DeuPen.ColAlignment(6) = flexAlignRightCenter
   grd_DeuPen.ColAlignment(7) = flexAlignRightCenter
   grd_DeuPen.ColAlignment(8) = flexAlignRightCenter
   grd_DeuPen.ColAlignment(9) = flexAlignRightCenter
   grd_DeuPen.ColAlignment(10) = flexAlignRightCenter
   grd_DeuPen.ColAlignment(11) = flexAlignRightCenter
   grd_DeuPen.ColAlignment(12) = flexAlignRightCenter
   grd_DeuPen.ColAlignment(13) = flexAlignRightCenter
   grd_DeuPen.ColAlignment(14) = flexAlignRightCenter
   grd_DeuPen.ColAlignment(15) = flexAlignRightCenter
   Call gs_LimpiaGrid(grd_DeuPen)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_MotPpg, 1, "115")
   ipp_FecPre.DateValue = date
End Sub

Private Sub fs_Buscar_Cuotas_Vencidas()
Dim r_dbl_ValCuo     As Double
Dim r_str_Parame     As String
Dim r_rst_Princi     As ADODB.Recordset
Dim r_dbl_DeuVen     As Double

   'Cuotas Vencidas
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT HIPCUO_NUMCUO, HIPCUO_FECVCT, HIPMAE_MONEDA, HIPCUO_CAPITA, HIPCUO_CAPPAG, HIPCUO_INTERE, HIPCUO_INTPAG, "
   r_str_Parame = r_str_Parame & "        HIPCUO_DESORG, HIPCUO_DESPAG, HIPCUO_VIVORG, HIPCUO_VIVPAG, HIPCUO_OTRORG, HIPCUO_OTRPAG, HIPCUO_CAPBBP, "
   r_str_Parame = r_str_Parame & "        HIPCUO_CBPPAG, HIPCUO_INTBBP, HIPCUO_IBPPAG, HIPCUO_INTMOR, HIPCUO_IMOPAG, HIPCUO_INTCOM, HIPCUO_ICOPAG,  "
   r_str_Parame = r_str_Parame & "        HIPCUO_GASCOB, HIPCUO_GCOPAG, HIPCUO_OTRGAS, HIPCUO_OTGPAG "
   r_str_Parame = r_str_Parame & "   FROM CRE_HIPCUO INNER JOIN CRE_HIPMAE ON HIPMAE_NUMOPE = HIPCUO_NUMOPE "
   r_str_Parame = r_str_Parame & "  WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   r_str_Parame = r_str_Parame & "    AND HIPCUO_TIPCRO = 1 "
   r_str_Parame = r_str_Parame & "    AND HIPCUO_SITUAC = 2 "
   r_str_Parame = r_str_Parame & "    AND HIPCUO_FECVCT <= " & Format(ipp_FecPre.Text, "yyyymmdd") & " "
   r_str_Parame = r_str_Parame & "  ORDER BY HIPCUO_NUMCUO ASC "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Exit Sub
   End If
   
   Call gs_LimpiaGrid(grd_DeuPen)
   pnl_DeuPen.Caption = "0.00 "
   
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      grd_DeuPen.Redraw = False
      grd_DeuPen.Cols = 16
      
      'Cabecera de la Grilla
      grd_DeuPen.Rows = grd_DeuPen.Rows + 2
      grd_DeuPen.FixedRows = 1
      grd_DeuPen.Rows = grd_DeuPen.Rows - 1
      grd_DeuPen.Row = 0
          
      grd_DeuPen.Col = 0:    grd_DeuPen.Text = "Cuota":        grd_DeuPen.CellAlignment = flexAlignCenterCenter
      grd_DeuPen.Col = 1:    grd_DeuPen.Text = "Estado":       grd_DeuPen.CellAlignment = flexAlignCenterCenter
      grd_DeuPen.Col = 2:    grd_DeuPen.Text = "F. Vcto.":     grd_DeuPen.CellAlignment = flexAlignCenterCenter
      'grd_DeuPen.Col = 3:    grd_DeuPen.Text = "Moneda":       grd_DeuPen.CellAlignment = flexAlignCenterCenter
      grd_DeuPen.Col = 4:    grd_DeuPen.Text = "Capital":      grd_DeuPen.CellAlignment = flexAlignCenterCenter
      grd_DeuPen.Col = 5:    grd_DeuPen.Text = "Interés":      grd_DeuPen.CellAlignment = flexAlignCenterCenter
      grd_DeuPen.Col = 6:    grd_DeuPen.Text = "Seg. Desg.":   grd_DeuPen.CellAlignment = flexAlignCenterCenter
      grd_DeuPen.Col = 7:    grd_DeuPen.Text = "Seg. Viv.":    grd_DeuPen.CellAlignment = flexAlignCenterCenter
      grd_DeuPen.Col = 8:    grd_DeuPen.Text = "Portes":       grd_DeuPen.CellAlignment = flexAlignCenterCenter
      grd_DeuPen.Col = 9:    grd_DeuPen.Text = "Capital BBP":  grd_DeuPen.CellAlignment = flexAlignCenterCenter
      grd_DeuPen.Col = 10:   grd_DeuPen.Text = "Interés BBP":  grd_DeuPen.CellAlignment = flexAlignCenterCenter
      'grd_DeuPen.Col = 11:   grd_DeuPen.Text = "Int. Morat.":  grd_DeuPen.CellAlignment = flexAlignCenterCenter
      'grd_DeuPen.Col = 12:   grd_DeuPen.Text = "Int. Comp.":   grd_DeuPen.CellAlignment = flexAlignCenterCenter
      grd_DeuPen.Col = 13:   grd_DeuPen.Text = "G. Cobr.":     grd_DeuPen.CellAlignment = flexAlignCenterCenter
      'grd_DeuPen.Col = 14:   grd_DeuPen.Text = "Otr. Gastos":  grd_DeuPen.CellAlignment = flexAlignCenterCenter
      grd_DeuPen.Col = 15:   grd_DeuPen.Text = "Total Cuota":  grd_DeuPen.CellAlignment = flexAlignCenterCenter
      
      r_rst_Princi.MoveFirst
      
      Do While Not r_rst_Princi.EOF
         grd_DeuPen.Rows = grd_DeuPen.Rows + 1
         grd_DeuPen.Row = grd_DeuPen.Rows - 1
         r_dbl_ValCuo = 0
         
         grd_DeuPen.Col = 0
         grd_DeuPen.Text = Format(r_rst_Princi!HIPCUO_NUMCUO, "000")
         
         grd_DeuPen.Col = 1
         grd_DeuPen.Text = IIf(CLng(r_rst_Princi!HIPCUO_FECVCT) < CLng(Format(date, "yyyymmdd")), "V", "P")
      
         grd_DeuPen.Col = 2
         grd_DeuPen.Text = gf_FormatoFecha(CStr(r_rst_Princi!HIPCUO_FECVCT))
      
         grd_DeuPen.Col = 3
         'grd_DeuPen.Text = moddat_gf_Consulta_ParDes("229", CStr(r_rst_Princi!HIPMAE_MONEDA))
         
         'Capital
         grd_DeuPen.Col = 4
         grd_DeuPen.Text = Format(r_rst_Princi!HIPCUO_CAPITA - r_rst_Princi!HIPCUO_CAPPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_DeuPen.Text)
      
         'Interes
         grd_DeuPen.Col = 5
         grd_DeuPen.Text = Format(r_rst_Princi!HIPCUO_INTERE - r_rst_Princi!HIPCUO_INTPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_DeuPen.Text)
      
         'Seguro de Desgravamen
         grd_DeuPen.Col = 6
         grd_DeuPen.Text = Format(r_rst_Princi!HIPCUO_DESORG - r_rst_Princi!HIPCUO_DESPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_DeuPen.Text)
      
         'Seguro de Vivienda
         grd_DeuPen.Col = 7
         grd_DeuPen.Text = Format(r_rst_Princi!HIPCUO_VIVORG - r_rst_Princi!HIPCUO_VIVPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_DeuPen.Text)
      
         'Otros Cargos
         grd_DeuPen.Col = 8
         grd_DeuPen.Text = Format(r_rst_Princi!HIPCUO_OTRORG - r_rst_Princi!HIPCUO_OTRPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_DeuPen.Text)
      
         'Capital PBP
         grd_DeuPen.Col = 9
         grd_DeuPen.Text = Format(r_rst_Princi!HIPCUO_CAPBBP - r_rst_Princi!HIPCUO_CBPPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_DeuPen.Text)
      
         'Interés PBP
         grd_DeuPen.Col = 10
         grd_DeuPen.Text = Format(r_rst_Princi!HIPCUO_INTBBP - r_rst_Princi!HIPCUO_IBPPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_DeuPen.Text)
      
         'Interes Moratorio
         grd_DeuPen.Col = 11
         'grd_DeuPen.Text = Format(r_rst_Princi!HIPCUO_INTMOR - r_rst_Princi!HIPCUO_IMOPAG, "###,###,##0.00")
         'r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_DeuPen.Text)
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(r_rst_Princi!HIPCUO_INTMOR - r_rst_Princi!HIPCUO_IMOPAG)
         
         'Interes Compensatorio
         grd_DeuPen.Col = 12
         'grd_DeuPen.Text = Format(r_rst_Princi!HIPCUO_INTCOM - r_rst_Princi!HIPCUO_ICOPAG, "###,###,##0.00")
         'r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_DeuPen.Text)
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(r_rst_Princi!HIPCUO_INTCOM - r_rst_Princi!HIPCUO_ICOPAG)
      
         'Gastos de Cobranza
         grd_DeuPen.Col = 13
         grd_DeuPen.Text = Format(r_rst_Princi!HIPCUO_GASCOB - r_rst_Princi!HIPCUO_GCOPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_DeuPen.Text)
      
         'Otros Gastos
         grd_DeuPen.Col = 14
         'grd_DeuPen.Text = Format(r_rst_Princi!HIPCUO_OTRGAS - r_rst_Princi!HIPCUO_OTGPAG, "###,###,##0.00")
         'r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_DeuPen.Text)
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(r_rst_Princi!HIPCUO_OTRGAS - r_rst_Princi!HIPCUO_OTGPAG)
                  
      
         'Valor Cuota
         grd_DeuPen.Col = 15
         grd_DeuPen.Text = Format(r_dbl_ValCuo, "###,###,##0.00")
         
         r_dbl_DeuVen = r_dbl_DeuVen + r_dbl_ValCuo
         r_rst_Princi.MoveNext
      Loop
      grd_DeuPen.Redraw = True
      
      pnl_DeuPen.Caption = Format(r_dbl_DeuVen, "###,###.00") & " "
      Call gs_UbiIniGrid(grd_DeuPen)
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub ipp_FecPre_LostFocus()
   If IsDate(pnl_UltPagTNC.Caption) Then
      pnl_DiasTNC.Caption = DateDiff("d", pnl_UltPagTNC.Caption, ipp_FecPre.Text) & " "
      Call fs_Cal_Interes
      Call fs_Cal_MtoItf
      Call fs_Cal_Prepago
   End If
   If IsDate(pnl_UltPagTC.Caption) Then
      pnl_DiasTC.Caption = DateDiff("d", pnl_UltPagTC.Caption, ipp_FecPre.Text) & " "
      Call fs_Cal_Interes
      Call fs_Cal_MtoItf
      Call fs_Cal_Prepago
   End If
   If IsDate(pnl_UltPagTNC.Caption) And IsDate(pnl_UltPagTC.Caption) Then
      Call fs_Buscar_Cuotas_Vencidas
      Call fs_Buscar
   End If
End Sub

Private Sub ipp_FecPre_Change()
   If IsDate(pnl_UltPagTNC.Caption) Then
      pnl_DiasTNC.Caption = DateDiff("d", pnl_UltPagTNC.Caption, ipp_FecPre.Text) & " "
      Call fs_Cal_Interes
      Call fs_Cal_MtoItf
      Call fs_Cal_Prepago
   End If
   If IsDate(pnl_UltPagTC.Caption) Then
      pnl_DiasTC.Caption = DateDiff("d", pnl_UltPagTC.Caption, ipp_FecPre.Text) & " "
      Call fs_Cal_Interes
      Call fs_Cal_MtoItf
      Call fs_Cal_Prepago
   End If
   If IsDate(pnl_UltPagTNC.Caption) And IsDate(pnl_UltPagTC.Caption) Then
      Call fs_Buscar_Cuotas_Vencidas
      Call fs_Buscar
   End If
End Sub

Private Sub ipp_FecPre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call fs_Cal_Interes
      Call gs_SetFocus(txt_InteresTNC)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_InteresTNC_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then
         Call gs_SetFocus(txt_SegDes)
      Else
         Call gs_SetFocus(txt_InteresTC)
      End If
   End If
End Sub

Private Sub txt_InteresTNC_LostFocus()
   Call fs_Cal_Prepago
End Sub

Private Sub txt_InteresTC_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_SegDes)
   End If
End Sub

Private Sub txt_InteresTC_LostFocus()
   Call fs_Cal_Prepago
End Sub

Private Sub txt_IntLeg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_MontoITF)
   End If
End Sub

Private Sub txt_IntLeg_LostFocus()
   Call fs_Cal_Prepago
End Sub

Private Sub txt_ObsPpg_KeyPress(KeyAscii As Integer)
   If Not KeyAscii = 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub txt_SegDes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_SegInm)
   End If
End Sub

Private Sub txt_SegDes_LostFocus()
   Call fs_Cal_Prepago
End Sub

Private Sub txt_SegInm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_MontoPortes)
   End If
End Sub

Private Sub txt_SegInm_LostFocus()
   Call fs_Cal_Prepago
End Sub

Private Sub txt_MontoPortes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_IntLeg)
   End If
End Sub

Private Sub txt_MontoPortes_LostFocus()
   Call fs_Cal_Prepago
End Sub

Private Sub txt_MontoITF_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_MotPpg)
   End If
End Sub

Private Sub txt_MontoITF_LostFocus()
   Call fs_Cal_Prepago
End Sub
 
Private Sub cmb_MotPre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_MotPpg.ListIndex > -1 Then
         Call gs_SetFocus(txt_ObsPpg)
      End If
   End If
End Sub
 
Private Sub fs_Buscar()
Dim r_str_CodPry     As String
Dim r_str_NomPry     As String
Dim r_str_CodBco     As String
Dim r_bol_Estado     As Boolean
   
   'Buscando Información del Crédito
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.HIPMAE_TASINT, A.HIPMAE_FOIPRE, A.HIPMAE_FOIVIV, A.HIPMAE_TIPSEG, A.HIPMAE_CODPRD, "
   g_str_Parame = g_str_Parame & "        A.HIPMAE_NUMCUO, A.HIPMAE_CUOPAG, A.HIPMAE_PERGRA, A.HIPMAE_IMPNCO, A.HIPMAE_TOTPRE, "
   g_str_Parame = g_str_Parame & "        A.HIPMAE_PRXVCT, A.HIPMAE_NUMOPE, A.HIPMAE_FECDES, A.HIPMAE_SALCAP, A.HIPMAE_SALCON, "
   g_str_Parame = g_str_Parame & "        NVL(B.SOLMAE_FMVBBP, 0) AS SOLMAE_FMVBBP, SOLMAE_PBPMTO "
   g_str_Parame = g_str_Parame & "   FROM CRE_HIPMAE A "
   g_str_Parame = g_str_Parame & "  INNER JOIN CRE_SOLMAE B ON B.SOLMAE_NUMERO = A.HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & "  WHERE A.HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "'"
   g_str_Parame = g_str_Parame & "    AND (A.HIPMAE_SITUAC = 2) "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If

   g_rst_Princi.MoveFirst
   
   'DATOS DE LA OPERACION DEL PREPAGO
   l_dbl_TasInt = g_rst_Princi!HIPMAE_TASINT
   l_dbl_SegDes = g_rst_Princi!HIPMAE_FOIPRE
   l_dbl_SegInm = g_rst_Princi!HIPMAE_FOIVIV
   l_int_TipSeg = g_rst_Princi!HIPMAE_TIPSEG  'seguro desgravamen
   l_int_CodPrd = g_rst_Princi!HIPMAE_CODPRD
   l_int_NumCuo = g_rst_Princi!HIPMAE_NUMCUO
   l_int_PagCuo = g_rst_Princi!HIPMAE_CUOPAG
   l_int_PerGra = g_rst_Princi!HIPMAE_PERGRA
   
   l_dbl_PorNco = Format(g_rst_Princi!HIPMAE_IMPNCO / g_rst_Princi!HIPMAE_TOTPRE, "##0.0000")
   l_dbl_PorCon = 1 - l_dbl_PorNco
   l_str_PrxVct = g_rst_Princi!HIPMAE_PRXVCT

   If pnl_DeuPen.Caption > 0# Then
'      r_bol_CuoVen = True
      l_dbl_SalNco = fs_Obtiene_Saldos(g_rst_Princi!HIPMAE_NUMOPE, 1, True)
      l_dbl_SalCon = fs_Obtiene_Saldos(g_rst_Princi!HIPMAE_NUMOPE, 2, True)
      pnl_UltPagTNC.Caption = gf_FormatoFecha(fs_Obtiene_FechaPago(g_rst_Princi!HIPMAE_NUMOPE, 1, g_rst_Princi!HIPMAE_FECDES, True))
      pnl_UltPagTC.Caption = gf_FormatoFecha(fs_Obtiene_FechaPago(g_rst_Princi!HIPMAE_NUMOPE, 2, g_rst_Princi!HIPMAE_FECDES, True))
   Else
'      r_bol_CuoVen = False
      l_dbl_SalNco = g_rst_Princi!HIPMAE_SALCAP
      l_dbl_SalCon = g_rst_Princi!HIPMAE_SALCON
      pnl_UltPagTNC.Caption = gf_FormatoFecha(fs_Obtiene_FechaPago(g_rst_Princi!HIPMAE_NUMOPE, 1, g_rst_Princi!HIPMAE_FECDES, False))
      pnl_UltPagTC.Caption = gf_FormatoFecha(fs_Obtiene_FechaPago(g_rst_Princi!HIPMAE_NUMOPE, 2, g_rst_Princi!HIPMAE_FECDES, False))
   End If
   pnl_DiasTNC.Caption = DateDiff("d", pnl_UltPagTNC.Caption, ipp_FecPre.Text) & " "
   pnl_DiasTC.Caption = DateDiff("d", pnl_UltPagTC.Caption, ipp_FecPre.Text) & " "
   
   pnl_SaldoTNC1.Caption = Format(l_dbl_SalNco, "###,###.00") & " "
   pnl_SaldoTC1.Caption = Format(l_dbl_SalCon, "###,###.00") & " "
         
   Call fs_ValidaTiempo(moddat_g_str_NumOpe, False, r_bol_Estado)
   If r_bol_Estado = False Then
      pnl_DevBBP.Caption = Format(g_rst_Princi!SOLMAE_FMVBBP + g_rst_Princi!SOLMAE_PBPMTO, "###,###.00") & " "
   End If

   'DETERMINA SI OPERACION ES MICASITA O MIVIVIENDA (SIN TC)
   If InStr(moddat_g_str_Agr1FMV, l_int_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, l_int_CodPrd) > 0 Then
      l_dbl_PorITF = opecaj_gf_Consulta_ITF(Format(CDate(moddat_g_str_FecSis), "yyyymmdd"), 1)
      l_dbl_PorNco = 1
      l_dbl_PorCon = 0
      Label21.Caption = "Saldo Actual"
      Label23.Caption = "Ultimo Pago"
      Label25.Caption = "Dias"
      Label6.Caption = "Interés a la fecha"
      Label20.Visible = False
      pnl_SaldoTC1.Visible = False
      Label22.Visible = False
      pnl_UltPagTC.Visible = False
      Label24.Visible = False
      pnl_DiasTC.Visible = False
      Label7.Visible = False
      txt_InteresTC.Visible = False
   Else
      Label21.Caption = "Saldo Actual TNC"
      Label23.Caption = "Ultimo Pago TNC"
      Label25.Caption = "Dias TNC"
      Label6.Caption = "Interés TNC a la fecha"
      Label20.Visible = True
      pnl_SaldoTC1.Visible = True
      Label22.Visible = True
      pnl_UltPagTC.Visible = True
      Label24.Visible = True
      pnl_DiasTC.Visible = True
      Label7.Visible = True
      txt_InteresTC.Visible = True
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
     
   'Información del Crédito
   Call modmip_gs_DatNumOpe(moddat_g_str_NumOpe, grd_Listad)
End Sub

Private Sub fs_Buscar_ant()
Dim r_str_CodPry     As String
Dim r_str_NomPry     As String
Dim r_str_CodBco     As String

   'Buscando Información del Crédito
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "(HIPMAE_SITUAC = 2)"
   
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
   moddat_g_str_NumSol = Trim(g_rst_Princi!hipmae_numsol)
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
         Case InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd):  grd_Listad.Text = "Nro. Operación Mivivienda"  '"001"
         Case InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd):  grd_Listad.Text = "Nro. Operación COFIDE"      '"003"
         Case InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd): grd_Listad.Text = "Nro. Operación COFIDE"      '"004", "007", "009", "010", "013", "014", "015", "016", "017", "018", "019", "020", "021", "022", "023"
      End Select
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!HIPMAE_OPEMVI & "")
      
      If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then '"003"
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
   
   'DATOS DE LA OPERACION DEL PREPAGO
   l_dbl_TasInt = g_rst_Princi!HIPMAE_TASINT
   l_dbl_SegDes = g_rst_Princi!HIPMAE_FOIPRE
   l_dbl_SegInm = g_rst_Princi!HIPMAE_FOIVIV
   l_int_CodPrd = g_rst_Princi!HIPMAE_CODPRD
   l_int_NumCuo = g_rst_Princi!HIPMAE_NUMCUO
   l_int_PagCuo = g_rst_Princi!HIPMAE_CUOPAG
   l_int_PerGra = g_rst_Princi!HIPMAE_PERGRA
   l_dbl_PorNco = Format(g_rst_Princi!HIPMAE_IMPNCO / g_rst_Princi!HIPMAE_TOTPRE, "##0.0000")
   l_dbl_PorCon = 1 - l_dbl_PorNco
   l_str_PrxVct = g_rst_Princi!HIPMAE_PRXVCT

   If pnl_DeuPen.Caption > 0# Then
'      r_bol_CuoVen = True
      l_dbl_SalNco = fs_Obtiene_Saldos(g_rst_Princi!HIPMAE_NUMOPE, 1, True)
      l_dbl_SalCon = fs_Obtiene_Saldos(g_rst_Princi!HIPMAE_NUMOPE, 2, True)
      pnl_UltPagTNC.Caption = gf_FormatoFecha(fs_Obtiene_FechaPago(g_rst_Princi!HIPMAE_NUMOPE, 1, g_rst_Princi!HIPMAE_FECDES, True))
      pnl_UltPagTC.Caption = gf_FormatoFecha(fs_Obtiene_FechaPago(g_rst_Princi!HIPMAE_NUMOPE, 2, g_rst_Princi!HIPMAE_FECDES, True))
   Else
'      r_bol_CuoVen = False
      l_dbl_SalNco = g_rst_Princi!HIPMAE_SALCAP
      l_dbl_SalCon = g_rst_Princi!HIPMAE_SALCON
      pnl_UltPagTNC.Caption = gf_FormatoFecha(fs_Obtiene_FechaPago(g_rst_Princi!HIPMAE_NUMOPE, 1, g_rst_Princi!HIPMAE_FECDES, False))
      pnl_UltPagTC.Caption = gf_FormatoFecha(fs_Obtiene_FechaPago(g_rst_Princi!HIPMAE_NUMOPE, 2, g_rst_Princi!HIPMAE_FECDES, False))
   End If
   
   pnl_DiasTNC.Caption = DateDiff("d", pnl_UltPagTNC.Caption, ipp_FecPre.Text) & " "
   pnl_DiasTC.Caption = DateDiff("d", pnl_UltPagTC.Caption, ipp_FecPre.Text) & " "
   pnl_SaldoTNC1.Caption = Format(l_dbl_SalNco, "###,###.00") & " "
   pnl_SaldoTC1.Caption = Format(l_dbl_SalCon, "###,###.00") & " "
   
   'DETERMINA SI OPERACION ES MICASITA O MIVIVIENDA (SIN TC)
   If InStr(moddat_g_str_Agr1FMV, l_int_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, l_int_CodPrd) > 0 Then
      l_dbl_PorITF = opecaj_gf_Consulta_ITF(Format(CDate(moddat_g_str_FecSis), "yyyymmdd"), 1)
      l_dbl_PorNco = 1
      l_dbl_PorCon = 0
      Label21.Caption = "Saldo Actual"
      Label23.Caption = "Ultimo Pago"
      Label25.Caption = "Dias"
      Label6.Caption = "Interés a la fecha"
      Label20.Visible = False
      pnl_SaldoTC1.Visible = False
      Label22.Visible = False
      pnl_UltPagTC.Visible = False
      Label24.Visible = False
      pnl_DiasTC.Visible = False
      Label7.Visible = False
      txt_InteresTC.Visible = False
   Else
      Label21.Caption = "Saldo Actual TNC"
      Label23.Caption = "Ultimo Pago TNC"
      Label25.Caption = "Dias TNC"
      Label6.Caption = "Interés TNC a la fecha"
      Label20.Visible = True
      pnl_SaldoTC1.Visible = True
      Label22.Visible = True
      pnl_UltPagTC.Visible = True
      Label24.Visible = True
      pnl_DiasTC.Visible = True
      Label7.Visible = True
      txt_InteresTC.Visible = True
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Call gs_UbiIniGrid(grd_Listad)
End Sub

Private Function fs_Obtiene_FechaPago(ByVal p_NumOpe As String, ByVal p_TipCro As Integer, ByVal p_FecDes As String, ByVal p_FlgCuoVen As Boolean) As String
Dim r_rst_Temp    As Recordset
   fs_Obtiene_FechaPago = p_FecDes
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT HIPCUO_FECVCT FROM CRE_HIPCUO  "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & p_NumOpe & "' "
   If p_FlgCuoVen = True Then
      g_str_Parame = g_str_Parame & "   AND HIPCUO_FECVCT <= " & Format(ipp_FecPre.Text, "yyyymmdd") & " "
   Else
      g_str_Parame = g_str_Parame & "   AND HIPCUO_SITUAC = 1 "
   End If
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = " & p_TipCro & " "
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

Private Function fs_Obtiene_Saldos(ByVal p_NumOpe As String, ByVal p_TipCro As Integer, ByVal p_FlgCuoVen As Boolean) As Double
Dim r_rst_Temp    As Recordset
   fs_Obtiene_Saldos = 0#
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT HIPCUO_SALCAP "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & p_NumOpe & "' "
   If p_FlgCuoVen = True Then
      g_str_Parame = g_str_Parame & "   AND HIPCUO_FECVCT <= " & Format(ipp_FecPre.Text, "yyyymmdd") & " "
   Else
      g_str_Parame = g_str_Parame & "   AND HIPCUO_SITUAC = 1 "
   End If
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = " & p_TipCro & " "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_FECVCT DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Temp, 3) Then
       Exit Function
   End If
   
   If Not (r_rst_Temp.BOF And r_rst_Temp.EOF) Then
      r_rst_Temp.MoveFirst
      fs_Obtiene_Saldos = r_rst_Temp!HIPCUO_SALCAP
   End If
   
   r_rst_Temp.Close
   Set r_rst_Temp = Nothing
End Function

Private Sub fs_Buscar_Credito()
Dim r_rst_Temp    As Recordset
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT EVATAS_TIPMON, EVATAS_SUMASE_INM, EVATAS_SUMASE_ES1, EVATAS_SUMASE_ES2, EVATAS_SUMASE_DEP "
   g_str_Parame = g_str_Parame & "  FROM TRA_EVATAS "
   g_str_Parame = g_str_Parame & " WHERE EVATAS_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Temp, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_Temp.BOF And r_rst_Temp.EOF) Then
      r_rst_Temp.MoveFirst
      pnl_Val_AsgInm.Caption = gf_FormatoNumero(r_rst_Temp!EVATAS_SUMASE_INM + r_rst_Temp!EVATAS_SUMASE_ES1 + r_rst_Temp!EVATAS_SUMASE_ES2 + r_rst_Temp!EVATAS_SUMASE_DEP, 12, 2) & " "
   End If
   
   r_rst_Temp.Close
   Set r_rst_Temp = Nothing
End Sub

Private Sub fs_Calcula_PbpPerdido(ByVal p_NumOpe As String)
Dim r_rst_Temp       As Recordset
Dim r_dbl_MtoCap     As Double
Dim r_dbl_MtoInt     As Double
   
   r_dbl_MtoCap = 0
   r_dbl_MtoInt = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT NVL(SUM(HIPCUO_CAPBBP), 0) AS CAPPBP_PERDIDO, NVL(SUM(HIPCUO_INTBBP), 0) AS INTPBP_PERDIDO "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & p_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 1 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_FECVCT > " & Format(ipp_FecPre.Text, "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_SITUAC = 2 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_CAPBBP > 0 "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Temp, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_Temp.BOF And r_rst_Temp.EOF) Then
      r_rst_Temp.MoveFirst
      r_dbl_MtoCap = r_rst_Temp!CAPPBP_PERDIDO
      r_dbl_MtoInt = r_rst_Temp!INTPBP_PERDIDO
   End If
   
   pnl_CapPbpPerdido.Caption = Format(r_dbl_MtoCap, "##,###,##0.00") & " "
   pnl_IntPbpPerdido.Caption = Format(r_dbl_MtoInt, "##,###,##0.00") & " "
   
   r_rst_Temp.Close
   Set r_rst_Temp = Nothing
End Sub

Private Sub fs_ValidaTiempo(ByVal p_NumOpe As String, ByVal p_VerMsg As Boolean, Optional ByRef p_EstVal As Boolean)
Dim r_rst_Temp       As Recordset
Dim r_int_NumCuo     As Integer
   
   r_int_NumCuo = 0
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT MAX(NVL(HIPCUO_NUMCUO,0)) AS NUM_CUOTA "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & p_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 1 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_SITUAC = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Temp, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_Temp.BOF And r_rst_Temp.EOF) Then
      r_rst_Temp.MoveFirst
      If Not IsNull(r_rst_Temp!NUM_CUOTA) Then
         r_int_NumCuo = r_rst_Temp!NUM_CUOTA
      End If
   End If
   
   p_EstVal = True
   If Mid(p_NumOpe, 1, 3) = "021" Or Mid(p_NumOpe, 1, 3) = "022" Or Mid(p_NumOpe, 1, 3) = "023" Then
      If r_int_NumCuo < 60 Then
         If p_VerMsg = True Then
            MsgBox "Segun el reglamento del FMV si aplica un prepago total antes de los 5 años debera devolver el bono otorgado.", vbInformation, modgen_g_str_NomPlt
         End If
         p_EstVal = False
      End If
   End If
   
   r_rst_Temp.Close
   Set r_rst_Temp = Nothing
End Sub

Private Sub fs_Cal_Interes()
   If CDbl(pnl_DiasTNC.Caption) > 0 Then
      txt_InteresTNC.Text = Format((CDbl(pnl_SaldoTNC1.Caption)) * (1 + (l_dbl_TasInt / 100)) ^ (CDbl(pnl_DiasTNC.Caption) / 360) - CDbl(pnl_SaldoTNC1.Caption), "###,##0.00")
   Else
      txt_InteresTNC.Text = 0
   End If
   If Not (InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0) Then
      If CDbl(pnl_DiasTC.Caption) > 0 Then
         txt_InteresTC.Text = Format(CDbl(pnl_SaldoTC1.Caption) * (1 + (l_dbl_TasInt / 100)) ^ (CDbl(pnl_DiasTC.Caption) / 360) - CDbl(pnl_SaldoTC1.Caption), "###,##0.00")
      Else
         txt_InteresTC.Text = 0
      End If
   End If
   If CDbl(pnl_DiasTNC.Caption) > 0 And l_int_TipSeg <> 13 Then
      txt_SegDes.Text = Format((CDbl(pnl_SaldoTNC1.Caption)) * (1 + (l_dbl_SegDes / 100)) ^ (CDbl(pnl_DiasTNC.Caption) / 30) - CDbl(pnl_SaldoTNC1.Caption), "###,##0.00")
   Else
      txt_SegDes.Text = 0
   End If
   If l_dbl_SegInm <> 0 Then
      txt_SegInm.Text = Format(CDbl(pnl_Val_AsgInm.Caption) * (l_dbl_SegInm / 100), "###,##0.00")
   Else
      txt_SegInm.Text = 0
   End If
End Sub

Private Sub fs_Cal_MtoItf()
   If InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then
      txt_MontoITF.Text = Format((l_dbl_PorITF / 100) * CDbl(pnl_MtoApl.Caption), "###,##0.00")
      'txt_MontoPortes.Text = 2
   Else
      txt_MontoITF.Text = 0
      'txt_MontoPortes.Text = 9
   End If
End Sub

Private Function fs_validar_ppgpnd() As Integer
   fs_validar_ppgpnd = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT PPGCAB_FLGEST FROM CRE_PPGCAB "
   g_str_Parame = g_str_Parame & " WHERE PPGCAB_NUMOPE = '" & moddat_g_str_NumOpe & "'"
   
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      fs_validar_ppgpnd = 0
      Exit Function
   End If
   
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      g_rst_GenAux.MoveFirst
      
      Do While Not g_rst_GenAux.EOF
         If g_rst_GenAux!PPGCAB_FLGEST <> 3 Then
            fs_validar_ppgpnd = 1
            Exit Function
         End If
         g_rst_GenAux.MoveNext
      Loop
   End If
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
End Function

Private Sub fs_Cal_Prepago()
  pnl_MtoApl.Caption = gf_FormatoNumero(CDbl(Trim(pnl_SaldoTNC1.Caption)) + CDbl(Trim(pnl_SaldoTC1.Caption)) + CDbl(txt_InteresTNC.Text) + CDbl(txt_InteresTC.Text) + CDbl(txt_SegDes.Text) + CDbl(txt_SegInm.Text) + CDbl(txt_MontoPortes.Text) + CDbl(pnl_DeuPen.Caption) + CDbl(Trim(pnl_CapPbpPerdido.Caption)) + _
                       CDbl(Trim(pnl_IntPbpPerdido.Caption)) + CDbl(Trim(pnl_DevBBP.Caption)) + CDbl(txt_IntLeg.Text), 12, 2) & " "
  pnl_MtoPrepagar.Caption = gf_FormatoNumero(CDbl(pnl_MtoApl.Caption) + CDbl(txt_MontoITF.Text), 12, 2) & " "
End Sub

Private Sub fs_Insert_ppgcab()
   'validar que no exista otro registro igual
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT PPGCAB_NUMOPE "
   g_str_Parame = g_str_Parame & "  FROM CRE_PPGCAB "
   g_str_Parame = g_str_Parame & " WHERE PPGCAB_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "  AND PPGCAB_FECPPG = '" & ipp_FecPre.Year & IIf(Len(Trim(ipp_FecPre.Month)) = 1, 0 & ipp_FecPre.Month, ipp_FecPre.Month) & IIf(Len(Trim(ipp_FecPre.Day)) = 1, 0 & ipp_FecPre.Day, ipp_FecPre.Day) & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Screen.MousePointer = 0
      moddat_g_int_CntErr = 1
      Exit Sub
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      moddat_g_int_CntErr = 0
   Else
      moddat_g_int_CntErr = 1
      MsgBox "El número de operación y la fecha del prepago ya existen, vuelva a ingresar otra fecha.", vbExclamation, modgen_g_con_OpeTra
      Call gs_SetFocus(ipp_FecPre)
   End If
   
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
   
   'inserta prepago total
   If moddat_g_int_CntErr = 0 Then
      Call fs_usp_cre_ppgcab
   End If
End Sub

Private Sub fs_usp_cre_ppgcab()
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "usp_cre_ppgcab ( "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
   g_str_Parame = g_str_Parame & ipp_FecPre.Year & IIf(Len(Trim(ipp_FecPre.Month)) = 1, 0 & ipp_FecPre.Month, ipp_FecPre.Month) & IIf(Len(Trim(ipp_FecPre.Day)) = 1, 0 & ipp_FecPre.Day, ipp_FecPre.Day) & ", "
   g_str_Parame = g_str_Parame & 2 & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & CDbl(pnl_Val_AsgInm.Caption) & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & CDbl(pnl_SaldoTNC1.Caption) & ", "
   g_str_Parame = g_str_Parame & IIf(InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0, 0, CDbl(pnl_SaldoTC1.Caption)) & ", "
   'Fecha del Ultimo Pago Realizado del TNC
   g_str_Parame = g_str_Parame & Year(CDate(pnl_UltPagTNC.Caption)) & IIf(Len(Trim(Month(CDate(pnl_UltPagTNC.Caption)))) = 1, 0 & Month(CDate(pnl_UltPagTNC.Caption)), Month(CDate(pnl_UltPagTNC.Caption))) & IIf(Len(Trim(Day(CDate(pnl_UltPagTNC.Caption)))) = 1, 0 & Day(CDate(pnl_UltPagTNC.Caption)), Day(CDate(pnl_UltPagTNC.Caption))) & ", "
   'Fecha del Ultimo Pago Realizado del TC
   If InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then
      g_str_Parame = g_str_Parame & 0 & ", "
   Else
      g_str_Parame = g_str_Parame & Year(CDate(pnl_UltPagTC.Caption)) & IIf(Len(Trim(Month(CDate(pnl_UltPagTC.Caption)))) = 1, 0 & Month(CDate(pnl_UltPagTC.Caption)), Month(CDate(pnl_UltPagTC.Caption))) & IIf(Len(Trim(Day(CDate(pnl_UltPagTC.Caption)))) = 1, 0 & Day(CDate(pnl_UltPagTC.Caption)), Day(CDate(pnl_UltPagTC.Caption))) & ", "
   End If
   g_str_Parame = g_str_Parame & CInt(pnl_DiasTNC.Caption) & ", "
   g_str_Parame = g_str_Parame & IIf(InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0, 0, CInt(pnl_DiasTC.Caption)) & ", "
   g_str_Parame = g_str_Parame & CDbl(txt_InteresTNC.Text) & ", "
   g_str_Parame = g_str_Parame & IIf(InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0, 0, CDbl(txt_InteresTC.Text)) & ", "
   g_str_Parame = g_str_Parame & CDbl(txt_SegDes.Text) & ", "
   g_str_Parame = g_str_Parame & CDbl(txt_SegInm.Text) & ", "
   g_str_Parame = g_str_Parame & CDbl(txt_MontoPortes.Text) & ", "
   g_str_Parame = g_str_Parame & CDbl(txt_MontoITF.Text) & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & CDbl(pnl_MtoApl.Caption) & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & CDbl(pnl_MtoPrepagar.Caption) & ", "
   g_str_Parame = g_str_Parame & CStr(cmb_MotPpg.ItemData(cmb_MotPpg.ListIndex)) & ", "
   g_str_Parame = g_str_Parame & "'" & Trim(txt_ObsPpg.Text) & "', "
   g_str_Parame = g_str_Parame & CDbl(pnl_CapPbpPerdido.Caption) & ", "
   g_str_Parame = g_str_Parame & CDbl(pnl_IntPbpPerdido.Caption) & ", "
   g_str_Parame = g_str_Parame & CDbl(pnl_DevBBP.Caption) & ", "
   g_str_Parame = g_str_Parame & CDbl(txt_IntLeg.Text) & ", "
   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                                       'Código Sucursal
   g_str_Parame = g_str_Parame & "1 ) "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      moddat_g_int_CntErr = 1
      MsgBox "No se pudo completar el procedimiento usp_cre_ppgcab.", vbCritical, modgen_g_str_NomPlt
   Else
      moddat_g_int_CntErr = 0
      If fs_update_hipmae = 0 Then MsgBox "El proceso se grabó exitosamente.", vbInformation, modgen_g_str_NomPlt
   End If
End Sub

Private Function fs_update_hipmae() As Integer
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "usp_ppgpar_cre_hipmae ( "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
   g_str_Parame = g_str_Parame & ipp_FecPre.Year & IIf(Len(Trim(ipp_FecPre.Month)) = 1, 0 & ipp_FecPre.Month, ipp_FecPre.Month) & IIf(Len(Trim(ipp_FecPre.Day)) = 1, 0 & ipp_FecPre.Day, ipp_FecPre.Day) & ", "
   g_str_Parame = g_str_Parame & "0,0,0,0,0,0,0,0,0,0,0,0,0, "
   '
   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                                       'Código Sucursal
   g_str_Parame = g_str_Parame & "2 ) "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 2) Then
      moddat_g_int_CntErr = 1
      fs_update_hipmae = 1
      MsgBox "No se pudo completar el procedimiento usp_ppgpar_cre_hipmae.", vbCritical, modgen_g_str_NomPlt
   Else
      fs_update_hipmae = 0
   End If
End Function

Private Sub fs_usp_cre_ppgsol()
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "usp_cre_ppgsol ( "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
   g_str_Parame = g_str_Parame & ipp_FecPre.Year & IIf(Len(Trim(ipp_FecPre.Month)) = 1, 0 & ipp_FecPre.Month, ipp_FecPre.Month) & IIf(Len(Trim(ipp_FecPre.Day)) = 1, 0 & ipp_FecPre.Day, ipp_FecPre.Day) & ", "
   g_str_Parame = g_str_Parame & 2 & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & CDbl(pnl_Val_AsgInm.Caption) & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & CDbl(pnl_SaldoTNC1.Caption) & ", "
   g_str_Parame = g_str_Parame & IIf(InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0, 0, CDbl(pnl_SaldoTC1.Caption)) & ", "
   'Fecha del Ultimo Pago Realizado del TNC
   g_str_Parame = g_str_Parame & Year(CDate(pnl_UltPagTNC.Caption)) & IIf(Len(Trim(Month(CDate(pnl_UltPagTNC.Caption)))) = 1, 0 & Month(CDate(pnl_UltPagTNC.Caption)), Month(CDate(pnl_UltPagTNC.Caption))) & IIf(Len(Trim(Day(CDate(pnl_UltPagTNC.Caption)))) = 1, 0 & Day(CDate(pnl_UltPagTNC.Caption)), Day(CDate(pnl_UltPagTNC.Caption))) & ", "
   'Fecha del Ultimo Pago Realizado del TC
   If InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then
      g_str_Parame = g_str_Parame & 0 & ", "
   Else
      g_str_Parame = g_str_Parame & Year(CDate(pnl_UltPagTC.Caption)) & IIf(Len(Trim(Month(CDate(pnl_UltPagTC.Caption)))) = 1, 0 & Month(CDate(pnl_UltPagTC.Caption)), Month(CDate(pnl_UltPagTC.Caption))) & IIf(Len(Trim(Day(CDate(pnl_UltPagTC.Caption)))) = 1, 0 & Day(CDate(pnl_UltPagTC.Caption)), Day(CDate(pnl_UltPagTC.Caption))) & ", "
   End If
   g_str_Parame = g_str_Parame & CInt(pnl_DiasTNC.Caption) & ", "
   g_str_Parame = g_str_Parame & IIf(InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0, 0, CInt(pnl_DiasTC.Caption)) & ", "
   g_str_Parame = g_str_Parame & CDbl(txt_InteresTNC.Text) & ", "
   g_str_Parame = g_str_Parame & IIf(InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0, 0, CDbl(txt_InteresTC.Text)) & ", "
   g_str_Parame = g_str_Parame & CDbl(txt_SegDes.Text) & ", "
   g_str_Parame = g_str_Parame & CDbl(txt_SegInm.Text) & ", "
   g_str_Parame = g_str_Parame & CDbl(txt_MontoPortes.Text) & ", "
   g_str_Parame = g_str_Parame & CDbl(txt_MontoITF.Text) & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & CDbl(pnl_MtoApl.Caption) & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & CDbl(pnl_MtoPrepagar.Caption) & ", "
   If cmb_MotPpg.ListIndex = -1 Then
      g_str_Parame = g_str_Parame & 0 & ", "
   Else
      g_str_Parame = g_str_Parame & CStr(cmb_MotPpg.ItemData(cmb_MotPpg.ListIndex)) & ", "
   End If
   g_str_Parame = g_str_Parame & "'" & Trim(txt_ObsPpg.Text) & "', "
   g_str_Parame = g_str_Parame & CDbl(pnl_CapPbpPerdido.Caption) & ", "
   g_str_Parame = g_str_Parame & CDbl(pnl_IntPbpPerdido.Caption) & ", "
   
   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                                       'Código Sucursal
   g_str_Parame = g_str_Parame & "1 ) "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      moddat_g_int_CntErr = 1
      MsgBox "No se pudo completar el procedimiento usp_cre_ppgsol.", vbCritical, modgen_g_str_NomPlt
   Else
      moddat_g_int_CntErr = 0
   End If
End Sub

'********************************
' LIQUIDACION CREDITO MIVIVIENDA
'********************************
Private Sub fs_PpgTot_Micasita(ByVal p_flg_guardar As Boolean, ByRef p_rut_Guardo As String)
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_NroFil     As Integer
   Dim r_int_ColumB     As Integer
   Dim r_int_ColumC     As Integer
   Dim r_int_ColumD     As Integer
   Dim r_int_ColumE     As Integer
   Dim r_int_ColumF     As Integer
   Dim r_int_ColumH     As Integer
   Dim r_int_ColumI     As Integer
   Dim r_int_ColumJ     As Integer
   Dim r_int_ColumK     As Integer
   Dim r_int_ColumL     As Integer
   Dim r_int_ColumM     As Integer
   Dim r_int_ColumO     As Integer
   Dim r_int_ColumP     As Integer

   r_int_NroFil = 3
   r_int_ColumB = 2
   r_int_ColumC = 3
   r_int_ColumD = 4
   r_int_ColumE = 5
   r_int_ColumF = 6
   r_int_ColumH = 8
   r_int_ColumI = 9
   r_int_ColumJ = 10
   r_int_ColumK = 11
   r_int_ColumL = 12
   r_int_ColumM = 13
   r_int_ColumO = 15
   r_int_ColumP = 16

   ''********************************************
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      'CENTRADO DE LA PAGINA
      '.PageSetup.CenterHorizontally = True
      .PageSetup.Orientation = xlLandscape

      'WIDTH
      .Columns("B").ColumnWidth = 50
      .Columns("C").ColumnWidth = 15
      
      'MARGENES
      .PageSetup.LeftMargin = Application.CentimetersToPoints(1)
      .PageSetup.RightMargin = Application.CentimetersToPoints(1)
      .PageSetup.TopMargin = Application.CentimetersToPoints(1)
      .PageSetup.BottomMargin = Application.CentimetersToPoints(1)

      .Range("A1:Y1").ColumnWidth = 6
      .Range("A8:A34").RowHeight = 12

      'BORDERS
      .Range("B3:R3").Borders(xlEdgeTop).Weight = xlMedium
      .Range("B4:R4").Borders(xlEdgeTop).Weight = xlMedium
      .Range("B5:R5").Borders(xlEdgeTop).Weight = xlMedium
      .Range("B6:R6").Borders(xlEdgeTop).Weight = xlMedium
      .Range("B34:R34").Borders(xlEdgeBottom).Weight = xlMedium
      .Range("B3:B34").Borders(xlEdgeLeft).Weight = xlMedium
      .Range("R3:R34").Borders(xlEdgeRight).Weight = xlMedium

      'Font
      .Range("A1:T45").Font.Name = "Arial"
      .Range("A1:T34").Font.Size = 9
      
      ' r_obj_Excel.Visible = True
      
      'Fecha de realizado el prepago
      .Range("L1") = "FECHA DE EMISIÓN: "
      .Range("L1").Font.Bold = True
      .Range("L1:O1").Merge
      .Range("P1:R1").Merge
      .Range("P1") = moddat_g_str_FecSis
      .Range("L1:O1").HorizontalAlignment = xlHAlignCenter
      .Range("P1:R1").HorizontalAlignment = xlHAlignCenter
      
      'Fila 3 - Titulo
      .Range("B3:R3").Merge
      .Range("B3") = "LIQUIDACION PREPAGO TOTAL - " & moddat_g_str_NomPrd
      .Range("B3:R3").HorizontalAlignment = xlHAlignCenter
      .Range("B3:R3").Font.Size = 12
      .Range("B5:R5").Font.Size = 12
      .Range("B7:R7").Font.Size = 12
      .Range("B3:R3").Font.Bold = True
      .Range("B5:R5").Font.Bold = True
      .Range("B7:R7").Font.Bold = True
      r_int_NroFil = r_int_NroFil + 2
      
      'Fila 5 - Datos Cliente
      .Cells(r_int_NroFil, r_int_ColumC) = "Cliente:"
      .Cells(r_int_NroFil, r_int_ColumC).HorizontalAlignment = xlHAlignLeft
      .Cells(r_int_NroFil, r_int_ColumE) = moddat_g_str_NomCli
      .Cells(r_int_NroFil, r_int_ColumO) = "DNI:"
      .Cells(r_int_NroFil, r_int_ColumO).HorizontalAlignment = xlHAlignLeft
      .Cells(r_int_NroFil, r_int_ColumP) = "'" & moddat_g_str_NumDoc
      .Cells(r_int_NroFil, r_int_ColumP).HorizontalAlignment = xlHAlignLeft
      r_int_NroFil = r_int_NroFil + 2
      
      'Fila 7 - Numero de Operacion, Moneda
      .Cells(r_int_NroFil, r_int_ColumC) = "N° de Operación:"
      .Cells(r_int_NroFil, r_int_ColumF) = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
      .Range("C" & r_int_NroFil & ":N" & r_int_NroFil & "").Font.Bold = True
      .Range("C" & r_int_NroFil & ":N" & r_int_NroFil & "").Font.Size = 11
      .Cells(r_int_NroFil, r_int_ColumK) = "Moneda:"
      .Cells(r_int_NroFil, r_int_ColumM) = moddat_g_str_Moneda
      r_int_NroFil = r_int_NroFil + 2
      
      'Fila 9 - Ultimo Pago, Saldo
      .Cells(r_int_NroFil, r_int_ColumC) = "Saldo al"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumI).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumI) = "'" & Format(CDate(pnl_UltPagTNC.Caption), "dd-mm-yy")
      
      .Cells(r_int_NroFil, r_int_ColumK) = Format(Me.pnl_SaldoTNC1.Caption, "###,###.00")  'moddat_g_dbl_SalCap
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumK).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
      .Range("K" & r_int_NroFil & ":M" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("K" & r_int_NroFil & ":M" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 1
      
      'Fila 10 - Fecha de Prepago
      .Cells(r_int_NroFil, r_int_ColumC) = "Fecha de corte"
      .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumI) = "'" & Format(CDate(ipp_FecPre), "dd-mm-yy")
      .Cells(r_int_NroFil, r_int_ColumI).HorizontalAlignment = xlHAlignRight
      r_int_NroFil = r_int_NroFil + 2
      
      'Fila 12 - Dias Interes
      .Cells(r_int_NroFil, r_int_ColumC) = "Días de interés"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumI) = CInt(pnl_DiasTNC.Caption)
      r_int_NroFil = r_int_NroFil + 2
      
      'Fila 14 - Valor inmueble
      .Cells(r_int_NroFil, r_int_ColumC) = "Valor asegurable inmueble"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumI) = Format(CDbl(pnl_Val_AsgInm.Caption), "###,###.00")
      r_int_NroFil = r_int_NroFil + 2
      
      'Fila 16 - Tasa interes anual
      .Cells(r_int_NroFil, r_int_ColumC) = "Tasa de interés anual (%)"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumI) = l_dbl_TasInt
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
      r_obj_Excel.Selection.NumberFormat = "###0.0000"
      r_int_NroFil = r_int_NroFil + 1
      
      'Fila 17 - Tasa seguro inmueble
      .Cells(r_int_NroFil, r_int_ColumC) = "Tasa de Seguro Inmueble (%)"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumI) = l_dbl_SegInm
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
      r_obj_Excel.Selection.NumberFormat = "###0.0000"
      r_int_NroFil = r_int_NroFil + 1
      
      'Fila 18 - Tasa seguro desgravamen
      .Cells(r_int_NroFil, r_int_ColumC) = "Tasa de Seguro Desgravamen (%)"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumI) = l_dbl_SegDes
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
      r_obj_Excel.Selection.NumberFormat = "###0.0000"
          
      r_int_NroFil = r_int_NroFil + 2
      
      'Fila 20 - Subtitulo gastos
      .Cells(r_int_NroFil, r_int_ColumC) = "Intereses, Seguros y Pendientes"
      .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
       
       'MARCO
      .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
      .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
      r_int_NroFil = r_int_NroFil + 1
      
      'Fila 21 - Portes
      .Cells(r_int_NroFil, r_int_ColumC) = "Portes"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumI) = CDbl(txt_MontoPortes.Text)
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
      r_obj_Excel.Selection.NumberFormat = "###0.00"
      r_int_NroFil = r_int_NroFil + 1
      
      'Fila 22 - Intereses
      .Cells(r_int_NroFil, r_int_ColumC) = "Intereses a la fecha"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumI) = CDbl(txt_InteresTNC.Text)
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
      r_obj_Excel.Selection.NumberFormat = "###0.00"
      r_int_NroFil = r_int_NroFil + 1
      
      'Fila 23 - Seguro desgravamen
      .Cells(r_int_NroFil, r_int_ColumC) = "Seguro Desgravamen"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumI) = CDbl(txt_SegDes.Text)
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
      r_obj_Excel.Selection.NumberFormat = "###0.00"
      r_int_NroFil = r_int_NroFil + 1
      
      'Fila 24 - Seguro inmueble
      .Cells(r_int_NroFil, r_int_ColumC) = "Seguro Inmueble"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumI) = CDbl(txt_SegInm.Text)
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
      r_obj_Excel.Selection.NumberFormat = "###0.00"
      r_int_NroFil = r_int_NroFil + 1
           
      'Linea - Deuda Pendiente
      .Cells(r_int_NroFil, r_int_ColumC) = "Deuda Pendiente"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumI) = CDbl(pnl_DeuPen.Caption)
      .Cells(r_int_NroFil, r_int_ColumI).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      r_int_NroFil = r_int_NroFil + 1
      
      'Devolución BBP
      .Cells(r_int_NroFil, r_int_ColumC) = "Devolución BBP"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumI) = CDbl(pnl_DevBBP.Caption)
      .Cells(r_int_NroFil, r_int_ColumI).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      r_int_NroFil = r_int_NroFil + 1
      
      'Interés Legal
      .Cells(r_int_NroFil, r_int_ColumC) = "Interés Legal"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumI) = CDbl(txt_IntLeg.Text)
      .Cells(r_int_NroFil, r_int_ColumI).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      r_int_NroFil = r_int_NroFil + 1
      
      'Fila 25 - Total gastos
      .Cells(r_int_NroFil, r_int_ColumC) = "Total a aplicar al saldo"
      .Range("C" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumK) = Format(CDbl(CDbl(txt_InteresTNC.Text) + CDbl(txt_MontoPortes.Text) + CDbl(txt_SegDes.Text) + CDbl(txt_SegInm.Text) + _
                                           CDbl(pnl_DeuPen.Caption) + CDbl(pnl_DevBBP.Caption) + CDbl(txt_IntLeg.Text)), "###,##0.00")
      .Range("C" & r_int_NroFil & ":J" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumK).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
      .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 2
      
      'Fila 29 - Total antes del ITF
      .Cells(r_int_NroFil, r_int_ColumC) = "Monto a prepagar antes de ITF"
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Range("C" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Range("K" & r_int_NroFil & ":M" & r_int_NroFil & "").Font.Bold = True
      .Range("C" & r_int_NroFil & ":J" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumK) = Format(CDbl(pnl_MtoApl.Caption), "###,###.00")
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumK).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
      .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 2
      
      'Fila 31 - Monto ITF
      .Cells(r_int_NroFil, r_int_ColumC) = "ITF (%)"
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumI) = "'0.0005"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumK) = Format$(CDbl(txt_MontoITF.Text), "##.00")
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Font.Bold = True
      r_int_NroFil = r_int_NroFil + 2
      
      'Fila 33 - Total Prepago
      .Cells(r_int_NroFil, r_int_ColumC) = "Monto total del prepago"
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Range("C" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Range("C" & r_int_NroFil & ":J" & r_int_NroFil & "").Font.Bold = True
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Font.Bold = True
      .Range("C" & r_int_NroFil & ":J" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumK) = Format$(CDbl(pnl_MtoPrepagar.Caption), "###,###.00")
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Select
      
      If CInt(moddat_g_int_TipMon) = 1 Then
         r_obj_Excel.Selection.NumberFormat = "$###,##0.00_);[Red]($###,##0.00)"
         .Cells(36, 2) = "'" & "- Esta liquidación es válida solo a la fecha de corte."
         .Cells(37, 2) = "'" & "- Cualquier gasto adicional que no contemple la presente liquidación, será informado al cliente y tendrá que ser "
         .Cells(38, 2) = "'" & "  cancelado antes de realizar el abono que consigna la presente liquidación."
         .Cells(39, 2) = "'" & "- Realizar el depósito en la cuenta Nº 0011-0369-02-00090532 del BBVA Banco Continental."
         .Cells(40, 2) = "'" & "- Realizar el depósito en la cuenta Nº 011-369-000200090532-69 del BBVA Banco Continental desde otro Banco."
         .Range(.Cells(36, 2), .Cells(40, 2)).Font.Size = 11
         .Range(.Cells(36, 2), .Cells(40, 2)).Font.Bold = True
      Else
         r_obj_Excel.Selection.NumberFormat = "[$$]#,##0.00;[Red][$$]#,##0.00"
         .Cells(36, 2) = "'" & "- Esta liquidación es válida solo a la fecha de corte."
         .Cells(37, 2) = "'" & "- Cualquier gasto adicional que no contemple la presente liquidación, será informado al cliente y tendrá que ser "
         .Cells(38, 2) = "'" & "  cancelado antes de realizar el abono que consigna la presente liquidación."
         .Cells(39, 2) = "'" & "- Realizar el depósito en la cuenta Nº 0011-0369-02-00090540 del BBVA Banco Continental."
         .Range(.Cells(36, 2), .Cells(39, 2)).Font.Size = 11
         .Range(.Cells(36, 2), .Cells(39, 2)).Font.Bold = True
      End If
      
      .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumK).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
      .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 2
   End With
   
   p_rut_Guardo = ""
   If p_flg_guardar = True Then
      p_rut_Guardo = Format(date, "yyyymmdd") & "_PPT_" & moddat_g_str_NumOpe & ".XLSX"
      r_obj_Excel.ActiveWorkbook.SaveAs (g_str_RutLog & "\" & p_rut_Guardo)
      
      r_obj_Excel.Application.Quit
      Set r_obj_Excel = Nothing
   Else
      r_obj_Excel.Visible = True
      Set r_obj_Excel = Nothing
   End If
   
End Sub

'********************************
' LIQUIDACION CREDITO MIVIVIENDA
'********************************
Private Sub fs_PpgTot_Mivivienda(ByVal p_flg_guardar As Boolean, ByRef p_rut_Guardo As String)
Dim r_obj_Excel      As Excel.Application
Dim r_int_NroFil     As Integer
Dim r_int_ColumC     As Integer
Dim r_int_ColumD     As Integer
Dim r_int_ColumE     As Integer
Dim r_int_ColumH     As Integer
Dim r_int_ColumI     As Integer
Dim r_int_ColumJ     As Integer
Dim r_int_ColumK     As Integer
Dim r_int_ColumL     As Integer
Dim r_int_ColumM     As Integer
Dim r_int_ColumN     As Integer
Dim r_int_ColumO     As Integer
Dim r_int_ColumP     As Integer
Dim r_int_ColumQ     As Integer

   r_int_NroFil = 3
   r_int_ColumC = 3
   r_int_ColumD = 4
   r_int_ColumE = 5
   r_int_ColumH = 8
   r_int_ColumI = 9
   r_int_ColumJ = 10
   r_int_ColumK = 11
   r_int_ColumL = 12
   r_int_ColumM = 13
   r_int_ColumN = 14
   r_int_ColumO = 15
   r_int_ColumP = 16
   r_int_ColumQ = 17
   ''********************************************
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      'CENTRADO DE LA PAGINA
      '.PageSetup.CenterHorizontally = True
      .PageSetup.Orientation = xlLandscape
      
      'WIDTH
      .Columns("B").ColumnWidth = 50
      .Columns("C").ColumnWidth = 15
      
      'MARGENES
      .PageSetup.LeftMargin = Application.CentimetersToPoints(1)
      .PageSetup.RightMargin = Application.CentimetersToPoints(1)
      .PageSetup.TopMargin = Application.CentimetersToPoints(1)
      .PageSetup.BottomMargin = Application.CentimetersToPoints(1)
      
      .Range("A1:A1").ColumnWidth = 2
      .Range("B1:Y1").ColumnWidth = 6
      .Range("A8:A33").RowHeight = 12.75
      
      'BORDERS
      .Range("B3:R3").Borders(xlEdgeTop).Weight = xlMedium
      .Range("B4:R4").Borders(xlEdgeTop).Weight = xlMedium
      .Range("B5:R5").Borders(xlEdgeTop).Weight = xlMedium
      .Range("B6:R6").Borders(xlEdgeTop).Weight = xlMedium
      .Range("B31:R34").Borders(xlEdgeBottom).Weight = xlMedium
      .Range("B3:B34").Borders(xlEdgeLeft).Weight = xlMedium
      .Range("R3:R34").Borders(xlEdgeRight).Weight = xlMedium
      
      'Font
      .Range("A1:T43").Font.Name = "Arial"
      .Range("A1:T34").Font.Size = 9
            
      'Fecha de realizado el prepago
      .Range("L1") = "FECHA DE EMISIÓN: "
      .Range("L1").Font.Bold = True
      .Range("L1:O1").Merge
      .Range("P1:R1").Merge
      .Range("P1") = moddat_g_str_FecSis
      .Range("L1:O1").HorizontalAlignment = xlHAlignCenter
      .Range("P1:R1").HorizontalAlignment = xlHAlignCenter
      
      'Fila 3 - Titulo
      .Range("B3:R3").Merge
      .Range("B3") = "LIQUIDACION PREPAGO TOTAL - " & moddat_g_str_NomPrd
      .Range("B3:R3").HorizontalAlignment = xlHAlignCenter
      .Range("B3:R3").Font.Size = 12
      .Range("B5:R5").Font.Size = 12
      .Range("B7:R7").Font.Size = 12
      .Range("B3:R3").Font.Bold = True
      .Range("B5:R5").Font.Bold = True
      .Range("B7:R7").Font.Bold = True
      r_int_NroFil = r_int_NroFil + 2
      
      'Fila 5 - Datos Cliente
      .Cells(r_int_NroFil, r_int_ColumC) = "Cliente:"
      .Cells(r_int_NroFil, r_int_ColumC).HorizontalAlignment = xlHAlignLeft
      .Cells(r_int_NroFil, r_int_ColumE) = moddat_g_str_NomCli
      .Cells(r_int_NroFil, r_int_ColumO) = "DNI:"
      .Cells(r_int_NroFil, r_int_ColumO).HorizontalAlignment = xlHAlignLeft
      .Cells(r_int_NroFil, r_int_ColumP) = "'" & moddat_g_str_NumDoc
      .Cells(r_int_NroFil, r_int_ColumP).HorizontalAlignment = xlHAlignLeft
      r_int_NroFil = r_int_NroFil + 2
      
      'Fila 7 - Numero Operacion
      .Cells(r_int_NroFil, r_int_ColumC) = "Operación:"
      .Cells(r_int_NroFil, r_int_ColumE) = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
      .Cells(r_int_NroFil, r_int_ColumK) = "Moneda:"
      .Cells(r_int_NroFil, r_int_ColumM) = moddat_g_str_Moneda
      r_int_NroFil = r_int_NroFil + 2
      
      'Fila 9 - Ultimo cuota TC, subtitulo saldo
      .Cells(r_int_NroFil, r_int_ColumC) = "Fecha de Desembolso o última cuota TC (A)"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumI) = "'" & Format(CDate(pnl_UltPagTC.Caption), "dd-mm-yy")
      .Cells(r_int_NroFil, r_int_ColumI).HorizontalAlignment = xlHAlignRight
      .Range("N" & r_int_NroFil & ":Q" & r_int_NroFil & "").Merge
      .Range("N" & r_int_NroFil & ":Q" & r_int_NroFil & "").Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumN) = "Saldo antes del prepago"
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
      .Range("N" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("N" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range("N" & r_int_NroFil & ":M" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
      r_int_NroFil = r_int_NroFil + 1
      
      'Fila 11 - Ultimo Pago TNC, Saldo Deuda, Saldo TNC
      .Cells(r_int_NroFil, r_int_ColumC) = "Saldo al (B)"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumI).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumI) = "'" & Format(CDate(pnl_UltPagTNC.Caption), "dd-mm-yy")
      
      .Cells(r_int_NroFil, r_int_ColumL) = Format(l_dbl_SalNco + l_dbl_SalCon, "###,###.00") 'moddat_g_dbl_SalCap
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
      
      .Range("N" & r_int_NroFil & ":O" & r_int_NroFil & "").Merge
      .Range("P" & r_int_NroFil & ":Q" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumN) = "TNC"
      .Cells(r_int_NroFil, r_int_ColumP) = Format(l_dbl_SalNco, "###,###.00")  'moddat_g_dbl_SalCap
      .Range("N" & r_int_NroFil & ":O" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumK).Font.Bold = True
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumK).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumL).Borders(xlEdgeRight).Weight = xlThin
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumO).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
      .Range("N" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("N" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 1
      
      'Fila 12 - Fecha Prepago, Saldo TC
      .Cells(r_int_NroFil, r_int_ColumC) = "Fecha de corte (fecha del prepago) (C)"
      .Range("C" & r_int_NroFil & ":L" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumI) = "'" & Format(CDate(ipp_FecPre), "dd-mm-yy")
      .Cells(r_int_NroFil, r_int_ColumI).HorizontalAlignment = xlHAlignRight
      
      .Range("N" & r_int_NroFil & ":O" & r_int_NroFil & "").Merge
      .Range("P" & r_int_NroFil & ":Q" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumN) = "TC"
      .Cells(r_int_NroFil, r_int_ColumP) = Format(l_dbl_SalCon, "###,###.00")
      .Range("N" & r_int_NroFil & ":O" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumO).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
      .Range("N" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 1
            
      .Range("N" & r_int_NroFil & ":O" & r_int_NroFil & "").Merge
      .Range("P" & r_int_NroFil & ":Q" & r_int_NroFil & "").Merge
      .Range("P" & r_int_NroFil & ":Q" & r_int_NroFil & "").Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumN) = "Total"
      .Cells(r_int_NroFil, r_int_ColumP) = Format(l_dbl_SalNco + l_dbl_SalCon, "###,###.00")
      .Range("N" & r_int_NroFil & ":O" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
      
      .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumO).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
      .Range("N" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 1

      'Fila 14 - Dias TNC
      .Cells(r_int_NroFil, r_int_ColumC) = "Días de interés TNC (C-B)"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumI) = CInt(pnl_DiasTNC.Caption)
      r_int_NroFil = r_int_NroFil + 1
      
      'Fila 15 - Dias TC
      .Cells(r_int_NroFil, r_int_ColumC) = "Días de interés TC (C-A)"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumI) = CInt(pnl_DiasTC.Caption)
      r_int_NroFil = r_int_NroFil + 2
      
      'Fila 17 - Valor inmueble
      .Cells(r_int_NroFil, r_int_ColumC) = "Valor asegurable inmueble"
      .Cells(r_int_NroFil, r_int_ColumI) = Format(CDbl(pnl_Val_AsgInm.Caption), "###,###.00")
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Font.Bold = True
      r_int_NroFil = r_int_NroFil + 2
      
      'Fila 19,20,21 - Tasas
      .Cells(r_int_NroFil, r_int_ColumC) = "Tasa de interés anual (%)"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumI) = l_dbl_TasInt
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
      r_obj_Excel.Selection.NumberFormat = "#,##0.0000"
      r_int_NroFil = r_int_NroFil + 1
      
      .Cells(r_int_NroFil, r_int_ColumC) = "Tasa de Seguro Inmueble (%)"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumI) = l_dbl_SegInm
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
      r_obj_Excel.Selection.NumberFormat = "#,##0.0000"
      r_int_NroFil = r_int_NroFil + 1
      
      .Cells(r_int_NroFil, r_int_ColumC) = "Tasa de Seguro Desgravamen (%)"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumI) = l_dbl_SegDes
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
      r_obj_Excel.Selection.NumberFormat = "#,##0.0000"
      r_int_NroFil = r_int_NroFil + 2
      
      'Fila 23 - Subtitulo gastos
      .Cells(r_int_NroFil, r_int_ColumC) = "Intereses, Seguros y Pendientes"
      .Range("C" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumL).Borders(xlEdgeRight).Weight = xlThin
      .Range("C" & r_int_NroFil & ":L" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C" & r_int_NroFil & ":L" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range("C" & r_int_NroFil & ":L" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
      r_int_NroFil = r_int_NroFil + 1
      
      'Fila 24 - Portes
      .Cells(r_int_NroFil, r_int_ColumC) = "Portes"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumI) = CDbl(txt_MontoPortes.Text)
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
      r_obj_Excel.Selection.NumberFormat = "#,##0.00"
      r_int_NroFil = r_int_NroFil + 1
      
      'Fila 25 - Interes TNC
      .Cells(r_int_NroFil, r_int_ColumC) = "Intereses TNC a la fecha"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumI) = CDbl(txt_InteresTNC.Text)
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
      r_obj_Excel.Selection.NumberFormat = "#,##0.00"
      r_int_NroFil = r_int_NroFil + 1
      
      'Fila 26 - Intereses TC
      .Cells(r_int_NroFil, r_int_ColumC) = "Intereses TC a la fecha"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumI) = CDbl(txt_InteresTC.Text)
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
      r_obj_Excel.Selection.NumberFormat = "#,##0.00"
      r_int_NroFil = r_int_NroFil + 1
      
      'Fila 27 - Seg. Desgravamen
      .Cells(r_int_NroFil, r_int_ColumC) = "Seguro Desgravamen"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumI) = CDbl(txt_SegDes.Text)
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
      r_obj_Excel.Selection.NumberFormat = "#,##0.00"
      r_int_NroFil = r_int_NroFil + 1
      
      'Fila 28 - Seg. Inmueble
      .Cells(r_int_NroFil, r_int_ColumC) = "Seguro Inmueble"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumI) = CDbl(txt_SegInm.Text)
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
      r_obj_Excel.Selection.NumberFormat = "#,##0.00"
      r_int_NroFil = r_int_NroFil + 1
      
      .Cells(r_int_NroFil, r_int_ColumC) = "Capital PBP Pendiente"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumI) = CDbl(pnl_CapPbpPerdido.Caption)
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
      r_obj_Excel.Selection.NumberFormat = "#,##0.00"
      r_int_NroFil = r_int_NroFil + 1
      
      .Cells(r_int_NroFil, r_int_ColumC) = "Interes PBP Pendiente"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumI) = CDbl(pnl_IntPbpPerdido.Caption)
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
      r_obj_Excel.Selection.NumberFormat = "#,##0.00"
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea - Deuda Pendiente
      .Cells(r_int_NroFil, r_int_ColumC) = "Deuda Pendiente"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumI) = CDbl(pnl_DeuPen.Caption)
      .Cells(r_int_NroFil, r_int_ColumI).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      r_int_NroFil = r_int_NroFil + 1
      
      'Devolución BBP
      .Cells(r_int_NroFil, r_int_ColumC) = "Devolución BBP"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumI) = CDbl(pnl_DevBBP.Caption)
      .Cells(r_int_NroFil, r_int_ColumI).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      r_int_NroFil = r_int_NroFil + 1
      
      'Interés Legal
      .Cells(r_int_NroFil, r_int_ColumC) = "Interés Legal"
      .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumI) = CDbl(txt_IntLeg.Text)
      .Cells(r_int_NroFil, r_int_ColumI).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      r_int_NroFil = r_int_NroFil + 1

      'Fila 29 - Total Gastos
      .Cells(r_int_NroFil, r_int_ColumC) = "Total de Interés, Seguros y Pendientes "
      .Range("C" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumK) = Format(CDbl(CDbl(txt_InteresTNC.Text) + CDbl(txt_InteresTC.Text) + CDbl(txt_MontoPortes.Text) + CDbl(txt_SegDes.Text) + CDbl(txt_SegInm.Text) + CDbl(pnl_CapPbpPerdido.Caption) + CDbl(pnl_IntPbpPerdido.Caption) + _
                                           CDbl(pnl_DevBBP.Caption) + CDbl(txt_IntLeg.Text) + CDbl(pnl_DeuPen.Caption)), "###,##0.00")
      .Range("C" & r_int_NroFil & ":J" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumK).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumL).Borders(xlEdgeRight).Weight = xlThin
      .Range("C" & r_int_NroFil & ":L" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C" & r_int_NroFil & ":L" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 2
      
      'Fila 31 - Total Prepago
      .Cells(r_int_NroFil, r_int_ColumC) = "Monto del Prepago Total"
      .Range("C" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Range("C" & r_int_NroFil & ":J" & r_int_NroFil & "").Font.Bold = True
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Font.Bold = True
      .Range("C" & r_int_NroFil & ":J" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumK) = Format(CDbl(pnl_MtoApl.Caption), "###,###.00")
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Select
      
      If CInt(moddat_g_int_TipMon) = 1 Then
         r_obj_Excel.Selection.NumberFormat = "$###,##0.00_);[Red]($###,##0.00)"
         .Cells(36, 2) = "'" & "- Esta liquidación es válida solo a la fecha de corte."
         .Cells(37, 2) = "'" & "- Cualquier gasto adicional que no contemple la presente liquidación, será informado al cliente y tendrá que ser "
         .Cells(38, 2) = "'" & "  cancelado antes de realizar el abono que consigna la presente liquidación."
         .Cells(39, 2) = "'" & "- Realizar el depósito en la cuenta Nº 0011-0369-02-00090532 del BBVA Banco Continental."
         .Cells(40, 2) = "'" & "- Realizar el depósito en la cuenta Nº 011-369-000200090532-69 del BBVA Banco Continental desde otro Banco."
         .Range(.Cells(36, 2), .Cells(40, 2)).Font.Size = 11
         .Range(.Cells(36, 2), .Cells(40, 2)).Font.Bold = True
      Else
         r_obj_Excel.Selection.NumberFormat = "[$$]#,##0.00;[Red][$$]#,##0.00"
         .Cells(36, 2) = "'" & "- Esta liquidación es válida solo a la fecha de corte."
         .Cells(37, 2) = "'" & "- Cualquier gasto adicional que no contemple la presente liquidación, será informado al cliente y tendrá que ser "
         .Cells(38, 2) = "'" & "  cancelado antes de realizar el abono que consigna la presente liquidación."
         .Cells(39, 2) = "'" & "- Realizar el depósito en la cuenta Nº 0011-0369-02-00090540 del BBVA Banco Continental."
         .Range(.Cells(36, 2), .Cells(39, 2)).Font.Size = 11
         .Range(.Cells(36, 2), .Cells(39, 2)).Font.Bold = True
      End If
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumK).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumL).Borders(xlEdgeRight).Weight = xlThin
      .Range("C" & r_int_NroFil & ":L" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C" & r_int_NroFil & ":L" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 2
   End With
   
   p_rut_Guardo = ""
   If p_flg_guardar = True Then
      p_rut_Guardo = Format(date, "yyyymmdd") & "_PPT_" & moddat_g_str_NumOpe & ".XLSX"
      r_obj_Excel.ActiveWorkbook.SaveAs (g_str_RutLog & "\" & p_rut_Guardo)
      
      r_obj_Excel.Application.Quit
      Set r_obj_Excel = Nothing
   Else
      r_obj_Excel.Visible = True
      Set r_obj_Excel = Nothing
   End If
   
End Sub
