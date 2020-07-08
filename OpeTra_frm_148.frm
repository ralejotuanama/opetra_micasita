VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Ges_CreHip_13 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3405
   ClientLeft      =   2040
   ClientTop       =   3030
   ClientWidth     =   11610
   Icon            =   "OpeTra_frm_148.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3405
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11640
      _Version        =   65536
      _ExtentX        =   20532
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
      Begin Threed.SSPanel SSPanel11 
         Height          =   1095
         Left            =   30
         TabIndex        =   9
         Top             =   2250
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin VB.TextBox txt_NumFia 
            Height          =   315
            Left            =   7590
            MaxLength       =   25
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   60
            Width           =   3225
         End
         Begin VB.ComboBox cmb_BanFia 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   3225
         End
         Begin VB.ComboBox cmb_MonFia 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   720
            Width           =   3225
         End
         Begin EditLib.fpDateTime ipp_FVcFia 
            Height          =   315
            Left            =   7590
            TabIndex        =   3
            Top             =   390
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
         Begin EditLib.fpDateTime ipp_FEmFia 
            Height          =   315
            Left            =   1680
            TabIndex        =   2
            Top             =   390
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
         Begin EditLib.fpDoubleSingle ipp_MtoFia 
            Height          =   315
            Left            =   7590
            TabIndex        =   5
            Top             =   720
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
         Begin VB.Label Label8 
            Caption         =   "Nro. Carta Fianza:"
            Height          =   285
            Left            =   5820
            TabIndex        =   15
            Top             =   60
            Width           =   1425
         End
         Begin VB.Label Label5 
            Caption         =   "Fecha Vcto.:"
            Height          =   315
            Left            =   5820
            TabIndex        =   14
            Top             =   390
            Width           =   1425
         End
         Begin VB.Label Label7 
            Caption         =   "Fecha Emisión:"
            Height          =   315
            Left            =   60
            TabIndex        =   13
            Top             =   390
            Width           =   1425
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Monto Fianza:"
            Height          =   285
            Index           =   1
            Left            =   5820
            TabIndex        =   12
            Top             =   720
            Width           =   1395
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Banco:"
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   11
            Top             =   60
            Width           =   1365
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Moneda Fianza:"
            Height          =   315
            Index           =   3
            Left            =   60
            TabIndex        =   10
            Top             =   720
            Width           =   1455
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   16
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
            Height          =   315
            Left            =   630
            TabIndex        =   23
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
         Begin Threed.SSPanel SSPanel15 
            Height          =   315
            Left            =   630
            TabIndex        =   24
            Top             =   330
            Width           =   4215
            _Version        =   65536
            _ExtentX        =   7435
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Cartas Fianza - Renovación"
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
            Left            =   60
            Picture         =   "OpeTra_frm_148.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel39 
         Height          =   645
         Left            =   30
         TabIndex        =   17
         Top             =   750
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_148.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10890
            Picture         =   "OpeTra_frm_148.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   18
         Top             =   1440
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
         Begin Threed.SSPanel pnl_NumOpe 
            Height          =   315
            Left            =   1680
            TabIndex        =   19
            Top             =   60
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
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   1680
            TabIndex        =   20
            Top             =   390
            Width           =   9795
            _Version        =   65536
            _ExtentX        =   17277
            _ExtentY        =   556
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
         Begin VB.Label Label3 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   22
            Top             =   390
            Width           =   1395
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Operación:"
            Height          =   315
            Left            =   60
            TabIndex        =   21
            Top             =   60
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_Ges_CreHip_13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_BanFia()      As moddat_tpo_Genera

Private Sub cmb_BanFia_Click()
   Call gs_SetFocus(txt_NumFia)
End Sub

Private Sub cmb_BanFia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_BanFia_Click
   End If
End Sub

Private Sub cmb_MonFia_Click()
   Call gs_SetFocus(ipp_MtoFia)
End Sub

Private Sub cmb_MonFia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_MonFia_Click
   End If
End Sub

Private Sub cmd_Grabar_Click()
   If cmb_BanFia.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Banco de la Fianza.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_BanFia)
      Exit Sub
   End If
   If Len(Trim(txt_NumFia.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Carta Fianza.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumFia)
      Exit Sub
   End If
   If CDate(ipp_FEmFia.Text) > date Then
      MsgBox "La Fecha de Emisión de la Carta Fianza no puede ser mayor a la Fecha Actual.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FEmFia)
      Exit Sub
   End If
   'Se comenta para que puedan registrar CF retrasadas
   'If CDate(ipp_FVcFia.Text) < date Then
   '   MsgBox "La Fecha de Vencimiento de la Carta Fianza no puede ser menor a la Fecha Actual.", vbExclamation, modgen_g_str_NomPlt
   '   Call gs_SetFocus(ipp_FVcFia)
   '   Exit Sub
   'End If
   If CDate(ipp_FVcFia.Text) < CDate(ipp_FEmFia.Text) Then
      MsgBox "La Fecha de Vencimiento de la Carta Fianza no puede ser menor a la Fecha de Emisión.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FVcFia)
      Exit Sub
   End If
   If cmb_MonFia.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Moneda de la Carta Fianza.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_MonFia)
      Exit Sub
   End If
   If ipp_MtoFia.Value = 0 Then
      MsgBox "Debe seleccionar el Monto de la Carta Fianza.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_MtoFia)
      Exit Sub
   End If
   If moddat_g_str_FecFia = "0" Then
      MsgBox "Debe realizar el registro inicial de la Carta Fianza en el modulo de Desembolso.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FEmFia)
      Exit Sub
   End If
   If frm_Ges_CreHip_12.l_str_Estado = "N" Then
      If CDate(ipp_FEmFia.Text) < CDate(gf_FormatoFecha(moddat_g_str_FecFia)) Then
         MsgBox "La Fecha de Emisión de la Carta Fianza no puede ser menor a la Fecha de Emisión de la Carta Fianza anterior.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_FEmFia)
         Exit Sub
      End If
   End If

   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   If frm_Ges_CreHip_12.l_str_Estado = "M" Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "USP_CRE_HIPFIA_ACTUALIZA ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
      g_str_Parame = g_str_Parame & "'" & Format(ipp_FEmFia.Text, "yyyymmdd") & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_BanFia(cmb_BanFia.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NumFia.Text & "', "
      g_str_Parame = g_str_Parame & "'" & Format(ipp_FVcFia.Text, "yyyymmdd") & "', "
      g_str_Parame = g_str_Parame & "'" & CStr(cmb_MonFia.ItemData(cmb_MonFia.ListIndex)) & "', "
      g_str_Parame = g_str_Parame & "'" & CStr(CDbl(ipp_MtoFia.Text)) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         MsgBox "No se pudo ejecutar el procedimiento.", vbExclamation, modgen_g_str_NomPlt
      End If
            
      MsgBox "La grabacion se realizo correctamente", vbInformation, modgen_g_str_NomPlt
      
      frm_Ges_CreHip_12.fs_Buscar
      
   ElseIf frm_Ges_CreHip_12.l_str_Estado = "N" Then
      'Cambiando Situación a Fianza Actual
      moddat_g_int_FlgGOK = False
      moddat_g_int_CntErr = 0
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " SELECT COUNT(*) as CTA FROM CRE_HIPFIA "
      g_str_Parame = g_str_Parame & "  WHERE HIPFIA_NUMOPE = '" & moddat_g_str_NumOpe & "'"
      g_str_Parame = g_str_Parame & "    AND HIPFIA_EMIFIA = '" & Format(ipp_FEmFia.Text, "yyyymmdd") & "' "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         MsgBox "No se pudo ejecutar la consulta.", vbExclamation, modgen_g_str_NomPlt
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         If g_rst_Princi!CTA > 0 Then
            MsgBox "La Fecha de Emisión ya ha sido registrada", vbInformation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_FEmFia)
            Exit Sub
         End If
      End If
      
      Do While moddat_g_int_FlgGOK = False
         Screen.MousePointer = 11
         Call moddat_gs_FecSis
         
         g_str_Parame = "USP_CRE_HIPFIA_RENUEVA ("
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_BanFia & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumFia & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_FecFia & "', "
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
               Exit Sub
            Else
               moddat_g_int_CntErr = 0
            End If
         End If
         
         Screen.MousePointer = 0
      Loop
   
      'Creando Nueva Fianza
      moddat_g_int_FlgGOK = False
      moddat_g_int_CntErr = 0
      
      Do While moddat_g_int_FlgGOK = False
         g_str_Parame = "USP_CRE_HIPFIA_NUEVA ("
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
         
         g_str_Parame = g_str_Parame & "'" & l_arr_BanFia(cmb_BanFia.ListIndex + 1).Genera_Codigo & "', "
         g_str_Parame = g_str_Parame & "'" & txt_NumFia.Text & "', "
         g_str_Parame = g_str_Parame & Format(CDate(ipp_FEmFia.Text), "yyyymmdd") & ", "
         g_str_Parame = g_str_Parame & Format(CDate(ipp_FVcFia.Text), "yyyymmdd") & ", "
         g_str_Parame = g_str_Parame & CStr(cmb_MonFia.ItemData(cmb_MonFia.ListIndex)) & ", "
         g_str_Parame = g_str_Parame & CStr(CDbl(ipp_MtoFia.Text)) & ", "
         g_str_Parame = g_str_Parame & "'" & CStr(moddat_g_int_TipDoc) & moddat_g_str_NumDoc & "', "
         
         'Datos de Auditoria
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                           'Código Usuario
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                           'Nombre Terminal
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                            'Nombre Ejecutable
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                           'Código Sucursal
         
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
   
      'Actualizando Tipo de Garantía Interna
      moddat_g_int_FlgGOK = False
      moddat_g_int_CntErr = 0
      
      Do While moddat_g_int_FlgGOK = False
         Screen.MousePointer = 11
         
         Call moddat_gs_FecSis
         
         g_str_Parame = "USP_CRE_HIPMAE_GARINT ("
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
         g_str_Parame = g_str_Parame & "4, "
         g_str_Parame = g_str_Parame & CStr(cmb_MonFia.ItemData(cmb_MonFia.ListIndex)) & ", "
         g_str_Parame = g_str_Parame & CStr(CDbl(ipp_MtoFia.Text)) & ", "
         g_str_Parame = g_str_Parame & "'" & txt_NumFia.Text & "', "
         g_str_Parame = g_str_Parame & "'" & l_arr_BanFia(cmb_BanFia.ListIndex + 1).Genera_Codigo & "', "
         
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
               Exit Sub
            Else
               moddat_g_int_CntErr = 0
            End If
         End If
         
         Screen.MousePointer = 0
      Loop
   
      moddat_g_int_FlgAct = 2
      MsgBox "Se renovo la Carta Fianza correctamente.", vbInformation, modgen_g_str_NomPlt
   End If
   
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   pnl_NumOpe.Caption = gf_Formato_NumOpe(moddat_g_str_NumOpe)
   pnl_NomCli.Caption = moddat_g_str_NomCli
   
   Call fs_Inicia
   Call fs_Limpia
   Call gs_CentraForm(Me)
      
   'Valida que no tenga registrada la garantia
'   If moddat_g_int_TipGar = 1 Or moddat_g_int_TipGar = 2 Then
'      MsgBox "Operación ya tiene registrada la HIPOTECA.", vbExclamation, modgen_g_str_NomPlt
'      cmd_Grabar.Enabled = False
'      Screen.MousePointer = 0
'      Exit Sub
'   End If
   
   If frm_Ges_CreHip_12.l_str_Estado = "M" Then
      ipp_FEmFia.Enabled = False
      frm_Ges_CreHip_13.cmb_BanFia.Text = frm_Ges_CreHip_12.grd_Listad.TextMatrix(frm_Ges_CreHip_12.grd_Listad.Row, 0)
      frm_Ges_CreHip_13.txt_NumFia.Text = frm_Ges_CreHip_12.grd_Listad.TextMatrix(frm_Ges_CreHip_12.grd_Listad.Row, 1)
      frm_Ges_CreHip_13.ipp_FEmFia.Text = frm_Ges_CreHip_12.grd_Listad.TextMatrix(frm_Ges_CreHip_12.grd_Listad.Row, 2)
      frm_Ges_CreHip_13.ipp_FVcFia.Text = frm_Ges_CreHip_12.grd_Listad.TextMatrix(frm_Ges_CreHip_12.grd_Listad.Row, 3)
      
      If frm_Ges_CreHip_12.grd_Listad.TextMatrix(frm_Ges_CreHip_12.grd_Listad.Row, 4) = "S/." Then
         frm_Ges_CreHip_13.cmb_MonFia.ListIndex = 0
      Else
         frm_Ges_CreHip_13.cmb_MonFia.ListIndex = 1
      End If
      
      frm_Ges_CreHip_13.ipp_MtoFia.Text = frm_Ges_CreHip_12.grd_Listad.TextMatrix(frm_Ges_CreHip_12.grd_Listad.Row, 5)
   End If
   
   Call gs_SetFocus(cmb_BanFia)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte(cmb_BanFia, l_arr_BanFia, 1, "505")
   Call moddat_gs_Carga_LisIte_Combo(cmb_MonFia, 1, "204")
End Sub

Private Sub fs_Limpia()
   cmb_BanFia.ListIndex = -1
   txt_NumFia.Text = ""
   ipp_FEmFia.Text = Format(date, "dd/mm/yyyy")
   ipp_FVcFia.Text = Format(date, "dd/mm/yyyy")
   cmb_MonFia.ListIndex = -1
   ipp_MtoFia.Value = 0
End Sub

Private Sub ipp_FEmFia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FVcFia)
   End If
End Sub

Private Sub ipp_FVcFia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_MonFia)
   End If
End Sub

Private Sub ipp_MtoFia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub txt_NumFia_GotFocus()
   Call gs_SelecTodo(txt_NumFia)
End Sub

Private Sub txt_NumFia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_FEmFia.Enabled = True Then
         Call gs_SetFocus(ipp_FEmFia)
      Else
         Call gs_SetFocus(ipp_FVcFia)
      End If
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "._-")
   End If
End Sub
