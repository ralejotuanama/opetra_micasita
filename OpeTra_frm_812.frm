VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Rpt_OpeFin_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7905
   Icon            =   "OpeTra_frm_812.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   6485
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   7905
      _Version        =   65536
      _ExtentX        =   13944
      _ExtentY        =   11439
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
         Height          =   3435
         Left            =   30
         TabIndex        =   15
         Top             =   2170
         Width           =   7815
         _Version        =   65536
         _ExtentX        =   13785
         _ExtentY        =   6059
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
         Begin VB.CheckBox chk_TipMov 
            Caption         =   "Todas los Tipos de Movimientos"
            Height          =   315
            Left            =   1200
            TabIndex        =   8
            Top             =   2340
            Width           =   2685
         End
         Begin VB.ComboBox cmb_TipMov 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   2010
            Width           =   6585
         End
         Begin VB.CheckBox chk_SucAge 
            Caption         =   "Todas las Sucursales"
            Height          =   315
            Left            =   1200
            TabIndex        =   6
            Top             =   1680
            Width           =   2685
         End
         Begin VB.ComboBox cmb_SucAge 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1380
            Width           =   6585
         End
         Begin VB.CheckBox chk_CodBan 
            Caption         =   "Todos los Bancos"
            Height          =   315
            Left            =   1200
            TabIndex        =   2
            Top             =   390
            Width           =   2685
         End
         Begin VB.ComboBox cmb_CodBan 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   90
            Width           =   6585
         End
         Begin VB.CheckBox chk_CtaBan 
            Caption         =   "Todas las Cuentas"
            Height          =   315
            Left            =   1200
            TabIndex        =   4
            Top             =   1020
            Width           =   2685
         End
         Begin VB.ComboBox cmb_CtaBan 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   720
            Width           =   6585
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1200
            TabIndex        =   9
            Top             =   2700
            Width           =   1395
            _Version        =   196608
            _ExtentX        =   2461
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
         Begin EditLib.fpDateTime ipp_FecFin 
            Height          =   315
            Left            =   1200
            TabIndex        =   10
            Top             =   3030
            Width           =   1395
            _Version        =   196608
            _ExtentX        =   2461
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
         Begin VB.Label Label4 
            Caption         =   "Fecha Inicio:"
            Height          =   315
            Left            =   60
            TabIndex        =   21
            Top             =   2700
            Width           =   1005
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Fin:"
            Height          =   315
            Left            =   60
            TabIndex        =   20
            Top             =   3000
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo de Movimiento:"
            Height          =   465
            Left            =   60
            TabIndex        =   19
            Top             =   2010
            Width           =   945
         End
         Begin VB.Label Label1 
            Caption         =   "Sucursal:"
            Height          =   225
            Left            =   60
            TabIndex        =   18
            Top             =   1350
            Width           =   945
         End
         Begin VB.Label Label5 
            Caption         =   "Banco:"
            Height          =   225
            Left            =   60
            TabIndex        =   17
            Top             =   60
            Width           =   945
         End
         Begin VB.Label Label6 
            Caption         =   "Cuenta Bancaria:"
            Height          =   495
            Left            =   60
            TabIndex        =   16
            Top             =   690
            Width           =   945
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   705
         Left            =   30
         TabIndex        =   22
         Top             =   30
         Width           =   7815
         _Version        =   65536
         _ExtentX        =   13785
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
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   7200
            Top             =   90
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
            Height          =   585
            Left            =   630
            TabIndex        =   23
            Top             =   30
            Width           =   5655
            _Version        =   65536
            _ExtentX        =   9975
            _ExtentY        =   1032
            _StockProps     =   15
            Caption         =   "Reporte de Operaciones Financieras por Cuenta Bancaria"
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
            Picture         =   "OpeTra_frm_812.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel15 
         Height          =   645
         Left            =   30
         TabIndex        =   24
         Top             =   780
         Width           =   7815
         _Version        =   65536
         _ExtentX        =   13785
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
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_812.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_812.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   7200
            Picture         =   "OpeTra_frm_812.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel pnl_GenTmp 
         Height          =   765
         Left            =   30
         TabIndex        =   25
         Top             =   5650
         Width           =   7815
         _Version        =   65536
         _ExtentX        =   13785
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
         Begin Threed.SSPanel pnl_BarPro 
            Height          =   345
            Left            =   60
            TabIndex        =   26
            Top             =   360
            Width           =   7695
            _Version        =   65536
            _ExtentX        =   13573
            _ExtentY        =   609
            _StockProps     =   15
            Caption         =   "SSPanel4"
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
            FloodType       =   1
            FloodColor      =   49152
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Generando Información"
            Height          =   315
            Left            =   60
            TabIndex        =   27
            Top             =   60
            Width           =   7665
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   675
         Left            =   30
         TabIndex        =   28
         Top             =   1470
         Width           =   7815
         _Version        =   65536
         _ExtentX        =   13785
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
         Begin VB.ComboBox cmb_TipRep 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   180
            Width           =   6585
         End
         Begin VB.Label Label8 
            Caption         =   "Tipo de Reporte:"
            Height          =   375
            Left            =   60
            TabIndex        =   29
            Top             =   150
            Width           =   990
         End
      End
   End
End
Attribute VB_Name = "frm_Rpt_OpeFin_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_SucAge()      As moddat_tpo_Genera
Dim l_arr_CodBan()      As moddat_tpo_Genera
Dim l_arr_CtaBan()      As moddat_tpo_Genera

Private Sub chk_CodBan_Click()
   If chk_CodBan.Value = 1 Then
      cmb_CodBan.ListIndex = -1
      cmb_CodBan.Enabled = False
      
      cmb_CtaBan.Clear
      cmb_CtaBan.Enabled = False
      chk_CtaBan.Value = 1
      chk_CtaBan.Enabled = False
      
      If cmb_SucAge.Enabled Then
         Call gs_SetFocus(cmb_SucAge)
      ElseIf cmb_TipMov.Enabled Then
         Call gs_SetFocus(cmb_TipMov)
      Else
         Call gs_SetFocus(ipp_FecIni)
      End If
   ElseIf chk_CodBan.Value = 0 Then
      cmb_CodBan.Enabled = True
      cmb_CtaBan.Enabled = True
      chk_CtaBan.Enabled = True
      
      chk_CtaBan.Value = 0
      
      Call gs_SetFocus(cmb_CodBan)
   End If
End Sub

Private Sub chk_CtaBan_Click()
   If chk_CtaBan.Value = 1 Then
      cmb_CtaBan.ListIndex = -1
      cmb_CtaBan.Enabled = False
      
      If cmb_SucAge.Enabled Then
         Call gs_SetFocus(cmb_SucAge)
      ElseIf cmb_TipMov.Enabled Then
         Call gs_SetFocus(cmb_TipMov)
      Else
         Call gs_SetFocus(ipp_FecIni)
      End If
   ElseIf chk_CtaBan.Value = 0 Then
      cmb_CtaBan.Enabled = True
      
      Call gs_SetFocus(cmb_CtaBan)
   End If
End Sub

Private Sub chk_SucAge_Click()
   
   If cmb_TipRep.ListIndex = -1 Then
      MsgBox "Debe seleccionar primero el tipo de reporte.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Or cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 3 Then
        If chk_SucAge.Value = 1 Then
             cmb_SucAge.ListIndex = -1
             cmb_SucAge.Enabled = False
                If cmb_TipMov.Enabled Then
              Call gs_SetFocus(cmb_TipMov)
           Else
              Call gs_SetFocus(ipp_FecIni)
           End If
        ElseIf chk_SucAge.Value = 0 Then
           cmb_SucAge.Enabled = True
           Call gs_SetFocus(cmb_SucAge)
        End If
      
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
        If chk_SucAge.Value = 1 Then
           cmb_SucAge.ListIndex = -1
           cmb_SucAge.Enabled = False
           Call gs_SetFocus(ipp_FecIni)
        ElseIf chk_SucAge.Value = 0 Then
           cmb_SucAge.Enabled = True
           Call gs_SetFocus(cmb_SucAge)
        End If
   End If
   
End Sub

Private Sub chk_TipMov_Click()
   If cmb_TipRep.ListIndex = -1 Then
      MsgBox "Debe seleccionar primero el tipo de reporte.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
    
    If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Or cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 3 Then
        If chk_TipMov.Value = 1 Then
           cmb_TipMov.ListIndex = -1
           cmb_TipMov.Enabled = False
           Call gs_SetFocus(ipp_FecIni)
        ElseIf chk_TipMov.Value = 0 Then
           cmb_TipMov.Enabled = True
           Call gs_SetFocus(cmb_TipMov)
        End If
       
    ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
        If chk_TipMov.Value = 1 Then
           cmb_TipMov.ListIndex = -1
           cmb_TipMov.Enabled = False
           If cmb_SucAge.Enabled Then
              Call gs_SetFocus(cmb_SucAge)
           Else
              Call gs_SetFocus(ipp_FecIni)
           End If
        ElseIf chk_TipMov.Value = 0 Then
           cmb_TipMov.Enabled = True
           Call gs_SetFocus(cmb_TipMov)
        End If
    End If
End Sub

Private Sub cmb_CodBan_Click()
   If cmb_CodBan.ListIndex > -1 Then
      Screen.MousePointer = 11
      Call moddat_gs_Carga_CtaBan(l_arr_CodBan(cmb_CodBan.ListIndex + 1).Genera_Codigo, cmb_CtaBan, l_arr_CtaBan)
      Screen.MousePointer = 0
   Else
      cmb_CtaBan.Clear
   End If
   
   If cmb_CtaBan.Enabled Then
      Call gs_SetFocus(cmb_CtaBan)
   ElseIf cmb_SucAge.Enabled Then
      Call gs_SetFocus(cmb_SucAge)
   ElseIf cmb_TipMov.Enabled Then
      Call gs_SetFocus(cmb_TipMov)
   Else
      Call gs_SetFocus(ipp_FecIni)
   End If
End Sub

Private Sub cmb_CodBan_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CodBan_Click
   End If
End Sub

Private Sub cmb_CtaBan_Click()
   If cmb_SucAge.Enabled Then
      Call gs_SetFocus(cmb_SucAge)
   ElseIf cmb_TipMov.Enabled Then
      Call gs_SetFocus(cmb_TipMov)
   Else
      Call gs_SetFocus(ipp_FecIni)
   End If
End Sub

Private Sub cmb_CtaBan_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CtaBan_Click
   End If
End Sub

Private Sub cmb_SucAge_Click()
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Or cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 3 Then
        If cmb_TipMov.Enabled Then
           Call gs_SetFocus(cmb_TipMov)
        Else
           Call gs_SetFocus(ipp_FecIni)
        End If
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
        Call gs_SetFocus(ipp_FecIni)
   End If
End Sub

Private Sub cmb_SucAge_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_SucAge_Click
   End If
End Sub

Private Sub cmb_TipMov_Click()
    If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Or cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 3 Then
       Call gs_SetFocus(ipp_FecIni)
    ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
       If cmb_SucAge.Enabled Then
           Call gs_SetFocus(cmb_SucAge)
       Else
           Call gs_SetFocus(ipp_FecIni)
       End If
    End If
End Sub

Private Sub cmb_TipMov_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipMov_Click
   End If
End Sub

Private Sub cmb_TipRep_Click()

 Call fs_Limpia
 
 If cmb_TipRep.ListIndex <> -1 Then
 
      Call moddat_gs_Carga_LisIte(cmb_CodBan, l_arr_CodBan, 1, "516")
      Call moddat_gs_Carga_SucAge(cmb_SucAge, l_arr_SucAge, moddat_g_str_Codigo)
      Call moddat_gs_Carga_LisIte_Combo(cmb_TipMov, 1, "301")
         
      If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Or cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
         
         chk_SucAge.Caption = "Todas las Sucursales"
         Label1.Caption = "Sucursal:"
         SSPanel7.Caption = "Reporte de Operaciones Financieras por Sucursal"
         Call gs_SetFocus(cmb_SucAge)
         
         chk_TipMov.Caption = "Todas los Tipos de Movimientos"
         Label2.Caption = "Tipo de Movimiento:"
         
         If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
            Label1.Top = 1350
            cmb_SucAge.Top = 1380
            chk_SucAge.Top = 1680
            Label2.Top = 2010
            cmb_TipMov.Top = 2010
            chk_TipMov.Top = 2340
         ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
            Label1.Top = 2010
            cmb_SucAge.Top = 2010
            chk_SucAge.Top = 2340
            Label2.Top = 1350
            cmb_TipMov.Top = 1380
            chk_TipMov.Top = 1680
            SSPanel7.Caption = "Reporte de Operaciones Financieras por Tipo de Movimiento"
            Call gs_SetFocus(cmb_TipMov)
         End If
                     
         cmb_CodBan.Enabled = False
         chk_CodBan.Enabled = False
         cmb_CtaBan.Enabled = False
         chk_CtaBan.Enabled = False
      Else
         SSPanel7.Caption = "Reporte de Operaciones Financieras por Cuenta Bancaria"
         Label1.Top = 1350
         cmb_SucAge.Top = 1380
         chk_SucAge.Top = 1680
         
         Label2.Top = 2010
         cmb_TipMov.Top = 2010
         chk_TipMov.Top = 2340
         
         cmb_CodBan.Enabled = True
         chk_CodBan.Enabled = True
         cmb_CtaBan.Enabled = True
         chk_CtaBan.Enabled = True
         Call gs_SetFocus(cmb_CodBan)
      End If
   End If
   
End Sub

Private Sub cmd_ExpExc_Click()
   
   If cmb_TipRep.ListIndex <> -1 Then
   
      If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
            If chk_SucAge.Value = 0 Then
               If cmb_SucAge.ListIndex = -1 Then
                  MsgBox "Debe seleccionar la Sucursal.", vbExclamation, modgen_g_str_NomPlt
                  Call gs_SetFocus(cmb_SucAge)
                  Exit Sub
               End If
            End If
            If chk_TipMov.Value = 0 Then
               If cmb_TipMov.ListIndex = -1 Then
                  MsgBox "Debe seleccionar el Tipo de Movimiento.", vbExclamation, modgen_g_str_NomPlt
                  Call gs_SetFocus(cmb_TipMov)
                  Exit Sub
               End If
            End If
            If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
               MsgBox "La Fecha de Fin es menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_FecFin)
               Exit Sub
            End If
      ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
            
            If chk_TipMov.Value = 0 Then
               If cmb_TipMov.ListIndex = -1 Then
                  MsgBox "Debe seleccionar el Tipo de Movimiento.", vbExclamation, modgen_g_str_NomPlt
                  Call gs_SetFocus(cmb_TipMov)
                  Exit Sub
               End If
            End If
            If chk_SucAge.Value = 0 Then
               If cmb_SucAge.ListIndex = -1 Then
                  MsgBox "Debe seleccionar la Sucursal.", vbExclamation, modgen_g_str_NomPlt
                  Call gs_SetFocus(cmb_SucAge)
                  Exit Sub
               End If
            End If
            If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
               MsgBox "La Fecha de Fin es menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_FecFin)
               Exit Sub
            End If
            
      ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 3 Then
      
            If chk_CodBan.Value = 0 Then
               If cmb_CodBan.ListIndex = -1 Then
                  MsgBox "Debe seleccionar el Banco.", vbExclamation, modgen_g_str_NomPlt
                  Call gs_SetFocus(cmb_CodBan)
                  
                  Exit Sub
               End If
               
               If chk_CtaBan.Value = 0 Then
                  If cmb_CtaBan.ListIndex = -1 Then
                     MsgBox "Debe seleccionar la Cuenta Bancaria.", vbExclamation, modgen_g_str_NomPlt
                     Call gs_SetFocus(cmb_CtaBan)
                     
                     Exit Sub
                  End If
               End If
            End If
            
            If chk_SucAge.Value = 0 Then
               If cmb_SucAge.ListIndex = -1 Then
                  MsgBox "Debe seleccionar la Sucursal.", vbExclamation, modgen_g_str_NomPlt
                  Call gs_SetFocus(cmb_SucAge)
                  
                  Exit Sub
               End If
            End If
            
            If chk_TipMov.Value = 0 Then
               If cmb_TipMov.ListIndex = -1 Then
                  MsgBox "Debe seleccionar el Tipo de Movimiento.", vbExclamation, modgen_g_str_NomPlt
                  Call gs_SetFocus(cmb_TipMov)
                  
                  Exit Sub
               End If
            End If
            
            If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
               MsgBox "La Fecha de Fin es menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_FecFin)
               
               Exit Sub
            End If
       End If
   Else
        MsgBox "Debe seleccionar Tipo de Reporte.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(cmb_TipRep)
        
        Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   Screen.MousePointer = 11
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Or cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
        Call fs_ExpExcel_SucTipMov
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 3 Then
       Call fs_ExpExcel_CtaBan
   End If
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Imprim_Click()

   If cmb_TipRep.ListIndex <> -1 Then
        If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
             If chk_SucAge.Value = 0 Then
                If cmb_SucAge.ListIndex = -1 Then
                   MsgBox "Debe seleccionar la Sucursal.", vbExclamation, modgen_g_str_NomPlt
                   Call gs_SetFocus(cmb_SucAge)
                   Exit Sub
                End If
             End If
             If chk_TipMov.Value = 0 Then
                If cmb_TipMov.ListIndex = -1 Then
                   MsgBox "Debe seleccionar el Tipo de Movimiento.", vbExclamation, modgen_g_str_NomPlt
                   Call gs_SetFocus(cmb_TipMov)
                   Exit Sub
                End If
             End If
             If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
                MsgBox "La Fecha de Fin es menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
                Call gs_SetFocus(ipp_FecFin)
                Exit Sub
             End If
        ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
             If chk_TipMov.Value = 0 Then
                If cmb_TipMov.ListIndex = -1 Then
                   MsgBox "Debe seleccionar el Tipo de Movimiento.", vbExclamation, modgen_g_str_NomPlt
                   Call gs_SetFocus(cmb_TipMov)
                   Exit Sub
                End If
             End If
             If chk_SucAge.Value = 0 Then
                If cmb_SucAge.ListIndex = -1 Then
                   MsgBox "Debe seleccionar la Sucursal.", vbExclamation, modgen_g_str_NomPlt
                   Call gs_SetFocus(cmb_SucAge)
                   Exit Sub
                End If
             End If
             If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
                MsgBox "La Fecha de Fin es menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
                Call gs_SetFocus(ipp_FecFin)
                Exit Sub
             End If
        ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 3 Then
             If chk_CodBan.Value = 0 Then
                If cmb_CodBan.ListIndex = -1 Then
                   MsgBox "Debe seleccionar el Banco.", vbExclamation, modgen_g_str_NomPlt
                   Call gs_SetFocus(cmb_CodBan)
                   
                   Exit Sub
                End If
                
                If chk_CtaBan.Value = 0 Then
                   If cmb_CtaBan.ListIndex = -1 Then
                      MsgBox "Debe seleccionar la Cuenta Bancaria.", vbExclamation, modgen_g_str_NomPlt
                      Call gs_SetFocus(cmb_CtaBan)
                      
                      Exit Sub
                   End If
                End If
             End If
             
             If chk_SucAge.Value = 0 Then
                If cmb_SucAge.ListIndex = -1 Then
                   MsgBox "Debe seleccionar la Sucursal.", vbExclamation, modgen_g_str_NomPlt
                   Call gs_SetFocus(cmb_SucAge)
                   
                   Exit Sub
                End If
             End If
             
             If chk_TipMov.Value = 0 Then
                If cmb_TipMov.ListIndex = -1 Then
                   MsgBox "Debe seleccionar el Tipo de Movimiento.", vbExclamation, modgen_g_str_NomPlt
                   Call gs_SetFocus(cmb_TipMov)
                   
                   Exit Sub
                End If
             End If
             
             If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
                MsgBox "La Fecha de Fin es menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
                Call gs_SetFocus(ipp_FecFin)
                
                Exit Sub
             End If
        End If
   Else
        MsgBox "Debe seleccionar el Tipo de Reporte.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(cmb_TipRep)
        Exit Sub
   End If
   
   If MsgBox("¿Está seguro de imprimir el Reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   Screen.MousePointer = 11
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Or cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
       Call fs_Rep_SucTipMov
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 3 Then
       Call fs_Rep_CtaBan
   End If
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call gs_CentraForm(Me)
   Call fs_Limpia
   Call fs_Inicio
   
   Screen.MousePointer = 0
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Imprim)
   End If
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFin)
   End If
End Sub

Private Sub fs_Inicio()
   moddat_g_str_Codigo = "000001"
   
   Call moddat_gs_Carga_LisIte(cmb_CodBan, l_arr_CodBan, 1, "516")
   
   Call moddat_gs_Carga_SucAge(cmb_SucAge, l_arr_SucAge, moddat_g_str_Codigo)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipMov, 1, "301")
   
   pnl_BarPro.FloodPercent = 0
   
   cmb_TipRep.AddItem "POR SUCURSAL"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 1
      
   'If modgen_g_int_TipUsu <> 20121 Then
      cmb_TipRep.AddItem "POR TIPO DE MOVIMIENTO"
      cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 2
   'End If
   cmb_TipRep.AddItem "POR CUENTA BANCARIA"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 3
   
   cmb_TipRep.ListIndex = -1
   
   
   ipp_FecIni.Text = Format(date, "dd/mm/yyyy")
   ipp_FecFin.Text = Format(date, "dd/mm/yyyy")
End Sub
Private Sub fs_Limpia()
   ipp_FecIni.Text = Format(date, "dd/mm/yyyy")
   ipp_FecFin.Text = Format(date, "dd/mm/yyyy")
   cmb_CodBan.Clear
   chk_CodBan.Value = 0
   cmb_CtaBan.Clear
   chk_CtaBan.Value = 0
   cmb_SucAge.Clear
   chk_SucAge.Value = 0
   cmb_TipMov.Clear
   chk_TipMov.Value = 0
End Sub
Private Sub fs_Rep_SucTipMov()
Dim r_rst_Princi     As ADODB.Recordset
Dim r_rst_Grabar     As ADODB.Recordset
Dim r_lng_TotReg     As Long
Dim r_lng_RegAct     As Long
Dim r_int_ValNeg     As Integer

   'Obteniendo Nro. de Registros
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT COUNT(*) AS TOTREG FROM OPE_CAJMOV WHERE "
   
   If chk_SucAge.Value = 0 Then
      g_str_Parame = g_str_Parame & "CAJMOV_SUCMOV = '" & l_arr_SucAge(cmb_SucAge.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   If chk_TipMov.Value = 0 Then
      g_str_Parame = g_str_Parame & "CAJMOV_TIPMOV = " & cmb_TipMov.ItemData(cmb_TipMov.ListIndex) & " AND "
   End If
   
   g_str_Parame = g_str_Parame & "CAJMOV_FECMOV >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "CAJMOV_FECMOV <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
       Exit Sub
   End If

   r_lng_TotReg = r_rst_Princi!TOTREG

   r_rst_Princi.Close
   Set r_rst_Princi = Nothing

   If r_lng_TotReg = 0 Then
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
     Call gs_SetFocus(ipp_FecIni)
     Exit Sub
   End If
   
   'Borrando Spool Anterior
   g_str_Parame = "DELETE FROM RPT_OPEFIN WHERE "
   g_str_Parame = g_str_Parame & "OPEFIN_CODTER = '" & modgen_g_str_NombPC & "' AND "
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
      g_str_Parame = g_str_Parame & " OPEFIN_NOMRPT = 'OPE_OPEFIN_01.RPT' "
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
      g_str_Parame = g_str_Parame & " OPEFIN_NOMRPT = 'OPE_OPEFIN_02.RPT' "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Grabar, 2) Then
       Exit Sub
   End If

   'Leyendo la Información
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM OPE_CAJMOV WHERE "
   
   If chk_SucAge.Value = 0 Then
      g_str_Parame = g_str_Parame & "CAJMOV_SUCMOV = '" & l_arr_SucAge(cmb_SucAge.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   If chk_TipMov.Value = 0 Then
      g_str_Parame = g_str_Parame & "CAJMOV_TIPMOV = " & cmb_TipMov.ItemData(cmb_TipMov.ListIndex) & " AND "
   End If
   g_str_Parame = g_str_Parame & "CAJMOV_FECMOV >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "CAJMOV_FECMOV <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "ORDER BY CAJMOV_NUMMOV ASC, CAJMOV_FECMOV ASC "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
       Exit Sub
   End If
   
   r_lng_RegAct = 1
   r_rst_Princi.MoveFirst
   Do While Not r_rst_Princi.EOF
      pnl_BarPro.FloodPercent = r_lng_RegAct / r_lng_TotReg * 100
   
      g_str_Parame = "INSERT INTO RPT_OPEFIN ("
      g_str_Parame = g_str_Parame & "OPEFIN_CODTER, "
      g_str_Parame = g_str_Parame & "OPEFIN_NOMRPT, "
      g_str_Parame = g_str_Parame & "OPEFIN_CODSUC, "
      g_str_Parame = g_str_Parame & "OPEFIN_NUMMOV, "
      g_str_Parame = g_str_Parame & "OPEFIN_FECMOV, "
      g_str_Parame = g_str_Parame & "OPEFIN_NOMSUC, "
      g_str_Parame = g_str_Parame & "OPEFIN_CODTMV, "
      g_str_Parame = g_str_Parame & "OPEFIN_TIPMOV, "
      g_str_Parame = g_str_Parame & "OPEFIN_DOCIDE, "
      g_str_Parame = g_str_Parame & "OPEFIN_NOMCLI, "
      g_str_Parame = g_str_Parame & "OPEFIN_NUMOPE, "
      g_str_Parame = g_str_Parame & "OPEFIN_NOMBAN, "
      g_str_Parame = g_str_Parame & "OPEFIN_NUMCTA, "
      g_str_Parame = g_str_Parame & "OPEFIN_NOMMON, "
      g_str_Parame = g_str_Parame & "OPEFIN_SIMMON, "
      g_str_Parame = g_str_Parame & "OPEFIN_IMPSOL, "
      g_str_Parame = g_str_Parame & "OPEFIN_ITFSOL, "
      g_str_Parame = g_str_Parame & "OPEFIN_TOTSOL, "
      g_str_Parame = g_str_Parame & "OPEFIN_IMPDOL, "
      g_str_Parame = g_str_Parame & "OPEFIN_ITFDOL, "
      g_str_Parame = g_str_Parame & "OPEFIN_TOTDOL, "
      g_str_Parame = g_str_Parame & "OPEFIN_FECINI, "
      g_str_Parame = g_str_Parame & "OPEFIN_FECFIN) "
      g_str_Parame = g_str_Parame & "VALUES ("
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      
      If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
            g_str_Parame = g_str_Parame & "'OPE_OPEFIN_01.RPT', "
      ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
            g_str_Parame = g_str_Parame & "'OPE_OPEFIN_02.RPT', "
      End If
      
      g_str_Parame = g_str_Parame & "'" & r_rst_Princi!CAJMOV_SUCMOV & "', "
      g_str_Parame = g_str_Parame & "'" & Mid(CStr(r_rst_Princi!CAJMOV_FECMOV), 3, 2) & Format(r_rst_Princi!CAJMOV_NUMMOV, "00000") & "', "
      g_str_Parame = g_str_Parame & "'" & gf_FormatoFecha(CStr(r_rst_Princi!CAJMOV_FECMOV)) & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_gf_ConsultaSucAge(moddat_g_str_Codigo, r_rst_Princi!CAJMOV_SUCMOV) & "', "
      g_str_Parame = g_str_Parame & "'" & CStr(r_rst_Princi!CAJMOV_TIPMOV) & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("301", CStr(r_rst_Princi!CAJMOV_TIPMOV)) & "', "
      g_str_Parame = g_str_Parame & "'" & CStr(r_rst_Princi!CAJMOV_TIPDOC) & "-" & Trim(r_rst_Princi!CAJMOV_NUMDOC) & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_gf_Buscar_NomCli(r_rst_Princi!CAJMOV_TIPDOC, Trim(r_rst_Princi!CAJMOV_NUMDOC)) & "', "
      
      If r_rst_Princi!CAJMOV_TIPMOV = 1101 Or r_rst_Princi!CAJMOV_TIPMOV = 2101 Then
         g_str_Parame = g_str_Parame & "'" & gf_Formato_NumSol(Trim(r_rst_Princi!CAJMOV_NUMOPE)) & "', "
      Else
         g_str_Parame = g_str_Parame & "'" & gf_Formato_NumOpe(Trim(r_rst_Princi!CAJMOV_NUMOPE)) & "', "
      End If
      
      If r_rst_Princi!CAJMOV_TIPMOV = 2101 Then
         r_int_ValNeg = -1
      Else
         r_int_ValNeg = 1
      End If
      
      g_str_Parame = g_str_Parame & IIf(Trim(r_rst_Princi!CAJMOV_CODBAN) = "000000" Or Len(Trim(r_rst_Princi!CAJMOV_CODBAN)) = 0, "'', ", "'" & moddat_gf_Consulta_ParDes("516", CStr(r_rst_Princi!CAJMOV_CODBAN)) & "', ")
      g_str_Parame = g_str_Parame & "'" & Trim(r_rst_Princi!CAJMOV_NUMCTA) & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("204", CStr(r_rst_Princi!CAJMOV_MONPAG)) & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("229", CStr(r_rst_Princi!CAJMOV_MONPAG)) & "', "
   
      If r_rst_Princi!CAJMOV_MONPAG = 1 Then
         g_str_Parame = g_str_Parame & CStr(r_rst_Princi!CAJMOV_IMPPAG * r_int_ValNeg) & ", "
         g_str_Parame = g_str_Parame & CStr(r_rst_Princi!CAJMOV_ITFIMP * r_int_ValNeg) & ", "
         g_str_Parame = g_str_Parame & CStr(r_rst_Princi!CAJMOV_IMPTOT * r_int_ValNeg) & ", "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      Else
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & CStr(r_rst_Princi!CAJMOV_IMPPAG * r_int_ValNeg) & ", "
         g_str_Parame = g_str_Parame & CStr(r_rst_Princi!CAJMOV_ITFIMP * r_int_ValNeg) & ", "
         g_str_Parame = g_str_Parame & CStr(r_rst_Princi!CAJMOV_IMPTOT * r_int_ValNeg) & ", "
      End If
      
      g_str_Parame = g_str_Parame & "'" & ipp_FecIni.Text & "', "
      g_str_Parame = g_str_Parame & "'" & ipp_FecFin.Text & "') "
      
      If Not gf_EjecutaSQL(g_str_Parame, r_rst_Grabar, 2) Then
          Exit Sub
      End If
      
      r_rst_Princi.MoveNext
      
      DoEvents
      r_lng_RegAct = r_lng_RegAct + 1
   Loop
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   
   'Se conecta al crystal report
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   'Se envia las tablas correspondientes en el orden que fueron utilizadas
   crp_Imprim.DataFiles(0) = UCase(moddat_g_str_EntDat) & ".RPT_OPEFIN"
      
   'Se pone la llamada del nombre del reporte y se escoge donde se destinara el reporte
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_OPEFIN_01.RPT"
      crp_Imprim.SelectionFormula = "{RPT_OPEFIN.OPEFIN_NOMRPT} = 'OPE_OPEFIN_01.RPT' AND "
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_OPEFIN_02.RPT"
      crp_Imprim.SelectionFormula = "{RPT_OPEFIN.OPEFIN_NOMRPT} = 'OPE_OPEFIN_02.RPT' AND "
   End If

   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_OPEFIN.OPEFIN_CODTER} = '" & modgen_g_str_NombPC & "' "
   
   pnl_BarPro.FloodPercent = 0
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Action = 1
End Sub

Private Sub fs_Rep_CtaBan()
   Dim r_rst_Princi     As ADODB.Recordset
   Dim r_rst_Grabar     As ADODB.Recordset
   Dim r_lng_TotReg     As Long
   Dim r_lng_RegAct     As Long
   Dim r_int_ValNeg     As Integer
   

   'Obteniendo Nro. de Registros
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT COUNT(*) AS TOTREG FROM OPE_CAJMOV WHERE "
   
   If chk_CodBan.Value = 0 Then
      g_str_Parame = g_str_Parame & "CAJMOV_CODBAN = '" & l_arr_CodBan(cmb_CodBan.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   If chk_CtaBan.Value = 0 Then
      g_str_Parame = g_str_Parame & "CAJMOV_NUMCTA = '" & l_arr_CtaBan(cmb_CtaBan.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   If chk_SucAge.Value = 0 Then
      g_str_Parame = g_str_Parame & "CAJMOV_SUCMOV = '" & l_arr_SucAge(cmb_SucAge.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   If chk_TipMov.Value = 0 Then
      g_str_Parame = g_str_Parame & "CAJMOV_TIPMOV = " & cmb_TipMov.ItemData(cmb_TipMov.ListIndex) & " AND "
   End If
   
   g_str_Parame = g_str_Parame & "CAJMOV_FECMOV >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "CAJMOV_FECMOV <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
       Exit Sub
   End If

   r_lng_TotReg = r_rst_Princi!TOTREG

   r_rst_Princi.Close
   Set r_rst_Princi = Nothing

   If r_lng_TotReg = 0 Then
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
     Call gs_SetFocus(ipp_FecIni)
     Exit Sub
   End If
   
   Screen.MousePointer = 11

   'Borrando Spool Anterior
   g_str_Parame = "DELETE FROM RPT_OPEFIN WHERE "
   g_str_Parame = g_str_Parame & "OPEFIN_CODTER = '" & modgen_g_str_NombPC & "' AND "
   g_str_Parame = g_str_Parame & "OPEFIN_NOMRPT = 'OPE_OPEFIN_03.RPT' "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Grabar, 2) Then
       Exit Sub
   End If

   'Leyendo la Información
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM OPE_CAJMOV WHERE "
   
   If chk_CodBan.Value = 0 Then
      g_str_Parame = g_str_Parame & "CAJMOV_CODBAN = '" & l_arr_CodBan(cmb_CodBan.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   If chk_CtaBan.Value = 0 Then
      g_str_Parame = g_str_Parame & "CAJMOV_NUMCTA = '" & l_arr_CtaBan(cmb_CtaBan.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   If chk_SucAge.Value = 0 Then
      g_str_Parame = g_str_Parame & "CAJMOV_SUCMOV = '" & l_arr_SucAge(cmb_SucAge.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   If chk_TipMov.Value = 0 Then
      g_str_Parame = g_str_Parame & "CAJMOV_TIPMOV = " & cmb_TipMov.ItemData(cmb_TipMov.ListIndex) & " AND "
   End If
   
   g_str_Parame = g_str_Parame & "CAJMOV_FECMOV >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "CAJMOV_FECMOV <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   
   g_str_Parame = g_str_Parame & "ORDER BY CAJMOV_NUMMOV ASC, CAJMOV_FECMOV ASC "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
       Exit Sub
   End If
   
   r_lng_RegAct = 1
   
   r_rst_Princi.MoveFirst
   Do While Not r_rst_Princi.EOF
      pnl_BarPro.FloodPercent = r_lng_RegAct / r_lng_TotReg * 100
   
      g_str_Parame = "INSERT INTO RPT_OPEFIN ("
      g_str_Parame = g_str_Parame & "OPEFIN_CODTER, "
      g_str_Parame = g_str_Parame & "OPEFIN_NOMRPT, "
      g_str_Parame = g_str_Parame & "OPEFIN_CODSUC, "
      g_str_Parame = g_str_Parame & "OPEFIN_NUMMOV, "
      g_str_Parame = g_str_Parame & "OPEFIN_FECMOV, "
      g_str_Parame = g_str_Parame & "OPEFIN_NOMSUC, "
      g_str_Parame = g_str_Parame & "OPEFIN_CODTMV, "
      g_str_Parame = g_str_Parame & "OPEFIN_TIPMOV, "
      g_str_Parame = g_str_Parame & "OPEFIN_DOCIDE, "
      g_str_Parame = g_str_Parame & "OPEFIN_NOMCLI, "
      g_str_Parame = g_str_Parame & "OPEFIN_NUMOPE, "
      g_str_Parame = g_str_Parame & "OPEFIN_NOMBAN, "
      g_str_Parame = g_str_Parame & "OPEFIN_NUMCTA, "
      g_str_Parame = g_str_Parame & "OPEFIN_NOMMON, "
      g_str_Parame = g_str_Parame & "OPEFIN_SIMMON, "
      g_str_Parame = g_str_Parame & "OPEFIN_IMPSOL, "
      g_str_Parame = g_str_Parame & "OPEFIN_ITFSOL, "
      g_str_Parame = g_str_Parame & "OPEFIN_TOTSOL, "
      g_str_Parame = g_str_Parame & "OPEFIN_IMPDOL, "
      g_str_Parame = g_str_Parame & "OPEFIN_ITFDOL, "
      g_str_Parame = g_str_Parame & "OPEFIN_TOTDOL, "
      g_str_Parame = g_str_Parame & "OPEFIN_FECINI, "
      g_str_Parame = g_str_Parame & "OPEFIN_FECFIN) "
      g_str_Parame = g_str_Parame & "VALUES ("
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'OPE_OPEFIN_03.RPT', "
      g_str_Parame = g_str_Parame & "'" & r_rst_Princi!CAJMOV_SUCMOV & "', "
      g_str_Parame = g_str_Parame & "'" & Mid(CStr(r_rst_Princi!CAJMOV_FECMOV), 3, 2) & Format(r_rst_Princi!CAJMOV_NUMMOV, "00000") & "', "
      g_str_Parame = g_str_Parame & "'" & gf_FormatoFecha(CStr(r_rst_Princi!CAJMOV_FECMOV)) & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_gf_ConsultaSucAge(moddat_g_str_Codigo, r_rst_Princi!CAJMOV_SUCMOV) & "', "
      g_str_Parame = g_str_Parame & "'" & CStr(r_rst_Princi!CAJMOV_TIPMOV) & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("301", CStr(r_rst_Princi!CAJMOV_TIPMOV)) & "', "
      g_str_Parame = g_str_Parame & "'" & CStr(r_rst_Princi!CAJMOV_TIPDOC) & "-" & Trim(r_rst_Princi!CAJMOV_NUMDOC) & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_gf_Buscar_NomCli(r_rst_Princi!CAJMOV_TIPDOC, Trim(r_rst_Princi!CAJMOV_NUMDOC)) & "', "
      
      If r_rst_Princi!CAJMOV_TIPMOV = 1101 Or r_rst_Princi!CAJMOV_TIPMOV = 2101 Then
         g_str_Parame = g_str_Parame & "'" & gf_Formato_NumSol(Trim(r_rst_Princi!CAJMOV_NUMOPE)) & "', "
      Else
         g_str_Parame = g_str_Parame & "'" & gf_Formato_NumOpe(Trim(r_rst_Princi!CAJMOV_NUMOPE)) & "', "
      End If
      
      If r_rst_Princi!CAJMOV_TIPMOV = 2101 Then
         r_int_ValNeg = -1
      Else
         r_int_ValNeg = 1
      End If
      
      g_str_Parame = g_str_Parame & IIf(Trim(r_rst_Princi!CAJMOV_CODBAN) = "000000" Or Len(Trim(r_rst_Princi!CAJMOV_CODBAN)) = 0, "'', ", "'" & moddat_gf_Consulta_ParDes("516", CStr(r_rst_Princi!CAJMOV_CODBAN)) & "', ")
      g_str_Parame = g_str_Parame & "'" & Trim(r_rst_Princi!CAJMOV_NUMCTA) & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("204", CStr(r_rst_Princi!CAJMOV_MONPAG)) & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("229", CStr(r_rst_Princi!CAJMOV_MONPAG)) & "', "
   
      If r_rst_Princi!CAJMOV_MONPAG = 1 Then
         g_str_Parame = g_str_Parame & CStr(r_rst_Princi!CAJMOV_IMPPAG * r_int_ValNeg) & ", "
         g_str_Parame = g_str_Parame & CStr(r_rst_Princi!CAJMOV_ITFIMP * r_int_ValNeg) & ", "
         g_str_Parame = g_str_Parame & CStr(r_rst_Princi!CAJMOV_IMPTOT * r_int_ValNeg) & ", "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      Else
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & CStr(r_rst_Princi!CAJMOV_IMPPAG * r_int_ValNeg) & ", "
         g_str_Parame = g_str_Parame & CStr(r_rst_Princi!CAJMOV_ITFIMP * r_int_ValNeg) & ", "
         g_str_Parame = g_str_Parame & CStr(r_rst_Princi!CAJMOV_IMPTOT * r_int_ValNeg) & ", "
      End If
      
      g_str_Parame = g_str_Parame & "'" & ipp_FecIni.Text & "', "
      g_str_Parame = g_str_Parame & "'" & ipp_FecFin.Text & "') "
      
      If Not gf_EjecutaSQL(g_str_Parame, r_rst_Grabar, 2) Then
          Exit Sub
      End If
      
      r_rst_Princi.MoveNext
      DoEvents
      
      r_lng_RegAct = r_lng_RegAct + 1
   Loop
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   
   'Se conecta al crystal report
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   'Se envia las tablas correspondientes en el orden que fueron utilizadas
   crp_Imprim.DataFiles(0) = UCase(moddat_g_str_EntDat) & ".RPT_OPEFIN"
      
   'Se pone la llamada del nombre del reporte y se escoge donde se destinara el reporte
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_OPEFIN_03.RPT"
        
   crp_Imprim.SelectionFormula = "{RPT_OPEFIN.OPEFIN_NOMRPT} = 'OPE_OPEFIN_03.RPT' AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_OPEFIN.OPEFIN_CODTER} = '" & modgen_g_str_NombPC & "' "
   
   pnl_BarPro.FloodPercent = 0
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Action = 1
End Sub

Private Sub fs_ExpExcel_SucTipMov()
Dim r_rst_Princi     As ADODB.Recordset
Dim r_lng_TotReg     As Long
Dim r_lng_RegAct     As Long
Dim r_int_ValNeg     As Integer
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer

   'Obteniendo Nro. de Registros
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT COUNT(*) AS TOTREG FROM OPE_CAJMOV WHERE "
   If chk_SucAge.Value = 0 Then
      g_str_Parame = g_str_Parame & "CAJMOV_SUCMOV = '" & l_arr_SucAge(cmb_SucAge.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   If chk_TipMov.Value = 0 Then
      g_str_Parame = g_str_Parame & "CAJMOV_TIPMOV = " & cmb_TipMov.ItemData(cmb_TipMov.ListIndex) & " AND "
   End If
   g_str_Parame = g_str_Parame & "CAJMOV_FECMOV >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "CAJMOV_FECMOV <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
       Exit Sub
   End If

   r_lng_TotReg = r_rst_Princi!TOTREG

   r_rst_Princi.Close
   Set r_rst_Princi = Nothing

   If r_lng_TotReg = 0 Then
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
     Call gs_SetFocus(ipp_FecIni)
     Exit Sub
   End If
   
   'Preparando Cabecera de Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM":                          .Columns("A").ColumnWidth = 6
      .Cells(1, 2) = "TIPO MOVIMIENTO":               .Columns("B").ColumnWidth = 35:     .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 3) = "SUCURSAL":                      .Columns("C").ColumnWidth = 25:     .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 4) = "NRO. MOVIMIENTO":               .Columns("D").ColumnWidth = 20:     .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 5) = "FECHA MOVIMIENTO":              .Columns("E").ColumnWidth = 20:     .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 6) = "DOC. IDENTIDAD":                .Columns("F").ColumnWidth = 17:     .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 7) = "NOMBRE CLIENTE":                .Columns("G").ColumnWidth = 50
      .Cells(1, 8) = "NRO. OPERACION":                .Columns("H").ColumnWidth = 20:     .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 9) = "BANCO":                         .Columns("I").ColumnWidth = 40:     .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 10) = "NUMERO CUENTA":                .Columns("J").ColumnWidth = 25:     .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 11) = "NRO. COMPROBANTE":             .Columns("K").ColumnWidth = 25:     .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 12) = "FECHA DE PAGO":                .Columns("L").ColumnWidth = 25:     .Columns("L").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 13) = "MONEDA":                       .Columns("M").ColumnWidth = 25:     .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 14) = "SUB-IMPORTE (S/.)":            .Columns("N").ColumnWidth = 20
      .Cells(1, 15) = "IMPORTE ITF (S/.)":            .Columns("O").ColumnWidth = 20
      .Cells(1, 16) = "IMPORTE TOTAL (S/.)":          .Columns("P").ColumnWidth = 20
      .Cells(1, 17) = "SUB-IMPORTE (US$)":            .Columns("Q").ColumnWidth = 20
      .Cells(1, 18) = "IMPORTE ITF (US$)":            .Columns("R").ColumnWidth = 20
      .Cells(1, 19) = "IMPORTE TOTAL (US$)":          .Columns("S").ColumnWidth = 20
      If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
            .Cells(1, 20) = "TIPO DE REGISTRO":             .Columns("T").ColumnWidth = 18:     .Columns("T").HorizontalAlignment = xlHAlignCenter
            .Range(.Cells(1, 1), .Cells(1, 20)).Font.Bold = True
            .Range(.Cells(1, 1), .Cells(1, 20)).HorizontalAlignment = xlHAlignCenter
      Else
            .Range(.Cells(1, 1), .Cells(1, 19)).Font.Bold = True
            .Range(.Cells(1, 1), .Cells(1, 19)).HorizontalAlignment = xlHAlignCenter
      End If
   End With

   'Leyendo la Información
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT A.CAJMOV_TIPMOV, A.CAJMOV_SUCMOV, A.CAJMOV_NUMMOV, A.CAJMOV_FECMOV, A.CAJMOV_NUMOPE, A.CAJMOV_CODBAN, "
   g_str_Parame = g_str_Parame & "       A.CAJMOV_NUMCTA, A.CAJMOV_IMPPAG, A.CAJMOV_ITFIMP, A.CAJMOV_IMPTOT, A.CAJMOV_TIPDOC, A.CAJMOV_NUMDOC, "
   g_str_Parame = g_str_Parame & "       A.CAJMOV_MONPAG, A.CAJMOV_FECDEP, A.CAJMOV_NUMCOM,"
   g_str_Parame = g_str_Parame & "       TRIM(B.DATGEN_APEPAT)||' '||TRIM(B.DATGEN_APEMAT)||' '||TRIM(B.DATGEN_NOMBRE) AS NOMBRE_CLIENTE, "
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
        g_str_Parame = g_str_Parame & " TRIM(F.PARDES_DESCRI) AS TIPO_REGISTRO, "
   End If
   g_str_Parame = g_str_Parame & "       TRIM(C.PARDES_DESCRI) AS TIPO_MOVIMIENTO, TRIM(D.SUCAGE_DESCRI) AS SUCURSAL, TRIM(E.PARDES_DESCRI) AS MONEDA "
   g_str_Parame = g_str_Parame & "  FROM OPE_CAJMOV A "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CLI_DATGEN B ON B.DATGEN_TIPDOC = A.CAJMOV_TIPDOC AND B.DATGEN_NUMDOC = A.CAJMOV_NUMDOC "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 301 AND C.PARDES_CODITE = A.CAJMOV_TIPMOV "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_SUCAGE D ON D.SUCAGE_CODEMP = " & moddat_g_str_Codigo & " AND D.SUCAGE_CODSUC = A.CAJMOV_SUCMOV "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 204 AND E.PARDES_CODITE = A.CAJMOV_MONPAG "
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
        g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES F ON F.PARDES_CODGRP = 239 AND F.PARDES_CODITE = A.CAJMOV_TIPREG "
   End If
   
   g_str_Parame = g_str_Parame & " WHERE A.CAJMOV_FECMOV >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "   AND A.CAJMOV_FECMOV <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   If chk_SucAge.Value = 0 Then
      g_str_Parame = g_str_Parame & "   AND A.CAJMOV_SUCMOV = '" & l_arr_SucAge(cmb_SucAge.ListIndex + 1).Genera_Codigo & "' "
   End If
   If chk_TipMov.Value = 0 Then
      g_str_Parame = g_str_Parame & "   AND A.CAJMOV_TIPMOV = " & cmb_TipMov.ItemData(cmb_TipMov.ListIndex) & " "
   End If
   g_str_Parame = g_str_Parame & "ORDER BY A.CAJMOV_TIPMOV ASC, A.CAJMOV_SUCMOV ASC, A.CAJMOV_NUMMOV ASC, A.CAJMOV_FECMOV ASC "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
       Exit Sub
   End If
   
   r_lng_RegAct = 1
   r_int_ConVer = 2
   
   r_rst_Princi.MoveFirst
   Do While Not r_rst_Princi.EOF
      pnl_BarPro.FloodPercent = r_lng_RegAct / r_lng_TotReg * 100
      If r_rst_Princi!CAJMOV_TIPMOV = 2101 Then
         r_int_ValNeg = -1
      Else
         r_int_ValNeg = 1
      End If
   
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(r_rst_Princi!TIPO_MOVIMIENTO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = Trim(r_rst_Princi!SUCURSAL)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = Mid(CStr(r_rst_Princi!CAJMOV_FECMOV), 3, 2) & Format(r_rst_Princi!CAJMOV_NUMMOV, "00000")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = CDate(gf_FormatoFecha(CStr(r_rst_Princi!CAJMOV_FECMOV)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = CStr(r_rst_Princi!CAJMOV_TIPDOC) & "-" & Trim(r_rst_Princi!CAJMOV_NUMDOC)
      If r_rst_Princi!CAJMOV_TIPMOV = 1105 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = moddat_gf_Buscar_NomCli_PlanAhorro(r_rst_Princi!CAJMOV_TIPDOC, Trim(r_rst_Princi!CAJMOV_NUMDOC & ""))
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = Trim(r_rst_Princi!NOMBRE_CLIENTE)
      End If
      
      If r_rst_Princi!CAJMOV_TIPMOV = 1101 Or r_rst_Princi!CAJMOV_TIPMOV = 2101 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = gf_Formato_NumSol(Trim(r_rst_Princi!CAJMOV_NUMOPE))
      ElseIf r_rst_Princi!CAJMOV_TIPMOV = 1105 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = "'" & CStr(Trim(r_rst_Princi!CAJMOV_NUMOPE))
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = gf_Formato_NumOpe(Trim(r_rst_Princi!CAJMOV_NUMOPE))
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = IIf(Trim(r_rst_Princi!CAJMOV_CODBAN) = "000000" Or Len(Trim(r_rst_Princi!CAJMOV_CODBAN)) = 0, "", moddat_gf_Consulta_ParDes("516", CStr(r_rst_Princi!CAJMOV_CODBAN)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = "'" & Trim(r_rst_Princi!CAJMOV_NUMCTA)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Trim(r_rst_Princi!CAJMOV_NUMCOM)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = CDate(gf_FormatoFecha(CStr(r_rst_Princi!CAJMOV_FECDEP)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Trim(r_rst_Princi!MONEDA)
      
      If r_rst_Princi!CAJMOV_MONPAG = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Format(r_rst_Princi!CAJMOV_IMPPAG * r_int_ValNeg, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(r_rst_Princi!CAJMOV_ITFIMP * r_int_ValNeg, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(r_rst_Princi!CAJMOV_IMPTOT * r_int_ValNeg, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = 0
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Format(r_rst_Princi!CAJMOV_IMPPAG * r_int_ValNeg, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = Format(r_rst_Princi!CAJMOV_ITFIMP * r_int_ValNeg, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = Format(r_rst_Princi!CAJMOV_IMPTOT * r_int_ValNeg, "###,###,##0.00")
      End If
      If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = Trim(r_rst_Princi!TIPO_REGISTRO)
      End If
      r_int_ConVer = r_int_ConVer + 1
      r_rst_Princi.MoveNext
      DoEvents
      r_lng_RegAct = r_lng_RegAct + 1
   Loop
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   Screen.MousePointer = 0
   
   pnl_BarPro.FloodPercent = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_ExpExcel_CtaBan()
   Dim r_rst_Princi     As ADODB.Recordset
   Dim r_lng_TotReg     As Long
   Dim r_lng_RegAct     As Long
   Dim r_int_ValNeg     As Integer
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer

   'Obteniendo Nro. de Registros
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT COUNT(*) AS TOTREG FROM OPE_CAJMOV WHERE "
   
   If chk_CodBan.Value = 0 Then
      g_str_Parame = g_str_Parame & "CAJMOV_CODBAN = '" & l_arr_CodBan(cmb_CodBan.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   If chk_CtaBan.Value = 0 Then
      g_str_Parame = g_str_Parame & "CAJMOV_NUMCTA = '" & l_arr_CtaBan(cmb_CtaBan.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   If chk_SucAge.Value = 0 Then
      g_str_Parame = g_str_Parame & "CAJMOV_SUCMOV = '" & l_arr_SucAge(cmb_SucAge.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   If chk_TipMov.Value = 0 Then
      g_str_Parame = g_str_Parame & "CAJMOV_TIPMOV = " & cmb_TipMov.ItemData(cmb_TipMov.ListIndex) & " AND "
   End If
   
   g_str_Parame = g_str_Parame & "CAJMOV_FECMOV >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "CAJMOV_FECMOV <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
       Exit Sub
   End If

   r_lng_TotReg = r_rst_Princi!TOTREG

   r_rst_Princi.Close
   Set r_rst_Princi = Nothing

   If r_lng_TotReg = 0 Then
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
     Call gs_SetFocus(ipp_FecIni)
     Exit Sub
   End If
   
   Screen.MousePointer = 11

   'Preparando Cabecera de Excel
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM":                          .Columns("A").ColumnWidth = 10
      .Cells(1, 2) = "BANCO":                         .Columns("B").ColumnWidth = 50
      .Cells(1, 3) = "NUMERO CUENTA":                 .Columns("C").ColumnWidth = 25:     .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 4) = "SUCURSAL":                      .Columns("D").ColumnWidth = 25:     .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 5) = "NRO. MOVIMIENTO":               .Columns("E").ColumnWidth = 20:     .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 6) = "FECHA MOVIMIENTO":              .Columns("F").ColumnWidth = 20:     .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 7) = "TIPO MOVIMIENTO":               .Columns("G").ColumnWidth = 30
      .Cells(1, 8) = "DOC. IDENTIDAD":                .Columns("H").ColumnWidth = 20:     .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 9) = "NOMBRE CLIENTE":                .Columns("I").ColumnWidth = 50
      .Cells(1, 10) = "NRO. OPERACION":               .Columns("J").ColumnWidth = 20:     .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 11) = "MONEDA":                       .Columns("K").ColumnWidth = 25:     .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 12) = "SUB-IMPORTE (S/.)":            .Columns("L").ColumnWidth = 20
      .Cells(1, 13) = "IMPORTE ITF (S/.)":            .Columns("M").ColumnWidth = 20
      .Cells(1, 14) = "IMPORTE TOTAL (S/.)":          .Columns("N").ColumnWidth = 20
      .Cells(1, 15) = "SUB-IMPORTE (US$)":            .Columns("O").ColumnWidth = 20
      .Cells(1, 16) = "IMPORTE ITF (US$)":            .Columns("P").ColumnWidth = 20
      .Cells(1, 17) = "IMPORTE TOTAL (US$)":          .Columns("Q").ColumnWidth = 20
      
      .Range(.Cells(1, 1), .Cells(1, 17)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 17)).HorizontalAlignment = xlHAlignCenter
   End With

   'Leyendo la Información
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM OPE_CAJMOV WHERE "
   
   If chk_CodBan.Value = 0 Then
      g_str_Parame = g_str_Parame & "CAJMOV_CODBAN = '" & l_arr_CodBan(cmb_CodBan.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   If chk_CtaBan.Value = 0 Then
      g_str_Parame = g_str_Parame & "CAJMOV_NUMCTA = '" & l_arr_CtaBan(cmb_CtaBan.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   If chk_SucAge.Value = 0 Then
      g_str_Parame = g_str_Parame & "CAJMOV_SUCMOV = '" & l_arr_SucAge(cmb_SucAge.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   If chk_TipMov.Value = 0 Then
      g_str_Parame = g_str_Parame & "CAJMOV_TIPMOV = " & cmb_TipMov.ItemData(cmb_TipMov.ListIndex) & " AND "
   End If
   
   g_str_Parame = g_str_Parame & "CAJMOV_FECMOV >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "CAJMOV_FECMOV <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   
   g_str_Parame = g_str_Parame & "ORDER BY CAJMOV_CODBAN ASC, CAJMOV_NUMCTA ASC, CAJMOV_SUCMOV ASC, CAJMOV_NUMMOV ASC, CAJMOV_FECMOV ASC "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
       Exit Sub
   End If
   
   r_lng_RegAct = 1
   r_int_ConVer = 2
   
   r_rst_Princi.MoveFirst
   Do While Not r_rst_Princi.EOF
      pnl_BarPro.FloodPercent = r_lng_RegAct / r_lng_TotReg * 100
   
      If r_rst_Princi!CAJMOV_TIPMOV = 2101 Then
         r_int_ValNeg = -1
      Else
         r_int_ValNeg = 1
      End If
   
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = IIf(Trim(r_rst_Princi!CAJMOV_CODBAN) = "000000" Or Len(Trim(r_rst_Princi!CAJMOV_CODBAN)) = 0, "", moddat_gf_Consulta_ParDes("516", CStr(r_rst_Princi!CAJMOV_CODBAN)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = "'" & Trim(r_rst_Princi!CAJMOV_NUMCTA)
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = moddat_gf_ConsultaSucAge(moddat_g_str_Codigo, r_rst_Princi!CAJMOV_SUCMOV)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Mid(CStr(r_rst_Princi!CAJMOV_FECMOV), 3, 2) & Format(r_rst_Princi!CAJMOV_NUMMOV, "00000")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = CDate(gf_FormatoFecha(CStr(r_rst_Princi!CAJMOV_FECMOV)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = moddat_gf_Consulta_ParDes("301", CStr(r_rst_Princi!CAJMOV_TIPMOV))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = CStr(r_rst_Princi!CAJMOV_TIPDOC) & "-" & Trim(r_rst_Princi!CAJMOV_NUMDOC)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = moddat_gf_Buscar_NomCli(r_rst_Princi!CAJMOV_TIPDOC, Trim(r_rst_Princi!CAJMOV_NUMDOC))
      
      If r_rst_Princi!CAJMOV_TIPMOV = 1101 Or r_rst_Princi!CAJMOV_TIPMOV = 2101 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = gf_Formato_NumSol(Trim(r_rst_Princi!CAJMOV_NUMOPE))
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = gf_Formato_NumOpe(Trim(r_rst_Princi!CAJMOV_NUMOPE))
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = moddat_gf_Consulta_ParDes("204", CStr(r_rst_Princi!CAJMOV_MONPAG))
      
      If r_rst_Princi!CAJMOV_MONPAG = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Format(r_rst_Princi!CAJMOV_IMPPAG * r_int_ValNeg, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(r_rst_Princi!CAJMOV_ITFIMP * r_int_ValNeg, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Format(r_rst_Princi!CAJMOV_IMPTOT * r_int_ValNeg, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = 0
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(r_rst_Princi!CAJMOV_IMPPAG * r_int_ValNeg, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(r_rst_Princi!CAJMOV_ITFIMP * r_int_ValNeg, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Format(r_rst_Princi!CAJMOV_IMPTOT * r_int_ValNeg, "###,###,##0.00")
      End If
   
      r_int_ConVer = r_int_ConVer + 1
   
      r_rst_Princi.MoveNext
      DoEvents
      
      r_lng_RegAct = r_lng_RegAct + 1
   Loop
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   
   pnl_BarPro.FloodPercent = 0
   
   r_obj_Excel.Visible = True
   
   Set r_obj_Excel = Nothing
End Sub
