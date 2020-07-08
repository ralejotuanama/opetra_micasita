VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Rpt_EvaTas_12 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   4800
   ClientLeft      =   12900
   ClientTop       =   2910
   ClientWidth     =   6030
   Icon            =   "OpeTra_frm_268.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   4785
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   6015
      _Version        =   65536
      _ExtentX        =   10610
      _ExtentY        =   8440
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
         Height          =   675
         Left            =   30
         TabIndex        =   12
         Top             =   30
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
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
            Height          =   285
            Left            =   660
            TabIndex        =   13
            Top             =   30
            Width           =   5235
            _Version        =   65536
            _ExtentX        =   9234
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Reporte de Solicitudes Evaluadas en Tasación del Inmueble"
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   285
            Left            =   660
            TabIndex        =   14
            Top             =   300
            Width           =   3345
            _Version        =   65536
            _ExtentX        =   5900
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Proyecto Inmobiliario"
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
            Picture         =   "OpeTra_frm_268.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   15
         Top             =   750
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
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
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_268.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   5340
            Picture         =   "OpeTra_frm_268.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_268.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   3540
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
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   795
         Left            =   30
         TabIndex        =   16
         Top             =   2280
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
         _ExtentY        =   1402
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
         Begin VB.ComboBox cmb_CodPry 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   60
            Width           =   4575
         End
         Begin VB.CheckBox chk_CodPry 
            Caption         =   "Todos los Proyectos"
            Height          =   315
            Left            =   1320
            TabIndex        =   3
            Top             =   420
            Width           =   3555
         End
         Begin VB.Label Label1 
            Caption         =   "Proyecto:"
            Height          =   255
            Left            =   60
            TabIndex        =   17
            Top             =   60
            Width           =   765
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   795
         Left            =   30
         TabIndex        =   18
         Top             =   1440
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
         _ExtentY        =   1402
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
         Begin VB.CheckBox chk_TipPry 
            Caption         =   "Todos los Tipos de Proyectos"
            Height          =   315
            Left            =   1320
            TabIndex        =   1
            Top             =   420
            Width           =   3555
         End
         Begin VB.ComboBox cmb_TipPry 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   4575
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo Proyecto:"
            Height          =   255
            Left            =   60
            TabIndex        =   19
            Top             =   60
            Width           =   1185
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   795
         Left            =   30
         TabIndex        =   20
         Top             =   3120
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
         _ExtentY        =   1402
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
         Begin VB.CheckBox chk_TipRes 
            Caption         =   "Todos los Resultados"
            Height          =   315
            Left            =   1320
            TabIndex        =   5
            Top             =   420
            Width           =   2685
         End
         Begin VB.ComboBox cmb_TipRes 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   60
            Width           =   4575
         End
         Begin VB.Label Label3 
            Caption         =   "Resultado:"
            Height          =   255
            Left            =   60
            TabIndex        =   21
            Top             =   60
            Width           =   765
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   22
         Top             =   3960
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
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
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1320
            TabIndex        =   6
            Top             =   60
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
            Text            =   "01/01/2008"
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
            Left            =   1320
            TabIndex        =   7
            Top             =   390
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
            Text            =   "01/01/2008"
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
         Begin VB.Label Label20 
            Caption         =   "Fecha de Fin:"
            Height          =   195
            Left            =   60
            TabIndex        =   24
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha de Inicio:"
            Height          =   225
            Left            =   60
            TabIndex        =   23
            Top             =   60
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frm_Rpt_EvaTas_12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_CodPry()       As moddat_tpo_Genera

Private Sub chk_CodPry_Click()
   If chk_CodPry.Value = 1 Then
      cmb_CodPry.ListIndex = -1
      cmb_CodPry.Enabled = False
      
      Call gs_SetFocus(cmd_Imprim)
   ElseIf chk_CodPry.Value = 0 Then
      cmb_CodPry.Enabled = True
      chk_CodPry.Enabled = True
      
      Call gs_SetFocus(cmb_CodPry)
   End If
End Sub

Private Sub chk_TipPry_Click()
   If chk_TipPry.Value = 1 Then
      cmb_TipPry.ListIndex = -1
      cmb_TipPry.Enabled = False
      
      cmb_CodPry.ListIndex = -1
      
      cmb_CodPry.Enabled = False
      chk_CodPry.Enabled = False
      
      Call gs_SetFocus(cmd_Imprim)
   ElseIf chk_TipPry.Value = 0 Then
      cmb_TipPry.Enabled = True
      
      cmb_CodPry.Enabled = True
      chk_CodPry.Enabled = True
      
      Call gs_SetFocus(cmb_TipPry)
   End If
End Sub

Private Sub chk_TipRes_Click()
   If chk_TipRes.Value = 1 Then
      cmb_TipRes.ListIndex = -1
      cmb_TipRes.Enabled = False
      Call gs_SetFocus(ipp_FecIni)
   ElseIf chk_TipRes.Value = 0 Then
      cmb_TipRes.Enabled = True
      Call gs_SetFocus(cmb_TipRes)
   End If
End Sub

Private Sub cmb_CodPry_Click()
   If cmb_TipRes.Enabled Then
      Call gs_SetFocus(cmb_TipRes)
   Else
      Call gs_SetFocus(ipp_FecIni)
   End If
End Sub

Private Sub cmb_CodPry_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CodPry_Click
   End If
End Sub

Private Sub cmb_TipPry_Click()
   If cmb_TipPry.ListIndex > -1 Then
      If cmb_TipPry.ItemData(cmb_TipPry.ListIndex) = 1 Then
         cmb_CodPry.ListIndex = -1
         cmb_CodPry.Enabled = False
         
         chk_CodPry.Value = 0
         chk_CodPry.Enabled = False
         
         Call gs_SetFocus(cmd_Imprim)
      Else
         Screen.MousePointer = 11
         
         If chk_CodPry.Value = 0 Then
            cmb_CodPry.Enabled = True
         End If
         
         chk_CodPry.Enabled = True
         
         If cmb_TipPry.ItemData(cmb_TipPry.ListIndex) = 2 Then
            Call modmip_gs_Carga_PryInm_Combo(cmb_CodPry, l_arr_CodPry, 2)
         Else
            Call modmip_gs_Carga_PryInm_Combo(cmb_CodPry, l_arr_CodPry, 1)
         End If
         
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_CodPry)
      End If
   End If
End Sub

Private Sub cmb_TipPry_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipPry_Click
   End If
End Sub

Private Sub cmb_TipRes_Click()
   Call gs_SetFocus(ipp_FecIni)
End Sub

Private Sub cmb_TipRes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipRes_Click
   End If
End Sub

Private Sub cmd_ExpExc_Click()
   Dim r_str_PrySel     As String
   Dim r_str_CodMod     As String
   Dim r_int_TipRes     As Integer
   
   If chk_TipPry.Value = 0 Then
      If cmb_TipPry.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Proyecto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipPry)
         Exit Sub
      End If
   End If
   
   If chk_CodPry.Value = 0 And chk_CodPry.Enabled Then
      If cmb_CodPry.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Proyecto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_CodPry)
         Exit Sub
      End If
   End If

   If chk_TipRes.Value = 0 Then
      If cmb_TipRes.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Resultado.", vbExclamation, modgen_g_str_NomPlt
         
         Call gs_SetFocus(cmb_TipRes)
         Exit Sub
      End If
   End If
   
   If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
      MsgBox "La Fecha de Fin no puede ser menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
      
      Call gs_SetFocus(ipp_FecIni)
      Exit Sub
   End If

   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   Screen.MousePointer = 11
   
   r_str_PrySel = ""
   r_str_CodMod = ""
   r_int_TipRes = 0
   
   If chk_CodPry.Value = 0 And chk_CodPry.Enabled Then
      r_str_PrySel = l_arr_CodPry(cmb_CodPry.ListIndex + 1).Genera_Codigo
   End If
   
   If chk_TipPry.Value = 0 Then
      r_str_CodMod = Format(cmb_TipPry.ItemData(cmb_TipPry.ListIndex), "00")
   End If

   If chk_TipRes.Value = 0 Then
      r_int_TipRes = cmb_TipRes.ItemData(cmb_TipRes.ListIndex)
   End If

   Screen.MousePointer = 11
   
   Call modmip_gs_Exc_EvaSol(41, 3, r_str_CodMod, r_int_TipRes, Format(CDate(ipp_FecIni.Text), "yyyymmdd"), Format(CDate(ipp_FecFin.Text), "yyyymmdd"), r_str_PrySel)
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Imprim_Click()
   Dim r_str_PrySel     As String
   Dim r_str_CodMod     As String
   Dim r_int_TipRes     As Integer
   
   If chk_TipPry.Value = 0 Then
      If cmb_TipPry.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Proyecto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipPry)
         Exit Sub
      End If
   End If
   
   If chk_CodPry.Value = 0 And chk_CodPry.Enabled Then
      If cmb_CodPry.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Proyecto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_CodPry)
         Exit Sub
      End If
   End If

   If chk_TipRes.Value = 0 Then
      If cmb_TipRes.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Resultado.", vbExclamation, modgen_g_str_NomPlt
         
         Call gs_SetFocus(cmb_TipRes)
         Exit Sub
      End If
   End If
   
   If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
      MsgBox "La Fecha de Fin no puede ser menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
      
      Call gs_SetFocus(ipp_FecIni)
      Exit Sub
   End If

   If MsgBox("¿Está seguro de imprimir el reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   Screen.MousePointer = 11
   
   r_str_PrySel = ""
   r_str_CodMod = ""
   r_int_TipRes = 0
   
   If chk_CodPry.Value = 0 And chk_CodPry.Enabled Then
      r_str_PrySel = l_arr_CodPry(cmb_CodPry.ListIndex + 1).Genera_Codigo
   End If
   
   If chk_TipPry.Value = 0 Then
      r_str_CodMod = Format(cmb_TipPry.ItemData(cmb_TipPry.ListIndex), "00")
   End If

   If chk_TipRes.Value = 0 Then
      r_int_TipRes = cmb_TipRes.ItemData(cmb_TipRes.ListIndex)
   End If
      
   Screen.MousePointer = 11
   
   Call modmip_gs_Rpt_SolEva("CRE_EVAHIP_12.RPT", 3, r_str_CodMod, 41, r_str_PrySel, r_int_TipRes, ipp_FecIni.Text, ipp_FecFin.Text)
   
   Screen.MousePointer = 0

   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   crp_Imprim.DataFiles(0) = "RPT_EVAHIP"
   
   crp_Imprim.SelectionFormula = "{RPT_EVAHIP.EVAHIP_NOMTER} = '" & modgen_g_str_NombPC & "' AND {RPT_EVAHIP.EVAHIP_NOMRPT} = 'CRE_EVAHIP_12.RPT'"
   
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CRE_EVAHIP_12.RPT"
   
   crp_Imprim.Action = 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Me.Caption = modgen_g_str_NomPlt
   
   Call gs_CentraForm(Me)

   ipp_FecIni.Text = Format(date - CDate(30), "dd/mm/yyyy")
   ipp_FecFin.Text = Format(date, "dd/mm/yyyy")

   Call moddat_gs_Carga_LisIte_Combo(cmb_TipPry, 1, "272")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipRes, 1, "273")
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




