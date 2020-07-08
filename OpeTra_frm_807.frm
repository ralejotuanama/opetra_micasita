VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Pro_SalCof_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10350
   Icon            =   "OpeTra_frm_807.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   6465
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   10365
      _Version        =   65536
      _ExtentX        =   18283
      _ExtentY        =   11404
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
         Height          =   1185
         Left            =   60
         TabIndex        =   9
         Top             =   2190
         Width           =   10245
         _Version        =   65536
         _ExtentX        =   18071
         _ExtentY        =   2090
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
         Begin VB.ComboBox cmb_PerMes 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   90
            Width           =   2685
         End
         Begin VB.ComboBox cmb_CodPro 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   750
            Width           =   6165
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1560
            TabIndex        =   2
            Top             =   420
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
            ButtonStyle     =   1
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
            Text            =   "2007"
            MaxValue        =   "9999"
            MinValue        =   "2007"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
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
            Caption         =   "Mes:"
            Height          =   315
            Left            =   150
            TabIndex        =   17
            Top             =   90
            Width           =   1245
         End
         Begin VB.Label Label4 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   150
            TabIndex        =   14
            Top             =   750
            Width           =   765
         End
         Begin VB.Label Label2 
            Caption         =   "Año:"
            Height          =   315
            Left            =   150
            TabIndex        =   10
            Top             =   420
            Width           =   765
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   705
         Left            =   60
         TabIndex        =   11
         Top             =   60
         Width           =   10245
         _Version        =   65536
         _ExtentX        =   18071
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   585
            Left            =   630
            TabIndex        =   12
            Top             =   90
            Width           =   8865
            _Version        =   65536
            _ExtentX        =   15637
            _ExtentY        =   1032
            _StockProps     =   15
            Caption         =   "Procesos - Carga Masiva"
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
            Left            =   90
            Picture         =   "OpeTra_frm_807.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   645
         Left            =   60
         TabIndex        =   13
         Top             =   810
         Width           =   10245
         _Version        =   65536
         _ExtentX        =   18071
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
            Left            =   9630
            Picture         =   "OpeTra_frm_807.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Proces 
            Height          =   585
            Left            =   60
            Picture         =   "OpeTra_frm_807.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Cargar Saldos COFIDE"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   2175
         Left            =   60
         TabIndex        =   15
         Top             =   3420
         Width           =   10245
         _Version        =   65536
         _ExtentX        =   18071
         _ExtentY        =   3836
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
         Begin VB.DriveListBox drv_LisUni 
            Height          =   315
            Left            =   6090
            TabIndex        =   22
            Top             =   60
            Width           =   4095
         End
         Begin VB.DirListBox dir_LisCar 
            Height          =   1665
            Left            =   6075
            TabIndex        =   4
            Top             =   420
            Width           =   4095
         End
         Begin VB.FileListBox fil_LisArc 
            Height          =   2040
            Left            =   1590
            TabIndex        =   5
            Top             =   90
            Width           =   4425
         End
         Begin VB.Label Label1 
            Caption         =   "Archivo a cargar:"
            Height          =   315
            Left            =   150
            TabIndex        =   16
            Top             =   90
            Width           =   1365
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   765
         Left            =   60
         TabIndex        =   18
         Top             =   5640
         Width           =   10245
         _Version        =   65536
         _ExtentX        =   18071
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
            Height          =   315
            Left            =   60
            TabIndex        =   23
            Top             =   390
            Width           =   10155
            _Version        =   65536
            _ExtentX        =   17912
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "SSPanel2"
            ForeColor       =   16777215
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
            Font3D          =   2
         End
         Begin VB.Label lbl_NomPro 
            Caption         =   "Proceso carga información Saldos COFIDE"
            Height          =   255
            Left            =   90
            TabIndex        =   19
            Top             =   120
            Width           =   5505
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   645
         Left            =   60
         TabIndex        =   20
         Top             =   1500
         Width           =   10245
         _Version        =   65536
         _ExtentX        =   18071
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
         Begin VB.ComboBox cmb_TipCar 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   180
            Width           =   6165
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Carga:"
            Height          =   195
            Left            =   150
            TabIndex        =   21
            Top             =   240
            Width           =   1050
         End
      End
   End
End
Attribute VB_Name = "frm_Pro_SalCof_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_bol_estado As Boolean

Private Type r_Arr_CargaSaldos
   CodCli               As String
   NomCli               As String
   NroCon               As String
   NroCon_C             As String
   FecExp               As String
   TipMon               As String
   ImpDes               As String
   PriFin               As String
   IntFin               As String
   ComFin               As String
   CuoPen               As String
   NroOri               As String
   ImpDes_TNC           As String
   PriFin_TNC           As String
   IntFin_TNC           As String
   ComFin_TNC           As String
   CuoPen_TNC           As String
   NroOri_TNC           As String
   ImpDes_TC            As String
   PriFin_TC            As String
   IntFin_TC            As String
   ComFin_TC            As String
   CuoPen_TC            As String
   NroOri_TC            As String
End Type

Dim r_str_Saldos()      As r_Arr_CargaSaldos
Private Type r_Arr_Calendario
   NumOpe               As String
   OpeMvi               As String
   Secuen               As String
   NroCuo               As String
   FecVct               As String
   NroDia               As String
   Moneda               As String
   TipCro               As Integer     '3 y 4
   Capita               As Double
   Intere               As Double
   ComCof               As Double
   MtoCuo               As Double
   SalCap               As Double
End Type

Dim r_str_CalDes()      As r_Arr_Calendario
Private Type r_Arr_DesErr
   NomArc               As String
   NumOpe               As String
   OpeMvi               As String
   MtoErr               As String
End Type
Dim r_str_DesErr()      As r_Arr_DesErr
Private Sub cmb_TipCar_Click()
   If cmb_TipCar.ItemData(cmb_TipCar.ListIndex) = 1 Then
      cmb_PerMes.Enabled = True
      ipp_PerAno.Enabled = True
      cmb_CodPro.Enabled = True
      lbl_NomPro.Caption = "Proceso carga información Saldos COFIDE"
   ElseIf cmb_TipCar.ItemData(cmb_TipCar.ListIndex) = 2 Then
      cmb_PerMes.Enabled = True
      ipp_PerAno.Enabled = True
      cmb_CodPro.Enabled = False
      ipp_PerAno.Value = 0
      lbl_NomPro.Caption = "Proceso carga información Cobranza COFIDE"
   ElseIf cmb_TipCar.ItemData(cmb_TipCar.ListIndex) = 3 Then
      cmb_PerMes.Enabled = False
      ipp_PerAno.Enabled = False
      cmb_CodPro.Enabled = False
      lbl_NomPro.Caption = "Proceso carga información Calendarios COFIDE"
      fil_LisArc.Pattern = "DET*.xls"
   End If
End Sub
Private Sub cmb_TipCar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_PerMes)
   End If
End Sub

Private Sub cmd_Proces_Click()
Dim r_lng_Contad           As Long
Dim r_str_NomArc           As String
Dim r_lng_NumReg           As Long
Dim r_lng_TotReg           As Long
Dim modprc_g_str_CadEje    As String

   If cmb_TipCar.ListIndex = -1 Then
      MsgBox "Debe seleccionar Tipo de Reporte.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipCar)
      Exit Sub
   End If

   If cmb_TipCar.ItemData(cmb_TipCar.ListIndex) = 1 Then
      If cmb_PerMes.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Mes.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_PerMes)
         Exit Sub
      End If
      If ipp_PerAno.Text = 0 Then
         MsgBox "Debe seleccionar Año.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_PerAno)
         Exit Sub
      End If
      If Me.cmb_CodPro.ListIndex = -1 Then
         MsgBox "Debe seleccionar un Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_CodPro)
         Exit Sub
      End If
      If Len(Trim(fil_LisArc.FileName & "")) = 0 Then
         MsgBox "Debe seleccionar el Archivo a cargar.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(fil_LisArc)
         Exit Sub
      End If
      
   ElseIf cmb_TipCar.ItemData(cmb_TipCar.ListIndex) = 2 Then
      If cmb_PerMes.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Mes.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_PerMes)
         Exit Sub
      End If
      If ipp_PerAno.Text = 0 Then
         MsgBox "Debe seleccionar Año.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_PerAno)
         Exit Sub
      End If
      If Len(Trim(fil_LisArc.FileName & "")) = 0 Then
         MsgBox "Debe seleccionar el Archivo a cargar.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(fil_LisArc)
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Está seguro de ejecutar el proceso?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   cmd_Proces.Enabled = False
   
   If cmb_TipCar.ItemData(cmb_TipCar.ListIndex) = 1 Then               'Carga Saldos COFIDE
      g_str_Parame = "SELECT NVL(COUNT(*),0) AS TOTREG FROM CRE_ARCCOF WHERE ARCCOF_PERMES = " & CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND ARCCOF_PERANO = " & CStr(ipp_PerAno.Text) & " AND ARCCOF_TIPARC = " & cmb_CodPro.ItemData(cmb_CodPro.ListIndex) & " "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      g_rst_Princi.MoveFirst
      r_lng_Contad = g_rst_Princi!TOTREG
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      If r_lng_Contad > 0 Then
         If MsgBox("La información de los Saldos COFIDE para este Período ya ha sido cargada. ¿Desea volver a cargar la información?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
            lbl_NomPro.Caption = "Eliminando información Saldos COFIDE...": DoEvents
            
            modprc_g_str_CadEje = "DELETE FROM CRE_ARCCOF WHERE "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "ARCCOF_PERMES = " & CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "ARCCOF_PERANO = " & CStr(ipp_PerAno.Text) & " AND "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "ARCCOF_TIPARC = " & cmb_CodPro.ItemData(cmb_CodPro.ListIndex) & " "
            
            If Not gf_EjecutaSQL(modprc_g_str_CadEje, g_rst_GenAux, 2) Then
               Exit Sub
            End If
            
            'g_rst_GenAux.Close
            'Set g_rst_GenAux = Nothing
            
            lbl_NomPro.Caption = "Proceso carga información Saldos COFIDE...": DoEvents
            Call fs_CargaCOFIDE_Trimestre(fil_LisArc.Path & "\" & fil_LisArc.FileName, cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text), cmb_CodPro.ItemData(cmb_CodPro.ListIndex), pnl_BarPro)
         End If
      Else
         lbl_NomPro.Caption = "Proceso carga información Saldos COFIDE...": DoEvents
         Call fs_CargaCOFIDE_Trimestre(fil_LisArc.Path & "\" & fil_LisArc.FileName, cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text), cmb_CodPro.ItemData(cmb_CodPro.ListIndex), pnl_BarPro)
      End If
      
      cmd_Proces.Enabled = True
      Screen.MousePointer = 0
         
      If l_bol_estado = False Then
         MsgBox "Proceso Terminado.", vbInformation, modgen_g_str_NomPlt
      Else
         MsgBox "No se encontró información.", vbInformation, modgen_g_str_NomPlt
      End If
      
   ElseIf cmb_TipCar.ItemData(cmb_TipCar.ListIndex) = 2 Then               'Carga Pagos Mensuales COFIDE
      
      g_str_Parame = "SELECT NVL(COUNT(*),0) AS TOTREG FROM CRE_ARCMEN WHERE ARCMEN_PERMES = " & CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND ARCMEN_PERANO = " & CStr(ipp_PerAno.Text)
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      g_rst_Princi.MoveFirst
      r_lng_Contad = g_rst_Princi!TOTREG
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   
      If r_lng_Contad > 0 Then
         If MsgBox("La información de cobranza COFIDE para este Período ya ha sido cargada. ¿Desea volver a cargar la información?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
            lbl_NomPro.Caption = "Eliminando información Cobranza COFIDE...": DoEvents
            
            modprc_g_str_CadEje = "DELETE FROM CRE_ARCMEN WHERE "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "ARCMEN_PERMES = " & CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "ARCMEN_PERANO = " & CStr(ipp_PerAno.Text)
            
            If Not gf_EjecutaSQL(modprc_g_str_CadEje, g_rst_GenAux, 2) Then
               Exit Sub
            End If
            
            lbl_NomPro.Caption = "Proceso carga información Cobranza COFIDE...": DoEvents
            Call fs_CargaCOFIDE_Mensual(fil_LisArc.Path & "\" & fil_LisArc.FileName, cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text), pnl_BarPro)
         End If
      Else
         lbl_NomPro.Caption = "Proceso carga información Cobranza COFIDE...": DoEvents
         Call fs_CargaCOFIDE_Mensual(fil_LisArc.Path & "\" & fil_LisArc.FileName, cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text), pnl_BarPro)
      End If
      
      cmd_Proces.Enabled = True
      Screen.MousePointer = 0
      MsgBox "Proceso Terminado.", vbInformation, modgen_g_str_NomPlt
   
   ElseIf cmb_TipCar.ItemData(cmb_TipCar.ListIndex) = 3 Then               'Carga Calendario Desembolso COFIDE
      
      ReDim r_str_DesErr(0)
      r_lng_NumReg = 0
      r_lng_TotReg = 0
      pnl_BarPro.FloodPercent = 0
   
      For r_lng_Contad = 0 To fil_LisArc.ListCount - 1
      
         r_lng_NumReg = r_lng_Contad
         r_lng_TotReg = (fil_LisArc.ListCount)
         r_str_NomArc = fil_LisArc.List(r_lng_Contad)
         
         If InStr(r_str_NomArc, "DET") > 0 Then
            Call fs_CargaCOFIDE_Calendario(fil_LisArc.Path & "\" & r_str_NomArc)
         End If
         
         pnl_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
      Next
      
      If UBound(r_str_DesErr) >= 1 Then
         Call fs_GenExc
      End If
      
      pnl_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
      cmd_Proces.Enabled = True
      Screen.MousePointer = 0
      MsgBox "Proceso Terminado.", vbInformation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_str_ParAux        As String
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
Dim r_int_NroFil        As Integer
Dim r_int_NoFlLi        As Integer
Dim r_lng_Contad        As Long

      r_int_NroFil = 3
      
      Set r_obj_Excel = New Excel.Application
      r_obj_Excel.SheetsInNewWorkbook = 1
      r_obj_Excel.Workbooks.Add
      
      With r_obj_Excel.ActiveSheet
      .Cells(1, 2) = "REPORTE DE CARGA DE CRONOGRAMAS"
      .Range(.Cells(1, 2), .Cells(1, 6)).Merge
      .Range(.Cells(1, 2), .Cells(1, 6)).Font.Bold = True
      .Range(.Cells(1, 2), .Cells(1, 6)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 2), .Cells(1, 6)).Font.Size = 14
      
      .Cells(r_int_NroFil, 2) = "ITEM"
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 2)).Merge
      .Cells(r_int_NroFil, 3) = "NOMBRE DE ARCHIVO"
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil + 1, 3)).Merge
      .Cells(r_int_NroFil, 4) = "OPERACION"
      .Range(.Cells(r_int_NroFil, 4), .Cells(r_int_NroFil + 1, 4)).Merge
      .Cells(r_int_NroFil, 5) = "NUMERO DE CONTRATO"
      .Range(.Cells(r_int_NroFil, 5), .Cells(r_int_NroFil + 1, 5)).Merge
      .Cells(r_int_NroFil, 6) = "DESCRIPCION"
      .Range(.Cells(r_int_NroFil, 6), .Cells(r_int_NroFil + 1, 6)).Merge
      
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 6)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 6)).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 6)).HorizontalAlignment = xlHAlignCenter
      '
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 12
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 30
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 15
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 15
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 60
      .Columns("F").HorizontalAlignment = xlHAlignLeft
      
      With .Range(.Cells(3, 2), .Cells(4, 6))
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
          .WrapText = True
      End With
      
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(2, 1), .Cells(99, 99)).Font.Size = 11
       
      r_int_NroFil = r_int_NroFil + 2
       
      For r_lng_Contad = 1 To UBound(r_str_DesErr)
         
         .Cells(r_int_NroFil, 2) = r_lng_Contad
         .Cells(r_int_NroFil, 3) = "'" & CStr(r_str_DesErr(r_lng_Contad).NomArc)
         .Cells(r_int_NroFil, 4) = "'" & CStr(r_str_DesErr(r_lng_Contad).NumOpe)
         .Cells(r_int_NroFil, 5) = CStr(r_str_DesErr(r_lng_Contad).OpeMvi)
         .Cells(r_int_NroFil, 6) = "'" & CStr(r_str_DesErr(r_lng_Contad).MtoErr)
         
         r_int_NroFil = r_int_NroFil + 1

      Next r_lng_Contad
      
      With .Range(.Cells(3, 2), .Cells(r_int_NroFil - 1, 6))
         .Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Borders(xlEdgeTop).LineStyle = xlContinuous
         .Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Borders(xlEdgeRight).LineStyle = xlContinuous
         .Borders(xlInsideVertical).LineStyle = xlContinuous
         .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
   End With
   
   r_obj_Excel.Visible = True
End Sub
Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub drv_LisUni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      On Error Resume Next
      dir_LisCar.Path = drv_LisUni.Drive
   End If
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_PerMes)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   cmb_PerMes.ListIndex = -1
   ipp_PerAno.Value = Year(date)
   
   cmb_TipCar.Clear
   cmb_TipCar.AddItem "CONCILIACION TRIMESTRAL SALDOS COFIDE"
   cmb_TipCar.ItemData(cmb_TipCar.NewIndex) = CInt(1)
   cmb_TipCar.AddItem "CONCILIACION PAGOS MENSUALES COFIDE"
   cmb_TipCar.ItemData(cmb_TipCar.NewIndex) = CInt(2)
   cmb_TipCar.AddItem "CALENDARIOS DESEMBOLSO COFIDE"
   cmb_TipCar.ItemData(cmb_TipCar.NewIndex) = CInt(3)
   cmb_TipCar.ListIndex = -1
   
   cmb_CodPro.Clear
   cmb_CodPro.AddItem "CREDITO CME"
   cmb_CodPro.ItemData(cmb_CodPro.NewIndex) = CInt(3)
   cmb_CodPro.AddItem "CREDITO MIHOGAR"
   cmb_CodPro.ItemData(cmb_CodPro.NewIndex) = CInt(4)
   cmb_CodPro.AddItem "CREDITO MIVIVIENDA"
   cmb_CodPro.ItemData(cmb_CodPro.NewIndex) = CInt(7)
   cmb_CodPro.AddItem "CREDITO MICASA MAS"
   cmb_CodPro.ItemData(cmb_CodPro.NewIndex) = CInt(19)
   cmb_CodPro.AddItem "CREDITO MIVIVIENDA BBP"
   cmb_CodPro.ItemData(cmb_CodPro.NewIndex) = CInt(21)
'   cmb_CodPro.AddItem "BBP - COMPLEMENTO CUOTA INICIAL"
'   cmb_CodPro.ItemData(cmb_CodPro.NewIndex) = CInt(22)
   cmb_CodPro.AddItem "CREDITO TECHO PROPIO"
   cmb_CodPro.ItemData(cmb_CodPro.NewIndex) = CInt(24)
   cmb_CodPro.ListIndex = -1
   
   dir_LisCar.Path = "C:\"
   lbl_NomPro.Caption = Empty
End Sub

Private Sub fs_Limpia()
Dim r_int_PerMes  As Integer
Dim r_int_PerAno  As Integer

   If Month(date) = 1 Then
      r_int_PerMes = 12
      r_int_PerAno = Year(date) - 1
   Else
      r_int_PerMes = Month(date) - 1
      r_int_PerAno = Year(date)
   End If

   ipp_PerAno.Text = Format(r_int_PerAno, "0000")
   fil_LisArc.Pattern = "*.xls"
End Sub

Private Sub fs_CargaCOFIDE_Trimestre(ByVal p_ArcCOFIDE, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, ByVal p_CodPro As Integer, Optional p_BarPro As SSPanel)
Dim r_obj_Excel         As Excel.Application
Dim r_int_FilExc        As Integer
Dim r_str_CodCof        As String
Dim r_str_NomCli        As String
Dim r_str_FecExp        As String
Dim r_dbl_ImpDesTot     As Double
Dim r_dbl_ImpDesTNC     As Double
Dim r_dbl_ImpDesTC      As Double
Dim r_dbl_SaldoTNC      As Double
Dim r_dbl_SaldoTC       As Double
Dim r_dbl_IntTNC        As Double
Dim r_dbl_IntTC         As Double
Dim r_dbl_ComTNC        As Double
Dim r_dbl_ComTC         As Double
Dim r_dbl_CuoPendTNC    As Double
Dim r_dbl_CuoPendTC     As Double
Dim r_str_NumOriTot     As String
'Dim r_str_CodPrd        As String
Dim r_int_Judici        As String
Dim r_int_Filaux        As Integer
Dim r_str_NumCon        As String
Dim r_str_NumCoC        As String
Dim r_str_TIPMON        As String
Dim r_dbl_PriFinTot     As Double
Dim r_dbl_IntFinTot     As Double
Dim r_dbl_ComFinTot     As Double
Dim r_int_CuoPenTot     As Integer

Dim r_dbl_PriFinTNC     As Double
Dim r_dbl_IntFinTNC     As Double
Dim r_dbl_ComFinTNC     As Double
Dim r_int_CuoPenTNC     As Integer
Dim r_str_NumOriTNC     As String

Dim r_dbl_PriFinTC      As Double
Dim r_dbl_IntFinTC      As Double
Dim r_dbl_ComFinTC      As Double
Dim r_int_CuoPenTC      As Integer
Dim r_str_NumOriTC      As String
Dim r_lng_Contad        As Long

Dim r_lng_NumReg        As Long
Dim r_lng_TotReg        As Long

    'Abriendo Archivo COFIDE
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Open FileName:=p_ArcCOFIDE
   
   r_int_FilExc = 2
   r_int_Filaux = 0
   r_int_Judici = 0
   r_lng_NumReg = 0
   r_lng_TotReg = 0
   p_BarPro.FloodPercent = 0
   ReDim r_str_Saldos(0)
   
   r_obj_Excel.Sheets(1).Select
   
   'HOJA PRINCIPAL
   Do While Trim(r_obj_Excel.Cells(r_int_FilExc, 1).Value) <> ""
      r_str_CodCof = Trim(r_obj_Excel.Cells(r_int_FilExc, 3).Value)
      r_str_NomCli = Trim(r_obj_Excel.Cells(r_int_FilExc, 4).Value)
      r_str_NumCon = Trim(r_obj_Excel.Cells(r_int_FilExc, 5).Value)
      r_str_NumCoC = Trim(r_obj_Excel.Cells(r_int_FilExc, 6).Value)
      r_str_FecExp = Format(Trim(r_obj_Excel.Cells(r_int_FilExc, 7).Value), "yyyymmdd")
      r_str_TIPMON = Trim(r_obj_Excel.Cells(r_int_FilExc, 8).Value)
      r_dbl_ImpDesTot = Trim(r_obj_Excel.Cells(r_int_FilExc, 9).Value)
      r_dbl_PriFinTot = Trim(r_obj_Excel.Cells(r_int_FilExc, 10).Value)
      r_dbl_IntFinTot = Trim(r_obj_Excel.Cells(r_int_FilExc, 11).Value)
      r_dbl_ComFinTot = Trim(r_obj_Excel.Cells(r_int_FilExc, 12).Value)
      r_int_CuoPenTot = Trim(r_obj_Excel.Cells(r_int_FilExc, 13).Value)
      r_str_NumOriTot = Trim(r_obj_Excel.Cells(r_int_FilExc, 14).Value)
        
      'Cargar la data en el array
      ReDim Preserve r_str_Saldos(UBound(r_str_Saldos) + 1)
      r_str_Saldos(UBound(r_str_Saldos)).CodCli = r_str_CodCof
      r_str_Saldos(UBound(r_str_Saldos)).NomCli = r_str_NomCli
      r_str_Saldos(UBound(r_str_Saldos)).NroCon = r_str_NumCon
      r_str_Saldos(UBound(r_str_Saldos)).NroCon_C = r_str_NumCoC
      r_str_Saldos(UBound(r_str_Saldos)).FecExp = r_str_FecExp
      r_str_Saldos(UBound(r_str_Saldos)).TipMon = r_str_TIPMON
      r_str_Saldos(UBound(r_str_Saldos)).ImpDes = r_dbl_ImpDesTot
      r_str_Saldos(UBound(r_str_Saldos)).PriFin = r_dbl_PriFinTot
      r_str_Saldos(UBound(r_str_Saldos)).IntFin = r_dbl_IntFinTot
      r_str_Saldos(UBound(r_str_Saldos)).ComFin = r_dbl_ComFinTot
      r_str_Saldos(UBound(r_str_Saldos)).CuoPen = r_int_CuoPenTot
      r_str_Saldos(UBound(r_str_Saldos)).NroOri = r_str_NumOriTot
      
      r_int_FilExc = r_int_FilExc + 1
   Loop
    
   'TNC
   r_obj_Excel.Sheets(2).Select
   r_int_FilExc = 2
   r_str_CodCof = ""
   r_dbl_ImpDesTNC = 0
   r_dbl_PriFinTNC = 0
   r_dbl_IntFinTNC = 0
   r_dbl_ComFinTNC = 0
   r_int_CuoPenTNC = 0
   r_str_NumOriTNC = ""
   r_dbl_ImpDesTC = 0
   r_dbl_PriFinTC = 0
   r_dbl_IntFinTC = 0
   r_dbl_ComFinTC = 0
   r_int_CuoPenTC = 0
   r_str_NumOriTC = ""
    
   Do While Trim(r_obj_Excel.Cells(r_int_FilExc, 1).Value) <> ""
      r_str_CodCof = Trim(r_obj_Excel.Cells(r_int_FilExc, 3).Value)
      r_dbl_ImpDesTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 8).Value)
      r_dbl_PriFinTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 9).Value)
      r_dbl_IntFinTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 10).Value)
      r_dbl_ComFinTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 11).Value)
      r_int_CuoPenTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 12).Value)
      r_str_NumOriTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 13).Value)
      
      For r_lng_Contad = 1 To UBound(r_str_Saldos)
         If r_str_Saldos(r_lng_Contad).CodCli = r_str_CodCof Then
             r_str_Saldos(r_lng_Contad).ImpDes_TNC = r_dbl_ImpDesTNC
             r_str_Saldos(r_lng_Contad).PriFin_TNC = r_dbl_PriFinTNC
             r_str_Saldos(r_lng_Contad).IntFin_TNC = r_dbl_IntFinTNC
             r_str_Saldos(r_lng_Contad).ComFin_TNC = r_dbl_ComFinTNC
             r_str_Saldos(r_lng_Contad).CuoPen_TNC = r_int_CuoPenTNC
             r_str_Saldos(r_lng_Contad).NroOri_TNC = r_str_NumOriTNC
             Exit For
         End If
      Next r_lng_Contad
      
      r_int_FilExc = r_int_FilExc + 1
   Loop
    
   If p_CodPro < 18 And p_CodPro <> 3 Then
      r_obj_Excel.Sheets(3).Select
      r_int_FilExc = 2
      
      r_str_CodCof = ""
      r_dbl_ImpDesTNC = 0
      r_dbl_PriFinTNC = 0
      r_dbl_IntFinTNC = 0
      r_dbl_ComFinTNC = 0
      r_int_CuoPenTNC = 0
      r_str_NumOriTNC = ""
      r_dbl_ImpDesTC = 0
      r_dbl_PriFinTC = 0
      r_dbl_IntFinTC = 0
      r_dbl_ComFinTC = 0
      r_int_CuoPenTC = 0
      r_str_NumOriTC = ""
       
      Do While Trim(r_obj_Excel.Cells(r_int_FilExc, 1).Value) <> ""
         r_str_CodCof = Trim(r_obj_Excel.Cells(r_int_FilExc, 3).Value)
         r_dbl_ImpDesTC = Trim(r_obj_Excel.Cells(r_int_FilExc, 8).Value)
         r_dbl_PriFinTC = Trim(r_obj_Excel.Cells(r_int_FilExc, 9).Value)
         r_dbl_IntFinTC = Trim(r_obj_Excel.Cells(r_int_FilExc, 10).Value)
         r_dbl_ComFinTC = Trim(r_obj_Excel.Cells(r_int_FilExc, 11).Value)
         r_int_CuoPenTC = Trim(r_obj_Excel.Cells(r_int_FilExc, 12).Value)
         r_str_NumOriTC = Trim(r_obj_Excel.Cells(r_int_FilExc, 13).Value)
         
         For r_lng_Contad = 1 To UBound(r_str_Saldos)
            If r_str_Saldos(r_lng_Contad).CodCli = r_str_CodCof Then
                r_str_Saldos(r_lng_Contad).ImpDes_TC = r_dbl_ImpDesTC
                r_str_Saldos(r_lng_Contad).PriFin_TC = r_dbl_PriFinTC
                r_str_Saldos(r_lng_Contad).IntFin_TC = r_dbl_IntFinTC
                r_str_Saldos(r_lng_Contad).ComFin_TC = r_dbl_ComFinTC
                r_str_Saldos(r_lng_Contad).CuoPen_TC = r_int_CuoPenTC
                r_str_Saldos(r_lng_Contad).NroOri_TC = r_str_NumOriTC
                Exit For
            End If
         Next r_lng_Contad
         r_int_FilExc = r_int_FilExc + 1
      Loop
   End If
    
   If p_CodPro = 7 Then
      r_obj_Excel.Sheets(4).Select
      r_int_FilExc = 2
      
      r_str_CodCof = ""
      r_dbl_ImpDesTNC = 0
      r_dbl_PriFinTNC = 0
      r_dbl_IntFinTNC = 0
      r_dbl_ComFinTNC = 0
      r_int_CuoPenTNC = 0
      r_str_NumOriTNC = ""
      r_dbl_ImpDesTC = 0
      r_dbl_PriFinTC = 0
      r_dbl_IntFinTC = 0
      r_dbl_ComFinTC = 0
      r_int_CuoPenTC = 0
      r_str_NumOriTC = ""
        
      Do While Trim(r_obj_Excel.Cells(r_int_FilExc, 1).Value) <> ""
         r_str_CodCof = Trim(r_obj_Excel.Cells(r_int_FilExc, 3).Value)
         r_dbl_ImpDesTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 8).Value)
         r_dbl_PriFinTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 9).Value)
         r_dbl_IntFinTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 10).Value)
         r_dbl_ComFinTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 11).Value)
         r_int_CuoPenTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 12).Value)
         r_str_NumOriTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 13).Value)
         
         For r_lng_Contad = 1 To UBound(r_str_Saldos)
            If r_str_Saldos(r_lng_Contad).CodCli = r_str_CodCof Then
                r_str_Saldos(r_lng_Contad).ImpDes_TNC = r_dbl_ImpDesTNC
                r_str_Saldos(r_lng_Contad).PriFin_TNC = r_dbl_PriFinTNC
                r_str_Saldos(r_lng_Contad).IntFin_TNC = r_dbl_IntFinTNC
                r_str_Saldos(r_lng_Contad).ComFin_TNC = r_dbl_ComFinTNC
                r_str_Saldos(r_lng_Contad).CuoPen_TNC = r_int_CuoPenTNC
                r_str_Saldos(r_lng_Contad).NroOri_TNC = r_str_NumOriTNC
                Exit For
            
            End If
         Next r_lng_Contad
          r_int_FilExc = r_int_FilExc + 1
      Loop
        
      r_obj_Excel.Sheets(5).Select
      r_int_FilExc = 2
      
      r_str_CodCof = ""
      r_dbl_ImpDesTNC = 0
      r_dbl_PriFinTNC = 0
      r_dbl_IntFinTNC = 0
      r_dbl_ComFinTNC = 0
      r_int_CuoPenTNC = 0
      r_str_NumOriTNC = ""
      r_dbl_ImpDesTC = 0
      r_dbl_PriFinTC = 0
      r_dbl_IntFinTC = 0
      r_dbl_ComFinTC = 0
      r_int_CuoPenTC = 0
      r_str_NumOriTC = ""
        
      Do While Trim(r_obj_Excel.Cells(r_int_FilExc, 1).Value) <> ""
         r_str_CodCof = Trim(r_obj_Excel.Cells(r_int_FilExc, 3).Value)
         r_dbl_ImpDesTC = Trim(r_obj_Excel.Cells(r_int_FilExc, 8).Value)
         r_dbl_PriFinTC = Trim(r_obj_Excel.Cells(r_int_FilExc, 9).Value)
         r_dbl_IntFinTC = Trim(r_obj_Excel.Cells(r_int_FilExc, 10).Value)
         r_dbl_ComFinTC = Trim(r_obj_Excel.Cells(r_int_FilExc, 11).Value)
         r_int_CuoPenTC = Trim(r_obj_Excel.Cells(r_int_FilExc, 12).Value)
         r_str_NumOriTC = Trim(r_obj_Excel.Cells(r_int_FilExc, 13).Value)
         
         For r_lng_Contad = 1 To UBound(r_str_Saldos)
            If r_str_Saldos(r_lng_Contad).CodCli = r_str_CodCof Then
                r_str_Saldos(r_lng_Contad).ImpDes_TC = r_dbl_ImpDesTC
                r_str_Saldos(r_lng_Contad).PriFin_TC = r_dbl_PriFinTC
                r_str_Saldos(r_lng_Contad).IntFin_TC = r_dbl_IntFinTC
                r_str_Saldos(r_lng_Contad).ComFin_TC = r_dbl_ComFinTC
                r_str_Saldos(r_lng_Contad).CuoPen_TC = r_int_CuoPenTC
                r_str_Saldos(r_lng_Contad).NroOri_TC = r_str_NumOriTC
                Exit For
            End If
         Next r_lng_Contad
         r_int_FilExc = r_int_FilExc + 1
      Loop
        
      r_obj_Excel.Sheets(6).Select
      r_int_FilExc = 2
      
      r_str_CodCof = ""
      r_dbl_ImpDesTNC = 0
      r_dbl_PriFinTNC = 0
      r_dbl_IntFinTNC = 0
      r_dbl_ComFinTNC = 0
      r_int_CuoPenTNC = 0
      r_str_NumOriTNC = ""
      r_dbl_ImpDesTC = 0
      r_dbl_PriFinTC = 0
      r_dbl_IntFinTC = 0
      r_dbl_ComFinTC = 0
      r_int_CuoPenTC = 0
      r_str_NumOriTC = ""
        
      Do While Trim(r_obj_Excel.Cells(r_int_FilExc, 1).Value) <> ""
          r_str_CodCof = Trim(r_obj_Excel.Cells(r_int_FilExc, 3).Value)
          r_dbl_ImpDesTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 8).Value)
          r_dbl_PriFinTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 9).Value)
          r_dbl_IntFinTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 10).Value)
          r_dbl_ComFinTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 11).Value)
          r_int_CuoPenTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 12).Value)
          r_str_NumOriTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 13).Value)
      
          For r_lng_Contad = 1 To UBound(r_str_Saldos)
              If r_str_Saldos(r_lng_Contad).CodCli = r_str_CodCof Then
                  r_str_Saldos(r_lng_Contad).ImpDes_TNC = r_dbl_ImpDesTNC
                  r_str_Saldos(r_lng_Contad).PriFin_TNC = r_dbl_PriFinTNC
                  r_str_Saldos(r_lng_Contad).IntFin_TNC = r_dbl_IntFinTNC
                  r_str_Saldos(r_lng_Contad).ComFin_TNC = r_dbl_ComFinTNC
                  r_str_Saldos(r_lng_Contad).CuoPen_TNC = r_int_CuoPenTNC
                  r_str_Saldos(r_lng_Contad).NroOri_TNC = r_str_NumOriTNC
                  Exit For
              End If
          Next r_lng_Contad
          r_int_FilExc = r_int_FilExc + 1
      Loop
        
      r_obj_Excel.Sheets(7).Select
      r_int_FilExc = 2
      
      r_str_CodCof = ""
      r_dbl_ImpDesTNC = 0
      r_dbl_PriFinTNC = 0
      r_dbl_IntFinTNC = 0
      r_dbl_ComFinTNC = 0
      r_int_CuoPenTNC = 0
      r_str_NumOriTNC = ""
      r_dbl_ImpDesTC = 0
      r_dbl_PriFinTC = 0
      r_dbl_IntFinTC = 0
      r_dbl_ComFinTC = 0
      r_int_CuoPenTC = 0
      r_str_NumOriTC = ""
        
      Do While Trim(r_obj_Excel.Cells(r_int_FilExc, 1).Value) <> ""
          r_str_CodCof = Trim(r_obj_Excel.Cells(r_int_FilExc, 3).Value)
          r_dbl_ImpDesTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 8).Value)
          r_dbl_PriFinTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 9).Value)
          r_dbl_IntFinTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 10).Value)
          r_dbl_ComFinTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 11).Value)
          r_int_CuoPenTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 12).Value)
          r_str_NumOriTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 13).Value)
      
          For r_lng_Contad = 1 To UBound(r_str_Saldos)
              If r_str_Saldos(r_lng_Contad).CodCli = r_str_CodCof Then
                  r_str_Saldos(r_lng_Contad).ImpDes_TNC = r_dbl_ImpDesTNC
                  r_str_Saldos(r_lng_Contad).PriFin_TNC = r_dbl_PriFinTNC
                  r_str_Saldos(r_lng_Contad).IntFin_TNC = r_dbl_IntFinTNC
                  r_str_Saldos(r_lng_Contad).ComFin_TNC = r_dbl_ComFinTNC
                  r_str_Saldos(r_lng_Contad).CuoPen_TNC = r_int_CuoPenTNC
                  r_str_Saldos(r_lng_Contad).NroOri_TNC = r_str_NumOriTNC
                  Exit For
              End If
          Next r_lng_Contad
          r_int_FilExc = r_int_FilExc + 1
      Loop
      
      r_obj_Excel.Sheets(8).Select
      r_int_FilExc = 2

      r_str_CodCof = ""
      r_dbl_ImpDesTNC = 0
      r_dbl_PriFinTNC = 0
      r_dbl_IntFinTNC = 0
      r_dbl_ComFinTNC = 0
      r_int_CuoPenTNC = 0
      r_str_NumOriTNC = ""
      r_dbl_ImpDesTC = 0
      r_dbl_PriFinTC = 0
      r_dbl_IntFinTC = 0
      r_dbl_ComFinTC = 0
      r_int_CuoPenTC = 0
      r_str_NumOriTC = ""
        
      Do While Trim(r_obj_Excel.Cells(r_int_FilExc, 1).Value) <> ""
          r_str_CodCof = Trim(r_obj_Excel.Cells(r_int_FilExc, 3).Value)
          r_dbl_ImpDesTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 8).Value)
          r_dbl_PriFinTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 9).Value)
          r_dbl_IntFinTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 10).Value)
          r_dbl_ComFinTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 11).Value)
          r_int_CuoPenTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 12).Value)
          r_str_NumOriTNC = Trim(r_obj_Excel.Cells(r_int_FilExc, 13).Value)
      
          For r_lng_Contad = 1 To UBound(r_str_Saldos)
              If r_str_Saldos(r_lng_Contad).CodCli = r_str_CodCof Then
                  r_str_Saldos(r_lng_Contad).ImpDes_TNC = r_dbl_ImpDesTNC
                  r_str_Saldos(r_lng_Contad).PriFin_TNC = r_dbl_PriFinTNC
                  r_str_Saldos(r_lng_Contad).IntFin_TNC = r_dbl_IntFinTNC
                  r_str_Saldos(r_lng_Contad).ComFin_TNC = r_dbl_ComFinTNC
                  r_str_Saldos(r_lng_Contad).CuoPen_TNC = r_int_CuoPenTNC
                  r_str_Saldos(r_lng_Contad).NroOri_TNC = r_str_NumOriTNC
                  Exit For
              End If
          Next r_lng_Contad
          r_int_FilExc = r_int_FilExc + 1
      Loop
   End If
   
   r_lng_TotReg = UBound(r_str_Saldos)
   
   For r_lng_Contad = 1 To UBound(r_str_Saldos)
   
      r_lng_NumReg = r_lng_Contad
      'Inserta
      g_str_Parame = ""
      g_str_Parame = "INSERT INTO CRE_ARCCOF ("
      g_str_Parame = g_str_Parame & "ARCCOF_PERMES, "
      g_str_Parame = g_str_Parame & "ARCCOF_PERANO, "
      g_str_Parame = g_str_Parame & "ARCCOF_CODCLI, "
      g_str_Parame = g_str_Parame & "ARCCOF_TIPARC, "
      g_str_Parame = g_str_Parame & "ARCCOF_NUMORI, "
      g_str_Parame = g_str_Parame & "ARCCOF_NOMCLI, "
      g_str_Parame = g_str_Parame & "ARCCOF_FECEXP, "
      g_str_Parame = g_str_Parame & "ARCCOF_IMPDES_TOT, "
      g_str_Parame = g_str_Parame & "ARCCOF_IMPDES_TNC, "
      g_str_Parame = g_str_Parame & "ARCCOF_IMPDES_TC, "
      g_str_Parame = g_str_Parame & "ARCCOF_SALDTN, "
      g_str_Parame = g_str_Parame & "ARCCOF_SALDTC, "
      g_str_Parame = g_str_Parame & "ARCCOF_INTETN, "
      g_str_Parame = g_str_Parame & "ARCCOF_INTETC, "
      g_str_Parame = g_str_Parame & "ARCCOF_COMITN, "
      g_str_Parame = g_str_Parame & "ARCCOF_COMITC, "
      g_str_Parame = g_str_Parame & "ARCCOF_CUOPEN_TNC, "
      g_str_Parame = g_str_Parame & "ARCCOF_CUOPEN_TC, "
      g_str_Parame = g_str_Parame & "SEGUSUCRE, "
      g_str_Parame = g_str_Parame & "SEGFECCRE, "
      g_str_Parame = g_str_Parame & "SEGHORCRE, "
      g_str_Parame = g_str_Parame & "SEGPLTCRE, "
      g_str_Parame = g_str_Parame & "SEGTERCRE, "
      g_str_Parame = g_str_Parame & "SEGSUCCRE ) "
      g_str_Parame = g_str_Parame & "VALUES ( "
      g_str_Parame = g_str_Parame & p_PerMes & " , "
      g_str_Parame = g_str_Parame & p_PerAno & ", "
      g_str_Parame = g_str_Parame & "'" & r_str_Saldos(r_lng_Contad).CodCli & "', "
      g_str_Parame = g_str_Parame & p_CodPro & ", "
      g_str_Parame = g_str_Parame & "'" & r_str_Saldos(r_lng_Contad).NroOri & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_Saldos(r_lng_Contad).NomCli & "', "
      g_str_Parame = g_str_Parame & r_str_Saldos(r_lng_Contad).FecExp & ", "
      g_str_Parame = g_str_Parame & r_str_Saldos(r_lng_Contad).ImpDes & ", "
      g_str_Parame = g_str_Parame & IIf(r_str_Saldos(r_lng_Contad).ImpDes_TNC = "", 0, r_str_Saldos(r_lng_Contad).ImpDes_TNC) & ", "
      g_str_Parame = g_str_Parame & IIf(r_str_Saldos(r_lng_Contad).ImpDes_TC = "", 0, r_str_Saldos(r_lng_Contad).ImpDes_TC) & ", "
      g_str_Parame = g_str_Parame & IIf(r_str_Saldos(r_lng_Contad).PriFin_TNC = "", 0, r_str_Saldos(r_lng_Contad).PriFin_TNC) & ", "
      g_str_Parame = g_str_Parame & IIf(r_str_Saldos(r_lng_Contad).PriFin_TC = "", 0, r_str_Saldos(r_lng_Contad).PriFin_TC) & ", "
      g_str_Parame = g_str_Parame & IIf(r_str_Saldos(r_lng_Contad).IntFin_TNC = "", 0, r_str_Saldos(r_lng_Contad).IntFin_TNC) & ", "
      g_str_Parame = g_str_Parame & IIf(r_str_Saldos(r_lng_Contad).IntFin_TC = "", 0, r_str_Saldos(r_lng_Contad).IntFin_TC) & ", "
      g_str_Parame = g_str_Parame & IIf(r_str_Saldos(r_lng_Contad).ComFin_TNC = "", 0, r_str_Saldos(r_lng_Contad).ComFin_TNC) & ", "
      g_str_Parame = g_str_Parame & IIf(r_str_Saldos(r_lng_Contad).ComFin_TC = "", 0, r_str_Saldos(r_lng_Contad).ComFin_TC) & ", "
      g_str_Parame = g_str_Parame & IIf(r_str_Saldos(r_lng_Contad).CuoPen_TNC = "", 0, r_str_Saldos(r_lng_Contad).CuoPen_TNC) & ", "
      g_str_Parame = g_str_Parame & IIf(r_str_Saldos(r_lng_Contad).CuoPen_TC = "", 0, r_str_Saldos(r_lng_Contad).CuoPen_TC) & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "' ,"
      g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & ", "
      g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & ", "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "' ,"
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
          Exit Sub
      End If
      DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
      
      p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
   Next r_lng_Contad
  
   r_obj_Excel.Workbooks.Close
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_CargaCOFIDE_Mensual(ByVal p_ArcCOFIDE, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, Optional p_BarPro As SSPanel)
Dim r_obj_Excel      As Excel.Application
Dim r_int_FilExc     As Integer
Dim r_int_FilTot     As Long
Dim r_int_IDCIPR     As Integer
Dim r_str_NOMPRO     As String
Dim r_str_CodCof     As String
Dim r_str_NomCli     As String
Dim r_str_NUMCTR     As String
Dim r_str_NUMALT     As String
Dim r_str_TIPMON     As String
Dim r_dbl_EXPINI     As Double
Dim r_dbl_Princi     As Double
Dim r_dbl_IMPINT     As Double
Dim r_dbl_IMTASA     As Double
Dim r_dbl_COMSIN     As Double
Dim r_dbl_ImpTot     As Double
Dim r_dbl_EXPFIN     As Double
Dim r_str_BUENPG     As String
Dim r_str_MALPAG     As String
Dim r_lng_NumReg     As Long
Dim r_lng_TotReg     As Long
       
   'Abriendo Archivo COFIDE
   r_lng_NumReg = 0
   r_lng_TotReg = 0
   p_BarPro.FloodPercent = 0
      
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Open FileName:=p_ArcCOFIDE
   r_int_FilExc = 1
   
   With r_obj_Excel.Sheets(1)
        r_int_FilTot = CStr(.Cells(.Rows.Count, 1).End(xlUp).Row)
        r_lng_TotReg = r_int_FilTot
        
        Do While r_int_FilTot <> r_int_FilExc
           r_lng_NumReg = r_int_FilExc
           
           If (Len(Trim(Cells(r_int_FilExc, 1).Value)) >= 8 And Len(Trim(Cells(r_int_FilExc, 3).Value)) >= 3 And Len(Trim(Cells(r_int_FilExc, 5).Value)) >= 8) Then
               If (IsNumeric(Trim(Cells(r_int_FilExc, 1).Value)) = True And IsNumeric(Trim(Cells(r_int_FilExc, 3).Value)) = True And IsNumeric(Trim(Cells(r_int_FilExc, 5).Value)) = True) Then
                   r_int_IDCIPR = 0:  r_str_NOMPRO = "": r_str_CodCof = ""
                   r_str_NomCli = "": r_str_NUMCTR = "": r_str_NUMALT = "": r_str_TIPMON = ""
                   r_dbl_EXPINI = 0:  r_dbl_Princi = 0:  r_dbl_IMPINT = 0
                   r_dbl_IMTASA = 0:  r_dbl_COMSIN = 0:  r_dbl_ImpTot = 0
                   r_dbl_EXPFIN = 0:  r_str_BUENPG = 0:  r_str_MALPAG = 0
                     
                   r_int_IDCIPR = Trim(Cells(r_int_FilExc, 3).Value)
                   r_str_NOMPRO = Trim(Cells(r_int_FilExc, 4).Value)
                   r_str_CodCof = Trim(Cells(r_int_FilExc, 5).Value)
                   r_str_NomCli = Trim(Cells(r_int_FilExc, 6).Value)
                   r_str_NUMCTR = Trim(Cells(r_int_FilExc, 7).Value)
                   r_str_NUMALT = Trim(Cells(r_int_FilExc, 8).Value)
                   r_str_TIPMON = Trim(Cells(r_int_FilExc, 9).Value)
                   r_dbl_EXPINI = Trim(Cells(r_int_FilExc, 10).Value)
                   r_dbl_Princi = Trim(Cells(r_int_FilExc, 11).Value)
                   r_dbl_IMPINT = Trim(Cells(r_int_FilExc, 12).Value)
                   r_dbl_IMTASA = Trim(Cells(r_int_FilExc, 13).Value)
                   r_dbl_COMSIN = Trim(Cells(r_int_FilExc, 14).Value)
                   r_dbl_ImpTot = Trim(Cells(r_int_FilExc, 15).Value)
                   r_dbl_EXPFIN = Trim(Cells(r_int_FilExc, 16).Value)
                   r_str_BUENPG = Trim(Cells(r_int_FilExc, 17).Value)
                   r_str_MALPAG = Trim(Cells(r_int_FilExc, 18).Value)
               
                   g_str_Parame = ""
                   g_str_Parame = "INSERT INTO CRE_ARCMEN ("
                   g_str_Parame = g_str_Parame & "ARCMEN_PERANO, "
                   g_str_Parame = g_str_Parame & "ARCMEN_PERMES, "
                   g_str_Parame = g_str_Parame & "ARCMEN_NUMCTR, "
                   g_str_Parame = g_str_Parame & "ARCMEN_IDCIPR, "
                   g_str_Parame = g_str_Parame & "ARCMEN_NOMPRO, "
                   g_str_Parame = g_str_Parame & "ARCMEN_CODCOF, "
                   g_str_Parame = g_str_Parame & "ARCMEN_NOMCLI, "
                   g_str_Parame = g_str_Parame & "ARCMEN_NUMALT, "
                   g_str_Parame = g_str_Parame & "ARCMEN_TIPMON, "
                   g_str_Parame = g_str_Parame & "ARCMEN_EXPINI, "
                   g_str_Parame = g_str_Parame & "ARCMEN_PRINCI, "
                   g_str_Parame = g_str_Parame & "ARCMEN_IMPINT, "
                   g_str_Parame = g_str_Parame & "ARCMEN_IMTASA, "
                   g_str_Parame = g_str_Parame & "ARCMEN_COMSIN, "
                   g_str_Parame = g_str_Parame & "ARCMEN_IMPTOT, "
                   g_str_Parame = g_str_Parame & "ARCMEN_EXPFIN, "
                   g_str_Parame = g_str_Parame & "ARCMEN_BUENPG, "
                   g_str_Parame = g_str_Parame & "ARCMEN_MALPAG, "
                   g_str_Parame = g_str_Parame & "SEGUSUCRE, "
                   g_str_Parame = g_str_Parame & "SEGFECCRE, "
                   g_str_Parame = g_str_Parame & "SEGHORCRE, "
                   g_str_Parame = g_str_Parame & "SEGPLTCRE, "
                   g_str_Parame = g_str_Parame & "SEGTERCRE, "
                   g_str_Parame = g_str_Parame & "SEGSUCCRE) "
                   g_str_Parame = g_str_Parame & "VALUES ( "
                   g_str_Parame = g_str_Parame & p_PerAno & " , "
                   g_str_Parame = g_str_Parame & p_PerMes & " , "
                   g_str_Parame = g_str_Parame & "'" & r_str_NUMCTR & "', "
                   g_str_Parame = g_str_Parame & "'" & r_int_IDCIPR & "', "
                   g_str_Parame = g_str_Parame & "'" & r_str_NOMPRO & "', "
                   g_str_Parame = g_str_Parame & "'" & r_str_CodCof & "', "
                   g_str_Parame = g_str_Parame & "'" & r_str_NomCli & "', "
                   g_str_Parame = g_str_Parame & "'" & r_str_NUMALT & "', "
                   g_str_Parame = g_str_Parame & "'" & r_str_TIPMON & "', "
                   g_str_Parame = g_str_Parame & r_dbl_EXPINI & ", "
                   g_str_Parame = g_str_Parame & r_dbl_Princi & ", "
                   g_str_Parame = g_str_Parame & r_dbl_IMPINT & ", "
                   g_str_Parame = g_str_Parame & r_dbl_IMTASA & ", "
                   g_str_Parame = g_str_Parame & r_dbl_COMSIN & ", "
                   g_str_Parame = g_str_Parame & r_dbl_ImpTot & ", "
                   g_str_Parame = g_str_Parame & r_dbl_EXPFIN & ", "
                   g_str_Parame = g_str_Parame & "'" & r_str_BUENPG & "', "
                   g_str_Parame = g_str_Parame & "'" & r_str_MALPAG & "', "
               
                   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "' ,"
                   g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & ", "
                   g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & ", "
                   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
                   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "' ,"
                   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
               
                   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
                      Exit Sub
                   End If
                   
               End If
           End If
           
           p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
           r_int_FilExc = r_int_FilExc + 1
        Loop
   End With
   
   r_obj_Excel.Workbooks.Close
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_CargaCOFIDE_Calendario(ByVal p_ArcCOFIDE As String)
Dim r_obj_Excel         As Excel.Application
Dim r_int_FilExc        As Integer
Dim r_int_FilTot        As Long
Dim r_lng_Contad        As Long
Dim r_str_NumOpe        As String
Dim r_int_TipCro        As Integer
Dim r_dbl_MtoPre        As Double
Dim r_dbl_MtoPre_TNC    As Double
Dim r_dbl_MtoPre_TC     As Double
Dim r_int_NumCuo_TNC    As Integer
Dim r_int_NumCuo_TC     As Integer
Dim r_int_NCuTNC        As Integer
Dim r_int_NCuoTC        As Integer
Dim r_int_CuoDob        As Integer
Dim r_dbl_MtoCDo        As Double
Dim r_int_PerGra        As Integer
Dim r_dbl_MtoPGr        As Double
Dim r_int_DifSal_TNC    As Integer
Dim r_int_DifSal_TC     As Integer

Dim r_int_NumCuo        As Integer
Dim r_str_FecVct        As String
Dim r_dbl_Capita        As Double
Dim r_dbl_Intere        As Double
Dim r_dbl_ComCof        As Double
Dim r_dbl_MtoCuo        As Double
Dim r_dbl_SalCap        As Double

Dim r_bol_flgMPr        As Boolean
Dim r_bol_flgSal        As Boolean
Dim r_bol_flgNCu        As Boolean
Dim r_bol_flgCDb        As Boolean
Dim r_bol_flgPGr        As Boolean
Dim r_bol_FlgIng        As Boolean

   'Abriendo Archivo COFIDE
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Open FileName:=p_ArcCOFIDE
   r_int_FilExc = 1
   
   ReDim r_str_CalDes(0)
   r_dbl_MtoPre = 0
   r_dbl_MtoPre_TNC = 0
   r_dbl_MtoPre_TC = 0
   r_int_NumCuo_TNC = 0
   r_int_NumCuo_TC = 0
   r_int_CuoDob = 0
   r_dbl_MtoCDo = 0
   r_int_PerGra = 0
   r_dbl_MtoPGr = 0
   r_int_DifSal_TNC = 0
   r_int_DifSal_TC = 0
   r_int_FilTot = 0

   With r_obj_Excel.Sheets(1)
      r_int_FilTot = CStr(.Cells(.Rows.Count, 1).End(xlUp).Row)
        
      Do While r_int_FilTot >= r_int_FilExc
                 
         If Len(Trim(.Cells(r_int_FilExc, 1).Value)) >= 13 Then
            If r_int_FilExc = 2 Then
               g_str_Parame = ""
               g_str_Parame = g_str_Parame & " SELECT HIPMAE_NUMOPE AS OPERACION, HIPMAE_MTOPRE AS MTOPRE, HIPMAE_IMPNCO AS MTOPRE_TNC, HIPMAE_IMPCON AS MTOPRE_TC, "
               g_str_Parame = g_str_Parame & "        HIPMAE_NUMCUO AS NUMCUO, HIPMAE_NUMCUO_CON AS NUMCUO_CON, HIPMAE_CUOANO AS CUODOB, HIPMAE_PERGRA AS PERGRA "
               g_str_Parame = g_str_Parame & "   FROM CRE_HIPMAE "
               g_str_Parame = g_str_Parame & "  WHERE HIPMAE_OPEMVI = " & Trim(.Cells(r_int_FilExc, 1).Value) & "' "
               g_str_Parame = g_str_Parame & "    AND HIPMAE_SITUAC = 2 "
               
               If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
                  Exit Sub
               End If
               
               If g_rst_Princi.BOF And g_rst_Princi.EOF Then
                  g_rst_Princi.Close
                  Set g_rst_Princi = Nothing
                  
                  ReDim Preserve r_str_DesErr(UBound(r_str_DesErr) + 1)
                  r_str_DesErr(UBound(r_str_DesErr)).NomArc = Trim(Mid(p_ArcCOFIDE, InStrRev(p_ArcCOFIDE, "\") + 1))
                  r_str_DesErr(UBound(r_str_DesErr)).NumOpe = Trim(r_str_NumOpe)
                  r_str_DesErr(UBound(r_str_DesErr)).OpeMvi = Trim(.Cells(r_int_FilExc, 1).Value)
                  r_str_DesErr(UBound(r_str_DesErr)).MtoErr = "No se encontró la Operación."
                  Exit Sub
               End If
            
               g_rst_Princi.MoveFirst
               r_str_NumOpe = g_rst_Princi!OPERACION
               r_dbl_MtoPre = g_rst_Princi!MTOPRE
               r_dbl_MtoPre_TNC = g_rst_Princi!MTOPRE_TNC
               r_dbl_MtoPre_TC = g_rst_Princi!MTOPRE_TC
               r_int_NumCuo_TNC = g_rst_Princi!NUMCUO
               r_int_NumCuo_TC = g_rst_Princi!NUMCUO_CON
               r_int_CuoDob = g_rst_Princi!CUODOB              '1- NO TIENE CUOTAS DOBLE, 2- JULIO, 3- DICIEMBRE, 4- JULIO/DICIEMBRE
               r_int_PerGra = g_rst_Princi!PERGRA
            End If
            
            ReDim Preserve r_str_CalDes(UBound(r_str_CalDes) + 1)
            r_str_CalDes(UBound(r_str_CalDes)).NumOpe = r_str_NumOpe
            r_str_CalDes(UBound(r_str_CalDes)).OpeMvi = Trim(.Cells(r_int_FilExc, 1).Value)
            r_str_CalDes(UBound(r_str_CalDes)).Secuen = Trim(.Cells(r_int_FilExc, 2).Value)
            r_str_CalDes(UBound(r_str_CalDes)).NroCuo = Trim(.Cells(r_int_FilExc, 3).Value)
            r_str_CalDes(UBound(r_str_CalDes)).FecVct = Trim(.Cells(r_int_FilExc, 4).Value)
            r_str_CalDes(UBound(r_str_CalDes)).NroDia = Trim(.Cells(r_int_FilExc, 5).Value)
            r_str_CalDes(UBound(r_str_CalDes)).Moneda = Trim(.Cells(r_int_FilExc, 6).Value)
            If Trim(.Cells(r_int_FilExc, 2).Value) = "" Or Trim(.Cells(r_int_FilExc, 2).Value) = 2 Then
               r_str_CalDes(UBound(r_str_CalDes)).TipCro = 3
            Else
               r_str_CalDes(UBound(r_str_CalDes)).TipCro = 4
            End If
            r_str_CalDes(UBound(r_str_CalDes)).Capita = CDbl(Trim(.Cells(r_int_FilExc, 7).Value))
            r_str_CalDes(UBound(r_str_CalDes)).Intere = CDbl(Trim(.Cells(r_int_FilExc, 8).Value))
            r_str_CalDes(UBound(r_str_CalDes)).ComCof = CDbl(Trim(.Cells(r_int_FilExc, 9).Value))
            r_str_CalDes(UBound(r_str_CalDes)).MtoCuo = CDbl(Trim(.Cells(r_int_FilExc, 10).Value))
            r_str_CalDes(UBound(r_str_CalDes)).SalCap = CDbl(Trim(.Cells(r_int_FilExc, 11).Value))
         End If
         r_int_FilExc = r_int_FilExc + 1
      Loop
   End With
   
   r_obj_Excel.Workbooks.Close
   Set r_obj_Excel = Nothing
   
   For r_lng_Contad = 1 To UBound(r_str_CalDes)
      
      'TNC
      If r_str_CalDes(r_lng_Contad).TipCro = 3 Then
         If r_lng_Contad = 1 Then
            'Obtiene Monto del Préstamo TNC
             r_dbl_MtoPre_TNC = CDbl(r_str_CalDes(r_lng_Contad).Capita) + CDbl(r_str_CalDes(r_lng_Contad).SalCap)
         End If
         
         If r_lng_Contad <> 2 And Mid(r_str_CalDes(r_lng_Contad).FecVct, 4, 2) <> "07" And Mid(r_str_CalDes(r_lng_Contad).FecVct, 4, 2) <> "12" Then
            r_dbl_MtoCuo = CDbl(r_str_CalDes(r_lng_Contad).MtoCuo)
         'Else 'If r_lng_Contad = 3 Then
          '  r_dbl_MtoCuo = CDbl(r_str_CalDes(r_lng_Contad).MtoCuo)
         End If
         
         'Compara si los saldos son correcto
         If (IsNumeric(r_str_CalDes(r_lng_Contad).Capita) = True And IsNumeric(r_str_CalDes(r_lng_Contad).SalCap) = True) Then
            If r_lng_Contad = 1 Then
               r_int_DifSal_TNC = r_int_DifSal_TNC + (r_dbl_MtoPre_TNC - CDbl(r_str_CalDes(r_lng_Contad).Capita) - CDbl(r_str_CalDes(r_lng_Contad).SalCap))
            Else
               r_int_DifSal_TNC = r_int_DifSal_TNC + (CDbl(r_str_CalDes(r_lng_Contad - 1).SalCap) - CDbl(r_str_CalDes(r_lng_Contad).Capita) - CDbl(r_str_CalDes(r_lng_Contad).SalCap))
            End If
         End If
         
         'Compara Periodo de Gracia
         If r_lng_Contad <= r_int_PerGra Then
            r_dbl_MtoPGr = r_dbl_MtoPGr + r_str_CalDes(r_lng_Contad).Capita + r_str_CalDes(r_lng_Contad).Intere
         End If
         
         'Compara Cuotas Dobles
         If r_int_CuoDob = 1 Then                                                'NINGUNA CUOTA DOBLE
            If Mid(r_str_CalDes(r_lng_Contad).FecVct, 4, 2) = "07" Or Mid(r_str_CalDes(r_lng_Contad).FecVct, 4, 2) = "12" Then
               r_dbl_MtoCDo = r_dbl_MtoCuo
               If CDbl(r_str_CalDes(r_lng_Contad).MtoCuo) = r_dbl_MtoCDo Then
                  r_bol_flgCDb = True
               Else
                  If r_lng_Contad = UBound(r_str_CalDes) Then                    'CUANDO ES ÚLTIMA CUOTA Y ES CUOTA DOBLE
                     r_bol_flgCDb = True
                  Else
                     r_bol_flgCDb = False
                  End If
               End If
            End If
         ElseIf r_int_CuoDob = 2 Then                                            'JULIO
            If Mid(r_str_CalDes(r_lng_Contad).FecVct, 4, 2) = "07" Then
               r_dbl_MtoCDo = 2 * r_dbl_MtoCuo
               If CDbl(r_str_CalDes(r_lng_Contad).MtoCuo) = r_dbl_MtoCDo Then
                  r_bol_flgCDb = True
               Else
                  If r_lng_Contad = UBound(r_str_CalDes) Then                    'CUANDO ES ÚLTIMA CUOTA Y ES CUOTA DOBLE
                     r_bol_flgCDb = True
                  Else
                     r_bol_flgCDb = False
                  End If
               End If
            End If
         ElseIf r_int_CuoDob = 3 Then                                            'DICIEMBRE
            If Mid(r_str_CalDes(r_lng_Contad).FecVct, 4, 2) = "12" Then
               r_dbl_MtoCDo = 2 * r_dbl_MtoCuo
               If CDbl(r_str_CalDes(r_lng_Contad).MtoCuo) = r_dbl_MtoCDo Then
                  r_bol_flgCDb = True
               Else
                  If r_lng_Contad = UBound(r_str_CalDes) Then                    'CUANDO ES ÚLTIMA CUOTA Y ES CUOTA DOBLE
                     r_bol_flgCDb = True
                  Else
                     r_bol_flgCDb = False
                  End If
               End If
            End If
         ElseIf r_int_CuoDob = 4 Then                                            'JULIO / DICIEMBRE
            If Mid(r_str_CalDes(r_lng_Contad).FecVct, 4, 2) = "07" Or Mid(r_str_CalDes(r_lng_Contad).FecVct, 4, 2) = "12" Then
               r_dbl_MtoCDo = 2 * r_dbl_MtoCuo
               If CDbl(r_str_CalDes(r_lng_Contad).MtoCuo) = r_dbl_MtoCDo Then
                  r_bol_flgCDb = True
               Else
                  If r_lng_Contad = UBound(r_str_CalDes) Then                    'CUANDO ES ÚLTIMA CUOTA Y ES CUOTA DOBLE
                     r_bol_flgCDb = True
                  Else
                     r_bol_flgCDb = False
                  End If
               End If
            End If
         End If
         
         r_int_NCuTNC = r_int_NCuTNC + 1
      
      'TC
      Else
         If r_lng_Contad = 1 Then
            'Obtiene Monto del Préstamo TC
             r_dbl_MtoPre_TC = CDbl(r_str_CalDes(r_lng_Contad).Capita) + CDbl(r_str_CalDes(r_lng_Contad).SalCap)
         End If
          
          'Compara si los saldos son correcto
         If (IsNumeric(r_str_CalDes(r_lng_Contad).Capita) = True And IsNumeric(r_str_CalDes(r_lng_Contad).SalCap) = True) Then
            If r_lng_Contad = r_int_NumCuo_TNC Then
               r_int_DifSal_TC = r_int_DifSal_TC + (r_dbl_MtoPre_TC - CDbl(r_str_CalDes(r_lng_Contad).Capita) - CDbl(r_str_CalDes(r_lng_Contad).SalCap))
            Else
               r_int_DifSal_TC = r_int_DifSal_TC + (CDbl(r_str_CalDes(r_lng_Contad - 1).SalCap) - CDbl(r_str_CalDes(r_lng_Contad).Capita) - CDbl(r_str_CalDes(r_lng_Contad).SalCap))
            End If
         End If
         
         r_int_NCuoTC = r_int_NCuoTC + 1
      End If
   Next r_lng_Contad
     
   'Compara Monto del Préstamo
   If r_dbl_MtoPre <> 0 Then
      If CDbl(CStr(CDbl(r_dbl_MtoPre_TNC) + CDbl(r_dbl_MtoPre_TC))) = CDbl(CStr(r_dbl_MtoPre)) Then
         r_bol_flgMPr = True
      Else
         r_bol_flgMPr = False
      End If
   Else
      r_bol_flgMPr = False
   End If
   
   'Compara Saldos
   If CInt(r_int_DifSal_TNC) + CInt(r_int_DifSal_TC) = 0 Then
      r_bol_flgSal = True
   Else
      r_bol_flgSal = False
   End If
   
   'Compara Periodo de Gracia
   If r_dbl_MtoPGr = 0 Then
      r_bol_flgPGr = True
   Else
      r_bol_flgPGr = False
   End If
   
   'Compara Número de Cuotas
   If CInt(r_int_NumCuo_TNC) + CInt(r_int_NumCuo_TC) <> 0 Then
      If CInt(r_int_NCuTNC) + CInt(r_int_NCuoTC) = CInt(r_int_NumCuo_TNC) + CInt(r_int_NumCuo_TC) Then
         r_bol_flgNCu = True
      Else
         r_bol_flgNCu = False
      End If
   Else
      r_bol_flgNCu = False
   End If
      
   If r_bol_flgMPr = True And r_bol_flgSal = True And r_bol_flgPGr = True And r_bol_flgNCu = True And r_bol_flgCDb = True Then
      
      'elimina cronograma de la BD
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "DELETE FROM CRE_HIPCUO "
      g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & r_str_NumOpe & "' "
      g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 3 "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 2) Then
         ReDim Preserve r_str_DesErr(UBound(r_str_DesErr) + 1)
         r_str_DesErr(UBound(r_str_DesErr)).NomArc = Trim(Mid(p_ArcCOFIDE, InStrRev(p_ArcCOFIDE, "\") + 1))
         r_str_DesErr(UBound(r_str_DesErr)).NumOpe = Trim(r_str_NumOpe)
         r_str_DesErr(UBound(r_str_DesErr)).OpeMvi = r_str_CalDes(r_lng_Contad).OpeMvi
         r_str_DesErr(UBound(r_str_DesErr)).MtoErr = "Error al eliminar las cuotas del cronograma FMV TNC."
      End If
         
      For r_lng_Contad = 1 To UBound(r_str_CalDes)
         'carga variables e inserta cuota
         r_int_TipCro = CInt(r_str_CalDes(r_lng_Contad).TipCro)
         r_int_NumCuo = CInt(r_str_CalDes(r_lng_Contad).NroCuo)
         r_str_FecVct = r_str_CalDes(r_lng_Contad).FecVct
         r_dbl_Capita = CDbl(r_str_CalDes(r_lng_Contad).Capita)
         r_dbl_Intere = CDbl(r_str_CalDes(r_lng_Contad).Intere)
         r_dbl_ComCof = CDbl(r_str_CalDes(r_lng_Contad).ComCof)
         r_dbl_MtoCuo = CDbl(r_str_CalDes(r_lng_Contad).MtoCuo)
         r_dbl_SalCap = CDbl(r_str_CalDes(r_lng_Contad).SalCap)
         
         If Not ff_Inserta_HipCuo(r_str_NumOpe, r_int_TipCro, r_int_NumCuo, r_str_FecVct, r_dbl_Capita, r_dbl_Intere, 0, 0, 0, r_dbl_SalCap, 0, 0, r_dbl_ComCof) Then
            ReDim Preserve r_str_DesErr(UBound(r_str_DesErr) + 1)
            r_str_DesErr(UBound(r_str_DesErr)).NomArc = Trim(Mid(p_ArcCOFIDE, InStrRev(p_ArcCOFIDE, "\") + 1))
            r_str_DesErr(UBound(r_str_DesErr)).NumOpe = Trim(r_str_NumOpe)
            r_str_DesErr(UBound(r_str_DesErr)).OpeMvi = r_str_CalDes(r_lng_Contad).OpeMvi
            r_str_DesErr(UBound(r_str_DesErr)).MtoErr = "No se pudo completar el procedimiento USP_CRE_HIPCUO_CREA. Nro. Cuota " & r_int_NumCuo
            Exit For
         Else
            r_bol_FlgIng = True
         End If
      Next r_lng_Contad
   
   Else
      If r_bol_flgMPr = False Then
         ReDim Preserve r_str_DesErr(UBound(r_str_DesErr) + 1)
         r_str_DesErr(UBound(r_str_DesErr)).NomArc = Trim(Mid(p_ArcCOFIDE, InStrRev(p_ArcCOFIDE, "\") + 1))
         r_str_DesErr(UBound(r_str_DesErr)).NumOpe = Trim(r_str_NumOpe)
         r_str_DesErr(UBound(r_str_DesErr)).OpeMvi = r_str_CalDes(r_lng_Contad - 1).OpeMvi
         r_str_DesErr(UBound(r_str_DesErr)).MtoErr = "El Monto del Préstamo no coincide."
      End If
      If r_bol_flgSal = False Then
         ReDim Preserve r_str_DesErr(UBound(r_str_DesErr) + 1)
         r_str_DesErr(UBound(r_str_DesErr)).NomArc = Trim(Mid(p_ArcCOFIDE, InStrRev(p_ArcCOFIDE, "\") + 1))
         r_str_DesErr(UBound(r_str_DesErr)).NumOpe = Trim(r_str_NumOpe)
         r_str_DesErr(UBound(r_str_DesErr)).OpeMvi = r_str_CalDes(r_lng_Contad - 1).OpeMvi
         r_str_DesErr(UBound(r_str_DesErr)).MtoErr = "El Saldo no es cero."
      End If
      If r_bol_flgPGr = False Then
         ReDim Preserve r_str_DesErr(UBound(r_str_DesErr) + 1)
         r_str_DesErr(UBound(r_str_DesErr)).NomArc = Trim(Mid(p_ArcCOFIDE, InStrRev(p_ArcCOFIDE, "\") + 1))
         r_str_DesErr(UBound(r_str_DesErr)).NumOpe = Trim(r_str_NumOpe)
         r_str_DesErr(UBound(r_str_DesErr)).OpeMvi = r_str_CalDes(r_lng_Contad - 1).OpeMvi
         r_str_DesErr(UBound(r_str_DesErr)).MtoErr = "No tiene Periodo de Gracia."
      End If
      If r_bol_flgCDb = False Then
         ReDim Preserve r_str_DesErr(UBound(r_str_DesErr) + 1)
         r_str_DesErr(UBound(r_str_DesErr)).NomArc = Trim(Mid(p_ArcCOFIDE, InStrRev(p_ArcCOFIDE, "\") + 1))
         r_str_DesErr(UBound(r_str_DesErr)).NumOpe = Trim(r_str_NumOpe)
         r_str_DesErr(UBound(r_str_DesErr)).OpeMvi = r_str_CalDes(r_lng_Contad - 1).OpeMvi
         r_str_DesErr(UBound(r_str_DesErr)).MtoErr = "Cuotas Dobles incorrectas."
      End If
      If r_bol_flgNCu = False Then
         ReDim Preserve r_str_DesErr(UBound(r_str_DesErr) + 1)
         r_str_DesErr(UBound(r_str_DesErr)).NomArc = Trim(Mid(p_ArcCOFIDE, InStrRev(p_ArcCOFIDE, "\") + 1))
         r_str_DesErr(UBound(r_str_DesErr)).NumOpe = Trim(r_str_NumOpe)
         r_str_DesErr(UBound(r_str_DesErr)).OpeMvi = r_str_CalDes(r_lng_Contad - 1).OpeMvi
         r_str_DesErr(UBound(r_str_DesErr)).MtoErr = "El número de cuotas no es correcto."
      End If
   End If
   
   If r_bol_FlgIng = True Then
      'Ingresos Correctos
      ReDim Preserve r_str_DesErr(UBound(r_str_DesErr) + 1)
      r_str_DesErr(UBound(r_str_DesErr)).NomArc = Trim(Mid(p_ArcCOFIDE, InStrRev(p_ArcCOFIDE, "\") + 1))
      r_str_DesErr(UBound(r_str_DesErr)).NumOpe = Trim(r_str_NumOpe)
      r_str_DesErr(UBound(r_str_DesErr)).OpeMvi = r_str_CalDes(r_lng_Contad - 1).OpeMvi
      r_str_DesErr(UBound(r_str_DesErr)).MtoErr = "Calendario cargado correctamente."
   End If
End Sub

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PerAno)
   End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CodPro)
   End If
End Sub

Private Sub cmb_CodPro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(drv_LisUni)
   End If
End Sub

Private Sub drv_LisUni_Change()
   On Error Resume Next
   dir_LisCar.Path = drv_LisUni.Drive
End Sub

Private Sub dir_LisCar_Change()
   fil_LisArc.Path = dir_LisCar.Path
End Sub

Private Function ff_Inserta_HipCuo(ByVal p_NumOpe As String, ByVal p_TipCro As Integer, ByVal p_NumCuo As Integer, ByVal p_FecVct As String, ByVal p_Capita As Double, ByVal p_intere As Double, ByVal p_SegDes As Double, ByVal p_SegViv As Double, ByVal p_OtrGas As Double, ByVal p_SalCap As Double, ByVal p_ComCrc As Double, ByVal p_ComPbp As Double, ByVal p_ComCof As Double) As Integer
'Dim r_lng_Contad     As Long

   ff_Inserta_HipCuo = False
   
'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & " SELECT NVL(COUNT(*),0) AS TOTREG  "
'   g_str_Parame = g_str_Parame & "   FROM CRE_HIPCUO"
'   g_str_Parame = g_str_Parame & "  WHERE HIPCUO_NUMOPE = '" & p_NumOpe & "'"
'   g_str_Parame = g_str_Parame & "    AND HIPCUO_TIPCRO = 3 "
'
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
'      Exit Function
'   End If
'
'   g_rst_Princi.MoveFirst
'   r_lng_Contad = g_rst_Princi!TOTREG
'
'   g_rst_Princi.Close
'   Set g_rst_Princi = Nothing
'
'   If r_lng_Contad = 0 Then
    
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CRE_HIPCUO_CREA ("
   g_str_Parame = g_str_Parame & "'" & p_NumOpe & "', "
   g_str_Parame = g_str_Parame & CStr(p_TipCro) & ", "
   g_str_Parame = g_str_Parame & CStr(p_NumCuo) & ", "
   g_str_Parame = g_str_Parame & Format(CDate(p_FecVct), "yyyymmdd") & ", "
   g_str_Parame = g_str_Parame & CStr(p_Capita) & ", "
   g_str_Parame = g_str_Parame & CStr(p_intere) & ", "
   g_str_Parame = g_str_Parame & CStr(p_SegDes) & ", "
   g_str_Parame = g_str_Parame & CStr(p_SegViv) & ", "
   g_str_Parame = g_str_Parame & CStr(p_OtrGas) & ", "
   g_str_Parame = g_str_Parame & CStr(p_SalCap) & ", "
   g_str_Parame = g_str_Parame & CStr(p_ComCrc) & ", "
   g_str_Parame = g_str_Parame & CStr(p_ComPbp) & ", "
   g_str_Parame = g_str_Parame & CStr(p_ComCof) & ", "
   
   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                           'Código Usuario
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                           'Nombre Terminal
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                            'Nombre Ejecutable
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                           'Código Sucursal
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      moddat_g_int_CntErr = moddat_g_int_CntErr + 1
       ff_Inserta_HipCuo = False
   Else
'      moddat_g_int_FlgGOK = True
      ff_Inserta_HipCuo = True
   End If
'   End If
End Function

