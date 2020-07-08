VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Tas_ActReg_04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   9240
   ClientLeft      =   1335
   ClientTop       =   2715
   ClientWidth     =   13620
   Icon            =   "OpeTra_frm_338.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   13620
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9240
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   13635
      _Version        =   65536
      _ExtentX        =   24051
      _ExtentY        =   16298
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
         Height          =   645
         Left            =   60
         TabIndex        =   9
         Top             =   810
         Width           =   13515
         _Version        =   65536
         _ExtentX        =   23839
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
            Left            =   1830
            Picture         =   "OpeTra_frm_338.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_338.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Buscar Registros"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_338.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   12915
            Picture         =   "OpeTra_frm_338.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_338.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   900
         Left            =   60
         TabIndex        =   10
         Top             =   1500
         Width           =   13515
         _Version        =   65536
         _ExtentX        =   23839
         _ExtentY        =   1587
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
            ItemData        =   "OpeTra_frm_338.frx":11AE
            Left            =   1620
            List            =   "OpeTra_frm_338.frx":11B0
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   120
            Width           =   2265
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1620
            TabIndex        =   1
            Top             =   465
            Width           =   825
            _Version        =   196608
            _ExtentX        =   1455
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
            MaxValue        =   "9999"
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
         Begin VB.Label Label4 
            Caption         =   "Mes de Proceso:"
            Height          =   315
            Left            =   150
            TabIndex        =   16
            Top             =   150
            Width           =   1395
         End
         Begin VB.Label Label3 
            Caption         =   "Año de Proceso:"
            Height          =   285
            Left            =   150
            TabIndex        =   15
            Top             =   510
            Width           =   1245
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   705
         Left            =   60
         TabIndex        =   11
         Top             =   60
         Width           =   13515
         _Version        =   65536
         _ExtentX        =   23839
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   315
            Left            =   660
            TabIndex        =   14
            Top             =   60
            Width           =   4995
            _Version        =   65536
            _ExtentX        =   8811
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Operaciones Financieras"
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Left            =   660
            TabIndex        =   12
            Top             =   360
            Width           =   4995
            _Version        =   65536
            _ExtentX        =   8811
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Proceso de Actualización de Garantías"
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
            Picture         =   "OpeTra_frm_338.frx":11B2
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   6735
         Left            =   60
         TabIndex        =   13
         Top             =   2445
         Width           =   13515
         _Version        =   65536
         _ExtentX        =   23839
         _ExtentY        =   11880
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
            Height          =   6630
            Left            =   60
            TabIndex        =   7
            Top             =   60
            Width           =   13410
            _ExtentX        =   23654
            _ExtentY        =   11695
            _Version        =   393216
            Rows            =   30
            Cols            =   14
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            SelectionMode   =   1
         End
      End
   End
End
Attribute VB_Name = "frm_Tas_ActReg_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private Sub cmd_Buscar_Click()
Dim r_str_Parame     As String
Dim r_int_Contad     As Integer
Dim r_rst_RegGar     As ADODB.Recordset
   
   'valida datos
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el mes de proceso.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If CInt(ipp_PerAno.Text) < 2012 Then
      MsgBox "Ingrese correctamente el año de proceso.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   
   'valida ejecucion
   r_str_Parame = ff_ObtieneDatos(CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)), CStr(ipp_PerAno.Text))
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_RegGar, 3) Then
      Exit Sub
   End If
   
   r_rst_RegGar.MoveFirst
   r_int_Contad = r_rst_RegGar!CONTADOR
   
   r_rst_RegGar.Close
   Set r_rst_RegGar = Nothing
   
   If r_int_Contad > 0 Then
      MsgBox "Proceso ya fue ejecutado para el periodo seleccionado.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Confirmacion
   If MsgBox("¿Está seguro de procesar la información?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call gs_LimpiaGrid(grd_Listad)
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_ExpExc_Click()
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el mes de proceso.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If CInt(ipp_PerAno.Text) < 2012 Then
      MsgBox "Ingrese correctamente el año de proceso.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
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
Dim r_int_Cont As Integer
   
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el mes de proceso.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If CInt(ipp_PerAno.Text) < 2012 Then
      MsgBox "Ingrese correctamente el año de proceso.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   If grd_Listad.Rows = 0 Then
      MsgBox "No existen registros que grabar.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If MsgBox("¿Este seguro de guardar la información procesada?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
      
   Screen.MousePointer = 11
   
   For r_int_Cont = 1 To grd_Listad.Rows - 1
      g_str_Parame = ""
      g_str_Parame = "USP_CRE_ACTGAR ("
      g_str_Parame = g_str_Parame & "'" & NumeroMes(cmb_PerMes.Text) & "',"
      g_str_Parame = g_str_Parame & "'" & ipp_PerAno.Text & "',"
      g_str_Parame = g_str_Parame & "'" & Left(grd_Listad.TextMatrix(r_int_Cont, 1), 3) & Mid(grd_Listad.TextMatrix(r_int_Cont, 1), 5, 2) & Right(grd_Listad.TextMatrix(r_int_Cont, 1), 5) & "', "
      g_str_Parame = g_str_Parame & "'" & Format(date, "yyyymmdd") & "',"
      g_str_Parame = g_str_Parame & "" & Left(grd_Listad.TextMatrix(r_int_Cont, 3), 1) & ","
      g_str_Parame = g_str_Parame & "'" & Right(grd_Listad.TextMatrix(r_int_Cont, 3), 8) & "',"
      g_str_Parame = g_str_Parame & "" & Format(grd_Listad.TextMatrix(r_int_Cont, 6), "########0.00") & ","
      g_str_Parame = g_str_Parame & "'" & Format(grd_Listad.TextMatrix(r_int_Cont, 8), "yyyymmdd") & "',"
      g_str_Parame = g_str_Parame & "'" & grd_Listad.TextMatrix(r_int_Cont, 9) & "',"
      g_str_Parame = g_str_Parame & "'" & grd_Listad.TextMatrix(r_int_Cont, 10) & "',"
      g_str_Parame = g_str_Parame & "" & Format(grd_Listad.TextMatrix(r_int_Cont, 11), "########0.00") & ","
      g_str_Parame = g_str_Parame & "" & Format(grd_Listad.TextMatrix(r_int_Cont, 12), "########0.00") & ","
      g_str_Parame = g_str_Parame & "" & Format(grd_Listad.TextMatrix(r_int_Cont, 13), "########0.00") & ","
      g_str_Parame = g_str_Parame & "'" & grd_Listad.TextMatrix(r_int_Cont, 14) & "',"
      g_str_Parame = g_str_Parame & "'" & grd_Listad.TextMatrix(r_int_Cont, 15) & "',"
      g_str_Parame = g_str_Parame & "'" & grd_Listad.TextMatrix(r_int_Cont, 16) & "',"
      g_str_Parame = g_str_Parame & "" & grd_Listad.TextMatrix(r_int_Cont, 17) & ","
      g_str_Parame = g_str_Parame & "" & Format(grd_Listad.TextMatrix(r_int_Cont, 18), "########0.00") & ","
      g_str_Parame = g_str_Parame & "" & Format(grd_Listad.TextMatrix(r_int_Cont, 19), "########0.00") & ","
      g_str_Parame = g_str_Parame & "" & Format(grd_Listad.TextMatrix(r_int_Cont, 20), "########0.00") & ","
      g_str_Parame = g_str_Parame & "" & Format(grd_Listad.TextMatrix(r_int_Cont, 21), "########0.00") & ","
      g_str_Parame = g_str_Parame & "" & Format(grd_Listad.TextMatrix(r_int_Cont, 22), "########0.00") & ","
      g_str_Parame = g_str_Parame & "" & Format(grd_Listad.TextMatrix(r_int_Cont, 23), "########0.00") & ","
      g_str_Parame = g_str_Parame & "" & Format(grd_Listad.TextMatrix(r_int_Cont, 24), "########0.00") & ","
      g_str_Parame = g_str_Parame & "" & Format(grd_Listad.TextMatrix(r_int_Cont, 25), "########0.00") & ","
      g_str_Parame = g_str_Parame & "" & Format(grd_Listad.TextMatrix(r_int_Cont, 26), "########0.00") & ","
      g_str_Parame = g_str_Parame & "" & Format(grd_Listad.TextMatrix(r_int_Cont, 27), "########0.00") & ","
      g_str_Parame = g_str_Parame & "" & Format(grd_Listad.TextMatrix(r_int_Cont, 28), "########0.00") & ","
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "',"          'Código Usuario
      g_str_Parame = g_str_Parame & "'" & Format(date, "yyyymmdd") & "',"     'Fecha
      g_str_Parame = g_str_Parame & "'" & Format(Time, "hhmmss") & "',"       'Hora/Minuto/Segundo
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "',"           'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "',"          'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"          'Nombre Sucursal

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         MsgBox "Error al ejecutar el procedimiento.", vbCritical, modgen_g_str_NomPlt
         Screen.MousePointer = 0
         Exit Sub
      End If
   Next
   
   MsgBox "La información se ha registrado satisfactoriamente.", vbInformation, modgen_g_str_NomPlt
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_PerMes)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Activa(True)
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.Cols = 29
   grd_Listad.ColWidth(0) = 800        'ITEM
   grd_Listad.ColWidth(1) = 1300       'OPERACION
   grd_Listad.ColWidth(2) = 1200       'PRODUCTO
   grd_Listad.ColWidth(3) = 1700       'TIPO DOCUMENTO
   grd_Listad.ColWidth(4) = 3500       'NOMBRE CLIENTE
   grd_Listad.ColWidth(5) = 2200       'TIPO MONEDA
   grd_Listad.ColWidth(6) = 1200       'SALDO CAPITAL
   grd_Listad.ColWidth(7) = 2200       'DISTRITO
   grd_Listad.ColWidth(8) = 1200       'FECHA TASACION
   grd_Listad.ColWidth(9) = 1200       'MES TASACION
   grd_Listad.ColWidth(10) = 1200      'AÑO TASACION
   grd_Listad.ColWidth(11) = 1200      'AREA CONTRUIDA
   grd_Listad.ColWidth(12) = 1200      'AREA TERRENO
   grd_Listad.ColWidth(13) = 1200      'VALOR COMERCIAL
   grd_Listad.ColWidth(14) = 1500      'AÑO CONSTRUCCION
   grd_Listad.ColWidth(15) = 1500      'MATERIAL CONSTRUCCION
   grd_Listad.ColWidth(16) = 1500      'ESTADO CONSERVACION
   grd_Listad.ColWidth(17) = 1400      'ANTIGUEDAD
   grd_Listad.ColWidth(18) = 1500      'DEPRECIACION
   grd_Listad.ColWidth(19) = 1500      'VALOR M2 TERRENO
   grd_Listad.ColWidth(20) = 1500      'VALOR M2 CONSTRUCCION
   grd_Listad.ColWidth(21) = 1500      'VALOR ACTUAL
   grd_Listad.ColWidth(22) = 1500      'VALOR ACTUALIZADO
   grd_Listad.ColWidth(23) = 1500      'RELACION VAL ACT / VAL COM
   grd_Listad.ColWidth(24) = 1500      'VALOR CORREGIDO - VALOR TERRENO
   grd_Listad.ColWidth(25) = 1500      'VALOR CORREGIDO - VALOR CONSTRUCCION
   grd_Listad.ColWidth(26) = 1500      'VALOR CORREGIDO - VALOR ACTUALIZADO
   grd_Listad.ColWidth(27) = 1600      'VALOR CORREGIDO - RELACION VAL ACT / VAL COM
   grd_Listad.ColWidth(28) = 1200      'LTV
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_Listad.ColAlignment(7) = flexAlignCenterCenter
   grd_Listad.ColAlignment(8) = flexAlignCenterCenter
   grd_Listad.ColAlignment(9) = flexAlignCenterCenter
   grd_Listad.ColAlignment(10) = flexAlignCenterCenter
   grd_Listad.ColAlignment(11) = flexAlignRightCenter
   grd_Listad.ColAlignment(12) = flexAlignRightCenter
   grd_Listad.ColAlignment(13) = flexAlignRightCenter
   grd_Listad.ColAlignment(14) = flexAlignCenterCenter
   grd_Listad.ColAlignment(15) = flexAlignCenterCenter
   grd_Listad.ColAlignment(16) = flexAlignCenterCenter
   grd_Listad.ColAlignment(17) = flexAlignCenterCenter
   grd_Listad.ColAlignment(18) = flexAlignCenterCenter
   grd_Listad.ColAlignment(19) = flexAlignRightCenter
   grd_Listad.ColAlignment(20) = flexAlignRightCenter
   grd_Listad.ColAlignment(21) = flexAlignRightCenter
   grd_Listad.ColAlignment(22) = flexAlignRightCenter
   grd_Listad.ColAlignment(23) = flexAlignRightCenter
   grd_Listad.ColAlignment(24) = flexAlignRightCenter
   grd_Listad.ColAlignment(25) = flexAlignRightCenter
   grd_Listad.ColAlignment(26) = flexAlignRightCenter
   grd_Listad.ColAlignment(27) = flexAlignRightCenter
   grd_Listad.ColAlignment(28) = flexAlignRightCenter
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
End Sub

Private Sub fs_Limpia()
   cmb_PerMes.ListIndex = -1
   ipp_PerAno.Text = Year(date)
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_PerMes.Enabled = p_Activa
   ipp_PerAno.Enabled = p_Activa
   grd_Listad.Enabled = Not p_Activa
   cmd_ExpExc.Enabled = Not p_Activa
   cmd_Grabar.Enabled = Not p_Activa
End Sub

Function NumeroMes(mes As String) As String
   Select Case mes
      Case "ENERO":     NumeroMes = "01"
      Case "FEBRERO":   NumeroMes = "02"
      Case "MARZO":     NumeroMes = "03"
      Case "ABRIL":     NumeroMes = "04"
      Case "MAYO":      NumeroMes = "05"
      Case "JUNIO":     NumeroMes = "06"
      Case "JULIO":     NumeroMes = "07"
      Case "AGOSTO":    NumeroMes = "08"
      Case "SETIEMBRE": NumeroMes = "09"
      Case "OCTUBRE":   NumeroMes = "10"
      Case "NOVIEMBRE": NumeroMes = "11"
      Case "DICIEMBRE": NumeroMes = "12"
   End Select
End Function

Private Function ff_ObtieneDatos(ByVal p_MesProc As String, ByVal p_AnioProc As String) As String
   ff_ObtieneDatos = ""
   ff_ObtieneDatos = ff_ObtieneDatos & "SELECT COUNT(*) AS CONTADOR "
   ff_ObtieneDatos = ff_ObtieneDatos & "  FROM CRE_ACTGAR "
   ff_ObtieneDatos = ff_ObtieneDatos & " WHERE ACTGAR_MESPRO = '" & Format(CInt(p_MesProc), "00") & "' "
   ff_ObtieneDatos = ff_ObtieneDatos & "   AND ACTGAR_ANOPRO = '" & p_AnioProc & "' "
End Function

Private Sub fs_Buscar()
Dim r_str_Param1     As String
Dim r_str_Param2     As String
Dim r_str_FecIni     As String
Dim r_str_FecFin     As String
Dim r_int_Contad     As Integer
Dim r_int_PerMes     As Integer
Dim r_int_PerAno     As Integer
Dim r_dbl_ValTer1    As Double
Dim r_dbl_ValCon1    As Double
Dim r_dbl_ValTer2    As Double
Dim r_dbl_ValCon2    As Double
Dim r_dbl_Tempor     As Double

   'Setea parametros de fechas de busqueda
   If CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) = 1 Then
      r_int_PerMes = 12
      r_int_PerAno = CInt(ipp_PerAno.Text) - 1
   Else
      r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) - 1
      r_int_PerAno = CInt(ipp_PerAno.Text)
   End If
   r_str_FecFin = Format(r_int_PerAno, "0000") & Format(r_int_PerMes, "00") & ff_Ultimo_Dia_Mes(r_int_PerMes, r_int_PerAno)
   r_str_FecIni = Format(r_int_PerAno - 1, "0000") & Format(r_int_PerMes, "00") & ff_Ultimo_Dia_Mes(r_int_PerMes, r_int_PerAno)
   
   'Obtiene datos de las operaciones
   r_str_Param1 = ff_Obtiene_Operaciones_Cierre(CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)), CInt(ipp_PerAno.Text), r_str_FecIni)
   
   If Not gf_EjecutaSQL(r_str_Param1, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se han encontrado registros de operaciones", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   
   'Obtiene datos de las tasaciones
   r_str_Param2 = ff_Obtiene_Tasaciones(r_str_FecIni, r_str_FecFin)
   
   If Not gf_EjecutaSQL(r_str_Param2, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      MsgBox "No se han encontrado registros de tasaciones", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Call fs_Activa(True)
      Exit Sub
   End If
   
   grd_Listad.Redraw = False
   Call fs_Activa(False)
   
   'Primera Linea
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.RowHeight(0) = 440
   grd_Listad.WordWrap = True
   
   grd_Listad.Col = 0:   grd_Listad.Text = "ITEM":                   grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 1:   grd_Listad.Text = "OPERACION":              grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 2:   grd_Listad.Text = "PRODUCTO":               grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 3:   grd_Listad.Text = "TIPO DOCUMENTO":         grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 4:   grd_Listad.Text = "NOMBRE CLIENTE":         grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 5:   grd_Listad.Text = "TIPO MONEDA":            grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 6:   grd_Listad.Text = "SALDO CAPITAL":          grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 7:   grd_Listad.Text = "DISTRITO":               grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 8:   grd_Listad.Text = "FECHA TASACION":         grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 9:   grd_Listad.Text = "MES TASACION":           grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 10:  grd_Listad.Text = "AÑO TASACION":           grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 11:  grd_Listad.Text = "AREA CONSTRUIDA":        grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 12:  grd_Listad.Text = "AREA TERRENO":           grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 13:  grd_Listad.Text = "VALOR COMERCIAL":        grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 14:  grd_Listad.Text = "AÑO CONSTRUCCION":       grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 15:  grd_Listad.Text = "MATERIAL CONSTRUCCION":  grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 16:  grd_Listad.Text = "ESTADO CONSERVACION":    grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 17:  grd_Listad.Text = "ANTIGUEDAD (AÑOS)":      grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 18:  grd_Listad.Text = "DEPRECIACION (%)":       grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 19:  grd_Listad.Text = "VALOR M2 TERRENO":       grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 20:  grd_Listad.Text = "VALOR M2 CONSTRUCCION":  grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 21:  grd_Listad.Text = "VALOR ACTUAL":           grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 22:  grd_Listad.Text = "VALOR ACTUALIZADO":      grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 23:  grd_Listad.Text = "VALACT/VALCOM (%)":      grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 24:  grd_Listad.Text = "CORREGIDO VAL TER":      grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 25:  grd_Listad.Text = "CORREGIDO VAL CONST":    grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 26:  grd_Listad.Text = "CORREGIDO VAL ACTUAL":   grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 27:  grd_Listad.Text = "CORREGIDO ACT/COM (%)":  grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 28:  grd_Listad.Text = "LTV (%)":                grd_Listad.CellAlignment = flexAlignCenterCenter
   
   grd_Listad.Redraw = False
   r_int_Contad = 0
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      
      If Trim(g_rst_Princi!OPERACION) = "0241900010" Then
         MsgBox "Pausa"
      End If
      
      If fs_ValidaOperacion(g_rst_Princi!OPERACION, CStr(ipp_PerAno.Text), CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)), r_str_FecIni) Then
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         r_int_Contad = r_int_Contad + 1
         
         'Numero de item
         grd_Listad.Col = 0:  grd_Listad.Text = Format(r_int_Contad, "0000")
         
         'Numero operacion
         grd_Listad.Col = 1:  grd_Listad.Text = gf_Formato_NumOpe(Trim(g_rst_Princi!OPERACION & ""))
         
         'Producto
         grd_Listad.Col = 2:  grd_Listad.Text = g_rst_Princi!PRODUCTO
         
         'Tipo de Documento
         grd_Listad.Col = 3:  grd_Listad.Text = g_rst_Princi!TIPO_DOCUMENTO
         
         'Nombre del Cliente
         grd_Listad.Col = 4:  grd_Listad.Text = g_rst_Princi!NOMBRE_CLIENTE
         
         'Tipo de Moneda
         grd_Listad.Col = 5:  grd_Listad.Text = g_rst_Princi!TIPO_MONEDA
         
         'Saldo Capital
         grd_Listad.Col = 6:  grd_Listad.Text = Format(g_rst_Princi!SALDO_CAPITAL, "###,##0.00")
         
         'Distrito
         grd_Listad.Col = 7:  grd_Listad.Text = g_rst_Princi!DISTRITO
         
         'Fecha tasacion
         grd_Listad.Col = 8:  grd_Listad.Text = g_rst_Princi!FECHA_TASACION
         
         'Mes tasacion
         grd_Listad.Col = 9:  grd_Listad.Text = g_rst_Princi!MES_TASACION
         
         'Año tasacion
         grd_Listad.Col = 10: grd_Listad.Text = g_rst_Princi!ANIO_TASACION
         
         'Area construida
         grd_Listad.Col = 11: grd_Listad.Text = Format(g_rst_Princi!AREA_CONSTRUIDA, "###,##0.00")
         
         'Area terreno
         grd_Listad.Col = 12: grd_Listad.Text = Format(g_rst_Princi!AREA_TERRENO, "###,##0.00")
         
         'Valor comercial
         grd_Listad.Col = 13: grd_Listad.Text = Format(g_rst_Princi!VALOR_COMERCIAL, "###,###,##0.00")
         
         'Año construccion
         grd_Listad.Col = 14: grd_Listad.Text = g_rst_Princi!ANIO_CONSTRUCCION
         
         'Material construccion
         grd_Listad.Col = 15: grd_Listad.Text = g_rst_Princi!MATERIAL_CONSTRUCCION
         
         'Estado conservacion
         grd_Listad.Col = 16: grd_Listad.Text = g_rst_Princi!ESTADO_CONSERVACION
         
         'Antiguedad
         grd_Listad.Col = 17: grd_Listad.Text = g_rst_Princi!ANTIGUEDAD_ACTUAL
         
         'Depreciacion
         grd_Listad.Col = 18: grd_Listad.Text = g_rst_Princi!DEPRECIACION
         
         r_dbl_ValTer1 = 0
         r_dbl_ValCon1 = 0
         r_dbl_ValTer2 = 0
         r_dbl_ValCon2 = 0
         
         g_rst_Listas.MoveFirst
         Do While Not g_rst_Listas.EOF
            If Trim(g_rst_Princi!DISTRITO) = Trim(g_rst_Listas!DISTRITO) Then
               r_dbl_ValTer1 = g_rst_Listas!VALOR_M2_TERRENO_1
               r_dbl_ValCon1 = g_rst_Listas!VALOR_M2_CONSTRUCCION_1
               r_dbl_ValTer2 = g_rst_Listas!VALOR_M2_TERRENO_2
               r_dbl_ValCon2 = g_rst_Listas!VALOR_M2_CONSTRUCCION_2
            End If
            g_rst_Listas.MoveNext
         Loop
         
         'Valor m2 terreno
         grd_Listad.Col = 19: grd_Listad.Text = Format(r_dbl_ValTer1, "###,###,##0.00")
         
         'Valor m2 construccion
         grd_Listad.Col = 20: grd_Listad.Text = Format(r_dbl_ValCon1, "###,###,##0.00")
         
         'Valor actual
         grd_Listad.Col = 21: grd_Listad.Text = Format((g_rst_Princi!AREA_CONSTRUIDA * r_dbl_ValCon1) + (g_rst_Princi!AREA_TERRENO * r_dbl_ValTer1), "###,###,##0.00")
         
         'Valor actualizado
         r_dbl_Tempor = ((g_rst_Princi!AREA_CONSTRUIDA * r_dbl_ValCon1) + (g_rst_Princi!AREA_TERRENO * r_dbl_ValTer1)) * ((100 - g_rst_Princi!DEPRECIACION) / 100)
         grd_Listad.Col = 22: grd_Listad.Text = Format(r_dbl_Tempor, "###,###,##0.00")
         
         'Relacion ValAct / ValCom
         If g_rst_Princi!VALOR_COMERCIAL = 0 Then
            r_dbl_Tempor = 0
         Else
            r_dbl_Tempor = (r_dbl_Tempor / g_rst_Princi!VALOR_COMERCIAL) - 1
         End If
         grd_Listad.Col = 23: grd_Listad.Text = Format(r_dbl_Tempor * 100, "###,###,##0.00")
         
         'Valor corregido - valor terreno
         grd_Listad.Col = 24: grd_Listad.Text = Format(g_rst_Princi!AREA_TERRENO * r_dbl_ValTer2, "###,###,##0.00")
         
         'Valor corregido - valor construccion
         grd_Listad.Col = 25: grd_Listad.Text = Format(g_rst_Princi!AREA_CONSTRUIDA * r_dbl_ValCon2, "###,###,##0.00")
         
         'Valor corregido - valor actualizado
         r_dbl_Tempor = ((g_rst_Princi!AREA_TERRENO * r_dbl_ValTer2) + (g_rst_Princi!AREA_CONSTRUIDA * r_dbl_ValCon2)) * ((100 - g_rst_Princi!DEPRECIACION) / 100)
         grd_Listad.Col = 26: grd_Listad.Text = Format(r_dbl_Tempor, "###,###,##0.00")
         
         'Valor corregido - Relacion ValAct / ValCom
         If g_rst_Princi!VALOR_COMERCIAL = 0 Then
            r_dbl_Tempor = 0
         Else
            r_dbl_Tempor = (r_dbl_Tempor / g_rst_Princi!VALOR_COMERCIAL) - 1
         End If
         grd_Listad.Col = 27: grd_Listad.Text = Format(r_dbl_Tempor * 100, "###,###,##0.00")
         
         'LTV
         r_dbl_Tempor = ((g_rst_Princi!AREA_TERRENO * r_dbl_ValTer2) + (g_rst_Princi!AREA_CONSTRUIDA * r_dbl_ValCon2)) * ((100 - g_rst_Princi!DEPRECIACION) / 100)
         
         If r_dbl_Tempor > 0 Then
            grd_Listad.Col = 28: grd_Listad.Text = Format((g_rst_Princi!SALDO_CAPITAL / r_dbl_Tempor) * 100, "##0.00")
         Else
            grd_Listad.Col = 28: grd_Listad.Text = "0.00"
         End If
         
      End If
      g_rst_Princi.MoveNext
   Loop
   
   If r_int_Contad = 0 Then
      With grd_Listad
         .FixedCols = 1
      End With
   Else
      With grd_Listad
         .FixedCols = 1
         .FixedRows = 1
      End With
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Function ff_Obtiene_Operaciones_Cierre(ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, p_FecIni As String) As String
   
   ff_Obtiene_Operaciones_Cierre = ""
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "SELECT TRIM(X.HIPCIE_NUMOPE)                                                                             AS OPERACION, "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "       CASE WHEN X.HIPCIE_CODPRD IN (" & moddat_g_str_AgrTMIC & ") THEN 'MICASITA' ELSE 'MIVIVIENDA' END AS PRODUCTO, "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "       TRIM(B.DATGEN_TIPDOC)||'-'||TRIM(B.DATGEN_NUMDOC)                                                 AS TIPO_DOCUMENTO, "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "       TRIM(B.DATGEN_APEPAT)||' '||TRIM(B.DATGEN_APEMAT)||' '||TRIM(B.DATGEN_NOMBRE)                     AS NOMBRE_CLIENTE, "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "       CASE WHEN X.HIPCIE_TIPMON = 1 THEN 'SOLES' ELSE 'DOLARES AMERICANOS' END                          AS TIPO_MONEDA, "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "       DECODE(X.HIPCIE_TIPMON, 1, X.HIPCIE_SALCAP+X.HIPCIE_SALCON, (X.HIPCIE_SALCAP+X.HIPCIE_SALCON)*X.HIPCIE_TIPCAM) AS SALDO_CAPITAL, "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "       TRIM(D.PARDES_DESCRI)                                                                             AS DISTRITO, "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "       SUBSTR(E.EVATAS_FECEVA,7,2)||'/'||SUBSTR(E.EVATAS_FECEVA,5,2)||'/'||SUBSTR(E.EVATAS_FECEVA,1,4)   AS FECHA_TASACION, "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "       SUBSTR(E.EVATAS_FECEVA,5,2)                                                                       AS MES_TASACION, "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "       SUBSTR(E.EVATAS_FECEVA,1,4)                                                                       AS ANIO_TASACION, "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "       E.EVATAS_ARECON_INM+E.EVATAS_ARECON_ES1+E.EVATAS_ARECON_ES2+E.EVATAS_ARECON_DEP                   AS AREA_CONSTRUIDA, "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "       E.EVATAS_ARETER_INM+E.EVATAS_ARETER_ES1+E.EVATAS_ARETER_ES2+E.EVATAS_ARETER_DEP                   AS AREA_TERRENO, "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "       ROUND(DECODE(E.EVATAS_TIPMON, 1, (E.EVATAS_VALCOM_INM+E.EVATAS_VALCOM_ES1+E.EVATAS_VALCOM_ES2+E.EVATAS_VALCOM_DEP), (E.EVATAS_VALCOM_INM+E.EVATAS_VALCOM_ES1+E.EVATAS_VALCOM_ES2+E.EVATAS_VALCOM_DEP)*E.EVATAS_TIPCAM),2) AS VALOR_COMERCIAL, "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "       E.EVATAS_ANOCON                                                                                   AS ANIO_CONSTRUCCION, "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "       CASE WHEN SUBSTR(X.HIPCIE_FECDES,1,4)-E.EVATAS_ANOCON > 20 THEN 'LADRILLO' ELSE 'CONCRETO' END    AS MATERIAL_CONSTRUCCION, "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "       CASE WHEN SUBSTR(X.HIPCIE_FECDES,1,4)-E.EVATAS_ANOCON > 20 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "            THEN "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                CASE WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 15 THEN 'MUY BUENO' "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                     WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 50 THEN 'BUENO' "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                     ELSE 'REGULAR' "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                END "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "            ELSE "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                CASE WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 20 THEN 'MUY BUENO' "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                     WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 50 THEN 'BUENO' "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                     ELSE 'REGULAR' "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                END "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "       END  AS ESTADO_CONSERVACION, "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "       SUBSTR(X.HIPCIE_FECDES,1,4)-E.EVATAS_ANOCON                                                      AS ANTIGUEDAD_INICIAL, "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "       SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON                                         AS ANTIGUEDAD_ACTUAL, "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "       CASE WHEN  SUBSTR(X.HIPCIE_FECDES,1,4)-E.EVATAS_ANOCON > 20 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "            THEN "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                CASE WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 15 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                     THEN CASE WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 5  THEN 0 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 10 THEN 3 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 15 THEN 6 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 20 THEN 9 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 25 THEN 12 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 30 THEN 15 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 35 THEN 18 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 40 THEN 21 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 45 THEN 24 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 50 THEN 27 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON >  50 THEN 30 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                          END "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                     WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 50 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                     THEN CASE WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 5  THEN 8 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 10 THEN 11 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 15 THEN 14 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 20 THEN 17 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 25 THEN 20 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 30 THEN 23 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 35 THEN 26 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 40 THEN 29 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 45 THEN 32 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 50 THEN 35 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON >  50 THEN 38 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                          END "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                     ELSE CASE WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 5  THEN 20 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 10 THEN 23 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 15 THEN 26 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 20 THEN 29 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 25 THEN 32 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 30 THEN 35 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 35 THEN 38 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 40 THEN 41 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 45 THEN 44 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 50 THEN 47 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON >  50 THEN 50 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                          END "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                END "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "            ELSE "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                CASE WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 20 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                     THEN CASE WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 5  THEN 0 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 10 THEN 0 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 15 THEN 3 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 20 THEN 6 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 25 THEN 9 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 30 THEN 12 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 35 THEN 15 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 40 THEN 18 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 45 THEN 21 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 50 THEN 24 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON >  50 THEN 27 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                          END "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                     WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 50 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                     THEN CASE WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 5  THEN 5 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 10 THEN 5 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 15 THEN 8 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 20 THEN 11 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 25 THEN 14 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 30 THEN 17 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 35 THEN 20 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 40 THEN 23 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 45 THEN 26 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 50 THEN 29 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON >  50 THEN 32 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                          END "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                     ELSE CASE WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 5  THEN 10 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 10 THEN 10 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 15 THEN 13 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 20 THEN 16 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 25 THEN 19 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 30 THEN 22 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 35 THEN 25 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 40 THEN 28 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 45 THEN 31 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON <= 50 THEN 34 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                               WHEN SUBSTR(TO_CHAR(SYSDATE, 'YYYYMMDD'),1,4)-E.EVATAS_ANOCON >  50 THEN 37 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                          END "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "                END "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "      END                                                                                              AS DEPRECIACION "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "  FROM CRE_HIPCIE X "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & " INNER JOIN CRE_HIPMAE A ON A.HIPMAE_NUMOPE = X.HIPCIE_NUMOPE "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & " INNER JOIN CLI_DATGEN B ON B.DATGEN_TIPDOC = A.HIPMAE_TDOCLI AND B.DATGEN_NUMDOC = A.HIPMAE_NDOCLI "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & " INNER JOIN CRE_SOLINM C ON C.SOLINM_NUMSOL = A.HIPMAE_NUMSOL "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & " INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = '101' AND D.PARDES_CODITE = C.SOLINM_UBIGEO "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & " INNER JOIN TRA_EVATAS E ON E.EVATAS_NUMSOL = A.HIPMAE_NUMSOL AND E.EVATAS_FECEVA < " & p_FecIni & " "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & " INNER JOIN CRE_HIPGAR F ON F.HIPGAR_NUMOPE = A.HIPMAE_NUMOPE AND F.HIPGAR_BIEGAR = 1 "  'AND F.HIPGAR_FECCON < " & p_FecIni & " "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & " WHERE X.HIPCIE_PERMES = " & p_PerMes & " "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "   AND X.HIPCIE_PERANO = " & p_PerAno & " "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & "   AND X.HIPCIE_TIPGAR = 1 "
   ff_Obtiene_Operaciones_Cierre = ff_Obtiene_Operaciones_Cierre & " ORDER BY DISTRITO, OPERACION "
End Function

Private Function ff_Obtiene_Tasaciones(ByVal p_FecIni As String, ByVal p_FecFin As String) As String
   ff_Obtiene_Tasaciones = ""
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & " SELECT DISTINCT TRIM(X.DISTRITO) AS DISTRITO, COUNT(*), "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "        SUM(ROUND(DECODE(X.EVATAS_TIPMON, 1, (X.EVATAS_VALTER_INM+X.EVATAS_VALTER_ES1+X.EVATAS_VALTER_ES2+X.EVATAS_VALTER_DEP), (X.EVATAS_VALTER_INM+X.EVATAS_VALTER_ES1+X.EVATAS_VALTER_ES2+X.EVATAS_VALTER_DEP)*X.EVATAS_TIPCAM), 2)) AS VALOR_TERRENO, "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "        SUM(X.EVATAS_ARETER_INM+X.EVATAS_ARETER_ES1+X.EVATAS_ARETER_ES2+X.EVATAS_ARETER_DEP) AS AREA_TERRENO, "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "        ROUND(SUM(DECODE(X.EVATAS_TIPMON, 1, (X.EVATAS_VALTER_INM+X.EVATAS_VALTER_ES1+X.EVATAS_VALTER_ES2+X.EVATAS_VALTER_DEP), (X.EVATAS_VALTER_INM+X.EVATAS_VALTER_ES1+X.EVATAS_VALTER_ES2+X.EVATAS_VALTER_DEP)*X.EVATAS_TIPCAM)) / (SUM(X.EVATAS_ARETER_INM+X.EVATAS_ARETER_ES1+X.EVATAS_ARETER_ES2+X.EVATAS_ARETER_DEP)), 2) AS VALOR_M2_TERRENO_1, "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "        ROUND(SUM(DECODE(X.EVATAS_TIPMON, 1, (X.EVATAS_VALTER_INM+X.EVATAS_VALTER_ES1+X.EVATAS_VALTER_ES2+X.EVATAS_VALTER_DEP), (X.EVATAS_VALTER_INM+X.EVATAS_VALTER_ES1+X.EVATAS_VALTER_ES2+X.EVATAS_VALTER_DEP)*X.EVATAS_TIPCAM) / (X.EVATAS_ARETER_INM+X.EVATAS_ARETER_ES1+X.EVATAS_ARETER_ES2+X.EVATAS_ARETER_DEP)) / COUNT(*), 2) AS VALOR_M2_TERRENO_2, "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "        ROUND(SUM(DECODE(X.EVATAS_TIPMON, 1, (X.EVATAS_VALEDI_INM+X.EVATAS_VALEDI_ES1+X.EVATAS_VALEDI_ES2+X.EVATAS_VALEDI_DEP), (X.EVATAS_VALEDI_INM+X.EVATAS_VALEDI_ES1+X.EVATAS_VALEDI_ES2+X.EVATAS_VALEDI_DEP)*X.EVATAS_TIPCAM) + DECODE(X.EVATAS_TIPMON, 1, (X.EVATAS_VALACO_INM+X.EVATAS_VALACO_ES1+X.EVATAS_VALACO_ES2+X.EVATAS_VALACO_DEP), (X.EVATAS_VALACO_INM+X.EVATAS_VALACO_ES1+X.EVATAS_VALACO_ES2+X.EVATAS_VALACO_DEP)*X.EVATAS_TIPCAM)) ,2) AS VALOR_CONSTRUCCION, "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "        SUM(X.EVATAS_ARECON_INM+X.EVATAS_ARECON_ES1+X.EVATAS_ARECON_ES2+X.EVATAS_ARECON_DEP) AS AREA_CONSTRUIDA, "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "        ROUND(SUM(DECODE(X.EVATAS_TIPMON, 1, (X.EVATAS_VALEDI_INM+X.EVATAS_VALEDI_ES1+X.EVATAS_VALEDI_ES2+X.EVATAS_VALEDI_DEP), (X.EVATAS_VALEDI_INM+X.EVATAS_VALEDI_ES1+X.EVATAS_VALEDI_ES2+X.EVATAS_VALEDI_DEP)*X.EVATAS_TIPCAM) + ROUND(DECODE(X.EVATAS_TIPMON, 1, (X.EVATAS_VALACO_INM+X.EVATAS_VALACO_ES1+X.EVATAS_VALACO_ES2+X.EVATAS_VALACO_DEP), (X.EVATAS_VALACO_INM+X.EVATAS_VALACO_ES1+X.EVATAS_VALACO_ES2+X.EVATAS_VALACO_DEP)*X.EVATAS_TIPCAM),2)) / (SUM(X.EVATAS_ARECON_INM+X.EVATAS_ARECON_ES1+X.EVATAS_ARECON_ES2+X.EVATAS_ARECON_DEP)) ,2)  AS VALOR_M2_CONSTRUCCION_1 , "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "        ROUND(SUM(DECODE(X.EVATAS_TIPMON, 1, (X.EVATAS_VALEDI_INM+X.EVATAS_VALEDI_ES1+X.EVATAS_VALEDI_ES2+X.EVATAS_VALEDI_DEP), (X.EVATAS_VALEDI_INM+X.EVATAS_VALEDI_ES1+X.EVATAS_VALEDI_ES2+X.EVATAS_VALEDI_DEP)*X.EVATAS_TIPCAM) / (X.EVATAS_ARECON_INM+X.EVATAS_ARECON_ES1+X.EVATAS_ARECON_ES2+X.EVATAS_ARECON_DEP)) / COUNT(*), 2) AS VALOR_M2_CONSTRUCCION_2 "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "  FROM ("
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "        SELECT DISTINCT TRIM(F.PARDES_DESCRI) AS DISTRITO, "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "               C.EVATAS_VALTER_INM, C.EVATAS_VALTER_ES1, C.EVATAS_VALTER_ES2, C.EVATAS_VALTER_DEP, "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "               C.EVATAS_ARETER_INM, C.EVATAS_ARETER_ES1, C.EVATAS_ARETER_ES2, C.EVATAS_ARETER_DEP, "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "               C.EVATAS_VALEDI_INM, C.EVATAS_VALEDI_ES1, C.EVATAS_VALEDI_ES2, C.EVATAS_VALEDI_DEP, "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "               C.EVATAS_VALACO_INM, C.EVATAS_VALACO_ES1, C.EVATAS_VALACO_ES2, C.EVATAS_VALACO_DEP, "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "               C.EVATAS_ARECON_INM, C.EVATAS_ARECON_ES1, C.EVATAS_ARECON_ES2, C.EVATAS_ARECON_DEP, "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "               C.EVATAS_TIPMON, C.EVATAS_TIPCAM "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "          FROM CRE_HIPMAE A "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "         INNER JOIN TRA_EVATAS C ON C.EVATAS_NUMSOL = A.HIPMAE_NUMSOL "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "         INNER JOIN CRE_SOLINM E ON E.SOLINM_NUMSOL = A.HIPMAE_NUMSOL "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "         INNER JOIN MNT_PARDES F ON F.PARDES_CODGRP = '101' AND F.PARDES_CODITE = E.SOLINM_UBIGEO "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "         WHERE A.HIPMAE_SITUAC = 2 "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "           AND C.EVATAS_FECEVA > " & p_FecIni & " AND C.EVATAS_FECEVA < " & p_FecFin & " "
   
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "         UNION ALL "
   
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "        SELECT DISTINCT TRIM(F.PARDES_DESCRI) AS DISTRITO, "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "               C.EVATAS_VALTER_INM, C.EVATAS_VALTER_ES1, C.EVATAS_VALTER_ES2, C.EVATAS_VALTER_DEP, "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "               C.EVATAS_ARETER_INM, C.EVATAS_ARETER_ES1, C.EVATAS_ARETER_ES2, C.EVATAS_ARETER_DEP, "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "               C.EVATAS_VALEDI_INM, C.EVATAS_VALEDI_ES1, C.EVATAS_VALEDI_ES2, C.EVATAS_VALEDI_DEP, "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "               C.EVATAS_VALACO_INM, C.EVATAS_VALACO_ES1, C.EVATAS_VALACO_ES2, C.EVATAS_VALACO_DEP, "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "               C.EVATAS_ARECON_INM, C.EVATAS_ARECON_ES1, C.EVATAS_ARECON_ES2, C.EVATAS_ARECON_DEP, "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "               C.EVATAS_TIPMON, C.EVATAS_TIPCAM "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "          FROM CRE_HIPMAE A "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "         INNER JOIN HIS_EVATAS C ON C.EVATAS_NUMSOL = A.HIPMAE_NUMSOL "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "         INNER JOIN CRE_SOLINM E ON E.SOLINM_NUMSOL = A.HIPMAE_NUMSOL "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "         INNER JOIN MNT_PARDES F ON F.PARDES_CODGRP = '101' AND F.PARDES_CODITE = E.SOLINM_UBIGEO "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "         WHERE C.EVATAS_FECEVA > " & p_FecIni & " AND C.EVATAS_FECEVA < " & p_FecFin & " ) X "
   'ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "        WHERE A.HIPMAE_SITUAC = 2 "
   'ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & "          AND C.EVATAS_FECEVA > " & p_FecIni & " AND C.EVATAS_FECEVA < " & p_FecFin & " ) X "
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & " GROUP BY DISTRITO"
   ff_Obtiene_Tasaciones = ff_Obtiene_Tasaciones & " ORDER BY DISTRITO"
     
End Function

'Private Function ff_Obtiene_Tasaciones2(ByVal p_FecIni As String, ByVal p_FecFin As String) As String
'   ff_Obtiene_Tasaciones2 = ""
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "SELECT DISTINCT DISTRITO AS DISTRITO, COUNT(*), "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "       ROUND(SUM(VALOR_TERRENO), 2) AS VALOR_TERRENO, "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "       ROUND(SUM(AREA_TERRENO), 2) AS AREA_TERRENO, "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "       ROUND(SUM(VALOR_TERRENO) / SUM(AREA_TERRENO), 2) AS VALOR_M2_TERREN0_1, "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "       ROUND(SUM(VALOR_TERRENO / AREA_TERRENO) / COUNT(*), 2) AS VALOR_M2_TERRENO_2, "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "       ROUND(SUM(VALOR_CONSTRUCCION), 2) AS VALOR_CONSTRUCCION, "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "       ROUND(SUM(AREA_CONSTRUIDA), 2) AS AREA_CONSTRUIDA, "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "       ROUND(SUM(VALOR_CONSTRUCCION) / SUM(AREA_CONSTRUIDA), 2) AS VALOR_M2_CONSTRUCCION_1, "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "       ROUND(SUM(VALOR_CONSTRUCCION / AREA_CONSTRUIDA) / COUNT(*), 2) AS VALOR_M2_CONSTRUCCION_2 "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "  FROM (SELECT TRIM(F.PARDES_DESCRI) AS DISTRITO, "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "               TRIM(HIPMAE_NUMSOL) AS NUMERO_SOLICITUD, "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "               DECODE(C.EVATAS_TIPMON, 1, (C.EVATAS_VALTER_INM+C.EVATAS_VALTER_ES1+C.EVATAS_VALTER_ES2+C.EVATAS_VALTER_DEP), (C.EVATAS_VALTER_INM+C.EVATAS_VALTER_ES1+C.EVATAS_VALTER_ES2+C.EVATAS_VALTER_DEP)*C.EVATAS_TIPCAM) AS VALOR_TERRENO, "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "               C.EVATAS_ARETER_INM+C.EVATAS_ARETER_ES1+C.EVATAS_ARETER_ES2+C.EVATAS_ARETER_DEP AS AREA_TERRENO, "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "               DECODE(C.EVATAS_TIPMON, 1, (C.EVATAS_VALTER_INM+C.EVATAS_VALTER_ES1+C.EVATAS_VALTER_ES2+C.EVATAS_VALTER_DEP), (C.EVATAS_VALTER_INM+C.EVATAS_VALTER_ES1+C.EVATAS_VALTER_ES2+C.EVATAS_VALTER_DEP)*C.EVATAS_TIPCAM) / (C.EVATAS_ARETER_INM+C.EVATAS_ARETER_ES1+C.EVATAS_ARETER_ES2+C.EVATAS_ARETER_DEP) AS VALOR_M2_TERRENO_1, "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "               DECODE(C.EVATAS_TIPMON, 1, (C.EVATAS_VALTER_INM+C.EVATAS_VALTER_ES1+C.EVATAS_VALTER_ES2+C.EVATAS_VALTER_DEP), (C.EVATAS_VALTER_INM+C.EVATAS_VALTER_ES1+C.EVATAS_VALTER_ES2+C.EVATAS_VALTER_DEP)*C.EVATAS_TIPCAM) / (C.EVATAS_ARETER_INM+C.EVATAS_ARETER_ES1+C.EVATAS_ARETER_ES2+C.EVATAS_ARETER_DEP) AS VALOR_M2_TERRENO_2, "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "               DECODE(C.EVATAS_TIPMON, 1, (C.EVATAS_VALEDI_INM+C.EVATAS_VALEDI_ES1+C.EVATAS_VALEDI_ES2+C.EVATAS_VALEDI_DEP), (C.EVATAS_VALEDI_INM+C.EVATAS_VALEDI_ES1+C.EVATAS_VALEDI_ES2+C.EVATAS_VALEDI_DEP)*C.EVATAS_TIPCAM) + DECODE(C.EVATAS_TIPMON, 1, (C.EVATAS_VALACO_INM+C.EVATAS_VALACO_ES1+C.EVATAS_VALACO_ES2+C.EVATAS_VALACO_DEP), (C.EVATAS_VALACO_INM+C.EVATAS_VALACO_ES1+C.EVATAS_VALACO_ES2+C.EVATAS_VALACO_DEP)*C.EVATAS_TIPCAM) AS VALOR_CONSTRUCCION, "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "               C.EVATAS_ARECON_INM+C.EVATAS_ARECON_ES1+C.EVATAS_ARECON_ES2+C.EVATAS_ARECON_DEP AS AREA_CONSTRUIDA, "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "               DECODE(C.EVATAS_TIPMON, 1, (C.EVATAS_VALEDI_INM+C.EVATAS_VALEDI_ES1+C.EVATAS_VALEDI_ES2+C.EVATAS_VALEDI_DEP), (C.EVATAS_VALEDI_INM+C.EVATAS_VALEDI_ES1+C.EVATAS_VALEDI_ES2+C.EVATAS_VALEDI_DEP)*C.EVATAS_TIPCAM) + DECODE(C.EVATAS_TIPMON, 1, (C.EVATAS_VALACO_INM+C.EVATAS_VALACO_ES1+C.EVATAS_VALACO_ES2+C.EVATAS_VALACO_DEP), (C.EVATAS_VALACO_INM+C.EVATAS_VALACO_ES1+C.EVATAS_VALACO_ES2+C.EVATAS_VALACO_DEP)*C.EVATAS_TIPCAM) / (C.EVATAS_ARECON_INM+C.EVATAS_ARECON_ES1+C.EVATAS_ARECON_ES2+C.EVATAS_ARECON_DEP)  AS VALOR_M2_CONSTRUCCION_1, "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "               DECODE(C.EVATAS_TIPMON, 1, (C.EVATAS_VALEDI_INM+C.EVATAS_VALEDI_ES1+C.EVATAS_VALEDI_ES2+C.EVATAS_VALEDI_DEP), (C.EVATAS_VALEDI_INM+C.EVATAS_VALEDI_ES1+C.EVATAS_VALEDI_ES2+C.EVATAS_VALEDI_DEP)*C.EVATAS_TIPCAM) / (C.EVATAS_ARECON_INM+C.EVATAS_ARECON_ES1+C.EVATAS_ARECON_ES2+C.EVATAS_ARECON_DEP) AS VALOR_M2_CONSTRUCCION_2 "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "          FROM CRE_HIPMAE A "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "         INNER JOIN TRA_EVATAS C ON C.EVATAS_NUMSOL = A.HIPMAE_NUMSOL "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "         INNER JOIN CRE_SOLINM E ON E.SOLINM_NUMSOL = A.HIPMAE_NUMSOL "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "         INNER JOIN MNT_PARDES F ON F.PARDES_CODGRP = '101' AND F.PARDES_CODITE = E.SOLINM_UBIGEO "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "         WHERE A.HIPMAE_SITUAC = 2 "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "           AND C.EVATAS_FECEVA > " & p_FecIni & " AND C.EVATAS_FECEVA < " & p_FecFin & " "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "        UNION "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "        SELECT TRIM(F.PARDES_DESCRI) AS DISTRITO, "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "               TRIM(C.EVATAS_NUMSOL) AS NUMERO_SOLICITUD, "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "               DECODE(C.EVATAS_TIPMON, 1, (C.EVATAS_VALTER_INM+C.EVATAS_VALTER_ES1+C.EVATAS_VALTER_ES2+C.EVATAS_VALTER_DEP), (C.EVATAS_VALTER_INM+C.EVATAS_VALTER_ES1+C.EVATAS_VALTER_ES2+C.EVATAS_VALTER_DEP)*C.EVATAS_TIPCAM) AS VALOR_TERRENO, "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "               C.EVATAS_ARETER_INM+C.EVATAS_ARETER_ES1+C.EVATAS_ARETER_ES2+C.EVATAS_ARETER_DEP AS AREA_TERRENO, "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "               DECODE(C.EVATAS_TIPMON, 1, (C.EVATAS_VALTER_INM+C.EVATAS_VALTER_ES1+C.EVATAS_VALTER_ES2+C.EVATAS_VALTER_DEP), (C.EVATAS_VALTER_INM+C.EVATAS_VALTER_ES1+C.EVATAS_VALTER_ES2+C.EVATAS_VALTER_DEP)*C.EVATAS_TIPCAM) / (C.EVATAS_ARETER_INM+C.EVATAS_ARETER_ES1+C.EVATAS_ARETER_ES2+C.EVATAS_ARETER_DEP) AS VALOR_M2_TERRENO_1, "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "               DECODE(C.EVATAS_TIPMON, 1, (C.EVATAS_VALTER_INM+C.EVATAS_VALTER_ES1+C.EVATAS_VALTER_ES2+C.EVATAS_VALTER_DEP), (C.EVATAS_VALTER_INM+C.EVATAS_VALTER_ES1+C.EVATAS_VALTER_ES2+C.EVATAS_VALTER_DEP)*C.EVATAS_TIPCAM) / (C.EVATAS_ARETER_INM+C.EVATAS_ARETER_ES1+C.EVATAS_ARETER_ES2+C.EVATAS_ARETER_DEP) AS VALOR_M2_TERRENO_2, "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "               DECODE(C.EVATAS_TIPMON, 1, (C.EVATAS_VALEDI_INM+C.EVATAS_VALEDI_ES1+C.EVATAS_VALEDI_ES2+C.EVATAS_VALEDI_DEP), (C.EVATAS_VALEDI_INM+C.EVATAS_VALEDI_ES1+C.EVATAS_VALEDI_ES2+C.EVATAS_VALEDI_DEP)*C.EVATAS_TIPCAM) + DECODE(C.EVATAS_TIPMON, 1, (C.EVATAS_VALACO_INM+C.EVATAS_VALACO_ES1+C.EVATAS_VALACO_ES2+C.EVATAS_VALACO_DEP), (C.EVATAS_VALACO_INM+C.EVATAS_VALACO_ES1+C.EVATAS_VALACO_ES2+C.EVATAS_VALACO_DEP)*C.EVATAS_TIPCAM) AS VALOR_CONSTRUCCION, "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "               C.EVATAS_ARECON_INM+C.EVATAS_ARECON_ES1+C.EVATAS_ARECON_ES2+C.EVATAS_ARECON_DEP AS AREA_CONSTRUIDA, "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "               DECODE(C.EVATAS_TIPMON, 1, (C.EVATAS_VALEDI_INM+C.EVATAS_VALEDI_ES1+C.EVATAS_VALEDI_ES2+C.EVATAS_VALEDI_DEP), (C.EVATAS_VALEDI_INM+C.EVATAS_VALEDI_ES1+C.EVATAS_VALEDI_ES2+C.EVATAS_VALEDI_DEP)*C.EVATAS_TIPCAM) + DECODE(C.EVATAS_TIPMON, 1, (C.EVATAS_VALACO_INM+C.EVATAS_VALACO_ES1+C.EVATAS_VALACO_ES2+C.EVATAS_VALACO_DEP), (C.EVATAS_VALACO_INM+C.EVATAS_VALACO_ES1+C.EVATAS_VALACO_ES2+C.EVATAS_VALACO_DEP)*C.EVATAS_TIPCAM) / (C.EVATAS_ARECON_INM+C.EVATAS_ARECON_ES1+C.EVATAS_ARECON_ES2+C.EVATAS_ARECON_DEP)  AS VALOR_M2_CONSTRUCCION_1, "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "               DECODE(C.EVATAS_TIPMON, 1, (C.EVATAS_VALEDI_INM+C.EVATAS_VALEDI_ES1+C.EVATAS_VALEDI_ES2+C.EVATAS_VALEDI_DEP), (C.EVATAS_VALEDI_INM+C.EVATAS_VALEDI_ES1+C.EVATAS_VALEDI_ES2+C.EVATAS_VALEDI_DEP)*C.EVATAS_TIPCAM) / (C.EVATAS_ARECON_INM+C.EVATAS_ARECON_ES1+C.EVATAS_ARECON_ES2+C.EVATAS_ARECON_DEP) AS VALOR_M2_CONSTRUCCION_2 "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "          FROM HIS_EVATAS C "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "         INNER JOIN CRE_SOLINM E ON E.SOLINM_NUMSOL = C.EVATAS_NUMSOL "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "         INNER JOIN MNT_PARDES F ON F.PARDES_CODGRP = '101' AND F.PARDES_CODITE = E.SOLINM_UBIGEO "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & "         WHERE C.EVATAS_FECEVA > " & p_FecIni & " AND C.EVATAS_FECEVA < " & p_FecFin & " "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & " GROUP BY DISTRITO "
'   ff_Obtiene_Tasaciones2 = ff_Obtiene_Tasaciones2 & " ORDER BY DISTRITO "
'End Function

Private Function fs_ValidaOperacion(ByVal p_NumOpe As String, ByVal p_AnoPro As String, ByVal p_MesPro As String, ByVal p_FecIni As String) As Boolean
Dim r_rst_UltGar     As ADODB.Recordset
Dim r_str_Cadena     As String
Dim r_str_UltEvaMes  As String
Dim r_str_UltEvaAno  As String

   fs_ValidaOperacion = False
   
   'Valida que la operacion no tenga una nueva tasacion registrada en tabla historica, menor a 1 año
   r_str_Cadena = ""
   r_str_Cadena = r_str_Cadena & "SELECT COUNT(*) AS CONTADOR "
   r_str_Cadena = r_str_Cadena & "  FROM HIS_EVATAS "
   r_str_Cadena = r_str_Cadena & " WHERE EVATAS_NUMSOL = (SELECT HIPMAE_NUMSOL FROM CRE_HIPMAE WHERE HIPMAE_NUMOPE = '" & p_NumOpe & "') "
   r_str_Cadena = r_str_Cadena & "   AND EVATAS_FECEVA > " & p_FecIni & " "
   
   If Not gf_EjecutaSQL(r_str_Cadena, r_rst_UltGar, 3) Then
      r_rst_UltGar.Close
      Set r_rst_UltGar = Nothing
      Exit Function
   End If
   
   r_rst_UltGar.MoveFirst
   If r_rst_UltGar!CONTADOR > 0 Then
      r_rst_UltGar.Close
      Set r_rst_UltGar = Nothing
      Exit Function
   End If
   
   r_rst_UltGar.Close
   Set r_rst_UltGar = Nothing
   
   'Valida que la operacion no este registrada en proceso de actualizacion de garantias, menor a 1 año
   r_str_Cadena = ""
   r_str_Cadena = r_str_Cadena & "SELECT * "
   r_str_Cadena = r_str_Cadena & "  FROM (SELECT DISTINCT ACTGAR_ANOPRO, ACTGAR_MESPRO "
   r_str_Cadena = r_str_Cadena & "          FROM CRE_ACTGAR "
   r_str_Cadena = r_str_Cadena & "         WHERE ACTGAR_NUMOPE = '" & p_NumOpe & "' "
   r_str_Cadena = r_str_Cadena & "         ORDER BY ACTGAR_ANOPRO DESC, ACTGAR_MESPRO DESC) "
   r_str_Cadena = r_str_Cadena & " WHERE ROWNUM < 2 "
   r_str_Cadena = r_str_Cadena & " ORDER BY ACTGAR_ANOPRO "
   
   If Not gf_EjecutaSQL(r_str_Cadena, r_rst_UltGar, 3) Then
      r_rst_UltGar.Close
      Set r_rst_UltGar = Nothing
      Exit Function
   End If
   
   If r_rst_UltGar.BOF And r_rst_UltGar.EOF Then
      fs_ValidaOperacion = True
      r_rst_UltGar.Close
      Set r_rst_UltGar = Nothing
      Exit Function
   End If
   
   r_rst_UltGar.MoveFirst
   r_str_UltEvaMes = r_rst_UltGar!ACTGAR_MESPRO
   r_str_UltEvaAno = r_rst_UltGar!ACTGAR_ANOPRO
   
   r_rst_UltGar.Close
   Set r_rst_UltGar = Nothing
   
   If CInt(p_AnoPro) - CInt(r_str_UltEvaAno) - 1 = 0 Then
      If CInt(r_str_UltEvaMes) <= CInt(p_MesPro) Then
         fs_ValidaOperacion = True
         Exit Function
      End If
   ElseIf CInt(p_AnoPro) - CInt(r_str_UltEvaAno) - 1 > 0 Then
      fs_ValidaOperacion = True
      Exit Function
   End If
End Function

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_Contad     As Integer
Dim r_int_NroFil     As Integer

   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      'Titulo
      .Cells(2, 1) = "REPORTE DE ACTUALIZACION DE GARANTIAS - PERIODO : " & Trim(cmb_PerMes.Text) & " / " & Trim(ipp_PerAno.Text)
      .Range(.Cells(2, 1), .Cells(2, 29)).Merge
      .Range("A2:AD2").HorizontalAlignment = xlHAlignCenter
      
      r_int_NroFil = 4
      .Columns("A").ColumnWidth = 5:    .Columns("A").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 1) = "ITEM"
      .Columns("B").ColumnWidth = 14:   .Columns("B").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 2) = "OPERACION"
      .Columns("C").ColumnWidth = 14:   .Columns("C").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 3) = "PRODUCTO"
      .Columns("D").ColumnWidth = 15:   .Columns("D").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 4) = "DOCUMENTO IDENTIFICACION"
      .Columns("E").ColumnWidth = 40:   .Columns("E").HorizontalAlignment = xlHAlignLeft:     .Cells(r_int_NroFil, 5) = "NOMBRE DEL CLIENTE"
      .Columns("F").ColumnWidth = 22:   .Columns("F").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 6) = "TIPO DE MONEDA"
      .Columns("G").ColumnWidth = 16:   .Columns("G").HorizontalAlignment = xlHAlignRight:    .Cells(r_int_NroFil, 7) = "SALDO CAPITAL"
      .Columns("H").ColumnWidth = 20:   .Columns("H").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 8) = "DISTRITO"
      .Columns("I").ColumnWidth = 14:   .Columns("I").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 9) = "FECHA DE TASACION"
      .Columns("J").ColumnWidth = 12:   .Columns("J").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 10) = "MES DE TASACION"
      .Columns("K").ColumnWidth = 12:   .Columns("K").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 11) = "AÑO DE TASACION"
      .Columns("L").ColumnWidth = 13:   .Columns("L").HorizontalAlignment = xlHAlignRight:    .Cells(r_int_NroFil, 12) = "AREA CONSTRUIDA"
      .Columns("M").ColumnWidth = 12:   .Columns("M").HorizontalAlignment = xlHAlignRight:    .Cells(r_int_NroFil, 13) = "AREA DEL TERRENO"
      .Columns("N").ColumnWidth = 13:   .Columns("N").HorizontalAlignment = xlHAlignRight:    .Cells(r_int_NroFil, 14) = "VALOR COMERCIAL"
      .Columns("O").ColumnWidth = 15:   .Columns("O").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 15) = "AÑO CONSTRUCCION"
      .Columns("P").ColumnWidth = 15:   .Columns("P").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 16) = "MATERIAL CONSTRUCCION"
      .Columns("Q").ColumnWidth = 15:   .Columns("Q").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 17) = "ESTADO CONSERVACION"
      .Columns("R").ColumnWidth = 13:   .Columns("R").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 18) = "ANTIGUEDAD (AÑOS)"
      .Columns("S").ColumnWidth = 14:   .Columns("S").HorizontalAlignment = xlHAlignRight:    .Cells(r_int_NroFil, 19) = "DEPRECIACION (%)"
      .Columns("T").ColumnWidth = 14:   .Columns("T").HorizontalAlignment = xlHAlignRight:    .Cells(r_int_NroFil, 20) = "VALOR M2 TERRENO"
      .Columns("U").ColumnWidth = 15:   .Columns("U").HorizontalAlignment = xlHAlignRight:    .Cells(r_int_NroFil, 21) = "VALOR M2 CONSTRUCCION"
      .Columns("V").ColumnWidth = 15:   .Columns("V").HorizontalAlignment = xlHAlignRight:    .Cells(r_int_NroFil, 22) = "VALOR ACTUAL"
      .Columns("W").ColumnWidth = 15:   .Columns("W").HorizontalAlignment = xlHAlignRight:    .Cells(r_int_NroFil, 23) = "VALOR ACTUALIZADO"
      .Columns("X").ColumnWidth = 16:   .Columns("X").HorizontalAlignment = xlHAlignRight:    .Cells(r_int_NroFil, 24) = "VALACT/VALCOM (%)"
      .Columns("Y").ColumnWidth = 16:   .Columns("Y").HorizontalAlignment = xlHAlignRight:    .Cells(r_int_NroFil, 25) = "CORREGIDO VALOR TERRENO"
      .Columns("Z").ColumnWidth = 16:   .Columns("Z").HorizontalAlignment = xlHAlignRight:    .Cells(r_int_NroFil, 26) = "CORREGIDO VALOR CONSTR"
      .Columns("AA").ColumnWidth = 16:  .Columns("AA").HorizontalAlignment = xlHAlignRight:   .Cells(r_int_NroFil, 27) = "CORREGIDO VALOR ACTUALIZ"
      .Columns("AB").ColumnWidth = 16:  .Columns("AB").HorizontalAlignment = xlHAlignRight:   .Cells(r_int_NroFil, 28) = "CORREGIDO ACT/COM (%)"
      .Columns("AC").ColumnWidth = 10:  .Columns("AC").HorizontalAlignment = xlHAlignRight:   .Cells(r_int_NroFil, 29) = "LTV (%)"
      
      .Range(.Cells(1, 1), .Cells(4, 29)).Font.Bold = True
      .Range(.Cells(4, 1), .Cells(4, 29)).WrapText = True
      .Range(.Cells(4, 1), .Cells(4, 29)).VerticalAlignment = xlCenter
      .Range(.Cells(4, 1), .Cells(4, 29)).HorizontalAlignment = xlCenter
      .Range(.Cells(4, 1), .Cells(4, 29)).Interior.Color = RGB(146, 208, 80)
            
      For r_int_Contad = 1 To grd_Listad.Rows - 1
         r_int_NroFil = r_int_NroFil + 1
         .Cells(r_int_NroFil, 1) = grd_Listad.TextMatrix(r_int_Contad, 0)
         .Cells(r_int_NroFil, 2) = grd_Listad.TextMatrix(r_int_Contad, 1)
         .Cells(r_int_NroFil, 3) = grd_Listad.TextMatrix(r_int_Contad, 2)
         .Cells(r_int_NroFil, 4) = grd_Listad.TextMatrix(r_int_Contad, 3)
         .Cells(r_int_NroFil, 5) = grd_Listad.TextMatrix(r_int_Contad, 4)
         .Cells(r_int_NroFil, 6) = grd_Listad.TextMatrix(r_int_Contad, 5)
         .Cells(r_int_NroFil, 7) = Format(grd_Listad.TextMatrix(r_int_Contad, 6), "###,###,##0.00")
         .Cells(r_int_NroFil, 8) = grd_Listad.TextMatrix(r_int_Contad, 7)
         .Cells(r_int_NroFil, 9) = "'" & grd_Listad.TextMatrix(r_int_Contad, 8)
         .Cells(r_int_NroFil, 10) = grd_Listad.TextMatrix(r_int_Contad, 9)
         .Cells(r_int_NroFil, 11) = grd_Listad.TextMatrix(r_int_Contad, 10)
         .Cells(r_int_NroFil, 12) = Format(grd_Listad.TextMatrix(r_int_Contad, 11), "###,##0.00")
         .Cells(r_int_NroFil, 13) = Format(grd_Listad.TextMatrix(r_int_Contad, 12), "###,##0.00")
         .Cells(r_int_NroFil, 14) = Format(grd_Listad.TextMatrix(r_int_Contad, 13), "###,##0.00")
         .Cells(r_int_NroFil, 15) = grd_Listad.TextMatrix(r_int_Contad, 14)
         .Cells(r_int_NroFil, 16) = grd_Listad.TextMatrix(r_int_Contad, 15)
         .Cells(r_int_NroFil, 17) = grd_Listad.TextMatrix(r_int_Contad, 16)
         .Cells(r_int_NroFil, 18) = grd_Listad.TextMatrix(r_int_Contad, 17)
         .Cells(r_int_NroFil, 19) = grd_Listad.TextMatrix(r_int_Contad, 18)
         .Cells(r_int_NroFil, 20) = Format(grd_Listad.TextMatrix(r_int_Contad, 19), "###,##0.00")
         .Cells(r_int_NroFil, 21) = Format(grd_Listad.TextMatrix(r_int_Contad, 20), "###,##0.00")
         .Cells(r_int_NroFil, 22) = Format(grd_Listad.TextMatrix(r_int_Contad, 21), "###,##0.00")
         .Cells(r_int_NroFil, 23) = Format(grd_Listad.TextMatrix(r_int_Contad, 22), "###,##0.00")
         .Cells(r_int_NroFil, 24) = Format(grd_Listad.TextMatrix(r_int_Contad, 23), "###,##0.00")
         .Cells(r_int_NroFil, 25) = Format(grd_Listad.TextMatrix(r_int_Contad, 24), "###,##0.00")
         .Cells(r_int_NroFil, 26) = Format(grd_Listad.TextMatrix(r_int_Contad, 25), "###,##0.00")
         .Cells(r_int_NroFil, 27) = Format(grd_Listad.TextMatrix(r_int_Contad, 26), "###,##0.00")
         .Cells(r_int_NroFil, 28) = Format(grd_Listad.TextMatrix(r_int_Contad, 27), "###,##0.00")
         .Cells(r_int_NroFil, 29) = Format(grd_Listad.TextMatrix(r_int_Contad, 28), "###,##0.00")
         
         .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 29)).Borders(xlEdgeTop).LineStyle = xlContinuous
      Next r_int_Contad
      
      .Range(.Cells(4, 1), .Cells(r_int_NroFil, 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 2), .Cells(r_int_NroFil, 2)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 3), .Cells(r_int_NroFil, 3)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 4), .Cells(r_int_NroFil, 4)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 5), .Cells(r_int_NroFil, 5)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 6), .Cells(r_int_NroFil, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 7), .Cells(r_int_NroFil, 7)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 8), .Cells(r_int_NroFil, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 9), .Cells(r_int_NroFil, 9)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 10), .Cells(r_int_NroFil, 10)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 11), .Cells(r_int_NroFil, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 12), .Cells(r_int_NroFil, 12)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 13), .Cells(r_int_NroFil, 13)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 14), .Cells(r_int_NroFil, 14)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 15), .Cells(r_int_NroFil, 15)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 16), .Cells(r_int_NroFil, 16)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 17), .Cells(r_int_NroFil, 17)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 18), .Cells(r_int_NroFil, 18)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 19), .Cells(r_int_NroFil, 19)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 20), .Cells(r_int_NroFil, 20)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 21), .Cells(r_int_NroFil, 21)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 22), .Cells(r_int_NroFil, 22)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 23), .Cells(r_int_NroFil, 23)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 24), .Cells(r_int_NroFil, 24)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 25), .Cells(r_int_NroFil, 25)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 26), .Cells(r_int_NroFil, 26)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 27), .Cells(r_int_NroFil, 27)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 28), .Cells(r_int_NroFil, 28)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 29), .Cells(r_int_NroFil, 29)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 30), .Cells(r_int_NroFil, 30)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      
      .Range(.Cells(4, 1), .Cells(4, 29)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(r_int_NroFil + 1, 1), .Cells(r_int_NroFil + 1, 29)).Borders(xlEdgeTop).LineStyle = xlContinuous

      .Range(.Cells(5, 12), .Cells(r_int_NroFil, 14)).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      
      .Range(.Cells(5, 19), .Cells(r_int_NroFil, 29)).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Cells(1, 1).Select
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_PerMes.ListIndex > -1 Then
         Call gs_SetFocus(ipp_PerAno)
      End If
   End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub
