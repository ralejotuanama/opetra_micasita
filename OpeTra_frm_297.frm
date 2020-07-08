VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_RptFia_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2850
   ClientLeft      =   3720
   ClientTop       =   5325
   ClientWidth     =   7215
   Icon            =   "OpeTra_frm_297.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2865
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7215
      _Version        =   65536
      _ExtentX        =   12726
      _ExtentY        =   5054
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
         TabIndex        =   6
         Top             =   780
         Width           =   7095
         _Version        =   65536
         _ExtentX        =   12515
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
            Left            =   30
            Picture         =   "OpeTra_frm_297.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   6480
            Picture         =   "OpeTra_frm_297.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   7
         Top             =   60
         Width           =   7095
         _Version        =   65536
         _ExtentX        =   12515
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
            Left            =   690
            TabIndex        =   8
            Top             =   30
            Width           =   6000
            _Version        =   65536
            _ExtentX        =   10583
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Reportes - Carta Fianza"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   315
            Left            =   690
            TabIndex        =   9
            Top             =   330
            Width           =   6000
            _Version        =   65536
            _ExtentX        =   10583
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Saldos de Garantias"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
            Picture         =   "OpeTra_frm_297.frx":0758
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1335
         Left            =   60
         TabIndex        =   10
         Top             =   1470
         Width           =   7095
         _Version        =   65536
         _ExtentX        =   12515
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
         Begin VB.ComboBox cmb_Permes 
            Height          =   315
            Left            =   990
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   180
            Width           =   5535
         End
         Begin VB.CheckBox Chk_FecAct 
            Caption         =   "A la Fecha"
            Height          =   285
            Left            =   990
            TabIndex        =   2
            Top             =   930
            Width           =   1995
         End
         Begin EditLib.fpDoubleSingle ipp_PerAno 
            Height          =   315
            Left            =   990
            TabIndex        =   1
            Top             =   540
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
            DecimalPlaces   =   0
            DecimalPoint    =   "."
            FixedPoint      =   0   'False
            LeadZero        =   0
            MaxValue        =   "9999"
            MinValue        =   "1900"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   0   'False
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
            Caption         =   "Mes:"
            Height          =   255
            Left            =   150
            TabIndex        =   12
            Top             =   210
            Width           =   675
         End
         Begin VB.Label Label5 
            Caption         =   "Año:"
            Height          =   255
            Left            =   150
            TabIndex        =   11
            Top             =   570
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frm_RptFia_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_Permes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_Permes.ListIndex > -1 Then
         Call gs_SetFocus(ipp_PerAno)
      End If
   End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_PerAno.Text > 2006 Then
         Call gs_SetFocus(Chk_FecAct)
      End If
   End If
End Sub

Private Sub Chk_FecAct_Click()
   If Chk_FecAct.Value = 1 Then
      cmb_Permes.ListIndex = -1
      cmb_Permes.Enabled = False
      ipp_PerAno.Value = 0
      ipp_PerAno.Enabled = False
   ElseIf Chk_FecAct.Value = 0 Then
      cmb_Permes.Enabled = True
      ipp_PerAno.Enabled = True
   End If
End Sub

Private Sub cmd_ExpExc_Click()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
   
   If Chk_FecAct.Value = 0 Then
      If cmb_Permes.ListIndex = -1 Then
         MsgBox "Debe seleccionar el mes de proceso.", vbInformation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Permes)
         Exit Sub
      End If
      If ipp_PerAno.Text < 2007 Then
         MsgBox "Debe ingresar el año de proceso.", vbInformation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_PerAno)
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Está seguro de exportar el reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   If Chk_FecAct.Value = 1 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT TRIM(B.PARDES_DESCRI) AS GARANTIA, TRIM(C.PARDES_DESCRI) AS MONEDA, "
      g_str_Parame = g_str_Parame & "       TRIM(D.PARDES_DESCRI) AS BANCO, COUNT(*) AS NUMERO, ROUND(SUM(HIPMAE_MTOGAR),2) AS MONTO "
      g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE A "
      g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = 241 AND B.PARDES_CODITE = A.HIPMAE_TIPGAR "
      g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 204 AND C.PARDES_CODITE = A.HIPMAE_MONGAR "
      g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 505 AND D.PARDES_CODITE = A.HIPMAE_BCOGAR "
      g_str_Parame = g_str_Parame & " WHERE A.HIPMAE_SITUAC = 2 "
      g_str_Parame = g_str_Parame & "   AND A.HIPMAE_TIPGAR NOT IN (1,2,5) "
      g_str_Parame = g_str_Parame & "GROUP BY B.PARDES_DESCRI, C.PARDES_DESCRI, D.PARDES_DESCRI "
   Else
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT TRIM(B.PARDES_DESCRI) AS GARANTIA, TRIM(C.PARDES_DESCRI) AS MONEDA, "
      g_str_Parame = g_str_Parame & "       TRIM(D.PARDES_DESCRI) AS BANCO, COUNT(*) AS NUMERO, ROUND(SUM(HIPCIE_MTOGAR),2) AS MONTO "
      g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE A "
      g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = 241 AND B.PARDES_CODITE = A.HIPCIE_TIPGAR "
      g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES C ON C.PARDES_CODGRP = 204 AND C.PARDES_CODITE = DECODE(A.HIPCIE_MONGAR, 0, 100, A.HIPCIE_MONGAR)"
      g_str_Parame = g_str_Parame & " INNER JOIN CRE_HIPMAE F ON F.HIPMAE_NUMOPE = A.HIPCIE_NUMOPE"
      g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 505 AND D.PARDES_CODITE = F.HIPMAE_BCOGAR "
      g_str_Parame = g_str_Parame & " WHERE A.HIPCIE_PERMES = " & cmb_Permes.ItemData(cmb_Permes.ListIndex) & " "
      g_str_Parame = g_str_Parame & "   AND A.HIPCIE_PERANO = " & Format(ipp_PerAno.Text, "0000") & " "
      g_str_Parame = g_str_Parame & "   AND A.HIPCIE_TIPGAR NOT IN (1,2,5)"
      g_str_Parame = g_str_Parame & "GROUP BY B.PARDES_DESCRI, C.PARDES_DESCRI, D.PARDES_DESCRI"
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron operaciones registradas.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      If Chk_FecAct.Value = 1 Then
         .Cells(2, 1) = "REPORTE DE SALDOS DE GARANTIA AL " & Format(date, "DD/MM/YYYY")
      Else
         .Cells(2, 1) = "REPORTE DE SALDOS DE GARANTIA A " & TRIM(cmb_Permes.Text) & " DEL " & TRIM(ipp_PerAno.Text)
      End If
      
      .Range(.Cells(2, 1), .Cells(2, 6)).Merge
      .Cells(4, 1) = "ITEM"
      .Cells(4, 2) = "GARANTIA"
      .Cells(4, 3) = "TIPO DE MONEDA"
      .Cells(4, 4) = "BANCO"
      .Cells(4, 5) = "NUMERO"
      .Cells(4, 6) = "MONTO"
      .Range(.Cells(1, 1), .Cells(4, 6)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(4, 6)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").ColumnWidth = 6
      .Columns("B").ColumnWidth = 30
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 25
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 30
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 12
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 19
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 5
   
   Do While Not g_rst_Princi.EOF
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 4
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = TRIM(g_rst_Princi!GARANTIA)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = TRIM(g_rst_Princi!moneda)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = TRIM(g_rst_Princi!BANCO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = TRIM(g_rst_Princi!NUMERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = Format(g_rst_Princi!MONTO, "###,###,##0.00")
      
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_Permes)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Chk_FecAct.Value = 0
   ipp_PerAno.Text = Year(date)
   Call moddat_gs_Carga_LisIte_Combo(cmb_Permes, 1, "033")
End Sub

