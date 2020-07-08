VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Ges_TecPro_13 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7260
   Icon            =   "OpeTra_frm_835.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   3045
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7305
      _Version        =   65536
      _ExtentX        =   12885
      _ExtentY        =   5371
      _StockProps     =   15
      BackColor       =   14215660
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
         Height          =   3015
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   7275
         _Version        =   65536
         _ExtentX        =   12832
         _ExtentY        =   5318
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
            Left            =   60
            TabIndex        =   8
            Top             =   60
            Width           =   7155
            _Version        =   65536
            _ExtentX        =   12621
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
               TabIndex        =   14
               Top             =   60
               Width           =   4215
               _Version        =   65536
               _ExtentX        =   7435
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Gestión de Crédito Hipotecario"
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
            Begin Threed.SSPanel SSPanel15 
               Height          =   315
               Left            =   630
               TabIndex        =   15
               Top             =   360
               Width           =   5505
               _Version        =   65536
               _ExtentX        =   9710
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Reporte de Cartas Fianza, Adendas y Cartas Seriedad Oferta"
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
               Left            =   60
               Picture         =   "OpeTra_frm_835.frx":000C
               Top             =   90
               Width           =   480
            End
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   645
            Left            =   60
            TabIndex        =   9
            Top             =   780
            Width           =   7155
            _Version        =   65536
            _ExtentX        =   12621
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
               Left            =   6540
               Picture         =   "OpeTra_frm_835.frx":0316
               Style           =   1  'Graphical
               TabIndex        =   5
               ToolTipText     =   "Salir de la Opción"
               Top             =   30
               Width           =   585
            End
            Begin VB.CommandButton cmd_ExpExc 
               Height          =   585
               Left            =   30
               Picture         =   "OpeTra_frm_835.frx":0758
               Style           =   1  'Graphical
               TabIndex        =   4
               ToolTipText     =   "Exportar a Excel"
               Top             =   30
               Width           =   585
            End
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   1485
            Left            =   60
            TabIndex        =   10
            Top             =   1470
            Width           =   7155
            _Version        =   65536
            _ExtentX        =   12621
            _ExtentY        =   2619
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
            Begin VB.CheckBox Chk_FecAct 
               Caption         =   "A la Fecha"
               Height          =   285
               Left            =   1200
               TabIndex        =   3
               Top             =   1200
               Width           =   1995
            End
            Begin VB.ComboBox cmb_TipRep 
               Height          =   315
               ItemData        =   "OpeTra_frm_835.frx":0A62
               Left            =   1200
               List            =   "OpeTra_frm_835.frx":0A64
               Style           =   2  'Dropdown List
               TabIndex        =   0
               Top             =   120
               Width           =   5775
            End
            Begin VB.ComboBox cmb_PerMes 
               Height          =   315
               ItemData        =   "OpeTra_frm_835.frx":0A66
               Left            =   1200
               List            =   "OpeTra_frm_835.frx":0A68
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   480
               Width           =   2535
            End
            Begin EditLib.fpLongInteger ipp_PerAno 
               Height          =   315
               Left            =   1200
               TabIndex        =   2
               Top             =   840
               Width           =   855
               _Version        =   196608
               _ExtentX        =   1508
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
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Reporte:"
               Height          =   195
               Left            =   120
               TabIndex        =   13
               Top             =   150
               Width           =   975
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Año:"
               Height          =   195
               Left            =   120
               TabIndex        =   12
               Top             =   870
               Width           =   330
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Periodo:"
               Height          =   195
               Left            =   120
               TabIndex        =   11
               Top             =   510
               Width           =   585
            End
         End
      End
   End
End
Attribute VB_Name = "frm_Ges_TecPro_13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_str_FecPer  As String

Private Sub Chk_FecAct_Click()
   If Chk_FecAct.Value = 1 Then
      cmb_PerMes.ListIndex = -1
      cmb_PerMes.Enabled = False
      ipp_PerAno.Value = 0
      ipp_PerAno.Enabled = False
   ElseIf Chk_FecAct.Value = 0 Then
      cmb_PerMes.Enabled = True
      ipp_PerAno.Enabled = True
   End If
   Call gs_SetFocus(cmb_PerMes)
End Sub

Private Sub Chk_FecAct_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_ExpExc)
  End If
End Sub

Private Sub cmb_PerMes_Click()
   Call gs_SetFocus(ipp_PerAno)
End Sub

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PerAno)
  End If
End Sub

Private Sub cmb_TipRep_Click()
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 3 Then
      Chk_FecAct.Enabled = False
   Else
      Chk_FecAct.Enabled = True
   End If
   Call gs_SetFocus(cmb_PerMes)
End Sub

Private Sub cmb_TipRep_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_PerMes)
  End If
End Sub

Private Sub cmd_ExpExc_Click()
Dim r_str_PerMes  As String
Dim r_str_PerAno  As String
   
   If cmb_TipRep.ListIndex = -1 Then
         MsgBox "Debe seleccionar Tipo de Reporte.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipRep)
         Exit Sub
   End If
      
   If Chk_FecAct.Value = 0 Then
      If cmb_PerMes.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Mes.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_PerMes)
         Exit Sub
      End If
      If ipp_PerAno.Text = 0 Then
         MsgBox "Debe seleccionar un Año.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_PerAno)
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   cmd_ExpExc.Enabled = False
   
   If Chk_FecAct.Value = 0 Then
      r_str_PerMes = Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00")
      r_str_PerAno = Format(ipp_PerAno.Text, "0000")
      l_str_FecPer = r_str_PerAno & r_str_PerMes & Format(ff_Ultimo_Dia_Mes(CInt(r_str_PerMes), CInt(r_str_PerAno)), "00")
   Else
      l_str_FecPer = Format(Now, "yyyymmdd")
   End If
   
   Select Case cmb_TipRep.ItemData(cmb_TipRep.ListIndex)
      Case 1: Call fs_GenExc_Resumen
      Case 2: Call fs_GenExc_Detalle
      Case 3: Call fs_GenExc_Riesgo
      Case 4: Call fs_GenExc_DocEle
   End Select

   cmd_ExpExc.Enabled = True
   Screen.MousePointer = 0
End Sub

Private Sub fs_GenExc_Resumen()
Dim r_obj_Excel         As Excel.Application
Dim r_str_ParAux        As String
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
Dim r_int_NroFil        As Integer
Dim r_int_NoFlLi        As Integer
   
   r_int_NroFil = 4
      
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
   
      .Cells(1, 2) = "REPORTE DE ENTIDADES TÉCNICAS "
      .Range(.Cells(1, 2), .Cells(1, 19)).Merge
      .Range(.Cells(1, 2), .Cells(1, 19)).Font.Bold = True
      .Range(.Cells(1, 2), .Cells(1, 19)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 2), .Cells(1, 19)).Font.Size = 14
      
      .Cells(r_int_NroFil, 2) = "TIPO DOC. - DOCUMENTO"
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 2)).Merge
      .Cells(r_int_NroFil, 3) = "RAZÓN SOCIAL"
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil + 1, 3)).Merge
      .Cells(r_int_NroFil, 4) = "LINEA ASIGNADA"
      .Range(.Cells(r_int_NroFil, 4), .Cells(r_int_NroFil + 1, 4)).Merge
      
      .Cells(r_int_NroFil - 1, 5) = "CREDITOS INDIRECTOS"
      .Range(.Cells(r_int_NroFil - 1, 5), .Cells(r_int_NroFil - 1, 12)).Merge
      .Range(.Cells(r_int_NroFil - 1, 5), .Cells(r_int_NroFil - 1, 12)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_NroFil - 1, 5), .Cells(r_int_NroFil - 1, 12)).Font.Bold = True
      .Range(.Cells(r_int_NroFil - 1, 5), .Cells(r_int_NroFil - 1, 12)).HorizontalAlignment = xlHAlignCenter
      
      .Cells(r_int_NroFil, 5) = "NRO. CARTA FIANZA"
      .Range(.Cells(r_int_NroFil, 5), .Cells(r_int_NroFil + 1, 5)).Merge
      .Cells(r_int_NroFil, 6) = "NRO. ADENDA"
      .Range(.Cells(r_int_NroFil, 6), .Cells(r_int_NroFil + 1, 6)).Merge
      .Cells(r_int_NroFil, 7) = "NRO. C. SERIEDAD OFERTA"
      .Range(.Cells(r_int_NroFil, 7), .Cells(r_int_NroFil + 1, 7)).Merge
      .Cells(r_int_NroFil, 8) = "GARANTIA"
      .Range(.Cells(r_int_NroFil, 8), .Cells(r_int_NroFil, 9)).Merge
      .Cells(r_int_NroFil + 1, 8) = "LIQUIDA"
      .Cells(r_int_NroFil + 1, 9) = "HIPOTECARIA"
      
      .Cells(r_int_NroFil, 10) = "LINEA UTILIZADA CF"
      .Range(.Cells(r_int_NroFil, 10), .Cells(r_int_NroFil + 1, 10)).Merge
      .Cells(r_int_NroFil, 11) = "LINEA UTILIZADA AD"
      .Range(.Cells(r_int_NroFil, 11), .Cells(r_int_NroFil + 1, 11)).Merge
      .Cells(r_int_NroFil, 12) = "LINEA UTILIZADA CSO"
      .Range(.Cells(r_int_NroFil, 12), .Cells(r_int_NroFil + 1, 12)).Merge
      
      .Cells(r_int_NroFil - 1, 13) = "CREDITOS DIRECTOS"
      .Range(.Cells(r_int_NroFil - 1, 13), .Cells(r_int_NroFil - 1, 17)).Merge
      .Range(.Cells(r_int_NroFil - 1, 13), .Cells(r_int_NroFil - 1, 17)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_NroFil - 1, 13), .Cells(r_int_NroFil - 1, 17)).Font.Bold = True
      .Range(.Cells(r_int_NroFil - 1, 13), .Cells(r_int_NroFil - 1, 17)).HorizontalAlignment = xlHAlignCenter
      
      .Cells(r_int_NroFil, 13) = "NRO. LC"
      .Range(.Cells(r_int_NroFil, 13), .Cells(r_int_NroFil + 1, 13)).Merge
      .Cells(r_int_NroFil, 14) = "NRO. CP"
      .Range(.Cells(r_int_NroFil, 14), .Cells(r_int_NroFil + 1, 14)).Merge
      .Cells(r_int_NroFil, 15) = "GARANTIA"
      .Range(.Cells(r_int_NroFil, 15), .Cells(r_int_NroFil + 1, 15)).Merge
      .Cells(r_int_NroFil, 16) = "LINEA UTILIZADA LC"
      .Range(.Cells(r_int_NroFil, 16), .Cells(r_int_NroFil + 1, 16)).Merge
      .Cells(r_int_NroFil, 17) = "LINEA UTILIZADA CP"
      .Range(.Cells(r_int_NroFil, 17), .Cells(r_int_NroFil + 1, 17)).Merge
      
      .Cells(r_int_NroFil, 18) = "SALDO"
      .Range(.Cells(r_int_NroFil, 18), .Cells(r_int_NroFil + 1, 18)).Merge
      .Cells(r_int_NroFil, 19) = "USUARIO"
      .Range(.Cells(r_int_NroFil, 19), .Cells(r_int_NroFil + 1, 19)).Merge
      
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 19)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 19)).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 19)).HorizontalAlignment = xlHAlignCenter
      '
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 16
      .Columns("C").ColumnWidth = 60
      .Columns("C").HorizontalAlignment = xlHAlignLeft
      .Columns("D").ColumnWidth = 15
      .Columns("D").NumberFormat = "###,###,###,##0.00"
      .Columns("D").HorizontalAlignment = xlHAlignRight
      .Columns("E").ColumnWidth = 16
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("E").NumberFormat = "#,##0"
      .Columns("F").ColumnWidth = 16
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("F").NumberFormat = "#,##0"
      .Columns("G").ColumnWidth = 16
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("G").NumberFormat = "#,##0"
      .Columns("H").ColumnWidth = 15
      .Columns("H").NumberFormat = "###,###,###,##0.00"
      .Columns("H").HorizontalAlignment = xlHAlignRight
      .Columns("I").ColumnWidth = 15
      .Columns("I").NumberFormat = "###,###,###,##0.00"
      .Columns("I").HorizontalAlignment = xlHAlignRight
      .Columns("J").ColumnWidth = 14.5
      .Columns("J").NumberFormat = "###,###,###,##0.00"
      .Columns("J").HorizontalAlignment = xlHAlignRight
      .Columns("K").ColumnWidth = 14.5
      .Columns("K").NumberFormat = "###,###,###,##0.00"
      .Columns("K").HorizontalAlignment = xlHAlignRight
      .Columns("L").ColumnWidth = 14.5
      .Columns("L").NumberFormat = "###,###,###,##0.00"
      .Columns("L").HorizontalAlignment = xlHAlignRight
      .Columns("M").ColumnWidth = 15
      .Columns("M").NumberFormat = "#,##0"
      .Columns("M").HorizontalAlignment = xlHAlignCenter
      .Columns("N").ColumnWidth = 16
      .Columns("N").HorizontalAlignment = xlHAlignCenter
      .Columns("N").NumberFormat = "#,##0"
      .Columns("O").ColumnWidth = 14.5
      .Columns("O").NumberFormat = "###,###,###,##0.00"
      .Columns("O").HorizontalAlignment = xlHAlignRight
      .Columns("P").ColumnWidth = 14.5
      .Columns("P").NumberFormat = "###,###,###,##0.00"
      .Columns("P").HorizontalAlignment = xlHAlignRight
      .Columns("Q").ColumnWidth = 14.5
      .Columns("Q").NumberFormat = "###,###,###,##0.00"
      .Columns("Q").HorizontalAlignment = xlHAlignRight
      .Columns("R").ColumnWidth = 15
      .Columns("R").NumberFormat = "###,###,###,##0.00"
      .Columns("R").HorizontalAlignment = xlHAlignRight
      .Columns("S").ColumnWidth = 15
      .Columns("S").HorizontalAlignment = xlHAlignCenter
           
      With .Range(.Cells(3, 2), .Cells(5, 19))
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
          .WrapText = True
      End With
      
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(2, 1), .Cells(99, 99)).Font.Size = 11
       
      r_int_NroFil = r_int_NroFil + 2
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " USP_RPT_TPR_CARFIA ("
      g_str_Parame = g_str_Parame & "'" & Format(Month(Now), "00") & "', "
      g_str_Parame = g_str_Parame & "'" & Format(Year(Now), "0000") & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(cmb_TipRep.Text) & "' , "
      g_str_Parame = g_str_Parame & CStr(l_str_FecPer) & " , "
      g_str_Parame = g_str_Parame & CStr(Chk_FecAct.Value) & " , "
      g_str_Parame = g_str_Parame & cmb_TipRep.ItemData(cmb_TipRep.ListIndex) & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "
        
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
        g_rst_Princi.Close
        Set g_rst_Princi = Nothing
        MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
        Exit Sub
      End If
       
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         
         Do While Not g_rst_Princi.EOF
            .Cells(r_int_NroFil, 2) = "'" & g_rst_Princi!DOCUMENTO
            .Cells(r_int_NroFil, 3) = "'" & g_rst_Princi!RAZON_SOCIAL
            .Cells(r_int_NroFil, 4) = Format(CDbl(g_rst_Princi!LINEA_ASIGNADA), "###,###,###,##0.00")
            .Cells(r_int_NroFil, 5) = g_rst_Princi!NRO_CARTA_CF
            .Cells(r_int_NroFil, 6) = g_rst_Princi!NRO_CARTA_AD
            .Cells(r_int_NroFil, 7) = g_rst_Princi!NRO_CARTA_CSO
            .Cells(r_int_NroFil, 8) = Format(CDbl(g_rst_Princi!GARANTIA_LIQUIDA_IND), "###,###,###,##0.00")
            .Cells(r_int_NroFil, 9) = Format(CDbl(g_rst_Princi!GARANTIA_HIPOTECARIO_IND), "###,###,###,##0.00")
            .Cells(r_int_NroFil, 10) = Format(CDbl(g_rst_Princi!LINEA_UTILIZADA_CF), "###,###,###,##0.00")
            .Cells(r_int_NroFil, 11) = Format(CDbl(g_rst_Princi!LINEA_UTILIZADA_AD), "###,###,###,##0.00")
            .Cells(r_int_NroFil, 12) = Format(CDbl(g_rst_Princi!LINEA_UTILIZADA_CSO), "###,###,###,##0.00")
            .Cells(r_int_NroFil, 13) = g_rst_Princi!NRO_DIR_LIN_CREDITO
            .Cells(r_int_NroFil, 14) = g_rst_Princi!NRO_DIR_CRED_PUNTUAL
            .Cells(r_int_NroFil, 15) = Format(CDbl(g_rst_Princi!GARANTIA_LIQUIDA_DIR), "###,###,###,##0.00")
            .Cells(r_int_NroFil, 16) = Format(CDbl(g_rst_Princi!LINEA_UTIL_DIR_LIN_CREDITO), "###,###,###,##0.00")
            .Cells(r_int_NroFil, 17) = Format(CDbl(g_rst_Princi!LINEA_UTIL_DIR_CRE_PUNTUAL), "###,###,###,##0.00")
            .Cells(r_int_NroFil, 18) = Format(CDbl(g_rst_Princi!LINEA_ASIGNADA) - CDbl(g_rst_Princi!LINEA_UTILIZADA_IND) - CDbl(g_rst_Princi!LINEA_UTILIZADA_DIR) + CDbl(g_rst_Princi!LINEA_UTIL_DIR_CRE_PUNTUAL), "###,###,###,##0.00")
            .Cells(r_int_NroFil, 19) = Trim(g_rst_Princi!USUARIO)
            
            r_int_NroFil = r_int_NroFil + 1
            g_rst_Princi.MoveNext
         Loop
      End If
      
      'SUMATORIA TOTAL
      .Cells(r_int_NroFil, 3) = "TOTAL"

      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil, 19)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil, 19)).Font.Bold = True
     ' .Range(.Cells(r_int_NroFil, 5), .Cells(r_int_NroFil, 10)).HorizontalAlignment = xlHAlignRight
      .Range(.Cells(r_int_NroFil, 4), .Cells(r_int_NroFil, 18)).FormulaR1C1 = "=SUM(R[-" & r_int_NroFil - 5 & "]C:R[-1]C)"
            
      With .Range(.Cells(3, 2), .Cells(r_int_NroFil, 19))
         .Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Borders(xlEdgeTop).LineStyle = xlContinuous
         .Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Borders(xlEdgeRight).LineStyle = xlContinuous
         .Borders(xlInsideVertical).LineStyle = xlContinuous
         .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
   End With
   r_obj_Excel.ActiveSheet.Range("D6").Select
   r_obj_Excel.ActiveWindow.FreezePanes = True
   r_obj_Excel.Visible = True
End Sub

Private Sub fs_GenExc_Detalle()
Dim r_obj_Excel         As Excel.Application
Dim r_int_NroFil        As Integer

   r_int_NroFil = 5
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 11 '6
   
   r_obj_Excel.Workbooks.Add
   r_obj_Excel.Sheets(1).Name = "TECHO PROPIO-CARTAS FIANZA"
   r_obj_Excel.Sheets(2).Name = "TECHO PROPIO-ADENDAS"
   r_obj_Excel.Sheets(3).Name = "TECHO PROPIO-C.SERIEDAD OFERTA"
   
   r_obj_Excel.Sheets(4).Name = "DEUDA SF-CARTAS FIANZA"
   r_obj_Excel.Sheets(5).Name = "DEUDA SF-ADENDAS"
   r_obj_Excel.Sheets(6).Name = "DEUDA SF-C.SERIEDAD OFERTA"
   r_obj_Excel.Sheets(7).Name = "COMISION"
   r_obj_Excel.Sheets(8).Name = "DEUDA SF-CRED.DIR LC"
   r_obj_Excel.Sheets(9).Name = "DEUDA SF-CRED.DIR CP"
   r_obj_Excel.Sheets(10).Name = "CREDITO DIRECTO-LC"
   r_obj_Excel.Sheets(11).Name = "CREDITO DIRECTO-CP"
   
   r_obj_Excel.Sheets(3).Select
   
   Call fs_GenExc_DeudaSF_CF(r_obj_Excel, r_int_NroFil)
   If Not (g_rst_Princi.State) = 0 Then
      Call fs_GenExc_DeudaSF_AD(r_obj_Excel, r_int_NroFil)
      Call fs_GenExc_DeudaSF_CSO(r_obj_Excel, r_int_NroFil)
      Call fs_GenExc_DeudaSF_CDir_LC(r_obj_Excel, r_int_NroFil)
      Call fs_GenExc_DeudaSF_CDir_CP(r_obj_Excel, r_int_NroFil)
      Call fs_GenExc_TechoPropio_AD(r_obj_Excel, r_int_NroFil)
      Call fs_GenExc_TechoPropio_CF(r_obj_Excel, r_int_NroFil)
      Call fs_GenExc_TechoPropio_CSO(r_obj_Excel, r_int_NroFil)
      Call fs_GenExc_TechoPropio_Comision(r_obj_Excel, r_int_NroFil)
      Call fs_GenExc_CDir_LC(r_obj_Excel, r_int_NroFil)
      Call fs_GenExc_CDir_CP(r_obj_Excel, r_int_NroFil)
   End If
End Sub
Private Sub fs_GenExc_Riesgo()
Dim r_obj_Excel         As Excel.Application
Dim r_str_ParAux        As String
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
Dim r_int_NroFil        As Integer
Dim r_int_NoFlLi        As Integer
   
   r_int_NroFil = 5
      
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
   
      .Cells(1, 2) = "REPORTE DE CENTRAL DE RIESGOS Y ALINEAMIENTOS "
      .Range(.Cells(1, 2), .Cells(1, 17)).Merge
      .Range(.Cells(1, 2), .Cells(1, 17)).Font.Bold = True
      .Range(.Cells(1, 2), .Cells(1, 17)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 2), .Cells(1, 17)).Font.Size = 14
      
      If Chk_FecAct.Value = 1 Then
         .Cells(3, 2) = " AL " & Format(date, "DD/MM/YYYY")
      Else
         .Cells(3, 2) = " AL " & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), ipp_PerAno.Text), "00") & " DE " & cmb_PerMes.Text & " DE " & Format(ipp_PerAno.Text, "0000")
      End If
      With .Range(.Cells(3, 2), .Cells(3, 2))
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
      End With
      
      .Range(.Cells(3, 2), .Cells(3, 17)).Merge
      
      .Cells(r_int_NroFil, 2) = "TIPO DOC. - DOCUMENTO"
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 2)).Merge
      
      .Cells(r_int_NroFil, 3) = "RAZÓN SOCIAL"
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil + 1, 3)).Merge
      
      .Cells(r_int_NroFil, 4) = "ENTIDAD"
      .Range(.Cells(r_int_NroFil, 4), .Cells(r_int_NroFil + 1, 4)).Merge
           
      .Cells(r_int_NroFil, 5) = "CARTA FIANZA"
      .Range(.Cells(r_int_NroFil, 5), .Cells(r_int_NroFil + 1, 5)).Merge
      
      .Cells(r_int_NroFil, 6) = "ADENDA"
      .Range(.Cells(r_int_NroFil, 6), .Cells(r_int_NroFil + 1, 6)).Merge
      
      .Cells(r_int_NroFil, 7) = "SERIEDAD OFERTA"
      .Range(.Cells(r_int_NroFil, 7), .Cells(r_int_NroFil + 1, 7)).Merge
      
      .Cells(r_int_NroFil, 8) = "TOTAL DEUDA"
      .Range(.Cells(r_int_NroFil, 8), .Cells(r_int_NroFil + 1, 8)).Merge
      
      .Cells(r_int_NroFil, 9) = "CLASIFICACION"
      .Range(.Cells(r_int_NroFil, 9), .Cells(r_int_NroFil + 1, 9)).Merge
      
      .Cells(r_int_NroFil, 10) = "%"
      .Range(.Cells(r_int_NroFil, 10), .Cells(r_int_NroFil + 1, 10)).Merge
      
      .Cells(r_int_NroFil, 11) = "CLASIFICACION INTERNA"
      .Range(.Cells(r_int_NroFil, 11), .Cells(r_int_NroFil + 1, 11)).Merge
      
      .Cells(r_int_NroFil, 12) = "DIAS ATRASO"
      .Range(.Cells(r_int_NroFil, 12), .Cells(r_int_NroFil + 1, 12)).Merge
      
      .Cells(r_int_NroFil, 13) = "SE ALINEA"
      .Range(.Cells(r_int_NroFil, 13), .Cells(r_int_NroFil + 1, 13)).Merge
      
      .Cells(r_int_NroFil, 14) = "NUEVA CLASIFICACION"
      .Range(.Cells(r_int_NroFil, 14), .Cells(r_int_NroFil + 1, 14)).Merge
      
      .Cells(r_int_NroFil, 15) = "NUEVA TASA"
      .Range(.Cells(r_int_NroFil, 15), .Cells(r_int_NroFil + 1, 15)).Merge
      
      .Cells(r_int_NroFil, 16) = "TIPO EMPRESA"
      .Range(.Cells(r_int_NroFil, 16), .Cells(r_int_NroFil + 1, 16)).Merge
      
      .Cells(r_int_NroFil, 17) = "USUARIO"
      .Range(.Cells(r_int_NroFil, 17), .Cells(r_int_NroFil + 1, 17)).Merge
      
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 17)).Interior.Color = 12611584 'RGB(146, 208, 80)
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 17)).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 17)).Font.ThemeColor = xlThemeColorDark1
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 17)).HorizontalAlignment = xlHAlignCenter
      '
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 16
      .Columns("C").ColumnWidth = 60
      .Columns("C").HorizontalAlignment = xlHAlignLeft
      .Columns("D").ColumnWidth = 25.5
      .Columns("D").HorizontalAlignment = xlHAlignLeft
      .Columns("E").ColumnWidth = 16
      .Columns("E").HorizontalAlignment = xlHAlignRight
      .Columns("E").NumberFormat = "###,###,###,##0.00"
      .Columns("F").ColumnWidth = 16
      .Columns("F").HorizontalAlignment = xlHAlignRight
      .Columns("F").NumberFormat = "###,###,###,##0.00"
      .Columns("G").ColumnWidth = 16
      .Columns("G").HorizontalAlignment = xlHAlignRight
      .Columns("G").NumberFormat = "###,###,###,##0.00"
      .Columns("H").ColumnWidth = 15
      .Columns("H").NumberFormat = "###,###,###,##0.00"
      .Columns("H").HorizontalAlignment = xlHAlignRight
      .Columns("I").ColumnWidth = 14
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("I").Font.Bold = True
      .Columns("J").ColumnWidth = 8
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("J").NumberFormat = "#,##0.00"
      .Columns("K").ColumnWidth = 14.5
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Columns("L").ColumnWidth = 14.5
      .Columns("L").NumberFormat = "0"
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      .Columns("M").ColumnWidth = 15
      .Columns("M").HorizontalAlignment = xlHAlignCenter
      .Columns("N").ColumnWidth = 16
      .Columns("N").HorizontalAlignment = xlHAlignCenter
      .Columns("O").ColumnWidth = 14.5
      .Columns("O").HorizontalAlignment = xlHAlignCenter
      .Columns("P").ColumnWidth = 14.5
      .Columns("P").HorizontalAlignment = xlHAlignCenter
      .Columns("Q").ColumnWidth = 14.5
      .Columns("Q").HorizontalAlignment = xlHAlignCenter
                
      With .Range(.Cells(5, 2), .Cells(6, 17))
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
          .WrapText = True
      End With
      
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(2, 1), .Cells(99, 99)).Font.Size = 11
       
      r_int_NroFil = r_int_NroFil + 2
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " USP_RPT_RIESGO_ALINEA ("
      g_str_Parame = g_str_Parame & "'" & Format(cmb_PerMes.ListIndex + 1, "00") & "', "
      g_str_Parame = g_str_Parame & "'" & Format(ipp_PerAno.Value, "0000") & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(cmb_TipRep.Text) & "' , "
      g_str_Parame = g_str_Parame & CStr(l_str_FecPer) & " , "
      g_str_Parame = g_str_Parame & CStr(Chk_FecAct.Value) & " , "
      g_str_Parame = g_str_Parame & cmb_TipRep.ItemData(cmb_TipRep.ListIndex) - 2 & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "
        
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
        g_rst_Princi.Close
        Set g_rst_Princi = Nothing
        MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
        Exit Sub
      End If
       
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         Dim r_str_Cadena As String
         Dim r_int_Cadena As String
         
         Do While Not g_rst_Princi.EOF

            .Cells(r_int_NroFil, 2) = "'" & g_rst_Princi!DOCUMENTO
            .Cells(r_int_NroFil, 3) = "'" & g_rst_Princi!RAZON_SOCIAL
            .Cells(r_int_NroFil, 4) = g_rst_Princi!EMPRESA_SUOERVISADA
            .Cells(r_int_NroFil, 5) = Format(CDbl(g_rst_Princi!LINEA_UTILIZADA_CF), "###,###,###,##0.00")
            .Cells(r_int_NroFil, 6) = Format(CDbl(g_rst_Princi!LINEA_UTILIZADA_AD), "###,###,###,##0.00")
            .Cells(r_int_NroFil, 7) = Format(CDbl(g_rst_Princi!LINEA_UTILIZADA_CSO), "###,###,###,##0.00")
            .Cells(r_int_NroFil, 8) = Format(CDbl(g_rst_Princi!DEUDA_SF), "###,###,###,##0.00")
            .Cells(r_int_NroFil, 9) = IIf(g_rst_Princi!COD_CLASIFICACION = 1, "CPP", Trim(g_rst_Princi!TIPO_CLASIFICACION))
            If g_rst_Princi!COD_CLASIFICACION = 0 Then
               .Cells(r_int_NroFil, 9).Interior.Color = 5287936
            ElseIf g_rst_Princi!COD_CLASIFICACION = 1 Then
               .Cells(r_int_NroFil, 9).Interior.Color = 65535
            ElseIf g_rst_Princi!COD_CLASIFICACION = 2 Then
               .Cells(r_int_NroFil, 9).Interior.Color = 49407
            ElseIf g_rst_Princi!COD_CLASIFICACION = 3 Then
               .Cells(r_int_NroFil, 9).Interior.Color = 255
            ElseIf g_rst_Princi!COD_CLASIFICACION = 4 Then
               .Cells(r_int_NroFil, 9).Interior.Color = 0
            End If
            If g_rst_Princi!COD_CLASIFICACION <> 1 Then .Cells(r_int_NroFil, 9).Font.ThemeColor = xlThemeColorDark1
            .Cells(r_int_NroFil, 10) = ""
            .Cells(r_int_NroFil, 11) = Trim(g_rst_Princi!CLASIFICACION_INTERNA)
            .Cells(r_int_NroFil, 12) = Trim(g_rst_Princi!DIAS_ATRASO)
            .Cells(r_int_NroFil, 13) = ""
            .Cells(r_int_NroFil, 14) = ""
            .Cells(r_int_NroFil, 15) = ""
            .Cells(r_int_NroFil, 16) = Trim(g_rst_Princi!TIPO_EMPRESA)
            .Cells(r_int_NroFil, 17) = Trim(g_rst_Princi!USUARIO)
            .Cells(r_int_NroFil, 18) = g_rst_Princi!COD_CLASIFICACION
            
            r_int_NroFil = r_int_NroFil + 1
            g_rst_Princi.MoveNext

         Loop
         
      End If
      r_int_NoFlLi = r_int_NroFil - 1
      
      Dim x                As String
      Dim i, j, k, L, m    As Integer
      Dim r_dbl_ValMax     As Double
      Dim r_dbl_MaxPor     As Double
      Dim r_int_CodCla     As Integer
      
      'Colocar fórmula para la columna %
      r_int_Contad = 7
      r_str_Cadena = .Cells(r_int_Contad, 2)
      
      k = 0
      For r_int_NroFil = 8 To r_int_NoFlLi
         If r_str_Cadena <> .Cells(r_int_NroFil, 2) Then
            L = 0
            For L = 0 To i
               For j = 0 To i
                  x = x & "+R[" & k & "]C[-2]"
                  k = k + 1
               Next
            
              .Cells(r_int_Contad + L, 10).FormulaR1C1 = "=+RC[-2]/" & "(" & x & ")*100"
              k = L * -1 - 1
              x = ""
            Next
            
            'Pintar el procentaje > 20 y diferente de Normal
'            'Valor del Porcentaje
'            .Cells(r_int_Contad, 13).FormulaR1C1 = "=+MAX(RC[-3]:R[" & l - 1 & "]C[-3])"
'            r_dbl_MaxPor = .Cells(r_int_Contad, 13)
'            .Cells(r_int_Contad, 13) = ""
            
            'Valor de Clasificación
            .Cells(r_int_Contad, 13).FormulaR1C1 = "=+MAX(RC[5]:R[" & L - 1 & "]C[5])"
            r_dbl_ValMax = .Cells(r_int_Contad, 13)
            .Cells(r_int_Contad, 13) = ""
            
            For m = 0 To i
               If r_dbl_ValMax > 1 And .Cells(r_int_Contad + m, 9) <> "NORMAL" And .Cells(r_int_Contad + m, 18) = r_dbl_ValMax Then 'If r_dbl_ValMax > 20 And .Cells(r_int_Contad + m, 9) <> "NORMAL" And .Cells(r_int_Contad + m, 10) = r_dbl_ValMax Then
                  
                  If .Cells(r_int_Contad + m, 10) > 20 Then
                  
                     .Cells(r_int_Contad + m, 10).Font.Color = -16776961
                     .Cells(r_int_Contad + m, 10).Font.Bold = True
                     
                     .Cells(r_int_Contad + m, 13).Value = "SI"
                     .Cells(r_int_Contad + m, 13).Font.Color = -16776961
                     .Cells(r_int_Contad + m, 13).Font.Bold = True
                     
                     r_int_CodCla = CInt(.Cells(r_int_Contad + m, 18)) - 1
                     .Cells(r_int_Contad + m, 14).Value = moddat_gf_ConsultaClasifCred("08", r_int_CodCla)
                     If r_int_CodCla = 1 Then .Cells(r_int_Contad + m, 14).Value = "CPP"
                     .Cells(r_int_Contad + m, 14).Font.Bold = True
                     
                     If r_int_CodCla = 0 Then
                        .Cells(r_int_Contad + m, 14).Interior.Color = 5287936
                     ElseIf r_int_CodCla = 1 Then
                        .Cells(r_int_Contad + m, 14).Interior.Color = 65535
                     ElseIf r_int_CodCla = 2 Then
                        .Cells(r_int_Contad + m, 14).Interior.Color = 49407
                     ElseIf r_int_CodCla = 3 Then
                        .Cells(r_int_Contad + m, 14).Interior.Color = 255
                     ElseIf r_int_CodCla = 4 Then
                        .Cells(r_int_Contad + m, 14).Interior.Color = 0
                     End If
                     
                     .Cells(r_int_Contad + m, 15).Value = "1%"
                  End If
               End If
            Next
            
            r_int_Contad = r_int_NroFil
            r_str_Cadena = .Cells(r_int_NroFil, 2)
            i = 0
            k = 0
         Else
            i = i + 1
         End If
      Next
      
      'última empresa para colocar fórmula
      For L = 0 To i
         For j = 0 To i
            x = x & "+R[" & k & "]C[-2]"
            k = k + 1
         Next
      
        .Cells(r_int_Contad + L, 10).FormulaR1C1 = "=+RC[-2]/" & "(" & x & ")*100"
        k = L * -1 - 1
        x = ""
      Next
      
       'Pintar el procentaje > 20 y diferente de Normal
'       'Valor del Porcentaje
'      .Cells(r_int_Contad, 13).FormulaR1C1 = "=+MAX(RC[-3]:R[" & l - 1 & "]C[-3])"
'      r_dbl_MaxPor = .Cells(r_int_Contad, 13)
'      .Cells(r_int_Contad, 13) = ""
      
      'Valor de Clasificación
      .Cells(r_int_Contad, 13).FormulaR1C1 = "=+MAX(RC[5]:R[" & L - 1 & "]C[5])"
      r_dbl_ValMax = .Cells(r_int_Contad, 13)
      .Cells(r_int_Contad, 13) = ""

      For m = 0 To i
         If r_dbl_ValMax > 1 And .Cells(r_int_Contad + m, 9) <> "NORMAL" And .Cells(r_int_Contad + m, 18) = r_dbl_ValMax Then 'If r_dbl_ValMax > 20 And .Cells(r_int_Contad + m, 9) <> "NORMAL" And .Cells(r_int_Contad + m, 10) = r_dbl_ValMax Then
            
            If .Cells(r_int_Contad + m, 10) > 20 Then
            
               .Cells(r_int_Contad + m, 10).Font.Color = -16776961
               .Cells(r_int_Contad + m, 10).Font.Bold = True
               
               .Cells(r_int_Contad + m, 13).Value = "SI"
               .Cells(r_int_Contad + m, 13).Font.Color = -16776961
               .Cells(r_int_Contad + m, 13).Font.Bold = True
               
               r_int_CodCla = CInt(.Cells(r_int_Contad + m, 18)) - 1
               .Cells(r_int_Contad + m, 14).Value = moddat_gf_ConsultaClasifCred("08", r_int_CodCla)
               If r_int_CodCla = 1 Then .Cells(r_int_Contad + m, 14).Value = "CPP"
               .Cells(r_int_Contad + m, 14).Font.Bold = True
               
               If r_int_CodCla = 0 Then
                  .Cells(r_int_Contad + m, 14).Interior.Color = 5287936
               ElseIf r_int_CodCla = 1 Then
                  .Cells(r_int_Contad + m, 14).Interior.Color = 65535
               ElseIf r_int_CodCla = 2 Then
                  .Cells(r_int_Contad + m, 14).Interior.Color = 49407
               ElseIf r_int_CodCla = 3 Then
                  .Cells(r_int_Contad + m, 14).Interior.Color = 255
               ElseIf r_int_CodCla = 4 Then
                  .Cells(r_int_Contad + m, 14).Interior.Color = 0
               End If
               
               .Cells(r_int_Contad + m, 15).Value = "1%"
            End If
         End If
      Next
      
      'recorre para colocar celdas combinadas
      r_int_Contad = 7
      r_str_Cadena = .Cells(r_int_Contad, 2)
      
      For r_int_NroFil = 8 To r_int_NoFlLi
         If r_str_Cadena <> .Cells(r_int_NroFil, 2) Then
            .Range(.Cells(r_int_Contad + 1, 2), .Cells(r_int_NroFil - 1, 2)) = ""
            .Range(.Cells(r_int_Contad, 2), .Cells(r_int_NroFil - 1, 2)).Merge
            
            .Range(.Cells(r_int_Contad + 1, 3), .Cells(r_int_NroFil - 1, 3)) = ""
            .Range(.Cells(r_int_Contad, 3), .Cells(r_int_NroFil - 1, 3)).Merge
            
            .Range(.Cells(r_int_Contad + 1, 16), .Cells(r_int_NroFil - 1, 16)) = ""
            .Range(.Cells(r_int_Contad, 16), .Cells(r_int_NroFil - 1, 16)).Merge
            
            .Range(.Cells(r_int_Contad + 1, 17), .Cells(r_int_NroFil - 1, 17)) = ""
            .Range(.Cells(r_int_Contad, 17), .Cells(r_int_NroFil - 1, 17)).Merge
            
'            .Cells(r_int_NroFil - 1, 17).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(r_int_NroFil - 1, 2), .Cells(r_int_NroFil - 1, 17)).Borders(xlEdgeBottom).Weight = xlMedium
            
            r_int_Contad = r_int_NroFil
            r_str_Cadena = .Cells(r_int_NroFil, 2)
         End If
      Next
      
     'última empresa
     .Range(.Cells(r_int_Contad + 1, 2), .Cells(r_int_NroFil - 1, 2)) = ""
     .Range(.Cells(r_int_Contad, 2), .Cells(r_int_NroFil - 1, 2)).Merge
      
     .Range(.Cells(r_int_Contad + 1, 3), .Cells(r_int_NroFil - 1, 3)) = ""
     .Range(.Cells(r_int_Contad, 3), .Cells(r_int_NroFil - 1, 3)).Merge
      
     .Range(.Cells(r_int_Contad + 1, 16), .Cells(r_int_NroFil - 1, 16)) = ""
     .Range(.Cells(r_int_Contad, 16), .Cells(r_int_NroFil - 1, 16)).Merge
      
     .Range(.Cells(r_int_Contad + 1, 17), .Cells(r_int_NroFil - 1, 17)) = ""
     .Range(.Cells(r_int_Contad, 17), .Cells(r_int_NroFil - 1, 17)).Merge
            
     '.Range(.Cells(5, 2), .Cells(r_int_NroFil, 3)).HorizontalAlignment = xlCenter
     .Range(.Cells(7, 2), .Cells(r_int_NroFil, 3)).VerticalAlignment = xlCenter
     .Range(.Cells(7, 2), .Cells(r_int_NroFil, 3)).Orientation = 0
     
     .Range(.Cells(7, 16), .Cells(r_int_NroFil, 17)).VerticalAlignment = xlCenter
     .Range(.Cells(7, 16), .Cells(r_int_NroFil, 17)).Orientation = 0
        
      With .Range(.Cells(5, 2), .Cells(r_int_NroFil - 1, 17))
         .Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Borders(xlEdgeLeft).Weight = xlMedium
         .Borders(xlEdgeTop).LineStyle = xlContinuous
         .Borders(xlEdgeTop).Weight = xlMedium
         .Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Borders(xlEdgeBottom).Weight = xlMedium
         .Borders(xlEdgeRight).LineStyle = xlContinuous
         .Borders(xlEdgeRight).Weight = xlMedium
'         .Borders(xlInsideVertical).LineStyle = xlContinuous
'         .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      .Columns("R:R").EntireColumn.Hidden = True
       
      .Cells(r_int_Contad + m, 2) = " Notas:"
      .Cells(r_int_Contad + m, 2).Font.Bold = True
      .Cells(r_int_Contad + m + 1, 2) = " Se alinean aquellos clientes que sus deudas represente un mínimo del 20%, se permite un nivel de discrepancia."
      .Cells(r_int_Contad + m + 2, 2) = " Se consideran deudas directas e indirectas (excepto los créditos no desembolsados y las líneas no utilizadas), incluye las carteras castigadas y las carteras en liquidación."

   End With
   r_obj_Excel.ActiveSheet.Range("D7").Select
   r_obj_Excel.ActiveWindow.FreezePanes = True
   r_obj_Excel.Visible = True
End Sub
Private Sub fs_GenExc_DocEle()
Dim r_obj_Excel         As Excel.Application
Dim r_str_ParAux        As String
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
Dim r_int_NroFil        As Integer
Dim r_int_NoFlLi        As Integer
Dim r_str_FecIni        As String
Dim r_str_FecFin        As String
   
   r_int_NroFil = 5
      
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
   
      .Cells(1, 2) = "REPORTE PARA EMISION DE DOCUMENTOS ELECTRONICOS "
      .Range(.Cells(1, 2), .Cells(1, 23)).Merge
      .Range(.Cells(1, 2), .Cells(1, 23)).Font.Bold = True
      .Range(.Cells(1, 2), .Cells(1, 23)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 2), .Cells(1, 23)).Font.Size = 14
      
      If Chk_FecAct.Value = 1 Then
         .Cells(3, 2) = " AL " & Format(date, "DD/MM/YYYY")
      Else
         .Cells(3, 2) = " AL " & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), ipp_PerAno.Text), "00") & " DE " & cmb_PerMes.Text & " DE " & Format(ipp_PerAno.Text, "0000")
      End If
      With .Range(.Cells(3, 2), .Cells(3, 2))
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
      End With
      
      .Range(.Cells(3, 2), .Cells(3, 23)).Merge
      
      .Cells(r_int_NroFil, 2) = "ITEM"
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 2)).Merge
      
      .Cells(r_int_NroFil, 3) = "TIPCOM"
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil + 1, 3)).Merge
      
      .Cells(r_int_NroFil, 4) = "TIPPRO"
      .Range(.Cells(r_int_NroFil, 4), .Cells(r_int_NroFil + 1, 4)).Merge
           
      .Cells(r_int_NroFil, 5) = "FECEMI"
      .Range(.Cells(r_int_NroFil, 5), .Cells(r_int_NroFil + 1, 5)).Merge
      
      .Cells(r_int_NroFil, 6) = "MONEDA"
      .Range(.Cells(r_int_NroFil, 6), .Cells(r_int_NroFil + 1, 6)).Merge
      
      .Cells(r_int_NroFil, 7) = "TIPCAM"
      .Range(.Cells(r_int_NroFil, 7), .Cells(r_int_NroFil + 1, 7)).Merge
      
      .Cells(r_int_NroFil, 8) = "TIPDOC"
      .Range(.Cells(r_int_NroFil, 8), .Cells(r_int_NroFil + 1, 8)).Merge
      
      .Cells(r_int_NroFil, 9) = "NUMDOC"
      .Range(.Cells(r_int_NroFil, 9), .Cells(r_int_NroFil + 1, 9)).Merge
      
      .Cells(r_int_NroFil, 10) = "EMPRESA"
      .Range(.Cells(r_int_NroFil, 10), .Cells(r_int_NroFil + 1, 10)).Merge
      
      .Cells(r_int_NroFil, 11) = "DIRECCION"
      .Range(.Cells(r_int_NroFil, 11), .Cells(r_int_NroFil + 1, 11)).Merge
      
      .Cells(r_int_NroFil, 12) = "DISTRITO"
      .Range(.Cells(r_int_NroFil, 12), .Cells(r_int_NroFil + 1, 12)).Merge
      
      .Cells(r_int_NroFil, 13) = "PROVINCIA"
      .Range(.Cells(r_int_NroFil, 13), .Cells(r_int_NroFil + 1, 13)).Merge
      
      .Cells(r_int_NroFil, 14) = "DEPARTAMENTO"
      .Range(.Cells(r_int_NroFil, 14), .Cells(r_int_NroFil + 1, 14)).Merge
      
      .Cells(r_int_NroFil, 15) = "CORREO"
      .Range(.Cells(r_int_NroFil, 15), .Cells(r_int_NroFil + 1, 15)).Merge
      
      .Cells(r_int_NroFil, 16) = "CANTIDAD"
      .Range(.Cells(r_int_NroFil, 16), .Cells(r_int_NroFil + 1, 16)).Merge
      
      .Cells(r_int_NroFil, 17) = "CODIGO"
      .Range(.Cells(r_int_NroFil, 17), .Cells(r_int_NroFil + 1, 17)).Merge
      
      .Cells(r_int_NroFil, 18) = "UM"
      .Range(.Cells(r_int_NroFil, 18), .Cells(r_int_NroFil + 1, 18)).Merge
      
      .Cells(r_int_NroFil, 19) = "GLOSA"
      .Range(.Cells(r_int_NroFil, 19), .Cells(r_int_NroFil + 1, 19)).Merge
      
      .Cells(r_int_NroFil, 20) = "VALUNI"
      .Range(.Cells(r_int_NroFil, 20), .Cells(r_int_NroFil + 1, 20)).Merge
      
      .Cells(r_int_NroFil, 21) = "VALVTA"
      .Range(.Cells(r_int_NroFil, 21), .Cells(r_int_NroFil + 1, 21)).Merge
      
      .Cells(r_int_NroFil, 22) = "NUMREF"
      .Range(.Cells(r_int_NroFil, 22), .Cells(r_int_NroFil + 1, 22)).Merge
            
      .Cells(r_int_NroFil, 23) = "OBSDOC"
      .Range(.Cells(r_int_NroFil, 23), .Cells(r_int_NroFil + 1, 23)).Merge

      .Cells(r_int_NroFil, 24) = "USUARIO"
      .Range(.Cells(r_int_NroFil, 24), .Cells(r_int_NroFil + 1, 24)).Merge
      
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 24)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 24)).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 24)).HorizontalAlignment = xlHAlignCenter
      '
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 16
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 12
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 25.5
      .Columns("D").HorizontalAlignment = xlHAlignLeft
      .Columns("E").ColumnWidth = 16
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 16
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 16
      .Columns("G").HorizontalAlignment = xlHAlignRight
      .Columns("G").NumberFormat = "###,###,###,##0.00"
      .Columns("H").ColumnWidth = 16
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 16
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 60
      .Columns("J").HorizontalAlignment = xlHAlignLeft
      .Columns("K").ColumnWidth = 60
      .Columns("K").HorizontalAlignment = xlHAlignLeft
      .Columns("L").ColumnWidth = 25
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      .Columns("M").ColumnWidth = 25
      .Columns("M").HorizontalAlignment = xlHAlignCenter
      .Columns("N").ColumnWidth = 25
      .Columns("N").HorizontalAlignment = xlHAlignCenter
      .Columns("O").ColumnWidth = 42
      .Columns("O").HorizontalAlignment = xlHAlignLeft
      .Columns("P").ColumnWidth = 14.5
      .Columns("P").HorizontalAlignment = xlHAlignCenter
      .Columns("Q").ColumnWidth = 14.5
      .Columns("Q").HorizontalAlignment = xlHAlignCenter
      .Columns("R").ColumnWidth = 14.5
      .Columns("R").HorizontalAlignment = xlHAlignCenter
      .Columns("S").ColumnWidth = 30
      .Columns("S").HorizontalAlignment = xlHAlignLeft
      .Columns("T").ColumnWidth = 16
      .Columns("T").HorizontalAlignment = xlHAlignRight
      .Columns("T").NumberFormat = "###,###,###,##0.00"
      .Columns("U").ColumnWidth = 16
      .Columns("U").HorizontalAlignment = xlHAlignRight
      .Columns("U").NumberFormat = "###,###,###,##0.00"
      .Columns("V").ColumnWidth = 20
      .Columns("V").HorizontalAlignment = xlHAlignCenter
      .Columns("W").ColumnWidth = 36
      .Columns("W").HorizontalAlignment = xlHAlignCenter
      .Columns("X").ColumnWidth = 20
      .Columns("X").HorizontalAlignment = xlHAlignCenter
      
      With .Range(.Cells(5, 2), .Cells(6, 24))
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
          .WrapText = True
      End With
      
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(2, 1), .Cells(99, 99)).Font.Size = 11
       
      r_int_NroFil = r_int_NroFil + 2
      
      r_str_FecIni = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "01"
      r_str_FecFin = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), ipp_PerAno.Text), "00")
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "  SELECT MAECFI_NUMREF, TIPCOM, MAERDE_FECASG, MONEDA, MAEETE_TIPDOC, MAEETE_NUMDOC, RAZON_SOCIAL, DIRECION, DISTRITO, PROVINCIA, DEPARTAMENTO, MAEETE_DIRELE, USUARIO, SUM(MAERDE_IMPORT) AS IMPORTE "
      g_str_Parame = g_str_Parame & "    FROM (SELECT MAECFI_NUMREF, 'F' AS TIPCOM, MAERDE_FECASG , TRIM(H.PARDES_DESCRI) AS MONEDA , MAEETE_TIPDOC, MAEETE_NUMDOC, TRIM(MAEPRV_RAZSOC) AS RAZON_SOCIAL, TRIM(MAEETE_DIRREP) AS DIRECION, "
      g_str_Parame = g_str_Parame & "                 TRIM(E.PARDES_DESCRI) AS DISTRITO, TRIM(F.PARDES_DESCRI) AS PROVINCIA, TRIM(G.PARDES_DESCRI) AS DEPARTAMENTO, MAEETE_DIRELE , MAERDE_IMPORT, TRIM(C.SEGUSUCRE) AS USUARIO "
      g_str_Parame = g_str_Parame & "            FROM TPR_MAECFI A "
      g_str_Parame = g_str_Parame & "                 INNER JOIN TPR_MAEETE B ON MAEETE_TIPDOC = MAECFI_TIPDOC AND MAEETE_NUMDOC = MAECFI_NUMDOC "
      g_str_Parame = g_str_Parame & "                 INNER JOIN TPR_MAERDE C ON MAERDE_NUMREF = MAECFI_NUMREF AND MAERDE_CODIGO IN(1, 2,13) AND MAERDE_FECASG >= " & r_str_FecIni & " AND MAERDE_FECASG <= " & r_str_FecFin & " "
      g_str_Parame = g_str_Parame & "                 INNER JOIN CNTBL_MAEPRV D ON MAEPRV_TIPDOC = MAECFI_TIPDOC AND MAEETE_NUMDOC = MAEPRV_NUMDOC "
      g_str_Parame = g_str_Parame & "                  LEFT JOIN MNT_PARDES E ON E.PARDES_CODGRP = 101 AND E.PARDES_CODITE = MAEETE_UBIGEO "
      g_str_Parame = g_str_Parame & "                  LEFT JOIN MNT_PARDES F ON F.PARDES_CODGRP = 101 AND F.PARDES_CODITE = SUBSTR(MAEETE_UBIGEO,1,4)||'00' "
      g_str_Parame = g_str_Parame & "                  LEFT JOIN MNT_PARDES G ON G.PARDES_CODGRP = 101 AND G.PARDES_CODITE = SUBSTR(MAEETE_UBIGEO,1,2)||'0000' "
      g_str_Parame = g_str_Parame & "                  LEFT JOIN MNT_PARDES H ON H.PARDES_CODGRP = 204 AND H.PARDES_CODITE = MAERDE_TIPMON "
      g_str_Parame = g_str_Parame & "         )"
      g_str_Parame = g_str_Parame & "    GROUP BY MAECFI_NUMREF, TIPCOM, MAERDE_FECASG, MONEDA, MAEETE_TIPDOC, MAEETE_NUMDOC, RAZON_SOCIAL, DIRECION, DISTRITO, PROVINCIA, DEPARTAMENTO, MAEETE_DIRELE, USUARIO "
      g_str_Parame = g_str_Parame & "    ORDER BY MAEETE_NUMDOC, MAERDE_FECASG, MAECFI_NUMREF "
       
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
        g_rst_Princi.Close
        Set g_rst_Princi = Nothing
        MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
        Exit Sub
      End If
       
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         
         Do While Not g_rst_Princi.EOF

            .Cells(r_int_NroFil, 2) = r_int_NroFil - 7 + 1
            .Cells(r_int_NroFil, 3) = "'" & g_rst_Princi!TIPCOM
            If Mid(g_rst_Princi!MAECFI_NUMREF, 1, 1) = 1 Then
               .Cells(r_int_NroFil, 4) = "3- COMISION CF"
            ElseIf Mid(g_rst_Princi!MAECFI_NUMREF, 1, 1) = 2 Then
               .Cells(r_int_NroFil, 4) = "6- COMISION AD"
            ElseIf Mid(g_rst_Princi!MAECFI_NUMREF, 1, 1) = 3 Then
               .Cells(r_int_NroFil, 4) = "4- COMISION CSO"
            End If
            .Cells(r_int_NroFil, 5) = "'" & Format(gf_FormatoFecha(CStr(g_rst_Princi!MAERDE_FECASG)), "dd/mm/yyyy")
            .Cells(r_int_NroFil, 6) = g_rst_Princi!Moneda
            .Cells(r_int_NroFil, 7) = 0
            .Cells(r_int_NroFil, 8) = g_rst_Princi!MAEETE_TIPDOC
            .Cells(r_int_NroFil, 9) = g_rst_Princi!MAEETE_NUMDOC
            .Cells(r_int_NroFil, 10) = g_rst_Princi!RAZON_SOCIAL
            .Cells(r_int_NroFil, 11) = g_rst_Princi!DIRECION
            .Cells(r_int_NroFil, 12) = g_rst_Princi!DISTRITO
            .Cells(r_int_NroFil, 13) = g_rst_Princi!PROVINCIA
            .Cells(r_int_NroFil, 14) = g_rst_Princi!DEPARTAMENTO
            .Cells(r_int_NroFil, 15) = Trim(g_rst_Princi!MAEETE_DIRELE)
            .Cells(r_int_NroFil, 16) = 1                                            'CANTIDAD
            .Cells(r_int_NroFil, 17) = "--"                                         'CODIGO
            .Cells(r_int_NroFil, 18) = "NIU"                                        'UM
            .Cells(r_int_NroFil, 19) = Trim(Mid(.Cells(r_int_NroFil, 4), 3))        'GLOSA
            .Cells(r_int_NroFil, 20) = Format(CDbl(g_rst_Princi!IMPORTE), "###,###,###,##0.00")
            .Cells(r_int_NroFil, 21) = Format(CDbl(g_rst_Princi!IMPORTE), "###,###,###,##0.00")
            .Cells(r_int_NroFil, 22) = "'" & gf_Formato_NumRef(g_rst_Princi!MAECFI_NUMREF, Mid(g_rst_Princi!MAECFI_NUMREF, 1, 1))
            .Cells(r_int_NroFil, 23) = ""
            .Cells(r_int_NroFil, 24) = g_rst_Princi!USUARIO
          
            r_int_NroFil = r_int_NroFil + 1
            g_rst_Princi.MoveNext

         Loop
         
      End If
      r_int_NoFlLi = r_int_NroFil - 1

      r_obj_Excel.Range("E7").Select
      r_obj_Excel.ActiveWindow.FreezePanes = True
      r_obj_Excel.Cells(1, 1).Select
      r_obj_Excel.Visible = True
   End With
End Sub

 '----------------------------------------------------------------    REPORTE DEUDA SF  - CARTAS FIANZA -------------------------------------------------------------------

Private Sub fs_GenExc_DeudaSF_CF(ByVal p_obj_Excel As Excel.Application, ByVal p_int_NroFil As Integer)
Dim r_dbl_PatEfe        As Double

     With p_obj_Excel.Sheets(4)
     
      .Cells(1, 2) = "REPORTE - SISTEMA FINANCIERO"
      .Range(.Cells(1, 2), .Cells(1, 17)).Merge
      
      If Chk_FecAct.Value = 1 Then
         .Cells(3, 2) = " AL " & Format(date, "DD/MM/YYYY")
      Else
         .Cells(3, 2) = " AL " & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), ipp_PerAno.Text), "00") & " DE " & cmb_PerMes.Text & " DE " & Format(ipp_PerAno.Text, "0000")
      End If
      .Range(.Cells(3, 2), .Cells(3, 17)).Merge
      
      .Range(.Cells(1, 2), .Cells(1, 17)).Font.Bold = True
      .Range(.Cells(1, 2), .Cells(3, 17)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 2), .Cells(1, 17)).Font.Size = 14
   
      .Cells(p_int_NroFil, 2) = "CODIGO SBS"
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 2)).Merge
      .Cells(p_int_NroFil, 3) = "TIPO - DOCUMENTO"
      .Range(.Cells(p_int_NroFil, 3), .Cells(p_int_NroFil + 1, 3)).Merge
      .Cells(p_int_NroFil, 4) = "ENTIDAD"
      .Range(.Cells(p_int_NroFil, 4), .Cells(p_int_NroFil + 1, 4)).Merge
      
      .Cells(p_int_NroFil, 5) = "SUB-PRODUCTO"
      .Range(.Cells(p_int_NroFil, 5), .Cells(p_int_NroFil + 1, 5)).Merge
      
      .Cells(p_int_NroFil, 6) = "PATRIMONIO EFECTIVO"
                
      .Cells(p_int_NroFil, 7) = "TIPO GARANTIA"
      .Range(.Cells(p_int_NroFil, 7), .Cells(p_int_NroFil + 1, 7)).Merge
      
      .Cells(p_int_NroFil, 8) = "GARANTIZADO"
      .Range(.Cells(p_int_NroFil, 8), .Cells(p_int_NroFil + 1, 8)).Merge
      
      .Cells(p_int_NroFil, 9) = "GARANTIAS"
      .Range(.Cells(p_int_NroFil, 9), .Cells(p_int_NroFil, 11)).Merge
      .Cells(p_int_NroFil + 1, 9) = "LIQUIDA"
      .Cells(p_int_NroFil + 1, 10) = "HIPOTECARIA"
      .Cells(p_int_NroFil + 1, 11) = "V.REALIZACION"
      
      .Cells(p_int_NroFil, 12) = "DEUDA SF"
      .Range(.Cells(p_int_NroFil, 12), .Cells(p_int_NroFil + 1, 12)).Merge
      
      .Cells(p_int_NroFil, 13) = "TOTAL"
      .Range(.Cells(p_int_NroFil, 13), .Cells(p_int_NroFil + 1, 13)).Merge
   
      .Cells(p_int_NroFil, 14) = "TIPO"
      .Range(.Cells(p_int_NroFil, 14), .Cells(p_int_NroFil + 1, 14)).Merge
      
      .Cells(p_int_NroFil, 15) = "CLASIFICACION"
      .Range(.Cells(p_int_NroFil, 15), .Cells(p_int_NroFil + 1, 15)).Merge
      
      .Cells(p_int_NroFil, 16) = "NRO. ENTIDADES"
      .Range(.Cells(p_int_NroFil, 16), .Cells(p_int_NroFil + 1, 16)).Merge
   
      .Cells(p_int_NroFil, 17) = "VENTAS"
      .Range(.Cells(p_int_NroFil, 17), .Cells(p_int_NroFil + 1, 17)).Merge
      
      .Cells(p_int_NroFil, 18) = "PATRIMONIO"
      .Range(.Cells(p_int_NroFil, 18), .Cells(p_int_NroFil + 1, 18)).Merge
   
      .Cells(p_int_NroFil, 19) = "USUARIO"
      .Range(.Cells(p_int_NroFil, 19), .Cells(p_int_NroFil + 1, 19)).Merge
      
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 19)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 19)).Font.Bold = True
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 19)).HorizontalAlignment = xlHAlignCenter
   
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 16.5
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 15
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 60
      
      .Columns("E").ColumnWidth = 15
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      
      .Columns("F").ColumnWidth = 15
      .Columns("F").NumberFormat = "###,###,###,##0.00"
      .Columns("F").HorizontalAlignment = xlHAlignRight
            
      .Columns("G").ColumnWidth = 15
      .Columns("H").ColumnWidth = 15
      .Columns("H").NumberFormat = "###,###,###,##0.00"
      .Columns("H").HorizontalAlignment = xlHAlignRight
      
      .Columns("I").ColumnWidth = 15
      .Columns("I").NumberFormat = "###,###,###,##0.00"
      .Columns("I").HorizontalAlignment = xlHAlignRight
      .Columns("J").ColumnWidth = 15
      .Columns("J").NumberFormat = "###,###,###,##0.00"
      .Columns("J").HorizontalAlignment = xlHAlignRight
      
      .Columns("K").ColumnWidth = 15
      .Columns("K").NumberFormat = "###,###,###,##0.00"
      .Columns("K").HorizontalAlignment = xlHAlignRight
      
      .Columns("L").ColumnWidth = 15
      .Columns("L").NumberFormat = "###,###,###,##0.00"
      .Columns("L").HorizontalAlignment = xlHAlignRight
      
      .Columns("M").ColumnWidth = 15
      .Columns("M").NumberFormat = "###,###,###,##0.00"
      .Columns("M").HorizontalAlignment = xlHAlignRight
      
      .Columns("N").ColumnWidth = 15
      .Columns("N").HorizontalAlignment = xlHAlignCenter
      
      .Columns("O").ColumnWidth = 25
      .Columns("O").HorizontalAlignment = xlHAlignCenter
      
      .Columns("P").ColumnWidth = 16
      .Columns("P").HorizontalAlignment = xlHAlignCenter
      
      .Columns("Q").ColumnWidth = 13.5
      .Columns("Q").NumberFormat = "###,###,###,##0.00"
      .Columns("Q").HorizontalAlignment = xlHAlignRight
      .Columns("R").ColumnWidth = 13.5
      .Columns("R").NumberFormat = "###,###,###,##0.00"
      .Columns("R").HorizontalAlignment = xlHAlignRight
      .Columns("S").ColumnWidth = 13.5
      .Columns("S").HorizontalAlignment = xlHAlignCenter
   
      With .Range(.Cells(5, 2), .Cells(6, 19))
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
          .WrapText = True
      End With
   
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(2, 1), .Cells(99, 99)).Font.Size = 11
                
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " USP_RPT_TPR_CARFIA ("
      g_str_Parame = g_str_Parame & "'" & Month(Format(gf_FormatoFecha(CStr(l_str_FecPer)), "dd/mm/yyyy")) & "', "   'Month(Now)
      g_str_Parame = g_str_Parame & "'" & Year(Format(gf_FormatoFecha(CStr(l_str_FecPer)), "dd/mm/yyyy")) & "', " 'Year(Now)
      g_str_Parame = g_str_Parame & "'" & Trim(cmb_TipRep.Text) & "' , "
      g_str_Parame = g_str_Parame & CStr(l_str_FecPer) & " , "
      g_str_Parame = g_str_Parame & CStr(Chk_FecAct.Value) & " , "
      g_str_Parame = g_str_Parame & cmb_TipRep.ItemData(cmb_TipRep.ListIndex) & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "
        
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
        g_rst_Princi.Close
        'Set g_rst_Princi = Nothing
        MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
        Exit Sub
      End If
      
      p_int_NroFil = p_int_NroFil + 2
      g_rst_Princi.MoveFirst
      
      If Not IsNull(g_rst_Princi!PATRIMONIO_EFECTIVO) Then
         r_dbl_PatEfe = CDbl(g_rst_Princi!PATRIMONIO_EFECTIVO)
      Else
         r_dbl_PatEfe = 0
      End If
      
      .Cells(6, 5) = r_dbl_PatEfe
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!TIPO = "CF" Then 'If g_rst_Princi!TIPO <> 2 And g_rst_Princi!TIPO <> 3 And g_rst_Princi!TIPO <> 0 Then
            .Cells(p_int_NroFil, 2) = g_rst_Princi!CODIGO_SBS
            .Cells(p_int_NroFil, 3) = g_rst_Princi!DOCUMENTO
            .Cells(p_int_NroFil, 4) = g_rst_Princi!RAZON_SOCIAL
            .Cells(p_int_NroFil, 5) = g_rst_Princi!SUB_PRODUCTO
            If r_dbl_PatEfe = 0 Then
               .Cells(p_int_NroFil, 6) = 0
            Else
               .Cells(p_int_NroFil, 6) = IIf(r_dbl_PatEfe = 0, 0, (CDbl(g_rst_Princi!GARANTIZADO) / CDbl(r_dbl_PatEfe)) * 100)
            End If
            .Cells(p_int_NroFil, 7) = "" 'g_rst_Princi!TIPO_GARANTIA
            .Cells(p_int_NroFil, 8) = Trim(g_rst_Princi!GARANTIZADO)
            .Cells(p_int_NroFil, 9) = g_rst_Princi!MONTO_GARANTIA_LIQUIDA
            .Cells(p_int_NroFil, 10) = g_rst_Princi!MONTO_GARANTIA_HIPOTECARIA
            .Cells(p_int_NroFil, 11) = g_rst_Princi!VALOR_REALIZACION
            .Cells(p_int_NroFil, 12) = g_rst_Princi!DEUDA_SF
            .Cells(p_int_NroFil, 13) = g_rst_Princi!Total
            .Cells(p_int_NroFil, 14) = g_rst_Princi!TIPO_EMPRESA
            .Cells(p_int_NroFil, 15) = g_rst_Princi!CLASIFICACION
            If g_rst_Princi!CLASIFICACION <> "NORMAL" Then
               .Cells(p_int_NroFil, 15).Font.Color = -16776961
               .Cells(p_int_NroFil, 15).Font.Bold = True
            End If
            .Cells(p_int_NroFil, 16) = g_rst_Princi!NRO_ENTIDADES
            .Cells(p_int_NroFil, 17) = g_rst_Princi!VENTAS
            .Cells(p_int_NroFil, 18) = g_rst_Princi!PATRIMONIO
            .Cells(p_int_NroFil, 19) = g_rst_Princi!USUARIO
            p_int_NroFil = p_int_NroFil + 1
         End If
         g_rst_Princi.MoveNext
      Loop
        
      'SUMATORIA TOTAL
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 19)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 19)).Font.Bold = True
      .Range(.Cells(p_int_NroFil, 5), .Cells(p_int_NroFil, 19)).HorizontalAlignment = xlHAlignRight
      
      .Cells(p_int_NroFil, 4) = "TOTAL"
      .Range(.Cells(p_int_NroFil, 8), .Cells(p_int_NroFil, 12)).FormulaR1C1 = "=SUM(R[-" & p_int_NroFil - 5 & "]C:R[-1]C)"
      
      'RESUMEN POR TIPO EMPRESA
      .Range(.Cells(p_int_NroFil + 2, 3), .Cells(p_int_NroFil + 2, 7)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil + 2, 3), .Cells(p_int_NroFil + 2, 7)).Font.Bold = True
      .Cells(p_int_NroFil + 2, 3) = "CANTIDAD"
      .Cells(p_int_NroFil + 2, 4) = "DESCRIPCION"
      .Cells(p_int_NroFil + 2, 7) = "GARANTIZADO"
      
      .Range(.Cells(p_int_NroFil + 2, 4), .Cells(p_int_NroFil + 7, 4)).Font.Bold = True
      

      .Cells(p_int_NroFil + 3, 3).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 4) & "]C[11]:R[-4]C[11],""PEQUEÑA"")"
      .Cells(p_int_NroFil + 3, 4) = "RESUMEN PEQUEÑA EMPRESA"
      .Cells(p_int_NroFil + 3, 7).FormulaR1C1 = "=SUMIF(R[-" & (p_int_NroFil - 4) & "]C[7]:R[-4]C[7],""PEQUEÑA"",R[-" & (p_int_NroFil - 4) & "]C[1]:R[-4]C[1])"
      
      .Cells(p_int_NroFil + 4, 3).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 3) & "]C[11]:R[-5]C[11],""MEDIANA"")"
      .Cells(p_int_NroFil + 4, 4) = "RESUMEN MEDIANA EMPRESA"
      .Cells(p_int_NroFil + 4, 7).FormulaR1C1 = "=SUMIF(R[-" & (p_int_NroFil - 3) & "]C[7]:R[-5]C[7],""MEDIANA"",R[-" & (p_int_NroFil - 3) & "]C[1]:R[-5]C[1])"
      
      .Cells(p_int_NroFil + 5, 3).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 2) & "]C[11]:R[-6]C[11],""GRANDE"")"
      .Cells(p_int_NroFil + 5, 4) = "RESUMEN GRANDE EMPRESA"
      .Cells(p_int_NroFil + 5, 7).FormulaR1C1 = "=+SUMIF(R[-" & (p_int_NroFil - 2) & "]C[7]:R[-6]C[7],""GRANDE"",R[-" & (p_int_NroFil - 2) & "]C[1]:R[-6]C[1])"
      
      .Cells(p_int_NroFil + 6, 3).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 1) & "]C[11]:R[-7]C[11],""MICRO"")"
      .Cells(p_int_NroFil + 6, 4) = "RESUMEN MICRO EMPRESA"
      .Cells(p_int_NroFil + 6, 7).FormulaR1C1 = "=+SUMIF(R[-" & (p_int_NroFil - 1) & "]C[7]:R[-7]C[7],""MICRO"",R[-" & (p_int_NroFil - 1) & "]C[1]:R[-7]C[1])"
      
      
      .Cells(p_int_NroFil + 7, 3).FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
      .Cells(p_int_NroFil + 7, 4) = "TOTAL"
      .Cells(p_int_NroFil + 7, 7).FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
      
      .Range(.Cells(p_int_NroFil + 7, 3), .Cells(p_int_NroFil + 7, 7)).Font.Bold = True
      .Range(.Cells(p_int_NroFil + 4, 7), .Cells(p_int_NroFil + 7, 7)).NumberFormat = "###,###,###,##0.00"
      .Range(.Cells(p_int_NroFil + 4, 7), .Cells(p_int_NroFil + 7, 7)).HorizontalAlignment = xlHAlignRight
      
      With .Range(.Cells(1, 2), .Cells(1, 19))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:            .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:          .Borders(xlEdgeRight).Weight = xlMedium
      End With
             
      With .Range(.Cells(5, 2), .Cells(p_int_NroFil - 1, 19))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:            .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:     .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
      With .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 19))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:            .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:     .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
      With .Range(.Cells(5, 2), .Cells(6, 19))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:            .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:     .Borders(xlInsideVertical).Weight = xlThin
          .Borders(xlInsideHorizontal).LineStyle = xlContinuous:   .Borders(xlInsideHorizontal).Weight = xlThin
      End With
      'BORDE RESUMEN
      With .Range(.Cells(p_int_NroFil + 2, 3), .Cells(p_int_NroFil + 6, 7))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlThin
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlThin
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlThin
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlThin
      End With
      'Borde del Total
      With .Range(.Cells(p_int_NroFil + 7, 3), .Cells(p_int_NroFil + 7, 7))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous:            .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous:          .Borders(xlEdgeRight).LineStyle = xlContinuous
      End With
      
   End With
   p_obj_Excel.ActiveSheet.Range("E7").Select
   p_obj_Excel.ActiveWindow.FreezePanes = True
   
End Sub

'---------------------------------------------------------------- REPORTE DEUDA SF - ADENDAS ---------------------------------------------------------
Private Sub fs_GenExc_DeudaSF_AD(ByVal p_obj_Excel As Excel.Application, ByVal p_int_NroFil As Integer)
   
Dim r_int_ConAux        As Integer
Dim r_dbl_PatEfe        As Double
Dim r_bol_Estado        As Boolean

'  p_int_NroFil = 5
   p_obj_Excel.Sheets(5).Select

   With p_obj_Excel.Sheets(5)
      .Cells(1, 2) = "REPORTE - SISTEMA FINANCIERO"
      .Range(.Cells(1, 2), .Cells(1, 17)).Merge
      
      If Chk_FecAct.Value = 1 Then
         .Cells(3, 2) = " AL " & Format(date, "DD/MM/YYYY")
      Else
         .Cells(3, 2) = " AL " & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), ipp_PerAno.Text), "00") & " DE " & cmb_PerMes.Text & " DE " & Format(ipp_PerAno.Text, "0000")
      End If
      .Range(.Cells(3, 2), .Cells(3, 17)).Merge
      
      .Range(.Cells(1, 2), .Cells(1, 17)).Font.Bold = True
      .Range(.Cells(1, 2), .Cells(3, 17)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 2), .Cells(1, 17)).Font.Size = 14
   
      .Cells(p_int_NroFil, 2) = "CODIGO SBS"
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 2)).Merge
      .Cells(p_int_NroFil, 3) = "TIPO - DOCUMENTO"
      .Range(.Cells(p_int_NroFil, 3), .Cells(p_int_NroFil + 1, 3)).Merge
      .Cells(p_int_NroFil, 4) = "ENTIDAD"
      .Range(.Cells(p_int_NroFil, 4), .Cells(p_int_NroFil + 1, 4)).Merge
      
      .Cells(p_int_NroFil, 5) = "SUB-PRODUCTO"
      .Range(.Cells(p_int_NroFil, 5), .Cells(p_int_NroFil + 1, 5)).Merge
      
      .Cells(p_int_NroFil, 6) = "PATRIMONIO EFECTIVO"
      .Cells(p_int_NroFil, 7) = "TIPO GARANTIA"
      .Range(.Cells(p_int_NroFil, 7), .Cells(p_int_NroFil + 1, 7)).Merge
      
      .Cells(p_int_NroFil, 8) = "GARANTIZADO"
      .Range(.Cells(p_int_NroFil, 8), .Cells(p_int_NroFil + 1, 8)).Merge
      
      .Cells(p_int_NroFil, 9) = "GARANTIAS"
      .Range(.Cells(p_int_NroFil, 9), .Cells(p_int_NroFil, 11)).Merge
      .Cells(p_int_NroFil + 1, 9) = "LIQUIDA"
      .Cells(p_int_NroFil + 1, 10) = "HIPOTECARIA"
      .Cells(p_int_NroFil + 1, 11) = "V.REALIZACION"
      
      .Cells(p_int_NroFil, 12) = "DEUDA SF"
      .Range(.Cells(p_int_NroFil, 12), .Cells(p_int_NroFil + 1, 12)).Merge
      .Cells(p_int_NroFil, 13) = "TOTAL"
      .Range(.Cells(p_int_NroFil, 13), .Cells(p_int_NroFil + 1, 13)).Merge
   
      .Cells(p_int_NroFil, 14) = "TIPO"
      .Range(.Cells(p_int_NroFil, 14), .Cells(p_int_NroFil + 1, 14)).Merge
      .Cells(p_int_NroFil, 15) = "CLASIFICACION"
      .Range(.Cells(p_int_NroFil, 15), .Cells(p_int_NroFil + 1, 15)).Merge
      .Cells(p_int_NroFil, 16) = "NRO. ENTIDADES"
      .Range(.Cells(p_int_NroFil, 16), .Cells(p_int_NroFil + 1, 16)).Merge
   
      .Cells(p_int_NroFil, 17) = "VENTAS"
      .Range(.Cells(p_int_NroFil, 17), .Cells(p_int_NroFil + 1, 17)).Merge
      .Cells(p_int_NroFil, 18) = "PATRIMONIO"
      .Range(.Cells(p_int_NroFil, 18), .Cells(p_int_NroFil + 1, 18)).Merge
   
      .Cells(p_int_NroFil, 19) = "USUARIO"
      .Range(.Cells(p_int_NroFil, 19), .Cells(p_int_NroFil + 1, 19)).Merge
      
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 19)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 19)).Font.Bold = True
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 19)).HorizontalAlignment = xlHAlignCenter
   
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 16.5
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 15
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 60
      
      .Columns("E").ColumnWidth = 15
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      
      .Columns("F").ColumnWidth = 15
      .Columns("F").NumberFormat = "###,###,###,##0.00"
      .Columns("F").HorizontalAlignment = xlHAlignRight
      .Columns("G").ColumnWidth = 15
      
      .Columns("H").ColumnWidth = 15
      .Columns("H").NumberFormat = "###,###,###,##0.00"
      .Columns("H").HorizontalAlignment = xlHAlignRight
      
      .Columns("I").ColumnWidth = 15
      .Columns("I").NumberFormat = "###,###,###,##0.00"
      .Columns("I").HorizontalAlignment = xlHAlignRight
      .Columns("J").ColumnWidth = 15
      .Columns("J").NumberFormat = "###,###,###,##0.00"
      .Columns("J").HorizontalAlignment = xlHAlignRight
      
      .Columns("K").ColumnWidth = 15
      .Columns("K").NumberFormat = "###,###,###,##0.00"
      .Columns("K").HorizontalAlignment = xlHAlignRight
   
      .Columns("L").ColumnWidth = 15
      .Columns("L").NumberFormat = "###,###,###,##0.00"
      .Columns("L").HorizontalAlignment = xlHAlignRight
      .Columns("M").ColumnWidth = 15
      .Columns("M").NumberFormat = "###,###,###,##0.00"
      .Columns("M").HorizontalAlignment = xlHAlignRight
      .Columns("N").ColumnWidth = 15
      .Columns("N").HorizontalAlignment = xlHAlignCenter
      .Columns("O").ColumnWidth = 25
      .Columns("O").HorizontalAlignment = xlHAlignCenter
      .Columns("P").ColumnWidth = 16
      .Columns("P").HorizontalAlignment = xlHAlignCenter
      .Columns("Q").ColumnWidth = 13.5
      .Columns("Q").NumberFormat = "###,###,###,##0.00"
      .Columns("Q").HorizontalAlignment = xlHAlignRight
      .Columns("R").ColumnWidth = 13.5
      .Columns("R").NumberFormat = "###,###,###,##0.00"
      .Columns("R").HorizontalAlignment = xlHAlignRight
      .Columns("S").ColumnWidth = 13.5
      .Columns("S").HorizontalAlignment = xlHAlignCenter
      
   
      With .Range(.Cells(5, 2), .Cells(6, 19))
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
          .WrapText = True
      End With
   
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(2, 1), .Cells(99, 99)).Font.Size = 11
                      
      p_int_NroFil = p_int_NroFil + 2
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
      End If
      
      If Not IsNull(g_rst_Princi!PATRIMONIO_EFECTIVO) Then
         r_dbl_PatEfe = CDbl(g_rst_Princi!PATRIMONIO_EFECTIVO)
      Else
         r_dbl_PatEfe = 0
      End If
      
      .Cells(6, 5) = r_dbl_PatEfe
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!TIPO = "AD" Then ' If g_rst_Princi!TIPO = 2 Then
            .Cells(p_int_NroFil, 2) = g_rst_Princi!CODIGO_SBS
            .Cells(p_int_NroFil, 3) = g_rst_Princi!DOCUMENTO
            .Cells(p_int_NroFil, 4) = g_rst_Princi!RAZON_SOCIAL
            .Cells(p_int_NroFil, 5) = g_rst_Princi!SUB_PRODUCTO
            If r_dbl_PatEfe = 0 Then
               .Cells(p_int_NroFil, 6) = 0
            Else
               .Cells(p_int_NroFil, 6) = IIf(r_dbl_PatEfe = 0, 0, (CDbl(g_rst_Princi!GARANTIZADO) / CDbl(r_dbl_PatEfe)) * 100)
            End If
            .Cells(p_int_NroFil, 7) = "" 'g_rst_Princi!TIPO_GARANTIA
            .Cells(p_int_NroFil, 8) = Trim(g_rst_Princi!GARANTIZADO)
            .Cells(p_int_NroFil, 9) = g_rst_Princi!MONTO_GARANTIA_LIQUIDA
            .Cells(p_int_NroFil, 10) = g_rst_Princi!MONTO_GARANTIA_HIPOTECARIA
            .Cells(p_int_NroFil, 11) = g_rst_Princi!VALOR_REALIZACION
            .Cells(p_int_NroFil, 12) = g_rst_Princi!DEUDA_SF
            .Cells(p_int_NroFil, 13) = g_rst_Princi!Total
            .Cells(p_int_NroFil, 14) = g_rst_Princi!TIPO_EMPRESA
            .Cells(p_int_NroFil, 15) = g_rst_Princi!CLASIFICACION
            If g_rst_Princi!CLASIFICACION <> "NORMAL" Then
               .Cells(p_int_NroFil, 15).Font.Color = -16776961
               .Cells(p_int_NroFil, 15).Font.Bold = True
            End If
            .Cells(p_int_NroFil, 16) = g_rst_Princi!NRO_ENTIDADES
            .Cells(p_int_NroFil, 17) = g_rst_Princi!VENTAS
            .Cells(p_int_NroFil, 18) = g_rst_Princi!PATRIMONIO
            .Cells(p_int_NroFil, 19) = g_rst_Princi!USUARIO
            p_int_NroFil = p_int_NroFil + 1
         End If
         g_rst_Princi.MoveNext
      Loop
        
      'SUMATORIA TOTAL
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 19)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 19)).Font.Bold = True
      .Range(.Cells(p_int_NroFil, 5), .Cells(p_int_NroFil, 19)).HorizontalAlignment = xlHAlignRight
      
      .Cells(p_int_NroFil, 4) = "TOTAL"
      .Range(.Cells(p_int_NroFil, 8), .Cells(p_int_NroFil, 12)).FormulaR1C1 = "=SUM(R[-" & p_int_NroFil - 5 & "]C:R[-1]C)"
      
      
      'RESUMEN POR TIPO EMPRESA
      .Range(.Cells(p_int_NroFil + 2, 3), .Cells(p_int_NroFil + 2, 7)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil + 2, 3), .Cells(p_int_NroFil + 2, 7)).Font.Bold = True
      .Cells(p_int_NroFil + 2, 3) = "CANTIDAD"
      .Cells(p_int_NroFil + 2, 4) = "DESCRIPCION"
      .Cells(p_int_NroFil + 2, 7) = "GARANTIZADO"
      
      .Range(.Cells(p_int_NroFil + 2, 4), .Cells(p_int_NroFil + 7, 4)).Font.Bold = True
      
      .Cells(p_int_NroFil + 3, 3).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 4) & "]C[11]:R[-4]C[11],""PEQUEÑA"")"
      .Cells(p_int_NroFil + 3, 4) = "RESUMEN PEQUEÑA EMPRESA"
      .Cells(p_int_NroFil + 3, 7).FormulaR1C1 = "=SUMIF(R[-" & (p_int_NroFil - 4) & "]C[7]:R[-4]C[7],""PEQUEÑA"",R[-" & (p_int_NroFil - 4) & "]C[1]:R[-4]C[1])"
      
      .Cells(p_int_NroFil + 4, 3).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 3) & "]C[11]:R[-5]C[11],""MEDIANA"")"
      .Cells(p_int_NroFil + 4, 4) = "RESUMEN MEDIANA EMPRESA"
      .Cells(p_int_NroFil + 4, 7).FormulaR1C1 = "=SUMIF(R[-" & (p_int_NroFil - 3) & "]C[7]:R[-5]C[7],""MEDIANA"",R[-" & (p_int_NroFil - 3) & "]C[1]:R[-5]C[1])"
      
      .Cells(p_int_NroFil + 5, 3).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 2) & "]C[11]:R[-6]C[11],""GRANDE"")"
      .Cells(p_int_NroFil + 5, 4) = "RESUMEN GRANDE EMPRESA"
      .Cells(p_int_NroFil + 5, 7).FormulaR1C1 = "=+SUMIF(R[-" & (p_int_NroFil - 2) & "]C[7]:R[-6]C[7],""GRANDE"",R[-" & (p_int_NroFil - 2) & "]C[1]:R[-6]C[1])"
      
      .Cells(p_int_NroFil + 6, 3).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 1) & "]C[11]:R[-7]C[11],""MICRO"")"
      .Cells(p_int_NroFil + 6, 4) = "RESUMEN MICRO EMPRESA"
      .Cells(p_int_NroFil + 6, 7).FormulaR1C1 = "=+SUMIF(R[-" & (p_int_NroFil - 1) & "]C[7]:R[-7]C[7],""MICRO"",R[-" & (p_int_NroFil - 1) & "]C[1]:R[-7]C[1])"
      
      .Cells(p_int_NroFil + 7, 3).FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
      .Cells(p_int_NroFil + 7, 4) = "TOTAL"
      .Cells(p_int_NroFil + 7, 7).FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
      
      .Range(.Cells(p_int_NroFil + 7, 3), .Cells(p_int_NroFil + 7, 7)).Font.Bold = True
      .Range(.Cells(p_int_NroFil + 3, 7), .Cells(p_int_NroFil + 7, 7)).NumberFormat = "###,###,###,##0.00"
      .Range(.Cells(p_int_NroFil + 3, 7), .Cells(p_int_NroFil + 7, 7)).HorizontalAlignment = xlHAlignRight
      
      With .Range(.Cells(1, 2), .Cells(1, 19))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
      End With
             
      With .Range(.Cells(5, 2), .Cells(p_int_NroFil - 1, 19))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
      With .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 19))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
      With .Range(.Cells(5, 2), .Cells(6, 19))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:    .Borders(xlInsideVertical).Weight = xlThin
          .Borders(xlInsideHorizontal).LineStyle = xlContinuous:  .Borders(xlInsideHorizontal).Weight = xlThin
      End With
      
      With .Range(.Cells(p_int_NroFil + 2, 3), .Cells(p_int_NroFil + 6, 7))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlThin
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlThin
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlThin
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlThin
      End With
      
      With .Range(.Cells(p_int_NroFil + 7, 3), .Cells(p_int_NroFil + 7, 7))
         .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeTop).LineStyle = xlContinuous
         .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeRight).LineStyle = xlContinuous
      End With
   End With
   p_obj_Excel.ActiveSheet.Range("E7").Select
   p_obj_Excel.ActiveWindow.FreezePanes = True
End Sub
Private Sub fs_GenExc_DeudaSF_CSO(ByVal p_obj_Excel As Excel.Application, ByVal p_int_NroFil As Integer)
   
Dim r_int_ConAux        As Integer
Dim r_dbl_PatEfe        As Double
Dim r_bol_Estado        As Boolean

'  p_int_NroFil = 5
   p_obj_Excel.Sheets(6).Select

   With p_obj_Excel.Sheets(6)
   
      .Cells(1, 2) = "REPORTE - SISTEMA FINANCIERO"
      .Range(.Cells(1, 2), .Cells(1, 17)).Merge
      
      If Chk_FecAct.Value = 1 Then
         .Cells(3, 2) = " AL " & Format(date, "DD/MM/YYYY")
      Else
         .Cells(3, 2) = " AL " & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), ipp_PerAno.Text), "00") & " DE " & cmb_PerMes.Text & " DE " & Format(ipp_PerAno.Text, "0000")
      End If
      .Range(.Cells(3, 2), .Cells(3, 17)).Merge
      
      .Range(.Cells(1, 2), .Cells(1, 17)).Font.Bold = True
      .Range(.Cells(1, 2), .Cells(3, 17)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 2), .Cells(1, 17)).Font.Size = 14
   
      .Cells(p_int_NroFil, 2) = "CODIGO SBS"
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 2)).Merge
      .Cells(p_int_NroFil, 3) = "TIPO - DOCUMENTO"
      .Range(.Cells(p_int_NroFil, 3), .Cells(p_int_NroFil + 1, 3)).Merge
      .Cells(p_int_NroFil, 4) = "ENTIDAD"
      .Range(.Cells(p_int_NroFil, 4), .Cells(p_int_NroFil + 1, 4)).Merge
      
      .Cells(p_int_NroFil, 5) = "SUB-PRODUCTO"
      .Range(.Cells(p_int_NroFil, 5), .Cells(p_int_NroFil + 1, 5)).Merge
      
      .Cells(p_int_NroFil, 6) = "PATRIMONIO EFECTIVO"
      .Cells(p_int_NroFil, 7) = "TIPO GARANTIA"
      .Range(.Cells(p_int_NroFil, 7), .Cells(p_int_NroFil + 1, 7)).Merge
      
      .Cells(p_int_NroFil, 8) = "GARANTIZADO"
      .Range(.Cells(p_int_NroFil, 8), .Cells(p_int_NroFil + 1, 8)).Merge
      
      .Cells(p_int_NroFil, 9) = "GARANTIAS"
      .Range(.Cells(p_int_NroFil, 9), .Cells(p_int_NroFil, 11)).Merge
      .Cells(p_int_NroFil + 1, 9) = "LIQUIDA"
      .Cells(p_int_NroFil + 1, 10) = "HIPOTECARIA"
      .Cells(p_int_NroFil + 1, 11) = "V.REALIZACION"
      
      .Cells(p_int_NroFil, 12) = "DEUDA SF"
      .Range(.Cells(p_int_NroFil, 12), .Cells(p_int_NroFil + 1, 12)).Merge
      .Cells(p_int_NroFil, 13) = "TOTAL"
      .Range(.Cells(p_int_NroFil, 13), .Cells(p_int_NroFil + 1, 13)).Merge
   
      .Cells(p_int_NroFil, 14) = "TIPO"
      .Range(.Cells(p_int_NroFil, 14), .Cells(p_int_NroFil + 1, 14)).Merge
      .Cells(p_int_NroFil, 15) = "CLASIFICACION"
      .Range(.Cells(p_int_NroFil, 15), .Cells(p_int_NroFil + 1, 15)).Merge
      .Cells(p_int_NroFil, 16) = "NRO. ENTIDADES"
      .Range(.Cells(p_int_NroFil, 16), .Cells(p_int_NroFil + 1, 16)).Merge
   
      .Cells(p_int_NroFil, 17) = "VENTAS"
      .Range(.Cells(p_int_NroFil, 17), .Cells(p_int_NroFil + 1, 17)).Merge
      .Cells(p_int_NroFil, 18) = "PATRIMONIO"
      .Range(.Cells(p_int_NroFil, 18), .Cells(p_int_NroFil + 1, 18)).Merge
   
      .Cells(p_int_NroFil, 19) = "USUARIO"
      .Range(.Cells(p_int_NroFil, 19), .Cells(p_int_NroFil + 1, 19)).Merge
      
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 19)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 19)).Font.Bold = True
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 19)).HorizontalAlignment = xlHAlignCenter
   
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 16.5
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 15
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 60
      
      .Columns("E").ColumnWidth = 15
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      
      .Columns("F").ColumnWidth = 15
      .Columns("F").NumberFormat = "###,###,###,##0.00"
      .Columns("F").HorizontalAlignment = xlHAlignRight
      .Columns("G").ColumnWidth = 15
      
      .Columns("H").ColumnWidth = 15
      .Columns("H").NumberFormat = "###,###,###,##0.00"
      .Columns("H").HorizontalAlignment = xlHAlignRight
      
      .Columns("I").ColumnWidth = 15
      .Columns("I").NumberFormat = "###,###,###,##0.00"
      .Columns("I").HorizontalAlignment = xlHAlignRight
      .Columns("J").ColumnWidth = 15
      .Columns("J").NumberFormat = "###,###,###,##0.00"
      .Columns("J").HorizontalAlignment = xlHAlignRight
      
      .Columns("K").ColumnWidth = 15
      .Columns("K").NumberFormat = "###,###,###,##0.00"
      .Columns("K").HorizontalAlignment = xlHAlignRight
   
      .Columns("L").ColumnWidth = 15
      .Columns("L").NumberFormat = "###,###,###,##0.00"
      .Columns("L").HorizontalAlignment = xlHAlignRight
      .Columns("M").ColumnWidth = 15
      .Columns("M").NumberFormat = "###,###,###,##0.00"
      .Columns("M").HorizontalAlignment = xlHAlignRight
      .Columns("N").ColumnWidth = 15
      .Columns("N").HorizontalAlignment = xlHAlignCenter
      .Columns("O").ColumnWidth = 25
      .Columns("O").HorizontalAlignment = xlHAlignCenter
      .Columns("P").ColumnWidth = 16
      .Columns("P").HorizontalAlignment = xlHAlignCenter
      .Columns("Q").ColumnWidth = 13.5
      .Columns("Q").NumberFormat = "###,###,###,##0.00"
      .Columns("Q").HorizontalAlignment = xlHAlignRight
      .Columns("R").ColumnWidth = 13.5
      .Columns("R").NumberFormat = "###,###,###,##0.00"
      .Columns("R").HorizontalAlignment = xlHAlignRight
      .Columns("S").ColumnWidth = 13.5
      .Columns("S").HorizontalAlignment = xlHAlignCenter
      
   
      With .Range(.Cells(5, 2), .Cells(6, 19))
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
          .WrapText = True
      End With
   
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(2, 1), .Cells(99, 99)).Font.Size = 11
                      
      p_int_NroFil = p_int_NroFil + 2
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
      End If
      
      If Not IsNull(g_rst_Princi!PATRIMONIO_EFECTIVO) Then
         r_dbl_PatEfe = CDbl(g_rst_Princi!PATRIMONIO_EFECTIVO)
      Else
         r_dbl_PatEfe = 0
      End If
      
      .Cells(6, 5) = r_dbl_PatEfe
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!TIPO = "CSO" Then 'If g_rst_Princi!TIPO = 3 Then
            .Cells(p_int_NroFil, 2) = g_rst_Princi!CODIGO_SBS
            .Cells(p_int_NroFil, 3) = g_rst_Princi!DOCUMENTO
            .Cells(p_int_NroFil, 4) = g_rst_Princi!RAZON_SOCIAL
            .Cells(p_int_NroFil, 5) = g_rst_Princi!SUB_PRODUCTO
            If r_dbl_PatEfe = 0 Then
               .Cells(p_int_NroFil, 6) = 0
            Else
               .Cells(p_int_NroFil, 6) = IIf(r_dbl_PatEfe = 0, 0, (CDbl(g_rst_Princi!GARANTIZADO) / CDbl(r_dbl_PatEfe)) * 100)
            End If
            .Cells(p_int_NroFil, 7) = "" 'g_rst_Princi!TIPO_GARANTIA
            .Cells(p_int_NroFil, 8) = Trim(g_rst_Princi!GARANTIZADO)
            .Cells(p_int_NroFil, 9) = g_rst_Princi!MONTO_GARANTIA_LIQUIDA
            .Cells(p_int_NroFil, 10) = g_rst_Princi!MONTO_GARANTIA_HIPOTECARIA
            .Cells(p_int_NroFil, 11) = g_rst_Princi!VALOR_REALIZACION
            .Cells(p_int_NroFil, 12) = g_rst_Princi!DEUDA_SF
            .Cells(p_int_NroFil, 13) = g_rst_Princi!Total
            .Cells(p_int_NroFil, 14) = g_rst_Princi!TIPO_EMPRESA
            .Cells(p_int_NroFil, 15) = g_rst_Princi!CLASIFICACION
            If g_rst_Princi!CLASIFICACION <> "NORMAL" Then
               .Cells(p_int_NroFil, 15).Font.Color = -16776961
               .Cells(p_int_NroFil, 15).Font.Bold = True
            End If
            .Cells(p_int_NroFil, 16) = g_rst_Princi!NRO_ENTIDADES
            .Cells(p_int_NroFil, 17) = g_rst_Princi!VENTAS
            .Cells(p_int_NroFil, 18) = g_rst_Princi!PATRIMONIO
            .Cells(p_int_NroFil, 19) = g_rst_Princi!USUARIO
            p_int_NroFil = p_int_NroFil + 1
         End If
         g_rst_Princi.MoveNext
      Loop
        
      'SUMATORIA TOTAL
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 19)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 19)).Font.Bold = True
      .Range(.Cells(p_int_NroFil, 5), .Cells(p_int_NroFil, 19)).HorizontalAlignment = xlHAlignRight
      
      .Cells(p_int_NroFil, 4) = "TOTAL"
      .Range(.Cells(p_int_NroFil, 8), .Cells(p_int_NroFil, 12)).FormulaR1C1 = "=SUM(R[-" & p_int_NroFil - 5 & "]C:R[-1]C)"
      
      
      'RESUMEN POR TIPO EMPRESA
      .Range(.Cells(p_int_NroFil + 2, 3), .Cells(p_int_NroFil + 2, 7)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil + 2, 3), .Cells(p_int_NroFil + 2, 7)).Font.Bold = True
      .Cells(p_int_NroFil + 2, 3) = "CANTIDAD"
      .Cells(p_int_NroFil + 2, 4) = "DESCRIPCION"
      .Cells(p_int_NroFil + 2, 7) = "GARANTIZADO"
      
      .Range(.Cells(p_int_NroFil + 2, 4), .Cells(p_int_NroFil + 7, 4)).Font.Bold = True
      
      .Cells(p_int_NroFil + 3, 3).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 4) & "]C[11]:R[-4]C[11],""PEQUEÑA"")"
      .Cells(p_int_NroFil + 3, 4) = "RESUMEN PEQUEÑA EMPRESA"
      .Cells(p_int_NroFil + 3, 7).FormulaR1C1 = "=SUMIF(R[-" & (p_int_NroFil - 4) & "]C[7]:R[-4]C[7],""PEQUEÑA"",R[-" & (p_int_NroFil - 4) & "]C[1]:R[-4]C[1])"
      
      .Cells(p_int_NroFil + 4, 3).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 3) & "]C[11]:R[-5]C[11],""MEDIANA"")"
      .Cells(p_int_NroFil + 4, 4) = "RESUMEN MEDIANA EMPRESA"
      .Cells(p_int_NroFil + 4, 7).FormulaR1C1 = "=SUMIF(R[-" & (p_int_NroFil - 3) & "]C[7]:R[-5]C[7],""MEDIANA"",R[-" & (p_int_NroFil - 3) & "]C[1]:R[-5]C[1])"
      
      .Cells(p_int_NroFil + 5, 3).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 2) & "]C[11]:R[-6]C[11],""GRANDE"")"
      .Cells(p_int_NroFil + 5, 4) = "RESUMEN GRANDE EMPRESA"
      .Cells(p_int_NroFil + 5, 7).FormulaR1C1 = "=+SUMIF(R[-" & (p_int_NroFil - 2) & "]C[7]:R[-6]C[7],""GRANDE"",R[-" & (p_int_NroFil - 2) & "]C[1]:R[-6]C[1])"
      
      .Cells(p_int_NroFil + 6, 3).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 1) & "]C[11]:R[-7]C[11],""MICRO"")"
      .Cells(p_int_NroFil + 6, 4) = "RESUMEN MICRO EMPRESA"
      .Cells(p_int_NroFil + 6, 7).FormulaR1C1 = "=+SUMIF(R[-" & (p_int_NroFil - 1) & "]C[7]:R[-7]C[7],""MICRO"",R[-" & (p_int_NroFil - 1) & "]C[1]:R[-7]C[1])"
      
      .Cells(p_int_NroFil + 7, 3).FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
      .Cells(p_int_NroFil + 7, 4) = "TOTAL"
      .Cells(p_int_NroFil + 7, 7).FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
      
      .Range(.Cells(p_int_NroFil + 7, 3), .Cells(p_int_NroFil + 7, 7)).Font.Bold = True
      .Range(.Cells(p_int_NroFil + 3, 7), .Cells(p_int_NroFil + 7, 7)).NumberFormat = "###,###,###,##0.00"
      .Range(.Cells(p_int_NroFil + 3, 7), .Cells(p_int_NroFil + 7, 7)).HorizontalAlignment = xlHAlignRight
      
      With .Range(.Cells(1, 2), .Cells(1, 19))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
      End With
             
      With .Range(.Cells(5, 2), .Cells(p_int_NroFil - 1, 19))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
      With .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 19))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
      With .Range(.Cells(5, 2), .Cells(6, 19))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:    .Borders(xlInsideVertical).Weight = xlThin
          .Borders(xlInsideHorizontal).LineStyle = xlContinuous:  .Borders(xlInsideHorizontal).Weight = xlThin
      End With
      
      With .Range(.Cells(p_int_NroFil + 2, 3), .Cells(p_int_NroFil + 6, 7))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlThin
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlThin
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlThin
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlThin
      End With
      
      With .Range(.Cells(p_int_NroFil + 7, 3), .Cells(p_int_NroFil + 7, 7))
         .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeTop).LineStyle = xlContinuous
         .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeRight).LineStyle = xlContinuous
      End With
   End With
   p_obj_Excel.ActiveSheet.Range("E7").Select
   p_obj_Excel.ActiveWindow.FreezePanes = True
End Sub
Private Sub fs_GenExc_DeudaSF_CDir_LC(ByVal p_obj_Excel As Excel.Application, ByVal p_int_NroFil As Integer)
   
Dim r_int_ConAux        As Integer
Dim r_dbl_PatEfe        As Double
Dim r_bol_Estado        As Boolean

'  p_int_NroFil = 5
   p_obj_Excel.Sheets(8).Select

   With p_obj_Excel.Sheets(8)
   
      .Cells(1, 2) = "REPORTE - SISTEMA FINANCIERO"
      .Range(.Cells(1, 2), .Cells(1, 17)).Merge
      
      If Chk_FecAct.Value = 1 Then
         .Cells(3, 2) = " AL " & Format(date, "DD/MM/YYYY")
      Else
         .Cells(3, 2) = " AL " & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), ipp_PerAno.Text), "00") & " DE " & cmb_PerMes.Text & " DE " & Format(ipp_PerAno.Text, "0000")
      End If
      .Range(.Cells(3, 2), .Cells(3, 17)).Merge
      
      .Range(.Cells(1, 2), .Cells(1, 17)).Font.Bold = True
      .Range(.Cells(1, 2), .Cells(3, 17)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 2), .Cells(1, 17)).Font.Size = 14
   
      .Cells(p_int_NroFil, 2) = "CODIGO SBS"
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 2)).Merge
      .Cells(p_int_NroFil, 3) = "TIPO - DOCUMENTO"
      .Range(.Cells(p_int_NroFil, 3), .Cells(p_int_NroFil + 1, 3)).Merge
      .Cells(p_int_NroFil, 4) = "ENTIDAD"
      .Range(.Cells(p_int_NroFil, 4), .Cells(p_int_NroFil + 1, 4)).Merge
      
      .Cells(p_int_NroFil, 5) = "SUB-PRODUCTO"
      .Range(.Cells(p_int_NroFil, 5), .Cells(p_int_NroFil + 1, 5)).Merge
      
      .Cells(p_int_NroFil, 6) = "PATRIMONIO EFECTIVO"
      .Cells(p_int_NroFil, 7) = "TIPO GARANTIA"
      .Range(.Cells(p_int_NroFil, 7), .Cells(p_int_NroFil + 1, 7)).Merge
      
      .Cells(p_int_NroFil, 8) = "MONTO PRESTAMO" '"GARANTIZADO"
      .Range(.Cells(p_int_NroFil, 8), .Cells(p_int_NroFil + 1, 8)).Merge
      
      .Cells(p_int_NroFil, 9) = "GARANTIAS"
      .Range(.Cells(p_int_NroFil, 9), .Cells(p_int_NroFil, 11)).Merge
      .Cells(p_int_NroFil + 1, 9) = "LIQUIDA"
      .Cells(p_int_NroFil + 1, 10) = "HIPOTECARIA"
      .Cells(p_int_NroFil + 1, 11) = "V.REALIZACION"
      
      .Cells(p_int_NroFil, 12) = "DEUDA SF"
      .Range(.Cells(p_int_NroFil, 12), .Cells(p_int_NroFil + 1, 12)).Merge
      .Cells(p_int_NroFil, 13) = "TOTAL"
      .Range(.Cells(p_int_NroFil, 13), .Cells(p_int_NroFil + 1, 13)).Merge
   
      .Cells(p_int_NroFil, 14) = "TIPO"
      .Range(.Cells(p_int_NroFil, 14), .Cells(p_int_NroFil + 1, 14)).Merge
      .Cells(p_int_NroFil, 15) = "CLASIFICACION"
      .Range(.Cells(p_int_NroFil, 15), .Cells(p_int_NroFil + 1, 15)).Merge
      .Cells(p_int_NroFil, 16) = "NRO. ENTIDADES"
      .Range(.Cells(p_int_NroFil, 16), .Cells(p_int_NroFil + 1, 16)).Merge
   
      .Cells(p_int_NroFil, 17) = "VENTAS"
      .Range(.Cells(p_int_NroFil, 17), .Cells(p_int_NroFil + 1, 17)).Merge
      .Cells(p_int_NroFil, 18) = "PATRIMONIO"
      .Range(.Cells(p_int_NroFil, 18), .Cells(p_int_NroFil + 1, 18)).Merge
   
      .Cells(p_int_NroFil, 19) = "USUARIO"
      .Range(.Cells(p_int_NroFil, 19), .Cells(p_int_NroFil + 1, 19)).Merge
      
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 19)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 19)).Font.Bold = True
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 19)).HorizontalAlignment = xlHAlignCenter
   
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 16.5
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 15
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 60
      
      .Columns("E").ColumnWidth = 15
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      
      .Columns("F").ColumnWidth = 15
      .Columns("F").NumberFormat = "###,###,###,##0.00"
      .Columns("F").HorizontalAlignment = xlHAlignRight
      .Columns("G").ColumnWidth = 15
      
      .Columns("H").ColumnWidth = 15
      .Columns("H").NumberFormat = "###,###,###,##0.00"
      .Columns("H").HorizontalAlignment = xlHAlignRight
      
      .Columns("I").ColumnWidth = 15
      .Columns("I").NumberFormat = "###,###,###,##0.00"
      .Columns("I").HorizontalAlignment = xlHAlignRight
      .Columns("J").ColumnWidth = 15
      .Columns("J").NumberFormat = "###,###,###,##0.00"
      .Columns("J").HorizontalAlignment = xlHAlignRight
      
      .Columns("K").ColumnWidth = 15
      .Columns("K").NumberFormat = "###,###,###,##0.00"
      .Columns("K").HorizontalAlignment = xlHAlignRight
   
      .Columns("L").ColumnWidth = 15
      .Columns("L").NumberFormat = "###,###,###,##0.00"
      .Columns("L").HorizontalAlignment = xlHAlignRight
      .Columns("M").ColumnWidth = 15
      .Columns("M").NumberFormat = "###,###,###,##0.00"
      .Columns("M").HorizontalAlignment = xlHAlignRight
      .Columns("N").ColumnWidth = 15
      .Columns("N").HorizontalAlignment = xlHAlignCenter
      .Columns("O").ColumnWidth = 25
      .Columns("O").HorizontalAlignment = xlHAlignCenter
      .Columns("P").ColumnWidth = 16
      .Columns("P").HorizontalAlignment = xlHAlignCenter
      .Columns("Q").ColumnWidth = 13.5
      .Columns("Q").NumberFormat = "###,###,###,##0.00"
      .Columns("Q").HorizontalAlignment = xlHAlignRight
      .Columns("R").ColumnWidth = 13.5
      .Columns("R").NumberFormat = "###,###,###,##0.00"
      .Columns("R").HorizontalAlignment = xlHAlignRight
      .Columns("S").ColumnWidth = 13.5
      .Columns("S").HorizontalAlignment = xlHAlignCenter
      
   
      With .Range(.Cells(5, 2), .Cells(6, 19))
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
          .WrapText = True
      End With
   
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(2, 1), .Cells(99, 99)).Font.Size = 11
                      
      p_int_NroFil = p_int_NroFil + 2
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
      End If
      
      If Not IsNull(g_rst_Princi!PATRIMONIO_EFECTIVO) Then
         r_dbl_PatEfe = CDbl(g_rst_Princi!PATRIMONIO_EFECTIVO)
      Else
         r_dbl_PatEfe = 0
      End If
      
      .Cells(6, 5) = r_dbl_PatEfe
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!TIPO = "LC" Then ' If g_rst_Princi!TIPO = 0 And InStr(g_rst_Princi!SUB_PRODUCTO, "CDIR-LC") > 0 Then '
            .Cells(p_int_NroFil, 2) = g_rst_Princi!CODIGO_SBS
            .Cells(p_int_NroFil, 3) = g_rst_Princi!DOCUMENTO
            .Cells(p_int_NroFil, 4) = g_rst_Princi!RAZON_SOCIAL
            .Cells(p_int_NroFil, 5) = g_rst_Princi!SUB_PRODUCTO
            If r_dbl_PatEfe = 0 Then
               .Cells(p_int_NroFil, 6) = 0
            Else
               .Cells(p_int_NroFil, 6) = IIf(r_dbl_PatEfe = 0, 0, (CDbl(g_rst_Princi!IMPORTE) / CDbl(r_dbl_PatEfe)) * 100)
            End If
            .Cells(p_int_NroFil, 7) = "" 'g_rst_Princi!TIPO_GARANTIA
            .Cells(p_int_NroFil, 8) = Trim(g_rst_Princi!IMPORTE) 'GARANTIZADO
            .Cells(p_int_NroFil, 9) = g_rst_Princi!MONTO_GARANTIA_LIQUIDA
            .Cells(p_int_NroFil, 10) = g_rst_Princi!MONTO_GARANTIA_HIPOTECARIA
            .Cells(p_int_NroFil, 11) = g_rst_Princi!VALOR_REALIZACION
            .Cells(p_int_NroFil, 12) = g_rst_Princi!DEUDA_SF
            .Cells(p_int_NroFil, 13) = g_rst_Princi!Total
            .Cells(p_int_NroFil, 14) = g_rst_Princi!TIPO_EMPRESA
            .Cells(p_int_NroFil, 15) = g_rst_Princi!CLASIFICACION
            If g_rst_Princi!CLASIFICACION <> "NORMAL" Then
               .Cells(p_int_NroFil, 15).Font.Color = -16776961
               .Cells(p_int_NroFil, 15).Font.Bold = True
            End If
            .Cells(p_int_NroFil, 16) = g_rst_Princi!NRO_ENTIDADES
            .Cells(p_int_NroFil, 17) = g_rst_Princi!VENTAS
            .Cells(p_int_NroFil, 18) = g_rst_Princi!PATRIMONIO
            .Cells(p_int_NroFil, 19) = g_rst_Princi!USUARIO
            p_int_NroFil = p_int_NroFil + 1
         End If
         g_rst_Princi.MoveNext
      Loop
        
      'SUMATORIA TOTAL
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 19)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 19)).Font.Bold = True
      .Range(.Cells(p_int_NroFil, 5), .Cells(p_int_NroFil, 19)).HorizontalAlignment = xlHAlignRight
      
      .Cells(p_int_NroFil, 4) = "TOTAL"
      .Range(.Cells(p_int_NroFil, 8), .Cells(p_int_NroFil, 12)).FormulaR1C1 = "=SUM(R[-" & p_int_NroFil - 5 & "]C:R[-1]C)"
      
      
      'RESUMEN POR TIPO EMPRESA
      .Range(.Cells(p_int_NroFil + 2, 3), .Cells(p_int_NroFil + 2, 7)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil + 2, 3), .Cells(p_int_NroFil + 2, 7)).Font.Bold = True
      .Cells(p_int_NroFil + 2, 3) = "CANTIDAD"
      .Cells(p_int_NroFil + 2, 4) = "DESCRIPCION"
      .Cells(p_int_NroFil + 2, 7) = "GARANTIZADO"
      
      .Range(.Cells(p_int_NroFil + 2, 4), .Cells(p_int_NroFil + 7, 4)).Font.Bold = True
      
      .Cells(p_int_NroFil + 3, 3).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 4) & "]C[11]:R[-4]C[11],""PEQUEÑA"")"
      .Cells(p_int_NroFil + 3, 4) = "RESUMEN PEQUEÑA EMPRESA"
      .Cells(p_int_NroFil + 3, 7).FormulaR1C1 = "=SUMIF(R[-" & (p_int_NroFil - 4) & "]C[7]:R[-4]C[7],""PEQUEÑA"",R[-" & (p_int_NroFil - 4) & "]C[1]:R[-4]C[1])"
      
      .Cells(p_int_NroFil + 4, 3).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 3) & "]C[11]:R[-5]C[11],""MEDIANA"")"
      .Cells(p_int_NroFil + 4, 4) = "RESUMEN MEDIANA EMPRESA"
      .Cells(p_int_NroFil + 4, 7).FormulaR1C1 = "=SUMIF(R[-" & (p_int_NroFil - 3) & "]C[7]:R[-5]C[7],""MEDIANA"",R[-" & (p_int_NroFil - 3) & "]C[1]:R[-5]C[1])"
      
      .Cells(p_int_NroFil + 5, 3).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 2) & "]C[11]:R[-6]C[11],""GRANDE"")"
      .Cells(p_int_NroFil + 5, 4) = "RESUMEN GRANDE EMPRESA"
      .Cells(p_int_NroFil + 5, 7).FormulaR1C1 = "=+SUMIF(R[-" & (p_int_NroFil - 2) & "]C[7]:R[-6]C[7],""GRANDE"",R[-" & (p_int_NroFil - 2) & "]C[1]:R[-6]C[1])"
      
      .Cells(p_int_NroFil + 6, 3).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 1) & "]C[11]:R[-7]C[11],""MICRO"")"
      .Cells(p_int_NroFil + 6, 4) = "RESUMEN MICRO EMPRESA"
      .Cells(p_int_NroFil + 6, 7).FormulaR1C1 = "=+SUMIF(R[-" & (p_int_NroFil - 1) & "]C[7]:R[-7]C[7],""MICRO"",R[-" & (p_int_NroFil - 1) & "]C[1]:R[-7]C[1])"
      
      .Cells(p_int_NroFil + 7, 3).FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
      .Cells(p_int_NroFil + 7, 4) = "TOTAL"
      .Cells(p_int_NroFil + 7, 7).FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
      
      .Range(.Cells(p_int_NroFil + 7, 3), .Cells(p_int_NroFil + 7, 7)).Font.Bold = True
      .Range(.Cells(p_int_NroFil + 3, 7), .Cells(p_int_NroFil + 7, 7)).NumberFormat = "###,###,###,##0.00"
      .Range(.Cells(p_int_NroFil + 3, 7), .Cells(p_int_NroFil + 7, 7)).HorizontalAlignment = xlHAlignRight
      
      With .Range(.Cells(1, 2), .Cells(1, 19))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
      End With
             
      With .Range(.Cells(5, 2), .Cells(p_int_NroFil - 1, 19))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
      With .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 19))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
      With .Range(.Cells(5, 2), .Cells(6, 19))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:    .Borders(xlInsideVertical).Weight = xlThin
          .Borders(xlInsideHorizontal).LineStyle = xlContinuous:  .Borders(xlInsideHorizontal).Weight = xlThin
      End With
      
      With .Range(.Cells(p_int_NroFil + 2, 3), .Cells(p_int_NroFil + 6, 7))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlThin
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlThin
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlThin
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlThin
      End With
      
      With .Range(.Cells(p_int_NroFil + 7, 3), .Cells(p_int_NroFil + 7, 7))
         .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeTop).LineStyle = xlContinuous
         .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeRight).LineStyle = xlContinuous
      End With
   End With
   p_obj_Excel.ActiveSheet.Range("E7").Select
   p_obj_Excel.ActiveWindow.FreezePanes = True
End Sub
Private Sub fs_GenExc_DeudaSF_CDir_CP(ByVal p_obj_Excel As Excel.Application, ByVal p_int_NroFil As Integer)
   
Dim r_int_ConAux        As Integer
Dim r_dbl_PatEfe        As Double
Dim r_bol_Estado        As Boolean

'  p_int_NroFil = 5
   p_obj_Excel.Sheets(9).Select

   With p_obj_Excel.Sheets(9)
   
      .Cells(1, 2) = "REPORTE - SISTEMA FINANCIERO"
      .Range(.Cells(1, 2), .Cells(1, 17)).Merge
      
      If Chk_FecAct.Value = 1 Then
         .Cells(3, 2) = " AL " & Format(date, "DD/MM/YYYY")
      Else
         .Cells(3, 2) = " AL " & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), ipp_PerAno.Text), "00") & " DE " & cmb_PerMes.Text & " DE " & Format(ipp_PerAno.Text, "0000")
      End If
      .Range(.Cells(3, 2), .Cells(3, 17)).Merge
      
      .Range(.Cells(1, 2), .Cells(1, 17)).Font.Bold = True
      .Range(.Cells(1, 2), .Cells(3, 17)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 2), .Cells(1, 17)).Font.Size = 14
   
      .Cells(p_int_NroFil, 2) = "CODIGO SBS"
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 2)).Merge
      .Cells(p_int_NroFil, 3) = "TIPO - DOCUMENTO"
      .Range(.Cells(p_int_NroFil, 3), .Cells(p_int_NroFil + 1, 3)).Merge
      .Cells(p_int_NroFil, 4) = "ENTIDAD"
      .Range(.Cells(p_int_NroFil, 4), .Cells(p_int_NroFil + 1, 4)).Merge
      
      .Cells(p_int_NroFil, 5) = "SUB-PRODUCTO"
      .Range(.Cells(p_int_NroFil, 5), .Cells(p_int_NroFil + 1, 5)).Merge
      
      .Cells(p_int_NroFil, 6) = "PATRIMONIO EFECTIVO"
      .Cells(p_int_NroFil, 7) = "TIPO GARANTIA"
      .Range(.Cells(p_int_NroFil, 7), .Cells(p_int_NroFil + 1, 7)).Merge
      
      .Cells(p_int_NroFil, 8) = "MONTO PRESTAMO" '"GARANTIZADO"
      .Range(.Cells(p_int_NroFil, 8), .Cells(p_int_NroFil + 1, 8)).Merge
      
      .Cells(p_int_NroFil, 9) = "GARANTIAS"
      .Range(.Cells(p_int_NroFil, 9), .Cells(p_int_NroFil, 11)).Merge
      .Cells(p_int_NroFil + 1, 9) = "LIQUIDA"
      .Cells(p_int_NroFil + 1, 10) = "HIPOTECARIA"
      .Cells(p_int_NroFil + 1, 11) = "V.REALIZACION"
      
      .Cells(p_int_NroFil, 12) = "DEUDA SF"
      .Range(.Cells(p_int_NroFil, 12), .Cells(p_int_NroFil + 1, 12)).Merge
      .Cells(p_int_NroFil, 13) = "TOTAL"
      .Range(.Cells(p_int_NroFil, 13), .Cells(p_int_NroFil + 1, 13)).Merge
   
      .Cells(p_int_NroFil, 14) = "TIPO"
      .Range(.Cells(p_int_NroFil, 14), .Cells(p_int_NroFil + 1, 14)).Merge
      .Cells(p_int_NroFil, 15) = "CLASIFICACION"
      .Range(.Cells(p_int_NroFil, 15), .Cells(p_int_NroFil + 1, 15)).Merge
      .Cells(p_int_NroFil, 16) = "NRO. ENTIDADES"
      .Range(.Cells(p_int_NroFil, 16), .Cells(p_int_NroFil + 1, 16)).Merge
   
      .Cells(p_int_NroFil, 17) = "VENTAS"
      .Range(.Cells(p_int_NroFil, 17), .Cells(p_int_NroFil + 1, 17)).Merge
      .Cells(p_int_NroFil, 18) = "PATRIMONIO"
      .Range(.Cells(p_int_NroFil, 18), .Cells(p_int_NroFil + 1, 18)).Merge
   
      .Cells(p_int_NroFil, 19) = "USUARIO"
      .Range(.Cells(p_int_NroFil, 19), .Cells(p_int_NroFil + 1, 19)).Merge
      
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 19)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 19)).Font.Bold = True
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 19)).HorizontalAlignment = xlHAlignCenter
   
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 16.5
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 15
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 60
      
      .Columns("E").ColumnWidth = 15
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      
      .Columns("F").ColumnWidth = 15
      .Columns("F").NumberFormat = "###,###,###,##0.00"
      .Columns("F").HorizontalAlignment = xlHAlignRight
      .Columns("G").ColumnWidth = 15
      
      .Columns("H").ColumnWidth = 15
      .Columns("H").NumberFormat = "###,###,###,##0.00"
      .Columns("H").HorizontalAlignment = xlHAlignRight
      
      .Columns("I").ColumnWidth = 15
      .Columns("I").NumberFormat = "###,###,###,##0.00"
      .Columns("I").HorizontalAlignment = xlHAlignRight
      .Columns("J").ColumnWidth = 15
      .Columns("J").NumberFormat = "###,###,###,##0.00"
      .Columns("J").HorizontalAlignment = xlHAlignRight
      
      .Columns("K").ColumnWidth = 15
      .Columns("K").NumberFormat = "###,###,###,##0.00"
      .Columns("K").HorizontalAlignment = xlHAlignRight
   
      .Columns("L").ColumnWidth = 15
      .Columns("L").NumberFormat = "###,###,###,##0.00"
      .Columns("L").HorizontalAlignment = xlHAlignRight
      .Columns("M").ColumnWidth = 15
      .Columns("M").NumberFormat = "###,###,###,##0.00"
      .Columns("M").HorizontalAlignment = xlHAlignRight
      .Columns("N").ColumnWidth = 15
      .Columns("N").HorizontalAlignment = xlHAlignCenter
      .Columns("O").ColumnWidth = 25
      .Columns("O").HorizontalAlignment = xlHAlignCenter
      .Columns("P").ColumnWidth = 16
      .Columns("P").HorizontalAlignment = xlHAlignCenter
      .Columns("Q").ColumnWidth = 13.5
      .Columns("Q").NumberFormat = "###,###,###,##0.00"
      .Columns("Q").HorizontalAlignment = xlHAlignRight
      .Columns("R").ColumnWidth = 13.5
      .Columns("R").NumberFormat = "###,###,###,##0.00"
      .Columns("R").HorizontalAlignment = xlHAlignRight
      .Columns("S").ColumnWidth = 13.5
      .Columns("S").HorizontalAlignment = xlHAlignCenter
      
   
      With .Range(.Cells(5, 2), .Cells(6, 19))
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
          .WrapText = True
      End With
   
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(2, 1), .Cells(99, 99)).Font.Size = 11
                      
      p_int_NroFil = p_int_NroFil + 2
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
      End If
      
      If Not IsNull(g_rst_Princi!PATRIMONIO_EFECTIVO) Then
         r_dbl_PatEfe = CDbl(g_rst_Princi!PATRIMONIO_EFECTIVO)
      Else
         r_dbl_PatEfe = 0
      End If
      
      .Cells(6, 5) = r_dbl_PatEfe
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!TIPO = "CP" Then 'If g_rst_Princi!TIPO = 0 And InStr(g_rst_Princi!SUB_PRODUCTO, "CDIR-CP") > 0 Then
            .Cells(p_int_NroFil, 2) = g_rst_Princi!CODIGO_SBS
            .Cells(p_int_NroFil, 3) = g_rst_Princi!DOCUMENTO
            .Cells(p_int_NroFil, 4) = g_rst_Princi!RAZON_SOCIAL
            .Cells(p_int_NroFil, 5) = g_rst_Princi!SUB_PRODUCTO
            If r_dbl_PatEfe = 0 Then
               .Cells(p_int_NroFil, 6) = 0
            Else
               .Cells(p_int_NroFil, 6) = IIf(r_dbl_PatEfe = 0, 0, (CDbl(g_rst_Princi!IMPORTE) / CDbl(r_dbl_PatEfe)) * 100)
            End If
            .Cells(p_int_NroFil, 7) = "" 'g_rst_Princi!TIPO_GARANTIA
            .Cells(p_int_NroFil, 8) = Trim(g_rst_Princi!IMPORTE) 'GARANTIZADO
            .Cells(p_int_NroFil, 9) = g_rst_Princi!MONTO_GARANTIA_LIQUIDA
            .Cells(p_int_NroFil, 10) = g_rst_Princi!MONTO_GARANTIA_HIPOTECARIA
            .Cells(p_int_NroFil, 11) = g_rst_Princi!VALOR_REALIZACION
            .Cells(p_int_NroFil, 12) = g_rst_Princi!DEUDA_SF
            .Cells(p_int_NroFil, 13) = g_rst_Princi!Total
            .Cells(p_int_NroFil, 14) = g_rst_Princi!TIPO_EMPRESA
            .Cells(p_int_NroFil, 15) = g_rst_Princi!CLASIFICACION
            If g_rst_Princi!CLASIFICACION <> "NORMAL" Then
               .Cells(p_int_NroFil, 15).Font.Color = -16776961
               .Cells(p_int_NroFil, 15).Font.Bold = True
            End If
            .Cells(p_int_NroFil, 16) = g_rst_Princi!NRO_ENTIDADES
            .Cells(p_int_NroFil, 17) = g_rst_Princi!VENTAS
            .Cells(p_int_NroFil, 18) = g_rst_Princi!PATRIMONIO
            .Cells(p_int_NroFil, 19) = g_rst_Princi!USUARIO
            p_int_NroFil = p_int_NroFil + 1
         End If
         g_rst_Princi.MoveNext
      Loop
        
      'SUMATORIA TOTAL
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 19)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 19)).Font.Bold = True
      .Range(.Cells(p_int_NroFil, 5), .Cells(p_int_NroFil, 19)).HorizontalAlignment = xlHAlignRight
      
      .Cells(p_int_NroFil, 4) = "TOTAL"
      .Range(.Cells(p_int_NroFil, 8), .Cells(p_int_NroFil, 12)).FormulaR1C1 = "=SUM(R[-" & p_int_NroFil - 5 & "]C:R[-1]C)"
      
      
      'RESUMEN POR TIPO EMPRESA
      .Range(.Cells(p_int_NroFil + 2, 3), .Cells(p_int_NroFil + 2, 7)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil + 2, 3), .Cells(p_int_NroFil + 2, 7)).Font.Bold = True
      .Cells(p_int_NroFil + 2, 3) = "CANTIDAD"
      .Cells(p_int_NroFil + 2, 4) = "DESCRIPCION"
      .Cells(p_int_NroFil + 2, 7) = "GARANTIZADO"
      
      .Range(.Cells(p_int_NroFil + 2, 4), .Cells(p_int_NroFil + 7, 4)).Font.Bold = True
      
      .Cells(p_int_NroFil + 3, 3).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 4) & "]C[11]:R[-4]C[11],""PEQUEÑA"")"
      .Cells(p_int_NroFil + 3, 4) = "RESUMEN PEQUEÑA EMPRESA"
      .Cells(p_int_NroFil + 3, 7).FormulaR1C1 = "=SUMIF(R[-" & (p_int_NroFil - 4) & "]C[7]:R[-4]C[7],""PEQUEÑA"",R[-" & (p_int_NroFil - 4) & "]C[1]:R[-4]C[1])"
      
      .Cells(p_int_NroFil + 4, 3).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 3) & "]C[11]:R[-5]C[11],""MEDIANA"")"
      .Cells(p_int_NroFil + 4, 4) = "RESUMEN MEDIANA EMPRESA"
      .Cells(p_int_NroFil + 4, 7).FormulaR1C1 = "=SUMIF(R[-" & (p_int_NroFil - 3) & "]C[7]:R[-5]C[7],""MEDIANA"",R[-" & (p_int_NroFil - 3) & "]C[1]:R[-5]C[1])"
      
      .Cells(p_int_NroFil + 5, 3).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 2) & "]C[11]:R[-6]C[11],""GRANDE"")"
      .Cells(p_int_NroFil + 5, 4) = "RESUMEN GRANDE EMPRESA"
      .Cells(p_int_NroFil + 5, 7).FormulaR1C1 = "=+SUMIF(R[-" & (p_int_NroFil - 2) & "]C[7]:R[-6]C[7],""GRANDE"",R[-" & (p_int_NroFil - 2) & "]C[1]:R[-6]C[1])"
      
      .Cells(p_int_NroFil + 6, 3).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 1) & "]C[11]:R[-7]C[11],""MICRO"")"
      .Cells(p_int_NroFil + 6, 4) = "RESUMEN MICRO EMPRESA"
      .Cells(p_int_NroFil + 6, 7).FormulaR1C1 = "=+SUMIF(R[-" & (p_int_NroFil - 1) & "]C[7]:R[-7]C[7],""MICRO"",R[-" & (p_int_NroFil - 1) & "]C[1]:R[-7]C[1])"
      
      .Cells(p_int_NroFil + 7, 3).FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
      .Cells(p_int_NroFil + 7, 4) = "TOTAL"
      .Cells(p_int_NroFil + 7, 7).FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
      
      .Range(.Cells(p_int_NroFil + 7, 3), .Cells(p_int_NroFil + 7, 7)).Font.Bold = True
      .Range(.Cells(p_int_NroFil + 3, 7), .Cells(p_int_NroFil + 7, 7)).NumberFormat = "###,###,###,##0.00"
      .Range(.Cells(p_int_NroFil + 3, 7), .Cells(p_int_NroFil + 7, 7)).HorizontalAlignment = xlHAlignRight
      
      With .Range(.Cells(1, 2), .Cells(1, 19))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
      End With
             
      With .Range(.Cells(5, 2), .Cells(p_int_NroFil - 1, 19))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
      With .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 19))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
      With .Range(.Cells(5, 2), .Cells(6, 19))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:    .Borders(xlInsideVertical).Weight = xlThin
          .Borders(xlInsideHorizontal).LineStyle = xlContinuous:  .Borders(xlInsideHorizontal).Weight = xlThin
      End With
      
      With .Range(.Cells(p_int_NroFil + 2, 3), .Cells(p_int_NroFil + 6, 7))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlThin
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlThin
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlThin
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlThin
      End With
      
      With .Range(.Cells(p_int_NroFil + 7, 3), .Cells(p_int_NroFil + 7, 7))
         .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeTop).LineStyle = xlContinuous
         .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeRight).LineStyle = xlContinuous
      End With
   End With
   p_obj_Excel.ActiveSheet.Range("E7").Select
   p_obj_Excel.ActiveWindow.FreezePanes = True
End Sub
 '*************************************************************** REPORTE GENERAL - CARTAS FIANZA **************************************************************
Private Sub fs_GenExc_TechoPropio_CF(ByVal p_obj_Excel As Excel.Application, ByVal p_int_NroFil As Integer)
Dim r_int_ConAux        As Integer
Dim r_bol_Estado        As Boolean
'Cartas:
 '  p_int_NroFil = 5
   p_obj_Excel.Sheets(1).Select
   
   With p_obj_Excel.Sheets(1)
      .Cells(1, 2) = "REPORTE GENERAL DE CARTAS FIANZA"
      .Range(.Cells(1, 2), .Cells(1, 21)).Merge
      
      If Chk_FecAct.Value = 1 Then
         .Cells(3, 2) = " AL " & Format(date, "DD/MM/YYYY")
      Else
         .Cells(3, 2) = " AL " & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), ipp_PerAno.Text), "00") & " DE " & cmb_PerMes.Text & " DE " & Format(ipp_PerAno.Text, "0000")
      End If
      .Range(.Cells(3, 2), .Cells(3, 21)).Merge
      
      .Range(.Cells(1, 2), .Cells(1, 21)).Font.Bold = True
      .Range(.Cells(1, 2), .Cells(3, 21)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 2), .Cells(1, 21)).Font.Size = 14
      
      .Cells(p_int_NroFil, 2) = "TIPO - NRO. DOCUMENTO"
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 2)).Merge
      .Cells(p_int_NroFil, 3) = "RAZON SOCIAL"
      .Range(.Cells(p_int_NroFil, 3), .Cells(p_int_NroFil + 1, 3)).Merge
      .Cells(p_int_NroFil, 4) = "TIPO EMPRESA"
      .Range(.Cells(p_int_NroFil, 4), .Cells(p_int_NroFil + 1, 4)).Merge
      .Cells(p_int_NroFil, 5) = "MODALIDAD"
      .Range(.Cells(p_int_NroFil, 5), .Cells(p_int_NroFil + 1, 5)).Merge
      
      .Cells(p_int_NroFil, 6) = "NÚMERO"
      .Range(.Cells(p_int_NroFil, 6), .Cells(p_int_NroFil + 1, 6)).Merge
      .Cells(p_int_NroFil, 7) = "FECHA EMISIÓN"
      .Range(.Cells(p_int_NroFil, 7), .Cells(p_int_NroFil + 1, 7)).Merge
      .Cells(p_int_NroFil, 8) = "FECHA     VCTO."
      .Range(.Cells(p_int_NroFil, 8), .Cells(p_int_NroFil + 1, 8)).Merge
          
      .Cells(p_int_NroFil, 9) = "MONEDA"
      .Range(.Cells(p_int_NroFil, 9), .Cells(p_int_NroFil + 1, 9)).Merge
      .Cells(p_int_NroFil, 10) = "TASA ANUAL"
      .Range(.Cells(p_int_NroFil, 10), .Cells(p_int_NroFil + 1, 10)).Merge
      .Cells(p_int_NroFil, 11) = "PLAZO"
      .Range(.Cells(p_int_NroFil, 11), .Cells(p_int_NroFil + 1, 11)).Merge
      
      .Cells(p_int_NroFil, 12) = "VALOR" 'CARTA FIANZA
      .Range(.Cells(p_int_NroFil, 12), .Cells(p_int_NroFil + 1, 12)).Merge
      .Cells(p_int_NroFil, 13) = "GARANTIZADO"
      .Range(.Cells(p_int_NroFil, 13), .Cells(p_int_NroFil + 1, 13)).Merge
           
      .Cells(p_int_NroFil, 14) = "GARANTIA"
      .Range(.Cells(p_int_NroFil, 14), .Cells(p_int_NroFil, 16)).Merge
      .Cells(p_int_NroFil + 1, 14) = "IMPORTE"
      .Cells(p_int_NroFil + 1, 15) = "PAGADO"
      .Cells(p_int_NroFil + 1, 16) = "SALDO"
      
      .Cells(p_int_NroFil, 17) = "GARANTIA HIPOTECARIA"
      .Range(.Cells(p_int_NroFil, 17), .Cells(p_int_NroFil + 1, 17)).Merge
      
      .Cells(p_int_NroFil, 18) = "COMISIONES"
      .Range(.Cells(p_int_NroFil, 18), .Cells(p_int_NroFil, 21)).Merge
      .Cells(p_int_NroFil + 1, 18) = "IMPORTE"
      .Cells(p_int_NroFil + 1, 19) = "PAGADO"
      .Cells(p_int_NroFil + 1, 20) = "SALDO"
      .Cells(p_int_NroFil + 1, 21) = "FECHA PAGO"
         
      .Cells(p_int_NroFil, 22) = "FONDOS RECIBIDOS" ' - FMV
      .Range(.Cells(p_int_NroFil, 22), .Cells(p_int_NroFil, 24)).Merge
      .Cells(p_int_NroFil + 1, 22) = "IMPORTE"
      .Cells(p_int_NroFil + 1, 23) = "RECIBIDO"
      .Cells(p_int_NroFil + 1, 24) = "SALDO"
      .Cells(p_int_NroFil, 25) = "DESEMBOLSOS - ET"
      .Range(.Cells(p_int_NroFil, 25), .Cells(p_int_NroFil, 27)).Merge
      .Cells(p_int_NroFil + 1, 25) = "IMPORTE"
      .Cells(p_int_NroFil + 1, 26) = "PAGADO"
      .Cells(p_int_NroFil + 1, 27) = "SALDO"
      .Cells(p_int_NroFil, 28) = "ESTADO"
      .Range(.Cells(p_int_NroFil, 28), .Cells(p_int_NroFil + 1, 28)).Merge
      .Cells(p_int_NroFil, 29) = "Factor"
      .Range(.Cells(p_int_NroFil, 29), .Cells(p_int_NroFil + 1, 29)).Merge
      .Cells(p_int_NroFil, 30) = "Exposición"
      .Range(.Cells(p_int_NroFil, 30), .Cells(p_int_NroFil + 1, 30)).Merge
      .Cells(p_int_NroFil, 31) = "Tasa"
      .Range(.Cells(p_int_NroFil, 31), .Cells(p_int_NroFil + 1, 31)).Merge
      .Cells(p_int_NroFil, 32) = "Prov. Actual"
      .Range(.Cells(p_int_NroFil, 32), .Cells(p_int_NroFil + 1, 32)).Merge
      .Cells(p_int_NroFil, 33) = "Prov. Anterior"
      .Range(.Cells(p_int_NroFil, 33), .Cells(p_int_NroFil + 1, 33)).Merge
      .Cells(p_int_NroFil, 34) = "Diferencia"
      .Range(.Cells(p_int_NroFil, 34), .Cells(p_int_NroFil + 1, 34)).Merge
      .Cells(p_int_NroFil, 35) = "USUARIO"
      .Range(.Cells(p_int_NroFil, 35), .Cells(p_int_NroFil + 1, 35)).Merge
            
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 35)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 35)).Font.Bold = True
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 35)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 16.5
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 60
      .Columns("D").ColumnWidth = 17
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 17.5
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      
      .Columns("F").ColumnWidth = 13
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      
      .Columns("G").ColumnWidth = 13
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      
      .Columns("H").ColumnWidth = 13
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      
      .Columns("I").ColumnWidth = 21
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      
      .Columns("J").ColumnWidth = 13
      .Columns("J").NumberFormat = "0.00%"
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      
      .Columns("K").ColumnWidth = 13
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      
      .Columns("L").ColumnWidth = 13
      .Columns("L").NumberFormat = "###,###,###,##0.00"
      .Columns("L").HorizontalAlignment = xlHAlignRight
      
      .Columns("M").ColumnWidth = 13.5
      .Columns("M").NumberFormat = "###,###,###,##0.00"
      .Columns("M").HorizontalAlignment = xlHAlignRight
      
      .Columns("N").ColumnWidth = 13.5
      .Columns("N").NumberFormat = "###,###,###,##0.00"
      .Columns("N").HorizontalAlignment = xlHAlignRight
      
      .Columns("O").ColumnWidth = 13.5
      .Columns("O").NumberFormat = "###,###,###,##0.00"
      .Columns("O").HorizontalAlignment = xlHAlignRight
      
      .Columns("P").ColumnWidth = 13.5
      .Columns("P").NumberFormat = "###,###,###,##0.00"
      .Columns("P").HorizontalAlignment = xlHAlignRight
      
      .Columns("Q").ColumnWidth = 13.5
      .Columns("Q").HorizontalAlignment = xlHAlignCenter
      
      .Columns("R").ColumnWidth = 13.5
      .Columns("R").NumberFormat = "###,###,###,##0.00"
      .Columns("R").HorizontalAlignment = xlHAlignRight
      
      .Columns("S").ColumnWidth = 13.5
      .Columns("S").NumberFormat = "###,###,###,##0.00"
      .Columns("S").HorizontalAlignment = xlHAlignRight
      
      .Columns("T").ColumnWidth = 13.5
      .Columns("T").NumberFormat = "###,###,###,##0.00"
      .Columns("T").HorizontalAlignment = xlHAlignRight
      
      .Columns("U").ColumnWidth = 13
      .Columns("U").HorizontalAlignment = xlHAlignCenter
      
      .Columns("V").ColumnWidth = 13.5
      .Columns("V").NumberFormat = "###,###,###,##0.00"
      .Columns("V").HorizontalAlignment = xlHAlignRight
      
      .Columns("W").ColumnWidth = 13.5
      .Columns("W").NumberFormat = "###,###,###,##0.00"
      .Columns("W").HorizontalAlignment = xlHAlignRight
      
      .Columns("X").ColumnWidth = 13.5
      .Columns("X").NumberFormat = "###,###,###,##0.00"
      .Columns("X").HorizontalAlignment = xlHAlignRight
      
      .Columns("Y").ColumnWidth = 13.5
      .Columns("Y").NumberFormat = "###,###,###,##0.00"
      .Columns("Y").HorizontalAlignment = xlHAlignRight
      
      .Columns("Z").ColumnWidth = 13.5
      .Columns("Z").NumberFormat = "###,###,###,##0.00"
      .Columns("Z").HorizontalAlignment = xlHAlignRight
      
      .Columns("AA").ColumnWidth = 13.5
      .Columns("AA").NumberFormat = "###,###,###,##0.00"
      .Columns("AA").HorizontalAlignment = xlHAlignRight
      
      .Columns("AB").ColumnWidth = 20 '13.5
      .Columns("AB").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AC").ColumnWidth = 6
      .Columns("AC").NumberFormat = "0.00%"
      .Columns("AC").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AD").ColumnWidth = 13.5
      .Columns("AD").NumberFormat = "###,###,###,##0.00"
      .Columns("AD").HorizontalAlignment = xlHAlignRight
      
      .Columns("AE").ColumnWidth = 6
      .Columns("AE").NumberFormat = "0.00%"
      .Columns("AE").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AF").ColumnWidth = 13.5
      .Columns("AF").NumberFormat = "###,###,###,##0.00"
      .Columns("AF").HorizontalAlignment = xlHAlignRight
      
      .Columns("AG").ColumnWidth = 13.5
      .Columns("AG").NumberFormat = "###,###,###,##0.00"
      .Columns("AG").HorizontalAlignment = xlHAlignRight
      
      .Columns("AH").ColumnWidth = 13.5
      .Columns("AH").NumberFormat = "###,###,###,##0.00"
      .Columns("AH").HorizontalAlignment = xlHAlignRight
      
      .Columns("AI").ColumnWidth = 13.5
      .Columns("AI").HorizontalAlignment = xlHAlignCenter
      
      
      With .Range(.Cells(5, 2), .Cells(6, 35))
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
          .WrapText = True
      End With
      
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(2, 1), .Cells(99, 99)).Font.Size = 11
            
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "   SELECT RPT_PERMES   MES               , RPT_PERANO   ANNO               , RPT_DESCRI   DOCUMENTO             , RPT_VALCAD01 RAZON_SOCIAL   , "
      g_str_Parame = g_str_Parame & "          RPT_VALCAD02 TIPO_EMPRESA      , RPT_VALCAD03 MONEDA             , RPT_VALNUM25 TASA_ANUAL            , RPT_VALNUM01 PLAZO          , "
      g_str_Parame = g_str_Parame & "          RPT_VALCAD05 NUMREF            , RPT_VALCAD06 MAECFI_EMIFIA      , RPT_VALCAD07 MAECFI_VTOFIA         , RPT_VALNUM02 MAECFI_IMPFIA  , "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM03 MAECFI_GARFIA     , RPT_VALNUM04 IMPORTE_GARANTIA   , RPT_VALNUM05 PAGADO_GARANTIA       , RPT_VALNUM06 SALDO_GARANTIA , "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM07 IMPORTE_COMISION  , RPT_VALNUM08 PAGADO_COMISION    , RPT_VALNUM09 SALDO_COMISION        , RPT_VALNUM10 IMPORTE_FONDOS , "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM11 RECIBIDO_FONDOS   , RPT_VALNUM12 SALDO_FONDOS       , RPT_VALNUM13 IMPORTE_DESEMBOLSADO  ,   "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM14 PAGADO_DESEMBOLSO , RPT_VALNUM15 DEVOLUCION_GARANTIA, RPT_VALNUM16 SALDO_DESEMBOLSO      , RPT_VALCAD08 SITUACION    , "
      g_str_Parame = g_str_Parame & "          RPT_VALCAD09 FACTOR            , RPT_VALNUM17 EXPOSICION         , RPT_VALCAD10 TASA                  , RPT_VALNUM18 PROV_ACTUAL  , "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM19 PROV_ANTERIOR     , RPT_VALCAD11 PRODUCTO           , RPT_VALNUM23 MONTO_GARANTIA_LIQUIDA, RPT_VALNUM24 MONTO_GARANTIA_HIPOTECARIA , "
      g_str_Parame = g_str_Parame & "          RPT_VALCAD14 USUARIO           , RPT_VALCAD17 FECHA_PAG_COMISION "
      g_str_Parame = g_str_Parame & "     FROM RPT_TABLA_TEMP "
      g_str_Parame = g_str_Parame & "    WHERE RPT_PERMES = " & Month(Format(gf_FormatoFecha(CStr(l_str_FecPer)), "dd/mm/yyyy")) & "" 'Month(Now)
      g_str_Parame = g_str_Parame & "      AND RPT_PERANO = " & Year(Format(gf_FormatoFecha(CStr(l_str_FecPer)), "dd/mm/yyyy")) & " " 'Year(Now)
      g_str_Parame = g_str_Parame & "      AND RPT_TERCRE = " & "'" & modgen_g_str_NombPC & "'"
      g_str_Parame = g_str_Parame & "      AND RPT_USUCRE = " & "'" & modgen_g_str_CodUsu & "'"
      g_str_Parame = g_str_Parame & "      AND RPT_NOMBRE = " & "'" & Me.cmb_TipRep.Text & "2'"
      g_str_Parame = g_str_Parame & "      AND SUBSTR(RPT_VALCAD05,1,1) <> 2 AND SUBSTR(RPT_VALCAD05,1,1) <> 3 AND SUBSTR(RPT_VALCAD05,1,1) <> 0 "
      g_str_Parame = g_str_Parame & "    ORDER BY RPT_VALCAD02, RPT_VALCAD06 "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
          
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         'g_rst_Princi.Close
         'Set g_rst_Princi = Nothing
         r_bol_Estado = True
'         Call fs_GenExc_TechoPropio_CSO(p_obj_Excel, 5)
'         MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
'         Exit Sub
      End If
      
      p_int_NroFil = p_int_NroFil + 2
      
      If Not g_rst_Princi.EOF Then
         g_rst_Princi.MoveFirst
         
         Do While Not g_rst_Princi.EOF
               If .Cells(p_int_NroFil - 1, 4) <> g_rst_Princi!TIPO_EMPRESA And .Cells(p_int_NroFil - 1, 4) <> "" Then
                  .Range(.Cells(p_int_NroFil, 30), .Cells(p_int_NroFil, 34)).FormulaR1C1 = "=SUM(R[-" & p_int_NroFil - r_int_ConAux - 6 & "]C:R[-1]C)"
                  .Range(.Cells(p_int_NroFil, 30), .Cells(p_int_NroFil, 34)).Font.Bold = True
                  .Range(.Cells(p_int_NroFil, 31), .Cells(p_int_NroFil, 31)).Value = ""
                  r_int_ConAux = p_int_NroFil - 3
                  p_int_NroFil = p_int_NroFil + 3
                  
               End If
               
               .Cells(p_int_NroFil, 2) = g_rst_Princi!DOCUMENTO
               .Cells(p_int_NroFil, 3) = g_rst_Princi!RAZON_SOCIAL
               .Cells(p_int_NroFil, 4) = g_rst_Princi!TIPO_EMPRESA
               .Cells(p_int_NroFil, 5) = IIf(Mid(g_rst_Princi!NUMREF, 1, 1) = 2, "AD", "CF")
               .Cells(p_int_NroFil, 6) = "'" & gf_Formato_NumRef(Trim(g_rst_Princi!NUMREF), Mid(g_rst_Princi!NUMREF, 1, 1))
               .Cells(p_int_NroFil, 7) = "'" & Format(gf_FormatoFecha(CStr(g_rst_Princi!MAECFI_EMIFIA)), "dd/mm/yyyy")
               .Cells(p_int_NroFil, 8) = "'" & Format(gf_FormatoFecha(CStr(g_rst_Princi!MAECFI_VTOFIA)), "dd/mm/yyyy")
               .Cells(p_int_NroFil, 9) = g_rst_Princi!Moneda
               .Cells(p_int_NroFil, 10) = g_rst_Princi!TASA_ANUAL / 100
               .Cells(p_int_NroFil, 11) = g_rst_Princi!PLAZO
               .Cells(p_int_NroFil, 12) = g_rst_Princi!MAECFI_IMPFIA
               .Cells(p_int_NroFil, 13) = g_rst_Princi!MAECFI_GARFIA
               .Cells(p_int_NroFil, 14) = g_rst_Princi!IMPORTE_GARANTIA
               .Cells(p_int_NroFil, 15) = g_rst_Princi!PAGADO_GARANTIA
               .Cells(p_int_NroFil, 16) = g_rst_Princi!SALDO_GARANTIA
               .Cells(p_int_NroFil, 17) = IIf(g_rst_Princi!MONTO_GARANTIA_HIPOTECARIA > 0, "SI", "")
               .Cells(p_int_NroFil, 18) = g_rst_Princi!IMPORTE_COMISION
               .Cells(p_int_NroFil, 19) = g_rst_Princi!PAGADO_COMISION
               .Cells(p_int_NroFil, 20) = g_rst_Princi!SALDO_COMISION
               If Not IsNull(g_rst_Princi!FECHA_PAG_COMISION) Then
                  .Cells(p_int_NroFil, 21) = "'" & Format(gf_FormatoFecha(CStr(g_rst_Princi!FECHA_PAG_COMISION)), "dd/mm/yyyy")
               Else
                  .Cells(p_int_NroFil, 21) = ""
               End If
               .Cells(p_int_NroFil, 22) = g_rst_Princi!IMPORTE_FONDOS
               .Cells(p_int_NroFil, 23) = g_rst_Princi!RECIBIDO_FONDOS
               .Cells(p_int_NroFil, 24) = g_rst_Princi!SALDO_FONDOS
               .Cells(p_int_NroFil, 25) = g_rst_Princi!IMPORTE_DESEMBOLSADO
               .Cells(p_int_NroFil, 26) = g_rst_Princi!PAGADO_DESEMBOLSO
               .Cells(p_int_NroFil, 27) = g_rst_Princi!SALDO_DESEMBOLSO
               .Cells(p_int_NroFil, 28) = g_rst_Princi!SITUACION
               .Cells(p_int_NroFil, 29) = g_rst_Princi!FACTOR
               .Cells(p_int_NroFil, 30) = g_rst_Princi!EXPOSICION
               .Cells(p_int_NroFil, 31) = g_rst_Princi!TASA
               .Cells(p_int_NroFil, 32) = g_rst_Princi!PROV_ACTUAL
               .Cells(p_int_NroFil, 33) = g_rst_Princi!PROV_ANTERIOR
               .Cells(p_int_NroFil, 34) = .Cells(p_int_NroFil, 32) - .Cells(p_int_NroFil, 33)
               .Cells(p_int_NroFil, 35) = g_rst_Princi!USUARIO
               
            p_int_NroFil = p_int_NroFil + 1
            g_rst_Princi.MoveNext
         Loop
      End If
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 35)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 35)).Font.Bold = True
      .Range(.Cells(p_int_NroFil, 5), .Cells(p_int_NroFil, 35)).HorizontalAlignment = xlHAlignRight
            
      'SUMATORIA TOTAL
      .Cells(p_int_NroFil, 3) = "TOTAL"
      .Cells(p_int_NroFil, 13).FormulaR1C1 = "=SUM(R[-" & p_int_NroFil - 5 & "]C:R[-1]C)"
      
      .Range(.Cells(p_int_NroFil, 30), .Cells(p_int_NroFil, 34)).FormulaR1C1 = "=SUM(R[-" & p_int_NroFil - r_int_ConAux - 4 & "]C:R[-1]C)" '-3
      .Range(.Cells(p_int_NroFil, 30), .Cells(p_int_NroFil, 34)).Font.Bold = True
      .Range(.Cells(p_int_NroFil, 31), .Cells(p_int_NroFil, 31)).Value = ""
      
      'RESUMEN POR TIPO EMPRESA
      .Range(.Cells(p_int_NroFil + 2, 2), .Cells(p_int_NroFil + 2, 5)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil + 2, 2), .Cells(p_int_NroFil + 2, 5)).Font.Bold = True
      .Cells(p_int_NroFil + 2, 2) = "CANTIDAD"
      .Cells(p_int_NroFil + 2, 3) = "DESCRIPCION"
      .Cells(p_int_NroFil + 2, 4) = "GARANTIZADO"
      .Cells(p_int_NroFil + 2, 5) = "PROVISION ACTUAL"
      
      .Range(.Cells(p_int_NroFil + 3, 3), .Cells(p_int_NroFil + 6, 3)).Font.Bold = True
      .Range(.Cells(p_int_NroFil + 3, 4), .Cells(p_int_NroFil + 6, 5)).NumberFormat = "###,###,###,##0.00"
      .Range(.Cells(p_int_NroFil + 3, 4), .Cells(p_int_NroFil + 6, 5)).HorizontalAlignment = xlHAlignRight
      
      .Cells(p_int_NroFil + 3, 2).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 4) & "]C[2]:R[-4]C[2],""PEQUEÑA"")"
      .Cells(p_int_NroFil + 3, 3) = "RESUMEN PEQUEÑA EMPRESA"
      .Cells(p_int_NroFil + 3, 4).FormulaR1C1 = "=SUMIF(R[-" & (p_int_NroFil - 4) & "]C:R[-4]C[9],""PEQUEÑA"",R[-" & (p_int_NroFil - 4) & "]C[9]:R[-4]C[9])"
      .Cells(p_int_NroFil + 3, 5).FormulaR1C1 = "=SUMIF(R[-" & (p_int_NroFil - 4) & "]C[-1]:R[-4]C[27],""PEQUEÑA"",R[-" & (p_int_NroFil - 4) & "]C[27]:R[-4]C[27])"
      
      .Cells(p_int_NroFil + 4, 2).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 3) & "]C[2]:R[-5]C[2],""MEDIANA"")"
      .Cells(p_int_NroFil + 4, 3) = "RESUMEN MEDIANA EMPRESA"
      .Cells(p_int_NroFil + 4, 4).FormulaR1C1 = "=SUMIF(R[-" & (p_int_NroFil - 3) & "]C:R[-5]C[9],""MEDIANA"",R[-" & (p_int_NroFil - 3) & "]C[9]:R[-5]C[9])"
      .Cells(p_int_NroFil + 4, 5).FormulaR1C1 = "=SUMIF(R[-" & (p_int_NroFil - 3) & "]C[-1]:R[-5]C[27],""MEDIANA"",R[-" & (p_int_NroFil - 3) & "]C[27]:R[-5]C[27])"
      
      .Cells(p_int_NroFil + 5, 2).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 2) & "]C[2]:R[-6]C[2],""GRANDE"")"
      .Cells(p_int_NroFil + 5, 3) = "RESUMEN GRANDE EMPRESA"
      .Cells(p_int_NroFil + 5, 4).FormulaR1C1 = "=+SUMIF(R[-" & (p_int_NroFil - 2) & "]C:R[-6]C[9],""GRANDE"",R[-" & (p_int_NroFil - 2) & "]C[9]:R[-6]C[9])"
      .Cells(p_int_NroFil + 5, 5).FormulaR1C1 = "=+SUMIF(R[-" & (p_int_NroFil - 2) & "]C[-1]:R[-6]C[27],""GRANDE"",R[-" & (p_int_NroFil - 2) & "]C[27]:R[-6]C[27])"
      
      .Cells(p_int_NroFil + 6, 2).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 1) & "]C[2]:R[-7]C[2],""MICRO"")"
      .Cells(p_int_NroFil + 6, 3) = "RESUMEN MICRO EMPRESA"
      .Cells(p_int_NroFil + 6, 4).FormulaR1C1 = "=+SUMIF(R[-" & (p_int_NroFil - 1) & "]C:R[-7]C[9],""MICRO"",R[-" & (p_int_NroFil - 1) & "]C[9]:R[-7]C[9])"
      .Cells(p_int_NroFil + 6, 5).FormulaR1C1 = "=+SUMIF(R[-" & (p_int_NroFil - 1) & "]C[-1]:R[-7]C[27],""MICRO"",R[-" & (p_int_NroFil - 1) & "]C[27]:R[-7]C[27])"
      
      .Cells(p_int_NroFil + 7, 2).FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
      .Cells(p_int_NroFil + 7, 3) = "TOTAL"
      .Range(.Cells(p_int_NroFil + 7, 4), .Cells(p_int_NroFil + 7, 5)).FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
      
      .Range(.Cells(p_int_NroFil + 7, 4), .Cells(p_int_NroFil + 7, 5)).HorizontalAlignment = xlHAlignRight
      .Range(.Cells(p_int_NroFil + 7, 2), .Cells(p_int_NroFil + 7, 5)).Font.Bold = True
      
      With .Range(.Cells(1, 2), .Cells(1, 35))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
      End With
      
      With .Range(.Cells(5, 2), .Cells(p_int_NroFil - 1, 35))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:            .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:     .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
      With .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 35))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:            .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:     .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
             
      With .Range(.Cells(5, 14), .Cells(p_int_NroFil, 16))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:            .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:     .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
      With .Range(.Cells(5, 18), .Cells(p_int_NroFil, 21))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:            .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:     .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
      With .Range(.Cells(5, 22), .Cells(p_int_NroFil, 24))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:            .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:     .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
      With .Range(.Cells(5, 25), .Cells(p_int_NroFil, 27))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:            .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:     .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      'RESUMEN
      With .Range(.Cells(p_int_NroFil + 2, 2), .Cells(p_int_NroFil + 6, 5))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeLeft).Weight = xlThin
          .Borders(xlEdgeTop).LineStyle = xlContinuous:            .Borders(xlEdgeTop).Weight = xlThin
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlThin
          .Borders(xlEdgeRight).LineStyle = xlContinuous:          .Borders(xlEdgeRight).Weight = xlThin
      End With
      
      With .Range(.Cells(p_int_NroFil + 7, 2), .Cells(p_int_NroFil + 7, 5))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeTop).LineStyle = xlContinuous:
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeRight).LineStyle = xlContinuous:
      End With
      
      With .Range(.Cells(5, 2), .Cells(4, 35))
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlMedium
      End With
   End With
     
   p_obj_Excel.Sheets(1).Range("E7").Select
   p_obj_Excel.ActiveWindow.FreezePanes = True
   
End Sub
   '***************************************************************** REPORTE GENERAL - ADENDAS *****************************************************************
Private Sub fs_GenExc_TechoPropio_AD(ByVal p_obj_Excel As Excel.Application, ByVal p_int_NroFil As Integer)
Dim r_int_ConAux        As Integer
Dim r_bol_Estado        As Boolean

'   p_int_NroFil = 5
   p_obj_Excel.Sheets(2).Select

   With p_obj_Excel.Sheets(2)
      .Cells(1, 2) = "REPORTE GENERAL DE ADENDAS"
      .Range(.Cells(1, 2), .Cells(1, 21)).Merge

      If Chk_FecAct.Value = 1 Then
         .Cells(3, 2) = " AL " & Format(date, "DD/MM/YYYY")
      Else
         .Cells(3, 2) = " AL " & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), ipp_PerAno.Text), "00") & " DE " & cmb_PerMes.Text & " DE " & Format(ipp_PerAno.Text, "0000")
      End If
      .Range(.Cells(3, 2), .Cells(3, 21)).Merge

      .Range(.Cells(1, 2), .Cells(1, 21)).Font.Bold = True
      .Range(.Cells(1, 2), .Cells(3, 21)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 2), .Cells(1, 21)).Font.Size = 14

      .Cells(p_int_NroFil, 2) = "TIPO - NRO. DOCUMENTO"
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 2)).Merge
      .Cells(p_int_NroFil, 3) = "RAZON SOCIAL"
      .Range(.Cells(p_int_NroFil, 3), .Cells(p_int_NroFil + 1, 3)).Merge
      .Cells(p_int_NroFil, 4) = "TIPO EMPRESA"
      .Range(.Cells(p_int_NroFil, 4), .Cells(p_int_NroFil + 1, 4)).Merge
      .Cells(p_int_NroFil, 5) = "MODALIDAD"
      .Range(.Cells(p_int_NroFil, 5), .Cells(p_int_NroFil + 1, 5)).Merge

      .Cells(p_int_NroFil, 6) = "NÚMERO"
      .Range(.Cells(p_int_NroFil, 6), .Cells(p_int_NroFil + 1, 6)).Merge
      .Cells(p_int_NroFil, 7) = "FECHA EMISIÓN"
      .Range(.Cells(p_int_NroFil, 7), .Cells(p_int_NroFil + 1, 7)).Merge
      .Cells(p_int_NroFil, 8) = "FECHA     VCTO."
      .Range(.Cells(p_int_NroFil, 8), .Cells(p_int_NroFil + 1, 8)).Merge

      .Cells(p_int_NroFil, 9) = "MONEDA"
      .Range(.Cells(p_int_NroFil, 9), .Cells(p_int_NroFil + 1, 9)).Merge
      .Cells(p_int_NroFil, 10) = "TASA ANUAL"
      .Range(.Cells(p_int_NroFil, 10), .Cells(p_int_NroFil + 1, 10)).Merge
      .Cells(p_int_NroFil, 11) = "PLAZO"
      .Range(.Cells(p_int_NroFil, 11), .Cells(p_int_NroFil + 1, 11)).Merge

      .Cells(p_int_NroFil, 12) = "VALOR" 'CARTA FIANZA
      .Range(.Cells(p_int_NroFil, 12), .Cells(p_int_NroFil + 1, 12)).Merge
      .Cells(p_int_NroFil, 13) = "GARANTIZADO"
      .Range(.Cells(p_int_NroFil, 13), .Cells(p_int_NroFil + 1, 13)).Merge
           
      .Cells(p_int_NroFil, 14) = "GARANTIA"
      .Range(.Cells(p_int_NroFil, 14), .Cells(p_int_NroFil, 16)).Merge
      .Cells(p_int_NroFil + 1, 14) = "IMPORTE"
      .Cells(p_int_NroFil + 1, 15) = "PAGADO"
      .Cells(p_int_NroFil + 1, 16) = "SALDO"
      
      .Cells(p_int_NroFil, 17) = "GARANTIA HIPOTECARIA"
      .Range(.Cells(p_int_NroFil, 17), .Cells(p_int_NroFil + 1, 17)).Merge
      
      .Cells(p_int_NroFil, 18) = "COMISIONES"
      .Range(.Cells(p_int_NroFil, 18), .Cells(p_int_NroFil, 21)).Merge
      .Cells(p_int_NroFil + 1, 18) = "IMPORTE"
      .Cells(p_int_NroFil + 1, 19) = "PAGADO"
      .Cells(p_int_NroFil + 1, 20) = "SALDO"
      .Cells(p_int_NroFil + 1, 21) = "FECHA PAGO"
      .Cells(p_int_NroFil, 22) = "FONDOS RECIBIDOS" ' - FMV
      .Range(.Cells(p_int_NroFil, 22), .Cells(p_int_NroFil, 24)).Merge
      .Cells(p_int_NroFil + 1, 22) = "IMPORTE"
      .Cells(p_int_NroFil + 1, 23) = "RECIBIDO"
      .Cells(p_int_NroFil + 1, 24) = "SALDO"
      .Cells(p_int_NroFil, 25) = "DESEMBOLSOS - ET"
      .Range(.Cells(p_int_NroFil, 25), .Cells(p_int_NroFil, 27)).Merge
      .Cells(p_int_NroFil + 1, 25) = "IMPORTE"
      .Cells(p_int_NroFil + 1, 26) = "PAGADO"
      .Cells(p_int_NroFil + 1, 27) = "SALDO"
      .Cells(p_int_NroFil, 28) = "ESTADO"
      .Range(.Cells(p_int_NroFil, 28), .Cells(p_int_NroFil + 1, 28)).Merge
      .Cells(p_int_NroFil, 29) = "Factor"
      .Range(.Cells(p_int_NroFil, 29), .Cells(p_int_NroFil + 1, 29)).Merge
      .Cells(p_int_NroFil, 30) = "Exposición"
      .Range(.Cells(p_int_NroFil, 30), .Cells(p_int_NroFil + 1, 30)).Merge
      .Cells(p_int_NroFil, 31) = "Tasa"
      .Range(.Cells(p_int_NroFil, 31), .Cells(p_int_NroFil + 1, 31)).Merge
      .Cells(p_int_NroFil, 32) = "Prov. Actual"
      .Range(.Cells(p_int_NroFil, 32), .Cells(p_int_NroFil + 1, 32)).Merge
      .Cells(p_int_NroFil, 33) = "Prov. Anterior"
      .Range(.Cells(p_int_NroFil, 33), .Cells(p_int_NroFil + 1, 33)).Merge
      .Cells(p_int_NroFil, 34) = "Diferencia"
      .Range(.Cells(p_int_NroFil, 34), .Cells(p_int_NroFil + 1, 34)).Merge
      .Cells(p_int_NroFil, 35) = "USUARIO"
      .Range(.Cells(p_int_NroFil, 35), .Cells(p_int_NroFil + 1, 35)).Merge
      
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 35)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 35)).Font.Bold = True
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 35)).HorizontalAlignment = xlHAlignCenter

      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 16.5
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 60
      .Columns("D").ColumnWidth = 17
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 17.5
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      
      .Columns("F").ColumnWidth = 13
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      
      .Columns("G").ColumnWidth = 13
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      
      .Columns("H").ColumnWidth = 13
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      
      .Columns("I").ColumnWidth = 21
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      
      .Columns("J").ColumnWidth = 13
      .Columns("J").NumberFormat = "0.00%"
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      
      .Columns("K").ColumnWidth = 13
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      
      .Columns("L").ColumnWidth = 13
      .Columns("L").NumberFormat = "###,###,###,##0.00"
      .Columns("L").HorizontalAlignment = xlHAlignRight
      
      .Columns("M").ColumnWidth = 13.5
      .Columns("M").NumberFormat = "###,###,###,##0.00"
      .Columns("M").HorizontalAlignment = xlHAlignRight
      
      .Columns("N").ColumnWidth = 13.5
      .Columns("N").NumberFormat = "###,###,###,##0.00"
      .Columns("N").HorizontalAlignment = xlHAlignRight
      
      .Columns("O").ColumnWidth = 13.5
      .Columns("O").NumberFormat = "###,###,###,##0.00"
      .Columns("O").HorizontalAlignment = xlHAlignRight
      
      .Columns("P").ColumnWidth = 13.5
      .Columns("P").NumberFormat = "###,###,###,##0.00"
      .Columns("P").HorizontalAlignment = xlHAlignRight
      
      .Columns("Q").ColumnWidth = 13.5
      .Columns("Q").HorizontalAlignment = xlHAlignCenter
      
      .Columns("R").ColumnWidth = 13.5
      .Columns("R").NumberFormat = "###,###,###,##0.00"
      .Columns("R").HorizontalAlignment = xlHAlignRight
      
      .Columns("S").ColumnWidth = 13.5
      .Columns("S").NumberFormat = "###,###,###,##0.00"
      .Columns("S").HorizontalAlignment = xlHAlignRight
      
      .Columns("T").ColumnWidth = 13.5
      .Columns("T").NumberFormat = "###,###,###,##0.00"
      .Columns("T").HorizontalAlignment = xlHAlignRight
      
      .Columns("U").ColumnWidth = 13
      .Columns("U").HorizontalAlignment = xlHAlignCenter
      
      .Columns("V").ColumnWidth = 13.5
      .Columns("V").NumberFormat = "###,###,###,##0.00"
      .Columns("V").HorizontalAlignment = xlHAlignRight
      
      .Columns("W").ColumnWidth = 13.5
      .Columns("W").NumberFormat = "###,###,###,##0.00"
      .Columns("W").HorizontalAlignment = xlHAlignRight
      
      .Columns("X").ColumnWidth = 13.5
      .Columns("X").NumberFormat = "###,###,###,##0.00"
      .Columns("X").HorizontalAlignment = xlHAlignRight
      
      .Columns("Y").ColumnWidth = 13.5
      .Columns("Y").NumberFormat = "###,###,###,##0.00"
      .Columns("Y").HorizontalAlignment = xlHAlignRight
      
      .Columns("Z").ColumnWidth = 13.5
      .Columns("Z").NumberFormat = "###,###,###,##0.00"
      .Columns("Z").HorizontalAlignment = xlHAlignRight
      
      .Columns("AA").ColumnWidth = 13.5
      .Columns("AA").NumberFormat = "###,###,###,##0.00"
      .Columns("AA").HorizontalAlignment = xlHAlignRight
      
      .Columns("AB").ColumnWidth = 20 '13.5
      .Columns("AB").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AC").ColumnWidth = 6
      .Columns("AC").NumberFormat = "0.00%"
      .Columns("AC").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AD").ColumnWidth = 13.5
      .Columns("AD").NumberFormat = "###,###,###,##0.00"
      .Columns("AD").HorizontalAlignment = xlHAlignRight
      
      .Columns("AE").ColumnWidth = 6
      .Columns("AE").NumberFormat = "0.00%"
      .Columns("AE").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AF").ColumnWidth = 13.5
      .Columns("AF").NumberFormat = "###,###,###,##0.00"
      .Columns("AF").HorizontalAlignment = xlHAlignRight
      
      .Columns("AG").ColumnWidth = 13.5
      .Columns("AG").NumberFormat = "###,###,###,##0.00"
      .Columns("AG").HorizontalAlignment = xlHAlignRight
      
      .Columns("AH").ColumnWidth = 13.5
      .Columns("AH").NumberFormat = "###,###,###,##0.00"
      .Columns("AH").HorizontalAlignment = xlHAlignRight
      
      .Columns("AI").ColumnWidth = 13.5
      .Columns("AI").HorizontalAlignment = xlHAlignCenter

      With .Range(.Cells(5, 2), .Cells(6, 35))
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
          .WrapText = True
      End With

      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(2, 1), .Cells(99, 99)).Font.Size = 11

      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "   SELECT RPT_PERMES   MES               , RPT_PERANO   ANNO               , RPT_DESCRI   DOCUMENTO             , RPT_VALCAD01 RAZON_SOCIAL   , "
      g_str_Parame = g_str_Parame & "          RPT_VALCAD02 TIPO_EMPRESA      , RPT_VALCAD03 MONEDA             , RPT_VALNUM25 TASA_ANUAL            , RPT_VALNUM01 PLAZO          , "
      g_str_Parame = g_str_Parame & "          RPT_VALCAD05 NUMREF            , RPT_VALCAD06 MAECFI_EMIFIA      , RPT_VALCAD07 MAECFI_VTOFIA         , RPT_VALNUM02 MAECFI_IMPFIA  , "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM03 MAECFI_GARFIA     , RPT_VALNUM04 IMPORTE_GARANTIA   , RPT_VALNUM05 PAGADO_GARANTIA       , RPT_VALNUM06 SALDO_GARANTIA , "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM07 IMPORTE_COMISION  , RPT_VALNUM08 PAGADO_COMISION    , RPT_VALNUM09 SALDO_COMISION        , RPT_VALNUM10 IMPORTE_FONDOS , "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM11 RECIBIDO_FONDOS   , RPT_VALNUM12 SALDO_FONDOS       , RPT_VALNUM13 IMPORTE_DESEMBOLSADO  ,   "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM14 PAGADO_DESEMBOLSO , RPT_VALNUM15 DEVOLUCION_GARANTIA, RPT_VALNUM16 SALDO_DESEMBOLSO      , RPT_VALCAD08 SITUACION    , "
      g_str_Parame = g_str_Parame & "          RPT_VALCAD09 FACTOR            , RPT_VALNUM17 EXPOSICION         , RPT_VALCAD10 TASA                  , RPT_VALNUM18 PROV_ACTUAL  , "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM19 PROV_ANTERIOR     , RPT_VALCAD11 PRODUCTO           , RPT_VALNUM23 MONTO_GARANTIA_LIQUIDA, RPT_VALNUM24 MONTO_GARANTIA_HIPOTECARIA, "
      g_str_Parame = g_str_Parame & "          RPT_VALCAD14 USUARIO           , RPT_VALCAD17 FECHA_PAG_COMISION "
      g_str_Parame = g_str_Parame & "     FROM RPT_TABLA_TEMP "
      g_str_Parame = g_str_Parame & "    WHERE RPT_PERMES = " & Month(Format(gf_FormatoFecha(CStr(l_str_FecPer)), "dd/mm/yyyy")) & "" 'Month(Now)
      g_str_Parame = g_str_Parame & "      AND RPT_PERANO = " & Year(Format(gf_FormatoFecha(CStr(l_str_FecPer)), "dd/mm/yyyy")) & " " 'Year(Now)
      g_str_Parame = g_str_Parame & "      AND RPT_TERCRE = " & "'" & modgen_g_str_NombPC & "'"
      g_str_Parame = g_str_Parame & "      AND RPT_USUCRE = " & "'" & modgen_g_str_CodUsu & "'"
      g_str_Parame = g_str_Parame & "      AND RPT_NOMBRE = " & "'" & Me.cmb_TipRep.Text & "2'"
      g_str_Parame = g_str_Parame & "      AND SUBSTR(RPT_VALCAD05,1,1) = 2 "
      g_str_Parame = g_str_Parame & "    ORDER BY RPT_VALCAD02, RPT_VALCAD06 "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If

      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         'g_rst_Princi.Close
         'Set g_rst_Princi = Nothing
         r_bol_Estado = True
'         Call fs_GenExc_TechoPropio_CF(p_obj_Excel, 5) 'GoTo Cartas
         'MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
'         Exit Sub
      End If
      
      p_int_NroFil = p_int_NroFil + 2
      
      If Not g_rst_Princi.EOF Then
         g_rst_Princi.MoveFirst
   
         Do While Not g_rst_Princi.EOF
               If .Cells(p_int_NroFil - 1, 4) <> g_rst_Princi!TIPO_EMPRESA And .Cells(p_int_NroFil - 1, 4) <> "" Then
                  .Range(.Cells(p_int_NroFil, 30), .Cells(p_int_NroFil, 34)).FormulaR1C1 = "=SUM(R[-" & p_int_NroFil - r_int_ConAux - 6 & "]C:R[-1]C)"
                  .Range(.Cells(p_int_NroFil, 30), .Cells(p_int_NroFil, 34)).Font.Bold = True
                  .Range(.Cells(p_int_NroFil, 31), .Cells(p_int_NroFil, 31)).Value = ""
                  r_int_ConAux = p_int_NroFil - 3
                  p_int_NroFil = p_int_NroFil + 3
               End If
   
               .Cells(p_int_NroFil, 2) = g_rst_Princi!DOCUMENTO
               .Cells(p_int_NroFil, 3) = g_rst_Princi!RAZON_SOCIAL
               .Cells(p_int_NroFil, 4) = g_rst_Princi!TIPO_EMPRESA
               .Cells(p_int_NroFil, 5) = IIf(Mid(g_rst_Princi!NUMREF, 1, 1) = 2, "AD", "CF")
               .Cells(p_int_NroFil, 6) = "'" & gf_Formato_NumRef(Trim(g_rst_Princi!NUMREF), Mid(g_rst_Princi!NUMREF, 1, 1))
               .Cells(p_int_NroFil, 7) = "'" & Format(gf_FormatoFecha(CStr(g_rst_Princi!MAECFI_EMIFIA)), "dd/mm/yyyy")
               .Cells(p_int_NroFil, 8) = "'" & Format(gf_FormatoFecha(CStr(g_rst_Princi!MAECFI_VTOFIA)), "dd/mm/yyyy")
               .Cells(p_int_NroFil, 9) = g_rst_Princi!Moneda
               .Cells(p_int_NroFil, 10) = g_rst_Princi!TASA_ANUAL / 100
               .Cells(p_int_NroFil, 11) = g_rst_Princi!PLAZO
               .Cells(p_int_NroFil, 12) = g_rst_Princi!MAECFI_IMPFIA
               .Cells(p_int_NroFil, 13) = g_rst_Princi!MAECFI_GARFIA
               .Cells(p_int_NroFil, 14) = g_rst_Princi!IMPORTE_GARANTIA
               .Cells(p_int_NroFil, 15) = g_rst_Princi!PAGADO_GARANTIA
               .Cells(p_int_NroFil, 16) = g_rst_Princi!SALDO_GARANTIA
               .Cells(p_int_NroFil, 17) = IIf(g_rst_Princi!MONTO_GARANTIA_HIPOTECARIA > 0, "SI", "")
               .Cells(p_int_NroFil, 18) = g_rst_Princi!IMPORTE_COMISION
               .Cells(p_int_NroFil, 19) = g_rst_Princi!PAGADO_COMISION
               .Cells(p_int_NroFil, 20) = g_rst_Princi!SALDO_COMISION
               If Not IsNull(g_rst_Princi!FECHA_PAG_COMISION) Then
                  .Cells(p_int_NroFil, 21) = "'" & Format(gf_FormatoFecha(CStr(g_rst_Princi!FECHA_PAG_COMISION)), "dd/mm/yyyy")
               Else
                  .Cells(p_int_NroFil, 21) = ""
               End If
               .Cells(p_int_NroFil, 22) = g_rst_Princi!IMPORTE_FONDOS
               .Cells(p_int_NroFil, 23) = g_rst_Princi!RECIBIDO_FONDOS
               .Cells(p_int_NroFil, 24) = g_rst_Princi!SALDO_FONDOS
               .Cells(p_int_NroFil, 25) = g_rst_Princi!IMPORTE_DESEMBOLSADO
               .Cells(p_int_NroFil, 26) = g_rst_Princi!PAGADO_DESEMBOLSO
               .Cells(p_int_NroFil, 27) = g_rst_Princi!SALDO_DESEMBOLSO
               .Cells(p_int_NroFil, 28) = g_rst_Princi!SITUACION
               .Cells(p_int_NroFil, 29) = g_rst_Princi!FACTOR
               .Cells(p_int_NroFil, 30) = g_rst_Princi!EXPOSICION
               .Cells(p_int_NroFil, 31) = g_rst_Princi!TASA
               .Cells(p_int_NroFil, 32) = g_rst_Princi!PROV_ACTUAL
               .Cells(p_int_NroFil, 33) = g_rst_Princi!PROV_ANTERIOR
               .Cells(p_int_NroFil, 34) = .Cells(p_int_NroFil, 32) - .Cells(p_int_NroFil, 33)
               .Cells(p_int_NroFil, 35) = g_rst_Princi!USUARIO
               
            p_int_NroFil = p_int_NroFil + 1
            g_rst_Princi.MoveNext
         Loop
      End If
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 35)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 35)).Font.Bold = True
      .Range(.Cells(p_int_NroFil, 5), .Cells(p_int_NroFil, 35)).HorizontalAlignment = xlHAlignRight

      'SUMATORIA TOTAL
      .Cells(p_int_NroFil, 3) = "TOTAL"
      .Cells(p_int_NroFil, 13).FormulaR1C1 = "=SUM(R[-" & p_int_NroFil - 5 & "]C:R[-1]C)"

      .Range(.Cells(p_int_NroFil, 30), .Cells(p_int_NroFil, 34)).FormulaR1C1 = "=SUM(R[-" & p_int_NroFil - r_int_ConAux - 4 & "]C:R[-1]C)" '-3
      .Range(.Cells(p_int_NroFil, 30), .Cells(p_int_NroFil, 34)).Font.Bold = True
      .Range(.Cells(p_int_NroFil, 31), .Cells(p_int_NroFil, 31)).Value = ""

      'RESUMEN POR TIPO EMPRESA
      .Range(.Cells(p_int_NroFil + 2, 2), .Cells(p_int_NroFil + 2, 5)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil + 2, 2), .Cells(p_int_NroFil + 2, 5)).Font.Bold = True
      .Cells(p_int_NroFil + 2, 2) = "CANTIDAD"
      .Cells(p_int_NroFil + 2, 3) = "DESCRIPCION"
      .Cells(p_int_NroFil + 2, 4) = "GARANTIZADO"
      .Cells(p_int_NroFil + 2, 5) = "PROVISION ACTUAL"

      .Range(.Cells(p_int_NroFil + 3, 3), .Cells(p_int_NroFil + 6, 3)).Font.Bold = True
      .Range(.Cells(p_int_NroFil + 3, 4), .Cells(p_int_NroFil + 6, 5)).NumberFormat = "###,###,###,##0.00"
      .Range(.Cells(p_int_NroFil + 3, 4), .Cells(p_int_NroFil + 6, 5)).HorizontalAlignment = xlHAlignRight
      

      .Cells(p_int_NroFil + 3, 2).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 4) & "]C[2]:R[-4]C[2],""PEQUEÑA"")"
      .Cells(p_int_NroFil + 3, 3) = "RESUMEN PEQUEÑA EMPRESA"
      .Cells(p_int_NroFil + 3, 4).FormulaR1C1 = "=SUMIF(R[-" & (p_int_NroFil - 4) & "]C:R[-4]C[9],""PEQUEÑA"",R[-" & (p_int_NroFil - 4) & "]C[9]:R[-4]C[9])"
      .Cells(p_int_NroFil + 3, 5).FormulaR1C1 = "=SUMIF(R[-" & (p_int_NroFil - 4) & "]C[-1]:R[-4]C[27],""PEQUEÑA"",R[-" & (p_int_NroFil - 4) & "]C[27]:R[-4]C[27])"

      .Cells(p_int_NroFil + 4, 2).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 3) & "]C[2]:R[-5]C[2],""MEDIANA"")"
      .Cells(p_int_NroFil + 4, 3) = "RESUMEN MEDIANA EMPRESA"
      .Cells(p_int_NroFil + 4, 4).FormulaR1C1 = "=SUMIF(R[-" & (p_int_NroFil - 3) & "]C:R[-5]C[9],""MEDIANA"",R[-" & (p_int_NroFil - 3) & "]C[9]:R[-5]C[9])"
      .Cells(p_int_NroFil + 4, 5).FormulaR1C1 = "=SUMIF(R[-" & (p_int_NroFil - 3) & "]C[-1]:R[-5]C[27],""MEDIANA"",R[-" & (p_int_NroFil - 3) & "]C[27]:R[-5]C[27])"

      .Cells(p_int_NroFil + 5, 2).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 2) & "]C[2]:R[-6]C[2],""GRANDE"")"
      .Cells(p_int_NroFil + 5, 3) = "RESUMEN GRANDE EMPRESA"
      .Cells(p_int_NroFil + 5, 4).FormulaR1C1 = "=+SUMIF(R[-" & (p_int_NroFil - 2) & "]C:R[-6]C[9],""GRANDE"",R[-" & (p_int_NroFil - 2) & "]C[9]:R[-6]C[9])"
      .Cells(p_int_NroFil + 5, 5).FormulaR1C1 = "=+SUMIF(R[-" & (p_int_NroFil - 2) & "]C[-1]:R[-6]C[27],""GRANDE"",R[-" & (p_int_NroFil - 2) & "]C[27]:R[-6]C[27])"

      .Cells(p_int_NroFil + 6, 2).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 1) & "]C[2]:R[-7]C[2],""MICRO"")"
      .Cells(p_int_NroFil + 6, 3) = "RESUMEN MICRO EMPRESA"
      .Cells(p_int_NroFil + 6, 4).FormulaR1C1 = "=+SUMIF(R[-" & (p_int_NroFil - 1) & "]C:R[-7]C[9],""MICRO"",R[-" & (p_int_NroFil - 1) & "]C[9]:R[-7]C[9])"
      .Cells(p_int_NroFil + 6, 5).FormulaR1C1 = "=+SUMIF(R[-" & (p_int_NroFil - 1) & "]C[-1]:R[-7]C[27],""MICRO"",R[-" & (p_int_NroFil - 1) & "]C[27]:R[-7]C[27])"
      
      .Cells(p_int_NroFil + 7, 2).FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
      .Cells(p_int_NroFil + 7, 3) = "TOTAL"
      .Range(.Cells(p_int_NroFil + 7, 4), .Cells(p_int_NroFil + 7, 5)).FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
      
      .Range(.Cells(p_int_NroFil + 7, 4), .Cells(p_int_NroFil + 7, 5)).HorizontalAlignment = xlHAlignRight
      .Range(.Cells(p_int_NroFil + 7, 2), .Cells(p_int_NroFil + 7, 7)).Font.Bold = True
      
      With .Range(.Cells(1, 2), .Cells(1, 35))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
      End With

      With .Range(.Cells(5, 2), .Cells(p_int_NroFil - 1, 35))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With

      With .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 35))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With

      With .Range(.Cells(5, 14), .Cells(p_int_NroFil, 16))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With

      With .Range(.Cells(5, 18), .Cells(p_int_NroFil, 21))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:    .Borders(xlInsideHorizontal).LineStyle = xlContinuous:
      End With

      With .Range(.Cells(5, 22), .Cells(p_int_NroFil, 24))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With

      With .Range(.Cells(5, 25), .Cells(p_int_NroFil, 27))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
      'RESUMEN
      With .Range(.Cells(p_int_NroFil + 2, 2), .Cells(p_int_NroFil + 6, 5))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlThin
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlThin
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlThin
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlThin
      End With
      
      With .Range(.Cells(p_int_NroFil + 7, 2), .Cells(p_int_NroFil + 7, 5))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:         ' .Borders(xlEdgeLeft).Weight = xlThin
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           '.Borders(xlEdgeTop).Weight = xlThin
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:       ' .Borders(xlEdgeBottom).Weight = xlThin
          .Borders(xlEdgeRight).LineStyle = xlContinuous:        ' .Borders(xlEdgeRight).Weight = xlThin
      End With
      
      With .Range(.Cells(5, 2), .Cells(4, 35))
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
      End With
   End With
   
   p_obj_Excel.Sheets(2).Range("E7").Select
   p_obj_Excel.ActiveWindow.FreezePanes = True
End Sub
'***************************************************************** REPORTE GENERAL - CARTA SERIEDAD OFERTA *****************************************************************
Private Sub fs_GenExc_TechoPropio_CSO(ByVal p_obj_Excel As Excel.Application, ByVal p_int_NroFil As Integer)
Dim r_int_ConAux        As Integer
Dim r_bol_Estado        As Boolean

'   p_int_NroFil = 5
   p_obj_Excel.Sheets(3).Select

   With p_obj_Excel.Sheets(3)
      .Cells(1, 2) = "REPORTE GENERAL DE CARTA SERIEDAD OFERTA"
      .Range(.Cells(1, 2), .Cells(1, 21)).Merge

      If Chk_FecAct.Value = 1 Then
         .Cells(3, 2) = " AL " & Format(date, "DD/MM/YYYY")
      Else
         .Cells(3, 2) = " AL " & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), ipp_PerAno.Text), "00") & " DE " & cmb_PerMes.Text & " DE " & Format(ipp_PerAno.Text, "0000")
      End If
      .Range(.Cells(3, 2), .Cells(3, 21)).Merge

      .Range(.Cells(1, 2), .Cells(1, 21)).Font.Bold = True
      .Range(.Cells(1, 2), .Cells(3, 21)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 2), .Cells(1, 21)).Font.Size = 14

      .Cells(p_int_NroFil, 2) = "TIPO - NRO. DOCUMENTO"
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 2)).Merge
      .Cells(p_int_NroFil, 3) = "RAZON SOCIAL"
      .Range(.Cells(p_int_NroFil, 3), .Cells(p_int_NroFil + 1, 3)).Merge
      .Cells(p_int_NroFil, 4) = "TIPO EMPRESA"
      .Range(.Cells(p_int_NroFil, 4), .Cells(p_int_NroFil + 1, 4)).Merge
      .Cells(p_int_NroFil, 5) = "MODALIDAD"
      .Range(.Cells(p_int_NroFil, 5), .Cells(p_int_NroFil + 1, 5)).Merge

      .Cells(p_int_NroFil, 6) = "NÚMERO"
      .Range(.Cells(p_int_NroFil, 6), .Cells(p_int_NroFil + 1, 6)).Merge
      .Cells(p_int_NroFil, 7) = "FECHA EMISIÓN"
      .Range(.Cells(p_int_NroFil, 7), .Cells(p_int_NroFil + 1, 7)).Merge
      .Cells(p_int_NroFil, 8) = "FECHA     VCTO."
      .Range(.Cells(p_int_NroFil, 8), .Cells(p_int_NroFil + 1, 8)).Merge

      .Cells(p_int_NroFil, 9) = "MONEDA"
      .Range(.Cells(p_int_NroFil, 9), .Cells(p_int_NroFil + 1, 9)).Merge
      .Cells(p_int_NroFil, 10) = "TASA ANUAL"
      .Range(.Cells(p_int_NroFil, 10), .Cells(p_int_NroFil + 1, 10)).Merge
      .Cells(p_int_NroFil, 11) = "PLAZO"
      .Range(.Cells(p_int_NroFil, 11), .Cells(p_int_NroFil + 1, 11)).Merge

      .Cells(p_int_NroFil, 12) = "VALOR" 'CARTA FIANZA
      .Range(.Cells(p_int_NroFil, 12), .Cells(p_int_NroFil + 1, 12)).Merge
      .Cells(p_int_NroFil, 13) = "GARANTIZADO"
      .Range(.Cells(p_int_NroFil, 13), .Cells(p_int_NroFil + 1, 13)).Merge
           
      .Cells(p_int_NroFil, 14) = "GARANTIA"
      .Range(.Cells(p_int_NroFil, 14), .Cells(p_int_NroFil, 16)).Merge
      .Cells(p_int_NroFil + 1, 14) = "IMPORTE"
      .Cells(p_int_NroFil + 1, 15) = "PAGADO"
      .Cells(p_int_NroFil + 1, 16) = "SALDO"
      
      .Cells(p_int_NroFil, 17) = "GARANTIA HIPOTECARIA"
      .Range(.Cells(p_int_NroFil, 17), .Cells(p_int_NroFil + 1, 17)).Merge
      
      .Cells(p_int_NroFil, 18) = "COMISIONES"
      .Range(.Cells(p_int_NroFil, 18), .Cells(p_int_NroFil, 21)).Merge
      .Cells(p_int_NroFil + 1, 18) = "IMPORTE"
      .Cells(p_int_NroFil + 1, 19) = "PAGADO"
      .Cells(p_int_NroFil + 1, 20) = "SALDO"
      .Cells(p_int_NroFil + 1, 21) = "FECHA PAGO"
      
      .Cells(p_int_NroFil, 22) = "FONDOS RECIBIDOS" ' - FMV
      .Range(.Cells(p_int_NroFil, 22), .Cells(p_int_NroFil, 24)).Merge
      .Cells(p_int_NroFil + 1, 22) = "IMPORTE"
      .Cells(p_int_NroFil + 1, 23) = "RECIBIDO"
      .Cells(p_int_NroFil + 1, 24) = "SALDO"
      .Cells(p_int_NroFil, 25) = "DESEMBOLSOS - ET"
      .Range(.Cells(p_int_NroFil, 25), .Cells(p_int_NroFil, 27)).Merge
      .Cells(p_int_NroFil + 1, 25) = "IMPORTE"
      .Cells(p_int_NroFil + 1, 26) = "PAGADO"
      .Cells(p_int_NroFil + 1, 27) = "SALDO"
      .Cells(p_int_NroFil, 28) = "ESTADO"
      .Range(.Cells(p_int_NroFil, 28), .Cells(p_int_NroFil + 1, 28)).Merge
      .Cells(p_int_NroFil, 29) = "Factor"
      .Range(.Cells(p_int_NroFil, 29), .Cells(p_int_NroFil + 1, 29)).Merge
      .Cells(p_int_NroFil, 30) = "Exposición"
      .Range(.Cells(p_int_NroFil, 30), .Cells(p_int_NroFil + 1, 30)).Merge
      .Cells(p_int_NroFil, 31) = "Tasa"
      .Range(.Cells(p_int_NroFil, 31), .Cells(p_int_NroFil + 1, 31)).Merge
      .Cells(p_int_NroFil, 32) = "Prov. Actual"
      .Range(.Cells(p_int_NroFil, 32), .Cells(p_int_NroFil + 1, 32)).Merge
      .Cells(p_int_NroFil, 33) = "Prov. Anterior"
      .Range(.Cells(p_int_NroFil, 33), .Cells(p_int_NroFil + 1, 33)).Merge
      .Cells(p_int_NroFil, 34) = "Diferencia"
      .Range(.Cells(p_int_NroFil, 34), .Cells(p_int_NroFil + 1, 34)).Merge
      .Cells(p_int_NroFil, 35) = "USUARIO"
      .Range(.Cells(p_int_NroFil, 35), .Cells(p_int_NroFil + 1, 35)).Merge
      
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 35)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 35)).Font.Bold = True
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 35)).HorizontalAlignment = xlHAlignCenter

      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 16.5
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 60
      .Columns("D").ColumnWidth = 17
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 17.5
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      
      .Columns("F").ColumnWidth = 13
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      
      .Columns("G").ColumnWidth = 13
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      
      .Columns("H").ColumnWidth = 13
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      
      .Columns("I").ColumnWidth = 21
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      
      .Columns("J").ColumnWidth = 13
      .Columns("J").NumberFormat = "0.00%"
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      
      .Columns("K").ColumnWidth = 13
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      
      .Columns("L").ColumnWidth = 13
      .Columns("L").NumberFormat = "###,###,###,##0.00"
      .Columns("L").HorizontalAlignment = xlHAlignRight
      
      .Columns("M").ColumnWidth = 13.5
      .Columns("M").NumberFormat = "###,###,###,##0.00"
      .Columns("M").HorizontalAlignment = xlHAlignRight
      
      .Columns("N").ColumnWidth = 13.5
      .Columns("N").NumberFormat = "###,###,###,##0.00"
      .Columns("N").HorizontalAlignment = xlHAlignRight
      
      .Columns("O").ColumnWidth = 13.5
      .Columns("O").NumberFormat = "###,###,###,##0.00"
      .Columns("O").HorizontalAlignment = xlHAlignRight
      
      .Columns("P").ColumnWidth = 13.5
      .Columns("P").NumberFormat = "###,###,###,##0.00"
      .Columns("P").HorizontalAlignment = xlHAlignRight
      
      .Columns("Q").ColumnWidth = 13.5
      .Columns("Q").HorizontalAlignment = xlHAlignCenter
      
      .Columns("R").ColumnWidth = 13.5
      .Columns("R").NumberFormat = "###,###,###,##0.00"
      .Columns("R").HorizontalAlignment = xlHAlignRight
      
      .Columns("S").ColumnWidth = 13.5
      .Columns("S").NumberFormat = "###,###,###,##0.00"
      .Columns("S").HorizontalAlignment = xlHAlignRight
      
      .Columns("T").ColumnWidth = 13.5
      .Columns("T").NumberFormat = "###,###,###,##0.00"
      .Columns("T").HorizontalAlignment = xlHAlignRight
      
      .Columns("U").ColumnWidth = 13
      .Columns("U").HorizontalAlignment = xlHAlignCenter
      
      .Columns("V").ColumnWidth = 13.5
      .Columns("V").NumberFormat = "###,###,###,##0.00"
      .Columns("V").HorizontalAlignment = xlHAlignRight
      
      .Columns("W").ColumnWidth = 13.5
      .Columns("W").NumberFormat = "###,###,###,##0.00"
      .Columns("W").HorizontalAlignment = xlHAlignRight
      
      .Columns("X").ColumnWidth = 13.5
      .Columns("X").NumberFormat = "###,###,###,##0.00"
      .Columns("X").HorizontalAlignment = xlHAlignRight
      
      .Columns("Y").ColumnWidth = 13.5
      .Columns("Y").NumberFormat = "###,###,###,##0.00"
      .Columns("Y").HorizontalAlignment = xlHAlignRight
      
      .Columns("Z").ColumnWidth = 13.5
      .Columns("Z").NumberFormat = "###,###,###,##0.00"
      .Columns("Z").HorizontalAlignment = xlHAlignRight
      
      .Columns("AA").ColumnWidth = 13.5
      .Columns("AA").NumberFormat = "###,###,###,##0.00"
      .Columns("AA").HorizontalAlignment = xlHAlignRight
      
      .Columns("AB").ColumnWidth = 20 '13.5
      .Columns("AB").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AC").ColumnWidth = 6
      .Columns("AC").NumberFormat = "0.00%"
      .Columns("AC").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AD").ColumnWidth = 13.5
      .Columns("AD").NumberFormat = "###,###,###,##0.00"
      .Columns("AD").HorizontalAlignment = xlHAlignRight
      
      .Columns("AE").ColumnWidth = 6
      .Columns("AE").NumberFormat = "0.00%"
      .Columns("AE").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AF").ColumnWidth = 13.5
      .Columns("AF").NumberFormat = "###,###,###,##0.00"
      .Columns("AF").HorizontalAlignment = xlHAlignRight
      
      .Columns("AG").ColumnWidth = 13.5
      .Columns("AG").NumberFormat = "###,###,###,##0.00"
      .Columns("AG").HorizontalAlignment = xlHAlignRight
      
      .Columns("AH").ColumnWidth = 13.5
      .Columns("AH").NumberFormat = "###,###,###,##0.00"
      .Columns("AH").HorizontalAlignment = xlHAlignRight
      
      .Columns("AI").ColumnWidth = 13.5
      .Columns("AI").HorizontalAlignment = xlHAlignCenter

      With .Range(.Cells(5, 2), .Cells(6, 35))
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
          .WrapText = True
      End With

      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(2, 1), .Cells(99, 99)).Font.Size = 11

      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "   SELECT RPT_PERMES   MES               , RPT_PERANO   ANNO               , RPT_DESCRI   DOCUMENTO             , RPT_VALCAD01 RAZON_SOCIAL   , "
      g_str_Parame = g_str_Parame & "          RPT_VALCAD02 TIPO_EMPRESA      , RPT_VALCAD03 MONEDA             , RPT_VALNUM25 TASA_ANUAL            , RPT_VALNUM01 PLAZO          , "
      g_str_Parame = g_str_Parame & "          RPT_VALCAD05 NUMREF            , RPT_VALCAD06 MAECFI_EMIFIA      , RPT_VALCAD07 MAECFI_VTOFIA         , RPT_VALNUM02 MAECFI_IMPFIA  , "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM03 MAECFI_GARFIA     , RPT_VALNUM04 IMPORTE_GARANTIA   , RPT_VALNUM05 PAGADO_GARANTIA       , RPT_VALNUM06 SALDO_GARANTIA , "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM07 IMPORTE_COMISION  , RPT_VALNUM08 PAGADO_COMISION    , RPT_VALNUM09 SALDO_COMISION        , RPT_VALNUM10 IMPORTE_FONDOS , "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM11 RECIBIDO_FONDOS   , RPT_VALNUM12 SALDO_FONDOS       , RPT_VALNUM13 IMPORTE_DESEMBOLSADO  ,   "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM14 PAGADO_DESEMBOLSO , RPT_VALNUM15 DEVOLUCION_GARANTIA, RPT_VALNUM16 SALDO_DESEMBOLSO      , RPT_VALCAD08 SITUACION    , "
      g_str_Parame = g_str_Parame & "          RPT_VALCAD09 FACTOR            , RPT_VALNUM17 EXPOSICION         , RPT_VALCAD10 TASA                  , RPT_VALNUM18 PROV_ACTUAL  , "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM19 PROV_ANTERIOR     , RPT_VALCAD11 TIPO_GARANTIA      , RPT_VALNUM23 MONTO_GARANTIA_LIQUIDA, RPT_VALNUM24 MONTO_GARANTIA_HIPOTECARIA, "
      g_str_Parame = g_str_Parame & "          RPT_VALCAD14 USUARIO           , RPT_VALCAD17 FECHA_PAG_COMISION "
      g_str_Parame = g_str_Parame & "     FROM RPT_TABLA_TEMP "
      g_str_Parame = g_str_Parame & "    WHERE RPT_PERMES = " & Month(Format(gf_FormatoFecha(CStr(l_str_FecPer)), "dd/mm/yyyy")) & "" 'Month(Now)
      g_str_Parame = g_str_Parame & "      AND RPT_PERANO = " & Year(Format(gf_FormatoFecha(CStr(l_str_FecPer)), "dd/mm/yyyy")) & " " 'Year(Now)
      g_str_Parame = g_str_Parame & "      AND RPT_TERCRE = " & "'" & modgen_g_str_NombPC & "'"
      g_str_Parame = g_str_Parame & "      AND RPT_USUCRE = " & "'" & modgen_g_str_CodUsu & "'"
      g_str_Parame = g_str_Parame & "      AND RPT_NOMBRE = " & "'" & Me.cmb_TipRep.Text & "2'"
      g_str_Parame = g_str_Parame & "      AND SUBSTR(RPT_VALCAD05,1,1) = 3 "
      g_str_Parame = g_str_Parame & "    ORDER BY RPT_VALCAD02, RPT_VALCAD06 "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If

      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         'g_rst_Princi.Close
         'Set g_rst_Princi = Nothing
         r_bol_Estado = True
'         Call fs_GenExc_TechoPropio_Comision(p_obj_Excel, 5) 'GoTo Cartas
         'MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
'         Exit Sub
      End If
      
      p_int_NroFil = p_int_NroFil + 2
      
      If Not g_rst_Princi.EOF Then
         g_rst_Princi.MoveFirst
   
         Do While Not g_rst_Princi.EOF
               If .Cells(p_int_NroFil - 1, 4) <> g_rst_Princi!TIPO_EMPRESA And .Cells(p_int_NroFil - 1, 4) <> "" Then
                  .Range(.Cells(p_int_NroFil, 30), .Cells(p_int_NroFil, 34)).FormulaR1C1 = "=SUM(R[-" & p_int_NroFil - r_int_ConAux - 6 & "]C:R[-1]C)"
                  .Range(.Cells(p_int_NroFil, 30), .Cells(p_int_NroFil, 34)).Font.Bold = True
                  .Range(.Cells(p_int_NroFil, 31), .Cells(p_int_NroFil, 31)).Value = ""
                  r_int_ConAux = p_int_NroFil - 3
                  p_int_NroFil = p_int_NroFil + 3
               End If
   
               .Cells(p_int_NroFil, 2) = g_rst_Princi!DOCUMENTO
               .Cells(p_int_NroFil, 3) = g_rst_Princi!RAZON_SOCIAL
               .Cells(p_int_NroFil, 4) = g_rst_Princi!TIPO_EMPRESA
               .Cells(p_int_NroFil, 5) = IIf(Mid(g_rst_Princi!NUMREF, 1, 1) = 2, "AD", IIf(Mid(g_rst_Princi!NUMREF, 1, 1) = 3, "CSO", "CF"))
               .Cells(p_int_NroFil, 6) = "'" & gf_Formato_NumRef(Trim(g_rst_Princi!NUMREF), Mid(g_rst_Princi!NUMREF, 1, 1))
               .Cells(p_int_NroFil, 7) = "'" & Format(gf_FormatoFecha(CStr(g_rst_Princi!MAECFI_EMIFIA)), "dd/mm/yyyy")
               .Cells(p_int_NroFil, 8) = "'" & Format(gf_FormatoFecha(CStr(g_rst_Princi!MAECFI_VTOFIA)), "dd/mm/yyyy")
               .Cells(p_int_NroFil, 9) = g_rst_Princi!Moneda
               .Cells(p_int_NroFil, 10) = g_rst_Princi!TASA_ANUAL / 100
               .Cells(p_int_NroFil, 11) = g_rst_Princi!PLAZO
               .Cells(p_int_NroFil, 12) = g_rst_Princi!MAECFI_IMPFIA
               .Cells(p_int_NroFil, 13) = g_rst_Princi!MAECFI_GARFIA
               .Cells(p_int_NroFil, 14) = g_rst_Princi!IMPORTE_GARANTIA
               .Cells(p_int_NroFil, 15) = g_rst_Princi!PAGADO_GARANTIA
               .Cells(p_int_NroFil, 16) = g_rst_Princi!SALDO_GARANTIA
               .Cells(p_int_NroFil, 17) = IIf(g_rst_Princi!MONTO_GARANTIA_HIPOTECARIA > 0, "SI", "")
               .Cells(p_int_NroFil, 18) = g_rst_Princi!IMPORTE_COMISION
               .Cells(p_int_NroFil, 19) = g_rst_Princi!PAGADO_COMISION
               .Cells(p_int_NroFil, 20) = g_rst_Princi!SALDO_COMISION
               If Not IsNull(g_rst_Princi!FECHA_PAG_COMISION) Then
                  .Cells(p_int_NroFil, 21) = "'" & Format(gf_FormatoFecha(CStr(g_rst_Princi!FECHA_PAG_COMISION)), "dd/mm/yyyy")
               Else
                  .Cells(p_int_NroFil, 21) = ""
               End If
               .Cells(p_int_NroFil, 22) = g_rst_Princi!IMPORTE_FONDOS
               .Cells(p_int_NroFil, 23) = g_rst_Princi!RECIBIDO_FONDOS
               .Cells(p_int_NroFil, 24) = g_rst_Princi!SALDO_FONDOS
               .Cells(p_int_NroFil, 25) = g_rst_Princi!IMPORTE_DESEMBOLSADO
               .Cells(p_int_NroFil, 26) = g_rst_Princi!PAGADO_DESEMBOLSO
               .Cells(p_int_NroFil, 27) = g_rst_Princi!SALDO_DESEMBOLSO
               .Cells(p_int_NroFil, 28) = g_rst_Princi!SITUACION
               .Cells(p_int_NroFil, 29) = g_rst_Princi!FACTOR
               .Cells(p_int_NroFil, 30) = g_rst_Princi!EXPOSICION
               .Cells(p_int_NroFil, 31) = g_rst_Princi!TASA
               .Cells(p_int_NroFil, 32) = g_rst_Princi!PROV_ACTUAL
               .Cells(p_int_NroFil, 33) = g_rst_Princi!PROV_ANTERIOR
               .Cells(p_int_NroFil, 34) = .Cells(p_int_NroFil, 32) - .Cells(p_int_NroFil, 33)
               .Cells(p_int_NroFil, 35) = g_rst_Princi!USUARIO
            p_int_NroFil = p_int_NroFil + 1
            g_rst_Princi.MoveNext
         Loop
      End If
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 35)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 35)).Font.Bold = True
      .Range(.Cells(p_int_NroFil, 5), .Cells(p_int_NroFil, 35)).HorizontalAlignment = xlHAlignRight

      'SUMATORIA TOTAL
      .Cells(p_int_NroFil, 3) = "TOTAL"
      .Cells(p_int_NroFil, 13).FormulaR1C1 = "=SUM(R[-" & p_int_NroFil - 5 & "]C:R[-1]C)"

      .Range(.Cells(p_int_NroFil, 30), .Cells(p_int_NroFil, 34)).FormulaR1C1 = "=SUM(R[-" & p_int_NroFil - r_int_ConAux - 4 & "]C:R[-1]C)" '-3
      .Range(.Cells(p_int_NroFil, 30), .Cells(p_int_NroFil, 34)).Font.Bold = True
      .Range(.Cells(p_int_NroFil, 31), .Cells(p_int_NroFil, 31)).Value = ""
      
      'RESUMEN POR TIPO EMPRESA
      .Range(.Cells(p_int_NroFil + 2, 2), .Cells(p_int_NroFil + 2, 5)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil + 2, 2), .Cells(p_int_NroFil + 2, 5)).Font.Bold = True
      .Cells(p_int_NroFil + 2, 2) = "CANTIDAD"
      .Cells(p_int_NroFil + 2, 3) = "DESCRIPCION"
      .Cells(p_int_NroFil + 2, 4) = "GARANTIZADO"
      .Cells(p_int_NroFil + 2, 5) = "PROVISION ACTUAL"

      .Range(.Cells(p_int_NroFil + 3, 3), .Cells(p_int_NroFil + 6, 3)).Font.Bold = True
      .Range(.Cells(p_int_NroFil + 3, 4), .Cells(p_int_NroFil + 6, 5)).NumberFormat = "###,###,###,##0.00"
      .Range(.Cells(p_int_NroFil + 3, 4), .Cells(p_int_NroFil + 6, 5)).HorizontalAlignment = xlHAlignRight

      .Cells(p_int_NroFil + 3, 2).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 4) & "]C[2]:R[-4]C[2],""PEQUEÑA"")"
      .Cells(p_int_NroFil + 3, 3) = "RESUMEN PEQUEÑA EMPRESA"
      .Cells(p_int_NroFil + 3, 4).FormulaR1C1 = "=SUMIF(R[-" & (p_int_NroFil - 4) & "]C:R[-4]C[9],""PEQUEÑA"",R[-" & (p_int_NroFil - 4) & "]C[9]:R[-4]C[9])"
      .Cells(p_int_NroFil + 3, 5).FormulaR1C1 = "=SUMIF(R[-" & (p_int_NroFil - 4) & "]C[-1]:R[-4]C[27],""PEQUEÑA"",R[-" & (p_int_NroFil - 4) & "]C[27]:R[-4]C[27])"

      .Cells(p_int_NroFil + 4, 2).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 3) & "]C[2]:R[-5]C[2],""MEDIANA"")"
      .Cells(p_int_NroFil + 4, 3) = "RESUMEN MEDIANA EMPRESA"
      .Cells(p_int_NroFil + 4, 4).FormulaR1C1 = "=SUMIF(R[-" & (p_int_NroFil - 3) & "]C:R[-5]C[9],""MEDIANA"",R[-" & (p_int_NroFil - 3) & "]C[9]:R[-5]C[9])"
      .Cells(p_int_NroFil + 4, 5).FormulaR1C1 = "=SUMIF(R[-" & (p_int_NroFil - 3) & "]C[-1]:R[-5]C[27],""MEDIANA"",R[-" & (p_int_NroFil - 3) & "]C[27]:R[-5]C[27])"

      .Cells(p_int_NroFil + 5, 2).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 2) & "]C[2]:R[-6]C[2],""GRANDE"")"
      .Cells(p_int_NroFil + 5, 3) = "RESUMEN GRANDE EMPRESA"
      .Cells(p_int_NroFil + 5, 4).FormulaR1C1 = "=+SUMIF(R[-" & (p_int_NroFil - 2) & "]C:R[-6]C[9],""GRANDE"",R[-" & (p_int_NroFil - 2) & "]C[9]:R[-6]C[9])"
      .Cells(p_int_NroFil + 5, 5).FormulaR1C1 = "=+SUMIF(R[-" & (p_int_NroFil - 2) & "]C[-1]:R[-6]C[27],""GRANDE"",R[-" & (p_int_NroFil - 2) & "]C[27]:R[-6]C[27])"
      
      .Cells(p_int_NroFil + 6, 2).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 1) & "]C[2]:R[-7]C[2],""MICRO"")"
      .Cells(p_int_NroFil + 6, 3) = "RESUMEN MICRO EMPRESA"
      .Cells(p_int_NroFil + 6, 4).FormulaR1C1 = "=+SUMIF(R[-" & (p_int_NroFil - 1) & "]C:R[-7]C[9],""MICRO"",R[-" & (p_int_NroFil - 1) & "]C[9]:R[-7]C[9])"
      .Cells(p_int_NroFil + 6, 5).FormulaR1C1 = "=+SUMIF(R[-" & (p_int_NroFil - 1) & "]C[-1]:R[-7]C[27],""MICRO"",R[-" & (p_int_NroFil - 1) & "]C[27]:R[-7]C[27])"
      
      .Cells(p_int_NroFil + 7, 2).FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
      .Cells(p_int_NroFil + 7, 3) = "TOTAL"
      .Range(.Cells(p_int_NroFil + 7, 4), .Cells(p_int_NroFil + 7, 5)).FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
      
      .Range(.Cells(p_int_NroFil + 7, 4), .Cells(p_int_NroFil + 7, 5)).HorizontalAlignment = xlHAlignRight
      .Range(.Cells(p_int_NroFil + 7, 2), .Cells(p_int_NroFil + 7, 7)).Font.Bold = True
      
      With .Range(.Cells(1, 2), .Cells(1, 35))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
      End With

      With .Range(.Cells(5, 2), .Cells(p_int_NroFil - 1, 35))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With

      With .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 35))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With

      With .Range(.Cells(5, 14), .Cells(p_int_NroFil, 16))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With

      With .Range(.Cells(5, 18), .Cells(p_int_NroFil, 21))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:    .Borders(xlInsideHorizontal).LineStyle = xlContinuous:
      End With

      With .Range(.Cells(5, 22), .Cells(p_int_NroFil, 24))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With

      With .Range(.Cells(5, 25), .Cells(p_int_NroFil, 27))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      'RESUMEN
      With .Range(.Cells(p_int_NroFil + 2, 2), .Cells(p_int_NroFil + 6, 5))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlThin
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlThin
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlThin
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlThin
      End With
      
      With .Range(.Cells(p_int_NroFil + 7, 2), .Cells(p_int_NroFil + 7, 5))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:         ' .Borders(xlEdgeLeft).Weight = xlThin
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           '.Borders(xlEdgeTop).Weight = xlThin
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:       ' .Borders(xlEdgeBottom).Weight = xlThin
          .Borders(xlEdgeRight).LineStyle = xlContinuous:        ' .Borders(xlEdgeRight).Weight = xlThin
      End With
      
      With .Range(.Cells(5, 2), .Cells(4, 35))
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
      End With
   End With
   
   p_obj_Excel.Sheets(3).Range("E7").Select
   p_obj_Excel.ActiveWindow.FreezePanes = True
 End Sub
'********************************************************** COMISIONES ************************************************************************
Private Sub fs_GenExc_TechoPropio_Comision(ByVal p_obj_Excel As Excel.Application, ByVal p_int_NroFil As Integer)
Dim r_bol_Estado        As Boolean
'   p_int_NroFil = 5
   p_obj_Excel.Sheets(7).Select
   
   '*************************************************************** REPORTE DE COMISIONES DE CARTAS FIANZA Y ADENDAS **************************************************************
   With p_obj_Excel.Sheets(7)
   
      .Cells(1, 2) = "REPORTE GENERAL DE COMISIONES DE CARTAS FIANZA Y ADENDAS"
      .Range(.Cells(1, 2), .Cells(1, 16)).Merge
      
      If Chk_FecAct.Value = 1 Then
         .Cells(3, 2) = " AL " & Format(date, "DD/MM/YYYY")
      Else
         .Cells(3, 2) = " AL " & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), ipp_PerAno.Text), "00") & " DE " & cmb_PerMes.Text & " DE " & Format(ipp_PerAno.Text, "0000")
      End If
      .Range(.Cells(3, 2), .Cells(3, 16)).Merge
      
      .Range(.Cells(1, 2), .Cells(1, 16)).Font.Bold = True
      .Range(.Cells(1, 2), .Cells(3, 16)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 2), .Cells(1, 16)).Font.Size = 14
      
      .Cells(p_int_NroFil, 2) = "TIPO - NRO. DOCUMENTO"
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 2)).Merge
      
      .Cells(p_int_NroFil, 3) = "RAZON SOCIAL"
      .Range(.Cells(p_int_NroFil, 3), .Cells(p_int_NroFil + 1, 3)).Merge
      
      .Cells(p_int_NroFil, 4) = "TIPO EMPRESA"
      .Range(.Cells(p_int_NroFil, 4), .Cells(p_int_NroFil + 1, 4)).Merge
      
      .Cells(p_int_NroFil, 5) = "MODALIDAD"
      .Range(.Cells(p_int_NroFil, 5), .Cells(p_int_NroFil + 1, 5)).Merge
      
      .Cells(p_int_NroFil, 6) = "NÚMERO"
      .Range(.Cells(p_int_NroFil, 6), .Cells(p_int_NroFil + 1, 6)).Merge
      
      .Cells(p_int_NroFil, 7) = "FECHA EMISIÓN"
      .Range(.Cells(p_int_NroFil, 7), .Cells(p_int_NroFil + 1, 7)).Merge
      
      .Cells(p_int_NroFil, 8) = "FECHA     VCTO."
      .Range(.Cells(p_int_NroFil, 8), .Cells(p_int_NroFil + 1, 8)).Merge
      
      .Cells(p_int_NroFil, 9) = "MONEDA"
      .Range(.Cells(p_int_NroFil, 9), .Cells(p_int_NroFil + 1, 9)).Merge
      
      .Cells(p_int_NroFil, 10) = "TASA ANUAL"
      .Range(.Cells(p_int_NroFil, 10), .Cells(p_int_NroFil + 1, 10)).Merge
      
      .Cells(p_int_NroFil, 11) = "PLAZO"
      .Range(.Cells(p_int_NroFil, 11), .Cells(p_int_NroFil + 1, 11)).Merge
      
      .Cells(p_int_NroFil, 12) = "VALOR"
      .Range(.Cells(p_int_NroFil, 12), .Cells(p_int_NroFil + 1, 12)).Merge
      
      .Cells(p_int_NroFil, 13) = "GARANTIZADO"
      .Range(.Cells(p_int_NroFil, 13), .Cells(p_int_NroFil + 1, 13)).Merge
      
      .Cells(p_int_NroFil, 14) = "COMISION"
      .Range(.Cells(p_int_NroFil, 14), .Cells(p_int_NroFil + 1, 14)).Merge
      
      .Cells(p_int_NroFil, 15) = "F. PAGO COMISION"
      .Range(.Cells(p_int_NroFil, 15), .Cells(p_int_NroFil + 1, 15)).Merge
            
      .Cells(p_int_NroFil, 16) = "ESTADO"
      .Range(.Cells(p_int_NroFil, 16), .Cells(p_int_NroFil + 1, 16)).Merge
      
      .Cells(p_int_NroFil, 17) = "USUARIO"
      .Range(.Cells(p_int_NroFil, 17), .Cells(p_int_NroFil + 1, 17)).Merge
      
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 17)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 17)).Font.Bold = True
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 17)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").ColumnWidth = 1
      
      .Columns("B").ColumnWidth = 16.5
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      
      .Columns("C").ColumnWidth = 60
      
      .Columns("D").ColumnWidth = 17
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      
      .Columns("E").ColumnWidth = 13
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      
      .Columns("F").ColumnWidth = 13
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      
      .Columns("G").ColumnWidth = 13
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      
      .Columns("H").ColumnWidth = 13
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      
      .Columns("I").ColumnWidth = 18
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      
      .Columns("J").ColumnWidth = 13
      .Columns("J").NumberFormat = "0.00%"
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      
      .Columns("K").ColumnWidth = 13
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      
      .Columns("L").ColumnWidth = 13
      .Columns("L").NumberFormat = "###,###,###,##0.00"
      .Columns("L").HorizontalAlignment = xlHAlignRight
      
      .Columns("M").ColumnWidth = 13.5
      .Columns("M").NumberFormat = "###,###,###,##0.00"
      .Columns("M").HorizontalAlignment = xlHAlignRight
      
      .Columns("N").ColumnWidth = 13.5
      .Columns("N").NumberFormat = "###,###,###,##0.00"
      .Columns("N").HorizontalAlignment = xlHAlignRight
      
      .Columns("O").ColumnWidth = 13
      .Columns("O").HorizontalAlignment = xlHAlignCenter
      
      .Columns("P").ColumnWidth = 20 '15
      .Columns("P").HorizontalAlignment = xlHAlignCenter
      
      .Columns("Q").ColumnWidth = 15
      .Columns("Q").HorizontalAlignment = xlHAlignCenter
     
      With .Range(.Cells(5, 2), .Cells(6, 17))
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
          .WrapText = True
      End With
      
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(2, 1), .Cells(99, 99)).Font.Size = 11
            
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "   SELECT DISTINCT RPT_PERMES   MES      , RPT_PERANO   ANNO               , RPT_DESCRI   REFERENCIA          , RPT_VALCAD01 DOCUMENTO      , "
      g_str_Parame = g_str_Parame & "          RPT_VALCAD02 RAZON_SOCIAL      , RPT_VALCAD03 TIPO_EMPRESA       , RPT_VALCAD04 MONEDA              , RPT_VALNUM01 TASA_ANUAL     , "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM02 PLAZO             , RPT_VALCAD05 FECHA_EMISION      , RPT_VALCAD06 FECHA_VENCIMIENTO   , RPT_VALNUM03 VALOR_CARTA_FIANZA  , RPT_VALNUM04 GARANTIZADO    , "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM05 IMPORTE_COMISION  , RPT_VALCAD07 SITUACION          , RPT_VALCAD08 USUARIO             , RPT_VALCAD09 FECHA_PAGO_COMISION "
      g_str_Parame = g_str_Parame & "     FROM RPT_TABLA_TEMP "
      g_str_Parame = g_str_Parame & "    WHERE RPT_PERMES = " & Month(Format(gf_FormatoFecha(CStr(l_str_FecPer)), "dd/mm/yyyy")) & ""
      g_str_Parame = g_str_Parame & "      AND RPT_PERANO = " & Year(Format(gf_FormatoFecha(CStr(l_str_FecPer)), "dd/mm/yyyy")) & " "
      g_str_Parame = g_str_Parame & "      AND RPT_TERCRE = " & "'" & modgen_g_str_NombPC & "'"
      g_str_Parame = g_str_Parame & "      AND RPT_USUCRE = " & "'" & modgen_g_str_CodUsu & "'"
      g_str_Parame = g_str_Parame & "      AND RPT_NOMBRE = " & "'" & Me.cmb_TipRep.Text & " COMISION'"
      g_str_Parame = g_str_Parame & "    ORDER BY RPT_VALCAD02, RPT_VALCAD06 "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
          
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         'g_rst_Princi.Close
         'Set g_rst_Princi = Nothing
         r_bol_Estado = True
'         Call fs_GenExc_CDir_LC(p_obj_Excel, 5)
'         MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
'         Exit Sub
      End If
      
      p_int_NroFil = p_int_NroFil + 2
      
      If Not g_rst_Princi.EOF Then
         g_rst_Princi.MoveFirst
         
         Do While Not g_rst_Princi.EOF
               
             .Cells(p_int_NroFil, 2) = g_rst_Princi!DOCUMENTO
             .Cells(p_int_NroFil, 3) = g_rst_Princi!RAZON_SOCIAL
             .Cells(p_int_NroFil, 4) = g_rst_Princi!TIPO_EMPRESA
             .Cells(p_int_NroFil, 5) = IIf(Mid(g_rst_Princi!REFERENCIA, 1, 1) = 2, "AD", "CF")
             If Trim(g_rst_Princi!REFERENCIA) = "008" Then
               .Cells(p_int_NroFil, 6) = "'" & gf_Formato_NumRef(Trim(g_rst_Princi!REFERENCIA), Mid(g_rst_Princi!REFERENCIA, 1, 1))
             Else
               .Cells(p_int_NroFil, 6) = "'" & gf_Formato_NumRef(Trim(g_rst_Princi!REFERENCIA), 1)
             End If
             .Cells(p_int_NroFil, 7) = "'" & Format(gf_FormatoFecha(CStr(g_rst_Princi!FECHA_EMISION)), "dd/mm/yyyy")
             .Cells(p_int_NroFil, 8) = "'" & Format(gf_FormatoFecha(CStr(g_rst_Princi!FECHA_VENCIMIENTO)), "dd/mm/yyyy")
             .Cells(p_int_NroFil, 9) = g_rst_Princi!Moneda
             .Cells(p_int_NroFil, 10) = g_rst_Princi!TASA_ANUAL / 100
             .Cells(p_int_NroFil, 11) = g_rst_Princi!PLAZO
             .Cells(p_int_NroFil, 12) = g_rst_Princi!VALOR_CARTA_FIANZA
             .Cells(p_int_NroFil, 13) = g_rst_Princi!GARANTIZADO
             .Cells(p_int_NroFil, 14) = g_rst_Princi!IMPORTE_COMISION
             If Not IsNull(g_rst_Princi!FECHA_PAGO_COMISION) Then
               .Cells(p_int_NroFil, 15) = "'" & Format(gf_FormatoFecha(CStr(g_rst_Princi!FECHA_PAGO_COMISION)), "dd/mm/yyyy")
             Else
               .Cells(p_int_NroFil, 15) = ""
             End If
             .Cells(p_int_NroFil, 16) = g_rst_Princi!SITUACION
             .Cells(p_int_NroFil, 17) = g_rst_Princi!USUARIO
         
            p_int_NroFil = p_int_NroFil + 1
            g_rst_Princi.MoveNext
         Loop
      End If
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 17)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 17)).Font.Bold = True
      .Range(.Cells(p_int_NroFil, 5), .Cells(p_int_NroFil, 17)).HorizontalAlignment = xlHAlignRight
            
      'SUMATORIA TOTAL
      .Cells(p_int_NroFil, 3) = "TOTAL"
      .Range(.Cells(p_int_NroFil, 12), .Cells(p_int_NroFil, 14)).FormulaR1C1 = "=SUM(R[-" & p_int_NroFil - 5 & "]C:R[-1]C)"

      .Range(.Cells(p_int_NroFil + 3, 3), .Cells(p_int_NroFil + 5, 3)).Font.Bold = True
      .Range(.Cells(p_int_NroFil + 3, 4), .Cells(p_int_NroFil + 5, 4)).NumberFormat = "###,###,###,##0.00"
      .Range(.Cells(p_int_NroFil + 3, 4), .Cells(p_int_NroFil + 5, 4)).HorizontalAlignment = xlHAlignRight
      
      With .Range(.Cells(1, 2), .Cells(1, 17))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
      End With
      
      With .Range(.Cells(5, 2), .Cells(p_int_NroFil - 1, 17))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
      With .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 17))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
             
      With .Range(.Cells(5, 2), .Cells(4, 17))
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
      End With
      
   End With
   
   p_obj_Excel.Sheets(7).Range("E7").Select
   p_obj_Excel.ActiveWindow.FreezePanes = True
  
'   p_obj_Excel.Sheets(1).Select
'
'   If r_bol_Estado = True Then
'      p_obj_Excel.Sheets(2).Select
'      p_obj_Excel.ActiveWindow.SelectedSheets.Visible = False
'      p_obj_Excel.Sheets(1).Select
'   End If
'
'   p_obj_Excel.Visible = True
End Sub
 '*************************************************************** REPORTE GENERAL - CARTAS FIANZA **************************************************************
Private Sub fs_GenExc_CDir_LC(ByVal p_obj_Excel As Excel.Application, ByVal p_int_NroFil As Integer)
Dim r_int_ConAux        As Integer
Dim r_bol_Estado        As Boolean
'Cartas:
 '  p_int_NroFil = 5
   p_obj_Excel.Sheets(10).Select
   
   With p_obj_Excel.Sheets(10)
   
      .Cells(1, 2) = "REPORTE GENERAL DE CARTAS FIANZA"
      .Range(.Cells(1, 2), .Cells(1, 21)).Merge
      
      If Chk_FecAct.Value = 1 Then
         .Cells(3, 2) = " AL " & Format(date, "DD/MM/YYYY")
      Else
         .Cells(3, 2) = " AL " & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), ipp_PerAno.Text), "00") & " DE " & cmb_PerMes.Text & " DE " & Format(ipp_PerAno.Text, "0000")
      End If
      .Range(.Cells(3, 2), .Cells(3, 21)).Merge
      
      .Range(.Cells(1, 2), .Cells(1, 21)).Font.Bold = True
      .Range(.Cells(1, 2), .Cells(3, 21)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 2), .Cells(1, 21)).Font.Size = 14
      
      .Cells(p_int_NroFil, 2) = "TIPO - NRO. DOCUMENTO"
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 2)).Merge
      .Cells(p_int_NroFil, 3) = "RAZON SOCIAL"
      .Range(.Cells(p_int_NroFil, 3), .Cells(p_int_NroFil + 1, 3)).Merge
      .Cells(p_int_NroFil, 4) = "TIPO EMPRESA"
      .Range(.Cells(p_int_NroFil, 4), .Cells(p_int_NroFil + 1, 4)).Merge
      .Cells(p_int_NroFil, 5) = "MODALIDAD"
      .Range(.Cells(p_int_NroFil, 5), .Cells(p_int_NroFil + 1, 5)).Merge
      
      .Cells(p_int_NroFil, 6) = "NÚMERO"
      .Range(.Cells(p_int_NroFil, 6), .Cells(p_int_NroFil + 1, 6)).Merge
      .Cells(p_int_NroFil, 7) = "FECHA EMISIÓN"
      .Range(.Cells(p_int_NroFil, 7), .Cells(p_int_NroFil + 1, 7)).Merge
      .Cells(p_int_NroFil, 8) = "FECHA     VCTO."
      .Range(.Cells(p_int_NroFil, 8), .Cells(p_int_NroFil + 1, 8)).Merge
      
      .Cells(p_int_NroFil, 9) = "MONEDA"
      .Range(.Cells(p_int_NroFil, 9), .Cells(p_int_NroFil + 1, 9)).Merge
      .Cells(p_int_NroFil, 10) = "TASA ANUAL"
      .Range(.Cells(p_int_NroFil, 10), .Cells(p_int_NroFil + 1, 10)).Merge
      .Cells(p_int_NroFil, 11) = "PLAZO"
      .Range(.Cells(p_int_NroFil, 11), .Cells(p_int_NroFil + 1, 11)).Merge
      
      .Cells(p_int_NroFil, 12) = "MONTO PRESTAMO" 'VALOR CARTA FIANZA
      .Range(.Cells(p_int_NroFil, 12), .Cells(p_int_NroFil + 1, 12)).Merge
      .Cells(p_int_NroFil, 13) = "GARANTIZADO"
      .Range(.Cells(p_int_NroFil, 13), .Cells(p_int_NroFil + 1, 13)).Merge
           
      .Cells(p_int_NroFil, 14) = "GARANTIA"
      .Range(.Cells(p_int_NroFil, 14), .Cells(p_int_NroFil, 16)).Merge
      .Cells(p_int_NroFil + 1, 14) = "IMPORTE"
      .Cells(p_int_NroFil + 1, 15) = "PAGADO"
      .Cells(p_int_NroFil + 1, 16) = "SALDO"
      
      .Cells(p_int_NroFil, 17) = "GARANTIA HIPOTECARIA"
      .Range(.Cells(p_int_NroFil, 17), .Cells(p_int_NroFil + 1, 17)).Merge
      
      .Cells(p_int_NroFil, 18) = "COMISIONES"
      .Range(.Cells(p_int_NroFil, 18), .Cells(p_int_NroFil, 20)).Merge
      .Cells(p_int_NroFil + 1, 18) = "IMPORTE"
      .Cells(p_int_NroFil + 1, 19) = "PAGADO"
      .Cells(p_int_NroFil + 1, 20) = "SALDO"
      .Cells(p_int_NroFil, 21) = "FONDOS RECIBIDOS" ' - FMV
      .Range(.Cells(p_int_NroFil, 21), .Cells(p_int_NroFil, 23)).Merge
      .Cells(p_int_NroFil + 1, 21) = "IMPORTE"
      .Cells(p_int_NroFil + 1, 22) = "RECIBIDO"
      .Cells(p_int_NroFil + 1, 23) = "SALDO"
      .Cells(p_int_NroFil, 24) = "DESEMBOLSOS - ET"
      .Range(.Cells(p_int_NroFil, 24), .Cells(p_int_NroFil, 26)).Merge
      .Cells(p_int_NroFil + 1, 24) = "IMPORTE"
      .Cells(p_int_NroFil + 1, 25) = "PAGADO"
      .Cells(p_int_NroFil + 1, 26) = "SALDO"
      
      .Cells(p_int_NroFil, 27) = "SALDO LINEA CREDITO"
      .Range(.Cells(p_int_NroFil, 27), .Cells(p_int_NroFil + 1, 27)).Merge
      .Cells(p_int_NroFil, 28) = "SALDO CAPITAL"
      .Range(.Cells(p_int_NroFil, 28), .Cells(p_int_NroFil + 1, 28)).Merge
      .Cells(p_int_NroFil, 29) = "INTERES DEVENGADO"
      .Range(.Cells(p_int_NroFil, 29), .Cells(p_int_NroFil + 1, 29)).Merge
      '27
      .Cells(p_int_NroFil, 30) = "ESTADO"
      .Range(.Cells(p_int_NroFil, 30), .Cells(p_int_NroFil + 1, 30)).Merge
      .Cells(p_int_NroFil, 31) = "Factor"
      .Range(.Cells(p_int_NroFil, 31), .Cells(p_int_NroFil + 1, 31)).Merge
      .Cells(p_int_NroFil, 32) = "Exposición"
      .Range(.Cells(p_int_NroFil, 32), .Cells(p_int_NroFil + 1, 32)).Merge
      .Cells(p_int_NroFil, 33) = "Tasa"
      .Range(.Cells(p_int_NroFil, 33), .Cells(p_int_NroFil + 1, 33)).Merge
      .Cells(p_int_NroFil, 34) = "Prov. Gen. Actual"
      .Range(.Cells(p_int_NroFil, 34), .Cells(p_int_NroFil + 1, 34)).Merge
      .Cells(p_int_NroFil, 35) = "Prov. Gen. Anterior"
      .Range(.Cells(p_int_NroFil, 35), .Cells(p_int_NroFil + 1, 35)).Merge
      .Cells(p_int_NroFil, 36) = "Diferencia"
      .Range(.Cells(p_int_NroFil, 36), .Cells(p_int_NroFil + 1, 36)).Merge
      .Cells(p_int_NroFil, 37) = "USUARIO"
      .Range(.Cells(p_int_NroFil, 37), .Cells(p_int_NroFil + 1, 37)).Merge
      
      
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 37)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 37)).Font.Bold = True
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 37)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 16.5
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 60
      .Columns("D").ColumnWidth = 17
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 13
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      
      .Columns("F").ColumnWidth = 13
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      
      .Columns("G").ColumnWidth = 13
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      
      .Columns("H").ColumnWidth = 13
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      
      .Columns("I").ColumnWidth = 21
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      
      .Columns("J").ColumnWidth = 13
      .Columns("J").NumberFormat = "0.00%"
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      
      .Columns("K").ColumnWidth = 13
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      
      .Columns("L").ColumnWidth = 13
      .Columns("L").NumberFormat = "###,###,###,##0.00"
      .Columns("L").HorizontalAlignment = xlHAlignRight
      
      .Columns("M").ColumnWidth = 13.5
      .Columns("M").NumberFormat = "###,###,###,##0.00"
      .Columns("M").HorizontalAlignment = xlHAlignRight
      
      .Columns("N").ColumnWidth = 13.5
      .Columns("N").NumberFormat = "###,###,###,##0.00"
      .Columns("N").HorizontalAlignment = xlHAlignRight
      
      .Columns("O").ColumnWidth = 13.5
      .Columns("O").NumberFormat = "###,###,###,##0.00"
      .Columns("O").HorizontalAlignment = xlHAlignRight
      
      .Columns("P").ColumnWidth = 13.5
      .Columns("P").NumberFormat = "###,###,###,##0.00"
      .Columns("P").HorizontalAlignment = xlHAlignRight
      
      .Columns("Q").ColumnWidth = 13.5
      .Columns("Q").HorizontalAlignment = xlHAlignCenter
      
      .Columns("R").ColumnWidth = 13.5
      .Columns("R").NumberFormat = "###,###,###,##0.00"
      .Columns("R").HorizontalAlignment = xlHAlignRight
      
      .Columns("S").ColumnWidth = 13.5
      .Columns("S").NumberFormat = "###,###,###,##0.00"
      .Columns("S").HorizontalAlignment = xlHAlignRight
      
      .Columns("T").ColumnWidth = 13.5
      .Columns("T").NumberFormat = "###,###,###,##0.00"
      .Columns("T").HorizontalAlignment = xlHAlignRight
      
      .Columns("U").ColumnWidth = 13.5
      .Columns("U").NumberFormat = "###,###,###,##0.00"
      .Columns("U").HorizontalAlignment = xlHAlignRight
      
      .Columns("V").ColumnWidth = 13.5
      .Columns("V").NumberFormat = "###,###,###,##0.00"
      .Columns("V").HorizontalAlignment = xlHAlignRight
      
      .Columns("W").ColumnWidth = 13.5
      .Columns("W").NumberFormat = "###,###,###,##0.00"
      .Columns("W").HorizontalAlignment = xlHAlignRight
      
      .Columns("X").ColumnWidth = 13.5
      .Columns("X").NumberFormat = "###,###,###,##0.00"
      .Columns("X").HorizontalAlignment = xlHAlignRight
      
      .Columns("Y").ColumnWidth = 13.5
      .Columns("Y").NumberFormat = "###,###,###,##0.00"
      .Columns("Y").HorizontalAlignment = xlHAlignRight
      
      .Columns("Z").ColumnWidth = 13.5
      .Columns("Z").NumberFormat = "###,###,###,##0.00"
      .Columns("Z").HorizontalAlignment = xlHAlignRight
      
      .Columns("AA").ColumnWidth = 13.5
      .Columns("AA").NumberFormat = "###,###,###,##0.00"
      .Columns("AA").HorizontalAlignment = xlHAlignRight
      
      .Columns("AB").ColumnWidth = 13.5
      .Columns("AB").NumberFormat = "###,###,###,##0.00"
      .Columns("AB").HorizontalAlignment = xlHAlignRight
      
      .Columns("AC").ColumnWidth = 13.5
      .Columns("AC").NumberFormat = "###,###,###,##0.00"
      .Columns("AC").HorizontalAlignment = xlHAlignRight
      
      .Columns("AD").ColumnWidth = 13.5
      .Columns("AD").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AE").ColumnWidth = 6
      .Columns("AE").NumberFormat = "0.00%"
      .Columns("AE").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AF").ColumnWidth = 13.5
      .Columns("AF").NumberFormat = "###,###,###,##0.00"
      .Columns("AF").HorizontalAlignment = xlHAlignRight
      
      .Columns("AG").ColumnWidth = 6
      .Columns("AG").NumberFormat = "0.00%"
      .Columns("AG").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AH").ColumnWidth = 13.5
      .Columns("AH").NumberFormat = "###,###,###,##0.00"
      .Columns("AH").HorizontalAlignment = xlHAlignRight
      
      .Columns("AI").ColumnWidth = 13.5
      .Columns("AI").NumberFormat = "###,###,###,##0.00"
      .Columns("AI").HorizontalAlignment = xlHAlignRight
      
      .Columns("AJ").ColumnWidth = 13.5
      .Columns("AJ").NumberFormat = "###,###,###,##0.00"
      .Columns("AJ").HorizontalAlignment = xlHAlignRight
      
      .Columns("AK").ColumnWidth = 13.5
      .Columns("AK").HorizontalAlignment = xlHAlignCenter
      
      
      With .Range(.Cells(5, 2), .Cells(6, 37))
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
          .WrapText = True
      End With
      
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(2, 1), .Cells(99, 99)).Font.Size = 11
            
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "   SELECT RPT_PERMES   MES               , RPT_PERANO   ANNO               , RPT_DESCRI   DOCUMENTO             , RPT_VALCAD01 RAZON_SOCIAL   , "
      g_str_Parame = g_str_Parame & "          RPT_VALCAD02 TIPO_EMPRESA      , RPT_VALCAD03 MONEDA             , RPT_VALNUM25 TASA_ANUAL            , RPT_VALNUM01 PLAZO          , "
      g_str_Parame = g_str_Parame & "          RPT_VALCAD05 NUMREF            , RPT_VALCAD06 MAECFI_EMIFIA      , RPT_VALCAD07 MAECFI_VTOFIA         , RPT_VALNUM02 MAECFI_IMPFIA  , "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM03 MAECFI_GARFIA     , RPT_VALNUM04 IMPORTE_GARANTIA   , RPT_VALNUM05 PAGADO_GARANTIA       , RPT_VALNUM06 SALDO_GARANTIA , "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM07 IMPORTE_COMISION  , RPT_VALNUM08 PAGADO_COMISION    , RPT_VALNUM09 SALDO_COMISION        , RPT_VALNUM10 IMPORTE_FONDOS , "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM11 RECIBIDO_FONDOS   , RPT_VALNUM12 SALDO_FONDOS       , RPT_VALNUM13 IMPORTE_DESEMBOLSADO  ,   "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM14 PAGADO_DESEMBOLSO , RPT_VALNUM15 DEVOLUCION_GARANTIA, RPT_VALNUM16 SALDO_DESEMBOLSO      , RPT_VALCAD08 SITUACION    , "
      g_str_Parame = g_str_Parame & "          RPT_VALCAD09 FACTOR            , RPT_VALNUM17 EXPOSICION         , RPT_VALCAD10 TASA                  , RPT_VALNUM18 PROV_ACTUAL  , "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM19 PROV_ANTERIOR     , RPT_VALCAD11 TIPO_GARANTIA      , RPT_VALNUM23 MONTO_GARANTIA_LIQUIDA, RPT_VALNUM24 MONTO_GARANTIA_HIPOTECARIA , "
      g_str_Parame = g_str_Parame & "          RPT_VALCAD14 USUARIO           , RPT_VALCAD13 MODALIDAD          , RPT_VALNUM26 INTERES_ACUMULADO     , RPT_VALNUM27 LINEA_CREDITO, "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM28 SALDO_CAPITAL "
      g_str_Parame = g_str_Parame & "     FROM RPT_TABLA_TEMP "
      g_str_Parame = g_str_Parame & "    WHERE RPT_PERMES = " & Month(Format(gf_FormatoFecha(CStr(l_str_FecPer)), "dd/mm/yyyy")) & "" 'Month(Now)
      g_str_Parame = g_str_Parame & "      AND RPT_PERANO = " & Year(Format(gf_FormatoFecha(CStr(l_str_FecPer)), "dd/mm/yyyy")) & " " 'Year(Now)
      g_str_Parame = g_str_Parame & "      AND RPT_TERCRE = " & "'" & modgen_g_str_NombPC & "'"
      g_str_Parame = g_str_Parame & "      AND RPT_USUCRE = " & "'" & modgen_g_str_CodUsu & "'"
      g_str_Parame = g_str_Parame & "      AND RPT_NOMBRE = " & "'" & Me.cmb_TipRep.Text & "2'"
      g_str_Parame = g_str_Parame & "      AND RPT_VALCAD11 = '008'"
      g_str_Parame = g_str_Parame & "      AND RPT_VALCAD12 = '008'"
      g_str_Parame = g_str_Parame & "      AND RPT_VALCAD13 = '001'"
      g_str_Parame = g_str_Parame & "    ORDER BY RPT_VALCAD02, RPT_VALCAD01, RPT_VALCAD06 "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
          
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         'g_rst_Princi.Close
         'Set g_rst_Princi = Nothing
         r_bol_Estado = True
'         MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
         'Exit Sub
      End If
      
      p_int_NroFil = p_int_NroFil + 2
      
      If Not g_rst_Princi.EOF Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
            If .Cells(p_int_NroFil - 1, 4) <> g_rst_Princi!TIPO_EMPRESA And .Cells(p_int_NroFil - 1, 4) <> "" Then
               .Range(.Cells(p_int_NroFil, 27), .Cells(p_int_NroFil, 29)).FormulaR1C1 = "=SUM(R[-" & p_int_NroFil - 6 & "]C:R[-1]C)"
               .Range(.Cells(p_int_NroFil, 27), .Cells(p_int_NroFil, 29)).Font.Bold = True
               
               .Range(.Cells(p_int_NroFil, 33), .Cells(p_int_NroFil, 36)).FormulaR1C1 = "=SUM(R[-" & p_int_NroFil - 6 & "]C:R[-1]C)"
               .Range(.Cells(p_int_NroFil, 33), .Cells(p_int_NroFil, 36)).Font.Bold = True
               r_int_ConAux = p_int_NroFil
               p_int_NroFil = p_int_NroFil + 3
            End If
            
            .Cells(p_int_NroFil, 2) = g_rst_Princi!DOCUMENTO
            .Cells(p_int_NroFil, 3) = g_rst_Princi!RAZON_SOCIAL
            .Cells(p_int_NroFil, 4) = g_rst_Princi!TIPO_EMPRESA
            .Cells(p_int_NroFil, 5) = IIf(g_rst_Princi!MODALIDAD = "001", "LC", "CP")
            .Cells(p_int_NroFil, 6) = "'" & gf_Formato_NumRef(Trim(g_rst_Princi!NUMREF), 1)
            .Cells(p_int_NroFil, 7) = "'" & Format(gf_FormatoFecha(CStr(g_rst_Princi!MAECFI_EMIFIA)), "dd/mm/yyyy")
            .Cells(p_int_NroFil, 8) = "'" & Format(gf_FormatoFecha(CStr(g_rst_Princi!MAECFI_VTOFIA)), "dd/mm/yyyy")
            .Cells(p_int_NroFil, 9) = g_rst_Princi!Moneda
            .Cells(p_int_NroFil, 10) = g_rst_Princi!TASA_ANUAL / 100
            .Cells(p_int_NroFil, 11) = g_rst_Princi!PLAZO
            .Cells(p_int_NroFil, 12) = g_rst_Princi!MAECFI_IMPFIA
            .Cells(p_int_NroFil, 13) = g_rst_Princi!MAECFI_GARFIA
            .Cells(p_int_NroFil, 14) = g_rst_Princi!IMPORTE_GARANTIA
            .Cells(p_int_NroFil, 15) = g_rst_Princi!PAGADO_GARANTIA
            .Cells(p_int_NroFil, 16) = g_rst_Princi!SALDO_GARANTIA
            .Cells(p_int_NroFil, 17) = IIf(g_rst_Princi!MONTO_GARANTIA_HIPOTECARIA > 0, "SI", "")
            .Cells(p_int_NroFil, 18) = g_rst_Princi!IMPORTE_COMISION
            .Cells(p_int_NroFil, 19) = g_rst_Princi!PAGADO_COMISION
            .Cells(p_int_NroFil, 20) = g_rst_Princi!SALDO_COMISION
            .Cells(p_int_NroFil, 21) = g_rst_Princi!IMPORTE_FONDOS
            .Cells(p_int_NroFil, 22) = g_rst_Princi!RECIBIDO_FONDOS
            .Cells(p_int_NroFil, 23) = g_rst_Princi!SALDO_FONDOS
            .Cells(p_int_NroFil, 24) = g_rst_Princi!IMPORTE_DESEMBOLSADO
            .Cells(p_int_NroFil, 25) = g_rst_Princi!PAGADO_DESEMBOLSO
            .Cells(p_int_NroFil, 26) = g_rst_Princi!SALDO_DESEMBOLSO
            
            .Cells(p_int_NroFil, 27) = g_rst_Princi!LINEA_CREDITO
            .Cells(p_int_NroFil, 28) = g_rst_Princi!SALDO_CAPITAL
            .Cells(p_int_NroFil, 29) = g_rst_Princi!INTERES_ACUMULADO
            
            .Cells(p_int_NroFil, 30) = g_rst_Princi!SITUACION
            .Cells(p_int_NroFil, 31) = g_rst_Princi!FACTOR
            .Cells(p_int_NroFil, 32) = g_rst_Princi!EXPOSICION
            .Cells(p_int_NroFil, 33) = g_rst_Princi!TASA
            .Cells(p_int_NroFil, 34) = g_rst_Princi!PROV_ACTUAL
            .Cells(p_int_NroFil, 35) = g_rst_Princi!PROV_ANTERIOR
            .Cells(p_int_NroFil, 36) = .Cells(p_int_NroFil, 34) - .Cells(p_int_NroFil, 35)
            .Cells(p_int_NroFil, 37) = g_rst_Princi!USUARIO
            
         p_int_NroFil = p_int_NroFil + 1
         g_rst_Princi.MoveNext
      Loop
      End If
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 37)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 37)).Font.Bold = True
      .Range(.Cells(p_int_NroFil, 5), .Cells(p_int_NroFil, 37)).HorizontalAlignment = xlHAlignRight
            
      'SUMATORIA TOTAL
      .Cells(p_int_NroFil, 3) = "TOTAL"
      .Cells(p_int_NroFil, 12).FormulaR1C1 = "=SUM(R[-" & p_int_NroFil - 5 & "]C:R[-1]C)"
      
      .Range(.Cells(p_int_NroFil, 27), .Cells(p_int_NroFil, 29)).FormulaR1C1 = "=SUM(R[-" & p_int_NroFil - r_int_ConAux - 3 & "]C:R[-1]C)"
      .Range(.Cells(p_int_NroFil, 27), .Cells(p_int_NroFil, 29)).Font.Bold = True
      
      .Range(.Cells(p_int_NroFil, 33), .Cells(p_int_NroFil, 37)).FormulaR1C1 = "=SUM(R[-" & p_int_NroFil - r_int_ConAux - 3 & "]C:R[-1]C)"
      .Range(.Cells(p_int_NroFil, 33), .Cells(p_int_NroFil, 37)).Font.Bold = True
      
      'RESUMEN POR TIPO EMPRESA
      .Range(.Cells(p_int_NroFil + 2, 2), .Cells(p_int_NroFil + 2, 4)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil + 2, 2), .Cells(p_int_NroFil + 2, 4)).Font.Bold = True
      .Cells(p_int_NroFil + 2, 2) = "CANTIDAD"
      .Cells(p_int_NroFil + 2, 3) = "DESCRIPCION"
      .Cells(p_int_NroFil + 2, 4) = "GARANTIZADO"
      
      .Range(.Cells(p_int_NroFil + 3, 3), .Cells(p_int_NroFil + 6, 3)).Font.Bold = True
      .Range(.Cells(p_int_NroFil + 3, 4), .Cells(p_int_NroFil + 6, 4)).NumberFormat = "###,###,###,##0.00"
      .Range(.Cells(p_int_NroFil + 3, 4), .Cells(p_int_NroFil + 6, 4)).HorizontalAlignment = xlHAlignRight
      
      .Cells(p_int_NroFil + 3, 2).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 4) & "]C[2]:R[-4]C[2],""PEQUEÑA"")"
      .Cells(p_int_NroFil + 3, 3) = "RESUMEN PEQUEÑA EMPRESA"
      .Cells(p_int_NroFil + 3, 4).FormulaR1C1 = "=SUMIF(R[-" & (p_int_NroFil - 4) & "]C:R[-4]C[8],""PEQUEÑA"",R[-" & (p_int_NroFil - 4) & "]C[8]:R[-4]C[8])"
      
      .Cells(p_int_NroFil + 4, 2).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 3) & "]C[2]:R[-5]C[2],""MEDIANA"")"
      .Cells(p_int_NroFil + 4, 3) = "RESUMEN MEDIANA EMPRESA"
      .Cells(p_int_NroFil + 4, 4).FormulaR1C1 = "=SUMIF(R[-" & (p_int_NroFil - 3) & "]C:R[-5]C[8],""MEDIANA"",R[-" & (p_int_NroFil - 3) & "]C[8]:R[-5]C[8])"
      
      .Cells(p_int_NroFil + 5, 2).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 2) & "]C[2]:R[-6]C[2],""GRANDE"")"
      .Cells(p_int_NroFil + 5, 3) = "RESUMEN GRANDE EMPRESA"
      .Cells(p_int_NroFil + 5, 4).FormulaR1C1 = "=+SUMIF(R[-" & (p_int_NroFil - 2) & "]C:R[-6]C[8],""GRANDE"",R[-" & (p_int_NroFil - 2) & "]C[8]:R[-6]C[8])"
      
      .Cells(p_int_NroFil + 6, 2).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 1) & "]C[2]:R[-7]C[2],""MICRO"")"
      .Cells(p_int_NroFil + 6, 3) = "RESUMEN MICRO EMPRESA"
      .Cells(p_int_NroFil + 6, 4).FormulaR1C1 = "=+SUMIF(R[-" & (p_int_NroFil - 1) & "]C:R[-7]C[8],""MICRO"",R[-" & (p_int_NroFil - 1) & "]C[8]:R[-7]C[8])"
      
      .Cells(p_int_NroFil + 7, 2).FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
      .Cells(p_int_NroFil + 7, 3) = "TOTAL"
      .Cells(p_int_NroFil + 7, 4).FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
      
      .Range(.Cells(p_int_NroFil + 7, 4), .Cells(p_int_NroFil + 7, 4)).HorizontalAlignment = xlHAlignRight
      .Range(.Cells(p_int_NroFil + 7, 2), .Cells(p_int_NroFil + 7, 4)).Font.Bold = True
      
      With .Range(.Cells(1, 2), .Cells(1, 37))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
      End With
      
      With .Range(.Cells(5, 2), .Cells(p_int_NroFil - 1, 37))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:            .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:     .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
      With .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 37))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:            .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:     .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
             
      With .Range(.Cells(5, 14), .Cells(p_int_NroFil, 16))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:            .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:     .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
      With .Range(.Cells(5, 18), .Cells(p_int_NroFil, 20))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:            .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:     .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
      With .Range(.Cells(5, 21), .Cells(p_int_NroFil, 23))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:            .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:     .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
      With .Range(.Cells(5, 24), .Cells(p_int_NroFil, 26))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:            .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:     .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
      With .Range(.Cells(p_int_NroFil + 2, 2), .Cells(p_int_NroFil + 6, 4))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeLeft).Weight = xlThin
          .Borders(xlEdgeTop).LineStyle = xlContinuous:            .Borders(xlEdgeTop).Weight = xlThin
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlThin
          .Borders(xlEdgeRight).LineStyle = xlContinuous:          .Borders(xlEdgeRight).Weight = xlThin
      End With
      
      With .Range(.Cells(p_int_NroFil + 7, 2), .Cells(p_int_NroFil + 7, 4))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeTop).LineStyle = xlContinuous:
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeRight).LineStyle = xlContinuous:
      End With
      
      With .Range(.Cells(5, 2), .Cells(4, 37))
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlMedium
      End With
   End With
     
   p_obj_Excel.Sheets(10).Range("E7").Select
   p_obj_Excel.ActiveWindow.FreezePanes = True
   
   If r_bol_Estado = True Then
      p_obj_Excel.Sheets(10).Select
      p_obj_Excel.ActiveWindow.SelectedSheets.Visible = False
      
      p_obj_Excel.Sheets(8).Select
      p_obj_Excel.ActiveWindow.SelectedSheets.Visible = False
   End If
   
End Sub
Private Sub fs_GenExc_CDir_CP(ByVal p_obj_Excel As Excel.Application, ByVal p_int_NroFil As Integer)
Dim r_int_ConAux        As Integer
Dim r_bol_Estado        As Boolean
'Cartas:
 '  p_int_NroFil = 5
   p_obj_Excel.Sheets(11).Select
   
   With p_obj_Excel.Sheets(11)
   
      .Cells(1, 2) = "REPORTE GENERAL DE CARTAS FIANZA"
      .Range(.Cells(1, 2), .Cells(1, 21)).Merge
      
      If Chk_FecAct.Value = 1 Then
         .Cells(3, 2) = " AL " & Format(date, "DD/MM/YYYY")
      Else
         .Cells(3, 2) = " AL " & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), ipp_PerAno.Text), "00") & " DE " & cmb_PerMes.Text & " DE " & Format(ipp_PerAno.Text, "0000")
      End If
      .Range(.Cells(3, 2), .Cells(3, 21)).Merge
      
      .Range(.Cells(1, 2), .Cells(1, 21)).Font.Bold = True
      .Range(.Cells(1, 2), .Cells(3, 21)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 2), .Cells(1, 21)).Font.Size = 14
      
      .Cells(p_int_NroFil, 2) = "TIPO - NRO. DOCUMENTO"
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 2)).Merge
      .Cells(p_int_NroFil, 3) = "RAZON SOCIAL"
      .Range(.Cells(p_int_NroFil, 3), .Cells(p_int_NroFil + 1, 3)).Merge
      .Cells(p_int_NroFil, 4) = "TIPO EMPRESA"
      .Range(.Cells(p_int_NroFil, 4), .Cells(p_int_NroFil + 1, 4)).Merge
      .Cells(p_int_NroFil, 5) = "MODALIDAD"
      .Range(.Cells(p_int_NroFil, 5), .Cells(p_int_NroFil + 1, 5)).Merge
      
      .Cells(p_int_NroFil, 6) = "NÚMERO"
      .Range(.Cells(p_int_NroFil, 6), .Cells(p_int_NroFil + 1, 6)).Merge
      .Cells(p_int_NroFil, 7) = "FECHA EMISIÓN"
      .Range(.Cells(p_int_NroFil, 7), .Cells(p_int_NroFil + 1, 7)).Merge
      .Cells(p_int_NroFil, 8) = "FECHA     VCTO."
      .Range(.Cells(p_int_NroFil, 8), .Cells(p_int_NroFil + 1, 8)).Merge
      
      .Cells(p_int_NroFil, 9) = "MONEDA"
      .Range(.Cells(p_int_NroFil, 9), .Cells(p_int_NroFil + 1, 9)).Merge
      .Cells(p_int_NroFil, 10) = "TASA ANUAL"
      .Range(.Cells(p_int_NroFil, 10), .Cells(p_int_NroFil + 1, 10)).Merge
      .Cells(p_int_NroFil, 11) = "PLAZO"
      .Range(.Cells(p_int_NroFil, 11), .Cells(p_int_NroFil + 1, 11)).Merge
      
      .Cells(p_int_NroFil, 12) = "MONTO PRESTAMO" ' VALOR CARTA FIANZA
      .Range(.Cells(p_int_NroFil, 12), .Cells(p_int_NroFil + 1, 12)).Merge
      .Cells(p_int_NroFil, 13) = "GARANTIZADO"
      .Range(.Cells(p_int_NroFil, 13), .Cells(p_int_NroFil + 1, 13)).Merge
           
      .Cells(p_int_NroFil, 14) = "GARANTIA"
      .Range(.Cells(p_int_NroFil, 14), .Cells(p_int_NroFil, 16)).Merge
      .Cells(p_int_NroFil + 1, 14) = "IMPORTE"
      .Cells(p_int_NroFil + 1, 15) = "PAGADO"
      .Cells(p_int_NroFil + 1, 16) = "SALDO"
      
      .Cells(p_int_NroFil, 17) = "GARANTIA HIPOTECARIA"
      .Range(.Cells(p_int_NroFil, 17), .Cells(p_int_NroFil + 1, 17)).Merge
      
      .Cells(p_int_NroFil, 18) = "COMISIONES"
      .Range(.Cells(p_int_NroFil, 18), .Cells(p_int_NroFil, 20)).Merge
      .Cells(p_int_NroFil + 1, 18) = "IMPORTE"
      .Cells(p_int_NroFil + 1, 19) = "PAGADO"
      .Cells(p_int_NroFil + 1, 20) = "SALDO"
      .Cells(p_int_NroFil, 21) = "FONDOS RECIBIDOS" ' - FMV
      .Range(.Cells(p_int_NroFil, 21), .Cells(p_int_NroFil, 23)).Merge
      .Cells(p_int_NroFil + 1, 21) = "IMPORTE"
      .Cells(p_int_NroFil + 1, 22) = "RECIBIDO"
      .Cells(p_int_NroFil + 1, 23) = "SALDO"
      .Cells(p_int_NroFil, 24) = "DESEMBOLSOS - ET"
      .Range(.Cells(p_int_NroFil, 24), .Cells(p_int_NroFil, 26)).Merge
      .Cells(p_int_NroFil + 1, 24) = "IMPORTE"
      .Cells(p_int_NroFil + 1, 25) = "PAGADO"
      .Cells(p_int_NroFil + 1, 26) = "SALDO"
      
      .Cells(p_int_NroFil, 27) = "SALDO LINEA CREDITO"
      .Range(.Cells(p_int_NroFil, 27), .Cells(p_int_NroFil + 1, 27)).Merge
      .Cells(p_int_NroFil, 28) = "SALDO CAPITAL"
      .Range(.Cells(p_int_NroFil, 28), .Cells(p_int_NroFil + 1, 28)).Merge
      .Cells(p_int_NroFil, 29) = "INTERES DEVENGADO"
      .Range(.Cells(p_int_NroFil, 29), .Cells(p_int_NroFil + 1, 29)).Merge
      
      .Cells(p_int_NroFil, 30) = "ESTADO"
      .Range(.Cells(p_int_NroFil, 30), .Cells(p_int_NroFil + 1, 30)).Merge
      .Cells(p_int_NroFil, 31) = "Factor"
      .Range(.Cells(p_int_NroFil, 31), .Cells(p_int_NroFil + 1, 31)).Merge
      .Cells(p_int_NroFil, 32) = "Exposición"
      .Range(.Cells(p_int_NroFil, 32), .Cells(p_int_NroFil + 1, 32)).Merge
      .Cells(p_int_NroFil, 33) = "Tasa"
      .Range(.Cells(p_int_NroFil, 33), .Cells(p_int_NroFil + 1, 33)).Merge
      .Cells(p_int_NroFil, 34) = "Prov. Gen. Actual"
      .Range(.Cells(p_int_NroFil, 34), .Cells(p_int_NroFil + 1, 34)).Merge
      .Cells(p_int_NroFil, 35) = "Prov. Gen. Anterior"
      .Range(.Cells(p_int_NroFil, 35), .Cells(p_int_NroFil + 1, 35)).Merge
      .Cells(p_int_NroFil, 36) = "Diferencia"
      .Range(.Cells(p_int_NroFil, 36), .Cells(p_int_NroFil + 1, 36)).Merge
      .Cells(p_int_NroFil, 37) = "USUARIO"
      .Range(.Cells(p_int_NroFil, 37), .Cells(p_int_NroFil + 1, 37)).Merge
      
      
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 37)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 37)).Font.Bold = True
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil + 1, 37)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 16.5
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 60
      .Columns("D").ColumnWidth = 17
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 13
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      
      .Columns("F").ColumnWidth = 13
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      
      .Columns("G").ColumnWidth = 13
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      
      .Columns("H").ColumnWidth = 13
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      
      .Columns("I").ColumnWidth = 21
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      
      .Columns("J").ColumnWidth = 13
      .Columns("J").NumberFormat = "0.00%"
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      
      .Columns("K").ColumnWidth = 13
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      
      .Columns("L").ColumnWidth = 13
      .Columns("L").NumberFormat = "###,###,###,##0.00"
      .Columns("L").HorizontalAlignment = xlHAlignRight
      
      .Columns("M").ColumnWidth = 13.5
      .Columns("M").NumberFormat = "###,###,###,##0.00"
      .Columns("M").HorizontalAlignment = xlHAlignRight
      
      .Columns("N").ColumnWidth = 13.5
      .Columns("N").NumberFormat = "###,###,###,##0.00"
      .Columns("N").HorizontalAlignment = xlHAlignRight
      
      .Columns("O").ColumnWidth = 13.5
      .Columns("O").NumberFormat = "###,###,###,##0.00"
      .Columns("O").HorizontalAlignment = xlHAlignRight
      
      .Columns("P").ColumnWidth = 13.5
      .Columns("P").NumberFormat = "###,###,###,##0.00"
      .Columns("P").HorizontalAlignment = xlHAlignRight
      
      .Columns("Q").ColumnWidth = 13.5
      .Columns("Q").HorizontalAlignment = xlHAlignCenter
      
      .Columns("R").ColumnWidth = 13.5
      .Columns("R").NumberFormat = "###,###,###,##0.00"
      .Columns("R").HorizontalAlignment = xlHAlignRight
      
      .Columns("S").ColumnWidth = 13.5
      .Columns("S").NumberFormat = "###,###,###,##0.00"
      .Columns("S").HorizontalAlignment = xlHAlignRight
      
      .Columns("T").ColumnWidth = 13.5
      .Columns("T").NumberFormat = "###,###,###,##0.00"
      .Columns("T").HorizontalAlignment = xlHAlignRight
      
      .Columns("U").ColumnWidth = 13.5
      .Columns("U").NumberFormat = "###,###,###,##0.00"
      .Columns("U").HorizontalAlignment = xlHAlignRight
      
      .Columns("V").ColumnWidth = 13.5
      .Columns("V").NumberFormat = "###,###,###,##0.00"
      .Columns("V").HorizontalAlignment = xlHAlignRight
      
      .Columns("W").ColumnWidth = 13.5
      .Columns("W").NumberFormat = "###,###,###,##0.00"
      .Columns("W").HorizontalAlignment = xlHAlignRight
      
      .Columns("X").ColumnWidth = 13.5
      .Columns("X").NumberFormat = "###,###,###,##0.00"
      .Columns("X").HorizontalAlignment = xlHAlignRight
      
      .Columns("Y").ColumnWidth = 13.5
      .Columns("Y").NumberFormat = "###,###,###,##0.00"
      .Columns("Y").HorizontalAlignment = xlHAlignRight
      
      .Columns("Z").ColumnWidth = 13.5
      .Columns("Z").NumberFormat = "###,###,###,##0.00"
      .Columns("Z").HorizontalAlignment = xlHAlignRight
      
      .Columns("AA").ColumnWidth = 13.5
      .Columns("AA").NumberFormat = "###,###,###,##0.00"
      .Columns("AA").HorizontalAlignment = xlHAlignRight
      
      .Columns("AB").ColumnWidth = 13.5
      .Columns("AB").NumberFormat = "###,###,###,##0.00"
      .Columns("AB").HorizontalAlignment = xlHAlignRight
      
      .Columns("AC").ColumnWidth = 13.5
      .Columns("AC").NumberFormat = "###,###,###,##0.00"
      .Columns("AC").HorizontalAlignment = xlHAlignRight
      
      .Columns("AD").ColumnWidth = 13.5
      .Columns("AD").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AE").ColumnWidth = 6
      .Columns("AE").NumberFormat = "0.00%"
      .Columns("AE").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AF").ColumnWidth = 13.5
      .Columns("AF").NumberFormat = "###,###,###,##0.00"
      .Columns("AF").HorizontalAlignment = xlHAlignRight
      
      .Columns("AG").ColumnWidth = 6
      .Columns("AG").NumberFormat = "0.00%"
      .Columns("AG").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AH").ColumnWidth = 13.5
      .Columns("AH").NumberFormat = "###,###,###,##0.00"
      .Columns("AH").HorizontalAlignment = xlHAlignRight
      
      .Columns("AI").ColumnWidth = 13.5
      .Columns("AI").NumberFormat = "###,###,###,##0.00"
      .Columns("AI").HorizontalAlignment = xlHAlignRight
      
      .Columns("AJ").ColumnWidth = 13.5
      .Columns("AJ").NumberFormat = "###,###,###,##0.00"
      .Columns("AJ").HorizontalAlignment = xlHAlignRight
      
      .Columns("AK").ColumnWidth = 13.5
      .Columns("AK").HorizontalAlignment = xlHAlignCenter
      
      
      With .Range(.Cells(5, 2), .Cells(6, 37))
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
          .WrapText = True
      End With
      
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(2, 1), .Cells(99, 99)).Font.Size = 11
            
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "   SELECT RPT_PERMES   MES               , RPT_PERANO   ANNO               , RPT_DESCRI   DOCUMENTO             , RPT_VALCAD01 RAZON_SOCIAL   , "
      g_str_Parame = g_str_Parame & "          RPT_VALCAD02 TIPO_EMPRESA      , RPT_VALCAD03 MONEDA             , RPT_VALNUM25 TASA_ANUAL            , RPT_VALNUM01 PLAZO          , "
      g_str_Parame = g_str_Parame & "          RPT_VALCAD05 NUMREF            , RPT_VALCAD06 MAECFI_EMIFIA      , RPT_VALCAD07 MAECFI_VTOFIA         , RPT_VALNUM02 MAECFI_IMPFIA  , "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM03 MAECFI_GARFIA     , RPT_VALNUM04 IMPORTE_GARANTIA   , RPT_VALNUM05 PAGADO_GARANTIA       , RPT_VALNUM06 SALDO_GARANTIA , "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM07 IMPORTE_COMISION  , RPT_VALNUM08 PAGADO_COMISION    , RPT_VALNUM09 SALDO_COMISION        , RPT_VALNUM10 IMPORTE_FONDOS , "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM11 RECIBIDO_FONDOS   , RPT_VALNUM12 SALDO_FONDOS       , RPT_VALNUM13 IMPORTE_DESEMBOLSADO  ,   "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM14 PAGADO_DESEMBOLSO , RPT_VALNUM15 DEVOLUCION_GARANTIA, RPT_VALNUM16 SALDO_DESEMBOLSO      , RPT_VALCAD08 SITUACION    , "
      g_str_Parame = g_str_Parame & "          RPT_VALCAD09 FACTOR            , RPT_VALNUM17 EXPOSICION         , RPT_VALCAD10 TASA                  , RPT_VALNUM18 PROV_ACTUAL  , "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM19 PROV_ANTERIOR     , RPT_VALCAD11 TIPO_GARANTIA      , RPT_VALNUM23 MONTO_GARANTIA_LIQUIDA, RPT_VALNUM24 MONTO_GARANTIA_HIPOTECARIA , "
      g_str_Parame = g_str_Parame & "          RPT_VALCAD14 USUARIO           , RPT_VALCAD13 MODALIDAD          , RPT_VALCAD13 MODALIDAD              , RPT_VALNUM26 INTERES_ACUMULADO         , "
      g_str_Parame = g_str_Parame & "          RPT_VALNUM27 LINEA_CREDITO     , RPT_VALNUM28 SALDO_CAPITAL "
      g_str_Parame = g_str_Parame & "     FROM RPT_TABLA_TEMP "
      g_str_Parame = g_str_Parame & "    WHERE RPT_PERMES = " & Month(Format(gf_FormatoFecha(CStr(l_str_FecPer)), "dd/mm/yyyy")) & "" 'Month(Now)
      g_str_Parame = g_str_Parame & "      AND RPT_PERANO = " & Year(Format(gf_FormatoFecha(CStr(l_str_FecPer)), "dd/mm/yyyy")) & " " 'Year(Now)
      g_str_Parame = g_str_Parame & "      AND RPT_TERCRE = " & "'" & modgen_g_str_NombPC & "'"
      g_str_Parame = g_str_Parame & "      AND RPT_USUCRE = " & "'" & modgen_g_str_CodUsu & "'"
      g_str_Parame = g_str_Parame & "      AND RPT_NOMBRE = " & "'" & Me.cmb_TipRep.Text & "2'"
      g_str_Parame = g_str_Parame & "      AND RPT_VALCAD11 = '008'"
      g_str_Parame = g_str_Parame & "      AND RPT_VALCAD12 = '008'"
      g_str_Parame = g_str_Parame & "      AND RPT_VALCAD13 = '002'"
      g_str_Parame = g_str_Parame & "    ORDER BY RPT_VALCAD02, RPT_VALCAD01, RPT_VALCAD06 "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
          
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
'         g_rst_Princi.Close
'         Set g_rst_Princi = Nothing
         r_bol_Estado = True
'         MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
'         Exit Sub
      End If
      
      p_int_NroFil = p_int_NroFil + 2
      
      If Not g_rst_Princi.EOF Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
            If .Cells(p_int_NroFil - 1, 4) <> g_rst_Princi!TIPO_EMPRESA And .Cells(p_int_NroFil - 1, 4) <> "" Then
               .Range(.Cells(p_int_NroFil, 27), .Cells(p_int_NroFil, 29)).FormulaR1C1 = "=SUM(R[-" & p_int_NroFil - 6 & "]C:R[-1]C)"
               .Range(.Cells(p_int_NroFil, 27), .Cells(p_int_NroFil, 29)).Font.Bold = True
               
               .Range(.Cells(p_int_NroFil, 33), .Cells(p_int_NroFil, 36)).FormulaR1C1 = "=SUM(R[-" & p_int_NroFil - 6 & "]C:R[-1]C)"
               .Range(.Cells(p_int_NroFil, 33), .Cells(p_int_NroFil, 36)).Font.Bold = True
               r_int_ConAux = p_int_NroFil
               p_int_NroFil = p_int_NroFil + 3
            End If
            
            .Cells(p_int_NroFil, 2) = g_rst_Princi!DOCUMENTO
            .Cells(p_int_NroFil, 3) = g_rst_Princi!RAZON_SOCIAL
            .Cells(p_int_NroFil, 4) = g_rst_Princi!TIPO_EMPRESA
            .Cells(p_int_NroFil, 5) = IIf(g_rst_Princi!MODALIDAD = "001", "LC", "CP")
            .Cells(p_int_NroFil, 6) = "'" & gf_Formato_NumRef(Trim(g_rst_Princi!NUMREF), 1)
            .Cells(p_int_NroFil, 7) = "'" & Format(gf_FormatoFecha(CStr(g_rst_Princi!MAECFI_EMIFIA)), "dd/mm/yyyy")
            .Cells(p_int_NroFil, 8) = "'" & Format(gf_FormatoFecha(CStr(g_rst_Princi!MAECFI_VTOFIA)), "dd/mm/yyyy")
            .Cells(p_int_NroFil, 9) = g_rst_Princi!Moneda
            .Cells(p_int_NroFil, 10) = g_rst_Princi!TASA_ANUAL / 100
            .Cells(p_int_NroFil, 11) = g_rst_Princi!PLAZO
            .Cells(p_int_NroFil, 12) = g_rst_Princi!MAECFI_IMPFIA
            .Cells(p_int_NroFil, 13) = g_rst_Princi!MAECFI_GARFIA
            .Cells(p_int_NroFil, 14) = g_rst_Princi!IMPORTE_GARANTIA
            .Cells(p_int_NroFil, 15) = g_rst_Princi!PAGADO_GARANTIA
            .Cells(p_int_NroFil, 16) = g_rst_Princi!SALDO_GARANTIA
            .Cells(p_int_NroFil, 17) = IIf(g_rst_Princi!MONTO_GARANTIA_HIPOTECARIA > 0, "SI", "")
            .Cells(p_int_NroFil, 18) = g_rst_Princi!IMPORTE_COMISION
            .Cells(p_int_NroFil, 19) = g_rst_Princi!PAGADO_COMISION
            .Cells(p_int_NroFil, 20) = g_rst_Princi!SALDO_COMISION
            .Cells(p_int_NroFil, 21) = g_rst_Princi!IMPORTE_FONDOS
            .Cells(p_int_NroFil, 22) = g_rst_Princi!RECIBIDO_FONDOS
            .Cells(p_int_NroFil, 23) = g_rst_Princi!SALDO_FONDOS
            .Cells(p_int_NroFil, 24) = g_rst_Princi!IMPORTE_DESEMBOLSADO
            .Cells(p_int_NroFil, 25) = g_rst_Princi!PAGADO_DESEMBOLSO
            .Cells(p_int_NroFil, 26) = g_rst_Princi!SALDO_DESEMBOLSO
            .Cells(p_int_NroFil, 27) = g_rst_Princi!LINEA_CREDITO
            .Cells(p_int_NroFil, 28) = g_rst_Princi!SALDO_CAPITAL
            .Cells(p_int_NroFil, 29) = g_rst_Princi!INTERES_ACUMULADO
            .Cells(p_int_NroFil, 30) = g_rst_Princi!SITUACION
            .Cells(p_int_NroFil, 31) = g_rst_Princi!FACTOR
            .Cells(p_int_NroFil, 32) = g_rst_Princi!EXPOSICION
            .Cells(p_int_NroFil, 33) = g_rst_Princi!TASA
            .Cells(p_int_NroFil, 34) = g_rst_Princi!PROV_ACTUAL
            .Cells(p_int_NroFil, 35) = g_rst_Princi!PROV_ANTERIOR
            .Cells(p_int_NroFil, 36) = .Cells(p_int_NroFil, 34) - .Cells(p_int_NroFil, 35)
            .Cells(p_int_NroFil, 37) = g_rst_Princi!USUARIO
            
         p_int_NroFil = p_int_NroFil + 1
         g_rst_Princi.MoveNext
      Loop
      End If
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 37)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 37)).Font.Bold = True
      .Range(.Cells(p_int_NroFil, 5), .Cells(p_int_NroFil, 37)).HorizontalAlignment = xlHAlignRight
            
      'SUMATORIA TOTAL
      .Cells(p_int_NroFil, 3) = "TOTAL"
      .Cells(p_int_NroFil, 12).FormulaR1C1 = "=SUM(R[-" & p_int_NroFil - 5 & "]C:R[-1]C)"
      
      .Range(.Cells(p_int_NroFil, 27), .Cells(p_int_NroFil, 29)).FormulaR1C1 = "=SUM(R[-" & p_int_NroFil - r_int_ConAux - 3 & "]C:R[-1]C)"
      .Range(.Cells(p_int_NroFil, 27), .Cells(p_int_NroFil, 29)).Font.Bold = True
      
      .Range(.Cells(p_int_NroFil, 33), .Cells(p_int_NroFil, 37)).FormulaR1C1 = "=SUM(R[-" & p_int_NroFil - r_int_ConAux - 3 & "]C:R[-1]C)"
      .Range(.Cells(p_int_NroFil, 33), .Cells(p_int_NroFil, 37)).Font.Bold = True
      
      'RESUMEN POR TIPO EMPRESA
      .Range(.Cells(p_int_NroFil + 2, 2), .Cells(p_int_NroFil + 2, 4)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(p_int_NroFil + 2, 2), .Cells(p_int_NroFil + 2, 4)).Font.Bold = True
      .Cells(p_int_NroFil + 2, 2) = "CANTIDAD"
      .Cells(p_int_NroFil + 2, 3) = "DESCRIPCION"
      .Cells(p_int_NroFil + 2, 4) = "GARANTIZADO"
      
      .Range(.Cells(p_int_NroFil + 3, 3), .Cells(p_int_NroFil + 6, 3)).Font.Bold = True
      .Range(.Cells(p_int_NroFil + 3, 4), .Cells(p_int_NroFil + 6, 4)).NumberFormat = "###,###,###,##0.00"
      .Range(.Cells(p_int_NroFil + 3, 4), .Cells(p_int_NroFil + 6, 4)).HorizontalAlignment = xlHAlignRight
      
      .Cells(p_int_NroFil + 3, 2).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 4) & "]C[2]:R[-4]C[2],""PEQUEÑA"")"
      .Cells(p_int_NroFil + 3, 3) = "RESUMEN PEQUEÑA EMPRESA"
      .Cells(p_int_NroFil + 3, 4).FormulaR1C1 = "=SUMIF(R[-" & (p_int_NroFil - 4) & "]C:R[-4]C[8],""PEQUEÑA"",R[-" & (p_int_NroFil - 4) & "]C[8]:R[-4]C[8])"
      
      .Cells(p_int_NroFil + 4, 2).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 3) & "]C[2]:R[-5]C[2],""MEDIANA"")"
      .Cells(p_int_NroFil + 4, 3) = "RESUMEN MEDIANA EMPRESA"
      .Cells(p_int_NroFil + 4, 4).FormulaR1C1 = "=SUMIF(R[-" & (p_int_NroFil - 3) & "]C:R[-5]C[8],""MEDIANA"",R[-" & (p_int_NroFil - 3) & "]C[8]:R[-5]C[8])"
      
      .Cells(p_int_NroFil + 5, 2).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 2) & "]C[2]:R[-6]C[2],""GRANDE"")"
      .Cells(p_int_NroFil + 5, 3) = "RESUMEN GRANDE EMPRESA"
      .Cells(p_int_NroFil + 5, 4).FormulaR1C1 = "=+SUMIF(R[-" & (p_int_NroFil - 2) & "]C:R[-6]C[8],""GRANDE"",R[-" & (p_int_NroFil - 2) & "]C[8]:R[-6]C[8])"
      
      .Cells(p_int_NroFil + 6, 2).FormulaR1C1 = "=COUNTIF(R[-" & (p_int_NroFil - 1) & "]C[2]:R[-7]C[2],""MICRO"")"
      .Cells(p_int_NroFil + 6, 3) = "RESUMEN MICRO EMPRESA"
      .Cells(p_int_NroFil + 6, 4).FormulaR1C1 = "=+SUMIF(R[-" & (p_int_NroFil - 1) & "]C:R[-7]C[8],""MICRO"",R[-" & (p_int_NroFil - 1) & "]C[8]:R[-7]C[8])"
      
      .Cells(p_int_NroFil + 7, 2).FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
      .Cells(p_int_NroFil + 7, 3) = "TOTAL"
      .Cells(p_int_NroFil + 7, 4).FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
      
      .Range(.Cells(p_int_NroFil + 7, 4), .Cells(p_int_NroFil + 7, 4)).HorizontalAlignment = xlHAlignRight
      .Range(.Cells(p_int_NroFil + 7, 2), .Cells(p_int_NroFil + 7, 4)).Font.Bold = True
      
      With .Range(.Cells(1, 2), .Cells(1, 37))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:           .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:        .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:         .Borders(xlEdgeRight).Weight = xlMedium
      End With
      
      With .Range(.Cells(5, 2), .Cells(p_int_NroFil - 1, 37))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:            .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:     .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
      With .Range(.Cells(p_int_NroFil, 2), .Cells(p_int_NroFil, 37))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:            .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:     .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
             
      With .Range(.Cells(5, 14), .Cells(p_int_NroFil, 16))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:            .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:     .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
      With .Range(.Cells(5, 18), .Cells(p_int_NroFil, 20))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:            .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:     .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
      With .Range(.Cells(5, 21), .Cells(p_int_NroFil, 23))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:            .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:     .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
      With .Range(.Cells(5, 24), .Cells(p_int_NroFil, 26))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous:            .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous:          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous:     .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
      With .Range(.Cells(p_int_NroFil + 2, 2), .Cells(p_int_NroFil + 6, 4))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeLeft).Weight = xlThin
          .Borders(xlEdgeTop).LineStyle = xlContinuous:            .Borders(xlEdgeTop).Weight = xlThin
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlThin
          .Borders(xlEdgeRight).LineStyle = xlContinuous:          .Borders(xlEdgeRight).Weight = xlThin
      End With
      
      With .Range(.Cells(p_int_NroFil + 7, 2), .Cells(p_int_NroFil + 7, 4))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous:           .Borders(xlEdgeTop).LineStyle = xlContinuous:
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeRight).LineStyle = xlContinuous:
      End With
      
      With .Range(.Cells(5, 2), .Cells(4, 37))
          .Borders(xlEdgeBottom).LineStyle = xlContinuous:         .Borders(xlEdgeBottom).Weight = xlMedium
      End With
   End With
     
   p_obj_Excel.Sheets(11).Range("E7").Select
   p_obj_Excel.ActiveWindow.FreezePanes = True
   
   If r_bol_Estado = True Then
      p_obj_Excel.Sheets(11).Select
      p_obj_Excel.ActiveWindow.SelectedSheets.Visible = False
      
      p_obj_Excel.Sheets(9).Select
      p_obj_Excel.ActiveWindow.SelectedSheets.Visible = False
'      p_obj_Excel.Sheets(1).Select
   End If
   
   p_obj_Excel.Sheets(1).Select
   p_obj_Excel.Visible = True
   
End Sub
Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   cmb_TipRep.Clear
   cmb_TipRep.AddItem "REPORTE RESUMEN DE ENTIDADES TECNICAS"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = CInt(1)
   cmb_TipRep.AddItem "REPORTE DETALLADO DE CARTAS FIANZA, ADENDAS Y CSO"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = CInt(2)
   cmb_TipRep.AddItem "REPORTE CENTRAL DE RIESGOS Y ALINEAMIENTO"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = CInt(3)
   
   cmb_TipRep.AddItem "REPORTE DE FACTURACIÓN"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = CInt(4)
   
   cmb_TipRep.ListIndex = -1
End Sub

Private Sub fs_Limpia()
   cmb_TipRep.ListIndex = -1
   Chk_FecAct.Value = 0
   ipp_PerAno.Text = Year(date)
End Sub
'
'Private Function fs_Formato_NumRef(ByVal p_Numref As String) As String
'   p_Numref = Format(p_Numref, "0000000000")
'   'fs_Formato_NumRef = Left(p_Numref, 4) & "-" & Mid(p_Numref, 5, 2) & "-" & Right(p_Numref, 4)
'   fs_Formato_NumRef = Mid(p_Numref, 1, 1) & Mid(p_Numref, 2, 2) & "-" & Mid(p_Numref, 4, 2) & "-" & Right(p_Numref, 5)
'End Function

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      Call gs_SetFocus(Chk_FecAct)
  End If
End Sub


