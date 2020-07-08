VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.ocx"
Begin VB.Form frm_Ges_TecPro_11 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13155
   Icon            =   "OpeTra_frm_832.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   13155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   5715
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   13125
      _Version        =   65536
      _ExtentX        =   23151
      _ExtentY        =   10081
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
      Begin Threed.SSPanel SSPanel8 
         Height          =   675
         Left            =   30
         TabIndex        =   1
         Top             =   780
         Width           =   13035
         _Version        =   65536
         _ExtentX        =   22992
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   12420
            Picture         =   "OpeTra_frm_832.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_832.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   4
         Top             =   60
         Width           =   13035
         _Version        =   65536
         _ExtentX        =   22992
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
            TabIndex        =   5
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
            TabIndex        =   6
            Top             =   330
            Width           =   4215
            _Version        =   65536
            _ExtentX        =   7435
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Techo Propio - Histórico de Renovaciones"
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
            Picture         =   "OpeTra_frm_832.frx":0758
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   1155
         Left            =   30
         TabIndex        =   7
         Top             =   1500
         Width           =   13035
         _Version        =   65536
         _ExtentX        =   22992
         _ExtentY        =   2037
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
         Begin Threed.SSPanel pnl_RazSoc 
            Height          =   315
            Left            =   1620
            TabIndex        =   8
            Top             =   450
            Width           =   5625
            _Version        =   65536
            _ExtentX        =   9922
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
         Begin Threed.SSPanel pnl_TipDoc 
            Height          =   315
            Left            =   1620
            TabIndex        =   9
            Top             =   120
            Width           =   5625
            _Version        =   65536
            _ExtentX        =   9922
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
         Begin Threed.SSPanel pnl_NroDoc 
            Height          =   315
            Left            =   9360
            TabIndex        =   10
            Top             =   120
            Width           =   2895
            _Version        =   65536
            _ExtentX        =   5106
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
         Begin Threed.SSPanel pnl_TipEmp 
            Height          =   315
            Left            =   1620
            TabIndex        =   11
            Top             =   780
            Width           =   2895
            _Version        =   65536
            _ExtentX        =   5106
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
         Begin Threed.SSPanel pnl_NumRef 
            Height          =   315
            Left            =   9360
            TabIndex        =   12
            Top             =   450
            Width           =   2895
            _Version        =   65536
            _ExtentX        =   5106
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
         Begin VB.Label lbl_TipEmp 
            Caption         =   "Tipo Empresa:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   810
            Width           =   1335
         End
         Begin VB.Label lbl_RazSoc 
            Caption         =   "Razón Social:"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lbl_NumDoc 
            Caption         =   "Nro. Documento:"
            Height          =   225
            Left            =   7770
            TabIndex        =   15
            Top             =   150
            Width           =   1335
         End
         Begin VB.Label lbl_TipDoc 
            Caption         =   "Tipo Documento:"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   150
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Nro. Referencia:"
            Height          =   255
            Left            =   7770
            TabIndex        =   13
            Top             =   480
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   2955
         Left            =   30
         TabIndex        =   18
         Top             =   2700
         Width           =   13035
         _Version        =   65536
         _ExtentX        =   22992
         _ExtentY        =   5212
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
            Height          =   2415
            Left            =   90
            TabIndex        =   19
            Top             =   450
            Width           =   12930
            _ExtentX        =   22807
            _ExtentY        =   4260
            _Version        =   393216
            Rows            =   21
            Cols            =   9
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin Threed.SSPanel pnl_Tit_TipGar 
            Height          =   285
            Left            =   570
            TabIndex        =   20
            Top             =   150
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Referencia"
            ForeColor       =   16777215
            BackColor       =   16384
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
         Begin Threed.SSPanel pnl_Tit_NumRef 
            Height          =   285
            Left            =   2010
            TabIndex        =   21
            Top             =   150
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Emisión"
            ForeColor       =   16777215
            BackColor       =   16384
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
         Begin Threed.SSPanel pnl_Tit_FecEmi 
            Height          =   285
            Left            =   3600
            TabIndex        =   22
            Top             =   150
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Plazo"
            ForeColor       =   16777215
            BackColor       =   16384
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
         Begin Threed.SSPanel pnl_Tit_TipMon 
            Height          =   285
            Left            =   4800
            TabIndex        =   23
            Top             =   150
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Vencimiento"
            ForeColor       =   16777215
            BackColor       =   16384
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
         Begin Threed.SSPanel pnl_Tit_MtoGar 
            Height          =   285
            Left            =   6390
            TabIndex        =   24
            Top             =   150
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Proceso"
            ForeColor       =   16777215
            BackColor       =   16384
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   285
            Left            =   7950
            TabIndex        =   25
            Top             =   150
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Valor"
            ForeColor       =   16777215
            BackColor       =   16384
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
         Begin Threed.SSPanel SSPanel3 
            Height          =   285
            Left            =   9540
            TabIndex        =   26
            Top             =   150
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Garantizado"
            ForeColor       =   16777215
            BackColor       =   16384
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   285
            Left            =   11100
            TabIndex        =   27
            Top             =   150
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Estado"
            ForeColor       =   16777215
            BackColor       =   16384
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
            Left            =   90
            TabIndex        =   28
            Top             =   150
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo"
            ForeColor       =   16777215
            BackColor       =   16384
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
   End
End
Attribute VB_Name = "frm_Ges_TecPro_11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_ExpExc_Click()
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   Call gs_CentraForm(Me)
   Call fs_Inicia
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 500
   grd_Listad.ColWidth(1) = 1445
   grd_Listad.ColWidth(2) = 1575
   grd_Listad.ColWidth(3) = 1220
   grd_Listad.ColWidth(4) = 1575
   grd_Listad.ColWidth(5) = 1575
   grd_Listad.ColWidth(6) = 1575
   grd_Listad.ColWidth(7) = 1565
   grd_Listad.ColWidth(8) = 1575
  
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_Listad.ColAlignment(8) = flexAlignCenterCenter
   
   Call gs_LimpiaGrid(grd_Listad)
   
   pnl_TipDoc.Caption = moddat_gf_Consulta_ParDes("118", moddat_g_int_TipDoc)
   pnl_NroDoc.Caption = moddat_g_str_NumDoc
   pnl_RazSoc.Caption = moddat_g_str_NomCli
   pnl_TipEmp.Caption = moddat_g_str_Descri
   pnl_NumRef.Caption = gf_Formato_NumRef(moddat_g_str_DesIte, Mid(moddat_g_str_DesIte, 1, 1)) 'moddat_g_str_NumFia
   
End Sub
'Private Function fs_Formato_NumRef(ByVal p_Numref As String) As String
'   p_Numref = Format(p_Numref, "0000000000")
'   'fs_Formato_NumRef = Left(p_Numref, 4) & "-" & Mid(p_Numref, 5, 2) & "-" & Right(p_Numref, 4)
'   fs_Formato_NumRef = Mid(p_Numref, 1, 1) & Mid(p_Numref, 2, 2) & "-" & Mid(p_Numref, 4, 2) & "-" & Right(p_Numref, 5)
'End Function
Private Sub fs_Buscar()

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT MAECFI_NUMREF, MAECFI_EMIFIA, MAECFI_PLZFIA, MAECFI_VTOFIA, MAECFI_MONFIA, MAECFI_CODPRD, MAECFI_CODSUB, "
   g_str_Parame = g_str_Parame & "        MAECFI_CODMOD, MAECFI_IMPFIA, MAECFI_GARFIA, MAECFI_TASFIA, MAECFI_COMFIA, MAECFI_MINFIA, MAECFI_PORRET, "
   g_str_Parame = g_str_Parame & "        MAECFI_REFORI, MAECFI_REFANT, A.SEGFECCRE  , TRIM(PARDES_DESCRI) SITUACION, MAECFI_NUMANT "
   g_str_Parame = g_str_Parame & "   FROM TPR_MAECFI A "
   g_str_Parame = g_str_Parame & "               INNER JOIN MNT_PARDES ON PARDES_CODGRP = '529' AND PARDES_CODITE = MAECFI_SITUAC "
   g_str_Parame = g_str_Parame & "  WHERE MAECFI_REFORI = (SELECT MAECFI_REFORI " 'CASE MAECFI_REFORI WHEN MAECFI_SITUAC = 4 THEN MAECFI_REFORI ELSE MAECFI_REFANT END
   g_str_Parame = g_str_Parame & "                           FROM TPR_MAECFI"
   g_str_Parame = g_str_Parame & "                          WHERE MAECFI_NUMREF = '" & CStr(moddat_g_str_DesIte) & "' " 'moddat_g_str_NumFia
   g_str_Parame = g_str_Parame & "                            AND MAECFI_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " "
   g_str_Parame = g_str_Parame & "                            AND MAECFI_NUMDOC = '" & CStr(moddat_g_str_NumDoc) & "' ) "
   g_str_Parame = g_str_Parame & "    AND MAECFI_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " "
   g_str_Parame = g_str_Parame & "    AND MAECFI_NUMDOC = '" & CStr(moddat_g_str_NumDoc) & "' "
   g_str_Parame = g_str_Parame & "  ORDER BY MAECFI_NUMREN DESC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
       
   grd_Listad.Redraw = False
   Call gs_LimpiaGrid(grd_Listad)
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
      grd_Listad.Redraw = True
     Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
          
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
                 
         grd_Listad.Col = 0
         If g_rst_Princi!MAECFI_CODMOD = "005" Then
            grd_Listad.Text = "AD"
         Else
            grd_Listad.Text = "CF"
         End If
         
         grd_Listad.Col = 1
         If Not IsNull(g_rst_Princi!MAECFI_NUMANT) Then
            grd_Listad.Text = gf_Formato_NumRef(CStr(Trim(g_rst_Princi!MAECFI_NUMANT)), Mid(g_rst_Princi!MAECFI_NUMANT, 1, 1))
         Else
            grd_Listad.Text = gf_Formato_NumRef(CStr(Trim(g_rst_Princi!MAECFI_NUMREF)), Mid(g_rst_Princi!MAECFI_NUMREF, 1, 1))
         End If
         grd_Listad.Col = 2
         grd_Listad.Text = Format(gf_FormatoFecha(CStr(g_rst_Princi!MAECFI_EMIFIA)), "dd/mm/yyyy")
         
         grd_Listad.Col = 3
         grd_Listad.Text = CStr(g_rst_Princi!MAECFI_PLZFIA)
         
         grd_Listad.Col = 4
         grd_Listad.Text = Format(gf_FormatoFecha(CStr(g_rst_Princi!MAECFI_VTOFIA)), "dd/mm/yyyy")
         
         grd_Listad.Col = 5
         grd_Listad.Text = Format(gf_FormatoFecha(CStr(g_rst_Princi!SEGFECCRE)), "dd/mm/yyyy")
         
         grd_Listad.Col = 6
         grd_Listad.Text = Format(CStr(g_rst_Princi!MAECFI_IMPFIA), "###,###,###,##0.00")
         
         grd_Listad.Col = 7
         grd_Listad.Text = Format(CStr(g_rst_Princi!MAECFI_GARFIA), "###,###,###,##0.00")
         
         grd_Listad.Col = 8
         grd_Listad.Text = CStr(g_rst_Princi!SITUACION)
         
         g_rst_Princi.MoveNext
      Loop
   End If
   
   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)
   
End Sub
Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_str_ParAux        As String
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
Dim r_int_NroFil        As Integer
Dim r_int_NoFlLi        As Integer
   
    r_int_NroFil = 8
    
    Set r_obj_Excel = New Excel.Application
    r_obj_Excel.SheetsInNewWorkbook = 1
    r_obj_Excel.Workbooks.Add

    With r_obj_Excel.ActiveSheet
        .Cells(2, 2) = "REPORTE DE HISTÓRICO DE RENOVACIONES"
        .Range(.Cells(2, 2), .Cells(2, 9)).Merge
        .Range(.Cells(2, 2), .Cells(2, 9)).Font.Bold = True
        .Range(.Cells(2, 2), .Cells(2, 9)).HorizontalAlignment = xlHAlignCenter
        .Range(.Cells(2, 2), .Cells(2, 9)).Font.Size = 14

        .Cells(4, 2) = "TIPO DE DOCUMENTO"
        .Cells(4, 3) = Trim(pnl_TipDoc.Caption)
        .Cells(5, 2) = "NRO. DOCUMENTO"
        .Cells(5, 3) = "'" & Trim(pnl_NroDoc.Caption)
        .Cells(6, 2) = "RAZÓN SOCIAL"
        .Cells(6, 3) = Trim(pnl_RazSoc.Caption)
        .Range(.Cells(3, 2), .Cells(6, 2)).Font.Bold = True
        
        .Cells(r_int_NroFil, 2) = "NRO. REFERENCIA"
        .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 2)).Merge
        .Cells(r_int_NroFil, 3) = "FECHA EMISION"
        .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil + 1, 3)).Merge
        .Cells(r_int_NroFil, 4) = "PLAZO"
        .Range(.Cells(r_int_NroFil, 4), .Cells(r_int_NroFil + 1, 4)).Merge
        .Cells(r_int_NroFil, 5) = "FECHA VENCIMIENTO"
        .Range(.Cells(r_int_NroFil, 5), .Cells(r_int_NroFil + 1, 5)).Merge
        .Cells(r_int_NroFil, 6) = "FECHA PROCESO"
        .Range(.Cells(r_int_NroFil, 6), .Cells(r_int_NroFil + 1, 6)).Merge
        .Cells(r_int_NroFil, 7) = "VALOR"
        .Range(.Cells(r_int_NroFil, 7), .Cells(r_int_NroFil + 1, 7)).Merge
        .Cells(r_int_NroFil, 8) = "GARANTIZADO"
        .Range(.Cells(r_int_NroFil, 8), .Cells(r_int_NroFil + 1, 8)).Merge
        .Cells(r_int_NroFil, 9) = "ESTADO"
        .Range(.Cells(r_int_NroFil, 9), .Cells(r_int_NroFil + 1, 9)).Merge
        
        .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 9)).Interior.Color = RGB(146, 208, 80)
        .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 9)).Font.Bold = True
        .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 9)).HorizontalAlignment = xlHAlignCenter
        
        .Columns("A").ColumnWidth = 1
        .Columns("B").ColumnWidth = 20
        .Columns("C").ColumnWidth = 20
        .Columns("C").HorizontalAlignment = xlHAlignCenter
        .Columns("D").ColumnWidth = 15
        .Columns("D").HorizontalAlignment = xlHAlignCenter
        .Columns("E").ColumnWidth = 20
        .Columns("E").HorizontalAlignment = xlHAlignCenter
        .Columns("F").ColumnWidth = 20
        .Columns("F").HorizontalAlignment = xlHAlignCenter
        .Columns("G").ColumnWidth = 13.5
        .Columns("G").NumberFormat = "###,###,###,##0.00"
        .Columns("G").HorizontalAlignment = xlHAlignRight
        .Columns("H").ColumnWidth = 13.5
        .Columns("H").NumberFormat = "###,###,###,##0.00"
        .Columns("H").HorizontalAlignment = xlHAlignRight
        .Columns("I").ColumnWidth = 22
        .Columns("I").HorizontalAlignment = xlHAlignCenter
        
        With .Range(.Cells(8, 2), .Cells(9, 9))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
        
        .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
        .Range(.Cells(3, 1), .Cells(99, 99)).Font.Size = 11
         
        r_int_NroFil = r_int_NroFil + 2
         
        For r_int_NoFlLi = 0 To grd_Listad.Rows - 1

            .Cells(r_int_NroFil, 2) = "'" & grd_Listad.TextMatrix(r_int_NoFlLi, 0)
            .Cells(r_int_NroFil, 3) = "'" & grd_Listad.TextMatrix(r_int_NoFlLi, 1)
            .Cells(r_int_NroFil, 4) = grd_Listad.TextMatrix(r_int_NoFlLi, 2)
            .Cells(r_int_NroFil, 5) = "'" & grd_Listad.TextMatrix(r_int_NoFlLi, 3)
            .Cells(r_int_NroFil, 6) = "'" & grd_Listad.TextMatrix(r_int_NoFlLi, 4)
            .Cells(r_int_NroFil, 7) = grd_Listad.TextMatrix(r_int_NoFlLi, 5)
            .Cells(r_int_NroFil, 8) = grd_Listad.TextMatrix(r_int_NoFlLi, 6)
            .Cells(r_int_NroFil, 9) = grd_Listad.TextMatrix(r_int_NoFlLi, 7)
            
            r_int_NroFil = r_int_NroFil + 1
        Next r_int_NoFlLi
        
        With .Range(.Cells(10, 2), .Cells(r_int_NroFil, 3))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
    
        With .Range(.Cells(8, 2), .Cells(r_int_NroFil - 1, 9))
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        End With
        
        With .Range(.Cells(4, 3), .Cells(6, 3))
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
            .WrapText = False
        End With
   End With
   
   r_obj_Excel.Visible = True
End Sub
