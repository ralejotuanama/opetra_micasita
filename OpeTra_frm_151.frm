VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_Ges_CreHip_16 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7155
   ClientLeft      =   1950
   ClientTop       =   1860
   ClientWidth     =   14820
   Icon            =   "OpeTra_frm_151.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   14820
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7155
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   14865
      _Version        =   65536
      _ExtentX        =   26220
      _ExtentY        =   12621
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
      Begin Threed.SSPanel SSPanel39 
         Height          =   645
         Left            =   60
         TabIndex        =   3
         Top             =   780
         Width           =   14745
         _Version        =   65536
         _ExtentX        =   26009
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
            Left            =   14100
            Picture         =   "OpeTra_frm_151.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   4
         Top             =   60
         Width           =   14745
         _Version        =   65536
         _ExtentX        =   26009
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
            TabIndex        =   5
            Top             =   30
            Width           =   5685
            _Version        =   65536
            _ExtentX        =   10028
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
            Left            =   690
            TabIndex        =   6
            Top             =   330
            Width           =   5505
            _Version        =   65536
            _ExtentX        =   9710
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Exnoración de Cargos por Cobranza Morosa"
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
            Picture         =   "OpeTra_frm_151.frx":044E
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   60
         TabIndex        =   7
         Top             =   1470
         Width           =   14745
         _Version        =   65536
         _ExtentX        =   26009
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
            Left            =   1560
            TabIndex        =   8
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
            Left            =   1560
            TabIndex        =   9
            Top             =   390
            Width           =   13095
            _Version        =   65536
            _ExtentX        =   23098
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
         Begin VB.Label Label5 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   11
            Top             =   390
            Width           =   1395
         End
         Begin VB.Label Label12 
            Caption         =   "Nro. Operación:"
            Height          =   315
            Left            =   60
            TabIndex        =   10
            Top             =   60
            Width           =   1245
         End
      End
      Begin Threed.SSPanel SSPanel31 
         Height          =   4815
         Left            =   60
         TabIndex        =   12
         Top             =   2280
         Width           =   14745
         _Version        =   65536
         _ExtentX        =   26009
         _ExtentY        =   8493
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
            Height          =   3735
            Left            =   60
            TabIndex        =   0
            Top             =   630
            Width           =   14625
            _ExtentX        =   25797
            _ExtentY        =   6588
            _Version        =   393216
            Rows            =   21
            Cols            =   10
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel32 
            Height          =   285
            Left            =   90
            TabIndex        =   13
            Top             =   330
            Width           =   765
            _Version        =   65536
            _ExtentX        =   1349
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cuota"
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
            Left            =   6960
            TabIndex        =   14
            Top             =   330
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Int. Comp."
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
         Begin Threed.SSPanel SSPanel40 
            Height          =   285
            Left            =   8010
            TabIndex        =   15
            Top             =   330
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Int. Morat."
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
         Begin Threed.SSPanel pnl_TotPag 
            Height          =   315
            Left            =   13260
            TabIndex        =   16
            Top             =   4410
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   192
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
            Alignment       =   4
         End
         Begin Threed.SSPanel SSPanel18 
            Height          =   285
            Left            =   9060
            TabIndex        =   17
            Top             =   330
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Gastos Cobr."
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
         Begin Threed.SSPanel SSPanel19 
            Height          =   285
            Left            =   10110
            TabIndex        =   18
            Top             =   330
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Otros Gastos"
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
         Begin Threed.SSPanel SSPanel20 
            Height          =   285
            Left            =   13260
            TabIndex        =   19
            Top             =   330
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Total"
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
            Left            =   11160
            TabIndex        =   20
            Top             =   330
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Capital PBP"
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
         Begin Threed.SSPanel SSPanel4 
            Height          =   285
            Left            =   12210
            TabIndex        =   21
            Top             =   330
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Interés PBP"
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
            Left            =   840
            TabIndex        =   23
            Top             =   330
            Width           =   3765
            _Version        =   65536
            _ExtentX        =   6641
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Motivo Exoneración"
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   4590
            TabIndex        =   24
            Top             =   330
            Width           =   2385
            _Version        =   65536
            _ExtentX        =   4207
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nivel Autorización"
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
         Begin VB.Label lbl_Totale 
            Alignment       =   1  'Right Justify
            Caption         =   "Total Pagado ==> US$"
            Height          =   315
            Left            =   11460
            TabIndex        =   22
            Top             =   4410
            Width           =   1755
         End
      End
   End
End
Attribute VB_Name = "frm_Ges_CreHip_16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   pnl_NumOpe.Caption = ""
   pnl_NomCli.Caption = ""
   
   pnl_NumOpe.Caption = gf_Formato_NumOpe(moddat_g_str_NumOpe)
   pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   Call fs_Inicia
   Call fs_Buscar_LisExo
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 745
   grd_Listad.ColWidth(1) = 3755
   grd_Listad.ColWidth(2) = 2375
   grd_Listad.ColWidth(3) = 1055
   grd_Listad.ColWidth(4) = 1055
   grd_Listad.ColWidth(5) = 1055
   grd_Listad.ColWidth(6) = 1055
   grd_Listad.ColWidth(7) = 1055
   grd_Listad.ColWidth(8) = 1055
   grd_Listad.ColWidth(9) = 1055
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_Listad.ColAlignment(8) = flexAlignRightCenter
   grd_Listad.ColAlignment(9) = flexAlignRightCenter
End Sub

Private Sub fs_Buscar_LisExo()
   Dim r_dbl_TotCuo     As Double
   Dim r_dbl_TotPag     As Double

   Call gs_LimpiaGrid(grd_Listad)
   lbl_Totale.Caption = "Total ===> " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " "

   'Obteniendo Información de Pagos
   g_str_Parame = "SELECT * FROM CRE_HIPEXO WHERE "
   g_str_Parame = g_str_Parame & "HIPEXO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "ORDER BY HIPEXO_NUMCUO DESC, HIPEXO_NUMEXO DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_dbl_TotPag = 0
      
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         r_dbl_TotCuo = 0
         
         grd_Listad.Col = 0
         grd_Listad.Text = CStr(g_rst_Princi!HIPEXO_NUMCUO)
                  
         grd_Listad.Col = 1
         grd_Listad.Text = moddat_gf_Consulta_ParDes("252", CStr(g_rst_Princi!HIPEXO_MOTEXO))
         
         grd_Listad.Col = 2
         grd_Listad.Text = moddat_gf_Consulta_ParDes("253", CStr(g_rst_Princi!HIPEXO_AUTEXO))
         
         grd_Listad.Col = 3
         grd_Listad.Text = Format(g_rst_Princi!HIPEXO_ANTMOR, "###,###,##0.00")
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_Listad.Text)
         
         grd_Listad.Col = 4
         grd_Listad.Text = Format(g_rst_Princi!HIPEXO_ANTCOM, "###,###,##0.00")
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_Listad.Text)
         
         grd_Listad.Col = 5
         grd_Listad.Text = Format(g_rst_Princi!HIPEXO_ANTCOB, "###,###,##0.00")
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_Listad.Text)
         
         grd_Listad.Col = 6
         grd_Listad.Text = Format(g_rst_Princi!HIPEXO_ANTOTR, "###,###,##0.00")
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_Listad.Text)
         
         grd_Listad.Col = 7
         grd_Listad.Text = Format(g_rst_Princi!HIPEXO_ANTCBP, "###,###,##0.00")
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_Listad.Text)
         
         grd_Listad.Col = 8
         grd_Listad.Text = Format(g_rst_Princi!HIPEXO_ANTIBP, "###,###,##0.00")
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_Listad.Text)
         
         grd_Listad.Col = 9
         grd_Listad.Text = Format(r_dbl_TotCuo, "###,###,##0.00")
         
         r_dbl_TotPag = r_dbl_TotPag + r_dbl_TotCuo
      
         g_rst_Princi.MoveNext
      Loop
      
      Call gs_UbiIniGrid(grd_Listad)
      
      pnl_TotPag.Caption = Format(r_dbl_TotPag, "###,###,##0.00") & " "
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

