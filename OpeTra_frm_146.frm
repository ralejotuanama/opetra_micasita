VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frm_Ges_CreHip_11 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6420
   ClientLeft      =   3825
   ClientTop       =   2835
   ClientWidth     =   11595
   Icon            =   "OpeTra_frm_146.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   6435
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11610
      _Version        =   65536
      _ExtentX        =   20479
      _ExtentY        =   11351
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
      Begin Threed.SSPanel SSPanel5 
         Height          =   765
         Left            =   30
         TabIndex        =   5
         Top             =   5610
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
         Begin VB.ComboBox cmb_BanChq 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   390
            Width           =   3225
         End
         Begin VB.TextBox txt_NumChq 
            Height          =   315
            Left            =   1590
            MaxLength       =   25
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   60
            Width           =   3225
         End
         Begin VB.ComboBox cmb_CtaChq 
            Height          =   315
            Left            =   7590
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   390
            Width           =   3225
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Nro. Cuenta:"
            Height          =   285
            Index           =   11
            Left            =   5820
            TabIndex        =   8
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Banco:"
            Height          =   315
            Index           =   7
            Left            =   60
            TabIndex        =   7
            Top             =   390
            Width           =   1365
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Nro. Cheque:"
            Height          =   285
            Index           =   16
            Left            =   60
            TabIndex        =   6
            Top             =   60
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   9
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
            Picture         =   "OpeTra_frm_146.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "OpeTra_frm_146.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   11
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
            Left            =   600
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
            Left            =   600
            TabIndex        =   24
            Top             =   330
            Width           =   4215
            _Version        =   65536
            _ExtentX        =   7435
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Retención de Fondos"
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
            Picture         =   "OpeTra_frm_146.frx":0890
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   3315
         Left            =   30
         TabIndex        =   12
         Top             =   2250
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   5847
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
         Begin TabDlg.SSTab SSTab1 
            Height          =   3195
            Left            =   60
            TabIndex        =   13
            Top             =   60
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   5636
            _Version        =   393216
            Style           =   1
            Tabs            =   7
            TabsPerRow      =   7
            TabHeight       =   520
            TabCaption(0)   =   "Datos del Cliente"
            TabPicture(0)   =   "OpeTra_frm_146.frx":0B9A
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Datos del Cónyuge"
            TabPicture(1)   =   "OpeTra_frm_146.frx":0BB6
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(5)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Datos del Inmueble"
            TabPicture(2)   =   "OpeTra_frm_146.frx":0BD2
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(1)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Informe Legal"
            TabPicture(3)   =   "OpeTra_frm_146.frx":0BEE
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "txt_InfLeg"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Datos Legal"
            TabPicture(4)   =   "OpeTra_frm_146.frx":0C0A
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "grd_Listad(2)"
            Tab(4).ControlCount=   1
            TabCaption(5)   =   "Datos del Crédito"
            TabPicture(5)   =   "OpeTra_frm_146.frx":0C26
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "grd_Listad(4)"
            Tab(5).ControlCount=   1
            TabCaption(6)   =   "Datos Desembolso"
            TabPicture(6)   =   "OpeTra_frm_146.frx":0C42
            Tab(6).ControlEnabled=   0   'False
            Tab(6).Control(0)=   "txt_ObsDes"
            Tab(6).Control(1)=   "grd_Listad(3)"
            Tab(6).ControlCount=   2
            Begin VB.TextBox txt_InfLeg 
               Height          =   2775
               Left            =   -74940
               MaxLength       =   8000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   29
               Text            =   "OpeTra_frm_146.frx":0C5E
               Top             =   360
               Width           =   11295
            End
            Begin VB.TextBox txt_ObsDes 
               Height          =   675
               Left            =   -74970
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   25
               Text            =   "OpeTra_frm_146.frx":0C62
               Top             =   2490
               Width           =   11325
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   2775
               Index           =   0
               Left            =   30
               TabIndex        =   14
               Top             =   360
               Width           =   11325
               _ExtentX        =   19976
               _ExtentY        =   4895
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   2055
               Index           =   3
               Left            =   -74970
               TabIndex        =   26
               Top             =   390
               Width           =   11325
               _ExtentX        =   19976
               _ExtentY        =   3625
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   2775
               Index           =   4
               Left            =   -74970
               TabIndex        =   27
               Top             =   360
               Width           =   11325
               _ExtentX        =   19976
               _ExtentY        =   4895
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   2775
               Index           =   2
               Left            =   -74970
               TabIndex        =   28
               Top             =   360
               Width           =   11295
               _ExtentX        =   19923
               _ExtentY        =   4895
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   2775
               Index           =   1
               Left            =   -74970
               TabIndex        =   30
               Top             =   360
               Width           =   11325
               _ExtentX        =   19976
               _ExtentY        =   4895
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   2775
               Index           =   5
               Left            =   -74970
               TabIndex        =   31
               Top             =   360
               Width           =   11325
               _ExtentX        =   19976
               _ExtentY        =   4895
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin VB.Label Label6 
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
               Left            =   -74970
               TabIndex        =   17
               Top             =   2160
               Width           =   2805
            End
            Begin VB.Label Label59 
               Caption         =   "Comité de Créditos"
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
               Left            =   -74970
               TabIndex        =   16
               Top             =   360
               Width           =   2805
            End
            Begin VB.Label Label3 
               Caption         =   "Contratos y Bloqueo Registral"
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
               Left            =   -74970
               TabIndex        =   15
               Top             =   1530
               Width           =   2805
            End
         End
      End
      Begin Threed.SSPanel SSPanel3 
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
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   1560
            TabIndex        =   19
            Top             =   390
            Width           =   9945
            _Version        =   65536
            _ExtentX        =   17542
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
         Begin Threed.SSPanel pnl_NumOpe 
            Height          =   315
            Left            =   1560
            TabIndex        =   20
            Top             =   60
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "001-01-00005"
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
         End
         Begin VB.Label Label5 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   22
            Top             =   390
            Width           =   1395
         End
         Begin VB.Label Label7 
            Caption         =   "Nro. de Operación:"
            Height          =   315
            Left            =   60
            TabIndex        =   21
            Top             =   60
            Width           =   1395
         End
      End
   End
End
Attribute VB_Name = "frm_Ges_CreHip_11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_BanChq()      As moddat_tpo_Genera
Dim l_arr_CtaChq()      As moddat_tpo_Genera

Private Sub cmb_BanChq_Click()
   Call gs_SetFocus(cmb_CtaChq)
   
   If cmb_BanChq.ListIndex > -1 Then
      Screen.MousePointer = 11
      Call moddat_gs_Carga_CtaBan(l_arr_BanChq(cmb_BanChq.ListIndex + 1).Genera_Codigo, cmb_CtaChq, l_arr_CtaChq)
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmb_BanChq_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_BanChq_Click
   End If
End Sub

Private Sub cmb_CtaChq_Click()
   Call gs_SetFocus(cmd_Grabar)
End Sub

Private Sub cmb_CtaChq_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CtaChq_Click
   End If
End Sub

Private Sub cmd_Grabar_Click()
   If Len(Trim(txt_NumChq.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Cheque.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumChq)
      Exit Sub
   End If
   
   If cmb_BanChq.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Banco del Cheque.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_BanChq)
      Exit Sub
   End If

   If cmb_CtaChq.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Cuenta del Cheque.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CtaChq)
      Exit Sub
   End If

   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   'Grabando Cabecera de Credito
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CRE_HIPDES_REGCHQ ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NumChq.Text & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_CtaChq(cmb_CtaChq.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_BanChq(cmb_BanChq.ListIndex + 1).Genera_Codigo & "', "
      
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

   MsgBox "Se registro la información correctamente.", vbInformation, modgen_g_str_NomPlt
   
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   
   Dim r_arr_Mtz()      As moddat_g_tpo_DatCom
   
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_NumOpe.Caption = ""
   pnl_NomCli.Caption = ""
   
   pnl_NumOpe.Caption = gf_Formato_NumOpe(moddat_g_str_NumOpe)
   pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicia
   
   'Buscando información de la solicitud
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
      
   Call modmip_gs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(0), 0)      'Buscar Información del Cliente
   Call modmip_gs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, grd_Listad(5), 1)      'Buscar Información del Cónyuge
   
   'Buscar Datos del Inmueble
   Call modmip_gs_DatInm(grd_Listad(1), True)

   Call fs_DatLeg
   Call fs_DatDes
   'Call fs_DatCre
   Call modmip_gs_DatCre(grd_Listad(4), r_arr_Mtz)
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub grd_Listad_SelChange(Index As Integer)
   If grd_Listad(Index).Rows > 2 Then
      grd_Listad(Index).RowSel = grd_Listad(Index).Row
   End If
End Sub

Private Sub fs_DatLeg()
   Call gs_LimpiaGrid(grd_Listad(2))

   g_str_Parame = "SELECT * FROM TRA_EVALEG WHERE "
   g_str_Parame = g_str_Parame & "EVALEG_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      txt_InfLeg.Text = Trim(g_rst_Princi!EVALEG_INFLG1 & "") & Trim(g_rst_Princi!EVALEG_INFLG2 & "") & Trim(g_rst_Princi!EVALEG_INFLG3 & "") & Trim(g_rst_Princi!EVALEG_INFLG4 & "")
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Fecha Firma Contrato"
      
      grd_Listad(2).Col = 1
      grd_Listad(2).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FIRCON))
   
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Notaria"
      
      grd_Listad(2).Col = 1
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("509", g_rst_Princi!EVALEG_CODNOT)
   
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Representante Legal 1"
      
      grd_Listad(2).Col = 1
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("512", g_rst_Princi!EVALEG_REPLG1)
   
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Representante Legal 2"
      
      grd_Listad(2).Col = 1
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("512", g_rst_Princi!EVALEG_REPLG2)
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Monto Hipoteca "
      
      grd_Listad(2).Col = 1
      grd_Listad(2).CellFontName = "Lucida Console"
      grd_Listad(2).CellFontSize = 8
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("229", g_rst_Princi!EVALEG_MONHIP) & " " & gf_FormatoNumero(g_rst_Princi!EVALEG_MTOHIP, 12, 2)
         
      Call gs_UbiIniGrid(grd_Listad(2))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatDes()
   Call gs_LimpiaGrid(grd_Listad(3))
   txt_ObsDes.Text = ""
   
   g_str_Parame = "SELECT * FROM CRE_HIPDES WHERE "
   g_str_Parame = g_str_Parame & "HIPDES_NUMOPE = '" & moddat_g_str_NumOpe & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_Listad(3).Rows = grd_Listad(3).Rows + 1
      grd_Listad(3).Row = grd_Listad(3).Rows - 1
      grd_Listad(3).Col = 0
      grd_Listad(3).Text = "Tipo de Desembolso"
      
      grd_Listad(3).Col = 1
      grd_Listad(3).Text = moddat_gf_Consulta_ParDes("241", g_rst_Princi!HIPDES_TIPGAR)
      
      grd_Listad(3).Rows = grd_Listad(3).Rows + 1
      grd_Listad(3).Row = grd_Listad(3).Rows - 1
      grd_Listad(3).Col = 0
      grd_Listad(3).Text = "Fecha de Desembolso"
      
      grd_Listad(3).Col = 1
      grd_Listad(3).Text = gf_FormatoFecha(g_rst_Princi!HIPDES_FECDES)
      
      
      If g_rst_Princi!HIPDES_TIPDES = 1 Then
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).Text = "Nro. de Cheque"
         
         grd_Listad(3).Col = 1
         grd_Listad(3).Text = Trim(g_rst_Princi!HIPDES_CHECGO & "")
         
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).Text = "Banco Emisor (Cuenta)"
         
         grd_Listad(3).Col = 1
         grd_Listad(3).Text = moddat_gf_Consulta_ParDes("505", g_rst_Princi!HIPDES_BANCGO & "") & " (" & Trim(g_rst_Princi!HIPDES_CTACGO & "") & ")"
      End If
      
      grd_Listad(3).Rows = grd_Listad(3).Rows + 1
      grd_Listad(3).Row = grd_Listad(3).Rows - 1
      grd_Listad(3).Col = 0
      grd_Listad(3).Text = "Importe Desembolsado"
      
      grd_Listad(3).Col = 1
      grd_Listad(3).CellFontName = "Lucida Console"
      grd_Listad(3).CellFontSize = 8
      grd_Listad(3).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPDES_DESMPR, 12, 2)
      
      
      If g_rst_Princi!HIPDES_TIPGAR = 4 Then
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).Text = "Banco Emisor Carta Fianza"
         
         grd_Listad(3).Col = 1
         If Not IsNull(g_rst_Princi!HIPDES_BANFIA) Then
            grd_Listad(3).Text = moddat_gf_Consulta_ParDes("505", g_rst_Princi!HIPDES_BANFIA)
         End If
      
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).Text = "Nro. Carta Fianza"
         
         grd_Listad(3).Col = 1
         grd_Listad(3).Text = Trim(g_rst_Princi!HIPDES_NUMFIA & "")
      
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).Text = "Fecha Emisión"
         
         grd_Listad(3).Col = 1
         grd_Listad(3).Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPDES_EMIFIA))
      
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).Text = "Fecha Vencimiento"
         
         grd_Listad(3).Col = 1
         grd_Listad(3).Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPDES_VCTFIA))
         
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).Text = "Importe Carta Fianza"
         
         grd_Listad(3).Col = 1
         grd_Listad(3).CellFontName = "Lucida Console"
         grd_Listad(3).CellFontSize = 8
         grd_Listad(3).Text = moddat_gf_Consulta_ParDes("229", g_rst_Princi!HIPDES_MONFIA) & " " & gf_FormatoNumero(g_rst_Princi!HIPDES_IMPFIA, 12, 2)
      End If
      
      If g_rst_Princi!HIPDES_TIPGAR = 5 Then
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).Text = "Banco Emisor Certificado"
         
         If Not IsNull(g_rst_Princi!HIPDES_BCOGAR) Then
            grd_Listad(3).Col = 1
            grd_Listad(3).Text = moddat_gf_Consulta_ParDes("505", CStr(g_rst_Princi!HIPDES_BCOGAR))
         End If
      
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).Text = "Nro. Certificado"
         
         grd_Listad(3).Col = 1
         grd_Listad(3).Text = Trim(g_rst_Princi!HIPDES_DOCGAR & "")
         
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).Text = "Importe Certificado"
         
         grd_Listad(3).Col = 1
         grd_Listad(3).CellFontName = "Lucida Console"
         grd_Listad(3).CellFontSize = 8
         grd_Listad(3).Text = moddat_gf_Consulta_ParDes("229", g_rst_Princi!HIPDES_MONGAR) & " " & gf_FormatoNumero(g_rst_Princi!HIPDES_MTOGAR, 12, 2)
      End If
      
      Call gs_UbiIniGrid(grd_Listad(3))
      
      txt_ObsDes.Text = Trim(g_rst_Princi!HIPDES_OBSERV & "")
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Inicia()
   Dim r_int_Contad     As Integer

   'Datos del Cliente
   grd_Listad(0).ColWidth(0) = 3060:   grd_Listad(0).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(0).ColWidth(1) = 7940:   grd_Listad(0).ColAlignment(1) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_Listad(0))

   'Datos del Conyuge
   grd_Listad(5).ColWidth(0) = 3060:   grd_Listad(5).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(5).ColWidth(1) = 7940:   grd_Listad(5).ColAlignment(1) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_Listad(5))
   
   'Datos del Inmueble
   grd_Listad(1).ColWidth(0) = 3060:   grd_Listad(1).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(1).ColWidth(1) = 7940:   grd_Listad(1).ColAlignment(1) = flexAlignLeftCenter

   Call gs_LimpiaGrid(grd_Listad(1))
   
   'Datos Legal
   grd_Listad(2).ColWidth(0) = 3060:   grd_Listad(2).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(2).ColWidth(1) = 7940:   grd_Listad(2).ColAlignment(1) = flexAlignLeftCenter

   Call gs_LimpiaGrid(grd_Listad(2))

   'Datos del Crédito
   grd_Listad(4).ColWidth(0) = 3060:   grd_Listad(4).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(4).ColWidth(1) = 7940:   grd_Listad(4).ColAlignment(1) = flexAlignLeftCenter

   Call gs_LimpiaGrid(grd_Listad(4))

   'Datos del Desembolso
   grd_Listad(3).ColWidth(0) = 3060:   grd_Listad(3).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(3).ColWidth(1) = 7940:   grd_Listad(3).ColAlignment(1) = flexAlignLeftCenter

   Call gs_LimpiaGrid(grd_Listad(3))

   Call moddat_gs_Carga_LisIte(cmb_BanChq, l_arr_BanChq, 1, "516")

   txt_NumChq.Text = ""
   cmb_BanChq.ListIndex = -1
   cmb_CtaChq.Clear
End Sub

Private Sub fs_DatCre()
   Dim r_str_CodPry     As String
   Dim r_str_NomPry     As String
   Dim r_str_CodBco     As String
   
   'Buscando Información del Crédito
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If

   g_rst_Princi.MoveFirst
   
   'Cargando en Grid
   grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   grd_Listad(4).Row = grd_Listad(4).Rows - 1
   grd_Listad(4).Col = 0
   grd_Listad(4).Text = "Moneda Préstamo"
   
   grd_Listad(4).Col = 1
   grd_Listad(4).Text = moddat_gf_Consulta_ParDes("204", CStr(moddat_g_int_TipMon))
   
   If moddat_g_int_TipMon = 1 Then
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).Text = "Valor Compra Venta"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellFontName = "Lucida Console"
      grd_Listad(4).CellFontSize = 8
      grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_CVTSOL, 12, 2)
   
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).Text = "Aporte Propio"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellFontName = "Lucida Console"
      grd_Listad(4).CellFontSize = 8
      grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APOSOL, 12, 2)
   Else
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).Text = "Valor Compra Venta"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellFontName = "Lucida Console"
      grd_Listad(4).CellFontSize = 8
      grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_CVTDOL, 12, 2)
   
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).Text = "Aporte Propio"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellFontName = "Lucida Console"
      grd_Listad(4).CellFontSize = 8
      grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APODOL, 12, 2)
   End If
   
   grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   grd_Listad(4).Row = grd_Listad(4).Rows - 1
   grd_Listad(4).Col = 0
   grd_Listad(4).Text = "Monto Préstamo"
   
   grd_Listad(4).Col = 1
   grd_Listad(4).CellFontName = "Lucida Console"
   grd_Listad(4).CellFontSize = 8
   grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_MTOPRE, 12, 2)
   
   If g_rst_Princi!HIPMAE_FECESC > 0 Then
      grd_Listad(4).Rows = grd_Listad(4).Rows + 2
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).Text = "Fecha Firma EE.PP"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECESC))
   End If
   
   grd_Listad(4).Rows = grd_Listad(4).Rows + 2
   grd_Listad(4).Row = grd_Listad(4).Rows - 1
   grd_Listad(4).Col = 0
   grd_Listad(4).Text = "Plazo"
   
   grd_Listad(4).Col = 1
   grd_Listad(4).Text = CStr(g_rst_Princi!HIPMAE_PLAANO) & " Años"
   
   grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   grd_Listad(4).Row = grd_Listad(4).Rows - 1
   grd_Listad(4).Col = 0
   grd_Listad(4).Text = "Tasa de Interés"
   
   grd_Listad(4).Col = 1
   grd_Listad(4).Text = Format(g_rst_Princi!HIPMAE_TASINT, "##0.00") & " %"
   
   grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   grd_Listad(4).Row = grd_Listad(4).Rows - 1
   grd_Listad(4).Col = 0
   grd_Listad(4).Text = "Nro. de Cuotas"
   
   grd_Listad(4).Col = 1
   grd_Listad(4).Text = CStr(g_rst_Princi!HIPMAE_NUMCUO)
   
   grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   grd_Listad(4).Row = grd_Listad(4).Rows - 1
   grd_Listad(4).Col = 0
   grd_Listad(4).Text = "Período de Gracia"
   
   grd_Listad(4).Col = 1
   grd_Listad(4).Text = CStr(g_rst_Princi!HIPMAE_PERGRA) & " Meses"
   
   grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   grd_Listad(4).Row = grd_Listad(4).Rows - 1
   grd_Listad(4).Col = 0
   grd_Listad(4).Text = "Compañía de Seguros"
   
   grd_Listad(4).Col = 1
   grd_Listad(4).Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!HIPMAE_SEGPRE & "")
   
   grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   grd_Listad(4).Row = grd_Listad(4).Rows - 1
   grd_Listad(4).Col = 0
   grd_Listad(4).Text = "Tipo de Seguro Desg."
   
   grd_Listad(4).Col = 1
   grd_Listad(4).Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!HIPMAE_SEGPRE, g_rst_Princi!HIPMAE_TIPSEG)
   
   grd_Listad(4).Rows = grd_Listad(4).Rows + 2
   grd_Listad(4).Row = grd_Listad(4).Rows - 1
   grd_Listad(4).Col = 0
   grd_Listad(4).Text = "Consejero Hipotecario"
   
   grd_Listad(4).Col = 1
   grd_Listad(4).Text = moddat_gf_Buscar_NomEje(g_rst_Princi!HIPMAE_CONHIP)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_Listad(4))
End Sub

Private Sub txt_NumChq_GotFocus()
   Call gs_SelecTodo(txt_NumChq)
End Sub

Private Sub txt_NumChq_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_BanChq)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
   End If
End Sub



