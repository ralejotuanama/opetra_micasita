VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_SolCre_51 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10095
   ClientLeft      =   3090
   ClientTop       =   525
   ClientWidth     =   11460
   Icon            =   "OpeTra_frm_156.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10095
   ScaleWidth      =   11460
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10095
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11475
      _Version        =   65536
      _ExtentX        =   20241
      _ExtentY        =   17806
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
         Height          =   1695
         Left            =   60
         TabIndex        =   33
         Top             =   8340
         Width           =   11355
         _Version        =   65536
         _ExtentX        =   20029
         _ExtentY        =   2990
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
            Height          =   1305
            Index           =   2
            Left            =   60
            TabIndex        =   34
            Top             =   360
            Width           =   11235
            _ExtentX        =   19817
            _ExtentY        =   2302
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.Label Label7 
            Caption         =   "Información del Apoderado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   60
            TabIndex        =   35
            Top             =   60
            Width           =   3195
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   645
         Left            =   60
         TabIndex        =   19
         Top             =   780
         Width           =   11355
         _Version        =   65536
         _ExtentX        =   20029
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
         Begin VB.CommandButton cmd_Patrim 
            Height          =   585
            Left            =   1830
            Picture         =   "OpeTra_frm_156.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Información Financiera"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_156.frx":012F
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Buscar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_156.frx":0439
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10740
            Picture         =   "OpeTra_frm_156.frx":0743
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_SimCre 
            Height          =   585
            Left            =   4830
            Picture         =   "OpeTra_frm_156.frx":0B85
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Simulación de Créditos Hipotecarios"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DatCli 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_156.frx":0E8F
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Datos del Cónyuge"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Refere 
            Height          =   585
            Left            =   2430
            Picture         =   "OpeTra_frm_156.frx":1199
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Referencias Personales"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DatInm 
            Height          =   585
            Left            =   3030
            Picture         =   "OpeTra_frm_156.frx":14A3
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Datos del Inmueble"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DatCre 
            Height          =   585
            Left            =   3630
            Picture         =   "OpeTra_frm_156.frx":1D6D
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Datos del Crédito"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   4230
            Picture         =   "OpeTra_frm_156.frx":2077
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   2715
         Left            =   60
         TabIndex        =   20
         Top             =   3300
         Width           =   11355
         _Version        =   65536
         _ExtentX        =   20029
         _ExtentY        =   4789
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
         Begin VB.CommandButton cmd_LisRec_Cli 
            Caption         =   "Ver Solicitudes Rechazadas"
            Height          =   285
            Left            =   8220
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Lista de Solicitudes Rechazadas"
            Top             =   60
            Width           =   3045
         End
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   2295
            Index           =   0
            Left            =   60
            TabIndex        =   21
            Top             =   360
            Width           =   11235
            _ExtentX        =   19817
            _ExtentY        =   4048
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.Label Label4 
            Caption         =   "Información del Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   90
            TabIndex        =   22
            Top             =   90
            Width           =   2235
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   23
         Top             =   60
         Width           =   11355
         _Version        =   65536
         _ExtentX        =   20029
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
            Height          =   495
            Left            =   630
            TabIndex        =   24
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Solicitud de Crédito Hipotecario"
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
            Picture         =   "OpeTra_frm_156.frx":24B9
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1785
         Left            =   60
         TabIndex        =   25
         Top             =   1470
         Width           =   11355
         _Version        =   65536
         _ExtentX        =   20029
         _ExtentY        =   3149
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
         Begin VB.ComboBox cmb_FerEve 
            Height          =   315
            Left            =   7590
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1410
            Width           =   3705
         End
         Begin VB.ComboBox cmb_ForCon 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1410
            Width           =   3705
         End
         Begin VB.ComboBox cmb_OfiCom 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1080
            Width           =   9405
         End
         Begin VB.ComboBox cmb_Produc 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   60
            Width           =   9405
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   750
            Width           =   3705
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   7560
            MaxLength       =   12
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   750
            Width           =   3705
         End
         Begin VB.ComboBox cmb_SubPrd 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   390
            Width           =   9405
         End
         Begin VB.Label Label10 
            Caption         =   "Feria / Evento:"
            Height          =   315
            Left            =   5760
            TabIndex        =   39
            Top             =   1410
            Width           =   1695
         End
         Begin VB.Label Label9 
            Caption         =   "Forma Contacto:"
            Height          =   315
            Left            =   90
            TabIndex        =   38
            Top             =   1410
            Width           =   1695
         End
         Begin VB.Label Label8 
            Caption         =   "Ofic. At. Comercial:"
            Height          =   315
            Left            =   90
            TabIndex        =   37
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   90
            TabIndex        =   29
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   28
            Top             =   750
            Width           =   1755
         End
         Begin VB.Label Label3 
            Caption         =   "Nro. Docum. Identidad:"
            Height          =   285
            Left            =   5760
            TabIndex        =   27
            Top             =   750
            Width           =   1725
         End
         Begin VB.Label Label6 
            Caption         =   "Sub-Producto:"
            Height          =   315
            Left            =   90
            TabIndex        =   26
            Top             =   390
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   2235
         Left            =   60
         TabIndex        =   30
         Top             =   6060
         Width           =   11355
         _Version        =   65536
         _ExtentX        =   20029
         _ExtentY        =   3942
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
         Begin VB.CommandButton cmd_LisRec_Cyg 
            Caption         =   "Ver Solicitudes Rechazadas"
            Height          =   285
            Left            =   8220
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Lista de Solicitudes Rechazadas"
            Top             =   60
            Width           =   3045
         End
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   1845
            Index           =   1
            Left            =   60
            TabIndex        =   31
            Top             =   360
            Width           =   11235
            _ExtentX        =   19817
            _ExtentY        =   3254
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.Label Label5 
            Caption         =   "Información del Cónyuge"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   60
            TabIndex        =   32
            Top             =   60
            Width           =   2235
         End
      End
   End
End
Attribute VB_Name = "frm_SolCre_51"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_Produc()      As moddat_tpo_Genera
Dim l_arr_SubPrd()      As moddat_tpo_Genera
Dim l_arr_Parame()      As moddat_tpo_Genera
Dim l_int_ActPri_Cyg    As Integer
Dim l_int_Flg_ActTit    As Integer
Dim l_int_Flg_ActCyg    As Integer

Private Sub cmb_FerEve_Click()
   Call gs_SetFocus(cmd_Buscar)
End Sub

Private Sub cmb_FerEve_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_FerEve_Click
   End If
End Sub

Private Sub cmb_ForCon_Click()
   If cmb_ForCon.ListIndex > -1 Then
      If cmb_ForCon.ItemData(cmb_ForCon.ListIndex) = 21 Then
         cmb_FerEve.Enabled = True
         Call gs_SetFocus(cmb_FerEve)
      Else
         cmb_FerEve.ListIndex = -1
         cmb_FerEve.Enabled = False
         Call gs_SetFocus(cmd_Buscar)
      End If
   End If
End Sub

Private Sub cmb_ForCon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_ForCon_Click
   End If
End Sub

Private Sub cmb_OfiCom_Click()
   Call gs_SetFocus(cmb_ForCon)
End Sub

Private Sub cmb_OfiCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_OfiCom_Click
   End If
End Sub

Private Sub cmb_Produc_Click()
   cmb_SubPrd.Clear
   moddat_g_str_CodPrd = ""
   
   If cmb_Produc.ListIndex > -1 Then
      Screen.MousePointer = 11
      moddat_g_str_CodPrd = l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo
      Call moddat_gs_Carga_SubPrd(cmb_SubPrd, l_arr_SubPrd, moddat_g_str_CodPrd)
      Call gs_SetFocus(cmb_SubPrd)
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmb_Produc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Produc_Click
   End If
End Sub

Private Sub cmb_SubPrd_Click()
   Call gs_SetFocus(cmb_TipDoc)
End Sub

Private Sub cmb_SubPrd_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_SubPrd_Click
   End If
End Sub

Private Sub cmb_TipDoc_Click()
   If cmb_TipDoc.ListIndex > -1 Then
      Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
         Case 1:     txt_NumDoc.MaxLength = 8
         Case 7:     txt_NumDoc.MaxLength = 12
         Case Else:  txt_NumDoc.MaxLength = 12
      End Select
   End If
   
   Call gs_SetFocus(txt_NumDoc)
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumDoc)
   End If
End Sub

Private Sub cmd_Buscar_Click()
Dim r_int_valida As Integer

   If cmb_Produc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Produc)
      Exit Sub
   End If
   If cmb_SubPrd.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Sub-Producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_SubPrd)
      Exit Sub
   End If
   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   If Len(Trim(txt_NumDoc.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   Else
      Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
         Case 1
            g_str_Parame = ""
            g_str_Parame = g_str_Parame & "SELECT PROMAE_TIPDOC,PROMAE_NUMDOC, TRIM(B.DATGEN_APEPAT) ||' '|| TRIM(B.DATGEN_APEMAT) ||' '||  TRIM(B.DATGEN_NOMBRE) AS NOMBRE_CLIENTE "
            g_str_Parame = g_str_Parame & "  FROM CRE_PROMAE A "
            g_str_Parame = g_str_Parame & "  LEFT JOIN CLI_DATGEN B ON B.DATGEN_TIPDOC = A.PROMAE_TIPDOC AND TRIM(B.DATGEN_NUMDOC) = TRIM(A.PROMAE_NUMDOC) "
            g_str_Parame = g_str_Parame & " WHERE PROMAE_TIPDOC = " & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & " "
            g_str_Parame = g_str_Parame & "   AND PROMAE_NUMDOC = " & Trim(txt_NumDoc.Text) & " "
   
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
               Exit Sub
            End If
      
            If g_rst_Princi.BOF And g_rst_Princi.EOF Then
               MsgBox "El cliente debe estar registrado en el 'Mantenedor de Prospectos'.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(txt_NumDoc)
               Screen.MousePointer = 0
               Exit Sub
            End If
      End Select
   End If
   
   'r_int_valida = validacionclientexproblemas(Trim(txt_NumDoc.Text), "")
   r_int_valida = PostWebservice(Trim(txt_NumDoc.Text), "")
         
   g_str_Parame = ""
   g_str_Parame = "USP_CRE_INSPEK ("
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'110', "
   g_str_Parame = g_str_Parame & "'" & Trim(txt_NumDoc.Text) & "', "
   g_str_Parame = g_str_Parame & "'" & Trim(modgen_g_str_rptwebservice) & "', "
   g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!NOMBRE_CLIENTE & "") & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
               
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      Screen.MousePointer = 0
      Exit Sub
   End If
     
   If r_int_valida = 0 Then
      MsgBox "El Cliente fue encontrado en la base inspektor.", vbExclamation, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   Screen.MousePointer = 0
   '---------------------------------------
   If cmb_OfiCom.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Oficina de Atención Comercial.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_OfiCom)
      Exit Sub
   End If
   If cmb_ForCon.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Forma de Contacto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_ForCon)
      Exit Sub
   End If
   If cmb_ForCon.ItemData(cmb_ForCon.ListIndex) = 21 Then
      If cmb_FerEve.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Feria o Evento de Contacto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_FerEve)
         Exit Sub
      End If
   End If

   If moddat_gf_Obtiene_TipCam(1, 2) = 0 Then
      MsgBox "No se encuentra registrado el Tipo de Cambio. No puede ingresar Solicitud para este Producto.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

   l_int_Flg_ActTit = 0
   l_int_Flg_ActCyg = 0
   
   If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1 Then
      txt_NumDoc.Text = Format(txt_NumDoc.Text, "00000000")
   End If
   
   moddat_g_str_CodSub = l_arr_SubPrd(cmb_SubPrd.ListIndex + 1).Genera_Codigo
   moddat_g_int_TipMon = l_arr_SubPrd(cmb_SubPrd.ListIndex + 1).Genera_TipMon
   moddat_g_int_TipDoc = cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
   moddat_g_str_TipDoc = Trim(cmb_TipDoc.Text)
   moddat_g_str_NumDoc = Trim(txt_NumDoc.Text)
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
   moddat_g_str_FecNac_Tit = ""
   moddat_g_str_FecNac_Cyg = ""
   moddat_g_int_RegCyg = 0
   moddat_g_int_EstCiv = 0
   moddat_g_int_ComRta = 0
   modmip_g_int_PaiRes = 0
   l_int_ActPri_Cyg = 0
   ReDim modatecli_g_arr_LisRec(0)
   ReDim modatecli_g_arr_CygRec(0)

   'Verificando que Cliente no haya sido ingresado como Cliente Negativo
   If Not atecli_gf_Buscar_BasNeg(moddat_g_int_TipDoc, moddat_g_str_NumDoc) Then
      Call cmd_Limpia_Click
      Exit Sub
   End If

   'Validar que Cliente no tenga una Solicitud de Crédito en Evaluación Como Titular
   If Not atecli_gf_Buscar_SolVig(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 1) Then
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   'Validar que Cliente no tenga una Solicitud de Crédito en Evaluación Como Cónyuge
   If Not atecli_gf_Buscar_SolVig(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 2) Then
      Call cmd_Limpia_Click
      Exit Sub
   End If

   'Buscando Operaciones de Crédito
   Call atecli_gs_Buscar_CreHip(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 1)
   
   If UBound(modatecli_g_arr_TitOpe) > 0 Then
      MsgBox "El Cliente ya tiene un Crédito Hipotecario registrado.", vbInformation, modgen_g_str_NomPlt
      Call cmd_Limpia_Click
      Exit Sub
   End If

   'Buscando Solicitudes Rechazadas de Cliente
   Call atecli_gs_Buscar_SolRec(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 1)

   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & moddat_g_str_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      moddat_g_int_FlgAct = 1
      moddat_g_int_FlgGrb = 1
      
      frm_MntCli_52.Show 1
      
      If moddat_g_int_FlgAct = 1 Then
         MsgBox "No ha registrado información del Cliente.", vbExclamation, modgen_g_str_NomPlt
         Call cmd_Limpia_Click
         Exit Sub
      End If
   Else
      moddat_g_int_CygTDo = g_rst_Princi!DATGEN_CYGTDO
      moddat_g_str_CygNDo = Trim(g_rst_Princi!DATGEN_CYGNDO & "")
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   Call fs_Activa(False)
   
   'Buscar Información del Cliente
   Call fs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(0), 0)
   
   'Buscando Información del Cónyuge
   If moddat_g_int_CygTDo > 0 Then
      'Verificando que Cónyuge no haya sido ingresado como Cliente Negativo
      If Not atecli_gf_Buscar_BasNeg(moddat_g_int_CygTDo, moddat_g_str_CygNDo) Then
         Call cmd_Limpia_Click
         Exit Sub
      End If
   
      'Validar que Cónyuge no tenga una Solicitud de Crédito en Evaluación Como Titular
      If Not atecli_gf_Buscar_SolVig(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 1) Then
         Call cmd_Limpia_Click
         Exit Sub
      End If
      
      'Validar que Cónyuge no tenga una Solicitud de Crédito en Evaluación Como Cónyuge
      If Not atecli_gf_Buscar_SolVig(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 2) Then
         Call cmd_Limpia_Click
         Exit Sub
      End If
      
      'Buscando Operaciones de Crédito
      Call atecli_gs_Buscar_CreHip(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 2)
      
      If UBound(modatecli_g_arr_CygOpe) > 0 Then
         MsgBox "El Cónyuge ya tiene un Crédito Hipotecario registrado.", vbInformation, modgen_g_str_NomPlt
         Call cmd_Limpia_Click
         Exit Sub
      End If
   
      'Buscando Solicitudes Rechazadas del Cónyuge
      Call atecli_gs_Buscar_SolRec(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 2)
      Call fs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, grd_Listad(1), 1)
   End If
   
   'Buscar Información del Apoderado
   Call fs_DatApo(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(2))
   
   If UBound(modatecli_g_arr_LisRec) = 0 Then
      cmd_LisRec_Cli.Enabled = False
   End If
   
   If UBound(modatecli_g_arr_CygRec) = 0 Then
      cmd_LisRec_Cyg.Enabled = False
   End If
   
   'Inicializando Arreglos
   Call modatecli_gs_Limpia_Refere(1)     'Referencias - Familiar
   Call modatecli_gs_Limpia_Refere(2)     'Referencias - No Familiar
   Call modatecli_gs_Limpia_DatInm        'Datos del Inmueble
   Call modatecli_gs_Limpia_DatCre        'Datos del Crédito
   
   ReDim modatecli_g_arr_IngresInv(0)     'Ingresos - Inversiones
   ReDim modatecli_g_arr_IngresInm(0)     'Ingresos - Inmuebles
   ReDim modatecli_g_arr_IngresAut(0)     'Ingresos - Autos
   ReDim modatecli_g_arr_IngresEns(0)     'Ingresos - Enseres
   ReDim modatecli_g_arr_GastosTar(0)     'Gastos - Tarjetas
   ReDim modatecli_g_arr_GastosFin(0)     'Gastos - Deudas Financieras
   ReDim modatecli_g_arr_GastosNFi(0)     'Gastos - Deudas No Financieras
   ReDim modatecli_g_arr_GastosGas(0)     'Gastos - Gastos Mensuales
   ReDim modatecli_g_arr_DocCre(0)        'Documentos Recibidos

   'Inicializando Flag de Datos Ingresados
   modatecli_g_int_GastosTit = 1
   modatecli_g_int_IngRegInm = 1
   modatecli_g_int_GasRegTar = 1
   modatecli_g_int_GasRegFin = 1
   modatecli_g_int_GasRegGas = 1
   modatecli_g_int_RefereTit = 1
   modatecli_g_int_DatInmTit = 1
   modatecli_g_int_DatCreTit = 1
   Screen.MousePointer = 0
End Sub

Private Sub cmd_DatCli_Click()
   moddat_g_int_FlgAct = 1
   moddat_g_int_FlgGrb = 2
   
   frm_MntCli_52.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      moddat_g_str_FecNac_Tit = ""
      moddat_g_str_FecNac_Cyg = ""
      moddat_g_int_RegCyg = 0
      moddat_g_int_EstCiv = 0
      moddat_g_int_ComRta = 0
      
      Call gs_LimpiaGrid(grd_Listad(0))
      Call gs_LimpiaGrid(grd_Listad(1))
   
      'Buscar Información del Cliente
      Call fs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(0), 0)
   
      'Verificando que Cliente no haya sido ingresado como Cliente Negativo
      If Not atecli_gf_Buscar_BasNeg(moddat_g_int_TipDoc, moddat_g_str_NumDoc) Then
         Call cmd_Limpia_Click
         Exit Sub
      End If
   
      'Validar que Cliente no tenga una Solicitud de Crédito en Evaluación Como Titular
      If Not atecli_gf_Buscar_SolVig(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 1) Then
         Call cmd_Limpia_Click
         Exit Sub
      End If
      
      'Validar que Cliente no tenga una Solicitud de Crédito en Evaluación Como Cónyuge
      If Not atecli_gf_Buscar_SolVig(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 2) Then
         Call cmd_Limpia_Click
         Exit Sub
      End If
   
      'Buscando Operaciones de Crédito
      Call atecli_gs_Buscar_CreHip(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 1)
      If UBound(modatecli_g_arr_TitOpe) > 0 Then
         MsgBox "El Cliente ya tiene un Crédito Hipotecario registrado.", vbInformation, modgen_g_str_NomPlt
         Call cmd_Limpia_Click
         Exit Sub
      End If
      
      'Buscando Solicitudes Rechazadas de Cliente
      Call atecli_gs_Buscar_SolRec(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 1)
      If UBound(modatecli_g_arr_LisRec) = 0 Then
         cmd_LisRec_Cli.Enabled = False
      End If
      
      'Buscando Información del Cónyuge
      If moddat_g_int_CygTDo > 0 Then
         Call fs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, grd_Listad(1), 1)
         
         'Verificando que Cónyuge no haya sido ingresado como Cliente Negativo
         If Not atecli_gf_Buscar_BasNeg(moddat_g_int_CygTDo, moddat_g_str_CygNDo) Then
            Call cmd_Limpia_Click
            Exit Sub
         End If
      
         'Validar que Cónyuge no tenga una Solicitud de Crédito en Evaluación Como Titular
         If Not atecli_gf_Buscar_SolVig(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 1) Then
            Call cmd_Limpia_Click
            Exit Sub
         End If
         
         'Validar que Cónyuge no tenga una Solicitud de Crédito en Evaluación Como Cónyuge
         If Not atecli_gf_Buscar_SolVig(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 2) Then
            Call cmd_Limpia_Click
            Exit Sub
         End If
         
         'Buscando Operaciones de Crédito
         Call atecli_gs_Buscar_CreHip(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 2)
         
         If UBound(modatecli_g_arr_CygOpe) > 0 Then
            MsgBox "El Cónyuge ya tiene un Crédito Hipotecario registrado.", vbInformation, modgen_g_str_NomPlt
            Call cmd_Limpia_Click
            Exit Sub
         End If
      
         'Buscando Solicitudes Rechazadas del Cónyuge
         Call atecli_gs_Buscar_SolRec(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 2)
         If UBound(modatecli_g_arr_CygRec) = 0 Then
            cmd_LisRec_Cyg.Enabled = False
         End If
      End If
      
      'Buscar Información del Apoderado
      Call fs_DatApo(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(2))
   End If
   'Call gs_SetFocus(cmd_Patrim)
End Sub

Private Sub cmd_DatCre_Click()
   If grd_Listad(0).Rows = 0 Then
      MsgBox "Debe ingresar los datos del cliente.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   frm_SolCre_54.Show 1
End Sub

Private Sub cmd_DatInm_Click()
   If grd_Listad(0).Rows = 0 Then
      MsgBox "Debe ingresar los datos del cliente.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   frm_SolCre_55.Show 1
End Sub

Private Sub cmd_Grabar_Click()
Dim r_str_NumSol     As String
Dim r_dbl_CuoRta     As Double
Dim r_dbl_IngMin     As Double
Dim r_dbl_TipCam     As Double
Dim r_str_CodPrd     As String
Dim r_int_TipEva     As Integer
Dim r_str_CodPry     As String
Dim r_dbl_ValPre     As Double
Dim r_dbl_ValInm     As Double
Dim r_int_PlzPre     As Integer
Dim r_dbl_PorIni     As Double
Dim r_dbl_TasInt     As Double
Dim r_str_Cadena     As String
Dim r_str_Parame     As String
Dim r_rst_Genera     As ADODB.Recordset
   
   'Validacion que compara el nombre y apellido ingresado en prospecto y el de la base de clientes
   If Not ValidaNombreCliente(moddat_g_int_TipDoc, moddat_g_str_NumDoc) Then
      If MsgBox("El nombre del cliente no coincide con el registrado en Prospectos." & vbCrLf & "¿Está seguro de continuar?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
   End If
   
   If moddat_g_int_EstCiv = 2 Or moddat_g_int_EstCiv = 5 Then
      If grd_Listad(1).Rows = 0 Then
         MsgBox "Debe ingresar la información del Cónyuge.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_DatCli)
         Exit Sub
      End If
   End If
   
   If modmip_g_int_PaiRes <> 4028 Then
      If grd_Listad(2).Rows = 0 Then
         MsgBox "Debe ingresar la información del Apoderado.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_DatCli)
         Exit Sub
      End If
   End If
   
   If moddat_g_int_ComRta = 1 And l_int_ActPri_Cyg = 0 Then
      MsgBox "Debe ingresar la Actividad Económica del Cónyuge.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_DatCli)
      Exit Sub
   End If
   
   'Obteniendo Ingreso Mínimo de Parámetro por Producto
   r_dbl_IngMin = 0
   r_dbl_TasInt = 0
   
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "014") Then
      r_dbl_IngMin = moddat_g_arr_Genera(1).Genera_Cantid
   End If
   
   If modatecli_g_int_GastosTit = 1 Then
      MsgBox "Debe ingresar la Información de Inmuebles, Tarjetas, Deudas y Egresos del Cliente.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Patrim)
      Exit Sub
   End If
   If modatecli_g_int_RefereTit = 1 Then
      MsgBox "Debe ingresar la Información de Referencias Personales del Cliente.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Refere)
      Exit Sub
   End If
                           
   If modatecli_g_int_DatInmTit = 1 Then
      MsgBox "Debe ingresar la Información del Inmueble.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_DatInm)
      Exit Sub
   Else
      If modatecli_g_arr_DatInm(1).DatInm_InmIde = 2 Then
         MsgBox "Debe ingresar la Información del Inmueble.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_DatInm)
         Exit Sub
      End If
   End If
   
   If modatecli_g_int_DatCreTit = 1 Then
      MsgBox "Debe ingresar la Información del Crédito.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_DatCre)
      Exit Sub
   End If
      
   If modatecli_g_arr_DatInm(1).DatInm_FlgEst = 1 Then
      If modatecli_g_arr_DatCre(1).DatCre_MtoEst = 0 Then
         MsgBox "Debe ingresar el valor del estacionamiento, en datos del crédito.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_DatCre)
         Exit Sub
      End If
   Else
      If modatecli_g_arr_DatCre(1).DatCre_MtoEst <> 0 Then
         MsgBox "El valor del estacionamiento debe de ser cero, en datos del crédito.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_DatCre)
         Exit Sub
      End If
   End If
   
   If modatecli_g_arr_DatCre(1).DatCre_TipEva = 1 Then
      If moddat_g_dbl_IngDec < r_dbl_IngMin Then
         MsgBox "El Ingreso Declarado es menor al Ingreso Mínimo solicitado para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   ElseIf modatecli_g_arr_DatCre(1).DatCre_TipEva = 2 Then
      r_dbl_CuoRta = 0
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "013") Then
         r_dbl_CuoRta = moddat_g_arr_Genera(1).Genera_Cantid
      End If
      
      If modatecli_g_arr_DatCre(1).DatCre_MonAho = 1 Then
         moddat_g_dbl_IngDec = 1 / r_dbl_CuoRta * 100 * modatecli_g_arr_DatCre(1).DatCre_MtoAho
      Else
         r_dbl_TipCam = moddat_gf_Obtiene_TipCam(1, modatecli_g_arr_DatCre(1).DatCre_MonAho)
         moddat_g_dbl_IngDec = 1 / r_dbl_CuoRta * 100 * modatecli_g_arr_DatCre(1).DatCre_MtoAho * r_dbl_TipCam
      End If
      
      moddat_g_dbl_IngDec = CDbl(Format(moddat_g_dbl_IngDec, "###,###,##0.00"))
   
      If moddat_g_dbl_IngDec < r_dbl_IngMin Then
         MsgBox "La Cuota de Ahorro Mensual no cubre el Ingreso Mínimo solicitado para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   End If
   
'   'obtiene los parámetros de ingreso: Valor del Préstamo, Valor de Inmueble, Palzo del Préstamo, Porcentaje Inicial y Tasa de Interés.
'   Call ff_Obtener_Tasa_Interes(r_str_CodPrd, r_int_TipEva, r_str_CodPry, r_dbl_ValPre, r_dbl_ValInm, r_int_PlzPre, r_dbl_PorIni, r_dbl_TasInt)
'
'   If r_dbl_TasInt = 0 Then
'      MsgBox "Los parámetros que se asignaron son: " & vbLf & vbLf & r_str_Cadena, vbExclamation, modgen_g_str_NomPlt
'      Exit Sub
'   End If
'
'   r_str_Cadena = ""
'   r_str_Cadena = r_str_Cadena & gs_modsec_Genera("Valor del Préstamo: ", 2, " ", 20) & gs_modsec_Genera(" ", 1, " ", 8) & Format(r_dbl_ValPre, "###,###,###,##0.00") & vbLf
'   r_str_Cadena = r_str_Cadena & gs_modsec_Genera("Valor del Inmueble: ", 2, " ", 20) & gs_modsec_Genera(" ", 1, " ", 8) & Format(r_dbl_ValInm, "###,###,###,##0.00") & vbLf
'   r_str_Cadena = r_str_Cadena & gs_modsec_Genera("Plazo Préstamo: ", 2, " ", 15) & gs_modsec_Genera(" ", 1, " ", 13) & r_int_PlzPre & vbLf
'   r_str_Cadena = r_str_Cadena & gs_modsec_Genera("Porcentaje Inicial: ", 2, " ", 15) & gs_modsec_Genera(" ", 1, " ", 10) & r_dbl_PorIni & "%" & vbLf
'   r_str_Cadena = r_str_Cadena & gs_modsec_Genera("Tasa de Interés: ", 2, " ", 15) & gs_modsec_Genera(" ", 1, " ", 12) & r_dbl_TasInt & "%"
'
'   MsgBox "Los parámetros que se asignaron son: " & vbLf & vbLf & r_str_Cadena, vbExclamation, modgen_g_str_NomPlt
      
   If MsgBox("¿Está seguro de grabar la Solicitud de Crédito?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   
   'Generando Número de Solicitud
   r_str_NumSol = ff_Genera_NumSol()

   'Grabando en Maestro de Solicitudes
   If Not ff_Graba_SolMae(r_str_NumSol, r_dbl_TasInt) Then
      Exit Sub
   End If
   
   'Grabando Información de Ingresos - Inmuebles
   If Not ff_Graba_IngInm(r_str_NumSol) Then
      Exit Sub
   End If

   'Grabando Información de Gastos Tarjetas
   If Not ff_Graba_GasTrj(r_str_NumSol) Then
      Exit Sub
   End If

   'Grabando Información de Gastos Deudas
   If Not ff_Graba_GasDeu(r_str_NumSol) Then
      Exit Sub
   End If

   'Grabando Información de Gastos Mensuales
   If Not ff_Graba_GasGas(r_str_NumSol) Then
      Exit Sub
   End If

   'Grabando Información de Referencias
   If Not ff_Graba_Refere(r_str_NumSol) Then
      Exit Sub
   End If

   'Grabando Información de Ejecutivo de Ventas
   If Not ff_Graba_SolEje(r_str_NumSol) Then
      Exit Sub
   End If
   
   'Grabando Lista de Documentos Recibidos
   If Not ff_Graba_SolDoc(r_str_NumSol) Then
      Exit Sub
   End If
   
   'Grabando en Seguimiento
   If Not ff_Graba_Seguim(r_str_NumSol) Then
      Exit Sub
   End If
   
   'Grabando Información de Inmueble si tiene identificado el inmueble
   If modatecli_g_arr_DatInm(1).DatInm_InmIde = 1 Then
      If Not ff_Graba_Inmueb(r_str_NumSol) Then
         Exit Sub
      End If
   End If
   
   
   MsgBox "Los datos fueron registrados correctamente. El Número de Solicitud generado es el: " & Left(r_str_NumSol, 3) & "-" & Mid(r_str_NumSol, 4, 3) & "-" & Mid(r_str_NumSol, 7, 2) & "-" & Right(r_str_NumSol, 4), vbInformation, modgen_g_str_NomPlt
   Call cmd_Limpia_Click
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_Produc)
End Sub

Private Sub cmd_LisRec_Cli_Click()
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   frm_LisRec_01.Show 1
End Sub

Private Sub cmd_LisRec_Cyg_Click()
   moddat_g_str_CygNom = moddat_gf_Buscar_NomCli(moddat_g_int_CygTDo, moddat_g_str_CygTDo)
   frm_LisRec_02.Show 1
End Sub

Private Sub cmd_Patrim_Click()
   If grd_Listad(0).Rows = 0 Then
      MsgBox "Debe ingresar los datos del cliente.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   frm_SolCre_52.Show 1
End Sub

Private Sub cmd_Refere_Click()
   If grd_Listad(0).Rows = 0 Then
      MsgBox "Debe ingresar los datos del cliente.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   frm_SolCre_53.Show 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_SimCre_Click()
   If moddat_gf_Obtiene_TipCam(1, 2) = 0 Then
      MsgBox "No se encuentra registrado el Tipo de Cambio. No puede ingresar a esta opción.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   frm_SimCre_11.Show 1
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
   Dim r_int_Contad     As Integer

   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "230")
   Call moddat_gs_Carga_Produc_Comerc(cmb_Produc, l_arr_Produc, 4)
   Call moddat_gs_Carga_LisIte_Combo(cmb_OfiCom, 1, "518")
   Call moddat_gs_Carga_LisIte_Combo(cmb_ForCon, 1, "519")
   Call moddat_gs_Carga_FerEve(cmb_FerEve)
   
   For r_int_Contad = 0 To 2
      grd_Listad(r_int_Contad).ColWidth(0) = 3000
      grd_Listad(r_int_Contad).ColWidth(1) = 8000
      grd_Listad(r_int_Contad).ColAlignment(0) = flexAlignLeftCenter
      grd_Listad(r_int_Contad).ColAlignment(1) = flexAlignLeftCenter
   Next r_int_Contad
End Sub

Private Sub fs_Limpia()
   Dim r_int_Contad     As Integer
   
   cmb_Produc.ListIndex = -1
   cmb_SubPrd.ListIndex = -1
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   cmb_OfiCom.ListIndex = -1
   cmb_ForCon.ListIndex = -1
   cmb_FerEve.ListIndex = -1
   
   For r_int_Contad = 0 To 2
      Call gs_LimpiaGrid(grd_Listad(r_int_Contad))
   Next r_int_Contad
End Sub

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   cmb_Produc.Enabled = p_Habilita
   cmb_SubPrd.Enabled = p_Habilita
   cmb_TipDoc.Enabled = p_Habilita
   txt_NumDoc.Enabled = p_Habilita
   cmb_OfiCom.Enabled = p_Habilita
   cmb_ForCon.Enabled = p_Habilita
   cmb_FerEve.Enabled = p_Habilita
   grd_Listad(0).Enabled = Not p_Habilita
   grd_Listad(1).Enabled = Not p_Habilita
   grd_Listad(2).Enabled = Not p_Habilita
   cmd_LisRec_Cli.Enabled = Not p_Habilita
   cmd_LisRec_Cyg.Enabled = Not p_Habilita
   cmd_DatCli.Enabled = Not p_Habilita
   cmd_Patrim.Enabled = Not p_Habilita
   cmd_Refere.Enabled = Not p_Habilita
   cmd_DatInm.Enabled = Not p_Habilita
   cmd_DatCre.Enabled = Not p_Habilita
   cmd_Grabar.Enabled = Not p_Habilita
End Sub

Private Function ValidaNombreCliente(ByVal p_TipDoc As String, ByVal p_NumDoc As String) As Boolean
Dim r_str_Parame     As String
Dim r_rst_TipCli     As ADODB.Recordset
Dim r_str_NomPro     As String
Dim r_str_NomCli     As String

   ValidaNombreCliente = False
   r_str_NomPro = ""
   r_str_NomCli = ""
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "SELECT PROCLI_TIPDOC, PROCLI_NUMDOC, PROCLI_APEPAT, PROCLI_APEMAT, PROCLI_NOMBRE "
   r_str_Parame = r_str_Parame & "  FROM CRE_PROCLI "
   r_str_Parame = r_str_Parame & " WHERE PROCLI_TIPDOC = " & p_TipDoc & " "
   r_str_Parame = r_str_Parame & "   AND PROCLI_NUMDOC = " & Trim(p_NumDoc) & " "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_TipCli, 3) Then
      Exit Function
   End If
   
   If r_rst_TipCli.BOF And r_rst_TipCli.EOF Then
      Exit Function
   End If
   
   r_str_NomPro = Trim(r_rst_TipCli!PROCLI_APEPAT) & " " & Trim(r_rst_TipCli!PROCLI_APEMAT) & " " & Trim(r_rst_TipCli!PROCLI_NOMBRE) & " "
   
   r_rst_TipCli.Close
   Set r_rst_TipCli = Nothing
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "SELECT DATGEN_TIPDOC, DATGEN_NUMDOC, DATGEN_APEPAT, DATGEN_APEMAT, DATGEN_NOMBRE "
   r_str_Parame = r_str_Parame & "  FROM CLI_DATGEN "
   r_str_Parame = r_str_Parame & " WHERE DATGEN_TIPDOC = " & p_TipDoc & " "
   r_str_Parame = r_str_Parame & "   AND DATGEN_NUMDOC = " & Trim(p_NumDoc) & " "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_TipCli, 3) Then
      Exit Function
   End If
   
   If r_rst_TipCli.BOF And r_rst_TipCli.EOF Then
      Exit Function
   End If
   
   r_str_NomCli = Trim(r_rst_TipCli!DATGEN_APEPAT) & " " & Trim(r_rst_TipCli!DATGEN_APEMAT) & " " & Trim(r_rst_TipCli!DATGEN_NOMBRE) & " "
   
   r_rst_TipCli.Close
   Set r_rst_TipCli = Nothing
   
   If r_str_NomPro = r_str_NomCli Then
      ValidaNombreCliente = True
   End If
End Function

Private Sub grd_Listad_SelChange(Index As Integer)
   If grd_Listad(Index).Rows > 2 Then
      grd_Listad(Index).RowSel = grd_Listad(Index).Row
   End If
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)

      If KeyAscii = 13 Then
         Call gs_SetFocus(cmb_OfiCom)
      End If
   
      If cmb_TipDoc.ListIndex > -1 Then
         Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
            Case 1: KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case 7: KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case Else:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
         End Select
      Else
         KeyAscii = 0
      End If
'   End If
End Sub

Private Sub fs_DatCli(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, p_Grid As MSFlexGrid, p_Indice As Integer)
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & p_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      p_Grid.Redraw = False
      g_rst_Princi.MoveFirst
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Documento de Identidad"
      p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!DatGen_TipDoc)) & " - " & Trim(g_rst_Princi!DATGEN_NUMDOC & "")
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Apellidos y Nombres"
      p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!DATGEN_APEPAT) & " " & Trim(g_rst_Princi!DATGEN_APEMAT) & IIf(Len(Trim(g_rst_Princi!DatGen_ApeCas)) > 0, " DE " & Trim(g_rst_Princi!DatGen_ApeCas), "") & " " & Trim(g_rst_Princi!DATGEN_NOMBRE)
      
      If g_rst_Princi!DatGen_FLGDOA = 1 Then
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Documento Adicional de Identidad"
         p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!DatGen_FLGDOA)) & IIf(g_rst_Princi!DatGen_FLGDOA = 1, " ( " & moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!DatGen_TIPDOA)) & " - " & Trim(g_rst_Princi!DatGen_NUMDOA) & ")", "")
      End If
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Sexo"
      p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("207", CStr(g_rst_Princi!DatGen_CodSex))
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Fecha de Nacimiento"
      p_Grid.Col = 1:                     p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_NACFEC))
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Nacionalidad"
      p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("500", Trim(g_rst_Princi!DATGEN_NACPAI))
   
      If Trim(g_rst_Princi!DATGEN_NACPAI) = "004028" Then
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Lugar de Nacimiento"
         p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_NACLUG, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_NACLUG, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DATGEN_NACLUG))
      End If
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Estado Civil"
      p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("205", CStr(g_rst_Princi!DATGEN_ESTCIV)) & IIf(g_rst_Princi!DATGEN_ESTCIV = 2, " / " & moddat_gf_Consulta_ParDes("206", g_rst_Princi!DATGEN_REGCYG), "")
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Nivel de Estudios"
      p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("209", CStr(g_rst_Princi!DatGen_NivEst))

      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Profesión"
      p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("501", CStr(g_rst_Princi!DatGen_Profes))
      
      If p_Indice = 0 Then
         If g_rst_Princi!DatGen_DepEco > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Nro. Dependientes Económicos"
            p_Grid.Col = 1:                  p_Grid.Text = CStr(g_rst_Princi!DatGen_DepEco)
         
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Edades"
            p_Grid.Col = 1:                  p_Grid.Text = IIf(g_rst_Princi!DatGen_EDAD01 > 0, CStr(g_rst_Princi!DatGen_EDAD01), "") & IIf(g_rst_Princi!DatGen_EDAD02 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD02), "") & IIf(g_rst_Princi!DatGen_EDAD03 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD03), "") & IIf(g_rst_Princi!DatGen_EDAD04 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD04), "") & IIf(g_rst_Princi!DatGen_EDAD05 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD05), "")
         End If
      End If
      
      If p_Indice = 0 Then
         p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "País Residencia"
         p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("500", CStr(g_rst_Princi!DATGEN_PAIRES))
         
         If Trim(g_rst_Princi!DATGEN_PAIRES & "") = "004028" Then
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Domicilio"
            p_Grid.Col = 1
            p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!DatGen_TipVia)) & " " & Trim(g_rst_Princi!DatGen_NomVia) & " " & Trim(g_rst_Princi!DatGen_Numero) & _
                                        IIf(Len(Trim(g_rst_Princi!DATGEN_INTDPT)) > 0, " (" & Trim(g_rst_Princi!DATGEN_INTDPT) & ")", "") & _
                                        IIf(Len(Trim(g_rst_Princi!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!DatGen_TipZon)) & " " & Trim(g_rst_Princi!DatGen_NomZon), "")
            
            If Len(Trim(g_rst_Princi!DATGEN_REFERE & "")) > 0 Then
               p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
               p_Grid.Col = 0:                  p_Grid.Text = "Referencia"
               p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!DATGEN_REFERE & "")
            End If
            
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Departamento / Provincia / Distrito"
            p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DatGen_Ubigeo))
                                        
            moddat_g_str_UbiGeo = Left(Format(g_rst_Princi!DatGen_Ubigeo, "000000"), 4)
         Else
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Domicilio"
            p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!DATGEN_EXTDIR & "")
         
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Ciudad"
            p_Grid.Col = 1:                  p_Grid.Text = modmip_gf_Consulta_CiuExt(Format(g_rst_Princi!DATGEN_PAIRES, "000000"), Trim(g_rst_Princi!DATGEN_EXTCIU))
         
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Código Postal"
            p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!DATGEN_EXTCPO & "")
         End If
      
         If Len(Trim(g_rst_Princi!DatGen_Telefo & "")) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                     p_Grid.Text = "Teléfono Domicilio"
            p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!DatGen_Telefo & "")
         End If
      
         If Len(Trim(g_rst_Princi!DATGEN_NUMCEL & "")) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                     p_Grid.Text = "Teléfono Celular"
            p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!DATGEN_NUMCEL & "")
         End If
         
         If Len(Trim(g_rst_Princi!DatGen_DirEle & "")) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "E-mail"
            p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!DatGen_DirEle & "")
            
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Autorización Envío"
            p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!DATGEN_AUTENV))
         End If
      End If
      
      If g_rst_Princi!DATGEN_TDOVIN > 0 Then
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Vinculado"
         
         p_Grid.Col = 1
         
         If g_rst_Princi!DATGEN_TIPVIN = 1 Then
            p_Grid.Text = "TRABAJADOR"
         ElseIf g_rst_Princi!DATGEN_TIPVIN = 2 Or g_rst_Princi!DATGEN_TIPVIN = 3 Then
            p_Grid.Text = "VINCULADO A TRABAJADOR (" & modmip_gf_Consulta_NomTra(g_rst_Princi!DATGEN_TDOVIN, Trim(g_rst_Princi!DATGEN_NDOVIN)) & ")"
         ElseIf g_rst_Princi!DATGEN_TIPVIN = 4 Then
            p_Grid.Text = "FUNCIONARIO"
         ElseIf g_rst_Princi!DATGEN_TIPVIN = 5 Then
            p_Grid.Text = "VINCULADO A FUNCIONARIO (" & modmip_gf_Consulta_NomOtrFun(g_rst_Princi!DATGEN_TDOVIN, Trim(g_rst_Princi!DATGEN_NDOVIN)) & ")"
         End If
      End If
      
      If g_rst_Princi!DATGEN_TDOACC > 0 Then
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Accionista"
         p_Grid.Col = 1
         
         If g_rst_Princi!DATGEN_ACCVIN = 1 Then
            p_Grid.Text = "ACCIONISTA"
         ElseIf g_rst_Princi!DATGEN_ACCVIN = 2 Then
            p_Grid.Text = "VINCULADO A ACCIONISTA (" & modmip_gf_Consulta_NomAcc(g_rst_Princi!DATGEN_TDOACC, Trim(g_rst_Princi!DATGEN_NDOACC)) & ")"
         End If
      End If
      
      modmip_g_int_PaiRes = CInt(g_rst_Princi!DATGEN_PAIRES)
      moddat_g_int_EstCiv = g_rst_Princi!DATGEN_ESTCIV
      moddat_g_int_RegCyg = g_rst_Princi!DATGEN_REGCYG
      
      If p_Indice = 0 Then
         moddat_g_str_FecNac_Tit = gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_NACFEC))
      Else
         moddat_g_str_FecNac_Cyg = gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_NACFEC))
      End If
      
      p_Grid.Redraw = True
      Call gs_UbiIniGrid(p_Grid)
   End If
   
   If p_Indice = 1 Then
      moddat_g_int_ComRta = 0
      
      If g_rst_Princi!DATGEN_ACTECO = 1 Then
         moddat_g_int_ComRta = 1
         
         l_int_ActPri_Cyg = 0
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call fs_ActEco(p_TipDoc, p_NumDoc, 1, p_Indice, p_Grid)
   Call fs_ActEco(p_TipDoc, p_NumDoc, 2, p_Indice, p_Grid)
End Sub

Private Sub fs_ActEco(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_OrdAct As Integer, ByVal p_Indice As Integer, p_Grid As MSFlexGrid)
   g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
   g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & p_NumDoc & "' AND "
   g_str_Parame = g_str_Parame & "ACTECO_ORDACT = " & CStr(p_OrdAct)

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      If p_Indice = 1 Then
         l_int_ActPri_Cyg = 1
      End If
      
      p_Grid.Redraw = False
      p_Grid.Rows = p_Grid.Rows + 2:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Ocupación " & Left(moddat_gf_Consulta_ParDes("007", CStr(g_rst_Princi!ActEco_OrdAct)), 1) & Mid(LCase(moddat_gf_Consulta_ParDes("007", CStr(g_rst_Princi!ActEco_OrdAct))), 2)
      p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("008", g_rst_Princi!ACTECO_CODACT)
      
      Select Case g_rst_Princi!ACTECO_CODACT
         Case 11: Call fs_ActEco_Dep(p_Grid)
         Case 21: Call fs_ActEco_Ind(p_Grid)
         Case 31: Call fs_ActEco_Com(p_Grid)
         Case 41: Call fs_ActEco_Acc(p_Grid)
         Case 51: Call fs_ActEco_Ren(p_Grid)
         Case 61: Call fs_ActEco_Otr(p_Grid)
      End Select
      
      p_Grid.Redraw = True
      Call gs_UbiIniGrid(p_Grid)
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_ActEco_Dep(p_Grid As MSFlexGrid)
   p_Grid.Rows = p_Grid.Rows + 1
   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0
   p_Grid.Text = "Documento Identidad Empleador"

   p_Grid.Col = 1
   p_Grid.Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Dep_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Dep_NumDoc & "")

   p_Grid.Rows = p_Grid.Rows + 1
   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0
   p_Grid.Text = "Situación como Trabajador"

   p_Grid.Col = 1
   p_Grid.Text = moddat_gf_Consulta_ParDes("235", g_rst_Princi!ActEco_Dep_SitTra)

   'Buscar en Maestro de Empresas
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Dep_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Dep_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst

      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Razón Social"
      p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Nombre Comercial"
      p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "CIIU"
      p_Grid.Col = 1:                  p_Grid.Text = g_rst_Genera!DATGEN_CODCIU & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Genera!DATGEN_CODCIU))
      
      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Tipo de Oficina"
      p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("234", CStr(g_rst_Princi!ActEco_Dep_TipOfi))
      
      If g_rst_Princi!ActEco_Dep_TipOfi = 1 Then
         If modmip_g_int_PaiRes = 4028 Then
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Dirección"
            p_Grid.Col = 1
            p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Genera!DatGen_TipVia)) & " " & Trim(g_rst_Genera!DatGen_NomVia) & " " & Trim(g_rst_Genera!DatGen_numVia) & _
                                        IIf(Len(Trim(g_rst_Genera!DATGEN_INTDPT)) > 0, " (" & Trim(g_rst_Genera!DATGEN_INTDPT) & ")", "") & _
                                        IIf(Len(Trim(g_rst_Genera!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Genera!DatGen_TipZon)) & " " & Trim(g_rst_Genera!DatGen_NomZon), "")
   
            If Len(Trim(g_rst_Genera!DATGEN_REFERE)) > 0 Then
               p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
               p_Grid.Col = 0:                     p_Grid.Text = "Referencia"
               p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Genera!DATGEN_REFERE & "")
            End If
   
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Departamento / Provincia / Distrito"
            p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo & "", 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo & "", 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Genera!DatGen_Ubigeo & ""))
         Else
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Domicilio"
            p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_EXTDIR & "")
         
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Ciudad"
            p_Grid.Col = 1:                        p_Grid.Text = modmip_gf_Consulta_CiuExt(g_rst_Genera!DATGEN_PAIRES, Trim(g_rst_Genera!DATGEN_EXTCIU))
         
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Código Postal"
            p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_EXTCPO & "")
         End If
      
         p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                           p_Grid.Text = "Teléfono(s)"
         p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
      
         If Len(Trim(g_rst_Genera!DATGEN_NUMFAX & "")) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                           p_Grid.Text = "Fax"
            p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Genera!DATGEN_NUMFAX & "")
         End If
      Else
         If modmip_g_int_PaiRes = 4028 Then
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Dirección"
            p_Grid.Col = 1
            p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Dep_TipVia)) & " " & Trim(g_rst_Princi!ActEco_Dep_NomVia) & " " & Trim(g_rst_Princi!ActEco_Dep_NumVia) & _
                                        IIf(Len(Trim(g_rst_Princi!ActEco_Dep_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Dep_IntDpt) & ")", "") & _
                                        IIf(Len(Trim(g_rst_Princi!ActEco_Dep_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Dep_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Dep_NomZon), "")
      
            If Len(Trim(g_rst_Princi!ActEco_Dep_Refere & "")) > 0 Then
               p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
               p_Grid.Col = 0:                     p_Grid.Text = "Referencia"
               p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!ActEco_Dep_Refere & "")
            End If
      
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Departamento / Provincia / Distrito"
      
            p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Dep_UbiGeo, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Dep_UbiGeo, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Dep_UbiGeo))
         Else
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Domicilio"
            p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ACTECO_DEP_EXTDIR & "")
         
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Ciudad"
            p_Grid.Col = 1:                        p_Grid.Text = modmip_gf_Consulta_CiuExt(Format(modmip_g_int_PaiRes, "000000"), Trim(g_rst_Princi!ACTECO_DEP_EXTCIU))
         
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Código Postal"
            p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ACTECO_DEP_EXTCPO & "")
         End If
         
         p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                           p_Grid.Text = "Teléfono(s)"
         p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Princi!ActEco_Dep_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Dep_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Dep_Telef2 & ""), "")
         
         If Len(Trim(g_rst_Princi!ActEco_Dep_NumFax & "")) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                           p_Grid.Text = "Fax"
            p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Princi!ActEco_Dep_NumFax & "")
         End If
      End If
      
      If Len(Trim(g_rst_Genera!DATGEN_TELERH & "")) > 0 Then
         p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                        p_Grid.Text = "Teléfono RR.HH"
         p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_TELERH & "")
      End If
   
      If Len(Trim(g_rst_Genera!DATGEN_ANEXRH & "")) > 0 Then
         p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                        p_Grid.Text = "Anexo RR.HH"
         p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_ANEXRH & "")
      End If
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing

   p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                           p_Grid.Text = "Moneda de Ingresos"
   p_Grid.Col = 1:                           p_Grid.Text = moddat_gf_Consulta_ParDes("113", CStr(g_rst_Princi!ActEco_Dep_MonIng))

   p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                           p_Grid.Text = "Ingreso Neto"
   p_Grid.Col = 1:                           p_Grid.Text = gf_FormatoNumero(g_rst_Princi!ActEco_Dep_IngNet, 15, 2)

   p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                           p_Grid.Text = "Fecha de Ingreso"
   p_Grid.Col = 1:                           p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecIng))

   p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                           p_Grid.Text = "Cargo"
   p_Grid.Col = 1:                           p_Grid.Text = IIf(g_rst_Princi!ActEco_Dep_CodCar = "999999", Trim(g_rst_Princi!ActEco_Dep_NomCar & ""), moddat_gf_Consulta_ParDes("503", g_rst_Princi!ActEco_Dep_CodCar))

   If Len(Trim(g_rst_Princi!ActEco_Dep_NumAnx & "")) > 0 Then
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Anexo"
      p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ActEco_Dep_NumAnx & "")
   End If
   
   If Len(Trim(g_rst_Princi!ActEco_Dep_TelDir & "")) > 0 Then
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Teléfono Directo"
      p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ActEco_Dep_TelDir & "")
   End If
   
   moddat_g_dbl_IngDec = moddat_g_dbl_IngDec + g_rst_Princi!ActEco_Dep_IngNet
End Sub

Private Sub fs_ActEco_Ind(p_Grid As MSFlexGrid)
   p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                     p_Grid.Text = "Documento Identidad"
   p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Ind_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Ind_NumDoc & "")

   If modmip_g_int_PaiRes = 4028 Then
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Dirección"
   
      p_Grid.Col = 1
      p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Ind_TipVia)) & " " & Trim(g_rst_Princi!ActEco_Ind_NomVia) & " " & Trim(g_rst_Princi!ActEco_Ind_NumVia) & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Ind_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Ind_IntDpt) & ")", "") & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Ind_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Ind_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Ind_NomZon), "")
   
      If Len(Trim(g_rst_Princi!ActEco_Ind_Refere & "")) > 0 Then
         p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "Referencia"
         p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!ActEco_Ind_Refere & "")
      End If
   
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Indartamento / Provincia / Distrito"
      p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ind_UbiGeo, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ind_UbiGeo, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Ind_UbiGeo))
   Else
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Domicilio"
      p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ACTECO_IND_EXTDIR & "")
   
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Ciudad"
      p_Grid.Col = 1:                        p_Grid.Text = modmip_gf_Consulta_CiuExt(Format(modmip_g_int_PaiRes, "000000"), Trim(g_rst_Princi!ACTECO_IND_EXTCIU))
   
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Código Postal"
      p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ACTECO_IND_EXTCPO & "")
   End If
   
   p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                           p_Grid.Text = "Teléfono(s)"
   p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Princi!ActEco_Ind_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ind_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ind_Telef2 & ""), "")
   
   If Len(Trim(g_rst_Princi!ActEco_Ind_NumFax & "")) > 0 Then
      p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                           p_Grid.Text = "Fax"
      p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Princi!ActEco_Ind_NumFax & "")
   End If

   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "CIIU"
   p_Grid.Col = 1:                        p_Grid.Text = g_rst_Princi!ActEco_Ind_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Ind_CodCiu))

   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Moneda de Ingresos"
   p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("113", CStr(g_rst_Princi!ActEco_ind_MonIng))

   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Ingreso Neto"
   p_Grid.Col = 1:                        p_Grid.Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ind_IngNet, 15, 2)
   
   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Fecha de Inicio de Actividades"
   p_Grid.Col = 1:                        p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ind_IniAct))
   
   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Contrato de Locación de Servicios"
   p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!ActEco_Ind_ConLoc))
   
   If g_rst_Princi!ActEco_Ind_ConLoc = 1 Then
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Documento Identidad Empleador"
      p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Ind_TipDoc_Emp) & " - " & Trim(g_rst_Princi!ActEco_Ind_NumDoc_Emp & "")
      
      'Buscar en Maestro de Empresas
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Ind_TipDoc_Emp) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Ind_NumDoc_Emp & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Razón Social"
         p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")
   
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Nombre Comercial"
         p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")
   
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "CIIU"
         p_Grid.Col = 1:                  p_Grid.Text = g_rst_Genera!DATGEN_CODCIU & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Genera!DATGEN_CODCIU))
         
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Tipo de Oficina"
         p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("234", CStr(g_rst_Princi!ActEco_Dep_TipOfi))
         
         If modmip_g_int_PaiRes = 4028 Then
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Dirección"
            p_Grid.Col = 1
            p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Genera!DatGen_TipVia)) & " " & Trim(g_rst_Genera!DatGen_NomVia) & " " & Trim(g_rst_Genera!DatGen_numVia) & _
                                        IIf(Len(Trim(g_rst_Genera!DATGEN_INTDPT)) > 0, " (" & Trim(g_rst_Genera!DATGEN_INTDPT) & ")", "") & _
                                        IIf(Len(Trim(g_rst_Genera!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Genera!DatGen_TipZon)) & " " & Trim(g_rst_Genera!DatGen_NomZon), "")
   
            If Len(Trim(g_rst_Genera!DATGEN_REFERE)) > 0 Then
               p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
               p_Grid.Col = 0:                     p_Grid.Text = "Referencia"
               p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Genera!DATGEN_REFERE & "")
            End If
   
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Departamento / Provincia / Distrito"
            p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Genera!DatGen_Ubigeo))
         Else
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Domicilio"
            p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_EXTDIR & "")
         
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Ciudad"
            p_Grid.Col = 1:                        p_Grid.Text = modmip_gf_Consulta_CiuExt(g_rst_Genera!DATGEN_PAIRES, Trim(g_rst_Genera!DATGEN_EXTCIU))
         
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Código Postal"
            p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_EXTCPO & "")
         End If
         
         p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                           p_Grid.Text = "Teléfono(s)"
         p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
      
         If Len(Trim(g_rst_Genera!DATGEN_NUMFAX & "")) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Fax"
            p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_NUMFAX & "")
         End If
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   
      p_Grid.Rows = p_Grid.Rows + 1:               p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                              p_Grid.Text = "Fecha de Ingreso"
      p_Grid.Col = 1:                              p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ind_FecIng_Emp))
   
      p_Grid.Rows = p_Grid.Rows + 1:               p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                              p_Grid.Text = "Cargo"
      p_Grid.Col = 1:                              p_Grid.Text = IIf(g_rst_Princi!ActEco_Ind_CodCar = "999999", Trim(g_rst_Princi!ActEco_Ind_NomCar & ""), moddat_gf_Consulta_ParDes("503", g_rst_Princi!ActEco_Ind_CodCar))
   End If

   moddat_g_dbl_IngDec = moddat_g_dbl_IngDec + g_rst_Princi!ActEco_Ind_IngNet
End Sub

Private Sub fs_ActEco_Com(p_Grid As MSFlexGrid)
   p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                     p_Grid.Text = "Documento Identidad"
   p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Com_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Com_NumDoc & "")

   p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                     p_Grid.Text = "Razón Social"
   p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!ActEco_Com_RazSoc & "")

   p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                     p_Grid.Text = "Nombre Comercial"
   p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!ActEco_Com_NomCom & "")
   
   If modmip_g_int_PaiRes = 4028 Then
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Dirección"
      p_Grid.Col = 1
      p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Com_TipVia)) & " " & Trim(g_rst_Princi!ActEco_Com_NomVia) & " " & Trim(g_rst_Princi!ActEco_Com_NumVia) & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Com_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Com_IntDpt) & ")", "") & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Com_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Com_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Com_NomZon), "")
   
      If Len(Trim(g_rst_Princi!ActEco_Com_Refere & "")) > 0 Then
         p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "Referencia"
         p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!ActEco_Com_Refere & "")
      End If
   
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Departamento / Provincia / Distrito"
      p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Com_UbiGeo, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Com_UbiGeo, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Com_UbiGeo))
   Else
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Dirección"
      p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ACTECO_COM_EXTDIR & "")
   
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Ciudad"
      p_Grid.Col = 1:                        p_Grid.Text = modmip_gf_Consulta_CiuExt(Format(modmip_g_int_PaiRes, "000000"), Trim(g_rst_Princi!ACTECO_COM_EXTCIU))
   
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Código Postal"
      p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ACTECO_COM_EXTCPO & "")
   End If
   
   p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                  p_Grid.Text = "Teléfono(s)"
   p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!ActEco_Com_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Com_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Com_Telef2 & ""), "")
   
   If Len(Trim(g_rst_Princi!ActEco_Com_NumFax & "")) > 0 Then
      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Fax"
      p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!ActEco_Com_NumFax & "")
   End If

   p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                  p_Grid.Text = "CIIU"
   p_Grid.Col = 1:                  p_Grid.Text = g_rst_Princi!ActEco_Com_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Com_CodCiu))

   p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                  p_Grid.Text = "Giro Comercial"
   p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_GirCom(g_rst_Princi!ActEco_Com_GirCom)

   p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                  p_Grid.Text = "Moneda de Ingresos"
   p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("113", CStr(g_rst_Princi!ActEco_com_MonIng))
   
   p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                  p_Grid.Text = "Ingreso Neto"
   p_Grid.Col = 1:                  p_Grid.Text = gf_FormatoNumero(g_rst_Princi!ActEco_Com_IngNet, 15, 2)

   p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                  p_Grid.Text = "Fecha de Inicio de Operaciones"
   p_Grid.Col = 1:                  p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Com_IniOpe))
   
   p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                  p_Grid.Text = "Cargo"
   p_Grid.Col = 1:                  p_Grid.Text = IIf(g_rst_Princi!ActEco_Com_CodCar = "999999", Trim(g_rst_Princi!ActEco_Com_NomCar & ""), moddat_gf_Consulta_ParDes("503", g_rst_Princi!ActEco_Com_CodCar))
   
   moddat_g_dbl_IngDec = moddat_g_dbl_IngDec + g_rst_Princi!ActEco_Com_IngNet
End Sub

Private Sub fs_ActEco_Acc(p_Grid As MSFlexGrid)
   p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                  p_Grid.Text = "Documento Identidad"
   p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Acc_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Acc_NumDoc & "")

   'Buscar en Maestro de Empresas
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Acc_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Acc_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Razón Social"
      p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Nombre Comercial"
      p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "CIIU"
      p_Grid.Col = 1:                  p_Grid.Text = g_rst_Genera!DATGEN_CODCIU & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Genera!DATGEN_CODCIU))
      
      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Tipo de Oficina"
      p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("234", CStr(g_rst_Princi!ActEco_Dep_TipOfi))
      
      If modmip_g_int_PaiRes = 4028 Then
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Dirección"
         p_Grid.Col = 1
         p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Genera!DatGen_TipVia)) & " " & Trim(g_rst_Genera!DatGen_NomVia) & " " & Trim(g_rst_Genera!DatGen_numVia) & _
                                     IIf(Len(Trim(g_rst_Genera!DATGEN_INTDPT)) > 0, " (" & Trim(g_rst_Genera!DATGEN_INTDPT) & ")", "") & _
                                     IIf(Len(Trim(g_rst_Genera!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Genera!DatGen_TipZon)) & " " & Trim(g_rst_Genera!DatGen_NomZon), "")

         If Len(Trim(g_rst_Genera!DATGEN_REFERE)) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                     p_Grid.Text = "Referencia"
            p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Genera!DATGEN_REFERE & "")
         End If

         p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                        p_Grid.Text = "Departamento / Provincia / Distrito"
         p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Genera!DatGen_Ubigeo))
      Else
         p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                        p_Grid.Text = "Domicilio"
         p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_EXTDIR & "")
      
         p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                        p_Grid.Text = "Ciudad"
         p_Grid.Col = 1:                        p_Grid.Text = modmip_gf_Consulta_CiuExt(g_rst_Genera!DATGEN_PAIRES, Trim(g_rst_Genera!DATGEN_EXTCIU))
      
         p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                        p_Grid.Text = "Código Postal"
         p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_EXTCPO & "")
      End If
      
      p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                           p_Grid.Text = "Teléfono(s)"
      p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
   
      If Len(Trim(g_rst_Genera!DATGEN_NUMFAX & "")) > 0 Then
         p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                        p_Grid.Text = "Fax"
         p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_NUMFAX & "")
      End If
   End If

   g_rst_Genera.Close
   Set g_rst_Genera = Nothing

   p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                     p_Grid.Text = "Moneda de Ingresos"
   p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("113", CStr(g_rst_Princi!ActEco_acc_MonIng))

   p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                     p_Grid.Text = "Ingreso Neto"
   p_Grid.Col = 1:                     p_Grid.Text = gf_FormatoNumero(g_rst_Princi!ActEco_Acc_IngNet, 15, 2)
   
   p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                     p_Grid.Text = "Fecha de Antigüedad"
   p_Grid.Col = 1:                     p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Acc_FecAnt))
   
   moddat_g_dbl_IngDec = moddat_g_dbl_IngDec + g_rst_Princi!ActEco_Acc_IngNet
End Sub

Private Sub fs_ActEco_Ren(p_Grid As MSFlexGrid)
   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Moneda de Ingresos"
   p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("113", CStr(g_rst_Princi!ActEco_ren_MonIng))
   
   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Ingreso Neto"
   p_Grid.Col = 1:                        p_Grid.Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ren_IngNet, 15, 2)

   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Dirección de Propiedad 01"
   p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ActEco_Ren_Direc1 & "")

   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Nombre de Arrendatario"
   p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ActEco_Ren_NomAr1 & "")
   
   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Fecha de Inicio de Alquiler"
   p_Grid.Col = 1:                        p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ren_IniAl1))
   
   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Teléfono(s)"
   p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ActEco_Ren_Tele11 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ren_Tele21 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ren_Tele21 & ""), "")

   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Alquiler Mensual"
   p_Grid.Col = 1:                        p_Grid.Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ren_AlqMe1, 15, 2)
   
   If g_rst_Princi!ActEco_Ren_SegPro = 1 Then
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Dirección de Propiedad 02"
      p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!ActEco_Ren_Direc2 & "")

      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Nombre de Arrendatario"
      p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!ActEco_Ren_NomAr2 & "")
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Fecha de Inicio de Alquiler"
      p_Grid.Col = 1:                     p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ren_IniAl2))
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Teléfono(s)"
      p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!ActEco_Ren_Tele12 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ren_Tele22 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ren_Tele22 & ""), "")

      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Alquiler Mensual"
      p_Grid.Col = 1:                     p_Grid.Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ren_AlqMe2, 15, 2)
   End If

   moddat_g_dbl_IngDec = moddat_g_dbl_IngDec + g_rst_Princi!ActEco_Ren_IngNet
End Sub

Private Sub fs_ActEco_Otr(p_Grid As MSFlexGrid)
   p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                           p_Grid.Text = "Moneda de Ingresos"
   p_Grid.Col = 1:                           p_Grid.Text = moddat_gf_Consulta_ParDes("113", CStr(g_rst_Princi!ActEco_otr_MonIng))
   
   p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                           p_Grid.Text = "Ingreso Neto"
   p_Grid.Col = 1:                           p_Grid.Text = gf_FormatoNumero(g_rst_Princi!ActEco_Otr_IngNet, 15, 2)

   p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                           p_Grid.Text = "Actividad"
   p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Princi!ActEco_Otr_Activi & "")

   p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                           p_Grid.Text = "CIIU"
   p_Grid.Col = 1:                           p_Grid.Text = g_rst_Princi!ActEco_Otr_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Otr_CodCiu))
   
   p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                           p_Grid.Text = "Observaciones"
   p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Princi!ActEco_Otr_Observ & "")
   
   moddat_g_dbl_IngDec = moddat_g_dbl_IngDec + g_rst_Princi!ActEco_Otr_IngNet
End Sub

Private Sub fs_DatApo(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, p_Grid As MSFlexGrid)
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & p_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      If g_rst_Princi!DATGEN_APOTDO > 0 Then
         p_Grid.Redraw = False
         g_rst_Princi.MoveFirst
         
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Documento de Identidad"
         p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!DATGEN_APOTDO)) & " - " & Trim(g_rst_Princi!DATGEN_APONDO & "")
         
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Apellidos y Nombres"
         p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!DATGEN_APOAPP) & " " & Trim(g_rst_Princi!DATGEN_APOAPM) & " " & Trim(g_rst_Princi!DATGEN_APONOM)
         
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Sexo"
         p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("207", CStr(g_rst_Princi!DATGEN_APOSEX))
         
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Fecha de Nacimiento"
         p_Grid.Col = 1:                  p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_APOFNC))
         
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Nacionalidad"
         p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("500", Trim(g_rst_Princi!DATGEN_APONAC))
      
         If Trim(g_rst_Princi!DATGEN_APONAC) = "004028" Then
            p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                     p_Grid.Text = "Lugar de Nacimiento"
            p_Grid.Col = 1
            p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_APOLNC, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_APOLNC, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DATGEN_APOLNC))
         End If
         
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Estado Civil"
         p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("205", CStr(g_rst_Princi!DATGEN_APOECV))
         
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Nivel de Estudios"
         p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("209", CStr(g_rst_Princi!DATGEN_APOEST))
   
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Profesión"
         p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("501", CStr(g_rst_Princi!DATGEN_APOPRF))
         
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Domicilio"
         p_Grid.Col = 1
         p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!DatGen_TipVia)) & " " & Trim(g_rst_Princi!DatGen_NomVia) & " " & Trim(g_rst_Princi!DatGen_Numero) & _
                                     IIf(Len(Trim(g_rst_Princi!DATGEN_INTDPT)) > 0, " (" & Trim(g_rst_Princi!DATGEN_INTDPT) & ")", "") & _
                                     IIf(Len(Trim(g_rst_Princi!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!DatGen_TipZon)) & " " & Trim(g_rst_Princi!DatGen_NomZon), "")
         
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Referencia"
         p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!DATGEN_REFERE & "")
         
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Departamento / Provincia / Distrito"
         p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(" " & g_rst_Princi!DatGen_Ubigeo, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(" " & g_rst_Princi!DatGen_Ubigeo))
                                           
      
         If Len(Trim(g_rst_Princi!DATGEN_APOTEL & "")) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                     p_Grid.Text = "Teléfono Domicilio"
            p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!DATGEN_APOTEL & "")
         End If
      
         If Len(Trim(g_rst_Princi!DATGEN_APOCEL & "")) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                     p_Grid.Text = "Teléfono Celular"
            p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!DATGEN_APOCEL & "")
         End If
            
         If Len(Trim(g_rst_Princi!DatGen_DirEle & "")) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "E-mail"
            p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!DATGEN_APOCOR & "")
            
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Autorización Envío"
            p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!DATGEN_APOAEN))
         End If
         
         p_Grid.Redraw = True
         Call gs_UbiIniGrid(p_Grid)
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Function ff_Obtener_Tasa_Interes(ByRef p_CodPrd As String, ByRef p_TipEva As Integer, ByRef p_CodPry As String, ByRef p_ValPre As Double, _
                                 ByRef p_ValInm As Double, ByRef p_PlzPre As Integer, ByRef p_PorIni As Double, ByRef p_TasInt As Double)
Dim r_dbl_TasInt     As Double
Dim r_dbl_PorIni     As Double
Dim r_int_PlzPre     As Integer
Dim r_int_TipBon     As Integer
  
   r_dbl_TasInt = 0
   r_int_PlzPre = CStr(modatecli_g_arr_DatCre(1).DatCre_PlaAno) * 12
   
   If moddat_g_str_CodPrd = "023" Then
      r_dbl_PorIni = Format((CDbl(modatecli_g_arr_DatCre(1).DatCre_ApoPro) + CDbl(modatecli_g_arr_DatCre(1).DatCre_MtoAFP) + CDbl(modatecli_g_arr_DatCre(1).DatCre_FmvBbp) + CDbl(modatecli_g_arr_DatCre(1).DatCre_MefPbp) + CDbl(modatecli_g_arr_DatCre(1).DatCre_MtoBMS)) / CDbl(modatecli_g_arr_DatCre(1).DatCre_ComVta) * 100, "##0.0000")
   Else
      r_dbl_PorIni = Format((CDbl(modatecli_g_arr_DatCre(1).DatCre_ApoPro) + CDbl(modatecli_g_arr_DatCre(1).DatCre_MtoAFP)) / CDbl(modatecli_g_arr_DatCre(1).DatCre_ComVta) * 100, "##0.0000")
   End If
   '0.- No asignado
   '1.- Sin Bono
   '2.- Bono Verde
   
   If CDbl(modatecli_g_arr_DatCre(1).DatCre_MtoBMS) > 0 Then
      r_int_TipBon = 2
   Else
      r_int_TipBon = 0
   End If
   
   'ITERACION 11
   If moddat_gf_Consulta_ParTasInt(l_arr_Parame, moddat_g_str_CodPrd, moddat_g_str_Codigo, CDbl(modatecli_g_arr_DatCre(1).DatCre_MtoPre), CDbl(modatecli_g_arr_DatCre(1).DatCre_MtoPre), CDbl(modatecli_g_arr_DatCre(1).DatCre_ComVta), CDbl(modatecli_g_arr_DatCre(1).DatCre_ComVta), r_int_PlzPre, r_int_PlzPre, r_dbl_PorIni, r_dbl_PorIni, r_int_TipBon) Then
      r_dbl_TasInt = l_arr_Parame(1).Genera_Cantid
   End If
   If r_dbl_TasInt = 0 Then
      'ITERACION 00
      If moddat_gf_Consulta_ParTasInt(l_arr_Parame, 0, 0, CDbl(modatecli_g_arr_DatCre(1).DatCre_MtoPre), CDbl(modatecli_g_arr_DatCre(1).DatCre_MtoPre), CDbl(modatecli_g_arr_DatCre(1).DatCre_ComVta), CDbl(modatecli_g_arr_DatCre(1).DatCre_ComVta), r_int_PlzPre, r_int_PlzPre, r_dbl_PorIni, r_dbl_PorIni, r_int_TipBon) Then
         r_dbl_TasInt = l_arr_Parame(1).Genera_Cantid
      End If
   End If

   p_ValPre = CDbl(modatecli_g_arr_DatCre(1).DatCre_MtoPre)
   p_ValInm = CDbl(modatecli_g_arr_DatCre(1).DatCre_ComVta)
   p_PlzPre = r_int_PlzPre
   p_PorIni = r_dbl_PorIni
   p_TasInt = r_dbl_TasInt
End Function

Private Function ff_Genera_NumSol() As String
Dim r_lng_NumSol     As Long
Dim r_str_NumSol     As String
   
   ff_Genera_NumSol = ""
   
   'Obteniendo Número de Solicitud
   Call moddat_gs_FecSis
   
   g_str_Parame = "SELECT * FROM CRE_FOLIOS WHERE "
   g_str_Parame = g_str_Parame & "FOLIOS_TIPFOL = 1 AND "
   g_str_Parame = g_str_Parame & "FOLIOS_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "FOLIOS_CODSUC = '" & modgen_g_str_CodSuc & "' AND "
   g_str_Parame = g_str_Parame & "FOLIOS_PERANO = " & Right(Format(Year(CDate(moddat_g_str_FecSis)), "0000"), 2)

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If

   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      r_lng_NumSol = 1
   Else
      r_lng_NumSol = g_rst_Genera!FOLIOS_NUMERO + 1
   End If

   r_str_NumSol = moddat_g_str_CodPrd & modgen_g_str_CodSuc & Right(Format(Year(CDate(moddat_g_str_FecSis)), "0000"), 2) & Format(r_lng_NumSol, "0000")
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      'Actualizando Correlativo
      g_str_Parame = "USP_CRE_FOLIOS ("
      g_str_Parame = g_str_Parame & "1, "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodPrd & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & Right(Format(Year(CDate(moddat_g_str_FecSis)), "0000"), 2) & ", "
      g_str_Parame = g_str_Parame & CStr(r_lng_NumSol) & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & "1, "
      If r_lng_NumSol = 1 Then
         g_str_Parame = g_str_Parame & "1) "
      Else
         g_str_Parame = g_str_Parame & "2) "
      End If
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_CRE_FOLIOS. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   ff_Genera_NumSol = r_str_NumSol
End Function

Private Function ff_Graba_SolMae(ByVal p_NumSol As String, ByVal p_TasInt As Double) As Integer
Dim r_dbl_TasInt     As Double
   
   ff_Graba_SolMae = False
   r_dbl_TasInt = p_TasInt
      
   'Tasa de Interes de Producto
   If moddat_gf_Consulta_ParSubPrd(l_arr_Parame, moddat_g_str_CodPrd, moddat_g_str_CodSub, "002", "101") Then
      r_dbl_TasInt = l_arr_Parame(1).Genera_Cantid
   End If
   
   Call moddat_gs_FecSis
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CRE_SOLMAE_INSERTA ("
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                                         'Numero de Solicitud
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodPrd & "', "                              'Código Producto
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodSub & "', "                              'Código SubProducto
      g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "          'Fecha Solicitud
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_CygTDo) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CygNDo & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipMon) & ", "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatCre(1).DatCre_ConHip & "', "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatCre(1).DatCre_EjeSeg & "', "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatCre(1).DatCre_Observ & "', "
      g_str_Parame = g_str_Parame & CStr(r_dbl_TasInt) & ", "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatCre(1).DatCre_ESgDes & "', "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_TipSeg) & ", "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatCre(1).DatCre_ESgViv & "', "
      g_str_Parame = g_str_Parame & CStr(CInt(modatecli_g_arr_DatCre(1).DatCre_DiaPag)) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_CuoExt) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_PlaAno) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_PlaAno * 12) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_PerGra) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_TipMon) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_ComVta_Dol) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_ComVta_Sol) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_CuoIni_Dol) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_CuoIni_Sol) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_MtoPre_Dol) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_MtoPre_Sol) & ", "
      If moddat_g_int_TipMon = 1 Then
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_MtoPre_Sol) & ", "
      Else
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_MtoPre_Dol) & ", "
      End If
      g_str_Parame = g_str_Parame & CStr(1) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_int_IngRegInm) & ", "                         'Flag de Registro de Ingresos Inmuebles
      g_str_Parame = g_str_Parame & CStr(modatecli_g_int_GasRegTar) & ", "                         'Flag de Registro de Gastos Tarjetas
      g_str_Parame = g_str_Parame & CStr(modatecli_g_int_GasRegFin) & ", "                         'Flag de Registro de Gastos Deudas
      g_str_Parame = g_str_Parame & CStr(modatecli_g_int_GasRegGas) & ", "                         'Flag de Registro de Gastos Mensuales
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_InmIde) & ", "           'Flag de Inmueble Identificado
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_TipEva) & ", "           '
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatCre(1).DatCre_InsFin & "', "          '
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_MonAho) & ", "           '
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_MtoAho) & ", "           '
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_MesAho) & ", "           '
      g_str_Parame = g_str_Parame & CStr(cmb_OfiCom.ItemData(cmb_OfiCom.ListIndex)) & ", "         '
      g_str_Parame = g_str_Parame & CStr(cmb_ForCon.ItemData(cmb_ForCon.ListIndex)) & ", "         '
      If cmb_ForCon.ItemData(cmb_ForCon.ListIndex) = 21 Then
         g_str_Parame = g_str_Parame & CStr(cmb_FerEve.ItemData(cmb_FerEve.ListIndex)) & ", "      '
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_PriViv) & ", "           '
      g_str_Parame = g_str_Parame & "' ', "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_FmvBbp) & ", "           '
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_MefPbp) & ", "           '
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_BMSTas) & ", "           '
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_MtoBMS) & ", "           '
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_MtoAFP) & ", "           '
            
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_MtoInm) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_MtoEst) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_MtoGCi) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_PreMto) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_TasEsp) & ", "           'Tasa Especial
                  
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                              'Código Sucursal
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_CRE_SOLMAE_INSERTA. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   ff_Graba_SolMae = True
End Function

Private Function ff_Graba_IngInm(ByVal p_NumSol As String) As Integer
   Dim r_int_Contad     As Integer

   ff_Graba_IngInm = False

   If modatecli_g_int_IngRegInm = 1 Then
      For r_int_Contad = 1 To UBound(modatecli_g_arr_IngresInm)
         moddat_g_int_FlgGOK = False
         moddat_g_int_CntErr = 0
         
         Do While moddat_g_int_FlgGOK = False
            g_str_Parame = "USP_INSERTA_CRE_SOLINB ("
            g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "
            g_str_Parame = g_str_Parame & CStr(r_int_Contad) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_IngresInm(r_int_Contad).IngInm_TipInm) & ", "
            g_str_Parame = g_str_Parame & Format(CDate(modatecli_g_arr_IngresInm(r_int_Contad).IngInm_FecAdq), "yyyymmdd") & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_IngresInm(r_int_Contad).IngInm_ImpVal) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_IngresInm(r_int_Contad).IngInm_TipMon) & ", "
            g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_IngresInm(r_int_Contad).IngInm_Direcc & "', "
         
            'Datos de Auditoria
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
            g_str_Parame = g_str_Parame & "1)"
         
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
               moddat_g_int_CntErr = moddat_g_int_CntErr + 1
            Else
               moddat_g_int_FlgGOK = True
            End If
      
            If moddat_g_int_CntErr = 6 Then
               If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRE_SOLINB. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
                  Exit Function
               Else
                  moddat_g_int_CntErr = 0
               End If
            End If
         Loop
      Next r_int_Contad
   End If
   
   ff_Graba_IngInm = True
End Function

Private Function ff_Graba_GasTrj(ByVal p_NumSol As String) As Integer
   Dim r_int_Contad     As Integer

   ff_Graba_GasTrj = False

   If modatecli_g_int_GasRegTar = 1 Then
      For r_int_Contad = 1 To UBound(modatecli_g_arr_GastosTar)
         moddat_g_int_FlgGOK = False
         moddat_g_int_CntErr = 0
         
         Do While moddat_g_int_FlgGOK = False
            g_str_Parame = "USP_INSERTA_CRE_SOLTRJ ("
            g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "
            g_str_Parame = g_str_Parame & CStr(r_int_Contad) & ", "
            g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_GastosTar(r_int_Contad).GasTar_InsFin & "', "
            g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_GastosTar(r_int_Contad).GasTar_NumTar & "', "
            g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_GastosTar(r_int_Contad).GasTar_TipTar & "', "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosTar(r_int_Contad).GasTar_TipMon) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosTar(r_int_Contad).GasTar_LinCre) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosTar(r_int_Contad).GasTar_SalPag) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosTar(r_int_Contad).GasTar_MonMin) & ", "
         
            'Datos de Auditoria
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
            g_str_Parame = g_str_Parame & "1)"
         
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
               moddat_g_int_CntErr = moddat_g_int_CntErr + 1
            Else
               moddat_g_int_FlgGOK = True
            End If
      
            If moddat_g_int_CntErr = 6 Then
               If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRE_SOLTRJ. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
                  Exit Function
               Else
                  moddat_g_int_CntErr = 0
               End If
            End If
         Loop
      Next r_int_Contad
   End If
   
   ff_Graba_GasTrj = True
End Function

Private Function ff_Graba_GasDeu(ByVal p_NumSol As String) As Integer
   Dim r_int_Contad     As Integer

   ff_Graba_GasDeu = False

   If modatecli_g_int_GasRegFin = 1 Then
      For r_int_Contad = 1 To UBound(modatecli_g_arr_GastosFin)
         moddat_g_int_FlgGOK = False
         moddat_g_int_CntErr = 0
         
         Do While moddat_g_int_FlgGOK = False
            g_str_Parame = "USP_INSERTA_CRE_SOLDEU ("
            g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "
            g_str_Parame = g_str_Parame & CStr(r_int_Contad) & ", "
            g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_GastosFin(r_int_Contad).GasFin_InsFin & "', "
            g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_GastosFin(r_int_Contad).GasFin_NumOpe & "', "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosFin(r_int_Contad).GasFin_TipMon) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosFin(r_int_Contad).GasFin_MonOto) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosFin(r_int_Contad).GasFin_SalPag) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosFin(r_int_Contad).GasFin_CuoMen) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosFin(r_int_Contad).GasFin_MesPag) & ", "
         
            'Datos de Auditoria
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
            g_str_Parame = g_str_Parame & "1)"
         
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
               moddat_g_int_CntErr = moddat_g_int_CntErr + 1
            Else
               moddat_g_int_FlgGOK = True
            End If
      
            If moddat_g_int_CntErr = 6 Then
               If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRE_SOLDEU. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
                  Exit Function
               Else
                  moddat_g_int_CntErr = 0
               End If
            End If
         Loop
      Next r_int_Contad
   End If
   
   ff_Graba_GasDeu = True
End Function

Private Function ff_Graba_GasGas(ByVal p_NumSol As String) As Integer
   Dim r_int_Contad     As Integer

   ff_Graba_GasGas = False

   If modatecli_g_int_GasRegGas = 1 Then
      For r_int_Contad = 1 To UBound(modatecli_g_arr_GastosGas)
         moddat_g_int_FlgGOK = False
         moddat_g_int_CntErr = 0
         
         Do While moddat_g_int_FlgGOK = False
            g_str_Parame = "USP_INSERTA_CRE_SOLEYM ("
            g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "
            g_str_Parame = g_str_Parame & CStr(r_int_Contad) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosGas(r_int_Contad).GasGas_TipGas) & ", "
            g_str_Parame = g_str_Parame & "1, "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosGas(r_int_Contad).GasGas_ImpVal) & ", "
         
            'Datos de Auditoria
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
            g_str_Parame = g_str_Parame & "1)"
         
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
               moddat_g_int_CntErr = moddat_g_int_CntErr + 1
            Else
               moddat_g_int_FlgGOK = True
            End If
      
            If moddat_g_int_CntErr = 6 Then
               If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRE_SOLEYM. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
                  Exit Function
               Else
                  moddat_g_int_CntErr = 0
               End If
            End If
         Loop
      Next r_int_Contad
   End If
   
   ff_Graba_GasGas = True
End Function

Private Function ff_Graba_Refere(ByVal p_NumSol As String) As Integer
   ff_Graba_Refere = False
   
   'Grabando Referencia Familiar 1
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_INSERTA_CRE_SOLREF ("
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                                                  'Número de Solicitud
      g_str_Parame = g_str_Parame & "3, "                                                                   'Tipo de Referencia (Familiar)
      g_str_Parame = g_str_Parame & "1, "                                                                   'Número de Referencia (Familiar)
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Refere(1).Refere_TipPar) & ", "                    'Tipo de Parentesco
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(1).Refere_ApePat & "', "                   'Apellido Paterno
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(1).Refere_ApeMat & "', "                   'Apellido Materno
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(1).Refere_Nombre & "', "                   'Nombres
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(1).Refere_Telefo & "', "                   'Teléfono
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(1).Refere_Celula & "', "                   'Celular
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                                       'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRE_SOLREF. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   'Grabando Referencia Familiar 2
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_INSERTA_CRE_SOLREF ("
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                                                  'Número de Solicitud
      g_str_Parame = g_str_Parame & "3, "                                                                   'Tipo de Referencia (Familiar)
      g_str_Parame = g_str_Parame & "2, "                                                                   'Número de Referencia (Familiar)
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Refere(2).Refere_TipPar) & ", "                    'Tipo de Parentesco
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(2).Refere_ApePat & "', "                   'Apellido Paterno
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(2).Refere_ApeMat & "', "                   'Apellido Materno
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(2).Refere_Nombre & "', "                   'Nombres
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(2).Refere_Telefo & "', "                   'Teléfono
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(2).Refere_Celula & "', "                   'Celular
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                                       'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRE_SOLREF. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   'Grabando Referencia No Familiar 1
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_INSERTA_CRE_SOLREF ("
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                                                  'Número de Solicitud
      g_str_Parame = g_str_Parame & "3, "                                                                   'Tipo de Referencia (Familiar)
      g_str_Parame = g_str_Parame & "3, "                                                                   'Número de Referencia (Familiar)
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Refere(3).Refere_TipPar) & ", "                    'Tipo de Parentesco
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(3).Refere_ApePat & "', "                   'Apellido Paterno
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(3).Refere_ApeMat & "', "                   'Apellido Materno
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(3).Refere_Nombre & "', "                   'Nombres
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(3).Refere_Telefo & "', "                   'Teléfono
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(3).Refere_Celula & "', "                   'Celular
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                                       'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRE_SOLREF. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   'Grabando Referencia No Familiar 2
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_INSERTA_CRE_SOLREF ("
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                                                  'Número de Solicitud
      g_str_Parame = g_str_Parame & "3, "                                                                   'Tipo de Referencia (Familiar)
      g_str_Parame = g_str_Parame & "4, "                                                                   'Número de Referencia (Familiar)
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Refere(4).Refere_TipPar) & ", "                    'Tipo de Parentesco
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(4).Refere_ApePat & "', "                   'Apellido Paterno
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(4).Refere_ApeMat & "', "                   'Apellido Materno
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(4).Refere_Nombre & "', "                   'Nombres
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(4).Refere_Telefo & "', "                   'Teléfono
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(4).Refere_Celula & "', "                   'Celular
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                                       'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRE_SOLREF. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   ff_Graba_Refere = True
End Function

Private Function ff_Graba_SolEje(ByVal p_NumSol As String) As Integer
   ff_Graba_SolEje = False
   
   Call moddat_gs_FecSis
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CRE_SOLEJE_INSERTA ("
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                                                  'Número de Solicitud
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatCre(1).DatCre_EjeSeg & "', "                   'Ejecutivo de Seguimiento
      g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "                   'Fecha de Asignación
      g_str_Parame = g_str_Parame & "0, "                                                                   'Fecha de Baja
      g_str_Parame = g_str_Parame & "1, "                                                                   'Situación
      g_str_Parame = g_str_Parame & "'', "                                                                  'Observaciones
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                                       'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_CRE_SOLEJE_INSERTA. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   ff_Graba_SolEje = True
End Function

Private Function ff_Graba_SolDoc(ByVal p_NumSol As String) As Integer
   Dim r_int_Contad     As Integer

   Call moddat_gs_FecSis
   ff_Graba_SolDoc = False
   
   For r_int_Contad = 1 To UBound(modatecli_g_arr_DocCre)
      If Not moddat_gf_Inserta_SolDoc(p_NumSol, modatecli_g_arr_DocCre(r_int_Contad).DocCre_TipDoc, moddat_g_str_CodPrd, moddat_g_str_CodSub, modatecli_g_arr_DocCre(r_int_Contad).DocCre_CodAct, modatecli_g_arr_DocCre(r_int_Contad).DocCre_CodGrp, modatecli_g_arr_DocCre(r_int_Contad).DocCre_CodIte, Format(CDate(moddat_g_str_FecSis), "yyyymmdd")) Then
         Exit Function
      End If
   Next r_int_Contad
   
   ff_Graba_SolDoc = True
End Function

Private Function ff_Graba_Seguim(ByVal p_NumSol As String) As Integer
   ff_Graba_Seguim = False

   If Not moddat_gf_Inserta_Seguim(p_NumSol, 11) Then
      Exit Function
   End If
   
   If Not moddat_gf_Inserta_SegDet(p_NumSol, 11, 16, 0, "", 0, 0) Then
      Exit Function
   End If

   ff_Graba_Seguim = True
End Function

Private Function ff_Graba_Inmueb(ByVal p_NumSol As String) As Integer
   ff_Graba_Inmueb = False
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CRE_SOLINM ("
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                                         'Número de Solicitud
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_TipInm) & ", "           'Tipo de Inmueble
      g_str_Parame = g_str_Parame & "'" & Format(CInt(modatecli_g_arr_DatInm(1).DatInm_Modali), "00") & "', "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_TipVia) & ", "           'Tipo de Via
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_NomVia & "', "          'Nombre de Via
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Numero & "', "          'Número en Via
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Interi & "', "          'Interior / Dpto
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_TipZon) & ", "           'Tipo de Zona
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_NomZon & "', "          'Nombre de Zona
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Refere & "', "          'Referencia
      g_str_Parame = g_str_Parame & modatecli_g_arr_DatInm(1).DatInm_FlgEst & ", "                 'Estacionamiento Flag
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Estaci & "', "          'Estacionamiento
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_UbiGeo & "', "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_PryMCs) & ", "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_BcoPry & "', "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_CodPry & "', "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_NomPry & "', "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_FlgPro) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_TipDoc_Pro) & ", "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_NumDoc_Pro & "', "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_RazSoc_Pro & "', "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_TipVia_Pro) & ", "           'Tipo de Via
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_NomVia_Pro & "', "          'Nombre de Via
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_NumVia_Pro & "', "          'Número en Via
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_IntDpt_Pro & "', "          'Interior / Dpto
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_TipZon_Pro) & ", "           'Tipo de Zona
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_NomZon_Pro & "', "          'Nombre de Zona
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Refere_Pro & "', "          'Nombre de Zona
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_UbiGeo_Pro & "', "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Telefo_Pro & "', "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_FlgCon) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_TipDoc_Con) & ", "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_NumDoc_Con & "', "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_RazSoc_Con & "', "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_TipVia_Con) & ", "           'Tipo de Via
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_NomVia_Con & "', "          'Nombre de Via
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_NumVia_Con & "', "          'Número en Via
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_IntDpt_Con & "', "          'Interior / Dpto
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_TipZon_Con) & ", "           'Tipo de Zona
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_NomZon_Con & "', "          'Nombre de Zona
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Refere_Con & "', "          'Nombre de Zona
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_UbiGeo_Con & "', "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Telefo_Con & "', "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                               'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_CRE_SOLINM. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   ff_Graba_Inmueb = True
End Function


