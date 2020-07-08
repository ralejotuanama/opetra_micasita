VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_MntEmp_52 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   7305
   ClientLeft      =   2790
   ClientTop       =   1965
   ClientWidth     =   11640
   Icon            =   "OpeTra_frm_167.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7335
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   11685
      _Version        =   65536
      _ExtentX        =   20611
      _ExtentY        =   12938
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
      Begin Threed.SSPanel SSPanel8 
         Height          =   435
         Left            =   30
         TabIndex        =   57
         Top             =   3090
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
         _ExtentY        =   767
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
         Begin VB.ComboBox cmb_PaiRes 
            Height          =   315
            Left            =   1980
            TabIndex        =   3
            Text            =   "cmb_PaiRes"
            Top             =   60
            Width           =   3315
         End
         Begin VB.Label Label4 
            Caption         =   "País:"
            Height          =   315
            Left            =   60
            TabIndex        =   58
            Top             =   60
            Width           =   1455
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1755
         Left            =   30
         TabIndex        =   43
         Top             =   3570
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
         _ExtentY        =   3096
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
         Begin VB.ComboBox cmb_DstDir 
            Height          =   315
            Left            =   1980
            TabIndex        =   12
            Text            =   "cmb_DstDir"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DptDir 
            Height          =   315
            Left            =   1980
            TabIndex        =   10
            Text            =   "cmb_DptDir"
            Top             =   1050
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipZon 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   720
            Width           =   3315
         End
         Begin VB.TextBox txt_NomVia 
            Height          =   315
            Left            =   1980
            MaxLength       =   120
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   390
            Width           =   3315
         End
         Begin VB.TextBox txt_Refere 
            Height          =   315
            Left            =   8190
            MaxLength       =   250
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PrvDir 
            Height          =   315
            Left            =   8190
            TabIndex        =   11
            Text            =   "cmb_PrvDir"
            Top             =   1050
            Width           =   3315
         End
         Begin VB.TextBox txt_NomZon 
            Height          =   315
            Left            =   8190
            MaxLength       =   120
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   720
            Width           =   3315
         End
         Begin VB.TextBox txt_IntDpt 
            Height          =   315
            Left            =   9840
            MaxLength       =   30
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   390
            Width           =   1665
         End
         Begin VB.TextBox txt_NumVia 
            Height          =   315
            Left            =   8190
            MaxLength       =   30
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   390
            Width           =   1640
         End
         Begin VB.ComboBox cmb_TipVia 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   60
            Width           =   3315
         End
         Begin VB.Label lbl_General 
            Caption         =   "Distrito:"
            Height          =   315
            Index           =   45
            Left            =   60
            TabIndex        =   52
            Top             =   1380
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Departamento:"
            Height          =   315
            Index           =   44
            Left            =   60
            TabIndex        =   51
            Top             =   1050
            Width           =   1425
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo de Zona:"
            Height          =   315
            Index           =   43
            Left            =   60
            TabIndex        =   50
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Vía:"
            Height          =   285
            Index           =   42
            Left            =   60
            TabIndex        =   49
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Referencia:"
            Height          =   285
            Index           =   54
            Left            =   6180
            TabIndex        =   48
            Top             =   1380
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Provincia:"
            Height          =   315
            Index           =   53
            Left            =   6180
            TabIndex        =   47
            Top             =   1050
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Zona:"
            Height          =   285
            Index           =   52
            Left            =   6180
            TabIndex        =   46
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nro - Int/Dpto/Mza/Lote:"
            Height          =   285
            Index           =   51
            Left            =   6180
            TabIndex        =   45
            Top             =   390
            Width           =   1935
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo de Vía:"
            Height          =   285
            Index           =   41
            Left            =   60
            TabIndex        =   44
            Top             =   60
            Width           =   1545
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   1095
         Left            =   30
         TabIndex        =   38
         Top             =   6180
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
         _ExtentY        =   1931
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
         Begin VB.TextBox txt_TeleRH 
            Height          =   315
            Left            =   1980
            MaxLength       =   25
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   390
            Width           =   1640
         End
         Begin VB.TextBox txt_AnexRH 
            Height          =   315
            Left            =   3630
            MaxLength       =   5
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   390
            Width           =   1640
         End
         Begin VB.TextBox txt_Telef1 
            Height          =   315
            Left            =   1980
            MaxLength       =   25
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   60
            Width           =   1640
         End
         Begin VB.TextBox txt_NumFax 
            Height          =   315
            Left            =   8190
            MaxLength       =   25
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   60
            Width           =   1640
         End
         Begin VB.TextBox txt_Telef2 
            Height          =   315
            Left            =   3630
            MaxLength       =   25
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   60
            Width           =   1640
         End
         Begin VB.TextBox txt_PagWeb 
            Height          =   315
            Left            =   1980
            MaxLength       =   120
            TabIndex        =   22
            Text            =   "Text1"
            Top             =   720
            Width           =   3315
         End
         Begin VB.Label lbl_General 
            Caption         =   "Teléfono/Anexo RR.HH:"
            Height          =   285
            Index           =   47
            Left            =   60
            TabIndex        =   42
            Top             =   390
            Width           =   1815
         End
         Begin VB.Label lbl_General 
            Caption         =   "Teléfono (s):"
            Height          =   285
            Index           =   46
            Left            =   60
            TabIndex        =   41
            Top             =   60
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fax:"
            Height          =   285
            Index           =   55
            Left            =   6180
            TabIndex        =   40
            Top             =   60
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Página Web:"
            Height          =   285
            Index           =   0
            Left            =   60
            TabIndex        =   39
            Top             =   720
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1095
         Left            =   30
         TabIndex        =   28
         Top             =   1950
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
         _ExtentY        =   1931
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
         Begin VB.ComboBox cmb_CodCiu 
            Height          =   315
            Left            =   1980
            TabIndex        =   2
            Text            =   "cmb_DptDir"
            Top             =   720
            Width           =   9525
         End
         Begin VB.TextBox txt_NomCom 
            Height          =   315
            Left            =   1980
            MaxLength       =   250
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   390
            Width           =   9525
         End
         Begin VB.TextBox txt_RazSoc 
            Height          =   315
            Left            =   1980
            MaxLength       =   250
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   60
            Width           =   9525
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Comercial:"
            Height          =   285
            Index           =   49
            Left            =   60
            TabIndex        =   31
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Razón Social:"
            Height          =   285
            Index           =   37
            Left            =   60
            TabIndex        =   30
            Top             =   60
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "CIIU:"
            Height          =   285
            Index           =   39
            Left            =   60
            TabIndex        =   29
            Top             =   720
            Width           =   1365
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   32
         Top             =   30
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   555
            Left            =   630
            TabIndex        =   33
            Top             =   60
            Width           =   4485
            _Version        =   65536
            _ExtentX        =   7911
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "Mantenimiento de Empresas"
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
         Begin Threed.SSPanel pnl_RelAcc 
            Height          =   555
            Left            =   5700
            TabIndex        =   59
            Top             =   60
            Width           =   5835
            _Version        =   65536
            _ExtentX        =   10292
            _ExtentY        =   979
            _StockProps     =   15
            ForeColor       =   16777215
            BackColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_167.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   34
         Top             =   1470
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
         _ExtentY        =   767
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
         Begin Threed.SSPanel pnl_DocIde 
            Height          =   315
            Left            =   1980
            TabIndex        =   35
            Top             =   60
            Width           =   9555
            _Version        =   65536
            _ExtentX        =   16854
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "7-20511904162 - EDPYME MICASITA S.A."
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
         Begin VB.Label Label1 
            Caption         =   "Documento Identidad:"
            Height          =   315
            Left            =   60
            TabIndex        =   36
            Top             =   60
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   675
         Left            =   30
         TabIndex        =   37
         Top             =   750
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
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
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_167.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Modificar Datos del Cliente"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Cancel 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_167.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Cancelar Modificación de Datos Generales"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_167.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10980
            Picture         =   "OpeTra_frm_167.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   765
         Left            =   30
         TabIndex        =   53
         Top             =   5370
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
         Begin VB.TextBox txt_CodPos 
            Height          =   315
            Left            =   8190
            MaxLength       =   250
            TabIndex        =   16
            Text            =   "Text1"
            Top             =   390
            Width           =   1640
         End
         Begin VB.ComboBox cmb_PrvEst 
            Height          =   315
            Left            =   1980
            TabIndex        =   15
            Text            =   "cmb_DptDir"
            Top             =   390
            Width           =   3315
         End
         Begin VB.TextBox txt_Direcc 
            Height          =   315
            Left            =   1980
            MaxLength       =   250
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   60
            Width           =   9555
         End
         Begin VB.Label Label28 
            Caption         =   "Código Postal:"
            Height          =   285
            Left            =   6180
            TabIndex        =   56
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label24 
            Caption         =   "Provincia / Estado:"
            Height          =   315
            Left            =   60
            TabIndex        =   55
            Top             =   390
            Width           =   1905
         End
         Begin VB.Label Label20 
            Caption         =   "Dirección:"
            Height          =   285
            Left            =   60
            TabIndex        =   54
            Top             =   60
            Width           =   1485
         End
      End
   End
End
Attribute VB_Name = "frm_MntEmp_52"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_PaiRes()   As moddat_tpo_Genera
Dim l_arr_PrvEst()   As moddat_tpo_Genera
Dim l_str_CodCiu     As String
Dim l_str_PaiRes     As String
Dim l_str_DptDir     As String
Dim l_str_PrvDir     As String
Dim l_str_DstDir     As String
Dim l_str_PrvEst     As String
Dim l_int_FlgCmb     As Integer
Dim l_int_RelAcc     As Integer
Dim l_int_AccTDo     As Integer
Dim l_str_AccNDo     As String
Dim l_int_AccVin     As Integer

Private Sub cmb_CodCiu_Change()
   l_str_CodCiu = cmb_CodCiu.Text
End Sub

Private Sub cmb_CodCiu_Click()
   If cmb_CodCiu.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(cmb_PaiRes)
      End If
   End If
End Sub

Private Sub cmb_CodCiu_GotFocus()
   l_int_FlgCmb = True
   l_str_CodCiu = cmb_CodCiu.Text
End Sub

Private Sub cmb_CodCiu_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + modgen_g_con_NUMERO + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_CodCiu, l_str_CodCiu)
      l_int_FlgCmb = True
      
      If cmb_CodCiu.ListIndex > -1 Then
         l_str_CodCiu = ""
      End If
      
      Call gs_SetFocus(cmb_PaiRes)
   End If
End Sub

Private Sub cmb_DptDir_Change()
   l_str_DptDir = cmb_DptDir.Text
End Sub

Private Sub cmb_DptDir_Click()
   If cmb_DptDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_PrvDir.Clear
         cmb_DstDir.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_PrvDir)
      End If
   End If
End Sub

Private Sub cmb_DptDir_GotFocus()
   l_int_FlgCmb = True
   l_str_DptDir = cmb_DptDir.Text
End Sub

Private Sub cmb_DptDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DptDir, l_str_DptDir)
      l_int_FlgCmb = True
      
      cmb_PrvDir.Clear
      cmb_DstDir.Clear
      If cmb_DptDir.ListIndex > -1 Then
         l_str_DptDir = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_PrvDir)
   End If
End Sub

Private Sub cmb_DstDir_Change()
   l_str_DstDir = cmb_DstDir.Text
End Sub

Private Sub cmb_DstDir_Click()
   If cmb_DstDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_Refere)
      End If
   End If
End Sub

Private Sub cmb_DstDir_GotFocus()
   l_int_FlgCmb = True
   l_str_DstDir = cmb_DstDir.Text
End Sub

Private Sub cmb_DstDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DstDir, l_str_DstDir)
      l_int_FlgCmb = True
      
      If cmb_DstDir.ListIndex > -1 Then
         l_str_DstDir = ""
      End If
      
      Call gs_SetFocus(txt_Refere)
   End If
End Sub

Private Sub cmb_PaiRes_Change()
   l_str_PaiRes = cmb_PaiRes.Text
   
   cmb_PaiRes.SelLength = Len(l_str_PaiRes)
End Sub

Private Sub cmb_PaiRes_Click()
   If cmb_PaiRes.ListIndex > -1 Then
      If l_int_FlgCmb Then
         If l_arr_PaiRes(cmb_PaiRes.ListIndex + 1).Genera_Codigo <> "004028" Then
            cmb_TipVia.ListIndex = -1
            txt_NomVia.Text = ""
            txt_NumVia.Text = ""
            txt_IntDpt.Text = ""
            cmb_TipZon.ListIndex = -1
            txt_NomZon.Text = ""
            cmb_DptDir.ListIndex = -1
            cmb_PrvDir.Clear
            cmb_DstDir.Clear
            txt_Refere.Text = ""
            
            cmb_TipVia.Enabled = False
            txt_NomVia.Enabled = False
            txt_NumVia.Enabled = False
            txt_IntDpt.Enabled = False
            cmb_TipZon.Enabled = False
            txt_NomZon.Enabled = False
            cmb_DptDir.Enabled = False
            cmb_PrvDir.Enabled = False
            cmb_DstDir.Enabled = False
            txt_Refere.Enabled = False
            
            txt_Direcc.Enabled = True
            cmb_PrvEst.Enabled = True
            txt_CodPos.Enabled = True
            
            'Cargar Provincia / Estado segñun País seleccionado
            Call modmip_gs_Carga_CiuExt(l_arr_PrvEst, cmb_PrvEst, l_arr_PaiRes(cmb_PaiRes.ListIndex + 1).Genera_Codigo)
            
            Call gs_SetFocus(txt_Direcc)
         Else
            cmb_TipVia.Enabled = True
            txt_NomVia.Enabled = True
            txt_NumVia.Enabled = True
            txt_IntDpt.Enabled = True
            cmb_TipZon.Enabled = True
            txt_NomZon.Enabled = True
            cmb_DptDir.Enabled = True
            cmb_PrvDir.Enabled = True
            cmb_DstDir.Enabled = True
            txt_Refere.Enabled = True
         
            txt_Direcc.Text = ""
            cmb_PrvEst.Clear
            txt_CodPos.Text = ""
            
            txt_Direcc.Enabled = False
            cmb_PrvEst.Enabled = False
            txt_CodPos.Enabled = False
            
            Call gs_SetFocus(cmb_TipVia)
         End If
      End If
   Else
      cmb_TipVia.ListIndex = -1
      txt_NomVia.Text = ""
      txt_NumVia.Text = ""
      txt_IntDpt.Text = ""
      cmb_TipZon.ListIndex = -1
      txt_NomZon.Text = ""
      cmb_DptDir.ListIndex = -1
      cmb_PrvDir.Clear
      cmb_DstDir.Clear
      txt_Refere.Text = ""
      
      cmb_TipVia.Enabled = False
      txt_NomVia.Enabled = False
      txt_NumVia.Enabled = False
      txt_IntDpt.Enabled = False
      cmb_TipZon.Enabled = False
      txt_NomZon.Enabled = False
      cmb_DptDir.Enabled = False
      cmb_PrvDir.Enabled = False
      cmb_DstDir.Enabled = False
      txt_Refere.Enabled = False
   
      txt_Direcc.Text = ""
      cmb_PrvEst.ListIndex = -1
      txt_CodPos.Text = ""
      
      txt_Direcc.Enabled = False
      cmb_PrvEst.Enabled = False
      txt_CodPos.Enabled = False
      
      Call gs_SetFocus(txt_Telef1)
   End If
End Sub

Private Sub cmb_PaiRes_GotFocus()
   l_int_FlgCmb = True
End Sub

Private Sub cmb_PaiRes_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_PaiRes, l_str_PaiRes)
      l_int_FlgCmb = True
      
      If cmb_PaiRes.ListIndex > -1 Then
         l_str_PaiRes = ""
         
         If l_arr_PaiRes(cmb_PaiRes.ListIndex + 1).Genera_Codigo <> "004028" Then
            cmb_TipVia.ListIndex = -1
            txt_NomVia.Text = ""
            txt_NumVia.Text = ""
            txt_IntDpt.Text = ""
            cmb_TipZon.ListIndex = -1
            txt_NomZon.Text = ""
            cmb_DptDir.ListIndex = -1
            cmb_PrvDir.Clear
            cmb_DstDir.Clear
            txt_Refere.Text = ""
            
            cmb_TipVia.Enabled = False
            txt_NomVia.Enabled = False
            txt_NumVia.Enabled = False
            txt_IntDpt.Enabled = False
            cmb_TipZon.Enabled = False
            txt_NomZon.Enabled = False
            cmb_DptDir.Enabled = False
            cmb_PrvDir.Enabled = False
            cmb_DstDir.Enabled = False
            txt_Refere.Enabled = False
            
            txt_Direcc.Enabled = True
            cmb_PrvEst.Enabled = True
            txt_CodPos.Enabled = True
            
            'Cargar Provincia/estado
            Call modmip_gs_Carga_CiuExt(l_arr_PrvEst, cmb_PrvEst, l_arr_PaiRes(cmb_PaiRes.ListIndex + 1).Genera_Codigo)
            
            Call gs_SetFocus(txt_Direcc)
         Else
            cmb_TipVia.Enabled = True
            txt_NomVia.Enabled = True
            txt_NumVia.Enabled = True
            txt_IntDpt.Enabled = True
            cmb_TipZon.Enabled = True
            txt_NomZon.Enabled = True
            cmb_DptDir.Enabled = True
            cmb_PrvDir.Enabled = True
            cmb_DstDir.Enabled = True
            txt_Refere.Enabled = True
         
            txt_Direcc.Text = ""
            cmb_PrvEst.Clear
            txt_CodPos.Text = ""
            
            txt_Direcc.Enabled = False
            cmb_PrvEst.Enabled = False
            txt_CodPos.Enabled = False
            
            Call gs_SetFocus(cmb_TipVia)
         End If
      Else
         cmb_TipVia.ListIndex = -1
         txt_NomVia.Text = ""
         txt_NumVia.Text = ""
         txt_IntDpt.Text = ""
         cmb_TipZon.ListIndex = -1
         txt_NomZon.Text = ""
         cmb_DptDir.ListIndex = -1
         cmb_PrvDir.Clear
         cmb_DstDir.Clear
         txt_Refere.Text = ""
         
         cmb_TipVia.Enabled = False
         txt_NomVia.Enabled = False
         txt_NumVia.Enabled = False
         txt_IntDpt.Enabled = False
         cmb_TipZon.Enabled = False
         txt_NomZon.Enabled = False
         cmb_DptDir.Enabled = False
         cmb_PrvDir.Enabled = False
         cmb_DstDir.Enabled = False
         txt_Refere.Enabled = False
      
         txt_Direcc.Text = ""
         cmb_PrvEst.ListIndex = -1
         txt_CodPos.Text = ""
         
         txt_Direcc.Enabled = False
         cmb_PrvEst.Enabled = False
         txt_CodPos.Enabled = False
         
         Call gs_SetFocus(txt_Telef1)
      End If
   End If
End Sub

Private Sub cmb_PrvDir_Change()
   l_str_PrvDir = cmb_PrvDir.Text
End Sub

Private Sub cmb_PrvDir_Click()
   If cmb_PrvDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_DstDir.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"), Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_DstDir)
      End If
   End If
End Sub

Private Sub cmb_PrvDir_GotFocus()
   l_int_FlgCmb = True
   l_str_PrvDir = cmb_PrvDir.Text
End Sub

Private Sub cmb_PrvDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_PrvDir, l_str_PrvDir)
      l_int_FlgCmb = True
      
      cmb_DstDir.Clear
      If cmb_PrvDir.ListIndex > -1 Then
         l_str_DstDir = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"), Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_DstDir)
   End If
End Sub

Private Sub cmb_PrvEst_Change()
   l_str_PrvEst = cmb_PrvEst.Text
End Sub

Private Sub cmb_PrvEst_Click()
   If cmb_PrvEst.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_CodPos)
      End If
   End If
End Sub

Private Sub cmb_PrvEst_GotFocus()
   l_int_FlgCmb = True
   l_str_PrvEst = cmb_PrvEst.Text
End Sub

Private Sub cmb_PrvEst_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_PrvEst, l_str_PrvEst)
      l_int_FlgCmb = True
      
      If cmb_PrvEst.ListIndex > -1 Then
         l_str_PrvEst = ""
      End If
      
      Call gs_SetFocus(txt_CodPos)
   End If
End Sub

Private Sub cmb_TipVia_Click()
   Call gs_SetFocus(txt_NomVia)
End Sub

Private Sub cmb_TipVia_KeyPress(KeyAscii As Integer)
   Call cmb_TipVia_Click
End Sub

Private Sub cmb_TipZon_Click()
   Call gs_SetFocus(txt_NomZon)
End Sub

Private Sub cmb_TipZon_KeyPress(KeyAscii As Integer)
   Call cmb_TipZon_Click
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_Limpia
   Call fs_Cargar_Datos
   Call fs_Activa(False)
End Sub

Private Sub cmd_Editar_Click()
   Call fs_Activa(True)
   
   If l_arr_PaiRes(cmb_PaiRes.ListIndex + 1).Genera_Codigo <> "004028" Then
      cmb_TipVia.Enabled = False
      txt_NomVia.Enabled = False
      txt_NumVia.Enabled = False
      txt_IntDpt.Enabled = False
      cmb_TipZon.Enabled = False
      txt_NomZon.Enabled = False
      cmb_DptDir.Enabled = False
      cmb_PrvDir.Enabled = False
      cmb_DstDir.Enabled = False
      txt_Refere.Enabled = False
   Else
      txt_Direcc.Enabled = False
      cmb_PrvEst.Enabled = False
      txt_CodPos.Enabled = False
   End If
   'modmip_g_int_FlgGrb = 2
   Call gs_SetFocus(txt_RazSoc)
End Sub

Private Sub cmd_Grabar_Click()
   If Len(Trim(txt_RazSoc.Text)) = 0 Then
      MsgBox "Debe ingresar la Razón Social.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_RazSoc)
      Exit Sub
   End If
   
   If Len(Trim(txt_NomCom.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre Comercial.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomCom)
      Exit Sub
   End If
   
   If cmb_CodCiu.ListIndex = -1 Then
      MsgBox "Debe seleccionar el CIIU.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodCiu)
      Exit Sub
   End If
   
   If cmb_PaiRes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el País.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PaiRes)
      Exit Sub
   End If
   
   If CInt(l_arr_PaiRes(cmb_PaiRes.ListIndex + 1).Genera_Codigo) = 4028 Then
      If cmb_TipVia.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Vía.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipVia)
         Exit Sub
      End If
      
      If Len(Trim(txt_NomVia.Text)) = 0 Then
         MsgBox "Debe ingresar el Nombre de Vía.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NomVia)
         Exit Sub
      End If
      
      If Len(Trim(txt_NumVia.Text)) = 0 Then
         MsgBox "Debe ingresar el Número.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumVia)
         Exit Sub
      End If
      
      If cmb_TipZon.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Zona.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipZon)
         Exit Sub
      End If
      
      If cmb_TipZon.ItemData(cmb_TipZon.ListIndex) <> 12 Then
         If Len(Trim(txt_NomZon.Text)) = 0 Then
            MsgBox "Debe ingresar el Nombre de Zona.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NomZon)
            Exit Sub
         End If
      End If
      
      If cmb_DptDir.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Departamento de la Dirección.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_DptDir)
         Exit Sub
      End If
      
      If cmb_PrvDir.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Provincia de la Dirección.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_PrvDir)
         Exit Sub
      End If
      
      If cmb_DstDir.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Distrito de la Dirección.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_DstDir)
         Exit Sub
      End If
   Else
      If Len(Trim(txt_Direcc.Text)) = 0 Then
         MsgBox "Debe ingresar la Dirección.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Direcc)
         Exit Sub
      End If
      
      If cmb_PrvEst.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Provincia / Estado.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_PrvEst)
         Exit Sub
      End If
      
      If Len(Trim(txt_CodPos.Text)) = 0 Then
         MsgBox "Debe ingresar el Código Postal.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_CodPos)
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Grabando Información de la Empresa
   g_str_Parame = "USP_EMP_DATGEN_MANTEN ("
   g_str_Parame = g_str_Parame & CStr(modmip_g_int_TDoEmp) & ", "
   g_str_Parame = g_str_Parame & "'" & modmip_g_str_NDoEmp & "', "
   
   g_str_Parame = g_str_Parame & "'" & txt_RazSoc & "', "
   g_str_Parame = g_str_Parame & "'" & txt_NomCom & "', "
   g_str_Parame = g_str_Parame & CStr(cmb_CodCiu.ItemData(cmb_CodCiu.ListIndex)) & ", "
   
   g_str_Parame = g_str_Parame & "'" & l_arr_PaiRes(cmb_PaiRes.ListIndex + 1).Genera_Codigo & "', "
   
   If CInt(l_arr_PaiRes(cmb_PaiRes.ListIndex + 1).Genera_Codigo) = 4028 Then
      g_str_Parame = g_str_Parame & CStr(cmb_TipVia.ItemData(cmb_TipVia.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NomVia.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NumVia.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_IntDpt.Text & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_TipZon.ItemData(cmb_TipZon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NomZon.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Refere.Text & "', "
      g_str_Parame = g_str_Parame & "'" & Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00") & Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00") & Format(cmb_DstDir.ItemData(cmb_DstDir.ListIndex), "00") & "', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
   Else
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'" & txt_Direcc.Text & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_PrvEst(cmb_PrvEst.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & "'" & txt_CodPos.Text & "', "
   End If
   
   g_str_Parame = g_str_Parame & "'" & txt_Telef1.Text & "', "
   g_str_Parame = g_str_Parame & "'" & txt_Telef2.Text & "', "
   g_str_Parame = g_str_Parame & "'" & txt_NumFax.Text & "', "
   g_str_Parame = g_str_Parame & "'" & txt_TeleRH.Text & "', "
   g_str_Parame = g_str_Parame & "'" & txt_AnexRH.Text & "', "
   g_str_Parame = g_str_Parame & "'" & txt_PagWeb.Text & "', "
   
   g_str_Parame = g_str_Parame & CStr(l_int_AccTDo) & ", "
   g_str_Parame = g_str_Parame & "'" & l_str_AccNDo & "', "
   g_str_Parame = g_str_Parame & CStr(l_int_AccVin) & ", "
   g_str_Parame = g_str_Parame & "'" & CStr(l_int_RelAcc) & "', "

   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
   g_str_Parame = g_str_Parame & CStr(modmip_g_int_FlgGrb) & ")"
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar el Procedimiento USP_EMP_DATGEN_MANTEN.", vbCritical, modgen_g_str_NomPlt
      MsgBox g_str_Parame & vbEnter & modmip_g_int_FlgGrb & vbEnter & modmip_g_int_FlgAct_2
      Exit Sub
   End If

   modmip_g_int_FlgGrb = 2
   modmip_g_int_FlgAct_2 = 2
   
   MsgBox "Los datos se grabaron correctamente.", vbInformation, modgen_g_str_NomPlt
   
   Call fs_Activa(False)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Dim r_str_NomVin     As String
   Dim r_str_NomAcc     As String
   
   Screen.MousePointer = 11
   
   Call gs_CentraForm(Me)
   
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_DocIde.Caption = modmip_g_str_TDoEmp & " - " & modmip_g_str_NDoEmp
   pnl_RelAcc.Visible = False
   
   Call fs_Inicio
   
   If modmip_g_int_FlgGrb = 1 Then
      Call fs_Activa(True)
      cmd_Cancel.Enabled = False
      
      Call fs_Limpia
   Else
      Call fs_Limpia
      Call fs_Cargar_Datos
      Call fs_Activa(False)
   End If
   
   'Verificando Relación con Accionistas
   Call modmip_gs_RelAcc(modmip_g_int_TDoEmp, modmip_g_str_NDoEmp, l_int_RelAcc, l_int_AccTDo, l_str_AccNDo, l_int_AccVin)

   If l_int_RelAcc > 0 Then
      pnl_RelAcc.Visible = True
      
      If l_int_AccVin = 1 Then
         pnl_RelAcc.Caption = "Empresa es Accionista"
      ElseIf l_int_AccVin = 2 Then
         pnl_RelAcc.Caption = "Relación con Accionista (" & modmip_gf_Consulta_NomAcc(l_int_AccTDo, l_str_AccNDo) & ")"
      End If
   End If
   
   Screen.MousePointer = 0
End Sub

Private Sub txt_CodPos_GotFocus()
   Call gs_SelecTodo(txt_CodPos)
End Sub

Private Sub txt_CodPos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telef1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_Direcc_GotFocus()
   Call gs_SelecTodo(txt_Direcc)
End Sub

Private Sub txt_Direcc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_PrvEst)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_IntDpt_GotFocus()
   Call gs_SelecTodo(txt_IntDpt)
End Sub

Private Sub txt_IntDpt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipZon)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_NomCom_GotFocus()
   Call gs_SelecTodo(txt_NomCom)
End Sub

Private Sub txt_NomCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CodCiu)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ._()=/&%$@")
   End If
End Sub

Private Sub txt_NomVia_GotFocus()
   Call gs_SelecTodo(txt_NomVia)
End Sub

Private Sub txt_NomVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumVia)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_NomZon_GotFocus()
   Call gs_SelecTodo(txt_NomZon)
End Sub

Private Sub txt_NomZon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_DptDir)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_NumVia_GotFocus()
   Call gs_SelecTodo(txt_NumVia)
End Sub

Private Sub txt_NumVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_IntDpt)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_Refere_GotFocus()
   Call gs_SelecTodo(txt_Refere)
End Sub

Private Sub txt_Refere_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telef1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_RazSoc_GotFocus()
   Call gs_SelecTodo(txt_RazSoc)
End Sub
Private Sub txt_RazSoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NomCom)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- )(/=.,;:_$%&@#")
   End If
End Sub

Private Sub txt_Telef1_GotFocus()
   Call gs_SelecTodo(txt_Telef1)
End Sub

Private Sub txt_Telef1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telef2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Telef2_GotFocus()
   Call gs_SelecTodo(txt_Telef2)
End Sub

Private Sub txt_Telef2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumFax)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_NumFax_GotFocus()
   Call gs_SelecTodo(txt_NumFax)
End Sub

Private Sub txt_NumFax_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_TeleRH)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-()")
   End If
End Sub

Private Sub txt_TeleRH_GotFocus()
   Call gs_SelecTodo(txt_TeleRH)
End Sub

Private Sub txt_TeleRH_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_AnexRH)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_AnexRH_GotFocus()
   Call gs_SelecTodo(txt_AnexRH)
End Sub

Private Sub txt_AnexRH_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_PagWeb)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_PagWeb_GotFocus()
   Call gs_SelecTodo(txt_PagWeb)
End Sub

Private Sub txt_PagWeb_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "@_.-")
   End If
End Sub

Private Sub fs_Inicio()
   Call moddat_gs_Carga_CdCIIU(cmb_CodCiu)

   Call moddat_gs_Carga_LisIte_Combo(cmb_TipVia, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipZon, 1, "202")
   
   Call moddat_gs_Carga_LisIte(cmb_PaiRes, l_arr_PaiRes, 1, "500")
      
   Call moddat_gs_Carga_Depart(cmb_DptDir)
End Sub
   
Private Sub fs_Activa(ByVal p_Habilita As Integer)
   txt_RazSoc.Enabled = p_Habilita
   txt_NomCom.Enabled = p_Habilita
   cmb_CodCiu.Enabled = p_Habilita
   cmb_PaiRes.Enabled = p_Habilita
   cmb_TipVia.Enabled = p_Habilita
   txt_NomVia.Enabled = p_Habilita
   txt_NumVia.Enabled = p_Habilita
   txt_IntDpt.Enabled = p_Habilita
   cmb_TipZon.Enabled = p_Habilita
   txt_NomZon.Enabled = p_Habilita
   cmb_DptDir.Enabled = p_Habilita
   cmb_PrvDir.Enabled = p_Habilita
   cmb_DstDir.Enabled = p_Habilita
   txt_Refere.Enabled = p_Habilita
   txt_Direcc.Enabled = p_Habilita
   cmb_PrvEst.Enabled = p_Habilita
   txt_CodPos.Enabled = p_Habilita
   txt_Telef1.Enabled = p_Habilita
   txt_Telef2.Enabled = p_Habilita
   txt_NumFax.Enabled = p_Habilita
   txt_TeleRH.Enabled = p_Habilita
   txt_AnexRH.Enabled = p_Habilita
   txt_PagWeb.Enabled = p_Habilita
   
   cmd_Grabar.Enabled = p_Habilita
   cmd_Cancel.Enabled = p_Habilita
   
   cmd_Editar.Enabled = Not p_Habilita
End Sub

Private Sub fs_Limpia()
   txt_RazSoc.Text = ""
   txt_NomCom.Text = ""
   cmb_CodCiu.ListIndex = -1
   
   cmb_PaiRes.ListIndex = -1
   
   cmb_TipVia.ListIndex = -1
   txt_NomVia.Text = ""
   txt_NumVia.Text = ""
   txt_IntDpt.Text = ""
   cmb_TipZon.ListIndex = -1
   txt_NomZon.Text = ""
   cmb_DptDir.ListIndex = -1
   cmb_PrvDir.Clear
   cmb_DstDir.Clear
   txt_Refere.Text = ""
   cmb_TipVia.Enabled = False
   txt_NomVia.Enabled = False
   txt_NumVia.Enabled = False
   txt_IntDpt.Enabled = False
   cmb_TipZon.Enabled = False
   txt_NomZon.Enabled = False
   cmb_DptDir.Enabled = False
   cmb_PrvDir.Enabled = False
   cmb_DstDir.Enabled = False
   txt_Refere.Enabled = False
   
   txt_Direcc.Text = ""
   cmb_PrvEst.Clear
   txt_CodPos.Text = ""
   
   txt_Direcc.Enabled = False
   cmb_PrvEst.Enabled = False
   txt_CodPos.Enabled = False
   
   txt_Telef1.Text = ""
   txt_Telef2.Text = ""
   txt_NumFax.Text = ""
   txt_TeleRH.Text = ""
   txt_AnexRH.Text = ""
   txt_PagWeb.Text = ""
End Sub

Private Sub fs_Cargar_Datos()
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(modmip_g_int_TDoEmp) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & modmip_g_str_NDoEmp & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      txt_RazSoc.Text = Trim(g_rst_Princi!DATGEN_RAZSOC & "")
      txt_NomCom.Text = Trim(g_rst_Princi!DATGEN_NOMCOM & "")
      
      Call gs_BuscarCombo_Item(cmb_CodCiu, g_rst_Princi!DATGEN_CODCIU)
      
      If Not IsNull(g_rst_Princi!DATGEN_PAIRES) Then
        cmb_PaiRes.ListIndex = gf_Busca_Arregl(l_arr_PaiRes, g_rst_Princi!DATGEN_PAIRES) - 1
      End If
      
      If l_arr_PaiRes(cmb_PaiRes.ListIndex + 1).Genera_Codigo = "004028" Then
         Call gs_BuscarCombo_Item(cmb_TipVia, g_rst_Princi!DatGen_TipVia)
         txt_NomVia.Text = Trim(g_rst_Princi!DatGen_NomVia & "")
         txt_NumVia.Text = Trim(g_rst_Princi!DatGen_numVia & "")
         txt_IntDpt.Text = Trim(g_rst_Princi!DATGEN_INTDPT & "")
         
         Call gs_BuscarCombo_Item(cmb_TipZon, g_rst_Princi!DatGen_TipZon)
         txt_NomZon.Text = Trim(g_rst_Princi!DatGen_NomZon & "")
         
         If CLng(g_rst_Princi!DatGen_Ubigeo) > 0 Then
            Call gs_BuscarCombo_Item(cmb_DptDir, CInt(Left(g_rst_Princi!DatGen_Ubigeo, 2)))
            Call moddat_gs_Carga_Provin(cmb_PrvDir, Left(g_rst_Princi!DatGen_Ubigeo, 2))
            Call gs_BuscarCombo_Item(cmb_PrvDir, CInt(Mid(g_rst_Princi!DatGen_Ubigeo, 3, 2)))
            Call moddat_gs_Carga_Distri(cmb_DstDir, Left(g_rst_Princi!DatGen_Ubigeo, 2), Mid(g_rst_Princi!DatGen_Ubigeo, 3, 2))
            Call gs_BuscarCombo_Item(cmb_DstDir, CInt(Right(g_rst_Princi!DatGen_Ubigeo, 2)))
         End If
         txt_Refere.Text = Trim(g_rst_Princi!DATGEN_REFERE & "")
      Else
         txt_Direcc.Text = Trim(g_rst_Princi!DATGEN_EXTDIR & "")
         
         Call modmip_gs_Carga_CiuExt(l_arr_PrvEst, cmb_PrvEst, l_arr_PaiRes(cmb_PaiRes.ListIndex + 1).Genera_Codigo)
         If Not IsNull(g_rst_Princi!DATGEN_EXTCIU) Then
            cmb_PrvEst.ListIndex = gf_Busca_Arregl(l_arr_PrvEst, g_rst_Princi!DATGEN_EXTCIU) - 1
         End If
         txt_CodPos.Text = Trim(g_rst_Princi!DATGEN_EXTCPO & "")
      End If
      
      txt_Telef1.Text = Trim(g_rst_Princi!DATGEN_TELEF1 & "")
      txt_Telef2.Text = Trim(g_rst_Princi!DATGEN_TELEF2 & "")
      txt_NumFax.Text = Trim(g_rst_Princi!DATGEN_NUMFAX & "")
      
      txt_TeleRH.Text = Trim(g_rst_Princi!DATGEN_TELERH & "")
      txt_AnexRH.Text = Trim(g_rst_Princi!DATGEN_ANEXRH & "")
      
      txt_PagWeb.Text = Trim(g_rst_Princi!DATGEN_PAGWEB & "")
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub


