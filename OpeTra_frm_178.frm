VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_MntCli_66 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2430
   ClientLeft      =   3585
   ClientTop       =   5265
   ClientWidth     =   11670
   Icon            =   "OpeTra_frm_178.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2445
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11685
      _Version        =   65536
      _ExtentX        =   20611
      _ExtentY        =   4313
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
      Begin Threed.SSPanel SSPanel9 
         Height          =   435
         Left            =   30
         TabIndex        =   1
         Top             =   1950
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
         Begin VB.ComboBox cmb_ActPri 
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   60
            Width           =   9615
         End
         Begin VB.Label Label19 
            Caption         =   "Actividad Económica:"
            Height          =   315
            Left            =   90
            TabIndex        =   3
            Top             =   60
            Width           =   1605
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   4
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
            Height          =   255
            Left            =   660
            TabIndex        =   5
            Top             =   60
            Width           =   6165
            _Version        =   65536
            _ExtentX        =   10874
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Mantenimiento de Clientes"
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
            Height          =   255
            Left            =   660
            TabIndex        =   6
            Top             =   330
            Width           =   6165
            _Version        =   65536
            _ExtentX        =   10874
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Actividades Económicas"
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
            Picture         =   "OpeTra_frm_178.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   7
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
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   1950
            TabIndex        =   8
            Top             =   60
            Width           =   9615
            _Version        =   65536
            _ExtentX        =   16960
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1 - 07522154 / IKEHARA PUNK MIGUEL ANGEL"
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
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   90
            TabIndex        =   9
            Top             =   60
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   30
         TabIndex        =   10
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10980
            Picture         =   "OpeTra_frm_178.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Datos del Crédito"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_178.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Cargar Plantilla para Ingreso de Datos"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_MntCli_66"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_ActPri_Click()
   Call gs_SetFocus(cmd_Buscar)
End Sub

Private Sub cmb_ActPri_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_ActPri_Click
   End If
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_ActPri.ListIndex = -1 Then
      MsgBox "Debe seleccionar una Actividad Económica.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_ActPri)
      
      Exit Sub
   End If

   If modmip_g_int_PaiRes = 4028 Then
      Select Case cmb_ActPri.ItemData(cmb_ActPri.ListIndex)
         Case 11:    frm_MntCli_57.Show 1
         Case 21:    frm_MntCli_58.Show 1
         Case 31:    frm_MntCli_59.Show 1
         Case 41:    frm_MntCli_60.Show 1
         Case 51:    frm_MntCli_61.Show 1
         Case 61:    frm_MntCli_62.Show 1
      End Select
   Else
      Select Case cmb_ActPri.ItemData(cmb_ActPri.ListIndex)
         Case 11:    frm_MntCli_63.Show 1
         Case 21:    frm_MntCli_64.Show 1
         Case 31:    frm_MntCli_65.Show 1
         Case 61:    frm_MntCli_62.Show 1
      End Select
   End If
   
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   If modmip_g_int_TipCli = 1 Then
      pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   Else
      pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli & " (" & CStr(moddat_g_int_CygTDo) & " - " & moddat_g_str_CygNDo & " / " & moddat_g_str_CygNom & ")"
   End If
   
   Call fs_Inicio
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicio()
   If modmip_g_int_PaiRes = 4028 Then
      Call moddat_gs_Carga_LisIte_Combo(cmb_ActPri, 1, "008")
   Else
      Call moddat_gs_Carga_LisIte_Combo(cmb_ActPri, 1, "114")
   End If
End Sub

