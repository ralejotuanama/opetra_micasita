VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frm_Gar_CreHip_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   375
   ClientTop       =   2760
   ClientWidth     =   14550
   Icon            =   "OpeTra_frm_009.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   14550
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   5415
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   14535
      _Version        =   65536
      _ExtentX        =   25638
      _ExtentY        =   9551
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
         Height          =   1515
         Left            =   30
         TabIndex        =   26
         Top             =   3030
         Width           =   14445
         _Version        =   65536
         _ExtentX        =   25479
         _ExtentY        =   2672
         _StockProps     =   15
         BackColor       =   13160660
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisHip 
            Height          =   1125
            Left            =   30
            TabIndex        =   7
            Top             =   360
            Width           =   14385
            _ExtentX        =   25374
            _ExtentY        =   1984
            _Version        =   393216
            Rows            =   21
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   60
            TabIndex        =   27
            Top             =   60
            Width           =   2445
            _Version        =   65536
            _ExtentX        =   4313
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Bien en Garantía"
            ForeColor       =   16777215
            BackColor       =   32768
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
            Left            =   7620
            TabIndex        =   28
            Top             =   60
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Inscripción"
            ForeColor       =   16777215
            BackColor       =   32768
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
            Left            =   10260
            TabIndex        =   29
            Top             =   60
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Importe"
            ForeColor       =   16777215
            BackColor       =   32768
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
         Begin Threed.SSPanel SSPanel12 
            Height          =   285
            Left            =   11850
            TabIndex        =   30
            Top             =   60
            Width           =   2235
            _Version        =   65536
            _ExtentX        =   3942
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Situación"
            ForeColor       =   16777215
            BackColor       =   32768
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
         Begin Threed.SSPanel SSPanel14 
            Height          =   285
            Left            =   8940
            TabIndex        =   32
            Top             =   60
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Constitución"
            ForeColor       =   16777215
            BackColor       =   32768
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   285
            Left            =   2490
            TabIndex        =   43
            Top             =   60
            Width           =   5145
            _Version        =   65536
            _ExtentX        =   9075
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Documento Registral"
            ForeColor       =   16777215
            BackColor       =   32768
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   12
         Top             =   30
         Width           =   14445
         _Version        =   65536
         _ExtentX        =   25479
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
            TabIndex        =   13
            Top             =   60
            Width           =   4965
            _Version        =   65536
            _ExtentX        =   8758
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Gestión de Garantías"
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
            Left            =   60
            Picture         =   "OpeTra_frm_009.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel20 
         Height          =   765
         Left            =   30
         TabIndex        =   14
         Top             =   750
         Width           =   14445
         _Version        =   65536
         _ExtentX        =   25479
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
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   12360
            Picture         =   "OpeTra_frm_009.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Buscar Registros"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   13050
            Picture         =   "OpeTra_frm_009.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   13740
            Picture         =   "OpeTra_frm_009.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin VB.ComboBox cmb_TipBus 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   2775
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   7530
            MaxLength       =   12
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   390
            Width           =   2775
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   7530
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   60
            Width           =   2775
         End
         Begin MSMask.MaskEdBox msk_NumOpe 
            Height          =   315
            Left            =   1620
            TabIndex        =   3
            Top             =   390
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            Mask            =   "###-##-#####"
            PromptChar      =   " "
         End
         Begin VB.Label Label17 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   60
            TabIndex        =   19
            Top             =   1740
            Width           =   1065
         End
         Begin VB.Label Label18 
            Caption         =   "Tipo de Búsqueda:"
            Height          =   315
            Left            =   90
            TabIndex        =   18
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label19 
            Caption         =   "Nro. Doc. Ident.:"
            Height          =   285
            Left            =   6150
            TabIndex        =   17
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Doc. Ident.:"
            Height          =   315
            Left            =   6150
            TabIndex        =   16
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label lbl_Numero 
            Caption         =   "Nro. Operación:"
            Height          =   285
            Left            =   90
            TabIndex        =   15
            Top             =   390
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1425
         Left            =   30
         TabIndex        =   20
         Top             =   1560
         Width           =   14445
         _Version        =   65536
         _ExtentX        =   25479
         _ExtentY        =   2514
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
            Left            =   1620
            TabIndex        =   21
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
         Begin Threed.SSPanel pnl_Situac 
            Height          =   315
            Left            =   3210
            TabIndex        =   22
            Top             =   60
            Width           =   4035
            _Version        =   65536
            _ExtentX        =   7117
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "CREDITO VIGENTE"
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
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   315
            Left            =   10440
            TabIndex        =   23
            Top             =   720
            Width           =   3945
            _Version        =   65536
            _ExtentX        =   6959
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   1620
            TabIndex        =   33
            Top             =   390
            Width           =   7155
            _Version        =   65536
            _ExtentX        =   12621
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
         Begin Threed.SSPanel pnl_Direcc 
            Height          =   660
            Left            =   1620
            TabIndex        =   34
            Top             =   720
            Width           =   7155
            _Version        =   65536
            _ExtentX        =   12621
            _ExtentY        =   1164
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
         Begin Threed.SSPanel pnl_MtoPre 
            Height          =   315
            Left            =   10440
            TabIndex        =   35
            Top             =   1050
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "99999.99 "
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   10440
            TabIndex        =   36
            Top             =   60
            Width           =   3945
            _Version        =   65536
            _ExtentX        =   6959
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "DOLARES AMERICANOS"
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
         Begin Threed.SSPanel pnl_Modali 
            Height          =   315
            Left            =   10440
            TabIndex        =   37
            Top             =   390
            Width           =   3945
            _Version        =   65536
            _ExtentX        =   6959
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "DOLARES AMERICANOS"
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
         Begin VB.Label Label4 
            Caption         =   "Cliente Titular:"
            Height          =   315
            Left            =   60
            TabIndex        =   42
            Top             =   390
            Width           =   1395
         End
         Begin VB.Label Label6 
            Caption         =   "Dirección Inmueble:"
            Height          =   405
            Left            =   60
            TabIndex        =   41
            Top             =   720
            Width           =   1395
         End
         Begin VB.Label Label8 
            Caption         =   "Monto Préstamo:"
            Height          =   315
            Left            =   9060
            TabIndex        =   40
            Top             =   1080
            Width           =   1245
         End
         Begin VB.Label Label20 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   9060
            TabIndex        =   39
            Top             =   60
            Width           =   855
         End
         Begin VB.Label Label21 
            Caption         =   "Modalidad:"
            Height          =   315
            Left            =   9060
            TabIndex        =   38
            Top             =   390
            Width           =   945
         End
         Begin VB.Label Label3 
            Caption         =   "Nro. de Operación:"
            Height          =   315
            Left            =   60
            TabIndex        =   25
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label Label2 
            Caption         =   "Moneda:"
            Height          =   315
            Left            =   9060
            TabIndex        =   24
            Top             =   720
            Width           =   945
         End
      End
      Begin Threed.SSPanel SSPanel13 
         Height          =   765
         Left            =   30
         TabIndex        =   31
         Top             =   4590
         Width           =   14445
         _Version        =   65536
         _ExtentX        =   25479
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
         Begin VB.CommandButton cmd_DatInm 
            Height          =   675
            Left            =   13740
            Picture         =   "OpeTra_frm_009.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Datos del Inmueble"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_NueHip 
            Height          =   675
            Left            =   12360
            Picture         =   "OpeTra_frm_009.frx":1636
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Registrar Hipotecas"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_LevHip 
            Height          =   675
            Left            =   13050
            Picture         =   "OpeTra_frm_009.frx":1F00
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Liberar Hipoteca"
            Top             =   30
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frm_Gar_CreHip_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_TipBus_Click()
   If cmb_TipBus.ListIndex > -1 Then
      If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
         cmb_TipDoc.Enabled = True
         txt_NumDoc.Enabled = True
         msk_NumOpe.Enabled = False
         
         msk_NumOpe.Mask = ""
         msk_NumOpe.Text = ""
         msk_NumOpe.Mask = "###-##-#####"
         
         Call gs_SetFocus(cmb_TipDoc)
      Else
         cmb_TipDoc.Enabled = False
         txt_NumDoc.Enabled = False
         msk_NumOpe.Enabled = True
         
         cmb_TipDoc.ListIndex = -1
         txt_NumDoc.Text = ""
         
         Call gs_SetFocus(msk_NumOpe)
      End If
   Else
      cmb_TipDoc.Enabled = False
      txt_NumDoc.Enabled = False
      
      msk_NumOpe.Enabled = False
   
      cmb_TipDoc.ListIndex = -1
      txt_NumDoc.Text = ""
      msk_NumOpe.Mask = ""
      msk_NumOpe.Text = ""
      msk_NumOpe.Mask = "###-##-#####"
   End If
End Sub

Private Sub cmb_TipBus_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipBus_Click
   End If
End Sub


Private Sub cmb_TipDoc_Click()
   If cmb_TipDoc.ListIndex > -1 Then
      Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
         Case 1:  txt_NumDoc.MaxLength = 8
         Case 2:  txt_NumDoc.MaxLength = 12
         Case 3:  txt_NumDoc.MaxLength = 12
      End Select
   End If
   Call gs_SetFocus(txt_NumDoc)
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Click
   End If
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_TipBus.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Búsqueda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipBus)
      Exit Sub
   End If
   
   If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
      If cmb_TipDoc.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipDoc)
         Exit Sub
      End If
      
      If Len(Trim(txt_NumDoc.Text)) = 0 Then
         MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         Exit Sub
      End If
      
      If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1 Then
         txt_NumDoc.Text = Format(txt_NumDoc.Text, "00000000")
      End If
      
      moddat_g_int_TipDoc = cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
      moddat_g_str_TipDoc = cmb_TipDoc.Text
      moddat_g_str_NumDoc = txt_NumDoc.Text
   Else
      If Len(Trim(msk_NumOpe.Text)) < 10 Then
         MsgBox "Debe ingresar el Número de Operación.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(msk_NumOpe)
         Exit Sub
      End If
      
      moddat_g_str_NumOpe = msk_NumOpe.Text
   End If
   
   If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
      g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
      g_str_Parame = g_str_Parame & "HIPMAE_TDOCLI = " & CStr(moddat_g_int_TipDoc) & " AND "
      g_str_Parame = g_str_Parame & "HIPMAE_NDOCLI = '" & moddat_g_str_NumDoc & "' AND "
      g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 "
      
   Else
      g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
      g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
      g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No existe ninguna Operación para amortizar para la Búsqueda deseada. ", vbExclamation, modgen_g_str_NomPlt
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   Screen.MousePointer = 11

   Call fs_Buscar_DatGen

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   pnl_NumOpe.Caption = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
   pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   pnl_Moneda.Caption = moddat_g_str_Moneda
   pnl_Produc.Caption = moddat_g_str_NomPrd
   pnl_Modali.Caption = moddat_g_str_DesMod
   pnl_MtoPre.Caption = Format(moddat_g_dbl_MtoPre, "###,###,##0.00") & " "
   pnl_Situac.Caption = moddat_g_str_Situac
   
   pnl_Direcc.Caption = moddat_g_str_Direcc & Chr(10) & Chr(13) & moddat_g_str_Distri
   
   
   Call fs_Activa(False)
   
   Call fs_Buscar_Hipotecas
   
   If grd_LisHip.Rows > 0 Then
      cmd_NueHip.Enabled = False
      cmd_DatInm.Enabled = False
   End If
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_DatInm_Click()
   moddat_g_int_FlgAct = 1
   
   frm_Gar_CreHip_02.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Call moddat_gs_Consulta_DatInm(moddat_g_str_NumSol, moddat_g_str_Direcc, moddat_g_str_Distri)
      
      pnl_Direcc.Caption = moddat_g_str_Direcc & Chr(10) & Chr(13) & moddat_g_str_Distri
   End If
End Sub

Private Sub cmd_LevHip_Click()
   Dim r_str_BieGar     As String
   Dim r_int_Situac     As Integer
   Dim r_dbl_MtoHip     As Double
   Dim r_str_Descri     As String
   Dim r_str_SitCre     As String
   Dim r_str_SitAnt     As String
   Dim r_str_Operac     As String
   
   If grd_LisHip.Rows = 0 Then
      Exit Sub
   End If
   
   grd_LisHip.Col = 1
   r_str_Descri = grd_LisHip.Text
   
   grd_LisHip.Col = 4
   r_dbl_MtoHip = CDbl(grd_LisHip.Text)
   
   grd_LisHip.Col = 6
   r_str_BieGar = grd_LisHip.Text
   
   grd_LisHip.Col = 7
   r_int_Situac = CInt(grd_LisHip.Text)
   
   Call gs_RefrescaGrid(grd_LisHip)
   
   If r_int_Situac = 2 Then
      MsgBox "Esta Hipoteca ya ha sido liberada.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(grd_LisHip)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   'Buscando Información en Maestro de Creditos
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      r_str_SitCre = CStr(g_rst_Genera!HIPMAE_SITCRE)
      r_str_SitAnt = CStr(g_rst_Genera!HIPMAE_SITANT)
   End If

   g_rst_Genera.Close
   Set g_rst_Genera = Nothing

   
   r_str_Operac = moddat_gf_Consulta_Operac(moddat_g_str_CodPrd, "42")
   
   'Grabando en Garantias
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CRE_HIPGAR_LIBERA ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
      g_str_Parame = g_str_Parame & r_str_BieGar & ", "
            
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
         If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   'Registrando Archivo para Contabilización
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CREDITO_MOV_CONTAB_INSERTA ("
      g_str_Parame = g_str_Parame & "'" & r_str_SitCre & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_SitAnt & "', "
      g_str_Parame = g_str_Parame & "1,"
      g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & "'" & Left(r_str_Operac, 3) & "002', "
      g_str_Parame = g_str_Parame & "'002', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
      g_str_Parame = g_str_Parame & CStr(r_dbl_MtoHip) & ", "
      
      g_str_Parame = g_str_Parame & "'" & Mid(r_str_Descri, 1, 50) & "')"
         
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
   
   Screen.MousePointer = 11
   Call fs_Buscar_Hipotecas
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call gs_SetFocus(cmb_TipBus)
End Sub

Private Sub cmd_NueHip_Click()
   moddat_g_int_FlgAct = 1
   
   frm_Gar_CreHip_03.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar_Hipotecas
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call cmd_Limpia_Click
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub grd_LisHip_SelChange()
   If grd_LisHip.Rows > 2 Then
      grd_LisHip.RowSel = grd_LisHip.Row
   End If
End Sub

Private Sub msk_NumOpe_GotFocus()
   Call gs_SelecTodo(msk_NumOpe)
End Sub

Private Sub msk_NumOpe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   Else
      If cmb_TipDoc.ListIndex > -1 Then
         Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
            Case 1:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case 2:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 3:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
         End Select
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub fs_Inicia()
   Call modsis_gs_Carga_TipBus_1(cmb_TipBus)
   Call moddat_gs_Carga_TipDocIde(cmb_TipDoc, 1)

   grd_LisHip.ColWidth(0) = 2430
   grd_LisHip.ColWidth(1) = 5130
   grd_LisHip.ColWidth(2) = 1305
   grd_LisHip.ColWidth(3) = 1305
   grd_LisHip.ColWidth(4) = 1575
   grd_LisHip.ColWidth(5) = 2205
   grd_LisHip.ColWidth(6) = 0
   grd_LisHip.ColWidth(7) = 0
   
   grd_LisHip.ColAlignment(0) = flexAlignCenterCenter
   grd_LisHip.ColAlignment(1) = flexAlignLeftCenter
   grd_LisHip.ColAlignment(2) = flexAlignCenterCenter
   grd_LisHip.ColAlignment(3) = flexAlignCenterCenter
   grd_LisHip.ColAlignment(4) = flexAlignRightCenter
   grd_LisHip.ColAlignment(5) = flexAlignLeftCenter
End Sub

Private Sub fs_Limpia()
   Call fs_Activa(True)
   
   cmb_TipBus.ListIndex = -1
   cmb_TipDoc.Enabled = False
   txt_NumDoc.Enabled = False
   msk_NumOpe.Enabled = False

   msk_NumOpe.Mask = ""
   msk_NumOpe.Text = ""
   msk_NumOpe.Mask = "###-##-#####"
   
   txt_NumDoc.Text = ""
   
   pnl_NumOpe.Caption = ""
   pnl_Situac.Caption = ""
   pnl_NomCli.Caption = ""
   
   pnl_Produc.Caption = ""
   pnl_Modali.Caption = ""
   pnl_Moneda.Caption = ""
   pnl_Direcc.Caption = ""
   pnl_MtoPre.Caption = "0.00 "
   
   Call gs_LimpiaGrid(grd_LisHip)
End Sub

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   cmb_TipBus.Enabled = p_Habilita
   cmb_TipDoc.Enabled = p_Habilita
   txt_NumDoc.Enabled = p_Habilita
   msk_NumOpe.Enabled = p_Habilita
   cmd_Buscar.Enabled = p_Habilita
   
   grd_LisHip.Enabled = Not p_Habilita
   cmd_NueHip.Enabled = Not p_Habilita
   cmd_LevHip.Enabled = Not p_Habilita
   cmd_DatInm.Enabled = Not p_Habilita
End Sub

Private Sub fs_Buscar_DatGen()
   g_rst_Princi.MoveFirst
   
   moddat_g_int_TipDoc = g_rst_Princi!HIPMAE_TDOCLI
   moddat_g_str_NumDoc = Trim(g_rst_Princi!HIPMAE_NDOCLI)
   moddat_g_str_NumSol = Trim(g_rst_Princi!HIPMAE_NUMSOL)
   moddat_g_str_NumOpe = Trim(g_rst_Princi!HIPMAE_NUMOPE)
   
   'Obteniendo Nombre de Cliente
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   'Obteniendo Nombre y DOI de Cónyuge
   moddat_g_int_CygTDo = g_rst_Princi!HIPMAE_TDOCYG
   
   If moddat_g_int_CygTDo > 0 Then
      moddat_g_str_CygNDo = Trim(g_rst_Princi!HIPMAE_NDOCYG & "")
      moddat_g_str_CygNom = moddat_gf_Buscar_NomCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo)
   End If
   
   'Obteniendo Descripción de Producto
   moddat_g_str_CodPrd = Trim(g_rst_Princi!HIPMAE_CODPRD)
   moddat_g_str_NomPrd = moddat_gf_Consulta_Produc(Trim(g_rst_Princi!HIPMAE_CODPRD))

   'Obeniendo Modalidad de Producto
   moddat_g_str_CodMod = Trim(g_rst_Princi!HIPMAE_CODMOD)
   moddat_g_str_DesMod = moddat_gf_Buscar_NomMod(Trim(g_rst_Princi!HIPMAE_CODPRD), moddat_g_str_CodMod)

   'Moneda
   moddat_g_str_Moneda = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPMAE_MONEDA))
   moddat_g_int_TipMon = g_rst_Princi!HIPMAE_MONEDA
   
   'Monto Préstamo
   moddat_g_dbl_MtoPre = g_rst_Princi!HIPMAE_MTOPRE

   'Situación de Crédito
   moddat_g_int_Situac = g_rst_Princi!HIPMAE_SITUAC
   'moddat_g_str_Situac = moddat_gf_Consulta_SitCre(moddat_g_int_Situac)
   
   'Obteniendo Información del Inmueble
   Call moddat_gs_Consulta_DatInm(moddat_g_str_NumSol, moddat_g_str_Direcc, moddat_g_str_Distri)
End Sub

Private Sub fs_Buscar_Fianzas()

End Sub

Private Sub fs_Buscar_Hipotecas()
   Call gs_LimpiaGrid(grd_LisHip)
   
   g_str_Parame = "SELECT * FROM CRE_HIPGAR WHERE "
   g_str_Parame = g_str_Parame & "HIPGAR_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "ORDER BY HIPGAR_BIEGAR ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_LisHip.Redraw = False
      
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         grd_LisHip.Rows = grd_LisHip.Rows + 1
         grd_LisHip.Row = grd_LisHip.Rows - 1
         
         grd_LisHip.Col = 0
         grd_LisHip.Text = moddat_gf_Consulta_ParDes("030", CStr(g_rst_Princi!HIPGAR_BIEGAR))
      
         grd_LisHip.Col = 1
         
         Select Case g_rst_Princi!HIPGAR_TDOREG
            Case 1, 2:  grd_LisHip.Text = Trim(moddat_gf_Consulta_ParDes("026", CStr(g_rst_Princi!HIPGAR_TDOREG))) & " NRO. " & Trim(g_rst_Princi!HIPGAR_PARFIC) & " ASIENTO NRO. " & Trim(g_rst_Princi!HIPGAR_NUMASI)
            Case 3:     grd_LisHip.Text = "TOMO NRO. " & Trim(g_rst_Princi!HIPGAR_NUMTOM) & " FOJA NRO. " & Trim(g_rst_Princi!HIPGAR_NUMFOJ) & " LIBRO NRO. " & Trim(g_rst_Princi!HIPGAR_NUMLIB)
         End Select
      
         grd_LisHip.Col = 2
         grd_LisHip.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPGAR_FECINS))
         
         grd_LisHip.Col = 3
         grd_LisHip.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPGAR_FECCON))
         
         grd_LisHip.Col = 4
         grd_LisHip.Text = Format(g_rst_Princi!HIPGAR_MTOHIP, "###,###,##0.00")
      
         grd_LisHip.Col = 5
         grd_LisHip.Text = moddat_gf_Consulta_ParDes("031", CStr(g_rst_Princi!HIPGAR_SITUAC))
         
         grd_LisHip.Col = 6
         grd_LisHip.Text = CStr(g_rst_Princi!HIPGAR_BIEGAR)
         
         grd_LisHip.Col = 7
         grd_LisHip.Text = CStr(g_rst_Princi!HIPGAR_SITUAC)
         
         g_rst_Princi.MoveNext
      Loop
   
      grd_LisHip.Redraw = True
      Call gs_UbiIniGrid(grd_LisHip)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

