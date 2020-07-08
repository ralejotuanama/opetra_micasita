VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_MntEmp_53 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   5430
   ClientLeft      =   7530
   ClientTop       =   3315
   ClientWidth     =   11685
   Icon            =   "OpeTra_frm_174.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11685
      _Version        =   65536
      _ExtentX        =   20611
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
      Begin Threed.SSPanel SSPanel5 
         Height          =   1125
         Left            =   30
         TabIndex        =   25
         Top             =   4230
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
         _ExtentY        =   1984
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
            TabIndex        =   31
            Text            =   "Text1"
            Top             =   390
            Width           =   1640
         End
         Begin VB.TextBox txt_AnexRH 
            Height          =   315
            Left            =   3630
            MaxLength       =   5
            TabIndex        =   30
            Text            =   "Text1"
            Top             =   390
            Width           =   1640
         End
         Begin VB.TextBox txt_Telef1 
            Height          =   315
            Left            =   1980
            MaxLength       =   25
            TabIndex        =   29
            Text            =   "Text1"
            Top             =   60
            Width           =   1640
         End
         Begin VB.TextBox txt_NumFax 
            Height          =   315
            Left            =   8190
            MaxLength       =   25
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   60
            Width           =   1640
         End
         Begin VB.TextBox txt_Telef2 
            Height          =   315
            Left            =   3630
            MaxLength       =   25
            TabIndex        =   27
            Text            =   "Text1"
            Top             =   60
            Width           =   1640
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Left            =   1980
            MaxLength       =   120
            TabIndex        =   26
            Text            =   "Text1"
            Top             =   720
            Width           =   3315
         End
         Begin VB.Label lbl_General 
            Caption         =   "Teléfono/Anexo RR.HH:"
            Height          =   285
            Index           =   47
            Left            =   60
            TabIndex        =   35
            Top             =   390
            Width           =   1815
         End
         Begin VB.Label lbl_General 
            Caption         =   "Teléfono (s):"
            Height          =   285
            Index           =   46
            Left            =   60
            TabIndex        =   34
            Top             =   60
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fax:"
            Height          =   285
            Index           =   55
            Left            =   6180
            TabIndex        =   33
            Top             =   60
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Página Web:"
            Height          =   285
            Index           =   0
            Left            =   60
            TabIndex        =   32
            Top             =   720
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1095
         Left            =   30
         TabIndex        =   1
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
         Begin VB.TextBox txt_RazSoc 
            Height          =   315
            Left            =   1980
            MaxLength       =   250
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   60
            Width           =   9525
         End
         Begin VB.TextBox txt_NomCom 
            Height          =   315
            Left            =   1980
            MaxLength       =   250
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   390
            Width           =   9525
         End
         Begin VB.ComboBox cmb_CodCiu 
            Height          =   315
            Left            =   1980
            TabIndex        =   2
            Text            =   "cmb_DptDir"
            Top             =   720
            Width           =   9525
         End
         Begin VB.Label lbl_General 
            Caption         =   "CIIU:"
            Height          =   285
            Index           =   39
            Left            =   60
            TabIndex        =   7
            Top             =   720
            Width           =   1365
         End
         Begin VB.Label lbl_General 
            Caption         =   "Razón Social:"
            Height          =   285
            Index           =   37
            Left            =   60
            TabIndex        =   6
            Top             =   60
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Comercial:"
            Height          =   285
            Index           =   49
            Left            =   60
            TabIndex        =   5
            Top             =   390
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   8
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
            Left            =   660
            TabIndex        =   9
            Top             =   60
            Width           =   6165
            _Version        =   65536
            _ExtentX        =   10874
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
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_174.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   10
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
            Left            =   1980
            TabIndex        =   11
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
         Begin VB.Label Label1 
            Caption         =   "Documento Identidad:"
            Height          =   315
            Left            =   60
            TabIndex        =   12
            Top             =   60
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   675
         Left            =   30
         TabIndex        =   13
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
            Picture         =   "OpeTra_frm_174.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_174.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   1095
         Left            =   30
         TabIndex        =   16
         Top             =   3090
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
         Begin VB.TextBox txt_NomVia 
            Height          =   315
            Left            =   1980
            MaxLength       =   250
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   60
            Width           =   9555
         End
         Begin VB.ComboBox cmb_TipZon 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   390
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DptDir 
            Height          =   315
            Left            =   8190
            TabIndex        =   18
            Text            =   "cmb_DptDir"
            Top             =   420
            Width           =   3315
         End
         Begin VB.TextBox txt_Refere 
            Height          =   315
            Left            =   1980
            MaxLength       =   250
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label20 
            Caption         =   "Dirección:"
            Height          =   285
            Left            =   60
            TabIndex        =   24
            Top             =   60
            Width           =   1485
         End
         Begin VB.Label Label22 
            Caption         =   "País:"
            Height          =   315
            Left            =   60
            TabIndex        =   23
            Top             =   390
            Width           =   1455
         End
         Begin VB.Label Label24 
            Caption         =   "Provincia / Estado:"
            Height          =   315
            Left            =   6180
            TabIndex        =   22
            Top             =   420
            Width           =   1905
         End
         Begin VB.Label Label28 
            Caption         =   "Código Postal:"
            Height          =   285
            Left            =   60
            TabIndex        =   21
            Top             =   720
            Width           =   1485
         End
      End
   End
End
Attribute VB_Name = "frm_MntEmp_53"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

