VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   7200
   Begin Threed.SSPanel SSPanel1 
      Height          =   7815
      Left            =   690
      TabIndex        =   0
      Top             =   2190
      Width           =   14535
      _Version        =   65536
      _ExtentX        =   25638
      _ExtentY        =   13785
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   1545
         Left            =   30
         TabIndex        =   1
         Top             =   5400
         Width           =   14445
         _Version        =   65536
         _ExtentX        =   25479
         _ExtentY        =   2725
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisFia 
            Height          =   1125
            Left            =   30
            TabIndex        =   2
            Top             =   360
            Width           =   14385
            _ExtentX        =   25374
            _ExtentY        =   1984
            _Version        =   393216
            Rows            =   21
            Cols            =   6
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel9 
            Height          =   285
            Left            =   60
            TabIndex        =   3
            Top             =   60
            Width           =   5025
            _Version        =   65536
            _ExtentX        =   8864
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Entidad Financiera"
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
         Begin Threed.SSPanel SSPanel15 
            Height          =   285
            Left            =   7620
            TabIndex        =   4
            Top             =   60
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Emisión"
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
         Begin Threed.SSPanel SSPanel17 
            Height          =   285
            Left            =   10260
            TabIndex        =   5
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
         Begin Threed.SSPanel SSPanel19 
            Height          =   285
            Left            =   8940
            TabIndex        =   6
            Top             =   60
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Vencimiento"
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
         Begin Threed.SSPanel SSPanel18 
            Height          =   285
            Left            =   11850
            TabIndex        =   7
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   5070
            TabIndex        =   8
            Top             =   60
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Carta"
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
      Begin Threed.SSPanel SSPanel4 
         Height          =   1515
         Left            =   30
         TabIndex        =   9
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
            TabIndex        =   10
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
            TabIndex        =   11
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
            TabIndex        =   12
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
            TabIndex        =   13
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
            TabIndex        =   14
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
            TabIndex        =   15
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
            TabIndex        =   16
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
         TabIndex        =   17
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
            TabIndex        =   18
            Top             =   60
            Width           =   4965
            _Version        =   65536
            _ExtentX        =   8758
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Registro de Garantías"
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
            Picture         =   "OpeTra_frm_015.frx":0000
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel20 
         Height          =   765
         Left            =   30
         TabIndex        =   19
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
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   7530
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   60
            Width           =   2775
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   7530
            MaxLength       =   12
            TabIndex        =   24
            Text            =   "Text1"
            Top             =   390
            Width           =   2775
         End
         Begin VB.ComboBox cmb_TipBus 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   60
            Width           =   2775
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   13740
            Picture         =   "OpeTra_frm_015.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   13050
            Picture         =   "OpeTra_frm_015.frx":074C
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   12360
            Picture         =   "OpeTra_frm_015.frx":0A56
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Buscar Registros"
            Top             =   30
            Width           =   675
         End
         Begin MSMask.MaskEdBox msk_NumOpe 
            Height          =   315
            Left            =   1620
            TabIndex        =   26
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
         Begin VB.Label lbl_Numero 
            Caption         =   "Nro. Operación:"
            Height          =   285
            Left            =   90
            TabIndex        =   31
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Doc. Ident.:"
            Height          =   315
            Left            =   6150
            TabIndex        =   30
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label Label19 
            Caption         =   "Nro. Doc. Ident.:"
            Height          =   285
            Left            =   6150
            TabIndex        =   29
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label18 
            Caption         =   "Tipo de Búsqueda:"
            Height          =   315
            Left            =   90
            TabIndex        =   28
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label17 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   60
            TabIndex        =   27
            Top             =   1740
            Width           =   1065
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1425
         Left            =   30
         TabIndex        =   32
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
            TabIndex        =   33
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
            TabIndex        =   34
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
            TabIndex        =   35
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
            TabIndex        =   36
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
            TabIndex        =   37
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
            TabIndex        =   38
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
            TabIndex        =   39
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
            TabIndex        =   40
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
         Begin VB.Label Label2 
            Caption         =   "Moneda:"
            Height          =   315
            Left            =   9060
            TabIndex        =   47
            Top             =   720
            Width           =   945
         End
         Begin VB.Label Label3 
            Caption         =   "Nro. de Operación:"
            Height          =   315
            Left            =   60
            TabIndex        =   46
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label Label21 
            Caption         =   "Modalidad:"
            Height          =   315
            Left            =   9060
            TabIndex        =   45
            Top             =   390
            Width           =   945
         End
         Begin VB.Label Label20 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   9060
            TabIndex        =   44
            Top             =   60
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Monto Préstamo:"
            Height          =   315
            Left            =   9060
            TabIndex        =   43
            Top             =   1080
            Width           =   1245
         End
         Begin VB.Label Label6 
            Caption         =   "Dirección Inmueble:"
            Height          =   405
            Left            =   60
            TabIndex        =   42
            Top             =   720
            Width           =   1395
         End
         Begin VB.Label Label4 
            Caption         =   "Cliente Titular:"
            Height          =   315
            Left            =   60
            TabIndex        =   41
            Top             =   390
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel13 
         Height          =   765
         Left            =   30
         TabIndex        =   48
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
         Begin VB.CommandButton cmd_LevHip 
            Height          =   675
            Left            =   13050
            Picture         =   "OpeTra_frm_015.frx":0D60
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_NueHip 
            Height          =   675
            Left            =   12360
            Picture         =   "OpeTra_frm_015.frx":162A
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_DatInm 
            Height          =   675
            Left            =   13740
            Picture         =   "OpeTra_frm_015.frx":1EF4
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel21 
         Height          =   765
         Left            =   30
         TabIndex        =   52
         Top             =   6990
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
         Begin VB.CommandButton cmd_LibFia 
            Height          =   675
            Left            =   13740
            Picture         =   "OpeTra_frm_015.frx":27BE
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_NueFia 
            Height          =   675
            Left            =   12360
            Picture         =   "OpeTra_frm_015.frx":2AC8
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_RenFia 
            Height          =   675
            Left            =   13050
            Picture         =   "OpeTra_frm_015.frx":2DD2
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   675
         End
      End
   End
   Begin Threed.SSPanel SSPanel22 
      Height          =   2865
      Left            =   0
      TabIndex        =   56
      Top             =   0
      Width           =   11535
      _Version        =   65536
      _ExtentX        =   20346
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
      BevelOuter      =   1
      Begin TabDlg.SSTab tab_Cronog 
         Height          =   2775
         Left            =   60
         TabIndex        =   57
         Top             =   60
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   4895
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Cliente - Tramo No Concesional"
         TabPicture(0)   =   "OpeTra_frm_015.frx":30DC
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "pnl_CliNCo_TotCuo"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "pnl_CliNCo_OtrCar"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "SSPanel62"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "SSPanel61"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "SSPanel59"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "SSPanel36"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "SSPanel35"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "SSPanel34"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "SSPanel33"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "SSPanel24"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "grd_CliNCo_Listad"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "pnl_CliNCo_Intere"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "pnl_CliNCo_SegPre"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "pnl_CliNCo_SegViv"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "pnl_CliNCo_Capita"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "SSPanel30"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).ControlCount=   16
         TabCaption(1)   =   "Cliente - Tramo Concesional"
         TabPicture(1)   =   "OpeTra_frm_015.frx":30F8
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "Mivivienda - Tramo No Concesional"
         TabPicture(2)   =   "OpeTra_frm_015.frx":3114
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         TabCaption(3)   =   "MiVivienda - Tramo Concesional"
         TabPicture(3)   =   "OpeTra_frm_015.frx":3130
         Tab(3).ControlEnabled=   0   'False
         Tab(3).ControlCount=   0
         Begin Threed.SSPanel SSPanel30 
            Height          =   285
            Left            =   3450
            TabIndex        =   58
            Top             =   360
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Interés"
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
         Begin Threed.SSPanel pnl_CliNCo_Capita 
            Height          =   285
            Left            =   2550
            TabIndex        =   59
            Top             =   2370
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2487
            _ExtentY        =   503
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
         Begin Threed.SSPanel pnl_CliNCo_SegViv 
            Height          =   285
            Left            =   6690
            TabIndex        =   60
            Top             =   2370
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2487
            _ExtentY        =   503
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
         Begin Threed.SSPanel pnl_CliNCo_SegPre 
            Height          =   285
            Left            =   5310
            TabIndex        =   61
            Top             =   2370
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2487
            _ExtentY        =   503
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
         Begin Threed.SSPanel pnl_CliNCo_Intere 
            Height          =   285
            Left            =   3930
            TabIndex        =   62
            Top             =   2370
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2487
            _ExtentY        =   503
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
         Begin MSFlexGridLib.MSFlexGrid grd_CliNCo_Listad 
            Height          =   1575
            Left            =   30
            TabIndex        =   63
            Top             =   660
            Width           =   11265
            _ExtentX        =   19870
            _ExtentY        =   2778
            _Version        =   393216
            Rows            =   21
            Cols            =   9
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel23 
            Height          =   285
            Left            =   -67530
            TabIndex        =   64
            Top             =   360
            Width           =   2370
            _Version        =   65536
            _ExtentX        =   4180
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Total Cuota"
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
         Begin Threed.SSPanel SSPanel25 
            Height          =   285
            Left            =   -65190
            TabIndex        =   65
            Top             =   360
            Width           =   2370
            _Version        =   65536
            _ExtentX        =   4180
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Saldo Capital"
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
         Begin Threed.SSPanel pnl_CliCon_TotCuo 
            Height          =   285
            Left            =   -67530
            TabIndex        =   66
            Top             =   1470
            Width           =   2370
            _Version        =   65536
            _ExtentX        =   4180
            _ExtentY        =   503
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
         Begin Threed.SSPanel pnl_CliCon_Intere 
            Height          =   285
            Left            =   -69870
            TabIndex        =   67
            Top             =   1470
            Width           =   2370
            _Version        =   65536
            _ExtentX        =   4180
            _ExtentY        =   503
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
         Begin Threed.SSPanel pnl_CliCon_Capita 
            Height          =   285
            Left            =   -72210
            TabIndex        =   68
            Top             =   1470
            Width           =   2370
            _Version        =   65536
            _ExtentX        =   4180
            _ExtentY        =   503
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
         Begin Threed.SSPanel SSPanel26 
            Height          =   285
            Left            =   -74940
            TabIndex        =   69
            Top             =   360
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cuota"
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
         Begin Threed.SSPanel SSPanel27 
            Height          =   285
            Left            =   -73770
            TabIndex        =   70
            Top             =   360
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Vcto"
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
         Begin Threed.SSPanel SSPanel28 
            Height          =   285
            Left            =   -71970
            TabIndex        =   71
            Top             =   360
            Width           =   1860
            _Version        =   65536
            _ExtentX        =   3281
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Capital"
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
         Begin Threed.SSPanel SSPanel29 
            Height          =   285
            Left            =   -70140
            TabIndex        =   72
            Top             =   360
            Width           =   1860
            _Version        =   65536
            _ExtentX        =   3281
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Interés"
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
         Begin Threed.SSPanel SSPanel31 
            Height          =   285
            Left            =   -66480
            TabIndex        =   73
            Top             =   360
            Width           =   1860
            _Version        =   65536
            _ExtentX        =   3281
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Total Cuota"
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
         Begin Threed.SSPanel SSPanel32 
            Height          =   285
            Left            =   -64650
            TabIndex        =   74
            Top             =   360
            Width           =   1860
            _Version        =   65536
            _ExtentX        =   3281
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Saldo Capital"
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
         Begin Threed.SSPanel pnl_CofNCo_Comisi 
            Height          =   285
            Left            =   -68310
            TabIndex        =   75
            Top             =   1470
            Width           =   1860
            _Version        =   65536
            _ExtentX        =   3281
            _ExtentY        =   503
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
         Begin Threed.SSPanel pnl_CofNCo_Intere 
            Height          =   285
            Left            =   -70140
            TabIndex        =   76
            Top             =   1470
            Width           =   1860
            _Version        =   65536
            _ExtentX        =   3281
            _ExtentY        =   503
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
         Begin Threed.SSPanel pnl_CofNCo_Capita 
            Height          =   285
            Left            =   -71970
            TabIndex        =   77
            Top             =   1470
            Width           =   1860
            _Version        =   65536
            _ExtentX        =   3281
            _ExtentY        =   503
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
         Begin Threed.SSPanel SSPanel37 
            Height          =   285
            Left            =   -68310
            TabIndex        =   78
            Top             =   360
            Width           =   1860
            _Version        =   65536
            _ExtentX        =   3281
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Comisión"
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
         Begin Threed.SSPanel pnl_CofNCo_TotCuo 
            Height          =   285
            Left            =   -66480
            TabIndex        =   79
            Top             =   1470
            Width           =   1860
            _Version        =   65536
            _ExtentX        =   3281
            _ExtentY        =   503
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
         Begin Threed.SSPanel SSPanel40 
            Height          =   285
            Left            =   -74940
            TabIndex        =   80
            Top             =   360
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cuota"
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
         Begin Threed.SSPanel SSPanel48 
            Height          =   285
            Left            =   -73770
            TabIndex        =   81
            Top             =   360
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Vcto"
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
         Begin Threed.SSPanel SSPanel50 
            Height          =   285
            Left            =   -71970
            TabIndex        =   82
            Top             =   360
            Width           =   1860
            _Version        =   65536
            _ExtentX        =   3281
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Capital"
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
         Begin Threed.SSPanel SSPanel51 
            Height          =   285
            Left            =   -70140
            TabIndex        =   83
            Top             =   360
            Width           =   1860
            _Version        =   65536
            _ExtentX        =   3281
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Interés"
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
         Begin Threed.SSPanel SSPanel52 
            Height          =   285
            Left            =   -66480
            TabIndex        =   84
            Top             =   360
            Width           =   1860
            _Version        =   65536
            _ExtentX        =   3281
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Total Cuota"
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
         Begin Threed.SSPanel SSPanel53 
            Height          =   285
            Left            =   -64650
            TabIndex        =   85
            Top             =   360
            Width           =   1860
            _Version        =   65536
            _ExtentX        =   3281
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Saldo Capital"
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
         Begin Threed.SSPanel pnl_CofCon_Comisi 
            Height          =   285
            Left            =   -68310
            TabIndex        =   86
            Top             =   1470
            Width           =   1860
            _Version        =   65536
            _ExtentX        =   3281
            _ExtentY        =   503
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
         Begin Threed.SSPanel pnl_CofCon_Intere 
            Height          =   285
            Left            =   -70140
            TabIndex        =   87
            Top             =   1470
            Width           =   1860
            _Version        =   65536
            _ExtentX        =   3281
            _ExtentY        =   503
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
         Begin Threed.SSPanel pnl_CofCon_Capita 
            Height          =   285
            Left            =   -71970
            TabIndex        =   88
            Top             =   1470
            Width           =   1860
            _Version        =   65536
            _ExtentX        =   3281
            _ExtentY        =   503
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
         Begin Threed.SSPanel SSPanel57 
            Height          =   285
            Left            =   -68310
            TabIndex        =   89
            Top             =   360
            Width           =   1860
            _Version        =   65536
            _ExtentX        =   3281
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Comisión"
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
         Begin Threed.SSPanel pnl_CofCon_TotCuo 
            Height          =   285
            Left            =   -66480
            TabIndex        =   90
            Top             =   1470
            Width           =   1860
            _Version        =   65536
            _ExtentX        =   3281
            _ExtentY        =   503
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
         Begin Threed.SSPanel SSPanel24 
            Height          =   285
            Left            =   60
            TabIndex        =   91
            Top             =   360
            Width           =   795
            _Version        =   65536
            _ExtentX        =   1402
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cuota"
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
         Begin Threed.SSPanel SSPanel33 
            Height          =   285
            Left            =   840
            TabIndex        =   92
            Top             =   360
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Vcto"
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
         Begin Threed.SSPanel SSPanel34 
            Height          =   285
            Left            =   2280
            TabIndex        =   93
            Top             =   360
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Capital"
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
         Begin Threed.SSPanel SSPanel35 
            Height          =   285
            Left            =   8130
            TabIndex        =   94
            Top             =   360
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Total Cuota"
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
         Begin Threed.SSPanel SSPanel36 
            Height          =   285
            Left            =   9420
            TabIndex        =   95
            Top             =   360
            Width           =   1560
            _Version        =   65536
            _ExtentX        =   2752
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Saldo Capital"
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
         Begin Threed.SSPanel SSPanel59 
            Height          =   285
            Left            =   4620
            TabIndex        =   96
            Top             =   360
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Seg. Prest."
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
         Begin Threed.SSPanel SSPanel61 
            Height          =   285
            Left            =   5790
            TabIndex        =   97
            Top             =   360
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Seg. Vivienda"
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
         Begin Threed.SSPanel SSPanel62 
            Height          =   285
            Left            =   6960
            TabIndex        =   98
            Top             =   360
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Portes"
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
         Begin Threed.SSPanel pnl_CliNCo_OtrCar 
            Height          =   285
            Left            =   8070
            TabIndex        =   99
            Top             =   2370
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2487
            _ExtentY        =   503
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
         Begin Threed.SSPanel pnl_CliNCo_TotCuo 
            Height          =   285
            Left            =   9450
            TabIndex        =   100
            Top             =   2370
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2487
            _ExtentY        =   503
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
         Begin VB.Label Label15 
            Caption         =   "Totales ==>"
            Height          =   285
            Left            =   -73230
            TabIndex        =   103
            Top             =   1470
            Width           =   945
         End
         Begin VB.Label Label14 
            Caption         =   "Totales ==>"
            Height          =   285
            Left            =   -72930
            TabIndex        =   102
            Top             =   1470
            Width           =   945
         End
         Begin VB.Label Label13 
            Caption         =   "Totales ==>"
            Height          =   285
            Left            =   -72930
            TabIndex        =   101
            Top             =   1470
            Width           =   945
         End
      End
   End
End
Attribute VB_Name = "Form1"
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
   
      cmd_NueFia.Enabled = False
      cmd_RenFia.Enabled = False
      cmd_LibFia.Enabled = False
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
   
   If grd_LisHip.Rows = 0 Then
      Exit Sub
   End If
   
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
   
   grd_LisFia.ColWidth(0) = 4995
   grd_LisFia.ColWidth(1) = 2550
   grd_LisFia.ColWidth(2) = 1305
   grd_LisFia.ColWidth(3) = 1305
   grd_LisFia.ColWidth(4) = 1575
   grd_LisFia.ColWidth(5) = 2205
   
   grd_LisFia.ColAlignment(0) = flexAlignLeftCenter
   grd_LisFia.ColAlignment(1) = flexAlignLeftCenter
   grd_LisFia.ColAlignment(2) = flexAlignCenterCenter
   grd_LisFia.ColAlignment(3) = flexAlignCenterCenter
   grd_LisFia.ColAlignment(4) = flexAlignRightCenter
   grd_LisFia.ColAlignment(5) = flexAlignLeftCenter
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
   Call gs_LimpiaGrid(grd_LisFia)
End Sub

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   cmb_TipBus.Enabled = p_Habilita
   cmb_TipDoc.Enabled = p_Habilita
   txt_NumDoc.Enabled = p_Habilita
   msk_NumOpe.Enabled = p_Habilita
   cmd_Buscar.Enabled = p_Habilita
   
   grd_LisHip.Enabled = Not p_Habilita
   grd_LisFia.Enabled = Not p_Habilita
   cmd_NueHip.Enabled = Not p_Habilita
   cmd_LevHip.Enabled = Not p_Habilita
   cmd_DatInm.Enabled = Not p_Habilita
   cmd_NueFia.Enabled = Not p_Habilita
   cmd_RenFia.Enabled = Not p_Habilita
   cmd_LibFia.Enabled = Not p_Habilita
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


