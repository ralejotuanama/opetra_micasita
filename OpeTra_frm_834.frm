VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_Ges_TecPro_12 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13155
   Icon            =   "OpeTra_frm_834.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10260
   ScaleWidth      =   13155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   10305
      Left            =   30
      TabIndex        =   17
      Top             =   0
      Width           =   13155
      _Version        =   65536
      _ExtentX        =   23204
      _ExtentY        =   18177
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
         Height          =   10275
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   13125
         _Version        =   65536
         _ExtentX        =   23151
         _ExtentY        =   18124
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
         Begin Threed.SSPanel SSPanel14 
            Height          =   1785
            Left            =   30
            TabIndex        =   60
            Top             =   6120
            Width           =   13035
            _Version        =   65536
            _ExtentX        =   22992
            _ExtentY        =   3149
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
            BevelOuter      =   1
            Begin VB.ComboBox cmb_TipRec 
               Height          =   315
               Left            =   8820
               Style           =   2  'Dropdown List
               TabIndex        =   79
               Top             =   90
               Width           =   4095
            End
            Begin VB.TextBox txt_ApeCas 
               Height          =   315
               Left            =   8820
               MaxLength       =   30
               TabIndex        =   68
               Text            =   "Text1"
               Top             =   750
               Width           =   4095
            End
            Begin VB.TextBox txt_ApeMat 
               Height          =   315
               Left            =   1950
               MaxLength       =   30
               TabIndex        =   67
               Text            =   "Text1"
               Top             =   750
               Width           =   4095
            End
            Begin VB.TextBox txt_ApePat 
               Height          =   315
               Left            =   1950
               MaxLength       =   30
               TabIndex        =   66
               Text            =   "Text1"
               Top             =   420
               Width           =   4095
            End
            Begin VB.TextBox txt_Nombre 
               Height          =   315
               Left            =   1950
               MaxLength       =   30
               TabIndex        =   65
               Text            =   "Text1"
               Top             =   1080
               Width           =   4095
            End
            Begin VB.TextBox txt_TelFij 
               Height          =   315
               Left            =   1950
               MaxLength       =   25
               TabIndex        =   64
               Text            =   "Text1"
               Top             =   1410
               Width           =   1545
            End
            Begin VB.TextBox txt_DirEle 
               Height          =   315
               Left            =   8820
               MaxLength       =   120
               TabIndex        =   63
               Text            =   "Text1"
               Top             =   1410
               Width           =   4095
            End
            Begin VB.TextBox txt_TelCel 
               Height          =   315
               Left            =   4440
               MaxLength       =   25
               TabIndex        =   62
               Text            =   "Text1"
               Top             =   1410
               Width           =   1605
            End
            Begin VB.TextBox txt_CodPry 
               Height          =   315
               Left            =   1950
               MaxLength       =   30
               TabIndex        =   61
               Text            =   "Text1"
               Top             =   90
               Width           =   4095
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Recurso:"
               Height          =   195
               Left            =   6840
               TabIndex        =   80
               Top             =   195
               Width           =   645
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               Caption         =   "Apellido Casada:"
               Height          =   195
               Left            =   6840
               TabIndex        =   76
               Top             =   810
               Width           =   1185
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Apellido Materno:"
               Height          =   195
               Left            =   90
               TabIndex        =   75
               Top             =   810
               Width           =   1230
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Apellido Paterno:"
               Height          =   195
               Left            =   90
               TabIndex        =   74
               Top             =   480
               Width           =   1200
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Nombres:"
               Height          =   195
               Left            =   90
               TabIndex        =   73
               Top             =   1140
               Width           =   675
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Teléfono  Fijo:"
               Height          =   195
               Left            =   90
               TabIndex        =   72
               Top             =   1470
               Width           =   1005
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "E-mail:"
               Height          =   195
               Left            =   6840
               TabIndex        =   71
               Top             =   1470
               Width           =   465
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Celular:"
               Height          =   195
               Left            =   3690
               TabIndex        =   70
               Top             =   1470
               Width           =   525
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Código Proyecto:"
               Height          =   195
               Left            =   90
               TabIndex        =   69
               Top             =   150
               Width           =   1215
            End
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   675
            Left            =   30
            TabIndex        =   19
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
            Begin VB.CommandButton cmd_Import 
               Height          =   585
               Left            =   2310
               Picture         =   "OpeTra_frm_834.frx":000C
               Style           =   1  'Graphical
               TabIndex        =   77
               Top             =   30
               Width           =   585
            End
            Begin VB.CommandButton cmd_Salida 
               Height          =   585
               Left            =   12420
               Picture         =   "OpeTra_frm_834.frx":044E
               Style           =   1  'Graphical
               TabIndex        =   55
               ToolTipText     =   "Salir"
               Top             =   30
               Width           =   585
            End
            Begin VB.CommandButton cmd_Cancel 
               Height          =   585
               Left            =   11850
               Picture         =   "OpeTra_frm_834.frx":0890
               Style           =   1  'Graphical
               TabIndex        =   54
               ToolTipText     =   "Cancelar"
               Top             =   30
               Width           =   585
            End
            Begin VB.CommandButton cmd_ExpExc 
               Height          =   585
               Left            =   1740
               Picture         =   "OpeTra_frm_834.frx":0B9A
               Style           =   1  'Graphical
               TabIndex        =   15
               ToolTipText     =   "Exportar a Excel"
               Top             =   30
               Width           =   585
            End
            Begin VB.CommandButton cmd_Agrega 
               Height          =   585
               Left            =   30
               Picture         =   "OpeTra_frm_834.frx":0EA4
               Style           =   1  'Graphical
               TabIndex        =   12
               ToolTipText     =   "Nuevo Registro"
               Top             =   30
               Width           =   585
            End
            Begin VB.CommandButton cmd_Editar 
               Height          =   585
               Left            =   600
               Picture         =   "OpeTra_frm_834.frx":11AE
               Style           =   1  'Graphical
               TabIndex        =   13
               ToolTipText     =   "Modificar Registro"
               Top             =   30
               Width           =   585
            End
            Begin VB.CommandButton cmd_Borrar 
               Height          =   585
               Left            =   1170
               Picture         =   "OpeTra_frm_834.frx":14B8
               Style           =   1  'Graphical
               TabIndex        =   14
               ToolTipText     =   "Borrar Registro"
               Top             =   30
               Width           =   585
            End
            Begin VB.CommandButton cmd_Grabar 
               Height          =   585
               Left            =   11280
               Picture         =   "OpeTra_frm_834.frx":17C2
               Style           =   1  'Graphical
               TabIndex        =   16
               ToolTipText     =   "Grabar Datos"
               Top             =   30
               Width           =   585
            End
         End
         Begin Threed.SSPanel SSPanel6 
            Height          =   675
            Left            =   30
            TabIndex        =   20
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
               TabIndex        =   21
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
               Left            =   630
               TabIndex        =   22
               Top             =   330
               Width           =   4215
               _Version        =   65536
               _ExtentX        =   7435
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Carta Fianza - Beneficiarios"
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
            Begin MSComDlg.CommonDialog dlg_Guarda 
               Left            =   12390
               Top             =   90
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.Image Image1 
               Height          =   480
               Left            =   60
               Picture         =   "OpeTra_frm_834.frx":1C04
               Top             =   60
               Width           =   480
            End
         End
         Begin Threed.SSPanel SSPanel9 
            Height          =   1155
            Left            =   30
            TabIndex        =   23
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
               TabIndex        =   24
               Top             =   450
               Width           =   5625
               _Version        =   65536
               _ExtentX        =   9922
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
            Begin Threed.SSPanel pnl_TipDoc 
               Height          =   315
               Left            =   1620
               TabIndex        =   25
               Top             =   120
               Width           =   5625
               _Version        =   65536
               _ExtentX        =   9922
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
            Begin Threed.SSPanel pnl_NroDoc 
               Height          =   315
               Left            =   9360
               TabIndex        =   26
               Top             =   120
               Width           =   2895
               _Version        =   65536
               _ExtentX        =   5106
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
            Begin Threed.SSPanel pnl_TipEmp 
               Height          =   315
               Left            =   1620
               TabIndex        =   27
               Top             =   780
               Width           =   2895
               _Version        =   65536
               _ExtentX        =   5106
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
            Begin Threed.SSPanel pnl_NumRef 
               Height          =   315
               Left            =   9360
               TabIndex        =   28
               Top             =   450
               Width           =   2895
               _Version        =   65536
               _ExtentX        =   5106
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
            Begin VB.Label Label3 
               Caption         =   "Nro. Referencia:"
               Height          =   255
               Left            =   7770
               TabIndex        =   33
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label lbl_TipDoc 
               Caption         =   "Tipo Documento:"
               Height          =   255
               Left            =   120
               TabIndex        =   32
               Top             =   150
               Width           =   1335
            End
            Begin VB.Label lbl_NumDoc 
               Caption         =   "Nro. Documento:"
               Height          =   225
               Left            =   7770
               TabIndex        =   31
               Top             =   150
               Width           =   1335
            End
            Begin VB.Label lbl_RazSoc 
               Caption         =   "Razón Social:"
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label lbl_TipEmp 
               Caption         =   "Tipo Empresa:"
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   810
               Width           =   1335
            End
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   2805
            Left            =   30
            TabIndex        =   34
            Top             =   2700
            Width           =   13035
            _Version        =   65536
            _ExtentX        =   22992
            _ExtentY        =   4948
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
               Height          =   2265
               Left            =   90
               TabIndex        =   35
               Top             =   450
               Width           =   12930
               _ExtentX        =   22807
               _ExtentY        =   3995
               _Version        =   393216
               Rows            =   21
               Cols            =   8
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
               Appearance      =   0
            End
            Begin Threed.SSPanel pnl_Tit_CodIte 
               Height          =   285
               Left            =   90
               TabIndex        =   36
               Top             =   150
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Items"
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
            Begin Threed.SSPanel pnl_Tit_NumDoc 
               Height          =   285
               Left            =   1050
               TabIndex        =   37
               Top             =   150
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Documento"
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
            Begin Threed.SSPanel pnl_Tit_ApePat 
               Height          =   285
               Left            =   2610
               TabIndex        =   38
               Top             =   150
               Width           =   2445
               _Version        =   65536
               _ExtentX        =   4313
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Apellido Paterno"
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
            Begin Threed.SSPanel pnl_Tit_ApeMat 
               Height          =   285
               Left            =   5040
               TabIndex        =   39
               Top             =   150
               Width           =   2445
               _Version        =   65536
               _ExtentX        =   4313
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Apellido Materno"
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
            Begin Threed.SSPanel pnl_Tit_Nombre 
               Height          =   285
               Left            =   7470
               TabIndex        =   40
               Top             =   150
               Width           =   2445
               _Version        =   65536
               _ExtentX        =   4313
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Nombres"
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
            Begin Threed.SSPanel pnl_Tit_Importe 
               Height          =   285
               Left            =   11100
               TabIndex        =   56
               Top             =   150
               Width           =   1665
               _Version        =   65536
               _ExtentX        =   2937
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Importe BFH / Ahorro"
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
            Begin Threed.SSPanel pnl_Tit_Recurso 
               Height          =   285
               Left            =   9840
               TabIndex        =   78
               Top             =   150
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Recurso"
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   525
            Left            =   30
            TabIndex        =   41
            Top             =   5550
            Width           =   13035
            _Version        =   65536
            _ExtentX        =   22992
            _ExtentY        =   926
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
               Left            =   1950
               Style           =   2  'Dropdown List
               TabIndex        =   0
               Top             =   120
               Width           =   4095
            End
            Begin VB.TextBox txt_NumDoc 
               Height          =   315
               Left            =   8820
               MaxLength       =   12
               TabIndex        =   1
               Text            =   "Text1"
               Top             =   120
               Width           =   4095
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Docum. Identidad:"
               Height          =   195
               Left            =   90
               TabIndex        =   43
               Top             =   180
               Width           =   1665
            End
            Begin VB.Label Label33 
               AutoSize        =   -1  'True
               Caption         =   "Nro. Doc. Id.:"
               Height          =   195
               Left            =   6840
               TabIndex        =   42
               Top             =   180
               Width           =   960
            End
         End
         Begin Threed.SSPanel SSPanel12 
            Height          =   1725
            Left            =   30
            TabIndex        =   44
            Top             =   7950
            Width           =   13035
            _Version        =   65536
            _ExtentX        =   22992
            _ExtentY        =   3043
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
            Begin VB.TextBox txt_Refere 
               Height          =   315
               Left            =   8820
               MaxLength       =   250
               TabIndex        =   11
               Text            =   "Text1"
               Top             =   1380
               Width           =   4095
            End
            Begin VB.ComboBox cmb_DstDir 
               Height          =   315
               Left            =   1950
               TabIndex        =   10
               Text            =   "cmb_DstDir"
               Top             =   1380
               Width           =   4125
            End
            Begin VB.ComboBox cmb_PrvDir 
               Height          =   315
               Left            =   8820
               TabIndex        =   9
               Text            =   "cmb_PrvDir"
               Top             =   1050
               Width           =   4095
            End
            Begin VB.ComboBox cmb_DptDir 
               Height          =   315
               Left            =   1950
               TabIndex        =   8
               Text            =   "cmb_DptDir"
               Top             =   1050
               Width           =   4125
            End
            Begin VB.TextBox txt_NomZon 
               Height          =   315
               Left            =   8820
               MaxLength       =   120
               TabIndex        =   7
               Text            =   "Text1"
               Top             =   720
               Width           =   4095
            End
            Begin VB.ComboBox cmb_TipZon 
               Height          =   315
               Left            =   1950
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   720
               Width           =   4125
            End
            Begin VB.TextBox txt_IntDpt 
               Height          =   315
               Left            =   10860
               MaxLength       =   30
               TabIndex        =   5
               Text            =   "Text1"
               Top             =   390
               Width           =   2025
            End
            Begin VB.TextBox txt_NumVia 
               Height          =   315
               Left            =   8820
               MaxLength       =   30
               TabIndex        =   4
               Text            =   "Text1"
               Top             =   390
               Width           =   2025
            End
            Begin VB.TextBox txt_NomVia 
               Height          =   315
               Left            =   1950
               MaxLength       =   120
               TabIndex        =   3
               Text            =   "Text1"
               Top             =   390
               Width           =   4125
            End
            Begin VB.ComboBox cmb_TipVia 
               Height          =   315
               Left            =   1950
               Style           =   2  'Dropdown List
               TabIndex        =   2
               Top             =   60
               Width           =   4125
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               Caption         =   "Referencia:"
               Height          =   195
               Left            =   6840
               TabIndex        =   53
               Top             =   1440
               Width           =   825
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               Caption         =   "Distrito:"
               Height          =   195
               Left            =   90
               TabIndex        =   52
               Top             =   1440
               Width           =   525
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               Caption         =   "Provincia:"
               Height          =   195
               Left            =   6840
               TabIndex        =   51
               Top             =   1110
               Width           =   705
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "Departamento:"
               Height          =   195
               Left            =   90
               TabIndex        =   50
               Top             =   1110
               Width           =   1050
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               Caption         =   "Nombre Zona:"
               Height          =   195
               Left            =   6840
               TabIndex        =   49
               Top             =   780
               Width           =   1020
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Zona:"
               Height          =   195
               Left            =   90
               TabIndex        =   48
               Top             =   780
               Width           =   1005
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "Nro - Int/Dpto/Mza/Lote:"
               Height          =   195
               Left            =   6840
               TabIndex        =   47
               Top             =   450
               Width           =   1800
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Nombre Vía:"
               Height          =   195
               Left            =   90
               TabIndex        =   46
               Top             =   450
               Width           =   900
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Vía:"
               Height          =   195
               Left            =   90
               TabIndex        =   45
               Top             =   120
               Width           =   885
            End
         End
         Begin Threed.SSPanel SSPanel10 
            Height          =   495
            Left            =   30
            TabIndex        =   57
            Top             =   9720
            Width           =   13035
            _Version        =   65536
            _ExtentX        =   22992
            _ExtentY        =   873
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
            Begin EditLib.fpDoubleSingle ipp_ImpBFH 
               Height          =   315
               Left            =   1950
               TabIndex        =   59
               Top             =   90
               Width           =   2085
               _Version        =   196608
               _ExtentX        =   3678
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
               ButtonStyle     =   0
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
               Text            =   "0.00"
               DecimalPlaces   =   2
               DecimalPoint    =   "."
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "9000000000"
               MinValue        =   "0"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ","
               UseSeparator    =   -1  'True
               IncInt          =   1
               IncDec          =   1
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
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Importe BFH / Ahorro:"
               Height          =   195
               Left            =   90
               TabIndex        =   58
               Top             =   150
               Width           =   1560
            End
         End
      End
   End
End
Attribute VB_Name = "frm_Ges_TecPro_12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_str_DptDir     As String
Dim l_str_PrvDir     As String
Dim l_str_DstDir     As String
Dim l_int_FlgCmb     As Integer
Dim l_int_Codigo     As Integer

Private Type r_Arr_CargaDatBen
   Codigo            As String
   Tipdoc            As String
   NumDoc            As String
   ApePat            As String
   ApeMat            As String
   Nombre            As String
   CodPry            As String
   ImpBHF            As String
   TipRec            As Integer
End Type

Dim r_str_DatBen()   As r_Arr_CargaDatBen
Dim l_arr_TipRec()   As moddat_tpo_Genera

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

Private Sub cmb_TipRec_Click()
   Call gs_SetFocus(txt_ApePat)
End Sub

Private Sub cmb_TipRec_KeyPress(KeyAscii As Integer)
   Call cmb_TipRec_Click
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

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgAct_2 = 1
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   Call fs_Activa(True)
   If Me.grd_Listad.Row <= 0 Then
      cmd_Editar.Enabled = False
      cmd_Borrar.Enabled = False
      cmd_ExpExc.Enabled = False
   Else
      cmd_Editar.Enabled = True
      cmd_Borrar.Enabled = True
   End If
   Call gs_SetFocus(cmb_TipDoc)
End Sub

Private Sub cmd_Borrar_Click()
   
   If MsgBox("¿Está seguro de eliminar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
   
      'Grabando Información de Carta Fianza
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " USP_TPR_DATBEN_ELIMINA ("
      g_str_Parame = g_str_Parame & "'" & CStr(moddat_g_str_DesIte) & "', " 'moddat_g_str_NumFia
      g_str_Parame = g_str_Parame & CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0)) & ") "
                  
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la eliminación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      Screen.MousePointer = 0
   Loop
     
   'Actualiza la Grilla
   Call fs_Buscar
   Call fs_Activa(False)
   
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_Limpia
   Call fs_Buscar
   Call fs_Activa(False)
End Sub

Private Sub cmd_Editar_Click()
   moddat_g_int_FlgAct_2 = 2
   Call fs_Activa(True)
   Call fs_Cargar_Datos(grd_Listad.TextMatrix(grd_Listad.Row, 7))
   
End Sub
Private Sub fs_Cargar_Datos(ByRef p_Codigo As Integer)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT DATBEN_CODIGO, DATBEN_TIPDOC, DATBEN_NUMDOC, DATBEN_APEPAT, DATBEN_APEMAT, DATBEN_APECAS, "
   g_str_Parame = g_str_Parame & "        DATBEN_NOMBRE, DATBEN_TELFIJ, DATBEN_NUMCEL, DATBEN_DIRELE, DATBEN_TIPVIA, DATBEN_NOMVIA, "
   g_str_Parame = g_str_Parame & "        DATBEN_NUMERO, DATBEN_INTDPT, DATBEN_TIPZON, DATBEN_NOMZON, DATBEN_UBIGEO, DATBEN_REFERE, "
   g_str_Parame = g_str_Parame & "        DATBEN_IMPBFH, DATBEN_CODPRY, DATBEN_TIPREC "
   g_str_Parame = g_str_Parame & "   FROM TPR_DATBEN "
   g_str_Parame = g_str_Parame & "  WHERE DATBEN_NUMREF = '" & moddat_g_str_DesIte & "'" 'moddat_g_str_NumFia
   g_str_Parame = g_str_Parame & "    AND DATBEN_CODIGO = " & CStr(p_Codigo) & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      l_int_Codigo = CStr(g_rst_Princi!DATBEN_CODIGO)
      Call gs_BuscarCombo_Item(cmb_TipDoc, g_rst_Princi!DATBEN_TIPDOC)
      txt_NumDoc.Text = Trim(g_rst_Princi!DATBEN_NUMDOC)
      
      cmb_TipDoc.Enabled = False
      txt_NumDoc.Enabled = False
      
      txt_CodPry.Text = Trim(g_rst_Princi!DATBEN_CODPRY & "") 'DATBEN_CODIGO
      txt_ApePat.Text = Trim(g_rst_Princi!DATBEN_APEPAT & "")
      txt_ApeMat.Text = Trim(g_rst_Princi!DATBEN_APEMAT & "")
      txt_ApeCas.Text = Trim(g_rst_Princi!DATBEN_APECAS & "")
      txt_Nombre.Text = Trim(g_rst_Princi!DATBEN_NOMBRE & "")
      txt_TelFij.Text = Trim(g_rst_Princi!DATBEN_TELFIJ & "")
      txt_TelCel.Text = Trim(g_rst_Princi!DATBEN_NUMCEL & "")
      txt_DirEle.Text = Trim(g_rst_Princi!DATBEN_DIRELE & "")
      
      If Not IsNull((g_rst_Princi!DATBEN_TIPREC)) Then
         cmb_TipRec.ListIndex = gf_Busca_Arregl(l_arr_TipRec, Trim(CStr(g_rst_Princi!DATBEN_TIPREC) & "")) - 1
      End If
   
      If Not IsNull(g_rst_Princi!DATBEN_TIPVIA) Then
         Call gs_BuscarCombo_Item(cmb_TipVia, g_rst_Princi!DATBEN_TIPVIA)
      End If
      txt_NomVia.Text = Trim(g_rst_Princi!DATBEN_NOMVIA & "")
      txt_NumVia.Text = Trim(g_rst_Princi!DATBEN_NUMERO & "")
      txt_IntDpt.Text = Trim(g_rst_Princi!DATBEN_INTDPT & "")
      
      If Not IsNull(g_rst_Princi!DATBEN_TIPZON) Then
         Call gs_BuscarCombo_Item(cmb_TipZon, g_rst_Princi!DATBEN_TIPZON)
      End If

      txt_NomZon.Text = Trim(g_rst_Princi!DATBEN_NOMZON & "")
      
      If Not IsNull(g_rst_Princi!DATBEN_UBIGEO) Then
         Call gs_BuscarCombo_Item(cmb_DptDir, CInt(Left(g_rst_Princi!DATBEN_UBIGEO, 2)))
         Call moddat_gs_Carga_Provin(cmb_PrvDir, Left(g_rst_Princi!DATBEN_UBIGEO, 2))
         Call gs_BuscarCombo_Item(cmb_PrvDir, CInt(Mid(g_rst_Princi!DATBEN_UBIGEO, 3, 2)))
         Call moddat_gs_Carga_Distri(cmb_DstDir, Left(g_rst_Princi!DATBEN_UBIGEO, 2), Mid(g_rst_Princi!DATBEN_UBIGEO, 3, 2))
         Call gs_BuscarCombo_Item(cmb_DstDir, CInt(Right(g_rst_Princi!DATBEN_UBIGEO, 2)))
      End If
      
      txt_Refere.Text = Trim(g_rst_Princi!DATBEN_REFERE & "")
      ipp_ImpBFH.Value = Format(CStr(g_rst_Princi!DATBEN_IMPBFH), "###,###,###,##0.00")

      Call gs_SetFocus(txt_ApePat)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub cmd_ExpExc_Click()
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub
Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_str_ParAux        As String
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
Dim r_int_NroFil        As Integer
Dim r_int_NoFlLi        As Integer
   
    r_int_NroFil = 9
    
    Set r_obj_Excel = New Excel.Application
    r_obj_Excel.SheetsInNewWorkbook = 1
    r_obj_Excel.Workbooks.Add

    With r_obj_Excel.ActiveSheet
        .Cells(2, 2) = "REPORTE DE BENEFICIARIOS" ' POR CARTA FIANZA
        .Range(.Cells(2, 2), .Cells(2, 8)).Merge
        .Range(.Cells(2, 2), .Cells(2, 8)).Font.Bold = True
        .Range(.Cells(2, 2), .Cells(2, 8)).HorizontalAlignment = xlHAlignCenter
        .Range(.Cells(2, 2), .Cells(2, 8)).Font.Size = 14

        .Cells(4, 2) = "TIPO DE DOCUMENTO"
        .Cells(4, 3) = Trim(pnl_TipDoc.Caption)
        .Cells(5, 2) = "NRO. DOCUMENTO"
        .Cells(5, 3) = "'" & Trim(pnl_NroDoc.Caption)
        .Cells(6, 2) = "RAZÓN SOCIAL"
        .Cells(6, 3) = Trim(pnl_RazSoc.Caption)
        .Cells(7, 2) = "NÚMERO" ' CARTA FIANZA
        .Cells(7, 3) = "'" & pnl_NumRef.Caption
        
        .Range(.Cells(3, 2), .Cells(7, 2)).Font.Bold = True
        
        .Cells(r_int_NroFil, 2) = "CODIGO"
        .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 2)).Merge
        .Cells(r_int_NroFil, 3) = "TIPO - NRO. DOCUMENTO"
        .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil + 1, 3)).Merge
        .Cells(r_int_NroFil, 4) = "APELLIDO PATERNO"
        .Range(.Cells(r_int_NroFil, 4), .Cells(r_int_NroFil + 1, 4)).Merge
        .Cells(r_int_NroFil, 5) = "APELLIDO MATERNO"
        .Range(.Cells(r_int_NroFil, 5), .Cells(r_int_NroFil + 1, 5)).Merge
        .Cells(r_int_NroFil, 6) = "NOMBRES"
        .Range(.Cells(r_int_NroFil, 6), .Cells(r_int_NroFil + 1, 6)).Merge
        .Cells(r_int_NroFil, 7) = "TIPO RECURSO"
        .Range(.Cells(r_int_NroFil, 7), .Cells(r_int_NroFil + 1, 7)).Merge
        .Cells(r_int_NroFil, 8) = "IMPORTE"
        .Range(.Cells(r_int_NroFil, 8), .Cells(r_int_NroFil + 1, 8)).Merge
        
        .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 8)).Interior.Color = RGB(146, 208, 80)
        .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 8)).Font.Bold = True
        .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 8)).HorizontalAlignment = xlHAlignCenter
        
        .Columns("A").ColumnWidth = 1
        .Columns("B").ColumnWidth = 22
        .Columns("C").ColumnWidth = 22
        .Columns("D").ColumnWidth = 25
        .Columns("D").HorizontalAlignment = xlHAlignCenter
        .Columns("E").ColumnWidth = 25
        .Columns("E").HorizontalAlignment = xlHAlignCenter
        .Columns("F").ColumnWidth = 25
        .Columns("F").HorizontalAlignment = xlHAlignCenter
        .Columns("G").ColumnWidth = 25
        .Columns("G").HorizontalAlignment = xlHAlignCenter
        .Columns("H").ColumnWidth = 25
        .Columns("H").HorizontalAlignment = xlHAlignRight
        .Columns("H").NumberFormat = "###,###,##0.00"
        
        With .Range(.Cells(8, 2), .Cells(9, 8))
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
            .Cells(r_int_NroFil, 4) = "'" & grd_Listad.TextMatrix(r_int_NoFlLi, 2)
            .Cells(r_int_NroFil, 5) = grd_Listad.TextMatrix(r_int_NoFlLi, 3)
            .Cells(r_int_NroFil, 6) = grd_Listad.TextMatrix(r_int_NoFlLi, 4)
            .Cells(r_int_NroFil, 7) = grd_Listad.TextMatrix(r_int_NoFlLi, 5)
            .Cells(r_int_NroFil, 8) = grd_Listad.TextMatrix(r_int_NoFlLi, 6)
            
            r_int_NroFil = r_int_NroFil + 1
        Next r_int_NoFlLi
        
        With .Range(.Cells(10, 2), .Cells(r_int_NroFil, 3))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
    
        With .Range(.Cells(9, 2), .Cells(r_int_NroFil - 1, 8))
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        End With
   End With
   
   r_obj_Excel.Visible = True
End Sub
Private Sub cmd_Grabar_Click()
   If moddat_g_str_CodPrd <> "026" And moddat_g_str_CodSub <> "001" And moddat_g_str_CodMod <> "004" Then
      If Len(Trim(txt_CodPry.Text)) = 0 Then
         MsgBox "Debe ingresar el Código del Proyecto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_CodPry)
         Exit Sub
      End If
   End If
   If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1 Then
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
   End If
   
   If Len(Trim(txt_ApePat.Text)) = 0 Then
      MsgBox "Debe ingresar el Apellido Paterno.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_ApePat)
      Exit Sub
   End If
   
   If Len(Trim(txt_ApeMat.Text)) = 0 Then
      MsgBox "Debe ingresar el Apellido Materno.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_ApeMat)
      Exit Sub
   End If
   
   If Len(Trim(txt_Nombre.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Nombre)
      Exit Sub
   End If
   
'   If Len(Trim(txt_DirEle.Text)) = 0 Then
'      MsgBox "Debe ingresar el E-mail del cliente.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(txt_DirEle)
'      Exit Sub
'   End If
'   If Not gf_ValidarEmail(txt_DirEle.Text) Then
'      MsgBox "El E-mail del cliente no tiene el formato correcto.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(txt_DirEle)
'      Exit Sub
'   End If
   
'   If cmb_TipVia.ListIndex = -1 Then
'      MsgBox "Debe seleccionar el Tipo de Vía.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(cmb_TipVia)
'      Exit Sub
'   End If
   
'   If Len(Trim(txt_NomVia.Text)) = 0 Then
'      MsgBox "Debe ingresar el Nombre de Vía.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(txt_NomVia)
'      Exit Sub
'   End If
   
'   If Len(Trim(txt_NumVia.Text)) = 0 Then
'      MsgBox "Debe ingresar el Número.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(txt_NumVia)
'      Exit Sub
'   End If
   
'   If cmb_TipZon.ListIndex = -1 Then
'      MsgBox "Debe seleccionar el Tipo de Zona.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(cmb_TipZon)
'      Exit Sub
'   End If
   
'   If cmb_TipZon.ItemData(cmb_TipZon.ListIndex) <> 12 Then
'     If Len(Trim(txt_NomZon.Text)) = 0 Then
'        MsgBox "Debe ingresar el Nombre de Zona.", vbExclamation, modgen_g_str_NomPlt
'        Call gs_SetFocus(txt_NomZon)
'        Exit Sub
'     End If
'   End If
    
'   If cmb_DptDir.ListIndex = -1 Then
'       MsgBox "Debe seleccionar el Departamento de la Dirección.", vbExclamation, modgen_g_str_NomPlt
'       Call gs_SetFocus(cmb_DptDir)
'       Exit Sub
'   End If
    
'   If cmb_PrvDir.ListIndex = -1 Then
'       MsgBox "Debe seleccionar la Provincia de la Dirección.", vbExclamation, modgen_g_str_NomPlt
'       Call gs_SetFocus(cmb_PrvDir)
'       Exit Sub
'   End If
    
'   If cmb_DstDir.ListIndex = -1 Then
'       MsgBox "Debe seleccionar el Distrito de la Dirección.", vbExclamation, modgen_g_str_NomPlt
'       Call gs_SetFocus(cmb_DstDir)
'       Exit Sub
'   End If
   
'   If CDbl(ipp_ImpBFH.Value) = 0 Then
'       MsgBox "Debe ingresar Importe BFH o Ahorro.", vbExclamation, modgen_g_str_NomPlt
'       Call gs_SetFocus(ipp_ImpBFH)
'       Exit Sub
'   End If
    
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
    'Grabando Información del Beneficiaario
   g_str_Parame = "USP_TPR_DATBEN ("
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_DesIte & "', " 'moddat_g_str_NumFia
   
   If moddat_g_int_FlgAct_2 = 2 Then
      g_str_Parame = g_str_Parame & "'" & CStr(l_int_Codigo) & "', "
   Else
      g_str_Parame = g_str_Parame & "'" & CStr(0) & "', "
   End If
   g_str_Parame = g_str_Parame & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & ", "
   g_str_Parame = g_str_Parame & "'" & Trim(txt_NumDoc.Text) & "', "
   g_str_Parame = g_str_Parame & "'" & txt_ApePat & "', "
   g_str_Parame = g_str_Parame & "'" & txt_ApeMat & "', "
   g_str_Parame = g_str_Parame & "'" & txt_ApeCas & "', "
   g_str_Parame = g_str_Parame & "'" & txt_Nombre & "', "
   g_str_Parame = g_str_Parame & "'" & CStr(txt_CodPry.Text) & "', "
   g_str_Parame = g_str_Parame & "'" & txt_TelFij.Text & "', "
   g_str_Parame = g_str_Parame & "'" & txt_TelCel.Text & "', "
   g_str_Parame = g_str_Parame & "'" & txt_DirEle.Text & "', "
   If cmb_TipVia.ListIndex <> -1 Then
      g_str_Parame = g_str_Parame & CStr(cmb_TipVia.ItemData(cmb_TipVia.ListIndex)) & ", "
   Else
      g_str_Parame = g_str_Parame & CStr(0) & ", "
   End If
   g_str_Parame = g_str_Parame & "'" & txt_NomVia.Text & "', "
   g_str_Parame = g_str_Parame & "'" & txt_NumVia.Text & "', "
   g_str_Parame = g_str_Parame & "'" & txt_IntDpt.Text & "', "
   If cmb_TipZon.ListIndex <> -1 Then
      g_str_Parame = g_str_Parame & CStr(cmb_TipZon.ItemData(cmb_TipZon.ListIndex)) & ", "
   Else
      g_str_Parame = g_str_Parame & CStr(0) & ", "
   End If
   g_str_Parame = g_str_Parame & "'" & txt_NomZon.Text & "', "
   g_str_Parame = g_str_Parame & "'" & txt_Refere.Text & "', "
   If cmb_DptDir.ListIndex <> -1 Then
      g_str_Parame = g_str_Parame & "'" & Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00") & Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00") & Format(cmb_DstDir.ItemData(cmb_DstDir.ListIndex), "00") & "', "
   Else
      g_str_Parame = g_str_Parame & "'', "
   End If
   g_str_Parame = g_str_Parame & CDbl(ipp_ImpBFH.Value) & ", "
   
   If l_arr_TipRec(cmb_TipRec.ListIndex + 1).Genera_Codigo <> "" Then
      g_str_Parame = g_str_Parame & CStr(l_arr_TipRec(cmb_TipRec.ListIndex + 1).Genera_Codigo) & ", "
   Else
      g_str_Parame = g_str_Parame & 0 & ", "
   End If
   
   g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgAct_2) & ", "

   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar el Procedimiento USP_TPR_DATBEN.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_int_FlgAct_2 = 2
   
   MsgBox "Los datos se grabaron correctamente.", vbInformation, modgen_g_str_NomPlt
   
   Call fs_Limpia
   Call fs_Buscar
   Call fs_Activa(False)
End Sub

Private Sub cmd_Import_Click()
Dim r_str_NomArc     As String
' On Error GoTo cmd_BusArc_Error
   
   dlg_Guarda.Filter = "Archivos Excel |*.xlsx;*.xls"
   dlg_Guarda.ShowOpen
   r_str_NomArc = UCase(dlg_Guarda.FileName)
'   Exit Sub
   
   If Len(Trim(r_str_NomArc)) = 0 Then
      MsgBox "Debe seleccionar el archivo a importar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Import)
      Exit Sub
   End If
      
   If MsgBox("¿Desea realizar la carga del archivo seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenImp(r_str_NomArc)
   Screen.MousePointer = 0
End Sub
Private Sub fs_GenImp(ByVal p_NomArc As String)

Dim r_Fila        As Integer
Dim r_int_ConSel  As Integer
Dim r_int_Contad  As Integer

   'validaciones
   Screen.MousePointer = 11
   
   If fs_Carga_ArchivoBeneficiarios(p_NomArc) Then
       MsgBox "Proceso de carga de archivo finalizada satisfactoriamente.", vbInformation, modgen_g_str_NomPlt
       Call fs_Buscar
   End If
   
   Screen.MousePointer = 0
End Sub
Private Function fs_Carga_ArchivoBeneficiarios(ByVal p_NomArc As String) As Boolean
Dim r_obj_Excel       As Excel.Application
Dim r_int_FilExc      As Integer
Dim r_str_Codigo      As String
Dim r_str_TipDoc      As String
Dim r_str_NumDoc      As String
Dim r_str_ApePat      As String
Dim r_str_ApeMat      As String
Dim r_str_Nombre      As String
Dim r_str_CodPry      As String
Dim r_dbl_ImpBHF      As Double
Dim r_int_TipRec      As Integer

Dim r_lng_Contad      As Long
Dim r_lng_NumReg      As Long

   fs_Carga_ArchivoBeneficiarios = False
   
   'Abriendo Archivo COFIDE
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Open FileName:=p_NomArc
    
'    For r_lng_NumReg = 1 To r_obj_Excel.Sheets.Count
      
   ReDim r_str_DatBen(0)
        
   r_int_FilExc = 2
   'r_lng_NumReg
   If moddat_g_int_TipRep = 1 Then
      r_obj_Excel.Sheets(1).Select
   Else
      r_obj_Excel.Sheets(2).Select
   End If
       
   Do While Trim(r_obj_Excel.Cells(r_int_FilExc, 1).Value) <> ""
   
      '        r_str_Codigo = Trim(r_obj_Excel.Cells(r_int_FilExc, 1).Value)
      r_str_TipDoc = IIf(Trim(r_obj_Excel.Cells(r_int_FilExc, 2).Value) = "DNI", 1, 0)
      r_str_NumDoc = Trim(r_obj_Excel.Cells(r_int_FilExc, 3).Value)
      r_str_ApePat = Trim(r_obj_Excel.Cells(r_int_FilExc, 4).Value)
      r_str_ApeMat = Trim(r_obj_Excel.Cells(r_int_FilExc, 5).Value)
      r_str_Nombre = Trim(r_obj_Excel.Cells(r_int_FilExc, 6).Value)
      
      If moddat_g_str_CodPrd = "026" And moddat_g_str_CodSub = "001" Then
         r_dbl_ImpBHF = Trim(r_obj_Excel.Cells(r_int_FilExc, 7).Value)
      Else
         r_str_CodPry = Trim(r_obj_Excel.Cells(r_int_FilExc, 7).Value) 'Mid(Trim(r_obj_Excel.Cells(r_int_FilExc, 7).Value), InStrRev(Trim(r_obj_Excel.Cells(r_int_FilExc, 7).Value), "-") + 1)
         r_dbl_ImpBHF = Trim(r_obj_Excel.Cells(r_int_FilExc, 8).Value)
      End If
   
'      If r_obj_Excel.Sheets(r_lng_NumReg).Name = "BFH" Then
'         r_int_TipRec = 1
'      ElseIf r_obj_Excel.Sheets(r_lng_NumReg).Name = "AHORRO" Then
'         r_int_TipRec = 2
'      End If
   
      ReDim Preserve r_str_DatBen(UBound(r_str_DatBen) + 1)
      '        r_str_DatBen(UBound(r_str_DatBen)).Codigo = r_str_Codigo
      r_str_DatBen(UBound(r_str_DatBen)).Tipdoc = r_str_TipDoc
      r_str_DatBen(UBound(r_str_DatBen)).NumDoc = r_str_NumDoc
      r_str_DatBen(UBound(r_str_DatBen)).ApePat = r_str_ApePat
      r_str_DatBen(UBound(r_str_DatBen)).ApeMat = r_str_ApeMat
      r_str_DatBen(UBound(r_str_DatBen)).Nombre = r_str_Nombre
      r_str_DatBen(UBound(r_str_DatBen)).CodPry = r_str_CodPry
      r_str_DatBen(UBound(r_str_DatBen)).ImpBHF = r_dbl_ImpBHF
      r_str_DatBen(UBound(r_str_DatBen)).TipRec = moddat_g_int_TipRep 'r_int_TipRec
      r_int_FilExc = r_int_FilExc + 1
   Loop
   
   For r_lng_Contad = 1 To UBound(r_str_DatBen)
   
      'Grabando Información del Beneficiaario
      g_str_Parame = "USP_TPR_DATBEN ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_DesIte & "', " 'moddat_g_str_NumFia
      g_str_Parame = g_str_Parame & "'" & r_str_DatBen(r_lng_Contad).Codigo & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_DatBen(r_lng_Contad).Tipdoc & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_DatBen(r_lng_Contad).NumDoc & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_DatBen(r_lng_Contad).ApePat & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_DatBen(r_lng_Contad).ApeMat & "', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'" & r_str_DatBen(r_lng_Contad).Nombre & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_DatBen(r_lng_Contad).CodPry & "', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & CDbl(r_str_DatBen(r_lng_Contad).ImpBHF) & ", "
      g_str_Parame = g_str_Parame & CInt(r_str_DatBen(r_lng_Contad).TipRec) & ", "
      g_str_Parame = g_str_Parame & CStr(1) & ", "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         MsgBox "Error al ejecutar el Procedimiento USP_TPR_DATBEN.", vbCritical, modgen_g_str_NomPlt
         Exit Function
      End If
   
      DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
   
   Next r_lng_Contad

'   Next r_lng_NumReg
   
   r_obj_Excel.Workbooks.Close
   Set r_obj_Excel = Nothing
   
   fs_Carga_ArchivoBeneficiarios = True
End Function
Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   Call gs_CentraForm(Me)
   Call fs_Inicia
   
   If moddat_g_int_FlgGrb_1 = 8 Then 'moddat_g_int_FlgGrb = 1
      Call fs_Activa(False)
      cmd_Cancel.Enabled = False
      Call fs_Limpia
   Else
      Call fs_Limpia
      Call fs_Activa(True)
   End If
   
   Call fs_Buscar
   
   If moddat_g_str_DesObs <> "VIGENTE" Then
      cmd_Agrega.Enabled = False
      cmd_Editar.Enabled = False
      cmd_Borrar.Enabled = False
   Else
      moddat_g_str_DesObs = ""
   End If
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   
   grd_Listad.ColWidth(0) = 975
   grd_Listad.ColWidth(1) = 1560
   grd_Listad.ColWidth(2) = 2425
   grd_Listad.ColWidth(3) = 2425
   grd_Listad.ColWidth(4) = 2425
   grd_Listad.ColWidth(5) = 1200
   grd_Listad.ColWidth(6) = 1630
   grd_Listad.ColWidth(7) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
      
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "230")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipVia, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipZon, 1, "202")
   Call moddat_gs_Carga_Depart(cmb_DptDir)
   
     'Recurso
   cmb_TipRec.Clear
   ReDim l_arr_TipRec(0)
   
   ReDim Preserve l_arr_TipRec(UBound(l_arr_TipRec) + 1)
   cmb_TipRec.AddItem Trim$("BONO")
   l_arr_TipRec(UBound(l_arr_TipRec)).Genera_Codigo = CInt(1)
   l_arr_TipRec(UBound(l_arr_TipRec)).Genera_Nombre = Trim$("BONO")
   
   ReDim Preserve l_arr_TipRec(UBound(l_arr_TipRec) + 1)
   cmb_TipRec.AddItem Trim$("AHORRO")
   l_arr_TipRec(UBound(l_arr_TipRec)).Genera_Codigo = CInt(2)
   l_arr_TipRec(UBound(l_arr_TipRec)).Genera_Nombre = Trim$("AHORRO")
   
'   ReDim Preserve l_arr_TipRec(UBound(l_arr_TipRec) + 1)
'   cmb_TipRec.AddItem Trim$("ABONO/AHORRO")
'   l_arr_TipRec(UBound(l_arr_TipRec)).Genera_Codigo = CInt(3)
'   l_arr_TipRec(UBound(l_arr_TipRec)).Genera_Nombre = Trim$("ABONO/AHORRO")

   Call gs_LimpiaGrid(grd_Listad)
   
   pnl_TipDoc.Caption = moddat_gf_Consulta_ParDes("118", moddat_g_int_TipDoc)
   pnl_NroDoc.Caption = moddat_g_str_NumDoc
   pnl_RazSoc.Caption = moddat_g_str_NomCli
   pnl_TipEmp.Caption = moddat_g_str_Descri
   pnl_NumRef.Caption = gf_Formato_NumRef(moddat_g_str_NumFia, Mid(moddat_g_str_NumFia, 1, 1))
   
End Sub
'Private Function fs_Formato_NumRef(ByVal p_Numref As String) As String
'   p_Numref = Format(p_Numref, "0000000000")
'   'fs_Formato_NumRef = Left(p_Numref, 4) & "-" & Mid(p_Numref, 5, 2) & "-" & Right(p_Numref, 4)
'   fs_Formato_NumRef = Mid(p_Numref, 1, 1) & Mid(p_Numref, 2, 2) & "-" & Mid(p_Numref, 4, 2) & "-" & Right(p_Numref, 5)
'End Function
Private Sub fs_Limpia()
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   txt_CodPry.Text = ""
   txt_ApePat.Text = ""
   txt_ApeMat.Text = ""
   txt_ApeCas.Text = ""
   txt_Nombre.Text = ""
   
   txt_TelCel.Text = ""
   txt_TelFij.Text = ""
   txt_DirEle.Text = ""
   
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
   ipp_ImpBFH.Text = 0#
   
   cmb_TipRec.ListIndex = -1
End Sub
Private Sub fs_Activa(ByVal p_Habilita As Integer)
   cmb_TipDoc.Enabled = p_Habilita
   txt_NumDoc.Enabled = p_Habilita
   txt_CodPry.Enabled = p_Habilita
   txt_ApePat.Enabled = p_Habilita
   txt_ApeMat.Enabled = p_Habilita
   txt_ApeCas.Enabled = p_Habilita
   txt_Nombre.Enabled = p_Habilita
   
   txt_TelFij.Enabled = p_Habilita
   txt_TelCel.Enabled = p_Habilita
   txt_DirEle.Enabled = p_Habilita
   
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
   ipp_ImpBFH.Enabled = p_Habilita
   cmb_TipRec.Enabled = p_Habilita
   
   cmd_Agrega.Enabled = Not p_Habilita
   If Me.grd_Listad.Row < 0 Then
      cmd_Editar.Enabled = p_Habilita
      cmd_Borrar.Enabled = p_Habilita
      cmd_ExpExc.Enabled = p_Habilita
   Else
      cmd_Editar.Enabled = Not p_Habilita
      cmd_Borrar.Enabled = Not p_Habilita
      cmd_ExpExc.Enabled = Not p_Habilita
   End If
   
   cmd_Import.Enabled = Not p_Habilita
   cmd_Grabar.Enabled = p_Habilita
   cmd_Cancel.Enabled = p_Habilita
  
End Sub

Private Sub fs_Buscar()

   cmd_Agrega.Enabled = True
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   grd_Listad.Enabled = False
   
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT DATBEN_CODIGO, DATBEN_TIPDOC, DATBEN_NUMDOC, DATBEN_APEPAT, DATBEN_APEMAT, DATBEN_APECAS,  "
   g_str_Parame = g_str_Parame & "       DATBEN_NOMBRE, DATBEN_TIPVIA, DATBEN_NOMVIA, DATBEN_NUMERO, DATBEN_INTDPT, DATBEN_TIPZON,  "
   g_str_Parame = g_str_Parame & "       DATBEN_NOMZON, DATBEN_REFERE, DATBEN_UBIGEO, DATBEN_NUMCEL, DATBEN_TELFIJ, DATBEN_DIRELE, "
   g_str_Parame = g_str_Parame & "       DATBEN_IMPBFH, DATBEN_TIPREC "
   g_str_Parame = g_str_Parame & "  FROM TPR_DATBEN "
   g_str_Parame = g_str_Parame & " WHERE DATBEN_NUMREF = '" & moddat_g_str_DesIte & "'" 'moddat_g_str_NumFia
   g_str_Parame = g_str_Parame & " ORDER BY DATBEN_NUMREF, DATBEN_CODIGO ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
     Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = grd_Listad.Row + 1 'Trim(g_rst_Princi!DATBEN_CODIGO)
            
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!DATBEN_TIPDOC) & " - " & Trim(g_rst_Princi!DATBEN_NUMDOC)
      
      grd_Listad.Col = 2
      If Not IsNull((g_rst_Princi!DATBEN_APEPAT)) Then
         grd_Listad.Text = Trim(Trim(g_rst_Princi!DATBEN_APEPAT) & " " & Trim(g_rst_Princi!DATBEN_APECAS))
      Else
         grd_Listad.Text = ""
      End If
      
      grd_Listad.Col = 3
      If Not IsNull((g_rst_Princi!DATBEN_APEMAT)) Then
         grd_Listad.Text = Trim(g_rst_Princi!DATBEN_APEMAT)
      Else
         grd_Listad.Text = ""
      End If
      
      grd_Listad.Col = 4
      If Not IsNull((g_rst_Princi!DATBEN_NOMBRE)) Then
         grd_Listad.Text = Trim(g_rst_Princi!DATBEN_NOMBRE)
      Else
         grd_Listad.Text = ""
      End If
      
      grd_Listad.Col = 5
      If Not IsNull((g_rst_Princi!DATBEN_TIPREC)) Then
         grd_Listad.Text = l_arr_TipRec(g_rst_Princi!DATBEN_TIPREC).Genera_Nombre
      Else
         grd_Listad.Text = ""
      End If
      
      grd_Listad.Col = 6
      grd_Listad.Text = "S/ " & Format(CStr(g_rst_Princi!DATBEN_IMPBFH), "###,###,###,##0.00")
      
      grd_Listad.Col = 7
       grd_Listad.Text = Trim(g_rst_Princi!DATBEN_CODIGO)
       
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If grd_Listad.Rows > 0 Then
      cmd_Editar.Enabled = True
      cmd_Borrar.Enabled = True
      cmd_ExpExc.Enabled = True
      grd_Listad.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub grd_Listad_Click()
   cmd_Editar_Click
End Sub

Private Sub ipp_ImpBFH_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub


Private Sub pnl_Tit_ApeMat_Click()
    If Len(Trim(pnl_Tit_ApeMat.Tag)) = 0 Or pnl_Tit_ApeMat.Tag = "D" Then
        pnl_Tit_ApeMat.Tag = "A"
        Call gs_SorteaGrid(grd_Listad, 3, "C")
    Else
        pnl_Tit_ApeMat.Tag = "D"
        Call gs_SorteaGrid(grd_Listad, 3, "C-")
    End If
End Sub

Private Sub pnl_Tit_ApePat_Click()
    If Len(Trim(pnl_Tit_ApePat.Tag)) = 0 Or pnl_Tit_ApePat.Tag = "D" Then
        pnl_Tit_ApePat.Tag = "A"
        Call gs_SorteaGrid(grd_Listad, 2, "C")
    Else
        pnl_Tit_ApePat.Tag = "D"
        Call gs_SorteaGrid(grd_Listad, 2, "C-")
    End If
End Sub

Private Sub pnl_Tit_CodPry_Click()
   If Len(Trim(pnl_Tit_Importe.Tag)) = 0 Or pnl_Tit_Importe.Tag = "D" Then
      pnl_Tit_Importe.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "N")
   Else
      pnl_Tit_Importe.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "N-")
   End If
End Sub

Private Sub pnl_Tit_Importe_Click()
   If Len(Trim(pnl_Tit_Importe.Tag)) = 0 Or pnl_Tit_Importe.Tag = "D" Then
      pnl_Tit_Importe.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 6, "N")
   Else
      pnl_Tit_Importe.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 6, "N-")
   End If
End Sub

Private Sub pnl_Tit_Nombre_Click()
    If Len(Trim(pnl_Tit_Nombre.Tag)) = 0 Or pnl_Tit_Nombre.Tag = "D" Then
        pnl_Tit_Nombre.Tag = "A"
        Call gs_SorteaGrid(grd_Listad, 4, "C")
    Else
        pnl_Tit_Nombre.Tag = "D"
        Call gs_SorteaGrid(grd_Listad, 4, "C-")
    End If
End Sub

Private Sub pnl_Tit_NumDoc_Click()
    If Len(Trim(pnl_Tit_NumDoc.Tag)) = 0 Or pnl_Tit_NumDoc.Tag = "D" Then
        pnl_Tit_NumDoc.Tag = "A"
        Call gs_SorteaGrid(grd_Listad, 1, "C")
    Else
        pnl_Tit_NumDoc.Tag = "D"
        Call gs_SorteaGrid(grd_Listad, 1, "C-")
    End If
End Sub

Private Sub pnl_Tit_Recurso_Click()
    If Len(Trim(pnl_Tit_Recurso.Tag)) = 0 Or pnl_Tit_Recurso.Tag = "D" Then
        pnl_Tit_Recurso.Tag = "A"
        Call gs_SorteaGrid(grd_Listad, 5, "C")
    Else
        pnl_Tit_Recurso.Tag = "D"
        Call gs_SorteaGrid(grd_Listad, 5, "C-")
    End If
End Sub

Private Sub txt_ApeCas_GotFocus()
   Call gs_SelecTodo(txt_ApeCas)
End Sub

Private Sub txt_ApeCas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Nombre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_ApeMat_GotFocus()
    Call gs_SelecTodo(txt_ApeMat)
End Sub

Private Sub txt_ApeMat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ApeCas)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_ApePat_GotFocus()
   Call gs_SelecTodo(txt_ApePat)
End Sub

Private Sub txt_ApePat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ApeMat)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_CodPry_GotFocus()
   Call gs_SelecTodo(txt_CodPry)
End Sub

Private Sub txt_CodPry_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipRec)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_DirEle_GotFocus()
   Call gs_SelecTodo(txt_DirEle)
End Sub

Private Sub txt_DirEle_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipVia)
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

Private Sub txt_Nombre_GotFocus()
   Call gs_SelecTodo(txt_Nombre)
End Sub

Private Sub txt_Nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_TelFij)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
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

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_CodPry) 'txt_ApePat
   Else
      If cmb_TipDoc.ListIndex > -1 Then
         Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
            Case 1:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case 2:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 3:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 4:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
         End Select
      Else
         KeyAscii = 0
      End If
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
      Call gs_SetFocus(ipp_ImpBFH) 'cmd_Grabar
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_TelCel_GotFocus()
   Call gs_SelecTodo(txt_DirEle)
End Sub

Private Sub txt_TelCel_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_DirEle)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_TelFij_GotFocus()
   Call gs_SelecTodo(txt_TelFij)
End Sub

Private Sub txt_TelFij_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_TelCel)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub
