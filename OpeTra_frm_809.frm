VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frm_RegDes_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9690
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11580
   Icon            =   "OpeTra_frm_809.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9690
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel111 
      Height          =   10095
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11610
      _Version        =   65536
      _ExtentX        =   20479
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   645
         Left            =   30
         TabIndex        =   15
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
            Picture         =   "OpeTra_frm_809.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Rechazar 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_809.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Rechazar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10875
            Picture         =   "OpeTra_frm_809.frx":0890
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Aprobar 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_809.frx":0CD2
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Aprobar"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   675
         Left            =   30
         TabIndex        =   18
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
         Begin Threed.SSPanel pnl_Titulo 
            Height          =   555
            Left            =   720
            TabIndex        =   19
            Top             =   30
            Width           =   7395
            _Version        =   65536
            _ExtentX        =   13044
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "Créditos Hipotecarios - Registro legal"
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
         Begin MSMAPI.MAPISession mps_Sesion 
            Left            =   10350
            Top             =   60
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   10935
            Top             =   60
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            AddressEditFieldCount=   1
            AddressModifiable=   0   'False
            AddressResolveUI=   0   'False
            FetchSorted     =   0   'False
            FetchUnreadOnly =   0   'False
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_809.frx":0FDC
            Top             =   120
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   6135
         Left            =   30
         TabIndex        =   20
         Top             =   2250
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   10821
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
            Height          =   5985
            Index           =   1
            Left            =   60
            TabIndex        =   23
            Top             =   60
            Width           =   11400
            _ExtentX        =   20108
            _ExtentY        =   10557
            _Version        =   393216
            Style           =   1
            Tabs            =   9
            Tab             =   7
            TabsPerRow      =   9
            TabHeight       =   520
            TabCaption(0)   =   "Cliente"
            TabPicture(0)   =   "OpeTra_frm_809.frx":12E6
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Cónyuge"
            TabPicture(1)   =   "OpeTra_frm_809.frx":1302
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(2)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Inmueble"
            TabPicture(2)   =   "OpeTra_frm_809.frx":131E
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(1)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Crédito"
            TabPicture(3)   =   "OpeTra_frm_809.frx":133A
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "grd_Listad(4)"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Desembolso"
            TabPicture(4)   =   "OpeTra_frm_809.frx":1356
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "grd_Listad(3)"
            Tab(4).Control(1)=   "txt_ObsDes"
            Tab(4).ControlCount=   2
            TabCaption(5)   =   "Informe Legal"
            TabPicture(5)   =   "OpeTra_frm_809.frx":1372
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "txt_InfLeg"
            Tab(5).ControlCount=   1
            TabCaption(6)   =   "Ev. Legal"
            TabPicture(6)   =   "OpeTra_frm_809.frx":138E
            Tab(6).ControlEnabled=   0   'False
            Tab(6).Control(0)=   "Label7"
            Tab(6).Control(1)=   "Label5"
            Tab(6).Control(2)=   "grd_Listad(6)"
            Tab(6).Control(3)=   "txt_ComCre"
            Tab(6).ControlCount=   4
            TabCaption(7)   =   "Datos del Desembolso"
            TabPicture(7)   =   "OpeTra_frm_809.frx":13AA
            Tab(7).ControlEnabled=   -1  'True
            Tab(7).Control(0)=   "Label11"
            Tab(7).Control(0).Enabled=   0   'False
            Tab(7).Control(1)=   "SSPanel21"
            Tab(7).Control(1).Enabled=   0   'False
            Tab(7).Control(2)=   "SSPanel22"
            Tab(7).Control(2).Enabled=   0   'False
            Tab(7).Control(3)=   "SSPanel15"
            Tab(7).Control(3).Enabled=   0   'False
            Tab(7).Control(4)=   "pnl_Prycto_Dsm"
            Tab(7).Control(4).Enabled=   0   'False
            Tab(7).Control(5)=   "cmd_Dsm_ExpExc"
            Tab(7).Control(5).Enabled=   0   'False
            Tab(7).ControlCount=   6
            TabCaption(8)   =   "Seguimientos"
            TabPicture(8)   =   "OpeTra_frm_809.frx":13C6
            Tab(8).ControlEnabled=   0   'False
            Tab(8).Control(0)=   "grd_Listad(7)"
            Tab(8).ControlCount=   1
            Begin VB.TextBox txt_ObsDes 
               Height          =   975
               Left            =   -74910
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   59
               Top             =   4935
               Width           =   11235
            End
            Begin VB.TextBox txt_InfLeg 
               Height          =   5475
               Left            =   -74910
               MaxLength       =   8000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   58
               Top             =   420
               Width           =   11200
            End
            Begin VB.TextBox txt_ComCre 
               Height          =   705
               Left            =   -74940
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   54
               Top             =   675
               Width           =   11200
            End
            Begin VB.CommandButton cmd_Dsm_ExpExc 
               Height          =   450
               Left            =   10710
               Picture         =   "OpeTra_frm_809.frx":13E2
               Style           =   1  'Graphical
               TabIndex        =   43
               ToolTipText     =   "Exportar a Excel"
               Top             =   330
               Width           =   585
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   5535
               Index           =   0
               Left            =   -74910
               TabIndex        =   24
               Top             =   370
               Width           =   11250
               _ExtentX        =   19844
               _ExtentY        =   9763
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
               Height          =   5535
               Index           =   7
               Left            =   -74910
               TabIndex        =   42
               Top             =   370
               Width           =   11250
               _ExtentX        =   19844
               _ExtentY        =   9763
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel pnl_Prycto_Dsm 
               Height          =   315
               Left            =   1650
               TabIndex        =   44
               Top             =   420
               Width           =   6660
               _Version        =   65536
               _ExtentX        =   11747
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
            Begin Threed.SSPanel SSPanel15 
               Height          =   645
               Left            =   60
               TabIndex        =   45
               Top             =   2835
               Width           =   11235
               _Version        =   65536
               _ExtentX        =   19817
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
               Begin VB.CommandButton cmd_Dsm_Borrar 
                  Height          =   585
                  Left            =   9990
                  Picture         =   "OpeTra_frm_809.frx":16EC
                  Style           =   1  'Graphical
                  TabIndex        =   11
                  ToolTipText     =   "Eliminar Registro"
                  Top             =   40
                  Width           =   585
               End
               Begin VB.CommandButton cmd_Dsm_Nuevo 
                  Height          =   585
                  Left            =   9390
                  Picture         =   "OpeTra_frm_809.frx":19F6
                  Style           =   1  'Graphical
                  TabIndex        =   0
                  ToolTipText     =   "Adicionar Registro"
                  Top             =   40
                  Width           =   585
               End
               Begin VB.CommandButton cmd_Dsm_Editar 
                  Height          =   585
                  Left            =   10590
                  Picture         =   "OpeTra_frm_809.frx":1D00
                  Style           =   1  'Graphical
                  TabIndex        =   12
                  ToolTipText     =   "Modificar Registro"
                  Top             =   40
                  Width           =   585
               End
            End
            Begin Threed.SSPanel SSPanel22 
               Height          =   2010
               Left            =   60
               TabIndex        =   46
               Top             =   810
               Width           =   11250
               _Version        =   65536
               _ExtentX        =   19844
               _ExtentY        =   3545
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
               Begin MSFlexGridLib.MSFlexGrid grd_Listad_Dsm 
                  Height          =   1650
                  Left            =   0
                  TabIndex        =   47
                  Top             =   20
                  Width           =   11230
                  _ExtentX        =   19817
                  _ExtentY        =   2910
                  _Version        =   393216
                  Rows            =   1
                  Cols            =   14
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  SelectionMode   =   1
                  Appearance      =   0
               End
               Begin Threed.SSPanel pnl_SumTot_Dsm 
                  Height          =   285
                  Left            =   6060
                  TabIndex        =   48
                  Top             =   1680
                  Width           =   1200
                  _Version        =   65536
                  _ExtentX        =   2117
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "9,999,999.99 "
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
               Begin Threed.SSPanel pnl_TotPtmo_Dsm 
                  Height          =   285
                  Left            =   1095
                  TabIndex        =   49
                  Top             =   1680
                  Width           =   1200
                  _Version        =   65536
                  _ExtentX        =   2117
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "9,999,999.99 "
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
               Begin VB.Label lbl_Totale 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Distribuir ==> "
                  Height          =   195
                  Index           =   1
                  Left            =   135
                  TabIndex        =   52
                  Top             =   1740
                  Width           =   960
               End
               Begin VB.Label lbl_Total 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Total ==> "
                  Height          =   195
                  Left            =   5325
                  TabIndex        =   51
                  Top             =   1740
                  Width           =   720
               End
               Begin VB.Label lbl_Bono_Dsm 
                  AutoSize        =   -1  'True
                  Caption         =   ".."
                  Height          =   195
                  Left            =   2400
                  TabIndex        =   50
                  Top             =   1740
                  Width           =   90
               End
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   4275
               Index           =   6
               Left            =   -74940
               TabIndex        =   55
               Top             =   1635
               Width           =   11250
               _ExtentX        =   19844
               _ExtentY        =   7541
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
               Height          =   4485
               Index           =   3
               Left            =   -74910
               TabIndex        =   60
               Top             =   420
               Width           =   11250
               _ExtentX        =   19844
               _ExtentY        =   7911
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
               Height          =   5535
               Index           =   4
               Left            =   -74910
               TabIndex        =   61
               Top             =   370
               Width           =   11250
               _ExtentX        =   19844
               _ExtentY        =   9763
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
               Height          =   5535
               Index           =   1
               Left            =   -74910
               TabIndex        =   62
               Top             =   370
               Width           =   11250
               _ExtentX        =   19844
               _ExtentY        =   9763
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
               Height          =   5535
               Index           =   2
               Left            =   -74910
               TabIndex        =   63
               Top             =   370
               Width           =   11250
               _ExtentX        =   19844
               _ExtentY        =   9763
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel21 
               Height          =   2445
               Left            =   40
               TabIndex        =   64
               Top             =   3480
               Width           =   11265
               _Version        =   65536
               _ExtentX        =   19861
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
               BorderWidth     =   1
               BevelOuter      =   0
               BevelInner      =   1
               Begin VB.ComboBox cmb_TipMto_Dsm 
                  Height          =   315
                  Left            =   5880
                  Style           =   2  'Dropdown List
                  TabIndex        =   2
                  Top             =   60
                  Width           =   2970
               End
               Begin VB.TextBox txt_Descrp_Dsm 
                  Height          =   315
                  Left            =   1680
                  MaxLength       =   250
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   7
                  Top             =   1725
                  Width           =   7170
               End
               Begin VB.TextBox txt_NroDsm_Dsm 
                  Height          =   315
                  Left            =   1680
                  MaxLength       =   250
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   8
                  Top             =   2055
                  Width           =   3000
               End
               Begin VB.ComboBox cmb_NroCta_Dsm 
                  Height          =   315
                  Left            =   1680
                  Style           =   2  'Dropdown List
                  TabIndex        =   5
                  Top             =   1065
                  Width           =   3000
               End
               Begin VB.CommandButton cmd_Dsm_Insert 
                  Height          =   585
                  Left            =   10020
                  Picture         =   "OpeTra_frm_809.frx":200A
                  Style           =   1  'Graphical
                  TabIndex        =   10
                  Tag             =   "2"
                  Top             =   1800
                  Width           =   585
               End
               Begin VB.CommandButton cmd_Dsm_Cancel 
                  Height          =   585
                  Left            =   10620
                  Picture         =   "OpeTra_frm_809.frx":2314
                  Style           =   1  'Graphical
                  TabIndex        =   13
                  Top             =   1800
                  Width           =   585
               End
               Begin VB.TextBox txt_ANombre_Dsm 
                  Height          =   315
                  Left            =   1680
                  MaxLength       =   250
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   6
                  Top             =   1395
                  Width           =   7170
               End
               Begin VB.ComboBox cmb_FrmDsm_Dsm 
                  Height          =   315
                  Left            =   1680
                  Style           =   2  'Dropdown List
                  TabIndex        =   1
                  Top             =   60
                  Width           =   3000
               End
               Begin VB.ComboBox cmb_EntFin_Dsm 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   1680
                  Style           =   2  'Dropdown List
                  TabIndex        =   4
                  Top             =   735
                  Width           =   7170
               End
               Begin EditLib.fpDoubleSingle ipp_Import_Dsm 
                  Height          =   315
                  Left            =   1680
                  TabIndex        =   3
                  Top             =   400
                  Width           =   1980
                  _Version        =   196608
                  _ExtentX        =   3492
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
               Begin Threed.SSPanel pnl_Moneda_Dsm 
                  Height          =   315
                  Left            =   5880
                  TabIndex        =   65
                  Top             =   405
                  Width           =   2970
                  _Version        =   65536
                  _ExtentX        =   5239
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
               Begin EditLib.fpDateTime ipp_FecDsm_Dsm 
                  Height          =   315
                  Left            =   6870
                  TabIndex        =   9
                  Top             =   2055
                  Width           =   1980
                  _Version        =   196608
                  _ExtentX        =   3492
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
                  ButtonStyle     =   3
                  ButtonWidth     =   0
                  ButtonWrap      =   -1  'True
                  ButtonDefaultAction=   -1  'True
                  ThreeDText      =   0
                  ThreeDTextHighlightColor=   -2147483637
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   0
                  AlignTextV      =   0
                  AllowNull       =   -1  'True
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
                  Text            =   "24/04/2015"
                  DateCalcMethod  =   0
                  DateTimeFormat  =   0
                  UserDefinedFormat=   ""
                  DateMax         =   "00000000"
                  DateMin         =   "00000000"
                  TimeMax         =   "000000"
                  TimeMin         =   "000000"
                  TimeString1159  =   ""
                  TimeString2359  =   ""
                  DateDefault     =   "00000000"
                  TimeDefault     =   "000000"
                  TimeStyle       =   0
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483637
                  Appearance      =   2
                  BorderDropShadow=   0
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  PopUpType       =   0
                  DateCalcY2KSplit=   60
                  CaretPosition   =   0
                  IncYear         =   1
                  IncMonth        =   1
                  IncDay          =   1
                  IncHour         =   1
                  IncMinute       =   1
                  IncSecond       =   1
                  ButtonColor     =   -2147483637
                  AutoMenu        =   0   'False
                  StartMonth      =   4
                  ButtonAlign     =   0
                  BoundDataType   =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin Threed.SSPanel pnl_NroCCI_Dsm 
                  Height          =   315
                  Left            =   5880
                  TabIndex        =   66
                  Top             =   1065
                  Width           =   2970
                  _Version        =   65536
                  _ExtentX        =   5239
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
               Begin VB.Label Label14 
                  AutoSize        =   -1  'True
                  Caption         =   "Tipo Monto:"
                  Height          =   195
                  Left            =   4980
                  TabIndex        =   77
                  Top             =   135
                  Width           =   855
               End
               Begin VB.Label Label26 
                  AutoSize        =   -1  'True
                  Caption         =   "Descripción:"
                  Height          =   195
                  Left            =   165
                  TabIndex        =   76
                  Top             =   1785
                  Width           =   885
               End
               Begin VB.Label Label25 
                  AutoSize        =   -1  'True
                  Caption         =   "Importe Desembolso:"
                  Height          =   195
                  Left            =   165
                  TabIndex        =   75
                  Top             =   450
                  Width           =   1485
               End
               Begin VB.Label Label23 
                  AutoSize        =   -1  'True
                  Caption         =   "Forma Desembolso:"
                  Height          =   195
                  Left            =   165
                  TabIndex        =   74
                  Top             =   135
                  Width           =   1395
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  Caption         =   "Nro Cuenta:"
                  Height          =   195
                  Left            =   165
                  TabIndex        =   73
                  Top             =   1125
                  Width           =   855
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  Caption         =   "Moneda:"
                  Height          =   195
                  Left            =   4980
                  TabIndex        =   72
                  Top             =   450
                  Width           =   630
               End
               Begin VB.Label Label10 
                  AutoSize        =   -1  'True
                  Caption         =   "A Nombre de:"
                  Height          =   195
                  Left            =   165
                  TabIndex        =   71
                  Top             =   1455
                  Width           =   975
               End
               Begin VB.Label lbl_NumDsm_Dsm 
                  AutoSize        =   -1  'True
                  Caption         =   "Nro Desembolso:"
                  Height          =   195
                  Left            =   165
                  TabIndex        =   70
                  Top             =   2115
                  Width           =   1215
               End
               Begin VB.Label lbl_FchDsm_Dsm 
                  AutoSize        =   -1  'True
                  Caption         =   "Fecha Reg Desemb.:"
                  Height          =   195
                  Left            =   5280
                  TabIndex        =   69
                  Top             =   2115
                  Width           =   1515
               End
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  Caption         =   "Entidad Financiera:"
                  Height          =   195
                  Left            =   165
                  TabIndex        =   68
                  Top             =   795
                  Width           =   1365
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  Caption         =   "Nro CCI:"
                  Height          =   195
                  Left            =   4980
                  TabIndex        =   67
                  Top             =   1125
                  Width           =   600
               End
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Datos de la Evaluación"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   -74940
               TabIndex        =   57
               Top             =   1410
               Width           =   1980
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
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
               Height          =   195
               Left            =   -74940
               TabIndex        =   56
               Top             =   420
               Width           =   1590
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Nombre Proyecto:"
               Height          =   195
               Left            =   210
               TabIndex        =   53
               Top             =   495
               Width           =   1275
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
               TabIndex        =   27
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
               TabIndex        =   26
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
               TabIndex        =   25
               Top             =   1530
               Width           =   2805
            End
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1230
         Left            =   30
         TabIndex        =   28
         Top             =   8430
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   2170
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
         Begin VB.TextBox txt_Comentario 
            Height          =   700
            Left            =   2400
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   29
            Top             =   450
            Width           =   9045
         End
         Begin EditLib.fpDateTime ipp_FecConstancia 
            Height          =   315
            Left            =   2400
            TabIndex        =   32
            Top             =   120
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
            ButtonStyle     =   3
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
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
            Text            =   "24/04/2015"
            DateCalcMethod  =   0
            DateTimeFormat  =   0
            UserDefinedFormat=   ""
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label lblComentario 
            AutoSize        =   -1  'True
            Caption         =   "Comentario:"
            Height          =   195
            Left            =   150
            TabIndex        =   31
            Top             =   690
            Width           =   840
         End
         Begin VB.Label lblFechaConstancia 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Recep. Const. Desemb:"
            Height          =   195
            Left            =   150
            TabIndex        =   30
            Top             =   180
            Width           =   2190
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   33
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
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   900
            TabIndex        =   34
            Top             =   390
            Width           =   6435
            _Version        =   65536
            _ExtentX        =   11351
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "CREDITO HIPOTECARIO MICASITA"
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
         Begin Threed.SSPanel pnl_NumOpe 
            Height          =   315
            Left            =   8730
            TabIndex        =   35
            Top             =   60
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
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
         End
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   900
            TabIndex        =   36
            Top             =   60
            Width           =   6435
            _Version        =   65536
            _ExtentX        =   11351
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
         Begin Threed.SSPanel pnl_EstadoActual 
            Height          =   315
            Left            =   8730
            TabIndex        =   37
            Top             =   390
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
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
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   150
            TabIndex        =   41
            Top             =   120
            Width           =   525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro.Operación:"
            Height          =   195
            Left            =   7560
            TabIndex        =   40
            Top             =   120
            Width           =   1080
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Producto:"
            Height          =   195
            Left            =   150
            TabIndex        =   39
            Top             =   450
            Width           =   690
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Instancia:"
            Height          =   195
            Left            =   7560
            TabIndex        =   38
            Top             =   450
            Width           =   690
         End
      End
   End
End
Attribute VB_Name = "frm_RegDes_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_dbl_MtoHip     As Double
Dim l_str_MonBlq     As String
Dim l_dbl_ImpTas     As Double
Dim l_dbl_ImpNot     As Double
Dim l_dbl_ImpEst     As Double
Dim l_dbl_ImpEva     As Double
Dim l_dbl_ImpAdm     As Double
Dim l_dbl_ImpRed     As Double
Dim l_dbl_ImpBlq     As Double
Dim l_int_ChqReg     As Integer
Dim l_int_PolReg     As Integer
Dim l_int_FiaReg     As Integer
Dim l_int_CerReg     As Integer
Dim l_int_FlgCVt     As Integer
Dim l_int_MonCvt     As Integer
Dim l_str_CodMod     As String
Dim l_str_Moneda     As String
Dim l_dbl_ImpPtm     As Double
Dim l_str_Prmtor     As String
Dim l_str_PryBan     As String
Dim l_str_CodBan     As String
Dim l_arr_CtaBco()   As moddat_tpo_Genera
Dim l_arr_Bancos()   As moddat_tpo_Genera

Private Sub cmb_NroCta_Dsm_Click()
Dim r_int_fila As Integer

   pnl_NroCCI_Dsm.Caption = ""
   For r_int_fila = 1 To UBound(l_arr_CtaBco)
       If (l_arr_CtaBco(r_int_fila).Genera_Codigo = CStr(l_arr_Bancos(cmb_EntFin_Dsm.ListIndex + 1).Genera_Codigo) And _
           Trim(l_arr_CtaBco(r_int_fila).Genera_Nombre) = Trim(cmb_NroCta_Dsm.Text)) Then
           pnl_NroCCI_Dsm.Caption = Trim(l_arr_CtaBco(r_int_fila).Genera_Refere)
           
           If cmb_FrmDsm_Dsm.ItemData(cmb_FrmDsm_Dsm.ListIndex) = 2 Then 'transferencia
              txt_ANombre_Dsm.Text = ""
              txt_ANombre_Dsm.Text = Trim(l_arr_CtaBco(r_int_fila).Genera_NomCli & "")
           End If
           Exit For
       End If
   Next
End Sub

Private Sub cmb_TipMto_Dsm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Import_Dsm)
   End If
End Sub

Private Sub cmd_Grabar_Click()
Dim r_bol_Estado As Boolean

    If Len(Trim(moddat_g_str_NumOpe)) = 0 Then
       MsgBox "Tiene que Haber un Nro. Operación.", vbExclamation, modgen_g_str_NomPlt
       Screen.MousePointer = 0
       Exit Sub
    End If
    
    If (moddat_g_str_CodIte = 6) Then
       If (Trim(ipp_FecConstancia.Text) = Trim("")) Then
          MsgBox "Debe ingresar una fecha.", vbExclamation, modgen_g_str_NomPlt
          Screen.MousePointer = 0
          Call gs_SetFocus(ipp_FecConstancia)
          Exit Sub
       End If
    End If
    
    'Validando
    If (CDbl(pnl_SumTot_Dsm.Caption) > CDbl(CStr(l_dbl_ImpPtm))) Then
       SSTab1(1).Tab = 6
       MsgBox "El préstamo total no es igual al distribuido en la pestaña datos del desembolso", vbExclamation, modgen_g_str_NomPlt
       Exit Sub
    End If
             
    If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
       Screen.MousePointer = 0
       Exit Sub
    End If
    
    Screen.MousePointer = 11
    moddat_g_int_FlgGOK = False
    moddat_g_int_CntErr = 0
    
    Do While moddat_g_int_FlgGOK = False
       Screen.MousePointer = 11
       '--Call moddat_gs_FecSis
       g_str_Parame = ""
       g_str_Parame = g_str_Parame & " UPDATE cre_desprocab SET "
       'tesoreria nroCheque / Tranferencia
       If (moddat_g_str_CodIte = 6) Then '---Tesoreria Aprobada
           g_str_Parame = g_str_Parame & " descab_ferece = '" & Format(ipp_FecConstancia.Text, "yyyymmdd") & "', " 'Fecha de entrega de notaria
           g_str_Parame = g_str_Parame & " descab_cmnop2 = '" & txt_Comentario.Text & "' "
       Else '--legal aprobado(3)
           g_str_Parame = g_str_Parame & " descab_cmnope = '" & txt_Comentario.Text & "' "
       End If
       g_str_Parame = g_str_Parame & " WHERE "
       g_str_Parame = g_str_Parame & " DESCAB_NUMOPE = '" & moddat_g_str_NumOpe & "' and "
       g_str_Parame = g_str_Parame & " DESCAB_FECREG = '" & moddat_g_str_FecRec & "' and "
       g_str_Parame = g_str_Parame & " DESCAB_HORREG = '" & moddat_g_str_FecHip & "' "
       
       If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
          moddat_g_int_CntErr = moddat_g_int_CntErr + 1
          Else
          moddat_g_int_FlgGOK = True
       End If
        
       If moddat_g_int_CntErr = 6 Then
          If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
             Screen.MousePointer = 0
             Exit Sub
          Else
             moddat_g_int_CntErr = 0
          End If
       End If
       Screen.MousePointer = 0
    Loop
          
    r_bol_Estado = fs_guardar_ctaPromotor
    If (r_bol_Estado = False) Then
       Exit Sub
    End If
          
    Screen.MousePointer = 0
    If (moddat_g_int_CntErr = 0) Then
       MsgBox "El proceso se grabó exitosamente.", vbInformation, modgen_g_str_NomPlt
       frm_RegDes_01.fs_Buscar_Creditos
       Unload Me
    End If
End Sub

Private Function fs_guardar_ctaPromotor() As Boolean
Dim r_int_fila As Integer

      'GUARDAR CUENTAS BANCARIAS
      fs_guardar_ctaPromotor = True
      
      If (Len(Trim(moddat_g_str_NumOpe)) > 0 And Len(Trim(moddat_g_str_FecRec)) > 0 And Len(Trim(moddat_g_str_FecHip)) > 0) Then
          For r_int_fila = 1 To grd_Listad_Dsm.Rows - 1
              If (UCase(Trim(grd_Listad_Dsm.TextMatrix(r_int_fila, 12))) = Trim("I")) Then
              
                  g_str_Parame = ""
                  g_str_Parame = "INSERT INTO CRE_DESPRODAT ("
                  g_str_Parame = g_str_Parame & "DESDAT_NUMOPE, "
                  g_str_Parame = g_str_Parame & "DESDAT_FECREG, "
                  g_str_Parame = g_str_Parame & "DESDAT_HORREG, "
                  g_str_Parame = g_str_Parame & "DESDAT_CODBCO, "
                  g_str_Parame = g_str_Parame & "DESDAT_NUMCTA, "
                  g_str_Parame = g_str_Parame & "DESDAT_FRMDES, "
                  g_str_Parame = g_str_Parame & "DESDAT_NUMDES, "
                  g_str_Parame = g_str_Parame & "DESDAT_FCHDES, "
                  g_str_Parame = g_str_Parame & "DESDAT_IMPORT, "
                  g_str_Parame = g_str_Parame & "DESDAT_ANOMBR, "
                  g_str_Parame = g_str_Parame & "DESDAT_DESCRI, "
                  g_str_Parame = g_str_Parame & "DESDAT_NUMITE, "
                  g_str_Parame = g_str_Parame & "DESDAT_TIPMTO, "
                  g_str_Parame = g_str_Parame & "SEGUSUCRE, "
                  g_str_Parame = g_str_Parame & "SEGFECCRE, "
                  g_str_Parame = g_str_Parame & "SEGHORCRE, "
                  g_str_Parame = g_str_Parame & "SEGPLTCRE, "
                  g_str_Parame = g_str_Parame & "SEGTERCRE, "
                  g_str_Parame = g_str_Parame & "SEGSUCCRE, "
                  g_str_Parame = g_str_Parame & "SEGUSUACT, "
                  g_str_Parame = g_str_Parame & "SEGFECACT, "
                  g_str_Parame = g_str_Parame & "SEGHORACT, "
                  g_str_Parame = g_str_Parame & "SEGPLTACT, "
                  g_str_Parame = g_str_Parame & "SEGTERACT, "
                  g_str_Parame = g_str_Parame & "SEGSUCACT) "
                  g_str_Parame = g_str_Parame & "VALUES ( "
                  g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
                  g_str_Parame = g_str_Parame & moddat_g_str_FecRec & ", "
                  g_str_Parame = g_str_Parame & moddat_g_str_FecHip & ", "
                  g_str_Parame = g_str_Parame & "'" & Trim(grd_Listad_Dsm.TextMatrix(r_int_fila, 5)) & "', " 'DESDAT_CODBCO
                  g_str_Parame = g_str_Parame & "'" & Trim(grd_Listad_Dsm.TextMatrix(r_int_fila, 7)) & "', " 'DESDAT_NUMCTA
                  g_str_Parame = g_str_Parame & Trim(grd_Listad_Dsm.TextMatrix(r_int_fila, 0)) & ", " 'DESDAT_FRMDES
                  g_str_Parame = g_str_Parame & "'" & Trim(grd_Listad_Dsm.TextMatrix(r_int_fila, 9)) & "', " 'DESDAT_NUMDES
                  g_str_Parame = g_str_Parame & "'" & Format(Trim(grd_Listad_Dsm.TextMatrix(r_int_fila, 10)), "yyyymmdd") & "', " 'DESDAT_FCHDES
                  g_str_Parame = g_str_Parame & Format(Trim(grd_Listad_Dsm.TextMatrix(r_int_fila, 4)), "########0.00") & ", " 'DESDAT_IMPORT
                  g_str_Parame = g_str_Parame & "'" & Trim(grd_Listad_Dsm.TextMatrix(r_int_fila, 8)) & "', " 'DESDAT_ANOMBR
                  g_str_Parame = g_str_Parame & "'" & Trim(grd_Listad_Dsm.TextMatrix(r_int_fila, 11)) & "', " 'DESDAT_DESCRI
                  g_str_Parame = g_str_Parame & grd_Listad_Dsm.TextMatrix(r_int_fila, 13) & ", " 'DESDAT_NUMITE
                  g_str_Parame = g_str_Parame & grd_Listad_Dsm.TextMatrix(r_int_fila, 2) & ", " 'DESDAT_TIPMTO
                  g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
                  g_str_Parame = g_str_Parame & "'" & Format(date, "YYYYMMDD") & "', "
                  g_str_Parame = g_str_Parame & "'" & Format(Time, "HHMMSS") & "', "
                  g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
                  g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
                  g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
                  g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
                  g_str_Parame = g_str_Parame & "'" & Format(date, "YYYYMMDD") & "', "
                  g_str_Parame = g_str_Parame & "'" & Format(Time, "HHMMSS") & "', "
                  g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
                  g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
                  g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
                  
                  If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
                     moddat_g_int_FlgGOK = False
                  End If
              ElseIf (UCase(Trim(grd_Listad_Dsm.TextMatrix(r_int_fila, 12))) = Trim("U")) Then
                  g_str_Parame = ""
                  g_str_Parame = g_str_Parame & "UPDATE CRE_DESPRODAT SET "
                  g_str_Parame = g_str_Parame & "       DESDAT_FRMDES =  " & grd_Listad_Dsm.TextMatrix(r_int_fila, 0) & ","
                  g_str_Parame = g_str_Parame & "       DESDAT_NUMDES = '" & grd_Listad_Dsm.TextMatrix(r_int_fila, 9) & "',"
                  g_str_Parame = g_str_Parame & "       DESDAT_FCHDES = '" & Format(grd_Listad_Dsm.TextMatrix(r_int_fila, 10), "yyyymmdd") & "',"
                  g_str_Parame = g_str_Parame & "       DESDAT_IMPORT =  " & Format(grd_Listad_Dsm.TextMatrix(r_int_fila, 4), "########0.00") & ","
                  g_str_Parame = g_str_Parame & "       DESDAT_DESCRI = '" & grd_Listad_Dsm.TextMatrix(r_int_fila, 11) & "',"
                  g_str_Parame = g_str_Parame & "       DESDAT_ANOMBR = '" & grd_Listad_Dsm.TextMatrix(r_int_fila, 8) & "',"
                  g_str_Parame = g_str_Parame & "       DESDAT_CODBCO = '" & Trim(grd_Listad_Dsm.TextMatrix(r_int_fila, 5)) & "', "
                  g_str_Parame = g_str_Parame & "       DESDAT_NUMCTA = '" & Trim(grd_Listad_Dsm.TextMatrix(r_int_fila, 7)) & "', "
                  g_str_Parame = g_str_Parame & "       DESDAT_TIPMTO =  " & grd_Listad_Dsm.TextMatrix(r_int_fila, 2) & ","
                  g_str_Parame = g_str_Parame & "       SEGUSUACT='" & modgen_g_str_CodUsu & "',"
                  g_str_Parame = g_str_Parame & "       SEGFECACT='" & Format(date, "YYYYMMDD") & "',"
                  g_str_Parame = g_str_Parame & "       SEGHORACT='" & Format(Time, "HHMMSS") & "',"
                  g_str_Parame = g_str_Parame & "       SEGPLTACT='" & UCase(App.EXEName) & "',"
                  g_str_Parame = g_str_Parame & "       SEGTERACT='" & modgen_g_str_NombPC & "',"
                  g_str_Parame = g_str_Parame & "       SEGSUCACT='" & modgen_g_str_CodSuc & "' "
                  g_str_Parame = g_str_Parame & " WHERE TRIM(DESDAT_NUMOPE) ='" & Trim(moddat_g_str_NumOpe) & "' "
                  g_str_Parame = g_str_Parame & "   AND DESDAT_FECREG = " & moddat_g_str_FecRec
                  g_str_Parame = g_str_Parame & "   AND DESDAT_HORREG = " & moddat_g_str_FecHip
                  g_str_Parame = g_str_Parame & "   AND DESDAT_NUMITE = " & grd_Listad_Dsm.TextMatrix(r_int_fila, 13)

                  If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
                     moddat_g_int_FlgGOK = False
                  End If
              ElseIf (UCase(Trim(grd_Listad_Dsm.TextMatrix(r_int_fila, 12))) = Trim("D")) Then
                  g_str_Parame = ""
                  g_str_Parame = g_str_Parame & "DELETE FROM CRE_DESPRODAT "
                  g_str_Parame = g_str_Parame & " WHERE DESDAT_NUMOPE = '" & Trim(moddat_g_str_NumOpe) & "' "
                  g_str_Parame = g_str_Parame & "   AND DESDAT_FECREG =  " & moddat_g_str_FecRec
                  g_str_Parame = g_str_Parame & "   AND DESDAT_HORREG =  " & moddat_g_str_FecHip
                  g_str_Parame = g_str_Parame & "   AND DESDAT_NUMITE =  " & grd_Listad_Dsm.TextMatrix(r_int_fila, 13)
                  
                  If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
                     moddat_g_int_FlgGOK = False
                  End If
              End If
              
              If moddat_g_int_FlgGOK = False Then
                 Screen.MousePointer = 0
                 MsgBox "No se pudo completar la grabación de los datos.", vbInformation, modgen_g_str_NomPlt
                 fs_guardar_ctaPromotor = False
                 Exit Function
              End If
          Next
      End If
End Function

Private Sub cmd_Rechazar_Click()
Dim r_bol_Estado As Boolean

   If Len(Trim(moddat_g_str_NumOpe)) = 0 Then
      MsgBox "Tiene que Haber un Nro. Operación.", vbExclamation, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   'Validando
   'If (CDbl(pnl_SumTot_Dsm.Caption) <> 0) Then
   '    If (CDbl(pnl_SumTot_Dsm.Caption) <> l_dbl_ImpPtm) Then
   '        SSTab1(1).Tab = 6
   '        MsgBox "El préstamo total no es igual al distribuido en la pestaña datos del desembolso", vbExclamation, modgen_g_str_NomPlt
   '        Exit Sub
   '    End If
   'End If
   
   Screen.MousePointer = 11
   g_str_Parame = "select DESCAB_NUMOPE, DESCAB_CODEST from CRE_DESPROCAB where DESCAB_NUMOPE = '" & moddat_g_str_NumOpe & "'   "
   g_str_Parame = g_str_Parame & " AND descab_fecreg = '" & moddat_g_str_FecRec & "'    "
   g_str_Parame = g_str_Parame & " AND descab_horreg = '" & moddat_g_str_FecHip & "'    "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   If (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   If (Trim(CStr(g_rst_Princi!DESCAB_CODEST)) <> Trim(moddat_g_str_CodIte)) Then
      MsgBox "Este registro ya ha cambiado de estado.", vbExclamation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   'Actualizando
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
   
      g_str_Parame = "usp_Actualiza_cre_desprocab("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_FecRec & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_FecHip & "', "
      
      If (moddat_g_int_CodIns = "2") Then
          g_str_Parame = g_str_Parame & "'2', " 'DESCAB_CODAREA
          g_str_Parame = g_str_Parame & "'5', " 'DESCAB_CODEST
          
          g_str_Parame = g_str_Parame & "'', " 'DESCAB_FESLNT
          g_str_Parame = g_str_Parame & "'', " 'DESCAB_CMNLEG
          g_str_Parame = g_str_Parame & "'', " 'DESCAB_FEENNT
          g_str_Parame = g_str_Parame & "'', " 'DESCAB_CMNLE2
          g_str_Parame = g_str_Parame & "'', " 'DESCAB_FERELG
          g_str_Parame = g_str_Parame & "'" & txt_Comentario.Text & "', "
          g_str_Parame = g_str_Parame & "'', " 'DESCAB_FERECE
          g_str_Parame = g_str_Parame & "'', " 'DESCAB_CMNOP2
      Else
          g_str_Parame = g_str_Parame & "'4', " 'DESCAB_CODAREA
          g_str_Parame = g_str_Parame & "'9', " 'DESCAB_CODEST
          g_str_Parame = g_str_Parame & "'', " 'DESCAB_FESLNT
          g_str_Parame = g_str_Parame & "'', " 'DESCAB_CMNLEG
          g_str_Parame = g_str_Parame & "'', " 'DESCAB_FEENNT
          g_str_Parame = g_str_Parame & "'', " 'DESCAB_CMNLE2
          g_str_Parame = g_str_Parame & "'', " 'DESCAB_FERELG
          g_str_Parame = g_str_Parame & "'', " 'DESCAB_CMNOPE
          If (Len(Trim(ipp_FecConstancia.Text)) = 0) Then
             g_str_Parame = g_str_Parame & "'', "
          Else
             g_str_Parame = g_str_Parame & "'" & Format(ipp_FecConstancia.Text, "yyyymmdd") & "', "
          End If
          g_str_Parame = g_str_Parame & "'" & txt_Comentario.Text & "', " 'DESCAB_CMNOP2
      End If
      g_str_Parame = g_str_Parame & "'', " 'DESCAB_CMNLE2
      '-----------------------------------
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      Screen.MousePointer = 0
   Loop
   
   r_bol_Estado = fs_guardar_ctaPromotor
   If (r_bol_Estado = False) Then
       Exit Sub
   End If
   
  'Enviando Correo Electrónico
   Call fs_Envia_Correo("RECHAZO")
   
   'Imprime liquidacion
   Screen.MousePointer = 0
   If (moddat_g_int_CntErr = 0) Then
       MsgBox "El proceso se grabó exitosamente.", vbInformation, modgen_g_str_NomPlt
       frm_RegDes_01.fs_Buscar_Creditos
       Unload Me
   End If
End Sub

Private Sub cmd_Aprobar_Click()
Dim r_bol_Estado As Boolean

   If Len(Trim(moddat_g_str_NumOpe)) = 0 Then
      MsgBox "Tiene que Haber un Nro. Operación.", vbExclamation, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   'Validando
   If (CDbl(pnl_SumTot_Dsm.Caption) <> CDbl(CStr(l_dbl_ImpPtm))) Then
       SSTab1(1).Tab = 7
       MsgBox "El préstamo total no es igual al distribuido en la pestaña datos del desembolso", vbExclamation, modgen_g_str_NomPlt
       Exit Sub
   End If

   Screen.MousePointer = 11
   g_str_Parame = "select DESCAB_NUMOPE, DESCAB_CODEST from CRE_DESPROCAB where DESCAB_NUMOPE = '" & moddat_g_str_NumOpe & "'   "
   g_str_Parame = g_str_Parame & " AND descab_fecreg = '" & moddat_g_str_FecRec & "'    "
   g_str_Parame = g_str_Parame & " AND descab_horreg = '" & moddat_g_str_FecHip & "'    "
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   If (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   If (Trim(CStr(g_rst_Princi!DESCAB_CODEST)) <> Trim(moddat_g_str_CodIte)) Then
      MsgBox "Este registro ya ha cambiado de estado.", vbExclamation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   If (moddat_g_str_CodIte = 6) Then '---Tesoreria Aprobada
       If (Trim(ipp_FecConstancia.Text) = Trim("")) Then
           MsgBox "Debe ingresar una fecha.", vbExclamation, modgen_g_str_NomPlt
           Screen.MousePointer = 0
           Call gs_SetFocus(ipp_FecConstancia)
           Exit Sub
       End If
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   'Actualizando
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      g_str_Parame = "usp_Actualiza_cre_desprocab("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_FecRec & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_FecHip & "', "
      If (moddat_g_int_CodIns = 2) Then 'operaciones
          g_str_Parame = g_str_Parame & "'3', "
          g_str_Parame = g_str_Parame & "'4', "
          g_str_Parame = g_str_Parame & "'', " 'fecha solicitud notaria
          g_str_Parame = g_str_Parame & "'', " 'comentario legal 1
          g_str_Parame = g_str_Parame & "'', " 'fecha entrega notaria
          g_str_Parame = g_str_Parame & "'', " 'comentario legal 2
          g_str_Parame = g_str_Parame & "'', " 'fecha de recepcion 2
          g_str_Parame = g_str_Parame & "'" & txt_Comentario.Text & "', "
          g_str_Parame = g_str_Parame & "'', "
          g_str_Parame = g_str_Parame & "'', "
      Else
          g_str_Parame = g_str_Parame & "'5', "
          g_str_Parame = g_str_Parame & "'8', "
          g_str_Parame = g_str_Parame & "'', " 'fecha solicitud notaria
          g_str_Parame = g_str_Parame & "'', " 'comentario legal 1
          g_str_Parame = g_str_Parame & "'', " 'fecha entrega notaria
          g_str_Parame = g_str_Parame & "'', " 'comentario legal 2
          g_str_Parame = g_str_Parame & "'', " 'fecha de recepcion 2
          g_str_Parame = g_str_Parame & "'', "
          g_str_Parame = g_str_Parame & "'" & Format(ipp_FecConstancia.Text, "yyyymmdd") & "', "
          g_str_Parame = g_str_Parame & "'" & txt_Comentario.Text & "', "
      End If
      
      g_str_Parame = g_str_Parame & "'', " 'DESCAB_CMNLE2
      '------------
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      Screen.MousePointer = 0
   Loop
   
   r_bol_Estado = fs_guardar_ctaPromotor
   If (r_bol_Estado = False) Then
       Exit Sub
   End If
   
   'Enviando Correo Electrónico
   Call fs_Envia_Correo("APROBACION")
   
   'Imprime liquidacion
   Screen.MousePointer = 0
   If (moddat_g_int_CntErr = 0) Then
       MsgBox "El proceso se grabó exitosamente.", vbInformation, modgen_g_str_NomPlt
       frm_RegDes_01.fs_Buscar_Creditos
       Unload Me
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub fs_Envia_Correo(p_Estado As String)
   modgen_g_str_Mail_Asunto = "PAGO PROMOTOR - AREA OPERACIONES - " & p_Estado & " (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
   
   modgen_g_str_Mail_Mensaj = ""
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & moddat_g_str_NumSol & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE OPERACION : " & moddat_g_str_NumOpe & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
   
   Call fs_Envia_Correo_Prom(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, "", "", False, True, False, False, True, True)
End Sub

Private Sub Form_Load()
Dim r_arr_Mtz()      As moddat_g_tpo_DatCom

   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   pnl_NumOpe.Caption = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
   pnl_Produc.Caption = Trim(moddat_g_str_NomPrd)
   pnl_EstadoActual.Caption = moddat_g_str_Situac
   pnl_NomCli.Caption = Trim(CStr(moddat_g_int_TipDoc)) & "-" & Trim(moddat_g_str_NumDoc) & " / " & Trim(moddat_g_str_NomCli)
   
   Call fs_Inicia
   Call moddat_gf_Cargar_AgrPrd
   
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
   Call modmip_gs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(0), 0)      'Buscar Información del Cliente
   Call modmip_gs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, grd_Listad(2), 1)      'Buscar Información del Cónyuge
   Call modmip_gs_DatInm(grd_Listad(1), True)
   
   Call fs_DatInm_Aux
   Call fs_DatLeg
   Call fs_DatDes
   Call fs_CalcMto
   Call modmip_gs_DatCre(grd_Listad(4), r_arr_Mtz)
   Call fs_Dat_Evaluacion
   Call fs_PryCta
   
   SSTab1(1).Tab = 7
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_PryCta()
Dim r_rst_Princi  As ADODB.Recordset
Dim r_str_Cadena  As String
    
    Call moddat_gs_Carga_LisIte_Combo(cmb_FrmDsm_Dsm, 1, "376")
    Call moddat_gs_Carga_LisIte_Combo(cmb_TipMto_Dsm, 1, "132")
    
    g_str_Parame = ""
    g_str_Parame = g_str_Parame & " SELECT DISTINCT A.Ctaban_Codbco, TRIM(B.PARDES_DESCRI) AS NOM_BANCO  "
    g_str_Parame = g_str_Parame & "   FROM PRY_CTABAN A  "
    g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = 513 AND B.PARDES_CODITE = A.Ctaban_Codbco  "
    g_str_Parame = g_str_Parame & "  WHERE A.CTABAN_CODPRY = '" & CStr(pnl_Prycto_Dsm.Tag) & "'"
    g_str_Parame = g_str_Parame & "    AND A.CTABAN_TIPMON = " & CStr(l_str_CodMod)
    
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
    End If
   
    ReDim l_arr_Bancos(0)
    cmb_EntFin_Dsm.Clear
    
    If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
       g_rst_Princi.MoveFirst
       Do While Not g_rst_Princi.EOF
          ReDim Preserve l_arr_Bancos(UBound(l_arr_Bancos) + 1)
          l_arr_Bancos(UBound(l_arr_Bancos)).Genera_Codigo = Trim$(g_rst_Princi!Ctaban_Codbco)
          l_arr_Bancos(UBound(l_arr_Bancos)).Genera_Nombre = Trim$(g_rst_Princi!NOM_BANCO & "")
          cmb_EntFin_Dsm.AddItem Trim$(g_rst_Princi!NOM_BANCO & "")
          g_rst_Princi.MoveNext
       Loop
    End If

    g_rst_Princi.Close
    Set g_rst_Princi = Nothing
    '-------------------------------------------------------
    g_str_Parame = ""
    g_str_Parame = g_str_Parame & "SELECT CT.* "
    g_str_Parame = g_str_Parame & "  FROM PRY_CTABAN CT "
    g_str_Parame = g_str_Parame & " WHERE CT.CTABAN_CODPRY = '" & CStr(pnl_Prycto_Dsm.Tag) & "'"
    g_str_Parame = g_str_Parame & "   AND CT.CTABAN_TIPMON = " & CStr(l_str_CodMod)
    'g_str_Parame = g_str_Parame & "   AND CT.CTABAN_SITUAC = 1 "
    
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
    End If
   
    ReDim l_arr_CtaBco(0)
    
    If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
       g_rst_Princi.MoveFirst
       Do While Not g_rst_Princi.EOF
          ReDim Preserve l_arr_CtaBco(UBound(l_arr_CtaBco) + 1)
          l_arr_CtaBco(UBound(l_arr_CtaBco)).Genera_Codigo = Trim$(g_rst_Princi!Ctaban_Codbco)
          l_arr_CtaBco(UBound(l_arr_CtaBco)).Genera_Nombre = Trim$(g_rst_Princi!CtaBan_NumCta & "")
          l_arr_CtaBco(UBound(l_arr_CtaBco)).Genera_FlgAso = Trim$(g_rst_Princi!CTABAN_SITUAC)
          l_arr_CtaBco(UBound(l_arr_CtaBco)).Genera_Refere = Trim$(g_rst_Princi!CTABAN_NUMCCI & "")
          l_arr_CtaBco(UBound(l_arr_CtaBco)).Genera_NomCli = Trim$(g_rst_Princi!CTABAN_ANOMDE & "")
          l_arr_CtaBco(UBound(l_arr_CtaBco)).Genera_ConHip = Trim$(g_rst_Princi!CTABAN_NOMCHQ & "")
          
          g_rst_Princi.MoveNext
       Loop
    End If
    
    Call cmd_Dsm_Cancel_Click
    Call fs_sumarDesemPrmt

    g_rst_Princi.Close
    Set g_rst_Princi = Nothing
    
    '******************
    'Detalle de cuentas
    '******************
    If (Trim(moddat_g_str_FecHip) <> "" And Trim(moddat_g_str_FecRec) <> "") Then
        g_str_Parame = ""
        g_str_Parame = g_str_Parame & " SELECT DT.DESDAT_NUMOPE, DT.DESDAT_FECREG, DT.DESDAT_HORREG, DT.DESDAT_CODBCO,  "
        g_str_Parame = g_str_Parame & "        (SELECT A.PARDES_DESCRI FROM MNT_PARDES A WHERE A.PARDES_CODGRP = 513 AND A.PARDES_CODITE = DT.DESDAT_CODBCO) AS BANCO,  "
        g_str_Parame = g_str_Parame & "        DT.DESDAT_NUMCTA, DT.DESDAT_FRMDES, DT.DESDAT_NUMITE, DT.DESDAT_TIPMTO,  "
        g_str_Parame = g_str_Parame & "        (SELECT A.PARDES_DESCRI FROM MNT_PARDES A WHERE A.PARDES_CODGRP = 376 AND A.PARDES_CODITE = DT.DESDAT_FRMDES) AS TIPODESEMBOLSO,  "
        g_str_Parame = g_str_Parame & "        (SELECT A.PARDES_DESCRI FROM MNT_PARDES A WHERE A.PARDES_CODGRP = 132 AND A.PARDES_CODITE = DT.DESDAT_TIPMTO) AS TIPOMONTO,  "
        g_str_Parame = g_str_Parame & "        DT.DESDAT_NUMDES, DT.DESDAT_FCHDES, DT.DESDAT_DESCRI, "
        g_str_Parame = g_str_Parame & "        DT.DESDAT_ANOMBR , DT.DESDAT_IMPORT  "
        g_str_Parame = g_str_Parame & "   FROM CRE_DESPRODAT DT "
        g_str_Parame = g_str_Parame & "  WHERE DT.DESDAT_NUMOPE = '" & moddat_g_str_NumOpe & "' "
        g_str_Parame = g_str_Parame & "    AND DT.DESDAT_FECREG = " & moddat_g_str_FecRec
        g_str_Parame = g_str_Parame & "    AND DT.DESDAT_HORREG = " & moddat_g_str_FecHip
   
        If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
           Exit Sub
        End If
               
        If Not (r_rst_Princi.EOF And r_rst_Princi.BOF) Then
           r_rst_Princi.MoveFirst
           Do While Not r_rst_Princi.EOF
              grd_Listad_Dsm.Rows = grd_Listad_Dsm.Rows + 1
              grd_Listad_Dsm.Row = grd_Listad_Dsm.Rows - 1
 
              grd_Listad_Dsm.Col = 0
              grd_Listad_Dsm.Text = r_rst_Princi!DESDAT_FRMDES
              
              grd_Listad_Dsm.Col = 1
              grd_Listad_Dsm.Text = IIf(IsNull(r_rst_Princi!TIPODESEMBOLSO) = True, "", Trim(r_rst_Princi!TIPODESEMBOLSO))
              
              grd_Listad_Dsm.Col = 2
              grd_Listad_Dsm.Text = IIf(IsNull(r_rst_Princi!DESDAT_TIPMTO) = True, "", Trim(r_rst_Princi!DESDAT_TIPMTO))
              
              grd_Listad_Dsm.Col = 3
              grd_Listad_Dsm.Text = IIf(IsNull(r_rst_Princi!TIPOMONTO) = True, "", Trim(r_rst_Princi!TIPOMONTO))
  
              grd_Listad_Dsm.Col = 4
              grd_Listad_Dsm.Text = gf_FormatoNumero(r_rst_Princi!DESDAT_IMPORT, 12, 2)
              
              grd_Listad_Dsm.Col = 5
              grd_Listad_Dsm.Text = Trim(r_rst_Princi!DESDAT_CODBCO & "")
            
              grd_Listad_Dsm.Col = 6
              grd_Listad_Dsm.Text = IIf(IsNull(r_rst_Princi!BANCO) = True, "", Trim(r_rst_Princi!BANCO))
              
              grd_Listad_Dsm.Col = 7
              grd_Listad_Dsm.Text = IIf(IsNull(r_rst_Princi!DESDAT_NUMCTA) = True, "", Trim(r_rst_Princi!DESDAT_NUMCTA))
              
              grd_Listad_Dsm.Col = 8
              grd_Listad_Dsm.Text = IIf(IsNull(r_rst_Princi!DESDAT_ANOMBR) = True, "", Trim(r_rst_Princi!DESDAT_ANOMBR))
                            
              grd_Listad_Dsm.Col = 9
              grd_Listad_Dsm.Text = IIf(IsNull(r_rst_Princi!DESDAT_NUMDES) = True, "", Trim(r_rst_Princi!DESDAT_NUMDES))
                    
              grd_Listad_Dsm.Col = 10
              grd_Listad_Dsm.Text = IIf(IsNull(r_rst_Princi!DESDAT_FCHDES) = True, "", _
                                        Right(r_rst_Princi!DESDAT_FCHDES, 2) & "/" & Mid(r_rst_Princi!DESDAT_FCHDES, 5, 2) & _
                                        "/" & Left(r_rst_Princi!DESDAT_FCHDES, 4))
              
              grd_Listad_Dsm.Col = 11
              grd_Listad_Dsm.Text = IIf(IsNull(r_rst_Princi!DESDAT_DESCRI) = True, "", Trim(r_rst_Princi!DESDAT_DESCRI))
                        
              grd_Listad_Dsm.Col = 12
              grd_Listad_Dsm.Text = "S"
              
              grd_Listad_Dsm.Col = 13
              grd_Listad_Dsm.Text = Trim(r_rst_Princi!DESDAT_NUMITE)
              
              r_rst_Princi.MoveNext
              DoEvents
           Loop
           Call gs_UbiIniGrid(grd_Listad_Dsm)
        End If
    End If
    Call fs_sumarDesemPrmt
    
    If (Trim(pnl_Prycto_Dsm.Tag) = "") Then
        cmd_Dsm_Nuevo.Enabled = False
        cmd_Dsm_Borrar.Enabled = False
        cmd_Dsm_Editar.Enabled = False
    End If
    If (CInt(moddat_g_int_CodIns) = CInt("000005")) Then
        'LEGAL 2DA PARTE
        If (CInt(moddat_g_str_CodIte) = CInt("000010") Or CInt(moddat_g_str_CodIte) = CInt("000009")) Then
            cmd_Dsm_Nuevo.Enabled = False
            cmd_Dsm_Borrar.Enabled = False
            cmd_Dsm_Editar.Enabled = False
        End If
    End If
    If (cmd_Grabar.Enabled = False) Then
        cmd_Dsm_Nuevo.Enabled = False
        cmd_Dsm_Borrar.Enabled = False
        cmd_Dsm_Editar.Enabled = False
    End If
    If (moddat_g_int_CodIns = CInt("000002") Or moddat_g_int_CodIns = CInt("000001")) Then
        'OPERACIONES  -  LEGAL
        lbl_NumDsm_Dsm.Visible = False
        txt_NroDsm_Dsm.Visible = False
        lbl_FchDsm_Dsm.Visible = False
        ipp_FecDsm_Dsm.Visible = False
    ElseIf (moddat_g_int_CodIns = CInt("000003") Or moddat_g_int_CodIns = CInt("000004") Or moddat_g_int_CodIns = CInt("000005")) Then
        'TESORERIA  -  OPERACIONES 2DA PARTE   -   LEGAL 2DA PARTE
        lbl_NumDsm_Dsm.Visible = True
        txt_NroDsm_Dsm.Visible = True
        lbl_FchDsm_Dsm.Visible = True
        ipp_FecDsm_Dsm.Visible = True
    End If
    If (moddat_g_int_CodIns = CInt("000004") Or moddat_g_int_CodIns = CInt("000005")) Then ''LEGAL 2DA PARTE
        'OPERACIONES 2DA PARTE    -   LEGAL 2DA PARTE
        cmd_Dsm_Nuevo.Enabled = False
        cmd_Dsm_Borrar.Enabled = False
        cmd_Dsm_Editar.Enabled = False
    End If
End Sub

Private Sub gs_Carga_EntFin()
   cmb_EntFin_Dsm.Clear
   ReDim l_arr_Bancos(0)
      
   ReDim Preserve l_arr_Bancos(UBound(l_arr_Bancos) + 1)
   l_arr_Bancos(UBound(l_arr_Bancos)).Genera_Codigo = Trim$(l_str_CodBan)
   l_arr_Bancos(UBound(l_arr_Bancos)).Genera_Nombre = Trim$(l_str_PryBan)
   l_arr_Bancos(UBound(l_arr_Bancos)).Genera_TipVal = 0
   l_arr_Bancos(UBound(l_arr_Bancos)).Genera_Cantid = 0
End Sub

Public Sub fs_Dat_Evaluacion()
Dim r_str_frmDesem As String
Dim r_str_feslnt As String
Dim r_str_nrodes As String
Dim r_str_ferece As String
Dim r_str_feennt As String
Dim r_str_cmnleg As String
Dim r_str_cmnope As String
Dim r_str_cmntes As String
Dim r_str_cmnop2 As String
Dim r_str_cmnle2 As String

Dim r_str_Legal As String
Dim r_str_Oper As String
Dim r_str_Teso As String
Dim r_str_Oper_2 As String
Dim r_str_Legal_2 As String

Dim r_str_LegFec As String
Dim r_str_OpeFec As String
Dim r_str_TesFec As String
Dim r_str_OpeFec_2 As String
Dim r_str_LegFec_2 As String

   g_str_Parame = "  SELECT to_number(PARDES_CODITE) PARDES_CODITE, trim(PARDES_DESCRI) as Instancia,  "
   g_str_Parame = g_str_Parame & "  (select det.desdet_fecfin from cre_desprodet det  "
   g_str_Parame = g_str_Parame & "  where det.desdet_numope = '" & moddat_g_str_NumOpe & "'  "
   g_str_Parame = g_str_Parame & "  and det.desdet_fecreg = '" & moddat_g_str_FecRec & "'  "
   g_str_Parame = g_str_Parame & "  and det.desdet_horreg = '" & moddat_g_str_FecHip & "'  "
   g_str_Parame = g_str_Parame & "  and par.PARDES_CODITE = det.desdet_codarea) as fechaEnvio  "
   
   g_str_Parame = g_str_Parame & "  FROM MNT_PARDES par WHERE par.PARDES_CODGRP = '374'  "
   g_str_Parame = g_str_Parame & "  and par.PARDES_CODITE <> '000000' AND par.PARDES_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "  ORDER BY PARDES_CODITE ASC  "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
      Select Case g_rst_Princi!PARDES_CODITE
             Case 1
                   r_str_Legal = g_rst_Princi!INSTANCIA
                   If (IsNull(g_rst_Princi!fechaEnvio) = False) Then
                       r_str_LegFec = g_rst_Princi!fechaEnvio
                   End If
             Case 2
                   r_str_Oper = g_rst_Princi!INSTANCIA
                   If (IsNull(g_rst_Princi!fechaEnvio) = False) Then
                       r_str_OpeFec = g_rst_Princi!fechaEnvio
                   End If
             Case 3
                   r_str_Teso = g_rst_Princi!INSTANCIA
                   If (IsNull(g_rst_Princi!fechaEnvio) = False) Then
                       r_str_TesFec = g_rst_Princi!fechaEnvio
                   End If
             Case 4
                   r_str_Oper_2 = g_rst_Princi!INSTANCIA
                   If (IsNull(g_rst_Princi!fechaEnvio) = False) Then
                       r_str_OpeFec_2 = g_rst_Princi!fechaEnvio
                   End If
             Case 5
                   r_str_Legal_2 = g_rst_Princi!INSTANCIA
                   If (IsNull(g_rst_Princi!fechaEnvio) = False) Then
                       r_str_LegFec_2 = g_rst_Princi!fechaEnvio
                   End If
      End Select
      g_rst_Princi.MoveNext
      Loop
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   '--------------------------------------------------------------------------------------------------------------
   Call gs_LimpiaGrid(grd_Listad(7))
 
   g_str_Parame = "" '-----trae la Cabecera----
   g_str_Parame = g_str_Parame & "SELECT descab_numope, descab_codarea, descab_codest,descab_fecreg, descab_horreg, "
   g_str_Parame = g_str_Parame & "       descab_feslnt, descab_ferece, descab_feennt, descab_cmnleg, descab_cmnope, "
   g_str_Parame = g_str_Parame & "       descab_cmntes , descab_cmnop2, descab_cmnle2, descab_FERELG "
   g_str_Parame = g_str_Parame & "  FROM cre_desprocab cab "
   g_str_Parame = g_str_Parame & " WHERE cab.descab_numope = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND cab.DESCAB_FECREG = '" & moddat_g_str_FecRec & "' "
   g_str_Parame = g_str_Parame & "   AND cab.DESCAB_HORREG = '" & moddat_g_str_FecHip & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      If (moddat_g_int_CodIns = 2) Then
          txt_Comentario.Text = IIf(IsNull(g_rst_Princi!descab_cmnope) = True, "", g_rst_Princi!descab_cmnope)
      Else '4
          txt_Comentario.Text = IIf(IsNull(g_rst_Princi!descab_cmnop2) = True, "", g_rst_Princi!descab_cmnop2)
          If (IsNull(g_rst_Princi!descab_ferece) = False) Then
              ipp_FecConstancia.Text = gf_FormatoFecha(g_rst_Princi!descab_ferece)
          End If
      End If
       
      Do While Not g_rst_Princi.EOF
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
             grd_Listad(7).Text = "Instancia"
             grd_Listad(7).Col = 1
             grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
             grd_Listad(7).Text = r_str_Legal
                  
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).Text = "Fecha de Solicitud Notaria"
             If (IsNull(g_rst_Princi!descab_feslnt) = False) Then
                 grd_Listad(7).Col = 1
                 grd_Listad(7).Text = gf_FormatoFecha(g_rst_Princi!descab_feslnt)
             End If
        
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).Text = "Comentario Legal"
             grd_Listad(7).Col = 1
             grd_Listad(7).Text = IIf(IsNull(g_rst_Princi!descab_cmnleg), "", g_rst_Princi!descab_cmnleg)
             
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).Text = "Fecha Envío"
             grd_Listad(7).Col = 1
             
             If (Len(Trim((r_str_LegFec))) <> 0) Then
                 grd_Listad(7).Text = gf_FormatoFecha(r_str_LegFec)
             End If
             
'-----------------------------------------------------------------------------------------------------
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
             grd_Listad(7).Text = "Instancia"
             grd_Listad(7).Col = 1
             grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
             grd_Listad(7).Text = r_str_Oper
         
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).Text = "Comentario Operaciones"
             grd_Listad(7).Col = 1
             grd_Listad(7).Text = IIf(IsNull(g_rst_Princi!descab_cmnope), "", g_rst_Princi!descab_cmnope)
             
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).Text = "Fecha Envío"
             grd_Listad(7).Col = 1
             If (Len(Trim((r_str_OpeFec))) <> 0) Then
                 grd_Listad(7).Text = gf_FormatoFecha(r_str_OpeFec)
             End If
'-----------------------------------------------------------------------------------------------------
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
             grd_Listad(7).Text = "Instancia"
             grd_Listad(7).Col = 1
             grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
             grd_Listad(7).Text = r_str_Teso
                      
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).Text = "Comentario Tesoreria"
             grd_Listad(7).Col = 1
             grd_Listad(7).Text = IIf(IsNull(g_rst_Princi!descab_cmntes), "", g_rst_Princi!descab_cmntes)
             
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).Text = "Fecha Envío"
             grd_Listad(7).Col = 1
             If (Len(Trim((r_str_TesFec))) <> 0) Then
                 grd_Listad(7).Text = gf_FormatoFecha(r_str_TesFec)
             End If
'-----------------------------------------------------------------------------------------------------
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
             grd_Listad(7).Text = "Instancia"
             grd_Listad(7).Col = 1
             grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
             grd_Listad(7).Text = r_str_Oper_2
                  
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).Text = "Fecha Recepcion Const. Desembolso"
             If (IsNull(g_rst_Princi!descab_ferece) = False) Then
                 grd_Listad(7).Col = 1
                 grd_Listad(7).Text = gf_FormatoFecha(g_rst_Princi!descab_ferece)
             End If
         
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).Text = "Comentario Operaciones 2da Parte"
             grd_Listad(7).Col = 1
             grd_Listad(7).Text = IIf(IsNull(g_rst_Princi!descab_cmnop2), "", g_rst_Princi!descab_cmnop2)
             
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).Text = "Fecha Envío"
             grd_Listad(7).Col = 1
             If (Len(Trim((r_str_OpeFec_2))) <> 0) Then
                 grd_Listad(7).Text = gf_FormatoFecha(r_str_OpeFec_2)
             End If
'-----------------------------------------------------------------------------------------------------
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
             grd_Listad(7).Text = "Instancia"
             grd_Listad(7).Col = 1
             grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
             grd_Listad(7).Text = r_str_Legal_2
         
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).Text = "Fecha de Recepción"
             If (IsNull(g_rst_Princi!descab_FERELG) = False) Then
                 grd_Listad(7).Col = 1
                 grd_Listad(7).Text = gf_FormatoFecha(g_rst_Princi!descab_FERELG)
             End If
             
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).Text = "Fecha de Entrega Notaria"
             If (IsNull(g_rst_Princi!descab_feennt) = False) Then
                 grd_Listad(7).Col = 1
                 grd_Listad(7).Text = gf_FormatoFecha(g_rst_Princi!descab_feennt)
             End If
                                                
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).Text = "Comentario Legal 2da Parte"
             grd_Listad(7).Col = 1
             grd_Listad(7).Text = IIf(IsNull(g_rst_Princi!descab_cmnle2), "", g_rst_Princi!descab_cmnle2)
             
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).Text = "Fecha Termino"
             grd_Listad(7).Col = 1
             If (Len(Trim((r_str_LegFec_2))) <> 0) Then
                 grd_Listad(7).Text = gf_FormatoFecha(r_str_LegFec_2)
             End If
             
         g_rst_Princi.MoveNext
      Loop
      
      Call gs_UbiIniGrid(grd_Listad(7))
   End If
      
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatInm_Aux()
Dim r_str_Cadena As String
   l_str_PryBan = ""
   l_str_CodBan = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT SL.*, EL.EVALEG_FEENIN, PY.DATGEN_VENTDO, PY.DATGEN_VENNDO, PY.DATGEN_CONTDO, PY.DATGEN_CONNDO, "
   g_str_Parame = g_str_Parame & "       (SELECT TRIM(A.PARDES_DESCRI) FROM MNT_PARDES A WHERE A.PARDES_CODGRP = 513 AND A.PARDES_CODITE = SL.SOLINM_PRYBCO) AS BANCO, "
   g_str_Parame = g_str_Parame & "       (SELECT TRIM(B.DATGEN_TITULO) FROM PRY_DATGEN B WHERE B.DATGEN_CODIGO = SL.SOLINM_PRYCOD) AS PROYECTO "
   g_str_Parame = g_str_Parame & "  FROM CRE_SOLINM SL "
   g_str_Parame = g_str_Parame & "   LEFT JOIN PRY_DATGEN PY ON PY.DATGEN_CODIGO = SL.SOLINM_PRYCOD "
   g_str_Parame = g_str_Parame & "   LEFT JOIN TRA_EVALEG EL ON EL.EVALEG_NUMSOL = SL.SOLINM_NUMSOL "
   g_str_Parame = g_str_Parame & " WHERE SL.SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      pnl_Prycto_Dsm.Caption = IIf(IsNull(g_rst_Princi!PROYECTO) = True, "", g_rst_Princi!PROYECTO)
      pnl_Prycto_Dsm.Tag = IIf(IsNull(g_rst_Princi!SOLINM_PRYCOD) = True, "", g_rst_Princi!SOLINM_PRYCOD)
      If Not IsNull(g_rst_Princi!BANCO) And Trim(CStr(g_rst_Princi!SOLINM_PRYBCO & "")) <> "888888" Then
         l_str_PryBan = Trim(CStr(g_rst_Princi!BANCO))
         l_str_CodBan = Trim(CStr(g_rst_Princi!SOLINM_PRYBCO))
      End If
      
      If g_rst_Princi!SOLINM_TABPRY = 2 Then
         'CREDITOS ANTIGUOS
         If (Len(Trim(g_rst_Princi!SOLINM_TIPDOC_PRO)) > 0 And Len(Trim(g_rst_Princi!SOLINM_NUMDOC_PRO)) > 0 And Len(Trim(g_rst_Princi!SOLINM_RAZSOC_PRO)) > 0) Then
             l_str_Prmtor = CStr(g_rst_Princi!SOLINM_TIPDOC_PRO) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO) & _
                            " / " & moddat_gf_Consulta_RazSoc(g_rst_Princi!SOLINM_TIPDOC_PRO, g_rst_Princi!SOLINM_NUMDOC_PRO)
         Else
             If (Len(Trim(g_rst_Princi!SOLINM_PRYCOD)) > 0) Then
                 r_str_Cadena = moddat_gf_Consulta_RazSoc(g_rst_Princi!DATGEN_VENTDO, g_rst_Princi!DATGEN_VENNDO)
                 If (Len(Trim(r_str_Cadena)) > 0) Then
                     l_str_Prmtor = CStr(g_rst_Princi!DATGEN_VENTDO) & "-" & Trim(g_rst_Princi!DATGEN_VENNDO) & _
                                    " / " & moddat_gf_Consulta_RazSoc(g_rst_Princi!SOLINM_TIPDOC_PRO, g_rst_Princi!SOLINM_NUMDOC_PRO)
                 End If
             End If
         End If
      Else
      'CREDITOS NUEVOS
         If CInt(g_rst_Princi!SOLINM_CODMOD) = 1 Then
            If (Len(Trim(g_rst_Princi!SOLINM_TIPDOC_PRO)) > 0 And Len(Trim(g_rst_Princi!SOLINM_NUMDOC_PRO)) > 0 And Len(Trim(g_rst_Princi!SOLINM_RAZSOC_PRO)) > 0) Then
                l_str_Prmtor = CStr(g_rst_Princi!SOLINM_TIPDOC_PRO) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO) & _
                               " / " & Trim(g_rst_Princi!SOLINM_RAZSOC_PRO & "")
            Else
                If (Len(Trim(g_rst_Princi!SOLINM_PRYCOD)) > 0) Then
                    r_str_Cadena = moddat_gf_Consulta_RazSoc(g_rst_Princi!DATGEN_VENTDO, g_rst_Princi!DATGEN_VENNDO)
                    If (Len(Trim(r_str_Cadena)) > 0) Then
                        l_str_Prmtor = CStr(g_rst_Princi!DATGEN_VENTDO) & "-" & Trim(g_rst_Princi!DATGEN_VENNDO) & _
                               " / " & r_str_Cadena
                    End If
                End If
            End If
         Else
            '*********BIEN FUTURO**********
            r_str_Cadena = moddat_gf_Consulta_RazSoc(g_rst_Princi!SOLINM_TIPDOC_PRO, g_rst_Princi!SOLINM_NUMDOC_PRO)
            If (Len(Trim(r_str_Cadena)) > 0) Then
                l_str_Prmtor = CStr(g_rst_Princi!SOLINM_TIPDOC_PRO) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO) & _
                               " / " & r_str_Cadena
            End If
         End If
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatLeg()
   Call gs_LimpiaGrid(grd_Listad(6))
   l_int_MonCvt = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM TRA_EVALEG "
   g_str_Parame = g_str_Parame & " WHERE EVALEG_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      txt_InfLeg.Text = Trim(g_rst_Princi!EVALEG_INFLG1 & "") & Trim(g_rst_Princi!EVALEG_INFLG2 & "") & Trim(g_rst_Princi!EVALEG_INFLG3 & "") & Trim(g_rst_Princi!EVALEG_INFLG4 & "")
      txt_ComCre.Text = "Fecha de Comité de Créditos: " & gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECCOM)) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Trim(g_rst_Princi!EVALEG_OBSCOM & "")
      
      If g_rst_Princi!EVALEG_FECCVT > 0 Then
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Fecha Firma Contrato Compra Venta"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECCVT))
         
         If Not IsNull(g_rst_Princi!EVALEG_TCASBS) Then
            If g_rst_Princi!EVALEG_TCASBS > 0 Then
               grd_Listad(6).Rows = grd_Listad(6).Rows + 1
               grd_Listad(6).Row = grd_Listad(6).Rows - 1
               grd_Listad(6).Col = 0
               grd_Listad(6).Text = "Tipo de Cambio SBS"
               
               grd_Listad(6).Col = 1
               grd_Listad(6).Text = Format(g_rst_Princi!EVALEG_TCASBS, "###,##0.0000")
            End If
         End If
      
         If g_rst_Princi!EVALEG_TCACVT > 0 Then
            grd_Listad(6).Rows = grd_Listad(6).Rows + 1
            grd_Listad(6).Row = grd_Listad(6).Rows - 1
            grd_Listad(6).Col = 0
            grd_Listad(6).Text = "Tipo de Cambio aplicado"
            
            grd_Listad(6).Col = 1
            grd_Listad(6).Text = Format(g_rst_Princi!EVALEG_TCACVT, "###,##0.0000")
         End If
      End If
      
      If Not IsNull(g_rst_Princi!EVALEG_MONCVT) Then
         grd_Listad(6).Rows = grd_Listad(6).Rows + 2
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Moneda Compra-Venta"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("204", g_rst_Princi!EVALEG_MONCVT)
         
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Valor Compra-Venta"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", g_rst_Princi!EVALEG_MONCVT) & " " & gf_FormatoNumero(g_rst_Princi!EVALEG_COMVTA, 12, 2)
      
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Aporte Propio"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", g_rst_Princi!EVALEG_MONCVT) & " " & gf_FormatoNumero(g_rst_Princi!EVALEG_APOPRO, 12, 2)
      
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Monto Préstamo"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", g_rst_Princi!EVALEG_MONCVT) & " " & gf_FormatoNumero(g_rst_Princi!EVALEG_MTOPRE, 12, 2)
      End If
      
      If grd_Listad(6).Rows = 0 Then
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      Else
         grd_Listad(6).Rows = grd_Listad(6).Rows + 2
      End If
      
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Fecha Firma Contrato (Crédito)"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FIRCON))
   
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Notaria"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("509", g_rst_Princi!EVALEG_CODNOT & "")
   
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Representante Legal 1"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("512", g_rst_Princi!EVALEG_REPLG1 & "")
   
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Representante Legal 2"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("512", g_rst_Princi!EVALEG_REPLG2 & "")
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Monto Hipoteca "
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellFontName = "Lucida Console"
      grd_Listad(6).CellFontSize = 8
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", g_rst_Princi!EVALEG_MONHIP) & " " & gf_FormatoNumero(g_rst_Princi!EVALEG_MTOHIP, 12, 2)
      
      If g_rst_Princi!EVALEG_FECBLQ_INM > 0 Then
         grd_Listad(6).Rows = grd_Listad(6).Rows + 2
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Bloqueo Registral Inscrito"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).Text = "SI"
      
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Sede Registral"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("511", CStr(g_rst_Princi!EVALEG_SEDREG & ""))
         
         grd_Listad(6).Rows = grd_Listad(6).Rows + 2
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Fecha Bloqueo (Inmueble)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECBLQ_INM))
                  
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Doc. Registral (Inmueble)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("026", g_rst_Princi!EVALEG_TIPDOC_INM)
                  
         Select Case g_rst_Princi!EVALEG_TIPDOC_INM
            Case 1: grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_INM & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_INM & "")
            Case 2: grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_INM & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_INM & "")
            Case 3: grd_Listad(6).Text = grd_Listad(6).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_INM & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_INM & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_INM & "") & ")"
         End Select
         
         If g_rst_Princi!EVALEG_FLGEST_ES1 = 1 Then
            grd_Listad(6).Rows = grd_Listad(6).Rows + 2
            grd_Listad(6).Row = grd_Listad(6).Rows - 1
            grd_Listad(6).Col = 0
            grd_Listad(6).Text = "Fecha Bloqueo (Estac. 1)"
            
            grd_Listad(6).Col = 1
            grd_Listad(6).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECBLQ_ES1))
                       
            grd_Listad(6).Rows = grd_Listad(6).Rows + 1
            grd_Listad(6).Row = grd_Listad(6).Rows - 1
            grd_Listad(6).Col = 0
            grd_Listad(6).Text = "Doc. Registral (Estac. 1)"
            
            grd_Listad(6).Col = 1
            grd_Listad(6).Text = moddat_gf_Consulta_ParDes("026", g_rst_Princi!EVALEG_TIPDOC_ES1)
            
            Select Case g_rst_Princi!EVALEG_TIPDOC_ES1
               Case 1: grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_ES1 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_ES1 & "")
               Case 2: grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_ES1 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_ES1 & "")
               Case 3: grd_Listad(6).Text = grd_Listad(6).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_ES1 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_ES1 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_ES1 & "") & ")"
            End Select
         End If
         
         If g_rst_Princi!EVALEG_FLGEST_ES2 = 1 Then
            grd_Listad(6).Rows = grd_Listad(6).Rows + 2
            grd_Listad(6).Row = grd_Listad(6).Rows - 1
            grd_Listad(6).Col = 0
            grd_Listad(6).Text = "Fecha Bloqueo (Estac. 2)"
            
            grd_Listad(6).Col = 1
            grd_Listad(6).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECBLQ_ES2))
                        
            grd_Listad(6).Rows = grd_Listad(6).Rows + 1
            grd_Listad(6).Row = grd_Listad(6).Rows - 1
            grd_Listad(6).Col = 0
            grd_Listad(6).Text = "Doc. Registral (Estac. 2)"
            
            grd_Listad(6).Col = 1
            grd_Listad(6).Text = moddat_gf_Consulta_ParDes("026", g_rst_Princi!EVALEG_TIPDOC_ES2)
            
            Select Case g_rst_Princi!EVALEG_TIPDOC_ES2
               Case 1: grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_ES2 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_ES2 & "")
               Case 2: grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_ES2 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_ES2 & "")
               Case 3: grd_Listad(6).Text = grd_Listad(6).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_ES2 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_ES2 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_ES2 & "") & ")"
            End Select
         End If
         
         If g_rst_Princi!EVALEG_FLGEST_DEP = 1 Then
            grd_Listad(6).Rows = grd_Listad(6).Rows + 2
            grd_Listad(6).Row = grd_Listad(6).Rows - 1
            grd_Listad(6).Col = 0
            grd_Listad(6).Text = "Fecha Bloqueo (Depósito)"
            
            grd_Listad(6).Col = 1
            grd_Listad(6).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECBLQ_DEP))
                        
            grd_Listad(6).Rows = grd_Listad(6).Rows + 1
            grd_Listad(6).Row = grd_Listad(6).Rows - 1
            grd_Listad(6).Col = 0
            grd_Listad(6).Text = "Doc. Registral (Depósito)"
            
            grd_Listad(6).Col = 1
            grd_Listad(6).Text = moddat_gf_Consulta_ParDes("026", g_rst_Princi!EVALEG_TIPDOC_DEP)
                        
            Select Case g_rst_Princi!EVALEG_TIPDOC_DEP
               Case 1: grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_DEP & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_DEP & "")
               Case 2: grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_DEP & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_DEP & "")
               Case 3: grd_Listad(6).Text = grd_Listad(6).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_DEP & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_DEP & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_DEP & "") & ")"
            End Select
         End If
      End If
      
      If Not IsNull(g_rst_Princi!EVALEG_MONCVT) Then
         l_int_MonCvt = g_rst_Princi!EVALEG_MONCVT
      End If
      
      Call gs_UbiIniGrid(grd_Listad(6))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatDes()
   Call gs_LimpiaGrid(grd_Listad(3))
   txt_ObsDes.Text = ""
   l_int_FlgCVt = 0

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPDES "
   g_str_Parame = g_str_Parame & " WHERE HIPDES_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_Listad(3).Rows = grd_Listad(3).Rows + 1
      grd_Listad(3).Row = grd_Listad(3).Rows - 1
      grd_Listad(3).Col = 0
      grd_Listad(3).Text = "Fecha de Desembolso"
      
      grd_Listad(3).Col = 1
      grd_Listad(3).Text = gf_FormatoFecha(g_rst_Princi!HIPDES_FECDES)
      
      grd_Listad(3).Rows = grd_Listad(3).Rows + 1
      grd_Listad(3).Row = grd_Listad(3).Rows - 1
      grd_Listad(3).Col = 0
      grd_Listad(3).Text = "Tipo de Desembolso"
      
      grd_Listad(3).Col = 1
      grd_Listad(3).Text = "CONTRA " & moddat_gf_Consulta_ParDes("241", g_rst_Princi!HIPDES_TIPGAR)
      
      If g_rst_Princi!HIPDES_TIPGAR = 2 Or g_rst_Princi!HIPDES_TIPGAR = 4 Or g_rst_Princi!HIPDES_TIPGAR = 5 Or g_rst_Princi!HIPDES_TIPGAR = 3 Then
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).Text = "Forma de Desembolso"
         
         grd_Listad(3).Col = 1
         grd_Listad(3).Text = moddat_gf_Consulta_ParDes("226", g_rst_Princi!HIPDES_TIPDES)
      End If
      
      If g_rst_Princi!HIPDES_TIPDES = 1 Then
         If Len(Trim(g_rst_Princi!HIPDES_CHECGO & "")) > 0 Then
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
            grd_Listad(3).Text = moddat_gf_Consulta_ParDes("516", g_rst_Princi!HIPDES_BANCGO & "") & " (" & Trim(g_rst_Princi!HIPDES_CTACGO & "") & ")"
         Else
            grd_Listad(3).Rows = grd_Listad(3).Rows + 1
            grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0
            grd_Listad(3).Text = "Nro. de Cheque"
            
            grd_Listad(3).Col = 1
            grd_Listad(3).Text = "CHEQUE NO EMITIDO"
            
            l_int_ChqReg = 1
         End If
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
         If Len(Trim(g_rst_Princi!HIPDES_NUMFIA & "")) > 0 Then
            grd_Listad(3).Rows = grd_Listad(3).Rows + 2
            grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0
            grd_Listad(3).Text = "Nro. Carta Fianza"
            
            grd_Listad(3).Col = 1
            grd_Listad(3).Text = Trim(g_rst_Princi!HIPDES_NUMFIA & "")
            
            grd_Listad(3).Rows = grd_Listad(3).Rows + 1
            grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0
            grd_Listad(3).Text = "Banco Emisor "
            
            grd_Listad(3).Col = 1
            grd_Listad(3).Text = moddat_gf_Consulta_ParDes("505", g_rst_Princi!HIPDES_BANFIA)
         
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
         Else
            grd_Listad(3).Rows = grd_Listad(3).Rows + 2
            grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0
            grd_Listad(3).Text = "Nro. Carta Fianza"
            
            grd_Listad(3).Col = 1
            grd_Listad(3).Text = "CARTA FIANZA NO RECIBIDA"
            
            l_int_FiaReg = 1
         End If
      End If
      
      If g_rst_Princi!HIPDES_TIPGAR = 5 Then
         If Len(Trim(g_rst_Princi!HIPDES_DOCGAR & "")) > 0 Then
            grd_Listad(3).Rows = grd_Listad(3).Rows + 2
            grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0
            grd_Listad(3).Text = "Nro. Certificado de Participación"
            
            grd_Listad(3).Col = 1
            grd_Listad(3).Text = Trim(g_rst_Princi!HIPDES_DOCGAR & "")
            
            grd_Listad(3).Rows = grd_Listad(3).Rows + 1
            grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0
            grd_Listad(3).Text = "Banco Emisor "
            
            grd_Listad(3).Col = 1
            grd_Listad(3).Text = moddat_gf_Consulta_ParDes("505", g_rst_Princi!HIPDES_BCOGAR)
            
            grd_Listad(3).Rows = grd_Listad(3).Rows + 1
            grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0
            grd_Listad(3).Text = "Importe Certificado"
            
            grd_Listad(3).Col = 1
            grd_Listad(3).CellFontName = "Lucida Console"
            grd_Listad(3).CellFontSize = 8
            grd_Listad(3).Text = moddat_gf_Consulta_ParDes("229", g_rst_Princi!HIPDES_MONGAR) & " " & gf_FormatoNumero(g_rst_Princi!HIPDES_MTOGAR, 12, 2)
         Else
            grd_Listad(3).Rows = grd_Listad(3).Rows + 2
            grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0
            grd_Listad(3).Text = "Nro. Certificado de Participación"
            
            grd_Listad(3).Col = 1
            grd_Listad(3).Text = "CERTIFICADO NO RECIBIDO"
            
            l_int_CerReg = 1
         End If
      End If
            
      Call gs_UbiIniGrid(grd_Listad(3))
      txt_ObsDes.Text = Trim(g_rst_Princi!HIPDES_OBSERV & "")
      
      If Not IsNull(g_rst_Princi!HIPDES_MONCVT) Then
         l_int_FlgCVt = 1
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_CalcMto()
   'Buscando Información del Crédito
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT HIPMAE_MTOPRE,SOLMAE_FMVBBP, SOLMAE_PBPMTO, SOLMAE_AFPMTO, SOLMAE_BMSMTO, HIPMAE_PRYMCS,  "
   g_str_Parame = g_str_Parame & "       HIPMAE_CVTSOL, HIPMAE_APOSOL, HIPMAE_CVTDOL, HIPMAE_APODOL, HIPMAE_FECESC, HIPMAE_PLAANO, "
   g_str_Parame = g_str_Parame & "       HIPMAE_TASINT, HIPMAE_NUMCUO, HIPMAE_PERGRA, HIPMAE_SEGPRE, HIPMAE_TIPSEG, HIPMAE_CONHIP, SOLMAE_MTOGCI "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE A"
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_SOLMAE B ON SOLMAE_NUMERO = HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND (HIPMAE_SITUAC = 2 OR HIPMAE_SITUAC = 3 OR HIPMAE_SITUAC = 6 OR HIPMAE_SITUAC = 7 OR HIPMAE_SITUAC = 9)"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If

   g_rst_Princi.MoveFirst
   
   'Datos_Promotor
   pnl_Moneda_Dsm.Caption = moddat_gf_Consulta_ParDes("204", CStr(moddat_g_int_TipMon))
   pnl_Moneda_Dsm.Tag = moddat_g_int_TipMon
   l_str_CodMod = moddat_g_int_TipMon
   l_str_Moneda = pnl_Moneda_Dsm.Caption
   lbl_Bono_Dsm.Caption = ".."
   l_dbl_ImpPtm = CDbl(g_rst_Princi!HIPMAE_MTOPRE)
   
   If moddat_g_int_TipMon = 1 Then
      If moddat_g_str_CodPrd = "024" Then
         If g_rst_Princi!HIPMAE_PRYMCS = 1 Then
            'VINCULADO
            l_dbl_ImpPtm = CDbl(g_rst_Princi!HIPMAE_MTOPRE) + CDbl(g_rst_Princi!SOLMAE_FMVBBP) + CDbl(g_rst_Princi!SOLMAE_PBPMTO) + CDbl(g_rst_Princi!SOLMAE_AFPMTO) + CDbl(g_rst_Princi!SOLMAE_BMSMTO)
            lbl_Bono_Dsm.Caption = "(INCLUYE BONOS " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & Format(g_rst_Princi!SOLMAE_FMVBBP + g_rst_Princi!SOLMAE_PBPMTO, "##,###,##0.00") & ")"
         Else
            'NO VINCULADO
            l_dbl_ImpPtm = CDbl(g_rst_Princi!HIPMAE_MTOPRE) + CDbl(g_rst_Princi!SOLMAE_BMSMTO) + CDbl(g_rst_Princi!SOLMAE_AFPMTO)
         End If
      ElseIf InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 And moddat_g_str_CodPrd <> "019" Then
         l_dbl_ImpPtm = CDbl(g_rst_Princi!HIPMAE_MTOPRE) + CDbl(g_rst_Princi!SOLMAE_FMVBBP) + CDbl(g_rst_Princi!SOLMAE_AFPMTO) + CDbl(g_rst_Princi!SOLMAE_PBPMTO) + CDbl(g_rst_Princi!SOLMAE_BMSMTO)
         lbl_Bono_Dsm.Caption = "(INCLUYE BONOS " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & Format(g_rst_Princi!SOLMAE_FMVBBP + g_rst_Princi!SOLMAE_PBPMTO, "##,###,##0.00") & ")"
      Else
         If moddat_g_str_CodPrd = "011" Then
            l_dbl_ImpPtm = CDbl(g_rst_Princi!HIPMAE_MTOPRE) + CDbl(g_rst_Princi!SOLMAE_AFPMTO)
         End If
      End If
   End If
   l_dbl_ImpPtm = l_dbl_ImpPtm - CDbl(g_rst_Princi!SOLMAE_MTOGCI)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Inicia()
    '1 = Primera Aprobacion, 2 = Segunda Aprobacion(Confirmacion)
    If (moddat_g_int_TipRep = 1) Then
       If (moddat_g_str_CodIte = "3") Then
           cmd_Aprobar.Enabled = True
           cmd_Rechazar.Enabled = True
           pnl_Titulo.Caption = "Créditos Hipotecarios - Evaluación de Operaciones"
        Else
           cmd_Aprobar.Enabled = True
           cmd_Rechazar.Enabled = True
           pnl_Titulo.Caption = "Créditos Hipotecarios - Evaluación de Operaciones 2DA Parte"
        End If
    Else
        cmd_Aprobar.Enabled = False
        cmd_Rechazar.Enabled = False
        pnl_Titulo.Caption = "Créditos Hipotecarios - Consultar Operación"
    End If
    
   If (moddat_g_str_CodIte = 3) Then '---legal aprobado
       ipp_FecConstancia.Visible = False
       lblFechaConstancia.Visible = False
       
       txt_Comentario.Top = 180
       txt_Comentario.Left = 1120
       txt_Comentario.Width = 10275
       txt_Comentario.Height = 900
       
       lblComentario.Top = 180
   End If
  
   ipp_FecConstancia.Text = date
  
   'Datos del Cliente
   grd_Listad(0).ColWidth(0) = 3060:   grd_Listad(0).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(0).ColWidth(1) = 7940:   grd_Listad(0).ColAlignment(1) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_Listad(0))

   'Datos del Cliente
   grd_Listad(2).ColWidth(0) = 3060:   grd_Listad(2).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(2).ColWidth(1) = 7940:   grd_Listad(2).ColAlignment(1) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_Listad(2))
   
   'Datos del Inmueble
   grd_Listad(1).ColWidth(0) = 3060:   grd_Listad(1).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(1).ColWidth(1) = 7940:   grd_Listad(1).ColAlignment(1) = flexAlignLeftCenter

   Call gs_LimpiaGrid(grd_Listad(1))
   
   'Datos Legal
   grd_Listad(6).ColWidth(0) = 3060:   grd_Listad(6).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(6).ColWidth(1) = 7940:   grd_Listad(6).ColAlignment(1) = flexAlignLeftCenter

   Call gs_LimpiaGrid(grd_Listad(6))

   'Datos del Crédito
   grd_Listad(4).ColWidth(0) = 3060:   grd_Listad(4).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(4).ColWidth(1) = 7940:   grd_Listad(4).ColAlignment(1) = flexAlignLeftCenter

   Call gs_LimpiaGrid(grd_Listad(4))

   'Datos del Desembolso
   grd_Listad(3).ColWidth(0) = 3060:   grd_Listad(3).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(3).ColWidth(1) = 7940:   grd_Listad(3).ColAlignment(1) = flexAlignLeftCenter

   Call gs_LimpiaGrid(grd_Listad(3))
            
   'Datos de la Evaluacion
   grd_Listad(7).ColWidth(0) = 3000:   grd_Listad(7).ColAlignment(1) = flexAlignLeftCenter
   grd_Listad(7).ColWidth(1) = 8000:   grd_Listad(7).ColAlignment(1) = flexAlignLeftCenter

   Call gs_LimpiaGrid(grd_Listad(7))
   
  'Datos Desembolso Promotor
   grd_Listad_Dsm.TextMatrix(0, 0) = "ID_FormaPago"
   grd_Listad_Dsm.TextMatrix(0, 1) = "Forma Pago"
   grd_Listad_Dsm.TextMatrix(0, 2) = "ID_TipoMonto"
   grd_Listad_Dsm.TextMatrix(0, 3) = "Tipo Monto"
   grd_Listad_Dsm.TextMatrix(0, 4) = "Importe"
   grd_Listad_Dsm.TextMatrix(0, 5) = "ID_BANCO"
   grd_Listad_Dsm.TextMatrix(0, 6) = "Entidad Financiera"
   grd_Listad_Dsm.TextMatrix(0, 7) = "Nro Cuenta"
   grd_Listad_Dsm.TextMatrix(0, 8) = "A Nombre de"
   grd_Listad_Dsm.TextMatrix(0, 9) = "Nro Desembolso"
   grd_Listad_Dsm.TextMatrix(0, 10) = "Fecha Reg."
   grd_Listad_Dsm.TextMatrix(0, 11) = "Descripcion"
   grd_Listad_Dsm.TextMatrix(0, 12) = "Flag"
   grd_Listad_Dsm.TextMatrix(0, 13) = "NumItem"
   
   If (moddat_g_int_CodIns = CInt("000001") Or moddat_g_int_CodIns = CInt("000002") Or moddat_g_int_CodIns = 0) Then
       'Legal 1 y Operaciones 1
       grd_Listad_Dsm.TextMatrix(0, 9) = ""
       grd_Listad_Dsm.TextMatrix(0, 10) = ""
   End If
   
   grd_Listad_Dsm.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad_Dsm.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad_Dsm.ColAlignment(4) = flexAlignRightCenter
   grd_Listad_Dsm.ColAlignment(6) = flexAlignCenterCenter
   grd_Listad_Dsm.ColAlignment(7) = flexAlignCenterCenter
   grd_Listad_Dsm.ColAlignment(8) = flexAlignLeftCenter
   grd_Listad_Dsm.ColAlignment(9) = flexAlignLeftCenter
   grd_Listad_Dsm.ColAlignment(10) = flexAlignLeftCenter
   grd_Listad_Dsm.ColAlignment(11) = flexAlignLeftCenter
   
   grd_Listad_Dsm.ColWidth(0) = 0      'Id-FormaPago
   grd_Listad_Dsm.ColWidth(1) = 1500   'FormaPago
   grd_Listad_Dsm.ColWidth(2) = 0      'ID_TipoMonto
   grd_Listad_Dsm.ColWidth(3) = 1500   'TipoMonto
   grd_Listad_Dsm.ColWidth(4) = 1200   'Importe
   grd_Listad_Dsm.ColWidth(5) = 0      'ID_Banco
   grd_Listad_Dsm.ColWidth(6) = 2500   'Nom_Banco
   grd_Listad_Dsm.ColWidth(7) = 1800   'Nro_Cuenta
   grd_Listad_Dsm.ColWidth(8) = 3200   'A_Nombre_DE
   grd_Listad_Dsm.ColWidth(9) = 0      'Nro_Desembolso
   grd_Listad_Dsm.ColWidth(10) = 0     'Fec_Desembolso
   grd_Listad_Dsm.ColWidth(11) = 3500  'Descripcion
   grd_Listad_Dsm.ColWidth(12) = 0     'Flag
   grd_Listad_Dsm.ColWidth(13) = 0     'NumItem
   
   If (moddat_g_int_CodIns = CInt("000001") Or moddat_g_int_CodIns = CInt("000002") Or moddat_g_int_CodIns = 0) Then
       'Legal 1 y Operaciones 1
       grd_Listad_Dsm.ColWidth(9) = 0 'Nro Desembolso
       grd_Listad_Dsm.ColWidth(10) = 0 'Fecha Reg.
   Else
       grd_Listad_Dsm.ColWidth(9) = 1900 'Nro Desembolso
       grd_Listad_Dsm.ColWidth(10) = 1020  'Fecha Reg.
   End If
      
   grd_Listad_Dsm.Row = 0
   grd_Listad_Dsm.Col = 1
   grd_Listad_Dsm.CellAlignment = flexAlignCenterCenter
   grd_Listad_Dsm.CellBackColor = &HE0E0E0
   grd_Listad_Dsm.Row = 0
   grd_Listad_Dsm.Col = 3
   grd_Listad_Dsm.CellAlignment = flexAlignCenterCenter
   grd_Listad_Dsm.CellBackColor = &HE0E0E0
   grd_Listad_Dsm.Row = 0
   grd_Listad_Dsm.Col = 4
   grd_Listad_Dsm.CellAlignment = flexAlignCenterCenter
   grd_Listad_Dsm.CellBackColor = &HE0E0E0
   grd_Listad_Dsm.Row = 0
   grd_Listad_Dsm.Col = 6
   grd_Listad_Dsm.CellAlignment = flexAlignCenterCenter
   grd_Listad_Dsm.CellBackColor = &HE0E0E0
   grd_Listad_Dsm.Row = 0
   grd_Listad_Dsm.Col = 7
   grd_Listad_Dsm.CellAlignment = flexAlignCenterCenter
   grd_Listad_Dsm.CellBackColor = &HE0E0E0
   grd_Listad_Dsm.Row = 0
   grd_Listad_Dsm.Col = 8
   grd_Listad_Dsm.CellAlignment = flexAlignCenterCenter
   grd_Listad_Dsm.CellBackColor = &HE0E0E0
   grd_Listad_Dsm.Row = 0
   grd_Listad_Dsm.Col = 9
   grd_Listad_Dsm.CellAlignment = flexAlignCenterCenter
   grd_Listad_Dsm.CellBackColor = &HE0E0E0
   grd_Listad_Dsm.Row = 0
   grd_Listad_Dsm.Col = 10
   grd_Listad_Dsm.CellAlignment = flexAlignCenterCenter
   grd_Listad_Dsm.CellBackColor = &HE0E0E0
   grd_Listad_Dsm.Row = 0
   grd_Listad_Dsm.Col = 11
   grd_Listad_Dsm.CellAlignment = flexAlignCenterCenter
   grd_Listad_Dsm.CellBackColor = &HE0E0E0
   
   Call gs_UbiIniGrid(grd_Listad_Dsm)
End Sub

Private Sub ipp_FecConstancia_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 13) Then
       txt_Comentario.SetFocus
    End If
End Sub

Private Sub txt_Comentario_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cmd_Dsm_Nuevo_Click()
   Call fs_MntBnt_Dsm(1) 'Cancelar
   cmb_FrmDsm_Dsm.ListIndex = 0
   cmd_Dsm_Insert.Tag = 1
   
   If (moddat_g_int_CodIns = CInt("000003")) Then
      'Codigo de area Tesoreria
       txt_NroDsm_Dsm.Enabled = True
       ipp_FecDsm_Dsm.Enabled = True
   Else
       txt_NroDsm_Dsm.Enabled = False
       ipp_FecDsm_Dsm.Enabled = False
   End If
End Sub
    
Private Sub cmd_Dsm_Borrar_Click()
   If grd_Listad_Dsm.Rows = 1 Then
      Exit Sub
   End If
   If grd_Listad_Dsm.Row = 0 Then
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de borrar el item ?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
       
   If (grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 12) = "I") Then
       grd_Listad_Dsm.RemoveItem (grd_Listad_Dsm.Row)
   Else
       grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 12) = "D"
       grd_Listad_Dsm.RowHeight(grd_Listad_Dsm.Row) = 0
   End If
   
   Call fs_sumarDesemPrmt
   Call fs_MntBnt_Dsm(4) 'Cancelar
End Sub

Private Sub cmd_Dsm_Editar_Click()
   If (grd_Listad_Dsm.Rows = 1) Then
       Exit Sub
   End If
   If (grd_Listad_Dsm.Row = 0) Then
       Exit Sub
   End If
   Call fs_MntBnt_Dsm(2) 'Editar
   
   cmd_Dsm_Insert.Tag = 2
   
   Call fs_HabFormDsm
   Call fs_mostrar_Datos
   
   If (moddat_g_int_CodIns = CInt("000003")) Then
      'Codigo de area Tesoreria
       txt_NroDsm_Dsm.Enabled = True
       ipp_FecDsm_Dsm.Enabled = True
   Else
       txt_NroDsm_Dsm.Enabled = False
       ipp_FecDsm_Dsm.Enabled = False
   End If
End Sub

Private Sub grd_Listad_Dsm_SelChange()
   cmb_EntFin_Dsm.ListIndex = -1
   cmb_NroCta_Dsm.ListIndex = -1
   pnl_NroCCI_Dsm.Caption = ""
   cmb_FrmDsm_Dsm.ListIndex = -1
   txt_NroDsm_Dsm.Text = ""
   ipp_FecDsm_Dsm.Text = ""
   txt_Descrp_Dsm.Text = ""
   txt_ANombre_Dsm.Text = ""
   ipp_Import_Dsm.Text = "0.00"
   pnl_Moneda_Dsm.Caption = ""
   
   If (grd_Listad_Dsm.Rows = 1) Then
       Exit Sub
   End If
   If (grd_Listad_Dsm.Row = 0) Then
       Exit Sub
   End If
       
   Call fs_mostrar_Datos
   cmb_EntFin_Dsm.Enabled = False
   cmb_NroCta_Dsm.Enabled = False
End Sub

Private Sub fs_mostrar_Datos()
   Call gs_BuscarCombo_Item(cmb_FrmDsm_Dsm, grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 0))
   Call gs_BuscarCombo_Item(cmb_TipMto_Dsm, grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 2))
   cmb_EntFin_Dsm.Text = Trim(grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 6))
   Call cmb_EntFin_Dsm_Click
   If grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 7) <> "" Then
      cmb_NroCta_Dsm.Text = Trim(grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 7) & "")
   End If
   ipp_Import_Dsm.Text = grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 4)
   txt_ANombre_Dsm.Text = grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 8)
   txt_NroDsm_Dsm.Text = grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 9)
   ipp_FecDsm_Dsm.Text = grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 10)
   txt_Descrp_Dsm.Text = grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 11)
   
   pnl_Moneda_Dsm.Caption = l_str_Moneda
End Sub

Private Sub cmd_Dsm_Cancel_Click()
   Call fs_MntBnt_Dsm(3) 'Cancelar
End Sub

Private Sub cmd_Dsm_Insert_Click()
Dim r_bol_Estado   As Boolean
Dim r_int_fila     As Integer
Dim r_dbl_suma     As Double
Dim r_int_NumIte   As Integer
   
   If cmb_FrmDsm_Dsm.ListIndex = -1 Then
      MsgBox "Debe seleccionar el tipo de desembolso.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_FrmDsm_Dsm)
      Exit Sub
   End If
   
   If cmb_TipMto_Dsm.ListIndex = -1 Then
      MsgBox "Debe seleccionar el tipo de monto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipMto_Dsm)
      Exit Sub
   End If
   
   If CDbl(Trim(ipp_Import_Dsm.Text)) = 0 Then
      MsgBox "Debe digitar un importe.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Import_Dsm)
      Exit Sub
   End If
   
   If cmb_EntFin_Dsm.ListIndex = -1 Then
      MsgBox "Debe seleccionar una entidad financiera.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_EntFin_Dsm)
      Exit Sub
   End If
   
   If cmb_FrmDsm_Dsm.ItemData(cmb_FrmDsm_Dsm.ListIndex) = 2 Then
      If cmb_NroCta_Dsm.ListIndex = -1 Then
         MsgBox "Debe seleccionar el nro de cuenta.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_NroCta_Dsm)
         Exit Sub
      End If
   End If
   
   If Len(Trim(txt_ANombre_Dsm.Text)) = 0 Then
      MsgBox "Debe digitar a nombre de quien va " & IIf(UCase(Left(cmb_FrmDsm_Dsm.Text, 1)) = "T", "la ", "el ") & Trim(cmb_FrmDsm_Dsm.Text), vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_ANombre_Dsm)
      Exit Sub
   End If

   If (moddat_g_int_CodIns = CInt("000003")) Then
      'Codigo de area Tesoreria
       If Len(Trim(txt_NroDsm_Dsm.Text)) = 0 Then
          If (cmb_FrmDsm_Dsm.ListIndex = 0) Then
              MsgBox "Debe digitar el nro de cheque.", vbExclamation, modgen_g_str_NomPlt
          Else
              MsgBox "Debe digitar el nro transferencia.", vbExclamation, modgen_g_str_NomPlt
          End If
          Call gs_SetFocus(txt_NroDsm_Dsm)
          Exit Sub
       End If
      
       If Len(Trim(ipp_FecDsm_Dsm.Text)) = 0 Then
          If (cmb_FrmDsm_Dsm.ListIndex = 0) Then
              MsgBox "Debe Digitar la fecha de registro del Cheque.", vbExclamation, modgen_g_str_NomPlt
          Else
              MsgBox "Debe Digitar la fecha de registro de Transferencia.", vbExclamation, modgen_g_str_NomPlt
          End If
          Call gs_SetFocus(ipp_FecDsm_Dsm)
          Exit Sub
       End If
   End If
      
   If (cmd_Dsm_Insert.Tag = 2) Then
       'ACTUALIZAR
       r_dbl_suma = 0
       r_dbl_suma = CDbl(grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 4))
       r_dbl_suma = CDbl(pnl_SumTot_Dsm.Caption) + CDbl(ipp_Import_Dsm.Text) - r_dbl_suma
       If (l_dbl_ImpPtm < r_dbl_suma) Then
           MsgBox "La suma de registros sobrepasa el importe del préstamo.", vbExclamation, modgen_g_str_NomPlt
           Call gs_SetFocus(ipp_Import_Dsm)
           Exit Sub
       End If
   
       grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 0) = cmb_FrmDsm_Dsm.ItemData(cmb_FrmDsm_Dsm.ListIndex)
       grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 1) = cmb_FrmDsm_Dsm.Text
       grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 2) = cmb_TipMto_Dsm.ItemData(cmb_TipMto_Dsm.ListIndex)
       grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 3) = cmb_TipMto_Dsm.Text
       grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 4) = ipp_Import_Dsm.Text
       grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 5) = CStr(l_arr_Bancos(cmb_EntFin_Dsm.ListIndex + 1).Genera_Codigo)
       grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 6) = cmb_EntFin_Dsm.Text
       grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 7) = cmb_NroCta_Dsm.Text
       grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 8) = txt_ANombre_Dsm.Text
       grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 9) = txt_NroDsm_Dsm.Text
       grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 10) = ipp_FecDsm_Dsm.Text
       grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 11) = txt_Descrp_Dsm.Text
       
       pnl_Moneda_Dsm.Caption = ""
           
       If (UCase(Trim(grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 12))) <> UCase(Trim("I"))) Then
           grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 12) = "U"
       End If
   Else
   'INSERTAR
       r_dbl_suma = 0
       r_dbl_suma = CDbl(pnl_SumTot_Dsm.Caption) + CDbl(ipp_Import_Dsm.Text)
       If (l_dbl_ImpPtm < r_dbl_suma) Then
           MsgBox "La suma de registros sobrepasa el importe del préstamo.", vbExclamation, modgen_g_str_NomPlt
           Call gs_SetFocus(ipp_Import_Dsm)
           Exit Sub
       End If
       
       'Genera correlativo
       r_int_NumIte = 0
       For r_int_fila = 1 To grd_Listad_Dsm.Rows - 1
           If (r_int_NumIte <= grd_Listad_Dsm.TextMatrix(r_int_fila, 13)) Then
               r_int_NumIte = grd_Listad_Dsm.TextMatrix(r_int_fila, 13)
           End If
       Next
       r_int_NumIte = r_int_NumIte + 1
       
       grd_Listad_Dsm.Rows = grd_Listad_Dsm.Rows + 1
       grd_Listad_Dsm.Row = grd_Listad_Dsm.Rows - 1
       
       grd_Listad_Dsm.Col = 0
       grd_Listad_Dsm.Text = cmb_FrmDsm_Dsm.ItemData(cmb_FrmDsm_Dsm.ListIndex)
       grd_Listad_Dsm.Col = 1
       grd_Listad_Dsm.Text = cmb_FrmDsm_Dsm.Text
       
       grd_Listad_Dsm.Col = 2
       grd_Listad_Dsm.Text = cmb_TipMto_Dsm.ItemData(cmb_TipMto_Dsm.ListIndex)
       grd_Listad_Dsm.Col = 3
       grd_Listad_Dsm.Text = cmb_TipMto_Dsm.Text
       
       grd_Listad_Dsm.Col = 4
       grd_Listad_Dsm.Text = ipp_Import_Dsm.Text
       
       grd_Listad_Dsm.Col = 5
       grd_Listad_Dsm.Text = CStr(l_arr_Bancos(cmb_EntFin_Dsm.ListIndex + 1).Genera_Codigo)
   
       grd_Listad_Dsm.Col = 6
       grd_Listad_Dsm.Text = cmb_EntFin_Dsm.Text
       
       grd_Listad_Dsm.Col = 7
       grd_Listad_Dsm.Text = cmb_NroCta_Dsm.Text
       
       grd_Listad_Dsm.Col = 8
       grd_Listad_Dsm.Text = txt_ANombre_Dsm.Text
                     
       grd_Listad_Dsm.Col = 9
       grd_Listad_Dsm.Text = txt_NroDsm_Dsm.Text
       grd_Listad_Dsm.Col = 10
       grd_Listad_Dsm.Text = ipp_FecDsm_Dsm.Text
   
       grd_Listad_Dsm.Col = 11
       grd_Listad_Dsm.Text = txt_Descrp_Dsm.Text
          
       grd_Listad_Dsm.Col = 12
       grd_Listad_Dsm.Text = "I"
       
       grd_Listad_Dsm.Col = 13
       grd_Listad_Dsm.Text = r_int_NumIte
   End If
      
   Call fs_sumarDesemPrmt
   Call fs_MntBnt_Dsm(3) 'Agregar
End Sub

Private Sub fs_sumarDesemPrmt()
Dim r_int_fila   As Integer
Dim r_dbl_suma   As Double
    
    r_dbl_suma = 0
    For r_int_fila = 1 To grd_Listad_Dsm.Rows - 1
        If (grd_Listad_Dsm.RowHeight(r_int_fila) > 0) Then
            r_dbl_suma = r_dbl_suma + CDbl(grd_Listad_Dsm.TextMatrix(r_int_fila, 4))
        End If
    Next
    pnl_SumTot_Dsm.Caption = gf_FormatoNumero(r_dbl_suma, 12, 2) & " "
    
    pnl_TotPtmo_Dsm.Caption = l_dbl_ImpPtm - r_dbl_suma
    pnl_TotPtmo_Dsm.Caption = gf_FormatoNumero(pnl_TotPtmo_Dsm.Caption, 12, 2) & " "
End Sub

Private Sub fs_MntBnt_Dsm(p_Tipo As Integer)
'Desabilitar = 0; Nuevo = 1; Editar = 2; Agregar = 3; Cancelar = 4
   If (p_Tipo = 0) Then '---desabilitar----
       cmd_Dsm_Nuevo.Enabled = False
       cmd_Dsm_Borrar.Enabled = False
       cmd_Dsm_Editar.Enabled = False
       cmd_Dsm_Insert.Enabled = False
       cmd_Dsm_Cancel.Enabled = False
       cmb_NroCta_Dsm.ListIndex = -1
       pnl_NroCCI_Dsm.Caption = ""
       cmb_FrmDsm_Dsm.ListIndex = -1
       cmb_TipMto_Dsm.ListIndex = -1
       cmb_EntFin_Dsm.ListIndex = -1
       txt_NroDsm_Dsm.Text = ""
       ipp_FecDsm_Dsm.Text = ""
       txt_Descrp_Dsm.Text = ""
       txt_ANombre_Dsm.Text = ""
       ipp_Import_Dsm.Text = "0.00"
       pnl_Moneda_Dsm.Caption = l_str_Moneda
       cmb_EntFin_Dsm.Enabled = False
       cmb_NroCta_Dsm.Enabled = False
       cmb_FrmDsm_Dsm.Enabled = False
       cmb_TipMto_Dsm.Enabled = False
       txt_NroDsm_Dsm.Enabled = False
       ipp_FecDsm_Dsm.Enabled = False
       txt_Descrp_Dsm.Enabled = False
       txt_ANombre_Dsm.Enabled = False
       ipp_Import_Dsm.Enabled = False
   ElseIf (p_Tipo = 1) Then '---nuevo----
       cmd_Dsm_Nuevo.Enabled = False
       cmd_Dsm_Borrar.Enabled = False
       cmd_Dsm_Editar.Enabled = False
       cmd_Dsm_Insert.Enabled = True
       cmd_Dsm_Cancel.Enabled = True
       cmb_NroCta_Dsm.ListIndex = -1
       pnl_NroCCI_Dsm.Caption = ""
       cmb_FrmDsm_Dsm.ListIndex = -1
       cmb_TipMto_Dsm.ListIndex = -1
       cmb_EntFin_Dsm.ListIndex = -1
       txt_NroDsm_Dsm.Text = ""
       ipp_FecDsm_Dsm.Text = ""
       txt_Descrp_Dsm.Text = ""
       txt_ANombre_Dsm.Text = ""
       ipp_Import_Dsm.Text = "0.00"
       pnl_Moneda_Dsm.Caption = ""
       cmb_EntFin_Dsm.Enabled = True
       cmb_NroCta_Dsm.Enabled = True
       cmb_FrmDsm_Dsm.Enabled = True
       cmb_TipMto_Dsm.Enabled = True
       txt_NroDsm_Dsm.Enabled = True
       ipp_FecDsm_Dsm.Enabled = True
       txt_Descrp_Dsm.Enabled = True
       txt_ANombre_Dsm.Enabled = True
       ipp_Import_Dsm.Enabled = True
       pnl_Moneda_Dsm.Caption = l_str_Moneda
       grd_Listad_Dsm.Enabled = False
       Call gs_UbiIniGrid(grd_Listad_Dsm)
       Call gs_SetFocus(cmb_FrmDsm_Dsm)
   ElseIf (p_Tipo = 2) Then '---Editar-----
       cmd_Dsm_Nuevo.Enabled = False
       cmd_Dsm_Borrar.Enabled = False
       cmd_Dsm_Editar.Enabled = False
       cmd_Dsm_Insert.Enabled = True
       cmd_Dsm_Cancel.Enabled = True
       cmb_EntFin_Dsm.Enabled = True
       cmb_NroCta_Dsm.Enabled = True
       cmb_FrmDsm_Dsm.Enabled = True
       cmb_TipMto_Dsm.Enabled = True
       txt_NroDsm_Dsm.Enabled = True
       ipp_FecDsm_Dsm.Enabled = True
       txt_Descrp_Dsm.Enabled = True
       txt_ANombre_Dsm.Enabled = True
       ipp_Import_Dsm.Enabled = True
       pnl_Moneda_Dsm.Caption = l_str_Moneda
       grd_Listad_Dsm.Enabled = False
       Call gs_SetFocus(cmb_FrmDsm_Dsm)
   ElseIf (p_Tipo = 3) Then '---Agregar-----
       cmd_Dsm_Nuevo.Enabled = True
       cmd_Dsm_Borrar.Enabled = True
       cmd_Dsm_Editar.Enabled = True
       cmd_Dsm_Insert.Enabled = False
       cmd_Dsm_Cancel.Enabled = False
       cmb_EntFin_Dsm.ListIndex = -1
       cmb_NroCta_Dsm.ListIndex = -1
       pnl_NroCCI_Dsm.Caption = ""
       cmb_FrmDsm_Dsm.ListIndex = -1
       cmb_TipMto_Dsm.ListIndex = -1
       txt_NroDsm_Dsm.Text = ""
       ipp_FecDsm_Dsm.Text = ""
       txt_Descrp_Dsm.Text = ""
       txt_ANombre_Dsm.Text = ""
       ipp_Import_Dsm.Text = "0.00"
       pnl_Moneda_Dsm.Caption = ""
       cmb_EntFin_Dsm.Enabled = False
       cmb_NroCta_Dsm.Enabled = False
       cmb_FrmDsm_Dsm.Enabled = False
       cmb_TipMto_Dsm.Enabled = False
       txt_NroDsm_Dsm.Enabled = False
       ipp_FecDsm_Dsm.Enabled = False
       txt_Descrp_Dsm.Enabled = False
       txt_ANombre_Dsm.Enabled = False
       ipp_Import_Dsm.Enabled = False
       grd_Listad_Dsm.Enabled = True
       Call gs_UbiIniGrid(grd_Listad_Dsm)
       Call gs_SetFocus(cmd_Dsm_Nuevo)
   ElseIf (p_Tipo = 4) Then '---Cancelar-----
       cmd_Dsm_Nuevo.Enabled = True
       cmd_Dsm_Borrar.Enabled = True
       cmd_Dsm_Editar.Enabled = True
       cmd_Dsm_Insert.Enabled = False
       cmd_Dsm_Cancel.Enabled = False
       cmb_EntFin_Dsm.ListIndex = -1
       cmb_NroCta_Dsm.ListIndex = -1
       pnl_NroCCI_Dsm.Caption = ""
       cmb_FrmDsm_Dsm.ListIndex = -1
       cmb_TipMto_Dsm.ListIndex = -1
       txt_NroDsm_Dsm.Text = ""
       ipp_FecDsm_Dsm.Text = ""
       txt_Descrp_Dsm.Text = ""
       txt_ANombre_Dsm.Text = ""
       ipp_Import_Dsm.Text = "0.00"
       pnl_Moneda_Dsm.Caption = ""
       cmb_EntFin_Dsm.Enabled = False
       cmb_NroCta_Dsm.Enabled = False
       cmb_FrmDsm_Dsm.Enabled = False
       cmb_TipMto_Dsm.Enabled = False
       txt_NroDsm_Dsm.Enabled = False
       ipp_FecDsm_Dsm.Enabled = False
       txt_Descrp_Dsm.Enabled = False
       txt_ANombre_Dsm.Enabled = False
       ipp_Import_Dsm.Enabled = False
       grd_Listad_Dsm.Enabled = True
       Call gs_UbiIniGrid(grd_Listad_Dsm)
   End If
End Sub

Private Sub cmb_FrmDsm_Dsm_Click()
    If (cmb_FrmDsm_Dsm.ListIndex <> -1) Then
        lbl_NumDsm_Dsm.Caption = "Nro " & UCase(Left(Trim(cmb_FrmDsm_Dsm.Text), 1)) & LCase(Right(Trim(cmb_FrmDsm_Dsm.Text), Len(Trim(cmb_FrmDsm_Dsm.Text)) - 1))
        lbl_FchDsm_Dsm.Caption = "Fecha Reg " & UCase(Left(Trim(cmb_FrmDsm_Dsm.Text), 1)) & LCase(Right(Left(Trim(cmb_FrmDsm_Dsm.Text), 4), 3))
        'lbl_NumDsm_Dsm.Caption = lbl_NumDsm_Dsm.Caption & ":"
        lbl_NumDsm_Dsm.Caption = Mid(Trim(lbl_NumDsm_Dsm.Caption), 1, 18) & ":"
        lbl_FchDsm_Dsm.Caption = lbl_FchDsm_Dsm.Caption & ":"
    Else
        lbl_NumDsm_Dsm.Caption = "Nro Desembolso:"
        lbl_FchDsm_Dsm.Caption = "Fecha Reg Dsm:"
    End If
    Call fs_HabFormDsm
End Sub

Private Sub fs_HabFormDsm()
    If cmd_Dsm_Editar.Enabled = False And cmd_Grabar.Enabled = True Then
       If cmb_FrmDsm_Dsm.ListIndex > -1 Then
          If cmb_FrmDsm_Dsm.ItemData(cmb_FrmDsm_Dsm.ListIndex) = 1 Or cmb_FrmDsm_Dsm.ItemData(cmb_FrmDsm_Dsm.ListIndex) = 3 Or _
             cmb_FrmDsm_Dsm.ItemData(cmb_FrmDsm_Dsm.ListIndex) = 4 Then
             cmb_EntFin_Dsm.ListIndex = -1
             cmb_NroCta_Dsm.ListIndex = -1
             pnl_NroCCI_Dsm.Caption = ""
             cmb_NroCta_Dsm.Enabled = False
          Else
             cmb_EntFin_Dsm.ListIndex = -1
             cmb_EntFin_Dsm.Enabled = True
             cmb_NroCta_Dsm.Enabled = True
          End If
       End If
    End If
End Sub

Private Sub cmd_Dsm_ExpExc_Click()
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
      
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmb_EntFin_Dsm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_NroCta_Dsm.Enabled = False Then
          Call gs_SetFocus(txt_ANombre_Dsm)
      Else
          Call gs_SetFocus(cmb_NroCta_Dsm)
      End If
   End If
End Sub

Private Sub cmb_NroCta_Dsm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ANombre_Dsm)
   End If
End Sub

Private Sub cmb_FrmDsm_Dsm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipMto_Dsm)
   End If
End Sub

Private Sub ipp_Import_Dsm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_EntFin_Dsm.Enabled = False Then
         Call gs_SetFocus(txt_ANombre_Dsm)
      Else
         Call gs_SetFocus(cmb_EntFin_Dsm)
      End If
   End If
End Sub

Private Sub txt_ANombre_Dsm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Descrp_Dsm)
   Else
      'KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & " '")
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- _.")
   End If
End Sub

Private Sub txt_Descrp_Dsm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If (txt_NroDsm_Dsm.Visible = False) Then
          Call gs_SetFocus(cmd_Dsm_Insert)
      Else
          Call gs_SetFocus(txt_NroDsm_Dsm)
      End If
   Else
      'KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & " '")
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub

Private Sub txt_NroDsm_Dsm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecDsm_Dsm)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-'")
   End If
End Sub

Private Sub ipp_FecDsm_Dsm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Dsm_Insert)
   End If
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel   As EXCEL.Application
Dim r_int_Filaux  As Integer
Dim r_int_FilExl  As Integer
Dim r_int_totExl  As Integer
Dim r_str_Cadena  As String
                
   Screen.MousePointer = 11
   Set r_obj_Excel = New EXCEL.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
        'Unir celdas
        .Range("B2") = "NRO OPERACION:"
        .Range("B3") = "CLIENTE:"
        .Range("B4") = "PRODUCTO:"
        .Range("B5") = "PROYECTO:"
        .Range("B6") = "PROMOTOR:"
        
        .Range("C2") = pnl_NumOpe.Caption
        .Range("C3") = Trim(pnl_NomCli.Caption)
        .Range("C4") = Trim(pnl_Produc.Caption)
        .Range("C5") = Trim(pnl_Prycto_Dsm.Caption)
        .Range("C6") = Trim(l_str_Prmtor)
        
        If (moddat_g_int_CodIns = CInt("000001") Or moddat_g_int_CodIns = CInt("000002") Or moddat_g_int_CodIns = 0) Then
           'Legal 1 y Operaciones 1
           r_int_totExl = 8
           r_str_Cadena = "H"
        Else
           r_int_totExl = 10
           r_str_Cadena = "J"
        End If
        
        r_int_FilExl = 8
        .Range("B" & r_int_FilExl) = "DATOS DE DESEMBOLSO A PROMOTOR"
        .Range("B" & r_int_FilExl & ":" & r_str_Cadena & r_int_FilExl).Font.Bold = True
        .Range("B" & r_int_FilExl & ":" & r_str_Cadena & r_int_FilExl).Merge
        .Range("B" & r_int_FilExl).HorizontalAlignment = xlHAlignCenter
        
        r_int_FilExl = r_int_FilExl + 1
        .Columns("G").HorizontalAlignment = xlHAlignLeft
        .Range("B" & r_int_FilExl & ":" & r_str_Cadena & r_int_FilExl).HorizontalAlignment = xlHAlignCenter
        .Range("B" & r_int_FilExl & ":" & r_str_Cadena & r_int_FilExl).Font.Bold = True
        .Range("B" & r_int_FilExl & ":" & r_str_Cadena & r_int_FilExl).Interior.Color = RGB(146, 208, 80)
        
        For r_int_Filaux = 2 To r_int_totExl
            .Cells(r_int_FilExl, r_int_Filaux).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Cells(r_int_FilExl, r_int_Filaux).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Cells(r_int_FilExl, r_int_Filaux).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Cells(r_int_FilExl, r_int_Filaux).Borders(xlEdgeBottom).LineStyle = xlContinuous
        Next
                    
        .Columns("A").ColumnWidth = 3
        .Columns("B").ColumnWidth = 14
        .Columns("C").ColumnWidth = 12
        .Columns("D").ColumnWidth = 26
        .Columns("E").ColumnWidth = 18
        .Columns("F").ColumnWidth = 13
        .Columns("G").ColumnWidth = 30
        If (moddat_g_int_CodIns = CInt("000001") Or moddat_g_int_CodIns = CInt("000002") Or moddat_g_int_CodIns = 0) Then
           'Legal 1 y Operaciones 1
           .Columns("H").ColumnWidth = 30
        Else
           .Columns("H").ColumnWidth = 30
           .Columns("I").ColumnWidth = 11
           .Columns("J").ColumnWidth = 42
        End If
              
        .Cells(r_int_FilExl, 2) = "Forma Pago"
        .Cells(r_int_FilExl, 3) = "Tipo Monto"
        .Cells(r_int_FilExl, 4) = "Banco"
        .Cells(r_int_FilExl, 5) = "Nro Cuenta"
        .Cells(r_int_FilExl, 6) = "Importe"
        .Cells(r_int_FilExl, 7) = "A Nombre de"
        If (moddat_g_int_CodIns = CInt("000001") Or moddat_g_int_CodIns = CInt("000002") Or moddat_g_int_CodIns = 0) Then
           'Legal 1 y Operaciones 1
           .Cells(r_int_FilExl, 8) = "Descripción"
        Else
           .Cells(r_int_FilExl, 8) = "Nro Desembolso"
           .Cells(r_int_FilExl, 9) = "Fecha Reg."
           .Cells(r_int_FilExl, 10) = "Descripción"
        End If
                
         For r_int_Filaux = 1 To grd_Listad_Dsm.Rows - 1
             r_int_FilExl = r_int_FilExl + 1
             .Cells(r_int_FilExl, 2).NumberFormat = "@"
             .Cells(r_int_FilExl, 3).NumberFormat = "@"
             .Cells(r_int_FilExl, 5).NumberFormat = "@"
             .Cells(r_int_FilExl, 6).NumberFormat = "###,###,##0.00" '"@"
             .Cells(r_int_FilExl, 7).NumberFormat = "@"
             .Cells(r_int_FilExl, 8).NumberFormat = "@"
             .Cells(r_int_FilExl, 9).NumberFormat = "@"
             .Cells(r_int_FilExl, 10).NumberFormat = "@"
             
             .Cells(r_int_FilExl, 2) = grd_Listad_Dsm.TextMatrix(r_int_Filaux, 1)
             .Cells(r_int_FilExl, 3) = grd_Listad_Dsm.TextMatrix(r_int_Filaux, 3)
             .Cells(r_int_FilExl, 4) = grd_Listad_Dsm.TextMatrix(r_int_Filaux, 6)
             .Cells(r_int_FilExl, 5) = grd_Listad_Dsm.TextMatrix(r_int_Filaux, 7)
             .Cells(r_int_FilExl, 6) = grd_Listad_Dsm.TextMatrix(r_int_Filaux, 4)
             .Cells(r_int_FilExl, 7) = grd_Listad_Dsm.TextMatrix(r_int_Filaux, 8)
             If (moddat_g_int_CodIns = CInt("000001") Or moddat_g_int_CodIns = CInt("000002") Or moddat_g_int_CodIns = 0) Then
                'Legal 1 y Operaciones 1
                .Cells(r_int_FilExl, 8) = grd_Listad_Dsm.TextMatrix(r_int_Filaux, 11)
             Else
                .Cells(r_int_FilExl, 8) = grd_Listad_Dsm.TextMatrix(r_int_Filaux, 9)
                .Cells(r_int_FilExl, 9) = grd_Listad_Dsm.TextMatrix(r_int_Filaux, 10)
                .Cells(r_int_FilExl, 10) = grd_Listad_Dsm.TextMatrix(r_int_Filaux, 11)
             End If
         Next
         
         r_int_FilExl = r_int_FilExl + 1
         .Cells(r_int_FilExl, 5) = "Suma Total ==>"
         .Cells(r_int_FilExl, 6) = pnl_SumTot_Dsm.Caption
         .Range("E" & r_int_FilExl & ":F" & r_int_FilExl).Interior.Color = RGB(146, 208, 80)
         .Range("E" & r_int_FilExl & ":F" & r_int_FilExl).Font.Bold = True
         
         .Range("A1:J" & r_int_FilExl).Font.Name = "Arial"
         .Range("A1:J" & r_int_FilExl).Font.Size = 8
   End With

   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub cmb_EntFin_Dsm_Click()
Dim r_int_fila As Integer

    cmb_NroCta_Dsm.Clear
    txt_ANombre_Dsm.Text = ""
    If cmb_FrmDsm_Dsm.ListIndex > -1 Then
       If cmb_FrmDsm_Dsm.ItemData(cmb_FrmDsm_Dsm.ListIndex) = 2 Then
          'transferencia
          For r_int_fila = 1 To UBound(l_arr_CtaBco)
              If (cmd_Dsm_Insert.Enabled = False) Then
                  If (l_arr_CtaBco(r_int_fila).Genera_Codigo = CStr(l_arr_Bancos(cmb_EntFin_Dsm.ListIndex + 1).Genera_Codigo)) Then
                      cmb_NroCta_Dsm.AddItem (Trim(l_arr_CtaBco(r_int_fila).Genera_Nombre))
                  End If
              Else
                  If (cmd_Dsm_Insert.Tag = 1) Then
                      If (l_arr_CtaBco(r_int_fila).Genera_FlgAso = 1 And _
                          l_arr_CtaBco(r_int_fila).Genera_Codigo = CStr(l_arr_Bancos(cmb_EntFin_Dsm.ListIndex + 1).Genera_Codigo)) Then
                          cmb_NroCta_Dsm.AddItem (Trim(l_arr_CtaBco(r_int_fila).Genera_Nombre))
                      End If
                  Else
                      If (l_arr_CtaBco(r_int_fila).Genera_Codigo = CStr(l_arr_Bancos(cmb_EntFin_Dsm.ListIndex + 1).Genera_Codigo)) Then
                          cmb_NroCta_Dsm.AddItem (Trim(l_arr_CtaBco(r_int_fila).Genera_Nombre))
                      End If
                  End If
              End If
         Next
       Else
        'cheque
        For r_int_fila = 1 To UBound(l_arr_CtaBco)
            If (l_arr_CtaBco(r_int_fila).Genera_Codigo = CStr(l_arr_Bancos(cmb_EntFin_Dsm.ListIndex + 1).Genera_Codigo)) Then
               txt_ANombre_Dsm.Text = Trim(l_arr_CtaBco(r_int_fila).Genera_ConHip)
               Exit For
            End If
        Next
       End If
    End If
End Sub

