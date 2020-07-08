VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Begin VB.Form frm_Ges_TecPro_15 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14070
   Icon            =   "OpeTra_frm_837.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   14070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   9060
      Left            =   30
      TabIndex        =   59
      Top             =   30
      Width           =   14040
      _Version        =   65536
      _ExtentX        =   24765
      _ExtentY        =   15981
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
         Height          =   1425
         Left            =   60
         TabIndex        =   60
         Top             =   7590
         Width           =   13905
         _Version        =   65536
         _ExtentX        =   24527
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
         Begin Threed.SSPanel pnl_AreTer 
            Height          =   315
            Left            =   1890
            TabIndex        =   61
            Top             =   270
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
         Begin Threed.SSPanel pnl_AreCon 
            Height          =   315
            Left            =   6420
            TabIndex        =   62
            Top             =   270
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   60
            Left            =   30
            TabIndex        =   63
            Top             =   630
            Width           =   13845
            _Version        =   65536
            _ExtentX        =   24421
            _ExtentY        =   106
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
         End
         Begin Threed.SSPanel pnl_SumAse 
            Height          =   315
            Left            =   1890
            TabIndex        =   64
            Top             =   720
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
         Begin Threed.SSPanel pnl_ValCom 
            Height          =   315
            Left            =   6420
            TabIndex        =   65
            Top             =   720
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
         Begin Threed.SSPanel pnl_ValRea 
            Height          =   315
            Left            =   11280
            TabIndex        =   66
            Top             =   720
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
         Begin Threed.SSPanel pnl_ValTer 
            Height          =   315
            Left            =   1890
            TabIndex        =   67
            Top             =   1050
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
         Begin Threed.SSPanel pnl_ValEdi 
            Height          =   315
            Left            =   6420
            TabIndex        =   68
            Top             =   1050
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
         Begin Threed.SSPanel pnl_ValACo 
            Height          =   315
            Left            =   11280
            TabIndex        =   69
            Top             =   1050
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            Caption         =   "Valor Areas Comunes:"
            Height          =   195
            Left            =   9180
            TabIndex        =   78
            Top             =   1110
            Width           =   1560
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            Caption         =   "Valor Edificación:"
            Height          =   195
            Left            =   4560
            TabIndex        =   77
            Top             =   1110
            Width           =   1230
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            Caption         =   "Valor Terreno:"
            Height          =   195
            Left            =   90
            TabIndex        =   76
            Top             =   1110
            Width           =   1005
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            Caption         =   "Valor Realización:"
            Height          =   195
            Left            =   9180
            TabIndex        =   75
            Top             =   780
            Width           =   1275
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            Caption         =   "Valor Comercial:"
            Height          =   195
            Left            =   4560
            TabIndex        =   74
            Top             =   780
            Width           =   1140
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            Caption         =   "Suma Asegurada:"
            Height          =   195
            Left            =   90
            TabIndex        =   73
            Top             =   780
            Width           =   1260
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "Area Construcción:"
            Height          =   195
            Left            =   4560
            TabIndex        =   72
            Top             =   330
            Width           =   1350
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Area Terreno:"
            Height          =   195
            Left            =   90
            TabIndex        =   71
            Top             =   330
            Width           =   975
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "Totales"
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
            Left            =   90
            TabIndex        =   70
            Top             =   30
            Width           =   645
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   60
         TabIndex        =   79
         Top             =   750
         Width           =   13905
         _Version        =   65536
         _ExtentX        =   24527
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
            Left            =   13290
            Picture         =   "OpeTra_frm_837.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   101
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Cancel 
            Height          =   585
            Left            =   12720
            Picture         =   "OpeTra_frm_837.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   100
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   12120
            Picture         =   "OpeTra_frm_837.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   99
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   600
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   600
            Picture         =   "OpeTra_frm_837.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   58
            ToolTipText     =   "Modificar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_837.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   2445
         Left            =   60
         TabIndex        =   80
         Top             =   5100
         Width           =   13905
         _Version        =   65536
         _ExtentX        =   24527
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
         BevelOuter      =   1
         Begin TabDlg.SSTab tab_Genera 
            Height          =   2295
            Left            =   60
            TabIndex        =   102
            Top             =   90
            Width           =   13800
            _ExtentX        =   24342
            _ExtentY        =   4048
            _Version        =   393216
            Style           =   1
            Tabs            =   6
            Tab             =   1
            TabsPerRow      =   6
            TabHeight       =   520
            TabCaption(0)   =   "Datos Generales"
            TabPicture(0)   =   "OpeTra_frm_837.frx":11AE
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "SSPanel3"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Inmueble"
            TabPicture(1)   =   "OpeTra_frm_837.frx":11CA
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "Label20"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "Label19"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "Label18"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "Label17"
            Tab(1).Control(3).Enabled=   0   'False
            Tab(1).Control(4)=   "Label16"
            Tab(1).Control(4).Enabled=   0   'False
            Tab(1).Control(5)=   "Label15"
            Tab(1).Control(5).Enabled=   0   'False
            Tab(1).Control(6)=   "Label23"
            Tab(1).Control(6).Enabled=   0   'False
            Tab(1).Control(7)=   "Label22"
            Tab(1).Control(7).Enabled=   0   'False
            Tab(1).Control(8)=   "SSPanel20"
            Tab(1).Control(8).Enabled=   0   'False
            Tab(1).Control(9)=   "ipp_ValRea_Inm"
            Tab(1).Control(9).Enabled=   0   'False
            Tab(1).Control(10)=   "ipp_ValCom_Inm"
            Tab(1).Control(10).Enabled=   0   'False
            Tab(1).Control(11)=   "ipp_ValACo_Inm"
            Tab(1).Control(11).Enabled=   0   'False
            Tab(1).Control(12)=   "ipp_ValEdi_Inm"
            Tab(1).Control(12).Enabled=   0   'False
            Tab(1).Control(13)=   "ipp_ValTer_Inm"
            Tab(1).Control(13).Enabled=   0   'False
            Tab(1).Control(14)=   "ipp_SumAse_Inm"
            Tab(1).Control(14).Enabled=   0   'False
            Tab(1).Control(15)=   "ipp_AreCon_Inm"
            Tab(1).Control(15).Enabled=   0   'False
            Tab(1).Control(16)=   "ipp_AreTer_Inm"
            Tab(1).Control(16).Enabled=   0   'False
            Tab(1).Control(17)=   "SSPanel5"
            Tab(1).Control(17).Enabled=   0   'False
            Tab(1).ControlCount=   18
            TabCaption(2)   =   "Estacionamiento 1"
            TabPicture(2)   =   "OpeTra_frm_837.frx":11E6
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Label30"
            Tab(2).Control(1)=   "Label29"
            Tab(2).Control(2)=   "Label28"
            Tab(2).Control(3)=   "Label27"
            Tab(2).Control(4)=   "Label26"
            Tab(2).Control(5)=   "Label25"
            Tab(2).Control(6)=   "Label24"
            Tab(2).Control(7)=   "Label10"
            Tab(2).Control(8)=   "Label9"
            Tab(2).Control(9)=   "SSPanel10"
            Tab(2).Control(10)=   "ipp_ValRea_Es1"
            Tab(2).Control(11)=   "ipp_ValCom_Es1"
            Tab(2).Control(12)=   "ipp_ValACo_Es1"
            Tab(2).Control(13)=   "ipp_ValEdi_Es1"
            Tab(2).Control(14)=   "ipp_ValTer_Es1"
            Tab(2).Control(15)=   "ipp_SumAse_Es1"
            Tab(2).Control(16)=   "ipp_AreCon_Es1"
            Tab(2).Control(17)=   "ipp_AreTer_Es1"
            Tab(2).Control(18)=   "SSPanel9"
            Tab(2).Control(19)=   "cmb_FlgEst_Es1"
            Tab(2).ControlCount=   20
            TabCaption(3)   =   "Estacionamiento 2"
            TabPicture(3)   =   "OpeTra_frm_837.frx":1202
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "Label31"
            Tab(3).Control(1)=   "Label39"
            Tab(3).Control(2)=   "Label38"
            Tab(3).Control(3)=   "Label37"
            Tab(3).Control(4)=   "Label36"
            Tab(3).Control(5)=   "Label35"
            Tab(3).Control(6)=   "Label34"
            Tab(3).Control(7)=   "Label33"
            Tab(3).Control(8)=   "Label32"
            Tab(3).Control(9)=   "SSPanel12"
            Tab(3).Control(10)=   "ipp_ValRea_Es2"
            Tab(3).Control(11)=   "ipp_ValCom_Es2"
            Tab(3).Control(12)=   "ipp_ValACo_Es2"
            Tab(3).Control(13)=   "ipp_ValEdi_Es2"
            Tab(3).Control(14)=   "ipp_ValTer_Es2"
            Tab(3).Control(15)=   "ipp_SumAse_Es2"
            Tab(3).Control(16)=   "ipp_AreCon_Es2"
            Tab(3).Control(17)=   "ipp_AreTer_Es2"
            Tab(3).Control(18)=   "SSPanel11"
            Tab(3).Control(19)=   "cmb_FlgEst_Es2"
            Tab(3).ControlCount=   20
            TabCaption(4)   =   "Depósito 1"
            TabPicture(4)   =   "OpeTra_frm_837.frx":121E
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "Label48"
            Tab(4).Control(1)=   "Label47"
            Tab(4).Control(2)=   "Label46"
            Tab(4).Control(3)=   "Label45"
            Tab(4).Control(4)=   "Label44"
            Tab(4).Control(5)=   "Label43"
            Tab(4).Control(6)=   "Label42"
            Tab(4).Control(7)=   "Label41"
            Tab(4).Control(8)=   "Label40"
            Tab(4).Control(9)=   "SSPanel14"
            Tab(4).Control(10)=   "ipp_ValRea_Dep1"
            Tab(4).Control(11)=   "ipp_ValCom_Dep1"
            Tab(4).Control(12)=   "ipp_ValACo_Dep1"
            Tab(4).Control(13)=   "ipp_ValEdi_Dep1"
            Tab(4).Control(14)=   "ipp_ValTer_Dep1"
            Tab(4).Control(15)=   "ipp_SumAse_Dep1"
            Tab(4).Control(16)=   "ipp_AreCon_Dep1"
            Tab(4).Control(17)=   "ipp_AreTer_Dep1"
            Tab(4).Control(18)=   "SSPanel13"
            Tab(4).Control(19)=   "cmb_FlgEst_Dep1"
            Tab(4).ControlCount=   20
            TabCaption(5)   =   "Depósito 2"
            TabPicture(5)   =   "OpeTra_frm_837.frx":123A
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "Label65"
            Tab(5).Control(1)=   "Label64"
            Tab(5).Control(2)=   "Label63"
            Tab(5).Control(3)=   "Label21"
            Tab(5).Control(4)=   "Label7"
            Tab(5).Control(5)=   "Label6"
            Tab(5).Control(6)=   "Label3"
            Tab(5).Control(7)=   "Label2"
            Tab(5).Control(8)=   "Label1"
            Tab(5).Control(9)=   "SSPanel17"
            Tab(5).Control(10)=   "SSPanel18"
            Tab(5).Control(11)=   "ipp_ValRea_Dep2"
            Tab(5).Control(12)=   "ipp_ValCom_Dep2"
            Tab(5).Control(13)=   "ipp_ValACo_Dep2"
            Tab(5).Control(14)=   "ipp_ValEdi_Dep2"
            Tab(5).Control(15)=   "ipp_ValTer_Dep2"
            Tab(5).Control(16)=   "ipp_SumAse_Dep2"
            Tab(5).Control(17)=   "ipp_AreCon_Dep2"
            Tab(5).Control(18)=   "ipp_AreTer_Dep2"
            Tab(5).Control(19)=   "cmb_FlgEst_Dep2"
            Tab(5).ControlCount=   20
            Begin VB.ComboBox cmb_FlgEst_Es1 
               Height          =   315
               ItemData        =   "OpeTra_frm_837.frx":1256
               Left            =   -73170
               List            =   "OpeTra_frm_837.frx":1258
               Style           =   2  'Dropdown List
               TabIndex        =   21
               Top             =   660
               Width           =   975
            End
            Begin VB.ComboBox cmb_FlgEst_Es2 
               Height          =   315
               Left            =   -73170
               Style           =   2  'Dropdown List
               TabIndex        =   30
               Top             =   660
               Width           =   975
            End
            Begin VB.ComboBox cmb_FlgEst_Dep1 
               Height          =   315
               Left            =   -73170
               Style           =   2  'Dropdown List
               TabIndex        =   39
               Top             =   660
               Width           =   975
            End
            Begin VB.ComboBox cmb_FlgEst_Dep2 
               Height          =   315
               Left            =   -73170
               Style           =   2  'Dropdown List
               TabIndex        =   48
               Top             =   660
               Width           =   975
            End
            Begin Threed.SSPanel SSPanel3 
               Height          =   1905
               Left            =   -74940
               TabIndex        =   103
               Top             =   330
               Width           =   13665
               _Version        =   65536
               _ExtentX        =   24104
               _ExtentY        =   3360
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
               Begin VB.ComboBox cmb_TieAzo 
                  Height          =   315
                  Left            =   12600
                  Style           =   2  'Dropdown List
                  TabIndex        =   7
                  Top             =   810
                  Width           =   975
               End
               Begin VB.ComboBox cmb_MatCon 
                  Height          =   315
                  Left            =   1860
                  Style           =   2  'Dropdown List
                  TabIndex        =   10
                  Top             =   1470
                  Width           =   5865
               End
               Begin VB.ComboBox cmb_UsoInm 
                  Height          =   315
                  Left            =   9360
                  Style           =   2  'Dropdown List
                  TabIndex        =   9
                  Top             =   1140
                  Width           =   4215
               End
               Begin VB.ComboBox cmb_TipInm 
                  Height          =   315
                  Left            =   1860
                  Style           =   2  'Dropdown List
                  TabIndex        =   8
                  Top             =   1140
                  Width           =   3315
               End
               Begin VB.ComboBox cmb_TipMon 
                  Height          =   315
                  Left            =   9360
                  Style           =   2  'Dropdown List
                  TabIndex        =   11
                  Top             =   1470
                  Width           =   2235
               End
               Begin VB.TextBox txt_NumInf 
                  Height          =   315
                  Left            =   1860
                  MaxLength       =   25
                  TabIndex        =   2
                  Text            =   "Text1"
                  Top             =   480
                  Width           =   3315
               End
               Begin VB.ComboBox cmb_EmpPer 
                  Height          =   315
                  Left            =   1860
                  Style           =   2  'Dropdown List
                  TabIndex        =   0
                  Top             =   150
                  Width           =   5865
               End
               Begin VB.ComboBox cmb_PerTas 
                  Height          =   315
                  Left            =   9360
                  Style           =   2  'Dropdown List
                  TabIndex        =   1
                  Top             =   150
                  Width           =   4215
               End
               Begin EditLib.fpDateTime ipp_FecEva 
                  Height          =   315
                  Left            =   9360
                  TabIndex        =   3
                  Top             =   480
                  Width           =   1335
                  _Version        =   196608
                  _ExtentX        =   2355
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
                  Text            =   "24/01/2014"
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
               Begin EditLib.fpDoubleSingle ipp_TipCam 
                  Height          =   315
                  Left            =   12720
                  TabIndex        =   12
                  Top             =   1470
                  Width           =   825
                  _Version        =   196608
                  _ExtentX        =   1455
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
                  Text            =   "0.0000"
                  DecimalPlaces   =   4
                  DecimalPoint    =   "."
                  FixedPoint      =   -1  'True
                  LeadZero        =   0
                  MaxValue        =   "9000000000"
                  MinValue        =   "-9000000000"
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
               Begin EditLib.fpDoubleSingle ipp_AnoCon 
                  Height          =   315
                  Left            =   1860
                  TabIndex        =   4
                  Top             =   810
                  Width           =   1005
                  _Version        =   196608
                  _ExtentX        =   1773
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
                  ButtonStyle     =   1
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
                  Text            =   "0"
                  DecimalPlaces   =   0
                  DecimalPoint    =   "."
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "9999"
                  MinValue        =   "1900"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   0   'False
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
               Begin EditLib.fpDoubleSingle ipp_NumPis 
                  Height          =   315
                  Left            =   4350
                  TabIndex        =   5
                  Top             =   810
                  Width           =   825
                  _Version        =   196608
                  _ExtentX        =   1455
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
                  ButtonStyle     =   1
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
                  Text            =   "0"
                  DecimalPlaces   =   0
                  DecimalPoint    =   "."
                  FixedPoint      =   -1  'True
                  LeadZero        =   0
                  MaxValue        =   "99"
                  MinValue        =   "1"
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
               Begin EditLib.fpDoubleSingle ipp_NumSot 
                  Height          =   315
                  Left            =   9360
                  TabIndex        =   6
                  Top             =   810
                  Width           =   765
                  _Version        =   196608
                  _ExtentX        =   1349
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
                  ButtonStyle     =   1
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
                  Text            =   "0"
                  DecimalPlaces   =   0
                  DecimalPoint    =   "."
                  FixedPoint      =   -1  'True
                  LeadZero        =   0
                  MaxValue        =   "99"
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
               Begin VB.Label Label67 
                  AutoSize        =   -1  'True
                  Caption         =   "Tiene Azotea"
                  Height          =   195
                  Left            =   11370
                  TabIndex        =   116
                  Top             =   870
                  Width           =   945
               End
               Begin VB.Label Label62 
                  AutoSize        =   -1  'True
                  Caption         =   "Material Construcción:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   115
                  Top             =   1530
                  Width           =   1575
               End
               Begin VB.Label Label61 
                  AutoSize        =   -1  'True
                  Caption         =   "Uso Inmueble:"
                  Height          =   195
                  Left            =   8070
                  TabIndex        =   114
                  Top             =   1200
                  Width           =   1020
               End
               Begin VB.Label Label60 
                  AutoSize        =   -1  'True
                  Caption         =   "Tipo Inmueble:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   113
                  Top             =   1200
                  Width           =   1050
               End
               Begin VB.Label Label59 
                  AutoSize        =   -1  'True
                  Caption         =   "Nro. Sótanos:"
                  Height          =   195
                  Left            =   8070
                  TabIndex        =   112
                  Top             =   870
                  Width           =   975
               End
               Begin VB.Label Label58 
                  AutoSize        =   -1  'True
                  Caption         =   "Nro. Pisos:"
                  Height          =   195
                  Left            =   3390
                  TabIndex        =   111
                  Top             =   870
                  Width           =   765
               End
               Begin VB.Label Label57 
                  AutoSize        =   -1  'True
                  Caption         =   "Año Construcción:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   110
                  Top             =   870
                  Width           =   1305
               End
               Begin VB.Label Label14 
                  AutoSize        =   -1  'True
                  Caption         =   "T. Cambio:"
                  Height          =   195
                  Left            =   11730
                  TabIndex        =   109
                  Top             =   1530
                  Width           =   765
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  Caption         =   "Moneda:"
                  Height          =   195
                  Left            =   8070
                  TabIndex        =   108
                  Top             =   1530
                  Width           =   630
               End
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  Caption         =   "F. Evaluación:"
                  Height          =   195
                  Left            =   8070
                  TabIndex        =   107
                  Top             =   540
                  Width           =   1020
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  Caption         =   "Número Informe:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   106
                  Top             =   540
                  Width           =   1170
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  Caption         =   "Empresa Peritaje:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   105
                  Top             =   210
                  Width           =   1230
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  Caption         =   "Perito Tasador:"
                  Height          =   195
                  Left            =   8070
                  TabIndex        =   104
                  Top             =   210
                  Width           =   1080
               End
            End
            Begin EditLib.fpDoubleSingle ipp_AreTer_Dep2 
               Height          =   315
               Left            =   -68670
               TabIndex        =   49
               Top             =   660
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_AreCon_Dep2 
               Height          =   315
               Left            =   -63810
               TabIndex        =   50
               Top             =   660
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_SumAse_Dep2 
               Height          =   315
               Left            =   -73170
               TabIndex        =   51
               Top             =   1290
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValTer_Dep2 
               Height          =   315
               Left            =   -73170
               TabIndex        =   54
               Top             =   1620
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValEdi_Dep2 
               Height          =   315
               Left            =   -68670
               TabIndex        =   55
               Top             =   1620
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValACo_Dep2 
               Height          =   315
               Left            =   -63810
               TabIndex        =   56
               Top             =   1620
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValCom_Dep2 
               Height          =   315
               Left            =   -68670
               TabIndex        =   52
               Top             =   1290
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValRea_Dep2 
               Height          =   315
               Left            =   -63810
               TabIndex        =   53
               Top             =   1290
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin Threed.SSPanel SSPanel18 
               Height          =   60
               Left            =   -75000
               TabIndex        =   117
               Top             =   1110
               Width           =   13725
               _Version        =   65536
               _ExtentX        =   24209
               _ExtentY        =   106
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
            End
            Begin Threed.SSPanel SSPanel17 
               Height          =   60
               Left            =   -75000
               TabIndex        =   127
               Top             =   480
               Width           =   13725
               _Version        =   65536
               _ExtentX        =   24209
               _ExtentY        =   106
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
            End
            Begin Threed.SSPanel SSPanel13 
               Height          =   60
               Left            =   -75000
               TabIndex        =   128
               Top             =   480
               Width           =   13725
               _Version        =   65536
               _ExtentX        =   24209
               _ExtentY        =   106
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
            End
            Begin EditLib.fpDoubleSingle ipp_AreTer_Dep1 
               Height          =   315
               Left            =   -68670
               TabIndex        =   40
               Top             =   660
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_AreCon_Dep1 
               Height          =   315
               Left            =   -63810
               TabIndex        =   41
               Top             =   660
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_SumAse_Dep1 
               Height          =   315
               Left            =   -73170
               TabIndex        =   42
               Top             =   1290
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValTer_Dep1 
               Height          =   315
               Left            =   -73170
               TabIndex        =   45
               Top             =   1620
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValEdi_Dep1 
               Height          =   315
               Left            =   -68670
               TabIndex        =   46
               Top             =   1620
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValACo_Dep1 
               Height          =   315
               Left            =   -63810
               TabIndex        =   47
               Top             =   1620
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValCom_Dep1 
               Height          =   315
               Left            =   -68670
               TabIndex        =   43
               Top             =   1290
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValRea_Dep1 
               Height          =   315
               Left            =   -63810
               TabIndex        =   44
               Top             =   1290
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin Threed.SSPanel SSPanel14 
               Height          =   60
               Left            =   -75000
               TabIndex        =   129
               Top             =   1110
               Width           =   13725
               _Version        =   65536
               _ExtentX        =   24209
               _ExtentY        =   106
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
            End
            Begin Threed.SSPanel SSPanel11 
               Height          =   60
               Left            =   -75000
               TabIndex        =   139
               Top             =   480
               Width           =   13725
               _Version        =   65536
               _ExtentX        =   24209
               _ExtentY        =   106
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
            End
            Begin EditLib.fpDoubleSingle ipp_AreTer_Es2 
               Height          =   315
               Left            =   -68670
               TabIndex        =   31
               Top             =   660
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_AreCon_Es2 
               Height          =   315
               Left            =   -63810
               TabIndex        =   32
               Top             =   660
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_SumAse_Es2 
               Height          =   315
               Left            =   -73170
               TabIndex        =   33
               Top             =   1290
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValTer_Es2 
               Height          =   315
               Left            =   -73170
               TabIndex        =   36
               Top             =   1620
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValEdi_Es2 
               Height          =   315
               Left            =   -68670
               TabIndex        =   37
               Top             =   1620
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValACo_Es2 
               Height          =   315
               Left            =   -63810
               TabIndex        =   38
               Top             =   1620
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValCom_Es2 
               Height          =   315
               Left            =   -68670
               TabIndex        =   34
               Top             =   1290
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValRea_Es2 
               Height          =   315
               Left            =   -63810
               TabIndex        =   35
               Top             =   1290
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin Threed.SSPanel SSPanel12 
               Height          =   60
               Left            =   -75000
               TabIndex        =   140
               Top             =   1110
               Width           =   13725
               _Version        =   65536
               _ExtentX        =   24209
               _ExtentY        =   106
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
            End
            Begin Threed.SSPanel SSPanel9 
               Height          =   60
               Left            =   -75000
               TabIndex        =   150
               Top             =   480
               Width           =   13725
               _Version        =   65536
               _ExtentX        =   24209
               _ExtentY        =   106
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
            End
            Begin EditLib.fpDoubleSingle ipp_AreTer_Es1 
               Height          =   315
               Left            =   -68670
               TabIndex        =   22
               Top             =   660
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_AreCon_Es1 
               Height          =   315
               Left            =   -63810
               TabIndex        =   23
               Top             =   660
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_SumAse_Es1 
               Height          =   315
               Left            =   -73170
               TabIndex        =   24
               Top             =   1290
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValTer_Es1 
               Height          =   315
               Left            =   -73170
               TabIndex        =   27
               Top             =   1620
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValEdi_Es1 
               Height          =   315
               Left            =   -68670
               TabIndex        =   28
               Top             =   1620
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValACo_Es1 
               Height          =   315
               Left            =   -63810
               TabIndex        =   29
               Top             =   1620
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValCom_Es1 
               Height          =   315
               Left            =   -68670
               TabIndex        =   25
               Top             =   1290
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValRea_Es1 
               Height          =   315
               Left            =   -63810
               TabIndex        =   26
               Top             =   1290
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin Threed.SSPanel SSPanel10 
               Height          =   60
               Left            =   -75000
               TabIndex        =   151
               Top             =   1110
               Width           =   13725
               _Version        =   65536
               _ExtentX        =   24209
               _ExtentY        =   106
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
            End
            Begin Threed.SSPanel SSPanel5 
               Height          =   60
               Left            =   30
               TabIndex        =   161
               Top             =   1110
               Width           =   13695
               _Version        =   65536
               _ExtentX        =   24156
               _ExtentY        =   106
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
            End
            Begin EditLib.fpDoubleSingle ipp_AreTer_Inm 
               Height          =   315
               Left            =   1830
               TabIndex        =   13
               Top             =   660
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_AreCon_Inm 
               Height          =   315
               Left            =   6330
               TabIndex        =   14
               Top             =   660
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_SumAse_Inm 
               Height          =   315
               Left            =   1830
               TabIndex        =   15
               Top             =   1290
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValTer_Inm 
               Height          =   315
               Left            =   1830
               TabIndex        =   18
               Top             =   1620
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValEdi_Inm 
               Height          =   315
               Left            =   6330
               TabIndex        =   19
               Top             =   1620
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValACo_Inm 
               Height          =   315
               Left            =   11190
               TabIndex        =   20
               Top             =   1620
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValCom_Inm 
               Height          =   315
               Left            =   6330
               TabIndex        =   16
               Top             =   1290
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin EditLib.fpDoubleSingle ipp_ValRea_Inm 
               Height          =   315
               Left            =   11190
               TabIndex        =   17
               Top             =   1290
               Width           =   1335
               _Version        =   196608
               _ExtentX        =   2355
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
               MinValue        =   "-9000000000"
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
            Begin Threed.SSPanel SSPanel20 
               Height          =   60
               Left            =   0
               TabIndex        =   162
               Top             =   480
               Width           =   13725
               _Version        =   65536
               _ExtentX        =   24209
               _ExtentY        =   106
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
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "Area Terreno:"
               Height          =   195
               Left            =   90
               TabIndex        =   170
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               Caption         =   "Area Construcción:"
               Height          =   195
               Left            =   4530
               TabIndex        =   169
               Top             =   720
               Width           =   1350
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Suma Asegurada:"
               Height          =   195
               Left            =   90
               TabIndex        =   168
               Top             =   1350
               Width           =   1260
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Valor Terreno:"
               Height          =   195
               Left            =   90
               TabIndex        =   167
               Top             =   1680
               Width           =   1005
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Valor Edificación:"
               Height          =   195
               Left            =   4530
               TabIndex        =   166
               Top             =   1680
               Width           =   1230
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Valor Areas Comunes:"
               Height          =   195
               Left            =   9120
               TabIndex        =   165
               Top             =   1680
               Width           =   1560
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Valor Comercial:"
               Height          =   195
               Left            =   4530
               TabIndex        =   164
               Top             =   1350
               Width           =   1140
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Valor Realización:"
               Height          =   195
               Left            =   9120
               TabIndex        =   163
               Top             =   1350
               Width           =   1275
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Estacionamiento 1:"
               Height          =   195
               Left            =   -74910
               TabIndex        =   160
               Top             =   720
               Width           =   1350
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Area Terreno:"
               Height          =   195
               Left            =   -70470
               TabIndex        =   159
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "Area Construcción:"
               Height          =   195
               Left            =   -65880
               TabIndex        =   158
               Top             =   720
               Width           =   1350
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               Caption         =   "Suma Asegurada:"
               Height          =   195
               Left            =   -74910
               TabIndex        =   157
               Top             =   1350
               Width           =   1260
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               Caption         =   "Valor Terreno:"
               Height          =   195
               Left            =   -74910
               TabIndex        =   156
               Top             =   1680
               Width           =   1005
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               Caption         =   "Valor Edificación:"
               Height          =   195
               Left            =   -70470
               TabIndex        =   155
               Top             =   1680
               Width           =   1230
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               Caption         =   "Valor Areas Comunes:"
               Height          =   195
               Left            =   -65880
               TabIndex        =   154
               Top             =   1680
               Width           =   1560
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               Caption         =   "Valor Comercial:"
               Height          =   195
               Left            =   -70470
               TabIndex        =   153
               Top             =   1350
               Width           =   1140
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               Caption         =   "Valor Realización:"
               Height          =   195
               Left            =   -65880
               TabIndex        =   152
               Top             =   1350
               Width           =   1275
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               Caption         =   "Area Terreno:"
               Height          =   195
               Left            =   -70470
               TabIndex        =   149
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label33 
               AutoSize        =   -1  'True
               Caption         =   "Area Construcción:"
               Height          =   195
               Left            =   -65880
               TabIndex        =   148
               Top             =   720
               Width           =   1350
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               Caption         =   "Suma Asegurada:"
               Height          =   195
               Left            =   -74910
               TabIndex        =   147
               Top             =   1350
               Width           =   1260
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               Caption         =   "Valor Terreno:"
               Height          =   195
               Left            =   -74910
               TabIndex        =   146
               Top             =   1680
               Width           =   1005
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               Caption         =   "Valor Edificación:"
               Height          =   195
               Left            =   -70470
               TabIndex        =   145
               Top             =   1680
               Width           =   1230
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               Caption         =   "Valor Areas Comunes:"
               Height          =   195
               Left            =   -65880
               TabIndex        =   144
               Top             =   1680
               Width           =   1560
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               Caption         =   "Valor Comercial:"
               Height          =   195
               Left            =   -70470
               TabIndex        =   143
               Top             =   1350
               Width           =   1140
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               Caption         =   "Valor Realización:"
               Height          =   195
               Left            =   -65880
               TabIndex        =   142
               Top             =   1350
               Width           =   1275
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "Estacionamiento 2:"
               Height          =   195
               Left            =   -74910
               TabIndex        =   141
               Top             =   720
               Width           =   1350
            End
            Begin VB.Label Label40 
               AutoSize        =   -1  'True
               Caption         =   "Depósito 1:"
               Height          =   195
               Left            =   -74910
               TabIndex        =   138
               Top             =   720
               Width           =   810
            End
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               Caption         =   "Area Terreno:"
               Height          =   195
               Left            =   -70470
               TabIndex        =   137
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label42 
               AutoSize        =   -1  'True
               Caption         =   "Area Construcción:"
               Height          =   195
               Left            =   -65880
               TabIndex        =   136
               Top             =   720
               Width           =   1350
            End
            Begin VB.Label Label43 
               AutoSize        =   -1  'True
               Caption         =   "Suma Asegurada:"
               Height          =   195
               Left            =   -74910
               TabIndex        =   135
               Top             =   1350
               Width           =   1260
            End
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               Caption         =   "Valor Terreno:"
               Height          =   195
               Left            =   -74910
               TabIndex        =   134
               Top             =   1680
               Width           =   1005
            End
            Begin VB.Label Label45 
               AutoSize        =   -1  'True
               Caption         =   "Valor Edificación:"
               Height          =   195
               Left            =   -70470
               TabIndex        =   133
               Top             =   1680
               Width           =   1230
            End
            Begin VB.Label Label46 
               AutoSize        =   -1  'True
               Caption         =   "Valor Areas Comunes:"
               Height          =   195
               Left            =   -65880
               TabIndex        =   132
               Top             =   1680
               Width           =   1560
            End
            Begin VB.Label Label47 
               AutoSize        =   -1  'True
               Caption         =   "Valor Comercial:"
               Height          =   195
               Left            =   -70470
               TabIndex        =   131
               Top             =   1350
               Width           =   1140
            End
            Begin VB.Label Label48 
               AutoSize        =   -1  'True
               Caption         =   "Valor Realización:"
               Height          =   195
               Left            =   -65880
               TabIndex        =   130
               Top             =   1350
               Width           =   1275
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor Realización:"
               Height          =   195
               Left            =   -65880
               TabIndex        =   126
               Top             =   1350
               Width           =   1275
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Valor Comercial:"
               Height          =   195
               Left            =   -70470
               TabIndex        =   125
               Top             =   1350
               Width           =   1140
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Valor Areas Comunes:"
               Height          =   195
               Left            =   -65880
               TabIndex        =   124
               Top             =   1680
               Width           =   1560
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Valor Edificación:"
               Height          =   195
               Left            =   -70470
               TabIndex        =   123
               Top             =   1680
               Width           =   1230
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Valor Terreno:"
               Height          =   195
               Left            =   -74910
               TabIndex        =   122
               Top             =   1680
               Width           =   1005
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "Suma Asegurada:"
               Height          =   195
               Left            =   -74910
               TabIndex        =   121
               Top             =   1350
               Width           =   1260
            End
            Begin VB.Label Label63 
               AutoSize        =   -1  'True
               Caption         =   "Area Construcción:"
               Height          =   195
               Left            =   -65880
               TabIndex        =   120
               Top             =   720
               Width           =   1350
            End
            Begin VB.Label Label64 
               AutoSize        =   -1  'True
               Caption         =   "Area Terreno:"
               Height          =   195
               Left            =   -70470
               TabIndex        =   119
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label65 
               AutoSize        =   -1  'True
               Caption         =   "Depósito 2:"
               Height          =   195
               Left            =   -74910
               TabIndex        =   118
               Top             =   720
               Width           =   810
            End
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   81
         Top             =   30
         Width           =   13905
         _Version        =   65536
         _ExtentX        =   24527
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
            TabIndex        =   82
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
            TabIndex        =   83
            Top             =   330
            Width           =   4215
            _Version        =   65536
            _ExtentX        =   7435
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Techo Propio - Registro de Tasación"
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
            Picture         =   "OpeTra_frm_837.frx":125A
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel19 
         Height          =   885
         Left            =   60
         TabIndex        =   84
         Top             =   1440
         Width           =   13905
         _Version        =   65536
         _ExtentX        =   24527
         _ExtentY        =   1561
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
            TabIndex        =   85
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
            TabIndex        =   86
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
            TabIndex        =   87
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
            Left            =   9360
            TabIndex        =   88
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
         Begin VB.Label lbl_TipEmp 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Empresa:"
            Height          =   195
            Left            =   7770
            TabIndex        =   92
            Top             =   510
            Width           =   1020
         End
         Begin VB.Label lbl_RazSoc 
            AutoSize        =   -1  'True
            Caption         =   "Razón Social:"
            Height          =   195
            Left            =   120
            TabIndex        =   91
            Top             =   510
            Width           =   990
         End
         Begin VB.Label lbl_NumDoc 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Documento:"
            Height          =   195
            Left            =   7770
            TabIndex        =   90
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label lbl_TipDoc 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento:"
            Height          =   195
            Left            =   120
            TabIndex        =   89
            Top             =   180
            Width           =   1230
         End
      End
      Begin Threed.SSPanel SSPanel21 
         Height          =   2685
         Left            =   60
         TabIndex        =   93
         Top             =   2370
         Width           =   13905
         _Version        =   65536
         _ExtentX        =   24527
         _ExtentY        =   4736
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
            Height          =   2175
            Left            =   90
            TabIndex        =   94
            Top             =   450
            Width           =   13710
            _ExtentX        =   24183
            _ExtentY        =   3836
            _Version        =   393216
            Rows            =   21
            Cols            =   5
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin Threed.SSPanel pnl_Tit_TipGar 
            Height          =   285
            Left            =   90
            TabIndex        =   95
            Top             =   150
            Width           =   2145
            _Version        =   65536
            _ExtentX        =   3784
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Número Informe"
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
         Begin Threed.SSPanel pnl_Tit_NumRef 
            Height          =   285
            Left            =   8730
            TabIndex        =   96
            Top             =   150
            Width           =   4710
            _Version        =   65536
            _ExtentX        =   8308
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Perito Tasador"
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
         Begin Threed.SSPanel pnl_Tit_FecEmi 
            Height          =   285
            Left            =   2220
            TabIndex        =   97
            Top             =   150
            Width           =   1830
            _Version        =   65536
            _ExtentX        =   3228
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Evaluación"
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
         Begin Threed.SSPanel pnl_Tit_TipMon 
            Height          =   285
            Left            =   4020
            TabIndex        =   98
            Top             =   150
            Width           =   4740
            _Version        =   65536
            _ExtentX        =   8361
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Empresa Peritaje"
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
   End
End
Attribute VB_Name = "frm_Ges_TecPro_15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_EmpPer()   As moddat_tpo_Genera
Dim l_arr_PerTas()   As moddat_tpo_Genera
Dim l_int_CodTas     As Integer

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte(cmb_EmpPer, l_arr_EmpPer, 1, "507")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipInm, 1, "221")
   Call moddat_gs_Carga_LisIte_Combo(cmb_UsoInm, 1, "222")
   Call moddat_gs_Carga_LisIte_Combo(cmb_MatCon, 1, "223")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipMon, 1, "204")
   Call moddat_gs_Carga_LisIte_Combo(cmb_FlgEst_Es1, 1, "214")
   Call moddat_gs_Carga_LisIte_Combo(cmb_FlgEst_Es2, 1, "214")
   Call moddat_gs_Carga_LisIte_Combo(cmb_FlgEst_Dep1, 1, "214")
   Call moddat_gs_Carga_LisIte_Combo(cmb_FlgEst_Dep2, 1, "214")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TieAzo, 1, "214")
   
   pnl_TipDoc.Caption = moddat_gf_Consulta_ParDes("118", moddat_g_int_TipDoc)
   pnl_NroDoc.Caption = moddat_g_str_NumDoc
   pnl_RazSoc.Caption = moddat_g_str_NomCli
   pnl_TipEmp.Caption = moddat_g_str_Descri
  
   grd_Listad.ColWidth(0) = 2150
   grd_Listad.ColWidth(1) = 1810
   grd_Listad.ColWidth(2) = 4700
   grd_Listad.ColWidth(3) = 4700
   grd_Listad.ColWidth(4) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
   
   'Inicializando Rejilla
   Call gs_LimpiaGrid(grd_Listad)
   
End Sub

'Private Sub fs_Buscar_Credito()
'Dim r_int_TipGar     As Integer
'
'
'   'Call modmip_gs_DatNumOpe(moddat_g_str_NumOpe, grd_Listad, r_int_TipGar)
'
'   If Not (r_int_TipGar = 1 Or r_int_TipGar = 2) Then
'      MsgBox "El tipo de garantia debe ser HIPOTECA.", vbExclamation, modgen_g_str_NomPlt
'      cmd_Grabar.Enabled = False
'      cmd_HisTas.Enabled = False
'   End If
'
'End Sub
Private Sub fs_Activa(ByVal p_Activa As Integer)
   
   cmb_EmpPer.Enabled = p_Activa
   cmb_PerTas.Enabled = p_Activa
   txt_NumInf.Enabled = p_Activa
   ipp_FecEva.Enabled = p_Activa
   ipp_AnoCon.Enabled = p_Activa
   ipp_NumPis.Enabled = p_Activa
   ipp_NumSot.Enabled = p_Activa
   cmb_TieAzo.Enabled = p_Activa
   cmb_TipInm.Enabled = p_Activa
   cmb_UsoInm.Enabled = p_Activa
   cmb_MatCon.Enabled = p_Activa
   cmb_TipMon.Enabled = p_Activa
   ipp_TipCam.Enabled = p_Activa
'   tab_Genera.Enabled = p_Activa
   ipp_AreTer_Inm.Enabled = p_Activa
   ipp_AreCon_Inm.Enabled = p_Activa
   ipp_SumAse_Inm.Enabled = p_Activa
   ipp_ValCom_Inm.Enabled = p_Activa
   ipp_ValRea_Inm.Enabled = p_Activa
   ipp_ValTer_Inm.Enabled = p_Activa
   ipp_ValEdi_Inm.Enabled = p_Activa
   ipp_ValACo_Inm.Enabled = p_Activa
   
   cmb_FlgEst_Es1.Enabled = p_Activa
   ipp_AreTer_Es1.Enabled = p_Activa
   ipp_AreCon_Es1.Enabled = p_Activa
   ipp_SumAse_Es1.Enabled = p_Activa
   ipp_ValCom_Es1.Enabled = p_Activa
   ipp_ValRea_Es1.Enabled = p_Activa
   ipp_ValTer_Es1.Enabled = p_Activa
   ipp_ValEdi_Es1.Enabled = p_Activa
   ipp_ValACo_Es1.Enabled = p_Activa
   
   cmb_FlgEst_Es2.Enabled = p_Activa
   ipp_AreTer_Es2.Enabled = p_Activa
   ipp_AreCon_Es2.Enabled = p_Activa
   ipp_SumAse_Es2.Enabled = p_Activa
   ipp_ValCom_Es2.Enabled = p_Activa
   ipp_ValRea_Es2.Enabled = p_Activa
   ipp_ValTer_Es2.Enabled = p_Activa
   ipp_ValEdi_Es2.Enabled = p_Activa
   ipp_ValACo_Es2.Enabled = p_Activa
   
   cmb_FlgEst_Dep1.Enabled = p_Activa
   ipp_AreTer_Dep1.Enabled = p_Activa
   ipp_AreCon_Dep1.Enabled = p_Activa
   ipp_SumAse_Dep1.Enabled = p_Activa
   ipp_ValCom_Dep1.Enabled = p_Activa
   ipp_ValRea_Dep1.Enabled = p_Activa
   ipp_ValTer_Dep1.Enabled = p_Activa
   ipp_ValEdi_Dep1.Enabled = p_Activa
   ipp_ValACo_Dep1.Enabled = p_Activa
   
   cmb_FlgEst_Dep2.Enabled = p_Activa
   ipp_AreTer_Dep2.Enabled = p_Activa
   ipp_AreCon_Dep2.Enabled = p_Activa
   ipp_SumAse_Dep2.Enabled = p_Activa
   ipp_ValCom_Dep2.Enabled = p_Activa
   ipp_ValRea_Dep2.Enabled = p_Activa
   ipp_ValTer_Dep2.Enabled = p_Activa
   ipp_ValEdi_Dep2.Enabled = p_Activa
   ipp_ValACo_Dep2.Enabled = p_Activa
   
   cmd_Agrega.Enabled = Not p_Activa
      
   If grd_Listad.Row < 0 Then
      cmd_Editar.Enabled = p_Activa
      'cmd_ExpExc.Enabled = p_Activa
   Else
      cmd_Editar.Enabled = Not p_Activa
      'cmd_ExpExc.Enabled = Not p_Activa
   End If
   cmd_Grabar.Enabled = p_Activa
   cmd_Cancel.Enabled = p_Activa
End Sub
Private Sub fs_Buscar()

Dim r_str_Cadena  As String
Dim r_str_CadAux  As String
Dim r_str_CadRef  As String
Dim r_dbl_MtoCfi  As Double

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT EVATAS_CODTAS    , EVATAS_CODEMP    , EVATAS_NOMPER    , EVATAS_CODPER    , EVATAS_NUMINF    , EVATAS_FECEVA   " ' , EVATAS_ANOCON    , "
'   g_str_Parame = g_str_Parame & "        EVATAS_NUMPIS    , EVATAS_NUMSOT    , EVATAS_FLGAZO    , EVATAS_TIPINM    , EVATAS_USOINM    , EVATAS_MATCON    , EVATAS_TIPMON    , EVATAS_TIPCAM    , EVATAS_TCAMPR    , "
'   g_str_Parame = g_str_Parame & "        EVATAS_ARETER_INM, EVATAS_ARECON_INM, EVATAS_SUMASE_INM, EVATAS_VALCOM_INM, EVATAS_VALREA_INM, EVATAS_VALTER_INM, EVATAS_VALEDI_INM, EVATAS_VALACO_INM, EVATAS_FLGEST_ES1, "
'   g_str_Parame = g_str_Parame & "        EVATAS_ARETER_ES1, EVATAS_ARECON_ES1, EVATAS_SUMASE_ES1, EVATAS_VALCOM_ES1, EVATAS_VALREA_ES1, EVATAS_VALTER_ES1, EVATAS_VALEDI_ES1, EVATAS_VALACO_ES1, EVATAS_FLGEST_ES2, "
'   g_str_Parame = g_str_Parame & "        EVATAS_ARETER_ES2, EVATAS_ARECON_ES2, EVATAS_SUMASE_ES2, EVATAS_VALCOM_ES2, EVATAS_VALREA_ES2, EVATAS_VALTER_ES2, EVATAS_VALEDI_ES2, EVATAS_VALACO_ES2, EVATAS_FLGEST_DE1, "
'   g_str_Parame = g_str_Parame & "        EVATAS_ARETER_DE1, EVATAS_ARECON_DE1, EVATAS_SUMASE_DE1, EVATAS_VALCOM_DE1, EVATAS_VALREA_DE1, EVATAS_VALTER_DE1, EVATAS_VALEDI_DE1, EVATAS_VALACO_DE1, EVATAS_FLGEST_DE2, "
'   g_str_Parame = g_str_Parame & "        EVATAS_ARETER_DE2, EVATAS_ARECON_DE2, EVATAS_SUMASE_DE2, EVATAS_VALCOM_DE2, EVATAS_VALREA_DE2, EVATAS_VALTER_DE2, EVATAS_VALEDI_DE2, EVATAS_VALACO_DE2 "
   g_str_Parame = g_str_Parame & "   FROM TPR_EVATAS "
   g_str_Parame = g_str_Parame & "  WHERE EVATAS_TIPDOC = " & CStr(moddat_g_int_TipDoc) & ""
   g_str_Parame = g_str_Parame & "    AND EVATAS_NUMDOC = '" & CStr(moddat_g_str_NumDoc) & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   Call gs_LimpiaGrid(grd_Listad)
    
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     grd_Listad.Redraw = True
     Exit Sub
   End If
     
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      cmd_Editar.Enabled = True
      g_rst_Princi.MoveFirst
     
      Do While Not g_rst_Princi.EOF
         
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         If Not IsNull(g_rst_Princi!EVATAS_NUMINF) Then
            grd_Listad.Text = Trim(g_rst_Princi!EVATAS_NUMINF)
         Else
            grd_Listad.Text = ""
         End If
         
         grd_Listad.Col = 1
         If Not IsNull(g_rst_Princi!EVATAS_FECEVA) Then
            grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!EVATAS_FECEVA)
         Else
            grd_Listad.Text = ""
         End If
         
         grd_Listad.Col = 2
         If Not IsNull(g_rst_Princi!EVATAS_CODEMP) Then
            grd_Listad.Text = Trim(moddat_gf_Consulta_ParDes("507", g_rst_Princi!EVATAS_CODEMP))
         Else
            grd_Listad.Text = ""
         End If
         
         grd_Listad.Col = 3
         If Not IsNull(g_rst_Princi!EVATAS_NOMPER) Then
            grd_Listad.Text = Trim(g_rst_Princi!EVATAS_NOMPER)
         Else
            grd_Listad.Text = ""
         End If
         
         grd_Listad.Col = 4
         grd_Listad.Text = g_rst_Princi!EVATAS_CODTAS
                  
         g_rst_Princi.MoveNext
      Loop
   End If
   
   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)
   
End Sub

Private Sub fs_Calcul()
   pnl_AreTer.Caption = Format(CDbl(ipp_AreTer_Inm.Text) + CDbl(ipp_AreTer_Es1.Text) + CDbl(ipp_AreTer_Es2.Text) + CDbl(ipp_AreTer_Dep1.Text) + CDbl(ipp_AreTer_Dep2.Text), "###,###,##0.00") & " "
   pnl_AreCon.Caption = Format(CDbl(ipp_AreCon_Inm.Text) + CDbl(ipp_AreCon_Es1.Text) + CDbl(ipp_AreCon_Es2.Text) + CDbl(ipp_AreCon_Dep1.Text) + CDbl(ipp_AreCon_Dep2.Text), "###,###,##0.00") & " "
   pnl_SumAse.Caption = Format(CDbl(ipp_SumAse_Inm.Text) + CDbl(ipp_SumAse_Es1.Text) + CDbl(ipp_SumAse_Es2.Text) + CDbl(ipp_SumAse_Dep1.Text) + CDbl(ipp_SumAse_Dep2.Text), "###,###,##0.00") & " "
   pnl_ValCom.Caption = Format(CDbl(ipp_ValCom_Inm.Text) + CDbl(ipp_ValCom_Es1.Text) + CDbl(ipp_ValCom_Es2.Text) + CDbl(ipp_ValCom_Dep1.Text) + CDbl(ipp_ValCom_Dep2.Text), "###,###,##0.00") & " "
   pnl_ValRea.Caption = Format(CDbl(ipp_ValRea_Inm.Text) + CDbl(ipp_ValRea_Es1.Text) + CDbl(ipp_ValRea_Es2.Text) + CDbl(ipp_ValRea_Dep1.Text) + CDbl(ipp_ValRea_Dep2.Text), "###,###,##0.00") & " "
   pnl_ValTer.Caption = Format(CDbl(ipp_ValTer_Inm.Text) + CDbl(ipp_ValTer_Es1.Text) + CDbl(ipp_ValTer_Es2.Text) + CDbl(ipp_ValTer_Dep1.Text) + CDbl(ipp_ValTer_Dep2.Text), "###,###,##0.00") & " "
   pnl_ValEdi.Caption = Format(CDbl(ipp_ValEdi_Inm.Text) + CDbl(ipp_ValEdi_Es1.Text) + CDbl(ipp_ValEdi_Es2.Text) + CDbl(ipp_ValEdi_Dep1.Text) + CDbl(ipp_ValEdi_Dep2.Text), "###,###,##0.00") & " "
   pnl_ValACo.Caption = Format(CDbl(ipp_ValACo_Inm.Text) + CDbl(ipp_ValACo_Es1.Text) + CDbl(ipp_ValACo_Es2.Text) + CDbl(ipp_ValACo_Dep1.Text) + CDbl(ipp_ValACo_Dep2.Text), "###,###,##0.00") & " "
End Sub

Private Sub cmb_EmpPer_Click()
   If cmb_EmpPer.ListIndex > -1 Then
      Screen.MousePointer = 11
      Call moddat_gs_Carga_PerTas(cmb_PerTas, l_arr_PerTas(), l_arr_EmpPer(cmb_EmpPer.ListIndex + 1).Genera_Codigo)
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmb_EmpPer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_PerTas)
   End If
End Sub

Private Sub cmb_FlgEst_Dep1_Click()
   If cmb_FlgEst_Dep1.ListIndex > -1 Then
      If cmb_FlgEst_Dep1.ItemData(cmb_FlgEst_Dep1.ListIndex) = 1 Then
         ipp_AreTer_Dep1.Enabled = True
         ipp_AreCon_Dep1.Enabled = True
         ipp_SumAse_Dep1.Enabled = True
         ipp_ValCom_Dep1.Enabled = True
         ipp_ValRea_Dep1.Enabled = True
         ipp_ValTer_Dep1.Enabled = True
         ipp_ValEdi_Dep1.Enabled = True
         ipp_ValACo_Dep1.Enabled = True
         Call gs_SetFocus(ipp_AreTer_Dep1)
      Else
         ipp_AreTer_Dep1.Enabled = False
         ipp_AreCon_Dep1.Enabled = False
         ipp_SumAse_Dep1.Enabled = False
         ipp_ValCom_Dep1.Enabled = False
         ipp_ValRea_Dep1.Enabled = False
         ipp_ValTer_Dep1.Enabled = False
         ipp_ValEdi_Dep1.Enabled = False
         ipp_ValACo_Dep1.Enabled = False
         ipp_AreTer_Dep1.Value = 0
         ipp_AreCon_Dep1.Value = 0
         ipp_SumAse_Dep1.Value = 0
         ipp_ValCom_Dep1.Value = 0
         ipp_ValRea_Dep1.Value = 0
         ipp_ValTer_Dep1.Value = 0
         ipp_ValEdi_Dep1.Value = 0
         ipp_ValACo_Dep1.Value = 0
         tab_Genera.Tab = 5
         Call gs_SetFocus(cmb_FlgEst_Dep2)
      End If
   End If
End Sub


Private Sub cmb_FlgEst_Dep2_Click()
   If cmb_FlgEst_Dep2.ListIndex > -1 Then
      If cmb_FlgEst_Dep2.ItemData(cmb_FlgEst_Dep2.ListIndex) = 1 Then
         ipp_AreTer_Dep2.Enabled = True
         ipp_AreCon_Dep2.Enabled = True
         ipp_SumAse_Dep2.Enabled = True
         ipp_ValCom_Dep2.Enabled = True
         ipp_ValRea_Dep2.Enabled = True
         ipp_ValTer_Dep2.Enabled = True
         ipp_ValEdi_Dep2.Enabled = True
         ipp_ValACo_Dep2.Enabled = True
         Call gs_SetFocus(ipp_AreTer_Dep2)
      Else
         ipp_AreTer_Dep2.Enabled = False
         ipp_AreCon_Dep2.Enabled = False
         ipp_SumAse_Dep2.Enabled = False
         ipp_ValCom_Dep2.Enabled = False
         ipp_ValRea_Dep2.Enabled = False
         ipp_ValTer_Dep2.Enabled = False
         ipp_ValEdi_Dep2.Enabled = False
         ipp_ValACo_Dep2.Enabled = False
         ipp_AreTer_Dep2.Value = 0
         ipp_AreCon_Dep2.Value = 0
         ipp_SumAse_Dep2.Value = 0
         ipp_ValCom_Dep2.Value = 0
         ipp_ValRea_Dep2.Value = 0
         ipp_ValTer_Dep2.Value = 0
         ipp_ValEdi_Dep2.Value = 0
         ipp_ValACo_Dep2.Value = 0
         tab_Genera.Tab = 0
         Call gs_SetFocus(cmd_Grabar)
      End If
   End If
End Sub

Private Sub cmb_FlgEst_Es1_Click()
   If cmb_FlgEst_Es1.ListIndex > -1 Then
      If cmb_FlgEst_Es1.ItemData(cmb_FlgEst_Es1.ListIndex) = 1 Then
         ipp_AreTer_Es1.Enabled = True
         ipp_AreCon_Es1.Enabled = True
         ipp_SumAse_Es1.Enabled = True
         ipp_ValCom_Es1.Enabled = True
         ipp_ValRea_Es1.Enabled = True
         ipp_ValTer_Es1.Enabled = True
         ipp_ValEdi_Es1.Enabled = True
         ipp_ValACo_Es1.Enabled = True
         Call gs_SetFocus(ipp_AreTer_Es1)
      Else
         ipp_AreTer_Es1.Enabled = False
         ipp_AreCon_Es1.Enabled = False
         ipp_SumAse_Es1.Enabled = False
         ipp_ValCom_Es1.Enabled = False
         ipp_ValRea_Es1.Enabled = False
         ipp_ValTer_Es1.Enabled = False
         ipp_ValEdi_Es1.Enabled = False
         ipp_ValACo_Es1.Enabled = False
         ipp_AreTer_Es1.Value = 0
         ipp_AreCon_Es1.Value = 0
         ipp_SumAse_Es1.Value = 0
         ipp_ValCom_Es1.Value = 0
         ipp_ValRea_Es1.Value = 0
         ipp_ValTer_Es1.Value = 0
         ipp_ValEdi_Es1.Value = 0
         ipp_ValACo_Es1.Value = 0
         tab_Genera.Tab = 3
         Call gs_SetFocus(cmb_FlgEst_Es2)
      End If
   End If
End Sub

Private Sub cmb_FlgEst_Es2_Click()
   If cmb_FlgEst_Es2.ListIndex > -1 Then
      If cmb_FlgEst_Es2.ItemData(cmb_FlgEst_Es2.ListIndex) = 1 Then
         ipp_AreTer_Es2.Enabled = True
         ipp_AreCon_Es2.Enabled = True
         ipp_SumAse_Es2.Enabled = True
         ipp_ValCom_Es2.Enabled = True
         ipp_ValRea_Es2.Enabled = True
         ipp_ValTer_Es2.Enabled = True
         ipp_ValEdi_Es2.Enabled = True
         ipp_ValACo_Es2.Enabled = True
         Call gs_SetFocus(ipp_AreTer_Es2)
      Else
         ipp_AreTer_Es2.Enabled = False
         ipp_AreCon_Es2.Enabled = False
         ipp_SumAse_Es2.Enabled = False
         ipp_ValCom_Es2.Enabled = False
         ipp_ValRea_Es2.Enabled = False
         ipp_ValTer_Es2.Enabled = False
         ipp_ValEdi_Es2.Enabled = False
         ipp_ValACo_Es2.Enabled = False
         ipp_AreTer_Es2.Value = 0
         ipp_AreCon_Es2.Value = 0
         ipp_SumAse_Es2.Value = 0
         ipp_ValCom_Es2.Value = 0
         ipp_ValRea_Es2.Value = 0
         ipp_ValTer_Es2.Value = 0
         ipp_ValEdi_Es2.Value = 0
         ipp_ValACo_Es2.Value = 0
         
         tab_Genera.Tab = 4
         Call gs_SetFocus(cmb_FlgEst_Dep1)
      End If
   End If
End Sub

Private Sub cmb_MatCon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipMon)
   End If
End Sub

Private Sub cmb_PerTas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumInf)
   End If
End Sub

Private Sub cmb_TieAzo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipInm)
   End If
End Sub

Private Sub cmb_TipInm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_UsoInm)
   End If
End Sub

Private Sub cmb_TipMon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_TipCam)
   End If
End Sub

Private Sub cmb_UsoInm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_MatCon)
   End If
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb_2 = 1
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_EmpPer)
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_Activa(False)
   Call fs_Limpia
   Call fs_Buscar
End Sub

Private Sub cmd_Editar_Click()
   moddat_g_int_FlgGrb_2 = 2
   Call fs_Activa(True)
   If grd_Listad.Row = -1 Then Exit Sub
   Call fs_Buscar_DatEva(grd_Listad.TextMatrix(grd_Listad.Row, 4))
   l_int_CodTas = grd_Listad.TextMatrix(grd_Listad.Row, 4)
   cmd_Editar.Enabled = False
   Call gs_SetFocus(cmb_EmpPer)
End Sub
Private Sub fs_Limpia()

   tab_Genera.Tab = 0
   cmb_EmpPer.ListIndex = -1
   cmb_PerTas.Clear
   txt_NumInf.Text = ""
   ipp_FecEva.Text = Format(date, "dd/mm/yyyy")
   ipp_AnoCon.Value = 0
   ipp_NumPis.Value = 0
   ipp_NumSot.Value = 0
   cmb_TieAzo.ListIndex = -1
   cmb_TipInm.ListIndex = -1
   cmb_UsoInm.ListIndex = -1
   cmb_MatCon.ListIndex = -1
   cmb_TipMon.ListIndex = -1
   ipp_TipCam.Value = 0
   
   ipp_AreTer_Inm.Value = 0
   ipp_AreCon_Inm.Value = 0
   ipp_SumAse_Inm.Value = 0
   ipp_ValCom_Inm.Value = 0
   ipp_ValRea_Inm.Value = 0
   ipp_ValTer_Inm.Value = 0
   ipp_ValEdi_Inm.Value = 0
   ipp_ValACo_Inm.Value = 0
   
   cmb_FlgEst_Es1.ListIndex = -1
   ipp_AreTer_Es1.Value = 0
   ipp_AreCon_Es1.Value = 0
   ipp_SumAse_Es1.Value = 0
   ipp_ValCom_Es1.Value = 0
   ipp_ValRea_Es1.Value = 0
   ipp_ValTer_Es1.Value = 0
   ipp_ValEdi_Es1.Value = 0
   ipp_ValACo_Es1.Value = 0
   ipp_AreTer_Es1.Enabled = False
   ipp_AreCon_Es1.Enabled = False
   ipp_SumAse_Es1.Enabled = False
   ipp_ValCom_Es1.Enabled = False
   ipp_ValRea_Es1.Enabled = False
   ipp_ValTer_Es1.Enabled = False
   ipp_ValEdi_Es1.Enabled = False
   ipp_ValACo_Es1.Enabled = False
   cmb_FlgEst_Es2.ListIndex = -1
   ipp_AreTer_Es2.Value = 0
   ipp_AreCon_Es2.Value = 0
   ipp_SumAse_Es2.Value = 0
   ipp_ValCom_Es2.Value = 0
   ipp_ValRea_Es2.Value = 0
   ipp_ValTer_Es2.Value = 0
   ipp_ValEdi_Es2.Value = 0
   ipp_ValACo_Es2.Value = 0
   ipp_AreTer_Es2.Enabled = False
   ipp_AreCon_Es2.Enabled = False
   ipp_SumAse_Es2.Enabled = False
   ipp_ValCom_Es2.Enabled = False
   ipp_ValRea_Es2.Enabled = False
   ipp_ValTer_Es2.Enabled = False
   ipp_ValEdi_Es2.Enabled = False
   ipp_ValACo_Es2.Enabled = False
   cmb_FlgEst_Dep1.ListIndex = -1
   ipp_AreTer_Dep1.Value = 0
   ipp_AreCon_Dep1.Value = 0
   ipp_SumAse_Dep1.Value = 0
   ipp_ValCom_Dep1.Value = 0
   ipp_ValRea_Dep1.Value = 0
   ipp_ValTer_Dep1.Value = 0
   ipp_ValEdi_Dep1.Value = 0
   ipp_ValACo_Dep1.Value = 0
   ipp_AreTer_Dep1.Enabled = False
   ipp_AreCon_Dep1.Enabled = False
   ipp_SumAse_Dep1.Enabled = False
   ipp_ValCom_Dep1.Enabled = False
   ipp_ValRea_Dep1.Enabled = False
   ipp_ValTer_Dep1.Enabled = False
   ipp_ValEdi_Dep1.Enabled = False
   ipp_ValACo_Dep1.Enabled = False
   cmb_FlgEst_Dep2.ListIndex = -1
   ipp_AreTer_Dep2.Value = 0
   ipp_AreCon_Dep2.Value = 0
   ipp_SumAse_Dep2.Value = 0
   ipp_ValCom_Dep2.Value = 0
   ipp_ValRea_Dep2.Value = 0
   ipp_ValTer_Dep2.Value = 0
   ipp_ValEdi_Dep2.Value = 0
   ipp_ValACo_Dep2.Value = 0
   ipp_AreTer_Dep2.Enabled = False
   ipp_AreCon_Dep2.Enabled = False
   ipp_SumAse_Dep2.Enabled = False
   ipp_ValCom_Dep2.Enabled = False
   ipp_ValRea_Dep2.Enabled = False
   ipp_ValTer_Dep2.Enabled = False
   ipp_ValEdi_Dep2.Enabled = False
   ipp_ValACo_Dep2.Enabled = False
   
End Sub
Private Sub cmd_Grabar_Click()
Dim r_dbl_TCaMpr           As Double
Dim r_int_CodTas           As Integer
   
   'Valida campos
   If cmb_EmpPer.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Empresa de Peritaje.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 0
      Call gs_SetFocus(cmb_EmpPer)
      Exit Sub
   End If
   If Len(Trim(cmb_PerTas.Text)) = 0 Then
      MsgBox "Debe seleccionar el Perito Tasador.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 0
      Call gs_SetFocus(cmb_PerTas)
      Exit Sub
   End If
   If Len(Trim(txt_NumInf.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Informe del Perito.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 0
      Call gs_SetFocus(txt_NumInf)
      Exit Sub
   End If
   If CDate(ipp_FecEva.Text) > date Then
      MsgBox "La Fecha de Evaluación no puede ser mayor a la Fecha Actual.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 0
      Call gs_SetFocus(ipp_FecEva)
      Exit Sub
   End If
   If ipp_AnoCon.Value = 0 Then
      MsgBox "Debe ingresar el Año de Construcción.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 0
      Call gs_SetFocus(ipp_AnoCon)
      Exit Sub
   End If
   If ipp_NumPis.Value = 0 Then
      MsgBox "Debe ingresar el Nro. de Pisos.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 0
      Call gs_SetFocus(ipp_NumPis)
      Exit Sub
   End If
   
   If cmb_TieAzo.ListIndex = -1 Then
      MsgBox "Debe seleccionar Tiene Azotea.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 0
      Call gs_SetFocus(cmb_TieAzo)
      Exit Sub
   End If
   
   
   If cmb_TipInm.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Inmueble.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 0
      Call gs_SetFocus(cmb_TipInm)
      Exit Sub
   End If
   If cmb_UsoInm.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Uso del Inmueble.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 0
      Call gs_SetFocus(cmb_UsoInm)
      Exit Sub
   End If
   If cmb_MatCon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Material de Construcción.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 0
      Call gs_SetFocus(cmb_MatCon)
      Exit Sub
   End If
   If cmb_TipMon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Moneda de la Valuación.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 0
      Call gs_SetFocus(cmb_TipMon)
      Exit Sub
   End If
'   If cmb_TipMon.ItemData(cmb_TipMon.ListIndex) <> moddat_g_int_TipMon Then
'      MsgBox "La Moneda de la Tasación debe ser igual a la Moneda del Préstamo.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(cmb_TipMon)
'      Exit Sub
'   End If
'   If ipp_TipCam.Value = 0 Then
'      MsgBox "Debe ingresar el Tipo de Cambio.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(ipp_TipCam)
'      Exit Sub
'   End If
   If ipp_AreTer_Inm.Value = 0 Then
      MsgBox "Debe ingresar el Area del Terreno.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 1
      Call gs_SetFocus(ipp_AreTer_Inm)
      Exit Sub
   End If
   If ipp_AreCon_Inm.Value = 0 Then
      MsgBox "Debe ingresar el Area Construida.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 1
      Call gs_SetFocus(ipp_AreCon_Inm)
      Exit Sub
   End If
   If ipp_SumAse_Inm.Value = 0 Then
      MsgBox "Debe ingresar la Suma Asegurada.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 1
      Call gs_SetFocus(ipp_SumAse_Inm)
      Exit Sub
   End If
   If ipp_ValCom_Inm.Value = 0 Then
      MsgBox "Debe ingresar el Valor Comercial.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 1
      Call gs_SetFocus(ipp_ValCom_Inm)
      Exit Sub
   End If
   If ipp_ValRea_Inm.Value = 0 Then
      MsgBox "Debe ingresar el Valor de Realización.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 1
      Call gs_SetFocus(ipp_ValRea_Inm)
      Exit Sub
   End If
   If ipp_ValTer_Inm.Value = 0 Then
      MsgBox "Debe ingresar el Valor del Terreno.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 1
      Call gs_SetFocus(ipp_ValTer_Inm)
      Exit Sub
   End If
   If ipp_ValEdi_Inm.Value = 0 Then
      MsgBox "Debe ingresar el Valor de Edificación.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 1
      Call gs_SetFocus(ipp_ValEdi_Inm)
      Exit Sub
   End If
   
   'Estacionamiento 1
   If cmb_FlgEst_Es1.ListIndex = -1 Then
      MsgBox "Debe seleccionar si hay valorización por Estacionamiento.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 2
      Call gs_SetFocus(cmb_FlgEst_Es1)
      Exit Sub
   End If
   
   If cmb_FlgEst_Es1.ItemData(cmb_FlgEst_Es1.ListIndex) = 1 Then
      If ipp_AreTer_Es1.Value = 0 Then
         MsgBox "Debe ingresar el Area del Terreno.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 2
         Call gs_SetFocus(ipp_AreTer_Es1)
         Exit Sub
      End If
      If ipp_ValCom_Es1.Value = 0 Then
         MsgBox "Debe ingresar el Valor Comercial.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 2
         Call gs_SetFocus(ipp_ValCom_Es1)
         Exit Sub
      End If
      If ipp_ValRea_Es1.Value = 0 Then
         MsgBox "Debe ingresar el Valor de Realización.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 2
         Call gs_SetFocus(ipp_ValRea_Es1)
         Exit Sub
      End If
      If ipp_ValTer_Es1.Value = 0 Then
         MsgBox "Debe ingresar el Valor del Terreno.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 2
         Call gs_SetFocus(ipp_ValTer_Es1)
         Exit Sub
      End If
      If ipp_ValEdi_Es1.Value = 0 Then
         MsgBox "Debe ingresar el Valor de Edificación.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 2
         Call gs_SetFocus(ipp_ValEdi_Es1)
         Exit Sub
      End If
   End If
   
   'Estacionamiento 2
   If cmb_FlgEst_Es2.ListIndex = -1 Then
      MsgBox "Debe seleccionar si hay valorización por Estacionamiento.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 3
      Call gs_SetFocus(cmb_FlgEst_Es2)
      Exit Sub
   End If
   
   If cmb_FlgEst_Es2.ItemData(cmb_FlgEst_Es2.ListIndex) = 1 Then
      If ipp_AreTer_Es2.Value = 0 Then
         MsgBox "Debe ingresar el Area del Terreno.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 3
         Call gs_SetFocus(ipp_AreTer_Es2)
         Exit Sub
      End If
      If ipp_ValCom_Es2.Value = 0 Then
         MsgBox "Debe ingresar el Valor Comercial.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 3
         Call gs_SetFocus(ipp_ValCom_Es2)
         Exit Sub
      End If
      If ipp_ValRea_Es2.Value = 0 Then
         MsgBox "Debe ingresar el Valor de Realización.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 3
         Call gs_SetFocus(ipp_ValRea_Es2)
         Exit Sub
      End If
      If ipp_ValTer_Es2.Value = 0 Then
         MsgBox "Debe ingresar el Valor del Terreno.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 3
         Call gs_SetFocus(ipp_ValTer_Es2)
         Exit Sub
      End If
      If ipp_ValEdi_Es2.Value = 0 Then
         MsgBox "Debe ingresar el Valor de Edificación.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 3
         Call gs_SetFocus(ipp_ValEdi_Es2)
         Exit Sub
      End If
   End If
   
   'Depósito 1
   If cmb_FlgEst_Dep1.ListIndex = -1 Then
      MsgBox "Debe seleccionar si hay valorización por el Depósito.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 4
      Call gs_SetFocus(cmb_FlgEst_Dep1)
      Exit Sub
   End If
   
   If cmb_FlgEst_Dep1.ItemData(cmb_FlgEst_Dep1.ListIndex) = 1 Then
      If ipp_AreTer_Dep1.Value = 0 Then
         MsgBox "Debe ingresar el Area del Terreno.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 4
         Call gs_SetFocus(ipp_AreTer_Dep1)
         Exit Sub
      End If
      If ipp_ValCom_Dep1.Value = 0 Then
         MsgBox "Debe ingresar el Valor Comercial.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 4
         Call gs_SetFocus(ipp_ValCom_Dep1)
         Exit Sub
      End If
      If ipp_ValRea_Dep1.Value = 0 Then
         MsgBox "Debe ingresar el Valor de Realización.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 4
         Call gs_SetFocus(ipp_ValRea_Dep1)
         Exit Sub
      End If
      If ipp_ValTer_Dep1.Value = 0 Then
         MsgBox "Debe ingresar el Valor del Terreno.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 4
         Call gs_SetFocus(ipp_ValTer_Dep1)
         Exit Sub
      End If
      If ipp_ValEdi_Dep1.Value = 0 Then
         MsgBox "Debe ingresar el Valor de Edificación.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 4
         Call gs_SetFocus(ipp_ValEdi_Dep1)
         Exit Sub
      End If
   End If
   
   'Depósito 2
   If cmb_FlgEst_Dep2.ListIndex = -1 Then
      MsgBox "Debe seleccionar si hay valorización por el Depósito.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 5
      Call gs_SetFocus(cmb_FlgEst_Dep2)
      Exit Sub
   End If
   
   If cmb_FlgEst_Dep2.ItemData(cmb_FlgEst_Dep2.ListIndex) = 1 Then
      If ipp_AreTer_Dep2.Value = 0 Then
         MsgBox "Debe ingresar el Area del Terreno.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 5
         Call gs_SetFocus(ipp_AreTer_Dep2)
         Exit Sub
      End If
      If ipp_ValCom_Dep2.Value = 0 Then
         MsgBox "Debe ingresar el Valor Comercial.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 5
         Call gs_SetFocus(ipp_ValCom_Dep2)
         Exit Sub
      End If
      If ipp_ValRea_Dep2.Value = 0 Then
         MsgBox "Debe ingresar el Valor de Realización.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 5
         Call gs_SetFocus(ipp_ValRea_Dep2)
         Exit Sub
      End If
      If ipp_ValTer_Dep2.Value = 0 Then
         MsgBox "Debe ingresar el Valor del Terreno.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 5
         Call gs_SetFocus(ipp_ValTer_Dep2)
         Exit Sub
      End If
      If ipp_ValEdi_Dep2.Value = 0 Then
         MsgBox "Debe ingresar el Valor de Edificación.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 5
         Call gs_SetFocus(ipp_ValEdi_Dep2)
         Exit Sub
      End If
   End If
   
   'Valida que el Valor de Realización sea como mínimo el 5% del PE.
'   If fs_Valida_ValRea(CDbl(CDbl(ipp_ValRea_Inm.Value) + CDbl(ipp_ValRea_Es1.Value) + CDbl(ipp_ValRea_Es2.Value) + CDbl(ipp_ValRea_Dep1.Value) + CDbl(ipp_ValRea_Dep2.Value))) = False Then
'      MsgBox "El Valor de Realización debe ser como mínimo el 5% del PE.", vbExclamation, modgen_g_str_NomPlt
'      tab_Genera.Tab = 1
'      Call gs_SetFocus(ipp_ValRea_Inm)
'      Exit Sub
'   End If
      
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   'Registro de Tasación
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
         
      'Grabando Información de Fondos Recibidos
      If moddat_g_int_FlgGrb_2 = 1 Then
         r_int_CodTas = fs_GeneraCodTas
      Else
         r_int_CodTas = l_int_CodTas
      End If
            
      r_dbl_TCaMpr = 0
      
      g_str_Parame = ""
      g_str_Parame = "USP_TPR_EVATAS ("
      g_str_Parame = g_str_Parame & moddat_g_int_TipDoc & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
      g_str_Parame = g_str_Parame & CStr(r_int_CodTas) & ", "
      g_str_Parame = g_str_Parame & "'" & l_arr_EmpPer(cmb_EmpPer.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_PerTas(cmb_PerTas.ListIndex + 1).Genera_Nombre & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_PerTas(cmb_PerTas.ListIndex + 1).Genera_Prefij & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NumInf.Text & "', "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecEva.Text), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AnoCon.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_NumPis.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_NumSot.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_TieAzo.ItemData(cmb_TieAzo.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_TipInm.ItemData(cmb_TipInm.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_UsoInm.ItemData(cmb_UsoInm.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_MatCon.ItemData(cmb_MatCon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_TipMon.ItemData(cmb_TipMon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_TipCam.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_TCaMpr) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AreTer_Inm.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AreCon_Inm.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_SumAse_Inm.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValCom_Inm.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValRea_Inm.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValTer_Inm.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValEdi_Inm.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValACo_Inm.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_FlgEst_Es1.ItemData(cmb_FlgEst_Es1.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AreTer_Es1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AreCon_Es1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_SumAse_Es1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValCom_Es1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValRea_Es1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValTer_Es1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValEdi_Es1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValACo_Es1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_FlgEst_Es2.ItemData(cmb_FlgEst_Es2.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AreTer_Es2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AreCon_Es2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_SumAse_Es2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValCom_Es2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValRea_Es2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValTer_Es2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValEdi_Es2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValACo_Es2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_FlgEst_Dep1.ItemData(cmb_FlgEst_Dep1.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AreTer_Dep1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AreCon_Dep1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_SumAse_Dep1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValCom_Dep1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValRea_Dep1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValTer_Dep1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValEdi_Dep1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValACo_Dep1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_FlgEst_Dep2.ItemData(cmb_FlgEst_Dep2.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AreTer_Dep2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AreCon_Dep2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_SumAse_Dep2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValCom_Dep2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValRea_Dep2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValTer_Dep2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValEdi_Dep2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValACo_Dep2.Value) & ", "
      
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb_2) & ") "
            
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
  
      MsgBox "Tasación se registro satisfactoriamente.", vbInformation, modgen_g_str_NomPlt
      Screen.MousePointer = 0
   Loop
   tab_Genera.Tab = 0
   Call fs_Limpia
   Call fs_Activa(False)
   Call fs_Buscar
'   Unload Me
End Sub
Private Function fs_Valida_ValRea(ByVal p_ValRea As Double) As Boolean
Dim r_dbl_ValExpo    As Double

   fs_Valida_ValRea = False
         
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "           SELECT NVL(CONLIM_PATEFE, 0) AS PATRIMONIO_EFECTIVO"
   g_str_Parame = g_str_Parame & "              FROM CTB_CONLIM "
   g_str_Parame = g_str_Parame & "             WHERE CONLIM_CODANO = (SELECT CONLIM_CODANO "
   g_str_Parame = g_str_Parame & "                                      FROM (SELECT DISTINCT CONLIM_CODANO, CONLIM_CODMES "
   g_str_Parame = g_str_Parame & "                                              FROM CTB_CONLIM "
   g_str_Parame = g_str_Parame & "                                             ORDER BY CONLIM_CODANO DESC, CONLIM_CODMES DESC) "
   g_str_Parame = g_str_Parame & "                                     WHERE ROWNUM < 2) "
   g_str_Parame = g_str_Parame & "               AND CONLIM_CODMES = (SELECT CONLIM_CODMES "
   g_str_Parame = g_str_Parame & "                                      FROM (SELECT DISTINCT CONLIM_CODANO, CONLIM_CODMES "
   g_str_Parame = g_str_Parame & "                                              FROM CTB_CONLIM "
   g_str_Parame = g_str_Parame & "                                             ORDER BY CONLIM_CODANO DESC, CONLIM_CODMES DESC)"
   g_str_Parame = g_str_Parame & "                                     WHERE ROWNUM < 2)"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Function
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      r_dbl_ValExpo = CDbl(Trim$(g_rst_Genera!PATRIMONIO_EFECTIVO)) * 0.05
      
      If CDbl(p_ValRea) < CDbl(r_dbl_ValExpo) Then    'Debe tener como mínimo el 5% del PE
         fs_Valida_ValRea = False
      Else
         fs_Valida_ValRea = True
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
   
End Function
Private Function fs_GeneraCodTas() As Integer
Dim r_str_Parame     As String

   fs_GeneraCodTas = 0
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT NVL(MAX(EVATAS_CODTAS),0) CODTAS FROM TPR_EVATAS "
   r_str_Parame = r_str_Parame & "  WHERE EVATAS_TIPDOC =  " & moddat_g_int_TipDoc & ""
   r_str_Parame = r_str_Parame & "    AND EVATAS_NUMDOC =  '" & CStr(moddat_g_str_NumDoc) & "'"
   
   If Not gf_EjecutaSQL(r_str_Parame, g_rst_GenAux, 3) Then
       Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Function
   End If
     
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      fs_GeneraCodTas = g_rst_GenAux!CODTAS + 1
   End If
End Function
Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Activa(False)
   Call fs_Buscar
   'Call fs_Buscar_Credito
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub
Private Sub fs_Buscar_DatEva(p_CodTas As Integer)

'   moddat_g_int_FlgGrb = 1

   'Datos de la tasación
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT * FROM TPR_EVATAS "
   g_str_Parame = g_str_Parame & "  WHERE EVATAS_TIPDOC =  " & moddat_g_int_TipDoc & ""
   g_str_Parame = g_str_Parame & "    AND EVATAS_NUMDOC =  '" & CStr(moddat_g_str_NumDoc) & "'"
   g_str_Parame = g_str_Parame & "    AND EVATAS_CODTAS =  " & p_CodTas & ""

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      cmd_Editar.Enabled = False
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      cmd_Editar.Enabled = True
      'moddat_g_int_FlgGrb = 2
      If Not IsNull(g_rst_Princi!EVATAS_CODEMP) Then
         cmb_EmpPer.ListIndex = gf_Busca_Arregl(l_arr_EmpPer, g_rst_Princi!EVATAS_CODEMP) - 1
         Call moddat_gs_Carga_PerTas(cmb_PerTas, l_arr_PerTas(), l_arr_EmpPer(cmb_EmpPer.ListIndex + 1).Genera_Codigo)
      End If
      
      If Not IsNull(g_rst_Princi!EVATAS_NOMPER) Then
         Call gs_BuscarCombo(cmb_PerTas, Trim(g_rst_Princi!EVATAS_NOMPER & ""))
      End If
      
      If Not IsNull(g_rst_Princi!EVATAS_NUMINF) Then
         txt_NumInf.Text = Trim(g_rst_Princi!EVATAS_NUMINF & "")
      End If
      
      If Not IsNull(g_rst_Princi!EVATAS_FECEVA) Then
         ipp_FecEva.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVATAS_FECEVA))
      End If
      
      If Not IsNull(g_rst_Princi!EVATAS_ANOCON) Then
         ipp_AnoCon.Value = g_rst_Princi!EVATAS_ANOCON
      End If
      
      If Not IsNull(g_rst_Princi!EVATAS_NUMPIS) Then
         ipp_NumPis.Value = g_rst_Princi!EVATAS_NUMPIS
      End If
      
      If Not IsNull(g_rst_Princi!EVATAS_NUMSOT) Then
         ipp_NumSot.Value = g_rst_Princi!EVATAS_NUMSOT
      End If
      
      If Not IsNull(g_rst_Princi!EVATAS_TIPINM) Then
         Call gs_BuscarCombo_Item(cmb_TipInm, g_rst_Princi!EVATAS_TIPINM)
      End If
      
      If Not IsNull(g_rst_Princi!EVATAS_USOINM) Then
         Call gs_BuscarCombo_Item(cmb_UsoInm, g_rst_Princi!EVATAS_USOINM)
      End If
      
      If Not IsNull(g_rst_Princi!EVATAS_MATCON) Then
         Call gs_BuscarCombo_Item(cmb_MatCon, g_rst_Princi!EVATAS_MATCON)
      End If
      
      If Not IsNull(g_rst_Princi!EVATAS_TIPMON) Then
         Call gs_BuscarCombo_Item(cmb_TipMon, g_rst_Princi!EVATAS_TIPMON)
      End If
      
      If Not IsNull(g_rst_Princi!EVATAS_FLGAZO) Then
         Call gs_BuscarCombo_Item(cmb_TieAzo, IIf(IsNull(g_rst_Princi!EVATAS_FLGAZO), 0, g_rst_Princi!EVATAS_FLGAZO))
      End If
      
      If Not IsNull(g_rst_Princi!EVATAS_TIPCAM) Then
         ipp_TipCam.Value = g_rst_Princi!EVATAS_TIPCAM
      End If
      
      If Not IsNull(g_rst_Princi!EVATAS_ARETER_INM) Then
         ipp_AreTer_Inm.Value = g_rst_Princi!EVATAS_ARETER_INM
      End If
      
      If Not IsNull(g_rst_Princi!EVATAS_ARECON_INM) Then
         ipp_AreCon_Inm.Value = g_rst_Princi!EVATAS_ARECON_INM
      End If
      
      If Not IsNull(g_rst_Princi!EVATAS_SUMASE_INM) Then
         ipp_SumAse_Inm.Value = g_rst_Princi!EVATAS_SUMASE_INM
      End If
      
      If Not IsNull(g_rst_Princi!EVATAS_VALCOM_INM) Then
         ipp_ValCom_Inm.Value = g_rst_Princi!EVATAS_VALCOM_INM
      End If
      
      If Not IsNull(g_rst_Princi!EVATAS_VALREA_INM) Then
         ipp_ValRea_Inm.Value = g_rst_Princi!EVATAS_VALREA_INM
      End If
      
      If Not IsNull(g_rst_Princi!EVATAS_VALTER_INM) Then
         ipp_ValTer_Inm.Value = g_rst_Princi!EVATAS_VALTER_INM
      End If
      
      If Not IsNull(g_rst_Princi!EVATAS_VALEDI_INM) Then
         ipp_ValEdi_Inm.Value = g_rst_Princi!EVATAS_VALEDI_INM
      End If
      
      If Not IsNull(g_rst_Princi!EVATAS_VALACO_INM) Then
         ipp_ValACo_Inm.Value = g_rst_Princi!EVATAS_VALACO_INM
      End If
      
      Call gs_BuscarCombo_Item(cmb_FlgEst_Es1, g_rst_Princi!EVATAS_FLGEST_ES1)
      If cmb_FlgEst_Es1.ItemData(cmb_FlgEst_Es1.ListIndex) = 1 Then
         ipp_AreTer_Es1.Value = g_rst_Princi!EVATAS_ARETER_ES1
         ipp_AreCon_Es1.Value = g_rst_Princi!EVATAS_ARECON_ES1
         ipp_SumAse_Es1.Value = g_rst_Princi!EVATAS_SUMASE_ES1
         ipp_ValCom_Es1.Value = g_rst_Princi!EVATAS_VALCOM_ES1
         ipp_ValRea_Es1.Value = g_rst_Princi!EVATAS_VALREA_ES1
         ipp_ValTer_Es1.Value = g_rst_Princi!EVATAS_VALTER_ES1
         ipp_ValEdi_Es1.Value = g_rst_Princi!EVATAS_VALEDI_ES1
         ipp_ValACo_Es1.Value = g_rst_Princi!EVATAS_VALACO_ES1
         ipp_AreTer_Es1.Enabled = True
         ipp_AreCon_Es1.Enabled = True
         ipp_SumAse_Es1.Enabled = True
         ipp_ValCom_Es1.Enabled = True
         ipp_ValRea_Es1.Enabled = True
         ipp_ValTer_Es1.Enabled = True
         ipp_ValEdi_Es1.Enabled = True
         ipp_ValACo_Es1.Enabled = True
      End If
   
      Call gs_BuscarCombo_Item(cmb_FlgEst_Es2, g_rst_Princi!EVATAS_FLGEST_ES2)
      If cmb_FlgEst_Es2.ItemData(cmb_FlgEst_Es2.ListIndex) = 1 Then
         ipp_AreTer_Es2.Value = g_rst_Princi!EVATAS_ARETER_ES2
         ipp_AreCon_Es2.Value = g_rst_Princi!EVATAS_ARECON_ES2
         ipp_SumAse_Es2.Value = g_rst_Princi!EVATAS_SUMASE_ES2
         ipp_ValCom_Es2.Value = g_rst_Princi!EVATAS_VALCOM_ES2
         ipp_ValRea_Es2.Value = g_rst_Princi!EVATAS_VALREA_ES2
         ipp_ValTer_Es2.Value = g_rst_Princi!EVATAS_VALTER_ES2
         ipp_ValEdi_Es2.Value = g_rst_Princi!EVATAS_VALEDI_ES2
         ipp_ValACo_Es2.Value = g_rst_Princi!EVATAS_VALACO_ES2
         ipp_AreTer_Es2.Enabled = True
         ipp_AreCon_Es2.Enabled = True
         ipp_SumAse_Es2.Enabled = True
         ipp_ValCom_Es2.Enabled = True
         ipp_ValRea_Es2.Enabled = True
         ipp_ValTer_Es2.Enabled = True
         ipp_ValEdi_Es2.Enabled = True
         ipp_ValACo_Es2.Enabled = True
      End If
      
      Call gs_BuscarCombo_Item(cmb_FlgEst_Dep1, g_rst_Princi!EVATAS_FLGEST_DE1)
      If cmb_FlgEst_Dep1.ItemData(cmb_FlgEst_Dep1.ListIndex) = 1 Then
         ipp_AreTer_Dep1.Value = g_rst_Princi!EVATAS_ARETER_DE1
         ipp_AreCon_Dep1.Value = g_rst_Princi!EVATAS_ARECON_DE1
         ipp_SumAse_Dep1.Value = g_rst_Princi!EVATAS_SUMASE_DE1
         ipp_ValCom_Dep1.Value = g_rst_Princi!EVATAS_VALCOM_DE1
         ipp_ValRea_Dep1.Value = g_rst_Princi!EVATAS_VALREA_DE1
         ipp_ValTer_Dep1.Value = g_rst_Princi!EVATAS_VALTER_DE1
         ipp_ValEdi_Dep1.Value = g_rst_Princi!EVATAS_VALEDI_DE1
         ipp_ValACo_Dep1.Value = g_rst_Princi!EVATAS_VALACO_DE1
         ipp_AreTer_Dep1.Enabled = True
         ipp_AreCon_Dep1.Enabled = True
         ipp_SumAse_Dep1.Enabled = True
         ipp_ValCom_Dep1.Enabled = True
         ipp_ValRea_Dep1.Enabled = True
         ipp_ValTer_Dep1.Enabled = True
         ipp_ValEdi_Dep1.Enabled = True
         ipp_ValACo_Dep1.Enabled = True
      End If
      
      Call gs_BuscarCombo_Item(cmb_FlgEst_Dep2, g_rst_Princi!EVATAS_FLGEST_DE2)
      If cmb_FlgEst_Dep2.ItemData(cmb_FlgEst_Dep2.ListIndex) = 1 Then
         ipp_AreTer_Dep2.Value = g_rst_Princi!EVATAS_ARETER_DE2
         ipp_AreCon_Dep2.Value = g_rst_Princi!EVATAS_ARECON_DE2
         ipp_SumAse_Dep2.Value = g_rst_Princi!EVATAS_SUMASE_DE2
         ipp_ValCom_Dep2.Value = g_rst_Princi!EVATAS_VALCOM_DE2
         ipp_ValRea_Dep2.Value = g_rst_Princi!EVATAS_VALREA_DE2
         ipp_ValTer_Dep2.Value = g_rst_Princi!EVATAS_VALTER_DE2
         ipp_ValEdi_Dep2.Value = g_rst_Princi!EVATAS_VALEDI_DE2
         ipp_ValACo_Dep2.Value = g_rst_Princi!EVATAS_VALACO_DE2
         ipp_AreTer_Dep2.Enabled = True
         ipp_AreCon_Dep2.Enabled = True
         ipp_SumAse_Dep2.Enabled = True
         ipp_ValCom_Dep2.Enabled = True
         ipp_ValRea_Dep2.Enabled = True
         ipp_ValTer_Dep2.Enabled = True
         ipp_ValEdi_Dep2.Enabled = True
         ipp_ValACo_Dep2.Enabled = True
      End If
      Call fs_Calcul
   Else
      cmd_Editar.Enabled = False
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   tab_Genera.Tab = 0
End Sub

Private Sub grd_Listad_Click()
   Call fs_Limpia
   Call cmd_Editar_Click
   Call fs_Activa(False)
End Sub

Private Sub ipp_AnoCon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_NumPis)
   End If
End Sub

Private Sub ipp_AreCon_Dep1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_AreCon_Dep1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_SumAse_Dep1)
   End If
End Sub

Private Sub ipp_AreCon_Dep2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_AreCon_Dep2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_SumAse_Dep2)
   End If
End Sub

Private Sub ipp_AreCon_Es1_Change()
   Call fs_Calcul
End Sub


Private Sub ipp_AreCon_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_SumAse_Es1)
   End If
End Sub

Private Sub ipp_AreCon_Es2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_AreCon_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_SumAse_Es2)
   End If
End Sub

Private Sub ipp_AreCon_Inm_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_AreCon_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_SumAse_Inm)
   End If
End Sub

Private Sub ipp_AreTer_Dep1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_AreTer_Dep1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AreCon_Dep1)
   End If
End Sub

Private Sub ipp_AreTer_Dep2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_AreTer_Dep2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AreCon_Dep2)
   End If
End Sub

Private Sub ipp_AreTer_Es1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_AreTer_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AreCon_Es1)
   End If
End Sub

Private Sub ipp_AreTer_Es2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_AreTer_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AreCon_Es2)
   End If
End Sub

Private Sub ipp_AreTer_Inm_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_AreTer_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AreCon_Inm)
   End If
End Sub

Private Sub ipp_FecEva_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AnoCon)
   End If
End Sub

Private Sub ipp_NumPis_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_NumSot)
   End If
End Sub

Private Sub ipp_NumSot_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TieAzo)
   End If
End Sub

Private Sub ipp_SumAse_Dep1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_SumAse_Dep1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValCom_Dep1)
   End If
End Sub

Private Sub ipp_SumAse_Dep2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_SumAse_Dep2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValCom_Dep2)
   End If
End Sub

Private Sub ipp_SumAse_Es1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_SumAse_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValCom_Es1)
   End If
End Sub

Private Sub ipp_SumAse_Es2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_SumAse_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValCom_Es2)
   End If
End Sub

Private Sub ipp_SumAse_Inm_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_SumAse_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValCom_Inm)
   End If
End Sub

Private Sub ipp_TipCam_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      tab_Genera.Tab = 1
      Call gs_SetFocus(ipp_AreTer_Inm)
   End If
End Sub

Private Sub ipp_ValACo_Dep1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValACo_Dep1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      tab_Genera.Tab = 5
      Call gs_SetFocus(cmb_FlgEst_Dep2)
   End If
End Sub

Private Sub ipp_ValACo_Dep2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValACo_Dep2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      tab_Genera.Tab = 0
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub ipp_ValACo_Es1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValACo_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      tab_Genera.Tab = 3
      Call gs_SetFocus(cmb_FlgEst_Es2)
   End If
End Sub

Private Sub ipp_ValACo_Es2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValACo_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      tab_Genera.Tab = 4
      Call gs_SetFocus(cmb_FlgEst_Dep1)
   End If
End Sub

Private Sub ipp_ValACo_Inm_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValACo_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      tab_Genera.Tab = 2
      Call gs_SetFocus(cmb_FlgEst_Es1)
   End If
End Sub

Private Sub ipp_ValCom_Dep1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValCom_Dep1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValRea_Dep1)
   End If
End Sub


Private Sub ipp_ValCom_Dep2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValCom_Dep2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValRea_Dep2)
   End If
End Sub

Private Sub ipp_ValCom_Es1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValCom_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValRea_Es1)
   End If
End Sub

Private Sub ipp_ValCom_Es2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValCom_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValRea_Es2)
   End If
End Sub

Private Sub ipp_ValCom_Inm_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValCom_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValRea_Inm)
   End If
End Sub

Private Sub ipp_ValEdi_Dep1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValEdi_Dep1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValACo_Dep1)
   End If
End Sub

Private Sub ipp_ValEdi_Dep2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValEdi_Dep2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValACo_Dep2)
   End If
End Sub

Private Sub ipp_ValEdi_Es1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValEdi_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValACo_Es1)
   End If
End Sub

Private Sub ipp_ValEdi_Es2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValEdi_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValACo_Es2)
   End If
End Sub

Private Sub ipp_ValEdi_Inm_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValEdi_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValACo_Inm)
   End If
End Sub

Private Sub ipp_ValRea_Dep1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValRea_Dep1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValTer_Dep1)
   End If
End Sub

Private Sub ipp_ValRea_Dep2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValRea_Dep2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValTer_Dep2)
   End If
End Sub

Private Sub ipp_ValRea_Es1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValRea_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValTer_Es1)
   End If
End Sub

Private Sub ipp_ValRea_Es2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValRea_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValTer_Es2)
   End If
End Sub

Private Sub ipp_ValRea_Inm_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValRea_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValTer_Inm)
   End If
End Sub

Private Sub ipp_ValTer_Dep1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValTer_Dep1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValEdi_Dep1)
   End If
End Sub

Private Sub ipp_ValTer_Dep2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValTer_Dep2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValEdi_Dep2)
   End If
End Sub

Private Sub ipp_ValTer_Es1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValTer_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValEdi_Es1)
   End If
End Sub

Private Sub ipp_ValTer_Es2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValTer_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValEdi_Es2)
   End If
End Sub

Private Sub ipp_ValTer_Inm_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValTer_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValEdi_Inm)
   End If
End Sub

Private Sub txt_NumInf_GotFocus()
   Call gs_SelecTodo(txt_NumInf)
End Sub

Private Sub txt_NumInf_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecEva)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-/():;:;")
   End If
End Sub
