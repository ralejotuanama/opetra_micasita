VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Seg_SolHip_64 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10125
   ClientLeft      =   7905
   ClientTop       =   1395
   ClientWidth     =   11595
   Icon            =   "OpeTra_frm_188.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10125
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10125
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11630
      _Version        =   65536
      _ExtentX        =   20514
      _ExtentY        =   17859
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
         Height          =   3465
         Left            =   30
         TabIndex        =   3
         Top             =   4620
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   6112
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
         Begin TabDlg.SSTab tab_Seguim 
            Height          =   3345
            Left            =   60
            TabIndex        =   4
            Top             =   60
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   5900
            _Version        =   393216
            Style           =   1
            Tabs            =   5
            TabsPerRow      =   5
            TabHeight       =   520
            TabCaption(0)   =   "Seguimiento Instancia"
            TabPicture(0)   =   "OpeTra_frm_188.frx":000C
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label11"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label8"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label7"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "pnl_DesOcu"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "SSPanel8"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "SSPanel14"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "SSPanel13"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "grd_LisOcu"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "SSPanel10"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "txt_Descar"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "txt_Observ"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).ControlCount=   11
            TabCaption(1)   =   "Excepciones Aplicadas"
            TabPicture(1)   =   "OpeTra_frm_188.frx":0028
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "txt_ObsExc"
            Tab(1).Control(1)=   "grd_LisExc"
            Tab(1).Control(2)=   "SSPanel4"
            Tab(1).Control(3)=   "SSPanel5"
            Tab(1).Control(4)=   "SSPanel9"
            Tab(1).Control(5)=   "SSPanel11"
            Tab(1).Control(6)=   "SSPanel12"
            Tab(1).Control(7)=   "pnl_DesExc"
            Tab(1).Control(8)=   "pnl_TipAut"
            Tab(1).Control(9)=   "pnl_motivo"
            Tab(1).Control(10)=   "lbl_motivo"
            Tab(1).Control(11)=   "Label2"
            Tab(1).Control(12)=   "Label3"
            Tab(1).Control(13)=   "Label4"
            Tab(1).ControlCount=   14
            TabCaption(2)   =   "Aprobación Condicionada"
            TabPicture(2)   =   "OpeTra_frm_188.frx":0044
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "txt_ObsCon"
            Tab(2).Control(1)=   "txt_LevCon"
            Tab(2).Control(2)=   "SSPanel17"
            Tab(2).Control(3)=   "grd_LisCon"
            Tab(2).Control(4)=   "SSPanel18"
            Tab(2).Control(5)=   "SSPanel19"
            Tab(2).Control(6)=   "SSPanel20"
            Tab(2).Control(7)=   "pnl_InsCon"
            Tab(2).Control(8)=   "Label15"
            Tab(2).Control(9)=   "Label14"
            Tab(2).Control(10)=   "Label12"
            Tab(2).ControlCount=   11
            TabCaption(3)   =   "Rechazo Solicitud"
            TabPicture(3)   =   "OpeTra_frm_188.frx":0060
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "txt_ObsRec"
            Tab(3).Control(1)=   "pnl_MotRec"
            Tab(3).Control(2)=   "Label19"
            Tab(3).Control(3)=   "Label10"
            Tab(3).ControlCount=   4
            TabCaption(4)   =   "Información Microempresarios"
            TabPicture(4)   =   "OpeTra_frm_188.frx":007C
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "SSPanel23"
            Tab(4).Control(1)=   "SSPanel22"
            Tab(4).ControlCount=   2
            Begin VB.TextBox txt_Observ 
               Height          =   645
               Left            =   1320
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   10
               Text            =   "OpeTra_frm_188.frx":0098
               Top             =   1980
               Width           =   10005
            End
            Begin VB.TextBox txt_Descar 
               Height          =   645
               Left            =   1320
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   9
               Text            =   "OpeTra_frm_188.frx":009C
               Top             =   2640
               Width           =   10005
            End
            Begin VB.TextBox txt_ObsExc 
               Height          =   975
               Left            =   -73770
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   8
               Text            =   "OpeTra_frm_188.frx":00A0
               Top             =   1980
               Width           =   10065
            End
            Begin VB.TextBox txt_ObsCon 
               Height          =   645
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   7
               Text            =   "OpeTra_frm_188.frx":00A4
               Top             =   1980
               Width           =   10005
            End
            Begin VB.TextBox txt_LevCon 
               Height          =   645
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   6
               Text            =   "OpeTra_frm_188.frx":00A8
               Top             =   2640
               Width           =   10005
            End
            Begin VB.TextBox txt_ObsRec 
               Height          =   2595
               Left            =   -73560
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   5
               Top             =   690
               Width           =   9915
            End
            Begin Threed.SSPanel SSPanel10 
               Height          =   45
               Left            =   30
               TabIndex        =   11
               Top             =   1560
               Width           =   11325
               _Version        =   65536
               _ExtentX        =   19976
               _ExtentY        =   79
               _StockProps     =   15
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.21
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_LisOcu 
               Height          =   855
               Left            =   30
               TabIndex        =   12
               Top             =   660
               Width           =   11325
               _ExtentX        =   19976
               _ExtentY        =   1508
               _Version        =   393216
               Rows            =   21
               Cols            =   5
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel13 
               Height          =   285
               Left            =   60
               TabIndex        =   13
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Ocurrencia"
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
            Begin Threed.SSPanel SSPanel14 
               Height          =   285
               Left            =   2400
               TabIndex        =   14
               Top             =   360
               Width           =   8595
               _Version        =   65536
               _ExtentX        =   15161
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Descripción Ocurrencia"
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
            Begin Threed.SSPanel SSPanel8 
               Height          =   285
               Left            =   1230
               TabIndex        =   15
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "H. Ocurrencia"
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
            Begin Threed.SSPanel pnl_DesOcu 
               Height          =   315
               Left            =   1320
               TabIndex        =   16
               Top             =   1650
               Width           =   10005
               _Version        =   65536
               _ExtentX        =   17648
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Día: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
               ForeColor       =   32768
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
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
            Begin MSFlexGridLib.MSFlexGrid grd_LisExc 
               Height          =   855
               Left            =   -74970
               TabIndex        =   17
               Top             =   660
               Width           =   11325
               _ExtentX        =   19976
               _ExtentY        =   1508
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
            Begin Threed.SSPanel SSPanel4 
               Height          =   285
               Left            =   -74940
               TabIndex        =   18
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Excepción"
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
            Begin Threed.SSPanel SSPanel5 
               Height          =   285
               Left            =   -69330
               TabIndex        =   19
               Top             =   360
               Width           =   5325
               _Version        =   65536
               _ExtentX        =   9393
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Descripción Excepción"
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
            Begin Threed.SSPanel SSPanel9 
               Height          =   285
               Left            =   -73770
               TabIndex        =   20
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "H. Excepción"
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
            Begin Threed.SSPanel SSPanel11 
               Height          =   285
               Left            =   -72600
               TabIndex        =   21
               Top             =   360
               Width           =   3285
               _Version        =   65536
               _ExtentX        =   5794
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Instancia"
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
            Begin Threed.SSPanel SSPanel12 
               Height          =   45
               Left            =   -74970
               TabIndex        =   22
               Top             =   1560
               Width           =   11325
               _Version        =   65536
               _ExtentX        =   19976
               _ExtentY        =   79
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
            Begin Threed.SSPanel pnl_DesExc 
               Height          =   315
               Left            =   -73770
               TabIndex        =   23
               Top             =   1650
               Width           =   10065
               _Version        =   65536
               _ExtentX        =   17754
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Día: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
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
            Begin Threed.SSPanel pnl_TipAut 
               Height          =   315
               Left            =   -73770
               TabIndex        =   24
               Top             =   2970
               Width           =   4005
               _Version        =   65536
               _ExtentX        =   7064
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "INGRESO A INSTANCIA"
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
            Begin Threed.SSPanel SSPanel17 
               Height          =   45
               Left            =   -74970
               TabIndex        =   25
               Top             =   1560
               Width           =   11325
               _Version        =   65536
               _ExtentX        =   19976
               _ExtentY        =   79
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
            Begin MSFlexGridLib.MSFlexGrid grd_LisCon 
               Height          =   855
               Left            =   -74970
               TabIndex        =   26
               Top             =   660
               Width           =   11325
               _ExtentX        =   19976
               _ExtentY        =   1508
               _Version        =   393216
               Rows            =   21
               Cols            =   4
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel18 
               Height          =   285
               Left            =   -74940
               TabIndex        =   27
               Top             =   360
               Width           =   2745
               _Version        =   65536
               _ExtentX        =   4842
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Instancia"
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
            Begin Threed.SSPanel SSPanel19 
               Height          =   285
               Left            =   -65610
               TabIndex        =   28
               Top             =   360
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Situación"
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
            Begin Threed.SSPanel SSPanel20 
               Height          =   285
               Left            =   -72210
               TabIndex        =   29
               Top             =   360
               Width           =   6615
               _Version        =   65536
               _ExtentX        =   11668
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Condiciones de Aprobación"
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
            Begin Threed.SSPanel pnl_InsCon 
               Height          =   315
               Left            =   -73680
               TabIndex        =   30
               Top             =   1650
               Width           =   10005
               _Version        =   65536
               _ExtentX        =   17648
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Día: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
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
            Begin Threed.SSPanel pnl_MotRec 
               Height          =   315
               Left            =   -73560
               TabIndex        =   31
               Top             =   360
               Width           =   9915
               _Version        =   65536
               _ExtentX        =   17489
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
            Begin Threed.SSPanel pnl_motivo 
               Height          =   315
               Left            =   -68940
               TabIndex        =   66
               Top             =   2970
               Width           =   5265
               _Version        =   65536
               _ExtentX        =   9287
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "MOTIVO DE EXCEPCION"
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
               Begin Threed.SSPanel SSPanel21 
                  Height          =   315
                  Left            =   6090
                  TabIndex        =   67
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   5205
                  _Version        =   65536
                  _ExtentX        =   9181
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "INGRESOS 4A CATEG. NO SE PUEDEN CONFIRMAR "
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
            End
            Begin Threed.SSPanel SSPanel22 
               Height          =   1560
               Left            =   -74940
               TabIndex        =   69
               Top             =   390
               Width           =   11280
               _Version        =   65536
               _ExtentX        =   19897
               _ExtentY        =   2752
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
               Begin VB.CommandButton cmd_VerArc_1 
                  Caption         =   "Ver"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   10500
                  TabIndex        =   84
                  Top             =   150
                  Width           =   465
               End
               Begin VB.CommandButton cmd_VerArc_2 
                  Caption         =   "Ver"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   10500
                  TabIndex        =   83
                  Top             =   480
                  Width           =   465
               End
               Begin VB.CommandButton cmd_VerArc_3 
                  Caption         =   "Ver"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   10500
                  TabIndex        =   82
                  Top             =   810
                  Width           =   465
               End
               Begin VB.CommandButton cmd_VerArc_4 
                  Caption         =   "Ver"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   10500
                  TabIndex        =   81
                  Top             =   1140
                  Width           =   465
               End
               Begin Threed.SSPanel pnl_ArcItem_3 
                  Height          =   315
                  Left            =   2130
                  TabIndex        =   77
                  Top             =   810
                  Width           =   8325
                  _Version        =   65536
                  _ExtentX        =   14684
                  _ExtentY        =   556
                  _StockProps     =   15
                  ForeColor       =   16711680
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
               Begin Threed.SSPanel pnl_ArcItem_4 
                  Height          =   315
                  Left            =   2130
                  TabIndex        =   78
                  Top             =   1140
                  Width           =   8325
                  _Version        =   65536
                  _ExtentX        =   14684
                  _ExtentY        =   556
                  _StockProps     =   15
                  ForeColor       =   16711680
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
               Begin Threed.SSPanel pnl_ArcItem_2 
                  Height          =   315
                  Left            =   2130
                  TabIndex        =   79
                  Top             =   480
                  Width           =   8325
                  _Version        =   65536
                  _ExtentX        =   14684
                  _ExtentY        =   556
                  _StockProps     =   15
                  ForeColor       =   16711680
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
               Begin Threed.SSPanel pnl_ArcItem_1 
                  Height          =   315
                  Left            =   2130
                  TabIndex        =   80
                  Top             =   150
                  Width           =   8325
                  _Version        =   65536
                  _ExtentX        =   14684
                  _ExtentY        =   556
                  _StockProps     =   15
                  ForeColor       =   16711680
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
               Begin VB.Label Label17 
                  AutoSize        =   -1  'True
                  Caption         =   "Estados Financieros:"
                  Height          =   195
                  Left            =   150
                  TabIndex        =   73
                  Top             =   210
                  Width           =   1470
               End
               Begin VB.Label Label16 
                  AutoSize        =   -1  'True
                  Caption         =   "Flujo de Caja:"
                  Height          =   195
                  Left            =   150
                  TabIndex        =   72
                  Top             =   540
                  Width           =   960
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  Caption         =   "Hoja de Trabajo N°1:"
                  Height          =   195
                  Left            =   150
                  TabIndex        =   71
                  Top             =   870
                  Width           =   1500
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  Caption         =   "Hoja de Trabajo N°2:"
                  Height          =   195
                  Left            =   150
                  TabIndex        =   70
                  Top             =   1200
                  Width           =   1500
               End
            End
            Begin Threed.SSPanel SSPanel23 
               Height          =   1270
               Left            =   -74940
               TabIndex        =   74
               Top             =   1995
               Width           =   11280
               _Version        =   65536
               _ExtentX        =   19897
               _ExtentY        =   2240
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
               Begin VB.TextBox txt_Coment 
                  Height          =   1005
                  Left            =   2100
                  MaxLength       =   250
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   75
                  Top             =   150
                  Width           =   8835
               End
               Begin VB.Label Label18 
                  AutoSize        =   -1  'True
                  Caption         =   "Comentario:"
                  Height          =   195
                  Left            =   150
                  TabIndex        =   76
                  Top             =   510
                  Width           =   840
               End
            End
            Begin VB.Label lbl_motivo 
               Caption         =   "Motivo:"
               Height          =   255
               Left            =   -69630
               TabIndex        =   68
               Top             =   3030
               Width           =   645
            End
            Begin VB.Label Label7 
               Caption         =   "Comentario u Observación:"
               Height          =   495
               Left            =   60
               TabIndex        =   42
               Top             =   1980
               Width           =   1155
            End
            Begin VB.Label Label8 
               Caption         =   "Ocurrencia:"
               Height          =   315
               Left            =   60
               TabIndex        =   41
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label11 
               Caption         =   "Descargo:"
               Height          =   315
               Left            =   60
               TabIndex        =   40
               Top             =   2640
               Width           =   1035
            End
            Begin VB.Label Label2 
               Caption         =   "Autorizado por:"
               Height          =   315
               Left            =   -74910
               TabIndex        =   39
               Top             =   2970
               Width           =   1095
            End
            Begin VB.Label Label3 
               Caption         =   "Excepción:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   38
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label4 
               Caption         =   "Descripción:"
               Height          =   495
               Left            =   -74940
               TabIndex        =   37
               Top             =   1980
               Width           =   1155
            End
            Begin VB.Label Label15 
               Caption         =   "Condiciones de Aprobación:"
               Height          =   495
               Left            =   -74940
               TabIndex        =   36
               Top             =   1980
               Width           =   1155
            End
            Begin VB.Label Label14 
               Caption         =   "Instancia:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   35
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label12 
               Caption         =   "Levantamiento de Condiciones:"
               Height          =   615
               Left            =   -74940
               TabIndex        =   34
               Top             =   2640
               Width           =   1215
            End
            Begin VB.Label Label19 
               Caption         =   "Observaciones de Rechazo:"
               Height          =   555
               Left            =   -74940
               TabIndex        =   33
               Top             =   690
               Width           =   1155
            End
            Begin VB.Label Label10 
               Caption         =   "Motivo Rechazo:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   32
               Top             =   360
               Width           =   1275
            End
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   2325
         Left            =   30
         TabIndex        =   43
         Top             =   2250
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   4101
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
            Height          =   2205
            Left            =   60
            TabIndex        =   44
            Top             =   60
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   3889
            _Version        =   393216
            Style           =   1
            Tabs            =   5
            TabsPerRow      =   5
            TabHeight       =   520
            TabCaption(0)   =   "Cliente"
            TabPicture(0)   =   "OpeTra_frm_188.frx":00AC
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Cónyuge"
            TabPicture(1)   =   "OpeTra_frm_188.frx":00C8
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Apoderado"
            TabPicture(2)   =   "OpeTra_frm_188.frx":00E4
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(2)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Inmueble"
            TabPicture(3)   =   "OpeTra_frm_188.frx":0100
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "grd_Listad(3)"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Datos del Crédito"
            TabPicture(4)   =   "OpeTra_frm_188.frx":011C
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "txt_ObsSol"
            Tab(4).Control(1)=   "grd_Listad(4)"
            Tab(4).Control(2)=   "Label5"
            Tab(4).ControlCount=   3
            Begin VB.TextBox txt_ObsSol 
               Height          =   675
               Left            =   -73710
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   63
               Text            =   "OpeTra_frm_188.frx":0138
               Top             =   1470
               Width           =   10005
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1815
               Index           =   0
               Left            =   60
               TabIndex        =   45
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3201
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
               Height          =   1815
               Index           =   1
               Left            =   -74940
               TabIndex        =   46
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3201
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
               Height          =   1815
               Index           =   3
               Left            =   -74940
               TabIndex        =   47
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3201
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1815
               Index           =   2
               Left            =   -74940
               TabIndex        =   48
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3201
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
               Height          =   1095
               Index           =   4
               Left            =   -74940
               TabIndex        =   64
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   1931
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
               Caption         =   "Observaciones:"
               Height          =   495
               Left            =   -74940
               TabIndex        =   65
               Top             =   1470
               Width           =   1155
            End
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   49
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
         Begin Threed.SSPanel pnl_AprCon 
            Height          =   555
            Left            =   8460
            TabIndex        =   50
            Top             =   60
            Width           =   3015
            _Version        =   65536
            _ExtentX        =   5318
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "CLIENTE CON APROBACION CONDICIONADA PENDIENTE"
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
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Left            =   630
            TabIndex        =   51
            Top             =   30
            Width           =   5685
            _Version        =   65536
            _ExtentX        =   10028
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Seguimiento de Solicitud de Crédito Hipotecario"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
            Left            =   660
            TabIndex        =   52
            Top             =   330
            Width           =   5505
            _Version        =   65536
            _ExtentX        =   9710
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Evaluación Crediticia"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   6780
            Top             =   180
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Presentación Preliminar"
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_188.frx":013C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   53
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1440
            TabIndex        =   54
            Top             =   390
            Width           =   10035
            _Version        =   65536
            _ExtentX        =   17701
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1440
            TabIndex        =   55
            Top             =   60
            Width           =   1875
            _Version        =   65536
            _ExtentX        =   3307
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
         Begin Threed.SSPanel pnl_Situac 
            Height          =   315
            Left            =   7620
            TabIndex        =   56
            Top             =   60
            Width           =   3855
            _Version        =   65536
            _ExtentX        =   6800
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "SOLICITUD EN TRAMITE"
            ForeColor       =   16711680
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Solicitud"
            Height          =   195
            Left            =   150
            TabIndex        =   59
            Top             =   90
            Width           =   945
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   150
            TabIndex        =   58
            Top             =   390
            Width           =   525
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Situación:"
            Height          =   195
            Left            =   6210
            TabIndex        =   57
            Top             =   90
            Width           =   705
         End
      End
      Begin Threed.SSPanel SSPanel39 
         Height          =   645
         Left            =   30
         TabIndex        =   60
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "OpeTra_frm_188.frx":0446
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_188.frx":0888
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Nueva Observación"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel16 
         Height          =   1935
         Left            =   30
         TabIndex        =   61
         Top             =   8130
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   3413
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisEva 
            Height          =   1815
            Left            =   60
            TabIndex        =   62
            Top             =   60
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   3201
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
      End
   End
End
Attribute VB_Name = "frm_Seg_SolHip_64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_AprCon        As Integer
Dim l_int_FlgRec        As Integer
Dim l_str_EmpSeg        As String
Dim l_int_CodIns        As Integer

Private Sub cmd_Imprim_Click()
Dim r_rst_Princi     As ADODB.Recordset
Dim r_rst_Genera     As ADODB.Recordset

   If MsgBox("¿Está seguro de imprimir la Carta de Aprobación?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   '-----------------------------------------------------
   crp_Imprim.Reset
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "CRE_SOLMAE"
   crp_Imprim.DataFiles(1) = "TRA_EVACRE"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "TRA_SEGUIM"
   crp_Imprim.DataFiles(4) = "CRE_SUBPRD"
   crp_Imprim.DataFiles(5) = "CRE_PRODUC"
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT F.SOLMAE_NUMERO, TRIM(M.PARDES_DESCRI) AS TIPO_AFP, F.SOLMAE_AFPMTO, "
   g_str_Parame = g_str_Parame & "        (SELECT NVL(CASE WHEN N.SOLINM_TIPDOC_PRO = 7 THEN TRIM(X.DATGEN_RAZSOC) ELSE TRIM(N.SOLINM_RAZSOC_PRO) END, '-') "
   g_str_Parame = g_str_Parame & "           FROM EMP_DATGEN X WHERE X.DATGEN_EMPTDO = N.SOLINM_TIPDOC_PRO AND X.DATGEN_EMPNDO = N.SOLINM_NUMDOC_PRO) AS NOM_PROMOTOR "
   g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE F "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN A ON A.DATGEN_TIPDOC = F.SOLMAE_TITTDO AND A.DATGEN_NUMDOC = F.SOLMAE_TITNDO "
   g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES M ON M.PARDES_CODGRP = 517 AND M.PARDES_CODITE = A.DATGEN_TIPAFP "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_SOLINM N ON N.SOLINM_NUMSOL = F.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & " WHERE F.SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
       Exit Sub
   End If
   
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      crp_Imprim.ParameterFields(0) = "P_NOMPRO;" & "PROMOTOR: " & r_rst_Genera!NOM_PROMOTOR & ";True"
      If r_rst_Genera!SOLMAE_AFPMTO > 0 Then
         crp_Imprim.ParameterFields(1) = "P_NOMAFP;" & "AFP: " & r_rst_Genera!TIPO_AFP & ";True"
         Else
         crp_Imprim.ParameterFields(1) = "P_NOMAFP;" & "" & ";True"
      End If
   End If
   
   crp_Imprim.SelectionFormula = "{CRE_SOLMAE.SOLMAE_NUMERO} = '" & moddat_g_str_NumSol & "' AND {TRA_SEGUIM.SEGUIM_NUMSOL} = '" & moddat_g_str_NumSol & "' AND {TRA_SEGUIM.SEGUIM_CODINS} = 21"
   
   If Format(moddat_g_str_FecIng, "YYYYMMDD") > 20130822 Then
      If Mid(moddat_g_str_NumSol, 1, 3) = "021" Or Mid(moddat_g_str_NumSol, 1, 3) = "022" Or Mid(moddat_g_str_NumSol, 1, 3) = "023" Or Mid(moddat_g_str_NumSol, 1, 3) = "024" Or Mid(moddat_g_str_NumSol, 1, 3) = "025" Then
         crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_CARAPR_14.RPT"
      Else
         crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_CARAPR_12.RPT"
      End If
   Else
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_CARAPR_11.RPT"
   End If
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Action = 1
   
   '---------------------------------------------
   If Mid(moddat_g_str_NumSol, 1, 3) = "024" Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " SELECT F.SOLMAE_NUMERO, B.SEGUIM_FECFIN, "
      g_str_Parame = g_str_Parame & "        TRIM(H.PARDES_DESCRI) AS TIT_TIPDOC, TRIM(A.DATGEN_APEPAT) ||' '|| TRIM(A.DATGEN_APEMAT) AS TIT_APE, TRIM(A.DATGEN_NOMBRE) AS TIT_NOM, "
      g_str_Parame = g_str_Parame & "        TRIM(E.PARDES_DESCRI) AS TIT_ORDEN, TRIM(D.PARDES_DESCRI) AS TIT_OCUPACION, "
      g_str_Parame = g_str_Parame & "        TRIM(I.PARDES_DESCRI) AS CYG_TIPDOC, TRIM(G.DATGEN_APEPAT) ||' '|| TRIM(G.DATGEN_APEMAT) AS CYG_APE, TRIM(G.DATGEN_NOMBRE) AS CYG_NOM, TRIM(G.DATGEN_APECAS) AS CYG_APE_CAS, "
      g_str_Parame = g_str_Parame & "        TRIM(L.PARDES_DESCRI) AS CYG_ORDEN, TRIM(K.PARDES_DESCRI) AS CYG_OCUPACION "
      g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE F "
      g_str_Parame = g_str_Parame & "  INNER JOIN CLI_DATGEN A ON A.DATGEN_TIPDOC = F.SOLMAE_TITTDO AND A.DATGEN_NUMDOC = F.SOLMAE_TITNDO "
      g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES H ON H.PARDES_CODGRP = 203 AND H.PARDES_CODITE = A.DATGEN_TIPDOC "
      g_str_Parame = g_str_Parame & "  INNER JOIN TRA_SEGUIM B ON B.SEGUIM_NUMSOL = F.SOLMAE_NUMERO AND B.SEGUIM_CODINS = 21 AND B.SEGUIM_SITUAC = 1 "
      g_str_Parame = g_str_Parame & "   LEFT JOIN CLI_ACTECO C ON C.ACTECO_CLITDO = A.DATGEN_TIPDOC AND C.ACTECO_CLINDO = A.DATGEN_NUMDOC "
      g_str_Parame = g_str_Parame & "   LEFT JOIN MNT_PARDES D ON D.PARDES_CODGRP = 8 AND D.PARDES_CODITE = C.ACTECO_CODACT "
      g_str_Parame = g_str_Parame & "   LEFT JOIN MNT_PARDES E ON E.PARDES_CODGRP = 7 AND E.PARDES_CODITE = C.ActEco_OrdAct "
      g_str_Parame = g_str_Parame & "   LEFT JOIN CLI_DATGEN G ON G.DATGEN_TIPDOC = F.SOLMAE_CYGTDO AND G.DATGEN_NUMDOC = F.SOLMAE_CYGNDO "
      g_str_Parame = g_str_Parame & "   LEFT JOIN MNT_PARDES I ON I.PARDES_CODGRP = 203 AND I.PARDES_CODITE = G.DATGEN_TIPDOC "
      g_str_Parame = g_str_Parame & "   LEFT JOIN CLI_ACTECO J ON J.ACTECO_CLITDO = G.DATGEN_TIPDOC AND J.ACTECO_CLINDO = G.DATGEN_NUMDOC "
      g_str_Parame = g_str_Parame & "   LEFT JOIN MNT_PARDES K ON K.PARDES_CODGRP = 8 AND K.PARDES_CODITE = J.ACTECO_CODACT "
      g_str_Parame = g_str_Parame & "   LEFT JOIN MNT_PARDES L ON L.PARDES_CODGRP = 7 AND L.PARDES_CODITE = J.ActEco_OrdAct "
  
      g_str_Parame = g_str_Parame & "  WHERE F.SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "'"
   
      If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
          Exit Sub
      End If
      
      If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
         r_rst_Princi.MoveFirst
         crp_Imprim.Reset
         crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
         crp_Imprim.DataFiles(0) = "CRE_SOLMAE"
         crp_Imprim.DataFiles(1) = "TRA_EVACRE"
         
         crp_Imprim.ParameterFields(0) = "P_FecEva;" & Mid(r_rst_Princi!SEGUIM_FECFIN, 7, 2) & "/" & Mid(r_rst_Princi!SEGUIM_FECFIN, 5, 2) & "/" & Mid(r_rst_Princi!SEGUIM_FECFIN, 1, 4) & ";True"
         crp_Imprim.ParameterFields(1) = "P_ApeNom;" & r_rst_Princi!TIT_APE & " " & r_rst_Princi!TIT_NOM & "," & ";True"
         crp_Imprim.ParameterFields(2) = "P_ApeTit;" & r_rst_Princi!TIT_APE & ";True"
         crp_Imprim.ParameterFields(3) = "P_NomTit;" & r_rst_Princi!TIT_NOM & ";True"
         crp_Imprim.ParameterFields(4) = "P_TipDoc_Tit;" & r_rst_Princi!TIT_TIPDOC & ";True"
         If Not IsNull(r_rst_Princi!CYG_APE_CAS) Then
            crp_Imprim.ParameterFields(5) = "P_ApeCyg;" & r_rst_Princi!CYG_APE & " DE " & r_rst_Princi!CYG_APE_CAS & ";True"
         Else
            crp_Imprim.ParameterFields(5) = "P_ApeCyg;" & r_rst_Princi!CYG_APE & ";True"
         End If
         crp_Imprim.ParameterFields(6) = "P_NomCyg;" & r_rst_Princi!CYG_NOM & ";True"
         crp_Imprim.ParameterFields(7) = "P_TipDoc_Cyg;" & r_rst_Princi!CYG_TIPDOC & ";True"
         crp_Imprim.ParameterFields(8) = "P_SitLab_Tit;" & r_rst_Princi!TIT_OCUPACION & ";True"
         crp_Imprim.ParameterFields(9) = "P_SitLab_Cyg;" & r_rst_Princi!CYG_OCUPACION & ";True"
            
            
         If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
            r_rst_Genera.MoveFirst
            crp_Imprim.ParameterFields(10) = "P_NOMPRO;" & "PROMOTOR: " & r_rst_Genera!NOM_PROMOTOR & ";True"
            If r_rst_Genera!SOLMAE_AFPMTO > 0 Then
               crp_Imprim.ParameterFields(11) = "P_NOMAFP;" & "AFP: " & r_rst_Genera!TIPO_AFP & ";True"
               Else
               crp_Imprim.ParameterFields(11) = "P_NOMAFP;" & "" & ";True"
            End If
         End If
            
         crp_Imprim.SelectionFormula = "{CRE_SOLMAE.SOLMAE_NUMERO} = '" & moddat_g_str_NumSol & "'" & " AND {TRA_EVACRE.EVACRE_NUMSOL} = '" & moddat_g_str_NumSol & "'"
         crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_CARAPR_17.RPT"
         crp_Imprim.WindowShowPrintSetupBtn = True
         crp_Imprim.Action = 1
      End If
      
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
   End If
End Sub

Private Sub cmd_Imprim_Click_old()
Dim r_rst_Princi     As ADODB.Recordset

   If MsgBox("¿Está seguro de imprimir la Carta de Aprobación?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   crp_Imprim.Reset
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "CRE_SOLMAE"
   crp_Imprim.DataFiles(1) = "TRA_EVACRE"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "TRA_SEGUIM"
   crp_Imprim.DataFiles(4) = "CRE_SUBPRD"
   crp_Imprim.DataFiles(5) = "CRE_PRODUC"
   crp_Imprim.SelectionFormula = "{CRE_SOLMAE.SOLMAE_NUMERO} = '" & moddat_g_str_NumSol & "' AND {TRA_SEGUIM.SEGUIM_NUMSOL} = '" & moddat_g_str_NumSol & "' AND {TRA_SEGUIM.SEGUIM_CODINS} = 21"
   
   If Format(moddat_g_str_FecIng, "YYYYMMDD") > 20130822 Then
      If Mid(moddat_g_str_NumSol, 1, 3) = "021" Or Mid(moddat_g_str_NumSol, 1, 3) = "022" Or Mid(moddat_g_str_NumSol, 1, 3) = "023" Or Mid(moddat_g_str_NumSol, 1, 3) = "024" Or Mid(moddat_g_str_NumSol, 1, 3) = "025" Then
         crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_CARAPR_14.RPT"
      Else
         crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_CARAPR_12.RPT"
      End If
   Else
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_CARAPR_11.RPT"
   End If
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Action = 1
   
   '---------------------------------------------
   If Mid(moddat_g_str_NumSol, 1, 3) = "024" Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " SELECT F.SOLMAE_NUMERO, B.SEGUIM_FECFIN, "
      g_str_Parame = g_str_Parame & "        TRIM(H.PARDES_DESCRI) AS TIT_TIPDOC, TRIM(A.DATGEN_APEPAT) ||' '|| TRIM(A.DATGEN_APEMAT) AS TIT_APE, TRIM(A.DATGEN_NOMBRE) AS TIT_NOM, "
      g_str_Parame = g_str_Parame & "        TRIM(E.PARDES_DESCRI) AS TIT_ORDEN, TRIM(D.PARDES_DESCRI) AS TIT_OCUPACION, "
      g_str_Parame = g_str_Parame & "        TRIM(I.PARDES_DESCRI) AS CYG_TIPDOC, TRIM(G.DATGEN_APEPAT) ||' '|| TRIM(G.DATGEN_APEMAT) AS CYG_APE, TRIM(G.DATGEN_NOMBRE) AS CYG_NOM, "
      g_str_Parame = g_str_Parame & "        TRIM(L.PARDES_DESCRI) AS CYG_ORDEN, TRIM(K.PARDES_DESCRI) AS CYG_OCUPACION "
      g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE F "
      g_str_Parame = g_str_Parame & "  INNER JOIN CLI_DATGEN A ON A.DATGEN_TIPDOC = F.SOLMAE_TITTDO AND A.DATGEN_NUMDOC = F.SOLMAE_TITNDO "
      g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES H ON H.PARDES_CODGRP = 203 AND H.PARDES_CODITE = A.DATGEN_TIPDOC "
      g_str_Parame = g_str_Parame & "  INNER JOIN TRA_SEGUIM B ON B.SEGUIM_NUMSOL = F.SOLMAE_NUMERO AND B.SEGUIM_CODINS = 21 AND B.SEGUIM_SITUAC = 1 "
      g_str_Parame = g_str_Parame & "   LEFT JOIN CLI_ACTECO C ON C.ACTECO_CLITDO = A.DATGEN_TIPDOC AND C.ACTECO_CLINDO = A.DATGEN_NUMDOC "
      g_str_Parame = g_str_Parame & "   LEFT JOIN MNT_PARDES D ON D.PARDES_CODGRP = 8 AND D.PARDES_CODITE = C.ACTECO_CODACT "
      g_str_Parame = g_str_Parame & "   LEFT JOIN MNT_PARDES E ON E.PARDES_CODGRP = 7 AND E.PARDES_CODITE = C.ActEco_OrdAct "
      g_str_Parame = g_str_Parame & "   LEFT JOIN CLI_DATGEN G ON G.DATGEN_TIPDOC = F.SOLMAE_CYGTDO AND G.DATGEN_NUMDOC = F.SOLMAE_CYGNDO "
      g_str_Parame = g_str_Parame & "   LEFT JOIN MNT_PARDES I ON I.PARDES_CODGRP = 203 AND I.PARDES_CODITE = G.DATGEN_TIPDOC "
      g_str_Parame = g_str_Parame & "   LEFT JOIN CLI_ACTECO J ON J.ACTECO_CLITDO = G.DATGEN_TIPDOC AND J.ACTECO_CLINDO = G.DATGEN_NUMDOC "
      g_str_Parame = g_str_Parame & "   LEFT JOIN MNT_PARDES K ON K.PARDES_CODGRP = 8 AND K.PARDES_CODITE = J.ACTECO_CODACT "
      g_str_Parame = g_str_Parame & "   LEFT JOIN MNT_PARDES L ON L.PARDES_CODGRP = 7 AND L.PARDES_CODITE = J.ActEco_OrdAct "
      g_str_Parame = g_str_Parame & "  WHERE F.SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "'"
   
      If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
          Exit Sub
      End If
      
      If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
         r_rst_Princi.MoveFirst
         crp_Imprim.Reset
         crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
         crp_Imprim.DataFiles(0) = "CRE_SOLMAE"
         crp_Imprim.DataFiles(1) = "TRA_EVACRE"
         
         crp_Imprim.ParameterFields(0) = "P_FecEva;" & Mid(r_rst_Princi!SEGUIM_FECFIN, 7, 2) & "/" & Mid(r_rst_Princi!SEGUIM_FECFIN, 5, 2) & "/" & Mid(r_rst_Princi!SEGUIM_FECFIN, 1, 4) & ";True"
         crp_Imprim.ParameterFields(1) = "P_ApeNom;" & r_rst_Princi!TIT_APE & " " & r_rst_Princi!TIT_NOM & "," & ";True"
         crp_Imprim.ParameterFields(2) = "P_ApeTit;" & r_rst_Princi!TIT_APE & ";True"
         crp_Imprim.ParameterFields(3) = "P_NomTit;" & r_rst_Princi!TIT_NOM & ";True"
         crp_Imprim.ParameterFields(4) = "P_TipDoc_Tit;" & r_rst_Princi!TIT_TIPDOC & ";True"
         crp_Imprim.ParameterFields(5) = "P_ApeCyg;" & r_rst_Princi!CYG_APE & ";True"
         crp_Imprim.ParameterFields(6) = "P_NomCyg;" & r_rst_Princi!CYG_NOM & ";True"
         crp_Imprim.ParameterFields(7) = "P_TipDoc_Cyg;" & r_rst_Princi!CYG_TIPDOC & ";True"
         crp_Imprim.ParameterFields(8) = "P_SitLab_Tit;" & r_rst_Princi!TIT_OCUPACION & ";True"
         crp_Imprim.ParameterFields(9) = "P_SitLab_Cyg;" & r_rst_Princi!CYG_OCUPACION & ";True"
            
         crp_Imprim.SelectionFormula = "{CRE_SOLMAE.SOLMAE_NUMERO} = '" & moddat_g_str_NumSol & "'" & " AND {TRA_EVACRE.EVACRE_NUMSOL} = '" & moddat_g_str_NumSol & "'"
         crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_CARAPR_17.RPT"
         crp_Imprim.WindowShowPrintSetupBtn = True
         crp_Imprim.Action = 1
      End If
      
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
Dim r_arr_Mtz()      As moddat_g_tpo_DatCom

   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt

   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_Situac.Caption = moddat_g_str_Situac
   
   Call fs_Inicia
   
   'Buscar Información de Solicitud de Crédito
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""

   Call modmip_gs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(0), 0)  'Buscar Información del Cliente
   Call modmip_gs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, grd_Listad(1), 1)  'Buscar Información del Cónyuge
   Call modmip_gs_DatApo(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(2))     'Buscar Información del Apoderado
   Call modmip_gs_DatInm(grd_Listad(3), False)                                        'Buscar Información del Inmueble
   
   'Buscar Información del Crédito
   Call modmip_gs_DatCre(grd_Listad(4), r_arr_Mtz)
   txt_ObsSol.Text = r_arr_Mtz(0).DatCom_Observ
   moddat_g_str_CodEjeSeg = r_arr_Mtz(0).DatCom_EjeSeg
   moddat_g_str_CodConHip = r_arr_Mtz(0).DatCom_ConHip
   moddat_g_str_FecIng = r_arr_Mtz(0).DatCom_FecSol
   l_str_EmpSeg = r_arr_Mtz(0).DatCom_EsgDes
   l_int_CodIns = r_arr_Mtz(0).DatCom_CodIns
   'Call fs_DatCre
   
   Call fs_Buscar_LisOcu      'Buscando Ocurrencias de Instancia
   Call fs_Buscar_LisExc      'Buscando Excepciones
   Call fs_Buscar_LisCon      'Buscando Aprobaciones Condicionadas
   Call fs_Buscar_InfMic      'Buscando Informacion microempresario

   'Si Situación no es Rechazado
   If moddat_g_int_Situac = 1 Then
      cmd_Imprim.Visible = True
      pnl_Situac.ForeColor = modgen_g_con_ColAzu
   ElseIf moddat_g_int_Situac = 2 Then
      cmd_Imprim.Visible = False
      pnl_Situac.ForeColor = modgen_g_con_ColVer
   ElseIf moddat_g_int_Situac = 3 Then
      cmd_Imprim.Visible = False
      pnl_Situac.ForeColor = modgen_g_con_ColRoj
   End If
   
   If moddat_g_int_Situac <> 1 Then
      cmd_Imprim.Visible = False
   End If
   
   'Si existe Rechazo en la Instancia
   If l_int_FlgRec = 0 Then
      tab_Seguim.TabVisible(3) = False
   End If
   
   'Si no hay Excepciones aplicadas
   If grd_LisExc.Rows = 0 Then
      tab_Seguim.TabVisible(1) = False
   End If

   'Si no hay Aprobaciones Condicionadas
   If grd_LisCon.Rows = 0 Then
      tab_Seguim.TabVisible(2) = False
   End If
   
   'Si no hay Aprobaciones Condicionadas Pendiente
   If l_int_AprCon = 0 Then
      pnl_AprCon.Visible = False
   End If
   
   Call fs_Buscar_EvaCre
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Dim r_int_Contad     As Integer

   'Inicializando Grid de Cliente y de Cónyuge
   For r_int_Contad = 0 To 4
      grd_Listad(r_int_Contad).ColWidth(0) = 2900:    grd_Listad(r_int_Contad).ColAlignment(0) = flexAlignLeftCenter
      grd_Listad(r_int_Contad).ColWidth(1) = 7950:    grd_Listad(r_int_Contad).ColAlignment(1) = flexAlignLeftCenter
      
      Call gs_LimpiaGrid(grd_Listad(r_int_Contad))
   Next r_int_Contad

   'Lista de Ocurrencias
   grd_LisOcu.ColWidth(0) = 1155
   grd_LisOcu.ColWidth(1) = 1185
   grd_LisOcu.ColWidth(2) = 8595
   grd_LisOcu.ColWidth(3) = 0
   grd_LisOcu.ColWidth(4) = 0
   grd_LisOcu.ColAlignment(0) = flexAlignCenterCenter
   grd_LisOcu.ColAlignment(1) = flexAlignCenterCenter
   grd_LisOcu.ColAlignment(2) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_LisOcu)

   pnl_DesOcu.Caption = ""
   txt_Observ.Text = ""
   txt_Descar.Text = ""

   'Lista de Excepciones
   grd_LisExc.ColWidth(0) = 1175
   grd_LisExc.ColWidth(1) = 1175
   grd_LisExc.ColWidth(2) = 3275
   grd_LisExc.ColWidth(3) = 5325
   grd_LisExc.ColWidth(4) = 0
   grd_LisExc.ColWidth(5) = 0
   grd_LisExc.ColAlignment(0) = flexAlignCenterCenter
   grd_LisExc.ColAlignment(1) = flexAlignCenterCenter
   grd_LisExc.ColAlignment(2) = flexAlignLeftCenter
   grd_LisExc.ColAlignment(3) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_LisExc)

   pnl_DesExc.Caption = ""
   txt_ObsExc.Text = ""
   pnl_TipAut.Caption = ""
   pnl_motivo.Caption = ""

   'Lista de Aprobaciones Condicionadas
   grd_LisCon.ColWidth(0) = 2735
   grd_LisCon.ColWidth(1) = 6605
   grd_LisCon.ColWidth(2) = 1625
   grd_LisCon.ColWidth(3) = 0
   grd_LisCon.ColAlignment(0) = flexAlignLeftCenter
   grd_LisCon.ColAlignment(1) = flexAlignLeftCenter
   grd_LisCon.ColAlignment(2) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_LisCon)

   pnl_InsCon.Caption = ""
   txt_ObsCon.Text = ""
   txt_LevCon.Text = ""

   'Lista de Evaluación
   grd_LisEva.ColWidth(0) = 3300
   grd_LisEva.ColWidth(1) = 7940
   grd_LisEva.ColAlignment(0) = flexAlignLeftCenter
   grd_LisEva.ColAlignment(1) = flexAlignLeftCenter
End Sub

Private Sub grd_LisCon_Click()
   Dim r_str_FecOcu     As String
   Dim r_str_HorOcu     As String
   Dim r_str_DesOcu     As String

   If grd_LisCon.Rows > 0 Then
      grd_LisCon.Col = 0
      pnl_InsCon.Caption = grd_LisCon.Text
      grd_LisCon.Col = 1
      txt_ObsCon.Text = grd_LisCon.Text
      grd_LisCon.Col = 3
      txt_LevCon.Text = grd_LisCon.Text
      
      Call gs_RefrescaGrid(grd_LisCon)
   End If
End Sub

Private Sub grd_LisCon_SelChange()
   If grd_LisCon.Rows > 2 Then
      grd_LisCon.RowSel = grd_LisCon.Row
   End If
   
   Call grd_LisCon_Click
End Sub

Private Sub grd_LisEva_SelChange()
   If grd_LisEva.Rows > 2 Then
      grd_LisEva.RowSel = grd_LisEva.Row
   End If
End Sub

Private Sub grd_LisExc_Click()
   Dim r_str_FecExc     As String
   Dim r_str_HorExc     As String
   Dim r_str_InsExc     As String

   If grd_LisExc.Rows > 0 Then
      grd_LisExc.Col = 0
      r_str_FecExc = grd_LisExc.Text
      
      grd_LisExc.Col = 1
      r_str_HorExc = grd_LisExc.Text
      
      grd_LisExc.Col = 2
      r_str_InsExc = grd_LisExc.Text
      
      pnl_DesExc.Caption = "Día: " & r_str_FecExc & " - " & r_str_HorExc & " hrs. - " & r_str_InsExc
   
      grd_LisExc.Col = 3
      txt_ObsExc.Text = grd_LisExc.Text
      
      grd_LisExc.Col = 4
      pnl_TipAut.Caption = grd_LisExc.Text
      
      If LCase(TRIM(r_str_InsExc)) = LCase("EVALUACION CREDITICIA") Then
         grd_LisExc.Col = 5
         pnl_motivo.Caption = IIf(grd_LisExc.Text = "0", " ", grd_LisExc.Text)
         pnl_motivo.Visible = True
         lbl_motivo.Visible = True
      Else
         pnl_motivo.Visible = False
         lbl_motivo.Visible = False
         pnl_motivo.Caption = ""
      End If
      
      Call gs_SetFocus(grd_LisExc)
      Call gs_RefrescaGrid(grd_LisExc)
   Else
      pnl_DesExc.Caption = ""
      txt_ObsExc.Text = ""
      pnl_TipAut.Caption = ""
   End If
End Sub

Private Sub grd_LisExc_SelChange()
   If grd_LisExc.Rows > 2 Then
      grd_LisExc.RowSel = grd_LisExc.Row
   End If
   
   Call grd_LisExc_Click
End Sub

Private Sub grd_LisOcu_Click()
   Dim r_str_FecOcu     As String
   Dim r_str_HorOcu     As String
   Dim r_str_DesOcu     As String

   If grd_LisOcu.Rows > 0 Then
      grd_LisOcu.Col = 0
      r_str_FecOcu = grd_LisOcu.Text
      
      grd_LisOcu.Col = 1
      r_str_HorOcu = grd_LisOcu.Text
      
      grd_LisOcu.Col = 2
      r_str_DesOcu = grd_LisOcu.Text
      
      pnl_DesOcu.Caption = "Día: " & r_str_FecOcu & " - " & r_str_HorOcu & " hrs. - " & r_str_DesOcu
   
      grd_LisOcu.Col = 3
      txt_Observ.Text = grd_LisOcu.Text
      
      grd_LisOcu.Col = 4
      txt_Descar.Text = grd_LisOcu.Text
      
      Call gs_RefrescaGrid(grd_LisOcu)
   End If
End Sub

Private Sub grd_LisOcu_SelChange()
   If grd_LisOcu.Rows > 2 Then
      grd_LisOcu.RowSel = grd_LisOcu.Row
   End If
   
   Call grd_LisOcu_Click
End Sub

Private Sub grd_Listad_SelChange(Index As Integer)
   If grd_Listad(Index).Rows > 2 Then
      grd_Listad(Index).RowSel = grd_Listad(Index).Row
   End If
End Sub

Private Sub fs_Buscar_LisOcu()
   l_int_FlgRec = 0
   Call gs_LimpiaGrid(grd_LisOcu)
   moddat_g_int_NumObs = 0
   
   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = 21 "
   g_str_Parame = g_str_Parame & "ORDER BY SEGFECCRE DESC, SEGHORCRE DESC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     Exit Sub
   End If
   
   grd_LisOcu.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_LisOcu.Rows = grd_LisOcu.Rows + 1
      grd_LisOcu.Row = grd_LisOcu.Rows - 1
      
      'Fecha de Ocurrencia
      grd_LisOcu.Col = 0
      grd_LisOcu.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECCRE))
      
      'Hora de Ocurrencia
      grd_LisOcu.Col = 1
      grd_LisOcu.Text = gf_FormatoHora(Format(g_rst_Princi!SEGHORCRE, "000000"))
      
      'Descripción Ocurrencia
      grd_LisOcu.Col = 2
      grd_LisOcu.Text = moddat_gf_Consulta_ParDes("004", Format(g_rst_Princi!SEGDET_CODOCU, "000000"))
      
      If g_rst_Princi!SEGDET_CODOCU = 21 Then
         If g_rst_Princi!SEGFECACT > 0 Then
            grd_LisOcu.Text = grd_LisOcu.Text & " (DESCARGO EFECTUADO - " & gf_FormatoFecha(CStr(g_rst_Princi!SEGFECACT))
            grd_LisOcu.Text = grd_LisOcu.Text & " / " & gf_FormatoHora(Format(g_rst_Princi!SEGHORACT, "000000")) & ")"
         Else
            moddat_g_int_NumObs = g_rst_Princi!SEGDET_NUMOBS
         End If
      ElseIf g_rst_Princi!SEGDET_CODOCU = 13 Then
         'Si la Solicitud está rechazada en la Instancia
         pnl_MotRec.Caption = moddat_gf_Consulta_ParDes("003", CStr(g_rst_Princi!SEGDET_MOTREC))
         txt_ObsRec.Text = TRIM(g_rst_Princi!SEGDET_OBSERV & "")
         
         l_int_FlgRec = 1
      End If
      
      grd_LisOcu.Col = 3
      grd_LisOcu.Text = TRIM(g_rst_Princi!SEGDET_OBSERV & "")
      
      grd_LisOcu.Col = 4
      grd_LisOcu.Text = TRIM(g_rst_Princi!SEGDET_OBSDES & "")
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_LisOcu.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_LisOcu)
   Call grd_LisOcu_Click
End Sub

Private Sub fs_Buscar_LisExc()
   Dim r_str_FecOcu  As String
   
   Call gs_LimpiaGrid(grd_LisExc)
   
   g_str_Parame = modgen_gf_Buscar_Excepc
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     Exit Sub
   End If
   
   grd_LisExc.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_LisExc.Rows = grd_LisExc.Rows + 1
      grd_LisExc.Row = grd_LisExc.Rows - 1
      
      'Fecha de Excepción
      grd_LisExc.Col = 0
      grd_LisExc.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECCRE))
      
      'Hora de Excepción
      grd_LisExc.Col = 1
      grd_LisExc.Text = gf_FormatoHora(Format(g_rst_Princi!SEGHORCRE, "000000"))
      
      'Instancia
      grd_LisExc.Col = 2
      grd_LisExc.Text = moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SEGEXC_CODINS))
      
      'Descripción Excepción
      grd_LisExc.Col = 3
      grd_LisExc.Text = TRIM(g_rst_Princi!SEGEXC_DESCRI & "")
      
      'Tipo Autorización
      grd_LisExc.Col = 4
      grd_LisExc.Text = moddat_gf_Consulta_ParDes("243", CStr(g_rst_Princi!SEGEXC_TIPAUT))
            
      'Motivo de Excepción
      grd_LisExc.Col = 5
      grd_LisExc.Text = TRIM(g_rst_Princi!PARDES_DESCRI)
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_LisExc.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_LisExc)
   Call grd_LisExc_Click
End Sub

Private Sub fs_Buscar_LisCon()
   l_int_AprCon = 0
   
   Call gs_LimpiaGrid(grd_LisCon)
   
   g_str_Parame = "SELECT * FROM TRA_SEGCON WHERE "
   g_str_Parame = g_str_Parame & "SEGCON_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & "ORDER BY SEGCON_SITUAC ASC, SEGCON_CODINS DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     Exit Sub
   End If
   
   grd_LisCon.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_LisCon.Rows = grd_LisCon.Rows + 1
      grd_LisCon.Row = grd_LisCon.Rows - 1
      
      'Instancia
      grd_LisCon.Col = 0
      grd_LisCon.Text = moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SEGCON_CODINS))
      
      'Descripción Condiciones
      grd_LisCon.Col = 1
      grd_LisCon.Text = TRIM(g_rst_Princi!SEGCON_OBSCON & "")
      
      'Situación
      grd_LisCon.Col = 2
      grd_LisCon.Text = moddat_gf_Consulta_ParDes("244", CStr(g_rst_Princi!SEGCON_SITUAC))
      
      If g_rst_Princi!SEGCON_SITUAC = 1 Then
         l_int_AprCon = 1
      End If
      
      'Descripción Levantamiento Condiciones
      grd_LisCon.Col = 3
      grd_LisCon.Text = TRIM(g_rst_Princi!SEGCON_OBSLEV & "")
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_LisCon.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_LisCon)
   Call grd_LisCon_Click
End Sub

Private Sub fs_Buscar_InfMic()
Dim r_str_Parame     As String
Dim r_rst_GenAux     As ADODB.Recordset

   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT A.ANXVAR_NUMSOL, ANXVAR_ANXITE1, ANXVAR_ANXITE2, ANXVAR_ANXITE3, ANXVAR_ANXITE4, ANXVAR_COMANX "
   r_str_Parame = r_str_Parame & "   FROM MIC_ANXVAR A "
   r_str_Parame = r_str_Parame & "  WHERE A.ANXVAR_NUMSOL = '" & moddat_g_str_NumSol & "'"
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_GenAux, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_GenAux.BOF And r_rst_GenAux.EOF) Then
      r_rst_GenAux.MoveFirst
      
      pnl_ArcItem_1.Caption = TRIM(r_rst_GenAux!ANXVAR_ANXITE1 & "")
      pnl_ArcItem_2.Caption = TRIM(r_rst_GenAux!ANXVAR_ANXITE2 & "")
      pnl_ArcItem_3.Caption = TRIM(r_rst_GenAux!ANXVAR_ANXITE3 & "")
      pnl_ArcItem_4.Caption = TRIM(r_rst_GenAux!ANXVAR_ANXITE4 & "")
      txt_Coment.Text = TRIM(r_rst_GenAux!ANXVAR_COMANX & "")
   End If
   
   r_rst_GenAux.Close
   Set r_rst_GenAux = Nothing
End Sub

Private Sub cmd_VerArc_1_Click()
   If TRIM(pnl_ArcItem_1.Caption) <> "" Then
      ShellExecute Me.hwnd, "Open", moddat_g_str_RutAnx & "\" & TRIM(pnl_ArcItem_1.Caption), "", "", 1
   End If
End Sub

Private Sub cmd_VerArc_2_Click()
   If TRIM(pnl_ArcItem_2.Caption) <> "" Then
      ShellExecute Me.hwnd, "Open", moddat_g_str_RutAnx & "\" & TRIM(pnl_ArcItem_2.Caption), "", "", 1
   End If
End Sub

Private Sub cmd_VerArc_3_Click()
   If TRIM(pnl_ArcItem_3.Caption) <> "" Then
      ShellExecute Me.hwnd, "Open", moddat_g_str_RutAnx & "\" & TRIM(pnl_ArcItem_3.Caption), "", "", 1
   End If
End Sub

Private Sub cmd_VerArc_4_Click()
   If TRIM(pnl_ArcItem_4.Caption) <> "" Then
      ShellExecute Me.hwnd, "Open", moddat_g_str_RutAnx & "\" & TRIM(pnl_ArcItem_4.Caption), "", "", 1
   End If
End Sub

Private Sub fs_Buscar_EvaCre()
   Call gs_LimpiaGrid(grd_LisEva)

   'Mostrar Todos los Documentos Recibidos
   g_str_Parame = "SELECT * FROM TRA_EVACRE WHERE "
   g_str_Parame = g_str_Parame & "EVACRE_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   'Llenando Grid
   If g_rst_Princi!EVACRE_INGTOT > 0 Then
      grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Total Ingreso Líquido"
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGTOT, 12, 2)
   
      grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Total Obligaciones Mensuales"
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_OBLMEN, 12, 2)
   
      grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Total Ingreso Neto"
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGNET, 12, 2)
   
      grd_LisEva.Rows = grd_LisEva.Rows + 2:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Cuota Máxima Aprob."
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CUOSOL, 12, 2)
   
      If moddat_g_int_TipMon <> 1 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1: grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                    grd_LisEva.Text = "Cuota Máximo Aprob. (M. Prest.)"
         grd_LisEva.Col = 1:                    grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:           grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!EVACRE_CUOMPR, 12, 2) & " (Tipo de Cambio: S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TCAING, 14, 4)
      End If
   End If
   
   If g_rst_Princi!EVACRE_FECCAL > 0 Then
      grd_LisEva.Rows = grd_LisEva.Rows + 2:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Monto Préstamo Aprobado"
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!EVACRE_MTOPRE_CAL, 12, 2)
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Plazo Aprobado"
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = CStr(g_rst_Princi!EVACRE_PLAANO_CAL) & " Años "
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Período de Gracia Aprobado"
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = CStr(g_rst_Princi!EVACRE_PERGRA_CAL) & IIf(g_rst_Princi!EVACRE_PERGRA_CAL = 1, " Mes", " Meses")
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Cuota Extraordinaria Aprobado"
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = moddat_gf_Consulta_ParDes("277", CStr(g_rst_Princi!EVACRE_CUODBL_CAL))
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Tipo de Seguro Aprobado"
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = moddat_gf_Consulta_TipSeg(l_str_EmpSeg, g_rst_Princi!EVACRE_TIPSEG_CAL)
      
      If moddat_g_int_TipMon <> 1 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1: grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                    grd_LisEva.Text = "Tipo Cambio de Aprobación"
         grd_LisEva.Col = 1:                    grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:           grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIPCAM, 14, 4)
      End If
   End If
   
   'Verificación Domiciliaria
   If Len(TRIM(g_rst_Princi!EVACRE_TIPVDM & "")) > 0 Then
      If g_rst_Princi!EVACRE_TIPVDM > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 2:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Verificación Domiciliaria"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = moddat_gf_Consulta_ParDes("067", CStr(g_rst_Princi!EVACRE_TIPVDM))
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
         grd_LisEva.Col = 1:                          grd_LisEva.Text = TRIM(g_rst_Princi!EVACRE_OBSVDM & "")
      End If
   End If
   
   'Central de Riesgo
   If Not IsNull(g_rst_Princi!EVACRE_TIT_CRIFLG) Then
      If g_rst_Princi!EVACRE_TIT_CRIFLG > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 2:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Central de Riesgo - Titular Reportado"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!EVACRE_TIT_CRIFLG))
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Fecha Reporte"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACRE_TIT_CRIFEC))
      End If
      
      If g_rst_Princi!EVACRE_TIT_CRIFLG = 1 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Nro. Entidades:"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = CStr(g_rst_Princi!EVACRE_TIT_CRIENT)
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Clasificación"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = "Nor: " & Format(g_rst_Princi!EVACRE_TIT_CRICL0, "##0.00") & "%" & " / " & "Cpp: " & Format(g_rst_Princi!EVACRE_TIT_CRICL1, "##0.00") & "%" & " / " & "Def: " & Format(g_rst_Princi!EVACRE_TIT_CRICL2, "##0.00") & "%" & " / " & "Dud: " & Format(g_rst_Princi!EVACRE_TIT_CRICL3, "##0.00") & "%" & " / " & "Per: " & Format(g_rst_Princi!EVACRE_TIT_CRICL4, "##0.00") & "%"
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Total Deuda MN"
         grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_TOTDMN, 12, 2)
      
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Total Deuda ME"
         grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_TOTDME, 12, 2)
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (1)"
         grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_NomEntFin(TRIM(g_rst_Princi!EVACRE_TIT_CODEN1) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_TIT_CLAEN1)) & ")"
         ' grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_ParDes("505", Trim(g_rst_Princi!EVACRE_TIT_CODEN1) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_TIT_CLAEN1)) & ")"

         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
         grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_DEUEN1, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!EVACRE_TIT_LIMDE1), 0, g_rst_Princi!EVACRE_TIT_LIMDE1), 12, 2)

         
         If Len(TRIM(g_rst_Princi!EVACRE_TIT_CODEN2 & "")) > 0 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (2)"
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_NomEntFin(TRIM(g_rst_Princi!EVACRE_TIT_CODEN2) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_TIT_CLAEN2)) & ")"

            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_DEUEN2, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!EVACRE_TIT_LIMDE2), 0, g_rst_Princi!EVACRE_TIT_LIMDE2), 12, 2)

         End If
      
         If Len(TRIM(g_rst_Princi!EVACRE_TIT_CODEN3 & "")) > 0 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (3)"
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_NomEntFin(TRIM(g_rst_Princi!EVACRE_TIT_CODEN3) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_TIT_CLAEN3)) & ")"
            
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_DEUEN3, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!EVACRE_TIT_LIMDE3), 0, g_rst_Princi!EVACRE_TIT_LIMDE3), 12, 2)

         End If
      
         If Len(TRIM(g_rst_Princi!EVACRE_TIT_CODEN4 & "")) > 0 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (4)"
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_NomEntFin(TRIM(g_rst_Princi!EVACRE_TIT_CODEN4) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_TIT_CLAEN4)) & ")"
         
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_DEUEN4, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!EVACRE_TIT_LIMDE4), 0, g_rst_Princi!EVACRE_TIT_LIMDE4), 12, 2)

         End If
      
         If Len(TRIM(g_rst_Princi!EVACRE_TIT_CODEN5 & "")) > 0 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (5)"
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_NomEntFin(TRIM(g_rst_Princi!EVACRE_TIT_CODEN5) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_TIT_CLAEN5)) & ")"
         
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_DEUEN5, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!EVACRE_TIT_LIMDE5), 0, g_rst_Princi!EVACRE_TIT_LIMDE5), 12, 2)

         End If
      
         If Len(TRIM(g_rst_Princi!EVACRE_TIT_CODEN6 & "")) > 0 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (6)"
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_NomEntFin(TRIM(g_rst_Princi!EVACRE_TIT_CODEN6) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_TIT_CLAEN6)) & ")"
         
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_DEUEN6, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!EVACRE_TIT_LIMDE6), 0, g_rst_Princi!EVACRE_TIT_LIMDE6), 12, 2)

         End If
      End If
      
      'Central de Riesgo (Cónyuge)
      If g_rst_Princi!EVACRE_CYG_CRIFLG > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 2:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Central de Riesgo - Cónyuge Reportado"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!EVACRE_CYG_CRIFLG))
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Fecha Reporte"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACRE_CYG_CRIFEC))
         
         If g_rst_Princi!EVACRE_CYG_CRIFLG = 1 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Nro. Entidades:"
            grd_LisEva.Col = 1:                          grd_LisEva.Text = CStr(g_rst_Princi!EVACRE_CYG_CRIENT)
            
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Clasificación"
            grd_LisEva.Col = 1:                          grd_LisEva.Text = "Nor: " & Format(g_rst_Princi!EVACRE_CYG_CRICL0, "##0.00") & "%" & " / " & "Cpp: " & Format(g_rst_Princi!EVACRE_CYG_CRICL1, "##0.00") & "%" & " / " & "Def: " & Format(g_rst_Princi!EVACRE_CYG_CRICL2, "##0.00") & "%" & " / " & "Dud: " & Format(g_rst_Princi!EVACRE_CYG_CRICL3, "##0.00") & "%" & " / " & "Per: " & Format(g_rst_Princi!EVACRE_CYG_CRICL4, "##0.00") & "%"
            
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Total Deuda MN"
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_TOTDMN, 12, 2)
         
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Total Deuda ME"
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_TOTDME, 12, 2)
            
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (1)"
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_NomEntFin(TRIM(g_rst_Princi!EVACRE_CYG_CODEN1) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_CYG_CLAEN1)) & ")"
            
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_DEUEN1, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!EVACRE_CYG_LIMDE1), 0, g_rst_Princi!EVACRE_CYG_LIMDE1), 12, 2)

            
            If Len(TRIM(g_rst_Princi!EVACRE_CYG_CODEN2 & "")) > 0 Then
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (2)"
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_NomEntFin(TRIM(g_rst_Princi!EVACRE_CYG_CODEN2) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_CYG_CLAEN2)) & ")"
               
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_DEUEN2, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!EVACRE_CYG_LIMDE2), 0, g_rst_Princi!EVACRE_CYG_LIMDE2), 12, 2)

            End If
         
            If Len(TRIM(g_rst_Princi!EVACRE_CYG_CODEN3 & "")) > 0 Then
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (3)"
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_NomEntFin(TRIM(g_rst_Princi!EVACRE_CYG_CODEN3) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_CYG_CLAEN3)) & ")"
               
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_DEUEN3, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!EVACRE_CYG_LIMDE3), 0, g_rst_Princi!EVACRE_CYG_LIMDE3), 12, 2)

            End If
         
            If Len(TRIM(g_rst_Princi!EVACRE_CYG_CODEN4 & "")) > 0 Then
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (4)"
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_NomEntFin(TRIM(g_rst_Princi!EVACRE_CYG_CODEN4) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_CYG_CLAEN4)) & ")"
               
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_DEUEN4, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!EVACRE_CYG_LIMDE4), 0, g_rst_Princi!EVACRE_CYG_LIMDE4), 12, 2)

            End If
         
            If Len(TRIM(g_rst_Princi!EVACRE_CYG_CODEN5 & "")) > 0 Then
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (5)"
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_NomEntFin(TRIM(g_rst_Princi!EVACRE_CYG_CODEN5) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_CYG_CLAEN5)) & ")"
               
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_DEUEN5, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!EVACRE_CYG_LIMDE5), 0, g_rst_Princi!EVACRE_CYG_LIMDE5), 12, 2)

            End If
         
            If Len(TRIM(g_rst_Princi!EVACRE_CYG_CODEN6 & "")) > 0 Then
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (6)"
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_NomEntFin(TRIM(g_rst_Princi!EVACRE_CYG_CODEN6) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_CYG_CLAEN6)) & ")"
               
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_DEUEN6, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!EVACRE_CYG_LIMDE6), 0, g_rst_Princi!EVACRE_CYG_LIMDE6), 12, 2)

            End If
         End If
      End If
   End If
   
   'Referencias Personales
   If Len(TRIM(g_rst_Princi!EVACRE_REFPER & "")) > 0 Then
      grd_LisEva.Rows = grd_LisEva.Rows + 2:       grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                          grd_LisEva.Text = "Verificación Referencias"
      grd_LisEva.Col = 1:                          grd_LisEva.Text = TRIM(g_rst_Princi!EVACRE_REFPER & "")
   End If
   
   'Actividades Económicas
   If Not IsNull(g_rst_Princi!EVACRE_TIT_TIPVE1) Then
      If g_rst_Princi!EVACRE_TIT_TIPVE1 > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 2:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Titular - Verif. Lab. (Act. Princ.)"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = moddat_gf_Consulta_ParDes("068", g_rst_Princi!EVACRE_TIT_TIPVE1)
      
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 1:                          grd_LisEva.Text = TRIM(g_rst_Princi!EVACRE_TIT_LABVE1 & "")
      End If
      
      If g_rst_Princi!EVACRE_TIT_TIPVE2 > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Titular - Verif. Lab. (Act. Secund.)"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = moddat_gf_Consulta_ParDes("068", g_rst_Princi!EVACRE_TIT_TIPVE2)
      
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 1:                          grd_LisEva.Text = TRIM(g_rst_Princi!EVACRE_TIT_LABVE2 & "")
      End If
   
      If g_rst_Princi!EVACRE_CYG_TIPVE1 > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 2:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Cónyuge - Verif. Lab. (Act. Princ.)"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = moddat_gf_Consulta_ParDes("068", g_rst_Princi!EVACRE_CYG_TIPVE1)
      
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 1:                          grd_LisEva.Text = TRIM(g_rst_Princi!EVACRE_CYG_LABVE1 & "")
      End If
   
      If g_rst_Princi!EVACRE_CYG_TIPVE2 > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Cónyuge - Verif. Lab. (Act. Secund.)"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = moddat_gf_Consulta_ParDes("068", g_rst_Princi!EVACRE_CYG_TIPVE2)
      
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 1:                          grd_LisEva.Text = TRIM(g_rst_Princi!EVACRE_CYG_LABVE2 & "")
      End If
   End If
   
   'Ingresos
   If g_rst_Princi!EVACRE_INGDP1 > 0 Then
      grd_LisEva.Rows = grd_LisEva.Rows + 2:       grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                          grd_LisEva.Text = "Ingreso Líquido Titular"
      grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGDP1, 12, 2) & " + " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGDP2, 12, 2) & " + " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGAD1, 12, 2) & " = " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGDP1 + g_rst_Princi!EVACRE_INGDP2 + g_rst_Princi!EVACRE_INGAD1, 12, 2)
   
      grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                          grd_LisEva.Text = "Obligaciones Mensuales Titular"
      grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_OMETIT, 12, 2)
   
      grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                          grd_LisEva.Text = "Ingreso Neto Titular"
      grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_INTTIT, 12, 2)
   
      If g_rst_Princi!EVACRE_INGDP3 > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                       grd_LisEva.Text = "Ingreso Líquido Cónyuge"
         grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGDP3, 12, 2) & " + " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGDP4, 12, 2) & " + " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGAD2, 12, 2) & " = " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGDP3 + g_rst_Princi!EVACRE_INGDP4 + g_rst_Princi!EVACRE_INGAD2, 12, 2)
      End If
   
      If g_rst_Princi!EVACRE_OMECYG > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                       grd_LisEva.Text = "Obligaciones Mensuales Cónyuge"
         grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_OMECYG, 12, 2)
      End If
      
      If g_rst_Princi!EVACRE_INTCYG > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                       grd_LisEva.Text = "Ingreso Neto Cónyuge"
         grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_INTCYG, 12, 2)
      End If
   
      grd_LisEva.Rows = grd_LisEva.Rows + 2:       grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                          grd_LisEva.Text = "Total Deuda"
      grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_MTODEU, 12, 2)
   
      grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                          grd_LisEva.Text = "Ratio Ingreso / Deuda"
      grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVACRE_RINGDE, 12, 2)
   
      grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                          grd_LisEva.Text = "Ratio Inicial / Deuda"
      grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVACRE_RINIDE, 12, 2) & "%"
   End If
   
   Call gs_UbiIniGrid(grd_LisEva)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub txt_Coment_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_Descar_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_LevCon_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_ObsCon_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_ObsExc_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_ObsRec_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_ObsSol_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

