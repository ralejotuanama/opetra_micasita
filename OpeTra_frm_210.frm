VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_Con_SolHip_54 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10125
   ClientLeft      =   825
   ClientTop       =   1455
   ClientWidth     =   11595
   Icon            =   "OpeTra_frm_210.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10125
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10125
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
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
         TabIndex        =   1
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
            Height          =   3375
            Left            =   60
            TabIndex        =   2
            Top             =   60
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   5953
            _Version        =   393216
            Style           =   1
            Tabs            =   5
            TabsPerRow      =   5
            TabHeight       =   520
            TabCaption(0)   =   "Seguimiento en Instancia"
            TabPicture(0)   =   "OpeTra_frm_210.frx":000C
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label7"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label8"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label11"
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
            Tab(0).Control(9)=   "txt_Observ"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "txt_Descar"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).ControlCount=   11
            TabCaption(1)   =   "Excepciones Aplicadas"
            TabPicture(1)   =   "OpeTra_frm_210.frx":0028
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
            Tab(1).Control(11)=   "Label4"
            Tab(1).Control(12)=   "Label3"
            Tab(1).Control(13)=   "Label2"
            Tab(1).ControlCount=   14
            TabCaption(2)   =   "Aprobaci�n Condicionada"
            TabPicture(2)   =   "OpeTra_frm_210.frx":0044
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "txt_LevCon"
            Tab(2).Control(1)=   "txt_ObsCon"
            Tab(2).Control(2)=   "SSPanel17"
            Tab(2).Control(3)=   "grd_LisCon"
            Tab(2).Control(4)=   "SSPanel18"
            Tab(2).Control(5)=   "SSPanel19"
            Tab(2).Control(6)=   "SSPanel20"
            Tab(2).Control(7)=   "pnl_InsCon"
            Tab(2).Control(8)=   "Label12"
            Tab(2).Control(9)=   "Label14"
            Tab(2).Control(10)=   "Label15"
            Tab(2).ControlCount=   11
            TabCaption(3)   =   "Rechazo de Solicitud"
            TabPicture(3)   =   "OpeTra_frm_210.frx":0060
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "txt_ObsRec"
            Tab(3).Control(1)=   "pnl_MotRec"
            Tab(3).Control(2)=   "Label10"
            Tab(3).Control(3)=   "Label19"
            Tab(3).ControlCount=   4
            TabCaption(4)   =   "Informaci�n Microempresario"
            TabPicture(4)   =   "OpeTra_frm_210.frx":007C
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "SSPanel23"
            Tab(4).Control(0).Enabled=   0   'False
            Tab(4).Control(1)=   "SSPanel22"
            Tab(4).Control(1).Enabled=   0   'False
            Tab(4).ControlCount=   2
            Begin VB.TextBox txt_ObsRec 
               Height          =   2595
               Left            =   -73560
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   8
               Top             =   690
               Width           =   9915
            End
            Begin VB.TextBox txt_LevCon 
               Height          =   645
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   7
               Text            =   "OpeTra_frm_210.frx":0098
               Top             =   2640
               Width           =   10005
            End
            Begin VB.TextBox txt_ObsCon 
               Height          =   645
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   6
               Text            =   "OpeTra_frm_210.frx":009C
               Top             =   1980
               Width           =   10005
            End
            Begin VB.TextBox txt_ObsExc 
               Height          =   975
               Left            =   -73770
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   5
               Text            =   "OpeTra_frm_210.frx":00A0
               Top             =   1980
               Width           =   10065
            End
            Begin VB.TextBox txt_Descar 
               Height          =   645
               Left            =   1320
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   4
               Text            =   "OpeTra_frm_210.frx":00A6
               Top             =   2640
               Width           =   10005
            End
            Begin VB.TextBox txt_Observ 
               Height          =   645
               Left            =   1320
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   3
               Text            =   "OpeTra_frm_210.frx":00AA
               Top             =   1980
               Width           =   10005
            End
            Begin Threed.SSPanel SSPanel10 
               Height          =   45
               Left            =   30
               TabIndex        =   9
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
               TabIndex        =   10
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
               TabIndex        =   11
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
               TabIndex        =   12
               Top             =   360
               Width           =   8595
               _Version        =   65536
               _ExtentX        =   15161
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Descripci�n Ocurrencia"
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
               TabIndex        =   13
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
               TabIndex        =   14
               Top             =   1650
               Width           =   10005
               _Version        =   65536
               _ExtentX        =   17648
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "D�a: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
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
               TabIndex        =   15
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
               TabIndex        =   16
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Excepci�n"
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
               TabIndex        =   17
               Top             =   360
               Width           =   5325
               _Version        =   65536
               _ExtentX        =   9393
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Descripci�n Excepci�n"
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
               TabIndex        =   18
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "H. Excepci�n"
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
               TabIndex        =   19
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
               TabIndex        =   20
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
               TabIndex        =   21
               Top             =   1650
               Width           =   10050
               _Version        =   65536
               _ExtentX        =   17727
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "D�a: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
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
               TabIndex        =   22
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
               TabIndex        =   23
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
               TabIndex        =   24
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
               TabIndex        =   25
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
               TabIndex        =   26
               Top             =   360
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Situaci�n"
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
               TabIndex        =   27
               Top             =   360
               Width           =   6615
               _Version        =   65536
               _ExtentX        =   11668
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Condiciones de Aprobaci�n"
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
               TabIndex        =   28
               Top             =   1650
               Width           =   10005
               _Version        =   65536
               _ExtentX        =   17648
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "D�a: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
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
               TabIndex        =   29
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
               Left            =   -68880
               TabIndex        =   66
               Top             =   2970
               Width           =   5205
               _Version        =   65536
               _ExtentX        =   9181
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
               Height          =   1575
               Left            =   -74940
               TabIndex        =   68
               Top             =   410
               Width           =   11280
               _Version        =   65536
               _ExtentX        =   19897
               _ExtentY        =   2787
               _StockProps     =   15
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.28
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   1
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
                  TabIndex        =   83
                  Top             =   1140
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
                  TabIndex        =   81
                  Top             =   480
                  Width           =   465
               End
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
                  TabIndex        =   80
                  Top             =   150
                  Width           =   465
               End
               Begin Threed.SSPanel pnl_ArcItem_3 
                  Height          =   315
                  Left            =   2130
                  TabIndex        =   69
                  Top             =   810
                  Width           =   8325
                  _Version        =   65536
                  _ExtentX        =   14684
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
               Begin Threed.SSPanel pnl_ArcItem_4 
                  Height          =   315
                  Left            =   2130
                  TabIndex        =   70
                  Top             =   1140
                  Width           =   8325
                  _Version        =   65536
                  _ExtentX        =   14684
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
               Begin Threed.SSPanel pnl_ArcItem_2 
                  Height          =   315
                  Left            =   2130
                  TabIndex        =   71
                  Top             =   480
                  Width           =   8325
                  _Version        =   65536
                  _ExtentX        =   14684
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
               Begin Threed.SSPanel pnl_ArcItem_1 
                  Height          =   315
                  Left            =   2130
                  TabIndex        =   72
                  Top             =   150
                  Width           =   8325
                  _Version        =   65536
                  _ExtentX        =   14684
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
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  Caption         =   "Hoja de Trabajo N�2:"
                  Height          =   195
                  Left            =   150
                  TabIndex        =   76
                  Top             =   1200
                  Width           =   1500
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  Caption         =   "Hoja de Trabajo N�1:"
                  Height          =   195
                  Left            =   150
                  TabIndex        =   75
                  Top             =   870
                  Width           =   1500
               End
               Begin VB.Label Label16 
                  AutoSize        =   -1  'True
                  Caption         =   "Flujo de Caja:"
                  Height          =   195
                  Left            =   150
                  TabIndex        =   74
                  Top             =   540
                  Width           =   960
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
            End
            Begin Threed.SSPanel SSPanel23 
               Height          =   1275
               Left            =   -74940
               TabIndex        =   77
               Top             =   2025
               Width           =   11280
               _Version        =   65536
               _ExtentX        =   19897
               _ExtentY        =   2240
               _StockProps     =   15
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.22
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
                  TabIndex        =   78
                  Top             =   150
                  Width           =   8835
               End
               Begin VB.Label Label18 
                  AutoSize        =   -1  'True
                  Caption         =   "Comentario:"
                  Height          =   195
                  Left            =   150
                  TabIndex        =   79
                  Top             =   510
                  Width           =   840
               End
            End
            Begin VB.Label lbl_motivo 
               Caption         =   "Motivo:"
               Height          =   255
               Left            =   -69630
               TabIndex        =   65
               Top             =   3030
               Width           =   675
            End
            Begin VB.Label Label10 
               Caption         =   "Motivo Rechazo:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   40
               Top             =   360
               Width           =   1275
            End
            Begin VB.Label Label19 
               Caption         =   "Observaciones de Rechazo:"
               Height          =   555
               Left            =   -74940
               TabIndex        =   39
               Top             =   690
               Width           =   1155
            End
            Begin VB.Label Label12 
               Caption         =   "Levantamiento de Condiciones:"
               Height          =   615
               Left            =   -74940
               TabIndex        =   38
               Top             =   2640
               Width           =   1215
            End
            Begin VB.Label Label14 
               Caption         =   "Instancia:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   37
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label15 
               Caption         =   "Condiciones de Aprobaci�n:"
               Height          =   495
               Left            =   -74940
               TabIndex        =   36
               Top             =   1980
               Width           =   1155
            End
            Begin VB.Label Label4 
               Caption         =   "Descripci�n:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   35
               Top             =   2040
               Width           =   1155
            End
            Begin VB.Label Label3 
               Caption         =   "Excepci�n:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   34
               Top             =   1680
               Width           =   1155
            End
            Begin VB.Label Label2 
               Caption         =   "Autorizado por:"
               Height          =   255
               Left            =   -74940
               TabIndex        =   33
               Top             =   3030
               Width           =   1095
            End
            Begin VB.Label Label11 
               Caption         =   "Descargo:"
               Height          =   315
               Left            =   60
               TabIndex        =   32
               Top             =   2640
               Width           =   1035
            End
            Begin VB.Label Label8 
               Caption         =   "Ocurrencia:"
               Height          =   315
               Left            =   60
               TabIndex        =   31
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label7 
               Caption         =   "Comentario u Observaci�n:"
               Height          =   495
               Left            =   60
               TabIndex        =   30
               Top             =   1980
               Width           =   1155
            End
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   2325
         Left            =   30
         TabIndex        =   41
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
            TabIndex        =   42
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
            TabPicture(0)   =   "OpeTra_frm_210.frx":00AE
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "C�nyuge"
            TabPicture(1)   =   "OpeTra_frm_210.frx":00CA
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Apoderado"
            TabPicture(2)   =   "OpeTra_frm_210.frx":00E6
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(2)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Inmueble"
            TabPicture(3)   =   "OpeTra_frm_210.frx":0102
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "grd_Listad(3)"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Datos del Cr�dito"
            TabPicture(4)   =   "OpeTra_frm_210.frx":011E
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
               TabIndex        =   43
               Text            =   "OpeTra_frm_210.frx":013A
               Top             =   1470
               Width           =   10005
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1815
               Index           =   0
               Left            =   60
               TabIndex        =   44
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
               Index           =   3
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
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1095
               Index           =   4
               Left            =   -74940
               TabIndex        =   48
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
               TabIndex        =   49
               Top             =   1470
               Width           =   1155
            End
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   50
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
            TabIndex        =   51
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
            TabIndex        =   52
            Top             =   30
            Width           =   5685
            _Version        =   65536
            _ExtentX        =   10028
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Consulta de Solicitud de Cr�dito Hipotecario"
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
            TabIndex        =   53
            Top             =   330
            Width           =   5505
            _Version        =   65536
            _ExtentX        =   9710
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Evaluaci�n Crediticia"
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
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_210.frx":013E
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   54
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
            TabIndex        =   55
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
            TabIndex        =   56
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
            TabIndex        =   57
            Top             =   30
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
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
         Begin VB.Label Label6 
            Caption         =   "Situaci�n:"
            Height          =   315
            Left            =   6210
            TabIndex        =   60
            Top             =   30
            Width           =   1005
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   59
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   58
            Top             =   60
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel39 
         Height          =   645
         Left            =   30
         TabIndex        =   61
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
            Picture         =   "OpeTra_frm_210.frx":0448
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel16 
         Height          =   1935
         Left            =   30
         TabIndex        =   63
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
            TabIndex        =   64
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
Attribute VB_Name = "frm_Con_SolHip_54"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_AprCon        As Integer
Dim l_int_FlgRec        As Integer
Dim l_str_EmpSeg        As String
Dim l_int_CodIns        As Integer

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
   
   'Buscar Informaci�n de Solicitud de Cr�dito
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""

   Call modmip_gs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(0), 0)      'Buscar Informaci�n del Cliente
   Call modmip_gs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, grd_Listad(1), 1)      'Buscar Informaci�n del C�nyuge
   Call modmip_gs_DatApo(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(2))         'Buscar Informaci�n del Apoderado
   Call modmip_gs_DatInm(grd_Listad(3), False)                                                   'Buscar Informaci�n del Inmueble
       
   'Datos del Cr�dito
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
   
   'Si Situaci�n no es Rechazado
   If moddat_g_int_Situac = 1 Then
      pnl_Situac.ForeColor = modgen_g_con_ColAzu
   ElseIf moddat_g_int_Situac = 2 Then
      pnl_Situac.ForeColor = modgen_g_con_ColVer
   ElseIf moddat_g_int_Situac = 3 Then
      pnl_Situac.ForeColor = modgen_g_con_ColRoj
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

   'Inicializando Grid de Cliente y de C�nyuge
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
   pnl_Motivo.Caption = ""

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

   'Lista de Evaluaci�n
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
      
      pnl_DesExc.Caption = "D�a: " & r_str_FecExc & " - " & r_str_HorExc & " hrs. - " & r_str_InsExc
   
      grd_LisExc.Col = 3
      txt_ObsExc.Text = grd_LisExc.Text
      
      grd_LisExc.Col = 4
      pnl_TipAut.Caption = grd_LisExc.Text
        
      If LCase(Trim(r_str_InsExc)) = LCase("EVALUACION CREDITICIA") Then
         grd_LisExc.Col = 5
         pnl_Motivo.Caption = IIf(grd_LisExc.Text = "0", " ", grd_LisExc.Text)
         pnl_Motivo.Visible = True
         lbl_motivo.Visible = True
      Else
         pnl_Motivo.Visible = False
         lbl_motivo.Visible = False
         pnl_Motivo.Caption = ""
      End If
      
      Call gs_SetFocus(grd_LisExc)
      Call gs_RefrescaGrid(grd_LisExc)
   Else
      pnl_DesExc.Caption = ""
      txt_ObsExc.Text = ""
      pnl_TipAut.Caption = ""
      pnl_Motivo.Caption = ""
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
      
      pnl_DesOcu.Caption = "D�a: " & r_str_FecOcu & " - " & r_str_HorOcu & " hrs. - " & r_str_DesOcu
   
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
      
      'Descripci�n Ocurrencia
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
         'Si la Solicitud est� rechazada en la Instancia
         pnl_MotRec.Caption = moddat_gf_Consulta_ParDes("003", CStr(g_rst_Princi!SEGDET_MOTREC))
         txt_ObsRec.Text = Trim(g_rst_Princi!SEGDET_OBSERV & "")
         l_int_FlgRec = 1
      End If
      
      grd_LisOcu.Col = 3
      grd_LisOcu.Text = Trim(g_rst_Princi!SEGDET_OBSERV & "")
      
      grd_LisOcu.Col = 4
      grd_LisOcu.Text = Trim(g_rst_Princi!SEGDET_OBSDES & "")
      
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
      
      'Fecha de Excepci�n
      grd_LisExc.Col = 0
      grd_LisExc.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECCRE))
      
      'Hora de Excepci�n
      grd_LisExc.Col = 1
      grd_LisExc.Text = gf_FormatoHora(Format(g_rst_Princi!SEGHORCRE, "000000"))
      
      'Instancia
      grd_LisExc.Col = 2
      grd_LisExc.Text = moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SEGEXC_CODINS))
      
      'Descripci�n Excepci�n
      grd_LisExc.Col = 3
      grd_LisExc.Text = Trim(g_rst_Princi!SEGEXC_DESCRI & "")
      
      'Tipo Autorizaci�n
      grd_LisExc.Col = 4
      grd_LisExc.Text = moddat_gf_Consulta_ParDes("243", CStr(g_rst_Princi!SEGEXC_TIPAUT))
      
      'Motivo de Excepci�n
      grd_LisExc.Col = 5
      grd_LisExc.Text = Trim(g_rst_Princi!PARDES_DESCRI)
      
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
      
      'Descripci�n Condiciones
      grd_LisCon.Col = 1
      grd_LisCon.Text = Trim(g_rst_Princi!SEGCON_OBSCON & "")
      
      'Situaci�n
      grd_LisCon.Col = 2
      grd_LisCon.Text = moddat_gf_Consulta_ParDes("244", CStr(g_rst_Princi!SEGCON_SITUAC))
      
      If g_rst_Princi!SEGCON_SITUAC = 1 Then
         l_int_AprCon = 1
      End If
      
      'Descripci�n Levantamiento Condiciones
      grd_LisCon.Col = 3
      grd_LisCon.Text = Trim(g_rst_Princi!SEGCON_OBSLEV & "")
      
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
      
      pnl_ArcItem_1.Caption = Trim(r_rst_GenAux!ANXVAR_ANXITE1 & "")
      pnl_ArcItem_2.Caption = Trim(r_rst_GenAux!ANXVAR_ANXITE2 & "")
      pnl_ArcItem_3.Caption = Trim(r_rst_GenAux!ANXVAR_ANXITE3 & "")
      pnl_ArcItem_4.Caption = Trim(r_rst_GenAux!ANXVAR_ANXITE4 & "")
      txt_Coment.Text = Trim(r_rst_GenAux!ANXVAR_COMANX & "")
   End If
   
   r_rst_GenAux.Close
   Set r_rst_GenAux = Nothing
End Sub

Private Sub cmd_VerArc_1_Click()
   If Trim(pnl_ArcItem_1.Caption) <> "" Then
      ShellExecute Me.hwnd, "Open", moddat_g_str_RutAnx & "\" & Trim(pnl_ArcItem_1.Caption), "", "", 1
   End If
End Sub

Private Sub cmd_VerArc_2_Click()
   If Trim(pnl_ArcItem_2.Caption) <> "" Then
      ShellExecute Me.hwnd, "Open", moddat_g_str_RutAnx & "\" & Trim(pnl_ArcItem_2.Caption), "", "", 1
   End If
End Sub

Private Sub cmd_VerArc_3_Click()
   If Trim(pnl_ArcItem_3.Caption) <> "" Then
      ShellExecute Me.hwnd, "Open", moddat_g_str_RutAnx & "\" & Trim(pnl_ArcItem_3.Caption), "", "", 1
   End If
End Sub

Private Sub cmd_VerArc_4_Click()
   If Trim(pnl_ArcItem_4.Caption) <> "" Then
      ShellExecute Me.hwnd, "Open", moddat_g_str_RutAnx & "\" & Trim(pnl_ArcItem_4.Caption), "", "", 1
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
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Total Ingreso L�quido"
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
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Cuota M�xima Aprob."
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CUOSOL, 12, 2)
   
      If moddat_g_int_TipMon <> 1 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1: grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                    grd_LisEva.Text = "Cuota M�ximo Aprob. (M. Prest.)"
         grd_LisEva.Col = 1:                    grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:           grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!EVACRE_CUOMPR, 12, 2) & " (Tipo de Cambio: S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TCAING, 14, 4)
      End If
   End If
   
   If g_rst_Princi!EVACRE_FECCAL > 0 Then
      grd_LisEva.Rows = grd_LisEva.Rows + 2:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Monto Pr�stamo Aprobado"
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!EVACRE_MTOPRE_CAL, 12, 2)
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Plazo Aprobado"
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = CStr(g_rst_Princi!EVACRE_PLAANO_CAL) & " A�os "
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Per�odo de Gracia Aprobado"
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = CStr(g_rst_Princi!EVACRE_PERGRA_CAL) & IIf(g_rst_Princi!EVACRE_PERGRA_CAL = 1, " Mes", " Meses")
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Tipo de Seguro Aprobado"
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = moddat_gf_Consulta_TipSeg(l_str_EmpSeg, g_rst_Princi!EVACRE_TIPSEG_CAL)
      
      If moddat_g_int_TipMon <> 1 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1: grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                    grd_LisEva.Text = "Tipo Cambio de Aprobaci�n"
         grd_LisEva.Col = 1:                    grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:           grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIPCAM, 14, 4)
      End If
   End If
   
   'Verificaci�n Domiciliaria
   If Len(Trim(g_rst_Princi!EVACRE_TIPVDM & "")) > 0 Then
      If g_rst_Princi!EVACRE_TIPVDM > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 2:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Verificaci�n Domiciliaria"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = moddat_gf_Consulta_ParDes("067", CStr(g_rst_Princi!EVACRE_TIPVDM))
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
         grd_LisEva.Col = 1:                          grd_LisEva.Text = Trim(g_rst_Princi!EVACRE_OBSVDM & "")
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
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Clasificaci�n"
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
         grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_NomEntFin(Trim(g_rst_Princi!EVACRE_TIT_CODEN1) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_TIT_CLAEN1)) & ")"
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
         grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_DEUEN1, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!EVACRE_TIT_LIMDE1), 0, g_rst_Princi!EVACRE_TIT_LIMDE1), 12, 2)


         If Len(Trim(g_rst_Princi!EVACRE_TIT_CODEN2 & "")) > 0 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (2)"
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_NomEntFin(Trim(g_rst_Princi!EVACRE_TIT_CODEN2) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_TIT_CLAEN2)) & ")"
            
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_DEUEN2, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!EVACRE_TIT_LIMDE2), 0, g_rst_Princi!EVACRE_TIT_LIMDE2), 12, 2)

         End If
      
         If Len(Trim(g_rst_Princi!EVACRE_TIT_CODEN3 & "")) > 0 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (3)"
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_NomEntFin(Trim(g_rst_Princi!EVACRE_TIT_CODEN3) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_TIT_CLAEN3)) & ")"
         
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_DEUEN3, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!EVACRE_TIT_LIMDE3), 0, g_rst_Princi!EVACRE_TIT_LIMDE3), 12, 2)

         End If
      
         If Len(Trim(g_rst_Princi!EVACRE_TIT_CODEN4 & "")) > 0 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (4)"
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_NomEntFin(Trim(g_rst_Princi!EVACRE_TIT_CODEN4) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_TIT_CLAEN4)) & ")"
            
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_DEUEN4, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!EVACRE_TIT_LIMDE4), 0, g_rst_Princi!EVACRE_TIT_LIMDE4), 12, 2)
         
         End If
      
         If Len(Trim(g_rst_Princi!EVACRE_TIT_CODEN5 & "")) > 0 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (5)"
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_NomEntFin(Trim(g_rst_Princi!EVACRE_TIT_CODEN5) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_TIT_CLAEN5)) & ")"
         
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_DEUEN5, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!EVACRE_TIT_LIMDE5), 0, g_rst_Princi!EVACRE_TIT_LIMDE5), 12, 2)

         End If
      
         If Len(Trim(g_rst_Princi!EVACRE_TIT_CODEN6 & "")) > 0 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (6)"
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_NomEntFin(Trim(g_rst_Princi!EVACRE_TIT_CODEN6) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_TIT_CLAEN6)) & ")"
            
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_DEUEN6, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!EVACRE_TIT_LIMDE6), 0, g_rst_Princi!EVACRE_TIT_LIMDE6), 12, 2)

         End If
      End If
      
      'Central de Riesgo (C�nyuge)
      If g_rst_Princi!EVACRE_CYG_CRIFLG > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 2:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Central de Riesgo - C�nyuge Reportado"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!EVACRE_CYG_CRIFLG))
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Fecha Reporte"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACRE_CYG_CRIFEC))
         
         If g_rst_Princi!EVACRE_CYG_CRIFLG = 1 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Nro. Entidades:"
            grd_LisEva.Col = 1:                          grd_LisEva.Text = CStr(g_rst_Princi!EVACRE_CYG_CRIENT)
            
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Clasificaci�n"
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
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_NomEntFin(Trim(g_rst_Princi!EVACRE_CYG_CODEN1) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_CYG_CLAEN1)) & ")"
            
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_DEUEN1, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!EVACRE_CYG_LIMDE1), 0, g_rst_Princi!EVACRE_CYG_LIMDE1), 12, 2)

            
            If Len(Trim(g_rst_Princi!EVACRE_CYG_CODEN2 & "")) > 0 Then
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (2)"
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_NomEntFin(Trim(g_rst_Princi!EVACRE_CYG_CODEN2) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_CYG_CLAEN2)) & ")"
               
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_DEUEN2, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!EVACRE_CYG_LIMDE2), 0, g_rst_Princi!EVACRE_CYG_LIMDE2), 12, 2)

            End If
         
            If Len(Trim(g_rst_Princi!EVACRE_CYG_CODEN3 & "")) > 0 Then
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (3)"
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_NomEntFin(Trim(g_rst_Princi!EVACRE_CYG_CODEN3) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_CYG_CLAEN3)) & ")"
               
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_DEUEN3, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!EVACRE_CYG_LIMDE3), 0, g_rst_Princi!EVACRE_CYG_LIMDE3), 12, 2)

            End If
         
            If Len(Trim(g_rst_Princi!EVACRE_CYG_CODEN4 & "")) > 0 Then
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (4)"
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_NomEntFin(Trim(g_rst_Princi!EVACRE_CYG_CODEN4) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_CYG_CLAEN4)) & ")"
            
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_DEUEN4, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!EVACRE_CYG_LIMDE4), 0, g_rst_Princi!EVACRE_CYG_LIMDE4), 12, 2)

            End If
         
            If Len(Trim(g_rst_Princi!EVACRE_CYG_CODEN5 & "")) > 0 Then
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (5)"
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_NomEntFin(Trim(g_rst_Princi!EVACRE_CYG_CODEN5) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_CYG_CLAEN5)) & ")"
            
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_DEUEN5, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!EVACRE_CYG_LIMDE5), 0, g_rst_Princi!EVACRE_CYG_LIMDE5), 12, 2)

            End If
         
            If Len(Trim(g_rst_Princi!EVACRE_CYG_CODEN6 & "")) > 0 Then
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (6)"
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_NomEntFin(Trim(g_rst_Princi!EVACRE_CYG_CODEN6) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_CYG_CLAEN6)) & ")"
            
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_DEUEN6, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!EVACRE_CYG_LIMDE6), 0, g_rst_Princi!EVACRE_CYG_LIMDE6), 12, 2)

            End If
         End If
      End If
   End If
   
   'Referencias Personales
   If Len(Trim(g_rst_Princi!EVACRE_REFPER & "")) > 0 Then
      grd_LisEva.Rows = grd_LisEva.Rows + 2:       grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                          grd_LisEva.Text = "Verificaci�n Referencias"
      grd_LisEva.Col = 1:                          grd_LisEva.Text = Trim(g_rst_Princi!EVACRE_REFPER & "")
   End If
   
   'Actividades Econ�micas
   If Not IsNull(g_rst_Princi!EVACRE_TIT_TIPVE1) Then
      If g_rst_Princi!EVACRE_TIT_TIPVE1 > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 2:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Titular - Verif. Lab. (Act. Princ.)"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = moddat_gf_Consulta_ParDes("068", g_rst_Princi!EVACRE_TIT_TIPVE1)
      
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 1:                          grd_LisEva.Text = Trim(g_rst_Princi!EVACRE_TIT_LABVE1 & "")
      End If
      
      If g_rst_Princi!EVACRE_TIT_TIPVE2 > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Titular - Verif. Lab. (Act. Secund.)"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = moddat_gf_Consulta_ParDes("068", g_rst_Princi!EVACRE_TIT_TIPVE2)
      
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 1:                          grd_LisEva.Text = Trim(g_rst_Princi!EVACRE_TIT_LABVE2 & "")
      End If
   
      If g_rst_Princi!EVACRE_CYG_TIPVE1 > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 2:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "C�nyuge - Verif. Lab. (Act. Princ.)"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = moddat_gf_Consulta_ParDes("068", g_rst_Princi!EVACRE_CYG_TIPVE1)
      
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 1:                          grd_LisEva.Text = Trim(g_rst_Princi!EVACRE_CYG_LABVE1 & "")
      End If
   
      If g_rst_Princi!EVACRE_CYG_TIPVE2 > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "C�nyuge - Verif. Lab. (Act. Secund.)"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = moddat_gf_Consulta_ParDes("068", g_rst_Princi!EVACRE_CYG_TIPVE2)
      
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 1:                          grd_LisEva.Text = Trim(g_rst_Princi!EVACRE_CYG_LABVE2 & "")
      End If
   End If
   
   'Ingresos
   If g_rst_Princi!EVACRE_INGDP1 > 0 Then
      grd_LisEva.Rows = grd_LisEva.Rows + 2:       grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                          grd_LisEva.Text = "Ingreso L�quido Titular"
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
         grd_LisEva.Col = 0:                       grd_LisEva.Text = "Ingreso L�quido C�nyuge"
         grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGDP3, 12, 2) & " + " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGDP4, 12, 2) & " + " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGAD2, 12, 2) & " = " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGDP3 + g_rst_Princi!EVACRE_INGDP4 + g_rst_Princi!EVACRE_INGAD2, 12, 2)
      End If
   
      If g_rst_Princi!EVACRE_OMECYG > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                       grd_LisEva.Text = "Obligaciones Mensuales C�nyuge"
         grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_OMECYG, 12, 2)
      End If
      
      If g_rst_Princi!EVACRE_INTCYG > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                       grd_LisEva.Text = "Ingreso Neto C�nyuge"
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
