VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Con_PrePgo_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   10335
   ClientLeft      =   5355
   ClientTop       =   1485
   ClientWidth     =   11685
   Icon            =   "OpeTra_frm_324.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10335
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel111 
      Height          =   10335
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   11685
      _Version        =   65536
      _ExtentX        =   20611
      _ExtentY        =   18230
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   2955
         Left            =   30
         TabIndex        =   33
         Top             =   7320
         Width           =   11625
         _Version        =   65536
         _ExtentX        =   20505
         _ExtentY        =   5212
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
            Height          =   2565
            Left            =   90
            TabIndex        =   16
            Top             =   330
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   4524
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Cronograma - Cliente TNC"
            TabPicture(0)   =   "OpeTra_frm_324.frx":000C
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
            Tab(0).Control(9)=   "SSPanel2"
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
            TabPicture(1)   =   "OpeTra_frm_324.frx":0028
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "pnl_CliCon_TotCuo"
            Tab(1).Control(1)=   "SSPanel9"
            Tab(1).Control(2)=   "grd_CliCon_Listad"
            Tab(1).Control(3)=   "SSPanel10"
            Tab(1).Control(4)=   "SSPanel11"
            Tab(1).Control(5)=   "SSPanel12"
            Tab(1).Control(6)=   "SSPanel13"
            Tab(1).Control(7)=   "SSPanel21"
            Tab(1).Control(8)=   "pnl_CliCon_Intere"
            Tab(1).Control(9)=   "pnl_CliCon_Capita"
            Tab(1).ControlCount=   10
            Begin Threed.SSPanel SSPanel30 
               Height          =   285
               Left            =   3450
               TabIndex        =   35
               Top             =   360
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
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
            Begin Threed.SSPanel pnl_CliNCo_Capita 
               Height          =   285
               Left            =   2280
               TabIndex        =   36
               Top             =   2190
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
            Begin Threed.SSPanel pnl_CliNCo_SegViv 
               Height          =   285
               Left            =   5790
               TabIndex        =   37
               Top             =   2190
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
            Begin Threed.SSPanel pnl_CliNCo_SegPre 
               Height          =   285
               Left            =   4620
               TabIndex        =   38
               Top             =   2190
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
            Begin Threed.SSPanel pnl_CliNCo_Intere 
               Height          =   285
               Left            =   3450
               TabIndex        =   39
               Top             =   2190
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
            Begin MSFlexGridLib.MSFlexGrid grd_CliNCo_Listad 
               Height          =   1515
               Left            =   30
               TabIndex        =   17
               Top             =   660
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   2672
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
               TabIndex        =   40
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
               TabIndex        =   41
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
            Begin Threed.SSPanel SSPanel26 
               Height          =   285
               Left            =   -74940
               TabIndex        =   42
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
               TabIndex        =   43
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
               TabIndex        =   44
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
               TabIndex        =   45
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
               TabIndex        =   46
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
               TabIndex        =   47
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
            Begin Threed.SSPanel SSPanel37 
               Height          =   285
               Left            =   -68310
               TabIndex        =   48
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
            Begin Threed.SSPanel SSPanel40 
               Height          =   285
               Left            =   -74940
               TabIndex        =   49
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
               TabIndex        =   50
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
               TabIndex        =   51
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
               TabIndex        =   52
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
               TabIndex        =   53
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
               TabIndex        =   54
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
            Begin Threed.SSPanel SSPanel57 
               Height          =   285
               Left            =   -68310
               TabIndex        =   55
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
            Begin Threed.SSPanel SSPanel2 
               Height          =   285
               Left            =   60
               TabIndex        =   56
               Top             =   360
               Width           =   795
               _Version        =   65536
               _ExtentX        =   1402
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
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
            Begin Threed.SSPanel SSPanel33 
               Height          =   285
               Left            =   840
               TabIndex        =   57
               Top             =   360
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
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
            Begin Threed.SSPanel SSPanel34 
               Height          =   285
               Left            =   2280
               TabIndex        =   58
               Top             =   360
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
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
            Begin Threed.SSPanel SSPanel35 
               Height          =   285
               Left            =   8130
               TabIndex        =   59
               Top             =   360
               Width           =   1320
               _Version        =   65536
               _ExtentX        =   2328
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
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
            Begin Threed.SSPanel SSPanel36 
               Height          =   285
               Left            =   9420
               TabIndex        =   60
               Top             =   360
               Width           =   1560
               _Version        =   65536
               _ExtentX        =   2752
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
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
            Begin Threed.SSPanel SSPanel59 
               Height          =   285
               Left            =   4620
               TabIndex        =   61
               Top             =   360
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg. Prest."
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
            Begin Threed.SSPanel SSPanel61 
               Height          =   285
               Left            =   5790
               TabIndex        =   62
               Top             =   360
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg. Vivienda"
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
            Begin Threed.SSPanel SSPanel62 
               Height          =   285
               Left            =   6960
               TabIndex        =   63
               Top             =   360
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Portes"
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
            Begin Threed.SSPanel pnl_CliNCo_OtrCar 
               Height          =   285
               Left            =   6960
               TabIndex        =   64
               Top             =   2190
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
            Begin Threed.SSPanel pnl_CliNCo_TotCuo 
               Height          =   285
               Left            =   8130
               TabIndex        =   65
               Top             =   2190
               Width           =   1320
               _Version        =   65536
               _ExtentX        =   2328
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
            Begin Threed.SSPanel pnl_CliCon_TotCuo 
               Height          =   285
               Left            =   -68370
               TabIndex        =   66
               Top             =   2190
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3828
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
            Begin Threed.SSPanel SSPanel9 
               Height          =   285
               Left            =   -70530
               TabIndex        =   67
               Top             =   360
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3828
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interes"
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
            Begin MSFlexGridLib.MSFlexGrid grd_CliCon_Listad 
               Height          =   1515
               Left            =   -74970
               TabIndex        =   18
               Top             =   660
               Width           =   11355
               _ExtentX        =   20029
               _ExtentY        =   2672
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
            Begin Threed.SSPanel SSPanel10 
               Height          =   285
               Left            =   -74940
               TabIndex        =   68
               Top             =   360
               Width           =   765
               _Version        =   65536
               _ExtentX        =   1349
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
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
               Left            =   -74190
               TabIndex        =   69
               Top             =   360
               Width           =   1515
               _Version        =   65536
               _ExtentX        =   2672
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
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
               Height          =   285
               Left            =   -72690
               TabIndex        =   70
               Top             =   360
               Width           =   2170
               _Version        =   65536
               _ExtentX        =   3828
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
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
            Begin Threed.SSPanel SSPanel13 
               Height          =   285
               Left            =   -68370
               TabIndex        =   71
               Top             =   360
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3828
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
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
            Begin Threed.SSPanel SSPanel21 
               Height          =   285
               Left            =   -66210
               TabIndex        =   72
               Top             =   360
               Width           =   2235
               _Version        =   65536
               _ExtentX        =   3942
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
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
            Begin Threed.SSPanel pnl_CliCon_Intere 
               Height          =   285
               Left            =   -70530
               TabIndex        =   73
               Top             =   2190
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3828
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
            Begin Threed.SSPanel pnl_CliCon_Capita 
               Height          =   285
               Left            =   -72690
               TabIndex        =   74
               Top             =   2190
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3828
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
            Begin VB.Label Label4 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   -72930
               TabIndex        =   77
               Top             =   1470
               Width           =   945
            End
            Begin VB.Label Label14 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   -72930
               TabIndex        =   76
               Top             =   1470
               Width           =   945
            End
            Begin VB.Label Label15 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   -73230
               TabIndex        =   75
               Top             =   1470
               Width           =   945
            End
         End
         Begin VB.Label Label12 
            Caption         =   "Cronograma"
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
            Left            =   120
            TabIndex        =   79
            Top             =   60
            Width           =   1875
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   4035
         Left            =   30
         TabIndex        =   32
         Top             =   3270
         Width           =   11625
         _Version        =   65536
         _ExtentX        =   20505
         _ExtentY        =   7117
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
         Begin TabDlg.SSTab Tab_Deuda 
            Height          =   1995
            Index           =   1
            Left            =   90
            TabIndex        =   4
            Top             =   1140
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   3519
            _Version        =   393216
            TabHeight       =   520
            TabCaption(0)   =   "Deuda Vigente"
            TabPicture(0)   =   "OpeTra_frm_324.frx":0044
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label9"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label24"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label25"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Label32"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "Label33"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "Label26"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "Label19"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "Label7"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "Label6"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "Label17"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "Label23"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "Label22"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "Label3"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "Label34"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "pnl_MtoApl_Fin"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "pnl_MtoApl"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "pnl_DeuPen"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).Control(17)=   "pnl_DiasTC"
            Tab(0).Control(17).Enabled=   0   'False
            Tab(0).Control(18)=   "pnl_DiasTNC"
            Tab(0).Control(18).Enabled=   0   'False
            Tab(0).Control(19)=   "pnl_UltPagTC"
            Tab(0).Control(19).Enabled=   0   'False
            Tab(0).Control(20)=   "pnl_UltPagTNC"
            Tab(0).Control(20).Enabled=   0   'False
            Tab(0).Control(21)=   "pnl_IntPbp"
            Tab(0).Control(21).Enabled=   0   'False
            Tab(0).Control(22)=   "pnl_CapPbp"
            Tab(0).Control(22).Enabled=   0   'False
            Tab(0).Control(23)=   "txt_MontoITF"
            Tab(0).Control(23).Enabled=   0   'False
            Tab(0).Control(24)=   "txt_SegInm"
            Tab(0).Control(24).Enabled=   0   'False
            Tab(0).Control(25)=   "txt_SegDes"
            Tab(0).Control(25).Enabled=   0   'False
            Tab(0).Control(26)=   "txt_InteresTNC"
            Tab(0).Control(26).Enabled=   0   'False
            Tab(0).Control(27)=   "txt_InteresTC"
            Tab(0).Control(27).Enabled=   0   'False
            Tab(0).ControlCount=   28
            TabCaption(1)   =   "Deuda Pendiente (*)"
            TabPicture(1)   =   "OpeTra_frm_324.frx":0060
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_DeuPen"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Origen de Fondos"
            TabPicture(2)   =   "OpeTra_frm_324.frx":007C
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "txt_ObsPpg"
            Tab(2).Control(1)=   "cmb_MotPpg"
            Tab(2).Control(2)=   "Label37"
            Tab(2).Control(3)=   "Label36"
            Tab(2).ControlCount=   4
            Begin VB.TextBox txt_ObsPpg 
               Height          =   1065
               Left            =   -73410
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   12
               Top             =   810
               Width           =   9585
            End
            Begin VB.ComboBox cmb_MotPpg 
               Height          =   315
               Left            =   -73410
               TabIndex        =   11
               Text            =   "MOTIVO DEL PREPAGO"
               Top             =   480
               Width           =   9600
            End
            Begin EditLib.fpDoubleSingle txt_InteresTC 
               Height          =   315
               Left            =   7500
               TabIndex        =   6
               Top             =   840
               Width           =   1170
               _Version        =   196608
               _ExtentX        =   2064
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
            Begin EditLib.fpDoubleSingle txt_InteresTNC 
               Height          =   315
               Left            =   7500
               TabIndex        =   5
               Top             =   510
               Width           =   1170
               _Version        =   196608
               _ExtentX        =   2064
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
            Begin EditLib.fpDoubleSingle txt_SegDes 
               Height          =   315
               Left            =   1470
               TabIndex        =   7
               Top             =   1170
               Width           =   1290
               _Version        =   196608
               _ExtentX        =   2275
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
            Begin EditLib.fpDoubleSingle txt_SegInm 
               Height          =   315
               Left            =   4260
               TabIndex        =   8
               Top             =   1170
               Width           =   1140
               _Version        =   196608
               _ExtentX        =   2011
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
            Begin EditLib.fpDoubleSingle txt_MontoITF 
               Height          =   315
               Left            =   7500
               TabIndex        =   9
               Top             =   1170
               Width           =   1170
               _Version        =   196608
               _ExtentX        =   2064
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
            Begin Threed.SSPanel pnl_CapPbp 
               Height          =   315
               Left            =   4260
               TabIndex        =   102
               Top             =   1500
               Width           =   1140
               _Version        =   65536
               _ExtentX        =   2011
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "0.00 "
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
               RoundedCorners  =   0   'False
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_IntPbp 
               Height          =   345
               Left            =   10110
               TabIndex        =   103
               Top             =   810
               Width           =   1110
               _Version        =   65536
               _ExtentX        =   1958
               _ExtentY        =   609
               _StockProps     =   15
               Caption         =   "0.00 "
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
               RoundedCorners  =   0   'False
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_UltPagTNC 
               Height          =   315
               Left            =   1470
               TabIndex        =   104
               Top             =   510
               Width           =   1290
               _Version        =   65536
               _ExtentX        =   2275
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "31/12/2011"
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
               RoundedCorners  =   0   'False
               Font3D          =   2
            End
            Begin Threed.SSPanel pnl_UltPagTC 
               Height          =   315
               Left            =   1470
               TabIndex        =   105
               Top             =   840
               Width           =   1290
               _Version        =   65536
               _ExtentX        =   2275
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "31/12/2011"
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
               RoundedCorners  =   0   'False
               Font3D          =   2
            End
            Begin Threed.SSPanel pnl_DiasTNC 
               Height          =   315
               Left            =   4260
               TabIndex        =   106
               Top             =   510
               Width           =   1140
               _Version        =   65536
               _ExtentX        =   2011
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "30 "
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
               RoundedCorners  =   0   'False
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_DiasTC 
               Height          =   315
               Left            =   4260
               TabIndex        =   107
               Top             =   840
               Width           =   1140
               _Version        =   65536
               _ExtentX        =   2011
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "180 "
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
               RoundedCorners  =   0   'False
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_DeuPen 
               Height          =   315
               Left            =   10110
               TabIndex        =   108
               Top             =   1170
               Width           =   1110
               _Version        =   65536
               _ExtentX        =   1958
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "0.00 "
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
               RoundedCorners  =   0   'False
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_MtoApl 
               Height          =   315
               Left            =   1470
               TabIndex        =   109
               Top             =   1500
               Width           =   1290
               _Version        =   65536
               _ExtentX        =   2275
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "0.00 "
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
               RoundedCorners  =   0   'False
               Font3D          =   2
               Alignment       =   4
            End
            Begin MSFlexGridLib.MSFlexGrid grd_DeuPen 
               Height          =   1365
               Left            =   -74880
               TabIndex        =   10
               Top             =   480
               Width           =   11145
               _ExtentX        =   19659
               _ExtentY        =   2408
               _Version        =   393216
               Rows            =   5
               Cols            =   16
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               SelectionMode   =   1
            End
            Begin Threed.SSPanel pnl_MtoApl_Fin 
               Height          =   315
               Left            =   7500
               TabIndex        =   123
               Top             =   1500
               Width           =   1170
               _Version        =   65536
               _ExtentX        =   2064
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "0.00 "
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
               RoundedCorners  =   0   'False
               Font3D          =   2
               Alignment       =   4
            End
            Begin VB.Label Label37 
               Caption         =   "Origen Fondos"
               Height          =   315
               Left            =   -74760
               TabIndex        =   129
               Top             =   510
               Width           =   1290
            End
            Begin VB.Label Label36 
               Caption         =   "Comentarios"
               Height          =   315
               Left            =   -74760
               TabIndex        =   128
               Top             =   840
               Width           =   1290
            End
            Begin VB.Label Label34 
               Caption         =   "Monto a Aplicar Final"
               Height          =   315
               Left            =   5730
               TabIndex        =   124
               Top             =   1560
               Width           =   1755
            End
            Begin VB.Label Label3 
               Caption         =   "Monto a Aplicar"
               Height          =   315
               Left            =   120
               TabIndex        =   122
               Top             =   1560
               Width           =   1350
            End
            Begin VB.Label Label22 
               Caption         =   "Último Vcto. TC"
               Height          =   315
               Left            =   120
               TabIndex        =   121
               Top             =   900
               Width           =   1350
            End
            Begin VB.Label Label23 
               Caption         =   "Último Vcto. TNC"
               Height          =   315
               Left            =   120
               TabIndex        =   120
               Top             =   570
               Width           =   1350
            End
            Begin VB.Label Label17 
               Caption         =   "Seguro Desgrav."
               Height          =   315
               Left            =   120
               TabIndex        =   119
               Top             =   1230
               Width           =   1350
            End
            Begin VB.Label Label6 
               Caption         =   "Interés TNC a la fecha"
               Height          =   315
               Left            =   5730
               TabIndex        =   118
               Top             =   570
               Width           =   1755
            End
            Begin VB.Label Label7 
               Caption         =   "Interés TC a la fecha"
               Height          =   315
               Left            =   5730
               TabIndex        =   117
               Top             =   900
               Width           =   1755
            End
            Begin VB.Label Label19 
               Caption         =   "Seg. Inmueble"
               Height          =   315
               Left            =   3090
               TabIndex        =   116
               Top             =   1230
               Width           =   1155
            End
            Begin VB.Label Label26 
               Caption         =   "Monto del ITF"
               Height          =   315
               Left            =   5730
               TabIndex        =   115
               Top             =   1230
               Width           =   1755
            End
            Begin VB.Label Label33 
               Caption         =   "Interés PBP"
               Height          =   315
               Left            =   8910
               TabIndex        =   114
               Top             =   900
               Width           =   1185
            End
            Begin VB.Label Label32 
               Caption         =   "Capital PBP"
               Height          =   315
               Left            =   3090
               TabIndex        =   113
               Top             =   1560
               Width           =   1155
            End
            Begin VB.Label Label25 
               Caption         =   "Dias TNC"
               Height          =   315
               Left            =   3090
               TabIndex        =   112
               Top             =   570
               Width           =   1155
            End
            Begin VB.Label Label24 
               Caption         =   "Dias TC"
               Height          =   315
               Left            =   3090
               TabIndex        =   111
               Top             =   900
               Width           =   1155
            End
            Begin VB.Label Label9 
               Caption         =   "Deuda Pend.(*)"
               Height          =   285
               Left            =   8910
               TabIndex        =   110
               Top             =   1230
               Width           =   1185
            End
         End
         Begin VB.ComboBox cmb_RedPlz 
            Height          =   315
            Left            =   7260
            TabIndex        =   2
            Text            =   "1 AÑO"
            Top             =   390
            Width           =   1230
         End
         Begin VB.ComboBox cmb_TipPre 
            Height          =   315
            Left            =   4350
            TabIndex        =   1
            Text            =   "RED. MONTO"
            Top             =   390
            Width           =   1440
         End
         Begin VB.CommandButton cmd_Recalc 
            Caption         =   "Regenerar Cronograma"
            Height          =   345
            Left            =   8670
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Procesar"
            Top             =   3270
            Width           =   2835
         End
         Begin EditLib.fpDateTime ipp_FecPre 
            Height          =   315
            Left            =   1500
            TabIndex        =   0
            Top             =   390
            Width           =   1320
            _Version        =   196608
            _ExtentX        =   2328
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
            Text            =   "28/09/2004"
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
         Begin EditLib.fpDoubleSingle txt_Mto_Deposito 
            Height          =   315
            Left            =   10320
            TabIndex        =   3
            Top             =   390
            Width           =   1170
            _Version        =   196608
            _ExtentX        =   2064
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
            Text            =   "175,000.00"
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
         Begin EditLib.fpDoubleSingle txt_ApliTC 
            Height          =   315
            Left            =   4350
            TabIndex        =   14
            Top             =   3630
            Width           =   1440
            _Version        =   196608
            _ExtentX        =   2540
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
         Begin EditLib.fpDoubleSingle txt_AplTNC 
            Height          =   315
            Left            =   4350
            TabIndex        =   13
            Top             =   3300
            Width           =   1440
            _Version        =   196608
            _ExtentX        =   2540
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
         Begin Threed.SSPanel pnl_NuevoSaldoTNC 
            Height          =   315
            Left            =   7260
            TabIndex        =   82
            Top             =   3300
            Width           =   1230
            _Version        =   65536
            _ExtentX        =   2170
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
            RoundedCorners  =   0   'False
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_NuevoSaldoTC 
            Height          =   315
            Left            =   7260
            TabIndex        =   83
            Top             =   3630
            Width           =   1230
            _Version        =   65536
            _ExtentX        =   2170
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
            RoundedCorners  =   0   'False
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_NuevaCuota 
            Height          =   315
            Left            =   10350
            TabIndex        =   84
            Top             =   3630
            Width           =   1140
            _Version        =   65536
            _ExtentX        =   2011
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
            RoundedCorners  =   0   'False
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_SaldoTNC1 
            Height          =   315
            Left            =   1500
            TabIndex        =   89
            Top             =   720
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "135,000.00 "
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
            RoundedCorners  =   0   'False
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_SaldoTC1 
            Height          =   315
            Left            =   4350
            TabIndex        =   90
            Top             =   720
            Width           =   1440
            _Version        =   65536
            _ExtentX        =   2540
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "15,000.00 "
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
            RoundedCorners  =   0   'False
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_SaldoTNC2 
            Height          =   315
            Left            =   1500
            TabIndex        =   93
            Top             =   3300
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "135,000.00 "
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
            RoundedCorners  =   0   'False
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_SaldoTC2 
            Height          =   315
            Left            =   1500
            TabIndex        =   94
            Top             =   3630
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "15,000.00 "
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
            RoundedCorners  =   0   'False
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Val_AsgInm 
            Height          =   315
            Left            =   7260
            TabIndex        =   100
            Top             =   750
            Width           =   1230
            _Version        =   65536
            _ExtentX        =   2170
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1,250,000.00 "
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
            RoundedCorners  =   0   'False
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_CuoPen 
            Height          =   315
            Left            =   10320
            TabIndex        =   125
            Top             =   720
            Width           =   1170
            _Version        =   65536
            _ExtentX        =   2064
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0 "
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
            RoundedCorners  =   0   'False
            Font3D          =   2
            Alignment       =   4
         End
         Begin VB.Label Label35 
            Caption         =   "Cuota Pendientes"
            Height          =   315
            Left            =   8760
            TabIndex        =   127
            Top             =   750
            Width           =   1305
         End
         Begin VB.Label Label1 
            Caption         =   "Monto Depositado"
            Height          =   315
            Left            =   8760
            TabIndex        =   126
            Top             =   450
            Width           =   1305
         End
         Begin VB.Label Label5 
            Caption         =   "V. Asegur. Inm."
            Height          =   315
            Left            =   6060
            TabIndex        =   101
            Top             =   780
            Width           =   1215
         End
         Begin VB.Label Label30 
            Caption         =   "Reducc. Años"
            Height          =   315
            Left            =   6060
            TabIndex        =   99
            Top             =   450
            Width           =   1215
         End
         Begin VB.Label Label31 
            Caption         =   "Tipo de Prepago"
            Height          =   315
            Left            =   3060
            TabIndex        =   98
            Top             =   450
            Width           =   1320
         End
         Begin VB.Label Label29 
            Caption         =   "Datos del Prepago"
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
            Left            =   120
            TabIndex        =   97
            Top             =   60
            UseMnemonic     =   0   'False
            Width           =   1875
         End
         Begin VB.Label Label28 
            Caption         =   "Saldo TNC"
            Height          =   315
            Left            =   120
            TabIndex        =   96
            Top             =   3360
            Width           =   1500
         End
         Begin VB.Label Label27 
            Caption         =   "Saldo TC"
            Height          =   315
            Left            =   120
            TabIndex        =   95
            Top             =   3690
            Width           =   1500
         End
         Begin VB.Label Label21 
            Caption         =   "Saldo Actual TNC"
            Height          =   315
            Left            =   120
            TabIndex        =   92
            Top             =   780
            Width           =   1395
         End
         Begin VB.Label Label20 
            Caption         =   "Saldo Actual TC"
            Height          =   315
            Left            =   3060
            TabIndex        =   91
            Top             =   780
            Width           =   1320
         End
         Begin VB.Label Label8 
            Caption         =   "Aplica. PP TNC"
            Height          =   315
            Left            =   3060
            TabIndex        =   88
            Top             =   3360
            Width           =   1500
         End
         Begin VB.Label Label18 
            Caption         =   "Nueva Monto Cuota"
            Height          =   315
            Left            =   8670
            TabIndex        =   87
            Top             =   3690
            Width           =   1665
         End
         Begin VB.Label Label16 
            Caption         =   "Nuevo TC"
            Height          =   315
            Left            =   6090
            TabIndex        =   86
            Top             =   3690
            Width           =   1500
         End
         Begin VB.Label Label10 
            Caption         =   "Nuevo TNC"
            Height          =   315
            Left            =   6090
            TabIndex        =   85
            Top             =   3360
            Width           =   1500
         End
         Begin VB.Label Label13 
            Caption         =   "Aplica. PP TC"
            Height          =   315
            Left            =   3060
            TabIndex        =   81
            Top             =   3690
            Width           =   1500
         End
         Begin VB.Label Label11 
            Caption         =   "Fecha Prepago"
            Height          =   315
            Left            =   120
            TabIndex        =   80
            Top             =   450
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   1935
         Left            =   30
         TabIndex        =   31
         Top             =   1320
         Width           =   11625
         _Version        =   65536
         _ExtentX        =   20505
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   1545
            Left            =   60
            TabIndex        =   19
            Top             =   330
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   2725
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.Label Label2 
            Caption         =   "Datos del Crédito"
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
            Left            =   120
            TabIndex        =   78
            Top             =   60
            Width           =   1875
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   645
         Left            =   30
         TabIndex        =   30
         Top             =   660
         Width           =   11625
         _Version        =   65536
         _ExtentX        =   20505
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
         Begin VB.CommandButton cmd_ForCon 
            Height          =   585
            Left            =   2430
            Picture         =   "OpeTra_frm_324.frx":0098
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Imprime Formato AFP"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   1830
            Picture         =   "OpeTra_frm_324.frx":04DA
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Exportar simulación de liquidación"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   3600
            Picture         =   "OpeTra_frm_324.frx":07E4
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Limpiar datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   3000
            Picture         =   "OpeTra_frm_324.frx":0AEE
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_PolSeg 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_324.frx":0F30
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Consulta Pólizas de Seguros"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   11010
            Picture         =   "OpeTra_frm_324.frx":123A
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_VerPag 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_324.frx":167C
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Consulta Pagos del Cliente"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ImpCro 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_324.frx":1986
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Consulta Cronograma de Pagos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   615
         Left            =   30
         TabIndex        =   29
         Top             =   30
         Width           =   11625
         _Version        =   65536
         _ExtentX        =   20505
         _ExtentY        =   1085
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
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   11130
            Top             =   90
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   555
            Left            =   690
            TabIndex        =   34
            Top             =   30
            Width           =   4755
            _Version        =   65536
            _ExtentX        =   8387
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "Prepago Parcial de Crédito Hipotecario - Registro"
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
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   10410
            Top             =   30
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            AddressEditFieldCount=   1
            AddressModifiable=   0   'False
            AddressResolveUI=   0   'False
            FetchSorted     =   0   'False
            FetchUnreadOnly =   0   'False
         End
         Begin MSMAPI.MAPISession mps_Sesion 
            Left            =   9840
            Top             =   30
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   90
            Picture         =   "OpeTra_frm_324.frx":1C90
            Top             =   60
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_Con_PrePgo_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Dim l_Arr_TNC_Cli()        As String
Dim l_Arr_TC_Cli()         As String
Dim l_Arr_TNC_Cof()        As String
Dim l_Arr_TC_Cof()         As String

Dim l_dbl_PorNco           As Double
Dim l_dbl_PorCon           As Double
Dim l_int_PlaAno           As Integer
Dim l_dbl_ComCof           As Double
Dim l_dbl_TasCof           As Double
Dim l_str_DesCof           As String
Dim l_str_PrxVct           As String
Dim l_int_NumCuo           As Integer
Dim l_int_PagCuo           As Integer
Dim l_int_PerGra           As Integer
Dim l_dbl_SalNco           As Double
Dim l_dbl_SalCon           As Double
Dim l_int_CodPrd           As Integer
Dim l_dbl_TasInt           As Double
Dim l_dbl_SegDes           As Double
Dim l_dbl_SegInm           As Double
Dim l_dbl_PorITF           As Double
Dim l_str_UltVct           As String
Dim l_dbl_CuoFij           As Double
Dim l_dbl_PagCap           As Double
Dim l_dbl_PagInt           As Double
Dim l_int_NCuota           As Integer
Dim l_int_TipSeg           As Double 'seguro desgravamen
Dim l_str_NomAch           As String

Private Type r_Arr_CroTNC
   str_NumOpe         As String
   int_NumCuo         As Integer
   dbl_SegPre         As Double
   dbl_SegViv         As Double
End Type
Dim arr_CroTNC()       As r_Arr_CroTNC

Private Sub cmd_VerPag_Click()
   frm_Ges_CreHip_05.Show 1
End Sub

Private Sub cmd_ImpCro_Click()
   modmip_g_int_OrdAct = 1
   frm_Ges_CreHip_07.Show 1
End Sub

Private Sub cmd_PolSeg_Click()
   frm_Con_PolSeg_01.Show 1
End Sub

Private Sub cmd_ExpExc_Click()
   If grd_CliNCo_Listad.Rows = 0 Then
      MsgBox "Debe generar el cronograma de tramo no concesional.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de generar la simulación de la liquidación y el cronograma?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'exporta liquidacion
   Screen.MousePointer = 11
   If InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then
      Call fs_PpgPar_Micasita(False, l_str_NomAch)
   Else
      Call fs_PpgPar_Mivivienda(False, l_str_NomAch)
   End If
   
   'exporta cronograma
   Call fs_Exportar
    
   'Graba datos en Solicitud de tabla prepagos
   If fs_usp_cre_ppgsol = 1 Then
      MsgBox "No se pudo completar el procedimiento 'usp_cre_ppgsol'.", vbCritical, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_ExpExc)
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   'Envio de correo preventivo
   Call fs_Envia_Correo_Preventivo
   Screen.MousePointer = 0
End Sub

Private Sub cmd_ForCon_Click()
   If (Len(Trim(cmb_TipPre.Text)) <= 0) Then
      MsgBox "Seleccione el tipo de prepago.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipPre)
      Exit Sub
   End If
   If (CDbl(txt_Mto_Deposito.Text) <= 0) Then
      MsgBox "El Monto del depósito debe ser mayor a 0.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Mto_Deposito)
      Exit Sub
   End If
   If CDbl(pnl_NuevaCuota.Caption) = 0 Then
      MsgBox "No se tiene el valor de la nueva cuota.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Mto_Deposito)
      Exit Sub
   End If
   
   moddat_g_str_CodIte = cmb_TipPre.ItemData(cmb_TipPre.ListIndex)
   moddat_g_str_TipPar = cmb_TipPre.Text
   frm_Con_PrePgo_06.Show 1
End Sub

Private Sub cmd_Grabar_Click()
   'Validaciones
   If cmb_TipPre.ListIndex = -1 Then
      MsgBox "Seleccione el tipo de prepago.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipPre)
      Exit Sub
   End If
   If CDbl(txt_Mto_Deposito.Text) = 0 Then
      MsgBox "El Monto del depósito debe ser mayor a 0.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Mto_Deposito)
      Exit Sub
   End If
   If InStr(moddat_g_str_AgrTMIC, moddat_g_str_CodPrd) > 0 Then
      If CDbl(txt_MontoITF.Text) = 0 Then
         MsgBox "El Monto de ITF debe ser mayor a 0.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_MontoITF)
         Exit Sub
      End If
   End If
   If CDbl(pnl_DeuPen.Caption) > 0 Then
      MsgBox "El cliente tiene Deuda Pendiente, favor de regularizarla antes de realizar el prepago.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Mto_Deposito)
      Exit Sub
   End If
   If cmb_MotPpg.ListIndex = -1 Then
      MsgBox "Debe seleccionar el origen de los fondos del prepago.", vbExclamation, modgen_g_str_NomPlt
      Tab_Deuda(1).Tab = 2
      Call gs_SetFocus(cmb_MotPpg)
      Exit Sub
   End If
   
   'Validaciones reduccion de plazo
   If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) = 2 Then
      If cmb_RedPlz.ListIndex = -1 Then
         MsgBox "Seleccione el numero de años a reducir.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_RedPlz)
         Exit Sub
      End If
   End If
   If CDbl(pnl_NuevaCuota.Caption) > l_dbl_CuoFij Then
      If CDbl(pnl_NuevaCuota.Caption) - l_dbl_CuoFij > 5 Then
         MsgBox "La nueva cuota generada es mayor que la cuota del cronograma actual.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Mto_Deposito)
         Exit Sub
      Else
         If MsgBox("La nueva cuota generada es mayor que la cuota del cronograma actual." & vbCrLf & "¿Está seguro de continuar?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Call gs_SetFocus(txt_Mto_Deposito)
            Exit Sub
         End If
      End If
   End If
   
    'Valida que no se tenga un prepago sin regularizar
'   If fs_validar_ppgpnd = 1 Then
'      MsgBox "Existe un prepago pendiente de regularización de COFIDE.", vbExclamation, modgen_g_str_NomPlt
'         Call gs_SetFocus(cmd_Grabar)
'         Exit Sub
'   End If

   'Validar Datos del Prepago
   If fs_validar_ppg = 0 Then
  
      If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
      
      Screen.MousePointer = 11
      
      'Valida que la operacion no exista
      If fs_Validar_NumOpe = 1 Then
         MsgBox "El número de operación y la fecha del prepago ya existen, vuelva a ingresar otra fecha.", vbExclamation, modgen_g_con_OpeTra
         Call gs_SetFocus(ipp_FecPre)
         Screen.MousePointer = 0
         Exit Sub
      End If
      
      'Graba datos en cabecera de tabla prepagos
      If fs_usp_cre_ppgcab = 1 Then
         MsgBox "No se pudo completar el procedimiento 'usp_cre_ppgcab'.", vbCritical, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_Grabar)
         Screen.MousePointer = 0
         Exit Sub
      End If
      
      'Graba datos en detalle de tabla prepagos
      If fs_usp_cre_ppgdet = 1 Then
         MsgBox "No se pudo completar el procedimiento 'usp_cre_ppgdet'.", vbCritical, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_Grabar)
         Screen.MousePointer = 0
         Exit Sub
      End If
      
      'Graba datos en detalle de tabla cuotas
      If fs_ppgpar_hipcuo = 1 Then
         Call gs_SetFocus(cmd_Grabar)
         Screen.MousePointer = 0
         Exit Sub
      End If
      
      'Actualiza maestro de creditos hipotecarios
      If fs_update_hipmae = 1 Then
         MsgBox "No se pudo completar el procedimiento usp_ppgpar_cre_hipmae.", vbCritical, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_Grabar)
         Screen.MousePointer = 0
         Exit Sub
      End If
      
      'Actualiza estado en tabla prepagos para seguimientos
      If (InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) = 0) Or (InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) = 0) Or (InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) = 0) Then
         If fs_usp_actualiza_cre_ppgcab = 1 Then
            MsgBox "No se pudo completar el procedimiento 'usp_actualiza_cre_ppgcab'.", vbCritical, modgen_g_str_NomPlt
            Call gs_SetFocus(cmd_Grabar)
            Screen.MousePointer = 0
            Exit Sub
         End If
      End If
      
      'Enviando Correo usuarios
      Call fs_Envia_Correo
      
      'Enviando Correo plat
      Call fs_Envia_Correo_Plaft
      
      'Imprime liquidacion
      cmd_Grabar.Enabled = False
      cmd_Recalc.Enabled = False
      MsgBox "El proceso se grabó exitosamente.", vbInformation, modgen_g_str_NomPlt
      
      If InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then
         Call fs_PpgPar_Micasita(False, l_str_NomAch)
      Else
         Call fs_PpgPar_Mivivienda(False, l_str_NomAch)
      End If
         
      Screen.MousePointer = 0
      Unload Me
   End If
End Sub

Private Function fs_validar_ppg() As Integer
   fs_validar_ppg = 0
   
   'valida que las cuotas totales(cre_hipcuo) sea la suma cuotas pendientes y pagadas (cre_hipmae)
   If fs_validar_CuoTot(moddat_g_str_NumOpe) = 1 Then
      MsgBox "El total de Cuotas no corresponde a la Suma de las Cuotas Pagadas y Pendientes.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Grabar)
      fs_validar_ppg = 1
      Exit Function
   End If
   
   'valida que el nro. cuotas pendientes sea igual al nro. de cuotas de la grilla
   If fs_validar_CuoGrd(moddat_g_str_NumOpe) = 1 Then
      MsgBox "El total de Cuotas Pendientes no corresponde a las Cuotas del Cronograma Generado.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Grabar)
      fs_validar_ppg = 1
      Exit Function
   End If
   
   'valida que la sumatoria de capitales sean el Nuevo TNC
   If fs_validar_CapTNC = 1 Then
      MsgBox "El Total Capital, no corresponde al Nuevo TNC", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Grabar)
      fs_validar_ppg = 1
      Exit Function
   End If
   
   'valida que Capital + Saldo Capital de Primer Cuota sea el Nuevo TNC
    If fs_validar_CapSalTNC = 1 Then
      MsgBox "El Capital + Saldo Capital de 1º Cuota, no corresponde al Nuevo TNC.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Grabar)
      fs_validar_ppg = 1
      Exit Function
   End If
   
   'valida que el Saldo Capital de la última cuota del Cronograma generado sea cero.
    If fs_validar_SalCer = 1 Then
      MsgBox "El Saldo Capital de la última Cuota, no es cero.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Grabar)
      fs_validar_ppg = 1
      Exit Function
   End If
   
   'valida que el primer vencimiento (1º cuota pendiente de pago) sea igual a la 1ª Cuota del Cronograma Generado.
    If fs_validar_PriVct(moddat_g_str_NumOpe) = 1 Then
      MsgBox "El Primer Vencimiento, no es igual a la Fecha de Vencimiento de la 1ª Cuota del Cronograma Generado.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Grabar)
      fs_validar_ppg = 1
      Exit Function
   End If
   
   'valida que la primera cuota del Cronograma Generado sea la primera cuota pendiente de pago CRE_HIPCUO.
    If fs_validar_PriHip(moddat_g_str_NumOpe) = 1 Then
      MsgBox "El Primer Vencimiento, no es igual a la Fecha de Vencimiento de la 1ª Cuota del Cronograma Generado.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Grabar)
      fs_validar_ppg = 1
      Exit Function
   End If
   
   If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) = 1 Then
      'Solo para Red. de Monto
      'valida que la última cuota del Cronograma Generado debe ser igual a la última Cuota de la Cre_Hipcuo.
      If fs_validar_UltCuo(moddat_g_str_NumOpe) = 1 Then
         MsgBox "El Primer Vencimiento, no es igual a la Fecha de Vencimiento de la 1ª Cuota del Cronograma Generado.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_Grabar)
         fs_validar_ppg = 1
         Exit Function
      End If
      
      'valida que la último vencimiento del Cronograma Generado debe ser igual a la última Cuota de la Cre_Hipcuo.
      If fs_validar_UltVct(moddat_g_str_NumOpe) = 1 Then
         MsgBox "El Primer Vencimiento, no es igual a la Fecha de Vencimiento de la 1ª Cuota del Cronograma Generado.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_Grabar)
         fs_validar_ppg = 1
         Exit Function
      End If
   End If
End Function

'valida que las cuotas totales(cre_hipcuo) sea la suma cuotas pendientes y pagadas (cre_hipmae)
Private Function fs_validar_CuoTot(p_NumOpe As String) As Integer
Dim r_str_Parame  As String
Dim r_rst_AuxCon  As Recordset

   fs_validar_CuoTot = 0
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "  SELECT (CASE WHEN (A.HIPMAE_CUOPAG + A.HIPMAE_CUOPEN) = B.CUOTAS THEN 0 ELSE 1 END) AS TOT_CUOTA"
   r_str_Parame = r_str_Parame & "    FROM CRE_HIPMAE A "
   r_str_Parame = r_str_Parame & "         LEFT JOIN (SELECT HIPCUO_NUMOPE, COUNT(*) AS CUOTAS "
   r_str_Parame = r_str_Parame & "                      FROM CRE_HIPCUO B "
   r_str_Parame = r_str_Parame & "                     WHERE HIPCUO_TIPCRO = 1"
   r_str_Parame = r_str_Parame & "                     GROUP BY HIPCUO_NUMOPE) B ON A.HIPMAE_NUMOPE = B.HIPCUO_NUMOPE"
   r_str_Parame = r_str_Parame & "   WHERE A.HIPMAE_NUMOPE = '" & p_NumOpe & "'"

   If gf_EjecutaSQL(r_str_Parame, r_rst_AuxCon, 3) Then
      fs_validar_CuoTot = r_rst_AuxCon!TOT_CUOTA
   End If
End Function

'valida que el nro. cuotas pendientes sea igual al nro. de cuotas de la grilla
Private Function fs_validar_CuoGrd(p_NumOpe As String) As Integer
Dim r_int_NumCuo  As Integer
Dim r_str_Parame  As String
Dim r_rst_AuxCon  As Recordset

   fs_validar_CuoGrd = 0
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT COUNT(HIPCUO_NUMOPE) AS CUOPEND "
   r_str_Parame = r_str_Parame & "   FROM CRE_HIPCUO A "
   r_str_Parame = r_str_Parame & "  WHERE A.HIPCUO_NUMOPE = '" & p_NumOpe & "'"
   r_str_Parame = r_str_Parame & "    AND A.HIPCUO_TIPCRO = 1 "
   r_str_Parame = r_str_Parame & "    AND A.HIPCUO_SITUAC = 2 "

   If gf_EjecutaSQL(r_str_Parame, r_rst_AuxCon, 3) Then
      If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) = 2 Then
         'Reduccion de plazo
         r_int_NumCuo = cmb_RedPlz.ItemData(cmb_RedPlz.ListIndex) * 12
         If grd_CliNCo_Listad.Rows = CStr(r_rst_AuxCon!CUOPEND - r_int_NumCuo) Then
            fs_validar_CuoGrd = 0
         Else
            fs_validar_CuoGrd = 1
         End If
      Else
         'Reduccion de monto
         If grd_CliNCo_Listad.Rows = CStr(r_rst_AuxCon!CUOPEND) Then
            fs_validar_CuoGrd = 0
         Else
            fs_validar_CuoGrd = 1
         End If
      End If
   End If
End Function
  
'valida que la sumatoria de capitales sean el Nuevo TNC
Private Function fs_validar_CapTNC() As Integer
   fs_validar_CapTNC = 0

   If CDbl(CStr(pnl_NuevoSaldoTNC.Caption)) = CDbl(CStr(pnl_CliNCo_Capita.Caption)) Then
      fs_validar_CapTNC = 0
   Else
      fs_validar_CapTNC = 1
   End If
End Function

'valida que Capital + Saldo Capital de Primer Cuota sea el Nuevo TNC
Private Function fs_validar_CapSalTNC() As Integer
Dim r_dbl_Capita  As Double
Dim r_dbl_SalCap  As Double

   fs_validar_CapSalTNC = 0
   r_dbl_Capita = CDbl(grd_CliNCo_Listad.TextMatrix(0, 2))
   r_dbl_SalCap = CDbl(grd_CliNCo_Listad.TextMatrix(0, 8))
   
   If Trim(CDbl(Trim(pnl_NuevoSaldoTNC.Caption))) = Trim(r_dbl_Capita + r_dbl_SalCap) Then
      fs_validar_CapSalTNC = 0
   Else
      fs_validar_CapSalTNC = 1
   End If
End Function

'valida que el Saldo Capital de la última cuota del Cronograma generado sea cero.
Private Function fs_validar_SalCer() As Integer
Dim r_dbl_SalCap  As Double

   fs_validar_SalCer = 0
   r_dbl_SalCap = grd_CliNCo_Listad.TextMatrix(grd_CliNCo_Listad.Rows - 1, 8)
   
   If r_dbl_SalCap = 0 Then
      fs_validar_SalCer = 0
   Else
      fs_validar_SalCer = 1
   End If
End Function

'valida que el primer vencimiento (1º cuota pendiente de pago) sea igual a la 1ª Cuota del Cronograma Generado.
Private Function fs_validar_PriVct(p_NumOpe As String) As Integer
Dim r_str_Parame  As String
Dim r_rst_AuxCon  As Recordset

   fs_validar_PriVct = 0
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT HIPCUO_FECVCT AS PRIMER_VCTO "
   r_str_Parame = r_str_Parame & "   FROM CRE_HIPCUO A "
   r_str_Parame = r_str_Parame & "  WHERE A.HIPCUO_NUMOPE = '" & p_NumOpe & "'"
   r_str_Parame = r_str_Parame & "    AND A.HIPCUO_TIPCRO = 1 "
   r_str_Parame = r_str_Parame & "    AND A.HIPCUO_SITUAC = 2 "
   r_str_Parame = r_str_Parame & "    AND ROWNUM = 1 "

   If gf_EjecutaSQL(r_str_Parame, r_rst_AuxCon, 3) Then
      If grd_CliNCo_Listad.TextMatrix(0, 1) = gf_FormatoFecha(r_rst_AuxCon!PRIMER_VCTO) Then
         fs_validar_PriVct = 0
      Else
         fs_validar_PriVct = 1
      End If
   End If
End Function

'valida que la primera cuota del Cronograma Generado sea la primera cuota pendiente de pago CRE_HIPCUO.
Private Function fs_validar_PriHip(p_NumOpe As String) As Integer
Dim r_str_Parame  As String
Dim r_rst_AuxCon  As Recordset

   fs_validar_PriHip = 0
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT HIPCUO_NUMCUO AS PRIMER_CUOTA "
   r_str_Parame = r_str_Parame & "   FROM CRE_HIPCUO A "
   r_str_Parame = r_str_Parame & "  WHERE A.HIPCUO_NUMOPE = '" & p_NumOpe & "'"
   r_str_Parame = r_str_Parame & "    AND A.HIPCUO_TIPCRO = 1 "
   r_str_Parame = r_str_Parame & "    AND A.HIPCUO_SITUAC = 2 "
   r_str_Parame = r_str_Parame & "    AND ROWNUM = 1 "
   
   If gf_EjecutaSQL(r_str_Parame, r_rst_AuxCon, 3) Then
      If CInt(grd_CliNCo_Listad.TextMatrix(0, 0)) = r_rst_AuxCon!PRIMER_CUOTA Then
         fs_validar_PriHip = 0
      Else
         fs_validar_PriHip = 1
      End If
   End If
End Function

'Solo para Red. de Monto
'valida que la ùltima cuota del Cronograma Generado debe ser igual a la última Cuota de la Cre_Hipcuo.
Private Function fs_validar_UltCuo(p_NumOpe As String) As Integer
Dim r_str_Parame  As String
Dim r_rst_AuxCon  As Recordset

   fs_validar_UltCuo = 0
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT HIPCUO_NUMCUO AS ULT_CUOTA "
   r_str_Parame = r_str_Parame & "   FROM CRE_HIPCUO A "
   r_str_Parame = r_str_Parame & "  WHERE A.HIPCUO_NUMOPE = '" & p_NumOpe & "'"
   r_str_Parame = r_str_Parame & "    AND A.HIPCUO_TIPCRO = 1 "
   r_str_Parame = r_str_Parame & "    AND A.HIPCUO_SITUAC = 2 "
   r_str_Parame = r_str_Parame & "    AND ROWNUM = 1 "
   r_str_Parame = r_str_Parame & "  ORDER BY HIPCUO_NUMCUO DESC "

   If gf_EjecutaSQL(r_str_Parame, r_rst_AuxCon, 3) Then
      If CDbl(grd_CliNCo_Listad.TextMatrix(grd_CliNCo_Listad.Rows - 1, 0)) = r_rst_AuxCon!ULT_CUOTA Then
         fs_validar_UltCuo = 0
      Else
         fs_validar_UltCuo = 1
      End If
   End If
End Function

'Solo para Red. de Monto
'valida que la ùltima cuota del Cronograma Generado debe ser igual a la última Cuota de la Cre_Hipcuo.
Private Function fs_validar_UltVct(p_NumOpe As String) As Integer
Dim r_str_Parame  As String
Dim r_rst_AuxCon  As Recordset

   fs_validar_UltVct = 0
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT HIPCUO_FECVCT AS ULT_VCT "
   r_str_Parame = r_str_Parame & "   FROM CRE_HIPCUO A "
   r_str_Parame = r_str_Parame & "  WHERE A.HIPCUO_NUMOPE = '" & p_NumOpe & "'"
   r_str_Parame = r_str_Parame & "    AND A.HIPCUO_TIPCRO = 1 "
   r_str_Parame = r_str_Parame & "    AND A.HIPCUO_SITUAC = 2 "
   r_str_Parame = r_str_Parame & "    AND ROWNUM = 1 "
   r_str_Parame = r_str_Parame & "  ORDER BY HIPCUO_NUMCUO DESC "

   If gf_EjecutaSQL(r_str_Parame, r_rst_AuxCon, 3) Then
      If grd_CliNCo_Listad.TextMatrix(grd_CliNCo_Listad.Rows - 1, 1) = gf_FormatoFecha(r_rst_AuxCon!ULT_VCT) Then
         fs_validar_UltVct = 0
      Else
         fs_validar_UltVct = 1
      End If
   End If
End Function

Private Sub cmd_Limpia_Click()
   Call fs_Inicia
   Call fs_Inicia_Cronog
   Call fs_Limpia_Cronog
   Call fs_Buscar
   Call fs_Buscar_Tasacion
   cmd_Recalc.Enabled = True
   cmd_Grabar.Enabled = False
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_Recalc_Click()
Dim r_dbl_MtoApl     As Double
Dim r_dbl_AplPpg     As Double
Dim r_int_RedPlz     As Integer
Dim r_dbl_MtoApl_Fin As Double

   'Validaciones
   If cmb_TipPre.ListIndex = -1 Then
      MsgBox "Seleccione el tipo de prepago.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipPre)
      Exit Sub
   End If
   If CDbl(txt_Mto_Deposito) = 0 Then
      MsgBox "Ingrese el monto a depositar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Mto_Deposito)
      Exit Sub
   End If
   If CDbl(txt_Mto_Deposito) >= CDbl(pnl_SaldoTNC1.Caption) + CDbl(pnl_SaldoTC1.Caption) + CDbl(txt_InteresTC.Text) + CDbl(txt_InteresTNC.Text) + CDbl(Me.txt_SegDes.Text) Then
      MsgBox "Favor verifique el monto depositado, puede aplicar un prepago total", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Mto_Deposito)
      Exit Sub
   End If
   If CDbl(pnl_NuevoSaldoTNC.Caption) <= 0 Then
      MsgBox "Favor verifique que el monto de 'Aplicacion PP TNC' sea positivo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_AplTNC)
      Exit Sub
   End If
   If CInt(pnl_DiasTNC.Caption) > 30 Then
      MsgBox "El campo 'Dias TNC' no puede ser mayor a 30.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecPre)
      Exit Sub
   End If
   If Not (InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0) Then
      'Se Comenta por solicitud de JMARANGUREN - 15/09/2015
      'If CDbl(pnl_NuevoSaldoTC.Caption) <= 0 Then
      '   MsgBox "Favor verifique que el monto de 'Aplicacion PP TC' sea positivo.", vbExclamation, modgen_g_str_NomPlt
      '   Call gs_SetFocus(txt_ApliTC)
      '   Exit Sub
      'End If
      
      'Se Comenta por solicitud de RRICALDI - 19/09/2013
      'If CInt(pnl_DiasTC.Caption) > 180 Then
      '   MsgBox "El campo 'Dias TC' no puede ser mayor a 180.", vbExclamation, modgen_g_str_NomPlt
      '   Call gs_SetFocus(ipp_FecPre)
      '   Exit Sub
      'End If
   End If
   If CDbl(pnl_MtoApl) = 0 Then
      MsgBox "El monto de aplicación del prepago es incorrecto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Mto_Deposito)
      Exit Sub
   End If
   If CDbl(pnl_CuoPen.Caption) = 0 Then
      MsgBox "Las cuotas pendientes de pago no pueden ser cero.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   If CDbl(txt_AplTNC) = 0 Then
      MsgBox "El monto de aplicación TNC del prepago es incorrecto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_AplTNC)
      Exit Sub
   End If
   If CDbl(pnl_NuevoSaldoTNC.Caption) < 0 Then
      MsgBox "El monto del Nuevo Saldo TNC no puede ser menor a cero.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_AplTNC)
      Exit Sub
  
   End If
   If CDbl(pnl_NuevoSaldoTC.Caption) < 0 Then
      MsgBox "El monto del Nuevo Saldo TC no puede ser menor a cero.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_AplTNC)
      Exit Sub
   End If
      
   'agregar
   Call fs_Cal_Prpago
   
   'Validaciones de reduccion de plazo
   If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) = 2 Then
      If cmb_RedPlz.ListIndex = -1 Then
         MsgBox "Seleccione el numero de años a reducir.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_RedPlz)
         Exit Sub
      End If
      
      r_int_RedPlz = cmb_RedPlz.ItemData(cmb_RedPlz.ListIndex) * 12
      
      'productos micasita
      If InStr(moddat_g_str_AgrTMIC, moddat_g_str_CodPrd) > 0 Then
         'productos micasita
         If l_int_PagCuo + (CInt(Trim(pnl_CuoPen.Caption)) - r_int_RedPlz) < 60 Then
            MsgBox "Para productos micasita el plazo minimo es 5 años, no puede realizar prepagos con reduccion de plazo.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_RedPlz)
            Exit Sub
         End If
      Else
         'productos mivivienda
         If l_int_PagCuo + (CInt(Trim(pnl_CuoPen.Caption)) - r_int_RedPlz) < 60 Then
            MsgBox "Para productos mivivienda el plazo minimo es 5 años, no puede realizar prepagos con reduccion de plazo.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_RedPlz)
            Exit Sub
         End If
      End If
   End If
      
   r_dbl_MtoApl = CDbl(txt_Mto_Deposito.Text) - CDbl(txt_InteresTNC.Text) - CDbl(txt_InteresTC.Text) - CDbl(txt_SegDes.Text) - CDbl(txt_SegInm.Text) - CDbl(txt_MontoITF.Text) - CDbl(Trim(pnl_IntPbp.Caption)) - CDbl(Trim(pnl_DeuPen.Caption))
   r_dbl_MtoApl_Fin = CDbl(r_dbl_MtoApl) - CDbl(Trim(pnl_CapPbp.Caption))
   r_dbl_AplPpg = CDbl(txt_AplTNC.Text) + CDbl(txt_ApliTC.Text)
   
   If Round(CDbl(r_dbl_MtoApl_Fin), 2) <> Round(CDbl(r_dbl_AplPpg), 2) Then 'r_dbl_MtoApl
      MsgBox "El monto depositado menos los interes y gastos debe ser igual al monto de aplicacion del prepago.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_AplTNC)
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Recalc
   
   cmd_ExpExc.Enabled = True
   If CDbl(pnl_NuevaCuota.Caption) > l_dbl_CuoFij Then
      If CDbl(pnl_NuevaCuota.Caption) - l_dbl_CuoFij > 5 Then
         MsgBox "La nueva cuota generada es mayor que la cuota del cronograma actual.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Mto_Deposito)
         cmd_ExpExc.Enabled = False
      Else
         If MsgBox("La nueva cuota generada es mayor que la cuota del cronograma actual." & vbCrLf & "¿Está seguro de continuar?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Call gs_SetFocus(txt_Mto_Deposito)
            cmd_ExpExc.Enabled = False
         End If
      End If
   End If
   
   If UCase(App.EXEName) = "OPETRA" Then
      cmd_Grabar.Enabled = True
   End If
   Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Inicia_Cronog
   Call fs_Limpia_Cronog
   Call fs_Buscar_Cuotas_Vencidas
   Call fs_Buscar
   Call fs_Buscar_Tasacion
   Call gs_CentraForm(Me)
   cmd_Grabar.Enabled = False
   cmd_ForCon.Enabled = False
   
   Call gs_SetFocus(ipp_FecPre)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Grid de Datos del Crédito
   grd_Listad.ColWidth(0) = 2900
   grd_Listad.ColWidth(1) = 8150
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   
   'Cuotas Pendientes
   grd_DeuPen.ColWidth(0) = 575
   grd_DeuPen.ColWidth(1) = 630
   grd_DeuPen.ColWidth(2) = 1020
   grd_DeuPen.ColWidth(3) = 0  '735
   grd_DeuPen.ColWidth(4) = 965
   grd_DeuPen.ColWidth(5) = 965
   grd_DeuPen.ColWidth(6) = 965
   grd_DeuPen.ColWidth(7) = 965
   grd_DeuPen.ColWidth(8) = 840
   grd_DeuPen.ColWidth(9) = 965
   grd_DeuPen.ColWidth(10) = 965
   grd_DeuPen.ColWidth(11) = 0 '965
   grd_DeuPen.ColWidth(12) = 0 '965
   grd_DeuPen.ColWidth(13) = 965
   grd_DeuPen.ColWidth(14) = 0 '965
   grd_DeuPen.ColWidth(15) = 965
   grd_DeuPen.ColAlignment(0) = flexAlignCenterCenter
   grd_DeuPen.ColAlignment(1) = flexAlignCenterCenter
   grd_DeuPen.ColAlignment(2) = flexAlignCenterCenter
   grd_DeuPen.ColAlignment(3) = flexAlignCenterCenter
   grd_DeuPen.ColAlignment(4) = flexAlignRightCenter
   grd_DeuPen.ColAlignment(5) = flexAlignRightCenter
   grd_DeuPen.ColAlignment(6) = flexAlignRightCenter
   grd_DeuPen.ColAlignment(7) = flexAlignRightCenter
   grd_DeuPen.ColAlignment(8) = flexAlignRightCenter
   grd_DeuPen.ColAlignment(9) = flexAlignRightCenter
   grd_DeuPen.ColAlignment(10) = flexAlignRightCenter
   grd_DeuPen.ColAlignment(11) = flexAlignRightCenter
   grd_DeuPen.ColAlignment(12) = flexAlignRightCenter
   grd_DeuPen.ColAlignment(13) = flexAlignRightCenter
   grd_DeuPen.ColAlignment(14) = flexAlignRightCenter
   grd_DeuPen.ColAlignment(15) = flexAlignRightCenter
   Call gs_LimpiaGrid(grd_DeuPen)
   
   cmb_TipPre.Clear
   cmb_TipPre.AddItem "RED. MONTO"
   cmb_TipPre.ItemData(cmb_TipPre.NewIndex) = 1
   cmb_TipPre.AddItem "RED. PLAZO"
   cmb_TipPre.ItemData(cmb_TipPre.NewIndex) = 2
   cmb_TipPre.ListIndex = -1
   
   cmb_RedPlz.Clear
   cmb_RedPlz.AddItem "1 AÑO"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 1
   cmb_RedPlz.AddItem "2 AÑOS"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 2
   cmb_RedPlz.AddItem "3 AÑOS"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 3
   cmb_RedPlz.AddItem "4 AÑOS"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 4
   cmb_RedPlz.AddItem "5 AÑOS"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 5
   cmb_RedPlz.AddItem "6 AÑOS"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 6
   cmb_RedPlz.AddItem "7 AÑOS"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 7
   cmb_RedPlz.AddItem "8 AÑOS"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 8
   cmb_RedPlz.AddItem "9 AÑOS"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 9
   cmb_RedPlz.AddItem "10 AÑOS"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 10
   cmb_RedPlz.AddItem "11 AÑOS"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 11
   cmb_RedPlz.AddItem "12 AÑOS"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 12
   cmb_RedPlz.AddItem "13 AÑOS"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 13
   cmb_RedPlz.AddItem "14 AÑOS"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 14
   cmb_RedPlz.ListIndex = -1
   cmb_RedPlz.Enabled = False
   
   txt_Mto_Deposito.Text = "0.00 "
   pnl_Val_AsgInm.Caption = "0.00 "
   pnl_SaldoTNC1.Caption = "0.00 "
   pnl_SaldoTC1.Caption = "0.00 "
   pnl_UltPagTNC.Caption = " "
   pnl_UltPagTC.Caption = " "
   pnl_DiasTNC.Caption = "0 "
   pnl_DiasTC.Caption = "0 "
   txt_InteresTNC.Text = 0
   txt_InteresTC.Text = 0
   txt_SegDes.Text = 0
   txt_SegInm.Text = 0
   txt_MontoITF.Text = 0
   pnl_CuoPen.Caption = "0 "
   pnl_MtoApl.Caption = "0.00 "
   pnl_SaldoTNC2.Caption = "0.00 "
   pnl_SaldoTC2.Caption = "0.00 "
   txt_AplTNC.Text = 0
   txt_ApliTC.Text = 0
   pnl_IntPbp.Caption = "0.00 "
   pnl_CapPbp.Caption = "0.00 "
   pnl_NuevoSaldoTNC.Caption = "0.00 "
   pnl_NuevoSaldoTC.Caption = "0.00 "
   pnl_NuevaCuota.Caption = "0.00 "
   pnl_DeuPen.Caption = "0.00 "
   pnl_MtoApl_Fin.Caption = "0.00 "
   ipp_FecPre.DateValue = date
   
   If UCase(App.EXEName) = "OPETRA" Then
      txt_InteresTNC.Enabled = True
      txt_InteresTC.Enabled = True
      txt_SegDes.Enabled = True
      txt_SegInm.Enabled = True
      txt_MontoITF.Enabled = True
      txt_AplTNC.Enabled = True
      txt_ApliTC.Enabled = True
   Else
      txt_InteresTNC.Enabled = False
      txt_InteresTC.Enabled = False
      txt_SegDes.Enabled = False
      txt_SegInm.Enabled = False
      txt_MontoITF.Enabled = False
      txt_AplTNC.Enabled = False
      txt_ApliTC.Enabled = False
   End If
   
   txt_ObsPpg.Text = " "
   cmb_MotPpg.ListIndex = -1
   Call moddat_gs_Carga_LisIte_Combo(cmb_MotPpg, 1, "115")
   Call gs_LimpiaGrid(grd_Listad)
   
   moddat_g_str_CodIte = ""
   moddat_g_str_TipPar = ""
   moddat_g_str_DesObs = ""
End Sub

Private Sub fs_Inicia_Cronog()
   'Cliente No Concesional
   grd_CliNCo_Listad.ColWidth(0) = 795
   grd_CliNCo_Listad.ColWidth(1) = 1425
   grd_CliNCo_Listad.ColWidth(2) = 1180
   grd_CliNCo_Listad.ColWidth(3) = 1170
   grd_CliNCo_Listad.ColWidth(4) = 1160
   grd_CliNCo_Listad.ColWidth(5) = 1160
   grd_CliNCo_Listad.ColWidth(6) = 1160
   grd_CliNCo_Listad.ColWidth(7) = 1320
   grd_CliNCo_Listad.ColWidth(8) = 1560
   
   'Cliente Concesional
   grd_CliCon_Listad.ColWidth(0) = 770
   grd_CliCon_Listad.ColWidth(1) = 1485
   grd_CliCon_Listad.ColWidth(2) = 2170
   grd_CliCon_Listad.ColWidth(3) = 2160
   grd_CliCon_Listad.ColWidth(4) = 2170
   grd_CliCon_Listad.ColWidth(5) = 2170
   grd_CliCon_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CliCon_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CliCon_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(5) = flexAlignRightCenter
End Sub

Private Sub fs_Limpia_Cronog()
   Call gs_LimpiaGrid(grd_CliNCo_Listad)
   Call gs_LimpiaGrid(grd_CliCon_Listad)
   
   pnl_CliNCo_Capita.Caption = "0.00 "
   pnl_CliNCo_Intere.Caption = "0.00 "
   pnl_CliNCo_SegPre.Caption = "0.00 "
   pnl_CliNCo_SegViv.Caption = "0.00 "
   pnl_CliNCo_OtrCar.Caption = "0.00 "
   pnl_CliNCo_TotCuo.Caption = "0.00 "
   pnl_CliCon_Capita.Caption = "0.00 "
   pnl_CliCon_Intere.Caption = "0.00 "
   pnl_CliCon_TotCuo.Caption = "0.00 "
End Sub

Private Sub fs_Buscar()
Dim r_str_CodPry     As String
Dim r_str_NomPry     As String
Dim r_str_CodBco     As String
Dim r_dbl_CapPBP     As Double
Dim r_dbl_IntPBP     As Double
Dim r_bol_CuoVen     As Boolean
   
   'Buscando Información del Crédito
   Call modmip_gs_DatNumOpe(moddat_g_str_NumOpe, grd_Listad)
    
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT * FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & "  WHERE HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "    AND (HIPMAE_SITUAC = 2)"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst

   'DATOS DE LA OPERACION DEL PREPAGO
   l_dbl_TasInt = g_rst_Princi!HIPMAE_TASINT
   l_dbl_SegDes = g_rst_Princi!HIPMAE_FOIPRE
   l_dbl_SegInm = g_rst_Princi!HIPMAE_FOIVIV
   l_int_TipSeg = g_rst_Princi!HIPMAE_TIPSEG  'seguro desgravamen

   l_int_CodPrd = g_rst_Princi!HIPMAE_CODPRD
   l_int_NumCuo = g_rst_Princi!HIPMAE_NUMCUO
   l_int_PagCuo = g_rst_Princi!HIPMAE_CUOPAG
   l_int_PerGra = g_rst_Princi!HIPMAE_PERGRA
   l_str_UltVct = g_rst_Princi!HIPMAE_UlTVCT
   l_dbl_CuoFij = CDbl(g_rst_Princi!HIPMAE_CUOFIJ)
   l_dbl_PagCap = CDbl(g_rst_Princi!HIPMAE_PAGCAP)
   l_dbl_PagInt = CDbl(g_rst_Princi!HIPMAE_PAGINT)
   l_int_NCuota = g_rst_Princi!HIPMAE_NUMCUO
   
   pnl_CuoPen.Caption = CInt(g_rst_Princi!HIPMAE_CUOPEN) & " "
   
   If pnl_DeuPen.Caption > 0# Then
      r_bol_CuoVen = True
      l_dbl_SalNco = fs_Obtiene_Saldos(g_rst_Princi!HIPMAE_NUMOPE, 1, True)
      l_dbl_SalCon = fs_Obtiene_Saldos(g_rst_Princi!HIPMAE_NUMOPE, 2, True)
      pnl_CuoPen.Caption = CInt(g_rst_Princi!HIPMAE_CUOPEN) - CInt(Me.grd_DeuPen.Rows - 1) & " "
      pnl_UltPagTNC.Caption = gf_FormatoFecha(fs_Obtiene_FechaPago(g_rst_Princi!HIPMAE_NUMOPE, 1, g_rst_Princi!HIPMAE_FECDES, True))
      pnl_UltPagTC.Caption = gf_FormatoFecha(fs_Obtiene_FechaPago(g_rst_Princi!HIPMAE_NUMOPE, 2, g_rst_Princi!HIPMAE_FECDES, True))
   Else
      r_bol_CuoVen = False
      l_dbl_SalNco = g_rst_Princi!HIPMAE_SALCAP
      l_dbl_SalCon = g_rst_Princi!HIPMAE_SALCON
      pnl_UltPagTNC.Caption = gf_FormatoFecha(fs_Obtiene_FechaPago(g_rst_Princi!HIPMAE_NUMOPE, 1, g_rst_Princi!HIPMAE_FECDES, False))
      pnl_UltPagTC.Caption = gf_FormatoFecha(fs_Obtiene_FechaPago(g_rst_Princi!HIPMAE_NUMOPE, 2, g_rst_Princi!HIPMAE_FECDES, False))
   End If
   
   pnl_DiasTNC.Caption = DateDiff("d", pnl_UltPagTNC.Caption, ipp_FecPre.Text) & " "
   pnl_DiasTC.Caption = DateDiff("d", pnl_UltPagTC.Caption, ipp_FecPre.Text) & " "
   pnl_SaldoTNC1.Caption = Format(l_dbl_SalNco, "###,###.00") & " "
   pnl_SaldoTC1.Caption = Format(l_dbl_SalCon, "###,###.00") & " "
   pnl_SaldoTNC2.Caption = Format(l_dbl_SalNco, "###,###.00") & " "
   pnl_SaldoTC2.Caption = Format(l_dbl_SalCon, "###,###.00") & " "
   
   If pnl_DeuPen.Caption > 0# Then
      Call fs_Obtiene_PBPPerdido(g_rst_Princi!HIPMAE_NUMOPE, r_dbl_CapPBP, r_dbl_IntPBP, True)
   Else
      Call fs_Obtiene_PBPPerdido(g_rst_Princi!HIPMAE_NUMOPE, r_dbl_CapPBP, r_dbl_IntPBP, False)
   End If
   
   pnl_CapPbp.Caption = Format(r_dbl_CapPBP, "###,###.00") & " "
   pnl_IntPbp.Caption = Format(r_dbl_IntPBP, "###,###.00") & " "
   l_dbl_PorNco = Format(g_rst_Princi!HIPMAE_IMPNCO / g_rst_Princi!HIPMAE_TOTPRE, "##0.0000")
   l_dbl_PorCon = 1 - l_dbl_PorNco
   l_str_PrxVct = g_rst_Princi!HIPMAE_PRXVCT
   
   'DETERMINA SI OPERACION ES MICASITA O MIVIVIENDA (UN SOLO TRAMO)
   If InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then
      l_dbl_PorITF = opecaj_gf_Consulta_ITF(Format(CDate(moddat_g_str_FecSis), "yyyymmdd"), 1)
      l_dbl_PorNco = 1
      l_dbl_PorCon = 0
      Label21.Caption = "Saldo Actual"
      Label23.Caption = "Ultimo Pago"
      Label25.Caption = "Dias"
      Label6.Caption = "Interés a la fecha"
      Label20.Visible = False
      pnl_SaldoTC1.Visible = False
      Label22.Visible = False
      pnl_UltPagTC.Visible = False
      Label24.Visible = False
      pnl_DiasTC.Visible = False
      Label7.Visible = False
      txt_InteresTC.Visible = False
      Label28.Caption = "Saldo Actual"
      Label8.Caption = "Aplicación PP"
      Label10.Caption = "Nuevo Saldo"
      Label33.Visible = False
      pnl_IntPbp.Visible = False
      Label32.Visible = False
      pnl_CapPbp.Visible = False
      Label27.Visible = False
      pnl_SaldoTC2.Visible = False
      Label13.Visible = False
      txt_ApliTC.Visible = False
      Label16.Visible = False
      pnl_NuevoSaldoTC.Visible = False
      tab_Cronog.TabVisible(0) = True
      tab_Cronog.TabVisible(1) = False
      tab_Cronog.TabCaption(0) = "Cliente"
      tab_Cronog.TabCaption(1) = ""
   Else
      Label21.Caption = "Saldo Actual TNC"
      Label23.Caption = "Ultimo Pago TNC"
      Label25.Caption = "Dias TNC"
      Label6.Caption = "Interés TNC a la fecha"
      Label20.Visible = True
      pnl_SaldoTC1.Visible = True
      Label22.Visible = True
      pnl_UltPagTC.Visible = True
      Label24.Visible = True
      pnl_DiasTC.Visible = True
      Label7.Visible = True
      txt_InteresTC.Visible = True
      Label28.Caption = "Saldo Actual TNC"
      Label8.Caption = "Aplica. PP TNC"
      Label10.Caption = "Nuevo TNC"
      Label27.Visible = True
      pnl_SaldoTC2.Visible = True
      Label13.Visible = True
      txt_ApliTC.Visible = True
      Label16.Visible = True
      Label33.Visible = True
      pnl_IntPbp.Visible = True
      Label32.Visible = True
      pnl_CapPbp.Visible = True
      pnl_NuevoSaldoTC.Visible = True
      tab_Cronog.TabVisible(0) = True
      tab_Cronog.TabVisible(1) = True
      tab_Cronog.TabCaption(0) = "Cliente - No Concesional"
      tab_Cronog.TabCaption(1) = "Cliente - Concesional"
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Call gs_UbiIniGrid(grd_Listad)
End Sub

Private Sub fs_Buscar_ant()
Dim r_str_CodPry     As String
Dim r_str_NomPry     As String
Dim r_str_CodBco     As String
Dim r_dbl_CapPBP     As Double
Dim r_dbl_IntPBP     As Double
Dim r_bol_CuoVen     As Boolean
  
   'Buscando Información del Crédito
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT * FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & "  WHERE HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "    AND (HIPMAE_SITUAC = 2)"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   'Almacenando en Variables Globales
   moddat_g_int_TipDoc = g_rst_Princi!HIPMAE_TDOCLI
   moddat_g_str_NumDoc = Trim(g_rst_Princi!HIPMAE_NDOCLI)
   moddat_g_str_NumSol = Trim(g_rst_Princi!HIPMAE_NUMSOL)
   moddat_g_str_NumOpe = Trim(g_rst_Princi!HIPMAE_NUMOPE)
   
   'Obteniendo Nombre de Cliente
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   'Obteniendo Nombre y DOI de Cónyuge
   moddat_g_int_CygTDo = g_rst_Princi!HIPMAE_TDOCYG
   moddat_g_str_CygNDo = ""
   moddat_g_str_CygNom = ""
   
   If moddat_g_int_CygTDo > 0 Then
      moddat_g_str_CygNDo = Trim(g_rst_Princi!HIPMAE_NDOCYG & "")
      moddat_g_str_CygNom = moddat_gf_Buscar_NomCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo)
   End If
   
   'Obteniendo Descripción de Producto
   moddat_g_str_CodPrd = Trim(g_rst_Princi!HIPMAE_CODPRD)
   moddat_g_str_NomPrd = moddat_gf_Consulta_Produc(Trim(g_rst_Princi!HIPMAE_CODPRD))
   moddat_g_str_CodSub = Trim(g_rst_Princi!HIPMAE_CODSUB)

   'Obeniendo Modalidad de Producto
   moddat_g_str_CodMod = Trim(g_rst_Princi!HIPMAE_CODMOD)
   moddat_g_str_DesMod = moddat_gf_Buscar_NomMod(Trim(g_rst_Princi!HIPMAE_CODPRD), moddat_g_str_CodMod)
   
   'Ejecutivo de Seguimiento
   moddat_g_str_CodEjeSeg = Trim(g_rst_Princi!HIPMAE_EJESEG & "")
   moddat_g_str_NomEjeSeg = moddat_gf_Buscar_NomEje(moddat_g_str_CodEjeSeg)
   
   'Consejero Hipotecario
   moddat_g_str_CodConHip = Trim(g_rst_Princi!HIPMAE_CONHIP & "")
   moddat_g_str_NomConHip = moddat_gf_Buscar_NomEje(moddat_g_str_CodConHip)
   
   'Moneda
   moddat_g_str_Moneda = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPMAE_MONEDA))
   moddat_g_int_TipMon = g_rst_Princi!HIPMAE_MONEDA                           'Moneda Préstamo
   moddat_g_dbl_MtoPre = g_rst_Princi!HIPMAE_MTOPRE                           'Monto Préstamo
   moddat_g_int_CuoPen = g_rst_Princi!HIPMAE_CUOPEN                           'Cuotas Pendientes
   moddat_g_int_TotCuo = g_rst_Princi!HIPMAE_NUMCUO                           'Total de Cuotas
   moddat_g_dbl_SalCap = g_rst_Princi!HIPMAE_SALCAP                           'Saldo Capital
   moddat_g_str_FecApr = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))    'Fecha Desembolso
   
   'Situación de Crédito
   moddat_g_int_Situac = g_rst_Princi!HIPMAE_SITUAC
   moddat_g_str_Situac = moddat_gf_Consulta_ParDes("027", CStr(g_rst_Princi!HIPMAE_SITUAC))
   
   'Obteniendo Información del Inmueble
   Call moddat_gs_Consulta_DatInm(moddat_g_str_NumSol, moddat_g_str_Direcc, moddat_g_str_Distri, r_str_CodPry, r_str_NomPry, r_str_CodBco)
   
   'Cargando en Grid
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.CellFontBold = True
   grd_Listad.Text = "Número de Operación"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontBold = True
   grd_Listad.Text = gf_Formato_NumOpe(g_rst_Princi!HIPMAE_NUMOPE)
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.CellFontBold = True
   grd_Listad.Text = "Situación"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontBold = True
   grd_Listad.Text = moddat_g_str_Situac
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.CellFontBold = True
   grd_Listad.Text = "Cliente"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontBold = True
   grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_TDOCLI) & " - " & Trim(g_rst_Princi!HIPMAE_NDOCLI) & " / " & moddat_g_str_NomCli
   
   If g_rst_Princi!HIPMAE_TDOCYG > 0 Then
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Cónyuge"
      
      grd_Listad.Col = 1
      grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_TDOCYG) & " - " & Trim(g_rst_Princi!HIPMAE_NDOCYG) & " / " & moddat_g_str_CygNom
   End If
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Producto"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_NomPrd & " / " & moddat_gf_Consulta_SubPrd(g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB)
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Moneda Préstamo"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_Moneda
   
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Modalidad"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_DesMod
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Dirección Inmueble"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_Direcc
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Distrito"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_Distri
   
   If g_rst_Princi!HIPMAE_PRYMCS = 1 Or (g_rst_Princi!HIPMAE_PRYMCS = 2 And CInt(g_rst_Princi!HIPMAE_CODMOD) = 2 Or CInt(g_rst_Princi!HIPMAE_CODMOD) = 3) Then
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Proyecto Inmobiliario"
      
      grd_Listad.Col = 1
      grd_Listad.Text = moddat_gf_Consulta_NomPry(g_rst_Princi!HIPMAE_PRYINM & "")
      
      If g_rst_Princi!HIPMAE_PRYMCS = 2 Then
         grd_Listad.Text = grd_Listad.Text & " (" & moddat_gf_Consulta_ParDes("513", r_str_CodBco) & ")"
      End If
   End If
   
   If moddat_g_int_TipMon = 1 Then
      grd_Listad.Rows = grd_Listad.Rows + 2
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Valor Compra Venta"
      
      grd_Listad.Col = 1
      grd_Listad.CellFontName = "Lucida Console"
      grd_Listad.CellFontSize = 8
      grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_CVTSOL, 12, 2)
      
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Aporte Propio"
      
      grd_Listad.Col = 1
      grd_Listad.CellFontName = "Lucida Console"
      grd_Listad.CellFontSize = 8
      grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APOSOL, 12, 2)
   Else
      grd_Listad.Rows = grd_Listad.Rows + 2
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Valor Compra Venta"
      
      grd_Listad.Col = 1
      grd_Listad.CellFontName = "Lucida Console"
      grd_Listad.CellFontSize = 8
      grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_CVTDOL, 12, 2)
   
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Aporte Propio"
      
      grd_Listad.Col = 1
      grd_Listad.CellFontName = "Lucida Console"
      grd_Listad.CellFontSize = 8
      grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APODOL, 12, 2)
   End If
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Monto Desembolsado"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPDES, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Monto Préstamo"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_MTOPRE, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Interés Capitalizado"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_INTCAP, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Total Préstamo"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_TOTPRE, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Fecha Activación"
   
   grd_Listad.Col = 1
   grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECACT))
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Fecha Desembolso"
   
   grd_Listad.Col = 1
   grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))
   
   If g_rst_Princi!HIPMAE_FECESC > 0 Then
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Fecha Firma EE.PP"
      
      grd_Listad.Col = 1
      grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECESC))
   End If
   
   If moddat_g_str_CodPrd <> "002" Then
      grd_Listad.Rows = grd_Listad.Rows + 2
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      
      Select Case moddat_g_str_CodPrd > 0
         Case InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd):  grd_Listad.Text = "Nro. Operación Mivivienda"  '"001"
         Case InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd):  grd_Listad.Text = "Nro. Operación COFIDE"      '"003"
         Case InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd): grd_Listad.Text = "Nro. Operación COFIDE"      '"004", "007", "009", "010", "013", "014", "015", "016", "017", "018", "019", "020", "021", "022", "023"
      End Select
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!HIPMAE_OPEMVI & "")
      
      If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0
         grd_Listad.Text = "Nro. Operación Mivivienda"
         
         grd_Listad.Col = 1
         grd_Listad.Text = Trim(g_rst_Princi!HIPMAE_OPEMV1 & "")
      End If
      
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Monto Préstamo (Tramo No Conces.)"
      
      grd_Listad.Col = 1
      grd_Listad.CellFontName = "Lucida Console"
      grd_Listad.CellFontSize = 8
      grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPNCO, 12, 2)
   
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Monto Préstamo (Tramo Conces.)"
      
      grd_Listad.Col = 1
      grd_Listad.CellFontName = "Lucida Console"
      grd_Listad.CellFontSize = 8
      grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPCON, 12, 2)
      
      If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "001" Or moddat_g_str_CodPrd = "003" Then
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0
         grd_Listad.Text = "Tasa de Interés Mivivienda"
      
         grd_Listad.Col = 1
         grd_Listad.Text = Format(g_rst_Princi!HIPMAE_TASMVI, "##0.00") & " %"
      End If
      
      If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "003" Or moddat_g_str_CodPrd = "007" Or moddat_g_str_CodPrd = "009" Or moddat_g_str_CodPrd = "010" Or moddat_g_str_CodPrd = "013" Or moddat_g_str_CodPrd = "014" Or moddat_g_str_CodPrd = "015" Or moddat_g_str_CodPrd = "016" Or moddat_g_str_CodPrd = "017" Or moddat_g_str_CodPrd = "018" Or moddat_g_str_CodPrd = "019" Or moddat_g_str_CodPrd = "020" Or moddat_g_str_CodPrd = "021" Or moddat_g_str_CodPrd = "022" Or moddat_g_str_CodPrd = "023" Then
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0
         grd_Listad.Text = "Tasa de Interés COFIDE"
      
         grd_Listad.Col = 1
         grd_Listad.Text = Format(g_rst_Princi!HIPMAE_TASCOF, "##0.00") & " %"
         
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0
         grd_Listad.Text = "Tasa de Comisión COFIDE"
         
         grd_Listad.Col = 1
         grd_Listad.Text = Format(g_rst_Princi!HIPMAE_COMCOF, "##0.00") & " %"
      End If
   End If
   
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Plazo"
   
   grd_Listad.Col = 1
   grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_PLAANO) & " Años"
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Tasa de Interés"
   
   grd_Listad.Col = 1
   grd_Listad.Text = Format(g_rst_Princi!HIPMAE_TASINT, "##0.00") & " %"
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Nro. de Cuotas"
   
   grd_Listad.Col = 1
   grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_NUMCUO)
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Período de Gracia"
   
   grd_Listad.Col = 1
   grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_PERGRA) & " Meses"
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Compañía de Seguros"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!HIPMAE_SEGPRE & "")
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Tipo de Seguro Desg."
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!HIPMAE_SEGPRE, g_rst_Princi!HIPMAE_TIPSEG)
      
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Tipo Garantía"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_gf_Consulta_ParDes("241", CStr(g_rst_Princi!HIPMAE_TIPGAR))
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Monto Garantía"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!HIPMAE_MONGAR)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_MTOGAR, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Saldo Capital"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCAP + g_rst_Princi!HIPMAE_SALCON, 12, 2)
      
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Cuotas Pendientes de Pago"
   
   grd_Listad.Col = 1
   grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_CUOPEN)
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Días de Atraso"
   
   grd_Listad.Col = 1
   grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_DIAMOR) & " Días"
   
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Saldo Capital (Tramo No Conces.)"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCAP, 12, 2)
      
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Saldo Capital (Tramo Conces.)"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCON, 12, 2)
 
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Consejero Hipotecario"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_NomConHip
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Ejecutivo de Seguimiento"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_NomEjeSeg
   
   'DATOS DE LA OPERACION DEL PREPAGO
   l_dbl_TasInt = g_rst_Princi!HIPMAE_TASINT
   l_dbl_SegDes = g_rst_Princi!HIPMAE_FOIPRE
   l_dbl_SegInm = g_rst_Princi!HIPMAE_FOIVIV
   l_int_CodPrd = g_rst_Princi!HIPMAE_CODPRD
   l_int_NumCuo = g_rst_Princi!HIPMAE_NUMCUO
   l_int_PagCuo = g_rst_Princi!HIPMAE_CUOPAG
   l_int_PerGra = g_rst_Princi!HIPMAE_PERGRA
   l_str_UltVct = g_rst_Princi!HIPMAE_UlTVCT
   l_dbl_CuoFij = CDbl(g_rst_Princi!HIPMAE_CUOFIJ)
   l_dbl_PagCap = CDbl(g_rst_Princi!HIPMAE_PAGCAP)
   l_dbl_PagInt = CDbl(g_rst_Princi!HIPMAE_PAGINT)
   l_int_NCuota = g_rst_Princi!HIPMAE_NUMCUO
   
   pnl_CuoPen.Caption = CInt(g_rst_Princi!HIPMAE_CUOPEN) & " "
   
   If pnl_DeuPen.Caption > 0# Then
      r_bol_CuoVen = True
      l_dbl_SalNco = fs_Obtiene_Saldos(g_rst_Princi!HIPMAE_NUMOPE, 1, True)
      l_dbl_SalCon = fs_Obtiene_Saldos(g_rst_Princi!HIPMAE_NUMOPE, 2, True)
      pnl_CuoPen.Caption = CInt(g_rst_Princi!HIPMAE_CUOPEN) - CInt(Me.grd_DeuPen.Rows - 1) & " "
      pnl_UltPagTNC.Caption = gf_FormatoFecha(fs_Obtiene_FechaPago(g_rst_Princi!HIPMAE_NUMOPE, 1, g_rst_Princi!HIPMAE_FECDES, True))
      pnl_UltPagTC.Caption = gf_FormatoFecha(fs_Obtiene_FechaPago(g_rst_Princi!HIPMAE_NUMOPE, 2, g_rst_Princi!HIPMAE_FECDES, True))
   Else
      r_bol_CuoVen = False
      l_dbl_SalNco = g_rst_Princi!HIPMAE_SALCAP
      l_dbl_SalCon = g_rst_Princi!HIPMAE_SALCON
      pnl_UltPagTNC.Caption = gf_FormatoFecha(fs_Obtiene_FechaPago(g_rst_Princi!HIPMAE_NUMOPE, 1, g_rst_Princi!HIPMAE_FECDES, False))
      pnl_UltPagTC.Caption = gf_FormatoFecha(fs_Obtiene_FechaPago(g_rst_Princi!HIPMAE_NUMOPE, 2, g_rst_Princi!HIPMAE_FECDES, False))
   End If
   
   pnl_DiasTNC.Caption = DateDiff("d", pnl_UltPagTNC.Caption, ipp_FecPre.Text) & " "
   pnl_DiasTC.Caption = DateDiff("d", pnl_UltPagTC.Caption, ipp_FecPre.Text) & " "
   pnl_SaldoTNC1.Caption = Format(l_dbl_SalNco, "###,###.00") & " "
   pnl_SaldoTC1.Caption = Format(l_dbl_SalCon, "###,###.00") & " "
   pnl_SaldoTNC2.Caption = Format(l_dbl_SalNco, "###,###.00") & " "
   pnl_SaldoTC2.Caption = Format(l_dbl_SalCon, "###,###.00") & " "
   
   If pnl_DeuPen.Caption > 0# Then
      Call fs_Obtiene_PBPPerdido(g_rst_Princi!HIPMAE_NUMOPE, r_dbl_CapPBP, r_dbl_IntPBP, True)
   Else
      Call fs_Obtiene_PBPPerdido(g_rst_Princi!HIPMAE_NUMOPE, r_dbl_CapPBP, r_dbl_IntPBP, False)
   End If
   
   pnl_CapPbp.Caption = Format(r_dbl_CapPBP, "###,###.00") & " "
   pnl_IntPbp.Caption = Format(r_dbl_IntPBP, "###,###.00") & " "
   l_dbl_PorNco = Format(g_rst_Princi!HIPMAE_IMPNCO / g_rst_Princi!HIPMAE_TOTPRE, "##0.0000")
   l_dbl_PorCon = 1 - l_dbl_PorNco
   l_str_PrxVct = g_rst_Princi!HIPMAE_PRXVCT
   
   'DETERMINA SI OPERACION ES MICASITA O MIVIVIENDA (UN SOLO TRAMO)
   If InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then
      l_dbl_PorITF = opecaj_gf_Consulta_ITF(Format(CDate(moddat_g_str_FecSis), "yyyymmdd"), 1)
      l_dbl_PorNco = 1
      l_dbl_PorCon = 0
      Label21.Caption = "Saldo Actual"
      Label23.Caption = "Ultimo Pago"
      Label25.Caption = "Dias"
      Label6.Caption = "Interés a la fecha"
      Label20.Visible = False
      pnl_SaldoTC1.Visible = False
      Label22.Visible = False
      pnl_UltPagTC.Visible = False
      Label24.Visible = False
      pnl_DiasTC.Visible = False
      Label7.Visible = False
      txt_InteresTC.Visible = False
      Label28.Caption = "Saldo Actual"
      Label8.Caption = "Aplicación PP"
      Label10.Caption = "Nuevo Saldo"
      Label33.Visible = False
      pnl_IntPbp.Visible = False
      Label32.Visible = False
      pnl_CapPbp.Visible = False
      Label27.Visible = False
      pnl_SaldoTC2.Visible = False
      Label13.Visible = False
      txt_ApliTC.Visible = False
      Label16.Visible = False
      pnl_NuevoSaldoTC.Visible = False
      tab_Cronog.TabVisible(0) = True
      tab_Cronog.TabVisible(1) = False
      tab_Cronog.TabCaption(0) = "Cliente"
      tab_Cronog.TabCaption(1) = ""
   Else
      Label21.Caption = "Saldo Actual TNC"
      Label23.Caption = "Ultimo Pago TNC"
      Label25.Caption = "Dias TNC"
      Label6.Caption = "Interés TNC a la fecha"
      Label20.Visible = True
      pnl_SaldoTC1.Visible = True
      Label22.Visible = True
      pnl_UltPagTC.Visible = True
      Label24.Visible = True
      pnl_DiasTC.Visible = True
      Label7.Visible = True
      txt_InteresTC.Visible = True
      Label28.Caption = "Saldo Actual TNC"
      Label8.Caption = "Aplica. PP TNC"
      Label10.Caption = "Nuevo TNC"
      Label27.Visible = True
      pnl_SaldoTC2.Visible = True
      Label13.Visible = True
      txt_ApliTC.Visible = True
      Label16.Visible = True
      Label33.Visible = True
      pnl_IntPbp.Visible = True
      Label32.Visible = True
      pnl_CapPbp.Visible = True
      pnl_NuevoSaldoTC.Visible = True
      tab_Cronog.TabVisible(0) = True
      tab_Cronog.TabVisible(1) = True
      tab_Cronog.TabCaption(0) = "Cliente - No Concesional"
      tab_Cronog.TabCaption(1) = "Cliente - Concesional"
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Call gs_UbiIniGrid(grd_Listad)
End Sub

Private Function fs_Obtiene_FechaPago(ByVal p_NumOpe As String, ByVal p_TipCro As Integer, ByVal p_FecDes As String, ByVal p_FlgCuoVen As Boolean) As String
Dim r_rst_Temp    As Recordset
   fs_Obtiene_FechaPago = p_FecDes
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT HIPCUO_FECVCT "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & p_NumOpe & "' "
   If p_FlgCuoVen = True Then
      g_str_Parame = g_str_Parame & "   AND HIPCUO_FECVCT <= " & Format(ipp_FecPre.Text, "yyyymmdd") & " "  'moddat_g_str_FecSis
   Else
      g_str_Parame = g_str_Parame & "   AND HIPCUO_SITUAC = 1 "
   End If
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = " & p_TipCro & " "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_FECVCT DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Temp, 3) Then
       Exit Function
   End If
   
   If Not (r_rst_Temp.BOF And r_rst_Temp.EOF) Then
      r_rst_Temp.MoveFirst
      fs_Obtiene_FechaPago = r_rst_Temp!HIPCUO_FECVCT
   End If
   
   r_rst_Temp.Close
   Set r_rst_Temp = Nothing
End Function

Private Function fs_Obtiene_Saldos(ByVal p_NumOpe As String, ByVal p_TipCro As Integer, ByVal p_FlgCuoVen As Boolean) As Double
Dim r_rst_Temp    As Recordset
   fs_Obtiene_Saldos = 0#
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT HIPCUO_SALCAP "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & p_NumOpe & "' "
   If p_FlgCuoVen = True Then
      g_str_Parame = g_str_Parame & "   AND HIPCUO_FECVCT <= " & Format(ipp_FecPre.Text, "yyyymmdd") & " "
   Else
      g_str_Parame = g_str_Parame & "   AND HIPCUO_SITUAC = 1 "
   End If
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = " & p_TipCro & " "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_FECVCT DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Temp, 3) Then
       Exit Function
   End If
   
   If Not (r_rst_Temp.BOF And r_rst_Temp.EOF) Then
      r_rst_Temp.MoveFirst
      fs_Obtiene_Saldos = r_rst_Temp!HIPCUO_SALCAP
   End If
   
   r_rst_Temp.Close
   Set r_rst_Temp = Nothing
End Function

Private Function fs_Obtiene_ProxVenc(ByVal p_NumOpe As String, ByVal p_TipCro As Integer) As String
Dim r_rst_Temp    As Recordset
Dim r_str_FecVct  As String
   
   'Ubica cuotas pagadas parcialmente
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT HIPCUO_NUMCUO, HIPCUO_FECVCT   "
   g_str_Parame = g_str_Parame & "   FROM CRE_HIPCUO  "
   g_str_Parame = g_str_Parame & "  WHERE HIPCUO_NUMOPE = '" & p_NumOpe & "' "
   g_str_Parame = g_str_Parame & "    AND HIPCUO_IMPPAG > 0 "
   g_str_Parame = g_str_Parame & "    AND HIPCUO_IMPPAG > HIPCUO_CAPITA "
   g_str_Parame = g_str_Parame & "    AND HIPCUO_SITUAC = 2 "
   g_str_Parame = g_str_Parame & "    AND HIPCUO_TIPCRO = 1 "
   g_str_Parame = g_str_Parame & "  ORDER BY HIPCUO_FECVCT ASC"

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Temp, 3) Then
       Exit Function
   End If
   
   If Not (r_rst_Temp.BOF And r_rst_Temp.EOF) Then
      r_rst_Temp.MoveFirst
      r_str_FecVct = r_rst_Temp!HIPCUO_FECVCT
   End If
   
   r_rst_Temp.Close
   Set r_rst_Temp = Nothing
   
   'Vencidos y por vencer
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT HIPCUO_NUMCUO, HIPCUO_FECVCT "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & p_NumOpe & "' "
   If r_str_FecVct = "" Then
      g_str_Parame = g_str_Parame & "   AND HIPCUO_FECVCT > " & Format(ipp_FecPre.Text, "yyyymmdd") & " "
   Else
      g_str_Parame = g_str_Parame & "   AND HIPCUO_FECVCT > " & r_str_FecVct & " "
   End If
   g_str_Parame = g_str_Parame & "   AND HIPCUO_SITUAC = 2 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = " & p_TipCro & " "
   g_str_Parame = g_str_Parame & " ORDER BY HIPCUO_FECVCT ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Temp, 3) Then
      Exit Function
   End If

   If Not (r_rst_Temp.BOF And r_rst_Temp.EOF) Then
      r_rst_Temp.MoveFirst
      fs_Obtiene_ProxVenc = r_rst_Temp!HIPCUO_FECVCT
   End If
   
   r_rst_Temp.Close
   Set r_rst_Temp = Nothing
End Function

Private Function fs_Obtiene_ProxVenc_old(ByVal p_NumOpe As String, ByVal p_TipCro As Integer) As String
Dim r_rst_Temp    As Recordset

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT HIPCUO_NUMCUO, HIPCUO_FECVCT "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & p_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_FECVCT >= " & Format(ipp_FecPre.Text, "yyyymmdd") & " "  'moddat_g_str_FecSis
   g_str_Parame = g_str_Parame & "   AND HIPCUO_SITUAC = 2 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = " & p_TipCro & " "
   g_str_Parame = g_str_Parame & " ORDER BY HIPCUO_FECVCT ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Temp, 3) Then
       Exit Function
   End If
   
   If Not (r_rst_Temp.BOF And r_rst_Temp.EOF) Then
      r_rst_Temp.MoveFirst
      fs_Obtiene_ProxVenc_old = r_rst_Temp!HIPCUO_FECVCT
   End If
   
   r_rst_Temp.Close
   Set r_rst_Temp = Nothing
End Function

Private Sub fs_Buscar_Tasacion()
Dim r_rst_Temp    As Recordset
   
   moddat_g_str_DesObs = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT EVATAS_TIPMON, EVATAS_SUMASE_INM, EVATAS_SUMASE_ES1, EVATAS_SUMASE_ES2, EVATAS_SUMASE_DEP, "
   g_str_Parame = g_str_Parame & "       EVATAS_VALCOM_INM, EVATAS_VALCOM_ES1, EVATAS_VALCOM_ES2, EVATAS_VALCOM_DEP "
   g_str_Parame = g_str_Parame & "  FROM TRA_EVATAS "
   g_str_Parame = g_str_Parame & " WHERE EVATAS_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Temp, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_Temp.BOF And r_rst_Temp.EOF) Then
      r_rst_Temp.MoveFirst
      pnl_Val_AsgInm.Caption = gf_FormatoNumero(r_rst_Temp!EVATAS_SUMASE_INM + r_rst_Temp!EVATAS_SUMASE_ES1 + r_rst_Temp!EVATAS_SUMASE_ES2 + r_rst_Temp!EVATAS_SUMASE_DEP, 12, 2) & " "
      moddat_g_str_DesObs = CStr(r_rst_Temp!EVATAS_VALCOM_INM + r_rst_Temp!EVATAS_VALCOM_ES1 + r_rst_Temp!EVATAS_VALCOM_ES2 + r_rst_Temp!EVATAS_VALCOM_DEP)
   End If
   
   r_rst_Temp.Close
   Set r_rst_Temp = Nothing
End Sub

Private Sub fs_Obtiene_PBPPerdido(ByVal p_NumOpe As String, ByRef p_CapPBP As Double, ByRef p_IntPBP As Double, ByVal p_FlgCuoVen As Boolean)
Dim r_rst_Temp    As Recordset

   p_CapPBP = 0
   p_IntPBP = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT SUM(HIPCUO_CAPBBP) AS CAP_PBP_PDTE, SUM(HIPCUO_INTBBP) AS INT_PBP_PDTE "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & p_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 1 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_SITUAC = 2 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_CAPBBP > 0 "
   If p_FlgCuoVen = True Then
      g_str_Parame = g_str_Parame & "   AND HIPCUO_FECVCT > " & Format(ipp_FecPre.Text, "yyyymmdd") & " "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Temp, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_Temp.BOF And r_rst_Temp.EOF) Then
      r_rst_Temp.MoveFirst
      If Not IsNull(r_rst_Temp!CAP_PBP_PDTE) Then
         p_CapPBP = r_rst_Temp!CAP_PBP_PDTE
      End If
      If Not IsNull(r_rst_Temp!INT_PBP_PDTE) Then
         p_IntPBP = r_rst_Temp!INT_PBP_PDTE
      End If
   End If
   
   r_rst_Temp.Close
   Set r_rst_Temp = Nothing
End Sub

Private Sub fs_Buscar_Cuotas_Vencidas()
Dim r_dbl_ValCuo     As Double
Dim r_str_Parame     As String
Dim r_rst_Princi     As ADODB.Recordset
Dim r_dbl_DeuVen     As Double

   'Cuotas Vencidas
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT HIPCUO_NUMCUO, HIPCUO_FECVCT, HIPMAE_MONEDA, HIPCUO_CAPITA, HIPCUO_CAPPAG, HIPCUO_INTERE, HIPCUO_INTPAG, "
   r_str_Parame = r_str_Parame & "        HIPCUO_DESORG, HIPCUO_DESPAG, HIPCUO_VIVORG, HIPCUO_VIVPAG, HIPCUO_OTRORG, HIPCUO_OTRPAG, HIPCUO_CAPBBP, "
   r_str_Parame = r_str_Parame & "        HIPCUO_CBPPAG, HIPCUO_INTBBP, HIPCUO_IBPPAG, HIPCUO_INTMOR, HIPCUO_IMOPAG, HIPCUO_INTCOM, HIPCUO_ICOPAG,  "
   r_str_Parame = r_str_Parame & "        HIPCUO_GASCOB, HIPCUO_GCOPAG, HIPCUO_OTRGAS, HIPCUO_OTGPAG "
   r_str_Parame = r_str_Parame & "   FROM CRE_HIPCUO INNER JOIN CRE_HIPMAE ON HIPMAE_NUMOPE = HIPCUO_NUMOPE "
   r_str_Parame = r_str_Parame & "  WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   r_str_Parame = r_str_Parame & "    AND HIPCUO_TIPCRO = 1 "
   r_str_Parame = r_str_Parame & "    AND HIPCUO_SITUAC = 2 "
   r_str_Parame = r_str_Parame & "    AND HIPCUO_IMPPAG = 0 "
   r_str_Parame = r_str_Parame & "    AND HIPCUO_FECVCT <= " & Format(ipp_FecPre.Text, "yyyymmdd") & " "
   r_str_Parame = r_str_Parame & "  UNION ALL "
   r_str_Parame = r_str_Parame & " SELECT HIPCUO_NUMCUO, HIPCUO_FECVCT, HIPMAE_MONEDA, HIPCUO_CAPITA, HIPCUO_CAPPAG, HIPCUO_INTERE, HIPCUO_INTPAG,"
   r_str_Parame = r_str_Parame & "        HIPCUO_DESORG, HIPCUO_DESPAG, HIPCUO_VIVORG, HIPCUO_VIVPAG, HIPCUO_OTRORG, HIPCUO_OTRPAG, HIPCUO_CAPBBP,"
   r_str_Parame = r_str_Parame & "        HIPCUO_CBPPAG, HIPCUO_INTBBP, HIPCUO_IBPPAG, HIPCUO_INTMOR, HIPCUO_IMOPAG, HIPCUO_INTCOM, HIPCUO_ICOPAG,"
   r_str_Parame = r_str_Parame & "        HIPCUO_GASCOB , HIPCUO_GCOPAG, HIPCUO_OTRGAS, HIPCUO_OTGPAG"
   r_str_Parame = r_str_Parame & "   FROM CRE_HIPCUO "
   r_str_Parame = r_str_Parame & "  INNER JOIN CRE_HIPMAE ON HIPMAE_NUMOPE = HIPCUO_NUMOPE"
   r_str_Parame = r_str_Parame & "  WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   r_str_Parame = r_str_Parame & "    AND HIPCUO_TIPCRO = 1  "
   r_str_Parame = r_str_Parame & "    AND HIPCUO_SITUAC = 2  "
   r_str_Parame = r_str_Parame & "    AND HIPCUO_IMPPAG > 0 "
   r_str_Parame = r_str_Parame & "    AND HIPCUO_IMPPAG > HIPCUO_CAPITA "
   r_str_Parame = r_str_Parame & "  ORDER BY HIPCUO_NUMCUO ASC "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Exit Sub
   End If
   
   Call gs_LimpiaGrid(grd_DeuPen)
   pnl_DeuPen.Caption = "0.00 "
   
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      grd_DeuPen.Redraw = False
      grd_DeuPen.Cols = 16
      
      'Cabecera de la Grilla
      grd_DeuPen.Rows = grd_DeuPen.Rows + 2
      grd_DeuPen.FixedRows = 1
      grd_DeuPen.Rows = grd_DeuPen.Rows - 1
      grd_DeuPen.Row = 0
          
      grd_DeuPen.Col = 0:    grd_DeuPen.Text = "Cuota":        grd_DeuPen.CellAlignment = flexAlignCenterCenter
      grd_DeuPen.Col = 1:    grd_DeuPen.Text = "Estado":       grd_DeuPen.CellAlignment = flexAlignCenterCenter
      grd_DeuPen.Col = 2:    grd_DeuPen.Text = "F. Vcto.":     grd_DeuPen.CellAlignment = flexAlignCenterCenter
      'grd_DeuPen.Col = 3:    grd_DeuPen.Text = "Moneda":       grd_DeuPen.CellAlignment = flexAlignCenterCenter
      grd_DeuPen.Col = 4:    grd_DeuPen.Text = "Capital":      grd_DeuPen.CellAlignment = flexAlignCenterCenter
      grd_DeuPen.Col = 5:    grd_DeuPen.Text = "Interés":      grd_DeuPen.CellAlignment = flexAlignCenterCenter
      grd_DeuPen.Col = 6:    grd_DeuPen.Text = "Seg. Desg.":   grd_DeuPen.CellAlignment = flexAlignCenterCenter
      grd_DeuPen.Col = 7:    grd_DeuPen.Text = "Seg. Viv.":    grd_DeuPen.CellAlignment = flexAlignCenterCenter
      grd_DeuPen.Col = 8:    grd_DeuPen.Text = "Portes":       grd_DeuPen.CellAlignment = flexAlignCenterCenter
      grd_DeuPen.Col = 9:    grd_DeuPen.Text = "Capital BBP":  grd_DeuPen.CellAlignment = flexAlignCenterCenter
      grd_DeuPen.Col = 10:   grd_DeuPen.Text = "Interés BBP":  grd_DeuPen.CellAlignment = flexAlignCenterCenter
      'grd_DeuPen.Col = 11:   grd_DeuPen.Text = "Int. Morat.":  grd_DeuPen.CellAlignment = flexAlignCenterCenter
      'grd_DeuPen.Col = 12:   grd_DeuPen.Text = "Int. Comp.":   grd_DeuPen.CellAlignment = flexAlignCenterCenter
      grd_DeuPen.Col = 13:   grd_DeuPen.Text = "G. Cobr.":     grd_DeuPen.CellAlignment = flexAlignCenterCenter
      'grd_DeuPen.Col = 14:   grd_DeuPen.Text = "Otr. Gastos":  grd_DeuPen.CellAlignment = flexAlignCenterCenter
      grd_DeuPen.Col = 15:   grd_DeuPen.Text = "Total Cuota":  grd_DeuPen.CellAlignment = flexAlignCenterCenter
      
      r_rst_Princi.MoveFirst
      
      Do While Not r_rst_Princi.EOF
         grd_DeuPen.Rows = grd_DeuPen.Rows + 1
         grd_DeuPen.Row = grd_DeuPen.Rows - 1
         r_dbl_ValCuo = 0
         
         grd_DeuPen.Col = 0
         grd_DeuPen.Text = Format(r_rst_Princi!HIPCUO_NUMCUO, "000")
         
         grd_DeuPen.Col = 1
         grd_DeuPen.Text = IIf(CLng(r_rst_Princi!HIPCUO_FECVCT) < CLng(Format(date, "yyyymmdd")), "V", "PV")
      
         grd_DeuPen.Col = 2
         grd_DeuPen.Text = gf_FormatoFecha(CStr(r_rst_Princi!HIPCUO_FECVCT))
      
         grd_DeuPen.Col = 3
         'grd_DeuPen.Text = moddat_gf_Consulta_ParDes("229", CStr(r_rst_Princi!HIPMAE_MONEDA))

         'Capital
         grd_DeuPen.Col = 4
         grd_DeuPen.Text = Format(r_rst_Princi!HIPCUO_CAPITA - r_rst_Princi!HIPCUO_CAPPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_DeuPen.Text)
      
         'Interes
         grd_DeuPen.Col = 5
         grd_DeuPen.Text = Format(r_rst_Princi!HIPCUO_INTERE - r_rst_Princi!HIPCUO_INTPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_DeuPen.Text)
         
         'Seguro de Desgravamen
         grd_DeuPen.Col = 6
         grd_DeuPen.Text = Format(r_rst_Princi!HIPCUO_DESORG - r_rst_Princi!HIPCUO_DESPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_DeuPen.Text)
         
         'Seguro de Vivienda
         grd_DeuPen.Col = 7
         grd_DeuPen.Text = Format(r_rst_Princi!HIPCUO_VIVORG - r_rst_Princi!HIPCUO_VIVPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_DeuPen.Text)
         
         'Otros Cargos
         grd_DeuPen.Col = 8
         grd_DeuPen.Text = Format(r_rst_Princi!HIPCUO_OTRORG - r_rst_Princi!HIPCUO_OTRPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_DeuPen.Text)
         
         'Capital PBP
         grd_DeuPen.Col = 9
         grd_DeuPen.Text = Format(r_rst_Princi!HIPCUO_CAPBBP - r_rst_Princi!HIPCUO_CBPPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_DeuPen.Text)
         
         'Interés PBP
         grd_DeuPen.Col = 10
         grd_DeuPen.Text = Format(r_rst_Princi!HIPCUO_INTBBP - r_rst_Princi!HIPCUO_IBPPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_DeuPen.Text)
         
         'Interes Moratorio
         grd_DeuPen.Col = 11
         'grd_DeuPen.Text = Format(r_rst_Princi!HIPCUO_INTMOR - r_rst_Princi!HIPCUO_IMOPAG, "###,###,##0.00")
         'r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_DeuPen.Text)
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(r_rst_Princi!HIPCUO_INTMOR - r_rst_Princi!HIPCUO_IMOPAG)
         
         'Interes Compensatorio
         grd_DeuPen.Col = 12
         'grd_DeuPen.Text = Format(r_rst_Princi!HIPCUO_INTCOM - r_rst_Princi!HIPCUO_ICOPAG, "###,###,##0.00")
         'r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_DeuPen.Text)
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(r_rst_Princi!HIPCUO_INTCOM - r_rst_Princi!HIPCUO_ICOPAG)
         
         'Gastos de Cobranza
         grd_DeuPen.Col = 13
         grd_DeuPen.Text = Format(r_rst_Princi!HIPCUO_GASCOB - r_rst_Princi!HIPCUO_GCOPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_DeuPen.Text)
         
         'Otros Gastos
         grd_DeuPen.Col = 14
         'grd_DeuPen.Text = Format(r_rst_Princi!HIPCUO_OTRGAS - r_rst_Princi!HIPCUO_OTGPAG, "###,###,##0.00")
         'r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_DeuPen.Text)
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(r_rst_Princi!HIPCUO_OTRGAS - r_rst_Princi!HIPCUO_OTGPAG)
         
         'Valor Cuota
         grd_DeuPen.Col = 15
         grd_DeuPen.Text = Format(r_dbl_ValCuo, "###,###,##0.00")
         
         r_dbl_DeuVen = r_dbl_DeuVen + r_dbl_ValCuo
         r_rst_Princi.MoveNext
      Loop
      grd_DeuPen.Redraw = True
      
      pnl_DeuPen.Caption = Format(r_dbl_DeuVen, "###,###.00") & " "
      Call gs_UbiIniGrid(grd_DeuPen)
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
End Sub

Private Sub fs_Envia_Correo()
Dim r_str_Mensaj     As String
Dim r_str_Asunto     As String

   ReDim moddat_g_arr_Genera(0)

   r_str_Asunto = "PREPAGO PARCIAL DE CREDITO HIPOTECARIO (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
   r_str_Mensaj = ""
   r_str_Mensaj = r_str_Mensaj & "NUMERO DE OPERACION : " & moddat_g_str_NumOpe & Chr(13)
   r_str_Mensaj = r_str_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
   r_str_Mensaj = r_str_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
   r_str_Mensaj = r_str_Mensaj & Chr(13)

   'Evaluador de Operaciones
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(221, moddat_g_arr_Genera)
   
   'Plataforma
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(122, moddat_g_arr_Genera)

   If InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then
      Call fs_PpgPar_Micasita(True, l_str_NomAch)
   Else
      Call fs_PpgPar_Mivivienda(True, l_str_NomAch)
   End If
      
   Call moddat_gs_EnvCor(mps_Sesion, mps_Mensaj, moddat_g_arr_Genera, r_str_Asunto, r_str_Mensaj, l_str_NomAch, g_str_RutLog & "\")
End Sub

Private Sub fs_Envia_Correo_Plaft()
Dim r_str_Mensaj     As String
Dim r_str_Asunto     As String

   If Mid(moddat_g_str_NumOpe, 1, 3) = "001" Or Mid(moddat_g_str_NumOpe, 1, 3) = "002" Then
      If CDbl(txt_Mto_Deposito.Value) < 10000 Then
         Exit Sub
      End If
   Else
      If CDbl(txt_Mto_Deposito.Value) < 30000 Then
         Exit Sub
      End If
   End If
   
   ReDim moddat_g_arr_Genera(0)

   r_str_Asunto = "ALERTA DE PREPAGO (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
   r_str_Mensaj = ""
   r_str_Mensaj = r_str_Mensaj & "NUMERO DE OPERACION : " & moddat_g_str_NumOpe & Chr(13)
   r_str_Mensaj = r_str_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
   r_str_Mensaj = r_str_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
   r_str_Mensaj = r_str_Mensaj & Chr(13)
   r_str_Mensaj = r_str_Mensaj & "El sistema alerto el prepago detallado en el adjunto"
  
   'Jefe de Legal
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(230, moddat_g_arr_Genera)
   
   If InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then
      Call fs_PpgPar_Micasita(True, l_str_NomAch)
   Else
      Call fs_PpgPar_Mivivienda(True, l_str_NomAch)
   End If
      
   Call moddat_gs_EnvCor(mps_Sesion, mps_Mensaj, moddat_g_arr_Genera, r_str_Asunto, r_str_Mensaj, l_str_NomAch, g_str_RutLog & "\")
End Sub

Private Sub fs_Envia_Correo_Preventivo()
Dim r_str_Mensaj     As String
Dim r_str_Asunto     As String

   ReDim moddat_g_arr_Genera(0)

   r_str_Asunto = "CLIENTE EN PROCESO DUE DILIGENCE"
   r_str_Mensaj = ""
   r_str_Mensaj = r_str_Mensaj & "NUMERO DE OPERACION : " & moddat_g_str_NumOpe & Chr(13)
   r_str_Mensaj = r_str_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
   r_str_Mensaj = r_str_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
   r_str_Mensaj = r_str_Mensaj & Chr(13)
   r_str_Mensaj = r_str_Mensaj & "El sistema alerto el prepago registrado"
     
   'Evaluador de Operaciones
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(221, moddat_g_arr_Genera)
  
   'Call moddat_gs_EnvCor(mps_Sesion, mps_Mensaj, moddat_g_arr_Genera, r_str_Asunto, r_str_Mensaj, l_str_NomAch, g_str_RutLog & "\")
   Call moddat_gs_EnvCor(mps_Sesion, mps_Mensaj, moddat_g_arr_Genera, r_str_Asunto, r_str_Mensaj, "", "")
End Sub


Private Sub fs_Cal_Prpago()
  pnl_MtoApl.Caption = gf_FormatoNumero(CDbl(txt_Mto_Deposito.Text) - (CDbl(txt_InteresTNC.Text) + CDbl(txt_InteresTC.Text) + CDbl(txt_SegDes.Text) + CDbl(txt_SegInm.Text) + CDbl(txt_MontoITF.Text) + CDbl(Trim(pnl_DeuPen.Caption)) + CDbl(Trim(pnl_IntPbp.Caption))), 12, 2) & " "
  pnl_MtoApl_Fin.Caption = Format(CDbl(pnl_MtoApl.Caption) - CDbl(Trim(pnl_CapPbp.Caption)), "###,###.00") & " "
End Sub

Private Sub fs_Cal_Prctaj()
   If Mid(moddat_g_str_NumOpe, 1, 3) = "001" Or Mid(moddat_g_str_NumOpe, 1, 3) = "003" Then
      txt_AplTNC.Text = l_dbl_SalNco - ((CDbl(pnl_SaldoTNC1.Caption) + CDbl(pnl_SaldoTC1.Caption) - CDbl(pnl_MtoApl_Fin.Caption)) * 0.85)
      txt_ApliTC.Text = CDbl(pnl_MtoApl_Fin.Caption) - CDbl(txt_AplTNC.Text)
      If txt_AplTNC.Text < 0 Then txt_AplTNC.Text = 0
      If txt_ApliTC.Text < 0 Then txt_ApliTC.Text = 0
   Else
      txt_AplTNC.Text = CDbl(pnl_MtoApl_Fin) * l_dbl_PorNco
      txt_ApliTC.Text = CDbl(pnl_MtoApl_Fin) - CDbl(txt_AplTNC.Text)
      If txt_AplTNC.Text < 0 Then txt_AplTNC.Text = 0
      If txt_ApliTC.Text < 0 Then txt_ApliTC.Text = 0
   End If
   pnl_NuevoSaldoTNC.Caption = gf_FormatoNumero(l_dbl_SalNco - CDbl(txt_AplTNC.Text), 12, 2) & " "
   pnl_NuevoSaldoTC.Caption = gf_FormatoNumero(l_dbl_SalCon - CDbl(txt_ApliTC.Text), 12, 2) & " "

'   l_dbl_PorNco = l_dbl_SalNco / (CDbl(pnl_SaldoTNC1.Caption) + CDbl(pnl_SaldoTC1.Caption))
'
'   If Mid(moddat_g_str_NumOpe, 1, 3) = "001" Or Mid(moddat_g_str_NumOpe, 1, 3) = "003" Then
'      txt_AplTNC.Text = l_dbl_SalNco - (CDbl(pnl_SaldoTNC1.Caption) + CDbl(pnl_SaldoTC1.Caption) - CDbl(pnl_MtoApl_Fin.Caption)) * l_dbl_PorNco
'      txt_ApliTC.Text = CDbl(pnl_MtoApl_Fin.Caption) - CDbl(txt_AplTNC.Text)
'      If txt_AplTNC.Text < 0 Then txt_AplTNC.Text = 0
'      If txt_ApliTC.Text < 0 Then txt_ApliTC.Text = 0
'   Else
'      txt_AplTNC.Text = CDbl(pnl_MtoApl_Fin) * l_dbl_PorNco
'      txt_ApliTC.Text = CDbl(pnl_MtoApl_Fin) - CDbl(txt_AplTNC.Text)
'      If txt_AplTNC.Text < 0 Then txt_AplTNC.Text = 0
'      If txt_ApliTC.Text < 0 Then txt_ApliTC.Text = 0
'   End If
'   pnl_NuevoSaldoTNC.Caption = gf_FormatoNumero(l_dbl_SalNco - CDbl(txt_AplTNC.Text), 12, 2) & " "
'   pnl_NuevoSaldoTC.Caption = gf_FormatoNumero(l_dbl_SalCon - CDbl(txt_ApliTC.Text), 12, 2) & " "
End Sub

Private Sub fs_Cal_MtoItf()
   If InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then
      txt_MontoITF.Text = txt_Mto_Deposito * (l_dbl_PorITF / 100)
   End If
End Sub

Private Sub fs_Cal_Interes()
   If CDbl(pnl_DiasTNC.Caption) > 0 Then
      txt_InteresTNC.Text = Format((CDbl(pnl_SaldoTNC1.Caption)) * (1 + (l_dbl_TasInt / 100)) ^ (CDbl(pnl_DiasTNC.Caption) / 360) - CDbl(pnl_SaldoTNC1.Caption), "###,##0.00")
   Else
      txt_InteresTNC.Text = 0
   End If
   If Not (InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0) Then
      If CDbl(pnl_DiasTC.Caption) > 0 Then
         txt_InteresTC.Text = Format(CDbl(pnl_SaldoTC1.Caption) * (1 + (l_dbl_TasInt / 100)) ^ (CDbl(pnl_DiasTC.Caption) / 360) - CDbl(pnl_SaldoTC1.Caption), "###,##0.00")
      Else
         txt_InteresTC.Text = 0
      End If
   End If
   If CDbl(pnl_DiasTNC.Caption) > 0 And l_int_TipSeg <> 13 Then  '13 = endosado
      txt_SegDes.Text = Format((CDbl(pnl_SaldoTNC1.Caption)) * (1 + (l_dbl_SegDes / 100)) ^ (CDbl(pnl_DiasTNC.Caption) / 30) - CDbl(pnl_SaldoTNC1.Caption), "###,##0.00")
   Else
      txt_SegDes.Text = 0
   End If
End Sub

Private Sub fs_Recalc()
Dim r_int_CuoExt        As Integer
Dim r_dbl_ValViv        As Double
Dim r_dbl_CuoIni        As Double
Dim r_int_TipSeg        As Integer
Dim r_dbl_FoiViv        As Double
Dim r_dbl_FoIDes        As Double
Dim r_dbl_Portes        As Double
Dim r_str_FecDes        As String
Dim r_int_DiaPag        As Integer
Dim r_dbl_MtoAse        As Double
Dim r_int_PriVct        As String
Dim r_int_NumCuo        As Integer

'variables nueva para la generacion del cronograma
Dim obj_Cronog          As Object
Dim int_Produc          As Integer
Dim int_CuoDbl          As Integer
Dim dbl_ValInm          As Double
Dim dbl_CuoIni          As Double
Dim dbl_MtoCon          As Double
Dim dbl_MtoTas          As Double
Dim int_PlaPre          As Integer
Dim dbl_TasInt          As Double
Dim dbl_TasCof          As Double
Dim dbl_ComCof          As Double
Dim dat_FecDes          As Date
Dim int_DiaVct          As Integer
Dim int_PerGra          As Integer
Dim str_PriVct          As String
Dim dbl_Portes          As Double
Dim dbl_SegViv          As Double
Dim int_TipSDe          As Integer
Dim dbl_SegDes          As Double

   Call fs_Inicia_Cronog
   Call fs_Limpia_Cronog
   
   'Obteniendo Datos del Préstamo
   l_int_PlaAno = 0
   l_int_NumCuo = 0
   l_int_PagCuo = 0
   r_int_CuoExt = 0
   r_dbl_FoIDes = 0
   r_dbl_FoiViv = 0
   r_dbl_Portes = 0
   r_str_FecDes = ""
   r_int_DiaPag = 0
   l_int_PerGra = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se encontro información de la Operación.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   l_int_PlaAno = g_rst_Princi!HIPMAE_PLAANO
   l_int_NumCuo = CInt(Trim(pnl_CuoPen.Caption))
   l_int_PagCuo = g_rst_Princi!HIPMAE_CUOPAG
   
   r_int_CuoExt = g_rst_Princi!HIPMAE_CUOANO
   If g_rst_Princi!HIPMAE_MONEDA = 1 Then
      r_dbl_ValViv = g_rst_Princi!HIPMAE_CVTSOL
      r_dbl_CuoIni = g_rst_Princi!HIPMAE_APOSOL
   Else
      r_dbl_ValViv = g_rst_Princi!HIPMAE_CVTDOL
      r_dbl_CuoIni = g_rst_Princi!HIPMAE_APODOL
   End If
   
   r_dbl_FoIDes = g_rst_Princi!HIPMAE_FOIPRE
   r_int_TipSeg = g_rst_Princi!HIPMAE_TIPSEG
   r_dbl_FoiViv = g_rst_Princi!HIPMAE_FOIVIV
   r_dbl_Portes = g_rst_Princi!HIPMAE_OTRIMP
   r_int_DiaPag = g_rst_Princi!HIPMAE_DIAPAG
   
   If Me.pnl_DeuPen.Caption > 0# Then
      r_int_PriVct = gf_FormatoFecha(CStr(fs_Obtiene_ProxVenc(g_rst_Princi!HIPMAE_NUMOPE, 1)))
      l_int_PagCuo = CInt(grd_DeuPen.TextMatrix(grd_DeuPen.Rows - 1, 0))
   Else
      r_int_PriVct = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_PRXVCT))
   End If

   'Prepago reduccion de plazo
   If (cmb_TipPre.ItemData(cmb_TipPre.ListIndex)) = 2 Then
      l_int_NumCuo = CInt(Trim(pnl_CuoPen.Caption)) - (cmb_RedPlz.ItemData(cmb_RedPlz.ListIndex) * 12)
   End If
   
   If l_int_NumCuo < 0 Then
      MsgBox "El plazo esta errado, favor verificar.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   l_dbl_ComCof = g_rst_Princi!HIPMAE_COMCOF
   l_dbl_TasCof = g_rst_Princi!HIPMAE_TASCOF
   
   'Obteniendo Valor de Inmueble
   r_dbl_MtoAse = CDbl(Trim(pnl_Val_AsgInm.Caption))
   
   'fecha de desembolo
   l_str_DesCof = ipp_FecPre.Text
   
   If CDbl(pnl_DiasTNC.Caption) > 0 Then
      r_str_FecDes = ipp_FecPre.Text
      l_str_DesCof = ipp_FecPre.Text
   Else
      r_str_FecDes = pnl_UltPagTNC.Caption
      l_str_DesCof = pnl_UltPagTNC.Caption
   End If
   
   'Generando Cronogramas de Pago
   Select Case moddat_g_str_CodPrd > 0
      Case InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd)
         tab_Cronog.TabCaption(0) = "Cliente - No Concesional"
         tab_Cronog.TabCaption(1) = "Cliente - Concesional"
         
      Case InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd)
         tab_Cronog.TabCaption(0) = "Cliente - No Concesional"
         tab_Cronog.TabCaption(1) = "Cliente - Concesional"
         
      Case InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) Or InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd)
         tab_Cronog.TabCaption(0) = "Cliente"
         tab_Cronog.TabCaption(1) = ""
         
      Case InStr(moddat_g_str_AgrMIHG, moddat_g_str_CodPrd) Or InStr(moddat_g_str_Agr2FMV, moddat_g_str_CodPrd)
         tab_Cronog.TabCaption(0) = "Cliente - No Concesional"
         tab_Cronog.TabCaption(1) = "Cliente - Concesional"
         
      Case InStr(moddat_g_str_Agr2MIC, moddat_g_str_CodPrd)
         tab_Cronog.TabCaption(0) = "Cliente"
         tab_Cronog.TabCaption(1) = "Cliente - Concesional"
   End Select
   
   '********************************************************************************************************
   '*************************** GENERACION DE CRONOGRAMAS SEGUN TIPO DE PRODUCTO ***************************
   '********************************************************************************************************
   Select Case moddat_g_str_CodPrd > 0
      Case InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd)    '"001"
         'NUEVA rutina de generacion de cronogramas
         int_Produc = 1
         int_CuoDbl = r_int_CuoExt
         dbl_ValInm = r_dbl_ValViv + CDbl(pnl_NuevoSaldoTC.Caption)
         dbl_CuoIni = r_dbl_ValViv - CDbl(pnl_NuevoSaldoTNC.Caption)
         dbl_MtoCon = CDbl(pnl_NuevoSaldoTC.Caption)
         dbl_MtoTas = r_dbl_MtoAse
         int_PlaPre = l_int_NumCuo
         dbl_TasInt = l_dbl_TasInt
         dbl_TasCof = l_dbl_TasCof
         dbl_ComCof = l_dbl_ComCof
         dat_FecDes = CDate(Format(r_str_FecDes, "dd/mm/yyyy"))
         int_DiaVct = r_int_DiaPag
         int_PerGra = 0
         str_PriVct = r_int_PriVct
         dbl_Portes = CDbl(r_dbl_Portes)
         dbl_SegViv = CDbl(r_dbl_FoiViv)
         int_TipSDe = r_int_TipSeg - 10
         dbl_SegDes = CDbl(r_dbl_FoIDes)
         
         'Calculando cronogramas
         Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
         Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, dbl_MtoTas, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
         
         'Mostrando Cronograma 1
         Call fs_Muestra_Cronograma1
         
         'Validando Monto de Seguros
         Call fs_Validar_Seguro
      
      Case InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) '"003"
         'NUEVA rutina de generacion de cronogramas
         int_Produc = 1
         int_CuoDbl = r_int_CuoExt
         dbl_ValInm = r_dbl_ValViv + CDbl(pnl_NuevoSaldoTC.Caption)
         dbl_CuoIni = r_dbl_ValViv - CDbl(pnl_NuevoSaldoTNC.Caption)
         dbl_MtoCon = CDbl(pnl_NuevoSaldoTC.Caption)
         dbl_MtoTas = r_dbl_MtoAse
         int_PlaPre = l_int_NumCuo
         dbl_TasInt = l_dbl_TasInt
         dbl_TasCof = l_dbl_TasCof
         dbl_ComCof = l_dbl_ComCof
         dat_FecDes = CDate(Format(r_str_FecDes, "dd/mm/yyyy"))
         int_DiaVct = r_int_DiaPag
         int_PerGra = 0
         str_PriVct = r_int_PriVct
         dbl_Portes = CDbl(r_dbl_Portes)
         dbl_SegViv = CDbl(r_dbl_FoiViv)
         int_TipSDe = r_int_TipSeg - 10
         dbl_SegDes = CDbl(r_dbl_FoIDes)
         
         'Calculando cronogramas
         Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
         Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, dbl_MtoTas, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
         
         'Mostrando Cronograma 1
         Call fs_Muestra_Cronograma1
         
         'Validando Monto de Seguros
         Call fs_Validar_Seguro
      
      Case InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd)   '"002", "011",
         'NUEVA rutina de generacion de cronogramas
         int_Produc = 2
         int_CuoDbl = r_int_CuoExt
         dbl_ValInm = r_dbl_ValViv
         dbl_CuoIni = r_dbl_ValViv - CDbl(pnl_NuevoSaldoTNC.Caption)
         dbl_MtoCon = 0
         dbl_MtoTas = r_dbl_MtoAse
         int_PlaPre = l_int_NumCuo
         dbl_TasInt = l_dbl_TasInt
         dbl_TasCof = l_dbl_TasCof
         dbl_ComCof = l_dbl_ComCof
         dat_FecDes = CDate(Format(r_str_FecDes, "dd/mm/yyyy"))
         int_DiaVct = r_int_DiaPag
         int_PerGra = 0
         str_PriVct = r_int_PriVct
         dbl_Portes = CDbl(r_dbl_Portes)
         dbl_SegViv = CDbl(r_dbl_FoiViv)
         int_TipSDe = r_int_TipSeg - 10
         dbl_SegDes = CDbl(r_dbl_FoIDes)
         
         'Calculando cronogramas
         Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
         Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, dbl_MtoTas, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
         
         'Mostrando Cronograma 1
         Call fs_Muestra_Cronograma1
         
         'Validando Monto de Seguros
         Call fs_Validar_Seguro
         
         
      Case InStr(moddat_g_str_AgrMIHG, moddat_g_str_CodPrd) Or InStr(moddat_g_str_Agr2MIC, moddat_g_str_CodPrd) Or InStr(moddat_g_str_Agr2MIC, moddat_g_str_CodPrd) Or InStr(moddat_g_str_Agr2FMV, moddat_g_str_CodPrd)  '"004", "006", "007", "009", "010", "012", "013", "014", "015", "016", "017", "018"
         'NUEVA rutina de generacion de cronogramas
         int_Produc = 1
         int_CuoDbl = r_int_CuoExt
         dbl_ValInm = r_dbl_ValViv + CDbl(pnl_NuevoSaldoTC.Caption)
         dbl_CuoIni = r_dbl_ValViv - CDbl(pnl_NuevoSaldoTNC.Caption)
         dbl_MtoCon = CDbl(pnl_NuevoSaldoTC.Caption)
         dbl_MtoTas = r_dbl_MtoAse
         int_PlaPre = l_int_NumCuo
         dbl_TasInt = l_dbl_TasInt
         dbl_TasCof = l_dbl_TasCof
         dbl_ComCof = l_dbl_ComCof
         dat_FecDes = CDate(Format(r_str_FecDes, "dd/mm/yyyy"))
         int_DiaVct = r_int_DiaPag
         int_PerGra = 0
         str_PriVct = r_int_PriVct
         dbl_Portes = CDbl(r_dbl_Portes)
         dbl_SegViv = CDbl(r_dbl_FoiViv)
         int_TipSDe = r_int_TipSeg - 10
         dbl_SegDes = CDbl(r_dbl_FoIDes)
         
         'Calculando cronogramas
         Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
         Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, dbl_MtoTas, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
         
         'Mostrando Cronograma 1
         Call fs_Muestra_Cronograma1
         
         'Validando Monto de Seguros
         Call fs_Validar_Seguro
         
      Case InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) '"019", "020", "021", "022", "023", "024", "025"
         'NUEVA rutina de generacion de cronogramas
         int_Produc = 3
         int_CuoDbl = r_int_CuoExt
         dbl_ValInm = r_dbl_ValViv
         dbl_CuoIni = r_dbl_ValViv - CDbl(pnl_NuevoSaldoTNC.Caption)
         dbl_MtoCon = 0
         dbl_MtoTas = r_dbl_MtoAse
         int_PlaPre = l_int_NumCuo
         dbl_TasInt = l_dbl_TasInt
         dbl_TasCof = l_dbl_TasCof
         dbl_ComCof = l_dbl_ComCof
         dat_FecDes = CDate(Format(r_str_FecDes, "dd/mm/yyyy"))
         int_DiaVct = r_int_DiaPag
         int_PerGra = 0
         str_PriVct = r_int_PriVct
         dbl_Portes = CDbl(r_dbl_Portes)
         dbl_SegViv = CDbl(r_dbl_FoiViv)
         int_TipSDe = r_int_TipSeg - 10
         dbl_SegDes = CDbl(r_dbl_FoIDes)
         
         'Calculando cronogramas
         Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
         Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, dbl_MtoTas, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
         
         'Mostrando Cronograma 1
         Call fs_Muestra_Cronograma1
         
         'Validando Monto de Seguros
         Call fs_Validar_Seguro
         
   End Select
   
   'Determina cuota mensual
   Select Case r_int_CuoExt
      Case 1:
          pnl_NuevaCuota.Caption = Format(grd_CliNCo_Listad.TextMatrix(1, 7), "###,###,##0.00") & " "          'pnl_NuevaCuota.Caption = Format(l_Arr_TNC_Cli(2, 9), "###,###,##0.00") & " "
      Case 2:
         r_int_NumCuo = Month(grd_CliNCo_Listad.TextMatrix(1, 1))                                              'Month(l_Arr_TNC_Cli(2, 2))
         If r_int_NumCuo = 7 Then
            pnl_NuevaCuota.Caption = Format(grd_CliNCo_Listad.TextMatrix(3, 7), "###,###,##0.00") & " "        'pnl_NuevaCuota.Caption = Format(l_Arr_TNC_Cli(4, 9), "###,###,##0.00") & " "
         Else
            pnl_NuevaCuota.Caption = Format(grd_CliNCo_Listad.TextMatrix(1, 7), "###,###,##0.00") & " "        'pnl_NuevaCuota.Caption = Format(l_Arr_TNC_Cli(2, 9), "###,###,##0.00") & " "
         End If
      Case 3:
         r_int_NumCuo = Month(grd_CliNCo_Listad.TextMatrix(1, 1))                                              'Month(l_Arr_TNC_Cli(2, 2))
         If r_int_NumCuo = 12 Then
            pnl_NuevaCuota.Caption = Format(grd_CliNCo_Listad.TextMatrix(3, 7), "###,###,##0.00") & " "        'pnl_NuevaCuota.Caption = Format(l_Arr_TNC_Cli(4, 9), "###,###,##0.00") & " "
         Else
            pnl_NuevaCuota.Caption = Format(grd_CliNCo_Listad.TextMatrix(1, 7), "###,###,##0.00") & " "        'pnl_NuevaCuota.Caption = Format(l_Arr_TNC_Cli(2, 9), "###,###,##0.00") & " "
         End If
      Case 4:
         r_int_NumCuo = Month(grd_CliNCo_Listad.TextMatrix(1, 1))                                              'Month(l_Arr_TNC_Cli(2, 2))
         If r_int_NumCuo = 7 Then
            pnl_NuevaCuota.Caption = Format(grd_CliNCo_Listad.TextMatrix(3, 7), "###,###,##0.00") & " "        'pnl_NuevaCuota.Caption = Format(l_Arr_TNC_Cli(4, 9), "###,###,##0.00") & " "
         Else
            If r_int_NumCuo = 12 Then
               pnl_NuevaCuota.Caption = Format(grd_CliNCo_Listad.TextMatrix(3, 7), "###,###,##0.00") & " "     'pnl_NuevaCuota.Caption = Format(l_Arr_TNC_Cli(4, 9), "###,###,##0.00") & " "
            Else
               pnl_NuevaCuota.Caption = Format(grd_CliNCo_Listad.TextMatrix(1, 7), "###,###,##0.00") & " "     'pnl_NuevaCuota.Caption = Format(l_Arr_TNC_Cli(2, 9), "###,###,##0.00") & " "
            End If
         End If
   End Select
End Sub

Private Sub fs_Muestra_Cronograma1()
Dim r_dbl_Cuo_Capita    As Double
Dim r_dbl_Cuo_Intere    As Double
Dim r_dbl_Cuo_SegPre    As Double
Dim r_dbl_Cuo_SegViv    As Double
Dim r_dbl_Cuo_Portes    As Double
Dim r_dbl_Cuo_TotCuo    As Double
Dim r_dbl_Tot_Capita    As Double
Dim r_dbl_Tot_Intere    As Double
Dim r_dbl_Tot_SegPre    As Double
Dim r_dbl_Tot_SegViv    As Double
Dim r_dbl_Tot_Portes    As Double
Dim r_dbl_Tot_TotCuo    As Double
Dim r_int_Contad        As Integer

   grd_CliNCo_Listad.Redraw = False
   Call gs_LimpiaGrid(grd_CliNCo_Listad)
   r_dbl_Tot_Capita = 0
   r_dbl_Tot_Intere = 0
   r_dbl_Tot_SegPre = 0
   r_dbl_Tot_SegViv = 0
   r_dbl_Tot_Portes = 0
   r_dbl_Tot_TotCuo = 0
   
   For r_int_Contad = 1 To UBound(l_Arr_TNC_Cli)
      grd_CliNCo_Listad.Rows = grd_CliNCo_Listad.Rows + 1
      grd_CliNCo_Listad.Row = grd_CliNCo_Listad.Rows - 1
      
      r_dbl_Cuo_Capita = CDbl(Format(l_Arr_TNC_Cli(r_int_Contad, 4), "###,##0.00"))
      r_dbl_Cuo_Intere = CDbl(Format(l_Arr_TNC_Cli(r_int_Contad, 5), "###,##0.00"))
      r_dbl_Cuo_SegPre = CDbl(Format(l_Arr_TNC_Cli(r_int_Contad, 6), "###,##0.00"))
      r_dbl_Cuo_SegViv = CDbl(Format(l_Arr_TNC_Cli(r_int_Contad, 7), "###,##0.00"))
      r_dbl_Cuo_Portes = CDbl(Format(l_Arr_TNC_Cli(r_int_Contad, 8), "###,##0.00"))
      r_dbl_Cuo_TotCuo = CDbl(Format(l_Arr_TNC_Cli(r_int_Contad, 9), "###,##0.00"))
      r_dbl_Tot_Capita = r_dbl_Tot_Capita + r_dbl_Cuo_Capita
      r_dbl_Tot_Intere = r_dbl_Tot_Intere + r_dbl_Cuo_Intere
      r_dbl_Tot_SegPre = r_dbl_Tot_SegPre + r_dbl_Cuo_SegPre
      r_dbl_Tot_SegViv = r_dbl_Tot_SegViv + r_dbl_Cuo_SegViv
      r_dbl_Tot_Portes = r_dbl_Tot_Portes + r_dbl_Cuo_Portes
      r_dbl_Tot_TotCuo = r_dbl_Tot_TotCuo + r_dbl_Cuo_TotCuo
      
      grd_CliNCo_Listad.Col = 0
      grd_CliNCo_Listad.Text = Format(r_int_Contad + l_int_PagCuo, "000")
      
      grd_CliNCo_Listad.Col = 1
      grd_CliNCo_Listad.Text = Format(l_Arr_TNC_Cli(r_int_Contad, 2), "dd/mm/yyyy")
      
      grd_CliNCo_Listad.Col = 2
      grd_CliNCo_Listad.Text = Format(r_dbl_Cuo_Capita, "###,##0.00")
      
      grd_CliNCo_Listad.Col = 3
      grd_CliNCo_Listad.Text = Format(r_dbl_Cuo_Intere, "###,##0.00")
      
      grd_CliNCo_Listad.Col = 4
      grd_CliNCo_Listad.Text = Format(r_dbl_Cuo_SegPre, "###,##0.00")
      
      grd_CliNCo_Listad.Col = 5
      grd_CliNCo_Listad.Text = Format(r_dbl_Cuo_SegViv, "###,##0.00")
      
      grd_CliNCo_Listad.Col = 6
      grd_CliNCo_Listad.Text = Format(r_dbl_Cuo_Portes, "###,##0.00")
      
      grd_CliNCo_Listad.Col = 7
      grd_CliNCo_Listad.Text = Format(r_dbl_Cuo_TotCuo, "###,##0.00")
      
      grd_CliNCo_Listad.Col = 8
      grd_CliNCo_Listad.Text = Format(l_Arr_TNC_Cli(r_int_Contad, 10), "###,##0.00")
   Next r_int_Contad
   
   grd_CliNCo_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_CliNCo_Listad)
   pnl_CliNCo_Capita.Caption = Format(r_dbl_Tot_Capita, "###,##0.00") & " "
   pnl_CliNCo_Intere.Caption = Format(r_dbl_Tot_Intere, "###,##0.00") & " "
   pnl_CliNCo_SegPre.Caption = Format(r_dbl_Tot_SegPre, "###,##0.00") & " "
   pnl_CliNCo_SegViv.Caption = Format(r_dbl_Tot_SegViv, "###,##0.00") & " "
   pnl_CliNCo_OtrCar.Caption = Format(r_dbl_Tot_Portes, "###,##0.00") & " "
   pnl_CliNCo_TotCuo.Caption = Format(r_dbl_Tot_TotCuo, "###,##0.00") & " "
End Sub

'********************************
' LIQUIDACION CREDITO MIVIVIENDA
'********************************
Private Sub fs_PpgPar_Mivivienda(ByVal p_flg_guardar As Boolean, ByRef p_rut_Guardo As String)
Dim r_obj_Excel      As Excel.Application
Dim r_int_NroFil     As Integer
Dim r_int_ColumC     As Integer
Dim r_int_ColumF     As Integer
Dim r_int_ColumK     As Integer
Dim r_int_ColumL     As Integer
Dim r_int_ColumM     As Integer
Dim r_int_ColumN     As Integer
Dim r_int_ColumO     As Integer
Dim r_int_ColumP     As Integer
Dim r_int_ColumQ     As Integer
Dim r_int_ColumR     As Integer
Dim r_int_ColumS     As Integer
Dim r_int_ColumT     As Integer
Dim r_int_ColumV     As Integer
Dim r_int_ColumW     As Integer
Dim r_int_ColumX     As Integer
Dim r_int_ColumY     As Integer
Dim r_int_ColumZ     As Integer

   r_int_NroFil = 2
   r_int_ColumC = 3
   r_int_ColumF = 6
   r_int_ColumK = 11
   r_int_ColumL = 12
   r_int_ColumM = 13
   r_int_ColumN = 14
   r_int_ColumO = 15
   r_int_ColumP = 16
   r_int_ColumQ = 17
   r_int_ColumR = 18
   r_int_ColumS = 19
   r_int_ColumT = 20
   r_int_ColumV = 22
   r_int_ColumW = 23
   r_int_ColumX = 24
   r_int_ColumY = 25
   r_int_ColumZ = 26
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      'CENTRADO DE LA PAGINA
      '.PageSetup.CenterHorizontally = True
      .PageSetup.Orientation = xlLandscape
      
      'MARGENES
      .PageSetup.LeftMargin = Application.CentimetersToPoints(1)
      .PageSetup.RightMargin = Application.CentimetersToPoints(1)
      .PageSetup.TopMargin = Application.CentimetersToPoints(1)
      .PageSetup.BottomMargin = Application.CentimetersToPoints(1)
      
      .Range("A1:A1").ColumnWidth = 1.14
      .Range("B1:AB1").ColumnWidth = 3.57
      .Range("A2:A46").RowHeight = 12
      
      'BORDERS
      .Range("B2:AA2").Borders(xlEdgeTop).Weight = xlMedium
      .Range("B4:AA4").Borders(xlEdgeTop).Weight = xlMedium
      .Range("B6:AA6").Borders(xlEdgeTop).Weight = xlMedium
      
      If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) = 2 Then
        .Range("B41:AA41").Borders(xlEdgeBottom).Weight = xlMedium
        .Range("B2:AA41").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("AA2:AA41").Borders(xlEdgeRight).Weight = xlMedium
      Else
        .Range("B40:AA40").Borders(xlEdgeBottom).Weight = xlMedium
        .Range("B2:AA40").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("AA2:AA40").Borders(xlEdgeRight).Weight = xlMedium
        .Range("B1:AA46").Font.Name = "Arial"
      End If
      
      'Font
      .Range("B1:AA46").Font.Name = "Arial"
      .Range("B1:AA40").Font.Size = 9
      
      'Fecha de realizado el prepago
      .Range("T1") = "FECHA DE EMISIÓN: "
      .Range("T1").Font.Bold = True
      .Range("T1:X1").Merge
      .Range("Y1:AA1").Merge
      .Range("Y1") = "'" & moddat_g_str_FecSis
      .Range("T1:X1").HorizontalAlignment = xlHAlignCenter
      .Range("Y1:AA1").HorizontalAlignment = xlHAlignCenter
      
      'Linea 1 - Titulo
      .Range("B2:AA3").Merge
      .Range("B2") = "LIQUIDACION PREPAGO PARCIAL - " & moddat_g_str_NomPrd & " - MONEDA " & moddat_g_str_Moneda
      .Range("B2:AA3").HorizontalAlignment = xlHAlignCenter
      .Range("B2:AA3").VerticalAlignment = xlCenter
      .Range("B2:AA3").Font.Size = 12
      .Range("B4:AA4").Font.Size = 12
      .Range("B6:AA6").Font.Size = 12
      .Range("B2:AA2").Font.Bold = True
      .Range("B4:AA4").Font.Bold = True
      .Range("B6:AA6").Font.Bold = True
'      .Range("B2:AA3").RowHeight = 12
      
      r_int_NroFil = r_int_NroFil + 2
      
      'Linea 2 - Datos del Cliente
      .Range("B4:AA5").Merge
      .Range("B4") = "OPERACIÓN: " & moddat_g_str_NumOpe & " - CLIENTE: (DNI-" & moddat_g_str_NumDoc & ") " & moddat_g_str_NomCli
      .Cells(r_int_NroFil, r_int_ColumC).HorizontalAlignment = xlCenter
      .Cells(r_int_NroFil, r_int_ColumC).VerticalAlignment = xlCenter
      .Range("B4:AA5").Font.Size = 11
'      .Range("B4:AA5").RowHeight = 12
       r_int_NroFil = r_int_NroFil + 3
        
      'Linea 3 - Fecha de Desembolso o última cuota TC (A)
      .Cells(r_int_NroFil, r_int_ColumC) = "Fecha de Desembolso o última cuota TC (A)"
      .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumL) = "'" & Format(CDate(pnl_UltPagTC.Caption), "dd-mm-yy")
      .Cells(r_int_NroFil, r_int_ColumL).HorizontalAlignment = xlHAlignRight

      'Linea 3 - Días de interés TNC (C-B)
      .Cells(r_int_NroFil, r_int_ColumR) = "Días de interés TNC (C-B)"
      .Range("Y" & r_int_NroFil & ":Z" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumY) = CInt(pnl_DiasTNC.Caption)
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 4 - Saldo TNC al (B)
      r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, r_int_ColumC) = "Saldo TNC al (B)"
      .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumL).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumT).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumL) = "'" & Format(CDate(pnl_UltPagTNC.Caption), "dd-mm-yy")
      
       'Linea 4 - Sumatoria pnl_SaldoTNC2 + pnl_SaldoTC2
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(pnl_SaldoTNC2) + CDbl(pnl_SaldoTC2), "###,###.00")

      'MARCO
      .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumP).Borders(xlEdgeRight).Weight = xlThin
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous

      'Linea 4 - Días de interés TC (C-A)
      .Cells(r_int_NroFil, r_int_ColumR) = "Días de interés TC (C-A)"
      .Range("Y" & r_int_NroFil & ":Z" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumY) = CInt(pnl_DiasTC.Caption)
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 5 - Fecha de corte (fecha del prepago) (C)
      .Cells(r_int_NroFil, r_int_ColumC) = "Fecha de corte (fecha del prepago) (C)"
'      If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) <> 2 Then
'        .Range("C" & r_int_nrofil & ":P" & r_int_nrofil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
'      End If
      .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumL) = "'" & Format(CDate(ipp_FecPre), "dd-mm-yy")
      .Cells(r_int_NroFil, r_int_ColumL).HorizontalAlignment = xlHAlignRight
      
      'Linea 5 - Tasa de interés anual (%)
      .Cells(r_int_NroFil, r_int_ColumR) = "Tasa de interés anual (%)"
      .Range("Y" & r_int_NroFil & ":Z" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumY) = l_dbl_TasInt
      .Cells(r_int_NroFil, r_int_ColumY).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.0000"
      r_int_NroFil = r_int_NroFil + 1
      
      'Si es reducción de plazo
      If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) = 2 Then
        .Cells(r_int_NroFil, r_int_ColumC) = "Reducción de Plazo"
        .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
        .Cells(r_int_NroFil, r_int_ColumL) = cmb_RedPlz.Text
        .Cells(r_int_NroFil, r_int_ColumL).HorizontalAlignment = xlHAlignRight
'        .Range("C" & r_int_nrofil & ":P" & r_int_nrofil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      End If
     
      'Linea 6 - Tasa de Seguro Inmueble (%)
      .Cells(r_int_NroFil, r_int_ColumR) = "Tasa de Seguro Inmueble (%)"
      .Range("Y" & r_int_NroFil & ":Z" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumY) = l_dbl_SegInm
      .Range("Y" & r_int_NroFil & ":Y" & r_int_NroFil & "").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.0000"
      r_int_NroFil = r_int_NroFil + 1

      'Linea 7 - Tasa de Seguro Desgravamen (%)
      .Cells(r_int_NroFil, r_int_ColumR) = "Tasa de Seguro Desgravamen (%)"
      .Range("Y" & r_int_NroFil & ":Z" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumY) = l_dbl_SegDes
      .Range("Y" & r_int_NroFil & ":Y" & r_int_NroFil & "").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.0000"
      r_int_NroFil = r_int_NroFil + 2
      
      'Linea 8 - Monto Depositado
      .Range("C" & r_int_NroFil & ":M" & (r_int_NroFil) & "").Merge
      .Range("N" & r_int_NroFil & ":P" & (r_int_NroFil) & "").Merge
      .Cells(r_int_NroFil, r_int_ColumC) = "Monto Depositado"
      .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(txt_Mto_Deposito.Text), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumC).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True

       'MARCO
      .Cells(r_int_NroFil, r_int_ColumC).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumC).VerticalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumN).VerticalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumP).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
      .Range("C" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumP).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
      .Range("C" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 2
      
      'Linea 9 - Intereses TNC a la fecha
      .Cells(r_int_NroFil, r_int_ColumC) = "Intereses TNC a la fecha"
      .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumL) = CDbl(txt_InteresTNC.Text)
      .Cells(r_int_NroFil, r_int_ColumL).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 10 - Intereses TC a la fecha
      .Cells(r_int_NroFil, r_int_ColumC) = "Intereses TC a la fecha"
      .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumL) = CDbl(txt_InteresTC.Text)
      .Cells(r_int_NroFil, r_int_ColumL).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      
      'Linea 10 - Saldo antes del Prepago
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumS) = "Saldo antes del prepago"
      
       'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, r_int_ColumT).HorizontalAlignment = xlHAlignCenter

      r_int_NroFil = r_int_NroFil + 1
          
      'Linea 11 - Seguro Desgravamen
      .Cells(r_int_NroFil, r_int_ColumC) = "Seguro Desgravamen"
      .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumL) = CDbl(txt_SegDes.Text)
      .Cells(r_int_NroFil, r_int_ColumL).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      
      'Linea 11 - Saldo antes del Prepago - TNC
      .Range("S" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
      .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS) = "TNC"
      .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_SaldoTNC1.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 12 - Seguro inmueble
      .Cells(r_int_NroFil, r_int_ColumC) = "Seguro Inmueble"
      .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumL) = CDbl(txt_SegInm.Text)
      .Cells(r_int_NroFil, r_int_ColumL).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
     
      'Linea 12 - Saldo antes del Prepago - TC
      .Range("S" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
      .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS) = "TC"
      .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_SaldoTC1.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
'      If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) <> 2 Then
'        .Range("C" & r_int_nrofil & ":Q" & r_int_nrofil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
'      End If
      r_int_NroFil = r_int_NroFil + 1
    
      'Linea 13 - Interes PBP
      .Cells(r_int_NroFil, r_int_ColumC) = "Interés PBP"
      .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumL) = CDbl(pnl_IntPbp.Caption)
      .Cells(r_int_NroFil, r_int_ColumL).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      
      'Linea 13 - Saldo antes del prepago - Capital PBP
      .Range("S" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
      .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS) = "PBP"
      .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_CapPbp.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
            
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 14 - Deuda Pendiente
      .Cells(r_int_NroFil, r_int_ColumC) = "Deuda Pendiente"
      .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumL) = CDbl(pnl_DeuPen.Caption)
      .Cells(r_int_NroFil, r_int_ColumL).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"

       'Linea 14 - Total Saldo antes del Prepago
      .Range("S" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
      .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumS) = "Total"
      .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_SaldoTNC2) + CDbl(pnl_SaldoTC2) + CDbl(pnl_CapPbp), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
      .Range("N" & r_int_NroFil & ":O" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumW).Font.Bold = True
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 1
      
'      If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) = 2 Then
'         r_int_nrofil = r_int_nrofil + 1
'      End If
      
      'Linea 15 - Total de Interés, Seguros y Vencidos
      .Cells(r_int_NroFil, r_int_ColumC) = "Total de Interés, seguros y vencidos"
      .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumN) = CDbl(CDbl(txt_InteresTNC.Text) + CDbl(txt_InteresTC.Text) + CDbl(txt_SegDes.Text) + CDbl(txt_SegInm.Text) + CDbl(pnl_IntPbp.Caption)) + CDbl(pnl_DeuPen.Caption)
      .Cells(r_int_NroFil, r_int_ColumC).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignRight

      'MARCO
      .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumP).Borders(xlEdgeRight).Weight = xlThin
      .Range("C" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, r_int_ColumC).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
      r_int_NroFil = r_int_NroFil + 3
      
      'Linea 16 - Monto de Prepago a Aplicar
      .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumC).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumC) = "Monto de Prepago a Aplicar"
      .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(pnl_MtoApl.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumC).HorizontalAlignment = xlHAlignCenter
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumP).Borders(xlEdgeRight).Weight = xlThin
      .Range("C" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      
      'Línea 16 - Cancelación de PBP x Cobrar
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumS) = "Cancelación de PBP x Cobrar"
      
       'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      
      r_int_NroFil = r_int_NroFil + 1
      
      'Línea 17 - Cancelación de PBP x Cobrar - Monto Aplicar
      .Range("S" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
      .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS) = "Monto Aplicar"
      .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_MtoApl.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 1
      
      'Línea 18 - Cancelación de PBP x Cobrar - PBP
      .Range("S" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
      .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS) = "PBP"
      .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_CapPbp.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 19 - Saldo Distribuir
      .Range("S" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
      .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumS) = "Saldo Distribuir"
      .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_MtoApl) - CDbl(pnl_CapPbp), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
      .Range("N" & r_int_NroFil & ":O" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumW).Font.Bold = True
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 2
         
      'Linea 20 - Monto de Prepago a Distribuir
      .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumC).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumC) = "Monto de Prepago a Distribuir"
      .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(pnl_MtoApl_Fin.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumC).HorizontalAlignment = xlHAlignCenter
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumP).Borders(xlEdgeRight).Weight = xlThin
      .Range("C" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      
       'Linea 20 - Distribución del Prepago
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumS) = "Distribución del Prepago"
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 21 - Distribución TNC prepago
      .Range("S" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
      .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS) = "TNC"
      .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(txt_AplTNC.Text), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 22 - Distribución TC prepago
      .Range("S" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
      .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS) = "TC"
      .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(txt_ApliTC.Text), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 23 - Distribución total prepago
      .Range("S" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
      .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumW).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumS) = "Total"
      .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_MtoApl_Fin.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 2
      
      'Linea 24 - Saldo después prepago
      .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumC).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumC).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumC) = "Saldo después del prepago"
      .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(pnl_NuevoSaldoTNC.Caption) + CDbl(pnl_NuevoSaldoTC.Caption), "###,###.00")
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumP).Borders(xlEdgeRight).Weight = xlThin
      .Range("C" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      
      'Linea 24 - Saldo después prepago
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumS) = "Saldo después del prepago"
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 25 - Saldo TNC después prepago
      .Range("S" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
      .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS) = "TNC"
      .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_NuevoSaldoTNC.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 26 - Saldo TC después prepago
      .Range("S" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
      .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS) = "TC"
      .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_NuevoSaldoTC.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 1
     
      'Linea 27 - Total Saldo después prepago
      .Range("S" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
      .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumW).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumS) = "Total"
      .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_NuevoSaldoTNC.Caption) + CDbl(pnl_NuevoSaldoTC.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 2
      
      'Linea 28 - Importe nueva cuota
      .Cells(r_int_NroFil, r_int_ColumC) = "IMPORTE NUEVA CUOTA"
      .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(pnl_NuevaCuota.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumC).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumC).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
'      .Cells(r_int_nrofil, r_int_ColumC).RowHeight = 12
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumP).Borders(xlEdgeRight).Weight = xlThin
      .Range("C" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, r_int_ColumN).Select

'      'Linea - Valor inmueble
'      .Cells(r_int_nrofil, r_int_ColumC) = "Valor asegurable inmueble"
'      .Range("K" & r_int_nrofil & ":M" & r_int_nrofil & "").Merge
'      .Cells(r_int_nrofil, r_int_ColumK).Font.Bold = True
'      .Cells(r_int_nrofil, r_int_ColumK) = Format(CDbl(pnl_Val_AsgInm.Caption), "###,###.00")
'      r_int_nrofil = r_int_nrofil + 2

      If CInt(moddat_g_int_TipMon) = 1 Then
         r_obj_Excel.Selection.NumberFormat = "$###,##0.00_);[Red]($###,##0.00)"
         If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) = 2 Then
            .Cells(43, 2) = "'" & "- Esta liquidación es válida solo a la fecha de corte."
            .Cells(44, 2) = "'" & "- Cualquier gasto adicional que no contemple la presente liquidación, será informado al cliente y tendrá que ser cancelado"
            .Cells(45, 2) = "'" & "  antes de realizar el abono que consigna la presente liquidación."
            .Cells(46, 2) = "'" & "- Realizar el depósito en la cuenta Nº 0011-0369-02-00090532 del BBVA Banco Continental."
            .Cells(47, 2) = "'" & "- Realizar el depósito en la cuenta Nº 011-369-000200090532-69 del BBVA Banco Continental desde otro Banco."
            .Range(.Cells(43, 2), .Cells(47, 2)).Font.Size = 11
            .Range(.Cells(43, 2), .Cells(47, 2)).Font.Bold = True
         Else
            .Cells(42, 2) = "'" & "- Esta liquidación es válida solo a la fecha de corte."
            .Cells(43, 2) = "'" & "- Cualquier gasto adicional que no contemple la presente liquidación, será informado al cliente y tendrá que ser cancelado"
            .Cells(44, 2) = "'" & "  antes de realizar el abono que consigna la presente liquidación."
            .Cells(45, 2) = "'" & "- Realizar el depósito en la cuenta Nº 0011-0369-02-00090532 del BBVA Banco Continental."
            .Cells(46, 2) = "'" & "- Realizar el depósito en la cuenta Nº 011-369-000200090532-69 del BBVA Banco Continental desde otro Banco."
            .Range(.Cells(42, 2), .Cells(46, 2)).Font.Size = 11
            .Range(.Cells(42, 2), .Cells(46, 2)).Font.Bold = True
         End If
      Else
         r_obj_Excel.Selection.NumberFormat = "[$$]#,##0.00;[Red][$$]#,##0.00"
         If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) = 2 Then
            .Cells(43, 2) = "'" & "- Esta liquidación es válida solo a la fecha de corte."
            .Cells(44, 2) = "'" & "- Cualquier gasto adicional que no contemple la presente liquidación, será informado al cliente y tendrá que ser cancelado"
            .Cells(45, 2) = "'" & "  antes de realizar el abono que consigna la presente liquidación."
            .Cells(46, 2) = "'" & "- Realizar el depósito en la cuenta Nº 0011-0369-02-00090540 del BBVA Banco Continental."
            .Range(.Cells(43, 2), .Cells(46, 2)).Font.Size = 11
            .Range(.Cells(43, 2), .Cells(46, 2)).Font.Bold = True
         Else
            .Cells(42, 2) = "'" & "- Esta liquidación es válida solo a la fecha de corte."
            .Cells(43, 2) = "'" & "- Cualquier gasto adicional que no contemple la presente liquidación, será informado al cliente y tendrá que ser cancelado"
            .Cells(44, 2) = "'" & "  antes de realizar el abono que consigna la presente liquidación."
            .Cells(45, 2) = "'" & "- Realizar el depósito en la cuenta Nº 0011-0369-02-00090540 del BBVA Banco Continental."
            .Range(.Cells(42, 2), .Cells(45, 2)).Font.Size = 11
            .Range(.Cells(42, 2), .Cells(45, 2)).Font.Bold = True
         End If
      End If
   End With

   p_rut_Guardo = ""
   If p_flg_guardar = True Then
      p_rut_Guardo = Format(date, "yyyymmdd") & "_PPG_" & moddat_g_str_NumOpe & ".XLSX"
      r_obj_Excel.ActiveWorkbook.SaveAs (g_str_RutLog & "\" & p_rut_Guardo)
      
      r_obj_Excel.Application.Quit
      Set r_obj_Excel = Nothing
   Else
      r_obj_Excel.Visible = True
      Set r_obj_Excel = Nothing
   End If

End Sub

'******************************
' LIQUIDACION CREDITO MICASITA
'******************************
Private Sub fs_PpgPar_Micasita(ByVal p_flg_guardar As Boolean, ByRef p_rut_Guardo As String)
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_NroFil     As Integer
   Dim r_int_ColumC     As Integer
   Dim r_int_ColumD     As Integer
   Dim r_int_ColumE     As Integer
   Dim r_int_ColumF     As Integer
   Dim r_int_ColumK     As Integer
   Dim r_int_ColumL     As Integer
   Dim r_int_ColumM     As Integer
   Dim r_int_ColumN     As Integer
   Dim r_int_ColumO     As Integer
   Dim r_int_ColumP     As Integer
   Dim r_int_ColumQ     As Integer
   Dim r_int_ColumR     As Integer

   r_int_NroFil = 3
   r_int_ColumC = 3
   r_int_ColumD = 4
   r_int_ColumE = 5
   r_int_ColumF = 6
   r_int_ColumK = 11
   r_int_ColumL = 12
   r_int_ColumM = 13
   r_int_ColumN = 14
   r_int_ColumO = 15
   r_int_ColumP = 16
   r_int_ColumQ = 17
   r_int_ColumR = 18
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      'CENTRADO DE LA PAGINA
      '.PageSetup.CenterHorizontally = True
      .PageSetup.Orientation = xlLandscape
      
      'MARGENES
      .PageSetup.LeftMargin = Application.CentimetersToPoints(1)
      .PageSetup.RightMargin = Application.CentimetersToPoints(1)
      .PageSetup.TopMargin = Application.CentimetersToPoints(1)
      .PageSetup.BottomMargin = Application.CentimetersToPoints(1)
      
      .Range("A1:W1").ColumnWidth = 4.57
      .Range("A8:A36").RowHeight = 13.5
      
      'BORDERS
      .Range("C2:V3").Borders(xlEdgeTop).Weight = xlMedium
      .Range("C4:V4").Borders(xlEdgeTop).Weight = xlMedium
      .Range("C6:V6").Borders(xlEdgeTop).Weight = xlMedium
      If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) = 2 Then
        .Range("C32:V32").Borders(xlEdgeBottom).Weight = xlMedium
        .Range("C2:V32").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("V2:V32").Borders(xlEdgeRight).Weight = xlMedium
      Else
        .Range("C31:V31").Borders(xlEdgeBottom).Weight = xlMedium
        .Range("C2:V31").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("V2:V31").Borders(xlEdgeRight).Weight = xlMedium
      End If

      'Font
      .Range("A1:V44").Font.Name = "Arial"
      .Range("A1:V35").Font.Size = 9
      
      'Fecha de realizado el prepago
      .Range("O1") = "FECHA DE EMISIÓN: "
      .Range("O1").Font.Bold = True
      .Range("O1:R1").Merge
      .Range("S1:V1").Merge
      .Range("S1") = "'" & moddat_g_str_FecSis
      .Range("S1:V1").HorizontalAlignment = xlHAlignCenter
      .Range("O1:R1").HorizontalAlignment = xlHAlignCenter
      
      'Linea 3 - Titulo
      .Range("C2:V3").Merge
      .Range("C2") = "LIQUIDACION PREPAGO PARCIAL - " & moddat_g_str_NomPrd
      .Range("C2:V3").HorizontalAlignment = xlHAlignCenter
      .Range("C2:V3").VerticalAlignment = xlCenter
      .Range("C2:V3").Font.Size = 12
      .Range("C2:V3").Font.Bold = True
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 5 - Datos del cliente
      .Range("C4:V5").Merge
      .Range("C4") = "OPERACIÓN: " & moddat_g_str_NumOpe & " - CLIENTE: (DNI-" & moddat_g_str_NumDoc & ") " & moddat_g_str_NomCli
      .Range("C4:V5").HorizontalAlignment = xlHAlignCenter
      .Range("C4:V5").VerticalAlignment = xlCenter
      .Range("C4:V5").Font.Size = 11
      .Range("C4:V5").Font.Bold = True
      r_int_NroFil = r_int_NroFil + 3
      
      'Linea 9 - Saldo
      r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, r_int_ColumD) = "Saldo al"
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumK).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumK) = "'" & Format(CDate(pnl_UltPagTNC.Caption), "dd-mm-yy")
      .Range("N" & r_int_NroFil & ":Q" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(pnl_SaldoTNC2) + CDbl(pnl_SaldoTC2), "###,###.00") 'Format(moddat_g_dbl_SalCap + l_dbl_SalCon, "###,###.00")
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
      .Range("N" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("N" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 10 - Fecha prepago
      .Cells(r_int_NroFil, r_int_ColumD) = "Fecha de corte (fecha del prepago)"
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumK) = "'" & Format(CDate(ipp_FecPre), "dd-mm-yy")
      .Cells(r_int_NroFil, r_int_ColumK).HorizontalAlignment = xlHAlignRight
      If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) <> 2 Then
        .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      End If
      
      'Si es reducción de plazo
      If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) = 2 Then
        r_int_NroFil = r_int_NroFil + 1
        .Cells(r_int_NroFil, r_int_ColumD) = "Reducción de Plazo"
        .Cells(r_int_NroFil, r_int_ColumK) = cmb_RedPlz.Text
        .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
        .Cells(r_int_NroFil, r_int_ColumK).HorizontalAlignment = xlHAlignRight
        .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      End If
      r_int_NroFil = r_int_NroFil + 2
      
      'Linea 12 - Dias interes
      .Cells(r_int_NroFil, r_int_ColumD) = "Días de interés"
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumK) = CInt(pnl_DiasTNC.Caption)
      r_int_NroFil = r_int_NroFil + 1
      
'      'Linea  14 - Valor inmueble
'      .Cells(r_int_nrofil, r_int_ColumD) = "Valor asegurable inmueble"
'      .Range("K" & r_int_nrofil & ":L" & r_int_nrofil & "").Merge
'      .Cells(r_int_nrofil, r_int_ColumK).Font.Bold = True
'      .Cells(r_int_nrofil, r_int_ColumK) = Format(CDbl(pnl_Val_AsgInm.Caption), "###,###.00")
'      r_int_nrofil = r_int_nrofil + 2
      
      'Linea 16 - Tasa Interes anual
      .Cells(r_int_NroFil, r_int_ColumD) = "Tasa de interés anual (%)"
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumK).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumK) = l_dbl_TasInt
      .Cells(r_int_NroFil, r_int_ColumK).Select
      r_obj_Excel.Selection.NumberFormat = "###0.0000"
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 17 - Tasa seguro inmueble
      .Cells(r_int_NroFil, r_int_ColumD) = "Tasa de Seguro Inmueble (%)"
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumK).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumK) = l_dbl_SegInm
      .Cells(r_int_NroFil, r_int_ColumK).Select
      r_obj_Excel.Selection.NumberFormat = "###0.0000"
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 18 - Tasa seguro desgravamen
      .Cells(r_int_NroFil, r_int_ColumD) = "Tasa de Seguro Desgravamen (%)"
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumK).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumK) = l_dbl_SegDes
      .Cells(r_int_NroFil, r_int_ColumK).Select
      r_obj_Excel.Selection.NumberFormat = "###0.0000"
      r_int_NroFil = r_int_NroFil + 2
      
      'Linea 20 - monto depositado
      .Range("D" & r_int_NroFil & ":M" & (r_int_NroFil) & "").Merge
      .Range("N" & r_int_NroFil & ":P" & (r_int_NroFil) & "").Merge
      .Cells(r_int_NroFil, r_int_ColumD) = "Monto Depositado"
      .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(txt_Mto_Deposito.Text), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumD).VerticalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumN).VerticalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumD).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
      
       'MARCO
      .Cells(r_int_NroFil, r_int_ColumE).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumO).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumD).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
      .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Cells(r_int_NroFil, r_int_ColumD).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
      .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 2
      
      'Linea 22 - interes
      .Cells(r_int_NroFil, r_int_ColumD) = "Intereses a la fecha"
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumK) = CDbl(txt_InteresTNC.Text)
      .Cells(r_int_NroFil, r_int_ColumK).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 23 - seguro desgravamen
      .Cells(r_int_NroFil, r_int_ColumD) = "Seguro Desgravamen"
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumK) = CDbl(txt_SegDes.Text)
      .Cells(r_int_NroFil, r_int_ColumK).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 24 - seguro inmueble
      .Cells(r_int_NroFil, r_int_ColumD) = "Seguro Inmueble"
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumK) = CDbl(txt_SegInm.Text)
      .Cells(r_int_NroFil, r_int_ColumK).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 25 - Deuda Pendiente
      .Cells(r_int_NroFil, r_int_ColumD) = "Deuda Pendiente"
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumK) = CDbl(pnl_DeuPen.Caption)
      .Cells(r_int_NroFil, r_int_ColumK).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 26 - total gastos
      .Cells(r_int_NroFil, r_int_ColumD) = "Total de Interés y Seguros"
      .Range("D" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(CDbl(txt_InteresTNC.Text) + CDbl(txt_SegDes.Text) + CDbl(txt_SegInm.Text) + CDbl(pnl_DeuPen.Caption)), "###,##0.00")
      .Cells(r_int_NroFil, r_int_ColumD).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumD).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
      .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 2
      
      'Linea 27 - ITF
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumD) = "ITF (%)"
      .Cells(r_int_NroFil, r_int_ColumK) = "'0.005"
      .Cells(r_int_NroFil, r_int_ColumN) = CDbl(txt_MontoITF.Text)
      .Cells(r_int_NroFil, r_int_ColumN).Select
      r_obj_Excel.Selection.NumberFormat = "###0.00"
      .Cells(r_int_NroFil, r_int_ColumK).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignRight
      r_int_NroFil = r_int_NroFil + 2
      
      'Linea 29 - monto prepago
      .Cells(r_int_NroFil, r_int_ColumD) = "Monto del Prepago a Aplicar"
      .Range("D" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(pnl_MtoApl.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumD).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumD).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumD).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
      .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 2
      
      'Linea 31 - saldo reprogramar
      .Cells(r_int_NroFil, r_int_ColumD) = "Saldo después del prepago"
      .Range("D" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(pnl_NuevoSaldoTNC.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumD).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumD).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumD).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
      .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 2
      
      'Linea 33 - monto cuota
      .Cells(r_int_NroFil, r_int_ColumD) = "MONTO DE LA NUEVA CUOTA"
      .Range("D" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumN) = CDbl(pnl_NuevaCuota.Caption)
      .Cells(r_int_NroFil, r_int_ColumD).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumD).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumD).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
      .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Select
      r_int_NroFil = r_int_NroFil + 2
      
      If CInt(moddat_g_int_TipMon) = 1 Then
         r_obj_Excel.Selection.NumberFormat = "$###,##0.00_);[Red]($###,##0.00)"
         If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) = 2 Then
            .Cells(34, 3) = "'" & "- Esta liquidación es válida solo a la fecha de corte."
            .Cells(35, 3) = "'" & "- Cualquier gasto adicional que no contemple la presente liquidación, será informado al cliente y "
            .Cells(36, 3) = "'" & "  tendrá que ser cancelado antes de realizar el abono que consigna la presente liquidación."
            .Cells(37, 3) = "'" & "- Realizar el depósito en la cuenta Nº 0011-0369-02-00090532 del BBVA Banco Continental."
            .Cells(38, 3) = "'" & "- Realizar el depósito en la cuenta Nº 011-369-000200090532-69 del BBVA Banco Continental desde otro Banco."
            .Range(.Cells(34, 3), .Cells(38, 3)).Font.Size = 10
            .Range(.Cells(34, 3), .Cells(38, 3)).Font.Bold = True
         Else
            .Cells(33, 3) = "'" & "- Esta liquidación es válida solo a la fecha de corte."
            .Cells(34, 3) = "'" & "- Cualquier gasto adicional que no contemple la presente liquidación, será informado al cliente y "
            .Cells(35, 3) = "'" & "  tendrá que ser cancelado antes de realizar el abono que consigna la presente liquidación."
            .Cells(36, 3) = "'" & "- Realizar el depósito en la cuenta Nº 0011-0369-02-00090532 del BBVA Banco Continental."
            .Cells(37, 3) = "'" & "- Realizar el depósito en la cuenta Nº 011-369-000200090532-69 del BBVA Banco Continental desde otro Banco."
            .Range(.Cells(33, 3), .Cells(37, 3)).Font.Size = 10
            .Range(.Cells(33, 3), .Cells(37, 3)).Font.Bold = True
         End If
      Else
         r_obj_Excel.Selection.NumberFormat = "[$$]#,##0.00;[Red][$$]#,##0.00"
         If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) = 2 Then
            .Cells(34, 3) = "'" & "- Esta liquidación es válida solo a la fecha de corte."
            .Cells(35, 3) = "'" & "- Cualquier gasto adicional que no contemple la presente liquidación, será informado al cliente y "
            .Cells(36, 3) = "'" & "  tendrá que ser cancelado antes de realizar el abono que consigna la presente liquidación."
            .Cells(37, 3) = "'" & "- Realizar el depósito en la cuenta Nº 0011-0369-02-00090540 del BBVA Banco Continental."
            .Cells(38, 3) = "'" & "- Realizar el depósito en la cuenta Nº 011-369-000200090540-62 del BBVA Banco Continental desde otro Banco."
            .Range(.Cells(34, 3), .Cells(38, 3)).Font.Size = 10
            .Range(.Cells(34, 3), .Cells(38, 3)).Font.Bold = True
         Else
            .Cells(33, 3) = "'" & "- Esta liquidación es válida solo a la fecha de corte."
            .Cells(34, 3) = "'" & "- Cualquier gasto adicional que no contemple la presente liquidación, será informado al cliente y "
            .Cells(35, 3) = "'" & "  tendrá que ser cancelado antes de realizar el abono que consigna la presente liquidación."
            .Cells(36, 3) = "'" & "- Realizar el depósito en la cuenta Nº 0011-0369-02-00090540 del BBVA Banco Continental."
            .Cells(37, 3) = "'" & "- Realizar el depósito en la cuenta Nº 011-369-000200090540-62 del BBVA Banco Continental desde otro Banco."
            .Range(.Cells(33, 3), .Cells(37, 3)).Font.Size = 10
            .Range(.Cells(33, 3), .Cells(37, 3)).Font.Bold = True
         End If
      End If
   End With
      
   p_rut_Guardo = ""
   If p_flg_guardar = True Then
      p_rut_Guardo = Format(date, "yyyymmdd") & "_PPG_" & moddat_g_str_NumOpe & ".XLSX"
      r_obj_Excel.ActiveWorkbook.SaveAs (g_str_RutLog & "\" & p_rut_Guardo)
      
      r_obj_Excel.Application.Quit
      Set r_obj_Excel = Nothing
   Else
      r_obj_Excel.Visible = True
      Set r_obj_Excel = Nothing
   End If
   
End Sub

'*****************************************************
'* validar que no exista otro registro igual
Private Function fs_Validar_NumOpe() As Integer
   fs_Validar_NumOpe = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT PPGCAB_NUMOPE "
   g_str_Parame = g_str_Parame & "  FROM CRE_PPGCAB "
   g_str_Parame = g_str_Parame & " WHERE PPGCAB_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND PPGCAB_FECPPG = '" & ipp_FecPre.Year & IIf(Len(Trim(ipp_FecPre.Month)) = 1, 0 & ipp_FecPre.Month, ipp_FecPre.Month) & IIf(Len(Trim(ipp_FecPre.Day)) = 1, 0 & ipp_FecPre.Day, ipp_FecPre.Day) & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      fs_Validar_NumOpe = 1
      Exit Function
   End If
   
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      fs_Validar_NumOpe = 1
   End If
   
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
End Function

'**************************************************
'* graba datos en la tabla de cabecera de prepagos
Private Function fs_usp_cre_ppgcab() As Integer
   fs_usp_cre_ppgcab = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "usp_cre_ppgcab ( "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
   g_str_Parame = g_str_Parame & ipp_FecPre.Year & IIf(Len(Trim(ipp_FecPre.Month)) = 1, 0 & ipp_FecPre.Month, ipp_FecPre.Month) & IIf(Len(Trim(ipp_FecPre.Day)) = 1, 0 & ipp_FecPre.Day, ipp_FecPre.Day) & ", "
   g_str_Parame = g_str_Parame & 1 & ", "
   g_str_Parame = g_str_Parame & "'" & cmb_TipPre.ItemData(cmb_TipPre.ListIndex) & "', "
   g_str_Parame = g_str_Parame & CDbl(pnl_Val_AsgInm.Caption) & ", "
   g_str_Parame = g_str_Parame & CDbl(txt_Mto_Deposito.Text) & ", "
   g_str_Parame = g_str_Parame & CDbl(pnl_SaldoTNC1.Caption) & ", "
   g_str_Parame = g_str_Parame & IIf((InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0), 0, CDbl(pnl_SaldoTC1.Caption)) & ", "
   
   'Fecha del Ultimo Pago Realizado del TNC
   g_str_Parame = g_str_Parame & Year(CDate(pnl_UltPagTNC.Caption)) & IIf(Len(Trim(Month(CDate(pnl_UltPagTNC.Caption)))) = 1, 0 & Month(CDate(pnl_UltPagTNC.Caption)), Month(CDate(pnl_UltPagTNC.Caption))) & IIf(Len(Trim(Day(CDate(pnl_UltPagTNC.Caption)))) = 1, 0 & Day(CDate(pnl_UltPagTNC.Caption)), Day(CDate(pnl_UltPagTNC.Caption))) & ", "
   
   'Fecha del Ultimo Pago Realizado del TC
   If InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then
      g_str_Parame = g_str_Parame & 0 & ", "
   Else
      g_str_Parame = g_str_Parame & Year(CDate(pnl_UltPagTC.Caption)) & IIf(Len(Trim(Month(CDate(pnl_UltPagTC.Caption)))) = 1, 0 & Month(CDate(pnl_UltPagTC.Caption)), Month(CDate(pnl_UltPagTC.Caption))) & IIf(Len(Trim(Day(CDate(pnl_UltPagTC.Caption)))) = 1, 0 & Day(CDate(pnl_UltPagTC.Caption)), Day(CDate(pnl_UltPagTC.Caption))) & ", "
   End If
   
   g_str_Parame = g_str_Parame & CInt(pnl_DiasTNC.Caption) & ", "
   g_str_Parame = g_str_Parame & IIf((InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0), 0, CInt(pnl_DiasTC.Caption)) & ", "
   g_str_Parame = g_str_Parame & CDbl(txt_InteresTNC.Text) & ", "
   g_str_Parame = g_str_Parame & IIf((InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0), 0, CDbl(txt_InteresTC.Text)) & ", "
   g_str_Parame = g_str_Parame & CDbl(txt_SegDes.Text) & ", "
   g_str_Parame = g_str_Parame & CDbl(txt_SegInm.Text) & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & CDbl(txt_MontoITF.Text) & ", "
   g_str_Parame = g_str_Parame & CInt(Trim(pnl_CuoPen.Caption)) & ", "
   g_str_Parame = g_str_Parame & CInt(moddat_g_int_TotCuo - CInt(Trim(pnl_CuoPen.Caption))) & ", "
   
   If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) = 1 Then
      g_str_Parame = g_str_Parame & 0 & ", "
   Else
      g_str_Parame = g_str_Parame & cmb_RedPlz.ItemData(cmb_RedPlz.ListIndex) & ", "
   End If
   g_str_Parame = g_str_Parame & CDbl(pnl_MtoApl_Fin.Caption) & ", " 'CDbl(pnl_MtoApl.Caption) & ", "
   g_str_Parame = g_str_Parame & CDbl(txt_AplTNC.Text) & ", "
   g_str_Parame = g_str_Parame & CDbl(txt_ApliTC.Text) & ", "
   g_str_Parame = g_str_Parame & CDbl(pnl_NuevaCuota.Caption) & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
    g_str_Parame = g_str_Parame & CStr(cmb_MotPpg.ItemData(cmb_MotPpg.ListIndex)) & ", " 'g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & "'" & Trim(txt_ObsPpg.Text) & "', " 'g_str_Parame = g_str_Parame & "' ',"
   g_str_Parame = g_str_Parame & CDbl(Trim(pnl_CapPbp.Caption)) & ", "
   g_str_Parame = g_str_Parame & CDbl(Trim(pnl_IntPbp.Caption)) & ", "
   g_str_Parame = g_str_Parame & "0, "
   g_str_Parame = g_str_Parame & "0, "
   
   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                  'Código Usuario
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                  'Nombre Terminal
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                   'Nombre Ejecutable
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                  'Código Sucursal
   g_str_Parame = g_str_Parame & "1 ) "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      fs_usp_cre_ppgcab = 1
   End If
End Function

'*************************************************************
'* graba datos en la tabla de detalle de prepagos (TNC actual)
Private Function fs_usp_cre_ppgdet() As Integer
   fs_usp_cre_ppgdet = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "usp_cre_ppgdet ( "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
   g_str_Parame = g_str_Parame & "'" & ipp_FecPre.Year & IIf(Len(Trim(ipp_FecPre.Month)) = 1, 0 & ipp_FecPre.Month, ipp_FecPre.Month) & IIf(Len(Trim(ipp_FecPre.Day)) = 1, 0 & ipp_FecPre.Day, ipp_FecPre.Day) & "', "
   
   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                  'Código Usuario
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                  'Nombre Terminal
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                   'Nombre Ejecutable
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                  'Código Sucursal
   g_str_Parame = g_str_Parame & "1 ) "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      fs_usp_cre_ppgdet = 1
   End If
End Function

'**************************************************
'* graba datos en la tabla de cabecera de prepagos - solicitud
Private Function fs_usp_cre_ppgsol() As Integer
   fs_usp_cre_ppgsol = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "usp_cre_ppgsol ( "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
   g_str_Parame = g_str_Parame & ipp_FecPre.Year & IIf(Len(Trim(ipp_FecPre.Month)) = 1, 0 & ipp_FecPre.Month, ipp_FecPre.Month) & IIf(Len(Trim(ipp_FecPre.Day)) = 1, 0 & ipp_FecPre.Day, ipp_FecPre.Day) & ", "
   g_str_Parame = g_str_Parame & 1 & ", "
   g_str_Parame = g_str_Parame & "'" & cmb_TipPre.ItemData(cmb_TipPre.ListIndex) & "', "
   g_str_Parame = g_str_Parame & CDbl(pnl_Val_AsgInm.Caption) & ", "
   g_str_Parame = g_str_Parame & CDbl(txt_Mto_Deposito.Text) & ", "
   g_str_Parame = g_str_Parame & CDbl(pnl_SaldoTNC1.Caption) & ", "
   g_str_Parame = g_str_Parame & IIf((InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0), 0, CDbl(pnl_SaldoTC1.Caption)) & ", "
   
   'Fecha del Ultimo Pago Realizado del TNC
   g_str_Parame = g_str_Parame & Year(CDate(pnl_UltPagTNC.Caption)) & IIf(Len(Trim(Month(CDate(pnl_UltPagTNC.Caption)))) = 1, 0 & Month(CDate(pnl_UltPagTNC.Caption)), Month(CDate(pnl_UltPagTNC.Caption))) & IIf(Len(Trim(Day(CDate(pnl_UltPagTNC.Caption)))) = 1, 0 & Day(CDate(pnl_UltPagTNC.Caption)), Day(CDate(pnl_UltPagTNC.Caption))) & ", "
   
   'Fecha del Ultimo Pago Realizado del TC
   If InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then
      g_str_Parame = g_str_Parame & 0 & ", "
   Else
      g_str_Parame = g_str_Parame & Year(CDate(pnl_UltPagTC.Caption)) & IIf(Len(Trim(Month(CDate(pnl_UltPagTC.Caption)))) = 1, 0 & Month(CDate(pnl_UltPagTC.Caption)), Month(CDate(pnl_UltPagTC.Caption))) & IIf(Len(Trim(Day(CDate(pnl_UltPagTC.Caption)))) = 1, 0 & Day(CDate(pnl_UltPagTC.Caption)), Day(CDate(pnl_UltPagTC.Caption))) & ", "
   End If
   
   g_str_Parame = g_str_Parame & CInt(pnl_DiasTNC.Caption) & ", "
   g_str_Parame = g_str_Parame & IIf((InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0), 0, CInt(pnl_DiasTC.Caption)) & ", "
   g_str_Parame = g_str_Parame & CDbl(txt_InteresTNC.Text) & ", "
   g_str_Parame = g_str_Parame & IIf((InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0), 0, CDbl(txt_InteresTC.Text)) & ", "
   g_str_Parame = g_str_Parame & CDbl(txt_SegDes.Text) & ", "
   g_str_Parame = g_str_Parame & CDbl(txt_SegInm.Text) & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & CDbl(txt_MontoITF.Text) & ", "
   g_str_Parame = g_str_Parame & CInt(Trim(pnl_CuoPen.Caption)) & ", "
   g_str_Parame = g_str_Parame & CInt(moddat_g_int_TotCuo - CInt(Trim(pnl_CuoPen.Caption))) & ", "
   
   If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) = 1 Then
      g_str_Parame = g_str_Parame & 0 & ", "
   Else
      g_str_Parame = g_str_Parame & cmb_RedPlz.ItemData(cmb_RedPlz.ListIndex) & ", "
   End If
   g_str_Parame = g_str_Parame & CDbl(pnl_MtoApl.Caption) & ", "
   g_str_Parame = g_str_Parame & CDbl(txt_AplTNC.Text) & ", "
   g_str_Parame = g_str_Parame & CDbl(txt_ApliTC.Text) & ", "
   g_str_Parame = g_str_Parame & CDbl(pnl_NuevaCuota.Caption) & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   If cmb_MotPpg.ListIndex = -1 Then
      g_str_Parame = g_str_Parame & 0 & ", "
   Else
      g_str_Parame = g_str_Parame & CStr(cmb_MotPpg.ItemData(cmb_MotPpg.ListIndex)) & ", "
   End If
   g_str_Parame = g_str_Parame & "'" & Trim(txt_ObsPpg.Text) & "', " 'g_str_Parame = g_str_Parame & "' ',"
   g_str_Parame = g_str_Parame & CDbl(Trim(pnl_CapPbp.Caption)) & ", "
   g_str_Parame = g_str_Parame & CDbl(Trim(pnl_IntPbp.Caption)) & ", "
   
   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                  'Código Usuario
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                  'Nombre Terminal
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                   'Nombre Ejecutable
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                  'Código Sucursal
   g_str_Parame = g_str_Parame & "1 ) "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      fs_usp_cre_ppgsol = 1
   End If
 
End Function

'********************************************************
'* actualiza las nuevas cuotas en la tabla TABLA CRE_HIPCUO
Private Function fs_ppgpar_hipcuo() As Integer
   Dim r_int_NroFil   As Integer
   fs_ppgpar_hipcuo = 0
   
   For r_int_NroFil = 0 To grd_CliNCo_Listad.Rows - 1
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "usp_ppgpar_cre_hipcuo ( "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
      g_str_Parame = g_str_Parame & CInt(grd_CliNCo_Listad.TextMatrix(r_int_NroFil, 0)) & ", "
      g_str_Parame = g_str_Parame & CDbl(grd_CliNCo_Listad.TextMatrix(r_int_NroFil, 2)) & ", "
      g_str_Parame = g_str_Parame & CDbl(grd_CliNCo_Listad.TextMatrix(r_int_NroFil, 3)) & ", "
      g_str_Parame = g_str_Parame & CDbl(grd_CliNCo_Listad.TextMatrix(r_int_NroFil, 4)) & ", "
      g_str_Parame = g_str_Parame & CDbl(grd_CliNCo_Listad.TextMatrix(r_int_NroFil, 5)) & ", "
      g_str_Parame = g_str_Parame & CDbl(grd_CliNCo_Listad.TextMatrix(r_int_NroFil, 6)) & ", "
      g_str_Parame = g_str_Parame & CDbl(grd_CliNCo_Listad.TextMatrix(r_int_NroFil, 8)) & ", "
   
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "               'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "               'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "               'Código Sucursal
      g_str_Parame = g_str_Parame & "1 ) "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         fs_ppgpar_hipcuo = fs_ppgpar_hipcuo + 1
      End If
   Next
   
   If fs_ppgpar_hipcuo > 0 Then
      MsgBox "No se pudo completar el procedimiento 'usp_ppgpar_cre_hipcuo'.", vbCritical, modgen_g_str_NomPlt
   Else
      If CInt(cmb_TipPre.ListIndex) = 1 Then
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "usp_ppgpar_delcuo ( "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
         g_str_Parame = g_str_Parame & CInt(grd_CliNCo_Listad.TextMatrix(r_int_NroFil - 1, 0)) & " ) "
   
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
            fs_ppgpar_hipcuo = fs_ppgpar_hipcuo + 1
         End If
   
         If fs_ppgpar_hipcuo > 0 Then
            MsgBox "No se pudo completar el procedimiento 'usp_ppgpar_delcuo'.", vbCritical, modgen_g_str_NomPlt
         End If
      End If
   End If
End Function

'* actualiza el estado en la tabla de cabecera de prepagos
Private Function fs_usp_actualiza_cre_ppgcab() As Integer
   fs_usp_actualiza_cre_ppgcab = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "USP_ACTUALIZA_CRE_PPGCAB ( "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
   g_str_Parame = g_str_Parame & "" & ipp_FecPre.Year & IIf(Len(Trim(ipp_FecPre.Month)) = 1, 0 & ipp_FecPre.Month, ipp_FecPre.Month) & IIf(Len(Trim(ipp_FecPre.Day)) = 1, 0 & ipp_FecPre.Day, ipp_FecPre.Day) & " , 1, 0, 0, 0, 0) "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      fs_usp_actualiza_cre_ppgcab = 1
   End If
End Function

'*************************************************
'* Actualizar maestro de credito hipotecarios
Private Function fs_update_hipmae() As Integer
   fs_update_hipmae = 0
   
   If cmb_TipPre.ListIndex = 0 Then
      '* REDUCCION DE MONTO
      If InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then
         l_dbl_PagCap = l_dbl_PagCap + CDbl(pnl_MtoApl.Caption)
      Else
         l_dbl_SalCon = CDbl(pnl_NuevoSaldoTC.Caption)
         l_dbl_PagCap = l_dbl_PagCap + CDbl(pnl_NuevoSaldoTNC.Caption)
      End If
   Else
      '* REDUCCION DE PLAZO
      l_int_PlaAno = l_int_PlaAno - CInt(cmb_RedPlz.ItemData(cmb_RedPlz.ListIndex))
      l_int_NCuota = CInt(l_int_PlaAno * 12)
      moddat_g_int_CuoPen = CInt(l_int_NCuota - l_int_PagCuo)
      
      If l_int_CodPrd = 2 Or l_int_CodPrd = 11 Then
         l_dbl_PagCap = l_dbl_PagCap + CDbl(pnl_MtoApl.Caption)
      Else
         l_dbl_SalCon = CDbl(pnl_NuevoSaldoTC.Caption)
         l_dbl_PagCap = l_dbl_PagCap + CDbl(pnl_NuevoSaldoTNC.Caption)
      End If
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "usp_ppgpar_cre_hipmae ( "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
   g_str_Parame = g_str_Parame & ipp_FecPre.Year & IIf(Len(Trim(ipp_FecPre.Month)) = 1, 0 & ipp_FecPre.Month, ipp_FecPre.Month) & IIf(Len(Trim(ipp_FecPre.Day)) = 1, 0 & ipp_FecPre.Day, ipp_FecPre.Day) & ", "
   g_str_Parame = g_str_Parame & "'" & l_int_PlaAno & "', "
   g_str_Parame = g_str_Parame & "'" & Right(Format(CStr(grd_CliNCo_Listad.TextMatrix(grd_CliNCo_Listad.Rows - 1, 1)), "yyyymmdd"), 2) & "', "  'Left(l_str_UltVct, 2) & "', "
   g_str_Parame = g_str_Parame & "'" & l_int_NCuota & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_g_int_CuoPen & "', "
   g_str_Parame = g_str_Parame & "'" & Format(CStr(grd_CliNCo_Listad.TextMatrix(grd_CliNCo_Listad.Rows - 1, 1)), "yyyymmdd") & "', "            'l_str_UltVct & "', "
   g_str_Parame = g_str_Parame & CDbl(pnl_NuevaCuota.Caption) & ", "
   g_str_Parame = g_str_Parame & CDbl(pnl_NuevoSaldoTNC.Caption) & ", "
   g_str_Parame = g_str_Parame & CDbl(pnl_NuevoSaldoTNC.Caption) & ", "
   g_str_Parame = g_str_Parame & l_dbl_SalCon & ", "
   g_str_Parame = g_str_Parame & l_dbl_PagCap & ", "
   g_str_Parame = g_str_Parame & l_dbl_PagInt + CInt(pnl_DiasTNC.Caption) & ", "
   g_str_Parame = g_str_Parame & "'" & l_int_PagCuo & "', "
     
   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                  'Código Usuario
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                  'Nombre Terminal
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                   'Nombre Ejecutable
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                  'Código Sucursal
   g_str_Parame = g_str_Parame & "1 ) "
     
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 2) Then
      fs_update_hipmae = 1
      MsgBox "No se pudo completar el procedimiento usp_ppgpar_cre_hipmae.", vbCritical, modgen_g_str_NomPlt
   End If
   
End Function

Private Sub fs_Exportar()
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_NroFil     As Integer
   Dim r_int_nroaux     As Integer
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   r_int_NroFil = 1
   
   With r_obj_Excel.ActiveSheet
      .Cells(r_int_NroFil, 1) = "CUOTA":            .Columns("A").ColumnWidth = 8
      .Cells(r_int_NroFil, 2) = "FECHA VCTO.":      .Columns("B").ColumnWidth = 16
      .Cells(r_int_NroFil, 3) = "CAPITAL":          .Columns("C").ColumnWidth = 18
      .Cells(r_int_NroFil, 4) = "INTERES":          .Columns("D").ColumnWidth = 18
      .Cells(r_int_NroFil, 5) = "SEG. PRESTAMO":    .Columns("E").ColumnWidth = 18
      .Cells(r_int_NroFil, 6) = "SEG. VIVIENDA":    .Columns("F").ColumnWidth = 18
      .Cells(r_int_NroFil, 7) = "PORTES":           .Columns("G").ColumnWidth = 18
      .Cells(r_int_NroFil, 8) = "TOTAL CUOTA":      .Columns("H").ColumnWidth = 18
      .Cells(r_int_NroFil, 9) = "SALDO CAPITAL":    .Columns("I").ColumnWidth = 18
      
      .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 2).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 3).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 4).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 5).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 6).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 7).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 8).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 9).HorizontalAlignment = xlHAlignCenter
      
      .Cells(r_int_NroFil, 1).Font.Bold = True
      .Cells(r_int_NroFil, 2).Font.Bold = True
      .Cells(r_int_NroFil, 3).Font.Bold = True
      .Cells(r_int_NroFil, 4).Font.Bold = True
      .Cells(r_int_NroFil, 5).Font.Bold = True
      .Cells(r_int_NroFil, 6).Font.Bold = True
      .Cells(r_int_NroFil, 7).Font.Bold = True
      .Cells(r_int_NroFil, 8).Font.Bold = True
      .Cells(r_int_NroFil, 9).Font.Bold = True
 
      r_int_NroFil = r_int_NroFil + 1
      
      For r_int_nroaux = 0 To grd_CliNCo_Listad.Rows - 1
         .Cells(r_int_NroFil, 1) = grd_CliNCo_Listad.TextMatrix(r_int_nroaux, 0)
         .Cells(r_int_NroFil, 2) = "'" & grd_CliNCo_Listad.TextMatrix(r_int_nroaux, 1)
         .Cells(r_int_NroFil, 3) = grd_CliNCo_Listad.TextMatrix(r_int_nroaux, 2)
         .Cells(r_int_NroFil, 4) = grd_CliNCo_Listad.TextMatrix(r_int_nroaux, 3)
         .Cells(r_int_NroFil, 5) = grd_CliNCo_Listad.TextMatrix(r_int_nroaux, 4)
         .Cells(r_int_NroFil, 6) = grd_CliNCo_Listad.TextMatrix(r_int_nroaux, 5)
         .Cells(r_int_NroFil, 7) = grd_CliNCo_Listad.TextMatrix(r_int_nroaux, 6)
         .Cells(r_int_NroFil, 8) = grd_CliNCo_Listad.TextMatrix(r_int_nroaux, 7)
         .Cells(r_int_NroFil, 9) = grd_CliNCo_Listad.TextMatrix(r_int_nroaux, 8)
         r_int_NroFil = r_int_NroFil + 1
      Next
      
      .Columns("C").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Columns("D").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Columns("E").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Columns("F").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Columns("G").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Columns("H").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Columns("I").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      r_int_NroFil = r_int_NroFil + 2
      
      .Range(.Cells(r_int_NroFil, 11), .Cells(1, 23)).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 11), .Cells(1, 23)).HorizontalAlignment = xlHAlignCenter
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub ipp_FecPre_LostFocus()
   If IsDate(pnl_UltPagTNC.Caption) Then
      pnl_DiasTNC.Caption = DateDiff("d", pnl_UltPagTNC.Caption, ipp_FecPre.Text) & " "
      Call fs_Cal_Interes
   End If
   If IsDate(pnl_UltPagTC.Caption) Then
      pnl_DiasTC.Caption = DateDiff("d", pnl_UltPagTC.Caption, ipp_FecPre.Text) & " "
      Call fs_Cal_Interes
   End If
    If IsDate(pnl_UltPagTNC.Caption) And IsDate(pnl_UltPagTC.Caption) Then
      Call fs_Buscar_Cuotas_Vencidas
      Call fs_Buscar
      Call fs_Cal_Prpago
      Call fs_Cal_Prctaj
      Call fs_Cal_MtoItf
      Call fs_Limpia_Cronog
   End If
End Sub

Private Sub ipp_FecPre_Change()
   If IsDate(pnl_UltPagTNC.Caption) Then
      pnl_DiasTNC.Caption = DateDiff("d", pnl_UltPagTNC.Caption, ipp_FecPre.Text) & " "
      Call fs_Cal_Interes
   End If
   If IsDate(pnl_UltPagTC.Caption) Then
      pnl_DiasTC.Caption = DateDiff("d", pnl_UltPagTC.Caption, ipp_FecPre.Text) & " "
      Call fs_Cal_Interes
   End If
    If IsDate(pnl_UltPagTNC.Caption) And IsDate(pnl_UltPagTC.Caption) Then
      Call fs_Buscar_Cuotas_Vencidas
      Call fs_Buscar
      Call fs_Cal_Prpago
      Call fs_Cal_Prctaj
      Call fs_Cal_MtoItf
      Call fs_Limpia_Cronog
   End If
End Sub

Private Sub ipp_FecPre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call fs_Cal_Interes
      Call gs_SetFocus(cmb_TipPre)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub cmb_TipPre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_TipPre.ListIndex > -1 Then
         If cmb_RedPlz.Enabled Then
            Call gs_SetFocus(cmb_RedPlz)
         Else
            Call gs_SetFocus(txt_Mto_Deposito)
         End If
      End If
   End If
End Sub

Private Sub cmb_TipPre_Change()
   If cmb_TipPre.ListIndex > -1 Then
      Select Case cmb_TipPre.ItemData(cmb_TipPre.ListIndex)
         Case 1: cmb_RedPlz.Enabled = False
         Case 2: cmb_RedPlz.Enabled = True
      End Select
      cmb_RedPlz.ListIndex = -1
   End If
End Sub

Private Sub cmb_TipPre_LostFocus()
   If cmb_TipPre.ListIndex > -1 Then
      Select Case cmb_TipPre.ItemData(cmb_TipPre.ListIndex)
         Case 1: cmb_RedPlz.Enabled = False
         Case 2: cmb_RedPlz.Enabled = True
      End Select
      cmb_RedPlz.ListIndex = -1
      If cmb_RedPlz.Enabled Then
         Call gs_SetFocus(cmb_RedPlz)
      Else
         Call gs_SetFocus(txt_Mto_Deposito)
      End If
   End If
End Sub

Private Sub cmb_RedPlz_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_RedPlz.ListIndex > -1 Then
         Call gs_SetFocus(txt_Mto_Deposito)
      End If
   End If
End Sub

Private Sub txt_Mto_Deposito_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_Mto_Deposito_LostFocus
      Call gs_SetFocus(txt_InteresTNC)
   End If
End Sub

Private Sub txt_Mto_Deposito_LostFocus()
   Call fs_Cal_Prpago
   Call fs_Cal_Prctaj
   Call fs_Cal_MtoItf
   Call fs_Limpia_Cronog
End Sub

Private Sub txt_InteresTNC_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If (InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0) Then
         Call gs_SetFocus(txt_SegDes)
      Else
         Call gs_SetFocus(txt_InteresTC)
      End If
   End If
End Sub

Private Sub txt_InteresTNC_LostFocus()
   Call fs_Cal_Prpago
   Call fs_Cal_Prctaj
End Sub

Private Sub txt_InteresTC_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_SegDes)
   End If
End Sub

Private Sub txt_InteresTC_LostFocus()
   Call fs_Cal_Prpago
   Call fs_Cal_Prctaj
End Sub

Private Sub txt_ObsPpg_KeyPress(KeyAscii As Integer)
   If Not KeyAscii = 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub txt_SegDes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_SegInm)
   End If
End Sub

Private Sub txt_SegDes_LostFocus()
   Call fs_Cal_Prpago
   Call fs_Cal_Prctaj
End Sub

Private Sub txt_SegInm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_MontoITF)
   End If
End Sub

Private Sub txt_SegInm_LostFocus()
   Call fs_Cal_Prpago
   Call fs_Cal_Prctaj
End Sub

Private Sub txt_MontoITF_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_AplTNC)
   End If
End Sub

Private Sub txt_MontoITF_LostFocus()
   Call fs_Cal_Prpago
   Call fs_Cal_Prctaj
End Sub
 
Private Sub txt_AplTNC_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If (InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0) Then
         Call gs_SetFocus(cmd_Recalc)
      Else
         pnl_NuevoSaldoTNC.Caption = gf_FormatoNumero(l_dbl_SalNco - CDbl(txt_AplTNC.Text), 12, 2) & " "
         Call gs_SetFocus(txt_ApliTC)
      End If
   End If
End Sub
 
Private Sub txt_ApliTC_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      pnl_NuevoSaldoTC.Caption = gf_FormatoNumero(l_dbl_SalCon - CDbl(txt_ApliTC.Text), 12, 2) & " "
      Call gs_SetFocus(cmd_Recalc)
   End If
End Sub
Private Sub fs_Validar_Seguro()
Dim r_int_Contad        As Integer
Dim r_int_ConAux        As Integer
Dim r_dbl_Cuo_Capita    As Double
Dim r_dbl_Cuo_Intere    As Double
Dim r_dbl_Cuo_SegPre    As Double
Dim r_dbl_Cuo_SegViv    As Double
Dim r_dbl_Cuo_Portes    As Double
Dim r_dbl_Cuo_TotCuo    As Double
Dim r_dbl_TotSeg_SegPre    As Double
Dim r_dbl_TotSeg_SegViv    As Double
Dim r_dbl_Val_SegPre    As Double
Dim r_dbl_Val_SegViv    As Double

Dim r_dbl_TotSeg_TotCuo    As Double
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT HIPCUO_NUMOPE, HIPCUO_NUMCUO, HIPCUO_DESORG, HIPCUO_VIVORG FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & "  WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "    AND HIPCUO_TIPCRO = 1 "
   g_str_Parame = g_str_Parame & "    AND HIPCUO_NUMCUO > " & l_int_PagCuo & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   g_rst_Princi.MoveFirst
   
   ReDim arr_CroTNC(0)
   
   Do While Not g_rst_Princi.EOF
   
      ReDim Preserve arr_CroTNC(UBound(arr_CroTNC) + 1)
      arr_CroTNC(UBound(arr_CroTNC)).str_NumOpe = Trim(g_rst_Princi!HIPCUO_NUMOPE)
      arr_CroTNC(UBound(arr_CroTNC)).int_NumCuo = Trim(g_rst_Princi!HIPCUO_NUMCUO)
      arr_CroTNC(UBound(arr_CroTNC)).dbl_SegPre = Trim(g_rst_Princi!HIPCUO_DESORG)
      arr_CroTNC(UBound(arr_CroTNC)).dbl_SegViv = Trim(g_rst_Princi!HIPCUO_VIVORG)
      
      r_dbl_Val_SegPre = r_dbl_Val_SegPre + Trim(g_rst_Princi!HIPCUO_DESORG)
      r_dbl_Val_SegViv = r_dbl_Val_SegViv + Trim(g_rst_Princi!HIPCUO_VIVORG)
      
      g_rst_Princi.MoveNext
   Loop

   r_int_ConAux = 1
   If UBound(arr_CroTNC) > 0 Then
      
      With grd_CliNCo_Listad
         For r_int_Contad = 0 To .Rows - 1
         
            If Trim(arr_CroTNC(r_int_ConAux).int_NumCuo) = CInt(.TextMatrix(r_int_Contad, 0)) Then
            
               If Format(arr_CroTNC(r_int_ConAux).dbl_SegPre, "###,##0.00") = 0 Then
               
                  If Format(arr_CroTNC(r_int_ConAux).dbl_SegPre, "###,##0.00") = 0 Then
                     .TextMatrix(r_int_Contad, 4) = Format(arr_CroTNC(r_int_ConAux).dbl_SegPre, "###,##0.00")
                  End If
              
                  r_dbl_Cuo_Capita = CDbl(Format(.TextMatrix(r_int_Contad, 2), "###,##0.00"))
                  r_dbl_Cuo_Intere = CDbl(Format(.TextMatrix(r_int_Contad, 3), "###,##0.00"))
                  r_dbl_Cuo_SegPre = CDbl(Format(.TextMatrix(r_int_Contad, 4), "###,##0.00"))
                  r_dbl_Cuo_SegViv = CDbl(Format(.TextMatrix(r_int_Contad, 5), "###,##0.00"))
                  r_dbl_Cuo_Portes = CDbl(Format(.TextMatrix(r_int_Contad, 6), "###,##0.00"))
                  r_dbl_Cuo_TotCuo = CDbl(Format(r_dbl_Cuo_Capita + r_dbl_Cuo_Intere + r_dbl_Cuo_SegPre + r_dbl_Cuo_SegViv + r_dbl_Cuo_Portes, "###,##0.00"))
                  
                  .TextMatrix(r_int_Contad, 7) = Format(r_dbl_Cuo_TotCuo, "###,##0.00")
                  
               End If
               
               If Format(arr_CroTNC(r_int_ConAux).dbl_SegViv, "###,##0.00") = 0 Then
                  
                  If Format(arr_CroTNC(r_int_ConAux).dbl_SegViv, "###,##0.00") = 0 Then
                     .TextMatrix(r_int_Contad, 5) = Format(arr_CroTNC(r_int_ConAux).dbl_SegViv, "###,##0.00")
                  End If
                  
                  r_dbl_Cuo_Capita = CDbl(Format(.TextMatrix(r_int_Contad, 2), "###,##0.00"))
                  r_dbl_Cuo_Intere = CDbl(Format(.TextMatrix(r_int_Contad, 3), "###,##0.00"))
                  r_dbl_Cuo_SegPre = CDbl(Format(.TextMatrix(r_int_Contad, 4), "###,##0.00"))
                  r_dbl_Cuo_SegViv = CDbl(Format(.TextMatrix(r_int_Contad, 5), "###,##0.00"))
                  r_dbl_Cuo_Portes = CDbl(Format(.TextMatrix(r_int_Contad, 6), "###,##0.00"))
                  r_dbl_Cuo_TotCuo = CDbl(Format(r_dbl_Cuo_Capita + r_dbl_Cuo_Intere + r_dbl_Cuo_SegPre + r_dbl_Cuo_SegViv + r_dbl_Cuo_Portes, "###,##0.00"))
                  
                  .TextMatrix(r_int_Contad, 7) = Format(r_dbl_Cuo_TotCuo, "###,##0.00")
                  
               End If
            
            End If
            r_int_ConAux = r_int_ConAux + 1
         
         Next r_int_Contad
      End With
      
      If r_dbl_Cuo_TotCuo > 0 Then
      
         For r_int_Contad = 0 To grd_CliNCo_Listad.Rows - 1
             r_dbl_TotSeg_SegPre = r_dbl_TotSeg_SegPre + grd_CliNCo_Listad.TextMatrix(r_int_Contad, 4)
             r_dbl_TotSeg_SegViv = r_dbl_TotSeg_SegViv + grd_CliNCo_Listad.TextMatrix(r_int_Contad, 5)
             r_dbl_TotSeg_TotCuo = r_dbl_TotSeg_TotCuo + grd_CliNCo_Listad.TextMatrix(r_int_Contad, 7)
         Next r_int_Contad

         pnl_CliNCo_SegPre.Caption = Format(r_dbl_TotSeg_SegPre, "###,##0.00") & " "
         pnl_CliNCo_SegViv.Caption = Format(r_dbl_TotSeg_SegViv, "###,##0.00") & " "
         pnl_CliNCo_TotCuo.Caption = Format(r_dbl_TotSeg_TotCuo, "###,##0.00") & " "
      End If
   End If
End Sub

Private Sub fs_Validar_Seguro_ant()

Dim r_int_Contad        As Integer
Dim r_int_ConAux        As Integer
Dim r_dbl_Cuo_Capita    As Double
Dim r_dbl_Cuo_Intere    As Double
Dim r_dbl_Cuo_SegPre    As Double
Dim r_dbl_Cuo_SegViv    As Double
Dim r_dbl_Cuo_Portes    As Double
Dim r_dbl_Cuo_TotCuo    As Double
Dim r_dbl_Tot_Capita    As Double
Dim r_dbl_Tot_Intere    As Double
Dim r_dbl_Tot_SegPre    As Double
Dim r_dbl_Tot_SegViv    As Double
Dim r_dbl_Val_SegPre    As Double
Dim r_dbl_Val_SegViv    As Double
Dim r_dbl_Tot_Portes    As Double
Dim r_dbl_Tot_TotCuo    As Double
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT HIPCUO_NUMOPE, HIPCUO_NUMCUO, HIPCUO_DESORG, HIPCUO_VIVORG FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & "  WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "    AND HIPCUO_TIPCRO = 1 "
   g_str_Parame = g_str_Parame & "    AND HIPCUO_NUMCUO > " & l_int_PagCuo & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   g_rst_Princi.MoveFirst
   
   ReDim arr_CroTNC(0)
   
   Do While Not g_rst_Princi.EOF
   
      ReDim Preserve arr_CroTNC(UBound(arr_CroTNC) + 1)
      arr_CroTNC(UBound(arr_CroTNC)).str_NumOpe = Trim(g_rst_Princi!HIPCUO_NUMOPE)
      arr_CroTNC(UBound(arr_CroTNC)).int_NumCuo = Trim(g_rst_Princi!HIPCUO_NUMCUO)
      arr_CroTNC(UBound(arr_CroTNC)).dbl_SegPre = Trim(g_rst_Princi!HIPCUO_DESORG)
      arr_CroTNC(UBound(arr_CroTNC)).dbl_SegViv = Trim(g_rst_Princi!HIPCUO_VIVORG)
      
      r_dbl_Val_SegPre = r_dbl_Val_SegPre + Trim(g_rst_Princi!HIPCUO_DESORG)
      r_dbl_Val_SegViv = r_dbl_Val_SegViv + Trim(g_rst_Princi!HIPCUO_VIVORG)
      
      g_rst_Princi.MoveNext
   Loop

   r_int_ConAux = 1
   If UBound(arr_CroTNC) > 0 Then
      
      With grd_CliNCo_Listad
         For r_int_Contad = 0 To .Rows - 1
         
            If Trim(arr_CroTNC(r_int_ConAux).int_NumCuo) = CInt(.TextMatrix(r_int_Contad, 0)) Then
            
               If Format(arr_CroTNC(r_int_ConAux).dbl_SegPre, "###,##0.00") = 0 Or Format(arr_CroTNC(r_int_ConAux).dbl_SegViv, "###,##0.00") = 0 Then
               
                  If Format(arr_CroTNC(r_int_ConAux).dbl_SegPre, "###,##0.00") = 0 Then
                     .TextMatrix(r_int_Contad, 4) = Format(arr_CroTNC(r_int_ConAux).dbl_SegPre, "###,##0.00")
                  End If
                  
                  If Format(arr_CroTNC(r_int_ConAux).dbl_SegViv, "###,##0.00") = 0 Then
                     .TextMatrix(r_int_Contad, 5) = Format(arr_CroTNC(r_int_ConAux).dbl_SegViv, "###,##0.00")
                  End If
                  
                  r_dbl_Cuo_Capita = CDbl(Format(.TextMatrix(r_int_Contad, 2), "###,##0.00"))
                  r_dbl_Cuo_Intere = CDbl(Format(.TextMatrix(r_int_Contad, 3), "###,##0.00"))
                  r_dbl_Cuo_SegPre = CDbl(Format(.TextMatrix(r_int_Contad, 4), "###,##0.00"))
                  r_dbl_Cuo_SegViv = CDbl(Format(.TextMatrix(r_int_Contad, 5), "###,##0.00"))
                  r_dbl_Cuo_Portes = CDbl(Format(.TextMatrix(r_int_Contad, 6), "###,##0.00"))
                  r_dbl_Cuo_TotCuo = CDbl(Format(r_dbl_Cuo_Capita + r_dbl_Cuo_Intere + r_dbl_Cuo_SegPre + r_dbl_Cuo_SegViv + r_dbl_Cuo_Portes, "###,##0.00"))
                  
                  .TextMatrix(r_int_Contad, 7) = Format(r_dbl_Cuo_TotCuo, "###,##0.00")
                  
                  r_dbl_Tot_Capita = r_dbl_Tot_Capita + r_dbl_Cuo_Capita
                  r_dbl_Tot_Intere = r_dbl_Tot_Intere + r_dbl_Cuo_Intere
                  r_dbl_Tot_SegPre = r_dbl_Tot_SegPre + r_dbl_Cuo_SegPre
                  r_dbl_Tot_SegViv = r_dbl_Tot_SegViv + r_dbl_Cuo_SegViv
                  r_dbl_Tot_Portes = r_dbl_Tot_Portes + r_dbl_Cuo_Portes
                  
               End If
            
            End If
            r_int_ConAux = r_int_ConAux + 1
         
         Next r_int_Contad
      End With
      
      If r_dbl_Cuo_TotCuo > 0 Then
         pnl_CliNCo_Capita.Caption = Format(r_dbl_Tot_Capita, "###,##0.00") & " "
         pnl_CliNCo_Intere.Caption = Format(r_dbl_Tot_Intere, "###,##0.00") & " "
         pnl_CliNCo_SegPre.Caption = Format(r_dbl_Tot_SegPre, "###,##0.00") & " "
         pnl_CliNCo_SegViv.Caption = Format(r_dbl_Tot_SegViv, "###,##0.00") & " "
         pnl_CliNCo_OtrCar.Caption = Format(r_dbl_Tot_Portes, "###,##0.00") & " "
         pnl_CliNCo_TotCuo.Caption = Format(r_dbl_Tot_TotCuo, "###,##0.00") & " "
      End If
   End If
End Sub

