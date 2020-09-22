VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{0229E69D-587E-485B-A9A8-795AB1BFE516}#1.0#0"; "Phantom.ocx"
Begin VB.Form frmMain 
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   12495
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList i16x16 
      Left            =   5400
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":370C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":411E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":44B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4852
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4BEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F86
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5998
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":63AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6DBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":77CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":81E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8BF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9604
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9BA0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin prjStudentIMS.ucStatusbar ucStatus 
      Height          =   375
      Left            =   0
      Top             =   8100
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   661
   End
   Begin MSComctlLib.ImageList imlPanel 
      Left            =   4680
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   46
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A13C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BACE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C7AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E13C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FACE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11460
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12DF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13ACC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":147A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15082
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15D5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16A3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1731E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17FFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":188D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":195B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AF46
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C8DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D1B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DA90
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E36A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EC44
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F51E
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1FAB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1FDD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":200EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":209C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":212A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21B7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22854
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22B6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22FC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23E12
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24264
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24B3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25418
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2AC0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B4E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B7FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C4D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2CDB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D68C
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DF66
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E840
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F11A
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F9F4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin prjStudentIMS.ucHorizontal3DLine ucToolbar 
      Height          =   30
      Left            =   -30
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   360
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   53
   End
   Begin prjStudentIMS.ucVertical3DLine ucDate_Sep 
      Height          =   255
      Left            =   11520
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
   End
   Begin MSComctlLib.TreeView tvwReport 
      Height          =   1455
      Left            =   6600
      TabIndex        =   25
      Top             =   2640
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2566
      _Version        =   393217
      Indentation     =   265
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "i16x16"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjStudentIMS.ucVertical3DLine ucVertical3DLine8 
      Height          =   255
      Left            =   9390
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
   End
   Begin prjStudentIMS.ucVertical3DLine ucVertical3DLine7 
      Height          =   255
      Left            =   8220
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
   End
   Begin PhantomPanel.PanelControl pMenu 
      Height          =   3225
      Left            =   180
      TabIndex        =   16
      Top             =   1230
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   5689
      PanelIconPictureSize=   32
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PanelStyle      =   2
      ToolTipStyle    =   1
      SelectedItemColor=   -2147483630
   End
   Begin prjStudentIMS.ucVertical3DLine ucVertical3DLine6 
      Height          =   255
      Left            =   7020
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
   End
   Begin prjStudentIMS.ucVertical3DLine ucVertical3DLine5 
      Height          =   255
      Left            =   5850
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
   End
   Begin prjStudentIMS.ucVertical3DLine ucVertical3DLine4 
      Height          =   255
      Left            =   4680
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
   End
   Begin prjStudentIMS.ucVertical3DLine ucVertical3DLine3 
      Height          =   255
      Left            =   3510
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
   End
   Begin prjStudentIMS.ucVertical3DLine ucVertical3DLine2 
      Height          =   255
      Left            =   2340
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
   End
   Begin prjStudentIMS.ucVertical3DLine ucVertical3DLine1 
      Height          =   255
      Left            =   1170
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
   End
   Begin prjStudentIMS.lvButtons_H cmdTools 
      Height          =   375
      Left            =   7050
      TabIndex        =   6
      Top             =   0
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Tools"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin prjStudentIMS.lvButtons_H cmdRefresh 
      Height          =   375
      Left            =   3510
      TabIndex        =   3
      Top             =   0
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Refresh"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin prjStudentIMS.lvButtons_H cmdDelete 
      Height          =   375
      Left            =   2340
      TabIndex        =   2
      Top             =   0
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Delete"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin prjStudentIMS.lvButtons_H cmdEdit 
      Height          =   375
      Left            =   1170
      TabIndex        =   1
      Top             =   0
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Edit"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin prjStudentIMS.lvButtons_H cmdAdd 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Add"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.PictureBox picData 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   4620
      ScaleHeight     =   2415
      ScaleWidth      =   4875
      TabIndex        =   9
      Top             =   1860
      Width           =   4875
      Begin MSComctlLib.ListView lvListView 
         Height          =   1935
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   3413
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         SmallIcons      =   "i16x16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin prjStudentIMS.lvButtons_H cmdNavigate 
         Height          =   315
         Index           =   0
         Left            =   0
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "First Record"
         Top             =   2100
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         CapAlign        =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Focus           =   0   'False
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin prjStudentIMS.lvButtons_H cmdNavigate 
         Height          =   315
         Index           =   1
         Left            =   360
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Previous Record"
         Top             =   2100
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         CapAlign        =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Focus           =   0   'False
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin prjStudentIMS.lvButtons_H cmdNavigate 
         Height          =   315
         Index           =   2
         Left            =   720
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Next Record"
         Top             =   2100
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         CapAlign        =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Focus           =   0   'False
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin prjStudentIMS.lvButtons_H cmdNavigate 
         Height          =   315
         Index           =   3
         Left            =   1080
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Last Record"
         Top             =   2100
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         CapAlign        =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Focus           =   0   'False
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin prjStudentIMS.lvButtons_H cmdFooter 
         Height          =   315
         Left            =   0
         TabIndex        =   20
         Top             =   2070
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   556
         CapAlign        =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Focus           =   0   'False
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
   End
   Begin MSComctlLib.ImageList imlTrans 
      Left            =   3960
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":302CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30668
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30A02
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30D9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":31136
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":314D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3186A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":361A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3653E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3BDC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C7D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3CAEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D086
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D620
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3DBBA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin prjStudentIMS.lvButtons_H cmdHelp 
      Height          =   375
      Left            =   8220
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   0
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Help"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin prjStudentIMS.lvButtons_H cmdPrint 
      Height          =   375
      Left            =   5850
      TabIndex        =   5
      Top             =   0
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Print"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin prjStudentIMS.lvButtons_H cmdSearch 
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   0
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Search"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin prjStudentIMS.ucVertical3DLine ucUser_Sep 
      Height          =   255
      Left            =   9600
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
   End
   Begin prjStudentIMS.lvButtons_H cmdUser 
      Height          =   375
      Left            =   9480
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   0
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   661
      Caption         =   "User"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16711680
      cFHover         =   16711680
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmMain.frx":3E154
   End
   Begin prjStudentIMS.lvButtons_H cmdToolbar 
      Height          =   375
      Left            =   9060
      TabIndex        =   7
      Top             =   0
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   661
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Visible         =   0   'False
      Begin VB.Menu mnuTools_Item 
         Caption         =   "Backup Database"
         Index           =   0
      End
      Begin VB.Menu mnuTools_Item 
         Caption         =   "Restore Database"
         Index           =   1
      End
      Begin VB.Menu mnuTools_Item 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuTools_Item 
         Caption         =   "Settings Options"
         Index           =   6
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Visible         =   0   'False
      Begin VB.Menu mnuHelp_Item 
         Caption         =   "Online Manual"
         Index           =   1
      End
      Begin VB.Menu mnuHelp_Item 
         Caption         =   "Goto Website"
         Index           =   2
      End
      Begin VB.Menu mnuHelp_Item 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuHelp_Item 
         Caption         =   "About..."
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-- Local Variables.
Private m_sSort         As String
Private m_lFieldIdx     As Long
Dim z, itemToAdd


'//---------------------------------------------------------------------------------------
'//--Procedure : initPanel
Private Sub initPanel()

On Error GoTo errHandler

    Dim c As pTab

    Me.pMenu.Clear

    Me.pMenu.LockUpdate = True
    
    Set c = Me.pMenu.Panels.Add("List of all existing students", imlPanel.ListImages(13).Picture, "Existing students list", "Students", picData, True)
    'In the line above "Students" is the pMenu.SelectedItem.Key which will be required when the toolbar buttons
    'are clicked to Add, Edit, Delete, Search Data. Note it carefully. This Key value should be
    'different and unique for all menu pMenu items. See the Click event of Add or Edit button.
    'Set c = Me.pMenu.Panels.Add("List of college ex-students", imlPanel.ListImages(19).Picture, "Ex-students list", "ExStudents", picData)
    Set c = Me.pMenu.Panels.Add("List of current departmental staffs", imlPanel.ListImages(18).Picture, "Department's staff list", "DeptStaffs", picData)
    Set c = Me.pMenu.Panels.Add("List of assignments alloted to the students", imlPanel.ListImages(26).Picture, "Alloted assignments list", "Alloted_Assignms", picData)
    Set c = Me.pMenu.Panels.Add("Assignments salted for submission today", imlPanel.ListImages(38).Picture, "Today's assignments", "Todays_Assignms", picData)
    Set c = Me.pMenu.Panels.Add("List of all laboratory books and reference materials", imlPanel.ListImages(30).Picture, "Laboratoty books and references", "Book_Ref", picData)
    Set c = Me.pMenu.Panels.Add("List of borrowed laboratory books & reference materials", imlPanel.ListImages(12).Picture, "Borrowed books and references", "Borr_BookRef", picData)
    Set c = Me.pMenu.Panels.Add("List of students' internal assessment marks", imlPanel.ListImages(9).Picture, "Internal assessment marks", "IntAssmt_Marks", picData)
    Set c = Me.pMenu.Panels.Add("Students Fee Collection Records", imlPanel.ListImages(5).Picture, "Student fee records", "StudFee_Records", picData)
    Set c = Me.pMenu.Panels.Add("Various application specific reports", imlPanel.ListImages(11).Picture, "System Reports", "Sys_Rpts", tvwReport)
    Me.pMenu.LockUpdate = False

errHandler:
    Set c = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

'//---------------------------------------------------------------------------------------
'//--Procedure : initStatusbar
'
Private Sub initStatusbar()
    Dim tmpFont As StdFont
    
On Error GoTo errHandler
    
    With ucStatus
        
        '-- Initialize statusbar
        Call .Initialize(SizeGrip:=True, ToolTips:=True)
        
        .Font.Size = 8
        .Font.Name = "Tahoma"
        
        '-- Initialize icons list
        Call .InitializeIconList
    
        '-- Add icons
        'Call .AddIcon(LoadResPicture("MAIL", vbResIcon))
        'Call .AddIcon(LoadResPicture("USER", vbResIcon))
        'Call .AddIcon(LoadResPicture("TIP", vbResIcon))
        'Call .AddIcon(LoadResPicture(105, 0))
        
        '-- Add panels
        Call .AddPanel(, , , [sbSpring], CPYRYT, "Program designer and developer: Partha S. Paul", 0)
        'Call .AddPanel(, , , [sbContents], Format(Date, "MMM dd, yyyy"), "")
        Call .AddPanel(, , , [sbContents], "  Academic Year: " & cur_acadYear & "  ", "")
        If CInt(user_level) = 1 Then
            Call .AddPanel(, , , [sbContents], "  User Level: Administrator  ", "")
        Else 'If CInt(user_level) = 0 Then
            Call .AddPanel(, , 400, [sbContents], "  User Level: User  ", "")
        End If
    End With

errHandler:
    Set tmpFont = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

'//---------------------------------------------------------------------------------------
'//--Procedure : initTreeview
'//--DateTime  : 9/2/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   :
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Private Sub initTreeview()

On Error GoTo errHandler
Dim tv1 As Node
Set tv1 = tvwReport.Nodes.Add(, , "main_menu", "System Reports Menu", 5, 4)
tv1.Tag = "main_menu"
tv1.ForeColor = vbRed
'tv1.EnsureVisible
'uncommenting line above and commenting line below
'causes the treeview explorer to collapse
tv1.Expanded = True
Set tv1 = tvwReport.Nodes.Add("main_menu", tvwChild, "students", "Students Inventory", 1, 5)
tv1.Tag = "students"
tv1.ForeColor = vbBlue
Set tv1 = tvwReport.Nodes.Add("students", tvwChild, "assignmt", "Assignments And Tasks Submission History", 7, 15)
tv1.Tag = "assignmt"
tv1.Expanded = True
Set tv1 = tvwReport.Nodes.Add("students", tvwChild, "booktrans", "Books & References Transection History", 7, 15)
tv1.Tag = "booktrans"
tv1.Expanded = True
'---------------------------------
Set tv1 = tvwReport.Nodes.Add("main_menu", tvwChild, "staffs", "Staffs Inventory", 1, 5)
tv1.Tag = "staffs"
tv1.ForeColor = vbBlue
Set tv1 = tvwReport.Nodes.Add("staffs", tvwChild, "purTrans", "Assignments And Tasks Allotment History", 7, 15)
tv1.Tag = "purTrans"
tv1.Expanded = True
Set tv1 = tvwReport.Nodes.Add("staffs", tvwChild, "purRetTrans", "Refresher And Orientation Courses Attendence Report", 7, 15)
tv1.Tag = "purRetTrans"
tv1.Expanded = True
Set tv1 = tvwReport.Nodes.Add("staffs", tvwChild, "searchPurTrans", "List Of Retired/Resigned Staffs", 7, 15)
tv1.Tag = "searchPurTrans"
tv1.Expanded = True
'---------------------------------
Set tv1 = tvwReport.Nodes.Add("main_menu", tvwChild, "accounts", "Books And References", 1, 5)
tv1.Tag = "accounts"
tv1.ForeColor = vbBlue
Set tv1 = tvwReport.Nodes.Add("accounts", tvwChild, "cashbook", "List Of All Departmental Books & References", 7, 15)
tv1.Tag = "cashbook"
tv1.Expanded = True
Set tv1 = tvwReport.Nodes.Add("accounts", tvwChild, "cashdisburse", "List Of Books & References Lying With Students", 7, 15)
tv1.Tag = "cashdisburse"
tv1.Expanded = True
Set tv1 = tvwReport.Nodes.Add("accounts", tvwChild, "journel", "List Of Books & References Lying With Staffs", 7, 15)
tv1.Tag = "journel"
tv1.Expanded = True
'---------------------------------
Set tv1 = tvwReport.Nodes.Add("main_menu", tvwChild, "sys_reports", "Lab Equipment Inventory", 1, 5)
tv1.Tag = "sys_reports"
tv1.ForeColor = vbBlue
Set tv1 = tvwReport.Nodes.Add("sys_reports", tvwChild, "salesReport", "Lab Equipment Purchase Report", 7, 15)
tv1.Tag = "salesReport"
tv1.Expanded = True
Set tv1 = tvwReport.Nodes.Add("sys_reports", tvwChild, "purchaseReport", "Lab Equipment Usage Report", 7, 15)
tv1.Tag = "purchaseReport"
tv1.Expanded = True
Set tv1 = tvwReport.Nodes.Add("sys_reports", tvwChild, "incomeReport", "Lab Equipment Status Report", 7, 15)
tv1.Tag = "incomeReport"
tv1.Expanded = True
Set tv1 = tvwReport.Nodes.Add("sys_reports", tvwChild, "expensesReport", "Miscellaneous Lab Expenses Report", 7, 15)
tv1.Tag = "expensesReport"
tv1.Expanded = True

errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmMain", "initTreeview", True)
    End If
End Sub


'//---------------------------------------------------------------------------------------
'//--Procedure : setToolbar
'//--DateTime  : 7/12/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   :
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Private Sub setToolbar()

On Error GoTo errHandler

    Select Case pMenu.SelectedItem.Key
        Case "Sys_Rpts", "Lab_Equipmts"
            cmdAdd.Enabled = False
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            cmdRefresh.Enabled = False
            cmdSearch.Enabled = False
        Case "Todays_Assignms"
            cmdAdd.Enabled = False
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            cmdRefresh.Enabled = True
            cmdSearch.Enabled = True
        Case Else
            cmdAdd.Enabled = True
            cmdEdit.Enabled = True
            cmdDelete.Enabled = True
            cmdRefresh.Enabled = True
            cmdSearch.Enabled = True
    End Select

errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmMain", "setToolbar", True)
    End If
End Sub


'//---------------------------------------------------------------------------------------
'//--Procedure : ShowRecordInfo
'//--DateTime  : 8/26/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   :
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Public Sub ShowRecordInfo(rs As ADODB.Recordset, cmdLavolpe As lvButtons_H)

On Error GoTo errHandler

    'Dim lPos    As Long
    'Dim lCount  As Long
    
    
    'lPos = rs.AbsolutePosition
    'Extra Line added
    'If lPos < 0 Then lPos = 1
    'lCount = rs.RecordCount
    
    'cmdLavolpe.Caption = IIf(lCount < 1, 0, lPos) & " of " & lCount
    cmdLavolpe.Caption = "Total number of record(s) found : " & lvListView.ListItems.Count
errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmMain", "ShowRecordInfo", True)
    End If
End Sub


Private Sub cmdAdd_Click()
    Select Case pMenu.SelectedItem.Key
        Case "Students" 'g_eMenu.Patient
            With frmPatient
                .g_lAddState = 1    '-- Add State.
                .Caption = frmMain.Caption & " [Add New Student]"
                .Show vbModal
            End With
        Case "DeptStaffs"
            With frmStaff
                .g_StaffAddState = 1    '-- Add State.
                .Caption = Me.Caption & " [Add New Staff]"
                .Show vbModal
            End With
        Case "Alloted_Assignms"
            With frmAssignmt
                .g_AssignAddState = 1    '-- Add State.
                .Caption = Me.Caption & " [Add New Assignment]"
                .Show vbModal
            End With
        Case "IntAssmt_Marks"
            With frmIntAssmt
                .g_AssmtAddState = 1    '-- Add State.
                .Caption = Me.Caption & " [Add Internal Assessment Mark]"
                .Show vbModal
            End With
        Case "Book_Ref"
            With frmBookJR
                .g_BookRefAddState = 1    '-- Add State.
                .Caption = Me.Caption & " [Add New Book/Reference]"
                .Show vbModal
            End With
        Case "Borr_BookRef"
            With frmBookJrTrans
                .g_BJrAddState = 1    '-- Add State.
                .Caption = Me.Caption & " [Add Book/Ref Transection Data]"
                .Show vbModal
            End With
    End Select
End Sub


Private Sub cmdDelete_Click()
    Dim lResult As Long
    Dim sel_ID As Long
    If lvListView.ListItems.Count = 0 Then MsgBox "There is currently no record on display to delete.", vbExclamation, "No Record...": Exit Sub
    Select Case pMenu.SelectedItem.Key
        Case "Students"
            lResult = MsgBox("Are you sure you want to delete the selected record?", vbQuestion + vbYesNo, "Delete Student Data...")
            If lResult = vbYes Then
                Set rsTemp = New ADODB.Recordset
                sel_ID = CLng(lvListView.SelectedItem.Text)
                rsTemp.Open "Select * FROM BJTransection WHERE StudentID=" & CLng(sel_ID) & "", g_cn, adOpenKeyset, adLockOptimistic
                    If rsTemp.EOF Or rsTemp.BOF Then
                        If rsTemp.State = adStateOpen Then rsTemp.Close
                        rsTemp.Open "Delete * FROM Students WHERE StudentID=" & CLng(sel_ID) & "", g_cn, adOpenKeyset, adLockOptimistic
                        'rsTemp.Close
                        Set rsTemp = Nothing
                        Call frmMain.cmdRefresh_Click
                        MsgBox "The selected student's data has been successfully deleted from the database.", vbInformation, "Delete Operation Successful..."
                        Exit Sub
                    Else
                        rsTemp.Close
                        Set rsTemp = Nothing
                        MsgBox "It seems that some book(s) or reference material(s) is/are still" & _
                               "lying with the student. Hence the student data cannot be deleted.", vbExclamation, "Delete Operation Unsuccessful..."
                       
                    End If
            End If
        Case "DeptStaffs"
            lResult = MsgBox("Are you sure you want to delete the selected record?", vbQuestion + vbYesNo, "Delete Staff Data...")
            If lResult = vbYes Then
                Set rsTemp = New ADODB.Recordset
                sel_ID = CLng(lvListView.SelectedItem.Text)
                rsTemp.Open "Select * FROM IntAssmt WHERE StaffID=" & CLng(sel_ID) & "", g_cn, adOpenKeyset, adLockOptimistic
                    If rsTemp.EOF Or rsTemp.BOF Then
                        If rsTemp.State = adStateOpen Then rsTemp.Close
                        rsTemp.Open "Delete * FROM Staff WHERE StaffID=" & CLng(sel_ID) & "", g_cn, adOpenKeyset, adLockOptimistic
                        'rsTemp.Close
                        Set rsTemp = Nothing
                        Call frmMain.cmdRefresh_Click
                        MsgBox "The selected staff's data has been successfully deleted from the database.", vbInformation, "Delete Operation Successful..."
                        Exit Sub
                    Else
                        rsTemp.Close
                        Set rsTemp = Nothing
                        MsgBox "It seems that the selected staff has carried out internal assessments of some of " & _
                               "the students and has marks assigned to them. If the staff record is deleted then " & _
                               "those student's internal assessment marks will become invalid. Hence the staff's " & _
                               "data cannot be deleted.", vbExclamation, "Delete Operation Unsuccessful..."
                       
                    End If
            End If
        Case "Alloted_Assignms"
            lResult = MsgBox("Are you sure you want to delete the selected record?", vbQuestion + vbYesNo, "Delete Assignment Record...")
            If lResult = vbYes Then
                Set rsTemp = New ADODB.Recordset
                sel_ID = CLng(lvListView.SelectedItem.Text)
                rsTemp.Open "Delete * FROM Assignments WHERE AssignID=" & CLng(sel_ID) & "", g_cn, adOpenKeyset, adLockOptimistic
                'rsTemp.Close
                Set rsTemp = Nothing
                Call frmMain.cmdRefresh_Click
                MsgBox "The selected assignment record has been successfully deleted from the database.", vbInformation, "Delete Operation Successful..."
                Exit Sub
            End If
        Case "IntAssmt_Marks"
            lResult = MsgBox("Are you sure you want to delete the selected record?", vbQuestion + vbYesNo, "Delete Internal Assessment Record...")
            If lResult = vbYes Then
                Set rsTemp = New ADODB.Recordset
                sel_ID = CLng(lvListView.SelectedItem.Text)
                rsTemp.Open "Delete * FROM IntAssmt WHERE IntAssmtID=" & CLng(sel_ID) & "", g_cn, adOpenKeyset, adLockOptimistic
                'rsTemp.Close
                Set rsTemp = Nothing
                Call frmMain.cmdRefresh_Click
                MsgBox "The selected internal assessment record has been successfully deleted from the database.", vbInformation, "Delete Operation Successful..."
                Exit Sub
            End If
        Case "Book_Ref"
            lResult = MsgBox("Are you sure you want to delete the selected record?", vbQuestion + vbYesNo, "Delete Book/Reference Record...")
            If lResult = vbYes Then
                Set rsTemp = New ADODB.Recordset
                sel_ID = CLng(lvListView.SelectedItem.Text)
                rsTemp.Open "Select * FROM BJTransection WHERE BookJrID=" & CLng(sel_ID) & " AND Returned = 0", g_cn, adOpenKeyset, adLockOptimistic
                    If rsTemp.EOF Or rsTemp.BOF Then
                        If rsTemp.State = adStateOpen Then rsTemp.Close
                        rsTemp.Open "Delete * FROM BookJournel WHERE BookJrID=" & CLng(sel_ID) & "", g_cn, adOpenKeyset, adLockOptimistic
                        'rsTemp.Close
                        Set rsTemp = Nothing
                        Call frmMain.cmdRefresh_Click
                        MsgBox "The selected book/reference record has been successfully deleted from the database.", vbInformation, "Delete Operation Successful..."
                        Exit Sub
                    Else
                        rsTemp.Close
                        Set rsTemp = Nothing
                        MsgBox "It seems that one or more copies of  the book/reference is/are still" & vbCrLf & _
                               "lying with the student(s). Hence the record cannot be deleted.", vbExclamation, "Delete Operation Unsuccessful..."
                       
                    End If
            End If
        
        Case "Borr_BookRef"
            Dim book_returned_status, book_ref_id As Long, count_copies_issued  As Integer
            
            lResult = MsgBox("Are you sure you want to delete the selected record?", vbQuestion + vbYesNo, "Delete Book/Ref Transection Record...")
            If lResult = vbYes Then
                Set rsTemp = New ADODB.Recordset
                sel_ID = CLng(lvListView.SelectedItem.Text)
                book_returned_status = CStr(lvListView.SelectedItem.SubItems(9))
                
                rsTemp.Open "Select * FROM BJTransection WHERE TransID=" & CLng(sel_ID) & "", g_cn, adOpenKeyset, adLockOptimistic
                book_ref_id = CLng(rsTemp!BookJrID)
                    If (book_returned_status = "Yes. Returned back.") Then
                        If rsTemp.State = adStateOpen Then rsTemp.Close
                        rsTemp.Open "Delete * FROM BJTransection WHERE TransID=" & CLng(sel_ID) & "", g_cn, adOpenKeyset, adLockOptimistic
                        'rsTemp.Close
                        Set rsTemp = Nothing
                        Call frmMain.cmdRefresh_Click
                        MsgBox "The selected book/ref transection record has been successfully deleted from the database.", vbInformation, "Delete Operation Successful..."
                        Exit Sub
                    Else
                        If rsTemp.State = adStateOpen Then rsTemp.Close
                        rsTemp.Open "Select * FROM BookJournel WHERE BookJrID=" & CLng(book_ref_id) & "", g_cn, adOpenKeyset, adLockOptimistic
                        count_copies_issued = rsTemp!CopiesIssued
                        count_copies_issued = count_copies_issued - 1
                        rsTemp!CopiesIssued = count_copies_issued
                        rsTemp.Update
                        'then delete the transection record
                        If rsTemp.State = adStateOpen Then rsTemp.Close
                        rsTemp.Open "Delete * FROM BJTransection WHERE TransID=" & CLng(sel_ID) & "", g_cn, adOpenKeyset, adLockOptimistic
                        'rsTemp.Close
                        Set rsTemp = Nothing
                        Call frmMain.cmdRefresh_Click
                        MsgBox "The selected book/ref transection record has been successfully deleted from the database.", vbInformation, "Delete Operation Successful..."
                        Exit Sub
                    End If
            End If
     End Select
    
End Sub

Private Sub cmdEdit_Click()
    Call lvListView_DblClick
End Sub

Private Sub cmdHelp_Click()
    Call PopupMenu(mnuHelp, , cmdHelp.Left, cmdHelp.Top + cmdHelp.Height)
End Sub

Private Sub cmdNavigate_Click(Index As Integer)
Set rsTemp = New ADODB.Recordset
If lvListView.ColumnHeaders.Item(1).Text = "Student ID" Then
    Select Case Index
        Case 0
        rsTemp.Open "Select * FROM qryStudents", g_cn, adOpenKeyset, adLockOptimistic
        'Call Navigate(Index, rsTemp)
        'Call ShowRecordInfo(rsTemp, cmdFooter)
        rsTemp.Close
    End Select
End If
Set rsTemp = Nothing
End Sub

Private Sub cmdPrint_Click()
    MsgBox "Viewing and printing of report(s) is/are not available with the demo version of this" & vbCrLf & "program. You need to use the full version for getting these facilities.", vbExclamation, "Demo Version Limitations..."
    Exit Sub
End Sub

Public Sub cmdRefresh_Click()
    Select Case pMenu.SelectedItem.Key
        Case "Students"
            lvListView.ListItems.Clear
            Call initFrame_Students
            Call displayStudents
        Case "DeptStaffs"
            lvListView.ListItems.Clear
            Call initFrame_Staffs
            Call displayStaffs
        Case "Alloted_Assignms"
            lvListView.ListItems.Clear
            Call initFrame_Assignments
            Call displayAssignments
        Case "Todays_Assignms"
            lvListView.ListItems.Clear
            Call initFrame_Assignments
            Call displayTodaysAssignms
        Case "IntAssmt_Marks"
            lvListView.ListItems.Clear
            Call initFrame_IntAssmt
            Call displayIntAssmtMarks
        Case "Book_Ref"
            lvListView.ListItems.Clear
            Call initFrame_BookRef
            Call displayBookRef
        Case "Borr_BookRef"
            lvListView.ListItems.Clear
            Call initFrame_BJTrans
            Call displayBJTrans
    End Select
End Sub

Private Sub cmdSearch_Click()
    If lvListView.ListItems.Count = 0 Then
        MsgBox "There is no item in the list currently on display to search for.", vbExclamation, "No Records..."
        Exit Sub
    Else
        Select Case pMenu.SelectedItem.Key
            Case "Students"
                Load frmSearch
                frmSearch.Show vbModal
            Case "DeptStaffs"
                Load frmSearch_Staff
                frmSearch_Staff.Show vbModal
            Case "Alloted_Assignms"
                frmSearch_Assignms.g_TodaysAssignms = 0
                Load frmSearch_Assignms
                frmSearch_Assignms.Show vbModal
            Case "Todays_Assignms"
                frmSearch_Assignms.g_TodaysAssignms = 1
                Load frmSearch_Assignms
                frmSearch_Assignms.Show vbModal
            Case "IntAssmt_Marks"
                Load frmSearch_IntAssmtMarks
                frmSearch_IntAssmtMarks.Show vbModal
            Case "Book_Ref"
                Load frmSearch_BookRef
                frmSearch_BookRef.Show vbModal
            Case "Borr_BookRef"
                Load frmSearch_BJTrans
                frmSearch_BJTrans.Show vbModal
        End Select
    End If
End Sub

Private Sub cmdTools_Click()
    Call PopupMenu(mnuTools, , cmdTools.Left, cmdTools.Top + cmdTools.Height)
End Sub

Private Sub cmdUser_Click()
    If MsgBox("Are you sure you that you want to log off?", vbQuestion + vbYesNo, "Log Off...") = vbYes Then
        g_lLogOff = 1   '-- Set LogOff flag.
        Call Unload(Me)
    End If
End Sub

Public Sub Form_Load()
    Dim lidx   As Long
    
    On Error GoTo errHandler
    
    Caption = g_sAppName
     
    '-- Initialize.
    Call initPanel
    
    Call Detect_AcadYRSetting
    
    Call initStatusbar
    Call initTreeview
    Call initToolbar(cmdToolbar, imlTrans)
    Call initToolbar(cmdAdd, imlTrans, , g_eIcon.Add)
    Call initToolbar(cmdEdit, imlTrans, , g_eIcon.Edit)
    Call initToolbar(cmdDelete, imlTrans, , g_eIcon.Delete)
    Call initToolbar(cmdRefresh, imlTrans, , g_eIcon.Refresh)
    Call initToolbar(cmdSearch, imlTrans, , g_eIcon.Search)
    Call initToolbar(cmdPrint, imlTrans, , g_eIcon.Print)
    Call initToolbar(cmdTools, imlTrans, , g_eIcon.Options)
'   Call initToolbar(cmdAbout, imlTrans)
    Call initToolbar(cmdUser, imlTrans)
    Call initToolbar(cmdHelp, imlTrans, , g_eIcon.Help)
    Call initToolbar(cmdFooter, imlTrans, 1)
    Call initToolbar(cmdNavigate(0), imlTrans, 1, g_eIcon.BOF, lv_Fill_Stretch)
    Call initToolbar(cmdNavigate(1), imlTrans, 1, g_eIcon.Previous, lv_Fill_Stretch)
    Call initToolbar(cmdNavigate(2), imlTrans, 1, g_eIcon.Next, lv_Fill_Stretch)
    Call initToolbar(cmdNavigate(3), imlTrans, 1, g_eIcon.EOF, lv_Fill_Stretch)

    '-- Default.
    Call initFrame_Students
    Call displayStudents
errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmMain", "Form_Load", True)
        Call Unload(Me)
    End If
    
End Sub

Private Sub Form_Resize()
On Error GoTo errHandler

    ucStatus.SizeGrip = (WindowState <> vbMaximized)
    pMenu.Move 20, cmdToolbar.Height + 20, ScaleWidth - 30, ScaleHeight - (30 + cmdToolbar.Height + ucStatus.Height)
    cmdUser.Left = ScaleWidth - cmdUser.Width
    ucDate_Sep.Left = cmdUser.Left + cmdUser.Width
    'cmdUser.Left = cmdDate.Left - cmdUser.Width
    ucUser_Sep.Left = cmdUser.Left
    cmdToolbar.Width = ScaleWidth - cmdToolbar.Left
    ucToolbar.Width = ScaleWidth + 20
    
errHandler: Err.Clear: Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
     If g_lLogOff = 1 Then
        g_cn.Close
        Set g_cn = Nothing
        Unload Me
        Load frmLogin
        frmLogin.Show
    Else
        If MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo, "Exit Application") = vbNo Then
            Cancel = 1
            Me.Show
        Else
            g_cn.Close
            Set g_cn = Nothing
            Unload Me
            End
        End If
    End If
End Sub


Private Sub lvListView_DblClick()
Dim staff_id As Long, student_id As Long, book_ref_id As Long
If lvListView.ListItems.Count = 0 Then Exit Sub
If pMenu.SelectedItem.Key = "Todays_Assignms" Then Exit Sub
'-- Student Records
If lvListView.ColumnHeaders.Item(1).Text = "Student ID" Then
    frmPatient.g_lAddState = 0 '-- Edit State.
    Load frmPatient
    frmPatient.Text2.Text = lvListView.SelectedItem.Text
    frmPatient.txtName.Text = lvListView.SelectedItem.SubItems(1)
    frmPatient.txtAddress.Text = lvListView.SelectedItem.SubItems(11)
    frmPatient.txtAY.Text = lvListView.SelectedItem.SubItems(7)
    frmPatient.txtCPhone.Text = lvListView.SelectedItem.SubItems(9)
    frmPatient.txtEMail.Text = lvListView.SelectedItem.SubItems(8)
    frmPatient.txtMPhone.Text = lvListView.SelectedItem.SubItems(10)
    frmPatient.txtSection.Text = lvListView.SelectedItem.SubItems(6)
    frmPatient.txtRollNo.Text = lvListView.SelectedItem.SubItems(4)
    frmPatient.txtRegdNo.Text = lvListView.SelectedItem.SubItems(5)
    frmPatient.txtClass.Text = lvListView.SelectedItem.SubItems(3)
    frmPatient.DTPicker1.Value = CDate(lvListView.SelectedItem.SubItems(2))
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open "SELECT * FROM Students WHERE StudentID=" & CLng(lvListView.SelectedItem.Text) & "", g_cn, adOpenKeyset, adLockOptimistic
    frmPatient.txtClassID.Text = rsTemp!ClassID
    frmPatient.txtAcadYRID.Text = rsTemp!AcadYearID
    rsTemp.Close
    Set rsTemp = Nothing
    frmPatient.Caption = frmMain.Caption & " [Edit Student Record]"
    frmPatient.Show vbModal
    
ElseIf lvListView.ColumnHeaders.Item(1).Text = "Staff ID" Then
    frmStaff.g_StaffAddState = 0 '-- Edit State.
    Load frmStaff
    frmStaff.Text2.Text = lvListView.SelectedItem.Text
    frmStaff.txtName.Text = lvListView.SelectedItem.SubItems(1)
    frmStaff.txtAddress.Text = lvListView.SelectedItem.SubItems(9)
    frmStaff.txtCPhone.Text = lvListView.SelectedItem.SubItems(4)
    frmStaff.txtMPhone.Text = lvListView.SelectedItem.SubItems(5)
    frmStaff.txtDesignation.Text = lvListView.SelectedItem.SubItems(3)
    frmStaff.txtQual.Text = lvListView.SelectedItem.SubItems(7)
    frmStaff.txtSpecial.Text = lvListView.SelectedItem.SubItems(8)
    frmStaff.txtEMail.Text = lvListView.SelectedItem.SubItems(6)
    frmStaff.txtRemarks.Text = lvListView.SelectedItem.SubItems(10)
    frmStaff.DTPicker1.Value = CDate(lvListView.SelectedItem.SubItems(2))
    frmStaff.cboStatus.Text = lvListView.SelectedItem.SubItems(11)
    frmStaff.Caption = frmMain.Caption & " [Edit Staff Record]"
    frmStaff.Show vbModal
    
ElseIf lvListView.ColumnHeaders.Item(1).Text = "Assignmt ID" Then
    If user_level <> 1 Then MsgBox "You don't have the sufficient privilege to edit students' assignment data." & vbCrLf & "Only an administrative user can perform such a task.", vbExclamation, "Warning...": Exit Sub
    frmAssignmt.g_AssignAddState = 0 '-- Edit State.
    '----for finding out the StudentID & StaffID of the selected assignment
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open "SELECT * FROM Assignments WHERE AssignID=" & CLng(lvListView.SelectedItem.Text) & "", g_cn, adOpenKeyset, adLockOptimistic
    staff_id = CLng(rsTemp!StaffID)
    student_id = CLng(rsTemp!StudentID)
    rsTemp.Close
    Set rsTemp = Nothing
    '--- now load the Form for Editing
    Load frmAssignmt
    frmAssignmt.Text2.Text = lvListView.SelectedItem.Text
    frmAssignmt.txtDescp.Text = lvListView.SelectedItem.SubItems(7)
    frmAssignmt.DTPicker1.Value = CDate(lvListView.SelectedItem.SubItems(1))
    frmAssignmt.DTPicker2.Value = CDate(lvListView.SelectedItem.SubItems(2))
    frmAssignmt.txtAT.Text = lvListView.SelectedItem.SubItems(3)
    frmAssignmt.txtAB.Text = lvListView.SelectedItem.SubItems(6)
    frmAssignmt.txtStaffID = staff_id
    frmAssignmt.txtStudID = student_id
    frmAssignmt.Caption = frmMain.Caption & " [Edit Assignment Record]"
    frmAssignmt.Show vbModal
    
ElseIf lvListView.ColumnHeaders.Item(1).Text = "Assessment ID" Then
    If user_level <> 1 Then MsgBox "You don't have the sufficient privilege to edit the internal assessment marks." & vbCrLf & "Only an administrative user can perform such a task.", vbExclamation, "Warning...": Exit Sub
    frmIntAssmt.g_AssmtAddState = 0 '-- Edit State.
    '----for finding out the StudentID & StaffID of the selected assignment
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open "SELECT * FROM IntAssmt WHERE IntAssmtID=" & CLng(lvListView.SelectedItem.Text) & "", g_cn, adOpenKeyset, adLockOptimistic
    staff_id = CLng(rsTemp!StaffID)
    student_id = CLng(rsTemp!StudentID)
    rsTemp.Close
    Set rsTemp = Nothing
    '--- now load the Form for Editing
    
    Load frmIntAssmt
    frmIntAssmt.Text2.Text = lvListView.SelectedItem.Text
    frmIntAssmt.txtTerm.Text = lvListView.SelectedItem.SubItems(1)
    frmIntAssmt.txtStudent.Text = lvListView.SelectedItem.SubItems(3)
    frmIntAssmt.txtClass.Text = lvListView.SelectedItem.SubItems(5)
    frmIntAssmt.txtSection.Text = lvListView.SelectedItem.SubItems(6)
    frmIntAssmt.txtRoll.Text = lvListView.SelectedItem.SubItems(4)
    frmIntAssmt.txtMarks.Text = lvListView.SelectedItem.SubItems(8)
    frmIntAssmt.txtStaff.Text = lvListView.SelectedItem.SubItems(7)
    frmIntAssmt.DTPicker1.Value = FormatDateTime(lvListView.SelectedItem.SubItems(2))
    frmIntAssmt.txtRemarks.Text = lvListView.SelectedItem.SubItems(9)
    frmIntAssmt.txtStaffID.Text = staff_id
    frmIntAssmt.txtStudID.Text = student_id
    frmIntAssmt.Caption = frmMain.Caption & " [Edit Internal Assessment Record]"
    frmIntAssmt.Show vbModal
    
ElseIf lvListView.ColumnHeaders.Item(1).Text = "Book/Ref ID" Then
    Load frmBookJR
    frmBookJR.Text2.Text = lvListView.SelectedItem.Text
    frmBookJR.txtTitle.Text = lvListView.SelectedItem.SubItems(1)
    frmBookJR.txtAuthors.Text = lvListView.SelectedItem.SubItems(2)
    frmBookJR.txtPages.Text = lvListView.SelectedItem.SubItems(4)
    frmBookJR.txtPrice.Text = lvListView.SelectedItem.SubItems(5)
    frmBookJR.txtNOC.Text = lvListView.SelectedItem.SubItems(6)
    frmBookJR.txtPublisher.Text = lvListView.SelectedItem.SubItems(3)
    frmBookJR.txtDescp.Text = lvListView.SelectedItem.SubItems(8)
    frmBookJR.txtCI.Text = CInt(lvListView.SelectedItem.SubItems(6) - lvListView.SelectedItem.SubItems(7))
    frmBookJR.Caption = frmMain.Caption & " [Edit Book/Reference Record]"
    frmBookJR.Show vbModal

ElseIf lvListView.ColumnHeaders.Item(1).Text = "Trans. ID" Then
    frmBookJrTrans.g_BJrAddState = 0 '-- Edit State.
    '----for finding out the BookID & StudentID of the selected transection
    Dim date_of_issue As Date
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open "SELECT * FROM BJTransection WHERE TransID=" & CLng(lvListView.SelectedItem.Text) & "", g_cn, adOpenKeyset, adLockOptimistic
    book_ref_id = CLng(rsTemp!BookJrID)
    student_id = CLng(rsTemp!StudentID)
    date_of_issue = CDate(rsTemp!DOI)
    rsTemp.Close
    Set rsTemp = Nothing
    '--- now load the Form for Editing
    Load frmBookJrTrans
    frmBookJrTrans.Text2.Text = lvListView.SelectedItem.Text
    frmBookJrTrans.txtTitle.Text = lvListView.SelectedItem.SubItems(1)
    frmBookJrTrans.txtBookID.Text = book_ref_id
    frmBookJrTrans.txtAuthors.Text = lvListView.SelectedItem.SubItems(2)
    frmBookJrTrans.txtStudent.Text = lvListView.SelectedItem.SubItems(3)
    frmBookJrTrans.txtStudID.Text = student_id
    frmBookJrTrans.txtClass.Text = lvListView.SelectedItem.SubItems(4)
    frmBookJrTrans.txtRoll.Text = lvListView.SelectedItem.SubItems(5)
    frmBookJrTrans.txtSection.Text = lvListView.SelectedItem.SubItems(6)
    frmBookJrTrans.dtReturn.Value = CDate(Date)
    frmBookJrTrans.dtReturn.Enabled = False
    frmBookJrTrans.dtIssue.Value = date_of_issue
    frmBookJrTrans.dtIssue.Enabled = False
    frmBookJrTrans.cboRetrn.Text = "Returning now..."
    frmBookJrTrans.cboRetrn.Enabled = False
    frmBookJrTrans.cmdAT.Enabled = False
    frmBookJrTrans.cmdAB.Enabled = False
    frmBookJrTrans.Caption = frmMain.Caption & " [Update Book/Reference Return Record]"
    frmBookJrTrans.Show vbModal
    
End If
End Sub

Private Sub mnuHelp_Item_Click(Index As Integer)
    Select Case Index
        Case 0
            '
        Case 1
            MsgBox "The online help manual is not available with the demo version of this application.", vbExclamation, "Online Help Missing..."
        Case 2
            OpenURL "http://www.prosventech.net/", Me.hWnd
        Case 3
            '
        Case 4  '-- About.
            frmAbout.Show vbModal
    End Select
End Sub

Private Sub mnuTools_Item_Click(Index As Integer)
    Select Case Index
        Case 0  '-- Backup.
            If user_level <> 1 Then MsgBox "You don't have the sufficient privilege to take a database backup." & vbCrLf & _
                                           "Only an administrative user can perform such a task.", vbExclamation, "Warning...": Exit Sub
            frmBackup.Show vbModal
        Case 1  '-- Restore.
            If user_level <> 1 Then MsgBox "You don't have the sufficient privilege to restore the database from an" & vbCrLf & _
                                           "earlier backup. Only an administrative user can perform such a task.", vbExclamation, "Warning...": Exit Sub
            frmRestore.Show vbModal
        'Case 3  '-- Calculator.
            'Call OpenThisFile("Calc", "1", "", hWnd)
        'Case 4  '-- Notepad.
            'Call OpenThisFile("Notepad", "1", "", hWnd)
        Case 6  '-- Options.
            frmOptions.Show vbModal
        Case Else
    End Select
End Sub

Private Sub picData_Resize()
On Error GoTo errHandler
    lvListView.Move 0, 0, picData.ScaleWidth, picData.ScaleHeight - cmdFooter.Height
    'dgData.Move 0, 0, picData.ScaleWidth, picData.ScaleHeight - cmdFooter.Height
    cmdFooter.Move 0, lvListView.Height, picData.ScaleWidth, cmdFooter.Height
    cmdNavigate(0).Top = cmdFooter.Top
    cmdNavigate(1).Top = cmdFooter.Top
    cmdNavigate(2).Top = cmdFooter.Top
    cmdNavigate(3).Top = cmdFooter.Top

errHandler: Err.Clear: Resume Next
End Sub

Private Sub pMenu_PanelSelected(oPanel As PhantomPanel.pTab)
    Select Case oPanel.Key
        Case "Students" 'g_eMenu.Patient
            lvListView.ListItems.Clear
            Call initFrame_Students
            Call displayStudents
        Case "DeptStaffs" 'g_eMenu.Fee
            lvListView.ListItems.Clear
            Call initFrame_Staffs
            Call displayStaffs
        Case "Alloted_Assignms"
            lvListView.ListItems.Clear
            Call initFrame_Assignments
            Call displayAssignments
        Case "Todays_Assignms"
            lvListView.ListItems.Clear
            Call initFrame_Assignments
            Call displayTodaysAssignms
        Case "Book_Ref"
            lvListView.ListItems.Clear
            lvListView.ColumnHeaders.Clear
            Call initFrame_BookRef
            Call displayBookRef
            'MsgBox "This facility is not available with the demo version of this application.", vbExclamation, "Demo Version Limitation..."
        Case "Borr_BookRef"
            lvListView.ListItems.Clear
            lvListView.ColumnHeaders.Clear
            Call initFrame_BJTrans
            Call displayBJTrans
            'MsgBox "This facility is not available with the demo version of this application.", vbExclamation, "Demo Version Limitation..."
        Case "IntAssmt_Marks"
            lvListView.ListItems.Clear
            Call initFrame_IntAssmt
            Call displayIntAssmtMarks
        Case "StudFee_Records"
            lvListView.ListItems.Clear
            lvListView.ColumnHeaders.Clear
            Call initFrame_Demo
            Call ShowRecordInfo(rsTemp, cmdFooter)
            MsgBox "This facility is not available with the demo version of this application.", vbExclamation, "Demo Version Limitation..."
        Case "Sys_Rpts" 'g_eMenu.Report
            tvwReport.Nodes.Clear
            Call initTreeview
            MsgBox "Viewing and printing of report(s) is/are not available with the demo version of this" & vbCrLf & "program. You need to use the full version for getting these facilities.", vbExclamation, "Demo Version Limitations..."
    End Select
    Call setToolbar
End Sub
Private Sub initFrame_Students()
    lvListView.ColumnHeaders.Clear
    lvListView.View = lvwReport
    lvListView.FullRowSelect = True
    lvListView.LabelEdit = lvwManual
    lvListView.Sorted = True
    Set z = lvListView.ColumnHeaders.Add(1, , "Student ID", 1000)
    Set z = lvListView.ColumnHeaders.Add(2, , "Student Name", 2500)
    Set z = lvListView.ColumnHeaders.Add(3, , "Date Of Admn", 1500)
    Set z = lvListView.ColumnHeaders.Add(4, , "Class", 2000)
    Set z = lvListView.ColumnHeaders.Add(5, , "Roll No", 1200, 2)
    Set z = lvListView.ColumnHeaders.Add(6, , "Regd. No", 1700)
    Set z = lvListView.ColumnHeaders.Add(7, , "Section", 1200)
    Set z = lvListView.ColumnHeaders.Add(8, , "Acad. Year", 1700)
    Set z = lvListView.ColumnHeaders.Add(9, , "E-Mail ID", 1700)
    Set z = lvListView.ColumnHeaders.Add(10, , "Contact Phone", 1500)
    Set z = lvListView.ColumnHeaders.Add(11, , "Cell Phone", 1500)
    Set z = lvListView.ColumnHeaders.Add(12, , "Contact Address", 3000)
End Sub
Private Sub displayStudents()

    Dim current_academic_yr As String
   '-- Recordsets.
    Set rsTemp = New ADODB.Recordset
    '-- Students
    g_sSQL = "SELECT * FROM qryStudents WHERE AcadYR_Duration='" & CStr(cur_acadYear) & "'"
    rsTemp.Open g_sSQL, g_cn, adOpenDynamic, adLockOptimistic
      
    lvListView.ListItems.Clear
    If Not rsTemp.EOF Or Not rsTemp.BOF Then
        'On Error Resume Next
        'Screen.MousePointer = vbHourglass
        Do While Not rsTemp.EOF
            On Error Resume Next
            Set itemToAdd = lvListView.ListItems.Add(, , rsTemp!StudentID, , 1)
            itemToAdd.SubItems(1) = rsTemp!StudName
            itemToAdd.SubItems(2) = Format(rsTemp!DOA, "dd-MMM-yyyy")
            itemToAdd.SubItems(3) = rsTemp!ClassName
            itemToAdd.SubItems(4) = rsTemp!RollNo
            itemToAdd.SubItems(5) = rsTemp!RegdNo
            itemToAdd.SubItems(6) = rsTemp!Section
            itemToAdd.SubItems(7) = rsTemp!AcadYR_Duration
            itemToAdd.SubItems(8) = rsTemp!EMail
            itemToAdd.SubItems(9) = rsTemp!LPhone
            itemToAdd.SubItems(10) = rsTemp!CPhone
            itemToAdd.SubItems(11) = rsTemp!Address
            rsTemp.MoveNext
        Loop
        DoEvents
        Call ShowRecordInfo(rsTemp, cmdFooter)
        rsTemp.Close
        Set rsTemp = Nothing
    Else
        Call ShowRecordInfo(rsTemp, cmdFooter)
        If rsTemp.State = adStateOpen Then rsTemp.Close
        Set rsTemp = Nothing
    End If
    
End Sub

Private Sub initFrame_Staffs()
    lvListView.ColumnHeaders.Clear
    lvListView.View = lvwReport
    lvListView.FullRowSelect = True
    lvListView.LabelEdit = lvwManual
    lvListView.Sorted = True
    Set z = lvListView.ColumnHeaders.Add(1, , "Staff ID", 1000)
    Set z = lvListView.ColumnHeaders.Add(2, , "Staff Name", 2500)
    Set z = lvListView.ColumnHeaders.Add(3, , "Date Of Joining", 1500)
    Set z = lvListView.ColumnHeaders.Add(4, , "Designation", 1500)
    Set z = lvListView.ColumnHeaders.Add(5, , "Land Ph No", 1500)
    Set z = lvListView.ColumnHeaders.Add(6, , "Cell No", 1500)
    Set z = lvListView.ColumnHeaders.Add(7, , "E-mail Address", 1700)
    Set z = lvListView.ColumnHeaders.Add(8, , "Qualification", 2500)
    Set z = lvListView.ColumnHeaders.Add(9, , "Specialisation", 1700)
    Set z = lvListView.ColumnHeaders.Add(10, , "Contact Adress", 2500)
    Set z = lvListView.ColumnHeaders.Add(11, , "Remarks", 3000)
    Set z = lvListView.ColumnHeaders.Add(12, , "Staff Status", 1700)
End Sub
Private Sub displayStaffs()

   '-- Recordsets.
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
    '-- Patient.
    g_sSQL = "SELECT * FROM Staff"
    rsTemp.Open g_sSQL, g_cn, adOpenDynamic, adLockOptimistic
      
    lvListView.ListItems.Clear
    If Not rsTemp.EOF Or Not rsTemp.BOF Then
        'On Error Resume Next
        'Screen.MousePointer = vbHourglass
        Do While Not rsTemp.EOF
            On Error Resume Next
            Set itemToAdd = lvListView.ListItems.Add(, , rsTemp!StaffID, , 1)
            itemToAdd.SubItems(1) = rsTemp!StaffName
            itemToAdd.SubItems(2) = Format(rsTemp!DOJ, "dd-MMM-yyyy")
            itemToAdd.SubItems(3) = rsTemp!Designation
            itemToAdd.SubItems(4) = rsTemp!Phone_No
            itemToAdd.SubItems(5) = rsTemp!MPhone_No
            itemToAdd.SubItems(6) = rsTemp!EMail
            itemToAdd.SubItems(7) = rsTemp!Qualification
            itemToAdd.SubItems(8) = rsTemp!Specialisation
            itemToAdd.SubItems(9) = rsTemp!Address
            itemToAdd.SubItems(10) = rsTemp!Remarks
            itemToAdd.SubItems(11) = rsTemp!Status
            rsTemp.MoveNext
        Loop
        DoEvents
        Call ShowRecordInfo(rsTemp, cmdFooter)
        rsTemp.Close
        Set rsTemp = Nothing
    Else
        Call ShowRecordInfo(rsTemp, cmdFooter)
        If rsTemp.State = adStateOpen Then rsTemp.Close
        Set rsTemp = Nothing
    End If
    
    '-- Clear Panel.
    'With ucStatus
    '   .PanelText(2) = vbNullString
    '   .PanelText(3) = vbNullString
    'End With
End Sub

Private Sub initFrame_Assignments()
    lvListView.ColumnHeaders.Clear
    lvListView.View = lvwReport
    lvListView.FullRowSelect = True
    lvListView.LabelEdit = lvwManual
    lvListView.Sorted = True
    Set z = lvListView.ColumnHeaders.Add(1, , "Assignmt ID", 1400)
    Set z = lvListView.ColumnHeaders.Add(2, , "Assignment Dated", 1700)
    Set z = lvListView.ColumnHeaders.Add(3, , "Date Of Submission", 1700)
    Set z = lvListView.ColumnHeaders.Add(4, , "Assignment To", 2000)
    Set z = lvListView.ColumnHeaders.Add(5, , "Roll No", 1500)
    Set z = lvListView.ColumnHeaders.Add(6, , "Class", 1700)
    Set z = lvListView.ColumnHeaders.Add(7, , "Assigned By", 2000)
    Set z = lvListView.ColumnHeaders.Add(8, , "Assignment Description", 3500)
    'Set z = lvListView.ColumnHeaders.Add(9, , "Academic Year", 1700)
End Sub
Private Sub displayAssignments()

   '-- Recordsets.
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
    '-- Patient.
    g_sSQL = "SELECT * FROM qryAssignments WHERE AcadYR_Duration='" & CStr(cur_acadYear) & "'"
    rsTemp.Open g_sSQL, g_cn, adOpenDynamic, adLockOptimistic
      
    lvListView.ListItems.Clear
    If Not rsTemp.EOF Or Not rsTemp.BOF Then
        'On Error Resume Next
        'Screen.MousePointer = vbHourglass
        Do While Not rsTemp.EOF
            On Error Resume Next
            Set itemToAdd = lvListView.ListItems.Add(, , rsTemp!AssignID, , 1)
            itemToAdd.SubItems(1) = Format(rsTemp!DOA, "dd-MMM-yyyy")
            itemToAdd.SubItems(2) = Format(rsTemp!DOS, "dd-MMM-yyyy")
            itemToAdd.SubItems(3) = rsTemp!StudName
            itemToAdd.SubItems(4) = rsTemp!RollNo
            itemToAdd.SubItems(5) = rsTemp!ClassName
            itemToAdd.SubItems(6) = rsTemp!StaffName
            itemToAdd.SubItems(7) = rsTemp!Description
            rsTemp.MoveNext
        Loop
        DoEvents
        Call ShowRecordInfo(rsTemp, cmdFooter)
        rsTemp.Close
        Set rsTemp = Nothing
    Else
        Call ShowRecordInfo(rsTemp, cmdFooter)
        If rsTemp.State = adStateOpen Then rsTemp.Close
        Set rsTemp = Nothing
    End If
    
    '-- Clear Panel.
    'With ucStatus
    '   .PanelText(2) = vbNullString
    '   .PanelText(3) = vbNullString
    'End With
End Sub

Private Sub displayTodaysAssignms()

   '-- Recordsets.
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
    '-- Patient.
    g_sSQL = "SELECT * FROM qryAssignments WHERE AcadYR_Duration='" & CStr(cur_acadYear) & "' AND DOS = #" & CDate(Date) & "#"
    rsTemp.Open g_sSQL, g_cn, adOpenDynamic, adLockOptimistic
      
    lvListView.ListItems.Clear
    If Not rsTemp.EOF Or Not rsTemp.BOF Then
        'On Error Resume Next
        'Screen.MousePointer = vbHourglass
        Do While Not rsTemp.EOF
            On Error Resume Next
            Set itemToAdd = lvListView.ListItems.Add(, , rsTemp!AssignID, , 1)
            itemToAdd.SubItems(1) = Format(rsTemp!DOA, "dd-MMM-yyyy")
            itemToAdd.SubItems(2) = Format(rsTemp!DOS, "dd-MMM-yyyy")
            itemToAdd.SubItems(3) = rsTemp!StudName
            itemToAdd.SubItems(4) = rsTemp!RollNo
            itemToAdd.SubItems(5) = rsTemp!ClassName
            itemToAdd.SubItems(6) = rsTemp!StaffName
            itemToAdd.SubItems(7) = rsTemp!Description
            rsTemp.MoveNext
        Loop
        DoEvents
        Call ShowRecordInfo(rsTemp, cmdFooter)
        rsTemp.Close
        Set rsTemp = Nothing
    Else
        Call ShowRecordInfo(rsTemp, cmdFooter)
        If rsTemp.State = adStateOpen Then rsTemp.Close
        Set rsTemp = Nothing
    End If
    
    '-- Clear Panel.
    'With ucStatus
    '   .PanelText(2) = vbNullString
    '  .PanelText(3) = vbNullString
    'End With
End Sub

Private Sub initFrame_Demo()
    lvListView.ColumnHeaders.Clear
    lvListView.View = lvwReport
    lvListView.FullRowSelect = True
    lvListView.LabelEdit = lvwManual
    lvListView.Sorted = True
    Set z = lvListView.ColumnHeaders.Add(1, , "Demo Col# 1", 1500)
    Set z = lvListView.ColumnHeaders.Add(2, , "Demo Col# 2", 1500)
    Set z = lvListView.ColumnHeaders.Add(3, , "Demo Col# 3", 1500)
    Set z = lvListView.ColumnHeaders.Add(4, , "Demo Col# 4", 1500)
    Set z = lvListView.ColumnHeaders.Add(5, , "Demo Col# 5", 1500)
End Sub

Private Sub initFrame_IntAssmt()
    lvListView.ColumnHeaders.Clear
    lvListView.View = lvwReport
    lvListView.FullRowSelect = True
    lvListView.LabelEdit = lvwManual
    lvListView.Sorted = True
    Set z = lvListView.ColumnHeaders.Add(1, , "Assessment ID", 1400)
    Set z = lvListView.ColumnHeaders.Add(2, , "Assmt. Term", 2000)
    Set z = lvListView.ColumnHeaders.Add(3, , "Date Of Assessment", 1700)
    Set z = lvListView.ColumnHeaders.Add(4, , "Student Name", 2000)
    Set z = lvListView.ColumnHeaders.Add(5, , "Roll No", 1400)
    Set z = lvListView.ColumnHeaders.Add(6, , "Class", 1700)
    Set z = lvListView.ColumnHeaders.Add(7, , "Section", 1700)
    Set z = lvListView.ColumnHeaders.Add(8, , "Assessed By", 2000)
    Set z = lvListView.ColumnHeaders.Add(9, , "Marks Obtained", 1400)
    Set z = lvListView.ColumnHeaders.Add(10, , "Remarks", 2500)
End Sub
Private Sub displayIntAssmtMarks()

   '-- Recordsets.
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
    '-- Patient.
    g_sSQL = "SELECT * FROM qryIntAssmt WHERE AcadYR_Duration='" & CStr(cur_acadYear) & "'"
    rsTemp.Open g_sSQL, g_cn, adOpenDynamic, adLockOptimistic
      
    lvListView.ListItems.Clear
    If Not rsTemp.EOF Or Not rsTemp.BOF Then
        'On Error Resume Next
        'Screen.MousePointer = vbHourglass
        Do While Not rsTemp.EOF
            On Error Resume Next
            Set itemToAdd = lvListView.ListItems.Add(, , rsTemp!IntAssmtID, , 1)
            itemToAdd.SubItems(1) = rsTemp!IntAssmt_Term
            itemToAdd.SubItems(2) = Format(rsTemp!DOA, "dd-MMM-yyyy")
            itemToAdd.SubItems(3) = rsTemp!StudName
            itemToAdd.SubItems(4) = rsTemp!RollNo
            itemToAdd.SubItems(5) = rsTemp!ClassName
            itemToAdd.SubItems(6) = rsTemp!Section
            itemToAdd.SubItems(7) = rsTemp!StaffName
            itemToAdd.SubItems(8) = rsTemp!IntAssmt_Marks
            itemToAdd.SubItems(9) = rsTemp!Remarks
            rsTemp.MoveNext
        Loop
        DoEvents
        Call ShowRecordInfo(rsTemp, cmdFooter)
        rsTemp.Close
        Set rsTemp = Nothing
    Else
        Call ShowRecordInfo(rsTemp, cmdFooter)
        If rsTemp.State = adStateOpen Then rsTemp.Close
        Set rsTemp = Nothing
    End If
    
    '-- Clear Panel.
    'With ucStatus
    '   .PanelText(2) = vbNullString
    '   .PanelText(3) = vbNullString
    'End With
End Sub

Private Sub initFrame_BookRef()
    lvListView.ColumnHeaders.Clear
    lvListView.View = lvwReport
    lvListView.FullRowSelect = True
    lvListView.LabelEdit = lvwManual
    lvListView.Sorted = True
    Set z = lvListView.ColumnHeaders.Add(1, , "Book/Ref ID", 1400)
    Set z = lvListView.ColumnHeaders.Add(2, , "Title", 2500)
    Set z = lvListView.ColumnHeaders.Add(3, , "Author(s)", 2500)
    Set z = lvListView.ColumnHeaders.Add(4, , "Publisher", 2500)
    Set z = lvListView.ColumnHeaders.Add(5, , "Pages", 1400, 2)
    Set z = lvListView.ColumnHeaders.Add(6, , "Price Rs/-", 1700, 1)
    Set z = lvListView.ColumnHeaders.Add(7, , "Total Copies", 1700, 2)
    Set z = lvListView.ColumnHeaders.Add(8, , "Copies Available", 1700, 2)
    Set z = lvListView.ColumnHeaders.Add(9, , "Description", 3500)
End Sub
Private Sub displayBookRef()

   '-- Recordsets.
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
    '-- Patient.
    g_sSQL = "SELECT * FROM BookJournel"
    rsTemp.Open g_sSQL, g_cn, adOpenDynamic, adLockOptimistic
      
    lvListView.ListItems.Clear
    If Not rsTemp.EOF Or Not rsTemp.BOF Then
        'On Error Resume Next
        'Screen.MousePointer = vbHourglass
        Do While Not rsTemp.EOF
            On Error Resume Next
            Set itemToAdd = lvListView.ListItems.Add(, , rsTemp!BookJrID, , 1)
            itemToAdd.SubItems(1) = rsTemp!BookJrTitle
            itemToAdd.SubItems(2) = rsTemp!AuthorName
            itemToAdd.SubItems(3) = rsTemp!Publisher
            itemToAdd.SubItems(4) = rsTemp!NOP
            itemToAdd.SubItems(5) = Format(rsTemp!Price, "#,##0.00")
            itemToAdd.SubItems(6) = rsTemp!NOC
            itemToAdd.SubItems(7) = (rsTemp!NOC - rsTemp!CopiesIssued)
            itemToAdd.SubItems(8) = rsTemp!Description
            rsTemp.MoveNext
        Loop
        DoEvents
        Call ShowRecordInfo(rsTemp, cmdFooter)
        rsTemp.Close
        Set rsTemp = Nothing
    Else
        Call ShowRecordInfo(rsTemp, cmdFooter)
        If rsTemp.State = adStateOpen Then rsTemp.Close
        Set rsTemp = Nothing
    End If
    
End Sub


Private Sub initFrame_BJTrans()
    lvListView.ColumnHeaders.Clear
    lvListView.View = lvwReport
    lvListView.FullRowSelect = True
    lvListView.LabelEdit = lvwManual
    lvListView.Sorted = True
    Set z = lvListView.ColumnHeaders.Add(1, , "Trans. ID", 1400)
    Set z = lvListView.ColumnHeaders.Add(2, , "Title", 3500)
    Set z = lvListView.ColumnHeaders.Add(3, , "Author(s)", 3000)
    Set z = lvListView.ColumnHeaders.Add(4, , "Student Name", 2500)
    Set z = lvListView.ColumnHeaders.Add(5, , "Class", 2500)
    Set z = lvListView.ColumnHeaders.Add(6, , "Roll No", 1600, 2)
    Set z = lvListView.ColumnHeaders.Add(7, , "Section", 1600, 2)
    Set z = lvListView.ColumnHeaders.Add(8, , "Date Of Issue", 2700)
    Set z = lvListView.ColumnHeaders.Add(9, , "Date Of Return", 2700)
    Set z = lvListView.ColumnHeaders.Add(10, , "Has Returned ?", 2000)
End Sub
Private Sub displayBJTrans()

   '-- Recordsets.
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
    '-- Patient.
    g_sSQL = "SELECT * FROM qryBJTrans"
    rsTemp.Open g_sSQL, g_cn, adOpenDynamic, adLockOptimistic
      
    lvListView.ListItems.Clear
    If Not rsTemp.EOF Or Not rsTemp.BOF Then
        'On Error Resume Next
        'Screen.MousePointer = vbHourglass
        Do While Not rsTemp.EOF
            On Error Resume Next
            Set itemToAdd = lvListView.ListItems.Add(, , rsTemp!TransID, , 1)
            itemToAdd.SubItems(1) = rsTemp!BookJrTitle
            itemToAdd.SubItems(2) = rsTemp!AuthorName
            itemToAdd.SubItems(3) = rsTemp!StudName
            itemToAdd.SubItems(4) = rsTemp!ClassName
            itemToAdd.SubItems(5) = rsTemp!RollNo
            itemToAdd.SubItems(6) = rsTemp!Section
            itemToAdd.SubItems(7) = FormatDateTime(rsTemp!DOI, vbLongDate)
            itemToAdd.SubItems(8) = FormatDateTime(rsTemp!DOR, vbLongDate)
            If (rsTemp!Returned) Then
                itemToAdd.SubItems(9) = "Yes. Returned back."
            Else
                itemToAdd.SubItems(9) = "No. Yet to return."
            End If
            rsTemp.MoveNext
        Loop
        DoEvents
        Call ShowRecordInfo(rsTemp, cmdFooter)
        rsTemp.Close
        Set rsTemp = Nothing
    Else
        Call ShowRecordInfo(rsTemp, cmdFooter)
        If rsTemp.State = adStateOpen Then rsTemp.Close
        Set rsTemp = Nothing
    End If

End Sub
Private Sub Detect_AcadYRSetting()
Set rsTemp = New ADODB.Recordset
rsTemp.Open "Select * From Settings", g_cn, adOpenKeyset, adLockOptimistic
If Not rsTemp.EOF Or Not rsTemp.BOF Then
    cur_acadYear = rsTemp!AcadYR_Setting
Else
    Set rsTemp = Nothing
    MsgBox "It seems that the data in the database has been tempered with manually." & vbCrLf & _
           "Please uninstall the application then reinstall it again so that the " & vbCrLf & _
           "program can functions properly.", vbCritical, "Corrupt Database..."
    g_cn.Close
    Set g_cn = Nothing
    Unload Me
    End
End If
rsTemp.Close
Set rsTemp = Nothing
End Sub

