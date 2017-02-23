VERSION 5.00
Begin VB.Form frmKlax 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Klax"
   ClientHeight    =   10410
   ClientLeft      =   4065
   ClientTop       =   2175
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10410
   ScaleWidth      =   10935
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   8175
      Left            =   360
      ScaleHeight     =   8115
      ScaleWidth      =   9795
      TabIndex        =   0
      Top             =   1680
      Width           =   9855
      Begin VB.Frame fraInstruct 
         BackColor       =   &H00404040&
         Caption         =   "Instructions"
         ForeColor       =   &H80000014&
         Height          =   7695
         Left            =   360
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   9015
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Thank you for playing!"
            ForeColor       =   &H8000000F&
            Height          =   2415
            Left            =   1320
            TabIndex        =   18
            Top             =   360
            Width           =   6255
         End
      End
      Begin VB.Timer tmrScore 
         Enabled         =   0   'False
         Interval        =   150
         Left            =   9360
         Top             =   3600
      End
      Begin VB.Timer tmrKlax 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   8880
         Top             =   3600
      End
      Begin VB.Frame fraLose 
         Height          =   3495
         Left            =   960
         TabIndex        =   6
         Top             =   720
         Visible         =   0   'False
         Width           =   7695
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "You Lose"
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   24
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1800
            TabIndex        =   7
            Top             =   360
            Width           =   4095
         End
      End
      Begin VB.Timer tmrA 
         Interval        =   50
         Left            =   9360
         Top             =   4080
      End
      Begin VB.Timer tmrTest 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   8880
         Top             =   4080
      End
      Begin VB.Label lblWave 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000012&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   9120
         TabIndex        =   16
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label lblTotalPoints 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   8280
         TabIndex        =   15
         Top             =   7200
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000012&
         Caption         =   "1x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   8280
         TabIndex        =   14
         Top             =   6600
         Width           =   375
      End
      Begin VB.Label lblGain 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   8760
         TabIndex        =   13
         Top             =   6600
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000012&
         Caption         =   "Points Gained:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   8280
         TabIndex        =   12
         Top             =   6240
         Width           =   1335
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H80000001&
         FillColor       =   &H80000001&
         Height          =   1935
         Left            =   8160
         Top             =   6120
         Width           =   1575
      End
      Begin VB.Label lblKlax 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   8280
         TabIndex        =   11
         Top             =   5400
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000012&
         Caption         =   "Klaxes Left:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   8280
         TabIndex        =   10
         Top             =   5040
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000012&
         Caption         =   "Wave"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   8280
         TabIndex        =   9
         Top             =   4560
         Width           =   735
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H80000001&
         FillColor       =   &H80000001&
         Height          =   1215
         Left            =   8160
         Top             =   4920
         Width           =   1575
      End
      Begin VB.Image imgWh 
         Height          =   960
         Left            =   8400
         Picture         =   "frmKlax.frx":0000
         Stretch         =   -1  'True
         Top             =   2520
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Image imgtesT 
         Height          =   105
         Left            =   7680
         Stretch         =   -1  'True
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Image imgCy 
         Height          =   960
         Left            =   120
         Picture         =   "frmKlax.frx":082A
         Stretch         =   -1  'True
         Top             =   3000
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Image imgStorage 
         Height          =   960
         Index           =   0
         Left            =   8400
         Picture         =   "frmKlax.frx":1054
         Stretch         =   -1  'True
         Top             =   1440
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Shape shpStack 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   1
         Left            =   4080
         Top             =   5040
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shpHold 
         BorderColor     =   &H00808080&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   4200
         Top             =   5350
         Width           =   1095
      End
      Begin VB.Shape shpPlayer 
         BackColor       =   &H80000006&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   3960
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Image imgPu 
         Height          =   960
         Left            =   120
         Picture         =   "frmKlax.frx":187E
         Stretch         =   -1  'True
         Top             =   1080
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Image imgGrn 
         Height          =   960
         Left            =   120
         Picture         =   "frmKlax.frx":20A8
         Stretch         =   -1  'True
         Top             =   120
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Image imgO 
         Height          =   960
         Left            =   120
         Picture         =   "frmKlax.frx":28D2
         Stretch         =   -1  'True
         Top             =   2040
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Image imgR2 
         Height          =   105
         Index           =   0
         Left            =   5100
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Image imgL2 
         Height          =   105
         Index           =   0
         Left            =   4065
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Image imgPnk 
         Height          =   960
         Left            =   1500
         Picture         =   "frmKlax.frx":30FC
         Stretch         =   -1  'True
         Top             =   3960
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Image imgL1 
         Height          =   105
         Index           =   0
         Left            =   4330
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Image imgGry 
         Height          =   960
         Left            =   2885
         Picture         =   "frmKlax.frx":3926
         Stretch         =   -1  'True
         Top             =   3960
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Image imgR 
         Height          =   960
         Left            =   4150
         Picture         =   "frmKlax.frx":4150
         Stretch         =   -1  'True
         Top             =   3960
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Image imgC 
         Height          =   105
         Index           =   0
         Left            =   4590
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Line lnPath 
         BorderColor     =   &H80000000&
         Index           =   1
         Visible         =   0   'False
         X1              =   4920
         X2              =   6000
         Y1              =   0
         Y2              =   4440
      End
      Begin VB.Image imgR1 
         Height          =   105
         Index           =   0
         Left            =   4850
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Image imgB 
         Height          =   960
         Left            =   5440
         Picture         =   "frmKlax.frx":497A
         Stretch         =   -1  'True
         Top             =   3960
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Line lnPath 
         BorderColor     =   &H80000000&
         Index           =   0
         Visible         =   0   'False
         X1              =   5190
         X2              =   7387
         Y1              =   0
         Y2              =   4920
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000001&
         X1              =   2760
         X2              =   4320
         Y1              =   4920
         Y2              =   0
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000001&
         X1              =   4080
         X2              =   4560
         Y1              =   4920
         Y2              =   0
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000001&
         X1              =   5400
         X2              =   4800
         Y1              =   4920
         Y2              =   0
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000001&
         X1              =   6720
         X2              =   5040
         Y1              =   4920
         Y2              =   0
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000001&
         X1              =   8055
         X2              =   5280
         Y1              =   4920
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000001&
         X1              =   1440
         X2              =   4080
         Y1              =   4920
         Y2              =   0
      End
      Begin VB.Line lnBreak 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         X1              =   1440
         X2              =   8055
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Image imgY 
         Height          =   960
         Left            =   6840
         Picture         =   "frmKlax.frx":51A4
         Stretch         =   -1  'True
         Top             =   3960
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Shape shpCol 
         BorderColor     =   &H00404040&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   5
         Left            =   6840
         Top             =   7740
         Width           =   1095
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H80000006&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   6720
         Top             =   7680
         Width           =   1335
      End
      Begin VB.Shape shpCol 
         BorderColor     =   &H00404040&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   4
         Left            =   5520
         Top             =   7740
         Width           =   1095
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H80000006&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   5400
         Top             =   7680
         Width           =   1335
      End
      Begin VB.Shape shpCol 
         BorderColor     =   &H00404040&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   3
         Left            =   4200
         Top             =   7740
         Width           =   1095
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H80000006&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   4080
         Top             =   7680
         Width           =   1335
      End
      Begin VB.Shape shpCol 
         BorderColor     =   &H00404040&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   2
         Left            =   2880
         Top             =   7740
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000006&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   2760
         Top             =   7680
         Width           =   1335
      End
      Begin VB.Shape shpCol 
         BorderColor     =   &H00404040&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   1
         Left            =   1560
         Top             =   7740
         Width           =   1095
      End
      Begin VB.Shape shpCol 
         BorderColor     =   &H00808080&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   8400
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000006&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   1440
         Top             =   7680
         Width           =   1335
      End
      Begin VB.Shape shpStack 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   8280
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shpPlace 
         BorderColor     =   &H80000007&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   25
         Left            =   6720
         Top             =   6480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shpPlace 
         BorderColor     =   &H80000007&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   24
         Left            =   6720
         Top             =   6720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shpPlace 
         BorderColor     =   &H80000007&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   23
         Left            =   6720
         Top             =   6960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shpPlace 
         BorderColor     =   &H80000007&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   22
         Left            =   6720
         Top             =   7200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shpPlace 
         BorderColor     =   &H80000007&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   21
         Left            =   6720
         Top             =   7440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shpPlace 
         BorderColor     =   &H80000007&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   20
         Left            =   5400
         Top             =   6480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shpPlace 
         BorderColor     =   &H80000007&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   19
         Left            =   5400
         Top             =   6720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shpPlace 
         BorderColor     =   &H80000007&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   18
         Left            =   5400
         Top             =   6960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shpPlace 
         BorderColor     =   &H80000007&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   17
         Left            =   5400
         Top             =   7200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shpPlace 
         BorderColor     =   &H80000007&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   16
         Left            =   5400
         Top             =   7440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shpPlace 
         BorderColor     =   &H80000007&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   15
         Left            =   4080
         Top             =   6480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shpPlace 
         BorderColor     =   &H80000007&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   14
         Left            =   4080
         Top             =   6720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shpPlace 
         BorderColor     =   &H80000007&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   13
         Left            =   4080
         Top             =   6960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shpPlace 
         BorderColor     =   &H80000007&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   12
         Left            =   4080
         Top             =   7200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shpPlace 
         BorderColor     =   &H80000007&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   11
         Left            =   4080
         Top             =   7440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shpPlace 
         BorderColor     =   &H80000007&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   10
         Left            =   2760
         Top             =   6480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shpPlace 
         BorderColor     =   &H80000007&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   9
         Left            =   2760
         Top             =   6720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shpPlace 
         BorderColor     =   &H80000007&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   8
         Left            =   2760
         Top             =   6960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shpPlace 
         BorderColor     =   &H80000007&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   7
         Left            =   2760
         Top             =   7200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shpPlace 
         BorderColor     =   &H80000007&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   6
         Left            =   2760
         Top             =   7440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shpPlace 
         BorderColor     =   &H80000007&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   5
         Left            =   1440
         Top             =   6480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shpPlace 
         BorderColor     =   &H80000007&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   4
         Left            =   1440
         Top             =   6720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shpPlace 
         BorderColor     =   &H80000007&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   3
         Left            =   1440
         Top             =   6960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shpPlace 
         BorderColor     =   &H80000007&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   2
         Left            =   1440
         Top             =   7200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shpPlace 
         BorderColor     =   &H80000007&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   1
         Left            =   1440
         Top             =   7440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape shpPlace 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   8280
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         Caption         =   "Drop Meter"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   0
         TabIndex        =   8
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Shape shpDrop 
         BorderColor     =   &H8000000B&
         FillStyle       =   0  'Solid
         Height          =   495
         Index           =   4
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   7560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Shape shpDrop 
         BorderColor     =   &H8000000B&
         FillStyle       =   0  'Solid
         Height          =   495
         Index           =   3
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   6960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Shape shpDrop 
         BorderColor     =   &H8000000B&
         FillStyle       =   0  'Solid
         Height          =   495
         Index           =   2
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   6360
         Width           =   1215
      End
      Begin VB.Shape shpDrop 
         BorderColor     =   &H8000000B&
         FillStyle       =   0  'Solid
         Height          =   495
         Index           =   1
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   5760
         Width           =   1215
      End
      Begin VB.Shape shpDrop 
         BorderColor     =   &H8000000B&
         FillStyle       =   0  'Solid
         Height          =   495
         Index           =   0
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Shape Shape6 
         FillColor       =   &H80000001&
         FillStyle       =   6  'Cross
         Height          =   3135
         Left            =   0
         Top             =   4920
         Width           =   1455
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "Player Score:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "KLAX"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1095
      Left            =   3600
      TabIndex        =   3
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label lblHighScore 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "32500"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   7920
      TabIndex        =   2
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "High Score:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   6840
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "frmKlax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'There are 20 Spaces, 5 Col
'Quintin Reed
'Computer Science 3
Dim location As Integer
Dim intervals As Integer
Dim seconds As Single
Dim fast As Boolean

Dim stacknum As Integer
Dim blockclr(10) As ColorConstants
Dim caught(5) As Boolean
Dim generateInt As Integer

'counters
Dim leftest As Integer  'counter for blocks on L2
Dim lefter As Integer   'counter for blocks on L1
Dim centerer As Integer 'counter for blocks on C
Dim righter As Integer  'counter for blocks on R1
Dim rightest As Integer 'counter for blocks on R2

Dim prevl2 As Integer  'counter for last level blocks on L2
Dim prevl1 As Integer   'counter for last level blocks on L1
Dim prevc As Integer 'counter for last level blocks on C
Dim prevr1 As Integer  'counter for last level blocks on R1
Dim prevr2 As Integer 'counter for last level blocks on R2

Dim level As Integer 'Level of Difficulty
Dim destroyed(5) As Integer '1 - 5 denotes where on board

Dim create As Integer 'Creation setting as to when to make a new one

Dim caughtnum As Integer 'caught level
Dim placenum(5) As Integer 'Amount placed on each

'for destoryed placed blocks
Dim placedestroy(25) As Integer
Dim destroyvar As Integer
Dim flash(8) As Boolean
Dim flasher(8) As Integer
Dim flashconst(8) As Integer 'flash count that keeps track of how many blocks are using the timer
                             ' designed for having multiple timers going at the same time
Dim flashingcount As Integer 'integer value of which timers are being used starting at 0 to 8
Dim flashingmin As Integer 'integer denoting the beginnign variable of for loop

Dim dropmeter As Integer 'measures drops

Dim destroyloop(25) As Integer 'count of which spaces have been destroyed as to avoid subtracting too much
Dim destroyplus As Integer 'counts the amount of blocks destroyed by klax

Dim tmrdone As Boolean

Dim klaxcount As Integer 'counts klaxes for level design
Dim levelklax As Integer 'klaxes for level/or points

Dim multiplier As Integer

Dim beltdone As Boolean 'is timer done counting belt tiles
Dim bindone As Boolean 'is timer done counting empty spaces



                            



Private Sub Form_Load()
seconds = 7.75 ' there is a two second delay in the interval movement
            'each level will decrease the seconds slighlty
            'factor of .2 maybe?
            'need tag system to implement colors accordingly
location = 3
fast = False
Dim X As Integer

beltdone = False
bindone = False

dropmeter = 0

imgR2(0) = imgY
imgR1(0) = imgB
imgC(0) = imgR
imgL1(0) = imgGry
imgL2(0) = imgPnk

destroyplus = 0
level = 1
stacknum = 1
create = 0

tmrdone = True
caught(0) = False
'colors
blockclr(0) = RGB(247, 73, 65) 'Red
blockclr(1) = RGB(249, 135, 23) 'Orange
blockclr(2) = RGB(229, 194, 0) 'Yellow
blockclr(3) = RGB(34, 250, 66) 'Green
blockclr(4) = RGB(191, 191, 191) 'Gray
blockclr(5) = RGB(243, 171, 190) 'Pink
blockclr(6) = RGB(227, 5, 245) 'Purple
blockclr(7) = RGB(73, 125, 248) 'Blue
blockclr(8) = RGB(9, 243, 248) 'Cyan


For X = 1 To 10
    Load imgStorage(X)
Next X
For X = 0 To 8
    flasher(X) = 0
    flashconst(X) = 0
    flash(X) = 0

Next X
flashingcount = 0

imgStorage(0) = imgR
imgStorage(1) = imgO
imgStorage(2) = imgY
imgStorage(3) = imgGrn
imgStorage(4) = imgGry
imgStorage(5) = imgPnk
imgStorage(6) = imgPu
imgStorage(7) = imgB
imgStorage(8) = imgCy

'shpStack(1).FillColor = blockclr(8)
'shpStack(1).Visible = True
For X = 2 To 5 ' load other stack blocks
    Load shpStack(X)
    shpStack(X).Visible = False
    shpStack(X).Top = shpStack(X - 1).Top - shpStack(0).Height
    shpStack(X).Left = shpStack(X - 1).Left
Next X
'Load imgR1(1)
'imgR1(1) = imgR1(0)
'imgR1(1).Visible = True
'

For X = 1 To 25
    shpPlace(X).Tag = 300
Next X

prevl2 = 0
prevl1 = 0
prevc = 0
prevr1 = 0
prevr2 = 0

leftest = 0
lefter = 0
centerer = 0
righter = 0
rightest = 0

flashingmin = 1

levelklax = 3


multiplier = 0 ' add one for every klax

destroyvar = 0 ' a variable set to maintiain count of destroyed placed blocks that form klaxes

For X = 0 To 2
    If X = 0 Then
        shpDrop(X).Top = Shape6.Top + 100
        shpDrop(X).Height = (Shape6.Height - 400) / 3
    Else
        shpDrop(X).Top = shpDrop(X - 1).Top + shpDrop(X - 1).Height + 100
        shpDrop(X).Height = shpDrop(X - 1).Height
    End If
    'shpDrop(x).Visible = True
Next X

intervals = seconds * 1000 / 50 'The amount of times the timer goes through, for 10 seconds, 50 intervals

End Sub


Private Sub ScoreChange()

    lblTotalPoints = Val(lblGain) * multiplier

End Sub

Private Sub totalscore()
    
    lblScore = Val(lblScore) + Val(lblTotalPoints)
    
End Sub


Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim inte As Integer
Dim X As Integer
inte = KeyCode

If inte = 37 And location > 1 Then
    shpPlayer.Left = shpPlace((location - 1) * 5).Left - 120
    shpHold.Left = shpPlace((location - 1) * 5).Left + 120
    For X = 1 To stacknum - 1
        shpStack(X).Left = shpPlace((location - 1) * 5).Left
    Next X
    location = location - 1
End If
'If inte = 38 And stacknum > 1 Then
'    Generate location
'    If location = 1 Then
'        imgL2(leftest - 1).Left = Int((imgPnk.Left - 4065) / 2) + 4065
'        imgL2(leftest - 1).Top = Int((imgPnk.Top) / 2) '38
'        imgL2(leftest - 1).Height = Int((imgPnk.Height - 105) / 2) + 105
'        imgL2(leftest - 1).Width = Int((imgPnk.Width - 195) / 2) + 195
'        imgL2(leftest - 1) = imgStorage(shpStack(stacknum - 1).Tag)
'        imgL2(leftest - 1).Tag = shpStack(stacknum - 1).Tag
'    End If
'    If location = 2 Then
'
'
'    End If
'    If location = 3 Then
'
'
'    End If
'    If location = 4 Then
'
'
'    End If
'    If location = 5 Then
'
'
'    End If
'    stacknum = stacknum - 1
'End If
If inte = 39 And location < 5 Then
    shpPlayer.Left = shpPlace((location + 1) * 5).Left - 120
    shpHold.Left = shpPlace((location + 1) * 5).Left + 120
    For X = 1 To stacknum - 1
        shpStack(X).Left = shpPlace((location + 1) * 5).Left
    Next X
    location = location + 1
End If
If inte = 40 And fast = False Then
'    seconds = seconds / 4
'    intervals = seconds * 1000 / tmrA.Interval
    tmrA.Interval = tmrA.Interval / 4
    fast = True
End If

End Sub

Private Sub Level_end()

'the end of level code

If klaxcount = levelklax Then
    Picture1.Enabled = False
    If tmrdone = True Then
        For X = 1 To 25
            'shpPlace(x).Visible = true
            If shpPlace(X).FillColor = vbWhite Then shpPlace(X).FillColor = vbBlack: shpPlace(X).BorderColor = vbRed
        Next X
        tmrScore.Interval = 150
        tmrScore.Enabled = True
    End If
End If


End Sub

Private Sub Reset()

seconds = 7.75 - 0.05 * (level + 1) ' there is a two second delay in the interval movement
            'each level will decrease the seconds slighlty
            'factor of .2 maybe?
            'need tag system to implement colors accordingly
fast = False

Dim y As Integer
Dim X As Integer

beltdone = False
bindone = False

dropmeter = 0

destroyplus = 0
level = level + 1
stacknum = 1
create = 0

tmrdone = True
caught(0) = False

For X = 0 To 8
    flasher(X) = 0
    flashconst(X) = 0
    flash(X) = 0
Next X

flashingcount = 0

For X = 1 To 25
    shpPlace(X).Tag = 300
Next X


flashingmin = 1

klaxcount = 0

levelklax = 3 + 2

lblKlax = levelklax

multiplier = 0 ' add one for every klax

destroyvar = 0 ' a variable set to maintiain count of destroyed placed blocks that form klaxes

intervals = seconds * 1000 / 50 'The amount of times the timer goes through, for 10 seconds, 50 intervals

tmrA.Interval = 50

level = level + 1

For X = 1 To 25
    shpPlace(X).FillColor = vbWhite
    shpPlace(X).BorderColor = vbBlack
    shpPlace(X).Visible = False
Next X

shpPlayer.Top = 5160
shpHold.Top = 5350

Label7 = "Klaxes Left:"

'reset the blocks too

For X = 0 To leftest - 1
    imgL2(X).Top = 0
    imgL2(X).Left = 4065
   ' imgL2(X).Visible = True
    imgL2(X).Height = 105
    imgL2(X).Width = 195

Next X
For X = 0 To lefter - 1
    imgL1(X).Top = 0
    imgL1(X).Left = 4330
    'imgL1(X).Visible = True
    imgL1(X).Height = 105
    imgL1(X).Width = 195

Next X
For X = 0 To centerer - 1
    imgC(X).Top = 0
    imgC(X).Left = 4590
   ' imgC(X).Visible = True
    imgC(X).Height = 105
    imgC(X).Width = 195

Next X
For X = 0 To righter - 1

    imgR1(X).Top = 0
    imgR1(X).Left = 4850
    'imgR1(X).Visible = True
    imgR1(X).Height = 105
    imgR1(X).Width = 195
    
Next X
For X = 0 To rightest - 1

    imgR2(X).Top = 0
    imgR2(X).Left = 5100
    imgR2(X).Height = 105
    imgR2(X).Width = 195
   ' imgR2(X).Visible = True
   
Next X

If prevl2 < leftest Then prevl2 = leftest
If prevl1 < lefter Then prevl1 = lefter
If prevc < centerer Then prevc = centerer
If prevr1 < righter Then prevr1 = righter
If prevr2 < rightest Then prevr2 = rightest

leftest = 0
lefter = 0
centerer = 0
righter = 0
rightest = 0

For y = 0 To 5
    placenum(y) = 0
    destroyed(y) = 0
    caught(y) = False
Next y


Picture1.Enabled = True

tmrA.Enabled = True


End Sub

Private Sub Filled()
Dim X As Integer
Dim counter As Integer

'for losing

For X = 1 To 25
    If shpPlace(X).Visible = True Then counter = counter + 1
Next X

If counter = 25 Then
    fraLose.Visible = True
    Picture1.Enabled = False
End If

For y = 1 To 5
    counter = 0
    For X = 1 + (5 * (y - 1)) To 5 * y
        If shpPlace(X).Visible = True Then counter = counter + 1
    Next X
    If counter = 5 Then
        shpCol(y).FillColor = vbRed
    End If
Next y



End Sub

Private Sub Klax()

'to check for a klax (3 or more in a row, column, or diagonal
Dim X As Integer
Dim found As Boolean
Dim n As Integer
Dim a As Integer
Dim b As Integer
'found = False
Dim z As Integer
Dim flashholder As Integer
Dim newklax As Boolean
newklax = False

lblGain = 0

For z = 0 To 8
    flashconst(z) = 0
Next z
For z = 0 To 24
    destroyloop(z) = -1
Next z
For X = 1 To 25

If flashingmin = 1 Then
    flashholder = flashingmin
Else
    flashholder = flashingmin - 1
End If

destroyplus = 0
    found = False
'    Do
        For y = 0 To destroyvar - 1
            If placedestroy(y) = X Then GoTo nextplacE:  'And newklax = True
        '**************************************************************Duplicate destroyvars showing up: Fixed
        Next y
        'column
        If (X + 2) Mod 5 >= 3 Mod 5 Or (X + 2) Mod 5 = 0 Then
            n = shpPlace(X).Tag
            a = shpPlace(X + 1).Tag
            b = shpPlace(X + 2).Tag
            
            If shpPlace(X).Tag = shpPlace(X + 1).Tag And shpPlace(X + 2).Tag = shpPlace(X).Tag Then
                If shpPlace(X).Tag <> 300 And shpPlace(X).Visible = True And shpPlace(X + 1).Visible = True And shpPlace(X + 2).Visible = True Then
                    If found = False Then flashingcount = flashingcount + 1
                    tmrKlax.Enabled = True
                    tmrA.Enabled = False
                    placedestroy(destroyvar) = X
                    placedestroy(destroyvar + 1) = X + 1
                    placedestroy(destroyvar + 2) = X + 2
                    destroyvar = destroyvar + 3
                    tmrdone = False
                    flashconst(flashholder) = flashconst(flashholder) + 3
                    newklax = True
                    multiplier = multiplier + 1
                    lblGain = Val(lblGain) + 50
                    klaxcount = klaxcount + 1
                    lblKlax = levelklax - klaxcount
                    If (X + 3) Mod 5 >= 4 Mod 5 Or (X + 3) Mod 5 = 0 Then
                        If shpPlace(X + 3).Tag = shpPlace(X).Tag And shpPlace(X + 3).Visible = True Then
                            placedestroy(destroyvar) = X + 3
                            destroyvar = destroyvar + 1
                            flashconst(flashholder) = flashconst(flashholder) + 1
                            
                            If (X + 4) Mod 5 = 0 Then
                                If shpPlace(X + 4).Tag = shpPlace(X).Tag And shpPlace(X + 4).Visible = True Then
                                    placedestroy(destroyvar) = X + 4
                                    destroyvar = destroyvar + 1
                                    flashconst(flashholder) = flashconst(flashholder) + 1
                                End If
                            End If
                            
                        End If
                    End If
                    If flashingmin > 1 Then flashingmin = flashingmin - 1

'                    denoting the use of flasher(flashingcount) with the amount of blocks being flashconst
'                    shpPlace(x).Tag = 300
'                    shpPlace(x + 1).Tag = shpPlace(x).Tag
'                    shpPlace(x + 2).Tag = shpPlace(x).Tag
'                    If destroyvar > 3 Then
'                        flasher = 3
'                    End If
                    
                End If
                
            Else
            
            End If
        End If
            'horizontal
        If X + 10 < 26 Then
            n = shpPlace(X).Tag
            a = shpPlace(X + 5).Tag
            b = shpPlace(X + 10).Tag
            
            If shpPlace(X).Tag = shpPlace(X + 5).Tag And shpPlace(X).Tag = shpPlace(X + 10).Tag Then
                If shpPlace(X).Tag <> 300 And shpPlace(X).Visible = True And shpPlace(X + 5).Visible = True And shpPlace(X + 10).Visible = True Then
                    If found = False Then flashingcount = flashingcount + 1
                    tmrKlax.Enabled = True
                    tmrA.Enabled = False
                    placedestroy(destroyvar) = X
                    placedestroy(destroyvar + 1) = X + 5
                    placedestroy(destroyvar + 2) = X + 10
                    destroyvar = destroyvar + 3
                    tmrdone = False
                    multiplier = multiplier + 1
                    flashconst(flashingcount) = flashconst(flashingcount) + 3
                    newklax = True
                    lblGain = Val(lblGain) + 1000
                    klaxcount = klaxcount + 1
                    lblKlax = levelklax - klaxcount
                    If (X + 15) < 26 Then
                    
                        If shpPlace(X + 15).Tag = shpPlace(X).Tag And shpPlace(X + 15).Visible = True Then
                            placedestroy(destroyvar) = X + 15
                            destroyvar = destroyvar + 1
                            flashconst(flashholder) = flashconst(flashholder) + 1
                            
                            If (X + 20) < 26 Then
                                If shpPlace(X + 20).Tag = shpPlace(X).Tag And shpPlace(X + 20).Visible = True Then
                                    placedestroy(destroyvar) = X + 20
                                    destroyvar = destroyvar + 1
                                    flashconst(flashholder) = flashconst(flashholder) + 1
                                End If
                            End If
                            
                        End If
                    End If
                    found = True
                    If flashingmin > 1 Then flashingmin = flashingmin - 1

'                    shpPlace(x).Tag = 300
'                    shpPlace(x + 5).Tag = shpPlace(x).Tag
'                    shpPlace(x + 10).Tag = shpPlace(x).Tag
                End If
            Else
            
            End If
        End If
            'right diagonal
        If ((X + 12) Mod 5 >= 3 Or (X + 12) Mod 5 = 0) And X + 12 < 26 Then
'            n = shpPlace(x).Tag
'            a = shpPlace(x + 6).Tag
'            b = shpPlace(x + 12).Tag
            
            If shpPlace(X).Tag = shpPlace(X + 6).Tag And shpPlace(X).Tag = shpPlace(X + 12).Tag Then
                If shpPlace(X).Tag <> 300 And shpPlace(X).Visible = True And shpPlace(X + 6).Visible = True And shpPlace(X + 12).Visible = True Then
                    If found = False Then flashingcount = flashingcount + 1
                    tmrKlax.Enabled = True
                    tmrA.Enabled = False
                    placedestroy(destroyvar) = X
                    placedestroy(destroyvar + 1) = X + 6
                    placedestroy(destroyvar + 2) = X + 12
                    destroyvar = destroyvar + 3
                    tmrdone = False
                    multiplier = multiplier + 1
                    flashconst(flashingcount) = flashconst(flashingcount) + 3
                    newklax = True
                    lblGain = Val(lblGain) + 5000
                    klaxcount = klaxcount + 1
                    lblKlax = levelklax - klaxcount
                    If ((X + 18) Mod 5 >= 3 Or (X + 18) Mod 5 = 0) And X + 18 < 26 Then
                        If shpPlace(X + 18).Tag = shpPlace(X).Tag And shpPlace(X + 18).Visible = True Then
                            placedestroy(destroyvar) = X + 18
                            destroyvar = destroyvar + 1
                            flashconst(flashholder) = flashconst(flashholder) + 1

                            If (X + 24) Mod 5 = 0 And X + 24 < 26 Then
                                If shpPlace(X + 24).Tag = shpPlace(X).Tag And shpPlace(X + 24).Visible = True Then
                                    placedestroy(destroyvar) = X + 24
                                    destroyvar = destroyvar + 1
                                    flashconst(flashholder) = flashconst(flashholder) + 1
                                End If
                            End If
                            
                        End If
                    End If
                    
                    
                    found = True
                    If flashingmin > 1 Then flashingmin = flashingmin - 1

'                    shpPlace(x).Tag = 300
'                    shpPlace(x + 5).Tag = shpPlace(x).Tag
'                    shpPlace(x + 10).Tag = shpPlace(x).Tag
                End If
            Else
            
            End If
        End If
            'left diagonal
        If ((X - 8) Mod 5 >= 3 Or (X - 8) Mod 5 = 0) And X - 8 > 1 Then
'            n = shpPlace(x).Tag
'            a = shpPlace(x + 6).Tag
'            b = shpPlace(x + 12).Tag
            
            If shpPlace(X).Tag = shpPlace(X - 4).Tag And shpPlace(X).Tag = shpPlace(X - 8).Tag Then
                If shpPlace(X).Tag <> 300 And shpPlace(X).Visible = True And shpPlace(X - 4).Visible = True And shpPlace(X - 8).Visible = True Then
                    If found = False Then flashingcount = flashingcount + 1
                    tmrKlax.Enabled = True
                    tmrA.Enabled = False
                    placedestroy(destroyvar) = X
                    placedestroy(destroyvar + 1) = X - 4
                    placedestroy(destroyvar + 2) = X - 8
                    destroyvar = destroyvar + 3
                    flashconst(flashingcount) = flashconst(flashingcount) + 3
                    tmrdone = False
                    newklax = True
                    multiplier = multiplier + 1
                    lblGain = Val(lblGain) + 5000
                    klaxcount = klaxcount + 1
                    lblKlax = levelklax - klaxcount
                    If ((X - 12) Mod 5 >= 3 Or (X - 12) Mod 5 = 0) And X - 12 > 1 Then
                        If shpPlace(X - 12).Tag = shpPlace(X).Tag And shpPlace(X - 12).Visible = True Then
                            placedestroy(destroyvar) = X - 12
                            destroyvar = destroyvar + 1
                            flashconst(flashholder) = flashconst(flashholder) + 1
                            
                                If (X - 16) Mod 5 = 0 And X - 16 > 1 Then
                                    If shpPlace(X - 16).Tag = shpPlace(X).Tag And shpPlace(X - 16).Visible = True Then
                                        placedestroy(destroyvar) = X - 16
                                        destroyvar = destroyvar + 1
                                        flashconst(flashholder) = flashconst(flashholder) + 1
                                    End If
                                End If
                                
                        End If
                    End If
                    
                    found = True
                    If flashingmin > 1 Then flashingmin = flashingmin - 1

'                    shpPlace(x).Tag = 300
'                    shpPlace(x + 5).Tag = shpPlace(x).Tag
'                    shpPlace(x + 10).Tag = shpPlace(x).Tag
                End If
            Else
            
            End If
        End If

nextplacE:
  '  Loop Until (found = True)
  
Next X

ScoreChange

totalscore

Level_end


End Sub

Private Sub tmrKlax_Timer()
Dim X As Integer
Dim locate As Integer
Dim y As Integer
Dim n As Integer
Dim m As Integer

If 1 = 2 Then
    'The World will explode
    End
End If
Dim z As Integer

For m = 0 To flashconst(z) - 1
    destroyloop(m) = -1
Next m

For z = flashingmin To flashingcount

    If flasher(z) < 11 Then
        'Flash white for 2 seconds
        If flash(z) = False Then
            For X = flashconst(z - 1) To flashconst(z) + flashconst(z - 1) - 1
                shpPlace(placedestroy(X)).FillColor = vbWhite
            Next X
            flash(z) = True: GoTo endofTimeR:
        End If
        If flash(z) = True Then
            For X = flashconst(z - 1) To flashconst(z) + flashconst(z - 1) - 1
                n = shpPlace(placedestroy(X)).Tag
                shpPlace(placedestroy(X)).FillColor = blockclr(shpPlace(placedestroy(X)).Tag)
            Next X
            flash(z) = False
        End If
endofTimeR:
        flasher(z) = flasher(z) + 1
    End If
    n = 0
    If flasher(z) = 11 Then
        '(5 * (location - 1)) + 1 + placenum(location)
        For X = flashconst(z) - 1 + destroyplus To destroyplus Step -1
            For m = 0 To n
                    If placedestroy(X) = destroyloop(m) Then GoTo getnextX:
            Next m
                shpPlace(placedestroy(X)).Tag = 300
                shpPlace(placedestroy(X)).FillColor = vbWhite
                shpPlace(placedestroy(X)).Visible = False
                locate = Int((placedestroy(X) - 1) / 5) + 1
                placenum(locate) = placenum(locate) - 1
                destroyloop(n) = placedestroy(X)
                n = n + 1
                For y = placedestroy(X) To (locate) * 5 - 1
                    If locate * 5 <> y Then
                        shpPlace(y).FillColor = shpPlace(y + 1).FillColor
                        shpPlace(y).Tag = shpPlace(y + 1).Tag
'                        shpPlace(y + 1).Tag = 300
'                        shpPlace(y + 1).FillColor = vbWhite
'                        shpPlace(y + 1).Visible = False
                        shpPlace(y).Visible = True
                    End If
                    If shpPlace(y).FillColor = vbWhite Then shpPlace(y).Visible = False
                Next y
                shpCol(locate).FillColor = vbGreen
getnextX:
        Next X
        flasher(z) = 0
        
        flash(z) = False
        destroyvar = destroyvar - flashconst(z)
        'flashconst(z) = 0
        destroyplus = destroyplus + flashconst(z)
        flashingmin = flashingmin + 1
        If flashingcount < flashingmin Then
            For X = 0 To destroyvar - 1
                placedestroy(X) = -1
            Next X
            flashingcount = 0
            tmrA.Enabled = True
            flashingmin = 1
            tmrdone = True
            Klax
            If flashingcount = 0 Then
                multiplier = 0
                If Picture1.Enabled = False Then tmrA.Enabled = False: Level_end
                tmrKlax.Enabled = False
            End If
        Else
'            Klax
        End If
    End If
Next z
End Sub



Private Sub Picture1_KeyPress(KeyAscii As Integer)
Dim inte As Integer
Dim y As Integer
inte = KeyAscii

If (inte = 120 Or inte = 88) And stacknum > 1 Then

    If placenum(location) < 5 Then
    
        shpPlace((5 * (location - 1)) + 1 + placenum(location)).FillColor = shpStack(stacknum - 1).FillColor 'Placing code figures position, need variable denoting placeheight currently
                                            'i.e. 4 blocks, etc.
        shpPlace((5 * (location - 1)) + 1 + placenum(location)).Visible = True
        
        shpPlace((5 * (location - 1)) + 1 + placenum(location)).Tag = shpStack(stacknum - 1).Tag
        
        placenum(location) = placenum(location) + 1
        
        stacknum = stacknum - 1
        
        shpStack(stacknum).Visible = False
        
        
            shpPlayer.Top = shpPlayer.Top - shpStack(0).Height / 1.5
            'shpStack(stacknum).Top = shpStack(stacknum).Top + shpStack(stacknum).Height / 1.5
            shpHold.Top = shpHold.Top - shpStack(0).Height / 1.5
            For y = 1 To stacknum
                    If y > 1 Then shpStack(y).Top = shpStack(y - 1).Top - shpStack(0).Height
                    If y = 1 Then shpStack(y).Top = shpStack(y).Top - shpStack(0).Height / 1.5
                    shpStack(y).Left = shpPlace((location - 1) * 5 + 1).Left
            Next y
        'End If
        'If causes end of game
        
        If stacknum < 6 Then shpHold.FillColor = vbGreen
        
        If tmrdone = True Then Klax
        
        If tmrdone = True Then Filled
        
    End If
    
End If

End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim inte As Integer

inte = KeyCode

If inte = 40 And fast = True Then
'    seconds = seconds * 4
'    intervals = seconds * 1000 / tmrA.Interval
    tmrA.Interval = tmrA.Interval * 4
    fast = False
End If

End Sub


Private Sub Generate(loc As Integer)
Randomize
'load other blocks and send them down via timer
'loc = 1
Dim COLORRND As Single
If loc = 1 Then
    If leftest <> 0 Then
        If leftest + 1 > prevl2 Then
            Load imgL2(leftest + 1)
        End If
        imgL2(leftest + 1).Top = imgL2(leftest).Top
        imgL2(leftest + 1).Left = imgL2(leftest).Left
        imgL2(leftest + 1).Height = imgL2(leftest).Height
        imgL2(leftest + 1).Width = imgL2(leftest).Width
        imgL2(leftest + 1).Visible = False
        imgL2(leftest).Visible = True
        COLORRND = Rnd * (4 + Int(level / 3))
        If Int(COLORRND) <> Round(COLORRND, 0) Then
'            COLORRND = Int(COLORRND - 1)
'        Else
            COLORRND = Int(COLORRND)
        End If
        imgL2(leftest) = imgStorage(COLORRND)
        imgL2(leftest).Tag = Int(COLORRND)
        leftest = leftest + 1
    End If
End If
If loc = 2 Then
    If lefter <> 0 Then
        If lefter + 1 > prevl1 Then
            Load imgL1(lefter + 1)
        End If
        imgL1(lefter + 1).Top = imgL1(lefter).Top
        imgL1(lefter + 1).Left = imgL1(lefter).Left
        imgL1(lefter + 1).Height = imgL1(lefter).Height
        imgL1(lefter + 1).Width = imgL1(lefter).Width
        imgL1(lefter + 1).Visible = False
        imgL1(lefter).Visible = True
        COLORRND = Rnd * (4 + Int(level / 3))
        If Int(COLORRND) <> Round(COLORRND, 0) Then
'            COLORRND = Int(COLORRND - 1)
'        Else
            COLORRND = Int(COLORRND)
        End If

        imgL1(lefter) = imgStorage(COLORRND)
        imgL1(lefter).Tag = Int(COLORRND)
        lefter = lefter + 1
    End If
End If
If loc = 3 Then
    If centerer <> 0 Then
        If centerer + 1 > prevc Then
            Load imgC(centerer + 1)
        End If
        imgC(centerer + 1).Top = imgC(centerer).Top
        imgC(centerer + 1).Left = imgC(centerer).Left
        imgC(centerer + 1).Height = imgC(centerer).Height
        imgC(centerer + 1).Width = imgC(centerer).Width
        imgC(centerer + 1).Visible = False
        imgC(centerer).Visible = True
        COLORRND = Rnd * (4 + Int(level / 3))
        If Int(COLORRND) <> Round(COLORRND, 0) Then
'            COLORRND = Int(COLORRND - 1)
'        Else
            COLORRND = Int(COLORRND)
        End If
        imgC(centerer) = imgStorage(COLORRND)
        imgC(centerer).Tag = Int(COLORRND)
        centerer = centerer + 1
    End If
End If
If loc = 4 Then
    If righter <> 0 Then
        If righter + 1 > prevr1 Then
            Load imgR1(righter + 1)
        End If
        imgR1(righter + 1).Top = imgR1(righter).Top
        imgR1(righter + 1).Left = imgR1(righter).Left
        imgR1(righter + 1).Height = imgR1(righter).Height
        imgR1(righter + 1).Width = imgR1(righter).Width
        imgR1(righter + 1).Visible = False
        imgR1(righter).Visible = True
        COLORRND = Rnd * (4 + Int(level / 3))
        If Int(COLORRND) <> Round(COLORRND, 0) Then
'            COLORRND = Int(COLORRND - 1)
'        Else
            COLORRND = Int(COLORRND)
        End If
        imgR1(righter) = imgStorage(COLORRND)
        imgR1(righter).Tag = Int(COLORRND)
        righter = righter + 1
    End If
End If
If loc = 5 Then
    If rightest <> 0 Then
        If rightest + 1 > prevr2 Then
            Load imgR2(rightest + 1)
        End If
        imgR2(rightest + 1).Top = imgR2(rightest).Top
        imgR2(rightest + 1).Left = imgR2(rightest).Left
        imgR2(rightest + 1).Height = imgR2(rightest).Height
        imgR2(rightest + 1).Width = imgR2(rightest).Width
        imgR2(rightest + 1).Visible = False
        imgR2(rightest).Visible = True
        COLORRND = Rnd * (4 + Int(level / 3))
        If Int(COLORRND) <> Round(COLORRND, 0) Then
'            COLORRND = Int(COLORRND - 1)
'        Else
            COLORRND = Int(COLORRND)
        End If
        imgR2(rightest) = imgStorage(COLORRND)
        imgR2(rightest).Tag = Int(COLORRND)
        rightest = rightest + 1
    End If
End If
 ' First Blocks have to leave behind the next one
If leftest = 0 And loc = 1 Then
    If prevl2 = 0 Then Load imgL2(leftest + 1)
    imgL2(leftest + 1).Visible = False
    imgL2(leftest).Visible = True
    COLORRND = Rnd * (4 + Int(level / 3))
    If Int(COLORRND) <> Round(COLORRND, 0) Then
'            COLORRND = Int(COLORRND - 1)
'        Else
            COLORRND = Int(COLORRND)
        End If
    imgL2(leftest) = imgStorage(COLORRND)
    imgL2(leftest).Tag = Int(COLORRND)
    leftest = leftest + 1
End If
If lefter = 0 And loc = 2 Then
    If prevl1 = 0 Then Load imgL1(lefter + 1)
    imgL1(lefter + 1).Visible = False
    imgL1(lefter).Visible = True
    COLORRND = Rnd * (4 + Int(level / 3))
    If Int(COLORRND) <> Round(COLORRND, 0) Then
'            COLORRND = Int(COLORRND - 1)
'        Else
            COLORRND = Int(COLORRND)
        End If
    imgL1(lefter) = imgStorage(COLORRND)
    imgL1(lefter).Tag = Int(COLORRND)
    lefter = lefter + 1
End If
If centerer = 0 And loc = 3 Then
    If prevc = 0 Then Load imgC(centerer + 1)
    imgC(centerer + 1).Visible = False
    imgC(centerer).Visible = True
    COLORRND = Rnd * (4 + Int(level / 5))
    If Int(COLORRND) <> Round(COLORRND, 0) Then
'            COLORRND = Int(COLORRND - 1)
'        Else
            COLORRND = Int(COLORRND)
        End If
    imgC(centerer) = imgStorage(COLORRND)
    imgC(centerer).Tag = Int(COLORRND)
    centerer = centerer + 1
End If
If righter = 0 And loc = 4 Then
    If prevr1 = 0 Then Load imgR1(righter + 1)
    imgR1(righter + 1).Visible = False
    imgR1(righter).Visible = True
    COLORRND = Rnd * (4 + Int(level / 3))
    If Int(COLORRND) <> Round(COLORRND, 0) Then
'            COLORRND = Int(COLORRND)
'        Else
            COLORRND = Int(COLORRND)
        End If
    imgR1(righter) = imgStorage(COLORRND)
    imgR1(righter).Tag = Int(COLORRND)
    righter = righter + 1
End If
If rightest = 0 And loc = 5 Then
    If prevr2 = 0 Then Load imgR2(rightest + 1)
    imgR2(rightest + 1).Visible = False
    imgR2(rightest).Visible = True
    COLORRND = Rnd * (4 + Int(level / 3))
    If Int(COLORRND) <> Round(COLORRND, 0) Then
'            COLORRND = Int(COLORRND - 1)
'        Else
            COLORRND = Int(COLORRND)
        End If
    imgR2(rightest) = imgStorage(COLORRND)
    imgR2(rightest).Tag = Int(COLORRND)
    rightest = rightest + 1
End If





End Sub

Private Sub tmrA_Timer()

'lblHighScore = Chr(7)

Dim X As Integer
Dim y As Integer
If create Mod 40 - Int(level / 4) = 0 Or create = 0 Then create = 0: Generate (Int(Rnd * 5) + 1)

'left
For X = destroyed(1) To leftest - 1
    If imgL2(X).Top + imgL2(X).Height >= lnBreak.Y1 + (200 - Int((imgPnk.Top) / intervals) / 2) And location = 1 And imgL2(X).Top + imgL2(X).Height <= lnBreak.Y1 + (275 + Int((imgPnk.Top) / intervals) / 2) + 30 And stacknum < 6 Then
        shpStack(stacknum).FillColor = blockclr(imgL2(X).Tag)
        shpStack(stacknum).Visible = True
        shpStack(stacknum).Tag = imgL2(X).Tag
        imgL2(X).Visible = False
        imgL2(X).Left = 0
        imgL2(X).Top = 0
        shpPlayer.Top = shpPlayer.Top + shpStack(0).Height / 1.5
        'shpStack(stacknum).Top = shpStack(stacknum).Top + shpStack(stacknum).Height / 1.5
        shpHold.Top = shpHold.Top + shpStack(0).Height / 1.5
            For y = 1 To stacknum
                If y > 1 Then shpStack(y).Top = shpStack(y - 1).Top - shpStack(0).Height
                If y = 1 Then shpStack(y).Top = shpStack(y).Top + shpStack(0).Height / 1.5
                shpStack(y).Left = shpPlace((location - 1) * 5 + 1).Left
            Next y
        lblScore = lblScore + 5
        'End If
        stacknum = stacknum + 1
        If stacknum = 6 Then shpHold.FillColor = vbRed
        destroyed(1) = destroyed(1) + 1
'        caught(0) = True
    End If
    If imgL2(X).Top + imgL2(X).Height <= lnBreak.Y1 + 275 Then
        imgL2(X).Left = imgL2(X).Left + Int((imgPnk.Left - 4065) / intervals) + 1
        imgL2(X).Top = imgL2(X).Top + Int((imgPnk.Top) / intervals) '38
            If imgL2(X).Top + imgL2(X).Height <= lnBreak.Y1 + 50 Then
                imgL2(X).Width = imgL2(X).Width + Int((imgPnk.Width - 195) / intervals)
                imgL2(X).Height = imgL2(X).Height + Int((imgPnk.Height - 105) / intervals) ' 105 is initial height
            End If
    End If
    If imgL2(X).Top + imgL2(X).Height > lnBreak.Y1 + 275 Then
        imgL2(X).Top = imgL2(X).Top + 100 '38
        If imgL2(X).Top + imgL2(X).Height + 100 >= 6480 Then
            imgL2(X).Visible = False
            destroyed(1) = destroyed(1) + 1
            imgL2(X).Left = 0
            imgL2(X).Top = 0
            If dropmeter < 3 Then
                shpDrop(dropmeter).FillColor = RGB(222, 222, 64)
                dropmeter = dropmeter + 1
            End If
        End If
    End If
Next X
'LEFT 1
For X = destroyed(2) To lefter - 1
    If imgL1(X).Top + imgL1(X).Height >= lnBreak.Y1 + (200 - Int((imgGry.Top) / intervals) / 2) And location = 2 And imgL1(X).Top + imgL1(X).Height <= lnBreak.Y1 + (275 + Int((imgGry.Top) / intervals) / 2) + 30 And stacknum < 6 Then
        shpStack(stacknum).FillColor = blockclr(imgL1(X).Tag)
        shpStack(stacknum).Visible = True
        shpStack(stacknum).Tag = imgL1(X).Tag
        imgL1(X).Visible = False
        imgL1(X).Left = 0
        imgL1(X).Top = 0
        shpPlayer.Top = shpPlayer.Top + shpStack(0).Height / 1.5
        'shpStack(stacknum).Top = shpStack(stacknum).Top + shpStack(stacknum).Height / 1.5
        shpHold.Top = shpHold.Top + shpStack(0).Height / 1.5
            For y = 1 To stacknum
                If y > 1 Then shpStack(y).Top = shpStack(y - 1).Top - shpStack(0).Height
                If y = 1 Then shpStack(y).Top = shpStack(y).Top + shpStack(0).Height / 1.5
                shpStack(y).Left = shpPlace((location - 1) * 5 + 1).Left
            Next y
        'End If
        lblScore = lblScore + 5
        stacknum = stacknum + 1
        If stacknum = 6 Then shpHold.FillColor = vbRed

        destroyed(2) = destroyed(2) + 1
'        caught(0) = True
    End If
    If imgL1(X).Top + imgL1(X).Height <= lnBreak.Y1 + 275 Then
        imgL1(X).Left = imgL1(X).Left + Int((imgGry.Left - 4330) / intervals) + 1
        imgL1(X).Top = imgL1(X).Top + Int((imgGry.Top) / intervals) '38
            If imgL1(X).Top + imgL1(X).Height <= lnBreak.Y1 + 50 Then
                imgL1(X).Width = imgL1(X).Width + Int((imgGry.Width - 195) / intervals)
                imgL1(X).Height = imgL1(X).Height + Int((imgGry.Height - 105) / intervals) ' 105 is initial height
            End If
    End If
    If imgL1(X).Top + imgL1(X).Height > lnBreak.Y1 + 275 Then
        imgL1(X).Top = imgL1(X).Top + 100 '38
        If imgL1(X).Top + imgL1(X).Height + 100 >= 6480 Then
            imgL1(X).Visible = False
            destroyed(2) = destroyed(2) + 1
            imgL1(X).Left = 0
            imgL1(X).Top = 0
            If dropmeter < 3 Then
                shpDrop(dropmeter).FillColor = RGB(222, 222, 64)
                dropmeter = dropmeter + 1
            End If
        End If
    End If
Next X
'center
For X = destroyed(3) To centerer - 1
    If imgC(X).Top + imgC(X).Height >= lnBreak.Y1 + (200 - Int((imgR.Top) / intervals) / 2) And location = 3 And imgC(X).Top + imgC(X).Height <= lnBreak.Y1 + (275 + Int((imgR.Top) / intervals) / 2) + 30 And stacknum < 6 Then
        shpStack(stacknum).FillColor = blockclr(imgC(X).Tag)
        shpStack(stacknum).Visible = True
        shpStack(stacknum).Tag = imgC(X).Tag
        imgC(X).Visible = False
        imgC(X).Left = 0
        imgC(X).Top = 0
        shpPlayer.Top = shpPlayer.Top + shpStack(0).Height / 1.5
        'shpStack(stacknum).Top = shpStack(stacknum).Top + shpStack(stacknum).Height / 1.5
        shpHold.Top = shpHold.Top + shpStack(0).Height / 1.5
            For y = 1 To stacknum
                If y > 1 Then shpStack(y).Top = shpStack(y - 1).Top - shpStack(0).Height
                If y = 1 Then shpStack(y).Top = shpStack(y).Top + shpStack(0).Height / 1.5
                shpStack(y).Left = shpPlace((location - 1) * 5 + 1).Left
            Next y
        'End If
        lblScore = lblScore + 5
        stacknum = stacknum + 1
        If stacknum = 6 Then shpHold.FillColor = vbRed
        destroyed(3) = destroyed(3) + 1
'        caught(0) = True
    End If
    If imgC(X).Top + imgC(X).Height <= lnBreak.Y1 + 275 Then
        imgC(X).Left = imgC(X).Left + Int((imgR.Left - 4590) / intervals) + 1
        imgC(X).Top = imgC(X).Top + Int((imgR.Top) / intervals) '38
            If imgC(X).Top + imgC(X).Height <= lnBreak.Y1 + 50 Then
                imgC(X).Width = imgC(X).Width + Int((imgR.Width - 195) / intervals)
                imgC(X).Height = imgC(X).Height + Int((imgR.Height - 105) / intervals) ' 105 is initial height
            End If
    End If
    If imgC(X).Top + imgC(X).Height > lnBreak.Y1 + 275 Then
        imgC(X).Top = imgC(X).Top + 100 '38
        If imgC(X).Top + imgC(X).Height + 100 >= 6480 Then
            imgC(X).Visible = False
            destroyed(3) = destroyed(3) + 1
            imgC(X).Left = 0
            imgC(X).Top = 0
            If dropmeter < 3 Then
                shpDrop(dropmeter).FillColor = RGB(222, 222, 64)
                dropmeter = dropmeter + 1
            End If
        End If
    End If
Next X
'R1
For X = destroyed(4) To righter - 1
    If imgR1(X).Top + imgR1(X).Height >= lnBreak.Y1 + (200 - Int((imgB.Top) / intervals) / 2) And location = 4 And imgR1(X).Top + imgR1(X).Height <= lnBreak.Y1 + (275 + Int((imgR.Top) / intervals) / 2) + 30 And stacknum < 6 Then
        shpStack(stacknum).FillColor = blockclr(imgR1(X).Tag)
        shpStack(stacknum).Visible = True
        shpStack(stacknum).Tag = imgR1(X).Tag
        imgR1(X).Visible = False
        imgR1(X).Left = 0
        imgR1(X).Top = 0
        shpPlayer.Top = shpPlayer.Top + shpStack(0).Height / 1.5
        'shpStack(stacknum).Top = shpStack(stacknum).Top + shpStack(stacknum).Height / 1.5
        shpHold.Top = shpHold.Top + shpStack(0).Height / 1.5
            For y = 1 To stacknum
                If y > 1 Then shpStack(y).Top = shpStack(y - 1).Top - shpStack(0).Height
                If y = 1 Then shpStack(y).Top = shpStack(y).Top + shpStack(0).Height / 1.5
                shpStack(y).Left = shpPlace((location - 1) * 5 + 1).Left
            Next y
        'End If
        lblScore = lblScore + 5
        stacknum = stacknum + 1
        If stacknum = 6 Then shpHold.FillColor = vbRed
        destroyed(4) = destroyed(4) + 1
'        caught(0) = True
    End If
    If imgR1(X).Top + imgR1(X).Height <= lnBreak.Y1 + 275 Then
        imgR1(X).Left = imgR1(X).Left + Int((imgB.Left - 4850) / intervals) + 1
        imgR1(X).Top = imgR1(X).Top + Int((imgB.Top) / intervals) '38
            If imgR1(X).Top + imgR1(X).Height <= lnBreak.Y1 + 50 Then
                imgR1(X).Width = imgR1(X).Width + Int((imgB.Width - 195) / intervals)
                imgR1(X).Height = imgR1(X).Height + Int((imgB.Height - 105) / intervals) ' 105 is initial height
            End If
    End If
    If imgR1(X).Top + imgR1(X).Height > lnBreak.Y1 + 275 Then
        imgR1(X).Top = imgR1(X).Top + 100 '38
        If imgR1(X).Top + imgR1(X).Height + 100 >= 6480 Then
            imgR1(X).Visible = False
            destroyed(4) = destroyed(4) + 1
            imgR1(X).Left = 0
            imgR1(X).Top = 0
            If dropmeter < 3 Then
                shpDrop(dropmeter).FillColor = RGB(222, 222, 64)
                dropmeter = dropmeter + 1
            End If
        End If
    End If
Next X
'Rightest code
For X = destroyed(5) To rightest - 1
    If imgR2(X).Top + imgR2(X).Height >= lnBreak.Y1 + (200 - Int((imgY.Top) / intervals) / 2) And location = 5 And imgR2(X).Top + imgR2(X).Height <= lnBreak.Y1 + (275 + Int((imgR.Top) / intervals) / 2) + 30 And stacknum < 6 Then
        shpStack(stacknum).FillColor = blockclr(imgR2(X).Tag)
        shpStack(stacknum).Visible = True
        shpStack(stacknum).Tag = imgR2(X).Tag
        imgR2(X).Visible = False
        imgR2(X).Left = 0
        imgR2(X).Top = 0
        shpPlayer.Top = shpPlayer.Top + shpStack(0).Height / 1.5
        'shpStack(stacknum).Top = shpStack(stacknum).Top + shpStack(stacknum).Height / 1.5
        shpHold.Top = shpHold.Top + shpStack(0).Height / 1.5
            For y = 1 To stacknum
                If y > 1 Then shpStack(y).Top = shpStack(y - 1).Top - shpStack(0).Height
                If y = 1 Then shpStack(y).Top = shpStack(y).Top + shpStack(0).Height / 1.5
                shpStack(y).Left = shpPlace((location - 1) * 5 + 1).Left
            Next y
        'End If
        lblScore = lblScore + 5
        stacknum = stacknum + 1
        If stacknum = 6 Then shpHold.FillColor = vbRed
        destroyed(5) = destroyed(5) + 1
'        caught(0) = True
    End If
    If imgR2(X).Top + imgR2(X).Height <= lnBreak.Y1 + 275 Then
        imgR2(X).Left = imgR2(X).Left + Int((imgY.Left - 5100) / intervals) '+ 1
        imgR2(X).Top = imgR2(X).Top + Int((imgY.Top) / intervals) '38
            If imgR2(X).Top + imgR2(X).Height <= lnBreak.Y1 + 50 Then
                imgR2(X).Width = imgR2(X).Width + Int((imgY.Width - 195) / intervals)
                imgR2(X).Height = imgR2(X).Height + Int((imgY.Height - 105) / intervals) ' 105 is initial height
            End If
    End If
    If imgR2(X).Top + imgR2(X).Height > lnBreak.Y1 + 275 Then
        imgR2(X).Top = imgR2(X).Top + 100 '38
        If imgR2(X).Top + imgR2(X).Height + 100 >= 6480 Then
            imgR2(X).Visible = False
            destroyed(5) = destroyed(5) + 1
            imgR2(X).Left = 0
            imgR2(X).Top = 0
            If dropmeter < 3 Then
                shpDrop(dropmeter).FillColor = RGB(222, 222, 64)
                dropmeter = dropmeter + 1
            End If
        End If
    End If
Next X

If dropmeter = 3 Then
    fraLose.Visible = True
    Picture1.Enabled = False
End If


create = create + 1

End Sub


Private Sub tmrScore_Timer() 'shows calculations for end of level

Dim X As Integer
Dim y As Integer
y = 0
Dim counter As Integer
counter = 0
Label7 = "Bonus Points"
    If beltdone = False Then
        'by timer
        
        'rightest
        counter = 0
        For X = destroyed(5) To rightest - 1
            If imgR2(X).Tag = 300 Then counter = counter + 1: y = y + 1
        Next X
        If counter = rightest - destroyed(5) Then GoTo skipR2:
        y = counter + destroyed(5)
        imgR2(y) = imgWh
        imgR2(y).Tag = 300
        lblTotalPoints = Val(lblTotalPoints) + 25
        GoTo endTimR:
skipR2:
        'righter
        counter = 0
        For X = destroyed(4) To righter - 1
            If imgR1(X).Tag = 300 Then counter = counter + 1: y = y + 1
        Next X
        If counter = righter - destroyed(4) Then GoTo skipR1:
        y = counter + destroyed(4)
        imgR1(y) = imgWh
        imgR1(y).Tag = 300
        lblTotalPoints = Val(lblTotalPoints) + 25
        GoTo endTimR:
skipR1:
        'centerer
        counter = 0
        For X = destroyed(3) To centerer - 1
            If imgC(X).Tag = 300 Then counter = counter + 1: y = y + 1
        Next X
        If counter = centerer - destroyed(3) Then GoTo skipC:
        y = counter + destroyed(3)
        imgC(y) = imgWh
        imgC(y).Tag = 300
        lblTotalPoints = Val(lblTotalPoints) + 25
        GoTo endTimR:
skipC:

        'lefter
        counter = 0
        For X = destroyed(2) To lefter - 1
            If imgL1(X).Tag = 300 Then counter = counter + 1: y = y + 1
        Next X
        If counter = lefter - destroyed(2) Then GoTo skipl1:
        y = counter + destroyed(2)
        imgL1(y) = imgWh
        imgL1(y).Tag = 300
        lblTotalPoints = Val(lblTotalPoints) + 25
        GoTo endTimR:
skipl1:

        'leftest
        counter = 0
        For X = destroyed(1) To leftest - 1
            If imgL2(X).Tag = 300 Then counter = counter + 1: y = y + 1
        Next X
        If counter = leftest - destroyed(1) Then GoTo skipl2:
        y = counter + destroyed(1)
        imgL2(y) = imgWh
        imgL2(y).Tag = 300
        lblTotalPoints = Val(lblTotalPoints) + 25
        GoTo endTimR:
skipl2:
        'holder
        counter = 0
        For X = 1 To stacknum - 1
            If shpStack(X).Tag <> 300 Then counter = counter + 1: y = y + 1
        Next X
        If counter = 0 Then GoTo SKIPHOLD:
        y = counter
        shpStack(y).FillColor = vbWhite
        shpStack(y).Tag = 300
        lblTotalPoints = Val(lblTotalPoints) + 25
        GoTo endTimR:
SKIPHOLD:

    'now make them invisible
    
        'rightest
        counter = 0
        For X = destroyed(5) To rightest - 1
            If imgR2(X).Visible = False Then counter = counter + 1: y = y + 1
        Next X
        If counter = rightest - destroyed(5) Then GoTo skipR22:
        y = counter + destroyed(5)
        imgR2(y).Visible = False
        GoTo endTimR:
skipR22:
        'righter
        counter = 0
        For X = destroyed(4) To righter - 1
            If imgR1(X).Visible = False Then counter = counter + 1: y = y + 1
        Next X
        If counter = righter - destroyed(4) Then GoTo skipR12:
        y = counter + destroyed(4)
        imgR1(y).Visible = False
        GoTo endTimR:
skipR12:
        'centerer
        counter = 0
        For X = destroyed(3) To centerer - 1
            If imgC(X).Visible = False Then counter = counter + 1: y = y + 1
        Next X
        If counter = centerer - destroyed(3) Then GoTo skipC2:
        y = counter + destroyed(3)
        imgC(y).Visible = False
        GoTo endTimR:
skipC2:

        'lefter
        counter = 0
        For X = destroyed(2) To lefter - 1
            If imgL1(X).Visible = False Then counter = counter + 1: y = y + 1
        Next X
        If counter = lefter - destroyed(2) Then GoTo skipl12:
        y = counter + destroyed(2)
        imgL1(y).Visible = False
        GoTo endTimR:
skipl12:

        'leftest
        counter = 0
        For X = destroyed(1) To leftest - 1
            If imgL2(X).Visible = False Then counter = counter + 1: y = y + 1
        Next X
        If counter = leftest - destroyed(1) Then GoTo skipl22:
        y = counter + destroyed(1)
        imgL2(y).Visible = False
        GoTo endTimR:
skipl22:
        'holder
        counter = 0
        For X = 1 To stacknum - 1
            If shpStack(X).Visible = True Then counter = counter + 1: y = y + 1
        Next X
        If counter = 0 Then GoTo SKIPHOLD2:
        y = counter
        shpStack(y).Visible = False
        GoTo endTimR:
SKIPHOLD2:

        beltdone = True
        lblKlax = Val(lblTotalPoints) + Val(lblKlax)
        lblTotalPoints = 0
        GoTo endTimR:
    End If
    If bindone = False And beltdone = True Then
        counter = 0
        y = 25
        'starting from 25 to 1
        For X = 25 To 1 Step -1
            If X > 5 Then If shpPlace(X).Visible = False Then y = X: GoTo outoLoop:
            If X <= 5 Then If shpPlace(X).Visible = True And shpPlace(X).FillColor <> vbBlack Then GoTo NONONO:
            If X <= 5 Then If shpPlace(X).Visible = False Then y = X: GoTo outoLoop:
        Next X
        GoTo NONONO:
outoLoop:
        shpPlace(y).Visible = True
        shpPlace(y).FillColor = vbBlack
        lblTotalPoints = Val(lblTotalPoints) + 200
        'shpPlace(y).BorderColor = vbRed
       GoTo endTimR:
       
       
NONONO:
        bindone = True
        lblKlax = Val(lblTotalPoints) + Val(lblKlax)
        lblTotalPoints = 0
        GoTo endTimR:
    End If
    If bindone = True And beltdone = True Then
        If Val(lblKlax) >= 50 Then
            lblKlax = Val(lblKlax) - 50
            lblScore = Val(lblScore) + 50
            tmrScore.Interval = 20
        Else
            lblScore = Val(lblScore) + Val(lblKlax)
            lblKlax = 0
            tmrScore.Interval = 300
            Reset
            tmrScore.Enabled = False
        End If
    End If
'    For x = 1 To 25
'        If shpPlace(x).FillColor = vbWhite Then shpPlace(x).FillColor = vbBlack: shpPlace(x).BorderColor = vbRed
'    Next x

endTimR:




End Sub

Private Sub tmrTest_Timer()
'each image starts at Top = 0
'R2 starts at L = 5100, must get to 6720
'Create algorithm to automate this

'tmrA.Enabled = True

Dim X, y, w, h As Integer
y = Int((imgY.Top) / intervals)
X = Int((imgY.Left - 5100) / intervals)
w = Int((imgY.Width - 195) / intervals)
h = Int((imgY.Height - 105) / intervals)
'Far Right
If imgR2(0).Top + imgR2(0).Height >= lnBreak.Y1 + 200 And location = 5 And imgR2(0).Top + imgR2(0).Height <= lnBreak.Y1 + 300 And caught(0) = False Then
    shpStack(stacknum).FillColor = blockclr(8)
    shpStack(stacknum).Visible = True
    imgR2(0).Visible = False
    shpPlayer.Top = shpPlayer.Top + shpStack(0).Height / 2
    shpStack(stacknum).Top = shpStack(stacknum).Top + shpStack(stacknum).Height / 2
    shpHold.Top = shpHold.Top + shpStack(0).Height / 2
    For X = 1 To stacknum
        If X > 1 Then shpStack(X).Top = shpStack(X - 1).Top - shpStack(0).Height
        If X = 1 Then shpStack(X).Top = shpStack(X).Top + shpStack(0).Height / 2
    Next X
    stacknum = stacknum + 1
    caught(0) = True
End If
If imgR2(0).Top + imgR2(0).Height <= lnBreak.Y1 + 300 Then
    imgR2(0).Left = imgR2(0).Left + Int((imgY.Left - 5100) / intervals)
    imgR2(0).Top = imgR2(0).Top + Int((imgY.Top) / intervals) '38
    If imgR2(0).Top + imgR2(0).Height <= lnBreak.Y1 + 50 Then
        imgR2(0).Width = imgR2(0).Width + Int((imgY.Width - 195) / intervals)
        imgR2(0).Height = imgR2(0).Height + Int((imgY.Height - 105) / intervals) ' 105 is initial height
    End If
End If
If imgR2(0).Top + imgR2(0).Height > lnBreak.Y1 + 300 Then
    imgR2(0).Top = imgR2(0).Top + 100 '38
    If imgR2(0).Top + imgR2(0).Height + 100 >= 6480 Then imgR2(0).Visible = False
End If
'Inner Right
If imgR1(0).Top + imgR1(0).Height >= lnBreak.Y1 + 200 And location = 4 And imgR1(0).Top + imgR1(0).Height <= lnBreak.Y1 + 300 And caught(0) = False Then
    shpStack(stacknum).FillColor = blockclr(0)
    shpStack(stacknum).Visible = True
    imgR1(0).Visible = False
    shpPlayer.Top = shpPlayer.Top + shpStack(0).Height / 2
    shpStack(stacknum).Top = shpStack(stacknum).Top + shpStack(stacknum).Height / 2
    shpHold.Top = shpHold.Top + shpStack(0).Height / 2
    stacknum = stacknum + 1
    caught(stacknum - 2) = True
End If
If imgR1(0).Top + imgR1(0).Height <= lnBreak.Y1 + 300 Then
    imgR1(0).Left = imgR1(0).Left + Int((imgB.Left - 4850) / intervals) + 1
    imgR1(0).Top = imgR1(0).Top + Int((imgB.Top) / intervals) '38
    If imgR1(0).Top + imgR1(0).Height <= lnBreak.Y1 + 50 Then
        imgR1(0).Width = imgR1(0).Width + Int((imgB.Width - 195) / intervals)
        imgR1(0).Height = imgR1(0).Height + Int((imgB.Height - 105) / intervals) ' 105 is initial height
    End If
End If
If imgR1(0).Top + imgR1(0).Height > lnBreak.Y1 + 300 Then
    imgR1(0).Top = imgR1(0).Top + 100 '38
    If imgR1(0).Top + imgR1(0).Height + 100 >= 6480 Then imgR1(0).Visible = False
End If

''Test
'If caught(0) = True Then
'    If imgR1(1).Top + imgR1(1).Height >= lnBreak.Y1 + 200 And location = 4 And imgR1(1).Top + imgR1(1).Height <= lnBreak.Y1 + 300 And caught(1) = False Then
'        shpStack(stacknum).FillColor = blockclr(0)
'        shpStack(stacknum).Visible = True
'        imgR1(1).Visible = False
'        shpPlayer.Top = shpPlayer.Top + shpStack(0).Height / 2
'        'shpStack(stacknum).Top = shpStack(stacknum - 1).Top - shpStack(0).Height
'        For x = 1 To stacknum
'            If x > 1 Then shpStack(x).Top = shpStack(x - 1).Top - shpStack(0).Height
'            If x = 1 Then shpStack(x).Top = shpStack(x).Top + shpStack(0).Height / 2
'        Next x
'        shpStack(stacknum).Left = shpStack(stacknum - 1).Left
'        shpHold.Top = shpHold.Top + shpStack(0).Height / 2
'        stacknum = stacknum + 1
'        caught(stacknum - 2) = True
'    End If
'    If imgR1(1).Top + imgR1(1).Height <= lnBreak.Y1 + 300 Then
'        imgR1(1).Left = imgR1(1).Left + Int((imgB.Left - 4850) / intervals) + 1
'        imgR1(1).Top = imgR1(1).Top + Int((imgB.Top) / intervals) '38
'        If imgR1(1).Top + imgR1(1).Height <= lnBreak.Y1 + 50 Then
'            imgR1(1).Width = imgR1(1).Width + Int((imgB.Width - 195) / intervals)
'            imgR1(1).Height = imgR1(1).Height + Int((imgB.Height - 105) / intervals) ' 105 is initial height
'        End If
'    End If
'    If imgR1(1).Top + imgR1(1).Height > lnBreak.Y1 + 300 Then
'        imgR1(1).Top = imgR1(1).Top + 100 '38
'        If imgR1(1).Top + imgR1(1).Height + 100 >= 6480 Then imgR1(1).Visible = False
'    End If
'End If

'Center
If imgC(0).Top + imgC(0).Height <= lnBreak.Y1 + 300 Then
    imgC(0).Left = imgC(0).Left + Int((imgR.Left - 4590) / intervals) + 1
    imgC(0).Top = imgC(0).Top + Int((imgR.Top) / intervals) '38
    If imgC(0).Top + imgC(0).Height <= lnBreak.Y1 + 50 Then
        imgC(0).Width = imgC(0).Width + Int((imgR.Width - 195) / intervals)
        imgC(0).Height = imgC(0).Height + Int((imgR.Height - 105) / intervals) ' 105 is initial height
    End If
End If
If imgC(0).Top + imgC(0).Height > lnBreak.Y1 + 300 Then
    imgC(0).Top = imgC(0).Top + 100 '38
    If imgC(0).Top + imgC(0).Height + 100 >= 6480 Then imgC(0).Visible = False
End If

'Inner Left
If imgL1(0).Top + imgL1(0).Height <= lnBreak.Y1 + 300 Then
    imgL1(0).Left = imgL1(0).Left + Int((imgGry.Left - 4330) / intervals) + 1
    imgL1(0).Top = imgL1(0).Top + Int((imgGry.Top) / intervals) '38
        If imgL1(0).Top + imgL1(0).Height <= lnBreak.Y1 + 50 Then
            imgL1(0).Width = imgL1(0).Width + Int((imgGry.Width - 195) / intervals)
            imgL1(0).Height = imgL1(0).Height + Int((imgGry.Height - 105) / intervals) ' 105 is initial height
        End If
End If
If imgL1(0).Top + imgL1(0).Height > lnBreak.Y1 + 300 Then
    imgL1(0).Top = imgL1(0).Top + 100 '38
    If imgL1(0).Top + imgL1(0).Height + 100 >= 6480 Then imgL1(0).Visible = False
End If

'Far Left
If imgL2(0).Top + imgL2(0).Height <= lnBreak.Y1 + 300 Then
    imgL2(0).Left = imgL2(0).Left + Int((imgPnk.Left - 4065) / intervals) + 1
    imgL2(0).Top = imgL2(0).Top + Int((imgPnk.Top) / intervals) '38
        If imgL2(0).Top + imgL2(0).Height <= lnBreak.Y1 + 50 Then
            imgL2(0).Width = imgL2(0).Width + Int((imgPnk.Width - 195) / intervals)
            imgL2(0).Height = imgL2(0).Height + Int((imgPnk.Height - 105) / intervals) ' 105 is initial height
        End If
End If
If imgL2(0).Top + imgL2(0).Height > lnBreak.Y1 + 300 Then
    imgL2(0).Top = imgL2(0).Top + 100 '38
    If imgL2(0).Top + imgL2(0).Height + 100 >= 6480 Then imgL2(0).Visible = False
End If

End Sub
                                                                                                                                           
                                                                                                                                           
