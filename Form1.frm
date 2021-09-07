VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Super Connect4"
   ClientHeight    =   4965
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2880
      Top             =   4080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stop Find Winning"
      Height          =   375
      Left            =   1080
      TabIndex        =   19
      Top             =   4560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Players"
      Height          =   1335
      Left            =   3720
      TabIndex        =   15
      Top             =   360
      Width           =   2415
      Begin VB.Label Label7 
         Caption         =   "Computer"
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Human"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   615
      End
      Begin VB.Shape Shape4 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   1200
         Shape           =   3  'Circle
         Top             =   720
         Width           =   495
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   240
         Shape           =   3  'Circle
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "INFORMATION"
      Height          =   2655
      Left            =   3720
      TabIndex        =   6
      Top             =   1920
      Width           =   2415
      Begin VB.Label Label10 
         Caption         =   "Depth"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   20
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Move Value"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Nodes /s"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Time"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Nodes"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   2160
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Move Now"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Setup Mode "
      Height          =   2655
      Left            =   3720
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   2415
      Begin VB.OptionButton Option2 
         Caption         =   "Red Move"
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Yellow Move"
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   960
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "End Setup"
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Clear All"
         Height          =   375
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   0
         Left            =   240
         Shape           =   3  'Circle
         Top             =   360
         Width           =   495
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   43
         Left            =   240
         Shape           =   3  'Circle
         Top             =   1200
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   615
         Left            =   240
         Top             =   240
         Width           =   495
      End
      Begin VB.Image Image3 
         Height          =   615
         Left            =   240
         Top             =   1080
         Width           =   495
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   44
         Left            =   240
         Shape           =   3  'Circle
         Top             =   2040
         Width           =   495
      End
      Begin VB.Image Image4 
         Height          =   615
         Left            =   240
         Top             =   1920
         Width           =   495
      End
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Wait Please"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   22
      Top             =   4560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "CONNECT 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   0
      TabIndex        =   18
      Top             =   240
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   3255
      Index           =   7
      Left            =   3000
      Top             =   600
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   3255
      Index           =   6
      Left            =   2520
      Top             =   600
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   3255
      Index           =   5
      Left            =   2040
      Top             =   600
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   3255
      Index           =   4
      Left            =   1560
      Top             =   600
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   3255
      Index           =   3
      Left            =   1080
      Top             =   600
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   3255
      Index           =   2
      Left            =   600
      Top             =   600
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   3255
      Index           =   1
      Left            =   120
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   42
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   1020
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   42
      Left            =   3015
      Top             =   960
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   41
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   1020
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   41
      Left            =   2535
      Top             =   960
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   40
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   1020
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   40
      Left            =   2055
      Top             =   960
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   39
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   1020
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   39
      Left            =   1575
      Top             =   960
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   38
      Left            =   1080
      Shape           =   3  'Circle
      Top             =   1020
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   38
      Left            =   1095
      Top             =   960
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   37
      Left            =   600
      Shape           =   3  'Circle
      Top             =   1020
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   37
      Left            =   615
      Top             =   960
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   36
      Left            =   120
      Shape           =   3  'Circle
      Top             =   1020
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   36
      Left            =   135
      Top             =   960
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   35
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   1500
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   35
      Left            =   3015
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   34
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   1500
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   34
      Left            =   2535
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   33
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   1500
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   33
      Left            =   2055
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   32
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   1500
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   32
      Left            =   1575
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   31
      Left            =   1080
      Shape           =   3  'Circle
      Top             =   1500
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   31
      Left            =   1095
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   30
      Left            =   600
      Shape           =   3  'Circle
      Top             =   1500
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   30
      Left            =   615
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   29
      Left            =   120
      Shape           =   3  'Circle
      Top             =   1500
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   29
      Left            =   135
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   28
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   1980
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   28
      Left            =   3015
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   27
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   1980
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   27
      Left            =   2535
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   26
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   1980
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   26
      Left            =   2055
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   25
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   1980
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   25
      Left            =   1575
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   24
      Left            =   1080
      Shape           =   3  'Circle
      Top             =   1980
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   24
      Left            =   1095
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   23
      Left            =   600
      Shape           =   3  'Circle
      Top             =   1980
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   23
      Left            =   615
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   22
      Left            =   120
      Shape           =   3  'Circle
      Top             =   1980
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   22
      Left            =   135
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   21
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   2460
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   21
      Left            =   3015
      Top             =   2400
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   20
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   2460
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   20
      Left            =   2535
      Top             =   2400
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   19
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   2460
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   19
      Left            =   2055
      Top             =   2400
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   18
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   2460
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   18
      Left            =   1575
      Top             =   2400
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   17
      Left            =   1080
      Shape           =   3  'Circle
      Top             =   2460
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   17
      Left            =   1095
      Top             =   2400
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   16
      Left            =   600
      Shape           =   3  'Circle
      Top             =   2460
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   16
      Left            =   615
      Top             =   2400
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   15
      Left            =   120
      Shape           =   3  'Circle
      Top             =   2460
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   15
      Left            =   135
      Top             =   2400
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   14
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   2940
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   14
      Left            =   3015
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   13
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   2940
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   13
      Left            =   2535
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   12
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   2940
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   12
      Left            =   2055
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   11
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   2940
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   11
      Left            =   1575
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   10
      Left            =   1080
      Shape           =   3  'Circle
      Top             =   2940
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   10
      Left            =   1095
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   9
      Left            =   600
      Shape           =   3  'Circle
      Top             =   2940
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   9
      Left            =   615
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   8
      Left            =   120
      Shape           =   3  'Circle
      Top             =   2940
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   8
      Left            =   135
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   7
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   3420
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   7
      Left            =   3015
      Top             =   3360
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   6
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   3420
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   6
      Left            =   2535
      Top             =   3360
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   5
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   3420
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   5
      Left            =   2055
      Top             =   3360
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   4
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   3420
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   4
      Left            =   1575
      Top             =   3360
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   3
      Left            =   1080
      Shape           =   3  'Circle
      Top             =   3420
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   3
      Left            =   1095
      Top             =   3360
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   2
      Left            =   600
      Shape           =   3  'Circle
      Top             =   3420
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   2
      Left            =   615
      Top             =   3360
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   1
      Left            =   120
      Shape           =   3  'Circle
      Top             =   3420
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   1
      Left            =   135
      Top             =   3360
      Width           =   495
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuNewGame 
         Caption         =   "&New Game"
      End
      Begin VB.Menu mnuSetup 
         Caption         =   "&Setup"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuPlayMethod 
      Caption         =   "&Play Method"
      Begin VB.Menu mnuBruteForce 
         Caption         =   "&Brute Force"
         Begin VB.Menu mnuBVeryEasy 
            Caption         =   "&B Very Easy"
         End
         Begin VB.Menu mnuBEasy 
            Caption         =   "B &Easy"
         End
         Begin VB.Menu mnuBGood 
            Caption         =   "B &Good"
         End
         Begin VB.Menu mnuBHard 
            Caption         =   "B &Hard"
         End
         Begin VB.Menu mnuBVeryHard 
            Caption         =   "B &Very Hard"
         End
      End
      Begin VB.Menu mnuAlphaBetaSearch 
         Caption         =   "&Alpha Beta Search"
         Begin VB.Menu mnuAVeryEasy 
            Caption         =   "&A Very Easy"
         End
         Begin VB.Menu mnuAEasy 
            Caption         =   "A &Easy"
         End
         Begin VB.Menu mnuAGood 
            Caption         =   "A &Good"
         End
         Begin VB.Menu mnuAHard 
            Caption         =   "A &Hard"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuAVeryHard 
            Caption         =   "A &Very Hard"
         End
      End
      Begin VB.Menu mnuFindForceWinning 
         Caption         =   "&Find Force Winning"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub PlayNow() '··⁄» «·‰ﬁ·… ··ﬂÊ„»ÌÊ —
Dim a As Integer
Dim g As Boolean '‘—ÿ  ﬂ„·… «·»ÕÀ ›Ì «·»ÕÀ ⁄‰ ‰ﬁ·… —«»Õ…

mnuGame.Enabled = False
mnuPlayMethod.Enabled = False


If PlayMethod = 1 Then
    Nodes = 0
    MousePointer = 11
    QStart = Timer
    Best = Search(Depth)
    If WhiteToMove Then
        Call MakeMove(PLine(Ply), Cwhite): Shape2(PLine(Ply - 1)).FillColor = &H80FFFF
    Else
        Call MakeMove(PLine(Ply), Cblack): Shape2(PLine(Ply - 1)).FillColor = &HFF&
    End If
    QFinish = Timer
    Label1 = Nodes
    Label2 = QFinish - QStart
    If Label2 > 0 Then Label3 = Int(Label1 / Label2) Else Label3 = Label1
    Label6 = Best
    Label9 = Depth
    MousePointer = 0
End If
   
If PlayMethod = 2 Then
    QNodes = 0
    MousePointer = 11
    QStart = Timer
    Best = QSearch(Depth, -11000, 11000)
    If WhiteToMove Then
     Call MakeMove(PLine(Ply), Cwhite): Shape2(PLine(Ply - 1)).FillColor = &H80FFFF
    Else
     Call MakeMove(PLine(Ply), Cblack): Shape2(PLine(Ply - 1)).FillColor = &HFF&
    End If
    QFinish = Timer
    Label1 = QNodes
    Label2 = QFinish - QStart
    If Label2 > 0 Then Label3 = Int(Label1 / Label2) Else Label3 = Label1
    Label6 = Best
    Label9 = Depth
    MousePointer = 0
End If

'··»ÕÀ ⁄‰ «·—»Õ «·«Ã»«—Ì Ê·Ì” »«·⁄„ﬁ
If PlayMethod = 3 Then
    Label1 = "": Label2 = "": Label3 = "": Label6 = "": Label9 = ""
    Timer1.Enabled = True
    Command1.Visible = True
    Command7.Enabled = False
    Label11.Visible = True
    Best = 0
    StopFindWin = False
    g = True
    QNodes = 0
    MousePointer = 11
    Depth = 0
    Do While g = True
     If Depth > (42 - Ply) Then Exit Do
     Depth = Depth + 1
     Best = QSearch(Depth, -11000, 11000)
     If (Best < 9000 And Best > -9000 And Not StopFindWin) Then g = True Else g = False
    Loop

    If Not StopFindWin Then '«–« ·„ ÌÊﬁ› «·„” Œœ„ »ÕÀÂ
    '›ﬁœ ÊÃœ «·‰ﬁ·…
    Timer1.Enabled = False
    Label11.Visible = False
    Command7.Enabled = True
   'Õ—ﬂ «·‰ﬁ·…
    If WhiteToMove Then
     Call MakeMove(PLine(Ply), Cwhite)
     Shape2(PLine(Ply - 1)).FillColor = &H80FFFF
     Else
     Call MakeMove(PLine(Ply), Cblack)
     Shape2(PLine(Ply - 1)).FillColor = &HFF&
    End If
     MousePointer = 0
  Else '√Êﬁ› «·„” Œœ„ «·»ÕÀ
    MousePointer = 0
    Timer1.Enabled = False
    Label11.Visible = False
    Command7.Enabled = True
  End If
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Â· «‰ Â  «··⁄»…
If EvalWin > 9000 Then
    MsgBox "Yellow Wins"
    mnuGame.Enabled = True: mnuPlayMethod.Enabled = False: Command7.Enabled = False: Exit Sub
End If
If EvalWin < -9000 Then
    MsgBox "Red Wins"
    mnuGame.Enabled = True: mnuPlayMethod.Enabled = False: Command7.Enabled = False: Exit Sub
End If

If Ply <> 43 Then
    '«·”„«Õ ··»‘—Ì »«··⁄»
    For a = 1 To 7: Image1(a).Enabled = True: Next
    Command7.Enabled = True
    mnuGame.Enabled = True
    mnuPlayMethod.Enabled = True
Else
    mnuGame.Enabled = True
    mnuPlayMethod.Enabled = False
    MsgBox "The game over by Draw"
End If
End Sub








Private Sub Command1_Click()
Command1.Visible = False
StopFindWin = True

End Sub


Private Sub Command3_Click()
'setup mode
If EvalWin <> 0 Then MsgBox "This position is ending by connect 4": Exit Sub

If Option1 = True Then WhiteToMove = True Else WhiteToMove = False
SetupMode = False
mnuPlayMethod.Enabled = True
Frame1.Visible = False
Frame2.Visible = True
Command7.Enabled = True
'‰»œ· «·«·Ê«‰
If WhiteToMove Then
   Shape3.FillColor = &HFF&
   Shape4.FillColor = &HFFFF&
Else
    Shape3.FillColor = &HFFFF&
    Shape4.FillColor = &HFF&
End If

End Sub






Private Sub Command5_Click()
Dim i As Integer
Erase Moves()
Ply = 1
For i = 1 To 42
    Moves(i).Target = i
    Moves(i).File = (i Mod 7): If Moves(i).File = 0 Then Moves(i).File = 7
    Moves(i).Rank = (i + 6) \ 7
    Moves(i).PPlayed = False
Next
For i = 1 To 42: Shape2(i).FillColor = &HFFFFFF: Next

End Sub



Private Sub Command7_Click()
Command7.Enabled = False
'‰»œ· «·«·Ê«‰
If WhiteToMove Then
   Shape3.FillColor = &HFF&
   Shape4.FillColor = &HFFFF&
Else
    Shape3.FillColor = &HFFFF&
    Shape4.FillColor = &HFF&
End If
'‰Ã⁄· «·ﬂÊ„»ÌÊ — Ì·⁄»
PlayNow

End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

Dim i As Integer
Erase Moves()
Ply = 1
For i = 1 To 42
    Moves(i).Target = i
    Moves(i).File = (i Mod 7): If Moves(i).File = 0 Then Moves(i).File = 7
    Moves(i).Rank = (i + 6) \ 7
    Moves(i).PPlayed = False
Next

WhiteToMove = True
PlayMethod = 2
Depth = 8
End Sub

Private Sub Form_Terminate()
End
End Sub

Private Sub Image1_Click(Index As Integer)
Dim a As Byte, b As Byte
SaveEvalWin = 0
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If SetupMode Then
    GenerateMoves
    For a = Image1(Index).Index To 42 Step 7
         For b = 1 To 7
          If LegalMoves(b) = a Then GoTo 300
         Next
    Next
If b = 8 Then Exit Sub
300 If LastColo = CEmpty Then
    If LegalMoves(b) < 7 Then Exit Sub '·« —Ã⁄ «·‰ﬁ·«  «·›«—€… ›Ì «·’› «·«Ê·
    Call UnMakeMove(LegalMoves(b) - 7)
    Shape2(LegalMoves(b) - 7).FillColor = &HFFFFFF
    Else
    Call MakeMove(LegalMoves(b), LastColo)
    If LastColo = Cblack Then Shape2(LegalMoves(b)).FillColor = &HFF& Else Shape2(LegalMoves(b)).FillColor = &HFFFF&
    End If
  Exit Sub
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
GenerateMoves
For a = Image1(Index).Index To 42 Step 7
    For b = 1 To 7
    If LegalMoves(b) = a Then GoTo 500
    Next
Next
If b = 8 Then Exit Sub
500 If Not WhiteToMove Then
        Shape3.FillColor = &HFF&
        Shape4.FillColor = &HFFFF&
        Call MakeMove(LegalMoves(b), Cblack): Shape2(LegalMoves(b)).FillColor = &HFF&
    Else
        Call MakeMove(LegalMoves(b), Cwhite): Shape2(LegalMoves(b)).FillColor = &HFFFF&
         Shape3.FillColor = &HFFFF&
         Shape4.FillColor = &HFF&
    End If
    
    If EvalWin > 9000 Then
        MsgBox "Yellow Wins"
        mnuGame.Enabled = True: mnuPlayMethod.Enabled = False: Command7.Enabled = False: Exit Sub
    End If
    If EvalWin < -9000 Then
        MsgBox "Red Wins"
        mnuGame.Enabled = True: mnuPlayMethod.Enabled = False: Command7.Enabled = False: Exit Sub
    End If
For a = 1 To 7: Image1(a).Enabled = False: Next
Command7.Enabled = False
If Ply <> 43 Then '·Ì·⁄» «·ﬂÊ„»ÌÊ —
    PlayNow
Else
    mnuPlayMethod.Enabled = False
    MsgBox "The game over by Draw"
End If
End Sub

Private Sub Image2_Click()

LastColo = Cwhite
End Sub

Private Sub Image3_Click()
LastColo = Cblack
End Sub

Private Sub Image4_Click()
LastColo = CEmpty
End Sub

Private Sub mnuAbout_Click()
Load frmAbout
frmAbout.Show
End Sub

Private Sub mnuAEasy_Click()
mnuAVeryEasy.Checked = False
mnuAEasy.Checked = True
mnuAGood.Checked = False
mnuAHard.Checked = False
mnuAVeryHard.Checked = False

mnuBVeryEasy.Checked = False
mnuBEasy.Checked = False
mnuBGood.Checked = False
mnuBHard.Checked = False
mnuBVeryHard.Checked = False
mnuFindForceWinning.Checked = False

PlayMethod = 2
Depth = 4

End Sub

Private Sub mnuAGood_Click()
mnuAVeryEasy.Checked = False
mnuAEasy.Checked = False
mnuAGood.Checked = True
mnuAHard.Checked = False
mnuAVeryHard.Checked = False

mnuBVeryEasy.Checked = False
mnuBEasy.Checked = False
mnuBGood.Checked = False
mnuBHard.Checked = False
mnuBVeryHard.Checked = False
mnuFindForceWinning.Checked = False

PlayMethod = 2
Depth = 6

End Sub

Private Sub mnuAHard_Click()
mnuAVeryEasy.Checked = False
mnuAEasy.Checked = False
mnuAGood.Checked = False
mnuAHard.Checked = True
mnuAVeryHard.Checked = False

mnuBVeryEasy.Checked = False
mnuBEasy.Checked = False
mnuBGood.Checked = False
mnuBHard.Checked = False
mnuBVeryHard.Checked = False
mnuFindForceWinning.Checked = False

PlayMethod = 2
Depth = 8

End Sub

Private Sub mnuAVeryEasy_Click()
mnuAVeryEasy.Checked = True
mnuAEasy.Checked = False
mnuAGood.Checked = False
mnuAHard.Checked = False
mnuAVeryHard.Checked = False

mnuBVeryEasy.Checked = False
mnuBEasy.Checked = False
mnuBGood.Checked = False
mnuBHard.Checked = False
mnuBVeryHard.Checked = False
mnuFindForceWinning.Checked = False

PlayMethod = 2
Depth = 2

End Sub

Private Sub mnuAVeryHard_Click()
mnuAVeryEasy.Checked = False
mnuAEasy.Checked = False
mnuAGood.Checked = False
mnuAHard.Checked = False
mnuAVeryHard.Checked = True

mnuBVeryEasy.Checked = False
mnuBEasy.Checked = False
mnuBGood.Checked = False
mnuBHard.Checked = False
mnuBVeryHard.Checked = False
mnuFindForceWinning.Checked = False

PlayMethod = 2
Depth = 10

End Sub

Private Sub mnuBEasy_Click()
mnuAVeryEasy.Checked = False
mnuAEasy.Checked = False
mnuAGood.Checked = False
mnuAHard.Checked = False
mnuAVeryHard.Checked = False

mnuBVeryEasy.Checked = False
mnuBEasy.Checked = True
mnuBGood.Checked = False
mnuBHard.Checked = False
mnuBVeryHard.Checked = False
mnuFindForceWinning.Checked = False

PlayMethod = 1
Depth = 3

End Sub

Private Sub mnuBGood_Click()
mnuAVeryEasy.Checked = False
mnuAEasy.Checked = False
mnuAGood.Checked = False
mnuAHard.Checked = False
mnuAVeryHard.Checked = False

mnuBVeryEasy.Checked = False
mnuBEasy.Checked = False
mnuBGood.Checked = True
mnuBHard.Checked = False
mnuBVeryHard.Checked = False
mnuFindForceWinning.Checked = False

PlayMethod = 1
Depth = 4

End Sub

Private Sub mnuBHard_Click()
mnuAVeryEasy.Checked = False
mnuAEasy.Checked = False
mnuAGood.Checked = False
mnuAHard.Checked = False
mnuAVeryHard.Checked = False

mnuBVeryEasy.Checked = False
mnuBEasy.Checked = False
mnuBGood.Checked = False
mnuBHard.Checked = True
mnuBVeryHard.Checked = False
mnuFindForceWinning.Checked = False

PlayMethod = 1
Depth = 5

End Sub

Private Sub mnuBVeryEasy_Click()

mnuAVeryEasy.Checked = False
mnuAEasy.Checked = False
mnuAGood.Checked = False
mnuAHard.Checked = False
mnuAVeryHard.Checked = False

mnuBVeryEasy.Checked = True
mnuBEasy.Checked = False
mnuBGood.Checked = False
mnuBHard.Checked = False
mnuBVeryHard.Checked = False
mnuFindForceWinning.Checked = False

PlayMethod = 1
Depth = 2

End Sub

Private Sub mnuBVeryHard_Click()
mnuAVeryEasy.Checked = False
mnuAEasy.Checked = False
mnuAGood.Checked = False
mnuAHard.Checked = False
mnuAVeryHard.Checked = False

mnuBVeryEasy.Checked = False
mnuBEasy.Checked = False
mnuBGood.Checked = False
mnuBHard.Checked = False
mnuBVeryHard.Checked = True
mnuFindForceWinning.Checked = False

PlayMethod = 1
Depth = 6

End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuFindForceWinning_Click()
Command1.Visible = True
Command7.Enabled = False
Label11.Visible = True


mnuAVeryEasy.Checked = False
mnuAEasy.Checked = False
mnuAGood.Checked = False
mnuAHard.Checked = False
mnuAVeryHard.Checked = False

mnuBVeryEasy.Checked = False
mnuBEasy.Checked = False
mnuBGood.Checked = False
mnuBHard.Checked = False
mnuBVeryHard.Checked = False
mnuFindForceWinning.Checked = True

PlayMethod = 3
PlayNow
End Sub

Private Sub mnuNewGame_Click()
Dim i As Integer
Erase Moves()
Ply = 1
For i = 1 To 42
    Moves(i).Target = i
    Moves(i).File = (i Mod 7): If Moves(i).File = 0 Then Moves(i).File = 7
    Moves(i).Rank = (i + 6) \ 7
    Moves(i).PPlayed = False
Next
For i = 1 To 42: Shape2(i).FillColor = &HFFFFFF: Next


WhiteToMove = True
Shape3.FillColor = &HFFFF&
Shape4.FillColor = &HFF&
Command7.Enabled = True
mnuPlayMethod.Enabled = True
For i = 1 To 7: Image1(i).Enabled = True: Next
End Sub

Private Sub mnuSetup_Click()
SetupMode = True
mnuPlayMethod.Enabled = False
Frame2.Visible = False
Frame1.Visible = True
For i = 1 To 7: Image1(i).Enabled = True: Next
Command7.Enabled = False
End Sub



Private Sub Timer1_Timer()
If PlayMethod = 3 Then
    Label1 = QNodes
    Label2 = Val(Label2) + 2
    If Label2 > 0 Then Label3 = Int(Label1 / Label2) Else Label3 = Label1
    Label9 = Depth
End If
End Sub
