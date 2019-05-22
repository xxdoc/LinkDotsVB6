VERSION 5.00
Begin VB.Form frmImages 
   Caption         =   "Form1"
   ClientHeight    =   7590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14580
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   14580
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgTutorial 
      Height          =   5925
      Index           =   3
      Left            =   8760
      Picture         =   "frmImages.frx":0000
      Top             =   4440
      Width           =   9510
   End
   Begin VB.Image imgTutorial 
      Height          =   5925
      Index           =   2
      Left            =   8880
      Picture         =   "frmImages.frx":3D996
      Top             =   3480
      Width           =   9510
   End
   Begin VB.Image imgTutorial 
      Height          =   5925
      Index           =   1
      Left            =   9120
      Picture         =   "frmImages.frx":7B32C
      Top             =   2280
      Width           =   9510
   End
   Begin VB.Image imgTutorial 
      Height          =   5925
      Index           =   0
      Left            =   8760
      Picture         =   "frmImages.frx":B8CC2
      Top             =   1560
      Width           =   9510
   End
   Begin VB.Image imgWilloWisp 
      Height          =   525
      Index           =   4
      Left            =   8520
      Picture         =   "frmImages.frx":F6658
      Top             =   2160
      Width           =   645
   End
   Begin VB.Image imgWilloWisp 
      Height          =   660
      Index           =   3
      Left            =   7680
      Picture         =   "frmImages.frx":F709E
      Top             =   2040
      Width           =   525
   End
   Begin VB.Image imgWilloWisp 
      Height          =   660
      Index           =   2
      Left            =   6960
      Picture         =   "frmImages.frx":F7B10
      Top             =   2040
      Width           =   525
   End
   Begin VB.Image imgWilloWisp 
      Height          =   660
      Index           =   1
      Left            =   6120
      Picture         =   "frmImages.frx":F8582
      Top             =   2040
      Width           =   525
   End
   Begin VB.Image imgWilloWisp 
      Height          =   660
      Index           =   0
      Left            =   5280
      Picture         =   "frmImages.frx":F8FF4
      Top             =   2040
      Width           =   525
   End
   Begin VB.Image imgEndlessBackground 
      Height          =   12000
      Index           =   1
      Left            =   4200
      Picture         =   "frmImages.frx":F9A66
      Top             =   3360
      Width           =   12000
   End
   Begin VB.Image imgEndlessBackground 
      Height          =   12000
      Index           =   0
      Left            =   120
      Picture         =   "frmImages.frx":1962A8
      Top             =   3360
      Width           =   12000
   End
   Begin VB.Shape shpLightGreen 
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   6120
      Top             =   720
      Width           =   975
   End
   Begin VB.Image imgEnemyKnight 
      Height          =   990
      Index           =   0
      Left            =   3480
      Picture         =   "frmImages.frx":232AEA
      Top             =   1080
      Width           =   960
   End
   Begin VB.Image imgCheckerBackground 
      Height          =   5760
      Index           =   0
      Left            =   3480
      Picture         =   "frmImages.frx":233FAC
      Top             =   3960
      Width           =   10560
   End
   Begin VB.Image imgCheckerFloor 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "frmImages.frx":2763EE
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image imgBrick 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "frmImages.frx":276C30
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image imgWizard 
      Height          =   1215
      Index           =   1
      Left            =   1200
      Picture         =   "frmImages.frx":2D3042
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1020
   End
   Begin VB.Image imgWizard 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "frmImages.frx":2D6F2B
      Stretch         =   -1  'True
      Top             =   720
      Width           =   480
   End
   Begin VB.Label lblUserWarning 
      Caption         =   "How are you seeing this?"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Author: William Chan
'Date: May 17th, 2019
'Purpose: ICS4U Culminating Assignment
