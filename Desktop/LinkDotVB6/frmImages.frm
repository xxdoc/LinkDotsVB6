VERSION 5.00
Begin VB.Form frmImages 
   Caption         =   "Form1"
   ClientHeight    =   7590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   11550
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgEnemyKnight 
      Height          =   990
      Index           =   0
      Left            =   3480
      Picture         =   "frmImages.frx":0000
      Top             =   1080
      Width           =   960
   End
   Begin VB.Image imgCheckerBackground 
      Height          =   5760
      Index           =   0
      Left            =   3480
      Picture         =   "frmImages.frx":14C2
      Top             =   3960
      Width           =   10560
   End
   Begin VB.Image imgBkgForest 
      Height          =   15000
      Index           =   0
      Left            =   360
      Picture         =   "frmImages.frx":43904
      Top             =   3960
      Width           =   15000
   End
   Begin VB.Image imgCheckerFloor 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "frmImages.frx":320006
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image imgBrick 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "frmImages.frx":320848
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image imgWizard 
      Height          =   1215
      Index           =   1
      Left            =   1200
      Picture         =   "frmImages.frx":37CC5A
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1020
   End
   Begin VB.Image imgWizard 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "frmImages.frx":380B43
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
