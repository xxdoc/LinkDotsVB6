VERSION 5.00
Begin VB.Form frmImages 
   Caption         =   "Form1"
   ClientHeight    =   4485
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgBrick 
      Height          =   720
      Index           =   0
      Left            =   120
      Picture         =   "frmImages.frx":0000
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image imgWizard 
      Height          =   1215
      Index           =   1
      Left            =   1200
      Picture         =   "frmImages.frx":5C412
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1020
   End
   Begin VB.Image imgWizard 
      Height          =   735
      Index           =   0
      Left            =   120
      Picture         =   "frmImages.frx":602FB
      Stretch         =   -1  'True
      Top             =   720
      Width           =   540
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
