VERSION 5.00
Begin VB.Form frmTutorial 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Link Dots V1 - How to play"
   ClientHeight    =   6555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>"
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   5880
      Width           =   4695
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<<"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   5880
      Width           =   4695
   End
   Begin VB.Image imgTutorialDisplay 
      Height          =   5295
      Left            =   0
      Top             =   0
      Width           =   9255
   End
End
Attribute VB_Name = "frmTutorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Author: William Chan
'Date: May 17th, 2019
'Purpose: ICS4U Culminating Assignment

Dim Frame As Integer

Private Sub LoadFrame()
    imgTutorialDisplay.Picture = frmImages.imgTutorial(Frame)
End Sub

Private Sub cmdBack_Click()
    If (Frame > 0) Then
        Frame = Frame - 1
    End If
    LoadFrame
End Sub

Private Sub cmdNext_Click()
    If (Frame < 3) Then
        Frame = Frame + 1
    End If
    LoadFrame
End Sub

Private Sub Form_Load()
    Frame = 0
    LoadFrame
End Sub
