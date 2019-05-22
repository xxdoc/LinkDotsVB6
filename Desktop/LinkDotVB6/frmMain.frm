VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LINK DOTS V1"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDisplay 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   0
      ScaleHeight     =   5175
      ScaleWidth      =   7215
      TabIndex        =   0
      Top             =   0
      Width           =   7215
   End
   Begin VB.Timer tmrLoop 
      Interval        =   1
      Left            =   9000
      Top             =   960
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Author: William Chan
'Date: May 17th, 2019
'Purpose: ICS4U Culminating Assignment

Dim KeyDown() As Boolean

Public Function IsKeyDownMAIN(ByVal Key As Integer) As Boolean
    IsKeyDownMAIN = KeyDown(Key)
End Function

Public Sub ResetKeys()
    ReDim KeyDown(0 To 1000)
End Sub

Private Sub Form_Load()
    Running = True
    
    ReDim KeyDown(0 To 1000)
    
    picDisplay.Top = 0
    picDisplay.Left = 0
    picDisplay.Width = frmMain.Width
    picDisplay.Height = frmMain.Height
    
    Init
    LoadNewLevel
    
    LevelsCompleted = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload frmImages
    Unload frmDebug
    Unload frmMain
End Sub

Private Sub picDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyDown(KeyCode) = True
End Sub

Private Sub picDisplay_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyDown(KeyCode) = False
End Sub

Private Sub tmrLoop_Timer()
    If (Running) Then
        picDisplay.Cls
        Update
        Render
    Else
        Form_QueryUnload 0, 0
    End If
End Sub
