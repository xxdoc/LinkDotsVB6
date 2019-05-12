VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
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

Dim KeyDown(0 To 1000) As Boolean

Public Function IsKeyDownMAIN(ByVal Key As Integer) As Boolean
    IsKeyDownMAIN = KeyDown(Key)
End Function

Private Sub Form_Load()
    frmDebug.Show
    frmImages.Show

    picDisplay.Top = 0
    picDisplay.Left = 0
    picDisplay.Width = frmMain.Width
    picDisplay.Height = frmMain.Height
    
    Init
    
    Dim TestObj As GameObject
    
    Create TestObj, 3000, 1000, PLAYER
    
    Dim WallObj As GameObject
    
    Create WallObj, 5000, 1000, WALL
    
    'TypeIDs: 0 = Player

End Sub

Private Sub picDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyDown(KeyCode) = True
End Sub

Private Sub picDisplay_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyDown(KeyCode) = False
End Sub

Private Sub tmrLoop_Timer()
    picDisplay.Cls
    Update
    Render
End Sub
