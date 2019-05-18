VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LINK DOTS V0.1A"
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

Dim KeyDown(0 To 1000) As Boolean

Public Function IsKeyDownMAIN(ByVal Key As Integer) As Boolean
    IsKeyDownMAIN = KeyDown(Key)
End Function

Private Sub Form_Load()

    Randomize

    frmDebug.Show
    frmImages.Show

    picDisplay.Top = 0
    picDisplay.Left = 0
    picDisplay.Width = frmMain.Width
    picDisplay.Height = frmMain.Height
    
    Init
    
   ' Dim Level() As Integer
   ' Level = GenerateLevel(100, 100, 0)
  '  LoadLevel Level, 100, 100
    Dim TestObj As GameObject
    Dim TestEnemy As GameObject
    Dim TestEnemy2 As GameObject
    Dim TestRoom As Room
   ' Dim WallObj As GameObject
    Dim A As Long
    Dim B As Long
    
    For A = 0 To 9
        For B = 0 To 9
         '   Dim WallObj As GameObject
         '   Create WallObj, 1000 + A * 480, 1000 + B * 480, CHECKER
        Next B
    Next A
    TestRoom.CorrectPath = "U"
    
    Create TestObj, ToTwips(500), ToTwips(500), Player
    Create TestEnemy, ToTwips(600), ToTwips(500), ENEMY1
    Create TestEnemy2, ToTwips(800), ToTwips(500), ENEMY1
    
   ' LoadRoom TestRoom
    
  '  Create WallObj, 5000, 1000, CHECKER
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
