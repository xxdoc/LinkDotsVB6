Attribute VB_Name = "mdlLevel"
Option Explicit

Const TILELENGTH = 480

Global Const LEVELWIDTH = 10080
Global Const LEVELHEIGHT = 5280

Global Const MAXOBJECTS = 200

Public Type Room
    Objects(0 To MAXOBJECTS) As GameObject
    N As Integer 'Number of objects
    CorrectPath As String * 1
End Type

Public Function GenerateLevel(ByVal LWidth As Integer, ByVal LHeight As Integer, ByVal Difficulty As Integer) As Integer()
    Dim Level(0 To 100, 0 To 100) As Integer
    Dim I As Integer
    Dim CX As Integer
    Dim CY As Integer
    Dim PlayerSpawned As Boolean
    Dim Direction
    Direction = 0
    PlayerSpawned = False
    CX = (LWidth) / 32
    CY = (LHeight) / 32
    For I = 0 To 999
        If (Rnd() > 0.5 Or CX > LWidth - 1 Or CX < 2 Or CY > LHeight - 1 Or CY < 2) Then
            If (Rnd() > 0.5) Then
                Direction = (Direction + 1) Mod 4
            Else
                Direction = (Direction + 3) Mod 4
            End If
        End If
        
        If (Direction = 1 Or Direction = 3) Then
            CX = CX + Direction - 2
        ElseIf (Direction = 2 Or Direction = 0) Then
            CY = CY + Direction - 1
        End If
        
        'Place tile
        'If (Level(CX, CY) <> "W") Then
        If (CX >= 0 And CX < LWidth And CY >= 0 And CY < LHeight) Then
            Level(CX, CY) = 1
            If (Not PlayerSpawned) Then
                Level(CX, CY) = 3
               ' MsgBox "AY"
                PlayerSpawned = True
            End If
        
       ' End If
        'Place Walls
        If (CX < LWidth) Then
            If (Level(CX + 1, CY) <> 1 And Not (CX + 1 > LWidth - 1) And Direction <> 1) Then
                Level(CX + 1, CY) = 2
            End If
        End If
        If (CX > 1) Then
            If (Level(CX - 1, CY) <> 1 And Not (CX - 1 < 1) And Direction <> 3) Then
                Level(CX - 1, CY) = 2
            End If
        End If
        If (CY < LHeight) Then
            If (Level(CX, CY + 1) <> 1 And Not (CY + 1 > LHeight - 1) And Direction <> 2) Then
                Level(CX, CY + 1) = 2
            End If
        End If
        If (CY > 1) Then
            If (Level(CX, CY - 1) <> 1 And Not (CY - 1 < 1) And Direction <> 0) Then
                Level(CX, CY - 1) = 2
            End If
        End If
        End If
    Next I
    
    GenerateLevel = Level
End Function

Public Sub RenderLevel(Level() As Integer, ByVal LWidth As Integer, ByVal LHeight As Integer)
    Dim I As Long
    Dim J As Long
    For I = 0 To LWidth - 1
        For J = 0 To LHeight - 1
            If (CamCorrectX(I * TILELENGTH) < INTEGERLIMIT And CamCorrectX(I * TILELENGTH) > -INTEGERLIMIT) _
                        And (CamCorrectY(I * TILELENGTH) < INTEGERLIMIT And CamCorrectY(I * TILELENGTH) > -INTEGERLIMIT) Then
                Select Case Level(I, J)
                    Case 1 'Floor
                    '    frmMain.picDisplay.PaintPicture frmImages.imgCheckerFloor(0), CamCorrectX(I * TILELENGTH), CamCorrectY(J * TILELENGTH), 480, 480
                    Case 2 'Wall
                    '    frmMain.picDisplay.PaintPicture frmImages.imgBrick(0), CamCorrectX(I * TILELENGTH), CamCorrectY(J * TILELENGTH), 480, 480
                End Select
            End If
        Next J
    Next I
End Sub

Public Sub LoadLevel(Level() As Integer, ByVal LWidth As Integer, ByVal LHeight As Integer)
    ReDim Objects(0)
    Dim I As Long
    Dim J As Long
    
    For I = 0 To LWidth - 1
        For J = 0 To LHeight - 1
            Dim TileObj As GameObject
            Select Case Level(I, J)
                Case 3
                    CreateQueue TileObj, I * TILELENGTH, J * TILELENGTH, Player
            End Select
        Next J
    Next I
    CurrentLevel = Level
End Sub

Public Sub LoadRoomQueue(ToLoad As Room)
    LoadingNext = True
    NextRoom = ToLoad
End Sub

Public Sub LoadRoom(ToLoad As Room)
    ReDim Objects(0)
    
    Dim I As Integer
    
    Dim DoorUp As GameObject
    Dim DoorRight As GameObject
    Dim DoorLeft As GameObject
    Dim DoorDown As GameObject
    Dim PlayerObject As GameObject
    
    For I = 0 To ToLoad.N
        AddObject ToLoad.Objects(I)
    Next I
    
    DoorUp.DataTag = "U"
    DoorRight.DataTag = "R"
    DoorLeft.DataTag = "L"
    DoorDown.DataTag = "D"
    
    Create DoorUp, ToTwips(425), 0, DOOR
    Create DoorRight, 0, ToTwips(430), DOOR
    Create DoorLeft, ToTwips(425), ToTwips(968), DOOR
    Create DoorDown, ToTwips(968), ToTwips(430), DOOR
    Create PlayerObject, ToTwips(500), ToTwips(500), Player
    
    CurrentRoom = ToLoad
End Sub
