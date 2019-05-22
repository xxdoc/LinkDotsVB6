Attribute VB_Name = "mdlLevel"
Option Explicit

'Author: William Chan
'Date: May 17th, 2019
'Purpose: ICS4U Culminating Assignment

Global Const TILELENGTH = 480
Global Const MAXOBJECTS = 200

Global MapCollision() As EuclideanLine

Public Type SeamlessRoom
    Path As String * 5
    Spot As Integer
End Type


Public Sub SetMapCollisions()
    If (BackgroundID = 0) Then
        MapCollision(0) = CreateLine(ToTwips(200 + MacroX * 800), ToTwips(178 + MacroY * 800), ToTwips(331 + MacroX * 800), ToTwips(178 + MacroY * 800))
        MapCollision(1) = CreateLine(ToTwips(331 + MacroX * 800), ToTwips(178 + MacroY * 800), ToTwips(331 + MacroX * 800), ToTwips(0 + MacroY * 800))
        MapCollision(2) = CreateLine(ToTwips(460 + MacroX * 800), ToTwips(0 + MacroY * 800), ToTwips(460 + MacroX * 800), ToTwips(178 + MacroY * 800))
        MapCollision(3) = CreateLine(ToTwips(460 + MacroX * 800), ToTwips(178 + MacroY * 800), ToTwips(600 + MacroX * 800), ToTwips(178 + MacroY * 800))
        MapCollision(4) = CreateLine(ToTwips(600 + MacroX * 800), ToTwips(178 + MacroY * 800), ToTwips(600 + MacroX * 800), ToTwips(331 + MacroY * 800))
        MapCollision(5) = CreateLine(ToTwips(600 + MacroX * 800), ToTwips(331 + MacroY * 800), ToTwips(800 + MacroX * 800), ToTwips(331 + MacroY * 800))
        MapCollision(6) = CreateLine(ToTwips(800 + MacroX * 800), ToTwips(450 + MacroY * 800), ToTwips(600 + MacroX * 800), ToTwips(450 + MacroY * 800))
        MapCollision(7) = CreateLine(ToTwips(600 + MacroX * 800), ToTwips(550 + MacroY * 800), ToTwips(460 + MacroX * 800), ToTwips(550 + MacroY * 800))
        MapCollision(8) = CreateLine(ToTwips(460 + MacroX * 800), ToTwips(550 + MacroY * 800), ToTwips(460 + MacroX * 800), ToTwips(800 + MacroY * 800))
        MapCollision(9) = CreateLine(ToTwips(331 + MacroX * 800), ToTwips(800 + MacroY * 800), ToTwips(331 + MacroX * 800), ToTwips(550 + MacroY * 800))
        MapCollision(10) = CreateLine(ToTwips(331 + MacroX * 800), ToTwips(550 + MacroY * 800), ToTwips(205 + MacroX * 800), ToTwips(550 + MacroY * 800))
        MapCollision(11) = CreateLine(ToTwips(205 + MacroX * 800), ToTwips(450 + MacroY * 800), ToTwips(0 + MacroX * 800), ToTwips(450 + MacroY * 800))
        MapCollision(12) = CreateLine(ToTwips(205 + MacroX * 800), ToTwips(550 + MacroY * 800), ToTwips(205 + MacroX * 800), ToTwips(450 + MacroY * 800))
        MapCollision(13) = CreateLine(ToTwips(600 + MacroX * 800), ToTwips(450 + MacroY * 800), ToTwips(600 + MacroX * 800), ToTwips(550 + MacroY * 800))
        MapCollision(14) = CreateLine(ToTwips(200 + MacroX * 800), ToTwips(178 + MacroY * 800), ToTwips(200 + MacroX * 800), ToTwips(331 + MacroY * 800))
        MapCollision(15) = CreateLine(ToTwips(200 + MacroX * 800), ToTwips(331 + MacroY * 800), ToTwips(0 + MacroX * 800), ToTwips(331 + MacroY * 800))
    ElseIf (BackgroundID = 1) Then
        '2D Arrays don't work here for some reason. Hence:
        Dim Pillar1() As EuclideanLine
        Dim Pillar2() As EuclideanLine
        Dim Pillar3() As EuclideanLine
        Dim Pillar4() As EuclideanLine
        Dim Statue1() As EuclideanLine
        Dim Statue2() As EuclideanLine

        Pillar1 = GetLines(ToTwips(295 + MacroX * 800), ToTwips(241 + MacroY * 800), ToTwips(43), ToTwips(96))
        Pillar2 = GetLines(ToTwips(466 + MacroX * 800), ToTwips(241 + MacroY * 800), ToTwips(43), ToTwips(96))
        Pillar3 = GetLines(ToTwips(295 + MacroX * 800), ToTwips(418 + MacroY * 800), ToTwips(43), ToTwips(96))
        Pillar4 = GetLines(ToTwips(466 + MacroX * 800), ToTwips(418 + MacroY * 800), ToTwips(43), ToTwips(96))
        
        Statue1 = GetLines(ToTwips(247 + MacroX * 800), ToTwips(31 + MacroY * 800), ToTwips(35), ToTwips(63))
        Statue2 = GetLines(ToTwips(735 + MacroX * 800), ToTwips(222 + MacroY * 800), ToTwips(35), ToTwips(63))
        
        MapCollision(0) = Pillar1(0)
        MapCollision(1) = Pillar1(1)
        MapCollision(2) = Pillar1(2)
        MapCollision(3) = Pillar1(3)
        
        MapCollision(4) = Pillar2(0)
        MapCollision(5) = Pillar2(1)
        MapCollision(6) = Pillar2(2)
        MapCollision(7) = Pillar2(3)
        
        MapCollision(8) = Pillar3(0)
        MapCollision(9) = Pillar3(1)
        MapCollision(10) = Pillar3(2)
        MapCollision(11) = Pillar3(3)
        
        MapCollision(12) = Pillar4(0)
        MapCollision(13) = Pillar4(1)
        MapCollision(14) = Pillar4(2)
        MapCollision(15) = Pillar4(3)
        
        MapCollision(16) = Statue1(0)
        MapCollision(17) = Statue1(1)
        MapCollision(18) = Statue1(2)
        MapCollision(19) = Statue1(3)
        
        MapCollision(20) = Statue2(0)
        MapCollision(21) = Statue2(1)
        MapCollision(22) = Statue2(2)
        MapCollision(23) = Statue2(3)
    End If
End Sub

Public Sub LoadNewLevel()
    Dim NewLevel As SeamlessRoom
    Dim Rng As Double
    
    Rng = Rnd()
    
    If (Rng <= 0.5) Then
        BackgroundID = 0
        ReDim MapCollision(0 To 15)
    Else
        BackgroundID = 1
        ReDim MapCollision(24)
    End If
    
    LoadRoomQueue
    
    NewLevel = GenerateLevel()
    CurrentRoom = NewLevel
End Sub

Public Function GenerateLevel() As SeamlessRoom

    Dim MaxSize As Integer

    Dim NewLevel As SeamlessRoom
    Dim I As Integer
    Dim Rng As Double
    Dim Path As String
    
    Dim Visited(0 To 20, 0 To 20) As Boolean
    
    Dim GX As Integer
    Dim GY As Integer
    
    MaxSize = 20 'Make this more flexible in later version.
    
    GX = Int(MaxSize / 2) + 1
    GY = Int(MaxSize / 2) + 1
    
    Visited(GX, GY) = True
    
    Path = ""
    NewLevel.Spot = 1
    
    For I = 0 To Int(MaxSize / 2)
        Rng = Rnd()
        If (Rng <= 0.25 And Not Visited(GX, GY - 1)) Then
            Path = Path & "U"
            Visited(GX, GY - 1) = True
            GY = GY - 1
        ElseIf (Rng <= 0.5 And Not Visited(GX + 1, GY)) Then
            Path = Path & "R"
            Visited(GX + 1, GY) = True
            GX = GX + 1
        ElseIf (Rng <= 0.75 And Not Visited(GX - 1, GY)) Then
            Path = Path & "L"
            Visited(GX - 1, GY) = True
            GX = GX - 1
        ElseIf (Not Visited(GX, GY + 1)) Then
            Path = Path & "D"
            Visited(GX, GY + 1) = True
            GY = GY + 1
        End If
    Next
    
    NewLevel.Path = Path
    GenerateLevel = NewLevel
End Function

Public Sub LoadRoomQueue()
    LoadingNext = True
End Sub
