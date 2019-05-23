Attribute VB_Name = "mdlObject"
Option Explicit

'Author: William Chan
'Date: May 17th, 2019
'Purpose: ICS4U Culminating Assignment

'TYPE IDs:
Global Const PLAYER = 0
Global Const WALL = 1
Global Const DOT = 2
Global Const LINK = 3
Global Const CHECKER = 4
Global Const ENEMY1 = 5
Global Const DOOR = 6
Global Const FLOOR = 7

Global Const MAXENEMYKNOCKBACKSTUN = 50

Const ENEMYSPEED1 = 20

Global PlayerX As Long
Global PlayerY As Long

Global Running As Boolean

Global ActiveLink As Boolean

Global LinkDots() As GameObject

Public Type Point
    X As Long
    Y As Long
End Type

Public Type EuclideanLine
    Point1 As Point
    Point2 As Point
End Type

Public Type GameObject
    ID As Integer
    TypeID As Integer
    X As Long
    Y As Long
    Depth As Integer
    DataTag As String * 50
    CustomDraw As Boolean
    Removed As Boolean
    SpriteID As Integer
    SpriteFrame As Integer
    KnockBackX As Double
    KnockBackY As Double
    AnimationSpeed As Integer
    Width As Integer
    Height As Integer
    PhantomDraw As Boolean
End Type

Public Function CreateLine(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As EuclideanLine
    Dim Point1 As Point
    Dim Point2 As Point
    Dim ReturnLine As EuclideanLine
    
    Point1.X = X1
    Point1.Y = Y1
    Point2.X = X2
    Point2.Y = Y2
    
    ReturnLine.Point1 = Point1
    ReturnLine.Point2 = Point2
    
    CreateLine = ReturnLine
End Function

Public Function OppositeDirection(ByVal Direction As String) As String
    Select Case Direction
        Case "U"
            OppositeDirection = "D"
        Case "D"
            OppositeDirection = "U"
        Case "R"
            OppositeDirection = "L"
        Case "L"
            OppositeDirection = "R"
    End Select
End Function

Public Function BackTrack(ByVal Direction As String) As Boolean
    If (CurrentRoom.Spot > 1) Then
        If (Mid$(CurrentRoom.Path, CurrentRoom.Spot - 1, 1) = OppositeDirection(Direction)) Then
            CurrentRoom.Spot = CurrentRoom.Spot - 1
            BackTrack = True
        Else
            BackTrack = False
        End If
    Else
        BackTrack = False
    End If
End Function

Public Sub Create(Object As GameObject, ByVal X As Long, ByVal Y As Long, ByVal TypeID As Integer)
    Object.X = X
    Object.Y = Y
    Select Case TypeID
        Case PLAYER
            Object.SpriteID = PLAYER
            Object.SpriteFrame = 0
            Object.DataTag = "0|500"
        Case WALL
            Object.SpriteID = CHECKERFLOOR
            Object.SpriteFrame = 0
            Object.CustomDraw = True
        Case DOT
            Object.CustomDraw = True
        Case LINK
            Object.CustomDraw = True
        Case CHECKER
            Object.SpriteID = CHECKERFLOOR
            Object.SpriteFrame = 0
            Object.CustomDraw = True
        Case ENEMY1
            Object.SpriteID = ENEMYWILLO
            Object.SpriteFrame = 0
            Object.DataTag = "300"
        Case DOOR
            Object.SpriteID = CHECKERFLOOR
            Object.SpriteFrame = 0
            Object.CustomDraw = True
    End Select
    Object.TypeID = TypeID
    AddObject Object
End Sub

Public Sub CreateQueue(Object As GameObject, ByVal X As Long, ByVal Y As Long, ByVal TypeID As Integer)
    Object.X = X
    Object.Y = Y
    Select Case TypeID
        Case PLAYER
            Object.SpriteID = PLAYER
            Object.SpriteFrame = 0
        Case WALL
            Object.SpriteID = WALL
            Object.SpriteFrame = 0
    End Select
    Object.TypeID = TypeID
    AddObjectQueue Object
End Sub

Public Function GetLines(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As EuclideanLine()
    Dim Lines(4) As EuclideanLine
    
    With Lines(0)
        .Point1.X = X
        .Point1.Y = Y
        .Point2.X = X
        .Point2.Y = Y + Height
    End With
    
    With Lines(1)
        .Point1.X = X + Width
        .Point1.Y = Y
        .Point2.X = X + Width
        .Point2.Y = Y + Height
    End With
    
    With Lines(2)
        .Point1.X = X
        .Point1.Y = Y
        .Point2.X = X + Width
        .Point2.Y = Y
    End With
    
    With Lines(3)
        .Point1.X = X
        .Point1.Y = Y + Height
        .Point2.X = X + Width
        .Point2.Y = Y + Height
    End With
    
    GetLines = Lines
End Function

Public Sub ResetDraw()
    frmMain.picDisplay.FillColor = vbBlack
    frmMain.picDisplay.ForeColor = vbBlack
    frmMain.picDisplay.DrawWidth = 1
End Sub

Public Sub RenderObject(Object As GameObject)
    Dim I As Integer
    Select Case Object.TypeID
        Case PLAYER
            frmMain.picDisplay.FillStyle = vbFSSolid
            frmMain.picDisplay.FillColor = vbBlack
            frmMain.picDisplay.Line (0, 0)-(500 * 8, 75), , B
            frmMain.picDisplay.FillColor = vbGreen
            frmMain.picDisplay.ForeColor = vbBlack
            frmMain.picDisplay.Line (0, 0)-(Val(Split(Trim$(Object.DataTag), "|")(1)) * 8, 75), , B
        Case DOT
            frmMain.picDisplay.FillStyle = vbFSSolid
            frmMain.picDisplay.FillColor = vbCyan
            frmMain.picDisplay.Circle (CamCorrectX(Object.X), CamCorrectY(Object.Y)), 100, vbBlue
        Case LINK
            frmMain.picDisplay.DrawWidth = Val(Object.DataTag)
            frmMain.picDisplay.ForeColor = vbBlue
            For I = 0 To UBound(LinkDots) - 1
                frmMain.picDisplay.Line (CamCorrectX(LinkDots(I).X), CamCorrectY(LinkDots(I).Y))-(CamCorrectX(LinkDots((I + 1) Mod UBound(LinkDots)).X), CamCorrectY(LinkDots((I + 1) Mod UBound(LinkDots)).Y))
            Next I
        Case WALL
        Case DOOR
        Case ENEMY1
        frmMain.picDisplay.FillStyle = vbFSSolid
            frmMain.picDisplay.FillColor = vbBlack
            frmMain.picDisplay.ForeColor = vbBlack
            frmMain.picDisplay.Line (CamCorrectX(Object.X - 300), CamCorrectY(Object.Y - 50))-(CamCorrectX(Object.X + 300 * 4), CamCorrectY(Object.Y - 100)), , B
            frmMain.picDisplay.FillStyle = vbFSSolid
            frmMain.picDisplay.FillColor = vbGreen
            frmMain.picDisplay.ForeColor = vbBlack
            frmMain.picDisplay.Line (CamCorrectX(Object.X - 300), CamCorrectY(Object.Y - 50))-(CamCorrectX(Object.X - 300 + Val(Object.DataTag) * 5), CamCorrectY(Object.Y - 100)), , B
        Case CHECKER
            frmMain.picDisplay.FillColor = vbWhite
            frmMain.picDisplay.ForeColor = vbBlack
            frmMain.picDisplay.Line (CamCorrectX(Object.X), CamCorrectY(Object.Y))-(CamCorrectX(Object.X + 480), CamCorrectY(Object.Y + 480)), , B
        Case FLOOR
            frmMain.picDisplay.FillColor = frmImages.shpLightGreen.FillColor
            frmMain.picDisplay.ForeColor = vbBlack
            frmMain.picDisplay.Line (CamCorrectX(Object.X), CamCorrectY(Object.Y))-(CamCorrectX(Object.X + ToTwips(32)), CamCorrectY(Object.Y + ToTwips(32))), , B
    End Select
    ResetDraw
End Sub

Public Function CollisionCheck(Object1 As GameObject, Object2 As GameObject) As Boolean
    CollisionCheck = (Abs(Object1.X - Object2.X) * 2 < (GetSprite(Object1.SpriteID, Object1.SpriteFrame).Width + _
                    GetSprite(Object2.SpriteID, Object2.SpriteFrame).Width)) And _
                    (Abs(Object1.Y - Object2.Y) * 2 < (GetSprite(Object1.SpriteID, Object1.SpriteFrame).Width + GetSprite(Object2.SpriteID, Object2.SpriteFrame).Height))
End Function

'Where the second object does not have an image
Public Function CollisionCheckImg1(Object1 As GameObject, Object2 As GameObject) As Boolean
    CollisionCheckImg1 = (Abs(Object1.X - Object2.X) * 2 < (GetSprite(Object1.SpriteID, Object1.SpriteFrame).Width + _
                    Object2.Width)) And _
                    (Abs(Object1.Y - Object2.Y) * 2 < (GetSprite(Object1.SpriteID, Object1.SpriteFrame).Width + Object2.Height))
End Function

Public Function CollisionCheckLineLine(Line1 As EuclideanLine, Line2 As EuclideanLine) As Boolean
    Dim UA As Double
    Dim UB As Double
    
    If (((Line2.Point2.Y - Line2.Point1.Y) * (Line1.Point2.X - Line1.Point1.X) _
                    - (Line2.Point2.X - Line2.Point1.X) * (Line1.Point2.Y - Line1.Point1.Y)) <> 0 _
                    And ((Line2.Point2.Y - Line2.Point1.Y) * (Line1.Point2.X - Line1.Point1.X) _
                        - (Line2.Point2.X - Line2.Point1.X) * (Line1.Point2.Y - Line1.Point1.Y)) <> 0) Then
    
        UA = ((Line2.Point2.X - Line2.Point1.X) * (Line1.Point1.Y - Line2.Point1.Y) _
                - (Line2.Point2.Y - Line2.Point1.Y) * (Line1.Point1.X - Line2.Point1.X)) _
                    / ((Line2.Point2.Y - Line2.Point1.Y) * (Line1.Point2.X - Line1.Point1.X) _
                        - (Line2.Point2.X - Line2.Point1.X) * (Line1.Point2.Y - Line1.Point1.Y))
                        
                        
        UB = ((Line1.Point2.X - Line1.Point1.X) * (Line1.Point1.Y - Line2.Point1.Y) _
                - (Line1.Point2.Y - Line1.Point1.Y) * (Line1.Point1.X - Line2.Point1.X)) _
                    / ((Line2.Point2.Y - Line2.Point1.Y) * (Line1.Point2.X - Line1.Point1.X) _
                        - (Line2.Point2.X - Line2.Point1.X) * (Line1.Point2.Y - Line1.Point1.Y))
                        
        If (UA >= 0 And UA <= 1 And UB >= 0 And UB <= 1) Then
            CollisionCheckLineLine = True
        Else
            CollisionCheckLineLine = False
        End If
    
    End If
End Function

Public Function CollisionCheckObjectLine(ObjectTest As GameObject, LineTest As EuclideanLine) As Boolean
    Dim Left As Boolean
    Dim Right As Boolean
    Dim Up As Boolean
    Dim Down As Boolean
    
    Dim ObjectLine1 As EuclideanLine
    Dim ObjectLine2 As EuclideanLine
    Dim ObjectLine3 As EuclideanLine
    Dim ObjectLine4 As EuclideanLine

    Dim ObjSprite As Sprite
    
    ObjSprite = GetSprite(ObjectTest.SpriteID, ObjectTest.SpriteFrame)
    
    'TODO: Make gameobjects use points in later version.
    
    With ObjectLine1
        .Point1.X = ObjectTest.X
        .Point1.Y = ObjectTest.Y
        .Point2.X = ObjectTest.X
        .Point2.Y = ObjectTest.Y + ObjSprite.Height
    End With
    
    With ObjectLine2
        .Point1.X = ObjectTest.X + ObjSprite.Width
        .Point1.Y = ObjectTest.Y
        .Point2.X = ObjectTest.X + ObjSprite.Width
        .Point2.Y = ObjectTest.Y + ObjSprite.Height
    End With
    
    With ObjectLine3
        .Point1.X = ObjectTest.X
        .Point1.Y = ObjectTest.Y
        .Point2.X = ObjectTest.X + ObjSprite.Width
        .Point2.Y = ObjectTest.Y
    End With
    
    With ObjectLine4
        .Point1.X = ObjectTest.X
        .Point1.Y = ObjectTest.Y + ObjSprite.Height
        .Point2.X = ObjectTest.X + ObjSprite.Width
        .Point2.Y = ObjectTest.Y + ObjSprite.Height
    End With
    
    Left = CollisionCheckLineLine(LineTest, ObjectLine1)
    Right = CollisionCheckLineLine(LineTest, ObjectLine2)
    Up = CollisionCheckLineLine(LineTest, ObjectLine3)
    Down = CollisionCheckLineLine(LineTest, ObjectLine4)
    
    If (Left Or Right Or Up Or Down _
        Or (((LineTest.Point1.X >= ObjectTest.X) And (LineTest.Point1.X <= ObjectTest.X + ObjSprite.Width)) _
            And (LineTest.Point1.Y >= ObjectTest.Y) And (LineTest.Point1.Y <= ObjectTest.Y + ObjSprite.Height) _
            And (LineTest.Point2.X >= ObjectTest.X) And (LineTest.Point2.X <= ObjectTest.X + ObjSprite.Width)) _
            And (LineTest.Point2.Y >= ObjectTest.Y) And (LineTest.Point2.Y <= ObjectTest.Y + ObjSprite.Height)) Then
        CollisionCheckObjectLine = True
    Else
        CollisionCheckObjectLine = False
    End If
End Function

Public Sub UpdateObject(Object As GameObject)

    Dim I As Integer

    Select Case Object.TypeID
        Case PLAYER
            Dim HSpeed As Integer
            Dim VSpeed As Integer
            
            Dim MacroXPrev As Integer
            Dim MacroYPrev As Integer
            
            Dim XPrev As Long
            Dim YPrev As Long
            
            Dim CanPlace As Boolean
            
            CanPlace = False

            If (Val(Trim$(Split(Object.DataTag, "|")(0))) = 0) Then
                CanPlace = True
            End If
            
            If (Val(Split(Object.DataTag, "|")(1)) <= 0 And Len(Trim$(Split(Object.DataTag, "|")(1))) >= 1) Then
                Object.Removed = True
                MsgBox "Game Over! You completed " & LevelsCompleted & " levels.", vbCritical, "Game Over!"
                Running = False
            End If
            
            Object.KnockBackX = (Object.KnockBackX / 1.5)
            Object.KnockBackY = (Object.KnockBackY / 1.5)
            
            XPrev = Object.X
            YPrev = Object.Y
            MacroXPrev = MacroX
            MacroYPrev = MacroY
            
            HSpeed = 0
            VSpeed = 0
            
            If (IsKeyDown(vbKeyW)) Then
                VSpeed = -50
            End If
            If (IsKeyDown(vbKeyA)) Then
                HSpeed = -50
            End If
            If (IsKeyDown(vbKeyS)) Then
                VSpeed = 50
            End If
            If (IsKeyDown(vbKeyD)) Then
                HSpeed = 50
            End If
            
            If (IsKeyDown(vbKeySpace) And CanPlace And Not ActiveLink And UBound(Dots) < 100) Then
                Dim DotObj As GameObject
                CanPlace = False
                Mid$(Object.DataTag, 1, 1) = "1"
                DotObj.DataTag = Format$(CurrentRoom.Spot, "0")
                CreateQueue DotObj, Object.X, Object.Y, DOT
            End If
            
            If (IsKeyDown(vbKeyH) And CanPlace And Not ActiveLink) Then
                ReDim LinkDots(UBound(Dots))
                Dim LinkObj As GameObject
                Dim Life As String
                
                Life = Format$(10, "0")
                
                CanPlace = False
                Mid$(Object.DataTag, 1, 1) = "1"
                For I = 0 To UBound(Dots) - 1
                    LinkDots(I) = Dots(I)
                Next I
                
                LinkObj.DataTag = Life
                
                CreateQueue LinkObj, 0, 0, LINK
            End If
            
            If (Not IsKeyDown(vbKeySpace) And Not IsKeyDown(vbKeyH)) Then
                CanPlace = True
                Mid$(Object.DataTag, 1, 1) = "0"
            End If
            
            Object.X = Object.X + HSpeed + Object.KnockBackX
            Object.Y = Object.Y + VSpeed + Object.KnockBackY

            PlayerX = Object.X
            PlayerY = Object.Y
            
            MacroX = Int(Object.X / ToTwips(800))
            MacroY = Int(Object.Y / ToTwips(800))
            
            Dim Direction As String

            If (MacroX - MacroXPrev <> 0 Or MacroY - MacroYPrev <> 0) Then
                If (MacroX > MacroXPrev) Then
                    Direction = "R"
                ElseIf (MacroX < MacroXPrev) Then
                    Direction = "L"
                ElseIf (MacroY > MacroYPrev) Then
                    Direction = "D"
                ElseIf (MacroY < MacroYPrev) Then
                    Direction = "U"
                End If
                
                Dim RelativeX As Long
                Dim RelativeY As Long
                
                RelativeX = 0
                RelativeY = 0
                
                If (Direction = Mid$(CurrentRoom.Path, CurrentRoom.Spot, 1)) Then
                    If (CurrentRoom.Spot < 5) Then
                        CurrentRoom.Spot = CurrentRoom.Spot + 1
                        
                        Dim NumEnemies As Integer
                        
                        NumEnemies = Int((Rnd() * 3)) + 1
                        
                        For I = 0 To NumEnemies - 1
                            Dim EnemyObject As GameObject
                            CreateQueue EnemyObject, Object.X + ToTwips(Int(Rnd() * 800)), Object.Y + ToTwips(Int(Rnd() * 800)), ENEMY1
                        Next I
                    Else
                        MsgBox "You have reached the end. Starting next level.", vbInformation
                        LevelsCompleted = LevelsCompleted + 1
                        LoadNewLevel
                        frmMain.ResetKeys
                    End If
                ElseIf (Not BackTrack(Direction)) Then
                    Select Case Direction
                        Case "R"
                            Object.X = Object.X - ToTwips(800)
                            RelativeX = -ToTwips(800)
                        Case "L"
                            Object.X = Object.X + ToTwips(800)
                            RelativeX = ToTwips(800)
                        Case "D"
                            Object.Y = Object.Y - ToTwips(800)
                            RelativeY = -ToTwips(800)
                        Case "U"
                            Object.Y = Object.Y + ToTwips(800)
                            RelativeY = ToTwips(800)
                    End Select
                    
                    For I = 0 To UBound(Objects) - 1
                        If (Objects(I).ID <> Object.ID And Objects(I).TypeID <> DOT) Then  '
                            Objects(I).X = Objects(I).X + RelativeX
                            Objects(I).Y = Objects(I).Y + RelativeY
                        End If
                    Next I
                    
                    MacroX = MacroXPrev
                    MacroY = MacroYPrev
                End If
            End If

            For I = 0 To UBound(MapCollision)
                If (CollisionCheckObjectLine(Object, MapCollision(I))) Then
                    Object.X = XPrev
                    Object.Y = YPrev
                End If
            Next I

            For I = 0 To UBound(Objects) - 1
                If (Objects(I).TypeID = WALL And CollisionCheck(Object, Objects(I))) Then
                    Object.X = XPrev
                    Object.Y = YPrev
                ElseIf (Objects(I).TypeID = ENEMY1 And CollisionCheck(Object, Objects(I))) Then
                
                    Dim DotState As String
                    Dim Health As String
                    Dim DataArr() As String
                    Dim Theta As Double
                    
                    DataArr = Split(Object.DataTag, "|")
                    
                    DotState = DataArr(0)
                    Health = DataArr(1)
                    
                    Health = Format$(Val(Health) - 50, "0")
                    
                    Object.DataTag = DotState & "|" & Health
                    Objects(I).DataTag = Format$(Val(Objects(I).DataTag) - 15, "0")
                    
                    Object.KnockBackX = (Object.X - Objects(I).X) * 0.75
                    Object.KnockBackY = (Object.Y - Objects(I).Y) * 0.75
                    
                    Objects(I).KnockBackX = -Object.KnockBackX * 0.75
                    Objects(I).KnockBackY = -Object.KnockBackY * 0.75
                    
                End If
            Next I
            
            CamX = Object.X + CAMXOFFSET
            CamY = Object.Y + CAMYOFFSET
        Case LINK
            ActiveLink = True
            Object.DataTag = Format$((Val(Object.DataTag) - 1), "0")
            If (Val(Object.DataTag) <= 0) Then
                Object.Removed = True
                ActiveLink = False
                For I = 0 To UBound(Objects) - 1
                    If (Objects(I).TypeID = DOT) Then
                        Objects(I).Removed = True
                    End If
                Next
                ReDim Dots(0)
                ReDim LinkDots(0)
            End If
        Case ENEMY1
            Object.AnimationSpeed = ((Object.AnimationSpeed + 1) Mod 10) + 1
            Object.SpriteFrame = (Object.SpriteFrame + (Int(Object.AnimationSpeed / 10))) Mod 4
        
            Object.KnockBackX = Object.KnockBackX / 1.5
            Object.KnockBackY = Object.KnockBackY / 1.5
            
            If (Int(Abs(Object.KnockBackX)) > 0 Or Int(Abs(Object.KnockBackY)) > 0) Then
                Object.SpriteFrame = 4
            End If
            
            XPrev = Object.X
            YPrev = Object.Y
        
            If (Val(Trim$(Object.DataTag)) <= 0) Then
                Object.Removed = True
            End If
            
            Dim DeltaX As Double
            Dim DeltaY As Double
            Dim ScaleDelta As Double
            
            For I = 0 To UBound(Objects) - 1
                If (CollisionCheck(Object, Objects(I)) And (Objects(I).TypeID = ENEMY1) And I <> Object.ID) Then
                    Object.KnockBackX = -(Objects(I).X - Object.X)
                    Object.KnockBackY = -(Objects(I).Y - Object.Y)
                    Objects(I).KnockBackX = (Objects(I).X - Object.X)
                    Objects(I).KnockBackY = (Objects(I).Y - Object.Y)
                    Object.DataTag = Format$(Val(Object.DataTag) - 30, "0")
                    Objects(I).DataTag = Format$(Val(Objects(I).DataTag) - 30, "0")
                End If
                
                'Follow
            
                If (Objects(I).TypeID = PLAYER) Then
                
                    DeltaX = Object.X - Objects(I).X
                    DeltaY = Object.Y - Objects(I).Y
                    
                    ScaleDelta = ENEMYSPEED1 * 1.25 / (Sqr(DeltaX ^ 2 + DeltaY ^ 2))
                    
                    Object.X = Object.X - (DeltaX * ScaleDelta)
                    Object.Y = Object.Y - (DeltaY * ScaleDelta)
                End If
            Next I
            
            Object.X = Object.X + Object.KnockBackX
            Object.Y = Object.Y + Object.KnockBackY
            
            If (ActiveLink) Then
                Dim Point1 As Point
                Dim Point2 As Point
                Dim LinkLine As EuclideanLine
                For I = 0 To UBound(Dots) - 1
                    Point1.X = Dots(I).X
                    Point1.Y = Dots(I).Y
                    Point2.X = Dots((I + 1) Mod UBound(Dots)).X
                    Point2.Y = Dots((I + 1) Mod UBound(Dots)).Y
                    LinkLine.Point1 = Point1
                    LinkLine.Point2 = Point2
                    If (CollisionCheckObjectLine(Object, LinkLine)) Then
                        Object.DataTag = Format$(Val(Object.DataTag) - 20, "0")
                        Object.KnockBackX = (XPrev - Object.X) * ToTwips(1.5) * (MAXENEMYKNOCKBACKSTUN / (Sqr((XPrev - Object.X) ^ 2 + (YPrev - Object.Y) ^ 2)))
                        Object.KnockBackY = (YPrev - Object.Y) * ToTwips(1.5) * (MAXENEMYKNOCKBACKSTUN / (Sqr((XPrev - Object.X) ^ 2 + (YPrev - Object.Y) ^ 2)))
                    End If
                Next
            End If
    End Select
End Sub
