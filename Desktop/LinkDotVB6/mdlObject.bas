Attribute VB_Name = "mdlObject"
Option Explicit

'TYPE IDs:
'0 = Test Object, 1 = Player

Global Const Player = 0
Global Const WALL = 1
Global Const DOT = 2
Global Const LINK = 3
Global Const CHECKER = 4
Global Const ENEMY1 = 5
Global Const DOOR = 6

Global PlayerX As Long
Global PlayerY As Long

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
    KnockBackX As Long
    KnockBackY As Long
End Type

Public Sub Create(Object As GameObject, ByVal X As Long, ByVal Y As Long, ByVal TypeID As Integer)
    Object.X = X
    Object.Y = Y
    Select Case TypeID
        Case Player
            Object.SpriteID = Player
            Object.SpriteFrame = 0
            Object.DataTag = "0|500"
        Case WALL
            Object.SpriteID = CHECKERFLOOR
            Object.SpriteFrame = 0
            'Object.CustomDraw = True
        Case DOT
            Object.CustomDraw = True
        Case LINK
            Object.CustomDraw = True
        Case CHECKER
            'Object.SpriteID = CHECKERFLOOR
            'Object.SpriteFrame = 0
            Object.CustomDraw = True
        Case ENEMY1
            Object.SpriteID = ENEMYKNIGHT
            Object.SpriteFrame = 0
            Object.DataTag = "300"
        Case DOOR
            Object.SpriteID = CHECKERFLOOR
            Object.SpriteFrame = 0
            'Object.CustomDraw = True
    End Select
    Object.TypeID = TypeID
    AddObject Object
End Sub

Public Sub CreateQueue(Object As GameObject, ByVal X As Long, ByVal Y As Long, ByVal TypeID As Integer)
    Object.X = X
    Object.Y = Y
    Select Case TypeID
        Case Player
            Object.SpriteID = Player
            Object.SpriteFrame = 0
        Case WALL
            Object.SpriteID = WALL
            Object.SpriteFrame = 0
    End Select
    Object.TypeID = TypeID
    AddObjectQueue Object
End Sub

Public Sub ResetDraw()
    frmMain.picDisplay.FillColor = vbBlack
    frmMain.picDisplay.ForeColor = vbBlack
    frmMain.picDisplay.DrawWidth = 1
End Sub

Public Sub RenderObject(Object As GameObject)
    Dim I As Integer
    Select Case Object.TypeID
        Case Player
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
          '  frmMain.picDisplay.FillStyle = vbFSSolid
          '  frmMain.picDisplay.FillColor = vbBlack
          '  frmMain.picDisplay.Line (CamCorrectX(Object.X), CamCorrectY(Object.Y))-(CamCorrectX(Object.X + 500), CamCorrectY(Object.Y + 500)), , B
        Case ENEMY1
            frmMain.picDisplay.FillStyle = vbFSSolid
            frmMain.picDisplay.FillColor = vbGreen
            frmMain.picDisplay.ForeColor = vbBlack
            frmMain.picDisplay.Line (CamCorrectX(Object.X - 300), CamCorrectY(Object.Y - 50))-(CamCorrectX(Object.X - 300 + Val(Object.DataTag) * 5), CamCorrectY(Object.Y - 100)), , B
        Case CHECKER
            frmMain.picDisplay.FillColor = vbWhite
            frmMain.picDisplay.ForeColor = vbBlack
            frmMain.picDisplay.Line (CamCorrectX(Object.X), CamCorrectY(Object.Y))-(CamCorrectX(Object.X + 480), CamCorrectY(Object.Y + 480)), , B
    End Select
    ResetDraw
End Sub

Public Function CollisionCheck(Object1 As GameObject, Object2 As GameObject) As Boolean
    CollisionCheck = (Abs(Object1.X - Object2.X) * 2 < (GetSprite(Object1.SpriteID, Object1.SpriteFrame).Width + _
                    GetSprite(Object2.SpriteID, Object2.SpriteFrame).Width)) And _
                    (Abs(Object1.Y - Object2.Y) * 2 < (GetSprite(Object1.SpriteID, Object1.SpriteFrame).Width + GetSprite(Object2.SpriteID, Object2.SpriteFrame).Height))
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
    
    If (Left Or Right Or Up Or Down) Then
        CollisionCheckObjectLine = True
    Else
        CollisionCheckObjectLine = False
    End If
End Function

Public Sub UpdateObject(Object As GameObject)

    Dim I As Integer

    Select Case Object.TypeID
        Case Player
            Dim HSpeed As Integer
            Dim VSpeed As Integer
            
            Dim XPrev As Long
            Dim YPrev As Long
            
            Dim CanPlace As Boolean
            
            CanPlace = False
            
            frmDebug.DebugPrint Object.X
            
            If (Val(Trim$(Split(Object.DataTag, "|")(0))) = 0) Then
                CanPlace = True
            End If
            
            If (Val(Split(Object.DataTag, "|")(1)) <= 0) Then
                Object.Removed = True
                MsgBox "Game Over!", vbCritical
            End If
            
            If (Object.KnockBackX <> 0) Then
                Object.KnockBackX = Object.KnockBackX / 5
            End If
            
            If (Object.KnockBackY <> 0) Then
                Object.KnockBackY = Object.KnockBackY / 5
            End If
            
            XPrev = Object.X
            YPrev = Object.Y
            
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
            
            If (IsKeyDown(vbKeySpace) And CanPlace And Not ActiveLink) Then
                Dim DotObj As GameObject
                CanPlace = False
                Mid$(Object.DataTag, 1, 1) = "1"
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
               ' MsgBox Val(LinkObj.DataTag) & " AYSDSADSA"
                'ReDim Dots(0)
            End If
            
            If (Not IsKeyDown(vbKeySpace) And Not IsKeyDown(vbKeyH)) Then
                CanPlace = True
                Mid$(Object.DataTag, 1, 1) = "0"
            End If
            
            Object.X = Object.X + HSpeed + Object.KnockBackX
            Object.Y = Object.Y + VSpeed + Object.KnockBackY

            PlayerX = Object.X
            PlayerY = Object.Y
            
            MacroX = Int(Object.X / LEVELWIDTH)
            MacroY = Int(Object.Y / LEVELHEIGHT)
            
            'frmDebug.DebugPrint MacroX & " " & MacroY & " " & Object.X & " " & Object.Y
            frmDebug.DebugPrint UBound(Objects)

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
                    
                    Object.KnockBackX = (Object.X - Objects(I).X) * 3
                    Object.KnockBackY = (Object.Y - Objects(I).Y) * 3
                    
                   ' frmDebug.DebugPrint KnockBackX
                    
                    'Object.X = XPrev
                    'Object.Y = YPrev
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
            If (Val(Trim$(Object.DataTag)) <= 0) Then
                Object.Removed = True
            End If
            
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
                       ' Object.Removed = True
                        'RemoveObjectIndexQueue Object.ID
                        'MsgBox "CRITICAL HIT!"
                        Object.DataTag = Format$(Val(Object.DataTag) - 10, "0")
                    End If
                Next
            End If
    End Select
End Sub

Public Function ParseData(ByVal Data As String) As String()

End Function
