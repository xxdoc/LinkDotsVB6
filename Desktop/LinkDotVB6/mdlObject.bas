Attribute VB_Name = "mdlObject"
'TYPE IDs:
'0 = Test Object, 1 = Player

Global Const PLAYER = 0
Global Const WALL = 1
Global Const DOT = 2
Global Const LINK = 3

Global ActiveLink As Boolean

Global LinkDots() As GameObject

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
End Type

Public Sub Create(Object As GameObject, ByVal X As Long, ByVal Y As Long, ByVal TypeID As Integer)
    Object.X = X
    Object.Y = Y
    Select Case TypeID
        Case PLAYER
            Object.SpriteID = PLAYER
            Object.SpriteFrame = 0
        Case WALL
            Object.SpriteID = WALL
            Object.SpriteFrame = 0
        Case DOT
            Object.CustomDraw = True
        Case LINK
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

Public Sub RenderObject(Object As GameObject)
    Select Case Object.TypeID
        Case DOT
            frmMain.picDisplay.FillStyle = vbFSSolid
            frmMain.picDisplay.FillColor = vbCyan
            frmMain.picDisplay.Circle (Object.X, Object.Y), 100, vbBlue
        Case LINK
            frmMain.picDisplay.DrawWidth = Val(Object.DataTag)
            frmMain.picDisplay.ForeColor = vbBlue
            For I = 0 To UBound(LinkDots) - 1
                frmMain.picDisplay.Line (LinkDots(I).X, LinkDots(I).Y)-(LinkDots((I + 1) Mod UBound(LinkDots)).X, LinkDots((I + 1) Mod UBound(LinkDots)).Y)
            Next I
            frmMain.picDisplay.DrawWidth = 1
    End Select
End Sub

Public Function CollisionCheck(Object1 As GameObject, Object2 As GameObject) As Boolean
    CollisionCheck = (Abs(Object1.X - Object2.X) * 2 < (GetSprite(Object1.SpriteID, Object1.SpriteFrame).Width + _
                    GetSprite(Object2.SpriteID, Object2.SpriteFrame).Width)) And _
                    (Abs(Object1.Y - Object2.Y) * 2 < (GetSprite(Object1.SpriteID, Object1.SpriteFrame).Width + GetSprite(Object2.SpriteID, Object2.SpriteFrame).Height))
End Function

Public Sub UpdateObject(Object As GameObject)

    Dim I As Integer

    Select Case Object.TypeID
        Case PLAYER
            Dim HSpeed As Integer
            Dim VSpeed As Integer
            
            Dim XPrev As Long
            Dim YPrev As Long
            
            Dim CanPlace As Boolean
            
            CanPlace = False
            
            If (Val(Trim$(Object.DataTag)) = 0) Then
                CanPlace = True
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
                Object.DataTag = "1"
                CreateQueue DotObj, Object.X, Object.Y, DOT
            End If
            
            If (IsKeyDown(vbKeyH) And CanPlace And Not ActiveLink) Then
                ReDim LinkDots(UBound(Dots))
                Dim LinkObj As GameObject
                Dim Life As String
                
                Life = Format$(10, "0")
                
                CanPlace = False
                Object.DataTag = "1"
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
                Object.DataTag = "0"
            End If
            
            Object.X = Object.X + HSpeed
            Object.Y = Object.Y + VSpeed

            For I = 0 To UBound(Objects) - 1
                If (Objects(I).TypeID = WALL And CollisionCheck(Object, Objects(I))) Then
                    Object.X = XPrev
                    Object.Y = YPrev
                End If
            Next I
        Case LINK
            ActiveLink = True
            Object.DataTag = Format$((Val(Object.DataTag) - 1), "0")
            If (Val(Object.DataTag) <= 0) Then
                Object.Removed = True
                ActiveLink = False
                For I = 0 To UBound(Dots) - 1
                    Objects(Dots(I).ID).Removed = True
                Next
                ReDim Dots(0)
                ReDim LinkDots(0)
            End If
    End Select
End Sub

Public Function ParseData(ByVal Data As String) As String()

End Function
