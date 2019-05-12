Attribute VB_Name = "mdlMain"
Global Objects() As GameObject
Global ObjCreateQueue() As GameObject
Global ObjRemoveQueue() As GameObject
Global Dots() As GameObject
Global CamX, CamY As Long

Global Const CAMXOFFSET = 200
Global Const CAMYOFFSET = 700

Public Function CamCorrectX(ByVal X As Long) As Long
    CamCorrectX = (frmMain.Width / 2) + (X - CamX)
End Function

Public Function CamCorrectY(ByVal Y As Long) As Long
    CamCorrectY = (frmMain.Height / 2) + (Y - CamY)
End Function

Public Sub Init()
    ReDim Objects(0)
    ReDim ObjCreateQueue(0)
    ReDim ObjRemoveQueue(0)
    ReDim Dots(0)
    CamX = 3500
    CamY = 1700
End Sub

Public Function IsKeyDown(ByVal Key As Integer) As Boolean
    IsKeyDown = frmMain.IsKeyDownMAIN(Key)
End Function

Public Sub AddDot(DOT As GameObject)
    Dim Upper As Integer
    Upper = UBound(Dots) + 1
    
    ReDim Preserve Dots(Upper)

    Dots(UBound(Dots) - 1) = DOT
End Sub

Public Sub AddObject(GObject As GameObject)
    Dim Upper As Integer
    Upper = UBound(Objects) + 1
    
    ReDim Preserve Objects(Upper)
    
    GObject.ID = UBound(Objects) - 1

    Objects(UBound(Objects) - 1) = GObject
End Sub

Public Sub RemoveObjectQueue(GObject As GameObject)
    Dim Upper As Integer
    Upper = UBound(ObjRemoveQueue) + 1
    
    ReDim Preserve ObjRemoveQueue(Upper)

    ObjRemoveQueue(UBound(ObjRemoveQueue) - 1) = GObject
End Sub

Public Sub AddObjectQueue(GObject As GameObject)
    Dim Upper As Integer
    Upper = UBound(ObjCreateQueue) + 1
    
    ReDim Preserve ObjCreateQueue(Upper)
    
    ObjCreateQueue(UBound(ObjCreateQueue) - 1) = GObject
End Sub

Public Function GetObject(ByVal Index As Integer) As GameObject
    GetObject = Objects(Index)
End Function


Public Sub RemoveObject(Object As GameObject)
    Dim I As Integer
    Dim Lower As Integer
    Dim Index As Integer
    
    Index = Object.ID
    
    For I = Index To UBound(Objects) - 2
        Objects(I) = Objects(I + 1)
    Next I
    Lower = UBound(Objects) - 1
    ReDim Preserve Objects(Lower)
End Sub

Public Sub RemoveObjectIndex(ByVal Index As Integer)
    Dim I As Integer
    Dim Lower As Integer
        For I = Index To UBound(Objects) - 2
        Objects(I) = Objects(I + 1)
    Next I
    Lower = UBound(Objects) - 1
    ReDim Preserve Objects(Lower)
End Sub

Public Sub RemoveObjectIndexQueue(ByVal Index As Integer)
    Dim I As Integer
    Dim Lower As Integer
        For I = Index To UBound(ObjCreateQueue) - 2
        ObjCreateQueue(I) = ObjCreateQueue(I + 1)
    Next I
    Lower = UBound(ObjCreateQueue) - 1
    ReDim Preserve ObjCreateQueue(Lower)
End Sub

Public Sub SortZOrder(Objects() As GameObject)

End Sub

Public Sub Update()
    Dim I As Integer
    Dim S As String
    S = ""
    For I = 0 To UBound(Objects) - 1
        UpdateObject Objects(I)
        Objects(I).ID = I
        S = S & Objects(I).TypeID & " "
        If (Objects(I).Removed) Then
            RemoveObjectQueue Objects(I)
        End If
    Next I
    For I = 0 To UBound(Dots) - 1
        'S = S & Dots(I).ID & " "
    Next I
    frmDebug.DebugPrint S
    For I = 0 To UBound(ObjCreateQueue) - 1
        Create ObjCreateQueue(I), ObjCreateQueue(I).X, ObjCreateQueue(I).Y, ObjCreateQueue(I).TypeID
        If (ObjCreateQueue(I).TypeID = DOT) Then
            AddDot ObjCreateQueue(I)
        End If
    Next I
    For I = 0 To UBound(ObjRemoveQueue) - 1
        RemoveObject ObjRemoveQueue(I)
    Next I
    ReDim ObjCreateQueue(0)
    ReDim ObjRemoveQueue(0)
End Sub

Public Sub Render()
    Dim I As Integer
    Dim Obj As GameObject
    For I = 0 To UBound(Objects) - 1
        Obj = Objects(I)
        If (Not Obj.CustomDraw) Then
            Dim DrawX, DrawY As Integer
            DrawX = CamCorrectX(Obj.X)
            DrawY = CamCorrectY(Obj.Y)
            DrawImage GetSprite(Obj.SpriteID, Obj.SpriteFrame), DrawX, DrawY
        Else
            RenderObject Obj
        End If
    Next I
End Sub

Public Function DataToString(Data() As String)
    Dim I As Integer
    Dim St As String
    
    St = ""
    
    For I = 0 To Data.UBound - 1
        St = St & Data(I) & "|"
    Next I
    
    DataToString = St
End Function
