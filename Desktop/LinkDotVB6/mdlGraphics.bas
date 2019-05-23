Attribute VB_Name = "mdlGraphics"
Option Explicit

'Author: William Chan
'Date: May 17th, 2019
'Purpose: ICS4U Culminating Assignment

'SpriteIDs

Global Const WIZARD = 0
Global Const BRICK = 1
Global Const CHECKERFLOOR = 2
Global Const ENEMYKNIGHT = 3
Global Const ENEMYWILLO = 4

Public Type Sprite
    Width As Integer
    Height As Integer
    Pic As Picture
End Type

Public Sub DrawImage(Img As Sprite, ByVal X As Integer, ByVal Y As Integer)
    frmMain.picDisplay.PaintPicture Img.Pic, X, Y, Img.Width, Img.Height
End Sub

Public Function GetSprite(ByVal SpriteID As Integer, ByVal Frame As Integer) As Sprite
    Dim ReturnImage As Sprite
    Select Case SpriteID
        Case WIZARD
            Set ReturnImage.Pic = frmImages.imgWizard(Frame).Picture
            ReturnImage.Width = frmImages.imgWizard(Frame).Width
            ReturnImage.Height = frmImages.imgWizard(Frame).Height
        Case BRICK
            Set ReturnImage.Pic = frmImages.imgBrick(Frame).Picture
            ReturnImage.Width = frmImages.imgBrick(Frame).Width
            ReturnImage.Height = frmImages.imgBrick(Frame).Height
        Case CHECKERFLOOR
            Set ReturnImage.Pic = frmImages.imgCheckerFloor(Frame).Picture
            ReturnImage.Width = frmImages.imgCheckerFloor(Frame).Width
            ReturnImage.Height = frmImages.imgCheckerFloor(Frame).Height
        Case ENEMYKNIGHT
            Set ReturnImage.Pic = frmImages.imgEnemyKnight(Frame).Picture
            ReturnImage.Width = frmImages.imgEnemyKnight(Frame).Width
            ReturnImage.Height = frmImages.imgEnemyKnight(Frame).Height
        Case ENEMYWILLO
            Set ReturnImage.Pic = frmImages.imgWilloWisp(Frame).Picture
            ReturnImage.Width = frmImages.imgWilloWisp(Frame).Width
            ReturnImage.Height = frmImages.imgWilloWisp(Frame).Height
    End Select
    GetSprite = ReturnImage
End Function
