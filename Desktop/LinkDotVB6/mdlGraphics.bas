Attribute VB_Name = "mdlGraphics"
Option Explicit

'SpriteIDs

Global Const WIZARD = 0
Global Const BRICK = 1

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
    End Select
    GetSprite = ReturnImage
End Function
