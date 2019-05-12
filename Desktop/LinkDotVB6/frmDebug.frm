VERSION 5.00
Begin VB.Form frmDebug 
   Caption         =   "Form1"
   ClientHeight    =   4140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub DebugPrint(ByVal St As String)
    Cls
    Print St
End Sub
