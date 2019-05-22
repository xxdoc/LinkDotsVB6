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
   Begin VB.Label Label1 
      Caption         =   "How are you seeing this?"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   3120
      Width           =   3495
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Author: William Chan
'Date: May 17th, 2019
'Purpose: ICS4U Culminating Assignment

Public Sub DebugPrint(ByVal St As String)
    Cls
    Print St
End Sub

