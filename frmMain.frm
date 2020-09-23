VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Mario Just Won The Lottery"
   ClientHeight    =   4050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5790
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            '// Exit
            mbRunning = False
        Case vbKeyAdd
            LIGHT_MAXSTEPS = LIGHT_MAXSTEPS + 1
        Case vbKeySubtract
            If LIGHT_MAXSTEPS > 1 Then
                LIGHT_MAXSTEPS = LIGHT_MAXSTEPS - 1
            End If
        Case vbKeyA
            If LIGHT_BEGIN < 255 Then
                LIGHT_BEGIN = LIGHT_BEGIN + 5
            End If
        Case vbKeyZ
            If LIGHT_BEGIN >= 5 Then
                LIGHT_BEGIN = LIGHT_BEGIN - 5
            End If
        Case vbKeyS
            If LIGHT_END < 255 Then
                LIGHT_END = LIGHT_END + 5
            End If
        Case vbKeyX
            If LIGHT_END >= 5 Then
                LIGHT_END = LIGHT_END - 5
            End If
    End Select
End Sub


