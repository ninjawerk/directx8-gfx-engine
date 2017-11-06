VERSION 5.00
Begin VB.Form Fmain 
   Caption         =   "ARCC Imagio Effect Combo : Particle Flame &  Lens Blast"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8250
   Icon            =   "Imagio 1.0.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Imagio 1.0.frx":030A
   ScaleHeight     =   463
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   550
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   3240
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2400
      Top             =   120
   End
End
Attribute VB_Name = "Fmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim im As New Imagio
Dim my As Single
Dim spl As New Particle
 

Private Sub Form_Load()
im.Screen_Prop Me.scalewidth, Me.scaleheight, Me.hwnd
im.Initialize
 End Sub
Public Function UPDATE_PARTICLES()
If spl.Lifetime = 1 Then


im.fire spl.sngX, spl.sngY, spl.sngA, spl.sngAlphaDecay, spl.sngZ, spl.XWind, spl.Gravity
im.SuperParticle Gwin.scalewidth / 2, Gwin.scaleheight / 2, 10, 1.01, 15, 2, 0.1, 20
Else
im.SuperParticle spl.sngX, spl.sngY, spl.sngA, spl.sngAlphaDecay, spl.sngZ, spl.XWind, spl.Gravity, 256
End If
 
End Function
Private Sub Form_Resize()
Timer1.Enabled = False
im.Screen_Prop Me.scalewidth, Me.scaleheight, Me.hwnd
im.Initialize
Set im.m = im.Load_Texture(55, 55, App.Path & "\vf.jpg")
Set im.cc = im.Load_Texture(55, 55, App.Path & "\vc.bmp")
Timer1.Enabled = True
End Sub
Private Sub Timer1_Timer()
DoEvents
im.Begin_Scene
UPDATE_PARTICLES
im.End_Scene
End Sub
Private Sub Timer2_Timer()
If spl.sngX < Gwin.scalewidth / 2 Then
spl.Lifetime = 1
spl.sngA = 5
spl.sngY = Gwin.scaleheight / 2
spl.sngAlphaDecay = 1.03
spl.sngZ = 10
spl.sngX = spl.sngX + 3
spl.XWind = -5
spl.Gravity = 0.4
Else
If spl.Lifetime = 1 Then spl.Lifetime = 3
spl.sngA = 15
spl.sngY = Gwin.scaleheight / 2
spl.sngAlphaDecay = 1.02
If spl.sngZ > 45 Then spl.sngZ = spl.sngZ / 1.005
spl.sngX = spl.sngX
spl.XWind = 3
spl.Gravity = 0.15
End If

If spl.Lifetime = 3 Then
If spl.sngZ < 375 Then
spl.sngZ = spl.sngZ * 1.15
Else
spl.Lifetime = 2
End If
End If
End Sub
