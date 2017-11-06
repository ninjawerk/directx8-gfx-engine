VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Particle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
 
Option Explicit

'Values Stored For Each Particle
Private mvarused As Boolean
Private mvarsngX As Single
Private mvarsngY As Single
Private mvarsngXSpeed As Single
Private mvarsngYSpeed As Single
Private mvarsngXAccel As Single
Private mvarsngYAccel As Single
Private mvarsngR As Single
Private mvarsngG As Single
Private mvarsngB As Single
Private mvarsngA As Single
Private mvarsngAlphaDecay As Single
'local variable(s) to hold property value(s)
Private mvarXWind As Single 'local copy
Private mvarYWind As Single 'local copy
Private mvarGravity As Single 'local copy
'local variable(s) to hold property value(s)
  'local copy
'local variable(s) to hold property value(s)
Private mvarCreated As Long 'local copy
Private mvarLifetime As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarsngZ As Single 'local copy
Private mvarsngZSpeed As Single 'local copy
Public Property Let sngZSpeed(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sngZSpeed = 5
    mvarsngZSpeed = vData
End Property


Public Property Get sngZSpeed() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sngZSpeed
    sngZSpeed = mvarsngZSpeed
End Property



Public Property Let sngZ(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sngZ = 5
    mvarsngZ = vData
End Property


Public Property Get sngZ() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sngZ
    sngZ = mvarsngZ
End Property



Public Property Let Lifetime(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Lifetime = 5
    mvarLifetime = vData
End Property


Public Property Get Lifetime() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Lifetime
    Lifetime = mvarLifetime
End Property



Public Property Let Created(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Created = 5
    mvarCreated = vData
End Property


Public Property Get Created() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Created
    Created = mvarCreated
End Property







Public Property Let Gravity(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Gravity = 5
    mvarGravity = vData
End Property


Public Property Get Gravity() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Gravity
    Gravity = mvarGravity
End Property



Public Property Let YWind(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.YWind = 5
    mvarYWind = vData
End Property


Public Property Get YWind() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.YWind
    YWind = mvarYWind
End Property



Public Property Let XWind(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.XWind = 5
    mvarXWind = vData
End Property


Public Property Get XWind() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.XWind
    XWind = mvarXWind
End Property




Public Sub ResetColor(sngRed As Single, sngGreen As Single, sngBlue As Single, sngAlpha As Single, sngDecay As Single)
'************************************************************
'Change the particle's color values
'More info: http://www.vbgore.com/CommonCode.Particle.ResetColor
'************************************************************

    sngR = sngRed
    sngG = sngGreen
    sngB = sngBlue
    sngA = sngAlpha
    sngAlphaDecay = sngDecay

End Sub

Public Sub ResetIt(X As Single, Y As Single, XSpeed As Single, YSpeed As Single, XAcc As Single, YAcc As Single)
'************************************************************
'Reset the particle's variables
'More info: http://www.vbgore.com/CommonCode.Particle.ResetIt
'************************************************************

    sngX = X
    sngY = Y
    sngXSpeed = XSpeed
    sngYSpeed = YSpeed
    sngXAccel = XAcc
    sngYAccel = YAcc

End Sub

Public Property Get sngA() As Single

    sngA = mvarsngA

End Property

Public Property Let sngA(ByVal vData As Single)

    mvarsngA = vData

End Property

Public Property Get sngAlphaDecay() As Single

    sngAlphaDecay = mvarsngAlphaDecay

End Property

Public Property Let sngAlphaDecay(ByVal vData As Single)

    mvarsngAlphaDecay = vData

End Property

Public Property Get sngB() As Single

    sngB = mvarsngB

End Property

Public Property Let sngB(ByVal vData As Single)

    mvarsngB = vData

End Property

Public Property Get sngG() As Single

    sngG = mvarsngG

End Property

Public Property Let sngG(ByVal vData As Single)

    mvarsngG = vData

End Property

Public Property Get sngR() As Single

    sngR = mvarsngR

End Property

Public Property Let sngR(ByVal vData As Single)

    mvarsngR = vData

End Property

Public Property Get sngX() As Single

    sngX = mvarsngX

End Property

Public Property Let sngX(ByVal vData As Single)

    mvarsngX = vData

End Property

Public Property Get sngXAccel() As Single

    sngXAccel = mvarsngXAccel

End Property

Public Property Let sngXAccel(ByVal vData As Single)

    mvarsngXAccel = vData

End Property

Public Property Get sngXSpeed() As Single

    sngXSpeed = mvarsngXSpeed

End Property

Public Property Let sngXSpeed(ByVal vData As Single)

    mvarsngXSpeed = vData

End Property

Public Property Get sngY() As Single

    sngY = mvarsngY

End Property

Public Property Let sngY(ByVal vData As Single)

    mvarsngY = vData

End Property

Public Property Get sngYAccel() As Single

    sngYAccel = mvarsngYAccel

End Property

Public Property Let sngYAccel(ByVal vData As Single)

    mvarsngYAccel = vData

End Property

Public Property Get sngYSpeed() As Single

    sngYSpeed = mvarsngYSpeed

End Property

Public Property Let sngYSpeed(ByVal vData As Single)

    mvarsngYSpeed = vData

End Property

 

Public Property Get Used() As Boolean

    Used = mvarused

End Property

Public Property Let Used(ByVal vData As Boolean)

    mvarused = vData

End Property

