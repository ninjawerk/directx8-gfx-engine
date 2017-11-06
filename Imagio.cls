VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Imagio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Directx Declarations
Private DirectX8 As DirectX8
Private Direct3D As Direct3D8
Private Direct3D_Device As Direct3DDevice8
Private Fullscreen_Width As Long
Private Fullscreen_Height As Long
Private Fullscreen_Enabled As Boolean
Private Running As Boolean
Public m As Direct3DTexture8
Public cc As Direct3DTexture8
Private Scalar As D3DVECTOR2
Private Direct3DX As D3DX8
Private Vx4Vertex_List(3) As TLVERTEX
Private Vertex_List(3) As TLVERTEX
Private LineVertex_List(3) As TLVERTEX
Dim Irow(4096) As Single
Public lock_Rend As Boolean
 Const FVF As Integer = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
Dim ps(1024) As New Particle
Dim ss(1024) As New Particle
Dim wa(1024) As New Particle
Public Function Screen_Prop(Sw As Long, Sh As Long, hwnd As Long) 'GWIN Properties
Gwin.hwnd = hwnd
Gwin.scaleheight = Sh
Gwin.scalewidth = Sw
End Function
Public Function update_colors(index As Integer, v As Single, pr() As Particle)
On Error Resume Next
Dim r As Single
r = Rnd * 10
If pr(index).sngA > 1 Then pr(index).sngA = pr(index).sngA / v
If pr(index).sngR > 1 Then pr(index).sngR = pr(index).sngR / v
If pr(index).sngG > 1 Then pr(index).sngG = pr(index).sngG / v
If pr(index).sngB > 1 Then pr(index).sngB = pr(index).sngB / v
End Function
Public Function fire(X As Single, Y As Single, baseArea As Single, Decay As Single, particle_size As Single, wind As Single, Gravity As Single)
Dim i As Integer
Dim mi As Integer
If Rnd * 2 > 0.5 Then
mi = -1
Else
mi = 1
End If
For i = 0 To 256

If ps(i).Used = True Then
ps(i).sngXSpeed = ps(i).sngXSpeed + (ps(i).sngXAccel) / 100 * 2 + ps(i).XWind + (1) * Val(wind) / 100
ps(i).sngX = ps(i).sngX + ps(i).sngXSpeed
ps(i).sngYSpeed = ps(i).sngYSpeed + (ps(i).sngYAccel * Rnd) / -0.5
ps(i).sngY = ps(i).sngY + ps(i).sngYSpeed + ps(i).Gravity
update_colors i, Decay, ps
 Draw_Vx4 ps(i).sngX, ps(i).sngY, particle_size, particle_size, ps(i).sngR, ps(i).sngB, ps(i).sngG, m
'If GetTickCount >= ps(i).Created + ps(i).Lifetime Then ps(i).Used = False
If (ps(i).sngA) < 3 Then ps(i).Used = False
If (ps(i).sngY) > Gwin.scaleheight Then ps(i).Used = False
 If (ps(i).sngX) < 0 Then ps(i).Used = False
 If (ps(i).sngX) > Gwin.scalewidth Then ps(i).Used = False
 If (ps(i).sngY) < 0 Then ps(i).Used = False
Else
ps(i).Created = GetTickCount
ps(i).Lifetime = Rnd * Rnd * 2000 * Rnd
ps(i).Gravity = Val(Gravity)
ps(i).XWind = Rnd / 40 * Val(wind)
ps(i).ResetIt 0, 0, 0, 0, 0, 0
ps(i).Used = True
ps(i).sngXAccel = 0.5 * mi * Val(wind) / 100
ps(i).sngYAccel = 0
ps(i).sngXSpeed = Rnd * mi / 5 * Val(wind) / 100
ps(i).sngYSpeed = -Rnd
ps(i).sngB = Rnd(20) * 250
ps(i).sngG = Rnd(2156) * 250
ps(i).sngR = Rnd(256) * 250
ps(i).sngA = Rnd(256) * 250
ps(i).sngX = X + -Rnd * baseArea
ps(i).sngY = Y + -Rnd * mi
 ps(i).sngAlphaDecay = Rnd * 5
End If
Next i
my = my + 0.5
End Function
Public Function water(X As Single, Y As Single, baseArea As Single, Decay As Single, particle_size As Single, we As Single, g As Single)
Dim i As Integer
Dim mi As Integer
If Rnd * 2 > 0.9 Then
mi = -1
Else
mi = 1
End If
For i = 0 To 555

If wa(i).Used = True Then
wa(i).sngXSpeed = wa(i).sngXSpeed
wa(i).sngX = wa(i).sngX + wa(i).sngXSpeed + wa(i).XWind
wa(i).sngYSpeed = wa(i).sngYSpeed + (wa(i).sngYAccel * Rnd) / -0.5 + wa(i).Gravity
wa(i).sngY = wa(i).sngY + wa(i).sngYSpeed + wa(i).YWind
update_colors i, Decay, wa
 Draw_Vx4 wa(i).sngX, wa(i).sngY, particle_size, particle_size, wa(i).sngR, wa(i).sngB, wa(i).sngG, m
'If GetTickCount >= wa(i).Created + wa(i).Lifetime Then wa(i).Used = False
If (wa(i).sngY) > Gwin.scaleheight Then wa(i).Used = False
 If (wa(i).sngX) < 0 Then wa(i).Used = False
 If (wa(i).sngX) > Gwin.scalewidth Then wa(i).Used = False

Else
wa(i).Created = GetTickCount
wa(i).Lifetime = Rnd * Rnd * 2000 * Rnd
wa(i).Gravity = 0.2 * Rnd * Val(g)
wa(i).XWind = (Rnd + 1) * Val(we)
wa(i).ResetIt 0, 0, 0, 0, 0, 0
wa(i).Used = True
wa(i).sngXAccel = 0.5 * mi
wa(i).sngYAccel = Rnd(2) / 20
wa(i).sngXSpeed = Rnd * mi / 5
wa(i).sngYSpeed = -Rnd
wa(i).sngB = Rnd(2156) * 200
wa(i).sngG = wa(i).sngB
wa(i).sngR = wa(i).sngB
wa(i).sngA = Rnd(256) * 250
wa(i).sngX = X + Rnd * baseArea
wa(i).sngY = Y
 wa(i).sngAlphaDecay = Rnd * 5
End If
Next i
my = my + 0.5
End Function
Public Function SuperParticle(X As Single, Y As Single, baseArea As Single, Decay As Single, particle_size As Single, we As Single, g As Single, num As Long)
Dim i As Integer
Dim mi As Integer
If Rnd * 2 > 1.5 Then
mi = -1
Else
mi = 1
End If
Dim gi As Integer
If Rnd * 2 > 1.5 Then
gi = -1
Else
gi = 1
End If
Dim zi As Integer
If Rnd * 2 > 1.5 Then
zi = -1
Else
zi = 1
End If
For i = 0 To num

If ss(i).Used = True Then
ss(i).sngXSpeed = ss(i).sngXSpeed + ss(i).XWind
ss(i).sngX = ss(i).sngX + ss(i).sngXSpeed
ss(i).sngYSpeed = ss(i).sngYSpeed + (ss(i).sngYAccel * Rnd) / -0.5 + ss(i).Gravity
ss(i).sngY = ss(i).sngY + ss(i).sngYSpeed + ss(i).YWind
update_colors i, Decay, ss
ss(i).sngZ = ss(i).sngZ - ss(i).sngZSpeed
If ss(i).sngZ > 0 Then Draw_Vx4Acolor ss(i).sngX, ss(i).sngY, ss(i).sngZ, ss(i).sngZ, D3DColorRGBA(ss(i).sngR, ss(i).sngG, ss(i).sngB, 255), D3DColorRGBA(ss(i).sngR, ss(i).sngG, ss(i).sngB, 255), D3DColorRGBA(ss(i).sngR, ss(i).sngG, ss(i).sngB, 255), D3DColorRGBA(ss(i).sngR, ss(i).sngG, ss(i).sngB, 255), m

'If GetTickCount >= ss(i).Created + ss(i).Lifetime Then ss(i).Used = False
If (ss(i).sngA) < 3 Then ss(i).Used = False
If (ss(i).sngY) > Gwin.scaleheight Then ss(i).Used = False
 If (ss(i).sngX) < 0 Then ss(i).Used = False
 If (ss(i).sngX) > Gwin.scalewidth Then ss(i).Used = False
 If (ss(i).sngY) < 0 Then ss(i).Used = False

Else
ss(i).Created = GetTickCount
ss(i).Lifetime = Rnd * Rnd * 2000 * Rnd
ss(i).Gravity = Val(g) * Rnd * gi
ss(i).XWind = Rnd / 40 * Val(we) * mi
ss(i).ResetIt 0, 0, 0, 0, 0, 0
ss(i).Used = True
ss(i).sngXAccel = 0.5 * mi * Val(we) / 100 * mi
ss(i).sngYAccel = Rnd(2) / 20
ss(i).sngXSpeed = Rnd * mi / 5 * Val(we) / 100
ss(i).sngYSpeed = -Rnd
ss(i).sngB = Rnd(20) * 120
ss(i).sngG = Rnd(20) * 150
ss(i).sngR = Rnd(20) * 250
ss(i).sngA = Rnd(256) * 250
ss(i).sngX = X + -Rnd * baseArea
ss(i).sngY = Y + 0
ss(i).sngZ = Rnd * particle_size
 ss(i).sngAlphaDecay = Rnd
 ss(i).sngZSpeed = Rnd(1) * zi / 2
End If
Next i
Draw_Vx4Acolor 0, 0, Gwin.scalewidth * 2, Gwin.scaleheight * 2, 0, D3DColorRGBA(255, 50, 150, 255), 0, D3DColorRGBA(180, 40, 100, 205), Nothing

my = my + 0.5
End Function
Private Function Create_TLVertex(X As Single, Y As Single, z As Single, RHW As Single, Color As Long, Specular As Long, TU As Single, tv As Single) As TLVERTEX
    Create_TLVertex.X = X
    Create_TLVertex.Y = Y
    Create_TLVertex.z = z
    Create_TLVertex.RHW = RHW
    Create_TLVertex.Color = Color
    Create_TLVertex.Specular = Specular
    Create_TLVertex.TU = TU
    Create_TLVertex.tv = tv
End Function

Public Function Initialize()
 Dim Display_Mode As D3DDISPLAYMODE
    Dim Direct3D_Window As D3DPRESENT_PARAMETERS
    
    Set DirectX8 = New DirectX8
    Set Direct3D = DirectX8.Direct3DCreate()
     Set Direct3DX = New D3DX8
        Direct3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display_Mode
        
        Direct3D_Window.Windowed = True
 
        Scalar.X = 1
        Scalar.Y = 1
 Direct3D_Window.SwapEffect = D3DSWAPEFFECT_FLIP
    Direct3D_Window.BackBufferFormat = Display_Mode.Format
    Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Gwin.hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, Direct3D_Window)
        Direct3D_Device.SetVertexShader FVF_TLVERTEX
        Direct3D_Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
    Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_SRCALPHA
    


    Direct3D_Device.SetRenderState D3DRS_ALPHAREF, 255
    Direct3D_Device.SetRenderState D3DRS_ALPHAFUNC, D3DCMP_GREATEREQUAL

    
    Direct3D_Device.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
    Direct3D_Device.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
End Function
'Loop Section Begin
Public Function Begin_Scene()
Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1, 0
Direct3D_Device.BeginScene
End Function
Public Function End_Scene()
Direct3D_Device.EndScene
Direct3D_Device.Present ByVal 0, ByVal 0, 0, ByVal 0
End Function
'Loop Section End


'Draw vertex
Public Function Draw_Vx4(X As Single, Y As Single, vx As Single, vy As Single, r As Single, b As Single, g As Single, Texture1 As Direct3DTexture8) As Boolean
Vx4Vertex_List(0) = Create_TLVertex(X, Y, 0, 1, D3DColorRGBA(r, g, b, 255), 0, 0, 0)
Vx4Vertex_List(1) = Create_TLVertex(X + (vx), Y, 0, 1, D3DColorRGBA(r, g, b, 255), 255, 1, 0)
Vx4Vertex_List(2) = Create_TLVertex(X, Y + (vy), 0, 1, D3DColorRGBA(r, g, b, 255), 255, 0, 1)
Vx4Vertex_List(3) = Create_TLVertex(X + (vx), Y + (vy), 0, 1, D3DColorRGBA(r, g, b, 255), 2, 1, 1)
 
Direct3D_Device.SetRenderState D3DRS_POINTSIZE, v
Direct3D_Device.SetRenderState D3DRS_POINTSIZE_MIN, 5
Direct3D_Device.SetRenderState D3DRS_POINTSCALE_A, X
Direct3D_Device.SetRenderState D3DRS_POINTSCALE_B, Y
Direct3D_Device.SetRenderState D3DRS_POINTSCALE_C, r
Direct3D_Device.SetTexture 0, Texture1
            
                
Direct3D_Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 3, Vx4Vertex_List(0), Len(Vx4Vertex_List(1))
Draw_Vx4 = True
End Function
Public Function Draw_Vx4Acolor(X As Single, Y As Single, vx As Single, vy As Single, vc1 As Long, vc2 As Long, vc3 As Long, vc4 As Long, Texture1 As Direct3DTexture8) As Boolean
Vertex_List(0) = Create_TLVertex(X - (vx / 2), Y - (vy / 2), 0, 1, vc1, 0, 0, 0)
Vertex_List(1) = Create_TLVertex(X + (vx / 2), Y - (vy / 2), 0, 1, vc2, 0, 1, 0)
Vertex_List(2) = Create_TLVertex(X - (vx / 2), Y + (vy / 2), 0, 1, vc3, 0, 0, 1)
Vertex_List(3) = Create_TLVertex(X + (vx / 2), Y + (vy / 2), 0, 1, vc4, 0, 1, 1)
 
 
Direct3D_Device.SetTexture 0, Texture1
            
                
Direct3D_Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 3, Vertex_List(0), Len(Vertex_List(1))
Draw_Vx4Acolor = True
End Function
'Draw Line
Public Function Draw_Line(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, r As Single, b As Single, g As Single) As Boolean
LineVertex_List(0) = Create_TLVertex(X1, Y1, 0, 1, D3DColorRGBA(r, g, b, 255), 0, 0, 0)
LineVertex_List(1) = Create_TLVertex(X2, Y2, 0, 1, D3DColorRGBA(r, g, b, 255), 0, 1, 0)
 
Direct3D_Device.DrawPrimitiveUP D3DPT_LINELIST, 2, LineVertex_List(0), Len(LineVertex_List(0))
Draw_Line = True
End Function
        



'Textures
Public Function Load_Texture(Width As Long, Height As Long, File_Path As String) As Direct3DTexture8
On Error Resume Next
Direct3D_Device.SetVertexShader FVF   'Set The Vertex Shader
Direct3D_Device.SetRenderState D3DRS_LIGHTING, False  'Disable Hardware Lighting (Using Vertex Lighting)
Direct3D_Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA  'Set The Source Blending Method
Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE  'Set The Destination Blending Method
Direct3D_Device.SetRenderState D3DRS_ALPHABLENDENABLE, True   'Enable Alphablending
Direct3D_Device.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1    'For The Particle Engine
Direct3D_Device.SetRenderState D3DRS_POINTSCALE_ENABLE, 0 'Also For The Particle Engine
Direct3D_Device.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE   'Allows ARGB To Work

Dim Transparency_Color As Long
    Transparency_Color = D3DColorRGBA(0, 0, 0, 255)

     Set Load_Texture = Direct3DX.CreateTextureFromFileEx(Direct3D_Device, File_Path, 35, 35, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, Transparency_Color, ByVal 0, ByVal 0)

                                                    
                                                        
End Function
   
        

