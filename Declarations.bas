Attribute VB_Name = "Declarations"
' Type Declaration
Public Type PPwin
hwnd As Long
scaleheight As Long
scalewidth As Long
End Type

Public Type TLVERTEX
    X As Single
    Y As Single
    z As Single
    RHW As Single
    Color As Long
    Specular As Long
    TU As Single
    tv As Single
End Type

'Custom Type variable
Public Gwin As PPwin

'Constants
Public Const COLOR_DEPTH_16_BIT As Long = D3DFMT_R5G6B5
Public Const COLOR_DEPTH_24_BIT As Long = D3DFMT_A8R8G8B8
Public Const COLOR_DEPTH_32_BIT As Long = D3DFMT_X8R8G8B8
Public Vertex_List_1(2) As TLVERTEX
Public Const FVF_TLVERTEX As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
Public Declare Function GetTickCount Lib "kernel32" () As Long













Private bufVertex As Direct3DVertexBuffer8 '=> VertexBuffer
Const fvfUDVertex = (D3DFVF_XYZRHW Or D3DFVF_DIFFUSE)
