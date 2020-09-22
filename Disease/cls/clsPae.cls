VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParticle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type PARTICLEVERTEX
    v As D3DVECTOR
    Color As Long
End Type

Dim FVF_PARTICLEVERTEX
Dim ParticleTex As Direct3DTexture8
Dim particle() As PARTICLEVERTEX
Dim intCount As Integer
Dim matworld As D3DMATRIX
Dim v() As D3DVECTOR
Dim a() As Integer

Public Sub Init(Image As String, Amount As Integer, ParticleSize As Single)
intCount = Amount
Dim i As Integer
ReDim particle(Amount) As PARTICLEVERTEX
ReDim v(Amount) As D3DVECTOR
ReDim a(Amount) As Integer

FVF_PARTICLEVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE)

D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 1
D3DDevice.SetRenderState D3DRS_POINTSIZE, FtoDW(ParticleSize)

Set ParticleTex = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\text\" + Image, 32, 32, D3DX_DEFAULT, 0, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &H0, ByVal 0, ByVal 0)

End Sub

Public Sub drawit(Position As D3DVECTOR, Angle As D3DVECTOR, Scaling As D3DVECTOR)

Dim i As Integer
For i = 0 To intCount
particle(i).v.x = particle(i).v.x + v(i).x
particle(i).v.y = particle(i).v.y + v(i).y
particle(i).v.z = particle(i).v.z + v(i).z
particle(i).Color = D3DColorARGB(a(i), a(i), a(i), a(i))
a(i) = a(i) - 1
If a(i) < 0 Then
particle(i).v = Position
v(i).x = (Rnd * (-0.03 / 3)) + (0.1 / 3)
v(i).y = (Rnd * (-0.02 / 3)) + (0.1 / 3)
v(i).z = (Rnd * (-0.03 / 3)) + (0.015 / 3)
a(i) = Int(Rnd * 255)
End If
Next i

SetWorld Position, Angle, Scaling

D3DDevice.SetRenderState D3DRS_ZWRITEENABLE, 0
D3DDevice.SetRenderState D3DRS_LIGHTING, 0
D3DDevice.SetVertexShader FVF_PARTICLEVERTEX
D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
D3DDevice.SetTexture 0, ParticleTex
D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, intCount, particle(0), Len(particle(0))
D3DDevice.SetRenderState D3DRS_ZWRITEENABLE, 1
D3DDevice.SetRenderState D3DRS_LIGHTING, 1
D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0

End Sub

Public Function FtoDW(f As Single) As Long
    Dim buf As D3DXBuffer
    Dim l As Long
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, f
    D3DX.BufferGetData buf, 0, 4, 1, l
    FtoDW = l
End Function


