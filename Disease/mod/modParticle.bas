Attribute VB_Name = "modParticle"
Option Explicit

Public Type PARTICLEVERTEX
    v As D3DVECTOR
    Color As Long
End Type

Public ParticleSize As Single
Public FVF_PARTICLEVERTEX
Public ParticleTex As Direct3DTexture8
Public particle() As PARTICLEVERTEX
Const amount = 250
Dim matworld As D3DMATRIX
Dim v() As D3DVECTOR
Dim a() As Integer

Public Sub LoadParticle()
Dim i As Integer
ReDim particle(amount) As PARTICLEVERTEX
ReDim v(amount) As D3DVECTOR
ReDim a(amount) As Integer
ParticleSize = 0.03
FVF_PARTICLEVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE)

D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 1
D3DDevice.SetRenderState D3DRS_POINTSIZE, FtoDW(ParticleSize)

Set ParticleTex = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\text\particle.bmp", 32, 32, D3DX_DEFAULT, 0, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &H0, ByVal 0, ByVal 0)

End Sub

Public Sub drawit()
D3DDevice.SetRenderState D3DRS_ZWRITEENABLE, 0
D3DDevice.SetRenderState D3DRS_LIGHTING, 0
Dim i As Integer
For i = 0 To amount
particle(i).v.x = particle(i).v.x + v(i).x
particle(i).v.y = particle(i).v.y + v(i).y
particle(i).v.z = particle(i).v.z + v(i).z
particle(i).Color = D3DColorARGB(a(i), a(i), a(i), a(i))
a(i) = a(i) - 1
If a(i) < 0 Then
particle(i).v.x = 0
particle(i).v.y = 0
particle(i).v.z = 0
v(i).x = (Rnd * (-0.03 / 3)) + (0.1 / 3)
v(i).y = (Rnd * (-0.02 / 3)) + (0.1 / 3)
v(i).z = (Rnd * (-0.03 / 3)) + (0.015 / 3)
a(i) = Int(Rnd * 255)
End If
Next i

D3DDevice.SetVertexShader FVF_PARTICLEVERTEX
D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1

'D3DXMatrixIdentity matworld
'D3DDevice.SetTransform D3DTS_WORLD, matworld

D3DDevice.SetTexture 0, ParticleTex
D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, amount, particle(0), Len(particle(0))
D3DDevice.SetRenderState D3DRS_ZWRITEENABLE, 1
D3DDevice.SetRenderState D3DRS_LIGHTING, 1
End Sub

Public Function FtoDW(f As Single) As Long
    Dim buf As D3DXBuffer
    Dim l As Long
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, f
    D3DX.BufferGetData buf, 0, 4, 1, l
    FtoDW = l
End Function
