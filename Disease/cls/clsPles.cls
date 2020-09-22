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

Const nParticles As Long = 1000 '//How many particles we are going to use
Const ParticleSize As Single = 0.01
Const Gravity As Single = -9.8
Const XWind As Single = 0
Const ZWind As Single = 0
Const variation As Single = 0.5
Const mass As Single = 0
Dim VBuffer As Direct3DVertexBuffer8

Private Type PARTICLEVERTEX
    v As D3DVECTOR
    Color As Long
    tu As Single
    tv As Single
End Type

Const FVF_PARTICLEVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)

Private Enum PARTICLE_STATUS
    Alive = 0
    Dead = 1
End Enum


Private Type particle
    X As Single     'World Space Coordinates
    Y As Single
    Z As Single
    vX As Single    'Speed and Direction
    vY As Single
    vZ As Single
    StartColor As D3DCOLORVALUE
    EndColor As D3DCOLORVALUE
    CurrentColor As D3DCOLORVALUE
    LifeTime As Long    'How long Mr. Particle Exists
    Created As Long 'When this particle was created...
    Status As PARTICLE_STATUS 'Does he even exist?
End Type

Dim PrtData(0 To nParticles - 1) As particle
Dim PrtVertList(0 To nParticles - 1) As PARTICLEVERTEX

Dim ParticleTex As Direct3DTexture8

Dim LastUpdatedParticles As Single

Public Sub Render()

    D3DDevice.SetRenderState D3DRS_LIGHTING, 0
    D3DDevice.SetRenderState D3DRS_ZWRITEENABLE, 0
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1

    D3DDevice.SetVertexShader FVF_PARTICLEVERTEX
    D3DDevice.SetTexture 0, ParticleTex
    D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, nParticles, PrtVertList(0), Len(PrtVertList(0))
    
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
    D3DDevice.SetRenderState D3DRS_ZWRITEENABLE, 1
    D3DDevice.SetRenderState D3DRS_LIGHTING, 0
    
End Sub

Public Sub UpdateParticles()

    Dim I As Long
    Dim ExactTime As Single
    ExactTime = CurrentTime / 5000

    For I = 0 To nParticles - 1
        If PrtData(I).Status = Alive Then
                
                '//Update the velocities
                    PrtData(I).vX = PrtData(I).vX + ((XWind) * (ExactTime - LastUpdatedParticles))
                    PrtData(I).vY = PrtData(I).vY + ((Gravity + mass) * (ExactTime - LastUpdatedParticles))
                    PrtData(I).vZ = PrtData(I).vZ + ((ZWind) * (ExactTime - LastUpdatedParticles))
                '//Update the positions
                    PrtData(I).X = PrtData(I).X + ((PrtData(I).vX) * (ExactTime - LastUpdatedParticles))
                    PrtData(I).Y = PrtData(I).Y + ((PrtData(I).vY) * (ExactTime - LastUpdatedParticles))
                    PrtData(I).Z = PrtData(I).Z + ((PrtData(I).vZ) * (ExactTime - LastUpdatedParticles))
       
                '//Update The color values
                    D3DXColorLerp PrtData(I).CurrentColor, PrtData(I).StartColor, PrtData(I).EndColor, (ExactTime - PrtData(I).Created) / PrtData(I).LifeTime
                    
                '//Check if the particle has gone below ground level...
                    If PrtData(I).Y < -1 Then PrtData(I).Status = Dead
                    
                '//Check if it's lifetime has expired
                    If ExactTime - PrtData(I).Created >= PrtData(I).LifeTime Then PrtData(I).Status = Dead
        Else
                '//We need to recreate our particle...
                PrtData(I).Status = Alive
                PrtData(I).LifeTime = 25 + ((Rnd * 12.5) - 5)
                PrtData(I).Created = ExactTime
                PrtData(I).X = 0
                PrtData(I).Y = 2
                PrtData(I).Z = 0
                PrtData(I).vX = (Rnd * -1) + 0.5
                PrtData(I).vY = (Rnd * 5) + 3
                PrtData(I).vZ = (Rnd * -1) + 0.5
                PrtData(I).StartColor = ColorValue4(0.5, 0, 0, 1)
                PrtData(I).EndColor = ColorValue4(0.5, 0, 0, 1)
                PrtData(I).CurrentColor = PrtData(I).StartColor
        End If
    Next I

LastUpdatedParticles = ExactTime

For I = 0 To nParticles - 1
    PrtVertList(I).v = Vec3(PrtData(I).X, PrtData(I).Y, PrtData(I).Z)
    PrtVertList(I).Color = D3DColorARGB(PrtData(I).CurrentColor.a * 255, PrtData(I).CurrentColor.R * 255, PrtData(I).CurrentColor.G * 255, PrtData(I).CurrentColor.B * 255)
Next I

End Sub

Public Sub InitParticles()
Dim I As Integer

D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1 '//Enable point sprite rendering
D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 1 '//Allow Direct3D to set/alter the size of the Psprites
D3DDevice.SetRenderState D3DRS_POINTSIZE, FtoDW(ParticleSize)

Set ParticleTex = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\text\part0001.bmp", 32, 32, D3DX_DEFAULT, 0, D3DFMT_A1R5G5B5, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFFFF00FF, ByVal 0, ByVal 0)

For I = 0 To nParticles - 1
    PrtData(I).Status = Dead
Next I

End Sub

Private Function FtoDW(f As Single) As Long
    Dim buf As D3DXBuffer
    Dim l As Long
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, f
    D3DX.BufferGetData buf, 0, 4, 1, l
    FtoDW = l
End Function
