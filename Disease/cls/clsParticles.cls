VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'           SSSSS AAAAA DDDD  AAAAA M   M H   H U   U
'           S   S A   A D   D A   A MM MM H   H U   U
'           S     A   A D   D A   A M M M H   H U   U
'           SSSSS AAAAA D   D AAAAA M   M HHHHH U   U
'               S A   A D   D A   A M   M H   H U   U
'           S   S A   A D   D A   A M   M H   H U   U
'           SSSSS A   A DDDD  A   A M   M H   H  UUU

'              !!THANK YOU FOR DOWNLOADING THIS!!

'      Please learn from this code, everyone else tought me it.

'                Some of it is M$, dont tell Bill!!

'         Print out the skinmesh classes and read them on the floor
'       (thats what I did(after making it more simple from the SDK))



'        ..jj
'        iiWW..            iiLL
'        tt##;;            ttDD
'        ffWWtt            ttEE
'        LLDDLL            iiKK
'        GGffEE            iiKK                          iiKKGG..
'        DD;;KK;;          ;;WW                  ..;;tt..EEttLLGG
'        EE,,LLLLDDtt      ,,WW        LLKKii    iiWWKKKKKK    DDii
'        KK..ff##tt      ..ttWW..    ttDDiiWW..  ..##ttGGKK    ttLL
'        KKiiKKKK..    ;;KKEE##..    LLtt  KKii    ##ii;;ii    ;;ii
'      ..KKKKttGGii    GGtt..##ii    ffff  DDDD....KK..
'      tt##tt  ffLL  ..KK....WWLL    ;;KKiiKKGGKKKK;;
'    ttKKKK    iiDD  ..KK..iiEEEEffjj..ttEELL  ..,,                                      iiGGffLLtt
'  ttEE;;KK..    KK..  GGffLLff;;LLff                ..;;  tt;;                      LL;;KKjjfftt..
'  LL;;  KK..    LLii  ..EEKK,,                      iiDD  ttGG    ..      ..fftt    EEDDff
'      ..KK      iiGG      ..                        ..KK  ;;EE  ;;KK    ;;KKffKK..  LL##ii
'      ;;EE        ..                                  KK..  KK::  ;;    KKttLLGG  iijj##tt
'      ;;DD                      ..                    DD;;  LLtt  ;;  ;;WWDDLL..;;KK;;KKff
'      ..tt                  ,,GGKK;;        ..jjDDii  LLtt  jjLL..KK  ..WWjj  iiEEtt  LLGG
'                          ..EEff..          ff##jjKK..jjLL  iiGG  DD..  jjKKEEEEtt    ;;ff
'                          ffLL              GGff  LLtt;;DD  ;;EE  LLtt    ..;;
'                          DDii        ::;;  GGtt  GGtt  KK;;  KK;;tttt
'                        ..KK..        iiGG  iiKKKKGG..  LLLL  GGtt
'                          KK::        GGff      ..      iitt  ;;::
'                          ffKKttiiiiGGDD..
'                            iiGGDDGGtt
'                                                                               Adam Collier
Option Explicit

Const nParticles As Long = 100  '//How many particles we are going to use
Const ParticleSize As Single = 1
Const Gravity As Single = -0.05
Const XWind As Single = 0
Const ZWind As Single = 0

Dim VBuffer As Direct3DVertexBuffer8

Private Type PARTICLEVERTEX
    P As D3DVECTOR
    color As Long
End Type

Const FVF_PARTICLEVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE)

Private Type particle
    V As D3DVECTOR
    startcolor As D3DCOLORVALUE
    endcolor As D3DCOLORVALUE
    currentcolor As D3DCOLORVALUE
    LifeTime As Long
    Created As Long
    Status As Boolean
End Type

Dim PrtData(0 To nParticles - 1) As particle
Dim PrtVertList(0 To nParticles - 1) As PARTICLEVERTEX

Dim PART0000 As Direct3DTexture8

Dim LastUpdatedParticles As Single

Public Sub InitParticles()
Dim I As Integer

D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 1
D3DDevice.SetRenderState D3DRS_POINTSIZE, FtoDW(ParticleSize)
D3DDevice.SetRenderState D3DRS_POINTSIZE_MIN, FtoDW(ParticleSize)
D3DDevice.SetRenderState D3DRS_POINTSCALE_C, FtoDW(ParticleSize)
Set PART0000 = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\text\part0000.bmp", 32, 32, D3DX_DEFAULT, 0, D3DFMT_A1R5G5B5, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFFFF00FF, ByVal 0, ByVal 0)

For I = 0 To nParticles - 1
    PrtData(I).Status = False
Next I

End Sub

Public Sub UpdateParticles()

    Dim I As Long
    Dim ExactTime As Single
    ExactTime = CurrentTime

    For I = 0 To nParticles - 1
        If PrtData(I).Status = True Then
                '//Update the positions
                    PrtVertList(I).P.X = PrtVertList(I).P.X + ((PrtData(I).V.X / 500) * (ExactTime - LastUpdatedParticles))
                    PrtVertList(I).P.Y = PrtVertList(I).P.Y + ((PrtData(I).V.Y / 500) * (ExactTime - LastUpdatedParticles))
                    PrtVertList(I).P.Z = PrtVertList(I).P.Z + ((PrtData(I).V.Z / 500) * (ExactTime - LastUpdatedParticles))
                    
                '//Update the velocities
                    PrtData(I).V.X = PrtData(I).V.X + ((XWind / 500) * (ExactTime - LastUpdatedParticles))
                    PrtData(I).V.Y = PrtData(I).V.Y + ((Gravity / 500) * (ExactTime - LastUpdatedParticles))
                    PrtData(I).V.Z = PrtData(I).V.Z + ((ZWind / 500) * (ExactTime - LastUpdatedParticles))
                    
                '//Update The color values
                    D3DXColorLerp PrtData(I).currentcolor, PrtData(I).startcolor, PrtData(I).endcolor, (ExactTime - PrtData(I).Created) / PrtData(I).LifeTime
                    
                '//Check if the particle has gone below ground level...
                    If PrtVertList(I).P.Y < -1 Then PrtData(I).Status = False
                    
                '//Check if it's lifetime has expired
                    If ExactTime - PrtData(I).Created >= PrtData(I).LifeTime Then PrtData(I).Status = False
        Else
                PrtData(I).Status = True
                PrtData(I).LifeTime = Int(Rnd * 1000) + 1
                PrtData(I).Created = ExactTime
                PrtVertList(I).P.X = 0 + (Rnd * -0.25) + 0.125
                PrtVertList(I).P.Y = 0 + (Rnd * -0.25) + 0.125
                PrtVertList(I).P.Z = 0 + (Rnd * -0.25) + 0.125
                PrtData(I).V.X = (Rnd * -0.5) + 0.35
                PrtData(I).V.Y = (Rnd * 3) + 1
                PrtData(I).V.Z = (Rnd * -0.5) + 0.25
                PrtData(I).startcolor = ColorValue4(1, 1, 0.75, 0.5)
                PrtData(I).endcolor = ColorValue4(0.08, 0.08, 0.08, 0.08)
                PrtData(I).currentcolor = PrtData(I).startcolor
        End If
    Next I

LastUpdatedParticles = ExactTime

For I = 0 To nParticles - 1
    PrtVertList(I).color = D3DColorARGB(PrtData(I).currentcolor.a * 255, PrtData(I).currentcolor.R * 255, PrtData(I).currentcolor.g * 255, PrtData(I).currentcolor.B * 255)
Next I

End Sub

Public Sub Render(Position As D3DVECTOR, Angle As D3DVECTOR, Scaling As D3DVECTOR) ' put this in the d3ddevice.begin, d3ddevice end
SetWorld Position, Angle, Scaling

    D3DDevice.SetRenderState D3DRS_LIGHTING, 0
    D3DDevice.SetRenderState D3DRS_ZWRITEENABLE, 0
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1

    D3DDevice.SetVertexShader FVF_PARTICLEVERTEX
    D3DDevice.SetTexture 0, PART0000
    D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, nParticles, PrtVertList(0), Len(PrtVertList(0))
    
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
    D3DDevice.SetRenderState D3DRS_ZWRITEENABLE, 1
    D3DDevice.SetRenderState D3DRS_LIGHTING, 0
End Sub


