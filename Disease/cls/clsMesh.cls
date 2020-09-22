VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMesh"
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

'when you want to use a static mesh in your game type
'dim mesh as new clsmesh
'dont forget to deinit it

Public Mesh As D3DXMesh
Private MeshMaterials() As D3DMATERIAL8
Private MeshTextures() As Direct3DTexture8
Private NMaterials As Long

Public Sub InitMesh(filename As String) ' put this line in form load
    Dim MtrlBuffer As D3DXBuffer
    Dim I As Long
    Dim TextureFile As String
    
    Set Mesh = D3DX.LoadMeshFromX(App.Path & "\mesh\" + filename, D3DXMESH_MANAGED, D3DDevice, Nothing, MtrlBuffer, NMaterials)

    ReDim MeshMaterials(NMaterials) As D3DMATERIAL8
    ReDim MeshTextures(NMaterials) As Direct3DTexture8

    For I = 0 To NMaterials - 1
        D3DX.BufferGetMaterial MtrlBuffer, I, MeshMaterials(I)
        MeshMaterials(I).Ambient = MeshMaterials(I).diffuse
        TextureFile = D3DX.BufferGetTextureName(MtrlBuffer, I)
        If TextureFile <> "" Then
            Set MeshTextures(I) = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\text\" & TextureFile, 128, 128, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
        End If
        Next I
End Sub

Public Sub Render(Position As D3DVECTOR, Angle As D3DVECTOR, Scaling As D3DVECTOR) ' put this in the d3ddevice.begin, d3ddevice end
SetWorld Position, Angle, Scaling

Dim I As Long
    For I = 0 To NMaterials - 1
        D3DDevice.SetMaterial MeshMaterials(I)
        D3DDevice.SetTexture 0, MeshTextures(I)
        Mesh.DrawSubset I
    Next I
End Sub

Public Sub DeInitMesh() ' place this line in the form unload
Erase MeshTextures
Erase MeshMaterials
Set Mesh = Nothing
End Sub
