Attribute VB_Name = "modDirectX"

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

Public DX As DirectX8
Public D3D As Direct3D8
Public D3DX As New D3DX8
Public D3DDevice As Direct3DDevice8      ' Our rendering device

Public sngPi As Single
Public sngRad As Single
Dim sngStartTime As Single

Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub InitD3D(hWnd As Long)
Dim MODE As D3DDISPLAYMODE
Dim D3DWINDOW As D3DPRESENT_PARAMETERS

Set DX = New DirectX8
Set D3D = DX.Direct3DCreate()

D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, MODE

D3DWINDOW.Windowed = 1
D3DWINDOW.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
D3DWINDOW.BackBufferFormat = MODE.Format
D3DWINDOW.BackBufferCount = 1
D3DWINDOW.EnableAutoDepthStencil = 1
D3DWINDOW.AutoDepthStencilFormat = D3DFMT_D16

Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWINDOW)

D3DDevice.SetRenderState D3DRS_ZENABLE, 1
D3DDevice.SetRenderState D3DRS_LIGHTING, 1
D3DDevice.SetRenderState D3DRS_CULLMODE, 1
D3DDevice.SetRenderState D3DRS_AMBIENT, RGB(100, 100, 100)
D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_CURRENT
D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_TEXTURE

D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR

sngStartTime = GetTickCount
sngPi = 3.1415
sngRad = sngPi * 180
End Sub

Public Sub DeInitD3D()
Set DX = Nothing
Set D3D = Nothing
Set D3DDevice = Nothing
End
End Sub

Public Function ColorValue4(a As Single, R As Single, g As Single, B As Single) As D3DCOLORVALUE
ColorValue4.a = a
ColorValue4.R = R
ColorValue4.g = g
ColorValue4.B = B
End Function

Public Function Vec3(X As Single, Y As Single, Z As Single) As D3DVECTOR
Vec3.X = X
Vec3.Y = Y
Vec3.Z = Z
End Function

Public Function MatSubtract(mat2 As D3DMATRIX, mat1 As D3DMATRIX) As D3DMATRIX
MatSubtract.m11 = mat2.m11 - mat1.m11
MatSubtract.m12 = mat2.m12 - mat1.m12
MatSubtract.m13 = mat2.m13 - mat1.m13
MatSubtract.m14 = mat2.m14 - mat1.m14
MatSubtract.m21 = mat2.m21 - mat1.m21
MatSubtract.m22 = mat2.m22 - mat1.m22
MatSubtract.m23 = mat2.m23 - mat1.m23
MatSubtract.m24 = mat2.m24 - mat1.m24
MatSubtract.m31 = mat2.m31 - mat1.m31
MatSubtract.m32 = mat2.m32 - mat1.m32
MatSubtract.m33 = mat2.m33 - mat1.m33
MatSubtract.m34 = mat2.m34 - mat1.m34
MatSubtract.m41 = mat2.m41 - mat1.m41
MatSubtract.m42 = mat2.m42 - mat1.m42
MatSubtract.m43 = mat2.m43 - mat1.m43
MatSubtract.m44 = mat2.m44 - mat1.m44
End Function

Public Function MatScalar(mat1 As D3DMATRIX, scalar As Single) As D3DMATRIX
MatScalar.m11 = mat1.m11 * scalar
MatScalar.m12 = mat1.m12 * scalar
MatScalar.m13 = mat1.m13 * scalar
MatScalar.m14 = mat1.m14 * scalar
MatScalar.m21 = mat1.m21 * scalar
MatScalar.m22 = mat1.m22 * scalar
MatScalar.m23 = mat1.m23 * scalar
MatScalar.m24 = mat1.m24 * scalar
MatScalar.m31 = mat1.m31 * scalar
MatScalar.m32 = mat1.m32 * scalar
MatScalar.m33 = mat1.m33 * scalar
MatScalar.m34 = mat1.m34 * scalar
MatScalar.m41 = mat1.m41 * scalar
MatScalar.m42 = mat1.m42 * scalar
MatScalar.m43 = mat1.m43 * scalar
MatScalar.m44 = mat1.m44 * scalar
End Function

Public Function MatAdd(mat1 As D3DMATRIX, mat2 As D3DMATRIX) As D3DMATRIX
MatAdd.m11 = mat2.m11 + mat1.m11
MatAdd.m12 = mat2.m12 + mat1.m12
MatAdd.m13 = mat2.m13 + mat1.m13
MatAdd.m14 = mat2.m14 + mat1.m14
MatAdd.m21 = mat2.m21 + mat1.m21
MatAdd.m22 = mat2.m22 + mat1.m22
MatAdd.m23 = mat2.m23 + mat1.m23
MatAdd.m24 = mat2.m24 + mat1.m24
MatAdd.m31 = mat2.m31 + mat1.m31
MatAdd.m32 = mat2.m32 + mat1.m32
MatAdd.m33 = mat2.m33 + mat1.m33
MatAdd.m34 = mat2.m34 + mat1.m34
MatAdd.m41 = mat2.m41 + mat1.m41
MatAdd.m42 = mat2.m42 + mat1.m42
MatAdd.m43 = mat2.m43 + mat1.m43
MatAdd.m44 = mat2.m44 + mat1.m44
End Function

Public Function CurrentTime() As Single
CurrentTime = GetTickCount
End Function

Public Function MilkTime() As Single
MilkTime = CurrentTime / 30
End Function

Public Function FtoDW(f As Single) As Long
    Dim buf As D3DXBuffer
    Dim l As Long
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, f
    D3DX.BufferGetData buf, 0, 4, 1, l
    FtoDW = l
End Function
