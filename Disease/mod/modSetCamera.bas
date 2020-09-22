Attribute VB_Name = "modSetCamera"

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

Public Sub SetCamera(AutoLook As Boolean, Position As D3DVECTOR, Angle As D3DVECTOR, LookAt As D3DVECTOR, zoom As Single)
Dim matProj As D3DMATRIX
Dim matView As D3DMATRIX
Dim matTemp As D3DMATRIX

If AutoLook = True Then
D3DXMatrixLookAtLH matView, Position, LookAt, Vec3(0, 1, 0)
Else

With Angle
D3DXMatrixIdentity matTemp
D3DXMatrixRotationY matTemp, .X * sngRad
D3DXMatrixMultiply matView, matView, matTemp

D3DXMatrixIdentity matTemp
D3DXMatrixRotationX matTemp, .Y * sngRad
D3DXMatrixMultiply matView, matView, matTemp
End With

With Position
D3DXMatrixIdentity matTemp
D3DXMatrixTranslation matTemp, .X, .Y, .Z
D3DXMatrixMultiply matView, matView, matTemp
End With
End If

D3DXMatrixIdentity matTemp
D3DXMatrixRotationZ matTemp, Angle.Z * sngRad
D3DXMatrixMultiply matView, matView, matTemp

D3DDevice.SetTransform D3DTS_VIEW, matView

D3DXMatrixPerspectiveFovLH matProj, sngPi / zoom, 1, 1, 100000
D3DDevice.SetTransform D3DTS_PROJECTION, matProj
End Sub
