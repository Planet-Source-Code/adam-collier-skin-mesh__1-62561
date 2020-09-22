VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Meshes"
   ClientHeight    =   10980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10980
   ScaleWidth      =   15075
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Dim character1 As New clsFrame
Dim run1 As New clsAnimation
Dim idle1 As New clsAnimation

Dim character2 As New clsFrame
Dim run2 As New clsAnimation
Dim idle2 As New clsAnimation

Dim Floor As New clsMesh

Dim fire As New clsFire

Dim cameramode As Integer
Dim zoom As Single

Dim keys(7) As Boolean

Dim pos1 As D3DVECTOR
Dim pos2 As D3DVECTOR
Dim ang1 As Single
Dim ang2 As Single

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then keys(0) = True
If KeyCode = vbKeyDown Then keys(1) = True
If KeyCode = vbKeyLeft Then keys(2) = True
If KeyCode = vbKeyRight Then keys(3) = True
If KeyCode = vbKeyW Then keys(4) = True
If KeyCode = vbKeyS Then keys(5) = True
If KeyCode = vbKeyA Then keys(6) = True
If KeyCode = vbKeyD Then keys(7) = True
If KeyCode = vbKey1 Then cameramode = 1
If KeyCode = vbKey2 Then cameramode = 2
If KeyCode = vbKey3 Then cameramode = 3
If KeyCode = vbKey4 Then cameramode = 4
If KeyCode = vbKey5 Then cameramode = 5
If zoom < 40 Then If KeyCode = vbKeyPageUp Then zoom = zoom + 1
If zoom > 2 Then If KeyCode = vbKeyPageDown Then zoom = zoom - 1
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then keys(0) = False
If KeyCode = vbKeyDown Then keys(1) = False
If KeyCode = vbKeyLeft Then keys(2) = False
If KeyCode = vbKeyRight Then keys(3) = False
If KeyCode = vbKeyW Then keys(4) = False
If KeyCode = vbKeyS Then keys(5) = False
If KeyCode = vbKeyA Then keys(6) = False
If KeyCode = vbKeyD Then keys(7) = False
End Sub

Private Sub Form_Load()
zoom = 4
cameramode = 5
pos1.Z = 5
pos2.Z = -5

Me.Show
DoEvents

InitD3D Me.hWnd
InitLight

character1.InitFromFile "drake.x"
run1.InitFromFile "run2.x", character1
idle1.InitFromFile "idle.x", character1

character2.InitFromFile "chris.x"
run2.InitFromFile "run2.x", character2
idle2.InitFromFile "idle.x", character2

fire.InitParticles

Floor.InitMesh "scene2.x"

Do

If keys(0) = True Then
pos1.X = pos1.X + Cos(-ang1)
pos1.Z = pos1.Z + Sin(-ang1)
run1.SetTime MilkTime
Else
idle1.SetTime MilkTime
End If

If keys(1) = True Then
pos1.X = pos1.X - Cos(-ang1)
pos1.Z = pos1.Z - Sin(-ang1)
run1.SetTime MilkTime
Else

End If

If keys(2) = True Then ang1 = ang1 + 5 * sngRad
If keys(3) = True Then ang1 = ang1 - 5 * sngRad

If keys(4) = True Then
pos2.X = pos2.X + Cos(-ang2)
pos2.Z = pos2.Z + Sin(-ang2)
run2.SetTime MilkTime
Else
idle2.SetTime MilkTime
End If

If keys(5) = True Then
pos2.X = pos2.X - Cos(-ang2)
pos2.Z = pos2.Z - Sin(-ang2)
run2.SetTime MilkTime
Else

End If

If keys(6) = True Then ang2 = ang2 + 5 * sngRad
If keys(7) = True Then ang2 = ang2 - 5 * sngRad




character1.AddRotation COMBINE_replace, 0, 1, 0, ang1 - (sngPi / 2)
character1.AddTranslation COMBINE_after, pos1.X, pos1.Y, pos1.Z
character1.UpdateFrames

character2.AddRotation COMBINE_replace, 0, 1, 0, ang2 - (sngPi / 2)
character2.AddTranslation COMBINE_after, pos2.X, pos2.Y, pos2.Z
character2.UpdateFrames

fire.UpdateParticles
If cameramode = 1 Then SetCamera True, Vec3(pos1.X + Cos(-ang1 + sngPi) * 20, 10, pos1.Z + Sin(-ang1 + sngPi) * 20), Vec3(0, 0, 0), Vec3(pos1.X + 0.1, 10, pos1.Z), zoom
If cameramode = 2 Then SetCamera True, Vec3(pos2.X + Cos(-ang2 + sngPi) * 20, 10, pos2.Z + Sin(-ang2 + sngPi) * 20), Vec3(0, 0, 0), Vec3(pos2.X + 0.1, 10, pos2.Z), zoom
If cameramode = 3 Then SetCamera True, Vec3(0, 50, 0), Vec3(0, 0, 0), Vec3(pos1.X + 0.1, 7, pos1.Z), zoom
If cameramode = 4 Then SetCamera True, Vec3(0, 50, 0), Vec3(0, 0, 0), Vec3(pos2.X + 0.1, 7, pos2.Z), zoom
If cameramode = 5 Then SetCamera True, Vec3(100, 50, 0), Vec3(0, 0, 0), Vec3(0, 7, 0), zoom
     
DoEvents
Render
Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
DeInitD3D
Floor.DeInitMesh
End Sub

Sub Render()
D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, &H0, 1#, 0
D3DDevice.BeginScene

character1.RenderSkins
character2.RenderSkins
Floor.Render Vec3(0, 0, 0), Vec3(0, 0, 0), Vec3(1, 1, 1)
fire.Render Vec3(50, 0, 0), Vec3(0, 0, 0), Vec3(1, 1, 1)

D3DDevice.EndScene
D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
End Sub

