VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFrameMesh"
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

Private Type BoneComboIdList
    List(4) As Long
End Type

Public objectname As String

' Base Objects
Public Mesh As D3DXMesh                         'if not skinned, regular mesh object
Public SkinMesh As D3DXSkinMesh                 'if skinned - skinned mesh object

' Material and Adjacency information
Dim m_NumMaterials As Long                      'Number of materials in object
Dim m_Materials() As D3DMATERIAL8               'Material List
Dim m_Textures() As Direct3DTexture8            'Texture List

' Bone and skinning informatin
Dim m_BoneNames As D3DXBuffer                   'Names of Frame objects that are bones
Dim m_BoneFrames() As clsFrame           'Frame objects that are bones
Dim m_maxFaceInfl As Long                       'Number of matrices that will influence a vertex
Dim m_worldMatHandles(4) As Long                'handle to multiple world matrices
Dim m_BoneOffsetMat() As D3DMATRIX              'Bone offset matrices from D3DXBuffers
Dim m_NumBoneCombos As Long                     'Size of bonecombo table
Dim m_BoneCombos() As D3DXBONECOMBINATION       'Combo table as returned from D3DX
Dim m_BoneComboIds() As BoneComboIdList         'BoneId portion of combotable
Dim m_BoneCount As Long                         'Number of bones

'-----------------------------------------------------------------------------
' Name: InitFromXOF
' Desc: Load mesh from data provided by XOF api
'       Called from D3DUtil_LoadFromFile
'-----------------------------------------------------------------------------
Public Function InitFromXOF(meshdata As DirectXFileData)
    Dim MtrlBuffer As D3DXBuffer
    Dim m_AdjBuffer As D3DXBuffer                   'Adjacency buffer for the mesh
    Dim m_BoneMatrices As D3DXBuffer                'Matrix object per bone
    Dim m_BoneComboTable As D3DXBuffer              'Groupings of bone material and triangles
    Dim TextureFile As String
    Dim strPath As String
    Dim bonename As String
    Dim I As Long
    Dim q As Long
    
    objectname = meshdata.GetName()
 
        Set SkinMesh = D3DX.LoadSkinMeshFromXof(meshdata, D3DXMESH_MANAGED, D3DDevice, m_AdjBuffer, MtrlBuffer, m_NumMaterials, m_BoneNames, m_BoneMatrices)
        
        Dim pAdj As Long
        pAdj = m_AdjBuffer.GetBufferPointer
        
        m_BoneCount = SkinMesh.GetNumBones()

        Set Mesh = SkinMesh.ConvertToBlendedMesh(D3DXMESH_SYSTEMMEM, ByVal pAdj, ByVal 0, m_NumBoneCombos, m_BoneComboTable, ByVal 0&, Nothing)
            
            m_maxFaceInfl = SkinMesh.GetMaxFaceInfluences()
            
            m_worldMatHandles(0) = D3DTS_WORLD
            m_worldMatHandles(1) = D3DTS_WORLD1
            m_worldMatHandles(2) = D3DTS_WORLD2
            m_worldMatHandles(3) = D3DTS_WORLD3
            
            ReDim m_BoneCombos(m_NumBoneCombos)
            ReDim m_BoneComboIds(m_NumBoneCombos)
            
            For q = 0 To m_NumBoneCombos - 1
                D3DX.BufferGetBoneCombo m_BoneComboTable, q, m_BoneCombos(q)
                D3DX.BufferGetBoneComboBoneIds m_BoneComboTable, q, m_maxFaceInfl, m_BoneComboIds(q).List(0)
            Next
            
            ReDim m_BoneOffsetMat(m_BoneCount)
            D3DX.BufferGetData m_BoneMatrices, 0, Len(m_BoneOffsetMat(0)), m_BoneCount, m_BoneOffsetMat(0)
         
        ReDim m_Materials(m_NumMaterials)
        ReDim m_Textures(m_NumMaterials)

        strPath = App.Path + "\Text\"
        
        For I = 0 To m_NumMaterials - 1

        D3DX.BufferGetMaterial MtrlBuffer, I, m_Materials(I)
        
        m_Materials(I).Ambient = m_Materials(I).diffuse
        TextureFile = D3DX.BufferGetTextureName(MtrlBuffer, I)
        If TextureFile <> "" Then
            Set m_Textures(I) = D3DX.CreateTextureFromFileEx(D3DDevice, strPath + TextureFile, 128, 128, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
        End If
        Next
        
    Set m_AdjBuffer = Nothing
    Set m_BoneComboTable = Nothing
    Set m_BoneMatrices = Nothing
    Set SkinMesh = Nothing
End Function

'-----------------------------------------------------------------------------
' Name: AttatchBonesToMesh
' Desc: Called to attach bones to a skin.
'       The BoneNames table is used to search out bone frames
'       in the children of the given parent frame
'
'       This must be done for any skinning animation to work
'-----------------------------------------------------------------------------
Friend Sub AttatchBonesToMesh(parent As clsFrame)
      ' get links to all the frames (bones)
        Dim I As Long
        Dim bonename As String
        
        ReDim m_BoneFrames(m_BoneCount)
        
        For I = 0 To m_BoneCount - 1
            bonename = D3DX.BufferGetBoneName(m_BoneNames, I)
            Set m_BoneFrames(I) = parent.FindChildObject(bonename)
            If m_BoneFrames(I) Is Nothing Then
                Debug.Print "unable to find " + bonename
                Stop
            End If
        Next
        Set m_BoneNames = Nothing
End Sub

'-----------------------------------------------------------------------------
' Name: RenderSkin
' Desc: Render the Mesh as skin
' Note: The mesh must have been loaded as a skin and bones must have been attached
'-----------------------------------------------------------------------------
Sub RenderSkin()
    Dim ipAttr As Long  'bonecombo attribute
    Dim matId As Long   'matrix id
    
    Dim I As Long
    Dim mat2 As D3DMATRIX
    Dim mat1 As D3DMATRIX
    Dim mat0 As D3DMATRIX
    
    D3DDevice.SetRenderState D3DRS_VERTEXBLEND, m_maxFaceInfl - 1
        
    For ipAttr = 0 To m_NumBoneCombos - 1
        For I = 0 To m_maxFaceInfl - 1
                matId = m_BoneComboIds(ipAttr).List(I)
                
                'If we get a MatId of -1 then all the vertex weights are 0
                'and we dont need to set the transform for this bone
                If matId <> -1 Then
                    mat0 = m_BoneFrames(matId).GetUpdatedMatrix()
                    mat1 = m_BoneOffsetMat(matId)
                    D3DXMatrixMultiply mat2, mat1, mat0
                    D3DDevice.SetTransform m_worldMatHandles(I), mat2
                End If
        Next
        D3DDevice.SetTexture 0, m_Textures(m_BoneCombos(ipAttr).AttribId)
        D3DDevice.SetMaterial m_Materials(m_BoneCombos(ipAttr).AttribId)
        
        Mesh.DrawSubset ipAttr
        
    Next
    
    D3DDevice.SetRenderState D3DRS_VERTEXBLEND, 0
End Sub

 
