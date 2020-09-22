VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFrame"
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

Public objectname As String             'Name of object

Dim m_Matrix As D3DMATRIX               'local coordinate system matrix
Dim m_MatConcat As D3DMATRIX            'world matrix coordinate system
Dim m_MatConcatInv As D3DMATRIX         'inv world matrix

Dim m_D3DMesh() As clsFrameMesh         'List of child Meshes of this frame
Dim m_NumMesh As Long                   'Number of child meshes that this frame contains

Dim m_D3DFrame() As clsFrame            'List of Child Frames of this frame
Dim m_NumFrames As Long                 'Number of child frames that this frame contains

Dim m_pos As D3DVECTOR                  'Position as set by SetPosition
Dim m_quat As D3DQUATERNION             'Orientations as set by SetOrientation
Dim m_rotcenter As D3DVECTOR            '
Dim m_scale As Single                   'Uniform scale as set by SetScale
Dim g_identityMatrix As D3DMATRIX  ' TODO what is this?

Enum CombineTypeEnum
    COMBINE_before = 0
    COMBINE_after = 1 ' this is the best setting i feel (but i may change my thoughts)
    COMBINE_replace = 2
End Enum

'-----------------------------------------------------------------------------
' Name: AddChild
' Desc: adds another frame as a child
'-----------------------------------------------------------------------------
Public Sub AddChild(child As clsFrame)
    ReDim Preserve m_D3DFrame(m_NumFrames)
    Set m_D3DFrame(m_NumFrames) = child
    m_NumFrames = m_NumFrames + 1
End Sub

'-----------------------------------------------------------------------------
' Name: AddMesh
' Desc: Adds a Child Mesh object to the frame
'-----------------------------------------------------------------------------
Public Sub AddMesh(childMesh As clsFrameMesh)
    ReDim Preserve m_D3DMesh(m_NumMesh)
    Set m_D3DMesh(m_NumMesh) = childMesh
    m_NumMesh = m_NumMesh + 1
End Sub

'-----------------------------------------------------------------------------
' Name: FindChildObject
' Desc: Given an object name (flags are ignored) return the first child with
'       a matching name. use the ClassName property to see if its a frame object
'-----------------------------------------------------------------------------
Public Function FindChildObject(Name As String) As Object

    Dim I As Long
    Dim aMesh As clsFrameMesh
    Dim aFrame As clsFrame
    
    'check self
    If objectname = Name Then
        Set FindChildObject = Me
        Exit Function
    End If

    
    'check child mesh
    For I = 0 To m_NumMesh - 1
        Set aMesh = m_D3DMesh(I)
        If Not aMesh Is Nothing Then
            If aMesh.objectname = Name Then
                Set FindChildObject = aMesh
                Exit Function
            End If
        End If
    Next
    
    Dim o As Object
    'check child frames
    For I = 0 To m_NumFrames - 1
        Set aFrame = m_D3DFrame(I)
        If Not aFrame Is Nothing Then
            Set o = aFrame.FindChildObject(Name)
            If Not o Is Nothing Then
                Set FindChildObject = o
                Exit Function
            End If
        End If
    Next
    Set FindChildObject = Nothing
    
End Function

'-----------------------------------------------------------------------------
' Name: InitFromFile
' Desc: character called this
'-----------------------------------------------------------------------------
Public Function InitFromFile(strFile As String) 'i am the ROOT
    Dim ChildData As DirectXFileData
    Dim xfile As DirectXFile
    Dim enumX As DirectXFileEnum
    Dim strPath As String
    
    strPath = App.Path + "\Mesh\"
    
    Set xfile = DX.DirectXFileCreate() 'create the X file object
    
    xfile.RegisterDefaultTemplates 'make sure it understand MeshBuilders et
    
    Set enumX = xfile.CreateEnumObject(strPath + strFile) 'Open the file
    
    Do 'Inspect each top level object in the file
        Set ChildData = enumX.GetNextDataObject()
        If ChildData Is Nothing Then Exit Do
                
        If ChildData.GetType = "TID_D3DRMFrame" Then
            Dim newFrame As clsFrame
            Set newFrame = New clsFrame
            newFrame.InitFromXOF ChildData, Me
            Set newFrame = Nothing
            
        ElseIf ChildData.GetType = "TID_D3DRMMesh" Then
            Dim newmesh As clsFrameMesh
            Set newmesh = New clsFrameMesh
            newmesh.InitFromXOF ChildData
            AddMesh newmesh
            Set newmesh = Nothing
            
        ElseIf ChildData.GetType = "TID_D3DRMMaterialSet" Then 'TODO Parse TopLevel Material Data
                        
        End If
    Loop
    
Me.AttatchBonesToMesh Me
End Function

'-----------------------------------------------------------------------------
' Name: InitFromXOF
' Desc: called from InitFromFile
'-----------------------------------------------------------------------------
Friend Sub InitFromXOF(FileData As DirectXFileData, parent As clsFrame)
    Dim ChildData As DirectXFileData
    Dim newmesh As clsFrameMesh
    Dim newFrame As clsFrame

    If Not parent Is Nothing Then parent.AddChild Me
    
    objectname = FileData.GetName() 'this frame holds the bone name
    
    Debug.Print objectname

    Do
        'list of object - should return nothing at end of list
        Set ChildData = FileData.GetNextObject()
        If ChildData Is Nothing Then Exit Do
    
        Select Case ChildData.GetType()
               
            Case "TID_D3DRMFrame"
                Set newFrame = New clsFrame
                newFrame.InitFromXOF ChildData, Me
                Set newFrame = Nothing
  
            Case "TID_D3DRMMesh"
                Set newmesh = New clsFrameMesh
                newmesh.InitFromXOF ChildData
                AddMesh newmesh
                Set newmesh = Nothing
  
            Case "TID_D3DRMFrameTransformMatrix"
                ChildData.GetDataFromOffset "", 0, Len(m_Matrix), m_Matrix
            
        End Select
            
    Loop
End Sub

'-----------------------------------------------------------------------------
' Name: UpdateFrames
' Desc: Called to precompute the WorldMatrices for each child object
' Note: Must be used prior to calling Mesh.RenderSkin
'-----------------------------------------------------------------------------
Public Sub UpdateFrames()
ComputeChildMatricesEx g_identityMatrix
End Sub

'-----------------------------------------------------------------------------
' Name: ComputeChildMatricesEx
' Desc: Aux function for UpdateFrames
'-----------------------------------------------------------------------------
Friend Sub ComputeChildMatricesEx(matSavedWorld As D3DMATRIX)
    Dim I As Long

    D3DXMatrixMultiply m_MatConcat, m_Matrix, matSavedWorld
    
    For I = 0 To m_NumFrames - 1
        If Not m_D3DFrame(I) Is Nothing Then
            m_D3DFrame(I).ComputeChildMatricesEx m_MatConcat
        End If
    Next
End Sub

'-----------------------------------------------------------------------------
' Name: RenderSkins
' Desc: Searches the children for any mesh with skinned objects
'       and renders them
' Note: BoneMatrice must be precalulated before rendering
'       and bones must also be attached to appropriate mesh
'-----------------------------------------------------------------------------
Friend Sub RenderSkins()
    Dim I As Long
   
    For I = 0 To m_NumMesh - 1
        If Not m_D3DMesh(I) Is Nothing Then
            m_D3DMesh(I).RenderSkin
        End If
    Next
    
    For I = 0 To m_NumFrames - 1
        If Not m_D3DFrame(I) Is Nothing Then
            m_D3DFrame(I).RenderSkins
        End If
    Next
End Sub

'-----------------------------------------------------------------------------
' Name: AttatchBonesToMesh
' Desc:
'-----------------------------------------------------------------------------
Friend Sub AttatchBonesToMesh(parent As clsFrame)
    Dim I As Long

    For I = 0 To m_NumMesh - 1
        If Not m_D3DMesh(I) Is Nothing Then
            m_D3DMesh(I).AttatchBonesToMesh parent
        End If
    Next
    
    For I = 0 To m_NumFrames - 1
        If Not m_D3DFrame(I) Is Nothing Then
            m_D3DFrame(I).AttatchBonesToMesh parent
        End If
    Next
End Sub

'-----------------------------------------------------------------------------
' Name: Class_Initialize()
'-----------------------------------------------------------------------------
Private Sub Class_Initialize()
    D3DXMatrixIdentity m_Matrix
    D3DXMatrixIdentity g_identityMatrix
        
    Call D3DXQuaternionIdentity(m_quat)
    m_scale = 1
End Sub

'-----------------------------------------------------------------------------
' Name: GetUpdatedMatrix
' Desc: Returns the precomputed matrix
'-----------------------------------------------------------------------------
Public Function GetUpdatedMatrix() As D3DMATRIX
    GetUpdatedMatrix = m_MatConcat
End Function

'-----------------------------------------------------------------------------
' Name: AddRotation
' Param:
'       combinetype     PreMultiply/PostMultiply/or replace current matrix
'       x,y,z           axis of rotation
'       rot             radians to rotate
'-----------------------------------------------------------------------------
Public Function AddRotation(combinetype As CombineTypeEnum, X As Single, Y As Single, Z As Single, rot As Single)
    Dim mat As D3DMATRIX
    D3DXMatrixRotationAxis mat, Vec3(X, Y, Z), rot
    
    Select Case combinetype
        Case COMBINE_before
            D3DXMatrixMultiply m_Matrix, mat, m_Matrix
        Case COMBINE_after
            D3DXMatrixMultiply m_Matrix, m_Matrix, mat
        Case COMBINE_replace
            m_Matrix = mat
    End Select
End Function

'-----------------------------------------------------------------------------
' Name: AddScale
' Param:
'       combinetype     PreMultiply/PostMultiply/or replace current matrix
'       x,y,z           scale
'-----------------------------------------------------------------------------
Public Function AddScale(combinetype As CombineTypeEnum, X As Single, Y As Single, Z As Single)
    Dim mat As D3DMATRIX
    D3DXMatrixScaling mat, X, Y, Z
    Select Case combinetype
        Case COMBINE_before
            D3DXMatrixMultiply m_Matrix, mat, m_Matrix
        Case COMBINE_after
            D3DXMatrixMultiply m_Matrix, m_Matrix, mat
        Case COMBINE_replace
            m_Matrix = mat
    End Select
End Function

'-----------------------------------------------------------------------------
' Name: AddTranslation
' Param:
'       combinetype     PreMultiply/PostMultiply/or replace current matrix
'       x,y,z           translation
'-----------------------------------------------------------------------------
Public Function AddTranslation(combinetype As CombineTypeEnum, X As Single, Y As Single, Z As Single)
    Dim mat As D3DMATRIX
    D3DXMatrixTranslation mat, X, Y, Z
    Select Case combinetype
        Case COMBINE_before
            D3DXMatrixMultiply m_Matrix, mat, m_Matrix
        Case COMBINE_after
            D3DXMatrixMultiply m_Matrix, m_Matrix, mat
        Case COMBINE_replace
            m_Matrix = mat
    End Select
End Function

'-----------------------------------------------------------------------------
' Name: SetPosition
' Desc: sets position of the object in the parents coordinate system
' Note: will replace whatever matrix is set
'-----------------------------------------------------------------------------
Sub SetPosition(vec As D3DVECTOR)
    m_pos = vec
    UpdateMatrix
End Sub
'-----------------------------------------------------------------------------
' Name: SetOrientation
' Desc: sets the orientation of the object
' Note: use D3DXQuaternionRotationAxis to create the quaternion
' Note: will replace whatever matrix is set
'-----------------------------------------------------------------------------
Sub SetOrientation(quat As D3DQUATERNION)
    m_quat = quat
    UpdateMatrix
End Sub

'-----------------------------------------------------------------------------
' Name: GetOrientation
' Desc: returns what ever was passed to SetOrientation
' Note: does not inspect the matrix to obtain orientation
'-----------------------------------------------------------------------------
Function GetOrientation() As D3DQUATERNION
    GetOrientation = m_quat
End Function

Function RotationAxis(X As Single, Y As Single, Z As Single, rotationInDegrees As Single) As D3DQUATERNION
    Dim quat As D3DQUATERNION
    D3DXQuaternionRotationAxis quat, Vec3(X, Y, Z), (rotationInDegrees / 180) * sngPi
    RotationAxis = quat
End Function
'-----------------------------------------------------------------------------
' Name: UpdateMatrix
' Desc: Aux function for SetPosition, SetOrientation, SetScale
'-----------------------------------------------------------------------------
Private Sub UpdateMatrix()
    D3DXMatrixAffineTransformation m_Matrix, m_scale, m_rotcenter, m_quat, m_pos
End Sub

'-----------------------------------------------------------------------------
' Name: SetMatrix
' Desc: Sets the matrix for the local coordinate system
' Note: This overrides any previous calls to functions such as
'       SetPosition, SetScale, SetOrientation
'-----------------------------------------------------------------------------
Public Function SetMatrix(m As D3DMATRIX)
    m_Matrix = m
End Function
