VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAnimation"
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

Public objectname As String

Private Type KEYHEADER
    keytype As Long
    keycount As Long
End Type

Private Type D3DMATRIXKEY
    time As Long
    nFloats As Long
    matrix As D3DMATRIX
End Type

Dim m_MatrixKeys() As D3DMATRIXKEY
Dim m_NumMatrixKeys As Long

Dim m_strFrameName As String
Dim m_frame As clsFrame

Dim m_Children() As clsAnimation
Dim m_NumChildren As Long
'-----------------------------------------------------------------------------
' Name: InitFromFile
' Desc: character called this
'-----------------------------------------------------------------------------
Public Function InitFromFile(strFile As String, framework As clsFrame) 'i am the ROOT
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
  
        If ChildData.GetType = "TID_D3DRMAnimationSet" Then
                Me.ParseAnimSet ChildData, framework
        End If
    Loop
End Function
'-----------------------------------------------------------------------------
' Name: ParseAnimSet
' Desc: called from D3DUtil_LoadFromFile
'-----------------------------------------------------------------------------
Friend Sub ParseAnimSet(FileData As DirectXFileData, parentFrame As clsFrame)
    objectname = FileData.GetName()
    
    Dim ChildData As DirectXFileData
    Dim NewAnim As clsAnimation

     Do
        Set ChildData = FileData.GetNextObject()
        If ChildData Is Nothing Then Exit Do

            If ChildData.GetType = "TID_D3DRMAnimation" Then
                Set NewAnim = New clsAnimation
                AddChild NewAnim
                NewAnim.ParseAnim ChildData, Me, parentFrame
            End If
    Loop
End Sub
'-----------------------------------------------------------------------------
' Name: ParseAnim
' Desc: Called by ParseAnimSet
'-----------------------------------------------------------------------------
Friend Sub ParseAnim(FileData As DirectXFileData, parentAnimation As clsAnimation, parentFrame As clsFrame)
    On Local Error Resume Next
    objectname = FileData.GetName()
    
    Dim KeyHead As KEYHEADER

    Dim ChildObj As DirectXFileObject
    Dim ChildData As DirectXFileData
    Dim ChildReference As DirectXFileReference
    Dim I As Long

    Do
        Set ChildObj = FileData.GetNextObject()
        If ChildObj Is Nothing Then Exit Do
  
        Set ChildReference = ChildObj

            Set ChildData = ChildReference.Resolve()
                    
            If ChildData.GetType = "TID_D3DRMFrame" Then
                m_strFrameName = ChildData.GetName()
                Set m_frame = parentFrame.FindChildObject(m_strFrameName)
            End If
            
        Set ChildData = ChildObj
        
        If ChildData.GetType = "TID_D3DRMAnimationKey" Then
                    ChildData.GetDataFromOffset "", 0, 8, KeyHead
                    Select Case KeyHead.keytype
                        Case 4      'MATRIX KEY
                            ReDim m_MatrixKeys(KeyHead.keycount)
                            ChildData.GetDataFromOffset "", 8, Len(m_MatrixKeys(0)) * KeyHead.keycount, m_MatrixKeys(0)
                            m_NumMatrixKeys = KeyHead.keycount
                    End Select
        End If
                    
        Set ChildData = Nothing
        Set ChildReference = Nothing
        
    Loop
End Sub
'-----------------------------------------------------------------------------
' Name: AddChild
' Desc: Add child animation
'-----------------------------------------------------------------------------
Public Sub AddChild(child As clsAnimation)
    ReDim Preserve m_Children(m_NumChildren)
    Set m_Children(m_NumChildren) = child
    m_NumChildren = m_NumChildren + 1
End Sub
'-----------------------------------------------------------------------------
' Name: SetTime
' Desc: Sets the matrix of the frame being animated
'-----------------------------------------------------------------------------
Public Sub SetTime(t As Single)
    Dim LastT As Single
    Dim key1 As Single
    Dim key2 As Single
    Dim timediff As Single
    Dim scalar As Single
    Dim t2 As Single
    Dim I As Long

    'Check children
    Dim child As clsAnimation
    For I = 0 To m_NumChildren - 1
        Set child = m_Children(I)
        If Not child Is Nothing Then
            child.SetTime t
            Set child = Nothing
        End If
    Next
        
    If m_frame Is Nothing Then Exit Sub


    'loop matrix keys

        LastT = m_MatrixKeys(m_NumMatrixKeys - 1).time
        If t > LastT Then
            I = t \ LastT
            t2 = t - I * LastT
        End If

    Dim mat1 As D3DMATRIX
    Dim mat2 As D3DMATRIX
    Dim mat3 As D3DMATRIX

        'linear search
        For I = 0 To m_NumMatrixKeys
        If m_MatrixKeys(I).time > t2 Then
        
    If I = 0 Then
        mat1 = m_MatrixKeys(m_NumMatrixKeys - 1).matrix
        key1 = m_MatrixKeys(m_NumMatrixKeys - 1).time
    Else
        mat1 = m_MatrixKeys(I - 1).matrix
        key1 = m_MatrixKeys(I - 1).time
    End If

        mat2 = m_MatrixKeys(I).matrix
        key2 = m_MatrixKeys(I).time
        mat3 = MatSubtract(mat2, mat1)
        timediff = key2 - key1
        scalar = (t2 - key1) / timediff
        
    Dim matTemp As D3DMATRIX

    matTemp = MatScalar(mat3, scalar)
    mat3 = MatAdd(mat1, matTemp)

                m_frame.SetMatrix mat3
                Exit Sub
            End If
        Next
End Sub
