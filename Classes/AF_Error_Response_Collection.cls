VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AF_Error_Resp_Collection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"AF_Error_Response"
Attribute VB_Ext_KEY = "Member0" ,"AF_Error_Response"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection

Public Sub Clear()

    Class_Terminate
    Class_Initialize
    
End Sub

Public Function GetErrorResponse(ByRef af_ramp_error As af_ramp_error, ByVal code_level As String) As AF_Error_Response

    Set GetErrorResponse = New AF_Error_Response

    If af_ramp_error Is Nothing Then
        
        GetErrorResponse.ErrorAction = SuppressError
        Exit Function
        
    End If
    
    If af_ramp_error.ErrorType = NoError Then
        
        GetErrorResponse.ErrorAction = SuppressError
        Exit Function
        
    End If
    
    If mCol Is Nothing Then
    
        GetErrorResponse.ErrorAction = ExpressError
        Exit Function
        
    End If
    
    If mCol.Count <= 0 Then
    
        GetErrorResponse.ErrorAction = ExpressError
        Exit Function
        
    End If
    
    Dim af_error_resp As AF_Error_Response
    
    For Each af_error_resp In mCol
    
        If af_error_resp.ErrorType = af_ramp_error.ErrorType And _
           af_error_resp.CoilType = af_ramp_error.CoilType And _
           af_error_resp.CodeLevel = code_level Then
        
           Set GetErrorResponse = af_error_resp
           Exit Function
            
        End If
    
    Next af_error_resp
    
    'If error not found in look-up table, suppress it
    GetErrorResponse.ErrorAction = SuppressError

End Function

Public Function GenerateCollectionIndex(ErrorType As AFErrorTypeEnum, CoilType As coil_type, CodeLevel As String) As String

    Dim index As String

    index = ""

    Select Case ErrorType
    
        Case AFErrorTypeEnum.FatalError
                
            index = index & "Fatal_"
            
        Case AFErrorTypeEnum.TargetOvershoot
        
            index = index & "Overshoot_"
        
        Case AFErrorTypeEnum.TargetUndershoot
        
            index = index & "Undershoot_"
        
        Case AFErrorTypeEnum.ZeroMonitorVoltage
        
            index = index & "Zero_"
        
        Case Else
        
            index = index & "Other_"
            
    End Select

    Select Case CoilType
    
        Case coil_type.Axial
        
            index = index & "Axial_"
        
        Case coil_type.Transverse
        
            index = index & "Transverse_"
            
        Case Else
        
            index = index & "NoCoil_"
            
    End Select
    
    GenerateCollectionIndex = index & CodeLevel

End Function

Public Function Add(error_type As AFErrorTypeEnum, _
                    coil_type As coil_type, _
                    code_level As String, _
                    error_action As AFErrorActionEnum) As AF_Error_Response
                    
    'create a new object
    Dim objNewMember As AF_Error_Response
    Set objNewMember = New AF_Error_Response


    'set the properties passed into the method
    objNewMember.ErrorType = error_type
    objNewMember.CoilType = coil_type
    objNewMember.CodeLevel = code_level
    objNewMember.ErrorAction = error_action
    
    Dim index As String
    index = GenerateCollectionIndex(error_type, coil_type, code_level)
        
    On Error Resume Next
        mCol.Add objNewMember, index
        
        'If error type already exists in collection
        'Update it
        If Err.number <> 0 Then
        
            mCol.Remove index
            mCol.Add objNewMember, index
            
        End If
    On Error GoTo 0
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(ErrorType As AFErrorTypeEnum, CoilType As coil_type, CodeLevel As String) As AF_Error_Response
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(GenerateCollectionIndex(ErrorType, CoilType, CodeLevel))
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(ErrorType As AFErrorTypeEnum, CoilType As coil_type, CodeLevel As String)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove GenerateCollectionIndex(ErrorType, CoilType, CodeLevel)
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

