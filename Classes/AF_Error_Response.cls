VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AF_Error_Response"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarErrorType As AFErrorTypeEnum 'local copy
Private mvarCoilType As coil_type 'local copy
Private mvarCodeLevel As String 'local copy
Private mvarErrorAction As AFErrorActionEnum 'local copy

Public Enum AFErrorActionEnum
    ExpressError = 0
    SuppressError = 1
End Enum

Public Property Let ErrorAction(ByVal vData As AFErrorActionEnum)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ErrorType = Form1
    mvarErrorAction = vData
End Property


Public Property Get ErrorAction() As AFErrorActionEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ErrorType
    ErrorAction = mvarErrorAction
End Property

Public Property Let CodeLevel(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CodeLevel = 5
    mvarCodeLevel = vData
End Property


Public Property Get CodeLevel() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CodeLevel
    CodeLevel = mvarCodeLevel
End Property



Public Property Let CoilType(ByVal vData As coil_type)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.CoilType = Form1
    mvarCoilType = vData
End Property


Public Property Get CoilType() As coil_type
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CoilType
    CoilType = mvarCoilType
End Property



Public Property Let ErrorType(ByVal vData As AFErrorTypeEnum)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ErrorType = Form1
    mvarErrorType = vData
End Property


Public Property Get ErrorType() As AFErrorTypeEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ErrorType
    ErrorType = mvarErrorType
End Property



