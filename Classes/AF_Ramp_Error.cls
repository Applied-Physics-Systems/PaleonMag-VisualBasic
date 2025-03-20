VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AF_Ramp_Error"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarMessage As String 'local copy
Private mvarErrorType As AFErrorTypeEnum 'local copy
Private mvarSource As String 'local copy
Private mvarStackTrace As String 'local copy
Private mvarErrNumber As Long 'local copy
Private mvarCoilType As coil_type 'local copy
Private mvarTimeStamp As Date 'local copy


Public Enum AFErrorTypeEnum
    NoError = 0
    ZeroMonitorVoltage = 1
    TargetUndershoot = 2
    TargetOvershoot = 4
    FatalError = 8
End Enum

Public Property Let TimeStamp(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TimeStamp = 5
    mvarTimeStamp = vData
End Property

Public Property Get TimeStamp() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TimeStamp
    TimeStamp = mvarTimeStamp
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

Public Property Let ErrNumber(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ErrNumber = 5
    mvarErrNumber = vData
End Property


Public Property Get ErrNumber() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ErrNumber
    ErrNumber = mvarErrNumber
End Property



Public Property Let StackTrace(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StackTrace = 5
    mvarStackTrace = vData
End Property


Public Property Get StackTrace() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StackTrace
    StackTrace = mvarStackTrace
End Property



Public Property Let Source(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Source = 5
    mvarSource = vData
End Property


Public Property Get Source() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Source
    Source = mvarSource
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



Public Property Let Message(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Message = 5
    mvarMessage = vData
End Property


Public Property Get Message() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Message
    Message = mvarMessage
End Property



Private Sub Class_Initialize()
    
    mvarCoilType = coil_type.Unknown
    mvarErrorType = NoError
    mvarMessage = ""
    mvarSource = ""
    mvarErrNumber = 0
    mvarStackTrace = ""
    mvarTimeStamp = Now
    
End Sub
