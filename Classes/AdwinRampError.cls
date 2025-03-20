VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdwinRampError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'local variable(s) to hold property value(s)
Private mvarErrorType As ADwinRampErrorType 'local copy
Private mvarMessage As String
Private mvarErrorName As String

Public Enum ADwinRampErrorType
    None = 0
    MonitorVoltage_BelowZeroThreshold = 1
    MonitorVoltage_DidNotReachTarget = 2
    MonitorVoltage_OvershootTarget = 3
End Enum

Public Property Let ErrorType(ByVal value As ADwinRampErrorType)

    mvarErrorType = value

End Property

Public Property Get ErrorType() As ADwinRampErrorType

    ErrorType = mvarErrorType

End Property

Public Property Let Message(ByVal value As String)

    mvarMessage = value

End Property

Public Property Get Message() As String

    Message = mvarMessage

End Property

Public Property Let ErrorName(ByVal value As String)

    mvarErrorName = value

End Property

Public Property Get ErrorName() As String

    ErrorName = mvarErrorName

End Property

