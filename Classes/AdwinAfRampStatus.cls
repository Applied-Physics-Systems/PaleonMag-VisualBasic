VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdwinAfRampStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarRamp_Start_Time As Date 'local copy
Private mvarRamp_End_Time As Date 'local copy
Private mvarWasSuccessful As Boolean 'local copy
Private mvarErrorType As AFErrorTypeEnum
Private mvarTargetField As String
Private mvarCoil As String

Public Property Let ErrorType(ByVal value As AFErrorTypeEnum)

    mvarErrorType = value

End Property

Public Property Get ErrorType() As AFErrorTypeEnum

    ErrorType = mvarErrorType

End Property

Public Property Let Coil(ByVal value As String)

    mvarCoil = value

End Property

Public Property Get Coil() As String

    Coil = mvarCoil

End Property

Public Property Let TargetPeakField(ByVal value As String)

    mvarTargetField = value

End Property

Public Property Get TargetPeakField() As String

    TargetPeakField = mvarTargetField

End Property

Public Property Let WasSuccessful(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WasSuccessful = 5
    mvarWasSuccessful = vData
End Property


Public Property Get WasSuccessful() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WasSuccessful
    WasSuccessful = mvarWasSuccessful
End Property

Public Function GetProgramRampDurationInSeconds() As Long
    
    Dim ret_val As Long
    
    On Error Resume Next
    
    ret_val = DateDiff("s", mvarRamp_Start_Time, mvarRamp_End_Time)
    
    If Err.number <> 0 Then
    
        ret_val = 0
        
    End If
    
    On Error GoTo 0
    
    GetProgramRampDurationInSeconds = ret_val
        
End Function

Public Property Let Ramp_End_Time(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Ramp_End_Time = 5
    mvarRamp_End_Time = vData
End Property


Public Property Get Ramp_End_Time() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Ramp_End_Time
    Ramp_End_Time = mvarRamp_End_Time
End Property



Public Property Let Ramp_Start_Time(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Ramp_Start_Time = 5
    mvarRamp_Start_Time = vData
End Property


Public Property Get Ramp_Start_Time() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Ramp_Start_Time
    Ramp_Start_Time = mvarRamp_Start_Time
End Property

Public Function GetErrorType_AsString() As String

    Select Case mvarErrorType
    
        Case AFErrorTypeEnum.FatalError
        
            GetErrorType_AsString = "FatalError"
        
        Case AFErrorTypeEnum.NoError
        
            GetErrorType_AsString = "NoError"
        
        Case AFErrorTypeEnum.TargetOvershoot
        
            GetErrorType_AsString = "TargetOvershoot"
        
        Case AFErrorTypeEnum.TargetUndershoot
        
            GetErrorType_AsString = "TargetUndershoot"
        
        Case AFErrorTypeEnum.ZeroMonitorVoltage
        
            GetErrorType_AsString = "ZeroMonitorVoltage"
        
        Case Else
        
            GetErrorType_AsString = "Unknown"
        
    End Select
    
End Function

Public Function GetErrorShortType_AsString() As String

    Select Case mvarErrorType
    
        Case AFErrorTypeEnum.FatalError
        
            GetErrorShortType_AsString = "Fatal"
        
        Case AFErrorTypeEnum.NoError
        
            GetErrorShortType_AsString = "NoError"
        
        Case AFErrorTypeEnum.TargetOvershoot
        
            GetErrorShortType_AsString = "Overshoot"
        
        Case AFErrorTypeEnum.TargetUndershoot
        
            GetErrorShortType_AsString = "Undershoot"
        
        Case AFErrorTypeEnum.ZeroMonitorVoltage
        
            GetErrorShortType_AsString = "ZeroVoltage"
        
        Case Else
        
            GetErrorShortType_AsString = "Unknown"
        
    End Select
    
End Function




