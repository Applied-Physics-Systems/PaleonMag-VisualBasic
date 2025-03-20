VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmDCMotors 
   Caption         =   "Motor control"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   8985
   Begin VB.CommandButton LoadButton 
      Caption         =   "Load"
      Height          =   255
      Left            =   5520
      TabIndex        =   61
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtNewY 
      Height          =   285
      Left            =   4440
      TabIndex        =   60
      Text            =   "0"
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton ButtonYSet 
      Caption         =   "Go to Y:"
      Height          =   255
      Left            =   3600
      TabIndex        =   59
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtNewX 
      Height          =   285
      Left            =   2760
      TabIndex        =   58
      Text            =   "0"
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton ButtonXSet 
      Caption         =   "Go to X:"
      Height          =   255
      Left            =   1920
      TabIndex        =   57
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton HomeToCenterButton 
      Caption         =   "Home to Center"
      Height          =   255
      Left            =   120
      TabIndex        =   56
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox chkConnectMotor 
      Caption         =   "Changer(Y)"
      Height          =   252
      Index           =   4
      Left            =   6480
      TabIndex        =   55
      Top             =   1080
      Width           =   1452
   End
   Begin VB.CommandButton cmdSetSCoilPos 
      Caption         =   "S Coil"
      Height          =   252
      Left            =   5520
      TabIndex        =   54
      Top             =   1200
      Width           =   612
   End
   Begin VB.CommandButton cmdSetIRMHi 
      Caption         =   "IRM Hi"
      Height          =   252
      Left            =   5520
      TabIndex        =   53
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdSetIRMLo 
      Caption         =   "IRM Lo"
      Height          =   252
      Left            =   5520
      TabIndex        =   52
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtSpinRPS 
      Height          =   288
      Left            =   3360
      TabIndex        =   51
      Top             =   3000
      Width           =   612
   End
   Begin VB.CommandButton cmdSpinTurningMotor 
      Caption         =   "Spin Sample (rps):"
      Height          =   252
      Left            =   1800
      TabIndex        =   50
      Top             =   3000
      Width           =   1452
   End
   Begin VB.CommandButton cmdSetChangerHole 
      Caption         =   "Set Current Hole"
      Height          =   255
      Left            =   2280
      TabIndex        =   49
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Frame frmActiveControls 
      Caption         =   "Active Controls"
      Height          =   1335
      Left            =   6360
      TabIndex        =   44
      Top             =   1560
      Width           =   1692
      Begin VB.OptionButton optMotorActive 
         Caption         =   "Changer (Y)"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   48
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optMotorActive 
         Caption         =   "Up/Down"
         Enabled         =   0   'False
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   47
         Top             =   720
         Width           =   1452
      End
      Begin VB.OptionButton optMotorActive 
         Caption         =   "Turning"
         Enabled         =   0   'False
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   46
         Top             =   480
         Width           =   1452
      End
      Begin VB.OptionButton optMotorActive 
         Caption         =   "Changer (X)"
         Enabled         =   0   'False
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   1452
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connections"
      Height          =   1335
      Left            =   6360
      TabIndex        =   40
      Top             =   120
      Width           =   1692
      Begin VB.CheckBox chkConnectMotor 
         Caption         =   "Up/Down"
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   43
         Top             =   720
         Width           =   1452
      End
      Begin VB.CheckBox chkConnectMotor 
         Caption         =   "Turning"
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   42
         Top             =   480
         Width           =   1452
      End
      Begin VB.CheckBox chkConnectMotor 
         Caption         =   "Changer (X)"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   1452
      End
   End
   Begin VB.CommandButton cmdSampleDropOff 
      Caption         =   "Sample Dropoff"
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   2040
      Width           =   1572
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   372
      Left            =   360
      TabIndex        =   38
      Top             =   5040
      Width           =   1332
   End
   Begin VB.CommandButton buttonHeightSet 
      Caption         =   "Change Height:"
      Height          =   255
      Left            =   1920
      TabIndex        =   37
      Top             =   1080
      Width           =   1572
   End
   Begin VB.TextBox txtNewHeight 
      Height          =   285
      Left            =   3600
      TabIndex        =   36
      Text            =   "0"
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton RelabelPosButton 
      Caption         =   "Relabel P"
      Height          =   255
      Left            =   1200
      TabIndex        =   35
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton SetMeasButton 
      Caption         =   "Meas."
      Height          =   252
      Left            =   4800
      TabIndex        =   34
      Top             =   1560
      Width           =   612
   End
   Begin VB.CommandButton SetZeroButton 
      Caption         =   "Zero"
      Height          =   252
      Left            =   4800
      TabIndex        =   33
      Top             =   1200
      Width           =   612
   End
   Begin VB.CommandButton SetAfButton 
      Caption         =   "Af Coil"
      Height          =   252
      Left            =   4800
      TabIndex        =   32
      Top             =   840
      Width           =   612
   End
   Begin VB.CommandButton SetTopButton 
      Caption         =   "Top"
      Height          =   252
      Left            =   4800
      TabIndex        =   31
      Top             =   480
      Width           =   612
   End
   Begin VB.TextBox HoleTargetText 
      Height          =   285
      Left            =   3600
      TabIndex        =   29
      Text            =   "0"
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox AngleTargetText 
      Height          =   285
      Left            =   3600
      TabIndex        =   28
      Text            =   "0"
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton ChangeHoleButton 
      Caption         =   "Change Hole:"
      Height          =   255
      Left            =   1920
      TabIndex        =   27
      Top             =   2040
      Width           =   1572
   End
   Begin VB.CommandButton ChangeTurnAngleButton 
      Caption         =   "Change Turn Angle:"
      Height          =   255
      Left            =   1920
      TabIndex        =   26
      Top             =   1560
      Width           =   1572
   End
   Begin VB.CommandButton ReadHoleButton 
      Caption         =   "Read hole"
      Height          =   252
      Left            =   7080
      TabIndex        =   25
      Top             =   4560
      Width           =   1332
   End
   Begin VB.CommandButton ReadAngleButton 
      Caption         =   "Read angle"
      Height          =   252
      Left            =   7080
      TabIndex        =   24
      Top             =   4200
      Width           =   1332
   End
   Begin VB.TextBox ChangerHoleBox 
      Height          =   285
      Left            =   5160
      TabIndex        =   22
      Top             =   4560
      Width           =   1812
   End
   Begin VB.TextBox TurningAngleBox 
      Height          =   285
      Left            =   5160
      TabIndex        =   20
      Top             =   4200
      Width           =   1812
   End
   Begin VB.TextBox txtPollPosition 
      Height          =   285
      Left            =   5160
      TabIndex        =   19
      Top             =   3840
      Width           =   1812
   End
   Begin VB.CommandButton ReadPosButton 
      Caption         =   "Read position"
      Height          =   252
      Left            =   7080
      TabIndex        =   17
      Top             =   3840
      Width           =   1332
   End
   Begin VB.CommandButton MotorResetButton 
      Caption         =   "Reset"
      Height          =   372
      Left            =   5640
      TabIndex        =   16
      Top             =   5040
      Width           =   852
   End
   Begin VB.CommandButton MotorHaltButton 
      BackColor       =   &H000000FF&
      Caption         =   "HALT!"
      Height          =   372
      Left            =   7680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5040
      UseMaskColor    =   -1  'True
      Width           =   852
   End
   Begin VB.CommandButton MotorStopButton 
      Caption         =   "Stop"
      Height          =   372
      Left            =   6600
      TabIndex        =   14
      Top             =   5040
      Width           =   852
   End
   Begin VB.CommandButton ClearPollStatusButton 
      Caption         =   "Clear Poll Status"
      Height          =   492
      Left            =   1920
      TabIndex        =   13
      Top             =   3960
      Width           =   852
   End
   Begin VB.CommandButton PollMotorButton 
      Caption         =   "Poll Motor"
      Height          =   372
      Left            =   720
      TabIndex        =   12
      Top             =   4080
      Width           =   852
   End
   Begin VB.TextBox MoveMotorVelocityEdit 
      Height          =   285
      Left            =   2880
      TabIndex        =   9
      Text            =   "4000000"
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton ZeroTargetPosButton 
      Caption         =   "Zero T/P"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton SamplePickupButton 
      Caption         =   "Sample Pickup"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   1572
   End
   Begin VB.CommandButton HomeToTopButton 
      Caption         =   "Home to Top"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1572
   End
   Begin VB.TextBox MoveMotorPosEdit 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Text            =   "0"
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton MoveMotorButton 
      Caption         =   "Move to Position:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1572
   End
   Begin VB.TextBox txtInputText 
      Height          =   285
      Left            =   5160
      TabIndex        =   2
      Top             =   3480
      Width           =   3015
   End
   Begin VB.TextBox OutputText 
      Height          =   285
      Left            =   5160
      TabIndex        =   0
      Top             =   3000
      Width           =   3015
   End
   Begin MSCommLib.MSComm MSCommMotor 
      Index           =   3
      Left            =   2880
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      BaudRate        =   57600
   End
   Begin MSCommLib.MSComm MSCommMotor 
      Index           =   1
      Left            =   3600
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      BaudRate        =   57600
   End
   Begin MSCommLib.MSComm MSCommMotor 
      Index           =   2
      Left            =   4440
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      BaudRate        =   57600
   End
   Begin MSCommLib.MSComm MSCommMotor 
      Index           =   4
      Left            =   2160
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      BaudRate        =   57600
   End
   Begin VB.Label Label5 
      Caption         =   "Set position"
      Height          =   252
      Left            =   4800
      TabIndex        =   30
      Top             =   120
      Width           =   1212
   End
   Begin VB.Label Label4 
      Caption         =   "Last Hole Read"
      Height          =   255
      Left            =   3840
      TabIndex        =   23
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Last Turn Angle"
      Height          =   255
      Left            =   3840
      TabIndex        =   21
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Last Pos Read"
      Height          =   255
      Left            =   3840
      TabIndex        =   18
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label13 
      Caption         =   "Velocity"
      Height          =   252
      Left            =   2880
      TabIndex        =   11
      Top             =   240
      Width           =   972
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Target Pos"
      Height          =   372
      Left            =   1920
      TabIndex        =   10
      Top             =   120
      Width           =   612
   End
   Begin VB.Label Label2 
      Caption         =   "Input:"
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Output:"
      Height          =   255
      Left            =   4440
      TabIndex        =   1
      Top             =   3000
      Width           =   615
   End
End
Attribute VB_Name = "frmDCMotors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const lastCmdMove As Integer = 1
Const lastCmdDropoff As Integer = 2
Const lastCmdPickup As Integer = 3

Const MoveXYMotorsToLimitSwitch_TimeoutSeconds As Integer = 60
Const HomeToCenter_NoLoadCorner_PreHomeHole As Integer = 55

'7/21/23Const MotorPositionMoveToLoadCorner = 50000
'7/21/23Const MotorPositionMoveToCenter = -50000


Const MotorPositionMoveToLoadCorner = 900000
Const MotorPositionMoveToCenter = -900000

Dim MotorsLocked As Integer
Private InputText(4) As String
Private lastPosition(4) As Long
Public lastMoveCommand As Integer
Public lastMoveMotor As Integer
Public lastMoveTarget As Long
Private currenthole As Double
Private ComPortAssignments(4) As Integer
Private OverlappingComPorts As Boolean
Dim UpDownSpeeds(2) As Long

Private Function ActiveMotorControls() As Integer
    Dim i As Integer
    ActiveMotorControls = 0
    For i = 1 To 4
        If optMotorActive(i) Then ActiveMotorControls = i
    Next i
End Function

Public Sub AssignCommPorts()
    OverlappingComPorts = False
    ComPortAssignments(MotorChanger) = MotorChanger
    If COMPortTurning = COMPortChanger Then
        ComPortAssignments(MotorTurning) = ComPortAssignments(MotorChanger)
        OverlappingComPorts = True
    Else
        ComPortAssignments(MotorTurning) = MotorTurning
    End If
    If COMPortUpDown = COMPortChanger Then
        ComPortAssignments(MotorUpDown) = ComPortAssignments(MotorChanger)
        OverlappingComPorts = True
    ElseIf COMPortUpDown = COMPortTurning Then
        ComPortAssignments(MotorUpDown) = ComPortAssignments(MotorTurning)
        OverlappingComPorts = True
    Else
        ComPortAssignments(MotorUpDown) = MotorUpDown
    End If
    If COMPortChangerY = COMPortChanger Then
        ComPortAssignments(MotorChangerY) = ComPortAssignments(MotorChanger)
        OverlappingComPorts = True
    ElseIf COMPortChangerY = COMPortTurning Then
        ComPortAssignments(MotorChangerY) = ComPortAssignments(MotorTurning)
        OverlappingComPorts = True
    ElseIf COMPortChangerY = COMPortUpDown Then
        ComPortAssignments(MotorChangerY) = ComPortAssignments(MotorUpDown)
        OverlappingComPorts = True
    Else
        ComPortAssignments(MotorChangerY) = MotorChangerY
    End If
End Sub

Private Sub buttonHeightSet_Click()
    UpDownMove txtNewHeight, 0
End Sub

Private Sub ButtonXSet_Click()
    If Not modConfig.HasXYTableBeenHomed Then
    
        frmProgram.StatBarNew "Home to top..."
        modMotor.MotorUPDN_TopReset
        modMotor.MotorXYTable_CenterReset
    
    End If

    MoveMotorXY MotorChanger, val(LTrim$(txtNewX.text)), ChangerSpeed, False, True
End Sub

Private Sub ButtonYSet_Click()
    If Not modConfig.HasXYTableBeenHomed Then
    
        frmProgram.StatBarNew "Home to top..."
        modMotor.MotorUPDN_TopReset
        modMotor.MotorXYTable_CenterReset
    
    End If

    MoveMotorAbsoluteXY MotorChangerY, val(LTrim$(txtNewY.text)), ChangerSpeed, False, True
End Sub

Private Sub ChangeHoleButton_Click()
    ChangerMotortoHole val(HoleTargetText)
End Sub

Public Function ChangerHole() As Double
    Dim curhole As Double
    Dim curpos As Long
    Dim curposY As Long
    
    If NOCOMM_MODE Then
        ChangerHole = currenthole
    Else
    If Not UseXYTableAPS Then
    'Chain Drive
        curpos = ReadPosition(MotorChanger)
        curhole = ConvertPosToHole(curpos)
        currenthole = curhole
        ChangerHoleBox = curhole
        ChangerHole = curhole
    Else
    'We are using an XY Table
        curpos = ReadPosition(MotorChanger)
        curposY = ReadPosition(MotorChangerY)
        curhole = ConvertPosToHoleXY(curpos, curposY)
        currenthole = curhole
        ChangerHoleBox = curhole
        ChangerHole = curhole
    End If
    End If
End Function

Public Sub ChangerMotortoHole(ByVal hole As Double, Optional ByVal waitingForStop As Boolean = True, Optional ByVal pauseOverride As Boolean = True)
    Dim curhole As Double
    Dim startinghole As Double
    Dim startingPos As Double
    Dim curpos As Long
    Dim startingPosX As Double
    Dim curposX As Long
    Dim startingPosY As Double
    Dim curposY As Long
    Dim FullLoop As Long
    Dim target As Long
    Dim targetX As Long
    Dim targetY As Long
    Dim ErrorMessage As String
    If Not Prog_paused Or Prog_halted Then
        lastMoveCommand = lastCmdMove
        lastMoveMotor = MotorChanger
        lastMoveTarget = hole
    End If
    
    ' Let's get the sample rod out of the way if necessary
    If Abs(UpDownHeight) > Abs(SampleBottom) * 0.1 Then HomeToTop
    
    If Not UseXYTableAPS Then
        'This is the routine to use a Chain Drive System
        FullLoop = (SlotMax - SlotMin + 1) * OneStep
        startingPos = ReadPosition(MotorChanger)
        curpos = startingPos
        
        If curpos \ FullLoop <> 0 Then RelabelPos MotorChanger, (curpos Mod FullLoop)
        startinghole = ChangerHole
        target = ConvertHoletoPos(hole)
        SetSCurveNext MotorChanger, SCurveFactor
        MoveMotor MotorChanger, target, ChangerSpeed, waitingForStop, pauseOverride
        
        If Not waitingForStop Then Exit Sub
        curpos = ReadPosition(MotorChanger)
        
        If curpos \ FullLoop <> 0 Then RelabelPos MotorChanger, (curpos Mod FullLoop)
        
        If NOCOMM_MODE Then currenthole = hole
        curhole = ChangerHole
        
        'Because OneStep is negative, the criteria was never reach (always <0 and not >0.02) till the asolute value (May 2007 L Carporzen)
        If Not NOCOMM_MODE And (Abs(curpos - target) / Abs(OneStep)) > 0.02 Then
            ' First try to move to move back to the desired position
            SetSCurveNext MotorChanger, SCurveFactor
            MoveMotor MotorChanger, target, 0.5 * ChangerSpeed, pauseOverride:=True
            curpos = ReadPosition(MotorChanger)
        End If
        
        '  Quit if fail here, backing off a bit first ...
        If Not NOCOMM_MODE And (Abs(curpos - target) / Abs(OneStep)) > 0.02 Then
            ErrorMessage = "Unacceptable slop moving changer" & vbCrLf & _
                "from hole " & str(startinghole) & " to hole " & str(hole) & "." & _
                vbCrLf & vbCrLf & "Target position: " & str(target) & vbCrLf & _
                "Current position: " & str(curpos) & vbCrLf & vbCrLf & _
                "Execution has been paused. Please check machine."
            SetSCurveNext MotorChanger, SCurveFactor
            MoveMotor MotorChanger, (curpos - (curpos - startingPos) * 0.1), 0.1 * ChangerSpeed, pauseOverride:=True
            DelayTime 0.2
            Flow_Pause
            SetCodeLevel CodeRed
            frmSendMail.MailNotification "Unacceptable slop", ErrorMessage, CodeRed
            MsgBox ErrorMessage
            SetCodeLevel StatusCodeColorLevelPrior, True
        End If
    Else
        'We Are using the XY Table
        
        If Not modConfig.HasXYTableBeenHomed Then
        
            modMotor.MotorXYTable_CenterReset
        
        End If
        
        startingPosX = ReadPosition(MotorChanger)
        startingPosY = ReadPosition(MotorChangerY)
        curposX = startingPosX
        curposY = startingPosY
        
        startinghole = ChangerHole
        targetX = ConvertHoletoPosX(hole)
        targetY = ConvertHoletoPosY(hole)
        
        modListenAndLog.AppendRS232MessageToLogFile _
            "Move motor to hole: " & str(hole) & _
            ", current pos (X,Y): (" & str(curposX) & ", " & str(curposY) & _
            "), target pos (X,Y): (" & str(targetX) & ", " & str(targetY), _
            -1, _
            ""
        
        'Then Move Directly To the position, always approaching from the corner
        MoveMotorAbsoluteXY MotorChanger, targetX, ChangerSpeed, False, pauseOverride
        MoveMotorAbsoluteXY MotorChangerY, targetY, ChangerSpeed, waitingForStop, pauseOverride
        
        If Not waitingForStop Then Exit Sub
        curposX = ReadPosition(MotorChanger)
        curposY = ReadPosition(MotorChangerY)
        
        If NOCOMM_MODE Then currenthole = hole
        curhole = ChangerHole
        'Because OneStep is negative, the criteria was never reach (always <0 and not >0.02) till the asolute value (May 2007 L Carporzen)
        If Not NOCOMM_MODE And ((Abs(curposX - targetX) / Abs(OneStep)) > 0.02 Or (Abs(curposY - targetY) / Abs(OneStep)) > 0.02) Then
            ' First try to move to move back to the desired position
            MoveMotorAbsoluteXY MotorChanger, targetX, 0.5 * ChangerSpeed, pauseOverride:=True
            MoveMotorAbsoluteXY MotorChangerY, targetY, 0.5 * ChangerSpeed, pauseOverride:=True
            curposX = ReadPosition(MotorChanger)
            curposY = ReadPosition(MotorChanger)
        End If
        '  Quit if fail here, backing off a bit first ...
        If Not NOCOMM_MODE And ((Abs(curpos - target) / Abs(OneStep)) > 0.02 Or (Abs(curpos - target) / Abs(OneStep)) > 0.02) Then
            ErrorMessage = "Unacceptable slop moving changer" & vbCrLf & _
                "from hole " & str(startinghole) & " to hole " & str(hole) & "." & _
                vbCrLf & vbCrLf & "Target position X: " & str(targetX) & vbCrLf & _
                "Current position X: " & str(curposX) & vbCrLf & vbCrLf & _
                "Target position Y: " & str(targetY) & vbCrLf & _
                "Current position Y: " & str(curposY) & vbCrLf & vbCrLf & _
                "Execution has been paused. Please check machine."
            MoveMotorAbsoluteXY MotorChanger, (curposX - (curposX - startingPosX) * 0.1), 0.1 * ChangerSpeed, pauseOverride:=True
            MoveMotorAbsoluteXY MotorChangerY, (curposY - (curposY - startingPosY) * 0.1), 0.1 * ChangerSpeed, pauseOverride:=True
            DelayTime 0.2
            Flow_Pause
            SetCodeLevel CodeRed
            frmSendMail.MailNotification "Unacceptable slop", ErrorMessage, CodeRed
            MsgBox ErrorMessage
            SetCodeLevel StatusCodeColorLevelPrior, True
            
        End If
    End If
    
End Sub

Private Sub ChangeTurnAngleButton_Click()
    TurningMotorRotate val(AngleTargetText)
End Sub

'*************************************************************************************************

Public Function CheckInternalStatus(motorid As Integer, bit As Integer) As Integer
    Dim inputchar As String, hexpos As String
    Dim decpos As Long
    SendCommand motorid, (MotorAddress(motorid) & "20") 'T Shuma 7-31-12 fixed to be compatible with all motors
    GetResponse motorid
    inputchar = InputText(motorid)
    CheckInternalStatus = -1
    If Len(inputchar) = 15 Then
        hexpos = Mid$(inputchar, 11, 4)
        decpos = CLng("&H" + hexpos)
    Else
        Exit Function
    End If
    decpos = (decpos \ 2 ^ (bit)) Mod 2
    CheckInternalStatus = decpos
    modListenAndLog.AppendRS232MessageToLogFile "Check Motor Status", MotorCOMPort(motorid), "Bit " & Trim(str$(bit)) & ", Status: " & Trim(str$(CheckInternalStatus))
End Function

Private Sub chkConnectMotor_Click(index As Integer)
    If chkConnectMotor(index) Then
        MotorCommConnect index
        If (index = 4) Then
        HomeToCenterButton.Visible = True
        LoadButton.Visible = True
        ButtonXSet.Visible = True
        txtNewX.Visible = True
        ButtonYSet.Visible = True
        txtNewY.Visible = True
        End If
    Else
        MotorCommDisconnect index
        If (index = 4) Then
        HomeToCenterButton.Visible = False
        LoadButton.Visible = False
        ButtonXSet.Visible = False
        txtNewX.Visible = False
        ButtonYSet.Visible = False
        txtNewY.Visible = False
        End If
    End If
End Sub

Private Sub ClearPollStatus(motorid As Integer)
    'Clear the polling status
    SendCommand motorid, (MotorAddress & "1 65535")
    GetResponse motorid
End Sub

Private Sub ClearPollStatusButton_Click()
    ClearPollStatus ActiveMotorControls
End Sub

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub cmdSampleDropOff_Click()
    SampleDropOff pauseOverride:=True
End Sub

Private Sub cmdSetChangerHole_Click()
    SetChangerHole val(HoleTargetText)
End Sub

Private Sub cmdSetIRMHi_Click()
    'MoveMotorPosEdit.text = Str(IRMHiPos)
    'txtNewHeight = Str(IRMHiPos)
End Sub

Private Sub cmdSetIRMLo_Click()
    'MoveMotorPosEdit.text = Str(IRMPos)
    'txtNewHeight = Str(IRMPos)
End Sub

Private Sub cmdSetSCoilPos_Click()
    MoveMotorPosEdit.text = str(SCoilPos)
    txtNewHeight = str(SCoilPos)
End Sub

Private Sub cmdSpinTurningMotor_Click()
    TurningMotorSpin val(txtSpinRPS), pauseOverride:=True
End Sub

Public Sub ConnectDCMotors()

    MotorCommConnect MotorChanger
    MotorCommConnect MotorChangerY
    MotorCommConnect MotorUpDown
    MotorCommConnect MotorTurning

End Sub

Public Function ConvertAngleToPos(angle As Double) As Long
    ConvertAngleToPos = Int(-TurningMotorFullRotation * angle / 360)
End Function

Public Function ConvertHoletoPos(hole As Double) As Long
    Dim currenthole As Double
    Dim currentPos As Long
    Dim TargetHolePosRaw As Long
    Dim TargetHole As Double
    Dim FullLoop As Double
    Dim StepsToGo As Double
    FullLoop = Abs((SlotMax - SlotMin + 1) * OneStep)
    currentPos = ReadPosition(MotorChanger)
    currenthole = ChangerHole()
    TargetHole = hole
    TargetHolePosRaw = OneStep * hole
    StepsToGo = (TargetHolePosRaw - currentPos) Mod FullLoop
    If Abs(StepsToGo) > (FullLoop / 2) Then
        If StepsToGo > 0 Then
            StepsToGo = StepsToGo - FullLoop
        Else
            StepsToGo = StepsToGo + FullLoop
        End If
    End If
    If Not Changer_isHole(TargetHole) Then
        If StepsToGo > 0 Then
            StepsToGo = StepsToGo + SampleHoleAlignmentOffset * OneStep
        Else
            StepsToGo = StepsToGo + SampleHoleAlignmentOffset * OneStep
        End If
    End If
    ConvertHoletoPos = Int(StepsToGo + currentPos)
End Function

Public Function ConvertHoletoPosX(hole As Double) As Long
    
    ConvertHoletoPosX = modConfig.XYTablePositions(hole, 0)
End Function

Public Function ConvertHoletoPosY(hole As Double) As Long
    
    ConvertHoletoPosY = modConfig.XYTablePositions(hole, 1)
End Function

Public Function ConvertPosToAngle(pos As Long) As Double
    ConvertPosToAngle = (pos / -TurningMotorFullRotation) * 360
End Function

Public Function ConvertPosToHole(pos As Long) As Double
    Dim hole As Double
    Dim FullLoop As Double
    FullLoop = (SlotMax - SlotMin + 1)
    hole = (pos / OneStep) Mod FullLoop
    If hole <= 0 Then hole = hole + (SlotMax - SlotMin + 1)
    ConvertPosToHole = hole
End Function

Public Function ConvertPosToHoleXY(posX As Long, posY As Long) As Double
    Dim hole As Double
    Dim i As Integer
    Dim TestX As Double
    Dim TestXb As Boolean
    Dim TestY As Double
    Dim TestYb As Boolean
    
    ConvertPosToHoleXY = -1
    For i = 1 To 101
        
        TestX = Abs(posX - modConfig.XYTablePositions(i, 0))
        TestXb = TestX < 1000
        
        TestY = Abs(posY - modConfig.XYTablePositions(i, 1))
        TestYb = TestY < 1000
        
        If TestXb And TestYb Then
            ConvertPosToHoleXY = i
            i = 101
        End If
        
    Next i
    If ConvertPosToHoleXY = -1 Then
    ConvertPosToHoleXY = SlotMin
    End If
End Function

Private Sub Form_Load()
    
    MotorsLocked = False
    AssignCommPorts
    
    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me
    
    'Display XY buttons if needed
    If (UseXYTableAPS) Then
        HomeToCenterButton.Visible = True
        LoadButton.Visible = True
        ButtonXSet.Visible = True
        txtNewX.Visible = True
        ButtonYSet.Visible = True
        txtNewY.Visible = True
    Else
        HomeToCenterButton.Visible = False
        LoadButton.Visible = False
        ButtonXSet.Visible = False
        txtNewX.Visible = False
        ButtonYSet.Visible = False
        txtNewY.Visible = False
    End If
    
End Sub

Private Sub Form_Resize()
    Me.Height = 6000
    Me.Width = 9150
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    For i = 1 To 4
        If MSCommMotor(i).PortOpen = True Then
            MSCommMotor(i).PortOpen = False
        End If
    Next i
End Sub

Private Function GetResponse(motorid As Integer) As Boolean
    Dim Delay As Double
    Dim inputchar As String
    
    On Error GoTo CommError
    If Not IsValidMotorid(motorid) Then Exit Function
    If MSCommMotor(motorid).PortOpen = False Then
        MotorCommConnect motorid
    End If
    Delay = Timer   ' Set delaystart time.
    inputchar = vbNullString
    Do While (Not NOCOMM_MODE) And Right$(inputchar, 1) <> vbCr
        DoEvents
        If MSCommMotor(motorid).InBufferCount > 0 Then
            inputchar = inputchar + MSCommMotor(motorid).Input
        End If
        If Timer < Delay Then Delay = Delay - 86400
        If Timer - Delay > 0.3 Then
            'MsgBox "Timeout sending motor command"
            Exit Do
        End If
    Loop
    InputText(motorid) = inputchar
    txtInputText = inputchar
    
    If modConfig.LogMessages Then
        modListenAndLog.AppendRS232MessageToLogFile inputchar, MSCommMotor(motorid).CommPort, "Input"
    End If
    
    If (Left$(inputchar, 1) = "#") Then
        inputchar = Mid$(inputchar, 11)
        If (inputchar <> "0") Then
            GetResponse = True
        Else
            GetResponse = False
        End If
    Else
        If (Left$(inputchar, 1) = "*") Then
            GetResponse = True
        Else
            GetResponse = False
        End If
    End If
    If NOCOMM_MODE Then GetResponse = True
    
    On Error GoTo 0
    Exit Function
    
CommError:

    MotorCommDisconnect motorid
    MotorCommConnect motorid
    GetResponse = GetResponse(motorid)

End Function

Public Function GetUpDownPos() As Long

    GetUpDownPos = ReadPosition(MotorUpDown)
    
End Function

'*************************************************************************************************************
'
'   Home To Center  7/27/2012
'         Author: Toni Shuma
'
'         Purpose: To enable the Quicksilver motors to return to the home position, center of the XY Table
'                 correctly.  Both X and Y axis must find the center of the table and stop there correctly
'                 (Stop Condition = if logic bit #1 becomes true, stop all motion on Up/Down motor;
'                 When the home position limit switch is pressed, logic bit #1 changes from false to true.)
'
'                   Note: untested.  Requires planning of limit switches to fix code and verify code works properly
'
'*************************************************************************************************************

Public Function HomeToCenter(ByRef xpos As Long, ByRef ypos As Long, Optional pauseOveride As Boolean = False)
    
    'No homing to center if the program has been halted
    If Prog_halted Then Exit Function
        
    Dim start_time As Double
   
    Dim stop_state As Integer
    If modConfig.DCMotorHomeToTop_StopOnTrue Then
        stop_state = 1
    Else
        stop_state = 0
    End If
   
    'No homing to center if the Up/Down Motor is not homed
    If CheckInternalStatus(MotorUpDown, 4) <> stop_state Then
         SetCodeLevel CodeRed
        frmSendMail.MailNotification "Homing error!", "Tried to Home to center but UP/Down Motor not homed to top.  This could break control rod.", CodeRed
       
        MsgBox "Could not home to center!  Home to top not complete!"
       
        SetCodeLevel StatusCodeColorLevelPrior, True
        Exit Function
    End If
        
    'Reset both X and Y motor controllers
    ' This will reset both motor positions to 0 motor encoder counts, but this is okay
    ' as the homing process will rezero the motor positions relative to the
    ' center X and Y optical limit switches
    MotorReset MotorChanger
    MotorReset MotorChangerY
        
    'Wait 1 second for motor power cycle process to finish
    DelayTime 1
        
    start_time = Timer
        
    LockMotor MotorChanger
    ' Move motor a relative number of motor units, and stop if signal from the load corner limit switch is logic low
    MoveMotorXY MotorChanger, MotorPositionMoveToLoadCorner, ChangerSpeed, False, pauseOveride, -1, 0
    
    DelayTime 0.1
    
    LockMotor MotorChangerY
    MoveMotorXY MotorChangerY, MotorPositionMoveToLoadCorner, ChangerSpeed, False, pauseOveride, -2, 0
      
    'T Shuma - Move to load corner limit switches
    Do While ((CheckInternalStatus(MotorChanger, 4) <> 0) Or _
              (CheckInternalStatus(MotorChangerY, 5) <> 0)) _
             And Not NOCOMM_MODE _
             And Not Prog_halted _
             And Not HasMoveToXYLimit_Timedout(start_time)
    
        DelayTime 0.1
        
    Loop
        
    
    If Not (CheckInternalStatus(MotorChanger, 4) = 0 And CheckInternalStatus(MotorChangerY, 5) = 0) And Not NOCOMM_MODE Then
       
        SetCodeLevel CodeRed
       
        If (HasMoveToXYLimit_Timedout(start_time)) Then
        
           'Home to center has timed out
           frmSendMail.MailNotification "Homing error!", "Home XY Stage, move motors to Load Corner limit switches timed-out after: " & _
                                                         str$(MoveXYMotorsToLimitSwitch_TimeoutSeconds) & _
                                                         " seconds.", CodeRed
            
           MsgBox "Home XY Stage, move motors to Load Corner limit switches timed-out after: " & _
                  str$(MoveXYMotorsToLimitSwitch_TimeoutSeconds) & _
                  "seconds."
          
        Else
               
           frmSendMail.MailNotification "Homing error!", "Homed XY Stage to center but did not hit load corner limit switch(es).", CodeRed
            
           MsgBox "Homed XY Stage to center but did not hit load corner limit switch(es)!"
           
        End If
        
        SetCodeLevel StatusCodeColorLevelPrior, True
        LockMotor 0
        Exit Function
           
    End If
    
    
    'Reset both X and Y motor controllers
    ' This will reset both motor positions to 0 motor encoder counts, but this is okay
    ' as the homing process will rezero the motor positions relative to the
    ' center X and Y optical limit switches
    MotorReset MotorChanger
    MotorReset MotorChangerY
        
    'Wait 1 second for motor power cycle process to finish
    DelayTime 1
    
    start_time = Timer
    
    'T Shuma - Now Move to center limit switches
    LockMotor MotorChanger
    ' Move motor a relative number of motor units, and stop if signal from the center limit switch is logic low
    MoveMotorXY MotorChanger, MotorPositionMoveToCenter, ChangerSpeed, False, pauseOveride, -2, 0
    
    DelayTime 0.1
    
    LockMotor MotorChangerY
    MoveMotorXY MotorChangerY, MotorPositionMoveToCenter, ChangerSpeed, False, pauseOveride, -3, 0
    
    'start_time = Timer
    
    Do While ((CheckInternalStatus(MotorChanger, 5) <> 0) Or (CheckInternalStatus(MotorChangerY, 6) <> 0)) _
             And Not NOCOMM_MODE _
             And Not Prog_halted _
             And Not HasMoveToXYLimit_Timedout(start_time)
             
            
        DelayTime 0.1
    
    Loop
   
    If Not (CheckInternalStatus(MotorChanger, 5) = 0 And CheckInternalStatus(MotorChangerY, 6) = 0) And Not NOCOMM_MODE Then
       
        SetCodeLevel CodeRed
       
        If (HasMoveToXYLimit_Timedout(start_time)) Then
        
           'Home to center has timed out
           frmSendMail.MailNotification "Homing error!", "Home XY Stage, move motors to Center limit switches timed-out after: " & _
                                                         str$(MoveXYMotorsToLimitSwitch_TimeoutSeconds) & _
                                                         "seconds.", CodeRed
            
           MsgBox "Home XY Stage, move motors to Center limit switches timed-out after: " & _
                  str$(MoveXYMotorsToLimitSwitch_TimeoutSeconds) & _
                  "seconds."
          
        Else
       
            frmSendMail.MailNotification "Homing error!", "Homed XY Stage to center but did not hit center limit switch(es).", CodeRed
       
            MsgBox "Homed XY Stage to center but did not hit positive limit switch(es)!"
       
        End If
        
        SetCodeLevel StatusCodeColorLevelPrior, True
        
    Else
    
        ' Zero X and Y motors so that 0 motor position corresponds to the X and Y
        ' center limit switch positions
        ZeroTargetPos MotorChanger
        DelayTime 0.25
        ZeroTargetPos MotorChangerY
        DelayTime 0.25
           
        xpos = ReadPosition(MotorChanger)
        ypos = ReadPosition(MotorChangerY)
        
    '    frmSettings.XYHolePositionsFlexGrid.row = 1
    '    frmSettings.XYHolePositionsFlexGrid.Col = 1
    '    frmSettings.XYHolePositionsFlexGrid.text = xPos
    '    modConfig.XYTablePositions(frmSettings.XYHolePositionsFlexGrid.row - 1, 0) = xPos
    '    frmSettings.XYHolePositionsFlexGrid.Col = 2
    '    frmSettings.XYHolePositionsFlexGrid.text = yPos
    '    modConfig.XYTablePositions(frmSettings.XYHolePositionsFlexGrid.row - 1, 1) = yPos
        
        modConfig.HasXYTableBeenHomed = True
        
    'frmSettings.RecalculateCupPositions xPos, yPos
    End If

    LockMotor 0
           
End Function

Private Function HasMoveToXYLimit_Timedout(ByVal StartTime As Double) As Boolean

    Dim current_time, start_time

    start_time = StartTime
    current_time = Timer

    If current_time < start_time Then start_time = start_time - 86400

    If current_time < start_time + MoveXYMotorsToLimitSwitch_TimeoutSeconds Then
    
        HasMoveToXYLimit_Timedout = False
        
    Else
    
        HasMoveToXYLimit_Timedout = True
    
    End If

End Function

Public Function HomeToCenter_NoLoadCorner(Optional pauseOveride As Boolean = False)
    
    'No homing to center if the program has been halted
    If Prog_halted Then Exit Function
   
    Dim start_time
   
    Dim stop_state As Integer
    If modConfig.DCMotorHomeToTop_StopOnTrue Then
        stop_state = 1
    Else
        stop_state = 0
    End If
   
    'No homing to center if the Up/Down Motor is not homed
    If CheckInternalStatus(MotorUpDown, 4) <> stop_state Then
         SetCodeLevel CodeRed
        frmSendMail.MailNotification "Homing error!", "Tried to Home to center but UP/Down Motor not homed to top.  This could break control rod.", CodeRed
       
        MsgBox "Could not home to center!  Home to top not complete!"
       
        SetCodeLevel StatusCodeColorLevelPrior, True
        Exit Function
    End If
   
    'Move changer to Pre-Center Hole, wait for movement to stop
    Me.ChangerMotortoHole HomeToCenter_NoLoadCorner_PreHomeHole
    
    DelayTime 0.1
   
    start_time = Timer '7/18/23
   
    LockMotor MotorChanger
    MoveMotorXY MotorChanger, MotorPositionMoveToCenter, ChangerSpeed, False, pauseOveride, -2, 0 '7/18/23 sign changed
    
    LockMotor MotorChangerY
    MoveMotorXY MotorChangerY, MotorPositionMoveToCenter, ChangerSpeed, False, pauseOveride, -3, 0 '7/18/23 sign changed
       
    'T Shuma - Now Move to center limit switches
    Do While ((CheckInternalStatus(MotorChanger, 5) <> 0) Or _
              (CheckInternalStatus(MotorChangerY, 6) <> 0)) _
             And Not NOCOMM_MODE _
             And Not Prog_halted _
             And Not HasMoveToXYLimit_Timedout(start_time) '7/18/23
             
            
        DelayTime 0.1
    
    Loop
    
    If Not (CheckInternalStatus(MotorChanger, 5) = 0 And CheckInternalStatus(MotorChangerY, 6) = 0) And Not NOCOMM_MODE Then
       
        SetCodeLevel CodeRed
       
        If (HasMoveToXYLimit_Timedout(start_time)) Then
        
           'Home to center has timed out
           frmSendMail.MailNotification "Homing error!", "Home XY Stage, move motors to Center limit switches timed-out after: " & _
                                                         str$(MoveXYMotorsToLimitSwitch_TimeoutSeconds) & _
                                                         "seconds.", CodeRed
            
           MsgBox "Home XY Stage, move motors to Center limit switches timed-out after: " & _
                  str$(MoveXYMotorsToLimitSwitch_TimeoutSeconds) & _
                  "seconds."
          
        Else
       
           frmSendMail.MailNotification "Homing error!", "Homed XY Stage to center but did not hit center limit switch(es).", CodeRed
        
           MsgBox "Homed XY Stage to center but did not hit center limit switch(es)!"
       
        End If
        
        SetCodeLevel StatusCodeColorLevelPrior, True
       
    Else
    
        ZeroTargetPos MotorChanger
        ZeroTargetPos MotorChangerY
        
    End If
    
    LockMotor 0

End Function

Private Sub HomeToCenterButton_Click()
Dim xpos As Long
Dim ypos As Long
    HomeToTop
    HomeToCenter xpos, ypos, pauseOveride:=True
    'modConfig.XYTablePositions(0, 0) = xPos
    'modConfig.XYTablePositions(0, 1) = yPos
End Sub

'*************************************************************************************************************
'
'   Home To Top Code rewrite
'   Modification: 12/3/2010
'         Author: Isaac Hilburn
'
'         Reason: To enable the new DC servo motors from Quicksilver to home to top
'                 correctly.  The old motors somehow did not need to have the stop-condition
'                 coded in directly in the Motor Move command sent to the Up/Down motor.
'                 (Stop Condition = if logic bit #1 becomes true, stop all motion on Up/Down motor;
'                  When the home position limit switch is pressed, logic bit #1 changes from false to true.)
'
'*************************************************************************************************************

'*****************  Old Code  *******************

'Public Function HomeToTop(Optional pauseOveride As Boolean = False) As Long
'    Dim offset As Integer
'    Dim speed As Long
'    If Prog_halted Then Exit Function
'    ' if switch already tripped, do nothing
'    If CheckInternalStatus(MotorUpDown, 4) = 1 Then Exit Function
'    Do While ((MotorsLocked <> 0) And (MotorsLocked <> MotorUpDown))
'        DelayTime 0.05
'    Loop
'    LockMotor MotorUpDown
'    If Abs(ReadPosition(MotorUpDown)) > Abs(SampleBottom) Then
'        speed = LiftSpeedNormal
'    Else
'        speed = 0.25 * (LiftSpeedNormal + 3 * LiftSpeedSlow)
'    End If
'    MoveMotor MotorUpDown, -2 * MeasPos, speed, True, pauseOveride
'    If Not CheckInternalStatus(MotorUpDown, 4) = 1 Then MoveMotor MotorUpDown, -2 * MeasPos, LiftSpeedSlow, True, pauseOveride
'    If Not CheckInternalStatus(MotorUpDown, 4) = 1 And Not NOCOMM_MODE Then
'        SetCodeLevel CodeRed
'        frmSendMail.MailNotification "Homing error!", "Homed to top but did not hit switch.", CodeRed
'        MsgBox "Homed to top but did not hit switch!"
'        SetCodeLevel StatusCodeColorLevelPrior, True
'    End If
'    HomeToTop = ReadPosition(MotorUpDown)
'    ZeroTargetPos MotorUpDown ' (September 2007) previously MotorReset MotorUpDown
'    LockMotor 0
'End Function

'*****************  New Code  *******************

Public Function HomeToTop(Optional pauseOveride As Boolean = False) As Long
    Dim offset As Integer
    Dim speed As Long
   
    'No homing to top if the program has been halted
    If Prog_halted Then Exit Function
   
    Dim stop_state As Integer
    If modConfig.DCMotorHomeToTop_StopOnTrue Then
        stop_state = 1
    Else
        stop_state = 0
    End If
   
    ' if switch already tripped, do nothing
    If CheckInternalStatus(MotorUpDown, 4) = stop_state Then
    
        Dim current_updown_pos As Long
        current_updown_pos = GetUpDownPos()
        
        If Abs(current_updown_pos) > (modConfig.UpDownMotor1cm / 10) Then
        
            ZeroTargetPos MotorUpDown
            
        End If
        Exit Function
        
    End If
   
    Do While ((MotorsLocked <> 0) And (MotorsLocked <> MotorUpDown))
       
        DelayTime 0.1
   
    Loop
   
    'Lock the up/down motor so that no other commands can be sent to it???
    LockMotor MotorUpDown
   
    'If up/down position is greater than the sample bottom, use the normal lift speed
    'otherwise, a sample has just been dropped off on the changer belt
    'and the home to top speed needs to be slower
    If Abs(ReadPosition(MotorUpDown)) > Abs(SampleBottom) Then
       
        speed = LiftSpeedNormal
   
    Else
       
        '???????????? - Why this particular speed?
        speed = 0.25 * (LiftSpeedNormal + 3 * LiftSpeedSlow)
   
    End If
   
    
    
    MoveMotor MotorUpDown, -2 * MeasPos, speed, True, pauseOveride, -1, stop_state
   
    'Check to see if the motor has reached the limit switch
    If Not CheckInternalStatus(MotorUpDown, 4) = stop_state Then
   
        MoveMotor MotorUpDown, _
                  -2 * MeasPos, _
                  LiftSpeedSlow, _
                  True, _
                  pauseOveride
       
    End If
   
    If Not CheckInternalStatus(MotorUpDown, 4) = stop_state And Not NOCOMM_MODE Then
       
        SetCodeLevel CodeRed
        frmSendMail.MailNotification "Homing error!", "Homed to top but did not hit switch.", CodeRed
       
        MsgBox "Homed to top but did not hit switch!"
       
        SetCodeLevel StatusCodeColorLevelPrior, True
       
    End If
   
    HomeToTop = ReadPosition(MotorUpDown)
   
    ZeroTargetPos MotorUpDown ' (September 2007) previously MotorReset MotorUpDown
   
    LockMotor 0
   
End Function

Private Sub HomeToTopButton_Click()
    HomeToTop pauseOveride:=True
End Sub

Private Function IsValidMotorid(motorid As Integer)
    IsValidMotorid = False
    If motorid = MotorTurning Or motorid = MotorChanger Or motorid = MotorChangerY Or motorid = MotorUpDown Then
        IsValidMotorid = True
    End If
End Function

Private Sub LoadButton_Click()
' Move to the corner
    frmDCMotors.MoveToCorner pauseOveride:=True
End Sub

Public Sub LockMotor(motorid As Integer)
    If IsValidMotorid(motorid) Then
        MotorsLocked = motorid
    Else
        MotorsLocked = 0
    End If
End Sub

Private Function MotorAddress(Optional motorid As Integer) As String
    Select Case motorid
    Case MotorChanger
        MotorAddress = MotorIDChanger
    Case MotorChangerY
        MotorAddress = MotorIDChangerY
    Case MotorTurning
        MotorAddress = MotorIDTurning
    Case MotorUpDown
        MotorAddress = MotorIDUpDown
    Case Else
        MotorAddress = MotorIDTurning
    End Select
    MotorAddress = "@" & LTrim$(str(MotorAddress)) & " "
End Function

Public Sub MotorCommConnect(motorid As Integer)
    If MSCommMotor(ComPortAssignments(motorid)).PortOpen = False And Not NOCOMM_MODE Then
        On Error GoTo ErrorHandler  ' Enable error-handling routine.
        MSCommMotor(ComPortAssignments(motorid)).CommPort = MotorCOMPort(motorid)
        MSCommMotor(ComPortAssignments(motorid)).Settings = "57600,N,8,2"
        MSCommMotor(ComPortAssignments(motorid)).SThreshold = 1
        MSCommMotor(ComPortAssignments(motorid)).RThreshold = 0
        MSCommMotor(ComPortAssignments(motorid)).inputlen = 1
        MSCommMotor(ComPortAssignments(motorid)).PortOpen = True
        On Error GoTo 0 ' Turn off error trapping.
        If MSCommMotor(ComPortAssignments(motorid)).PortOpen = True Then
            chkConnectMotor(motorid).value = Checked
            ' Set an Ack Delay of 50 msec
            SendCommand motorid, ("@255 173 416") '50 msec delay
            If motorid = MotorUpDown Then SetTorques MotorUpDown, UpDownTorqueFactor, UpDownTorqueFactor, UpDownTorqueFactor, UpDownTorqueFactor
            ReadPosition (motorid)
            optMotorActive(motorid).Enabled = True
        Else
            chkConnectMotor(motorid).value = Unchecked
            optMotorActive(motorid).Enabled = False
            If Not NOCOMM_MODE Then MsgBox "Motor Comm port not open in connection routine."
        End If
    End If
Exit Sub        ' Exit to avoid handler.
ErrorHandler:   ' Error-handling routine.
    Select Case Err.number  ' Evaluate error number.
        Case 8002
            MsgBox "Invalid Port Number"
        Case 8005
            MsgBox "Port already open" + Chr(13) + "(Already is use?)"
        Case 8010
            MsgBox "The hardware is not available (locked by another device)"
        Case 8012
            MsgBox "The device is not open"
        Case 8013
            MsgBox "The device is already open"
        Case Else
            MsgBox "Unknown error trying to Connect Comm Port"
    End Select
    
     'Prompt the user to see if they would like to turn on NOCOMM mode
    Prompt_NOCOMM
    
End Sub

Public Sub MotorCommDisconnect(Optional ByVal motorid As Integer = 0)
    If Not (motorid = MotorTurning Or motorid = MotorChanger Or motorid = MotorChangerY Or motorid = MotorUpDown) Then
        MotorCommDisconnect MotorTurning
        MotorCommDisconnect MotorChanger
        MotorCommDisconnect MotorChangerY
        MotorCommDisconnect MotorUpDown
        AssignCommPorts
        Exit Sub
    End If
    If MSCommMotor(ComPortAssignments(motorid)).PortOpen Then
        MSCommMotor(ComPortAssignments(motorid)).InBufferCount = 0
        MSCommMotor(ComPortAssignments(motorid)).OutBufferCount = 0
        MSCommMotor(ComPortAssignments(motorid)).PortOpen = False
        chkConnectMotor(ComPortAssignments(motorid)).value = Unchecked
        optMotorActive(ComPortAssignments(motorid)).Enabled = False
    End If
End Sub

Private Function MotorCOMPort(Optional motorid As Integer = MotorTurning) As Integer
    Select Case motorid
    Case MotorChanger
        MotorCOMPort = COMPortChanger
    Case MotorChangerY
        MotorCOMPort = COMPortChangerY
    Case MotorTurning
        MotorCOMPort = COMPortTurning
    Case MotorUpDown
        MotorCOMPort = COMPortUpDown
    Case Else
        MotorCOMPort = COMPortTurning
    End Select
End Function

Public Sub MotorHalt(Optional motorid As Integer = 0)
    If motorid = 0 Then
        LockMotor 0
        MotorHalt MotorUpDown
        MotorHalt MotorChanger
        MotorHalt MotorChangerY
        MotorHalt MotorTurning
        Exit Sub
    End If
    LockMotor 0
    SendCommand motorid, MotorAddress(motorid) & "2"
    GetResponse (motorid)
End Sub

Private Sub MotorHaltButton_Click()
    MotorHalt ActiveMotorControls
End Sub

Public Sub MotorRestart(Optional motorid As Integer = 0)
    If motorid = 0 Then
        LockMotor 0
        MotorRestart MotorUpDown
        MotorRestart MotorChanger
        MotorRestart MotorChangerY
        MotorRestart MotorTurning
        Exit Sub
    End If
    
    LockMotor 0
    SendCommand motorid, MotorAddress(motorid) & "255"
    GetResponse (motorid)
    DelayTime 0.2
End Sub

Public Sub MotorReset(Optional motorid As Integer = 0)
    If motorid = 0 Then
        LockMotor 0
        MotorReset MotorUpDown
        MotorReset MotorChanger
        MotorReset MotorChangerY
        MotorReset MotorTurning
        Exit Sub
    End If
    SendCommand motorid, MotorAddress(motorid) & "4"
    GetResponse (motorid)
End Sub

Private Sub MotorResetButton_Click()
    MotorReset ActiveMotorControls
End Sub

Public Sub MotorStop(Optional motorid As Integer = 0)
    If motorid = 0 Then
        LockMotor 0
        MotorStop MotorUpDown
        MotorStop MotorChanger
        MotorStop MotorChangerY
        MotorStop MotorTurning
        Exit Sub
    End If
    LockMotor 0
    SendCommand motorid, MotorAddress(motorid) & "3 0"
    GetResponse (motorid)
End Sub

Private Sub MotorStopButton_Click()
    MotorStop ActiveMotorControls
End Sub

                  
'***********  Old Code  **************

'Private Sub MoveMotor(motorid As Integer, MoveMotorPos As Long, MoveMotorVelocity As Long, Optional waitingForStop As Boolean = True, Optional ByVal pauseOverride As Boolean = False)
'    Do While ((MotorsLocked <> 0) And (MotorsLocked <> motorid))
'        DelayTime 0.1
'    Loop
'    If DEBUG_MODE Then frmDebug.Msg "Motor " & Str$(motorid) & " to " & Str$(MoveMotorPos) & " at " & Str$(MoveMotorVelocity)
'    LockMotor motorid
'    MoveMotorPosEdit = Str(MoveMotorPos)
'    MoveMotorVelocityEdit = Str(MoveMotorVelocity)
'    PollMotor motorid
'    ClearPollStatus motorid
'    SendCommand motorid, (MotorAddress(motorid) & "134 " + MoveMotorPosEdit + " 96637 " + MoveMotorVelocityEdit + " 0 0")
'    GetResponse motorid
'    If waitingForStop Then WaitForMotorStop motorid, pauseOverride
'    LockMotor 0
'End Sub

'************  New Code  *************

Public Sub MoveMotor(motorid As Integer, _
                      MoveMotorPos As Long, _
                      MoveMotorVelocity As Long, _
                      Optional waitingForStop As Boolean = True, _
                      Optional ByVal pauseOverride As Boolean = False, _
                      Optional ByVal StopEnable As Integer = 0, _
                      Optional ByVal StopCondition As Integer = 0)
    Do While ((MotorsLocked <> 0) And (MotorsLocked <> motorid))
        DelayTime 0.1
    Loop
    If DEBUG_MODE Then frmDebug.msg "Motor " & str$(motorid) & " to " & str$(MoveMotorPos) & " at " & str$(MoveMotorVelocity)
    LockMotor motorid
    MoveMotorPosEdit = str(MoveMotorPos)
    MoveMotorVelocityEdit = str(MoveMotorVelocity)
    PollMotor motorid
    ClearPollStatus motorid
    '96637   @16 135 -100000 483184 536869570 0 0
    If UseXYTableAPS And motorid = MotorUpDown Then
    SendCommand motorid, _
               (MotorAddress(motorid) & "134 " + _
                MoveMotorPosEdit + str$(modConfig.LiftAcceleration) + _
                MoveMotorVelocityEdit + " " + _
                Trim(str(StopEnable)) + " " + _
                Trim(str(StopCondition)))
    Else
        SendCommand motorid, _
               (MotorAddress(motorid) & "134 " + _
                MoveMotorPosEdit + " 96637 " + _
                MoveMotorVelocityEdit + " " + _
                Trim(str(StopEnable)) + " " + _
                Trim(str(StopCondition)))
    End If
                

    GetResponse motorid
    If waitingForStop Then WaitForMotorStop motorid, pauseOverride
    LockMotor 0
End Sub

Public Sub MoveMotorAbsoluteXY(ByVal motorid As Integer, _
                      ByVal MoveMotorPos As Long, _
                      ByVal MoveMotorVelocity As Long, _
                      Optional waitingForStop As Boolean = True, _
                      Optional ByVal pauseOverride As Boolean = False, _
                      Optional ByVal StopEnable As Integer = 0, _
                      Optional ByVal StopCondition As Integer = 0)
                      
    'Verify that the up/down motor is homed to top
    Dim stop_state As Integer
    If modConfig.DCMotorHomeToTop_StopOnTrue Then
        stop_state = 1
    Else
        stop_state = 0
    End If
    
    'Home up/down motor to top
    HomeToTop pauseOveride:=True
    
    'No homing to center if the Up/Down Motor is not homed
    If CheckInternalStatus(MotorUpDown, 4) <> stop_state Then
         SetCodeLevel CodeRed
        frmSendMail.MailNotification "Homing error!", _
                                     "Tried to move to X,Y motors, but UP/Down Motor not homed to top.  " & _
                                     "This could break control rod.", CodeRed
       
        MsgBox "Could not move X,Y motors!  Home to top not complete!"
       
        SetCodeLevel StatusCodeColorLevelPrior, True
        Exit Sub
    End If
                     
    LockMotor motorid
                      
    Do While ((MotorsLocked <> 0) And (MotorsLocked <> motorid))
        DelayTime 0.1
    Loop
    If DEBUG_MODE Then frmDebug.msg "Motor " & str$(motorid) & " to " & str$(MoveMotorPos) & " at " & str$(MoveMotorVelocity)
    
    MoveMotorPosEdit = str(MoveMotorPos)
    MoveMotorVelocityEdit = str(MoveMotorVelocity)
    PollMotor motorid
    ClearPollStatus motorid
    '96637   @16 135 -100000 483184 536869570 0 0
    SendCommand motorid, _
               (MotorAddress(motorid) & "134 " + _
                MoveMotorPosEdit + " 483184 " + _
                MoveMotorVelocityEdit + " " + _
                Trim(str(StopEnable)) + " " + _
                Trim(str(StopCondition)))

    GetResponse motorid
    If waitingForStop Then WaitForMotorStop motorid, pauseOverride
    LockMotor 0
End Sub

Private Sub MoveMotorButton_Click()
    MoveMotor ActiveMotorControls, val(MoveMotorPosEdit), val(MoveMotorVelocityEdit), pauseOverride:=True
End Sub

Public Sub MoveMotorXY(motorid As Integer, _
                      MoveMotorPos As Long, _
                      MoveMotorVelocity As Long, _
                      Optional waitingForStop As Boolean = True, _
                      Optional ByVal pauseOverride As Boolean = False, _
                      Optional ByVal StopEnable As Integer = 0, _
                      Optional ByVal StopCondition As Integer = 0)
                  
    'If Settings form and XY Motors tab are active, and Override Home to Top is clicked,
    'then only need to check position of up/down tube (greater than or equal to sample bottom)
    If frmSettings.Visible = True And _
       frmSettings.frameOptions(2).Visible = True And _
       frmSettings.chkOverrideHomeToTop_ForMoveMotorAbsoluteXY.value = Checked Then
       
        Dim up_down_position As Long
        up_down_position = ReadPosition(MotorUpDown)
        
        If Abs(up_down_position) >= Abs(modConfig.SampleBottom) + 50 Then
        
            UpDownMove modConfig.SampleTop, 0
            
        End If
        
        up_down_position = ReadPosition(MotorUpDown)
        
        If Abs(up_down_position) >= Abs(modConfig.SampleBottom) + 50 Then
        
             SetCodeLevel CodeRed
             frmSendMail.MailNotification "Up/Down Motor Error!", _
                                          "Tried to move to X,Y motors, but UP/Down Motor " & _
                                          "is in the way and will not respond to a move motor command.  This could break control rod.", _
                                          CodeRed
            
             MsgBox "Tried to move to X,Y motors, but UP/Down Motor is in the way and will not respond to a move motor command."
            
             SetCodeLevel StatusCodeColorLevelPrior, True
             Exit Sub
             
        End If
    
    Else
    
        'Verify that the up/down motor is homed to top
        Dim stop_state As Integer
        If modConfig.DCMotorHomeToTop_StopOnTrue Then
            stop_state = 1
        Else
            stop_state = 0
        End If
        
        'Home up/down motor to top
        LockMotor MotorUpDown
        HomeToTop pauseOveride:=True
        
        'No homing to center if the Up/Down Motor is not homed
        If CheckInternalStatus(MotorUpDown, 4) <> stop_state Then
             SetCodeLevel CodeRed
            frmSendMail.MailNotification "Homing error!", _
                                         "Tried to move to X,Y motors, but UP/Down Motor not homed to top.  " & _
                                         "This could break control rod.", CodeRed
           
            MsgBox "Could not move X,Y motors!  Home to top not complete!"
           
            SetCodeLevel StatusCodeColorLevelPrior, True
            Exit Sub
        End If
    
    End If
                          
    'Do While ((MotorsLocked <> 0) And (MotorsLocked <> motorid))
    '    DelayTime 0.1
    'Loop
    If DEBUG_MODE Then frmDebug.msg "Motor " & str$(motorid) & " to " & str$(MoveMotorPos) & " at " & str$(MoveMotorVelocity)
    LockMotor motorid
    MoveMotorPosEdit = str(MoveMotorPos)
    MoveMotorVelocityEdit = str(MoveMotorVelocity)
    
    PollMotor motorid
    ClearPollStatus motorid
    '96637   @16 135 -100000 3000 536869570 0 0
    SendCommand motorid, _
               (MotorAddress(motorid) + " 135 " + _
                MoveMotorPosEdit + " 3000 " + _
                Trim(str(MoveMotorVelocityEdit)) + " " + _
                Trim(str(StopEnable)) + " " + _
                Trim(str(StopCondition)))

    GetResponse motorid
    If waitingForStop Then WaitForMotorStop motorid, pauseOverride
    LockMotor 0
End Sub

'*************************************************************************************************************
'
'   Home To Center  7/27/2012
'         Author: Toni Shuma
'
'         Purpose: To enable the Quicksilver motors to return to the home position, center of the XY Table
'                 correctly.  Both X and Y axis must find the center of the table and stop there correctly
'                 (Stop Condition = if logic bit #1 becomes true, stop all motion on Up/Down motor;
'                 When the home position limit switch is pressed, logic bit #1 changes from false to true.)
'
'                   Note: untested.  Requires planning of limit switches to fix code and verify code works properly
'
'*************************************************************************************************************


Public Function MoveToCorner(Optional pauseOveride As Boolean = False)
    
    'No homing to center if the program has been halted
    If Prog_halted Then Exit Function
   
    HomeToTop pauseOveride:=True
    
    Dim stop_state As Integer
    If modConfig.DCMotorHomeToTop_StopOnTrue Then
        stop_state = 1
    Else
        stop_state = 0
    End If
    
    'No homing to center if the Up/Down Motor is not homed
    If CheckInternalStatus(MotorUpDown, 4) <> stop_state Then
         SetCodeLevel CodeRed
        frmSendMail.MailNotification "Homing error!", "Tried to move to load corner but UP/Down Motor not homed to top.  This could break control rod.", CodeRed
       
        MsgBox "Could not move to corner!  Home to top not complete!"
       
        SetCodeLevel StatusCodeColorLevelPrior, True
        Exit Function
    End If

    Dim start_time As Double
    
    'start_time = Timer

    LockMotor MotorChanger
    MoveMotorXY MotorChanger, MotorPositionMoveToLoadCorner, ChangerSpeed, False, pauseOveride, -1, 0
   
    LockMotor MotorChangerY
    MoveMotorXY MotorChangerY, MotorPositionMoveToLoadCorner, ChangerSpeed, False, pauseOveride, -2, 0
   
    'T Shuma - Move to load corner
    Do While ((CheckInternalStatus(MotorChanger, 4) <> 0) Or (CheckInternalStatus(MotorChangerY, 5) <> 0)) _
             And Not NOCOMM_MODE _
             And Not Prog_halted
             'And Not HasMoveToXYLimit_Timedout(start_time)
             
        DelayTime 0.1
    Loop
    
    If Not (CheckInternalStatus(MotorChanger, 4) = 0 And CheckInternalStatus(MotorChangerY, 5) = 0) And Not NOCOMM_MODE Then
       
        SetCodeLevel CodeRed
       
        If (HasMoveToXYLimit_Timedout(start_time)) Then
        
           'Home to center has timed out
           frmSendMail.MailNotification "Homing error!", "Move XY Stage to Load Corner timed-out after: " & _
                                                         str$(MoveXYMotorsToLimitSwitch_TimeoutSeconds) & _
                                                         "seconds.", CodeRed
            
           MsgBox "Move XY Stage to Load Corner timed-out after: " & _
                  str$(MoveXYMotorsToLimitSwitch_TimeoutSeconds) & _
                  "seconds."
          
        Else
       
            frmSendMail.MailNotification "Homing error!", "Moved XY Stage to Load Corner but did not hit load corner limit switch(es).", CodeRed
       
            MsgBox "Moved XY Stage to Load Corner but did not hit load corner limit switch(es)!"
       
        End If
        
        SetCodeLevel StatusCodeColorLevelPrior, True
        
    End If
    
    LockMotor 0

End Function
'******************************************************************************************
'
'   Part of Home To Top code re-write
'   Modification: 12/3/2010
'         Author: Isaac Hilburn
'
'         Reason: To use the limit switch digital input on the Quicksilver motor box
'                 to stop the motor movement (as opposed to waiting for the motor to
'                 encounter too much resistance to stop itself).
'                 Also, this code is necessary for the new Quicksilver Up/Down motor to home to top
'******************************************************************************************

Private Function PollMotor(motorid As Integer) As String
    'Poll the motor once
    SendCommand motorid, (MotorAddress & "0")
    GetResponse motorid
End Function

Private Sub PollMotorButton_Click()
    PollMotor ActiveMotorControls
End Sub

Public Function PromptUser_DoHomeXYStage() As VbMsgBoxResult

    PromptUser_DoHomeXYStage = _
        MsgBox("The XY stage needs to be homed to center, now." & vbCrLf & vbCrLf & _
               "The code will home the Up/Down glass tube to the top limit " & _
               "switch before moving the XY Stage.  HOWEVER, if their are " & _
               "cables or other impediments in the way, the XY stage should NOT be homed." & _
               vbCrLf & vbCrLf & _
               "Do you want to home the XY stage to the center position, now?", vbYesNo, _
               "Warning: XY Stage Homing!")
               
End Function

Private Sub ReadAngleButton_Click()
    Dim dummy As Double
    dummy = TurningMotorAngle()
End Sub

Private Sub ReadHoleButton_Click()
    Dim dummy As Double
    dummy = ChangerHole()
End Sub

Private Sub ReadPosButton_Click()
    Dim dummy As Long
    dummy = ReadPosition(ActiveMotorControls)
End Sub

Public Function ReadPosition(motorid As Integer) As Long
    Dim hexpos As String
    Dim decpos As Long
    ' If we have an error, we'll report the last position read
    decpos = lastPosition(motorid)
    SendCommand motorid, MotorAddress(motorid) & "12 1"
    GetResponse motorid
    If Len(InputText(motorid)) = 20 Then
        hexpos = Mid(InputText(motorid), 11, 4) + Mid(InputText(motorid), 16, 4)
        decpos = CLng("&H" + hexpos)
    End If
    txtPollPosition = decpos
    lastPosition(motorid) = decpos
    ReadPosition = decpos
End Function

Public Sub RelabelPos(motorid As Integer, pos As Long)
    Do Until Abs(ReadPosition(motorid) - pos) < 10
        ZeroTargetPos motorid
        ' @16 11 10 num 'load register
        SendCommand motorid, (MotorAddress & "11 10 " & str(-pos))
        GetResponse motorid
        ' @16 165 1802 ' subtract register 10 from T&P
        SendCommand motorid, (MotorAddress & "165 1802")
        GetResponse motorid
    Loop
End Sub

Private Sub RelabelPosButton_Click()
    RelabelPos ActiveMotorControls, (val(MoveMotorPosEdit))
End Sub

Public Sub ResumeMove()
    Select Case lastMoveCommand
    Case lastCmdPickup
        SamplePickup
    Case lastCmdDropoff
        SampleDropOff
    Case lastCmdMove
        Select Case lastMoveMotor
        Case MotorChanger
            ChangerMotortoHole lastMoveTarget
        Case MotorChangerY
            ChangerMotortoHole lastMoveTarget
        Case MotorTurning
            TurningMotorRotate lastMoveTarget
        Case MotorUpDown
            UpDownMove lastMoveTarget, 0
        End Select
    End Select
End Sub

Public Sub SampleDropOff(Optional pauseOverride As Boolean = False)
    If Prog_halted Then Exit Sub
    If Not Prog_paused Or Prog_halted Then
        lastMoveCommand = lastCmdDropoff
        lastMoveMotor = MotorUpDown
    End If
    PollMotor MotorUpDown
    ClearPollStatus MotorUpDown
    
    If UseXYTableAPS Then
        MoveMotor MotorUpDown, SampleBottom + (SampleHeight - 0.1 * SampleHeight), LiftSpeedSlow, True, pauseOverride
    Else
        MoveMotor MotorUpDown, SampleBottom + 1.1 * SampleHeight, LiftSpeedSlow, True, pauseOverride
    End If
    
    Dim stop_state As Integer
    If modConfig.DCMotorHomeToTop_StopOnTrue Then
        stop_state = 1
    Else
        stop_state = 0
    End If
    
    ' (September 2007) check to see if the up/down switch is stuck
    If CheckInternalStatus(MotorUpDown, 4) = stop_state And Not NOCOMM_MODE Then
        SetCodeLevel CodeRed
        frmSendMail.MailNotification "Switch Failure!", "Dropped off sample, but homing switch still set. Check for switch failure.", CodeRed
        MsgBox "Dropped off sample, but homing switch still set. Check for switch failure."
        SetCodeLevel StatusCodeColorLevelPrior, True
    End If
End Sub

Public Sub SamplePickup(Optional pauseOverride As Boolean = False)
    Dim currentPos As Long
    If Prog_halted Then Exit Sub
    If Not Prog_paused Or Prog_halted Then
        lastMoveCommand = lastCmdPickup
        lastMoveMotor = MotorUpDown
    End If
    LockMotor MotorUpDown
    SetTorques MotorUpDown, PickupTorqueThrottle * UpDownTorqueFactor, PickupTorqueThrottle * UpDownTorqueFactor, PickupTorqueThrottle * UpDownTorqueFactor, PickupTorqueThrottle * UpDownTorqueFactor
    MoveMotor MotorUpDown, SampleBottom, LiftSpeedSlow, True, pauseOverride
    currentPos = UpDownHeight
    
    ZeroTargetPos MotorUpDown

    RelabelPos MotorUpDown, currentPos
    SetTorques MotorUpDown, UpDownTorqueFactor, UpDownTorqueFactor, UpDownTorqueFactor, UpDownTorqueFactor
    'MoveMotor MotorUpDown, currentPos + (currentPos - SampleBottom) * 0.3, LiftSpeedSlow, False, pauseOverride
        ' (September 2007) check to see if the up/down switch is stuck
        
    Dim stop_state As Integer
    If modConfig.DCMotorHomeToTop_StopOnTrue Then
        stop_state = 1
    Else
        stop_state = 0
    End If
        
    If CheckInternalStatus(MotorUpDown, 4) = stop_state And Not NOCOMM_MODE Then
        SetCodeLevel CodeRed
        frmSendMail.MailNotification "Switch Failure!", "Quartz tube at sample top but homing switch still set. Check for switch failure.", CodeRed
        MsgBox "Quartz tube at sample top but homing switch still set. Check for switch failure."
        SetCodeLevel StatusCodeColorLevelPrior, True
    End If
    
    LockMotor 0
End Sub

Private Sub SamplePickupButton_Click()
    SamplePickup pauseOverride:=True
End Sub

Private Sub SendCommand(motorid As Integer, outstring As String)

    On Error GoTo CommError

    If Not IsValidMotorid(motorid) Then Exit Sub
    optMotorActive(motorid).value = True
    If MSCommMotor(motorid).PortOpen = False Then MotorCommConnect motorid
    If OverlappingComPorts Then
        Do While ((MotorsLocked <> 0) And (MotorsLocked <> motorid))
            DelayTime 0.1
        Loop
        LockMotor motorid
    End If
    If MSCommMotor(motorid).PortOpen = True Then
    
        MSCommMotor(motorid).RTSEnable = True
        MSCommMotor(motorid).InBufferCount = 0
        MSCommMotor(motorid).OutBufferCount = 0
        'outstring = "@16 135 -100000 483184 536869570 0 0 "
        MSCommMotor(motorid).Output = outstring + vbCrLf
        OutputText = outstring
        
        If modConfig.LogMessages Then
            modListenAndLog.AppendRS232MessageToLogFile outstring, MSCommMotor(motorid).CommPort, "Output"
        End If
        
    Else
        If Not NOCOMM_MODE Then MsgBox "Motor Comm Port Not Open sending " + outstring + " to comm port " + str(MSCommMotor(motorid).CommPort)
    End If
    If OverlappingComPorts Then LockMotor 0
    
    On Error GoTo 0
    Exit Sub
    
CommError:

    MotorCommDisconnect motorid
    MotorCommConnect motorid
    SendCommand motorid, outstring
    
End Sub

Private Sub SetAfButton_click()
    Dim dummystring As String
    ' Set the value of the place to move to the center of the Af coils
    MoveMotorPosEdit.text = str(AFPos)
    txtNewHeight = str(AFPos)
End Sub

Public Sub SetChangerHole(hole As Double)
    If Changer_ValidStart(hole) Then
    If Not UseXYTableAPS Then
        RelabelPos MotorChanger, (ConvertHoletoPos(hole))
        Else
        RelabelPos MotorChanger, (ConvertHoletoPosX(hole))
        RelabelPos MotorChangerY, (ConvertHoletoPosY(hole))
        End If
        currenthole = hole
        ChangerHole
    End If
End Sub

Public Sub SetChangerHoleXY(hole As Double)
    If Changer_ValidStart(hole) Then
        RelabelPos MotorChanger, (ConvertHoletoPos(hole))
        currenthole = hole
        ChangerHole
    End If
End Sub

Private Sub SetMeasButton_click()
    ' Set the value of the place to move to the center of the Af coils
    MoveMotorPosEdit.text = str(Int(MeasPos + SampleHeight / 2))
    txtNewHeight = MoveMotorPosEdit.text
End Sub

Private Sub SetSCurveNext(motorid As Integer, sValue As Integer)
    ' full S curve is 32767
    'SendCommand motorid, MotorAddress(motorid) & "195 " + Str(sValue)
    'GetResponse (motorid)
End Sub

Private Sub SetTopButton_click()
    ' Set the value of the place to move to the top
    MoveMotorPosEdit.text = "0"
    txtNewHeight = "0"
End Sub

Private Sub SetTorques(motorid As Integer, ClosedHold As Integer, ClosedMove As Integer, OpenHold As Integer, OpenMove As Integer)
    'Set the appropriate torque limits
    '100% torque is entered by user under settings T.S. 2012
    Dim PerTorque As Integer
    PerTorque = 0.01 * UpDownMaxTorque 'Value for 1% torque
    
    Dim CH As Integer
    Dim CM As Integer
    Dim OH As Integer
    Dim OM As Integer
    CH = CInt(ClosedHold) * PerTorque
    CM = CInt(ClosedMove) * PerTorque
    OH = CInt(OpenHold) * PerTorque
    OM = CInt(OpenMove) * PerTorque
    SendCommand motorid, (MotorAddress & "149 " + CStr(CH) + " " + CStr(CM) + " " + CStr(OH) + " " + CStr(OM))
    GetResponse motorid
End Sub

Public Sub SetTurningMotorAngle(angle As Double)
    Dim pos As Long
    Dim turnSign As Integer
    
    ' (November 2009 L Carporzen) fix the rotation roudings
    ' (July 2012 I Hilburn) Altered the code to allow it to handle negative values for
    '                       TurningMotorFullRotation
    
    'Get the sign of the turning motor full rotation setting
    turnSign = Math.Sgn(TurningMotorFullRotation)
    
    'Get the current motor position
    pos = ReadPosition(MotorTurning)
    
    'If the motor position is less than zero and TurningMotorFullRotation is positive
    'OR, if the motor position is greater than zero and TurningMotorFullRotation is negative
    'Then add TurningMotorFullRotation to pos until pos is within 5% of +/- TurningMotorFullRotation
    If (pos < 0 And turnSign = 1) Then
       'If TurningMotorFullRotation > 0, then pos is increasing in value to more than
       '-0.95 * TurningMotorFullRotation.
        Do Until pos > -TurningMotorFullRotation * 0.95
            pos = pos + TurningMotorFullRotation
       Loop
    ElseIf (pos > 0 And turnSign = -1) Then
       'If TurningMotorFullRotation < 0, then pos is decreasing in value to less than
       '-0.05 * TurningMotorFullRotation.
       Do Until pos < -TurningMotorFullRotation * 0.05
            pos = pos + TurningMotorFullRotation
       Loop
    ElseIf (pos < 0 And turnSign = -1) Then
       'If TurningMotorFullRotation < 0, then pos is increasing in value to more than
        '0.95 * TurningMotorFullRotation.
        Do Until pos > TurningMotorFullRotation * 0.95
            pos = pos - TurningMotorFullRotation
        Loop
    ElseIf Abs(pos) < Abs(TurningMotorFullRotation * 0.05) Then
        'Do nothing
        'This is the close enough
    Else
        'If TurningMotorFullRotation > 0, then pos is decreasing in value to less than
        '0.95 * Abs(TurningMotorFullRotation).
        Do Until pos < Abs(TurningMotorFullRotation * 0.05)
            pos = pos - Abs(TurningMotorFullRotation)
        Loop
    End If
    
    RelabelPos MotorTurning, pos
    TurningAngleBox = ConvertPosToAngle(pos)
     'If TurningMotorAngle <> angle Then RelabelPos MotorTurning, pos
     'TurningAngleBox = angle
End Sub

Private Sub SetZeroButton_click()
    ' Set the value of the place to move to the measurement Zero position
    MoveMotorPosEdit.text = str(ZeroPos)
    txtNewHeight = str(ZeroPos)
End Sub

Public Function TurningMotorAngle() As Double
    Dim angle As Double
    angle = ConvertPosToAngle(ReadPosition(MotorTurning))
    TurningAngleBox = angle
    TurningMotorAngle = angle
End Function

Public Sub TurningMotorAngleOffset(ByVal angle As Double)
    TurningMotorRotate angle  ' (November 2009 L Carporzen) change - angle to + angle for clarity
    SetTurningMotorAngle angle
    ZeroTargetPos MotorTurning
End Sub

Public Sub TurningMotorRotate(ByVal angle As Double, Optional ByVal waitingForStop As Boolean = True, Optional ByVal pauseOverride As Boolean = True)
    Dim CurAngle As Double
    Dim startingangle As Double
    Dim startingPos As Long
    Dim target As Long
    Dim curpos As Long
    Dim ErrorMessage As String
    If Prog_halted Then Exit Sub
    If Not Prog_paused Or Prog_halted Then
        lastMoveCommand = lastCmdMove
        lastMoveMotor = MotorTurning
        lastMoveTarget = angle
    End If
    
    startingPos = ReadPosition(MotorTurning)
    startingangle = TurningMotorAngle
    
    target = ConvertAngleToPos(angle)
    SetSCurveNext MotorTurning, SCurveFactor
    MoveMotor MotorTurning, target, TurnerSpeed, waitingForStop, pauseOverride
    If Not waitingForStop Then Exit Sub
    CurAngle = TurningMotorAngle
    If Not NOCOMM_MODE And (Abs((CurAngle - angle) Mod 360) > 3) Then
        ' First try to move to move back to the desired position
        'MoveMotor MotorTurning, startingPos, TurnerSpeed, pauseOverride:=True
        'curpos = ReadPosition(MotorTurning)
        'SetSCurveNext MotorTurning, SCurveFactor
        MoveMotor MotorTurning, target, TurnerSpeed, pauseOverride:=True
        CurAngle = TurningMotorAngle
    End If
    ' Quit here if this is bad ...
    If Not NOCOMM_MODE And (Abs((CurAngle - angle) Mod 360) > 3) Then ' (November 2009 L Carporzen) Changed 3 to 1 degree
        curpos = ReadPosition(MotorTurning)
        ErrorMessage = "Unacceptable slop on turning motor from" & vbCrLf & _
            str(startingangle) & "degrees to " & str(angle) & " degrees." & _
            vbCrLf & vbCrLf & "Target position: " & str(target) & vbCrLf & _
            "Current position: " & str(curpos) & vbCrLf & vbCrLf & _
            "Execution has been paused. Please check machine."
        SetSCurveNext MotorTurning, SCurveFactor
        ' Here start the lines which could be remove for rockmag measurements:
        MoveMotor MotorTurning, (curpos - (curpos - startingPos) * 0.1), 0.1 * TurnerSpeed, pauseOverride:=True ' ???
        DelayTime 0.2 ' Can be remove to don't wait
        Flow_Pause ' Need to be remove to don't pause the measurement
        SetCodeLevel CodeRed ' Need to be remove to don't make the screen red
        frmSendMail.MailNotification "Unacceptable slop", ErrorMessage, CodeRed ' Send the email, could be keep for record
        MsgBox ErrorMessage ' Need to be remove to don't pop the window with the OK button
        SetCodeLevel StatusCodeColorLevelPrior, True ' Need to be remove to don't mess with the color code
    End If
    CurAngle = TurningMotorAngle
    CurAngle = CurAngle - (CurAngle \ 360) * 360 ' Integer division
    If CurAngle <> TurningMotorAngle Then SetTurningMotorAngle CurAngle
End Sub

Public Sub TurningMotorSpin(ByVal speedRPS As Double, Optional ByVal Duration As Double = 60, Optional ByVal pauseOverride As Boolean = False)
    Dim CurAngle As Double
    Dim startingangle As Double
    Dim startingPos As Long
    Dim target As Long
    Dim curpos As Long
    Dim ErrorMessage As String
    Dim activeAngle As Double
    If Prog_halted Then Exit Sub
'    If Not Prog_paused Or Prog_halted Then
'        lastMoveCommand = lastCmdMove
'        lastMoveMotor = MotorTurning
'        lastMoveTarget = angle
'    End If
    If speedRPS = 0 Then
        MotorStop MotorTurning
        activeAngle = TurningMotorAngle
        CurAngle = activeAngle - (activeAngle \ 360) * 360
        If CurAngle <> activeAngle Then SetTurningMotorAngle CurAngle
        activeAngle = TurningMotorAngle
        If CurAngle <> activeAngle Then SetTurningMotorAngle CurAngle
        If Abs(CurAngle) > 10 Then
            target = 360
        Else
            target = 0
        End If
        TurningMotorRotate target
        WaitForMotorStop MotorTurning
        SetTurningMotorAngle 0
        CurAngle = TurningMotorAngle
    Else
        startingPos = ReadPosition(MotorTurning)
        startingangle = TurningMotorAngle
        target = startingPos - TurningMotorFullRotation * speedRPS * Duration
        SetSCurveNext MotorTurning, SCurveFactor
        MoveMotor MotorTurning, target, Abs(TurningMotor1rps * speedRPS), False, pauseOverride
    End If
End Sub

Public Function UpDownHeight() As Double
    UpDownHeight = ReadPosition(MotorUpDown)
End Function

Public Sub UpDownMove(ByVal position As Long, ByVal speed As Integer, Optional ByVal waitingForStop As Boolean = True, Optional ByVal pauseOverride = False)
    Dim startingPos As Long
    Dim ErrorMessage As String
    Dim curpos As Long
    Dim movementSign As Integer
    If Prog_halted Then Exit Sub
    If Not Prog_paused Or Prog_halted Then
        lastMoveCommand = lastCmdMove
        lastMoveMotor = MotorUpDown
        lastMoveTarget = position
    End If
    startingPos = ReadPosition(MotorUpDown)
    If position < startingPos Then ' (Janv 2009) was If Position < curpos Then
        movementSign = -1
    Else
        movementSign = 1
    End If
    UpDownSpeeds(0) = LiftSpeedSlow
    UpDownSpeeds(1) = LiftSpeedNormal
    UpDownSpeeds(2) = LiftSpeedFast
    SetSCurveNext MotorUpDown, SCurveFactor
    MoveMotor MotorUpDown, position, UpDownSpeeds(speed), waitingForStop, pauseOverride
    If Not waitingForStop Then Exit Sub
    curpos = ReadPosition(MotorUpDown)
    ' back off a bit and try again if off
    If Not NOCOMM_MODE And Abs(curpos - position) > 100 And position <> 0 Then
        SetSCurveNext MotorUpDown, SCurveFactor
        MoveMotor MotorUpDown, (curpos + startingPos) / 2, LiftSpeedSlow, pauseOverride:=True
        SetSCurveNext MotorUpDown, SCurveFactor
        MoveMotor MotorUpDown, position, UpDownSpeeds(speed), pauseOverride:=True
        curpos = ReadPosition(MotorUpDown)
    End If
    ' quit here if this is bad
    If Not NOCOMM_MODE And (Abs(curpos - position) > 150) And position <> 0 Then
        ErrorMessage = "Unacceptable slop on up/down motor moving from" & vbCrLf & _
            str(startingPos) & " to " & str(position) & " at speed " & str(UpDownSpeeds(speed)) & "." & _
            vbCrLf & vbCrLf & "Target position: " & str(position) & vbCrLf & _
            "Current position: " & str(curpos) & vbCrLf & vbCrLf & _
            "Execution has been paused. Please check machine."
        MoveMotor MotorUpDown, (curpos - 100 * movementSign), 0.5 * LiftSpeedSlow, pauseOverride:=True
        DelayTime 0.2
        Flow_Pause
        MotorStop MotorUpDown
        SetCodeLevel CodeRed
        frmSendMail.MailNotification "Unacceptable slop", ErrorMessage, CodeRed
        MsgBox ErrorMessage
        SetCodeLevel StatusCodeColorLevelPrior, True
    End If
End Sub

Private Sub WaitForMotorStop(motorid As Integer, Optional pauseOveride As Boolean = False)
    Dim finished As Boolean
    Dim dummy As Boolean
    Dim curmotor As String
    Dim PollPosition As Long
    Dim oldPosition(1) As Long
    Dim pausedInitial As Boolean
    oldPosition(0) = 2 ^ 7
    oldPosition(1) = -2 ^ 7
    pausedInitial = Prog_paused
    frmProgram.StatusBar "Wait for stop...", 2
    'Now wait for motor to indicate that it is finished before continuing
    'We do this by polling the motor repeatedly until the appropriate bit is set in the polling word
    Do While ((finished = False))
        'having this delay inside the do loop instead of outside
        'appears to fix early drop program (according to Scott)
        DelayTime 0.1
        DoEvents
        If Not pauseOveride Then Flow_WaitForUnpaused
        oldPosition(1) = oldPosition(0)
        oldPosition(0) = PollPosition
        frmProgram.StatusBar str$(oldPosition(0)), 3
        Select Case motorid
            Case MotorUpDown: dummy = str(UpDownHeight)
            Case MotorTurning: dummy = str(TurningMotorAngle)
            Case MotorChanger: dummy = str(ChangerHole)
            Case MotorChangerY: dummy = str(ChangerHole)
        End Select
        PollPosition = ReadPosition(motorid)
        ' if we're not moving, we're done
        If oldPosition(1) = oldPosition(0) And oldPosition(0) = PollPosition Then
            finished = True
        ElseIf (Abs(oldPosition(1) - oldPosition(0)) < 5) And (Abs(oldPosition(0) - PollPosition) < 5) Then
            finished = True
        End If
        If finished Then frmProgram.StatusBar "Stopped", 2
    Loop
    frmProgram.StatusBar vbNullString, 2
    frmProgram.StatusBar vbNullString, 3
    MotorStop motorid ' just make sure we're really stopped
End Sub

Private Sub ZeroTargetPos(motorid As Integer, Optional attempt_number = 1)
    'Zero target and position
    Dim dummy As Long
    Dim is_command_successful As Boolean
    SendCommand motorid, (MotorAddress & "145")
    is_command_successful = GetResponse(motorid)
    
    If Not is_command_successful And attempt_number < 5 Then
        DelayTime 0.25
        ZeroTargetPos motorid, attempt_number + 1
    End If
    
    dummy = ReadPosition(motorid)
End Sub

Private Sub ZeroTargetPosButton_Click()
    'Zero target and position
    ZeroTargetPos ActiveMotorControls
End Sub

